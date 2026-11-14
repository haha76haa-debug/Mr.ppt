/**
 * PPT Studio — Host-only bundle for DeepSeek Harness.
 *
 * Registers one model-facing tool, `pptx`, over a vendored copy of the
 * `pptx-native` toolkit. The tool only shells out to the bundled Python CLI:
 * no network, no telemetry, and writes go to the paths the caller passes.
 */
import { spawn, spawnSync } from 'node:child_process'
import { existsSync, mkdirSync, readdirSync } from 'node:fs'
import { dirname, isAbsolute, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const ROOT = dirname(fileURLToPath(import.meta.url))
const PYTHON_HOME = join(ROOT, 'python')
const EXPORT_SCRIPT = join(ROOT, 'scripts', 'export_slide.ps1')

const MAX_OUTPUT = 12000
const MAX_ERROR = 4000
const CLI_TIMEOUT_MS = 180000
const IMPORT_CHECK = 'import pptx, PIL, numpy, lxml'

export const name = 'pptx-studio'
export const inject = ['tools']

/**
 * Model-facing parameter schema, already in the registry's raw JSON-Schema form.
 * `defineTool` normally compiles the authoring DSL into this shape; doing it
 * here keeps the plugin free of an `@deepseek-ai/*` import, which a
 * profile-installed bundle cannot resolve.
 */
const PARAMETERS = {
  type: 'object',
  properties: {
    action: {
      type: 'string',
      enum: ['sample', 'probe', 'lint', 'render', 'themes'],
      description: 'Which operation to run.',
    },
    out: {
      type: 'string',
      description: 'Absolute output path: the .pptx for sample, the .png for render.',
    },
    deck: {
      type: 'string',
      description: 'Absolute path of an existing .pptx; required by probe, lint and render.',
    },
    slide: {
      type: 'integer',
      description: 'One-based slide number for probe, lint and render; defaults to 1.',
    },
    theme: {
      type: 'string',
      description: 'Theme preset for sample: blue (default), dark, minimal, or warm.',
    },
  },
  required: ['action'],
}

const OUTPUT_SCHEMA = {
  type: 'object',
  additionalProperties: false,
  properties: {
    ok: { type: 'boolean' },
    action: { type: 'string' },
    message: { type: 'string' },
  },
  required: ['ok', 'action', 'message'],
}

/** Absolute path for a caller-supplied path; relative paths resolve against the Host cwd. */
export function absolute(value) {
  return isAbsolute(value) ? value : resolve(process.cwd(), value)
}

/** Candidate Python interpreters, best first, with an explanation for each. */
export function pythonCandidates(config = {}) {
  const list = []
  if (config && typeof config.python === 'string' && config.python.trim() !== '') {
    list.push({ cmd: config.python.trim(), args: [], label: 'plugin config python' })
  }
  if (process.env.DSH_PYTHON) {
    list.push({ cmd: process.env.DSH_PYTHON, args: [], label: 'DSH_PYTHON' })
  }
  const home = process.env.DSH_HOME
  if (home) {
    const runtimes = join(home, 'dsh-runtimes')
    try {
      for (const entry of readdirSync(runtimes, { withFileTypes: true })) {
        if (!entry.isDirectory()) continue
        const exe = join(runtimes, entry.name, 'dependencies', 'python', 'python.exe')
        if (existsSync(exe)) list.push({ cmd: exe, args: [], label: 'DSH bundled runtime ' + entry.name })
      }
    } catch {
      // No bundled runtime directory: fall through to PATH candidates.
    }
  }
  list.push({ cmd: 'python', args: [], label: 'PATH python' })
  if (process.platform === 'win32') list.push({ cmd: 'py', args: ['-3'], label: 'PATH py -3' })
  list.push({ cmd: 'python3', args: [], label: 'PATH python3' })
  return list
}

let resolved = null

/** Pick the first Python that can import the toolkit dependencies. */
export function detectPython(config = {}) {
  if (resolved && (!config || !config.python)) return resolved
  const tried = []
  for (const candidate of pythonCandidates(config)) {
    try {
      const probe = spawnSync(candidate.cmd, [...candidate.args, '-c', IMPORT_CHECK], {
        timeout: 20000,
        windowsHide: true,
      })
      if (probe.status === 0) {
        resolved = candidate
        return candidate
      }
      tried.push(`${candidate.cmd}: exit ${probe.status}`)
    } catch (error) {
      tried.push(`${candidate.cmd}: ${error && error.code ? error.code : error}`)
    }
  }
  throw new Error(
    'pptx: no Python interpreter with python-pptx / pillow / numpy / lxml was found. '
    + 'Tried: ' + tried.join('; ') + '. '
    + 'Install the packages, or set the plugin config "python" to an interpreter path.',
  )
}

/** Run the bundled `pptx-native` CLI and resolve with its bounded stdout. */
export function runCli(python, args, exec = {}) {
  return new Promise((resolvePromise, reject) => {
    let child
    try {
      child = spawn(python.cmd, [...python.args, '-m', 'pptx_native', ...args], {
        cwd: ROOT,
        windowsHide: true,
        signal: exec.signal,
        env: { ...process.env, PYTHONPATH: PYTHON_HOME, PYTHONIOENCODING: 'utf-8' },
      })
    } catch (error) {
      reject(error)
      return
    }
    let stdout = ''
    let stderr = ''
    let settled = false
    const settle = (fn, value) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      fn(value)
    }
    const timer = setTimeout(() => {
      child.kill('SIGKILL')
      settle(reject, new Error(`pptx-native timed out after ${CLI_TIMEOUT_MS} ms`))
    }, CLI_TIMEOUT_MS)
    child.stdout.on('data', (chunk) => {
      if (stdout.length >= MAX_OUTPUT) return
      stdout += chunk.toString('utf8').slice(0, MAX_OUTPUT - stdout.length)
    })
    child.stderr.on('data', (chunk) => {
      if (stderr.length >= MAX_ERROR) return
      stderr += chunk.toString('utf8').slice(0, MAX_ERROR - stderr.length)
    })
    child.on('error', (error) => {
      settle(reject, new Error(`pptx-native failed to start: ${error.message}`))
    })
    child.on('close', (code) => {
      if (code === 0) {
        settle(resolvePromise, stdout.trim())
      } else {
        settle(reject, new Error(`pptx-native exited with ${code}`
          + (stderr.trim() ? `\n${stderr.trim()}` : '')))
      }
    })
  })
}

function requirePath(args, key, label) {
  const value = args[key]
  if (typeof value !== 'string' || value.trim() === '') {
    throw new Error(`pptx: "${key}" is required for action=${label}`)
  }
  return absolute(value.trim())
}

function slideNumber(args) {
  return Number.isInteger(args.slide) && args.slide > 0 ? args.slide : 1
}

function ensureParent(file) {
  mkdirSync(dirname(file), { recursive: true })
}

/** Execute one tool call and return the rendered message. */
export async function runAction(config, args, exec = {}) {
  const python = detectPython(config)
  const action = args.action
  switch (action) {
    case 'themes': {
      return runCli(python, ['themes'], exec)
    }
    case 'sample': {
      const out = requirePath(args, 'out', 'sample')
      ensureParent(out)
      const theme = typeof args.theme === 'string' && args.theme.trim() !== '' ? args.theme.trim() : 'blue'
      const log = await runCli(python, ['sample', out, '--theme', theme], exec)
      return log + `\n\n示例 deck 已生成：${out}（theme=${theme}）`
    }
    case 'probe': {
      const deck = requirePath(args, 'deck', 'probe')
      if (!existsSync(deck)) throw new Error(`pptx: deck not found: ${deck}`)
      return runCli(python, ['probe', deck, '--slide', String(slideNumber(args)), '--text-only'], exec)
    }
    case 'lint': {
      const deck = requirePath(args, 'deck', 'lint')
      if (!existsSync(deck)) throw new Error(`pptx: deck not found: ${deck}`)
      return runCli(python, ['lint', deck, '--slide', String(slideNumber(args))], exec)
    }
    case 'render': {
      if (process.platform !== 'win32') {
        throw new Error('pptx: render requires Windows with Microsoft PowerPoint')
      }
      const deck = requirePath(args, 'deck', 'render')
      if (!existsSync(deck)) throw new Error(`pptx: deck not found: ${deck}`)
      const out = requirePath(args, 'out', 'render')
      ensureParent(out)
      if (!existsSync(EXPORT_SCRIPT)) {
        throw new Error(`pptx: missing export script: ${EXPORT_SCRIPT}`)
      }
      const slide = slideNumber(args)
      const log = await runCli(python, ['export', deck, out, '--slide', String(slide), '--script', EXPORT_SCRIPT], exec)
      return log + `\n\n真机渲染图：${out}（slide=${slide}）`
    }
    default:
      throw new Error(`pptx: unknown action ${JSON.stringify(action)}`)
  }
}

/**
 * Register the `pptx` tool on the Host tool registry.
 *
 * The definition is assembled by hand instead of `defineTool` because a
 * profile-installed bundle cannot resolve `@deepseek-ai/dsh-tools`; the registry
 * accepts the same raw definition shape that `defineTool` produces.
 */
export async function apply(ctx, config = {}) {
  ctx.tools.register({
    name: 'pptx',
    description:
      'Build and verify PowerPoint slides with the bundled pptx-native toolkit. '
      + 'Actions: sample writes a 12-slide demo deck; probe lists every shape on one slide '
      + 'with absolute coordinates; lint audits a slide for off-canvas elements, text '
      + 'overflow risk, overlapping text boxes and too-small fonts; render exports one '
      + 'slide as PNG through PowerPoint (Windows + Office only); themes lists the bundled '
      + 'theme presets. Pass absolute paths. The result returns the written paths, and a '
      + 'rendered PNG can be viewed with the image-reading tool.',
    parameters: PARAMETERS,
    output: {
      schema: OUTPUT_SCHEMA,
      render: (_args, value) => [{ type: 'text', text: value.message }],
    },
    async execute(args, exec) {
      if (typeof args.action !== 'string' || args.action === '') {
        throw new Error('pptx: "action" is required (sample | probe | lint | render | themes)')
      }
      const message = await runAction(config, args, exec)
      return { ok: true, action: args.action, message }
    },
  })
}
