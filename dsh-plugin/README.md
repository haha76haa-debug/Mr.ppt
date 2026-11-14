# dsh-plugin — PPT Studio

A DeepSeek Harness Host-only bundle that adds one model tool, `pptx`, over the
toolkit in this repository. One plugin install gives the agent a reliable
build / inspect / lint / render loop instead of hand-written shell commands.

This directory is the **source** of the bundle. `python/` is generated from
`src/pptx_native/` by `scripts/build_plugin.py`, so the Python package has one
home and is not duplicated in git.

## Build and install

```bash
pip install -e .                              # toolkit itself (for CLI use)
python scripts/build_plugin.py                # → build/pptx-studio-bundle/
```

Then in DeepSeek Harness: **Plugins → Add plugin** and pick the absolute path of
`build/pptx-studio-bundle` (or install it through the plugin manager API).

The bundle appends these files at build time:

```
build/pptx-studio-bundle/
├── index.js                 # Host plugin: registers the `pptx` tool
├── package.json             # dsh.bundle.patch manifest
├── cordis.patch.yml         # inserts the plugin row
├── locale/{en,zh}.json      # Plugin Manager title/description
├── icon.svg                 # Plugin Manager icon
├── README.md                # Plugin details fallback
├── python/pptx_native/      # vendored copy of src/pptx_native
└── scripts/export_slide.ps1 # PowerPoint COM export helper
```

## Tool

`pptx` takes an `action`:

| action | what it does | required params |
|---|---|---|
| `sample` | writes a 12-slide neutral demo deck | `out`, optional `theme` |
| `probe` | lists every shape on one slide (groups flattened, absolute coordinates) | `deck`, optional `slide` |
| `lint` | audits one slide: off-canvas elements, text overflow risk, overlapping text boxes, too-small fonts | `deck`, optional `slide` |
| `render` | exports one slide as PNG through PowerPoint COM | `deck`, `out`, optional `slide` |
| `themes` | lists the bundled theme presets | — |

Paths are absolute. Rendering needs Windows + Microsoft PowerPoint; everything
else is cross-platform.

## Runtime

Python discovery order: plugin config `python` → `DSH_PYTHON` →
`$DSH_HOME/dsh-runtimes/*/dependencies/python/python.exe` → `python` / `py -3` /
`python3` on PATH. The first interpreter that can
`import pptx, PIL, numpy, lxml` wins. To pin one, edit the plugin row config:

```yaml
- id: pptx-studio
  name: '@local/pptx-studio'
  config:
    python: 'C:\path\to\python.exe'
```

The tool runs `python -m pptx_native` with `PYTHONPATH=<bundle>/python`; it never
installs anything and sends nothing over the network.

## Limitations

* `render` requires Windows with Microsoft PowerPoint.
* The tool writes files directly to the paths it is given and does not pass
  through the session sandbox. Pass paths inside the workspace you intend to
  modify.
* Calls time out after 180 s; a large deck takes a few seconds to render.
* The vendored Python package is a snapshot. Rebuild after changing `src/`:

  ```bash
  python scripts/build_plugin.py --out <installed bundle dir>
  ```

  Restart DeepSeek Harness (or reinstall the bundle) so the Host loads the new
  JavaScript generation.
* This is a local-only bundle (`@local/...`). It is not published to a registry.

## Uninstall

**Plugins** → remove the `pptx-studio` bundle. Nothing else is written outside
the paths you asked the tool to write.

## Related

* The methodology lives in [SKILL.md](../SKILL.md) and [docs/](../docs/) — the
  plugin is the execution shell, the skill is the know-how.
* `<https://github.com/haha76haa-debug/Mr.ppt>`
