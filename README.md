# pptx-native

**Build PowerPoint decks as native, editable shapes — then check them in PowerPoint itself.**

[![CI](https://github.com/haha76haa-debug/Mr.ppt/actions/workflows/ci.yml/badge.svg)](https://github.com/haha76haa-debug/Mr.ppt/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

[中文文档](README_CN.md) · [Design system](docs/design-system.md) · [Layouts](docs/layouts.md) · [Workflow](docs/workflow.md) · [Skill](SKILL.md)

A small design system + layout library + verification toolkit for building decks
programmatically. Every slide is made of real text boxes, real shapes, native charts
and native tables — so the person receiving the `.pptx` can edit it instead of
staring at a screenshot.

Typical uses:

- generate a deck from data / a script / an agent;
- rebuild a slide 1:1 from a screenshot as editable modules;
- restyle an existing deck (theme colours, fonts, captions, image grids);
- verify that what you generated actually renders right in PowerPoint.

## Install

```bash
git clone https://github.com/haha76haa-debug/Mr.ppt.git
cd Mr.ppt
pip install -e .
```

Requires `python-pptx`, `pillow`, `lxml`, `numpy`. Real-device rendering
(`pptx-native export`) needs **Windows + Microsoft PowerPoint**; everything else
is cross-platform.

## 60-second quickstart

```bash
# 1) build a neutral 12-slide sample deck (no external assets needed)
pptx-native sample demo.pptx

# 2) list theme presets
pptx-native themes

# 3) inspect an existing slide (all shapes, groups flattened, real coordinates)
pptx-native probe deck.pptx --slide 3

# 4) export a slide as PNG using PowerPoint itself (the only trustworthy check)
pptx-native export deck.pptx slide3.png --slide 3

# 5) audit and unify stray colours
pptx-native colors deck.pptx --slide 3
pptx-native sweep deck.pptx out.pptx --slide 3 --to 2E7BE8
```

## Use it from Python

```python
from pptx_native import layouts
from pptx_native.theme import Theme

theme = Theme.blue()                       # or dark / minimal / warm
prs = layouts.new_deck(theme)

layouts.cover(prs, theme, kicker="QUARTERLY REVIEW",
              title="Q1 Review", subtitle="Delivery and quality",
              meta="Product team · 2026-04")
layouts.agenda(prs, theme, title="Agenda",
               items=["Context", "Method", "Data", "Next"], page=2)
layouts.stats(prs, theme, title="Key numbers", page=3,
              items=[("128", "items", "delivered"), ("96", "%", "on time")],
              note="Source: internal dashboard")
layouts.table(prs, theme, title="Targets", page=4,
              header=["Metric", "Target", "Actual"],
              rows=[["Delivery", ">= 120", "128"], ["On time", ">= 95%", "96%"]])
layouts.chart(prs, theme, title="Quarterly trend", page=5, kind="column",
              categories=["Q1", "Q2", "Q3", "Q4"],
              series=[("Plan", [24, 30, 35, 39]), ("Actual", [22, 29, 34, 38])])
layouts.end(prs, theme, subtitle="Questions welcome")

prs.save("review.pptx")
```

13 slide archetypes are available: `cover`, `agenda`, `section`, `bullets`,
`two_col`, `stats`, `timeline`, `process`, `image_grid`, `quote`, `table`,
`chart`, `end`. See [docs/layouts.md](docs/layouts.md).

## CLI

| Command | Purpose |
|---|---|
| `sample` | generate a 12-slide sample deck (`--theme blue/dark/minimal/warm`) |
| `themes` | list theme presets |
| `probe` | dump every shape on a slide (groups flattened, absolute coordinates) |
| `preview` | keep selected slides into a small deck (render fast, open big decks) |
| `export` | export a slide as PNG through PowerPoint COM (Windows + Office) |
| `colors` | colour census: which colours, in which context |
| `sweep` | replace reddish colours on one slide, keeping alpha |
| `borders` | repaint 1–2 px frames baked into photo files |
| `diff` | structural diff between two versions of a slide |
| `overlaps` | overlapping-shape audit |

## Why a toolkit, not just python-pptx

python-pptx is fine until the result has to *look right* and *stay editable*.
The hard parts are not the API calls:

| Trap | Symptom |
|---|---|
| group `chOff != off` | everything is off by ~0.13 in and nothing lines up |
| `scene3d` on a group | photos are skewed in PowerPoint, but flat in LibreOffice |
| theme fonts / autofit | Chinese falls back to another font, sizes shrink |
| stray theme colours | old red lives in text gradients, outlines, shadows, photo pixels |
| baked-in photo borders | a 1–2 px frame around every JPEG |
| too-large decks | renderers report `maxInputBytes`, PowerPoint refuses to open |

This repo packages the fixes for those traps as functions, plus a design system so
slides don't drift apart, plus a verification loop that ends in PowerPoint.

## Project layout

```
src/pptx_native/
├── theme.py       design system: colours, type scale, 12-column grid
├── layouts.py     13 slide archetypes built from native shapes
├── styling.py     fonts, gradients, alpha, layered text, vertical text
├── geometry.py    group coordinate transforms (abs_rect)
├── groups.py      move shapes into groups, clone + renumber
├── colors.py      colour census and reddish sweep
├── images.py      cover crop, pad, baked-in border repair
├── checks.py      snapshot, diff, overlap audit
├── probe.py       slide inspector
├── render.py      preview extraction + PowerPoint COM export
├── sample.py      neutral 12-slide sample deck
└── cli.py         command line
tests/             pytest suite (no Office required)
scripts/           selfcheck, PowerPoint export helper
docs/              design system, layouts, workflow, verification, ...
SKILL.md           the workflow as an agent skill
```

## Docs

- [docs/design-system.md](docs/design-system.md) — grid, type scale, colour roles
- [docs/layouts.md](docs/layouts.md) — 13 archetypes with parameters and geometry
- [docs/workflow.md](docs/workflow.md) — from a request to a delivered deck
- [docs/native-editing.md](docs/native-editing.md) — editing existing decks safely
- [docs/xml-cookbook.md](docs/xml-cookbook.md) — raw OOXML recipes
- [docs/verification.md](docs/verification.md) — three-layer verification
- [docs/image-prep.md](docs/image-prep.md) — image handling
- [docs/anti-patterns.md](docs/anti-patterns.md) — bad design → good design

## Tests

```bash
pytest -q                  # CI: Python 3.9–3.12 on Ubuntu
python scripts/selfcheck.py  # pure-assert fallback, no pytest needed
```

CI runs on every push; PowerPoint rendering is Windows-only and verified manually.

## License

MIT — see [LICENSE](LICENSE).
