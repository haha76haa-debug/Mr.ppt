# pptx-native

**用原生形状做 PPT——文字能改、形状能改、图片能换，最后在 PowerPoint 里验收。**

[![CI](https://github.com/haha76haa-debug/Mr.ppt/actions/workflows/ci.yml/badge.svg)](https://github.com/haha76haa-debug/Mr.ppt/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

[English](README.md) · [设计系统](docs/design-system.md) · [版式库](docs/layouts.md) · [工作流](docs/workflow.md) · [Skill](SKILL.md)

一套「设计系统 + 版式库 + 校验工具链」，用来**程序化地做 PPT**。
每一页都是真文本框、真形状、原生图表、原生表格——收到 `.pptx` 的人可以直接改，
而不是对着一张截图发呆。

典型场景：

- 用脚本 / 数据 / agent 生成一套 deck；
- 按截图把某一页 1:1 复刻成可编辑模块；
- 改造既有 deck（统一主题色、换字体、加注释条、做图片墙）；
- 确认生成的文件在 PowerPoint 里真的渲染正确。

## 安装

```bash
git clone https://github.com/haha76haa-debug/Mr.ppt.git
cd Mr.ppt
pip install -e .
```

依赖：`python-pptx`、`pillow`、`lxml`、`numpy`。
真机渲染（`pptx-native export`）需要 **Windows + Microsoft PowerPoint**，其余功能跨平台。

## 60 秒上手

```bash
# 1) 生成一份 12 页中性示例 deck（不依赖任何外部素材）
pptx-native sample demo.pptx

# 2) 看有哪些主题预设
pptx-native themes

# 3) 探明某一页的结构（含组合内层、真实渲染坐标）
pptx-native probe deck.pptx --slide 3

# 4) 用 PowerPoint 真机导出 PNG（唯一可信的验收图）
pptx-native export deck.pptx slide3.png --slide 3

# 5) 颜色普查 + 统一
pptx-native colors deck.pptx --slide 3
pptx-native sweep deck.pptx out.pptx --slide 3 --to 2E7BE8
```

## 在代码里用

```python
from pptx_native import layouts
from pptx_native.theme import Theme

theme = Theme.blue()                       # 或 dark / minimal / warm
prs = layouts.new_deck(theme)

layouts.cover(prs, theme, kicker="季度回顾",
              title="Q1 回顾", subtitle="交付与质量",
              meta="产品团队 · 2026-04")
layouts.agenda(prs, theme, title="本次内容",
               items=["背景", "方法", "数据", "下一步"], page=2)
layouts.stats(prs, theme, title="关键数据", page=3,
              items=[("128", "个", "需求"), ("96", "%", "按时率")],
              note="来源：内部看板")
layouts.table(prs, theme, title="指标对照", page=4,
              header=["指标", "目标", "实际"],
              rows=[["交付", "≥120", "128"], ["按时", "≥95%", "96%"]])
layouts.chart(prs, theme, title="季度趋势", page=5, kind="column",
              categories=["Q1", "Q2", "Q3", "Q4"],
              series=[("计划", [24, 30, 35, 39]), ("实际", [22, 29, 34, 38])])
layouts.end(prs, theme, subtitle="欢迎提问")

prs.save("review.pptx")
```

13 种版式：`cover`、`agenda`、`section`、`bullets`、`two_col`、`stats`、
`timeline`、`process`、`image_grid`、`quote`、`table`、`chart`、`end`。
参数见 [docs/layouts.md](docs/layouts.md)。

## 命令行

| 命令 | 作用 |
|---|---|
| `sample` | 生成 12 页示例 deck（`--theme blue/dark/minimal/warm`） |
| `themes` | 列出主题预设 |
| `probe` | 导出某页全部形状（组合展开、绝对坐标） |
| `lint` | 页面体检：越界 / 文字溢出风险 / 文本框重叠 / 字号过小 |
| `preview` | 只保留指定页，输出小 deck（渲染快、大文件也能打开） |
| `export` | 用 PowerPoint COM 真机导出某页 PNG（Windows + Office） |
| `colors` | 颜色普查：有哪些颜色、分别出现在什么上下文 |
| `sweep` | 把某页的红色系统一替换（保留透明度） |
| `borders` | 修掉图片文件自带的 1–2px 边框 |
| `diff` | 两版单页结构 diff |
| `overlaps` | 形状/文字重叠自检 |

## 为什么需要一套工具，而不是直接调 python-pptx

python-pptx 的 API 本身不难，难的是「看起来对 + 仍然可编辑」。坑集中在这几处：

| 坑 | 表现 |
|---|---|
| 组合 `chOff ≠ off` | 元素整体偏 0.1~0.2in，怎么调都对不齐 |
| 组上的 `scene3d` | PowerPoint 里照片是斜的，LibreOffice 里却是平的 |
| 主题字体 / autofit | 中文被换成默认字体，字号被自动缩小 |
| 主题色残留 | 旧红色藏在文字渐变、描边、阴影、照片像素里 |
| 图片自带边框 | 每张 JPEG 最外 1–2px 一圈杂色 |
| deck 太大 | 渲染器报 `maxInputBytes`，PowerPoint 打不开 |

本仓库把这些修复封装成函数，再加一套设计系统防止页面跑偏，最后用
「PowerPoint 真机导出」收尾，形成闭环。

## DeepSeek Harness 插件（可选）

`dsh-plugin/` 把工具链变成一个 DSH 模型工具：

```bash
python scripts/build_plugin.py          # → build/pptx-studio-bundle/
```

然后在 **Plugins → Add plugin** 选择 `build/pptx-studio-bundle`。工具名 `pptx`，
action 为 `sample` / `probe` / `lint` / `render` / `themes`；插件自带 Python 包，
不需要额外 `pip install`。详见 [docs/dsh-plugin.md](docs/dsh-plugin.md) 与
[dsh-plugin/README.md](dsh-plugin/README.md)。

## 项目结构

```
src/pptx_native/
├── theme.py       设计系统：颜色角色、字号阶梯、12 栏网格
├── layouts.py     13 种原生版式
├── styling.py     字体 / 渐变 / 半透明 / 叠层文字 / 竖排
├── geometry.py    组合坐标换算（abs_rect）
├── groups.py      把形状搬进组合、复制并重编号
├── colors.py      颜色普查 + 红色清扫
├── images.py      cover 裁切 / 补边 / 自带边框修复
├── checks.py      快照、diff、重叠自检
├── probe.py       页面探针
├── render.py      抽页预览 + PowerPoint COM 导出
├── sample.py      中性 12 页示例 deck
└── cli.py         命令行
tests/             pytest 测试（无需 Office）
scripts/           selfcheck、PowerPoint 导出脚本、插件构建
dsh-plugin/        DeepSeek Harness Host 插件源码（用 scripts/build_plugin.py 构建）
docs/              设计系统、版式、工作流、校验、反模式……
SKILL.md           把整套流程写成 agent skill
```

## 文档

- [docs/design-system.md](docs/design-system.md) —— 网格、字号、配色、间距
- [docs/layouts.md](docs/layouts.md) —— 13 种版式的参数与几何
- [docs/workflow.md](docs/workflow.md) —— 从需求到交付的完整流程
- [docs/native-editing.md](docs/native-editing.md) —— 改既有 deck 的安全做法
- [docs/xml-cookbook.md](docs/xml-cookbook.md) —— 原始 OOXML 配方
- [docs/verification.md](docs/verification.md) —— 三层校验与交付清单
- [docs/image-prep.md](docs/image-prep.md) —— 图片素材处理
- [docs/anti-patterns.md](docs/anti-patterns.md) —— 坏设计 → 好设计对照
- [docs/dsh-plugin.md](docs/dsh-plugin.md) —— DeepSeek Harness 插件 bundle

## 测试

```bash
pytest -q                     # CI：Python 3.9~3.12（Ubuntu）
python scripts/selfcheck.py   # 没装 pytest 时的纯 assert 版
```

CI 每次 push 自动跑；PowerPoint 真机渲染只在 Windows 上手动验证。

## 许可

MIT —— 见 [LICENSE](LICENSE)。
