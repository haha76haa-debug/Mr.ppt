---
name: pptx-native
description: 通用 PPT 制作 / 改版 skill。做一套演示文稿、按参考图复刻或改造某一页、统一主题色与字体、加注释条/图片墙/图表页，最终产出「文字可改、形状可改、图片可换」的原生 .pptx，并用 PowerPoint 真机导出验收。Use when asked to build a PowerPoint deck or slides, redesign/restyle a page 1:1 from a screenshot, rebuild a slide as editable native shapes, unify stray colours or fonts, add photo captions, fix crooked/overlapping/clipped elements, or verify a .pptx actually renders right in PowerPoint. 触发词：PPT、幻灯片、演示文稿、deck、汇报、答辩、路演、分享会、按图复刻、做成可编辑、主题色、微软雅黑、统一风格、加注释、照片墙、数据页、歪了/溢出/红色残留。
---

# 用原生组件做 PPT（通用版）

两条主线，同一套工具：

* **A. 从零做一套**：需求 → 大纲 → 设计系统 → 逐页版式 → 校验 → 交付。
* **B. 改 / 复刻既有页**：量测参考图 → 重建原生元素（或直接改 XML）→ 真机校验 → 交付。

产出标准只有一句：**用户在 PowerPoint 里能直接改字、改色、换图，而且和设计稿一致。**
不是「导出一张图贴上去」。

```
pip install -e .
pptx-native sample demo.pptx     # 先看 12 页通用版式
pptx-native themes               # 看主题预设
pptx-native lint demo.pptx --slide 1   # 页面体检
```

> 装了 DSH 插件（`dsh-plugin/`）时，以上命令对应 `pptx` 工具的
> `sample` / `themes` / `lint` / `probe` / `render`，无需 pip 安装。

## 准备（首次使用）

```bash
git clone https://github.com/haha76haa-debug/Mr.ppt.git
cd Mr.ppt && pip install -e .
```

* `pip` 不可用时：`PYTHONPATH=<repo>/src python -m pptx_native ...`
* 真机验收（`export` / 插件 `render`）需要 Windows + Microsoft PowerPoint。
* 装成用户级技能：把本文件与 `docs/` 放进 `~/.agents/skills/pptx-native/`
  （或项目里的 `.dsh/skills/`），模型会在 PPT 任务上自动加载。

---

## 0. 先判断是哪类任务

| 任务 | 入口 | 关键风险 |
|---|---|---|
| 从零做一套 deck | `theme.py` + `layouts.py` | 内容结构、版式节奏 |
| 按截图复刻一页 | `probe` + 量测 + `build_*.py` | 坐标、字体、渲染差异 |
| 改既有 deck 的配色/字体 | `colors` / `sweep` / XML | 主题色藏在描边、阴影、渐变里 |
| 给照片加注释 / 图片墙 | `layouts.image_grid` / `groups` | 组合变换、3D 透视 |
| 「帮我检查一下」 | `lint` / `probe` / `diff` / `overlaps` / `export` | 只看预览会漏 |

---

## 1. 十条铁律

1. **先定设计系统，再排页面。** 12 栏网格、字号阶梯、颜色角色先写死；页面只引用角色，不写裸值。
2. **一页一个论点。** 标题写结论（「交付准时率 96%」），不是写主题（「交付情况」）。
3. **文字必须真文字。** 图标/描边/阴影用形状和 OOXML 做，不截图。
4. **图片只 cover 填满。** 不拉伸、不留白边、不加阴影（除非用户明确要）。
5. **几何靠量测。** 参考图先标定 px/in，写常量并注明出处；不要目测。
6. **组合是最大陷阱。** 组合有 `off/ext/chOff/chExt`，子元素坐标 ≠ 渲染坐标；要「贴合」就进同一个组。
7. **验收用 PowerPoint 真机。** LibreOffice / 在线预览只配当快速检查。
8. **改完必做结构 diff。** 形状数、几何、文字没变才算「只改了颜色/图片」。
9. **颜色问题全量扫。** 文字、描边、阴影、渐变、图片像素，一处都别漏。
10. **别猜内容。** 数据、文案、哪张图放哪，问用户；样式与量测自己判断。

---

## 2. 流程总览

```
0) 问清 5 件事：受众 / 场合 / 时长 / 页数 / 素材与禁忌
1) 写大纲：每页一句 takeaway，标出页面类型（封面/议程/过渡/要点/数据/…）
2) 定设计系统：Theme（颜色角色、字体、字号、边距）
3) 逐页搭：layouts.* 或自定义版式；先搭骨架，后填内容
4) 素材：照片 cover 裁切、截图补边、图表用原生 Chart、表格用原生 Table
5) 自检：文字溢出 / 重叠 / 越界 / 对齐 / 颜色数量
6) 真机：pptx-native export 导出关键页 PNG，逐页看
7) 交付：复制到目标位置，核对大小与时间，写清改了什么
```

页数参考：10 分钟讲 8~12 页；20 分钟 15~20 页。一页讲超过 90 秒，就拆。

---

## 3. 设计系统（`src/pptx_native/theme.py`）

```python
from pptx_native.theme import Theme
from pptx_native import layouts

theme = Theme.blue().with_fonts("Microsoft YaHei")   # 或 Theme.dark/minimal/warm
prs = layouts.new_deck(theme)                        # 16:9
layouts.cover(prs, theme, title="季度回顾", subtitle="2026 Q1", meta="团队名")
layouts.stats(prs, theme, title="关键数据", items=[("128", "个", "需求")])
```

四件事必须显式定义：

| 维度 | 规则 | 默认值 |
|---|---|---|
| 画布 | 16:9 | 13.333 × 7.5 in |
| 网格 | 12 栏，边距 0.72in，槽 0.28in | `theme.col(0, 6)` 取半栏 |
| 字号 | 阶梯制，一页最多 4 级 | display 88 / hero 44 / h1 34 / h3 19 / body 16 / small 13 / caption 11 |
| 颜色 | 5 个角色，最多再加 1 个强调色 | primary / primary_dark / primary_light / accent / text / muted / bg / surface / line |

细则见 [docs/design-system.md](docs/design-system.md)。

---

## 4. 版式库（`src/pptx_native/layouts.py`）

全部是「参数进、原生页面出」，换主题不用改版式代码：

| 函数 | 用途 | 关键参数 |
|---|---|---|
| `cover` | 封面 | `title, subtitle, kicker, meta` |
| `agenda` | 目录（5 项起自动两列） | `items` |
| `section` | 章节过渡 | `number, title, subtitle` |
| `bullets` | 要点页（小标题 + 一句说明） | `items=[(head, body)]` |
| `two_col` | 左右对比 | `left=(head, body), right=...` |
| `stats` | 关键数字并排（3~5 个） | `items=[(value, unit, label)]` |
| `timeline` | 时间线（均分节点） | `items=[(when, head, body)]` |
| `process` | 流程（N 步卡片） | `steps=[(head, body)]` |
| `image_grid` | 图片墙（可留空占位） | `items=[(path/bytes/None, caption)], cols` |
| `quote` | 引用 / 金句 | `text, author` |
| `table` | 原生表格 | `header, rows, weight` |
| `chart` | 原生图表（柱/条/折线） | `categories, series, kind` |
| `end` | 结尾页 | `title, subtitle, contact` |

版式参数与几何见 [docs/layouts.md](docs/layouts.md)；不满意就直接在函数里改，
所有坐标都是英寸，改完 `pptx-native export` 看真机效果。

---

## 5. 从参考图反推参数（复刻页）

1. **标定比例**：参考图宽 `W px` 对应 13.333in → `k = 13.333 / W`（in/px）。截图一般是 150dpi 左右。
2. **量 bbox**：每个元素量 `x1,y1,x2,y2`（px）→ 乘 `k` 得英寸，写进常量表，注明「来自参考图第几行」。
3. **反推字号**：中文 ink 高 ≈ 字号的 0.72 倍 → `size_pt ≈ ink_px × k × 72 / 0.72`；行距 `line_spacing = 基线间距 / 字号`。
4. **取色**：直接取填充区中心像素（避开抗锯齿边缘）；半透明色要算出 alpha。
5. **对齐关系**：先判断元素属于哪条基线/哪一栏，再决定是「绝对坐标」还是「相对偏移」。

```python
# 例：量到标题 ink 高 43px，图宽 2000px
k = 13.333 / 2000
size = 43 * k * 72 / 0.72        # ≈ 28.7pt → 取 28 或 29
```

---

## 6. 改既有 deck（不是重画）

```bash
pptx-native probe deck.pptx --slide 3          # 全量形状：id/类型/绝对坐标/文字
pptx-native colors deck.pptx --slide 3         # 颜色普查：哪些红/杂色、在什么上下文
pptx-native sweep deck.pptx out.pptx --slide 3 --to 2E7BE8
pptx-native borders deck.pptx out.pptx --slide 3 --to 1B3E78   # 照片自带红边
pptx-native diff before.pptx after.pptx --slide 3              # 结构 diff
```

* **组合坐标**用 `geometry.abs_rect()`；要搬元素进组用 `groups.move_into_group()`。
* **组上有 `scene3d`**（透视相机）时，PowerPoint 会把整组画成斜的；要贴合的装饰必须进同一个组。
* **主题色残留**常在：文字 `gradFill` 的双 stop、文字描边 `a:ln`、阴影 `a:outerShdw`、组内小图标阴影。
* 改颜色/换图后，**几何与文字必须完全不变**（`diff` 的 GEOM/TEXT 都应为 0）。

细节见 [docs/native-editing.md](docs/native-editing.md)。

---

## 7. 素材

* 照片：`images.cover(img, w*300, h*300)` 按槽位比例裁切，再 `add_picture`；不拉伸、不留边。
* 截图：`images.pad` 整图可见 + 采样背景色补边（不要裁掉正文）。
* 图表：用原生 `chart`（右键可编辑数据），不要贴图。
* 表格：用原生 `table`，显式清掉模板样式，颜色自控。
* 图标：优先用形状 + 文字；不要 emoji（跨平台字形不一致）。
* 字体：中文用微软雅黑 / 思源黑体；显式写 `latin/ea/cs` 三处，别依赖主题字体。

细则见 [docs/image-prep.md](docs/image-prep.md)。

---

## 8. 校验（交付前必跑）

```bash
pptx-native lint deck.pptx --slide 3            # 越界 / 溢出风险 / 重叠 / 字号
pptx-native export deck.pptx p3.png --slide 3   # PowerPoint 真机渲染（唯一验收依据）
pptx-native overlaps deck.pptx --slide 3        # 文字/形状重叠自检
pptx-native diff a.pptx b.pptx --slide 3        # 改版前后结构 diff
```

`lint` 是启发式体检（会给出 warning），用来在建完就发现问题；
它不替代真机导出——文字是否真的被裁掉，只有 PowerPoint 说了算。

三层：

| 层 | 手段 | 查什么 |
|---|---|---|
| 快 | LibreOffice / 预览渲染 | 溢出、错位、空白页 |
| 真 | PowerPoint COM 导出 | 3D、描边/阴影、字体替换、透明叠层 |
| 数 | 结构 diff + 像素量测 | 只改了预期区域 |

交付前逐项过一遍 [docs/verification.md](docs/verification.md) 的清单。

---

## 9. 反模式（看到就改）

* 一页塞 200 字、正文 10pt 以下 → 拆页、提层级。
* 5 种蓝、7 种字号 → 收敛到 5 个颜色角色、4 级字号。
* 标题写主题不写结论 → 改成一句可复述的结论。
* 图片拉伸变形 / 留黑边 / 加投影 → cover 裁切、去效果。
* 图表没标题、没单位、3D 饼图 → 加标题单位、换柱/折线。
* 动画逐条飞入、音效 → 除非演讲需要，一律删。
* 每页版式都不一样 → 同一用途固定用同一种版式。
* 页面元素贴边、不对齐 → 回到 12 栏网格。

完整清单见 [docs/anti-patterns.md](docs/anti-patterns.md)。

---

## 10. 踩坑速查

| 现象 | 原因 | 处理 |
|---|---|---|
| 元素整体偏移 0.1~0.2in | group 的 `chOff ≠ off` | `geometry.abs_rect()` |
| 预览对、真机歪 | group 有 `scene3d` 透视 | 装饰搬进同组 / 删 scene3d |
| 文字描边/阴影不显示 | 渲染器不支持文本效果 | `styling.layered_text` 叠层；看真机 |
| 文字变左/漂移 | `wrap=none` 时居中失效 | 打开 wrap，或加宽文本框 |
| PowerPoint 打不开生成的文件 | 形状 id 重复 / 抽页丢 rel | 全树重编号；`pptx-native preview` |
| 字体被换 / 字号变 | 主题字体 + autofit | 显式 typeface + `noAutofit` |
| 照片边缘一圈杂色 | 图片文件自带边框 | `pptx-native borders` |
| 中文乱码 / gbk 报错 | 控制台编码 | `PYTHONIOENCODING=utf-8` |
| 渲染报 maxInputBytes | deck 太大 | `preview` 抽单页再渲染 |
| CI 上 Python 3.9 装不上 | ubuntu-24.04 无对应 wheel | 固定 `ubuntu-22.04` |

---

## 11. 文档索引

* [docs/design-system.md](docs/design-system.md) —— 网格 / 字号 / 配色 / 间距的完整规范
* [docs/layouts.md](docs/layouts.md) —— 13 种版式的参数、几何与适用场景
* [docs/workflow.md](docs/workflow.md) —— 从需求到交付的完整流程与大纲模板
* [docs/native-editing.md](docs/native-editing.md) —— 改既有 deck：坐标、组合、3D、颜色清扫
* [docs/xml-cookbook.md](docs/xml-cookbook.md) —— 原始 OOXML 配方
* [docs/verification.md](docs/verification.md) —— 三层校验与交付清单
* [docs/image-prep.md](docs/image-prep.md) —— 图片素材处理
* [docs/anti-patterns.md](docs/anti-patterns.md) —— 常见坏设计对照表
* [docs/dsh-plugin.md](docs/dsh-plugin.md) —— DeepSeek Harness 插件（`pptx` 工具）
