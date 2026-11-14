# PROJECT.md — 设计说明

## 问题定义

用 python-pptx 或任何库「生成 PPT」都不难，难在生成物要同时满足：

1. **可编辑**：文字是文本框、形状是形状、图表可改数据、图片可换；
2. **好看且一致**：多页之间不能各写各的，必须有统一网格 / 字号 / 配色；
3. **真的对**：用户在看 PowerPoint，而你在看预览器——两者会不一样；
4. **改得动**：拿到一个既有 deck 要改配色/加注释，不能重画。

本项目的回答：**把设计约束写成代码（Theme + Layouts），把渲染差异写成校验工具。**

## 分层

| 层 | 模块 | 职责 |
|---|---|---|
| 设计层 | `theme.py` | 颜色角色、字号阶梯、12 栏网格、主题预设 |
| 版式层 | `layouts.py` | 13 种页面原型，只依赖 Theme |
| 样式层 | `styling.py` | 字体（latin/ea/cs）、渐变、半透明、叠层文字、竖排 |
| 几何层 | `geometry.py` | 组合坐标换算、绝对 bbox、贴合判断 |
| 结构层 | `groups.py` | 形状搬进/搬出组合、克隆重编号 |
| 素材层 | `images.py` | cover 裁切、截图补边、图片自带边框修复 |
| 审计层 | `colors.py` / `checks.py` | 颜色普查与清扫、快照 diff、重叠自检 |
| 观测层 | `probe.py` / `render.py` | 页面探针、抽页预览、PowerPoint COM 导出 |
| 入口层 | `sample.py` / `cli.py` | 示例 deck、命令行 |

依赖方向：上层依赖下层；`theme`/`geometry` 不依赖业务；`cli` 只做参数解析。

## 关键设计决策

### 1. 为什么 Theme 是 frozen dataclass

主题是「一组常量」，不是可变状态。冻结 + `with_fonts()` / `from_dict()` 派生，
保证换主题不会在运行时被某页偷偷改掉；也方便序列化进配置。

### 2. 为什么版式统一签名 `(prs, theme, **内容)`

* 换主题不改版式；
* 测试可以逐个调用、检查形状数与越界；
* 用户想改某一页，直接复制函数体改坐标，比对着 GUI 调更快。

### 3. 为什么 `abs_rect` 是核心

OOXML 里 group 的子元素坐标是「child 空间」的，渲染时按
`abs = off + (child - chOff) * (ext / chExt)` 映射。模板里 `chOff ≠ off` 很常见，
python-pptx 的 `shape.left/top` 读的是 raw 值——所有对齐计算必须过 `abs_rect()`。

### 4. 为什么「搬进组合」而不是「算透视」

组上的 `a:scene3d` 是透视相机，硬算投影矩阵复杂且易错。把要贴合的装饰搬进同组，
它继承同一套 2D/3D 变换，任何渲染器下都贴合，一次解决。

### 5. 为什么验收必须用 PowerPoint COM

LibreOffice / 在线预览不支持：组合 3D、文字描边、部分文字阴影、内嵌字体。
这些恰好是「看起来对/不对」的根源。所以 `export` 用 COM 真机导出，作为唯一验收图。

### 6. 为什么颜色清扫要遍历整页 XML

颜色可能藏在：形状 fill、形状 line、文字 fill、文字 line、文字/形状 effectLst、
`schemeClr` 主题引用、图片像素。只看 `shape.fill` 一定漏。`colors.scan_red`
遍历整页并标注上下文，`sweep_red` 保留 alpha 地替换。

## 边界（不做什么）

* 不做批量套模板 / 数据合并（用 python-pptx 或 docxtpl 更合适）；
* 不做动画与切换（API 支持弱、收益低）；
* 不做 PPT 转 HTML/图片的通用渲染（`render` 只导出单页 PNG）；
* 不自带字体文件（用户自行安装或用系统自带字体）。

## 质量门槛

| 门槛 | 手段 |
|---|---|
| 单测 | `pytest`，无需 PowerPoint；覆盖几何、样式、颜色、图片、版式、主题 |
| 无 pytest 环境 | `scripts/selfcheck.py`（纯 assert） |
| CI | Python 3.9 / 3.10 / 3.11 / 3.12，Ubuntu |
| 快速渲染 | LibreOffice（只查溢出） |
| 真机验收 | PowerPoint COM 导出（Windows 手动） |
| 结构回归 | `diff`：只改颜色时 GEOM/TEXT 必须为 0 |

## 路线

- [x] 生成前的「内容溢出预测」——`checks.audit_slide()` / `lint`（0.2.1）
- [ ] `layouts.py` 增加 `grid` / `cards` / `gantt` 版式
- [ ] `Theme` 支持从 PPTX 主题（`theme1.xml`）反读
- [ ] `probe --json` 输出可直接喂给布局重算（从截图到参数）
- [ ] 跨平台渲染（LibreOffice → PNG）一键化
- [ ] DSH 插件：工具结果内联渲染图——需接 `ctx.attachments` 把 PNG 转成 `ImageBlock`，
      省掉一次 `read_image`；等高频生成/逐页验收时再做
- [ ] DSH 插件：UI 面板（缩略图墙 / 主题切换）——Client 插件 + slot + React，
      面向非 agent 使用场景，当前与侧边栏预览、`office_render` 重复，暂不做

