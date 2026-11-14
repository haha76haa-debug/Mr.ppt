# Changelog

所有重要改动记录在这里。格式参考 [Keep a Changelog](https://keepachangelog.com/)，
版本号遵循 [SemVer](https://semver.org/lang/zh-CN/)。

## [0.2.0] - 2026-11-14

首个公开版本：从「单页复刻工具」泛化为「通用 PPT 设计系统 + 版式库 + 校验工具链」。

### Added

- `theme.py`：设计系统——颜色角色、字号阶梯、12 栏网格、4 套主题预设（blue / dark / minimal / warm）。
- `layouts.py`：13 种原生版式——`cover` / `agenda` / `section` / `bullets` / `two_col` /
  `stats` / `timeline` / `process` / `image_grid` / `quote` / `table` / `chart` / `end`。
- `sample.py`：中性 12 页示例 deck，全部内容为可替换占位。
- CLI 新增 `sample`、`themes` 命令。
- 测试：`test_theme.py`、`test_layouts.py`（版式生成、文字越界、图表/表格/图片墙）。
- 文档：`design-system.md`、`layouts.md`、`workflow.md`、`native-editing.md`、`anti-patterns.md`。
- `SKILL.md` 重写为通用 PPT 工作流（做新 deck / 复刻单页 / 改既有 deck 三类任务）。

### Changed

- 示例策略：不再内置任何图片素材，示例只依赖代码生成的占位图与占位框。
- `.gitignore`：不再对 `docs/*.png` 开白名单，仓库不存二进制产物。
- `pyproject.toml`：版本 0.2.0，描述更新为「设计系统 + 版式库 + 校验工具链」。

### Removed

- 旧版单页演示脚本（被 `sample` + `layouts` 取代）。

## [0.1.0] - 2026-11-13

内部迭代：单页复刻的工具集基础。

### Added

- `geometry.py`：组合坐标换算（`abs_rect`、`to_child_space`）。
- `styling.py`：字体（latin/ea/cs）、渐变、半透明、叠层文字、竖排。
- `groups.py`：把形状搬进组合、克隆并全树重编号。
- `images.py`：cover 裁切、截图补边、图片自带边框修复。
- `colors.py` / `checks.py`：颜色普查与清扫、快照 diff、重叠自检。
- `probe.py` / `render.py`：页面探针、抽页预览、PowerPoint COM 真机导出。
- CLI：`probe` / `preview` / `export` / `colors` / `sweep` / `borders` / `diff` / `overlaps`。
- `scripts/selfcheck.py`、`scripts/export_slide.ps1`、GitHub Actions CI。
