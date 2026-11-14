# AGENT.md

给在本仓库工作的 AI agent 的规则。目标：**改完能跑、能验、能交付**。

## 开始前

1. 读 [SKILL.md](SKILL.md)（工作流）与 [PROJECT.md](PROJECT.md)（设计约束）。
2. 看 `git status` 和 `git log --oneline -5`，不要覆盖未提交的改动。
3. 确认改动属于哪一层：设计系统 / 版式 / 工具 / 文档。跨层的改动先写计划。

## 硬规则

* **不引入个人数据。** 仓库里不要出现真实姓名、班级、单位、内部数据；
  示例一律用中性占位（「示例演示文稿」「128 个」）。
* **不提交生成物。** `*.pptx` / `*.pdf` / `*.png` 一律进 `.gitignore`；
  文档需要图就现场渲染，不往仓库放二进制。
* **新能力必须带测试。** 版式函数至少验证：能生成、形状数合理、文字在画布内。
* **版式函数保持纯粹**：只接收 `(prs, theme, **内容)`，不读全局状态、不写文件。
* **颜色/字体只从 Theme 取**，不在版式代码里写裸色值（`RGBColor(0x...)`）。
* **改 XML 时注意元素顺序**（见 [docs/xml-cookbook.md](docs/xml-cookbook.md) 第 1 节）。
* **中文注释可以**，但公开 API 的 docstring 保持简短、可英文阅读。

## 每次改动后

```bash
pytest -q                     # CI 用的标准测试
python scripts/selfcheck.py   # 没有 pytest 时的纯 assert 版
python -m pptx_native sample build/sample.pptx
python -m pptx_native lint build/sample.pptx --slide 1
python scripts/build_plugin.py  # DSH 插件 bundle 能构建
```

涉及渲染 / 版式的改动，在 Windows 上再做一次真机检查：

```powershell
python -m pptx_native export build/sample.pptx build/p1.png --slide 1
```

看三件事：文字有没有溢出、元素有没有越界、和设计意图是否一致。

## 交付检查清单

- [ ] `pytest -q` 全绿（或 selfcheck 全 PASS）
- [ ] `pptx-native sample` 能生成且能重新打开
- [ ] `pptx-native lint` 在示例页上无 error
- [ ] 改了 `dsh-plugin/` 或 `src/` 后 `python scripts/build_plugin.py` 能通过
- [ ] 新增/修改的版式有测试
- [ ] README / SKILL / docs 的命令与实际 CLI 一致
- [ ] `CHANGELOG.md` 有对应条目
- [ ] 没有把生成物、个人数据、临时文件提交进去

## 不要做

* ❌ 直接改 `build/` 里的产物当交付（产物可重建）
* ❌ 在既有 pptx 上反复打补丁而不修构建脚本
* ❌ 用截图冒充原生元素
* ❌ 未经确认删除用户内容（数据/文案/页面）
* ❌ 把令牌、路径、用户名写进代码或文档
