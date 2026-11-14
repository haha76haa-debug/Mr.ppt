# Contributing to pptx-native

感谢贡献！这个项目来自「程序化做 PPT 时反复被同一批坑绊倒」的实践，
任何「又踩到一个坑」的补充都很欢迎。

## 开发环境

```bash
git clone https://github.com/haha76haa-debug/Mr.ppt.git
cd Mr.ppt
python -m venv .venv
# Windows: .venv\Scripts\activate    macOS/Linux: source .venv/bin/activate
pip install -e ".[dev]"
```

## 跑测试

```bash
pytest -q                     # 标准方式
python scripts/selfcheck.py   # 没装 pytest / 受限环境（纯 assert 版）
```

CI 在 Ubuntu + Python 3.9/3.10/3.11/3.12 上跑 pytest。
PowerPoint COM 相关功能只在 Windows 手动验证（CI 没有 Office）。

## 提交规范

`type(scope): 简述`，type ∈ `feat|fix|docs|test|chore|refactor`。

示例：

- `feat(layouts): 新增 gantt 版式`
- `fix(geometry): 处理嵌套 group 的多级缩放`
- `docs: 补充 scene3d 透视的排查方法`

## 代码约定

* 只在 `src/pptx_native/` 里加库代码；脚本放 `scripts/`，示例放 `examples/`。
* 新能力必须带测试；几何/颜色/图片改动优先复用现有模块。
* 中文注释可以；**不要在仓库里写入个人或组织隐私数据**。
* 不要提交大文件（pptx/png/jpeg）；示例素材用代码生成。

## 提 Issue

- 🐞 Bug：附最小可复现 deck（脱敏）、期望 vs 实际、`pptx-native probe` 输出片段。
- 💡 Feature：说明使用场景与期望的 CLI/API 形态。

## Pull Request

1. 关联 Issue（如有）
2. 说明改了什么、怎么验证的
3. 附测试结果（`pytest` / `selfcheck` 输出）
4. 更新 `CHANGELOG.md` 和相应文档
