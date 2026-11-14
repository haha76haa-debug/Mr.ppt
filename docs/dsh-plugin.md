# DeepSeek Harness 插件（PPT Studio）

把工具链打包成一个 DSH **Host 插件**：装一次，agent 就有一个 `pptx` 工具，
不用每次写 shell / 临时脚本。

* 源码在 [`dsh-plugin/`](../dsh-plugin/)（清单、patch、Host 代码、locale、图标）
* `python/` 由构建脚本从 `src/pptx_native/` 生成，**不在仓库里重复提交**
* 方法论仍在 [SKILL.md](../SKILL.md) 与 [docs/](.)：插件是执行壳，skill 是知识

## 构建

```bash
python scripts/build_plugin.py                # → build/pptx-studio-bundle/
python scripts/build_plugin.py --out DIR      # 指定输出目录（例如已安装的 bundle 目录）
```

构建把 `src/pptx_native/`、`scripts/export_slide.ps1` 复制进 bundle，并校验
`package.json`、patch、Python 包、导出脚本都在。

## 安装

* **Plugins → Add plugin** → 选 `build/pptx-studio-bundle` 的绝对路径；
* 或在 agent 里用插件管理接口 `install_bundle`，`target` 传同一个绝对路径。

安装后 **Plugins** 里会出现 `pptx-studio` 一行；工具 `pptx` 出现在模型工具列表里。

## 工具

| action | 作用 | 必需参数 |
|---|---|---|
| `sample` | 生成 12 页中性示例 deck | `out`，可选 `theme` |
| `probe` | 列出某页全部形状（组合展开、绝对坐标） | `deck`，可选 `slide` |
| `lint` | 页面体检：越界 / 文字溢出风险 / 文字框重叠 / 字号过小 | `deck`，可选 `slide` |
| `render` | 用 PowerPoint COM 真机导出某页 PNG | `deck`、`out`，可选 `slide` |
| `themes` | 列出主题预设 | — |

路径用绝对路径。`render` 需要 Windows + Microsoft PowerPoint；其余跨平台。

`lint` 输出示例：

```
[warn] canvas   #3 Oval 2   元素超出画布右/下边界（right=15.83, bottom=9.60）
[error] canvas  #7 TextBox  文字框超出画布边界
[warn] overlap  #9 TextBox  文字框与 #8「标题」重叠 0.21 in²
[warn] font     #11 TextBox 字号 8.0pt 小于建议值 9pt
```

它是启发式检查（只提醒，不替代真机渲染）；`--fail-on error` 可让 CLI 在
存在 error 时返回非零退出码。

## 运行时

Python 查找顺序：插件配置 `python` → `DSH_PYTHON` →
`$DSH_HOME/dsh-runtimes/*/dependencies/python/python.exe` → `PATH` 的
`python` / `py -3` / `python3`；取第一个能 `import pptx, PIL, numpy, lxml` 的。

固定解释器（编辑插件行 config）：

```yaml
- id: pptx-studio
  name: '@local/pptx-studio'
  config:
    python: 'C:\path\to\python.exe'
```

工具执行 `python -m pptx_native`（`PYTHONPATH=<bundle>/python`），不联网、不装包。

## 更新

改了 `src/` 或 `dsh-plugin/` 之后：

```bash
python scripts/build_plugin.py --out "<已安装的 bundle 目录>"
```

然后**重启 DSH**（或移除后重新安装 bundle），让 Host 加载新的 JavaScript 代码。
只改 Python 代码时，重启同样最稳妥。

## 限制

* `render` 仅 Windows + PowerPoint；
* 工具直接写调用方给的路径，**不经过会话沙箱**——传 workspace 内的路径；
* 单次调用 180s 超时；
* 本地 bundle（`@local/...`）不发布到 npm registry；
* 工具结果的图片目前以路径返回，需要 `read_image` 再读一次
  （内联图片需要接入 attachment 服务，暂未做）。

## 卸载

**Plugins** → 移除 `pptx-studio`。除了你让工具写过的文件，不会留下别的东西。
