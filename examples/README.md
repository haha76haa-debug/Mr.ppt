# examples

最小可运行示例，不依赖任何外部素材。

## build_sample_deck.py

用版式库搭一份 12 页示例 deck，覆盖所有版式：

```bash
python examples/build_sample_deck.py build/sample.pptx
# 等价命令
pptx-native sample build/sample.pptx
pptx-native sample build/sample-dark.pptx --theme dark
```

## 自己改一份

```python
from pptx_native import layouts
from pptx_native.theme import Theme

theme = Theme.dark()
prs = layouts.new_deck(theme)
layouts.cover(prs, theme, title="你的标题", subtitle="副标题")
layouts.bullets(prs, theme, title="要点", items=[("标题一", "一句说明")])
prs.save("my.pptx")
```

版式参数见 [docs/layouts.md](../docs/layouts.md)，完整流程见
[docs/workflow.md](../docs/workflow.md)。
