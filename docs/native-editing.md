# 改造既有 deck：坐标、组合、3D、颜色

任务形态：用户给一个现成 `.pptx`（或截图），要求「改成可编辑的」「统一主题色」
「给照片加注释」「移一下位置」，但不许重画整份。原则：**先探明结构，再动刀；
每动一刀都能量化对比。**

## 1. 先探明结构

```bash
pptx-native probe deck.pptx --slide 6          # 形状清单：id/类型/名称/绝对坐标/文字
pptx-native probe deck.pptx --slide 6 --json s6.json
pptx-native colors deck.pptx --slide 6         # 颜色普查（含主题色、预设色、红色）
pptx-native overlaps deck.pptx --slide 6       # 有没有已经重叠的元素
```

`probe` 输出的是**渲染坐标**（已换算组合变换），不是 python-pptx 的原始 `left/top`。
先看三件事：形状总数、层级（是否在组里）、颜色是否来自主题。

## 2. 组合坐标：最容易错的 0.13in

`p:grpSpPr/a:xfrm` 同时有 `off/ext` 与 `chOff/chExt`：

```
abs = off + (child - chOff) * (ext / chExt)
```

* `off/ext`：组合在页面上的真实位置与缩放。
* `chOff/chExt`：组合「内部坐标系」的原点与尺寸。
* `child`：子元素自己记的坐标（python-pptx 的 `shape.left/top` 读的就是它）。

模板里 `chOff ≠ off` 很常见（PowerPoint 编辑过组合就会这样），差 0.1~0.2in。
**所有对齐/贴合计算都必须先 `geometry.abs_rect(shape)`。**

反向：把页面坐标写回子元素空间用 `geometry.to_child_space(group_el, x, y, w, h)`。

把元素搬进组合（保持视觉位置不变、继承组合变换）：

```python
from pptx_native import groups
groups.move_into_group(bar, photo, flush="bottom")   # 注释条贴照片底边并进组
```

## 3. `scene3d` 透视：预览对、真机歪

模板可能给整组照片挂透视相机：

```xml
<a:scene3d><a:camera prst="perspectiveRight">
  <a:rot lat="0" lon="20099999" rev="0"/>
</a:camera><a:lightRig rig="threePt" dir="t"/></a:scene3d>
```

PowerPoint 会把整组按透视投影（照片变斜/梯形），LibreOffice 直接忽略 →
**预览看起来贴合，真机是歪的**。

两种处理：

1. **推荐**：把需要贴合的元素（注释条、边框、角标）搬进同一个组，继承同一套 2D/3D 变换。
2. 确认后删掉 `scene3d` 让页面变平（会改变设计观感，必须问用户）。

判断方法：`pptx-native export` 导出 PNG，和 LibreOffice 渲染对比。

## 4. 颜色：红/杂色为什么会残留

颜色可能藏在 6 个地方，只改形状填充一定漏：

| 位置 | XML |
|---|---|
| 形状填充 | `p:spPr/a:solidFill` 或 `a:gradFill` |
| 形状描边 | `p:spPr/a:ln/a:solidFill` |
| 文字填充 | `a:rPr/a:solidFill` 或 `a:gradFill` |
| 文字描边 | `a:rPr/a:ln/a:solidFill` |
| 文字/形状阴影 | `a:effectLst/a:outerShdw/a:srgbClr` |
| 主题色 | `a:schemeClr` 引用（改主题或显式覆盖） |
| 图片来源 | 照片像素本身（JPEG 自带边框、红模板留的虚线） |

```bash
pptx-native colors deck.pptx --slide 8              # 普查
pptx-native sweep deck.pptx out.pptx --slide 8 --to 2E7BE8   # 统一换色
pptx-native borders deck.pptx out.pptx --slide 8 --to 1B3E78 # 照片自带边框
```

规则：

* 只改用户指定的页；删页后页码会变，先确认。
* 换成同一「语义」的颜色：标题用主色、次要信息用浅一档，不要全刷成同一个蓝。
* 改完用 `diff` 确认：`GEOM`/`TEXT` 变化必须为 0，只有颜色/图片字节变。
* 半透明不要改成不透明：`sweep` 保留 alpha。

## 5. 字体与主题继承

* 新加的文字显式写字体（latin/ea/cs），否则会跟主题字体走。
* 既有文字被主题字体覆盖 → 在 `rPr` 上写 `a:latin/a:ea/a:cs` 的 `typeface`。
* 字号被 autofit 缩小 → 加 `<a:noAutofit/>`（`styling.no_autofit`）。
* 批量替换字体可遍历 `p:cNvPr` 对应的 `rPr`，但**不要**改数字 run 的字体（保留等宽数字观感）。

## 6. 复制页面 / 形状

```python
from pptx_native import groups
groups.clone_shape(src_shape, dst_slide, start_id)   # 深拷贝并全树重编号
```

* 同一页所有 `p:cNvPr@id` 必须唯一，否则 PowerPoint 报「无法读取」。
* 复制整页后，图片/媒体关系要重建（`add_picture` 重新引用），不能只拷贝 XML。
* 新建形状取「当前最大 id + 1」→ **先建新形状，再插入复制组**，否则 id 撞车。

## 7. 两版对比

```bash
pptx-native diff before.pptx after.pptx --slide 6
```

输出：形状数变化、只在一侧出现的 id、几何变化、文字变化。
「只改颜色」的场景下四类都应为 0；有 GEOM 说明动了不该动的坐标。

## 8. 交付核对

```powershell
Copy-Item out.pptx "D:\Desktop\交付.pptx" -Force
Get-Item "D:\Desktop\交付.pptx" | Select-Object Length, LastWriteTime
```

* 大小/时间与刚生成的一致；
* 目标文件被 PowerPoint 打开时会写不进去，先让用户关掉；
* 报告写清：改了哪些页、哪些元素、用了什么方法验证、有什么副作用。
