# OOXML 配方速查（python-pptx + lxml）

命名空间统一：

```python
from pptx.oxml.ns import qn
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
```

## 1. 插入位置（最关键）

`spPr` 的子元素顺序有严格要求，乱插 PowerPoint 会报「无法读取」：

```
xfrm → prstGeom → (fill) → ln → effectLst
```

所以渐变/纯色这类 fill **必须紧跟在 `prstGeom` 后面**：

```python
spPr.find(qn("a:prstGeom")).addnext(fill_el)
```

`rPr` 里顺序是：`a:ln → solidFill → effectLst`。

## 2. 多段渐变（可带透明度）

```xml
<a:gradFill rotWithShape="1">
  <a:gsLst>
    <a:gs pos="0"><a:srgbClr val="0030E6"/></a:gs>
    <a:gs pos="56000"><a:srgbClr val="FFFFFF"/></a:gs>
    <a:gs pos="100000"><a:srgbClr val="0135E7"/></a:gs>
  </a:gsLst>
  <a:lin ang="0" scaled="0"/>
</a:gradFill>
```

`pos` 0~100000；`ang` 单位 1/60000 度（0 = 左→右，5400000 = 上→下）。
半透明在 `srgbClr` 里加 `<a:alpha val="78000"/>`（78%）。

## 3. 纯色 + 透明度（磨砂注释条）

```xml
<a:solidFill><a:srgbClr val="0B52C4"><a:alpha val="80000"/></a:srgbClr></a:solidFill>
```

## 4. 去掉主题默认阴影

`shape.shadow.inherit = False` → 写入空 `<a:effectLst/>`。
空 effectLst 表示「显式无效果」，覆盖 `p:style` 里 `effectRef` 带来的阴影。

## 5. 文字描边 / 文字阴影

```xml
<a:rPr sz="2400" b="1">
  <a:ln w="32480"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>
  <a:solidFill><a:srgbClr val="0B52C4"/></a:solidFill>
  <a:effectLst>
    <a:outerShdw blurRad="106299" dist="35431" dir="2700000" rotWithShape="0">
      <a:srgbClr val="12325F"><a:alpha val="29000"/></a:srgbClr>
    </a:outerShdw>
  </a:effectLst>
</a:rPr>
```

PowerPoint 认，LibreOffice 不认 → 跨渲染器场景改用叠层文字。

## 6. 叠层文字（跨渲染器一致的描边/阴影）

用真形状叠出来，三层：

1. 最底层：同文字、阴影色、偏移 1~2pt（阴影）
2. 中间层：同文字、白/底色、加粗描边（描边）
3. 最上层：真文字、目标色（用户编辑这一层）

```python
styling.layered_text(slide, x, y, w, h, "标题", size=28,
                     color=RGBColor(0x0B, 0x52, 0xC4))
```

代价是改字要改三处（或重新生成）；只有确实需要才用，纯色文字不要叠。

## 7. 竖排文字

**不要**用 `<a:bodyPr vert="eaVert"/>` / `vert270`（不同渲染器侧躺/倒置）。
一字一段，行距压到 0.9~1.0：

```python
tf = styling.vertical_label(slide, x, y, 0.4, 2.0, "竖排示例",
                            size=13, color=RGBColor(0xFF, 0xFF, 0xFF))
```

## 8. 组合（group）坐标

```xml
<p:grpSpPr><a:xfrm>
  <a:off x="114673" y="1557663"/> <a:ext cx="1860220" cy="3290690"/>
  <a:chOff x="114673" y="1681068"/> <a:chExt cx="1860220" cy="3290690"/>
</a:xfrm></p:grpSpPr>
```

渲染位置：`abs = off + (child - chOff) * (ext / chExt)`。
`child` 是 python-pptx 读到的 `left/top`（raw），**不等于**渲染位置。
用 `pptx_native.geometry.abs_rect()` 读；`groups.move_into_group()` 写。

## 9. 3D 透视相机（scene3d）

```xml
<a:scene3d>
  <a:camera prst="perspectiveRight"><a:rot lat="0" lon="20099999" rev="0"/></a:camera>
  <a:lightRig rig="threePt" dir="t"/>
</a:scene3d>
```

挂在 `p:grpSp/p:grpSpPr` 上时，PowerPoint 把整组按透视投影，LibreOffice 忽略。
处理：要贴合的装饰搬进同组；或（确认后）删掉 scene3d 让页面变平。

## 10. 图片填充

```xml
<p:pic>
  <p:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
  <p:spPr>
    <a:xfrm><a:off .../><a:ext .../></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

`stretch/fillRect` = 把图拉伸到框；配合事先 cover 裁好的素材，效果就是 cover。
要圆角/圆形：换 `prstGeom prst="roundRect"` / `"ellipse"`，但图片本身会被裁。

## 11. 表格：去掉模板样式

```python
tblPr = table._tbl.tblPr
tblPr.set("firstRow", "0")
tblPr.set("bandRow", "0")
for el in tblPr.findall(qn("a:tableStyleId")):
    tblPr.remove(el)
```

然后逐格 `cell.fill` + 显式写 `a:lnL/lnR/lnT/lnB`（见 `layouts._cell_borders`），
否则会残留模板的竖线/底色。

## 12. 文本框内边距

`a:bodyPr` 不写 `lIns/rIns/tIns/bIns` 时用主题默认值（左右各 0.1in），
窄框会让居中漂移。要精确对齐就显式设 0，并加 `<a:noAutofit/>`：

```python
tf.margin_left = tf.margin_right = 0
tf.word_wrap = True
styling.no_autofit(tf)
```

## 13. 形状 id 不可重复

同一页所有 `p:cNvPr@id` 必须唯一。复制组后全树重编号：

```python
n = max(int(e.get("id")) for e in slide._element.iter(qn("p:cNvPr")))
for c in new_el.iter(qn("p:cNvPr")):
    n += 1
    c.set("id", str(n))
```

`groups.clone_shape()` 已封装。注意 python-pptx 新建形状取「当前最大 id + 1」，
所以**先建新形状，再插入复制组**。

## 14. 主题色引用

```xml
<a:schemeClr val="accent1"><a:lumMod val="60000"/></a:schemeClr>
```

模板里大量元素是 `schemeClr`，改 `theme1.xml` 的 `a:clrScheme` 会整份 deck 变色。
只想改一页时，把该页的 `schemeClr` 显式替换成 `srgbClr` 更可控。
