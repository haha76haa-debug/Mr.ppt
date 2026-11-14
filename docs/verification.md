# 校验：怎么确认「做对了」

核心原则：**用户看到的是 PowerPoint（或 WPS），不是你看到的预览。**
LibreOffice / 在线预览只用来快速发现排版问题，不能当验收依据。

## 1. 三层校验

| 层 | 手段 | 查什么 | 耗时 |
|---|---|---|---|
| 快 | LibreOffice / 预览渲染 | 溢出、错位、空白页、明显缺字 | 秒级 |
| 真 | PowerPoint COM 导出 PNG | 3D 透视、文本描边/阴影、透明叠层、字体替换 | 十秒级 |
| 数 | 结构 diff + 像素量测 | 只改了预期区域、贴边误差、颜色替换是否彻底 | 秒级 |

**任何交付前必须跑「真」+「数」。** 只跑「快」一定会漏。

## 2. PowerPoint 真机导出

```bash
pptx-native export deck.pptx out.png --slide 8
# 等价脚本（Windows）：
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/export_slide.ps1 `
  -Path "C:\path\deck.pptx" -Out "C:\temp\p8.png" -Slide 8
```

要点：

* 必须 `Visible = msoTrue`；`WithWindow=false` 时大文件常报
  `PowerPoint could not open the file.`
* 100MB+ 先抽单页：`pptx-native preview deck.pptx one.pptx --keep 8`（会正确 drop rels）。
* 导出尺寸 `2001×1126`（13.333in @150dpi）够做像素比对。
* 结束后 `Quit()` 并释放 COM，否则 `POWERPNT.exe` 残留会锁文件。
* CI/无 Office 环境跑不了这一步，只能做结构级测试（见 `.github/workflows/ci.yml`）。

## 3. 结构自检

```bash
pptx-native lint deck.pptx --slide 8          # 页面体检（越界/溢出/重叠/字号）
pptx-native overlaps deck.pptx --slide 8      # 重叠对（面积、双方 id/名称）
pptx-native diff before.pptx after.pptx --slide 8
pptx-native probe deck.pptx --slide 8         # 形状清单（含绝对坐标）
```

`lint` 返回四类问题：

| kind | 含义 | 级别 |
|---|---|---|
| `canvas` | 元素超出画布 | 文字 = error，装饰/图片 = warn |
| `overflow` | 文字可能撑破文本框（按字符宽度 + 字号估算） | warn |
| `overlap` | 两个文字框互相重叠 | warn |
| `font` | 显式字号小于 9pt | warn |

它是启发式检查，用来在构建后立刻发现问题；**不能替代真机渲染**——
文字是不是真的被裁掉、字体有没有被替换，只有 PowerPoint 说了算。

期望：

* 只改颜色/换图 → `diff` 的 `GEOM = 0`、`TEXT = 0`；
* 新增元素 → 形状数 +N，其余几何不变；
* 未改动的页 XML 完全相同：

```python
assert prs_a.slides[0]._element.xml == prs_b.slides[0]._element.xml
```

## 4. 像素校验

```python
from collections import Counter
import numpy as np
from PIL import Image

arr = np.asarray(Image.open("out.png").convert("RGB"))
sub = arr[y0:y1, x0:x1]
print(Counter(map(tuple, sub.reshape(-1, 3))).most_common(5))   # 颜色是否到位

d = np.abs(arr_a.astype(int) - arr_b.astype(int)).sum(axis=2)
print("changed px", (d > 30).sum())                              # 只应出现在预期区域
```

* 改色后，目标色应成为该区域的主色；
* 改版前后差异只应出现在改动区域，整页都在变说明改错了东西；
* 半透明叠加区域的像素值取决于背景，不要用单一 RGB 判断。

## 5. 贴边/对齐量测

```python
from pptx_native.geometry import abs_rect

px, py, pw, ph = abs_rect(photo)
bx, by, bw, bh = abs_rect(bar)
assert abs((by + bh) - (py + ph)) < 0.004     # 底边贴合（0.004in ≈ 0.1mm）
assert abs(bx - px) < 0.004
```

## 6. 交付核对

```powershell
Copy-Item out.pptx "D:\Desktop\交付.pptx" -Force
Get-Item "D:\Desktop\交付.pptx" | Select-Object Length, LastWriteTime
```

* 大小、时间与刚生成的源文件一致；
* 目标被 PowerPoint 打开时写不进去，先让用户关闭；
* 交付说明里写清：改了哪些页、怎么验证、已知限制。

## 7. 常见「预览对、真机不对」

| 现象 | 原因 | 对策 |
|---|---|---|
| 照片斜了 / 梯形 | group 上 `scene3d` 透视 | 同组 / 删 scene3d，真机看 |
| 装饰条不贴照片 | 装饰在组外，不跟 3D 变换 | 搬进同组 |
| 文字描边/阴影没出现 | LibreOffice 不支持文本效果 | 叠层文字；只用 PowerPoint 验收 |
| 字体被替换 | 系统没装该字体 | 换常见字体，或内嵌/随文件分发 |
| 中文变方框 | 渲染器缺中文字体 | PowerPoint 导出可避免 |
| 半透明变灰/变实 | 预览器不支持 alpha | 真机看，或改用叠层色块 |
| 生成的文件打不开 | shape id 重复 / 抽页丢 rel | 全树重编号 / 用 `preview` |

## 8. 每页 30 秒自检

- [ ] 标题是结论
- [ ] 无文字溢出 / 无贴边 / 无意外重叠
- [ ] 字号、颜色只来自 Theme
- [ ] 图片无拉伸、无留白边
- [ ] 页码、页脚一致
- [ ] 图表有标题/单位；表格不超 5 列
