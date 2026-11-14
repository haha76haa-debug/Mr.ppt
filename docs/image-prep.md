# 图片素材处理

原则：**PPT 里的照片只有一个正确形态——充满框、不拉伸、不留边、不加效果。**
工具在 `src/pptx_native/images.py`。

## 1. 照片：cover 裁切

槽位 `w × h`（英寸），输出像素按 300dpi 足够（投影 1920 宽已超需求）：

```python
from pptx_native.images import cover, encode
img = cover(Image.open("photo.jpg"), int(w * 300), int(h * 300))
data = encode(img, "jpeg", 92)          # bytes，可直接 add_picture
```

裁切方向：素材比槽位「更宽」→ 裁左右；「更高」→ 裁上下。默认居中；
带人脸的素材可以按锚点上移，保住头部。

**反面例子（都踩过）：**

* `add_picture` 只给一个宽高 → 按原比例缩，留白边；
* 先 pad 再放 → 看起来像两张图；
* 加阴影 / 模糊补边 → 投影上一团脏，用户会说「不要阴影」。

## 2. 截图：整图可见 + 背景补边

聊天记录、表格、网页截图不能裁（会切掉正文）。做法：

1. 缩放到槽位宽度；
2. 采样截图四条边的中位色作为背景色；
3. 新建槽位尺寸画布，填背景色，把截图贴顶部居中；
4. 空出的部分正好被注释条/底板盖住。

```python
from pptx_native.images import pad
img = pad(shot, int(w * 300), int(h * 300))
```

## 3. 透明素材（logo / 图标）

* 输出 RGBA PNG；
* **合成前先铺底色再贴**，否则透明区变黑（Pillow 的 `convert("RGB")` 会丢 alpha）：

```python
canvas = Image.new("RGB", size, BG)
canvas.paste(logo_rgba, pos, logo_rgba)   # 用 alpha 当 mask
```

* 圆形裁切用 4× 超采样 mask 再缩小，边缘才平滑；
* SVG 先转 PNG（500px 以上），不要直接放 SVG（PowerPoint 版本兼容不稳）。

## 4. 多图拼一张

* 统一高度横向拼，留 2~4px 浅色分隔；
* 拼完的比例必须等于槽位比例，否则又回到拉伸/留白；
* 超过 4 张就别拼，改用 `layouts.image_grid`。

## 5. 图片文件自带边框

模板导出的 JPEG 有时最外 1~2px 有一圈深色/红色边，缩略图看不出，投影上很明显。

```bash
pptx-native borders deck.pptx out.pptx --slide 3 --to 1B3E78
```

判据：整行/整列平均色偏某色，且位于最外 5 行内（逐像素阈值会被 JPEG 噪声干扰，
必须按行/列均值判）。只改边框像素，不动内容。

## 6. 尺寸与体积

| 用途 | 建议像素宽 | 说明 |
|---|---|---|
| 全屏背景图 | 1920~2560 | 16:9，JPEG q85 |
| 半屏配图 | 1200 | |
| 图片墙单格 | 700~900 | 槽位 × 300dpi |
| 图标/logo | 500（透明 PNG） | |

* 20 张 1000px 照片约 +20MB；再大对投影没有意义。
* 重新编码 JPEG 用 `quality=90~95`，肉眼看不出、体积可控。
* deck 超过 100MB 时，很多渲染/预览工具会直接报错。

## 7. 素材检查清单

- [ ] 每张图都按槽位比例裁切，没有拉伸
- [ ] 没有留白边 / 黑边
- [ ] 截图没被裁掉正文
- [ ] 照片没有自带边框/水印残留
- [ ] 同一页图片的色调/亮度大致统一（必要时统一调色）
- [ ] 图片文件名可读（`team-01.jpg` 而不是 `IMG_8842.JPG`）
