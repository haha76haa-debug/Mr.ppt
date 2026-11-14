# -*- coding: utf-8 -*-
"""素材预处理：cover 裁切、截图补边、照片自带边框修复。

三条实战规则：
  1. 照片一律 cover（充满槽位，不留边、不拉伸）
  2. 截图一律 pad（整图可见，用自身背景色补边）
  3. 素材输出像素尺寸 = PPT 槽位尺寸（宽高比必须一致，误差 < 0.5%）
"""
from __future__ import annotations

import io

import numpy as np
from PIL import Image

BORDER_SCAN = 5          # 检测自带边框时，每个边看的最多像素行/列


def sample_bg(img):
    """采样四周边缘的中位色，作为截图补边颜色。"""
    w, h = img.size
    pts = []
    for x in range(0, w, max(1, w // 24)):
        for y in (1, 2, h - 3, h - 2):
            pts.append(img.getpixel((x, y)))
    for y in range(0, h, max(1, h // 24)):
        for x in (1, 2, w - 3, w - 2):
            pts.append(img.getpixel((x, y)))
    pts = [(p[0], p[1], p[2]) if isinstance(p, tuple) else (p, p, p) for p in pts]
    return tuple(sorted(v[i] for v in pts)[len(pts) // 2] for i in range(3))


def cover(img, w, h):
    """按比例裁切填满 w×h（居中裁）。"""
    src_r, dst_r = img.width / img.height, w / h
    if src_r > dst_r:
        nw, nh = int(round(img.height * dst_r)), img.height
        x0 = (img.width - nw) // 2
        box = (x0, 0, x0 + nw, nh)
    else:
        nw, nh = img.width, int(round(img.width / dst_r))
        y0 = (img.height - nh) // 2
        box = (0, y0, nw, y0 + nh)
    return img.crop(box).resize((w, h), Image.LANCZOS)


def pad(img, w, h):
    """整图可见 + 自身背景色补边（截图用）。"""
    bg = sample_bg(img)
    scale = w / img.width
    img = img.resize((w, max(1, int(img.height * scale))), Image.LANCZOS)
    canvas = Image.new("RGB", (w, h), bg)
    canvas.paste(img, (0, 0))
    return canvas


def flatten(img, bg=(255, 255, 255)):
    """把 RGBA 合成到纯色底上（Pillow 直接 convert('RGB') 会把透明处变黑）。"""
    canvas = Image.new("RGB", img.size, bg)
    canvas.paste(img, (0, 0), img if img.mode == "RGBA" else None)
    return canvas


def _line_mean(arr, idx, axis):
    return arr[idx, :].mean(axis=0) if axis == 1 else arr[:, idx].mean(axis=0)


def _is_red(mean_rgb):
    r, g, b = mean_rgb
    return r > 60 and r >= g + 30 and r >= b + 30


def repaint_border_lines(img, target=(0x1B, 0x3E, 0x78), scan=BORDER_SCAN):
    """把图片最外圈自带的细边框（如红模板遗留的 1~2px 红线）重涂成目标色。

    返回 (新图, 重涂的行/列数)。只动边框像素，不影响画面内容。
    """
    a = np.array(img.convert("RGB")).astype(np.uint8)
    h, w = a.shape[:2]
    target = np.array(target, dtype=np.uint8)
    painted = 0

    sides = (
        (1, [range(scan), range(h - 1, h - 1 - scan, -1)]),      # 上、下 行
        (0, [range(scan), range(w - 1, w - 1 - scan, -1)]),      # 左、右 列
    )
    for axis, orders in sides:
        for order in orders:
            order = list(order)
            start = next((i for i, idx in enumerate(order) if _is_red(_line_mean(a, idx, axis))), None)
            if start is None:
                continue
            for idx in order[start:start + 3]:
                m = _line_mean(a, idx, axis)
                if _line_mean(a, order[start], axis)[0] > 0 and (
                        idx == order[start] or (m[0] >= m[1] + 15 and m[0] >= m[2] + 15)):
                    if axis == 1:
                        a[idx, :] = target
                    else:
                        a[:, idx] = target
                    painted += 1
                else:
                    break
    return Image.fromarray(a), painted


def encode(img, ext="png", quality=95):
    """按目标格式编码成 bytes（JPEG 用高质量、4:4:4 采样）。"""
    buf = io.BytesIO()
    if ext.lower() in ("jpg", "jpeg"):
        img.convert("RGB").save(buf, "JPEG", quality=quality, subsampling=1)
    else:
        img.save(buf, "PNG")
    return buf.getvalue()
