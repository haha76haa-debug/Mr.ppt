# -*- coding: utf-8 -*-
"""素材：cover 裁切、补边、照片自带边框修复。"""
from PIL import Image

from pptx_native.images import cover, encode, pad, repaint_border_lines, sample_bg


def _img(w=120, h=90, colour=(200, 210, 220)):
    return Image.new("RGB", (w, h), colour)


def test_cover_fills_slot():
    src = _img(120, 90)
    out = cover(src, 200, 100)          # 目标比例 2:1，源 4:3 → 裁上下
    assert out.size == (200, 100)


def test_pad_keeps_full_picture():
    src = _img(120, 90)
    out = pad(src, 200, 200)
    assert out.size == (200, 200)
    bg = sample_bg(src)
    assert out.getpixel((199, 199)) == bg


def test_repaint_border_lines_removes_red_frame(photo_bytes):
    import io
    img = Image.open(io.BytesIO(photo_bytes)).convert("RGB")
    new, n = repaint_border_lines(img, target=(0x1B, 0x3E, 0x78))
    assert n > 0
    pixel = new.getpixel((0, 0))
    assert pixel[2] > pixel[0], "边框应变成蓝色系"
    # 中间内容不受影响
    assert new.getpixel((60, 45)) == img.getpixel((60, 45))


def test_encode_formats():
    img = _img(20, 20)
    assert encode(img, "png")[:4] == b"\x89PNG"
    assert encode(img, "jpg")[:2] == b"\xff\xd8"
