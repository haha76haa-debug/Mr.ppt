# -*- coding: utf-8 -*-
"""pytest fixtures：全部用程序生成，不依赖外部素材，可在 CI 上跑。"""
from __future__ import annotations

import io

import pytest
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

E = 914400


@pytest.fixture()
def photo_bytes():
    """120x90 测试图，四周 2px 红框（模拟红模板遗留的图片边框）。"""
    img = Image.new("RGB", (120, 90), (200, 210, 220))
    red = (122, 29, 24)
    for i in range(2):
        for x in range(120):
            img.putpixel((x, i), red)
            img.putpixel((x, 89 - i), red)
        for y in range(90):
            img.putpixel((i, y), red)
            img.putpixel((119 - i, y), red)
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def _add_group_with_picture(slide, pic_el, rect=(2, 2, 2, 1.5), ch_shift=0.135):
    """把已存在的图片元素搬进一个 group，并造出 chOff ≠ off 的模板状态。"""
    x, y, w, h = rect
    grp = slide.shapes.add_group_shape()
    g_el = grp._element
    pic_el.getparent().remove(pic_el)
    g_el.append(pic_el)

    # 子元素在 child 空间的位置 = 目标位置 + shift
    xf = pic_el.spPr.find(qn("a:xfrm"))
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    off.set("x", str(int((x) * E)))
    off.set("y", str(int((y + ch_shift) * E)))
    ext.set("cx", str(int(w * E)))
    ext.set("cy", str(int(h * E)))

    gxf = g_el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    for tag, attrs in (
            ("a:off", {"x": int(x * E), "y": int(y * E)}),
            ("a:ext", {"cx": int(w * E), "cy": int(h * E)}),
            ("a:chOff", {"x": int(x * E), "y": int((y + ch_shift) * E)}),
            ("a:chExt", {"cx": int(w * E), "cy": int(h * E)})):
        el = gxf.find(qn(tag))
        for k, v in attrs.items():
            el.set(k, str(int(v)))
    return grp, pic_el


@pytest.fixture()
def deck(tmp_path, photo_bytes):
    """最小演示 deck。

    第 1 页：红字 + 红块 + 带红框照片（用于颜色清扫 / 边框修复测试）
    第 2 页：group(chOff≠off) 内照片 + 组外注释条（用于几何 / 搬组测试）
    """
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    pic_path = tmp_path / "photo.png"
    pic_path.write_bytes(photo_bytes)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(0.6))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "红色残留"
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(0xC0, 0x03, 0x03)

    red_box = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2),
                                  Inches(2), Inches(1))
    red_box.fill.solid()
    red_box.fill.fore_color.rgb = RGBColor(0xD5, 0x1A, 0x01)
    red_box.shadow.inherit = False

    s1.shapes.add_picture(str(pic_path), Inches(1), Inches(4), Inches(2), Inches(1.5))
    # 一个与红块重叠的形状，用于 overlaps 测试
    s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(2.2), Inches(2), Inches(1.5))

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    pic = s2.shapes.add_picture(str(pic_path), Inches(2), Inches(2), Inches(2), Inches(1.5))
    _add_group_with_picture(s2, pic._element)

    bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(3.15),
                              Inches(2), Inches(0.35))
    bar.name = "Rectangle 1"
    bar.text_frame.text = "注释"

    path = tmp_path / "deck.pptx"
    prs.save(str(path))
    return path


@pytest.fixture()
def deck_path(deck):
    return str(deck)
