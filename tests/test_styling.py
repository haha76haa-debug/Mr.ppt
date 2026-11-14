# -*- coding: utf-8 -*-
"""样式：渐变/半透明/叠层文字/竖排。"""
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches

from pptx_native.styling import (grad, layered_text, set_alpha_fill, set_font,
                                 shape, textbox, vertical_label)

A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _deck():
    prs = Presentation()
    return prs, prs.slides.add_slide(prs.slide_layouts[6])


def test_gradient_fill_written():
    prs, slide = _deck()
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    grad(sp, [(0.0, (0x00, 0x30, 0xE6)), (0.5, (0xFF, 0xFF, 0xFF)), (1.0, (0x01, 0x35, 0xE7))])
    assert sp._element.spPr.find(qn("a:gradFill")) is not None
    stops = sp._element.spPr.findall(".//{%s}gs" % A)
    assert len(stops) == 3


def test_alpha_fill():
    prs, slide = _deck()
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    set_alpha_fill(sp, RGBColor(0x04, 0x5D, 0xE2), 0.78)
    alpha = sp._element.spPr.find(".//{%s}alpha" % A)
    assert alpha is not None and int(alpha.get("val")) == 78000


def test_set_font_writes_ea_typeface():
    prs, slide = _deck()
    tf, p = textbox(slide, 1, 1, 3, 0.5, align=PP_ALIGN.CENTER)
    r = p.add_run()
    r.text = "测试"
    rPr = set_font(r, 18, RGBColor(0, 0, 0), name="微软雅黑", spc=2.5)
    ea = rPr.find(qn("a:ea"))
    assert ea is not None and ea.get("typeface") == "微软雅黑"
    assert rPr.get("spc") == "250"


def test_layered_text_creates_three_layers():
    prs, slide = _deck()
    before = len(slide.shapes)
    layered_text(slide, 1, 1, 4, 0.8, "标题",
                 size=28, color=RGBColor(0x0A, 0x2C, 0x5A))
    assert len(slide.shapes) - before == 3        # 阴影 + 描边 + 彩色


def test_vertical_label_one_char_per_paragraph():
    prs, slide = _deck()
    tf = vertical_label(slide, 1, 1, 0.4, 2, "班长支书", size=13, color=RGBColor(255, 255, 255))
    assert [p.text for p in tf.paragraphs] == list("班长支书")


def test_shape_without_shadow_has_empty_effect():
    prs, slide = _deck()
    sp = shape(slide, MSO_SHAPE.RECTANGLE, 1, 1, 2, 1, fill=RGBColor(0, 0, 255))
    eff = sp._element.spPr.find(qn("a:effectLst"))
    assert eff is not None and len(eff) == 0
