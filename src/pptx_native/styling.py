# -*- coding: utf-8 -*-
"""样式原语：字体、渐变、半透明、叠层文字、形状。

约定：坐标与尺寸均为英寸；颜色用 ``RGBColor``。
"""
from __future__ import annotations

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
DEFAULT_FONT = "微软雅黑"


# --------------------------------------------------------------------- 字体
def set_font(run, size, color, bold=True, name=DEFAULT_FONT, spc=None, italic=False):
    """设置一个 run（显式写 latin/ea/cs，避免主题字体覆盖中文）。"""
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    run.font.name = name
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set("typeface", name)
    if spc is not None:                       # 字间距：pt * 100
        rPr.set("spc", str(int(spc * 100)))
    return rPr


def no_autofit(text_frame):
    """固定字号，禁止 PowerPoint/WPS 自动缩放。"""
    bodyPr = text_frame._txBody.find(qn("a:bodyPr"))
    for tag in ("a:normAutofit", "a:spAutoFit", "a:noAutofit"):
        for el in bodyPr.findall(qn(tag)):
            bodyPr.remove(el)
    bodyPr.append(bodyPr.makeelement(qn("a:noAutofit"), {}))
    return bodyPr


def textbox(slide, x, y, w, h, *, anchor=MSO_ANCHOR.MIDDLE, wrap=True,
            margins=(0, 0, 0, 0), align=PP_ALIGN.LEFT, line_spacing=1.0):
    """新建文本框，返回 (text_frame, first_paragraph)。"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.margin_left, tf.margin_right = Inches(margins[0]), Inches(margins[1])
    tf.margin_top, tf.margin_bottom = Inches(margins[2]), Inches(margins[3])
    tf.vertical_anchor = anchor
    no_autofit(tf)
    p = tf.paragraphs[0]
    p.alignment = align
    p.line_spacing = line_spacing
    p.space_before = Pt(0)
    p.space_after = Pt(0)
    return tf, p


def fill_text(text_frame, lines):
    """lines = [[{"t": "文字", "size": 12, "c": RGBColor, ...}, ...], ...]"""
    first = True
    for runs in lines:
        p = text_frame.paragraphs[0] if first else text_frame.add_paragraph()
        first = False
        for rd in runs:
            r = p.add_run()
            r.text = rd["t"]
            set_font(r, rd.get("size", 14), rd.get("c", RGBColor(0, 0, 0)),
                     rd.get("b", True), rd.get("f", DEFAULT_FONT),
                     rd.get("spc"), rd.get("i", False))
    return text_frame


def vertical_label(slide, x, y, w, h, text, *, size, color,
                   name=DEFAULT_FONT, line_spacing=0.94):
    """竖排文字：一个字一个段落（不要用 bodyPr@vert，跨渲染器会歪/倒）。"""
    tf, p = textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.MIDDLE,
                    align=PP_ALIGN.CENTER, line_spacing=line_spacing)
    first = True
    for ch in text:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = ch
        set_font(r, size, color, True, name)
    return tf


# --------------------------------------------------------------------- 形状
def shape(slide, kind, x, y, w, h, fill=None, line=None, line_w=1.0,
          radius=None, dash=None, no_shadow=True):
    """新建形状；默认去掉主题阴影（用户通常明确要求无阴影）。"""
    sp = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    if no_shadow:
        sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
        if dash:
            sp.line.dash_style = dash
    if radius is not None:
        sp.adjustments[0] = radius
    return sp


def set_alpha_fill(sp, rgb, alpha=0.78):
    """纯色 + 透明度（0~1）。替换已有填充，插在 prstGeom 之后（OOXML 顺序要求）。"""
    spPr = sp._element.spPr
    for el in list(spPr.findall(qn("a:solidFill"))) + list(spPr.findall(qn("a:noFill"))):
        spPr.remove(el)
    fill = etree.fromstring(
        f'<a:solidFill xmlns:a="{A}"><a:srgbClr val="{str(rgb)}">'
        f'<a:alpha val="{int(alpha * 100000)}"/></a:srgbClr></a:solidFill>')
    spPr.find(qn("a:prstGeom")).addnext(fill)


def grad(sp, stops, angle=0.0):
    """多段渐变。stops = [(pos0~1, (r,g,b) | (r,g,b,alpha0~255)), ...]"""
    spPr = sp._element.spPr
    for tag in ("a:solidFill", "a:noFill", "a:gradFill"):
        for el in list(spPr.findall(qn(tag))):
            spPr.remove(el)
    parts = []
    for pos, c in stops:
        if len(c) == 4:
            col = (f'<a:srgbClr val="{str(RGBColor(c[0], c[1], c[2]))}">'
                   f'<a:alpha val="{int(c[3] / 255 * 100000)}"/></a:srgbClr>')
        else:
            col = f'<a:srgbClr val="{str(RGBColor(*c))}"/>'
        parts.append(f'<a:gs pos="{int(pos * 100000)}">{col}</a:gs>')
    el = etree.fromstring(
        f'<a:gradFill xmlns:a="{A}" rotWithShape="1"><a:gsLst>{"".join(parts)}</a:gsLst>'
        f'<a:lin ang="{int(angle * 60000)}" scaled="0"/></a:gradFill>')
    spPr.find(qn("a:prstGeom")).addnext(el)


def add_shadow(sp, blur=10, dist=3, color="0A2C5A", alpha=16000):
    spPr = sp._element.spPr
    for el in spPr.findall(qn("a:effectLst")):
        spPr.remove(el)
    spPr.append(etree.fromstring(
        f'<a:effectLst xmlns:a="{A}"><a:outerShdw blurRad="{int(blur * 12700)}" '
        f'dist="{int(dist * 12700)}" dir="5400000" rotWithShape="0">'
        f'<a:srgbClr val="{color}"><a:alpha val="{alpha}"/></a:srgbClr>'
        f"</a:outerShdw></a:effectLst>"))


# --------------------------------------------------------------------- 叠层文字
def outline_run(run, color, width_pt):
    """文字描边（PowerPoint 认，LibreOffice 不认）。"""
    rPr = run._r.get_or_add_rPr()
    ln = etree.fromstring(
        f'<a:ln xmlns:a="{A}" w="{int(width_pt * 12700)}"><a:solidFill>'
        f'<a:srgbClr val="{str(color)}"/></a:solidFill></a:ln>')
    rPr.insert(0, ln)


def layered_text(slide, x, y, w, h, text, *, size, color,
                 font=DEFAULT_FONT,
                 outline_color=RGBColor(0xFF, 0xFF, 0xFF), outline_w=6.0,
                 shadow_color=RGBColor(0x80, 0x80, 0x80), shadow_dx=0.018,
                 align=PP_ALIGN.CENTER):
    """跨渲染器一致的「描边 + 阴影」文字：用真形状叠出来。

    层次：阴影层 → 描边层（白字+粗描边）→ 彩色真文字。
    返回最上层文本框；用户改字改这一层。
    """
    def _layer(dx, dy, col, with_outline=False):
        tf, p = textbox(slide, x + dx, y + dy, w, h, align=align)
        r = p.add_run()
        r.text = text
        set_font(r, size, col, True, font)
        if with_outline:
            outline_run(r, outline_color, outline_w)
        return tf

    _layer(shadow_dx, shadow_dx, shadow_color)
    _layer(0, 0, outline_color, with_outline=True)
    return _layer(0, 0, color)
