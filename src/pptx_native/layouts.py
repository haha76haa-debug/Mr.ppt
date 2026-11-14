# -*- coding: utf-8 -*-
"""通用版式库：把一页 PPT 做成「参数化函数」。

每个函数统一签名 ``layout(prs, theme, **内容) -> slide``，只用原生形状和真文字：
* 文字是 TextBox（可直接改字）；
* 色块是 Shape（可直接改色）；
* 图表是原生 Chart（可直接改数据）；
* 照片是 Picture（右键即可替换）。

坐标系：左上角原点，单位英寸，画布 13.333 x 7.5（16:9）。
"""
from __future__ import annotations

import io
import math
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

from . import styling
from .images import cover as crop_cover, encode
from .styling import grad, set_alpha_fill, set_font, shape, textbox
from .theme import SLIDE_H, SLIDE_W, Theme

# 版面纵向节奏（英寸）
TITLE_Y = 0.62
BODY_TOP = 1.82
BODY_BOTTOM = 6.82

#: 图表 / 统计数字的配色循环（固定，避免每页配色漂移）
_ACCENTS = ("primary", "primary_light", "accent")


# --------------------------------------------------------------------- 基础件
def new_deck(theme: Theme | None = None) -> Presentation:
    """新建 16:9 空白 deck（沿用 python-pptx 默认模板，不含占位符版式）。"""
    _ = theme                                    # 保留参数，方便以后套模板
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def add_slide(prs, theme: Theme, page: int | None = None):
    """新建空白页并铺背景色；``page`` 给页码，None 不显示。"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = theme.rgb("bg")
    if page is not None:
        _text(slide, SLIDE_W - theme.margin - 0.9, SLIDE_H - 0.62, 0.9, 0.3,
              "%02d" % page, theme.size("caption"), theme.rgb("muted"),
              font=theme.font_body, bold=False, align=PP_ALIGN.RIGHT)
    return slide


def _text(slide, x, y, w, h, text, size, color, *, font=None, bold=True,
          align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spc=None, line_spacing=1.0):
    """单段文字（一个跑）。返回 run，方便追加。"""
    if not text:
        return None
    tf, p = textbox(slide, x, y, w, h, anchor=anchor, align=align,
                    line_spacing=line_spacing)
    run = p.add_run()
    run.text = str(text)
    set_font(run, size, color, bold, font or styling.DEFAULT_FONT, spc)
    return run


def _paragraphs(slide, x, y, w, h, lines, size, color, *, font=None,
                line_spacing=1.3, bold=False):
    """多段文字（每行一段）。"""
    lines = [ln for ln in lines if ln]
    if not lines:
        return None
    tf, first = textbox(slide, x, y, w, h, anchor=MSO_ANCHOR.TOP,
                        line_spacing=line_spacing)
    for i, line in enumerate(lines):
        p = first if i == 0 else tf.add_paragraph()
        p.line_spacing = line_spacing
        r = p.add_run()
        r.text = str(line)
        set_font(r, size, color, bold, font or styling.DEFAULT_FONT)
    return tf


def _header(slide, theme: Theme, title: str, kicker: str | None = None):
    """页眉：小标签 + 标题 + 短强调线。"""
    if kicker:
        _text(slide, theme.margin, TITLE_Y - 0.38, theme.content_width, 0.28,
              kicker, theme.size("caption") + 1, theme.rgb("primary"),
              font=theme.font_body, spc=2.2)
    _text(slide, theme.margin, TITLE_Y, theme.content_width - 1.0, 0.66,
          title, theme.size("h1"), theme.rgb("text"), font=theme.font_title)
    shape(slide, MSO_SHAPE.RECTANGLE, theme.margin, TITLE_Y + 0.8,
          0.72, 0.05, fill=theme.rgb("primary"))


def _hairline(slide, theme: Theme, x, y, w, color=None):
    shape(slide, MSO_SHAPE.RECTANGLE, x, y, w, 0.012,
          fill=theme.rgb(color or "line"))


def _card(slide, theme: Theme, x, y, w, h, *, fill=None, line=None, radius=0.045):
    return shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                 fill=theme.rgb(fill or "surface"),
                 line=theme.rgb(line) if line else None,
                 line_w=1.0, radius=radius)


def _dot(slide, theme: Theme, cx, cy, d=0.18, *, fill="primary", line="bg", line_w=2.0):
    return shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d,
                 fill=theme.rgb(fill), line=theme.rgb(line), line_w=line_w)


def _load_image(source):
    """路径 / bytes / BytesIO / PIL.Image → RGB 的 PIL.Image（None 表示占位）。"""
    if source is None:
        return None
    if isinstance(source, Image.Image):
        return source.convert("RGB")
    if isinstance(source, (bytes, bytearray)):
        return Image.open(io.BytesIO(source)).convert("RGB")
    return Image.open(str(Path(source))).convert("RGB")


# --------------------------------------------------------------------- 版式
def cover(prs, theme: Theme, *, title: str, subtitle: str = "", meta: str = "",
          kicker: str = "", page: int | None = None):
    """封面：深色渐变底 + 主标题 + 副标题 + 底部信息。"""
    slide = add_slide(prs, theme, page)

    bg = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
               fill=theme.rgb("primary_dark"))
    grad(bg, [(0.0, theme.tuple("primary_dark")), (0.58, theme.tuple("primary")),
              (1.0, theme.tuple("primary_light"))], 38)

    # 右下装饰：半透明圆 + 细环（克制，不抢字）
    halo = shape(slide, MSO_SHAPE.OVAL, SLIDE_W - 4.1, SLIDE_H - 4.5, 6.6, 6.6,
                 fill=theme.rgb("primary_light"))
    set_alpha_fill(halo, theme.rgb("primary_light"), 0.20)
    shape(slide, MSO_SHAPE.OVAL, SLIDE_W - 2.55, SLIDE_H - 2.85, 3.9, 3.9,
          fill=None, line=theme.rgb("on_dark"), line_w=1.0)

    if kicker:
        _text(slide, theme.margin, 2.30, 6.4, 0.32, kicker, theme.size("small"),
              theme.rgb("on_dark"), font=theme.font_body, spc=3.0)
    _text(slide, theme.margin, 2.66, 8.2, 1.55, title, theme.size("hero"),
          theme.rgb("on_dark"), font=theme.font_title)
    shape(slide, MSO_SHAPE.RECTANGLE, theme.margin, 4.36, 1.15, 0.06,
          fill=theme.rgb("accent"))
    if subtitle:
        _text(slide, theme.margin, 4.60, 7.6, 0.8, subtitle, theme.size("lead"),
              theme.rgb("on_dark"), font=theme.font_body, bold=False,
              line_spacing=1.25)
    if meta:
        _text(slide, theme.margin, SLIDE_H - 0.92, 9.0, 0.32, meta,
              theme.size("caption") + 1, theme.rgb("on_dark"),
              font=theme.font_body, bold=False)
    return slide


def section(prs, theme: Theme, *, number: int, title: str, subtitle: str = "",
            page: int | None = None):
    """章节过渡页：大序号 + 章节名。"""
    slide = add_slide(prs, theme, page)
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, 4.25, SLIDE_H, fill=theme.rgb("surface"))
    shape(slide, MSO_SHAPE.RECTANGLE, 4.25, 0, 0.05, SLIDE_H, fill=theme.rgb("primary"))

    _text(slide, theme.margin, 2.38, 3.2, 1.7, "%02d" % number,
          theme.size("display"), theme.rgb("primary_light"), font=theme.font_title)
    _text(slide, 4.78, 2.62, 7.7, 0.9, title, theme.size("h1"),
          theme.rgb("text"), font=theme.font_title)
    if subtitle:
        _text(slide, 4.78, 3.58, 7.3, 0.9, subtitle, theme.size("lead"),
              theme.rgb("muted"), font=theme.font_body, bold=False,
              line_spacing=1.3)
    return slide


def agenda(prs, theme: Theme, *, title: str, items, kicker: str = "议程",
           page: int | None = None):
    """目录页：编号 + 条目，1~4 项单列，5 项起分两列。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    items = list(items)
    cols = 1 if len(items) <= 4 else 2
    rows = math.ceil(len(items) / cols)
    row_h = (BODY_BOTTOM - BODY_TOP) / rows
    cell_w = theme.content_width / cols

    for i, item in enumerate(items):
        col, row = divmod(i, rows) if cols == 2 else (0, i)
        x = theme.margin + col * cell_w
        y = BODY_TOP + row * row_h
        _text(slide, x, y + (row_h - 0.55) / 2, 0.78, 0.55, "%02d" % (i + 1),
              theme.size("h3"), theme.rgb("primary_light"), font=theme.font_title)
        _text(slide, x + 0.82, y + (row_h - 0.55) / 2, cell_w - 1.1, 0.55, item,
              theme.size("lead") - 1, theme.rgb("text"), font=theme.font_body)
        if row < rows - 1 or cols == 2:
            _hairline(slide, theme, x, y + row_h - 0.1, cell_w - 0.32)
    return slide


def bullets(prs, theme: Theme, *, title: str, items, kicker: str = "",
            page: int | None = None):
    """要点页：每行「小标题（可带一句说明）」，行距均分。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    items = list(items)
    row_h = (BODY_BOTTOM - BODY_TOP) / max(len(items), 1)
    for i, item in enumerate(items):
        head, body = (item, "") if isinstance(item, str) else item
        y = BODY_TOP + i * row_h
        shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, theme.margin + 0.04, y + 0.13,
              0.16, 0.16, fill=theme.rgb("primary"), radius=0.2)
        _text(slide, theme.margin + 0.44, y, theme.content_width - 0.6, 0.46, head,
              theme.size("h3"), theme.rgb("text"), font=theme.font_title)
        if body:
            _paragraphs(slide, theme.margin + 0.44, y + 0.48,
                        theme.content_width - 0.6, row_h - 0.6, [body],
                        theme.size("small") + 1, theme.rgb("muted"),
                        font=theme.font_body, line_spacing=1.35)
        if i < len(items) - 1:
            _hairline(slide, theme, theme.margin, y + row_h - 0.12,
                      theme.content_width)
    return slide


def two_col(prs, theme: Theme, *, title: str, left, right, kicker: str = "",
            page: int | None = None):
    """对比页：左右两张卡片，参数为 ``(标题, 正文或正文行列表)``。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    card_w = (theme.content_width - theme.gutter) / 2
    card_h = BODY_BOTTOM - BODY_TOP - 0.1
    for i, (head, body) in enumerate((left, right)):
        x = theme.margin + i * (card_w + theme.gutter)
        _card(slide, theme, x, BODY_TOP, card_w, card_h)
        shape(slide, MSO_SHAPE.RECTANGLE, x, BODY_TOP, 0.09, card_h,
              fill=theme.rgb(_ACCENTS[i] if i < len(_ACCENTS) else "primary"))
        _text(slide, x + 0.42, BODY_TOP + 0.38, card_w - 0.8, 0.5, head,
              theme.size("h3"), theme.rgb("text"), font=theme.font_title)
        lines = body if isinstance(body, (list, tuple)) else [body]
        _paragraphs(slide, x + 0.42, BODY_TOP + 1.02, card_w - 0.84,
                    card_h - 1.4, lines, theme.size("body"), theme.rgb("muted"),
                    font=theme.font_body, line_spacing=1.45)
    return slide


def stats(prs, theme: Theme, *, title: str, items, note: str = "", kicker: str = "",
          page: int | None = None):
    """数据页：3~5 个关键数字并排（数字大、标签小，主次分明）。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    items = list(items)
    cell_w = theme.content_width / max(len(items), 1)
    for i, item in enumerate(items):
        value, unit, label = (item[0], "", item[1]) if len(item) == 2 else item
        x = theme.margin + i * cell_w
        shape(slide, MSO_SHAPE.RECTANGLE, x, BODY_TOP + 0.05,
              cell_w - 0.34, 0.05,
              fill=theme.rgb(_ACCENTS[i % len(_ACCENTS)]))
        tf, p = textbox(slide, x, BODY_TOP + 0.35, cell_w - 0.22, 0.95,
                        anchor=MSO_ANCHOR.BOTTOM)
        r1 = p.add_run()
        r1.text = str(value)
        set_font(r1, theme.size("stat"), theme.rgb("primary"), True, theme.font_title)
        if unit:
            r2 = p.add_run()
            r2.text = " " + str(unit)
            set_font(r2, theme.size("stat_unit"), theme.rgb("muted"), False,
                     theme.font_body)
        _text(slide, x, BODY_TOP + 1.45, cell_w - 0.22, 0.34, label,
              theme.size("small"), theme.rgb("muted"), font=theme.font_body)
    if note:
        _text(slide, theme.margin, BODY_BOTTOM - 0.62, theme.content_width, 0.3,
              note, theme.size("caption"), theme.rgb("muted"),
              font=theme.font_body, bold=False)
    return slide


def timeline(prs, theme: Theme, *, title: str, items, kicker: str = "",
             page: int | None = None):
    """时间线：一条横线 + 均分节点，参数为 ``(时间, 标题, 说明)``。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    items = list(items)
    x0 = theme.margin + 0.2
    x1 = SLIDE_W - theme.margin - 0.2
    line_y = 3.55
    shape(slide, MSO_SHAPE.RECTANGLE, x0, line_y, x1 - x0, 0.035,
          fill=theme.rgb("primary_light"))
    slot = (x1 - x0) / max(len(items), 1)
    box_w = slot - 0.14

    for i, item in enumerate(items):
        when, head, body = (item[0], "", item[1]) if len(item) == 2 else item
        cx = x0 + (i + 0.5) * slot
        bx = cx - box_w / 2
        if when:
            _text(slide, bx, line_y - 0.78, box_w, 0.34, when,
                  theme.size("small"), theme.rgb("primary"), font=theme.font_body,
                  align=PP_ALIGN.CENTER)
        _dot(slide, theme, cx, line_y + 0.02, 0.2, fill="bg", line="primary", line_w=3.0)
        _text(slide, bx, line_y + 0.34, box_w, 0.42, head,
              theme.size("h3"), theme.rgb("text"), font=theme.font_title,
              align=PP_ALIGN.CENTER)
        if body:
            _paragraphs(slide, bx, line_y + 0.82, box_w, 1.3, [body],
                        theme.size("small"), theme.rgb("muted"),
                        font=theme.font_body, line_spacing=1.25)
    return slide


def process(prs, theme: Theme, *, title: str, steps, kicker: str = "",
            page: int | None = None):
    """流程页：N 步横向卡片（序号徽标 + 小标题 + 说明）。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    steps = list(steps)
    gap = 0.3
    card_w = (theme.content_width - (len(steps) - 1) * gap) / max(len(steps), 1)
    card_h = 2.35
    y = BODY_TOP + 0.3

    for i, (head, body) in enumerate(steps):
        x = theme.margin + i * (card_w + gap)
        _card(slide, theme, x, y, card_w, card_h)
        badge = shape(slide, MSO_SHAPE.OVAL, x + 0.34, y + 0.32, 0.62, 0.62,
                      fill=theme.rgb("primary"))
        tf = badge.text_frame
        tf.word_wrap = False
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(i + 1)
        set_font(r, theme.size("h3"), theme.rgb("on_dark"), True, theme.font_title)
        _text(slide, x + 0.34, y + 1.1, card_w - 0.68, 0.42, head,
              theme.size("h3"), theme.rgb("text"), font=theme.font_title)
        if body:
            _paragraphs(slide, x + 0.34, y + 1.56, card_w - 0.68, 0.72, [body],
                        theme.size("small"), theme.rgb("muted"),
                        font=theme.font_body, line_spacing=1.25)
        if i < len(steps) - 1:
            shape(slide, MSO_SHAPE.RIGHT_ARROW, x + card_w + 0.045,
                  y + card_h / 2 - 0.12, 0.21, 0.24,
                  fill=theme.rgb("primary_light"))
    return slide


def image_grid(prs, theme: Theme, *, title: str, items, cols: int = 3,
               captions: bool = True, gap: float = 0.22, kicker: str = "",
               page: int | None = None):
    """图片墙：等分网格，``items`` 为 ``(图片路径/bytes/None, 注释)``。

    图片按 cover 裁切填满格子（不拉伸、不留边）；``None`` 生成空占位框，
    方便先排版、后让用户自己插图。
    """
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    items = list(items)
    rows = math.ceil(len(items) / cols)
    avail_h = BODY_BOTTOM - BODY_TOP - 0.25
    cell_w = (theme.content_width - (cols - 1) * gap) / cols
    cell_h = min((avail_h - (rows - 1) * gap) / rows, cell_w * 0.74)
    grid_h = rows * cell_h + (rows - 1) * gap
    top = BODY_TOP + max(0.0, (avail_h - grid_h) / 2)

    for i, item in enumerate(items):
        source, caption = item
        row, col = divmod(i, cols)
        x = theme.margin + col * (cell_w + gap)
        y = top + row * (cell_h + gap)
        _photo_cell(slide, theme, x, y, cell_w, cell_h, source,
                    caption if captions else "")
    return slide


def _photo_cell(slide, theme: Theme, x, y, w, h, source, caption):
    img = _load_image(source)
    if img is None:
        pic = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h,
                    fill=theme.rgb("surface"), line=theme.rgb("primary_light"),
                    line_w=1.0, radius=0.03)
        pic.line.dash_style = MSO_LINE.DASH
        _text(slide, x, y + h / 2 - 0.22, w, 0.44, "图片占位",
              theme.size("small"), theme.rgb("muted"), font=theme.font_body,
              align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    else:
        px_w, px_h = max(int(w * 300), 32), max(int(h * 300), 32)
        data = io.BytesIO(encode(crop_cover(img, px_w, px_h), "jpeg", 92))
        pic = slide.shapes.add_picture(data, Inches(x), Inches(y),
                                       Inches(w), Inches(h))
    if caption:
        bar_h = max(0.34, min(0.46, h * 0.2))
        bar = shape(slide, MSO_SHAPE.RECTANGLE, x, y + h - bar_h, w, bar_h,
                    fill=theme.rgb("primary"))
        set_alpha_fill(bar, theme.rgb("primary"), 0.8)
        _text(slide, x, y + h - bar_h, w, bar_h, caption, theme.size("small"),
              theme.rgb("on_dark"), font=theme.font_body, align=PP_ALIGN.CENTER,
              anchor=MSO_ANCHOR.MIDDLE)
    return pic


def quote(prs, theme: Theme, *, text: str, author: str = "", kicker: str = "",
          page: int | None = None):
    """引用页：大字引言 + 引号 + 出处。"""
    slide = add_slide(prs, theme, page)
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H, fill=theme.rgb("surface"))
    if kicker:
        _text(slide, theme.margin, 1.34, 6.0, 0.3, kicker, theme.size("small"),
              theme.rgb("primary"), font=theme.font_body, spc=2.4)
    _text(slide, theme.margin - 0.06, 1.42, 3.0, 1.6, "“", 120,
          theme.rgb("primary_light"), font=theme.font_title)
    _text(slide, theme.margin + 0.1, 2.62, theme.content_width - 0.4, 2.7, text,
          theme.size("h2"), theme.rgb("text"), font=theme.font_title,
          bold=False, line_spacing=1.45)
    if author:
        shape(slide, MSO_SHAPE.RECTANGLE, theme.margin + 0.1, 5.72, 0.62, 0.05,
              fill=theme.rgb("primary"))
        _text(slide, theme.margin + 0.1, 5.9, theme.content_width - 0.4, 0.4,
              author, theme.size("small"), theme.rgb("muted"),
              font=theme.font_body)
    return slide


def table(prs, theme: Theme, *, title: str, header, rows, weight=None,
          note: str = "", kicker: str = "", page: int | None = None):
    """表格页：原生表格（可编辑），首行深色、隔行浅底、细下线。"""
    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    rows = list(rows)
    n_rows, n_cols = len(rows) + 1, len(header)
    table_h = min(BODY_BOTTOM - BODY_TOP - 0.45, 0.52 + 0.46 * len(rows))
    gf = slide.shapes.add_table(n_rows, n_cols, Inches(theme.margin),
                                Inches(BODY_TOP + 0.12),
                                Inches(theme.content_width), Inches(table_h))
    tbl = gf.table

    # 去掉模板表格样式，颜色全部自己控制
    tblPr = tbl._tbl.tblPr
    tblPr.set("firstRow", "0")
    tblPr.set("bandRow", "0")
    for el in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(el)

    weights = list(weight) if weight else [1] * n_cols
    total = float(sum(weights))
    for i, wgt in enumerate(weights):
        tbl.columns[i].width = Inches(theme.content_width * wgt / total)

    tbl.rows[0].height = Inches(0.52)
    body_row_h = (table_h - 0.52) / max(len(rows), 1)
    for r in range(1, n_rows):
        tbl.rows[r].height = Inches(body_row_h)

    for c, text in enumerate(header):
        cell = tbl.cell(0, c)
        cell.fill.solid()
        cell.fill.fore_color.rgb = theme.rgb("primary")
        _cell_borders(cell)
        _cell(slide, theme, cell, text, theme.size("small") + 1,
              theme.rgb("on_dark"), theme.font_title, bold=True)
    for r, row in enumerate(rows):
        for c, text in enumerate(row):
            cell = tbl.cell(r + 1, c)
            cell.fill.solid()
            cell.fill.fore_color.rgb = theme.rgb("bg" if r % 2 == 0 else "surface")
            _cell_borders(cell, theme.rgb("line"), 0.75)
            _cell(slide, theme, cell, text, theme.size("small") + 1,
                  theme.rgb("text"), theme.font_body, bold=(c == 0))
    if note:
        _text(slide, theme.margin, min(BODY_BOTTOM - 0.3, BODY_TOP + 0.12 + table_h + 0.16),
              theme.content_width, 0.3, note, theme.size("caption"),
              theme.rgb("muted"), font=theme.font_body, bold=False)
    return slide


def _cell(slide, theme: Theme, cell, text, size, color, font, bold=False):
    cell.margin_left = cell.margin_right = Inches(0.16)
    cell.margin_top = cell.margin_bottom = Inches(0.04)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = str(text)
    set_font(r, size, color, bold, font)


def _cell_borders(cell, color=None, bottom_pt=0.75):
    """清掉表格样式带的边框，只保留一条细下边线（全部显式写 XML）。"""
    tcPr = cell._tc.get_or_add_tcPr()
    tags = ("a:lnL", "a:lnR", "a:lnT", "a:lnB")
    for tag in tags:
        for el in tcPr.findall(qn(tag)):
            tcPr.remove(el)
    for i, tag in enumerate(tags):
        has_line = tag == "a:lnB" and color is not None
        ln = tcPr.makeelement(qn(tag), {"w": str(int((bottom_pt if has_line else 0) * 12700))})
        if has_line:
            fill = ln.makeelement(qn("a:solidFill"), {})
            fill.append(ln.makeelement(qn("a:srgbClr"), {"val": str(color)}))
            ln.append(fill)
        else:
            ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.insert(i, ln)


def chart(prs, theme: Theme, *, title: str, categories, series, kind: str = "column",
          note: str = "", kicker: str = "", page: int | None = None):
    """图表页：原生柱状 / 条形 / 折线图（右键即可改数据）。"""
    from pptx.chart.data import CategoryChartData

    slide = add_slide(prs, theme, page)
    _header(slide, theme, title, kicker)

    data = CategoryChartData()
    data.categories = list(categories)
    for name, values in series:
        data.add_series(name, list(values))

    kinds = {"column": XL_CHART_TYPE.COLUMN_CLUSTERED,
             "bar": XL_CHART_TYPE.BAR_CLUSTERED,
             "line": XL_CHART_TYPE.LINE_MARKERS}
    if kind not in kinds:
        raise ValueError("kind 只支持 %s" % ", ".join(kinds))

    gf = slide.shapes.add_chart(
        kinds[kind], Inches(theme.margin), Inches(BODY_TOP + 0.05),
        Inches(theme.content_width), Inches(BODY_BOTTOM - BODY_TOP - 0.45), data)
    chart_obj = gf.chart
    chart_obj.has_title = False
    chart_obj.has_legend = len(series) > 1
    if chart_obj.has_legend:
        chart_obj.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart_obj.legend.include_in_layout = False
    chart_obj.font.size = Pt(theme.size("caption"))
    chart_obj.font.name = theme.font_body
    chart_obj.font.color.rgb = theme.rgb("text")

    plot = chart_obj.plots[0]
    plot.vary_by_categories = False
    if kind != "line":
        plot.gap_width = 60
    palette = [theme.rgb(_ACCENTS[i % len(_ACCENTS)]) for i in range(len(series))]
    for i, s in enumerate(chart_obj.series):
        if kind == "line":
            s.format.line.color.rgb = palette[i]
            s.format.line.width = Pt(2.5)
        else:
            s.format.fill.solid()
            s.format.fill.fore_color.rgb = palette[i]

    if kind != "line" and len(series) == 1:
        plot.has_data_labels = True
        labels = plot.data_labels
        labels.font.size = Pt(11)
        labels.font.bold = True
        labels.font.color.rgb = theme.rgb("text")
        labels.position = XL_LABEL_POSITION.OUTSIDE_END

    value_axis = chart_obj.value_axis
    value_axis.has_major_gridlines = True
    value_axis.major_gridlines.format.line.color.rgb = theme.rgb("line")
    value_axis.format.line.fill.background()
    value_axis.tick_labels.font.size = Pt(theme.size("caption"))
    category_axis = chart_obj.category_axis
    category_axis.format.line.color.rgb = theme.rgb("line")
    category_axis.tick_labels.font.size = Pt(theme.size("caption"))

    if note:
        _text(slide, theme.margin, BODY_BOTTOM - 0.32, theme.content_width, 0.3,
              note, theme.size("caption"), theme.rgb("muted"),
              font=theme.font_body, bold=False)
    return slide


def end(prs, theme: Theme, *, title: str = "谢谢", subtitle: str = "",
        contact: str = "", page: int | None = None):
    """结尾页：与封面同一色系，收得干净。"""
    slide = add_slide(prs, theme, page)
    bg = shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H,
               fill=theme.rgb("primary_dark"))
    grad(bg, [(0.0, theme.tuple("primary_dark")), (0.62, theme.tuple("primary")),
              (1.0, theme.tuple("primary_light"))], 38)
    shape(slide, MSO_SHAPE.RECTANGLE, theme.margin, 3.04, 1.15, 0.06,
          fill=theme.rgb("accent"))
    _text(slide, theme.margin, 3.28, 8.0, 1.0, title, theme.size("hero"),
          theme.rgb("on_dark"), font=theme.font_title)
    if subtitle:
        _text(slide, theme.margin, 4.34, 7.6, 0.6, subtitle, theme.size("lead"),
              theme.rgb("on_dark"), font=theme.font_body, bold=False)
    if contact:
        _text(slide, theme.margin, SLIDE_H - 0.95, 9.0, 0.34, contact,
              theme.size("small"), theme.rgb("on_dark"), font=theme.font_body,
              bold=False)
    return slide


__all__ = ["new_deck", "add_slide", "cover", "section", "agenda", "bullets",
           "two_col", "stats", "timeline", "process", "image_grid", "quote",
           "table", "chart", "end", "TITLE_Y", "BODY_TOP", "BODY_BOTTOM"]
