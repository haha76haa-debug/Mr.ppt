# -*- coding: utf-8 -*-
"""校验：结构快照 / diff / 重叠自检 / 贴合断言。

交付前建议顺序：
    python -m pptx_native.cli diff --a before.pptx --b after.pptx --slide 8
    python -m pptx_native.cli overlaps deck.pptx --slide 8
"""
from __future__ import annotations

from pptx.oxml.ns import qn

from .geometry import abs_rect


def snapshot(slide):
    """把整页形状拍成可比对的 dict：id -> (类型, 几何, 文字, 图片尺寸)。"""
    out = {}

    def walk(shapes):
        for sh in shapes:
            xf = None
            try:
                xf = sh._element.spPr.find(qn("a:xfrm"))
            except AttributeError:
                xf = None
            if xf is None and sh.shape_type == 6:
                xf = sh._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
            geo = None
            if xf is not None:
                off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
                geo = (int(off.get("x")), int(off.get("y")),
                       int(ext.get("cx")), int(ext.get("cy")), xf.get("rot"))
            img = list(sh.image.size) if sh.shape_type == 13 else None
            out[sh.shape_id] = (str(sh.shape_type), geo,
                                sh.text_frame.text if sh.has_text_frame else "", img)
            if sh.shape_type == 6:
                walk(sh.shapes)

    walk(slide.shapes)
    return out


def diff(a, b):
    """比较两份 snapshot，返回 (新增, 缺失, 几何变化, 文字变化)。"""
    only_a = sorted(set(a) - set(b))
    only_b = sorted(set(b) - set(a))
    geom, text = [], []
    for k in sorted(set(a) & set(b)):
        if a[k][:2] != b[k][:2]:
            geom.append((k, a[k][1], b[k][1]))
        elif a[k][2] != b[k][2]:
            text.append((k, a[k][2], b[k][2]))
    return only_a, only_b, geom, text


def _area(rect_a, rect_b):
    ax, ay, aw, ah = rect_a
    bx, by, bw, bh = rect_b
    x = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    y = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return x * y


def overlaps(slide, min_area=0.02):
    """列出互相重叠的形状对，按重叠面积降序。"""
    items = []
    for sh in slide.shapes:
        try:
            rect = abs_rect(sh)
        except TypeError:
            continue
        if rect[2] > 0 and rect[3] > 0:
            items.append((sh.shape_id, sh.name, rect,
                          sh.text_frame.text if sh.has_text_frame else ""))
    pairs = []
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            area = _area(items[i][2], items[j][2])
            if area >= min_area:
                pairs.append((area, items[i], items[j]))
    pairs.sort(reverse=True, key=lambda t: t[0])
    return pairs


def flush_pairs(bars, photos, tol=0.004):
    """检查注释条是否贴合照片底边，返回 [(bar, photo, 偏差)]。"""
    result = []
    for bar in bars:
        bx, by, bw, bh = abs_rect(bar)
        best = None
        for photo in photos:
            px, py, pw, ph = abs_rect(photo)
            d = (abs((py + ph) - (by + bh)) + abs(px - bx) + abs(pw - bw))
            if best is None or d < best[0]:
                best = (d, photo)
        if best:
            result.append((bar, best[1], best[0]))
    return result
