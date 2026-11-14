# -*- coding: utf-8 -*-
"""组合（group）操作：把形状搬进组、复制组并重编号。

为什么要把装饰搬进组：模板照片组常带 ``a:scene3d`` 透视相机，PowerPoint 会把组内照片
画成梯形/斜的，而组外平放的注释条不会跟着变形 → 看着「歪」「不贴合」。
搬进同一组后，它继承同一套 2D/3D 变换，任何渲染器下都贴合。
"""
from __future__ import annotations

import copy

from pptx.oxml.ns import qn

from .geometry import abs_rect, parent_group, to_child_space


def max_shape_id(slide):
    return max(int(e.get("id")) for e in slide._element.iter(qn("p:cNvPr")))


def renumber(element, start):
    """给 element 内全部 p:cNvPr 重新编号（同一页 id 必须唯一）。"""
    n = start
    for c in element.iter(qn("p:cNvPr")):
        n += 1
        c.set("id", str(n))
    return n


def clone_shape(shape, slide, start_id):
    """复制形状（含 group 子树）并追加到页面末尾，返回 (element, new_max_id)。"""
    el = copy.deepcopy(shape._element)
    n = renumber(el, start_id)
    slide.shapes._spTree.append(el)
    return el, n


def set_child_bounds(element, x, y, w, h):
    """设置元素自身 xfrm（child 坐标系，EMU）。"""
    xf = element.spPr.find(qn("a:xfrm"))
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    off.set("x", str(int(x)))
    off.set("y", str(int(y)))
    ext.set("cx", str(int(w)))
    ext.set("cy", str(int(h)))


def move_into_group(bar, photo, *, flush="bottom"):
    """把 bar 搬进 photo 所在的组，并在 child 空间里让 bar 贴住 photo 的底边。

    ``flush``：bottom（底边对齐）/ top / left / right。
    返回目标组的元素。
    """
    grp = photo._element.getparent()
    if grp is None or grp.tag != qn("p:grpSp"):
        raise ValueError("photo 不在一个 group 里，无法继承组合变换")

    bx, by, bw, bh = abs_rect(bar)
    px, py, pw, ph = abs_rect(photo)
    if flush == "bottom":
        bx, by, bw = px, py + ph - bh, pw
    elif flush == "top":
        bx, by, bw = px, py, pw
    elif flush == "left":
        bx, by, bh = px, py, ph
    elif flush == "right":
        bx, by, bh = px + pw - bw, py, ph

    cx, cy, cw, ch = to_child_space(grp, bx, by, bw, bh)

    el = bar._element
    el.getparent().remove(el)
    grp.append(el)                     # 放在组末尾 → 盖在照片上方
    set_child_bounds(el, cx, cy, cw, ch)
    return grp


def move_all_captions(slide, prefix="Rectangle", tolerance=0.15):
    """把页面上名字以 prefix 开头的文字条，按几何关系搬进对应的照片组。

    匹配规则：条的水平中心落在照片内，且条的底边与照片底边相差 < tolerance。
    返回成功搬入的条数。
    """
    photos = []
    for sh in slide.shapes:
        if sh.shape_type == 6:
            for k in sh.shapes:
                if k.shape_type == 13:              # PICTURE
                    photos.append(k)

    moved = 0
    for bar in list(slide.shapes):
        if not (bar.name.startswith(prefix) and bar.has_text_frame
                and bar.text_frame.text.strip()):
            continue
        bx, by, bw, bh = abs_rect(bar)
        cx = bx + bw / 2
        best = None
        for photo in photos:
            px, py, pw, ph = abs_rect(photo)
            if px - 0.01 <= cx <= px + pw + 0.01 and abs((py + ph) - (by + bh)) < tolerance:
                d = abs((py + ph) - (by + bh)) + abs(px - bx)
                if best is None or d < best[0]:
                    best = (d, photo)
        if best is None:
            continue
        move_into_group(bar, best[1], flush="bottom")
        moved += 1
    return moved


def group_of(shape):
    return parent_group(shape)
