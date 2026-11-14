# -*- coding: utf-8 -*-
"""绝对几何：把 group 的 off/chOff 换算成本模板里所有位置计算的唯一入口。

python-pptx 的 ``shape.left/top`` 是 group 子坐标系里的原始值，模板里常见
``chOff != off``（差 0.13in 量级），直接拿来对齐会整体偏移。
"""
from __future__ import annotations

from pptx.oxml.ns import qn

EMU = 914400


def group_xfrm(group_el):
    """返回 group 元素里的 ``a:xfrm``。"""
    return group_el.find(qn("p:grpSpPr")).find(qn("a:xfrm"))


def transform_of(group_el):
    """(off_x, off_y, sx, sy, chOff_x, chOff_y)，单位 EMU / 倍率。"""
    xf = group_xfrm(group_el)
    off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
    chOff, chExt = xf.find(qn("a:chOff")), xf.find(qn("a:chExt"))
    sx = int(ext.get("cx")) / int(chExt.get("cx"))
    sy = int(ext.get("cy")) / int(chExt.get("cy"))
    return (int(off.get("x")), int(off.get("y")), sx, sy,
            int(chOff.get("x")), int(chOff.get("y")))


def abs_rect(shape):
    """形状在幻灯片坐标系里的真实位置，返回 (x, y, w, h)，单位英寸。"""
    x, y = shape.left, shape.top
    w, h = shape.width, shape.height
    parent = shape._element.getparent()
    while parent is not None and parent.tag == qn("p:grpSp"):
        ox, oy, sx, sy, cx, cy = transform_of(parent)
        x = ox + (x - cx) * sx
        y = oy + (y - cy) * sy
        w *= sx
        h *= sy
        parent = parent.getparent()
    return x / EMU, y / EMU, w / EMU, h / EMU


def to_child_space(group_el, x, y, w, h):
    """把绝对英寸坐标换算成 group 的 child 坐标（EMU），用于把形状搬进组。"""
    ox, oy, sx, sy, cx, cy = transform_of(group_el)
    return (
        cx + (x * EMU - ox) / sx,
        cy + (y * EMU - oy) / sy,
        w * EMU / sx,
        h * EMU / sy,
    )


def iter_shapes(shapes):
    """深度遍历（含 group 内层）。"""
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:          # GROUP
            yield from iter_shapes(sh.shapes)


def find_shape(slide, shape_id):
    for sh in iter_shapes(slide.shapes):
        if sh.shape_id == shape_id:
            return sh
    return None


def parent_group(shape):
    el = shape._element.getparent()
    return el if el is not None and el.tag == qn("p:grpSp") else None


def bbox_of(shape):
    return abs_rect(shape)


def is_flush(a, b, tol=0.004):
    """两个矩形是否「贴边」（用于检查注释条是否贴合照片底边）。"""
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    return (abs((ay + ah) - (by + bh)) < tol
            and abs(ax - bx) < tol
            and abs(aw - bw) < tol)
