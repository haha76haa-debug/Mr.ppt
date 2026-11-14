# -*- coding: utf-8 -*-
"""几何：group 的 off/chOff 换算。"""
import pytest
from pptx import Presentation

from pptx_native.geometry import abs_rect, is_flush, to_child_space, transform_of

E = 914400


def test_abs_rect_handles_group_offset(deck_path):
    """group 的 chOff.y = off.y + 0.135in 时，照片渲染位置仍是 (2,2,2,1.5)。"""
    prs = Presentation(deck_path)
    slide = prs.slides[1]
    grp = slide.shapes[0]
    assert grp.shape_type == 6
    pic = [k for k in grp.shapes if k.shape_type == 13][0]
    assert abs_rect(pic) == pytest.approx((2.0, 2.0, 2.0, 1.5), abs=1e-3)


def test_to_child_space_roundtrip(deck_path):
    prs = Presentation(deck_path)
    slide = prs.slides[1]
    grp = slide.shapes[0]
    cx, cy, cw, ch = to_child_space(grp._element, 3.0, 3.0, 1.0, 0.5)
    ox, oy, sx, sy, chx, chy = transform_of(grp._element)
    # 反算回绝对坐标应当一致
    assert (ox + (cx - chx) * sx) / E == pytest.approx(3.0, abs=1e-3)
    assert (oy + (cy - chy) * sy) / E == pytest.approx(3.0, abs=1e-3)
    assert cw * sx / E == pytest.approx(1.0, abs=1e-3)


def test_is_flush():
    assert is_flush((2, 3.15, 2, 0.35), (2, 2.0, 2, 1.5))
    assert not is_flush((2, 3.20, 2, 0.35), (2, 2.0, 2, 1.5))
