# -*- coding: utf-8 -*-
"""把注释条搬进照片组：搬完必须仍然贴合照片底边，且继承组的变换。"""
import pytest
from pptx import Presentation

from pptx_native.geometry import abs_rect
from pptx_native.groups import move_all_captions, move_into_group


def _slide2(deck_path):
    prs = Presentation(deck_path)
    return prs, prs.slides[1]


def test_move_into_group_keeps_flush(deck_path):
    prs, slide = _slide2(deck_path)
    grp = slide.shapes[0]
    pic = [k for k in grp.shapes if k.shape_type == 13][0]
    bar = [sh for sh in slide.shapes if sh.name.startswith("Rectangle")][0]

    before = abs_rect(bar)
    assert before[1] + before[3] == pytest.approx(3.5, abs=1e-3)   # 贴着照片底边

    move_into_group(bar, pic, flush="bottom")

    # 已进入组
    assert bar._element.getparent() is grp._element
    after = abs_rect(bar)
    assert after[1] + after[3] == pytest.approx(abs_rect(pic)[1] + abs_rect(pic)[3], abs=2e-3)
    assert after[0] == pytest.approx(abs_rect(pic)[0], abs=2e-3)


def test_move_all_captions(deck_path):
    prs, slide = _slide2(deck_path)
    n = move_all_captions(slide)
    assert n == 1
    prs.save(deck_path)
    prs2 = Presentation(deck_path)
    grp2 = prs2.slides[1].shapes[0]
    assert grp2.shape_type == 6
    names = [sh.name for sh in grp2.shapes]
    assert any(nm.startswith("Rectangle") for nm in names)
