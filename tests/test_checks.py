# -*- coding: utf-8 -*-
"""校验：重叠、快照 diff、贴合检查。"""
from pptx import Presentation

from pptx_native.checks import diff, flush_pairs, overlaps, snapshot
from pptx_native.geometry import abs_rect


def test_overlaps_finds_pair(deck_path):
    prs = Presentation(deck_path)
    pairs = overlaps(prs.slides[0])
    assert pairs, "红块与椭圆应该重叠"
    area, a, b = pairs[0]
    assert area > 0.5


def test_snapshot_and_diff(deck_path, tmp_path):
    prs = Presentation(deck_path)
    snap_a = snapshot(prs.slides[0])

    slide = prs.slides[0]
    for sh in slide.shapes:
        if sh.has_text_frame and sh.text_frame.text:
            sh.text_frame.paragraphs[0].runs[0].text = "改了"
            break
    snap_b = snapshot(slide)

    only_a, only_b, geom, text = diff(snap_a, snap_b)
    assert not only_a and not only_b
    assert not geom
    assert text, "文字变化应被 diff 捕获"


def test_flush_pairs(deck_path):
    prs = Presentation(deck_path)
    slide = prs.slides[1]
    grp = slide.shapes[0]
    pic = [k for k in grp.shapes if k.shape_type == 13][0]
    bar = [sh for sh in slide.shapes if sh.name.startswith("Rectangle")][0]
    res = flush_pairs([bar], [pic])
    assert len(res) == 1
    _, _, deviation = res[0]
    assert deviation < 0.01
