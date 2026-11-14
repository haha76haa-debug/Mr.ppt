# -*- coding: utf-8 -*-
"""页面体检（checks.audit_slide）测试：越界 / 字号 / 重叠 / 溢出。"""
from __future__ import annotations

from pptx import Presentation
from pptx.util import Inches, Pt

from pptx_native import checks, layouts
from pptx_native.sample import build_sample
from pptx_native.theme import SLIDE_H, SLIDE_W, Theme


def _prs():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def _slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _text(slide, x, y, w, h, text, size=16.0):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(size)
    return box


def _kinds(issues, level=None):
    return {item["kind"] for item in issues
            if level is None or item["level"] == level}


def test_clean_layout_has_no_errors():
    prs = _prs()
    theme = Theme.blue()
    layouts.bullets(prs, theme, title="要点",
                    items=[("标题一", "一句说明"), ("标题二", "一句说明")])
    issues = checks.audit_slide(prs.slides[0])
    assert not any(item["level"] == "error" for item in issues), issues


def test_off_canvas_text_is_error():
    prs = _prs()
    _text(_slide(prs), SLIDE_W - 0.5, 1.0, 2.0, 0.5, "跑到画布外面的文字")
    issues = checks.audit_slide(prs.slides[0])
    assert "canvas" in _kinds(issues, "error"), issues


def test_small_font_is_warning():
    prs = _prs()
    _text(_slide(prs), 1.0, 1.0, 4.0, 0.4, "小字提醒", size=7.0)
    issues = checks.audit_slide(prs.slides[0])
    assert "font" in _kinds(issues), issues


def test_overlapping_text_boxes_are_reported():
    prs = _prs()
    slide = _slide(prs)
    _text(slide, 1.0, 1.0, 4.0, 1.0, "第一段文字内容")
    _text(slide, 1.2, 1.2, 4.0, 1.0, "第二段文字内容")
    issues = checks.audit_slide(slide)
    assert "overlap" in _kinds(issues), issues


def test_long_text_in_small_box_is_overflow_risk():
    prs = _prs()
    slide = _slide(prs)
    _text(slide, 1.0, 1.0, 1.2, 0.3, "这是一段很长的中文说明文字" * 4, size=18.0)
    issues = checks.audit_slide(slide)
    assert "overflow" in _kinds(issues), issues


def test_format_issues_one_line_per_issue():
    issues = [
        {"level": "error", "kind": "canvas", "shape_id": 2,
         "name": "TextBox 1", "message": "越界"},
        {"level": "warn", "kind": "font", "shape_id": 3,
         "name": "TextBox 2", "message": "字太小"},
    ]
    lines = checks.format_issues(issues)
    assert len(lines) == 2
    assert lines[0].startswith("[error]") and lines[1].startswith("[warn]")


def test_audit_survives_grouped_shapes(deck_path):
    prs = Presentation(deck_path)
    issues = checks.audit_slide(prs.slides[1])
    assert isinstance(issues, list)


def test_sample_deck_has_no_errors(tmp_path):
    out = tmp_path / "sample.pptx"
    build_sample(str(out))
    prs = Presentation(str(out))
    for slide in prs.slides:
        issues = checks.audit_slide(slide)
        errors = [item for item in issues if item["level"] == "error"]
        assert not errors, errors
