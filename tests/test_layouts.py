# -*- coding: utf-8 -*-
"""版式库测试：每个版式都能生成原生形状，文字可编辑，且不越界。"""
from __future__ import annotations

import io

import pytest
from PIL import Image
from pptx import Presentation
from pptx.util import Inches

from pptx_native import geometry, layouts
from pptx_native.sample import THEMES, build_sample
from pptx_native.theme import SLIDE_H, SLIDE_W, Theme


def _prs():
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    return prs


def _texts(slide):
    out = []
    stack = list(slide.shapes)
    while stack:
        sh = stack.pop()
        if sh.shape_type == 6:          # group
            stack.extend(sh.shapes)
        elif sh.has_text_frame:
            out.append(sh.text_frame.text)
    return out


def _photo_bytes():
    img = Image.new("RGB", (320, 200), (120, 150, 190))
    buf = io.BytesIO()
    img.save(buf, "PNG")
    return buf.getvalue()


def test_cover_and_end_are_dark_with_titles():
    prs = _prs()
    theme = Theme.blue()
    layouts.cover(prs, theme, title="示例标题", subtitle="副标题", meta="2026",
                  kicker="KICKER")
    layouts.end(prs, theme, title="谢谢", subtitle="再见")
    assert len(prs.slides) == 2
    assert "示例标题" in _texts(prs.slides[0])
    assert "谢谢" in _texts(prs.slides[1])


def test_each_layout_creates_native_shapes():
    prs = _prs()
    theme = Theme.dark()
    builders = [
        lambda: layouts.agenda(prs, theme, title="目录", items=["一", "二", "三"]),
        lambda: layouts.section(prs, theme, number=2, title="章节"),
        lambda: layouts.bullets(prs, theme, title="要点",
                                items=[("小标题", "一句说明"), "纯文字条目"]),
        lambda: layouts.two_col(prs, theme, title="对比",
                                left=("左", ["甲", "乙"]), right=("右", "丙")),
        lambda: layouts.stats(prs, theme, title="数据",
                              items=[("128", "个", "需求"), ("96", "%", "按时")]),
        lambda: layouts.timeline(prs, theme, title="时间线",
                                 items=[("第 1 周", "启动", "对齐目标"),
                                        ("第 4 周", "上线", "灰度")]),
        lambda: layouts.process(prs, theme, title="流程",
                                steps=[("一", "说明一"), ("二", "说明二")]),
        lambda: layouts.image_grid(prs, theme, title="图片",
                                   items=[(None, "占位"), (_photo_bytes(), "图")]),
        lambda: layouts.quote(prs, theme, text="一句引用", author="某人"),
        lambda: layouts.table(prs, theme, title="表",
                              header=["A", "B"], rows=[["1", "2"]]),
        lambda: layouts.chart(prs, theme, title="图",
                              categories=["Q1", "Q2"], series=[("系列", [1, 2])]),
    ]
    for build in builders:
        before = len(prs.slides)
        slide = build()
        assert slide is not None
        assert len(prs.slides) == before + 1
        assert len(slide.shapes) >= 3, slide.shapes


def test_agenda_numbers_five_items_in_two_columns():
    prs = _prs()
    slide = layouts.agenda(prs, Theme.blue(), title="目录",
                           items=["甲", "乙", "丙", "丁", "戊"])
    texts = _texts(slide)
    for n in ("01", "05"):
        assert n in texts
    # 分两列：左列 3 个，右列 2 个
    nums = [sh for sh in slide.shapes if sh.has_text_frame and
            sh.text_frame.text in ("01", "02", "03", "04", "05")]
    xs = sorted({round(sh.left / 914400, 1) for sh in nums})
    assert len(xs) == 2


def test_stats_renders_value_unit_and_label():
    prs = _prs()
    slide = layouts.stats(prs, Theme.blue(), title="数据",
                          items=[("128", "个", "需求"), ("96", "%", "按时")],
                          note="示例数据")
    texts = _texts(slide)
    assert any(t.startswith("128") and "个" in t for t in texts)
    assert any(t.startswith("96") and "%" in t for t in texts)
    assert "需求" in texts and "按时" in texts
    assert "示例数据" in texts


def test_image_grid_placeholder_and_picture():
    prs = _prs()
    slide = layouts.image_grid(prs, Theme.blue(), title="图片",
                               items=[(None, "图片 01"), (_photo_bytes(), "图片 02")])
    kinds = [sh.shape_type for sh in slide.shapes]
    assert 13 in kinds            # 图片
    texts = _texts(slide)
    assert "图片占位" in texts and "图片 01" in texts and "图片 02" in texts


def test_table_has_expected_cells():
    prs = _prs()
    slide = layouts.table(prs, Theme.blue(), title="表",
                          header=["指标", "目标"], rows=[["交付", "120"]])
    table = [sh for sh in slide.shapes if sh.has_table][0].table
    assert len(table.rows) == 2 and len(table.columns) == 2
    assert table.cell(1, 0).text_frame.text == "交付"


def test_chart_is_native_and_editable():
    prs = _prs()
    slide = layouts.chart(prs, Theme.blue(), title="图",
                          categories=["Q1", "Q2"],
                          series=[("计划", [1, 2]), ("实际", [1, 3])])
    frame = [sh for sh in slide.shapes if sh.has_chart][0]
    assert list(frame.chart.plots[0].categories) == ["Q1", "Q2"]
    assert len(frame.chart.series) == 2


def test_all_text_stays_inside_canvas(tmp_path):
    out = tmp_path / "sample.pptx"
    build_sample(str(out))
    prs = Presentation(str(out))
    for slide in prs.slides:
        for sh in slide.shapes:
            if not sh.has_text_frame or not sh.text_frame.text:
                continue
            x, y, w, h = geometry.abs_rect(sh)
            assert -0.02 <= x and x + w <= SLIDE_W + 0.02, (sh.text_frame.text, x, w)
            assert -0.02 <= y and y + h <= SLIDE_H + 0.02, (sh.text_frame.text, y, h)


def test_sample_deck_and_all_themes(tmp_path):
    for name in THEMES:
        out = tmp_path / ("sample-%s.pptx" % name)
        build_sample(str(out), THEMES[name]())
        prs = Presentation(str(out))
        assert len(prs.slides) == 12
        for slide in prs.slides:
            assert len(slide.shapes) >= 3
        # 每页至少有一段真文字（图片墙只有“图片占位”也算）
        assert all(_texts(s) for s in prs.slides)
