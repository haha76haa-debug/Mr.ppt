# -*- coding: utf-8 -*-
"""设计系统测试：颜色角色、字号阶梯、12 栏网格、主题预设。"""
from __future__ import annotations

import pytest
from pptx.dml.color import RGBColor

from pptx_native.theme import SLIDE_H, SLIDE_W, TYPE_SCALE, Theme


def test_presets_are_distinct_and_complete():
    names = set()
    for factory in (Theme.blue, Theme.dark, Theme.minimal, Theme.warm):
        theme = factory()
        assert theme.name and theme.primary and theme.bg
        names.add(theme.primary)
    assert len(names) == 4


def test_rgb_accepts_role_alias_and_hex():
    theme = Theme.blue()
    assert theme.rgb("primary") == RGBColor.from_string("0B52C4")
    assert theme.rgb("dark") == theme.rgb("primary_dark")
    assert theme.rgb("white") == theme.rgb("on_dark")
    assert theme.rgb("2E7BE8") == RGBColor.from_string("2E7BE8")
    assert theme.tuple("primary") == (0x0B, 0x52, 0xC4)


def test_type_scale_is_descending():
    order = ["display", "hero", "h1", "h2", "lead", "h3", "body", "small", "caption"]
    sizes = [TYPE_SCALE[k] for k in order]
    assert sizes == sorted(sizes, reverse=True)
    assert TYPE_SCALE["stat"] > TYPE_SCALE["body"]


def test_grid_columns_fill_content_width():
    theme = Theme.blue()
    x0, w0 = theme.col(0, 12)
    assert x0 == pytest.approx(theme.margin)
    assert w0 == pytest.approx(theme.content_width)
    # 两栏之间留一个 gutter
    xa, wa = theme.col(0, 6)
    xb, wb = theme.col(6, 6)
    assert xb - (xa + wa) == pytest.approx(theme.gutter)
    assert xb + wb == pytest.approx(x0 + w0)


def test_grid_rejects_out_of_range():
    theme = Theme.blue()
    with pytest.raises(ValueError):
        theme.col(0, 13)
    with pytest.raises(ValueError):
        theme.col(-1, 1)


def test_font_override_and_roundtrip():
    theme = Theme.blue().with_fonts("Source Han Sans SC", "Songti SC")
    assert theme.font("title") == "Source Han Sans SC"
    assert theme.font() == "Songti SC"
    assert Theme.from_dict(theme.as_dict()) == theme


def test_slide_size_is_16_9():
    assert SLIDE_W == pytest.approx(13.333, abs=1e-3)
    assert SLIDE_H == pytest.approx(7.5, abs=1e-3)
