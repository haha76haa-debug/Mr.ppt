# -*- coding: utf-8 -*-
"""颜色：红色识别、扫描、清扫。"""
from pptx import Presentation

from pptx_native.colors import colour_census, is_reddish, scan_red, sweep_red


def test_is_reddish():
    assert is_reddish("C00303")
    assert is_reddish("D51A01")
    assert not is_reddish("2E7BE8")
    assert not is_reddish("0A2C5A")
    assert not is_reddish("FFFFFF")
    assert not is_reddish("123456")


def test_scan_and_sweep(deck_path):
    prs = Presentation(deck_path)
    el = prs.slides[0]._element
    hits = scan_red(el)
    found = {k[0] for k in hits}
    assert "C00303" in found          # 红字
    assert "D51A01" in found          # 红块

    n = sweep_red(el, "2E7BE8")
    assert n >= 2
    assert not scan_red(el), "清扫后不应再有红色矢量颜色"


def test_colour_census(deck_path):
    prs = Presentation(deck_path)
    census = colour_census(prs.slides[0]._element)
    assert census["srgb"]["2E7BE8"] >= 1 or census["srgb"]
