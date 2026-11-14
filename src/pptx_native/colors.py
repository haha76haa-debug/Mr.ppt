# -*- coding: utf-8 -*-
"""颜色审计与清扫。

模板里的红色常常不是「一个红形状」，而是藏在：
  * 文字 run 的 solidFill / gradFill（叠层文字的一层）
  * 文字描边 a:ln（→ 文字周围一圈粉色光晕）
  * 阴影 a:outerShdw（→ 暖色光晕）
  * 组合里小图标的阴影
所以清扫必须遍历整页 XML，而不是只看形状 fill。
"""
from __future__ import annotations

from collections import Counter

from pptx.oxml.ns import qn

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
RED_PRESETS = {"red", "darkred"}


def hex_rgb(value):
    return tuple(int(value[i:i + 2], 16) for i in (0, 2, 4))


def is_reddish(value, gap=35, min_r=60, max_gb=170):
    """判断 RRGGBB 是否明显偏红（不是照片内容，只看矢量颜色）。"""
    if not value or len(value) != 6:
        return False
    try:
        r, g, b = hex_rgb(value)
    except ValueError:
        return False
    return r > min_r and r >= g + gap and r >= b + gap and g < max_gb and b < max_gb


def scan_red(slide_element):
    """返回 {(颜色, 上下文): 次数}，上下文 ∈ text/shadow/line/fill/grad。"""
    out = Counter()
    for el in slide_element.iter():
        tag = el.tag.split("}")[-1]
        if tag not in ("srgbClr", "prstClr"):
            continue
        val = el.get("val") or ""
        hit = (is_reddish(val.upper()) if tag == "srgbClr"
               else val.lower() in RED_PRESETS)
        if not hit:
            continue
        chain = []
        p = el.getparent()
        while p is not None and len(chain) < 8:
            chain.append(p.tag.split("}")[-1])
            p = p.getparent()
        if "outerShdw" in chain or "reflection" in chain:
            ctx = "shadow"
        elif "rPr" in chain:
            ctx = "text"
        elif "ln" in chain:
            ctx = "line"
        elif "gradFill" in chain:
            ctx = "grad"
        else:
            ctx = "fill"
        out[(val.upper(), ctx)] += 1
    return out


def sweep_red(slide_element, to_hex="2E7BE8"):
    """把所有红色系矢量颜色替换成目标色，返回替换次数。"""
    n = 0
    for el in slide_element.iter():
        tag = el.tag.split("}")[-1]
        if tag == "srgbClr":
            v = (el.get("val") or "").upper()
            if is_reddish(v):
                el.set("val", to_hex.upper())
                n += 1
        elif tag == "prstClr" and (el.get("val") or "").lower() in RED_PRESETS:
            el.set("val", to_hex.upper())
            n += 1
    return n


def theme_colors(prs):
    """读取母版主题的 12 个配色（排查「主题色其实是红色」的情况）。"""
    try:
        theme_part = prs.slide_masters[0].part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme")
    except Exception:
        return {}
    from lxml import etree
    root = etree.fromstring(theme_part.blob)
    scheme = root.find(".//{%s}clrScheme" % A)
    if scheme is None:
        return {}
    out = {}
    for child in scheme:
        c = child[0] if len(child) else None
        if c is None:
            continue
        out[child.tag.split("}")[-1]] = c.get("val") or c.get("lastClr")
    return out


def colour_census(slide_element):
    """统计整页用到的 srgbClr / schemeClr / prstClr 次数（排查杂色用）。"""
    srgb, scheme, prst = Counter(), Counter(), Counter()
    for el in slide_element.iter():
        tag = el.tag.split("}")[-1]
        if tag == "srgbClr":
            srgb[(el.get("val") or "").upper()] += 1
        elif tag == "schemeClr":
            scheme[el.get("val")] += 1
        elif tag == "prstClr":
            prst[el.get("val")] += 1
    return {"srgb": srgb, "scheme": scheme, "prst": prst}
