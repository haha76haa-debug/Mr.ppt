# -*- coding: utf-8 -*-
"""形状探针：把一页的所有形状（含 group 内层）导出成结构化数据。"""
from __future__ import annotations

from .geometry import abs_rect


def dump_slide(slide):
    rows = []

    def walk(shapes, depth=0):
        for sh in shapes:
            rec = {
                "depth": depth,
                "id": sh.shape_id,
                "name": sh.name,
                "type": str(sh.shape_type),
                "text": sh.text_frame.text if sh.has_text_frame else "",
            }
            try:
                rec["abs"] = [round(v, 4) for v in abs_rect(sh)]
                rec["raw"] = [round(v / 914400, 4) for v in
                              (sh.left, sh.top, sh.width, sh.height)]
            except TypeError:
                rec["abs"] = rec["raw"] = None
            if sh.shape_type == 13:
                rec["image"] = {"size": list(sh.image.size), "ext": sh.image.ext}
            xf = None
            try:
                xf = sh._element.spPr.find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}xfrm")
            except AttributeError:
                xf = None
            if xf is not None and xf.get("rot"):
                rec["rot"] = xf.get("rot")
            rows.append(rec)
            if sh.shape_type == 6:
                walk(sh.shapes, depth + 1)

    walk(slide.shapes)
    return rows


def format_rows(rows, text_only=False):
    lines = []
    for r in rows:
        if text_only and not r["text"].strip():
            continue
        a = r["abs"] or [0, 0, 0, 0]
        extra = ""
        if "image" in r:
            extra = " IMG %dx%d %s" % (r["image"]["size"][0], r["image"]["size"][1],
                                       r["image"]["ext"])
        if r.get("rot"):
            extra += " rot=%s" % r["rot"]
        lines.append("%s%4d %-16s %-20s x=%.2f y=%.2f w=%.2f h=%.2f%s %s"
                     % ("  " * r["depth"], r["id"], r["type"], r["name"][:20],
                        a[0], a[1], a[2], a[3], extra, repr(r["text"][:40])))
    return "\n".join(lines)
