# -*- coding: utf-8 -*-
"""命令行入口：``pptx-native <command>``。

    probe     导出某页全部形状（含 group 内层）
    preview   抽单页成小 deck（渲染 / 让 PowerPoint 打得开）
    export    用 PowerPoint 真机导出某页 PNG
    colors    颜色普查（找出杂色/红色）
    sweep     把红色系统一替换成蓝色系
    borders   修照片文件自带的细边框
    diff      两版单页结构 diff
    overlaps  形状重叠自检
    sample    生成示例 deck（不依赖外部素材）
    themes    列出可用的主题预设
"""
from __future__ import annotations

import argparse
import json
import sys

from pptx import Presentation

from . import checks, colors, groups, images, probe, render
from .geometry import abs_rect


def _slide(prs, index):
    if not 1 <= index <= len(prs.slides):
        raise SystemExit("页码 %d 超出范围（共 %d 页）" % (index, len(prs.slides)))
    return prs.slides[index - 1]


# ------------------------------------------------------------------ commands
def cmd_probe(args):
    prs = Presentation(args.deck)
    rows = probe.dump_slide(_slide(prs, args.slide))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(rows, fh, ensure_ascii=False, indent=1)
        print("wrote %s (%d shapes)" % (args.json, len(rows)))
    else:
        print(probe.format_rows(rows, text_only=args.text_only))


def cmd_preview(args):
    keep = [int(x) for x in args.keep.split(",")]
    render.preview_deck(args.deck, args.out, keep)
    import os
    print("saved %s (%.1f MB)" % (args.out, os.path.getsize(args.out) / 1e6))


def cmd_export(args):
    out = render.export_slide_via_powerpoint(args.deck, args.out, args.slide,
                                             args.width, args.height, script=args.script or
                                             render.PS1)
    print("exported", out)


def cmd_colors(args):
    prs = Presentation(args.deck)
    census = colors.colour_census(_slide(prs, args.slide)._element)
    print("srgb  :", dict(census["srgb"].most_common(25)))
    print("scheme:", dict(census["scheme"]))
    if census["prst"]:
        print("preset:", dict(census["prst"]))
    red = colors.scan_red(_slide(prs, args.slide)._element)
    print("red-ish:", {"%s@%s" % k: v for k, v in red.items()} or "none")


def cmd_sweep(args):
    prs = Presentation(args.deck)
    slide = _slide(prs, args.slide)
    found = colors.scan_red(slide._element)
    print("red-ish found:", {"%s@%s" % k: v for k, v in found.items()} or "none")
    if args.list:
        return
    n = colors.sweep_red(slide._element, args.to)
    prs.save(args.out)
    print("replaced %d colour(s) -> #%s, saved %s" % (n, args.to.upper(), args.out))


def cmd_borders(args):
    prs = Presentation(args.deck)
    slide = _slide(prs, args.slide)
    target = colors.hex_rgb(args.to)
    photos = []
    for sh in slide.shapes:
        if sh.shape_type == 6:
            photos += [k for k in sh.shapes if k.shape_type == 13]
    seen, total = set(), 0
    for pic in photos:
        part = pic.part.related_part(pic._element.blipFill.blip.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"))
        key = str(part.partname)
        if key in seen:
            continue
        seen.add(key)
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(part.blob)).convert("RGB")
        new_img, n = images.repaint_border_lines(img, target)
        total += n
        print("  %-34s %-11s border lines: %d" % (key, str(img.size), n))
        if args.report_only:
            continue
        part._blob = images.encode(new_img, part.ext)
    print("photos: %d, lines repainted: %d" % (len(seen), total))
    if not args.report_only:
        prs.save(args.out)
        print("saved", args.out)


def cmd_diff(args):
    pa, pb = Presentation(args.a), Presentation(args.b)
    a, b = checks.snapshot(_slide(pa, args.slide)), checks.snapshot(_slide(pb, args.slide))
    only_a, only_b, geom, text = checks.diff(a, b)
    print("shapes: %d -> %d" % (len(a), len(b)))
    if only_a:
        print("only in A:", only_a)
    if only_b:
        print("only in B:", only_b)
    for k, ga, gb in geom:
        print("GEOM id=%s\n   A %s\n   B %s" % (k, ga, gb))
    for k, ta, tb in text:
        print("TEXT id=%s  %r -> %r" % (k, ta[:40], tb[:40]))
    print("geometry changes: %d, text changes: %d" % (len(geom), len(text)))
    return 0 if not geom else 1


def cmd_overlaps(args):
    prs = Presentation(args.deck)
    pairs = checks.overlaps(_slide(prs, args.slide), args.min_area)
    print("overlapping pairs:", len(pairs))
    for area, a, b in pairs:
        print("  %6.3f in²  #%-4s %-16s %-14r <-> #%-4s %-16s %-14r"
              % (area, a[0], a[1], a[3], b[0], b[1], b[3]))


def cmd_sample(args):
    from .sample import THEMES, build_sample
    if args.theme not in THEMES:
        raise SystemExit("未知主题 %r，可选：%s" % (args.theme, ", ".join(THEMES)))
    out = build_sample(args.out, THEMES[args.theme]())
    print("sample deck:", out)


def cmd_themes(args):
    from .sample import THEMES
    for name, factory in THEMES.items():
        t = factory()
        print("%-8s %-6s primary=#%s  surface=#%s" % (name, t.name, t.primary, t.surface))


# ------------------------------------------------------------------ parser
def build_parser():
    p = argparse.ArgumentParser(prog="pptx-native",
                                description="PPT 原生构建与校验工具集")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("probe", help="导出某页全部形状")
    s.add_argument("deck")
    s.add_argument("--slide", type=int, required=True)
    s.add_argument("--json")
    s.add_argument("--text-only", action="store_true")
    s.set_defaults(func=cmd_probe)

    s = sub.add_parser("preview", help="抽单页成小 deck")
    s.add_argument("deck")
    s.add_argument("out")
    s.add_argument("--keep", default="1")
    s.set_defaults(func=cmd_preview)

    s = sub.add_parser("export", help="PowerPoint 真机导出 PNG（Windows+Office）")
    s.add_argument("deck")
    s.add_argument("out")
    s.add_argument("--slide", type=int, default=1)
    s.add_argument("--width", type=int, default=2001)
    s.add_argument("--height", type=int, default=1126)
    s.add_argument("--script", help="export_slide.ps1 路径")
    s.set_defaults(func=cmd_export)

    s = sub.add_parser("colors", help="颜色普查")
    s.add_argument("deck")
    s.add_argument("--slide", type=int, required=True)
    s.set_defaults(func=cmd_colors)

    s = sub.add_parser("sweep", help="红色系统一替换")
    s.add_argument("deck")
    s.add_argument("out")
    s.add_argument("--slide", type=int, required=True)
    s.add_argument("--to", default="2E7BE8")
    s.add_argument("--list", action="store_true")
    s.set_defaults(func=cmd_sweep)

    s = sub.add_parser("borders", help="修照片自带细边框")
    s.add_argument("deck")
    s.add_argument("out")
    s.add_argument("--slide", type=int, required=True)
    s.add_argument("--to", default="1B3E78")
    s.add_argument("--report-only", action="store_true")
    s.set_defaults(func=cmd_borders)

    s = sub.add_parser("diff", help="两版单页结构 diff")
    s.add_argument("a")
    s.add_argument("b")
    s.add_argument("--slide", type=int, required=True)
    s.set_defaults(func=cmd_diff)

    s = sub.add_parser("overlaps", help="形状重叠自检")
    s.add_argument("deck")
    s.add_argument("--slide", type=int, required=True)
    s.add_argument("--min-area", type=float, default=0.02)
    s.set_defaults(func=cmd_overlaps)

    s = sub.add_parser("sample", help="生成示例 deck")
    s.add_argument("out", nargs="?", default="sample.pptx")
    s.add_argument("--theme", default="blue", help="blue / dark / minimal / warm")
    s.set_defaults(func=cmd_sample)

    s = sub.add_parser("themes", help="列出主题预设")
    s.set_defaults(func=cmd_themes)

    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    return args.func(args) or 0


if __name__ == "__main__":
    sys.exit(main())
