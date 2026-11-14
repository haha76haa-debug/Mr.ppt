# -*- coding: utf-8 -*-
"""无 pytest 环境的自检脚本（CI 用 pytest，本机没 pytest 时用这个）。

    python scripts/selfcheck.py

检查项与 tests/ 一致；临时文件放在仓库内 build/_selfcheck/（受限环境里系统 temp 目录
可能不允许删除，所以不用 tempfile）。
"""
from __future__ import annotations

import sys
import uuid
import traceback
from pathlib import Path

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from pptx_native import checks, colors, geometry, groups, images, probe, styling  # noqa: E402

E = 914400
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
ROOT = Path(__file__).resolve().parents[1] / "build" / "_selfcheck"
RESULTS = []


def workdir() -> Path:
    ROOT.mkdir(parents=True, exist_ok=True)
    d = ROOT / uuid.uuid4().hex[:8]
    d.mkdir()
    return d


def check(name):
    def deco(fn):
        try:
            fn()
            RESULTS.append((name, True, ""))
        except Exception as exc:                     # noqa: BLE001
            RESULTS.append((name, False, "%s: %s" % (type(exc).__name__, exc)))
            traceback.print_exc()
        return fn
    return deco


# ------------------------------------------------------------------ fixtures
def make_photo(path):
    img = Image.new("RGB", (120, 90), (200, 210, 220))
    red = (122, 29, 24)
    for i in range(2):
        for x in range(120):
            img.putpixel((x, i), red)
            img.putpixel((x, 89 - i), red)
        for y in range(90):
            img.putpixel((i, y), red)
            img.putpixel((119 - i, y), red)
    img.save(path)


def make_deck(d: Path) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    photo = d / "photo.png"
    make_photo(photo)

    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(0.6))
    run = tb.text_frame.paragraphs[0].add_run()
    run.text = "红色残留"
    run.font.size = Pt(24)
    run.font.color.rgb = RGBColor(0xC0, 0x03, 0x03)
    box = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(2), Inches(2), Inches(1))
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xD5, 0x1A, 0x01)
    box.shadow.inherit = False
    s1.shapes.add_picture(str(photo), Inches(1), Inches(4), Inches(2), Inches(1.5))
    s1.shapes.add_shape(MSO_SHAPE.OVAL, Inches(2), Inches(2.2), Inches(2), Inches(1.5))

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    pic = s2.shapes.add_picture(str(photo), Inches(2), Inches(2), Inches(2), Inches(1.5))
    grp = s2.shapes.add_group_shape()
    pic._element.getparent().remove(pic._element)
    grp._element.append(pic._element)
    xf = pic._element.spPr.find(qn("a:xfrm"))
    xf.find(qn("a:off")).set("x", str(int(2 * E)))
    xf.find(qn("a:off")).set("y", str(int(2.135 * E)))
    xf.find(qn("a:ext")).set("cx", str(int(2 * E)))
    xf.find(qn("a:ext")).set("cy", str(int(1.5 * E)))
    gxf = grp._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
    for tag, vals, axis in (("a:off", (2, 2), "off"), ("a:chOff", (2, 2.135), "off"),
                            ("a:ext", (2, 1.5), "ext"), ("a:chExt", (2, 1.5), "ext")):
        el = gxf.find(qn(tag))
        if axis == "off":
            el.set("x", str(int(vals[0] * E)))
            el.set("y", str(int(vals[1] * E)))
        else:
            el.set("cx", str(int(vals[0] * E)))
            el.set("cy", str(int(vals[1] * E)))
    bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(2), Inches(3.15),
                              Inches(2), Inches(0.35))
    bar.name = "Rectangle 1"
    bar.text_frame.text = "注释"

    out = d / "deck.pptx"
    prs.save(str(out))
    return out


# ------------------------------------------------------------------ checks
@check("geometry.abs_rect 处理 group 偏移")
def _():
    d = workdir()
    prs = Presentation(str(make_deck(d)))
    grp = prs.slides[1].shapes[0]
    pic = [k for k in grp.shapes if k.shape_type == 13][0]
    x, y, w, h = geometry.abs_rect(pic)
    assert abs(x - 2) < 1e-3 and abs(y - 2) < 1e-3, (x, y)
    assert abs(w - 2) < 1e-3 and abs(h - 1.5) < 1e-3, (w, h)


@check("geometry.to_child_space 往返一致")
def _():
    d = workdir()
    grp = Presentation(str(make_deck(d))).slides[1].shapes[0]
    cx, cy, cw, ch = geometry.to_child_space(grp._element, 3.0, 3.0, 1.0, 0.5)
    ox, oy, sx, sy, chx, chy = geometry.transform_of(grp._element)
    assert abs((ox + (cx - chx) * sx) / E - 3.0) < 1e-3
    assert abs((oy + (cy - chy) * sy) / E - 3.0) < 1e-3
    assert abs(cw * sx / E - 1.0) < 1e-3


@check("groups.move_into_group 搬进组并保持贴合")
def _():
    d = workdir()
    prs = Presentation(str(make_deck(d)))
    slide = prs.slides[1]
    grp = slide.shapes[0]
    pic = [k for k in grp.shapes if k.shape_type == 13][0]
    bar = [sh for sh in slide.shapes if sh.name.startswith("Rectangle")][0]
    groups.move_into_group(bar, pic, flush="bottom")
    assert bar._element.getparent() is grp._element
    bx, by, bw, bh = geometry.abs_rect(bar)
    px, py, pw, ph = geometry.abs_rect(pic)
    assert abs((by + bh) - (py + ph)) < 2e-3, (by + bh, py + ph)
    assert abs(bx - px) < 2e-3


@check("groups.move_all_captions 批量搬入")
def _():
    d = workdir()
    prs = Presentation(str(make_deck(d)))
    assert groups.move_all_captions(prs.slides[1]) == 1


@check("colors 识别与清扫红色")
def _():
    d = workdir()
    el = Presentation(str(make_deck(d))).slides[0]._element
    found = {k[0] for k in colors.scan_red(el)}
    assert "C00303" in found and "D51A01" in found, found
    assert colors.sweep_red(el, "2E7BE8") >= 2
    assert not colors.scan_red(el)
    assert not colors.is_reddish("2E7BE8") and colors.is_reddish("D51A01")


@check("images 修照片自带红框")
def _():
    d = workdir()
    p = d / "p.png"
    make_photo(p)
    img = Image.open(p).convert("RGB")
    new, n = images.repaint_border_lines(img, target=(0x1B, 0x3E, 0x78))
    assert n > 0
    px = new.getpixel((0, 0))
    assert px[2] > px[0], px
    assert new.getpixel((60, 45)) == img.getpixel((60, 45))


@check("images cover / pad / encode")
def _():
    img = Image.new("RGB", (120, 90), (200, 210, 220))
    assert images.cover(img, 200, 100).size == (200, 100)
    assert images.pad(img, 200, 200).size == (200, 200)
    assert images.encode(img, "png")[:4] == b"\x89PNG"


@check("styling 渐变/半透明/叠层/竖排")
def _():
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(1))
    styling.grad(sp, [(0.0, (0, 0x30, 0xE6)), (0.5, (255, 255, 255)), (1.0, (0x01, 0x35, 0xE7))])
    assert len(sp._element.spPr.findall(".//{%s}gs" % A)) == 3
    styling.set_alpha_fill(sp, RGBColor(0x04, 0x5D, 0xE2), 0.78)
    assert sp._element.spPr.find(".//{%s}alpha" % A).get("val") == "78000"

    n0 = len(slide.shapes)
    styling.layered_text(slide, 1, 1, 4, 0.8, "标题", size=28, color=RGBColor(0x0A, 0x2C, 0x5A))
    assert len(slide.shapes) - n0 == 3

    tf = styling.vertical_label(slide, 1, 1, 0.4, 2, "班长支书",
                                size=13, color=RGBColor(255, 255, 255))
    assert [p.text for p in tf.paragraphs] == list("班长支书")


@check("checks 重叠 / 快照 diff")
def _():
    d = workdir()
    prs = Presentation(str(make_deck(d)))
    assert checks.overlaps(prs.slides[0])
    snap_a = checks.snapshot(prs.slides[0])
    for sh in prs.slides[0].shapes:
        if sh.has_text_frame and sh.text_frame.text:
            sh.text_frame.paragraphs[0].runs[0].text = "改了"
            break
    only_a, only_b, geom, text = checks.diff(snap_a, checks.snapshot(prs.slides[0]))
    assert not only_a and not only_b and not geom and text


@check("checks.audit_slide 体检越界 / 重叠 / 字号")
def _():
    from pptx_native.theme import SLIDE_W
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    for x, y, size in ((SLIDE_W - 0.4, 1.0, 8.0), (2.0, 2.0, 16.0), (2.2, 2.2, 16.0)):
        box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(2), Inches(0.5))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = "体检文字内容"
        run.font.size = Pt(size)
        run.font.name = "Microsoft YaHei"
    issues = checks.audit_slide(slide)
    kinds = {item["kind"] for item in issues}
    assert "canvas" in kinds and "overlap" in kinds and "font" in kinds, kinds


@check("probe 导出结构")
def _():
    d = workdir()
    rows = probe.dump_slide(Presentation(str(make_deck(d))).slides[1])
    assert len(rows) >= 2 and any(r.get("abs") for r in rows)


@check("theme 网格 / 字号 / 颜色角色")
def _():
    from pptx_native.theme import Theme
    t = Theme.blue()
    assert abs(t.col(0, 12)[1] - t.content_width) < 1e-6
    assert abs(t.col(0, 6)[1] + t.gutter + t.col(6, 6)[1] - t.content_width) < 1e-6
    assert t.size("h1") > t.size("body") > t.size("caption")
    assert str(t.rgb("primary")) == "0B52C4" and t.tuple("dark") == t.tuple("primary_dark")


@check("layouts 生成原生形状 / 真文字")
def _():
    from pptx_native import layouts
    from pptx_native.theme import SLIDE_W, Theme
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    t = Theme.minimal()
    layouts.cover(prs, t, title="标题", subtitle="副标题")
    layouts.agenda(prs, t, title="目录", items=["一", "二", "三", "四", "五"])
    layouts.stats(prs, t, title="数据", items=[("1", "个", "甲"), ("2", "%", "乙")])
    layouts.timeline(prs, t, title="时间", items=[("W1", "启动", "对齐"), ("W2", "上线", "灰度")])
    layouts.process(prs, t, title="流程", steps=[("一", "甲"), ("二", "乙")])
    layouts.table(prs, t, title="表", header=["A", "B"], rows=[["1", "2"]])
    layouts.chart(prs, t, title="图", categories=["Q1", "Q2"], series=[("S", [1, 2])])
    layouts.image_grid(prs, t, title="图墙", items=[(None, "占位")])
    d = workdir()
    out = d / "layouts.pptx"
    prs.save(str(out))
    again = Presentation(str(out))
    assert len(again.slides) == 8
    assert all(len(s.shapes) >= 3 for s in again.slides)
    assert SLIDE_W > 0


@check("sample 生成中性示例 deck")
def _():
    from pptx_native.sample import build_sample
    d = workdir()
    out = d / "sample.pptx"
    build_sample(str(out))
    prs = Presentation(str(out))
    assert len(prs.slides) == 12 and len(prs.slides[0].shapes) >= 6


@check("cli probe/preview/sample/themes 可运行")
def _():
    from pptx_native.cli import main
    d = workdir()
    deck = make_deck(d)
    assert main(["probe", str(deck), "--slide", "1", "--text-only"]) == 0
    prev = d / "prev.pptx"
    assert main(["preview", str(deck), str(prev), "--keep", "1"]) == 0
    assert prev.exists()
    assert main(["sample", str(d / "sample.pptx")]) == 0
    assert main(["sample", str(d / "warm.pptx"), "--theme", "warm"]) == 0
    assert main(["themes"]) == 0


def main():
    for name, passed, err in RESULTS:
        print(("PASS  " if passed else "FAIL  ") + name + ("" if passed else "   <- " + err))
    ok = sum(1 for _, p, _ in RESULTS if p)
    print("\n%d/%d checks passed" % (ok, len(RESULTS)))
    return 0 if ok == len(RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
