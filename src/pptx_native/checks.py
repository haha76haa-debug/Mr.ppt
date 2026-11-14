# -*- coding: utf-8 -*-
"""校验：结构快照 / diff / 重叠自检 / 贴合断言 / 页面体检。

交付前建议顺序：
    python -m pptx_native lint deck.pptx --slide 3
    python -m pptx_native overlaps deck.pptx --slide 3
    python -m pptx_native diff before.pptx after.pptx --slide 3
"""
from __future__ import annotations

import math

from pptx.oxml.ns import qn

from .geometry import abs_rect, iter_shapes
from .theme import SLIDE_H, SLIDE_W

#: 交付体检的默认阈值
TOLERANCE_IN = 0.02          # 越界容差（in）
MIN_FONT_PT = 9.0            # 小于这个字号就提醒
OVERLAP_MIN_AREA = 0.04      # 文字重叠判定最小面积（in²）
OVERLAP_RATIO = 0.15         # 重叠面积 / 较小文本框面积

#: 全角标点（按一个汉字宽计算）
_WIDE_PUNCT = set("，。、；：？！“”‘’（）《》【】〈〉…—·")


def _char_width_pt(ch: str, size: float) -> float:
    code = ord(ch)
    if code >= 0x2E80 or ch in _WIDE_PUNCT:
        return size
    if ch == " ":
        return size * 0.28
    return size * 0.55


def _text_width_pt(text: str, size: float) -> float:
    return sum(_char_width_pt(ch, size) for ch in text)


def _line_spacing(paragraph) -> float:
    value = paragraph.line_spacing
    if value is None:
        return 1.0
    if isinstance(value, (int, float)):
        return float(value)
    try:                                 # Length 对象 → 行高倍数
        return max(1.0, float(value.pt) / 12.0)
    except AttributeError:
        return 1.0


def _font_points(text_frame) -> list[float]:
    """该文本框里显式声明过的字号（None 表示继承主题，不参与判断）。"""
    sizes = []
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            size = run.font.size
            if size is not None and size.pt > 0:
                sizes.append(float(size.pt))
    return sizes


def _overflow_risk(shape, text: str) -> str | None:
    """估算文字是否撑破文本框；返回提示文本或 None。"""
    if not text.strip() or not shape.has_text_frame:
        return None
    tf = shape.text_frame
    width_in = shape.width / 914400.0
    height_in = shape.height / 914400.0
    usable_w = width_in - (tf.margin_left + tf.margin_right) / 914400.0
    usable_h = height_in - (tf.margin_top + tf.margin_bottom) / 914400.0
    if usable_w <= 0.05 or usable_h <= 0.02:
        return None
    usable_w_pt, usable_h_pt = usable_w * 72.0, usable_h * 72.0
    total_h = 0.0
    for paragraph in tf.paragraphs:
        if not paragraph.text.strip():
            total_h += 12.0
            continue
        sizes = [float(r.font.size.pt) for r in paragraph.runs
                 if r.font.size is not None]
        size = max(sizes) if sizes else 12.0
        line_w = _text_width_pt(paragraph.text, size)
        lines = max(1, math.ceil(line_w / max(usable_w_pt, 1.0)))
        spacing = _line_spacing(paragraph)
        total_h += lines * size * 1.35 * spacing
    if total_h > usable_h_pt * 1.02:
        return "文字可能溢出文本框（估算需 %.0fpt 高，框内仅 %.0fpt）" % (total_h, usable_h_pt)
    return None


def _is_text_shape(shape) -> bool:
    return bool(getattr(shape, "has_text_frame", False)) and bool(shape.text_frame.text.strip())


def _text_bbox(shape):
    try:
        x, y, w, h = abs_rect(shape)
    except (TypeError, ValueError):
        return None
    if w <= 0 or h <= 0:
        return None
    return (x, y, w, h)


def audit_slide(slide, *, canvas=(SLIDE_W, SLIDE_H),
                tolerance=TOLERANCE_IN, min_font=MIN_FONT_PT):
    """交付前体检：越界 / 文字溢出风险 / 文字重叠 / 字号过小。

    返回 ``[{"level": "error"|"warn", "kind": ..., "shape_id": ..., "name": ..., "message": ...}]``；
    空列表表示没发现问题。启发式检查只用于提醒，不替代 PowerPoint 真机渲染。
    """
    issues = []
    canvas_w, canvas_h = canvas
    text_boxes = []

    for shape in iter_shapes(slide.shapes):
        if shape.shape_type == 6:            # 组合本身不参与
            continue
        text = shape.text_frame.text if shape.has_text_frame else ""
        is_text = bool(text.strip())
        bbox = _text_bbox(shape)
        if bbox is None:
            continue
        x, y, w, h = bbox

        # 1) 越界：文字越界算 error，装饰/图片越界算 warn
        if x < -tolerance or y < -tolerance:
            level = "error" if is_text else "warn"
            issues.append({
                "level": level, "kind": "canvas",
                "shape_id": shape.shape_id, "name": shape.name,
                "message": "元素超出画布左/上边界（x=%.2f, y=%.2f）" % (x, y),
            })
        elif x + w > canvas_w + tolerance or y + h > canvas_h + tolerance:
            level = "error" if is_text else "warn"
            issues.append({
                "level": level, "kind": "canvas",
                "shape_id": shape.shape_id, "name": shape.name,
                "message": "元素超出画布右/下边界（right=%.2f, bottom=%.2f）"
                           % (x + w, y + h),
            })

        # 2) 字号过小
        sizes = _font_points(shape.text_frame) if is_text else []
        if sizes and min(sizes) < min_font:
            issues.append({
                "level": "warn", "kind": "font",
                "shape_id": shape.shape_id, "name": shape.name,
                "message": "字号 %.1fpt 小于建议值 %.0fpt" % (min(sizes), min_font),
            })

        # 3) 文字溢出风险
        if is_text:
            risk = _overflow_risk(shape, text)
            if risk:
                issues.append({
                    "level": "warn", "kind": "overflow",
                    "shape_id": shape.shape_id, "name": shape.name,
                    "message": risk,
                })
            text_boxes.append((shape.shape_id, shape.name, bbox, text))

    # 4) 文字框互相重叠
    for i in range(len(text_boxes)):
        for j in range(i + 1, len(text_boxes)):
            id_a, name_a, rect_a, text_a = text_boxes[i]
            id_b, name_b, rect_b, text_b = text_boxes[j]
            area = _area(rect_a, rect_b)
            if area < OVERLAP_MIN_AREA:
                continue
            smaller = min(rect_a[2] * rect_a[3], rect_b[2] * rect_b[3])
            if smaller > 0 and area / smaller < OVERLAP_RATIO:
                continue
            issues.append({
                "level": "warn", "kind": "overlap",
                "shape_id": id_b, "name": name_b,
                "message": "文字框与 #%s「%s」重叠 %.2f in²" % (id_a, name_a[:16], area),
            })

    issues.sort(key=lambda item: (item["level"] != "error", item["kind"]))
    return issues


def format_issues(issues) -> list[str]:
    """把 audit_slide 的结果格式化成逐行文本。"""
    lines = []
    for issue in issues:
        lines.append("[%s] %-8s #%s %-18s %s" % (
            issue["level"], issue["kind"], issue["shape_id"],
            issue["name"][:18], issue["message"]))
    return lines


def snapshot(slide):
    """把整页形状拍成可比对的 dict：id -> (类型, 几何, 文字, 图片尺寸)。"""
    out = {}

    def walk(shapes):
        for sh in shapes:
            xf = None
            try:
                xf = sh._element.spPr.find(qn("a:xfrm"))
            except AttributeError:
                xf = None
            if xf is None and sh.shape_type == 6:
                xf = sh._element.find(qn("p:grpSpPr")).find(qn("a:xfrm"))
            geo = None
            if xf is not None:
                off, ext = xf.find(qn("a:off")), xf.find(qn("a:ext"))
                geo = (int(off.get("x")), int(off.get("y")),
                       int(ext.get("cx")), int(ext.get("cy")), xf.get("rot"))
            img = list(sh.image.size) if sh.shape_type == 13 else None
            out[sh.shape_id] = (str(sh.shape_type), geo,
                                sh.text_frame.text if sh.has_text_frame else "", img)
            if sh.shape_type == 6:
                walk(sh.shapes)

    walk(slide.shapes)
    return out


def diff(a, b):
    """比较两份 snapshot，返回 (新增, 缺失, 几何变化, 文字变化)。"""
    only_a = sorted(set(a) - set(b))
    only_b = sorted(set(b) - set(a))
    geom, text = [], []
    for k in sorted(set(a) & set(b)):
        if a[k][:2] != b[k][:2]:
            geom.append((k, a[k][1], b[k][1]))
        elif a[k][2] != b[k][2]:
            text.append((k, a[k][2], b[k][2]))
    return only_a, only_b, geom, text


def _area(rect_a, rect_b):
    ax, ay, aw, ah = rect_a
    bx, by, bw, bh = rect_b
    x = max(0.0, min(ax + aw, bx + bw) - max(ax, bx))
    y = max(0.0, min(ay + ah, by + bh) - max(ay, by))
    return x * y


def overlaps(slide, min_area=0.02):
    """列出互相重叠的形状对，按重叠面积降序。"""
    items = []
    for sh in slide.shapes:
        try:
            rect = abs_rect(sh)
        except TypeError:
            continue
        if rect[2] > 0 and rect[3] > 0:
            items.append((sh.shape_id, sh.name, rect,
                          sh.text_frame.text if sh.has_text_frame else ""))
    pairs = []
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            area = _area(items[i][2], items[j][2])
            if area >= min_area:
                pairs.append((area, items[i], items[j]))
    pairs.sort(reverse=True, key=lambda t: t[0])
    return pairs


def flush_pairs(bars, photos, tol=0.004):
    """检查注释条是否贴合照片底边，返回 [(bar, photo, 偏差)]。"""
    result = []
    for bar in bars:
        bx, by, bw, bh = abs_rect(bar)
        best = None
        for photo in photos:
            px, py, pw, ph = abs_rect(photo)
            d = (abs((py + ph) - (by + bh)) + abs(px - bx) + abs(pw - bw))
            if best is None or d < best[0]:
                best = (d, photo)
        if best:
            result.append((bar, best[1], best[0]))
    return result
