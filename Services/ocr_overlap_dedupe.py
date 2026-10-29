# -*- coding: utf-8 -*-
"""去掉同一行相邻检测框交界处被识别两次的字。

检测框扩边后会互相压住一个字，识别时这个字会同时出现在左右两个框里，
例如「出生」和「生1654年12月20日」。这里只在几何重叠和文字边界都对得上时，
把这个字留给更靠近它的那个框，另一边删掉。检测框保持原文文字框，不再按字宽另裁一条边。
"""


def _char_units(ch):
    if ch.isspace() and ord(ch) <= 0x007E:
        return 0.33
    if ord(ch) <= 0x007E:
        return 0.55
    return 1.0


def _text_units(text):
    return sum(_char_units(ch) for ch in text) or 1.0


def _point_xy(p):
    if isinstance(p, dict) and "x" in p and "y" in p:
        try:
            return float(p["x"]), float(p["y"])
        except (TypeError, ValueError):
            return None
    if isinstance(p, (list, tuple)) and len(p) >= 2:
        try:
            return float(p[0]), float(p[1])
        except (TypeError, ValueError):
            return None
    return None


def _bounds(item):
    box = item.get("box") if isinstance(item, dict) else None
    if not isinstance(box, (list, tuple)) or len(box) < 4:
        return None
    pts = []
    for p in box:
        xy = _point_xy(p)
        if xy is None:
            return None
        pts.append(xy)
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    left, right = min(xs), max(xs)
    top, bottom = min(ys), max(ys)
    if right - left < 1 or bottom - top < 1:
        return None
    return {
        "left": left,
        "right": right,
        "top": top,
        "bottom": bottom,
    }


def _core_text(text, side):
    """匹配时忽略交界外侧的空白。返回 (参与匹配的文本, 被忽略的空白长度)。"""
    if side == "left":
        core = text.rstrip()
        return core, len(text) - len(core)
    core = text.lstrip()
    return core, len(text) - len(core)


def _v_overlap_ratio(a, b):
    inter = min(a["bottom"], b["bottom"]) - max(a["top"], b["top"])
    if inter <= 0:
        return 0.0
    shorter = min(a["bottom"] - a["top"], b["bottom"] - b["top"])
    if shorter <= 0:
        return 0.0
    return inter / shorter


def _same_line(a, b):
    ha = a["bottom"] - a["top"]
    hb = b["bottom"] - b["top"]
    if min(ha, hb) <= 0:
        return False
    if max(ha, hb) / min(ha, hb) > 2.2:
        return False
    return _v_overlap_ratio(a, b) >= 0.5


def _cluster_lines(records):
    ordered = sorted(records, key=lambda r: ((r["top"] + r["bottom"]) / 2, r["left"]))
    lines = []
    for rec in ordered:
        placed = False
        for line in lines:
            anchor = line[0]
            if _same_line(anchor, rec):
                line.append(rec)
                placed = True
                break
        if not placed:
            lines.append([rec])
    return lines


def _matching_prefix_len(a_text, b_text, shared_units):
    """重叠区往往比真正重复的字更宽（检测框扩边）。只要这段字装得进重叠区就去掉。"""
    if shared_units < 0.45:
        return 0
    max_n = min(len(a_text), len(b_text))
    best = 0
    for n in range(1, max_n + 1):
        piece = a_text[-n:]
        if piece != b_text[:n]:
            continue
        if any(ch.isspace() for ch in piece):
            continue
        units = _text_units(piece)
        if units <= shared_units + 0.55:
            best = n
    return best


def _keep_on_left(a, b, glyph_x):
    a_mid = (a["left"] + a["right"]) / 2
    b_mid = (b["left"] + b["right"]) / 2
    da = abs(glyph_x - a_mid)
    db = abs(glyph_x - b_mid)
    slack = max(2.0, 0.15 * min(a["right"] - a["left"], b["right"] - b["left"]))
    if abs(da - db) > slack:
        return da < db
    ca = float(a["item"].get("confidence") or 0)
    cb = float(b["item"].get("confidence") or 0)
    if abs(ca - cb) > 0.02:
        return ca >= cb
    return True


def _try_pair(a, b):
    if a.get("is_formula") or b.get("is_formula"):
        return
    if not _same_line(a, b):
        return
    overlap = min(a["right"], b["right"]) - max(a["left"], b["left"])
    a_text = a["item"].get("text") or ""
    b_text = b["item"].get("text") or ""
    a_core, a_trail = _core_text(a_text, "left")
    b_core, b_lead = _core_text(b_text, "right")
    if not a_core or not b_core:
        return
    a_unit = (a["right"] - a["left"]) / _text_units(a_core)
    b_unit = (b["right"] - b["left"]) / _text_units(b_core)
    if a_unit <= 0 or b_unit <= 0:
        return
    if overlap > 1:
        shared = min(overlap / a_unit, overlap / b_unit)
    else:
        # 框几乎贴住，中间只留一条缝，交界上的字仍会进两个框
        gap = max(a["left"], b["left"]) - min(a["right"], b["right"])
        if gap > 0.45 * min(a_unit, b_unit):
            return
        shared = 1.0
    n = _matching_prefix_len(a_core, b_core, shared)
    if n <= 0:
        return
    piece = a_core[-n:]
    piece_units = _text_units(piece)
    span_a = piece_units * a_unit
    span_b = piece_units * b_unit
    glyph_x = ((a["right"] - span_a / 2) + (b["left"] + span_b / 2)) / 2
    if _keep_on_left(a, b, glyph_x):
        new_text = b_text[b_lead + n:]
        b["item"]["text"] = new_text
        _sync_original(b["item"], b_text, new_text)
    else:
        new_text = a_text[:len(a_text) - a_trail - n]
        a["item"]["text"] = new_text
        _sync_original(a["item"], a_text, new_text)


def _sync_original(item, previous, new_text):
    if "originalText" not in item:
        return
    original = item.get("originalText")
    if original is None or original == previous:
        item["originalText"] = new_text


def _quad_points(quad):
    if not isinstance(quad, (list, tuple)) or len(quad) < 4:
        return None
    pts = []
    for p in quad[:4]:
        xy = _point_xy(p)
        if xy is None:
            return None
        pts.append(xy)
    return pts


def _nearly_horizontal(pts):
    dx = pts[1][0] - pts[0][0]
    dy = pts[1][1] - pts[0][1]
    if abs(dx) < 1 and abs(dy) < 1:
        return True
    return abs(dy) <= abs(dx) * 0.25


def exclusive_recognition_quads(quads):
    """检测框照原样留给显示。这里只决定送进识别器的区域。

    同一行上，右边的框如果压住了左边的框，交叠的那一截只留给左边识别。
    右边送进去的图从左边框的右缘开始。不看具体是哪个字。
    """
    if not isinstance(quads, (list, tuple)):
        return quads
    parsed = [_quad_points(q) for q in quads]
    rec = [[list(p) for p in pts] if pts else None for pts in parsed]
    records = []
    for i, pts in enumerate(parsed):
        if pts is None or not _nearly_horizontal(pts):
            continue
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        left, right = min(xs), max(xs)
        top, bottom = min(ys), max(ys)
        if right - left < 2 or bottom - top < 2:
            continue
        records.append({
            "i": i,
            "left": left,
            "right": right,
            "top": top,
            "bottom": bottom,
        })
    for line in _cluster_lines(records):
        line.sort(key=lambda r: r["left"])
        for k in range(1, len(line)):
            cur = line[k]
            cut = cur["left"]
            for prev in line[:k]:
                overlap = min(prev["right"], cur["right"]) - max(prev["left"], cur["left"])
                if overlap <= 1:
                    continue
                if prev["right"] < cur["right"] - 2:
                    cut = max(cut, prev["right"])
            if cut <= cur["left"] + 1 or cut >= cur["right"] - 2:
                continue
            rec[cur["i"]] = [
                [cut, cur["top"]],
                [cur["right"], cur["top"]],
                [cur["right"], cur["bottom"]],
                [cut, cur["bottom"]],
            ]
    out = []
    for i, q in enumerate(rec):
        out.append(quads[i] if q is None else q)
    return out


def dedupe_overlapping_items(items):
    """返回新列表。没有坐标或公式块的项原样保留，顺序不变。"""
    if not isinstance(items, list) or len(items) < 2:
        return items
    indexed = []
    active = []
    for item in items:
        if not isinstance(item, dict):
            indexed.append({"skip": True, "item": item})
            continue
        copied = dict(item)
        box = copied.get("box")
        if isinstance(box, list):
            copied["box"] = [dict(p) if isinstance(p, dict) else list(p) for p in box]
        bounds = _bounds(copied)
        rec = {"skip": bounds is None or bool(copied.get("is_formula")), "item": copied}
        if bounds is not None:
            rec.update(bounds)
        indexed.append(rec)
        if not rec["skip"]:
            active.append(rec)
    for line in _cluster_lines(active):
        line.sort(key=lambda r: r["left"])
        for i in range(len(line)):
            seam = max(1.0, 0.6 * (line[i]["bottom"] - line[i]["top"]))
            for j in range(i + 1, len(line)):
                if line[j]["left"] >= line[i]["right"] + seam:
                    break
                _try_pair(line[i], line[j])
    out = []
    for rec in indexed:
        text = rec["item"].get("text") if isinstance(rec["item"], dict) else None
        if rec["skip"] or text:
            out.append(rec["item"])
    return out
