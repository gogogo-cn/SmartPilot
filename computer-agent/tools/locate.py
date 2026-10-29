"""
定位层 Tool: locate_text / locate_image / locate_color
让 Agent 精准找到目标元素的坐标
"""
from __future__ import annotations

import sys
from difflib import SequenceMatcher
from typing import Any, Optional

sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.dirname(__file__)))

from core.screen_capture import ScreenCapture
from core.ocr_engine import OcrEngine
from core.utils import image_to_base64

# 复用 perception 模块里的单例模式
_sct = None
_ocr = None


def _get_screen() -> ScreenCapture:
    global _sct
    if _sct is None:
        _sct = ScreenCapture()
    return _sct


def _get_ocr() -> OcrEngine:
    global _ocr
    if _ocr is None:
        cfg = __import__("core.utils", fromlist=["load_config"]).load_config()
        _ocr = OcrEngine(lang=cfg.get("ocr", {}).get("lang", "ch"))
    return _ocr


def locate_text(text: str, match_mode: str = "contains",
                region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    OCR 搜索文字，返回所有匹配项的中心点坐标。

    Args:
        text: 要搜索的文字
        match_mode: exact (完全匹配) / contains (包含) / fuzzy (模糊匹配)
        region: 可选 [x,y,w,h] 限制搜索区域

    Returns:
        {"success": True, "matches": [{text, confidence, center, bbox}...]}
    """
    sct = _get_screen()
    ocr = _get_ocr()

    if region is None:
        img = sct.capture_full()
        offset_x, offset_y = 0, 0
    else:
        x, y, w, h = region
        img = sct.capture_region(x, y, w, h)
        offset_x, offset_y = x, y

    raw_items = ocr.recognize(img)
    matches = []

    mode = match_mode.lower()
    target = text.strip()

    for r in raw_items:
        cand = r.text.strip()
        is_match = False
        if mode == "exact":
            is_match = (cand == target)
        elif mode == "contains":
            is_match = (target in cand) or (cand in target)
        elif mode == "fuzzy":
            ratio = SequenceMatcher(None, cand, target).ratio()
            is_match = ratio >= 0.6
        else:
            is_match = (target in cand)  # fallback

        if is_match:
            abs_center = [r.center[0] + offset_x, r.center[1] + offset_y]
            abs_bbox = [r.bbox[i] + (offset_x if i % 2 == 0 else offset_y)
                        for i in range(len(r.bbox))]
            matches.append({
                "text": r.text,
                "confidence": r.confidence,
                "center": abs_center,
                "bbox": abs_bbox,
            })

    return {
        "success": True,
        "query": text,
        "match_mode": match_mode,
        "match_count": len(matches),
        "matches": matches,
    }


def locate_image(image_path: str, confidence: float = 0.9,
                 region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    模板匹配 — 在屏幕上找目标图片的位置。

    Args:
        image_path: 模板图片路径
        confidence: 匹配阈值 0-1，越高越严格
        region: 可选 [x,y,w,h] 限制搜索区域
    """
    import cv2
    import numpy as np
    from PIL import Image

    sct = _get_screen()

    if region is None:
        screen_bgr = sct.capture_to_numpy()
        offset_x, offset_y = 0, 0
    else:
        x, y, w, h = region
        screen_bgr = sct.capture_to_numpy(x, y, w, h)
        offset_x, offset_y = x, y

    # 读模板
    tpl = cv2.imread(image_path)
    if tpl is None:
        return {"success": False, "error": f"无法读取模板图片: {image_path}"}

    th, tw = tpl.shape[:2]
    sh, sw = screen_bgr.shape[:2]

    if tw > sw or th > sh:
        return {"success": False, "error": "模板图片比屏幕还大"}

    # 多尺度匹配（尝试 0.5x ~ 1.5x）
    best_val = 0.0
    best_pos = None  # (cx, cy)
    best_scale = 1.0

    for scale in [0.6, 0.7, 0.8, 0.9, 1.0, 1.1, 1.2, 1.3, 1.4]:
        new_w = int(tw * scale)
        new_h = int(th * scale)
        if new_w < 10 or new_h < 10 or new_w > sw or new_h > sh:
            continue
        tpl_scaled = cv2.resize(tpl, (new_w, new_h))
        res = cv2.matchTemplate(screen_bgr, tpl_scaled, cv2.TM_CCOEFF_NORMED)
        _, max_val, _, max_loc = cv2.minMaxLoc(res)
        if max_val > best_val:
            best_val = max_val
            cx = max_loc[0] + new_w // 2 + offset_x
            cy = max_loc[1] + new_h // 2 + offset_y
            best_pos = (cx, cy)
            best_scale = scale

    if best_val < confidence:
        return {
            "success": False,
            "error": f"匹配置信度 {best_val:.3f} < 阈值 {confidence}",
            "best_found": best_val,
        }

    return {
        "success": True,
        "confidence": round(best_val, 4),
        "center": list(best_pos),
        "scale": best_scale,
        "template_size": [tw, th],
    }


def locate_color(color: str, tolerance: int = 30,
                 region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    颜色查找 — 在屏幕上找到指定颜色的所有像素中心点。

    Args:
        color: HEX 颜色如 '#FF5733' 或 'FF5733'
        tolerance: 容差 0-255
        region: 可选 [x,y,w,h] 限制区域
    """
    import numpy as np

    sct = _get_screen()

    # 解析 HEX → RGB
    color = color.lstrip("#")
    if len(color) == 6:
        r, g, b = int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16)
    else:
        return {"success": False, "error": f"颜色格式错误: {color}（需要 #RRGGBB）"}

    if region is None:
        img = sct.capture_full()
        offset_x, offset_y = 0, 0
    else:
        x, y, w, h = region
        img = sct.capture_region(x, y, w, h)
        offset_x, offset_y = x, y

    arr = np.array(img)  # RGB
    # 颜色距离（曼哈顿）
    diff = np.abs(arr[:, :, 0].astype(int) - r) + \
           np.abs(arr[:, :, 1].astype(int) - g) + \
           np.abs(arr[:, :, 2].astype(int) - b)
    mask = diff <= tolerance * 3  # 通道合并后容差扩大

    ys, xs = np.where(mask)
    if len(xs) == 0:
        return {
            "success": True,
            "query_color": f"#{color}",
            "tolerance": tolerance,
            "match_count": 0,
            "matches": [],
        }

    # 聚类：把相近的匹配聚在一起，返回簇中心
    points = list(zip(xs.tolist(), ys.tolist()))
    clusters = _cluster_points(points, distance=15)
    matches = []
    for cx, cy, count in clusters:
        matches.append({
            "center": [cx + offset_x, cy + offset_y],
            "pixel_count": count,
        })

    return {
        "success": True,
        "query_color": f"#{color}",
        "tolerance": tolerance,
        "match_count": len(matches),
        "matches": matches,
    }


def _cluster_points(points: list[tuple[int, int]], distance: int = 15) -> list[tuple[int, int, int]]:
    """简单的近邻聚类 — 返回 [(cx, cy, count)]"""
    if not points:
        return []
    clusters: list[list[tuple[int, int]]] = []
    for p in points:
        placed = False
        for cl in clusters:
            cx = sum(q[0] for q in cl) / len(cl)
            cy = sum(q[1] for q in cl) / len(cl)
            if abs(p[0] - cx) <= distance and abs(p[1] - cy) <= distance:
                cl.append(p)
                placed = True
                break
        if not placed:
            clusters.append([p])
    result = []
    for cl in clusters:
        cx = int(sum(q[0] for q in cl) / len(cl))
        cy = int(sum(q[1] for q in cl) / len(cl))
        result.append((cx, cy, len(cl)))
    # 按数量降序，返回前 20 个
    result.sort(key=lambda t: -t[2])
    return result[:20]
