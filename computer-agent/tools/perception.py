"""
感知层 Tool: screen_capture / screen_ocr / screen_describe / screen_get_info
让 Agent "看见" 屏幕上的内容
"""
from __future__ import annotations

import json
import logging
import sys
from typing import Any, Optional

# 让 core 模块可以被导入
sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.dirname(__file__)))

from core.screen_capture import ScreenCapture
from core.ocr_engine import OcrEngine
from core.utils import image_to_base64, load_config

logger = logging.getLogger(__name__)

# 模块级单例
_sct: Optional[ScreenCapture] = None
_ocr: Optional[OcrEngine] = None


def _get_screen() -> ScreenCapture:
    global _sct
    if _sct is None:
        _sct = ScreenCapture()
    return _sct


def _get_ocr() -> OcrEngine:
    global _ocr
    if _ocr is None:
        cfg = load_config()
        _ocr = OcrEngine(lang=cfg.get("ocr", {}).get("lang", "ch"))
    return _ocr


def screen_capture(region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    截图并返回 base64 PNG 图片。

    Args:
        region: 可选 [x, y, w, h] 区域；不传则全屏

    Returns:
        {"success": True, "image_base64": "...", "width": w, "height": h}
    """
    sct = _get_screen()
    if region is None:
        img = sct.capture_full()
    else:
        x, y, w, h = region
        img = sct.capture_region(x, y, w, h)

    b64 = image_to_base64(img)
    return {
        "success": True,
        "image_base64": b64,
        "width": img.width,
        "height": img.height,
        "format": "png",
    }


def screen_ocr(region: Optional[list[int]] = None,
               lang: Optional[str] = None) -> dict[str, Any]:
    """
    截图 + OCR 识别，返回每个文本块的坐标和置信度。

    Args:
        region: 可选 [x, y, w, h] 只识别区域
        lang: 语言代码，如 'ch' 'en' 'ch_en'

    Returns:
        {"success": True, "screen_size": [w, h], "items": [{text, confidence, bbox, center}...]}
        注意: bbox 坐标是相对屏幕左上角的绝对坐标 (OCR 内部偏移已加回)
    """
    sct = _get_screen()
    ocr = _get_ocr()

    screen_w, screen_h = sct.screen_size

    if region is None:
        img = sct.capture_full()
        offset_x, offset_y = 0, 0
    else:
        x, y, w, h = region
        img = sct.capture_region(x, y, w, h)
        offset_x, offset_y = x, y

    raw_items = ocr.recognize(img)

    # 把相对图片坐标转为屏幕绝对坐标
    items = []
    for r in raw_items:
        abs_bbox = [r.bbox[i] + (offset_x if i % 2 == 0 else offset_y)
                    for i in range(len(r.bbox))]
        abs_center = (r.center[0] + offset_x, r.center[1] + offset_y)
        items.append({
            "text": r.text,
            "confidence": r.confidence,
            "bbox": abs_bbox,
            "center": list(abs_center),
        })

    return {
        "success": True,
        "screen_size": [screen_w, screen_h],
        "region_offset": [offset_x, offset_y],
        "item_count": len(items),
        "items": items,
    }


def screen_describe(region: Optional[list[int]] = None,
                    detail: str = "low") -> dict[str, Any]:
    """
    截图 + 视觉模型描述 UI 内容。
    视觉模型可选，未配置时返回降级提示。
    """
    cfg = load_config()
    vision_cfg = cfg.get("vision", {})
    if not vision_cfg.get("enabled", False):
        return {
            "success": False,
            "error": "视觉模型未启用。在 config/default.yaml 中设置 vision.enabled=true",
            "hint": "也可以用 screen_ocr + LLM 自己推理屏幕含义",
        }

    # 暂时返回占位 —— 视觉模型在 Step 6 扩展
    return {
        "success": True,
        "description": "视觉模型模块待接入 (Step 6)",
        "note": "请在 vision_client.py 实现后替换",
    }


def screen_get_info() -> dict[str, Any]:
    """获取屏幕状态速查信息"""
    sct = _get_screen()
    import pyautogui

    screen_w, screen_h = sct.screen_size
    mx, my = pyautogui.position()

    try:
        import psutil
        # 活动窗口信息（简化版，跨平台）
        active_win = _get_active_window_info()
    except Exception:
        active_win = None

    return {
        "success": True,
        "screen_size": [screen_w, screen_h],
        "mouse_position": [mx, my],
        "primary_monitor": sct._primary,
        "active_window": active_win,
    }


def _get_active_window_info() -> Optional[dict]:
    """获取活动窗口信息（Windows 优先）"""
    try:
        if sys.platform == "win32":
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            hwnd = user32.GetForegroundWindow()
            length = user32.GetWindowTextLengthW(hwnd)
            buf = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buf, length + 1)
            return {"title": buf.value, "handle": hwnd}
    except Exception:
        pass
    return None
