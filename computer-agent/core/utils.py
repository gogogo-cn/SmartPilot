"""
通用工具函数：坐标转换 / base64 编码 / 错误分类 / 安全边界
"""
from __future__ import annotations

import base64
import io
import os
import sys
from dataclasses import dataclass
from typing import Optional


# ─── 错误分类 ────────────────────────────────────────────────────────

@dataclass
class AgentError(Exception):
    """Agent 操作错误基类"""
    message: str
    category: str = "unknown"   # retry | user_intervention | fatal
    suggestion: str = ""

    def __str__(self) -> str:
        base = f"[{self.category}] {self.message}"
        if self.suggestion:
            base += f" (建议: {self.suggestion})"
        return base


class RetryableError(AgentError):
    """可重试的错误"""
    def __init__(self, message: str, suggestion: str = "重试一次"):
        super().__init__(message, category="retry", suggestion=suggestion)


class UserInterventionError(AgentError):
    """需要用户介入的错误"""
    def __init__(self, message: str, suggestion: str = ""):
        super().__init__(message, category="user_intervention", suggestion=suggestion)


class FatalError(AgentError):
    """致命错误"""
    def __init__(self, message: str, suggestion: str = "停止操作"):
        super().__init__(message, category="fatal", suggestion=suggestion)


# ─── 坐标校验 ────────────────────────────────────────────────────────

def clamp(x: int, lo: int, hi: int) -> int:
    """限制坐标范围"""
    return max(lo, min(hi, x))


def validate_region(region: Optional[list[int]], screen_w: int, screen_h: int) -> tuple[int, int, int, int]:
    """校验 region 参数，返回规范化的 (x, y, w, h)"""
    if region is None:
        return 0, 0, screen_w, screen_h
    if len(region) != 4:
        raise AgentError(f"region 需要 4 个元素 [x,y,w,h]，实际 {len(region)} 个",
                         category="user_intervention")
    x, y, w, h = region
    if w <= 0 or h <= 0:
        raise AgentError(f"region 宽高必须正数，实际 w={w} h={h}",
                         category="user_intervention")
    # clamp 到屏幕范围
    x = clamp(x, 0, screen_w - 1)
    y = clamp(y, 0, screen_h - 1)
    w = min(w, screen_w - x)
    h = min(h, screen_h - y)
    return x, y, w, h


def validate_coords(x: int, y: int, screen_w: int, screen_h: int) -> tuple[int, int]:
    """校验并 clamp 坐标"""
    x = clamp(int(x), 0, screen_w - 1)
    y = clamp(int(y), 0, screen_h - 1)
    return x, y


# ─── 图片编解码 ───────────────────────────────────────────────────────

def image_to_base64(img) -> str:
    """PIL.Image → base64 PNG"""
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return base64.b64encode(buf.getvalue()).decode("ascii")


def base64_to_image(b64: str):
    """base64 PNG → PIL.Image"""
    from PIL import Image
    data = base64.b64decode(b64)
    return Image.open(io.BytesIO(data))


# ─── 配置加载 ────────────────────────────────────────────────────────

def load_config(config_path: Optional[str] = None) -> dict:
    """加载 YAML 配置，找不到则返回默认"""
    try:
        import yaml
    except ImportError:
        return _default_config()

    if config_path is None:
        config_path = os.path.join(os.path.dirname(__file__), "..", "config", "default.yaml")

    config_path = os.path.normpath(config_path)
    if not os.path.exists(config_path):
        return _default_config()

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f) or {}
        # 合并默认值
        return _merge_default(cfg)
    except Exception:
        return _default_config()


def _default_config() -> dict:
    return {
        "ocr": {"engine": "paddleocr", "lang": "ch", "use_angle_cls": True},
        "security": {"safe_mode": False, "confirm_required": False, "fail_safe": True},
        "vision": {"enabled": False, "provider": "glm-4v", "api_key": "", "api_base": ""},
    }


def _merge_default(cfg: dict) -> dict:
    d = _default_config()
    for k, v in cfg.items():
        if isinstance(v, dict) and k in d and isinstance(d[k], dict):
            d[k].update(v)
        else:
            d[k] = v
    return d


# ─── 安全模式标志（进程级） ──────────────────────────────────────────

_SAFE_MODE = False


def set_safe_mode(enabled: bool) -> None:
    global _SAFE_MODE
    _SAFE_MODE = bool(enabled)


def is_safe_mode() -> bool:
    return _SAFE_MODE
