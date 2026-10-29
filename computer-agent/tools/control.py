"""
操作层 Tool: mouse_* / key_* / type_text / wait
让 Agent 控制键鼠
"""
from __future__ import annotations

import sys
from typing import Any, Optional

sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.dirname(__file__)))

from core.mouse_keyboard import MouseController, KeyboardController, wait_ms


def mouse_move(x: int, y: int) -> dict[str, Any]:
    """鼠标移动到绝对坐标"""
    return MouseController.move(x, y)


def mouse_click(x: Optional[int] = None, y: Optional[int] = None,
                button: str = "left", clicks: int = 1) -> dict[str, Any]:
    """鼠标点击。x,y 可选 — 不填则点击当前位置"""
    return MouseController.click(x, y, button, clicks)


def mouse_drag(from_x: int, from_y: int, to_x: int, to_y: int,
               duration: float = 0.5) -> dict[str, Any]:
    """拖拽鼠标"""
    return MouseController.drag(from_x, from_y, to_x, to_y, duration)


def mouse_scroll(amount: int, direction: str = "down") -> dict[str, Any]:
    """滚动鼠标滚轮"""
    return MouseController.scroll(amount, direction)


def key_press(key: str) -> dict[str, Any]:
    """按下并释放一个键"""
    return KeyboardController.press(key)


def key_hotkey(keys: list[str]) -> dict[str, Any]:
    """按下组合键，如 ['ctrl', 'c']"""
    return KeyboardController.hotkey(keys)


def type_text(text: str) -> dict[str, Any]:
    """输入文本（通过剪贴板粘贴，支持中文）"""
    return KeyboardController.type_text(text)


def wait(milliseconds: int) -> dict[str, Any]:
    """等待若干毫秒，让 UI 响应"""
    return wait_ms(milliseconds)
