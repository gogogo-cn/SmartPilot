"""
键鼠控制封装 — pyautogui + 安全边界 + safe_mode 支持
"""
from __future__ import annotations

import logging
import time
from typing import Optional, Tuple

import pyautogui

from .utils import validate_coords, is_safe_mode, RetryableError, FatalError

logger = logging.getLogger(__name__)

# pyautogui 全局配置
pyautogui.FAILSAFE = True       # 鼠标甩到 (0,0) 自动中断
pyautogui.PAUSE = 0.05          # 每个 pyautogui 操作后短暂停顿


def _screen_size() -> Tuple[int, int]:
    return pyautogui.size()


class MouseController:
    """鼠标操作封装"""

    @staticmethod
    def move(x: int, y: int, duration: float = 0.3) -> dict:
        screen_w, screen_h = _screen_size()
        x, y = validate_coords(x, y, screen_w, screen_h)

        if is_safe_mode():
            return {"action": "mouse_move", "x": x, "y": y, "executed": False, "safe_mode": True}

        pyautogui.moveTo(x, y, duration=duration)
        return {"action": "mouse_move", "x": x, "y": y, "executed": True}

    @staticmethod
    def click(x: Optional[int] = None, y: Optional[int] = None,
              button: str = "left", clicks: int = 1) -> dict:
        """点击；如果提供 x,y 则先移动再点击"""
        button = button.lower()
        if button not in ("left", "right", "middle"):
            raise RetryableError(f"无效鼠标键: {button}，可选 left/right/middle")

        screen_w, screen_h = _screen_size()
        if x is not None and y is not None:
            x, y = validate_coords(int(x), int(y), screen_w, screen_h)
        else:
            x, y = pyautogui.position()

        if is_safe_mode():
            return {"action": "mouse_click", "x": x, "y": y, "button": button,
                    "clicks": clicks, "executed": False, "safe_mode": True}

        pyautogui.click(x=x, y=y, button=button, clicks=clicks)
        return {"action": "mouse_click", "x": x, "y": y, "button": button,
                "clicks": clicks, "executed": True}

    @staticmethod
    def drag(from_x: int, from_y: int, to_x: int, to_y: int,
             duration: float = 0.5) -> dict:
        screen_w, screen_h = _screen_size()
        from_x, from_y = validate_coords(from_x, from_y, screen_w, screen_h)
        to_x, to_y = validate_coords(to_x, to_y, screen_w, screen_h)

        if is_safe_mode():
            return {"action": "mouse_drag", "from": [from_x, from_y],
                    "to": [to_x, to_y], "executed": False, "safe_mode": True}

        pyautogui.moveTo(from_x, from_y)
        pyautogui.drag(to_x - from_x, to_y - from_y, duration=duration)
        return {"action": "mouse_drag", "from": [from_x, from_y],
                "to": [to_x, to_y], "executed": True}

    @staticmethod
    def scroll(amount: int, direction: str = "down") -> dict:
        direction = direction.lower()
        if direction not in ("up", "down", "left", "right"):
            raise RetryableError(f"无效滚动方向: {direction}")

        if is_safe_mode():
            return {"action": "mouse_scroll", "amount": amount,
                    "direction": direction, "executed": False, "safe_mode": True}

        if direction == "up":
            pyautogui.scroll(amount)
        elif direction == "down":
            pyautogui.scroll(-amount)
        elif direction == "left":
            pyautogui.hscroll(-amount)
        else:
            pyautogui.hscroll(amount)

        return {"action": "mouse_scroll", "amount": amount,
                "direction": direction, "executed": True}

    @staticmethod
    def position() -> Tuple[int, int]:
        return pyautogui.position()


class KeyboardController:
    """键盘操作封装"""

    @staticmethod
    def press(key: str) -> dict:
        if is_safe_mode():
            return {"action": "key_press", "key": key, "executed": False, "safe_mode": True}
        pyautogui.press(key)
        return {"action": "key_press", "key": key, "executed": True}

    @staticmethod
    def hotkey(keys: list[str]) -> dict:
        if is_safe_mode():
            return {"action": "key_hotkey", "keys": keys, "executed": False, "safe_mode": True}
        pyautogui.hotkey(*keys)
        return {"action": "key_hotkey", "keys": keys, "executed": True}

    @staticmethod
    def type_text(text: str, interval: float = 0.02) -> dict:
        if is_safe_mode():
            return {"action": "type_text", "text": text[:50] + ("..." if len(text) > 50 else ""),
                    "executed": False, "safe_mode": True}
        # pyautogui.typewrite 不支持中文，用 pyperclip 粘贴
        import pyperclip
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.1)
        return {"action": "type_text", "length": len(text), "executed": True}


def wait_ms(milliseconds: int) -> dict:
    time.sleep(max(0, milliseconds) / 1000.0)
    return {"action": "wait", "milliseconds": milliseconds, "executed": True}
