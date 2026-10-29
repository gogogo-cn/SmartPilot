"""
屏幕截图模块 — 基于 mss 高性能截屏
比 PIL.ImageGrab 快 3-5x，直接走操作系统缓冲区
"""
from __future__ import annotations

from typing import Optional, Tuple

import mss
import numpy as np
from PIL import Image


class ScreenCapture:
    """屏幕截图封装"""

    def __init__(self):
        self._sct = mss.mss()
        self._monitors = self._sct.monitors  # [monitor0, monitor1, ...]
        self._primary = self._monitors[0]      # 主屏幕

    @property
    def screen_size(self) -> Tuple[int, int]:
        """主屏幕分辨率 (width, height)"""
        return self._primary["width"], self._primary["height"]

    def capture_full(self) -> Image.Image:
        """全屏截图 → PIL.Image"""
        shot = self._sct.grab(self._primary)
        return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    def capture_region(self, x: int, y: int, w: int, h: int) -> Image.Image:
        """区域截图 → PIL.Image"""
        monitor = {"left": x, "top": y, "width": w, "height": h}
        shot = self._sct.grab(monitor)
        return Image.frombytes("RGB", shot.size, shot.bgra, "raw", "BGRX")

    def capture_to_numpy(self, x: int = 0, y: int = 0,
                         w: Optional[int] = None, h: Optional[int] = None) -> np.ndarray:
        """截图到 numpy ndarray (BGR 格式，OpenCV 友好)"""
        if w is None:
            w, h = self.screen_size
        monitor = {"left": x, "top": y, "width": w, "height": h}
        shot = self._sct.grab(monitor)
        # BGRA → BGR (去掉 alpha)
        arr = np.array(shot, dtype=np.uint8)
        return arr[:, :, :3]

    def close(self) -> None:
        self._sct.close()

    def __enter__(self) -> "ScreenCapture":
        return self

    def __exit__(self, *args) -> None:
        self.close()
