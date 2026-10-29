"""
OCR 引擎封装 — 双路径策略：
  路径 A: 直接 import paddleocr（如果当前 Python 装了）
  路径 B: subprocess 调 Python 3.9（你现成的稳定环境）
复用 AddIn 的架构思想：跨 Python 版本 + JSON 协议通信
"""
from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import tempfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import List, Optional, Tuple

from PIL import Image

logger = logging.getLogger(__name__)


# ─── 子进程 Python 3.9 路径（硬编码你机器上的） ─────────────────────
PYTHON39 = r"C:\Program Files\Python39\python.exe"


@dataclass
class OcrResult:
    text: str
    confidence: float
    bbox: List[int]           # [x1,y1,x2,y2,x3,y3,x4,y4]
    center: Tuple[int, int]

    def to_dict(self) -> dict:
        return asdict(self)


class OcrEngine:
    """OCR 引擎 — 自动选择最佳路径"""

    def __init__(self, lang: str = "ch", use_angle_cls: bool = True):
        self._lang = lang
        self._use_angle_cls = use_angle_cls
        self._mode: Optional[str] = None    # "direct" | "subprocess" | None
        self._ocr = None
        self._initialized = False

    def _ensure_init(self) -> None:
        if self._initialized:
            return

        # 路径 A: 直接 import
        try:
            import paddleocr  # noqa: F401
            self._ocr = paddleocr.PaddleOCR(
                use_angle_cls=self._use_angle_cls,
                lang=self._lang,
                use_gpu=False,
                show_log=False,
                drop_score=0.3,
            )
            self._mode = "direct"
            self._initialized = True
            logger.info("PaddleOCR 直接加载成功 (direct mode)")
            return
        except ImportError:
            logger.info("当前 Python 无 paddleocr，尝试子进程模式...")

        # 路径 B: subprocess 调 Python 3.9
        if os.path.exists(PYTHON39):
            self._mode = "subprocess"
            self._initialized = True
            logger.info("使用 Python 3.9 subprocess 模式: %s", PYTHON39)
            return

        raise RuntimeError(
            "OCR 不可用：当前 Python 没装 paddleocr，且 Python 3.9 也不存在。"
            f"预期路径: {PYTHON39}"
        )

    def recognize(self, img: Image.Image) -> List[OcrResult]:
        """识别 PIL.Image"""
        self._ensure_init()

        if self._mode == "direct":
            import numpy as np
            arr = np.array(img.convert("RGB"))[:, :, ::-1]
            result = self._ocr.ocr(arr, cls=self._use_angle_cls)
            return self._parse_result(result)

        # subprocess 模式：存临时文件 → 让 Python 3.9 识别 → 返回 JSON
        tmp_path = None
        try:
            fd, tmp_path = tempfile.mkstemp(suffix=".png")
            os.close(fd)
            img.save(tmp_path)
            return self._subprocess_ocr(tmp_path)
        finally:
            if tmp_path and os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except OSError:
                    pass

    def recognize_from_file(self, image_path: str) -> List[OcrResult]:
        """从文件路径识别"""
        self._ensure_init()
        if self._mode == "direct":
            result = self._ocr.ocr(image_path, cls=self._use_angle_cls)
            return self._parse_result(result)
        return self._subprocess_ocr(image_path)

    def _subprocess_ocr(self, image_path: str) -> List[OcrResult]:
        """调用 Python 3.9 执行识别脚本，返回解析后的结果"""
        # 内联脚本 — 省文件 IO
        script = f"""
import sys, json
sys.path.insert(0, r'{PYTHON39.replace("python.exe", "Lib/site-packages")}')
import paddleocr
ocr = paddleocr.PaddleOCR(use_angle_cls={str(self._use_angle_cls)}, lang='{self._lang}', use_gpu=False, show_log=False, drop_score=0.3)
result = ocr.ocr(r'{image_path.replace(chr(92), chr(47))}', cls={str(self._use_angle_cls)})
items = []
if result and result[0]:
    for line in result[0]:
        if not line: continue
        box, (text, conf) = line[0], line[1]
        xs = [int(p[0]) for p in box]
        ys = [int(p[1]) for p in box]
        cx, cy = sum(xs)//len(xs), sum(ys)//len(ys)
        flat = []
        for p in box: flat.extend([int(p[0]), int(p[1])])
        items.append({{"text": text, "confidence": round(float(conf),4), "bbox": flat, "center": [cx, cy]}})
print(json.dumps({{"items": items}}))
""".strip()

        proc = subprocess.run(
            [PYTHON39, "-c", script],
            capture_output=True, text=True, timeout=60,
            encoding="utf-8", errors="replace",
        )

        if proc.returncode != 0:
            raise RuntimeError(f"OCR subprocess failed: {proc.stderr[-500:]}")

        # 从 stdout 最后一行找 JSON
        for line in proc.stdout.strip().splitlines()[::-1]:
            line = line.strip()
            if line.startswith("{"):
                data = json.loads(line)
                items = data.get("items", [])
                return [
                    OcrResult(
                        text=i["text"],
                        confidence=i["confidence"],
                        bbox=i["bbox"],
                        center=tuple(i["center"]),
                    )
                    for i in items
                ]

        return []

    @staticmethod
    def _parse_result(result) -> List[OcrResult]:
        results: List[OcrResult] = []
        if not result or not result[0]:
            return results

        for line in result[0]:
            if line is None or len(line) < 2:
                continue
            box = line[0]
            text_conf = line[1]
            if box is None or text_conf is None:
                continue

            text = str(text_conf[0])
            conf = float(text_conf[1]) if len(text_conf) > 1 else 0.0

            flat_box: List[int] = []
            xs, ys = [], []
            for pt in box:
                flat_box.extend([int(pt[0]), int(pt[1])])
                xs.append(int(pt[0]))
                ys.append(int(pt[1]))

            center = (int(sum(xs) / len(xs)), int(sum(ys) / len(ys)))
            results.append(OcrResult(text=text, confidence=round(conf, 4),
                                     bbox=flat_box, center=center))

        return results
