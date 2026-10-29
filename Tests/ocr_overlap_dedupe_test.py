# -*- coding: utf-8 -*-
"""与 Tests/ocr-overlap-dedupe.test.js 同一组例子，锁住 Python 侧行为。"""
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "Services"))
from ocr_overlap_dedupe import dedupe_overlapping_items, exclusive_recognition_quads


def box(l, t, r, b):
    return [[l, t], [r, t], [r, b], [l, b]]


def texts(items):
    return [it["text"] for it in items]


class OverlapDedupeTest(unittest.TestCase):
    def test_birth_date_keeps_sheng_on_the_left(self):
        items = [
            {"text": "出生", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "生1654年12月20日", "confidence": 0.97, "box": box(132, 40, 400.8, 72)},
        ]
        out = dedupe_overlapping_items(items)
        self.assertEqual(texts(out), ["出生", "1654年12月20日"])
        self.assertEqual(out[1]["box"], box(132, 40, 400.8, 72))
        self.assertEqual(items[1]["text"], "生1654年12月20日")
        self.assertEqual(texts(dedupe_overlapping_items(out)), ["出生", "1654年12月20日"])

    def test_birth_date_wider_overlap_still_drops_one_sheng(self):
        # 日期框盖住「出生」将近两个字宽，文字只重复一个「生」
        items = [
            {"text": "出生", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "生1654年12月20日", "confidence": 0.97, "box": box(100, 40, 420, 72)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["出生", "1654年12月20日"])

    def test_birth_date_small_gap_still_drops_sheng(self):
        items = [
            {"text": "出生", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "生1654年12月20日", "confidence": 0.97, "box": box(172, 40, 440, 72)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["出生", "1654年12月20日"])

    def test_gap_keeps_both(self):
        items = [
            {"text": "甲乙", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "乙丙丁", "confidence": 0.97, "box": box(180, 40, 448, 72)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["甲乙", "乙丙丁"])

    def test_different_lines(self):
        items = [
            {"text": "出生", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "生1654年12月20日", "confidence": 0.97, "box": box(132, 90, 400.8, 122)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["出生", "生1654年12月20日"])

    def test_different_boundary_text(self):
        items = [
            {"text": "出生", "confidence": 1, "box": box(100, 40, 164, 72)},
            {"text": "于1654年12月20日", "confidence": 0.97, "box": box(132, 40, 400.8, 72)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["出生", "于1654年12月20日"])

    def test_glyph_closer_to_right_box(self):
        items = [
            {"text": "甲乙生", "confidence": 0.9, "box": box(0, 10, 96, 42)},
            {"text": "生丙", "confidence": 0.9, "box": box(64, 10, 128, 42)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["甲乙", "生丙"])

    def test_identical_boxes_collapse(self):
        items = [
            {"text": "出生", "confidence": 0.99, "box": box(0, 0, 64, 32)},
            {"text": "出生", "confidence": 0.91, "box": box(0, 0, 64, 32)},
        ]
        self.assertEqual(texts(dedupe_overlapping_items(items)), ["出生"])

    def test_overlap_is_cut_out_of_the_right_recognition_crop(self):
        quads = [box(100, 40, 164, 72), box(132, 40, 400.8, 72)]
        rec = exclusive_recognition_quads(quads)
        self.assertEqual(rec[0], quads[0])
        self.assertEqual(rec[1][0][0], 164)
        self.assertEqual(rec[1][1], [400.8, 40])
        self.assertEqual(quads[1], box(132, 40, 400.8, 72))

    def test_gap_recognition_crop_stays_the_detection_box(self):
        quads = [box(100, 40, 164, 72), box(200, 40, 468, 72)]
        self.assertEqual(exclusive_recognition_quads(quads), quads)

    def test_vertical_overlap_does_not_clip_recognition(self):
        quads = [box(100, 40, 164, 72), box(132, 90, 400, 122)]
        self.assertEqual(exclusive_recognition_quads(quads), quads)

    def test_clip_that_would_erase_the_box_is_skipped(self):
        quads = [box(100, 40, 400, 72), box(132, 40, 164, 72)]
        self.assertEqual(exclusive_recognition_quads(quads)[1], quads[1])


if __name__ == "__main__":
    unittest.main()
