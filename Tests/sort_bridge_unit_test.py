# -*- coding: utf-8 -*-
"""Tests/sort_bridge_unit_test.py — py-excel-bridge 排序"还原原生对话框"纯函数回归。

背景（2026-10-16 实测）：
  ① 桥里 XL_YES=-4163 / XL_NO=-4152 是从其它枚举（xlUp 等）抄错的常量，
     传给 Sort.Header 直接 E_INVALIDARG（现代 Sort 对象路径必崩，实测 e2e 全红）；
     正确值 xlYes=1 / xlNo=2。legacy rng.Sort 虽不报错，但会被当成 xlGuess——
     首行是数字时猜"无表头"，把表头行一起排进数据（静默错序）。
  ② 原生「排序」对话框的排序依据（单元格值/单元格颜色/字体颜色/条件格式图标）、
     选项（区分大小写/方向/排序方法）与排序后预览回传，都必须在桥侧可用。

运行：py -3.9 Tests/sort_bridge_unit_test.py
"""
import importlib.util
import os
import sys
import unittest

BRIDGE = os.path.join(os.path.dirname(__file__), "..", "Tools", "py-excel-bridge.py")


def _load_bridge():
    spec = importlib.util.spec_from_file_location("py_excel_bridge_under_test", BRIDGE)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)  # 有 __main__ 保护，导入不会起服务
    return mod


bridge = _load_bridge()


class SortConstantsTest(unittest.TestCase):
    def test_yes_no_constants_are_xlyesno_guess(self):
        # xlGuess=0 / xlNo=2 / xlYes=1（XlYesNoGuess）。旧值 -4163/-4152 是错的。
        self.assertEqual(bridge.XL_YES, 1)
        self.assertEqual(bridge.XL_NO, 2)

    def test_order_constants(self):
        self.assertEqual(bridge.XL_ASCENDING, 1)
        self.assertEqual(bridge.XL_DESCENDING, 2)

    def test_sort_on_map(self):
        self.assertEqual(bridge._SORT_ON["values"], 0)
        self.assertEqual(bridge._SORT_ON["cellcolor"], 1)
        self.assertEqual(bridge._SORT_ON["fontcolor"], 2)
        self.assertEqual(bridge._SORT_ON["icon"], 3)

    def test_icon_sets_list_matches_xliconset_enum(self):
        names = bridge._XL_ICON_SETS
        self.assertEqual(names[0], "xl3Arrows")
        self.assertEqual(names[2], "xl3Flags")
        self.assertIn("xl3TrafficLights1", names)
        self.assertIn("xl5Boxes", names)


class ParseSortColorTest(unittest.TestCase):
    def test_hex_and_bare(self):
        self.assertEqual(bridge._parse_sort_color("#FF0000"), 0xFF0000)
        self.assertEqual(bridge._parse_sort_color("FF0000"), 0xFF0000)
        self.assertEqual(bridge._parse_sort_color("008000"), 0x008000)

    def test_rgb_triple_and_decimal(self):
        self.assertEqual(bridge._parse_sort_color("255,0,0"), 0xFF0000)
        self.assertEqual(bridge._parse_sort_color(" 255 , 0 , 0 "), 0xFF0000)
        self.assertEqual(bridge._parse_sort_color("255"), 255)

    def test_chinese_color_names(self):
        self.assertEqual(bridge._parse_sort_color("红色"), 0xFF0000)
        self.assertEqual(bridge._parse_sort_color("绿"), 0x008000)
        self.assertEqual(bridge._parse_sort_color("蓝色"), 0x0000FF)

    def test_invalid_returns_none(self):
        self.assertIsNone(bridge._parse_sort_color(None))
        self.assertIsNone(bridge._parse_sort_color(""))
        self.assertIsNone(bridge._parse_sort_color("  "))
        self.assertIsNone(bridge._parse_sort_color("不是颜色"))
        self.assertIsNone(bridge._parse_sort_color("#GGGGGG"))


class AsListTest(unittest.TestCase):
    def test_list_passthrough(self):
        self.assertEqual(bridge._as_list([{"a": 1}]), [{"a": 1}])

    def test_json_array_string(self):
        self.assertEqual(bridge._as_list('[{"column":"B"}]'), [{"column": "B"}])

    def test_single_dict_wrapped(self):
        self.assertEqual(bridge._as_list({"column": "B"}), [{"column": "B"}])

    def test_json_single_object_string(self):
        self.assertEqual(bridge._as_list('{"column":"B"}'), [{"column": "B"}])

    def test_garbage_returns_empty(self):
        self.assertEqual(bridge._as_list("not json"), [])
        self.assertEqual(bridge._as_list(None), [])
        self.assertEqual(bridge._as_list(""), [])


if __name__ == "__main__":
    unittest.main(verbosity=2)
