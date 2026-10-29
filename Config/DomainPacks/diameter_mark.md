---
id: diameter_mark
name: 直径符号改写
description: 识别改写入口示例。默认不改字；按脚本注释打开规则后，把单独的 O 改成 Ø
icon: 径
category: 工程图
version: 1.0.0
author: SmartPilot
pipeline: python
module: diameter_mark.py
---

这段脚本在文字引擎出框之后运行。它拿到原图路径和每个文字框，只把要改的框写回去。
图号、图名这类字段抽取不要写在这里，用解析脚本。

```python
# 识别改写入口
# 输入（标准输入的一个 JSON）：
#   imagePath  原图绝对路径。要用像素时自己打开这张图
#   packId     模块 id
#   options    清单开关的当前值
#   elements   [{index, text, confidence, box:{left, top, right, bottom}}]
# 输出（标准输出的一个 JSON，放在最后一行）：
#   {"success": true, "elements": [{"index": 0, "text": "改后的文字"}]}
# 没返回的框保持原样。success 为 false、超时或脚本报错时，整份识别结果保持原样。
import json
import sys

def rewrite(image_path, elements, options):
    changed = []
    for el in elements:
        text = el.get("text") or ""
        new_text = text
        # 示例规则：单独一个 O / o / 0 当成直径符号。默认关闭。
        # 要看效果，把 _rule_enabled 改成 return True。
        if _rule_enabled() and text in ("O", "o", "0"):
            new_text = "Ø"
        if new_text != text:
            changed.append({"index": el["index"], "text": new_text})
    return changed

def _rule_enabled():
    return False

def main():
    raw = sys.stdin.read()
    data = json.loads(raw) if raw.strip() else {}
    changed = rewrite(
        data.get("imagePath") or "",
        data.get("elements") or [],
        data.get("options") or {},
    )
    json.dump({"success": True, "elements": changed}, sys.stdout, ensure_ascii=False)
    sys.stdout.write("\n")

if __name__ == "__main__":
    main()
```
