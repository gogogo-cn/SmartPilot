# -*- coding: utf-8 -*-
"""文字 OCR 之后的公式分流。RapidOCR 和 PaddleOCR 共用这一份。

文字引擎只负责出框。这里决定整图、公式行还是正文不动，再调用 pix2tex，
并把结果收成同一种条目：text 为 LaTeX，is_formula 为 True。
公式模型不可用时原样返回文字结果。
"""

import importlib.util
import json
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading

# 机械图、形位公差、工程图不进 pix2tex。括号、乘号和标题栏会被猜成空公式。
_ROUTING_MODES = ("auto", "formula")
_SKIP_MODES = ("text", "number", "date")

_STRONG_PATTERNS = [
    r"\\frac",
    r"\\sqrt",
    r"\\sum",
    r"\\int",
    r"\\prod",
    r"\\lim",
    r"\\infty",
    r"\\partial",
    r"\\nabla",
    r"\\begin\{",
    r"\\mathrm",
    r"\\mathbf",
    r"\\hat",
    r"\\vec",
    r"\\theta|\\alpha|\\beta|\\gamma|\\delta|\\omega|\\lambda|\\sigma|\\pi",
    r"∫|∑|∏|∂|∞|√|Δ|π|θ|α|β|γ|δ",
]

_MATH_PATTERNS = [
    r"=", r"\+", r"\-", r"[×x*/÷]",
    r"\bsin\b", r"\bcos\b", r"\btan\b", r"\blog\b", r"\bln\b",
    r"\blim\b", r"\bexp\b", r"\bmax\b", r"\bmin\b",
    r"\\alpha|\\beta|\\gamma|\\delta|\\epsilon|\\theta|\\mu|\\pi|\\sigma|\\omega",
    r"[_^]",
    r"∫|∑|∏|∂|∞|√|Δ|π",
    r"\(|\)|\[|\]|\{|\}",
]

_session_lock = threading.Lock()
_ocr = None
_ocr_state = "unknown"  # unknown / ready / subprocess / missing
_worker = None
_worker_lock = None


def _log(log, message):
    if log is not None:
        try:
            log(message)
            return
        except Exception:
            pass
    print(f"[formula] {message}", file=sys.stderr, flush=True)


def formula_model_missing():
    """公式模式的提示用。尚未尝试加载时，只看 pix2tex 是否装上。"""
    if _ocr_state == "missing":
        return True
    if _ocr_state in ("ready", "subprocess"):
        return False
    try:
        return importlib.util.find_spec("pix2tex") is None
    except Exception:
        return True


def route_formulas(image_path, items, recognition_mode="auto", log=None):
    """按同一套规则把文字框换成公式。

    返回:
      items: 原位替换后的条目
      formulas: [{latex, box}]
      markdown_text: 公式行用 $$ 包住；未替换时为 None
      applied: 是否真的换成了公式
    """
    items = list(items or [])
    unchanged = {"items": items, "formulas": [], "markdown_text": None, "applied": False}
    mode = (recognition_mode or "auto").strip().lower()
    if not items or mode in _SKIP_MODES or mode not in _ROUTING_MODES:
        return unchanged

    if mode == "formula" or _is_single_formula_page(items, log):
        latex = _align_simple_lhs(_recognize_file(image_path, log), items, log)
        if _latex_usable(latex):
            box = _union_box(items)
            latex = latex.strip()
            _log(log, f"[整图公式] {latex[:160]}")
            _log(log, "[整图公式] 置信度为合成占位已省略（pix2tex 无原生置信度）")
            formula_item = _formula_item(latex, box, None)
            return {
                "items": [formula_item],
                "formulas": [{"latex": latex, "box": box}],
                "markdown_text": f"$$\n{latex}\n$$",
                "applied": True,
            }
        if mode == "formula":
            return unchanged

    if _looks_like_equation_page(items):
        bands = _split_equation_bands(items)
        if len(bands) >= 2 or (len(bands) == 1 and len(bands[0]) >= 2):
            recognized = _recognize_bands(image_path, items, bands, log)
            if recognized:
                merged = _drop_brace_items(_replace_items(items, recognized))
                _log(log, f"[公式行] 识别到 {len(recognized)} 行")
                return _pack(merged, recognized)

    recognized = _recognize_regions(image_path, items, log)
    if not recognized:
        return unchanged
    merged = _replace_items(items, recognized)
    _log(log, f"[区域公式] 替换 {len(recognized)} 处")
    return _pack(merged, recognized)


def recognize_regions(image_path, items, log=None):
    """论文页：只裁公式行。返回 [{latex, box, indices}]。"""
    return _recognize_regions(image_path, items, log)


def recognize_whole(image_path, log=None):
    """没有文字框时的整图识别。返回 LaTeX 或空字符串。"""
    latex = _recognize_file(image_path, log)
    return latex.strip() if _latex_usable(latex) else ""


def _pack(items, recognized):
    formulas = [{"latex": f["latex"], "box": f["box"]} for f in recognized]
    return {
        "items": items,
        "formulas": formulas,
        "markdown_text": _markdown(items),
        "applied": True,
    }


def _markdown(items):
    parts = []
    for item in items:
        text = item.get("text") or ""
        if item.get("is_formula"):
            parts.append(f"$$\n{text}\n$$")
        else:
            parts.append(text)
    return "\n".join(parts)


def _formula_item(latex, box, confidence=None):
    # pix2tex 不提供原生置信度，缺失时干脆省略该键，
    # 让 VB 端 frm("confidence") IsNot Nothing 守卫安全跳过（不再伪造 0.95）。
    item = {
        "text": latex,
        "box": box,
        "is_formula": True,
    }
    if confidence is not None:
        item["confidence"] = confidence
    return item


def _detect_formula_lines(items):
    """True 公式行，False 正文，None 可并入相邻公式块的短行。"""
    flags = []
    for item in items:
        text = (item.get("text") or "").strip()
        if not text:
            flags.append(False)
            continue
        chinese = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
        if chinese > len(text) * 0.5:
            flags.append(False)
            continue
        if any(re.search(pat, text) for pat in _STRONG_PATTERNS):
            flags.append(True)
            continue
        if re.match(r"^[\d.,%\s]+$", text):
            flags.append(None)
            continue
        medium = 0
        if re.search(r"\b(sin|cos|tan|log|ln|exp|lim|max|min)\b", text):
            medium += 1
        if re.search(r"\b[a-zA-Z]\s*\(", text) and re.search(r"[a-zA-Z]\)", text):
            medium += 1
        if re.search(r"[_^]", text):
            medium += 1
        if re.search(r"\([^()]*\([^)]*\)\)", text):
            medium += 1
        if re.search(r"\{[^{}]*\{[^}]*\}\}", text):
            medium += 1
        if re.search(r"\d+[a-zA-Z]|[a-zA-Z]\d+", text) and not re.search(r"\d{4,}", text):
            medium += 1
        if re.search(r"\d+\.\d*[a-zA-Z(]", text):
            medium += 1
        if re.search(r"(?:^|\s|[=+\-*(])\d+[ \t]*/[ \t]*\d+", text):
            medium += 1
        if re.search(r"[a-zA-Z][ \t]*/[ \t]*\d+", text):
            medium += 1
        if (chinese == 0 and not re.search(r"[A-Za-z]{4,}", text)
                and re.search(r"[√∑∫παβγδμσθλ∞≤≥≠≈⁰¹²³⁴⁵⁶⁷⁸⁹]", text)):
            flags.append(True)
            continue
        hits = sum(1 for pat in _MATH_PATTERNS if re.search(pat, text))
        flags.append(medium >= 2 or hits >= 2)
    return flags


def _merge_formula_blocks(flags):
    blocks = []
    current = []
    for index, flag in enumerate(flags):
        if flag is True:
            current.append(index)
        elif flag is None and current:
            current.append(index)
        elif current:
            blocks.append(current)
            current = []
    if current:
        blocks.append(current)
    return [block for block in blocks if any(flags[i] is True for i in block)]


def _is_single_formula_page(items, log):
    if not items or len(items) > 16:
        return False
    texts = "".join((item.get("text") or "") for item in items)
    if not texts.strip():
        return False
    chinese = sum(1 for ch in texts if "\u4e00" <= ch <= "\u9fff")
    if chinese > max(2, int(len(texts) * 0.15)):
        return False
    centers = []
    heights = []
    for item in items:
        bounds = _bounds(item)
        if not bounds:
            continue
        centers.append((bounds[1] + bounds[3]) / 2.0)
        heights.append(max(bounds[3] - bounds[1], 1))
    if not centers:
        return False
    span = max(centers) - min(centers)
    band = max(heights) if heights else 1
    if span > max(band * 3.2, 48):
        return False
    return _contains_math_symbols(texts, log) or _unicode_math_line(texts)


def _unicode_math_line(text):
    if not text or len(text.strip()) < 3:
        return False
    markers = set("=+-±×÷√∑∫^_/()[]{}⁰¹²³⁴⁵⁶⁷⁸⁹")
    hits = 0
    has_greek = False
    for ch in text:
        if ch in markers:
            hits += 1
        elif "\u0370" <= ch <= "\u03ff":
            hits += 1
            has_greek = True
    return hits >= 2 and ("=" in text or "√" in text or has_greek)


def _contains_math_symbols(text, log):
    if not text or len(text.strip()) < 5:
        return False
    compact = text.replace(" ", "").replace("\n", "")
    chinese = sum(1 for ch in text if "\u4e00" <= ch <= "\u9fff")
    if compact and chinese / len(compact) > 0.3:
        return False
    words = re.findall(r"[A-Za-z]{6,}", text)
    if sum(len(word) for word in words) / max(len(text), 1) > 0.3:
        return False
    if any(re.search(pat, text) for pat in _STRONG_PATTERNS):
        return True
    medium = 0
    if re.search(r"\b(sin|cos|tan|log|ln|exp|lim|max|min)\b", text):
        medium += 1
    if re.search(r"\b[A-Za-z]\s*\(", text) and re.search(r"[A-Za-z]\)", text):
        medium += 1
    if re.search(r"[_^]", text):
        medium += 1
    if re.search(r"\([^()]*\([^)]*\)\)", text):
        medium += 1
    if re.search(r"\{[^{}]*\{[^}]*\}\}", text):
        medium += 1
    if re.search(r"\d+[A-Za-z]|[A-Za-z]\d+", text) and not re.search(r"\d{4,}", text):
        medium += 1
    if re.search(r"\d+\.\d*[A-Za-z(]", text):
        medium += 1
    if re.search(r"(?:^|\s|[=+\-*(])\d+[ \t]*/[ \t]*\d+", text):
        medium += 1
    if re.search(r"[A-Za-z][ \t]*/[ \t]*\d+", text):
        medium += 1
    return medium >= 2


def _looks_like_equation_page(items):
    if not items or len(items) < 3 or len(items) > 40:
        return False
    blob = "".join((item.get("text") or "") for item in items)
    compact = re.sub(r"\s+", "", blob)
    if len(compact) < 6:
        return False
    chinese = sum(1 for ch in compact if "\u4e00" <= ch <= "\u9fff")
    if chinese > max(2, int(len(compact) * 0.12)):
        return False
    if blob.count("=") + blob.count("≠") < 1:
        return False
    if len(re.findall(r"[A-Za-z]{8,}", blob)) >= 2:
        return False
    return "{" in blob and "(" in blob


def _split_equation_bands(items):
    median_h = _median_height(items)
    order = []
    for index, item in enumerate(items):
        if _is_brace_item(item, median_h):
            continue
        bounds = _bounds(item)
        if not bounds:
            continue
        order.append((bounds[1], index, bounds))
    order.sort(key=lambda row: (row[0], row[2][0]))
    if not order:
        return []
    gap = max(8, median_h * 0.16)
    bands = []
    current = []
    current_y1 = None
    for y0, index, bounds in order:
        if current and current_y1 is not None and y0 > current_y1 + gap:
            bands.append(current)
            current = [index]
            current_y1 = bounds[3]
        else:
            current.append(index)
            current_y1 = bounds[3] if current_y1 is None else max(current_y1, bounds[3])
    if current:
        bands.append(current)
    return bands


def _recognize_regions(image_path, items, log):
    if not items:
        return []
    flags = _detect_formula_lines(items)
    blocks = _merge_formula_blocks(flags)
    if not blocks:
        return []
    image = _open_image(image_path, log)
    if image is None:
        return []
    width, height = image.size
    results = []
    for block in blocks:
        xs, ys = [], []
        for index in block:
            bounds = _bounds(items[index])
            if not bounds:
                continue
            xs.extend((bounds[0], bounds[2]))
            ys.extend((bounds[1], bounds[3]))
        if not xs:
            continue
        x0, x1 = min(xs), max(xs)
        y0, y1 = min(ys), max(ys)
        pad_x = max(12, int((x1 - x0) * 0.08))
        pad_y = max(16, int((y1 - y0) * 0.45))
        crop = _crop(image, x0, y0, x1, y1, pad_x, pad_y, width, height)
        if crop is None:
            continue
        latex = _clean_latex(_recognize_image(crop, log))
        if not _latex_usable(latex):
            continue
        _log(log, f"[区域公式] 块{block} -> {latex[:80]}")
        results.append({
            "latex": latex,
            "box": [[x0, y0], [x1, y0], [x1, y1], [x0, y1]],
            "indices": list(block),
        })
    return results


def _recognize_bands(image_path, items, bands, log):
    image = _open_image(image_path, log)
    if image is None:
        return []
    width, height = image.size
    median_h = _median_height(items)
    content_x0, content_x1 = None, None
    for item in items:
        if _is_brace_item(item, median_h):
            continue
        bounds = _bounds(item)
        if not bounds:
            continue
        content_x0 = bounds[0] if content_x0 is None else min(content_x0, bounds[0])
        content_x1 = bounds[2] if content_x1 is None else max(content_x1, bounds[2])
    results = []
    for band in bands:
        xs, ys = [], []
        for index in band:
            bounds = _bounds(items[index])
            if not bounds:
                continue
            xs.extend((bounds[0], bounds[2]))
            ys.extend((bounds[1], bounds[3]))
        if not xs:
            continue
        x0 = content_x0 if content_x0 is not None else min(xs)
        x1 = content_x1 if content_x1 is not None else max(xs)
        y0, y1 = min(ys), max(ys)
        pad_x = max(12, int((x1 - x0) * 0.06))
        pad_y = max(10, int((y1 - y0) * 0.18))
        crop = _crop(image, x0, y0, x1, y1, pad_x, pad_y, width, height, min_size=8)
        if crop is None:
            continue
        latex = _align_simple_lhs(_clean_latex(_recognize_image(crop, log)), [items[i] for i in band], log)
        if not _latex_usable(latex):
            continue
        if not any(tok in latex for tok in ("\\frac", "\\partial", "\\sum", "\\int", "=")):
            continue
        results.append({
            "latex": latex,
            "box": [[x0, y0], [x1, y0], [x1, y1], [x0, y1]],
            "indices": list(band),
        })
    return results


def _replace_items(items, recognized):
    remove = set()
    inserts = {}
    for formula in recognized:
        indexes = formula.get("indices") or []
        if not indexes:
            continue
        remove.update(indexes)
        first = min(indexes)
        # 仅当被合并的原始 OCR 条目本身带有真实置信度时才平均；
        # 否则（如整图公式）传 None，由 _formula_item 省略该键，避免伪造高置信度。
        confs = [items[i].get("confidence") for i in indexes if 0 <= i < len(items) and items[i].get("confidence") is not None]
        confidence = round(sum(confs) / len(confs), 4) if confs else None
        inserts[first] = _formula_item(formula["latex"], formula["box"], confidence)
    merged = []
    for index, item in enumerate(items):
        if index in inserts:
            merged.append(inserts[index])
        if index in remove:
            continue
        merged.append(item)
    return merged


def _drop_brace_items(items):
    median_h = _median_height(items)
    kept = [item for item in items if not _is_brace_item(item, median_h)]
    return kept or items


def _recognize_file(image_path, log):
    image = _open_image(image_path, log)
    if image is None:
        return ""
    return _clean_latex(_recognize_image(image, log))


def _recognize_image(image, log):
    runner = _formula_runner(log)
    if runner is None or image is None:
        return ""
    prepared = _preprocess(image)
    if runner == "subprocess":
        return _recognize_via_worker(prepared, log) or ""
    try:
        latex = runner(prepared)
        if isinstance(latex, (list, tuple)):
            latex = latex[0] if latex else ""
        return str(latex or "")
    except Exception as exc:
        _log(log, f"公式识别失败，改走子进程: {exc}")
        return _recognize_via_worker(prepared, log) or ""


def _formula_runner(log):
    global _ocr, _ocr_state
    with _session_lock:
        if _ocr_state == "missing":
            return None
        if _ocr_state == "subprocess":
            return "subprocess"
        if _ocr is not None:
            return _ocr
        try:
            os.environ["SHM_DISABLE"] = "1"
            from pix2tex.cli import LatexOCR
            model = LatexOCR()
            try:
                model.args.temperature = 1e-6
            except Exception:
                pass
            _ocr = model
            _ocr_state = "ready"
            _log(log, "LaTeX-OCR 已加载")
            return _ocr
        except Exception as exc:
            _log(log, f"进程内公式模型加载失败，改用子进程: {exc}")
            _ocr_state = "subprocess"
            return "subprocess"


def _recognize_via_worker(image, log):
    path = None
    try:
        fd, path = tempfile.mkstemp(suffix=".png")
        os.close(fd)
        image.save(path)
        return _worker_latex(path, log)
    finally:
        if path and os.path.exists(path):
            try:
                os.remove(path)
            except Exception:
                pass


def _worker_script():
    return r'''
import os, sys, json, warnings
os.environ["SHM_DISABLE"] = "1"
warnings.filterwarnings("ignore")
from pix2tex.cli import LatexOCR
from PIL import Image

def preprocess_image(img):
    img = img.convert("RGB")
    w, h = img.size
    if h < 192:
        scale = 192 / float(max(h, 1))
        w = max(int(w * scale), 1)
        h = 192
        img = img.resize((w, h), Image.LANCZOS)
    long_side = max(w, h)
    if long_side > 1400:
        scale = 1400 / float(long_side)
        w = max(int(w * scale), 1)
        h = max(int(h * scale), 1)
        img = img.resize((w, h), Image.LANCZOS)
    margin = max(16, int(min(w, h) * 0.08))
    canvas = Image.new("RGB", (w + margin * 2, h + margin * 2), (255, 255, 255))
    canvas.paste(img, (margin, margin))
    return canvas

ocr = LatexOCR()
ocr.args.temperature = 1e-6
sys.stdout.write(json.dumps({"ready": True}) + "\n")
sys.stdout.flush()
for raw in sys.stdin:
    raw = raw.strip()
    if not raw:
        continue
    try:
        req = json.loads(raw)
    except Exception:
        sys.stdout.write(json.dumps({"ok": False, "error": "bad json"}) + "\n")
        sys.stdout.flush()
        continue
    if req.get("cmd") == "exit":
        break
    try:
        img = preprocess_image(Image.open(req["path"]).convert("RGB"))
        latex = (ocr(img) or "").replace("\n", " ").strip()
        sys.stdout.write(json.dumps({"ok": True, "latex": latex}, ensure_ascii=False) + "\n")
    except Exception as exc:
        sys.stdout.write(json.dumps({"ok": False, "error": str(exc)}) + "\n")
    sys.stdout.flush()
'''


def _stop_worker():
    global _worker
    proc = _worker
    _worker = None
    if proc is None:
        return
    try:
        if proc.poll() is None and proc.stdin:
            proc.stdin.write('{"cmd":"exit"}\n')
            proc.stdin.flush()
    except Exception:
        pass
    try:
        proc.kill()
    except Exception:
        pass


def _readline_timeout(stream, timeout):
    mailbox = queue.Queue()

    def read():
        try:
            mailbox.put(stream.readline())
        except Exception as exc:
            mailbox.put(exc)

    threading.Thread(target=read, daemon=True).start()
    try:
        item = mailbox.get(timeout=timeout)
    except queue.Empty:
        return None
    if isinstance(item, Exception):
        return None
    return item


def _ensure_worker(log):
    global _worker
    proc = _worker
    if proc is not None and proc.poll() is None:
        return proc
    _stop_worker()
    proc = subprocess.Popen(
        [sys.executable, "-c", _worker_script()],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
        encoding="utf-8",
        bufsize=1,
        env={**os.environ, "SHM_DISABLE": "1"},
    )
    _worker = proc
    ready_line = _readline_timeout(proc.stdout, 90)
    if not ready_line:
        _log(log, "公式进程启动超时")
        _stop_worker()
        return None
    try:
        ready = json.loads(ready_line)
    except Exception:
        _log(log, "公式进程就绪信号无效")
        _stop_worker()
        return None
    if not ready.get("ready"):
        _stop_worker()
        return None
    _log(log, "公式进程已就绪")
    return proc


def _worker_latex(path, log):
    global _worker_lock, _ocr_state
    if _worker_lock is None:
        _worker_lock = threading.Lock()
    with _worker_lock:
        proc = _ensure_worker(log)
        if proc is None:
            _ocr_state = "missing"
            return ""
        try:
            proc.stdin.write(json.dumps({"path": path}) + "\n")
            proc.stdin.flush()
        except Exception as exc:
            _log(log, f"公式进程写入失败: {exc}")
            _stop_worker()
            return ""
        line = _readline_timeout(proc.stdout, 60)
        if not line:
            _log(log, "公式进程响应超时")
            _stop_worker()
            return ""
        try:
            msg = json.loads(line)
        except Exception:
            _log(log, "公式进程返回无法解析")
            _stop_worker()
            return ""
        if not msg.get("ok"):
            _log(log, f"公式识别失败: {msg.get('error')}")
            return ""
        return msg.get("latex") or ""


def _open_image(image_path, log):
    try:
        from PIL import Image
        image = Image.open(image_path)
        image.load()
        return image.convert("RGB")
    except Exception as exc:
        _log(log, f"公式图无法打开: {exc}")
        return None


def _crop(image, x0, y0, x1, y1, pad_x, pad_y, width, height, min_size=4):
    left = max(0, int(x0) - pad_x)
    upper = max(0, int(y0) - pad_y)
    right = min(width, int(x1) + pad_x + 1)
    lower = min(height, int(y1) + pad_y + 1)
    if right - left < min_size or lower - upper < min_size:
        return None
    return image.crop((left, upper, right, lower))


def _preprocess(img):
    from PIL import Image

    img = img.convert("RGB")
    w, h = img.size
    if h < 192:
        scale = 192.0 / float(max(h, 1))
        w = max(int(w * scale), 1)
        h = 192
        img = img.resize((w, h), Image.LANCZOS)
    long_side = max(w, h)
    if long_side > 1400:
        scale = 1400.0 / float(long_side)
        w = max(int(w * scale), 1)
        h = max(int(h * scale), 1)
        img = img.resize((w, h), Image.LANCZOS)
    margin = max(16, int(min(w, h) * 0.08))
    canvas = Image.new("RGB", (w + margin * 2, h + margin * 2), (255, 255, 255))
    canvas.paste(img, (margin, margin))
    return canvas


def _bounds(item):
    xs, ys = [], []
    for point in item.get("box") or []:
        if isinstance(point, (list, tuple)) and len(point) >= 2:
            xs.append(float(point[0]))
            ys.append(float(point[1]))
    if not xs:
        return None
    return min(xs), min(ys), max(xs), max(ys)


def _union_box(items):
    xs, ys = [], []
    for item in items:
        bounds = _bounds(item)
        if not bounds:
            continue
        xs.extend((bounds[0], bounds[2]))
        ys.extend((bounds[1], bounds[3]))
    if not xs:
        return [[0, 0], [1, 0], [1, 1], [0, 1]]
    return [[min(xs), min(ys)], [max(xs), min(ys)], [max(xs), max(ys)], [min(xs), max(ys)]]


def _median_height(items):
    heights = []
    for item in items:
        bounds = _bounds(item)
        if bounds:
            heights.append(max(bounds[3] - bounds[1], 1))
    if not heights:
        return 1
    heights.sort()
    return heights[len(heights) // 2]


def _is_brace_item(item, median_h):
    bounds = _bounds(item)
    if bounds:
        width = max(bounds[2] - bounds[0], 1)
        height = max(bounds[3] - bounds[1], 1)
        if height > max(2.2 * median_h, median_h + 8) and width < height * 0.45:
            return True
    text = (item.get("text") or "").strip()
    return text in set("{ } ( ) [ ] ⎧ ⎨ ⎩ ⎫ ⎬ ⎭ ⎪ ⎮")


def _clean_latex(latex):
    text = (latex or "").replace("\n", " ").strip()
    if "\x0c" in text:
        text = text.replace("\x0c", "\\f")
    deficit = text.count("{") - text.count("}")
    if 0 < deficit <= 4:
        text += "}" * deficit
    return text


def _latex_usable(latex):
    text = (latex or "").strip()
    if len(text) < 6 or text.count("{") != text.count("}"):
        return False
    return "\\" in text or "^" in text or "_" in text


def _align_simple_lhs(latex, items, log):
    latex = (latex or "").strip()
    if not latex:
        return latex
    head_re = re.compile(
        r"^(?:\\int|\\(?:mathcal|mathrm|mathbf)\s*\{\s*([A-Za-z])\s*\}|([A-Za-z]))"
        r"\s*(?:\\left\s*)?\(\s*"
        r"(?:\\(?:mathcal|mathrm|mathbf)\s*\{\s*([A-Za-z])\s*\}|([A-Za-z]))"
        r"\s*(?:\\right\s*)?\)"
    )
    match = head_re.match(latex)
    if not match:
        return _fix_equals_read_as_arrow(latex, log)
    rest = latex[match.end():]
    is_int = latex.startswith("\\int")
    fn = match.group(1) or match.group(2) or ""
    arg = match.group(3) or match.group(4) or ""
    confused = (is_int or fn in ("J", "j")) and arg in ("Q", "q")
    body_has_x = "x" in rest and "J" not in rest and "Q" not in rest
    ocr_lhs = _lhs_from_text_ocr(items)
    if confused and body_has_x:
        if ocr_lhs and ocr_lhs[1] == "x" and ocr_lhs[0] not in ("j", "q"):
            fixed = f"{ocr_lhs[0]}({ocr_lhs[1]})" + rest
        else:
            fixed = "f(x)" + rest
        return _fix_equals_read_as_arrow(fixed, log)
    if ocr_lhs and not is_int and ocr_lhs != (fn, arg):
        return _fix_equals_read_as_arrow(f"{ocr_lhs[0]}({ocr_lhs[1]})" + rest, log)
    return _fix_equals_read_as_arrow(latex, log)


def _fix_equals_read_as_arrow(latex, log):
    raw = latex or ""
    text = raw.strip()
    text = re.sub(r"^(?:\$\$|\$|\\\[)\s*", "", text)
    text = re.sub(r"\s*(?:\$\$|\$|\\\])$", "", text)
    text = re.sub(r"^\\(?:displaystyle|textstyle|scriptstyle)\s*", "", text)
    spacing = r"(?:\\[,;:!]|\\ |\\quad|\\qquad|\\thinspace|\\:|\\>|\\hspace\s*\{[^{}]*\}|\s)*"
    relation = (
        r"(?:\\longrightarrow|\\rightarrow|\\Longrightarrow|\\Rightarrow"
        r"|\\longmapsto|\\mapsto|\\xrightarrow(?:\s*\[[^\]]*\])?(?:\s*\{[^{}]*\})?"
        r"|\\rightharpoonup|\\rightharpoondown|\\leadsto|\\rightsquigarrow"
        r"|\\relbar(?:\s*\\joinrel\s*\\(?:rightarrow|relbar))+"
        r"|\\to|\\bigcup|\\cup|\\bigcap|\\cap|\\sqcup|\\uplus"
        r"|→|⟶|⇒|⟹|↦|⟼|∪|⋃|∩|=)"
    )
    pattern = re.compile(
        r"^(?P<head>"
        r"(?P<fn>\\int|\\(?:mathcal|mathrm|mathbf|operatorname)\s*\{\s*(?P<styled>[A-Za-z])\s*\}|(?P<plain>[A-Za-z]))"
        r"\s*(?:\\left\s*)?\(\s*"
        r"(?P<arg>\\[a-zA-Z]+|[\u0370-\u03ff]|\\(?:mathcal|mathrm|mathbf)\s*\{\s*[A-Za-z]\s*\}|[A-Za-z])"
        r"\s*(?:\\right\s*)?\)"
        r")"
        + spacing + r"(?P<rel>" + relation + r")" + spacing
        + r"(?=\\(?:d|t)?frac\b|e\s*\^|\\exp\b)"
    )
    found = pattern.match(text)
    if found:
        letter = found.group("styled") or found.group("plain") or ""
        arg = found.group("arg") or ""
        body = text[found.end():]
        bogus = (
            found.group("fn") == "\\int"
            or arg.startswith("\\")
            or (len(arg) == 1 and "\u0370" <= arg <= "\u03ff")
            or (letter in ("J", "j") and arg in ("Q", "q"))
        )
        if letter and letter not in ("J", "j") and arg == "x" and found.group("fn") != "\\int":
            bogus = False
        if bogus and not re.search(r"\(\s*x\s*-", body):
            bogus = False
        head = found.group("head")
        if bogus:
            head = "f(x)"
        if bogus or found.group("rel") != "=":
            text = head + " = " + body
    return _fix_exp_base(text if text != raw else raw)


def _fix_exp_base(latex):
    return re.sub(
        r"\\(?:sigma|epsilon|varepsilon)(?=\s*\^\s*\{?\s*-\s*(?:\\frac\s*\{)?\s*\(\s*x\s*-)",
        "e",
        latex or "",
        count=1,
    )


def _lhs_from_text_ocr(items):
    ordered = sorted(items or [], key=_box_left)
    blob = ""
    for item in ordered[:3]:
        piece = re.sub(r"\s+", "", (item.get("text") or ""))
        if not piece:
            continue
        blob += piece
        found = re.match(r"^([a-z])[\(（]([a-z])[\)）]", blob)
        if found:
            return found.group(1), found.group(2)
        if len(blob) > 12:
            break
    return None


def _box_left(item):
    bounds = _bounds(item)
    return bounds[0] if bounds else 0
