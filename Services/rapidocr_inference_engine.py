#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
RapidOCR 统一推理引擎
====================
【架构说明】
- 长驻进程模式：启动时一次性加载所有模型，通过 stdin/stdout JSON 协议通信
- 基于 ONNX Runtime 的轻量级 OCR，模型体积小（约 30MB），启动快
- 支持通用文字识别、多语言支持
- 支持批量推理

【协议】
输入: 每行一个 JSON 对象
输出: 每行一个 JSON 对象（含 "success" 字段）

【通信命令】
{
  "cmd": "ocr",              // 纯文本 OCR
  "image_path": "..."        // 图片路径
}

{
  "cmd": "ocr_pos",          // 带坐标的 OCR
  "image_path": "..."
}

{
  "cmd": "document",         // 结构化文档识别
  "image_path": "...",
  "lang": "ch"              // 语言
}

{
  "cmd": "batch_ocr",        // 批量 OCR
  "image_paths": ["...", "..."],
  "mode": "text"            // "text" 或 "pos"
}

{
  "cmd": "status"            // 获取状态
}

{
  "cmd": "shutdown"          // 关闭引擎
}

{
  "cmd": "set_config",       // 动态配置
  "config": {
    "lang": "ch",
    "cpu_threads": 4
  }
}
"""

import os
import sys
import json
import time
import traceback

try:
    from ocr_overlap_dedupe import dedupe_overlapping_items
except Exception:
    def dedupe_overlapping_items(items):
        return items

try:
    from ocr_formula_router import route_formulas, formula_model_missing
except Exception:
    def route_formulas(image_path, items, recognition_mode="auto", log=None):
        return {"items": list(items or []), "formulas": [], "markdown_text": None, "applied": False}

    def formula_model_missing():
        return True

# ============================================================
# CRITICAL: 设置环境变量 - 必须在导入任何依赖之前执行
# ============================================================
os.environ["SHM_DISABLE"] = "1"

def _log(msg):
    """输出日志到 stderr"""
    print(f"[RapidOCR] {msg}", file=sys.stderr, flush=True)

def _ensure_rapidocr():
    """确保 RapidOCR 可用，不可用时尝试导入或安装"""
    try:
        from rapidocr_onnxruntime import RapidOCR
        return RapidOCR
    except ImportError:
        _log("RapidOCR 未安装，尝试安装...")
        try:
            import subprocess
            proc = subprocess.run(
                [sys.executable, "-m", "pip", "install", "rapidocr_onnxruntime", "-q"],
                capture_output=True, text=True, timeout=300
            )
            if proc.returncode != 0:
                _log(f"RapidOCR 安装失败(exit={proc.returncode}): {(proc.stderr or '')[:200]}")
        except Exception as e:
            _log(f"RapidOCR 安装异常: {e}")
        try:
            from rapidocr_onnxruntime import RapidOCR
            return RapidOCR
        except Exception:
            return None

# 全局 OCR 引擎实例
_ocr_engine = None
_config = {
    "lang": "ch",
    "cpu_threads": 4
}

def _get_ocr_engine():
    """获取或创建 OCR 引擎单例"""
    global _ocr_engine
    if _ocr_engine is None:
        RapidOCR_cls = _ensure_rapidocr()
        if RapidOCR_cls is None:
            raise RuntimeError("RapidOCR 无法加载，请运行: pip install rapidocr_onnxruntime")
        # 新版 rapidocr_onnxruntime 统一为无参构造（自带默认模型路径）；
        # 旧版 kwargs 签名（det_use_cuda 等）在新版上会抛 KeyError('model_path')。
        # 先试新版无参构造，失败再回退旧版 kwargs，兼容两个大版本。
        try:
            _ocr_engine = RapidOCR_cls()
            _log("RapidOCR 引擎初始化完成（新版无参构造）")
        except (TypeError, KeyError) as e:
            _log(f"新版无参构造不可用({e})，回退旧版 kwargs 签名")
            _ocr_engine = RapidOCR_cls(
                det_use_cuda=False,
                rec_use_cuda=False,
                cls_use_cuda=False,
                cpu_threads=_config.get("cpu_threads", 4)
            )
            _log(f"RapidOCR 引擎初始化完成 (legacy kwargs, threads={_config.get('cpu_threads')})")
    return _ocr_engine

def _parse_lang(lang):
    """解析语言参数"""
    lang_map = {
        "ch": "ch",      # 中文
        "en": "en",      # 英文
        "ja": "ja",      # 日语
        "ko": "ko",      # 韩语
        "fr": "fr",      # 法语
        "de": "de",      # 德语
        "es": "es",      # 西班牙语
        "ru": "ru",      # 俄语
        "ar": "ar",      # 阿拉伯语
    }
    return lang_map.get(lang, "ch")

def _to_jsonable(obj):
    """把 numpy 标量 / ndarray 转成可 json.dumps 的值"""
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, dict):
        return {str(k): _to_jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_to_jsonable(v) for v in obj]
    if hasattr(obj, "tolist"):
        return _to_jsonable(obj.tolist())
    if hasattr(obj, "item"):
        try:
            return _to_jsonable(obj.item())
        except Exception:
            return str(obj)
    return str(obj)


def _box_to_list(box):
    if hasattr(box, "tolist"):
        box = box.tolist()
    points = []
    if isinstance(box, (list, tuple)):
        for p in box:
            if hasattr(p, "tolist"):
                p = p.tolist()
            if isinstance(p, (list, tuple)) and len(p) >= 2:
                try:
                    points.append([float(p[0]), float(p[1])])
                except (TypeError, ValueError):
                    continue
    return points


def _elapse_seconds(elapse):
    """新版 RapidOCR 的 elapse 可能是 float、list 或 dict"""
    if elapse is None or isinstance(elapse, bool):
        return 0.0
    if isinstance(elapse, (int, float)):
        return float(elapse)
    if isinstance(elapse, dict):
        return sum(_elapse_seconds(v) for v in elapse.values())
    if isinstance(elapse, (list, tuple)):
        return sum(_elapse_seconds(v) for v in elapse)
    return 0.0


def _invoke_engine(engine, image_path):
    """兼容返回 (result, elapse) 或只返回 result 的版本"""
    out = engine(image_path)
    if isinstance(out, tuple):
        if len(out) >= 2:
            return out[0], out[1]
        if len(out) == 1:
            return out[0], None
        return None, None
    return out, None


def _parse_result_line(line):
    """兼容 RapidOCR 新旧版本返回格式
    旧版(v1.3-): [box, (text, score)]
    新版(v1.4+): [box, text, score]
    """
    if not line or len(line) < 2:
        return [], "", 0.0
    box = _box_to_list(line[0])
    second = line[1]
    if isinstance(second, (tuple, list)) and len(second) >= 2 and isinstance(second[0], str):
        text = second[0]
        confidence = second[1]
    elif len(line) >= 3:
        text = second
        confidence = line[2]
    else:
        text = second
        confidence = 0.0
    try:
        confidence = float(confidence)
    except (TypeError, ValueError):
        confidence = 0.0
    if text is None:
        text = ""
    elif not isinstance(text, str):
        text = str(text)
    return box, text, confidence


def _mode_warning(args):
    mode = args.get("recognition_mode") or "auto"
    if mode == "formula" and formula_model_missing():
        return "公式模型未安装，已按文本识别"
    return None


def _apply_formula_route(image_path, items, mode):
    """文字出框后交给共用公式分流。模型不可用时保持原文字框。"""
    routed = route_formulas(image_path, items, mode, log=_log)
    if isinstance(routed, dict) and "items" in routed:
        return routed["items"]
    return items



def _collect_lines(result):
    items = []
    if not result:
        return items
    for line in result:
        box, text, confidence = _parse_result_line(line)
        items.append({"text": text, "confidence": confidence, "box": box})
    before = [it.get("text") for it in items]
    items = dedupe_overlapping_items(items)
    after = [it.get("text") for it in items]
    if before != after:
        _log(f"[重叠字] 去掉检测框交界处的重复字: {before} -> {after}")
    return items


def _rerecognize_vertical(image_path, items):
    """竖排（旋转 90°）文字重识别，并定出阅读方向。与 Paddle 引擎的同名逻辑对齐。

    为什么需要：检测器给的是轴对齐矩形。文字竖排时整个框又高又窄，识别器按横排
    喂进去——实测机械图侧边的竖排 Ø78.5 直接读空、竖排 F8 读成 "B8"（0.43）。
    接近正方形的短行（F8/H7/6H）连现有「高>=宽*1.5 才转一下」的判断都过不了。

    做法：对又高又窄的框分别顺/逆时针转 90° 各识别一次（use_det/use_cls 都关掉，
    两个方向结果不同才开始比出阅读方向），置信度高的赢。
    同时给 item 打上 vdir：up=原文从下往上读，down=从上往下读——
    下游合并被切成两半的竖排标注时按这个顺序接文本。
    """
    if not items:
        return items
    try:
        import numpy as np
        from PIL import Image
    except ImportError:
        return items
    try:
        pil = Image.open(image_path).convert("RGB")
    except Exception as e:
        _log(f"[竖排重识别] 打不开图，跳过: {e}")
        return items
    W, H = pil.size
    engine = _get_ocr_engine()
    fixed = 0
    for item in items:
        box = item.get("box")
        if not box or len(box) < 4:
            continue
        try:
            xs = [float(p[0]) for p in box]
            ys = [float(p[1]) for p in box]
        except (TypeError, ValueError, IndexError):
            continue
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        # 只处理明显竖排的框；矮框（含普通单行文字）不碰，避免无谓的两次识别
        if h < w * 1.15 or h < 8:
            continue
        pad = 3
        l = max(0, int(min(xs)) - pad)
        t = max(0, int(min(ys)) - pad)
        r = min(W, int(round(max(xs))) + pad)
        b = min(H, int(round(max(ys))) + pad)
        if r - l < 3 or b - t < 3:
            continue
        try:
            crop = pil.crop((l, t, r, b))
        except Exception as e:
            _log(f"[竖排重识别] 裁剪失败: {e}")
            continue
        orig_text = item.get("text", "") or ""
        orig_conf = float(item.get("confidence") or 0.0)
        rot_best = ("", 0.0, "")   # (text, conf, vdir)
        # PIL rotate 是逆时针为正：270=顺时针 90°，90=逆时针 90°。
        # 顺时针能读通 → 原文是逆时针放的 → 从下往上读（up）。
        for angle, vdir in ((270, "up"), (90, "down")):
            try:
                rotated = crop.rotate(angle, expand=True)
                out = engine(np.asarray(rotated), use_det=False, use_cls=False, use_rec=True)
                res = out[0] if isinstance(out, tuple) else out
            except Exception as e:
                _log(f"[竖排重识别] 识别失败({angle}): {e}")
                continue
            text, conf = "", 0.0
            if isinstance(res, (list, tuple)) and res:
                first = res[0]
                if isinstance(first, (list, tuple)) and len(first) >= 2 and isinstance(first[0], str):
                    text, conf = first[0], first[1]
                elif isinstance(first, str) and len(res) >= 3:
                    text, conf = first, res[2]
            try:
                conf = float(conf)
            except (TypeError, ValueError):
                conf = 0.0
            if text and conf > rot_best[1]:
                rot_best = (text, conf, vdir)
        changed = False
        # 文本：旋转后明显更通才替换（原读法也不算差的就留住，避免无谓抖动）
        if rot_best[0] and rot_best[1] > orig_conf and rot_best[0] != orig_text:
            item["text"] = rot_best[0]
            item["confidence"] = round(float(rot_best[1]), 4)
            changed = True
        # 方向：两个旋转里赢的那个就是阅读方向，和文本换不换无关
        if rot_best[2] and rot_best[1] >= 0.3 and not item.get("vdir"):
            item["vdir"] = rot_best[2]
            changed = True
        if changed:
            fixed += 1
    if fixed:
        _log(f"[竖排重识别] 修正 {fixed} 个竖排文本框")
    return items


def cmd_ocr(args):
    """纯文本 OCR"""
    image_path = args.get("image_path", "")
    if not image_path or not os.path.exists(image_path):
        return {"success": False, "error": f"图片不存在: {image_path}"}
    
    try:
        engine = _get_ocr_engine()
        result, elapse = _invoke_engine(engine, image_path)
        items = _collect_lines(result)
        items = _rerecognize_vertical(image_path, items)
        items = _apply_formula_route(image_path, items, args.get("recognition_mode") or "auto")
        text = "\n".join(item["text"] for item in items)
        payload = {
            "success": True,
            "text": text,
            "raw_data": items,
            "elapse": _elapse_seconds(elapse),
            "recognition_mode": args.get("recognition_mode") or "auto",
            "lang": _config.get("lang", "ch")
        }
        warning = _mode_warning(args)
        if warning:
            payload["warning"] = warning
        return payload
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}

def cmd_ocr_pos(args):
    """带坐标的 OCR"""
    image_path = args.get("image_path", "")
    if not image_path or not os.path.exists(image_path):
        return {"success": False, "error": f"图片不存在: {image_path}"}
    
    try:
        engine = _get_ocr_engine()
        result, elapse = _invoke_engine(engine, image_path)
        items = _collect_lines(result)
        items = _rerecognize_vertical(image_path, items)
        items = _apply_formula_route(image_path, items, args.get("recognition_mode") or "auto")
        payload = {
            "success": True,
            "items": items,
            "raw_data": items,
            "elapse": _elapse_seconds(elapse),
            "recognition_mode": args.get("recognition_mode") or "auto",
            "lang": _config.get("lang", "ch")
        }
        warning = _mode_warning(args)
        if warning:
            payload["warning"] = warning
        return payload
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}

def cmd_document(args):
    """文档识别（结构化输出）"""
    image_path = args.get("image_path", "")
    lang = args.get("lang", "ch")
    
    if not image_path or not os.path.exists(image_path):
        return {"success": False, "error": f"图片不存在: {image_path}"}
    
    try:
        engine = _get_ocr_engine()
        result, elapse = _invoke_engine(engine, image_path)
        items = _collect_lines(result)
        items = _apply_formula_route(image_path, items, args.get("recognition_mode") or "auto")
        elements = [{
            "type": "text",
            "text": item["text"],
            "confidence": item["confidence"],
            "box": item["box"]
        } for item in items]
        return {
            "success": True,
            "text": "\n".join(item["text"] for item in items),
            "elements": elements,
            "raw_data": items,
            "elapse": _elapse_seconds(elapse)
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}

def _ensure_pymupdf():
    """确保 PyMuPDF 可用，不可用时尝试安装"""
    try:
        import fitz
        return fitz
    except ImportError:
        _log("PyMuPDF 未安装，尝试安装...")
        try:
            import subprocess
            proc = subprocess.run(
                [sys.executable, "-m", "pip", "install", "pymupdf", "-q"],
                capture_output=True, text=True, timeout=300
            )
            if proc.returncode != 0:
                _log(f"PyMuPDF 安装失败(exit={proc.returncode}): {(proc.stderr or '')[:200]}")
        except Exception as e:
            _log(f"PyMuPDF 安装异常: {e}")
        try:
            import fitz
            return fitz
        except Exception:
            return None


def _pdf_to_images(pdf_path, dpi=300, start_page=None, end_page=None, output_dir=None):
    """把 PDF 每页转成高分辨率 PNG 图片，返回 (图片路径列表, 错误)

    dpi: 输出分辨率，默认 300（图纸推荐）
    start_page / end_page: 可选的页码范围（从 1 开始，包含两端）
    output_dir: 可选输出目录，默认在 PDF 同目录建 _ocr_pages 子目录
    """
    fitz = _ensure_pymupdf()
    if fitz is None:
        return [], "PyMuPDF 无法加载，请运行: pip install pymupdf"

    if not os.path.exists(pdf_path):
        return [], f"PDF 不存在: {pdf_path}"

    # 确定输出目录
    if output_dir is None:
        out_dir = os.path.join(os.path.dirname(pdf_path), "_ocr_pages")
    else:
        out_dir = output_dir
    os.makedirs(out_dir, exist_ok=True)

    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        return [], f"PDF 打开失败: {e}"

    total = len(doc)
    # 页码范围
    p0 = 0
    p1 = total - 1
    if start_page is not None:
        p0 = max(0, min(total - 1, start_page - 1))
    if end_page is not None:
        p1 = max(p0, min(total - 1, end_page - 1))

    mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
    pdf_stem = os.path.splitext(os.path.basename(pdf_path))[0]
    images = []

    for i in range(p0, p1 + 1):
        page = doc.load_page(i)
        try:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            fname = f"{pdf_stem}_page_{i + 1:04d}.png"
            fpath = os.path.join(out_dir, fname)
            pix.save(fpath, "png")
            images.append({
                "page_index": i,
                "page_number": i + 1,
                "path": fpath,
                "width": pix.width,
                "height": pix.height,
                "dpi": dpi,
            })
        except Exception as e:
            _log(f"PDF 第 {i + 1} 页转图失败: {e}")

    doc.close()
    return images, None


def cmd_pdf_to_images(args):
    """PDF 转图，返回所有生成的图片路径和尺寸

    请求:
    {"cmd": "pdf_to_images", "pdf_path": "...", "dpi": 300, "start_page": 1, "end_page": 3, "output_dir": "..."}
    """
    pdf_path = args.get("pdf_path", "")
    if not pdf_path:
        return {"success": False, "error": "缺少 pdf_path 参数"}

    dpi = int(args.get("dpi", 300))
    start_page = args.get("start_page")
    end_page = args.get("end_page")
    output_dir = args.get("output_dir")

    try:
        pages, err = _pdf_to_images(
            pdf_path,
            dpi=dpi,
            start_page=int(start_page) if start_page is not None else None,
            end_page=int(end_page) if end_page is not None else None,
            output_dir=output_dir,
        )
        if err:
            return {"success": False, "error": err}
        return {
            "success": True,
            "pages": pages,
            "page_count": len(pages),
            "output_dir": os.path.dirname(pages[0]["path"]) if pages else "",
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


def cmd_ocr_pdf(args):
    """一步到位的 PDF OCR：PDF → 逐页转图 → 逐页识别 → 合并结果

    请求:
    {"cmd": "ocr_pdf", "pdf_path": "...", "dpi": 300, "mode": "pos",
     "recognition_mode": "mechanical", "lang": "ch"}

    mode: "pos"(带坐标，默认) / "text"(纯文本) / "document"(结构化)
    返回: 每页一条 items/text，附带 page_index 和合并结果
    """
    pdf_path = args.get("pdf_path", "")
    if not pdf_path:
        return {"success": False, "error": "缺少 pdf_path 参数"}

    mode = args.get("mode", "pos") or "pos"
    dpi = int(args.get("dpi", 300))
    lang = _parse_lang(args.get("lang") or _config.get("lang", "ch"))
    rec_mode = args.get("recognition_mode") or "auto"

    try:
        pages, err = _pdf_to_images(pdf_path, dpi=dpi)
        if err:
            return {"success": False, "error": err}

        engine = _get_ocr_engine()
        total_elapse = 0.0
        per_page = []

        for pg in pages:
            items = []
            elapse = 0.0
            try:
                result, elapse = _invoke_engine(engine, pg["path"])
                items = _collect_lines(result)
            except Exception as e:
                _log(f"第 {pg['page_number']} 页识别失败: {e}")
                items = []
            total_elapse += _elapse_seconds(elapse)

            if mode == "text":
                per_page.append({
                    "page_index": pg["page_index"],
                    "page_number": pg["page_number"],
                    "text": "\n".join(it["text"] for it in items),
                    "elapse": _elapse_seconds(elapse),
                })
            elif mode == "document":
                per_page.append({
                    "page_index": pg["page_index"],
                    "page_number": pg["page_number"],
                    "elements": [{
                        "type": "text",
                        "text": it["text"],
                        "confidence": it["confidence"],
                        "box": it["box"],
                    } for it in items],
                    "items": items,
                    "elapse": _elapse_seconds(elapse),
                })
            else:  # pos
                per_page.append({
                    "page_index": pg["page_index"],
                    "page_number": pg["page_number"],
                    "items": items,
                    "elapse": _elapse_seconds(elapse),
                })

        payload = {
            "success": True,
            "pdf_path": pdf_path,
            "dpi": dpi,
            "lang": lang,
            "recognition_mode": rec_mode,
            "mode": mode,
            "page_count": len(pages),
            "pages": per_page,
            "total_elapse": total_elapse,
        }
        if mode == "pos":
            # 顶层 items 合并所有页
            merged = []
            for p in per_page:
                for it in p["items"]:
                    item_copy = dict(it)
                    item_copy["page_index"] = p["page_index"]
                    item_copy["page_number"] = p["page_number"]
                    merged.append(item_copy)
            payload["items"] = merged
        if mode == "document":
            all_text = "\n\n".join(p.get("text", "") for p in per_page)
            payload["text"] = all_text
            all_elements = []
            for p in per_page:
                for el in p.get("elements", []):
                    el_copy = dict(el)
                    el_copy["page_index"] = p["page_index"]
                    all_elements.append(el_copy)
            payload["elements"] = all_elements
        if mode == "text":
            payload["text"] = "\n\n".join(p["text"] for p in per_page)

        return payload
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}


def cmd_batch_ocr(args):
    """批量 OCR"""
    image_paths = args.get("image_paths", [])
    mode = args.get("mode", "text")
    
    if not image_paths:
        return {"success": False, "error": "未提供图片路径列表"}
    
    try:
        engine = _get_ocr_engine()
        results = []
        total_elapse = 0
        
        for img_path in image_paths:
            if not os.path.exists(img_path):
                results.append({"path": img_path, "success": False, "error": "文件不存在"})
                continue
            
            try:
                result, elapse = _invoke_engine(engine, img_path)
                total_elapse += _elapse_seconds(elapse)
                items = _collect_lines(result)
                items = _apply_formula_route(img_path, items, args.get("recognition_mode") or "auto")

                if mode == "pos":
                    results.append({"path": img_path, "success": True, "items": items})
                else:
                    results.append({
                        "path": img_path,
                        "success": True,
                        "text": "\n".join(item["text"] for item in items)
                    })
            except Exception as e:
                results.append({"path": img_path, "success": False, "error": str(e)})
        
        return {
            "success": True,
            "results": results,
            "total_elapse": total_elapse
        }
    except Exception as e:
        return {"success": False, "error": str(e), "traceback": traceback.format_exc()}

def cmd_status(args):
    """获取引擎状态"""
    try:
        engine_ok = _ocr_engine is not None
        return {
            "success": True,
            "status": "ready" if engine_ok else "not_initialized",
            "config": _config,
            "engine": "RapidOCR"
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

def cmd_set_config(args):
    """设置配置"""
    global _config, _ocr_engine
    config = args.get("config", {})
    _config.update(config)
    
    # 如果引擎已初始化，需要重建才能应用某些配置（如线程数）
    if "cpu_threads" in config and _ocr_engine is not None:
        _log("线程数配置变更，需要重建引擎...")
        _ocr_engine = None
    
    return {"success": True, "config": _config}

def process_command(line):
    """处理单条命令"""
    try:
        cmd_data = json.loads(line.strip())
        cmd = cmd_data.get("cmd", "")
        
        if cmd == "ocr":
            result = cmd_ocr(cmd_data)
        elif cmd == "ocr_pos":
            result = cmd_ocr_pos(cmd_data)
        elif cmd == "document":
            result = cmd_document(cmd_data)
        elif cmd == "batch_ocr":
            result = cmd_batch_ocr(cmd_data)
        elif cmd == "pdf_to_images":
            result = cmd_pdf_to_images(cmd_data)
        elif cmd == "ocr_pdf":
            result = cmd_ocr_pdf(cmd_data)
        elif cmd == "status":
            result = cmd_status(cmd_data)
        elif cmd == "set_config":
            result = cmd_set_config(cmd_data)
        elif cmd == "shutdown":
            result = {"success": True, "message": "引擎关闭中"}
            print(json.dumps(result), flush=True)
            sys.exit(0)
        else:
            result = {"success": False, "error": f"未知命令: {cmd}"}
    except json.JSONDecodeError as e:
        result = {"success": False, "error": f"JSON 解析失败: {str(e)}"}
    except Exception as e:
        result = {"success": False, "error": str(e), "traceback": traceback.format_exc()}
    
    try:
        print(json.dumps(_to_jsonable(result), ensure_ascii=False), flush=True)
    except Exception as e:
        _log(f"响应序列化失败: {e}")
        print(json.dumps({"success": False, "error": f"响应序列化失败: {e}"}, ensure_ascii=False), flush=True)

def main():
    """主循环"""
    _log("RapidOCR 推理引擎启动")
    
    # 尝试预热引擎
    try:
        _get_ocr_engine()
        _log("引擎预热完成")
    except Exception as e:
        _log(f"引擎预热失败（将在首次识别时重试）: {e}")
    
    # 输出就绪信号（供宿主 WaitForReady 检测）
    print(json.dumps({"success": True, "status": "ready", "engine": "RapidOCR"}), flush=True)
    
    # 主循环：逐行读取命令
    while True:
        try:
            line = sys.stdin.readline()
            if not line:
                break
            line = line.strip()
            if not line:
                continue
            process_command(line)
        except EOFError:
            break
        except Exception as e:
            _log(f"处理命令异常: {e}")
            try:
                print(json.dumps({"success": False, "error": str(e)}), flush=True)
            except:
                pass

if __name__ == "__main__":
    main()
