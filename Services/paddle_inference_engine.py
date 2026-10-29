#!/usr/bin/env python
# -*- coding: utf-8 -*-
# 类型检查器说明（basedpyright/pyright）：
# 本脚本实际运行在 Python 3.9（paddleocr/pix2tex/paddle 仅装在该环境），而分析器
# 默认使用工作区 Python（3.14，无这些重型依赖、无类型存根）。以下抑制项均为
# "分析环境与运行环境不一致"或"第三方库无类型存根"导致的噪音，不属于真实缺陷；
# 真实错误规则（未绑定变量/未使用导入/无效转义序列等）保持开启。
# pyright: reportMissingImports=none, reportOptionalMemberAccess=none, reportAttributeAccessIssue=none, reportUnknownMemberType=none, reportUnknownVariableType=none, reportUnknownArgumentType=none, reportUnknownParameterType=none, reportAny=none, reportUnannotatedClassAttribute=none, reportMissingParameterType=none, reportOptionalCall=none, reportOperatorIssue=none, reportGeneralTypeIssues=none, reportUnreachable=none, reportUnusedCallResult=none, reportUnknownLambdaType=none, reportImplicitStringConcatenation=none, reportPrivateUsage=none
"""
PaddleOCR 统一推理引擎
======================
【架构说明】
- 长驻进程模式：启动时一次性加载所有模型，通过 stdin/stdout JSON 协议通信
- 支持 PP-StructureV3（版面分析 + 表格识别 + 公式识别）+ PaddleOCR 通用识别
- 支持 PaddlePaddle 3.x 原生运行（移除强制降级）
- 支持 OpenVINO 加速
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
  "cmd": "document",         // 结构化文档识别（版面分析 + OCR）
  "image_path": "...",
  "lang": "ch",              // 语言
  "enable_layout": true,     // 启用版面分析
  "enable_table": true,      // 启用表格识别
  "enable_formula": true     // 启用公式识别
}

{
  "cmd": "batch_ocr",        // 批量 OCR
  "image_paths": ["...", "..."],
  "mode": "text"             // "text" 或 "pos"
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
    "cpu_threads": 4,
    "enable_mkldnn": true
  }
}
"""

import os
import sys
import subprocess

# ============================================================
# CRITICAL: 设置环境变量 - 必须在导入任何依赖之前执行
# 原因：PyTorch 在 Windows 上需要禁用共享内存以避免 shm.dll 错误
# ============================================================
os.environ["SHM_DISABLE"] = "1"

# ============================================================
# CRITICAL: Python 版本检查 - 必须在导入任何依赖之前执行
# 原因：numpy 2.0 与 imgaug/PaddleOCR 不兼容，会导致引擎启动失败
# ============================================================
def _ensure_compatible_python():
    """诊断运行本脚本的 Python 版本是否兼容（3.9）。

    历史做法：版本不符时用 `py -3.9` 重启自身（sys.exit(proc.wait())）。该"自重启"
    把环境准备与识别逻辑耦合，且在被 import 时可能错误地重启 *宿主* 脚本
    （sys.argv[0] 未必是本文件），行为不可预测。改为：仅做诊断输出，绝不重启、绝不
    退出进程。生产环境始终以 `py -3.9` 拉起本引擎，本检查只是兜底告警；若确实在不兼容
    的 Python 下运行，后续依赖导入会以清晰错误失败，便于定位。
    """
    if sys.version_info >= (3, 9) and sys.version_info < (3, 10):
        return
    # 非 3.9：明确告警并给出修复建议，但不重启、不退出。
    print(
        f"[环境警告] 当前 Python {sys.version_info.major}.{sys.version_info.minor} "
        f"非兼容版本（需 3.9）。numpy 2.x 与 PaddleOCR/pix2tex 不兼容，引擎可能无法启动。"
        f"请用 `py -3.9` 启动本引擎（不要依赖自动重启，避免误重启宿主进程）。",
        file=sys.stderr, flush=True,
    )

def _force_fix_numpy():
    """诊断 numpy 大版本是否兼容（需 1.x）。

    历史做法：检测到 numpy 2.x 时自动 `pip install numpy==1.26.4 --force-reinstall`
    并 `sys.exit(1)` 强杀进程——既在未经用户同意下改动 Python 环境，又会直接终止
    引擎（依赖外部监控重启）。改为：仅诊断输出并建议手动降级，绝不自动改环境、绝不
    退出进程。若环境确实为 numpy 2.x，后续依赖导入会以清晰错误失败，便于定位。
    """
    try:
        import importlib.util
        spec = importlib.util.find_spec("numpy")
        if spec is not None and spec.loader is not None:
            numpy = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(numpy)
            if hasattr(numpy, '__version__') and numpy.__version__.startswith('2'):
                print(
                    f"[环境警告] 检测到 numpy 2.x ({numpy.__version__})，与 PaddleOCR/pix2tex 不兼容。"
                    f"请手动执行：pip install numpy==1.26.4  （勿用 --force-reinstall，以免破坏其他环境）。"
                    f"引擎将继续尝试启动，但 paddle 相关功能可能不可用。",
                    file=sys.stderr, flush=True,
                )
    except Exception:
        pass

# 先检查并切换到兼容的 Python 版本
_ensure_compatible_python()

# 然后修复 numpy 版本（只有在正确的 Python 环境中才执行）
_force_fix_numpy()

import json
import traceback
import gc

try:
    # 作为包内模块被导入时（import Services.paddle_inference_engine）
    from .ocr_overlap_dedupe import dedupe_overlapping_items, exclusive_recognition_quads
except ImportError:
    try:
        # 作为脚本直接运行时（py Services/paddle_inference_engine.py）无包上下文，
        # 脚本所在目录已在 sys.path，按普通模块导入
        from ocr_overlap_dedupe import dedupe_overlapping_items, exclusive_recognition_quads  # pyright: ignore[reportImplicitRelativeImport]
    except Exception:
        def dedupe_overlapping_items(items):
            """依赖缺失时的降级：不做去重"""
            return items

        def exclusive_recognition_quads(quads):
            return quads

try:
    from .ocr_formula_router import recognize_whole, route_formulas
except ImportError:
    try:
        from ocr_formula_router import recognize_whole, route_formulas  # pyright: ignore[reportImplicitRelativeImport]
    except Exception:
        def route_formulas(image_path, items, recognition_mode="auto", log=None):
            return {"items": list(items or []), "formulas": [], "markdown_text": None, "applied": False}

        def recognize_whole(image_path, log=None):
            return ""

# ============================================================
# 全局配置
# ============================================================
CONFIG = {
    "lang": "ch",
    "cpu_threads": 2,       # Paddle 2.x + OpenBLAS 限制
    "enable_mkldnn": False, # Paddle 2.x 禁用 MKLDNN 避免崩溃
    "use_gpu": False,
    "drop_score": 0.5,
    "use_angle_cls": True,
    "use_textline_orientation": True,
}

# ============================================================
# 初始化环境
# ============================================================
def setup_environment():
    """设置运行环境"""
    # 编码修复 - 使用 reconfigure() 安全设置 UTF-8，避免 BOM 问题
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except (AttributeError, ValueError):
        try:
            # Python < 3.7 回退方案：先 detach 再重新包装
            if sys.stdout.encoding != 'utf-8':
                import io
                sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
                sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8')
        except Exception:
            pass  # 如果都失败，依赖 PYTHONIOENCODING 环境变量

    # Paddle 2.x 兼容环境变量
    os.environ['OMP_NUM_THREADS'] = str(CONFIG["cpu_threads"])
    os.environ['KMP_DUPLICATE_LIB_OK'] = 'TRUE'
    os.environ['PYTHONIOENCODING'] = 'utf-8'
    # 禁用 MKLDNN/OneDNN（Paddle 2.x + Windows + OpenBLAS 稳定性）
    os.environ['FLAGS_use_mkldnn'] = '0'

# ============================================================
# 推理引擎核心
# ============================================================

def _protect_response_for_json(obj):
    """
    递归扫描 response 对象，修复所有字符串中被损坏的 LaTeX。
    
    核心问题：\f (反斜杠+f) 在某些传输环节可能被误解析为 
    form feed 控制字符 (U+000C)，导致 \frac{f}{700} → rac{...} 或更糟
    
    修复策略：
    - 检测已知的损坏模式并还原
    - 确保输出的 \f 安全通过 JSON 传输
    """
    import re
    
    if isinstance(obj, dict):
        return {k: _protect_response_for_json(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [_protect_response_for_json(v) for v in obj]
    elif isinstance(obj, str):
        original = obj
        
        obj = re.sub(r'(?<![a-zA-Z\\])rac\{', r'\\frac{', obj)
        
        if obj != original:
            pass
        
        return obj
    return obj


class OcrInferenceEngine:
    """OCR 推理引擎 - 持有所有模型实例。公式分流在 ocr_formula_router。"""

    def __init__(self):
        self.paddleocr = None
        self.pp_structure = None
        self.paddleocr_available = False
        self.pp_structure_available = False
        self.formula_ocr = None
        self.formula_available = False
        self.initialized = False
        self.init_time = 0.0

    def _log(self, message):
        """日志输出到 stderr（不干扰 stdout 协议）"""
        import time
        print(f"[{time.strftime('%H:%M:%S')}] {message}", file=__import__('sys').stderr, flush=True)

    def _log_error(self, context, exception):
        """错误日志"""
        import traceback
        self._log(f"❌ {context}: {exception}")
        self._log(traceback.format_exc())

    def initialize(self):
        """初始化所有模型（一次性加载）"""
        import time
        start = time.time()
        try:
            self._init_paddle_version()
            self._init_paddleocr()
            self._init_pp_structure()
            if not self.formula_available:
                self._log("自动检测公式增强包...")
                self._auto_install_formula_ocr()
            self.initialized = True
        except Exception as e:
            self._log_error(f"引擎初始化失败", e)
            if self.paddleocr is not None:
                self.initialized = True

        self.init_time = time.time() - start
        self._log(f"引擎初始化完成，耗时 {self.init_time:.1f}s")

    def _init_paddle_version(self):
        """检查 PaddlePaddle 版本"""
        try:
            import paddle
            ver = paddle.__version__
            self._log(f"PaddlePaddle 版本: {ver}")
            self._log(f"PaddlePaddle 2.x 兼容模式")
            CONFIG["enable_mkldnn"] = False
            CONFIG["cpu_threads"] = 1

            if CONFIG["use_gpu"]:
                if paddle.is_compiled_with_cuda():
                    paddle.set_device('gpu')
                    self._log("使用 GPU 推理")
                else:
                    self._log("GPU 不可用，回退到 CPU")
                    CONFIG["use_gpu"] = False
            else:
                paddle.set_device('cpu')
                self._log("使用 CPU 推理")
        except ImportError:
            self._log("PaddlePaddle 未安装")

    def _init_paddleocr(self):
        """初始化 PaddleOCR（兼容 2.x: PP-OCRv4）"""
        try:
            from paddleocr import PaddleOCR

            ocr_kwargs = {
                "use_angle_cls": CONFIG["use_angle_cls"],
                "lang": CONFIG["lang"],
                "use_gpu": CONFIG["use_gpu"],
                "enable_mkldnn": False,
                "cpu_threads": CONFIG["cpu_threads"],
                "show_log": False,
                "drop_score": CONFIG["drop_score"],
            }

            self._log(f"初始化 PaddleOCR，参数: {ocr_kwargs}")
            self.paddleocr = PaddleOCR(**ocr_kwargs)
            self.paddleocr_available = True
            self._widen_rec_input()
            self._log("PaddleOCR 初始化成功")
        except ImportError as e:
            self._log(f"PaddleOCR 未安装: {e}，尝试自动安装...")
            if self._auto_install_paddleocr():
                from paddleocr import PaddleOCR
                ocr_kwargs = {
                    "use_angle_cls": CONFIG["use_angle_cls"], "lang": CONFIG["lang"],
                    "use_gpu": CONFIG["use_gpu"], "enable_mkldnn": False,
                    "cpu_threads": CONFIG["cpu_threads"], "show_log": False,
                    "drop_score": CONFIG["drop_score"],
                }
                self.paddleocr = PaddleOCR(**ocr_kwargs)
                self.paddleocr_available = True
                self._widen_rec_input()
                self._log("✅ PaddleOCR 自动安装并初始化成功")
            else:
                self._log_error("PaddleOCR 自动安装失败", e)
        except Exception as e:
            self._log_error("PaddleOCR 初始化失败", e)

    def _widen_rec_input(self):
        """放宽识别模型的输入宽度上限（默认 320 → 640）

        rec 模型输入高度固定 48、宽度上限 320（predict_rec.py 里硬截断）。
        宽行（如 18 位身份证号）按高度 48 折算后宽度远超 320，被压扁后
        每字仅十几像素，退化图上末位 X 常被误识为 8。PP-OCRv4 识别骨干是
        全卷积网络，推理宽度可自由放宽；批量 padding 会按新宽度补白，代价很小。
        """
        try:
            recognizer = getattr(self.paddleocr, "text_recognizer", None)
            if recognizer is not None and hasattr(recognizer, "rec_image_shape"):
                recognizer.rec_image_shape = [3, 48, 640]
                self._log("识别模型输入宽度已放宽: 320 -> 640")
        except Exception as e:
            self._log(f"放宽识别输入宽度失败（忽略，使用默认 320）: {e}")

    def _apply_hires_recognition(self, image_path, result, rec_boxes=None):
        """只对又小又不确定的文字行做一次批量重识别。

        以前会把整图放大，再对每一个文字框单独调用 ocr()。图纸上几十个尺寸
        就会串行推理几十次，一次识别要多等几十秒。

        现在：行数多的页面直接用第一次结果。行数少时，只重识别高度不足或
        置信度偏低的行，并合成一次识别。只有新结果置信度更高才替换。
        """
        try:
            if not result or not result[0]:
                return result
            lines = result[0]
            if len(lines) > 20:
                self._log(f"[高分辨率重识别] 跳过：{len(lines)} 行，不再逐行重识别")
                return result

            from PIL import Image
            img = Image.open(image_path).convert("RGB")
            W, H = img.size
            if max(W, H) > 2400:
                return result

            candidates = []
            for i, line in enumerate(lines):
                if not line or len(line) < 2:
                    continue
                rec = line[1]
                if not rec or len(rec) < 2:
                    continue
                box = line[0]
                pad = 3
                if rec_boxes and i < len(rec_boxes) and rec_boxes[i]:
                    box = rec_boxes[i]
                    if box != line[0]:
                        pad = 0
                text, conf = rec[0], rec[1]
                try:
                    xs = [float(p[0]) for p in box]
                    ys = [float(p[1]) for p in box]
                    conf_f = float(conf)
                except (TypeError, ValueError):
                    continue
                l, u = max(0, int(min(xs)) - pad), max(0, int(min(ys)) - pad)
                r, d = min(W, int(max(xs)) + pad), min(H, int(max(ys)) + pad)
                bw, bh = r - l, d - u
                if bw < 2 or bh < 2:
                    continue
                if min(bw, bh) >= 36 and conf_f >= 0.9:
                    continue
                crop = img.crop((l, u, r, d))
                ch = crop.size[1]
                if 0 < ch < 96:
                    scale = min(3.0, 96.0 / float(ch))
                    if scale > 1.05:
                        crop = crop.resize(
                            (max(1, int(round(crop.size[0] * scale))), max(1, int(round(crop.size[1] * scale)))),
                            Image.LANCZOS,
                        )
                candidates.append((i, text, conf_f, crop))

            if not candidates:
                return result

            recs = self._recognize_crop_batch([item[3] for item in candidates])
            if not recs or len(recs) != len(candidates):
                self._log("[高分辨率重识别] 批量识别不可用，保留第一次结果")
                return result

            new_lines = list(lines)
            replaced = 0
            for (i, text, conf, _crop), rec in zip(candidates, recs):
                if not rec:
                    continue
                t2, c2 = rec
                try:
                    c2 = float(c2)
                except (TypeError, ValueError):
                    continue
                if c2 > conf:
                    new_lines[i] = [lines[i][0], [t2, c2]]
                    replaced += 1
                    self._log(f"[高分辨率重识别] {text!r}@{conf:.3f} -> {t2!r}@{c2:.3f}")
            if replaced:
                self._log(f"[高分辨率重识别] 完成：{replaced}/{len(lines)} 行采用高分辨率结果（检查 {len(candidates)} 行）")
            return [new_lines] + list(result[1:])
        except Exception as e:
            self._log(f"[高分辨率重识别] 跳过（使用 1x 结果）: {e}")
            return result

    def _recognize_oriented(self, crops):
        """竖排尺寸先转成横排再认。左右两侧的直径经常是立着的框，横着读会剩一个字。"""
        batch = []
        slots = []
        for index, crop in enumerate(crops):
            if crop is None:
                continue
            batch.append(crop)
            slots.append((index, 0))
            width, height = crop.size
            if height >= 18 and height >= width * 1.45:
                batch.append(crop.rotate(-90, expand=True))
                slots.append((index, 1))
                batch.append(crop.rotate(90, expand=True))
                slots.append((index, 2))
        recs = self._recognize_crop_batch(batch)
        if not recs:
            return None
        chosen = [None] * len(crops)
        for (index, variant), rec in zip(slots, recs):
            if not rec:
                continue
            text, score = rec[0], rec[1]
            try:
                score = float(score)
            except (TypeError, ValueError):
                score = 0.0
            text = (text or "").strip()
            current = chosen[index]
            if current is None or self._prefer_oriented(current[0], current[1], current[2], text, score, variant):
                chosen[index] = (text, score, variant)
        out = []
        turned = 0
        for item in chosen:
            if item is None:
                out.append(None)
                continue
            text, score, variant = item
            if variant:
                turned += 1
            out.append((text, score))
        if turned:
            self._log(f"[竖排尺寸] {turned} 个立着的框转正后采用")
        return out

    def _prefer_oriented(self, old_text, old_score, old_variant, new_text, new_score, new_variant):
        if not new_text:
            return False
        if old_variant == 0 and new_variant == 0:
            return new_score > old_score
        if old_variant != 0 and new_variant != 0:
            return new_score > old_score
        upright_text, upright_score = (old_text, old_score) if old_variant == 0 else (new_text, new_score)
        turned_text, turned_score = (new_text, new_score) if old_variant == 0 else (old_text, old_score)
        upright_digit = any(ch.isdigit() for ch in upright_text)
        turned_digit = any(ch.isdigit() for ch in turned_text)
        take_turned = False
        if turned_digit and not upright_digit:
            take_turned = True
        elif turned_digit and upright_digit and len(turned_text) >= len(upright_text) + 2 and turned_score >= 0.45:
            take_turned = True
        if old_variant == 0:
            return take_turned
        return not take_turned

    def _recognize_crop_batch(self, crops, use_angle_cls=None):
        """一次识别多张行图。返回与 crops 等长的 (text, score)，失败返回 None。

        use_angle_cls 传 False 时跳过方向分类器——两个旋转方向要各识一次比置信度，
        分类器会把两个方向都掰成同一个，就比不出阅读方向了。
        """
        import numpy as np
        recognizer = getattr(self.paddleocr, "text_recognizer", None)
        if recognizer is None or not crops:
            return None
        arrays = []
        for crop in crops:
            arr = np.asarray(crop)
            if getattr(arr, "ndim", 0) == 3 and arr.shape[2] >= 3:
                arr = arr[:, :, :3][:, :, ::-1]
            arrays.append(np.ascontiguousarray(arr))
        imgs = arrays
        if use_angle_cls is None:
            use_angle_cls = CONFIG.get("use_angle_cls")
        if use_angle_cls:
            classifier = getattr(self.paddleocr, "text_classifier", None)
            if classifier is not None:
                classified = classifier(imgs)
                if isinstance(classified, tuple) and classified and classified[0] is not None:
                    imgs = classified[0]
        rec = recognizer(imgs)
        rec_res = rec[0] if isinstance(rec, tuple) else rec
        if not rec_res or len(rec_res) != len(crops):
            return None
        out = []
        for item in rec_res:
            if isinstance(item, (list, tuple)) and len(item) >= 2 and isinstance(item[0], str):
                out.append((item[0], item[1]))
            else:
                out.append(None)
        return out

    def _rerecognize_vertical(self, image_path, items):
        """竖排（旋转 90°）文字重识别，并定出阅读方向。

        为什么需要：检测器给的是轴对齐矩形。文字竖排时整个框又高又窄，识别器按横排
        喂进去——实测机械图侧边的竖排 Ø78.5 直接读成空、竖排 F8 读成 "8."（0.51）。
        接近正方形的短行（F8/H7/6H）连现有「高>=宽*1.5 才转一下」的判断都过不了。

        做法：对又高又窄的框分别顺/逆时针转 90° 各识别一次（不走方向分类器，
        两个方向结果不同才开始比出阅读方向），置信度高的赢。
        同时给 item 打上 vdir：up=原文从下往上读，down=从上往下读——
        下游合并被切成两半的竖排标注时按这个顺序接文本。
        """
        if not items:
            return items
        try:
            from PIL import Image
            pil = Image.open(image_path).convert("RGB")
        except Exception as e:
            self._log(f"[竖排重识别] 打不开图，跳过: {e}")
            return items
        W, H = pil.size
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
                self._log(f"[竖排重识别] 裁剪失败: {e}")
                continue
            orig_text = item.get("text", "") or ""
            orig_conf = float(item.get("confidence") or 0.0)
            rot_best = ("", 0.0, "")   # (text, conf, vdir)
            # PIL rotate 是逆时针为正：270=顺时针 90°，90=逆时针 90°。
            # 顺时针能读通 → 原文是逆时针放的 → 从下往上读（up）。
            for angle, vdir in ((270, "up"), (90, "down")):
                try:
                    rotated = crop.rotate(angle, expand=True)
                    out = self._recognize_crop_batch([rotated], use_angle_cls=False)
                except Exception as e:
                    self._log(f"[竖排重识别] 识别失败({angle}): {e}")
                    continue
                if not out or not out[0]:
                    continue
                text, conf = out[0]
                if text and conf > rot_best[1]:
                    rot_best = (text, conf, vdir)
            changed = False
            # 文本：旋转后明显更通才替换（原读法也不算差的就留住，避免无谓抖动）
            if rot_best[0] and rot_best[1] > orig_conf and rot_best[0] != orig_text:
                item["text"] = rot_best[0]
                item["confidence"] = round(float(rot_best[1]), 4)
                changed = True
            # 方向：两个旋转里赢的那个就是阅读方向，和文本换不换无关——
            # 已经读对的框也要有方向，下游合并拆分盒、UI 画朝向都靠它。
            if rot_best[2] and rot_best[1] >= 0.3 and not item.get("vdir"):
                item["vdir"] = rot_best[2]
                changed = True
            if changed:
                fixed += 1
        if fixed:
            self._log(f"[竖排重识别] 修正 {fixed} 个竖排文本框")
        return items

    def _crop_quad(self, pil, quad):
        """按四边形裁出送去识别的图。水平框直接切矩形，斜的框做透视拉正。"""
        import numpy as np
        from PIL import Image
        W, H = pil.size
        xs = [float(p[0]) for p in quad]
        ys = [float(p[1]) for p in quad]
        dx = float(quad[1][0]) - float(quad[0][0])
        dy = float(quad[1][1]) - float(quad[0][1])
        horizontal = abs(dx) < 1 or abs(dy) <= abs(dx) * 0.08
        if horizontal:
            l = max(0, int(min(xs)))
            t = max(0, int(min(ys)))
            r = min(W, int(round(max(xs))))
            b = min(H, int(round(max(ys))))
            if r - l < 2 or b - t < 2:
                return None
            return pil.crop((l, t, r, b))
        try:
            import cv2
        except ImportError:
            return None
        rgb = np.asarray(pil)
        pts = np.array(quad, dtype="float32")
        width = max(np.linalg.norm(pts[0] - pts[1]), np.linalg.norm(pts[2] - pts[3]))
        height = max(np.linalg.norm(pts[0] - pts[3]), np.linalg.norm(pts[1] - pts[2]))
        width = int(max(2, round(float(width))))
        height = int(max(2, round(float(height))))
        dst = np.array([[0, 0], [width, 0], [width, height], [0, height]], dtype="float32")
        matrix = cv2.getPerspectiveTransform(pts, dst)
        warped = cv2.warpPerspective(rgb, matrix, (width, height), borderMode=cv2.BORDER_REPLICATE)
        if height >= width * 1.5:
            warped = np.rot90(warped)
        return Image.fromarray(np.ascontiguousarray(warped))

    def _ocr_with_exclusive_crops(self, image_path):
        """先检测，再识别。同一行交叠的那一截不送进右边的识别。

        返回 (paddle 结果, 实际送去识别的框)。失败返回 None，调用方改走整段 ocr()。
        结果里的框仍是检测框，显示位置不变。
        """
        import numpy as np
        from PIL import Image
        detector = getattr(self.paddleocr, "text_detector", None)
        if detector is None:
            return None
        try:
            pil = Image.open(image_path).convert("RGB")
            rgb = np.asarray(pil)
            if getattr(rgb, "ndim", 0) != 3 or rgb.shape[2] < 3:
                return None
            bgr = np.ascontiguousarray(rgb[:, :, :3][:, :, ::-1])
            detected = detector(bgr)
        except Exception as e:
            self._log(f"[识别裁切] 检测失败，改走整段识别: {e}")
            return None
        boxes = detected[0] if isinstance(detected, tuple) else detected
        if boxes is None:
            return ([[]], [])
        quads = []
        for box in boxes:
            try:
                pts = [[float(p[0]), float(p[1])] for p in list(box)[:4]]
            except (TypeError, ValueError, IndexError):
                continue
            if len(pts) >= 4:
                quads.append(pts)
        if not quads:
            return ([[]], [])
        quads.sort(key=lambda q: (min(p[1] for p in q), min(p[0] for p in q)))
        rec_quads = exclusive_recognition_quads(quads)
        crops = [self._crop_quad(pil, quad) for quad in rec_quads]
        if any(crop is None for crop in crops):
            self._log("[识别裁切] 有框裁不出来，改走整段识别")
            return None
        recs = self._recognize_oriented(crops)
        if not recs:
            self._log("[识别裁切] 识别器不可用，改走整段识别")
            return None
        lines = []
        clipped = 0
        for quad, rec_quad, rec in zip(quads, rec_quads, recs):
            if rec:
                text, score = rec
                try:
                    score = float(score)
                except (TypeError, ValueError):
                    score = 0.0
            else:
                text, score = "", 0.0
            lines.append([quad, [text, score]])
            if rec_quad != quad:
                clipped += 1
        if clipped:
            self._log(f"[识别裁切] {clipped} 个框的交叠区域没有送进识别")
        return ([lines], rec_quads)

    def _recognize_page(self, image_path):
        packed = self._ocr_with_exclusive_crops(image_path)
        if packed is None:
            result = self.paddleocr.ocr(image_path, cls=True)
            return self._apply_hires_recognition(image_path, result)
        result, rec_boxes = packed
        return self._apply_hires_recognition(image_path, result, rec_boxes)

    def _auto_install_paddleocr(self):
        """自动安装 PaddlePaddle + PaddleOCR"""
        import subprocess, sys
        try:
            self._log("正在安装 paddlepaddle（可能需要几分钟）...")
            r1 = subprocess.run(
                [sys.executable, "-m", "pip", "install", "paddlepaddle", "-q"],
                capture_output=True, text=True, timeout=600, encoding='utf-8'
            )
            if r1.returncode != 0:
                self._log(f"paddlepaddle 安装失败: {r1.stderr[-200:]}")
                return False
            self._log("paddlepaddle 安装成功，正在安装 paddleocr...")
            r2 = subprocess.run(
                [sys.executable, "-m", "pip", "install", "paddleocr", "-q"],
                capture_output=True, text=True, timeout=600, encoding='utf-8'
            )
            if r2.returncode != 0:
                self._log(f"paddleocr 安装失败: {r2.stderr[-200:]}")
                return False
            self._fix_numpy()
            return True
        except subprocess.TimeoutExpired:
            self._log("安装超时")
            return False
        except Exception as e:
            self._log(f"安装异常: {e}")
            return False

    def _fix_numpy(self):
        """确保 numpy 版本兼容（<2）"""
        import subprocess, sys
        try:
            import numpy
            if numpy.__version__.startswith('2'):
                subprocess.run(
                    [sys.executable, "-m", "pip", "install", "numpy==1.26.4", "--force-reinstall", "-q"],
                    capture_output=True, timeout=120
                )
                self._log("已修复 numpy 版本到 1.x")
        except Exception:
            pass

    def _init_pp_structure(self):
        """PP-Structure 初始化"""
        try:
            from paddleocr import PPStructure
            self.pp_structure = PPStructure(
                lang=CONFIG["lang"],
                use_gpu=CONFIG["use_gpu"],
                show_log=False,
                layout=True,
                table=True,
                ocr=True,
            )
            self.pp_structure_available = True
            self._log("PP-Structure 初始化成功，支持版面/表格/公式识别")
            return
        except ImportError:
            self._log("PP-Structure 不可用（paddleocr 版本较低），跳过版面分析/表格/公式")
        except Exception as e:
            self._log(f"PP-Structure 初始化失败: {e}")

        self._init_formula_ocr()

    def _init_formula_ocr(self):
        """标记公式识别可用（模型在首次识别时懒加载）"""
        try:
            import importlib.util
            if importlib.util.find_spec("pix2tex") is not None:
                self.formula_available = True
                self.formula_ocr = None
                self._log("LaTeX-OCR (pix2tex) 已就绪（模型将在首次识别时加载）")
            else:
                self._log("pix2tex 未安装，公式识别不可用")
        except Exception as e:
            self._log(f"公式识别检测失败: {e}")

    def ocr_text(self, image_path, recognition_mode="auto"):
        """纯文本 OCR（支持识别模式选择）
        
        参数:
        - recognition_mode: 识别模式
          - "auto": 自动检测（根据图片特征判断是否使用公式识别）
          - "text": 纯文本模式（禁用公式识别）
          - "formula": 公式模式（优先使用公式识别）
          - "number": 数字模式（纯文本，优化数字识别）
          - "date": 日期模式（纯文本，优化日期识别）
        
        返回格式说明：
        - text: 纯识别结果（LaTeX 或 OCR 文本）
        - markdown_text: 同 text（$$ 包裹由前端渲染时统一处理）
        """
        self._log(f"[ocr_text] 入口: mode={recognition_mode}, formula_available={self.formula_available}, paddleocr={self.paddleocr_available}")

        if not self.paddleocr_available:
            if recognition_mode == "formula":
                latex = recognize_whole(image_path, log=self._log)
                if latex:
                    return {"success": True, "text": latex, "markdown_text": f"$$\n{latex}\n$$"}
            return {"success": False, "error": "PaddleOCR 未初始化"}

        result = self._recognize_page(image_path)
        raw_items = []
        if result and result[0]:
            for line in result[0]:
                if line:
                    raw_items.append({
                        "box": [[float(p[0]), float(p[1])] for p in line[0]],
                        "text": line[1][0],
                        "confidence": float(line[1][1])
                    })

        before_texts = [it.get("text") for it in raw_items]
        raw_items = dedupe_overlapping_items(raw_items)
        if before_texts != [it.get("text") for it in raw_items]:
            self._log(f"[重叠字] 去掉检测框交界处的重复字: {before_texts} -> {[it.get('text') for it in raw_items]}")

        raw_items = self._rerecognize_vertical(image_path, raw_items)

        routed = route_formulas(image_path, raw_items, recognition_mode, log=self._log)
        raw_items = routed["items"]
        texts = [it.get("text", "") for it in raw_items]
        markdown = routed["markdown_text"] or "\n".join(texts)
        return {"success": True, "text": "\n".join(texts), "markdown_text": markdown, "raw_data": raw_items}

    def ocr_with_position(self, image_path, recognition_mode="auto"):
        """带坐标的 OCR（支持识别模式选择）
        
        参数:
        - recognition_mode: 识别模式 (auto/text/formula/number/date)
        """
        self._log(f"[ocr_with_position] 入口: mode={recognition_mode}")

        if not self.paddleocr_available:
            return {"success": False, "error": "PaddleOCR 未初始化"}

        result = self._recognize_page(image_path)
        data = []
        if result and result[0]:
            for line in result[0]:
                if line:
                    box = line[0]
                    text = line[1][0]
                    confidence = float(line[1][1])
                    data.append({
                        "text": text,
                        "confidence": round(confidence, 4),
                        "box": [[float(p[0]), float(p[1])] for p in box]
                    })

        before_texts = [it.get("text") for it in data]
        data = dedupe_overlapping_items(data)
        after_texts = [it.get("text") for it in data]
        if before_texts != after_texts:
            self._log(f"[重叠字] 去掉检测框交界处的重复字: {before_texts} -> {after_texts}")

        data = self._rerecognize_vertical(image_path, data)

        routed = route_formulas(image_path, data, recognition_mode, log=self._log)
        output: dict[str, object] = {
            "success": True,
            "items": routed["items"],
            "raw_data": routed["items"],
        }
        if routed["applied"]:
            output["formulas"] = routed["formulas"]
            output["markdown_text"] = routed["markdown_text"]
        return output

    def analyze_document(self, image_path):
        """结构化文档分析（版面+表格+公式+OCR）
        
        设计原则：后端不添加 $$ 标记，前端渲染时统一处理。
        """
        if self.pp_structure_available:
            return self._analyze_with_pp_structure(image_path)

        self._log("PP-Structure 不可用，使用常规 OCR + 公式识别")
        return self.ocr_with_position(image_path)

    def _analyze_with_pp_structure(self, image_path):
        """使用 PP-Structure 进行结构化文档分析"""
        try:
            results = self.pp_structure(image_path)
            output: dict[str, object] = {
                "success": True,
                "layout_boxes": [],
                "tables": [],
                "formulas": [],
                "overall_text": "",
                "markdown_text": "",
            }
            # 收集带框的 OCR 行（若 PP-Structure 提供 box），供区域级公式兜底使用
            overall_items = []

            for res in results:
                # 布局分析结果
                if hasattr(res, 'layout_det_res') and res.layout_det_res:
                    boxes = res.layout_det_res.get("boxes", [])
                    for box in boxes:
                        output["layout_boxes"].append({
                            "label": box.get("label", ""),
                            "score": round(float(box.get("score", 0)), 4),
                            "coordinate": box.get("coordinate", []),
                        })

                # 表格识别结果
                if hasattr(res, 'table_res_list') and res.table_res_list:
                    for table in res.table_res_list:
                        output["tables"].append({
                            "pred_html": table.get("pred_html", ""),
                            "cell_count": len(table.get("cell_box_list", [])),
                        })

                # 公式识别结果
                if hasattr(res, 'formula_res_list') and res.formula_res_list:
                    for formula in res.formula_res_list:
                        output["formulas"].append({
                            "latex": formula.get("rec_formula", ""),
                            "polygon": formula.get("rec_polys", []),
                        })

                # Markdown 输出（含表格和公式渲染）
                if hasattr(res, 'markdown') and res.markdown:
                    md_info = res.markdown
                    output["markdown_text"] = md_info.get("markdown_texts", "")

                # 整体 OCR 文本
                if hasattr(res, 'overall_ocr_res') and res.overall_ocr_res:
                    texts = []
                    if isinstance(res.overall_ocr_res, list):
                        for item in res.overall_ocr_res:
                            if isinstance(item, dict) and "text" in item:
                                texts.append(item["text"])
                                box = item.get("box")
                                if box:
                                    overall_items.append({
                                        "text": item["text"],
                                        "box": [[float(p[0]), float(p[1])] for p in box],
                                    })
                    output["overall_text"] = "\n".join(texts)

                # 尝试 save_to_json 保存并读取
                try:
                    import tempfile
                    tmp_dir = tempfile.mkdtemp()
                    res.save_to_json(save_path=tmp_dir)

                    # 保存 Markdown 文件
                    if hasattr(res, 'save_to_markdown'):
                        res.save_to_markdown(save_path=tmp_dir)

                    # 清理临时目录
                    import shutil
                    shutil.rmtree(tmp_dir, ignore_errors=True)
                except Exception:
                    pass  # 非关键操作

            # 如果没有 markdown 文本，从布局构造
            if not output["markdown_text"] and output["overall_text"]:
                output["markdown_text"] = output["overall_text"]

            # 版面分析没给出公式时，按文字框分流。整页只在完全没有文字框时才送公式模型。
            if not output["formulas"] and overall_items:
                routed = route_formulas(image_path, overall_items, "auto", log=self._log)
                if routed["applied"]:
                    for formula in routed["formulas"]:
                        output["formulas"].append({
                            "latex": formula.get("latex", ""),
                            "polygon": formula.get("box") or [],
                        })
                    extra = "\n".join(
                        f"$$\n{formula['latex']}\n$$" for formula in routed["formulas"] if formula.get("latex")
                    )
                    if extra:
                        output["markdown_text"] = ((output.get("markdown_text") or "") + "\n\n" + extra).strip()
            elif not output["formulas"]:
                latex = recognize_whole(image_path, log=self._log)
                if latex:
                    output["formulas"].append({"latex": latex, "polygon": []})
                    output["markdown_text"] = ((output.get("markdown_text") or "") + f"\n\n$$\n{latex}\n$$").strip()

            return output

        except Exception as e:
            self._log_error("文档分析失败", e)
            # 降级到带位置的 OCR
            return self.ocr_with_position(image_path)

    def batch_ocr(self, image_paths, mode="text"):
        """批量 OCR"""
        results = []
        for path in image_paths:
            try:
                if mode == "pos":
                    res = self.ocr_with_position(path)
                else:
                    res = self.ocr_text(path)
                res["file"] = path
                results.append(res)
            except Exception as e:
                results.append({
                    "file": path,
                    "success": False,
                    "error": str(e)
                })
        return {"success": True, "results": results}

    def get_status(self):
        """获取引擎状态"""
        # 尝试获取 paddle 版本
        paddle_version = "N/A"
        try:
            import paddle
            paddle_version = paddle.__version__
        except Exception:
            pass

        return {
            "success": True,
            "status": {
                "initialized": self.initialized,
                "init_time_s": round(self.init_time, 1),
                "paddle_version": paddle_version,
                "paddleocr_available": self.paddleocr_available,
                "pp_structure_available": self.pp_structure_available,
                "formula_available": self.formula_available,
                "config": CONFIG,
            }
        }

    def _auto_install_formula_ocr(self):
        """自动安装 pix2tex（引擎初始化时静默调用，用户无需手动操作）"""
        import subprocess, sys, importlib.util
        try:
            if importlib.util.find_spec("pix2tex") is not None:
                self._log("pix2tex 已安装")
                self.formula_available = True
                self.formula_ocr = None  # 懒加载
                return
            self._log("正在自动安装公式识别增强包（pix2tex），请稍候...")
            proc = subprocess.run(
                [sys.executable, "-m", "pip", "install", "pix2tex", "--no-deps", "-q"],
                capture_output=True, text=True, timeout=600, encoding='utf-8'
            )
            if proc.returncode == 0 and importlib.util.find_spec("pix2tex") is not None:
                self.formula_available = True
                self.formula_ocr = None
                self._log("✅ 公式识别增强包自动安装成功（模型将在首次使用时加载）")
            else:
                self._log("⚠️ 公式增强包安装失败，可稍后手动安装（不影响文字OCR功能）")
        except Exception as e:
            self._log(f"⚠️ 公式增强包自动安装异常: {e}（不影响文字OCR功能）")

    def install_formula_ocr(self):
        """安装 pix2tex（LaTeX-OCR）公式识别增强包"""
        import subprocess
        import sys

        if self.formula_available:
            return {"success": True, "message": "公式识别已安装", "installed": True}

        self._log("开始安装 pix2tex（公式识别增强包）...")

        def _try_install(extra_args, label):
            # 用 --no-deps 避免触发 opencv-python-headless 依赖导致 numpy 被升级
            # （opencv-python-headless>=4.9 强制要求 numpy>=2，与 paddle 冲突）
            pip_cmd = [sys.executable, "-m", "pip", "install", "pix2tex", "--no-deps"] + extra_args
            self._log(f"[{label}] 执行: {' '.join(pip_cmd)}")
            proc = subprocess.run(
                pip_cmd,
                capture_output=True,
                text=True,
                timeout=600,
                encoding='utf-8'
            )
            return proc

        try:
            # 先尝试 PyPI 官方源
            proc = _try_install([], "官方源")
            if proc.returncode != 0:
                err = proc.stderr[-800:] if proc.stderr else ""
                self._log(f"官方源安装失败: {err}")
                # 再尝试清华镜像
                proc = _try_install(["-i", "https://pypi.tuna.tsinghua.edu.cn/simple"], "清华镜像")

            if proc.returncode == 0:
                # 验证 pix2tex 可导入（不立即创建 LatexOCR，因为首次会下载116MB模型）
                try:
                    import importlib.util
                    spec = importlib.util.find_spec("pix2tex")
                    if spec is not None:
                        self._log("pix2tex 安装成功（模型将在首次使用时自动下载）")
                        return {"success": True, "message": "公式识别增强包安装成功，请重启引擎生效", "installed": True}
                    else:
                        return {"success": False, "message": "安装完成但模块未找到，请重启引擎", "installed": False}
                except Exception as e:
                    return {"success": False, "message": f"安装后验证失败: {e}", "installed": False}
            else:
                err = proc.stderr[-1000:] if proc.stderr else "未知错误"
                self._log(f"pix2tex 安装失败: {err}")
                hint = ""
                if "No matching distribution" in err:
                    hint = "（提示：pix2tex 可能不支持当前 Python 版本）"
                return {"success": False, "message": f"安装失败: {err} {hint}", "installed": False}
        except subprocess.TimeoutExpired:
            self._log("pix2tex 安装超时（超过10分钟）")
            return {"success": False, "message": "安装超时，请检查网络连接", "installed": False}
        except Exception as e:
            self._log(f"pix2tex 安装异常: {e}")
            return {"success": False, "message": f"安装异常: {str(e)}", "installed": False}

    def set_config(self, config_update):
        """动态更新配置。语言变更会重新加载识别模型。"""
        lang_changed = False
        for key, value in config_update.items():
            if key in CONFIG:
                old_val = CONFIG[key]
                if key == "lang" and str(old_val) != str(value):
                    lang_changed = True
                CONFIG[key] = value
                self._log(f"配置更新: {key}: {old_val} -> {value}")
            else:
                self._log(f"未知配置项: {key}")

        # 更新线程数和 MKLDNN 实时生效
        if "cpu_threads" in config_update:
            os.environ['OMP_NUM_THREADS'] = str(CONFIG["cpu_threads"])
        if "enable_mkldnn" in config_update:
            os.environ['FLAGS_use_mkldnn'] = '1' if CONFIG["enable_mkldnn"] else '0'

        # 语言写进 CONFIG 不会换掉已经加载的识别模型。中文模型没有英文词表，
        # 直径符号、圆圈和剖视标记会被读成「国」。语言变了就丢掉旧模型再加载。
        if lang_changed and self.paddleocr is not None:
            self._log(f"识别语言改为 {CONFIG['lang']}，重新加载模型")
            try:
                del self.paddleocr
            except Exception:
                pass
            self.paddleocr = None
            self.paddleocr_available = False
            if self.pp_structure is not None:
                try:
                    del self.pp_structure
                except Exception:
                    pass
                self.pp_structure = None
                self.pp_structure_available = False
            gc.collect()
            self._init_paddleocr()

        return {"success": True, "lang": CONFIG.get("lang")}

    def shutdown(self):
        """清理并关闭"""
        self._log("正在关闭推理引擎...")
        if self.paddleocr:
            try:
                del self.paddleocr
            except Exception:
                pass
            self.paddleocr = None
        if self.pp_structure:
            try:
                del self.pp_structure
            except Exception:
                pass
            self.pp_structure = None

        gc.collect()
        self._log("推理引擎已关闭")
        return {"success": True, "message": "引擎已关闭"}


# ============================================================
# JSON 协议主循环
# ============================================================
def main():
    """入口：JSON 协议主循环"""
    setup_environment()
    # numpy 版本检查已在文件开头执行（必须在导入任何依赖之前）

    engine = OcrInferenceEngine()

    # 发送就绪信号
    ready_msg = json.dumps({"event": "ready", "message": "PaddleOCR 推理引擎已就绪"}, ensure_ascii=False)
    sys.stdout.write(ready_msg + "\n")
    sys.stdout.flush()

    # 在后台初始化（首次命令触发，前台快速响应）
    engine.initialize()

    # BOM 处理辅助函数
    def _strip_bom(text):
        """去除字符串开头的 UTF-8 BOM"""
        if text and text[0] == '\ufeff':
            return text[1:]
        return text

    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue

        # 去除可能的 BOM 字符
        line = _strip_bom(line)

        try:
            request = json.loads(line)
        except json.JSONDecodeError as e:
            response = {"success": False, "error": f"JSON 解析错误: {e}"}
            safe_resp = _protect_response_for_json(response)
            sys.stdout.write(json.dumps(safe_resp, ensure_ascii=False) + "\n")
            sys.stdout.flush()
            continue

        cmd = request.get("cmd", "")
        response = None

        try:
            if cmd == "ocr":
                image_path = request.get("image_path", "")
                recognition_mode = request.get("recognition_mode", "auto")  # auto/text/formula/number/date
                if not os.path.exists(image_path):
                    response = {"success": False, "error": f"图片文件不存在: {image_path}"}
                else:
                    response = engine.ocr_text(image_path, recognition_mode=recognition_mode)

            elif cmd == "ocr_pos":
                image_path = request.get("image_path", "")
                recognition_mode = request.get("recognition_mode", "auto")  # auto/text/formula/number/date
                if not os.path.exists(image_path):
                    response = {"success": False, "error": f"图片文件不存在: {image_path}"}
                else:
                    response = engine.ocr_with_position(image_path, recognition_mode=recognition_mode)

            elif cmd == "document":
                image_path = request.get("image_path", "")
                if not os.path.exists(image_path):
                    response = {"success": False, "error": f"图片文件不存在: {image_path}"}
                else:
                    # 允许单次请求覆盖全局配置；finally 保证异常时也能恢复
                    old_lang = None
                    if "lang" in request:
                        old_lang = CONFIG["lang"]
                        CONFIG["lang"] = request["lang"]
                        engine._log(f"临时切换语言: {old_lang} -> {CONFIG['lang']}")
                    try:
                        response = engine.analyze_document(image_path)
                    finally:
                        if old_lang is not None:
                            CONFIG["lang"] = old_lang

            elif cmd == "batch_ocr":
                image_paths = request.get("image_paths", [])
                mode = request.get("mode", "text")
                valid_paths = [p for p in image_paths if os.path.exists(p)]
                if not valid_paths:
                    response = {"success": False, "error": "没有有效的图片文件"}
                else:
                    response = engine.batch_ocr(valid_paths, mode)

            elif cmd == "status":
                response = engine.get_status()

            elif cmd == "set_config":
                config_update = request.get("config", {})
                response = engine.set_config(config_update)

            elif cmd == "install_formula":
                response = engine.install_formula_ocr()

            elif cmd == "shutdown":
                response = engine.shutdown()
                safe_resp = _protect_response_for_json(response)
                sys.stdout.write(json.dumps(safe_resp, ensure_ascii=False) + "\n")
                sys.stdout.flush()
                break

            else:
                response = {"success": False, "error": f"未知命令: {cmd}"}

        except Exception as e:
            response = {
                "success": False,
                "cmd": cmd,
                "error": str(e),
                "traceback": traceback.format_exc()
            }
            engine._log_error(f"命令 '{cmd}' 执行失败", e)

        if response is not None:
            # 🔒 安全防护：修复 LaTeX 中的 \f 等危险序列，防止 JSON 传输中被吞噬
            safe_response = _protect_response_for_json(response)
            output_json = json.dumps(safe_response, ensure_ascii=False)
            sys.stdout.write(output_json + "\n")
            sys.stdout.flush()


if __name__ == "__main__":
    # 版本/环境兼容性诊断已在模块导入时由 _ensure_compatible_python() / _force_fix_numpy()
    # 完成（仅告警，不再自动重启或强杀进程）。此处直接启动识别主循环。
    main()
