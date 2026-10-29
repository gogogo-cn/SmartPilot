#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
StandaloneAgent - Excel 关闭后仍可执行定时任务、操作 Excel 文件
技术栈: Python + openpyxl + 标准库 http.server
启动: python standalone_agent.py [--config PATH_TO_llm_config.json] [--port 9529]
"""
import json
import os
import sys
import time
import threading
import datetime
import argparse
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse, parse_qs
from urllib.request import Request, urlopen

try:
    import openpyxl
except ImportError:
    print("[StandaloneAgent] 缺少 openpyxl，请执行: pip install openpyxl")
    sys.exit(1)

# ---- 先解析命令行参数（--config / --port），确保 find_config 能用到 CLI 覆盖 ----
_CMD_ARGS = None
try:
    _ap = argparse.ArgumentParser(add_help=False)
    _ap.add_argument('--config', type=str, default=None, help='显式指定 llm_config.json 路径')
    _ap.add_argument('--port', type=int, default=9529, help='HTTP 服务端口')
    _CMD_ARGS, _ = _ap.parse_known_args()
except Exception:
    _CMD_ARGS = argparse.Namespace(config=None, port=9529)

CMD_LINE_CONFIG_PATH = _CMD_ARGS.config
# 允许环境变量覆盖端口（优先级：命令行 > 环境变量 > 默认 9529）
_ENV_PORT = os.environ.get('STANDALONE_PORT')
if _CMD_ARGS.port == 9529 and _ENV_PORT and _ENV_PORT.isdigit():
    PORT_VAL = int(_ENV_PORT)
else:
    PORT_VAL = int(_CMD_ARGS.port or 9529)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = PORT_VAL

def find_config():
    # 1) 显式命令行参数 / 环境变量：优先级最高（MobileControlBridge 启动时 --config 必传）
    explicit_from_cli = globals().get('CMD_LINE_CONFIG_PATH')
    if explicit_from_cli and os.path.isfile(explicit_from_cli):
        return explicit_from_cli
    env = os.environ.get('STANDALONE_CONFIG_JSON') or os.environ.get('LLM_CONFIG_JSON')
    if env and os.path.isfile(env):
        return env
    # 2) 唯一权威位置（SSOT）：VSTO ConfigDirectory 下的 llm_config.json
    #    - 与 ExportLLMConfigForStandalone 写入位置、MobileControlBridge --config 传入位置保持一致
    ssot_candidates = [
        # 真实 VSTO AppData 配置目录（ConfigConstants.APP_DATA_FOLDER = "ExcelAIVSTOAddIn"）
        os.path.join(os.environ.get('APPDATA', ''), 'ExcelAIVSTOAddIn', 'Config', 'llm_config.json'),
        # bin/Debug|Release/Config/llm_config.json（开发/调试阶段 ConfigDirectory 就在这里）
        os.path.join(BASE_DIR, "..", "Config", "llm_config.json"),
        os.path.join(BASE_DIR, "..", "..", "bin", "Debug", "Config", "llm_config.json"),
        os.path.join(BASE_DIR, "..", "..", "bin", "Release", "Config", "llm_config.json"),
    ]
    for c in ssot_candidates:
        p = os.path.abspath(c)
        if os.path.isfile(p):
            return p
    # 3) 最后兜底：脚本目录自身的 llm_config.json（Excel 未运行、完全独立运行时用户手放）
    legacy = os.path.abspath(os.path.join(BASE_DIR, "llm_config.json"))
    if os.path.isfile(legacy):
        return legacy
    return None

CONFIG_PATH = find_config()
TASKS_FILE = os.path.join(BASE_DIR, "tasks.json")
# 删除墓碑：API 删除的任务 ID 持久化记录，前端 localStorage 回推时直接拒绝（410），防止已删除任务复活
DELETED_FILE = os.path.join(BASE_DIR, "deleted_ids.json")
_task_lock = threading.Lock()
_deleted_lock = threading.Lock()
_deleted_ids = set()
_tasks = []

def ssot_write_path():
    """返回写入 llm_config.json 的权威路径（优先 AppData），不存在则创建目录。"""
    # 与 ExportLLMConfigForStandalone 保持一致的 SSOT 第一候选
    appdata = os.environ.get('APPDATA', '')
    if appdata:
        d = os.path.abspath(os.path.join(appdata, 'ExcelAIVSTOAddIn', 'Config'))
        try:
            if not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            return os.path.join(d, 'llm_config.json')
        except Exception:
            pass
    # 兜底：如果 --config 已经指向一个有目录的有效路径，就写那里
    if CONFIG_PATH and os.path.isabs(CONFIG_PATH):
        try:
            d = os.path.dirname(CONFIG_PATH)
            if not os.path.isdir(d):
                os.makedirs(d, exist_ok=True)
            return CONFIG_PATH
        except Exception:
            pass
    # 最后兜底：脚本所在目录
    try:
        if not os.path.isdir(BASE_DIR):
            os.makedirs(BASE_DIR, exist_ok=True)
    except Exception:
        pass
    return os.path.join(BASE_DIR, 'llm_config.json')

def _load_tasks():
    global _tasks
    try:
        if os.path.isfile(TASKS_FILE):
            with open(TASKS_FILE, "r", encoding="utf-8") as f:
                _tasks = json.load(f)
    except Exception:
        _tasks = []

def _save_tasks():
    try:
        with open(TASKS_FILE, "w", encoding="utf-8") as f:
            json.dump(_tasks, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[StandaloneAgent] 保存任务失败:", e)

def _load_deleted():
    global _deleted_ids
    try:
        if os.path.isfile(DELETED_FILE):
            with open(DELETED_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            _deleted_ids = set(data if isinstance(data, list) else [])
    except Exception:
        _deleted_ids = set()

def _save_deleted():
    try:
        with open(DELETED_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(_deleted_ids), f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("[StandaloneAgent] 保存删除墓碑失败:", e)

def _next_run(task, now=None):
    now = now or datetime.datetime.now()
    freq = task.get("Frequency", "daily")
    if freq == "hourly":
        return now + datetime.timedelta(hours=1)
    time_str = task.get("Time") or "09:00"
    parts = time_str.split(":")
    hour = int(parts[0]) if parts and parts[0].isdigit() else 9
    minute = int(parts[1]) if len(parts) > 1 and parts[1].isdigit() else 0
    result = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if result <= now:
        result += datetime.timedelta(days=1)
    if freq == "weekdays":
        while result.weekday() >= 5:
            result += datetime.timedelta(days=1)
    elif freq == "weekly":
        while result.weekday() != 0:
            result += datetime.timedelta(days=1)
    return result

def llm_config():
    if not CONFIG_PATH:
        return None
    try:
        # utf-8-sig 兼容带 BOM 的导出（.NET Encoding.UTF8 会写 BOM）
        with open(CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            return json.load(f)
    except Exception:
        return None

def default_api_url(provider):
    url = {
        "openai": "https://api.openai.com/v1/chat/completions",
        "azure": "https://api.openai.azure.com/v1/chat/completions",
        "moonshot": "https://api.moonshot.cn/v1/chat/completions",
        "qianwen": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions",
        "zhipu": "https://open.bigmodel.cn/api/paas/v4/chat/completions",
        "ollama": "http://localhost:11434/v1/chat/completions",
    }
    return url.get((provider or "").lower(), "https://api.openai.com/v1/chat/completions")

SYSTEM_PROMPT = (
    "你是一个 Excel 自动化助手。根据用户指令，输出一个 JSON 数组，每个元素是一个操作对象 {op,...}。"
    "支持操作:\n"
    ' {"op":"write","sheet":"Sheet1","cell":"A1","value":123}\n'
    ' {"op":"write_range","sheet":"Sheet1","range":"A1:B2","values":[[1,2],[3,4]]}\n'
    ' {"op":"clear_range","sheet":"Sheet1","range":"A1:C5"}\n'
    ' {"op":"fill_down","sheet":"Sheet1","cell":"B2","count":10}\n'
    ' {"op":"ensure_sheet","sheet":"汇总"}\n'
    ' {"op":"set_cell_style","sheet":"Sheet1","cell":"A1","bold":true,"fill":"FFC000","font_color":"000000","align":"center"}\n'
    ' {"op":"noop","reason":"指令无需写入Excel"}\n'
    "要求：只返回 JSON 数组本身，不返回任何解释文字。若指令与Excel无关，返回 [{\"op\":\"noop\"}]。"
)

def call_llm(actions_prompt, task_model=None):
    cfg = llm_config() or {}
    # 任务级覆盖：task_model = "provider::model"；兼容只写 "model" 和 dict 对象
    _override_provider = None
    _override_model = None
    if task_model:
        if isinstance(task_model, dict):
            _override_provider = (task_model.get("provider") or "").strip().lower() or None
            _override_model = (task_model.get("model") or "").strip() or None
        elif isinstance(task_model, str):
            tm = task_model.strip()
            if "::" in tm:
                p, m = tm.split("::", 1)
                _override_provider = p.strip().lower() or None
                _override_model = m.strip() or None
            elif tm:
                _override_model = tm
    provider = (_override_provider or (cfg.get("provider") or cfg.get("apiType") or "") or "openai").lower()
    model = _override_model or cfg.get("model") or "gpt-4o"
    api_url = cfg.get("apiUrl") or cfg.get("baseUrl") or ""
    if not api_url:
        api_url = default_api_url(provider)
    api_key = cfg.get("apiKey") or ""
    # 没有任何配置时，若有任务模型但没有全局 llm_config → 仍提示缺失（避免拿 api_key 缺失的配置静默请求失败）
    if not cfg and not (provider and model and api_url):
        return False, "未找到 llm_config.json，请先在 Excel 中保存一次 AI 配置"
    if provider != "ollama" and not api_key:
        # Ollama 本地不需要 Key；其他 Provider 若没有 Key 仍尝试请求（有些自建代理无 Key），失败时给出提示
        pass
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": actions_prompt},
        ],
        "temperature": 0.1,
        "stream": False,
    }
    headers = {"Content-Type": "application/json"}
    if provider != "ollama" and api_key:
        headers["Authorization"] = "Bearer " + api_key
    req = Request(api_url, data=json.dumps(payload).encode("utf-8"), headers=headers, method="POST")
    try:
        with urlopen(req, timeout=120) as resp:
            body = resp.read().decode("utf-8")
    except Exception as e:
        return False, "LLM 请求失败: %s" % e
    try:
        obj = json.loads(body)
        content = None
        # OpenAI 兼容格式 choices[0].message.content
        if "choices" in obj and obj["choices"]:
            content = obj["choices"][0].get("message", {}).get("content")
        # Ollama 原生 /api/chat 格式 message.content
        if content is None and "message" in obj:
            content = obj["message"].get("content")
        if content is None:
            content = obj.get("response")
        if content is None:
            return False, "LLM 响应缺少内容: " + body[:300]
    except Exception:
        return False, "LLM 响应解析失败: " + body[:300]
    content = content.strip()
    if content.startswith("```"):
        lines = content.splitlines()
        content = "\n".join(lines[1:-1]).strip() if lines[-1].strip().startswith("```") else "\n".join(lines[1:]).strip()
    start = content.find("[")
    end = content.rfind("]")
    if start < 0 or end < 0:
        return False, "LLM 未返回 JSON 操作数组: " + content[:200]
    try:
        actions = json.loads(content[start:end + 1])
    except Exception as e:
        return False, "解析操作 JSON 失败: " + str(e)
    return True, actions

def ensure_wb(path):
    if os.path.isfile(path):
        return openpyxl.load_workbook(path)
    wb = openpyxl.Workbook()
    wb.active.title = "Sheet1"
    return wb

def _iter_cells(ws, ref):
    """按引用取单个或范围内所有 cell 迭代器；AI 常写 cell=A1:E1 表示整行。"""
    if not ref:
        return iter([])
    obj = ws[ref]
    if isinstance(obj, (tuple, list)):
        for row in obj:
            if isinstance(row, (tuple, list)):
                for c in row:
                    yield c
            else:
                yield row
    else:
        yield obj

def _merge_font(old, *, bold=None, color=None):
    """合并字体属性，避免 set bold 后再写 color 把 bold 冲掉。"""
    kwargs = {}
    try:
        if old and old.name: kwargs["name"] = old.name
        if old and old.size is not None: kwargs["size"] = old.size
        if old and old.bold is not None: kwargs["bold"] = old.bold
        if old and old.italic is not None: kwargs["italic"] = old.italic
        if old and old.underline is not None: kwargs["underline"] = old.underline
        if old and getattr(old, "strike", None) is not None: kwargs["strike"] = old.strike
    except Exception:
        pass
    if bold is not None:
        kwargs["bold"] = bool(bold)
    if color:
        kwargs["color"] = color
    return openpyxl.styles.Font(**kwargs)

def _fmt_val(v):
    """操作值展示：长文本截断，避免步骤卡片被超长内容撑爆。"""
    try:
        if isinstance(v, str):
            return v if len(v) <= 30 else v[:27] + "..."
        return repr(v)
    except Exception:
        return "?"

def _op_desc(op, a, ws_name):
    """把操作对象翻译成人类可读的步骤描述。"""
    try:
        if op == "write":
            return "写入 %s!%s = %s" % (ws_name, a.get("cell") or "A1", _fmt_val(a.get("value")))
        if op == "write_range":
            return "写入范围 %s!%s" % (ws_name, a.get("range") or "A1:A1")
        if op == "clear_range":
            return "清空范围 %s!%s" % (ws_name, a.get("range") or "A1:A1")
        if op == "fill_down":
            return "向下填充 %s!%s 共 %s 行" % (ws_name, a.get("cell") or "A1", a.get("count", 1))
        if op == "ensure_sheet":
            return "创建工作表「%s」" % (a.get("sheet") or "?")
        if op == "set_cell_style":
            attrs = []
            if a.get("bold") is not None or a.get("font_color"):
                attrs.append("字体")
            if a.get("fill"):
                attrs.append("填充" + str(a.get("fill")))
            if a.get("align"):
                attrs.append("对齐")
            if a.get("border") is not False:
                attrs.append("边框")
            if a.get("number_format"):
                attrs.append("数字格式")
            return "设置样式 %s!%s（%s）" % (ws_name, a.get("cell") or a.get("range") or "?", "、".join(attrs) or "默认")
        if op == "adjust_width":
            return "调整列宽 %s!%s = %s" % (ws_name, a.get("columns") or a.get("column") or a.get("range") or "?", a.get("width", 20))
        if op == "set_row_height":
            return "设置第 %s 行行高 = %s" % (a.get("row", 1), a.get("height", 24))
        if op == "set_border_range":
            return "为范围 %s!%s 添加边框" % (ws_name, a.get("range") or a.get("cell") or "A1:A1")
        return "执行操作 %s" % op
    except Exception:
        return "执行操作 %s" % op

def apply_actions(wb, actions):
    count = 0
    steps = []
    for a in actions or []:
        if not isinstance(a, dict):
            continue
        op = a.get("op")
        if op == "noop":
            continue
        sheet_name = a.get("sheet")
        try:
            ws = wb[sheet_name] if sheet_name and sheet_name in wb.sheetnames else wb.active
        except Exception:
            ws = wb.active
        ws_name = ws.title
        target = ("%s!%s" % (ws_name, a.get("cell") or a.get("range") or a.get("column") or "")).rstrip("!")
        desc = _op_desc(op, a, ws_name)
        ok_flag = True
        detail = ""
        try:
            if op == "write":
                ws[a["cell"]] = a.get("value")
                count += 1
            elif op == "write_range":
                rng = a.get("range") or "A1:A1"
                vals = a.get("values") or []
                idx = 0
                for row in ws[rng]:
                    for cell in row:
                        if idx < len(vals) and isinstance(vals[idx], list) and cell.column <= len(vals[idx]):
                            cell.value = vals[idx][cell.column - 1]
                        else:
                            break
                        idx += 1
                count += 1
            elif op == "clear_range":
                for row in ws[a.get("range") or "A1:A1"]:
                    for cell in row:
                        cell.value = None
                count += 1
            elif op == "fill_down":
                cell_ref = a.get("cell") or "A1"
                v = ws[cell_ref].value
                start_row = ws[cell_ref].row
                col = ws[cell_ref].column
                for i in range(1, int(a.get("count", 1)) + 1):
                    ws.cell(start_row + i, col, v)
                count += 1
            elif op == "ensure_sheet":
                if a.get("sheet") not in wb.sheetnames:
                    wb.create_sheet(a["sheet"])
                count += 1
            elif op == "set_cell_style":
                ref = a.get("cell") or a.get("range")
                processed = False
                for c in _iter_cells(ws, ref):
                    processed = True
                    try:
                        if a.get("bold") is not None or a.get("font_color"):
                            c.font = _merge_font(c.font, bold=a.get("bold"), color=a.get("font_color"))
                        if a.get("fill"):
                            c.fill = openpyxl.styles.PatternFill("solid", fgColor=str(a["fill"]).lstrip("#"))
                        if a.get("align"):
                            halign = a["align"] if a["align"] in ("center", "left", "right", "general") else "center"
                            valign = "center" if a.get("valign") in (None, True, "center") else (a.get("valign") or "center")
                            c.alignment = openpyxl.styles.Alignment(horizontal=halign, vertical=valign, wrap_text=bool(a.get("wrap")))
                        if a.get("border") is not False:
                            sides = {}
                            want_thin = False
                            if a.get("border") in (True, 1, "thin"):
                                want_thin = True
                            for sd in ("left", "right", "top", "bottom"):
                                val = a.get(sd) or ("thin" if want_thin else None)
                                if val:
                                    sides[sd] = openpyxl.styles.Side(style=val, color=a.get("border_color") or "000000")
                            if sides:
                                for sd in ("left", "right", "top", "bottom"):
                                    if want_thin and sd not in sides:
                                        sides[sd] = openpyxl.styles.Side(style="thin", color=a.get("border_color") or "000000")
                                c.border = openpyxl.styles.Border(**sides)
                        if a.get("number_format"):
                            c.number_format = str(a["number_format"])
                    except Exception:
                        pass
                if processed:
                    count += 1
                else:
                    ok_flag = False
                    detail = "未定位到目标单元格 %s" % (a.get("cell") or a.get("range") or "?")
            elif op == "adjust_width":
                ref = a.get("cell") or a.get("range") or a.get("column")
                cols_done = set()
                for c in _iter_cells(ws, ref):
                    letter = openpyxl.utils.get_column_letter(c.column)
                    if letter in cols_done:
                        continue
                    cols_done.add(letter)
                    try:
                        ws.column_dimensions[letter].width = float(a.get("width", 20))
                    except Exception:
                        pass
                cols = a.get("columns") or []
                if isinstance(cols, str):
                    cols = [x for x in cols.replace("，", ",").split(",") if x]
                for col in cols:
                    try:
                        letter = col.strip().upper()
                        if letter and letter not in cols_done:
                            cols_done.add(letter)
                            ws.column_dimensions[letter].width = float(a.get("width", 20))
                    except Exception:
                        pass
            elif op == "set_row_height":
                r = int(a.get("row") or 1)
                ws.row_dimensions[r].height = float(a.get("height", 24))
                count += 1
            elif op == "set_border_range":
                rng = a.get("range") or a.get("cell") or "A1:A1"
                style = a.get("style") or "thin"
                color = a.get("color") or "000000"
                side = openpyxl.styles.Side(style=style, color=color)
                border = openpyxl.styles.Border(left=side, right=side, top=side, bottom=side)
                for row in ws[rng]:
                    for cell in row:
                        try:
                            cell.border = border
                        except Exception:
                            pass
                count += 1
            else:
                ok_flag = False
                detail = "未知操作类型: %s" % op
        except Exception as e:
            ok_flag = False
            detail = str(e)
        step = {"op": op, "target": target, "desc": desc, "ok": ok_flag}
        if detail:
            step["detail"] = detail
        steps.append(step)
    return True, "已执行 %d 个操作" % count, count, steps

def _search_file(name):
    """按文件名在本机常见目录快速找完整路径。找不到返回 []。"""
    if not name:
        return []
    # 去掉可能的路径前缀，只按文件名的 basename 匹配
    needle = os.path.basename(name.strip())
    # 先看 basename 本身是不是绝对路径且存在（用户可能直接贴了路径过来）
    if os.path.isabs(name) and os.path.isfile(name):
        return [name]
    candidates = []
    # 用户目录下常见几个高频位置 —— 遍历但设超时 + 跳过深树
    home = os.path.expanduser("~")
    roots = []
    for sub in ["Desktop", "Documents", "Downloads", "OneDrive", "OneDrive\\Desktop",
                "OneDrive\\Documents", "OneDrive\\Downloads"]:
        p = os.path.join(home, sub)
        if os.path.isdir(p):
            roots.append(p)
    # 再加当前工作目录和 BASE_DIR 兜底
    for p in [os.getcwd(), BASE_DIR]:
        if os.path.isdir(p) and p not in roots:
            roots.append(p)

    import time as _time
    deadline = _time.time() + 5.0   # 5 秒硬超时，避免卡死在大目录
    for root in roots:
        if _time.time() > deadline:
            break
        try:
            for dirpath, dirnames, filenames in os.walk(root):
                if _time.time() > deadline:
                    break
                # 跳过 node_modules / .git / 回收站这些大目录
                dirnames[:] = [d for d in dirnames if d not in (
                    "node_modules", ".git", "$RECYCLE.BIN", "System Volume Information",
                    "__pycache__", ".vs", ".vscode", ".idea")]
                for fn in filenames:
                    if fn == needle:
                        full = os.path.join(dirpath, fn)
                        candidates.append(full)
                        # 找到一个就够，多了也没用 —— 后面候选会被截断
                        if len(candidates) >= 5:
                            break
                if len(candidates) >= 5:
                    break
        except (PermissionError, OSError):
            continue
    return candidates

def execute_command(file_path, command, task_model=None):
    steps = []

    def _fail(msg):
        return False, msg, steps

    if not file_path:
        return _fail("未指定目标文件")
    if not os.path.isfile(file_path):
        return _fail("目标文件不存在: " + file_path)
    if not command:
        return _fail("指令为空")

    # 不再做自定义 msvcrt 字节锁预检——Windows 的字节区间锁 + Excel 乐观写极易误报"被占用"，
    # 而且与真实 openpyxl write 的句柄竞争完全不同。让真实打开/保存决定成败，并在这里快速验证一次可写。
    try:
        _probe = open(file_path, 'a+b')
        _probe.close()
    except (PermissionError, OSError):
        return _fail("文件被占用，无法写入（请关闭该文件的 Excel 后重试）: " + file_path)

    # 步骤 1：调用 LLM 生成操作计划
    cfg = llm_config() or {}
    model_label = task_model or cfg.get("model") or "默认模型"
    prompt = "目标文件: %s\n用户指令: %s\n请输出操作数组。" % (file_path, command)
    ok, res = call_llm(prompt, task_model=task_model)
    plan_step = {"op": "llm_plan", "target": file_path,
                 "desc": "调用模型（%s）分析指令、生成操作计划" % model_label, "ok": ok}
    if not ok:
        plan_step["detail"] = str(res)[:200]
    steps.append(plan_step)
    if not ok:
        return _fail(res)
    wb = ensure_wb(file_path)
    try:
        ok2, msg, cnt, op_steps = apply_actions(wb, res if isinstance(res, list) else [])
        steps.extend(op_steps)
        if not ok2:
            return False, msg, steps
        # 保存：最多重试 3 次，间隔 0.4s（Excel 关闭释放锁通常 <1s，避免因释放抖动报错）
        last_err = None
        for _i in range(3):
            try:
                wb.save(file_path)
                last_err = None
                break
            except PermissionError as e:
                last_err = "文件被占用，无法保存（请关闭该文件的 Excel）: " + file_path
                time.sleep(0.4)
            except Exception as e:
                last_err = "保存失败: " + str(e)
                break
        if last_err:
            steps.append({"op": "save", "target": file_path,
                          "desc": "保存文件 %s" % os.path.basename(file_path), "ok": False, "detail": last_err})
            return False, last_err, steps
        steps.append({"op": "save", "target": file_path,
                      "desc": "保存文件 %s（写入磁盘）" % os.path.basename(file_path), "ok": True})
        return True, "已写入 %s（%d 个操作）" % (file_path, cnt), steps
    finally:
        # 无论成功还是异常，都显式关闭工作簿，释放 openpyxl 持有的文件句柄，避免常驻进程内残留锁
        try:
            wb.close()
        except Exception:
            pass

# 与 MobileControlBridge.IsTrustedOrigin 对齐：无 Origin（curl / 同源）才用 *；
# 带 Origin 的网页只有本机与宿主虚拟主机可读响应，避免任意网站读走 API 密钥。
_TRUSTED_ORIGIN_HOSTS = {
    "localhost",
    "127.0.0.1",
    "::1",
    "app.excelai.local",
    "screenassets.local",
}


def _cors_allow_origin(headers):
    origin = ""
    try:
        origin = (headers.get("Origin") or "").strip()
    except Exception:
        origin = ""
    if not origin:
        return "*"
    try:
        host = (urlparse(origin).hostname or "").lower()
    except Exception:
        return ""
    if host.startswith("[") and host.endswith("]"):
        host = host[1:-1]
    if host in _TRUSTED_ORIGIN_HOSTS:
        return origin
    return ""


def _public_llm_config(cfg):
    """给浏览器的配置视图：不回传明文 API 密钥。"""
    if not isinstance(cfg, dict):
        return {}
    key = str(cfg.get("apiKey") or "")
    configured = bool(key) and key != "local"
    masked = ""
    if configured:
        masked = "****" if len(key) <= 8 else (key[:3] + "****" + key[-4:])
    provider = cfg.get("provider") or ""
    return {
        "provider": provider,
        "apiType": cfg.get("apiType") or provider,
        "model": cfg.get("model") or "",
        "apiUrl": cfg.get("apiUrl") or "",
        "baseUrl": cfg.get("baseUrl") or cfg.get("apiUrl") or "",
        "apiKeyConfigured": configured,
        "apiKeyMasked": masked,
    }


def _key_should_keep_existing(api_key):
    text = (api_key or "").strip()
    return (not text) or ("****" in text)


def restrict_config_acl(path):
    """把 llm_config.json 收紧为当前用户可读，失败不阻断保存。"""
    if os.name != "nt" or not path or not os.path.isfile(path):
        return
    user = os.environ.get("USERNAME") or ""
    domain = os.environ.get("USERDOMAIN") or ""
    if not user:
        return
    ident = (domain + "\\" + user) if domain and domain.upper() != user.upper() else user
    try:
        subprocess.run(
            ["icacls", path, "/inheritance:r", "/grant:r", ident + ":(F)"],
            capture_output=True,
            timeout=8,
            check=False,
        )
    except Exception:
        pass


class Handler(BaseHTTPRequestHandler):
    def _apply_cors(self):
        origin = _cors_allow_origin(self.headers)
        if origin:
            self.send_header("Access-Control-Allow-Origin", origin)
            self.send_header("Vary", "Origin")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _json(self, code, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self._apply_cors()
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self._apply_cors()
        self.end_headers()

    def _read_body(self):
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return None
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return None

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        q = parse_qs(parsed.query)
        # 独立页面托管：Excel 关闭后仍可访问任务规划页（单文件，无需其它静态资源）
        if path in ("/", "/index", "/index.html"):
            # 根路径统一重定向到任务规划页，方便 Excel 关闭后手动访问 localhost:9529
            self.send_response(302)
            self.send_header("Location", "/task-planner")
            self.send_header("Content-Length", "0")
            self._apply_cors()
            self.end_headers()
            return
        if path in ("/task-planner", "/task-planner.html"):
            tp_candidates = [
                os.path.abspath(os.path.join(BASE_DIR, "..", "UI", "task-planner.html")),
                os.path.abspath(os.path.join(BASE_DIR, "task-planner.html")),
            ]
            tp_file = next((p for p in tp_candidates if os.path.isfile(p)), None)
            if tp_file is None:
                self._json(404, {"success": False, "message": "task-planner.html 未找到，请先同步 UI 目录"})
                return
            try:
                with open(tp_file, "r", encoding="utf-8") as f:
                    html = f.read()
                data = html.encode("utf-8")
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(data)))
                self.send_header("Cache-Control", "no-cache")
                self._apply_cors()
                self.end_headers()
                self.wfile.write(data)
            except Exception as e:
                self._json(500, {"success": False, "message": "读取页面失败：" + str(e)})
            return
        if path == "/health":
            cfg_ok = os.path.isfile(CONFIG_PATH) if CONFIG_PATH else False
            self._json(200, {"ok": True, "service": "standalone-agent", "port": PORT,
                             "config_found": cfg_ok, "tasks": len(_tasks)})
        elif path == "/tasks":
            with _task_lock:
                self._json(200, {"success": True, "data": _tasks})
        elif path == "/model-config":
            # 返回 provider/model/地址，以及密钥是否已配置。不回传明文 apiKey。
            cfg = llm_config() or {}
            if cfg:
                self._json(200, {"success": True, "data": _public_llm_config(cfg)})
            else:
                self._json(200, {"success": False, "message": "未找到 llm_config.json，请在 Excel 设置 → AI 模型 → 保存后重启独立服务"})
        elif path == "/search-file":
            # 按文件名在本机常见目录查完整路径（选一个文件但沙箱拿不到路径时用）
            name = (q.get("name") or [""])[0].strip()
            if not name:
                self._json(400, {"success": False, "message": "缺少 name 参数"})
                return
            self._json(200, {"success": True, "matches": _search_file(name)})
        elif path == "/workbook-info":
            # 独立页「引用区」用：传入本地 xlsx 绝对路径，返回 sheet 清单、行列数、表头预览
            fp = (q.get("path") or [""])[0].strip()
            if not fp:
                self._json(400, {"success": False, "message": "缺少 path 参数"})
                return
            try:
                if not os.path.isfile(fp):
                    self._json(404, {"success": False, "message": "文件不存在"})
                    return
                wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
                try:
                    sheets_info = []
                    for name in wb.sheetnames:
                        try:
                            ws = wb[name]
                            max_row = ws.max_row or 0
                            max_col = ws.max_column or 0
                            header_row = []
                            if max_row >= 1:
                                for col in range(1, max_col + 1):
                                    try:
                                        v = ws.cell(row=1, column=col).value
                                    except Exception:
                                        v = None
                                    header_row.append("" if v is None else str(v))
                            sheets_info.append({
                                "name": name,
                                "rows": max_row,
                                "cols": max_col,
                                "headerRow": header_row,
                            })
                        except Exception:
                            sheets_info.append({
                                "name": name,
                                "rows": 0, "cols": 0, "headerRow": []
                            })
                    self._json(200, {
                        "success": True,
                        "path": os.path.abspath(fp),
                        "fileName": os.path.basename(fp),
                        "sheets": sheets_info,
                    })
                finally:
                    try: wb.close()
                    except Exception: pass
            except Exception as e:
                self._json(500, {"success": False, "message": "读取失败：" + str(e)})
        else:
            self._json(404, {"success": False, "message": "not found"})

    def do_POST(self):
        path = urlparse(self.path).path
        body = self._read_body()
        if path == "/tasks":
            with _deleted_lock:
                new = dict(body or {})
                task_id = new.get("Id")
                # 已删除墓碑：前端 localStorage 补同步把已删任务回推时拒绝（410），前端据此自愈清除本地残留
                if task_id and task_id in _deleted_ids:
                    self._json(410, {"success": False, "message": "任务已删除（%s），不再接收回推" % task_id, "deleted": True})
                    return
            with _task_lock:
                new = dict(body or {})
                task_id = new.get("Id")
                existing = next((t for t in _tasks if t.get("Id") == task_id), None)
                if existing is not None:
                    existing.update(new)
                    task = existing
                else:
                    task = new
                task.setdefault("Id", "tsk_" + str(int(time.time())))
                task["NextRun"] = _next_run(task).strftime("%Y-%m-%d %H:%M:%S")
                _tasks[:] = [t for t in _tasks if t.get("Id") != task.get("Id")]
                _tasks.append(task)
                _save_tasks()
            self._json(200, {"success": True, "data": task})
        elif path.startswith("/tasks/") and path.endswith("/execute"):
            tid = path.split("/")[2]
            with _task_lock:
                task = next((t for t in _tasks if t.get("Id") == tid), None)
            if not task:
                self._json(404, {"success": False, "message": "任务不存在"})
                return
            file_path = task.get("TargetFile") or task.get("target_file") or ""
            command = task.get("Command") or ""
            tModel = task.get("Model") or task.get("model") or None
            with _task_lock:
                task["Running"] = True
                _save_tasks()
            try:
                ok, msg, steps = execute_command(file_path, command, task_model=tModel)
            except Exception as e:
                ok, msg, steps = False, "执行异常：" + str(e), []
            with _task_lock:
                task["Running"] = False
                task["LastRun"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                task["LastResult"] = msg
                task["LastSteps"] = steps
                task["UpdatedAt"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                task["NextRun"] = _next_run(task).strftime("%Y-%m-%d %H:%M:%S")
                _save_tasks()
            self._json(200, {"success": ok, "message": msg, "result": msg, "steps": steps})
        elif path == "/shutdown":
            with _task_lock:
                _save_tasks()
            self._json(200, {"ok": True, "message": "独立服务已停止"})
            # 响应发送后延迟退出进程
            threading.Timer(0.3, os._exit, args=(0,)).start()
        elif path == "/reload-config":
            global CONFIG_PATH
            CONFIG_PATH = find_config()
            found = bool(CONFIG_PATH and os.path.isfile(CONFIG_PATH))
            self._json(200, {
                "success": found,
                "message": "配置已重载" if found else "配置未找到，可直接在本页填写模型配置并保存（无需 Excel）",
                "configPath": CONFIG_PATH or "",
                "configFound": found
            })
        elif path == "/save-config":
            # 独立页面直接写入配置到 SSOT（无需 Excel 运行），反向同步到 Excel 下次启动时会自动导入
            try:
                cfg = dict(body or {})
            except Exception:
                cfg = {}
            provider = (cfg.get("provider") or cfg.get("apiType") or "").strip()
            model = (cfg.get("model") or "").strip()
            api_key = (cfg.get("apiKey") or cfg.get("api_key") or "").strip()
            api_url = (cfg.get("apiUrl") or cfg.get("api_url") or cfg.get("baseUrl") or "").strip()
            if not provider:
                self._json(400, {"success": False, "message": "缺少必填参数：provider（服务商）"})
                return
            if _key_should_keep_existing(api_key):
                existing = llm_config() or {}
                existing_provider = (existing.get("provider") or existing.get("apiType") or "").strip().lower()
                existing_key = (existing.get("apiKey") or "").strip()
                if existing_provider == provider.lower() and existing_key and existing_key != "local" and "****" not in existing_key:
                    api_key = existing_key
            if provider.lower() != "ollama" and not api_key:
                # 云端服务商必须有 API Key；留空且本机已有同服务商密钥时上面已沿用。Ollama 本地允许为空。
                self._json(400, {"success": False, "message": "云端服务商需提供 apiKey；本地 Ollama 可留空"})
                return
            if not model:
                self._json(400, {"success": False, "message": "缺少必填参数：model（模型名称）"})
                return
            if provider.lower() == "ollama" and not api_key:
                api_key = "local"
            payload = {
                "provider": provider,
                "apiType": provider,
                "model": model,
                "apiUrl": api_url,
                "baseUrl": api_url,
                "apiKey": api_key,
            }
            target_path = ssot_write_path()
            try:
                with open(target_path, "w", encoding="utf-8") as f:
                    json.dump(payload, f, ensure_ascii=False, indent=2)
                restrict_config_acl(target_path)
            except Exception as e:
                self._json(500, {"success": False, "message": "写入失败：" + str(e), "path": target_path})
                return

            # 🆕 多路径双写：与 Excel 端 ExportLLMConfigForStandalone 保持完全相同的 SSOT 覆盖策略
            #   除了 AppData 权威路径，还要同步写到 MobileControlBridge 启动时 --config 必传的
            #   bin\Debug|Release\Config\llm_config.json；否则独立页里保存后，Excel 重启/新起的
            #   Python 进程通过 --config 读的是旧文件，导致"配置改了但没生效"的幻觉。
            written = [os.path.abspath(target_path)]
            # A) <BASE_DIR>/../Config/llm_config.json —— 永远存在：
            #      在输出目录 BASE_DIR=bin\Debug\StandaloneAgent 时，对应 bin\Debug\Config（--config 指向的就是这里）
            #      在源码树 BASE_DIR=StandaloneAgent 时，对应根级 Config（开发模式兜底）
            runtime_cfg = os.path.abspath(os.path.join(BASE_DIR, "..", "Config", "llm_config.json"))
            secondary_targets = [runtime_cfg]
            # B) "切换构建配置"同步：当前输出目录若是 bin\Debug\StandaloneAgent，把 bin\Release\Config 也写一份（反之亦然）
            #    与 VB 端 ExportLLMConfigForStandalone 的 ③ 构建目录兄弟同步保持一致
            try:
                # BASE_DIR = <root>\bin\Debug\StandaloneAgent  →  dirname twice = <root>\bin （这就是 bin 根）
                bin_root = os.path.dirname(os.path.dirname(os.path.abspath(BASE_DIR)))
                for cfg_name in ("Debug", "Release"):
                    sibling_cfg_dir = os.path.join(bin_root, cfg_name, "Config")
                    if os.path.isdir(sibling_cfg_dir):
                        sibling_path = os.path.join(sibling_cfg_dir, "llm_config.json")
                        norm_sib = os.path.normcase(os.path.abspath(sibling_path))
                        already = any(os.path.normcase(w) == norm_sib for w in written)
                        if not already:
                            secondary_targets.append(sibling_path)
            except Exception:
                pass
            # C) 源码树副本（BASE_DIR=StandaloneAgent 时才需要；如果 BASE_DIR 已经在 bin\ 下，
            #    ../../bin/Debug|Release 会退化成 bin\bin\Debug|Release（冗余），故跳过）
            base_norm = os.path.normcase(os.path.abspath(BASE_DIR))
            in_output_tree = False
            for seg in ("\\bin\\debug\\", "\\bin\\release\\"):
                if seg in base_norm:
                    in_output_tree = True
                    break
            if not in_output_tree:
                secondary_targets.append(os.path.abspath(os.path.join(BASE_DIR, "..", "..", "bin", "Debug", "Config", "llm_config.json")))
                secondary_targets.append(os.path.abspath(os.path.join(BASE_DIR, "..", "..", "bin", "Release", "Config", "llm_config.json")))
            for sp in secondary_targets:
                try:
                    norm_sp = os.path.normcase(os.path.abspath(sp))
                    already = any(os.path.normcase(w) == norm_sp for w in written)
                    if already:
                        continue
                    sd = os.path.dirname(sp)
                    if not os.path.isdir(sd):
                        try:
                            os.makedirs(sd, exist_ok=True)
                        except Exception:
                            continue  # 目录建不出来跳过（Debug 目录可能不存在）
                    with open(sp, "w", encoding="utf-8") as f:
                        json.dump(payload, f, ensure_ascii=False, indent=2)
                    restrict_config_acl(sp)
                    written.append(os.path.abspath(sp))
                except Exception:
                    pass  # 任一从路径失败不影响主路径成功
            # 更新全局 CONFIG_PATH 以便下次 llm_config() 立即读到新文件
            CONFIG_PATH = target_path
            # 清理历史遗留副本（保持与 ExportLLMConfigForStandalone 一致）
            obsolete = [
                os.path.join(BASE_DIR, "llm_config.json"),
            ]
            for p in obsolete:
                try:
                    ap = os.path.normcase(os.path.abspath(p))
                    is_protected = any(os.path.normcase(w) == ap for w in written)
                    if os.path.isfile(ap) and not is_protected:
                        os.remove(ap)
                except Exception:
                    pass
            self._json(200, {
                "success": True,
                "message": "配置已保存到权威位置（下次启动 Excel 时会自动同步），已双写到 " + str(len(written)) + " 处",
                "path": target_path,
                "allPaths": written,
                "data": _public_llm_config(payload),
            })
        else:
            self._json(404, {"success": False, "message": "not found"})

    def do_DELETE(self):
        path = urlparse(self.path).path
        if path.startswith("/tasks/"):
            tid = path.split("/")[2]
            with _task_lock:
                before = len(_tasks)
                _tasks[:] = [t for t in _tasks if t.get("Id") != tid]
                _save_tasks()
            with _deleted_lock:
                _deleted_ids.add(tid)
                _save_deleted()
            self._json(200, {"success": True, "removed": before - len(_tasks)})
        else:
            self._json(404, {"success": False, "message": "not found"})

    def log_message(self, *a):
        pass

def scheduler_loop():
    while True:
        try:
            now = datetime.datetime.now()
            # 只在外层锁内收集到期任务；执行放到锁外，避免嵌套加锁死锁
            # （threading.Lock 不可重入，任务到点时会第二次 with _task_lock → 永久死锁）
            due = []
            with _task_lock:
                for t in _tasks:
                    if not t.get("Enabled", True):
                        continue
                    freq = t.get("Frequency") or "daily"
                    if freq == "once" and t.get("LastRun"):
                        continue
                    nr = t.get("NextRun")
                    try:
                        nr_dt = datetime.datetime.strptime(nr, "%Y-%m-%d %H:%M:%S") if nr else None
                    except Exception:
                        nr_dt = None
                    if nr_dt is None:
                        nr_dt = _next_run(t)
                        t["NextRun"] = nr_dt.strftime("%Y-%m-%d %H:%M:%S")
                    if now >= nr_dt:
                        due.append(t)
            for t in due:
                file_path = t.get("TargetFile") or t.get("target_file") or ""
                command = t.get("Command") or ""
                # 执行中标记：页面轮询时显示"正在执行"
                with _task_lock:
                    t["Running"] = True
                    _save_tasks()
                try:
                    ok, msg, steps = execute_command(file_path, command)
                except Exception as e:
                    ok, msg, steps = False, "执行异常：" + str(e), []
                with _task_lock:
                    t["Running"] = False
                    t["LastRun"] = now.strftime("%Y-%m-%d %H:%M:%S")
                    t["LastResult"] = msg
                    t["LastSteps"] = steps
                    t["UpdatedAt"] = now.strftime("%Y-%m-%d %H:%M:%S")
                    t["NextRun"] = _next_run(t).strftime("%Y-%m-%d %H:%M:%S")
                    _save_tasks()
                print("[Scheduler] run", t.get("Name"), ok, msg)
        except Exception as e:
            print("[Scheduler] err", e)
        time.sleep(20)

def _respawn_detached_if_excel_spawned():
    """若由 EXCEL.EXE 直接拉起，则以脱离进程组的方式重生为独立进程并退出本实例，
    使 9529 服务脱离 Excel 进程树——Excel 正常退出/崩溃/被结束进程树都不再影响它。"""
    import subprocess
    if '--smartpilot-detached' in sys.argv:
        return False  # 已是脱离后的副本，不再重生
    if not sys.platform.startswith('win'):
        return False
    try:
        import ctypes
        from ctypes import wintypes
        ppid = os.getppid()
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, ppid)
        if not h:
            return False
        try:
            buf = ctypes.create_unicode_buffer(512)
            size = wintypes.DWORD(512)
            if not ctypes.windll.kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                return False
            is_excel = os.path.basename(buf.value).lower() == 'excel.exe'
        finally:
            ctypes.windll.kernel32.CloseHandle(h)
        if not is_excel:
            return False
    except Exception:
        return False
    # 两层脱离策略：
    # 1) CREATE_BREAKAWAY_FROM_JOB | CREATE_NEW_PROCESS_GROUP | DETACHED_PROCESS：
    #    Excel 常被放入 Job Object（作业对象），直接 spawn 的子进程会继承该作业，
    #    Excel 退出时作业关闭会连坐杀掉 9529。携带 BREAKAWAY 标志可跳出作业。
    # 2) 若作业不允许脱离（CreateProcess 报 ERROR_ACCESS_DENIED），退化为仅脱离控制台/进程组。
    script = os.path.abspath(__file__)
    new_args = [sys.executable, script] + sys.argv[1:] + ['--smartpilot-detached']
    CREATE_BREAKAWAY_FROM_JOB = 0x01000000
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    DETACHED_PROCESS = 0x00000008
    for flags in (
        CREATE_BREAKAWAY_FROM_JOB | CREATE_NEW_PROCESS_GROUP | DETACHED_PROCESS,
        CREATE_NEW_PROCESS_GROUP | DETACHED_PROCESS,
    ):
        try:
            # std 全部 DEVNULL：脱离进程后不再依赖 Excel 的句柄，Excel 退出后 print 不会因句柄失效抛错
            subprocess.Popen(new_args, creationflags=flags, close_fds=True, cwd=BASE_DIR,
                             stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            return True
        except Exception:
            continue
    return False

def main():
    global CONFIG_PATH
    if _respawn_detached_if_excel_spawned():
        print("[StandaloneAgent] 已脱离 Excel 进程树，以独立进程继续运行")
        return
    if not CONFIG_PATH:
        print("[StandaloneAgent] WARNING: 未找到 llm_config.json（请先在 Excel 保存一次 AI 配置）")
    _load_tasks()
    _load_deleted()
    threading.Thread(target=scheduler_loop, daemon=True).start()
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    print("[StandaloneAgent] 已启动于 http://localhost:%d  任务数=%d  配置=%s" % (PORT, len(_tasks), CONFIG_PATH))
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()

if __name__ == "__main__":
    main()