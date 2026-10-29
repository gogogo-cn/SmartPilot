#!/usr/bin/env python
# -*- coding: utf-8 -*-
# pyright: reportConstantRedefinition=none, reportPossiblyUnboundVariable=none, reportCallIssue=none, reportGeneralTypeIssues=none, reportMissingImports=none, reportIncompatibleMethodOverride=none, reportUnusedImport=none, reportUnknownParameterType=none, reportMissingParameterType=none, reportUnknownVariableType=none, reportUnknownArgumentType=none, reportUnknownMemberType=none, reportAny=none, reportImplicitStringConcatenation=none, reportUnusedParameter=none, reportUnusedVariable=none, reportRedeclaration=none, reportArgumentType=none, reportOptionalMemberAccess=none, reportUnusedFunction=none, reportUnnecessaryComparison=none, reportUnusedCallResult=none, reportImplicitOverride=none
# ============================================================
# 【py-excel-bridge.py】本地 Excel 代理服务（Python + pywin32）
# ------------------------------------------------------------
# 替代 Tools/local-bridge-server.js 里的 winax 驱动：
#   winax 3.6.9 在 Node 20+ 无法编译（V8 API 不兼容），导致网页版
#   本地桥只能跑“模拟模式”，碰不到真实 Excel。
#   pywin32 走 COM 但与 Python 版本无关，稳定可靠。
#
# 与 local-bridge-server.js 保持同一 HTTP 契约（前端 excel-connector.js 零改动）：
#   GET  /api/excel/status
#   POST /api/excel/connect     body: {action:'test'}
#   POST /api/excel/invoke      body: {method, arguments:[...]}
#   其余 GET 请求静态托管 ../UI（即 http://localhost:9527/index.html）
#
# 用法：
#   python py-excel-bridge.py            # 默认 9527
#   PORT=9531 python py-excel-bridge.py  # 自定义端口
# 依赖：pywin32（Excel 自动化），标准库 http.server（无需额外安装）。
# ============================================================
import os
import sys
import io
import json
import uuid
import time
import base64
import threading
import traceback
import re
import subprocess
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer, ThreadingHTTPServer
from urllib.parse import urlparse, unquote

# pythonw（无控制台）下 sys.stdout/stderr 为 None，任何 print 都会 AttributeError 直接崩。
# 重定向到日志文件，让无控制台启动器（SmartPilot.vbs → pythonw）可用。
if sys.stdout is None:
    sys.stdout = open(os.path.join(tempfile.gettempdir(), "smartpilot-bridge-stdout.log"),
                      "a", encoding="utf-8", buffering=1)
if sys.stderr is None:
    sys.stderr = open(os.path.join(tempfile.gettempdir(), "smartpilot-bridge-stderr.log"),
                      "a", encoding="utf-8", buffering=1)

try:
    import pythoncom
    import win32com.client
    PYWIN32 = True
    PYWIN32_ERR = ""
except Exception as e:  # pragma: no cover
    PYWIN32 = False
    PYWIN32_ERR = str(e)

PORT = int(os.environ.get("PORT", "9527"))
HERE = os.path.dirname(os.path.abspath(__file__))
UI_DIR = os.path.normpath(os.path.join(HERE, "..", "UI"))
# 静默模式：--quiet 抑制每请求访问日志（仍输出错误与启动横幅），桌面版用
QUIET = ("--quiet" in sys.argv)

APP = None                      # Excel COM Application 单例
LOCK = threading.Lock()


def com_init():
    try:
        pythoncom.CoInitialize()
    except Exception:
        pass


def to_jsonable(v):
    """COM 变体 → JSON 可序列化类型。"""
    if v is None:
        return None
    t = type(v)
    if t is tuple:
        return [to_jsonable(x) for x in v]
    if t is list:
        return [to_jsonable(x) for x in v]
    if t in (int, float, str, bool):
        return v
    try:
        return str(v)
    except Exception:
        return None


# pythonw（无控制台）下，console 子进程默认会新建控制台窗口 → 黑窗反复闪烁。
# 所有子进程调用必须带 CREATE_NO_WINDOW（GUI 目标程序不受影响，仍正常显示窗口）。
_CREATE_NO_WINDOW = 0x08000000  # CREATE_NO_WINDOW


def _excel_process_running():
    """检测系统中是否有 EXCEL.EXE 进程在运行（不依赖 COM/ROT）。"""
    try:
        import subprocess
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE"],
            capture_output=True, text=True, timeout=5,
            creationflags=subprocess.CREATE_NO_WINDOW
        ).stdout or ""
        return "EXCEL.EXE" in out.upper()
    except Exception:
        return False


EXCEL_APP_CLSID = "00024500-0000-0000-c000-000000000046"
_EXCEL_EXT = (".xls", ".xlsx", ".xlsm", ".xlsb")


def _enum_excel_instances():
    """枚举 ROT 中所有正在运行的 Excel.Application 实例。

    两条通道（实践中缺一不可）：
    1. "!{clsid}" 项名字对象 → Application 对象。但 Excel 有时只为
       部分/自动化实例注册该条目，双击打开的实例可能缺席；
    2. 工作簿文件路径 moniker（如 C:\\...\\demo_test_data.xlsx）
       → Workbook 对象 → .Application。双击打开的 Excel 一定会注册。
    最后按窗口句柄去重。
    """
    if not PYWIN32:
        return []
    com_init()
    insts = []
    seen_hwnds = set()

    def _add(app):
        try:
            hwnd = int(app.Hwnd)
        except Exception:
            hwnd = None
        if hwnd and hwnd in seen_hwnds:
            return
        insts.append(app)
        if hwnd:
            seen_hwnds.add(hwnd)

    try:
        ctx = pythoncom.CreateBindCtx(0)
        rot = pythoncom.GetRunningObjectTable()
        for mk in rot.EnumRunning():
            try:
                name = mk.GetDisplayName(ctx, None)
            except Exception:
                continue
            if not name:
                continue
            low = name.lower()
            try:
                if low == "!{" + EXCEL_APP_CLSID + "}":
                    # 通道1：Application 项名字对象
                    unk = rot.GetObject(mk)
                    _add(win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch)))
                elif low.endswith(_EXCEL_EXT) and ("\\" in name or "/" in name):
                    # 通道2：工作簿文件 moniker → Workbook.Application
                    unk = rot.GetObject(mk)
                    wb = win32com.client.Dispatch(unk.QueryInterface(pythoncom.IID_IDispatch))
                    _add(wb.Application)
            except Exception:
                continue
    except Exception:
        pass
    return insts


def _score_excel_instance(app):
    """实例优先级：有工作簿 > 工作簿有已保存文件路径（含 \\）> 数量多。"""
    count = 0
    saved = 0
    try:
        count = app.Workbooks.Count
    except Exception:
        count = 0
    for i in range(1, count + 1):
        try:
            fn = str(app.Workbooks.Item(i).FullName or "")
            if "\\" in fn or "/" in fn:
                saved += 1
        except Exception:
            pass
    return (count > 0, saved, count)


def get_excel(create=False):
    """获取运行的 Excel 实例；create=True 时若不存在则新建。

    多实例场景下（桌面常有桥误建的白板 Excel 和用户真实 Excel 并存），
    通过 ROT 枚举所有实例，优先选择"带已保存文件工作簿"的那个，
    避免连到不含用户数据的幽灵实例。
    """
    global APP
    if APP is not None:
        try:
            _ = APP.Version
            # 缓存实例还有工作簿 → 直接复用；空实例则重新挑选
            if APP.Workbooks.Count > 0:
                return APP
        except Exception:
            APP = None
    if not PYWIN32:
        raise RuntimeError("pywin32 不可用: " + PYWIN32_ERR)

    # 1) 首选：枚举全部实例，挑最优（优先含已保存工作簿的）
    insts = _enum_excel_instances()
    if insts:
        best = max(insts, key=_score_excel_instance)
        APP = best
        return APP

    # 枚举不到任何 ROT 实例：若缓存 APP 仍存活（可能是空实例），继续用
    if APP is not None:
        try:
            _ = APP.Version
            return APP
        except Exception:
            APP = None

    # 2) 枚举失败时退回 GetActiveObject（保守路径）
    try:
        inst = win32com.client.GetActiveObject("Excel.Application")
        APP = inst
        return APP
    except Exception:
        inst = None

    if inst is None:
        excel_proc = _excel_process_running()
        if excel_proc:
            raise RuntimeError(
                "检测到 Excel 正在运行，但无法通过 COM 连接。"
                "常见原因：Excel 与本地桥权限不一致（Excel 以管理员身份运行，而本地桥不是，或反之）。"
                "请关闭 Excel 后，用与本地桥相同的权限重新打开 Excel 和 py-excel-bridge.py。"
            )
        if not create:
            raise RuntimeError("Excel 未运行，请先打开 Excel")
        inst = win32com.client.Dispatch("Excel.Application")
        inst.Visible = True
    APP = inst
    return APP


def ensure_workbook(app):
    """确保存在一个工作簿，返回 ActiveWorkbook。"""
    wb = app.ActiveWorkbook
    if wb is None:
        wb = app.Workbooks.Add()
    return wb


# ===================== 方法实现 =====================

def m_get_selected_range(args):
    app = get_excel(create=True)
    wb = ensure_workbook(app)
    sel = app.Selection
    try:
        addr = sel.Address
        vals = to_jsonable(sel.Value2)
        cols = int(sel.Columns.Count)
        rows = int(sel.Rows.Count)
    except Exception:
        # 选区可能指向整个 Application（无工作簿时），退回 A1
        app.ActiveSheet.Range("A1").Select()
        sel = app.Selection
        addr = sel.Address
        vals = to_jsonable(sel.Value2)
        cols = int(sel.Columns.Count)
        rows = int(sel.Rows.Count)
    return {"address": addr, "values": vals, "columnCount": cols, "rowCount": rows}


def _selection_payload(app):
    """读取当前选区摘要（地址/工作表/工作簿），供 GetSelection 及前端轮询使用。"""
    wb = app.ActiveWorkbook
    wb_name = str(wb.Name) if wb else ""
    try:
        sel = app.Selection
        addr = str(sel.Address)
    except Exception:
        return {"address": "", "sheetName": "", "workbookName": wb_name}
    sheet_name = ""
    try:
        sheet_name = str(sel.Worksheet.Name)
    except Exception:
        try:
            sheet_name = str(app.ActiveSheet.Name)
        except Exception:
            pass
    return {"address": addr, "sheetName": sheet_name, "workbookName": wb_name}


def m_get_selection(args):
    """对齐 VSTO 桥：返回当前选区 {address, sheetName, workbookName}。"""
    app = get_excel(create=False)
    return _selection_payload(app)


def m_get_multi_selection(args):
    """对齐 VSTO 桥：返回多选区 {areas: [{address, sheetName, workbookName}], count}。"""
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    wb_name = str(wb.Name) if wb else ""
    areas = []
    try:
        sel = app.Selection
        cnt = int(sel.Areas.Count)
        for i in range(1, cnt + 1):
            a = sel.Areas.Item(i)
            try:
                sn = str(a.Worksheet.Name)
            except Exception:
                sn = ""
            areas.append({"address": str(a.Address), "sheetName": sn, "workbookName": wb_name})
    except Exception as e:
        return {"areas": areas, "count": len(areas), "error": str(e)}
    return {"areas": areas, "count": len(areas)}


def m_start_selection_monitor(args):
    # 本地桥（HTTP）没有 VSTO 那样的脚本事件推送通道：
    # 前端 reference-area-manager 在本地桥模式下改为轮询 GetSelection 模拟 SelectionChange。
    return {"success": True, "started": True, "mode": "client-poll"}


def m_stop_selection_monitor(args):
    return {"success": True, "stopped": True}


def m_get_range_data(args):
    """对齐 VSTO 桥 GetRangeData：读取区域数据。
    地址支持 [工作簿]表!A1:C5 / '表 名'!A1:C5 / 表!A1:C5 / A1:C5。
    返回 {success, address, sheetName, workbookName, headers, rows, rowCount, colCount}。
    """
    addr = args[0] if args else None
    if not addr:
        raise RuntimeError("GetRangeData 缺少地址参数")
    app = get_excel(create=False)
    target_wb = app.ActiveWorkbook
    s = str(addr).strip()
    wb_name = ""
    sheet_name = ""
    range_part = s
    # 工作簿：以 [名称] 开头
    if range_part.startswith("["):
        end = range_part.find("]")
        if end > 0:
            wb_name = range_part[1:end]
            range_part = range_part[end + 1:].strip()
    # 工作表：! 之前部分（可能带引号）
    if "!" in range_part:
        sheet_part, range_part = range_part.split("!", 1)
        sheet_name = sheet_part.strip().strip("'")
    if not range_part:
        raise RuntimeError("GetRangeData 地址无效: " + str(addr))
    if wb_name:
        try:
            target_wb = app.Workbooks(wb_name)
        except Exception:
            pass
    try:
        ws = target_wb.Sheets(sheet_name) if sheet_name else target_wb.ActiveSheet
    except Exception:
        raise RuntimeError("GetRangeData 找不到工作表: " + sheet_name)
    rng = ws.Range(range_part)
    vals = to_jsonable(rng.Value2)
    # 规整为二维 rows
    if vals is None:
        rows = [[]]
    elif not isinstance(vals, list):
        rows = [[vals]]
    elif vals and not isinstance(vals[0], list):
        rows = [vals]
    else:
        rows = vals
    headers = [("" if h is None else str(h)) for h in (rows[0] if rows and rows[0] else [])]
    data_rows = rows[1:] if len(rows) > 1 else []
    return {
        "success": True,
        "address": str(rng.Address),
        "sheetName": str(ws.Name),
        "workbookName": str(target_wb.Name),
        "headers": headers,
        "rows": data_rows,
        "rowCount": len(data_rows),
        "colCount": (len(headers) if headers else (len(rows[0]) if rows and rows[0] else 0)),
    }


def m_write_to_range(args):
    rng = args[0] if len(args) > 0 else None
    data = args[1] if len(args) > 1 else None
    if not rng:
        raise RuntimeError("WriteToRange 缺少 rangeAddress 参数")
    if isinstance(data, str):
        data = json.loads(data)
    if isinstance(data, dict):
        data = data.get("values", data)
    if data is None:
        raise RuntimeError("WriteToRange 缺少数据")
    # 规整为二维列表
    if isinstance(data, list) and data and not isinstance(data[0], list):
        data = [data]
    app = get_excel(create=True)
    ensure_workbook(app)
    r = app.Range(rng)
    r.Value2 = data
    return {"success": True, "address": rng, "message": "数据已写入 " + str(rng)}


def m_get_workbook_info(args):
    app = get_excel()
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("当前没有打开的工作簿")
    sheets = [wb.Sheets(i).Name for i in range(1, wb.Sheets.Count + 1)]
    return {
        "name": wb.Name,
        "fullName": wb.FullName,
        "sheets": sheets,
        "activeSheet": wb.ActiveSheet.Name,
        "path": wb.Path,
    }


def m_get_active_sheet_name(args):
    app = get_excel()
    return app.ActiveSheet.Name


def m_get_open_workbooks(args):
    app = get_excel()
    workbooks = []
    for i in range(1, app.Workbooks.Count + 1):
        wb = app.Workbooks(i)
        sheets = [wb.Sheets(j).Name for j in range(1, wb.Sheets.Count + 1)]
        workbooks.append({
            "name": wb.Name,
            "fullName": wb.FullName,
            "path": wb.Path,
            "sheets": sheets,
            "activeSheet": wb.ActiveSheet.Name,
        })
    return {"workbooks": workbooks, "count": len(workbooks)}


def m_activate_workbook(args):
    name = args[0] if args else None
    if not name:
        raise RuntimeError("ActivateWorkbook 缺少工作簿名称")
    app = get_excel()
    target = None
    for i in range(1, app.Workbooks.Count + 1):
        wb = app.Workbooks(i)
        if wb.Name == name or wb.FullName == name:
            target = wb
            break
    if target is None:
        raise RuntimeError("找不到工作簿: " + str(name))
    target.Activate()
    sheets = [target.Sheets(j).Name for j in range(1, target.Sheets.Count + 1)]
    return {
        "success": True,
        "message": "已切换到工作簿: " + target.Name,
        "data": {
            "name": target.Name,
            "fullName": target.FullName,
            "sheets": sheets,
            "activeSheet": target.ActiveSheet.Name,
            "path": target.Path,
        },
    }


def _csv_field(v, delimiter):
    s = "" if v is None else str(v)
    if delimiter in s or "," in s or '"' in s or "\n" in s:
        s = '"' + s.replace('"', '""') + '"'
    return s


def m_export_csv(args):
    sheet_name = args[0] if len(args) > 0 else ""
    file_path = args[1] if len(args) > 1 else None
    delimiter = args[2] if len(args) > 2 and args[2] else ","
    include_header = args[3] if len(args) > 3 else True
    if not file_path:
        raise RuntimeError("ExportToCsv 缺少 filePath")
    # 【防假成功】展开开头的 ~（Unix 风格 home 路径），否则 open("~/...") 会按相对路径
    # 处理：要么 FileNotFoundError，要么在 CWD 下造出字面名为 "~" 的目录。
    file_path = _os.path.expanduser(file_path)
    file_path = _os.path.abspath(file_path)
    app = get_excel()
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("当前没有打开的工作簿")
    ws = wb.Sheets(sheet_name) if sheet_name else app.ActiveSheet
    import csv
    used = ws.UsedRange
    data = to_jsonable(used.Value2)
    rows = data if isinstance(data, list) and data and isinstance(data[0], list) else [data] if data else []
    with open(file_path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f, delimiter=delimiter)
        started = False
        for idx, row in enumerate(rows):
            if not include_header and idx == 0:
                continue
            w.writerow([("" if c is None else c) for c in (row if isinstance(row, list) else [row])])
    return {"taskCompleted": True, "message": "成功导出到 CSV: " + file_path}


def m_execute_vba(args):
    code = args[0] if args else None
    if not code:
        raise RuntimeError("ExecuteVBA 缺少代码")
    app = get_excel(create=True)
    # 必须是宏启用工作簿
    wb = app.ActiveWorkbook
    if wb is None or not str(wb.Name).lower().endswith(".xlsm"):
        wb = app.Workbooks.Add()
        wb.SaveAs(os.path.join(os.environ.get("USERPROFILE", "."), "SmartPilotBridge.xlsm"),
                  FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled = 52
    module_name = "SmartPilotBridge"
    vbcs = wb.VBProject.VBComponents
    target = None
    for i in range(1, vbcs.Count + 1):
        if vbcs.Item(i).Name == module_name:
            target = vbcs.Item(i)
            break
    if target is None:
        target = vbcs.Add(1)  # vbext_ct_StdModule
        target.Name = module_name
    target.CodeModule.DeleteLines(1, target.CodeModule.CountOfLines)
    target.CodeModule.AddFromString(code)
    # 尝试运行找到的第一个 Sub（无参）
    ran = None
    import re
    m = re.search(r"Sub\s+(\w+)\s*\(", code)
    if m:
        sub_name = m.group(1)
        try:
            app.Run(sub_name)
            ran = sub_name
        except Exception as e:
            ran = sub_name + " (未自动运行: " + str(e) + ")"
    return {
        "success": True,
        "message": "VBA 已写入模块 " + module_name + (("，已运行 " + ran) if ran else "，请手动运行"),
        "moduleName": module_name,
        "ran": ran,
    }


def m_check_python_environment(args):
    """代码编辑器的 Python 环境检测。桥本身就是 Python 进程，环境必然可用。"""
    return {"available": True, "version": "Python " + ".".join(map(str, sys.version_info[:3]))}


def m_execute_python(args):
    """执行编辑器提交的 Python 代码：临时文件 + 子进程，捕获 stdout/stderr/退出码。
    与 ExecuteVBA 同级（编辑器"运行"按钮），仅在本机执行，超时 60s。"""
    code = args[0] if args else ""
    if not str(code).strip():
        return {"success": False, "error": "代码为空"}
    fd, path = tempfile.mkstemp(suffix=".py", prefix="smartpilot_exec_")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(code)
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        proc = subprocess.run([sys.executable, path], capture_output=True,
                              timeout=60, env=env,
                              creationflags=subprocess.CREATE_NO_WINDOW)
        return {
            "success": proc.returncode == 0,
            "exitCode": proc.returncode,
            "stdout": proc.stdout.decode("utf-8", errors="replace"),
            "stderr": proc.stderr.decode("utf-8", errors="replace"),
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "error": "执行超时（60 秒）"}
    except Exception as e:
        return {"success": False, "error": str(e)}
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# ==================== 云端 OCR 配置（与插件版 CloudOcrConfig.json 归一） ====================
# 与插件版共用 %APPDATA%\ExcelAIVSTOAddIn\cloud_ocr_config.json，
# 密钥用 DPAPI(CurrentUser) 加密，熵 = SHA256("ExcelAIVSTOAddIn|v2.0|"+MachineName)，
# 与 Utils/EncryptionHelper.vb 完全互通（同机同用户可互解）。

import hashlib
import win32crypt  # pywin32；桥自身就依赖 pywin32 驱动 Excel COM


def _dpapi(data, entropy=None, protect=True):
    """DPAPI 加解密。pywin32 封装（ctypes 手写 DATA_BLOB 有段错误风险，勿改回）。"""
    if protect:
        return win32crypt.CryptProtectData(bytes(data), None, entropy, None, None, 0)
    return win32crypt.CryptUnprotectData(bytes(data), entropy, None, None, 0)[1]


def _enc_entropy():
    raw = ("ExcelAIVSTOAddIn|v2.0|" + (os.environ.get("COMPUTERNAME") or "")).encode("utf-8")
    return hashlib.sha256(raw).digest()


def _enc_is_encrypted(text):
    """对齐 EncryptionHelper.IsEncrypted：有效 base64 且 ≥32 字节、非可打印字符 >20%。"""
    try:
        b = base64.b64decode(text)
    except Exception:
        return False
    if len(b) < 32:
        return False
    non_printable = sum(1 for x in b[:256] if x < 32 or x > 126)
    return non_printable > min(len(b), 256) * 0.2


def _enc_encrypt(plain):
    return base64.b64encode(_dpapi(plain.encode("utf-8"), _enc_entropy(), True)).decode("ascii")


def _enc_decrypt(cipher):
    if not cipher:
        return ""
    if not _enc_is_encrypted(cipher):
        return cipher  # 明文（桌面版桥写入的旧数据）原样返回
    try:
        return _dpapi(base64.b64decode(cipher), _enc_entropy(), False).decode("utf-8", errors="replace")
    except Exception:
        return ""  # 对齐 VB Decrypt 的容错：解密失败返回空


_CLOUD_OCR_KEYS = ("baiduApiKey", "baiduSecretKey", "tencentSecretId", "tencentSecretKey")


def _cloud_ocr_cfg_path():
    return os.path.join(os.environ.get("APPDATA", "."), "ExcelAIVSTOAddIn", "cloud_ocr_config.json")


def _load_cloud_ocr_cfg():
    cfg = {k: "" for k in _CLOUD_OCR_KEYS}
    cfg["provider"] = "baidu"
    try:
        with open(_cloud_ocr_cfg_path(), "r", encoding="utf-8-sig") as f:
            raw = json.load(f)
        if isinstance(raw, dict):
            for k in list(cfg.keys()) + list(_CLOUD_OCR_KEYS):
                if raw.get(k) is not None:
                    cfg[k] = str(raw[k])
    except Exception:
        pass
    for k in _CLOUD_OCR_KEYS:
        cfg[k] = _enc_decrypt(cfg[k])
    return cfg


def _save_cloud_ocr_cfg(cfg):
    d = {k: str(cfg.get(k) or "") for k in list(_CLOUD_OCR_KEYS) + ["provider"]}
    for k in _CLOUD_OCR_KEYS:
        if d[k] and not _enc_is_encrypted(d[k]):
            d[k] = _enc_encrypt(d[k])
    p = _cloud_ocr_cfg_path()
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    return p


def m_get_cloud_ocr_config(args):
    """前端读 config.provider/apiKey/secretKey/secretId（app-ocr.js loadOcrSettings）。"""
    cfg = _load_cloud_ocr_cfg()
    return {"success": True, "provider": cfg.get("provider") or "baidu",
            "apiKey": cfg.get("baiduApiKey") or "", "secretKey": cfg.get("baiduSecretKey") or "",
            "secretId": cfg.get("tencentSecretId") or "", "tencentSecretKey": cfg.get("tencentSecretKey") or ""}


def m_save_cloud_ocr_config(args):
    try:
        data = json.loads(args[0]) if args and isinstance(args[0], str) else (args[0] if args else {})
    except Exception as e:
        return {"success": False, "error": "JSON 解析失败: " + str(e)}
    cfg = _load_cloud_ocr_cfg()
    if data.get("provider"):
        cfg["provider"] = str(data["provider"])
    if data.get("apiKey") is not None:
        cfg["baiduApiKey"] = str(data["apiKey"])
    if data.get("secretKey") is not None:
        cfg["baiduSecretKey"] = str(data["secretKey"])
    if data.get("secretId") is not None:
        cfg["tencentSecretId"] = str(data["secretId"])
    if data.get("tencentSecretKey") is not None:
        cfg["tencentSecretKey"] = str(data["tencentSecretKey"])
    _save_cloud_ocr_cfg(cfg)
    return {"success": True}


def m_test_cloud_ocr_config(args):
    """连通性测试：baidu 用 OAuth 换 access_token 验证凭据；tencent 签名复杂，提示保存后在插件版验证。"""
    try:
        data = json.loads(args[1]) if len(args) > 1 and isinstance(args[1], str) else {}
    except Exception:
        data = {}
    provider = str(args[0] or data.get("provider") or "baidu").lower()
    if provider == "baidu":
        ak = str(data.get("apiKey") or "") or _load_cloud_ocr_cfg().get("baiduApiKey") or ""
        sk = str(data.get("secretKey") or "") or _load_cloud_ocr_cfg().get("baiduSecretKey") or ""
        if not ak or not sk:
            return {"success": False, "error": "请先填写 API Key 和 Secret Key"}
        try:
            import urllib.parse
            import urllib.request
            url = ("https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&"
                   + urllib.parse.quote(ak) + "&" + urllib.parse.quote(sk))
            req = urllib.request.Request(url, headers={"User-Agent": "SmartPilotWeb"})
            with urllib.request.urlopen(req, timeout=10) as r:
                resp = json.loads(r.read().decode("utf-8", errors="replace"))
            if resp.get("access_token"):
                return {"success": True}
            return {"success": False, "error": resp.get("error_description") or resp.get("error") or "凭据验证失败"}
        except Exception as e:
            return {"success": False, "error": "网络请求失败: " + str(e)}
    return {"success": False, "error": "桌面版暂不支持该引擎的连通性测试，保存后请在 Excel 插件版中验证"}


def m_detect_environment(args):
    info = {"driver": "pywin32", "pywin32Available": PYWIN32}
    if PYWIN32:
        try:
            app = get_excel(create=False)
            info["excelVersion"] = str(app.Version)
            info["excelConnected"] = True
        except Exception as e:
            info["excelConnected"] = False
            info["reason"] = str(e)
    else:
        info["excelConnected"] = False
        info["reason"] = PYWIN32_ERR
    return info


# ---- 归一：多端共用同一份数据源（插件版已装则用它，否则用桌面版自带）----
_LOCALAPPDATA = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
_APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
_PLUGIN_INSTALL = os.path.join(_LOCALAPPDATA, "Programs", "ExcelAIVSTOAddIn")
_PLUGIN_APPDATA = os.path.join(_APPDATA, "ExcelAIVSTOAddIn")


def _resolve_shared_dir(candidates):
    """返回第一个存在的目录；都不存在时返回最后一个候选（桌面版自带路径）。"""
    for d in candidates:
        if d and os.path.isdir(d):
            return d
    return candidates[-1] if candidates else ""


SKILLS_DIR = _resolve_shared_dir([
    os.path.join(_PLUGIN_INSTALL, "Excel", "Skills"),          # 插件版安装目录
    os.path.join(_PLUGIN_APPDATA, "Skills"),                   # 插件版 AppData 部署
    os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Skills')),
])
_SKILLS_CACHE = None
_SKILL_APPDATA = os.path.join(_APPDATA, "ExcelAIVSTOAddIn")
_CUSTOM_SKILLS_DIR = os.path.join(_SKILL_APPDATA, "CustomSkills", "skills")
_CUSTOM_SNIPPETS_DIR = os.path.join(_SKILL_APPDATA, "CustomSkills", "snippets")
_DISABLED_SKILLS_FILE = os.path.join(_SKILL_APPDATA, "disabled_skills.json")

try:
    import yaml as _yaml
except Exception:  # pragma: no cover
    _yaml = None


def _parse_skill_md(fpath):
    """解析 Skill MD 的 YAML frontmatter，返回 dict；无 frontmatter 返回 None。"""
    try:
        with open(fpath, "r", encoding="utf-8-sig") as f:
            text = f.read()
    except Exception:
        return None
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return None
    fm_lines = []
    for line in lines[1:]:
        if line.strip() == "---":
            break
        fm_lines.append(line)
    else:
        return None
    if _yaml is None or not fm_lines:
        return None
    try:
        data = _yaml.safe_load("\n".join(fm_lines))
    except Exception:
        return None
    return data if isinstance(data, dict) else None


def _load_disabled_skills():
    """与 VSTO 端共用 %APPDATA%\ExcelAIVSTOAddIn\disabled_skills.json（JSON 数组）。"""
    try:
        with open(_DISABLED_SKILLS_FILE, "r", encoding="utf-8-sig") as f:
            arr = json.loads(f.read() or "[]")
        return set(str(x).strip() for x in arr if str(x).strip())
    except Exception:
        return set()


def _save_disabled_skills(names):
    try:
        d = os.path.dirname(_DISABLED_SKILLS_FILE)
        if d and not os.path.isdir(d):
            os.makedirs(d, exist_ok=True)
        with open(_DISABLED_SKILLS_FILE, "w", encoding="utf-8") as f:
            json.dump(sorted(set(names)), f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def _safe_skill_filename(name):
    safe = re.sub(r'[\\/:*?"<>|]+', "_", str(name or "").strip())
    return safe or "skill"


def m_get_all_skills(args):
    """对齐 VSTO SkillBridge.GetAllSkills 契约：{success, skills:[camelCase字段]}。
    数据源归一：内置 Skills/*.md + 自定义 CustomSkills/skills/*.md（自定义覆盖同名），
    并应用与插件版共用的 disabled_skills.json 停用名单。"""
    global _SKILLS_CACHE
    if _SKILLS_CACHE is not None:
        return {"success": True, "skills": _SKILLS_CACHE}

    merged = {}

    def _collect(directory, is_custom):
        if not directory or not os.path.isdir(directory):
            return
        for fn in sorted(os.listdir(directory)):
            if not fn.lower().endswith(".md"):
                continue
            data = _parse_skill_md(os.path.join(directory, fn))
            if not data or not data.get("name"):
                continue
            name = str(data.get("name"))
            existing = merged.get(name)
            if is_custom or existing is None:
                data.setdefault("isCustom", is_custom)
                data.setdefault("isEnabled", True)
                merged[name] = data

    _collect(SKILLS_DIR, False)
    _collect(_CUSTOM_SKILLS_DIR, True)

    disabled = _load_disabled_skills()
    skills = []
    for name in sorted(merged.keys()):
        s = merged[name]
        s["isEnabled"] = name not in disabled
        skills.append(s)
    _SKILLS_CACHE = skills
    return {"success": True, "skills": _SKILLS_CACHE}


def m_toggle_skill_enabled(args):
    """ToggleSkillEnabled(name, isEnabled)：固化到与插件版共用的停用名单。"""
    name = args[0] if args else None
    enabled = args[1] if len(args) > 1 else True
    if not name:
        return {"success": False, "error": "缺少技能名称"}
    name = str(name).strip()
    disabled = _load_disabled_skills()
    if enabled:
        disabled.discard(name)
    else:
        disabled.add(name)
    if not _save_disabled_skills(disabled):
        return {"success": False, "error": "写入停用名单失败: " + _DISABLED_SKILLS_FILE}
    global _SKILLS_CACHE
    _SKILLS_CACHE = None
    return {"success": True, "name": name, "isEnabled": bool(enabled)}


def m_save_skill(args):
    """SaveSkill(skillJson)：保存/更新自定义技能到 CustomSkills/skills/<name>.md。"""
    global _SKILLS_CACHE
    raw = args[0] if args else None
    if isinstance(raw, str):
        try:
            data = json.loads(raw)
        except Exception as e:
            return {"success": False, "error": "技能 JSON 解析失败: " + str(e)}
    elif isinstance(raw, dict):
        data = raw
    else:
        return {"success": False, "error": "缺少技能数据"}
    name = str(data.get("name") or data.get("Name") or "").strip()
    if not name:
        return {"success": False, "error": "缺少 name 字段"}
    body = (data.get("content") or data.get("body") or data.get("templateCode")
            or data.get("llmSystemPrompt") or "")
    try:
        import yaml as _y
        front = {k: v for k, v in data.items() if k not in ("content", "body")}
        front.setdefault("isCustom", True)
        front.setdefault("isEnabled", True)
        text = "---\n" + _y.safe_dump(front, allow_unicode=True, sort_keys=False) + "---\n\n" + str(body)
    except Exception as e:
        return {"success": False, "error": "生成技能文件失败（需要 PyYAML）: " + str(e)}
    try:
        if not os.path.isdir(_CUSTOM_SKILLS_DIR):
            os.makedirs(_CUSTOM_SKILLS_DIR, exist_ok=True)
        path = os.path.join(_CUSTOM_SKILLS_DIR, _safe_skill_filename(name) + ".md")
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
    except Exception as e:
        return {"success": False, "error": "写入技能文件失败: " + str(e)}
    _SKILLS_CACHE = None
    return {"success": True, "name": name, "path": path}


def m_delete_skill(args):
    """DeleteSkill(name)：删除自定义技能（内置 Skill 不可删，与 VSTO 语义一致）。"""
    global _SKILLS_CACHE
    name = args[0] if args else None
    if not name:
        return {"success": False, "error": "缺少技能名称"}
    name = str(name).strip()
    builtin = os.path.join(SKILLS_DIR, _safe_skill_filename(name) + ".md")
    if os.path.isfile(builtin):
        return {"success": False, "error": "内置 Skill 不可删除"}
    removed = False
    for directory in (_CUSTOM_SKILLS_DIR, _CUSTOM_SNIPPETS_DIR):
        p = os.path.join(directory, _safe_skill_filename(name) + ".md")
        if os.path.isfile(p):
            os.remove(p)
            removed = True
    disabled = _load_disabled_skills()
    disabled.discard(name)
    _save_disabled_skills(disabled)
    _SKILLS_CACHE = None
    if not removed:
        return {"success": False, "error": "未找到自定义 Skill: " + name}
    return {"success": True, "name": name}


def m_get_skill_by_name(args):
    """GetSkillByName(name)：按 name 在 Skills/*.md 与自定义 Skills 中查找，
    返回前端 editSkill 期望的 {success, skill}（Web API 路径直接消费）。"""
    name = args[0] if args else None
    if not name:
        return {"success": False, "error": "缺少技能名称"}
    name = str(name).strip()

    def _find(directory):
        if not directory or not os.path.isdir(directory):
            return None
        for fn in os.listdir(directory):
            if not fn.lower().endswith(".md"):
                continue
            data = _parse_skill_md(os.path.join(directory, fn))
            if data and str(data.get("name")) == name:
                return data
        return None

    skill = _find(_CUSTOM_SKILLS_DIR) or _find(SKILLS_DIR)
    if not skill:
        return {"success": False, "error": "Skill 不存在: " + name}
    return {"success": True, "skill": skill}


def m_get_available_skills(args):
    """对齐 VSTO SkillBridge.GetAvailableSkills：只返回 isEnabled 的技能。"""
    all_skills = m_get_all_skills(args).get("skills", [])
    enabled = []
    for s in all_skills:
        if str(s.get("isEnabled", True)).lower() not in ("false", "0", "no"):
            enabled.append(s)
    return {"success": True, "skills": enabled}


def _skill_by_name(name):
    """按 name 或 displayName 在仓库 Skills/ 中查找技能定义。"""
    if not name:
        return None
    if os.path.isdir(SKILLS_DIR):
        for fn in sorted(os.listdir(SKILLS_DIR)):
            if not fn.lower().endswith(".md"):
                continue
            d = _parse_skill_md(os.path.join(SKILLS_DIR, fn))
            if d and (d.get("name") == name or d.get("displayName") == name):
                return d
    return None


def _ts():
    import datetime as _dt
    return _dt.datetime.now().strftime("%Y%m%d%H%M%S")


def _col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _lookup_path(left, ctx, results):
    """点路径取值：ctx.x 取上下文；step.field / step.result.field 取前序步骤结果。
    容忍 VSTO 风格的 'result' 包裹缺失（桥里部分工具结果扁平，无 result 子对象）。"""
    left = left.strip()
    if left.lower().startswith("ctx"):
        rest = left[3:].lstrip(".")
        base = ctx
    else:
        toks = left.split(".")
        base = results.get(toks[0], {})
        rest = ".".join(toks[1:])
    if not rest:
        return base
    cur = base
    for part in rest.split("."):
        if isinstance(cur, dict):
            if part in cur:
                cur = cur[part]
            elif part == "result":
                # 容忍缺失的 result 包裹，仍以当前字典继续向下解析
                continue
            else:
                return None
        elif isinstance(cur, list):
            try:
                cur = cur[int(part)]
            except Exception:
                return None
        else:
            return None
    return cur


def _eval_template_expr(expr, ctx, results):
    """解析 {{expr}}，支持 a || default 默认语法。"""
    if "||" in expr:
        left, default = expr.split("||", 1)
    else:
        left, default = expr, None
    left = left.strip()
    default = default.strip() if default is not None else None
    if default is not None and len(default) >= 2 and default[0] in "'\"" and default[-1] == default[0]:
        default = default[1:-1]
    val = _lookup_path(left, ctx, results)
    if val is None and default is not None:
        return default
    return val


def _resolve_text(s, ctx, results):
    # 整串即单个占位符时，保留原始类型（不强制转字符串）
    mfull = re.fullmatch(r"\s*\{\{\s*(.*?)\s*\}\}\s*", s)
    if mfull:
        return _eval_template_expr(mfull.group(1), ctx, results)

    def repl(mm):
        v = _eval_template_expr(mm.group(1), ctx, results)
        return "" if v is None else (v if isinstance(v, str) else json.dumps(v, ensure_ascii=False))

    return re.sub(r"\{\{\s*(.*?)\s*\}\}", repl, s)


def _deep_resolve(node, ctx, results):
    if isinstance(node, str):
        return _resolve_text(node, ctx, results)
    if isinstance(node, dict):
        return {k: _deep_resolve(v, ctx, results) for k, v in node.items()}
    if isinstance(node, list):
        return [_deep_resolve(x, ctx, results) for x in node]
    return node


# ── VSTO↔桥 工具契约兼容垫片 ──
# 仓库 Skills/*.md 的步骤是按 VSTO 工具 I/O 契约写的；下面把 VSTO 风格参数/返回
# 适配到 pywin32 桥里工具的实际契约，使同一份技能可在桥中真实执行。
def _shim_aggregate(params):
    out = dict(params)
    gb = params.get("groupBy")
    if gb and isinstance(gb, list) and gb:
        out["dimColumn"] = gb[0]
    funcs = params.get("funcs")
    if funcs and isinstance(funcs, list):
        f0 = funcs[0]
        out["valueColumn"] = f0.get("column") or out.get("valueColumn")
        out["aggFunc"] = f0.get("func") or "sum"
    # VSTO 契约用 sourceSheet/sourceRange，桥用 sheet/range
    if params.get("sourceSheet"):
        out["sheet"] = params["sourceSheet"]
    if params.get("sourceRange"):
        out["range"] = params["sourceRange"]
    out.pop("groupBy", None)
    out.pop("funcs", None)
    out.pop("createChart", None)
    out.pop("sourceSheet", None)
    out.pop("sourceRange", None)
    return out


def _shim_read_range(params):
    out = dict(params)
    if str(out.get("sheet", "")).strip().lower() in ("active", ""):
        out["sheet"] = ""
    if str(out.get("range", "")).strip().lower() in ("selection", "selected", ""):
        out["range"] = ""
    return out


def _shim_chart(params):
    out = dict(params)
    dv = params.get("dataValues")
    if isinstance(dv, list) and dv:
        name = "ChartData_" + _ts()
        route_method("excel_create_sheet", [{"name": name}])
        route_method("excel_write_range", [{"sheet": name, "range": "A1", "values": dv}])
        nrows = len(dv)
        ncols = max((len(r) for r in dv if isinstance(r, list)), default=0)
        out["dataRange"] = "%s!A1:%s%d" % (name, _col_letter(ncols), nrows)
    out.pop("dataValues", None)
    out.pop("anchorSheet", None)
    out.pop("anchorRange", None)
    out.pop("resultSheet", None)
    return out


def _shim_tool(tool, params):
    if tool == "excel_aggregate":
        return _shim_aggregate(params)
    if tool == "excel_read_range":
        return _shim_read_range(params)
    if tool == "excel_create_chart":
        return _shim_chart(params)
    return params


def _adapt_result(tool, res):
    if tool == "excel_aggregate" and isinstance(res, dict):
        rr = res.get("result") or {}
        res = dict(res)
        res["data"] = rr.get("rows")
        res["headers"] = rr.get("headers")
        res["resultSheet"] = None
    if tool == "excel_read_range" and isinstance(res, dict):
        # VSTO 的 read 结果用 sheet 字段，桥里为 worksheet，补别名以兼容 {{step.result.sheet}}
        res = dict(res)
        if "sheet" not in res and res.get("worksheet") is not None:
            res["sheet"] = res["worksheet"]
    return res


def _run_vba_code(code):
    return m_execute_vba([code])


def _run_steps(skill, steps, ctx):
    results = {}
    log = []
    for step in steps:
        sname = step.get("name") or step.get("tool")
        tool = step.get("tool")
        strategy = str(step.get("strategy") or "direct").lower()
        params = step.get("params") or {}
        try:
            if strategy == "code":
                res = _run_vba_code(step.get("code") or "")
            else:
                rparams = _deep_resolve(params, ctx, results)
                rparams = _shim_tool(tool, rparams)
                res = route_method(tool, [rparams])
                res = _adapt_result(tool, res)
            results[sname] = res
            log.append({"name": sname, "tool": tool, "success": True, "result": res})
        except Exception as e:
            log.append({"name": sname, "tool": tool, "success": False, "error": str(e)})
            return {"success": False, "skill": skill.get("name"),
                    "failedStep": sname, "error": str(e), "steps": log}
    return {"success": True, "skill": skill.get("name"), "steps": log}


def m_execute_skill(args):
    """ExecuteSkill：真实执行技能（对齐 VSTO 端语义）。
    读取仓库 Skills/*.md：
      - 含 steps：按声明式步骤顺序执行（direct 调真实工具 / code 跑 VBA），
        参数模板 {{ctx.x}} 取上下文、{{步骤名.result.y}} 取前序结果；
      - 含 templateCode：写入并运行 VBA；
      - 纯 AI/Python 类技能：返回 notExecutable 说明需 VSTO 宿主后端，
        不再像此前那样静默返回假成功。
    由于直接读取与 VSTO 同一份 Skills/*.md，技能定义天然同步。"""
    p = _tool_params(args) if not isinstance(args, list) else {}
    if isinstance(args, list):
        name = args[0] if args else None
        inputs = args[1] if len(args) > 1 else None
    else:
        name = p.get("skillName") or p.get("name")
        inputs = p.get("inputsJson") or p.get("inputs")
    if isinstance(inputs, str):
        try:
            inputs = json.loads(inputs)
        except Exception:
            inputs = {}
    if not isinstance(inputs, dict):
        inputs = {}
    skill = _skill_by_name(name)
    if skill is None:
        raise RuntimeError("未找到技能: " + str(name))
    steps = skill.get("steps") or []
    if steps:
        return _run_steps(skill, steps, inputs)
    code = skill.get("templateCode")
    if code:
        return {"success": True, "skill": skill.get("name"), "vba": _run_vba_code(code)}
    pm = str(skill.get("processingMode") or "")
    pl = str(skill.get("pipeline") or "")
    if pm in ("MultimodalLLM", "SkillSpec") or pl in ("html", "code"):
        return {"success": False, "notExecutable": True,
                "error": "技能 %s 依赖 AI/大模型后端(%s)，本地桥(pywin32)不支持直接执行，需在 VSTO 宿主内运行。"
                         % (name, pm or pl)}
    if "Python" in pm or pl == "python":
        return {"success": False, "notExecutable": True,
                "error": "技能 %s 为 Python 代码生成/执行类，本地桥未内置 Python 运行时，需在 VSTO 宿主内运行。"
                         % name}
    return {"success": False,
            "error": "技能 %s 未包含可执行的 steps 或 VBA 模板（processingMode=%s）" % (name, pm)}


def m_get_mcp_status(args):
    """桌面版：桥进程本身就是 MCP Server（HTTP /invoke 执行 + /api/mcp/tools 清单）。
    返回与 VSTO McpManager 相同结构的 status，前端面板据此显示「运行中」。"""
    try:
        tools = load_mcp_tools()
    except Exception:
        tools = []
    n = len(tools) if isinstance(tools, list) else 0
    return {
        "success": True,
        "status": {
            "serverRunning": True,
            "serverPort": PORT,
            "serverToolCount": n,
            "totalAvailableTools": n,
            "clients": []
        }
    }


def m_start_mcp_server(args):
    return {"success": True,
            "status": {"serverRunning": True, "serverPort": PORT},
            "message": "桌面版的 MCP Server 随本地桥自动运行，无需手动启动"}


def m_stop_mcp_server(args):
    return {"success": True,
            "warning": "桌面版的 MCP Server 随本地桥运行，关闭 SmartPilot 桌面版窗口即停止"}


def m_call_local_tool(args):
    """对齐 VSTO McpManager.CallLocalTool：args = [toolName, argsJson]。
    返回前端约定的 {success, result, taskCompleted}；result 为文本（VSTO 契约，
    前端 mcp-dispatcher 靠 result 字符串做"验证正在播放/not_found"等判定）。"""
    tool_name = args[0] if len(args) > 0 else None
    params_raw = args[1] if len(args) > 1 else None
    if isinstance(params_raw, str):
        try:
            params = json.loads(params_raw) if params_raw.strip() else {}
        except Exception:
            params = {}
    elif isinstance(params_raw, dict):
        params = params_raw
    else:
        params = {}
    if not tool_name:
        raise RuntimeError("CallLocalTool 缺少工具名")
    try:
        result = route_method(str(tool_name), [params])
        ok = not (isinstance(result, dict) and result.get("success") is False)
        completed = result.get("taskCompleted", True) if isinstance(result, dict) else True
    except Exception as e:
        result = str(e)
        ok = False
        completed = True
    # dict → 文本：优先 message/error，否则序列化整个 JSON（tool-card 可再解析）
    if isinstance(result, dict):
        msg = result.get("message") or result.get("error")
        result = str(msg) if msg is not None else json.dumps(result, ensure_ascii=False)
    elif result is None:
        result = ""
    elif not isinstance(result, str):
        result = json.dumps(result, ensure_ascii=False)
    return {"success": ok, "result": result, "taskCompleted": completed}


# ===================== 数据插入方法（对齐 C#/VB 桥） =====================

import re as _re

# InsertParsedData / InsertDataSmart 等用到的字段同义词组（对齐 ExcelDataBridge.IsFieldMatch）
_FIELD_SYNONYM_GROUPS = [
    ["名称", "企业名称", "公司名称", "单位名称", "姓名", "名字", "营业执照名称", "证照名称"],
    ["统一社会信用代码", "信用代码", "社会信用代码", "税号", "纳税人识别号", "税务登记号"],
    ["法定代表人", "法人", "负责人", "代表人"],
    ["注册资本", "注册资金", "资本", "出资额"],
    ["成立日期", "成立时间", "注册日期", "设立日期", "开业日期"],
    ["营业期限", "经营期限", "有效期", "营业期限至"],
    ["经营范围", "业务范围", "经营业务", "主营业务"],
    ["地址", "住所", "所在地", "经营场所", "注册地址", "企业地址"],
    ["电话", "联系电话", "手机号", "手机号码", "联系方式", "传真"],
    ["金额", "价税合计", "总额", "合计", "总金额", "金额小写", "金额大写"],
    ["发票代码", "代码", "发票代码"],
    ["发票号码", "号码", "发票号码"],
    ["开票日期", "日期", "开票时间"],
    ["车票", "车次", "出发站", "到达站", "出发地", "目的地", "始发站", "终到站"],
    ["票价", "金额", "价格", "票价金额"],
    ["座位", "座位号", "席位", "席位号", "车厢座位"],
    ["席别", "座位类型", "席位类型", "舱位", "铺位"],
    ["日期", "乘车日期", "发车日期", "出行日期", "日期时间"],
    ["时间", "发车时间", "乘车时间", "时刻", "发车时刻"],
    ["类型", "公司类型", "企业类型", "主体类型"],
    ["登记机关", "注册机关", "工商局", "市场监管", "登记机关名称"],
    ["营业执照", "证照名称", "证件名称"],
    ["企业信息公示系统链接", "公示链接", "查询链接"],
    ["工商行政管理机构", "管理机构", "登记机构"],
]


def _sanitize_cell_value(value):
    """对齐 VB SanitizeCellValue：防公式注入 + 长数字/身份证/前导零数字作为文本保留。"""
    if value is None:
        return ""
    s = str(value)
    if s == "":
        return s
    t = s.strip()
    if _should_preserve_as_text(t):
        return "'" + s
    t2 = s.lstrip()
    if t2 and t2[0] in ("=", "+", "-", "@"):
        return "'" + s
    return s


def _should_preserve_as_text(value):
    if not value:
        return False
    if _re.match(r"^\d{17}[\dXx]$", value):
        return True
    if len(value) > 15 and _re.match(r"^\d+$", value):
        return True
    if len(value) > 1 and value.startswith("0") and _re.match(r"^\d+$", value):
        return True
    return False


def _is_field_match(header, field):
    """对齐 VB IsFieldMatch：精确 / 包含 / 同义词组。"""
    if not header or not field:
        return False
    h = header.strip().lower()
    f = field.strip().lower()
    if h == f:
        return True
    if h and f and (h in f or f in h):
        return True
    for group in _FIELD_SYNONYM_GROUPS:
        h_in = any((t and t in h) for t in group)
        f_in = any((t and t in f) for t in group)
        if h_in and f_in:
            return True
    return False


def _active_sheet(app):
    """确保有工作簿与活动工作表，返回 (workbook, worksheet)。"""
    wb = ensure_workbook(app)
    ws = app.ActiveSheet
    if ws is None:
        ws = wb.Worksheets(1)
    return wb, ws


def _set_cell(app, row, col, value):
    cell = app.Cells(row, col)
    cell.NumberFormat = "@"
    cell.Value2 = _sanitize_cell_value(value)


def _autofit(ws, start_col, col_count):
    try:
        for c in range(start_col, start_col + col_count):
            ws.Columns(c).AutoFit()
    except Exception:
        pass


def _write_object_array(app, records, start_row, start_col, with_header=True):
    """把对象数组写入 (start_row,start_col) 起的区域；with_header 时先写表头。返回 (行数, 列名列表)。"""
    if not records:
        return 0, []
    columns = list(records[0].keys())
    r = start_row
    if with_header:
        for ci, name in enumerate(columns):
            _set_cell(app, r, start_col + ci, name)
        r += 1
    for rec in records:
        for ci, name in enumerate(columns):
            _set_cell(app, r, start_col + ci, rec.get(name))
        r += 1
    return len(records), columns


def m_insert_table_data(args):
    """对齐 FormatBridge.InsertTableData：二维数组从活动单元格铺开，保持原始布局。"""
    rows_json = args[0] if args else None
    if not rows_json:
        raise RuntimeError("InsertTableData 缺少数据")
    rows = json.loads(rows_json) if isinstance(rows_json, str) else rows_json
    if not isinstance(rows, list) or not rows:
        raise RuntimeError("表格数据为空")
    app = get_excel(create=True)
    wb, ws = _active_sheet(app)
    sr = app.ActiveCell.Row
    sc = app.ActiveCell.Column
    for ri, row in enumerate(rows):
        if not isinstance(row, list):
            row = [row]
            rows[ri] = row
        for ci, val in enumerate(row):
            _set_cell(app, sr + ri, sc + ci, "" if val is None else val)
    max_cols = max(len(r) for r in rows)
    _autofit(ws, sc, max_cols)
    return {"success": True, "message": "表格数据已插入，共%d行" % len(rows)}


def m_insert_data_smart(args):
    """对齐 FormatBridge.InsertDataSmart：payload {records, mode}，records 可为二维数组或对象数组。"""
    payload_json = args[0] if args else None
    if not payload_json:
        raise RuntimeError("InsertDataSmart 缺少数据")
    payload = json.loads(payload_json) if isinstance(payload_json, str) else payload_json
    if not isinstance(payload, dict) or "records" not in payload:
        raise RuntimeError("数据格式不正确，缺少records字段")
    records = payload["records"]
    app = get_excel(create=True)
    wb, ws = _active_sheet(app)
    sr = app.ActiveCell.Row
    sc = app.ActiveCell.Column
    # 二维数组（表格数据）
    if isinstance(records, list) and records and isinstance(records[0], list):
        for ri, row in enumerate(records):
            if not isinstance(row, list):
                row = [row]
            for ci, val in enumerate(row):
                _set_cell(app, sr + ri, sc + ci, "" if val is None else val)
        max_cols = max(len(r) for r in records)
        _autofit(ws, sc, max_cols)
        return {"success": True, "message": "表格数据已插入，共%d行" % len(records)}
    # 对象数组（结构化数据）
    if isinstance(records, list) and records and isinstance(records[0], dict):
        n, cols = _write_object_array(app, records, sr, sc, with_header=True)
        _autofit(ws, sc, len(cols))
        return {"success": True, "message": "数据已插入，共%d行" % n}
    raise RuntimeError("不支持的数据格式")


def m_insert_parsed_data(args):
    """对齐 ExcelDataBridge.InsertParsedData：对象数组，自动匹配现有表头或新建带样式表。"""
    data_json = args[0] if args else None
    if not data_json:
        raise RuntimeError("InsertParsedData 缺少数据")
    data = json.loads(data_json) if isinstance(data_json, str) else data_json
    if isinstance(data, dict):
        data = [data]
    if not isinstance(data, list) or not data or not isinstance(data[0], dict):
        raise RuntimeError("数据格式不正确，需为对象数组")
    app = get_excel(create=True)
    wb, ws = _active_sheet(app)
    new_columns = list(data[0].keys())

    # 扫描前10行 x 50列寻找最佳匹配表头
    best_rate = 0.0
    best_row = -1
    best_mapping = None
    best_headers = None
    for scan in range(1, 11):
        headers = []
        for col in range(1, 51):
            v = app.Cells(scan, col).Value2
            if v is None or str(v).strip() == "":
                break
            headers.append(str(v).strip())
        if len(headers) >= 1:
            match_count = 0
            mapping = {}
            for header in headers:
                for nc in new_columns:
                    if nc not in mapping and _is_field_match(header, nc):
                        match_count += 1
                        mapping[nc] = headers.index(header) + 1
                        break
            rate = match_count / max(len(headers), len(new_columns))
            if rate > best_rate or (rate == best_rate and best_mapping is not None and match_count > len(best_mapping)):
                best_rate = rate
                best_row = scan
                best_mapping = mapping
                best_headers = headers

    matched = best_headers is not None and len(best_mapping) >= 1 and best_rate >= 0.2

    if matched:
        sc = 1
        insert_row = best_row + 1
        max_col = max(len(best_headers), len(new_columns))
        # 找到已有数据的最后一行后追加，避免覆盖
        while True:
            has = False
            for c in range(1, max_col + 1):
                v = app.Cells(insert_row, c).Value2
                if v is not None and str(v).strip() != "":
                    has = True
                    break
            if not has:
                break
            insert_row += 1
        for rec in data:
            for ci, nc in enumerate(new_columns):
                tc = best_mapping.get(nc, 1 + ci)
                _set_cell(app, insert_row, tc, rec.get(nc))
            insert_row += 1
        _autofit(ws, 1, max_col)
        return {"success": True, "message": "已插入%d行数据到现有表格" % len(data),
                "insertedRows": len(data), "foundExistingTable": True}

    # 未匹配：从活动单元格新建表（表头 + 钢蓝样式 + 边框）
    sr = app.ActiveCell.Row
    sc = app.ActiveCell.Column
    for ci, name in enumerate(new_columns):
        cell = app.Cells(sr, sc + ci)
        cell.Value2 = name
        try:
            cell.Font.Bold = True
            cell.Interior.Color = 0x4682B4  # 钢蓝
            cell.Font.Color = 0xFFFFFF
            cell.HorizontalAlignment = -4108  # xlCenter
            cell.VerticalAlignment = -4108
        except Exception:
            pass
    insert_row = sr + 1
    for rec in data:
        for ci, nc in enumerate(new_columns):
            _set_cell(app, insert_row, sc + ci, rec.get(nc))
        insert_row += 1
    try:
        total_rows = len(data) + 1
        total_cols = len(new_columns)
        rng = ws.Range(ws.Cells(sr, sc), ws.Cells(sr + total_rows - 1, sc + total_cols - 1))
        rng.Borders.LineStyle = 1   # xlContinuous
        rng.Borders.Weight = 2      # xlThin
        rng.Borders.Color = 0xC8C8C8
    except Exception:
        pass
    _autofit(ws, sc, len(new_columns))
    return {"success": True, "message": "已创建新表格并插入%d行数据" % len(data),
            "insertedRows": len(data), "foundExistingTable": False}


def m_fill_data_to_excel(args):
    """对齐 ExcelDataBridge.FillDataToExcel：dataJson 为对象或对象数组，dataType 为类型标签。"""
    data_json = args[0] if len(args) > 0 else None
    data_type = args[1] if len(args) > 1 else ""
    if not data_json:
        raise RuntimeError("FillDataToExcel 缺少数据")
    data = json.loads(data_json) if isinstance(data_json, str) else data_json
    if isinstance(data, dict):
        records = [data]
    elif isinstance(data, list):
        records = data
    else:
        raise RuntimeError("FillDataToExcel 不支持的数据类型")
    if not records:
        raise RuntimeError("数据为空")
    app = get_excel(create=True)
    wb, ws = _active_sheet(app)
    sr = app.ActiveCell.Row
    sc = app.ActiveCell.Column
    if isinstance(records[0], dict):
        n, cols = _write_object_array(app, records, sr, sc, with_header=True)
        _autofit(ws, sc, len(cols))
        msg = "数据已成功填入Excel，共%d行" % n
    elif isinstance(records[0], list):
        for ri, row in enumerate(records):
            if not isinstance(row, list):
                row = [row]
            for ci, val in enumerate(row):
                _set_cell(app, sr + ri, sc + ci, "" if val is None else val)
        _autofit(ws, sc, max(len(r) for r in records))
        msg = "数据已成功填入Excel，共%d行" % len(records)
    else:
        raise RuntimeError("不支持的数据格式")
    return {"success": True, "message": msg}


# ===================== 本地 OCR（直接调用 rapidocr_onnxruntime） =====================
# 说明：项目自带 Services/rapidocr_inference_engine 用旧版 RapidOCR(det/rec/cls_dir=...) 实例化签名，
# 与当前已装的 rapidocr_onnxruntime 新版（统一 model_path 风格）不兼容，会抛 KeyError('model_path')。
# 这里直接实例化新版 RapidOCR()（自带默认模型路径，约 0.6s 就绪），自行解析返回格式。
import os as _os
import time as _time
import threading as _threading
import base64 as _base64

_ocr_engine_inst = None
_ocr_engine_init_err = None
_ocr_lock = _threading.Lock()
_ocr_state = {"completed": False, "success": False, "error": None,
              "text": "", "items": [], "raw": None, "path": None}
_ocr_current_engine = "rapidocr"


def _get_rapidocr():
    global _ocr_engine_inst, _ocr_engine_init_err
    if _ocr_engine_inst is not None:
        return _ocr_engine_inst
    if _ocr_engine_init_err is not None:
        raise RuntimeError(_ocr_engine_init_err)
    try:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine_inst = RapidOCR()  # 新版自带模型路径
    except Exception as e:
        _ocr_engine_init_err = "RapidOCR 初始化失败: " + str(e)
        raise
    return _ocr_engine_inst


def _parse_ocr_line(line):
    box = line[0]
    if len(line) >= 3 and isinstance(line[1], (tuple, list)):
        text, confidence = line[1]
    else:
        text = line[1]
        confidence = line[2] if len(line) >= 3 else 0.0
    return box, text, float(confidence)


def m_save_base64_image(args):
    """对齐 ExcelBridge.SaveBase64Image：base64 图片落盘，返回 {success, path}。"""
    data = args[0] if args else None
    file_name = args[1] if len(args) > 1 else "image.png"
    if not data:
        raise RuntimeError("SaveBase64Image 缺少数据")
    if "," in data:
        data = data[data.index(",") + 1:]
    try:
        raw = _base64.b64decode(data)
    except Exception as e:
        raise RuntimeError("base64 解码失败: " + str(e))
    temp_dir = _os.path.join(_os.path.expanduser("~"), "AppData", "Local", "Temp", "ExcelAIOCR")
    try:
        _os.makedirs(temp_dir, exist_ok=True)
    except Exception:
        temp_dir = _os.path.join(_os.path.dirname(__file__), "tmp_ocr")
        _os.makedirs(temp_dir, exist_ok=True)
    safe = _re.sub(r'[^\w\-.]', '_', _os.path.splitext(file_name)[0])
    ext = _os.path.splitext(file_name)[1] or ".png"
    path = _os.path.join(temp_dir, "%s_%d%s" % (safe, int(_time.time() * 1000), ext))
    with open(path, "wb") as f:
        f.write(raw)
    return {"success": True, "path": path}


def _ocr_run(path):
    try:
        engine = _get_rapidocr()
        result, _elapse = engine(path)
        if not result:
            _ocr_state.update(completed=True, success=True, error=None, text="", items=[])
            return
        items = []
        for line in result:
            try:
                box, text, conf = _parse_ocr_line(line)
            except Exception:
                continue
            items.append({"text": text, "confidence": conf, "box": box})
        text = "\n".join((it["text"] or "") for it in items)
        _ocr_state.update(completed=True, success=True, error=None, text=text, items=items, raw=result)
    except Exception as e:
        _ocr_state.update(completed=True, success=False, error=str(e), text="")


def m_start_ocr_recognition(args):
    """对齐 ExcelBridge.StartOcrRecognition：后台启动本地 RapidOCR，立即返回 {started:true}。"""
    path = args[0] if args else None
    if not path or not _os.path.exists(path):
        raise RuntimeError("StartOcrRecognition 缺少图片路径")
    with _ocr_lock:
        _ocr_state.update(completed=False, success=False, error=None, text="", items=[], raw=None, path=path)
    _threading.Thread(target=_ocr_run, args=(path,), daemon=True).start()
    return {"started": True}


def _ocr_cluster_rows(items):
    """按 y 中心聚类成行（对齐插件版 BuildDocumentFromPositionedResults 行合并），
    行内按 x 排序。"""
    def cy(it):
        b = it.get("box") or [[0, 0], [0, 0], [0, 0], [0, 0]]
        ys = [float(p[1]) for p in b]
        return (min(ys) + max(ys)) / 2.0

    def hh(it):
        b = it.get("box") or [[0, 0], [0, 0], [0, 0], [0, 0]]
        ys = [float(p[1]) for p in b]
        return max(1.0, max(ys) - min(ys))

    rows = []
    for it in sorted(items, key=cy):
        placed = False
        for row in rows:
            ref = row[-1]
            tol = max(hh(it), hh(ref)) * 0.6
            if abs(cy(it) - cy(ref)) <= tol:
                row.append(it)
                placed = True
                break
        if not placed:
            rows.append([it])
    for row in rows:
        row.sort(key=lambda it: min(float(p[0]) for p in (it.get("box") or [[0]])))
    rows.sort(key=lambda r: cy(r[0]))
    return rows


def _ocr_row_cells(row):
    """行内按 x 间隙切列（对齐表格重建）：间隙大于字宽的 ~1.2 倍视为新列。"""
    cells, cur, prev_right = [], "", None
    for it in row:
        b = it.get("box") or [[0, 0], [0, 0], [0, 0], [0, 0]]
        left = min(float(p[0]) for p in b)
        right = max(float(p[0]) for p in b)
        txt = str(it.get("text") or "")
        charw = max(1.0, (right - left) / max(1, len(txt)))
        if prev_right is not None:
            gap = left - prev_right
            if gap > charw * 1.2:
                cells.append(cur)
                cur = ""
            elif gap > charw * 0.35:
                cur += " "
        cur += txt
        prev_right = right
    if cur:
        cells.append(cur)
    return cells


def _ocr_build_grid(items):
    """网格聚类（对齐插件版 OcrTableParser 语义）：全局 x 区间聚类出列边界，
    每个 item 按 x 中心归列，行×列 → 规则网格。空缺格为空串（合并单元格
    跨行首格为空，与插件版输出一致）。"""
    rows = _ocr_cluster_rows(items)
    spans = []
    for it in items:
        b = it.get("box")
        if not b:
            continue
        xs = [float(p[0]) for p in b]
        spans.append((min(xs), max(xs)))
    spans.sort()
    col_bounds = []
    for l, r in spans:
        w = max(1.0, r - l)
        best, best_ov = -1, 0.0
        for i, (cl, cr) in enumerate(col_bounds):
            ov = min(r, cr) - max(l, cl)
            if ov > best_ov:
                best, best_ov = i, ov
        # 与某列重叠明显（同列跨行的框 x 区间基本重合）→ 并入该列；否则新列
        if best >= 0 and best_ov >= w * 0.25:
            cl, cr = col_bounds[best]
            col_bounds[best] = (min(cl, l), max(cr, r))
        else:
            col_bounds.append((l, r))
    col_bounds.sort()

    def col_index(cx):
        best, bd = 0, None
        for i, (l, r) in enumerate(col_bounds):
            d = 0.0 if l <= cx <= r else min(abs(cx - l), abs(cx - r))
            if bd is None or d < bd:
                best, bd = i, d
        return best

    grid = []
    for row in rows:
        cells = [""] * len(col_bounds)
        for it in row:
            b = it.get("box")
            if not b:
                continue
            xs = [float(p[0]) for p in b]
            ci = col_index((min(xs) + max(xs)) / 2.0)
            txt = str(it.get("text") or "")
            cells[ci] = (cells[ci] + " " + txt) if cells[ci] else txt
        grid.append(cells)
    return grid


def _ocr_format_items(items, fmt):
    """对齐插件版 OcrFormatFormatter：
    plaintext → raw_ocr_data 坐标 JSON（前端 parseRawOcrDataItems 消费，带置信度+标注框）；
    excel/csv/markdown/html → 全局列边界网格（跨行列对齐，表格不错位）。"""
    if not items:
        return ""
    if fmt in ("excel", "csv", "markdown", "html"):
        grid = _ocr_build_grid(items)
        if fmt in ("excel", "csv"):
            lines = ["# 来源:本地识别 引擎:RapidOCR"]
            for cells in grid:
                lines.append("\t".join(cells))
            return "\ufeff" + "\n".join(lines) + "\n"
        if fmt == "markdown":
            lines = []
            for cells in grid:
                lines.append("| " + " | ".join(cells) + " |")
            if lines:
                lines.insert(0, "| " + " | ".join([""] * len(grid[0])) + " |")
                lines.insert(1, "|" + "|".join(["------"] * len(grid[0])) + "|")
            return "\n".join(lines)
        out = ['<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8"><title>OCR 识别结果</title>',
               '<style>table{border-collapse:collapse}td,th{border:1px solid #999;padding:4px 8px}</style></head><body><table>']
        for cells in grid:
            out.append("<tr>" + "".join("<td>%s</td>" % c for c in cells) + "</tr>")
        out.append("</table></body></html>")
        return "\n".join(out)
    rows = _ocr_cluster_rows(items)
    data = []
    for row in rows:
        for it in row:
            b = it.get("box")
            if not b:
                continue
            try:
                conf = float(it.get("confidence") or 0)
            except Exception:
                conf = 0
            data.append({"text": str(it.get("text") or ""), "confidence": conf, "box": b})
    return json.dumps(data, ensure_ascii=False)


def m_get_ocr_result(args):
    fmt = str(args[0] or "plaintext").strip().lower() if args else "plaintext"
    if fmt == "raw":
        fmt = "plaintext"
    if not _ocr_state["completed"]:
        return {"ready": False}
    if _ocr_state["error"] is not None:
        return {"ready": True, "success": False, "error": _ocr_state["error"]}
    text = _ocr_format_items(_ocr_state.get("items") or [], fmt)
    return {"ready": True, "success": True, "text": text,
            "rawText": _ocr_state.get("text") or "", "format": fmt}


def m_get_ocr_structured_result(args):
    if not _ocr_state["completed"]:
        return {"ready": False}
    if _ocr_state["error"] is not None:
        return {"ready": True, "success": False, "error": _ocr_state["error"], "items": []}
    return {"ready": True, "success": True, "items": _ocr_state["items"]}


def m_get_current_ocr_engine(args):
    return _ocr_current_engine


def m_set_ocr_engine(args):
    global _ocr_current_engine
    eng = args[0] if args else None
    if eng:
        _ocr_current_engine = eng
    return {"success": True, "engine": _ocr_current_engine}


def m_get_ocr_engine_status(args):
    eng = args[0] if args else _ocr_current_engine
    try:
        from rapidocr_onnxruntime import RapidOCR
        ok = True
    except Exception:
        ok = False
    return {"success": True, "status": "ready" if ok else "not_initialized",
            "engine": "RapidOCR", "available": ok, "requested": eng}


# ===================== 轻量查询 / 状态方法 =====================
def m_get_active_sheet_name(args):
    app = get_excel(create=False)
    if app is None or app.ActiveSheet is None:
        return ""
    return app.ActiveSheet.Name


def m_get_active_workbook_name(args):
    app = get_excel(create=False)
    if app is None or app.ActiveWorkbook is None:
        return ""
    return app.ActiveWorkbook.Name


def m_get_sheets(args):
    app = get_excel(create=False)
    if app is None:
        return []
    names = []
    try:
        for ws in app.Worksheets:
            names.append(ws.Name)
    except Exception:
        pass
    return names


def m_activate_sheet(args):
    """激活指定名称的工作表（网页版"选择工作表"入口）。"""
    name = args[0] if args else None
    if not name:
        raise RuntimeError("缺少工作表名")
    app = get_excel(create=False)
    if app is None:
        raise RuntimeError("Excel 未连接")
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    try:
        ws = wb.Worksheets(name)
    except Exception:
        raise RuntimeError("工作表不存在: " + str(name))
    ws.Activate()
    return {"name": ws.Name, "index": ws.Index}


def m_get_selected_range_data(args):
    app = get_excel(create=False)
    if app is None:
        return {"success": False, "error": "未连接到 Excel"}
    try:
        sel = app.Selection
        return {"success": True, "address": sel.Address, "value": sel.Value}
    except Exception as e:
        return {"success": False, "error": str(e)}


def m_get_undo_stack(args):
    return []


def m_get_redo_stack(args):
    return []


# ===================== web / time 本地工具 =====================

def m_time_now(args):
    import datetime as _dt
    p = _tool_params(args)
    with_s = _as_bool(p.get("withSeconds"), False)
    now = _dt.datetime.now().astimezone()
    off = now.utcoffset() or _dt.timedelta(0)
    total_min = int(off.total_seconds() // 60)
    sign = "+" if total_min >= 0 else "-"
    tz = "UTC%s%02d:%02d" % (sign, abs(total_min) // 60, abs(total_min) % 60)
    wd = (now.weekday() + 1) % 7  # 周日=0，与 C# DayOfWeek 一致
    week_names = "日一二三四五六"
    return {"success": True,
            "iso": now.isoformat(timespec="seconds" if with_s else "minutes"),
            "date": now.strftime("%Y-%m-%d"),
            "time": now.strftime("%H:%M:%S" if with_s else "%H:%M"),
            "weekday": wd, "weekdayName": "星期" + week_names[wd],
            "isWeekend": wd in (0, 6), "timezone": tz}


def _http_decode(b):
    if not b:
        return ""
    for enc in ("utf-8", "gbk"):
        try:
            return b.decode(enc)
        except Exception:
            pass
    return b.decode("utf-8", "ignore")


def m_web_search(args):
    import urllib.parse
    import urllib.request
    p = _tool_params(args)
    q = str(p.get("query") or "").strip()
    if not q:
        raise RuntimeError("缺少必要参数 query（搜索关键词）")
    maxr = max(1, min(10, int(p.get("maxResults") or 5)))
    url = "https://www.bing.com/search?q=" + urllib.parse.quote(q) + "&num=" + str(maxr)
    region = str(p.get("region") or "").strip()
    if region:
        url += "&setmkt=" + urllib.parse.quote(region)
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"})
    html = urllib.request.urlopen(req, timeout=15).read().decode("utf-8", "ignore")
    results = []
    for m in re.finditer(r'<li class="b_algo".*?</li>', html, re.S):
        block = m.group(0)
        a = re.search(r'<h2[^>]*>\s*<a[^>]+href="([^"]+)"[^>]*>(.*?)</a>', block, re.S)
        if not a:
            continue
        title = re.sub(r"<[^>]+>", "", a.group(2)).strip()
        snippet = ""
        pm = re.search(r"<p[^>]*>(.*?)</p>", block, re.S)
        if pm:
            snippet = re.sub(r"<[^>]+>", "", pm.group(1)).strip()
        results.append({"title": title, "url": a.group(1), "snippet": snippet})
        if len(results) >= maxr:
            break
    payload = {"query": q, "total": len(results), "results": results}
    if not results:
        payload["note"] = "未解析到网页结果（可能触发反爬、或确实无与该关键词匹配的独立网页），可尝试更换关键词。"
    return payload


# ===================== computer 本地工具（Windows 自动化） =====================

_COMPUTER_STATE = {"safeMode": True}

_PROTECTED_PROCESSES = {
    "excel", "winword", "powerpnt", "outlook", "onenote", "msaccess",
    "explorer", "dwm", "csrss", "winlogon", "lsass", "services", "smss",
    "system", "idle", "svchost", "wininit", "fontdrvhost",
    "cursor", "devenv", "wps", "et", "wpp", "wpspdf",
}
_BLOCKED_LAUNCH = {
    "cmd", "cmd.exe", "powershell", "powershell.exe", "pwsh", "pwsh.exe",
    "wscript", "wscript.exe", "cscript", "cscript.exe", "mshta", "mshta.exe",
    "rundll32", "rundll32.exe", "regsvr32", "regsvr32.exe",
    "bitsadmin", "bitsadmin.exe", "certutil", "certutil.exe",
    "msiexec", "msiexec.exe", "ftp", "ftp.exe", "curl", "curl.exe",
    "wget", "wget.exe", "bash", "bash.exe", "wsl", "wsl.exe",
    "命令行", "命令提示符",
}
_SAFE_CMD_ALLOW = {"hostname", "whoami", "ver", "date /t", "time /t"}
_BLOCKED_LAUNCH_EXT = {".bat", ".cmd", ".ps1", ".vbs", ".js", ".jse", ".wsf", ".hta", ".scr", ".com", ".msi"}


def _proc_base(name):
    n = (name or "").strip().lower()
    if n.endswith(".exe"):
        n = n[:-4]
    return n


def _is_protected_process(name):
    return not name or _proc_base(name) in _PROTECTED_PROCESSES


def _is_blocked_launch(target):
    t = (target or "").strip()
    if not t:
        return True
    low = t.lower()
    if low.startswith("http://") or low.startswith("https://"):
        return False
    import os
    fn = os.path.basename(t) or t
    base = os.path.splitext(fn)[0]
    ext = os.path.splitext(fn)[1].lower()
    if fn.lower() in _BLOCKED_LAUNCH or base.lower() in _BLOCKED_LAUNCH:
        return True
    return ext in _BLOCKED_LAUNCH_EXT


def _command_looks_safe(cmd):
    c = (cmd or "").strip()
    if not c:
        return False
    for ch in "|;&`$()<>\n\r":
        if ch in c:
            return False
    if "&&" in c or "||" in c:
        return False
    return c.lower() in _SAFE_CMD_ALLOW

_APP_ALIAS = {
    "qq音乐": ["QQMusic.exe"], "音乐": ["QQMusic.exe", "cloudmusic.exe", "kugou.exe", "kuwo.exe"],
    "浏览器": ["msedge.exe", "chrome.exe", "firefox.exe"], "edge": ["msedge.exe"], "chrome": ["chrome.exe"],
    "记事本": ["notepad.exe", "Notepad.exe"], "计算器": ["CalculatorApp.exe", "calc.exe"],
    "excel": ["EXCEL.EXE"], "word": ["WINWORD.EXE"], "ppt": ["POWERPNT.EXE"],
    "outlook": ["OUTLOOK.EXE"], "微信": ["WeChat.exe", "Weixin.exe"],
    "企业微信": ["WXWork.exe"], "钉钉": ["DingTalk.exe"], "画图": ["mspaint.exe"],
}


def _not_impl(tool, need):
    return {"success": False, "notImplemented": True,
            "error": "工具 %s 依赖 %s，网页端(本地桥)暂不支持，需在 VSTO 宿主内使用。" % (tool, need)}


def _get_cursor_pos():
    import ctypes

    class POINT(ctypes.Structure):
        _fields_ = [("x", ctypes.c_long), ("y", ctypes.c_long)]

    pt = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(pt))
    return pt.x, pt.y


def m_computer_screen_get_info(args):
    import ctypes
    u = ctypes.windll.user32
    u.SetProcessDPIAware()
    return {"success": True, "width": u.GetSystemMetrics(0), "height": u.GetSystemMetrics(1),
            "cursor": {"x": _get_cursor_pos()[0], "y": _get_cursor_pos()[1]}}


def _grab_screen(region=None):
    from PIL import ImageGrab
    return ImageGrab.grab(bbox=region, all_screens=True)


def _img_payload(img, p):
    fmt = str(p.get("returnFormat") or "base64").strip().lower()
    buf = io.BytesIO()
    img.save(buf, "PNG")
    data = buf.getvalue()
    if fmt == "path":
        desk = os.path.join(os.environ.get("USERPROFILE", "."), "Desktop")
        if not os.path.isdir(desk):
            desk = os.path.abspath(".")
        path = os.path.join(desk, "screen_%s.png" % _ts())
        with open(path, "wb") as f:
            f.write(data)
        return {"success": True, "path": path}
    return {"success": True, "format": "png", "base64": base64.b64encode(data).decode()}


def m_computer_screen_capture(args):
    p = _tool_params(args)
    region = None
    if p.get("x") is not None and p.get("y") is not None:
        x, y = int(p["x"]), int(p["y"])
        region = (x, y, x + int(p.get("width") or 0), y + int(p.get("height") or 0))
    return _img_payload(_grab_screen(region), p)


def _find_window_by_keyword(kw):
    import win32gui
    kw_l = kw.lower()
    hits = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t and kw_l in t.lower():
                hits.append((hwnd, t))
    win32gui.EnumWindows(cb, None)
    return hits


def m_computer_capture_window(args):
    p = _tool_params(args)
    app = str(p.get("app") or "").strip()
    if not app:
        raise RuntimeError("缺少 app 参数")
    hits = _find_window_by_keyword(app)
    if not hits:
        return {"success": False, "error": "未找到标题包含 '%s' 的可见窗口" % app}
    import win32gui
    import win32process
    import psutil
    hwnd, title = hits[0]
    # 找到该窗口所属进程名，再按进程的窗口聚合成区域（简化：直接用窗口矩形）
    from PIL import ImageGrab
    try:
        import win32gui as _w
        l, t, r, b = _w.GetWindowRect(hwnd)
        if r <= l or b <= t:
            return {"success": False, "error": "窗口区域无效（可能已最小化）"}
        img = ImageGrab.grab(bbox=(max(l, 0), max(t, 0), r, b), all_screens=True)
        return _img_payload(img, p)
    except Exception as e:
        return {"success": False, "error": "窗口截图失败: " + str(e)}


def m_computer_mouse_move(args):
    import ctypes
    p = _tool_params(args)
    x, y = int(p.get("x")), int(p.get("y"))
    ctypes.windll.user32.SetProcessDPIAware()
    ctypes.windll.user32.SetCursorPos(x, y)
    return {"success": True, "x": x, "y": y}


def m_computer_mouse_click(args):
    import ctypes
    p = _tool_params(args)
    u = ctypes.windll.user32
    u.SetProcessDPIAware()
    if p.get("x") is not None and p.get("y") is not None:
        u.SetCursorPos(int(p["x"]), int(p["y"]))
    btn = str(p.get("button") or "left").lower()
    clicks = int(p.get("clicks") or 1)
    down, up = {"left": (0x0002, 0x0004), "right": (0x0008, 0x0010), "middle": (0x0020, 0x0040)}[btn]
    for _ in range(max(1, clicks)):
        u.mouse_event(down, 0, 0, 0, 0)
        u.mouse_event(up, 0, 0, 0, 0)
    return {"success": True, "button": btn, "clicks": clicks,
            "at": {"x": _get_cursor_pos()[0], "y": _get_cursor_pos()[1]}}


def m_computer_mouse_drag(args):
    import ctypes
    p = _tool_params(args)
    u = ctypes.windll.user32
    u.SetProcessDPIAware()
    u.SetCursorPos(int(p["fromX"]), int(p["fromY"]))
    u.mouse_event(0x0002, 0, 0, 0, 0)  # left down
    u.SetCursorPos(int(p["toX"]), int(p["toY"]))
    u.mouse_event(0x0004, 0, 0, 0, 0)  # left up
    return {"success": True, "from": [int(p["fromX"]), int(p["fromY"])],
            "to": [int(p["toX"]), int(p["toY"])]}


def m_computer_mouse_scroll(args):
    import ctypes
    p = _tool_params(args)
    amount = int(p.get("amount") or 0)
    ctypes.windll.user32.mouse_event(0x0800, 0, 0, amount * 120, 0)  # WHEEL
    return {"success": True, "amount": amount}


_VK = {"enter": 0x0D, "tab": 0x09, "esc": 0x1B, "escape": 0x1B, "space": 0x20,
       "backspace": 0x08, "delete": 0x2E, "del": 0x2E, "insert": 0x2D,
       "home": 0x24, "end": 0x23, "pageup": 0x21, "pagedown": 0x22,
       "up": 0x26, "down": 0x28, "left": 0x25, "right": 0x27,
       "shift": 0x10, "ctrl": 0x11, "control": 0x11, "alt": 0x12, "win": 0x5B,
       # 多媒体键（全局生效，QQ音乐/网易云/浏览器媒体等通用；此前缺失导致 computer_key_press
       # 传 media_play_pause 会抛"无法识别的键名"，dispatcher 的播放/切歌硬路径全部失效）
       "media_play_pause": 0xB3, "media_play": 0xB3, "media_pause": 0xB3,
       "play_pause": 0xB3, "media_next": 0xB0, "next_track": 0xB0,
       "media_prev": 0xB1, "prev_track": 0xB1, "media_previous": 0xB1,
       "media_stop": 0xB2, "volume_mute": 0xAD, "mute": 0xAD,
       "volume_up": 0xAF, "volume_down": 0xAE}


def _vk_of(name):
    n = str(name).strip().lower()
    if n in _VK:
        return _VK[n]
    if len(n) == 1 and n.isalnum():
        return ord(n.upper())
    if re.fullmatch(r"f[1-9]|f1[0-2]", n):
        return 0x70 + int(n[1:]) - 1
    raise RuntimeError("无法识别的键名: " + str(name))


def _key_event(vk, up):
    import ctypes
    ctypes.windll.user32.keybd_event(vk, 0, 2 if up else 0, 0)


def m_computer_key_press(args):
    p = _tool_params(args)
    vk = _vk_of(p.get("key"))
    _key_event(vk, False)
    _key_event(vk, True)
    return {"success": True, "key": str(p.get("key"))}


def m_computer_key_hotkey(args):
    p = _tool_params(args)
    keys = [k.strip() for k in str(p.get("keys") or "").split("+") if k.strip()]
    if not keys:
        raise RuntimeError("缺少 keys 参数")
    vks = [_vk_of(k) for k in keys]
    for vk in vks:
        _key_event(vk, False)
    for vk in reversed(vks):
        _key_event(vk, True)
    return {"success": True, "keys": keys}


def m_computer_type_text(args):
    import ctypes
    p = _tool_params(args)
    text = str(p.get("text") or "")
    if not text:
        raise RuntimeError("缺少 text 参数")
    PUL = ctypes.POINTER(ctypes.c_ulong)

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [("wVk", ctypes.c_ushort), ("wScan", ctypes.c_ushort),
                    ("dwFlags", ctypes.c_ulong), ("time", ctypes.c_ulong),
                    ("dwExtraInfo", PUL)]

    class _I(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _anonymous_ = ("i",)
        _fields_ = [("type", ctypes.c_ulong), ("i", _I)]

    KEYEVENTF_UNICODE = 0x0004
    KEYEVENTF_KEYUP = 0x0002
    for ch in text:
        for flags in (KEYEVENTF_UNICODE, KEYEVENTF_UNICODE | KEYEVENTF_KEYUP):
            inp = INPUT(type=1)
            inp.ki = KEYBDINPUT(0, ord(ch), flags, 0, None)
            ctypes.windll.user32.SendInput(1, ctypes.byref(inp), ctypes.sizeof(INPUT))
    return {"success": True, "length": len(text)}


def m_computer_locate_color(args):
    p = _tool_params(args)
    color = str(p.get("color") or "").strip().lstrip("#")
    if len(color) != 6:
        raise RuntimeError("color 需为 #RRGGBB 格式")
    rgb = (int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))
    tol = int(p.get("tolerance") or 10)
    img = _grab_screen().convert("RGB")
    w, h = img.size
    px = img.load()
    matches = []
    for y in range(0, h, 2):
        for x in range(0, w, 2):
            r, g, b = px[x, y]
            if abs(r - rgb[0]) <= tol and abs(g - rgb[1]) <= tol and abs(b - rgb[2]) <= tol:
                matches.append({"x": x, "y": y})
                if len(matches) >= 20:
                    break
        if len(matches) >= 20:
            break
    return {"success": True, "found": bool(matches), "matches": matches}


def m_computer_clipboard_read(args):
    import win32clipboard
    win32clipboard.OpenClipboard()
    try:
        if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
            text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        else:
            text = ""
    finally:
        win32clipboard.CloseClipboard()
    return {"success": True, "text": text}


def m_computer_clipboard_write(args):
    import win32clipboard
    p = _tool_params(args)
    text = str(p.get("text") or "")
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()
    return {"success": True, "length": len(text)}


# 打开应用：口语别名 → 有序 exe 候选（镜像 VSTO ComputerAgentTools._appAliases）
_OPEN_ALIASES = {
    "浏览器": ["chrome.exe", "msedge.exe"], "browser": ["chrome.exe", "msedge.exe"],
    "chrome": ["chrome.exe", "msedge.exe"], "chrom": ["chrome.exe", "msedge.exe"],
    "谷歌浏览器": ["chrome.exe", "msedge.exe"], "googlechrome": ["chrome.exe", "msedge.exe"],
    "edge": ["msedge.exe"], "msedge": ["msedge.exe"], "edge浏览器": ["msedge.exe"],
    "火狐": ["firefox.exe"], "firefox": ["firefox.exe"],
    "记事本": ["notepad.exe"], "notepad": ["notepad.exe"],
    "计算器": ["calc.exe"], "calculator": ["calc.exe"],
    "画图": ["mspaint.exe"], "mspaint": ["mspaint.exe"],
    "cmd": ["cmd.exe"], "命令行": ["cmd.exe"], "命令提示符": ["cmd.exe"],
    "控制面板": ["control.exe"], "资源管理器": ["explorer.exe"], "explorer": ["explorer.exe"],
    "微信": ["WeChat.exe"], "wechat": ["WeChat.exe"], "weixin": ["WeChat.exe"],
    "qq音乐": ["QQMusic.exe"], "qqmusic": ["QQMusic.exe"], "音乐": ["QQMusic.exe"],
    "excel": ["excel.exe"], "excel表格": ["excel.exe"], "表格": ["excel.exe"],
    "word": ["winword.exe"], "word文档": ["winword.exe"],
    "ppt": ["powerpnt.exe"], "powerpoint": ["powerpnt.exe"], "幻灯片": ["powerpnt.exe"],
    "网易云": ["cloudmusic.exe"], "网易云音乐": ["cloudmusic.exe"],
    "qq": ["QQ.exe"], "钉钉": ["DingTalk.exe"], "企业微信": ["WXWork.exe"],
}

# 常见应用标准安装目录（键为不含扩展名的 exe 基名；镜像 VSTO _exeDirs）
_OPEN_EXE_DIRS = {
    "chrome": ["Google\\Chrome\\Application"],
    "msedge": ["Microsoft\\Edge\\Application"],
    "firefox": ["Mozilla Firefox"],
    "qqmusic": ["Tencent\\QQMusic"],
    "wechat": ["Tencent\\WeChat", "WeChat"],
    "weixin": ["Tencent\\WeChat", "Tencent\\Weixin", "WeChat"],
    "wxwork": ["Tencent\\WXWork"],
    "qq": ["Tencent\\QQ", "Tencent\\QQNT"],
    "dingtalk": ["DingTalk"],
    "cloudmusic": ["NetEase\\CloudMusic", "CloudMusic"],
}


def _open_known_dirs(base):
    import winreg
    dirs = []
    for root_name, env in (("LocalApplicationData", os.environ.get("LOCALAPPDATA", "")),
                           ("ProgramFiles", os.environ.get("ProgramFiles", "")),
                           ("ProgramFiles(x86)", os.environ.get("ProgramFiles(x86)", ""))):
        base_dir = env
        if not base_dir:
            continue
        for sub in _OPEN_EXE_DIRS.get(base, []):
            dirs.append(os.path.join(base_dir, sub))
    return dirs


def _locate_launch_target(exe_name, low_name):
    """把可执行名解析成可启动目标；找不到返回 ""（镜像 VSTO LocateLaunchTarget）。"""
    # ① 已含路径 → 校验存在
    if "\\" in exe_name or "/" in exe_name:
        return exe_name if os.path.isfile(exe_name) else ""
    base = os.path.splitext(os.path.basename(exe_name))[0].lower()
    # ② 标准安装目录
    for d in _open_known_dirs(base):
        probe = os.path.join(d, exe_name)
        if os.path.isfile(probe):
            return probe
    # ③ 开始菜单快捷方式（.lnk/.url，按 exe 基名/用户原词匹配）
    menus = []
    for env in ("ProgramData", "APPDATA"):
        base_dir = os.environ.get(env, "")
        if base_dir:
            menus.append(os.path.join(base_dir, "Microsoft", "Windows", "Start Menu", "Programs"))
    target_name = base
    for menu in menus:
        if not os.path.isdir(menu):
            continue
        for root, _dirs, files in os.walk(menu):
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext not in (".lnk", ".url"):
                    continue
                fname = os.path.splitext(f)[0].lower()
                if fname == target_name or fname == low_name:
                    return os.path.join(root, f)
                if len(low_name) >= 2 and low_name in fname:
                    return os.path.join(root, f)
    # ④ PATH where.exe
    if " " not in exe_name:
        try:
            r = subprocess.run(["where.exe", exe_name], capture_output=True, timeout=3,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            line = _http_decode(r.stdout).splitlines()
            if line and os.path.isfile(line[0].strip()):
                return line[0].strip()
        except Exception:
            pass
    # ⑤ 兜底：裸 exe 名交给 ShellExecute（系统 App Paths / System32 可解析，如 calc/notepad）
    return exe_name


def _is_browser_intent(app_name):
    low = str(app_name or "").strip().lower()
    return low in ("浏览器", "browser") or "chrome" in low or "edge" in low \
        or "火狐" in low or "firefox" in low or "谷歌" in low


def m_computer_open_application(args):
    """对齐 VSTO ExecuteOpenApplication：支持 path 或 app（口语别名）。
    解析链：别名候选 → 标准安装目录 → 开始菜单快捷方式 → PATH where.exe → 裸 exe；
    浏览器意图解析失败时唤起系统默认浏览器。"""
    p = _tool_params(args)
    app_name = str(p.get("app") or "").strip()
    target = str(p.get("path") or "").strip()
    if _COMPUTER_STATE["safeMode"] and target:
        return {"success": False, "error": "安全模式已开启，已拦截任意路径打开。请使用应用名（如 Excel、Chrome），不要传绝对路径。"}
    if _COMPUTER_STATE["safeMode"] and app_name.strip().lower() in _BLOCKED_LAUNCH:
        return {"success": False, "error": "安全模式已开启，已拦截打开命令行/脚本宿主（可用 computer_set_safe_mode 关闭）"}
    if not target:
        if not app_name:
            raise RuntimeError("未指定要打开的应用（缺 path 或 app）")
        low = app_name.lower()
        candidates = list(_OPEN_ALIASES.get(low, []))
        if not candidates:
            candidates.append(app_name if app_name.lower().endswith(".exe") else app_name + ".exe")
        for cand in candidates:
            t = _locate_launch_target(cand, low)
            if t:
                target = t
                break
    if not target and _is_browser_intent(app_name):
        subprocess.Popen("start \"\" \"https://www.bing.com\"", shell=True,
                         creationflags=subprocess.CREATE_NO_WINDOW)
        return {"success": True, "message": "使用系统默认浏览器打开"}
    if not target:
        return {"success": False,
                "error": "未指定要打开的应用（缺 path 或 app），且无法按名称解析到已安装程序: " + app_name}
    if _is_blocked_launch(target):
        return {"success": False, "error": "已拦截打开命令行/脚本宿主或脚本文件"}
    argstr = str(p.get("arguments") or "").strip()
    try:
        if argstr:
            proc = subprocess.Popen('"%s" %s' % (target, argstr), shell=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            proc = subprocess.Popen('start "" "%s"' % target, shell=True,
                                    creationflags=subprocess.CREATE_NO_WINDOW)
    except Exception as e:
        if _is_browser_intent(app_name) and not target.lower().startswith("http"):
            subprocess.Popen("start \"\" \"https://www.bing.com\"", shell=True,
                             creationflags=subprocess.CREATE_NO_WINDOW)
            return {"success": True, "message": "使用系统默认浏览器打开（%s 启动失败）" % target}
        return {"success": False, "error": "打开应用失败 [%s]: %s（可用 path 指定完整绝对路径重试）" % (target, e)}
    return {"success": True, "target": target, "app": app_name or None,
            "pid": proc.pid if hasattr(proc, "pid") else None}


def m_computer_execute_command(args):
    p = _tool_params(args)
    cmd = str(p.get("command") or "").strip()
    if not cmd:
        raise RuntimeError("缺少 command 参数")
    if _COMPUTER_STATE["safeMode"]:
        return {"success": False, "error": "安全模式已开启，已拦截命令执行（可用 computer_set_safe_mode 关闭）"}
    if not _command_looks_safe(cmd):
        return {"success": False, "error": "已拒绝任意命令。关闭安全模式后也仅允许 hostname / whoami / ver / date /t / time /t，禁止管道、重定向与复合命令。"}
    r = subprocess.run(cmd, shell=True, capture_output=True, timeout=8,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    return {"success": True, "exitCode": r.returncode,
            "stdout": _http_decode(r.stdout), "stderr": _http_decode(r.stderr)}


def m_computer_wait(args):
    p = _tool_params(args)
    ms = max(0, min(60000, int(p.get("ms") or 0)))
    time.sleep(ms / 1000.0)
    return {"success": True, "waitedMs": ms}


def m_computer_set_safe_mode(args):
    p = _tool_params(args)
    _COMPUTER_STATE["safeMode"] = _as_bool(p.get("enabled"), False)
    return {"success": True, "safeMode": _COMPUTER_STATE["safeMode"]}


def m_computer_mouse_position_get(args):
    x, y = _get_cursor_pos()
    return {"success": True, "x": x, "y": y}


def m_computer_window_list(args):
    import win32gui
    p = _tool_params(args)
    kw = str(p.get("keyword") or "").strip().lower()
    titles = []

    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            t = win32gui.GetWindowText(hwnd)
            if t:
                titles.append({"hwnd": hwnd, "title": t})
    win32gui.EnumWindows(cb, None)
    if kw:
        titles = [t for t in titles if kw in t["title"].lower()]
    return {"success": True, "count": len(titles), "windows": titles}


def m_computer_process_list(args):
    import psutil
    p = _tool_params(args)
    name = str(p.get("name") or "").strip().lower()
    procs = []
    for pr in psutil.process_iter(["pid", "name"]):
        try:
            n = pr.info.get("name") or ""
            if name and name not in n.lower():
                continue
            procs.append({"pid": pr.info["pid"], "name": n})
        except Exception:
            pass
    return {"success": True, "count": len(procs), "processes": procs}


def m_computer_process_kill(args):
    p = _tool_params(args)
    name = str(p.get("name") or "").strip()
    if not name:
        raise RuntimeError("缺少 name 参数")
    if _COMPUTER_STATE["safeMode"]:
        return {"success": False, "error": "安全模式已开启，已拦截进程结束（可用 computer_set_safe_mode 关闭）"}
    if not name.lower().endswith(".exe"):
        name += ".exe"
    r = subprocess.run(["taskkill", "/IM", name, "/F"], capture_output=True,
                       creationflags=subprocess.CREATE_NO_WINDOW)
    return {"success": r.returncode == 0, "name": name,
            "message": (_http_decode(r.stdout) or _http_decode(r.stderr)).strip()}


def _alias_processes(app):
    a = app.strip().lower()
    names = []
    for k, v in _APP_ALIAS.items():
        if k in a or a in k:
            names.extend(v)
    return list(dict.fromkeys(names)) or [app if app.lower().endswith(".exe") else app + ".exe"]


def m_computer_close_application(args):
    p = _tool_params(args)
    app = str(p.get("app") or "").strip()
    if not app:
        raise RuntimeError("缺少 app 参数")
    if _COMPUTER_STATE["safeMode"]:
        return {"success": False, "error": "安全模式已开启，已拦截关闭应用"}
    closed = []
    import psutil
    targets = _alias_processes(app)
    for pr in psutil.process_iter(["pid", "name"]):
        try:
            n = (pr.info.get("name") or "").lower()
            if any(n == t.lower() for t in targets):
                pr.kill()
                closed.append(pr.info["name"])
        except Exception:
            pass
    return {"success": True, "closed": closed,
            "message": ("已结束: " + ", ".join(closed)) if closed else ("未找到运行中的 %s" % app)}


def _ocr_screen_text():
    """整屏 OCR 文本（本地 RapidOCR）。"""
    try:
        items = _ocr_screen_items({})
        return "\n".join(it["text"] for it in items)
    except Exception:
        return ""


def _looks_playing(s):
    low = (s or "").lower()
    return ("暂停" in s) or ("正在播放" in s) or ("播放中" in s) or ("pause" in low)


def _find_main_window_by_process(cand_names):
    """按进程名（不含扩展名）轮询查找应用主窗口，至多约 4s。"""
    import psutil
    import win32gui
    import win32process
    names = {}
    for pr in psutil.process_iter(["pid", "name"]):
        try:
            n = (pr.info.get("name") or "")
            names[pr.info["pid"]] = os.path.splitext(n)[0].lower()
        except Exception:
            pass
    cand = set(cand_names)
    for _ in range(8):
        hit = [None]

        def cb(hwnd, _):
            if hit[0] or not win32gui.IsWindowVisible(hwnd):
                return
            if not win32gui.GetWindowText(hwnd):
                return
            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                return
            if names.get(pid, "") in cand:
                hit[0] = hwnd
        win32gui.EnumWindows(cb, None)
        if hit[0]:
            return hit[0]
        time.sleep(0.5)
    return None


def _force_foreground(hwnd):
    """尽最大努力把窗口置前，返回是否已成为前台窗口。
    后台进程直接 SetForegroundWindow 常被 Windows 静默拒绝（不抛异常但无效），
    故失败时用 AttachThreadInput 共享前台线程输入队列后再置前。"""
    import win32gui
    try:
        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
    except Exception:
        pass
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    try:
        if win32gui.GetForegroundWindow() == hwnd:
            return True
    except Exception:
        pass
    try:
        import ctypes
        u = ctypes.windll.user32
        fg = u.GetForegroundWindow()
        tid_fg = u.GetWindowThreadProcessId(fg, None)
        tid_cur = u.GetCurrentThreadId()
        u.AttachThreadInput(tid_fg, tid_cur, True)
        try:
            u.SetForegroundWindow(hwnd)
        finally:
            u.AttachThreadInput(tid_fg, tid_cur, False)
    except Exception:
        pass
    try:
        return win32gui.GetForegroundWindow() == hwnd
    except Exception:
        return False


def m_computer_play_media(args):
    """对齐 VSTO ExecutePlayMedia：聚焦应用 → OCR 判断是否已在播 → 触发 → OCR 验证。
    触发键择一（避免双触发把"刚播"又暂停）：
      - 已成功置前 → 空格键（桌面播放器通用切换）
      - 置前失败（后台进程抢不到焦点，空格会打到别的窗口）→ 全局媒体键 VK_MEDIA_PLAY_PAUSE
    """
    p = _tool_params(args)
    app = str(p.get("app") or "").strip()
    if not app:
        raise RuntimeError("需要指定 app（应用名/别名）")
    low = app.lower()
    cand = [os.path.splitext(os.path.basename(al))[0].lower() for al in _OPEN_ALIASES.get(low, [])]
    base_self = os.path.splitext(low)[0]
    if base_self and base_self not in cand:
        cand.append(base_self)
    if not cand:
        cand.append(low.replace(".exe", ""))
    hwnd = _find_main_window_by_process(cand)
    if hwnd is None:
        return {"success": False, "error": "未找到正在运行的应用主窗口: " + app}
    is_fg = _force_foreground(hwnd)
    time.sleep(1.2)
    start_text = _ocr_screen_text()
    if start_text and _looks_playing(start_text):
        return {"success": True, "message": "已正在播放，无需触发"}
    vk = 0x20 if is_fg else 0xB3  # 空格 / 全局媒体播放暂停键
    key_label = "空格键" if is_fg else "媒体键"
    _key_event(vk, False)
    _key_event(vk, True)
    time.sleep(1.0)
    if _looks_playing(_ocr_screen_text()):
        return {"success": True, "message": "已触发播放并验证正在播放（%s）" % key_label}
    # 键已发送成功（触发即视为已执行）；界面为纯图形按钮时 OCR 读不出播放态，如实转中性成功。
    return {"success": True,
            "message": "已发送%s触发播放（播放键为图形按钮，OCR 无法确认播放态，请以听感为准）" % key_label}


def m_computer_focus_window(args):
    p = _tool_params(args)
    app = str(p.get("app") or "").strip()
    if not app:
        raise RuntimeError("缺少 app 参数")
    import win32gui
    target = None
    for alias in [app] + [k for k in _APP_ALIAS if k in app.lower()]:
        hits = _find_window_by_keyword(alias)
        if hits:
            target = hits[0][0]
            break
    if target is None:
        return {"success": False, "error": "未找到标题匹配 '%s' 的可见窗口" % app}
    _force_foreground(target)
    return {"success": True, "window": win32gui.GetWindowText(target)}


# computer 组：OCR 类用本地 RapidOCR 真实实现（复用桥内 OCR 引擎）
def _ocr_screen_items(p):
    """截屏（可指定区域）→ 本地 RapidOCR → 返回带屏幕坐标的文本项。"""
    region = None
    if p.get("region_x") is not None and p.get("region_y") is not None:
        rx, ry = int(p["region_x"]), int(p["region_y"])
        region = (rx, ry, rx + int(p.get("region_w") or 0), ry + int(p.get("region_h") or 0))
    img = _grab_screen(region)
    path = os.path.join(os.environ.get("TEMP", "."), "screen_ocr_%s.png" % _ts())
    img.save(path, "PNG")
    try:
        engine = _get_rapidocr()
        result, _elapse = engine(path)
    finally:
        try:
            os.remove(path)
        except Exception:
            pass
    ox, oy = (region[0], region[1]) if region else (0, 0)
    items = []
    for line in result or []:
        try:
            box, text, conf = _parse_ocr_line(line)
        except Exception:
            continue
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        items.append({"text": text, "confidence": conf,
                      "x": int(min(xs)) + ox, "y": int(min(ys)) + oy,
                      "w": int(max(xs) - min(xs)), "h": int(max(ys) - min(ys))})
    return items


def m_computer_screen_ocr(args):
    p = _tool_params(args)
    items = _ocr_screen_items(p)
    return {"success": True, "text": "\n".join(it["text"] for it in items), "items": items}


def m_computer_describe(args):
    return _not_impl("computer_describe", "AI 视觉模型")


def _find_text_items(items, targets):
    for t in targets:
        t = t.strip()
        if not t:
            continue
        for it in items:
            if t.lower() in str(it["text"]).lower():
                return t, it
    return None, None


def m_computer_locate_text(args):
    p = _tool_params(args)
    text = str(p.get("text") or "").strip()
    if not text:
        raise RuntimeError("缺少 text 参数")
    hit, item = _find_text_items(_ocr_screen_items(p), [text])
    if item is None:
        return {"success": True, "found": False, "text": text}
    return {"success": True, "found": True, "text": text,
            "x": item["x"] + item["w"] // 2, "y": item["y"] + item["h"] // 2, "item": item}


def m_computer_click_text(args):
    p = _tool_params(args)
    texts = str(p.get("texts") or "")
    candidates = [t for t in re.split(r"[,，]", texts) if t.strip()]
    if not candidates:
        raise RuntimeError("缺少 texts 参数")
    hit, item = _find_text_items(_ocr_screen_items(p), candidates)
    if item is None:
        return {"success": False, "error": "屏幕上未找到任何候选文字: " + texts}
    import ctypes
    cx, cy = item["x"] + item["w"] // 2, item["y"] + item["h"] // 2
    ctypes.windll.user32.SetCursorPos(cx, cy)
    ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
    return {"success": True, "clicked": hit, "x": cx, "y": cy}


def m_computer_wait_text(args):
    p = _tool_params(args)
    text = str(p.get("text") or "").strip()
    if not text:
        raise RuntimeError("缺少 text 参数")
    timeout_ms = int(p.get("timeout_ms") or 10000)
    deadline = time.time() + timeout_ms / 1000.0
    while time.time() < deadline:
        hit, item = _find_text_items(_ocr_screen_items(p), [text])
        if item is not None:
            return {"success": True, "found": True, "text": text,
                    "x": item["x"] + item["w"] // 2, "y": item["y"] + item["h"] // 2}
        time.sleep(1.0)
    return {"success": True, "found": False, "text": text, "timeoutMs": timeout_ms}


def m_computer_assert_text(args):
    p = _tool_params(args)
    text = str(p.get("text") or "").strip()
    if not text:
        raise RuntimeError("缺少 text 参数")
    expect = str(p.get("expect") or "present").strip().lower()
    hit, item = _find_text_items(_ocr_screen_items(p), [text])
    found = item is not None
    ok = found if expect == "present" else (not found)
    return {"success": ok, "found": found, "expect": expect,
            "message": ("已出现: %s" % text) if found else ("未出现: %s" % text)}


# ===================== PPT 本地工具 =====================

def _ppt_themes():
    fpath = os.path.join(CONFIG_DIR, "ppt_templates.json")
    with open(fpath, "r", encoding="utf-8-sig") as f:
        return json.load(f)


def m_ppt_get_templates(args):
    data = _ppt_themes()
    themes = [{"id": t.get("id"), "name": t.get("name"), "font": t.get("font")}
              for t in data.get("themes", [])]
    return {"success": True, "defaultTheme": data.get("defaultTheme"), "themes": themes}


def _ppt_parse_data(p):
    headers = p.get("headers")
    if isinstance(headers, str):
        headers = json.loads(headers)
    rows = p.get("rows")
    if isinstance(rows, str):
        rows = json.loads(rows)
    return headers or [], rows or []


def m_ppt_analyze_data(args):
    p = _tool_params(args)
    headers, rows = _ppt_parse_data(p)
    if not headers:
        raise RuntimeError("缺少 headers 参数（可先用 excel_read_range 读取选区）")
    # 列画像：数值列 / 日期时间列 / 类别列
    profile = []
    for i, h in enumerate(headers):
        vals = [r[i] if isinstance(r, list) and i < len(r) else (r.get(h) if isinstance(r, dict) else None)
                for r in rows]
        nums = [v for v in vals if isinstance(v, (int, float)) and not isinstance(v, bool)]
        profile.append({"column": str(h), "type": "number" if len(nums) >= max(1, len(vals) // 2)
                        else ("category" if len(set(map(str, vals))) <= max(2, len(vals) // 3) else "text"),
                        "numericCount": len(nums)})
    metric_cols = [c["column"] for c in profile if c["type"] == "number"]
    cat_cols = [c["column"] for c in profile if c["type"] == "category"]
    slides = [
        {"type": "cover", "title": str(p.get("reportTitle") or "数据分析报告")},
        {"type": "kpi", "title": "核心指标", "metrics": [{"label": c, "value": "auto"} for c in metric_cols[:4]]},
        {"type": "chart", "title": "趋势与对比", "chartType": "column",
         "category": (cat_cols[0] if cat_cols else str(headers[0])),
         "series": metric_cols[:2]},
        {"type": "table", "title": "数据明细"},
        {"type": "closing", "title": "谢谢观看"},
    ]
    return {"success": True, "profile": profile, "slides": slides,
            "insights": ["建议按「封面 → 核心指标 → 图表 → 明细 → 结尾」结构组织"],
            "orientation": "landscape"}


def m_ppt_generate(args):
    """用 python-pptx 生成 .pptx（核心页型：cover/text/bullets/agenda/section/table/chart/kpi/closing）。
    完整版设计引擎（SmartArt/时间线/动画等）在 VSTO 宿主内，此处为网页端等效实现。"""
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.dml.color import RGBColor

    p = _tool_params(args)
    themes = _ppt_themes()
    theme_id = str(p.get("theme") or themes.get("defaultTheme") or "")
    theme = next((t for t in themes.get("themes", []) if t.get("id") == theme_id), None)
    colors = (theme or {}).get("colors") or {}
    primary = RGBColor(*colors.get("primary", [0, 90, 156]))
    text_color = RGBColor(*colors.get("text", [51, 51, 51]))
    font = (theme or {}).get("font") or "微软雅黑"

    headers, rows = _ppt_parse_data(p)
    slides = p.get("slides")
    if (not slides) and (headers or rows):
        slides = m_ppt_analyze_data(args).get("slides")
    if not slides:
        slides = [{"type": "cover", "title": str(p.get("reportTitle") or "演示文稿")},
                  {"type": "closing", "title": "谢谢观看"}]

    prs = Presentation()
    W, H = prs.slide_width, prs.slide_height

    def _bg(slide, rgb):
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = rgb

    def _add_title(slide, title):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), W / 914400 - 1.2, Inches(0.9))
        tf = tb.text_frame
        tf.text = str(title)
        r = tf.paragraphs[0].runs[0]
        r.font.size = Pt(28)
        r.font.bold = True
        r.font.color.rgb = primary
        r.font.name = font

    def _chart_type(t):
        return {"column": XL_CHART_TYPE.COLUMN_CLUSTERED, "bar": XL_CHART_TYPE.BAR_CLUSTERED,
                "line": XL_CHART_TYPE.LINE, "pie": XL_CHART_TYPE.PIE,
                "area": XL_CHART_TYPE.AREA, "doughnut": XL_CHART_TYPE.DOUGHNUT}.get(t, XL_CHART_TYPE.COLUMN_CLUSTERED)

    for spec in slides:
        st = str(spec.get("type") or "text").lower()
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式
        _bg(slide, RGBColor(*colors.get("bg", [255, 255, 255])))
        if st in ("title", "cover"):
            tb = slide.shapes.add_textbox(Inches(1), H / 914400 / 2 - 1, W / 914400 - 2, Inches(2))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.text = str(spec.get("title") or p.get("reportTitle") or "")
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(40)
            r.font.bold = True
            r.font.color.rgb = primary
            r.font.name = font
        elif st in ("closing", "end"):
            tb = slide.shapes.add_textbox(Inches(1), H / 914400 / 2 - 1, W / 914400 - 2, Inches(2))
            tf = tb.text_frame
            tf.word_wrap = True
            tf.text = str(spec.get("title") or "谢谢观看")
            r = tf.paragraphs[0].runs[0]
            r.font.size = Pt(36)
            r.font.bold = True
            r.font.color.rgb = primary
            r.font.name = font
        elif st in ("agenda", "section", "text", "bullets"):
            _add_title(slide, spec.get("title") or ("目录" if st == "agenda" else "内容"))
            items = spec.get("items") or spec.get("bullets") or spec.get("points") or []
            if isinstance(items, str):
                items = [ln.strip() for ln in items.splitlines() if ln.strip()]
            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), W / 914400 - 1.6, H / 914400 - 2.2)
            tf = tb.text_frame
            tf.word_wrap = True
            for i, it in enumerate(items):
                para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                para.text = "• " + str(it)
                para.runs[0].font.size = Pt(18)
                para.runs[0].font.color.rgb = text_color
                para.runs[0].font.name = font
        elif st == "kpi":
            _add_title(slide, spec.get("title") or "核心指标")
            metrics = spec.get("metrics") or []
            n = max(1, len(metrics))
            for i, m in enumerate(metrics[:6]):
                label = str(m.get("label") or "")
                val = m.get("value")
                if (val is None or val == "auto") and headers and rows:
                    idx = headers.index(label) if label in [str(h) for h in headers] else None
                    if idx is not None:
                        nums = [r[idx] for r in rows if isinstance(r, list) and idx < len(r)
                                and isinstance(r[idx], (int, float))]
                        val = round(sum(nums), 2) if nums else "-"
                tb = slide.shapes.add_textbox(Inches(0.6 + i * (W / 914400 - 1.2) / n),
                                              Inches(2.2), Inches((W / 914400 - 1.2) / n), Inches(1.6))
                tf = tb.text_frame
                tf.word_wrap = True
                tf.text = label
                pr = tf.paragraphs[0].runs[0]
                pr.font.size = Pt(14)
                pr.font.color.rgb = text_color
                para2 = tf.add_paragraph()
                para2.text = str(val if val is not None else "-")
                para2.runs[0].font.size = Pt(28)
                para2.runs[0].font.bold = True
                para2.runs[0].font.color.rgb = primary
        elif st == "table":
            _add_title(slide, spec.get("title") or "数据明细")
            if headers:
                nr, nc = min(len(rows) + 1, 12), len(headers)
                gfx = slide.shapes.add_table(nr, nc, Inches(0.6), Inches(1.5),
                                             W / 914400 - 1.2, Inches(0.4 * nr))
                tbl = gfx.table
                for j, h in enumerate(headers[:nc]):
                    tbl.cell(0, j).text = str(h)
                for i, r in enumerate(rows[:nr - 1]):
                    for j, v in enumerate(r[:nc] if isinstance(r, list) else []):
                        tbl.cell(i + 1, j).text = str(v)
        elif st == "chart":
            _add_title(slide, spec.get("title") or "图表")
            cat_col = spec.get("category")
            series = spec.get("series") or []
            if headers and rows:
                cd = CategoryChartData()
                ci = headers.index(cat_col) if cat_col in [str(h) for h in headers] else 0
                cats = [str(r[ci]) for r in rows if isinstance(r, list) and ci < len(r)]
                cd.categories = cats
                hdrs = [str(h) for h in headers]
                if not series:
                    # 未指定系列时自动推导数值列
                    series = [h for h in hdrs
                              if any(isinstance(r[hdrs.index(h)], (int, float))
                                     for r in rows[:20] if isinstance(r, list) and hdrs.index(h) < len(r))
                              and h != hdrs[ci]]
                for s in series[:4]:
                    if s in hdrs:
                        idx = headers.index(s)
                        vals = [r[idx] if isinstance(r, list) and idx < len(r)
                                and isinstance(r[idx], (int, float)) else 0 for r in rows]
                        cd.add_series(s, vals)
                slide.shapes.add_chart(_chart_type(spec.get("chartType")), Inches(0.8), Inches(1.5),
                                       W / 914400 - 1.6, H / 914400 - 2.2, cd)

    title = str(p.get("reportTitle") or "演示文稿").strip() or "演示文稿"
    desk = os.path.join(os.environ.get("USERPROFILE", "."), "Desktop")
    if not os.path.isdir(desk):
        desk = os.path.abspath(".")
    fname = os.path.join(desk, "%s_%s.pptx" % (re.sub(r'[\\/:*?"<>|]', "_", title), _ts()))
    prs.save(fname)
    try:
        os.startfile(fname)
    except Exception:
        pass
    return {"success": True, "filePath": fname, "slideCount": len(prs.slides._sldIdLst),
            "theme": theme_id or None}


def _ppt_not_impl(tool):
    return {"success": False, "notImplemented": True,
            "error": "工具 %s 的完整实现（SmartArt/动画/主题精调等设计引擎）在 VSTO 宿主内，"
                     "网页端(本地桥)暂不支持；核心生成请用 ppt_generate。" % tool}


# ===================== 路由 =====================

ROUTES = {
    "GetSelectedRange": m_get_selected_range,
    "GetSelectedRangeAsync": m_get_selected_range,
    "WriteToRange": m_write_to_range,
    "WriteToRangeAsync": m_write_to_range,
    "GetWorkbookInfo": m_get_workbook_info,
    "GetActiveSheetName": m_get_active_sheet_name,
    "GetOpenWorkbooks": m_get_open_workbooks,
    "ActivateWorkbook": m_activate_workbook,
    "ExportToCsv": m_export_csv,
    "excel_export_csv": m_export_csv,
    "ExecuteVBA": m_execute_vba,
    "ExecuteVBAAsync": m_execute_vba,
    "CheckPythonEnvironment": m_check_python_environment,
    "ExecutePython": m_execute_python,
    "GetCloudOcrConfig": m_get_cloud_ocr_config,
    "SaveCloudOcrConfig": m_save_cloud_ocr_config,
    "TestCloudOcrConfig": m_test_cloud_ocr_config,
    "ExecuteSkill": m_execute_skill,
    "ExecuteSkillAsync": m_execute_skill,
    "CallLocalTool": m_call_local_tool,
    "CallLocalToolAsync": m_call_local_tool,
    "DetectEnvironment": m_detect_environment,
    "GetAllSkills": m_get_all_skills,
    "GetAvailableSkills": m_get_available_skills,
    "ToggleSkillEnabled": m_toggle_skill_enabled,
    "SaveSkill": m_save_skill,
    "DeleteSkill": m_delete_skill,
    "GetSkillByName": m_get_skill_by_name,
    "GetMcpStatus": m_get_mcp_status,
    "StartMcpServer": m_start_mcp_server,
    "StopMcpServer": m_stop_mcp_server,
    # 数据插入（对齐 VB 桥）
    "InsertTableData": m_insert_table_data,
    "InsertDataSmart": m_insert_data_smart,
    "InsertDataSmartAsync": m_insert_data_smart,
    "InsertParsedData": m_insert_parsed_data,
    "InsertParsedDataAsync": m_insert_parsed_data,
    "FillDataToExcel": m_fill_data_to_excel,
    "FillDataToExcelAsync": m_fill_data_to_excel,
    # 本地 OCR（复用项目 Services/rapidocr_inference_engine）
    "SaveBase64Image": m_save_base64_image,
    "StartOcrRecognition": m_start_ocr_recognition,
    "StartOcrRecognitionWithFormat": m_start_ocr_recognition,
    "GetOcrResult": m_get_ocr_result,
    "GetOcrResultWithFormat": m_get_ocr_result,
    "GetOcrStructuredResult": m_get_ocr_structured_result,
    "GetCurrentOcrEngine": m_get_current_ocr_engine,
    "SetOcrEngine": m_set_ocr_engine,
    "GetOcrEngineStatus": m_get_ocr_engine_status,
    # 轻量查询 / 状态
    "GetActiveSheetName": m_get_active_sheet_name,
    "GetActiveWorkbookName": m_get_active_workbook_name,
    "GetSheets": m_get_sheets,
    "ActivateSheet": m_activate_sheet,
    "SelectSheet": m_activate_sheet,
    "ActivateSheetAsync": m_activate_sheet,
    "GetSelectedRangeData": m_get_selected_range_data,
    # 选区（对齐 VSTO 桥：引用区域管理器/实时选区依赖）
    "GetSelection": m_get_selection,
    "GetMultiSelection": m_get_multi_selection,
    "StartSelectionMonitor": m_start_selection_monitor,
    "StopSelectionMonitor": m_stop_selection_monitor,
    "GetRangeData": m_get_range_data,
    "GetUndoStack": m_get_undo_stack,
    "GetRedoStack": m_get_redo_stack,
}

# time / web 本地工具
ROUTES.update({
    "time_now": m_time_now,
    "web_search": m_web_search,
})

# computer 本地工具（Windows 自动化；OCR 类复用本地 RapidOCR）
ROUTES.update({
    "computer_screen_capture": m_computer_screen_capture,
    "computer_screen_get_info": m_computer_screen_get_info,
    "computer_screen_ocr": m_computer_screen_ocr,
    "computer_describe": m_computer_describe,
    "computer_mouse_move": m_computer_mouse_move,
    "computer_mouse_click": m_computer_mouse_click,
    "computer_mouse_drag": m_computer_mouse_drag,
    "computer_mouse_scroll": m_computer_mouse_scroll,
    "computer_key_press": m_computer_key_press,
    "computer_key_hotkey": m_computer_key_hotkey,
    "computer_type_text": m_computer_type_text,
    "computer_locate_text": m_computer_locate_text,
    "computer_locate_color": m_computer_locate_color,
    "computer_click_text": m_computer_click_text,
    "computer_clipboard_read": m_computer_clipboard_read,
    "computer_clipboard_write": m_computer_clipboard_write,
    "computer_open_application": m_computer_open_application,
    "computer_execute_command": m_computer_execute_command,
    "computer_wait": m_computer_wait,
    "computer_set_safe_mode": m_computer_set_safe_mode,
    "computer_mouse_position_get": m_computer_mouse_position_get,
    "computer_window_list": m_computer_window_list,
    "computer_process_list": m_computer_process_list,
    "computer_process_kill": m_computer_process_kill,
    "computer_close_application": m_computer_close_application,
    "computer_play_media": m_computer_play_media,
    "computer_wait_text": m_computer_wait_text,
    "computer_assert_text": m_computer_assert_text,
    "computer_focus_window": m_computer_focus_window,
    "computer_capture_window": m_computer_capture_window,
})

# PPT 本地工具
ROUTES.update({
    "ppt_get_templates": m_ppt_get_templates,
    "ppt_analyze_data": m_ppt_analyze_data,
    "ppt_generate": m_ppt_generate,
})


# ===================== License 桥方法（桌面版离线激活，路线①） =====================
# 与 Hosts/Shared/HostLicense.vb、Bridges/LicenseKey.vb、LicenseServer/lib/licenseKey.js
# 逐字节一致（改算法必须多处同步）：
#   key = "SPK1-" & base64url("v1|login|plan|seats|expUnix|nonce") & "." & base64url(HMAC-SHA256)
# 会话与插件版共用 %APPDATA%\SmartPilotOffice\license_session.json；
# activation 经 DPAPI（CurrentUser，无熵）保护，pywin32 的 win32crypt 与 .NET ProtectedData 互操作。
# 桌面版只支持离线激活码；在线订阅（登录/注册/支付）仍由插件版 + 授权服务器提供。

_LIC_KEY_PREFIX = "SPK1-"
_LIC_DEACT_PREFIX = "SPD1-"
_LIC_KEY_VERSION = "v1"
_LIC_KEY_FIELD_COUNT = 6
_LIC_KEY_SECRET = b"SmartPilot-Offline-License-Key-v1-CHANGE-ME"
_LIC_SESSION_PATH = os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"),
    "SmartPilotOffice", "license_session.json")
_LIC_FILE_LOCK = threading.Lock()


def _lic_b64url_encode(data):
    return base64.b64encode(data).decode("ascii").replace("+", "-").replace("/", "_").replace("=", "")


def _lic_b64url_decode(text):
    s = (text or "").replace("-", "+").replace("_", "/")
    if len(s) % 4 == 2:
        s += "=="
    elif len(s) % 4 == 3:
        s += "="
    return base64.b64decode(s)


def _lic_sign(payload_text):
    import hashlib
    import hmac
    return _lic_b64url_encode(
        hmac.new(_LIC_KEY_SECRET, payload_text.encode("utf-8"), hashlib.sha256).digest())


def _lic_key_hash(key):
    import hashlib
    return hashlib.sha256((key or "").encode("utf-8")).hexdigest()[:16]


def _lic_protect(text):
    """DPAPI 保护（与 .NET ProtectedData.CurrentUser 无熵互操作）；不可用返回空。"""
    try:
        import win32crypt
        blob = win32crypt.CryptProtectData(
            text.encode("utf-8"), "SmartPilot", None, None, None, 0)
        return base64.b64encode(blob).decode("ascii")
    except Exception:
        return ""


def _lic_unprotect(b64_text):
    try:
        import win32crypt
        blob = base64.b64decode(b64_text)
        raw = win32crypt.CryptUnprotectData(blob, None, None, None, 0)[1]
        return raw.decode("utf-8")
    except Exception:
        return ""


def _lic_read_session():
    try:
        with open(_LIC_SESSION_PATH, "r", encoding="utf-8-sig") as f:
            jo = json.loads(f.read() or "{}")
        if not isinstance(jo, dict):
            jo = {}
    except Exception:
        jo = {}
    act = ""
    prot = jo.get("activationProtected") or ""
    if prot:
        act = _lic_unprotect(prot)
    if not act:
        act = jo.get("activation") or ""
    jo["activation"] = act
    jo["activationSet"] = bool(jo.get("activationSet", False))
    jo["deactivated"] = jo.get("deactivated") or ""
    return jo


def _lic_write_session(jo):
    d = os.path.dirname(_LIC_SESSION_PATH)
    if d and not os.path.isdir(d):
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
    out = dict(jo)
    act = out.pop("activation", "") or ""
    if act:
        prot = _lic_protect(act)
        if prot:
            out["activationProtected"] = prot
            out["activation"] = ""
        else:
            out["activation"] = act  # DPAPI 不可用时明文兜底（只影响跨端共享）
    else:
        out["activationProtected"] = ""
        out["activation"] = ""
    with _LIC_FILE_LOCK:
        for _attempt in range(8):
            try:
                with open(_LIC_SESSION_PATH, "w", encoding="utf-8") as f:
                    json.dump(out, f, ensure_ascii=False, indent=2)
                return
            except Exception:
                time.sleep(0.2)


def _lic_deactivated_list(session):
    out = []
    for part in (session.get("deactivated") or "").split(","):
        part = part.strip()
        if part:
            out.append(part)
    return out


def _lic_verify_key(key):
    import datetime as _dt
    raw = (key or "").strip()
    if not raw.startswith(_LIC_KEY_PREFIX):
        return {"valid": False, "reason": "malformed"}
    body = raw[len(_LIC_KEY_PREFIX):]
    dot = body.rfind(".")
    if dot <= 0 or dot == len(body) - 1:
        return {"valid": False, "reason": "malformed"}
    try:
        payload_text = _lic_b64url_decode(body[:dot]).decode("utf-8")
    except Exception:
        return {"valid": False, "reason": "malformed"}
    sig = body[dot + 1:]
    if sig != _lic_sign(payload_text):
        return {"valid": False, "reason": "bad_signature"}
    parts = payload_text.split("|")
    if len(parts) != _LIC_KEY_FIELD_COUNT or parts[0] != _LIC_KEY_VERSION:
        return {"valid": False, "reason": "bad_payload"}
    try:
        exp_unix = int(parts[4])
    except Exception:
        return {"valid": False, "reason": "bad_payload"}
    if exp_unix <= 0:
        return {"valid": False, "reason": "bad_payload"}
    period_end = _dt.datetime(1970, 1, 1) + _dt.timedelta(seconds=exp_unix)
    if period_end <= _dt.datetime.utcnow():
        return {"valid": False, "reason": "expired"}
    try:
        seats = int(parts[3])
    except Exception:
        seats = 1
    if seats < 1:
        seats = 1
    return {"valid": True, "reason": "ok",
            "login": "" if parts[1] == "-" else parts[1],
            "plan": parts[2], "seats": seats, "period_end": period_end}


def _lic_first_param(args):
    """license.js 的 call() 把 body 作为单个 JSON 字符串/对象传入。"""
    if not args:
        return {}
    first = args[0]
    if isinstance(first, dict):
        return first
    if isinstance(first, str):
        try:
            v = json.loads(first)
            return v if isinstance(v, dict) else {}
        except Exception:
            return {}
    return {}


def _lic_entitlement_of(result):
    pe = result.get("period_end")
    return {
        "plan": result.get("plan", "none"),
        "seats": result.get("seats", 0),
        "status": "active" if result.get("valid") else "inactive",
        "periodEnd": pe.isoformat() if pe else None,
        "active": bool(result.get("valid")),
        "source": "offline",
    }


def _lic_inactive_local(reason):
    return json.dumps({
        "success": True, "active": False, "source": "offline", "reason": reason,
        "entitlement": {"plan": "none", "status": "inactive", "active": False,
                        "periodEnd": None, "source": "offline", "reason": reason}
    }, ensure_ascii=False)


def m_license_get_local_activation(args):
    """读取本地离线激活状态。本方法为【只读】：验证失败或已注销时只返回
    inactive，绝不清空 activation——否则用户打开页面（只读）就会被静默注销。"""
    with _LIC_FILE_LOCK:
        session = _lic_read_session()
    key = (session.get("activation") or "").strip()
    if not key:
        return _lic_inactive_local("not_activated")
    result = _lic_verify_key(key)
    if not result["valid"]:
        return _lic_inactive_local(result["reason"])
    if _lic_key_hash(key) in _lic_deactivated_list(session):
        return _lic_inactive_local("deactivated")
    return json.dumps({
        "success": True, "active": True, "source": "offline", "login": result["login"],
        "entitlement": _lic_entitlement_of(result),
    }, ensure_ascii=False)


def m_license_activate_key(args):
    p = _lic_first_param(args)
    key = (p.get("key") or "").strip()
    result = _lic_verify_key(key)
    if not result["valid"]:
        reason_text = {"expired": "激活码已过期",
                       "bad_signature": "激活码无效（签名不匹配）"}.get(result["reason"], "激活码格式不正确")
        return json.dumps({"success": False, "error": reason_text, "code": "invalid_key"},
                          ensure_ascii=False)
    with _LIC_FILE_LOCK:
        session = _lic_read_session()
        if _lic_key_hash(key) in _lic_deactivated_list(session):
            return json.dumps({"success": False,
                               "error": "此激活码已在本机注销（已申请退款），不能再次激活",
                               "code": "deactivated"}, ensure_ascii=False)
        session["activation"] = key
        session["activationSet"] = True
        _lic_write_session(session)
    return json.dumps({
        "success": True, "active": True, "source": "offline",
        "entitlement": _lic_entitlement_of(result),
    }, ensure_ascii=False)


def m_license_clear_local_activation(args):
    with _LIC_FILE_LOCK:
        session = _lic_read_session()
        session["activation"] = ""
        session["activationSet"] = False
        _lic_write_session(session)
    return json.dumps({"success": True}, ensure_ascii=False)


def m_license_deactivate_for_refund(args):
    import datetime as _dt
    import random
    with _LIC_FILE_LOCK:
        session = _lic_read_session()
        key = (session.get("activation") or "").strip()
        if not key:
            return json.dumps({"success": False, "error": "当前没有离线激活记录，无需注销",
                               "code": "not_activated"}, ensure_ascii=False)
        result = _lic_verify_key(key)
        plan = result.get("plan") or "pro_yearly"
        login = result.get("login") or ""
        h = _lic_key_hash(key)
        lst = _lic_deactivated_list(session)
        if h not in lst:
            lst.append(h)
        session["deactivated"] = ",".join(lst)
        session["activation"] = ""
        session["activationSet"] = False
        _lic_write_session(session)
    ts = int((_dt.datetime.utcnow() - _dt.datetime(1970, 1, 1)).total_seconds())
    nonce = "".join(random.choice("0123456789abcdef") for _ in range(16))
    payload = "|".join([_LIC_KEY_VERSION, h, plan, str(ts), nonce])
    proof = _LIC_DEACT_PREFIX + _lic_b64url_encode(payload.encode("utf-8")) + "." + _lic_sign(payload)
    return json.dumps({"success": True, "code": proof, "plan": plan, "login": login,
                       "destroyedAt": _dt.datetime.utcnow().isoformat()}, ensure_ascii=False)


def m_license_get_entitlement(args):
    return json.dumps({"success": True, "loggedIn": False,
                       "entitlement": {"plan": "none", "status": "inactive", "active": False,
                                       "periodEnd": None, "reason": "not_logged_in"}},
                      ensure_ascii=False)


def m_license_is_subscription_active(args):
    local = json.loads(m_license_get_local_activation(args))
    if local.get("active") is True:
        return json.dumps({"success": True, "active": True, "source": "offline"},
                          ensure_ascii=False)
    return json.dumps({"success": True, "active": False, "loggedIn": False},
                      ensure_ascii=False)


def m_license_get_api_base_url(args):
    return json.dumps({"success": True, "baseUrl": "http://localhost:9530"}, ensure_ascii=False)


def m_license_ping_health(args):
    return json.dumps({"success": True, "ok": False, "stripeReady": False,
                       "code": "server_down"}, ensure_ascii=False)


def m_license_logout(args):
    return json.dumps({"success": True}, ensure_ascii=False)


def _lic_desktop_only(args):
    return json.dumps({"success": False,
                       "error": "桌面版暂不支持在线账户功能，请使用离线激活码（SPK1- 开头）激活",
                       "code": "desktop_offline_only"}, ensure_ascii=False)


# ---- AI 模型配置（与插件版共用 %APPDATA%\ExcelAIVSTOAddIn\Config\llm_config.json）----
# 插件版保存 API Key 后会导出明文到该文件（ConfigBridge.ExportLLMConfigForStandalone，
# 字段 provider/apiType/model/apiUrl/baseUrl/apiKey，UTF-8 无 BOM）。
# 桌面版读写同一份，因此两个版本配置一次、两边通用。

_LLM_CONFIG_DIR = os.path.join(
    os.environ.get("APPDATA") or os.path.expanduser("~"), "ExcelAIVSTOAddIn", "Config")
_LLM_CONFIG_PATH = os.path.join(_LLM_CONFIG_DIR, "llm_config.json")


def _llm_read():
    try:
        with open(_LLM_CONFIG_PATH, "r", encoding="utf-8-sig") as f:
            jo = json.loads(f.read() or "{}")
        return jo if isinstance(jo, dict) else {}
    except Exception:
        return {}


def _llm_write(cfg):
    try:
        if not os.path.isdir(_LLM_CONFIG_DIR):
            os.makedirs(_LLM_CONFIG_DIR, exist_ok=True)
        with open(_LLM_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False


def _llm_mask(v):
    v = v or ""
    if not v:
        return ""
    if len(v) <= 8:
        return "*" * len(v)
    return v[:4] + "****" + v[-4:]


def _llm_apply(cfg, incoming):
    """与 ConfigBridge.SaveLLMConfig 同构：字段名兼容，空 apiKey 表示不覆盖。"""
    provider = incoming.get("provider") or incoming.get("apiType")
    if provider:
        cfg["provider"] = provider
        cfg["apiType"] = provider
    if incoming.get("model"):
        cfg["model"] = incoming["model"]
    api_url = incoming.get("apiUrl") or incoming.get("baseUrl")
    if api_url:
        cfg["apiUrl"] = api_url
        cfg["baseUrl"] = api_url
    if incoming.get("apiKey"):
        cfg["apiKey"] = incoming["apiKey"]
    return cfg


def m_save_ai_config(args):
    p = _lic_first_param(args)
    cfg = _llm_apply(_llm_read(), p)
    if not _llm_write(cfg):
        return json.dumps({"success": False,
                           "error": "写入本机 AI 配置失败: " + _LLM_CONFIG_PATH},
                          ensure_ascii=False)
    return json.dumps({"success": True}, ensure_ascii=False)


def m_get_ai_config(args):
    cfg = _llm_read()
    return json.dumps({"success": True, "config": {
        "provider": cfg.get("provider", ""),
        "apiType": cfg.get("apiType", cfg.get("provider", "")),
        "model": cfg.get("model", ""),
        "apiUrl": cfg.get("apiUrl", cfg.get("baseUrl", "")),
        "baseUrl": cfg.get("baseUrl", cfg.get("apiUrl", "")),
        "apiKey": cfg.get("apiKey", ""),
    }}, ensure_ascii=False)


def m_get_llm_config(args):
    cfg = _llm_read()
    ak = cfg.get("apiKey") or ""
    return json.dumps({"provider": cfg.get("provider", "ollama"),
                       "model": cfg.get("model", ""),
                       "apiKey": _llm_mask(ak),
                       "apiUrl": cfg.get("apiUrl", ""),
                       "hasApiKey": bool(ak)}, ensure_ascii=False)


def m_save_llm_config(args):
    return m_save_ai_config(args)


ROUTES.update({
    "SaveAIConfig": m_save_ai_config,
    "GetAIConfig": m_get_ai_config,
    "SaveLLMConfig": m_save_llm_config,
    "GetLLMConfig": m_get_llm_config,
})

ROUTES.update({
    "LicenseActivateKey": m_license_activate_key,
    "LicenseGetLocalActivation": m_license_get_local_activation,
    "LicenseClearLocalActivation": m_license_clear_local_activation,
    "LicenseDeactivateForRefund": m_license_deactivate_for_refund,
    "LicenseGetEntitlement": m_license_get_entitlement,
    "LicenseIsSubscriptionActive": m_license_is_subscription_active,
    "LicenseGetApiBaseUrl": m_license_get_api_base_url,
    "LicensePingHealth": m_license_ping_health,
    "LicenseRegister": _lic_desktop_only,
    "LicenseLogin": _lic_desktop_only,
    "LicenseLogout": m_license_logout,
    "LicenseStartCheckout": _lic_desktop_only,
    "LicenseOpenBillingPortal": _lic_desktop_only,
})


# ===================== MCP 本地工具（excel-local） =====================
# 前端 app-mcp.js: loadMcpToolsForAI() 经 GET /api/mcp/tools 拉工具清单注入 LLM；
# 工具执行经 /api/excel/invoke（根 /invoke）以 method=工具名 分发到下列实现。
CONFIG_DIR = _resolve_shared_dir([
    os.path.join(_PLUGIN_INSTALL, "Excel", "Config"),          # 插件版安装目录
    os.path.join(_PLUGIN_APPDATA, "Config"),                   # 插件版 AppData 部署
    os.path.normpath(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'Config')),
])
_MCP_TOOLS_CACHE = None


# ===================== OcrHost 兼容协议（/v1/call + /health） =====================
# 插件版由 OcrHost.exe(9530) 提供独立识图服务；桌面版桥在同协议上模拟，
# ocr-panel.html 从 9527 加载时走同源。解析脚本数据与插件版共用
# %APPDATA%\ExcelAIVSTOAddIn\ParseScripts\*.md（归一，两端同一份）。
_PARSE_SCRIPTS_DIR = os.path.join(_SKILL_APPDATA, "ParseScripts")
_PARSE_TEMPLATES_DIR = os.path.join(CONFIG_DIR, "DefaultParseScripts")
_PARSE_SCRIPTS_CACHE = None


def _ensure_parse_templates():
    """首次把内置模板缓存到共享 ParseScripts 目录（对齐 VSTO EnsureBuiltInTemplates）。"""
    try:
        if not os.path.isdir(_PARSE_SCRIPTS_DIR):
            os.makedirs(_PARSE_SCRIPTS_DIR, exist_ok=True)
        if os.path.isdir(_PARSE_TEMPLATES_DIR):
            for fn in os.listdir(_PARSE_TEMPLATES_DIR):
                if not fn.lower().endswith(".md"):
                    continue
                dest = os.path.join(_PARSE_SCRIPTS_DIR, fn)
                if not os.path.isfile(dest):
                    with open(os.path.join(_PARSE_TEMPLATES_DIR, fn), "rb") as f:
                        data = f.read()
                    with open(dest, "wb") as f:
                        f.write(data)
    except Exception:
        pass


def _builtin_parse_script_ids():
    ids = set()
    try:
        if os.path.isdir(_PARSE_TEMPLATES_DIR):
            for fn in os.listdir(_PARSE_TEMPLATES_DIR):
                if not fn.lower().endswith(".md"):
                    continue
                try:
                    with open(os.path.join(_PARSE_TEMPLATES_DIR, fn), "r", encoding="utf-8-sig") as f:
                        s = _parse_script_from_md(f.read())
                    if s and s.get("id"):
                        ids.add(str(s["id"]))
                except Exception:
                    continue
    except Exception:
        pass
    return ids


def _parse_script_from_md(text):
    """解析 VSTO ParseScriptManager 的 md 格式：frontmatter + 字段表 + js 块 + 示例块。"""
    s = {"id": "", "name": "", "description": "", "icon": "📝", "category": "通用",
         "outputFormat": "table", "tags": [], "fields": [], "script": "", "sampleText": "",
         "isBuiltIn": False}
    lines = text.replace("\ufeff", "").splitlines()
    if lines and lines[0].strip() == "---":
        i = 1
        while i < len(lines) and lines[i].strip() != "---":
            m = re.match(r'^([A-Za-z_][\w-]*)\s*:\s*(.*)$', lines[i])
            if m:
                k, v = m.group(1), m.group(2).strip()
                if k == "tags":
                    s["tags"] = [t.strip() for t in v.split(",") if t.strip()]
                elif k in s:
                    s[k] = v
            i += 1
        lines = lines[i + 1:]
    section = None
    for ln in lines:
        st = ln.strip()
        if st.startswith("## "):
            section = st[3:].strip()
            continue
        if section == "字段定义" and st.startswith("|"):
            cells = [c.strip() for c in st.strip("|").split("|")]
            if len(cells) >= 5 and cells[1] and cells[1] != "键名" and not set(cells[1]) <= set("-"):
                pat = cells[2]
                if pat.startswith("`") and pat.endswith("`") and len(pat) >= 2:
                    pat = pat[1:-1]
                if pat == "-":
                    pat = ""
                s["fields"].append({
                    "name": cells[0], "key": cells[1], "pattern": pat,
                    "required": cells[3] == "是", "visible": cells[4] != "否",
                    "defaultValue": (cells[5] if len(cells) > 5 else "") or "",
                })
        elif section == "解析脚本":
            if not st.startswith("```"):
                s["script"] += ln + "\n"
        elif section == "示例文本":
            if not st.startswith("```"):
                s["sampleText"] += ln + "\n"
    s["script"] = s["script"].strip("\n")
    s["sampleText"] = s["sampleText"].strip("\n")
    return s if s.get("id") else None


def _parse_script_to_md(s):
    out = ["---",
           "id: %s" % s.get("id", ""),
           "name: %s" % (s.get("name") or ""),
           "description: %s" % (s.get("description") or ""),
           "icon: %s" % (s.get("icon") or "📝"),
           "category: %s" % (s.get("category") or "通用"),
           "outputFormat: %s" % (s.get("outputFormat") or "table")]
    tags = s.get("tags") or []
    if tags:
        out.append("tags: " + ", ".join(str(t) for t in tags))
    out += ["---", "", "## 字段定义", "",
            "| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |",
            "|--------|------|-----------|------|------|"]
    for f in (s.get("fields") or []):
        pat = (f.get("pattern") or "").strip()
        pat = ("`%s`" % pat) if pat else "-"
        out.append("| %s | %s | %s | %s | %s |" % (
            f.get("name", ""), f.get("key", ""), pat,
            "是" if f.get("required") else "否",
            "否" if f.get("visible") is False else "是"))
    out += ["", "## 解析脚本", "", "```javascript", s.get("script") or "", "```", ""]
    if s.get("sampleText"):
        out += ["## 示例文本", "", "```", s["sampleText"], "```", ""]
    return "\n".join(out) + "\n"


def _load_parse_scripts():
    global _PARSE_SCRIPTS_CACHE
    if _PARSE_SCRIPTS_CACHE is not None:
        return _PARSE_SCRIPTS_CACHE
    _ensure_parse_templates()
    scripts = {}
    try:
        if os.path.isdir(_PARSE_SCRIPTS_DIR):
            for fn in sorted(os.listdir(_PARSE_SCRIPTS_DIR)):
                if not fn.lower().endswith(".md"):
                    continue
                try:
                    with open(os.path.join(_PARSE_SCRIPTS_DIR, fn), "r", encoding="utf-8-sig") as f:
                        s = _parse_script_from_md(f.read())
                    if s:
                        scripts[str(s["id"])] = s
                except Exception:
                    continue
    except Exception:
        pass
    builtin_ids = _builtin_parse_script_ids()
    for sid, s in scripts.items():
        s["isBuiltIn"] = sid in builtin_ids
    _PARSE_SCRIPTS_CACHE = scripts
    return _PARSE_SCRIPTS_CACHE


def _ocrh_get_parse_scripts(args):
    lst = sorted(_load_parse_scripts().values(), key=lambda x: str(x.get("name") or ""))
    return {"success": True, "scripts": lst, "count": len(lst)}


def _ocrh_get_parse_script(args):
    sid = str(args[0]) if args else ""
    s = _load_parse_scripts().get(sid)
    if not s:
        return {"success": False, "error": "找不到指定的解析脚本"}
    return {"success": True, "script": s}


def _ocrh_save_parse_script(args):
    raw = args[0] if args else None
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
    except Exception as e:
        return {"success": False, "error": "JSON 解析失败: " + str(e)}
    sid = str((data or {}).get("id") or "").strip()
    if not sid:
        return {"success": False, "error": "缺少 id"}
    try:
        if not os.path.isdir(_PARSE_SCRIPTS_DIR):
            os.makedirs(_PARSE_SCRIPTS_DIR, exist_ok=True)
        with open(os.path.join(_PARSE_SCRIPTS_DIR, _safe_skill_filename(sid) + ".md"),
                  "w", encoding="utf-8") as f:
            f.write(_parse_script_to_md(data))
    except Exception as e:
        return {"success": False, "error": str(e)}
    global _PARSE_SCRIPTS_CACHE
    _PARSE_SCRIPTS_CACHE = None
    return {"success": True, "id": sid}


def _ocrh_delete_parse_script(args):
    sid = str(args[0]) if args else ""
    s = _load_parse_scripts().get(sid)
    if s and s.get("isBuiltIn"):
        return {"success": False, "error": "内置模板不可删除"}
    try:
        p = os.path.join(_PARSE_SCRIPTS_DIR, _safe_skill_filename(sid) + ".md")
        if os.path.isfile(p):
            os.remove(p)
    except Exception as e:
        return {"success": False, "error": str(e)}
    global _PARSE_SCRIPTS_CACHE
    _PARSE_SCRIPTS_CACHE = None
    return {"success": True, "id": sid}


def _ocrh_duplicate_parse_script(args):
    sid = str(args[0]) if args else ""
    new_name = str(args[1]) if len(args) > 1 else ""
    src = _load_parse_scripts().get(sid)
    if not src:
        return {"success": False, "error": "找不到指定的解析脚本"}
    import copy as _copy
    cp = _copy.deepcopy(src)
    cp["id"] = uuid.uuid4().hex
    cp["name"] = new_name or (str(src.get("name") or "脚本") + " 副本")
    cp["isBuiltIn"] = False
    try:
        if not os.path.isdir(_PARSE_SCRIPTS_DIR):
            os.makedirs(_PARSE_SCRIPTS_DIR, exist_ok=True)
        with open(os.path.join(_PARSE_SCRIPTS_DIR, cp["id"] + ".md"), "w", encoding="utf-8") as f:
            f.write(_parse_script_to_md(cp))
    except Exception as e:
        return {"success": False, "error": str(e)}
    global _PARSE_SCRIPTS_CACHE
    _PARSE_SCRIPTS_CACHE = None
    return {"success": True, "script": cp}


def _ocrh_reset_builtin_script(args):
    sid = str(args[0]) if args else ""
    try:
        if os.path.isdir(_PARSE_TEMPLATES_DIR):
            for fn in os.listdir(_PARSE_TEMPLATES_DIR):
                if not fn.lower().endswith(".md"):
                    continue
                with open(os.path.join(_PARSE_TEMPLATES_DIR, fn), "r", encoding="utf-8-sig") as f:
                    s = _parse_script_from_md(f.read())
                if s and str(s.get("id")) == sid:
                    if not os.path.isdir(_PARSE_SCRIPTS_DIR):
                        os.makedirs(_PARSE_SCRIPTS_DIR, exist_ok=True)
                    with open(os.path.join(_PARSE_SCRIPTS_DIR, _safe_skill_filename(sid) + ".md"),
                              "w", encoding="utf-8") as f:
                        f.write(_parse_script_to_md(s))
                    global _PARSE_SCRIPTS_CACHE
                    _PARSE_SCRIPTS_CACHE = None
                    return {"success": True, "script": s}
    except Exception as e:
        return {"success": False, "error": str(e)}
    return {"success": False, "error": "找不到内置模板: " + sid}


def _ocrh_clear_parse_script_cache(args):
    global _PARSE_SCRIPTS_CACHE
    _PARSE_SCRIPTS_CACHE = None
    _load_parse_scripts()
    return {"success": True}


def _ocrh_import_parse_script_from_markdown(args):
    md = args[0] if args else ""
    s = _parse_script_from_md(md or "")
    if not s:
        return {"success": False, "error": "Markdown 格式无效（缺少 id）"}
    try:
        if not os.path.isdir(_PARSE_SCRIPTS_DIR):
            os.makedirs(_PARSE_SCRIPTS_DIR, exist_ok=True)
        with open(os.path.join(_PARSE_SCRIPTS_DIR, _safe_skill_filename(s["id"]) + ".md"),
                  "w", encoding="utf-8") as f:
            f.write(_parse_script_to_md(s))
    except Exception as e:
        return {"success": False, "error": str(e)}
    global _PARSE_SCRIPTS_CACHE
    _PARSE_SCRIPTS_CACHE = None
    return {"success": True, "script": s}


def _ocrh_get_ocr_elements_json(args):
    return {"success": False, "error": "桌面版暂不支持 OCR 坐标元素"}


def _ocrh_get_parse_script_categories(args):
    cats = sorted({str(s.get("category") or "通用")
                   for s in _load_parse_scripts().values()})
    return {"success": True, "categories": cats}


_PADDLE_LANGS = [{"id": "ch", "name": "中文"}, {"id": "en", "name": "英文"},
                 {"id": "japan", "name": "日语"}, {"id": "korean", "name": "韩语"},
                 {"id": "fr", "name": "法语"}, {"id": "ger", "name": "德语"},
                 {"id": "es", "name": "西班牙语"}, {"id": "ru", "name": "俄语"},
                 {"id": "arabic", "name": "阿拉伯语"}, {"id": "structure", "name": "自动检测"}]


def _ocrh_get_paddle_advanced_config(args):
    return {"success": True, "language": "ch", "enableLayoutAnalysis": True,
            "enableTableRecognition": True, "enableFormulaRecognition": False,
            "supportedLanguages": _PADDLE_LANGS}


def _ocrh_save_paddle_advanced_config(args):
    return {"success": True}


def _ocrh_get_ocr_raw(args):
    """对齐 OcrBridge.GetOcrRawData：{ready, success, raw_data:[{text,confidence,box}]}"""
    with _ocr_lock:
        done = bool(_ocr_state.get("completed"))
        ok = bool(_ocr_state.get("success"))
        raw = _ocr_state.get("items")
    if not done:
        return {"ready": False}
    return {"ready": True, "success": ok, "raw_data": raw or []}


# ==================== 自定义 PPT 主题（与 VSTO 共用 custom_ppt_themes.json） ====================
# 存 %APPDATA%\ExcelAIVSTOAddIn\custom_ppt_themes.json，与 VSTO 的 PptGenerationService.Theme.vb
# SaveCustomTheme 同一文件（PathHelper.ConfigDirectory），桌面版/插件版数据互通。

def _custom_ppt_theme_path():
    return os.path.join(os.environ.get("APPDATA", "."), "ExcelAIVSTOAddIn", "custom_ppt_themes.json")


def m_save_custom_ppt_theme(args):
    raw = args[0] if args else None
    try:
        t = json.loads(raw) if isinstance(raw, str) else (raw or {})
    except Exception as e:
        return {"success": False, "error": "JSON 解析失败: " + str(e)}
    if not isinstance(t, dict):
        return {"success": False, "error": "主题数据格式错误"}
    tid = str(t.get("id") or "").strip() or ("custom_" + uuid.uuid4().hex)
    t["id"] = tid
    path = _custom_ppt_theme_path()
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        root = {}
        if os.path.isfile(path):
            try:
                root = json.load(open(path, encoding="utf-8"))
            except Exception:
                root = {}
        if not isinstance(root.get("themes"), list):
            root["themes"] = []
        replaced = False
        for i, x in enumerate(root["themes"]):
            if str((x or {}).get("id")) == tid:
                root["themes"][i] = t
                replaced = True
                break
        if not replaced:
            root["themes"].append(t)
        root["version"] = 1
        json.dump(root, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    except Exception as e:
        return {"success": False, "error": str(e)}
    return {"success": True, "id": tid}


def m_list_custom_ppt_themes(args):
    path = _custom_ppt_theme_path()
    if os.path.isfile(path):
        try:
            root = json.load(open(path, encoding="utf-8"))
            return {"success": True, "themes": root.get("themes") or []}
        except Exception as e:
            return {"success": False, "error": str(e)}
    return {"success": True, "themes": []}


def m_delete_custom_ppt_theme(args):
    tid = str(args[0]) if args else ""
    path = _custom_ppt_theme_path()
    if os.path.isfile(path):
        try:
            root = json.load(open(path, encoding="utf-8"))
            root["themes"] = [x for x in (root.get("themes") or [])
                              if str((x or {}).get("id")) != tid]
            json.dump(root, open(path, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
        except Exception as e:
            return {"success": False, "error": str(e)}
    return {"success": True, "id": tid}


# 归一：OcrHost 兼容方法统一注册进 ROUTES，与桌面版桥共用同一张方法表（单一事实来源）。
# 放在所有 _ocrh_* 实现函数之后，避免模块加载时前向引用未定义符号。
ROUTES.update({
    "GetOcrRawData": _ocrh_get_ocr_raw,
    "GetOcrElementsJson": _ocrh_get_ocr_elements_json,
    "GetParseScriptCategories": _ocrh_get_parse_script_categories,
    "GetParseScripts": _ocrh_get_parse_scripts,
    "GetParseScript": _ocrh_get_parse_script,
    "SaveParseScript": _ocrh_save_parse_script,
    "DeleteParseScript": _ocrh_delete_parse_script,
    "DuplicateParseScript": _ocrh_duplicate_parse_script,
    "ResetBuiltinScript": _ocrh_reset_builtin_script,
    "ClearParseScriptCache": _ocrh_clear_parse_script_cache,
    "ImportParseScriptFromMarkdown": _ocrh_import_parse_script_from_markdown,
    "GetPaddleOcrAdvancedConfig": _ocrh_get_paddle_advanced_config,
    "SavePaddleOcrAdvancedConfig": _ocrh_save_paddle_advanced_config,
    "SetPaddleOcrAdvancedConfig": _ocrh_save_paddle_advanced_config,
    "SaveCustomPptTheme": m_save_custom_ppt_theme,
    "ListCustomPptThemes": m_list_custom_ppt_themes,
    "DeleteCustomPptTheme": m_delete_custom_ppt_theme,
})


def _tool_def(name, desc, params, required=None):
    """与 excel_tools.json 一致的工具定义形态。params: {名: (类型, 描述)}"""
    req = set(required or [])
    return {"name": name, "description": desc, "displayName": name,
            "parameters": [{"name": k, "type": t, "description": d, "required": k in req}
                           for k, (t, d) in (params or {}).items()]}


# computer 组定义：镜像 VSTO ComputerAgentTools.GetToolDefinitions()
COMPUTER_TOOL_DEFS = [
    _tool_def("computer_screen_capture", "截图返回 base64 或保存路径",
              {"x": ("integer", "左上X，可选"), "y": ("integer", "左上Y，可选"),
               "width": ("integer", "宽"), "height": ("integer", "高"),
               "returnFormat": ("string", "base64 或 path")}),
    _tool_def("computer_screen_get_info", "获取屏幕分辨率信息", {}),
    _tool_def("computer_screen_ocr", "截图+OCR返回识别文字（本地 RapidOCR）", {}),
    _tool_def("computer_describe", "AI视觉描述屏幕内容（需 AI 后端，网页端不支持）",
              {"prompt": ("string", "描述指令")}),
    _tool_def("computer_mouse_move", "移动鼠标到指定坐标",
              {"x": ("integer", "X"), "y": ("integer", "Y")}, ["x", "y"]),
    _tool_def("computer_mouse_click", "点击鼠标按钮",
              {"x": ("integer", "X"), "y": ("integer", "Y"),
               "button": ("string", "left/right/middle"), "clicks": ("integer", "次数")}),
    _tool_def("computer_mouse_drag", "拖动鼠标从起点到终点",
              {"fromX": ("integer", "起点X"), "fromY": ("integer", "起点Y"),
               "toX": ("integer", "终点X"), "toY": ("integer", "终点Y")},
              ["fromX", "fromY", "toX", "toY"]),
    _tool_def("computer_mouse_scroll", "鼠标滚轮滚动",
              {"amount": ("integer", "量")}, ["amount"]),
    _tool_def("computer_key_press", "按下并释放一个键",
              {"key": ("string", "键名如 Enter/Tab/F1")}, ["key"]),
    _tool_def("computer_key_hotkey", "组合键",
              {"keys": ("string", "如 Ctrl+C")}, ["keys"]),
    _tool_def("computer_type_text", "输入文字",
              {"text": ("string", "文字")}, ["text"]),
    _tool_def("computer_locate_text", "OCR定位屏幕上指定文字（本地 RapidOCR）",
              {"text": ("string", "搜索目标")}, ["text"]),
    _tool_def("computer_locate_color", "定位屏幕上指定颜色",
              {"color": ("string", "#RRGGBB"), "tolerance": ("integer", "容差0-100")}, ["color"]),
    _tool_def("computer_click_text", "OCR定位并点击屏幕上指定文字（候选词逗号分隔按优先级）",
              {"texts": ("string", "候选词，逗号分隔")}, ["texts"]),
    _tool_def("computer_clipboard_read", "读取剪贴板文字", {}),
    _tool_def("computer_clipboard_write", "写入剪贴板文字",
              {"text": ("string", "文字")}, ["text"]),
    _tool_def("computer_open_application", "打开应用程序或文件（传 app 别名如 QQ音乐/浏览器/记事本，或 path 完整路径）",
              {"app": ("string", "应用名/别名，如 QQ音乐、浏览器、记事本"),
               "path": ("string", "完整路径（优先于 app）"),
               "arguments": ("string", "启动参数")}),
    _tool_def("computer_execute_command", "执行命令行命令",
              {"command": ("string", "命令")}, ["command"]),
    _tool_def("computer_wait", "等待毫秒数",
              {"ms": ("integer", "毫秒")}, ["ms"]),
    _tool_def("computer_set_safe_mode", "切换安全模式",
              {"enabled": ("boolean", "true/false")}, ["enabled"]),
    _tool_def("computer_mouse_position_get", "获取当前鼠标光标坐标", {}),
    _tool_def("computer_window_list", "列出当前可见顶层窗口标题",
              {"keyword": ("string", "按标题关键词过滤，可选")}),
    _tool_def("computer_process_list", "列出运行中的进程",
              {"name": ("string", "按进程名过滤，可选")}),
    _tool_def("computer_process_kill", "结束指定进程",
              {"name": ("string", "进程名如 notepad/notepad.exe")}, ["name"]),
    _tool_def("computer_close_application", "按应用名（口语/别名）关闭程序",
              {"app": ("string", "应用名/别名，如 音乐、浏览器、QQ音乐")}, ["app"]),
    _tool_def("computer_play_media", "聚焦应用并触发播放（需 OCR 验证，网页端为简化实现）",
              {"app": ("string", "应用名/别名，如 QQ音乐、音乐")}, ["app"]),
    _tool_def("computer_wait_text", "等待屏幕出现指定文字（轮询OCR直到超时或命中）",
              {"text": ("string", "要等待出现的文字"), "timeout_ms": ("integer", "超时毫秒")}, ["text"]),
    _tool_def("computer_assert_text", "校验屏幕是否出现/消失指定文字（返回found）",
              {"text": ("string", "待校验文字"), "expect": ("string", "present 或 absent")}, ["text"]),
    _tool_def("computer_focus_window", "按应用名/别名将已打开窗口切换置前",
              {"app": ("string", "应用名/别名，如 QQ音乐、浏览器、计算器")}, ["app"]),
    _tool_def("computer_capture_window", "按应用名/别名截图其窗口区域",
              {"app": ("string", "应用名/别名"), "returnFormat": ("string", "base64 或 path")}, ["app"]),
]

# web/time 组定义：镜像 VSTO WebSearchTools / TimeTools
WEB_TIME_TOOL_DEFS = [
    _tool_def("web_search",
              "联网搜索网页知识并返回结构化结果标题/网址/摘要，适合查资料、查资讯、陌生概念等时效性信息",
              {"query": ("string", "搜索关键词（必填）"),
               "maxResults": ("integer", "返回结果数，默认 5，范围 1~10"),
               "region": ("string", "搜索地域代码，可选，如 cn-zh（中文）/ us-en（英文）")}, ["query"]),
    _tool_def("time_now",
              "获取当前的系统日期、时间、星期、时区与工作日状态，用于回答今天几号/几点/周几等问题",
              {"withSeconds": ("boolean", "是否包含秒，默认 false")}),
]


def load_mcp_tools():
    """读取 Config 工具定义，附加 source 供前端 buildMcpToolMeta 使用。
    对齐 VSTO McpCapabilityRegistry：excel-local / ppt-local / computer-local / web-local / time-local。"""
    global _MCP_TOOLS_CACHE
    if _MCP_TOOLS_CACHE is not None:
        return _MCP_TOOLS_CACHE
    tools = []
    fpath = os.path.join(CONFIG_DIR, "excel_tools.json")
    try:
        with open(fpath, "r", encoding="utf-8-sig") as f:
            data = json.load(f)
        for t in data.get("tools", []):
            t = dict(t)
            t["source"] = "excel-local"
            tools.append(t)
    except Exception as e:
        sys.stderr.write("[Bridge] 加载工具定义失败 %s: %s\n" % (fpath, e))
    # PPT 组定义（与 VSTO 共用 Config/ppt_tools.json）
    ppath = os.path.join(CONFIG_DIR, "ppt_tools.json")
    try:
        with open(ppath, "r", encoding="utf-8-sig") as f:
            pdata = json.load(f)
        for t in pdata.get("tools", []):
            t = dict(t)
            t["source"] = "ppt-local"
            tools.append(t)
    except Exception as e:
        sys.stderr.write("[Bridge] 加载PPT工具定义失败 %s: %s\n" % (ppath, e))
    # computer / web / time 组定义（镜像 VSTO 内置定义）
    for t in COMPUTER_TOOL_DEFS:
        t = dict(t)
        t["source"] = "computer-local"
        tools.append(t)
    for t in WEB_TIME_TOOL_DEFS:
        t = dict(t)
        t["source"] = "web-local" if t["name"].startswith("web_") else "time-local"
        tools.append(t)
    _MCP_TOOLS_CACHE = tools
    return tools


def m_get_available_mcp_tools(args):
    tools = load_mcp_tools()
    return {"success": True, "tools": tools, "count": len(tools)}


def _as_bool(v, default):
    if v is None or v == "":
        return default
    if isinstance(v, bool):
        return v
    return str(v).strip().lower() in ("true", "1", "yes", "y", "是")


def _tool_params(args):
    return args[0] if args and isinstance(args[0], dict) else {}


def _target_sheet(app, wb, sheet_name):
    if sheet_name:
        try:
            return wb.Sheets(sheet_name)
        except Exception:
            raise RuntimeError("找不到工作表: " + str(sheet_name))
    return wb.ActiveSheet


def _range_grid(rng, use_formula=False):
    vals = to_jsonable(rng.Formula if use_formula else rng.Value2)
    if vals is None:
        return [[]]
    if not isinstance(vals, list):
        return [[vals]]
    if vals and not isinstance(vals[0], list):
        return [vals]
    return vals


def m_excel_read_range(args):
    """excel_read_range：读取区域数据。range 为空时按 当前选区 → UsedRange 回退。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    if app is None:
        raise RuntimeError("未连接到 Excel")
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        try:
            sel = app.Selection
            if sel is not None and str(sel.Worksheet.Name) == str(ws.Name):
                range_addr = str(sel.Address)
        except Exception:
            pass
        if not range_addr:
            try:
                range_addr = str(ws.UsedRange.Address)
            except Exception:
                raise RuntimeError("无法确定读取范围（无选区且无 UsedRange）")
    rng = ws.Range(range_addr)
    max_rows = int(p.get("maxRows") or 10000)
    max_cols = int(p.get("maxColumns") or 100)
    has_headers = _as_bool(p.get("hasHeaders"), True)
    return_format = str(p.get("returnFormat") or "array").strip().lower()
    skip_empty = _as_bool(p.get("skipEmptyRows"), False)

    grid = _range_grid(rng, _as_bool(p.get("readFormulas"), False))
    total_rows = len(grid)
    truncated = False
    if total_rows > max_rows:
        grid = grid[:max_rows]
        truncated = True
    if grid and grid[0] and len(grid[0]) > max_cols:
        grid = [r[:max_cols] for r in grid]
        truncated = True

    def _row_empty(r):
        return all(c is None or (isinstance(c, str) and c.strip() == "") for c in r)

    if skip_empty:
        grid = [r for r in grid if not _row_empty(r)]

    headers = []
    data_rows = grid
    if has_headers and grid:
        headers = ["" if c is None else str(c) for c in grid[0]]
        data_rows = grid[1:]

    if return_format == "object" and headers:
        out_rows = [dict(zip(headers, r)) for r in data_rows]
    else:
        out_rows = data_rows

    return {
        "success": True,
        "range": str(rng.Address),
        "worksheet": str(ws.Name),
        "workbook": str(wb.Name),
        "headers": headers,
        "rows": out_rows,
        "rowCount": len(data_rows),
        "colCount": (len(grid[0]) if grid and grid[0] else 0),
        "totalRows": total_rows,
        "truncated": truncated,
        "returnFormat": return_format,
    }


def m_excel_get_cell_value(args):
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    cell = str(p.get("cell") or p.get("range") or p.get("address") or "").strip()
    if not cell:
        raise RuntimeError("缺少 cell/range 参数")
    rng = ws.Range(cell)
    return {
        "success": True,
        "value": to_jsonable(rng.Value2),
        "address": str(rng.Address),
        "worksheet": str(ws.Name),
    }


def m_excel_list_sheets(args):
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    names = [str(wb.Sheets.Item(i).Name) for i in range(1, int(wb.Sheets.Count) + 1)]
    return {"success": True, "workbook": str(wb.Name), "sheets": names,
            "active": str(app.ActiveSheet.Name), "count": len(names)}


def m_excel_write_range(args):
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or p.get("address") or "").strip()
    if not range_addr:
        raise RuntimeError("缺少 range 参数")
    values = p.get("values", p.get("data", p.get("rows")))
    if isinstance(values, str):
        try:
            values = json.loads(values)
        except Exception:
            raise RuntimeError("values 不是合法的 JSON")
    if values is None:
        raise RuntimeError("缺少 values 数据")
    if not isinstance(values, list):
        values = [[values]]
    elif values and not isinstance(values[0], list):
        values = [values]
    rows_n = len(values)
    cols_n = len(values[0]) if rows_n else 0
    # 注意：本环境对多单元格区域批量 .Value = 二维数组会静默写坏（pywin32/SAFEARRAY 兼容问题）；
    # 且 .Resize()/.Offset()/.Columns() 等索引访问方法返回的 Range 代理会报 -2146827284。
    # 故只用直接 Range 与 .Cells() 构造范围并逐格赋值，最稳定。
    tl = ws.Range(range_addr).Cells(1, 1)
    br = ws.Range(range_addr).Cells(rows_n, cols_n)
    out = ws.Range(tl, br)
    for i, row in enumerate(values):
        for j, v in enumerate(row):
            out.Cells(i + 1, j + 1).Value = ("" if v is None else v)
    return {"success": True, "written": str(out.Address),
            "worksheet": str(ws.Name), "rowCount": rows_n, "colCount": cols_n}


def m_excel_clear_range(args):
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or p.get("address") or "").strip()
    if not range_addr:
        try:
            sel = app.Selection
            if sel is not None and str(sel.Worksheet.Name) == str(ws.Name):
                range_addr = str(sel.Address)
        except Exception:
            pass
    if not range_addr:
        raise RuntimeError("缺少 range 参数（且无当前选区可复用）")
    rng = ws.Range(range_addr)
    rng.ClearContents()
    return {"success": True, "cleared": str(rng.Address), "worksheet": str(ws.Name)}


def m_excel_create_sheet(args):
    p = _tool_params(args)
    name = str(p.get("name") or "").strip()
    if not name:
        raise RuntimeError("缺少 name 参数")
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    for i in range(1, int(wb.Sheets.Count) + 1):
        if str(wb.Sheets.Item(i).Name) == name:
            return {"success": True, "created": False, "message": "工作表已存在: " + name, "worksheet": name}
    ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
    ws.Name = name
    return {"success": True, "created": True, "worksheet": name}


def _as_obj(v):
    """format/alignment 参数兼容 dict 与 JSON 字符串。"""
    if isinstance(v, dict):
        return v
    if isinstance(v, str) and v.strip():
        try:
            parsed = json.loads(v)
            return parsed if isinstance(parsed, dict) else {}
        except Exception:
            return {}
    return {}


def _ole_color(v, default=None):
    """'RRGGBB' → Excel OLE_COLOR (BGR 整数)。"""
    s = str(v or "").strip().lstrip("#")
    if len(s) == 3:
        s = "".join(c * 2 for c in s)
    if len(s) != 6:
        return default
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
    except ValueError:
        return default
    return r + g * 256 + b * 65536


_BORDER_LINE = {"dotted": -4118, "dashed": -4114, "double": -4115, "slantDashDot": 5}
_BORDER_WEIGHT = {"hair": 1, "thin": 2, "medium": -4138, "thick": 5}
_ALIGN_H = {"left": -4131, "center": -4108, "right": -4152, "justify": -4130}
_ALIGN_V = {"top": -4160, "middle": -4108, "center": -4108, "bottom": -4107}


def m_excel_format_range(args):
    """excel_format_range：字体/底色/边框/对齐/数字格式/条件格式，仅改样式不改数据。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        try:
            sel = app.Selection
            if sel is not None and str(sel.Worksheet.Name) == str(ws.Name):
                range_addr = str(sel.Address)
        except Exception:
            pass
    if not range_addr:
        raise RuntimeError("缺少 range 参数（且无当前选区可复用）")
    rng = ws.Range(range_addr)

    fmt = _as_obj(p.get("format"))
    align = _as_obj(p.get("alignment")) or _as_obj(fmt.get("alignment"))
    applied = []

    font = rng.Font
    if "bold" in fmt:
        font.Bold = _as_bool(fmt.get("bold"), False)
        applied.append("bold")
    if "italic" in fmt:
        font.Italic = _as_bool(fmt.get("italic"), False)
        applied.append("italic")
    if fmt.get("fontSize"):
        font.Size = int(fmt.get("fontSize"))
        applied.append("fontSize")
    if fmt.get("fontName"):
        font.Name = str(fmt.get("fontName"))
        applied.append("fontName")
    c = _ole_color(fmt.get("fontColor"))
    if c is not None:
        font.Color = c
        applied.append("fontColor")
    c = _ole_color(fmt.get("bgColor"))
    if c is not None:
        rng.Interior.Color = c
        applied.append("bgColor")
    if fmt.get("numberFormat"):
        rng.NumberFormat = str(fmt.get("numberFormat"))
        applied.append("numberFormat")

    # 边框（六条边 + 内部线，不含对角线）
    border_style = str(fmt.get("borderStyle") or "").strip().lower()
    if border_style and border_style != "none":
        line = _BORDER_LINE.get(border_style, 1)
        weight = _BORDER_WEIGHT.get(border_style, 2)
        color = _ole_color(fmt.get("borderColor"), 0)
        for bid in (7, 8, 9, 10, 11, 12):  # xlEdgeLeft/Top/Bottom/Right/InsideV/InsideH
            b = rng.Borders(bid)
            b.LineStyle = line
            b.Weight = weight
            b.Color = color
        applied.append("border")

    if align:
        if align.get("horizontal"):
            h = _ALIGN_H.get(str(align.get("horizontal")).strip().lower())
            if h is not None:
                rng.HorizontalAlignment = h
                applied.append("hAlign")
        if align.get("vertical"):
            v = _ALIGN_V.get(str(align.get("vertical")).strip().lower())
            if v is not None:
                rng.VerticalAlignment = v
                applied.append("vAlign")
        if "wrapText" in align:
            rng.WrapText = _as_bool(align.get("wrapText"), False)
            applied.append("wrapText")

    # 条件格式（可选）：dataBar / colorScale / iconSet
    cond = _as_obj(p.get("conditionalFormat"))
    cond_type = str(cond.get("type") or "").strip().lower()
    if cond_type:
        try:
            fcs = rng.FormatConditions
            if cond_type == "databar" or cond_type == "databar":
                fc = fcs.AddDatabar()
                c = _ole_color(cond.get("color"))
                if c is not None and hasattr(fc, "BarColor"):
                    fc.BarColor.Color = c
            elif cond_type == "colorscale":
                fcs.AddColorScale()
            elif cond_type == "iconset":
                fcs.AddIconSetCondition()
            applied.append("conditional:" + cond_type)
        except Exception as e:
            applied.append("conditional-failed:" + str(e))

    return {"success": True, "formatted": str(rng.Address), "worksheet": str(ws.Name),
            "applied": applied}


def m_excel_autofit(args):
    """excel_autofit：自动调整列宽（默认）或行高。range 缺省用 UsedRange。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        try:
            range_addr = str(ws.UsedRange.Address)
        except Exception:
            raise RuntimeError("无法确定调整范围（无 range 且无 UsedRange）")
    rng = ws.Range(range_addr)
    typ = str(p.get("type") or "column").strip().lower()
    if typ == "row":
        rng.Rows.AutoFit()
    else:
        rng.Columns.AutoFit()
    return {"success": True, "adjusted": str(rng.Address), "type": typ,
            "worksheet": str(ws.Name)}


def m_excel_save_workbook(args):
    """excel_save_workbook：无 filePath 保存原文件；有 filePath 另存为。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    fp = str(p.get("filePath") or "").strip()
    if fp:
        wb.SaveAs(fp)
        return {"success": True, "savedAs": fp, "workbook": str(wb.Name)}
    wb.Save()
    return {"success": True, "saved": str(wb.Name), "path": str(wb.FullName)}


# ---- 进阶 excel-local 工具：聚合/排序/筛选/复制/合并/图表/透视表 ----
# XlYesNoGuess：xlGuess=0 / xlNo=2 / xlYes=1
# ⚠️ 历史坑：旧值 XL_YES=-4163 / XL_NO=-4152 是从其它枚举（xlUp 等）抄错的，
#    传给 Sort.Header 直接 E_INVALIDARG（现代 Sort 对象路径必崩）；legacy rng.Sort 虽不报错，
#    但会被 Excel 当成 xlGuess —— 首行是数字时猜"无表头"，把表头行一起排进数据（静默错序）。
XL_YES = 1
XL_NO = 2
XL_ASCENDING = 1
XL_DESCENDING = 2
XL_CENTER = -4108


def _col_spec_to_index(header_row, spec):
    """列名/列字母(A,B,AA)/列序号 → 1-based 列号。
    注意：优先按表头名匹配（字段名如 'Cat' 会被误读为列字母），再退化为数字/字母。"""
    if spec is None:
        return None
    s = str(spec).strip()
    if not s:
        return None
    if header_row:
        for i, h in enumerate(header_row):
            if h is not None and str(h).strip().lower() == s.lower():
                return i + 1
    if s.isdigit():
        return int(s)
    if re.fullmatch(r'[A-Za-z]+', s):
        n = 0
        for ch in s.upper():
            n = n * 26 + (ord(ch) - 64)
        return n
    return None


def _as_list(v):
    if v is None:
        return []
    if isinstance(v, list):
        return v
    if isinstance(v, str) and v.strip():
        try:
            parsed = json.loads(v)
            return parsed if isinstance(parsed, list) else [v]
        except Exception:
            return [v]
    return []


def _write_table_to_sheet(wb, sheet_name, headers, rows):
    try:
        ws = wb.Sheets(sheet_name)
    except Exception:
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = sheet_name
    ncols = len(headers)
    for j, h in enumerate(headers):
        ws.Cells(1, j + 1).Value = ("" if h is None else h)
    for i, row in enumerate(rows):
        for j, v in enumerate(row):
            ws.Cells(i + 2, j + 1).Value = ("" if v is None else v)
    return ws


# ── excel_aggregate 专用：维度择优（2026-10-28）──
# 背景：源表「地区」列 20 行全是"华南"，模型按"地区"聚合后只有 1 组 →
# 图表只有一根柱/一个扇区，毫无信息量，而模型还会照常收尾说"已完成"。
# 旧逻辑只检查"列名能不能匹配上"，从不看数据分布，所以永远发现不了这种退化。
_AGG_NON_DIM_KEYS = ("日期", "时间", "年份", "年度", "月份", "编号", "序号", "编码", "id",
                     "销售额", "成本", "利润", "金额", "数量", "客户数", "单价", "价格", "收入", "业绩",
                     "备注", "说明", "地址", "电话", "邮箱")


def _agg_keyword_score(name):
    """表头"典型维度"关键词打分（与 VB 侧 KeywordDimScore 同口径）。"""
    n = str(name or "")
    for key, sc in (("部门", 100), ("地区", 90), ("区域", 88), ("大区", 87),
                    ("产品", 85), ("商品", 84), ("品类", 83), ("渠道", 80), ("来源", 79),
                    ("分类", 75), ("类别", 74), ("类型", 73), ("城市", 70),
                    ("省份", 68), ("国家", 65), ("客户", 60)):
        if key in n:
            return sc
    return 0


def _agg_is_numeric_col(data_rows, i):
    """第 i 列（0-based）是否数值列（Excel 序列日期本质是 float，一并排除出维度候选）。"""
    seen = num = 0
    for r in data_rows:
        if i >= len(r):
            continue
        v = r[i]
        if v is None or str(v).strip() == "":
            continue
        seen += 1
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            num += 1
    return seen > 0 and num >= seen * 0.8


def _agg_distinct(data_rows, i, cap=500):
    seen = set()
    for r in data_rows:
        if i >= len(r):
            continue
        v = r[i]
        if v is None or str(v).strip() == "":
            continue
        seen.add(str(v).strip())
        if len(seen) >= cap:
            return len(seen)
    return len(seen)


def _agg_dim_candidates(header_row, data_rows, val_idx_1based):
    """扫描所有分类列，给出 {index(1-based), column, distinct, score}，供择优与回传。"""
    out = []
    n_rows = len(data_rows)
    for i, h in enumerate(header_row):
        name = str(h or "").strip()
        if not name:
            continue
        if any(k.lower() in name.lower() for k in _AGG_NON_DIM_KEYS):
            continue
        if _agg_is_numeric_col(data_rows, i):
            continue
        if val_idx_1based is not None and (i + 1) == val_idx_1based:
            continue
        d = _agg_distinct(data_rows, i)
        if d <= 0:
            continue
        if d <= 1:
            score = -1
        else:
            base = 90 if d <= 20 else (60 if d <= 50 else 30)
            if n_rows > 0 and d >= n_rows:   # 每行唯一 = 明细列（姓名/流水号）
                base -= 20
            score = max(base, 0) * 2 + _agg_keyword_score(name)
        out.append({"index": i + 1, "column": name, "distinct": d, "score": score})
    return out


def _agg_group_rows(data_rows, dim_idx, val_idx, func, skip_empty=False):
    """按维度列（1-based）分组聚合，返回 [[key, value], ...]。"""
    groups = {}
    for r in data_rows:
        if max(dim_idx, val_idx) > len(r):
            continue
        dim_val = r[dim_idx - 1]
        if dim_val is None or (skip_empty and str(dim_val).strip() == ""):
            continue
        val = r[val_idx - 1]
        key = str(dim_val)
        g = groups.get(key)
        if g is None:
            g = groups[key] = {"sum": 0.0, "count": 0, "max": None, "min": None}
        if func in ("sum", "avg"):
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                g["sum"] += val
                g["count"] += 1
        elif func == "count":
            g["count"] += 1
        elif func == "max":
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                g["max"] = val if g["max"] is None else max(g["max"], val)
        elif func == "min":
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                g["min"] = val if g["min"] is None else min(g["min"], val)
    rows = []
    for key in groups:
        g = groups[key]
        if func == "sum":
            v = round(g["sum"], 6)
        elif func == "avg":
            v = round(g["sum"] / g["count"], 6) if g["count"] else 0
        elif func == "count":
            v = g["count"]
        elif func == "max":
            v = g["max"] if g["max"] is not None else 0
        else:
            v = g["min"] if g["min"] is not None else 0
        rows.append([key, v])
    return rows


def m_excel_aggregate(args):
    """excel_aggregate：按维度列分组聚合数值列（sum/avg/count/max/min）。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        try:
            sel = app.Selection
            if sel is not None and str(sel.Worksheet.Name) == str(ws.Name):
                range_addr = str(sel.Address)
        except Exception:
            pass
    if not range_addr:
        range_addr = str(ws.UsedRange.Address)
    rng = ws.Range(range_addr)
    grid = _range_grid(rng)
    if not grid:
        return {"success": True, "groups": [], "result": {"headers": ["维度", "聚合值"], "rows": []}}
    header_row = grid[0]
    data_rows = grid[1:]
    dim_spec = str(p.get("dimColumn") or "").strip()
    dim_idx = _col_spec_to_index(header_row, dim_spec)
    if dim_idx is None:
        raise RuntimeError("找不到分组维度列: " + str(dim_spec))
    val_spec = str(p.get("valueColumn") or "").strip()
    val_idx = _col_spec_to_index(header_row, val_spec) if val_spec else None
    if val_idx is None:
        for i, h in enumerate(header_row):
            col_data = [r[i] for r in data_rows if i < len(r)]
            if any(isinstance(x, (int, float)) and not isinstance(x, bool) for x in col_data):
                val_idx = i + 1
                val_spec = str(h)
                break
    if val_idx is None:
        raise RuntimeError("找不到可聚合的数值列")
    func = str(p.get("aggFunc") or "sum").strip().lower()
    func = {"sum": "sum", "avg": "avg", "average": "avg", "count": "count",
            "max": "max", "min": "min"}.get(func, "sum")
    skip_empty = _as_bool(p.get("skipEmptyRows"), False)
    out_rows = _agg_group_rows(data_rows, dim_idx, val_idx, func, skip_empty)
    # 【维度择优】结果只有 1 组 = 该维度列在本次数据里没有区分度 → 自动改用有区分度的列重算
    requested_dim = dim_spec
    dim_candidates = _agg_dim_candidates(header_row, data_rows, val_idx)
    dim_switched = False
    dim_switch_note = ""
    if len(out_rows) <= 1:
        best, best_score = None, 0
        for cand in dim_candidates:
            if cand["index"] == dim_idx or cand["distinct"] < 2:
                continue
            if cand["score"] > best_score:
                best_score, best = cand["score"], cand
        if best is not None:
            new_rows = _agg_group_rows(data_rows, best["index"], val_idx, func, skip_empty)
            if len(new_rows) > 1:
                dim_switch_note = ("维度列“%s”在本次数据中只有 %d 个取值，聚合结果没有区分度"
                                   "（图表只会有一根柱/一个扇区），已自动改用“%s”（%d 个取值）。"
                                   % (requested_dim, len(out_rows), best["column"], len(new_rows)))
                dim_spec = best["column"]
                dim_idx = best["index"]
                out_rows = new_rows
                dim_switched = True
    headers = [dim_spec, func + "(" + val_spec + ")"]
    target = str(p.get("targetSheet") or "").strip()
    if target:
        _write_table_to_sheet(wb, target, headers, out_rows)
    out = {"success": True, "dimColumn": dim_spec, "valueColumn": val_spec, "aggFunc": func,
           "groupCount": len(out_rows), "result": {"headers": headers, "rows": out_rows},
           "targetSheet": target or None}
    if dim_candidates:
        out["dimCandidates"] = dim_candidates
    if dim_switched:
        out["dimSwitched"] = True
        out["requestedDimColumn"] = requested_dim
        out["actualDimColumn"] = dim_spec
        out["dimSwitchNote"] = dim_switch_note
    return out


# ── excel_sort_range 专用：对齐 Excel 原生「排序」对话框（2026-10-16 重写）──
# 排序依据（SortOn）常量：本机实测 Excel 16.0 上 0=单元格值 1=单元格颜色 2=字体颜色 3=条件格式图标
_SORT_ON = {"values": 0, "cellcolor": 1, "fontcolor": 2, "icon": 3}
# 方向/排序方法（原生对话框「选项」）
_XL_SORT_COLUMNS, _XL_SORT_ROWS = 1, 2        # xlSortColumns=1 / xlSortRows=2
_XL_PINYIN, _XL_STROKE = 1, 2                  # xlPinYin=1 / xlStroke=2
# XlIconSet 枚举（值=集合序号，实测 Item(i).ID == i，故名称→序号直接用下表）
_XL_ICON_SETS = ["xl3Arrows", "xl3ArrowsGray", "xl3Flags", "xl3TrafficLights1",
                 "xl3TrafficLights2", "xl3Signs", "xl3Symbols", "xl3Symbols2",
                 "xl4Arrows", "xl4ArrowsGray", "xl4RedToBlack", "xl4CRV",
                 "xl4TrafficLights", "xl5Arrows", "xl5ArrowsGray", "xl5CRV",
                 "xl5Quarters", "xl3Stars", "xl3Triangles", "xl5Boxes"]
# 常见中文颜色名 → RGB 整数（与 Interior.Color/Font.Color 同一空间）
_CN_COLOR_RGB = {"红": 0xFF0000, "红色": 0xFF0000, "绿": 0x008000, "绿色": 0x008000,
                 "蓝": 0x0000FF, "蓝色": 0x0000FF, "黄": 0xFFFF00, "黄色": 0xFFFF00,
                 "橙": 0xFFA500, "橙色": 0xFFA500, "紫": 0x800080, "紫色": 0x800080,
                 "青": 0x008080, "青色": 0x008080, "黑": 0x000000, "黑色": 0x000000,
                 "白": 0xFFFFFF, "白色": 0xFFFFFF, "灰": 0x808080, "灰色": 0x808080}


def _as_list(v):
    """参数兼容 list / dict(单键) / JSON 数组字符串（_as_obj 只认 dict，排序键必须用它）。"""
    if isinstance(v, list):
        return v
    if isinstance(v, str) and v.strip():
        try:
            parsed = json.loads(v)
            return parsed if isinstance(parsed, list) else ([parsed] if isinstance(parsed, dict) else [])
        except Exception:
            return []
    return [v] if isinstance(v, dict) else []


def _parse_sort_color(v):
    """颜色参数 → Excel RGB 整数。支持 #RRGGBB / RRGGBB / 'r,g,b' / 十进制 / 中文颜色名。"""
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    if s in _CN_COLOR_RGB:
        return _CN_COLOR_RGB[s]
    hexpart = s[1:] if s.startswith("#") else s
    if re.fullmatch(r"[0-9a-fA-F]{6}", hexpart):
        return int(hexpart, 16)
    parts = [x.strip() for x in s.split(",")]
    if len(parts) == 3 and all(x.isdigit() for x in parts):
        r, g, b = (int(x) for x in parts)
        return (r << 16) | (g << 8) | b
    if s.isdigit():
        return int(s)
    return None


def _resolve_sort_icon(wb, set_raw, icon_raw):
    """解析条件格式图标：iconSet 支持序号(1基)或名称(3Arrows/xl3Arrows)，icon 为图标序号(1基)。"""
    set_s = str(set_raw or "").strip()
    icon_s = str(icon_raw or "").strip()
    if not set_s:
        raise RuntimeError("按条件格式图标排序缺少 iconSet（图标集序号或名称，如 3Arrows）")
    if not icon_s:
        raise RuntimeError("按条件格式图标排序缺少 icon（图标序号，从 1 开始）")
    if not icon_s.isdigit() or int(icon_s) < 1:
        raise RuntimeError("icon 参数『%s』不是合法序号（应为从 1 开始的整数）" % icon_s)
    icon_idx = int(icon_s)
    sets = wb.IconSets
    if set_s.isdigit():
        set_idx = int(set_s)
    else:
        cand = set_s if set_s.lower().startswith("xl") else "xl" + set_s
        set_idx = None
        for i, name in enumerate(_XL_ICON_SETS, start=1):
            if name.lower() == cand.lower():
                set_idx = i
                break
        if set_idx is None:
            raise RuntimeError("图标集『%s』不存在。可用：%s" % (set_s, "、".join(_XL_ICON_SETS)))
    iset = sets.Item(set_idx)
    if icon_idx > iset.Count:
        raise RuntimeError("图标序号 %d 超出图标集『%s』的范围（共 %d 个图标）" % (icon_idx, set_s, iset.Count))
    return iset.Item(icon_idx)


def m_excel_sort_range(args):
    """excel_sort_range：完整还原 Excel 原生「排序」对话框 —— 多级排序依据/次要关键字、
    排序依据（单元格值/单元格颜色/字体颜色/条件格式图标）、次序（升/降序）、
    选项（区分大小写 matchCase / 方向 orientation / 排序方法 sortMethod）。
    排序后回传表头+前 N 行预览，让 Agent 有真实数据可总结（不再输出占位空表格）。
    """
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        raise RuntimeError("缺少 range 参数")
    rng = ws.Range(range_addr)
    has_header = _as_bool(p.get("hasHeader"), True)
    header = XL_YES if has_header else XL_NO
    header_row = _range_grid(rng)[0] if has_header else None

    # ── 解析排序键（对齐原生对话框）──
    keys = []
    for item in _as_list(p.get("sortBy")):
        if not isinstance(item, dict):
            continue
        col = _col_spec_to_index(header_row, item.get("column"))
        if col is None:
            continue
        order = str(item.get("order") or "ascending").strip().lower()
        keys.append({
            "col": col,
            "order": XL_DESCENDING if order.startswith("desc") else XL_ASCENDING,
            "desc": order.startswith("desc"),
            "sortOn": str(item.get("sortOn") or "values").strip().lower().replace("_", ""),
            "color": item.get("color"),
            "iconSet": item.get("iconSet"),
            "icon": item.get("icon"),
        })
    if not keys:
        keys = [{"col": 1, "order": XL_ASCENDING, "desc": False,
                 "sortOn": "values", "color": None, "iconSet": None, "icon": None}]

    def _col_rng(idx):
        # 本环境 .Columns(idx) 返回的 Range 代理会失效，改用 .Cells 构造列范围作排序键
        return ws.Range(rng.Cells(1, idx), rng.Cells(rng.Rows.Count, idx))

    def _apply_via_sort_object():
        s = ws.Sort
        s.SortFields.Clear()
        # 选项三要素：每次显式重置（Sort 对象状态在 Worksheet 上跨次持久）
        s.MatchCase = _as_bool(p.get("matchCase"), False)
        orient = str(p.get("orientation") or "").strip().lower()
        s.Orientation = _XL_SORT_ROWS if orient in ("rows", "按行", "lefttoright") else _XL_SORT_COLUMNS
        method = str(p.get("sortMethod") or "").strip().lower()
        if method in ("stroke", "笔划", "笔画"):
            s.SortMethod = _XL_STROKE
        elif method in ("pinyin", "字母", "拼音"):
            s.SortMethod = _XL_PINYIN
        for k in keys:
            so = _SORT_ON.get(k["sortOn"], 0)
            sf = s.SortFields.Add(_col_rng(k["col"]), so, k["order"])
            if so in (1, 2):
                rgb = _parse_sort_color(k["color"])
                if rgb is None:
                    kind = "字体颜色" if so == 2 else "单元格颜色"
                    raise RuntimeError("按%s排序必须指定 color 参数（如 #FF0000）" % kind)
                # ⚠️ 不能整体给 SortOnValue 赋值；取其 Color 成员设置（本机实测）
                sf.SortOnValue.Color = rgb
            elif so == 3:
                sf.SetIcon(_resolve_sort_icon(wb, k["iconSet"], k["icon"]))
        s.Header = header
        # 【必须 SetRange】否则 Apply 只按第一个 Key 的 CurrentRegion 取数据区
        s.SetRange(rng)
        s.Apply()

    def _apply_legacy():
        """纯"按值"场景的 legacy 回退路径（本环境长期验证过）。"""
        rng.Sort(Key1=_col_rng(keys[0]["col"]), Order1=keys[0]["order"], Header=header)
        if len(keys) > 1:
            rng.Sort(Key1=_col_rng(keys[0]["col"]), Order1=keys[0]["order"],
                     Key2=_col_rng(keys[1]["col"]), Order2=keys[1]["order"], Header=header)
        if len(keys) > 2:
            rng.Sort(Key1=_col_rng(keys[0]["col"]), Order1=keys[0]["order"],
                     Key2=_col_rng(keys[1]["col"]), Order2=keys[1]["order"],
                     Key3=_col_rng(keys[2]["col"]), Order3=keys[2]["order"], Header=header)

    try:
        _apply_via_sort_object()
    except Exception:
        # 现代 Sort 对象失败时，仅"纯按值 + 无选项"请求回退 legacy（其它场景直接报错，避免静默降级）
        plain = all(k["sortOn"] in ("values", "") for k in keys) and \
            not p.get("matchCase") and not p.get("orientation") and not p.get("sortMethod")
        if not plain:
            raise
        _apply_legacy()

    # ── 排序后回传预览（表头 + 前 N 行）──
    preview_rows = int(p.get("previewRows") or 5)
    preview_rows = max(0, min(preview_rows, 50))
    grid = _range_grid(rng)
    if has_header and grid:
        headers = ["" if c is None else str(c) for c in grid[0]]
        data_rows = grid[1:]
    else:
        headers = ["列%d" % (i + 1) for i in range(len(grid[0]) if grid else 0)]
        data_rows = grid
    take = min(preview_rows, len(data_rows)) if preview_rows else 0
    preview = data_rows[:take]

    def _key_desc(k, first):
        on = {"cellcolor": "单元格颜色", "fontcolor": "字体颜色", "icon": "条件格式图标"}.get(k["sortOn"], "")
        # 有表头时用表头名（更可读），否则退回列字母
        if header_row and 1 <= k["col"] <= len(header_row) and header_row[k["col"] - 1] not in (None, ""):
            col_txt = str(header_row[k["col"] - 1])
        else:
            col_txt = _col_letter(k["col"])
        extra = ""
        if k["sortOn"] in ("cellcolor", "fontcolor") and _parse_sort_color(k["color"]) is not None:
            extra = " #%06X" % _parse_sort_color(k["color"])
        return ("" if first else "次要关键字 ") + "%s%s %s" % (
            col_txt, ("（%s%s）" % (on, extra)) if on else "", "降序" if k["desc"] else "升序")

    sorted_by = [_key_desc(k, i == 0) for i, k in enumerate(keys)]
    return {
        "success": True,
        "message": "已按 %s 排序（范围 %s）" % ("、".join(sorted_by), str(rng.Address)),
        "sorted": str(rng.Address),
        "range": str(rng.Address),
        "worksheet": str(ws.Name),
        "rowCount": len(data_rows),
        "sortedBy": sorted_by,
        "headers": headers,
        "previewRows": preview,
        "truncated": take < len(data_rows),
    }


def m_excel_apply_filter(args):
    """excel_apply_filter：对范围开启自动筛选并按 criteria（{字段:条件}）过滤。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    range_addr = str(p.get("range") or "").strip()
    if not range_addr:
        raise RuntimeError("缺少 range 参数")
    rng = ws.Range(range_addr)
    header_row = _range_grid(rng)[0] if _as_bool(p.get("hasHeader"), False) else None
    criteria = _as_obj(p.get("criteria"))
    if criteria:
        for key, cond in criteria.items():
            col = _col_spec_to_index(header_row, key)
            if col is None:
                continue
            try:
                rng.AutoFilter(Field=col, Criteria1=str(cond))
            except Exception:
                pass
    else:
        try:
            rng.AutoFilter()
        except Exception:
            pass
    return {"success": True, "filtered": str(rng.Address), "worksheet": str(ws.Name)}


def m_excel_clear_filter(args):
    """excel_clear_filter：清除工作表自动筛选（不删数据）。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    if ws.AutoFilterMode:
        ws.AutoFilterMode = False
    return {"success": True, "cleared": str(ws.Name)}


def m_excel_copy_range(args):
    """excel_copy_range：复制源范围到目标工作表（含格式）。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    src_ws = _target_sheet(app, wb, str(p.get("sourceSheet") or "").strip())
    src_addr = str(p.get("sourceRange") or "").strip() or str(src_ws.UsedRange.Address)
    src_rng = src_ws.Range(src_addr)
    dst_ws = _target_sheet(app, wb, str(p.get("destinationSheet") or "").strip())
    dst_addr = str(p.get("destinationRange") or "A1").strip()
    dst_rng = dst_ws.Range(dst_addr)
    src_rng.Copy(Destination=dst_rng)
    return {"success": True, "from": str(src_rng.Address), "to": str(dst_rng.Address),
            "destinationSheet": str(dst_ws.Name)}


def m_excel_merge_cells(args):
    """excel_merge_cells：合并单元格；mergeType=mergeAndCenter 时居中。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    rng_addr = str(p.get("range") or "").strip()
    if not rng_addr:
        raise RuntimeError("缺少 range 参数")
    rng = ws.Range(rng_addr)
    rng.Merge(False)
    if str(p.get("mergeType") or "merge").strip().lower() == "mergeandcenter":
        rng.HorizontalAlignment = XL_CENTER
        rng.VerticalAlignment = XL_CENTER
    return {"success": True, "merged": str(rng.Address), "worksheet": str(ws.Name)}


def m_excel_unmerge_cells(args):
    """excel_unmerge_cells：取消合并；fillAfterUnmerge=true 时把左上角值铺满。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    rng_addr = str(p.get("range") or "").strip()
    if not rng_addr:
        raise RuntimeError("缺少 range 参数")
    rng = ws.Range(rng_addr)
    fill = _as_bool(p.get("fillAfterUnmerge"), False)
    val = None
    try:
        if fill and rng.MergeCells:
            val = rng.MergeArea.Cells(1, 1).Value2
    except Exception:
        val = None
    rng.UnMerge()
    if fill and val is not None:
        rng.Value = val
    return {"success": True, "unmerged": str(rng.Address), "worksheet": str(ws.Name)}


def _last_used_col_letter(ws):
    """工作表 UsedRange 最后一列的列字母（如 'I'）。失败回退 'A'。"""
    try:
        uc = ws.UsedRange
        addr = str(ws.Cells(1, int(uc.Column) + int(uc.Columns.Count) - 1).Address)
        return addr.split('$')[1] if '$' in addr else 'A'
    except Exception:
        return 'A'


def _normalize_chart_data_range(ws, ref):
    """把模型常见的畸形区间写法修成合法 A1 区间。

    Excel COM 的 Range() 对 'A1:21'（列行:行，行区间简写）直接抛 0x800A03EC
    （实测 2026-10-27：弱模型把"前 21 行"写成 A1:21 → 建图失败 → 模型幻觉
    内嵌 base64 图 → 空图 + 假总结）。修复：
      'A1:21'  → 'A1:<UsedRange末列>21'（模型意图"行 1..21 全数据列"）
      '1:21'   → 'A1:<UsedRange末列>21'
    其余写法原样返回（带表名的由上层解析，非法地址让 COM 报真实错误）。
    """
    s = str(ref or '').strip()
    if '!' in s:
        return s
    m = re.fullmatch(r'\$?([A-Za-z]{1,3})\$?(\d+)\s*:\s*\$?(\d+)', s)
    if m:
        col = m.group(1).upper()
        r1, r2 = sorted((int(m.group(2)), int(m.group(3))))
        return '%s%d:%s%d' % (col, r1, _last_used_col_letter(ws), r2)
    m = re.fullmatch(r'\$?(\d+)\s*:\s*\$?(\d+)', s)
    if m:
        r1, r2 = sorted((int(m.group(1)), int(m.group(2))))
        return 'A%d:%s%d' % (r1, _last_used_col_letter(ws), r2)
    return s


def m_excel_create_chart(args):
    """excel_create_chart：在指定工作表创建图表（核心参数 + 容错高级参数）。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    data_range = str(p.get("dataRange") or "").strip()
    if not data_range:
        raise RuntimeError("缺少 dataRange 参数")
    chart_type = str(p.get("chartType") or "auto").strip().lower()
    CHART_TYPES = {"xlcolumnclustered": 51, "xlcolumnstacked": 52, "xlcolumnstacked100": 53,
                   "xlbarclustered": 57, "xlpie": 5, "xldoughnut": -4120, "xlline": 4,
                   "xllinemarkers": 65, "xlarea": 1, "xlxyScatter": -4169, "xlradar": -4151,
                   "xlbubble": 15}
    xtype = 51 if chart_type == "auto" or chart_type not in CHART_TYPES else CHART_TYPES[chart_type]
    sr = data_range
    src_ws = None
    if "!" in sr:
        sp, rp = sr.split("!", 1)
        sp = sp.strip().strip("'")
        try:
            src_ws = wb.Sheets(sp)
        except Exception:
            src_ws = None
        sr = rp
    if src_ws is None:
        src_ws = _target_sheet(app, wb, str(p.get("sheet") or "").strip())
    sr = _normalize_chart_data_range(src_ws, sr)
    data_rng = src_ws.Range(sr)
    pos = _as_obj(p.get("position"))
    left = int(pos.get("left", 0)) if pos else 0
    top = int(pos.get("top", 0)) if pos else 0
    width = int(pos.get("width", 480)) if pos else 480
    height = int(pos.get("height", 300)) if pos else 300
    # 【单图请求自清理】默认情况下（非看板/多图），建图前先删除历史 AI 图表，
    # 保证单图请求最终工作簿中至多只剩一张 AI 生成的图，杜绝重复出图。
    # 仅当显式 multiChart=true（看板/仪表盘等合法多图场景）时才跳过清理。
    # 这一层清理放在后端（而非前端注入参数），可覆盖所有调用路径
    # （agent-loop / skill 管道 / 兜底合成），无需依赖前端是否正确传参。
    _multiChart = _as_bool(p.get("multiChart"), False)
    if not _multiChart:
        # 1) 清理目标工作表上全部图表（替换旧图，含历史测试残留）
        try:
            for co in list(src_ws.ChartObjects()):
                co.Delete()
        except Exception:
            pass
        # 2) 全工作簿清理此前由 AI 生成的图表（命名以 AIChart_ 开头），
        #    覆盖"旧图落在不同工作表"导致的跨表重复。
        try:
            for _i in range(1, int(wb.Sheets.Count) + 1):
                _ws = wb.Sheets(_i)
                for co in list(_ws.ChartObjects()):
                    try:
                        if str(co.Name).startswith("AIChart_"):
                            co.Delete()
                    except Exception:
                        pass
        except Exception:
            pass
    try:
        shape = src_ws.Shapes.AddChart2(Style=201, XlChartType=xtype, Left=left,
                                        Top=top, Width=width, Height=height)
    except Exception:
        shape = src_ws.Shapes.AddChart(Type=xtype, Left=left, Top=top, Width=width, Height=height)
    try:
        shape.Name = "AIChart_" + str(int(time.time() * 1000))
    except Exception:
        pass
    chart = shape.Chart
    # 【pywin32 大坑·空图根因】SetSourceData(Source=Range对象) 在 pywin32 下实测
    # （2026-10-27，Range对象/带PlotBy关键字/ChartWizard 三种传法对照实验）全部产出
    # 1 个全 0 系列，系列名是整列/整行文本拼接——Range 被 marshal 成二维值数组后，
    # Excel 无法识别表头与行列方向（对应截图：图例是一条"2024/1/15…孙七"拼接文本
    # 的空图）。改为手动构建系列：读出二维值 → 数值列各建一个系列、首列作类别。
    try:
        _vals = to_jsonable(data_rng.Value)
    except Exception:
        _vals = None
    if _vals is None:
        grid = [[]]
    elif not isinstance(_vals, list):
        grid = [[_vals]]
    elif _vals and not isinstance(_vals[0], list):
        grid = [_vals]
    else:
        grid = _vals

    def _cat_str(v):
        """类别/系列名转显示文本（datetime 转日期串，None 转空串）。"""
        import datetime as _dt
        if isinstance(v, _dt.datetime):
            return v.strftime('%Y-%m-%d') if (v.hour, v.minute, v.second) == (0, 0, 0) \
                else v.strftime('%Y-%m-%d %H:%M:%S')
        if isinstance(v, _dt.date):
            return v.strftime('%Y-%m-%d')
        return '' if v is None else str(v)

    def _num(v):
        # 【None→VT_EMPTY】dataRange 声明范围常大于实际数据行（如聚合结果表只有 3 行
        # 却声明 $A$1:$I$21），越界处是 None。pywin32 把数组里的 None 转 VT_NULL，
        # Excel 的 Series.Values setter 不认 → DISP_E_TYPEMISMATCH（实测 21:47
        # "构建图表数据系列失败 (-2147352571)"）。pythoncom.Empty（VT_EMPTY）
        # 才会被 Excel 当作"空数据点"。
        import pythoncom as _pc
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return v
        return _pc.Empty

    if not grid or not grid[0]:
        try:
            shape.Delete()
        except Exception:
            pass
        raise RuntimeError("图表源数据 %s 为空，无法生成图表" % sr)

    series_by = str(p.get("seriesBy") or "").strip().lower()
    by_rows = (series_by == "rows")
    if _as_bool(p.get("switchRowColumn"), False):
        by_rows = not by_rows

    # 清掉 AddChart2 自带的空系列
    try:
        while int(chart.SeriesCollection().Count) > 0:
            chart.SeriesCollection(1).Delete()
    except Exception:
        pass

    try:
        if by_rows:
            # 每行一系列：第 0 列作系列名，首行作类别
            cats = [_cat_str(x) for x in grid[0][1:]]
            for row in grid[1:]:
                if all(v is None or v == '' for v in row[1:]):
                    continue
                ser = chart.SeriesCollection().NewSeries()
                ser.Name = _cat_str(row[0])
                ser.Values = [_num(v) for v in row[1:]]
                ser.XValues = cats
        else:
            # 每列一系列（默认）：首行表头、第 0 列作类别；只给含数值的列建系列
            header = grid[0]
            body = grid[1:]
            cats = [_cat_str(r[0]) for r in body]
            for c in range(1, len(header)):
                col_vals = [(r[c] if c < len(r) else None) for r in body]
                if not any(isinstance(v, (int, float)) and not isinstance(v, bool) for v in col_vals):
                    continue  # 文本列（如 客户类型）不建系列
                ser = chart.SeriesCollection().NewSeries()
                ser.Name = _cat_str(header[c])
                ser.Values = [_num(v) for v in col_vals]
                ser.XValues = cats
    except Exception as e:
        try:
            shape.Delete()
        except Exception:
            pass
        raise RuntimeError("构建图表数据系列失败（源 %s）: %s" % (sr, str(e)[:120]))
    if p.get("title"):
        chart.HasTitle = True
        chart.ChartTitle.Text = str(p.get("title"))
    else:
        chart.HasTitle = False
    if p.get("hasLegend") is not None:
        chart.HasLegend = _as_bool(p.get("hasLegend"), True)
    lp = {"xlbottom": -4107, "xlright": -4152, "xltop": -4160, "xlleft": -4161, "xlcorner": -4114}
    lv = lp.get(str(p.get("legendPosition") or "").strip().lower())
    if lv is not None:
        try:
            chart.Legend.Position = lv
        except Exception:
            pass
    try:
        if p.get("xAxisTitle"):
            ax = chart.Axes(1)
            ax.HasTitle = True
            ax.AxisTitle.Text = str(p.get("xAxisTitle"))
        if p.get("yAxisTitle"):
            ay = chart.Axes(2)
            ay.HasTitle = True
            ay.AxisTitle.Text = str(p.get("yAxisTitle"))
    except Exception:
        pass
    if _as_bool(p.get("showDataLabels"), False):
        try:
            for s in chart.SeriesCollection():
                s.DataLabels.ShowValue = True
        except Exception:
            pass
    if p.get("chartStyle"):
        try:
            chart.ChartStyle = int(p.get("chartStyle"))
        except Exception:
            pass
    if _as_bool(p.get("useSecondaryAxisLastSeries"), False):
        try:
            ss = chart.SeriesCollection()
            ss(ss.Count).AxisGroup = 2
        except Exception:
            pass
    # 【空图熔断】系列是手动构建的（只建有数值的列），理论必 ≥1；此校验兜底
    # "区间内连一个数值都没有"（如 H1:I21 纯文本）的情形：删图报错，让 agent-loop
    # 走失败重试，杜绝"假成功总结 + 空图"。
    def _series_count():
        try:
            return int(chart.SeriesCollection().Count)
        except Exception:
            return 0
    def _plottable_series():
        """至少含 1 个数值点的系列数（None/文本/错误值不算）。"""
        n = 0
        try:
            for s in chart.SeriesCollection():
                try:
                    for v in (s.Values or []):
                        if isinstance(v, (int, float)) and not isinstance(v, bool):
                            n += 1
                            break
                except Exception:
                    pass
        except Exception:
            pass
        return n
    if _plottable_series() == 0:
        try:
            shape.Delete()
        except Exception:
            pass
        raise RuntimeError(
            "图表源数据 %s 没有可绘制的数据系列（区间内可能没有数值列，或区域为空），"
            "请改选包含数值列的 dataRange" % sr)
    return {"success": True, "chartType": chart_type, "sheet": str(src_ws.Name),
            "chartName": str(shape.Name), "seriesCount": _series_count(),
            "dataRange": sr, "position": {"left": left, "top": top,
                                          "width": width, "height": height}}


def m_excel_create_pivot_table(args):
    """excel_create_pivot_table：创建原生数据透视表。"""
    p = _tool_params(args)
    app = get_excel(create=False)
    wb = app.ActiveWorkbook
    if wb is None:
        raise RuntimeError("没有打开的工作簿")
    src_range = str(p.get("sourceRange") or "").strip()
    if not src_range:
        raise RuntimeError("缺少 sourceRange 参数")
    sr = src_range
    src_ws = None
    if "!" in sr:
        sp, rp = sr.split("!", 1)
        sp = sp.strip().strip("'")
        try:
            src_ws = wb.Sheets(sp)
        except Exception:
            src_ws = None
        sr = rp
    if src_ws is None:
        src_ws = wb.ActiveSheet
    sr = _normalize_chart_data_range(src_ws, sr)
    source_rng = src_ws.Range(sr)
    # SourceData 必须传字符串地址（而非 COM Range 对象），否则 CreatePivotTable 后
    # PivotFields 会报 "方法无效"
    if re.fullmatch(r'[\w\u4e00-\u9fff]+', src_ws.Name):
        src_ref = src_ws.Name + "!" + sr
    else:
        src_ref = "'" + src_ws.Name + "'!" + sr
    target_sheet = str(p.get("targetSheet") or "数据透视表").strip()
    try:
        dst_ws = wb.Sheets(target_sheet)
    except Exception:
        dst_ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        dst_ws.Name = target_sheet
    cell = str(p.get("targetCell") or "A1").strip()
    dest_rng = dst_ws.Range(cell)
    pc = wb.PivotCaches().Create(SourceType=1, SourceData=src_ref)
    pt_name = "PivotTable" + str(dst_ws.PivotTables().Count + 1)
    pc.CreatePivotTable(TableDestination=dest_rng, TableName=pt_name)
    # 从目标单元格取引用，避免 CreatePivotTable 后 PivotFields 尚未物化的 COM 错误
    pt = dest_rng.PivotTable
    pt.ManualUpdate = True
    header_row = _range_grid(source_rng)[0] if _range_grid(source_rng) else []
    row_fields = _as_list(p.get("rowFields"))
    col_fields = _as_list(p.get("columnFields"))
    val_fields = _as_list(p.get("valueFields"))
    val_funcs = _as_list(p.get("valueFunctions"))
    FUNC = {"sum": -4157, "avg": -4106, "average": -4106, "count": -4112, "max": -4136, "min": -4139}

    def _fidx(spec):
        idx = _col_spec_to_index(header_row, spec)
        if idx is None:
            raise RuntimeError("找不到透视表字段: " + str(spec))
        return idx

    for i, f in enumerate(row_fields, 1):
        pf = pt.PivotFields(_fidx(f))
        pf.Orientation = 1
        pf.Position = i
    for i, f in enumerate(col_fields, 1):
        pf = pt.PivotFields(_fidx(f))
        pf.Orientation = 2
        pf.Position = i
    for i, vf in enumerate(val_fields):
        fn = str(val_funcs[i] if i < len(val_funcs) else "Sum").lower()
        # pywin32 下 AddDataField 不能带 Caption/Function 关键字或位置参数，
        # 否则报 -2146827284；先用「仅字段」创建，再设属性（Excel 自动按函数命名）。
        cp = pt.AddDataField(pt.PivotFields(_fidx(vf)))
        cp.Function = FUNC.get(fn, -4157)
        try:
            cp.NumberFormat = "#,##0.00"
        except Exception:
            pass
    pt.ManualUpdate = False
    pt.RefreshTable()
    return {"success": True, "pivotTableName": pt_name, "sheet": str(dst_ws.Name),
            "cell": cell, "rowFields": row_fields, "valueFields": val_fields}


# MCP 本地工具注册（函数定义在 ROUTES 字面量之后，故用 update 后置注册）
ROUTES.update({
    "GetAvailableMcpTools": m_get_available_mcp_tools,
    "excel_read_range": m_excel_read_range,
    "excel_get_cell_value": m_excel_get_cell_value,
    "excel_list_sheets": m_excel_list_sheets,
    "excel_get_sheets": m_excel_list_sheets,
    "excel_write_range": m_excel_write_range,
    "excel_clear_range": m_excel_clear_range,
    "excel_create_sheet": m_excel_create_sheet,
    "excel_format_range": m_excel_format_range,
    "excel_autofit": m_excel_autofit,
    "excel_save_workbook": m_excel_save_workbook,
    "excel_aggregate": m_excel_aggregate,
    "excel_sort_range": m_excel_sort_range,
    "excel_apply_filter": m_excel_apply_filter,
    "excel_clear_filter": m_excel_clear_filter,
    "excel_copy_range": m_excel_copy_range,
    "excel_merge_cells": m_excel_merge_cells,
    "excel_unmerge_cells": m_excel_unmerge_cells,
    "excel_create_chart": m_excel_create_chart,
    "excel_create_pivot_table": m_excel_create_pivot_table,
})


def route_method(method, args):
    if method in ROUTES:
        return ROUTES[method](args or [])
    if isinstance(method, str) and re.match(r'^(excel|ppt|word|pdf|computer|web|time)_', method):
        raise RuntimeError("本地桥(pywin32)暂未实现工具: " + method +
                           "（已实现: excel_read_range / excel_get_cell_value / excel_list_sheets / "
                           "excel_write_range / excel_clear_range / excel_create_sheet / "
                           "excel_format_range / excel_autofit / excel_save_workbook / "
                           "excel_aggregate / excel_sort_range / excel_apply_filter / "
                           "excel_clear_filter / excel_copy_range / excel_merge_cells / "
                           "excel_unmerge_cells / excel_create_chart / excel_create_pivot_table）")
    raise RuntimeError("方法在本地桥(pywin32)模式下未实现: " + str(method))


# ===================== 单线程 COM 串行化 =====================
# Excel 是 STA（单线程单元）COM 对象，跨线程调用会静默失败或抛 -2146827284 "发生意外"。
# 本地桥用 ThreadingHTTPServer，多个请求线程可能并发触碰同一个 Excel 实例；
# 因此把所有 Excel COM 操作集中到唯一一个 STA 后台线程串行执行，保证线程亲和。
import threading as _T
import queue as _Q
_COM_QUEUE = _Q.Queue()
_COM_THREAD = None


def _com_worker():
    pythoncom.CoInitialize()  # STA 单元
    while True:
        item = _COM_QUEUE.get()
        if item is None:
            _COM_QUEUE.task_done()
            break
        fn, args, kwargs, holder = item
        try:
            holder["value"] = fn(*args, **kwargs)
        except Exception as e:
            holder["error"] = e
        finally:
            holder["event"].set()
            _COM_QUEUE.task_done()


def run_in_com(fn, *args, **kwargs):
    holder = {"event": _T.Event()}
    _COM_QUEUE.put((fn, args, kwargs, holder))
    holder["event"].wait()
    if "error" in holder:
        raise holder["error"]
    return holder.get("value")


def _start_com_thread():
    global _COM_THREAD
    if _COM_THREAD is None or not _COM_THREAD.is_alive():
        _COM_THREAD = _T.Thread(target=_com_worker, daemon=True)
        _COM_THREAD.start()


def _status_excel():
    app = get_excel(create=False)
    return (True, "Excel " + str(app.Version))


# ===================== HTTP 服务 =====================

class Handler(BaseHTTPRequestHandler):
    def _cors(self):
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _send_json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self._cors()
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _status_payload(self):
        connected = False
        version = ""
        reason = None
        if PYWIN32:
            try:
                app = get_excel(create=False)
                connected = True
                version = "Excel " + str(app.Version)
            except Exception as e:
                connected = False
                reason = str(e)
        else:
            reason = "pywin32 不可用: " + PYWIN32_ERR
        return {
            "status": "ok",
            "service": "SmartPilot Local Bridge (pywin32)",
            "version": "1.1.0",
            "excelConnected": connected,
            "excelRunning": _excel_process_running(),
            "reason": reason,
            "winaxAvailable": False,
            "pywin32Available": PYWIN32,
            "driver": "pywin32",
            "mockMode": False,
            "mockReason": (None if PYWIN32 else "pywin32 不可用: " + PYWIN32_ERR),
            "timestamp": _now(),
        }

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        # 容错：前端有些地方把 baseUrl（已含 /api/excel）再拼上 /api/mcp/...，
        # 造出 /api/excel/api/mcp/tools 这类重复前缀路径 → 归并回标准路径，避免 404。
        path = parsed.path
        if path.startswith("/api/excel/api/"):
            path = path[len("/api/excel"):]
        if path == "/api/excel/status":
            self._send_json(self._status_payload())
            return
        if path == "/health":
            self._send_json({"ok": True, "service": "ocr-host-compat"})
            return
        if path == "/api/mcp/status":
            self._send_json(m_get_mcp_status([]))
            return
        if path == "/api/mcp/tools":
            tools = load_mcp_tools()
            self._send_json({"success": True, "tools": tools, "count": len(tools)})
            return
        if path.startswith("/api/skills/shared/"):
            nm = unquote(path[len("/api/skills/shared/"):])
            self._send_json(m_get_skill_by_name([nm]))
            return
        if path.startswith("/api/"):
            self._send_json({"success": False, "error": "未知 GET 接口: " + parsed.path}, 404)
            return
        self._serve_static(parsed.path)

    def do_POST(self):
        parsed = urlparse(self.path)
        try:
            length = int(self.headers.get("Content-Length", 0) or 0)
            raw = self.rfile.read(length) if length else b"{}"
            body = json.loads(raw.decode("utf-8") or "{}")
        except Exception as e:
            self._send_json({"success": False, "error": "请求体解析失败: " + str(e)}, 400)
            return

        # OcrHost 兼容协议：ocr-panel.html 独立页从 9527 加载时同源调用。
        # 与 /api/excel/invoke 共用同一张 ROUTES 表（单一事实来源），
        # 仅响应包格式不同（{ok, result:字符串}），由协议层统一 json.dumps。
        if parsed.path == "/v1/call":
            method = str(body.get("method") or "")
            for _pfx in ("OcrBridge.", "OcrServiceManager."):
                if method.startswith(_pfx):
                    method = method[len(_pfx):]
                    break
            handler = ROUTES.get(method)
            if not handler:
                self._send_json({"ok": False, "error": "方法不存在: " + method})
                return
            try:
                result = json.dumps(handler(body.get("args") or []), ensure_ascii=False)
                self._send_json({"ok": True, "result": result})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)})
            return

        if parsed.path == "/api/excel/connect":
            with LOCK:
                res = connect_excel()
            if res.get("success"):
                self._send_json({
                    "success": True,
                    "excelVersion": res.get("version", "Excel"),
                    "serviceVersion": "1.1.0",
                    "winaxAvailable": False,
                    "pywin32Available": PYWIN32,
                    "mockMode": False,
                    "mockReason": None,
                    "message": "Excel 连接成功 (pywin32)",
                })
            else:
                self._send_json({"success": False, "error": res.get("error", "连接失败"),
                                 "message": "无法连接到 Excel"}, 500)
            return

        if parsed.path == "/api/excel/invoke" or parsed.path == "/invoke":
            method = body.get("method")
            args = body.get("arguments", []) or []
            try:
                with LOCK:
                    # 非环境探测类方法自动连接 Excel；
                    # License*/AI 配置/Skills/MCP 状态/OCR 只读写本机文件或查本地包，不碰 Excel，
                    # 豁免避免点"激活"/"保存设置"/打开 MCP、OCR 面板时弹出 Excel 窗口或超时。
                    _local_only = isinstance(method, str) and (
                        method.startswith("License") or
                        method in ("SaveAIConfig", "GetAIConfig", "SaveLLMConfig", "GetLLMConfig",
                                   "GetAllSkills", "GetAvailableSkills", "GetMcpStatus",
                                   "StartMcpServer", "StopMcpServer", "GetUndoStack",
                                   "ToggleSkillEnabled", "SaveSkill", "DeleteSkill") or
                        method.startswith("GetOcr") or method.startswith("StartOcr") or
                        method.startswith("GetParseScript") or method.startswith("SaveParseScript") or
                        method.startswith("DeleteParseScript") or method.startswith("DuplicateParseScript") or
                        method.startswith("ResetBuiltinScript") or method.startswith("ClearParseScriptCache") or
                        method.startswith("GetPaddle") or method.startswith("SavePaddle") or
                        method.startswith("SetPaddle") or method == "ImportParseScriptFromMarkdown")
                    if method != "DetectEnvironment" and not _local_only:
                        try:
                            get_excel(create=True)
                        except Exception:
                            pass
                    result = route_method(method, args)
                resp = {"success": True, "data": result, "method": method}
                if isinstance(result, dict) and "taskCompleted" in result:
                    resp["taskCompleted"] = result["taskCompleted"]
                self._send_json(resp)
            except Exception as e:
                self._send_json({"success": False, "error": str(e), "method": method})
            return

        self._send_json({"success": False, "error": "未知 POST 接口: " + parsed.path}, 404)

    def _serve_static(self, path):
        rel = path.lstrip("/")
        if rel == "" or rel == "index.html":
            rel = "index.html"
        fpath = os.path.normpath(os.path.join(UI_DIR, rel))
        if not fpath.startswith(UI_DIR) or not os.path.isfile(fpath):
            self.send_response(404)
            self._cors()
            self.end_headers()
            return
        self.send_response(200)
        self._cors()
        # 禁用浏览器缓存：UI 资源（含 excel-connector.js 等）编辑后必须立即生效，
        # 否则旧 ?v= 戳会让浏览器一直用缓存里的旧 JS，垫片/数据插入代码无法加载。
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        if fpath.lower().endswith(".html"):
            self.send_header("Content-Type", "text/html; charset=utf-8")
        elif fpath.lower().endswith(".js"):
            self.send_header("Content-Type", "application/javascript; charset=utf-8")
        elif fpath.lower().endswith(".css"):
            self.send_header("Content-Type", "text/css; charset=utf-8")
        else:
            self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Length", str(os.path.getsize(fpath)))
        self.end_headers()
        with open(fpath, "rb") as f:
            self.wfile.write(f.read())

    def log_message(self, fmt, *args):
        if QUIET:
            return
        sys.stderr.write("[Bridge] " + (fmt % args) + "\n")


def _now():
    from datetime import datetime
    return datetime.now().isoformat()


def connect_excel():
    try:
        app = get_excel(create=True)
        return {"success": True, "version": "Excel " + str(app.Version), "mock": False}
    except Exception as e:
        return {"success": False, "error": str(e)}


def main():
    if not PYWIN32:
        sys.stderr.write("[Bridge] 警告: pywin32 不可用 (%s)，Excel 操作将失败。请安装: pip install pywin32\n" % PYWIN32_ERR)
    sys.stderr.write("============================================================\n")
    sys.stderr.write("🚀 SmartPilot 本地代理服务 (pywin32 驱动真实 Excel COM)\n")
    sys.stderr.write("📍 服务地址: http://localhost:%d\n" % PORT)
    sys.stderr.write("📄 Web界面:  http://localhost:%d/index.html\n" % PORT)
    sys.stderr.write("🔗 API端点:  http://localhost:%d/api/excel/*\n" % PORT)
    sys.stderr.write("============================================================\n")
    # 必须用 ThreadingHTTPServer：浏览器会并行拉取大量 defer 脚本（含 /js/*.js）并复用
    # keep-alive 连接；单线程 HTTPServer 一次只能处理一个连接，靠后的脚本请求会被饿死/超时，
    # 导致 side-toolbar.js / task-planner.js 等加载失败、内联 onclick 报“函数未定义”。
    # （Excel COM 操作已通过 _COM_QUEUE 串行化到唯一 STA 线程，多线程 HTTP 不会破坏线程亲和性。）
    server = ThreadingHTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.shutdown()


if __name__ == "__main__":
    main()
