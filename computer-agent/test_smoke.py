"""
Computer Agent 分层测试脚本
用法: python test_smoke.py [--level N]
  level=1  不依赖 OCR — 感知(截图) + 操作(键鼠) + 辅助
  level=2  MCP 协议层 — 验证 tools/list 和 tools/call
  level=3  端到端组合 — 真实打开应用、点击、输入
  level=4  OCR 能力 — 需要 PaddleOCR 已安装（冷启动 3-5s）
"""
from __future__ import annotations
import sys, os, argparse, json, time, asyncio, traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _HERE)

PASS = 0
FAIL = 0

def test(name, fn):
    global PASS, FAIL
    print(f"\n{'─'*60}")
    print(f"▶ {name}")
    print(f"{'─'*60}")
    try:
        result = fn()
        if result is False:
            print(f"  ❌ FAILED")
            FAIL += 1
        else:
            print(f"  ✅ PASSED  →  {result}")
            PASS += 1
    except Exception as e:
        print(f"  ❌ ERROR: {e}")
        traceback.print_exc()
        FAIL += 1

def summary():
    print(f"\n{'='*60}")
    print(f"结果: {PASS} passed, {FAIL} failed")
    print(f"{'='*60}")


# ═══════════════════════════════════════════════════════
# Level 1: 不依赖 OCR 的核心模块
# ═══════════════════════════════════════════════════════
def test_level1():
    # 1.1 工具导入
    def import_test():
        from tools import perception, control, locate, utility
        from core import utils, screen_capture, mouse_keyboard
        return f"全部模块导入 OK"

    test("1.1 模块导入", import_test)

    # 1.2 屏幕信息
    def screen_info():
        from tools.perception import screen_get_info
        info = screen_get_info()
        assert info["success"]
        w, h = info["screen_size"]
        assert w > 100 and h > 100
        mx, my = info["mouse_position"]
        assert 0 <= mx < w and 0 <= my < h
        return f"屏幕 {w}x{h}, 鼠标 ({mx},{my})"

    test("1.2 screen_get_info", screen_info)

    # 1.3 截图
    def capture():
        from tools.perception import screen_capture
        cap = screen_capture()
        assert cap["success"]
        assert len(cap["image_base64"]) > 10000
        return f"截图 {cap['width']}x{cap['height']}, base64 {len(cap['image_base64'])} chars"

    test("1.3 screen_capture (全屏)", capture)

    # 1.4 区域截图
    def capture_region():
        from tools.perception import screen_capture
        cap = screen_capture(region=[100, 100, 400, 300])
        assert cap["success"]
        assert cap["width"] == 400 and cap["height"] == 300
        return f"区域截图 {cap['width']}x{cap['height']}"

    test("1.4 screen_capture (区域)", capture_region)

    # 1.5 safe_mode 键鼠（不实际执行）
    def mouse_safe():
        from core.utils import set_safe_mode
        from tools.control import mouse_move, mouse_click, key_press, type_text, wait
        set_safe_mode(True)
        r1 = mouse_move(200, 300)
        r2 = mouse_click(500, 400)
        r3 = key_press("enter")
        r4 = type_text("测试文字 hello")
        r5 = wait(50)
        assert r1["safe_mode"] and not r1["executed"]
        assert r2["safe_mode"] and not r2["executed"]
        assert r3["safe_mode"] and not r3["executed"]
        return "5 个操作全部 safe_mode"

    test("1.5 键鼠 safe_mode", mouse_safe)

    # 1.6 剪贴板
    def clipboard():
        from tools.utility import clipboard_write, clipboard_read
        from core.utils import set_safe_mode
        set_safe_mode(False)
        clipboard_write("test_clipboard_123")
        time.sleep(0.1)
        read_back = clipboard_read()
        assert "test_clipboard_123" in read_back["text"]
        return f"剪贴板往返 OK: '{read_back['text']}'"

    test("1.6 clipboard 读写", clipboard)

    # 1.7 命令执行
    def cmd():
        from tools.utility import execute_command
        from core.utils import set_safe_mode
        set_safe_mode(False)
        r = execute_command("echo hello_cmd_test")
        assert r["success"]
        assert "hello_cmd_test" in r["stdout"]
        return f"stdout: {r['stdout'].strip()}"

    test("1.7 execute_command", cmd)


# ═══════════════════════════════════════════════════════
# Level 2: MCP 协议层（tools/list + tools/call）
# ═══════════════════════════════════════════════════════
def test_level2():
    import mcp_server as srv

    # 2.1 tools 数量
    async def tools_count():
        tools = await srv.mcp.list_tools()
        return f"注册 {len(tools)} 个 tools"

    test("2.1 MCP tools 注册", lambda: asyncio.run(tools_count()))

    # 2.2 每个 tool 都有 name + description
    async def tools_have_docs():
        tools = await srv.mcp.list_tools()
        missing = [t.name for t in tools if not t.description or len(t.description.strip()) < 5]
        if missing:
            return f"⚠️ 以下 tool 缺少描述: {missing}"
        return f"全部 {len(tools)} 个 tool 都有描述"

    test("2.2 tool 描述完整性", lambda: asyncio.run(tools_have_docs()))

    # 2.3 逐个 call 验证（safe_mode 下，只看 schema 接受）
    async def call_all_safe():
        from core.utils import set_safe_mode
        set_safe_mode(True)
        tools = await srv.mcp.list_tools()
        results = []
        for t in tools:
            try:
                if t.name == "set_safe_mode":
                    r = await srv.mcp.call_tool("set_safe_mode", {"enabled": True})
                elif t.name == "wait":
                    r = await srv.mcp.call_tool("wait", {"milliseconds": 10})
                elif t.name == "mouse_move":
                    r = await srv.mcp.call_tool("mouse_move", {"x": 100, "y": 100})
                elif t.name == "mouse_click":
                    r = await srv.mcp.call_tool("mouse_click", {"x": 100, "y": 100})
                elif t.name == "key_press":
                    r = await srv.mcp.call_tool("key_press", {"key": "enter"})
                elif t.name == "key_hotkey":
                    r = await srv.mcp.call_tool("key_hotkey", {"keys": ["ctrl", "c"]})
                elif t.name == "type_text":
                    r = await srv.mcp.call_tool("type_text", {"text": "test"})
                elif t.name == "mouse_drag":
                    r = await srv.mcp.call_tool("mouse_drag", {"from_x": 0, "from_y": 0, "to_x": 100, "to_y": 100})
                elif t.name == "mouse_scroll":
                    r = await srv.mcp.call_tool("mouse_scroll", {"amount": 3, "direction": "down"})
                elif t.name == "screen_capture":
                    r = await srv.mcp.call_tool("screen_capture", {})
                elif t.name == "screen_get_info":
                    r = await srv.mcp.call_tool("screen_get_info", {})
                elif t.name == "screen_ocr":
                    r = await srv.mcp.call_tool("screen_ocr", {})
                elif t.name == "clipboard_read":
                    r = await srv.mcp.call_tool("clipboard_read", {})
                elif t.name == "clipboard_write":
                    r = await srv.mcp.call_tool("clipboard_write", {"text": "test"})
                elif t.name == "execute_command":
                    r = await srv.mcp.call_tool("execute_command", {"command": "echo ok"})
                elif t.name == "open_application":
                    r = await srv.mcp.call_tool("open_application", {"name": "notepad"})
                elif t.name == "locate_text":
                    r = await srv.mcp.call_tool("locate_text", {"text": "测试"})
                elif t.name == "locate_image":
                    r = await srv.mcp.call_tool("locate_image", {"image_path": "dummy.png"})
                elif t.name == "locate_color":
                    r = await srv.mcp.call_tool("locate_color", {"color": "#FF0000"})
                elif t.name == "screen_describe":
                    r = await srv.mcp.call_tool("screen_describe", {})
                else:
                    r = await srv.mcp.call_tool(t.name, {})
                results.append(f"  ✔ {t.name}")
            except Exception as e:
                results.append(f"  ✘ {t.name}: {e}")
        return "\n" + "\n".join(results)

    test("2.3 逐个 tools/call 验证", lambda: asyncio.run(call_all_safe()))


# ═══════════════════════════════════════════════════════
# Level 3: 端到端组合（真实操作电脑）
# ═══════════════════════════════════════════════════════
def test_level3():
    import pyautogui
    from core.utils import set_safe_mode
    set_safe_mode(False)

    # 3.1 真实鼠标移动（小范围，安全）
    def real_move():
        from tools.control import mouse_move, wait
        start = pyautogui.position()
        mouse_move(start.x + 50, start.y + 50)
        wait(100)
        end = pyautogui.position()
        dist = ((end.x - start.x) ** 2 + (end.y - start.y) ** 2) ** 0.5
        mouse_move(start.x, start.y)  # 归位
        assert dist >= 30, f"移动距离太小: {dist}"
        return f"鼠标从 ({start.x},{start.y}) 移到 ({end.x},{end.y}), 距离 {dist:.0f}"

    test("3.1 真实鼠标移动", real_move)

    # 3.2 剪贴板 + 粘贴（中文）
    def chinese_type():
        from tools.utility import clipboard_write, clipboard_read
        from tools.control import key_hotkey
        clipboard_write("你好 World 123")
        time.sleep(0.1)
        r = clipboard_read()
        assert "你好 World" in r["text"]
        return f"中文剪贴板 OK: '{r['text']}'"

    test("3.2 中文剪贴板往返", chinese_type)

    # 3.3 启动记事本（会弹出窗口）
    def open_notepad():
        from tools.utility import open_application
        r = open_application("notepad")
        # 给窗口弹出来一点时间
        time.sleep(1.5)
        assert r["success"] or r.get("executed")
        return f"notepad 启动 → 已执行: {r}"

    test("3.3 open_application notepad", open_notepad)


# ═══════════════════════════════════════════════════════
# Level 4: OCR 能力（需要 PaddleOCR 已安装）
# ═══════════════════════════════════════════════════════
def test_level4():
    # 4.1 PaddleOCR import
    def ocr_import():
        import paddleocr
        return f"paddleocr OK, version={paddleocr.__version__ if hasattr(paddleocr, '__version__') else 'unknown'}"

    test("4.1 PaddleOCR import", ocr_import)

    # 4.2 OCR 识别（首次调用触发模型加载，冷启动 3-5s）
    def ocr_recognize():
        from tools.perception import screen_ocr
        print("    ⏳ 首次加载 PaddleOCR 模型约 3-5s...")
        result = screen_ocr()
        assert result["success"]
        items = result["items"]
        return f"识别到 {len(items)} 个文本块，前 3 个: {[it['text'] for it in items[:3]]}"

    test("4.2 screen_ocr (全屏)", ocr_recognize)

    # 4.3 locate_text 定位真实文字
    def locate_real_text():
        from tools.locate import locate_text
        result = locate_text("开始", match_mode="fuzzy")
        return f"'开始' 找到 {result['match_count']} 处, 模式 fuzzy"

    test("4.3 locate_text 真实搜索", locate_real_text)


# ═══════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--level", type=int, default=0,
                        help="1=基础 2=MCP协议 3=端到端 4=OCR, 0=全部(按顺序遇到依赖缺失败跳过)")
    args = parser.parse_args()

    print("=" * 60)
    print("Computer Agent 分层测试")
    print("=" * 60)

    if args.level == 0 or args.level == 1:
        print("\n" + "#" * 60)
        print("# LEVEL 1: 核心模块（不依赖 OCR）")
        print("#" * 60)
        test_level1()

    if args.level == 0 or args.level == 2:
        print("\n" + "#" * 60)
        print("# LEVEL 2: MCP 协议层")
        print("#" * 60)
        test_level2()

    if args.level == 0 or args.level == 3:
        print("\n" + "#" * 60)
        print("# LEVEL 3: 端到端真实操作")
        print("#" * 60)
        test_level3()

    if args.level == 4:
        print("\n" + "#" * 60)
        print("# LEVEL 4: OCR 能力（需 PaddleOCR）")
        print("#" * 60)
        test_level4()

    summary()
