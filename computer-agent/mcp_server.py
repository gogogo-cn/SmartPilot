"""
Computer Agent — MCP Server 入口
使用官方 mcp SDK (v2.x) 的 MCPServer，自动从 Python 函数签名生成 Tool Schema。
传输方式: stdio（Trae 通过子进程启动）
"""
from __future__ import annotations

import json
import logging
import sys
import os
from typing import Any, Optional

# 让 tools/core 可以被导入
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

# ─── 日志配置 ────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger("computer-agent")

# ─── 导入工具函数 ────────────────────────────────────────────────────
import tools.perception as P
import tools.control as C
import tools.locate as L
import tools.utility as U
from core.utils import load_config, set_safe_mode as _set_safe_mode, AgentError

# ─── 构建 MCP Server ────────────────────────────────────────────────

# 先检测 mcp 是否可用
try:
    from mcp.server.mcpserver import MCPServer
except ImportError:
    print("[FATAL] mcp SDK 未安装。运行: pip install mcp", file=sys.stderr)
    sys.exit(1)

# 兼容变量名 mcp (历史命名)
mcp = MCPServer(
    name="computer-agent",
    description="让 LLM 通过 PaddleOCR + pyautogui 操作电脑。"
                "提供屏幕感知、键鼠控制、精准定位三大能力。",
)


# ════════════════════════════════════════════════════════════════════
# 感知层 Tools
# ════════════════════════════════════════════════════════════════════

@mcp.tool()
def screen_capture(region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    截图并返回 base64 PNG 图片。可用于视觉模型分析或直接观察。

    Args:
        region: 可选 [x, y, w, h] 只截区域；不传则全屏
    """
    return P.screen_capture(region)


@mcp.tool()
def screen_ocr(region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    截图并用 PaddleOCR 识别所有文字，返回每个文本块的坐标和置信度。
    这是 Agent "看"屏幕的核心能力 — 用它来读 UI 上的按钮、菜单、状态。

    Args:
        region: 可选 [x, y, w, h] 只识别屏幕的某个区域
    """
    return P.screen_ocr(region)


@mcp.tool()
def screen_get_info() -> dict[str, Any]:
    """获取屏幕分辨率、鼠标当前位置、活动窗口标题。做决策前先了解屏幕状态。"""
    return P.screen_get_info()


# ════════════════════════════════════════════════════════════════════
# 操作层 Tools
# ════════════════════════════════════════════════════════════════════

@mcp.tool()
def mouse_move(x: int, y: int) -> dict[str, Any]:
    """移动鼠标到绝对坐标 (x, y)。坐标原点在屏幕左上角。"""
    return C.mouse_move(x, y)


@mcp.tool()
def mouse_click(x: Optional[int] = None, y: Optional[int] = None,
                button: str = "left", clicks: int = 1) -> dict[str, Any]:
    """
    点击鼠标。可指定坐标或点击当前位置。

    Args:
        x: 可选，点击 X 坐标
        y: 可选，点击 Y 坐标
        button: 'left' | 'right' | 'middle'，默认 left
        clicks: 点击次数，默认 1
    """
    return C.mouse_click(x, y, button, clicks)


@mcp.tool()
def mouse_drag(from_x: int, from_y: int, to_x: int, to_y: int,
               duration: float = 0.5) -> dict[str, Any]:
    """
    拖拽鼠标从一点到另一点。

    Args:
        from_x: 起点 X
        from_y: 起点 Y
        to_x: 终点 X
        to_y: 终点 Y
        duration: 拖拽耗时秒数，默认 0.5
    """
    return C.mouse_drag(from_x, from_y, to_x, to_y, duration)


@mcp.tool()
def mouse_scroll(amount: int, direction: str = "down") -> dict[str, Any]:
    """
    滚动鼠标滚轮。

    Args:
        amount: 滚动量，正数
        direction: 'up' | 'down' | 'left' | 'right'
    """
    return C.mouse_scroll(amount, direction)


@mcp.tool()
def key_press(key: str) -> dict[str, Any]:
    """按下并释放一个键。如 'enter' 'escape' 'f5' 'space'。"""
    return C.key_press(key)


@mcp.tool()
def key_hotkey(keys: list[str]) -> dict[str, Any]:
    """
    按下组合键。如 ['ctrl', 'c'] 复制、['ctrl', 'shift', 't'] 恢复标签页。

    Args:
        keys: 按键列表，按顺序同时按下
    """
    return C.key_hotkey(keys)


@mcp.tool()
def type_text(text: str) -> dict[str, Any]:
    """输入文本（支持中文，通过剪贴板粘贴实现）。"""
    return C.type_text(text)


@mcp.tool()
def wait(milliseconds: int) -> dict[str, Any]:
    """等待若干毫秒。执行操作后经常需要等 UI 响应。"""
    return C.wait(milliseconds)


# ════════════════════════════════════════════════════════════════════
# 定位层 Tools
# ════════════════════════════════════════════════════════════════════

@mcp.tool()
def locate_text(text: str, match_mode: str = "contains",
                region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    在屏幕上用 OCR 搜索文字，返回所有匹配项的中心点坐标。
    配合 mouse_click 实现 "找到按钮 → 点击"。

    Args:
        text: 要搜索的文字（中文/英文都行）
        match_mode: 'exact' (完全) | 'contains' (包含，推荐) | 'fuzzy' (模糊 60% 相似度)
        region: 可选 [x,y,w,h] 限制搜索区域，加快速度
    """
    return L.locate_text(text, match_mode, region)


@mcp.tool()
def locate_image(image_path: str, confidence: float = 0.9,
                 region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    模板匹配 — 在屏幕上找目标图片的位置。
    适合 OCR 找不到的图标、按钮、图片元素。

    Args:
        image_path: 模板图片的本地路径
        confidence: 匹配阈值 0-1，默认 0.9
        region: 可选 [x,y,w,h] 限制区域
    """
    return L.locate_image(image_path, confidence, region)


@mcp.tool()
def locate_color(color: str, tolerance: int = 30,
                 region: Optional[list[int]] = None) -> dict[str, Any]:
    """
    颜色查找 — 在屏幕上找到指定颜色出现的位置。
    适合找红点提示、状态指示器、特定色块。

    Args:
        color: HEX 颜色如 '#FF5733' 或 'FF5733'
        tolerance: 颜色容差 0-255，默认 30
        region: 可选 [x,y,w,h] 限制区域
    """
    return L.locate_color(color, tolerance, region)


# ════════════════════════════════════════════════════════════════════
# 辅助层 Tools
# ════════════════════════════════════════════════════════════════════

@mcp.tool()
def execute_command(command: str, timeout: int = 10) -> dict[str, Any]:
    """
    执行系统 shell 命令。用于文件操作、进程查询等。

    Args:
        command: 命令行字符串
        timeout: 超时秒数，默认 10
    """
    return U.execute_command(command, timeout)


@mcp.tool()
def open_application(name: str) -> dict[str, Any]:
    """
    启动应用程序。name 可以是:
    - 开始菜单名称 (如 'notepad', 'calc', 'chrome')
    - 可执行文件完整路径
    - URL (用默认浏览器打开)
    """
    return U.open_application(name)


@mcp.tool()
def clipboard_read() -> dict[str, Any]:
    """读取系统剪贴板的文本内容。"""
    return U.clipboard_read()


@mcp.tool()
def clipboard_write(text: str) -> dict[str, Any]:
    """向系统剪贴板写入文本。"""
    return U.clipboard_write(text)


# ════════════════════════════════════════════════════════════════════
# 系统控制 Tools
# ════════════════════════════════════════════════════════════════════

@mcp.tool()
def set_safe_mode(enabled: bool) -> dict[str, Any]:
    """
    设置安全模式。开启后所有键鼠操作只打印不执行 — 调试时有用。

    Args:
        enabled: True=开启(只看不动)，False=关闭(实际操作)
    """
    _set_safe_mode(enabled)
    logger.info("safe_mode = %s", enabled)
    return {"success": True, "safe_mode": enabled}


# ─── 启动 ─────────────────────────────────────────────────────────────

def main():
    # 加载配置并应用
    cfg = load_config()
    _set_safe_mode(cfg.get("security", {}).get("safe_mode", False))

    logger.info("=" * 50)
    logger.info("Computer Agent MCP Server starting...")
    logger.info("  ScreenCapture: mss (fast)")
    logger.info("  OCR: PaddleOCR (lazy load, lang=%s)",
                cfg.get("ocr", {}).get("lang", "ch"))
    logger.info("  Keyboard/Mouse: pyautogui")
    logger.info("  Safe mode: %s", is_safe_mode_from_cfg())
    logger.info("=" * 50)

    # FastMCP.run() 默认用 stdio 传输
    mcp.run(transport="stdio")


def is_safe_mode_from_cfg() -> bool:
    from core.utils import is_safe_mode
    return is_safe_mode()


if __name__ == "__main__":
    main()
