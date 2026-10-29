"""
辅助层 Tool: execute_command / open_application / clipboard_*
"""
from __future__ import annotations

import subprocess
import sys
from typing import Any

sys.path.insert(0, __import__("os").path.dirname(__import__("os").path.dirname(__file__)))

from core.utils import is_safe_mode


def execute_command(command: str, timeout: int = 10) -> dict[str, Any]:
    """
    执行 shell 命令（同步，有超时）。

    Args:
        command: 命令行字符串
        timeout: 超时秒数

    Returns:
        {"success": True, "stdout": "...", "stderr": "...", "returncode": 0}
    """
    if is_safe_mode():
        return {"action": "execute_command", "command": command,
                "executed": False, "safe_mode": True}

    try:
        result = subprocess.run(
            command,
            shell=True,
            capture_output=True,
            text=True,
            timeout=timeout,
            encoding="utf-8",
            errors="replace",
        )
        return {
            "success": True,
            "stdout": result.stdout.strip()[:5000],
            "stderr": result.stderr.strip()[:2000],
            "returncode": result.returncode,
        }
    except subprocess.TimeoutExpired:
        return {"success": False, "error": f"命令超时 (>{timeout}s)"}
    except Exception as e:
        return {"success": False, "error": str(e)}


def open_application(name: str) -> dict[str, Any]:
    """
    启动应用程序。name 可以是:
      - Windows 开始菜单名称 (如 'notepad', 'calculator')
      - 可执行文件完整路径
      - URL（会用默认浏览器打开）
    """
    if is_safe_mode():
        return {"action": "open_application", "name": name,
                "executed": False, "safe_mode": True}

    try:
        if sys.platform == "win32":
            import os
            if name.startswith(("http://", "https://", "file://")):
                os.startfile(name)
            elif os.path.exists(name):
                os.startfile(name)
            else:
                # 用 start 命令，能识别开始菜单注册名
                subprocess.Popen(f"start ""{name}""", shell=True)
        else:
            subprocess.Popen([name], shell=True)
        return {"success": True, "application": name, "executed": True}
    except Exception as e:
        return {"success": False, "error": str(e)}


def clipboard_read() -> dict[str, Any]:
    """读取剪贴板文本"""
    import pyperclip
    try:
        text = pyperclip.paste()
        return {"success": True, "text": text, "length": len(text)}
    except Exception as e:
        return {"success": False, "error": str(e)}


def clipboard_write(text: str) -> dict[str, Any]:
    """写入剪贴板"""
    import pyperclip
    if is_safe_mode():
        return {"action": "clipboard_write", "length": len(text),
                "executed": False, "safe_mode": True}
    try:
        pyperclip.copy(text)
        return {"success": True, "length": len(text)}
    except Exception as e:
        return {"success": False, "error": str(e)}
