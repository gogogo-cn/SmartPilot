@echo off
:: SmartPilot 本地桥一键启动（pywin32 驱动真实 Excel）
:: 用法：双击本文件，或命令行 start-bridge.bat
cd /d "%~dp0"

echo ============================================================
echo  SmartPilot 本地代理服务 (pywin32 驱动真实 Excel COM)
echo ============================================================

:: 检查 pywin32 是否可用，缺失则尝试安装
python -c "import win32com" >nul 2>&1
if errorlevel 1 (
    echo [提示] 未检测到 pywin32，正在尝试安装...
    pip install pywin32
    if errorlevel 1 (
        echo [错误] pywin32 安装失败，请手动执行: pip install pywin32
        pause
        exit /b 1
    )
)

echo 启动服务（默认端口 9527）...
echo 浏览器打开: http://localhost:9527/index.html
echo 按 Ctrl+C 停止服务。
echo ============================================================
python py-excel-bridge.py
pause
