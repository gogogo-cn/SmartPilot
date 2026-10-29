@echo off
REM ============================================================
REM SmartPilot 桌面版启动器（本地桥 + 浏览器，端口 9527）
REM 首次运行自动安装 Python 依赖（pywin32 / openpyxl）
REM ============================================================
cd /d "%~dp0Tools"

set "PORT=9527"
set "PY="
REM 优先 py -3.9：桥依赖（pywin32 等）装在 3.9；PATH 上的 python 可能是别的版本
py -3.9 --version >nul 2>nul && set "PY=py -3.9"
if not defined PY ( py -3 --version >nul 2>nul && set "PY=py -3" )
if not defined PY ( python --version >nul 2>nul && set "PY=python" )
if not defined PY (
    echo [错误] 未检测到 Python。请安装 Python 3.10+（安装时勾选 "Add Python to PATH"）后重试。
    pause
    exit /b 1
)

%PY% -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo 首次运行：正在安装依赖 pywin32（需要联网）...
    %PY% -m pip install --user pywin32
)
%PY% -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo 首次运行：正在安装依赖 openpyxl（需要联网）...
    %PY% -m pip install --user openpyxl
)
%PY% -c "import yaml" >nul 2>nul
if errorlevel 1 (
    echo 首次运行：正在安装依赖 PyYAML（技能清单解析，需要联网）...
    %PY% -m pip install --user PyYAML
)
%PY% -c "from rapidocr_onnxruntime import RapidOCR" >nul 2>nul
if errorlevel 1 (
    echo 首次运行：正在安装 OCR 引擎 rapidocr-onnxruntime（约 15MB，需要联网）...
    %PY% -m pip install --user rapidocr-onnxruntime
)

REM 端口被占用说明已有实例在跑：新版安装后必须关掉旧实例重启才会生效
%WINDIR%\System32\netstat.exe -ano | %WINDIR%\System32\findstr.exe ":9527 .*LISTENING" >nul 2>nul
if not errorlevel 1 (
    echo [提示] 端口 9527 已被占用：SmartPilot 桌面版已有实例在运行。
    echo        刚更新过版本的话，请先关闭旧的黑色窗口（或结束 python.exe）再双击本图标，
    echo        否则页面连到的仍是旧版本。
    echo.
    start "" http://127.0.0.1:9527/
    pause
    exit /b 0
)

echo.
echo 正在启动 SmartPilot 桌面版（http://127.0.0.1:9527/）...
echo 使用期间请保持本窗口开启（最小化即可），关闭窗口即停止服务。
echo.
start "" http://127.0.0.1:9527/
%PY% py-excel-bridge.py --quiet
pause
