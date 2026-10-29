@echo off
REM ============================================================
REM 启动本地桥（Python 版 py-excel-bridge.py，默认端口 9527）
REM
REM 适用场景：
REM   - 网页版（独立浏览器打开 UI/index.html）需要此桥，才能调用
REM     Excel 本地工具 / MCP 工具（/api/excel/* 、/api/mcp/*）。
REM   - VSTO 插件内前端走 COM 宿主桥（App.excelBridge），无需手动启动此桥。
REM
REM 说明：旧的 Tools/local-bridge-server.js（Node/winax）在 Node 20+ 无法编译，
REM       已被 py-excel-bridge.py（Python + pywin32）取代，二者 HTTP 契约一致。
REM ============================================================
cd /d "%~dp0Tools"

set "PORT=9527"

REM 探测 Python 解释器（优先 python，其次 python3；GitHub 某些环境需 py）
set "PY="
where python >nul 2>nul && set "PY=python"
if not defined PY ( where python3 >nul 2>nul && set "PY=python3" )
if not defined PY ( where py >nul 2>nul && set "PY=py -3" )
if not defined PY (
    echo [错误] 未检测到 Python。请安装 Python 3.10+ 并加入 PATH 后重试。
    pause
    exit /b 1
)

echo 正在启动本地桥 (Python) 于端口 %PORT% ...
echo 如需自定义端口，请先执行： set PORT=9531
%PY% py-excel-bridge.py
pause
