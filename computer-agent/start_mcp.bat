@echo off
REM Computer Agent MCP Server launcher
REM Auto-locates .venv python, no hardcoded paths
setlocal

set "HERE=%~dp0"
set "PYTHON=%HERE%.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo [FATAL] venv not found: %PYTHON%
    exit /b 1
)

set PYTHONDONTWRITEBYTECODE=1
"%PYTHON%" -B "%HERE%mcp_server.py"
endlocal