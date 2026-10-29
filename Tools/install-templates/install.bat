@echo off
echo Installing SmartPilot for Office (per-user, no administrator required)...
set "PS=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0register.ps1"
if errorlevel 1 (
    echo.
    echo [ERROR] Installation FAILED - see red messages above. Do NOT use the add-in yet.
) else (
    echo.
    echo Installation complete.
    echo Please FULLY close all Office apps (Excel / Word / PowerPoint), then reopen them.
    echo The add-ins will load automatically. (If Office was already open, restart it.)
)
echo.
pause
