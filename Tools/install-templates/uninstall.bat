@echo off
echo Uninstalling SmartPilot for Office...
set "PS=%WINDIR%\System32\WindowsPowerShell\v1.0\powershell.exe"
if not exist "%PS%" set "PS=powershell"
"%PS%" -NoProfile -ExecutionPolicy Bypass -File "%~dp0register.ps1" -Unregister

"%WINDIR%\System32\reg.exe" delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\SmartPilot for Office" /f >nul 2>nul

set "SMSM=%APPDATA%\Microsoft\Windows\Start Menu\Programs\SmartPilot for Office"
if exist "%SMSM%" rd /s /q "%SMSM%"
if exist "%USERPROFILE%\Desktop\SmartPilot.lnk" del /f /q "%USERPROFILE%\Desktop\SmartPilot.lnk"

echo.
echo Uninstall complete. Please restart Office.
echo The installation folder will be removed automatically in a few seconds.
echo.
pause

set "DELDIR=%~dp0"
if "%DELDIR:~-1%"=="\" set "DELDIR=%DELDIR:~0,-1%"
start "" /min cmd /c "timeout /t 3 /nobreak >nul & rd /s /q "%DELDIR%""
