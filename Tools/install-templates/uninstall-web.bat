@echo off
echo Uninstalling SmartPilot desktop edition...

"%WINDIR%\System32\reg.exe" delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Uninstall\SmartPilot" /f >nul 2>nul

set "SMSM=%APPDATA%\Microsoft\Windows\Start Menu\Programs\SmartPilot"
if exist "%SMSM%" rd /s /q "%SMSM%"
if exist "%USERPROFILE%\Desktop\SmartPilot.lnk" del /f /q "%USERPROFILE%\Desktop\SmartPilot.lnk"

echo.
echo Uninstall complete.
echo The installation folder will be removed automatically in a few seconds.
echo.
pause

set "DELDIR=%~dp0"
if "%DELDIR:~-1%"=="\" set "DELDIR=%DELDIR:~0,-1%"
start "" /min cmd /c "timeout /t 3 /nobreak >nul & rd /s /q "%DELDIR%""
