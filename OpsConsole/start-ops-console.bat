@echo off
setlocal
cd /d %~dp0

REM ============================================================
REM  SmartPilot OpsConsole one-click launcher
REM  Double-click -> start server -> open http://127.0.0.1:9600
REM  Close the "SmartPilot OpsConsole" window to stop the server.
REM  Custom port: set OPS_PORT=9610 first, then double-click.
REM
REM  NOTE: keep this file ASCII-only. cmd parses .bat using the
REM  system ANSI codepage; any non-ASCII line gets split into
REM  garbage "commands" and the window dies instantly (flash
REM  crash). The node child window switches to UTF-8 (chcp
REM  65001) so its Chinese logs still display correctly.
REM
REM  NOTE: readiness is probed with node itself, never curl.
REM  curl honors http_proxy/HTTP_PROXY env vars, so on a machine
REM  behind a proxy it reports success against a dead port and
REM  the launcher skips starting the server entirely (the
REM  browser then opens a page that does not exist).
REM ============================================================

if not defined OPS_PORT set "OPS_PORT=9600"

REM ---- 1. Find Node.js (PATH -> WorkBuddy bundled -> system install) ----
set "NODE="
where node >nul 2>nul && set "NODE=node"
if not defined NODE if exist "C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2-6\node.exe" set "NODE=C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2-6\node.exe"
if not defined NODE if exist "C:\Program Files\nodejs\node.exe" set "NODE=C:\Program Files\nodejs\node.exe"
if not defined NODE (
    echo [FAIL] Node.js not found.
    echo Install Node.js 18+ from https://nodejs.org then run this file again.
    pause
    exit /b 1
)

REM ---- 2. Already running? Just open the browser ----
call :Probe
if not errorlevel 1 (
    echo Server already running - opening browser...
    start "" http://127.0.0.1:%OPS_PORT%
    echo This window can be closed now.
    call :Sleep 3
    exit /b 0
)

REM ---- 3. Start server in its own window (cmd /k keeps it open after exit) ----
echo Starting SmartPilot OpsConsole...
if "%NODE%"=="node" (
    start "SmartPilot OpsConsole" cmd /k "chcp 65001 >nul & node server.js"
) else (
    start "SmartPilot OpsConsole" cmd /k "chcp 65001 >nul & ""%NODE%"" server.js"
)

REM ---- 4. Wait for readiness, then open browser (up to 15 tries) ----
set /a TRIES=0
:WAIT
set /a TRIES+=1
call :Sleep 1
call :Probe
if not errorlevel 1 (
    start "" http://127.0.0.1:%OPS_PORT%
    echo Opened http://127.0.0.1:%OPS_PORT%
    echo This window can be closed now. Close the server window to stop OpsConsole.
    call :Sleep 5
    exit /b 0
)
if %TRIES% geq 15 (
    echo [TIMEOUT] Server did not respond within 15 seconds.
    echo Check the "SmartPilot OpsConsole" window for errors ^(common cause: port %OPS_PORT% in use^).
    echo Use another port: run "set OPS_PORT=9610" first, then double-click this file again.
    pause
    exit /b 1
)
goto WAIT

REM ---- 5. Sleep helper ----
REM timeout.exe refuses to wait when stdin is redirected (prints a localized
REM error and returns instantly). Fall back to ping, which never touches stdin:
REM ping -n N 127.0.0.1 waits N-1 seconds.
:Sleep
set /a SLEEPN=%~1+1
"%SystemRoot%\System32\timeout.exe" /t %~1 /nobreak >nul 2>nul || ping -n %SLEEPN% 127.0.0.1 >nul
goto :eof

REM ---- 6. Probe helper: HTTP 200 on /api/overview means the console is up ----
REM Uses node's http module directly: no proxy env vars, no curl dependency.
REM The 3-second dead-man switch covers a port that answers but hangs.
:Probe
"%NODE%" -e "var t=setTimeout(function(){process.exit(1)},3000);require('http').get('http://127.0.0.1:%OPS_PORT%/api/overview',function(r){clearTimeout(t);process.exit(r.statusCode===200?0:1)}).on('error',function(){process.exit(1)})" >nul 2>nul
goto :eof
