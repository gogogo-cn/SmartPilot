' ============================================================
' SmartPilot desktop launcher (no console window).
' - Bridge already running -> just open the browser.
' - Deps missing (first run) -> fall back to start.bat (visible,
'   installs pywin32/openpyxl/PyYAML/rapidocr and prints errors).
' - Otherwise: start the bridge with pythonw.exe (hidden),
'   wait for http://127.0.0.1:9527/health, then open the browser.
' ============================================================
Option Explicit

Dim sh, fso, baseDir, bridge, pyw, i
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
baseDir = fso.GetParentFolderName(WScript.ScriptFullName)
bridge = baseDir & "\Tools\py-excel-bridge.py"
Log "launcher start, baseDir=" & baseDir

If Not fso.FileExists(bridge) Then
    Log "ERROR bridge not found"
    MsgBox "Bridge script not found: " & bridge, vbCritical, "SmartPilot"
    WScript.Quit 1
End If

' 1) already running -> open browser and exit
If HttpOk("http://127.0.0.1:9527/health") Then
    Log "already running, open browser"
    OpenBrowser
    WScript.Quit 0
End If

' 2) find pythonw.exe (no console window)
pyw = FindPythonW()
Log "pythonw=" & pyw
If pyw = "" Then
    Log "fallback start.bat (no pythonw)"
    sh.Run """" & baseDir & "\start.bat""", 1, False
    WScript.Quit 0
End If

' 3) dependency check with THE SAME interpreter that will run the bridge
'    (PATH "python" may be a different install without pywin32).
Dim pyExe, rc
pyExe = fso.GetParentFolderName(pyw) & "\python.exe"
If Not fso.FileExists(pyExe) Then pyExe = Replace(pyw, "pythonw.exe", "python.exe")
Log "deps check via " & pyExe
' 直接调 python.exe（Run 窗口样式 0 即隐藏控制台，不走 cmd 规避空格路径引号剥离坑）
rc = sh.Run("" & """" & pyExe & """ -c ""import win32com.client, openpyxl, yaml""", 0, True)
Log "deps rc=" & rc
If rc <> 0 Then
    Log "fallback start.bat (deps missing)"
    sh.Run """" & baseDir & "\start.bat""", 1, False
    WScript.Quit 0
End If

Log "start bridge via pythonw"
sh.Run """" & pyw & """ """ & bridge & """ --quiet", 0, False

' 4) wait until the bridge is ready (max ~25s), then open the browser
For i = 1 To 50
    WScript.Sleep 500
    If HttpOk("http://127.0.0.1:9527/health") Then Exit For
Next
OpenBrowser
WScript.Quit 0


Sub OpenBrowser()
    sh.Run "cmd /c start """" http://127.0.0.1:9527/", 0, False
End Sub

Sub Log(s)
    On Error Resume Next
    Dim f
    Set f = fso.OpenTextFile(sh.ExpandEnvironmentStrings("%TEMP%") & "\smartpilot-launcher.log", 8, True)
    f.WriteLine Now & "  " & s
    f.Close
    Set f = Nothing
    On Error GoTo 0
End Sub

Function HttpOk(url)
    On Error Resume Next
    HttpOk = False
    Dim x
    Set x = CreateObject("MSXML2.XMLHttp.6.0")
    x.open "GET", url, False
    x.send
    If Err.Number = 0 Then
        If x.status = 200 Then HttpOk = True
    End If
    Set x = Nothing
    On Error GoTo 0
End Function

Function FindPythonW()
    Dim candidates, root, sub_ , p
    candidates = Array( _
        "C:\Program Files\Python39\pythonw.exe", _
        "C:\Program Files\Python310\pythonw.exe", _
        "C:\Program Files\Python311\pythonw.exe", _
        "C:\Program Files\Python312\pythonw.exe", _
        "C:\Program Files\Python313\pythonw.exe", _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%\Programs\Python\Python39\pythonw.exe"), _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%\Programs\Python\Python310\pythonw.exe"), _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%\Programs\Python\Python311\pythonw.exe"), _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%\Programs\Python\Python312\pythonw.exe"), _
        sh.ExpandEnvironmentStrings("%LOCALAPPDATA%\Programs\Python\Python313\pythonw.exe"))
    Dim k
    For k = 0 To UBound(candidates)
        If candidates(k) <> "" Then
            If fso.FileExists(candidates(k)) Then
                FindPythonW = candidates(k)
                Exit Function
            End If
        End If
    Next
    ' registry hint from the py launcher (per-user / all-users installs)
    On Error Resume Next
    root = "HKLM\SOFTWARE\Python\PythonCore\"
    Dim vers, v
    vers = Array("3.9", "3.10", "3.11", "3.12", "3.13")
    For Each v In vers
        p = sh.RegRead(root & v & "\InstallPath\")
        If Err.Number = 0 Then
            If fso.FileExists(p & "pythonw.exe") Then
                FindPythonW = p & "pythonw.exe"
                Exit Function
            End If
        End If
        Err.Clear
        p = sh.RegRead("HKCU\SOFTWARE\Python\PythonCore\" & v & "\InstallPath\")
        If Err.Number = 0 Then
            If fso.FileExists(p & "pythonw.exe") Then
                FindPythonW = p & "pythonw.exe"
                Exit Function
            End If
        End If
        Err.Clear
    Next
    On Error GoTo 0
    FindPythonW = ""
End Function
