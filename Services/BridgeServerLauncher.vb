Option Strict On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Reflection
Imports System.Threading.Tasks

Namespace Services
    ''' <summary>
    ''' 容错启动本地桥 py-excel-bridge.py（端口 9527，Python + pywin32）。
    ''' 用途：为"独立网页版"（浏览器打开 UI/index.html）或共享桥场景提供开箱即用的本地桥，
    '''       消除手动运行 start.bat 的门槛。VSTO 插件内前端优先走 COM 宿主桥（App.excelBridge），
    '''       此桥仅作 HTTP fallback / 网页版入口，启动失败不影响插件主体。
    ''' 行为：
    '''   - 端口已占用 → 视为已运行，跳过（不重复拉起）。
    '''   - 未找到 python / py-excel-bridge.py → 仅记日志，静默跳过。
    '''   - 环境变量覆盖：EXCEL_AI_BRIDGE_PY（脚本路径）、EXCEL_AI_BRIDGE_PORT（端口）。
    ''' </summary>
    Public NotInheritable Class BridgeServerLauncher

        Private Shared ReadOnly _lock As New Object()
        Private Shared _process As Process = Nothing
        Private Shared _started As Boolean = False

        Public Shared ReadOnly Property IsStarted As Boolean
            Get
                SyncLock _lock
                    Return _started
                End SyncLock
            End Get
        End Property

        Public Shared Sub EnsureStarted()
            SyncLock _lock
                If _started Then Return
                _started = True
            End SyncLock
            Try
                Dim port As Integer = 9527
                Dim portEnv As String = Environment.GetEnvironmentVariable("EXCEL_AI_BRIDGE_PORT")
                If Not String.IsNullOrWhiteSpace(portEnv) Then
                    Integer.TryParse(portEnv.Trim(), port)
                End If

                LogLine($"触发本地桥自启动检测（端口={port}）")
                If IsPortInUse(port) Then
                    Debug.WriteLine($"[BridgeServerLauncher] 端口 {port} 已被占用，假定本地桥已运行，跳过自启动")
                    LogLine($"端口 {port} 已被占用，假定本地桥已运行，跳过自启动")
                    Return
                End If

                Dim scriptPath As String = ResolveScriptPath()
                If String.IsNullOrEmpty(scriptPath) OrElse Not File.Exists(scriptPath) Then
                    Dim skipMsg As String = "未找到 py-excel-bridge.py（开发/部署未附带 Tools 目录），跳过；网页版用户可手动运行 start.bat"
                    Debug.WriteLine("[BridgeServerLauncher] " & skipMsg)
                    LogLine(skipMsg)
                    Return
                End If

                Dim python As String = ResolvePython()
                If String.IsNullOrEmpty(python) Then
                    Dim noPyMsg As String = "未检测到 Python，跳过本地桥自启动；请安装 Python 3.10+ 后手动运行 start.bat"
                    Debug.WriteLine("[BridgeServerLauncher] " & noPyMsg)
                    LogLine(noPyMsg)
                    Return
                End If

                Dim psi As New ProcessStartInfo(python, """" & scriptPath & """") With {
                    .UseShellExecute = False,
                    .CreateNoWindow = True,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .WorkingDirectory = Path.GetDirectoryName(scriptPath)
                }
                psi.EnvironmentVariables("PORT") = port.ToString()

                Dim proc As Process = Process.Start(psi)
                If proc Is Nothing Then Return

                Dim logPath As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAIVSTOAddIn", "bridge_server.log")
                Try
                    Directory.CreateDirectory(Path.GetDirectoryName(logPath))
                Catch
                End Try
                AttachOutputLogging(proc, logPath)

                SyncLock _lock
                    _process = proc
                End SyncLock
                Debug.WriteLine($"[BridgeServerLauncher] 已启动本地桥 (pid={proc.Id}) → {scriptPath}")
                LogLine($"已启动本地桥 (pid={proc.Id}) python={python} script={scriptPath} port={port}")
            Catch ex As Exception
                Debug.WriteLine("[BridgeServerLauncher] 自启动失败（忽略）: " & ex.Message)
                LogLine("自启动失败（忽略）: " & ex.Message)
            End Try
        End Sub

        Public Shared Sub StopServer()
            SyncLock _lock
                If _process Is Nothing Then Return
                Dim proc As Process = _process
                _process = Nothing
                Try
                    If Not proc.HasExited Then
                        LogLine($"停止本地桥子进程 (pid={proc.Id})")
                        proc.Kill()
                        If Not proc.WaitForExit(3000) Then
                            Debug.WriteLine("[BridgeServerLauncher] 子进程未能在 3s 内退出")
                        End If
                    End If
                Catch ex As Exception
                    Debug.WriteLine("[BridgeServerLauncher] Stop 异常（忽略）: " & ex.Message)
                Finally
                    Try : proc.Dispose() : Catch : End Try
                End Try
            End SyncLock
        End Sub

        Private Shared Function IsPortInUse(port As Integer) As Boolean
            Try
                Using client As New TcpClient()
                    Dim ar As IAsyncResult = client.BeginConnect(IPAddress.Loopback, port, Nothing, Nothing)
                    If ar.AsyncWaitHandle.WaitOne(500) Then
                        client.EndConnect(ar)
                        Return True
                    End If
                    Return False
                End Using
            Catch
                Return False
            End Try
        End Function

        Private Shared Function ResolveScriptPath() As String
            Dim env As String = Environment.GetEnvironmentVariable("EXCEL_AI_BRIDGE_PY")
            If Not String.IsNullOrWhiteSpace(env) AndAlso File.Exists(env) Then Return env

            ' 1) AppData 部署目录（构建期由 SyncBridgeFiles 拷贝至此；与 DLL 实际加载位置
            '    无关，哪怕从 ClickOnce/VS 影子缓存加载也能稳定命中）
            Try
                Dim appDataTools As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAIVSTOAddIn", "Tools", "py-excel-bridge.py")
                If File.Exists(appDataTools) Then Return appDataTools
            Catch
            End Try

            ' 2) 程序集目录自身的 Tools 子目录（开发期 bin\Debug 直载场景）
            Try
                Dim asmDir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim direct As String = Path.Combine(asmDir, "Tools", "py-excel-bridge.py")
                If File.Exists(direct) Then Return direct
            Catch
            End Try

            ' 3) 从程序集目录向上回溯（开发期从 bin\Debug 上溯到仓库 ExcelAIVSTOAddIn\Tools）
            Try
                Dim dir As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                For i As Integer = 0 To 3
                    Dim candidate As String = Path.Combine(dir, "Tools", "py-excel-bridge.py")
                    If File.Exists(candidate) Then Return candidate
                    Dim parent As String = Path.GetDirectoryName(dir)
                    If String.IsNullOrEmpty(parent) Then Exit For
                    dir = parent
                Next
            Catch
            End Try
            Return ""
        End Function

        ''' <summary>
        ''' 仅探测、不下载的 Python 定位（复刻 McpManager.FindWorkingPython 的候选逻辑，但去掉
        ''' PythonBootstrapper 的按需下载副作用，避免插件启动时产生意外网络 I/O/阻塞）。
        ''' 优先级：config.ini 已验证路径 → 项目 vbenv → 用户/系统标准安装路径 → PATH 上的 python/python3。
        ''' </summary>
        Private Shared Function ResolvePython() As String
            Dim candidates As New List(Of String)()

            ' 1) config.ini 中 OCR 验证过的 Python 路径
            Try
                Dim configIni As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAIVSTOAddIn", "config.ini")
                If File.Exists(configIni) Then
                    For Each line As String In File.ReadAllLines(configIni)
                        If line.StartsWith("Python.ExecutionPath=") Then
                            Dim p As String = line.Substring("Python.ExecutionPath=".Length).Trim()
                            If File.Exists(p) Then candidates.Add(p)
                        End If
                    Next
                End If
            Catch
            End Try

            ' 2) 项目虚拟环境 vbenv
            Try
                Dim base As String = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
                Dim vbenv As String = Path.Combine(base, "vbenv", "Scripts", "python.exe")
                If File.Exists(vbenv) Then candidates.Add(vbenv)
            Catch
            End Try

            ' 3) 用户级 / 系统级标准安装路径（用环境变量派生，避免硬编码账号）
            Try
                Dim localAppData As String = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                Dim programFilesX86 As String = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
                Dim programW6432 As String = Environment.GetEnvironmentVariable("ProgramW6432")
                For Each ver As String In New String() {"Python313", "Python312", "Python311", "Python310", "Python39", "Python38"}
                    candidates.Add(Path.Combine(localAppData, "Programs\Python\", ver, "python.exe"))
                    candidates.Add(Path.Combine(localAppData, "Programs\", ver, "python.exe"))
                    If Not String.IsNullOrEmpty(programFilesX86) Then
                        candidates.Add(Path.Combine(programFilesX86, ver, "python.exe"))
                    End If
                    If Not String.IsNullOrEmpty(programW6432) AndAlso programW6432.ToLowerInvariant() <> programFilesX86.ToLowerInvariant() Then
                        candidates.Add(Path.Combine(programW6432, ver, "python.exe"))
                    End If
                Next
            Catch
            End Try

            ' 4) PATH 上的 python / python3
            For Each name As String In New String() {"python.exe", "python3.exe", "python", "python3"}
                Dim full As String = FindOnPath(name)
                If Not String.IsNullOrEmpty(full) Then candidates.Add(full)
            Next

            ' 返回第一个存在的候选（桥对 Python 版本容忍度高，仅做存在性校验即可，避免逐项 --version 拉起子进程拖慢启动）
            For Each p As String In candidates
                If Not String.IsNullOrEmpty(p) AndAlso File.Exists(p) Then Return p
            Next
            Return ""
        End Function

        Private Shared Function FindOnPath(name As String) As String
            Dim pathEnv As String = Environment.GetEnvironmentVariable("PATH")
            If String.IsNullOrEmpty(pathEnv) Then Return ""
            For Each dir As String In pathEnv.Split(Path.PathSeparator)
                Try
                    Dim full As String = Path.Combine(dir, name)
                    If File.Exists(full) Then Return full
                Catch
                End Try
            Next
            Return ""
        End Function

        Private Shared Sub AttachOutputLogging(proc As Process, logPath As String)
            Try
                Task.Run(Sub()
                    Try
                        Using w As New StreamWriter(logPath, True)
                            w.AutoFlush = True
                            w.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Bridge server started (pid={proc.Id})")
                            Dim line As String
                            line = proc.StandardOutput.ReadLine()
                            While line IsNot Nothing
                                w.WriteLine(line)
                                line = proc.StandardOutput.ReadLine()
                            End While
                        End Using
                    Catch
                    End Try
                End Sub)
                Task.Run(Sub()
                    Try
                        Using w As New StreamWriter(logPath, True)
                            w.AutoFlush = True
                            Dim line As String
                            line = proc.StandardError.ReadLine()
                            While line IsNot Nothing
                                w.WriteLine("[ERR] " & line)
                                line = proc.StandardError.ReadLine()
                            End While
                        End Using
                    Catch
                    End Try
                End Sub)
            Catch
            End Try
        End Sub

        ''' <summary>写入项目统一诊断日志 host_bridge.log（与桥方法惯例一致），便于确认自启动是否触发。</summary>
        Private Shared Sub LogLine(msg As String)
            Try
                Dim dir As String = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAIVSTOAddIn")
                Directory.CreateDirectory(dir)
                Using w As New StreamWriter(Path.Combine(dir, "host_bridge.log"), True)
                    w.AutoFlush = True
                    w.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [BridgeServerLauncher] {msg}")
                End Using
            Catch
            End Try
        End Sub

    End Class
End Namespace
