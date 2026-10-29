Option Strict Off
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils

Namespace ExcelAIVSTOAddIn

    Partial Public Class ExcelBridge

#Region "OCR 安装与外部工具检测"

        <ComVisible(True)>
        Public Function InstallPaddleOcr() As Task(Of String)
            Try
                Dim pythonPath = GetPythonPath()
                If String.IsNullOrEmpty(pythonPath) Then
                    Return Task.FromResult(JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "未检测到 Python 环境，请先安装Python 3.8+",
                        .helpUrl = "https://www.python.org/downloads/"
                    }))
                End If

                _installProgressPath = Path.Combine(Path.GetTempPath(), "paddleocr_install_progress.txt")
                If File.Exists(_installProgressPath) Then File.Delete(_installProgressPath)

                _installCompleted = False
                _installSuccess = False

                System.Diagnostics.Debug.WriteLine($"[InstallPaddleOcr] 后台启动安装，Python: {pythonPath}")

                WriteInstallLog("开始安装PaddleOCR...")
                WriteInstallLog("这可能需要5-10 分钟，请耐心等待")

                ' 后台安装（有意 fire-and-forget）：不阻塞本 API 返回，安装进度经 GetPaddleOcrInstallProgress 轮询
                Dim installTask = Task.Run(Async Function()
                    Try
                        Dim progress = New Progress(Of String)(Sub(msg)
                            WriteInstallLog(msg)
                        End Sub)

                        Dim ocrService As New Services.PaddleOcrService()
                        Dim success = Await ocrService.InstallPaddleOcrAsync(progress)

                        _installCompleted = True
                        _installSuccess = success

                        If success Then
                            WriteInstallLog("安装成功")
                        Else
                            WriteInstallLog("安装失败")
                        End If
                    Catch ex As Exception
                        _installCompleted = True
                        _installSuccess = False
                        WriteInstallLog("安装异常: " & ex.Message)
                    End Try
                End Function)

                Return Task.FromResult(JsonConvert.SerializeObject(New With {
                    .success = True, .started = True,
                    .message = "安装已启动，正在后台运行..."
                }))

            Catch ex As Exception
                Return Task.FromResult(JsonConvert.SerializeObject(New With {.success = False, .error = $"启动失败: {ex.Message}"}))
            End Try
        End Function

        <ComVisible(True)>
        Public Function GetPaddleOcrInstallProgress() As String
            Try
                Dim result As New Dictionary(Of String, Object)
                result("completed") = _installCompleted
                result("success") = _installSuccess

                If File.Exists(_installProgressPath) Then
                    Dim lines = File.ReadAllLines(_installProgressPath)
                    Dim recentLines = lines.Skip(Math.Max(0, lines.Length - 30)).ToArray()
                    result("log") = String.Join(Environment.NewLine, recentLines)
                    result("totalLines") = lines.Length
                Else
                    result("log") = ""
                    result("totalLines") = 0
                End If

                Return JsonConvert.SerializeObject(result)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .completed = True, .success = False,
                    .log = "读取进度失败: " & ex.Message, .totalLines = 0
                })
            End Try
        End Function

        Private Sub WriteInstallLog(message As String)
            Try
                If Not String.IsNullOrEmpty(_installProgressPath) Then
                    Using writer As New StreamWriter(_installProgressPath, True, System.Text.Encoding.UTF8)
                        writer.WriteLine($"[{DateTime.Now.ToString("HH:mm:ss")}] {message}")
                    End Using
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private _installProgressPath As String = ""
        Private _installCompleted As Boolean = True
        Private _installSuccess As Boolean = False

        <ComVisible(True)>
        Public Sub RunPaddleOCRInstaller()
            Try
                System.Diagnostics.Debug.WriteLine("[RunPaddleOCRInstaller] 启动 PaddleOCR 安装程序")
                Task.Run(Sub()
                    Try
                        Dim pythonPath = GetPythonPath()
                        If String.IsNullOrEmpty(pythonPath) Then Return

                        Dim psi As New ProcessStartInfo With {
                            .FileName = pythonPath,
                            .Arguments = "-m pip install paddlepaddle paddleocr --quiet",
                            .UseShellExecute = False,
                            .CreateNoWindow = True
                        }
                        Using process As Process = Process.Start(psi)
                            process.WaitForExit(180000)
                            System.Diagnostics.Debug.WriteLine($"[RunPaddleOCRInstaller] 完成，退出码: {process.ExitCode}")
                        End Using
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[RunPaddleOCRInstaller] 后台安装失败: {ex.Message}")
                    End Try
                End Sub)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[RunPaddleOCRInstaller] 启动失败: {ex.Message}")
            End Try
        End Sub

        Private Function CheckPaddleOcrInstallation() As Object
            Try
                Dim allPaths As New List(Of String)
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each p In Utils.PathHelper.GetPythonSearchPaths()
                    If File.Exists(p) AndAlso seen.Add(p) Then allPaths.Add(p)
                Next

                Dim regPath = PythonLocator.FindFromRegistry()
                If Not String.IsNullOrEmpty(regPath) AndAlso seen.Add(regPath) Then allPaths.Add(regPath)

                For Each line In PythonLocator.FindAllOnPath()
                    If File.Exists(line) AndAlso seen.Add(line) Then allPaths.Add(line)
                Next

                For Each ver In {"3.12", "3.11", "3.10", "3.9", "3.14"}
                    Try
                        Using proc As New Process()
                            proc.StartInfo.FileName = "py"
                            proc.StartInfo.Arguments = $"-{ver} -c ""import sys; print(sys.executable)"""
                            proc.StartInfo.RedirectStandardOutput = True
                            proc.StartInfo.UseShellExecute = False
                            proc.StartInfo.CreateNoWindow = True
                            proc.Start()
                            Dim out = proc.StandardOutput.ReadToEnd().Trim()
                            proc.WaitForExit(3000)
                            If Not String.IsNullOrEmpty(out) AndAlso File.Exists(out) AndAlso seen.Add(out) Then
                                allPaths.Add(out)
                            End If
                        End Using
                    Catch ex As Exception  ' 
                    End Try
                Next

                Dim pythonFound = allPaths.Count > 0
                If Not pythonFound Then
                    Return New With {
                        .available = False,
                        .engine = "paddleocr",
                        .python_found = "False",
                        .message = "未检测到 Python"
                    }
                End If

                Dim modelDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".paddleocr")
                Dim modelsReady = Directory.Exists(modelDir) AndAlso
                                   Directory.GetDirectories(modelDir).Length > 0

                For Each pyPath In allPaths
                    If IsPaddleOcrInstalled(pyPath) Then
                        If modelsReady Then
                            Return New With {
                                .available = True,
                                .engine = "paddleocr",
                                .python_found = "True",
                                .message = "PaddleOCR 已安装并可用 (" & Path.GetFileName(Path.GetDirectoryName(pyPath)) & ")"
                            }
                        Else
                            Return New With {
                                .available = False,
                                .engine = "paddleocr",
                                .python_found = "True",
                                .models_missing = True,
                                .message = "PaddleOCR 已安装，但模型未下载。请点击重新安装以下载模型"
                            }
                        End If
                    End If
                Next

                Return New With {
                    .available = False,
                    .engine = "paddleocr",
                    .python_found = "True",
                    .message = "PaddleOCR 未安装（已找到" & allPaths.Count & " 个Python，均不可用）"
                }

            Catch ex As Exception
                Return New With {
                    .available = False,
                    .engine = "paddleocr",
                    .python_found = "False",
                    .message = "检测失败: " & ex.Message
                }
            End Try
        End Function

        <ComVisible(True)>
        Public Function GetPaddleOcrStatus() As String
            Try
                Dim result = CheckPaddleOcrInstallation()
                Return JsonConvert.SerializeObject(result)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .available = False,
                    .engine = "paddleocr",
                    .python_found = "False",
                    .message = "检测异常: " & ex.Message
                })
            End Try
        End Function

        Private Function GetPythonPath() As String
            Return PythonLocator.ResolvePreferredPath()
        End Function

        Private Function IsPaddleOcrInstalled(pythonPath As String) As Boolean
            Try
                Dim process As New System.Diagnostics.Process()
                process.StartInfo.FileName = pythonPath
                process.StartInfo.Arguments = "-c ""import paddleocr; print('OK')"""
                process.StartInfo.RedirectStandardOutput = True
                process.StartInfo.RedirectStandardError = True
                process.StartInfo.UseShellExecute = False
                process.StartInfo.CreateNoWindow = True
                process.Start()

                Dim stdOutTask = process.StandardOutput.ReadToEndAsync()
                Dim stdErrTask = process.StandardError.ReadToEndAsync()

                Dim finished = process.WaitForExit(30000)

                If Not finished Then
                    System.Diagnostics.Debug.WriteLine("PaddleOCR检测: 超时30秒")
                    Try : process.Kill() : Catch : End Try
                    Return False
                End If

                Dim output = stdOutTask.Result
                Dim [error] = stdErrTask.Result

                System.Diagnostics.Debug.WriteLine($"PaddleOCR检测: exit={process.ExitCode} out={output.Trim()} err={[error].Trim()}")

                Return process.ExitCode = 0 AndAlso output.Contains("OK")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("PaddleOCR安装检查失败: " & ex.Message)
                Return False
            End Try
        End Function

        <ComVisible(True)>
        Public Function DiagnosePaddleOcr() As String
            Try
                Dim results As New List(Of Object)

                Dim allPaths As New List(Of String)
                Dim seen As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

                For Each p In Utils.PathHelper.GetPythonSearchPaths()
                    If File.Exists(p) AndAlso seen.Add(p) Then allPaths.Add(p)
                Next

                Dim regPath = PythonLocator.FindFromRegistry()
                If Not String.IsNullOrEmpty(regPath) AndAlso seen.Add(regPath) Then allPaths.Add(regPath)

                For Each line In PythonLocator.FindAllOnPath()
                    If File.Exists(line) AndAlso seen.Add(line) Then allPaths.Add(line)
                Next

                For Each ver In {"3.12", "3.11", "3.10", "3.9", "3.14"}
                    Try
                        Using proc As New Process()
                            proc.StartInfo.FileName = "py"
                            proc.StartInfo.Arguments = $"-{ver} -c ""import sys; print(sys.executable)"""
                            proc.StartInfo.RedirectStandardOutput = True
                            proc.StartInfo.UseShellExecute = False
                            proc.StartInfo.CreateNoWindow = True
                            proc.Start()
                            Dim out = proc.StandardOutput.ReadToEnd().Trim()
                            proc.WaitForExit(3000)
                            If Not String.IsNullOrEmpty(out) AndAlso File.Exists(out) AndAlso seen.Add(out) Then
                                allPaths.Add(out)
                            End If
                        End Using
                    Catch ex As Exception  ' 
                    End Try
                Next

                Dim anySuccess = False
                For Each pyPath In allPaths
                    Dim result = TestPythonForPaddleOcr(pyPath)
                    results.Add(result)
                    If CBool(result("paddleocr_ok")) Then anySuccess = True
                Next

                Return JsonConvert.SerializeObject(New With {
                    .total_pythons = allPaths.Count,
                    .any_success = anySuccess,
                    .details = results
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .error = ex.Message,
                    .total_pythons = 0,
                    .any_success = False,
                    .details = {}
                })
            End Try
        End Function

        Private Function TestPythonForPaddleOcr(pythonPath As String) As Dictionary(Of String, Object)
            Dim result As New Dictionary(Of String, Object)
            result("path") = pythonPath
            result("python_ok") = False
            result("paddleocr_ok") = False
            result("version") = ""
            result("pip_paddleocr") = False
            result("import_error") = ""

            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = pythonPath
                    proc.StartInfo.Arguments = "--version"
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    Dim stdOut = proc.StandardOutput.ReadToEndAsync()
                    Dim stdErr = proc.StandardError.ReadToEndAsync()
                    proc.WaitForExit(5000)
                    Dim ver = (stdOut.Result & stdErr.Result).Trim()
                    result("python_ok") = (proc.ExitCode = 0)
                    result("version") = ver
                End Using
            Catch ex As Exception
                result("import_error") = "无法启动: " & ex.Message
                Return result
            End Try

            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = pythonPath
                    proc.StartInfo.Arguments = "-m pip list"
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    Dim output = proc.StandardOutput.ReadToEnd()
                    proc.WaitForExit(10000)
                    result("pip_paddleocr") = output.ToLower().Contains("paddleocr")
                End Using
            Catch ex As Exception  ' 
            End Try

            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = pythonPath
                    proc.StartInfo.Arguments = "-c ""import paddleocr; print('OK')"""
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    Dim stdOutTask = proc.StandardOutput.ReadToEndAsync()
                    Dim stdErrTask = proc.StandardError.ReadToEndAsync()
                    Dim finished = proc.WaitForExit(15000)
                    If Not finished Then
                        Try : proc.Kill() : Catch : End Try
                        result("import_error") = "超时15秒"
                    Else
                        Dim out = stdOutTask.Result.Trim()
                        Dim [err] = stdErrTask.Result.Trim()
                        If proc.ExitCode = 0 AndAlso out.Contains("OK") Then
                            result("paddleocr_ok") = True
                        Else
                            result("import_error") = If([err], $"退出码={proc.ExitCode}")
                        End If
                    End If
                End Using
            Catch ex As Exception
                result("import_error") = ex.Message
            End Try

            System.Diagnostics.Debug.WriteLine($"[Diagnose] {pythonPath}: python={result("python_ok")} pip={result("pip_paddleocr")} import={result("paddleocr_ok")}")
            Return result
        End Function

#End Region

#Region "Python 代码执行 API"

        <ComVisible(True)>
        Public Function ExecutePython(pythonCode As String) As String
            Try
                If String.IsNullOrWhiteSpace(pythonCode) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Python 代码不能为空"})
                End If

                Dim pythonPath = GetPythonPath()
                If String.IsNullOrEmpty(pythonPath) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "未检测到 Python 环境。请先安装 Python 3.8+，或检查 PATH 配置",
                        .helpUrl = "https://www.python.org/downloads/"
                    })
                End If

                System.Diagnostics.Debug.WriteLine($"[ExecutePython] 使用 Python: {pythonPath}")

                Dim tempFile = Path.Combine(Path.GetTempPath(), $"ExcelAI_Python_{Guid.NewGuid()}.py")
                File.WriteAllText(tempFile, pythonCode, System.Text.Encoding.UTF8)

                Dim psi As New ProcessStartInfo With {
                    .FileName = pythonPath,
                    .Arguments = $"""{tempFile}""",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .StandardOutputEncoding = System.Text.Encoding.UTF8,
                    .StandardErrorEncoding = System.Text.Encoding.UTF8
                }

                Dim output As String = ""
                Dim errorOutput As String = ""
                Dim exitCode As Integer = -1

                Using process As Process = Process.Start(psi)
                    output = process.StandardOutput.ReadToEnd()
                    errorOutput = process.StandardError.ReadToEnd()
                    process.WaitForExit(60000)
                    exitCode = process.ExitCode
                End Using

                Try
                    If File.Exists(tempFile) Then File.Delete(tempFile)
                Catch ex As Exception  ' 
                End Try

                Dim success = (exitCode = 0)
                Dim resultMessage = If(success, output, If(String.IsNullOrEmpty(errorOutput), "执行失败（无错误输出）", errorOutput))

                Dim response = New Dictionary(Of String, Object) From {
                    {"success", success},
                    {"result", resultMessage},
                    {"stdout", output},
                    {"stderr", errorOutput},
                    {"exitCode", exitCode},
                    {"pythonPath", pythonPath},
                    {"executedAt", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}
                }

                Dim settings = New JsonSerializerSettings With {
                    .ContractResolver = New CamelCasePropertyNamesContractResolver(),
                    .Formatting = Formatting.Indented
                }

                Return JsonConvert.SerializeObject(response, settings)

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ExecutePython] 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        <ComVisible(True)>
        Public Function CheckPythonEnvironment() As String
            Try
                Dim pythonPath = GetPythonPath()
                If String.IsNullOrEmpty(pythonPath) Then
                    Return JsonConvert.SerializeObject(New With {
                        .available = False,
                        .pythonFound = False,
                        .message = "未检测到 Python 环境"
                    })
                End If

                Dim psi As New ProcessStartInfo With {
                    .FileName = pythonPath,
                    .Arguments = "--version",
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .CreateNoWindow = True
                }

                Dim version As String = ""
                Using process As Process = Process.Start(psi)
                    version = process.StandardOutput.ReadToEnd().Trim()
                    process.WaitForExit(5000)
                End Using

                Return JsonConvert.SerializeObject(New With {
                    .available = True,
                    .pythonFound = True,
                    .path = pythonPath,
                    .version = version,
                    .message = $"Python 已就绪: {version}"
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .available = False,
                    .pythonFound = False,
                    .message = "检测异常 " & ex.Message
                })
            End Try
        End Function

#End Region

    End Class

End Namespace