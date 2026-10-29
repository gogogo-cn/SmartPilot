Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports System.Net
Imports System.Net.Http
Imports Newtonsoft.Json

Namespace Services

    ''' <summary>
    ''' 【扩展6-架构预留】Python 脚本执行服务
    '''
    ''' 支持两种执行模式：
    '''   A) 本地 Python subprocess（需要本地安装 Python）
    '''   B) 远程 API 调用（推荐，安全隔离更好）
    '''
    ''' 使用方式：
    '''   1. 在 Skill 定义中设置 processingMode: "PythonCode"
    '''   2. templateCode 字段存放 Python 脚本代码
    '''   3. CustomSkillExecutor.ExecutePythonSkill() 会调用此服务
    ''' </summary>
    Public Class PythonExecutionService

        Private Shared ReadOnly _lazyInstance As New Lazy(Of PythonExecutionService)(
            Function() New PythonExecutionService(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As PythonExecutionService
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ' 共享 HttpClient，避免远程调用时套接字耗尽
        Private Shared ReadOnly _httpClient As HttpClient

        ' 静态构造函数：强制 TLS 1.2+ 并初始化 HttpClient
        Shared Sub New()
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
            ServicePointManager.CheckCertificateRevocationList = True

            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromSeconds(60)
            _httpClient.DefaultRequestHeaders.Add("Accept", "application/json")
        End Sub

        ' 配置项
        Private _pythonPath As String = "python"  ' 本地 Python 可执行文件路径
        Private _remoteEndpoint As String = ""     ' 远程 Python 服务端点 URL
        Private _executionMode As String = "local"  ' "local" 或 "remote"
        Private _timeoutSeconds As Integer = 60
        Private _pipInstallTimeoutSeconds As Integer = 180

        ' 包安装缓存：已验证/已安装的包名集合，避免每次都重新检查
        Private Shared ReadOnly _verifiedPackages As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _verifiedLock As New Object()

        ' 沙箱校验规则（正则表达式，IgnoreCase 匹配）
        Private Shared ReadOnly BlockedPatterns As String() = New String() {
            "\beval\s*\(",
            "\bexec\s*\(",
            "\bcompile\s*\(",
            "\b__import__\s*\(",
            "\bexecfile\s*\(",
            "\bfile\s*\(",
            "\bos\.system\b",
            "\bos\.popen\b",
            "\bos\.spawn\b",
            "\bos\.startfile\b",
            "\bsubprocess\.call\b",
            "\bsubprocess\.Popen\b",
            "\bsubprocess\.run\b",
            "\bctypes\.",
            "\bsys\.exit\b"
        }

        Private Shared ReadOnly HighRiskImports As String() = New String() {
            "os", "sys", "subprocess", "socket", "urllib", "urllib2", "requests",
            "httpx", "ftplib", "smtplib", "poplib", "imaplib", "telnetlib",
            "ctypes", "pickle", "cPickle", "marshal", "base64", "code", "codecs"
        }

        ' 注意：通用 import 规则已移除；pandas/numpy/matplotlib/sklearn 等数据分析库为正常 Low 风险导入
        Private Shared ReadOnly MediumRiskPatterns As String() = New String() {
            "\bopen\s*\(",
            "\bread_csv\b",
            "\bread_excel\b",
            "\bread_json\b",
            "\bpathlib\b",
            "\bshutil\.",
            "\bfile\s*\(",
            "\burllib\b",
            "\brequests\b",
            "\bhttpx\b",
            "\bsocket\b"
        }

        ' 导入名 -> pip 包名映射（自动安装时使用）
        Private Shared ReadOnly ImportToPipMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
            {"pandas", "pandas"},
            {"numpy", "numpy"},
            {"matplotlib", "matplotlib"},
            {"sklearn", "scikit-learn"},
            {"scipy", "scipy"},
            {"stats", "scipy"},
            {"seaborn", "seaborn"},
            {"openpyxl", "openpyxl"},
            {"yaml", "pyyaml"},
            {"PIL", "pillow"},
            {"cv2", "opencv-python"},
            {"requests", "requests"},
            {"sqlalchemy", "sqlalchemy"},
            {"xlrd", "xlrd"},
            {"xlwt", "xlwt"},
            {"plotly", "plotly"},
            {"bokeh", "bokeh"},
            {"statsmodels", "statsmodels"},
            {"sympy", "sympy"},
            {"networkx", "networkx"},
            {"dateutil", "python-dateutil"},
            {"pytz", "pytz"},
            {"docx", "python-docx"},
            {"pptx", "python-pptx"},
            {"lxml", "lxml"}
        }

        ' 标准库模块集合（不需要 pip 安装）
        Private Shared ReadOnly StdlibModules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "abc", "aifc", "argparse", "array", "ast", "asynchat", "asyncio", "asyncore",
            "atexit", "audioop", "base64", "bdb", "binascii", "binhex", "bisect", "builtins",
            "bz2", "calendar", "cgi", "cgitb", "chunk", "cmath", "cmd", "code", "codecs",
            "codeop", "collections", "colorsys", "compileall", "concurrent", "configparser",
            "contextlib", "contextvars", "copy", "copyreg", "cProfile", "crypt", "csv",
            "ctypes", "curses", "dataclasses", "datetime", "dbm", "decimal", "difflib",
            "dis", "distutils", "doctest", "email", "encodings", "enum", "errno",
            "faulthandler", "fcntl", "filecmp", "fileinput", "fnmatch", "fractions",
            "ftplib", "functools", "gc", "getopt", "getpass", "gettext", "glob", "grp",
            "gzip", "hashlib", "heapq", "hmac", "html", "http", "idlelib", "imaplib",
            "imghdr", "imp", "importlib", "inspect", "io", "ipaddress", "itertools",
            "json", "keyword", "lib2to3", "linecache", "locale", "logging", "lzma",
            "mailbox", "mailcap", "marshal", "math", "mimetypes", "mmap", "modulefinder",
            "multiprocessing", "netrc", "nis", "nntplib", "numbers", "operator", "optparse",
            "os", "ossaudiodev", "pathlib", "pdb", "pickle", "pickletools", "pipes",
            "pkgutil", "platform", "plistlib", "poplib", "posix", "posixpath", "pprint",
            "profile", "pstats", "pty", "pwd", "py_compile", "pyclbr", "pydoc",
            "queue", "quopri", "random", "re", "readline", "reprlib", "resource",
            "rlcompleter", "runpy", "sched", "secrets", "select", "selectors", "shelve",
            "shlex", "shutil", "signal", "site", "smtpd", "smtplib", "sndhdr", "socket",
            "socketserver", "spwd", "sqlite3", "sre_compile", "sre_constants", "sre_parse",
            "ssl", "stat", "statistics", "string", "stringprep", "struct", "subprocess",
            "sunau", "symtable", "sys", "sysconfig", "syslog", "tabnanny", "tarfile",
            "telnetlib", "tempfile", "termios", "test", "textwrap", "threading", "time",
            "timeit", "tkinter", "token", "tokenize", "trace", "traceback", "tracemalloc",
            "tty", "turtle", "turtledemo", "types", "typing", "unicodedata", "unittest",
            "urllib", "uu", "uuid", "venv", "warnings", "wave", "weakref", "webbrowser",
            "winreg", "winsound", "wsgiref", "xdrlib", "xml", "xmlrpc", "zipapp",
            "zipfile", "zipimport", "zlib", "_thread"
        }

        Private Sub New()
            LoadConfiguration()
        End Sub

        ''' <summary>
        ''' 从配置加载 Python 执行参数
        ''' 优先使用系统 Python → 按需下载的嵌入式 Python → 配置中的 Python 路径
        ''' </summary>
        Private Sub LoadConfiguration()
            Try
                Dim config = Utils.ConfigurationManager.Instance

                ' 优先使用系统 Python（通过 Bootstrapper 检测，支持 where/py launcher/注册表/Microsoft Store）
                Dim systemPython = PythonBootstrapper.Instance.DetectSystemPython()
                If Not String.IsNullOrEmpty(systemPython) Then
                    _pythonPath = systemPython
                    System.Diagnostics.Debug.WriteLine("[PythonService] 使用系统 Python: " & systemPython)
                Else
                    ' 降级：使用按需下载的嵌入式 Python
                    Dim bootstrapperPath = PythonBootstrapper.EmbeddedPythonExe
                    If Not String.IsNullOrEmpty(bootstrapperPath) AndAlso File.Exists(bootstrapperPath) Then
                        _pythonPath = bootstrapperPath
                        System.Diagnostics.Debug.WriteLine("[PythonService] 使用按需下载的 Python: " & bootstrapperPath)
                    Else
                        ' 最终降级：使用配置中的 Python 路径
                        _pythonPath = config.GetSetting("Python.ExecutionPath", "python")
                        System.Diagnostics.Debug.WriteLine("[PythonService] 使用配置的 Python: " & _pythonPath)
                    End If
                End If

                _remoteEndpoint = config.GetSetting("Python.RemoteEndpoint", "")
                _executionMode = config.GetSetting("Python.ExecutionMode", "local")
                _timeoutSeconds = CInt(Val(config.GetSetting("Python.TimeoutSeconds", "60")))
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[PythonService] 配置加载失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 执行 Python 脚本（主入口）
        ''' </summary>
        ''' <param name="scriptCode">Python 脚本代码</param>
        ''' <param name="inputData">输入数据（任意可 JSON 序列化对象，会注入为脚本中的 input_data 变量）</param>
        ''' <param name="requireConfirmation">中高风险脚本是否需要用户确认</param>
        ''' <param name="headers">表头信息（已弃用，优先使用 referenceData/selectionData 中的 headers）</param>
        ''' <param name="referenceData">引用区域数据列表，每个元素包含 workbook/sheet/address/headers/rows</param>
        ''' <param name="selectionData">当前选区数据对象，包含 workbook/sheet/address/headers/rows</param>
        ''' <returns>执行结果：包含 stdout、stderr、exitCode、outputData</returns>
    Public Async Function ExecuteScriptAsync(scriptCode As String,
                                              Optional inputData As Object = Nothing,
                                              Optional requireConfirmation As Boolean = True,
                                              Optional headers As List(Of String) = Nothing,
                                              Optional referenceData As Object = Nothing,
                                              Optional selectionData As Object = Nothing) As Task(Of PythonResult)
            Dim result As New PythonResult()

            ' 统一规范化上下文数据，确保 Python 侧拿到固定字段：workbook/sheet/address/headers/rows
            Dim normalizedReferenceData = NormalizeReferenceData(referenceData)
            Dim normalizedSelectionData = NormalizeSelectionData(selectionData)

            ' 1. 沙箱安全校验
            Dim checkResult = ValidateScript(scriptCode)
            If Not checkResult.IsSafe Then
                result.Success = False
                result.ErrorMessage = "Python 脚本未通过安全校验:" & Environment.NewLine &
                                      String.Join(Environment.NewLine, checkResult.Errors.Select(Function(e) "  - " & e))
                System.Diagnostics.Debug.WriteLine("[PythonService] 安全校验失败: " & result.ErrorMessage)
                Return result
            End If

            ' 2. 中高风险操作需用户确认
            '    前端「自动化免确认」开关开启时（AutoExecutionSettings.BypassConfirmation=True）跳过弹窗，
            '    供自动化任务自动执行；其余情况仍按原规则确认。
            If requireConfirmation AndAlso Not Global.ExcelAIVSTOAddIn.Security.AutoExecutionSettings.BypassConfirmation AndAlso
               (checkResult.RiskLevel = "Medium" OrElse checkResult.RiskLevel = "High") Then
                If Not ConfirmExecution(scriptCode, checkResult) Then
                    result.Success = False
                    result.ErrorMessage = "用户取消了 Python 脚本执行"
                    Return result
                End If
            End If

            Try
                Select Case _executionMode.ToLower()
                    Case "remote"
                        result = Await ExecuteRemoteAsync(scriptCode, inputData, headers, normalizedReferenceData, normalizedSelectionData).ConfigureAwait(False)
                    Case Else
                        result = Await ExecuteLocalAsync(scriptCode, inputData, headers, normalizedReferenceData, normalizedSelectionData).ConfigureAwait(False)
                End Select
            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "Python 执行异常: " & ex.Message
                System.Diagnostics.Debug.WriteLine("[PythonService] 执行失败: " & ex.ToString())
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 方案A：本地 subprocess 执行 Python 脚本
        '''
        ''' 流程:
        '''   1. 将 Excel 数据导出为临时 JSON 文件
        '''   2. 将用户 Python 代码包装为可执行脚本（注入 input_data 变量）
        '''   3. 通过 Process 启动 python 解释器执行
        '''   4. 捕获 stdout/stderr 和退出码
        '''   5. 解析输出中的 JSON 结果
        ''' </summary>
        Private Async Function ExecuteLocalAsync(scriptCode As String,
                                                   inputData As Object,
                                                   Optional headers As List(Of String) = Nothing,
                                                   Optional referenceData As List(Of Object) = Nothing,
                                                   Optional selectionData As Dictionary(Of String, Object) = Nothing) As Task(Of PythonResult)
            Dim result As New PythonResult()

            ' 创建临时目录和文件
            Dim tempDir = Path.Combine(Path.GetTempPath(), "SmartPilot_Python_" & Guid.NewGuid().ToString("N"))
            Directory.CreateDirectory(tempDir)

            Dim scriptPath = Path.Combine(tempDir, "script.py")
            Dim inputPath = Path.Combine(tempDir, "input.json")
            Dim outputPath = Path.Combine(tempDir, "output.json")

            System.Diagnostics.Debug.WriteLine($"[PythonService] 开始执行脚本: tempDir={tempDir}")
            System.Diagnostics.Debug.WriteLine($"[PythonService] Python路径: {_pythonPath}")
            System.Diagnostics.Debug.WriteLine($"[PythonService] 超时设置: {_timeoutSeconds}秒")
            System.Diagnostics.Debug.WriteLine($"[PythonService] 脚本长度: {scriptCode?.Length}字符")

            Try
                ' 检查 Python 路径是否有效
                If String.IsNullOrEmpty(_pythonPath) Then
                    result.Success = False
                    result.ErrorMessage = "Python 路径未配置。请在设置中配置 Python.ExecutionPath"
                    System.Diagnostics.Debug.WriteLine("[PythonService] 失败: Python 路径为空")
                    Return result
                End If

                If Not File.Exists(_pythonPath) Then
                    result.Success = False
                    result.ErrorMessage = $"Python 可执行文件不存在: {_pythonPath}"
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 失败: Python 文件不存在 - {_pythonPath}")
                    Return result
                End If

                ' 自动检测并安装脚本所需的第三方库
                Dim autoInstalled = Await EnsureRequiredPackagesAsync(scriptCode, _pythonPath).ConfigureAwait(False)
                If autoInstalled.Count > 0 Then
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 首次运行自动安装: {String.Join(", ", autoInstalled)}")
                End If

                ' 写入输入数据
                If inputData IsNot Nothing Then
                    Dim inputJson = JsonConvert.SerializeObject(inputData)
                    ' 将 NaN/Infinity 替换为 null，避免生成非法 JSON（Python json.load 会因 NaN 报错或产生 nan 污染统计结果）
                    inputJson = inputJson.Replace("-Infinity", "null").Replace("Infinity", "null").Replace("NaN", "null")
                    File.WriteAllText(inputPath, inputJson)
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 输入数据已写入: {inputPath}")
                End If

                ' 将用户脚本单独写入文件，供包装脚本 try/except 包裹执行（避免脚本异常导致进程崩溃、无法返回结构化错误）
                Dim userScriptPath = Path.Combine(tempDir, "user_script.py")
                File.WriteAllText(userScriptPath, scriptCode, New System.Text.UTF8Encoding(False))

                ' 包装用户脚本：添加标准 I/O 头尾，并注入 reference_data/selection_data
                Dim wrappedScript = WrapScriptForLocalExecution(scriptCode, inputPath, outputPath, headers, referenceData, selectionData, userScriptPath)
                File.WriteAllText(scriptPath, wrappedScript)
                System.Diagnostics.Debug.WriteLine($"[PythonService] 包装脚本已写入: {scriptPath}")
                System.Diagnostics.Debug.WriteLine($"[PythonService] 包装脚本长度: {wrappedScript.Length}字符")

                ' 创建进程并执行
                Dim psi As New ProcessStartInfo With {
                    .FileName = _pythonPath,
                    .Arguments = String.Format("-u ""{0}""", scriptPath),
                    .UseShellExecute = False,
                    .RedirectStandardOutput = True,
                    .RedirectStandardError = True,
                    .CreateNoWindow = True,
                    .WorkingDirectory = tempDir,
                    .StandardOutputEncoding = System.Text.Encoding.UTF8,
                    .StandardErrorEncoding = System.Text.Encoding.UTF8
                }

                ' 强制 Python 使用 UTF-8 编码输出，避免 Windows GBK 乱码
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"
                psi.EnvironmentVariables("PYTHONUTF8") = "1"

                System.Diagnostics.Debug.WriteLine($"[PythonService] 启动 Python 进程...")
                System.Diagnostics.Debug.WriteLine($"[PythonService] 参数: {psi.Arguments}")

                Using proc = Process.Start(psi)
                    If Not proc.WaitForExit(_timeoutSeconds * 1000) Then
                        proc.Kill()
                        result.Success = False
                        result.ErrorMessage = "Python 执行超时（" & _timeoutSeconds & "秒）。可能原因：脚本中存在死循环或计算过慢。"
                        System.Diagnostics.Debug.WriteLine($"[PythonService] 执行超时，进程已终止")
                        Return result
                    End If

                    result.StandardOutput = Await proc.StandardOutput.ReadToEndAsync().ConfigureAwait(False)
                    result.StandardError = Await proc.StandardError.ReadToEndAsync().ConfigureAwait(False)
                    result.ExitCode = proc.ExitCode

                    System.Diagnostics.Debug.WriteLine($"[PythonService] 进程退出码: {result.ExitCode}")
                    System.Diagnostics.Debug.WriteLine($"[PythonService] stdout 长度: {result.StandardOutput?.Length}")
                    System.Diagnostics.Debug.WriteLine($"[PythonService] stderr 长度: {result.StandardError?.Length}")
                End Using

                ' 读取输出结果
                If File.Exists(outputPath) Then
                    Try
                        Dim outputJson = File.ReadAllText(outputPath)
                        result.OutputData = Newtonsoft.Json.Linq.JObject.Parse(outputJson)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"[PythonService] 解析输出 JSON 失败: {ex.Message}")
                        ' 即使 JSON 解析失败，也尝试读取原始内容
                        Try
                            Dim rawContent = File.ReadAllText(outputPath)
                            result.RawOutput = rawContent
                        Catch
                        End Try
                    End Try
                End If

                ' 如果没有 output.json，尝试从 stdout 解析
                If result.OutputData Is Nothing AndAlso Not String.IsNullOrWhiteSpace(result.StandardOutput) Then
                    Try
                        result.OutputData = Newtonsoft.Json.Linq.JObject.Parse(result.StandardOutput.Trim())
                    Catch
                        result.RawOutput = result.StandardOutput
                    End Try
                End If

                ' 如果仍没有 OutputData，使用 stdout 作为原始输出
                If result.OutputData Is Nothing AndAlso result.RawOutput = "" Then
                    result.RawOutput = result.StandardOutput
                End If

                ' 修复：成功判定逻辑
                ' 只要退出码为 0 就视为成功，不再强制要求 OutputData 非空
                result.Success = (result.ExitCode = 0)

                ' 如果退出码为 0 但 OutputData 为空，使用 RawOutput 作为结果
                If result.Success AndAlso result.OutputData Is Nothing Then
                    result.OutputData = New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("result", If(result.RawOutput, "脚本执行成功"))
                    )
                End If

                ' 记录详细错误信息
                If Not result.Success Then
                    Dim errorDetails As New System.Text.StringBuilder()
                    errorDetails.AppendLine($"Python 脚本执行失败")
                    errorDetails.AppendLine($"退出码: {result.ExitCode}")
                    
                    If Not String.IsNullOrWhiteSpace(result.StandardError) Then
                        errorDetails.AppendLine($"错误详情:")
                        errorDetails.AppendLine(result.StandardError)
                        ' 分析常见错误类型
                        If result.StandardError.Contains("ModuleNotFoundError") Or result.StandardError.Contains("No module named") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: 缺少必要的 Python 库")
                            errorDetails.AppendLine("💡 建议: 运行 pip install 安装缺失的依赖包")
                        ElseIf result.StandardError.Contains("FileNotFoundError") Or result.StandardError.Contains("No such file") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: 引用的文件不存在")
                        ElseIf result.StandardError.Contains("SyntaxError") Or result.StandardError.Contains("IndentationError") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: 脚本语法错误")
                        ElseIf result.StandardError.Contains("ImportError") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: Python 环境配置问题")
                        ElseIf result.StandardError.Contains("PermissionError") Or result.StandardError.Contains("Access denied") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: 文件访问权限不足")
                        ElseIf result.StandardError.Contains("MemoryError") Then
                            errorDetails.AppendLine(vbCrLf & "💡 可能原因: 内存不足，脚本处理数据量过大")
                        Else
                            errorDetails.AppendLine(vbCrLf & "💡 请检查脚本逻辑或输入数据是否正确")
                        End If
                    End If
                    
                    result.ErrorMessage = errorDetails.ToString()
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 执行失败: {result.ErrorMessage}")
                ElseIf Not String.IsNullOrWhiteSpace(result.StandardError) Then
                    ' 即使成功，也记录 stderr 中的警告/信息
                    System.Diagnostics.Debug.WriteLine("[PythonService] 执行成功但有 stderr: " & result.StandardError)
                End If

                ' 写入调试日志，便于定位 Python 执行问题（含包装脚本、输入数据、stdout/stderr、退出码）
                Try
                    WritePythonDebugLog(tempDir, wrappedScript, inputPath, result)
                Catch
                End Try

            Finally
                ' 清理临时文件
                Try
                    If Directory.Exists(tempDir) Then Directory.Delete(tempDir, recursive:=True)
                Catch
                    System.Diagnostics.Debug.WriteLine("[PythonService] 清理临时文件失败")
                End Try
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 方案B：远程 API 调用（推荐用于生产环境）
        '''
        ''' 流程:
        '''   1. 将脚本和数据打包为 JSON POST 到远程端点
        '''   2. 远程服务在沙箱中执行 Python
        '''   3. 接收执行结果 JSON
        ''' </summary>
        Private Async Function ExecuteRemoteAsync(scriptCode As String,
                                                    inputData As Object,
                                                    Optional headers As List(Of String) = Nothing,
                                                    Optional referenceData As List(Of Object) = Nothing,
                                                    Optional selectionData As Dictionary(Of String, Object) = Nothing) As Task(Of PythonResult)
            Dim result As New PythonResult()

            If String.IsNullOrWhiteSpace(_remoteEndpoint) Then
                result.Success = False
                result.ErrorMessage = "未配置远程 Python 端点 (Python.RemoteEndpoint)"
                Return result
            End If

            Try
                Dim payload = New With {
                    .script = scriptCode,
                    .input_data = inputData,
                    .headers = headers,
                    .reference_data = referenceData,
                    .selection_data = selectionData,
                    .timeout = _timeoutSeconds
                }

                Using content As New StringContent(JsonConvert.SerializeObject(payload), Encoding.UTF8, "application/json")
                    Dim response = Await _httpClient.PostAsync(_remoteEndpoint, content).ConfigureAwait(False)
                    response.EnsureSuccessStatusCode()
                    Dim responseBytes = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

                    Dim responseObj = Newtonsoft.Json.Linq.JObject.Parse(responseBytes)

                    result.ExitCode = If(responseObj("exit_code")?.ToObject(Of Integer)(), -1)
                    result.StandardOutput = responseObj("stdout")?.ToString()
                    result.StandardError = responseObj("stderr")?.ToString()

                    If responseObj.ContainsKey("output_data") Then
                        result.OutputData = TryCast(responseObj("output_data"), Newtonsoft.Json.Linq.JObject)
                    End If

                    result.Success = (result.ExitCode = 0)
                    If Not result.Success Then
                        result.ErrorMessage = If(
                            result.StandardError,
                            "远程 Python 执行失败 (exit code: " & result.ExitCode & ")")
                    End If
                End Using

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "远程 Python 调用失败: " & ex.Message
            End Try

            Return result
        End Function

        ''' <summary>
        ''' 包装用户 Python 脚本为本地可执行形式
        ''' 注入 input_data 变量和结果输出逻辑
        ''' </summary>
        Private Function WrapScriptForLocalExecution(userScript As String,
                                                       inputJsonPath As String,
                                                       outputJsonPath As String,
                                                       Optional headers As List(Of String) = Nothing,
                                                       Optional referenceData As List(Of Object) = Nothing,
                                                       Optional selectionData As Dictionary(Of String, Object) = Nothing,
                                                       Optional userScriptPath As String = Nothing) As String
            Dim wrapper As New System.Text.StringBuilder()

            wrapper.AppendLine("# -*- coding: utf-8 -*-")
            wrapper.AppendLine("import json")
            wrapper.AppendLine("import sys")
            wrapper.AppendLine("import traceback")
            wrapper.AppendLine("import os")
            wrapper.AppendLine()

            ' 预置常用库自动导入（#2 借鉴外接 MCP）
            wrapper.AppendLine("# === SmartPilot 预置库（无需手动 import） ===")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    import pandas as pd")
            wrapper.AppendLine("except ImportError:")
            wrapper.AppendLine("    pd = None")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    import numpy as np")
            wrapper.AppendLine("except ImportError:")
            wrapper.AppendLine("    np = None")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    from scipy import stats")
            wrapper.AppendLine("except ImportError:")
            wrapper.AppendLine("    stats = None")
            wrapper.AppendLine()

            ' 图表自动展示（#4 借鉴外接 MCP）：设置 Agg 后端 + 拦截 plt.show() + 中文字体配置
            wrapper.AppendLine("# === SmartPilot 图表自动展示 ===")
            wrapper.AppendLine("plt = None")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    import matplotlib")
            wrapper.AppendLine("    matplotlib.use('Agg')")
            wrapper.AppendLine("    import matplotlib.pyplot as plt")
            wrapper.AppendLine("    import matplotlib.font_manager as fm")
            wrapper.AppendLine("except ImportError:")
            wrapper.AppendLine("    pass  # matplotlib 不可用时跳过图表功能")
            wrapper.AppendLine()

            ' 中文字体配置（仅当 matplotlib 可用时）
            wrapper.AppendLine("# === SmartPilot 中文字体配置 ===")
            wrapper.AppendLine("if plt is not None:")
            wrapper.AppendLine("    _cjk_fonts = ['Microsoft YaHei', 'SimHei', 'SimSun', 'KaiTi', 'FangSong',")
            wrapper.AppendLine("                 'Noto Sans CJK SC', 'WenQuanYi Zen Hei', 'Arial Unicode MS']")
            wrapper.AppendLine("    _available_fonts = {f.name for f in fm.fontManager.ttflist}")
            wrapper.AppendLine("    _selected_cjk = None")
            wrapper.AppendLine("    for _f in _cjk_fonts:")
            wrapper.AppendLine("        if _f in _available_fonts:")
            wrapper.AppendLine("            _selected_cjk = _f")
            wrapper.AppendLine("            break")
            wrapper.AppendLine("    if _selected_cjk:")
            wrapper.AppendLine("        plt.rcParams['font.sans-serif'] = [_selected_cjk, 'DejaVu Sans']")
            wrapper.AppendLine("        plt.rcParams['axes.unicode_minus'] = False")
            wrapper.AppendLine("    else:")
            wrapper.AppendLine("        plt.rcParams['font.sans-serif'] = ['DejaVu Sans']")
            wrapper.AppendLine("        plt.rcParams['axes.unicode_minus'] = False")
            wrapper.AppendLine()

            wrapper.AppendLine("_smartpilot_charts = []")
            wrapper.AppendLine("_smartpilot_chart_dir = os.path.join(os.environ.get('TEMP', '/tmp'), 'SmartPilot_Charts')")
            wrapper.AppendLine("os.makedirs(_smartpilot_chart_dir, exist_ok=True)")
            wrapper.AppendLine("_smartpilot_chart_idx = 0")
            wrapper.AppendLine()
            wrapper.AppendLine("def _smartpilot_show():")
            wrapper.AppendLine("    global _smartpilot_chart_idx")
            wrapper.AppendLine("    if plt is None:")
            wrapper.AppendLine("        print('[WARN] matplotlib 未安装，无法生成图表')")
            wrapper.AppendLine("        return")
            wrapper.AppendLine("    _smartpilot_chart_idx += 1")
            wrapper.AppendLine("    path = os.path.join(_smartpilot_chart_dir, f'chart_{_smartpilot_chart_idx}.png')")
            wrapper.AppendLine("    plt.savefig(path, dpi=150, bbox_inches='tight')")
            wrapper.AppendLine("    plt.close()")
            wrapper.AppendLine("    _smartpilot_charts.append(path)")
            wrapper.AppendLine("    print(f'[CHART]{path}')")
            wrapper.AppendLine()
            wrapper.AppendLine("if plt is not None:")
            wrapper.AppendLine("    plt.show = _smartpilot_show")
            wrapper.AppendLine()
            wrapper.AppendLine("def _save_chart(filename=None):")
            wrapper.AppendLine("    # 保存当前 matplotlib 图表，返回文件路径。推荐替代 plt.savefig() 直接调用。")
            wrapper.AppendLine("    global _smartpilot_chart_idx")
            wrapper.AppendLine("    if plt is None:")
            wrapper.AppendLine("        return None")
            wrapper.AppendLine("    _smartpilot_chart_idx += 1")
            wrapper.AppendLine("    if filename:")
            wrapper.AppendLine("        safe_name = os.path.splitext(os.path.basename(filename))[0]")
            wrapper.AppendLine("        path = os.path.join(_smartpilot_chart_dir, f'{safe_name}_{_smartpilot_chart_idx}.png')")
            wrapper.AppendLine("    else:")
            wrapper.AppendLine("        path = os.path.join(_smartpilot_chart_dir, f'chart_{_smartpilot_chart_idx}.png')")
            wrapper.AppendLine("    plt.savefig(path, dpi=150, bbox_inches='tight')")
            wrapper.AppendLine("    plt.close()")
            wrapper.AppendLine("    _smartpilot_charts.append(path)")
            wrapper.AppendLine("    print(f'[CHART]{path}')")
            wrapper.AppendLine("    return path")
            wrapper.AppendLine("# === 预置结束 ===")
            wrapper.AppendLine()

            ' 加载输入数据
            wrapper.AppendLine("# === SmartPilot 自动生成的输入加载代码 ===")
            wrapper.AppendLine("input_data = {}")
            Dim headersJson As String = If(headers IsNot Nothing, JsonConvert.SerializeObject(headers), "[]")
            ' 把 JSON 中的双引号转义为 Python 字符串中的 \"
            Dim quote As String = Chr(34)
            Dim backslash As String = "\"
            Dim escapedJson As String = headersJson.Replace(quote, backslash & quote)
            Dim pythonHeadersExpr As String = "json.loads(" & quote & escapedJson & quote & ")"
            wrapper.AppendLine("input_headers = " & pythonHeadersExpr)

            ' 注入引用区域数据（完整 rows，不是 sampleRows）
            Dim referenceJson As String = If(referenceData IsNot Nothing, JsonConvert.SerializeObject(referenceData), "[]")
            Dim escapedRefJson As String = referenceJson.Replace(quote, backslash & quote)
            Dim pythonReferenceExpr As String = "json.loads(" & quote & escapedRefJson & quote & ")"
            wrapper.AppendLine("reference_data = " & pythonReferenceExpr)

            ' 注入当前选区数据（提供安全访问方法）
            Dim selectionJson As String = If(selectionData IsNot Nothing, JsonConvert.SerializeObject(selectionData), "null")
            Dim escapedSelJson As String = selectionJson.Replace(quote, backslash & quote)
            Dim pythonSelectionExpr As String = "json.loads(" & quote & escapedSelJson & quote & ")"
            wrapper.AppendLine("selection_data = " & pythonSelectionExpr)
            wrapper.AppendLine()
            wrapper.AppendLine("# selection_data 安全访问助手")
            wrapper.AppendLine("def _sel(key, default=''):")
            wrapper.AppendLine("    if isinstance(selection_data, dict):")
            wrapper.AppendLine("        return selection_data.get(key, selection_data.get(key.lower(), default))")
            wrapper.AppendLine("    return default")
            wrapper.AppendLine()

            ' 修复：使用双引号包裹路径，正确转义反斜杠
            Dim safeInputPath = inputJsonPath.Replace("\", "\\").Replace("'", "\'")
            Dim safeOutputPath = outputJsonPath.Replace("\", "\\").Replace("'", "\'")
            
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    with open('" & safeInputPath & "', 'r', encoding='utf-8') as f:")
            wrapper.AppendLine("        input_data = json.load(f)")
            wrapper.AppendLine("        if isinstance(input_data, list) and len(input_data) > 0 and isinstance(input_data[0], dict):")
            wrapper.AppendLine("            input_headers = list(input_data[0].keys())")
            wrapper.AppendLine("except FileNotFoundError:")
            wrapper.AppendLine("    pass  # 无输入文件时使用空字典")
            wrapper.AppendLine("except Exception as e:")
            wrapper.AppendLine("    print(f'[INPUT_ERROR] {e}', file=sys.stderr)")
            wrapper.AppendLine("    pass")
        
            ' 用户脚本：单独文件 + try/except 包裹执行
            ' 任何脚本异常都会被捕获并写入结构化 output（success=False + message + traceback），
            ' 避免进程以非零码崩溃导致前端只能看到笼统的"模板执行错误"
            Dim safeUserScriptPath = userScriptPath.Replace("\", "\\").Replace("'", "\'")
            wrapper.AppendLine("# === 用户脚本开始 ===")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    exec(open('" & safeUserScriptPath & "', encoding='utf-8').read())")
            wrapper.AppendLine("except Exception as _sp_exc:")
            wrapper.AppendLine("    import traceback as _sp_tb")
            wrapper.AppendLine("    _sp_tb_str = _sp_tb.format_exc()")
            wrapper.AppendLine("    print(f'[SCRIPT_ERROR] {_sp_tb_str}', file=sys.stderr)")
            wrapper.AppendLine("    output = {'success': False, 'message': '脚本执行异常: ' + str(_sp_exc), 'traceback': _sp_tb_str}")
            wrapper.AppendLine("# === 用户脚本结束 ===")
            wrapper.AppendLine()
        
            ' 输出图表路径
            wrapper.AppendLine("# 输出图表信息")
            wrapper.AppendLine("if _smartpilot_charts:")
            wrapper.AppendLine("    for p in _smartpilot_charts:")
            wrapper.AppendLine("        print(f'[CHART]{p}')")
            wrapper.AppendLine()
        
            ' 简化版：查找 output 变量或 result 变量
            wrapper.AppendLine("# 收集输出结果（简化版）")
            wrapper.AppendLine("output_value = None")
            wrapper.AppendLine()
            wrapper.AppendLine("# 优先查找 output 变量")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    output_value = locals().get('output', None)")
            wrapper.AppendLine("except:")
            wrapper.AppendLine("    pass")
            wrapper.AppendLine()
            wrapper.AppendLine("# 如果 output 不存在，尝试 result 变量")
            wrapper.AppendLine("if output_value is None:")
            wrapper.AppendLine("    try:")
            wrapper.AppendLine("        output_value = locals().get('result', None)")
            wrapper.AppendLine("    except:")
            wrapper.AppendLine("        pass")
            wrapper.AppendLine()
            wrapper.AppendLine("# 如果仍没有，尝试 df 变量（pandas DataFrame）")
            wrapper.AppendLine("if output_value is None:")
            wrapper.AppendLine("    try:")
            wrapper.AppendLine("        df_val = locals().get('df', None)")
            wrapper.AppendLine("        if df_val is not None and hasattr(df_val, 'to_dict'):")
            wrapper.AppendLine("            output_value = df_val.to_dict()")
            wrapper.AppendLine("        elif df_val is not None:")
            wrapper.AppendLine("            output_value = df_val")
            wrapper.AppendLine("    except:")
            wrapper.AppendLine("        pass")
            wrapper.AppendLine()
            wrapper.AppendLine("if output_value is None:")
            wrapper.AppendLine("    output_value = '脚本执行成功，但未找到返回值。请确保你的脚本将结果赋值给 output 变量'")
            wrapper.AppendLine()
            wrapper.AppendLine("# 保存结果到文件")
            wrapper.AppendLine("try:")
            wrapper.AppendLine("    result_data = output_value")
            wrapper.AppendLine("    if not isinstance(result_data, (dict, list)):")
            wrapper.AppendLine("        result_data = {'result': result_data}")
            wrapper.AppendLine("    with open('" & safeOutputPath & "', 'w', encoding='utf-8') as f:")
            wrapper.AppendLine("        json.dump(result_data, f, ensure_ascii=False, indent=2, default=str)")
            wrapper.AppendLine("except Exception as e:")
            wrapper.AppendLine("    print(f'[OUTPUT_ERROR] {e}', file=sys.stderr)")
            wrapper.AppendLine("    import traceback")
            wrapper.AppendLine("    traceback.print_exc(file=sys.stderr)")
            wrapper.AppendLine("    sys.exit(1)")
        
            Return wrapper.ToString()
        End Function

        ''' <summary>
        ''' 将一次 Python 执行的关键信息写入固定调试日志，便于定位问题
        ''' 日志位置：%TEMP%\SmartPilot_Python_debug.log（追加模式）
        ''' </summary>
        Private Sub WritePythonDebugLog(tempDir As String, wrappedScript As String, inputPath As String, result As PythonResult)
            Dim logPath = Path.Combine(Path.GetTempPath(), "SmartPilot_Python_debug.log")
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine("=" & New String("="c, 70))
            sb.AppendLine("[Python 执行调试] " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            sb.AppendLine("tempDir: " & tempDir)
            sb.AppendLine("退出码: " & result.ExitCode)
            sb.AppendLine("Success: " & result.Success)
            sb.AppendLine("--- 输入数据（前 800 字符） ---")
            Try
                If File.Exists(inputPath) Then
                    Dim inputText = File.ReadAllText(inputPath)
                    sb.AppendLine(If(inputText.Length > 800, inputText.Substring(0, 800) & "...", inputText))
                Else
                    sb.AppendLine("(input.json 不存在)")
                End If
            Catch ex As Exception
                sb.AppendLine("读取输入数据失败: " & ex.Message)
            End Try
            sb.AppendLine("--- 包装脚本（前 1500 字符） ---")
            sb.AppendLine(If(wrappedScript IsNot Nothing AndAlso wrappedScript.Length > 1500, wrappedScript.Substring(0, 1500) & "...", wrappedScript))
            sb.AppendLine("--- stdout（前 1500 字符） ---")
            sb.AppendLine(If(result.StandardOutput IsNot Nothing AndAlso result.StandardOutput.Length > 1500, result.StandardOutput.Substring(0, 1500) & "...", result.StandardOutput))
            sb.AppendLine("--- stderr（前 3000 字符） ---")
            sb.AppendLine(If(result.StandardError IsNot Nothing AndAlso result.StandardError.Length > 3000, result.StandardError.Substring(0, 3000) & "...", result.StandardError))
            sb.AppendLine()
            File.AppendAllText(logPath, sb.ToString(), New System.Text.UTF8Encoding(False))
        End Sub

        ''' <summary>
        ''' 规范化引用区域数据：统一字段名为 workbook/sheet/address/headers/rows，并确保 rows 为完整数据
        ''' </summary>
        Private Function NormalizeReferenceData(referenceData As Object) As List(Of Object)
            Dim result As New List(Of Object)()
            If referenceData Is Nothing Then Return result

            Dim items As List(Of Object) = Nothing
            Try
                items = JsonConvert.DeserializeObject(Of List(Of Object))(JsonConvert.SerializeObject(referenceData))
            Catch
                Return result
            End Try

            If items Is Nothing Then Return result
            For Each item In items
                Dim normalized = NormalizeRegionData(item)
                If normalized IsNot Nothing Then result.Add(normalized)
            Next

            Return result
        End Function

        ''' <summary>
        ''' 规范化当前选区数据：统一字段名为 workbookName/sheetName/address/headers/rows
        ''' </summary>
        Private Function NormalizeSelectionData(selectionData As Object) As Dictionary(Of String, Object)
            If selectionData Is Nothing Then Return Nothing
            Return NormalizeRegionData(selectionData)
        End Function

        ''' <summary>
        ''' 将单个区域对象（可能来自前端 ReferenceAreaManager 或后端 ExcelDataBridge）规范化为统一字典
        ''' 输出字段：workbookName, sheetName, address, headers, rows
        ''' </summary>
        Private Function NormalizeRegionData(item As Object) As Dictionary(Of String, Object)
            If item Is Nothing Then Return Nothing

            Dim dict As Dictionary(Of String, Object) = Nothing
            Try
                dict = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(JsonConvert.SerializeObject(item))
            Catch
                Return Nothing
            End Try

            If dict Is Nothing Then Return Nothing

            Dim result As New Dictionary(Of String, Object)()
            result("workbookName") = GetStringValue(dict, "workbookName", "workbook")
            result("sheetName") = GetStringValue(dict, "sheetName", "sheet", "worksheet")
            result("address") = GetStringValue(dict, "address", "range")
            result("headers") = GetStringList(dict, "headers")
            result("rows") = GetObjectList(dict, "rows")

            ' 同时提供 sheet/workbook 别名，兼容旧版脚本
            result("sheet") = result("sheetName")
            result("workbook") = result("workbookName")
            Return result
        End Function

        Private Function GetStringValue(dict As Dictionary(Of String, Object), ParamArray keys As String()) As String
            If dict Is Nothing Then Return ""
            For Each key In keys
                If dict.ContainsKey(key) AndAlso dict(key) IsNot Nothing Then
                    Return dict(key).ToString()
                End If
            Next
            Return ""
        End Function

        Private Function GetStringList(dict As Dictionary(Of String, Object), key As String) As List(Of String)
            Dim result As New List(Of String)()
            If dict Is Nothing OrElse Not dict.ContainsKey(key) OrElse dict(key) Is Nothing Then Return result
            Try
                Dim list = JsonConvert.DeserializeObject(Of List(Of Object))(JsonConvert.SerializeObject(dict(key)))
                For Each item In list
                    If item IsNot Nothing Then result.Add(item.ToString())
                Next
            Catch
            End Try
            Return result
        End Function

        Private Function GetObjectList(dict As Dictionary(Of String, Object), key As String) As List(Of Object)
            Dim result As New List(Of Object)()
            If dict Is Nothing OrElse Not dict.ContainsKey(key) OrElse dict(key) Is Nothing Then Return result
            Try
                Return JsonConvert.DeserializeObject(Of List(Of Object))(JsonConvert.SerializeObject(dict(key)))
            Catch
                Return result
            End Try
        End Function

        ''' <summary>
        ''' 检查 Python 环境是否可用
        ''' </summary>
        Public Async Function CheckEnvironmentAsync() As Task(Of Dictionary(Of String, Object))
            Dim envInfo As New Dictionary(Of String, Object)()

            Try
                ' 检查本地 Python
                Dim localAvailable As Boolean = False
                Dim pythonVersion As String = ""

                Using proc = New Process()
                    proc.StartInfo.FileName = _pythonPath
                    proc.StartInfo.Arguments = "--version"
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    pythonVersion = Await proc.StandardOutput.ReadToEndAsync().ConfigureAwait(False)
                    proc.WaitForExit(5000)
                    localAvailable = (proc.ExitCode = 0)
                End Using

                envInfo("localAvailable") = localAvailable
                envInfo("pythonVersion") = pythonVersion.Trim()
                envInfo("executionMode") = _executionMode
                envInfo("remoteEndpointConfigured") = Not String.IsNullOrWhiteSpace(_remoteEndpoint)

                ' 测试远程连接
                If Not String.IsNullOrWhiteSpace(_remoteEndpoint) Then
                    Try
                        Dim healthUrl As String
                        If _remoteEndpoint.EndsWith("/") Then
                            healthUrl = _remoteEndpoint & "health"
                        Else
                            healthUrl = _remoteEndpoint & "/health"
                        End If
                        Dim testResponse = Await _httpClient.GetStringAsync(healthUrl).ConfigureAwait(False)
                        envInfo("remoteAvailable") = True
                    Catch
                        envInfo("remoteAvailable") = False
                    End Try
                Else
                    envInfo("remoteAvailable") = False
                End If

                envInfo("ready") = If(_executionMode = "remote", envInfo("remoteAvailable"), localAvailable)

            Catch ex As Exception
                envInfo("error") = ex.Message
                envInfo("ready") = False
            End Try

            Return envInfo
        End Function

        ''' <summary>
        ''' 对用户 Python 脚本进行静态安全校验
        ''' </summary>
        Private Function ValidateScript(scriptCode As String) As PythonSecurityCheckResult
            Dim result As New PythonSecurityCheckResult()
            If String.IsNullOrWhiteSpace(scriptCode) Then
                result.IsSafe = False
                result.RiskLevel = "High"
                result.Errors.Add("脚本内容为空")
                Return result
            End If

            Dim code As String = scriptCode

            ' 1. 完全阻断的极危险模式
            For Each pattern In BlockedPatterns
                If Regex.IsMatch(code, pattern, RegexOptions.IgnoreCase) Then
                    result.IsSafe = False
                    result.Errors.Add($"检测到高危操作: {pattern}")
                End If
            Next

            ' 2. 高风险导入
            For Each importName In HighRiskImports
                Dim importPattern As String = $"^\s*(import|from)\s+{Regex.Escape(importName)}\b"
                If Regex.IsMatch(code, importPattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline) Then
                    result.Warnings.Add($"检测到高风险模块导入: {importName}")
                    result.RiskLevel = "High"
                End If
            Next

            ' 3. 中风险模式（文件/网络/任意导入）
            For Each pattern In MediumRiskPatterns
                If Regex.IsMatch(code, pattern, RegexOptions.IgnoreCase Or RegexOptions.Multiline) Then
                    result.Warnings.Add($"检测到潜在风险操作: {pattern}")
                    If result.RiskLevel <> "High" Then
                        result.RiskLevel = "Medium"
                    End If
                End If
            Next

            ' 4. 如果没有发现任何风险，标记为 Low
            If result.RiskLevel = "None" AndAlso result.Warnings.Count = 0 Then
                result.RiskLevel = "Low"
            End If

            Return result
        End Function

        ''' <summary>
        ''' 在 Excel UI 线程上显示执行确认对话框（简化版，隐藏正则/代码等技术细节）
        ''' </summary>
        Private Function ConfirmExecution(scriptCode As String, checkResult As PythonSecurityCheckResult) As Boolean
            Dim confirmed As Boolean = False

            ' 生成用户可读的脚本概要
            Dim lineCount = scriptCode.Split({Environment.NewLine, vbLf}, StringSplitOptions.None).Length
            Dim operations = New List(Of String)()
            If Regex.IsMatch(scriptCode, "\bpd\.DataFrame\b|\bpandas\b", RegexOptions.IgnoreCase) Then operations.Add("数据分析")
            If Regex.IsMatch(scriptCode, "\bmatplotlib\b|\bplt\.\b|\bseaborn\b", RegexOptions.IgnoreCase) Then operations.Add("图表生成")
            If Regex.IsMatch(scriptCode, "\bsklearn\b|\bscikit-learn\b", RegexOptions.IgnoreCase) Then operations.Add("机器学习")
            If Regex.IsMatch(scriptCode, "\bopen\s*\(|\bread_csv\b|\bread_excel\b", RegexOptions.IgnoreCase) Then operations.Add("文件读取")
            If operations.Count = 0 Then operations.Add("数据处理")

            Dim message As String = String.Format(
                "AI 生成了一个 Python 脚本，是否执行？" & Environment.NewLine & Environment.NewLine &
                "脚本概要：" & Environment.NewLine &
                "  • 约 {0} 行代码" & Environment.NewLine &
                "  • 主要操作：{1}" & Environment.NewLine & Environment.NewLine &
                "点击「是」立即执行，点击「否」取消。",
                lineCount,
                String.Join("、", operations))

            ThisAddIn.RunOnUIThread(Sub()
                confirmed = (System.Windows.Forms.MessageBox.Show(
                    message,
                    "执行确认",
                    System.Windows.Forms.MessageBoxButtons.YesNo,
                    System.Windows.Forms.MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes)
            End Sub)

            Return confirmed
        End Function

#Region "自动安装缺失库"

        ''' <summary>
        ''' 从 Python 脚本中解析所有 import 的模块名
        ''' </summary>
        Private Function ExtractImports(scriptCode As String) As HashSet(Of String)
            Dim modules As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            Dim importRegex As New Regex("^\s*import\s+([a-zA-Z_][a-zA-Z0-9_]*)", RegexOptions.Multiline)
            For Each m As Match In importRegex.Matches(scriptCode)
                modules.Add(m.Groups(1).Value)
            Next

            Dim fromRegex As New Regex("^\s*from\s+([a-zA-Z_][a-zA-Z0-9_]*)\s+import", RegexOptions.Multiline)
            For Each m As Match In fromRegex.Matches(scriptCode)
                modules.Add(m.Groups(1).Value)
            Next

            Return modules
        End Function

        ''' <summary>
        ''' 检测并安装脚本所需的第三方库（按需自动安装）
        ''' </summary>
        ''' <param name="scriptCode">用户 Python 脚本代码</param>
        ''' <param name="pythonPath">Python 可执行文件路径</param>
        ''' <returns>本次新安装的包名列表</returns>
        Private Async Function EnsureRequiredPackagesAsync(scriptCode As String, pythonPath As String) As Task(Of List(Of String))
            Dim installedPackages As New List(Of String)()

            ' 收集需要检查的模块
            Dim modulesToCheck As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

            ' 预置 wrapper 脚本需要的库
            modulesToCheck.Add("pandas")
            modulesToCheck.Add("numpy")
            modulesToCheck.Add("matplotlib")

            ' 从用户脚本解析额外的 import
            Dim userImports = ExtractImports(scriptCode)
            For Each m In userImports
                modulesToCheck.Add(m)
            Next

            ' 过滤标准库
            Dim thirdParty = modulesToCheck.Where(Function(m) Not StdlibModules.Contains(m)).ToList()
            If thirdParty.Count = 0 Then Return installedPackages

            ' 检查哪些包缺失
            Dim missingPkgs As New List(Of String)()
            Dim pythonExe = pythonPath

            For Each modName In thirdParty
                Dim pipPkg = If(ImportToPipMap.ContainsKey(modName), ImportToPipMap(modName), modName)

                Try
                    Using checkProc = New Process()
                        checkProc.StartInfo.FileName = pythonExe
                        checkProc.StartInfo.Arguments = "-c """ & $"import {modName}" & """"
                        checkProc.StartInfo.UseShellExecute = False
                        checkProc.StartInfo.RedirectStandardOutput = True
                        checkProc.StartInfo.RedirectStandardError = True
                        checkProc.StartInfo.CreateNoWindow = True
                        checkProc.Start()
                        checkProc.WaitForExit(10000)

                        If checkProc.ExitCode <> 0 Then
                            missingPkgs.Add(pipPkg)
                        End If
                    End Using
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 检查 {modName} 失败: {ex.Message}")
                    missingPkgs.Add(pipPkg)
                End Try
            Next

            If missingPkgs.Count = 0 Then Return installedPackages

            ' 去重后安装
            Dim uniqueMissing = missingPkgs.Distinct().ToList()
            System.Diagnostics.Debug.WriteLine($"[PythonService] 检测到 {uniqueMissing.Count} 个缺失库: {String.Join(", ", uniqueMissing)}")

            For Each pkg In uniqueMissing
                System.Diagnostics.Debug.WriteLine($"[PythonService] 正在安装: {pkg}")
                Try
                    Using installProc = New Process()
                        installProc.StartInfo.FileName = pythonExe
                        installProc.StartInfo.Arguments = "-m pip install """ & pkg & """ --quiet"
                        installProc.StartInfo.UseShellExecute = False
                        installProc.StartInfo.RedirectStandardOutput = True
                        installProc.StartInfo.RedirectStandardError = True
                        installProc.StartInfo.CreateNoWindow = True
                        installProc.StartInfo.StandardOutputEncoding = System.Text.Encoding.UTF8
                        installProc.StartInfo.StandardErrorEncoding = System.Text.Encoding.UTF8
                        installProc.StartInfo.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"
                        installProc.StartInfo.EnvironmentVariables("PYTHONUTF8") = "1"
                        installProc.Start()

                        Dim stdoutTask = installProc.StandardOutput.ReadToEndAsync()
                        Dim stderrTask = installProc.StandardError.ReadToEndAsync()
                        installProc.WaitForExit(180000)

                        Dim stdout = Await stdoutTask.ConfigureAwait(False)
                        Dim stderr = Await stderrTask.ConfigureAwait(False)

                        If installProc.ExitCode = 0 Then
                            installedPackages.Add(pkg)
                            System.Diagnostics.Debug.WriteLine($"[PythonService] 安装成功: {pkg}")
                        Else
                            System.Diagnostics.Debug.WriteLine($"[PythonService] 安装失败: {pkg}")
                            System.Diagnostics.Debug.WriteLine($"  stderr: {stderr}")
                        End If
                    End Using
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[PythonService] 安装异常 {pkg}: {ex.Message}")
                End Try
            Next

            Return installedPackages
        End Function

#End Region

    End Class

    ''' <summary>
    ''' Python 脚本安全校验结果
    ''' </summary>
    Public Class PythonSecurityCheckResult
        Public Property IsSafe As Boolean = True
        Public Property RiskLevel As String = "None"
        Public Property Warnings As New List(Of String)()
        Public Property Errors As New List(Of String)()
    End Class

    ''' <summary>
    ''' Python 执行结果
    ''' </summary>
    Public Class PythonResult
        Public Property Success As Boolean = False
        Public Property ErrorMessage As String = ""
        Public Property StandardOutput As String = ""
        Public Property StandardError As String = ""
        Public Property ExitCode As Integer = -1
        Public Property OutputData As Newtonsoft.Json.Linq.JObject
        Public Property RawOutput As String = ""
    End Class

End Namespace