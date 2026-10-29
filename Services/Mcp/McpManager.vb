Option Explicit On
Option Strict Off
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services.Mcp

    '#Region "MCP 管理器"

    ''' <summary>
    ''' MCP 统一管理器 - 管理 Server 和 Clients
    ''' 作为 ExcelBridge 与 MCP 系统之间的中间层
    ''' </summary>
    Public Class McpManager
        Implements IDisposable

        Private _serverHost As McpServerHost
        Private _clientPool As McpClientPool
        Private _configManager As McpConfigManager
        Private _bridge As Object
        ' 【新增】用于取消后台任务的 Token
        Private _cancellationTokenSource As System.Threading.CancellationTokenSource
        Private _autoConnectTask As Task
    Private _bootstrapTask As Task
        Private _isDisposed As Boolean = False

        ''' <summary>
        ''' MCP 配置管理器
        ''' </summary>
        Public ReadOnly Property Config As McpConfigManager
            Get
                Return _configManager
            End Get
        End Property

        ''' <summary>
        ''' Server 主机
        ''' </summary>
        Public ReadOnly Property ServerHost As McpServerHost
            Get
                Return _serverHost
            End Get
        End Property

        ''' <summary>
        ''' Client 连接池
        ''' </summary>
        Public ReadOnly Property ClientPool As McpClientPool
            Get
                Return _clientPool
            End Get
        End Property

        Public Sub New(bridge As Object)
            _bridge = bridge
            _configManager = New McpConfigManager()
            _cancellationTokenSource = New System.Threading.CancellationTokenSource()
            Dim port = _configManager.GetServerPort()
            _serverHost = New McpServerHost(bridge, port)
            _clientPool = New McpClientPool(_configManager)

            ' 自动加载已保存的 MCP Client 配置（不自动连接，避免启动变慢）
            Try
                Dim savedClients = _configManager.GetServerConfigs()

                If savedClients IsNot Nothing AndAlso savedClients.Count > 0 Then
                    For Each kvp In savedClients
                        Dim clientConfig = kvp.Value
                        If clientConfig.Enabled Then
                            _clientPool.AddClientToPoolOnly(clientConfig)
                        End If
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"McpManager: 加载保存的 Client 失败: {ex.Message}")
            End Try

            ' 不在这里建立实际连接，保持完全懒加载；
            ' 前端通过 js-app-init.js 根据 autoStartClients 配置决定是否调用 AutoReconnect。
            _autoConnectTask = Task.CompletedTask
            Dim clientCount As Integer = If(_clientPool IsNot Nothing, _clientPool.Count, 0)
            System.Diagnostics.Debug.WriteLine($"McpManager 初始化完成（已加载 {clientCount} 个 client 配置，连接由前端按需触发）")

            ' Computer Agent 已内置于 .NET（ComputerAgentTools.vb），无需外部 bootstrap
            ' 旧的 BootstrapComputerAgentAsync 调用已废弃（原外部 Python MCP Server 方案残留）
        End Sub


        ''' <summary>从现有 OCR / PythonBootstrapper 配置里找已验证的 Python 路径
        ''' 要求: >= 3.10 (mcp 2.x 需要), 排除 3.14 (太新)
        ''' </summary>
        Private Shared Function FindWorkingPython() As String
            ' 候选池：config.ini (OCR 验证过的) → PythonBootstrapper 扫描
            Dim candidates = New List(Of String)()

            Try
                Dim configIni = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAIVSTOAddIn", "config.ini")
                If File.Exists(configIni) Then
                    For Each line In File.ReadAllLines(configIni)
                        If line.StartsWith("Python.ExecutionPath=") Then
                            Dim p = line.Substring("Python.ExecutionPath=".Length).Trim()
                            If File.Exists(p) Then candidates.Add(p)
                        End If
                    Next
                End If
            Catch
            End Try

            Try
                Dim pythonService = New Services.PythonBootstrapper()
                Dim path = pythonService.DetectSystemPython()
                If Not String.IsNullOrEmpty(path) AndAlso File.Exists(path) Then
                    candidates.Add(path)
                End If
            Catch
            End Try

            ' 再补一轮标准安装路径兜底，且使用当前用户/机器的环境变量派生路径，避免硬编码特定账号
            ' （原硬编码 C:\Users\Administrator\... 移他机即失效；PythonBootstrapper 已扫描系统级标准路径，
            ' 这里仅补 USER 级 LocalAppData 安装位，覆盖以当前登录用户身份 pip 安装的 Python。）
            Dim userExtras As New List(Of String)()
            Try
                Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                Dim programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86)
                Dim programW6432 = Environment.GetEnvironmentVariable("ProgramW6432")
                For Each ver In {"Python313", "Python312", "Python311", "Python310", "Python39", "Python38"}
                    userExtras.Add(Path.Combine(localAppData, "Programs\Python\", ver, "python.exe"))
                    userExtras.Add(Path.Combine(localAppData, "Programs\", ver, "python.exe"))
                    If Not String.IsNullOrEmpty(programFilesX86) Then
                        userExtras.Add(Path.Combine(programFilesX86, ver, "python.exe"))
                    End If
                    If Not String.IsNullOrEmpty(programW6432) AndAlso programW6432.ToLowerInvariant() <> programFilesX86.ToLowerInvariant() Then
                        userExtras.Add(Path.Combine(programW6432, ver, "python.exe"))
                    End If
                Next
            Catch
            End Try
            For Each p In userExtras
                If Not String.IsNullOrEmpty(p) AndAlso File.Exists(p) AndAlso Not candidates.Contains(p) Then candidates.Add(p)
            Next

            ' 逐个验证版本，选第一个 3.10 <= v < 3.14
            For Each p In candidates
                Dim v = GetPythonVersion(p)
                If v IsNot Nothing AndAlso v.Major = 3 AndAlso v.Minor >= 10 AndAlso v.Minor < 14 Then
                    Return p  ' 3.10, 3.11, 3.12, 3.13 OK
                End If
            Next
            For Each p In candidates
                Dim v = GetPythonVersion(p)
                If v IsNot Nothing AndAlso v.Major = 3 AndAlso v.Minor = 9 Then
                    ' 3.9 也勉强能用（mcp<2 或我们有 mcp 2.x 需升级）
                    Return p
                End If
            Next
            Return Nothing
        End Function

        Private Shared Function GetPythonVersion(pythonPath As String) As Version
            Try
                Using proc As New Process()
                    proc.StartInfo.FileName = pythonPath
                    proc.StartInfo.Arguments = "--version"
                    proc.StartInfo.UseShellExecute = False
                    proc.StartInfo.RedirectStandardOutput = True
                    proc.StartInfo.RedirectStandardError = True
                    proc.StartInfo.CreateNoWindow = True
                    proc.Start()
                    ' --version 输出到 stderr
                    Dim err = proc.StandardError.ReadToEnd().Trim()
                    proc.WaitForExit(5000)
                    ' 格式: "Python 3.12.9"
                    Dim match = System.Text.RegularExpressions.Regex.Match(err, "(\d+)\.(\d+)")
                    If match.Success Then
                        Return New Version(match.Groups(1).Value & "." & match.Groups(2).Value & ".0")
                    End If
                End Using
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[GetPythonVersion] 失败: " & ex.Message)
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' 自动重连所有已保存的 Client（可在 UI 中手动触发）
        ''' </summary>
        Public Async Function AutoReconnectAsync() As Task(Of String)
            Try
                ' 先从磁盘重新加载最新配置（支持外部修改配置文件后刷新）
                _configManager.ReloadConfig()
                Await _clientPool.AutoConnectAllAsync()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "自动重连完成"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        '#Region "Server 管理"

        ''' <summary>
        ''' 启动 MCP Server
        ''' </summary>
        Public Async Function StartServerAsync(Optional port As Integer = -1) As Task(Of String)
            Return Await Task.Run(Of String)(Function()
                Try
                    If _serverHost.IsRunning Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .error = "Server 已在运行中"
                        })
                    End If

                    If port > 0 Then
                        ' 使用指定端口重建
                        _serverHost.Dispose()
                        _serverHost = New McpServerHost(_bridge, port)
                        _configManager.SetServerPort(port)
                    End If

                    _serverHost.StartAsync()
                    _configManager.SetServerEnabled(True)

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = $"MCP Server 已启动在端口 {_serverHost.Port}",
                        .port = _serverHost.Port
                    })

                Catch ex As Exception
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = ex.Message
                    })
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 停止 MCP Server
        ''' </summary>
        Public Async Function StopServerAsync() As Task(Of String)
            Try
                If Not _serverHost.IsRunning Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "Server 未在运行"
                    })
                End If

                Await _serverHost.StopAsync()
                _configManager.SetServerEnabled(False)

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "MCP Server 已停止"
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        '#End Region

        '#Region "Client 管理"

        ''' <summary>
        ''' 添加并连接 MCP Client
        ''' </summary>
        Public Async Function AddClientAsync(clientConfigJson As String) As Task(Of String)
            Try
                Dim config = JsonConvert.DeserializeObject(Of McpClientConfig)(clientConfigJson)
                If config Is Nothing OrElse String.IsNullOrEmpty(config.Name) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "配置无效：需要提供 name"
                    })
                End If

                ' 保存配置
                _configManager.SaveServerConfigAsync(config.Name, config)

                ' 连接
                Dim connected = Await _clientPool.AddClientAsync(config)

                If connected Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = $"已连接 MCP Server: {config.Name}",
                        .toolCount = _clientPool.GetClientTools(config.Name).Count
                    })
                Else
                    Dim client = _clientPool.GetClient(config.Name)
                    Dim errMsg As String = "未知错误"
                    If client IsNot Nothing AndAlso Not String.IsNullOrEmpty(client.ErrorMessage) Then
                        errMsg = client.ErrorMessage
                    End If
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"连接失败: {errMsg}"
                    })
                End If

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 移除 MCP Client
        ''' </summary>
        Public Async Function RemoveClientAsync(name As String) As Task(Of String)
            Try
                Await _clientPool.RemoveClientAsync(name)
                _configManager.RemoveServerConfig(name)

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = $"已移除并断开: {name}"
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 断开 MCP Client（保留配置）
        ''' </summary>
        Public Async Function DisconnectClientAsync(name As String) As Task(Of String)
            Try
                Dim removed = Await _clientPool.DisconnectClientAsync(name)
                If removed Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = $"已断开: {name}"
                    })
                Else
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"Client '{name}' 不存在"
                    })
                End If
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 测试 Client 连接（支持传入完整配置 JSON 或仅名称）
        ''' </summary>
        Public Async Function TestClientConnectionAsync(configOrName As String) As Task(Of String)
            Try
                ' 尝试解析为配置 JSON
                Dim testConfig As McpClientConfig = Nothing
                Dim testName As String = configOrName

                If configOrName.TrimStart().StartsWith("{") Then
                    Try
                        testConfig = JsonConvert.DeserializeObject(Of McpClientConfig)(configOrName)
                        If testConfig IsNot Nothing AndAlso Not String.IsNullOrEmpty(testConfig.Name) Then
                            testName = testConfig.Name
                        End If
                    Catch ex As Exception  '  不是有效 JSON，当作名称处理
                    End Try
                End If

                Dim client = _clientPool.GetClient(testName)

                If testConfig IsNot Nothing Then
                    ' 有完整配置：创建临时 client 测试，不保存到 mcp.user.json
                    If client IsNot Nothing Then
                        Await _clientPool.RemoveClientAsync(testName)
                    End If
                    Dim connected = Await _clientPool.AddClientAsync(testConfig)
                    client = _clientPool.GetClient(testName)

                    If connected AndAlso client IsNot Nothing AndAlso client.IsConnected Then
                        Dim canPing = Await client.PingAsync()
                        ' 测试用临时 client，测试后移除
                        Await _clientPool.RemoveClientAsync(testName)
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .connected = True,
                            .canPing = canPing,
                            .toolCount = client.GetTools().Count,
                            .message = "连接正常"
                        })
                    Else
                        Dim errMsg = "连接失败"
                        If client IsNot Nothing AndAlso Not String.IsNullOrEmpty(client.ErrorMessage) Then
                            errMsg = client.ErrorMessage
                        End If
                        Await _clientPool.RemoveClientAsync(testName)
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .connected = False,
                            .error = errMsg
                        })
                    End If
                Else
                    ' 仅名称：从配置加载
                    If client Is Nothing Then
                        Dim config = _configManager.GetServerConfig(testName)
                        If config Is Nothing Then
                            Return JsonConvert.SerializeObject(New With {
                                .success = False,
                                .error = $"MCP Server '{testName}' 不存在"
                            })
                        End If
                        Await _clientPool.AddClientAsync(config)
                        client = _clientPool.GetClient(testName)
                    End If

                    If client IsNot Nothing AndAlso client.IsConnected Then
                        Dim canPing = Await client.PingAsync()
                        Return JsonConvert.SerializeObject(New With {
                            .success = True,
                            .connected = True,
                            .canPing = canPing,
                            .toolCount = client.GetTools().Count,
                            .message = "连接正常"
                        })
                    Else
                        Dim errMsg As String = "未连接"
                        If client IsNot Nothing AndAlso Not String.IsNullOrEmpty(client.ErrorMessage) Then
                            errMsg = client.ErrorMessage
                        End If
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .connected = False,
                            .error = errMsg
                        })
                    End If
                End If

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取所有 Client 的工具列表
        ''' </summary>
        Public Function GetClientTools(name As String) As String
            Try
                Dim tools = _clientPool.GetClientTools(name)
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .tools = tools,
                    .count = tools.Count
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 调用外部工具
        ''' </summary>
        Public Async Function CallToolAsync(clientName As String, toolName As String, argumentsJson As String) As Task(Of String)
            Try
                Dim arguments = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argumentsJson)
                If arguments Is Nothing Then
                    arguments = New Dictionary(Of String, Object)()
                End If

                Dim result = Await _clientPool.CallToolAsync(clientName, toolName, arguments)

                ' 提取文本内容，避免复杂对象序列化问题
                Dim resultText = ""
                If result.Content IsNot Nothing AndAlso result.Content.Count > 0 Then
                    Dim textContents = result.Content.Where(Function(c) TypeOf c Is McpTextContent).Cast(Of McpTextContent)()
                    resultText = String.Join(vbCrLf, textContents.Select(Function(c) c.Text))
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = Not result.IsError,
                    .result = resultText,
                    .isError = result.IsError
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 调用本地 Excel 工具（通过 McpServer.ExcelTools.ExecuteTool 直接执行，不经过 Client Pool）
        ''' </summary>
        Public Function CallLocalTool(toolName As String, argumentsJson As String) As String
            Try
                Dim arguments = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(argumentsJson)
                If arguments Is Nothing Then
                    arguments = New Dictionary(Of String, Object)()
                End If

                If _serverHost Is Nothing OrElse _serverHost.Server Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "MCP Server 未启动，无法调用本地工具"
                    })
                End If

                ' 按工具名前缀分发：ppt_ → PPT，computer_ → 电脑 Agent，web_ → 联网搜索，
                ' time_ → 系统时间，其余 → Excel
                Dim result As ToolsCallResult
                If toolName.StartsWith("ppt_", StringComparison.OrdinalIgnoreCase) Then
                    result = _serverHost.Server.PptTools.ExecuteTool(toolName, arguments)
                ElseIf toolName.StartsWith("computer_", StringComparison.OrdinalIgnoreCase) Then
                    result = ComputerAgentTools.ExecuteTool(toolName, arguments)
                ElseIf toolName.StartsWith("web_", StringComparison.OrdinalIgnoreCase) Then
                    result = WebSearchTools.ExecuteTool(toolName, arguments)
                ElseIf toolName.StartsWith("time_", StringComparison.OrdinalIgnoreCase) Then
                    result = TimeTools.ExecuteTool(toolName, arguments)
                Else
                    result = _serverHost.Server.ExcelTools.ExecuteTool(toolName, arguments)
                End If

                ' 提取文本内容
                Dim resultText = ""
                If result.Content IsNot Nothing AndAlso result.Content.Count > 0 Then
                    Dim textContents = result.Content.Where(Function(c) TypeOf c Is McpTextContent).Cast(Of McpTextContent)()
                    resultText = String.Join(vbCrLf, textContents.Select(Function(c) c.Text))
                End If

                ' 【修复 代码溢出】MCP 本地工具出口统一清洗：剥掉 LLM 会误抄进总结回复的样式字典 JSON / CSS 尾巴
                ' 只清洗 "非结构化/人类可读文本" 类工具的 result；对纯 JSON 类工具做保护
                ' PPT 工具全部返回结构化 JSON（主题配置/设计推荐/生成结果），统一保留
                Dim jsonToolsNeedPreserve As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
                    "excel_get_workbook_info", "excel_get_open_workbooks", "excel_read_range", "excel_read_selection",
                    "excel_split_by_column", "excel_split_workbook", "excel_aggregate", "excel_merge_workbooks",
                    "excel_import_csv", "excel_export_csv", "excel_get_used_range", "excel_list_files", "excel_search",
                    "excel_clean_data", "excel_remove_duplicates", "excel_check_errors",
                    "excel_create_slicer",
                    "web_search",
                    "time_now",
                    "ppt_get_templates", "ppt_analyze_data", "ppt_generate", "ppt_import_theme", "ppt_browse_template"
                }
                If Not jsonToolsNeedPreserve.Contains(toolName) AndAlso Not String.IsNullOrWhiteSpace(resultText) Then
                    resultText = Bridges.SkillBridge.SanitizeLeakedJsonSnippets(resultText)
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = Not result.IsError,
                    .result = resultText,
                    .isError = result.IsError,
                    .error = If(result.IsError, resultText, ""),
                    .taskCompleted = result.TaskCompleted
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 重置选区上下文（新会话开始时由前端调用，防止跨会话选区污染）
        ''' </summary>
        Public Function ResetRangeContext() As String
            ExcelMcpTools.RangeContext.Reset()
            Return JsonConvert.SerializeObject(New With {.success = True})
        End Function

        '#End Region

        '#Region "配置管理"

        ''' <summary>
        ''' 获取所有 Server 配置
        ''' </summary>
        Public Function GetServerConfigs() As String
            Try
                Dim configs = _configManager.GetServerConfigs()
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .servers = configs
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 保存 Server 配置
        ''' </summary>
        Public Async Function SaveServerConfigAsync(configJson As String) As Task(Of String)
            Return Await Task.Run(Of String)(Function()
                Try
                    Dim config = JsonConvert.DeserializeObject(Of McpClientConfig)(configJson)
                    If config Is Nothing OrElse String.IsNullOrEmpty(config.Name) Then
                        Return JsonConvert.SerializeObject(New With {
                            .success = False,
                            .error = "无效的配置"
                        })
                    End If

                    _configManager.SaveServerConfigAsync(config.Name, config)

                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .message = "配置已保存"
                    })

                Catch ex As Exception
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = ex.Message
                    })
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 获取 MCP 配置
        ''' </summary>
        Public Function GetMcpConfig() As String
            Try
                Dim config = _configManager.GetConfig()
                Dim json = JsonConvert.SerializeObject(New With {
                    .success = True,
                    .config = config
                })
                ' DEBUG: 写到桌面让前端能看到实际返回
                Try
                    Dim p = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) & "\mcp_diag.log"
                    IO.File.AppendAllText(p, "=== GetMcpConfig 返回 ===" & vbCrLf & json & vbCrLf)
                Catch
                End Try
                Return json
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        '#End Region

        '#Region "状态查询"

        ''' <summary>
        ''' 获取 MCP 整体状态（供前端查询）
        ''' </summary>
        Public Function GetStatus() As String
            Try
                Dim toolCount As Integer = 0
                If _serverHost IsNot Nothing AndAlso _serverHost.Server IsNot Nothing Then
                    toolCount = _serverHost.Server.GetToolCount()
                End If
                Dim status = New McpStatusInfo With {
                    .ServerRunning = _serverHost.IsRunning,
                    .ServerPort = _serverHost.Port,
                    .ServerToolCount = toolCount
                }

                ' 添加 Client 状态
                status.Clients = _clientPool.GetAllStatus()

                ' 计算总可用工具数：本地工具（Excel+PPT）从注册中心统计全量，另加已连接客户端工具数
                Dim registry = McpCapabilityRegistry.Instance
                Dim allTools = registry.GetAllCapabilities().
                    Where(Function(c) c.Kind = CapabilityKind.McpTool).ToList()
                status.TotalAvailableTools = allTools.Count
                For Each clientStatus In status.Clients
                    If clientStatus.Connected Then
                        status.TotalAvailableTools += clientStatus.ToolCount
                    End If
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .status = status
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取 Server 状态
        ''' </summary>
        Public Function GetServerStatus() As String
            Try
                ' 始终返回本地工具数量，即使 Server 未运行
                Dim localToolCount As Integer = 0
                If _serverHost IsNot Nothing AndAlso _serverHost.Server IsNot Nothing Then
                    localToolCount = _serverHost.Server.GetToolCount()
                End If

                If Not _serverHost.IsRunning Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .running = False,
                        .message = "MCP Server 未运行",
                        .port = _serverHost.Port,
                        .toolCount = localToolCount
                    })
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .running = True,
                    .port = _serverHost.Port,
                    .toolCount = _serverHost.Server.GetToolCount(),
                    .connectedClients = _serverHost.ConnectedClientCount
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取所有可用工具（供 AI 使用）
        ''' 本地 Excel 工具始终可用（不依赖 TCP Server 运行状态）
        ''' </summary>
        Public Function GetAllAvailableTools() As String
            Try
                ' 从统一注册中心读所有本地工具（加新工具集不用再改这里）
                ' 先确保注册完成（兜底）
                McpCapabilityRegistry.RegisterAllBuiltInToolSets()
                Dim allTools = _clientPool.GetAllAvailableTools()

                ' 本地工具：从注册中心一次读全（Excel + PPT + Computer + 未来扩展）
                Dim registry = McpCapabilityRegistry.Instance
                Dim localCaps = registry.GetAllCapabilities().
                    Where(Function(c) c.Kind = CapabilityKind.McpTool AndAlso c.Source.EndsWith("-local")).ToList()
                For Each c In localCaps
                    Dim params = New List(Of McpParameterInfo)()
                    If c.InputSchema IsNot Nothing AndAlso c.InputSchema.Properties IsNot Nothing Then
                        Dim reqSet As New HashSet(Of String)(If(c.InputSchema.Required, Enumerable.Empty(Of String)()))
                        For Each prop In c.InputSchema.Properties
                            params.Add(New McpParameterInfo With {
                                .Name = prop.Key,
                                .Type = prop.Value.Type,
                                .Description = prop.Value.Description,
                                .Required = reqSet.Contains(prop.Key)
                            })
                        Next
                    End If
                    ' 【修复·元数据丢失】与 McpBridge.GetAvailableMcpTools 保持一致，
                    ' 否则走这条降级路径时前端同样拿不到 intentKeywords。
                    allTools.Add(New McpToolInfo With {
                        .Name = c.Id,
                        .Description = c.Description,
                        .Source = c.Source,
                        .Parameters = params,
                        .Enabled = True,
                        .DisplayName = c.DisplayName,
                        .IntentKeywords = If(c.Keywords, New List(Of String)()),
                        .IsWrite = c.IsWriteOperation,
                        .IsDangerous = c.IsDangerous,
                        .DangerSeverity = c.DangerSeverity,
                        .DangerMessage = c.DangerMessage
                    })
                Next

                ' 去重：按 Source + Name 组合去重
                Dim seen = New Dictionary(Of String, Integer)()
                For i As Integer = allTools.Count - 1 To 0 Step -1
                    Dim key = allTools(i).Source & "__" & allTools(i).Name
                    If seen.ContainsKey(key) Then
                        allTools.RemoveAt(i)
                    Else
                        seen(key) = i
                    End If
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .tools = allTools,
                    .count = allTools.Count
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        '#End Region

        Public Sub Dispose() Implements IDisposable.Dispose
            If _isDisposed Then Return
            _isDisposed = True

            Try
                ' 【增强】首先取消所有后台任务
                Try
                    If _cancellationTokenSource IsNot Nothing Then
                        _cancellationTokenSource.Cancel()
                        ' 给后台 Task 最多 2 秒响应取消
                        If _bootstrapTask IsNot Nothing AndAlso Not _bootstrapTask.IsCompleted Then
                            System.Threading.Tasks.Task.WaitAny({ _bootstrapTask }, 2000)
                            System.Diagnostics.Debug.WriteLine("[McpManager] bootstrapTask 等待结束")
                        End If
                        If _autoConnectTask IsNot Nothing AndAlso Not _autoConnectTask.IsCompleted Then
                            System.Threading.Tasks.Task.WaitAny({ _autoConnectTask }, 2000)
                            System.Diagnostics.Debug.WriteLine("[McpManager] autoConnectTask 等待结束")
                        End If
                        System.Diagnostics.Debug.WriteLine("[McpManager] 已发送取消信号给后台任务，不等待")

                        _cancellationTokenSource.Dispose()
                        _cancellationTokenSource = Nothing
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[McpManager] 取消任务异常: {ex.Message}")
                End Try

                ' 【增强】停止 MCP Server（如果有运行中的 TCP 监听服务）
                Try
                    If _serverHost IsNot Nothing Then
                        _serverHost.Dispose()
                        _serverHost = Nothing
                        System.Diagnostics.Debug.WriteLine("[McpManager] McpServerHost 已释放")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[McpManager] 停止服务异常: {ex.Message}")
                    ' 确保即使异常也清理引用
                    _serverHost = Nothing
                End Try

                ' 【增强】断开所有 MCP Client 连接并释放连接池
                Try
                    If _clientPool IsNot Nothing Then
                        _clientPool.Dispose()
                        _clientPool = Nothing
                        System.Diagnostics.Debug.WriteLine("[McpManager] McpClientPool 已释放")
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"[McpManager] 释放 ClientPool 异常: {ex.Message}")
                    ' 确保即使异常也清理引用
                    _clientPool = Nothing
                End Try

                System.Diagnostics.Debug.WriteLine("[McpManager] Dispose 完成")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[McpManager] Dispose 过程中发生未预期异常: {ex.Message}")
            End Try
        End Sub

    End Class

    '#End Region

End Namespace



