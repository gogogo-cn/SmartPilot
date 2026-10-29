Option Strict On
Option Explicit On
Option Infer On
Imports System
Imports System.IO
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Namespace Services.Mcp

    '#Region "MCP 配置管理器"

    ''' <summary>
    ''' MCP 配置管理器 - 加载、保存、管理 MCP 配置
    ''' </summary>
    Public Class McpConfigManager

        Private _config As McpGlobalConfig
        Private _configPath As String
        Private _userConfigPath As String
        Private _lockObj As New Object()

        ''' <summary>
        ''' 配置文件路径（默认模板）
        ''' </summary>
        Public ReadOnly Property ConfigPath As String
            Get
                Return _configPath
            End Get
        End Property

        ''' <summary>
        ''' 用户配置路径（运行时修改，不被编译覆盖）
        ''' </summary>
        Public ReadOnly Property UserConfigPath As String
            Get
                Return _userConfigPath
            End Get
        End Property

        Public Sub New()
            _configPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "mcp.json")
            _userConfigPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "mcp.user.json")

            LoadConfig()
        End Sub

        Public Sub New(configPath As String)
            _configPath = configPath
            _userConfigPath = Path.Combine(Path.GetDirectoryName(configPath), "mcp.user.json")
            LoadConfig()
        End Sub

        '#Region "加载/保存"

        ''' <summary>
        ''' 加载配置：先加载默认模板 mcp.json，再合并用户配置 mcp.user.json
        ''' 用户配置优先级更高，编译不会覆盖用户配置
        ''' </summary>
        Private Sub LoadConfig()
            SyncLock _lockObj
                Try
                    Dim deskPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    Dim diagPath = IO.Path.Combine(deskPath, "mcp_diag.log")
                    Using sw As New IO.StreamWriter(diagPath, False)
                        sw.WriteLine($"=== MCP Config Load === {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
                        sw.WriteLine($"_configPath: {_configPath}")
                        sw.WriteLine($"_userConfigPath: {_userConfigPath}")

                        ' 1. 加载默认模板
                        _config = Nothing
                        If File.Exists(_configPath) Then
                            Dim json = File.ReadAllText(_configPath)
                            sw.WriteLine($"Default config length: {json.Length}")
                            _config = JsonConvert.DeserializeObject(Of McpGlobalConfig)(json)
                        End If

                        If _config Is Nothing Then
                            sw.WriteLine("No default config, creating new")
                            _config = New McpGlobalConfig()
                        End If

                        If _config.McpServers Is Nothing Then
                            _config.McpServers = New Dictionary(Of String, McpClientConfig)()
                        End If

                        sw.WriteLine($"Default McpServers count: {_config.McpServers.Count}")

                        ' 2. 合并用户配置（优先级更高）
                        If File.Exists(_userConfigPath) Then
                            sw.WriteLine($"Loading user config from: {_userConfigPath}")
                            Dim userJson = File.ReadAllText(_userConfigPath)
                            Dim userConfig = JsonConvert.DeserializeObject(Of McpGlobalConfig)(userJson)
                            If userConfig IsNot Nothing Then
                                ' 合并顶层字段（user.json 优先级更高）
                                sw.WriteLine($"  userConfig.AutoStartClients raw = {userConfig.AutoStartClients}")
                                If userConfig.AutoStartClients Then
                                    _config.AutoStartClients = True
                                    sw.WriteLine("  Merged: autoStartClients=True")
                                End If
                                ' 合并 McpServers
                                If userConfig.McpServers IsNot Nothing Then
                                    sw.WriteLine($"User McpServers count: {userConfig.McpServers.Count}")
                                    For Each kvp In userConfig.McpServers
                                        _config.McpServers(kvp.Key) = kvp.Value
                                        sw.WriteLine($"  Merged: {kvp.Key}")
                                    Next
                                End If
                            End If
                        Else
                            sw.WriteLine("No user config file found")
                        End If

                        ' 3. 展开配置中的环境变量占位符 (e.g. %USERPROFILE%, %PROGRAMFILES%)
                        ExpandEnvironmentVariables(_config, sw)

                        sw.WriteLine($"Final McpServers count: {_config.McpServers.Count}")
                        For Each serverKey In _config.McpServers.Keys
                            sw.WriteLine($"  Server: {serverKey}, enabled={_config.McpServers(serverKey).Enabled}")
                        Next
                        sw.Flush()
                    End Using

                Catch ex As Exception
                    Try
                        IO.File.WriteAllText(IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "mcp_diag_error.log"), $"LoadConfig异常: {ex.Message}{vbCrLf}堆栈: {ex.StackTrace}")
                    Catch
                    End Try
                    System.Diagnostics.Debug.WriteLine($"加载 MCP 配置失败: {ex.Message}")
                    _config = New McpGlobalConfig()
                    AddDefaultServers()
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' 添加默认的示例 Server 配置（仅作为文件不存在时的兜底）
        ''' </summary>
        Private Sub AddDefaultServers()
            Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
            Dim programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            Dim npxPath = Path.Combine(programFiles, "nodejs", "npx.cmd")
            If Not File.Exists(npxPath) Then npxPath = "npx"

            ' filesystem server - 需要安装 Node.js
            _config.McpServers("filesystem") = New McpClientConfig With {
                .Name = "filesystem",
                .Description = "文件系统访问工具（需要 npx 可用）",
                .Transport = "stdio",
                .Command = npxPath,
                .Args = New List(Of String) From {"-y", "@modelcontextprotocol/server-filesystem", userProfile},
                .Enabled = True
            }
        End Sub

        ''' <summary>
        ''' 展开配置字符串中的环境变量占位符：
        '''   %USERPROFILE%, %PROGRAMFILES%, %APPDATA%, %LOCALAPPDATA%, %DESKTOP%, %DOCUMENTS%
        ''' </summary>
        Private Shared Sub ExpandEnvironmentVariables(config As McpGlobalConfig, diagWriter As IO.StreamWriter)
            If config Is Nothing OrElse config.McpServers Is Nothing Then Return

            For Each kvp In config.McpServers
                Dim server = kvp.Value
                If server Is Nothing Then Continue For

                ' 展开 Command
                If Not String.IsNullOrEmpty(server.Command) Then
                    Dim originalCmd = server.Command
                    server.Command = ExpandSpecialFolders(server.Command)
                    If originalCmd <> server.Command Then
                        diagWriter.WriteLine($"  [{kvp.Key}] Command expanded: '{originalCmd}' -> '{server.Command}'")
                    End If
                End If

                ' 展开 Args
                If server.Args IsNot Nothing Then
                    For i = 0 To server.Args.Count - 1
                        If Not String.IsNullOrEmpty(server.Args(i)) Then
                            Dim originalArg = server.Args(i)
                            server.Args(i) = ExpandSpecialFolders(server.Args(i))
                            If originalArg <> server.Args(i) Then
                                diagWriter.WriteLine($"  [{kvp.Key}] Args[{i}] expanded: '{originalArg}' -> '{server.Args(i)}'")
                            End If
                        End If
                    Next
                End If
            Next
        End Sub

        ''' <summary>
        ''' 展开自定义环境变量 + 系统 Environment.ExpandEnvironmentVariables
        ''' 支持：%DESKTOP%, %DOCUMENTS%, %USERPROFILE%, %APPDATA%, %LOCALAPPDATA%,
        '''       %PROGRAMFILES%, %COMMONPROGRAMFILES%, %TMP%, %TEMP%
        ''' Public 可见性以便单元测试。
        ''' </summary>
        Public Shared Function ExpandSpecialFolders(path As String) As String
            If String.IsNullOrEmpty(path) Then Return path

            ' 1. 先处理我们自定义的 SpecialFolder 别名
            Dim mappings = New Dictionary(Of String, Environment.SpecialFolder) From {
                {"%DESKTOP%", Environment.SpecialFolder.DesktopDirectory},
                {"%DOCUMENTS%", Environment.SpecialFolder.MyDocuments},
                {"%MUSIC%", Environment.SpecialFolder.MyMusic},
                {"%PICTURES%", Environment.SpecialFolder.MyPictures},
                {"%VIDEOS%", Environment.SpecialFolder.MyVideos}
            }

            For Each map In mappings
                If path.IndexOf(map.Key, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Dim folder = Environment.GetFolderPath(map.Value)
                    If Not String.IsNullOrEmpty(folder) Then
                        path = Replace(path, map.Key, folder, Compare:=Microsoft.VisualBasic.CompareMethod.Text)
                    End If
                End If
            Next

            ' 2. 处理 %USERPROFILE% 即使在系统没设置时也能解析
            If path.IndexOf("%USERPROFILE%", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
                If Not String.IsNullOrEmpty(userProfile) Then
                    path = Replace(path, "%USERPROFILE%", userProfile, Compare:=Microsoft.VisualBasic.CompareMethod.Text)
                End If
            End If

            ' 3. 系统级环境变量展开（%PROGRAMFILES%, %APPDATA%, %LOCALAPPDATA% 等）
            path = Environment.ExpandEnvironmentVariables(path)

            ' 4. 将 / 统一为 Windows \
            If path.Contains("/") AndAlso Not path.StartsWith("http") Then
                path = path.Replace("/", "\")
            End If

            Return path
        End Function

        ''' <summary>
        ''' 保存配置到用户配置文件 mcp.user.json（不被编译覆盖）
        ''' </summary>
        Public Function SaveConfig() As Boolean
            SyncLock _lockObj
                Try
                    Dim dir = Path.GetDirectoryName(_userConfigPath)
                    If Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If

                    Dim json = JsonConvert.SerializeObject(_config, Formatting.Indented)
                    File.WriteAllText(_userConfigPath, json)
                    Return True

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"保存 MCP 配置失败: {ex.Message}")
                    Return False
                End Try
            End SyncLock
        End Function

        ''' <summary>
        ''' 从磁盘重新加载配置（用于配置文件被外部修改后刷新）
        ''' </summary>
        Public Sub ReloadConfig()
            LoadConfig()
        End Sub

        '#End Region

        '#Region "配置访问"

        ''' <summary>
        ''' 获取当前配置
        ''' </summary>
        Public Function GetConfig() As McpGlobalConfig
            SyncLock _lockObj
                Return Cloned(_config)
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有 MCP Server 配置
        ''' </summary>
        Public Function GetServerConfigs() As Dictionary(Of String, McpClientConfig)
            SyncLock _lockObj
                Dim result = New Dictionary(Of String, McpClientConfig)()
                For Each kvp In _config.McpServers
                    result(kvp.Key) = CloneServerConfig(kvp.Value)
                Next
                Return result
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取指定 Server 的配置
        ''' </summary>
        Public Function GetServerConfig(name As String) As McpClientConfig
            SyncLock _lockObj
                If _config.McpServers.ContainsKey(name) Then
                    Return CloneServerConfig(_config.McpServers(name))
                End If
                Return Nothing
            End SyncLock
        End Function

        ''' <summary>
        ''' 添加或更新 Server 配置
        ''' </summary>
        Public Function SaveServerConfigAsync(name As String, config As McpClientConfig) As Boolean
            SyncLock _lockObj
                config.Name = name
                _config.McpServers(name) = config
            End SyncLock
            Return SaveConfig()
        End Function

        ''' <summary>
        ''' 删除 Server 配置
        ''' </summary>
        Public Function RemoveServerConfig(name As String) As Boolean
            SyncLock _lockObj
                If _config.McpServers.ContainsKey(name) Then
                    _config.McpServers.Remove(name)
                    Return SaveConfig()
                End If
                Return False
            End SyncLock
        End Function

        ''' <summary>
        ''' 更新 Server 的启用状态
        ''' </summary>
        Public Function SetServerEnabled(name As String, enabled As Boolean) As Boolean
            SyncLock _lockObj
                If _config.McpServers.ContainsKey(name) Then
                    _config.McpServers(name).Enabled = enabled
                    Return SaveConfig()
                End If
                Return False
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取 Server 端口号
        ''' </summary>
        Public Function GetServerPort() As Integer
            Return _config.ServerPort
        End Function

        ''' <summary>
        ''' 设置 Server 端口号
        ''' </summary>
        Public Function SetServerPort(port As Integer) As Boolean
            _config.ServerPort = port
            Return SaveConfig()
        End Function

        ''' <summary>
        ''' 获取 Server 是否启用
        ''' </summary>
        Public Function IsServerEnabled() As Boolean
            Return _config.ServerEnabled
        End Function

        ''' <summary>
        ''' 设置 Server 是否启用
        ''' </summary>
        Public Function SetServerEnabled(enabled As Boolean) As Boolean
            _config.ServerEnabled = enabled
            Return SaveConfig()
        End Function

        '#End Region

        '#Region "辅助方法"

        Private Function Cloned(config As McpGlobalConfig) As McpGlobalConfig
            Dim json = JsonConvert.SerializeObject(config)
            Return JsonConvert.DeserializeObject(Of McpGlobalConfig)(json)
        End Function

        Private Function CloneServerConfig(config As McpClientConfig) As McpClientConfig
            Dim json = JsonConvert.SerializeObject(config)
            Return JsonConvert.DeserializeObject(Of McpClientConfig)(json)
        End Function

        '#End Region

    End Class

    '#End Region

End Namespace

