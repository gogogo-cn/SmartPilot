Imports System
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core
Imports ExcelAIVSTOAddIn.Utils

Namespace Utils

    ''' <summary>
    ''' WebView2 工厂类 - 统一创建和配置 WebView2 实例
    ''' 消除 ThisAddIn.vb 中的重复代码，提供一致的初始化流程
    ''' </summary>
    Public Class WebView2Factory

        ' 单例实例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of WebView2Factory)(
            Function() New WebView2Factory(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 默认用户数据文件夹
        Private ReadOnly _defaultUserDataFolder As String

        ' 日志记录器
        Private ReadOnly _logger As ChatLogger

        ''' <summary>
        ''' 获取单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As WebView2Factory
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数
        ''' </summary>
        Private Sub New()
            _logger = ChatLogger.Instance
            _defaultUserDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelAIVSTOAddIn",
                "WebView2")
        End Sub

        ''' <summary>
        ''' WebView2 配置选项
        ''' </summary>
        Public Class WebView2Options
            ''' <summary>是否启用脚本</summary>
            Public Property EnableScript As Boolean = True
            ''' <summary>是否启用 WebMessage</summary>
            Public Property EnableWebMessage As Boolean = True
            ''' <summary>是否启用开发者工具（默认仅在 Debug 模式启用）</summary>
            Public Property EnableDevTools As Boolean = False
            ''' <summary>是否禁用 HTTP 缓存</summary>
            Public Property DisableCache As Boolean = True
            ''' <summary>虚拟主机名称</summary>
            Public Property VirtualHostName As String = "app.excelai.local"
            ''' <summary>UI 文件夹路径（可选，自动检测）</summary>
            Public Property UiFolder As String = ""
            ''' <summary>要注入的 HostObject（可选）</summary>
            Public Property HostObject As Object = Nothing
            ''' <summary>HostObject 名称</summary>
            Public Property HostObjectName As String = "excelBridge"
            ''' <summary>初始页面 URL（可选）</summary>
            Public Property InitialUrl As String = ""
            ''' <summary>是否清除缓存</summary>
            Public Property ClearCacheOnInit As Boolean = True
        End Class

        ''' <summary>
        ''' 检测 WebView2 运行时是否可用
        ''' </summary>
        Public Function CheckRuntimeAvailable() As Boolean
            Try
                ' 方法1: 检查注册表
                Dim registryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
                Dim pv = Microsoft.Win32.Registry.GetValue(registryPath, "pv", Nothing)

                If pv IsNot Nothing Then
                    Dim version = pv.ToString()
                    _logger?.LogInfo("WEBVIEW2_FACTORY", String.Format("检测到 WebView2 运行时: {0}", version))
                    Return True
                End If

                ' 方法2: 尝试创建环境
                Dim createTask = Task.Run(Async Function()
                    Await CoreWebView2Environment.CreateAsync()
                End Function)
                createTask.Wait(5000)

                If Not createTask.IsFaulted Then
                    _logger?.LogInfo("WEBVIEW2_FACTORY", "WebView2 运行时检测成功")
                    Return True
                End If

                _logger?.LogWarning("WEBVIEW2_FACTORY", "WebView2 环境创建超时或失败")
                Return False

            Catch ex As Exception
                _logger?.LogWarning("WEBVIEW2_FACTORY", "WebView2 运行时检测失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 异步创建并初始化 WebView2 实例
        ''' </summary>
        ''' <param name="options">配置选项</param>
        ''' <returns>已初始化的 WebView2 实例</returns>
        Public Async Function CreateWebView2Async(Optional options As WebView2Options = Nothing) As Task(Of WebView2)
            If options Is Nothing Then
                options = New WebView2Options()
            End If

            Try
                _logger?.LogInfo("WEBVIEW2_FACTORY", "开始创建 WebView2 实例")

                ' 创建 WebView2 控件
                Dim webView As New WebView2()
                webView.Dock = DockStyle.Fill

                ' 创建环境
                Dim envOptions As New CoreWebView2EnvironmentOptions()
                envOptions.AllowSingleSignOnUsingOSPrimaryAccount = False

                Dim userDataFolder = _defaultUserDataFolder
                If Not Directory.Exists(userDataFolder) Then
                    Directory.CreateDirectory(userDataFolder)
                End If

                Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, envOptions)

                ' 初始化 CoreWebView2
                Await webView.EnsureCoreWebView2Async(env)

                ' 配置设置
                ConfigureWebView2(webView, options)

                _logger?.LogInfo("WEBVIEW2_FACTORY", "WebView2 实例创建完成")
                Return webView

            Catch ex As Exception
                _logger?.LogError("WEBVIEW2_FACTORY", "创建 WebView2 实例失败", ex)
                Throw
            End Try
        End Function

        ''' <summary>
        ''' 配置 WebView2 实例
        ''' </summary>
        Private Sub ConfigureWebView2(webView As WebView2, options As WebView2Options)
            Try
                If webView.CoreWebView2 Is Nothing Then
                    Return
                End If

                ' 基本设置
                webView.CoreWebView2.Settings.IsScriptEnabled = options.EnableScript
                webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                webView.CoreWebView2.Settings.IsWebMessageEnabled = options.EnableWebMessage
                webView.CoreWebView2.Settings.IsStatusBarEnabled = False

                ' 开发者工具
#If DEBUG Then
                webView.CoreWebView2.Settings.AreDevToolsEnabled = True
#Else
                webView.CoreWebView2.Settings.AreDevToolsEnabled = options.EnableDevTools
#End If

                ' 禁用缓存
                If options.DisableCache Then
                    Try
                        webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                            "Network.setCacheDisabled",
                            "{""cacheDisabled"":true}"
                        ).ConfigureAwait(False)
                    Catch ex As Exception
                        _logger?.LogWarning("WEBVIEW2_FACTORY", "禁用缓存失败: " & ex.Message)
                    End Try
                End If

                ' 清除缓存
                If options.ClearCacheOnInit Then
                    Try
                        webView.CoreWebView2.Profile.ClearBrowsingDataAsync(
                            CoreWebView2BrowsingDataKinds.CacheStorage)
                        _logger?.LogInfo("WEBVIEW2_FACTORY", "HTTP缓存已清除")
                    Catch ex As Exception
                        _logger?.LogWarning("WEBVIEW2_FACTORY", "清除缓存失败(非致命): " & ex.Message)
                    End Try
                End If

                ' 设置虚拟主机映射
                If Not String.IsNullOrEmpty(options.UiFolder) AndAlso Directory.Exists(options.UiFolder) Then
                    webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        options.VirtualHostName,
                        options.UiFolder,
                        CoreWebView2HostResourceAccessKind.Allow)
                    _logger?.LogInfo("WEBVIEW2_FACTORY", String.Format("虚拟主机映射已设置: {0} -> {1}", options.VirtualHostName, options.UiFolder))
                End If

                ' 注入 HostObject
                If options.HostObject IsNot Nothing Then
                    webView.CoreWebView2.AddHostObjectToScript(options.HostObjectName, options.HostObject)
                    _logger?.LogInfo("WEBVIEW2_FACTORY", String.Format("HostObject 已注入: {0}", options.HostObjectName))
                End If

            Catch ex As Exception
                _logger?.LogError("WEBVIEW2_FACTORY", "配置 WebView2 失败", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 获取程序集目录
        ''' </summary>
        Public Function GetAssemblyDirectory() As String
            Try
                Dim codeBase As String = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
                Dim uri As New Uri(codeBase)
                Dim path As String = uri.LocalPath
                Return IO.Path.GetDirectoryName(path)
            Catch ex As Exception  ' 
                Return IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            End Try
        End Function

        ''' <summary>
        ''' 获取 UI 文件夹路径
        ''' </summary>
        Public Function GetUiFolder(Optional htmlFileName As String = "index.html") As String
            Dim assemblyDir = GetAssemblyDirectory()
            Dim uiFolder = IO.Path.Combine(assemblyDir, "UI")
            Dim htmlPath = IO.Path.Combine(uiFolder, htmlFileName)

            If File.Exists(htmlPath) Then
                Return uiFolder
            End If

            ' 尝试其他路径
            Dim possiblePaths As String() = {
                IO.Path.GetFullPath(IO.Path.Combine(assemblyDir, "..", "..", "UI")),
                IO.Path.Combine(assemblyDir, "..", "UI")
            }

            For Each path In possiblePaths
                If Directory.Exists(path) AndAlso File.Exists(IO.Path.Combine(path, htmlFileName)) Then
                    Return path
                End If
            Next

            Return uiFolder
        End Function

        ''' <summary>
        ''' 构建带缓存破坏参数的 URL
        ''' </summary>
        Public Function BuildUrlWithCacheBuster(baseUrl As String) As String
            Dim cacheBuster = Guid.NewGuid().ToString("N")
            If baseUrl.Contains("?") Then
                Return baseUrl & "&_cb=" & cacheBuster
            Else
                Return baseUrl & "?_cb=" & cacheBuster
            End If
        End Function

        ''' <summary>
        ''' 创建任务窗格容器控件
        ''' </summary>
        Public Function CreateTaskPaneContainer(width As Integer, height As Integer) As UserControl
            Dim control As New UserControl()
            control.Width = width
            control.Height = height
            control.BackColor = System.Drawing.Color.White
            Return control
        End Function

        ''' <summary>
        ''' 显示加载状态面板
        ''' </summary>
        Public Sub ShowLoadingPanel(container As UserControl, message As String, Optional statusText As String = "初始化组件...")
            If container Is Nothing OrElse container.IsDisposed Then
                Return
            End If

            If container.InvokeRequired Then
                container.Invoke(New Action(Of UserControl, String, String)(AddressOf ShowLoadingPanel), container, message, statusText)
                Return
            End If

            container.Controls.Clear()

            Dim loadingPanel As New Panel()
            loadingPanel.Dock = DockStyle.Fill
            loadingPanel.BackColor = System.Drawing.Color.White

            Dim loadingLabel As New Label()
            loadingLabel.Text = message
            loadingLabel.Font = New System.Drawing.Font("Microsoft YaHei", 12, System.Drawing.FontStyle.Regular)
            loadingLabel.AutoSize = True
            loadingLabel.Location = New System.Drawing.Point(80, 200)

            Dim progressBar As New ProgressBar()
            progressBar.Style = ProgressBarStyle.Marquee
            progressBar.MarqueeAnimationSpeed = 30
            progressBar.Width = 260
            progressBar.Height = 20
            progressBar.Location = New System.Drawing.Point(80, 240)

            Dim statusLabel As New Label()
            statusLabel.Text = statusText
            statusLabel.Font = New System.Drawing.Font("Microsoft YaHei", 9, System.Drawing.FontStyle.Regular)
            statusLabel.ForeColor = System.Drawing.Color.Gray
            statusLabel.AutoSize = True
            statusLabel.Location = New System.Drawing.Point(80, 270)

            loadingPanel.Controls.Add(loadingLabel)
            loadingPanel.Controls.Add(progressBar)
            loadingPanel.Controls.Add(statusLabel)

            container.Controls.Add(loadingPanel)
        End Sub

    End Class

End Namespace