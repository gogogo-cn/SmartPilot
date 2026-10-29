Imports System
Imports System.IO
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class WebView2Manager
        Implements IDisposable

        Private _webView As WebView2
        Private _hostControl As UserControl
        Private _logger As Utils.ChatLogger
        Private _isReady As Boolean = False
        Private _loadingPanel As Panel
        Private _statusLabel As Label
        Private _excelBridge As ExcelAIVSTOAddIn.ExcelBridge

        Public Event WebMessageReceived As EventHandler(Of CoreWebView2WebMessageReceivedEventArgs)
        Public Event NavigationCompleted As EventHandler(Of CoreWebView2NavigationCompletedEventArgs)

        Public ReadOnly Property IsReady As Boolean
            Get
                Return _isReady AndAlso _webView IsNot Nothing AndAlso _webView.CoreWebView2 IsNot Nothing
            End Get
        End Property

        Public ReadOnly Property WebView As WebView2
            Get
                Return _webView
            End Get
        End Property

        Public Sub New(hostControl As UserControl, logger As Utils.ChatLogger)
            _hostControl = hostControl
            _logger = logger
        End Sub

        ''' <summary>
        ''' 检测 WebView2 运行时是否可用
        ''' </summary>
        Public Function CheckRuntime() As Boolean
            Try
                Dim registryPath = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"
                Dim pv = Microsoft.Win32.Registry.GetValue(registryPath, "pv", Nothing)

                If pv IsNot Nothing Then
                    _logger?.LogInfo("WEBVIEW2", $"检测到 WebView2 运行时: {pv.ToString()}")
                    Return True
                End If

                Dim createTask = Task.Run(Async Function()
                    Await CoreWebView2Environment.CreateAsync()
                End Function)
                createTask.Wait(5000)
                If Not createTask.IsFaulted Then
                    _logger?.LogInfo("WEBVIEW2", "WebView2 运行时检测成功(通过环境创建)")
                    Return True
                End If
                _logger?.LogWarning("WEBVIEW2", "WebView2 环境创建超时或失败: " & If(createTask.Exception?.InnerException?.Message, "未知错误"))
                Return False

            Catch ex As Exception
                _logger?.LogWarning("WEBVIEW2", "未检测到 WebView2 运行时: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 显示加载界面
        ''' </summary>
        Public Sub ShowLoadingPanel()
            If _hostControl Is Nothing OrElse _hostControl.IsDisposed Then Return

            _hostControl.Controls.Clear()

            _loadingPanel = New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White
            }

            Dim loadingLabel = New Label() With {
                .Text = "正在加载 SmartPilot 智航...",
                .Font = New Font("Microsoft YaHei", 12),
                .AutoSize = True,
                .Location = New Point(80, 200)
            }

            Dim progressBar = New ProgressBar() With {
                .Style = ProgressBarStyle.Marquee,
                .MarqueeAnimationSpeed = 30,
                .Width = 260,
                .Height = 20,
                .Location = New Point(80, 240)
            }

            _statusLabel = New Label() With {
                .Text = "初始化组件...",
                .Font = New Font("Microsoft YaHei", 9),
                .ForeColor = Color.Gray,
                .AutoSize = True,
                .Location = New Point(80, 270)
            }

            _loadingPanel.Controls.Add(loadingLabel)
            _loadingPanel.Controls.Add(progressBar)
            _loadingPanel.Controls.Add(_statusLabel)
            _hostControl.Controls.Add(_loadingPanel)
        End Sub

        ''' <summary>
        ''' 更新加载状态文本
        ''' </summary>
        Public Sub UpdateLoadingStatus(status As String)
            If _statusLabel IsNot Nothing AndAlso Not _statusLabel.IsDisposed Then
                If _statusLabel.InvokeRequired Then
                    _statusLabel.Invoke(New Action(Of String)(AddressOf UpdateLoadingStatus), status)
                Else
                    _statusLabel.Text = status
                End If
            End If
        End Sub

        ''' <summary>
        ''' 显示 WebView2 安装引导
        ''' </summary>
        Public Sub ShowInstallationGuide()
            Try
                UpdateLoadingStatus("WebView2 运行时未安装")

                If _hostControl Is Nothing OrElse _hostControl.IsDisposed Then Return

                _hostControl.Controls.Clear()

                Dim guidePanel = New Panel() With {
                    .Dock = DockStyle.Fill,
                    .BackColor = Color.White,
                    .Padding = New Padding(30)
                }

                Dim titleLabel = New Label() With {
                    .Text = "需要安装 WebView2 运行时",
                    .Font = New Font("Microsoft YaHei", 16, FontStyle.Bold),
                    .ForeColor = Color.FromArgb(&H333333),
                    .AutoSize = True,
                    .Location = New Point(30, 40)
                }
                guidePanel.Controls.Add(titleLabel)

                Dim descLabel = New Label() With {
                    .Text = "SmartPilot 智航需要 Microsoft Edge WebView2 运行时才能运行。" & vbCrLf & vbCrLf & "请按照以下步骤安装：",
                    .Font = New Font("Microsoft YaHei", 11),
                    .ForeColor = Color.FromArgb(&H666666),
                    .AutoSize = False,
                    .Size = New Size(360, 80),
                    .Location = New Point(30, 90)
                }
                guidePanel.Controls.Add(descLabel)

                Dim step1Label = New Label() With {
                    .Text = "步骤 1: 下载 WebView2 运行时",
                    .Font = New Font("Microsoft YaHei", 10, FontStyle.Bold),
                    .ForeColor = Color.FromArgb(&H333333),
                    .AutoSize = True,
                    .Location = New Point(30, 190)
                }
                guidePanel.Controls.Add(step1Label)

                Dim downloadBtn = New Button() With {
                    .Text = "前往微软官网下载",
                    .Font = New Font("Microsoft YaHei", 10),
                    .Width = 180,
                    .Height = 35,
                    .Location = New Point(30, 220),
                    .BackColor = Color.FromArgb(&H4A90D9),
                    .ForeColor = Color.White,
                    .FlatStyle = FlatStyle.Flat
                }
                AddHandler downloadBtn.Click, Sub(s, e)
                    System.Diagnostics.Process.Start("https://go.microsoft.com/fwlink/p/?LinkId=2124703")
                End Sub
                guidePanel.Controls.Add(downloadBtn)

                Dim step2Desc = New Label() With {
                    .Text = "步骤 2: 安装完成后重启 Excel" & vbCrLf & "安装完成后，请关闭所有 Excel 窗口，" & vbCrLf & "然后重新打开 Excel 即可使用。",
                    .Font = New Font("Microsoft YaHei", 9),
                    .ForeColor = Color.FromArgb(&H999999),
                    .AutoSize = False,
                    .Size = New Size(360, 60),
                    .Location = New Point(30, 310)
                }
                guidePanel.Controls.Add(step2Desc)

                Dim line = New Label() With {
                    .AutoSize = False,
                    .Size = New Size(360, 1),
                    .BackColor = Color.FromArgb(&HEEEEEE),
                    .Location = New Point(30, 390)
                }
                guidePanel.Controls.Add(line)

                Dim tipLabel = New Label() With {
                    .Text = "提示: WebView2 是微软官方组件，安全可信。" & vbCrLf & "大多数 Windows 10/11 系统已预装。",
                    .Font = New Font("Microsoft YaHei", 9),
                    .ForeColor = Color.FromArgb(&H999999),
                    .AutoSize = False,
                    .Size = New Size(360, 50),
                    .Location = New Point(30, 410)
                }
                guidePanel.Controls.Add(tipLabel)

                _hostControl.Controls.Add(guidePanel)
                _logger?.LogInfo("WEBVIEW2", "已显示 WebView2 安装引导界面")

            Catch ex As Exception
                _logger?.LogError("WEBVIEW2", "显示安装引导失败", ex)
                MessageBox.Show("初始化失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        ''' <summary>
        ''' 异步初始化 WebView2
        ''' </summary>
        Public Async Function InitializeAsync(uiPath As String, excelBridge As ExcelAIVSTOAddIn.ExcelBridge) As Task
            Try
                _excelBridge = excelBridge

                UpdateLoadingStatus("正在检测 WebView2 运行时...")
                If Not CheckRuntime() Then
                    ShowInstallationGuide()
                    Return
                End If

                UpdateLoadingStatus("正在初始化 WebView2 环境...")
                _logger?.LogInfo("WEBVIEW2", "开始异步初始化")

                Dim userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "ExcelAIVSTOAddIn",
                    "WebView2"
                )

                Dim options = New CoreWebView2EnvironmentOptions()
                options.AllowSingleSignOnUsingOSPrimaryAccount = False

                Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder, options)

                UpdateLoadingStatus("正在创建 WebView2 实例...")

                _webView = New WebView2() With {.Dock = DockStyle.Fill}

                AddHandler _webView.CoreWebView2InitializationCompleted,
                    Sub(s, args)
                        If args.IsSuccess Then
                            OnWebView2Initialized(uiPath)
                        Else
                            ShowError(args.InitializationException)
                        End If
                    End Sub

                Await _webView.EnsureCoreWebView2Async(env)

            Catch ex As Exception
                _logger?.LogError("WEBVIEW2", "异步初始化失败", ex)
                ShowError(ex)
            End Try
        End Function

        ''' <summary>
        ''' WebView2 初始化完成后的配置
        ''' </summary>
        Private Sub OnWebView2Initialized(uiPath As String)
            Try
                UpdateLoadingStatus("正在配置 WebView2...")
                _logger?.LogInfo("WEBVIEW2", "初始化完成")

                Try
                    _webView.CoreWebView2.Profile.ClearBrowsingDataAsync(CoreWebView2BrowsingDataKinds.CacheStorage)
                    _logger?.LogInfo("WEBVIEW2", "HTTP缓存已清除")
                Catch cacheEx As Exception
                    _logger?.LogWarning("WEBVIEW2", "清除缓存失败(非致命): " & cacheEx.Message)
                End Try

                _webView.CoreWebView2.Settings.IsScriptEnabled = True
                _webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                _webView.CoreWebView2.Settings.IsWebMessageEnabled = True
                _webView.CoreWebView2.Settings.IsStatusBarEnabled = False
                _webView.CoreWebView2.Settings.AreDevToolsEnabled = True

                _webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Network.setCacheDisabled",
                    "{""cacheDisabled"":true}"
                ).ConfigureAwait(False)

                Dim uiFolder = Path.GetDirectoryName(uiPath)
                _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    "app.excelai.local",
                    uiFolder,
                    CoreWebView2HostResourceAccessKind.Allow)

                ' 临时目录也映射为虚拟主机：截屏等生成在 Temp 下的文件可被同源 https 引用显示
                Try
                    Dim tempDir = System.IO.Path.GetTempPath()
                    If System.IO.Directory.Exists(tempDir) Then
                        _webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "screenassets.local",
                            tempDir,
                            CoreWebView2HostResourceAccessKind.Allow)
                    End If
                Catch ex As Exception
                    _logger?.LogWarning("WEBVIEW2", "临时目录虚拟主机映射失败(非致命): " & ex.Message)
                End Try

                If _excelBridge IsNot Nothing Then
                    _webView.CoreWebView2.AddHostObjectToScript("excelBridge", _excelBridge)
                End If

                ' 注册任务规划 Bridge
                Try
                    _webView.CoreWebView2.AddHostObjectToScript("taskScheduler",
                        Bridges.TaskSchedulerBridge.Instance)
                    _logger?.LogInfo("WEBVIEW2", "TaskSchedulerBridge registered as host object")
                Catch ex As Exception
                    _logger?.LogError("WEBVIEW2", "Failed to register TaskSchedulerBridge", ex)
                End Try

                AddHandler _webView.CoreWebView2.WebMessageReceived,
                    Sub(s, e) RaiseEvent WebMessageReceived(Me, e)

                AddHandler _webView.CoreWebView2.NavigationCompleted,
                    Sub(s, e)
                        RaiseEvent NavigationCompleted(Me, e)
                        If e.IsSuccess Then
                            _isReady = True
                            _logger?.LogInfo("WEBVIEW2", "页面导航完成")
                            FocusWebView()
                        End If
                    End Sub

                _hostControl.Controls.Clear()
                _hostControl.Controls.Add(_webView)

                AddHandler _webView.PreviewKeyDown,
                    Sub(s, e)
                        If e.KeyCode = Keys.F12 Then
                            Try : _webView.CoreWebView2.OpenDevToolsWindow() : Catch : End Try
                        End If
                    End Sub

                Dim cacheBuster = Guid.NewGuid().ToString("N")
                _webView.Source = New Uri($"https://app.excelai.local/index.html?_cb={cacheBuster}")

                _logger?.LogInfo("WEBVIEW2", "设置完成")

            Catch ex As Exception
                ShowError(ex)
            End Try
        End Sub

        ''' <summary>
        ''' 获取 WebView2 焦点
        ''' </summary>
        Public Sub FocusWebView()
            Try
                If _webView IsNot Nothing AndAlso Not _webView.IsDisposed Then
                    _webView.Focus()
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        ''' <summary>
        ''' 显示初始化错误
        ''' </summary>
        Public Sub ShowError(ex As Exception)
            _logger?.LogError("WEBVIEW2", "初始化错误", ex)

            If _hostControl Is Nothing OrElse _hostControl.IsDisposed Then Return

            If _hostControl.InvokeRequired Then
                _hostControl.Invoke(New Action(Of Exception)(AddressOf ShowError), ex)
                Return
            End If

            _hostControl.Controls.Clear()

            Dim errorPanel = New Panel() With {
                .Dock = DockStyle.Fill,
                .BackColor = Color.White
            }

            Dim errorLabel = New Label() With {
                .Text = "加载失败",
                .Font = New Font("Microsoft YaHei", 14, FontStyle.Bold),
                .ForeColor = Color.Red,
                .AutoSize = True,
                .Location = New Point(140, 150)
            }

            Dim messageLabel = New Label() With {
                .Text = If(ex IsNot Nothing, ex.Message, "未知错误"),
                .Font = New Font("Microsoft YaHei", 9),
                .ForeColor = Color.Gray,
                .AutoSize = True,
                .Location = New Point(50, 200),
                .MaximumSize = New Size(320, 0)
            }

            Dim retryButton = New Button() With {
                .Text = "重试",
                .Width = 100,
                .Height = 30,
                .Location = New Point(160, 280)
            }
            AddHandler retryButton.Click,
                Sub(s, e)
                    ShowLoadingPanel()
                    InitializeAsync(GetUIPath(), _excelBridge).ContinueWith(
                        Sub(t)
                            If t.IsFaulted Then ShowError(t.Exception?.InnerException)
                        End Sub, TaskScheduler.FromCurrentSynchronizationContext())
                End Sub

            errorPanel.Controls.Add(errorLabel)
            errorPanel.Controls.Add(messageLabel)
            errorPanel.Controls.Add(retryButton)
            _hostControl.Controls.Add(errorPanel)
        End Sub

        ''' <summary>
        ''' 获取 UI 文件路径
        ''' </summary>
        Public Shared Function GetUIPath() As String
            Dim assemblyDir = GetAssemblyDirectory()
            Dim uiFolder = Path.Combine(assemblyDir, "UI")
            Dim uiPath = Path.Combine(uiFolder, "index.html")

            If File.Exists(uiPath) Then Return uiPath

            Dim possibleSourcePaths = GetPossibleUIPaths(assemblyDir)
            For Each sourcePath In possibleSourcePaths
                If File.Exists(sourcePath) Then
                    Try
                        If Not Directory.Exists(uiFolder) Then Directory.CreateDirectory(uiFolder)
                        File.Copy(sourcePath, uiPath, True)
                        Return uiPath
                    Catch ex As Exception  ' 
                    End Try
                End If
            Next

            Throw New FileNotFoundException("UI 文件未找到", uiPath)
        End Function

        Private Shared Function GetAssemblyDirectory() As String
            Dim codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
            Dim uri = New Uri(codeBase)
            Return Path.GetDirectoryName(uri.LocalPath)
        End Function

        Private Shared Function GetPossibleUIPaths(assemblyDir As String) As String()
            Return New String() {
                Path.Combine(assemblyDir, "..", "UI", "index.html"),
                Path.Combine(assemblyDir, "..", "..", "UI", "index.html"),
                Path.Combine(Environment.CurrentDirectory, "UI", "index.html"),
                Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "UI", "index.html")
            }
        End Function

        ''' <summary>
        ''' 停止 WebView2 渲染
        ''' </summary>
        Public Sub StopWebView()
            Try
                If _webView IsNot Nothing AndAlso _webView.CoreWebView2 IsNot Nothing Then
                    _webView.CoreWebView2.Stop()
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

        ''' <summary>
        ''' 释放资源（WebView2 异步释放，避免阻塞 Excel 关闭）
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                StopWebView()
                If _webView IsNot Nothing Then
                    RemoveHandler _webView.PreviewKeyDown, Nothing
                    Dim wv = _webView
                    _webView = Nothing
                    Task.Run(Sub()
                                 Try : wv.Dispose() : Catch : End Try
                             End Sub)
                End If
            Catch ex As Exception  ' 
            End Try
        End Sub

    End Class