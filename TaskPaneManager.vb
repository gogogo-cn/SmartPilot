Imports System
Imports System.Windows.Forms
Imports System.Threading
Imports System.Threading.Tasks
Imports Microsoft.Office.Tools
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class TaskPaneManager
        Implements IDisposable

        Private _taskPanes As Dictionary(Of String, TaskPaneInfo) = New Dictionary(Of String, TaskPaneInfo)()
        Private _creatingKeys As New HashSet(Of String)()
        Private ReadOnly _createLock As New Object()
        Private _addIn As ThisAddIn
        Private _logger As Utils.ChatLogger
        Private _excelBridge As ExcelAIVSTOAddIn.ExcelBridge

        ''' <summary>P3：火与忘记异步的安全包装（捕获异常+记录日志+消除BC42358警告）</summary>
        Private Sub SafeFork(task As Task, Optional caller As String = NameOf(SafeFork))
            If task Is Nothing Then Return
            Dim __discard = task.ContinueWith(AddressOf SafeForkContinuation, Tuple.Create(_logger, caller), TaskScheduler.Default)
        End Sub

        ''' <summary>P3：SafeFork的ContinueWith辅助（单独方法避免多行Sub Lambda解析问题）</summary>
        Private Shared Sub SafeForkContinuation(antecedent As Task, state As Object)
            Dim tuple = TryCast(state, Tuple(Of Utils.ChatLogger, String))
            If tuple Is Nothing OrElse Not antecedent.IsFaulted OrElse antecedent.Exception Is Nothing Then Return
            Try
                Dim ex = antecedent.Exception.InnerException
                Dim logger = tuple.Item1
                Dim caller = tuple.Item2
                Dim msg As String = If(ex IsNot Nothing, ex.Message, "null")
                logger?.LogError("TASKPANE", caller & " 异步后台异常: " & msg, ex)
                System.Diagnostics.Debug.WriteLine("[TaskPaneManager] " & caller & " FAULT: " & If(ex IsNot Nothing, ex.ToString(), "null"))
            Catch ex As Exception  ' 
            End Try
        End Sub

        Private Const MAIN_PANE_KEY As String = "Main"
        Private Const OCR_PANE_KEY As String = "OCR"
        Private Const SETTINGS_PANE_KEY As String = "Settings"
        Private Const MOBILE_PANE_KEY As String = "Mobile"
        Private Const AI_SETTINGS_PANE_KEY As String = "AiSettings"
        Private Const OCR_SETTINGS_PANE_KEY As String = "OcrSettings"
        Private Const PPT_THEME_PANE_KEY As String = "PptTheme"

        Private Const DEFAULT_PANE_WIDTH As Integer = 480
        Private Const PPT_THEME_PANE_WIDTH As Integer = 960
        Private Const DEFAULT_PANE_HEIGHT As Integer = 600
        Private Const VIRTUAL_HOST_NAME As String = "app.excelai.local"
        Private Const SCREEN_ASSETS_HOST As String = "screenassets.local"
        Private Const UI_VERSION As String = "2026072514"

        Public ReadOnly Property MainTaskPane As CustomTaskPane
            Get
                Return GetTaskPane(MAIN_PANE_KEY)
            End Get
        End Property

        Public ReadOnly Property OcrTaskPane As CustomTaskPane
            Get
                Return GetTaskPane(OCR_PANE_KEY)
            End Get
        End Property

        ''' <summary>
        ''' 获取 OCR 面板的 WebView2 实例（供截图服务使用）
        ''' </summary>
        Public ReadOnly Property OcrWebView As WebView2
            Get
                If _taskPanes.ContainsKey(OCR_PANE_KEY) Then
                    Return _taskPanes(OCR_PANE_KEY).WebView
                End If
                Return Nothing
            End Get
        End Property

        Public ReadOnly Property MobileTaskPane As CustomTaskPane
            Get
                Return GetTaskPane(MOBILE_PANE_KEY)
            End Get
        End Property

        Public ReadOnly Property AiSettingsTaskPane As CustomTaskPane
            Get
                Return GetTaskPane(AI_SETTINGS_PANE_KEY)
            End Get
        End Property

        Public ReadOnly Property OcrSettingsTaskPane As CustomTaskPane
            Get
                Return GetTaskPane(OCR_SETTINGS_PANE_KEY)
            End Get
        End Property

        Public Sub New(addIn As ThisAddIn, logger As Utils.ChatLogger, excelBridge As ExcelAIVSTOAddIn.ExcelBridge)
            _addIn = addIn
            _logger = logger
            _excelBridge = excelBridge
            Current = Me
        End Sub

        ''' <summary>
        ''' 当前活动实例。供 Bridge/Services 层主动向任务窗格推送消息（例如 Skill 自动导入解析脚本后刷新 OCR 面板）。
        ''' </summary>
        Public Shared Property Current As TaskPaneManager

        ''' <summary>
        ''' 通知 OCR 智能解析面板：脚本库发生了变化（自动导入/同名覆盖），请刷新列表并选中该脚本。
        ''' 面板未打开时不做事（下次打开会从磁盘重新加载）。
        ''' </summary>
        Public Sub NotifyParseScriptUpdated(scriptId As String, scriptName As String, Optional replaced As Boolean = False)
            SendMessageToWebView(OCR_PANE_KEY, New With {
                .type = "parseScriptUpdated",
                .scriptId = If(scriptId, ""),
                .scriptName = If(scriptName, ""),
                .replaced = replaced
            })
        End Sub

        ''' <summary>
        ''' 创建主任务窗格
        ''' </summary>
        Public Function CreateMainTaskPane(hostControl As UserControl, title As String, width As Integer) As CustomTaskPane
            Dim taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(hostControl, title)
            taskPane.Width = width
            taskPane.Visible = False
            taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight

            _taskPanes(MAIN_PANE_KEY) = New TaskPaneInfo With {
                .TaskPane = taskPane,
                .HostControl = hostControl,
                .WebView = Nothing,
                .HtmlPath = Nothing
            }

            _logger?.LogInfo("TASKPANE", "主任务窗格创建完成")
            Return taskPane
        End Function

        ''' <summary>
        ''' 异步创建 OCR 任务窗格
        ''' </summary>
        Public Async Function CreateOcrTaskPaneAsync() As Task
            If TryReuseOpenPane(OCR_PANE_KEY, "识图入表") Then Return

            Dim iAmCreator As Boolean = False
            SyncLock _createLock
                If Not _creatingKeys.Contains(OCR_PANE_KEY) Then
                    _creatingKeys.Add(OCR_PANE_KEY)
                    iAmCreator = True
                End If
            End SyncLock

            If Not iAmCreator Then
                ' 另一次创建还在跑（WebView 初始化要几秒）。等它结束，不要再 Add 一个窗格。
                For i As Integer = 1 To 50
                    Await Task.Delay(100)
                    If TryReuseOpenPane(OCR_PANE_KEY, "识图入表") Then Return
                Next
                Return
            End If

            Try
                If TryReuseOpenPane(OCR_PANE_KEY, "识图入表") Then Return
                Await CreateWebViewPaneAsync(OCR_PANE_KEY, "ocr-panel.html", "识图入表", "OCR")
            Finally
                SyncLock _createLock
                    _creatingKeys.Remove(OCR_PANE_KEY)
                End SyncLock
            End Try
        End Function

        ''' <summary>
        ''' 异步创建设置任务窗格。
        ''' 注意：原统一的 settings-panel.html 已被拆分为 ai-settings-panel.html（AI 设置）
        ''' 与 ocr-settings-panel.html（OCR 设置）。此处按 initialTab 路由到对应面板：
        '''   - "ocr"  → OCR 模型设置
        '''   - 其他/空 → AI 模型设置（默认）
        ''' 旧的统一面板用 switchTab 切换内部标签，拆分后改为直接加载对应面板文件，
        ''' 故不再发送 switchTab 消息。
        ''' </summary>
        Public Async Function CreateSettingsTaskPaneAsync(Optional initialTab As String = "") As Task
            Dim useOcr As Boolean = String.Equals(initialTab, "ocr", StringComparison.OrdinalIgnoreCase)
            Dim htmlFile As String = If(useOcr, "ocr-settings-panel.html", "ai-settings-panel.html")
            Dim title As String = If(useOcr, "OCR 模型设置", "AI 模型设置")
            Dim displayName As String = If(useOcr, "OCR 模型设置", "AI 模型设置")

            If _taskPanes.ContainsKey(SETTINGS_PANE_KEY) AndAlso _taskPanes(SETTINGS_PANE_KEY).TaskPane IsNot Nothing Then
                ShowTaskPane(SETTINGS_PANE_KEY)
                Return
            End If

            Await CreateWebViewPaneAsync(SETTINGS_PANE_KEY, htmlFile, title, displayName)
        End Function

        ''' <summary>
        ''' 异步创建手机遥控任务窗格
        ''' </summary>
        Public Async Function CreateMobileTaskPaneAsync() As Task
            If _taskPanes.ContainsKey(MOBILE_PANE_KEY) AndAlso _taskPanes(MOBILE_PANE_KEY).TaskPane IsNot Nothing Then
                ShowTaskPane(MOBILE_PANE_KEY)
                Return
            End If
            Await CreateWebViewPaneAsync(MOBILE_PANE_KEY, "mobile-control-panel.html", "手机遥控 OCR", "手机遥控")
        End Function

        ''' <summary>
        ''' 异步创建 AI 模型设置任务窗格
        ''' </summary>
        Public Async Function CreateAiSettingsTaskPaneAsync() As Task
            If _taskPanes.ContainsKey(AI_SETTINGS_PANE_KEY) AndAlso _taskPanes(AI_SETTINGS_PANE_KEY).TaskPane IsNot Nothing Then
                ShowTaskPane(AI_SETTINGS_PANE_KEY)
                Return
            End If
            Await CreateWebViewPaneAsync(AI_SETTINGS_PANE_KEY, "ai-settings-panel.html", "AI 模型设置", "AI 模型设置")
        End Function

        ''' <summary>
        ''' 异步创建 OCR 识别设置任务窗格
        ''' </summary>
        Public Async Function CreateOcrSettingsTaskPaneAsync() As Task
            If _taskPanes.ContainsKey(OCR_SETTINGS_PANE_KEY) AndAlso _taskPanes(OCR_SETTINGS_PANE_KEY).TaskPane IsNot Nothing Then
                ShowTaskPane(OCR_SETTINGS_PANE_KEY)
                Return
            End If
            Await CreateWebViewPaneAsync(OCR_SETTINGS_PANE_KEY, "ocr-settings-panel.html", "OCR 模型设置", "OCR 模型设置")
        End Function

        ''' <summary>
        ''' 异步创建 PPT 主题模板编辑器（宽窗格，注入 excelBridge）
        ''' </summary>
        Public Async Function CreatePptThemeEditorTaskPaneAsync() As Task
            If _taskPanes.ContainsKey(PPT_THEME_PANE_KEY) AndAlso _taskPanes(PPT_THEME_PANE_KEY).TaskPane IsNot Nothing Then
                ShowTaskPane(PPT_THEME_PANE_KEY)
                Return
            End If
            Await CreateWebViewPaneAsync(PPT_THEME_PANE_KEY, "ppt-theme-editor.html", HostRibbonLocale.PptThemePaneTitle(), "PPT 主题模板", Nothing, PPT_THEME_PANE_WIDTH)
        End Function

        ''' <summary>
        ''' 核心方法：统一创建 WebView2 任务窗格，消除重复代码
        ''' </summary>
        ''' <param name="key">窗格键名</param>
        ''' <param name="htmlFileName">HTML 文件名</param>
        ''' <param name="title">窗格标题</param>
        ''' <param name="logDisplayName">日志显示名称</param>
        ''' <param name="postWebViewSetup">可选的 WebView 额外设置回调（如 NavigationCompleted 处理器）</param>
        ''' <param name="paneWidth">任务窗格宽度（默认 480；主题编辑器等宽布局可传更大值）</param>
        Private Async Function CreateWebViewPaneAsync(
            key As String,
            htmlFileName As String,
            title As String,
            logDisplayName As String,
            Optional postWebViewSetup As Action(Of WebView2) = Nothing,
            Optional paneWidth As Integer = 0) As Task

            Try
                _logger?.LogInfo("TASKPANE", $"开始异步创建{logDisplayName}任务窗格")
                Dim width As Integer = If(paneWidth > 0, paneWidth, DEFAULT_PANE_WIDTH)

                Dim webView = New WebView2() With {.Dock = DockStyle.Fill}

                Dim webViewEnvironment = Await CoreWebView2Environment.CreateAsync()
                Await webView.EnsureCoreWebView2Async(webViewEnvironment)

                webView.CoreWebView2.Settings.IsScriptEnabled = True
                webView.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = True
                webView.CoreWebView2.Settings.IsWebMessageEnabled = True
                webView.CoreWebView2.Settings.IsStatusBarEnabled = False
                webView.CoreWebView2.Settings.AreDevToolsEnabled = True

                ' 预创建窗格信息并跟踪页面加载状态，避免消息在 WebView 尚未就绪时丢失
                Dim info As New TaskPaneInfo With {.WebView = webView, .PreferredWidth = width}
                AddHandler webView.CoreWebView2.NavigationCompleted,
                    Async Sub(s, e)
                        info.IsReady = e.IsSuccess
                        System.Diagnostics.Debug.WriteLine($"[TaskPaneManager] {logDisplayName} 导航完成，成功={e.IsSuccess}")
                        If e.IsSuccess AndAlso info.PendingMessage IsNot Nothing Then
                            Try
                                Dim pendingJson = Newtonsoft.Json.JsonConvert.SerializeObject(info.PendingMessage)
                                Dim script = $"(function() {{ if (typeof window.handleMobileOcrResult === 'function') {{ window.handleMobileOcrResult({pendingJson}); return JSON.stringify({{status:'ok'}}); }} else {{ return JSON.stringify({{status:'function_not_found'}}); }} }})()"
                                Dim result = Await ExecuteScriptAsync(key, script)
                                Dim resultStr = If(result?.ToString(), "")
                                info.PendingMessage = Nothing
                                System.Diagnostics.Debug.WriteLine($"[TaskPaneManager] 待处理消息发送结果: {resultStr}")
                                If resultStr.Contains("function_not_found") Then
                                    MessageBox.Show("OCR 面板页面未正确加载，请关闭后重新打开 OCR 面板再试。", "移动端同步失败", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                End If
                            Catch ex As Exception
                                System.Diagnostics.Debug.WriteLine($"[TaskPaneManager] 发送待处理消息失败: {ex.Message}")
                            End Try
                        End If
                    End Sub

                ' 调用可选的额外设置回调（如 NavigationCompleted 处理器）
                postWebViewSetup?.Invoke(webView)

                Dim assemblyDir = GetAssemblyDirectory()
                Dim uiFolder = System.IO.Path.Combine(assemblyDir, "UI")
                Dim htmlPath = System.IO.Path.Combine(uiFolder, htmlFileName)

                webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                    VIRTUAL_HOST_NAME,
                    uiFolder,
                    CoreWebView2HostResourceAccessKind.Allow)

                ' 临时目录也映射为虚拟主机：截屏/导出等生成在 Temp 下的文件，
                ' 前端用 https://screenassets.local/xxx.png 同源引用即可显示（绕开 file:// 跨源与超大 base64）
                Try
                    Dim tempDir = System.IO.Path.GetTempPath()
                    If System.IO.Directory.Exists(tempDir) Then
                        webView.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            SCREEN_ASSETS_HOST,
                            tempDir,
                            CoreWebView2HostResourceAccessKind.Allow)
                        _logger?.LogInfo("TASKPANE", $"临时目录虚拟主机映射已设置: {SCREEN_ASSETS_HOST} -> {tempDir}")
                    End If
                Catch ex As Exception
                    _logger?.LogError("TASKPANE", "临时目录虚拟主机映射失败(非致命)", ex)
                End Try

                If _excelBridge Is Nothing Then
                    _excelBridge = ExcelAIVSTOAddIn.ExcelBridge.Instance
                End If
                webView.CoreWebView2.AddHostObjectToScript("excelBridge", _excelBridge)

                Dim cacheBust = DateTime.Now.ToString("yyyyMMddHHmmss")
                webView.Source = New Uri($"https://{VIRTUAL_HOST_NAME}/{htmlFileName}?v={cacheBust}")

                Dim control = New UserControl() With {
                    .Width = width,
                    .Height = DEFAULT_PANE_HEIGHT
                }
                control.Controls.Add(webView)

                If _taskPanes.ContainsKey(key) AndAlso PaneStillAlive(_taskPanes(key)) Then
                    Try
                        control.Controls.Remove(webView)
                        webView.Dispose()
                        control.Dispose()
                    Catch
                    End Try
                    ShowTaskPane(key)
                    _logger?.LogInfo("TASKPANE", $"{logDisplayName}窗格已存在，跳过重复创建")
                    Return
                End If

                Dim taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(control, title)
                taskPane.Width = width
                taskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight
                taskPane.Visible = True

                info.TaskPane = taskPane
                info.HostControl = control
                info.HtmlPath = htmlPath
                _taskPanes(key) = info

                _logger?.LogInfo("TASKPANE", $"{logDisplayName}任务窗格创建完成")

            Catch ex As Exception
                _logger?.LogError("TASKPANE", $"创建{logDisplayName}任务窗格失败", ex)
                MessageBox.Show($"创建{logDisplayName}面板失败: " & ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Function

        ''' <summary>
        ''' 识图入表已经在屏幕上时直接拿来用。Excel 里同名窗格再 Add 一次会并排多出一块空面板。
        ''' </summary>
        Private Function TryReuseOpenPane(key As String, title As String) As Boolean
            If _taskPanes.ContainsKey(key) AndAlso PaneStillAlive(_taskPanes(key)) Then
                ShowTaskPane(key)
                Return True
            End If

            Try
                For Each pane As CustomTaskPane In Globals.ThisAddIn.CustomTaskPanes
                    If pane Is Nothing Then Continue For
                    Dim paneTitle As String = Nothing
                    Try
                        paneTitle = pane.Title
                    Catch
                        Continue For
                    End Try
                    If Not String.Equals(paneTitle, title, StringComparison.Ordinal) Then Continue For

                    Dim host = pane.Control
                    Dim web = FindWebView(host)
                    _taskPanes(key) = New TaskPaneInfo With {
                        .TaskPane = pane,
                        .HostControl = TryCast(host, UserControl),
                        .WebView = web,
                        .PreferredWidth = If(pane.Width > 0, pane.Width, DEFAULT_PANE_WIDTH)
                    }
                    ShowTaskPane(key)
                    _logger?.LogInfo("TASKPANE", "复用已打开的窗格: " & title)
                    Return True
                Next
            Catch ex As Exception
                _logger?.LogError("TASKPANE", "查找已打开窗格失败", ex)
            End Try
            Return False
        End Function

        Private Shared Function PaneStillAlive(info As TaskPaneInfo) As Boolean
            If info Is Nothing OrElse info.TaskPane Is Nothing Then Return False
            Try
                Dim ignored = info.TaskPane.Control
                Return ignored IsNot Nothing
            Catch
                Return False
            End Try
        End Function

        Private Shared Function FindWebView(root As Control) As WebView2
            If root Is Nothing Then Return Nothing
            Dim direct = TryCast(root, WebView2)
            If direct IsNot Nothing Then Return direct
            For Each child As Control In root.Controls
                Dim found = FindWebView(child)
                If found IsNot Nothing Then Return found
            Next
            Return Nothing
        End Function

        Public Sub ShowTaskPane(key As String)
            If Not _taskPanes.ContainsKey(key) OrElse _taskPanes(key).TaskPane Is Nothing Then Return
            Dim info = _taskPanes(key)
            Try
                ' 已经可见时再设 Visible = True，Excel 会在旁边再停靠一个同名窗格。
                If Not info.TaskPane.Visible Then
                    info.TaskPane.Visible = True
                End If
                Dim width = If(info.PreferredWidth > 0, info.PreferredWidth, DEFAULT_PANE_WIDTH)
                If info.TaskPane.Width <> width Then
                    info.TaskPane.Width = width
                End If
            Catch ex As Exception
                _logger?.LogError("TASKPANE", "显示窗格失败: " & key, ex)
            End Try
            Try
                info.WebView?.Focus()
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 隐藏指定任务窗格
        ''' </summary>
        Public Sub HideTaskPane(key As String)
            If _taskPanes.ContainsKey(key) AndAlso _taskPanes(key).TaskPane IsNot Nothing Then
                _taskPanes(key).TaskPane.Visible = False
            End If
        End Sub

        ''' <summary>
        ''' 切换任务窗格显示/隐藏
        ''' </summary>
        Public Sub ToggleTaskPane(key As String)
            If _taskPanes.ContainsKey(key) AndAlso _taskPanes(key).TaskPane IsNot Nothing Then
                Dim taskPane = _taskPanes(key).TaskPane
                Try
                    Dim testVisible = taskPane.Visible
                    If taskPane.Visible Then
                        taskPane.Visible = False
                    Else
                        taskPane.Visible = True
                        taskPane.Width = DEFAULT_PANE_WIDTH
                        _taskPanes(key).WebView?.Focus()
                    End If
                Catch ex As Exception
                    DisposeTaskPane(key)
                    If key = OCR_PANE_KEY Then
                        SafeFork(CreateOcrTaskPaneAsync(), NameOf(ToggleTaskPane) & "_Catch_OCR")
                    ElseIf key = SETTINGS_PANE_KEY Then
                        SafeFork(CreateSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_Catch_Settings")
                    ElseIf key = MOBILE_PANE_KEY Then
                        SafeFork(CreateMobileTaskPaneAsync(), NameOf(ToggleTaskPane) & "_Catch_Mobile")
                    ElseIf key = AI_SETTINGS_PANE_KEY Then
                        SafeFork(CreateAiSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_Catch_Ai")
                    ElseIf key = OCR_SETTINGS_PANE_KEY Then
                        SafeFork(CreateOcrSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_Catch_OcrSet")
                    End If
                End Try
            Else
                If key = OCR_PANE_KEY Then
                    SafeFork(CreateOcrTaskPaneAsync(), NameOf(ToggleTaskPane) & "_New_OCR")
                ElseIf key = SETTINGS_PANE_KEY Then
                    SafeFork(CreateSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_New_Settings")
                ElseIf key = MOBILE_PANE_KEY Then
                    SafeFork(CreateMobileTaskPaneAsync(), NameOf(ToggleTaskPane) & "_New_Mobile")
                ElseIf key = AI_SETTINGS_PANE_KEY Then
                    SafeFork(CreateAiSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_New_Ai")
                ElseIf key = OCR_SETTINGS_PANE_KEY Then
                    SafeFork(CreateOcrSettingsTaskPaneAsync(), NameOf(ToggleTaskPane) & "_New_OcrSet")
                End If
            End If
        End Sub

        ''' <summary>
        ''' 发送消息到任务窗格的 WebView
        ''' </summary>
        Public Sub SendMessageToWebView(key As String, message As Object)
            If _taskPanes.ContainsKey(key) AndAlso _taskPanes(key).WebView IsNot Nothing Then
                Dim webView = _taskPanes(key).WebView
                If webView.CoreWebView2 IsNot Nothing Then
                    Try
                        Dim msgJson = Newtonsoft.Json.JsonConvert.SerializeObject(message)
                        webView.CoreWebView2.PostWebMessageAsString(msgJson)
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine($"发送消息到 {key} 失败: " & ex.Message)
                    End Try
                End If
            End If
        End Sub

        ''' <summary>
        ''' 确保 OCR 面板已创建并加载完成，然后发送消息（解决移动端结果到达时 WebView 尚未就绪的问题）
        ''' </summary>
        Public Sub SendMessageToOcrPanelWhenReady(message As Object)
            Try
                _logger?.LogInfo("TASKPANE", "准备发送移动端 OCR 结果到面板")
                ' 若 OCR 面板尚未创建，先创建（同步等待）
                If Not _taskPanes.ContainsKey(OCR_PANE_KEY) OrElse _taskPanes(OCR_PANE_KEY).TaskPane Is Nothing Then
                    _logger?.LogInfo("TASKPANE", "OCR 面板不存在，准备创建")
                    CreateOcrTaskPaneAsync().GetAwaiter().GetResult()
                Else
                    ShowTaskPane(OCR_PANE_KEY)
                End If

                Dim webView = If(_taskPanes.ContainsKey(OCR_PANE_KEY), _taskPanes(OCR_PANE_KEY).WebView, Nothing)
                If webView Is Nothing Then
                    _logger?.LogError("TASKPANE", "OCR WebView 为 Nothing", New Exception("WebView null"))
                    Return
                End If

                Dim msgJson = Newtonsoft.Json.JsonConvert.SerializeObject(message)
                _logger?.LogInfo("TASKPANE", $"OCR 消息长度={msgJson.Length}")

                ' 所有 CoreWebView2 访问都必须在 UI 线程，用 Control.Invoke 切换
                webView.Invoke(New Action(Sub()
                    Try
                        If webView.CoreWebView2 Is Nothing Then
                            _logger?.LogError("TASKPANE", "CoreWebView2 未初始化", New Exception("CoreWebView2 null"))
                            Return
                        End If

                        ' 通道 1：PostWebMessageAsString（同步，官方推荐）
                        webView.CoreWebView2.PostWebMessageAsString(msgJson)
                        _logger?.LogInfo("TASKPANE", "已通过 PostWebMessageAsString 发送消息")

                        ' 通道 2：ExecuteScriptAsync 直接调用页面函数
                        Dim script = $"(function() {{ try {{ if (typeof window.handleMobileOcrResult === 'function') {{ window.handleMobileOcrResult({msgJson}); return JSON.stringify({{status:'ok'}}); }} else {{ return JSON.stringify({{status:'function_not_found'}}); }} }} catch (err) {{ return JSON.stringify({{status:'error', message: err.message}}); }} }})()"
                        Dim task = webView.CoreWebView2.ExecuteScriptAsync(script)
                        task.ContinueWith(Sub(t)
                                              Try
                                                  If t.IsFaulted Then
                                                      _logger?.LogError("TASKPANE", "ExecuteScript 失败: " & t.Exception?.Message, t.Exception)
                                                  Else
                                                      _logger?.LogInfo("TASKPANE", $"ExecuteScript 结果: {t.Result}")
                                                  End If
                                              Catch ex As Exception  ' 
                                              End Try
                                          End Sub)
                    Catch ex1 As Exception
                        _logger?.LogError("TASKPANE", "Invoke 内部失败: " & ex1.Message, ex1)
                    End Try
                End Sub))

            Catch ex As Exception
                _logger?.LogError("TASKPANE", $"发送 OCR 面板消息失败: {ex.Message}", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 向每一个已打开的任务窗格执行脚本（语言切换要通知 PPT 主题模板等独立 WebView）。
        ''' </summary>
        Public Sub ExecuteScriptAll(script As String)
            If String.IsNullOrEmpty(script) Then Return
            For Each kvp In _taskPanes
                Try
                    Dim webView = kvp.Value.WebView
                    If webView IsNot Nothing AndAlso webView.CoreWebView2 IsNot Nothing Then
                        webView.CoreWebView2.ExecuteScriptAsync(script)
                    End If
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine($"向 {kvp.Key} 推送脚本失败: " & ex.Message)
                End Try
            Next
        End Sub

        ''' <summary>
        ''' Excel 的 CustomTaskPane.Title 是只读的（创建时由 CustomTaskPanes.Add 定死），
        ''' 不能在语言切换后赋值，否则 BC30526。窗格里的页面由 ExecuteScriptAll 刷新；
        ''' 标题在下次创建窗格时用 HostRibbonLocale.PptThemePaneTitle() 带上当前语言。
        ''' </summary>
        Public Sub SetPaneTitle(key As String, title As String)
            ' 参数保留，供调用方与 Word/PPT 的 HostPaneManager.SetPaneTitle 对齐。
            ' Title 只读，这里只校验入参，避免未使用参数告警。
            If String.IsNullOrEmpty(key) OrElse String.IsNullOrEmpty(title) Then Return
        End Sub

        ''' <summary>
        ''' 执行脚本到任务窗格的 WebView
        ''' </summary>
        Public Async Function ExecuteScriptAsync(key As String, script As String) As Task(Of Object)
            If _taskPanes.ContainsKey(key) AndAlso _taskPanes(key).WebView IsNot Nothing Then
                Dim webView = _taskPanes(key).WebView
                If webView.CoreWebView2 IsNot Nothing Then
                    Return Await webView.CoreWebView2.ExecuteScriptAsync(script)
                End If
            End If
            Return Nothing
        End Function

        ''' <summary>
        ''' 关闭指定的任务窗格
        ''' </summary>
        Public Sub CloseTaskPane(key As String)
            If _taskPanes.ContainsKey(key) Then
                Dim info = _taskPanes(key)
                Try
                    If info.TaskPane IsNot Nothing Then
                        info.TaskPane.Visible = False
                        Globals.ThisAddIn.CustomTaskPanes.Remove(info.TaskPane)
                    End If
                Catch ex As Exception  ' 
                End Try
                DisposeTaskPane(key)
            End If
        End Sub

        ''' <summary>
        ''' 停止所有任务窗格的 WebView2 渲染
        ''' </summary>
        Public Sub StopAllWebViews()
            For Each kvp In _taskPanes
                Try
                    If kvp.Value.WebView IsNot Nothing AndAlso kvp.Value.WebView.CoreWebView2 IsNot Nothing Then
                        kvp.Value.WebView.CoreWebView2.Stop()
                    End If
                Catch ex As Exception  ' 
                End Try
            Next
        End Sub

        ''' <summary>
        ''' 在关闭前移除所有 WebView2 子控件，防止阻塞
        ''' </summary>
        Public Sub RemoveWebViewControlsBeforeClose()
            For Each kvp In _taskPanes
                Dim info = kvp.Value
                If info.TaskPane IsNot Nothing Then
                    Try
                        Dim ctrl = info.TaskPane.Control
                        If ctrl IsNot Nothing AndAlso Not ctrl.IsDisposed Then
                            Dim wvControls = New List(Of Control)()
                            For Each child As Control In ctrl.Controls
                                If TypeOf child Is WebView2 Then
                                    wvControls.Add(child)
                                End If
                            Next
                            For Each child In wvControls
                                ctrl.Controls.Remove(child)
                            Next
                        End If
                    Catch ex As Exception  ' 
                    End Try
                    Try : Globals.ThisAddIn.CustomTaskPanes.Remove(info.TaskPane) : Catch : End Try
                End If
            Next
        End Sub

        ''' <summary>
        ''' 释放所有资源（WebView2 异步释放，避免阻塞 Excel 关闭）
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            StopAllWebViews()
            RemoveWebViewControlsBeforeClose()

            ' 把 WebView2/HostControl 的 Dispose 放到后台线程，避免同步等待浏览器进程退出阻塞 Excel
            Dim panesToDispose = _taskPanes.Values.ToList()
            _taskPanes.Clear()

            If panesToDispose.Count > 0 Then
                Task.Run(Sub()
                             For Each info In panesToDispose
                                 Try : info.WebView?.Dispose() : Catch : End Try
                                 Try : info.HostControl?.Dispose() : Catch : End Try
                             Next
                         End Sub)
            End If
        End Sub

        Private Function GetTaskPane(key As String) As CustomTaskPane
            If _taskPanes.ContainsKey(key) Then
                Return _taskPanes(key).TaskPane
            End If
            Return Nothing
        End Function

        Private Sub DisposeTaskPane(key As String)
            If _taskPanes.ContainsKey(key) Then
                _taskPanes(key).WebView?.Dispose()
                _taskPanes.Remove(key)
            End If
        End Sub

        Private Shared Function GetAssemblyDirectory() As String
            Dim codeBase = System.Reflection.Assembly.GetExecutingAssembly().CodeBase
            Dim uri = New Uri(codeBase)
            Return System.IO.Path.GetDirectoryName(uri.LocalPath)
        End Function

        ''' <summary>
        ''' 任务窗格信息
        ''' </summary>
        Private Class TaskPaneInfo
            Public Property TaskPane As CustomTaskPane
            Public Property HostControl As UserControl
            Public Property WebView As WebView2
            Public Property HtmlPath As String
            Public Property IsReady As Boolean = False
            Public Property PendingMessage As Object
            Public Property PreferredWidth As Integer = 0
        End Class

    End Class