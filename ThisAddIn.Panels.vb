Option Strict On
Option Infer On

Imports System
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Office.Tools
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel

Partial Public NotInheritable Class ThisAddIn

#Region "面板管理"

    Private Const DEFAULT_SIDEBAR_WIDTH As Integer = 480

    ''' <summary>
    ''' 确保 _taskPaneManager 已初始化(不依赖主侧边栏初始化)
    ''' </summary>
    Private Sub EnsureTaskPaneManager()
        If _taskPaneManager IsNot Nothing Then Return
        If _excelBridge Is Nothing Then
            _excelBridge = ExcelAIVSTOAddIn.ExcelBridge.Instance
        End If
        _taskPaneManager = New TaskPaneManager(Me, _chatLogger, _excelBridge)
        _chatLogger?.LogInfo("ADDIN", "TaskPaneManager initialized on demand")
    End Sub

    ''' <summary>
    ''' 切换主聊天面板显示/隐藏(WebView2聊天面板,加载UI/index.html)
    ''' </summary>
    Public Sub ToggleChatPanel()
        Try
            If _chatTaskPane Is Nothing Then
                InitializeChatPanel()
                If _chatTaskPane IsNot Nothing Then
                    _chatTaskPane.Visible = True
                    _chatTaskPane.Width = DEFAULT_SIDEBAR_WIDTH
                End If
                Return
            End If

            Try
                If _chatTaskPane.Visible Then
                    _chatTaskPane.Visible = False
                Else
                    _chatTaskPane.Visible = True
                    _chatTaskPane.Width = DEFAULT_SIDEBAR_WIDTH
                    _webViewManager?.FocusWebView()
                End If
            Catch ex As Exception
                Try : _chatTaskPane.Dispose() : Catch : End Try
                _chatTaskPane = Nothing
                _webViewManager?.Dispose()
                _webViewManager = Nothing
                InitializeChatPanel()
                If _chatTaskPane IsNot Nothing Then
                    _chatTaskPane.Visible = True
                    _chatTaskPane.Width = DEFAULT_SIDEBAR_WIDTH
                End If
            End Try
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Toggle chat panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 初始化主聊天面板(WebView2,加载UI/index.html)
    ''' </summary>
    Private Sub InitializeChatPanel()
        If _chatTaskPane IsNot Nothing Then Return

        _chatHostControl = New UserControl() With {
            .Width = DEFAULT_SIDEBAR_WIDTH,
            .Height = 600,
            .BackColor = Color.White
        }

        _chatTaskPane = CustomTaskPanes.Add(_chatHostControl, "SmartPilot")
        _chatTaskPane.Width = DEFAULT_SIDEBAR_WIDTH
        _chatTaskPane.Visible = False
        _chatTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight

        If _excelBridge Is Nothing Then
            _excelBridge = ExcelAIVSTOAddIn.ExcelBridge.Instance
        End If

        ' 如果_webViewManager已与其他hostControl关联,先释放,确保聊天面板有独立WebView2
        If _webViewManager IsNot Nothing Then
            Try : _webViewManager.Dispose() : Catch : End Try
            _webViewManager = Nothing
        End If
        _webViewManager = New WebView2Manager(_chatHostControl, _chatLogger)

        Dim uiPath As String = WebView2Manager.GetUIPath()
        ' WebView2必须在Excel UI线程(STA)初始化,不能使用Task.Run到后台线程
        Dim initTask As Task = _webViewManager.InitializeAsync(uiPath, _excelBridge)
    End Sub

    ''' <summary>
    ''' 关闭OCR面板
    ''' </summary>
    Public Sub CloseOcrPanel()
        Try
            _taskPaneManager?.HideTaskPane("OCR")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Close OCR panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 关闭设置面板
    ''' </summary>
    Public Sub CloseSettingsPanel()
        Try
            _taskPaneManager?.HideTaskPane("Settings")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Close settings panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 切换OCR面板显示/隐藏
    ''' </summary>
    Public Sub ToggleOcrPanel()
        Try
            Dim ocrPane As Microsoft.Office.Tools.CustomTaskPane = _taskPaneManager?.OcrTaskPane
            If ocrPane Is Nothing Then
                ' 面板不存在 - 创建并显示
                ShowOcrPanel()
            ElseIf Not ocrPane.Visible Then
                ' 面板存在但隐藏(如通过X按钮关闭) - 直接显示,无需重建
                ocrPane.Visible = True
                ocrPane.Width = DEFAULT_SIDEBAR_WIDTH
            Else
                ' 面板可见 - 关闭
                CloseOcrPanel()
            End If
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Toggle OCR panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 切换手机远程面板显示/隐藏
    ''' </summary>
    Public Sub ToggleMobilePanel()
        Try
            Dim mobilePane As Microsoft.Office.Tools.CustomTaskPane = _taskPaneManager?.MobileTaskPane
            If mobilePane Is Nothing Then
                ShowMobilePanel()
            ElseIf Not mobilePane.Visible Then
                mobilePane.Visible = True
                mobilePane.Width = DEFAULT_SIDEBAR_WIDTH
            Else
                CloseMobilePanel()
            End If
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Toggle mobile panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 显示独立手机远程面板侧边栏
    ''' </summary>
    Public Sub ShowMobilePanel()
        Try
            Try
                Bridges.MobileControlBridge.Instance.StartServer()
            Catch startEx As Exception
                _chatLogger?.LogError("ADDIN", "Mobile control server start failed", startEx)
            End Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreateMobileTaskPaneAsync(), NameOf(ShowMobilePanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show mobile panel failed", ex)
            MessageBox.Show("Show mobile panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 关闭手机远程面板(仅隐藏,不销毁)
    ''' </summary>
    Public Sub CloseMobilePanel()
        Try
            _taskPaneManager?.HideTaskPane("Mobile")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Close mobile panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 显示独立OCR识别侧边栏
    ''' </summary>
    Public Sub ShowOcrPanel()
        Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreateOcrTaskPaneAsync(), NameOf(ShowOcrPanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show OCR panel failed", ex)
            MessageBox.Show("Show OCR panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 显示独立设置面板侧边栏
    ''' </summary>
    Public Sub ShowSettingsPanel(Optional initialTab As String = "")
        Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreateSettingsTaskPaneAsync(initialTab), NameOf(ShowSettingsPanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show settings panel failed", ex)
            MessageBox.Show("Show settings panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 显示AI模型设置面板
    ''' </summary>
    Public Sub ShowAiSettingsPanel()
        Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreateAiSettingsTaskPaneAsync(), NameOf(ShowAiSettingsPanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show AI model settings panel failed", ex)
            MessageBox.Show("Show AI model settings panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 显示OCR识别设置面板
    ''' </summary>
    Public Sub ShowOcrSettingsPanel()
        Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreateOcrSettingsTaskPaneAsync(), NameOf(ShowOcrSettingsPanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show OCR settings panel failed", ex)
            MessageBox.Show("Show OCR settings panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 显示 PPT 主题模板编辑器（宽任务窗格，注入 excelBridge）
    ''' </summary>
    Public Sub ShowPptThemeEditorPanel()
        Try
            EnsureTaskPaneManager()
            SafeFork(_taskPaneManager?.CreatePptThemeEditorTaskPaneAsync(), NameOf(ShowPptThemeEditorPanel), _chatLogger)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Show PPT theme editor panel failed", ex)
            MessageBox.Show("打开 PPT 主题编辑器失败: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>关闭 PPT 主题模板编辑器（仅隐藏）</summary>
    Public Sub ClosePptThemeEditorPanel()
        Try
            _taskPaneManager?.HideTaskPane("PptTheme")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Close PPT theme editor panel failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 开始截图识别(先启动截图UI,OCR面板后台预加载)
    ''' </summary>
    Public Sub StartScreenshotOcr()
        Try
            EnsureTaskPaneManager()

            ' 确保截图服务已初始化
            If _screenshotService Is Nothing Then
                _screenshotService = New ScreenshotService(_chatLogger)
            End If

            ' 人已经在识图入表面板里点「截图识别」。直接用这块面板，不要再开一块。
            Dim openWeb = _taskPaneManager.OcrWebView
            If openWeb IsNot Nothing AndAlso openWeb.CoreWebView2 IsNot Nothing Then
                _screenshotService.SetOcrWebView(openWeb)
                _screenshotService.StartOcr()
                Return
            End If

            ' 面板还没开过：只创建一次，就绪后把 WebView 交给截图服务
            Dim preloadTask = _taskPaneManager.CreateOcrTaskPaneAsync()
            Dim uiScheduler As TaskScheduler
            Try
                uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()
            Catch
                uiScheduler = TaskScheduler.Default
            End Try
            preloadTask.ContinueWith(Sub(t)
                Try
                    If t.IsFaulted Then
                        _chatLogger?.LogError("ADDIN", "Preload OCR panel failed", t.Exception)
                        Return
                    End If
                    Dim ocrWebView = _taskPaneManager.OcrWebView
                    If ocrWebView IsNot Nothing Then
                        _screenshotService?.SetOcrWebView(ocrWebView)
                        _chatLogger?.LogInfo("ADDIN", "OCR panel WebView injected to screenshot service")
                    End If
                Catch ex As Exception
                    _chatLogger?.LogError("ADDIN", "Set OCR WebView failed", ex)
                End Try
            End Sub, uiScheduler)

            ' 立即启动截图,不等待OCR面板
            _screenshotService?.StartOcr()
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Start screenshot recognition failed", ex)
            MessageBox.Show("Start screenshot recognition failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' 打开浏览器调试面板(WebView2主面板,F12检查元素)
    ''' </summary>
    Private Sub OnOpenBrowserDebugPanel(sender As Object, e As EventArgs)
        Try
            Dim debugPane As Microsoft.Office.Tools.CustomTaskPane = CustomTaskPanes.FirstOrDefault(Function(p) p.Title = "SmartPilot Debug")
            If debugPane IsNot Nothing Then
                debugPane.Visible = True
                debugPane.Width = DEFAULT_SIDEBAR_WIDTH
                Return
            End If

            Dim hostControl As New UserControl() With {
                .Width = DEFAULT_SIDEBAR_WIDTH,
                .Height = 600,
                .BackColor = Color.White
            }

            debugPane = CustomTaskPanes.Add(hostControl, "SmartPilot Debug")
            debugPane.Width = DEFAULT_SIDEBAR_WIDTH
            debugPane.Visible = True
            debugPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight

            ' 调试面板使用独立WebView2Manager,避免覆盖主聊天面板的_webViewManager
            Dim debugWebViewManager As New WebView2Manager(hostControl, _chatLogger)

            Dim uiPath As String = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "UI", "index.html")
            ' WebView2初始化必须在UI线程(STA)
            Dim initTask As Task = debugWebViewManager.InitializeAsync(uiPath, _excelBridge)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Open browser debug panel failed", ex)
            MessageBox.Show("Open browser debug panel failed: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

#End Region

End Class