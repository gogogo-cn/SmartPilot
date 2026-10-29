' ============================================================================
' HostWebViewPanel - WebView2 面板封装（宿主无关）
'
' 复用现有 Excel 版前端：把 UI 目录映射为虚拟主机名 app.excelai.local，
' 并把宿主桥对象以 "excelBridge" 名字注入 JS（前端 200+ 处调用无需改动）。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.IO
Imports System.Windows.Forms
Imports System.Threading.Tasks
Imports Microsoft.Web.WebView2.WinForms
Imports Microsoft.Web.WebView2.Core

Public Class HostWebViewPanel
    Implements IDisposable

    Public Const VIRTUAL_HOST As String = "app.excelai.local"

    Private _host As UserControl
    Private _web As WebView2
    Private _ready As Boolean = False
    Private _hostTag As String = "shared"

    Public Event WebMessageReceived As EventHandler(Of CoreWebView2WebMessageReceivedEventArgs)

    Public Sub New(hostControl As UserControl, Optional hostTag As String = "shared")
        _host = hostControl
        _hostTag = If(String.IsNullOrWhiteSpace(hostTag), "shared", hostTag.Replace(" "c, "_"c))
    End Sub

    Public ReadOnly Property Control As UserControl
        Get
            Return _host
        End Get
    End Property

    Public ReadOnly Property WebView As WebView2
        Get
            Return _web
        End Get
    End Property

    Public ReadOnly Property IsReady As Boolean
        Get
            Return _ready AndAlso _web IsNot Nothing AndAlso _web.CoreWebView2 IsNot Nothing
        End Get
    End Property

    ''' <summary>初始化 WebView2 并导航到指定页面（如 index.html?host=word）</summary>
    Public Async Function InitializeAsync(uiDir As String, startPage As String, hostObject As Object) As Task
        Try
            _web = New WebView2() With {.Dock = DockStyle.Fill}
            _host.Controls.Add(_web)

            Dim cacheDir = Path.Combine(Path.GetTempPath(), "SmartPilotOffice_EBWebView_" & _hostTag)
            If Not Directory.Exists(cacheDir) Then Directory.CreateDirectory(cacheDir)

            Dim env = Await CoreWebView2Environment.CreateAsync(Nothing, cacheDir)
            Await _web.EnsureCoreWebView2Async(env)

            AddHandler _web.CoreWebView2.WebMessageReceived, AddressOf OnMessage

            If Directory.Exists(uiDir) Then
                Try
                    _web.CoreWebView2.SetVirtualHostNameToFolderMapping(VIRTUAL_HOST, uiDir, CoreWebView2HostResourceAccessKind.Allow)
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[HostWebViewPanel] 虚拟主机映射失败: " & ex.Message)
                End Try
            End If

            If hostObject IsNot Nothing Then
                _web.CoreWebView2.AddHostObjectToScript("excelBridge", hostObject)
            End If

            ' 禁用 HTTP 缓存（与 Excel 版 WebView2Manager 保持一致，见其 Network.setCacheDisabled 调用）。
            ' 必要性：本面板用【持久化】缓存目录 + 【固定】入口 URL（index.html?host=xxx），Chromium 会把
            ' index.html 本体缓存住，进而复用 HTML 里"旧版本戳"的一整套 JS —— 表现为"改了前端却一直跑旧代码"。
            ' 实测该问题曾导致多轮修复在 PPT/Word 里完全看不到效果。
            Try
                Await _web.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Network.setCacheDisabled",
                    "{""cacheDisabled"":true}")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[HostWebViewPanel] 禁用缓存失败(非致命): " & ex.Message)
            End Try

            _ready = True
            ' 入口 URL 追加一次性 cache buster（双保险）：即使缓存策略被环境改写，HTML 本体也不会被复用。
            Dim cb As String = Guid.NewGuid().ToString("N")
            Dim url As String = startPage
            If url.Contains("?") Then
                url &= "&_cb=" & cb
            Else
                url &= "?_cb=" & cb
            End If
            _web.CoreWebView2.Navigate(url)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostWebViewPanel] 初始化失败: " & ex.ToString())
            Throw
        End Try
    End Function

    Private Sub OnMessage(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        RaiseEvent WebMessageReceived(sender, e)
    End Sub

    ''' <summary>向页面推送消息（JSON 字符串）</summary>
    Public Sub PostJson(json As String)
        Try
            If IsReady Then _web.CoreWebView2.PostWebMessageAsJson(json)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostWebViewPanel] PostJson 失败: " & ex.Message)
        End Try
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Try
            If _web IsNot Nothing Then _web.Dispose()
        Catch
        End Try
        _web = Nothing
    End Sub
End Class
