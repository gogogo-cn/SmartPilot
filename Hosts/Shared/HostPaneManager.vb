' ============================================================================
' HostPaneManager - 宿主无关的面板管理（聊天 / 设置 / OCR 等）
'
' 通过工厂委托创建 IPaneHost（Word→CustomTaskPane，PPT→SetParent 停靠窗体），
' 每个面板内部是一个 WebView2 页面，桥对象以 excelBridge 注入。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms

Public Class HostPaneManager
    Implements IDisposable

    Private Class PaneEntry
        Public Host As IPaneHost
        Public View As HostWebViewPanel
        Public Page As String
        Public Title As String
    End Class

    Private ReadOnly _factory As Func(Of UserControl, String, IPaneHost)
    Private ReadOnly _uiDir As String
    Private ReadOnly _bridge As Object
    Private ReadOnly _panes As New Dictionary(Of String, PaneEntry)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _hostQuery As String

    Public Sub New(factory As Func(Of UserControl, String, IPaneHost), uiDir As String, bridge As Object, hostQuery As String)
        _factory = factory
        _uiDir = uiDir
        _bridge = bridge
        _hostQuery = If(hostQuery, "")
        Current = Me
    End Sub

    ''' <summary>当前宿主面板管理器（供 Bridge 层关闭面板等）</summary>
    Public Shared Property Current As HostPaneManager

    ''' <summary>可选日志委托（宿主注入，便于诊断异步初始化失败）</summary>
    Public Shared Property Logger As Action(Of String)

    Private Shared Sub Log(msg As String)
        Try
            If Logger IsNot Nothing Then Logger.Invoke("[HostPaneManager] " & msg)
        Catch
        End Try
        System.Diagnostics.Debug.WriteLine("[HostPaneManager] " & msg)
    End Sub

    Public ReadOnly Property UiDirectory As String
        Get
            Return _uiDir
        End Get
    End Property

    ''' <summary>显示（或创建后显示）面板；同步入口，内部异步初始化 WebView2</summary>
    Public Sub Toggle(key As String, page As String, title As String)
        Try
            Dim e As PaneEntry = Nothing
            If _panes.TryGetValue(key, e) AndAlso e IsNot Nothing AndAlso e.Host IsNot Nothing Then
                Dim willShow As Boolean = Not (e.Host.Visible AndAlso e.View IsNot Nothing AndAlso e.View.IsReady)
                Log("Toggle existing pane '" & key & "': willShow=" & willShow)
                e.Host.Visible = willShow
                Return
            End If
            Log("Toggle new pane '" & key & "' -> create")
            SafeRun(ShowAsync(key, page, title))
        Catch ex As Exception
            Log("Toggle FAILED: " & ex.ToString())
        End Try
    End Sub

    Public Sub Show(key As String, page As String, title As String)
        Try
            Dim e As PaneEntry = Nothing
            If _panes.TryGetValue(key, e) AndAlso e IsNot Nothing AndAlso e.Host IsNot Nothing Then
                e.Host.Visible = True
                Return
            End If
            SafeRun(ShowAsync(key, page, title))
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostPaneManager] Show 失败: " & ex.ToString())
        End Try
    End Sub

    Public Sub Hide(key As String)
        Dim e As PaneEntry = Nothing
        If _panes.TryGetValue(key, e) AndAlso e IsNot Nothing AndAlso e.Host IsNot Nothing Then
            e.Host.Visible = False
        End If
    End Sub

    Public Async Function ShowAsync(key As String, page As String, title As String) As Task
        Dim entry As PaneEntry = Nothing
        If _panes.TryGetValue(key, entry) AndAlso entry IsNot Nothing AndAlso entry.Host IsNot Nothing Then
            entry.Host.Visible = True
            Return
        End If

        Dim uc As New UserControl() With {.Dock = DockStyle.Fill}
        Dim tag = If(String.IsNullOrEmpty(_hostQuery), "shared", _hostQuery)
        Dim view As New HostWebViewPanel(uc, tag)
        Dim host = _factory(uc, title)
        If host Is Nothing Then Throw New Exception("窗格宿主创建失败")

        _panes(key) = New PaneEntry With {.Host = host, .View = view, .Page = page, .Title = title}

        Dim url = "https://" & HostWebViewPanel.VIRTUAL_HOST & "/" & page
        If Not String.IsNullOrEmpty(_hostQuery) Then
            url &= If(page.Contains("?"), "&", "?") & "host=" & _hostQuery
        End If

        Log("ShowAsync: initializing WebView2 -> " & url)
        Await view.InitializeAsync(_uiDir, url, _bridge)
        Log("ShowAsync: WebView2 initialized")
        host.Visible = True
        Log("ShowAsync: pane visible")
    End Function

    ''' <summary>向每个已打开的面板执行脚本（语言切换通知 PPT 主题模板等独立页面）。</summary>
    Public Sub ExecuteScriptAll(script As String)
        If String.IsNullOrEmpty(script) Then Return
        For Each kv As KeyValuePair(Of String, PaneEntry) In _panes
            Try
                Dim view As HostWebViewPanel = kv.Value.View
                If view IsNot Nothing AndAlso view.IsReady AndAlso view.WebView IsNot Nothing AndAlso view.WebView.CoreWebView2 IsNot Nothing Then
                    view.WebView.CoreWebView2.ExecuteScriptAsync(script)
                End If
            Catch ex As Exception
                Log("ExecuteScriptAll " & kv.Key & " 失败: " & ex.Message)
            End Try
        Next
    End Sub

    ''' <summary>更新面板标题。Word 的 CustomTaskPane 创建后标题不可变，赋值会被宿主忽略。</summary>
    Public Sub SetPaneTitle(key As String, title As String)
        If String.IsNullOrEmpty(key) OrElse String.IsNullOrEmpty(title) Then Return
        Dim e As PaneEntry = Nothing
        If Not _panes.TryGetValue(key, e) OrElse e Is Nothing OrElse e.Host Is Nothing Then Return
        Try
            e.Host.Title = title
        Catch ex As Exception
            Log("SetPaneTitle " & key & " 失败: " & ex.Message)
        End Try
    End Sub

    ''' <summary>关闭并释放全部面板（Shutdown 时调用）</summary>
    Public Sub Dispose() Implements IDisposable.Dispose
        For Each kv As KeyValuePair(Of String, PaneEntry) In _panes
            Try
                If kv.Value.View IsNot Nothing Then kv.Value.View.Dispose()
                If kv.Value.Host IsNot Nothing Then kv.Value.Host.Dispose()
            Catch
            End Try
        Next
        _panes.Clear()
    End Sub

    Private Shared Sub SafeRun(t As Task)
        If t Is Nothing Then Return
        Dim discard = t.ContinueWith(Sub(x)
                                         If x.IsFaulted AndAlso x.Exception IsNot Nothing Then
                                             Dim msg = If(x.Exception.InnerException IsNot Nothing, x.Exception.InnerException.ToString(), x.Exception.Message)
                                             System.Diagnostics.Debug.WriteLine("[HostPaneManager] 异步异常: " & msg)
                                             Log("异步异常: " & msg)
                                         End If
                                     End Sub, TaskScheduler.Default)
    End Sub
End Class
