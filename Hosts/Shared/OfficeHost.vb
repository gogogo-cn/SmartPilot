' ============================================================================
' OfficeHost - 多宿主（Excel / Word / PowerPoint）通用宿主抽象层
'
' 背景：VSTO 是单宿主模型（一个工程只能挂一个 Office 应用），且 PowerPoint
'       不提供 CustomTaskPane。因此：
'        - Word   → 使用 VSTO 原生 CustomTaskPane（CustomTaskPaneHost）
'        - PPT    → 使用 Win32 SetParent 把无边框窗体嵌到 PPT 主窗口右侧
'                   （DockedFormHost），视觉上与任务窗格一致
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Drawing
Imports System.Windows.Forms

''' <summary>宿主类型</summary>
Public Enum OfficeHostKind
    Excel = 0
    Word = 1
    PowerPoint = 2
End Enum

''' <summary>统一窗格宿主抽象：屏蔽 Word（CustomTaskPane）与 PPT（自绘窗体）的差异</summary>
Public Interface IPaneHost
    ReadOnly Property Control As UserControl
    Property Visible As Boolean
    Property Width As Integer
    Property Title As String
    Sub Dispose()
End Interface

' ---------------------------------------------------------------------------
' Win32 声明（PPT 停靠窗体需要）
' ---------------------------------------------------------------------------
Friend Module Win32
    Public Declare Function SetParent Lib "user32.dll" (hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
    Public Declare Function SetWindowPos Lib "user32.dll" (hWnd As IntPtr, hWndInsertAfter As IntPtr,
        x As Integer, y As Integer, cx As Integer, cy As Integer, uFlags As Integer) As Boolean
    Public Declare Function GetClientRect Lib "user32.dll" (hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    Public Declare Function GetWindowRect Lib "user32.dll" (hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    Public Declare Function IsIconic Lib "user32.dll" (hWnd As IntPtr) As Boolean
    Public Declare Function ShowWindow Lib "user32.dll" (hWnd As IntPtr, nCmdShow As Integer) As Boolean
    Public Declare Function IsWindow Lib "user32.dll" (hWnd As IntPtr) As Boolean
    Public Declare Function IsWindowVisible Lib "user32.dll" (hWnd As IntPtr) As Boolean
    Public Declare Function GetParent Lib "user32.dll" (hWnd As IntPtr) As IntPtr
    Public Declare Auto Function FindWindowEx Lib "user32.dll" (hWndParent As IntPtr, hWndChildAfter As IntPtr,
        lpszClass As String, lpszWindow As String) As IntPtr
    Public Declare Auto Function GetClassName Lib "user32.dll" (hWnd As IntPtr, lpClassName As System.Text.StringBuilder,
        nMaxCount As Integer) As Integer
    Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (hWnd As IntPtr, nIndex As Integer) As Integer
    Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (hWnd As IntPtr, nIndex As Integer,
        dwNewLong As Integer) As Integer

    Public Const GWL_STYLE As Integer = -16
    Public Const WS_CHILD As Integer = &H40000000
    Public Const WS_POPUP As Integer = &H80000000

    Public Delegate Function EnumChildProc(hWnd As IntPtr, lParam As IntPtr) As Boolean
    Public Declare Auto Function EnumChildWindows Lib "user32.dll" (hWndParent As IntPtr,
        proc As EnumChildProc, lParam As IntPtr) As Boolean

    <Runtime.InteropServices.StructLayout(Runtime.InteropServices.LayoutKind.Sequential)>
    Public Structure RECT
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    Public Const SWP_NOZORDER As Integer = &H4
    Public Const SWP_SHOWWINDOW As Integer = &H40
    Public Const SWP_NOACTIVATE As Integer = &H10
    Public Const SW_HIDE As Integer = 0
    Public Const SW_SHOW As Integer = 5
End Module

' ---------------------------------------------------------------------------
' PowerPoint（及未来非 VSTO 宿主）：无边框窗体 + SetParent 嵌入宿主窗口右侧
' （PowerPoint 既无 VSTO 宿主程序集，也不提供 CustomTaskPane，故用此方案）
' ---------------------------------------------------------------------------
''' <summary>把 IntPtr 包装成 IWin32Window，用于 Form.Show(owner)（owned 窗口随 owner 最小化、总在其之上）</summary>
Friend Class Win32WindowWrapper
    Implements IWin32Window

    Private _handle As IntPtr

    Public Sub New(handle As IntPtr)
        _handle = handle
    End Sub

    Public ReadOnly Property Handle As IntPtr Implements IWin32Window.Handle
        Get
            Return _handle
        End Get
    End Property
End Class

''' <summary>
''' 无边框但四边/四角可拖拽调整大小的窗体（WM_NCHITTEST 手工返回 HT* 边缘命中）。
''' 用于让 PPT 侧栏既"无边框贴合内容区"，又能自由调整宽高。
''' </summary>
Friend Class BorderlessResizeForm
    Inherits Form

    Private Const WM_NCHITTEST As Integer = &H84
    Private Const GripSize As Integer = 8

    Public Sub New()
        MyBase.New()
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_NCHITTEST Then
            Try
                ' LParam 屏幕坐标（含负值，用 Short 截断做符号扩展）
                Dim sx As Short = CShort(m.LParam.ToInt32() And &HFFFF)
                Dim sy As Short = CShort((m.LParam.ToInt32() >> 16) And &HFFFF)
                Dim client = PointToClient(New Point(sx, sy))
                Dim w = ClientSize.Width
                Dim h = ClientSize.Height
                Dim leftZone = client.X <= GripSize
                Dim rightZone = client.X >= w - GripSize
                Dim topZone = client.Y <= GripSize
                Dim bottomZone = client.Y >= h - GripSize

                If topZone AndAlso leftZone Then
                    m.Result = CType(13, IntPtr) : Return   ' HTTOPLEFT
                ElseIf topZone AndAlso rightZone Then
                    m.Result = CType(14, IntPtr) : Return   ' HTTOPRIGHT
                ElseIf bottomZone AndAlso leftZone Then
                    m.Result = CType(16, IntPtr) : Return   ' HTBOTTOMLEFT
                ElseIf bottomZone AndAlso rightZone Then
                    m.Result = CType(17, IntPtr) : Return   ' HTBOTTOMRIGHT
                ElseIf leftZone Then
                    m.Result = CType(10, IntPtr) : Return   ' HTLEFT
                ElseIf rightZone Then
                    m.Result = CType(11, IntPtr) : Return   ' HTRIGHT
                ElseIf topZone Then
                    m.Result = CType(12, IntPtr) : Return   ' HTTOP
                ElseIf bottomZone Then
                    m.Result = CType(15, IntPtr) : Return   ' HTBOTTOM
                End If
            Catch
            End Try
        End If
        MyBase.WndProc(m)
    End Sub
End Class

Public Class DockedFormHost
    Implements IPaneHost

    ''' <summary>诊断日志（宿主注入）</summary>
    Public Shared Property Logger As Action(Of String)

    Private Shared Sub Log(msg As String)
        Try
            If Logger IsNot Nothing Then Logger.Invoke("[DockedFormHost] " & msg)
        Catch
        End Try
        System.Diagnostics.Debug.WriteLine("[DockedFormHost] " & msg)
    End Sub

    Private _form As Form
    Private _control As UserControl
    Private _ownerHwnd As IntPtr
    Private _refHwnd As IntPtr
    Private _syncTimer As Timer
    Private _width As Integer
    Private _visible As Boolean
    Private _title As String
    Private _topOffset As Integer
    Private _shownOnce As Boolean
    Private _followRef As Boolean
    ' 用户拖拽后的自定义尺寸（>0 时同步只跟随位置，不再强制尺寸）
    Private _userWidth As Integer
    Private _userHeight As Integer

    ''' <summary>
    ''' Owned 浮动窗模式（PowerPoint 无任务窗格 API，SetParent 内嵌会被文档窗口遮挡）：
    '''   · owner = PPT 主窗口：面板永远浮在 PPT 之上、随其最小化/关闭
    '''   · ref   = 定位参考窗口（MDIClient 内容区）：面板贴其右缘，随窗口缩放跟随
    ''' </summary>
    Public Sub New(ownerHwnd As IntPtr, refHwnd As IntPtr, control As UserControl, title As String,
                   width As Integer, Optional topOffset As Integer = 0, Optional followRef As Boolean = False)
        _ownerHwnd = ownerHwnd
        _refHwnd = If(refHwnd <> IntPtr.Zero, refHwnd, ownerHwnd)
        _control = control
        _title = title
        _width = If(width <= 0, 480, width)
        _topOffset = If(topOffset < 0, 0, topOffset)
        _followRef = followRef

        _form = New Form() With {
            .FormBorderStyle = FormBorderStyle.SizableToolWindow,
            .ShowInTaskbar = False,
            .StartPosition = FormStartPosition.Manual,
            .Text = title,
            .Width = _width,
            .Height = 600,
            .BackColor = Color.White
        }
        _form.MinimumSize = New Size(280, 200)
        _control.Dock = DockStyle.Fill
        _form.Controls.Add(_control)
        AddHandler _form.ResizeEnd, AddressOf OnFormResizeEnd

        ' ✕ 关闭按钮 = 仅隐藏（不销毁）。否则窗体被 Dispose 后，
        ' 面板管理器仍持有旧实例，再点 Ribbon"智能聊天"将无法打开。
        AddHandler _form.FormClosing,
            Sub(s, e)
                Log("FormClosing reason=" & e.CloseReason.ToString())
                If e.CloseReason = CloseReason.UserClosing Then
                    e.Cancel = True
                    _visible = False
                    _form.Hide()
                End If
            End Sub

        _syncTimer = New Timer() With {.Interval = 400}
        AddHandler _syncTimer.Tick, AddressOf OnSyncTick
        _syncTimer.Start()
    End Sub

    Private Sub OnSyncTick(sender As Object, e As EventArgs)
        Try
            If _form Is Nothing OrElse _form.IsDisposed Then Return
            If Not _visible Then Return
            ' owner 最小化时跟随隐藏（无论是否吸附）
            If _ownerHwnd <> IntPtr.Zero AndAlso Win32.IsIconic(_ownerHwnd) Then
                If _form.Visible Then _form.Hide()
                Return
            End If
            ' owner 恢复：若逻辑上应可见则重新显示（不强制位置/尺寸）
            If _ownerHwnd <> IntPtr.Zero AndAlso Not Win32.IsIconic(_ownerHwnd) Then
                If Not _form.Visible Then Win32.ShowWindow(_form.Handle, Win32.SW_SHOW)
            End If
            ' 已取消吸附：位置/尺寸完全交给用户，不再强制同步
            If Not _followRef Then Return
            If _refHwnd <> IntPtr.Zero AndAlso Not Win32.IsWindow(_refHwnd) Then
                _visible = False
                _form.Hide()
                Return
            End If
            LayoutToRef()
        Catch
        End Try
    End Sub

    ''' <summary>用户拖拽结束时记住自定义尺寸；此后同步只跟随位置，不覆盖尺寸</summary>
    Private Sub OnFormResizeEnd(sender As Object, e As EventArgs)
        Try
            If _form Is Nothing OrElse _form.IsDisposed Then Return
            _userWidth = _form.Width
            _userHeight = _form.Height
        Catch
        End Try
    End Sub

    ''' <summary>按参考窗口（屏幕矩形）把面板贴到其右缘；默认高度与内容区完全对齐</summary>
    Private Sub LayoutToRef()
        If _refHwnd = IntPtr.Zero Then
            Dim wa = Screen.PrimaryScreen.WorkingArea
            Dim dw = If(_userWidth > 0, _userWidth, _width)
            Dim dh = If(_userHeight > 0, _userHeight, wa.Height)
            _form.SetBounds(wa.Right - dw, wa.Top, dw, dh)
            Return
        End If
        Dim r As Win32.RECT = Nothing
        Win32.GetWindowRect(_refHwnd, r)
        Dim refWidth As Integer = r.Right - r.Left
        Dim refHeight As Integer = r.Bottom - r.Top
        If refWidth <= 0 OrElse refHeight <= 0 Then Return

        Dim w As Integer
        If _userWidth > 0 Then
            w = _userWidth
        Else
            w = Math.Min(_width, Math.Max(240, CInt(refWidth * 0.4)))
        End If

        Dim top As Integer
        Dim h As Integer
        If _userHeight > 0 Then
            ' 手动调过高度：底边对齐内容区底部
            top = Math.Max(r.Top, r.Bottom - _userHeight)
            h = _userHeight
        Else
            ' 默认：顶部对齐内容区顶部、底部对齐内容区底部（完全等高）
            top = r.Top + Math.Min(_topOffset, Math.Max(0, refHeight - 120))
            h = Math.Max(120, refHeight - top)
        End If

        Win32.SetWindowPos(_form.Handle, IntPtr.Zero, r.Right - w, top, w, h,
                           Win32.SWP_NOZORDER Or Win32.SWP_NOACTIVATE)
    End Sub

    Public ReadOnly Property Control As UserControl Implements IPaneHost.Control
        Get
            Return _control
        End Get
    End Property

    Public Property Visible As Boolean Implements IPaneHost.Visible
        Get
            Return _visible
        End Get
        Set(value As Boolean)
            Try
                If _form Is Nothing OrElse _form.IsDisposed Then
                    Log("Visible set " & value & " IGNORED (form disposed)")
                    Return
                End If
                Log("Visible set: " & value & " (current visible=" & _form.Visible & ", shownOnce=" & _shownOnce & ")")
                _visible = value
                If value Then
                    ' 首次显示时给一个默认位置（内容区右缘）
                    If Not _shownOnce Then
                        LayoutToRef()
                    End If
                    ' 统一用 Show(owner)：幂等且可靠；owned 关系随 PPT 最小化、总在其上
                    If _ownerHwnd <> IntPtr.Zero Then
                        _form.Show(New Win32WindowWrapper(_ownerHwnd))
                    Else
                        _form.Show()
                    End If
                    _shownOnce = True
                    ' 仅吸附模式下每次显示都对齐；自由模式尊重用户上次的位置/尺寸
                    If _followRef Then LayoutToRef()
                    Log("shown, bounds=" & _form.Bounds.ToString())
                Else
                    _form.Hide()
                    Log("hidden")
                End If
            Catch ex As Exception
                Log("Visible set FAILED: " & ex.ToString())
            End Try
        End Set
    End Property

    Public Property Width As Integer Implements IPaneHost.Width
        Get
            Return _width
        End Get
        Set(value As Integer)
            If value > 0 Then
                _width = value
                If _visible Then LayoutToRef()
            End If
        End Set
    End Property

    Public Property Title As String Implements IPaneHost.Title
        Get
            Return _title
        End Get
        Set(value As String)
            _title = value
            If _form IsNot Nothing AndAlso Not _form.IsDisposed Then _form.Text = value
        End Set
    End Property

    Public Sub Dispose() Implements IPaneHost.Dispose
        Try
            If _syncTimer IsNot Nothing Then
                _syncTimer.Stop()
                _syncTimer.Dispose()
            End If
            If _form IsNot Nothing Then
                If Not _form.IsDisposed Then
                    Win32.ShowWindow(_form.Handle, Win32.SW_HIDE)
                    _form.Dispose()
                End If
            End If
        Catch
        End Try
        _form = Nothing
    End Sub
End Class
