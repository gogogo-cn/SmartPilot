' ============================================================================
' CustomTaskPaneHost - 使用 VSTO 原生 CustomTaskPane（Word / Excel 适用）
'
' 单独成文件：它依赖 Microsoft.Office.Tools；PowerPoint 工程（COM 加载项）
' 不引用 VSTO，只 Link 不依赖 Tools 的 OfficeHost.vb。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Windows.Forms
Imports Microsoft.Office.Tools

Public Class CustomTaskPaneHost
    Implements IPaneHost

    Private _pane As CustomTaskPane
    Private _control As UserControl

    Public Sub New(pane As CustomTaskPane)
        _pane = pane
        _control = TryCast(pane.Control, UserControl)
    End Sub

    Public ReadOnly Property Control As UserControl Implements IPaneHost.Control
        Get
            Return _control
        End Get
    End Property

    Public Property Visible As Boolean Implements IPaneHost.Visible
        Get
            Return _pane IsNot Nothing AndAlso _pane.Visible
        End Get
        Set(value As Boolean)
            If _pane IsNot Nothing Then _pane.Visible = value
        End Set
    End Property

    Public Property Width As Integer Implements IPaneHost.Width
        Get
            If _pane Is Nothing Then Return 0
            Return _pane.Width
        End Get
        Set(value As Integer)
            If _pane IsNot Nothing AndAlso value > 0 Then _pane.Width = value
        End Set
    End Property

    Public Property Title As String Implements IPaneHost.Title
        Get
            Return ""
        End Get
        Set(value As String)
            ' CustomTaskPane 标题创建后不可变
        End Set
    End Property

    Public Sub Dispose() Implements IPaneHost.Dispose
        Try
            If _pane IsNot Nothing Then _pane.Dispose()
        Catch
        End Try
        _pane = Nothing
    End Sub
End Class
