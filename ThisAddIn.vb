Option Strict On
Option Infer On

Imports System
Imports System.IO
Imports System.Linq
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports System.Windows.Forms.Integration
Imports System.Drawing
Imports Microsoft.Office.Tools
Imports Microsoft.Office.Core
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAIVSTOAddIn
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Skills.Testing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Partial Public NotInheritable Class ThisAddIn

#Region "核心字段"

    Private _ribbon As Ribbon
    Private _excelBridge As ExcelAIVSTOAddIn.ExcelBridge
    Private _chatLogger As Utils.ChatLogger

    Private _webViewManager As WebView2Manager
    Private _taskPaneManager As TaskPaneManager
    Private _screenshotService As ScreenshotService

    ' 主聊天面板使用WebView2
    Private _chatTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private _chatHostControl As UserControl

#End Region

#Region "核心辅助方法"

    ''' <summary>P3：火与忘记异步安全包装（捕获Task异常+日志+解决BC42358警告）</summary>
    Private Shared Sub SafeFork(task As Task, Optional caller As String = NameOf(SafeFork), _
                                 Optional logger As Utils.ChatLogger = Nothing)
        If task Is Nothing Then Return
        Dim __discard = task.ContinueWith(AddressOf SafeForkContinuation, Tuple.Create(logger, caller), TaskScheduler.Default)
    End Sub

    ''' <summary>P3：SafeFork ContinueWith辅助（AddressOf单独方法避免Lambda换行解析坑）</summary>
    Private Shared Sub SafeForkContinuation(antecedent As Task, state As Object)
        Dim tuple = TryCast(state, Tuple(Of Utils.ChatLogger, String))
        If tuple Is Nothing OrElse Not antecedent.IsFaulted OrElse antecedent.Exception Is Nothing Then Return
        Try
            Dim ex = antecedent.Exception.InnerException
            Dim logger = tuple.Item1
            Dim caller = tuple.Item2
            Dim msg As String = If(ex IsNot Nothing, ex.Message, "null")
            logger?.LogError("ADDIN", caller & " 异步后台异常: " & msg, ex)
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] " & caller & " FAULT: " & If(ex IsNot Nothing, ex.ToString(), "null"))
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 在插件UI线程(Excel主线程)上执行操作
    ''' </summary>
    Public Shared Sub RunOnUIThread(action As Action)
        If Globals.ThisAddIn Is Nothing Then
            action()
            Return
        End If

        Dim addIn = Globals.ThisAddIn

        If addIn._chatHostControl IsNot Nothing AndAlso Not addIn._chatHostControl.IsDisposed Then
            Dim control = addIn._chatHostControl
            If control.InvokeRequired Then
                control.Invoke(action)
            Else
                action()
            End If
        Else
            Try
                ' Excel COM Application 对象没有 WinForms 风格的 Invoke 方法。
                ' 此分支通常已经在 Excel 主线程中执行，直接调用 action；
                ' 失败时记录日志并重新抛出，绝不重复执行（避免副作用跑两遍）。
                action()
            Catch ex As Exception
                Dim logger = addIn._chatLogger
                logger?.LogError("ADDIN", "RunOnUIThread 直接执行分支失败: " & ex.Message, ex)
                Throw
            End Try
        End If
    End Sub

#End Region

#Region "核心公共API"

    ''' <summary>
    ''' 获取Excel应用程序对象
    ''' </summary>
    Public Function GetExcelApp() As Excel.Application
        Return Me.Application
    End Function

    ''' <summary>
    ''' 向WebView2主面板发送消息
    ''' </summary>
    Public Sub SendMessageToWebView(message As String)
        Try
            If _webViewManager Is Nothing Then
                Dim maxLen As Integer = Math.Min(100, message.Length)
                _chatLogger?.LogInfo("ADDIN", "WebView2 main panel not initialized, message not delivered: " & message.Substring(0, maxLen))
                Return
            End If
            _webViewManager?.WebView?.CoreWebView2?.PostWebMessageAsString(message)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Send message to WebView failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 在WebView2主面板上执行JS脚本
    ''' </summary>
    Public Sub ExecuteScriptOnWebView(script As String)
        Try
            If _webViewManager Is Nothing OrElse _webViewManager.WebView?.CoreWebView2 Is Nothing Then
                _chatLogger?.LogInfo("ADDIN", "WebView2 not ready, script not executed")
                Return
            End If
            _webViewManager.WebView.CoreWebView2.ExecuteScriptAsync(script)
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Execute WebView script failed", ex)
        End Try
    End Sub

#End Region

End Class