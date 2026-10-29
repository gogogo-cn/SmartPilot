Option Strict On
Option Infer On

Imports System
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Excel = Microsoft.Office.Interop.Excel

Partial Public NotInheritable Class ThisAddIn

#Region "Excel选择监听"

    Private _selectionMonitorActive As Boolean = False
    Private _selectionDebounceTimer As System.Windows.Forms.Timer = Nothing
    Private _pendingSelectionData As String = ""

    ''' <summary>
    ''' 开始监控Excel选择变化(参考区域添加模式)
    ''' </summary>
    Public Sub StartSelectionMonitor()
        Try
            _selectionMonitorActive = True
            _chatLogger?.LogInfo("ADDIN", "Start monitoring Excel selection changes")
            NotifyWebViewSelectionEvent("started")

            ' 激活Excel窗口,让用户直接在Excel中选择区域
            Try
                Dim excelApp As Excel.Application = GetExcelApp()
                If excelApp IsNot Nothing Then
                    ' 注意:COM Application.Activate() / Window.Activate() 在 Option Strict On 下
                    ' 即使声明了强类型 RCW,仍被视为 IDispatch 后期绑定,改用 CallByName 反射调用
                    Microsoft.VisualBasic.Interaction.CallByName(excelApp, "Activate", CallType.Method)
                    Dim winObj As Object = excelApp.ActiveWindow
                    If winObj IsNot Nothing Then
                        Dim win As Excel.Window = TryCast(winObj, Excel.Window)
                        If win IsNot Nothing Then
                            Microsoft.VisualBasic.Interaction.CallByName(win, "Activate", CallType.Method)
                            Utils.ComReleaseHelper.Release(win)
                        End If
                    End If
                    _chatLogger?.LogInfo("ADDIN", "Excel window activated")
                End If
            Catch ex As Exception
                _chatLogger?.LogError("ADDIN", "Activate Excel window failed", ex)
            End Try
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "StartSelectionMonitor exception", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 停止监控Excel选择变化
    ''' </summary>
    Public Sub StopSelectionMonitor()
        Try
            _selectionMonitorActive = False
            _pendingSelectionData = ""
            If _selectionDebounceTimer IsNot Nothing Then
                _selectionDebounceTimer.Stop()
            End If
            _chatLogger?.LogInfo("ADDIN", "Stop monitoring Excel selection changes")
            NotifyWebViewSelectionEvent("stopped")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "StopSelectionMonitor exception", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 通知前端选择监听器状态或变化
    ''' </summary>
    Public Sub NotifyWebViewSelectionEvent(eventType As String, Optional dataJson As String = "")
        Try
            Dim payloadObj As New Newtonsoft.Json.Linq.JObject()
            payloadObj("type") = eventType
            If Not String.IsNullOrEmpty(dataJson) Then
                payloadObj("data") = Newtonsoft.Json.Linq.JObject.Parse(dataJson)
            End If
            Dim payloadJson As String = payloadObj.ToString(Formatting.None)
            Dim jsStringLiteral As String = JsonConvert.ToString(payloadJson)
            ExecuteScriptOnWebView("if(window.__onExcelSelectionChanged){window.__onExcelSelectionChanged(JSON.parse(" & jsStringLiteral & "));}")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Notify frontend selection event failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 当Excel选择变化且侧边栏可见时,自动刷新上下文数据。
    ''' 同时处理参考区域添加模式的选择捕获。
    ''' </summary>
    Private Sub Application_SheetSelectionChange(Sh As Object, Target As Excel.Range)
        Try
            ' 参考区域添加模式:选择变化时通知前端
            If _selectionMonitorActive AndAlso Target IsNot Nothing Then
                Try
                    Dim sheetName As String = ""
                    Dim workbookName As String = ""
                    Try
                        Dim ws As Excel.Worksheet = CType(Sh, Excel.Worksheet)
                        sheetName = ws.Name
                        If ws.Parent IsNot Nothing Then
                            Dim wb As Excel.Workbook = TryCast(ws.Parent, Excel.Workbook)
                            If wb IsNot Nothing Then workbookName = wb.Name
                        End If
                    Catch
                        sheetName = ""
                    End Try

                    Dim targetText As String = ""
                    Try
                        If Target.Text IsNot Nothing Then targetText = Target.Text.ToString()
                    Catch
                    End Try

                    Dim data As String = JsonConvert.SerializeObject(New With {
                        .address = Target.Address.ToString(),
                        .sheetName = sheetName,
                        .workbookName = workbookName,
                        .text = targetText
                    })

                    ' 保存最新选择数据(拖拽期间多次触发,仅保留最后一次)
                    _pendingSelectionData = data

                    ' 启动/重置防抖计时器:300ms无新变化 = 最终选择
                    If _selectionDebounceTimer Is Nothing Then
                        _selectionDebounceTimer = New System.Windows.Forms.Timer()
                        _selectionDebounceTimer.Interval = 300
                        AddHandler _selectionDebounceTimer.Tick, AddressOf OnSelectionDebounceTick
                    End If
                    _selectionDebounceTimer.Stop()
                    _selectionDebounceTimer.Start()
                Catch ex As Exception
                    _chatLogger?.LogError("ADDIN", "Reference area selection change handling exception", ex)
                End Try
            End If
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Refresh sidebar data on selection change failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 防抖计时器回调:选择稳定300ms后触发最终捕获
    ''' </summary>
    Private Sub OnSelectionDebounceTick(sender As Object, e As EventArgs)
        Try
            If _selectionDebounceTimer IsNot Nothing Then
                _selectionDebounceTimer.Stop()
            End If

            If _selectionMonitorActive AndAlso Not String.IsNullOrEmpty(_pendingSelectionData) Then
                _chatLogger?.LogInfo("ADDIN", "Debounce capture selection (final): " & _pendingSelectionData)
                NotifyWebViewSelectionEvent("changed", _pendingSelectionData)
                StopSelectionMonitor()
                _pendingSelectionData = ""
            End If
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Debounce capture exception", ex)
        End Try
    End Sub

#End Region

End Class