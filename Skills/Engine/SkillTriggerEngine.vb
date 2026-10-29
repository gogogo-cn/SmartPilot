Imports System
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Timers

Namespace Skills

    ''' <summary>
    ''' 【扩展3】Skill 触发引擎 - 管理事件驱动和定时触发的 Skill 执行
    '''
    ''' 功能：
    '''   1. 注册/注销 Excel Application 事件（SheetChange、WorkbookOpen 等）
    '''   2. 管理定时器触发（System.Threading.Timer）
    '''   3. 触发条件评估（FilterCondition 解析）
    '''   4. 防抖（Debounce）机制防止频繁触发
    '''   5. 异步执行触发的 Skill 并通知结果
    ''' </summary>
    Public Class SkillTriggerEngine

        ' ====== 单例模式 ======
        Private Shared ReadOnly _lazyInstance As New Lazy(Of SkillTriggerEngine)(
            Function() New SkillTriggerEngine(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As SkillTriggerEngine
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ' ====== 字段 ======
        Private _registeredTriggers As New Dictionary(Of String, TriggerRegistration)()
        Private _timerTriggers As New Dictionary(Of String, TimerTriggerState)()
        Private _debounceTimers As New Dictionary(Of String, DateTime)()
        Private _excelApp As Excel.Application
        Private _isInitialized As Boolean = False
        Private _lock As New Object()

        ' ====== 内部类型 ======
        Private Class TriggerRegistration
            Public Property SkillName As String
            Public Property SkillDefinition As SkillDefinition
            Public Property EventType As String
            Public Property FilterCondition As String
            Public Property DebounceMs As Integer
            Public Property Enabled As Boolean
        End Class

        Private Class TimerTriggerState
            Public Property SkillName As String
            Public Property SkillDefinition As SkillDefinition
            Public Property IntervalSeconds As Integer
            Public Property Timer As Timers.Timer
            Public Property Enabled As Boolean
        End Class

        ''' <summary>
        ''' 私有构造函数
        ''' </summary>
        Private Sub New()
        End Sub

        ''' <summary>
        ''' 初始化触发引擎，绑定 Excel Application 事件
        ''' 必须在 Excel Application 就绪后调用（通常在 ThisAddIn_Startup 中）
        ''' </summary>
        Public Sub Initialize(excelApp As Excel.Application)
            SyncLock _lock
                If _isInitialized Then Return

                _excelApp = excelApp

                Try
                    ' 绑定 SheetChange 事件
                    AddHandler DirectCast(_excelApp, Excel.AppEvents_Event).SheetChange,
                        AddressOf OnSheetChangeEvent

                    ' 绑定 WorkbookOpen 事件
                    AddHandler _excelApp.WorkbookOpen,
                        AddressOf OnWorkbookOpenEvent

                    ' 绑定 WorkbookBeforeClose 事件
                    AddHandler _excelApp.WorkbookBeforeClose,
                        AddressOf OnWorkbookBeforeCloseEvent

                    ' 绑定 SelectionChange 事件
                    AddHandler DirectCast(_excelApp, Excel.AppEvents_Event).SheetSelectionChange,
                        AddressOf OnSelectionChangeEvent

                    _isInitialized = True
                    Utils.AppLogger.LogInfo("SkillTriggerEngine", "初始化完成,Excel事件已绑定")

                Catch ex As Exception
                    Utils.AppLogger.LogError("SkillTriggerEngine", "初始化失败", ex)
                End Try
            End SyncLock
        End Sub

        ''' <summary>
        ''' 注册一个事件驱动的 Skill 触发器
        ''' </summary>
        Public Sub RegisterEventTrigger(skillName As String, skill As SkillDefinition)
            SyncLock _lock
                If skill.TriggerConfig Is Nothing OrElse
                   skill.TriggerConfig.TriggerType <> TriggerTypeEnum.[Event] Then Return

                Dim reg As New TriggerRegistration With {
                    .SkillName = skillName,
                    .SkillDefinition = skill,
                    .EventType = skill.TriggerConfig.EventType,
                    .FilterCondition = skill.TriggerConfig.FilterCondition,
                    .DebounceMs = skill.TriggerConfig.DebounceMs,
                    .Enabled = skill.TriggerConfig.Enabled
                }

                _registeredTriggers(skillName) = reg
                System.Diagnostics.Debug.WriteLine(String.Format(
                    "[SkillTriggerEngine] 已注册事件触发: {0} (事件={1}, 条件={2})",
                    skillName, skill.TriggerConfig.EventType, skill.TriggerConfig.FilterCondition))
            End SyncLock
        End Sub

        ''' <summary>
        ''' 注册一个定时触发的 Skill 触发器
        ''' </summary>
        Public Sub RegisterTimerTrigger(skillName As String, skill As SkillDefinition)
            SyncLock _lock
                If skill.TriggerConfig Is Nothing OrElse
                   skill.TriggerConfig.TriggerType <> TriggerTypeEnum.Timer Then Return

                UnregisterTimerTrigger(skillName)

                Dim intervalMs = Math.Max(1000, skill.TriggerConfig.TimerIntervalSeconds * 1000)

                Dim timer = New System.Timers.Timer(intervalMs)
                AddHandler timer.Elapsed, AddressOf OnTimerElapsed

                Dim state As New TimerTriggerState With {
                    .SkillName = skillName,
                    .SkillDefinition = skill,
                    .IntervalSeconds = skill.TriggerConfig.TimerIntervalSeconds,
                    .Timer = timer,
                    .Enabled = skill.TriggerConfig.Enabled
                }

                _timerTriggers(skillName) = state

                If skill.TriggerConfig.Enabled Then
                    timer.Start()
                End If

                System.Diagnostics.Debug.WriteLine(String.Format(
                    "[SkillTriggerEngine] 已注册定时触发: {0} (间隔={1}s)",
                    skillName, skill.TriggerConfig.TimerIntervalSeconds))
            End SyncLock
        End Sub

        ''' <summary>
        ''' 定时器触发处理（安全包装，捕获所有异常）
        ''' </summary>
        Private Sub OnTimerElapsed(sender As Object, e As ElapsedEventArgs)
            Dim timer = TryCast(sender, System.Timers.Timer)
            If timer Is Nothing Then Return

            Dim skillName As String = Nothing
            SyncLock _lock
                For Each kvp In _timerTriggers
                    If kvp.Value.Timer Is timer Then
                        skillName = kvp.Key
                        Exit For
                    End If
                Next
            End SyncLock

            If String.IsNullOrEmpty(skillName) Then Return

            Task.Run(Async Function()
                         Try
                             Await ExecuteTimerTriggerAsync(skillName)
                         Catch ex As Exception
                             System.Diagnostics.Debug.WriteLine(String.Format(
                                 "[SkillTriggerEngine] 定时触发 {0} 异常: {1}",
                                 skillName, ex.Message))
                         End Try
                     End Function)
        End Sub

        ''' <summary>
        ''' 注销指定 Skill 的所有触发器
        ''' </summary>
        Public Sub UnregisterTrigger(skillName As String)
            SyncLock _lock
                If _registeredTriggers.ContainsKey(skillName) Then
                    _registeredTriggers.Remove(skillName)
                End If
                UnregisterTimerTrigger(skillName)
            End SyncLock
        End Sub

        ''' <summary>
        ''' 注销定时触发器
        ''' </summary>
        Private Sub UnregisterTimerTrigger(skillName As String)
            If _timerTriggers.ContainsKey(skillName) Then
                Dim state = _timerTriggers(skillName)
                If state.Timer IsNot Nothing Then
                    state.Timer.Stop()
                    state.Timer.Dispose()
                End If
                _timerTriggers.Remove(skillName)
            End If
        End Sub

        ''' <summary>
        ''' 停止所有触发器（在插件关闭时调用）
        ''' </summary>
        Public Sub Shutdown()
            SyncLock _lock
                Try
                    ' 停止所有定时器
                    For Each kvp In _timerTriggers.ToList()
                        UnregisterTimerTrigger(kvp.Key)
                    Next

                    ' 清空注册表
                    _registeredTriggers.Clear()
                    _debounceTimers.Clear()

                    ' 【修复】移除 Excel 事件处理器，防止关闭时事件回调到已卸载的插件导致卡死
                    If _excelApp IsNot Nothing Then
                        Try : RemoveHandler DirectCast(_excelApp, Excel.AppEvents_Event).SheetChange, AddressOf OnSheetChangeEvent : Catch : End Try
                        Try : RemoveHandler _excelApp.WorkbookOpen, AddressOf OnWorkbookOpenEvent : Catch : End Try
                        Try : RemoveHandler _excelApp.WorkbookBeforeClose, AddressOf OnWorkbookBeforeCloseEvent : Catch : End Try
                        Try : RemoveHandler DirectCast(_excelApp, Excel.AppEvents_Event).SheetSelectionChange, AddressOf OnSelectionChangeEvent : Catch : End Try
                        _excelApp = Nothing
                    End If

                    _isInitialized = False
                    System.Diagnostics.Debug.WriteLine("[SkillTriggerEngine] 已关闭")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[SkillTriggerEngine] 关闭异常: " & ex.Message)
                End Try
            End SyncLock
        End Sub

        ' ====== Excel 事件处理 ======

        Private Sub OnSheetChangeEvent(sh As Object, target As Excel.Range)
            HandleEventTrigger("Change", sh, target)
        End Sub

        Private Sub OnWorkbookOpenEvent(wb As Excel.Workbook)
            HandleEventTrigger("Open", wb, Nothing)
        End Sub

        Private Sub OnWorkbookBeforeCloseEvent(wb As Excel.Workbook, ByRef cancel As Boolean)
            HandleEventTrigger("Close", wb, Nothing)
        End Sub

        Private Sub OnSelectionChangeEvent(sh As Object, target As Excel.Range)
            HandleEventTrigger("SelectionChange", sh, target)
        End Sub

        ''' <summary>
        ''' 统一事件分发：遍历匹配的触发器并执行
        ''' </summary>
        Private Sub HandleEventTrigger(eventType As String, sourceObj As Object, targetRange As Excel.Range)
            SyncLock _lock
                For Each kvp In _registeredTriggers
                    Dim reg = kvp.Value
                    If Not reg.Enabled Then Continue For
                    If reg.EventType <> eventType Then Continue For

                    ' 防抖检查
                    Dim now = DateTime.Now
                    If _debounceTimers.ContainsKey(reg.SkillName) Then
                        Dim lastTime = _debounceTimers(reg.SkillName)
                        If (now - lastTime).TotalMilliseconds < reg.DebounceMs Then
                            Continue For ' 跳过（防抖中）
                        End If
                    End If
                    _debounceTimers(reg.SkillName) = now

                    ' 过滤条件检查
                    If Not String.IsNullOrWhiteSpace(reg.FilterCondition) Then
                        If Not EvaluateFilterCondition(reg.FilterCondition, sourceObj, targetRange) Then
                            Continue For
                        End If
                    End If

                    ' 异步执行 Skill
                    Task.Run(Async Function()
                        Try
                            Dim executor = New CustomSkillExecutor()
                            Dim inputs = New Dictionary(Of String, Object) From {
                                {"eventType", eventType},
                                {"source", "event_trigger"}
                            }
                            ' 如果有目标 Range 信息，传入上下文
                            If targetRange IsNot Nothing Then
                                inputs("targetAddress") = targetRange.Address
                                inputs("targetColumn") = GetColumnLetter(targetRange.Column)
                                inputs("targetRow") = targetRange.Row.ToString()
                                If targetRange.Worksheet IsNot Nothing Then
                                    inputs("targetSheet") = targetRange.Worksheet.Name
                                End If
                                If targetRange.Value IsNot Nothing Then
                                    inputs("cellValue") = targetRange.Value.ToString()
                                End If
                            End If

                            Dim result = Await executor.ExecuteCustomSkill(reg.SkillDefinition, inputs)

                            ' 通知前端结果
                            NotifyTriggerResult(reg.SkillName, result)
                        Catch ex As Exception
                            System.Diagnostics.Debug.WriteLine(String.Format(
                                "[SkillTriggerEngine] 触发执行异常 ({0}): {1}", reg.SkillName, ex.Message))
                        End Try
                    End Function)
                Next
            End SyncLock
        End Sub

        ''' <summary>
        ''' 定时触发执行
        ''' </summary>
        Private Async Function ExecuteTimerTriggerAsync(skillName As String) As Task
            If Not _timerTriggers.ContainsKey(skillName) Then Return

            Dim state = _timerTriggers(skillName)
            If Not state.Enabled Then Return

            Try
                Dim executor = New CustomSkillExecutor()
                Dim inputs = New Dictionary(Of String, Object) From {
                    {"triggerType", "timer"}
                }
                Dim result = Await executor.ExecuteCustomSkill(state.SkillDefinition, inputs)
                NotifyTriggerResult(skillName, result)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format(
                    "[SkillTriggerEngine] 定时触发异常 ({0}): {1}", skillName, ex.Message))
            End Try
        End Function

        ''' <summary>
        ''' 评估过滤条件（简化版）
        ''' 支持格式: targetColumn='B', targetSheet='订单表', cellValue!=''
        ''' </summary>
        Private Function EvaluateFilterCondition(condition As String, sourceObj As Object, targetRange As Excel.Range) As Boolean
            Try
                Dim ctx As New Dictionary(Of String, String)

                If targetRange IsNot Nothing Then
                    ctx("targetColumn") = GetColumnLetter(targetRange.Column)
                    ctx("targetRow") = targetRange.Row.ToString()
                    ctx("targetAddress") = targetRange.Address
                    If targetRange.Worksheet IsNot Nothing Then
                        ctx("targetSheet") = targetRange.Worksheet.Name
                    End If
                    If targetRange.Value IsNot Nothing Then
                        ctx("cellValue") = targetRange.Value.ToString()
                    Else
                        ctx("cellValue") = ""
                    End If
                End If

                ' 简单条件解析
                For Each op In {"=", "<>", "!=", "contains"}
                    If condition.Contains(op) Then
                        Dim pair = condition.Split({op}, 2, StringSplitOptions.None)
                        If pair.Length = 2 Then
                            Dim varName = pair(0).Trim().ToLower()
                            Dim expectedValue = pair(1).Trim().Trim("'"c, """"c).ToLower()
                            If ctx.ContainsKey(varName) Then
                                Dim actualValue = ctx(varName).ToLower()
                                Select Case op
                                    Case "=", "=="
                                        Return actualValue = expectedValue
                                    Case "<>", "!="
                                        Return actualValue <> expectedValue
                                    Case "contains"
                                        Return actualValue.Contains(expectedValue)
                                End Select
                            End If
                        End If
                    End If
                Next

                Return True ' 无法解析时默认通过
            Catch ex As Exception  ' 
                Return True
            End Try
        End Function

        ''' <summary>
        ''' 将列号转换为字母
        ''' </summary>
        Private Function GetColumnLetter(colNum As Integer) As String
            If colNum <= 0 Then Return ""
            Dim result = ""
            While colNum > 0
                colNum -= 1
                result = Chr(65 + (colNum Mod 26)) & result
                colNum \= 26
            End While
            Return result
        End Function

        ''' <summary>
        ''' 向前端 WebView 发送触发结果通知
        ''' </summary>
        Private Sub NotifyTriggerResult(skillName As String, result As WorkflowResult)
            Try
                Dim addIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    Dim notification = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                        .type = "triggerResult",
                        .skillName = skillName,
                        .success = result.Success,
                        .message = If(result.Success, result.Message, result.ErrorMessage),
                        .timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    })
                    addIn.SendMessageToWebView(notification)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[SkillTriggerEngine] 通知失败: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace
