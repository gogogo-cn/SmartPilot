Option Explicit On
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' TaskSchedulerBridge - 定时任务规划引擎
    ''' 管理定时任务的创建、持久化、定时触发和执行。
    ''' 支持外部 HTTP API 和内部 COM 调用两种访问方式。
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.TaskSchedulerBridge")>
    Public Class TaskSchedulerBridge
        Implements IBridge

        ' 单例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of TaskSchedulerBridge)(
            Function() New TaskSchedulerBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 独立任务规划服务（Standalone Agent）基址，端口与 ConfigConstants.LocalServices.STANDALONE_AGENT_PORT 对齐
        Private Shared ReadOnly SAR_BASE As String = "http://localhost:" & Utils.ConfigConstants.LocalServices.STANDALONE_AGENT_PORT.ToString()

        Public Shared ReadOnly Property Instance As TaskSchedulerBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "TaskSchedulerBridge"
            End Get
        End Property

        ' 任务列表
        Private _tasks As New List(Of ScheduledTask)
        Private _taskLock As New Object

        ' 定时器管理
        Private _timers As New Dictionary(Of String, Timer)
        Private _timerLock As New Object

        ' 持久化路径
        Private _configPath As String = ""

        ' 日志
        Private _logger As ChatLogger

        ' 是否已初始化
        Private _initialized As Boolean = False

        Private Sub New()
            _logger = ChatLogger.Instance
            _configPath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "Config", "scheduled_tasks.json")
        End Sub

        ''' <summary>
        ''' 初始化：加载持久化任务并恢复定时器
        ''' </summary>
        Public Sub Initialize()
            If _initialized Then Return
            _initialized = True
            LoadTasks()
            RestoreAllTimers()
            _logger?.LogInfo("TASK-SCHEDULER", "TaskSchedulerBridge initialized, " & _tasks.Count & " tasks loaded")
        End Sub

        ''' <summary>
        ''' 获取所有任务（JSON 格式）
        ''' </summary>
        Public Function GetAllTasks() As String
            SyncLock _taskLock
                Return JsonConvert.SerializeObject(_tasks, Formatting.Indented)
            End SyncLock
        End Function

        '''<summary>'''
        ''' 推送单条任务（创建/更新合并）到独立服务 9529，异步 fire-and-forget 不阻塞 UI'''
        '''</summary>'''
        Private Shared Sub PushTaskToStandalone(task As ScheduledTask)
            If task Is Nothing OrElse String.IsNullOrEmpty(task.Id) Then Return
            Try
                Dim payload As String = JsonConvert.SerializeObject(task)
                SubTaskFireForget("POST", SAR_BASE & "/tasks", payload)
            Catch : End Try
        End Sub

        '''<summary>'''
        ''' 从独立服务 9529 删除任务 ID'''
        '''</summary>'''
        Private Shared Sub DeleteTaskFromStandalone(taskId As String)
            If String.IsNullOrEmpty(taskId) Then Return
            Try
                SubTaskFireForget("DELETE", SAR_BASE & "/tasks/" & Uri.EscapeDataString(taskId), Nothing)
            Catch : End Try
        End Sub

        '''<summary>'''
        ''' 全量同步所有任务到独立服务（独立服务刚启动时调用一次，确保数据对齐）'''
        '''</summary>'''
        Public Shared Sub SyncAllToStandaloneAgent()
            Try
                Dim all = Instance.GetTaskList()
                For Each t In all
                    PushTaskToStandalone(t)
                Next
            Catch : End Try
        End Sub

        '''<summary>'''
        ''' Fire-and-forget HTTP 辅助（用后台线程避免阻塞主线程）'''
        '''</summary>'''
        Private Shared Sub SubTaskFireForget(method As String, url As String, body As String)
            Dim t As New Threading.Thread(
                Sub()
                    Try
                        Dim req As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(url), Net.HttpWebRequest)
                        req.Method = method
                        req.Timeout = 3000
                        req.ReadWriteTimeout = 3000
                        req.ContentType = "application/json; charset=utf-8"
                        If Not String.IsNullOrEmpty(body) Then
                            Dim bytes = Encoding.UTF8.GetBytes(body)
                            req.ContentLength = bytes.Length
                            Using s = req.GetRequestStream()
                                s.Write(bytes, 0, bytes.Length)
                            End Using
                        Else
                            req.ContentLength = 0
                        End If
                        Using resp = DirectCast(req.GetResponse(), Net.HttpWebResponse)
                            Dim code = CInt(resp.StatusCode)
                            System.Diagnostics.Debug.WriteLine("StandaloneSync " & method & " " & url & " -> " & code)
                        End Using
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("StandaloneSync FAIL " & method & " " & url & ": " & ex.Message)
                    End Try
                End Sub
            )
            t.IsBackground = True
            t.Start()
        End Sub

        ''' <summary>
        ''' 同步 HTTP 调用（用于 ExecuteTask：需要真实结果，不能 fire-and-forget）。
        ''' 超时可配；默认 180s 足够跑一次 LLM。
        ''' </summary>
        Private Shared Function SubTaskCall(method As String, url As String, body As String, Optional timeoutMs As Integer = 180000) As String
            Try
                Dim req As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(url), Net.HttpWebRequest)
                req.Method = method
                req.Timeout = timeoutMs
                req.ReadWriteTimeout = timeoutMs
                req.ContentType = "application/json; charset=utf-8"
                If Not String.IsNullOrEmpty(body) Then
                    Dim bytes = Encoding.UTF8.GetBytes(body)
                    req.ContentLength = bytes.Length
                    Using s = req.GetRequestStream()
                        s.Write(bytes, 0, bytes.Length)
                    End Using
                Else
                    req.ContentLength = 0
                End If
                Using resp = DirectCast(req.GetResponse(), Net.HttpWebResponse)
                    Using sr = New StreamReader(resp.GetResponseStream(), Encoding.UTF8)
                        Return sr.ReadToEnd()
                    End Using
                End Using
            Catch ex As Net.WebException
                Try
                    If ex.Response IsNot Nothing Then
                        Using sr = New StreamReader(ex.Response.GetResponseStream(), Encoding.UTF8)
                            Return sr.ReadToEnd()
                        End Using
                    End If
                Catch : End Try
                Return Nothing
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("StandaloneCall FAIL " & method & " " & url & ": " & ex.Message)
                Return Nothing
            End Try
        End Function


        ''' <summary>
        ''' 获取所有任务（对象列表）
        ''' </summary>
        Public Function GetTaskList() As List(Of ScheduledTask)
            SyncLock _taskLock
                Return New List(Of ScheduledTask)(_tasks)
            End SyncLock
        End Function

        ''' <summary>
        ''' 创建新任务
        ''' </summary>
        Public Function CreateTask(taskJson As String) As String
            Try
                Dim task = JsonConvert.DeserializeObject(Of ScheduledTask)(taskJson)
                If task Is Nothing Then
                    Return JsonError("任务数据解析失败")
                End If

                task.Id = "tsk_" & Guid.NewGuid().ToString("N").Substring(0, 12)
                task.CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                task.NextRun = CalculateNextRun(task)

                SyncLock _taskLock
                    _tasks.Add(task)
                    SaveTasks()
                End SyncLock

                StartTimer(task)
                PushTaskToStandalone(task)
                _logger?.LogInfo("TASK-SCHEDULER", "Task created: " & task.Name & " [" & task.Id & "]")
                Return JsonSuccess(task)
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "CreateTask failed", ex)
                Return JsonError("创建任务失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 更新任务
        ''' </summary>
        Public Function UpdateTask(taskId As String, taskJson As String) As String
            Try
                Dim updates = JsonConvert.DeserializeObject(Of ScheduledTask)(taskJson)
                If updates Is Nothing Then
                    Return JsonError("任务数据解析失败")
                End If

                SyncLock _taskLock
                    Dim idx = _tasks.FindIndex(Function(t) t.Id = taskId)
                    If idx < 0 Then
                        Return JsonError("任务不存在: " & taskId)
                    End If

                    Dim task = _tasks(idx)
                    If updates.Name IsNot Nothing Then task.Name = updates.Name
                    If updates.Frequency IsNot Nothing Then task.Frequency = updates.Frequency
                    If updates.Time IsNot Nothing Then task.Time = updates.Time
                    If updates.Cron IsNot Nothing Then task.Cron = updates.Cron
                    If updates.Command IsNot Nothing Then task.Command = updates.Command
                    If updates.TargetFile IsNot Nothing Then task.TargetFile = updates.TargetFile
                    If updates.Model IsNot Nothing Then task.Model = updates.Model
                    task.Enabled = updates.Enabled
                    PushTaskToStandalone(task)
                    task.NextRun = CalculateNextRun(task)
                    SaveTasks()
                End SyncLock

                ' 重启定时器
                RestartTimer(taskId)
                _logger?.LogInfo("TASK-SCHEDULER", "Task updated: " & taskId)
                Return GetTaskById(taskId)
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "UpdateTask failed", ex)
                Return JsonError("更新任务失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 删除任务
        ''' </summary>
        Public Function DeleteTask(taskId As String) As String
            Try
                StopTimer(taskId)
                SyncLock _taskLock
                    Dim removed = _tasks.RemoveAll(Function(t) t.Id = taskId)
                    If removed = 0 Then
                        Return JsonError("任务不存在: " & taskId)
                    End If
                    SaveTasks()
                    DeleteTaskFromStandalone(taskId)
                End SyncLock
                _logger?.LogInfo("TASK-SCHEDULER", "Task deleted: " & taskId)
                Return JsonSuccess(Nothing, "任务已删除")
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "DeleteTask failed", ex)
                Return JsonError("删除任务失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 立即执行任务（真执行：调用独立服务 9529 的 /tasks/{id}/execute）。
        ''' 
        ''' 之前的实现只改 LastRun + 广播消息就回 success，前端显示"已触发执行"
        ''' 但实际上没跑任何指令 → 典型「假运行」。现在改为：
        '''   1. 先把最新任务（最新 TargetFile/Command/Model）POST 给 9529 upsert，
        '''      保证执行器侧用的是卡片上那版数据，不脏；
        '''   2. 再同步 POST /tasks/{id}/execute 到 9529（180s 超时，足够 LLM 跑完）；
        '''   3. 解析 9529 返回的 {success, message}，写入 ScheduledTask.LastResult，
        '''      并把 lastResult/success/message 一起回传给前端 → toast/卡片显示一致、
        '''      不再"toast 成功但实际文件/卡片都没变化"。
        ''' 
        ''' 若独立服务未启动，直接失败并返回明确信息（不再默默 success）。
        ''' </summary>
        Public Function ExecuteTask(taskId As String) As String
            Try
                Dim task As ScheduledTask = Nothing
                SyncLock _taskLock
                    task = _tasks.FirstOrDefault(Function(t) t.Id = taskId)
                End SyncLock

                If task Is Nothing Then
                    Return JsonError("任务不存在: " & taskId)
                End If

                If String.IsNullOrWhiteSpace(task.Command) Then
                    Return JsonError("任务缺少指令内容，无法执行")
                End If

                _logger?.LogInfo("TASK-SCHEDULER", "ExecuteTask (real) start: " & task.Name & " [" & taskId & "] file=" & If(task.TargetFile, ""))

                ' 先 upsert 给独立服务（前端已做，后端再保一次）
                Try
                    Dim upsertPayload = JsonConvert.SerializeObject(task)
                    SubTaskFireForget("POST", SAR_BASE & "/tasks", upsertPayload)
                    ' fire-and-forget 发送即忘，给 300ms 让它尽量到达（不阻塞执行）
                    Threading.Thread.Sleep(300)
                Catch : End Try

                ' 调独立服务执行。失败响应（HTTP 4xx/5xx）WebException 分支也会读响应体。
                Dim executeUrl = SAR_BASE & "/tasks/" & Uri.EscapeDataString(taskId) & "/execute"
                Dim rawResult = SubTaskCall("POST", executeUrl, "{}", 180000)

                Dim runSuccess As Boolean = False
                Dim runMessage As String = ""
                Dim runResult As JObject = Nothing
                If Not String.IsNullOrWhiteSpace(rawResult) Then
                    Try
                        runResult = JsonConvert.DeserializeObject(Of JObject)(rawResult)
                    Catch : End Try
                    If runResult IsNot Nothing Then
                        Dim successToken = runResult("success")
                        If successToken IsNot Nothing AndAlso successToken.Type = JTokenType.Boolean Then
                            runSuccess = successToken.Value(Of Boolean)()
                        Else
                            ' success 字段缺失、null 或非布尔：兼容只有 message 的纯失败响应，按失败记
                            runSuccess = False
                        End If
                        runMessage = If(runResult("message") IsNot Nothing, CStr(runResult("message")), "")
                    Else
                        ' 非 JSON 的纯文本（异常兜底）
                        runMessage = rawResult
                    End If
                End If

                ' 独立服务不通？明确报错，不再假成功
                If String.IsNullOrWhiteSpace(rawResult) Then
                    runSuccess = False
                    runMessage = "独立执行服务(localhost:9529)未响应，请启动独立服务后再执行"
                    _logger?.LogWarning("TASK-SCHEDULER", "ExecuteTask abort: 9529 no response for task " & taskId)
                End If

                ' 通过 WebView 通知前端（保持旧广播用于日志），但广播不再等同于"执行成功"
                SendTaskExecutionToFrontend(task)

                ' 独立服务返回的逐步结果明细（LLM 计划 + 各操作 + 保存），回读保存到任务
                Dim stepsToken = If(runResult IsNot Nothing, runResult("steps"), Nothing)
                If stepsToken IsNot Nothing AndAlso stepsToken.Type = JTokenType.Array Then
                    task.LastSteps = stepsToken.ToObject(Of List(Of Object))()
                End If

                ' 更新最后执行时间 + 下次运行时间
                SyncLock _taskLock
                    task.LastRun = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                    task.NextRun = CalculateNextRun(task)
                    task.LastResult = If(runSuccess,
                        If(String.IsNullOrWhiteSpace(runMessage), "✅ 执行成功", "✅ 执行成功：" & runMessage),
                        If(String.IsNullOrWhiteSpace(runMessage), "❌ 执行失败", "❌ 执行失败：" & runMessage)
                    )
                    SaveTasks()
                End SyncLock

                ' 也把结果异步推到 9529（让独立服务自己存的那份也一致）——fire-and-forget
                Try
                    Dim syncPayload = JsonConvert.SerializeObject(task)
                    SubTaskFireForget("POST", SAR_BASE & "/tasks", syncPayload)
                Catch : End Try

                ' 返回结构与前端 executeTask 期望一致：{success, message, lastResult, lastRun, nextRun, steps}
                Dim resp = New With {
                    .success = runSuccess,
                    .message = runMessage,
                    .lastResult = task.LastResult,
                    .lastRun = task.LastRun,
                    .nextRun = task.NextRun,
                    .steps = task.LastSteps,
                    .data = New With {
                        .id = task.Id,
                        .targetFile = task.TargetFile,
                        .command = task.Command
                    }
                }
                Return JsonConvert.SerializeObject(resp, Formatting.Indented)
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "ExecuteTask failed", ex)
                Return JsonError("执行任务失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 获取单个任务
        ''' </summary>
        Private Function GetTaskById(taskId As String) As String
            SyncLock _taskLock
                Dim task = _tasks.FirstOrDefault(Function(t) t.Id = taskId)
                If task Is Nothing Then
                    Return JsonError("任务不存在: " & taskId)
                End If
                Return JsonConvert.SerializeObject(task, Formatting.Indented)
            End SyncLock
        End Function

        ''' <summary>
        ''' 停止所有定时器（关闭时调用）
        ''' </summary>
        Public Sub StopAllTimers()
            SyncLock _timerLock
                For Each kvp In _timers
                    Try
                        kvp.Value.Stop()
                        kvp.Value.Dispose()
                    Catch ex As Exception  ' 
                    End Try
                Next
                _timers.Clear()
            End SyncLock
            _logger?.LogInfo("TASK-SCHEDULER", "All timers stopped")
        End Sub

        ''' <summary>
        ''' 恢复所有已启用的任务的定时器
        ''' </summary>
        Private Sub RestoreAllTimers()
            SyncLock _taskLock
                For Each task In _tasks.Where(Function(t) t.Enabled)
                    StartTimer(task)
                Next
            End SyncLock
        End Sub

        ''' <summary>
        ''' 启动一个任务的定时器
        ''' </summary>
        Private Sub StartTimer(task As ScheduledTask)
            If Not task.Enabled Then Return

            Dim intervalMs = CalculateIntervalMs(task)
            If intervalMs <= 0 Then Return

            SyncLock _timerLock
                ' 停止已存在的定时器
                If _timers.ContainsKey(task.Id) Then
                    Try
                        _timers(task.Id).Stop()
                        _timers(task.Id).Dispose()
                    Catch ex As Exception  ' 
                    End Try
                    _timers.Remove(task.Id)
                End If

                Dim timer As New Timer()
                timer.Interval = intervalMs
                AddHandler timer.Tick, Sub()
                    ' 定时器触发
                    timer.Stop()
                    Dim currentTask As ScheduledTask = Nothing
                    SyncLock _taskLock
                        currentTask = _tasks.FirstOrDefault(Function(t) t.Id = task.Id)
                    End SyncLock
                    If currentTask IsNot Nothing AndAlso currentTask.Enabled Then
                        ExecuteTask(currentTask.Id)
                        ' 重新计算下次执行时间并重启定时器
                        SyncLock _taskLock
                            currentTask.NextRun = CalculateNextRun(currentTask)
                            SaveTasks()
                        End SyncLock
                        ' 如果是单次任务，不再重启
                        If currentTask.Frequency <> "once" Then
                            StartTimer(currentTask)
                        End If
                    End If
                End Sub
                timer.Start()
                _timers(task.Id) = timer
            End SyncLock
        End Sub

        ''' <summary>
        ''' 停止一个任务的定时器
        ''' </summary>
        Private Sub StopTimer(taskId As String)
            SyncLock _timerLock
                If _timers.ContainsKey(taskId) Then
                    Try
                        _timers(taskId).Stop()
                        _timers(taskId).Dispose()
                    Catch ex As Exception  ' 
                    End Try
                    _timers.Remove(taskId)
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 重启一个任务的定时器
        ''' </summary>
        Private Sub RestartTimer(taskId As String)
            Dim task As ScheduledTask = Nothing
            SyncLock _taskLock
                task = _tasks.FirstOrDefault(Function(t) t.Id = taskId)
            End SyncLock
            If task Is Nothing Then Return
            StopTimer(taskId)
            If task.Enabled Then
                StartTimer(task)
            End If
        End Sub

        ''' <summary>
        ''' 计算定时器间隔（毫秒）
        ''' 对于每天/工作日/每周任务，计算到下次执行时间的毫秒数
        ''' 对于每小时任务，使用 1 小时间隔
        ''' </summary>
        Private Function CalculateIntervalMs(task As ScheduledTask) As Integer
            If task.Frequency = "hourly" Then
                Return 3600000 ' 1 小时
            End If

            Dim now = DateTime.Now
            Dim nextTime = ParseNextRunTime(task, now)
            If nextTime <= now Then
                nextTime = nextTime.AddDays(1)
            End If

            Dim interval = CInt((nextTime - now).TotalMilliseconds)
            If interval < 60000 Then ' 最少 1 分钟
                interval = 60000
            End If
            Return interval
        End Function

        ''' <summary>
        ''' 解析下次执行时间
        ''' </summary>
        Private Function ParseNextRunTime(task As ScheduledTask, fromTime As DateTime) As DateTime
            Dim timeParts = task.Time.Split(":"c)
            Dim hour = 0
            Dim minute = 0
            If timeParts.Length >= 2 Then
                Integer.TryParse(timeParts(0), hour)
                Integer.TryParse(timeParts(1), minute)
            End If

            Dim result = New DateTime(fromTime.Year, fromTime.Month, fromTime.Day, hour, minute, 0)

            If task.Frequency = "daily" OrElse task.Frequency = "once" Then
                Return result
            ElseIf task.Frequency = "weekdays" Then
                While result.DayOfWeek = DayOfWeek.Saturday OrElse result.DayOfWeek = DayOfWeek.Sunday
                    result = result.AddDays(1)
                End While
                Return result
            ElseIf task.Frequency = "weekly" Then
                ' 默认周一
                While result.DayOfWeek <> DayOfWeek.Monday
                    result = result.AddDays(1)
                End While
                Return result
            End If

            Return result
        End Function

        ''' <summary>
        ''' 计算下次执行时间字符串
        ''' </summary>
        Private Function CalculateNextRun(task As ScheduledTask) As String
            If Not task.Enabled Then Return ""
            If task.Frequency = "once" AndAlso task.LastRun IsNot Nothing AndAlso task.LastRun <> "" Then
                Return "" ' 已执行过的单次任务
            End If

            Dim now = DateTime.Now
            Dim nextTime = ParseNextRunTime(task, now)
            If nextTime <= now Then
                nextTime = nextTime.AddDays(1)
            End If
            Return nextTime.ToString("yyyy-MM-dd HH:mm:ss")
        End Function

        ''' <summary>
        ''' 发送任务执行消息到前端 WebView
        ''' </summary>
        Private Sub SendTaskExecutionToFrontend(task As ScheduledTask)
            Try
                Dim msg = New With {
                    .Type = "task_execution",
                    .taskId = task.Id,
                    .taskName = task.Name,
                    .command = task.Command
                }
                Dim json = JsonConvert.SerializeObject(msg)
                ' 通过 ThisAddIn 发送消息到 WebView
                Dim addin = Globals.ThisAddIn
                If addin IsNot Nothing Then
                    addin.SendMessageToWebView(json)
                End If
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "SendTaskExecutionToFrontend failed", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 从文件加载任务
        ''' </summary>
        Private Sub LoadTasks()
            Try
                If File.Exists(_configPath) Then
                    Dim json = File.ReadAllText(_configPath, Encoding.UTF8)
                    If Not String.IsNullOrWhiteSpace(json) Then
                        SyncLock _taskLock
                            _tasks = If(JsonConvert.DeserializeObject(Of List(Of ScheduledTask))(json), New List(Of ScheduledTask)())
                        End SyncLock
                    End If
                End If
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "LoadTasks failed", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 保存任务到文件
        ''' </summary>
        Private Sub SaveTasks()
            Try
                Dim dir = Path.GetDirectoryName(_configPath)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
                Dim json = JsonConvert.SerializeObject(_tasks, Formatting.Indented)
                File.WriteAllText(_configPath, json, Encoding.UTF8)
            Catch ex As Exception
                _logger?.LogError("TASK-SCHEDULER", "SaveTasks failed", ex)
            End Try
        End Sub

        ''' <summary>
        ''' 构造成功 JSON 响应
        ''' </summary>
        Private Function JsonSuccess(data As Object, Optional message As String = "") As String
            Dim result As New Dictionary(Of String, Object)
            result("success") = True
            If data IsNot Nothing Then result("data") = data
            If message <> "" Then result("message") = message
            Return JsonConvert.SerializeObject(result)
        End Function

        ''' <summary>
        ''' 构造错误 JSON 响应
        ''' </summary>
        Private Function JsonError(message As String) As String
            Dim result As New Dictionary(Of String, Object)
            result("success") = False
            result("message") = message
            Return JsonConvert.SerializeObject(result)
        End Function

        ''' <summary>
        ''' 释放资源
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            StopAllTimers()
            _initialized = False
            _logger?.LogInfo("TASK-SCHEDULER", "TaskSchedulerBridge disposed")
        End Sub
    End Class

    ''' <summary>
    ''' 定时任务数据类
    ''' </summary>
    Public Class ScheduledTask
        Public Property Id As String = ""
        Public Property Name As String = ""
        Public Property Frequency As String = "daily"   ' once | hourly | daily | weekdays | weekly | custom
        Public Property Time As String = "09:00"
        Public Property Cron As String = ""
        Public Property Command As String = ""
        Public Property TargetFile As String = ""
        ''' <summary>
        ''' 任务级执行模型（选填）。格式：provider::model；空串=跟随全局默认执行模型。
        ''' 例：ollama::qwen2.5:7b / deepseek::deepseek-reasoner / qianwen::qwen3-30b
        ''' </summary>
        Public Property Model As String = ""
        Public Property Enabled As Boolean = True
        Public Property CreatedAt As String = ""
        Public Property LastRun As String = ""
        Public Property NextRun As String = ""
        ''' <summary>
        ''' 最近一次执行结果（成功/失败摘要），供前端卡片与回包展示，避免"假运行"。
        ''' </summary>
        Public Property LastResult As String = ""
        ''' <summary>
        ''' 最近一次执行的逐步结果明细（LLM 计划 + 各操作 + 保存），供前端展开查看。
        ''' 每个元素：{op, target, desc, ok, detail?}
        ''' </summary>
        Public Property LastSteps As List(Of Object) = New List(Of Object)()
    End Class
End Namespace
