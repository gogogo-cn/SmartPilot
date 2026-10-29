' ============================================================
' PptTaskQueue.vb — PPT 生成异步任务队列（P1-1）
'
' 目标：把 PptGenerationService.Generate 从「Excel 主线程同步执行」
' 改为「提交 taskId + 后台线程执行 + 前端轮询状态」，大 PPT 生成期间
' Excel UI 不再假死。
'
' 设计：
'   - 单 Worker 串行执行（队列语义，保持原 Generate 的"一次一个"并发约束）；
'   - Worker 通过 PptGenerationService.GenerateQueued 与同步 Generate 共用
'     同一把生成锁，两者互斥，避免静态状态（主题/画布/版式）被并发污染；
'   - 状态机：queued → running → done / error / cancelled；
'   - 仅允许取消排队中（queued）的任务；running 中的 PowerPoint COM 调用无法中断。
' ============================================================
Option Explicit On
Option Infer On
Option Strict On

Imports System
Imports System.Collections.Generic
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Services

    Public NotInheritable Class PptTaskQueue

        Private Const STATUS_QUEUED As String = "queued"
        Private Const STATUS_RUNNING As String = "running"
        Private Const STATUS_DONE As String = "done"
        Private Const STATUS_ERROR As String = "error"
        Private Const STATUS_CANCELLED As String = "cancelled"

        Private Shared ReadOnly _lock As New Object()
        Private Shared ReadOnly _tasks As New Dictionary(Of String, TaskEntry)()
        Private Shared ReadOnly _queue As New Queue(Of TaskEntry)()
        Private Shared _workerBusy As Boolean = False

        ''' <summary>任务条目（内存态，不持久化；崩溃后任务结果自然丢失）</summary>
        Private NotInheritable Class TaskEntry
            Public Property TaskId As String
            Public Property ArgsJson As String
            Public Property Status As String = STATUS_QUEUED
            Public Property Progress As Integer = 0
            Public Property Stage As String = "排队中"
            Public Property ResultJson As String = ""
            Public Property Error As String = ""
            Public Property CreatedAt As DateTime = DateTime.Now
            Public Property StartedAt As DateTime?
            Public Property FinishedAt As DateTime?
        End Class

        ''' <summary>
        ''' 提交 PPT 生成任务（立即返回，不阻塞调用线程）。
        ''' 返回：{success:true, result:{taskId}} 或 {success:false, error}
        ''' </summary>
        Public Shared Function Submit(argsJson As String) As String
            Try
                Dim entry As New TaskEntry()
                entry.TaskId = Guid.NewGuid().ToString("N").Substring(0, 16)
                entry.ArgsJson = If(argsJson, "")

                SyncLock _lock
                    _tasks(entry.TaskId) = entry
                    _queue.Enqueue(entry)
                    EnsureWorkerLocked()
                End SyncLock

                Return JsonConvert.SerializeObject(New With {
                    Key .success = True,
                    Key .result = New With {Key .taskId = entry.TaskId}
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    Key .success = False,
                    Key .error = "任务提交失败: " & ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 查询任务状态（前端轮询入口）。
        ''' 返回：{success:true, result:{taskId,status,progress,stage,output?,error?}}
        ''' output 仅在 status=done 时出现，值为 PptGenerationService.Generate 的原始返回 JSON 字符串，
        ''' 与同步 CallLocalTool 路径的 result 形状一致，前端可无缝复用既有解析逻辑。
        ''' </summary>
        Public Shared Function GetStatus(taskId As String) As String
            Dim entry As TaskEntry = Nothing
            SyncLock _lock
                If String.IsNullOrEmpty(taskId) OrElse Not _tasks.TryGetValue(taskId, entry) Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .success = False,
                        Key .error = "任务不存在: " & If(taskId, "(空)")
                    })
                End If

                ' 在锁内读取任务状态字段，保证与 Worker 的写入建立同步关系
                Dim result As New JObject()
                result("taskId") = entry.TaskId
                result("status") = entry.Status
                result("progress") = entry.Progress
                result("stage") = entry.Stage
                If entry.Status = STATUS_DONE Then
                    If Not String.IsNullOrEmpty(entry.ResultJson) Then
                        ' 保持与同步路径 data.result 一致的形状：Generate 原始返回 JSON 文本
                        result("output") = entry.ResultJson
                    End If
                ElseIf entry.Status = STATUS_ERROR OrElse entry.Status = STATUS_CANCELLED Then
                    result("error") = entry.Error
                End If
                Return JsonConvert.SerializeObject(New With {
                    Key .success = True,
                    Key .result = result
                })
            End SyncLock
        End Function

        ''' <summary>
        ''' 取消任务：仅对排队中（queued）的任务生效；已 running 的 COM 调用无法中断。
        ''' </summary>
        Public Shared Function Cancel(taskId As String) As String
            Dim entry As TaskEntry = Nothing
            SyncLock _lock
                If String.IsNullOrEmpty(taskId) OrElse Not _tasks.TryGetValue(taskId, entry) Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .success = False,
                        Key .error = "任务不存在: " & If(taskId, "(空)")
                    })
                End If
                If entry.Status = STATUS_QUEUED Then
                    entry.Status = STATUS_CANCELLED
                    entry.Error = "已取消"
                    entry.FinishedAt = DateTime.Now
                Else
                    Return JsonConvert.SerializeObject(New With {
                        Key .success = False,
                        Key .error = "仅能取消排队中的任务（当前状态: " & entry.Status & "）"
                    })
                End If
            End SyncLock
            Return JsonConvert.SerializeObject(New With {Key .success = True})
        End Function

        ''' <summary>
        ''' 清理已完成任务（供宿主关闭或定时清理调用）
        ''' </summary>
        Public Shared Sub CleanupCompleted()
            SyncLock _lock
                Dim expiredKeys As New List(Of String)()
                For Each kvp In _tasks
                    Dim st = kvp.Value.Status
                    Dim isFinished = (st = STATUS_DONE OrElse st = STATUS_ERROR OrElse st = STATUS_CANCELLED)
                    If isFinished AndAlso (DateTime.Now - kvp.Value.FinishedAt.GetValueOrDefault(DateTime.MinValue)).TotalMinutes > 30 Then
                        expiredKeys.Add(kvp.Key)
                    End If
                Next
                For Each k In expiredKeys
                    _tasks.Remove(k)
                Next
            End SyncLock
        End Sub

        ''' <summary>必须在持有 _lock 时调用</summary>
        Private Shared Sub EnsureWorkerLocked()
            If _workerBusy Then Return
            _workerBusy = True
            Dim discard = Task.Run(Sub() WorkerLoop())
        End Sub

        ''' <summary>串行 Worker：一次执行一个任务，直到队列清空</summary>
        Private Shared Sub WorkerLoop()
            Try
                While True
                    Dim entry As TaskEntry = Nothing
                    SyncLock _lock
                        While _queue.Count > 0 AndAlso _queue.Peek().Status = STATUS_CANCELLED
                            _queue.Dequeue()
                        End While
                        If _queue.Count = 0 Then
                            _workerBusy = False
                            Exit While
                        End If
                        entry = _queue.Dequeue()
                        entry.Status = STATUS_RUNNING
                        entry.StartedAt = DateTime.Now
                        entry.Progress = 0
                        entry.Stage = "准备中"
                    End SyncLock

                    If entry Is Nothing Then Continue While

                    ' 与同步 Generate 共用生成锁（阻塞等待），保证任意时刻只有一个 PPT 生成执行
                    Dim resultJson As String = ""
                    Try
                        resultJson = PptGenerationService.GenerateQueued(
                            entry.ArgsJson,
                            Sub(pct As Integer, stage As String)
                                SyncLock _lock
                                    entry.Progress = Math.Min(100, Math.Max(0, pct))
                                    If Not String.IsNullOrEmpty(stage) Then entry.Stage = stage
                                End SyncLock
                            End Sub)
                        SyncLock _lock
                            entry.Status = STATUS_DONE
                            entry.ResultJson = resultJson
                            entry.Progress = 100
                            entry.Stage = "完成"
                            entry.FinishedAt = DateTime.Now
                        End SyncLock
                    Catch ex As Exception
                        SyncLock _lock
                            entry.Status = STATUS_ERROR
                            entry.Error = ex.Message
                            entry.FinishedAt = DateTime.Now
                        End SyncLock
                        ChatLogger.Instance.LogError("PPT_TASK", "任务 " & entry.TaskId & " 执行异常: " & ex.ToString())
                    End Try
                End While
            Catch ex As Exception
                SyncLock _lock
                    _workerBusy = False
                End SyncLock
                ChatLogger.Instance.LogError("PPT_TASK", "WorkerLoop 异常: " & ex.ToString())
            End Try
        End Sub

    End Class

End Namespace
