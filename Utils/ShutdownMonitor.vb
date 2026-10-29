Imports System.Diagnostics
Imports System.Text
Imports System.IO

Namespace Utils

    ''' <summary>
    ''' 关闭流程监控器 - 精确定位插件关闭时的卡顿原因
    ''' 记录每个清理步骤的耗时、线程状态、资源占用
    ''' </summary>
    Public NotInheritable Class ShutdownMonitor

        Private Shared _instance As ShutdownMonitor
        Private Shared _lock As New Object()

        Private _steps As New List(Of ShutdownStep)()
        Private _currentStep As ShutdownStep = Nothing
        Private _startTime As DateTime
        Private _isMonitoring As Boolean = False

        ''' <summary>
        ''' 单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As ShutdownMonitor
            Get
                If _instance Is Nothing Then
                    SyncLock _lock
                        If _instance Is Nothing Then
                            _instance = New ShutdownMonitor()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        ''' <summary>
        ''' 开始监控关闭流程
        ''' </summary>
        Public Sub StartMonitoring()
            SyncLock _lock
                _steps.Clear()
                _startTime = DateTime.Now
                _isMonitoring = True
            End SyncLock

            Dim proc = Process.GetCurrentProcess()
            Dim msg = "===== 关闭流程监控开始 @ " & _startTime.ToString("HH:mm:ss.fff") & " ====="
            Debug.WriteLine(msg)
            Debug.WriteLine("[关闭监控] 当前内存: " & (proc.WorkingSet64 / 1024 / 1024).ToString("F1") & "MB, 线程数: " & proc.Threads.Count.ToString() & ", 句柄数: " & proc.HandleCount.ToString())
            AppLogger.LogInfo("SHUTDOWN", msg)
        End Sub

        ''' <summary>
        ''' 开始记录一个步骤
        ''' </summary>
        ''' <param name="stepName">步骤名称</param>
        Public Sub BeginStep(stepName As String)
            If Not _isMonitoring Then Return

            SyncLock _lock
                _currentStep = New ShutdownStep With {
                    .StepName = stepName,
                    .StartTime = DateTime.Now,
                    .Stopwatch = Stopwatch.StartNew(),
                    .ThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId,
                    .IsBackgroundThread = System.Threading.Thread.CurrentThread.IsBackground,
                    .MemoryBeforeMB = Process.GetCurrentProcess().WorkingSet64 / 1024 / 1024
                }
            End SyncLock

            Debug.WriteLine("[关闭监控] ▶ 开始: " & stepName & " (线程#" & _currentStep.ThreadId.ToString() & If(_currentStep.IsBackgroundThread, " [后台]", " [UI]") & ")")
        End Sub

        ''' <summary>
        ''' 完成当前步骤记录
        ''' </summary>
        ''' <param name="status">步骤状态</param>
        ''' <param name="details">详细信息</param>
        Public Sub EndStep(status As String, Optional details As String = "")
            If Not _isMonitoring OrElse _currentStep Is Nothing Then Return

            SyncLock _lock
                _currentStep.Stopwatch.Stop()
                _currentStep.EndTime = DateTime.Now
                _currentStep.ElapsedMs = _currentStep.Stopwatch.ElapsedMilliseconds
                _currentStep.Status = status
                _currentStep.Details = details
                _currentStep.MemoryAfterMB = Process.GetCurrentProcess().WorkingSet64 / 1024 / 1024
                _currentStep.MemoryDeltaMB = _currentStep.MemoryAfterMB - _currentStep.MemoryBeforeMB

                _steps.Add(_currentStep)

                Dim memChange = If(_currentStep.MemoryDeltaMB > 0, "+" & _currentStep.MemoryDeltaMB.ToString("F1") & "MB", _currentStep.MemoryDeltaMB.ToString("F1") & "MB")
                Dim warning = If(_currentStep.ElapsedMs > 500, " ⚠️ 耗时过长!", "")
                Debug.WriteLine("[关闭监控] ✓ 完成: " & _currentStep.StepName & " - " & _currentStep.ElapsedMs.ToString() & "ms [" & status & "] 内存变化:" & memChange & warning)
                If Not String.IsNullOrEmpty(details) Then
                    Debug.WriteLine("[关闭监控]   详情: " & details)
                End If

                AppLogger.LogInfo("SHUTDOWN", "步骤完成: " & _currentStep.StepName & " - " & _currentStep.ElapsedMs.ToString() & "ms [" & status & "]")

                _currentStep = Nothing
            End SyncLock
        End Sub

        ''' <summary>
        ''' 记录一个即时事件(无耗时)
        ''' </summary>
        Public Sub LogEvent(eventName As String, Optional details As String = "")
            If Not _isMonitoring Then Return

            Debug.WriteLine("[关闭监控] • 事件: " & eventName)
            If Not String.IsNullOrEmpty(details) Then
                Debug.WriteLine("[关闭监控]   详情: " & details)
            End If
            AppLogger.LogInfo("SHUTDOWN", "事件: " & eventName & " - " & details)
        End Sub

        ''' <summary>
        ''' 记录异常
        ''' </summary>
        Public Sub LogException(stepName As String, ex As Exception)
            If Not _isMonitoring Then Return

            Debug.WriteLine("[关闭监控] ✗ 异常: " & stepName & " - " & ex.GetType().Name & ": " & ex.Message)
            Debug.WriteLine("[关闭监控]   堆栈: " & ex.StackTrace)
            AppLogger.LogError("SHUTDOWN", "步骤异常: " & stepName, ex)

            ' 自动记录为失败步骤
            SyncLock _lock
                _steps.Add(New ShutdownStep With {
                    .StepName = stepName,
                    .StartTime = DateTime.Now,
                    .EndTime = DateTime.Now,
                    .ElapsedMs = 0,
                    .Status = "EXCEPTION",
                    .Details = ex.GetType().Name & ": " & ex.Message,
                    .ThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId
                })
            End SyncLock
        End Sub

        ''' <summary>
        ''' 结束监控并生成报告
        ''' </summary>
        Public Function GenerateReport() As String
            SyncLock _lock
                _isMonitoring = False

                Dim totalTime = (DateTime.Now - _startTime).TotalMilliseconds
                Dim proc = Process.GetCurrentProcess()

                Dim sb As New StringBuilder()
                sb.AppendLine("========== 关闭流程监控报告 ==========")
                sb.AppendLine("开始时间: " & _startTime.ToString("yyyy-MM-dd HH:mm:ss.fff"))
                sb.AppendLine("总耗时: " & totalTime.ToString("F0") & "ms (" & (totalTime / 1000).ToString("F2") & "秒)")
                sb.AppendLine("结束内存: " & (proc.WorkingSet64 / 1024 / 1024).ToString("F1") & "MB, 线程数: " & proc.Threads.Count.ToString() & ", 句柄数: " & proc.HandleCount.ToString())
                sb.AppendLine()

                ' 步骤详情
                sb.AppendLine("--- 步骤详情 ---")
                For i = 0 To _steps.Count - 1
                    Dim stp = _steps(i)
                    Dim percent = If(totalTime > 0, (stp.ElapsedMs / totalTime * 100), 0)
                    Dim warning = If(stp.ElapsedMs > 500, " [⚠️ 慢]", "")
                    sb.AppendLine("[" & (i + 1).ToString("D2") & "] " & stp.StepName)
                    sb.AppendLine("     耗时: " & stp.ElapsedMs.ToString() & "ms (" & percent.ToString("F1") & "%)" & warning)
                    sb.AppendLine("     状态: " & stp.Status)
                    sb.AppendLine("     线程: #" & stp.ThreadId.ToString() & If(stp.IsBackgroundThread, " [后台]", " [UI]") & " ")
                    If stp.MemoryDeltaMB <> 0 Then
                        sb.AppendLine("     内存: " & If(stp.MemoryDeltaMB > 0, "+", "") & stp.MemoryDeltaMB.ToString("F1") & "MB (Before:" & stp.MemoryBeforeMB.ToString("F1") & "MB → After:" & stp.MemoryAfterMB.ToString("F1") & "MB)")
                    End If
                    If Not String.IsNullOrEmpty(stp.Details) Then
                        sb.AppendLine("     详情: " & stp.Details)
                    End If
                    sb.AppendLine()
                Next

                ' 性能瓶颈分析
                sb.AppendLine("--- 性能瓶颈分析 ---")
                Dim slowSteps = _steps.Where(Function(s) s.ElapsedMs > 200).OrderByDescending(Function(s) s.ElapsedMs).ToList()
                If slowSteps.Count > 0 Then
                    sb.AppendLine("发现 " & slowSteps.Count.ToString() & " 个耗时超过 200ms 的步骤:")
                    For Each stp In slowSteps
                        sb.AppendLine("  • " & stp.StepName & ": " & stp.ElapsedMs.ToString() & "ms [" & stp.Status & "]")
                    Next
                Else
                    sb.AppendLine("未发现明显性能瓶颈(所有步骤 < 200ms)")
                End If
                sb.AppendLine()

                ' 阻塞风险警告
                sb.AppendLine("--- 阻塞风险警告 ---")
                Dim uiThreadSlowSteps = _steps.Where(Function(s) Not s.IsBackgroundThread AndAlso s.ElapsedMs > 300).ToList()
                If uiThreadSlowSteps.Count > 0 Then
                    sb.AppendLine("⚠️ 发现 " & uiThreadSlowSteps.Count.ToString() & " 个 UI 线程慢操作(>300ms),可能导致界面卡顿:")
                    For Each stp In uiThreadSlowSteps
                        sb.AppendLine("  • " & stp.StepName & ": " & stp.ElapsedMs.ToString() & "ms")
                    Next
                Else
                    sb.AppendLine("✓ UI 线程无慢操作")
                End If
                sb.AppendLine()

                ' 异常汇总
                Dim exceptions = _steps.Where(Function(s) s.Status = "EXCEPTION").ToList()
                If exceptions.Count > 0 Then
                    sb.AppendLine("--- 异常汇总 (" & exceptions.Count.ToString() & " 个) ---")
                    For Each stp In exceptions
                        sb.AppendLine("  ✗ " & stp.StepName & ": " & stp.Details)
                    Next
                    sb.AppendLine()
                End If

                sb.AppendLine("========================================")

                Dim report = sb.ToString()
                Debug.WriteLine(report)

                ' 写入日志文件
                Try
                    Dim logPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                        "ExcelAIVSTOAddIn",
                        "Logs",
                        "shutdown_report_" & DateTime.Now.ToString("yyyyMMdd_HHmmss") & ".txt"
                    )
                    File.WriteAllText(logPath, report, Encoding.UTF8)
                    Debug.WriteLine("[关闭监控] 报告已保存至: " & logPath)
                Catch ex As Exception
                    Debug.WriteLine("[关闭监控] 保存报告失败: " & ex.Message)
                End Try

                Return report
            End SyncLock
        End Function

        ''' <summary>
        ''' 重置监控器
        ''' </summary>
        Public Sub Reset()
            SyncLock _lock
                _steps.Clear()
                _currentStep = Nothing
                _isMonitoring = False
            End SyncLock
        End Sub

        ''' <summary>
        ''' 结束监控并生成报告
        ''' </summary>
        Public Sub EndMonitoring()
            GenerateReport()
        End Sub

        ''' <summary>
        ''' 关闭步骤信息
        ''' </summary>
        Private Class ShutdownStep
            Public Property StepName As String
            Public Property StartTime As DateTime
            Public Property EndTime As DateTime
            Public Property ElapsedMs As Long
            Public Property Status As String
            Public Property Details As String
            Public Property ThreadId As Integer
            Public Property IsBackgroundThread As Boolean
            Public Property MemoryBeforeMB As Double
            Public Property MemoryAfterMB As Double
            Public Property MemoryDeltaMB As Double
            Public Property Stopwatch As Stopwatch
        End Class

    End Class

End Namespace
