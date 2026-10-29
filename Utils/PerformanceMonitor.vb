Imports System.Diagnostics
Imports System.Threading

Namespace Utils

    ''' <summary>
    ''' 性能监控工具 - 用于监控操作耗时和内存使用
    ''' </summary>
    Public Class PerformanceMonitor

        Private Shared ReadOnly _lock As New Object()
        Private Shared _instance As PerformanceMonitor
        Private _operationTimings As New Dictionary(Of String, List(Of Long))()
        Private _isEnabled As Boolean = True

        Public Shared ReadOnly Property Instance As PerformanceMonitor
            Get
                If _instance Is Nothing Then
                    SyncLock _lock
                        If _instance Is Nothing Then
                            _instance = New PerformanceMonitor()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        ''' <summary>
        ''' 是否启用性能监控
        ''' </summary>
        Public Property IsEnabled As Boolean
            Get
                Return _isEnabled
            End Get
            Set(value As Boolean)
                _isEnabled = value
            End Set
        End Property

        ''' <summary>
        ''' 开始监控一个操作
        ''' </summary>
        ''' <param name="operationName">操作名称</param>
        ''' <returns>可释放的计时器对象</returns>
        Public Function StartOperation(operationName As String) As IDisposable
            If Not _isEnabled Then
                Return New NullTimer()
            End If
            Return New OperationTimer(operationName, Me)
        End Function

        ''' <summary>
        ''' 记录操作耗时
        ''' </summary>
        Private Sub RecordTiming(operationName As String, elapsedMilliseconds As Long)
            SyncLock _lock
                If Not _operationTimings.ContainsKey(operationName) Then
                    _operationTimings(operationName) = New List(Of Long)()
                End If
                _operationTimings(operationName).Add(elapsedMilliseconds)

                ' 只保留最近100次记录
                If _operationTimings(operationName).Count > 100 Then
                    _operationTimings(operationName).RemoveAt(0)
                End If
            End SyncLock

            ' 输出到调试窗口
            If elapsedMilliseconds > 100 Then
                System.Diagnostics.Debug.WriteLine($"[性能] {operationName}: {elapsedMilliseconds}ms")
            End If
        End Sub

        ''' <summary>
        ''' 获取操作的平均耗时
        ''' </summary>
        Public Function GetAverageTiming(operationName As String) As Double
            SyncLock _lock
                If Not _operationTimings.ContainsKey(operationName) OrElse _operationTimings(operationName).Count = 0 Then
                    Return 0
                End If
                Return _operationTimings(operationName).Average()
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有操作的统计信息
        ''' </summary>
        Public Function GetAllStatistics() As Dictionary(Of String, OperationStatistics)
            Dim result As New Dictionary(Of String, OperationStatistics)()

            SyncLock _lock
                For Each kvp In _operationTimings
                    Dim timings = kvp.Value
                    If timings.Count > 0 Then
                        result(kvp.Key) = New OperationStatistics With {
                            .OperationName = kvp.Key,
                            .Count = timings.Count,
                            .AverageMs = timings.Average(),
                            .MinMs = timings.Min(),
                            .MaxMs = timings.Max()
                        }
                    End If
                Next
            End SyncLock

            Return result
        End Function

        ''' <summary>
        ''' 记录当前内存使用情况
        ''' </summary>
        Public Shared Sub LogMemoryUsage(context As String)
            Try
                Dim proc = Process.GetCurrentProcess()
                Dim workingSetMB = proc.WorkingSet64 / 1024 / 1024
                Dim privateBytesMB = proc.PrivateMemorySize64 / 1024 / 1024
                Dim gcMemoryMB = GC.GetTotalMemory(False) / 1024 / 1024

                System.Diagnostics.Debug.WriteLine($"[内存] {context} - WorkingSet: {workingSetMB:F1}MB, PrivateBytes: {privateBytesMB:F1}MB, GC: {gcMemoryMB:F1}MB")
            Catch ex As Exception  '  忽略错误
            End Try
        End Sub

        ''' <summary>
        ''' 强制垃圾回收并记录内存变化
        ''' </summary>
        Public Shared Sub ForceGarbageCollection(context As String)
            Try
                Dim beforeMB = GC.GetTotalMemory(False) / 1024 / 1024
                GC.Collect()
                GC.WaitForPendingFinalizers()
                Dim afterMB = GC.GetTotalMemory(False) / 1024 / 1024

                System.Diagnostics.Debug.WriteLine($"[GC] {context} - Before: {beforeMB:F1}MB, After: {afterMB:F1}MB, Freed: {beforeMB - afterMB:F1}MB")
            Catch ex As Exception  '  忽略错误
            End Try
        End Sub

        ''' <summary>
        ''' 清空所有统计数据
        ''' </summary>
        Public Sub ClearStatistics()
            SyncLock _lock
                _operationTimings.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 操作统计信息
        ''' </summary>
        Public Class OperationStatistics
            Public Property OperationName As String
            Public Property Count As Integer
            Public Property AverageMs As Double
            Public Property MinMs As Long
            Public Property MaxMs As Long

            Public Overrides Function ToString() As String
                Return $"{OperationName}: 平均{AverageMs:F1}ms, 最小{MinMs}ms, 最大{MaxMs}ms, 次数{Count}"
            End Function
        End Class

        ''' <summary>
        ''' 操作计时器
        ''' </summary>
        Private Class OperationTimer
            Implements IDisposable

            Private _name As String
            Private _sw As Stopwatch
            Private _monitor As PerformanceMonitor

            Public Sub New(name As String, monitor As PerformanceMonitor)
                _name = name
                _monitor = monitor
                _sw = Stopwatch.StartNew()
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                _sw.Stop()
                _monitor.RecordTiming(_name, _sw.ElapsedMilliseconds)
            End Sub
        End Class

        ''' <summary>
        ''' 空计时器（用于禁用监控时）
        ''' </summary>
        Private Class NullTimer
            Implements IDisposable

            Public Sub Dispose() Implements IDisposable.Dispose
                ' 什么都不做
            End Sub
        End Class

    End Class

End Namespace
