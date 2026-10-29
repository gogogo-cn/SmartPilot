Option Strict On
Option Explicit On
Option Infer On
Imports System.Threading
Imports System.Diagnostics

Namespace Utils

    ''' <summary>
    ''' 轻量级调用频率限制器 — 防止前端恶意/异常高频率调用 COM 方法
    ''' 基于固定时间窗口 + 最大调用次数
    ''' 【修复】使用 Stopwatch 单调时钟避免系统时间回拨影响
    ''' </summary>
    Public Class RateLimiter

        Private ReadOnly _maxCalls As Integer
        Private ReadOnly _windowMs As Integer
        ' 使用 Long 存储 Stopwatch 单调时间戳（毫秒），替代 DateTime.UtcNow
        Private ReadOnly _callTimestamps As New Queue(Of Long)
        Private ReadOnly _lock As New Object()
        ' 用于将 Stopwatch ticks 转换为毫秒的系数（只计算一次）
        Private Shared ReadOnly _tickFrequencyMs As Double = 1000.0 / Stopwatch.Frequency

        ''' <param name="maxCalls">窗口内最大调用次数</param>
        ''' <param name="windowMs">时间窗口（毫秒）</param>
        Public Sub New(maxCalls As Integer, windowMs As Integer)
            _maxCalls = maxCalls
            _windowMs = windowMs
        End Sub

        ''' <summary>
        ''' 支持使用 TimeSpan 构造函数的重载（向后兼容测试）
        ''' </summary>
        Public Sub New(maxCalls As Integer, window As TimeSpan)
            ' 直接在构造器链表达式中计算毫秒值，确保 Me.New 为首条语句
            Me.New(maxCalls, CInt(Math.Min(Math.Max(window.TotalMilliseconds, 1.0), CDbl(Integer.MaxValue))))
        End Sub

        ''' <summary>
        ''' 尝试获取调用许可。返回 True 表示允许调用。
        ''' </summary>
        Public Function TryAcquire() As Boolean
            SyncLock _lock
                ' 【安全说明】Stopwatch.GetTimestamp() 返回 Long 类型的滴答数（通常为100ns精度）
                ' 乘以 _tickFrequencyMs（double）后 CLng 转换为毫秒。
                ' Long.MaxValue ≈ 9.2×10^18，按当前频率约需连续运行 292 年才会溢出，
                ' 实际运行中不存在溢出风险。Stopwatch 为单调时钟，不受系统时间回拨影响。
                Dim nowMs As Long = CLng(Stopwatch.GetTimestamp() * _tickFrequencyMs)
                Dim cutoff As Long = nowMs - CLng(_windowMs)

                ' 清理过期的时间戳（单调时钟只增不减，回拨不影响正确性）
                While _callTimestamps.Count > 0 AndAlso _callTimestamps.Peek() < cutoff
                    _callTimestamps.Dequeue()
                End While

                ' 检查是否超限
                If _callTimestamps.Count >= _maxCalls Then
                    Return False
                End If

                ' 记录本次调用
                _callTimestamps.Enqueue(nowMs)
                Return True
            End SyncLock
        End Function

        ''' <summary>
        ''' 当前窗口内剩余可用次数
        ''' </summary>
        Public ReadOnly Property RemainingCalls As Integer
            Get
                SyncLock _lock
                    Dim cutoff As Long = CLng(Stopwatch.GetTimestamp() * _tickFrequencyMs) - CLng(_windowMs)
                    While _callTimestamps.Count > 0 AndAlso _callTimestamps.Peek() < cutoff
                        _callTimestamps.Dequeue()
                    End While
                    Return Math.Max(0, _maxCalls - _callTimestamps.Count)
                End SyncLock
            End Get
        End Property

        ' —— 预定义限流策略（Shared 字段确保全局唯一实例） ——
        Public Shared ReadOnly ForVBAExecution As New RateLimiter(maxCalls:=10, windowMs:=60000)     ' 每分钟10次
        Public Shared ReadOnly ForDataWrite As New RateLimiter(maxCalls:=30, windowMs:=60000)       ' 每分钟30次
        Public Shared ReadOnly ForGeneral As New RateLimiter(maxCalls:=60, windowMs:=60000)         ' 每分钟60次

    End Class

End Namespace