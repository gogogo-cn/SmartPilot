Imports System
Imports System.Threading.Tasks

Namespace Utils

    ''' <summary>
    ''' 重试策略 - 支持指数退避重试
    ''' </summary>
    Public Class RetryPolicy

        Public Property MaxRetries As Integer = 3
        Public Property InitialDelayMs As Integer = 1000
        Public Property BackoffMultiplier As Double = 2.0
        Public Property MaxDelayMs As Integer = 10000

        ''' <summary>
        ''' 判断异常是否应该重试（编程错误类异常不应重试）
        ''' </summary>
        Private Shared Function ShouldRetryException(ex As Exception) As Boolean
            ' 这些是编程错误，重试不会解决问题
            If TypeOf ex Is ArgumentException OrElse
               TypeOf ex Is ArgumentNullException OrElse
               TypeOf ex Is ArgumentOutOfRangeException OrElse
               TypeOf ex Is ObjectDisposedException Then
                Return False
            End If
            Return True
        End Function

        ''' <summary>
        ''' 执行带重试的异步操作（指数退避: 1s -> 2s -> 4s -> 8s，最大10秒）
        ''' </summary>
        Public Shared Async Function ExecuteAsync(Of T)(
            func As Func(Of Task(Of T)),
            Optional maxRetries As Integer = 3,
            Optional initialDelayMs As Integer = 1000) As Task(Of T)
        
            Dim lastException As Exception = Nothing
        
            For attempt = 0 To maxRetries
                Dim shouldRetry As Boolean = False
                Dim retryDelayMs As Integer = 0

                Try
                    Return Await func()
                Catch ex As Exception
                    lastException = ex
                            
                    ' 编程错误类异常不重试，直接抛出
                    If Not ShouldRetryException(ex) Then
                        Throw
                    End If

                    If attempt < maxRetries Then
                        shouldRetry = True
                        ' 指数退避: 1s -> 2s -> 4s（最大10秒）
                        retryDelayMs = CInt(Math.Min(
                            initialDelayMs * Math.Pow(2, attempt),
                            10000))
                        System.Diagnostics.Debug.WriteLine($"重试 {attempt + 1}/{maxRetries}, 等待 {retryDelayMs}ms: {ex.Message}")
                    End If
                End Try

                ' 在 Catch 块外执行 Await（.NET Framework 4.7.2 不支持 Catch 内 Await）
                If shouldRetry Then
                    Await Task.Delay(retryDelayMs)
                End If
            Next
        
            Dim rootMessage = If(lastException?.Message, "未知错误")
            Throw New InvalidOperationException($"操作失败,已重试{maxRetries}次: {rootMessage}", lastException)
        End Function

        ''' <summary>
        ''' 执行带重试的异步操作(无返回值)
        ''' </summary>
        Public Shared Async Function ExecuteAsync(
            action As Func(Of Task),
            Optional maxRetries As Integer = 3,
            Optional initialDelayMs As Integer = 1000) As Task

            Dim lastException As Exception = Nothing

            For attempt = 0 To maxRetries
                Dim shouldRetry As Boolean = False
                Dim retryDelayMs As Integer = 0

                Try
                    Await action()
                    Return
                Catch ex As Exception
                    lastException = ex

                    ' 编程错误类异常不重试，直接抛出
                    If Not ShouldRetryException(ex) Then
                        Throw
                    End If

                    If attempt < maxRetries Then
                        shouldRetry = True
                        ' 指数退避: 1s -> 2s -> 4s（最大10秒）
                        retryDelayMs = CInt(Math.Min(
                            initialDelayMs * Math.Pow(2, attempt),
                            10000))
                        System.Diagnostics.Debug.WriteLine($"重试 {attempt + 1}/{maxRetries}, 等待 {retryDelayMs}ms: {ex.Message}")
                    End If
                End Try

                ' 在 Catch 块外执行 Await（.NET Framework 4.7.2 不支持 Catch 内 Await）
                If shouldRetry Then
                    Await Task.Delay(retryDelayMs)
                End If
            Next

            Dim rootMessage2 = If(lastException?.Message, "未知错误")
            Throw New InvalidOperationException($"操作失败,已重试{maxRetries}次: {rootMessage2}", lastException)
        End Function

    End Class

End Namespace
