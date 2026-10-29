Imports System
Imports System.Runtime.CompilerServices

Namespace Utils

    ''' <summary>
    ''' 应用程序结构化日志门面 - 统一委托到 LogServiceManager 作为唯一日志后端
    ''' 仅保留原公共 API 签名，保证调用方兼容，不再各自维护文件写入实现
    ''' </summary>
    Public NotInheritable Class AppLogger

        ''' <summary>
        ''' 记录调试信息
        ''' </summary>
        <MethodImpl(MethodImplOptions.AggressiveInlining)>
        Public Shared Sub LogDebug(category As String, message As String)
            LogServiceManager.Instance.LogDebug(category, message)
        End Sub

        ''' <summary>
        ''' 记录一般信息
        ''' </summary>
        <MethodImpl(MethodImplOptions.AggressiveInlining)>
        Public Shared Sub LogInfo(category As String, message As String)
            LogServiceManager.Instance.LogInfo(category, message)
        End Sub

        ''' <summary>
        ''' 记录警告信息
        ''' </summary>
        <MethodImpl(MethodImplOptions.AggressiveInlining)>
        Public Shared Sub LogWarning(category As String, message As String)
            LogServiceManager.Instance.LogWarning(category, message)
        End Sub

        ''' <summary>
        ''' 记录错误信息
        ''' </summary>
        <MethodImpl(MethodImplOptions.AggressiveInlining)>
        Public Shared Sub LogError(category As String, message As String, Optional ex As Exception = Nothing)
            LogServiceManager.Instance.LogError(category, message, ex)
        End Sub

        ''' <summary>
        ''' 获取当前日志输出文件路径
        ''' </summary>
        Public Shared Function GetLogFilePath() As String
            Return LogServiceManager.Instance.GetLogFilePath()
        End Function

        Private Sub New()
            ' 静态使用，禁止实例化
        End Sub

    End Class

End Namespace