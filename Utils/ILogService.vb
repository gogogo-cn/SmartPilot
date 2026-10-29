Option Strict On

Imports System

Namespace Utils

    ''' <summary>
    ''' 日志服务接口 - 统一的日志抽象层
    ''' 所有日志实现（ChatLogger、Logger、LogServiceManager）都应实现此接口
    ''' 注意：使用 LogLevel.vb 中定义的 LogLevel 枚举
    ''' </summary>
    Public Interface ILogService

        ''' <summary>
        ''' 当前最小日志级别（低于此级别的日志将被忽略）
        ''' </summary>
        Property MinLevel As LogLevel

        ''' <summary>
        ''' 日志来源标识
        ''' </summary>
        Property Source As String

        ''' <summary>
        ''' 记录调试日志
        ''' </summary>
        ''' <param name="category">日志分类</param>
        ''' <param name="message">日志消息</param>
        Sub LogDebug(category As String, message As String)

        ''' <summary>
        ''' 记录调试日志（带异常）
        ''' </summary>
        Sub LogDebug(category As String, message As String, ex As Exception)

        ''' <summary>
        ''' 记录信息日志
        ''' </summary>
        Sub LogInfo(category As String, message As String)

        ''' <summary>
        ''' 记录信息日志（带异常）
        ''' </summary>
        Sub LogInfo(category As String, message As String, ex As Exception)

        ''' <summary>
        ''' 记录警告日志
        ''' </summary>
        Sub LogWarning(category As String, message As String)

        ''' <summary>
        ''' 记录警告日志（带异常）
        ''' </summary>
        Sub LogWarning(category As String, message As String, ex As Exception)

        ''' <summary>
        ''' 记录错误日志
        ''' </summary>
        Sub LogError(category As String, message As String)

        ''' <summary>
        ''' 记录错误日志（带异常）
        ''' </summary>
        Sub LogError(category As String, message As String, ex As Exception)

        ''' <summary>
        ''' 记录严重错误日志
        ''' </summary>
        Sub LogCritical(category As String, message As String)

        ''' <summary>
        ''' 记录严重错误日志（带异常）
        ''' </summary>
        Sub LogCritical(category As String, message As String, ex As Exception)

        ''' <summary>
        ''' 通用日志方法
        ''' </summary>
        Sub Log(level As LogLevel, category As String, message As String, Optional ex As Exception = Nothing)

    End Interface

End Namespace