Option Strict On

Imports System
Imports System.Collections.Concurrent
Imports System.IO
Imports System.Threading.Tasks

Namespace Utils

    ''' <summary>
    ''' 统一日志服务管理器 - 管理多个日志服务实例
    ''' 支持同时输出到多个目标（文件、控制台、事件日志等）
    ''' </summary>
    Public Class LogServiceManager
        Implements ILogService

        ' 单例实例
        Private Shared ReadOnly _lazyInstance As New Lazy(Of LogServiceManager)(
            Function() New LogServiceManager(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' 日志服务集合
        Private ReadOnly _logServices As New ConcurrentBag(Of ILogService)

        ' 日志文件路径
        Private ReadOnly _logFilePath As String

        ' 日志队列（异步写入）
        Private ReadOnly _logQueue As New ConcurrentQueue(Of LogEntry)

        ' 后台写入任务
        Private _writeTask As Task
        Private _isRunning As Boolean = False

        ' 关闭唤醒信号：Stop 时立即打断1s轮询，避免Excel关闭固定多等1秒
        Private _stopSignal As New Threading.ManualResetEventSlim(False)

        ' 最小日志级别
        Private _minLevel As LogLevel = LogLevel.Info

        ''' <summary>
        ''' 获取单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As LogServiceManager
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 最小日志级别
        ''' </summary>
        Public Property MinLevel As LogLevel Implements ILogService.MinLevel
            Get
                Return _minLevel
            End Get
            Set(value As LogLevel)
                _minLevel = value
                ' 同步更新所有子服务的级别
                For Each service In _logServices
                    service.MinLevel = value
                Next
            End Set
        End Property

        ''' <summary>
        ''' 日志来源标识
        ''' </summary>
        Public Property Source As String Implements ILogService.Source
            Get
                Return "LogServiceManager"
            End Get
            Set(value As String)
                ' 不设置，保持固定值
            End Set
        End Property

        ''' <summary>
        ''' 私有构造函数
        ''' </summary>
        Private Sub New()
            ' 设置日志文件路径
            _logFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "ExcelAIVSTOAddIn",
                "Logs",
                String.Format("log_{0:yyyyMMdd}.log", DateTime.Now))

            ' 确保日志目录存在
            Dim logDir = Path.GetDirectoryName(_logFilePath)
            If Not Directory.Exists(logDir) Then
                Directory.CreateDirectory(logDir)
            End If

            ' 启动后台写入任务
            StartBackgroundWriter()

            ' 注册应用退出事件
            AddHandler AppDomain.CurrentDomain.ProcessExit, AddressOf OnProcessExit
        End Sub

        ''' <summary>
        ''' 注册日志服务
        ''' </summary>
        Public Sub RegisterService(service As ILogService)
            If service IsNot Nothing AndAlso Not _logServices.Contains(service) Then
                _logServices.Add(service)
            End If
        End Sub

        ''' <summary>
        ''' 移除日志服务
        ''' </summary>
        Public Sub RemoveService(service As ILogService)
            ' ConcurrentBag 不支持直接移除，使用新集合替换
            Dim newServices As New ConcurrentBag(Of ILogService)
            For Each s In _logServices
                If s IsNot service Then
                    newServices.Add(s)
                End If
            Next
            ' ConcurrentBag 没有 Clear 方法，使用Interlocked.Exchange 替换
            System.Threading.Interlocked.Exchange(_logServices, newServices)
        End Sub

        ''' <summary>
        ''' 启动后台写入任务
        ''' </summary>
        Private Sub StartBackgroundWriter()
            If _isRunning Then Return
            _isRunning = True
            _writeTask = Task.Run(Async Function()
                While _isRunning
                    Await WriteLogsAsync()
                    _stopSignal.Wait(1000) ' 每秒写入一次；Stop 时立即被唤醒
                End While
            End Function)
        End Sub

        ''' <summary>
        ''' 异步写入日志到文件
        ''' </summary>
        Private Async Function WriteLogsAsync() As Task
            Try
                Dim entries As New List(Of LogEntry)
                Dim entry As LogEntry = Nothing
                While _logQueue.TryDequeue(entry)
                    entries.Add(entry)
                End While

                If entries.Count > 0 Then
                    Using writer As New StreamWriter(_logFilePath, True, System.Text.Encoding.UTF8)
                        For Each entry In entries
                            Dim line = FormatLogEntry(entry)
                            Await writer.WriteLineAsync(line)
                        Next
                    End Using
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("[LogServiceManager] 写入日志失败: {0}", ex.Message))
            End Try
        End Function

        ''' <summary>
        ''' 格式化日志条目
        ''' </summary>
        Private Function FormatLogEntry(entry As LogEntry) As String
            Dim levelStr = entry.Level.ToString().ToUpper()
            Dim timestamp = entry.Timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Dim exStr = If(entry.Exception IsNot Nothing, String.Format(" | Exception: {0}", entry.Exception.Message), "")
            Return String.Format("[{0}] [{1}] [{2}] {3}{4}", timestamp, levelStr, entry.Category, entry.Message, exStr)
        End Function

        ''' <summary>
        ''' 应用退出时停止写入任务
        ''' </summary>
        Private Sub OnProcessExit(sender As Object, e As EventArgs)
            StopBackgroundWriter()
        End Sub

        ''' <summary>
        ''' 停止后台写入任务（最多等待2秒）
        ''' </summary>
        Public Sub StopBackgroundWriter()
            _isRunning = False
            _stopSignal.Set()
            Try
                If _writeTask IsNot Nothing Then
                    ' 循环被Set 唤醒后做最后一次写入并退出，通常毫秒级返回
                    _writeTask.Wait(2000)
                End If
                ' 兜底：清空剩余队列（通常为空或极少）
                WriteLogsAsync().Wait(2000)
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 获取当前日志输出文件路径
        ''' </summary>
        Public Function GetLogFilePath() As String
            Return _logFilePath
        End Function

        ''' <summary>
        ''' 清理历史日志文件（保留最近N天）
        ''' </summary>
        Public Sub CleanOldLogs(daysToKeep As Integer)
            Try
                Dim logDir = Path.GetDirectoryName(_logFilePath)
                If Not Directory.Exists(logDir) Then Return
                Dim cutoff = DateTime.Now.AddDays(-daysToKeep)
                For Each file In Directory.GetFiles(logDir, "log_*.log")
                    Try
                        Dim fileInfo = New FileInfo(file)
                        If fileInfo.LastWriteTime < cutoff Then
                            fileInfo.Delete()
                            LogInfo("LogServiceManager", "清理旧日志文件: " & file)
                        End If
                    Catch ex As Exception
                        LogWarning("LogServiceManager", "删除旧日志文件失败: " & file, ex)
                    End Try
                Next
            Catch ex As Exception
                LogWarning("LogServiceManager", "清理旧日志失败", ex)
            End Try
        End Sub

#Region "ILogService 实现"

        Public Sub LogDebug(category As String, message As String) Implements ILogService.LogDebug
            Log(LogLevel.Debug, category, message)
        End Sub

        Public Sub LogDebug(category As String, message As String, ex As Exception) Implements ILogService.LogDebug
            Log(LogLevel.Debug, category, message, ex)
        End Sub

        Public Sub LogInfo(category As String, message As String) Implements ILogService.LogInfo
            Log(LogLevel.Info, category, message)
        End Sub

        Public Sub LogInfo(category As String, message As String, ex As Exception) Implements ILogService.LogInfo
            Log(LogLevel.Info, category, message, ex)
        End Sub

        Public Sub LogWarning(category As String, message As String) Implements ILogService.LogWarning
            Log(LogLevel.Warning, category, message)
        End Sub

        Public Sub LogWarning(category As String, message As String, ex As Exception) Implements ILogService.LogWarning
            Log(LogLevel.Warning, category, message, ex)
        End Sub

        Public Sub LogError(category As String, message As String) Implements ILogService.LogError
            Log(LogLevel.Error, category, message)
        End Sub

        Public Sub LogError(category As String, message As String, ex As Exception) Implements ILogService.LogError
            Log(LogLevel.Error, category, message, ex)
        End Sub

        Public Sub LogCritical(category As String, message As String) Implements ILogService.LogCritical
            Log(LogLevel.Fatal, category, message)
        End Sub

        Public Sub LogCritical(category As String, message As String, ex As Exception) Implements ILogService.LogCritical
            Log(LogLevel.Fatal, category, message, ex)
        End Sub

        ''' <summary>
        ''' 通用日志方法 - 同时输出到所有注册的服务
        ''' </summary>
        Public Sub Log(level As LogLevel, category As String, message As String, Optional ex As Exception = Nothing) Implements ILogService.Log
            ' 检查日志级别
            If level < _minLevel Then Return

            ' 添加到队列（异步写入文件）
            _logQueue.Enqueue(New LogEntry(level, category, message, ex, DateTime.Now))

            ' 同时输出到所有注册的服务
            For Each service In _logServices
                Try
                    service.Log(level, category, message, ex)
                Catch
                    ' 忽略单个服务的错误
                End Try
            Next
        End Sub

#End Region

        ''' <summary>
        ''' 日志条目结构
        ''' </summary>
        Private Structure LogEntry
            Public ReadOnly Level As LogLevel
            Public ReadOnly Category As String
            Public ReadOnly Message As String
            Public ReadOnly Exception As Exception
            Public ReadOnly Timestamp As DateTime

            Public Sub New(level As LogLevel, category As String, message As String, ex As Exception, timestamp As DateTime)
                Me.Level = level
                Me.Category = category
                Me.Message = message
                Me.Exception = ex
                Me.Timestamp = timestamp
            End Sub
        End Structure

    End Class

End Namespace