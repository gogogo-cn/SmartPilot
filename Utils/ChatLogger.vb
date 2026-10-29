Option Strict On

Imports System
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Utils

    ''' <summary>
    ''' 聊天日志记录器 - 门面模式,统一委托到 LogServiceManager
    ''' 保留原有API签名确保调用方兼容,不再独立写文件
    ''' </summary>
    Public Class ChatLogger
        Private Shared _instance As ChatLogger
        Private Shared _lock As New Object()
        Private _logLevel As LogLevel = LogLevel.Debug
        Private _isEnabled As Boolean = True

        ''' <summary>
        ''' 单例实例
        ''' </summary>
        Public Shared ReadOnly Property Instance As ChatLogger
            Get
                If _instance Is Nothing Then
                    SyncLock _lock
                        If _instance Is Nothing Then
                            _instance = New ChatLogger()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        Private Sub New()
            ' 不再初始化独立日志文件,统一使用 LogServiceManager
            LogServiceManager.Instance.LogInfo("ChatLogger", "日志记录器初始化完成(统一后端)")
        End Sub

        ''' <summary>
        ''' 设置日志级别
        ''' </summary>
        Public Sub SetLogLevel(level As LogLevel)
            _logLevel = level
            LogServiceManager.Instance.MinLevel = level
            LogInfo("ChatLogger", $"日志级别设置为: {level}")
        End Sub

        ''' <summary>
        ''' 启用/禁用日志
        ''' </summary>
        Public Sub SetEnabled(enabled As Boolean)
            _isEnabled = enabled
        End Sub

        ''' <summary>
        ''' 静态日志方法 - 用于快速记录日志
        ''' </summary>
        Public Shared Sub Log(message As String, Optional category As String = "General")
            Instance.LogInfo(category, message)
        End Sub

        ''' <summary>
        ''' 记录调试信息
        ''' </summary>
        Public Sub LogDebug(category As String, message As String, Optional details As String = "")
            If Not _isEnabled Then Return
            If _logLevel <= LogLevel.Debug Then
                Dim fullMessage = BuildFullMessage(message, details)
                LogServiceManager.Instance.LogDebug(category, fullMessage)
            End If
        End Sub

        ''' <summary>
        ''' 记录一般信息
        ''' </summary>
        Public Sub LogInfo(category As String, message As String, Optional details As String = "")
            If Not _isEnabled Then Return
            If _logLevel <= LogLevel.Info Then
                Dim fullMessage = BuildFullMessage(message, details)
                LogServiceManager.Instance.LogInfo(category, fullMessage)
            End If
        End Sub

        ''' <summary>
        ''' 记录警告信息
        ''' </summary>
        Public Sub LogWarning(category As String, message As String, Optional details As String = "")
            If Not _isEnabled Then Return
            If _logLevel <= LogLevel.Warning Then
                Dim fullMessage = BuildFullMessage(message, details)
                LogServiceManager.Instance.LogWarning(category, fullMessage)
            End If
        End Sub

        ''' <summary>
        ''' 记录错误信息
        ''' </summary>
        Public Sub LogError(category As String, message As String, Optional exception As Exception = Nothing)
            If Not _isEnabled Then Return
            If _logLevel <= LogLevel.Error Then
                LogServiceManager.Instance.LogError(category, message, exception)
            End If
        End Sub

        ''' <summary>
        ''' 记录聊天消息发送
        ''' </summary>
        Public Sub LogChatSend(userMessage As String, Optional model As String = "", Optional context As String = "")
            Dim details As New StringBuilder()
            details.AppendLine($"用户消息: {SanitizeForLog(TruncateString(userMessage, 200))}")
            If Not String.IsNullOrEmpty(model) Then
                details.AppendLine($"使用模型: {model}")
            End If
            If Not String.IsNullOrEmpty(context) Then
                details.AppendLine($"上下文信息: {SanitizeForLog(TruncateString(context, 500))}")
            End If
            LogInfo("CHAT_SEND", "发送聊天消息", details.ToString())
        End Sub

        ''' <summary>
        ''' 记录 AI 响应接收
        ''' </summary>
        Public Sub LogChatReceive(aiResponse As String, Optional processingTime As String = "", Optional tokenCount As String = "")
            Dim details As New StringBuilder()
            details.AppendLine($"AI响应: {SanitizeForLog(TruncateString(aiResponse, 500))}")
            If Not String.IsNullOrEmpty(processingTime) Then
                details.AppendLine($"处理时间: {processingTime}")
            End If
            If Not String.IsNullOrEmpty(tokenCount) Then
                details.AppendLine($"Token数量: {tokenCount}")
            End If
            LogInfo("CHAT_RECEIVE", "接收AI响应", details.ToString())
        End Sub

        ''' <summary>
        ''' 记录 AI 思考（内部链路）
        ''' </summary>
        Public Sub LogThoughts(category As String, thoughts As String)
            If _logLevel <= LogLevel.Debug Then
                LogDebug(category, "模型思考过程", thoughts)
            End If
        End Sub

        ''' <summary>
        ''' 记录 Skill 执行
        ''' </summary>
        Public Sub LogSkillExecution(skillName As String, inputs As Dictionary(Of String, Object), result As String, success As Boolean)
            Dim details As New StringBuilder()
            details.AppendLine($"Skill名称: {skillName}")
            details.AppendLine($"执行结果: {If(success, "成功", "失败")}")
            details.AppendLine($"输入参数:")
            For Each kvp In inputs
                Dim value As String = If(kvp.Value IsNot Nothing, kvp.Value.ToString(), "null")
                If value.Length > 100 Then
                    value = value.Substring(0, 100) & "..."
                End If
                details.AppendLine($"  {kvp.Key}: {value}")
            Next
            details.AppendLine($"返回结果: {result}")

            If success Then
                LogInfo("SKILL_EXEC", $"执行Skill: {skillName}", details.ToString())
            Else
                LogError("SKILL_EXEC", $"执行Skill失败: {skillName}")
            End If
        End Sub

        ''' <summary>
        ''' 记录 API 调用
        ''' </summary>
        Public Sub LogApiCall(provider As String, endpoint As String, requestBody As String, responseBody As String, success As Boolean)
            Dim details As New StringBuilder()
            details.AppendLine($"提供商: {provider}")
            details.AppendLine($"端点: {endpoint}")
            details.AppendLine($"请求体: {SanitizeForLog(TruncateString(requestBody, 500))}")
            details.AppendLine($"响应体: {SanitizeForLog(TruncateString(responseBody, 500))}")
            details.AppendLine($"调用结果: {If(success, "成功", "失败")}")

            If success Then
                LogDebug("API_CALL", $"API调用: {provider}", details.ToString())
            Else
                LogError("API_CALL", $"API调用失败: {provider}")
            End If
        End Sub

        ''' <summary>
        ''' 记录 WebMessage 通信
        ''' </summary>
        Public Sub LogWebMessage(direction As String, messageType As String, messageData As String)
            Dim details As New StringBuilder()
            details.AppendLine($"方向: {direction}")
            details.AppendLine($"消息类型: {messageType}")
            details.AppendLine($"消息数据: {SanitizeForLog(TruncateString(messageData, 1000))}")
            LogDebug("WEBMESSAGE", $"WebMessage {direction}", details.ToString())
        End Sub

        ''' <summary>
        ''' 记录 Excel 操作
        ''' </summary>
        Public Sub LogExcelOperation(operation As String, sheetName As String, range As String, Optional details As String = "")
            Dim sb As New StringBuilder()
            sb.AppendLine($"操作: {operation}")
            sb.AppendLine($"工作表: {sheetName}")
            sb.AppendLine($"范围: {range}")
            If Not String.IsNullOrEmpty(details) Then
                sb.AppendLine($"详细信息: {details}")
            End If
            LogInfo("EXCEL_OP", $"Excel操作: {operation}", sb.ToString())
        End Sub

        ''' <summary>
        ''' 记录配置变更
        ''' </summary>
        Public Sub LogConfigChange(configName As String, oldValue As String, newValue As String)
            Dim details As New StringBuilder()
            details.AppendLine($"配置项: {configName}")
            details.AppendLine($"旧值: {oldValue}")
            details.AppendLine($"新值: {newValue}")
            LogInfo("CONFIG", $"配置变更: {configName}", details.ToString())
        End Sub

        ''' <summary>
        ''' 获取当前日志文件路径
        ''' </summary>
        Public Function GetLogFilePath() As String
            Return LogServiceManager.Instance.GetLogFilePath()
        End Function

        ''' <summary>
        ''' 获取最近的日志内容
        ''' </summary>
        Public Function GetRecentLogs(lineCount As Integer) As String
            Try
                Dim logFile = LogServiceManager.Instance.GetLogFilePath()
                If Not File.Exists(logFile) Then Return "日志文件不存在"

                Dim allLines As List(Of String) = File.ReadAllLines(logFile).ToList()
                If allLines.Count <= lineCount Then
                    Return String.Join(vbCrLf, allLines)
                Else
                    Return String.Join(vbCrLf, allLines.Skip(allLines.Count - lineCount))
                End If
            Catch ex As Exception
                Return $"读取日志失败: {ex.Message}"
            End Try
        End Function

        ''' <summary>
        ''' 清除所有日志文件
        ''' </summary>
        Public Sub ClearAllLogs()
            Try
                Dim logFile = LogServiceManager.Instance.GetLogFilePath()
                Dim logDir = Path.GetDirectoryName(logFile)
                If Directory.Exists(logDir) Then
                    Dim logFiles = Directory.GetFiles(logDir, "log_*.log")
                    For Each lf In logFiles
                        Try
                            File.Delete(lf)
                        Catch
                        End Try
                    Next
                End If
                LogInfo("ChatLogger", "所有日志已清除")
            Catch ex As Exception
                LogError("ChatLogger", "清除日志失败", ex)
            End Try
        End Sub

#Region "私有辅助方法"

        ''' <summary>
        ''' 构建完整日志消息(含详情)
        ''' </summary>
        Private Function BuildFullMessage(message As String, details As String) As String
            If String.IsNullOrEmpty(details) Then Return message
            Return $"{message}{vbCrLf}详情:{vbCrLf}{details}"
        End Function

        ''' <summary>
        ''' 截断字符串
        ''' </summary>
        Private Function TruncateString(value As String, maxLength As Integer) As String
            If String.IsNullOrEmpty(value) Then Return ""
            If value.Length <= maxLength Then Return value
            Return value.Substring(0, maxLength) & "..."
        End Function

        ''' <summary>
        ''' 脱敏处理 - 移除日志中的敏感信息
        ''' </summary>
        Private Function SanitizeForLog(value As String) As String
            If String.IsNullOrEmpty(value) Then Return ""

            ' 身份证号/银行卡号（15-19位数字）
            value = System.Text.RegularExpressions.Regex.Replace(value, "\d{15,19}", "[身份证号/卡号]")

            ' 手机号（1开头11位数字）
            value = System.Text.RegularExpressions.Regex.Replace(value, "1[3-9]\d{9}", "[手机号]")

            ' API密钥（sk-开头，20位以上字母数字）
            value = System.Text.RegularExpressions.Regex.Replace(value, "sk-[a-zA-Z0-9]{20,}", "[API密钥]")

            ' Bearer Token
            value = System.Text.RegularExpressions.Regex.Replace(value, "Bearer\s+[a-zA-Z0-9\-._~+/]{20,}", "[Bearer Token]")

            ' 邮箱地址
            value = System.Text.RegularExpressions.Regex.Replace(value, "[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", "[邮箱]")

            ' 通用密钥字段
            value = System.Text.RegularExpressions.Regex.Replace(
                value,
                "(?i)(""?(?:apiKey|secretKey|api_secret|api_key|accessKey|access_key|appKey|app_key|password|passwd|token)""?\s*[=:]\s*"")[^""]*""",
                "$1[已隐藏]""")

            Return value
        End Function

#End Region

    End Class

End Namespace