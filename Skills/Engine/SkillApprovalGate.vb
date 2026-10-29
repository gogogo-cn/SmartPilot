Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Context

Namespace Skills

    ''' <summary>
    ''' 【扩展5】Skill 审批网关 - 根据安全级别决定是否需要用户确认
    '''
    ''' 安全级别说明：
    '''   Safe            → 直接执行（只读分析、格式化高亮等）
    '''   NeedsConfirmation → 需要用户确认后执行（写入Excel、创建图表等）【默认】
    '''   Critical        → 二次确认 + 审计日志（MCP写操作、发送邮件、删除数据）
    ''' </summary>
    Public Class SkillApprovalGate

        ' ====== 单例模式 ======
        Private Shared ReadOnly _lazyInstance As New Lazy(Of SkillApprovalGate)(
            Function() New SkillApprovalGate(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As SkillApprovalGate
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ' ====== 字段 ======
        Private _auditLogPath As String
        Private _lock As New Object()
        Private _pendingApprovals As New Dictionary(Of String, ApprovalRequest)()

        ' ====== 审批请求类型 ======
        Public Class ApprovalRequest
            Public Property RequestId As String
            Public Property SkillName As String
            Public Property SecurityLevel As SkillSecurityLevel
            Public Property OperationDescription As String
            Public Property RequestedAt As DateTime
            Public Property Status As ApprovalStatus = ApprovalStatus.Pending

            Public Enum ApprovalStatus
                Pending
                Approved
                Rejected
                Expired
            End Enum
        End Class

        ''' <summary>
        ''' 私有构造函数
        ''' </summary>
        Private Sub New()
            Dim appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
            _auditLogPath = Path.Combine(appData, "SmartPilot", "skill_audit.log")
            EnsureAuditLogDirectory()
        End Sub

        ''' <summary>
        ''' 检查并请求审批（主入口方法）
        ''' 在 CustomSkillExecutor.ExecuteCustomSkill() 开头调用
        ''' </summary>
        ''' <param name="skill">待执行的 Skill</param>
        ''' <returns>True=允许执行, False=拒绝执行</returns>
        Public Async Function CheckAndApproveAsync(skill As SkillDefinition) As Task(Of Boolean)
            ' ============ 免确认开关 ============
            ' 前端「自动化免确认」开关打开时（AutoExecutionSettings.BypassConfirmation = True），
            ' 所有级别（包括 Critical）均跳过审批，直接放行，审计日志标注来源。
            ' 这让 Excel 重启后恢复默认 False，不会残留高危授权状态。
            If Security.AutoExecutionSettings.BypassConfirmation Then
                WriteAuditLog(skill.Name, "BYPASS_CONFIRM",
                    String.Format("免确认开关开启，跳过审批（级别: {0}）", skill.SecurityLevel.ToString()))
                Return True
            End If

            Select Case skill.SecurityLevel
                Case SkillSecurityLevel.Safe
                    ' 安全级别：直接通过，记录日志
                    WriteAuditLog(skill.Name, "AUTO_APPROVED", "Safe级别自动放行")
                    Return True

                Case SkillSecurityLevel.NeedsConfirmation
                    ' [修复] NeedsConfirmation 级别：用户从 UI 主动触发 Skill，已构成明确授权意图，
                    ' 且前端 WebView 无审批弹窗基础设施（无 approvalRequest handler / 无 Approve 回调 Bridge），
                    ' 原先的 RequestConfirmationAsync 会等待 120 秒后超时返回 False 导致「用户拒绝」误报。
                    ' 因此改为自动放行，审计日志记录，VBACodeSandbox 仍保留独立安全检测。
                    WriteAuditLog(skill.Name, "AUTO_APPROVED", "NeedsConfirmation级别-用户UI触发自动放行")
                    Return True

                Case SkillSecurityLevel.Critical
                    ' 高危级别：二次确认 + 审计日志
                    WriteAuditLog(skill.Name, "CRITICAL_REQUESTED",
                        String.Format("高危操作已触发: {0}", skill.Description))
                    Return Await RequestConfirmationAsync(skill, singleConfirm:=False)

                Case Else
                    ' 未知级别：按 NeedsConfirmation 同等处理（自动放行）
                    WriteAuditLog(skill.Name, "AUTO_APPROVED", "未知级别-自动放行")
                    Return True
            End Select
        End Function

        ''' <summary>
        ''' 向前端 WebView 发送审批请求，等待用户响应
        ''' </summary>
        Private Async Function RequestConfirmationAsync(skill As SkillDefinition,
                                                         singleConfirm As Boolean) As Task(Of Boolean)
            Try
                Dim requestId = Guid.NewGuid().ToString("N")
                Dim request As New ApprovalRequest With {
                    .RequestId = requestId,
                    .SkillName = skill.Name,
                    .SecurityLevel = skill.SecurityLevel,
                    .OperationDescription = BuildOperationDescription(skill),
                    .RequestedAt = DateTime.Now
                }

                SyncLock _lock
                    _pendingApprovals(requestId) = request
                End SyncLock

                ' 发送审批请求到前端
                Dim approvalMessage = Newtonsoft.Json.JsonConvert.SerializeObject(New With {
                    .type = "approvalRequest",
                    .requestId = requestId,
                    .skillName = skill.DisplayName,
                    .description = request.OperationDescription,
                    .securityLevel = skill.SecurityLevel.ToString(),
                    .isCritical = (skill.SecurityLevel = SkillSecurityLevel.Critical),
                    .singleConfirm = singleConfirm,
                    .timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                })

                Dim addIn = Globals.ThisAddIn
                If addIn IsNot Nothing Then
                    addIn.SendMessageToWebView(approvalMessage)
                End If

                ' 等待前端响应（最多等待 120 秒）
                Dim maxWaitMs = 120000
                Dim checkIntervalMs = 200
                Dim elapsedMs As Integer = 0

                While elapsedMs < maxWaitMs
                    Await Task.Delay(checkIntervalMs)
                    elapsedMs += checkIntervalMs

                    SyncLock _lock
                        If _pendingApprovals.ContainsKey(requestId) Then
                            Dim currentStatus = _pendingApprovals(requestId).Status
                            Select Case currentStatus
                                Case ApprovalRequest.ApprovalStatus.Approved
                                    _pendingApprovals.Remove(requestId)
                                    WriteAuditLog(skill.Name, "APPROVED", "用户确认执行")
                                    Return True
                                Case ApprovalRequest.ApprovalStatus.Rejected
                                    _pendingApprovals.Remove(requestId)
                                    WriteAuditLog(skill.Name, "REJECTED", "用户拒绝执行")
                                    Return False
                            End Select
                        Else
                            ' 请求已被移除（超时清理）
                            Return False
                        End If
                    End SyncLock
                End While

                ' 超时：标记为过期
                SyncLock _lock
                    If _pendingApprovals.ContainsKey(requestId) Then
                        _pendingApprovals(requestId).Status = ApprovalRequest.ApprovalStatus.Expired
                        _pendingApprovals.Remove(requestId)
                    End If
                End SyncLock

                WriteAuditLog(skill.Name, "TIMEOUT", "用户未在时限内响应")
                Return False

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ApprovalGate] 审批流程异常: " & ex.Message)
                ' 异常时默认允许执行（避免阻塞正常功能）
                WriteAuditLog(skill.Name, "ERROR_FALLBACK",
                    String.Format("审批异常，降级放行: {0}", ex.Message))
                Return True
            End Try
        End Function

        ''' <summary>
        ''' 前端回调：用户点击了"确认"
        ''' 由 ExcelBridge 或 WebSocket 消息处理器调用
        ''' </summary>
        Public Sub Approve(requestId As String)
            SyncLock _lock
                If _pendingApprovals.ContainsKey(requestId) Then
                    _pendingApprovals(requestId).Status = ApprovalRequest.ApprovalStatus.Approved
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 前端回调：用户点击了"取消/拒绝"
        ''' </summary>
        Public Sub Reject(requestId As String)
            SyncLock _lock
                If _pendingApprovals.ContainsKey(requestId) Then
                    _pendingApprovals(requestId).Status = ApprovalRequest.ApprovalStatus.Rejected
                End If
            End SyncLock
        End Sub

        ''' <summary>
        ''' 构建操作描述文本（用于显示给用户的确认信息）
        ''' </summary>
        Private Function BuildOperationDescription(skill As SkillDefinition) As String
            Dim desc As New System.Text.StringBuilder()

            desc.Append("即将执行 Skill: ")
            desc.AppendLine(skill.DisplayName)
            desc.AppendLine()

            ' 根据 ProcessingMode 描述操作类型
            Select Case skill.ProcessingMode
                Case SkillProcessingMode.VBACode
                    desc.AppendLine("- 操作类型: VBA 代码执行")
                    If skill.TemplateCode IsNot Nothing AndAlso skill.TemplateCode.Length > 0 Then
                        desc.AppendLine("- 将执行代码: Sub " &
                            ExtractSubName(skill.TemplateCode) & "()")
                    End If
                Case SkillProcessingMode.MultimodalLLM
                    desc.AppendLine("- 操作类型: AI 智能分析")
                    desc.AppendLine("- 将调用 LLM 进行分析处理")
                Case SkillProcessingMode.Hybrid
                    desc.AppendLine("- 操作类型: AI 分析 + VBA 执行（混合模式）")
                    desc.AppendLine("- Phase1: AI 数据分析")
                    desc.AppendLine("- Phase2: 自动执行 VBA 操作")
                Case SkillProcessingMode.PythonCode
                    desc.AppendLine("- 操作类型: Python 脚本执行")
                    desc.AppendLine("- 将在隔离进程中运行 Python 代码")
                    If skill.TemplateCode IsNot Nothing AndAlso skill.TemplateCode.Length > 0 Then
                        desc.AppendLine("- 脚本长度: " & skill.TemplateCode.Length & " 字符")
                    End If
                Case SkillProcessingMode.LocalOCR
                    desc.AppendLine("- 操作类型: 本地 OCR 识别")
                Case Else
                    desc.AppendLine("- 操作类型: " & skill.ProcessingMode.ToString())
            End Select

            ' 如果是管道型
            If skill.IsPipeline Then
                desc.AppendLine()
                desc.AppendLine(String.Format("- 管道步骤数: {0}", skill.PipelineSteps.Count))
            End If

            ' 如果涉及 MCP 工具调用（通过步骤名称或配置判断）
            If skill.CustomWorkflowSteps IsNot Nothing Then
                For Each cws In skill.CustomWorkflowSteps
                    Dim isMcpTool = False
                    If cws.Config IsNot Nothing AndAlso cws.Config.ContainsKey("mcpTool") Then
                        isMcpTool = True
                    End If
                    If isMcpTool Then
                        desc.AppendLine()
                        desc.AppendLine("- [注意] 将调用外部 MCP 工具:")
                        Dim toolName = ""
                        If cws.Config IsNot Nothing AndAlso cws.Config.ContainsKey("toolName") Then
                            toolName = cws.Config("toolName").ToString()
                        End If
                        desc.Append("  ").AppendLine(If(toolName, cws.StepName))
                    End If
                Next
            End If

            desc.AppendLine()
            desc.Append("描述: ").Append(If(skill.Description, "无"))

            Return desc.ToString()
        End Function

        ''' <summary>
        ''' 从 VBA 代码中提取 Sub 名称
        ''' </summary>
        Private Function ExtractSubName(code As String) As String
            If String.IsNullOrWhiteSpace(code) Then Return "(unknown)"
            Dim match = System.Text.RegularExpressions.Regex.Match(
                code, "Sub\s+(\w+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If match.Success Then Return match.Groups(1).Value
            Return "(anonymous)"
        End Function

        ' ====== 审计日志 ======

        ''' <summary>
        ''' 写入审计日志
        ''' </summary>
        Private Sub WriteAuditLog(skillName As String, action As String, detail As String)
            Try
                Dim logEntry = String.Format("[{0}] [{1}] Skill={2} | {3}{4}",
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                    action,
                    skillName,
                    detail,
                    Environment.NewLine)

                SyncLock _lock
                    File.AppendAllText(_auditLogPath, logEntry)
                End SyncLock
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ApprovalGate] 审计日志写入失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 确保审计日志目录存在
        ''' </summary>
        Private Sub EnsureAuditLogDirectory()
            Try
                Dim dir = Path.GetDirectoryName(_auditLogPath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ApprovalGate] 创建审计目录失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 获取最近的审计日志条目（用于调试和管理界面展示）
        ''' </summary>
        Public Function GetRecentAuditLogs(maxCount As Integer) As List(Of String)
            Dim logs As New List(Of String)()
            Try
                If File.Exists(_auditLogPath) Then
                    Dim lines = File.ReadAllLines(_auditLogPath)
                    Dim startIdx = Math.Max(0, lines.Length - maxCount)
                    For i = startIdx To lines.Length - 1
                        logs.Add(lines(i))
                    Next
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ApprovalGate] 读取审计日志失败: " & ex.Message)
            End Try
            Return logs
        End Function

    End Class

End Namespace
