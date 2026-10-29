Option Explicit On
Option Infer On

Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services

Namespace Skills

    Public Enum SkillExecutionPhase
        Initializing = 0
        Validating = 1
        Executing = 2
        PostProcessing = 3
        Completed = 4
        Failed = 5
        TimedOut = 6
        Retrying = 7
    End Enum

    Public Enum SkillErrorCategory
        Unknown = 0
        Transient = 1      ' 可重试：网络抖动、临时资源不可用、超时
        Permanent = 2      ' 不可重试：配置错误、参数非法、Skill不存在
        UserError = 3      ' 用户输入错误：Excel保护、无权限、范围无效
        SystemCrash = 4    ' COM异常、OutOfMemory：需要全局恢复
    End Enum

    Public Class SkillRetryPolicy
        Public Property MaxRetryCount As Integer = 3
        Public Property BaseDelayMs As Integer = 500
        Public Property EnableExponentialBackoff As Boolean = True
        Public Property RetryOnTransient As Boolean = True
        Public Property RetryOnTimeout As Boolean = False

        Public Shared Function [Default]() As SkillRetryPolicy
            Return New SkillRetryPolicy()
        End Function

        Public Shared Function NoRetry() As SkillRetryPolicy
            Return New SkillRetryPolicy() With {.MaxRetryCount = 0}
        End Function

        Public Function ComputeDelayMs(retryAttempt As Integer) As Integer
            If retryAttempt <= 0 Then Return 0
            If Not EnableExponentialBackoff Then Return BaseDelayMs
            Return CInt(BaseDelayMs * Math.Pow(2, Math.Min(retryAttempt - 1, 5)))
        End Function
    End Class

    Public Class SkillExecutionContext
        Public Property SkillName As String
        Public Property Skill As SkillDefinition
        Public Property Inputs As Dictionary(Of String, Object)
        Public Property Progress As IProgress(Of String)
        Public Property Phase As SkillExecutionPhase
        Public Property ExecutionStartTime As DateTime
        Public Property ExecutionEndTime As DateTime
        Public Property ErrorMessage As String
        Public Property IsComposite As Boolean
        Public Property ProcessingMode As SkillProcessingMode
        Public Property TimeoutSeconds As Integer
        Public Property CancellationToken As CancellationToken
        Friend Property CancellationTokenSource As CancellationTokenSource

        Public Property RetryPolicy As SkillRetryPolicy
        Public Property RetryCount As Integer
        Public Property LastErrorCategory As SkillErrorCategory
        Public Property LastException As Exception

        Public Sub New()
            TimeoutSeconds = 300
            RetryPolicy = SkillRetryPolicy.Default()
            RetryCount = 0
            LastErrorCategory = SkillErrorCategory.Unknown
        End Sub
    End Class

    Public Class SkillExecutionRouter
        Implements IDisposable

        Private _skillManager As SkillManager
        Private _workflowEngine As WorkflowEngine
        Private _llmEngine As MultimodalLLMEngine
        Private _isDisposed As Boolean

        Private _defaultTimeoutSeconds As Integer = 300
        Private _diagnosticLog As New List(Of SkillDiagnosticEntry)()
        Private _diagnosticSyncLock As New Object()
        Private Const MAX_DIAGNOSTIC_ENTRIES As Integer = 500

        Public Event ExecutionStarted(context As SkillExecutionContext)
        Public Event ExecutionCompleted(context As SkillExecutionContext, result As WorkflowResult)
        Public Event ExecutionFailed(context As SkillExecutionContext, ex As Exception)
        Public Event PhaseChanged(context As SkillExecutionContext, previousPhase As SkillExecutionPhase)
        Public Event ExecutionTimedOut(context As SkillExecutionContext, timeoutSeconds As Integer)
        Public Event ExecutionRetrying(context As SkillExecutionContext, retryAttempt As Integer, delayMs As Integer, ex As Exception)

        Private Sub New()
            _skillManager = SkillManager.Instance
        End Sub

        Private Shared ReadOnly _lazyInstance As New Lazy(Of SkillExecutionRouter)(
            Function() New SkillExecutionRouter(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As SkillExecutionRouter
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Public Sub SetDefaultTimeoutSeconds(seconds As Integer)
            If seconds > 0 Then
                _defaultTimeoutSeconds = seconds
            End If
        End Sub

        Public Function GetDefaultTimeoutSeconds() As Integer
            Return _defaultTimeoutSeconds
        End Function

        Public Async Function ExecuteSkill(skillName As String, inputs As Dictionary(Of String, Object),
                                          Optional progress As IProgress(Of String) = Nothing,
                                          Optional timeoutSeconds As Integer = 0,
                                          Optional cancellationToken As CancellationToken = Nothing,
                                          Optional retryPolicy As SkillRetryPolicy = Nothing) As Task(Of WorkflowResult)
            Dim context = CreateExecutionContext(skillName, inputs, progress, timeoutSeconds, cancellationToken, retryPolicy)
            Dim effectiveTimeout = If(context.TimeoutSeconds > 0, context.TimeoutSeconds, _defaultTimeoutSeconds)
            Dim pipelineStartTime = DateTime.Now

            Try
                RaiseEvent ExecutionStarted(context)

                While True
                    Dim ctsRecreated As Boolean = False
                    If context.CancellationToken.IsCancellationRequested AndAlso context.RetryCount < context.RetryPolicy.MaxRetryCount Then
                        Dim oldCts = context.CancellationTokenSource
                        Try
                            If cancellationToken = Nothing Then
                                context.CancellationTokenSource = New CancellationTokenSource()
                            Else
                                context.CancellationTokenSource = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)
                            End If
                            context.CancellationToken = context.CancellationTokenSource.Token
                            ctsRecreated = True
                        Finally
                            oldCts?.Dispose()
                        End Try
                    End If

                    Dim perAttemptTimeout = effectiveTimeout
                    If context.RetryCount > 0 AndAlso effectiveTimeout > 60 Then
                        perAttemptTimeout = Math.Max(60, effectiveTimeout - CInt((DateTime.Now - pipelineStartTime).TotalSeconds))
                    End If

                    Dim timeoutTask = Task.Delay(TimeSpan.FromSeconds(perAttemptTimeout), context.CancellationToken)
                    Dim executionTask = ExecuteSkillPipelineAsync(context)

                    Dim completedTask = Await Task.WhenAny(executionTask, timeoutTask)

                    If completedTask Is timeoutTask Then
                        context.Phase = SkillExecutionPhase.TimedOut
                        context.ExecutionEndTime = DateTime.Now
                        context.ErrorMessage = $"Skill 执行超时（超过 {perAttemptTimeout} 秒）：{context.SkillName}"
                        context.LastErrorCategory = SkillErrorCategory.Transient
                        If context.CancellationTokenSource IsNot Nothing Then
                            Try
                                context.CancellationTokenSource.Cancel()
                            Catch ex As Exception  ' 
                            End Try
                        End If

                        RaiseEvent ExecutionTimedOut(context, perAttemptTimeout)
                        WriteDiagnosticLog(context, "TIMEOUT", $"超时 {perAttemptTimeout}s，重试次数={context.RetryCount}/{context.RetryPolicy.MaxRetryCount}")

                        Dim canRetry As Boolean = context.RetryPolicy.RetryOnTimeout AndAlso
                                                  context.RetryCount < context.RetryPolicy.MaxRetryCount AndAlso
                                                  Not cancellationToken.IsCancellationRequested
                        If canRetry Then
                            If Await PrepareForRetry(context, New TimeoutException(context.ErrorMessage)) Then
                                Continue While
                            End If
                        End If

                        Return New WorkflowResult With {
                            .Success = False,
                            .ErrorMessage = context.ErrorMessage,
                            .TotalExecutionTime = context.ExecutionEndTime - pipelineStartTime
                        }
                    End If

                    Try
                        Dim result = Await executionTask
                        WriteDiagnosticLog(context, "SUCCESS", $"耗时 {(DateTime.Now - pipelineStartTime).TotalMilliseconds:F0}ms，重试{context.RetryCount}次")
                        result.TotalExecutionTime = DateTime.Now - pipelineStartTime
                        Return result
                    Catch ex As Exception
                        context.LastException = ex
                        context.LastErrorCategory = ClassifyException(ex)
                        context.Phase = SkillExecutionPhase.Failed
                        context.ExecutionEndTime = DateTime.Now
                        context.ErrorMessage = ex.Message

                        WriteDiagnosticLog(context, "FAILED", $"类型={ex.GetType().Name}，分类={context.LastErrorCategory}，消息={ex.Message}")

                        If context.RetryPolicy.RetryOnTransient AndAlso
                           context.LastErrorCategory = SkillErrorCategory.Transient AndAlso
                           context.RetryCount < context.RetryPolicy.MaxRetryCount AndAlso
                           Not cancellationToken.IsCancellationRequested Then
                            GoTo PrepareRetryLoop
                        End If

                        RaiseEvent ExecutionFailed(context, ex)
                        Dim failedResult = New WorkflowResult With {
                            .Success = False,
                            .ErrorMessage = BuildErrorMessageWithCategory(context, ex),
                            .TotalExecutionTime = DateTime.Now - pipelineStartTime
                        }
                        If failedResult.Metadata Is Nothing Then
                            failedResult.Metadata = New Dictionary(Of String, Object)()
                        End If
                        failedResult.Metadata("error_category") = context.LastErrorCategory.ToString()
                        failedResult.Metadata("retry_count") = context.RetryCount
                        Return failedResult
                    End Try
PrepareRetryLoop:
                    If context.LastException IsNot Nothing Then
                        Dim retryOk = Await PrepareForRetry(context, context.LastException)
                        If retryOk Then
                            Continue While
                        End If
                        RaiseEvent ExecutionFailed(context, context.LastException)
                        Dim retryFailed = New WorkflowResult With {
                            .Success = False,
                            .ErrorMessage = BuildErrorMessageWithCategory(context, context.LastException),
                            .TotalExecutionTime = DateTime.Now - pipelineStartTime
                        }
                        If retryFailed.Metadata Is Nothing Then
                            retryFailed.Metadata = New Dictionary(Of String, Object)()
                        End If
                        retryFailed.Metadata("error_category") = context.LastErrorCategory.ToString()
                        retryFailed.Metadata("retry_count") = context.RetryCount
                        Return retryFailed
                    End If
                End While

            Catch exCx As OperationCanceledException
                context.Phase = SkillExecutionPhase.Failed
                context.ExecutionEndTime = DateTime.Now
                context.ErrorMessage = "Skill 执行已被用户取消：" & context.SkillName
                WriteDiagnosticLog(context, "CANCELLED", exCx.Message)
                RaiseEvent ExecutionFailed(context, exCx)
                Return New WorkflowResult With {
                    .Success = False,
                    .ErrorMessage = context.ErrorMessage,
                    .TotalExecutionTime = context.ExecutionEndTime - pipelineStartTime
                }
            Catch ex As Exception
                context.Phase = SkillExecutionPhase.Failed
                context.ExecutionEndTime = DateTime.Now
                context.LastException = ex
                context.LastErrorCategory = ClassifyException(ex)
                context.ErrorMessage = ex.Message
                WriteDiagnosticLog(context, "FATAL", $"未处理异常: {ex.GetType().Name}: {ex.Message}")
                RaiseEvent ExecutionFailed(context, ex)
                Return New WorkflowResult With {
                    .Success = False,
                    .ErrorMessage = BuildErrorMessageWithCategory(context, ex),
                    .TotalExecutionTime = context.ExecutionEndTime - pipelineStartTime
                }
            End Try

            Return New WorkflowResult With {.Success = False, .ErrorMessage = "技能执行意外终止（不可达代码路径）"}
        End Function

        Private Async Function PrepareForRetry(context As SkillExecutionContext, ex As Exception) As Task(Of Boolean)
            context.RetryCount += 1
            Dim delayMs = context.RetryPolicy.ComputeDelayMs(context.RetryCount)
            Dim previousPhase = context.Phase
            context.Phase = SkillExecutionPhase.Retrying

            Try
                RaiseEvent ExecutionRetrying(context, context.RetryCount, delayMs, ex)
                context.Progress?.Report($"[{context.RetryCount}/{context.RetryPolicy.MaxRetryCount}] 遇到临时错误，{delayMs}ms 后重试：{ex.Message}")
                WriteDiagnosticLog(context, "RETRY", $"第{context.RetryCount}次重试，延迟={delayMs}ms，错误={ex.GetType().Name}")
                Await Task.Delay(delayMs, context.CancellationToken)
                Return True
            Catch cx As OperationCanceledException
                WriteDiagnosticLog(context, "RETRY_ABORT", "重试期间被取消")
                Return False
            End Try
        End Function

        Private Function ClassifyException(ex As Exception) As SkillErrorCategory
            If TypeOf ex Is ArgumentException OrElse
               TypeOf ex Is InvalidOperationException Then
                Dim msg = If(ex.Message, "").ToLowerInvariant()
                If msg.Contains("不存在") OrElse msg.Contains("未找到") OrElse msg.Contains("missing") OrElse
                   msg.Contains("必需参数") OrElse msg.Contains("required") OrElse
                   msg.Contains("配置") OrElse msg.Contains("config") OrElse msg.Contains("apikey") OrElse msg.Contains("key") Then
                    Return SkillErrorCategory.Permanent
                End If
                Return SkillErrorCategory.UserError
            End If

            If TypeOf ex Is System.IO.IOException OrElse
               TypeOf ex Is System.Net.Http.HttpRequestException OrElse
               TypeOf ex Is System.Net.Sockets.SocketException OrElse
               TypeOf ex Is System.Threading.Tasks.TaskCanceledException OrElse
               TypeOf ex Is TimeoutException Then
                Return SkillErrorCategory.Transient
            End If

            If TypeOf ex Is UnauthorizedAccessException OrElse
               TypeOf ex Is System.Security.SecurityException Then
                Return SkillErrorCategory.UserError
            End If

            If TypeOf ex Is OutOfMemoryException OrElse
               TypeOf ex Is StackOverflowException OrElse
               TypeOf ex Is SEHException Then
                Return SkillErrorCategory.SystemCrash
            End If

            If TypeOf ex Is COMException Then
                Dim msg = If(ex.Message, "").ToLowerInvariant()
                If msg.Contains("busy") OrElse msg.Contains("rpc") OrElse msg.Contains("忙") OrElse msg.Contains("重试") Then
                    Return SkillErrorCategory.Transient
                End If
                Return SkillErrorCategory.SystemCrash
            End If

            Dim exType = ex.GetType().Name.ToLowerInvariant()
            If exType.Contains("http") OrElse exType.Contains("network") OrElse exType.Contains("socket") OrElse exType.Contains("timeout") OrElse exType.Contains("transient") Then
                Return SkillErrorCategory.Transient
            End If

            Return SkillErrorCategory.Unknown
        End Function

        Private Function BuildErrorMessageWithCategory(context As SkillExecutionContext, ex As Exception) As String
            Dim categoryLabel As String = ""
            Dim suggestion As String = ""

            Select Case context.LastErrorCategory
                Case SkillErrorCategory.Transient : categoryLabel = "临时错误" : suggestion = "可稍后重试"
                Case SkillErrorCategory.Permanent : categoryLabel = "配置/定义错误" : suggestion = "请检查配置或Skill定义"
                Case SkillErrorCategory.UserError : categoryLabel = "输入错误" : suggestion = "请检查数据或权限后重试"
                Case SkillErrorCategory.SystemCrash : categoryLabel = "系统级错误" : suggestion = "可能需要重启Excel或释放资源"
                Case Else : categoryLabel = "未知错误" : suggestion = "请查看日志详情"
            End Select

            Dim retryInfo As String = ""
            If context.RetryCount > 0 Then
                retryInfo = $"（已自动重试{context.RetryCount}次）"
            End If

            Return $"[{categoryLabel}]{retryInfo} {ex.Message}{IIf(Not String.IsNullOrWhiteSpace(suggestion), " | 建议: " & suggestion, "")}"
        End Function

        Private Async Function ExecuteSkillPipelineAsync(context As SkillExecutionContext) As Task(Of WorkflowResult)
            Await ValidateSkill(context)
            Await ExecuteSkillInternal(context)
            Await PostProcessResult(context)

            context.Phase = SkillExecutionPhase.Completed
            context.ExecutionEndTime = DateTime.Now

            Dim result = BuildSuccessResult(context)
            RaiseEvent ExecutionCompleted(context, result)
            Return result
        End Function

        Private Function CreateExecutionContext(skillName As String, inputs As Dictionary(Of String, Object),
                                                progress As IProgress(Of String),
                                                timeoutSeconds As Integer,
                                                externalToken As CancellationToken,
                                                Optional retryPolicy As SkillRetryPolicy = Nothing) As SkillExecutionContext
            Dim cts As CancellationTokenSource = Nothing

            If externalToken = Nothing Then
                cts = New CancellationTokenSource()
            Else
                cts = CancellationTokenSource.CreateLinkedTokenSource(externalToken)
            End If

            Dim effectiveTimeout = If(timeoutSeconds > 0, timeoutSeconds, _defaultTimeoutSeconds)
            Dim effectivePolicy = If(retryPolicy, SkillRetryPolicy.Default())

            Return New SkillExecutionContext With {
                .SkillName = skillName,
                .Inputs = inputs,
                .Progress = progress,
                .Phase = SkillExecutionPhase.Initializing,
                .ExecutionStartTime = DateTime.Now,
                .TimeoutSeconds = effectiveTimeout,
                .CancellationToken = cts.Token,
                .CancellationTokenSource = cts,
                .RetryPolicy = effectivePolicy
            }
        End Function

        Private Async Function ValidateSkill(context As SkillExecutionContext) As Task
            ChangePhase(context, SkillExecutionPhase.Validating)

            Await Task.Run(Sub()
                Dim skill = _skillManager.GetSkill(context.SkillName)
                If skill Is Nothing Then
                    Throw New InvalidOperationException($"Skill '{context.SkillName}' 不存在")
                End If

                context.Skill = skill
                context.IsComposite = skill.IsComposite
                context.ProcessingMode = skill.ProcessingMode

                If skill.InputParameters IsNot Nothing Then
                    For Each param In skill.InputParameters
                        If param.IsRequired AndAlso (context.Inputs Is Nothing OrElse Not context.Inputs.ContainsKey(param.Name)) Then
                            Throw New ArgumentException($"缺少必需参数: {param.Name}")
                        End If
                    Next
                End If
            End Sub)
        End Function

        Private Async Function ExecuteSkillInternal(context As SkillExecutionContext) As Task
            ChangePhase(context, SkillExecutionPhase.Executing)

            Select Case context.Skill.ProcessingMode
                Case SkillProcessingMode.MultimodalLLM
                    Await ExecuteMultimodalSkill(context)
                Case SkillProcessingMode.LocalOCR
                    Await ExecuteLocalOCRSkill(context)
                Case SkillProcessingMode.Hybrid
                    Await ExecuteHybridSkill(context)
                Case SkillProcessingMode.PythonCode
                    Await ExecutePythonSkill(context)
                Case SkillProcessingMode.VBACode
                    Await ExecuteVBASkill(context)
                Case SkillProcessingMode.SkillSpec
                    ' SkillSpec 由前端 skill-pipeline 的 Spec 引擎处理（LLM 规划 + 确定性渲染 VBA）。
                    ' 后端不应执行占位逻辑；若被 LLM 直连调用（text-pipeline 兜底路径），返回明确提示，
                    ' 避免泄露"已识别，等待 AI 生成响应"占位 JSON。
                    Await ExecuteSkillSpecSkill(context)
                Case Else
                    If context.IsComposite Then
                        Await ExecuteCompositeSkill(context)
                    Else
                        Await ExecuteSimpleSkill(context)
                    End If
            End Select
        End Function

        Private Async Function ExecuteMultimodalSkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行多模态 Skill: {context.Skill.DisplayName}")

            Dim result = Await _skillManager.ExecuteSkill(context.SkillName, context.Inputs, context.Progress)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecuteLocalOCRSkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行本地 OCR Skill: {context.Skill.DisplayName}")

            Dim result = Await _skillManager.ExecuteSkill(context.SkillName, context.Inputs, context.Progress)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecuteHybridSkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行混合 Skill: {context.Skill.DisplayName}")

            Dim executor As New CustomSkillExecutor()
            Dim result = Await executor.ExecuteCustomSkill(context.Skill, context.Inputs)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecutePythonSkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行 Python Skill: {context.Skill.DisplayName}")

            Dim executor As New CustomSkillExecutor()
            Dim result = Await executor.ExecuteCustomSkill(context.Skill, context.Inputs)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecuteVBASkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行 VBA Skill: {context.Skill.DisplayName}")

            Dim executor As New CustomSkillExecutor()
            Dim result = Await executor.ExecuteCustomSkill(context.Skill, context.Inputs)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecuteCompositeSkill(context As SkillExecutionContext) As Task
            context.Progress?.Report($"正在执行复合 Skill 工作流: {context.Skill.DisplayName}")

            If _workflowEngine Is Nothing Then
                _workflowEngine = New WorkflowEngine(_skillManager)
            End If

            Dim result = Await _workflowEngine.ExecuteWorkflow(context.Skill, context.Inputs, context.Progress)
            context.Inputs("__executionResult") = result
        End Function

        Private Async Function ExecuteSimpleSkill(context As SkillExecutionContext) As Task
            Await Task.Run(Sub()
                Dim result As New WorkflowResult

                If context.Skill.IsGeneric Then
                    result.Success = True
                    result.Outputs("message") = $"Skill [{context.Skill.DisplayName}] 已识别，等待 AI 生成响应"
                Else
                    result.Success = False
                    result.ErrorMessage = "非通用 Skill 缺少工作流定义"
                End If

                context.Inputs("__executionResult") = result
            End Sub)
        End Function

        ''' <summary>
        ''' SkillSpec 模式的后端兜底处理。
        ''' SkillSpec 由前端 Spec 引擎规划并渲染 VBA 执行，后端仅透传该枚举；
        ''' 若被 LLM 直连调用（text-pipeline 兜底），返回明确提示而非占位 JSON。
        ''' </summary>
        Private Async Function ExecuteSkillSpecSkill(context As SkillExecutionContext) As Task
            Await Task.Run(Sub()
                Dim result As New WorkflowResult
                result.Success = True
                result.Outputs("message") = $"Skill [{context.Skill.DisplayName}] 由前端 Spec 引擎处理，请直接输入指令触发规划执行"
                context.Inputs("__executionResult") = result
            End Sub)
        End Function

        Private Async Function PostProcessResult(context As SkillExecutionContext) As Task
            ChangePhase(context, SkillExecutionPhase.PostProcessing)

            Await Task.Run(Sub()
                Dim result As WorkflowResult = Nothing
                If context.Inputs.ContainsKey("__executionResult") Then
                    result = CType(context.Inputs("__executionResult"), WorkflowResult)
                End If

                If result IsNot Nothing AndAlso result.Success Then
                    If context.Skill.Name = "SkillCreator" AndAlso context.ProcessingMode = SkillProcessingMode.MultimodalLLM Then
                        ProcessGeneratedSkill(result)
                    End If
                End If
            End Sub)
        End Function

        Private Sub ProcessGeneratedSkill(result As WorkflowResult)
            Try
                ' 如果 SkillManager.vb 已经成功处理，直接使用已有结果，不再重复处理
                If result.Outputs.ContainsKey("generatedSkillJson") AndAlso
                   result.Outputs("generatedSkillJson") IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(result.Outputs("generatedSkillJson").ToString()) Then
                    System.Diagnostics.Debug.WriteLine("[SkillExecutionRouter] SkillManager 已处理完成，跳过重复后处理")
                    Return
                End If

                Dim rawContent As String = ""
                If result.Outputs.ContainsKey("content") AndAlso result.Outputs("content") IsNot Nothing Then
                    rawContent = result.Outputs("content").ToString()
                End If

                If String.IsNullOrWhiteSpace(rawContent) Then
                    result.Success = False
                    result.ErrorMessage = "LLM 返回内容为空，无法生成 Skill"
                    Return
                End If

                ' 调试：将 LLM 原始输出写入临时日志，便于排查
                Try
                    Dim debugDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelAIVSTOAddIn_Debug")
                    If Not System.IO.Directory.Exists(debugDir) Then System.IO.Directory.CreateDirectory(debugDir)
                    Dim debugFile = System.IO.Path.Combine(debugDir, $"SkillCreator_Raw_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 6)}.txt")
                    System.IO.File.WriteAllText(debugFile, $"长度: {rawContent.Length}{vbCrLf}{vbCrLf}原始内容:{vbCrLf}{rawContent}{vbCrLf}")
                    System.Diagnostics.Debug.WriteLine($"[SkillCreator] 原始输出已保存: {debugFile}")
                Catch logEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[SkillCreator] 保存调试日志失败: {logEx.Message}")
                End Try

                ' 如果 content 已经是 HTML（SkillManager 可能已处理），直接返回
                If System.Text.RegularExpressions.Regex.IsMatch(rawContent, "<[a-z][\s\S]*>", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                    System.Diagnostics.Debug.WriteLine("[SkillExecutionRouter] content 已是 HTML，不再重复处理")
                    Return
                End If

                Dim jsonText = ExtractJsonFromLLMResponse(rawContent)
                If String.IsNullOrWhiteSpace(jsonText) Then
                    result.Success = False
                    result.ErrorMessage = "无法从 LLM 响应中提取有效的 SkillDefinition JSON"
                    Return
                End If

                ' 尝试多种方式解析 Skill
                Dim newSkill As SkillDefinition = Nothing

                ' 方式1: Function Calling 格式
                newSkill = TryParseFunctionCallSkill(jsonText)

                ' 方式2: 直接反序列化（处理嵌套 JSON）
                If newSkill Is Nothing Then
                    newSkill = TryDeserializeSkillFlexibly(jsonText)
                End If

                ' 方式3: 从原始文本中递归搜索包含 name 字段的嵌套 JSON
                If newSkill Is Nothing OrElse String.IsNullOrWhiteSpace(newSkill.Name) Then
                    newSkill = TryExtractSkillFromNestedJson(rawContent)
                End If

                If newSkill Is Nothing Then
                    result.Success = False
                    result.ErrorMessage = "JSON 反序列化失败，SkillDefinition 为 Nothing"
                    Return
                End If

                If String.IsNullOrWhiteSpace(newSkill.Name) Then
                    result.Success = False
                    result.ErrorMessage = "生成的 Skill 缺少 name 字段"
                    Return
                End If

                If String.IsNullOrWhiteSpace(newSkill.DisplayName) Then
                    newSkill.DisplayName = newSkill.Name
                End If

                newSkill.IsCustom = True
                newSkill.CreatedDate = DateTime.Now
                newSkill.LastModifiedDate = DateTime.Now
                newSkill.Version = 1
                newSkill.UsageCount = 0
                newSkill.CreatedBy = "SkillCreator"

                result.Outputs("generatedSkillName") = newSkill.Name
                result.Outputs("generatedSkillDisplayName") = newSkill.DisplayName
                result.Outputs("saved") = False
                ' 使用精简序列化，仅包含必要字段
                Dim skillJsonStr = SkillJsonHelper.SerializeSkillCompact(newSkill)
                result.Outputs("generatedSkillJson") = skillJsonStr

                ' content 设为结构化 JSON（含 skill 数据和显示信息），前端根据类型渲染
                Dim contentObj As New Dictionary(Of String, Object) From {
                    {"type", "skill_card"},
                    {"skillName", newSkill.Name},
                    {"skillDisplayName", newSkill.DisplayName},
                    {"skillDescription", If(newSkill.Description, "")},
                    {"skillCategory", If(newSkill.Category, "")},
                    {"processingMode", newSkill.ProcessingMode.ToDisplayName()},
                    {"saved", False},
                    {"skillJson", skillJsonStr}
                }
                result.Outputs("content") = Newtonsoft.Json.JsonConvert.SerializeObject(contentObj)

            Catch ex As Exception
                result.Success = False
                result.ErrorMessage = "生成 Skill 后处理失败: " & ex.Message
            End Try
        End Sub

        ''' <summary>
        ''' 从 LLM 输出中容错提取合法 JSON（处理 thinking 标签、代码块包裹、正文混杂、
        ''' 平衡括号、常见污染）。Friend Shared 供 SkillManager 复用——避免 SkillManager 用
        ''' 裸 JObject.Parse 对同一份输出做严格解析而在包裹/污染时误判"未返回结构化 JSON"。
        ''' </summary>
        Friend Shared Function ExtractJsonFromLLMResponse(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return Nothing

            text = text.Trim()

            ' 先清除 <thinking>...</thinking> 标签内容
            text = System.Text.RegularExpressions.Regex.Replace(text, "<thinking>[\s\S]*?</thinking>", "")
            ' 清除 <tool_call>...</tool_call> 等可能包裹的标签
            text = System.Text.RegularExpressions.Regex.Replace(text, "<[^>]+>[\s\S]*?</[^>]+>", "")
            text = text.Trim()

            ' 策略1: 尝试直接解析整个文本
            Try
                Dim testObj = Newtonsoft.Json.Linq.JObject.Parse(text)
                If testObj IsNot Nothing Then Return text
            Catch ex As Exception  ' 
            End Try

            ' 策略2: 匹配 ```json ... ``` 或 ``` ... ``` 代码块
            Dim codeBlockMatch = System.Text.RegularExpressions.Regex.Match(text, "```(?:json)?\s*([\s\S]*?)\s*```", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
            If codeBlockMatch.Success AndAlso codeBlockMatch.Groups.Count > 1 Then
                Dim candidate = SanitizeJsonString(codeBlockMatch.Groups(1).Value.Trim())
                Dim valid = TryParseJObject(candidate)
                If valid Then Return candidate
            End If

            ' 策略3: 平衡括号匹配
            Dim braceMatch = ExtractBalancedBraces(text)
            If Not String.IsNullOrWhiteSpace(braceMatch) Then
                Dim candidate = SanitizeJsonString(braceMatch)
                Dim valid = TryParseJObject(candidate)
                If valid Then Return candidate
            End If

            ' 策略4: 扫描文本中所有可能的 { ... } 片段，取最大的可解析片段
            Dim allMatches = FindAllBraceCandidates(text)
            For Each candidate In allMatches
                Dim sanitized = SanitizeJsonString(candidate)
                If TryParseJObject(sanitized) Then
                    Return sanitized
                End If
            Next

            ' 策略5: 包含 Skill 关键字段的 JSON
            Dim skillMatch = ExtractSkillJsonCandidate(text)
            If Not String.IsNullOrWhiteSpace(skillMatch) Then
                Return skillMatch
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 清理 LLM 输出中常见的 JSON 格式污染
        ''' </summary>
        Private Shared Function SanitizeJsonString(jsonText As String) As String
            If String.IsNullOrWhiteSpace(jsonText) Then Return jsonText

            Dim result = jsonText.Trim()

            ' 移除 BOM 和零宽字符
            result = New String(result.Where(Function(c) Not Char.IsControl(c) OrElse c = vbCr Or c = vbLf Or c = vbTab).ToArray())

            ' 修复 LLM 常见错误：中文引号替换成英文引号
            result = result.Replace(""""c, """"c).Replace(""""c, """"c)

            ' 修复多余的逗号：",}" → "}"
            result = System.Text.RegularExpressions.Regex.Replace(result, ",\s*}", "}")
            result = System.Text.RegularExpressions.Regex.Replace(result, ",\s*]", "]")

            Return result
        End Function

        ''' <summary>
        ''' 扫描文本中所有可能的 JSON 对象候选（优先返回最长可解析的）
        ''' </summary>
        Private Shared Function FindAllBraceCandidates(text As String) As List(Of String)
            Dim candidates As New List(Of String)()
            Dim idx As Integer = 0

            While idx < text.Length
                Dim braceStart = text.IndexOf("{"c, idx)
                If braceStart < 0 Then Exit While

                Dim balanced = ExtractBalancedBraces(text.Substring(braceStart))
                If Not String.IsNullOrWhiteSpace(balanced) Then
                    candidates.Add(balanced)
                    idx = braceStart + balanced.Length
                Else
                    idx = braceStart + 1
                End If
            End While

            ' 按长度降序，优先尝试大的完整 JSON
            Return candidates.OrderByDescending(Function(s) s.Length).ToList()
        End Function

        Private Shared Function TryParseJObject(jsonText As String) As Boolean
            If String.IsNullOrWhiteSpace(jsonText) Then Return False
            Try
                Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                Return True
            Catch ex As Exception  ' 
                Return False
            End Try
        End Function

        Private Shared Function ExtractBalancedBraces(text As String) As String
            Dim startIdx = text.IndexOf("{"c)
            If startIdx < 0 Then Return Nothing

            Dim depth As Integer = 0
            Dim inString As Boolean = False
            Dim escape As Boolean = False

            For i As Integer = startIdx To text.Length - 1
                Dim ch = text(i)

                If escape Then
                    escape = False
                    Continue For
                End If

                If ch = "\"c Then
                    escape = True
                    Continue For
                End If

                If ch = """"c Then
                    inString = Not inString
                    Continue For
                End If

                If Not inString Then
                    If ch = "{"c Then
                        depth += 1
                    ElseIf ch = "}"c Then
                        depth -= 1
                        If depth = 0 Then
                            Return text.Substring(startIdx, i - startIdx + 1)
                        End If
                    End If
                End If
            Next

            Return Nothing
        End Function

        Private Shared Function ExtractSkillJsonCandidate(text As String) As String
            Dim requiredFields = {"name", "displayName", "description", "examples", "keywords", "category", "processingMode", "outputFormat"}

            Dim braceRegex As New System.Text.RegularExpressions.Regex("\{[\s\S]*?\}")
            Dim matches = braceRegex.Matches(text)

            For Each match As System.Text.RegularExpressions.Match In matches
                Dim candidate = match.Value.Trim()
                If Not TryParseJObject(candidate) Then Continue For

                Try
                    Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(candidate)
                    If jObj Is Nothing Then Continue For

                    Dim fieldCount As Integer = 0
                    For Each field In requiredFields
                        If jObj.ContainsKey(field) Then fieldCount += 1
                    Next

                    If fieldCount >= 3 Then
                        Return candidate
                    End If
                Catch ex As Exception  ' 
                End Try
            Next

            Return Nothing
        End Function

        Private Function TryParseFunctionCallSkill(jsonText As String) As SkillDefinition
            Try
                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
                If jObj Is Nothing OrElse Not jObj.ContainsKey("function_call") Then Return Nothing

                Dim fc = jObj("function_call")
                If fc Is Nothing Then Return Nothing

                Dim funcName = fc("name")?.ToString()
                If String.IsNullOrWhiteSpace(funcName) Then Return Nothing

                Dim args = fc("arguments")
                Dim argsObj = TryCast(args, Newtonsoft.Json.Linq.JObject)
                If argsObj Is Nothing Then Return Nothing

                Dim skill As New SkillDefinition()

                skill.Name = argsObj("name")?.ToString()
                skill.DisplayName = argsObj("displayName")?.ToString()
                skill.Description = argsObj("description")?.ToString()
                skill.Category = argsObj("category")?.ToString()

                Dim kwArray = argsObj("keywords")
                If kwArray IsNot Nothing AndAlso kwArray.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    skill.Keywords = kwArray.Select(Function(t) t.ToString()).ToArray()
                End If

                Dim exArray = argsObj("examples")
                If exArray IsNot Nothing AndAlso exArray.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    skill.Examples = exArray.Select(Function(t) t.ToString()).ToArray()
                End If

                Dim tagArray = argsObj("tags")
                If tagArray IsNot Nothing AndAlso tagArray.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    skill.Tags = tagArray.Select(Function(t) t.ToString()).ToList()
                End If

                Dim supportsImage As Boolean = False
                Dim supportsFile As Boolean = False
                Dim imgVal = argsObj("supportsImageInput")
                If imgVal IsNot Nothing Then
                    If imgVal.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                        supportsImage = imgVal.ToObject(Of Boolean)()
                    ElseIf imgVal.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        supportsImage = imgVal.ToObject(Of Integer)() <> 0
                    Else
                        supportsImage = imgVal.ToString().Trim().Equals("true", StringComparison.OrdinalIgnoreCase)
                    End If
                End If
                Dim fileVal = argsObj("supportsFileInput")
                If fileVal IsNot Nothing Then
                    If fileVal.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                        supportsFile = fileVal.ToObject(Of Boolean)()
                    ElseIf fileVal.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                        supportsFile = fileVal.ToObject(Of Integer)() <> 0
                    Else
                        supportsFile = fileVal.ToString().Trim().Equals("true", StringComparison.OrdinalIgnoreCase)
                    End If
                End If

                Select Case funcName.ToLower()
                    Case "create_vba_skill"
                        skill.ProcessingMode = SkillProcessingMode.VBACode
                        skill.OutputFormat = "vba"
                        skill.TemplateCode = argsObj("templateCode")?.ToString()
                        skill.SupportsImageInput = supportsImage
                        skill.SupportsFileInput = supportsFile
                    Case "create_llm_skill"
                        skill.ProcessingMode = SkillProcessingMode.MultimodalLLM
                        Dim outputFormat = argsObj("outputFormat")?.ToString()
                        skill.OutputFormat = If(String.IsNullOrWhiteSpace(outputFormat), "markdown", outputFormat)
                        skill.LLMSystemPrompt = argsObj("llmSystemPrompt")?.ToString()
                        skill.SupportsImageInput = supportsImage
                        skill.SupportsFileInput = supportsFile
                    Case "create_hybrid_skill"
                        skill.ProcessingMode = SkillProcessingMode.Hybrid
                        Dim outputFormat = argsObj("outputFormat")?.ToString()
                        skill.OutputFormat = If(String.IsNullOrWhiteSpace(outputFormat), "vba", outputFormat)
                        skill.LLMSystemPrompt = argsObj("llmSystemPrompt")?.ToString()
                        skill.TemplateCode = argsObj("templateCode")?.ToString()
                        skill.SupportsImageInput = supportsImage
                        skill.SupportsFileInput = supportsFile
                End Select

                skill.IsGeneric = False
                skill.NeedsConfig = False
                skill.IsComposite = False
                skill.IsCustom = True
                skill.Version = 1
                skill.UsageCount = 0
                skill.CreatedBy = "SkillCreator"
                skill.CreatedDate = DateTime.Now
                skill.LastModifiedDate = DateTime.Now

                Return skill
        Catch ex As Exception  ' 
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 灵活反序列化 Skill（处理嵌套 JSON）
    ''' 【修复】使用 SkillJsonHelper 进行不区分大小写的反序列化
    ''' </summary>
    Private Function TryDeserializeSkillFlexibly(jsonText As String) As SkillDefinition
        Try
            Dim normalizedJson = NormalizeJsonForSkill(jsonText)

            ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
            Dim skill = SkillJsonHelper.DeserializeSkill(normalizedJson)

            If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                System.Diagnostics.Debug.WriteLine("[SkillCreator] 通过 SkillJsonHelper 反序列化成功: " & skill.Name)
                Return skill
            End If

            ' 反序列化成功但 name 为空 — 可能是嵌套结构，尝试逐层解包
            System.Diagnostics.Debug.WriteLine("[SkillCreator] SkillJsonHelper 返回的 Skill 名称为空，尝试解包嵌套结构")
            Try
                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(normalizedJson)
                Dim unwrapped = UnwrapNestedSkill(jObj)
                If unwrapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(unwrapped.Name) Then
                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 解包嵌套结构成功: " & unwrapped.Name)
                    Return unwrapped
                End If

                ' 尝试映射 LLM 使用的替代字段名 → 标准字段名
                Dim mapped = TryMapAlternativeNames(jObj)
                If mapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(mapped.Name) Then
                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 映射替代字段名成功: " & mapped.Name)
                    Return mapped
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[SkillCreator] 解包嵌套结构异常: " & ex.Message)
            End Try

            System.Diagnostics.Debug.WriteLine("[SkillCreator] 所有反序列化方式均失败")
            Return Nothing
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[SkillCreator] TryDeserializeSkillFlexibly 异常: " & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 尝试将 LLM 常用的替代字段名映射到 SkillDefinition 标准字段名
    ''' 【已修复】使用 PascalCase 属性名以匹配 SkillDefinition 类的属性
    ''' </summary>
    Private Function TryMapAlternativeNames(jObj As Newtonsoft.Json.Linq.JObject) As SkillDefinition
        Try
            ' 常见替代字段名映射表
            ' 格式: { PascalCase属性名, { 替代名称列表 } }
            Dim fieldMappings As New Dictionary(Of String, String()) From {
                {"Name", {"skillName", "SkillName", "skill_name", "name_id", "name", "Name"}},
                {"DisplayName", {"display_name", "displayName", "DisplayName", "Display_Name", "title", "Title"}},
                {"Description", {"description", "Description", "desc", "Desc", "function_description", "FunctionDescription"}},
                {"ProcessingMode", {"processing_mode", "processingMode", "ProcessingMode", "mode", "Mode", "type", "Type"}},
                {"TemplateCode", {"template_code", "templateCode", "TemplateCode", "code", "Code", "vba_code", "VbaCode"}},
                {"LLMSystemPrompt", {"llm_system_prompt", "llmSystemPrompt", "LlmSystemPrompt", "system_prompt", "SystemPrompt", "prompt", "Prompt"}},
                {"Keywords", {"keywords", "Keywords", "keyword", "Keyword"}},
                {"Examples", {"examples", "Examples", "example", "Example"}},
                {"Category", {"category", "Category", "categories", "Categories"}},
                {"Tags", {"tags", "Tags", "tag", "Tag"}},
                {"OutputFormat", {"output_format", "outputFormat", "OutputFormat", "format", "Format"}},
                {"IsGeneric", {"is_generic", "isGeneric", "IsGeneric", "generic", "Generic"}},
                {"NeedsConfig", {"needs_config", "needsConfig", "NeedsConfig", "needsConfiguration", "NeedsConfiguration"}},
                {"ConfigFields", {"config_fields", "configFields", "ConfigFields", "configuration", "Configuration", "inputParameters", "input_parameters", "InputParameters"}},
                {"SupportsImageInput", {"supports_image_input", "supportsImageInput", "SupportsImageInput", "imageInput", "ImageInput"}},
                {"SupportsFileInput", {"supports_file_input", "supportsFileInput", "SupportsFileInput", "fileInput", "FileInput"}},
                {"FieldMappings", {"field_mappings", "fieldMappings", "FieldMappings", "mappings", "Mappings"}}
            }

            ' 构建标准 JSON 对象（使用 PascalCase 属性名匹配 SkillDefinition）
            Dim standardJson As New Dictionary(Of String, Object)()

            For Each kvp In fieldMappings
                Dim standardKey = kvp.Key
                Dim alternativeNames = kvp.Value

                Dim foundValue As Newtonsoft.Json.Linq.JToken = Nothing
                For Each altName In alternativeNames
                    If jObj.TryGetValue(altName, StringComparison.OrdinalIgnoreCase, foundValue) Then
                        Exit For
                    End If
                Next

                If foundValue IsNot Nothing Then
                    Select Case standardKey
                        Case "Name", "DisplayName", "Description", "TemplateCode",
                             "LLMSystemPrompt", "Category", "OutputFormat"
                            standardJson(standardKey) = foundValue.ToString()
                        Case "ProcessingMode"
                            ' 转换 processingMode 为枚举整数值
                            Dim modeInt As Integer
                            If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then
                                modeInt = CInt(foundValue)
                            ElseIf foundValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                Dim modeStr = foundValue.ToString().Trim()
                                If Integer.TryParse(modeStr, modeInt) Then
                                    ' 成功解析为整数
                                Else
                                    ' 尝试按枚举名称解析
                                    Dim parsedMode As SkillProcessingMode
                                    If [Enum].TryParse(modeStr, True, parsedMode) Then
                                        modeInt = CInt(parsedMode)
                                    Else
                                        modeInt = 0 ' 默认为 VBACode
                                    End If
                                End If
                            Else
                                modeInt = 0
                            End If
                            standardJson(standardKey) = modeInt
                        Case "IsGeneric", "NeedsConfig", "SupportsImageInput", "SupportsFileInput"
                            If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then
                                standardJson(standardKey) = CBool(foundValue)
                            Else
                                Dim strVal = foundValue.ToString().ToLowerInvariant()
                                standardJson(standardKey) = (strVal = "true" OrElse strVal = "yes")
                            End If
                        Case "Keywords", "Examples", "Tags", "ConfigFields", "FieldMappings"
                            If foundValue.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                                standardJson(standardKey) = foundValue.ToObject(Of List(Of String))()
                            ElseIf foundValue.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                                standardJson(standardKey) = New List(Of String) From {foundValue.ToString()}
                            Else
                                standardJson(standardKey) = New List(Of String)()
                            End If
                    End Select
                End If
            Next

            ' === 特殊处理：LLM 常用但结构不同的替代字段 ===

            ' 1. actions → TemplateCode：从 actions 数组的第一个元素提取 code 字段
            If Not standardJson.ContainsKey("TemplateCode") Then
                Dim actionsToken As Newtonsoft.Json.Linq.JToken = Nothing
                If jObj.TryGetValue("actions", StringComparison.OrdinalIgnoreCase, actionsToken) AndAlso
                   actionsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    Dim actionsArray = CType(actionsToken, Newtonsoft.Json.Linq.JArray)
                    If actionsArray.Count > 0 AndAlso actionsArray(0).Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                        Dim firstAction = CType(actionsArray(0), Newtonsoft.Json.Linq.JObject)
                        Dim codeToken As Newtonsoft.Json.Linq.JToken = Nothing
                        If firstAction.TryGetValue("code", StringComparison.OrdinalIgnoreCase, codeToken) Then
                            standardJson("TemplateCode") = codeToken.ToString()
                        End If
                    End If
                End If
            End If

            ' 2. inputParameters → ConfigFields：从 inputParameters 数组提取各元素的 name 字段
            If Not standardJson.ContainsKey("ConfigFields") OrElse
               (TypeOf standardJson("ConfigFields") Is List(Of String) AndAlso
                DirectCast(standardJson("ConfigFields"), List(Of String)).Count = 0) Then
                Dim inputParamsToken As Newtonsoft.Json.Linq.JToken = Nothing
                If jObj.TryGetValue("inputParameters", StringComparison.OrdinalIgnoreCase, inputParamsToken) AndAlso
                   inputParamsToken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    Dim paramsArray = CType(inputParamsToken, Newtonsoft.Json.Linq.JArray)
                    Dim paramNames As New List(Of String)()
                    For Each param In paramsArray
                        If param.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                            Dim paramObj = CType(param, Newtonsoft.Json.Linq.JObject)
                            Dim nameToken As Newtonsoft.Json.Linq.JToken = Nothing
                            If paramObj.TryGetValue("name", StringComparison.OrdinalIgnoreCase, nameToken) Then
                                paramNames.Add(nameToken.ToString())
                            End If
                        End If
                    Next
                    If paramNames.Count > 0 Then
                        standardJson("ConfigFields") = paramNames
                    End If
                End If
            End If

            ' 如果连 Name 都没有，无法构建有效 Skill
            If Not standardJson.ContainsKey("Name") OrElse String.IsNullOrWhiteSpace(standardJson("Name").ToString()) Then
                Return Nothing
            End If

            ' 从 standardJson 序列化再反序列化为 SkillDefinition
            ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
            Dim mappedJsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(standardJson)
            Dim mappedSkill = SkillJsonHelper.DeserializeSkill(mappedJsonStr)
            Return mappedSkill
        Catch ex As Exception  ' 
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 从原始文本中递归搜索包含 Skill 必需字段的 JSON
    ''' </summary>
    Private Function TryExtractSkillFromNestedJson(rawText As String) As SkillDefinition
        If String.IsNullOrWhiteSpace(rawText) Then Return Nothing

        Dim cleaned = System.Text.RegularExpressions.Regex.Replace(rawText, "<thinking>[\s\S]*?</thinking>", "")

        Dim candidates As New List(Of String)()
        Dim idx As Integer = 0
        While idx < cleaned.Length
            Dim braceStart = cleaned.IndexOf("{"c, idx)
            If braceStart < 0 Then Exit While

            Dim balanced = ExtractBalancedBraces(cleaned.Substring(braceStart))
            If Not String.IsNullOrWhiteSpace(balanced) Then
                candidates.Add(balanced)
                idx = braceStart + balanced.Length
            Else
                idx = braceStart + 1
            End If
        End While

        For Each candidate In candidates
            Try
                Dim normalized = NormalizeJsonForSkill(candidate)
                ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                Dim skill = SkillJsonHelper.DeserializeSkill(normalized)
                If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                    Return skill
                End If

                Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(normalized)
                Dim unwrapped = UnwrapNestedSkill(jObj)
                If unwrapped IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(unwrapped.Name) Then
                    Return unwrapped
                End If
            Catch ex As Exception  ' 
            End Try
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 逐层解包嵌套 JSON 对象，找到包含 name 字段的那一层
    ''' 【修复】使用 SkillJsonHelper 进行不区分大小写的反序列化
    ''' </summary>
    Private Function UnwrapNestedSkill(jObj As Newtonsoft.Json.Linq.JObject) As SkillDefinition
        If jObj Is Nothing Then Return Nothing

        If jObj.ContainsKey("name") AndAlso Not String.IsNullOrWhiteSpace(jObj("name")?.ToString()) Then
            Try
                Dim normalized = NormalizeJsonForSkill(jObj.ToString(Newtonsoft.Json.Formatting.None))
                ' 使用 SkillJsonHelper 进行不区分大小写的反序列化
                Dim skill = SkillJsonHelper.DeserializeSkill(normalized)
                If skill IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(skill.Name) Then
                    Return skill
                End If
            Catch ex As Exception  ' 
            End Try
        End If

        For Each prop In jObj.Properties()
            If prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Dim childObj = CType(prop.Value, Newtonsoft.Json.Linq.JObject)
                Dim result = UnwrapNestedSkill(childObj)
                If result IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(result.Name) Then
                    Return result
                End If
            End If
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' 标准化 LLM 生成的 Skill JSON，修复常见类型问题：
    ''' 1. 数组字段（examples/keywords/tags）被错误生成为对象或字符串
    ''' 2. 整数字段（version/processingMode）被错误生成为字符串（如 "1.0"）
    ''' 3. 布尔字段被错误生成为字符串
    ''' </summary>
    Private Function NormalizeJsonForSkill(jsonText As String) As String
        Try
            If String.IsNullOrWhiteSpace(jsonText) Then Return jsonText

            Dim jObj = Newtonsoft.Json.Linq.JObject.Parse(jsonText)
            Dim changed As Boolean = False

            ' 1. 修复数组字段
            Dim arrayFields = {"examples", "keywords", "tags", "configFields", "inputContext"}
            For Each fieldName In arrayFields
                Dim token = jObj(fieldName)
                If token Is Nothing Then Continue For

                If token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then Continue For

                Dim newArray As New Newtonsoft.Json.Linq.JArray()
                If token.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                    For Each prop In CType(token, Newtonsoft.Json.Linq.JObject).Properties()
                        Dim val = prop.Value?.ToString()
                        If Not String.IsNullOrWhiteSpace(val) Then newArray.Add(val)
                    Next
                    changed = True
                ElseIf token.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    Dim strVal = token.ToString()
                    If Not String.IsNullOrWhiteSpace(strVal) Then newArray.Add(strVal)
                    changed = True
                ElseIf token.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then
                    newArray.Add(token.ToString())
                    changed = True
                Else
                    changed = True
                End If

                If changed Then jObj(fieldName) = newArray
            Next

            ' 2. 修复整数字段（如 version: "1.0" → version: 1）
            Dim intFields = {"version", "processingMode"}
            For Each fieldName In intFields
                Dim token = jObj(fieldName)
                If token Is Nothing Then Continue For

                If token.Type = Newtonsoft.Json.Linq.JTokenType.Integer Then Continue For

                Dim intVal As Integer
                Dim strRep = token?.ToString()?.Trim()
                If Integer.TryParse(strRep, intVal) Then
                    jObj(fieldName) = New Newtonsoft.Json.Linq.JValue(intVal)
                    changed = True
                ElseIf Decimal.TryParse(strRep, Nothing) Then
                    Dim decVal = Decimal.Parse(strRep)
                    intVal = CInt(Math.Round(decVal))
                    jObj(fieldName) = New Newtonsoft.Json.Linq.JValue(intVal)
                    changed = True
                End If
            Next

            ' 3. 修复布尔字段
            Dim boolFields = {"supportsImageInput", "supportsFileInput", "isGeneric", "needsConfig", "isCustom", "isComposite"}
            For Each fieldName In boolFields
                Dim token = jObj(fieldName)
                If token Is Nothing Then Continue For

                If token.Type = Newtonsoft.Json.Linq.JTokenType.Boolean Then Continue For

                Dim strVal = token?.ToString()?.Trim()
                If strVal IsNot Nothing Then
                    Dim boolVal As Boolean
                    If Boolean.TryParse(strVal, boolVal) Then
                        jObj(fieldName) = New Newtonsoft.Json.Linq.JValue(boolVal)
                        changed = True
                    End If
                End If
            Next

            If changed Then
                Return jObj.ToString(Newtonsoft.Json.Formatting.None)
            End If

            Return jsonText
        Catch ex As Exception  ' 
            Return jsonText
        End Try
    End Function

    Private Function BuildSuccessResult(context As SkillExecutionContext) As WorkflowResult
            Dim result As WorkflowResult = Nothing
            If context.Inputs.ContainsKey("__executionResult") Then
                result = CType(context.Inputs("__executionResult"), WorkflowResult)
            End If

            If result Is Nothing Then
                result = New WorkflowResult With {
                    .Success = True,
                    .Message = $"Skill '{context.SkillName}' 执行完成"
                }
            End If

            result.TotalExecutionTime = context.ExecutionEndTime - context.ExecutionStartTime
            Return result
        End Function

        Private Sub ChangePhase(context As SkillExecutionContext, newPhase As SkillExecutionPhase)
            Dim previousPhase = context.Phase
            context.Phase = newPhase
            RaiseEvent PhaseChanged(context, previousPhase)
        End Sub

#Region "诊断日志"

        Public Function GetDiagnosticLogs() As List(Of SkillDiagnosticEntry)
            SyncLock _diagnosticSyncLock
                Return New List(Of SkillDiagnosticEntry)(_diagnosticLog)
            End SyncLock
        End Function

        Public Sub ClearDiagnosticLogs()
            SyncLock _diagnosticSyncLock
                _diagnosticLog.Clear()
            End SyncLock
        End Sub

        Private Sub WriteDiagnosticLog(context As SkillExecutionContext, level As String, message As String)
            Dim entry As New SkillDiagnosticEntry() With {
                .Timestamp = DateTime.Now,
                .SkillName = If(context IsNot Nothing, context.SkillName, "(null)"),
                .Phase = If(context IsNot Nothing, context.Phase.ToString(), ""),
                .RetryCount = If(context IsNot Nothing, context.RetryCount, 0),
                .Level = level,
                .Message = If(message, ""),
                .ThreadId = Thread.CurrentThread.ManagedThreadId
            }
            SyncLock _diagnosticSyncLock
                _diagnosticLog.Add(entry)
                While _diagnosticLog.Count > MAX_DIAGNOSTIC_ENTRIES
                    _diagnosticLog.RemoveAt(0)
                End While
            End SyncLock

            Dim traceMsg = $"[SkillDiagnostic|{level}|{entry.SkillName}] {entry.Message}"
            System.Diagnostics.Debug.WriteLine(traceMsg)
        End Sub

        Public Function FlushDiagnosticLogsToTempFile() As String
            Try
                Dim entries = GetDiagnosticLogs()
                If entries.Count = 0 Then Return ""
                Dim tempPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    $"SkillDiag_{DateTime.Now:yyyyMMdd_HHmmss}.json")
                Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(entries, Newtonsoft.Json.Formatting.Indented)
                System.IO.File.WriteAllText(tempPath, json)
                Return tempPath
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[SkillDiagnostic] 刷新文件失败: " & ex.Message)
                Return ""
            End Try
        End Function

#End Region

#Region "Dispose + 兼容旧方法"

        Public Sub Dispose() Implements IDisposable.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If _isDisposed Then Return

            If disposing Then
                If _workflowEngine IsNot Nothing Then
                    _workflowEngine.Dispose()
                    _workflowEngine = Nothing
                End If
                If _llmEngine IsNot Nothing Then
                    _llmEngine.Dispose()
                    _llmEngine = Nothing
                End If
                ClearDiagnosticLogs()
            End If

            _isDisposed = True
        End Sub

        ''' <summary>
        ''' 旧API兼容：设置全局默认超时（秒）
        ''' </summary>
        <Obsolete("请使用 SetDefaultTimeoutSeconds 代替")>
        Public Sub SetDefaultTimeout(seconds As Integer)
            SetDefaultTimeoutSeconds(seconds)
        End Sub

#End Region

    End Class

    Public Class SkillDiagnosticEntry
        Public Property Timestamp As DateTime
        Public Property SkillName As String
        Public Property Phase As String
        Public Property RetryCount As Integer
        Public Property Level As String
        Public Property Message As String
        Public Property ThreadId As Integer
    End Class

End Namespace