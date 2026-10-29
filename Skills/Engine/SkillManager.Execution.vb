Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports System.Threading
Imports System.Threading.Tasks
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils
Imports YamlDotNet.RepresentationModel

Namespace Skills

    ' SkillManager.Execution.vb — Skill 执行调度（ExecuteSkill 入口区）
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        ''' <summary>
        ''' 执行 Skill（支持简单 Skill、复合 Skill 和多模态处理）- 使用队列确保严格顺序执行
        ''' </summary>
        Public Async Function ExecuteSkill(skillName As String, inputs As Dictionary(Of String, Object), Optional progress As IProgress(Of String) = Nothing) As Task(Of WorkflowResult)
            System.Diagnostics.Debug.WriteLine($"[ExecuteSkill] 被调用，Skill名称: {skillName}")

            ' 创建任务完成源，用于等待队列处理结果
            Dim tcs As New TaskCompletionSource(Of WorkflowResult)

            ' 将请求加入队列
            SyncLock _queueLock
                _skillQueue.Enqueue(Tuple.Create(skillName, inputs, progress, tcs))
                System.Diagnostics.Debug.WriteLine($"[ExecuteSkill] 请求已加入队列，当前队列长度: {_skillQueue.Count}")

                ' 如果队列没有在处理，启动处理
                If Not _isProcessingQueue Then
                    _isProcessingQueue = True
                    ' 启动后台任务处理队列（fire-and-forget）
#Disable Warning BC42358  ' Fire-and-forget调用，故意不等待
                    Task.Run(AddressOf ProcessSkillQueueAsync)
#Enable Warning BC42358
                End If
            End SyncLock

            ' 等待队列处理完成
            Return Await tcs.Task
        End Function

        ''' <summary>
        ''' 处理Skill队列 - 确保严格的顺序执行
        ''' </summary>
        Private Async Function ProcessSkillQueueAsync() As Task
            System.Diagnostics.Debug.WriteLine("[ProcessSkillQueueAsync] 队列处理启动")

            Try
                While True
                    Dim currentItem As Tuple(Of String, Dictionary(Of String, Object), IProgress(Of String), TaskCompletionSource(Of WorkflowResult)) = Nothing

                    ' 从队列取出一个请求
                    SyncLock _queueLock
                        If _skillQueue.Count = 0 Then
                            ' 队列为空，停止处理
                            _isProcessingQueue = False
                            System.Diagnostics.Debug.WriteLine("[ProcessSkillQueueAsync] 队列为空，处理停止")
                            Return
                        End If
                        currentItem = _skillQueue.Dequeue()
                        System.Diagnostics.Debug.WriteLine($"[ProcessSkillQueueAsync] 取出请求，剩余队列长度: {_skillQueue.Count}")
                    End SyncLock

                    ' 执行Skill
                    Try
                        Dim result = Await ExecuteSkillInternal(currentItem.Item1, currentItem.Item2, currentItem.Item3)
                        currentItem.Item4.SetResult(result)
                    Catch ex As Exception
                        currentItem.Item4.SetException(ex)
                    End Try

                    ' 智能等待：检查PaddleOCR进程是否已退出（最多等待5秒）
                    Await WaitForPaddleOcrProcessExitAsync()
                End While
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ProcessSkillQueueAsync] 队列处理异常: " & ex.Message)
                Throw
            Finally
                ' 【修复】确保无论何种原因退出，都重置处理状态
                SyncLock _queueLock
                    _isProcessingQueue = False
                    System.Diagnostics.Debug.WriteLine("[ProcessSkillQueueAsync] 重置 _isProcessingQueue")
                End SyncLock
            End Try
        End Function

        ''' <summary>
        ''' 智能等待PaddleOCR进程退出（最多等待5秒，每200ms检查一次）
        ''' </summary>
        Private Async Function WaitForPaddleOcrProcessExitAsync() As Task
            Const MAX_WAIT_MS As Integer = 5000
            Const CHECK_INTERVAL_MS As Integer = 200
            Dim totalWaited As Integer = 0

            System.Diagnostics.Debug.WriteLine("[WaitForPaddleOcr] 开始智能等待PaddleOCR进程退出...")

            While totalWaited < MAX_WAIT_MS
                Dim pythonProcesses = System.Diagnostics.Process.GetProcessesByName("python")
                Dim hasPaddleOcrProcess = False

                For Each proc As System.Diagnostics.Process In pythonProcesses
                    Try
                        ' 检查进程是否与PaddleOCR相关（通过进程名和窗口标题）
                        If proc.MainWindowTitle.Contains("PaddleOCR") OrElse
                           proc.ProcessName.Contains("paddleocr") Then
                            hasPaddleOcrProcess = True
                            Exit For
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine("检查 PaddleOCR 进程时出错: " & ex.Message)
                    End Try
                Next

                If Not hasPaddleOcrProcess Then
                    System.Diagnostics.Debug.WriteLine($"[WaitForPaddleOcr] ✓ PaddleOCR进程已退出（等待了{totalWaited}ms）")
                    Return
                End If

                Await Task.Delay(CHECK_INTERVAL_MS)
                totalWaited += CHECK_INTERVAL_MS
            End While

            System.Diagnostics.Debug.WriteLine($"[WaitForPaddleOcr] ⚠ 等待超时({MAX_WAIT_MS}ms)，继续处理下一个任务")
        End Function

        ''' <summary>
        ''' 实际执行Skill的内部方法
        ''' </summary>
        Private Async Function ExecuteSkillInternal(skillName As String, inputs As Dictionary(Of String, Object), progress As IProgress(Of String)) As Task(Of WorkflowResult)
            System.Diagnostics.Debug.WriteLine($"[ExecuteSkillInternal] 开始执行Skill: {skillName}")

            Dim skill = GetSkill(skillName)
            If skill Is Nothing Then
                System.Diagnostics.Debug.WriteLine($"[ExecuteSkillInternal] 错误: 找不到 Skill: {skillName}")
                Return New WorkflowResult With {
                    .Success = False,
                    .ErrorMessage = $"找不到 Skill: {skillName}"
                }
            End If

            System.Diagnostics.Debug.WriteLine($"[ExecuteSkillInternal] 找到 Skill: {skill.DisplayName}, 处理模式: {skill.ProcessingMode}")

            ' 如果是自定义Skill，用CustomSkillExecutor执行
            If skill.IsCustom Then
                System.Diagnostics.Debug.WriteLine("[ExecuteSkillInternal] 执行自定义Skill...")
                Dim executor As New CustomSkillExecutor()
                Return Await executor.ExecuteCustomSkill(skill, inputs)
            End If

            ' 根据处理模式选择执行方式
            Select Case skill.ProcessingMode
                Case SkillProcessingMode.MultimodalLLM
                    Dim mmResult = Await ExecuteMultimodalSkill(skill, inputs, progress)
                    ' Skill Creator 后处理：将 LLM 输出的 JSON 解析并保存为自定义 Skill
                    If skill.Name = "SkillCreator" AndAlso mmResult.Success Then
                        mmResult = Await PostProcessGeneratedSkill(mmResult)
                    End If
                    ' 【自动导入 OCR 解析脚本】声明 autoImportParseScript 的 Skill（非自定义副本走这条多模态路径）
                    If skill.AutoImportParseScript AndAlso mmResult.Success Then
                        Services.ParseScriptManager.AutoImportSkillParseScript(skill, mmResult)
                    End If
                    Return mmResult
                Case SkillProcessingMode.LocalOCR
                    Return Await ExecuteLocalOCRSkill(skill, inputs, progress)
                Case SkillProcessingMode.Hybrid
                    ' Hybrid Skill 统一委托给 CustomSkillExecutor 执行完整流程
                    ' （LLM 分析 → 结果注入 templateCode → VBA/VSTO 执行）。
                    ' 之前对非图片输入会降级为 ExecuteSimpleSkill，导致只返回"已识别，等待 AI 生成响应"占位消息，
                    ' 实际并未执行任何分析或 Excel 操作。
                    System.Diagnostics.Debug.WriteLine("[ExecuteSkillInternal] 执行 Hybrid Skill，委托 CustomSkillExecutor...")
                    Dim hybridExecutor As New CustomSkillExecutor()
                    Return Await hybridExecutor.ExecuteCustomSkill(skill, inputs)
                Case SkillProcessingMode.PythonCode
                    ' 【P0修复】PythonCode Skill 委托给 CustomSkillExecutor（已实现 ExecutePythonSkill）
                    ' 之前落到 Case Else 会被当 VBA 处理，导致 Python 代码无法执行
                    System.Diagnostics.Debug.WriteLine("[ExecuteSkillInternal] 执行 Python Skill，委托 CustomSkillExecutor...")
                    Dim pyExecutor As New CustomSkillExecutor()
                    Return Await pyExecutor.ExecuteCustomSkill(skill, inputs)
                Case Else
                    If skill.IsComposite Then
                        System.Diagnostics.Debug.WriteLine("[ExecuteSkillInternal] 执行复合 Skill 工作流...")
                        Dim result = Await _workflowEngine.ExecuteWorkflow(skill, inputs, progress)
                        System.Diagnostics.Debug.WriteLine($"[ExecuteSkillInternal] 复合 Skill 执行完成，Success: {result.Success}")
                        Return result
                    Else
                        System.Diagnostics.Debug.WriteLine("[ExecuteSkillInternal] 执行简单 Skill...")
                        Return ExecuteSimpleSkill(skill, inputs)
                    End If
            End Select
        End Function

        ''' <summary>
        ''' 后处理 SkillCreator 的 LLM 输出：解析 JSON 为 SkillDefinition 并保存
        ''' </summary>
        Private Async Function PostProcessGeneratedSkill(result As WorkflowResult) As Task(Of WorkflowResult)
            Return Await Task.Run(Function()
                Try
                    Dim rawContent As String = ""
                    If result.Outputs.ContainsKey("content") AndAlso result.Outputs("content") IsNot Nothing Then
                        rawContent = result.Outputs("content").ToString()
                    End If

                    If String.IsNullOrWhiteSpace(rawContent) Then
                        result.Success = False
                        result.ErrorMessage = "LLM 返回内容为空，无法生成 Skill"
                        Return result
                    End If

                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 原始输出长度: " & rawContent.Length)

                    ' 调试：将 LLM 原始输出写入临时日志
                    Try
                        Dim debugDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ExcelAIVSTOAddIn_Debug")
                        If Not System.IO.Directory.Exists(debugDir) Then System.IO.Directory.CreateDirectory(debugDir)
                        Dim debugFile = System.IO.Path.Combine(debugDir, $"SkillManager_Raw_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid().ToString("N").Substring(0, 6)}.txt")
                        System.IO.File.WriteAllText(debugFile, $"长度: {rawContent.Length}{vbCrLf}{vbCrLf}原始内容:{vbCrLf}{rawContent}{vbCrLf}")
                        System.Diagnostics.Debug.WriteLine($"[SkillCreator] 原始输出已保存: {debugFile}")
                    Catch logEx As Exception
                        System.Diagnostics.Debug.WriteLine($"[SkillCreator] 保存调试日志失败: {logEx.Message}")
                    End Try

                    ' 提取 JSON（可能包裹在 Markdown 代码块中）
                    Dim jsonText = ExtractJsonFromLLMResponse(rawContent)
                    If String.IsNullOrWhiteSpace(jsonText) Then
                        result.Success = False
                        result.ErrorMessage = "无法从 LLM 响应中提取有效的 SkillDefinition JSON"
                        Return result
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
                        Return result
                    End If

                    ' 验证必填字段
                    If String.IsNullOrWhiteSpace(newSkill.Name) Then
                        result.Success = False
                        result.ErrorMessage = "生成的 Skill 缺少 name 字段"
                        Return result
                    End If

                    If String.IsNullOrWhiteSpace(newSkill.DisplayName) Then
                        newSkill.DisplayName = newSkill.Name
                    End If

                    ' 设置元数据
                    newSkill.IsCustom = True
                    newSkill.CreatedDate = DateTime.Now
                    newSkill.LastModifiedDate = DateTime.Now
                    newSkill.Version = 1
                    newSkill.UsageCount = 0
                    newSkill.CreatedBy = "SkillCreator"

                    System.Diagnostics.Debug.WriteLine($"[SkillCreator] 已生成自定义 Skill: {newSkill.Name}")

                    ' 使用精简序列化，仅包含必要字段
                    Dim skillJsonStr = SkillJsonHelper.SerializeSkillCompact(newSkill)

                    ' 保存到 Outputs（顶层字段，扁平结构）
                    result.Outputs("generatedSkillName") = newSkill.Name
                    result.Outputs("generatedSkillDisplayName") = newSkill.DisplayName
                    result.Outputs("generatedSkillJson") = skillJsonStr
                    result.Outputs("saved") = False

                    ' content 仅包含显示信息（不嵌套 skillJson，避免转义混乱）
                    Dim contentObj As New Dictionary(Of String, Object) From {
                        {"type", "skill_card"},
                        {"skillName", newSkill.Name},
                        {"skillDisplayName", newSkill.DisplayName},
                        {"skillDescription", If(newSkill.Description, "")},
                        {"skillCategory", If(newSkill.Category, "")},
                        {"processingMode", newSkill.ProcessingMode.ToDisplayName()},
                        {"saved", False}
                    }
                    result.Outputs("content") = Newtonsoft.Json.JsonConvert.SerializeObject(contentObj)

                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[SkillCreator] 后处理失败: " & ex.Message)
                    result.Success = False
                    result.ErrorMessage = "生成 Skill 后处理失败: " & ex.Message
                End Try

                Return result
            End Function)
        End Function

        ' 【历史备注】尝试解析 Function Calling 格式的 JSON，提取 SkillDefinition
        ' 格式: {"function_call": {"name": "create_vba_skill", "arguments": {...}}}
        ' （该方法定义在其他分部文件中）

    End Class

End Namespace
