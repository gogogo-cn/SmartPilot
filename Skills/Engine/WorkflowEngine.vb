Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.IO
Imports System.Text.RegularExpressions
Imports ExcelAIVSTOAddIn.Services
Imports System.Reflection

Namespace Skills

    ''' <summary>
    ''' 步骤执行结果
    ''' </summary>
    Public Class StepResult
        Public Property Success As Boolean
        Public Property Output As Object
        Public Property ErrorMessage As String
        Public Property ExecutionTime As TimeSpan

        Public Shared Function CreateSuccess(output As Object) As StepResult
            Return New StepResult With {
                .Success = True,
                .Output = output
            }
        End Function

        Public Shared Function CreateFailure(errorMessage As String) As StepResult
            Return New StepResult With {
                .Success = False,
                .ErrorMessage = errorMessage
            }
        End Function
    End Class

    ' WorkflowResult 已拆分到独立的 Skills\Engine\WorkflowResult.vb
    ' （原因：本文件强依赖 Excel Interop，而 WorkflowResult 是纯 POCO，
    '   拆出后可让 ParseScriptManager 脱离 Excel 耦合进入公共核心库 Core）

    ''' <summary>
    ''' Excel 列配置
    ''' </summary>
    Public Class ExcelColumnConfig
        Public Property Field As String
        Public Property Header As String
        Public Property Width As Integer
        Public Property Format As String
    End Class

    ''' <summary>
    ''' 数据类型配置 - 支持可扩展的业务类型定义
    ''' </summary>
    Public Class DataTypeConfiguration
        Public Property TypeName As String
        Public Property DisplayName As String
        Public Property DefaultSheetName As String
        Public Property FieldMappings As Dictionary(Of String, String)
        Public Property IncludeImagePath As Boolean
        Public Property DetectionFields As String()
        Public Property ExcelColumns As List(Of ExcelColumnConfig)
        Public Property DynamicColumns As Boolean

        Public Sub New(typeName As String, displayName As String, defaultSheetName As String,
                       fieldMappings As Dictionary(Of String, String),
                       Optional includeImagePath As Boolean = False,
                       Optional detectionFields As String() = Nothing,
                       Optional excelColumns As List(Of ExcelColumnConfig) = Nothing,
                       Optional dynamicColumns As Boolean = False)
            Me.TypeName = typeName
            Me.DisplayName = displayName
            Me.DefaultSheetName = defaultSheetName
            Me.FieldMappings = fieldMappings
            Me.IncludeImagePath = includeImagePath
            Me.DetectionFields = If(detectionFields, New String() {})
            Me.ExcelColumns = If(excelColumns, New List(Of ExcelColumnConfig)())
            Me.DynamicColumns = dynamicColumns
        End Sub
    End Class

    ''' <summary>
    ''' 工作流引擎 - 执行复合 Skill 的工作流步骤
    ''' </summary>
    Public Class WorkflowEngine
        Implements IDisposable

        Private _skillManager As SkillManager
        Private _variableStore As Dictionary(Of String, Object)
        Private _executionLog As List(Of String)
        Private _httpClient As HttpClient
        Private _dataTypeConfigs As Dictionary(Of String, DataTypeConfiguration)
        Private _currentWorkflowId As String
        Private _checkpointIntervalSeconds As Integer
        Private _lastCheckpointTime As DateTime
        Private _isDisposed As Boolean
        Private _progress As IProgress(Of String) ' 进度报告器
        Private _llmEngine As MultimodalLLMEngine
        Private _currentSkill As SkillDefinition

        Public Event StepStarted(stepId As String, stepName As String)
        Public Event StepCompleted(stepId As String, result As StepResult)
        Public Event StepFailed(stepId As String, errorMessage As String)
        Public Event WorkflowCompleted(result As WorkflowResult)

        Public Sub New(skillManager As SkillManager)
            Try
                _skillManager = skillManager
                _llmEngine = New MultimodalLLMEngine()
                _executionLog = New List(Of String)()
                _httpClient = New HttpClient()
                ' 【修复】使用 Integer.TryParse 防御性解析配置值，避免非数字字符串导致 FormatException
                Dim httpTimeout As Integer = 30
                Dim checkpointInterval As Integer = 300
                If Not Integer.TryParse(
                    Utils.ConfigurationManager.Instance.GetSetting(
                        Utils.ConfigConstants.Workflow.HTTP_TIMEOUT,
                        Utils.ConfigConstants.Workflow.DEFAULT_HTTP_TIMEOUT.ToString()),
                    httpTimeout) Then
                    httpTimeout = 30
                End If
                _httpClient.Timeout = TimeSpan.FromSeconds(httpTimeout)
                _isDisposed = False
                _lastCheckpointTime = DateTime.Now
                If Not Integer.TryParse(
                    Utils.ConfigurationManager.Instance.GetSetting(
                        Utils.ConfigConstants.Workflow.CHECKPOINT_INTERVAL,
                        Utils.ConfigConstants.Workflow.DEFAULT_CHECKPOINT_INTERVAL.ToString()),
                    checkpointInterval) Then
                    checkpointInterval = 300
                End If
                _checkpointIntervalSeconds = checkpointInterval

                Utils.AppLogger.LogInfo("WorkflowEngine", "初始化完成")
            Catch ex As Exception
                Utils.AppLogger.LogError("WorkflowEngine", "初始化失败", ex)
                Throw
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            DisposeCore()
        End Sub

        Protected Overridable Sub DisposeCore()
            If _isDisposed Then Return

            _isDisposed = True

            If _httpClient IsNot Nothing Then
                _httpClient.Dispose()
            End If
        End Sub

        ''' <summary>
        ''' 初始化数据类型配置（优先从配置文件加载，失败则使用默认配置）
        ''' </summary>
        Private Sub InitializeDataTypeConfigs()
            _dataTypeConfigs = New Dictionary(Of String, DataTypeConfiguration)()

            ' 首先尝试从配置文件加载
            Dim configLoaded = LoadDataTypeConfigsFromFile()

            ' 如果配置文件加载失败，使用默认配置
            If Not configLoaded Then
                LoadDefaultDataTypeConfigs()
                LogMessage("使用默认数据类型配置")
            End If
        End Sub

        ''' <summary>
        ''' 从配置文件加载数据类型配置
        ''' </summary>
        Private Function LoadDataTypeConfigsFromFile() As Boolean
            Try
                Dim configPath = GetDataTypeConfigFilePath()

                If Not File.Exists(configPath) Then
                    LogMessage(String.Format("数据类型配置文件不存在: {0}", configPath))
                    Return False
                End If

                Dim json = File.ReadAllText(configPath)
                Dim configRoot = JsonConvert.DeserializeObject(Of DataTypeConfigRoot)(json)

                If configRoot Is Nothing OrElse configRoot.DataTypes Is Nothing Then
                    LogMessage("数据类型配置文件格式无效")
                    Return False
                End If

                For Each dataType In configRoot.DataTypes
                    Dim fieldMappings = New Dictionary(Of String, String)()
                    If dataType.FieldMappings IsNot Nothing Then
                        For Each mapping In dataType.FieldMappings
                            fieldMappings(mapping.Key) = mapping.Value
                        Next
                    End If

                    Dim excelColumns = New List(Of ExcelColumnConfig)()
                    If dataType.ExcelColumns IsNot Nothing Then
                        For Each col In dataType.ExcelColumns
                            excelColumns.Add(New ExcelColumnConfig With {
                                .Field = col.Field,
                                .Header = col.Header,
                                .Width = col.Width,
                                .Format = col.Format
                            })
                        Next
                    End If

                    Dim config = New DataTypeConfiguration(
                        dataType.TypeName,
                        dataType.DisplayName,
                        dataType.DefaultSheetName,
                        fieldMappings,
                        dataType.IncludeImagePath,
                        dataType.DetectionFields,
                        excelColumns,
                        dataType.DynamicColumns
                    )

                    _dataTypeConfigs(dataType.TypeName) = config
                    LogMessage(String.Format("从配置文件加载数据类型: {0}", dataType.TypeName))
                Next

                LogMessage(String.Format("成功从配置文件加载 {0} 个数据类型", _dataTypeConfigs.Count))
                Return True

            Catch ex As Exception
                LogMessage(String.Format("加载数据类型配置文件失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 获取数据类型配置文件路径
        ''' </summary>
        Private Function GetDataTypeConfigFilePath() As String
            ' 首先检查是否有自定义配置路径
            Dim customPath = Utils.ConfigurationManager.Instance.GetSetting(
                Utils.ConfigConstants.DataType.CONFIG_FILE, "")

            If Not String.IsNullOrEmpty(customPath) AndAlso File.Exists(customPath) Then
                Return customPath
            End If

            ' 使用统一配置目录
            Dim defaultPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "DataTypeConfigs.json")
            Return defaultPath
        End Function

        ''' <summary>
        ''' 加载默认数据类型配置（作为后备）
        ''' </summary>
        Private Sub LoadDefaultDataTypeConfigs()
            _dataTypeConfigs = New Dictionary(Of String, DataTypeConfiguration) From {
                {"发票", New DataTypeConfiguration(
                    "发票", "发票", "发票记录",
                    New Dictionary(Of String, String) From {
                        {"invoiceNumber", "invoiceNumber"}, {"invoiceCode", "invoiceCode"},
                        {"invoiceDate", "invoiceDate"}, {"amount", "amount"}, {"totalAmount", "totalAmount"},
                        {"taxRate", "taxRate"}, {"taxAmount", "taxAmount"}, {"netAmount", "netAmount"},
                        {"sellerName", "sellerName"}, {"sellerTaxId", "sellerTaxId"},
                        {"sellerAddress", "sellerAddress"}, {"sellerBank", "sellerBank"},
                        {"buyerName", "buyerName"}, {"buyerTaxId", "buyerTaxId"},
                        {"buyerAddress", "buyerAddress"}, {"buyerBank", "buyerBank"},
                        {"invoiceType", "invoiceType"}, {"specModel", "specModel"},
                        {"payee", "payee"}, {"reviewer", "reviewer"},
                        {"drawer", "drawer"}, {"verifyCode", "verifyCode"}, {"remarks", "remarks"}
                    },
                    True,
                    New String() {"invoiceNumber", "invoiceCode", "invoiceDate", "amount"}
                )},
                {"名片/身份证", New DataTypeConfiguration(
                    "名片/身份证", "名片/身份证", "",
                    New Dictionary(Of String, String) From {
                        {"姓名", "姓名"}, {"性别", "性别"}, {"民族", "民族"},
                        {"出生日期", "出生日期"}, {"出生", "出生"}, {"住址", "住址"},
                        {"公民身份号码", "公民身份号码"}, {"身份证号码", "身份证号码"},
                        {"身份证号", "身份证号"}, {"签发机关", "签发机关"}, {"有效期限", "有效期限"},
                        {"职位", "职位"}, {"公司名称", "公司名称"}, {"公司", "公司"},
                        {"电话", "电话"}, {"手机", "手机"}, {"邮箱", "邮箱"},
                        {"地址", "地址"}, {"部门", "部门"}
                    },
                    False,
                    New String() {"姓名", "公民身份号码", "身份证号码", "电话", "公司名称"}
                )}
            }
        End Sub

        ''' <summary>
        ''' 重新加载数据类型配置（支持热更新）
        ''' </summary>
        Public Function ReloadDataTypeConfigs() As Boolean
            Try
                LogMessage("重新加载数据类型配置...")
                _dataTypeConfigs.Clear()
                InitializeDataTypeConfigs()
                Return True
            Catch ex As Exception
                LogMessage(String.Format("重新加载数据类型配置失败: {0}", ex.Message))
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 数据类型配置根对象（用于JSON反序列化）
        ''' </summary>
        Private Class DataTypeConfigRoot
            Public Property Version As Integer
            Public Property Description As String
            Public Property LastUpdated As String
            Public Property DataTypes As List(Of DataTypeConfigItem)
        End Class

        ''' <summary>
        ''' 数据类型配置项（用于JSON反序列化）
        ''' </summary>
        Private Class DataTypeConfigItem
            Public Property TypeName As String
            Public Property DisplayName As String
            Public Property DefaultSheetName As String
            Public Property IncludeImagePath As Boolean
            Public Property DetectionFields As String()
            Public Property FieldMappings As Dictionary(Of String, String)
            Public Property ExcelColumns As List(Of ExcelColumnConfig)
            Public Property DynamicColumns As Boolean
        End Class

        ''' <summary>
        ''' 注册新的数据类型配置（支持运行时动态扩展）
        ''' </summary>
        Public Sub RegisterDataType(config As DataTypeConfiguration)
            If Not _dataTypeConfigs.ContainsKey(config.TypeName) Then
                _dataTypeConfigs.Add(config.TypeName, config)
                LogMessage(String.Format("已注册新数据类型: {0}", config.TypeName))
            Else
                _dataTypeConfigs(config.TypeName) = config
                LogMessage(String.Format("已更新数据类型配置: {0}", config.TypeName))
            End If
        End Sub

        ''' <summary>
        ''' 检测数据类型（基于特征字段匹配）
        ''' </summary>
        Private Function DetectDataType(structuredData As Object) As String
            If structuredData Is Nothing Then Return ""

            For Each config In _dataTypeConfigs.Values
                If CheckDetectionFields(structuredData, config.DetectionFields) Then
                    Return config.TypeName
                End If
            Next

            Return ""
        End Function

        ''' <summary>
        ''' 检查是否包含检测字段
        ''' </summary>
        Private Function CheckDetectionFields(data As Object, fields As String()) As Boolean
            If fields Is Nothing OrElse fields.Length = 0 Then Return False

            Dim matchCount As Integer = 0
            For Each field In fields
                Try
                    Dim value = CallByName(data, field, CallType.Get)
                    If value IsNot Nothing AndAlso Not String.IsNullOrEmpty(value.ToString()) Then
                        matchCount += 1
                    End If
                Catch ex As Exception
                    ' 忽略 MissingMemberException 和其他访问异常
                    System.Diagnostics.Debug.WriteLine(String.Format("CheckDetectionFields: 访问字段 {0} 失败: {1}", field, ex.Message))
                End Try
            Next

            Try
                Dim minMatchFields = Integer.Parse(
                    Utils.ConfigurationManager.Instance.GetSetting(
                        Utils.ConfigConstants.DataType.MIN_MATCH_FIELDS,
                        Utils.ConfigConstants.DataType.DEFAULT_MIN_MATCH_FIELDS.ToString()))
                Return matchCount >= minMatchFields
            Catch ex As Exception
                ' 配置解析失败时使用默认值
                Return matchCount >= 1
            End Try
        End Function

        ''' <summary>
        ''' 执行工作流
        ''' </summary>
        Public Async Function ExecuteWorkflow(skill As SkillDefinition, inputs As Dictionary(Of String, Object), Optional progress As IProgress(Of String) = Nothing) As Task(Of WorkflowResult)
            Dim result As New WorkflowResult
            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()
            _executionLog = New List(Of String)
            _variableStore = New Dictionary(Of String, Object)
            _progress = progress ' 保存进度报告器
            _currentSkill = skill

            _currentWorkflowId = Guid.NewGuid().ToString()
            Dim stepExecutionCount As New Dictionary(Of String, Integer)
            Dim persistedState As WorkflowPersistenceState = Nothing
            Dim currentStepId As String = Nothing

            Try
                LogMessage(String.Format("开始执行复合 Skill: {0}", skill.DisplayName))

                persistedState = _skillManager.GetPersistedWorkflow(_currentWorkflowId)

                If persistedState IsNot Nothing Then
                    LogMessage(String.Format("恢复工作流 from checkpoint at step: {0}", persistedState.CurrentStepId))
                    _variableStore = persistedState.VariableStore
                    _executionLog = persistedState.ExecutionLog
                    stepExecutionCount = persistedState.StepExecutionCount
                    currentStepId = persistedState.CurrentStepId
                Else
                    ValidateInputs(skill, inputs)
                    InitializeVariables(skill, inputs)
                    If skill.WorkflowSteps IsNot Nothing AndAlso skill.WorkflowSteps.Count > 0 Then
                        currentStepId = skill.WorkflowSteps(0).StepId
                    End If
                End If

                _lastCheckpointTime = DateTime.Now

                If skill.WorkflowSteps IsNot Nothing AndAlso skill.WorkflowSteps.Count > 0 Then
                    Dim maxStepRetries = Integer.Parse(
                        Utils.ConfigurationManager.Instance.GetSetting(
                            Utils.ConfigConstants.Workflow.MAX_STEP_RETRIES,
                            Utils.ConfigConstants.Workflow.DEFAULT_MAX_STEP_RETRIES.ToString()))

                    While Not String.IsNullOrEmpty(currentStepId)
                        If Not stepExecutionCount.ContainsKey(currentStepId) Then
                            stepExecutionCount(currentStepId) = 0
                        End If

                        stepExecutionCount(currentStepId) += 1

                        If stepExecutionCount(currentStepId) > maxStepRetries Then
                            Throw New InvalidOperationException(String.Format(
                                "步骤 {0} 执行次数超过最大限制 ({1} 次)，可能存在死循环",
                                currentStepId, maxStepRetries))
                        End If

                        Dim workflowStep = skill.WorkflowSteps.FirstOrDefault(Function(s) s.StepId = currentStepId)
                        If workflowStep Is Nothing Then
                            Throw New InvalidOperationException(String.Format("找不到步骤: {0}", currentStepId))
                        End If

                        Dim stepResult = Await ExecuteStep(workflowStep)

                        If (DateTime.Now - _lastCheckpointTime).TotalSeconds >= _checkpointIntervalSeconds Then
                            _skillManager.SaveWorkflowState(_currentWorkflowId, skill.Name, currentStepId,
                                _variableStore, _executionLog, stepExecutionCount)
                            _lastCheckpointTime = DateTime.Now
                            LogMessage("检查点已保存")
                        End If

                        If stepResult.Success Then
                            currentStepId = workflowStep.NextStepOnSuccess
                        Else
                            LogMessage(String.Format("步骤 {0} 执行失败 (第{1}次): {2}",
                                workflowStep.StepName,
                                stepExecutionCount(currentStepId),
                                stepResult.ErrorMessage))
                            If Not String.IsNullOrEmpty(workflowStep.NextStepOnFailure) Then
                                currentStepId = workflowStep.NextStepOnFailure
                            Else
                                Throw New InvalidOperationException(String.Format("步骤 {0} 执行失败: {1}",
                                    workflowStep.StepName, stepResult.ErrorMessage))
                            End If
                        End If
                    End While
                End If

                CollectOutputs(skill, result)

                result.Success = True
                stopwatch.Stop()
                result.TotalExecutionTime = stopwatch.Elapsed

                _skillManager.UpdateWorkflowStatus(_currentWorkflowId, WorkflowPersistenceState.WorkflowStatus.Completed)
                _skillManager.RemovePersistedWorkflow(_currentWorkflowId)

                LogMessage(String.Format("工作流执行完成，总耗时: {0:F2} 秒", stopwatch.Elapsed.TotalSeconds))

                RaiseEvent WorkflowCompleted(result)

            Catch ex As Exception
                stopwatch.Stop()
                result.Success = False
                result.ErrorMessage = FormatExceptionMessage(ex, "工作流执行")
                result.TotalExecutionTime = stopwatch.Elapsed
                LogMessage(String.Format("工作流执行失败: {0}{1}堆栈: {2}", ex.Message, vbCrLf, ex.StackTrace))

                _skillManager.UpdateWorkflowStatus(_currentWorkflowId, WorkflowPersistenceState.WorkflowStatus.Failed)
            End Try

            result.ExecutionLog = _executionLog
            Return result
        End Function

        ''' <summary>
        ''' 执行单个步骤
        ''' </summary>
        Private Async Function ExecuteStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()

            Try
                RaiseEvent StepStarted(workflowStep.StepId, workflowStep.StepName)
                LogMessage(String.Format("执行步骤: {0} (类型: {1})", workflowStep.StepName, workflowStep.StepType))

                Dim result As StepResult

                Select Case workflowStep.StepType
                    Case WorkflowStepType.ExecuteCode
                        result = Await ExecuteCodeStep(workflowStep)
                    Case WorkflowStepType.CallApi
                        result = Await ExecuteApiStep(workflowStep)
                    Case WorkflowStepType.InvokeSubSkill
                        result = Await ExecuteSubSkillStep(workflowStep)
                    Case WorkflowStepType.Condition
                        result = ExecuteConditionStep(workflowStep)
                    Case WorkflowStepType.[Loop]
                        result = ExecuteLoopStep(workflowStep)
                    Case WorkflowStepType.Wait
                        result = Await ExecuteWaitStep(workflowStep)
                    Case WorkflowStepType.InvokeMcpTool
                        ' 【优先级2扩展】调用 MCP 外部工具
                        result = Await ExecuteMcpToolStep(workflowStep)
                    ' === v2 统一体系新增：原 CustomStepType 中的领域步骤 ===
                    Case WorkflowStepType.UploadImage
                        result = ExecuteUploadImageStep(workflowStep)
                    Case WorkflowStepType.OCRRecognition
                        result = Await ExecuteOcrRecognitionStepAsync(workflowStep)
                    Case WorkflowStepType.ExtractFields
                        result = ExecuteExtractFieldsStep(workflowStep)
                    Case WorkflowStepType.WriteToExcel
                        result = ExecuteWriteToExcelStep(workflowStep)
                    Case WorkflowStepType.CustomLLM
                        result = Await ExecuteCustomLlmStepAsync(workflowStep)
                    Case WorkflowStepType.LLMProcess, WorkflowStepType.MultimodalProcess
                        result = Await ExecuteLlmProcessStepAsync(workflowStep)
                    Case Else
                        result = StepResult.CreateFailure(String.Format("未知的步骤类型: {0}", workflowStep.StepType))
                End Select

                stopwatch.Stop()
                result.ExecutionTime = stopwatch.Elapsed

                If result.Success Then
                    LogMessage(String.Format("步骤 {0} 执行成功，耗时: {1:F2} 秒", workflowStep.StepName, stopwatch.Elapsed.TotalSeconds))
                    RaiseEvent StepCompleted(workflowStep.StepId, result)
                Else
                    LogMessage(String.Format("步骤 {0} 执行失败: {1}", workflowStep.StepName, result.ErrorMessage))
                    RaiseEvent StepFailed(workflowStep.StepId, result.ErrorMessage)
                End If

                Return result

            Catch ex As Exception
                stopwatch.Stop()
                Dim fullError = FormatExceptionMessage(ex, String.Format("步骤 {0}", workflowStep.StepName))
                LogMessage(fullError)
                RaiseEvent StepFailed(workflowStep.StepId, fullError)
                Return StepResult.CreateFailure(fullError)
            End Try
        End Function

        ''' <summary>
        ''' 执行代码步骤（支持 VBA）- 使用配置驱动的数据类型检测
        ''' </summary>
        Private Async Function ExecuteCodeStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Return Await Task.Run(Function()
                Try
                    ' 替换变量占位符
                    Dim code = ReplaceVariables(workflowStep.Code)

                    ' 使用配置驱动的方式检测数据类型
                    Dim detectedType As String = ""
                    If _variableStore.ContainsKey("output.structuredData") Then
                        Dim structuredData As Object = _variableStore("output.structuredData")
                        If structuredData IsNot Nothing Then
                            detectedType = DetectDataType(structuredData)
                            Utils.AppLogger.LogDebug("WorkflowEngine", $"检测到数据类型: {detectedType}")
                        End If
                    End If

                    ' 也检查简单的发票标识
                    Dim hasSimpleInvoiceData As Boolean = _variableStore.ContainsKey("invoiceNumber")

                    If workflowStep.StepName.Contains("填入Excel") Then
                        Dim structuredDataForFill As Dictionary(Of String, Object) = Nothing
                        If _variableStore.ContainsKey("output.structuredData") Then
                            Try
                                structuredDataForFill = CType(_variableStore("output.structuredData"), Dictionary(Of String, Object))
                            Catch ex As Exception
                                Utils.AppLogger.LogError("WorkflowEngine", "转换structuredData失败", ex)
                            End Try
                        End If

                        Dim hasValidData As Boolean = structuredDataForFill IsNot Nothing AndAlso structuredDataForFill.Count > 0

                        If Not hasValidData AndAlso hasSimpleInvoiceData Then
                            hasValidData = True
                        End If

                        If hasValidData Then
                            If Not String.IsNullOrEmpty(detectedType) AndAlso _dataTypeConfigs.ContainsKey(detectedType) Then
                                Dim config = _dataTypeConfigs(detectedType)
                                LogMessage(String.Format("使用配置 [{0}] 填入Excel", config.DisplayName))
                                Return FillStructuredDataToExcel(workflowStep, config)
                            ElseIf hasSimpleInvoiceData Then
                                LogMessage("使用简单发票检测填入Excel")
                                Return FillInvoiceToExcel(workflowStep)
                            Else
                                LogMessage("使用通用处理填入Excel")
                                Return FillGenericDataToExcel(workflowStep, structuredDataForFill)
                            End If
                        Else
                            LogMessage("⚠️ 填入Excel跳过: 未识别到有效的结构化数据")
                        End If
                    End If

                    If workflowStep.StepName.Contains("保存工作簿") Then
                        Return SaveWorkbookStep()
                    Else
                        ' 在 VBA 沙箱中执行代码（全静态方法）
                        ' 【修复】工作流步骤由用户主动触发，跳过交互式安全确认弹窗
                        Dim output = Security.VBACodeSandbox.ExecuteCode(code, requireConfirmation:=False)

                        ' 处理输出映射
                        If workflowStep.OutputMapping IsNot Nothing Then
                            For Each mapping In workflowStep.OutputMapping
                                If _variableStore.ContainsKey(mapping.Value) Then
                                    _variableStore(mapping.Key) = _variableStore(mapping.Value)
                                End If
                            Next
                        End If

                        Return StepResult.CreateSuccess(output)
                    End If
                Catch ex As Exception
                    Return StepResult.CreateFailure(FormatExceptionMessage(ex, "代码执行"))
                End Try
            End Function)
        End Function

        ''' <summary>
        ''' 填入发票数据到Excel（兼容入口，统一调用通用方法）
        ''' </summary>
        Private Function FillInvoiceToExcel(workflowStep As WorkflowStep) As StepResult
            If _dataTypeConfigs.ContainsKey("发票") Then
                Return FillStructuredDataToExcel(workflowStep, _dataTypeConfigs("发票"))
            Else
                Return StepResult.CreateFailure("发票类型配置未找到")
            End If
        End Function

        ''' <summary>
        ''' 通用结构化数据填入Excel（配置驱动版本）
        ''' 通过 DataTypeConfiguration 配置，支持所有可扩展的数据类型
        ''' </summary>
        Private Function FillStructuredDataToExcel(workflowStep As WorkflowStep,
                                                        config As DataTypeConfiguration) As StepResult
            Try
                Dim excelService As New Services.ExcelService()
                Dim outputData As New Dictionary(Of String, Object)()

                Dim structuredData As Dictionary(Of String, Object) = Nothing
                If _variableStore.ContainsKey("output.structuredData") Then
                    Try
                        structuredData = CType(_variableStore("output.structuredData"), Dictionary(Of String, Object))
                        System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: 从 output.structuredData 获取数据", config.DisplayName))
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: 转换 structuredData 失败: {1}", config.DisplayName, ex.Message))
                    End Try
                End If

                If structuredData IsNot Nothing Then
                    For Each mapping In config.FieldMappings
                        If structuredData.ContainsKey(mapping.Value) AndAlso structuredData(mapping.Value) IsNot Nothing Then
                            outputData(mapping.Key) = structuredData(mapping.Value)
                            System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: {1} = {2}", config.DisplayName, mapping.Key, structuredData(mapping.Value).ToString()))
                        End If
                    Next
                End If

                If outputData.Count = 0 Then
                    System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: 从变量存储直接获取数据", config.DisplayName))
                    For Each mapping In config.FieldMappings
                        If _variableStore.ContainsKey(mapping.Value) AndAlso _variableStore(mapping.Value) IsNot Nothing Then
                            outputData(mapping.Key) = _variableStore(mapping.Value)
                        End If
                    Next
                End If

                System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: 最终包含 {1} 个字段", config.DisplayName, outputData.Count))

                Dim targetSheet As String = config.DefaultSheetName
                If _variableStore.ContainsKey("input.targetSheet") Then
                    targetSheet = _variableStore("input.targetSheet").ToString()
                End If

                Dim imagePath As String = ""
                If _variableStore.ContainsKey("input.imagePath") Then
                    imagePath = _variableStore("input.imagePath").ToString()
                End If

                Dim fillMode As Services.FillMode = Services.FillMode.Auto
                If _variableStore.ContainsKey("input.fillMode") Then
                    Dim modeStr = _variableStore("input.fillMode").ToString().ToLower()
                    Select Case modeStr
                        Case "createnew", "new", "create" : fillMode = Services.FillMode.CreateNew
                        Case "smartmatch", "match", "smart" : fillMode = Services.FillMode.SmartMatch
                        Case Else : fillMode = Services.FillMode.Auto
                    End Select
                End If

                If config.IncludeImagePath AndAlso Not String.IsNullOrEmpty(imagePath) Then
                    excelService.AddInvoiceRecordSmart(targetSheet, outputData, imagePath, fillMode)
                Else
                    excelService.AddGenericRecordSmart(targetSheet, outputData, fillMode)
                End If

                ' 将提取的数据写回 output.structuredData，让前端能显示
                If outputData.Count > 0 Then
                    _variableStore("output.structuredData") = outputData
                    System.Diagnostics.Debug.WriteLine(String.Format("FillStructuredDataToExcel[{0}]: 已写回 output.structuredData，字段数={1}", config.DisplayName, outputData.Count))
                End If

                Dim resultInfo = New Dictionary(Of String, Object) From {
                    {"dataType", config.TypeName},
                    {"targetSheet", targetSheet},
                    {"fillMode", fillMode.ToString()},
                    {"fieldCount", outputData.Count}
                }
                For Each kvp In outputData
                    resultInfo(kvp.Key) = kvp.Value
                Next

                If workflowStep.OutputMapping IsNot Nothing Then
                    For Each mapping In workflowStep.OutputMapping
                        _variableStore(mapping.Key) = resultInfo
                    Next
                End If

                Return StepResult.CreateSuccess(resultInfo)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, String.Format("填入{0}数据到Excel", config.DisplayName)))
            End Try
        End Function

        ''' <summary>
        ''' 保存工作簿步骤
        ''' </summary>
        Private Function SaveWorkbookStep() As StepResult
            Try
                Dim excelService As New Services.ExcelService()
                excelService.SaveWorkbook()
                Return StepResult.CreateSuccess(Nothing)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "保存工作簿"))
            End Try
        End Function

        ''' <summary>
        ''' 检测数据是否包含名片/身份证字段
        ''' </summary>
        Private Function CheckHasCardFields(data As Object) As Boolean
            Try
                If data Is Nothing Then Return False

                ' 定义名片/身份证的特征字段
                Dim cardFields As String() = {
                    "姓名", "性别", "民族", "出生日期", "出生",
                    "住址", "公民身份号码", "身份证号码", "身份证号",
                    "签发机关", "有效期限",
                    "职位", "公司名称", "公司", "电话", "手机", "邮箱", "地址", "部门"
                }

                ' 尝试通过反射获取属性值
                For Each field In cardFields
                    Try
                        Dim value = CallByName(data, field, CallType.Get)
                        If value IsNot Nothing AndAlso Not String.IsNullOrEmpty(value.ToString()) Then
                            System.Diagnostics.Debug.WriteLine(String.Format("CheckHasCardFields: 找到字段 {0} = {1}", field, value.ToString()))
                            Return True
                        End If
                    Catch ex As Exception  '  字段不存在，继续检查下一个
                    End Try
                Next

                Return False
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("CheckHasCardFields 异常: " & ex.Message)
                Return False
            End Try  ' ✨ 修复：添加缺失的 End Try
        End Function

        ''' <summary>
        ''' 填入名片/身份证数据到Excel（兼容入口，统一调用通用方法）
        ''' </summary>
        Private Async Function FillCardDataToExcel(workflowStep As WorkflowStep) As Task(Of StepResult)
            Return Await Task.Run(Function()
                If _dataTypeConfigs.ContainsKey("名片/身份证") Then
                    Return FillStructuredDataToExcel(workflowStep, _dataTypeConfigs("名片/身份证"))
                Else
                    Return StepResult.CreateFailure("名片/身份证类型配置未找到")
                End If
            End Function)
        End Function

        ''' <summary>
        ''' 通用数据填入Excel（无特定类型配置时使用）
        ''' </summary>
        Private Function FillGenericDataToExcel(workflowStep As WorkflowStep, data As Dictionary(Of String, Object)) As StepResult
            Try
                If data Is Nothing OrElse data.Count = 0 Then
                    Return StepResult.CreateFailure("没有可填入的数据")
                End If

                Dim excelService As New Services.ExcelService()
                Dim targetSheet As String = "识别数据"
                If _variableStore.ContainsKey("input.targetSheet") Then
                    targetSheet = _variableStore("input.targetSheet").ToString()
                End If

                Dim imagePath As String = ""
                If _variableStore.ContainsKey("input.imagePath") Then
                    imagePath = _variableStore("input.imagePath").ToString()
                End If

                Dim fillMode As Services.FillMode = Services.FillMode.Auto
                If _variableStore.ContainsKey("input.fillMode") Then
                    Dim modeStr = _variableStore("input.fillMode").ToString().ToLower()
                    Select Case modeStr
                        Case "createnew", "new", "create" : fillMode = Services.FillMode.CreateNew
                        Case "smartmatch", "match", "smart" : fillMode = Services.FillMode.SmartMatch
                        Case Else : fillMode = Services.FillMode.Auto
                    End Select
                End If

                If Not String.IsNullOrEmpty(imagePath) Then
                    excelService.AddInvoiceRecordSmart(targetSheet, data, imagePath, fillMode)
                Else
                    excelService.AddGenericRecordSmart(targetSheet, data, fillMode)
                End If

                LogMessage(String.Format("通用数据已成功填入Excel（{0}个字段）", data.Count))
                Return StepResult.CreateSuccess(String.Format("成功填入 {0} 个字段到 {1}", data.Count, targetSheet))
            Catch ex As Exception
                LogMessage("通用数据填入Excel失败: " & ex.Message)
                Return StepResult.CreateFailure("通用数据填入失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 执行 API 调用步骤
        ''' </summary>
        Private Async Function ExecuteApiStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                ' [已弃用 - 识别类Skill已移除] 不再走OCR/AI提取/图片识别分支
                ' 检查是否是OCR识别步骤
                ' If workflowStep.StepName.Contains("OCR") Then
                '     Return Await PerformOcrRecognition(workflowStep)
                ' End If

                ' 检查是否是AI提取步骤
                ' If workflowStep.StepName.Contains("AI") OrElse workflowStep.StepName.Contains("提取") Then
                '     Return Await PerformAiExtraction(workflowStep)
                ' End If

                ' 检查是否是通用图片识别
                ' System.Diagnostics.Debug.WriteLine("ExecuteApiStep: 检查 ApiEndpoint, value = " & workflowStep.ApiEndpoint)
                ' If workflowStep.ApiEndpoint = "internal://imageRecognition" Then
                '     System.Diagnostics.Debug.WriteLine("ExecuteApiStep: 匹配 internal://imageRecognition，调用 PerformImageRecognition")
                '     Return Await PerformImageRecognition(workflowStep)
                ' End If

                Dim url = ReplaceVariables(workflowStep.ApiEndpoint)
                Dim method = If(String.IsNullOrEmpty(workflowStep.ApiMethod), "POST", workflowStep.ApiMethod.ToUpper())

                ' 构建请求体
                Dim requestBody As New JObject
                If workflowStep.InputMapping IsNot Nothing Then
                    For Each mapping In workflowStep.InputMapping
                        Dim value = GetVariableValue(mapping.Value)
                        requestBody(mapping.Key) = JToken.FromObject(value)
                    Next
                End If

                Dim content = New StringContent(requestBody.ToString(), Encoding.UTF8, "application/json")
                Dim response As HttpResponseMessage

                Select Case method
                    Case "GET"
                        response = Await _httpClient.GetAsync(url)
                    Case "POST"
                        response = Await _httpClient.PostAsync(url, content)
                    Case "PUT"
                        response = Await _httpClient.PutAsync(url, content)
                    Case "DELETE"
                        response = Await _httpClient.DeleteAsync(url)
                    Case Else
                        Return StepResult.CreateFailure(String.Format("不支持的 HTTP 方法: {0}", method))
                End Select

                If response.IsSuccessStatusCode Then
                    Dim responseBody = Await response.Content.ReadAsStringAsync()
                    Dim jsonResponse = JObject.Parse(responseBody)

                    ' 处理输出映射
                    If workflowStep.OutputMapping IsNot Nothing Then
                        For Each mapping In workflowStep.OutputMapping
                            If jsonResponse.ContainsKey(mapping.Value) Then
                                _variableStore(mapping.Key) = jsonResponse(mapping.Value).ToObject(Of Object)()
                            End If
                        Next
                    End If

                    Return StepResult.CreateSuccess(jsonResponse)
                Else
                    Return StepResult.CreateFailure(String.Format("API 调用失败: {0}", response.StatusCode))
                End If

            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "API 调用"))
            End Try
        End Function

        ''' <summary>
        ''' 执行子 Skill 步骤
        ''' </summary>
        Private Async Function ExecuteSubSkillStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                Dim subSkill = _skillManager.GetSkill(workflowStep.SubSkillName)
                If subSkill Is Nothing Then
                    Return StepResult.CreateFailure(String.Format("找不到子 Skill: {0}", workflowStep.SubSkillName))
                End If

                ' 构建子 Skill 输入参数
                Dim subSkillInputs As New Dictionary(Of String, Object)
                If workflowStep.InputMapping IsNot Nothing Then
                    For Each mapping In workflowStep.InputMapping
                        subSkillInputs(mapping.Key) = GetVariableValue(mapping.Value)
                    Next
                End If

                If subSkill.IsComposite Then
                    ' 递归执行复合 Skill，传递进度报告器
                    Dim subResult = Await ExecuteWorkflow(subSkill, subSkillInputs, _progress)
                    If subResult.Success Then
                        ' 处理输出映射
                        If workflowStep.OutputMapping IsNot Nothing Then
                            For Each mapping In workflowStep.OutputMapping
                                If subResult.Outputs.ContainsKey(mapping.Value) Then
                                    _variableStore(mapping.Key) = subResult.Outputs(mapping.Value)
                                End If
                            Next
                        End If
                        Return StepResult.CreateSuccess(subResult.Outputs)
                    Else
                        Return StepResult.CreateFailure(subResult.ErrorMessage)
                    End If
                Else
                    ' 简单 Skill，直接返回模板代码
                    Return StepResult.CreateSuccess(subSkill.TemplateCode)
                End If

            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "子 Skill 执行"))
            End Try
        End Function

        ''' <summary>
        ''' 执行条件判断步骤
        ''' </summary>
        Private Function ExecuteConditionStep(workflowStep As WorkflowStep) As StepResult
            Try
                Dim condition = ReplaceVariables(workflowStep.Condition)
                ' 简单条件评估（可以扩展为更复杂的表达式引擎）
                Dim evalResult = EvaluateCondition(condition)

                Return StepResult.CreateSuccess(evalResult)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "条件判断"))
            End Try
        End Function

        ''' <summary>
        ''' 执行循环步骤
        ''' </summary>
        Private Function ExecuteLoopStep(workflowStep As WorkflowStep) As StepResult
            Try
                ' 【v2修复】强制使用 MaxLoopIterations 上限，防止LLM生成的配置造成死循环
                Dim maxIter = If(workflowStep.MaxLoopIterations > 0, workflowStep.MaxLoopIterations, 1000)
                Dim iterVar = If(workflowStep.Config?.ContainsKey("iteratorVar") = True, CStr(workflowStep.Config("iteratorVar")), "loop.i")
                Dim collectionVar = If(workflowStep.Config?.ContainsKey("collection") = True, CStr(workflowStep.Config("collection")), Nothing)
                Dim iterCount = 0
                Dim items As IEnumerable(Of Object) = Nothing

                If Not String.IsNullOrEmpty(collectionVar) AndAlso _variableStore.ContainsKey(collectionVar) Then
                    items = TryCast(_variableStore(collectionVar), IEnumerable(Of Object))
                    If items IsNot Nothing Then
                        Dim limit = Math.Min(items.Count(), maxIter)
                        For Each it In items.Take(limit)
                            _variableStore(iterVar) = it
                            iterCount += 1
                        Next
                    End If
                End If

                _variableStore("loop.totalIterations") = iterCount
                If iterCount = maxIter Then
                    LogMessage($"⚠️ Loop达到上限MaxLoopIterations={maxIter}，已提前终止。如需更大范围请调大workflowStep.MaxLoopIterations。")
                End If
                Return StepResult.CreateSuccess(iterCount)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "循环步骤"))
            End Try
        End Function

        ''' <summary>
        ''' 执行等待步骤
        ''' </summary>
        Private Async Function ExecuteWaitStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                Dim waitMs = workflowStep.Timeout * 1000
                Await Task.Delay(waitMs)
                Return StepResult.CreateSuccess(Nothing)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "等待步骤"))
            End Try
        End Function

        ''' <summary>
        ''' 【优先级2扩展】执行 MCP 工具调用步骤
        ''' 通过 McpBridge 调用已连接的 MCP Server 上的外部工具
        '''
        ''' 使用方式（在 Skill 的 WorkflowSteps JSON 中定义）：
        ''' {
        '''   "stepId": "mcp_list_files",
        '''   "stepName": "列出文件",
        '''   "stepType": "InvokeMcpTool",
        '''   "mcpClientName": "filesystem",    ' mcp.json 中定义的服务名
        '''   "mcpToolName": "list_directory",   ' 具体工具名
        '''   "mcpArguments": "{\"path\": \"{{input.filePath}}\"}",  ' 支持变量占位符
        '''   "outputMapping": {"files": "output.fileList"}
        ''' }
        ''' </summary>
        Private Async Function ExecuteMcpToolStep(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                ' 验证必要参数
                If String.IsNullOrWhiteSpace(workflowStep.McpClientName) Then
                    Return StepResult.CreateFailure("InvokeMcpTool 步骤缺少 mcpClientName 参数")
                End If
                If String.IsNullOrWhiteSpace(workflowStep.McpToolName) Then
                    Return StepResult.CreateFailure("InvokeMcpTool 步骤缺少 mcpToolName 参数")
                End If

                LogMessage(String.Format("调用 MCP 工具: client={0}, tool={1}",
                    workflowStep.McpClientName, workflowStep.McpToolName))

                ' 解析参数（支持变量占位符替换）
                Dim argumentsJson As String = workflowStep.McpArguments
                If Not String.IsNullOrWhiteSpace(argumentsJson) Then
                    argumentsJson = ReplaceVariables(argumentsJson)
                Else
                    argumentsJson = "{}"
                End If

                ' 经依赖倒置解析 McpBridge 接口调用 MCP 工具
                Dim mcpBridge = Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IMcpBridge)()
                Dim resultJson As String = Await mcpBridge.CallMcpToolAsync(
                    workflowStep.McpClientName,
                    workflowStep.McpToolName,
                    argumentsJson)

                ' 解析返回结果
                Dim responseObj = JObject.Parse(resultJson)

                ' 检查是否为错误响应
                If responseObj.ContainsKey("success") AndAlso
                   responseObj("success")?.ToString() = "False" Then
                    Dim errorMsg = ""
                    If responseObj.ContainsKey("error") Then
                        errorMsg = responseObj("error")?.ToString()
                    End If
                    Return StepResult.CreateFailure(String.Format(
                        "MCP 工具调用失败 [{0}/{1}]: {2}",
                        workflowStep.McpClientName, workflowStep.McpToolName,
                        If(errorMsg, "未知错误")))
                End If

                ' 将结果存入变量存储
                _variableStore(String.Format("output.mcp_{0}_{1}",
                    workflowStep.McpClientName, workflowStep.McpToolName)) = responseObj

                ' 处理输出映射
                If workflowStep.OutputMapping IsNot Nothing Then
                    For Each mapping In workflowStep.OutputMapping
                        If responseObj.ContainsKey(mapping.Value) Then
                            _variableStore(mapping.Key) =
                                responseObj(mapping.Value).ToObject(Of Object)()
                        ElseIf mapping.Value = "$raw" Then
                            ' $raw 特殊值：将整个结果映射到指定变量
                            _variableStore(mapping.Key) = resultJson
                        End If
                    Next
                End If

                LogMessage(String.Format("MCP 工具调用成功: {0}/{1}",
                    workflowStep.McpClientName, workflowStep.McpToolName))

                Return StepResult.CreateSuccess(responseObj)

            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, String.Format(
                    "MCP工具调用 [{0}/{1}]",
                    If(workflowStep?.McpClientName, "unknown"),
                    If(workflowStep?.McpToolName, "unknown"))))
            End Try
        End Function

        ''' <summary>
        ''' 验证输入参数
        ''' </summary>
        Private Sub ValidateInputs(skill As SkillDefinition, inputs As Dictionary(Of String, Object))
            If skill.InputParameters Is Nothing Then Return

            For Each param In skill.InputParameters
                If param.IsRequired AndAlso Not inputs.ContainsKey(param.Name) Then
                    Throw New ArgumentException(String.Format("缺少必需参数: {0}", param.Name))
                End If
            Next
        End Sub

        ''' <summary>
        ''' 初始化变量存储
        ''' </summary>
        Private Sub InitializeVariables(skill As SkillDefinition, inputs As Dictionary(Of String, Object))
            ' 设置输入参数
            If inputs IsNot Nothing Then
                For Each inputParam In inputs
                    _variableStore(String.Format("input.{0}", inputParam.Key)) = inputParam.Value
                Next
            End If

            ' 设置默认值
            If skill.InputParameters IsNot Nothing Then
                For Each param In skill.InputParameters
                    If Not _variableStore.ContainsKey(String.Format("input.{0}", param.Name)) AndAlso param.DefaultValue IsNot Nothing Then
                        _variableStore(String.Format("input.{0}", param.Name)) = param.DefaultValue
                    End If
                Next
            End If
        End Sub

        ''' <summary>
        ''' 收集输出参数
        ''' </summary>
        Private Sub CollectOutputs(skill As SkillDefinition, result As WorkflowResult)
            If skill.OutputParameters Is Nothing Then Return

            For Each param In skill.OutputParameters
                Dim key = String.Format("output.{0}", param.Name)
                If _variableStore.ContainsKey(key) Then
                    result.Outputs(param.Name) = _variableStore(key)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 替换变量占位符
        ''' </summary>
        Private Function ReplaceVariables(text As String) As String
            If String.IsNullOrEmpty(text) Then Return text

            Dim result = text

            ' 替换 {{input.xxx}} 格式的变量
            For Each variable In _variableStore
                Dim placeholder = String.Format("{{{{{0}}}}}", variable.Key)
                If result.Contains(placeholder) Then
                    Dim varValue As String = ""
                    If variable.Value IsNot Nothing Then
                        varValue = variable.Value.ToString()
                    End If
                    result = result.Replace(placeholder, varValue)
                End If
            Next

            ' 替换 {{vars.xxx}} 格式的变量
            For Each variable In _variableStore
                Dim shortKey = variable.Key.Replace("input.", "vars.").Replace("output.", "vars.")
                Dim placeholder = String.Format("{{{{{0}}}}}", shortKey)
                If result.Contains(placeholder) Then
                    Dim varValue As String = ""
                    If variable.Value IsNot Nothing Then
                        varValue = variable.Value.ToString()
                    End If
                    result = result.Replace(placeholder, varValue)
                End If
            Next

            Return result
        End Function

        ''' <summary>
        ''' 获取变量值
        ''' </summary>
        Private Function GetVariableValue(key As String) As Object
            If _variableStore.ContainsKey(key) Then
                Return _variableStore(key)
            End If

            ' 尝试不同的前缀
            If Not key.StartsWith("input.") AndAlso Not key.StartsWith("output.") AndAlso Not key.StartsWith("vars.") Then
                If _variableStore.ContainsKey(String.Format("input.{0}", key)) Then
                    Return _variableStore(String.Format("input.{0}", key))
                End If
                If _variableStore.ContainsKey(String.Format("vars.{0}", key)) Then
                    Return _variableStore(String.Format("vars.{0}", key))
                End If
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 评估条件表达式
        ''' 支持比较运算符(==, !=, &gt;, &lt;, &gt;=, &lt;=)、逻辑运算符(&amp;&amp;, ||, !)、
        ''' 字符串字面量('xx'或"xx")、数值、布尔常量、括号分组。
        ''' </summary>
        Private Function EvaluateCondition(condition As String) As Boolean
            If String.IsNullOrWhiteSpace(condition) Then Return True

            Try
                Dim tokens = TokenizeCondition(condition)
                Dim result As Object = Nothing
                Dim pos As Integer = 0
                result = ParseExpression(tokens, pos)
                If TypeOf result Is Boolean Then
                    Return CBool(result)
                End If
                Return result IsNot Nothing
            Catch ex As Exception
                LogMessage(String.Format("条件表达式评估失败 [{0}]: {1}", condition, ex.Message))
                Return True ' 安全兜底：表达式无法解析时允许继续执行
            End Try
        End Function

        ' ======================== 条件表达式词法/语法解析器 ========================

        Private Class TokenType
            Public Const T_STRING As String = "STRING"
            Public Const T_NUMBER As String = "NUMBER"
            Public Const T_BOOL As String = "BOOL"
            Public Const OP_EQ As String = "=="
            Public Const OP_NEQ As String = "!="
            Public Const OP_GE As String = ">="
            Public Const OP_LE As String = "<="
            Public Const OP_GT As String = ">"
            Public Const OP_LT As String = "<"
            Public Const OP_AND As String = "&&"
            Public Const OP_OR As String = "||"
            Public Const OP_NOT As String = "!"
            Public Const LPAREN As String = "("
            Public Const RPAREN As String = ")"
            Public Const EOF As String = "EOF"
        End Class

        Private Class Token
            Public Property Type As String
            Public Property Value As Object
            Public Sub New(t As String, v As Object)
                Type = t
                Value = v
            End Sub
        End Class

        ''' <summary>
        ''' 将条件字符串拆分为 Token 列表
        ''' </summary>
        Private Function TokenizeCondition(input As String) As List(Of Token)
            Dim tokens As New List(Of Token)()
            Dim i As Integer = 0
            Dim n As Integer = input.Length

            While i < n
                Dim ch As Char = input(i)

                ' 跳过空白
                If Char.IsWhiteSpace(ch) Then
                    i += 1
                    Continue While
                End If

                ' 字符串字面量: 'xxx' 或 "xxx"
                If ch = "'"c OrElse ch = """"c Then
                    Dim quote As Char = ch
                    Dim sb As New System.Text.StringBuilder()
                    i += 1
                    While i < n AndAlso input(i) <> quote
                        sb.Append(input(i))
                        i += 1
                    End While
                    If i < n Then i += 1 ' 跳过关闭引号
                    tokens.Add(New Token(TokenType.T_STRING, sb.ToString()))
                    Continue While
                End If

                ' 两字符运算符: == != >= <= && ||
                If i + 1 < n Then
                    Dim twoCh As String = input.Substring(i, 2)
                    If twoCh = "==" OrElse twoCh = "!=" OrElse
                       twoCh = ">=" OrElse twoCh = "<=" OrElse
                       twoCh = "&&" OrElse twoCh = "||" Then
                        tokens.Add(New Token(twoCh, Nothing))
                        i += 2
                        Continue While
                    End If
                End If

                ' 单字符运算符
                If ch = ">"c OrElse ch = "<"c OrElse ch = "!"c Then
                    tokens.Add(New Token(ch.ToString(), Nothing))
                    i += 1
                    Continue While
                End If
                If ch = "("c Then
                    tokens.Add(New Token(TokenType.LPAREN, Nothing))
                    i += 1
                    Continue While
                End If
                If ch = ")"c Then
                    tokens.Add(New Token(TokenType.RPAREN, Nothing))
                    i += 1
                    Continue While
                End If

                ' 数字字面量
                If Char.IsDigit(ch) OrElse (ch = "-"c AndAlso i + 1 < n AndAlso Char.IsDigit(input(i + 1))) Then
                    Dim sb As New System.Text.StringBuilder()
                    If ch = "-"c Then
                        sb.Append(ch)
                        i += 1
                    End If
                    While i < n AndAlso (Char.IsDigit(input(i)) OrElse input(i) = "."c)
                        sb.Append(input(i))
                        i += 1
                    End While
                    tokens.Add(New Token(TokenType.T_NUMBER, CDbl(sb.ToString())))
                    Continue While
                End If

                ' 标识符: true / false
                If Char.IsLetter(ch) Then
                    Dim sb As New System.Text.StringBuilder()
                    While i < n AndAlso (Char.IsLetterOrDigit(input(i)) OrElse input(i) = "_"c)
                        sb.Append(input(i))
                        i += 1
                    End While
                    Dim ident As String = sb.ToString().ToLowerInvariant()
                    If ident = "true" Then
                        tokens.Add(New Token(TokenType.T_BOOL, True))
                    ElseIf ident = "false" Then
                        tokens.Add(New Token(TokenType.T_BOOL, False))
                    Else
                        ' 未知标识符视为字符串值（用于变量展开后的结果）
                        tokens.Add(New Token(TokenType.T_STRING, sb.ToString()))
                    End If
                    Continue While
                End If

                ' 遇到无法识别的字符，跳过
                i += 1
            End While

            tokens.Add(New Token(TokenType.EOF, Nothing))
            Return tokens
        End Function

        ''' <summary>
        ''' 递归下降解析：表达式 → 逻辑或
        ''' </summary>
        Private Function ParseExpression(tokens As List(Of Token), ByRef pos As Integer) As Object
            Dim left = ParseAndOr(tokens, pos)
            Return left
        End Function

        ''' <summary>
        ''' 处理 || 逻辑或
        ''' </summary>
        Private Function ParseAndOr(tokens As List(Of Token), ByRef pos As Integer) As Object
            Dim left = ParseAnd(tokens, pos)

            While tokens(pos).Type = TokenType.OP_OR
                pos += 1
                Dim right = ParseAnd(tokens, pos)
                left = CBool(Coerce(left, True)) OrElse CBool(Coerce(right, True))
            End While

            Return left
        End Function

        ''' <summary>
        ''' 处理 AND 逻辑与
        ''' </summary>
        Private Function ParseAnd(tokens As List(Of Token), ByRef pos As Integer) As Object
            Dim left = ParseUnary(tokens, pos)

            While tokens(pos).Type = TokenType.OP_AND
                pos += 1
                Dim right = ParseUnary(tokens, pos)
                left = CBool(Coerce(left, True)) AndAlso CBool(Coerce(right, True))
            End While

            Return left
        End Function

        ''' <summary>
        ''' 处理一元 NOT
        ''' </summary>
        Private Function ParseUnary(tokens As List(Of Token), ByRef pos As Integer) As Object
            If tokens(pos).Type = TokenType.OP_NOT Then
                pos += 1
                Dim operand = ParseUnary(tokens, pos)
                Return Not CBool(Coerce(operand, True))
            End If
            Return ParseComparison(tokens, pos)
        End Function

        ''' <summary>
        ''' 处理比较运算: == != &gt; &lt; &gt;= &lt;=
        ''' </summary>
        Private Function ParseComparison(tokens As List(Of Token), ByRef pos As Integer) As Object
            Dim left = ParsePrimary(tokens, pos)

            Dim opType As String = tokens(pos).Type
            If opType = TokenType.OP_EQ OrElse opType = TokenType.OP_NEQ OrElse
               opType = TokenType.OP_GT OrElse opType = TokenType.OP_LT OrElse
               opType = TokenType.OP_GE OrElse opType = TokenType.OP_LE Then
                pos += 1
                Dim right = ParsePrimary(tokens, pos)
                Return CompareValues(left, right, opType)
            End If

            Return left
        End Function

        ''' <summary>
        ''' 解析基本值: 字面量 或 括号分组
        ''' </summary>
        Private Function ParsePrimary(tokens As List(Of Token), ByRef pos As Integer) As Object
            If tokens(pos).Type = TokenType.LPAREN Then
                pos += 1
                Dim result = ParseAndOr(tokens, pos)
                If tokens(pos).Type = TokenType.RPAREN Then
                    pos += 1
                End If
                Return result
            End If

            If tokens(pos).Type = TokenType.T_STRING OrElse
               tokens(pos).Type = TokenType.T_NUMBER OrElse
               tokens(pos).Type = TokenType.T_BOOL Then
                Dim val = tokens(pos).Value
                pos += 1
                Return val
            End If

            ' 未匹配到有效 token，返回 Nothing
            pos += 1
            Return Nothing
        End Function

        ''' <summary>
        ''' 将值强制转换为指定类型（用于布尔上下文）
        ''' </summary>
        Private Function Coerce(value As Object, Optional forBool As Boolean = False) As Object
            If value Is Nothing Then Return If(forBool, False, Nothing)
            If TypeOf value Is Boolean Then Return value
            If TypeOf value Is Double Then
                If forBool Then Return CDbl(value) <> 0.0
                Return value
            End If
            If TypeOf value Is String Then
                If forBool Then Return Not String.IsNullOrEmpty(CStr(value))
                ' 尝试解析为数字
                Dim d As Double
                If Double.TryParse(CStr(value), Globalization.NumberStyles.Float,
                                   Globalization.CultureInfo.InvariantCulture, d) Then
                    Return d
                End If
                Return value
            End If
            Return value
        End Function

        ''' <summary>
        ''' 比较两个值（支持字符串/数值混合比较）
        ''' </summary>
        Private Function CompareValues(left As Object, right As Object, op As String) As Boolean
            ' 尝试数值比较
            Dim lIsNum As Boolean = TypeOf left Is Double
            Dim rIsNum As Boolean = TypeOf right Is Double

            If lIsNum AndAlso rIsNum Then
                Dim l As Double = CDbl(left), r As Double = CDbl(right)
                Select Case op
                    Case TokenType.OP_EQ : Return l = r
                    Case TokenType.OP_NEQ : Return l <> r
                    Case TokenType.OP_GT : Return l > r
                    Case TokenType.OP_LT : Return l < r
                    Case TokenType.OP_GE : Return l >= r
                    Case TokenType.OP_LE : Return l <= r
                End Select
            End If

            ' 字符串比较
            Dim sl As String = If(left Is Nothing, "", left.ToString())
            Dim sr As String = If(right Is Nothing, "", right.ToString())
            Select Case op
                Case TokenType.OP_EQ : Return String.Equals(sl, sr, StringComparison.OrdinalIgnoreCase)
                Case TokenType.OP_NEQ : Return Not String.Equals(sl, sr, StringComparison.OrdinalIgnoreCase)
                Case TokenType.OP_GT : Return String.Compare(sl, sr, StringComparison.OrdinalIgnoreCase) > 0
                Case TokenType.OP_LT : Return String.Compare(sl, sr, StringComparison.OrdinalIgnoreCase) < 0
                Case TokenType.OP_GE : Return String.Compare(sl, sr, StringComparison.OrdinalIgnoreCase) >= 0
                Case TokenType.OP_LE : Return String.Compare(sl, sr, StringComparison.OrdinalIgnoreCase) <= 0
            End Select

            Return False
        End Function

        ''' <summary>
        ''' 格式化异常信息（包含完整堆栈和内部异常）
        ''' </summary>
        Private Function FormatExceptionMessage(ex As Exception, Optional context As String = "") As String
            Dim sb As New System.Text.StringBuilder()

            If Not String.IsNullOrEmpty(context) Then
                sb.AppendFormat("[{0}] ", context)
            End If

            sb.AppendLine(ex.Message)

            If ex.InnerException IsNot Nothing Then
                sb.AppendLine()
                sb.AppendLine("内部异常:")
                sb.AppendLine(ex.InnerException.Message)

                If ex.InnerException.InnerException IsNot Nothing Then
                    sb.AppendLine("根本原因:")
                    sb.AppendLine(ex.InnerException.InnerException.Message)
                End If
            End If

            sb.AppendLine()
            sb.AppendLine("堆栈跟踪:")
            sb.AppendLine(ex.StackTrace)

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 记录日志
        ''' </summary>
        Private Sub LogMessage(message As String)
            Dim logEntry = String.Format("[{0:HH:mm:ss.fff}] {1}", DateTime.Now, message)
            _executionLog.Add(logEntry)
            System.Diagnostics.Debug.WriteLine(logEntry)
            ' 同时发送进度消息到前端
            _progress?.Report(message)
        End Sub

        ' =========================================================
        ' v2 统一体系：原 CustomStepType 中的领域步骤处理器
        ' （把 CustomSkillExecutor.ExecuteWorkflowStep 里的分支吸收到统一 WorkflowEngine.ExecuteStep 路径下）
        ' =========================================================

        ''' <summary>
        ''' UploadImage步骤：校验图片路径存在，存入变量
        '''   Config: {"imagePath": "..."} 或变量 input.imagePath
        ''' </summary>
        Private Function ExecuteUploadImageStep(workflowStep As WorkflowStep) As StepResult
            Try
                Dim imagePath As String = ResolveConfigOrVar(workflowStep, "imagePath", "input.imagePath")
                If String.IsNullOrWhiteSpace(imagePath) OrElse Not File.Exists(imagePath) Then
                    Return StepResult.CreateFailure("UploadImage: 图片路径不存在 '" & imagePath & "'")
                End If
                _variableStore("input.imagePath") = imagePath
                _variableStore("input.imageName") = Path.GetFileName(imagePath)
                Return StepResult.CreateSuccess(imagePath)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "UploadImage步骤"))
            End Try
        End Function

        ''' <summary>
        ''' OCRRecognition步骤：调用OCR服务识别图片文本
        '''   Config: {"imagePath": "..."}
        '''   输出: output.ocrText
        ''' </summary>
        Private Async Function ExecuteOcrRecognitionStepAsync(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                Dim imagePath As String = ResolveConfigOrVar(workflowStep, "imagePath", "input.imagePath")
                If String.IsNullOrWhiteSpace(imagePath) OrElse Not File.Exists(imagePath) Then
                    Return StepResult.CreateFailure("OCRRecognition: 图片路径无效 '" & imagePath & "'。请确保 UploadImage 先执行或 input.imagePath 已设置。")
                End If
                LogMessage("OCR识别: " & Path.GetFileName(imagePath))
                Dim ocrText As String = Await Services.OcrServiceManager.Instance.RecognizeText(imagePath)
                _variableStore("output.ocrText") = If(ocrText, String.Empty)
                _variableStore("ocr.pageCount") = 1
                Return StepResult.CreateSuccess(ocrText)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "OCRRecognition步骤"))
            End Try
        End Function

        ''' <summary>
        ''' ExtractFields步骤：对已识别的文本做字段抽取（优先使用LLM）
        '''   Config: {"schema": "可选JSONSchema字符串", "useLlm": true, "prompt": "可选自定义提示词"}
        '''   输入: output.ocrText
        '''   输出: output.structuredData
        ''' </summary>
        Private Function ExecuteExtractFieldsStep(workflowStep As WorkflowStep) As StepResult
            Try
                Dim ocrText As String = If(_variableStore.ContainsKey("output.ocrText"), CStr(_variableStore("output.ocrText")), String.Empty)
                If String.IsNullOrWhiteSpace(ocrText) Then
                    Return StepResult.CreateFailure("ExtractFields: output.ocrText 为空，请先执行 OCRRecognition")
                End If

                ' 若配置了自定义LLM提示词，优先走MultimodalLLMEngine
                Dim cfg As Dictionary(Of String, Object) = workflowStep.Config
                Dim useLlm As Boolean = True
                Dim customPrompt As String = Nothing
                If cfg IsNot Nothing Then
                    If cfg.ContainsKey("useLlm") Then useLlm = CBool(cfg("useLlm"))
                    If cfg.ContainsKey("prompt") Then customPrompt = CStr(cfg("prompt"))
                End If

                If useLlm AndAlso _llmEngine IsNot Nothing Then
                    Dim prompt As String = If(Not String.IsNullOrEmpty(customPrompt),
                                              ReplaceVariables(customPrompt),
                                              "请从以下OCR文本中抽取结构化JSON字段，保留所有数字/日期/金额格式：" & vbCrLf & ocrText)
                    Try
                        Dim t As Task(Of LLMResponse) = _llmEngine.ProcessContentAsync(
                            New List(Of MultimodalContent) From {MultimodalContent.CreateText(prompt)})
                        t.Wait()
                        Dim llmRaw As LLMResponse = t.Result
                        Dim resultJson As String = If(llmRaw IsNot Nothing, llmRaw.Content, Nothing)
                        If Not String.IsNullOrWhiteSpace(resultJson) Then
                            Try
                                Dim parsed As JObject = JObject.Parse(resultJson)
                                _variableStore("output.structuredData") = parsed.ToObject(Of Dictionary(Of String, Object))()
                                Return StepResult.CreateSuccess(_variableStore("output.structuredData"))
                            Catch ex As Exception  ' 
                                _variableStore("output.extractedText") = resultJson
                            End Try
                        End If
                    Catch ex As Exception  '  LLM异常时走下面的正则Fallback，不吞整体异常
                    End Try
                End If

                ' Fallback：简单正则抽取（兜底）
                Dim simple = New Dictionary(Of String, Object)()
                For Each m As Match In Regex.Matches(ocrText, "(?<k>[\u4e00-\u9fa5A-Za-z0-9]{2,20})[:：]\s*(?<v>[^\r\n]{1,80})")
                    simple(m.Groups("k").Value) = m.Groups("v").Value.Trim()
                Next
                _variableStore("output.structuredData") = simple
                Return StepResult.CreateSuccess(simple)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "ExtractFields步骤"))
            End Try
        End Function

        ''' <summary>
        ''' WriteToExcel步骤：按Skill.FieldMappings把output.structuredData 映射写入 Excel
        '''   Config: {"targetCell": "A1", "clearBeforeWrite": false, "sheetName": ""}
        ''' </summary>
        Private Function ExecuteWriteToExcelStep(workflowStep As WorkflowStep) As StepResult
            Try
                Dim structured As Dictionary(Of String, Object) = Nothing
                If _variableStore.ContainsKey("output.structuredData") Then
                    structured = TryCast(_variableStore("output.structuredData"), Dictionary(Of String, Object))
                End If
                Dim fieldMaps As List(Of FieldMapping) = If(_currentSkill IsNot Nothing, _currentSkill.FieldMappings, Nothing)
                If fieldMaps Is Nothing OrElse fieldMaps.Count = 0 Then
                    Return StepResult.CreateFailure("WriteToExcel: Skill未定义 FieldMappings 映射表")
                End If

                Dim total As Integer = 0
                Dim excelApp As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.GetExcelApp()
                If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                    Return StepResult.CreateFailure("WriteToExcel: 当前没有打开的工作簿")
                End If

                ' ======= 【P0 Bug修复】Boolean? → 真正的 String sheetName =======
                Dim sheetName As String = Nothing
                Dim cfg As Dictionary(Of String, Object) = workflowStep.Config
                If cfg IsNot Nothing AndAlso cfg.ContainsKey("sheetName") Then
                    sheetName = TryCast(cfg("sheetName"), String)
                End If

                ' ======= 【后期绑定修复】ws As Worksheet，Range CType =======
                Dim ws As Microsoft.Office.Interop.Excel.Worksheet
                If Not String.IsNullOrWhiteSpace(sheetName) Then
                    ws = TryCast(excelApp.ActiveWorkbook.Sheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)
                    If ws Is Nothing Then
                        Return StepResult.CreateFailure("WriteToExcel: 找不到工作表 '" & sheetName & "'")
                    End If
                Else
                    ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                    If ws Is Nothing Then
                        Return StepResult.CreateFailure("WriteToExcel: 当前ActiveSheet不是Worksheet")
                    End If
                End If

                For Each fm As FieldMapping In fieldMaps
                    Dim valObj As Object = Nothing
                    If structured IsNot Nothing AndAlso structured.ContainsKey(fm.SourceField) Then
                        valObj = structured(fm.SourceField)
                    End If
                    Dim valStr As String = If(valObj Is Nothing, String.Empty, valObj.ToString())
                    If Not String.IsNullOrWhiteSpace(fm.Format) Then
                        valStr = ApplyFormat(valStr, fm.Format)
                    End If
                    Dim target As String = If(Not String.IsNullOrWhiteSpace(fm.TargetCell), fm.TargetCell, Nothing)
                    If target IsNot Nothing Then
                        Dim rng As Microsoft.Office.Interop.Excel.Range = CType(ws.Range(target), Microsoft.Office.Interop.Excel.Range)
                        rng.Value = valStr
                        total += 1
                    End If
                Next

                _variableStore("output.writtenFieldCount") = total
                LogMessage("WriteToExcel: 写入 " & total & " 个字段")
                Return StepResult.CreateSuccess(total)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "WriteToExcel步骤"))
            End Try
        End Function

        ''' <summary>
        ''' CustomLLM步骤：执行自定义提示词，将输出写入 output.llmResponse
        '''   Config: {"prompt": "...支持{{变量}}...", "temperature": 0.2}
        ''' </summary>
        Private Async Function ExecuteCustomLlmStepAsync(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                Dim cfg As Dictionary(Of String, Object) = workflowStep.Config
                If cfg Is Nothing OrElse Not cfg.ContainsKey("prompt") Then
                    Return StepResult.CreateFailure("CustomLLM: 必须在Config.prompt中指定提示词（支持{{变量}}占位符）")
                End If
                If _llmEngine Is Nothing Then
                    Return StepResult.CreateFailure("CustomLLM: MultimodalLLMEngine 未初始化，请确认配置")
                End If
                Dim prompt As String = ReplaceVariables(CStr(cfg("prompt")))
                Dim contents As New List(Of MultimodalContent) From {MultimodalContent.CreateText(prompt)}
                Dim llmResult As LLMResponse = Await _llmEngine.ProcessContentAsync(contents)
                If llmResult Is Nothing OrElse Not llmResult.Success Then
                    Return StepResult.CreateFailure("CustomLLM: " & If(llmResult IsNot Nothing, llmResult.ErrorMessage, "未知错误"))
                End If
                _variableStore("output.llmResponse") = llmResult.Content
                Return StepResult.CreateSuccess(llmResult.Content)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "CustomLLM步骤"))
            End Try
        End Function

        ''' <summary>
        ''' LLMProcess/MultimodalProcess步骤：统一走MultimodalLLMEngine，可带图片附件
        '''   Config: {"attachments": ["路径1","路径2"]} 或变量 input.attachments
        ''' </summary>
        Private Async Function ExecuteLlmProcessStepAsync(workflowStep As WorkflowStep) As Task(Of StepResult)
            Try
                If _llmEngine Is Nothing Then
                    Return StepResult.CreateFailure("LlmProcess: MultimodalLLMEngine 未初始化，请确认配置")
                End If
                Dim fallbackQuery As String = If(_variableStore.ContainsKey("input.userQuery"),
                                                 CStr(_variableStore("input.userQuery")),
                                                 String.Empty)
                Dim prompt As String = ReplaceVariables(
                    If(String.IsNullOrEmpty(workflowStep.Code), fallbackQuery, workflowStep.Code))
                Dim contents As New List(Of MultimodalContent)()

                If workflowStep.StepType = WorkflowStepType.MultimodalProcess Then
                    Dim atts As IEnumerable(Of String) = Nothing
                    If _variableStore.ContainsKey("input.attachments") Then
                        atts = TryCast(_variableStore("input.attachments"), IEnumerable(Of String))
                    End If
                    Dim cfg As Dictionary(Of String, Object) = workflowStep.Config
                    If atts Is Nothing AndAlso cfg IsNot Nothing AndAlso cfg.ContainsKey("attachments") Then
                        atts = TryCast(cfg("attachments"), IEnumerable(Of String))
                    End If
                    If atts IsNot Nothing Then
                        For Each p As String In atts
                            If File.Exists(p) Then
                                Try
                                    Dim fileBytes As Byte() = File.ReadAllBytes(p)
                                    Dim ext As String = Path.GetExtension(p).ToLowerInvariant()
                                    Dim mime As String = "image/png"
                                    Select Case ext
                                        Case ".jpg", ".jpeg" : mime = "image/jpeg"
                                        Case ".gif" : mime = "image/gif"
                                        Case ".bmp" : mime = "image/bmp"
                                        Case ".webp" : mime = "image/webp"
                                    End Select
                                    contents.Add(MultimodalContent.CreateImageFile(Path.GetFileName(p), fileBytes, mime))
                                Catch ex As Exception  '  单张图片读取失败时跳过，不影响整体
                                End Try
                            End If
                        Next
                    End If
                End If
                contents.Add(MultimodalContent.CreateText(prompt))

                Dim result As LLMResponse = Await _llmEngine.ProcessContentAsync(contents)
                If result Is Nothing OrElse Not result.Success Then
                    Return StepResult.CreateFailure("LlmProcess: " & If(result IsNot Nothing, result.ErrorMessage, "未知错误"))
                End If
                _variableStore("output.llmResponse") = result.Content
                Return StepResult.CreateSuccess(result.Content)
            Catch ex As Exception
                Return StepResult.CreateFailure(FormatExceptionMessage(ex, "LLMProcess步骤"))
            End Try
        End Function

        ' ======== v2 统一体系：辅助工具方法 ========

        ''' <summary>解析Config中的键（优先）或变量存储中的键</summary>
        Private Function ResolveConfigOrVar(workflowStep As WorkflowStep, configKey As String, varKey As String) As String
            If workflowStep.Config?.ContainsKey(configKey) = True Then
                Dim raw = CStr(workflowStep.Config(configKey))
                If Not String.IsNullOrWhiteSpace(raw) Then Return ReplaceVariables(raw)
            End If
            If _variableStore.ContainsKey(varKey) Then Return CStr(_variableStore(varKey))
            Return ""
        End Function

        ''' <summary>简单格式应用（大写/小写/金额/日期/数字千分位）</summary>
        Private Shared Function ApplyFormat(value As String, format As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Select Case format?.Trim().ToLowerInvariant()
                Case "upper", "uppercase" : Return value.ToUpperInvariant()
                Case "lower", "lowercase" : Return value.ToLowerInvariant()
                Case "trim" : Return value.Trim()
                Case "currency", "c", "c2"
                    Dim d As Decimal
                    If Decimal.TryParse(value, d) Then Return d.ToString("C2")
                Case "number", "n", "n2"
                    Dim d As Decimal
                    If Decimal.TryParse(value, d) Then Return d.ToString("N2")
                Case "date", "yyyy-mm-dd"
                    Dim dt As DateTime
                    If DateTime.TryParse(value, dt) Then Return dt.ToString("yyyy-MM-dd")
            End Select
            Return value
        End Function

    End Class

End Namespace