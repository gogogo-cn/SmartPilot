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

    ' SkillManager.Config.vb — 配置监控 / Skill 查询 API / 工作流持久化 / LLM 引擎初始化
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        Public Sub Dispose() Implements IDisposable.Dispose
            SaveAllPersistedWorkflows()

            ' 释放防抖计时器
            Try
                If _hotReloadDebounceTimer IsNot Nothing Then
                    _hotReloadDebounceTimer.Dispose()
                    _hotReloadDebounceTimer = Nothing
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("释放热重载计时器失败: " & ex.Message)
            End Try

            ' 取消订阅事件并释放配置监控器
            If _configWatcher IsNot Nothing Then
                RemoveHandler _configWatcher.ConfigFileChanged, AddressOf OnConfigFileChanged
                _configWatcher.Dispose()
            End If

            If _workflowEngine IsNot Nothing Then
                _workflowEngine.Dispose()
            End If

            If _llmEngine IsNot Nothing Then
                _llmEngine.Dispose()
            End If
        End Sub

        ''' <summary>
        ''' 启动配置监控
        ''' </summary>
        Private Sub StartConfigWatching()
            Try
                ' 确保配置监控器已初始化
                EnsureConfigWatcherInitialized()

                ' 监控 Markdown Skills 目录
                Dim mdDir = Utils.PathHelper.GetMarkdownSkillsDirectory()
                If Directory.Exists(mdDir) Then
                    _configWatcher.StartWatchingDirectory(Utils.ConfigType.Skills, mdDir, "*.md")
                    System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 已启动 Markdown Skills 目录监控: " & mdDir)
                End If

                ' 监控自定义 Skills 目录（用户自定义技能，AppData）
                Dim customSkillsDir = Path.Combine(Utils.PathHelper.CustomSkillsPath, "skills")
                If Directory.Exists(customSkillsDir) Then
                    _configWatcher.StartWatchingDirectory(Utils.ConfigType.CustomSkills, customSkillsDir, "*.md")
                    System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 已启动自定义 Skills 目录监控: " & customSkillsDir)
                End If

                ' 监控自定义代码片段目录（snippets）
                Dim customSnippetsDir = Path.Combine(Utils.PathHelper.CustomSkillsPath, "snippets")
                If Directory.Exists(customSnippetsDir) Then
                    _configWatcher.StartWatchingDirectory(Utils.ConfigType.CustomSkillSnippets, customSnippetsDir, "*.md")
                    System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 已启动自定义代码片段目录监控: " & customSnippetsDir)
                End If

                ' 监控数据类型配置文件
                Dim dataTypeConfigPath = GetDataTypeConfigFilePath()
                If File.Exists(dataTypeConfigPath) Then
                    _configWatcher.StartWatching(Utils.ConfigType.DataType, dataTypeConfigPath)
                End If

                System.Diagnostics.Debug.WriteLine("配置热更新监控已启动")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("启动配置监控失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 停止配置监控
        ''' </summary>
        Private Sub StopConfigWatching()
            Try
                _configWatcher.StopAllWatching()
                System.Diagnostics.Debug.WriteLine("配置热更新监控已停止")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("停止配置监控失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 配置文件变更处理
        ''' </summary>
        Private Sub OnConfigFileChanged(sender As Object, e As Utils.ConfigFileChangedEventArgs)
            Try
                System.Diagnostics.Debug.WriteLine(String.Format("检测到配置文件变更: {0} [{1}]", e.FilePath, e.ConfigType.ToString()))

                Select Case e.ConfigType
                    Case Utils.ConfigType.Skills, Utils.ConfigType.CustomSkills, Utils.ConfigType.CustomSkillSnippets
                        ' 【热更新】Skill 文件变更（系统/自定义技能/代码片段）→ 延迟防抖重载
                        ' 延迟 1 秒：合并短时间内的事件风暴，并等待写文件完成，避免读到半成品
                        ScheduleHotReload()

                    Case Utils.ConfigType.DataType
                        ' 重新加载数据类型配置
                        System.Diagnostics.Debug.WriteLine("正在热更新数据类型配置...")
                        If _workflowEngine IsNot Nothing Then
                            _workflowEngine.ReloadDataTypeConfigs()
                        End If

                    Case Utils.ConfigType.AppSettings
                        ' 应用程序配置变更，需要重启才能生效
                        System.Diagnostics.Debug.WriteLine("应用程序配置已变更，部分设置可能需要重启才能生效")

                End Select

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("处理配置文件变更失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 调度一次防抖热重载：1 秒内多次变更合并为一次，停止变化后才执行
        ''' </summary>
        Private Sub ScheduleHotReload()
            Try
                SyncLock _hotReloadSyncLock
                    If _hotReloadDebounceTimer Is Nothing Then
                        _hotReloadDebounceTimer = New System.Threading.Timer(
                            AddressOf OnHotReloadDebounceElapsed,
                            Nothing,
                            System.Threading.Timeout.Infinite,
                            System.Threading.Timeout.Infinite)
                    End If
                    ' trailing-edge：每次变更重置计时，停止变化 1 秒后执行一次
                    _hotReloadDebounceTimer.Change(1000, System.Threading.Timeout.Infinite)
                End SyncLock
                System.Diagnostics.Debug.WriteLine("已调度 Skill 热重载（1 秒防抖）")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("调度 Skill 热重载失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 防抖计时到点后执行热重载（Timer 回调线程，不触碰 UI）
        ''' </summary>
        Private Sub OnHotReloadDebounceElapsed(state As Object)
            Try
                System.Diagnostics.Debug.WriteLine("正在热更新 Skills（防抖后）...")
                ReloadSkills()
                InvalidateCache()
                System.Diagnostics.Debug.WriteLine("Skill 热重载完成")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine(String.Format("Skill 热重载失败: {0}", ex.Message))
            End Try
        End Sub

        ''' <summary>
        ''' 获取数据类型配置文件路径（供 WorkflowEngine 使用）
        ''' </summary>
        Private Function GetDataTypeConfigFilePath() As String
            Dim defaultPath = Path.Combine(Utils.PathHelper.ConfigDirectory, "DataTypeConfigs.json")
            Return defaultPath
        End Function

        ''' <summary>
        ''' 使所有缓存失效
        ''' </summary>
        Private Sub InvalidateCache()
            SyncLock _cacheLock
                _skillsListCacheValid = False
                _skillsListCache = Nothing
                _skillCategoriesCacheValid = False
                _skillCategoriesCache = Nothing
            End SyncLock
            System.Diagnostics.Debug.WriteLine("SkillManager 缓存已失效")
        End Sub



        ''' <summary>
        ''' 获取所有 Skills（使用缓存，含停用的）
        ''' </summary>
        Public Function GetAllSkills() As List(Of SkillDefinition)
            SyncLock _cacheLock
                If _skillsListCacheValid AndAlso _skillsListCache IsNot Nothing Then
                    Return _skillsListCache
                End If

                Using PerformanceMonitor.Instance.StartOperation("SkillManager.GetAllSkills")
                    _skillsListCache = _skills.Values.ToList()
                    _skillsListCacheValid = True
                    Return _skillsListCache
                End Using
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有启用的 Skills（IsEnabled=True）
        ''' 路由匹配、快捷入口、命令列表等场景使用此方法，
        ''' 排除用户在 Skill 管理界面停用的 Skill。
        ''' </summary>
        Public Function GetEnabledSkills() As List(Of SkillDefinition)
            Dim all = GetAllSkills()
            If all Is Nothing OrElse all.Count = 0 Then Return all
            Return all.Where(Function(s) s.IsEnabled).ToList()
        End Function

        ''' <summary>
        '''  将所有已加载的 Skill 注册到 McpCapabilityRegistry
        ''' 使 Skill 也能通过统一的后端匹配引擎被前端发现
        ''' 只注册启用的 Skill（IsEnabled=True）
        ''' </summary>
        Private Sub RegisterSkillsToCommandRegistry()
            Dim skills = GetEnabledSkills()
            If skills Is Nothing OrElse skills.Count = 0 Then Return

            Dim registry = Services.Mcp.McpCapabilityRegistry.Instance
            Dim registeredCount As Integer = 0

            For Each skill In skills
                If skill Is Nothing OrElse String.IsNullOrEmpty(skill.Name) Then Continue For

                Dim keywords As List(Of String) = Nothing
                If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Length > 0 Then
                    keywords = New List(Of String)(skill.Keywords)
                End If

                registry.RegisterSkill(
                    skillName:=skill.Name,
                    displayName:=If(String.IsNullOrEmpty(skill.DisplayName), skill.Name, skill.DisplayName),
                    description:=skill.Description,
                    category:=MapSkillCategoryToEnum(skill.Category),
                    keywords:=keywords,
                    isWriteOperation:=skill.ProcessingMode = SkillProcessingMode.VBACode OrElse
                                      skill.ProcessingMode = SkillProcessingMode.PythonCode
                )
                registeredCount += 1
            Next

            Utils.AppLogger.LogInfo("SkillManager", $"已将 {registeredCount} 个 Skill 注册到统一能力注册中心")
        End Sub

        ''' <summary>
        '''  将 Skill 字符串分类映射到 CapabilityCategory 枚举
        ''' </summary>
        Private Shared Function MapSkillCategoryToEnum(category As String) As Services.Mcp.CapabilityCategory
            If String.IsNullOrEmpty(category) Then Return Services.Mcp.CapabilityCategory.Utility

            Select Case category
                Case "数据分析", "数据分析工具"
                    Return Services.Mcp.CapabilityCategory.Analyze
                Case "数据处理"
                    Return Services.Mcp.CapabilityCategory.Transform
                Case "格式美化", "格式化"
                    Return Services.Mcp.CapabilityCategory.Format
                Case "效率工具"
                    Return Services.Mcp.CapabilityCategory.Utility
                Case "高级功能", "开发工具"
                    Return Services.Mcp.CapabilityCategory.Manage
                Case "Excel自动化工具"
                    Return Services.Mcp.CapabilityCategory.Write
                Case "导入导出"
                    Return Services.Mcp.CapabilityCategory.Export
                Case Else
                    Return Services.Mcp.CapabilityCategory.Utility
            End Select
        End Function

        ''' <summary>
        ''' 按类别获取 Skills（使用缓存）
        ''' </summary>
        Public Function GetSkillsByCategory(category As String) As List(Of SkillDefinition)
            SyncLock _cacheLock
                ' 确保类别缓存已初始化
                If Not _skillCategoriesCacheValid OrElse _skillCategoriesCache Is Nothing Then
                    BuildCategoryCache()
                End If

                If _skillCategoriesCache.ContainsKey(category) Then
                    Return _skillCategoriesCache(category)
                End If

                Return New List(Of SkillDefinition)()
            End SyncLock
        End Function

        ''' <summary>
        ''' 构建类别缓存
        ''' </summary>
        Private Sub BuildCategoryCache()
            _skillCategoriesCache = New Dictionary(Of String, List(Of SkillDefinition))()

            For Each skill In _skills.Values
                Dim category = If(String.IsNullOrEmpty(skill.Category), "其他", skill.Category)

                If Not _skillCategoriesCache.ContainsKey(category) Then
                    _skillCategoriesCache(category) = New List(Of SkillDefinition)()
                End If

                _skillCategoriesCache(category).Add(skill)
            Next

            _skillCategoriesCacheValid = True
        End Sub



        Private Sub LoadPersistedWorkflows()
            Try
                If Not File.Exists(_persistenceFilePath) Then Return

                Dim json = File.ReadAllText(_persistenceFilePath)
                Dim states = JsonConvert.DeserializeObject(Of List(Of WorkflowPersistenceState))(json)

                If states IsNot Nothing Then
                    For Each state In states
                        If state.Status = WorkflowPersistenceState.WorkflowStatus.InProgress Then
                            _activeWorkflows(state.WorkflowId) = state
                        End If
                    Next
                    System.Diagnostics.Debug.WriteLine($"Loaded {_activeWorkflows.Count} persisted workflows")
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("加载工作流状态失败: " & ex.Message)
            End Try
        End Sub

        Private Sub SaveAllPersistedWorkflows()
            SyncLock _persistenceLock
                Try
                    Dim dir = Path.GetDirectoryName(_persistenceFilePath)
                    If Not Directory.Exists(dir) Then
                        Directory.CreateDirectory(dir)
                    End If

                    Dim states = _activeWorkflows.Values.ToList()
                    Dim json = JsonConvert.SerializeObject(states, Formatting.Indented)
                    File.WriteAllText(_persistenceFilePath, json)
                    System.Diagnostics.Debug.WriteLine($"Saved {_activeWorkflows.Count} workflows to persistence")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("保存工作流状态失败: " + ex.Message)
                End Try
            End SyncLock
        End Sub

        Public Sub SaveWorkflowState(workflowId As String, skillName As String, currentStepId As String,
                                     variableStore As Dictionary(Of String, Object),
                                     executionLog As List(Of String),
                                     stepExecutionCount As Dictionary(Of String, Integer))
            SyncLock _persistenceLock
                Dim state As WorkflowPersistenceState

                If _activeWorkflows.ContainsKey(workflowId) Then
                    state = _activeWorkflows(workflowId)
                Else
                    state = New WorkflowPersistenceState With {
                        .WorkflowId = workflowId,
                        .SkillName = skillName
                    }
                    _activeWorkflows(workflowId) = state
                End If

                state.CurrentStepId = currentStepId
                state.VariableStore = variableStore
                state.ExecutionLog = executionLog
                state.StepExecutionCount = stepExecutionCount
                state.LastCheckpointAt = DateTime.Now

                SaveAllPersistedWorkflows()
            End SyncLock
        End Sub

        Public Function GetPersistedWorkflow(workflowId As String) As WorkflowPersistenceState
            SyncLock _persistenceLock
                If _activeWorkflows.ContainsKey(workflowId) Then
                    Return _activeWorkflows(workflowId)
                End If
                Return Nothing
            End SyncLock
        End Function

        Public Function GetAllPersistedWorkflows() As List(Of WorkflowPersistenceState)
            SyncLock _persistenceLock
                Return _activeWorkflows.Values.ToList()
            End SyncLock
        End Function

        Public Sub RemovePersistedWorkflow(workflowId As String)
            SyncLock _persistenceLock
                If _activeWorkflows.ContainsKey(workflowId) Then
                    _activeWorkflows.Remove(workflowId)
                    SaveAllPersistedWorkflows()
                End If
            End SyncLock
        End Sub

        Public Sub UpdateWorkflowStatus(workflowId As String, status As WorkflowPersistenceState.WorkflowStatus)
            SyncLock _persistenceLock
                If _activeWorkflows.ContainsKey(workflowId) Then
                    _activeWorkflows(workflowId).Status = status
                    SaveAllPersistedWorkflows()
                End If
            End SyncLock
        End Sub



        Private Sub InitializeLLMEngine()
            Try
                Dim configManager = Utils.ConfigurationManager.Instance
                Dim provider = configManager.GetSetting(Utils.ConfigConstants.LLM.PROVIDER, Utils.ConfigConstants.LLM.DEFAULT_PROVIDER)

                ' 根据提供商获取正确的 API URL 和模型
                Dim apiUrl As String = ""
                Dim model As String = ""
                If provider.ToLower() = "ollama" Then
                    apiUrl = configManager.GetSetting(Utils.ConfigConstants.Ollama.API_URL, Utils.ConfigConstants.Ollama.DEFAULT_API_URL)
                    model = configManager.GetSetting(Utils.ConfigConstants.LLM.MODEL, Utils.ConfigConstants.Ollama.DEFAULT_MODEL)
                Else
                    apiUrl = configManager.GetSetting(Utils.ConfigConstants.LLM.API_URL, "")
                    model = configManager.GetSetting(Utils.ConfigConstants.LLM.MODEL, Utils.ConfigConstants.LLM.DEFAULT_MODEL)
                End If

                Dim config As New LLMRequestConfig With {
                    .Provider = provider,
                    .Model = model,
                    .ApiKey = configManager.GetSetting(Utils.ConfigConstants.LLM.API_KEY, ""),
                    .ApiUrl = apiUrl,
                    .Temperature = Double.Parse(configManager.GetSetting(Utils.ConfigConstants.LLM.TEMPERATURE, Utils.ConfigConstants.LLM.DEFAULT_TEMPERATURE.ToString())),
                    .MaxTokens = Integer.Parse(configManager.GetSetting(Utils.ConfigConstants.LLM.MAX_TOKENS, Utils.ConfigConstants.LLM.DEFAULT_MAX_TOKENS.ToString()))
                }
                _llmEngine.Initialize(config)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("初始化 LLM Engine 失败: " & ex.Message)
            End Try
        End Sub

    End Class

End Namespace
