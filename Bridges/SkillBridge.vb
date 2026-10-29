Option Explicit On
Option Strict Off
Option Infer On

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading.Tasks
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Skills
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Services

Namespace Bridges

    ''' <summary>
    ''' 自定义 CamelCase 命名策略 — 正确处理缩写词（如 LLM → llm, UI → ui）
    ''' </summary>
    Public Class AcronymAwareNamingStrategy
        Inherits CamelCaseNamingStrategy

        Public Overrides Function GetPropertyName(name As String, overrideSpecified As Boolean) As String
            If String.IsNullOrEmpty(name) Then Return name

            ' 查找第一个小写字母的位置
            Dim firstLower As Integer = -1
            For i As Integer = 0 To name.Length - 1
                If Char.IsLower(name(i)) Then
                    firstLower = i
                    Exit For
                End If
            Next

            ' 正常 PascalCase：首字母大写后紧跟小写（如 Output, Processing）
            ' 直接小写首字母即可
            If firstLower <= 1 Then
                Return Char.ToLowerInvariant(name(0)) & name.Substring(1)
            End If

            ' 缩写词场景：多个大写字母开头后紧跟小写（如 LLMSystem, UIActions）
            ' 将缩写词部分全部小写，保留后续单词的首字母大写
            ' 例如：LLMSystemPrompt → llmSystemPrompt, UIActions → uiActions
            Dim acronym = name.Substring(0, firstLower - 1).ToLowerInvariant()
            Dim rest = name.Substring(firstLower - 1)
            Return acronym & rest
        End Function
    End Class

    ''' <summary>
    ''' SkillBridge - Skill 系统的桥梁类
    ''' 从 ExcelBridge 中提取的 Skill 相关 API 方法
    ''' 提供 COM 可见的方法供前端 JavaScript 调用
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    <ProgId("ExcelAIVSTOAddIn.SkillBridge")>
    Public Class SkillBridge
        Implements IBridge, ISkillBridge
        ' 注意: 此类实现 ISkillBridge 接口契约（见 IBridges.vb）。
        ' 契约方法以同名公共方法显式 Implements 绑定，供依赖注入与单元测试按接口解析。

        ' 单例实例 - 使用 Lazy(Of T) 确保线程安全
        Private Shared ReadOnly _lazyInstance As New Lazy(Of SkillBridge)(Function() New SkillBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        ' Skill 系统组件
        Private _skillManager As SkillManager
        Private _skillRecognizer As SkillRecognizer
        Private _customSkillManager As CustomSkillManager
        Private _customSkillExecutor As CustomSkillExecutor
        Private _skillsInitialized As Boolean = False
        Private _initializationLock As New Object()

        ' 意图识别引擎和规则管理器
        Private _intentEngine As IntentRecognitionEngine
        Private _rulesManager As RulesManager

        ' 调试日志：同时写 VS Output 和 AppData 文件
        Private Shared ReadOnly _dbgLogPath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAIVSTOAddIn", "skill-toggle-debug.log")

        Private Shared Sub DbgLog(msg As String)
            Dim line = $"[{DateTime.Now:HH:mm:ss.fff}] {msg}"
            Debug.WriteLine(line)
            Try
                Dim dir = Path.GetDirectoryName(_dbgLogPath)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                File.AppendAllText(_dbgLogPath, line & vbCrLf)
            Catch
            End Try
        End Sub

#Region "单例模式"

        ''' <summary>
        ''' 获取 SkillBridge 单例实例（线程安全 Lazy 初始化）
        ''' </summary>
        Public Shared ReadOnly Property Instance As SkillBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        ''' <summary>
        ''' 私有构造函数（单例模式）
        ''' </summary>
        Private Sub New()
            Try
                System.Diagnostics.Debug.WriteLine("========== SkillBridge 构造函数开始 ==========")
                ' Skill 系统采用延迟初始化策略，不在构造函数中加载
                System.Diagnostics.Debug.WriteLine("Skill 系统将采用延迟初始化")
                System.Diagnostics.Debug.WriteLine("========== SkillBridge 构造函数完成 ==========")
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("SkillBridge 构造函数异常: " & ex.Message)
                Throw
            End Try
        End Sub

#End Region

        ''' <summary>
        ''' 在 Excel 主线程（UI 线程）上执行有返回值的函数。
        ''' Skill 执行过程中若访问 Excel COM 对象，应通过此 Helper 切换到 UI 线程。
        ''' </summary>
        Private Shared Function ExecuteOnExcelUIThread(Of T)(func As Func(Of T)) As T
            Dim result As T = Nothing
            ThisAddIn.RunOnUIThread(Sub() result = func())
            Return result
        End Function

        ''' <summary>
        ''' 获取带有缩写词感知的 camelCase 序列化设置
        ''' </summary>
        Private Shared ReadOnly _camelSettings As JsonSerializerSettings = New JsonSerializerSettings With {
            .ContractResolver = New DefaultContractResolver With {
                .NamingStrategy = New AcronymAwareNamingStrategy()
            },
            .Formatting = Formatting.Indented
        }

        Private Shared Function GetCamelSettings() As JsonSerializerSettings
            Return _camelSettings
        End Function

#Region "初始化"

        ''' <summary>
        ''' 确保 Skill 系统已初始化（延迟加载）
        ''' </summary>
        Private Sub EnsureSkillsInitialized()
            If _skillsInitialized Then Return

            SyncLock _initializationLock
                If _skillsInitialized Then Return

                Try
                    System.Diagnostics.Debug.WriteLine("[Skill] 开始延迟初始化 Skill 系统...")

                    ' 初始化 SkillManager（单例模式）
                    _skillManager = SkillManager.Instance
                    If _skillManager Is Nothing Then
                        System.Diagnostics.Debug.WriteLine("[Skill] SkillManager.Instance 返回 Nothing")
                        Return
                    End If
                    System.Diagnostics.Debug.WriteLine("[Skill] SkillManager 实例获取成功")

                    ' 加载 Skills
                    _skillManager.LoadSkills()
                    System.Diagnostics.Debug.WriteLine("[Skill] Skills 加载完成")

                    ' 初始化其他组件
                    _skillRecognizer = New SkillRecognizer(_skillManager)
                    _customSkillManager = CustomSkillManager.GetInstance()
                    
                    ' 调试输出:CustomSkillManager 加载情况
                    Dim customSkills = _customSkillManager.GetAllCustomSkills()
                    System.Diagnostics.Debug.WriteLine($"[Skill] CustomSkillManager 路径: {Utils.PathHelper.CustomSkillsPath}")
                    System.Diagnostics.Debug.WriteLine($"[Skill] CustomSkillManager 加载数量: {customSkills.Count}")
                    For Each s In customSkills
                        System.Diagnostics.Debug.WriteLine($"[Skill]   - {s.Name}")
                    Next
                    
                    _customSkillExecutor = New CustomSkillExecutor()

                    ' 初始化意图识别引擎和规则管理器
                    _rulesManager = RulesManager.Instance
                    _intentEngine = New IntentRecognitionEngine()
                    System.Diagnostics.Debug.WriteLine("[Skill] 意图识别引擎初始化完成")

                    _skillsInitialized = True
                    System.Diagnostics.Debug.WriteLine("[Skill] Skill 系统延迟初始化成功")
                Catch ex As Exception
                    System.Diagnostics.Debug.WriteLine("[Skill] Skill 系统初始化失败: " & ex.Message)

                    ' 确保在失败时设置为 Nothing，以便调用方知道初始化失败
                    _skillManager = Nothing
                    _skillRecognizer = Nothing
                    _customSkillManager = Nothing
                    _customSkillExecutor = Nothing
                    Throw
                End Try
            End SyncLock
        End Sub

#End Region

#Region "Skill API"

        ''' <summary>
        ''' 获取所有可用的 Skill 列表（仅启用的，IsEnabled=True）
        ''' 路由匹配、快捷入口、命令列表等场景调用此方法。
        ''' Skill 管理界面需要完整列表（含停用的）请用 GetAllSkills。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAvailableSkills() As String Implements ISkillBridge.GetAvailableSkills
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "SkillManager 未初始化"})
                End If

                Dim skills = _skillManager.GetEnabledSkills()
                Dim customCount = skills.Where(Function(s) s.IsCustom).Count()
                Dim disabledCount = _skillManager.GetAllSkills().Count - skills.Count
                System.Diagnostics.Debug.WriteLine($"[GetAvailableSkills] 启用={skills.Count}, 自定义={customCount}, 停用={disabledCount}")

                ' 使用缩写词感知的 camelCase 命名策略序列化
                Return JsonConvert.SerializeObject(New With {.success = True, .skills = skills}, GetCamelSettings())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取识别类型列表（用于前端卡片类型配置）
        ''' 将 Skill 定义转换为前端 IMAGE_TYPE_CONFIG 期望的格式
        ''' </summary>
        <ComVisible(True)>
        Public Function GetRecognitionTypes() As String
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skills = _skillManager.GetAllSkills()
                Dim recognitionTypes = New List(Of Object)()

                For Each skill In skills
                    Dim displayName As String = If(String.IsNullOrWhiteSpace(skill.DisplayName), skill.Name, skill.DisplayName)
                    Dim defaultWorksheet As String = displayName & "信息"
                    Dim fieldMappingDict As New Dictionary(Of String, Object)()
                    Dim fieldNameMappingDict As New Dictionary(Of String, String)()

                    If skill.ExcelConfig IsNot Nothing Then
                        If Not String.IsNullOrWhiteSpace(skill.ExcelConfig.DefaultWorksheet) Then
                            defaultWorksheet = skill.ExcelConfig.DefaultWorksheet
                        End If

                        If skill.ExcelConfig.FieldMapping IsNot Nothing Then
                            For Each kv In skill.ExcelConfig.FieldMapping
                                fieldMappingDict(kv.Key) = New With {
                                    .column = kv.Value.Column,
                                    .header = kv.Value.Header
                                }
                            Next
                        End If

                        If skill.ExcelConfig.FieldNameMapping IsNot Nothing Then
                            For Each kv In skill.ExcelConfig.FieldNameMapping
                                fieldNameMappingDict(kv.Key) = kv.Value
                            Next
                        End If
                    End If

                    Dim extractFields = New List(Of String)()
                    If skill.FieldMappings IsNot Nothing Then
                        For Each fm In skill.FieldMappings
                            If Not String.IsNullOrWhiteSpace(fm.SourceField) Then
                                extractFields.Add(fm.SourceField)
                            End If
                        Next
                    End If

                    recognitionTypes.Add(New With {
                        .id = skill.Name,
                        .name = skill.Name,
                        .displayName = displayName,
                        .keywords = If(skill.Keywords, New String() {}),
                        .icon = "",
                        .extractFields = extractFields,
                        .excelConfig = New With {
                            .defaultWorksheet = defaultWorksheet,
                            .fieldMapping = fieldMappingDict,
                            .fieldNameMapping = fieldNameMappingDict
                        }
                    })
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .recognitionTypes = recognitionTypes
                }, GetCamelSettings())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取指定 Skill 的详细信息
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSkillDetails(skillName As String) As String Implements ISkillBridge.GetSkillDetails
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skill = _skillManager.GetSkill(skillName)
                If skill Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = $"Skill '{skillName}' 不存在"})
                End If

                ' 使用缩写词感知的 camelCase 命名策略序列化
                Dim skillJson = JsonConvert.SerializeObject(skill, GetCamelSettings())
                
                ' 包装成 {success: true, skill: {...}} 格式
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .skill = JsonConvert.DeserializeObject(Of Object)(skillJson)
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 通过名称获取 Skill（前端调用此方法）
        ''' </summary>
        <ComVisible(True)>
        Public Function GetSkillByName(skillName As String) As String Implements ISkillBridge.GetSkillByName
            ' 直接调用 GetSkillDetails 实现，保持代码复用
            Return GetSkillDetails(skillName)
        End Function

        ''' <summary>
        ''' 打开 Skill 的 Markdown 源文件
        ''' </summary>
        ''' <param name="skillName">Skill 名称</param>
        ''' <returns>JSON 格式的操作结果</returns>
        <ComVisible(True)>
        Public Function OpenSkillSourceFile(skillName As String) As String
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                ' 先检查自定义 Skill 目录
                Dim filePath = System.IO.Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAIVSTOAddIn", "CustomSkills", "skills", skillName & ".md")
                Dim isCustom = True

                If Not System.IO.File.Exists(filePath) Then
                    ' 再检查系统 Skill 目录
                    filePath = System.IO.Path.Combine(
                        AppDomain.CurrentDomain.BaseDirectory, "Skills", skillName & ".md")
                    isCustom = False
                End If

                If Not System.IO.File.Exists(filePath) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = $"Skill '{skillName}' 的源文件不存在"
                    })
                End If

                ' 使用默认程序打开文件
                System.Diagnostics.Process.Start(New System.Diagnostics.ProcessStartInfo(filePath) With {
                    .UseShellExecute = True
                })

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .path = filePath,
                    .isCustom = isCustom
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取所有标记为快捷入口的 Skill 列表
        ''' </summary>
        <ComVisible(True)>
        Public Function GetQuickShortcutSkills() As String Implements ISkillBridge.GetQuickShortcutSkills
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim allSkills = _skillManager.GetAllSkills()
                Dim quickShortcuts = allSkills.Where(Function(s) s.IsQuickShortcut).ToList()

                Return JsonConvert.SerializeObject(New With {.success = True, .skills = quickShortcuts}, GetCamelSettings())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 切换 Skill 的快捷入口标记状态
        ''' </summary>
        ''' <param name="skillName">Skill 名称</param>
        ''' <param name="isQuick">是否设为快捷入口</param>
        <ComVisible(True)>
        Public Function ToggleQuickShortcut(skillName As String, isQuick As Boolean) As String Implements ISkillBridge.ToggleQuickShortcut
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skill = _skillManager.GetSkill(skillName)
                If skill Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = $"Skill '{skillName}' 不存在"})
                End If

                skill.IsQuickShortcut = isQuick

                ' 【修复】所有技能的快捷入口状态都保存到 CustomSkillManager
                ' 内置技能也会被"提升"为自定义版本保存到 AppData，确保构建后不丢失
                Try
                    Dim customManager = Skills.CustomSkillManager.GetInstance()
                    skill.IsCustom = True
                    customManager.SaveCustomSkill(skill)
                    System.Diagnostics.Debug.WriteLine("[ToggleQuickShortcut] 已保存到 CustomSkillManager: " & skillName)
                Catch customEx As Exception
                    System.Diagnostics.Debug.WriteLine("[ToggleQuickShortcut] 保存到 CustomSkillManager 失败: " & customEx.Message)
                End Try

                ' 保存系统技能到 bin 目录（仅内置技能，自定义技能已跳过）
                _skillManager.SaveSkills()

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .skillName = skillName,
                    .isQuickShortcut = isQuick
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 执行 Skill（同步版本，供COM接口调用）
        ''' 【修复】使用 Task.Run + 线程池避免 STA 线程上 .GetAwaiter().GetResult() 死锁风险
        ''' 【注意】Task.Run 会将执行转移到线程池线程；若 Skill 内部访问 Excel COM 对象，
        '''        必须在该处使用 ExecuteOnExcelUIThread 切换回 Excel UI 线程，否则会出现
        '''        RPC_E_WRONG_THREAD 或数据损坏。
        ''' </summary>
        <ComVisible(True)>
        Public Function ExecuteSkill(skillName As String, inputsJson As String) As String Implements ISkillBridge.ExecuteSkill
            Try
                System.Diagnostics.Debug.WriteLine("[ExecuteSkill] 开始执行: " & skillName)

                ' 【修复】使用 Task.Run 在线程池线程上执行异步操作，避免在 STA COM 线程上死锁
                ' TODO: SkillManager.ExecuteSkill 内部若访问 Excel COM，应通过 ExecuteOnExcelUIThread 切换到 UI 线程
#Disable Warning BC42025 ' Task.Run在实例方法中调用时的已知VB.NET误报
                Dim task As Task(Of String) = Task.Run(Async Function()
                    Return Await ExecuteSkillAsync(skillName, inputsJson)
                End Function)
#Enable Warning BC42025
                Dim result = task.GetAwaiter().GetResult() ' 线程池线程上无 SynchronizationContext，不会死锁

                System.Diagnostics.Debug.WriteLine("[ExecuteSkill] 执行完成: " & result)
                Return result
            Catch aggEx As AggregateException
                ' 解包AggregateException
                Dim innerEx = aggEx.InnerException
                System.Diagnostics.Debug.WriteLine("[ExecuteSkill] AggregateException: " & innerEx.Message)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = innerEx.Message})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[ExecuteSkill] 异常: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 执行 Skill（异步版本）
        ''' </summary>
        <ComVisible(True)>
        Public Async Function ExecuteSkillAsync(skillName As String, inputsJson As String) As Task(Of String) Implements ISkillBridge.ExecuteSkillAsync
            Try
                EnsureSkillsInitialized()

                ' 【动态加载兜底】如果内存中没有，尝试从 Markdown 文件即时加载
                Dim skill = _skillManager.GetSkill(skillName)
                If skill Is Nothing Then
                    System.Diagnostics.Debug.WriteLine($"[ExecuteSkillAsync] Skill '{skillName}' 不在内存中，尝试动态加载...")
                    Try
                        _skillManager.ReloadSingleSkill(skillName)
                        skill = _skillManager.GetSkill(skillName)
                        If skill IsNot Nothing Then
                            System.Diagnostics.Debug.WriteLine($"[ExecuteSkillAsync] ✅ 动态加载成功: {skillName}")
                        End If
                    Catch reloadEx As Exception
                        System.Diagnostics.Debug.WriteLine($"[ExecuteSkillAsync] 动态加载失败: {reloadEx.Message}")
                    End Try
                End If

                If skill Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = $"Skill '{skillName}' 不存在"})
                End If

                Dim inputs = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(inputsJson)

                ' 通过统一路由执行 Skill
                Dim router = Skills.SkillExecutionRouter.Instance
                Dim result = Await router.ExecuteSkill(skillName, inputs)

                ' 【修复 代码溢出】出口端清洗 Message/Outputs("content")，防止 LLM 把内部 JSON / CSS 字典尾巴直接抄进总结回复
                Dim cleanMessage = SanitizeLeakedJsonSnippets(If(result.Message, ""))
                Dim cleanOutputs As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
                If result.Outputs IsNot Nothing Then
                    For Each kv In result.Outputs
                        If kv.Key Is Nothing Then Continue For
                        Dim vObj = kv.Value
                        If TypeOf vObj Is String Then
                            cleanOutputs(kv.Key) = SanitizeLeakedJsonSnippets(CStr(vObj))
                        Else
                            cleanOutputs(kv.Key) = vObj
                        End If
                    Next
                End If

                ' 【步骤规划兜底】若整棵 outputs 尚未包含任何 steps（美化表格/数据清洗/图表等暂未产出精细步骤的 Skill），
                ' 则用真实执行日志生成简略规划并挂到 outputs.steps，保证所有操作型 Skill 都能统一显示"步骤规划"样式。
                If Not TreeContainsSteps(cleanOutputs) Then
                    cleanOutputs("steps") = BuildFallbackSteps(result.ExecutionLog)
                End If

                Return JsonConvert.SerializeObject(New With {
                    .success = result.Success,
                    .message = cleanMessage,
                    .outputs = cleanOutputs,
                    .error = result.ErrorMessage
                })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 递归判断传入对象是否已包含 steps 字段（用于决定是否需要兜底生成步骤规划）。
        ''' 兼容对象、字典、数组以及内嵌 JSON 字符串（合并/拆分/关联的 steps 藏在 outputs.executionResult.Message 里）。
        ''' </summary>
        Private Function TreeContainsSteps(value As Object) As Boolean
            If value Is Nothing Then Return False
            Dim s = TryCast(value, String)
            If s IsNot Nothing Then
                If s.Contains("""steps""") Then Return True
                Dim trimmed = s.Trim()
                If trimmed.StartsWith("{") OrElse trimmed.StartsWith("[") Then
                    Try
                        Return TreeContainsSteps(Newtonsoft.Json.Linq.JToken.Parse(s))
                    Catch ex As Exception  ' 
                        Return False
                    End Try
                End If
                Return False
            End If
            Dim token = TryCast(value, Newtonsoft.Json.Linq.JToken)
            If token IsNot Nothing Then
                Dim obj = TryCast(token, Newtonsoft.Json.Linq.JObject)
                If obj IsNot Nothing Then
                    If obj("steps") IsNot Nothing Then Return True
                    For Each p In obj.Properties()
                        If TreeContainsSteps(p.Value) Then Return True
                    Next
                    Return False
                End If
                Dim arr = TryCast(token, Newtonsoft.Json.Linq.JArray)
                If arr IsNot Nothing Then
                    For Each item In arr
                        If TreeContainsSteps(item) Then Return True
                    Next
                End If
                Return False
            End If
            Dim dict = TryCast(value, System.Collections.IDictionary)
            If dict IsNot Nothing Then
                For Each key In dict.Keys
                    If String.Equals(Convert.ToString(key), "steps", StringComparison.OrdinalIgnoreCase) Then Return True
                    If TreeContainsSteps(dict(key)) Then Return True
                Next
            End If
            Return False
        End Function

        ''' <summary>
        ''' 用真实执行日志生成简略步骤规划。优先将已知的关键里程碑映射为友好友文案，
        ''' 其余记录行剥离日志前缀后原样保留——内容完全来自真实执行，不编造。
        ''' </summary>
        Private Function BuildFallbackSteps(log As List(Of String)) As List(Of String)
            Dim steps As New List(Of String)()
            If log Is Nothing Then Return steps
            For Each line In log
                If line Is Nothing Then Continue For
                Dim t = line.Trim()
                If String.IsNullOrWhiteSpace(t) Then Continue For
                Dim friendly = MapLogLineToStep(t)
                If String.IsNullOrWhiteSpace(friendly) Then Continue For
                If steps.Count > 0 AndAlso String.Equals(steps(steps.Count - 1), friendly, StringComparison.Ordinal) Then Continue For
                steps.Add(friendly)
                If steps.Count >= 8 Then Exit For
            Next
            If steps.Count = 0 Then steps.Add("执行完成")
            Return steps
        End Function

        Private Function MapLogLineToStep(line As String) As String
            Dim lower = line.ToLower()
            If lower.Contains("phase 1") AndAlso lower.Contains("分析") Then Return "解析请求并分析数据"
            If lower.Contains("phase 2") AndAlso lower.Contains("开始") Then Return "开始执行操作"
            If lower.Contains("使用 mcp") Then Return "调用工具执行"
            If lower.Contains("phase 2") AndAlso lower.Contains("执行完成") Then Return "操作执行完成"
            If lower.Contains("开始执行") Then Return "开始执行操作"
            If lower.Contains("执行完成") Then Return "操作执行完成"
            Dim idx = line.IndexOf("]")
            If idx >= 0 AndAlso idx < line.Length - 1 Then
                Dim rest = line.Substring(idx + 1).Trim()
                If rest.Length > 0 Then Return rest
            End If
            Return line.Trim()
        End Function

        ''' <summary>
        ''' 获取所有 Skill（别名方法，含停用的）
        ''' Skill 管理界面专用，返回完整列表。
        ''' 路由匹配等场景请用 GetAvailableSkills（仅启用）。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAllSkills() As String Implements ISkillBridge.GetAllSkills
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skills = _skillManager.GetAllSkills()
                Dim enabledCount = skills.Where(Function(s) s.IsEnabled).Count()
                Dim disabledCount = skills.Count - enabledCount
                System.Diagnostics.Debug.WriteLine($"[GetAllSkills] 总数={skills.Count}, 启用={enabledCount}, 停用={disabledCount}")

                Return JsonConvert.SerializeObject(New With {.success = True, .skills = skills}, GetCamelSettings())
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 切换 Skill 的启用/停用状态
        ''' 停用的 Skill 不参与路由匹配、快捷入口。
        ''' 状态持久化到 CustomSkillManager（AppData）。
        ''' </summary>
        ''' <param name="skillName">Skill 名称</param>
        ''' <param name="isEnabled">True=启用, False=停用</param>
        <ComVisible(True)>
        Public Function ToggleSkillEnabled(skillName As String, isEnabled As Boolean) As String Implements ISkillBridge.ToggleSkillEnabled
            Try
                DbgLog($"[ToggleSkillEnabled] ⭐⭐⭐ ENTRY: skillName='{skillName}', isEnabled={isEnabled}, SkillManager IsNothing={_skillManager Is Nothing}")
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    DbgLog("[ToggleSkillEnabled] SkillManager 未初始化，返回失败")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skill = _skillManager.GetSkill(skillName)
                If skill Is Nothing Then
                    DbgLog($"[ToggleSkillEnabled] Skill '{skillName}' 不存在")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = $"Skill '{skillName}' 不存在"})
                End If

                Dim oldState = skill.IsEnabled
                DbgLog($"[ToggleSkillEnabled] 找到 skill: IsEnabled={oldState}, IsCustom={skill.IsCustom}, Version={skill.Version}")
                skill.IsEnabled = isEnabled
                DbgLog($"[ToggleSkillEnabled] 内存已修改: {oldState} -> {isEnabled}")

                ' 固化到独立停用名单（与 .md 合并解耦，热重载/重读后依然生效）
                Try
                    _skillManager.SetSkillDisabled(skillName, isEnabled)
                Catch dx As Exception
                    DbgLog($"[ToggleSkillEnabled] SetSkillDisabled 异常: {dx.Message}")
                End Try

                ' 持久化到 CustomSkillManager（类似 IsQuickShortcut 的处理方式）
                Try
                    Dim customManager = Skills.CustomSkillManager.GetInstance()
                    DbgLog($"[ToggleSkillEnabled] CustomSkillManager IsNothing={customManager Is Nothing}")
                    skill.IsCustom = True
                    DbgLog($"[ToggleSkillEnabled] 准备调 SaveCustomSkill: name={skill.Name}, IsEnabled={skill.IsEnabled}, IsCustom={skill.IsCustom}")
                    customManager.SaveCustomSkill(skill)
                    DbgLog($"[ToggleSkillEnabled] ✅ SaveCustomSkill 完成: {skillName}: {oldState} -> {isEnabled}")
                Catch customEx As Exception
                    DbgLog($"[ToggleSkillEnabled] ❌ CustomSkillManager 保存异常: {customEx.Message}")
                    DbgLog($"[ToggleSkillEnabled] StackTrace: {customEx.StackTrace}")
                End Try

                ' 保存系统技能到 bin 目录（仅内置技能，自定义技能已跳过）
                DbgLog("[ToggleSkillEnabled] 准备调 SaveSkills()")
                _skillManager.SaveSkills()
                DbgLog("[ToggleSkillEnabled] ✅ SaveSkills() 完成")

                ' 清除意图识别缓存
                Try
                    If _intentEngine IsNot Nothing Then
                        _intentEngine.ClearCache()
                    End If
                Catch
                End Try

                DbgLog($"[ToggleSkillEnabled] ⭐⭐⭐ EXIT success: {skillName} now IsEnabled={isEnabled}")
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .skillName = skillName,
                    .isEnabled = isEnabled
                })
            Catch ex As Exception
                DbgLog($"[ToggleSkillEnabled] ❌ TOP-LEVEL EXCEPTION: {ex.Message}")
                DbgLog($"[ToggleSkillEnabled] StackTrace: {ex.StackTrace}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取所有代码片段（从 CustomSkillManager 的 snippets 目录直接返回，不经过 SkillManager 合并流程）
        ''' 前端代码库弹窗专用，避免因 SkillManager 过滤/版本保护导致代码片段丢失
        ''' </summary>
        <ComVisible(True)>
        Public Function GetAllCodeSnippets() As String
            Try
                EnsureSkillsInitialized()
                Skills.CustomSkillManager.ResetInstance()
                Dim customManager = Skills.CustomSkillManager.GetInstance()
                Dim snippets = customManager.GetAllCodeSnippets()
                System.Diagnostics.Debug.WriteLine($"[GetAllCodeSnippets] 从 CustomSkillManager 读取到 {snippets.Count} 个代码片段")
                For Each s In snippets
                    System.Diagnostics.Debug.WriteLine($"[GetAllCodeSnippets]   - {s.Name} (DisplayName={s.DisplayName}, HasCode={Not String.IsNullOrEmpty(s.TemplateCode)})")
                Next
                Return JsonConvert.SerializeObject(New With {.success = True, .snippets = snippets}, GetCamelSettings())
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[GetAllCodeSnippets] 异常: {ex.Message}")
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 重新加载所有 Skills（从 Markdown 文件强制刷新）
        ''' </summary>
        <ComVisible(True)>
        Public Function ReloadAllSkills() As String
            Try
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim result = _skillManager.ReloadSkills()
                If result Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "Skills 重新加载成功", .count = _skillManager.GetAllSkills().Count})
                Else
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Skills 重新加载失败"})
                End If
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 获取自定义 Skill 调试信息
        ''' </summary>
        <ComVisible(True)>
        Public Function GetCustomSkillDebugInfo() As String
            Try
                EnsureSkillsInitialized()
                If _customSkillManager Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.error = "CustomSkillManager 未初始化"})
                End If

                Dim customSkills = _customSkillManager.GetAllCustomSkills()
                Dim customSkillsPath = System.IO.Path.Combine(Utils.PathHelper.CustomSkillsPath, "skills")
                Dim snippetsPath = System.IO.Path.Combine(Utils.PathHelper.CustomSkillsPath, "snippets")
                
                Dim skillFiles As New List(Of String)
                Try
                    If System.IO.Directory.Exists(customSkillsPath) Then
                        skillFiles = System.IO.Directory.GetFiles(customSkillsPath, "*.md").Select(Function(f) System.IO.Path.GetFileName(f)).ToList()
                    End If
                Catch ex As Exception
                    skillFiles.Add("ERROR: " & ex.Message)
                End Try
                
                Dim snippetFiles As New List(Of String)
                Try
                    If System.IO.Directory.Exists(snippetsPath) Then
                        snippetFiles = System.IO.Directory.GetFiles(snippetsPath, "*.md").Select(Function(f) System.IO.Path.GetFileName(f)).ToList()
                    End If
                Catch ex As Exception
                    snippetFiles.Add("ERROR: " & ex.Message)
                End Try
                
                Dim allSkillsCount = 0
                Dim allSkillsCustomCount = 0
                If _skillManager IsNot Nothing Then
                    allSkillsCount = _skillManager.GetAllSkills().Count
                    allSkillsCustomCount = _skillManager.GetAllSkills().Where(Function(s) s.IsCustom).Count()
                End If
                
                Dim info = New With {
                    .customSkillsPath = customSkillsPath,
                    .snippetsPath = snippetsPath,
                    .customSkillsDirExists = System.IO.Directory.Exists(customSkillsPath),
                    .snippetsDirExists = System.IO.Directory.Exists(snippetsPath),
                    .skillFilesOnDisk = skillFiles,
                    .snippetFilesOnDisk = snippetFiles,
                    .customSkillManagerLoadedCount = customSkills.Count,
                    .skillManagerTotalCount = allSkillsCount,
                    .skillManagerCustomCount = allSkillsCustomCount,
                    .loadedCustomSkills = customSkills.Select(Function(s) New With {.name = s.Name, .isCustom = s.IsCustom, .isGeneric = s.IsGeneric}).ToArray()
                }
                
                Dim logPath = System.IO.Path.Combine(Utils.PathHelper.AppDataRoot, "skill_debug.log")
                System.IO.File.WriteAllText(logPath, JsonConvert.SerializeObject(info, Formatting.Indented))
                
                Return JsonConvert.SerializeObject(info)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.error = ex.Message, .stackTrace = ex.StackTrace})
            End Try
        End Function

        ''' <summary>
        ''' 删除指定 Skill
        ''' </summary>
        <ComVisible(True)>
        Public Function DeleteSkill(skillName As String) As String Implements ISkillBridge.DeleteSkill
            Try
                System.Diagnostics.Debug.WriteLine("[DeleteSkill] 开始删除 Skill: " & skillName)
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] 错误: SkillManager 未初始化")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                Dim skill = _skillManager.GetSkill(skillName)
                System.Diagnostics.Debug.WriteLine("[DeleteSkill] GetSkill 返回: " & If(skill Is Nothing, "Nothing", skill.Name & " (IsCustom=" & skill.IsCustom & ")"))

                Dim removedFromManager As Boolean = False

                If skill IsNot Nothing Then
                    ' 内存中找到了，检查是否为内置 Skill
                    If Not skill.IsCustom Then
                        System.Diagnostics.Debug.WriteLine("[DeleteSkill] 错误: 内置 Skill 不可删除")
                        Return JsonConvert.SerializeObject(New With {.success = False, .error = "内置 Skill 不可删除"})
                    End If

                    ' 从 SkillManager 内存中删除
                    removedFromManager = _skillManager.RemoveSkill(skillName)
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] RemoveSkill 返回: " & removedFromManager)
                    If removedFromManager Then
                        _skillManager.SaveSkills()
                    End If
                Else
                    ' 内存中没找到：可能是YAML解析失败导致没加载，仍尝试从磁盘删除
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] 内存中未找到 Skill: " & skillName & ", 尝试直接从磁盘删除")
                End If

                ' 无论内存中是否找到，都调用 CustomSkillManager 确保磁盘文件被删除
                Dim deletedFromCustomManager As Boolean = False
                Try
                    Dim customManager = Skills.CustomSkillManager.GetInstance()
                    deletedFromCustomManager = customManager.DeleteCustomSkill(skillName)
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] CustomSkillManager.DeleteCustomSkill 返回: " & deletedFromCustomManager)
                Catch customEx As Exception
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] 从 CustomSkillManager 删除失败: " & customEx.Message)
                End Try

                If removedFromManager OrElse deletedFromCustomManager Then
                    Return JsonConvert.SerializeObject(New With {.success = True, .message = "Skill 已删除"})
                Else
                    System.Diagnostics.Debug.WriteLine("[DeleteSkill] 错误: 内存和磁盘均未找到可删除的 Skill")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "找不到 Skill: " & skillName})
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("[DeleteSkill] 异常: " & ex.Message & ", 堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 保存 Skill（添加或更新）
        ''' </summary>
        <ComVisible(True)>
        Public Function SaveSkill(skillJson As String) As String Implements ISkillBridge.SaveSkill
            Try
                System.Diagnostics.Debug.WriteLine("[SaveSkill] 开始保存 Skill, JSON长度: " & skillJson.Length)
                ' 【关键修复】必须先初始化技能系统，否则 _skillManager 和 _customSkillManager 可能为 Nothing
                EnsureSkillsInitialized()
                If _skillManager Is Nothing Then
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 错误: SkillManager 未初始化")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "SkillManager 未初始化"})
                End If

                ' 解析 JSON（使用 SkillJsonHelper 不区分大小写的反序列化）
                Dim skill = Skills.SkillJsonHelper.DeserializeSkill(skillJson)
                System.Diagnostics.Debug.WriteLine("[SaveSkill] 反序列化结果: " & If(skill Is Nothing, "Nothing", "Name=" & skill.Name))
                If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 错误: 无法解析 Skill 数据或缺少 Name 字段")
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无法解析 Skill 数据或缺少 Name 字段"})
                End If

                ' 【完善 SkillCreator】可执行性校验——拒绝保存"空壳"技能。
                ' 允许的实现形态：① 声明式 steps（至少一步有 tool 或 code）② VBA/Python templateCode
                ' ③ LLM 提示词 llmSystemPrompt（多模态）④ pipeline 路由（html/python 等）。
                ' 四者全无 = 只有 description 的空定义，禁止落盘（防止生成"确立了但没法跑"的假 Skill）。
                Dim hasValidSteps = skill.Steps IsNot Nothing AndAlso
                                    skill.Steps.Any(Function(s) (Not String.IsNullOrWhiteSpace(s.Tool)) OrElse (Not String.IsNullOrWhiteSpace(s.Code)))
                Dim hasCode = Not String.IsNullOrWhiteSpace(skill.TemplateCode)
                Dim hasLLM = Not String.IsNullOrWhiteSpace(skill.LLMSystemPrompt)
                Dim hasPipeline = Not String.IsNullOrWhiteSpace(skill.Pipeline)
                ' spec 形态：由前端规划引擎按 specFields 生成步骤，本身不携带 steps/code，
                ' 属于合法可执行形态（SmartTableBeautifier 等内置技能即为此类）。
                Dim hasSpec = String.Equals(skill.ExecutionMode, "spec", StringComparison.OrdinalIgnoreCase) OrElse
                              (skill.SpecFields IsNot Nothing AndAlso skill.SpecFields.Count > 0)
                If Not (hasValidSteps OrElse hasCode OrElse hasLLM OrElse hasPipeline OrElse hasSpec) Then
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 可执行性校验失败（空壳技能）: " & skill.Name)
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .executable = False,
                        .reason = "NO_EXECUTABLE_IMPL",
                        .error = "该技能没有任何可执行实现（缺少 声明式 steps / VBA·Python 代码 / LLM 提示词 / pipeline / spec 参数），无法保存为空壳技能。请补充实现，或让 SkillCreator 重新生成。"
                    })
                End If

                ' 检查是否已存在
                Dim existingSkill = _skillManager.GetSkill(skill.Name)
                System.Diagnostics.Debug.WriteLine("[SaveSkill] existingSkill: " & If(existingSkill Is Nothing, "Nothing", "Name=" & existingSkill.Name & ", IsCustom=" & existingSkill.IsCustom))
                If existingSkill IsNot Nothing Then
                    ' 更新现有 Skill
                    skill.IsCustom = True  ' 用户保存的 Skill 标记为自定义
                    skill.LastModifiedDate = DateTime.Now
                    _skillManager.AddSkill(skill)
                    System.Diagnostics.Debug.WriteLine("更新 Skill: " & skill.Name)
                Else
                    ' 添加新 Skill
                    skill.IsCustom = True  ' 用户保存的 Skill 标记为自定义
                    skill.Version = 1
                    skill.CreatedDate = DateTime.Now
                    skill.LastModifiedDate = DateTime.Now
                    _skillManager.AddSkill(skill)
                    System.Diagnostics.Debug.WriteLine("添加新 Skill: " & skill.Name)
                End If

                ' 保存系统内置技能到 bin 目录（自定义技能会被自动跳过）
                _skillManager.SaveSkills()
                System.Diagnostics.Debug.WriteLine("[SaveSkill] 系统技能已保存到 Skills 目录（自定义技能已跳过）")

                ' 【关键修复】保存自定义技能到 AppData 目录
                Try
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 尝试保存到 CustomSkillManager, IsCustom=" & skill.IsCustom)
                    If _customSkillManager IsNot Nothing Then
                        _customSkillManager.SaveCustomSkill(skill)
                        System.Diagnostics.Debug.WriteLine("[SaveSkill] ✅ 自定义技能已保存到 CustomSkillManager: " & skill.Name)
                    Else
                        System.Diagnostics.Debug.WriteLine("[SaveSkill] ⚠️ CustomSkillManager 未初始化")
                    End If
                Catch customSaveEx As Exception
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 保存自定义技能失败: " & customSaveEx.Message & ", 堆栈: " & customSaveEx.StackTrace)
                End Try

                ' 验证自定义技能文件是否已保存到 AppData
                Try
                    Dim customSkillsDir = Path.Combine(Utils.PathHelper.CustomSkillsPath, "skills")
                    Dim snippetsDir = Path.Combine(Utils.PathHelper.CustomSkillsPath, "snippets")
                    Dim mdPath = Path.Combine(customSkillsDir, skill.Name & ".md")
                    If Not System.IO.File.Exists(mdPath) Then
                        mdPath = Path.Combine(snippetsDir, skill.Name & ".md")
                    End If
                    If System.IO.File.Exists(mdPath) Then
                        System.Diagnostics.Debug.WriteLine("[SaveSkill] ✅ 验证成功: 自定义技能文件已创建 - " & mdPath)
                    Else
                        System.Diagnostics.Debug.WriteLine("[SaveSkill] ⚠️ 警告: 未找到自定义技能文件, 期望路径: " & Path.Combine(customSkillsDir, skill.Name & ".md"))
                    End If
                Catch verifyEx As Exception
                    System.Diagnostics.Debug.WriteLine("[SaveSkill] 验证文件失败: " & verifyEx.Message)
                End Try

                ' 【关键修复】清除意图识别缓存,让新Skill立即生效
                Try
                    If _intentEngine IsNot Nothing Then
                        _intentEngine.ClearCache()
                        System.Diagnostics.Debug.WriteLine("已清除意图识别缓存")
                    End If
                Catch cacheEx As Exception
                    System.Diagnostics.Debug.WriteLine("清除缓存失败: " & cacheEx.Message)
                End Try

                Return JsonConvert.SerializeObject(New With {.success = True, .message = "Skill 已保存"})
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("SaveSkill 错误: " & ex.Message & ", 堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

#End Region

#Region "IBridge 生命周期"

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "SkillBridge"
            End Get
        End Property

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                ' SkillBridge 无外部资源需要释放
                _disposed = True
            End If
        End Sub

#End Region

#Region "代码溢出：工具结果出口 JSON / CSS 字典尾巴清洗"

        ''' <summary>
        ''' 与前端 sanitizeLeakedJsonSnippets 对称的 VB.NET 版清洗函数。
        ''' LLM（尤其是小模型）会把 tool_result 里「样式字典 JSON」当正文直接抄进总结回复——
        ''' 表现为用户消息末尾贴一段 ,"border":{"top":"1px solid black",...},"fill":{...},"font":{...}。
        ''' 这里在后端出口先对面向 LLM 的 Message / Outputs 字符串做剥离，形成"即使前端没缓存更新也不暴露"的兜底。
        ''' </summary>
        ''' <remarks>
        ''' 只做最典型的尾部剥离，不做激进的整段正则吞吃，避免误伤真正的 JSON 类结果（如 GetWorkbookInfo）：
        '''   1) 末尾一段包含 "border"/"fill"/"font"/"alignment" 键的嵌套对象（通常前面还有多余逗号）
        '''   2) 行尾孤立逗号 / 中文逗号
        '''   3) 多余空白
        ''' </remarks>
        Friend Shared Function SanitizeLeakedJsonSnippets(text As String) As String
            If String.IsNullOrWhiteSpace(text) Then Return text
            Dim s As String = text

            ' 1) 最典型：以"border":{"top":...} 或 ,"fill":{"color":...} 开头的整段尾巴，包到文件末尾
            '    用一次性 Replace 包装正则 + 末尾锚点 $，确保只清尾巴，清不到正文里的 JSON
            s = System.Text.RegularExpressions.Regex.Replace(
                s,
                ",?\s*\{?\s*""border""\s*:\s*\{[\s\S]*?\}(?:\s*,\s*""[^""]+""\s*:\s*(?:""[^""]*""|\{[^}]*\}|\[[^\]]*\]|true|false|null|-?\d+(?:\.\d+)?))*\s*\}?\s*$",
                "",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            ' 2) 兜底：行尾一段独立对象片段（深度>=2 且包含显式样式键）
            '    仅当尾巴里包含 2 个以上 '{' 才当裸尾巴清掉，避免误伤单行 JSON
            s = System.Text.RegularExpressions.Regex.Replace(
                s,
                ",\s*\{(?=[\s\S]{0,500}?(?:""border""|""fill""|""font""|""alignment""))[\s\S]{10,}\s*$",
                Function(m As System.Text.RegularExpressions.Match)
                     Dim inner = m.Value
                     Dim openBrace As Integer = 0
                     For Each ch In inner
                         If ch = "{"c Then openBrace += 1
                     Next
                     Return If(openBrace >= 2, "", m.Value)
                 End Function,
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            ' 3) 每行行尾孤立逗号 / 中文逗号（JSON 被截断后残留）
            s = System.Text.RegularExpressions.Regex.Replace(s, "[,，]\s*$", "", System.Text.RegularExpressions.RegexOptions.Multiline Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            ' 4) 多余空白（收尾整理）
            s = System.Text.RegularExpressions.Regex.Replace(s, "[ \t]{2,}", " ", System.Text.RegularExpressions.RegexOptions.CultureInvariant)
            s = System.Text.RegularExpressions.Regex.Replace(s, "\n{3,}", Environment.NewLine & Environment.NewLine, System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            Return s.TrimEnd()
        End Function

#End Region

    End Class

End Namespace
