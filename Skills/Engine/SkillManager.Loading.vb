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

    ' SkillManager.Loading.vb — Skill 加载与升级（Markdown/Excel/Ocr 解析 + 校验）
    ' （类级 XML 文档只在主文件 SkillManager.vb 保留一份，避免 BC42314 分部重复）
    Partial Public Class SkillManager

        ''' <summary>
        ''' 动态加载单个 Skill（从 Markdown 文件），用于运行时按需加载
        ''' </summary>
        Public Sub ReloadSingleSkill(skillName As String)
            Try
                Dim mdDir = Utils.PathHelper.GetMarkdownSkillsDirectory()
                If Not Directory.Exists(mdDir) Then Return

                ' 查找匹配的 .md 文件（不区分大小写）
                Dim mdFile As String = Nothing
                For Each file In Directory.GetFiles(mdDir, "*.md")
                    Dim fileName = Path.GetFileNameWithoutExtension(file)
                    If String.Equals(fileName, skillName, StringComparison.OrdinalIgnoreCase) Then
                        mdFile = file
                        Exit For
                    End If
                Next

                If mdFile Is Nothing Then
                    System.Diagnostics.Debug.WriteLine($"[ReloadSingleSkill] 未找到 {skillName}.md")
                    Return
                End If

                Dim skill = ParseMarkdownSkill(mdFile)
                If skill IsNot Nothing AndAlso Not _excludedSkills.Contains(skill.Name) Then
                    _skills(skill.Name) = skill
                    _skillsListCacheValid = False  ' 使列表缓存失效
                    _skillCategoriesCacheValid = False  ' 使分类缓存失效
                    System.Diagnostics.Debug.WriteLine($"[ReloadSingleSkill] ✅ 成功加载: {skill.Name}")
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ReloadSingleSkill] 加载失败: {ex.Message}")
            End Try
        End Sub

        Public Sub LoadSkills()
            ' 在 Try 块外部定义 loadedSkills,以便在 Catch 块中也能访问
            Dim loadedSkills As New Dictionary(Of String, SkillDefinition)(StringComparer.OrdinalIgnoreCase)
        
            Try
                System.Diagnostics.Debug.WriteLine("========== LoadSkills called ==========")
                System.Diagnostics.Debug.WriteLine("Current code version: " & CurrentSkillVersion.ToString())
        
                ' 首先从 Markdown 文件加载 Skills(主要数据源)
                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] ========== 步骤1: 从 Markdown 文件加载 Skills...")
                Dim mdCount = LoadMarkdownSkills(loadedSkills)
                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 从 Markdown 加载了 " & mdCount.ToString() & " 个 skills")
                For Each key In loadedSkills.Keys
                    System.Diagnostics.Debug.WriteLine("[SKILL DEBUG]   - 已加载: " & key)
                Next
        
                ' 如果没有从 Markdown 加载到任何 Skill,使用代码中的默认配置
                If loadedSkills.Count = 0 Then
                    System.Diagnostics.Debug.WriteLine("没有加载到任何 Skill,使用代码默认配置...")
                    CreateDefaultSkills()
                    System.Diagnostics.Debug.WriteLine("代码默认配置创建完成,_skills 数量: " & _skills.Count.ToString())
                    ' 将默认 Skills 复制到 loadedSkills
                    For Each kvp In _skills
                        loadedSkills(kvp.Key) = kvp.Value
                    Next
                    System.Diagnostics.Debug.WriteLine("复制到 loadedSkills 完成,数量: " & loadedSkills.Count.ToString())
                End If
        
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("Error in LoadSkills (Markdown部分): " & ex.Message)
                System.Diagnostics.Debug.WriteLine("异常堆栈: " & ex.StackTrace)
                CreateDefaultSkills()
                ' 将默认 Skills 复制到 loadedSkills
                For Each kvp In _skills
                    loadedSkills(kvp.Key) = kvp.Value
                Next
            End Try
        
            ' 【关键修复】无论上面是否出错,都必须加载 CustomSkill
            Utils.AppLogger.LogInfo("SkillManager", "步骤2: 加载自定义Skills")
            Try
                ' 【修复】重置单例,强制重新加载
                CustomSkillManager.ResetInstance()
                Dim customManager = CustomSkillManager.GetInstance()
                
                Dim customSkills = customManager.GetAllCustomSkills()
                Utils.AppLogger.LogInfo("SkillManager", $"自定义Skills路径: {Utils.PathHelper.CustomSkillsPath}")
                Utils.AppLogger.LogInfo("SkillManager", $"自定义Skills数量: {customSkills.Count}")
                For Each cs In customSkills
                    System.Diagnostics.Debug.WriteLine($"[LoadSkills] 自定义Skill: {cs.Name}, IsCustom={cs.IsCustom}, IsGeneric={cs.IsGeneric}")
                Next
                
                For Each skill In customSkills
                    Utils.AppLogger.LogDebug("CustomSkill", $"[{skill.Name}] {skill.Description.Substring(0, Math.Min(30, skill.Description.Length))}...")

                    If loadedSkills.ContainsKey(skill.Name) Then
                        ' 【修复】以 Markdown 系统版本为基础，只覆盖用户可编辑字段，
                        ' 确保系统新增字段（Pipeline、InputContext、Category、OutputFormat 等）不会被旧的自定义版本丢失。
                        Dim systemSkill = loadedSkills(skill.Name)

                        ' 行为类字段（TemplateCode/LLMSystemPrompt/SpecFields/QuickPrompts）
                        ' 是否允许被自定义版本覆盖：仅当自定义版本号 >= 系统版本号时（即用户基于当前/更新版编辑保存），
                        ' 才采用自定义内容；否则（旧残留自定义）保持系统版本，避免拖累系统升级。
                        Dim allowBehaviorOverride = (skill.Version >= systemSkill.Version)

                        ' 【同名合并防御】内置 .md 比 AppData 里的自定义副本更新 → 判定"系统已升级"，
                        ' 行为字段（LLMSystemPrompt/TemplateCode/SpecFields/QuickPrompts）一律保留内置版。
                        '
                        ' 为什么不能只比版本号：用户点「快捷方式 / 启用开关 / 保存」会把内置技能**整份快照**
                        ' 写进 AppData\CustomSkills\skills\{name}.md（SkillBridge.ToggleQuickShortcut /
                        ' ToggleSkillEnabled / SaveSkill → CustomSkillManager.SaveCustomSkill），快照连
                        ' Version 一起拷走，于是 `custom.Version >= system.Version` 恒成立。实测所有
                        ' 内置技能两侧都是 `version: 10`，旧提示词因此永久压住系统升级——改 Skills\*.md
                        ' 后编译、重启 Excel 都不生效（用户可见症状："我改了提示词，LLM 还在用旧的"）。
                        ' 文件时间能区分两种"更新"：系统内置文件刚被写入 = 系统升级，应以内置为准；
                        ' 用户刚在扩展中心保存 = 自定义副本更新，仍按版本规则让自定义生效。
                        Dim systemFileNewer As Boolean = (systemSkill.SourceLastWriteTime > DateTime.MinValue AndAlso
                                                          skill.SourceLastWriteTime > DateTime.MinValue AndAlso
                                                          systemSkill.SourceLastWriteTime > skill.SourceLastWriteTime)
                        If systemFileNewer Then
                            Utils.AppLogger.LogWarning("SkillManager",
                                $"[MergeCustomSkill] {skill.Name} 内置文件({systemSkill.SourceLastWriteTime:yyyy-MM-dd HH:mm:ss})比自定义副本({skill.SourceLastWriteTime:yyyy-MM-dd HH:mm:ss})新，行为字段保留系统版（避免旧快照压住升级）")
                            allowBehaviorOverride = False
                        End If

                        ' 用户快捷方式设置（持久化到自定义版本）→ 始终自定义优先（非行为类）
                        systemSkill.IsQuickShortcut = skill.IsQuickShortcut
                        systemSkill.QuickShortcutIcon = skill.QuickShortcutIcon
                        systemSkill.QuickShortcutPrompt = skill.QuickShortcutPrompt

                        ' Skill 开关：用户自定义版本的 IsEnabled 始终优先（运行时开关通过这里持久化）
                        System.Diagnostics.Debug.WriteLine($"[MergeCustomSkill] IsEnabled: system={systemSkill.IsEnabled} -> custom={skill.IsEnabled}, name='{skill.Name}'")
                        systemSkill.IsEnabled = skill.IsEnabled

                        ' 展示类字段：始终自定义优先（用户 UI 编辑的显示内容）
                        If Not String.IsNullOrWhiteSpace(skill.DisplayName) Then systemSkill.DisplayName = skill.DisplayName
                        If Not String.IsNullOrWhiteSpace(skill.Description) Then systemSkill.Description = skill.Description
                        If skill.Keywords IsNot Nothing AndAlso skill.Keywords.Length > 0 Then systemSkill.Keywords = skill.Keywords
                        If skill.Examples IsNot Nothing AndAlso skill.Examples.Length > 0 Then systemSkill.Examples = skill.Examples
                        If skill.UIActions IsNot Nothing Then systemSkill.UIActions = skill.UIActions

                        ' 多模态输入能力（用户在 Skill 编辑对话框里勾选）→ 始终自定义优先
                        ' 【修复】此前未合并这两个字段，导致用户在扩展面板勾选
                        '        「支持图片输入 / 支持文件输入」保存（已写入自定义副本）后，
                        '        重新加载又被系统 Markdown 的 false 覆盖，表现为复选框“勾不上”。
                        System.Diagnostics.Debug.WriteLine($"[MergeCustomSkill] SupportsImageInput: system={systemSkill.SupportsImageInput} -> custom={skill.SupportsImageInput}, name='{skill.Name}'")
                        systemSkill.SupportsImageInput = skill.SupportsImageInput
                        systemSkill.SupportsFileInput = skill.SupportsFileInput

                        ' 行为类字段：版本保护（仅当自定义版本 >= 系统版本时才覆盖）
                        If allowBehaviorOverride Then
                            If Not String.IsNullOrWhiteSpace(skill.TemplateCode) Then systemSkill.TemplateCode = skill.TemplateCode
                            ' 【残缺提示词防御】自定义副本里的 System Prompt 是按 MD 的 "## System Prompt" 段落抽出来的，
                            ' 而 GenerateOCRParseScript 这类"整篇正文即提示词"的 Skill，正文里那行 "## System Prompt"
                            ' 只是小节标题 → 抽出来只剩一行标题（"# 生成OCR解析脚本"）。一旦拿它覆盖系统提示词，
                            ' Skill 的指令就被清空成一行字（实测：扩展中心里保存一次后，该 Skill 再也生不出合规脚本，
                            ' 提示词从几千字变成 8 个字）。标题级残片一律不算覆盖，保留系统完整提示词。
                            If Not String.IsNullOrWhiteSpace(skill.LLMSystemPrompt) Then
                                If IsUsableCustomPrompt(skill.LLMSystemPrompt, skill.DisplayName, systemSkill.DisplayName, systemSkill.LLMSystemPrompt) Then
                                    systemSkill.LLMSystemPrompt = skill.LLMSystemPrompt
                                Else
                                    Utils.AppLogger.LogWarning("SkillManager",
                                        $"[MergeCustomSkill] {skill.Name} 自定义提示词疑似残缺（{skill.LLMSystemPrompt.Length} 字），已忽略，保留系统提示词")
                                End If
                            End If
                            ' SpecFields/QuickPrompts 仅在自定义声明了非空列表时才覆盖，
                            ' 避免旧副本（未声明这两个字段）把系统新字段清空。
                            If skill.SpecFields IsNot Nothing AndAlso skill.SpecFields.Count > 0 Then systemSkill.SpecFields = skill.SpecFields
                            If skill.QuickPrompts IsNot Nothing AndAlso skill.QuickPrompts.Count > 0 Then systemSkill.QuickPrompts = skill.QuickPrompts
                        End If

                        ' 标记为用户已自定义（但保留系统字段如 Pipeline/InputContext/Category 等来自 Markdown 的值）
                        systemSkill.IsCustom = True

                        ' 系统定义字段（Pipeline、InputContext、ProcessingMode、Category、OutputFormat、Version 等）
                        ' 不从自定义版本覆盖，始终以 Markdown 文件中定义的为准，确保系统升级后新字段生效

                        loadedSkills(skill.Name) = systemSkill
                    Else
                        ' 纯自定义技能（无同名 Markdown 版本），直接添加
                        loadedSkills(skill.Name) = skill
                    End If
                Next
                Utils.AppLogger.LogInfo("SkillManager", $"加载了{customSkills.Count}个自定义Skill")
            Catch ex As Exception
                Utils.AppLogger.LogError("SkillManager", "加载自定义Skill失败", ex)
            End Try
        
            ' 确保 loadedSkills 被赋值给 _skills(即使在异常情况下)
            _skills = loadedSkills
            ' 【停用状态固化】在 .md 合并之后强制套用独立停用名单，确保用户点"停用"后
            ' 在热重载/重读时依然保持停用（不被系统 .md 默认 IsEnabled=True 覆盖）。
            ApplyDisabledSkillNames(loadedSkills)
            InvalidateCache()  ' 【修复】清除列表缓存，确保 GetAllSkills 返回最新数据
            Utils.AppLogger.LogInfo("SkillManager", $"LoadSkills完成,总计:{_skills.Count}个Skill")
            LogToFile("========== LoadSkills 完成,总计: " & _skills.Count.ToString() & " 个 skills ==========")

            '  将所有 Skill 注册到统一能力注册中心
            Try
                RegisterSkillsToCommandRegistry()
            Catch ex As Exception
                Utils.AppLogger.LogError("SkillManager", "注册 Skill 到能力注册中心失败", ex)
            End Try
        End Sub

        ' ════════════════════════════════════════════════════════════════
        ' 【停用状态固化】独立的停用名单（与 .md 合并解耦）
        ' 背景：ToggleSkillEnabled 把 IsEnabled 写进 AppData 自定义副本，但
        ' SkillManager 在热重载/重读时按系统 .md 重新合并，自定义副本的
        ' isEnabled:false 可能未被稳定合入，导致"点了停用又弹回启用"。
        ' 这里用一份独立 JSON 名单，在合并之后强制套用，保证停用一定生效。
        ' ════════════════════════════════════════════════════════════════
        Private _disabledSkillNames As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        Private Shared ReadOnly DisabledSkillsFilePath As String =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                         "ExcelAIVSTOAddIn", "disabled_skills.json")

        Private Sub LoadDisabledSkillNames()
            _disabledSkillNames.Clear()
            Try
                If File.Exists(DisabledSkillsFilePath) Then
                    Dim arr = Newtonsoft.Json.Linq.JArray.Parse(File.ReadAllText(DisabledSkillsFilePath))
                    For Each t As Newtonsoft.Json.Linq.JToken In arr
                        Dim n = t.ToString().Trim()
                        If n <> "" Then _disabledSkillNames.Add(n)
                    Next
                End If
            Catch ex As Exception
                Utils.AppLogger.LogError("SkillManager", "加载停用名单失败: " & ex.Message, ex)
            End Try
        End Sub

        Private Sub SaveDisabledSkillNames()
            Try
                Dim dir = Path.GetDirectoryName(DisabledSkillsFilePath)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                Dim arr As New Newtonsoft.Json.Linq.JArray()
                For Each n In _disabledSkillNames
                    arr.Add(n)
                Next
                File.WriteAllText(DisabledSkillsFilePath, arr.ToString(Newtonsoft.Json.Formatting.None))
            Catch ex As Exception
                Utils.AppLogger.LogError("SkillManager", "保存停用名单失败: " & ex.Message, ex)
            End Try
        End Sub

        Private Sub ApplyDisabledSkillNames(loadedSkills As Dictionary(Of String, SkillDefinition))
            LoadDisabledSkillNames()
            If _disabledSkillNames.Count = 0 Then Return
            For Each dn In _disabledSkillNames
                If loadedSkills.ContainsKey(dn) Then
                    loadedSkills(dn).IsEnabled = False
                    Utils.AppLogger.LogInfo("SkillManager", $"[停用固化] {dn} 强制设为停用（用户停用名单）")
                End If
            Next
        End Sub

        ''' <summary>
        ''' 运行时切换技能启用/停用，并固化到独立停用名单（热重载后仍生效）
        ''' </summary>
        Public Sub SetSkillDisabled(skillName As String, isEnabled As Boolean)
            If isEnabled Then _disabledSkillNames.Remove(skillName) Else _disabledSkillNames.Add(skillName)
            SaveDisabledSkillNames()
            If _skills.ContainsKey(skillName) Then _skills(skillName).IsEnabled = isEnabled
            InvalidateCache()
        End Sub

        ''' <summary>
        ''' 从外部配置文件加载默认 Skills
        ''' </summary>
        Private Sub LogToFile(message As String)
            Try
                Dim logPath = Utils.PathHelper.DebugLogFilePath
                Utils.PathHelper.EnsureDirectoryExists(logPath)
                File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") & " " & message & Environment.NewLine)
            Catch ex As Exception
                ' 忽略日志写入错误，但记录到 Debug 输出以便诊断
                System.Diagnostics.Debug.WriteLine("LogToFile 失败: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' 自定义副本带来的 System Prompt 是否"可用"（不是标题级残片）。
        ''' 残片特征：单行 + 形如 "# 显示名" / 等于显示名 / 等于系统提示词首行，或短到无法承载指令。
        ''' 背景：自定义副本的提示词是按 MD 的 "## System Prompt" 段落抽取的，对"整篇正文即提示词"的
        ''' Skill（GenerateOCRParseScript）只能抽到那一行小节标题，若允许它覆盖，Skill 指令会被清空。
        ''' </summary>
        Private Function IsUsableCustomPrompt(customPrompt As String, customDisplayName As String,
                                              systemDisplayName As String, systemPrompt As String) As Boolean
            Dim p = If(customPrompt, "").Trim()
            If p.Length = 0 Then Return False

            ' 真正的提示词几乎都是多行；多行内容直接认（避免误伤用户自己写的短提示词）
            If p.Contains(vbLf) OrElse p.Contains(vbCr) Then Return True

            ' 单行 + 标题形式（"# xxx" / "## xxx"）→ 残片
            If p.StartsWith("#") Then Return False

            For Each candidate In New String() {customDisplayName, systemDisplayName}
                If String.IsNullOrWhiteSpace(candidate) Then Continue For
                Dim c = candidate.Trim()
                If String.Equals(p, c, StringComparison.OrdinalIgnoreCase) OrElse
                   String.Equals(p, "# " & c, StringComparison.OrdinalIgnoreCase) Then Return False
            Next

            If Not String.IsNullOrWhiteSpace(systemPrompt) Then
                Dim lines = systemPrompt.Trim().Split(New String() {vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                If lines.Length > 0 AndAlso String.Equals(p, lines(0).Trim(), StringComparison.OrdinalIgnoreCase) Then Return False
            End If

            ' 单行且极短 → 不可能是一条可用的 Skill 指令
            Return p.Length >= 40
        End Function

        ''' <summary>
        ''' 从 Markdown 文件加载 Skills（YAML frontmatter + Markdown 正文）
        ''' 【已修改】识别类Skill黑名单，这些Skill已被移到backup目录，不会加载
        ''' </summary>
        Private ReadOnly _excludedSkills As String() = New String() {
            "ExtractIDCard", "ExtractInvoice", "RecognizeTable",
            "ExtractBusinessCard", "ExtractBankCard", "ExtractDriverLicense",
            "ExtractBusinessLicense", "ExtractDocument", "ExtractTableData",
            "ExtractText", "AnalyzeImage", "SmartRecognize", "TableToExcel",
            "ExtractDate"
        }

        Private Function LoadMarkdownSkills(ByRef loadedSkills As Dictionary(Of String, SkillDefinition)) As Integer
            Dim count As Integer = 0
            Try
                Dim mdDir = Utils.PathHelper.GetMarkdownSkillsDirectory()
                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] Markdown Skills 目录: " & mdDir)
                LogToFile("[SKILL DEBUG] Markdown Skills 目录: " & mdDir)

                If Not Directory.Exists(mdDir) Then
                    System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] Markdown Skills 目录不存在，跳过")
                    Return 0
                End If

                Dim mdFiles = Directory.GetFiles(mdDir, "*.md")
                System.Diagnostics.Debug.WriteLine(String.Format("[SKILL DEBUG] 发现 {0} 个 .md 文件", mdFiles.Length))
                LogToFile(String.Format("[SKILL DEBUG] 发现 {0} 个 .md 文件", mdFiles.Length))

                For Each mdFile In mdFiles
                    Try
                        Dim skill = ParseMarkdownSkill(mdFile)
                        ' 【同名合并防御】记录来源文件与写入时间：合并 AppData 自定义副本时用它判断
                        ' "内置文件是否比那份旧快照新"（快照带同一 Version，光比版本号会被旧快照反超）
                        If skill IsNot Nothing Then
                            skill.SourceFilePath = mdFile
                            Try
                                skill.SourceLastWriteTime = File.GetLastWriteTime(mdFile)
                            Catch exTime As Exception
                                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 读取内置 Skill 文件时间失败: " & exTime.Message)
                            End Try
                        End If
                        ' 【已修改】跳过识别类Skill
                        If skill IsNot Nothing AndAlso Not _excludedSkills.Contains(skill.Name) Then
                            loadedSkills(skill.Name) = skill
                            count += 1
                            System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 从 Markdown 加载 Skill: " & skill.Name)
                            LogToFile("[SKILL DEBUG] 从 Markdown 加载 Skill: " & skill.Name)
                        ElseIf skill IsNot Nothing AndAlso _excludedSkills.Contains(skill.Name) Then
                            System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 跳过识别类Skill: " & skill.Name)
                            LogToFile("[SKILL DEBUG] 跳过识别类Skill: " & skill.Name)
                        End If
                    Catch ex As Exception
                        System.Diagnostics.Debug.WriteLine(String.Format("[SKILL DEBUG] 加载 Markdown Skill 失败 ({0}): {1}", Path.GetFileName(mdFile), ex.Message))
                        LogToFile(String.Format("[SKILL DEBUG] 加载 Markdown Skill 失败 ({0}): {1}", Path.GetFileName(mdFile), ex.Message))
                    End Try
                Next

                System.Diagnostics.Debug.WriteLine(String.Format("[SKILL DEBUG] 成功从 Markdown 加载 {0} 个 Skills", count))
                LogToFile(String.Format("[SKILL DEBUG] 成功从 Markdown 加载 {0} 个 Skills", count))
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("加载 Markdown Skills 失败: " & ex.Message)
                LogToFile("加载 Markdown Skills 失败: " & ex.Message)
            End Try
            Return count
        End Function

        ''' <summary>
        ''' 解析单个 Markdown Skill 文件（使用 Newtonsoft.Json 解析 YAML frontmatter）
        ''' </summary>
        Private Function ParseMarkdownSkill(filePath As String) As SkillDefinition
            Dim content = File.ReadAllText(filePath, System.Text.Encoding.UTF8)
            If String.IsNullOrWhiteSpace(content) Then Return Nothing

            Dim yamlText As String = ""
            Dim markdownBody As String = ""
            ExtractYamlFrontmatter(content, yamlText, markdownBody)
            If String.IsNullOrWhiteSpace(yamlText) Then Return Nothing

            Dim yamlObj = ConvertYamlToJObject(yamlText)
            If yamlObj Is Nothing OrElse yamlObj.Count = 0 Then
                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] YAML 解析失败: " & filePath)
                Return Nothing
            End If

            Dim skill = New SkillDefinition()
            PopulateSkillBasicFields(skill, yamlObj)
            PopulateSkillTextContent(skill, yamlObj, markdownBody, filePath)
            ParseExcelConfig(skill, yamlObj)
            ParseOcrConfig(skill, yamlObj)

            System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] 成功解析 Markdown Skill: " & skill.Name)
            ValidateSkillDefinition(skill, filePath)
            Return skill
        End Function

        Private Sub ExtractYamlFrontmatter(content As String, ByRef yamlText As String, ByRef markdownBody As String)
            If content.StartsWith("---") Then
                Dim endIndex = content.IndexOf("---", 3)
                If endIndex > 0 Then
                    yamlText = content.Substring(3, endIndex - 3).Trim()
                    markdownBody = content.Substring(endIndex + 3).Trim()
                End If
            End If
        End Sub

        Private Sub PopulateSkillBasicFields(skill As SkillDefinition, yamlObj As Newtonsoft.Json.Linq.JObject)
            skill.Name = GetJsonString(yamlObj, "name")
            skill.DisplayName = GetJsonString(yamlObj, "displayName")
            skill.Description = GetJsonString(yamlObj, "description")
            skill.Keywords = GetJsonStringArray(yamlObj, "keywords")
            skill.AllowedTools = GetJsonStringArray(yamlObj, "allowedTools")
            skill.Examples = GetJsonStringArray(yamlObj, "examples")
            skill.IsGeneric = GetJsonBool(yamlObj, "isGeneric", True)
            skill.NeedsConfig = GetJsonBool(yamlObj, "needsConfig", False)
            skill.ConfigFields = GetJsonStringArray(yamlObj, "configFields")
            skill.SupportsImageInput = GetJsonBool(yamlObj, "supportsImageInput", False)
            skill.SupportsFileInput = GetJsonBool(yamlObj, "supportsFileInput", False)
            ' 输出自动导入 OCR 智能解析面板（GenerateOCRParseScript 等）
            skill.AutoImportParseScript = GetJsonBool(yamlObj, "autoImportParseScript", False)
            skill.OutputFormat = GetJsonString(yamlObj, "outputFormat", "text")
            skill.IsCustom = GetJsonBool(yamlObj, "isCustom", False)
            skill.AllowMcpTools = GetJsonBool(yamlObj, "allowMcpTools", False)
            skill.Category = GetJsonString(yamlObj, "category", "")
            skill.Tags = GetJsonStringList(yamlObj, "tags")
            skill.IsQuickShortcut = GetJsonBool(yamlObj, "isQuickShortcut", False)
            skill.QuickShortcutIcon = GetJsonString(yamlObj, "quickShortcutIcon", "⚡")
            skill.QuickShortcutPrompt = GetJsonString(yamlObj, "quickShortcutPrompt", "")
            skill.Version = GetJsonInt(yamlObj, "version", CurrentSkillVersion)
            ' Skill 开关（默认启用）
            Dim rawEnabled = GetJsonBool(yamlObj, "isEnabled", True)
            skill.IsEnabled = rawEnabled
            System.Diagnostics.Debug.WriteLine($"[PopulateSkillBasicFields] Parse IsEnabled: name='{skill.Name}', raw={rawEnabled}, yamlHasKey={yamlObj.ContainsKey("isEnabled")}")

            skill.CreatedDate = DateTime.Now
            skill.LastModifiedDate = DateTime.Now
            skill.CreatedBy = "Markdown"
            skill.UsageCount = 0
            skill.Pipeline = GetJsonString(yamlObj, "pipeline", "")
            skill.InputContext = GetJsonStringArray(yamlObj, "inputContext")
            ' 解析 frontmatter 中的 quickPrompts 对象数组
            skill.QuickPrompts = GetJsonQuickPrompts(yamlObj, "quickPrompts")
            ' 解析 frontmatter 中的 specFields 对象数组
            skill.SpecFields = GetJsonSpecFields(yamlObj, "specFields")
            ' 解析 frontmatter 中的 uiActions 对象（复制/保存为MD/重新生成 按钮开关）
            skill.UIActions = GetJsonUIActions(yamlObj, "uiActions")

            Dim pmStr = GetJsonString(yamlObj, "processingMode", "VBACode")
            If Integer.TryParse(pmStr, Nothing) Then
                skill.ProcessingMode = CType(Integer.Parse(pmStr), SkillProcessingMode)
            Else
                skill.ProcessingMode = CType([Enum].Parse(GetType(SkillProcessingMode), pmStr), SkillProcessingMode)
            End If
        End Sub

        Private Sub PopulateSkillTextContent(skill As SkillDefinition, yamlObj As Newtonsoft.Json.Linq.JObject, markdownBody As String, filePath As String)
            Dim fileName = Path.GetFileNameWithoutExtension(filePath)
            If String.Equals(fileName, "GenerateOCRParseScript", StringComparison.OrdinalIgnoreCase) Then
                skill.LLMSystemPrompt = markdownBody
                System.Diagnostics.Debug.WriteLine("[SKILL DEBUG] GenerateOCRParseScript 使用完整 Markdown 正文作为 System Prompt")
            Else
                skill.LLMSystemPrompt = ExtractMarkdownSection(markdownBody, "System Prompt")
            End If
            skill.TemplateCode = ExtractMarkdownCodeBlock(markdownBody, "Template Code")

            If String.IsNullOrEmpty(skill.LLMSystemPrompt) Then
                skill.LLMSystemPrompt = GetJsonString(yamlObj, "llmSystemPrompt")
            End If
            If String.IsNullOrEmpty(skill.TemplateCode) Then
                skill.TemplateCode = GetJsonString(yamlObj, "templateCode")
            End If
        End Sub

        Private Sub ParseExcelConfig(skill As SkillDefinition, yamlObj As Newtonsoft.Json.Linq.JObject)
            If Not yamlObj.ContainsKey("excelConfig") Then Return

            Dim ecObj = yamlObj("excelConfig")
            If ecObj Is Nothing OrElse ecObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            Dim ecJObj = CType(ecObj, Newtonsoft.Json.Linq.JObject)
            skill.ExcelConfig = New ExcelConfig()
            skill.ExcelConfig.DefaultWorksheet = GetJsonString(ecJObj, "defaultWorksheet")

            ParseExcelFieldMapping(skill.ExcelConfig, ecJObj)
            ParseExcelFieldNameMapping(skill.ExcelConfig, ecJObj)
        End Sub

        Private Sub ParseExcelFieldMapping(excelConfig As ExcelConfig, ecJObj As Newtonsoft.Json.Linq.JObject)
            If Not ecJObj.ContainsKey("fieldMapping") Then Return

            Dim fmObj = ecJObj("fieldMapping")
            If fmObj Is Nothing OrElse fmObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            excelConfig.FieldMapping = New Dictionary(Of String, ExcelFieldMapping)()
            For Each prop In CType(fmObj, Newtonsoft.Json.Linq.JObject).Properties()
                Dim mapObj = prop.Value
                If mapObj IsNot Nothing AndAlso mapObj.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                    excelConfig.FieldMapping(prop.Name) = New ExcelFieldMapping With {
                        .Column = GetJsonString(CType(mapObj, Newtonsoft.Json.Linq.JObject), "column"),
                        .Header = GetJsonString(CType(mapObj, Newtonsoft.Json.Linq.JObject), "header")
                    }
                End If
            Next
        End Sub

        Private Sub ParseExcelFieldNameMapping(excelConfig As ExcelConfig, ecJObj As Newtonsoft.Json.Linq.JObject)
            If Not ecJObj.ContainsKey("fieldNameMapping") Then Return

            Dim fnmObj = ecJObj("fieldNameMapping")
            If fnmObj Is Nothing OrElse fnmObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            excelConfig.FieldNameMapping = New Dictionary(Of String, String)()
            For Each prop In CType(fnmObj, Newtonsoft.Json.Linq.JObject).Properties()
                Dim value = prop.Value?.ToString()
                If value IsNot Nothing Then
                    excelConfig.FieldNameMapping(prop.Name) = value
                End If
            Next
        End Sub

        Private Sub ParseOcrConfig(skill As SkillDefinition, yamlObj As Newtonsoft.Json.Linq.JObject)
            If Not yamlObj.ContainsKey("ocrConfig") Then Return

            Dim ocObj = yamlObj("ocrConfig")
            If ocObj Is Nothing OrElse ocObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            Dim ocJObj = CType(ocObj, Newtonsoft.Json.Linq.JObject)
            skill.OcrConfig = New OcrConfig()
            skill.OcrConfig.Engine = GetJsonString(ocJObj, "engine", "PaddleOCR")
            skill.OcrConfig.FallbackEngine = GetJsonString(ocJObj, "fallbackEngine", "PaddleOCR")

            ParseOcrPreprocessing(skill.OcrConfig, ocJObj)
            ParseOcrPostProcessing(skill.OcrConfig, ocJObj)
        End Sub

        Private Sub ParseOcrPreprocessing(ocrConfig As OcrConfig, ocJObj As Newtonsoft.Json.Linq.JObject)
            If Not ocJObj.ContainsKey("preprocessing") Then Return

            Dim preObj = ocJObj("preprocessing")
            If preObj Is Nothing OrElse preObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            Dim preJObj = CType(preObj, Newtonsoft.Json.Linq.JObject)
            ocrConfig.Preprocessing = New OcrPreprocessingConfig()
            ocrConfig.Preprocessing.Deskew = GetJsonBool(preJObj, "deskew", True)
            ocrConfig.Preprocessing.Denoise = GetJsonBool(preJObj, "denoise", True)
            ocrConfig.Preprocessing.ContrastEnhancement = GetJsonBool(preJObj, "contrastEnhancement", True)
        End Sub

        Private Sub ParseOcrPostProcessing(ocrConfig As OcrConfig, ocJObj As Newtonsoft.Json.Linq.JObject)
            If Not ocJObj.ContainsKey("postprocessing") Then Return

            Dim postObj = ocJObj("postprocessing")
            If postObj Is Nothing OrElse postObj.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Return

            Dim postJObj = CType(postObj, Newtonsoft.Json.Linq.JObject)
            ocrConfig.PostProcessing = New OcrPostProcessingConfig()
            ocrConfig.PostProcessing.UseLLM = GetJsonBool(postJObj, "useLLM", True)
            ocrConfig.PostProcessing.LLMModel = GetJsonString(postJObj, "llmModel")
            ocrConfig.PostProcessing.ExtractionPrompt = GetJsonString(postJObj, "extractionPrompt")
        End Sub

        ''' <summary>
        ''' 校验 Skill 定义的基础质量，发现问题时仅记录日志，不阻止加载
        ''' </summary>
        Private Sub ValidateSkillDefinition(skill As SkillDefinition, filePath As String)
            If skill Is Nothing Then Return

            Dim fileName = Path.GetFileName(filePath)
            Dim issues As New List(Of String)

            ' 1. VBACode/Hybrid/PythonCode 模式必须有 TemplateCode
            If (skill.ProcessingMode = SkillProcessingMode.VBACode OrElse
                skill.ProcessingMode = SkillProcessingMode.Hybrid OrElse
                skill.ProcessingMode = SkillProcessingMode.PythonCode) Then
                If String.IsNullOrWhiteSpace(skill.TemplateCode) Then
                    issues.Add($"{skill.Name}: 处理模式为 {skill.ProcessingMode}，但 TemplateCode 为空")
                ElseIf skill.ProcessingMode <> SkillProcessingMode.PythonCode Then
                    ' PythonCode 允许占位模板，其他模式检查未解析占位符
                    Dim placeholderMatches = System.Text.RegularExpressions.Regex.Matches(skill.TemplateCode, "\{\{[^}]+\}\}")
                    If placeholderMatches.Count > 0 Then
                        Dim placeholders = placeholderMatches.Cast(Of System.Text.RegularExpressions.Match)().Select(Function(m) m.Value).Distinct().Take(3)
                        issues.Add($"{skill.Name}: TemplateCode 包含未解析占位符: {String.Join(", ", placeholders)}")
                    End If
                End If
            End If

            ' 2. needsConfig=true 时 configFields 不能为空
            If skill.NeedsConfig AndAlso (skill.ConfigFields Is Nothing OrElse skill.ConfigFields.Length = 0) Then
                issues.Add($"{skill.Name}: needsConfig=true 但 configFields 为空")
            End If

            ' 3. 名称和描述不能为空
            If String.IsNullOrWhiteSpace(skill.Name) Then
                issues.Add($"{fileName}: Skill 名称为空")
            End If
            If String.IsNullOrWhiteSpace(skill.Description) Then
                issues.Add($"{skill.Name}: 描述为空")
            End If

            ' 4. 关键词/示例为空时警告
            If skill.Keywords Is Nothing OrElse skill.Keywords.Length = 0 Then
                issues.Add($"{skill.Name}: keywords 为空，可能影响意图识别")
            End If
            If skill.Examples Is Nothing OrElse skill.Examples.Length = 0 Then
                issues.Add($"{skill.Name}: examples 为空，可能影响意图识别")
            End If

            ' 输出所有问题到日志
            For Each issue In issues
                System.Diagnostics.Debug.WriteLine($"[SKILL VALIDATION] ⚠️ {issue}")
                LogToFile($"[SKILL VALIDATION] {issue}")
            Next
        End Sub

    End Class

End Namespace
