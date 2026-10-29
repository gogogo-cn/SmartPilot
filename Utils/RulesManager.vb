Imports System.IO
Imports System.Text.RegularExpressions

Namespace Utils

    ''' <summary>
    ''' 项目规则管理器 - 类似 Trae 的 .rules 机制
    ''' 用于持久化上下文规则和快速决策
    ''' </summary>
    Public Class RulesManager
        Implements IDisposable

        Private _projectRules As Dictionary(Of String, List(Of String))
        Private _userRules As Dictionary(Of String, List(Of String))
        Private _quickResponses As Dictionary(Of String, String)
        Private _cacheLock As New Object()
        Private _rulesLoaded As Boolean = False

        Private Shared _instance As RulesManager
        Private Shared ReadOnly _instanceLock As New Object()

        Public Shared ReadOnly Property Instance As RulesManager
            Get
                If _instance Is Nothing Then
                    SyncLock _instanceLock
                        If _instance Is Nothing Then
                            _instance = New RulesManager()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        Private Sub New()
            _projectRules = New Dictionary(Of String, List(Of String))()
            _userRules = New Dictionary(Of String, List(Of String))()
            _quickResponses = New Dictionary(Of String, String)()
            LoadRules()
        End Sub

        ''' <summary>
        ''' 加载项目规则和用户规则
        ''' </summary>
        Public Sub LoadRules()
            Try
                System.Diagnostics.Debug.WriteLine("========== RulesManager 开始加载规则 ==========")

                ' 加载项目规则
                Dim projectRulesPath = GetProjectRulesPath()
                If File.Exists(projectRulesPath) Then
                    LoadRulesFile(projectRulesPath, _projectRules)
                    System.Diagnostics.Debug.WriteLine($"已加载项目规则: {projectRulesPath}")
                Else
                    System.Diagnostics.Debug.WriteLine($"项目规则文件不存在: {projectRulesPath}")
                    ' 创建默认规则文件
                    CreateDefaultRules(projectRulesPath)
                End If

                ' 加载用户规则
                Dim userRulesPath = GetUserRulesPath()
                If File.Exists(userRulesPath) Then
                    LoadRulesFile(userRulesPath, _userRules)
                    System.Diagnostics.Debug.WriteLine($"已加载用户规则: {userRulesPath}")
                End If

                ' 初始化快速响应映射
                InitializeQuickResponses()

                _rulesLoaded = True
                System.Diagnostics.Debug.WriteLine("========== RulesManager 规则加载完成 ==========")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"加载规则失败: {ex.Message}")
                ' 使用默认规则
                InitializeDefaultQuickResponses()
            End Try
        End Sub

        ''' <summary>
        ''' 获取项目规则文件路径
        ''' </summary>
        Private Function GetProjectRulesPath() As String
            ' 首先检查项目目录
            Dim appDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
            Dim projectPath = Path.Combine(appDir, ".trae", "project_rules.md")

            If File.Exists(projectPath) Then
                Return projectPath
            End If

            ' 检查开发环境路径
            Dim devPath = Path.Combine(appDir, "..", "..", ".trae", "project_rules.md")
            If File.Exists(devPath) Then
                Return devPath
            End If

            Return projectPath
        End Function

        ''' <summary>
        ''' 获取用户规则文件路径
        ''' </summary>
        Private Function GetUserRulesPath() As String
            Dim localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
            Return Path.Combine(localAppData, "ExcelAIVSTOAddIn", ".trae", "user_rules.md")
        End Function

        ''' <summary>
        ''' 加载规则文件
        ''' </summary>
        Private Sub LoadRulesFile(filePath As String, rulesDict As Dictionary(Of String, List(Of String)))
            Dim content = File.ReadAllText(filePath)
            Dim lines = content.Split(New String() {Environment.NewLine}, StringSplitOptions.None)

            Dim currentSection As String = Nothing
            Dim currentRules As New List(Of String)()

            For Each line In lines
                Dim trimmedLine = line.Trim()

                ' 检测章节标题 (## 开头)
                If trimmedLine.StartsWith("## ") Then
                    ' 保存上一个章节
                    If currentSection IsNot Nothing AndAlso currentRules.Count > 0 Then
                        rulesDict(currentSection) = New List(Of String)(currentRules)
                    End If

                    ' 开始新章节
                    currentSection = trimmedLine.Substring(3).Trim()
                    currentRules.Clear()
                ElseIf trimmedLine.StartsWith("- ") OrElse trimmedLine.StartsWith("* ") Then
                    ' 规则项
                    currentRules.Add(trimmedLine.Substring(2).Trim())
                ElseIf Not String.IsNullOrWhiteSpace(trimmedLine) AndAlso Not trimmedLine.StartsWith("#") Then
                    ' 普通文本也作为规则
                    currentRules.Add(trimmedLine)
                End If
            Next

            ' 保存最后一个章节
            If currentSection IsNot Nothing AndAlso currentRules.Count > 0 Then
                rulesDict(currentSection) = currentRules
            End If
        End Sub

        ''' <summary>
        ''' 创建默认规则文件
        ''' </summary>
        Private Sub CreateDefaultRules(filePath As String)
            Try
                Dim dir = Path.GetDirectoryName(filePath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim defaultContent = "# Excel AI 助手项目规则" & Environment.NewLine & Environment.NewLine &
                    "## 快速响应场景" & Environment.NewLine &
                    "- 问候语：你好！我是 Excel AI 助手，可以帮助你处理 Excel 相关的任务。" & Environment.NewLine &
                    "- 感谢：不客气！有其他问题随时告诉我。" & Environment.NewLine &
                    "- 告别：再见！有需要随时找我。" & Environment.NewLine &
                    "- 自定义技能：当前的自定义技能数量和列表会自动查询并展示。" & Environment.NewLine & Environment.NewLine &
                    "## Excel 操作规范" & Environment.NewLine &
                    "- 处理大量数据前先备份工作簿" & Environment.NewLine &
                    "- 使用 UsedRange 而非全表遍历" & Environment.NewLine &
                    "- 操作完成后显示确认消息给用户" & Environment.NewLine

                File.WriteAllText(filePath, defaultContent)
                System.Diagnostics.Debug.WriteLine($"已创建默认规则文件: {filePath}")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"创建默认规则文件失败: {ex.Message}")
            End Try
        End Sub

        ''' <summary>
        ''' 初始化快速响应映射
        ''' </summary>
        Private Sub InitializeQuickResponses()
            SyncLock _cacheLock
                _quickResponses.Clear()

                ' 从项目规则中提取快速响应
                If _projectRules.ContainsKey("快速响应场景") Then
                    For Each rule In _projectRules("快速响应场景")
                        ParseQuickResponseRule(rule)
                    Next
                End If

                ' 从用户规则中提取（覆盖项目规则）
                If _userRules.ContainsKey("快速响应场景") Then
                    For Each rule In _userRules("快速响应场景")
                        ParseQuickResponseRule(rule)
                    Next
                End If

                ' 如果没有加载到规则，使用默认
                If _quickResponses.Count = 0 Then
                    InitializeDefaultQuickResponses()
                End If

                System.Diagnostics.Debug.WriteLine($"已加载 {_quickResponses.Count} 个快速响应规则")
            End SyncLock
        End Sub

        ''' <summary>
        ''' 解析快速响应规则
        ''' 格式："关键词：响应内容" 或 "- 关键词：响应内容"
        ''' </summary>
        Private Sub ParseQuickResponseRule(rule As String)
            Dim match = Regex.Match(rule, "^(.*?)[：:](.+)$")
            If match.Success Then
                Dim keyword = match.Groups(1).Value.Trim()
                Dim response = match.Groups(2).Value.Trim()
                _quickResponses(keyword.ToLower()) = response
            End If
        End Sub

        ''' <summary>
        ''' 初始化默认快速响应
        ''' </summary>
        Private Sub InitializeDefaultQuickResponses()
            _quickResponses("问候语") = "你好！我是 Excel AI 助手，可以帮助你处理 Excel 相关的任务。有什么我可以帮你的吗？"
            _quickResponses("感谢") = "不客气！有其他问题随时告诉我。"
            _quickResponses("告别") = "再见！有需要随时找我。"
            _quickResponses("确认") = "好的，请继续。"
            _quickResponses("询问能力") = "我可以帮你：
1. 执行 Excel 操作（拆分工作簿、合并工作表等）
2. 生成 VBA 代码来自动化任务
3. 回答 Excel 相关问题

直接告诉我你想做什么，我会尽力帮你完成！"
        End Sub

        ''' <summary>
        ''' 检查是否匹配快速响应
        ''' </summary>
        ''' <param name="input">用户输入</param>
        ''' <returns>快速响应内容，不匹配返回 null</returns>
        Public Function GetQuickResponse(input As String) As String
            If String.IsNullOrWhiteSpace(input) Then Return Nothing

            Dim lowerInput = input.ToLower().Trim()

            ' 检查问候语
            If IsGreeting(lowerInput) Then
                Return GetResponseByKeyword("问候语")
            End If

            ' 检查感谢
            If IsThanks(lowerInput) Then
                Return GetResponseByKeyword("感谢")
            End If

            ' 检查告别
            If IsGoodbye(lowerInput) Then
                Return GetResponseByKeyword("告别")
            End If

            ' 检查确认
            If IsConfirmation(lowerInput) Then
                Return GetResponseByKeyword("确认")
            End If

            ' 检查询问能力
            If IsAbilityQuestion(lowerInput) Then
                Return GetResponseByKeyword("询问能力")
            End If

            ' 检查身份询问
            If IsIdentityQuestion(lowerInput) Then
                Return "我是 Excel AI 助手，专门帮助你处理 Excel 相关的任务。我可以执行 Excel 操作、生成 VBA 代码、回答 Excel 问题。有什么可以帮你的吗？"
            End If

            ' 检查自定义技能相关问题
            If IsCustomSkillsQuestion(lowerInput) Then
                Return GetCustomSkillsCountResponse()
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 根据关键词获取响应
        ''' </summary>
        Private Function GetResponseByKeyword(keyword As String) As String
            SyncLock _cacheLock
                If _quickResponses.ContainsKey(keyword.ToLower()) Then
                    Return _quickResponses(keyword.ToLower())
                End If
                Return Nothing
            End SyncLock
        End Function

        ''' <summary>
        ''' 判断是否为问候语
        ''' </summary>
        Private Function IsGreeting(input As String) As Boolean
            Dim greetings = New String() {"你好", "您好", "嗨", "hello", "hi", "hey", "在吗", "在么", "在嘛"}
            Return greetings.Any(Function(g) input.Contains(g)) AndAlso input.Length < 15
        End Function

        ''' <summary>
        ''' 判断是否为感谢
        ''' </summary>
        Private Function IsThanks(input As String) As Boolean
            Dim thanks = New String() {"谢谢", "感谢", "多谢", "thank", "thx"}
            Return thanks.Any(Function(t) input.Contains(t)) AndAlso input.Length < 20
        End Function

        ''' <summary>
        ''' 判断是否为告别
        ''' </summary>
        Private Function IsGoodbye(input As String) As Boolean
            Dim goodbyes = New String() {"再见", "拜拜", "bye", "goodbye", "再会"}
            Return goodbyes.Any(Function(g) input.Contains(g)) AndAlso input.Length < 15
        End Function

        ''' <summary>
        ''' 判断是否为确认
        ''' </summary>
        Private Function IsConfirmation(input As String) As Boolean
            Dim confirmations = New String() {"好的", "ok", "okay", "行", "可以", "没问题", "是的", "对的"}
            Return confirmations.Any(Function(c) input = c OrElse input = c & "。")
        End Function

        ''' <summary>
        ''' 判断是否为询问能力
        ''' </summary>
        Private Function IsAbilityQuestion(input As String) As Boolean
            Dim questions = New String() {"你能做什么", "你会什么", "有什么功能", "怎么使用", "帮助", "help"}
            Return questions.Any(Function(q) input.Contains(q))
        End Function

        ''' <summary>
        ''' 判断是否为身份询问
        ''' </summary>
        Private Function IsIdentityQuestion(input As String) As Boolean
            Dim questions = New String() {"你是谁", "你是什么", "你是什么模型", "什么模型", "什么ai", "什么人工智能", "介绍一下自己"}
            Return questions.Any(Function(q) input.Contains(q))
        End Function

        ''' <summary>
        ''' 判断是否为自定义技能相关问题（只保留"自定义"关键词，其余交给AI处理）
        ''' 技能列表已经通过系统提示暴露给AI，让AI自己统计更智能
        ''' </summary>
        Private Function IsCustomSkillsQuestion(input As String) As Boolean
            Dim keywords = New String() {"自定义技能", "自定义skill", "多少自定义"}
            Return keywords.Any(Function(k) input.Contains(k))
        End Function

        ''' <summary>
        ''' 获取自定义技能数量的响应（只回答自定义技能相关问题）
        ''' 通用技能查询交给AI处理，AI会根据系统提示中的技能列表自动统计
        ''' </summary>
        Private Function GetCustomSkillsCountResponse() As String
            Try
                Dim customManager = Skills.CustomSkillManager.GetInstance()
                Dim customSkills = customManager.GetAllCustomSkills()
                Dim count = customSkills.Count

                If count = 0 Then
                    Return "目前没有自定义技能。你可以通过 Skill 创建器来创建自己的技能！"
                Else
                    Dim skillNames = String.Join("、", customSkills.Select(Function(s) s.DisplayName))
                    Return $"当前有 {count} 个自定义技能：{skillNames}"
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"获取自定义技能数量失败: {ex.Message}")
                Return "暂时无法获取自定义技能信息，请稍后再试。"
            End Try
        End Function

        ''' <summary>
        ''' 获取规则章节内容
        ''' </summary>
        Public Function GetRulesSection(sectionName As String) As List(Of String)
            SyncLock _cacheLock
                ' 优先从用户规则获取
                If _userRules.ContainsKey(sectionName) Then
                    Return _userRules(sectionName)
                End If

                ' 从项目规则获取
                If _projectRules.ContainsKey(sectionName) Then
                    Return _projectRules(sectionName)
                End If

                Return New List(Of String)()
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有规则章节名称
        ''' </summary>
        Public Function GetAllSectionNames() As List(Of String)
            SyncLock _cacheLock
                Dim sections = New HashSet(Of String)(_projectRules.Keys)
                sections.UnionWith(_userRules.Keys)
                Return sections.ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 重新加载规则
        ''' </summary>
        Public Sub ReloadRules()
            SyncLock _cacheLock
                _projectRules.Clear()
                _userRules.Clear()
                _quickResponses.Clear()
                _rulesLoaded = False
            End SyncLock

            LoadRules()
        End Sub

        ''' <summary>
        ''' 添加用户规则
        ''' </summary>
        Public Sub AddUserRule(sectionName As String, rule As String)
            SyncLock _cacheLock
                If Not _userRules.ContainsKey(sectionName) Then
                    _userRules(sectionName) = New List(Of String)()
                End If
                _userRules(sectionName).Add(rule)
            End SyncLock

            ' 保存到文件
            SaveUserRules()
        End Sub

        ''' <summary>
        ''' 保存用户规则到文件
        ''' </summary>
        Private Sub SaveUserRules()
            Try
                Dim filePath = GetUserRulesPath()
                Dim dir = Path.GetDirectoryName(filePath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim sb As New System.Text.StringBuilder()
                sb.AppendLine("# 用户自定义规则")
                sb.AppendLine()

                For Each kvp In _userRules
                    sb.AppendLine($"## {kvp.Key}")
                    For Each rule In kvp.Value
                        sb.AppendLine($"- {rule}")
                    Next
                    sb.AppendLine()
                Next

                File.WriteAllText(filePath, sb.ToString())
                System.Diagnostics.Debug.WriteLine($"用户规则已保存: {filePath}")

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存用户规则失败: {ex.Message}")
            End Try
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' 清理资源
            _projectRules.Clear()
            _userRules.Clear()
            _quickResponses.Clear()
        End Sub

    End Class

End Namespace
