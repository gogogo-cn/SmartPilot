Imports System.Text.RegularExpressions
Imports ExcelAIVSTOAddIn.Skills

Namespace Utils

    ''' <summary>
    ''' 分层意图识别引擎 - 语义匹配版本
    ''' 第一层：本地规则匹配（最快）
    ''' 第二层：Skill 语义匹配（快）
    ''' 第三层：LLM 意图分析（慢但精准）
    ''' </summary>
    Public Class IntentRecognitionEngine

        Private _rulesManager As RulesManager
        Private _skillManager As SkillManager
        Private _skillRecognizer As SkillRecognizer
        Private _recognitionCache As Dictionary(Of String, RecognitionResult)
        Private _cacheLock As New Object()
        Private Const CacheMaxSize As Integer = 100

        ''' <summary>
        ''' 否定词列表（多字词直接匹配，单字词需排除复合词）
        ''' </summary>
        Private ReadOnly _negationWords As String() = {"不要", "别", "取消", "拒绝", "禁止", "无法", "无需", "停止", "放弃"}

        ''' <summary>
        ''' 含“不”但非否定意图的常见复合词（避免误杀）
        ''' </summary>
        Private ReadOnly _notNegationCompounds As String() = {
            "不锈钢", "不容易", "不得了", "不起眼", "不耐烦", "不自觉", "不平等", "不正确",
            "不一定", "不好意思", "不仅如此", "不知不觉", "不可思议", "不由自主", "不约而同",
            "不计其数", "不言而喻", "不劳而获", "不速之客", "不遗余力", "不择手段"
        }

        ''' <summary>
        ''' 高风险操作关键词
        ''' </summary>
        Private ReadOnly _highRiskKeywords As String() = {"删除", "清空", "发送", "提交", "覆盖", "替换", "移除", "销毁"}

        ''' <summary>
        ''' 多候选竞争阈值差距
        ''' </summary>
        Private ReadOnly _competitionGap As Double = 0.15

        Public Enum RecognitionLevel
            Level1_LocalRules      ' 本地规则匹配
            Level2_SkillKeywords   ' Skill 语义匹配
            Level3_LLM_Analysis    ' LLM 意图分析
        End Enum

        Public Class RecognitionResult
            Public Property Level As RecognitionLevel
            Public Property Intent As String
            Public Property Confidence As Double
            Public Property SkillName As String
            Public Property Response As String
            Public Property ShouldCallLLM As Boolean
            Public Property Source As String
            Public Property ProcessingTime As TimeSpan
            Public Property MatchedKeyword As String
            Public Property MatchedKeywords As List(Of String) ' 新增：返回所有匹配关键词
            Public Property NeedsConfirmation As Boolean ' 新增：是否需要二次确认
        End Class

        Public Sub New()
            _rulesManager = RulesManager.Instance
            _skillManager = SkillManager.Instance
            _skillRecognizer = New SkillRecognizer(_skillManager)
            _recognitionCache = New Dictionary(Of String, RecognitionResult)()
        End Sub

        ''' <summary>
        ''' 分层意图识别主入口
        ''' </summary>
        ''' <param name="input">用户输入</param>
        ''' <param name="hasAttachment">是否有附件</param>
        ''' <param name="context">对话上下文（可选）</param>
        ''' <returns>识别结果</returns>
        Public Function RecognizeIntent(input As String, Optional hasAttachment As Boolean = False, Optional context As String = "") As RecognitionResult
            Dim stopwatch = System.Diagnostics.Stopwatch.StartNew()

            Try
                ' 输入长度校验
                If String.IsNullOrWhiteSpace(input) Then
                    Return CreateResult(RecognitionLevel.Level3_LLM_Analysis, "input_empty", 0.0, True, "输入为空")
                End If

                If input.Trim().Length < 2 Then
                    Return CreateResult(RecognitionLevel.Level3_LLM_Analysis, "input_too_short", 0.0, True, "输入过短,无法识别意图")
                End If

                ' 检测否定意图
                If DetectNegation(input) Then
                    Return CreateResult(RecognitionLevel.Level3_LLM_Analysis, "negation_detected", 0.0, True, "检测到否定意图,不匹配任何Skill")
                End If

                ' 检查缓存（包含上下文）
                Dim cacheKey = GetCacheKey(input, hasAttachment, context)
                SyncLock _cacheLock
                    If _recognitionCache.ContainsKey(cacheKey) Then
                        Dim cached = _recognitionCache(cacheKey)
                        cached.ProcessingTime = stopwatch.Elapsed
                        Return cached
                    End If
                End SyncLock

                Dim result As RecognitionResult

                ' ===== 第一层：本地规则匹配（最快，< 1ms）=====
                result = TryLocalRulesMatch(input)
                If result IsNot Nothing AndAlso result.Confidence >= 0.9 Then
                    stopwatch.Stop()
                    result.ProcessingTime = stopwatch.Elapsed
                    AddToCache(cacheKey, result)
                    Return result
                End If

                ' ===== 第二层：Skill 语义匹配（快，< 10ms）=====
                result = TrySkillSemanticMatch(input, context)
                If result IsNot Nothing AndAlso result.Confidence >= 0.6 Then
                    ' 高风险操作需二次确认
                    If IsHighRiskOperation(input) Then
                        result.NeedsConfirmation = True
                        result.Confidence = Math.Min(result.Confidence, 0.85) ' 降低置信度,强制确认
                    End If

                    stopwatch.Stop()
                    result.ProcessingTime = stopwatch.Elapsed
                    AddToCache(cacheKey, result)
                    Return result
                End If

                ' ===== 第三层：需要 LLM 分析（慢，> 500ms）=====
                result = CreateResult(RecognitionLevel.Level3_LLM_Analysis, "llm_required", 0.0, True, "需要LLM分析")

                stopwatch.Stop()
                result.ProcessingTime = stopwatch.Elapsed
                AddToCache(cacheKey, result)
                Return result

            Catch ex As Exception
                stopwatch.Stop()
                System.Diagnostics.Debug.WriteLine($"意图识别失败: {ex.Message}")

                ' 优雅降级，不直接调用LLM
                Return CreateResult(RecognitionLevel.Level3_LLM_Analysis, "error_fallback", 0.0, True, $"识别异常:{ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' 第一层：本地规则匹配
        ''' </summary>
        Private Function TryLocalRulesMatch(input As String) As RecognitionResult
            ' 使用 RulesManager 检查快速响应
            Dim quickResponse = _rulesManager.GetQuickResponse(input)

            If quickResponse IsNot Nothing Then
                Return New RecognitionResult With {
                    .Level = RecognitionLevel.Level1_LocalRules,
                    .Intent = "quick_response",
                    .Confidence = 1.0,
                    .ShouldCallLLM = False,
                    .Source = "local_rules",
                    .Response = quickResponse
                }
            End If

            ' 检查翻译请求（本地规则）
            If IsTranslateRequest(input) Then
                Return New RecognitionResult With {
                    .Level = RecognitionLevel.Level1_LocalRules,
                    .Intent = "translate",
                    .Confidence = 0.95,
                    .ShouldCallLLM = False,
                    .Source = "local_rules",
                    .Response = Nothing
                }
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 第二层：Skill 语义匹配（使用SkillRecognizer）
        ''' </summary>
        Private Function TrySkillSemanticMatch(input As String, context As String) As RecognitionResult
            Dim skillResult = _skillRecognizer.Recognize(input, context)

            If skillResult.IsMatch AndAlso skillResult.MatchedSkill IsNot Nothing Then
                Dim skill = skillResult.MatchedSkill
                Return New RecognitionResult With {
                    .Level = RecognitionLevel.Level2_SkillKeywords,
                    .Intent = "execute_skill",
                    .Confidence = skillResult.Confidence,
                    .SkillName = skill.Name,
                    .ShouldCallLLM = False,
                    .Source = "skill_semantic",
                    .MatchedKeyword = skillResult.MatchDetails,
                    .MatchedKeywords = GetAllMatchedKeywords(input, skill),
                    .Response = Nothing
                }
            End If

            Return Nothing
        End Function

        ''' <summary>
        ''' 计算 Skill 匹配分数（优化版）
        ''' </summary>
        Private Function CalculateSkillMatchScore(input As String, skill As SkillDefinition) As Double
            If String.IsNullOrWhiteSpace(input) OrElse skill Is Nothing Then
                Return 0
            End If

            Dim score As Double = 0
            Dim lowerInput = input.ToLower().Trim()

            ' 0. 【新增】描述完全匹配（最高优先级）
            If Not String.IsNullOrWhiteSpace(skill.Description) Then
                Dim lowerDesc = skill.Description.ToLower().Trim()
                
                ' 描述与输入完全相同或高度相似
                If lowerDesc = lowerInput Then
                    score += 1.0  ' 完全匹配，直接满分
                    Return Math.Min(score, 1.0)
                ElseIf lowerDesc.Contains(lowerInput) OrElse lowerInput.Contains(lowerDesc) Then
                    score += 0.8  ' 互相包含，高度匹配
                End If
            End If

            ' 1. 关键词匹配（权重最高）
            If skill.Keywords IsNot Nothing Then
                For Each keyword In skill.Keywords
                    If String.IsNullOrWhiteSpace(keyword) Then Continue For

                    Dim lowerKeyword = keyword.ToLower()

                    ' 完全匹配加分更多
                    If input = lowerKeyword Then
                        score += 0.4
                    ElseIf input.Contains(lowerKeyword) Then
                        score += 0.25
                    End If

                    ' 关键词在开头加分
                    If input.StartsWith(lowerKeyword) Then
                        score += 0.15
                    End If
                Next
            End If

            ' 2. 示例匹配
            If skill.Examples IsNot Nothing Then
                For Each example In skill.Examples
                    If String.IsNullOrWhiteSpace(example) Then Continue For

                    Dim lowerExample = example.ToLower()
                    
                    ' 双向匹配：输入包含示例 OR 示例包含输入
                    If input.Contains(lowerExample) OrElse lowerExample.Contains(input) Then
                        score += 0.3
                    End If
                Next
            End If

            ' 3. 名称匹配
            If Not String.IsNullOrWhiteSpace(skill.DisplayName) Then
                If input.Contains(skill.DisplayName.ToLower()) OrElse skill.DisplayName.ToLower().Contains(input) Then
                    score += 0.2
                End If
            End If

            ' 4. 通用 Skill 降权
            If skill.IsGeneric Then
                score *= 0.7
            End If

            ' 5. 高使用频率 Skill 加权
            If skill.UsageCount > 10 Then
                score *= 1.1
            End If

            Return Math.Min(score, 1.0)
        End Function

        ''' <summary>
        ''' 获取所有匹配的关键词（用于调试）
        ''' </summary>
        Private Function GetAllMatchedKeywords(input As String, skill As SkillDefinition) As List(Of String)
            Dim matched As New List(Of String)()
            If skill.Keywords Is Nothing Then Return matched

            Dim lowerInput = input.ToLower()
            For Each keyword In skill.Keywords
                If Not String.IsNullOrWhiteSpace(keyword) AndAlso lowerInput.Contains(keyword.ToLower()) Then
                    matched.Add(keyword)
                End If
            Next

            Return matched
        End Function

        ''' <summary>
        ''' 判断是否为翻译请求
        ''' </summary>
        Private Function IsTranslateRequest(input As String) As Boolean
            Dim translatePatterns = New String() {"翻译", "translate", "译文", "翻译成"}
            Dim selectedPatterns = New String() {"选中", "选择", "选", "cell", "单元格", "内容", "text"}

            Dim hasTranslate = translatePatterns.Any(Function(p) input.Contains(p))
            Dim hasSelected = selectedPatterns.Any(Function(p) input.Contains(p))

            Return hasTranslate AndAlso hasSelected
        End Function

        ''' <summary>
        ''' 检测否定意图（带复合词排除，避免"不锈钢"等误杀）
        ''' </summary>
        Private Function DetectNegation(input As String) As Boolean
            Dim lowerInput = input.ToLower()

            ' 先检查是否为含否定字的复合词（非否定意图）
            For Each compound In _notNegationCompounds
                If lowerInput.Contains(compound) Then
                    Return False
                End If
            Next

            ' 检查多字否定词（直接包含即命中）
            For Each word In _negationWords
                If lowerInput.Contains(word) Then
                    Return True
                End If
            Next

            ' 单字"不"需要更精细判断：后面跟形容词/动词才视为否定
            ' 简单策略：检查"不"后面一个字是否为常见否定搭配字
            Dim buIndex = lowerInput.IndexOf("不"c)
            If buIndex >= 0 Then
                ' 检查是否为复合词的一部分（前后都是中文字符且构成常见词）
                Dim isCompound = False
                If buIndex > 0 AndAlso buIndex < lowerInput.Length - 1 Then
                    Dim prevChar = lowerInput(buIndex - 1)
                    Dim nextChar = lowerInput(buIndex + 1)
                    ' 前后都是中文，大概率是复合词的一部分
                    If (prevChar >= "一"c AndAlso prevChar <= "龥"c) AndAlso
                       (nextChar >= "一"c AndAlso nextChar <= "龥"c) Then
                        isCompound = True
                    End If
                End If
                If Not isCompound Then
                    Return True
                End If
            End If

            Return False
        End Function

        ''' <summary>
        ''' 检测高风险操作
        ''' </summary>
        Private Function IsHighRiskOperation(input As String) As Boolean
            Dim lowerInput = input.ToLower()
            For Each keyword In _highRiskKeywords
                If lowerInput.Contains(keyword) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 创建识别结果
        ''' </summary>
        Private Function CreateResult(level As RecognitionLevel, intent As String, confidence As Double, shouldCallLLM As Boolean, source As String) As RecognitionResult
            Return New RecognitionResult With {
                .Level = level,
                .Intent = intent,
                .Confidence = confidence,
                .ShouldCallLLM = shouldCallLLM,
                .Source = source,
                .MatchedKeywords = New List(Of String)(),
                .NeedsConfirmation = False
            }
        End Function

        ''' <summary>
        ''' 生成缓存键（包含上下文）
        ''' </summary>
        Private Function GetCacheKey(input As String, hasAttachment As Boolean, context As String) As String
            Dim contextHash = If(String.IsNullOrWhiteSpace(context), "", System.Security.Cryptography.MD5.Create().ComputeHash(System.Text.Encoding.UTF8.GetBytes(context)).Aggregate("", Function(s, b) s + b.ToString("x2")))
            Return $"{input.ToLower().Trim()}|{hasAttachment}|{contextHash}"
        End Function

        ''' <summary>
        ''' 添加到缓存
        ''' </summary>
        Private Sub AddToCache(key As String, result As RecognitionResult)
            SyncLock _cacheLock
                ' 清理旧缓存
                If _recognitionCache.Count >= CacheMaxSize Then
                    Dim oldestKey = _recognitionCache.Keys.First()
                    _recognitionCache.Remove(oldestKey)
                End If

                _recognitionCache(key) = result
            End SyncLock
        End Sub

        ''' <summary>
        ''' 清除缓存
        ''' </summary>
        Public Sub ClearCache()
            SyncLock _cacheLock
                _recognitionCache.Clear()
            End SyncLock
        End Sub

        ''' <summary>
        ''' 获取缓存统计
        ''' </summary>
        Public Function GetCacheStats() As Dictionary(Of String, Object)
            SyncLock _cacheLock
                Return New Dictionary(Of String, Object) From {
                    {"cacheSize", _recognitionCache.Count},
                    {"maxSize", CacheMaxSize},
                    {"hitRate", "N/A"} ' 可以扩展命中率统计
                }
            End SyncLock
        End Function

#Region "RAG 集成支持"

        ''' <summary>
        ''' 使用 RAG 增强识别结果（异步）
        ''' 当意图可能是知识问答时，调用此方法进行 RAG 检索增强
        ''' 如果检索成功，返回增强后的上下文；如果失败或无结果，返回降级标志
        '''
        ''' 使用示例:
        ''' Dim ragResult = Await intentEngine.EnhanceWithRagAsync(userQuestion)
        ''' If ragResult.ShouldFallback Then
        '''     ' 降级到普通 chat 流程
        ''' Else
        '''     ' 使用 ragResult.ContextText 作为增强上下文
        ''' End If
        ''' </summary>
        ''' <param name="userQuestion">用户原始问题</param>
        ''' <param name="optionalKbId">可选的知识库 ID（不传则使用默认）</param>
        ''' <returns>RagResult 对象</returns>
        Public Async Function EnhanceWithRagAsync(userQuestion As String,
                                                     Optional optionalKbId As String = "") As System.Threading.Tasks.Task(Of Services.RAG.RagResult)
            Try
                ' 使用单例编排器（默认知识库从 RagConfig.DefaultKnowledgeBaseId 读取）
                Dim orchestrator As Services.RAG.RagOrchestrator = Services.RAG.RagOrchestrator.Instance
                Dim result = Await orchestrator.RetrieveAndEnhanceAsync(userQuestion, optionalKbId).ConfigureAwait(False)

                If result IsNot Nothing Then
                    If result.HasContext Then
                        System.Diagnostics.Debug.WriteLine($"[IntentEngine] RAG 增强: 检索到 {result.RetrievedChunks} 个相关块")
                    Else
                        System.Diagnostics.Debug.WriteLine("[IntentEngine] RAG 检索为空，将降级到 chat")
                    End If
                End If

                Return result

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[IntentEngine] EnhanceWithRagAsync 异常: {ex.Message}")

                ' 返回降级结果
                Return New Services.RAG.RagResult With {
                    .HasContext = False,
                    .ContextText = "",
                    .SourceReferences = New List(Of Services.RAG.SourceRef)(),
                    .RetrievedChunks = 0,
                    .ShouldFallback = True
                }
            End Try
        End Function

#End Region

    End Class

End Namespace
