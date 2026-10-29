Imports System.Linq

Namespace Skills

    ''' <summary>
    ''' Skill 识别器 - 语义匹配版本
    ''' 根据用户输入识别匹配的 Skill
    ''' </summary>
    Public Class SkillRecognizer
        Private _skillManager As SkillManager

        ''' <summary>
        ''' 匹配置信度阈值（提升至0.6避免误匹配）
        ''' </summary>
        Public Property ConfidenceThreshold As Double = 0.6

        ''' <summary>
        ''' 多候选竞争阈值差距
        ''' </summary>
        Private ReadOnly _competitionGap As Double = 0.15

        ''' <summary>
        ''' 否定词列表（避免单字"不"造成过度误杀）
        ''' </summary>
        Private ReadOnly _negationWords As String() = {"不要", "别", "取消", "拒绝", "禁止", "无法", "无需", "不用"}

        ''' <summary>
        ''' 分类置信度阈值（已废弃 - 不再使用硬编码分类关键词）
        ''' </summary>
        Public Property CategoryConfidenceThreshold As Double = 0.6

        ''' <summary>
        ''' 构造函数
        ''' </summary>
        ''' <param name="skillManager">Skill 管理器</param>
        Public Sub New(skillManager As SkillManager)
            _skillManager = skillManager
        End Sub

        ''' <summary>
        ''' 识别用户输入匹配的 Skill（直接全量匹配，不再使用硬编码分类关键词）
        ''' </summary>
        ''' <param name="input">用户输入</param>
        ''' <param name="context">对话上下文（可选）</param>
        ''' <returns>识别结果</returns>
        Public Function Recognize(input As String, Optional context As String = "") As SkillRecognitionResult
            If String.IsNullOrWhiteSpace(input) Then
                Return New SkillRecognitionResult With {
                    .IsMatch = False,
                    .Confidence = 0,
                    .MatchDetails = "输入为空"
                }
            End If

            ' 输入长度校验（至少2字符）
            If input.Trim().Length < 2 Then
                Return New SkillRecognitionResult With {
                    .IsMatch = False,
                    .Confidence = 0,
                    .MatchDetails = "输入过短,无法识别意图"
                }
            End If

            Dim skills = _skillManager.GetAllSkills()
            If skills Is Nothing OrElse skills.Count = 0 Then
                Return New SkillRecognitionResult With {
                    .IsMatch = False,
                    .Confidence = 0,
                    .MatchDetails = "没有可用的 Skill"
                }
            End If

            ' 检测否定意图
            Dim hasNegation = DetectNegation(input)
            If hasNegation Then
                Return New SkillRecognitionResult With {
                    .IsMatch = False,
                    .Confidence = 0,
                    .MatchDetails = "检测到否定意图,不匹配任何Skill"
                }
            End If

            ' 直接全量语义匹配（不再使用硬编码分类关键词预过滤）
            Dim candidates As New List(Of SkillRecognitionResult)()
            For Each skill In skills
                Dim score = CalculateSemanticScore(input, skill, context)
                candidates.Add(New SkillRecognitionResult With {
                    .MatchedSkill = skill,
                    .Confidence = score,
                    .IsMatch = False,
                    .MatchDetails = $"语义相似度: {score:F2}"
                })
            Next

            ' 按置信度降序排序
            candidates = candidates.OrderByDescending(Function(c) c.Confidence).ToList()

            If candidates.Count = 0 OrElse candidates(0).Confidence < ConfidenceThreshold Then
                Return New SkillRecognitionResult With {
                    .IsMatch = False,
                    .Confidence = 0,
                    .MatchDetails = $"最高置信度{If(candidates.Count > 0, candidates(0).Confidence.ToString("F2"), "0.00")}低于阈值{ConfidenceThreshold}"
                }
            End If

            ' 多候选竞争机制
            Dim bestCandidate = candidates(0)
            If candidates.Count > 1 Then
                Dim gap = bestCandidate.Confidence - candidates(1).Confidence
                If gap < _competitionGap Then
                    ' 置信度差距太小,需要二次确认
                    Return New SkillRecognitionResult With {
                        .IsMatch = False,
                        .Confidence = bestCandidate.Confidence,
                        .MatchDetails = $"存在竞争候选({bestCandidate.Confidence:F2} vs {candidates(1).Confidence:F2}),需用户澄清"
                    }
                End If
            End If

            ' 标记最佳候选为匹配
            bestCandidate.IsMatch = True
            Return bestCandidate
        End Function

        ''' <summary>
        ''' 语义相似度计算
        ''' </summary>
        Private Function CalculateSemanticScore(input As String, skill As SkillDefinition, context As String) As Double
            Dim inputLower = input.ToLower().Trim()
            Dim maxScore As Double = 0

            ' 1. 名称完全匹配（最高优先级，使用整词匹配）
            If Not String.IsNullOrWhiteSpace(skill.DisplayName) AndAlso ContainsWholeWord(inputLower, skill.DisplayName.ToLower()) Then
                maxScore = Math.Max(maxScore, 0.8)
            End If

            ' 2. 关键词语义匹配（累加制，同一 Skill 多个关键词命中更有优势）
            If skill.Keywords IsNot Nothing Then
                Dim keywordScore As Double = 0
                Dim matchedKeywordCount As Integer = 0
                For Each keyword In skill.Keywords
                    If ContainsWholeWord(inputLower, keyword.ToLower()) Then
                        matchedKeywordCount += 1
                        ' 关键词越长，权重越高（精确匹配优于模糊匹配）
                        Dim weight = 0.25 + (keyword.Length * 0.015)
                        keywordScore += Math.Min(weight, 0.4)
                    End If
                Next
                ' 多关键词命中给予少量额外奖励，但总分不超过 0.65
                If matchedKeywordCount > 1 Then
                    keywordScore += 0.05 * Math.Min(matchedKeywordCount - 1, 3)
                End If
                maxScore = Math.Max(maxScore, Math.Min(keywordScore, 0.65))
            End If

            ' 3. 示例匹配（使用字符串相似度，非简单Contains）
            If skill.Examples IsNot Nothing Then
                For Each example In skill.Examples
                    Dim similarity = CalculateStringSimilarity(inputLower, example.ToLower())
                    maxScore = Math.Max(maxScore, similarity * 0.7)
                Next
            End If

            ' 4. 描述分词匹配（使用中文分词）
            If Not String.IsNullOrWhiteSpace(skill.Description) Then
                Dim descTokens = TokenizeChinese(skill.Description)
                Dim inputTokens = TokenizeChinese(input)
                Dim overlap = descTokens.Intersect(inputTokens).Count()
                If overlap > 0 Then
                    Dim descScore = overlap / Math.Max(descTokens.Count, 1)
                    maxScore = Math.Max(maxScore, descScore * 0.4)
                End If
            End If

            ' 5. 上下文增强（如果有历史对话）
            If Not String.IsNullOrWhiteSpace(context) Then
                Dim contextTokens = TokenizeChinese(context)
                Dim inputTokens = TokenizeChinese(input)
                Dim contextRelevance = contextTokens.Intersect(inputTokens).Count() / Math.Max(contextTokens.Count, 1)
                maxScore = Math.Min(maxScore + (contextRelevance * 0.1), 1.0)
            End If

            ' 6. 通用Skill轻微微降权（边界已梳理，不再大幅压制）
            If skill.IsGeneric Then
                maxScore *= 0.85
            End If

            Return Math.Min(maxScore, 1.0)
        End Function

        ''' <summary>
        ''' 检测否定意图（中文直接子串匹配，英文做词边界检查）
        ''' </summary>
        Private Function DetectNegation(input As String) As Boolean
            Dim inputLower = input.ToLower()

            ' 先排除含否定字的常见复合词
            Dim nonNegationCompounds As String() = {"不锈钢", "不容易", "不得了", "不起眼", "不耐烦", "不自觉", "不平等", "不正确", "不一定", "不好意思"}
            For Each compound In nonNegationCompounds
                If inputLower.Contains(compound) Then Return False
            Next

            For Each word In _negationWords
                If String.IsNullOrWhiteSpace(word) Then Continue For

                ' 中文否定词直接子串匹配
                Dim wordIsChinese = word.Any(Function(c) c >= "一"c AndAlso c <= "龥"c)
                If wordIsChinese Then
                    If inputLower.Contains(word) Then Return True
                    Continue For
                End If

                ' 英文否定词做词边界检查
                Dim idx = 0
                Do
                    idx = inputLower.IndexOf(word, idx)
                    If idx < 0 Then Exit Do

                    If idx > 0 Then
                        Dim prevChar = inputLower(idx - 1)
                        If IsChineseOrAlnum(prevChar) Then
                            idx += word.Length
                            Continue Do
                        End If
                    End If

                    Return True
                Loop
            Next
            Return False
        End Function

        ''' <summary>
        ''' 计算字符串相似度（基于编辑距离）
        ''' </summary>
        Private Function CalculateStringSimilarity(s1 As String, s2 As String) As Double
            If String.IsNullOrWhiteSpace(s1) OrElse String.IsNullOrWhiteSpace(s2) Then Return 0

            Dim maxLength = Math.Max(s1.Length, s2.Length)
            If maxLength = 0 Then Return 1.0

            Dim distance = CalculateLevenshteinDistance(s1, s2)
            Return 1.0 - (distance / maxLength)
        End Function

        ''' <summary>
        ''' Levenshtein编辑距离
        ''' </summary>
        Private Function CalculateLevenshteinDistance(s1 As String, s2 As String) As Integer
            Dim n = s1.Length
            Dim m = s2.Length
            Dim d(n + 1, m + 1) As Integer

            For i = 0 To n
                d(i, 0) = i
            Next
            For j = 0 To m
                d(0, j) = j
            Next

            For i = 1 To n
                For j = 1 To m
                    Dim cost = If(s1(i - 1) = s2(j - 1), 0, 1)
                    d(i, j) = Math.Min(
                        Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1),
                        d(i - 1, j - 1) + cost
                    )
                Next
            Next

            Return d(n, m)
        End Function

        ''' <summary>
        ''' 检查 input 中是否包含完整的 word
        ''' 对于中文词（包含中文字符），直接做子串匹配（中文无词边界）
        ''' 对于英文/数字词，检查词边界避免"stat"匹配"statistics"
        ''' </summary>
        Private Function ContainsWholeWord(input As String, word As String) As Boolean
            If String.IsNullOrWhiteSpace(word) Then Return False

            ' 判断 word 是否包含中文字符
            Dim wordIsChinese = word.Any(Function(c) c >= "一"c AndAlso c <= "龥"c)

            ' 中文词直接做子串匹配（中文天然无空格分词）
            If wordIsChinese Then
                Return input.Contains(word)
            End If

            ' 英文/数字词做词边界检查
            Dim idx = 0
            Do
                idx = input.IndexOf(word, idx)
                If idx < 0 Then Return False

                ' 检查前一个字符
                If idx > 0 Then
                    Dim prevChar = input(idx - 1)
                    If IsChineseOrAlnum(prevChar) Then
                        idx += word.Length
                        Continue Do
                    End If
                End If

                ' 检查后一个字符
                Dim endIdx = idx + word.Length
                If endIdx < input.Length Then
                    Dim nextChar = input(endIdx)
                    If IsChineseOrAlnum(nextChar) Then
                        idx += word.Length
                        Continue Do
                    End If
                End If

                Return True
            Loop
        End Function

        ''' <summary>
        ''' 判断字符是否为中文、字母或数字
        ''' </summary>
        Private Function IsChineseOrAlnum(c As Char) As Boolean
            Return Char.IsLetterOrDigit(c) OrElse (c >= "一"c AndAlso c <= "龥"c)
        End Function

        ''' <summary>
        ''' 简单中文分词（按常见词边界）
        ''' </summary>
        Private Function TokenizeChinese(text As String) As List(Of String)
            Dim tokens As New List(Of String)()
            ' 简单实现：按空格、标点分词，保留2字以上词
            Dim words = text.Split({" "c, "　"c, "，"c, "。"c, "！"c, "？"c}, StringSplitOptions.RemoveEmptyEntries)
            For Each word In words
                If word.Length >= 2 Then
                    tokens.Add(word.ToLower())
                End If
            Next
            Return tokens
        End Function

    End Class

End Namespace