Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Diagnostics
Imports ExcelAIVSTOAddIn.Services.Parsers

Namespace Utils

    ''' <summary>
    ''' 文本块数据类
    ''' </summary>
    Public Class TextChunk
        Public Property Id As Integer
        Public Property Text As String
        Public Property CharCount As Integer
        Public Property StartPosition As Long          ' 在原文中的起始位置
        Public Property EndPosition As Long            ' 在原文中的结束位置
        Public Property SourceSection As String         ' 来源章节（如有）
        Public Property Level As Integer                ' 层级深度
    End Class

    ''' <summary>
    ''' 语义分块器 - 将长文档智能切分为语义完整的块
    ''' 单例模式（Lazy + SyncLock 线程安全）
    ''' </summary>
    Public Class SemanticChunker

        ' ==================== 配置常量 ====================
        Private _defaultChunkSize As Integer = 1500      ' 默认每块目标字符数
        Private _defaultOverlap As Integer = 200          ' 块间重叠字符数
        Private _maxChunkSize As Integer = 4000           ' 单块上限
        Private _minChunkSize As Integer = 200            ' 单块下限
        Private _maxRecursionDepth As Integer = 6         ' 最大递归深度（防止指数级递归）
        Private _maxTotalChunks As Integer = 5000         ' 最大块数（防止内存爆炸）

        ' 分隔符优先级列表（从高到低）
        Private Shared ReadOnly Separators As String() = {
            vbCrLf & vbCrLf,   ' 双换行（段落分隔）
            vbCrLf,            ' 单换行
            "。",               ' 中文句号
            ". ",              ' 英文句号+空格
            "；",              ' 中文分号
            "; ",              ' 英文分号
            "，",              ' 中文逗号
            ", ",              ' 英文逗号
            " "                ' 空格（最后手段）
        }

        ' 停用词列表
        Private Shared ReadOnly StopWords As HashSet(Of String) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "的", "了", "是", "在", "这", "那", "有", "和", "与", "或",
            "什么", "如何", "为什么", "哪", "哪个", "哪些", "怎样",
            "the", "a", "an", "is", "are", "was", "were", "be", "been",
            "being", "have", "has", "had", "do", "does", "did", "will",
            "would", "could", "should", "may", "might", "shall", "can",
            "to", "of", "in", "for", "on", "with", "at", "by", "from",
            "as", "into", "through", "during", "before", "after", "above",
            "below", "between", "out", "off", "over", "under", "again",
            "further", "then", "once", "here", "there", "when", "where",
            "why", "how", "all", "each", "few", "more", "most", "other",
            "some", "such", "no", "nor", "not", "only", "own", "same",
            "so", "than", "too", "very", "just", "because", "but", "and",
            "if", "or", "until", "while", "about", "against", "between"
        }

        ' ==================== 单例模式 ====================
        Private Shared _instance As Lazy(Of SemanticChunker) =
            New Lazy(Of SemanticChunker)(Function() New SemanticChunker())

        Public Shared ReadOnly Property Instance As SemanticChunker
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
            Debug.WriteLine("[SemanticChunker] 初始化完成")
        End Sub

        ' ==================== 主入口方法 ====================

        ''' <summary>
        ''' 主入口方法 - 根据模式自动选择分块策略
        ''' </summary>
        ''' <param name="text">待分块的文本</param>
        ''' <param name="toc">目录结构（可选）</param>
        ''' <param name="mode">分块模式：structure/paragraph/fixed/auto</param>
        ''' <returns>文本块列表</returns>
        Public Function Chunk(text As String,
                              Optional toc As List(Of HeadingNode) = Nothing,
                              Optional mode As String = "auto") As List(Of TextChunk)

            If String.IsNullOrWhiteSpace(text) Then
                Debug.WriteLine("[SemanticChunker] 输入文本为空")
                Return New List(Of TextChunk)()
            End If

            Debug.WriteLine(String.Format("[SemanticChunker] 开始分块: 模式={0}, 文本长度={1}", mode, text.Length))

            Dim result As List(Of TextChunk)
            Select Case mode.ToLower()
                Case "structure"
                    If toc IsNot Nothing AndAlso toc.Count > 0 Then
                        result = ChunkByStructure(text, toc)
                    Else
                        Debug.WriteLine("[SemanticChunker] structure模式但无目录，降级为paragraph模式")
                        result = ChunkByParagraph(text)
                    End If
                Case "paragraph"
                    result = ChunkByParagraph(text)
                Case "fixed"
                    result = ChunkByFixedLength(text)
                Case "auto"
                    result = AutoSelectStrategy(text, toc)
                Case Else
                    Debug.WriteLine(String.Format("[SemanticChunker] 未知模式 '{0}'，使用默认paragraph模式", mode))
                    result = ChunkByParagraph(text)
            End Select

            ' 安全限制：块数过多时截断
            If result IsNot Nothing AndAlso result.Count > _maxTotalChunks Then
                Debug.WriteLine(String.Format("[SemanticChunker] 块数({0})超限，截断为{1}", result.Count, _maxTotalChunks))
                result = result.GetRange(0, _maxTotalChunks)
            End If

            Return result
        End Function

        ' ==================== 策略选择 ====================

        ''' <summary>
        ''' 自动判断最佳分块策略
        ''' </summary>
        Private Function AutoSelectStrategy(text As String,
                                            Optional toc As List(Of HeadingNode) = Nothing) As List(Of TextChunk)

            ' 有目录且章节较多 → 使用结构化分块
            If toc IsNot Nothing AndAlso toc.Count > 0 Then
                Debug.WriteLine("[SemanticChunker] auto模式: 选择structure策略（有目录）")
                Return ChunkByStructure(text, toc)
            End If

            ' 文本包含明显的段落标记（双换行）→ 使用段落分块
            Dim doubleNewLineCount As Integer = CountOccurrences(text, vbCrLf & vbCrLf)
            If doubleNewLineCount > 3 Then
                Debug.WriteLine("[SemanticChunker] auto模式: 选择paragraph策略（检测到多段落）")
                Return ChunkByParagraph(text)
            End If

            ' 默认使用段落分块
            Debug.WriteLine("[SemanticChunker] auto模式: 选择默认paragraph策略")
            Return ChunkByParagraph(text)
        End Function

        ''' <summary>
        ''' 统计子串出现次数
        ''' </summary>
        Private Function CountOccurrences(text As String, substr As String) As Integer
            If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(substr) Then
                Return 0
            End If

            Dim count As Integer = 0
            Dim pos As Integer = 0
            Do
                pos = text.IndexOf(substr, pos, StringComparison.Ordinal)
                If pos < 0 Then Exit Do
                count += 1
                pos += substr.Length
            Loop
            Return count
        End Function

        ' ==================== 结构化分块 ====================

        ''' <summary>
        ''' 按文档结构（标题树）分块
        ''' 每个HeadingNode对应一个chunk，过长章节递归细分
        ''' </summary>
        Public Function ChunkByStructure(text As String, toc As List(Of HeadingNode)) As List(Of TextChunk)
            Dim chunks As New List(Of TextChunk)()
            Dim chunkId As Integer = 0

            Debug.WriteLine(String.Format("[SemanticChunker] Structure分块开始: {0}个顶层节点", toc.Count))

            ' 递归处理每个标题节点
            ProcessHeadingNodes(text, toc, chunks, chunkId, "")

            Debug.WriteLine(String.Format("[SemanticChunker] Structure分块完成: 共{0}个块", chunks.Count))
            Return chunks
        End Function

        ''' <summary>
        ''' 递归处理标题节点树
        ''' </summary>
        Private Sub ProcessHeadingNodes(text As String,
                                        nodes As List(Of HeadingNode),
                                        ByRef chunks As List(Of TextChunk),
                                        ByRef chunkId As Integer,
                                        parentPath As String)

            For Each node As HeadingNode In nodes
                ' 构建章节路径
                Dim sectionPath As String = If(parentPath.Length > 0,
                                               parentPath & "/" & node.Title,
                                               node.Title)

                Debug.WriteLine(String.Format("[SemanticChunker] 处理节点: Level={0}, Title={1}", node.Level, node.Title))

                ' 提取该节点的文本内容
                Dim sectionText As String = ExtractSectionText(text, node)

                If String.IsNullOrWhiteSpace(sectionText) Then
                    Debug.WriteLine(String.Format("[SemanticChunker] 节点 '{0}' 文本为空，跳过", node.Title))
                    Continue For
                End If

                ' 如果文本超过上限，需要进一步分割
                If sectionText.Length > _maxChunkSize Then
                    Debug.WriteLine(String.Format("[SemanticChunker] 节点 '{0}' 文本过长({1}字符)，递归细分",
                                                  node.Title, sectionText.Length))

                    Dim subChunks As List(Of TextChunk) = ChunkByParagraph(sectionText)
                    For Each subChunk As TextChunk In subChunks
                        subChunk.Id = System.Threading.Interlocked.Increment(chunkId)
                        subChunk.SourceSection = sectionPath
                        subChunk.Level = node.Level
                        ' 调整绝对位置
                        subChunk.StartPosition += node.StartPosition
                        subChunk.EndPosition += node.StartPosition
                        chunks.Add(subChunk)
                    Next
                Else
                    ' 直接作为一个块
                    Dim chunk As New TextChunk With {
                        .Id = System.Threading.Interlocked.Increment(chunkId),
                        .Text = sectionText.Trim(),
                        .CharCount = sectionText.Trim().Length,
                        .StartPosition = node.StartPosition,
                        .EndPosition = node.EndPosition,
                        .SourceSection = sectionPath,
                        .Level = node.Level
                    }
                    chunks.Add(chunk)
                End If

                ' 递归处理子节点
                If node.Children IsNot Nothing AndAlso node.Children.Count > 0 Then
                    ProcessHeadingNodes(text, node.Children, chunks, chunkId, sectionPath)
                End If
            Next
        End Sub

        ''' <summary>
        ''' 从原文中提取指定节点的文本内容
        ''' </summary>
        Private Function ExtractSectionText(text As String, node As HeadingNode) As String
            If node.EndPosition <= node.StartPosition OrElse
               node.StartPosition < 0 OrElse
               node.EndPosition > text.Length Then

                ' 回退：如果位置信息无效，尝试按标题查找
                Dim startIdx As Integer = text.IndexOf(node.Title, StringComparison.Ordinal)
                If startIdx >= 0 Then
                    ' 查找下一个同级或更高级标题作为结束位置
                    Dim endIdx As Integer = FindNextSiblingOrParentTitle(text, startIdx, node.Level)
                    If endIdx > startIdx Then
                        Return text.Substring(startIdx, endIdx - startIdx)
                    ElseIf endIdx = -1 Then
                        Return text.Substring(startIdx)
                    End If
                End If
                Return ""
            End If

            Return text.Substring(CInt(node.StartPosition), CInt(node.EndPosition - node.StartPosition))
        End Function

        ''' <summary>
        ''' 查找下一个同级或更高级标题的位置
        ''' </summary>
        Private Function FindNextSiblingOrParentTitle(text As String,
                                                       startPos As Integer,
                                                       currentLevel As Integer) As Integer

            ' 简单实现：搜索常见标题模式
            Dim patterns As String() = {
                vbCrLf & "第" & "\d+" & "[章节]",
                vbCrLf & "\d+(\.\d+)*[.、]",
                vbCrLf & "#{1,3}\s"
            }

            Dim searchStart As Integer = startPos + 1
            Dim minPos As Integer = text.Length

            For Each pattern As String In patterns
                Try
                    Dim regex As New System.Text.RegularExpressions.Regex(pattern,
                        System.Text.RegularExpressions.RegexOptions.None)
                    Dim match As System.Text.RegularExpressions.Match = regex.Match(text, searchStart)
                    If match.Success AndAlso match.Index < minPos Then
                        minPos = match.Index
                    End If
                Catch ex As Exception
                    ' 忽略正则错误
                End Try
            Next

            If minPos < text.Length Then
                Return minPos
            End If

            Return -1
        End Function

        ' ==================== 段落分块（核心算法）====================

        ''' <summary>
        ''' 按段落分块 - 核心递归分割算法
        ''' 使用多级分隔符优先级列表进行智能切割
        ''' </summary>
        Public Function ChunkByParagraph(text As String,
                                          Optional separatorIndex As Integer = 0,
                                          Optional chunkSize As Integer = -1,
                                          Optional overlap As Integer = -1,
                                          Optional depth As Integer = 0) As List(Of TextChunk)

            If chunkSize < 0 Then chunkSize = _defaultChunkSize
            If overlap < 0 Then overlap = _defaultOverlap

            ' 安全检查：递归深度超限 → 降级为固定长度分块
            If depth > _maxRecursionDepth Then
                Debug.WriteLine(String.Format("[SemanticChunker] 递归深度超限({0})，降级为固定长度分块", depth))
                Return ChunkByFixedLength(text, chunkSize, overlap)
            End If

            Dim chunks As New List(Of TextChunk)()
            Dim chunkId As Integer = 0
            Dim currentChunk As New StringBuilder()
            Dim globalPos As Long = 0           ' 在原文中的全局位置
            Dim chunkStartPosition As Long = 0  ' 当前块的起始位置

            Debug.WriteLine(String.Format("[SemanticChunker] Paragraph分块开始: 分隔符级别={0}, 目标大小={1}",
                                          separatorIndex, chunkSize))

            ' 获取当前级别的分隔符
            Dim separator As String = Separators(separatorIndex)
            Dim segments As String() = SplitWithDelimiter(text, separator)

            Debug.WriteLine(String.Format("[SemanticChunker] 使用分隔符 '{0}', 切割为{1}段",
                                          If(separator.Length > 5, separator.Substring(0, 5) & "...", separator),
                                          segments.Count))

            For i As Integer = 0 To segments.Length - 1
                Dim segment As String = segments(i)

                ' 跳过空白段
                If String.IsNullOrWhiteSpace(segment) Then
                    globalPos += segment.Length
                    Continue For
                End If

                ' 检查单段是否超长
                If segment.Length > chunkSize AndAlso separatorIndex < Separators.Length - 1 Then
                    ' 先保存当前累积的内容（如果有）
                    If currentChunk.Length > 0 Then
                        SaveChunk(chunks, currentChunk, chunkId, chunkStartPosition, globalPos, "", 0)
                        chunkId = chunks.Count
                        currentChunk.Clear()
                    End If

                    ' 递归使用下一级分隔符继续切割
                    Debug.WriteLine(String.Format(
                        "[SemanticChunker] 段落过长({0}字符)，递归降级到分隔符级别{1}",
                        segment.Length, separatorIndex + 1))

                    Dim subChunks As List(Of TextChunk) = ChunkByParagraph(segment, separatorIndex + 1, chunkSize, overlap, depth + 1)
                    For Each subChunk As TextChunk In subChunks
                        subChunk.Id = System.Threading.Interlocked.Increment(chunkId)
                        subChunk.StartPosition += globalPos
                        subChunk.EndPosition += globalPos
                        chunks.Add(subChunk)
                    Next

                    globalPos += segment.Length
                    chunkStartPosition = globalPos
                    Continue For
                End If

                ' 累积段落到当前块
                If currentChunk.Length > 0 Then
                    currentChunk.Append(separator)
                End If
                currentChunk.Append(segment)

                ' 检查是否达到目标大小
                If currentChunk.Length >= chunkSize Then
                    Dim savedLength As Integer = SaveChunk(chunks, currentChunk, chunkId, chunkStartPosition, globalPos + segment.Length, "", 0)
                    chunkId = chunks.Count

                    ' 计算重叠部分
                    Dim overlapText As String = ""
                    If overlap > 0 AndAlso savedLength > overlap Then
                        overlapText = currentChunk.ToString().Substring(savedLength - overlap, overlap)
                    End If

                    ' 新块以重叠内容开始
                    currentChunk.Clear()
                    If overlapText.Length > 0 Then
                        currentChunk.Append(overlapText)
                    End If
                    chunkStartPosition = globalPos + segment.Length - overlapText.Length
                End If

                globalPos += segment.Length
            Next

            ' 保存剩余内容
            If currentChunk.Length >= _minChunkSize Then
                SaveChunk(chunks, currentChunk, chunkId, chunkStartPosition, globalPos, "", 0)
            ElseIf currentChunk.Length > 0 AndAlso chunks.Count > 0 Then
                ' 内容太少，追加到最后一个块
                Dim lastChunk As TextChunk = chunks(chunks.Count - 1)
                lastChunk.Text &= currentChunk.ToString()
                lastChunk.CharCount = lastChunk.Text.Length
                lastChunk.EndPosition = globalPos
            End If

            ' 最终检查：确保没有块超过上限
            EnforceMaxChunkSize(chunks, chunkSize, overlap)

            Debug.WriteLine(String.Format("[SemanticChunker] Paragraph分块完成: 共{0}个块", chunks.Count))
            Return chunks
        End Function

        ''' <summary>
        ''' 使用指定分隔符切割文本，保留分隔符信息
        ''' </summary>
        Private Function SplitWithDelimiter(text As String, delimiter As String) As String()
            If String.IsNullOrEmpty(delimiter) Then
                Return {text}
            End If

            Dim result As New List(Of String)()
            Dim parts As String() = text.Split({delimiter}, StringSplitOptions.None)

            For i As Integer = 0 To parts.Length - 1
                result.Add(parts(i))
                ' 不在最后一段后添加分隔符
            Next

            Return result.ToArray()
        End Function

        ''' <summary>
        ''' 保存当前块到结果列表
        ''' </summary>
        Private Function SaveChunk(ByRef chunks As List(Of TextChunk),
                                    ByRef currentChunk As StringBuilder,
                                    ByRef chunkId As Integer,
                                    startPosition As Long,
                                    endPosition As Long,
                                    sourceSection As String,
                                    level As Integer) As Integer

            Dim chunkText As String = currentChunk.ToString().Trim()
            If chunkText.Length < _minChunkSize Then
                Return 0
            End If

            Dim chunk As New TextChunk With {
                .Id = System.Threading.Interlocked.Increment(chunkId),
                .Text = chunkText,
                .CharCount = chunkText.Length,
                .StartPosition = startPosition,
                .EndPosition = endPosition,
                .SourceSection = sourceSection,
                .Level = level
            }
            chunks.Add(chunk)

            Debug.WriteLine(String.Format("[SemanticChunker] 保存块 #{0}: {1}字符, 位置[{2}-{3}]",
                                          chunk.Id, chunk.CharCount, startPosition, endPosition))

            Return chunkText.Length
        End Function

        ''' <summary>
        ''' 强制执行最大块大小限制
        ''' </summary>
        Private Sub EnforceMaxChunkSize(ByRef chunks As List(Of TextChunk),
                                         chunkSize As Integer,
                                         overlap As Integer)

            Dim i As Integer = 0
            While i < chunks.Count
                If chunks(i).CharCount > _maxChunkSize Then
                    Debug.WriteLine(String.Format("[SemanticChunker] 块#{0}超限({1}字符)，强制分割",
                                                  chunks(i).Id, chunks(i).CharCount))

                    ' 强制按固定大小分割
                    Dim oversizedChunk As TextChunk = chunks(i)
                    chunks.RemoveAt(i)

                    Dim subChunks As List(Of TextChunk) = ForceSplitChunk(oversizedChunk, _maxChunkSize, overlap)
                    For Each subChunk As TextChunk In subChunks
                        chunks.Insert(i, subChunk)
                        i += 1
                    Next
                Else
                    i += 1
                End If
            End While
        End Sub

        ''' <summary>
        ''' 强制分割超大块
        ''' </summary>
        Private Function ForceSplitChunk(chunk As TextChunk,
                                          maxSize As Integer,
                                          overlap As Integer) As List(Of TextChunk)

            Dim result As New List(Of TextChunk)()
            Dim text As String = chunk.Text
            Dim pos As Integer = 0
            Dim id As Integer = chunk.Id

            While pos < text.Length
                Dim endPos As Integer = Math.Min(pos + maxSize, text.Length)

                ' 尝试在句子边界对齐
                If endPos < text.Length Then
                    Dim sentenceBoundaries As Char() = {"。"c, "."c, "!"c, "?"c, "；"c, ";"c}
                    For Each boundary As Char In sentenceBoundaries
                        Dim lastBoundary As Integer = text.LastIndexOf(boundary, endPos - 1)
                        If lastBoundary > pos + _minChunkSize Then
                            endPos = lastBoundary + 1
                            Exit For
                        End If
                    Next
                End If

                Dim chunkText As String = text.Substring(pos, endPos - pos).Trim()
                If chunkText.Length >= _minChunkSize Then
                    result.Add(New TextChunk With {
                        .Id = System.Threading.Interlocked.Increment(id),
                        .Text = chunkText,
                        .CharCount = chunkText.Length,
                        .StartPosition = chunk.StartPosition + pos,
                        .EndPosition = chunk.StartPosition + endPos,
                        .SourceSection = chunk.SourceSection,
                        .Level = chunk.Level
                    })
                End If

                ' 应用重叠（防止死循环：确保 pos 一定前进）
                Dim advance As Integer = endPos - pos
                If advance <= 0 Then
                    ' 没有前进，强制跳过
                    pos = endPos + 1
                Else
                    pos = endPos - Math.Min(overlap, advance \ 2)
                    If pos <= endPos - advance Then pos = endPos
                End If
            End While

            Return result
        End Function

        ' ==================== 固定长度分块 ====================

        ''' <summary>
        ''' 固定长度等距切割，带重叠和句子边界对齐
        ''' </summary>
        Public Function ChunkByFixedLength(text As String,
                                            Optional chunkSize As Integer = -1,
                                            Optional overlap As Integer = -1) As List(Of TextChunk)

            If chunkSize < 0 Then chunkSize = _defaultChunkSize
            If overlap < 0 Then overlap = _defaultOverlap

            Dim chunks As New List(Of TextChunk)()
            Dim chunkId As Integer = 0
            Dim pos As Integer = 0

            Debug.WriteLine(String.Format("[SemanticChunker] FixedLength分块开始: 块大小={0}, 重叠={1}",
                                          chunkSize, overlap))

            While pos < text.Length
                Dim endPos As Integer = Math.Min(pos + chunkSize, text.Length)

                ' 尝试在句子边界对齐（查找最近的 。. ! ? 切开）
                If endPos < text.Length Then
                    Dim aligned As Boolean = False
                    Dim searchStart As Integer = Math.Max(endPos - 50, pos + _minChunkSize)

                    For searchPos As Integer = endPos - 1 To searchStart Step -1
                        Dim ch As Char = text(searchPos)
                        If ch = "。"c OrElse ch = "."c OrElse ch = "!"c OrElse ch = "?"c OrElse
                           ch = "；"c OrElse ch = ";"c OrElse ch = "，"c OrElse ch = ","c Then
                            endPos = searchPos + 1
                            aligned = True
                            Exit For
                        End If
                    Next

                    If Not aligned Then
                        Debug.WriteLine(String.Format("[SemanticChunker] 位置{0}: 未找到句子边界，硬切割", endPos))
                    End If
                End If

                Dim chunkText As String = text.Substring(pos, endPos - pos).Trim()
                If chunkText.Length >= _minChunkSize Then
                    chunks.Add(New TextChunk With {
                        .Id = System.Threading.Interlocked.Increment(chunkId),
                        .Text = chunkText,
                        .CharCount = chunkText.Length,
                        .StartPosition = pos,
                        .EndPosition = endPos,
                        .SourceSection = "",
                        .Level = 0
                    })
                End If

                ' 移动到下一块（保留重叠，防止死循环）
                Dim advanceLen As Integer = endPos - pos
                If advanceLen <= 0 Then
                    pos = endPos + 1
                Else
                    pos = endPos - Math.Min(overlap, advanceLen \ 2)
                    If pos <= endPos - advanceLen Then pos = endPos
                End If
            End While

            Debug.WriteLine(String.Format("[SemanticChunker] FixedLength分块完成: 共{0}个块", chunks.Count))
            Return chunks
        End Function

        ' ==================== 关键词检索 ====================

        ''' <summary>
        ''' 轻量级关键词检索（无需向量库）
        ''' 基于关键词命中分数排序，返回最相关的topK个块
        ''' </summary>
        Public Function RetrieveRelevantChunks(question As String,
                                                 chunks As List(Of TextChunk),
                                                 Optional topK As Integer = 5) As List(Of TextChunk)

            If String.IsNullOrWhiteSpace(question) OrElse
               chunks Is Nothing OrElse chunks.Count = 0 Then
                Debug.WriteLine("[SemanticChunker] RetrieveRelevantChunks: 输入为空")
                Return New List(Of TextChunk)()
            End If

            topK = Math.Min(topK, chunks.Count)

            Debug.WriteLine(String.Format("[SemanticChunker] 关键词检索开始: 问题='{0}', topK={1}",
                                          question.Substring(0, Math.Min(50, question.Length)), topK))

            ' 1. 分词提取关键词
            Dim keywords As List(Of String) = ExtractKeywords(question)
            Debug.WriteLine(String.Format("[SemanticChunker] 提取关键词: [{0}]",
                                          String.Join(", ", keywords)))

            ' 2. 计算每个chunk的关键词命中分数
            Dim scoredChunks As New List(Of KeyValuePair(Of TextChunk, Integer))()

            For Each chunk As TextChunk In chunks
                Dim score As Integer = CalculateKeywordScore(chunk.Text, keywords)
                scoredChunks.Add(New KeyValuePair(Of TextChunk, Integer)(chunk, score))
            Next

            ' 3. 按分数降序排列
            scoredChunks.Sort(Function(a, b) b.Value.CompareTo(a.Value))

            ' 4. 取topK并保持原文顺序
            Dim result As New List(Of TextChunk)()
            Dim selectedIndices As New HashSet(Of Integer)()

            For i As Integer = 0 To Math.Min(topK, scoredChunks.Count) - 1
                Dim pair As KeyValuePair(Of TextChunk, Integer) = scoredChunks(i)
                If pair.Value > 0 Then
                    result.Add(pair.Key)
                    selectedIndices.Add(pair.Key.Id)
                    Debug.WriteLine(String.Format("[SemanticChunker] 命中块 #{0}: 分数={1}, 预览={2}",
                                                  pair.Key.Id, pair.Value,
                                                  pair.Key.Text.Substring(0, Math.Min(50, pair.Key.Text.Length))))
                End If
            Next

            ' 按原始ID排序（保持原文顺序）
            result.Sort(Function(a, b) a.Id.CompareTo(b.Id))

            Debug.WriteLine(String.Format("[SemanticChunker] 关键词检索完成: 返回{0}个相关块", result.Count))
            Return result
        End Function

        ''' <summary>
        ''' 从问题文本中提取关键词
        ''' 支持中文和英文分词
        ''' </summary>
        Private Function ExtractKeywords(question As String) As List(Of String)
            Dim keywords As New List(Of String)()

            ' 按空格分割英文单词
            Dim words As String() = question.Split({" "c, ControlChars.Tab,
                                                     "，"c, "、"c, "。"c,
                                                     "？"c, "！"c, ";"c,
                                                     "；"c, ","c},
                                                    StringSplitOptions.RemoveEmptyEntries)

            For Each word As String In words
                word = word.Trim()
                If word.Length < 2 Then Continue For
                If StopWords.Contains(word) Then Continue For

                keywords.Add(word.ToLower())
            Next

            ' 中文按2字切分（简单bigram）
            Dim chineseOnly As String = ExtractChineseCharacters(question)
            For i As Integer = 0 To chineseOnly.Length - 2
                Dim bigram As String = chineseOnly.Substring(i, 2)
                If Not StopWords.Contains(bigram) Then
                    keywords.Add(bigram)
                End If
            Next

            ' 去重
            Dim uniqueKeywords As New List(Of String)(New HashSet(Of String)(keywords))

            Return uniqueKeywords
        End Function

        ''' <summary>
        ''' 提取中文字符
        ''' </summary>
        Private Function ExtractChineseCharacters(text As String) As String
            Dim sb As New StringBuilder()
            For Each ch As Char In text
                If ch >= "一"c AndAlso ch <= "龥"c Then
                    sb.Append(ch)
                End If
            Next
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 计算关键词在文本中的命中分数
        ''' </summary>
        Private Function CalculateKeywordScore(text As String,
                                                keywords As List(Of String)) As Integer

            Dim score As Integer = 0
            Dim lowerText As String = text.ToLower()

            For Each keyword As String In keywords
                If String.IsNullOrEmpty(keyword) Then Continue For

                ' 统计关键词出现次数
                Dim idx As Integer = 0
                Do
                    idx = lowerText.IndexOf(keyword, idx, StringComparison.Ordinal)
                    If idx < 0 Then Exit Do
                    score += 1
                    idx += keyword.Length
                Loop
            Next

            Return score
        End Function

        ' ==================== 摘要生成 ====================

        ''' <summary>
        ''' 生成全局摘要文本：将所有chunks的前100字拼装成概览
        ''' </summary>
        Public Function BuildChunkSummary(chunks As List(Of TextChunk)) As String
            If chunks Is Nothing OrElse chunks.Count = 0 Then
                Debug.WriteLine("[SemanticChunker] BuildChunkSummary: 无块数据")
                Return ""
            End If

            Dim summary As New StringBuilder()
            Dim previewLength As Integer = 100  ' 每块预览字符数
            Dim totalPreviewed As Integer = 0

            summary.AppendLine("=== 文档概览 ===")
            summary.AppendLine(String.Format("总块数: {0}", chunks.Count))
            summary.AppendLine()

            For Each chunk As TextChunk In chunks
                Dim preview As String = ""
                If chunk.Text.Length > previewLength Then
                    preview = chunk.Text.Substring(0, previewLength) & "..."
                Else
                    preview = chunk.Text
                End If

                totalPreviewed += preview.Length

                Dim sectionInfo As String = If(Not String.IsNullOrWhiteSpace(chunk.SourceSection),
                                                "[" & chunk.SourceSection & "] ",
                                                "")

                summary.AppendLine(String.Format("--- 块 #{0} ({1}字符) {2}---",
                                                chunk.Id, chunk.CharCount, sectionInfo))
                summary.AppendLine(preview)
                summary.AppendLine()
            Next

            Debug.WriteLine(String.Format("[SemanticChunker] BuildChunkSummary完成: 总预览{0}字符", totalPreviewed))
            Return summary.ToString()
        End Function

    End Class

End Namespace