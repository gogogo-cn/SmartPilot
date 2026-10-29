Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json

Namespace Services.RAG

    ''' <summary>
    ''' RAG 检索增强生成编排器（单例）
    ''' 用于在对话流程中调用：检索默认知识库 → 命中则组装 Prompt 上下文 → 未命中则降级普通 chat
    ''' 默认知识库 ID 从 RagConfig.DefaultKnowledgeBaseId 读取（前端"RAG 设置"里配置）。
    ''' </summary>
    Public Class RagOrchestrator

        Private Shared ReadOnly _instance As New Lazy(Of RagOrchestrator)(Function() New RagOrchestrator())

        Private _kbManager As KnowledgeBaseManager

        Public Shared ReadOnly Property Instance As RagOrchestrator
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
            _kbManager = KnowledgeBaseManager.Instance
        End Sub

#Region "公开方法"

        ''' <summary>
        ''' RAG 主入口 - 在对话流程中调用
        '''
        ''' 流程:
        ''' 1. 检查 RAG 是否启用（config.EnableRagForChat）
        ''' 2. 确定目标知识库（传入 ID 优先，否则用配置的默认知识库）
        ''' 3. 向量检索
        ''' 4. 检索结果为空 → 返回 ShouldFallback=True（降级到普通 chat）
        ''' 5. 有结果 → 组装 Prompt 上下文，返回增强后的上下文字符串（由调用方送给 LLM）
        ''' </summary>
        ''' <param name="userQuestion">用户原始问题</param>
        ''' <param name="optionalKbId">可选的知识库 ID（不传则使用配置的默认知识库）</param>
        ''' <returns>RagResult: 包含上下文文本和来源引用，或 ShouldFallback=True 表示降级</returns>
        Public Async Function RetrieveAndEnhanceAsync(userQuestion As String,
                                                        Optional optionalKbId As String = "") As Task(Of RagResult)
            Dim result As New RagResult With {
                .HasContext = False,
                .ContextText = "",
                .SourceReferences = New List(Of SourceRef)(),
                .RetrievedChunks = 0,
                .ShouldFallback = True  ' 默认降级
            }

            If String.IsNullOrWhiteSpace(userQuestion) Then Return result

            Try
                ' 检查 RAG 是否启用
                Dim config = _kbManager.GetConfig()
                If Not config.EnableRagForChat Then
                    System.Diagnostics.Debug.WriteLine("[RagOrch] RAG 功能已禁用")
                    Return result
                End If

                ' 确定目标知识库
                Dim kbId = If(String.IsNullOrEmpty(optionalKbId), config.DefaultKnowledgeBaseId, optionalKbId)

                If String.IsNullOrEmpty(kbId) Then
                    System.Diagnostics.Debug.WriteLine("[RagOrch] 未配置默认知识库 ID")
                    Return result
                End If

                System.Diagnostics.Debug.WriteLine($"[RagOrch] 开始 RAG 检索: 问题='{userQuestion.Substring(0, Math.Min(50, userQuestion.Length))}...', KB={kbId}")

                ' 执行检索
                Dim searchResults = Await _kbManager.SearchAsync(
                    kbId,
                    userQuestion,
                    config.DefaultTopK,
                    New SearchOptions With {
                        .ScoreThreshold = config.ScoreThreshold
                    }
                ).ConfigureAwait(False)

                If searchResults Is Nothing OrElse searchResults.Count = 0 Then
                    ' 检索为空，返回降级标志（不返回"不知道"）
                    System.Diagnostics.Debug.WriteLine("[RagOrch] 检索结果为空，将降级到普通 chat")
                    result.ShouldFallback = True
                    Return result
                End If

                ' 检索成功！构建上下文
                result.HasContext = True
                result.ShouldFallback = False
                result.RetrievedChunks = searchResults.Count

                ' 截取前 N 个块作为上下文
                Dim takeCount As Integer = Math.Min(searchResults.Count, Math.Max(1, config.MaxContextChunks))
                Dim contextChunks = searchResults.Take(takeCount).ToList()

                ' 构建来源引用列表
                For Each chunk In contextChunks
                    result.SourceReferences.Add(New SourceRef With {
                        .DocTitle = If(chunk.DocTitle, "未知文档"),
                        .Section = If(chunk.SourceSection, ""),
                        .Score = chunk.Score
                    })
                Next

                ' 组装 RAG Prompt
                result.ContextText = BuildRagPrompt(userQuestion, contextChunks)

                System.Diagnostics.Debug.WriteLine($"[RagOrch] RAG 增强: 检索到 {result.RetrievedChunks} 个块, 使用 {contextChunks.Count} 个块作为上下文")

            Catch ex As Exception
                ' RAG 出错也降级
                System.Diagnostics.Debug.WriteLine($"[RagOrch] RAG 异常: {ex.Message}, 将降级到普通 chat")
                result.ShouldFallback = True
            End Try

            Return result
        End Function

#End Region

#Region "内部方法"

        ''' <summary>
        ''' 构建 RAG Prompt 模板
        ''' </summary>
        Private Function BuildRagPrompt(originalQuestion As String,
                                          contextChunks As List(Of SearchResult)) As String
            Dim contextBuilder As New StringBuilder()

            contextBuilder.AppendLine("【知识库参考资料】")
            contextBuilder.AppendLine("请基于以下参考资料回答用户问题，不要编造资料中没有的信息；")
            contextBuilder.AppendLine("如果参考资料与问题无关，请忽略参考资料直接回答：")
            contextBuilder.AppendLine()

            For i As Integer = 0 To contextChunks.Count - 1
                Dim chunk = contextChunks(i)
                Dim docTitle = If(String.IsNullOrEmpty(chunk.DocTitle), "未知", chunk.DocTitle)
                contextBuilder.AppendLine($"--- 资料 {i + 1} (相关度: {chunk.Score:P0}, 来源: {docTitle}) ---")
                contextBuilder.AppendLine(chunk.Text)
                contextBuilder.AppendLine()
            Next

            contextBuilder.AppendLine("---")
            contextBuilder.AppendLine($"用户问题: {originalQuestion}")
            contextBuilder.AppendLine("请回答:")

            Return contextBuilder.ToString()
        End Function

#End Region

    End Class

End Namespace
