Option Strict On
Option Explicit On
Option Infer On
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Services

    Public Class SseChunk
        Public Property Content As String = ""
        Public Property Reasoning As String = ""
        Public Property Model As String = ""
        Public Property FinishReason As String = ""
        Public Property IsDone As Boolean = False
        Public Property IsError As Boolean = False
        Public Property ErrorMessage As String = ""
        Public Property Usage As Dictionary(Of String, Integer)

        Public Sub New()
            Usage = New Dictionary(Of String, Integer)()
        End Sub
    End Class

    Public Class SseParser

        Private _buffer As String = ""
        ' <think> 标签状态机：某些部署（LM Studio / llama.cpp / 部分中转）不提供独立 reasoning 字段，
        ' 而是把思考以 <think>...</think> 混在 content 里输出。不剥离会污染正文、干扰 JSON 解析。
        Private _inThink As Boolean = False
        Private _pendingTail As String = ""   ' 被 chunk 边界截断的部分标签（如 "<thin"），暂存到下一块

        Public Function ProcessChunk(chunk As String) As List(Of SseChunk)
            If String.IsNullOrEmpty(chunk) Then Return New List(Of SseChunk)()

            _buffer &= chunk
            Dim results As New List(Of SseChunk)()
            Dim lines = _buffer.Split({vbLf}, StringSplitOptions.None)
            _buffer = lines(lines.Length - 1)

            For i = 0 To lines.Length - 2
                Dim line = lines(i).TrimEnd(CChar(vbCr))
                Dim parsed = ParseLine(line)
                If parsed IsNot Nothing Then results.Add(parsed)
            Next

            Return results
        End Function

        Private Function ParseLine(line As String) As SseChunk
            Dim trimmed = line.Trim()
            If String.IsNullOrEmpty(trimmed) Then Return Nothing

            If trimmed = "data: [DONE]" Or trimmed = "[DONE]" Then
                Return New SseChunk With {.IsDone = True}
            End If

            If trimmed.StartsWith("error:") Then
                Return New SseChunk With {
                    .IsDone = True,
                    .IsError = True,
                    .ErrorMessage = trimmed.Substring(6).Trim()
                }
            End If

            Dim jsonData As String = Nothing
            If trimmed.StartsWith("data:") Then
                jsonData = trimmed.Substring(5).Trim()
            ElseIf trimmed.StartsWith("{") OrElse trimmed.StartsWith("[") Then
                jsonData = trimmed
            Else
                Return Nothing
            End If

            If String.IsNullOrEmpty(jsonData) Then Return Nothing

            Try
                Dim json = JObject.Parse(jsonData)
                Return ExtractFromJson(json)
            Catch ex As Exception  ' 
                Return Nothing
            End Try
        End Function

        Private Function ExtractFromJson(json As JObject) As SseChunk
            Dim chunk As New SseChunk()

            If json("error") IsNot Nothing Then
                chunk.IsDone = True
                chunk.IsError = True
                chunk.ErrorMessage = If(TypeOf json("error") Is JObject,
                    json("error")("message")?.ToString(),
                    json("error").ToString())
                Return chunk
            End If

            If json("model") IsNot Nothing Then
                chunk.Model = json("model").ToString()
            End If

            Dim choices = json("choices")
            If choices IsNot Nothing AndAlso TypeOf choices Is JArray Then
                Dim arr = DirectCast(choices, JArray)
                If arr.Count > 0 Then
                    Dim choice = arr(0)
                    Dim delta = choice("delta")
                    If delta Is Nothing Then delta = choice("message")

                    If delta IsNot Nothing Then
                        If delta("content") IsNot Nothing Then
                            SplitThinkTags(delta("content").ToString(), chunk)
                        End If
                        If delta("reasoning_content") IsNot Nothing Then
                            chunk.Reasoning &= delta("reasoning_content").ToString()
                        ElseIf delta("reasoning") IsNot Nothing Then
                            chunk.Reasoning &= delta("reasoning").ToString()
                        End If
                    End If

                    If choice("finish_reason") IsNot Nothing Then
                        Dim fr = choice("finish_reason").ToString()
                        If fr = "stop" OrElse fr = "length" OrElse fr = "tool_calls" Then
                            chunk.IsDone = True
                            chunk.FinishReason = fr
                        End If
                    End If
                End If
            ElseIf json("message") IsNot Nothing Then
                Dim msg = json("message")
                If msg("content") IsNot Nothing Then
                    SplitThinkTags(msg("content").ToString(), chunk)
                End If
                If msg("thinking") IsNot Nothing Then
                    chunk.Reasoning &= msg("thinking").ToString()
                ElseIf msg("reasoning") IsNot Nothing Then
                    chunk.Reasoning &= msg("reasoning").ToString()
                End If
                If json("done") IsNot Nothing AndAlso json("done").Value(Of Boolean)() Then
                    chunk.IsDone = True
                End If
            End If

            If json("usage") IsNot Nothing Then
                Dim usageObj = json("usage")
                If usageObj("prompt_tokens") IsNot Nothing Then
                    chunk.Usage("prompt_tokens") = usageObj("prompt_tokens").Value(Of Integer)()
                End If
                If usageObj("completion_tokens") IsNot Nothing Then
                    chunk.Usage("completion_tokens") = usageObj("completion_tokens").Value(Of Integer)()
                End If
            End If

            Return chunk
        End Function

        ''' <summary>
        ''' 把混在 content 流里的 &lt;think&gt;...&lt;/think&gt; 思考段落剥离归入 Reasoning。
        ''' 标签可能被 chunk 边界截断（如一块结尾是 "&lt;thin"），用 _pendingTail 缓存到下一块拼接。
        ''' </summary>
        Private Sub SplitThinkTags(raw As String, chunk As SseChunk)
            Dim data = _pendingTail & If(raw, "")
            _pendingTail = ""
            If String.IsNullOrEmpty(data) Then Return

            Dim sbContent As New StringBuilder()
            Dim sbReason As New StringBuilder()
            Dim i = 0
            While i < data.Length
                If Not _inThink Then
                    Dim openIdx = data.IndexOf("<think>", i, StringComparison.Ordinal)
                    If openIdx >= 0 Then
                        sbContent.Append(data.Substring(i, openIdx - i))
                        _inThink = True
                        i = openIdx + 7
                    Else
                        ' 尾部可能是被截断的 "<think>" 前缀 → 暂存到下块
                        Dim tailLen = Math.Min(6, data.Length - i)
                        Dim matched = 0
                        For tryLen As Integer = tailLen To 1 Step -1
                            If "<think>".StartsWith(data.Substring(data.Length - tryLen), StringComparison.Ordinal) Then
                                matched = tryLen
                                Exit For
                            End If
                        Next
                        sbContent.Append(data.Substring(i, data.Length - i - matched))
                        If matched > 0 Then _pendingTail = data.Substring(data.Length - matched)
                        i = data.Length
                    End If
                Else
                    Dim closeIdx = data.IndexOf("</think>", i, StringComparison.Ordinal)
                    If closeIdx >= 0 Then
                        sbReason.Append(data.Substring(i, closeIdx - i))
                        _inThink = False
                        i = closeIdx + 8
                    Else
                        Dim tailLen = Math.Min(7, data.Length - i)
                        Dim matched = 0
                        For tryLen As Integer = tailLen To 1 Step -1
                            If "</think>".StartsWith(data.Substring(data.Length - tryLen), StringComparison.Ordinal) Then
                                matched = tryLen
                                Exit For
                            End If
                        Next
                        sbReason.Append(data.Substring(i, data.Length - i - matched))
                        If matched > 0 Then _pendingTail = data.Substring(data.Length - matched)
                        i = data.Length
                    End If
                End If
            End While

            chunk.Content = sbContent.ToString()
            If sbReason.Length > 0 Then
                chunk.Reasoning &= sbReason.ToString()
            End If
        End Sub

        ''' <summary>
        ''' 非流式一次性剥离 &lt;think&gt;...&lt;/think&gt; 块。思考内容通过 thinkingOut 返回；
        ''' 无 think 块时原样返回 text、thinkingOut 为空。
        ''' </summary>
        Public Shared Function StripThinkBlocks(text As String, ByRef thinkingOut As String) As String
            thinkingOut = ""
            If String.IsNullOrEmpty(text) OrElse text.IndexOf("<think>", StringComparison.Ordinal) < 0 Then Return text
            Dim sbContent As New StringBuilder()
            Dim sbThink As New StringBuilder()
            Dim rest = text
            While True
                Dim openIdx = rest.IndexOf("<think>", StringComparison.Ordinal)
                If openIdx < 0 Then Exit While
                Dim closeIdx = rest.IndexOf("</think>", openIdx + 7, StringComparison.Ordinal)
                If closeIdx < 0 Then
                    ' 未闭合（格式异常/截断）：剩余整段当思考
                    sbThink.Append(rest.Substring(openIdx + 7))
                    rest = rest.Substring(0, openIdx)
                    Exit While
                End If
                sbContent.Append(rest.Substring(0, openIdx))
                sbThink.Append(rest.Substring(openIdx + 7, closeIdx - openIdx - 7))
                rest = rest.Substring(closeIdx + 8)
            End While
            thinkingOut = sbThink.ToString()
            If sbThink.Length = 0 Then Return text
            Return sbContent.ToString() & rest
        End Function

        Public Sub Reset()
            _buffer = ""
            _inThink = False
            _pendingTail = ""
        End Sub
    End Class

End Namespace