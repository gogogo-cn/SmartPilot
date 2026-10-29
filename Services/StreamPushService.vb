Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    Public Class StreamPushService

        Private Shared ReadOnly _lazyInstance As New Lazy(Of StreamPushService)(
            Function() New StreamPushService(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As StreamPushService
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private _activeStreams As New Dictionary(Of String, StreamSession)()
        Private ReadOnly _lock As New Object()

        Public Class StreamSession
            Public Property StreamId As String
            Public Property CreatedAt As DateTime = DateTime.Now
            Public Property IsCompleted As Boolean = False
            Public Property ChunkCount As Integer = 0
            Public Property FullContent As New StringBuilder()
        End Class

        Public Function CreateStream(Optional streamId As String = Nothing) As String
            If String.IsNullOrEmpty(streamId) Then
                streamId = Guid.NewGuid().ToString("N").Substring(0, 8)
            End If

            Dim session = New StreamSession With {.StreamId = streamId}

            SyncLock _lock
                _activeStreams(streamId) = session
            End SyncLock

            PushEvent(streamId, "stream_start", New Dictionary(Of String, Object) From {
                {"streamId", streamId},
                {"timestamp", DateTime.Now.ToString("o")}
            })

            Return streamId
        End Function

        Public Sub PushChunk(streamId As String, chunk As SseChunk)
            Dim session = GetSession(streamId)
            If session Is Nothing Then Return

            session.ChunkCount += 1

            If Not String.IsNullOrEmpty(chunk.Content) Then
                session.FullContent.Append(chunk.Content)
            End If

            Dim data As New Dictionary(Of String, Object) From {
                {"streamId", streamId},
                {"content", chunk.Content},
                {"reasoning", chunk.Reasoning}
            }

            If Not String.IsNullOrEmpty(chunk.Model) Then
                data("model") = chunk.Model
            End If

            If chunk.IsDone Then
                data("finishReason") = chunk.FinishReason
                data("fullContent") = session.FullContent.ToString()
            End If

            If chunk.Usage IsNot Nothing AndAlso chunk.Usage.Count > 0 Then
                data("usage") = chunk.Usage
            End If

            Dim eventType = If(chunk.IsDone, "stream_end", "stream_delta")
            PushEvent(streamId, eventType, data)

            If chunk.IsDone Then
                session.IsCompleted = True
            End If
        End Sub

        Public Sub PushError(streamId As String, errorMessage As String)
            PushEvent(streamId, "stream_error", New Dictionary(Of String, Object) From {
                {"streamId", streamId},
                {"error", errorMessage},
                {"timestamp", DateTime.Now.ToString("o")}
            })

            Dim session = GetSession(streamId)
            If session IsNot Nothing Then
                session.IsCompleted = True
            End If
        End Sub

        Public Sub CompleteStream(streamId As String, Optional fullContent As String = Nothing)
            Dim session = GetSession(streamId)
            If session Is Nothing Then Return

            Dim data As New Dictionary(Of String, Object) From {
                {"streamId", streamId},
                {"finishReason", "stop"},
                {"fullContent", If(fullContent, session.FullContent.ToString())},
                {"timestamp", DateTime.Now.ToString("o")}
            }

            PushEvent(streamId, "stream_end", data)
            session.IsCompleted = True
        End Sub

        Private Sub PushEvent(streamId As String, eventType As String, data As Dictionary(Of String, Object))
            Try
                Dim payload As New Dictionary(Of String, Object) From {
                    {"type", "llm_stream"},
                    {"event", eventType},
                    {"data", data}
                }

                Dim json = JsonConvert.SerializeObject(payload, New JsonSerializerSettings With {
                    .NullValueHandling = NullValueHandling.Ignore
                })

                Dim script = $"if(typeof window.__onLLMStream==='function'){{window.__onLLMStream({EscapeForJs(json)});}}"

                Dim addin = Globals.ThisAddIn
                If addin IsNot Nothing Then
                    addin.ExecuteScriptOnWebView(script)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[StreamPushService] PushEvent 失败: {ex.Message}")
            End Try
        End Sub

        Private Function EscapeForJs(json As String) As String
            If String.IsNullOrEmpty(json) Then Return "''"
            Return "'" & json.Replace("\", "\\").Replace("'", "\'").Replace(vbLf, "\n").Replace(vbCr, "\r").Replace(vbTab, "\t") & "'"
        End Function

        Private Function GetSession(streamId As String) As StreamSession
            SyncLock _lock
                If _activeStreams.ContainsKey(streamId) Then
                    Return _activeStreams(streamId)
                End If
            End SyncLock
            Return Nothing
        End Function

        Public Sub RemoveStream(streamId As String)
            SyncLock _lock
                _activeStreams.Remove(streamId)
            End SyncLock
        End Sub

        Public Sub CleanupCompleted()
            SyncLock _lock
                Dim expired = _activeStreams.Where(Function(kvp) _
                    kvp.Value.IsCompleted AndAlso (DateTime.Now - kvp.Value.CreatedAt).TotalMinutes > 5).ToList()
                For Each kvp In expired
                    _activeStreams.Remove(kvp.Key)
                Next
            End SyncLock
        End Sub
    End Class

End Namespace