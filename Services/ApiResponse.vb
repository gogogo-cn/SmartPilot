Option Strict On
Option Explicit On
Option Infer On

Imports Newtonsoft.Json
Imports Newtonsoft.Json.Serialization

Namespace Services

    Public Class ApiResponse

        <JsonProperty("success")>
        Public Property Success As Boolean

        <JsonProperty("data", NullValueHandling:=NullValueHandling.Ignore)>
        Public Property Data As Object

        <JsonProperty("error", NullValueHandling:=NullValueHandling.Ignore)>
        Public Property [Error] As String

        <JsonProperty("errorMessage", NullValueHandling:=NullValueHandling.Ignore)>
        Public Property ErrorMessage As String

        <JsonProperty("message", NullValueHandling:=NullValueHandling.Ignore)>
        Public Property Message As String

        <JsonProperty("timestamp")>
        Public ReadOnly Property Timestamp As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Private Sub New()
        End Sub

        Public Shared Function Ok(Optional data As Object = Nothing, Optional message As String = Nothing) As String
            Dim resp = New ApiResponse With {
                .Success = True,
                .Data = data,
                .Message = message
            }
            Return resp.ToJson()
        End Function

        Public Shared Function Fail(errorMsg As String, Optional data As Object = Nothing) As String
            Dim resp = New ApiResponse With {
                .Success = False,
                .[Error] = errorMsg,
                .ErrorMessage = errorMsg,
                .Data = data
            }
            Return resp.ToJson()
        End Function

        Public Shared Function FromException(ex As Exception, Optional context As String = "") As String
            Dim msg = If(String.IsNullOrEmpty(context), ex.Message, $"{context}: {ex.Message}")
            Return Fail(msg)
        End Function

        Private Function ToJson() As String
            Dim settings As New JsonSerializerSettings With {
                .ContractResolver = New CamelCasePropertyNamesContractResolver(),
                .NullValueHandling = NullValueHandling.Ignore,
                .Formatting = Formatting.None
            }
            Return JsonConvert.SerializeObject(Me, settings)
        End Function

    End Class

End Namespace