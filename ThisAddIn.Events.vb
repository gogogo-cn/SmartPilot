Option Strict On
Option Infer On

Imports System
Imports Microsoft.Web.WebView2.Core
Imports Newtonsoft.Json.Linq

Partial Public NotInheritable Class ThisAddIn

#Region "事件处理"

    ''' <summary>
    ''' WebView2消息处理(旧主面板通道,新WPF面板不再使用)
    ''' </summary>
    Private Sub OnWebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim message As String = e.TryGetWebMessageAsString()
            If String.IsNullOrEmpty(message) Then Return

            Dim msgObj As JObject = JObject.Parse(message)
            Dim msgType As String = msgObj("type")?.ToString()

            Select Case msgType
                Case "showSettings"
                    ShowSettingsPanel()
                Case "showOcr"
                    ShowOcrPanel()
                Case "startScreenshot"
                    StartScreenshotOcr()
                Case "reloadConfig"
                    SendConfigToUI()
            End Select
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Handle WebMessage failed", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 手机OCR完成回调(切换到UI线程,等待OCR面板WebView就绪后发送结果)
    ''' </summary>
    Private Sub OnMobileOcrCompleted(text As String, fileName As String, imageBase64 As String)
        Dim textLen As Integer = If(text Is Nothing, 0, text.Length)
        _chatLogger?.LogInfo("ADDIN", $"Mobile recognition complete callback triggered, text length={textLen}, fileName={fileName}")
        Try
            EnsureTaskPaneManager()
            If _taskPaneManager Is Nothing Then
                _chatLogger?.LogError("ADDIN", "_taskPaneManager is Nothing", New Exception("TaskPaneManager null"))
                Return
            End If
            _chatLogger?.LogInfo("ADDIN", "Preparing to send mobile OCR result to panel")
            _taskPaneManager.SendMessageToOcrPanelWhenReady(New With {
                Key .type = "mobileOcrResult",
                Key .text = text,
                Key .fileName = fileName,
                Key .imageBase64 = imageBase64
            })
            _chatLogger?.LogInfo("ADDIN", "Mobile OCR result sent to OCR panel")
        Catch ex As Exception
            _chatLogger?.LogError("ADDIN", "Mobile OCR complete callback failed: " & ex.Message, ex)
        End Try
    End Sub

    ''' <summary>
    ''' 手机上传回调
    ''' </summary>
    Private Sub OnMobileUploadReceived(fileName As String)
        Try
            _chatLogger?.LogInfo("ADDIN", "Received mobile upload: " & fileName)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[ThisAddIn] Mobile upload callback failed: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' 请求服务(Ribbon扩展)
    ''' </summary>
    Protected Overrides Function RequestService(serviceGuid As Guid) As Object
        If serviceGuid = GetType(Microsoft.Office.Core.IRibbonExtensibility).GUID Then
            If _ribbon Is Nothing Then
                _ribbon = New Ribbon()
            End If
            Return _ribbon
        End If
        Return MyBase.RequestService(serviceGuid)
    End Function

#End Region

End Class