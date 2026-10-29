Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Sockets
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 插件进程访问独立 OCR 服务的客户端。服务只监听本机，请求带一次性令牌。
    ''' </summary>
    Public NotInheritable Class OcrHostClient

        Private Sub New()
        End Sub

        Private Shared ReadOnly _startLock As New Object()
        Private Const HealthTimeoutMs As Integer = 2000
        Private Const CallTimeoutMs As Integer = 900000
        Private Const StartWaitMs As Integer = 20000

        Public Shared ReadOnly JsonSettings As JsonSerializerSettings = CreateSettings()

        Private Shared Function CreateSettings() As JsonSerializerSettings
            Dim settings As New JsonSerializerSettings()
            settings.NullValueHandling = NullValueHandling.Include
            settings.MissingMemberHandling = MissingMemberHandling.Ignore
            settings.Converters.Add(New StringEnumConverter())
            Return settings
        End Function

        Public Shared ReadOnly Property EndpointFilePath As String
            Get
                Return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAIVSTOAddIn",
                    "ocr-host.json")
            End Get
        End Property

        Public Shared Sub EnsureStarted()
            If IsHealthy() Then Return
            SyncLock _startLock
                If IsHealthy() Then Return
                Dim exePath = FindHostExe()
                If String.IsNullOrEmpty(exePath) Then
                    Throw New FileNotFoundException("找不到 OcrHost.exe。请先编译 OcrHost 项目。")
                End If
                Dim pluginDir = FindPluginDir()
                Dim psi As New ProcessStartInfo()
                psi.FileName = exePath
                psi.WorkingDirectory = If(String.IsNullOrEmpty(pluginDir), Path.GetDirectoryName(exePath), pluginDir)
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.EnvironmentVariables("SMARTPILOT_OCR_HOST") = "1"
                If Not String.IsNullOrEmpty(pluginDir) Then
                    psi.EnvironmentVariables("SMARTPILOT_PLUGIN_DIR") = pluginDir
                End If
                Dim started = Process.Start(psi)
                If started IsNot Nothing Then
                    Try
                        started.PriorityClass = ProcessPriorityClass.BelowNormal
                    Catch
                    End Try
                End If
                Dim deadline = DateTime.UtcNow.AddMilliseconds(StartWaitMs)
                Do
                    If IsHealthy() Then Return
                    Thread.Sleep(200)
                Loop While DateTime.UtcNow < deadline
                Throw New TimeoutException("OCR 服务启动超时")
            End SyncLock
        End Sub

        Public Shared Function Invoke(Of T)(method As String, ParamArray args() As Object) As T
            EnsureStarted()
            Dim payload = JsonConvert.SerializeObject(New With {
                .method = method,
                .args = If(args, New Object() {})
            }, JsonSettings)
            Dim body = Send("POST", "/v1/call", payload, CallTimeoutMs)
            Dim root = JObject.Parse(body)
            If root.Value(Of Boolean?)("ok") <> True Then
                Dim err = root.Value(Of String)("error")
                Throw New InvalidOperationException(If(String.IsNullOrEmpty(err), "OCR 服务调用失败", err))
            End If
            Dim resultToken = root("result")
            If resultToken Is Nothing OrElse resultToken.Type = JTokenType.Null Then
                Return Nothing
            End If
            Return resultToken.ToObject(Of T)(JsonSerializer.Create(JsonSettings))
        End Function

        Private Shared Function IsHealthy() As Boolean
            Try
                Dim info = ReadEndpoint()
                If info Is Nothing Then Return False
                If Not IsProcessAlive(info.Pid) Then Return False
                Dim body = Send("GET", "/health", Nothing, HealthTimeoutMs, info, False)
                Return body.IndexOf("""ok"":true", StringComparison.OrdinalIgnoreCase) >= 0
            Catch
                Return False
            End Try
        End Function

        Private Class EndpointInfo
            Public Property Port As Integer
            Public Property Token As String
            Public Property Pid As Integer
        End Class

        Private Shared Function ReadEndpoint() As EndpointInfo
            If Not File.Exists(EndpointFilePath) Then Return Nothing
            Dim json = File.ReadAllText(EndpointFilePath, Encoding.UTF8)
            Dim info = JsonConvert.DeserializeObject(Of EndpointInfo)(json)
            If info Is Nothing OrElse info.Port <= 0 OrElse String.IsNullOrEmpty(info.Token) Then Return Nothing
            Return info
        End Function

        Private Shared Function IsProcessAlive(pid As Integer) As Boolean
            If pid <= 0 Then Return False
            Try
                Dim proc = Process.GetProcessById(pid)
                Return proc IsNot Nothing AndAlso Not proc.HasExited
            Catch
                Return False
            End Try
        End Function

        Private Shared Function Send(method As String, path As String, body As String, timeoutMs As Integer) As String
            Dim info = ReadEndpoint()
            If info Is Nothing Then Throw New InvalidOperationException("OCR 服务未就绪")
            Return Send(method, path, body, timeoutMs, info, True)
        End Function

        Private Shared Function Send(method As String, path As String, body As String, timeoutMs As Integer, info As EndpointInfo, sendToken As Boolean) As String
            Dim payload = If(body, "")
            Dim payloadBytes = Encoding.UTF8.GetBytes(payload)
            Dim header As New StringBuilder()
            header.Append(method).Append(" ").Append(path).Append(" HTTP/1.1").Append(vbCrLf)
            header.Append("Host: 127.0.0.1:").Append(info.Port).Append(vbCrLf)
            header.Append("Connection: close").Append(vbCrLf)
            If sendToken Then
                header.Append("X-Ocr-Token: ").Append(info.Token).Append(vbCrLf)
            End If
            If payloadBytes.Length > 0 Then
                header.Append("Content-Type: application/json; charset=utf-8").Append(vbCrLf)
                header.Append("Content-Length: ").Append(payloadBytes.Length).Append(vbCrLf)
            End If
            header.Append(vbCrLf)

            Using client As New TcpClient()
                client.ReceiveTimeout = timeoutMs
                client.SendTimeout = timeoutMs
                Dim connect = client.BeginConnect("127.0.0.1", info.Port, Nothing, Nothing)
                If Not connect.AsyncWaitHandle.WaitOne(timeoutMs) Then
                    Throw New TimeoutException("连接 OCR 服务超时")
                End If
                client.EndConnect(connect)
                Using stream = client.GetStream()
                    Dim headerBytes = Encoding.ASCII.GetBytes(header.ToString())
                    stream.Write(headerBytes, 0, headerBytes.Length)
                    If payloadBytes.Length > 0 Then stream.Write(payloadBytes, 0, payloadBytes.Length)
                    stream.Flush()
                    Using reader As New StreamReader(stream, Encoding.UTF8, False, 8192, True)
                        Dim status = reader.ReadLine()
                        If String.IsNullOrEmpty(status) Then Throw New IOException("OCR 服务没有返回")
                        Dim contentLength As Integer = 0
                        Dim line = reader.ReadLine()
                        While line IsNot Nothing AndAlso line.Length > 0
                            If line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase) Then
                                Integer.TryParse(line.Substring(15).Trim(), contentLength)
                            End If
                            line = reader.ReadLine()
                        End While
                        If contentLength <= 0 Then Return ""
                        Dim buffer(contentLength - 1) As Char
                        Dim readTotal = 0
                        While readTotal < contentLength
                            Dim n = reader.Read(buffer, readTotal, contentLength - readTotal)
                            If n <= 0 Then Exit While
                            readTotal += n
                        End While
                        Dim responseText = New String(buffer, 0, readTotal)
                        If status.IndexOf(" 401 ", StringComparison.Ordinal) >= 0 Then
                            Throw New InvalidOperationException("OCR 服务拒绝了令牌")
                        End If
                        If status.IndexOf(" 200 ", StringComparison.Ordinal) < 0 Then
                            Throw New InvalidOperationException("OCR 服务返回 " & status & " " & responseText)
                        End If
                        Return responseText
                    End Using
                End Using
            End Using
        End Function

        Private Shared Function FindHostExe() As String
            Dim dirs As New List(Of String)
            dirs.Add(AppDomain.CurrentDomain.BaseDirectory)
            Try
                Dim asmDir = Path.GetDirectoryName(GetType(OcrHostClient).Assembly.Location)
                If Not String.IsNullOrEmpty(asmDir) Then dirs.Add(asmDir)
            Catch
            End Try
            For Each dirPath In dirs
                Dim candidate = Path.Combine(dirPath, "OcrHost.exe")
                If File.Exists(candidate) Then Return candidate
            Next
            Return Nothing
        End Function

        Private Shared Function FindPluginDir() As String
            Dim dirs As New List(Of String)
            dirs.Add(AppDomain.CurrentDomain.BaseDirectory)
            Try
                Dim asmDir = Path.GetDirectoryName(GetType(OcrHostClient).Assembly.Location)
                If Not String.IsNullOrEmpty(asmDir) Then dirs.Add(asmDir)
            Catch
            End Try
            For Each dirPath In dirs
                If File.Exists(Path.Combine(dirPath, "Services", "rapidocr_inference_engine.py")) Then Return dirPath
                If File.Exists(Path.Combine(dirPath, "rapidocr_inference_engine.py")) Then Return dirPath
            Next
            Return AppDomain.CurrentDomain.BaseDirectory
        End Function

        Public Shared Function CreateToken() As String
            Dim bytes(31) As Byte
            Using rng As New RNGCryptoServiceProvider()
                rng.GetBytes(bytes)
            End Using
            Dim sb As New StringBuilder(64)
            For Each b In bytes
                sb.Append(b.ToString("x2"))
            Next
            Return sb.ToString()
        End Function

    End Class

End Namespace
