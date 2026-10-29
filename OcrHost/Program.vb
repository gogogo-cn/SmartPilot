Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Net
Imports System.Net.Sockets
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Bridges
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils

''' <summary>
''' 独立 OCR 进程。Excel、Word、PPT 通过本机回环地址共用这一份引擎。
''' </summary>
Public Module Program

        Private ReadOnly _allowed As New HashSet(Of String)(StringComparer.Ordinal) From {
            "RecognizeText",
            "RecognizeTextWithPositionAsync",
            "RecognizeTextWithEngine",
            "RecognizeDocumentAsync",
            "RecognizeTextWithFormatAsync",
            "IsEngineAvailable",
            "GetAvailableEngines",
            "GetEngineStatus",
            "CheckFormulaOcrStatusAsync",
            "StartFormulaOcrInstall",
            "GetFormulaInstallStatus",
            "InstallFormulaOcrAsync",
            "ReloadCloudOcrConfig",
            "ForceCheckRapidAvailable",
            "GetCloudOcrConfig",
            "SaveCloudOcrConfig",
            "TestCloudOcrConfig",
            "GetCloudProviders",
            "HostEcho",
            "SaveBase64Image",
            "GetOcrEngineStatus",
            "GetCurrentOcrEngine",
            "SetOcrEngine",
            "StartOcrRecognitionWithFormat",
            "GetOcrResultWithFormat",
            "GetOcrStructuredResult",
            "GetOcrRawData",
            "CheckFormulaOcrStatus",
            "StartFormulaOcrInstall",
            "GetFormulaInstallStatus",
            "GetParseScripts",
            "GetParseScript",
            "GetParseScriptCategories",
            "SaveParseScript",
            "DeleteParseScript",
            "ResetBuiltinScript",
            "DuplicateParseScript",
            "ImportParseScriptFromMarkdown",
            "ClearParseScriptCache",
            "SearchParseScripts",
            "GetOcrElementsJson",
            "GetPaddleOcrAdvancedConfig",
            "SetPaddleOcrAdvancedConfig",
            "CaptureScreenshots"
        }

        Private ReadOnly _callLock As New Object()
        Private _listener As TcpListener
        Private _token As String
        Private _activeCalls As Integer
        Private _lastActivityUtc As DateTime = DateTime.UtcNow
        Private _cleaned As Integer
        Private Const IdleMinutes As Integer = 10

        Public Sub Main(args As String())
            Environment.SetEnvironmentVariable("SMARTPILOT_OCR_HOST", "1")
            If args IsNot Nothing AndAlso args.Length > 0 AndAlso args(0) = "--self-test" Then
                RunSelfTest()
                Return
            End If

            Dim createdNew As Boolean = False
            Dim mutex As Mutex = Nothing
            Try
                mutex = New Mutex(False, "Local\SmartPilotOcrHost")
                Try
                    createdNew = mutex.WaitOne(0)
                Catch ex As AbandonedMutexException
                    createdNew = True
                End Try
            Catch ex As Exception
                WriteLog("创建互斥量失败: " & ex.Message)
                Environment.Exit(1)
            End Try
            If Not createdNew Then
                Environment.Exit(0)
            End If

            Try
                _token = OcrHostClient.CreateToken()
                _listener = New TcpListener(IPAddress.Loopback, ConfigConstants.LocalServices.OCR_HOST_PORT)
                _listener.Start()
                WriteEndpointFile()
                RegisterOpenProtocol()
                WriteLog("OCR 服务已监听 127.0.0.1:" & ConfigConstants.LocalServices.OCR_HOST_PORT)
                Dim idle As New Thread(AddressOf IdleLoop)
                idle.IsBackground = True
                idle.Start()
                Do
                    Dim client As TcpClient = Nothing
                    Try
                        client = _listener.AcceptTcpClient()
                    Catch
                        Exit Do
                    End Try
                    Dim accepted = client
                    ThreadPool.QueueUserWorkItem(Sub() HandleClient(accepted))
                Loop
            Catch ex As Exception
                WriteLog("OCR 服务退出: " & ex.Message)
            Finally
                Cleanup()
                Try
                    mutex.ReleaseMutex()
                Catch
                End Try
            End Try
        End Sub

        Private Sub RunSelfTest()
            Dim doc As New OcrDocument()
            doc.RawText = "hi"
            doc.RawEngineResponse = "{""raw"":true}"
            Dim element As New OcrElement()
            element.Text = "a"
            element.BoundingBox = New OcrBoundingBox(1, 2, 3, 4)
            element.TableData = New OcrTableData()
            element.TableData.Rows.Add(New List(Of String) From {"cell"})
            element.TableData.RowSpanMap(Tuple.Create(0, 0)) = 2
            doc.Elements.Add(element)
            Dim json = JsonConvert.SerializeObject(doc, OcrHostClient.JsonSettings)
            Dim back = JsonConvert.DeserializeObject(Of OcrDocument)(json, OcrHostClient.JsonSettings)
            If back Is Nothing OrElse back.RawText <> "hi" OrElse back.Elements Is Nothing OrElse back.Elements.Count <> 1 Then
                Console.Error.WriteLine("ROUNDTRIP_FAIL")
                Environment.Exit(1)
            End If
            If back.Elements(0).BoundingBox Is Nothing OrElse back.Elements(0).BoundingBox.Right <> 3 Then
                Console.Error.WriteLine("ROUNDTRIP_FAIL_BOX")
                Environment.Exit(1)
            End If
            If back.RawEngineResponse <> doc.RawEngineResponse Then
                Console.Error.WriteLine("ROUNDTRIP_FAIL_RAW")
                Environment.Exit(1)
            End If
            Dim engines As New Dictionary(Of OcrEngineType, String)
            engines(OcrEngineType.RapidOCR) = "ok"
            Dim engineJson = JsonConvert.SerializeObject(engines, OcrHostClient.JsonSettings)
            Dim engineBack = JsonConvert.DeserializeObject(Of Dictionary(Of OcrEngineType, String))(engineJson, OcrHostClient.JsonSettings)
            If engineBack Is Nothing OrElse Not engineBack.ContainsKey(OcrEngineType.RapidOCR) Then
                Console.Error.WriteLine("ROUNDTRIP_FAIL_ENGINE")
                Environment.Exit(1)
            End If
            Console.WriteLine("ROUNDTRIP_OK")
            Environment.Exit(0)
        End Sub

        Private Sub IdleLoop()
            Do
                Thread.Sleep(30000)
                If _activeCalls = 0 AndAlso (DateTime.UtcNow - _lastActivityUtc).TotalMinutes >= IdleMinutes Then
                    WriteLog("空闲 " & IdleMinutes & " 分钟，OCR 服务退出")
                    Try
                        _listener.Stop()
                    Catch
                    End Try
                    Exit Do
                End If
            Loop
        End Sub

        Private Sub HandleClient(client As TcpClient)
            Try
                client.ReceiveTimeout = 900000
                client.SendTimeout = 900000
                Using client
                    Using stream = client.GetStream()
                        Dim headerText = ReadHeaders(stream)
                        If String.IsNullOrEmpty(headerText) Then Return
                        Dim lines = headerText.Split({vbCrLf}, StringSplitOptions.None)
                        Dim requestLine = lines(0)
                        Dim parts = requestLine.Split(" "c)
                        If parts.Length < 2 Then
                            WriteResponse(stream, 400, "{""ok"":false,""error"":""bad request""}")
                            Return
                        End If
                        Dim method = parts(0)
                        Dim requestPath = parts(1)
                        Dim rawQuery = ""
                        Dim query = requestPath.IndexOf("?"c)
                        If query >= 0 Then
                            rawQuery = requestPath.Substring(query + 1)
                            requestPath = requestPath.Substring(0, query)
                        End If
                        Dim contentLength As Integer = 0
                        Dim tokenOk As Boolean = False
                        Dim origin As String = Nothing
                        For Each headerLine In lines.Skip(1)
                            If headerLine.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase) Then
                                Integer.TryParse(headerLine.Substring(15).Trim(), contentLength)
                            ElseIf headerLine.StartsWith("X-Ocr-Token:", StringComparison.OrdinalIgnoreCase) Then
                                tokenOk = TokensEqual(headerLine.Substring(12).Trim(), _token)
                            ElseIf headerLine.StartsWith("Origin:", StringComparison.OrdinalIgnoreCase) Then
                                origin = headerLine.Substring(7).Trim()
                            End If
                        Next
                        If Not OriginAllowed(origin) Then
                            WriteResponse(stream, 403, "{""ok"":false,""error"":""origin not allowed""}", Nothing, origin)
                            Return
                        End If
                        If method = "OPTIONS" Then
                            WriteResponse(stream, 204, "", Nothing, origin)
                            Return
                        End If
                        _lastActivityUtc = DateTime.UtcNow
                        If method = "GET" AndAlso (requestPath = "/" OrElse requestPath = "/ocr-panel.html" OrElse requestPath = "/ocr-settings-panel.html" OrElse requestPath = "/gdt-symbol-recognizer.vb" OrElse requestPath.StartsWith("/js/") OrElse requestPath.StartsWith("/css/") OrElse requestPath.StartsWith("/lib/") OrElse requestPath.StartsWith("/config/") OrElse requestPath.StartsWith("/images/") OrElse requestPath.StartsWith("/prompts/")) Then
                            If requestPath = "/" Then requestPath = "/ocr-panel.html"
                            If ServeStatic(stream, requestPath, origin) Then Return
                            WriteResponse(stream, 404, "{""ok"":false,""error"":""not found""}", Nothing, origin)
                            Return
                        End If
                        Dim body = ""
                        If contentLength > 0 Then
                            If contentLength > 64 * 1024 * 1024 Then
                                WriteResponse(stream, 413, "{""ok"":false,""error"":""payload too large""}", Nothing, origin)
                                Return
                            End If
                            Dim buffer(contentLength - 1) As Byte
                            Dim readTotal = 0
                            While readTotal < contentLength
                                Dim n = stream.Read(buffer, readTotal, contentLength - readTotal)
                                If n <= 0 Then Exit While
                                readTotal += n
                            End While
                            body = Encoding.UTF8.GetString(buffer, 0, readTotal)
                        End If

                        If method = "GET" AndAlso requestPath = "/health" Then
                            WriteResponse(stream, 200, "{""ok"":true,""pid"":" & ProcessId() & "}", Nothing, origin)
                            Return
                        End If
                        If method = "GET" AndAlso requestPath = "/v1/screenshots" Then
                            Dim after = 0
                            Integer.TryParse(QueryValue(rawQuery, "after"), after)
                            Dim images = ScreenCaptureSession.ShotsAfter(after)
                            Dim arr As New JArray()
                            For Each img In images
                                arr.Add(img)
                            Next
                            Dim payload As New JObject()
                            payload("ok") = True
                            payload("total") = ScreenCaptureSession.LiveCount()
                            payload("images") = arr
                            WriteResponse(stream, 200, payload.ToString(Formatting.None), "application/json", origin)
                            Return
                        End If
                        Dim browserCaller = Not String.IsNullOrEmpty(origin) AndAlso OriginAllowed(origin)
                        If Not tokenOk AndAlso Not browserCaller Then
                            WriteResponse(stream, 401, "{""ok"":false,""error"":""unauthorized""}", Nothing, origin)
                            Return
                        End If
                        If method = "POST" AndAlso requestPath = "/v1/call" Then
                            Interlocked.Increment(_activeCalls)
                            Try
                                Dim response = Dispatch(body)
                                WriteResponse(stream, 200, response, Nothing, origin)
                            Finally
                                Interlocked.Decrement(_activeCalls)
                                _lastActivityUtc = DateTime.UtcNow
                            End Try
                            Return
                        End If
                        WriteResponse(stream, 404, "{""ok"":false,""error"":""not found""}", Nothing, origin)
                    End Using
                End Using
            Catch ex As Exception
                WriteLog("请求失败: " & ex.Message)
            End Try
        End Sub

        Private Function Dispatch(body As String) As String
            Try
                Dim root = JObject.Parse(If(body, "{}"))
                Dim fullName = root.Value(Of String)("method")
                If String.IsNullOrEmpty(fullName) Then
                    Return "{""ok"":false,""error"":""missing method""}"
                End If
                Dim methodName = fullName
                Dim dot = fullName.LastIndexOf("."c)
                If dot >= 0 Then methodName = fullName.Substring(dot + 1)
                If Not _allowed.Contains(methodName) Then
                    Return "{""ok"":false,""error"":""method not allowed""}"
                End If
                If methodName = "CaptureScreenshots" Then
                    Dim captured = ScreenCaptureSession.CaptureScreenshots()
                    Dim shotEnvelope As New JObject()
                    shotEnvelope("ok") = True
                    shotEnvelope("result") = JToken.FromObject(captured)
                    Return shotEnvelope.ToString(Formatting.None)
                End If
                Dim targetType As Type = GetType(OcrServiceManager)
                Dim target As Object = OcrServiceManager.Instance
                If fullName.StartsWith("OcrBridge.", StringComparison.Ordinal) Then
                    targetType = GetType(OcrBridge)
                    target = OcrBridge.Instance
                End If
                Dim argToken = root("args")
                Dim argCount = 0
                If TypeOf argToken Is JArray Then argCount = DirectCast(argToken, JArray).Count
                Dim targetMethod = FindMethod(targetType, methodName, argCount)
                If targetMethod Is Nothing Then
                    Return "{""ok"":false,""error"":""method not found: " & EscapeJson(methodName) & """}"
                End If
                Dim parameters = targetMethod.GetParameters()
                Dim invokeArgs As Object() = Nothing
                If parameters.Length > 0 Then
                    ReDim invokeArgs(parameters.Length - 1)
                End If
                For i = 0 To parameters.Length - 1
                    Dim paramType = parameters(i).ParameterType
                    If argToken Is Nothing OrElse i >= argCount OrElse DirectCast(argToken, JArray)(i).Type = JTokenType.Null Then
                        If parameters(i).IsOptional Then
                            invokeArgs(i) = If(parameters(i).DefaultValue Is DBNull.Value, Nothing, parameters(i).DefaultValue)
                        ElseIf paramType.IsValueType AndAlso Nullable.GetUnderlyingType(paramType) Is Nothing Then
                            invokeArgs(i) = Activator.CreateInstance(paramType)
                        Else
                            invokeArgs(i) = Nothing
                        End If
                    Else
                        invokeArgs(i) = DirectCast(argToken, JArray)(i).ToObject(paramType, JsonSerializer.Create(OcrHostClient.JsonSettings))
                    End If
                Next
                Dim raw As Object
                SyncLock _callLock
                    raw = targetMethod.Invoke(target, invokeArgs)
                End SyncLock
                Dim value = Unwrap(raw)
                Dim envelope As New JObject()
                envelope("ok") = True
                envelope("result") = If(value Is Nothing, JValue.CreateNull(), JToken.FromObject(value, JsonSerializer.Create(OcrHostClient.JsonSettings)))
                Return envelope.ToString(Formatting.None)
            Catch ex As TargetInvocationException
                Dim inner = If(ex.InnerException, ex)
                Return "{""ok"":false,""error"":""" & EscapeJson(inner.Message) & """}"
            Catch ex As Exception
                Return "{""ok"":false,""error"":""" & EscapeJson(ex.Message) & """}"
            End Try
        End Function

        Private Function FindMethod(targetType As Type, name As String, argCount As Integer) As MethodInfo
            Dim methods = targetType.GetMethods(BindingFlags.Instance Or BindingFlags.Public)
            Dim exact = methods.Where(Function(m) m.Name = name AndAlso m.GetParameters().Length = argCount).ToList()
            If exact.Count = 1 Then Return exact(0)
            Dim optionalMatch = methods.Where(Function(m)
                                                    If m.Name <> name Then Return False
                                                    Dim ps = m.GetParameters()
                                                    If ps.Length < argCount Then Return False
                                                    For i = argCount To ps.Length - 1
                                                        If Not ps(i).IsOptional Then Return False
                                                    Next
                                                    Return True
                                                End Function).OrderBy(Function(m) m.GetParameters().Length).ToList()
            If optionalMatch.Count > 0 Then Return optionalMatch(0)
            Return Nothing
        End Function

        Private Function Unwrap(result As Object) As Object
            Dim task = TryCast(result, Threading.Tasks.Task)
            If task Is Nothing Then Return result
            task.GetAwaiter().GetResult()
            Dim resultType = result.GetType()
            If resultType.IsGenericType Then
                Return resultType.GetProperty("Result").GetValue(result, Nothing)
            End If
            Return Nothing
        End Function

        Private Function ReadHeaders(stream As NetworkStream) As String
            Dim buffer As New List(Of Byte)
            Dim one(0) As Byte
            Do
                Dim n = stream.Read(one, 0, 1)
                If n <= 0 Then Exit Do
                buffer.Add(one(0))
                If buffer.Count >= 4 Then
                    Dim c = buffer.Count
                    If buffer(c - 4) = 13 AndAlso buffer(c - 3) = 10 AndAlso buffer(c - 2) = 13 AndAlso buffer(c - 1) = 10 Then
                        Exit Do
                    End If
                End If
                If buffer.Count > 65536 Then Exit Do
            Loop
            Return Encoding.ASCII.GetString(buffer.ToArray())
        End Function

        Private Function OriginAllowed(origin As String) As Boolean
            If String.IsNullOrEmpty(origin) OrElse origin = "null" Then Return True
            Dim port = ConfigConstants.LocalServices.OCR_HOST_PORT.ToString()
            Return origin.StartsWith("http://127.0.0.1:" & port, StringComparison.OrdinalIgnoreCase) OrElse
                origin.StartsWith("http://localhost:" & port, StringComparison.OrdinalIgnoreCase)
        End Function

        Private Function CorsHeaders(origin As String) As String
            If Not OriginAllowed(origin) Then Return ""
            Dim allowOrigin = origin
            If String.IsNullOrEmpty(allowOrigin) Then allowOrigin = "null"
            Return "Access-Control-Allow-Origin: " & allowOrigin & vbCrLf &
                "Access-Control-Allow-Methods: GET, POST, OPTIONS" & vbCrLf &
                "Access-Control-Allow-Headers: Content-Type, X-Ocr-Token" & vbCrLf &
                "Access-Control-Allow-Private-Network: true" & vbCrLf &
                "Vary: Origin" & vbCrLf
        End Function

        Private Function UiRoot() As String
            Dim candidates As New List(Of String)
            Dim plugin = Environment.GetEnvironmentVariable("SMARTPILOT_PLUGIN_DIR")
            If Not String.IsNullOrEmpty(plugin) Then candidates.Add(Path.Combine(plugin, "UI"))
            candidates.Add(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UI"))
            Dim dir = AppDomain.CurrentDomain.BaseDirectory
            For i = 1 To 6
                dir = Path.GetDirectoryName(If(dir, "").TrimEnd("\"c, "/"c))
                If String.IsNullOrEmpty(dir) Then Exit For
                candidates.Add(Path.Combine(dir, "UI"))
            Next
            Dim best As String = Nothing
            Dim bestTime = DateTime.MinValue
            For Each candidate In candidates
                Dim panel = Path.Combine(candidate, "ocr-panel.html")
                If Not File.Exists(panel) Then Continue For
                Dim written = File.GetLastWriteTimeUtc(panel)
                If best Is Nothing OrElse written > bestTime Then
                    best = candidate
                    bestTime = written
                End If
            Next
            Return best
        End Function

        Private Function ServeStatic(stream As NetworkStream, requestPath As String, origin As String) As Boolean
            Dim root = UiRoot()
            If String.IsNullOrEmpty(root) Then Return False
            Dim relative = requestPath.TrimStart("/"c).Replace("/"c, Path.DirectorySeparatorChar)
            If relative.Contains("..") Then Return False
            Dim full = Path.GetFullPath(Path.Combine(root, relative))
            Dim rootFull = Path.GetFullPath(root)
            If Not full.StartsWith(rootFull, StringComparison.OrdinalIgnoreCase) Then Return False
            If Not File.Exists(full) Then Return False
            Dim ext = Path.GetExtension(full).ToLowerInvariant()
            Dim contentType = "application/octet-stream"
            Select Case ext
                Case ".html" : contentType = "text/html; charset=utf-8"
                Case ".js" : contentType = "application/javascript; charset=utf-8"
                Case ".css" : contentType = "text/css; charset=utf-8"
                Case ".json" : contentType = "application/json; charset=utf-8"
                Case ".png" : contentType = "image/png"
                Case ".jpg", ".jpeg" : contentType = "image/jpeg"
                Case ".svg" : contentType = "image/svg+xml"
                Case ".gif" : contentType = "image/gif"
                Case ".woff" : contentType = "font/woff"
                Case ".woff2" : contentType = "font/woff2"
                Case ".md", ".txt", ".vb" : contentType = "text/plain; charset=utf-8"
            End Select
            Dim payload = File.ReadAllBytes(full)
            WriteRaw(stream, 200, payload, contentType, origin)
            Return True
        End Function

        Private Function QueryValue(query As String, name As String) As String
            If String.IsNullOrEmpty(query) OrElse String.IsNullOrEmpty(name) Then Return ""
            For Each part In query.Split("&"c)
                Dim eq = part.IndexOf("="c)
                Dim key = If(eq >= 0, part.Substring(0, eq), part)
                If String.Equals(Uri.UnescapeDataString(key), name, StringComparison.OrdinalIgnoreCase) Then
                    Return If(eq >= 0, Uri.UnescapeDataString(part.Substring(eq + 1)), "")
                End If
            Next
            Return ""
        End Function

        Private Sub WriteResponse(stream As NetworkStream, status As Integer, body As String, Optional contentType As String = Nothing, Optional origin As String = Nothing)
            WriteRaw(stream, status, Encoding.UTF8.GetBytes(If(body, "")), If(contentType, "application/json; charset=utf-8"), origin)
        End Sub

        Private Sub WriteRaw(stream As NetworkStream, status As Integer, payload As Byte(), contentType As String, origin As String)
            Dim reason = "Error"
            Select Case status
                Case 200 : reason = "OK"
                Case 204 : reason = "No Content"
                Case 401 : reason = "Unauthorized"
                Case 403 : reason = "Forbidden"
                Case 404 : reason = "Not Found"
                Case 413 : reason = "Payload Too Large"
            End Select
            Dim cacheControl = ""
            If contentType.StartsWith("text/", StringComparison.OrdinalIgnoreCase) OrElse
                contentType.StartsWith("application/javascript", StringComparison.OrdinalIgnoreCase) OrElse
                contentType.StartsWith("application/json", StringComparison.OrdinalIgnoreCase) Then
                cacheControl = "Cache-Control: no-cache" & vbCrLf
            End If
            Dim header = "HTTP/1.1 " & status & " " & reason & vbCrLf &
                "Content-Type: " & contentType & vbCrLf &
                "Content-Length: " & payload.Length & vbCrLf &
                cacheControl &
                CorsHeaders(origin) &
                "Connection: close" & vbCrLf & vbCrLf
            Dim headerBytes = Encoding.ASCII.GetBytes(header)
            stream.Write(headerBytes, 0, headerBytes.Length)
            If payload.Length > 0 Then stream.Write(payload, 0, payload.Length)
            stream.Flush()
        End Sub

        Private Sub RegisterOpenProtocol()
            Try
                Dim ui = UiRoot()
                If String.IsNullOrEmpty(ui) Then Return
                Dim launcher = Path.Combine(ui, "open-ocr-panel.vbs")
                If Not File.Exists(launcher) Then Return
                Using key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\Classes\smartpilot-ocr")
                    key.SetValue("", "URL:SmartPilot OCR")
                    key.SetValue("URL Protocol", "")
                End Using
                Using key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey("Software\Classes\smartpilot-ocr\shell\open\command")
                    key.SetValue("", "wscript.exe """ & launcher & """ ""%1""")
                End Using
            Catch ex As Exception
                WriteLog("注册打开协议失败: " & ex.Message)
            End Try
        End Sub

        Private Sub WriteEndpointFile()
            Dim dir = Path.GetDirectoryName(OcrHostClient.EndpointFilePath)
            If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
            Dim json = JsonConvert.SerializeObject(New With {
                .port = ConfigConstants.LocalServices.OCR_HOST_PORT,
                .token = _token,
                .pid = ProcessId()
            })
            File.WriteAllText(OcrHostClient.EndpointFilePath, json, Encoding.UTF8)
        End Sub

        Private Function ProcessId() As Integer
            Return Process.GetCurrentProcess().Id
        End Function

        Private Sub Cleanup()
            If Interlocked.Exchange(_cleaned, 1) = 1 Then Return
            Try
                OcrServiceManager.Instance.Dispose()
            Catch
            End Try
            Try
                If File.Exists(OcrHostClient.EndpointFilePath) Then
                    Dim text = File.ReadAllText(OcrHostClient.EndpointFilePath, Encoding.UTF8)
                    If text.Contains("""" & "pid""" & """:" & ProcessId()) OrElse text.Contains("""" & "pid""" & """: " & ProcessId()) Then
                        File.Delete(OcrHostClient.EndpointFilePath)
                    End If
                End If
            Catch
            End Try
            Try
                _listener.Stop()
            Catch
            End Try
        End Sub

        Private Function TokensEqual(left As String, right As String) As Boolean
            If left Is Nothing OrElse right Is Nothing OrElse left.Length <> right.Length Then Return False
            Dim diff = 0
            For i = 0 To left.Length - 1
                diff = diff Or (AscW(left(i)) Xor AscW(right(i)))
            Next
            Return diff = 0
        End Function

        Private Function EscapeJson(value As String) As String
            If value Is Nothing Then Return ""
            Return value.Replace("\", "\\").Replace("""", "\""").Replace(vbCr, "").Replace(vbLf, " ")
        End Function

        Private Sub WriteLog(message As String)
            Try
                Dim logPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ExcelAIVSTOAddIn",
                    "ocr-host.log")
                Dim dir = Path.GetDirectoryName(logPath)
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " " & message & vbCrLf, Encoding.UTF8)
            Catch
            End Try
        End Sub

End Module
