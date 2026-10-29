Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Text
Imports System.Threading.Tasks
Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 云端 OCR API 服务 - 支持多个云服务商
    ''' 识别准确率高，但需要联网和 API Key
    ''' </summary>
    Public Class CloudOcrService
        Implements IDisposable

        Private _httpClient As HttpClient
        Private _config As CloudOcrConfig
        Private _isDisposed As Boolean = False

        Public Sub New()
            _httpClient = New HttpClient()
            _httpClient.Timeout = TimeSpan.FromSeconds(30)
            _config = CloudOcrConfig.Load()
            _isDisposed = False
        End Sub

        ''' <summary>
        ''' 释放资源
        ''' </summary>
        Public Sub Dispose() Implements IDisposable.Dispose
            DisposeCore(True)
            GC.SuppressFinalize(Me)
        End Sub

        Protected Overridable Sub DisposeCore(disposing As Boolean)
            If _isDisposed Then Return
            _isDisposed = True

            If disposing Then
                If _httpClient IsNot Nothing Then
                    _httpClient.Dispose()
                    _httpClient = Nothing
                End If
            End If
        End Sub

        ''' <summary>
        ''' 检查云端 OCR 是否可用
        ''' </summary>
        Public Function IsAvailable() As Boolean
            If String.IsNullOrEmpty(_config.Provider) Then
                Return False
            End If

            Select Case _config.Provider.ToLower()
                Case "baidu"
                    Return Not String.IsNullOrEmpty(_config.BaiduApiKey) AndAlso Not String.IsNullOrEmpty(_config.BaiduSecretKey)
                Case "tencent"
                    Return Not String.IsNullOrEmpty(_config.TencentSecretId) AndAlso Not String.IsNullOrEmpty(_config.TencentSecretKey)
                Case Else
                    Return False
            End Select
        End Function

        ''' <summary>
        ''' 获取配置
        ''' </summary>
        Public Function GetConfig() As CloudOcrConfig
            Return _config
        End Function

        ''' <summary>
        ''' 保存配置
        ''' </summary>
        Public Sub SaveConfig(config As CloudOcrConfig)
            _config = config
            config.Save()
        End Sub

        ''' <summary>
        ''' 从文件重新加载配置（用于同步外部保存的配置）
        ''' </summary>
        Public Sub ReloadConfig()
            _config = CloudOcrConfig.Load()
        End Sub

        ''' <summary>
        ''' 获取支持的提供商列表
        ''' </summary>
        Public Function GetProviders() As Dictionary(Of String, String)
            Return New Dictionary(Of String, String) From {
                {"baidu", "百度智能云 OCR"},
                {"tencent", "腾讯云 OCR"}
            }
        End Function

        ''' <summary>
        ''' 验证云端 OCR 凭据是否有效（使用传入的临时配置进行真实 API 调用）
        ''' </summary>
        Public Async Function ValidateCredentialsAsync(provider As String, apiKey As String, secretKey As String) As Task(Of Boolean)
            Try
                Select Case provider.ToLower()
                    Case "baidu"
                        Return Await ValidateBaiduCredentials(apiKey, secretKey)
                    Case "tencent"
                        Return Await ValidateTencentCredentials(apiKey, secretKey)
                    Case Else
                        Return False
                End Select
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[ValidateCredentialsAsync] 验证失败: {ex.Message}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 验证百度智能云凭据
        ''' </summary>
        Private Async Function ValidateBaiduCredentials(apiKey As String, secretKey As String) As Task(Of Boolean)
            If String.IsNullOrEmpty(apiKey) OrElse String.IsNullOrEmpty(secretKey) Then
                Return False
            End If

            Dim url = "https://aip.baidubce.com/oauth/2.0/token"

            Dim formData = $"grant_type=client_credentials&client_id={Uri.EscapeDataString(apiKey)}&client_secret={Uri.EscapeDataString(secretKey)}"
            Dim bodyBytes = System.Text.Encoding.UTF8.GetBytes(formData)
            Dim content = New ByteArrayContent(bodyBytes)
            content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded")

            Dim response = Await _httpClient.PostAsync(url, content)
            Dim json = Await response.Content.ReadAsStringAsync()

            Dim obj = JObject.Parse(json)
            If obj("access_token") IsNot Nothing Then
                Return True
            End If

            System.Diagnostics.Debug.WriteLine($"[ValidateBaiduCredentials] 失败: {json}")
            Return False
        End Function

        ''' <summary>
        ''' 验证腾讯云凭据（通过调用轻量级 API）
        ''' </summary>
        Private Async Function ValidateTencentCredentials(secretId As String, secretKey As String) As Task(Of Boolean)
            If String.IsNullOrEmpty(secretId) OrElse String.IsNullOrEmpty(secretKey) Then
                Return False
            End If

            Dim url = "https://ocr.tencentcloudapi.com/"
            Dim timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString()
            Dim jsonContent = "{}"
            Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

            content.Headers.Add("Authorization", GetTencentSignatureWithKeys("GeneralBasicOCR", timestamp, jsonContent, secretId, secretKey))
            content.Headers.Add("X-TC-Action", "GeneralBasicOCR")
            content.Headers.Add("X-TC-Version", "2018-11-19")
            content.Headers.Add("X-TC-Timestamp", timestamp)
            content.Headers.Add("X-TC-Region", "ap-guangzhou")

            Dim response = Await _httpClient.PostAsync(url, content)
            Dim json = Await response.Content.ReadAsStringAsync()

            Dim obj = JObject.Parse(json)
            If obj("Response") IsNot Nothing Then
                Dim responseObj = obj("Response")
                If responseObj("Error") Is Nothing Then
                    Return True
                End If
                Dim errorCode = If(responseObj("Error")("Code") IsNot Nothing, responseObj("Error")("Code").ToString(), "")
                If errorCode = "InvalidParameter" OrElse errorCode = "MissingParameter" Then
                    Return True
                End If
            End If

            System.Diagnostics.Debug.WriteLine($"[ValidateTencentCredentials] 失败: {json}")
            Return False
        End Function

        ''' <summary>
        ''' 获取腾讯云签名(TC3-HMAC-SHA256) - 使用传入的凭据
        ''' </summary>
        Private Function GetTencentSignatureWithKeys(action As String, timestamp As String, requestBody As String, secretId As String, secretKey As String) As String
            Try
                Dim service = "ocr"
                Dim host = service & ".tencentcloudapi.com"
                Dim algorithm = "TC3-HMAC-SHA256"
                Dim dateStr = DateTimeOffset.FromUnixTimeSeconds(Long.Parse(timestamp)).UtcDateTime.ToString("yyyy-MM-dd")

                Dim httpRequestMethod = "POST"
                Dim canonicalUri = "/"
                Dim canonicalQueryString = ""
                Dim canonicalHeaders = $"content-type:application/json{vbLf}host:{host}{vbLf}"
                Dim signedHeaders = "content-type;host"
                Dim hashedRequestPayload = HashSha256(requestBody)
                Dim canonicalRequest = $"{httpRequestMethod}{vbLf}{canonicalUri}{vbLf}{canonicalQueryString}{vbLf}{canonicalHeaders}{vbLf}{signedHeaders}{vbLf}{hashedRequestPayload}"

                Dim credentialScope = $"{dateStr}/{service}/tc3_request"
                Dim hashedCanonicalRequest = HashSha256(canonicalRequest)
                Dim stringToSign = $"{algorithm}{vbLf}{timestamp}{vbLf}{credentialScope}{vbLf}{hashedCanonicalRequest}"

                Dim secretDate = SignHmac($"TC3{secretKey}", dateStr)
                Dim secretService = SignHmac(secretDate, service)
                Dim secretSigning = SignHmac(secretService, "tc3_request")
                Dim signature = SignHmacHex(secretSigning, stringToSign)

                Return $"{algorithm} Credential={secretId}/{credentialScope}, SignedHeaders={signedHeaders}, Signature={signature}"
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"腾讯云签名失败: {ex.Message}")
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' 执行 OCR 识别
        ''' </summary>
        Public Async Function RecognizeText(imagePath As String) As Task(Of String)
            System.Diagnostics.Debug.WriteLine($"[CloudOCR] RecognizeText 入口，路径: {imagePath}")
            
            If Not File.Exists(imagePath) Then
                Return "【错误：图片文件不存在】"
            End If

            System.Diagnostics.Debug.WriteLine($"[CloudOCR] 配置 Provider: {_config.Provider}")
            System.Diagnostics.Debug.WriteLine($"[CloudOCR] BaiduApiKey: {If(String.IsNullOrEmpty(_config.BaiduApiKey), "空", "非空，长度:" & _config.BaiduApiKey.Length)}")
            System.Diagnostics.Debug.WriteLine($"[CloudOCR] BaiduSecretKey: {If(String.IsNullOrEmpty(_config.BaiduSecretKey), "空", "非空，长度:" & _config.BaiduSecretKey.Length)}")

            If Not IsAvailable() Then
                System.Diagnostics.Debug.WriteLine("[CloudOCR] IsAvailable 返回 False")
                Return "【错误：云端 OCR 未配置，请先在设置中配置 API Key】"
            End If

            System.Diagnostics.Debug.WriteLine("[CloudOCR] IsAvailable 返回 True，准备识别")

            Try
                Select Case _config.Provider.ToLower()
                    Case "baidu"
                        System.Diagnostics.Debug.WriteLine("[CloudOCR] 选择百度 OCR")
                        Return Await RecognizeWithBaidu(imagePath)
                    Case "tencent"
                        System.Diagnostics.Debug.WriteLine("[CloudOCR] 选择腾讯云 OCR")
                        Return Await RecognizeWithTencent(imagePath)
                    Case Else
                        Return "【错误：不支持的 OCR 提供商】"
                End Select
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("云端 OCR 异常: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("异常堆栈: " & ex.StackTrace)
                Return "【错误：云端 OCR 执行异常: " & ex.Message & "】"
            End Try
        End Function

#Region "百度智能云 OCR"

        ''' <summary>
        ''' 使用百度智能云 OCR
        ''' </summary>
        Private Async Function RecognizeWithBaidu(imagePath As String) As Task(Of String)
            Dim token = Await GetBaiduAccessToken()
            If String.IsNullOrEmpty(token) Then
                Return "【错误：百度 OCR 认证失败，请检查 API Key】"
            End If

            If token.Length > 500 Then
                System.Diagnostics.Debug.WriteLine($"[BaiduOCR] Token 长度异常: {token.Length}，可能是错误消息")
                Return "【错误：百度 OCR 认证失败，Token 获取异常】"
            End If
            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] Access Token 长度: {token.Length}")

            Dim imageBytes = File.ReadAllBytes(imagePath)
            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] 图片大小: {imageBytes.Length} bytes")

            If imageBytes.Length > 8 * 1024 * 1024 Then
                Return "【错误：图片过大（>8MB），请压缩后再试】"
            End If

            Dim base64Image = Convert.ToBase64String(imageBytes)
            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] Base64 长度: {base64Image.Length}")

            Dim url = $"https://aip.baidubce.com/rest/2.0/ocr/v1/general_basic?access_token={Uri.EscapeDataString(token)}"
            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] URL 长度: {url.Length}")

            If url.Length > 2000 Then
                System.Diagnostics.Debug.WriteLine($"[BaiduOCR] URL 过长: {url.Length}，可能是 Token 异常")
                Return "【错误：URL 过长，请检查配置】"
            End If

            ' 注意：图片 Base64 可能超过 65519 字符，Uri.EscapeDataString 会抛出
            ' "无效的 URI: URI 字符串太长"。改用 WebUtility.UrlEncode，无此长度限制。
            Dim formData = $"image={WebUtility.UrlEncode(base64Image)}&language_type=CHN_ENG"
            Dim bodyBytes = System.Text.Encoding.UTF8.GetBytes(formData)
            Dim content = New ByteArrayContent(bodyBytes)
            content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded")

            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] POST URL: {url}")
            System.Diagnostics.Debug.WriteLine($"[BaiduOCR] 请求体大小: {bodyBytes.Length} bytes")

            Dim response = Await _httpClient.PostAsync(url, content)
            Dim json = Await response.Content.ReadAsStringAsync()

            Dim obj = JObject.Parse(json)

            If obj("error_code") IsNot Nothing Then
                Dim errorMsg = obj("error_msg").ToString()
                Return "【错误：百度 OCR - " & errorMsg & "】"
            End If

            If obj("words_result") IsNot Nothing Then
                Dim lines As New List(Of String)
                For Each item In obj("words_result")
                    If item("words") IsNot Nothing Then
                        lines.Add(item("words").ToString())
                    End If
                Next

                If lines.Count = 0 Then
                    Return "【未识别到文字】"
                End If

                Return String.Join(Environment.NewLine, lines)
            End If

            Return "【错误：无法解析百度 OCR 响应】"
        End Function

        ''' <summary>
        ''' 获取百度 Access Token
        ''' </summary>
        Private Async Function GetBaiduAccessToken() As Task(Of String)
            Dim url = "https://aip.baidubce.com/oauth/2.0/token"

            Try
                Dim formData = $"grant_type=client_credentials&client_id={Uri.EscapeDataString(_config.BaiduApiKey)}&client_secret={Uri.EscapeDataString(_config.BaiduSecretKey)}"
                Dim bodyBytes = System.Text.Encoding.UTF8.GetBytes(formData)
                Dim content = New ByteArrayContent(bodyBytes)
                content.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/x-www-form-urlencoded")

                Dim response = Await _httpClient.PostAsync(url, content)
                
                If Not response.IsSuccessStatusCode Then
                    System.Diagnostics.Debug.WriteLine($"[BaiduOCR] Token 请求失败，状态码: {response.StatusCode}")
                    Return Nothing
                End If

                Dim json = Await response.Content.ReadAsStringAsync()
                
                If String.IsNullOrEmpty(json) Then
                    System.Diagnostics.Debug.WriteLine("[BaiduOCR] Token 响应为空")
                    Return Nothing
                End If

                Dim obj = JObject.Parse(json)
                If obj("access_token") IsNot Nothing Then
                    Dim token = obj("access_token").ToString()
                    System.Diagnostics.Debug.WriteLine($"[BaiduOCR] 获取 Token 成功，长度: {token.Length}")
                    Return token
                End If

                If obj("error_code") IsNot Nothing Then
                    System.Diagnostics.Debug.WriteLine($"[BaiduOCR] Token 获取失败: {obj("error_code")} - {obj("error_msg")}")
                End If

                Return Nothing
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[BaiduOCR] 获取 Token 异常: {ex.Message}")
                Return Nothing
            End Try
        End Function

#End Region

#Region "腾讯云 OCR"

        ''' <summary>
        ''' 使用腾讯云 OCR
        ''' </summary>
        Private Async Function RecognizeWithTencent(imagePath As String) As Task(Of String)
            Dim url = "https://ocr.tencentcloudapi.com/"

            Dim imageBytes = File.ReadAllBytes(imagePath)
            Dim base64Image = Convert.ToBase64String(imageBytes)

            Dim requestBody = New Dictionary(Of String, Object) From {
                {"ImageBase64", base64Image}
            }

            Dim timestamp = DateTimeOffset.UtcNow.ToUnixTimeSeconds().ToString()
            Dim jsonContent = JsonConvert.SerializeObject(requestBody)
            Dim content = New StringContent(jsonContent, Encoding.UTF8, "application/json")

            content.Headers.Add("Authorization", GetTencentSignature("GeneralBasicOCR", timestamp, jsonContent))
            content.Headers.Add("X-TC-Action", "GeneralBasicOCR")
            content.Headers.Add("X-TC-Version", "2018-11-19")
            content.Headers.Add("X-TC-Timestamp", timestamp)
            content.Headers.Add("X-TC-Region", "ap-guangzhou")

            Dim response = Await _httpClient.PostAsync(url, content)
            Dim json = Await response.Content.ReadAsStringAsync()

            Dim obj = JObject.Parse(json)

            If obj("Response") IsNot Nothing Then
                Dim responseObj = obj("Response")

                If responseObj("TextDetections") IsNot Nothing Then
                    Dim lines As New List(Of String)
                    For Each item In responseObj("TextDetections")
                        If item("DetectedText") IsNot Nothing Then
                            lines.Add(item("DetectedText").ToString())
                        End If
                    Next

                    If lines.Count = 0 Then
                        Return "【未识别到文字】"
                    End If

                    Return String.Join(Environment.NewLine, lines)
                End If

                If responseObj("Error") IsNot Nothing Then
                    Dim errorMsg = responseObj("Error")("Message").ToString()
                    Return "【错误：腾讯云 OCR - " & errorMsg & "】"
                End If
            End If

            Return "【错误：无法解析腾讯云 OCR 响应】"
        End Function

        ''' <summary>
        ''' 获取腾讯云签名(TC3-HMAC-SHA256)
        ''' </summary>
        Private Function GetTencentSignature(action As String, timestamp As String, requestBody As String) As String
            Try
                Dim service = "ocr"
                Dim host = service & ".tencentcloudapi.com"
                Dim algorithm = "TC3-HMAC-SHA256"
                Dim dateStr = DateTimeOffset.FromUnixTimeSeconds(Long.Parse(timestamp)).UtcDateTime.ToString("yyyy-MM-dd")
                        
                ' 步骤1: 拼接规范请求串
                Dim httpRequestMethod = "POST"
                Dim canonicalUri = "/"
                Dim canonicalQueryString = ""
                Dim canonicalHeaders = $"content-type:application/json{vbLf}host:{host}{vbLf}"
                Dim signedHeaders = "content-type;host"
                Dim hashedRequestPayload = HashSha256(requestBody)
                Dim canonicalRequest = $"{httpRequestMethod}{vbLf}{canonicalUri}{vbLf}{canonicalQueryString}{vbLf}{canonicalHeaders}{vbLf}{signedHeaders}{vbLf}{hashedRequestPayload}"
                        
                ' 步骤2: 拼接待签名字符串
                Dim credentialScope = $"{dateStr}/{service}/tc3_request"
                Dim hashedCanonicalRequest = HashSha256(canonicalRequest)
                Dim stringToSign = $"{algorithm}{vbLf}{timestamp}{vbLf}{credentialScope}{vbLf}{hashedCanonicalRequest}"
                        
                ' 步骤3: 计算签名
                Dim secretDate = SignHmac($"TC3{_config.TencentSecretKey}", dateStr)
                Dim secretService = SignHmac(secretDate, service)
                Dim secretSigning = SignHmac(secretService, "tc3_request")
                Dim signature = SignHmacHex(secretSigning, stringToSign)
                        
                ' 步骤4: 拼接Authorization
                Return $"{algorithm} Credential={_config.TencentSecretId}/{credentialScope}, SignedHeaders={signedHeaders}, Signature={signature}"
                        
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"腾讯云签名失败: {ex.Message}")
                Return ""
            End Try
        End Function
                
        ''' <summary>
        ''' SHA256哈希
        ''' </summary>
        Private Function HashSha256(data As String) As String
            Dim sha256 As SHA256 = SHA256.Create()
            Try
                Dim bytes = Encoding.UTF8.GetBytes(data)
                Dim hash = sha256.ComputeHash(bytes)
                Return BitConverter.ToString(hash).Replace("-", "").ToLower()
            Finally
                sha256.Dispose()
            End Try
        End Function
                
        ''' <summary>
        ''' HMAC-SHA256签名
        ''' </summary>
        Private Function SignHmac(key As String, data As String) As Byte()
            Using hmac = New HMACSHA256(Encoding.UTF8.GetBytes(key))
                Return hmac.ComputeHash(Encoding.UTF8.GetBytes(data))
            End Using
        End Function
        
        ''' <summary>
        ''' HMAC-SHA256签名(Byte[]密钥)
        ''' </summary>
        Private Function SignHmac(key As Byte(), data As String) As Byte()
            Using hmac = New HMACSHA256(key)
                Return hmac.ComputeHash(Encoding.UTF8.GetBytes(data))
            End Using
        End Function
        
        ''' <summary>
        ''' HMAC-SHA256签名(返回十六进制)
        ''' </summary>
        Private Function SignHmacHex(key As Byte(), data As String) As String
            Using hmac = New HMACSHA256(key)
                Dim hash = hmac.ComputeHash(Encoding.UTF8.GetBytes(data))
                Return BitConverter.ToString(hash).Replace("-", "").ToLower()
            End Using
        End Function

#End Region



    End Class

    ''' <summary>
    ''' 云端 OCR 配置
    ''' </summary>
    Public Class CloudOcrConfig

        Public Property Provider As String = "baidu"

        Public Property BaiduApiKey As String = ""
        Public Property BaiduSecretKey As String = ""

        Public Property TencentSecretId As String = ""
        Public Property TencentSecretKey As String = ""

        Private Shared ReadOnly ConfigPath As String = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelAIVSTOAddIn", "cloud_ocr_config.json")

        ''' <summary>
        ''' 加载配置
        ''' </summary>
        Public Shared Function Load() As CloudOcrConfig
            Try
                If File.Exists(ConfigPath) Then
                    Dim json = File.ReadAllText(ConfigPath)
                    Dim config = JsonConvert.DeserializeObject(Of CloudOcrConfig)(json)
                    
                    ' 自动解密API密钥
                    config.DecryptApiKeys()
                    
                    Return config
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"加载云端 OCR 配置失败: {ex.Message}")
            End Try

            Return New CloudOcrConfig()
        End Function

        ''' <summary>
        ''' 保存配置
        ''' </summary>
        Public Sub Save()
            Try
                Dim dir = Path.GetDirectoryName(ConfigPath)
                If Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                ' 创建副本用于序列化,避免修改原始对象
                Dim configForSave = New CloudOcrConfig With {
                    .Provider = Me.Provider,
                    .BaiduApiKey = Me.BaiduApiKey,
                    .BaiduSecretKey = Me.BaiduSecretKey,
                    .TencentSecretId = Me.TencentSecretId,
                    .TencentSecretKey = Me.TencentSecretKey
                }
                
                ' 加密API密钥后保存
                configForSave.EncryptApiKeys()

                Dim json = JsonConvert.SerializeObject(configForSave, Formatting.Indented)
                File.WriteAllText(ConfigPath, json)
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"保存云端 OCR 配置失败: {ex.Message}")
            End Try
        End Sub
        
        ''' <summary>
        ''' 加密所有API密钥
        ''' </summary>
        Private Sub EncryptApiKeys()
            If Not String.IsNullOrEmpty(BaiduApiKey) AndAlso Not Utils.EncryptionHelper.IsEncrypted(BaiduApiKey) Then
                BaiduApiKey = Utils.EncryptionHelper.Encrypt(BaiduApiKey)
            End If
            If Not String.IsNullOrEmpty(BaiduSecretKey) AndAlso Not Utils.EncryptionHelper.IsEncrypted(BaiduSecretKey) Then
                BaiduSecretKey = Utils.EncryptionHelper.Encrypt(BaiduSecretKey)
            End If
            If Not String.IsNullOrEmpty(TencentSecretId) AndAlso Not Utils.EncryptionHelper.IsEncrypted(TencentSecretId) Then
                TencentSecretId = Utils.EncryptionHelper.Encrypt(TencentSecretId)
            End If
            If Not String.IsNullOrEmpty(TencentSecretKey) AndAlso Not Utils.EncryptionHelper.IsEncrypted(TencentSecretKey) Then
                TencentSecretKey = Utils.EncryptionHelper.Encrypt(TencentSecretKey)
            End If
        End Sub
        
        ''' <summary>
        ''' 解密所有API密钥
        ''' </summary>
        Private Sub DecryptApiKeys()
            If Not String.IsNullOrEmpty(BaiduApiKey) AndAlso Utils.EncryptionHelper.IsEncrypted(BaiduApiKey) Then
                BaiduApiKey = Utils.EncryptionHelper.Decrypt(BaiduApiKey)
            End If
            If Not String.IsNullOrEmpty(BaiduSecretKey) AndAlso Utils.EncryptionHelper.IsEncrypted(BaiduSecretKey) Then
                BaiduSecretKey = Utils.EncryptionHelper.Decrypt(BaiduSecretKey)
            End If
            If Not String.IsNullOrEmpty(TencentSecretId) AndAlso Utils.EncryptionHelper.IsEncrypted(TencentSecretId) Then
                TencentSecretId = Utils.EncryptionHelper.Decrypt(TencentSecretId)
            End If
            If Not String.IsNullOrEmpty(TencentSecretKey) AndAlso Utils.EncryptionHelper.IsEncrypted(TencentSecretKey) Then
                TencentSecretKey = Utils.EncryptionHelper.Decrypt(TencentSecretKey)
            End If
        End Sub

    End Class

End Namespace
