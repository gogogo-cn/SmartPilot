Option Explicit On
Option Infer On

Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Utils

Namespace Bridges

    ''' <summary>
    ''' 席位订阅授权桥 — BYOK：只校验软件许可是否有效，不计量 Token，不上传 API Key。
    '''
    ''' 统一激活门禁：离线激活码（路线①）的唯一存储是 HostLicense 的共享会话
    ''' （%APPDATA%\SmartPilotOffice\license_session.json），Excel / Word / PPT /
    ''' 其它端共用同一份激活与同一份注销名单——任何一端激活即全端激活，
    ''' 任何一端注销退款即全端失效。本桥不再自建第二份存储；旧版 Excel 曾把
    ''' 激活记在 ConfigurationManager（LICENSE_ACTIVATION / LICENSE_DEACTIVATED），
    ''' 首次读取时由 EnsureSharedActivation 一次性搬迁进共享会话。
    ''' </summary>
    <ComVisible(True)>
    <ClassInterface(ClassInterfaceType.AutoDispatch)>
    Public Class LicenseBridge
        Implements IBridge

        Private Shared ReadOnly _lazyInstance As New Lazy(Of LicenseBridge)(
            Function() New LicenseBridge(),
            System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Private Shared ReadOnly _http As New Lazy(Of HttpClient)(Function() CreateClient())

        Private Const TOKEN_SETTING As String = "LICENSE_SESSION_TOKEN"
        Private Const BASE_URL_SETTING As String = "LICENSE_API_BASE_URL"
        Private Const DEFAULT_BASE As String = "http://localhost:9530"
        Private Const CACHE_SETTING As String = "LICENSE_ENTITLEMENT_CACHE"
        ' LICENSE_ACTIVATION 已登记进 ConfigurationManager._sensitiveKeys，SetSetting 会自动 DPAPI 加密
        Private Const ACTIVATION_SETTING As String = "LICENSE_ACTIVATION"
        ' 本机已注销（消 key）的激活码标识列表。只存 SHA256 前 16 位，不含激活码原文，
        ' 因此**不要**登记进 _sensitiveKeys（明文即可，体积小、便于排查）。
        ' ⚠️ 迁移说明：ACTIVATION_SETTING / DEACTIVATED_SETTING 是旧版 Excel 的私有存储，
        ' 现已降级为「迁移源」——首次读取时搬进 HostLicense 共享会话，之后全端统一。
        ' 不要再往这两个设置里写新的激活状态，否则会重新劈出双轨。
        Private Const DEACTIVATED_SETTING As String = "LICENSE_DEACTIVATED"

        Public Shared ReadOnly Property Instance As LicenseBridge
            Get
                Return _lazyInstance.Value
            End Get
        End Property

        Private Sub New()
        End Sub

        Private Shared Function CreateClient() As HttpClient
            Dim c As New HttpClient()
            c.Timeout = TimeSpan.FromSeconds(20)
            Return c
        End Function

        Public ReadOnly Property BridgeName As String Implements IBridge.BridgeName
            Get
                Return "LicenseBridge"
            End Get
        End Property

        Private Function BaseUrl() As String
            Try
                Dim sharedUrl = HostLicense.PeekBaseUrl()
                If Not String.IsNullOrWhiteSpace(sharedUrl) Then Return sharedUrl
                Dim cfg = ConfigurationManager.Instance
                If cfg IsNot Nothing Then
                    Dim u = cfg.GetSetting(BASE_URL_SETTING, "")
                    If Not String.IsNullOrWhiteSpace(u) Then
                        u = u.Trim().TrimEnd("/"c)
                        HostLicense.RememberBaseUrl(u)
                        Return u
                    End If
                End If
            Catch
            End Try
            Return DEFAULT_BASE
        End Function

        Private Function GetToken() As String
            Try
                If HostLicense.TokenIsAuthoritative() Then Return HostLicense.PeekToken()
                Dim cfg = ConfigurationManager.Instance
                If cfg Is Nothing Then Return ""
                Dim localToken = If(cfg.GetSetting(TOKEN_SETTING, ""), "")
                If Not String.IsNullOrEmpty(localToken) Then HostLicense.RememberToken(localToken)
                Return localToken
            Catch
                Return ""
            End Try
        End Function

        Private Sub SetToken(token As String)
            Try
                Dim value = If(token, "")
                HostLicense.RememberToken(value)
                Dim cfg = ConfigurationManager.Instance
                If cfg Is Nothing Then Return
                ' LICENSE_SESSION_TOKEN 含 token 关键字，SetSetting 会自动 DPAPI 加密
                cfg.SetSetting(TOKEN_SETTING, value)
            Catch
            End Try
        End Sub

        Private Sub CacheEntitlement(json As String)
            Try
                HostLicense.RememberEntitlement(If(json, ""))
                Dim cfg = ConfigurationManager.Instance
                If cfg Is Nothing Then Return
                cfg.SetSetting(CACHE_SETTING, If(json, ""))
            Catch
            End Try
        End Sub

        Private Function ReadCachedEntitlement() As String
            Try
                If HostLicense.TokenIsAuthoritative() Then Return HostLicense.PeekEntitlement()
                Dim cfg = ConfigurationManager.Instance
                If cfg Is Nothing Then Return ""
                Dim localCache = If(cfg.GetSetting(CACHE_SETTING, ""), "")
                If Not String.IsNullOrWhiteSpace(localCache) Then HostLicense.RememberEntitlement(localCache)
                Return localCache
            Catch
                Return ""
            End Try
        End Function

        Private Async Function SendJsonAsync(method As HttpMethod, relativePath As String, bodyObj As Object, Optional bearer As String = Nothing) As Task(Of JObject)
            Dim url = BaseUrl() & relativePath
            Using req As New HttpRequestMessage(method, url)
                If Not String.IsNullOrEmpty(bearer) Then
                    req.Headers.TryAddWithoutValidation("Authorization", "Bearer " & bearer)
                End If
                If bodyObj IsNot Nothing Then
                    Dim json = JsonConvert.SerializeObject(bodyObj)
                    req.Content = New StringContent(json, Encoding.UTF8, "application/json")
                End If
                Dim resp = Await _http.Value.SendAsync(req).ConfigureAwait(False)
                Dim text = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
                Dim parsed As JObject = Nothing
                Try
                    parsed = JObject.Parse(If(String.IsNullOrWhiteSpace(text), "{}", text))
                Catch
                    parsed = New JObject From {
                        {"success", False},
                        {"error", "Invalid JSON from license server (HTTP " & CInt(resp.StatusCode) & ")"}
                    }
                End Try
                If Not resp.IsSuccessStatusCode AndAlso parsed("success") Is Nothing Then
                    parsed("success") = False
                    If parsed("error") Is Nothing Then parsed("error") = "HTTP " & CInt(resp.StatusCode)
                End If
                Return parsed
            End Using
        End Function

        ''' <summary>注册并开始试用。body: { email|phone, password }</summary>
        <ComVisible(True)>
        Public Function Register(bodyJson As String) As String
            Return RegisterAsync(bodyJson).GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function RegisterAsync(bodyJson As String) As Task(Of String)
            Try
                Dim body = JObject.Parse(If(bodyJson, "{}"))
                Dim result = Await SendJsonAsync(HttpMethod.Post, "/auth/register", body).ConfigureAwait(False)
                ApplyAuthResult(result)
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>登录。body: { login|email|phone, password }</summary>
        <ComVisible(True)>
        Public Function Login(bodyJson As String) As String
            Return LoginAsync(bodyJson).GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function LoginAsync(bodyJson As String) As Task(Of String)
            Try
                Dim body = JObject.Parse(If(bodyJson, "{}"))
                Dim result = Await SendJsonAsync(HttpMethod.Post, "/auth/login", body).ConfigureAwait(False)
                ApplyAuthResult(result)
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        <ComVisible(True)>
        Public Function Logout() As String
            Return LogoutAsync().GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function LogoutAsync() As Task(Of String)
            Try
                Dim token = GetToken()
                If Not String.IsNullOrEmpty(token) Then
                    Await SendJsonAsync(HttpMethod.Post, "/auth/logout", New Dictionary(Of String, Object)(), token).ConfigureAwait(False)
                End If
                SetToken("")
                CacheEntitlement("")
                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                SetToken("")
                Return JsonConvert.SerializeObject(New With {.success = True, .warning = ex.Message})
            End Try
        End Function

        ''' <summary>拉取并缓存当前权益；离线时返回缓存。</summary>
        <ComVisible(True)>
        Public Function GetEntitlement() As String
            Return GetEntitlementAsync().GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function GetEntitlementAsync() As Task(Of String)
            Try
                Dim token = GetToken()
                If String.IsNullOrEmpty(token) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = True,
                        .loggedIn = False,
                        .entitlement = New With {
                            .plan = "none",
                            .status = "inactive",
                            .active = False,
                            .periodEnd = CType(Nothing, String),
                            .reason = "not_logged_in"
                        }
                    })
                End If
                Dim result = Await SendJsonAsync(HttpMethod.Get, "/me/entitlement", Nothing, token).ConfigureAwait(False)
                If result.Value(Of Boolean?)("success") = True Then
                    CacheEntitlement(result.ToString(Formatting.None))
                    result("loggedIn") = True
                    Return result.ToString(Formatting.None)
                End If
                Dim cached = ReadCachedEntitlement()
                If Not String.IsNullOrWhiteSpace(cached) Then
                    Dim c = JObject.Parse(cached)
                    c("success") = True
                    c("stale") = True
                    c("loggedIn") = True
                    Return c.ToString(Formatting.None)
                End If
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Dim cached = ReadCachedEntitlement()
                If Not String.IsNullOrWhiteSpace(cached) Then
                    Try
                        Dim c = JObject.Parse(cached)
                        c("success") = True
                        c("stale") = True
                        c("offline") = True
                        c("loggedIn") = True
                        Return c.ToString(Formatting.None)
                    Catch
                    End Try
                End If
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message, .loggedIn = Not String.IsNullOrEmpty(GetToken())})
            End Try
        End Function

        Private _gateRefreshBusy As Integer

        ''' <summary>
        ''' 发送消息会同步调用本方法。已有权益缓存时立刻返回，授权服务的 HTTP（超时 20 秒）改到后台刷新，
        ''' 避免授权服务慢或未启动时「发送」按钮一直没反应。
        ''' </summary>
        Private Sub RefreshEntitlementInBackground()
            If System.Threading.Interlocked.CompareExchange(_gateRefreshBusy, 1, 0) <> 0 Then Return
            Task.Run(Sub()
                         Try
                             GetEntitlementAsync().GetAwaiter().GetResult()
                         Catch
                         Finally
                             System.Threading.Interlocked.Exchange(_gateRefreshBusy, 0)
                         End Try
                     End Sub)
        End Sub

        Private Function TryCachedSubscriptionJson() As String
            If String.IsNullOrEmpty(GetToken()) Then Return Nothing
            Dim cached = ReadCachedEntitlement()
            If String.IsNullOrWhiteSpace(cached) Then Return Nothing
            Try
                Dim c = JObject.Parse(cached)
                Dim ent = TryCast(c("entitlement"), JObject)
                If ent Is Nothing AndAlso c("active") IsNot Nothing Then ent = c
                If ent Is Nothing Then Return Nothing
                Dim active = ent.Value(Of Boolean?)("active") = True
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .active = active,
                    .loggedIn = True,
                    .cached = True,
                    .entitlement = ent,
                    .error = If(active, CType(Nothing, String), If(ent.Value(Of String)("reason"), "inactive"))
                })
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>是否允许使用付费能力（试用或订阅在有效期内）。</summary>
        <ComVisible(True)>
        Public Function IsSubscriptionActive() As String
            Try
                Dim fast = TryCachedSubscriptionJson()
                If fast IsNot Nothing Then
                    RefreshEntitlementInBackground()
                    Return fast
                End If
                Dim raw = GetEntitlement()
                Dim jo = JObject.Parse(raw)
                If jo("loggedIn") Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .active = False,
                        .error = If(jo.Value(Of String)("error"), "entitlement unavailable")
                    })
                End If
                Dim ent = TryCast(jo("entitlement"), JObject)
                Dim active = ent IsNot Nothing AndAlso ent.Value(Of Boolean?)("active") = True
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .active = active,
                    .loggedIn = jo.Value(Of Boolean?)("loggedIn") = True,
                    .entitlement = ent,
                    .error = If(active, Nothing, If(ent IsNot Nothing, ent.Value(Of String)("reason"), "inactive"))
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .active = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>创建 Stripe / 国内结账并打开系统浏览器。body: { channel: stripe|cn, plan?: pro_monthly|pro_yearly }</summary>
        <ComVisible(True)>
        Public Function StartCheckout(bodyJson As String) As String
            Return StartCheckoutAsync(bodyJson).GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function StartCheckoutAsync(bodyJson As String) As Task(Of String)
            Try
                Dim token = GetToken()
                If String.IsNullOrEmpty(token) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "请先登录", .code = "not_logged_in"})
                End If
                Dim body = JObject.Parse(If(String.IsNullOrWhiteSpace(bodyJson), "{}", bodyJson))
                Dim channel = (If(body.Value(Of String)("channel"), "stripe")).Trim().ToLowerInvariant()
                Dim path = If(channel = "cn", "/checkout/cn", "/checkout/stripe")
                Dim result = Await SendJsonAsync(HttpMethod.Post, path, body, token).ConfigureAwait(False)
                Dim url = result.Value(Of String)("url")
                If result.Value(Of Boolean?)("success") = True AndAlso Not String.IsNullOrWhiteSpace(url) Then
                    Try
                        Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
                    Catch ex As Exception
                        result("openError") = ex.Message
                    End Try
                End If
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>打开 Stripe Customer Portal。</summary>
        <ComVisible(True)>
        Public Function OpenBillingPortal() As String
            Return OpenBillingPortalAsync().GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function OpenBillingPortalAsync() As Task(Of String)
            Try
                Dim token = GetToken()
                If String.IsNullOrEmpty(token) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "请先登录", .code = "not_logged_in"})
                End If
                Dim result = Await SendJsonAsync(HttpMethod.Post, "/checkout/stripe/portal", New Dictionary(Of String, Object)(), token).ConfigureAwait(False)
                Dim url = result.Value(Of String)("url")
                If result.Value(Of Boolean?)("success") = True AndAlso Not String.IsNullOrWhiteSpace(url) Then
                    Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
                End If
                Return result.ToString(Formatting.None)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>设置授权服务根地址，例如 https://license.example.com</summary>
        <ComVisible(True)>
        Public Function SetLicenseApiBaseUrl(url As String) As String
            Try
                Dim cfg = ConfigurationManager.Instance
                If cfg Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "ConfigurationManager 未初始化"})
                End If
                Dim normalized = If(url, "").Trim().TrimEnd("/"c)
                cfg.SetSetting(BASE_URL_SETTING, normalized)
                HostLicense.RememberBaseUrl(normalized)
                Return JsonConvert.SerializeObject(New With {.success = True, .baseUrl = BaseUrl()})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        <ComVisible(True)>
        Public Function GetLicenseApiBaseUrl() As String
            Return JsonConvert.SerializeObject(New With {.success = True, .baseUrl = BaseUrl()})
        End Function

        ''' <summary>
        ''' 由 .NET HttpClient 探测授权服务 /health。
        ''' WebView2 页面直连 localhost 常被拦截，不能用前端 fetch 判断服务是否可用。
        ''' 插件不负责启动授权服务。
        ''' </summary>
        <ComVisible(True)>
        Public Function PingHealth() As String
            Return PingHealthAsync().GetAwaiter().GetResult()
        End Function

        <ComVisible(True)>
        Public Async Function PingHealthAsync() As Task(Of String)
            Try
                Dim result = Await SendJsonAsync(HttpMethod.Get, "/health", Nothing).ConfigureAwait(False)
                Dim okFlag = result.Value(Of Boolean?)("success") = True
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .ok = okFlag,
                    .baseUrl = BaseUrl(),
                    .service = result.Value(Of String)("service"),
                    .stripeReady = result.Value(Of Boolean?)("stripeReady").GetValueOrDefault(False),
                    .cnReady = result.Value(Of Boolean?)("cnReady").GetValueOrDefault(False),
                    .error = If(okFlag, Nothing, result.Value(Of String)("error"))
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .ok = False,
                    .baseUrl = BaseUrl(),
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 离线激活码激活（路线①：人工收款后发 key，无需授权服务器在线）。
        ''' body: { key }。验签用本文件三端一致的算法，激活记录写进共享会话（DPAPI），
        ''' Word / PPT / 其它端立刻可见。
        ''' </summary>
        <ComVisible(True)>
        Public Function ActivateKey(bodyJson As String) As String
            Try
                Dim body = JObject.Parse(If(bodyJson, "{}"))
                Dim key = If(body.Value(Of String)("key"), "").Trim()
                Dim result = LicenseKey.Verify(key)
                If Not result.Valid Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = ActivationReasonText(result.Reason),
                        .code = "invalid_key"
                    })
                End If
                Dim cfg = ConfigurationManager.Instance
                EnsureSharedActivation(cfg)
                ' 共享注销名单：任何端注销过的 key，所有端都不能再激活
                If HostLicense.IsKeyInDeactivatedList(key) Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "此激活码已在本机注销（已申请退款），不能再次激活",
                        .code = "deactivated"
                    })
                End If
                HostLicense.RememberActivation(key)
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .active = True,
                    .source = "offline",
                    .entitlement = OfflineEntitlement(result)
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 本地离线激活状态（不打网络）。记录被篡改无所谓——每次从 key 重新验签。
        ''' 读的就是共享会话，与 Word / PPT / 其它端看到同一份激活。
        ''' </summary>
        <ComVisible(True)>
        Public Function GetLocalActivation() As String
            Try
                Dim cfg = ConfigurationManager.Instance
                EnsureSharedActivation(cfg)
                Return HostLicense.GetLocalActivation()
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>清除本地离线激活（换 key / 退款）。共享会话与旧版残留一并清掉。</summary>
        <ComVisible(True)>
        Public Function ClearLocalActivation() As String
            Try
                Dim cfg = ConfigurationManager.Instance
                EnsureSharedActivation(cfg)
                HostLicense.ClearLocalActivation()
                ForgetLegacyActivation(cfg)
                Return JsonConvert.SerializeObject(New With {.success = True})
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 「注销并申请退款」：清除本机激活 + 把该激活码永久记入本机已注销名单 + 返回消 key 码。
        '''
        ''' 用户把消 key 码发给客服，运营台验签通过后才放款。于是「退款」与「不能再使用」
        ''' 变成同一个动作的两面，不需要强制联网。
        ''' 关键：注销后同一张 key 在本机**不能再被激活**（ActivateKey / GetLocalActivation
        ''' 都会查共享注销名单），否则销完立刻重激活就等于没销。
        ''' 门禁统一：注销写进共享会话，Word / PPT / 其它端同步失效——退款不留跨端缺口。
        ''' </summary>
        <ComVisible(True)>
        Public Function DeactivateForRefund() As String
            Try
                Dim cfg = ConfigurationManager.Instance
                EnsureSharedActivation(cfg)
                Dim result = HostLicense.DeactivateForRefund()
                ForgetLegacyActivation(cfg)
                Return result
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 统一激活门禁的迁移入口：旧版 Excel 把激活/注销名单记在 ConfigurationManager，
        ''' 与 Word / PPT 的共享会话是两份独立存储（本机激活不互认、退款不互生效）。
        ''' 首次读取时把旧数据一次性搬进共享会话并标记权威，此后所有端同源。
        ''' 共享会话已权威时这是纯读，成本 = 一次小文件读。
        ''' </summary>
        Private Shared Sub EnsureSharedActivation(cfg As ConfigurationManager)
            If HostLicense.ActivationIsAuthoritative() Then Return
            Dim key As String = ""
            If cfg IsNot Nothing Then
                Try
                    Dim raw = If(cfg.GetSetting(ACTIVATION_SETTING, ""), "")
                    If Not String.IsNullOrWhiteSpace(raw) Then
                        Dim record As JObject = Nothing
                        Try
                            record = JObject.Parse(raw)
                        Catch
                            record = Nothing
                        End Try
                        If record IsNot Nothing Then key = If(record.Value(Of String)("key"), "").Trim()
                    End If
                Catch
                End Try
                Try
                    ' 旧注销名单并进共享名单：在 Excel 销过的 key，拿到 Word 也不能再激活
                    Dim legacy = ReadDeactivatedList(cfg)
                    If legacy.Count > 0 Then HostLicense.MergeDeactivatedList(legacy)
                Catch
                End Try
            End If
            ' 收尾把激活（可能为空）写进共享并标记权威。空也要写——否则 activationSet
            ' 永远为 False，每次门禁都会重跑这段迁移。
            HostLicense.RememberActivation(key)
        End Sub

        ''' <summary>
        ''' 清掉旧版 Excel 的私有激活记录（迁移已完成，留着只会让以后误迁移）。
        ''' 注销名单不清——它已在 EnsureSharedActivation 里并进共享名单。
        ''' </summary>
        Private Shared Sub ForgetLegacyActivation(cfg As ConfigurationManager)
            If cfg Is Nothing Then Return
            Try
                cfg.SetSetting(ACTIVATION_SETTING, "")
            Catch
            End Try
        End Sub

        ''' <summary>查询某张激活码是否已在本机注销过（激活流程用来拦截“销完再激活”）。查共享名单。</summary>
        <ComVisible(True)>
        Public Function IsKeyDeactivated(bodyJson As String) As String
            Try
                Dim cfg = ConfigurationManager.Instance
                EnsureSharedActivation(cfg)
                Return HostLicense.IsKeyDeactivated(bodyJson)
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>读取旧版 Excel 私有存储里的已注销标识列表（只存 SHA256 前 16 位）。仅迁移用。</summary>
        Private Shared Function ReadDeactivatedList(cfg As ConfigurationManager) As List(Of String)
            Dim result As New List(Of String)()
            If cfg Is Nothing Then Return result
            Dim raw = cfg.GetSetting(DEACTIVATED_SETTING, "")
            If String.IsNullOrWhiteSpace(raw) Then Return result
            For Each part In raw.Split(","c)
                Dim h = part.Trim()
                If h.Length > 0 Then result.Add(h)
            Next
            Return result
        End Function

        Private Shared Function OfflineEntitlement(result As LicenseKeyResult) As Object
            Return New With {
                .plan = result.Plan,
                .seats = result.Seats,
                .status = "active",
                .periodEnd = result.PeriodEnd.ToString("o"),
                .active = True,
                .source = "offline"
            }
        End Function

        Private Shared Function ActivationReasonText(reason As String) As String
            Select Case reason
                Case "expired" : Return "激活码已过期"
                Case "bad_signature" : Return "激活码无效（签名不匹配）"
                Case Else : Return "激活码格式不正确"
            End Select
        End Function

        Private Sub ApplyAuthResult(result As JObject)
            If result Is Nothing Then Return
            If result.Value(Of Boolean?)("success") <> True Then Return
            Dim token = result.Value(Of String)("token")
            If Not String.IsNullOrEmpty(token) Then SetToken(token)
            CacheEntitlement(result.ToString(Formatting.None))
        End Sub

        Private _disposed As Boolean = False
        Public Sub Dispose() Implements IBridge.Dispose
            Dispose(True)
            GC.SuppressFinalize(Me)
        End Sub
        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not _disposed Then
                _disposed = True
            End If
        End Sub
    End Class

End Namespace
