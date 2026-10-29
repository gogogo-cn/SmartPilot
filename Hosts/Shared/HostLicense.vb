' ============================================================================
' HostLicense - Word / PowerPoint / Excel 共用的席位授权客户端
'
' 会话写在 %APPDATA%\SmartPilotOffice\license_session.json。
' Excel 的 LicenseBridge 读写同一份文件，所以所有端共用一次登录、同一个激活门禁：
'   - 在线订阅（路线②）：token 共享；
'   - 离线激活码（路线①）：activation 共享，Excel 侧旧配置一次性迁移进来。
' 插件不启动授权服务，只访问已有的 /health 与 /auth、/me、/checkout。
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Module HostLicense

    Private ReadOnly SessionPath As String = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "SmartPilotOffice", "license_session.json")

    Private Const DEFAULT_BASE As String = "http://localhost:9530"

    ''' <summary>共享会话文件读写重试次数。三端 Office 可能同时读写同一份文件。</summary>
    Private Const SHARED_FILE_RETRY_TIMES As Integer = 8
    ''' <summary>共享会话文件读写重试间隔（毫秒）。</summary>
    Private Const SHARED_FILE_RETRY_DELAY_MS As Integer = 200

    ' ---- 离线激活码（路线①：无授权服务器也能卖） ----
    ' 算法与 Bridges/LicenseKey.vb、LicenseServer/lib/licenseKey.js 逐字节一致（改算法必须三处同步）：
    '   key = "SPK1-" & base64url("v1|login|plan|seats|expUnix|nonce") & "." & base64url(HMAC-SHA256)
    ' （第 6 段 nonce 为随机数，保证同参数批量生成的 key 互不相同；验签只计数不取值。）
    ' 激活记录写进 license_session.json 的 activationProtected 字段（DPAPI 保护，与 token 同款），
    ' 每次读取都从 key 重新验签，所以记录被篡改没有意义。
    ' 威胁模型：secret 内置于客户端，属轻度防护；强校验（吊销/换机）走在线链路。
    Private Const KEY_PREFIX As String = "SPK1-"
    ''' <summary>消 key 码（退款凭证）前缀，与激活码前缀严格隔离。</summary>
    Private Const DEACT_PREFIX As String = "SPD1-"
    Private Const KEY_VERSION As String = "v1"
    Private Const KEY_FIELD_COUNT As Integer = 6
    Private Const DEACT_FIELD_COUNT As Integer = 5
    Private Const KEY_SECRET As String = "SmartPilot-Offline-License-Key-v1-CHANGE-ME"
    Private ReadOnly KeyEpoch As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
    Private ReadOnly FileLock As New Object()
    Private ReadOnly Http As New Lazy(Of HttpClient)(Function()
                                                         Dim client As New HttpClient()
                                                         client.Timeout = TimeSpan.FromSeconds(20)
                                                         Return client
                                                     End Function)

    Public Function PeekToken() As String
        Return If(ReadSession().Value(Of String)("token"), "")
    End Function

    ''' <summary>共享会话里的 token 是否已经写过（含退出后的空 token）。未写过时 Excel 才从自己的配置迁移。</summary>
    Public Function TokenIsAuthoritative() As Boolean
        Return ReadSession().Value(Of Boolean?)("tokenSet").GetValueOrDefault(False)
    End Function

    Public Sub RememberToken(token As String)
        Dim session = ReadSession()
        session("token") = If(token, "")
        session("tokenSet") = True
        If String.IsNullOrEmpty(token) Then session("entitlement") = ""
        WriteSession(session)
    End Sub

    Public Function PeekBaseUrl() As String
        Dim url = ReadSession().Value(Of String)("baseUrl")
        If String.IsNullOrWhiteSpace(url) Then Return ""
        Return url.Trim().TrimEnd("/"c)
    End Function

    Public Sub RememberBaseUrl(url As String)
        Dim session = ReadSession()
        session("baseUrl") = If(url, "").Trim().TrimEnd("/"c)
        WriteSession(session)
    End Sub

    Public Function PeekEntitlement() As String
        Return If(ReadSession().Value(Of String)("entitlement"), "")
    End Function

    Public Sub RememberEntitlement(json As String)
        Dim session = ReadSession()
        session("entitlement") = If(json, "")
        WriteSession(session)
    End Sub

    Public Function Register(bodyJson As String) As String
        Return RegisterAsync(bodyJson).GetAwaiter().GetResult()
    End Function

    Public Function Login(bodyJson As String) As String
        Return LoginAsync(bodyJson).GetAwaiter().GetResult()
    End Function

    Public Function Logout() As String
        Return LogoutAsync().GetAwaiter().GetResult()
    End Function

    Public Function GetEntitlement() As String
        Return GetEntitlementAsync().GetAwaiter().GetResult()
    End Function

    Private GateRefreshBusy As Integer

    ''' <summary>已有权益缓存时立刻回答，授权 HTTP 改到后台，避免每次发送都堵住界面线程。</summary>
    Private Sub RefreshEntitlementInBackground()
        If System.Threading.Interlocked.CompareExchange(GateRefreshBusy, 1, 0) <> 0 Then Return
        System.Threading.Tasks.Task.Run(Sub()
                                            Try
                                                GetEntitlement()
                                            Catch
                                            Finally
                                                System.Threading.Interlocked.Exchange(GateRefreshBusy, 0)
                                            End Try
                                        End Sub)
    End Sub

    Private Function TryCachedSubscriptionJson() As String
        If String.IsNullOrEmpty(PeekToken()) Then Return Nothing
        Dim cached = PeekEntitlement()
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

    Public Function IsSubscriptionActive() As String
        Try
            Dim fast = TryCachedSubscriptionJson()
            If fast IsNot Nothing Then
                RefreshEntitlementInBackground()
                Return fast
            End If
            Dim jo = JObject.Parse(GetEntitlement())
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

    Public Function StartCheckout(bodyJson As String) As String
        Return StartCheckoutAsync(bodyJson).GetAwaiter().GetResult()
    End Function

    Public Function OpenBillingPortal() As String
        Return OpenBillingPortalAsync().GetAwaiter().GetResult()
    End Function

    Public Function SetApiBaseUrl(url As String) As String
        Try
            RememberBaseUrl(If(url, ""))
            Return JsonConvert.SerializeObject(New With {.success = True, .baseUrl = BaseUrl()})
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    Public Function GetApiBaseUrl() As String
        Return JsonConvert.SerializeObject(New With {.success = True, .baseUrl = BaseUrl()})
    End Function

    Public Function PingHealth() As String
        Return PingHealthAsync().GetAwaiter().GetResult()
    End Function

    ' ---- 离线激活码：激活 / 读取 / 清除 ----

    ''' <summary>离线激活码激活。body: { key }。记录写进 license_session.json（DPAPI）。</summary>
    Public Function ActivateKey(bodyJson As String) As String
        Try
            Dim body = JObject.Parse(If(bodyJson, "{}"))
            Dim key = If(body.Value(Of String)("key"), "").Trim()
            Dim result = VerifyKey(key)
            If Not result.Value(Of Boolean)("valid") Then
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ActivationReasonText(result.Value(Of String)("reason")),
                    .code = "invalid_key"
                })
            End If
            ' 已注销（消过 key 申请退款）的激活码不允许再次激活
            If IsDeactivated(key) Then
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = "此激活码已在本机注销（已申请退款），不能再次激活",
                    .code = "deactivated"
                })
            End If
            RememberActivation(key)
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .active = True,
                .source = "offline",
                .entitlement = New With {
                    .plan = result.Value(Of String)("plan"),
                    .seats = result.Value(Of Integer)("seats"),
                    .status = "active",
                    .periodEnd = result.Value(Of DateTime)("periodEnd").ToString("o"),
                    .active = True,
                    .source = "offline"
                }
            })
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    ''' <summary>本地离线激活状态（不打网络）。每次从 key 重新验签。</summary>
    Public Function GetLocalActivation() As String
        Try
            Dim session = ReadSession()
            Dim key = If(session.Value(Of String)("activation"), "")
            If String.IsNullOrWhiteSpace(key) Then Return InactiveLocal("not_activated")
            Dim result = VerifyKey(key)
            If Not result.Value(Of Boolean)("valid") Then
                RememberActivation("")
                Return InactiveLocal(result.Value(Of String)("reason"))
            End If
            ' 防「销完又手工恢复激活记录」：只要这张 key 在本机注销过，一律视为失效
            If IsDeactivated(key) Then
                RememberActivation("")
                Return InactiveLocal("deactivated")
            End If
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .active = True,
                .source = "offline",
                .login = result.Value(Of String)("login"),
                .entitlement = New With {
                    .plan = result.Value(Of String)("plan"),
                    .seats = result.Value(Of Integer)("seats"),
                    .status = "active",
                    .periodEnd = result.Value(Of DateTime)("periodEnd").ToString("o"),
                    .active = True,
                    .source = "offline"
                }
            })
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    ''' <summary>清除本地离线激活（换 key / 退款）。</summary>
    Public Function ClearLocalActivation() As String
        Try
            RememberActivation("")
            Return JsonConvert.SerializeObject(New With {.success = True})
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    ''' <summary>
    ''' 「注销并申请退款」：清除本机激活 + 把该激活码永久记入本机已注销名单 + 返回消 key 码。
    ''' 用户把消 key 码发给客服，运营台验签通过才放款，于是「退款」与「不能再使用」
    ''' 变成同一个动作的两面，不依赖强制联网。
    ''' 关键：注销后同一张 key 在本机不能再被激活，否则销完立刻重激活就等于没销。
    ''' </summary>
    Public Function DeactivateForRefund() As String
        Try
            Dim session = ReadSession()
            Dim key = If(session.Value(Of String)("activation"), "")
            If String.IsNullOrWhiteSpace(key) Then
                Return JsonConvert.SerializeObject(New With {
                    .success = False, .error = "当前没有离线激活记录，无需注销", .code = "not_activated"})
            End If
            Dim result = VerifyKey(key)
            Dim plan = If(result.Value(Of String)("plan"), "pro_yearly")
            Dim login = If(result.Value(Of String)("login"), "")

            ' 顺序要紧：先记入已注销名单，再清激活。反了会留出「注销后仍能再用」的窗口。
            MarkDeactivated(key)
            RememberActivation("")

            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .code = CreateDeactivationProof(key, plan),
                .plan = plan,
                .login = login,
                .destroyedAt = DateTime.UtcNow.ToString("o")
            })
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    ''' <summary>查询某张激活码是否已在本机注销过（激活流程用来拦截“销完再激活”）。</summary>
    Public Function IsKeyDeactivated(bodyJson As String) As String
        Try
            Dim body = JObject.Parse(If(bodyJson, "{}"))
            Dim key = If(body.Value(Of String)("key"), "").Trim()
            Dim h = KeyHash(key)
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .deactivated = ReadDeactivatedList().Contains(h)
            })
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    ''' <summary>本机已注销的激活码标识列表（只存 SHA256 前 16 位，无原文）。</summary>
    Private Function ReadDeactivatedList() As List(Of String)
        Dim result As New List(Of String)()
        Dim raw As String = If(ReadSession().Value(Of String)("deactivated"), "")
        If String.IsNullOrWhiteSpace(raw) Then Return result
        For Each part As String In raw.Split(","c)
            Dim h As String = part.Trim()
            If h.Length > 0 Then result.Add(h)
        Next
        Return result
    End Function

    Private Sub MarkDeactivated(key As String)
        Dim list = ReadDeactivatedList()
        Dim h = KeyHash(key)
        If Not list.Contains(h) Then list.Add(h)
        Dim session = ReadSession()
        session("deactivated") = String.Join(",", list)
        WriteSession(session)
    End Sub

    ''' <summary>该激活码是否已在本机注销过。</summary>
    Private Function IsDeactivated(key As String) As Boolean
        If String.IsNullOrEmpty(key) Then Return False
        Return ReadDeactivatedList().Contains(KeyHash(key))
    End Function

    ''' <summary>
    ''' 写入共享激活码（DPAPI 保护后落盘）。Excel / Word / PPT / 其它端共用这一份——
    ''' activationSet 一旦写过，共享会话即成为激活状态的唯一权威（见 ActivationIsAuthoritative）。
    ''' </summary>
    Public Sub RememberActivation(key As String)
        Dim session = ReadSession()
        session("activation") = If(key, "")
        session("activationSet") = True
        WriteSession(session)
    End Sub

    ''' <summary>
    ''' 共享会话是否已接管激活状态（写过 activationSet）。
    ''' 与 TokenIsAuthoritative 同款语义：为 False 说明本机还没被共享会话管过，
    ''' 调用方（Excel）应先把旧存储迁移进来。
    ''' </summary>
    Public Function ActivationIsAuthoritative() As Boolean
        Return ReadSession().Value(Of Boolean?)("activationSet").GetValueOrDefault(False)
    End Function

    ''' <summary>
    ''' 把外部（旧版 Excel ConfigurationManager）的已注销标识并进共享注销名单。
    ''' 任何端注销过的 key，所有端都不能再激活——退款闭环不能留跨端缺口。
    ''' </summary>
    Public Sub MergeDeactivatedList(hashes As IEnumerable(Of String))
        If hashes Is Nothing Then Return
        Dim list = ReadDeactivatedList()
        Dim changed As Boolean = False
        For Each h As String In hashes
            Dim t As String = If(h, "").Trim()
            If t.Length > 0 AndAlso Not list.Contains(t) Then
                list.Add(t)
                changed = True
            End If
        Next
        If Not changed Then Return
        Dim session = ReadSession()
        session("deactivated") = String.Join(",", list)
        WriteSession(session)
    End Sub

    ''' <summary>该激活码是否已在共享注销名单里（任何端注销都算）。</summary>
    Public Function IsKeyInDeactivatedList(key As String) As Boolean
        If String.IsNullOrEmpty(key) Then Return False
        Return ReadDeactivatedList().Contains(KeyHash(key))
    End Function

    Private Function InactiveLocal(reason As String) As String
        Return JsonConvert.SerializeObject(New With {
            .success = True,
            .active = False,
            .source = "offline",
            .reason = reason,
            .entitlement = New With {
                .plan = "none",
                .status = "inactive",
                .active = False,
                .periodEnd = CType(Nothing, String),
                .source = "offline",
                .reason = reason
            }
        })
    End Function

    Private Function ActivationReasonText(reason As String) As String
        Select Case reason
            Case "expired" : Return "激活码已过期"
            Case "bad_signature" : Return "激活码无效（签名不匹配）"
            Case Else : Return "激活码格式不正确"
        End Select
    End Function

    ''' <summary>
    ''' 验签离线激活码。返回 JObject：{ valid, reason, login, plan, seats, periodEnd }。
    ''' 与 Bridges/LicenseKey.vb 的 LicenseKey.Verify 同一算法；Option Strict Off 下用 JObject 传结果。
    ''' </summary>
    Private Function VerifyKey(key As String) As JObject
        Dim raw = If(key, "").Trim()
        If Not raw.StartsWith(KEY_PREFIX, StringComparison.Ordinal) Then Return KeyInvalid("malformed")
        Dim body = raw.Substring(KEY_PREFIX.Length)
        Dim dot = body.LastIndexOf("."c)
        If dot <= 0 OrElse dot = body.Length - 1 Then Return KeyInvalid("malformed")

        Dim payloadText As String
        Try
            payloadText = Encoding.UTF8.GetString(KeyBase64UrlDecode(body.Substring(0, dot)))
        Catch
            Return KeyInvalid("malformed")
        End Try

        Dim sig = body.Substring(dot + 1)
        Dim expected = KeySign(payloadText)
        If sig.Length <> expected.Length OrElse Not String.Equals(sig, expected, StringComparison.Ordinal) Then
            Return KeyInvalid("bad_signature")
        End If

        Dim parts = payloadText.Split("|"c)
        If parts.Length <> KEY_FIELD_COUNT Or Not String.Equals(parts(0), KEY_VERSION, StringComparison.Ordinal) Then
            Return KeyInvalid("bad_payload")
        End If
        Dim expUnix As Long
        If Not Long.TryParse(parts(4), expUnix) OrElse expUnix <= 0 Then Return KeyInvalid("bad_payload")
        Dim periodEnd = KeyEpoch.AddSeconds(expUnix)
        If periodEnd <= DateTime.UtcNow Then Return KeyInvalid("expired")

        Dim seats As Integer
        If Not Integer.TryParse(parts(3), seats) OrElse seats < 1 Then seats = 1
        Return New JObject From {
            {"valid", True},
            {"reason", "ok"},
            {"login", If(String.Equals(parts(1), "-", StringComparison.Ordinal), "", parts(1))},
            {"plan", parts(2)},
            {"seats", seats},
            {"periodEnd", periodEnd}
        }
    End Function

    Private Function KeyInvalid(reason As String) As JObject
        Return New JObject From {
            {"valid", False},
            {"reason", reason},
            {"login", ""},
            {"plan", ""},
            {"seats", 0},
            {"periodEnd", CType(Nothing, DateTime)}
        }
    End Function

    ''' <summary>
    ''' 激活码 → 不可逆标识：SHA256(key) 十六进制串的**前 16 个字符**（是字符数不是字节数，
    ''' 与 JS 端 digest('hex').slice(0,16) 对齐）。必须与 LicenseServer/lib/revocation.js
    ''' 的 hashKey 一致，否则运营台拿注销码对不上台账。
    ''' </summary>
    Private Function KeyHash(key As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(If(key, "")))
            Dim sb As New StringBuilder()
            For i As Integer = 0 To bytes.Length - 1
                sb.Append(bytes(i).ToString("x2"))
            Next
            Return sb.ToString().Substring(0, 16)
        End Using
    End Function

    ''' <summary>
    ''' 生成「消 key 码」（退款凭证）：SPD1-&lt;base64url(payload)&gt;.&lt;base64url(sig)&gt;
    ''' payload：v1|keyHash|plan|destroyedUnix|nonce
    ''' 用户在本机注销后把此码发给客服，运营台验签通过才退款。
    ''' </summary>
    Private Function CreateDeactivationProof(key As String, plan As String) As String
        Dim ts = CLng((DateTime.UtcNow - KeyEpoch).TotalSeconds)
        Dim nonce = Guid.NewGuid().ToString("N").Substring(0, 16)
        Dim payload = String.Join("|", New String() {
            KEY_VERSION,
            KeyHash(key),
            If(String.IsNullOrWhiteSpace(plan), "pro_yearly", plan),
            ts.ToString(),
            nonce
        })
        Return DEACT_PREFIX & KeyBase64UrlEncode(Encoding.UTF8.GetBytes(payload)) & "." & KeySign(payload)
    End Function

    Private Function KeySign(payloadText As String) As String
        Using hmac As New HMACSHA256(Encoding.UTF8.GetBytes(KEY_SECRET))
            Return KeyBase64UrlEncode(hmac.ComputeHash(Encoding.UTF8.GetBytes(payloadText)))
        End Using
    End Function

    ''' <summary>.NET Framework 没有 Replace(Char, Char) 重载，必须用字符串版本。</summary>
    Private Function KeyBase64UrlEncode(bytes As Byte()) As String
        Return Convert.ToBase64String(bytes) _
            .Replace("+", "-") _
            .Replace("/", "_") _
            .Replace("=", "")
    End Function

    Private Function KeyBase64UrlDecode(text As String) As Byte()
        Dim s = If(text, "").Replace("-", "+").Replace("_", "/")
        Select Case s.Length Mod 4
            Case 2 : s &= "=="
            Case 3 : s &= "="
        End Select
        Return Convert.FromBase64String(s)
    End Function

    Private Async Function RegisterAsync(bodyJson As String) As System.Threading.Tasks.Task(Of String)
        Try
            Dim body = JObject.Parse(If(bodyJson, "{}"))
            Dim result = Await SendJsonAsync(HttpMethod.Post, "/auth/register", body, Nothing).ConfigureAwait(False)
            ApplyAuthResult(result)
            Return result.ToString(Formatting.None)
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    Private Async Function LoginAsync(bodyJson As String) As System.Threading.Tasks.Task(Of String)
        Try
            Dim body = JObject.Parse(If(bodyJson, "{}"))
            Dim result = Await SendJsonAsync(HttpMethod.Post, "/auth/login", body, Nothing).ConfigureAwait(False)
            ApplyAuthResult(result)
            Return result.ToString(Formatting.None)
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    Private Async Function LogoutAsync() As System.Threading.Tasks.Task(Of String)
        Try
            Dim token = PeekToken()
            If Not String.IsNullOrEmpty(token) Then
                Await SendJsonAsync(HttpMethod.Post, "/auth/logout", New Dictionary(Of String, Object)(), token).ConfigureAwait(False)
            End If
            RememberToken("")
            Return JsonConvert.SerializeObject(New With {.success = True})
        Catch ex As Exception
            RememberToken("")
            Return JsonConvert.SerializeObject(New With {.success = True, .warning = ex.Message})
        End Try
    End Function

    Private Async Function GetEntitlementAsync() As System.Threading.Tasks.Task(Of String)
        Try
            Dim token = PeekToken()
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
                RememberEntitlement(result.ToString(Formatting.None))
                result("loggedIn") = True
                Return result.ToString(Formatting.None)
            End If
            Dim cached = PeekEntitlement()
            If Not String.IsNullOrWhiteSpace(cached) Then
                Dim c = JObject.Parse(cached)
                c("success") = True
                c("stale") = True
                c("loggedIn") = True
                Return c.ToString(Formatting.None)
            End If
            Return result.ToString(Formatting.None)
        Catch ex As Exception
            Dim cached = PeekEntitlement()
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
            Return JsonConvert.SerializeObject(New With {
                .success = False,
                .error = ex.Message,
                .loggedIn = Not String.IsNullOrEmpty(PeekToken())
            })
        End Try
    End Function

    Private Async Function StartCheckoutAsync(bodyJson As String) As System.Threading.Tasks.Task(Of String)
        Try
            Dim token = PeekToken()
            If String.IsNullOrEmpty(token) Then
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "请先登录", .code = "not_logged_in"})
            End If
            Dim body = JObject.Parse(If(String.IsNullOrWhiteSpace(bodyJson), "{}", bodyJson))
            Dim channel = (If(body.Value(Of String)("channel"), "stripe")).Trim().ToLowerInvariant()
            Dim path = If(channel = "cn", "/checkout/cn", "/checkout/stripe")
            Dim result = Await SendJsonAsync(HttpMethod.Post, path, body, token).ConfigureAwait(False)
            OpenResultUrl(result)
            Return result.ToString(Formatting.None)
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    Private Async Function OpenBillingPortalAsync() As System.Threading.Tasks.Task(Of String)
        Try
            Dim token = PeekToken()
            If String.IsNullOrEmpty(token) Then
                Return JsonConvert.SerializeObject(New With {.success = False, .error = "请先登录", .code = "not_logged_in"})
            End If
            Dim result = Await SendJsonAsync(HttpMethod.Post, "/checkout/stripe/portal", New Dictionary(Of String, Object)(), token).ConfigureAwait(False)
            OpenResultUrl(result)
            Return result.ToString(Formatting.None)
        Catch ex As Exception
            Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
        End Try
    End Function

    Private Async Function PingHealthAsync() As System.Threading.Tasks.Task(Of String)
        Try
            Dim result = Await SendJsonAsync(HttpMethod.Get, "/health", Nothing, Nothing).ConfigureAwait(False)
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

    Private Sub ApplyAuthResult(result As JObject)
        If result Is Nothing Then Return
        If result.Value(Of Boolean?)("success") <> True Then Return
        Dim token = result.Value(Of String)("token")
        If Not String.IsNullOrEmpty(token) Then RememberToken(token)
        RememberEntitlement(result.ToString(Formatting.None))
    End Sub

    Private Sub OpenResultUrl(result As JObject)
        If result Is Nothing Then Return
        Dim url = result.Value(Of String)("url")
        If result.Value(Of Boolean?)("success") <> True OrElse String.IsNullOrWhiteSpace(url) Then Return
        Try
            Process.Start(New ProcessStartInfo(url) With {.UseShellExecute = True})
        Catch ex As Exception
            result("openError") = ex.Message
        End Try
    End Sub

    Private Function BaseUrl() As String
        Dim url = PeekBaseUrl()
        If String.IsNullOrWhiteSpace(url) Then Return DEFAULT_BASE
        Return url
    End Function

    Private Async Function SendJsonAsync(method As HttpMethod, relativePath As String, bodyObj As Object, bearer As String) As System.Threading.Tasks.Task(Of JObject)
        Dim url = BaseUrl() & relativePath
        Using req As New HttpRequestMessage(method, New Uri(url))
            If Not String.IsNullOrEmpty(bearer) Then
                req.Headers.TryAddWithoutValidation("Authorization", "Bearer " & bearer)
            End If
            If bodyObj IsNot Nothing Then
                req.Content = New StringContent(JsonConvert.SerializeObject(bodyObj), Encoding.UTF8, "application/json")
            End If
            Dim resp = Await Http.Value.SendAsync(req).ConfigureAwait(False)
            Dim text = Await resp.Content.ReadAsStringAsync().ConfigureAwait(False)
            Dim parsed As JObject
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

    ''' <summary>
    ''' 读共享会话，带重试。Excel / Word / PPT 可能同时读写这一份文件：
    ''' File.ReadAllText 会撞 ERROR_SHARING_VIOLATION（IOException）；即便拿到文本，
    ''' 也可能正好读到对方写一半的内容（JSON 解析失败）。这两种情况都必须重试，
    ''' 不能降级成「空会话」——那会把已激活用户误判成未激活、已登录用户判成未登录。
    ''' </summary>
    Private Function ReadSession() As JObject
        SyncLock FileLock
            For attempt As Integer = 1 To SHARED_FILE_RETRY_TIMES
                Try
                    If Not File.Exists(SessionPath) Then Return EmptySession()
                    Dim text = File.ReadAllText(SessionPath)
                    If String.IsNullOrWhiteSpace(text) Then Return EmptySession()
                    Dim jo = JObject.Parse(text)
                    Dim protectedToken = jo.Value(Of String)("tokenProtected")
                    If Not String.IsNullOrEmpty(protectedToken) Then
                        jo("token") = Unprotect(protectedToken)
                    Else
                        jo("token") = If(jo.Value(Of String)("token"), "")
                    End If
                    Dim protectedActivation = jo.Value(Of String)("activationProtected")
                    If Not String.IsNullOrEmpty(protectedActivation) Then
                        Try
                            jo("activation") = Unprotect(protectedActivation)
                        Catch
                            jo("activation") = ""
                        End Try
                    Else
                        jo("activation") = If(jo.Value(Of String)("activation"), "")
                    End If
                    jo("activationSet") = jo.Value(Of Boolean?)("activationSet").GetValueOrDefault(False)
                    jo("deactivated") = If(jo.Value(Of String)("deactivated"), "")
                    Return jo
                Catch ex As FileNotFoundException
                    Return EmptySession()
                Catch ex As IOException
                    System.Threading.Thread.Sleep(SHARED_FILE_RETRY_DELAY_MS)
                Catch ex As JsonException
                    ' 读到对方写了一半的内容（torn read），重试而不是降级
                    System.Threading.Thread.Sleep(SHARED_FILE_RETRY_DELAY_MS)
                End Try
            Next
            ' 重试耗尽：降级空会话并留痕，不把异常炸进门禁调用链
            Debug.WriteLine("[HostLicense] 读取共享会话重试 " & SHARED_FILE_RETRY_TIMES & " 次仍失败: " & SessionPath)
            Return EmptySession()
        End SyncLock
    End Function

    Private Sub WriteSession(session As JObject)
        SyncLock FileLock
            Dim stored As New JObject()
            stored("baseUrl") = If(session.Value(Of String)("baseUrl"), "")
            stored("entitlement") = If(session.Value(Of String)("entitlement"), "")
            stored("tokenSet") = session.Value(Of Boolean?)("tokenSet").GetValueOrDefault(False)
            Dim token = If(session.Value(Of String)("token"), "")
            stored("token") = ""
            stored("tokenProtected") = If(String.IsNullOrEmpty(token), "", Protect(token))
            Dim activation = If(session.Value(Of String)("activation"), "")
            stored("activationSet") = session.Value(Of Boolean?)("activationSet").GetValueOrDefault(False)
            stored("activationProtected") = If(String.IsNullOrEmpty(activation), "", Protect(activation))
            stored("deactivated") = If(session.Value(Of String)("deactivated"), "")
            Dim dir = Path.GetDirectoryName(SessionPath)
            If Not String.IsNullOrEmpty(dir) Then Directory.CreateDirectory(dir)
            Dim text = stored.ToString(Formatting.None)
            Dim last As IOException = Nothing
            For attempt As Integer = 1 To SHARED_FILE_RETRY_TIMES
                Try
                    File.WriteAllText(SessionPath, text)
                    Return
                Catch ex As IOException
                    last = ex
                    System.Threading.Thread.Sleep(SHARED_FILE_RETRY_DELAY_MS)
                End Try
            Next
            Throw last
        End SyncLock
    End Sub

    Private Function EmptySession() As JObject
        Return New JObject From {
            {"token", ""},
            {"baseUrl", ""},
            {"entitlement", ""},
            {"tokenSet", False},
            {"activation", ""},
            {"activationSet", False},
            {"deactivated", ""}
        }
    End Function

    Private Function Protect(text As String) As String
        Dim raw = Encoding.UTF8.GetBytes(text)
        Dim protectedBytes = ProtectedData.Protect(raw, Nothing, DataProtectionScope.CurrentUser)
        Return Convert.ToBase64String(protectedBytes)
    End Function

    Private Function Unprotect(text As String) As String
        Dim protectedBytes = Convert.FromBase64String(text)
        Dim raw = ProtectedData.Unprotect(protectedBytes, Nothing, DataProtectionScope.CurrentUser)
        Return Encoding.UTF8.GetString(raw)
    End Function

End Module
