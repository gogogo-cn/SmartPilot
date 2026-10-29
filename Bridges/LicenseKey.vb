Option Explicit On
Option Infer On

Imports System
Imports System.Security.Cryptography
Imports System.Text

Namespace Bridges

    ''' <summary>
    ''' 离线激活码验签结果。
    ''' </summary>
    Friend NotInheritable Class LicenseKeyResult
        Public Property Valid As Boolean
        Public Property Reason As String
        Public Property Login As String
        Public Property Plan As String
        Public Property Seats As Integer
        Public Property PeriodEnd As DateTime
    End Class

    ''' <summary>
    ''' 离线激活码（路线①：无授权服务器也能卖）。
    ''' key 格式：SPK1-&lt;base64url(payload)&gt;.&lt;base64url(sig)&gt;
    ''' payload：v1|login|plan|seats|expUnixSeconds|nonce
    ''' （第 6 段 nonce 是随机数，只为保证同参数批量生成的 key 互不相同，
    '''  验签时只计数不取值。）
    ''' sig：HMAC-SHA256(payload, SECRET) 的 32 字节摘要（base64url）。
    '''
    ''' 算法与 LicenseServer/lib/licenseKey.js 逐字节一致（改算法必须三处同步：
    ''' 本文件、Hosts/Shared/HostLicense.vb、LicenseServer/lib/licenseKey.js）。
    '''
    ''' 威胁模型：SECRET 内置于客户端，纯本地验签是「防君子不防小人」的轻度防护；
    ''' 远程吊销、换机绑定等强校验走 LicenseServer 在线链路。生产环境更换 SECRET
    ''' 需同时改三处并重新编译发布，旧 key 会全部失效。
    ''' </summary>
    Friend NotInheritable Class LicenseKey

        Private Const PREFIX As String = "SPK1-"
        ''' <summary>消 key 码（退款凭证）前缀，与激活码前缀严格隔离。</summary>
        Private Const DEACT_PREFIX As String = "SPD1-"
        Private Const VERSION As String = "v1"
        Private Const FIELD_COUNT As Integer = 6
        Private Const DEACT_FIELD_COUNT As Integer = 5
        Private Const SECRET As String = "SmartPilot-Offline-License-Key-v1-CHANGE-ME"
        Private Shared ReadOnly EPOCH As New DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)

        ''' <summary>校验激活码。任何格式问题返回 Valid=False + Reason，不抛异常。</summary>
        Public Shared Function Verify(key As String) As LicenseKeyResult
            Dim raw = If(key, "").Trim()
            If Not raw.StartsWith(PREFIX, StringComparison.Ordinal) Then Return Invalid("malformed")
            Dim body = raw.Substring(PREFIX.Length)
            Dim dot = body.LastIndexOf("."c)
            If dot <= 0 OrElse dot = body.Length - 1 Then Return Invalid("malformed")

            Dim payloadText As String
            Try
                payloadText = Encoding.UTF8.GetString(Base64UrlDecode(body.Substring(0, dot)))
            Catch
                Return Invalid("malformed")
            End Try

            Dim sig = body.Substring(dot + 1)
            Dim expected = Sign(payloadText)
            If sig.Length <> expected.Length OrElse
               Not String.Equals(sig, expected, StringComparison.Ordinal) Then
                Return Invalid("bad_signature")
            End If

            Dim parts = payloadText.Split("|"c)
            If parts.Length <> FIELD_COUNT OrElse Not String.Equals(parts(0), VERSION, StringComparison.Ordinal) Then
                Return Invalid("bad_payload")
            End If
            Dim expUnix As Long
            If Not Long.TryParse(parts(4), expUnix) OrElse expUnix <= 0 Then Return Invalid("bad_payload")
            Dim periodEnd = EPOCH.AddSeconds(expUnix)
            If periodEnd <= DateTime.UtcNow Then Return Invalid("expired")

            Dim seats As Integer
            If Not Integer.TryParse(parts(3), seats) OrElse seats < 1 Then seats = 1
            Return New LicenseKeyResult With {
                .Valid = True,
                .Reason = "ok",
                .Login = If(String.Equals(parts(1), "-", StringComparison.Ordinal), "", parts(1)),
                .Plan = parts(2),
                .Seats = seats,
                .PeriodEnd = periodEnd
            }
        End Function

        ''' <summary>
        ''' 激活码 → 不可逆标识：SHA256(key) 十六进制串的**前 16 个字符**（注意是字符数，
        ''' 不是字节数——JS 端是 digest('hex').slice(0,16)，对应 8 个字节）。
        ''' 必须与 LicenseServer/lib/revocation.js 的 hashKey 完全一致，否则运营台拿注销码
        ''' 对不上台账（跨语言测试 Tests 会守这一条）。标识不含激活码原文，可安全外发。
        ''' </summary>
        Public Shared Function HashKey(key As String) As String
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
        '''
        ''' 业务闭环：用户在本机点「注销并申请退款」→ 本机激活被清除且该 key 永久失效 →
        ''' 生成此码 → 发给客服 → 运营台验签通过才退款。这样「退款」与「不能用」是同
        ''' 一个动作的两面，不依赖强制联网。
        ''' </summary>
        Public Shared Function CreateDeactivationProof(key As String, plan As String) As String
            Dim ts = CLng((DateTime.UtcNow - EPOCH).TotalSeconds)
            Dim nonce = Guid.NewGuid().ToString("N").Substring(0, 16)
            Dim payload = String.Join("|", New String() {
                VERSION,
                HashKey(key),
                If(String.IsNullOrWhiteSpace(plan), "pro_yearly", plan),
                ts.ToString(),
                nonce
            })
            Return DEACT_PREFIX & Base64UrlEncode(Encoding.UTF8.GetBytes(payload)) & "." & Sign(payload)
        End Function

        ''' <summary>消 key 码前缀（供桥层拼提示文案用）。</summary>
        Public Shared ReadOnly Property DeactivationPrefix As String
            Get
                Return DEACT_PREFIX
            End Get
        End Property

        Private Shared Function Sign(payloadText As String) As String
            Using hmac As New HMACSHA256(Encoding.UTF8.GetBytes(SECRET))
                Return Base64UrlEncode(hmac.ComputeHash(Encoding.UTF8.GetBytes(payloadText)))
            End Using
        End Function

        ''' <summary>.NET Framework 没有 Replace(Char, Char) 重载，必须用字符串版本。</summary>
        Private Shared Function Base64UrlEncode(bytes As Byte()) As String
            Return Convert.ToBase64String(bytes) _
                .Replace("+", "-") _
                .Replace("/", "_") _
                .Replace("=", "")
        End Function

        Private Shared Function Base64UrlDecode(text As String) As Byte()
            Dim s = If(text, "").Replace("-", "+").Replace("_", "/")
            Select Case s.Length Mod 4
                Case 2 : s &= "=="
                Case 3 : s &= "="
            End Select
            Return Convert.FromBase64String(s)
        End Function

        Private Shared Function Invalid(reason As String) As LicenseKeyResult
            Return New LicenseKeyResult With {.Valid = False, .Reason = reason}
        End Function

    End Class

End Namespace
