' ============================================================
' PptThemeDesignValidator - AI 生成 PPT 主题的配色校验与自动修正
'
' 目标：让 LLM 通过 ppt_create_theme 创建主题时，"自动设计"不等于"乱设计"。
'  - WCAG 2.1 对比度底线：
'      text     vs bg  >= 4.5  （正文/标题文本可读性，硬底线）
'      primary  vs bg  >= 3.0  （主色：标题图形/按钮/强调文本）
'      accent   vs bg  >= 1.5  （强调点缀色，装饰性图形可区分）
'      secondary vs bg >= 1.3  （图表系列/辅助色，仅防与背景同化）
'  - 颜色格式容错：[r,g,b] 数组 或 "#RRGGBB" 字符串
'  - 不达标自动微调（确定性单调收敛，最多 25 轮），仍不达标则报错返回
'  - 纯静态、无 COM / IO 依赖（仅 Newtonsoft.Json），可离线单测
'
' 使用方：PptMcpTools.ExecuteCreateTheme → ValidateAndFix → SaveCustomPptTheme
' ============================================================
Option Strict On
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports Newtonsoft.Json.Linq

Namespace Services

    ''' <summary>
    ''' 主题配色校验器：WCAG 对比度校验 + 确定性自动微调。
    ''' 输出 fixedTheme（标准化 [r,g,b] 的完整主题 JSON）、adjusted（调整说明）、errors（硬错误）。
    ''' </summary>
    Public NotInheritable Class PptThemeDesignValidator

        Private Sub New()
        End Sub

        ' ------------------------------------------------------------
        ' WCAG 2.1 相对亮度与对比度（sRGB）
        ' ------------------------------------------------------------

        ''' <summary>sRGB 通道线性化（WCAG 2.x 公式）</summary>
        Public Shared Function ChannelLinear(c As Integer) As Double
            Dim v As Double = Math.Max(0, Math.Min(255, c)) / 255.0
            If v <= 0.03928 Then Return v / 12.92
            Return Math.Pow((v + 0.055) / 1.055, 2.4)
        End Function

        ''' <summary>相对亮度 L = 0.2126R + 0.7152G + 0.0722B</summary>
        Public Shared Function RelativeLuminance(r As Integer, g As Integer, b As Integer) As Double
            Return 0.2126 * ChannelLinear(r) + 0.7152 * ChannelLinear(g) + 0.0722 * ChannelLinear(b)
        End Function

        ''' <summary>对比度 (L1+0.05)/(L2+0.05)，结果 >= 1.0</summary>
        Public Shared Function ContrastRatio(c1 As Integer(), c2 As Integer()) As Double
            If c1 Is Nothing OrElse c1.Length <> 3 OrElse c2 Is Nothing OrElse c2.Length <> 3 Then
                Return 1.0
            End If
            Dim l1 As Double = RelativeLuminance(c1(0), c1(1), c1(2))
            Dim l2 As Double = RelativeLuminance(c2(0), c2(1), c2(2))
            Dim hi As Double = Math.Max(l1, l2)
            Dim lo As Double = Math.Min(l1, l2)
            Return (hi + 0.05) / (lo + 0.05)
        End Function

        ' ------------------------------------------------------------
        ' 颜色解析：支持 [r,g,b] 数组 或 "#RRGGBB" 字符串；非法返回 False
        ' ------------------------------------------------------------

        Public Shared Function TryParseColor(value As JToken, ByRef rgb As Integer()) As Boolean
            rgb = Nothing
            If value Is Nothing Then Return False
            If TypeOf value Is JArray Then
                Dim arr = DirectCast(value, JArray)
                If arr.Count <> 3 Then Return False
                Dim out(2) As Integer
                For i As Integer = 0 To 2
                    Dim n As Integer
                    If Not Integer.TryParse(Convert.ToString(arr(i)), n) Then Return False
                    If n < 0 OrElse n > 255 Then Return False
                    out(i) = n
                Next
                rgb = out
                Return True
            End If
            Dim s As String = Convert.ToString(value).Trim()
            If s.StartsWith("#") Then s = s.Substring(1)
            If s.Length = 6 Then
                Dim r2 As Integer
                Dim g2 As Integer
                Dim b2 As Integer
                If Not TryParseHex(s.Substring(0, 2), r2) Then Return False
                If Not TryParseHex(s.Substring(2, 2), g2) Then Return False
                If Not TryParseHex(s.Substring(4, 2), b2) Then Return False
                rgb = New Integer() {r2, g2, b2}
                Return True
            End If
            Return False
        End Function

        Private Shared Function TryParseHex(hex As String, ByRef value As Integer) As Boolean
            Return Integer.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, value)
        End Function

        ' ------------------------------------------------------------
        ' 校验 + 自动微调主入口
        ' ------------------------------------------------------------

        ''' <summary>
        ''' 校验并自动修正主题 JSON。
        ''' 返回 JObject：success / fixedTheme / adjusted(JArray) / errors(JArray) / message
        ''' errors 非空时 success=False（调用方应拒绝落盘并回报）。
        ''' </summary>
        Public Shared Function ValidateAndFix(themeJson As String) As JObject
            Dim result As New JObject()
            result("success") = False
            result("adjusted") = New JArray()
            result("errors") = New JArray()

            If String.IsNullOrWhiteSpace(themeJson) Then
                DirectCast(result("errors"), JArray).Add("主题内容为空")
                result("message") = "主题内容为空"
                Return result
            End If

            Dim t As JObject
            Try
                t = JObject.Parse(themeJson)
            Catch ex As Exception
                DirectCast(result("errors"), JArray).Add("主题 JSON 解析失败: " & ex.Message)
                result("message") = "主题 JSON 解析失败"
                Return result
            End Try

            ' ---- 0) 风格方向展开：theme 含 styleDirection 时合并 preset（显式字段覆盖） ----
            If t("styleDirection") IsNot Nothing Then
                Try
                    t = PptStyleDirections.ExpandPreset(t)
                Catch ex As ArgumentException
                    DirectCast(result("errors"), JArray).Add(ex.Message)
                    result("message") = "主题未通过校验: " & ex.Message
                    Return result
                End Try
            End If

            ' ---- 1) 补全基础字段（与 SaveCustomPptTheme 同契约） ----
            Dim id As String = If(t.Value(Of String)("id"), "").Trim()
            If String.IsNullOrWhiteSpace(id) Then
                id = "custom_" & Guid.NewGuid().ToString("N").Substring(0, 8)
                t("id") = id
                DirectCast(result("adjusted"), JArray).Add("id 自动生成: " & id)
            End If
            If String.IsNullOrWhiteSpace(t.Value(Of String)("name")) Then t("name") = id
            If String.IsNullOrWhiteSpace(t.Value(Of String)("font")) Then t("font") = "微软雅黑"

            ' ---- 2) 解析/补缺 7 组颜色（默认值 = 商务蓝，与 PptGenerationService.FALLBACK_THEME 一致） ----
            Dim defaults As New Dictionary(Of String, Integer())() From {
                {"primary", New Integer() {31, 78, 121}},
                {"secondary", New Integer() {190, 167, 91}},
                {"accent", New Integer() {255, 165, 0}},
                {"bg", New Integer() {245, 247, 250}},
                {"light", New Integer() {239, 243, 248}},
                {"text", New Integer() {33, 37, 41}},
                {"gray", New Integer() {150, 150, 150}}
            }

            Dim colors As JObject = TryCast(t("colors"), JObject)
            If colors Is Nothing Then
                colors = New JObject()
                t("colors") = colors
            End If

            Dim parsed As New Dictionary(Of String, Integer())(StringComparer.OrdinalIgnoreCase)
            Dim errors As New List(Of String)()
            For Each key As String In defaults.Keys
                Dim token As JToken = Nothing
                Dim rgb As Integer() = Nothing
                If colors.TryGetValue(key, token) AndAlso TryParseColor(token, rgb) Then
                    parsed(key) = rgb
                Else
                    parsed(key) = defaults(key)
                    If colors.ContainsKey(key) Then
                        errors.Add(key & " 颜色格式非法，已回退默认")
                    End If
                End If
            Next

            ' ---- 3) 对比度校验 + 自动微调 ----
            ' 阈值与内置主题现状对齐（商务蓝/科技绿等全部通过），同时守住文本可读性硬底线
            Dim bg As Integer() = parsed("bg")
            Dim targets As New Dictionary(Of String, Double)() From {
                {"text", 4.5},
                {"primary", 3.0},
                {"accent", 1.5},
                {"secondary", 1.3}
            }
            For Each kv As KeyValuePair(Of String, Double) In targets
                Dim key As String = kv.Key
                Dim threshold As Double = kv.Value
                Dim color As Integer() = parsed(key)
                Dim ratio As Double = ContrastRatio(color, bg)
                Dim before As Double = ratio
                Dim iterations As Integer = 0
                While ratio < threshold AndAlso iterations < 25
                    color = StepAway(color, bg)
                    parsed(key) = color
                    ratio = ContrastRatio(color, bg)
                    iterations += 1
                End While
                If ratio >= threshold Then
                    If iterations > 0 Then
                        DirectCast(result("adjusted"), JArray).Add(
                            key & " 已自动调整（对比度 " & FormatRatio(before) & " → " & FormatRatio(ratio) &
                            "，目标 ≥ " & threshold.ToString("0.0") & "）")
                    End If
                Else
                    errors.Add(key & " 对比度无法达标（" & FormatRatio(ratio) & " < " & threshold.ToString("0.0") &
                               "）：请提供更亮的背景或与背景更对比的 " & key & " 色")
                End If
            Next

            ' ---- 4) 写回 colors（统一标准化为 [r,g,b]） ----
            Dim fixedColors As New JObject()
            For Each key As String In defaults.Keys
                Dim c As Integer() = parsed(key)
                fixedColors(key) = New JArray(c(0), c(1), c(2))
            Next
            t("colors") = fixedColors

            ' ---- 5) designTokens 校验（版式令牌：密度/字号/边距/装饰，非法即拒绝创建） ----
            ValidateDesignTokens(t, errors)
            result("fixedTheme") = t

            ' ---- 5) 汇总 ----
            Dim adjusted = DirectCast(result("adjusted"), JArray)
            If errors.Count = 0 Then
                result("success") = True
                result("message") = "配色校验通过" & (If(adjusted.Count > 0, "（含自动微调）", ""))
            Else
                For Each e As String In errors
                    DirectCast(result("errors"), JArray).Add(e)
                Next
                result("message") = "配色未通过校验: " & String.Join("; ", errors)
            End If
            Return result
        End Function

        ' ------------------------------------------------------------
        ' 内部工具
        ' ------------------------------------------------------------

        ''' <summary>
        ''' 校验主题 JSON 的 designTokens（版式令牌）。
        ''' - density: compact / comfortable / spacious（枚举）
        ''' - 数字令牌：titleSize/headingSize/bodySize/smallSize/footerSize（8-80）、
        '''   margin（8-120）、grid（4-60）、titleTop/contentTop（0-300）、ribbonWidth（20-600）
        ''' - ribbonStyle: line / bar / none（枚举）
        ''' 非法字段记入 errors（调用方据此拒绝创建）；不存在的字段不校验（引擎回退默认）。
        ''' 引擎侧 BuildDesignTokens 对 density 做预设展开 + 显式字段覆盖。
        ''' </summary>
        Private Shared Sub ValidateDesignTokens(t As JObject, errors As List(Of String))
            Dim token = t("designTokens")
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Return
            Dim dt As JObject = TryCast(token, JObject)
            If dt Is Nothing Then
                errors.Add("designTokens 必须是 JSON 对象")
                Return
            End If

            Dim density As String = If(dt.Value(Of String)("density"), "").Trim().ToLowerInvariant()
            If density <> "" AndAlso density <> "compact" AndAlso density <> "comfortable" AndAlso density <> "spacious" Then
                errors.Add("designTokens.density 仅支持 compact / comfortable / spacious")
            End If

            Dim intTokens As String() = {"titleSize", "headingSize", "bodySize", "smallSize", "footerSize",
                                          "margin", "grid", "titleTop", "contentTop", "ribbonWidth"}
            For Each key As String In intTokens
                Dim raw = dt(key)
                If raw Is Nothing OrElse raw.Type = JTokenType.Null Then Continue For
                Dim n As Integer
                If Not Integer.TryParse(Convert.ToString(raw), n) Then
                    errors.Add("designTokens." & key & " 必须为整数")
                    Continue For
                End If
                Dim minV As Integer = 8
                Dim maxV As Integer = 80
                Select Case key
                    Case "margin" : minV = 8 : maxV = 120
                    Case "grid" : minV = 4 : maxV = 60
                    Case "titleTop", "contentTop" : minV = 0 : maxV = 300
                    Case "ribbonWidth" : minV = 20 : maxV = 600
                End Select
                If n < minV OrElse n > maxV Then
                    errors.Add("designTokens." & key & " 超出范围 [" & minV.ToString() & ", " & maxV.ToString() & "]")
                End If
            Next

            Dim ribbon As String = If(dt.Value(Of String)("ribbonStyle"), "").Trim().ToLowerInvariant()
            If ribbon <> "" AndAlso ribbon <> "line" AndAlso ribbon <> "bar" AndAlso ribbon <> "none" Then
                errors.Add("designTokens.ribbonStyle 仅支持 line / bar / none")
            End If

            ' pageVariants：主题级页型变体默认（key=页型，value=合法变体）
            Dim pv = dt("pageVariants")
            If pv IsNot Nothing AndAlso pv.Type <> JTokenType.Null Then
                Dim pvObj As JObject = TryCast(pv, JObject)
                If pvObj Is Nothing Then
                    errors.Add("designTokens.pageVariants 必须是 JSON 对象（如 {""title"":""modern"",""kpi"":""hero""}）")
                Else
                    For Each prop As JProperty In pvObj.Properties()
                        Dim allowed As String() = PageVariantAllowed(prop.Name)
                        If allowed Is Nothing Then
                            errors.Add("designTokens.pageVariants." & prop.Name &
                                       " 不是支持的页型（title/text/kpi/chart/section）")
                            Continue For
                        End If
                        Dim val As String = Convert.ToString(prop.Value).Trim().ToLowerInvariant()
                        If Array.IndexOf(allowed, val) < 0 Then
                            errors.Add("designTokens.pageVariants." & prop.Name & " 仅支持 " & String.Join(" / ", allowed))
                        End If
                    Next
                End If
            End If
        End Sub

        ''' <summary>
        ''' 页型变体合法值（单一事实源：校验器校验 + 引擎 BuildDesignTokens 过滤共用）。
        ''' 未知页型返回 Nothing。
        ''' </summary>
        Public Shared Function PageVariantAllowed(pageKey As String) As String()
            Select Case pageKey.Trim().ToLowerInvariant()
                Case "title" : Return New String() {"split", "centered", "modern"}
                Case "text" : Return New String() {"bullets", "cards", "numbered"}
                Case "kpi" : Return New String() {"grid", "hero", "row"}
                Case "chart" : Return New String() {"standard", "full", "insight"}
                Case "section" : Return New String() {"number", "centered"}
                Case Else : Return Nothing
            End Select
        End Function

        ''' <summary>
        ''' 朝增大与 bg 对比度的方向移动一步（确定性单调收敛）。
        ''' 明显更亮则提亮、更暗则压暗；两者亮度几乎相等（近同化，如近白对近白）时
        ''' 优先压暗——深色方向动态范围大（245→0 对比度可从 1.0 到 21+），提亮受 255 上限约束可能无效。
        ''' 步长 15%。
        ''' </summary>
        Private Shared Function StepAway(color As Integer(), bg As Integer()) As Integer()
            Dim lc As Double = RelativeLuminance(color(0), color(1), color(2))
            Dim lb As Double = RelativeLuminance(bg(0), bg(1), bg(2))
            Dim darken As Boolean = (Math.Abs(lc - lb) < 0.02) OrElse (lc < lb)
            Dim out(2) As Integer
            For i As Integer = 0 To 2
                If darken Then
                    out(i) = CInt(Math.Round(color(i) * 0.85))
                Else
                    out(i) = Math.Min(255, CInt(Math.Round(color(i) + (255 - color(i)) * 0.15)))
                End If
            Next
            Return out
        End Function

        Private Shared Function FormatRatio(v As Double) As String
            Return v.ToString("0.00", CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
