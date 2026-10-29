' ============================================================
' PptThemeDesignValidatorTests.vb — AI 主题配色校验器单元测试
'
' 被测对象：直接编译的元数据生产代码 Services/PptThemeDesignValidator.vb
' （WCAG 对比度校验 + 自动微调），纯静态无 COM 依赖，可离线单测。
' 覆盖：内置主题基线不误调 / 低对比自动加深 / 深色主题自适应 /
'       同化色分离 / hex 格式 / 非法格式回退 / 无法达标报错 / 空与坏 JSON
' ============================================================
Imports System
Imports ExcelAIVSTOAddIn.Services
Imports Newtonsoft.Json.Linq

Module PptThemeDesignValidatorTests

    Private _failures As Integer = 0
    Private _total As Integer = 0

    Private Sub Check(condition As Boolean, label As String)
        _total += 1
        If condition Then
            Console.WriteLine("  ok " & label)
        Else
            _failures += 1
            Console.WriteLine("  FAIL " & label)
        End If
    End Sub

    ''' <summary>从校验结果提取固定后的主题 JSON 对象</summary>
    Private Function GetFixed(result As JObject) As JObject
        Return TryCast(result("fixedTheme"), JObject)
    End Function

    ''' <summary>从 fixedTheme.colors 取一组 [r,g,b]</summary>
    Private Function GetColor(t As JObject, key As String) As Integer()
        Dim colors As JObject = TryCast(t("colors"), JObject)
        Dim arr As JArray = TryCast(colors(key), JArray)
        Return New Integer() {CInt(arr(0)), CInt(arr(1)), CInt(arr(2))}
    End Function

    Private Function AdjustedCount(result As JObject) As Integer
        Dim arr = TryCast(result("adjusted"), JArray)
        Return If(arr Is Nothing, 0, arr.Count)
    End Function

    Public Function Run() As (Total As Integer, Failures As Integer)
        Console.WriteLine("== PptThemeDesignValidator 配色校验单元测试 ==")

        TestBusinessBlueBaseline()
        TestLowContrastTextAutoDarken()
        TestDarkThemeAdaptive()
        TestPrimarySimilarToBg()
        TestHexColorParsing()
        TestPartialColorsFallback()
        TestInvalidColorFormat()
        TestUnfixableContrast()
        TestEmptyAndBadJson()
        TestValidDesignTokens()
        TestInvalidDensityToken()
        TestOutOfRangeToken()
        TestInvalidRibbonStyle()
        TestNonObjectDesignTokens()
        TestTechDirectionPreset()
        TestMinimalDirectionPreset()
        TestDirectionExplicitOverride()
        TestInvalidStyleDirection()
        TestAllPresetsZeroAdjustment()
        TestValidPageVariants()
        TestInvalidPageVariantKey()
        TestInvalidPageVariantValue()
        TestNonObjectPageVariants()

        Console.WriteLine()
        If _failures = 0 Then
            Console.WriteLine("[PASS] PptThemeDesignValidator " & _total.ToString() & " 项断言通过")
        Else
            Console.WriteLine("[FAIL] PptThemeDesignValidator " & _failures.ToString() & " / " & _total.ToString() & " 项断言失败")
        End If
        Return (_total, _failures)
    End Function

    ' ------------------------------------------------------------
    ' 1) 内置商务蓝基线：全部通过、无自动调整
    ' ------------------------------------------------------------
    Private Sub TestBusinessBlueBaseline()
        Console.WriteLine("[Baseline] 商务蓝主题不应触发任何调整")
        Dim theme = "{""id"":""business_blue_test"",""name"":""商务蓝"",""font"":""微软雅黑"",""colors"":{" &
            """primary"":[0,90,156],""secondary"":[0,120,200],""accent"":[255,165,0]," &
            """bg"":[245,248,252],""light"":[230,240,250],""text"":[51,51,51],""gray"":[150,150,150]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "商务蓝校验通过")
        Check(AdjustedCount(result) = 0, "商务蓝无自动调整（adjusted 为空）")
    End Sub

    ' ------------------------------------------------------------
    ' 2) 低对比 text（浅灰字浅底）→ 自动加深到 >= 4.5
    ' ------------------------------------------------------------
    Private Sub TestLowContrastTextAutoDarken()
        Console.WriteLine("[Text] 低对比文本自动加深")
        Dim theme = "{""id"":""low_contrast"",""colors"":{" &
            """bg"":[255,255,255],""text"":[200,200,200]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Dim t = GetFixed(result)
        Dim text = GetColor(t, "text")
        Dim bg = GetColor(t, "bg")
        Dim ratio = PptThemeDesignValidator.ContrastRatio(text, bg)
        Check(ratio >= 4.5, "调整后 text/bg 对比度 " & ratio.ToString("0.00") & " >= 4.5")
        Check(Convert.ToString(result("adjusted")).Contains("text 已自动调整"), "adjusted 含 text 调整说明")
    End Sub

    ' ------------------------------------------------------------
    ' 3) 深色主题（深底浅字）→ 通过且方向自适应（提亮而非压暗）
    ' ------------------------------------------------------------
    Private Sub TestDarkThemeAdaptive()
        Console.WriteLine("[Dark] 深色主题自适应（text 应被提亮）")
        Dim theme = "{""id"":""dark_test"",""colors"":{" &
            """bg"":[20,24,35],""text"":[180,185,200]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Dim t = GetFixed(result)
        Dim text = GetColor(t, "text")
        Dim bg = GetColor(t, "bg")
        Dim ratio = PptThemeDesignValidator.ContrastRatio(text, bg)
        Check(ratio >= 4.5, "深色主题调整后对比度 " & ratio.ToString("0.00") & " >= 4.5")
        Check(text(0) >= 180, "深色主题 text 未被错误压暗（R=" & text(0).ToString() & "）")
    End Sub

    ' ------------------------------------------------------------
    ' 4) 主色与背景同化（primary ≈ bg）→ primary 被分离到 >= 3.0
    ' ------------------------------------------------------------
    Private Sub TestPrimarySimilarToBg()
        Console.WriteLine("[Primary] 主色与背景同化时自动分离")
        Dim theme = "{""id"":""same_primary"",""colors"":{" &
            """bg"":[245,247,250],""primary"":[246,247,249]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Dim t = GetFixed(result)
        Dim primary = GetColor(t, "primary")
        Dim bg = GetColor(t, "bg")
        Dim ratio = PptThemeDesignValidator.ContrastRatio(primary, bg)
        Check(ratio >= 3.0, "调整后 primary/bg 对比度 " & ratio.ToString("0.00") & " >= 3.0")
        Check(Convert.ToString(result("adjusted")).Contains("primary 已自动调整"), "adjusted 含 primary 调整说明")
    End Sub

    ' ------------------------------------------------------------
    ' 5) hex "#RRGGBB" 格式解析
    ' ------------------------------------------------------------
    Private Sub TestHexColorParsing()
        Console.WriteLine("[Hex] #RRGGBB 格式解析")
        Dim theme = "{""id"":""hex_test"",""colors"":{" &
            """primary"":""#005A9C"",""accent"":""#FFA500"",""bg"":""#F5F8FC""}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "hex 主题校验通过")
        Dim t = GetFixed(result)
        Dim primary = GetColor(t, "primary")
        Check(primary(0) = 0 AndAlso primary(1) = 90 AndAlso primary(2) = 156, "primary #005A9C 解析为 [0,90,156]")
        Dim colors As JObject = TryCast(t("colors"), JObject)
        Dim textArr As JArray = TryCast(colors("text"), JArray)
        Check(textArr IsNot Nothing AndAlso CInt(textArr(0)) = 33, "缺省 text 回退商务蓝默认 [33,37,41]")
    End Sub

    ' ------------------------------------------------------------
    ' 6) 部分颜色缺省 → 回退默认并成功
    ' ------------------------------------------------------------
    Private Sub TestPartialColorsFallback()
        Console.WriteLine("[Fallback] 只给 primary+bg，其余回退默认")
        Dim theme = "{""id"":""partial_test"",""colors"":{" &
            """primary"":[10,120,200],""bg"":[255,255,255]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "部分颜色主题校验通过")
        Dim t = GetFixed(result)
        Dim colors As JObject = TryCast(t("colors"), JObject)
        Check(colors("secondary") IsNot Nothing AndAlso colors("accent") IsNot Nothing, "secondary/accent 已补默认")
    End Sub

    ' ------------------------------------------------------------
    ' 7) 显式非法颜色格式 → 记 errors 并拒绝（success=False）
    ' ------------------------------------------------------------
    Private Sub TestInvalidColorFormat()
        Console.WriteLine("[Invalid] 非法颜色格式拒绝创建")
        Dim theme = "{""id"":""bad_color"",""colors"":{" &
            """bg"":[255,255,255],""text"":""not_a_color"",""primary"":[999,0,0]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "非法颜色格式导致 success=False")
        Dim errors = TryCast(result("errors"), JArray)
        Check(errors.Count >= 2, "errors 记录 text/primary 两处非法（count=" & errors.Count.ToString() & "）")
    End Sub

    ' ------------------------------------------------------------
    ' 8) 无法达标：bg 亮度≈0.175（25 轮压暗后 text 已到极值仍 < 4.5）→ 报错
    ' ------------------------------------------------------------
    Private Sub TestUnfixableContrast()
        Console.WriteLine("[Unfixable] 中间亮度背景 25 轮内无法达标时报错")
        Dim theme = "{""id"":""unfixable"",""colors"":{" &
            """bg"":[116,116,116],""text"":[116,116,116]}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "无法达标 success=False")
        Check(Convert.ToString(result("message")).Contains("对比度无法达标"), "message 含无法达标说明")
    End Sub

    ' ------------------------------------------------------------
    ' 9) 空内容 / 坏 JSON → success=False
    ' ------------------------------------------------------------
    Private Sub TestEmptyAndBadJson()
        Console.WriteLine("[BadJson] 空与坏 JSON 拒绝")
        Dim r1 = PptThemeDesignValidator.ValidateAndFix("")
        Check(Not CBool(r1("success")), "空主题 success=False")
        Dim r2 = PptThemeDesignValidator.ValidateAndFix("{not json!!")
        Check(Not CBool(r2("success")), "坏 JSON success=False")
    End Sub

    ' ------------------------------------------------------------
    ' 10) 合法 designTokens（density 枚举 + 显式令牌 + 数字字符串）→ 通过且保留
    ' ------------------------------------------------------------
    Private Sub TestValidDesignTokens()
        Console.WriteLine("[Tokens] 合法 designTokens 校验通过")
        Dim theme = "{""id"":""tokens_ok"",""colors"":{}," &
            """designTokens"":{""density"":""compact"",""titleSize"":24,""margin"":""32"",""ribbonStyle"":""bar"",""ribbonWidth"":220}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "合法 designTokens success=True")
        Dim t = GetFixed(result)
        Dim dt = TryCast(t("designTokens"), JObject)
        Check(dt IsNot Nothing, "fixedTheme 保留 designTokens")
        Check(dt IsNot Nothing AndAlso CInt(dt("titleSize")) = 24, "titleSize=24 保留")
        Check(dt IsNot Nothing AndAlso CInt(dt("margin")) = 32, "数字字符串 margin=""32"" 解析为 32")
    End Sub

    ' ------------------------------------------------------------
    ' 11) 非法 density 枚举 → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestInvalidDensityToken()
        Console.WriteLine("[Tokens] 非法 density 拒绝创建")
        Dim theme = "{""id"":""tokens_bad_density"",""colors"":{}," &
            """designTokens"":{""density"":""huge""}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "density=""huge"" success=False")
        Check(Convert.ToString(result("message")).Contains("density"), "message 提及 designTokens.density")
    End Sub

    ' ------------------------------------------------------------
    ' 12) 字号超范围 → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestOutOfRangeToken()
        Console.WriteLine("[Tokens] 超范围令牌拒绝创建")
        Dim theme = "{""id"":""tokens_out"",""colors"":{}," &
            """designTokens"":{""titleSize"":200,""margin"":5}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "titleSize=200 / margin=5 success=False")
        Check(Convert.ToString(result("message")).Contains("titleSize"), "message 提及 titleSize 超范围")
        Check(Convert.ToString(result("message")).Contains("margin"), "message 提及 margin 超范围")
    End Sub

    ' ------------------------------------------------------------
    ' 13) 非法 ribbonStyle 枚举 → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestInvalidRibbonStyle()
        Console.WriteLine("[Tokens] 非法 ribbonStyle 拒绝创建")
        Dim theme = "{""id"":""tokens_ribbon"",""colors"":{}," &
            """designTokens"":{""ribbonStyle"":""glow""}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "ribbonStyle=""glow"" success=False")
        Check(Convert.ToString(result("message")).Contains("ribbonStyle"), "message 提及 ribbonStyle")
    End Sub

    ' ------------------------------------------------------------
    ' 14) designTokens 非 JSON 对象（字符串）→ 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestNonObjectDesignTokens()
        Console.WriteLine("[Tokens] designTokens 非对象拒绝创建")
        Dim theme = "{""id"":""tokens_nobj"",""colors"":{},""designTokens"":""compact""}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "designTokens 为字符串 success=False")
        Check(Convert.ToString(result("message")).Contains("designTokens"), "message 提及 designTokens 必须是对象")
    End Sub

    ' ------------------------------------------------------------
    ' 15) tech 风格方向展开：density=compact + bar 装饰 + 深色配色
    ' ------------------------------------------------------------
    Private Sub TestTechDirectionPreset()
        Console.WriteLine("[Direction] tech 预设展开")
        Dim theme = "{""id"":""dir_tech"",""styleDirection"":""tech"",""colors"":{}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "tech 方向校验通过")
        Dim t = GetFixed(result)
        Dim dt = TryCast(t("designTokens"), JObject)
        Check(dt IsNot Nothing, "tech 方向补齐 designTokens")
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("density")) = "compact", "density 展开为 compact")
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("ribbonStyle")) = "bar", "ribbonStyle 展开为 bar")
        Dim colors As JObject = TryCast(t("colors"), JObject)
        Dim bg = TryCast(colors("bg"), JArray)
        Check(bg IsNot Nothing AndAlso CInt(bg(0)) = 13 AndAlso CInt(bg(1)) = 18 AndAlso CInt(bg(2)) = 30, "tech 深色背景 [13,18,30] 注入")
        Dim text = TryCast(colors("text"), JArray)
        Check(text IsNot Nothing AndAlso CInt(text(0)) >= 220, "tech 浅色文字注入")
    End Sub

    ' ------------------------------------------------------------
    ' 16) minimal 风格方向：无装饰线 + 宽松边距
    ' ------------------------------------------------------------
    Private Sub TestMinimalDirectionPreset()
        Console.WriteLine("[Direction] minimal 预设展开")
        Dim theme = "{""id"":""dir_minimal"",""styleDirection"":""minimal""}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "minimal 方向校验通过")
        Dim dt = TryCast(GetFixed(result)("designTokens"), JObject)
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("ribbonStyle")) = "none", "minimal ribbonStyle=none")
        Check(dt IsNot Nothing AndAlso CInt(dt("margin")) = 60, "minimal margin=60（宽松）")
        Check(dt IsNot Nothing AndAlso CInt(dt("titleSize")) = 30, "minimal titleSize=30")
    End Sub

    ' ------------------------------------------------------------
    ' 17) 显式字段覆盖 preset：主题给的 titleSize/colors 覆盖预设值
    ' ------------------------------------------------------------
    Private Sub TestDirectionExplicitOverride()
        Console.WriteLine("[Direction] 显式字段覆盖预设")
        Dim theme = "{""id"":""dir_override"",""styleDirection"":""tech""," &
            """colors"":{""bg"":[255,255,255]},""designTokens"":{""titleSize"":36}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "tech+覆盖 校验通过")
        Dim t = GetFixed(result)
        Dim dt = TryCast(t("designTokens"), JObject)
        Check(dt IsNot Nothing AndAlso CInt(dt("titleSize")) = 36, "显式 titleSize=36 覆盖 preset 28")
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("density")) = "compact", "未覆盖的 density 仍取 preset")
        Dim colors As JObject = TryCast(t("colors"), JObject)
        Dim bg = TryCast(colors("bg"), JArray)
        Check(bg IsNot Nothing AndAlso CInt(bg(0)) = 255 AndAlso CInt(bg(1)) = 255 AndAlso CInt(bg(2)) = 255, "显式 bg 覆盖 preset")
    End Sub

    ' ------------------------------------------------------------
    ' 18) 非法 styleDirection → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestInvalidStyleDirection()
        Console.WriteLine("[Direction] 非法 styleDirection 拒绝创建")
        Dim theme = "{""id"":""dir_bad"",""styleDirection"":""fancy"",""colors"":{}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "styleDirection=""fancy"" success=False")
        Check(Convert.ToString(result("message")).Contains("styleDirection"), "message 提及 styleDirection")
    End Sub

    ' ------------------------------------------------------------
    ' 19) 五个预设方向全部零误调（preset 配色自检）
    ' ------------------------------------------------------------
    Private Sub TestAllPresetsZeroAdjustment()
        Console.WriteLine("[Direction] 五个预设方向零误调")
        Dim directions As String() = {"tech", "academic", "business", "minimal", "vivid"}
        For Each d As String In directions
            Dim theme = "{""id"":""dir_" & d & """,""styleDirection"":""" & d & """,""colors"":{}}"
            Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
            Check(CBool(result("success")), d & " 方向校验通过")
            Check(AdjustedCount(result) = 0, d & " 预设配色零自动调整")
        Next
    End Sub

    ' ------------------------------------------------------------
    ' 20) 合法 pageVariants → 通过且保留
    ' ------------------------------------------------------------
    Private Sub TestValidPageVariants()
        Console.WriteLine("[Variants] 合法 pageVariants 校验通过")
        Dim theme = "{""id"":""pv_ok"",""colors"":{}," &
            """designTokens"":{""pageVariants"":{""title"":""modern"",""kpi"":""hero"",""chart"":""insight""}}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(CBool(result("success")), "合法 pageVariants success=True")
        Dim dt = TryCast(GetFixed(result)("designTokens"), JObject)
        Check(dt IsNot Nothing, "fixedTheme 保留 designTokens")
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("pageVariants")("title")) = "modern", "title=modern 保留")
        Check(dt IsNot Nothing AndAlso Convert.ToString(dt("pageVariants")("kpi")) = "hero", "kpi=hero 保留")
    End Sub

    ' ------------------------------------------------------------
    ' 21) 非法页型 key（table 不在支持集）→ 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestInvalidPageVariantKey()
        Console.WriteLine("[Variants] 非法页型 key 拒绝创建")
        Dim theme = "{""id"":""pv_badkey"",""colors"":{}," &
            """designTokens"":{""pageVariants"":{""table"":""compact""}}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "pageVariants.table success=False")
        Check(Convert.ToString(result("message")).Contains("table"), "message 提及不支持页型 table")
    End Sub

    ' ------------------------------------------------------------
    ' 22) 合法页型 + 非法变体值 → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestInvalidPageVariantValue()
        Console.WriteLine("[Variants] 非法变体值拒绝创建")
        Dim theme = "{""id"":""pv_badval"",""colors"":{}," &
            """designTokens"":{""pageVariants"":{""title"":""big"",""kpi"":""row""}}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "title=""big"" success=False")
        Check(Convert.ToString(result("message")).Contains("split"), "message 提示合法值 split/centered/modern")
        Check(Convert.ToString(result("message")).Contains("kpi") = False, "合法的 kpi=row 未报错")
    End Sub

    ' ------------------------------------------------------------
    ' 23) pageVariants 非 JSON 对象 → 拒绝创建
    ' ------------------------------------------------------------
    Private Sub TestNonObjectPageVariants()
        Console.WriteLine("[Variants] pageVariants 非对象拒绝创建")
        Dim theme = "{""id"":""pv_nobj"",""colors"":{}," &
            """designTokens"":{""pageVariants"":""hero""}}"
        Dim result = PptThemeDesignValidator.ValidateAndFix(theme)
        Check(Not CBool(result("success")), "pageVariants 为字符串 success=False")
        Check(Convert.ToString(result("message")).Contains("pageVariants"), "message 提及 pageVariants")
    End Sub

End Module
