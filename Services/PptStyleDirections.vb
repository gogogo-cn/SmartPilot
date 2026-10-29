' ============================================================
' PptStyleDirections.vb — 风格方向预设（styleDirection）
'
' 把「科技感 / 学术 / 商务 / 极简 / 活力」五个风格方向映射为
' 一套合理的默认 designTokens + 配色引导，AI 只需声明方向，
' 再按需覆盖显式字段；未覆盖项取 preset 默认，全部经过
' 与手写主题相同的 WCAG / 令牌校验，保证"合理"由系统兜底。
'
' 展开规则（ExpandPreset）：
'   - theme 含 styleDirection → 合并 preset 的 designTokens 与 colors
'   - 主题显式给出的 designTokens 字段覆盖 preset 对应字段（逐字段）
'   - 主题显式给出的 colors 键覆盖 preset 对应键
'   - preset 自身保证通过校验（零误调，由单测锁定）
' ============================================================
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services

    Public NotInheritable Class PptStyleDirections

        ''' <summary>
        ''' 支持的方向枚举（小写）。未知方向由调用方拒绝。
        ''' </summary>
        Public Shared ReadOnly SupportedDirections As String() = {
            "tech", "academic", "business", "minimal", "vivid"
        }

        ' preset 定义：每个方向含 designTokens（版式令牌）+ colors（配色引导）。
        ' 配色值均已按 WCAG 阈值自检（text/bg≥4.5、primary/bg≥3.0、accent/bg≥1.5、
        ' secondary/bg≥1.3），保证 preset 展开后零误调。
        Private Shared ReadOnly _presets As Dictionary(Of String, JObject) = BuildPresets()

        Private Shared Function BuildPresets() As Dictionary(Of String, JObject)
            Dim presets As New Dictionary(Of String, JObject)()

            ' 科技感：深色底、电光蓝主色、青绿强调、紧凑密度、粗色条装饰
            presets("tech") = New JObject(
                New JProperty("designTokens", New JObject(
                    New JProperty("density", "compact"),
                    New JProperty("ribbonStyle", "bar"),
                    New JProperty("ribbonWidth", 240),
                    New JProperty("titleSize", 28),
                    New JProperty("headingSize", 18),
                    New JProperty("bodySize", 14),
                    New JProperty("smallSize", 11),
                    New JProperty("footerSize", 9),
                    New JProperty("margin", 32),
                    New JProperty("grid", 16),
                    New JProperty("titleTop", 16),
                    New JProperty("contentTop", 96))),
                New JProperty("colors", New JObject(
                    New JProperty("primary", New JArray(0, 140, 255)),
                    New JProperty("secondary", New JArray(0, 110, 180)),
                    New JProperty("accent", New JArray(0, 229, 255)),
                    New JProperty("bg", New JArray(13, 18, 30)),
                    New JProperty("light", New JArray(22, 30, 48)),
                    New JProperty("text", New JArray(228, 238, 255)),
                    New JProperty("gray", New JArray(120, 132, 152)))))

            ' 学术：米白底、深蓝主色、棕褐辅助、舒适密度、细线装饰
            presets("academic") = New JObject(
                New JProperty("designTokens", New JObject(
                    New JProperty("density", "comfortable"),
                    New JProperty("ribbonStyle", "line"),
                    New JProperty("ribbonWidth", 160),
                    New JProperty("titleSize", 26),
                    New JProperty("headingSize", 22),
                    New JProperty("bodySize", 16),
                    New JProperty("smallSize", 12),
                    New JProperty("footerSize", 10),
                    New JProperty("margin", 44),
                    New JProperty("grid", 20),
                    New JProperty("titleTop", 20),
                    New JProperty("contentTop", 104))),
                New JProperty("colors", New JObject(
                    New JProperty("primary", New JArray(31, 78, 121)),
                    New JProperty("secondary", New JArray(139, 94, 60)),
                    New JProperty("accent", New JArray(180, 120, 40)),
                    New JProperty("bg", New JArray(250, 250, 247)),
                    New JProperty("light", New JArray(243, 244, 241)),
                    New JProperty("text", New JArray(35, 40, 45)),
                    New JProperty("gray", New JArray(140, 140, 140)))))

            ' 商务：经典商务蓝（与内置 business_blue 一致），舒适密度、细线装饰
            presets("business") = New JObject(
                New JProperty("designTokens", New JObject(
                    New JProperty("density", "comfortable"),
                    New JProperty("ribbonStyle", "line"),
                    New JProperty("ribbonWidth", 160),
                    New JProperty("titleSize", 26),
                    New JProperty("headingSize", 22),
                    New JProperty("bodySize", 16),
                    New JProperty("smallSize", 12),
                    New JProperty("footerSize", 10),
                    New JProperty("margin", 40),
                    New JProperty("grid", 20),
                    New JProperty("titleTop", 20),
                    New JProperty("contentTop", 100))),
                New JProperty("colors", New JObject(
                    New JProperty("primary", New JArray(31, 78, 121)),
                    New JProperty("secondary", New JArray(190, 167, 91)),
                    New JProperty("accent", New JArray(255, 165, 0)),
                    New JProperty("bg", New JArray(245, 247, 250)),
                    New JProperty("light", New JArray(239, 243, 248)),
                    New JProperty("text", New JArray(33, 37, 41)),
                    New JProperty("gray", New JArray(150, 150, 150)))))

            ' 极简：纯白底、近黑文字、无装饰线、宽松密度、大标题
            presets("minimal") = New JObject(
                New JProperty("designTokens", New JObject(
                    New JProperty("density", "spacious"),
                    New JProperty("ribbonStyle", "none"),
                    New JProperty("titleSize", 30),
                    New JProperty("headingSize", 20),
                    New JProperty("bodySize", 15),
                    New JProperty("smallSize", 12),
                    New JProperty("footerSize", 10),
                    New JProperty("margin", 60),
                    New JProperty("grid", 24),
                    New JProperty("titleTop", 24),
                    New JProperty("contentTop", 116))),
                New JProperty("colors", New JObject(
                    New JProperty("primary", New JArray(45, 45, 45)),
                    New JProperty("secondary", New JArray(110, 110, 110)),
                    New JProperty("accent", New JArray(200, 200, 200)),
                    New JProperty("bg", New JArray(255, 255, 255)),
                    New JProperty("light", New JArray(245, 245, 245)),
                    New JProperty("text", New JArray(35, 35, 35)),
                    New JProperty("gray", New JArray(160, 160, 160)))))

            ' 活力：暖白底、橙红主色、亮粉强调、宽松密度、粗色条装饰
            presets("vivid") = New JObject(
                New JProperty("designTokens", New JObject(
                    New JProperty("density", "spacious"),
                    New JProperty("ribbonStyle", "bar"),
                    New JProperty("ribbonWidth", 200),
                    New JProperty("titleSize", 30),
                    New JProperty("headingSize", 22),
                    New JProperty("bodySize", 16),
                    New JProperty("smallSize", 12),
                    New JProperty("footerSize", 10),
                    New JProperty("margin", 48),
                    New JProperty("grid", 24),
                    New JProperty("titleTop", 24),
                    New JProperty("contentTop", 116))),
                New JProperty("colors", New JObject(
                    New JProperty("primary", New JArray(190, 60, 40)),
                    New JProperty("secondary", New JArray(255, 170, 0)),
                    New JProperty("accent", New JArray(255, 60, 120)),
                    New JProperty("bg", New JArray(255, 250, 240)),
                    New JProperty("light", New JArray(255, 240, 230)),
                    New JProperty("text", New JArray(50, 40, 40)),
                    New JProperty("gray", New JArray(160, 150, 145)))))

            Return presets
        End Function

        ''' <summary>
        ''' 展开风格方向：theme 含 styleDirection 时，把 preset 的 designTokens/colors
        ''' 合并进主题（主题显式字段逐字段覆盖 preset）。未知方向抛 ArgumentException。
        ''' 不修改传入对象，返回新 JObject。
        ''' </summary>
        Public Shared Function ExpandPreset(theme As JObject) As JObject
            If theme Is Nothing Then Return theme
            Dim direction = theme.Value(Of String)("styleDirection")
            If String.IsNullOrWhiteSpace(direction) Then Return theme

            Dim dir = direction.Trim().ToLowerInvariant()
            If Not _presets.ContainsKey(dir) Then
                Throw New ArgumentException("styleDirection 仅支持: " & String.Join(" / ", SupportedDirections))
            End If

            Dim merged As New JObject(theme)
            merged.Remove("styleDirection")

            Dim preset As JObject = _presets(dir)

            ' 合并 colors：主题显式键覆盖 preset 键
            Dim presetColors As JObject = TryCast(preset("colors"), JObject)
            If presetColors IsNot Nothing Then
                Dim colors As JObject = TryCast(merged("colors"), JObject)
                If colors Is Nothing Then
                    colors = New JObject()
                    merged("colors") = colors
                End If
                For Each p As JProperty In presetColors.Properties()
                    If colors(p.Name) Is Nothing Then
                        colors(p.Name) = p.Value.DeepClone()
                    End If
                Next
            End If

            ' 合并 designTokens：主题显式字段覆盖 preset 字段
            Dim presetTokens As JObject = TryCast(preset("designTokens"), JObject)
            If presetTokens IsNot Nothing Then
                Dim tokens As JObject = TryCast(merged("designTokens"), JObject)
                If tokens Is Nothing Then
                    tokens = New JObject()
                    merged("designTokens") = tokens
                End If
                For Each p As JProperty In presetTokens.Properties()
                    If tokens(p.Name) Is Nothing Then
                        tokens(p.Name) = p.Value.DeepClone()
                    End If
                Next
            End If

            Return merged
        End Function

    End Class

End Namespace
