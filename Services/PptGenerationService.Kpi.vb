Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

Partial Public Class PptGenerationService

        Private Shared Sub BuildTextBullets(s As Object, toRelease As List(Of Object),
                                            lines As List(Of String), theme As PptTheme)
            Dim startY As Integer = ContentAreaTop()
            Dim lineH As Integer = 52
            For i As Integer = 0 To lines.Count - 1
                Dim y As Integer = startY + i * lineH
                If y > 660 Then Exit For
                AddRoundedRect(s, toRelease, MARGIN, y + 8, 12, 12, theme.Secondary, 3)
                AddTextBox(s, toRelease, lines(i), MARGIN + 28, y, 812, 44, 17, False, theme.Text, theme.Font)
            Next
        End Sub
        Private Shared Sub BuildTextSlide(pres As Object, toRelease As List(Of Object),
                                          spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "要点"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim lines As New List(Of String)()
            Dim bullets As JArray = spec.Value(Of JArray)("bullets")
            If bullets IsNot Nothing Then
                For Each b As JToken In bullets
                    ' 兼容对象型 bullet（{label,value}）：规范化成 "label：value"，避免整段 JSON 上屏
                    Dim txt As String = FormatBulletItem(b)
                    If Not String.IsNullOrWhiteSpace(txt) Then lines.Add(txt)
                Next
            End If
            If lines.Count = 0 Then
                Dim content As String = spec.Value(Of String)("content")
                If Not String.IsNullOrWhiteSpace(content) Then
                    For Each line As String In content.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                        lines.Add(line.Trim())
                    Next
                End If
            End If
            If lines.Count = 0 Then lines.Add("（无内容）")

            ' 容量限制：按版式容量截断要点数，避免溢出正文区
            Dim cap As PptCapacityProfile = CapacityOf(_activeLayout)
            Dim maxBullets As Integer = If(cap Is Nothing, 8, cap.MaxBullets)
            If lines.Count > maxBullets Then
                lines = lines.Take(maxBullets).ToList()
            End If

            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "bullets"
            Select Case slideVariant.ToLowerInvariant()
                Case "cards"
                    BuildTextCards(s, toRelease, lines, theme)
                Case "numbered"
                    BuildTextNumbered(s, toRelease, lines, theme)
                Case Else
                    BuildTextBullets(s, toRelease, lines, theme)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub
        Private Shared Sub BuildTextNumbered(s As Object, toRelease As List(Of Object),
                                             lines As List(Of String), theme As PptTheme)
            Dim startY As Integer = ContentAreaTop() + 20
            Dim lineH As Integer = 72
            For i As Integer = 0 To lines.Count - 1
                Dim y As Integer = startY + i * lineH
                If y > 660 Then Exit For
                Dim dot As Object = Nothing
                Try
                    dot = s.Shapes.AddShape(9, SX(MARGIN), SY(y), SW(40), SH(40)) ' msoShapeOval = 9
                    dot.Fill.ForeColor.RGB = theme.Primary
                    Try
                        dot.Line.Visible = False
                    Catch
                    End Try
                    toRelease.Add(dot)
                Catch
                End Try
                AddCenteredTextBox(s, toRelease, (i + 1).ToString(), MARGIN, y, 40, 40, 16, True, RGB(255, 255, 255), theme.Font)
                AddTextBox(s, toRelease, lines(i), MARGIN + 64, y + 6, 816, 56, 17, False, theme.Text, theme.Font)
            Next
        End Sub
        Private Shared Sub BuildTextCards(s As Object, toRelease As List(Of Object),
                                          lines As List(Of String), theme As PptTheme)
            Dim count As Integer = Math.Min(lines.Count, 4)
            Dim cardW As Integer = 420
            Dim cardH As Integer = 150
            Dim gapX As Integer = 40
            Dim gapY As Integer = 30
            Dim startX As Integer = MARGIN
            Dim startY As Integer = ContentAreaTop() + 20
            For i As Integer = 0 To count - 1
                Dim col As Integer = i Mod 2
                Dim rowIdx As Integer = i \ 2
                Dim x As Integer = startX + col * (cardW + gapX)
                Dim y As Integer = startY + rowIdx * (cardH + gapY)
                AddRoundedRect(s, toRelease, x, y, cardW, cardH, theme.Light, 14)
                AddRibbon(s, toRelease, theme.Primary, x, y, 6, cardH)
                AddTextBox(s, toRelease, (i + 1).ToString("00"), x + 26, y + 24, 60, 40, 26, True, theme.Secondary, theme.Font)
                AddTextBox(s, toRelease, lines(i), x + 96, y + 26, cardW - 122, cardH - 52, 15, False, theme.Text, theme.Font)
            Next
        End Sub
        Private Shared Sub BuildKpiGrid(s As Object, toRelease As List(Of Object),
                                        kpiList As List(Of KeyValuePair(Of String, String)),
                                        highlightList As List(Of Boolean), theme As PptTheme)
            Dim count As Integer = Math.Min(kpiList.Count, 6)
            Dim cardW As Integer = 280
            Dim gapX As Integer = 20
            Dim gapY As Integer = 20
            ' 高度自适应：避让母版顶部装饰后，按行数压缩卡片高度，防止溢出页面
            Dim rows As Integer = (count + 2) \ 3
            Dim cardH As Integer = Math.Max(120, Math.Min(180, (DESIGN_H - ContentAreaTop() - 80) \ rows - gapY))
            For i As Integer = 0 To count - 1
                Dim col As Integer = i Mod 3
                Dim rowIdx As Integer = i \ 3
                Dim x As Integer = MARGIN + col * (cardW + gapX)
                Dim y As Integer = ContentAreaTop() + rowIdx * (cardH + gapY)
                Dim isHi As Boolean = i < highlightList.Count AndAlso highlightList(i)
                AddRoundedRect(s, toRelease, x, y, cardW, cardH, theme.Light, 14)
                If isHi Then
                    AddGradientRect(s, toRelease, x, y, cardW, 8, theme.Accent, theme.Secondary, 0)
                    AddBadge(s, toRelease, "TOP", x + cardW - 72, y + 20, 52, 24, theme.Accent, RGB(255, 255, 255), 10, theme.Font)
                Else
                    AddGradientRect(s, toRelease, x, y, cardW, 8, theme.Primary, theme.Secondary, 0)
                End If
                ' 内部相对定位：随卡片高度缩放，避免压缩后溢出
                AddTextBox(s, toRelease, kpiList(i).Value, x + GRID, y + CInt(cardH * 0.22), cardW - 40, 60, 32, True, If(isHi, theme.Accent, theme.Primary), theme.Font)
                AddTextBox(s, toRelease, kpiList(i).Key, x + GRID, y + CInt(cardH * 0.62), cardW - 40, 40, 16, False, theme.Text, theme.Font)
            Next
        End Sub
        Private Shared Sub BuildKpiSlide(pres As Object, toRelease As List(Of Object),
                                         spec As JObject, theme As PptTheme,
                                         headers As List(Of String),
                                         rows As List(Of Dictionary(Of String, Object)))
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "关键指标"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim kpiList As New List(Of KeyValuePair(Of String, String))()
            Dim highlightList As New List(Of Boolean)()
            Dim items As JArray = spec.Value(Of JArray)("items")
            If items IsNot Nothing AndAlso items.Count > 0 Then
                For Each item As JToken In items
                    Dim obj As JObject = TryCast(item, JObject)
                    If obj Is Nothing Then Continue For
                    Dim label As String = obj.Value(Of String)("label")
                    If String.IsNullOrWhiteSpace(label) Then Continue For
                    kpiList.Add(New KeyValuePair(Of String, String)(label, If(obj.Value(Of String)("value"), "")))
                    highlightList.Add(obj.Value(Of Boolean)("highlight"))
                Next
            End If
            If kpiList.Count = 0 Then
                ' 自动从数据生成：数值列求和
                For Each h As String In headers
                    Dim sum As Double = 0
                    Dim hasNum As Boolean = False
                    For Each row As Dictionary(Of String, Object) In rows
                        Dim d As Double
                        If Double.TryParse(Convert.ToString(GetRowValue(row, h)), d) Then
                            sum += d
                            hasNum = True
                        End If
                    Next
                    If hasNum Then
                        kpiList.Add(New KeyValuePair(Of String, String)(h, FormatNumber(sum, 0)))
                    End If
                Next
            End If
            If kpiList.Count = 0 Then
                AddTextBox(s, toRelease, "没有可用于指标的数据", 40, ContentAreaTop(), 880, 40, 16, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            ' 防御：单 item value 含多值（换行/顿号分隔，如趋势预测的多个未来值）时拆分为独立 KPI 项，
            ' 避免单卡片多值溢出文本框覆盖 label（"预测趋势"重叠的根因）。
            Dim expanded As New List(Of KeyValuePair(Of String, String))()
            Dim expandedHi As New List(Of Boolean)()
            For i As Integer = 0 To kpiList.Count - 1
                Dim parts As String() = kpiList(i).Value.Split(New String() {vbCrLf, vbLf, vbCr, "、"}, StringSplitOptions.RemoveEmptyEntries)
                If parts.Length > 1 Then
                    For Each p As String In parts
                        expanded.Add(New KeyValuePair(Of String, String)(kpiList(i).Key, p.Trim()))
                        expandedHi.Add(If(i < highlightList.Count, highlightList(i), False))
                    Next
                Else
                    expanded.Add(kpiList(i))
                    expandedHi.Add(If(i < highlightList.Count, highlightList(i), False))
                End If
            Next
            kpiList = expanded
            highlightList = expandedHi

            Dim slideVariant As String = spec.Value(Of String)("variant")
            If String.IsNullOrWhiteSpace(slideVariant) Then slideVariant = "grid"
            Select Case slideVariant.ToLowerInvariant()
                Case "hero"
                    BuildKpiHero(s, toRelease, kpiList, highlightList, theme)
                Case "row"
                    BuildKpiRow(s, toRelease, kpiList, highlightList, theme)
                Case Else
                    BuildKpiGrid(s, toRelease, kpiList, highlightList, theme)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub
        Private Shared Sub BuildKpiRow(s As Object, toRelease As List(Of Object),
                                       kpiList As List(Of KeyValuePair(Of String, String)),
                                       highlightList As List(Of Boolean), theme As PptTheme)
            Dim count As Integer = Math.Min(kpiList.Count, 4)
            Dim gapX As Integer = 20
            Dim cardW As Integer = (DESIGN_W - 2 * MARGIN - (count - 1) * gapX) / count
            Dim cardH As Integer = 200
            ' 内容避让：卡片行下移到母版顶部装饰区之下（无装饰时维持原位置）
            Dim y As Integer = Math.Max(260, ContentAreaTop() + 60)
            For i As Integer = 0 To count - 1
                Dim x As Integer = MARGIN + i * (cardW + gapX)
                Dim isHi As Boolean = i < highlightList.Count AndAlso highlightList(i)
                AddRoundedRect(s, toRelease, x, y, cardW, cardH, theme.Light, 14)
                AddGradientRect(s, toRelease, x, y, cardW, 8, If(isHi, theme.Accent, theme.Primary), theme.Secondary, 0)
                AddTextBox(s, toRelease, kpiList(i).Value, x + GRID, y + 44, cardW - 40, 60, 30, True, If(isHi, theme.Accent, theme.Primary), theme.Font)
                AddTextBox(s, toRelease, kpiList(i).Key, x + GRID, y + 118, cardW - 40, 40, 16, False, theme.Text, theme.Font)
            Next
        End Sub
        Private Shared Sub BuildKpiHero(s As Object, toRelease As List(Of Object),
                                        kpiList As List(Of KeyValuePair(Of String, String)),
                                        highlightList As List(Of Boolean), theme As PptTheme)
            Dim count As Integer = Math.Min(kpiList.Count, 4)
            Dim heroW As Integer = 420
            ' 高度自适应：避让母版顶部装饰后压缩主卡高度，防止溢出页面（底部留 80 给页脚）
            Dim heroH As Integer = Math.Min(520, DESIGN_H - ContentAreaTop() - 80)
            AddGradientRect(s, toRelease, MARGIN, ContentAreaTop(), heroW, heroH, theme.Primary, theme.Secondary, 16)
            AddTextBox(s, toRelease, kpiList(0).Value, MARGIN + 30, ContentAreaTop() + CInt(heroH * 0.23), heroW - 60, 90, 48, True, RGB(255, 255, 255), theme.Font)
            AddTextBox(s, toRelease, kpiList(0).Key, MARGIN + 30, ContentAreaTop() + CInt(heroH * 0.46), heroW - 60, 40, 20, False, RGB(240, 245, 255), theme.Font)
            AddRibbon(s, toRelease, RGB(255, 255, 255), MARGIN + 30, ContentAreaTop() + CInt(heroH * 0.58), 80, 4)
            Dim rightX As Integer = MARGIN + heroW + 40
            Dim rightW As Integer = DESIGN_W - MARGIN - rightX
            Dim cardH As Integer = 160
            Dim gapY As Integer = 20
            For i As Integer = 1 To count - 1
                Dim y As Integer = ContentAreaTop() + (i - 1) * (cardH + gapY)
                Dim isHi As Boolean = i < highlightList.Count AndAlso highlightList(i)
                AddRoundedRect(s, toRelease, rightX, y, rightW, cardH, theme.Light, 14)
                AddRibbon(s, toRelease, If(isHi, theme.Accent, theme.Primary), rightX, y, 6, cardH)
                AddTextBox(s, toRelease, kpiList(i).Value, rightX + 30, y + 28, rightW - 60, 50, 30, True, If(isHi, theme.Accent, theme.Primary), theme.Font)
                AddTextBox(s, toRelease, kpiList(i).Key, rightX + 30, y + 92, rightW - 60, 30, 16, False, theme.Text, theme.Font)
            Next
        End Sub
End Class
End Namespace
