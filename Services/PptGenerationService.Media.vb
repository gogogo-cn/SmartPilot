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

        ' ------------------------------------------------------------
        ' 自由形状页：任意形状组合（矩形/椭圆/箭头/流程图/标注/线条）+ 阴影 + 文字
        ' spec: { type:"shapes", title, items:[{type,x,y,w,h,fill,textColor,text,shadow,rotate}...] }
        ' 坐标缺失时按网格自动布局；items 也可用 shapes 字段作为别名。
        ' ------------------------------------------------------------
        Private Shared Sub BuildShapesSlide(pres As Object, toRelease As List(Of Object),
                                            spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "形状示意图"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim items As New List(Of JToken)()
            Dim itemsArr As JArray = spec.Value(Of JArray)("items")
            If itemsArr Is Nothing Then itemsArr = spec.Value(Of JArray)("shapes")
            If itemsArr IsNot Nothing Then
                For Each it As JToken In itemsArr
                    items.Add(it)
                Next
            End If
            If items.Count = 0 Then
                AddTextBox(s, toRelease, "（未提供形状，可传 items 数组：{type,x,y,w,h,fill,text,shadow}）",
                           40, ContentAreaTop(), 880, 40, 14, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            Dim idx As Integer = 0
            For Each it As JToken In items
                Dim shapeType As String = ""
                Dim text As String = ""
                Dim fillToken As String = ""
                Dim textColorToken As String = ""
                Dim shadow As Boolean = False
                Dim rotate As Double = 0
                Dim lineW As Integer = 0
                Dim x As Integer = -1
                Dim y As Integer = -1
                Dim w As Integer = 200
                Dim h As Integer = 70
                Dim isLine As Boolean = False

                Dim obj As JObject = TryCast(it, JObject)
                If obj IsNot Nothing Then
                    shapeType = obj.Value(Of String)("type")
                    text = obj.Value(Of String)("text")
                    If String.IsNullOrWhiteSpace(text) Then text = obj.Value(Of String)("label")
                    fillToken = obj.Value(Of String)("fill")
                    textColorToken = obj.Value(Of String)("textColor")
                    shadow = obj.Value(Of Boolean)("shadow")
                    rotate = obj.Value(Of Double)("rotate")
                    lineW = obj.Value(Of Integer)("weight")
                    x = If(obj("x") Is Nothing, -1, obj.Value(Of Integer)("x"))
                    y = If(obj("y") Is Nothing, -1, obj.Value(Of Integer)("y"))
                    w = If(obj("w") Is Nothing, 200, obj.Value(Of Integer)("w"))
                    h = If(obj("h") Is Nothing, 70, obj.Value(Of Integer)("h"))
                Else
                    shapeType = "rect"
                    text = it.ToString()
                End If
                If String.IsNullOrWhiteSpace(shapeType) Then shapeType = "rect"
                Dim st As String = shapeType.ToLowerInvariant()
                If st = "line" Then isLine = True

                ' 自动布局：未指定坐标时按 4 列网格顺排
                If x < 0 OrElse y < 0 Then
                    Dim col As Integer = idx Mod 4
                    Dim row As Integer = idx \ 4
                    x = MARGIN + col * (w + 40)
                    y = 140 + row * (h + 40)
                End If

                If isLine Then
                    Dim lineColor As Long = ResolveColor(fillToken, theme, theme.Gray)
                    If lineW <= 0 Then lineW = 2
                    AddLineShape(s, toRelease, x, y, x + w, y + h, lineColor, lineW)
                Else
                    Dim autoShape As Integer = MapShapeType(shapeType)
                    Dim fillColor As Long = ResolveColor(fillToken, theme, If(idx Mod 2 = 0, theme.Primary, theme.Secondary))
                    Dim shape As Object = AddAutoShape(s, toRelease, autoShape, x, y, w, h, fillColor)
                    If shape IsNot Nothing Then
                        Try
                            If rotate <> 0 Then shape.Rotation = rotate
                        Catch
                        End Try
                        If shadow Then ApplyShapeShadow(shape)
                        Dim tc As Long = ResolveColor(textColorToken, theme, TextColorOn(fillColor, theme))
                        If Not String.IsNullOrWhiteSpace(text) Then
                            SetShapeText(shape, text, If(text.Length > 12, 13, 15), tc, theme.Font)
                        End If
                    End If
                End If
                idx += 1
            Next
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' ------------------------------------------------------------
        ' 图片页：插入本地图片（位置/大小/圆角/边框/阴影/等比适配 + 图注）
        ' spec: { type:"picture", title, images:[{src,caption,x,y,w,h,rounded,border,shadow,fit}...] }
        ' 或单图：{ type:"picture", title, src, caption, x, y, w, h, ... }
        ' ------------------------------------------------------------
        Private Shared Sub BuildPictureSlide(pres As Object, toRelease As List(Of Object),
                                             spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "图片展示"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim items As New List(Of JObject)()
            Dim imagesArr As JArray = spec.Value(Of JArray)("images")
            If imagesArr IsNot Nothing Then
                For Each it As JToken In imagesArr
                    Dim obj As JObject = TryCast(it, JObject)
                    If obj IsNot Nothing Then items.Add(obj)
                Next
            End If
            Dim singleSrc As String = spec.Value(Of String)("src")
            If items.Count = 0 AndAlso Not String.IsNullOrWhiteSpace(singleSrc) Then
                Dim singleItem As New JObject()
                singleItem("src") = singleSrc
                singleItem("caption") = spec.Value(Of String)("caption")
                If spec("x") IsNot Nothing Then singleItem("x") = spec.Value(Of Integer)("x")
                If spec("y") IsNot Nothing Then singleItem("y") = spec.Value(Of Integer)("y")
                If spec("w") IsNot Nothing Then singleItem("w") = spec.Value(Of Integer)("w")
                If spec("h") IsNot Nothing Then singleItem("h") = spec.Value(Of Integer)("h")
                If spec("rounded") IsNot Nothing Then singleItem("rounded") = spec.Value(Of Boolean)("rounded")
                If spec("border") IsNot Nothing Then singleItem("border") = spec("border").DeepClone()
                If spec("shadow") IsNot Nothing Then singleItem("shadow") = spec.Value(Of Boolean)("shadow")
                If spec("fit") IsNot Nothing Then singleItem("fit") = spec.Value(Of String)("fit")
                items.Add(singleItem)
            End If
            If items.Count = 0 Then
                AddTextBox(s, toRelease, "（未提供图片路径，可传 images 数组 {src,caption} 或单图 src 字段）",
                           40, ContentAreaTop(), 880, 40, 14, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            Dim idx As Integer = 0
            For Each it As JObject In items
                Dim src As String = it.Value(Of String)("src")
                If String.IsNullOrWhiteSpace(src) Then src = it.Value(Of String)("image")
                Dim caption As String = it.Value(Of String)("caption")
                Dim x As Integer = If(it("x") Is Nothing, -1, it.Value(Of Integer)("x"))
                Dim y As Integer = If(it("y") Is Nothing, -1, it.Value(Of Integer)("y"))
                Dim w As Integer = If(it("w") Is Nothing, 0, it.Value(Of Integer)("w"))
                Dim h As Integer = If(it("h") Is Nothing, 0, it.Value(Of Integer)("h"))
                Dim rounded As Boolean = it.Value(Of Boolean)("rounded")
                Dim shadow As Boolean = it.Value(Of Boolean)("shadow")
                Dim fit As String = it.Value(Of String)("fit")
                Dim borderColor As Long = 0
                Dim hasBorder As Boolean = False
                If it("border") IsNot Nothing Then
                    If it("border").Type = JTokenType.Boolean Then
                        hasBorder = it.Value(Of Boolean)("border")
                        If hasBorder Then borderColor = theme.Secondary
                    Else
                        hasBorder = True
                        borderColor = ResolveColor(it.Value(Of String)("border"), theme, theme.Secondary)
                    End If
                End If

                ' 自动网格布局：单图居中大图，多图 2 列
                If x < 0 OrElse y < 0 OrElse w <= 0 OrElse h <= 0 Then
                    If items.Count = 1 Then
                        x = 180 : y = 150 : w = 600 : h = 420
                    Else
                        Dim col As Integer = idx Mod 2
                        Dim row As Integer = idx \ 2
                        x = MARGIN + col * 460
                        y = 120 + row * 360
                        w = 420 : h = 300
                    End If
                End If

                If String.IsNullOrWhiteSpace(src) Then
                    Dim ph As Object = AddAutoShape(s, toRelease, 1, x, y, w, h, theme.Light)
                    If ph IsNot Nothing Then SetShapeText(ph, "缺少图片路径", 14, theme.Gray, theme.Font)
                Else
                    AddPictureOnSlide(s, toRelease, src, x, y, w, h, rounded, hasBorder, borderColor, shadow, fit, theme)
                End If

                If Not String.IsNullOrWhiteSpace(caption) Then
                    AddCenteredTextBox(s, toRelease, caption, x, y + h + 6, w, 24, 12, False, theme.Text, theme.Font)
                End If
                idx += 1
            Next
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' ------------------------------------------------------------
        ' SmartArt 替代页：形状组合模拟流程/循环/层次/矩阵
        ' spec: { type:"smartart", title, layout:"flow|cycle|hierarchy|matrix", items:[string|{title,desc}] }
        ' ------------------------------------------------------------
        Private Shared Sub BuildSmartArtSlide(pres As Object, toRelease As List(Of Object),
                                              spec As JObject, theme As PptTheme)
            Dim s As Object = pres.Slides.Add(pres.Slides.Count + 1, 12)
            SetSlideBackground(s, toRelease, theme.Bg)
            Dim title As String = spec.Value(Of String)("title")
            If String.IsNullOrWhiteSpace(title) Then title = "结构示意图"
            AddTextBox(s, toRelease, title, TitleLeft(), TitleTop(), TitleWidth(), TitleHeight(), 26, True, theme.Primary, theme.Font)
            AddRibbon(s, toRelease, theme.Secondary, 40, TitleRibbonTop(), 160, 4)

            Dim layout As String = spec.Value(Of String)("layout")
            If String.IsNullOrWhiteSpace(layout) Then layout = spec.Value(Of String)("type")
            If String.IsNullOrWhiteSpace(layout) Then layout = "flow"

            Dim nodes As New List(Of KeyValuePair(Of String, String))()
            Dim itemsArr As JArray = spec.Value(Of JArray)("items")
            If itemsArr IsNot Nothing Then
                For Each it As JToken In itemsArr
                    Dim obj As JObject = TryCast(it, JObject)
                    If obj IsNot Nothing Then
                        Dim t As String = obj.Value(Of String)("title")
                        If String.IsNullOrWhiteSpace(t) Then t = obj.Value(Of String)("text")
                        If String.IsNullOrWhiteSpace(t) Then t = obj.Value(Of String)("label")
                        nodes.Add(New KeyValuePair(Of String, String)(t, obj.Value(Of String)("desc")))
                    Else
                        nodes.Add(New KeyValuePair(Of String, String)(it.ToString(), ""))
                    End If
                Next
            End If
            If nodes.Count = 0 Then
                AddTextBox(s, toRelease, "（未提供节点，可传 items 数组：字符串或 {title,desc}）",
                           40, ContentAreaTop(), 880, 40, 14, False, theme.Gray, theme.Font)
                AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
                Return
            End If

            Select Case layout.ToLowerInvariant()
                Case "cycle"
                    BuildSmartArtCycle(s, toRelease, nodes, theme)
                Case "hierarchy", "tree"
                    BuildSmartArtHierarchy(s, toRelease, nodes, theme)
                Case "matrix"
                    BuildSmartArtMatrix(s, toRelease, nodes, theme)
                Case Else
                    BuildSmartArtFlow(s, toRelease, nodes, theme)
            End Select
            AddFooter(s, toRelease, pres.Slides.Count, theme.Gray, theme.Font)
        End Sub

        ' ---- SmartArt 流程：横向圆角盒 + 箭头 ----
        Private Shared Sub BuildSmartArtFlow(s As Object, toRelease As List(Of Object),
                                             nodes As List(Of KeyValuePair(Of String, String)), theme As PptTheme)
            Dim n As Integer = Math.Min(nodes.Count, 6)
            Dim arrowW As Integer = 40
            Dim boxW As Integer = (880 - (n - 1) * arrowW) / n
            Dim boxH As Integer = 84
            Dim y As Integer = 270
            For i As Integer = 0 To n - 1
                Dim x As Integer = MARGIN + i * (boxW + arrowW)
                Dim fill As Long = If(i Mod 2 = 0, theme.Primary, theme.Secondary)
                Dim box As Object = AddAutoShape(s, toRelease, 5, x, y, boxW, boxH, fill)
                If box IsNot Nothing Then
                    SetShapeText(box, SmartArtNodeText(nodes(i)), 15, RGB(255, 255, 255), theme.Font)
                End If
                If i < n - 1 Then
                    AddAutoShape(s, toRelease, 33, x + boxW + 8, y + (boxH - 36) / 2, arrowW - 16, 36, theme.Accent)
                End If
            Next
        End Sub

        ' ---- SmartArt 循环：环形椭圆 + 中心枢纽 + 连接线 ----
        Private Shared Sub BuildSmartArtCycle(s As Object, toRelease As List(Of Object),
                                              nodes As List(Of KeyValuePair(Of String, String)), theme As PptTheme)
            Dim n As Integer = Math.Min(nodes.Count, 6)
            Dim cx As Integer = 480
            Dim cy As Integer = 360
            Dim radius As Integer = 210
            Dim ovalW As Integer = 160
            Dim ovalH As Integer = 64
            ' 先画连接线（在椭圆下方，避免遮挡）
            For i As Integer = 0 To n - 1
                Dim angle As Double = -Math.PI / 2 + i * 2.0 * Math.PI / n
                Dim ox As Double = cx + radius * Math.Cos(angle)
                Dim oy As Double = cy + radius * Math.Sin(angle)
                AddLineShape(s, toRelease, cx, cy, CInt(Math.Round(ox)), CInt(Math.Round(oy)), theme.Gray, 1)
            Next
            Dim hub As Object = AddAutoShape(s, toRelease, 9, cx - 55, cy - 55, 110, 110, theme.Accent)
            If hub IsNot Nothing Then SetShapeText(hub, "循环", 18, RGB(255, 255, 255), theme.Font)
            For i As Integer = 0 To n - 1
                Dim angle As Double = -Math.PI / 2 + i * 2.0 * Math.PI / n
                Dim ox As Integer = CInt(Math.Round(cx + radius * Math.Cos(angle)))
                Dim oy As Integer = CInt(Math.Round(cy + radius * Math.Sin(angle)))
                Dim fill As Long = If(i Mod 3 = 0, theme.Primary, If(i Mod 3 = 1, theme.Secondary, theme.Light))
                Dim oval As Object = AddAutoShape(s, toRelease, 9, ox - ovalW / 2, oy - ovalH / 2, ovalW, ovalH, fill)
                If oval IsNot Nothing Then
                    Dim tc As Long = If(fill = theme.Light, theme.Text, RGB(255, 255, 255))
                    SetShapeText(oval, SmartArtNodeText(nodes(i)), 12, tc, theme.Font)
                End If
            Next
        End Sub

        ' ---- SmartArt 层次：根节点 + 两行子节点 + 连线 ----
        Private Shared Sub BuildSmartArtHierarchy(s As Object, toRelease As List(Of Object),
                                                  nodes As List(Of KeyValuePair(Of String, String)), theme As PptTheme)
            Dim n As Integer = Math.Min(nodes.Count, 9)
            If n = 0 Then Return
            ' 根节点
            Dim rootW As Integer = 240
            Dim rootH As Integer = 64
            Dim rootX As Integer = (DESIGN_W - rootW) / 2
            Dim rootY As Integer = Math.Max(110, ContentAreaTop() + 10)
            Dim root As Object = AddAutoShape(s, toRelease, 5, rootX, rootY, rootW, rootH, theme.Primary)
            If root IsNot Nothing Then SetShapeText(root, SmartArtNodeText(nodes(0)), 17, RGB(255, 255, 255), theme.Font)

            ' 第二层（最多 4 个）
            Dim row2 As Integer = Math.Min(n - 1, 4)
            Dim boxW As Integer = 200
            Dim boxH As Integer = 64
            Dim row2Y As Integer = 260
            Dim gap2 As Integer = (DESIGN_W - 2 * MARGIN - row2 * boxW) / Math.Max(1, row2 + 1)
            Dim row2Xs As New List(Of Integer)()
            For i As Integer = 0 To row2 - 1
                Dim x As Integer = MARGIN + gap2 + i * (boxW + gap2)
                row2Xs.Add(x)
                ' 连线：根底 → 子顶
                AddLineShape(s, toRelease, rootX + rootW / 2, rootY + rootH, x + boxW / 2, row2Y, theme.Gray, 1)
                Dim box As Object = AddAutoShape(s, toRelease, 5, x, row2Y, boxW, boxH, theme.Secondary)
                If box IsNot Nothing Then SetShapeText(box, SmartArtNodeText(nodes(1 + i)), 13, RGB(255, 255, 255), theme.Font)
            Next

            ' 第三层（最多 4 个，对应第二层索引取模）
            If n - 1 > row2 Then
                Dim row3 As Integer = Math.Min(n - 1 - row2, 4)
                Dim row3Y As Integer = 430
                For i As Integer = 0 To row3 - 1
                    Dim parentIdx As Integer = i Mod Math.Max(1, row2)
                    Dim x As Integer = row2Xs(parentIdx)
                    AddLineShape(s, toRelease, x + boxW / 2, row2Y + boxH, x + boxW / 2, row3Y, theme.Gray, 1)
                    Dim box As Object = AddAutoShape(s, toRelease, 5, x, row3Y, boxW, boxH, theme.Light)
                    If box IsNot Nothing Then SetShapeText(box, SmartArtNodeText(nodes(1 + row2 + i)), 13, theme.Text, theme.Font)
                Next
            End If
        End Sub

        ' ---- SmartArt 矩阵：2x2 象限 ----
        Private Shared Sub BuildSmartArtMatrix(s As Object, toRelease As List(Of Object),
                                               nodes As List(Of KeyValuePair(Of String, String)), theme As PptTheme)
            Dim count As Integer = Math.Min(nodes.Count, 4)
            Dim boxW As Integer = 400
            Dim boxH As Integer = 200
            Dim fills As Long() = {theme.Primary, theme.Secondary, theme.Accent, theme.Light}
            For i As Integer = 0 To count - 1
                Dim col As Integer = i Mod 2
                Dim row As Integer = i \ 2
                Dim x As Integer = MARGIN + col * (boxW + 40)
                Dim y As Integer = Math.Max(140, ContentAreaTop() + 40) + row * (boxH + 40)
                Dim fill As Long = fills(i Mod 4)
                Dim box As Object = AddAutoShape(s, toRelease, 5, x, y, boxW, boxH, fill)
                If box IsNot Nothing Then
                    Dim tc As Long = If(fill = theme.Light, theme.Text, RGB(255, 255, 255))
                    SetShapeText(box, SmartArtNodeText(nodes(i)), 15, tc, theme.Font)
                End If
            Next
        End Sub

        Private Shared Function SmartArtNodeText(node As KeyValuePair(Of String, String)) As String
            If String.IsNullOrWhiteSpace(node.Key) Then Return "（节点）"
            If String.IsNullOrWhiteSpace(node.Value) Then Return node.Key
            Return node.Key & vbCrLf & node.Value
        End Function

        ' ------------------------------------------------------------
        ' 动画：页切换 + 进入动画（挂到全部页型）
        ' spec: transition(fade/push/wipe/dissolve/cut/random/none)、animation(appear/fade/rise/float/grow/none)
        ' ------------------------------------------------------------
        Private Shared Sub ApplySlideAnimation(slide As Object, spec As JObject, theme As PptTheme)
            ' 页切换（默认淡入淡出）
            Try
                slide.SlideShowTransition.EntryEffect = MapTransition(spec.Value(Of String)("transition"))
                Try
                    slide.SlideShowTransition.Duration = 0.6
                Catch
                End Try
            Catch
            End Try

            ' 进入动画：对所有形状挂同种进入效果（withPrevious 同时淡入）
            Dim animation As String = spec.Value(Of String)("animation")
            If String.Equals(animation, "none", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(animation, "off", StringComparison.OrdinalIgnoreCase) Then
                Return
            End If
            Dim eff As Integer = MapEntranceEffect(animation)
            Try
                Dim seq As Object = slide.TimeLine.MainSequence
                Dim shapeCount As Integer = 0
                Try
                    shapeCount = slide.Shapes.Count
                Catch
                    Return
                End Try
                For i As Integer = 1 To shapeCount
                    Dim shp As Object = Nothing
                    Try
                        shp = slide.Shapes(i)
                        If shp IsNot Nothing Then
                            seq.AddEffect(shp, eff, 2, i) ' msoAnimTriggerWithPrevious=2
                        End If
                    Catch
                    Finally
                        Try
                            If shp IsNot Nothing Then ComReleaseHelper.Release(shp)
                        Catch
                        End Try
                    End Try
                Next
            Catch
            End Try
        End Sub

        Private Shared Function MapTransition(t As String) As Integer
            If String.IsNullOrWhiteSpace(t) Then Return 17 ' ppEffectFade
            Select Case t.Trim().ToLowerInvariant()
                Case "none", "off" : Return 0 ' ppEffectNone
                Case "cut" : Return 2 ' ppEffectCut
                Case "dissolve" : Return 12 ' ppEffectDissolve
                Case "push" : Return 23 ' ppEffectPushRight
                Case "wipe" : Return 41 ' ppEffectWipeRight
                Case "cover" : Return 5 ' ppEffectCoverRight
                Case "random" : Return 19 ' ppEffectRandom
                Case Else : Return 17
            End Select
        End Function

        Private Shared Function MapEntranceEffect(a As String) As Integer
            If String.IsNullOrWhiteSpace(a) Then Return 10 ' msoAnimEffectFade
            Select Case a.Trim().ToLowerInvariant()
                Case "appear" : Return 1 ' msoAnimEffectAppear
                Case "rise", "riseup" : Return 6 ' msoAnimEffectRiseUp
                Case "float", "floatup" : Return 8 ' msoAnimEffectFloatUp
                Case "grow" : Return 5 ' msoAnimEffectGrowAndTurn
                Case Else : Return 10
            End Select
        End Function

        ' ------------------------------------------------------------
        ' 通用形状辅助：颜色令牌 / 自动形状 / 线条 / 阴影 / 文字
        ' ------------------------------------------------------------
        Private Shared Function MapShapeType(t As String) As Integer
            Select Case t.Trim().ToLowerInvariant()
                Case "rect", "rectangle", "square" : Return 1 ' msoShapeRectangle
                Case "roundrect", "rounded", "round" : Return 5 ' msoShapeRoundedRectangle
                Case "oval", "ellipse", "circle" : Return 9 ' msoShapeOval
                Case "arrow", "arrowright" : Return 33 ' msoShapeRightArrow
                Case "arrowleft" : Return 34 ' msoShapeLeftArrow
                Case "arrowup" : Return 32 ' msoShapeUpArrow
                Case "arrowdown" : Return 36 ' msoShapeDownArrow
                Case "flow", "process" : Return 61 ' msoShapeFlowchartProcess
                Case "flowdecision", "decision" : Return 63 ' msoShapeFlowchartDecision
                Case "flowterminator", "terminator" : Return 74 ' msoShapeFlowchartTerminator
                Case "flowdocument", "document" : Return 60 ' msoShapeFlowchartDocument
                Case "callout", "calloutrect" : Return 106 ' msoShapeRoundedRectangularCallout
                Case "calloutoval", "ovalcallout" : Return 109 ' msoShapeOvalCallout
                Case "diamond" : Return 4 ' msoShapeDiamond
                Case "triangle" : Return 7 ' msoShapeIsocelesTriangle
                Case Else : Return 1
            End Select
        End Function

        Private Shared Function AddAutoShape(s As Object, toRelease As List(Of Object), shapeType As Integer,
                                             left As Integer, top As Integer, width As Integer, height As Integer,
                                             fillColor As Long) As Object
            Dim shape As Object = Nothing
            Try
                shape = s.Shapes.AddShape(shapeType, SX(left), SY(top), SW(width), SH(height))
                If fillColor >= 0 Then
                    shape.Fill.ForeColor.RGB = fillColor
                Else
                    Try
                        shape.Fill.Visible = 0 ' msoFalse
                    Catch
                    End Try
                End If
                Try
                    shape.Line.Visible = False
                Catch
                End Try
                toRelease.Add(shape)
                Return shape
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddAutoShape 失败(type=" & shapeType.ToString() & "): " & ex.Message)
                Try
                    If shape IsNot Nothing Then ComReleaseHelper.Release(shape)
                Catch
                End Try
                Return Nothing
            End Try
        End Function

        Private Shared Sub AddLineShape(s As Object, toRelease As List(Of Object),
                                        x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer,
                                        color As Long, weight As Integer)
            Dim line As Object = Nothing
            Try
                line = s.Shapes.AddLine(SX(x1), SY(y1), SX(x2), SY(y2))
                line.Line.ForeColor.RGB = color
                line.Line.Weight = weight
                toRelease.Add(line)
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddLineShape 失败: " & ex.Message)
                Try
                    If line IsNot Nothing Then ComReleaseHelper.Release(line)
                Catch
                End Try
            End Try
        End Sub

        Private Shared Sub AddPictureOnSlide(s As Object, toRelease As List(Of Object),
                                             src As String, x As Integer, y As Integer, w As Integer, h As Integer,
                                             rounded As Boolean, hasBorder As Boolean, borderColor As Long,
                                             shadow As Boolean, fit As String, theme As PptTheme)
            Dim filePath As String = ResolveImagePath(src)
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
                Dim ph As Object = AddAutoShape(s, toRelease, 1, x, y, w, h, theme.Light)
                If ph IsNot Nothing Then SetShapeText(ph, "图片不存在：" & src, 12, theme.Gray, theme.Font)
                Return
            End If

            Dim px As Integer = x
            Dim py As Integer = y
            Dim pw As Integer = w
            Dim ph2 As Integer = h
            If String.Equals(fit, "fit", StringComparison.OrdinalIgnoreCase) Then
                ' 等比适配：读取图片原始尺寸，按比例缩放并居中
                Dim natW As Integer = 0
                Dim natH As Integer = 0
                Try
                    Using img As System.Drawing.Image = System.Drawing.Image.FromFile(filePath)
                        natW = img.Width
                        natH = img.Height
                    End Using
                Catch
                End Try
                If natW > 0 AndAlso natH > 0 Then
                    Dim scale As Double = Math.Min(w / natW, h / natH)
                    pw = CInt(Math.Round(natW * scale))
                    ph2 = CInt(Math.Round(natH * scale))
                    px = x + (w - pw) / 2
                    py = y + (h - ph2) / 2
                End If
            End If

            Dim pic As Object = Nothing
            Try
                If rounded Then
                    pic = s.Shapes.AddShape(5, SX(px), SY(py), SW(pw), SH(ph2)) ' 圆角矩形 + 图片填充（裁剪为圆角）
                    pic.Fill.UserPicture(filePath)
                Else
                    pic = s.Shapes.AddPicture(filePath, 0, -1, SX(px), SY(py), SW(pw), SH(ph2)) ' msoFalse=0, msoTrue=-1
                End If
                If pic IsNot Nothing Then
                    If hasBorder Then
                        Try
                            pic.Line.Visible = True
                            pic.Line.ForeColor.RGB = borderColor
                            pic.Line.Weight = 2
                        Catch
                        End Try
                    Else
                        Try
                            pic.Line.Visible = False
                        Catch
                        End Try
                    End If
                    If shadow Then ApplyShapeShadow(pic)
                    toRelease.Add(pic)
                End If
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "AddPictureOnSlide 失败: " & ex.Message)
                Try
                    If pic IsNot Nothing Then ComReleaseHelper.Release(pic)
                Catch
                End Try
                Dim ph As Object = AddAutoShape(s, toRelease, 1, x, y, w, h, theme.Light)
                If ph IsNot Nothing Then SetShapeText(ph, "图片加载失败", 14, theme.Gray, theme.Font)
            End Try
        End Sub

        ''' <summary>
        ''' 图片路径解析：本地路径直接用；http(s) URL 先下载到临时目录再返回本地路径。
        ''' </summary>
        Private Shared Function ResolveImagePath(src As String) As String
            If String.IsNullOrWhiteSpace(src) Then Return ""
            Dim trimmed As String = src.Trim()
            If Not (trimmed.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
                    trimmed.StartsWith("https://", StringComparison.OrdinalIgnoreCase)) Then
                Return trimmed
            End If
            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(trimmed, UriKind.Absolute, uri) Then Return ""
            Try
                Dim ext As String = ""
                Dim urlPath As String = uri.AbsolutePath
                If Not String.IsNullOrWhiteSpace(urlPath) Then ext = Path.GetExtension(urlPath)
                If String.IsNullOrWhiteSpace(ext) Then ext = ".jpg"
                Dim tmp As String = Path.Combine(Path.GetTempPath(), "ppt_img_" & Guid.NewGuid().ToString("N") & ext)
                Using wc As New Net.WebClient()
                    wc.DownloadFile(uri, tmp)
                End Using
                If File.Exists(tmp) Then Return tmp
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "图片 URL 下载失败: " & ex.Message)
            End Try
            Return ""
        End Function

        Private Shared Sub ApplyShapeShadow(shape As Object, Optional offset As Integer = 3)
            Try
                shape.Shadow.Visible = True
                Try
                    shape.Shadow.Transparency = 0.45
                Catch
                End Try
                Try
                    shape.Shadow.OffsetX = offset
                    shape.Shadow.OffsetY = offset
                Catch
                End Try
            Catch
            End Try
        End Sub

        Private Shared Sub SetShapeText(shape As Object, text As String, fontSize As Integer, color As Long,
                                        fontName As String)
            Try
                shape.TextFrame.TextRange.Text = text
                shape.TextFrame.TextRange.Font.Name = fontName
                shape.TextFrame.TextRange.Font.Size = fontSize
                shape.TextFrame.TextRange.Font.Color.RGB = color
                Try
                    shape.TextFrame.TextRange.Font.Bold = True
                Catch
                End Try
                Try
                    shape.TextFrame.VerticalAnchor = 3 ' msoAnchorMiddle
                Catch
                End Try
                Try
                    shape.TextFrame.MarginLeft = 0
                    shape.TextFrame.MarginRight = 0
                    shape.TextFrame.MarginTop = 0
                    shape.TextFrame.MarginBottom = 0
                Catch
                End Try
            Catch
            End Try
        End Sub

        ''' <summary>
        ''' 颜色令牌解析：#RRGGBB 或主题令牌（primary/secondary/accent/bg/light/text/gray/white/none）
        ''' </summary>
        Private Shared Function ResolveColor(token As String, theme As PptTheme, fallback As Long) As Long
            If String.IsNullOrWhiteSpace(token) Then Return fallback
            Dim t As String = token.Trim()
            Select Case t.ToLowerInvariant()
                Case "primary" : Return theme.Primary
                Case "secondary" : Return theme.Secondary
                Case "accent" : Return theme.Accent
                Case "bg", "background" : Return theme.Bg
                Case "light" : Return theme.Light
                Case "text" : Return theme.Text
                Case "gray", "grey" : Return theme.Gray
                Case "white" : Return RGB(255, 255, 255)
                Case "none", "transparent" : Return -1
                Case Else
                    If t.StartsWith("#") Then t = t.Substring(1)
                    If t.Length = 6 Then
                        Try
                            Dim r As Integer = Convert.ToInt32(t.Substring(0, 2), 16)
                            Dim g As Integer = Convert.ToInt32(t.Substring(2, 2), 16)
                            Dim b As Integer = Convert.ToInt32(t.Substring(4, 2), 16)
                            Return RGB(r, g, b)
                        Catch
                        End Try
                    End If
                    Return fallback
            End Select
        End Function

        ''' <summary>
        ''' 根据填充色亮度决定文字颜色：浅色填充用主题文字色，深色填充用白色。
        ''' </summary>
        Private Shared Function TextColorOn(fillColor As Long, theme As PptTheme) As Long
            Dim r As Integer = (fillColor >> 16) And &HFF
            Dim g As Integer = (fillColor >> 8) And &HFF
            Dim b As Integer = fillColor And &HFF
            Dim luminance As Double = r * 0.299 + g * 0.587 + b * 0.114
            If luminance > 170 Then Return theme.Text
            Return RGB(255, 255, 255)
        End Function
    End Class
End Namespace
