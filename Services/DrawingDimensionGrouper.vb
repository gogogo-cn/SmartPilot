Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions

Namespace Services

    ''' <summary>
    ''' 机械图识别后的成块和归类。几何检出仍在 GdtSymbolRecognizer。
    ''' 这里把同一条标注的文字框收成一块，并写上种类、标称值和公差。
    ''' </summary>
    Public Class DrawingDimensionGrouper

        Private Shared ReadOnly GdtChars As String = "⊕⌖◎∥⊥▱⌒⌓∠↗⌭"
        Private Shared ReadOnly GdtNames As String() = {
            "直线度", "平面度", "圆度", "圆柱度", "线轮廓度", "面轮廓度",
            "平行度", "垂直度", "倾斜度", "位置度", "同轴度", "同心度",
            "对称度", "圆跳动", "全跳动"
        }
        Private Shared ReadOnly DeviationToken As New Regex("^(?:±\s*\d+(?:\.\d+)?|[+-]\d+(?:\.\d+)?|0)$", RegexOptions.Compiled)
        Private Shared ReadOnly GradeToken As New Regex("^(?:[A-Z]{1,2}\d{1,2}|[a-z]{1,2}\d{1,2}|\d[HhGg](?:/\d[HhGg])?)(?:/[A-Za-z]{1,2}\d{1,2})?$", RegexOptions.Compiled)
        Private Shared ReadOnly PrefixToken As New Regex("^(?:SØ|SR|Ø|ø|Φ|φ|⌀|∅|□|R|C|M|\d+\s*[xX×])$", RegexOptions.Compiled)

        Friend Shared Function JoinOnLeader(members As List(Of OcrElement)) As OcrElement
            Return BuildBlock(members, "", members(0))
        End Function

        Public Shared Sub Apply(doc As OcrDocument)
            If doc Is Nothing OrElse doc.Elements Is Nothing OrElse doc.Elements.Count = 0 Then Return
            Dim items = doc.Elements.Where(Function(el) el IsNot Nothing AndAlso el.BoundingBox IsNot Nothing).ToList()
            If items.Count = 0 Then Return

            Dim used As New HashSet(Of OcrElement)()
            Dim output As New List(Of OcrElement)()

            For Each el In items
                If used.Contains(el) Then Continue For
                If Not IsFcfWrapper(el) Then Continue For
                Dim inside = items.Where(Function(other) other IsNot el AndAlso Not used.Contains(other) AndAlso CenterInside(other, el.BoundingBox)).ToList()
                ' 标题栏外框也会被看成公差框。里面有英文单词或格子太多时，不把整框收成一条尺寸。
                If Not IsCalloutFrame(inside) Then Continue For
                used.Add(el)
                For Each member In inside
                    used.Add(member)
                Next
                output.Add(BuildBlock(inside, SymbolFromFcf(el.Text), el))
            Next

            Dim rest = items.Where(Function(el) Not used.Contains(el)).ToList()
            output.AddRange(GroupLoose(rest))

            For Each el In doc.Elements
                If el IsNot Nothing AndAlso el.BoundingBox Is Nothing Then output.Add(el)
            Next

            doc.Elements = output
        End Sub

        Private Shared Function GroupLoose(items As List(Of OcrElement)) As List(Of OcrElement)
            ' 名义值先当种子，裸偏差排到最后。
            ' 按 Top 排时上偏差（Top 更小）会先当种子，而 CanAttach 只允许偏差贴名义、
            ' 不允许偏差把名义吸过来——公差栈于是被拆成「+0.2」和「22.4 -0」两块，
            ' 正是截图里「22.4-0 +02」各占一行的原因。
            Dim pending = items.
                OrderBy(Function(el) If(IsDeviation(Compact(el.Text)), 1, 0)).
                ThenBy(Function(el) el.BoundingBox.Top).
                ThenBy(Function(el) el.BoundingBox.Left).ToList()
            Dim used As New HashSet(Of OcrElement)()
            Dim output As New List(Of OcrElement)()

            For Each seed In pending
                If used.Contains(seed) Then Continue For
                If seed.Type = OcrElementType.Balloon OrElse seed.Type = OcrElementType.DatumCircle Then
                    used.Add(seed)
                    output.Add(TagExisting(seed))
                    Continue For
                End If

                Dim members As New List(Of OcrElement) From {seed}
                used.Add(seed)
                Dim grew = True
                While grew
                    grew = False
                    For Each candidate In pending
                        If used.Contains(candidate) Then Continue For
                        If candidate.Type = OcrElementType.Balloon OrElse candidate.Type = OcrElementType.DatumCircle Then Continue For
                        If Not CanAttach(members, candidate) Then Continue For
                        members.Add(candidate)
                        used.Add(candidate)
                        grew = True
                    Next
                End While
                output.Add(BuildBlock(members, "", seed))
            Next
            Return output
        End Function

        Private Shared Function CanAttach(members As List(Of OcrElement), candidate As OcrElement) As Boolean
            Dim nominal = NominalMember(members)
            If nominal Is Nothing Then Return False
            Dim text = Compact(candidate.Text)
            If text = "" Then Return False

            If IsDeviation(text) Then
                If text = "0" AndAlso Not members.Any(Function(el) IsSignedDeviation(Compact(el.Text))) Then Return False
                Return IsStackedDeviation(nominal, candidate) AndAlso Not HasDeviationOnSameSide(members, nominal, candidate)
            End If

            If IsGrade(text) AndAlso SameRow(nominal, candidate) AndAlso SitsJustRight(members, candidate) AndAlso HasDigit(JoinTexts(members)) Then
                Return True
            End If

            Dim right = Rightmost(members)
            If IsPrefix(Compact(right.Text)) AndAlso Regex.IsMatch(text, "^\d") AndAlso SameRow(right, candidate) AndAlso SitsJustRight(members, candidate) Then
                Return True
            End If

            If Regex.IsMatch(Compact(right.Text), "^\d+\s*[xX×]$") AndAlso (text.StartsWith("Ø") OrElse text.StartsWith("M") OrElse Regex.IsMatch(text, "^\d")) AndAlso SameRow(right, candidate) AndAlso SitsJustRight(members, candidate) Then
                Return True
            End If

            ' 单个字母不是形位公差。只有组里已经有公差符号，才把旁边的数值和基准字母收进来。
            If HasGdtSymbol(members) AndAlso IsGdtAttachment(text) Then
                If SameRow(nominal, candidate) AndAlso SitsJustRight(members, candidate) Then Return True
                If StackedInFrame(members, candidate) Then Return True
            End If

            If IsRoughPiece(JoinTexts(members)) AndAlso IsRoughValue(text) AndAlso Near(nominal, candidate) Then Return True
            Return False
        End Function

        Private Shared Function BuildBlock(members As List(Of OcrElement), symbol As String, anchor As OcrElement) As OcrElement
            If members Is Nothing OrElse members.Count = 0 Then
                Dim alone = CloneShell(anchor)
                alone.Text = If(symbol <> "", symbol, Normalize(anchor.Text))
                ApplyCallout(alone, Classify(alone.Text))
                Return alone
            End If

            Dim nominal = NominalMember(members)
            ' 球标排在最前：解析脚本的 readBalloon 只认行首的 ①/⬡/(1)，
            ' 顺序乱了球标序号就丢（期望「① Ø25±0.02」而不是「Ø25±0.02 ①」）。
            Dim body = members.Where(Function(el) Not IsDeviation(Compact(el.Text))).
                OrderBy(Function(el) If(el.Type = OcrElementType.Balloon, 0, 1)).
                ThenBy(Function(el) BandKey(el, nominal)).
                ThenBy(Function(el) el.BoundingBox.Left).ToList()
            Dim uppers = members.Where(Function(el) IsDeviation(Compact(el.Text)) AndAlso CenterY(el) < CenterY(nominal)).OrderBy(Function(el) el.BoundingBox.Left).ToList()
            Dim lowers = members.Where(Function(el) IsDeviation(Compact(el.Text)) AndAlso CenterY(el) >= CenterY(nominal)).OrderBy(Function(el) el.BoundingBox.Left).ToList()

            Dim parts As New List(Of String)()
            If symbol <> "" AndAlso Not body.Any(Function(el) If(el.Text, "").Contains(symbol)) Then parts.Add(symbol)
            For Each el In body
                Dim piece = If(el.Text, "").Trim()
                If piece <> "" AndAlso Not piece.StartsWith("FCF:") Then parts.Add(piece)
            Next
            For Each el In uppers
                parts.Add(Compact(el.Text))
            Next
            For Each el In lowers
                parts.Add(Compact(el.Text))
            Next

            Dim block = CloneShell(anchor)
            block.Text = String.Join(" ", parts.Where(Function(p) p <> "")).Trim()
            block.BoundingBox = UnionBox(members)
            block.Confidence = members.Average(Function(el) el.Confidence)
            If members.Any(Function(el) el.Type = OcrElementType.Roughness) Then block.Type = OcrElementType.Roughness
            ApplyCallout(block, Classify(block.Text))
            Return block
        End Function

        Private Shared Sub ApplyCallout(el As OcrElement, callout As Callout)
            If callout Is Nothing OrElse callout.Kind = "" Then Return
            el.Kind = callout.Kind
            el.Nominal = callout.Nominal
            el.Lower = callout.Lower
            el.Upper = callout.Upper
            el.Datum = callout.Datum
            If el.Type = OcrElementType.Paragraph OrElse el.Type = OcrElementType.Image OrElse el.Type = OcrElementType.GdtSymbol Then
                el.Type = TypeForKind(callout.Kind)
            End If
        End Sub

        Private Shared Function TagExisting(el As OcrElement) As OcrElement
            If el.Type = OcrElementType.Balloon Then
                el.Kind = "球标"
                el.Nominal = Compact(el.Text)
            ElseIf el.Type = OcrElementType.DatumCircle Then
                el.Kind = "基准"
                el.Datum = Compact(el.Text)
                el.Nominal = el.Datum
            End If
            Return el
        End Function

        Private Shared Function TypeForKind(kind As String) As OcrElementType
            Select Case kind
                Case "形位公差" : Return OcrElementType.GdtSymbol
                Case "表面粗糙度" : Return OcrElementType.Roughness
                Case "基准" : Return OcrElementType.DatumCircle
                Case "球标" : Return OcrElementType.Balloon
                Case "未注公差" : Return OcrElementType.Paragraph
                Case Else : Return OcrElementType.Dimension
            End Select
        End Function

        Private Class Callout
            Public Property Kind As String = ""
            Public Property Nominal As String = ""
            Public Property Lower As String = ""
            Public Property Upper As String = ""
            Public Property Datum As String = ""
        End Class

        Private Shared Function Classify(raw As String) As Callout
            Dim text = Normalize(raw)
            Dim result As New Callout()
            If text = "" OrElse IsSkipped(text) Then Return result

            If Regex.IsMatch(text, "未注(?:尺寸|形位)?公差|GB/T\s*1804|GB/T\s*1184") Then
                result.Kind = "未注公差"
                result.Nominal = text
                Return result
            End If

            If Regex.IsMatch(text, "(?:表面)?粗糙度|(?:^|[^A-Za-z])R\s*[azqy](?![A-Za-z])|[√∨]") Then
                result.Kind = "表面粗糙度"
                result.Nominal = Regex.Replace(text, "\s+", "")
                Dim range = Regex.Match(text, "(\d+(?:\.\d+)?)\s*[/~～]\s*(\d+(?:\.\d+)?)")
                If range.Success Then
                    result.Upper = range.Groups(1).Value
                    result.Lower = range.Groups(2).Value
                End If
                Return result
            End If

            Dim gdtName = GdtNames.FirstOrDefault(Function(name) text.Contains(name))
            Dim gdtSymbol = FirstGdtSymbol(text)
            If gdtName IsNot Nothing OrElse gdtSymbol <> "" Then
                result.Kind = "形位公差"
                result.Nominal = If(gdtName, SymbolName(gdtSymbol))
                Dim zone = Regex.Match(text, "(?:Ø\s*)?(\d+(?:\.\d+)?)")
                If zone.Success Then result.Upper = zone.Groups(1).Value
                result.Datum = TrailingDatums(text)
                Return result
            End If

            If Regex.IsMatch(text, "^(?:基准\s*)?[A-Z](?:\s+[A-Z]){0,3}$") AndAlso text.Contains("基准") Then
                result.Kind = "基准"
                result.Datum = Regex.Replace(text, "基准", "").Trim()
                result.Nominal = result.Datum
                Return result
            End If

            FillDeviation(result, text)
            Dim body = StripDeviation(text)

            If Regex.IsMatch(body, "^(?:M|G|Tr|Rc|Rp|NPT)\s*\d", RegexOptions.IgnoreCase) Then
                result.Kind = "螺纹"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "Ø|ø|Φ|φ|⌀|∅|SØ") Then
                result.Kind = "直径"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "^(?:SR|R)\s*\d") Then
                result.Kind = "半径"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "^C\s*\d|×\s*45\s*°|\d+(?:\.\d+)?\s*[xX×]\s*\d+(?:\.\d+)?\s*°") Then
                result.Kind = "倒角"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "\d+(?:\.\d+)?\s*°") Then
                result.Kind = "角度"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "^□\s*\d") Then
                result.Kind = "线性尺寸"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "^\d+\s*[:：]\s*\d+$") Then
                result.Kind = "锥度"
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            If Regex.IsMatch(body, "[A-Z]\d{1,2}\s*/\s*[a-z]\d{1,2}") Then
                result.Kind = If(Regex.IsMatch(body, "Ø|Φ|φ|⌀"), "直径", "配合")
                result.Nominal = Regex.Replace(body, "\s+", "")
                Return result
            End If
            ' 只有整段都像一条尺寸时才标线性尺寸。标题栏、注释里带个数字不算。
            If Regex.IsMatch(body, "^\(?\d+(?:\.\d+)?\)?(?:[A-Z]{1,2}\d{1,2})?$") Then
                result.Kind = "线性尺寸"
                result.Nominal = Regex.Replace(body, "\s+", "")
            End If
            Return result
        End Function

        Private Shared Sub FillDeviation(result As Callout, text As String)
            Dim sym = Regex.Match(text, "±\s*(\d+(?:\.\d+)?)")
            If sym.Success Then
                result.Lower = "-" & sym.Groups(1).Value
                result.Upper = "+" & sym.Groups(1).Value
                Return
            End If
            Dim plus = Regex.Match(text, "(?:^|\s)(\+\d+(?:\.\d+)?)")
            Dim minus = Regex.Match(text, "(?:^|\s)(-\d+(?:\.\d+)?)")
            Dim zero = Regex.IsMatch(text, "(?:^|\s)0(?:\s|$)")
            If plus.Success Then result.Upper = plus.Groups(1).Value
            If minus.Success Then result.Lower = minus.Groups(1).Value
            If zero AndAlso result.Lower = "" AndAlso result.Upper <> "" Then result.Lower = "0"
            If zero AndAlso result.Upper = "" AndAlso result.Lower <> "" Then result.Upper = "0"
        End Sub

        Private Shared Function StripDeviation(text As String) As String
            Dim body = Regex.Replace(text, "±\s*\d+(?:\.\d+)?", " ")
            body = Regex.Replace(body, "(?:^|\s)[+-]\d+(?:\.\d+)?", " ")
            body = Regex.Replace(body, "(?:^|\s)0(?:\s|$)", " ")
            Return Regex.Replace(body, "\s+", " ").Trim()
        End Function

        Private Shared Function TrailingDatums(text As String) As String
            Dim match = Regex.Match(text, "(?:Ø\s*)?\d+(?:\.\d+)?(?:\s*\(?[MLS]\)?)?\s+([A-Z](?:\s+[A-Z]){0,5})\s*$")
            If Not match.Success Then Return ""
            Return Regex.Replace(match.Groups(1).Value, "\s+", " ").Trim()
        End Function

        Private Shared Function IsSkipped(text As String) As Boolean
            If Regex.IsMatch(text, "\d{4}[.\-/年]\d{1,2}") Then Return True
            If Regex.IsMatch(text, "^[A-Za-z]{1,8}[-.][A-Za-z0-9./_-]{2,}$") Then Return True
            Dim chinese = text.Count(Function(ch) ch >= "一"c AndAlso ch <= "龥"c)
            If chinese > 0 AndAlso Not text.Contains("未注") AndAlso Not text.Contains("公差") AndAlso Not text.Contains("粗糙") AndAlso GdtNames.All(Function(name) Not text.Contains(name)) Then
                Return True
            End If
            Return False
        End Function

        Private Shared Function IsFcfWrapper(el As OcrElement) As Boolean
            Return el.Text IsNot Nothing AndAlso el.Text.StartsWith("FCF:")
        End Function

        Private Shared Function SymbolFromFcf(text As String) As String
            Dim body = If(text, "").Substring(Math.Min(4, If(text, "").Length))
            Dim found = FirstGdtSymbol(body)
            If found <> "" Then Return found
            If body.IndexOf("Perpendicular", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "⊥"
            If body.IndexOf("Parallel", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "∥"
            If body.IndexOf("Position", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "⊕"
            If body.IndexOf("Runout", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "↗"
            If body.IndexOf("Profile", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "⌓"
            If body.IndexOf("Coaxial", StringComparison.OrdinalIgnoreCase) >= 0 OrElse body.IndexOf("Symmetry", StringComparison.OrdinalIgnoreCase) >= 0 Then Return "◎"
            Return ""
        End Function

        Private Shared Function SymbolName(symbol As String) As String
            Select Case symbol
                Case "⊕", "⌖" : Return "位置度"
                Case "◎" : Return "同轴度"
                Case "∥" : Return "平行度"
                Case "⊥" : Return "垂直度"
                Case "▱" : Return "平面度"
                Case "⌒" : Return "线轮廓度"
                Case "⌓" : Return "面轮廓度"
                Case "∠" : Return "倾斜度"
                Case "↗" : Return "圆跳动"
                Case "⌭" : Return "圆柱度"
                Case Else : Return "形位公差"
            End Select
        End Function

        Private Shared Function FirstGdtSymbol(text As String) As String
            For Each ch In If(text, "")
                If GdtChars.IndexOf(ch) >= 0 Then Return ch.ToString()
            Next
            Return ""
        End Function

        Private Shared Function IsDeviation(text As String) As Boolean
            Return DeviationToken.IsMatch(text)
        End Function

        Private Shared Function IsSignedDeviation(text As String) As Boolean
            Return Regex.IsMatch(text, "^[±+-]")
        End Function

        Private Shared Function IsGrade(text As String) As Boolean
            Return GradeToken.IsMatch(text)
        End Function

        Private Shared Function IsPrefix(text As String) As Boolean
            Return PrefixToken.IsMatch(text)
        End Function

        Private Shared Function HasGdtSymbol(members As List(Of OcrElement)) As Boolean
            Return members.Any(Function(el) FirstGdtSymbol(el.Text) <> "" OrElse GdtNames.Any(Function(name) If(el.Text, "").Contains(name)))
        End Function

        Private Shared Function IsGdtAttachment(text As String) As Boolean
            If text.Length = 1 AndAlso GdtChars.IndexOf(text(0)) >= 0 Then Return True
            If Regex.IsMatch(text, "^[A-Z](?:\s+[A-Z]){0,3}$") Then Return True
            If Regex.IsMatch(text, "^(?:Ø)?0\.\d+(?:\s*[MLS])?$") Then Return True
            Return False
        End Function

        Private Shared Function IsCalloutFrame(members As List(Of OcrElement)) As Boolean
            If members Is Nothing OrElse members.Count = 0 OrElse members.Count > 8 Then Return False
            Dim callouts = 0
            For Each el In members
                If Regex.IsMatch(If(el.Text, ""), "[A-Za-z]{4,}") Then Return False
                Dim text = Compact(el.Text)
                If text = "" OrElse IsDeviation(text) OrElse el.Type = OcrElementType.Balloon Then Continue For
                If Regex.IsMatch(text, "^(?:[A-Z]{1,2}\d{1,2}|[a-z]{1,2}\d{1,2})$") AndAlso Not Regex.IsMatch(text, "^(?:Ra|Rz|Rq|Rt)\d", RegexOptions.IgnoreCase) Then Continue For
                If Regex.IsMatch(text, "\d") Then callouts += 1
            Next
            ' 框里并排出现 11.3、0.05、Rz25，那是三条尺寸，不是一个公差框。
            Return callouts <= 1
        End Function

        Private Shared Function IsRoughPiece(text As String) As Boolean
            Return Regex.IsMatch(text, "粗糙|[√∨]|(?:^|[^A-Za-z])R\s*[azqy]", RegexOptions.IgnoreCase)
        End Function

        Private Shared Function IsRoughValue(text As String) As Boolean
            Return Regex.IsMatch(text, "^(?:R\s*[azqy]\s*)?\d+(?:\.\d+)?(?:\s*[/~～]\s*\d+(?:\.\d+)?)?$", RegexOptions.IgnoreCase)
        End Function

        Private Shared Function HasDigit(text As String) As Boolean
            Return Regex.IsMatch(If(text, ""), "\d")
        End Function

        Private Shared Function NominalMember(members As List(Of OcrElement)) As OcrElement
            Return members.
                OrderByDescending(Function(el) If(el.Type = OcrElementType.Balloon OrElse IsDeviation(Compact(el.Text)), 0, 1)).
                ThenByDescending(Function(el) Width(el)).
                FirstOrDefault()
        End Function

        Private Shared Function Rightmost(members As List(Of OcrElement)) As OcrElement
            Return members.OrderByDescending(Function(el) el.BoundingBox.Right).First()
        End Function

        Private Shared Function HasDeviationOnSameSide(members As List(Of OcrElement), nominal As OcrElement, candidate As OcrElement) As Boolean
            Dim above = CenterY(candidate) < CenterY(nominal)
            Return members.Any(Function(el) el IsNot candidate AndAlso IsDeviation(Compact(el.Text)) AndAlso (CenterY(el) < CenterY(nominal)) = above)
        End Function

        Private Shared Function IsStackedDeviation(nominal As OcrElement, candidate As OcrElement) As Boolean
            Dim nh = Math.Max(8, Height(nominal))
            Dim ch = Height(candidate)
            If ch > nh * 0.95 AndAlso Not IsSignedDeviation(Compact(candidate.Text)) Then Return False
            Dim dy = Math.Abs(CenterY(candidate) - CenterY(nominal))
            If dy < nh * 0.2 OrElse dy > nh * 1.35 Then Return False
            Dim cx = CenterX(candidate)
            If cx < nominal.BoundingBox.Left - nh * 0.3 OrElse cx > nominal.BoundingBox.Right + nh Then Return False
            ' 公差就写在名义值正上/正下方，必须和名义值水平重叠。
            ' 少了这条：合并块的框是整条公差栈的高度（80px），dy 阈值按 nh*1.35 放大到 108px，
            ' 会把隔着一条尺寸的另一个 +0.2 也吸进来——「22.4 +0.2 -0 +0.2」就是这么来的。
            Dim overlap = Math.Min(nominal.BoundingBox.Right, candidate.BoundingBox.Right) -
                          Math.Max(nominal.BoundingBox.Left, candidate.BoundingBox.Left)
            Return overlap > 0
        End Function

        Private Shared Function StackedInFrame(members As List(Of OcrElement), candidate As OcrElement) As Boolean
            Dim box = UnionBox(members)
            Dim overlap = Math.Min(box.Right, candidate.BoundingBox.Right) - Math.Max(box.Left, candidate.BoundingBox.Left)
            Dim narrower = Math.Min(box.Right - box.Left, Width(candidate))
            If narrower <= 0 OrElse overlap / narrower < 0.45 Then Return False
            Dim gap = VerticalGap(box, candidate.BoundingBox)
            Return gap >= -4 AndAlso gap <= Math.Max(8, Height(candidate))
        End Function

        Private Shared Function SameRow(a As OcrElement, b As OcrElement) As Boolean
            Dim overlap = Math.Min(a.BoundingBox.Bottom, b.BoundingBox.Bottom) - Math.Max(a.BoundingBox.Top, b.BoundingBox.Top)
            Dim shorter = Math.Min(Height(a), Height(b))
            Return shorter > 0 AndAlso overlap / shorter >= 0.4
        End Function

        Private Shared Function SitsJustRight(members As List(Of OcrElement), candidate As OcrElement) As Boolean
            Dim box = UnionBox(members)
            Dim gap = candidate.BoundingBox.Left - box.Right
            Dim h = Math.Max(8, Math.Max(box.Bottom - box.Top, Height(candidate)))
            Return gap >= -4 AndAlso gap <= h * 0.85
        End Function

        Private Shared Function Near(a As OcrElement, b As OcrElement) As Boolean
            Dim gapX = Math.Max(0, Math.Max(a.BoundingBox.Left, b.BoundingBox.Left) - Math.Min(a.BoundingBox.Right, b.BoundingBox.Right))
            Dim gapY = Math.Max(0, Math.Max(a.BoundingBox.Top, b.BoundingBox.Top) - Math.Min(a.BoundingBox.Bottom, b.BoundingBox.Bottom))
            Dim h = Math.Max(Height(a), Height(b))
            Return gapX <= h * 1.2 AndAlso gapY <= h * 1.2
        End Function

        Private Shared Function CenterInside(el As OcrElement, box As OcrBoundingBox) As Boolean
            Dim x = CenterX(el)
            Dim y = CenterY(el)
            Return x >= box.Left AndAlso x <= box.Right AndAlso y >= box.Top AndAlso y <= box.Bottom
        End Function

        Private Shared Function CenterX(el As OcrElement) As Integer
            Return (el.BoundingBox.Left + el.BoundingBox.Right) \ 2
        End Function

        Private Shared Function CenterY(el As OcrElement) As Integer
            Return (el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2
        End Function

        Private Shared Function Width(el As OcrElement) As Integer
            Return Math.Max(1, el.BoundingBox.Right - el.BoundingBox.Left)
        End Function

        Private Shared Function Height(el As OcrElement) As Integer
            Return Math.Max(1, el.BoundingBox.Bottom - el.BoundingBox.Top)
        End Function

        Private Shared Function VerticalGap(a As OcrBoundingBox, b As OcrBoundingBox) As Integer
            Return Math.Max(a.Top, b.Top) - Math.Min(a.Bottom, b.Bottom)
        End Function

        Private Shared Function BandKey(el As OcrElement, nominal As OcrElement) As Integer
            Dim h = Math.Max(8, Height(nominal))
            Return CInt(Math.Round(CenterY(el) / CDbl(h)))
        End Function

        Private Shared Function UnionBox(members As List(Of OcrElement)) As OcrBoundingBox
            Return New OcrBoundingBox(
                members.Min(Function(el) el.BoundingBox.Left),
                members.Min(Function(el) el.BoundingBox.Top),
                members.Max(Function(el) el.BoundingBox.Right),
                members.Max(Function(el) el.BoundingBox.Bottom))
        End Function

        Private Shared Function CloneShell(anchor As OcrElement) As OcrElement
            Return New OcrElement() With {
                .Type = anchor.Type,
                .Text = anchor.Text,
                .Confidence = anchor.Confidence,
                .BoundingBox = anchor.BoundingBox,
                .Kind = anchor.Kind,
                .Nominal = anchor.Nominal,
                .Lower = anchor.Lower,
                .Upper = anchor.Upper,
                .Datum = anchor.Datum
            }
        End Function

        Private Shared Function JoinTexts(members As List(Of OcrElement)) As String
            Return String.Join(" ", members.Select(Function(el) If(el.Text, "").Trim()))
        End Function

        Private Shared Function Compact(text As String) As String
            Return Regex.Replace(If(text, ""), "\s+", "")
        End Function

        Private Shared Function Normalize(text As String) As String
            Dim value = If(text, "")
            value = value.Replace("＋", "+").Replace("－", "-").Replace("Φ", "Ø").Replace("φ", "Ø").Replace("⌀", "Ø")
            value = Regex.Replace(value, "\s+", " ").Trim()
            Return value
        End Function

    End Class

End Namespace
