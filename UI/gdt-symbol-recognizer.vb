Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

Namespace Services

    ''' <summary>
    ''' 识别完成后按框线取出符号格。只在「机械图纸」模式下调用。
    ''' 补垂直度、轮廓度、位置度，并给公差值补上直径符号。
    ''' </summary>
    Public Class GdtSymbolRecognizer

        ''' <summary>SplitDatumGroup 拆分多字母基准组时的中间结果</summary>
        Private Class DatumLetterPiece
            Public Property Chars As String
            Public Property CenterX As Integer
            Public Property CenterY As Integer
        End Class

        Private Shared ReadOnly HasPosition As New Regex("[⊕⌖]", RegexOptions.Compiled)
        Private Shared ReadOnly GluedFrame As New Regex("^(?<sym>[⊕⌖◎∥⊥▱⌒⌓∠↗])?\s*(?<dia>[OØøΦφ⌀∅])?\s*(?<num>\d+\.\d+|\d+)\s*(?<dat>[A-Z]{0,6})$", RegexOptions.Compiled)
        Private Shared ReadOnly ToleranceCore As New Regex("^(?:\d+\s*[xX×]\s+)?(?:[OØøΦφ⌀∅]\s*)?\d+\.\d+(?:\s+[A-Z]){0,6}$", RegexOptions.Compiled)
        Private Shared ReadOnly SlashFrame As New Regex("^(?://+|∥)(?<num>0+\.?\d+)(?<dat>[A-Z]{1,3})$", RegexOptions.Compiled)
        Private Shared ReadOnly LetterFrame As New Regex("^(?<junk>[A-Za-zØ⌀○])(?<num>\d+\.\d+)(?<dat>[A-Z]{1,3})$", RegexOptions.Compiled)
        Private Shared ReadOnly FrameValue As New Regex("^(?:\d+\s*[xX×]\s+)?(?:[OØøΦφ⌀∅]\s*)?\d+(?:\.\d+)?(?:\s+[A-Z]){1,6}$", RegexOptions.Compiled)

        ''' <summary>
        ''' 球标内部数字（1-99，图纸编号通常不大），或单个带圈数字（①..⑳、❶..❿、⓪..⑳、㉑..㉟、㊱..㊿）。
        ''' 带圈数字保持原文不改写：解析脚本的 readBalloon 认 ① 不认裸 1，改写成数字会丢球标序号。
        ''' </summary>
        Private Shared ReadOnly BalloonDigit As New Regex("^(?:\d{1,2}|[\u2460-\u2473\u2776-\u277F\u24EA-\u24FF\u3251-\u325F\u32B1-\u32BF])$", RegexOptions.Compiled)

        ''' <summary>单个带圈数字（五套）。球标，绝不是形位符号——RestoreRow 不得改写它。</summary>
        Private Shared ReadOnly CircledText As New Regex("^[\u2460-\u2473\u2776-\u277F\u24EA-\u24FF\u3251-\u325F\u32B1-\u32BF]$", RegexOptions.Compiled)
        ''' <summary>基准圆圈内部的单个大写字母（A-Z，排除掉常见组合）</summary>
        Private Shared ReadOnly DatumLetter As New Regex("^[A-Z]$", RegexOptions.Compiled)
        ''' <summary>2-4 个大写字母（多个基准圆圈的字母被 OCR 合并时的 fallback）</summary>
        Private Shared ReadOnly DatumLetterGroup As New Regex("^[A-Z]{2,4}$", RegexOptions.Compiled)
        ''' <summary>粗糙度只认已经写出参数名的文本。单独的 1、12.57 是尺寸或标题栏，不补成 Ra。</summary>
        Private Shared ReadOnly RoughnessValue As New Regex("^(?:Ra|Rz|Rq|Rt|Rh|Sm)\s*\d+(?:\.\d+)?$", RegexOptions.Compiled Or RegexOptions.IgnoreCase)

        ''' <summary>值后面直接跟基准字母时补空格：0.05A → 0.05 A（下游解析脚本靠空格分基准）</summary>
        Private Shared ReadOnly GluedDatum As New Regex("(\d)([A-Z])", RegexOptions.Compiled)

        ''' <summary>
        ''' 符号表单一事实源：Classify 能产出、MarkGdtAndDimension 要标记、直径判定要排除的都在这里。
        ''' 之前 MarkGdtAndDimension 漏了 ⌭（圆柱度）和 ○（圆度），Classify 认出来却标不上类型。
        ''' </summary>
        Private Shared ReadOnly GdtSymbolChars As New HashSet(Of Char)() From {"⊕"c, "⌖"c, "◎"c, "∥"c, "⊥"c, "▱"c, "⌒"c, "⌓"c, "∠"c, "↗"c, "⌭"c, "○"c}
        ''' <summary>位置度：公差值允许带 Ø</summary>
        Private Shared ReadOnly PositionSymbolChars As New HashSet(Of Char)() From {"⊕"c, "⌖"c}
        ''' <summary>公差值不带 Ø 的符号</summary>
        Private Shared ReadOnly NonDiameterSymbolChars As New HashSet(Of Char)() From {"⌓"c, "⌒"c, "⊥"c, "∥"c, "∠"c, "▱"c, "⌭"c, "↗"c, "○"c}

        Public Shared Sub Apply(doc As OcrDocument)
            If doc Is Nothing Then Return
            Dim profile = Services.DomainPack.CorrectionProfile.LoadMechanical()
            Using Services.DomainPack.CorrectionProfile.Enter(profile)
                ApplyWithProfile(doc, profile)
            End Using
        End Sub

        Private Shared Sub ApplyWithProfile(doc As OcrDocument, profile As Services.DomainPack.CorrectionProfile)
            Dim hasImage = Not String.IsNullOrEmpty(doc.SourceFilePath) AndAlso File.Exists(doc.SourceFilePath)
            Dim targets = If(doc.RawCellElements IsNot Nothing AndAlso doc.RawCellElements.Count > 0, doc.RawCellElements, doc.Elements)
            If Not hasImage OrElse targets Is Nothing OrElse targets.Count = 0 Then
                ' 没有图可做几何检测时，文字层面的引线规则修整仍然要跑：
                ' 公差小数、箭头碎片都不需要墨迹。
                If targets IsNot Nothing AndAlso profile.IsStageOn("leaderRules") Then
                    RepairBareDeviations(doc, targets)
                    DropLeaderJunk(doc, targets)
                End If
                profile.ApplyRules(doc)
                Return
            End If

            Dim bmp As Bitmap = Nothing
            Try
                Using fs As New FileStream(doc.SourceFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    Using tmp As New Bitmap(fs)
                        bmp = New Bitmap(tmp)
                    End Using
                End Using

                ' 阶段由 CorrectionProfile 开关。默认全开，与原来的固定流程一致。
                If profile.IsStageOn("gdtFrame") Then
                    Dim rows = GroupRows(targets)
                    For Each row In rows
                        RestoreRow(bmp, doc, row)
                    Next
                End If

                If profile.IsStageOn("balloon") Then DetectBalloons(bmp, doc, targets)
                ' 竖排（旋转 90°）标注：检测器常把它切成上下两半、又多检一个小碎片，
                ' 两个盒子重叠着显示；Ø 也会被读成 0。趁元素还没被分组/收并，先合成一条。
                If profile.IsStageOn("vertical") Then
                    MergeSplitVerticalText(doc, targets)
                    RestoreVerticalDiameter(bmp, targets)
                End If
                If profile.IsStageOn("datum") Then DetectDatumCircles(bmp, doc, targets)
                ' datumScan 是「OCR 漏掉基准字母」时才用的全图扫圆兜底，慢且有误报，默认关。
                If profile.IsStageOn("datumScan") Then ScanAllDatumCircles(bmp, doc)
                If profile.IsStageOn("roughness") Then DetectRoughnessSymbols(bmp, doc, targets)
                If profile.IsStageOn("mark") Then MarkGdtAndDimension(doc, targets)

                If profile.IsStageOn("fcf") Then
                    Dim tolCandidates = targets.Where(Function(el) el IsNot Nothing AndAlso el.BoundingBox IsNot Nothing).ToList()
                    DetectFcfFrames(bmp, doc, tolCandidates)
                End If

                ' 引线规则前置修整第二步：公差小数修回、箭头碎片清掉。
                ' +02 不修成 +0.2，IsDeviation 不认，偏差永远叠不上名义值；
                ' 业、北、® 是箭头/引线末端的误识，不是标注文字，留着会把同一条尺寸线上的
                ' 几段文字连成一块——识别盒子就是这样脱离引线规则的。
                If profile.IsStageOn("leaderRules") Then
                    RepairBareDeviations(doc, targets)
                    DropLeaderJunk(doc, targets)
                End If

                If profile.IsStageOn("leader") Then LeaderInk.Apply(bmp, targets)
                If doc.RawCellElements IsNot Nothing AndAlso doc.RawCellElements.Count > 0 AndAlso Not Object.ReferenceEquals(targets, doc.Elements) Then
                    Dim carried = EnsureElements(doc).Where(Function(el) el IsNot Nothing AndAlso Not targets.Contains(el) AndAlso
                        (el.Type = OcrElementType.DatumCircle OrElse
                         (el.Text IsNot Nothing AndAlso el.Text.StartsWith("FCF:")))).ToList()
                    Dim merged As New List(Of OcrElement)(targets)
                    merged.AddRange(carried)
                    doc.Elements = merged
                End If

                If profile.IsStageOn("groupDimensions") Then DrawingDimensionGrouper.Apply(doc)
                profile.ApplyRules(doc)
            Catch ex As Exception
                ' 只写 Debug.WriteLine 时，整页增强会静默丢掉、只剩 ApplyRules，现场完全查不到原因。
                WriteLog("Apply 失败，本页增强已跳过：" & ex.Message & vbCrLf & ex.StackTrace)
            Finally
                If bmp IsNot Nothing Then bmp.Dispose()
            End Try
        End Sub

#Region "辅助：日志与元素容器"

        ''' <summary>
        ''' 排障日志：%AppData%\ExcelAIVSTOAddIn\Logs\gdt_recognizer.log。
        ''' </summary>
        Private Shared Sub WriteLog(message As String)
            Try
                Dim dir = Path.Combine(Utils.PathHelper.AppDataRoot, "Logs")
                Directory.CreateDirectory(dir)
                File.AppendAllText(Path.Combine(dir, "gdt_recognizer.log"),
                                   DateTime.Now.ToString("HH:mm:ss.fff") & " " & message & vbCrLf, Encoding.UTF8)
            Catch
                ' 日志本身不能影响识别
            End Try
        End Sub

        ''' <summary>符号 → 中文形位名。模块对外只写中文名（解析脚本示例就是「位置度 0.05 A B」这种写法）。</summary>
        Private Shared ReadOnly GdtSymbolNames As New Dictionary(Of String, String)(StringComparer.Ordinal) From {
            {"⊕", "位置度"}, {"⌖", "同轴度"}, {"◎", "同轴度"}, {"∥", "平行度"}, {"⊥", "垂直度"},
            {"▱", "平面度"}, {"⌒", "线轮廓度"}, {"⌓", "面轮廓度"}, {"∠", "倾斜度"},
            {"↗", "圆跳动"}, {"⌭", "圆柱度"}, {"○", "圆度"}
        }

        ''' <summary>Classify / InferFcfSymbol 认不出的符号返回空串，调用方跳过，不要写"形位公差"这种解析脚本不认的名字。</summary>
        Private Shared Function SymbolName(symbol As String) As String
            If String.IsNullOrEmpty(symbol) Then Return ""
            Dim name As String = Nothing
            If GdtSymbolNames.TryGetValue(symbol.Trim(), name) Then Return name
            Return ""
        End Function

        ''' <summary>文本里已经写了中文形位名时，不要再补名前缀。</summary>
        Private Shared Function HasGdtName(text As String) As Boolean
            Dim s = If(text, "")
            For Each pair In GdtSymbolNames
                If s.Contains(pair.Value) Then Return True
            Next
            Return False
        End Function

        ''' <summary>
        ''' 改写 OCR 元素文本时，把另一份镜像元素（RawCellElements / Elements 是两份等价列表）一起改掉，
        ''' 否则输出里会同时存在"0.05"和"位置度 0.05"。
        ''' </summary>
        Private Shared Sub SyncTwinText(doc As OcrDocument, targets As List(Of OcrElement), source As OcrElement, newText As String)
            Dim other = If(Object.ReferenceEquals(targets, doc.Elements), doc.RawCellElements, doc.Elements)
            If other Is Nothing OrElse source.BoundingBox Is Nothing Then Return
            For Each twin In other
                If twin Is Nothing OrElse twin.BoundingBox Is Nothing Then Continue For
                If twin.BoundingBox.Left = source.BoundingBox.Left AndAlso
                   twin.BoundingBox.Top = source.BoundingBox.Top AndAlso
                   twin.BoundingBox.Right = source.BoundingBox.Right AndAlso
                   twin.BoundingBox.Bottom = source.BoundingBox.Bottom Then
                    twin.Text = newText
                End If
            Next
        End Sub

        ''' <summary>RawText 里这段文字只出现一次时才同步，出现多次就留给整表重建，避免改错行。</summary>
        Private Shared Sub ReplaceUniqueText(doc As OcrDocument, oldText As String, newText As String)
            If doc Is Nothing OrElse String.IsNullOrEmpty(oldText) OrElse oldText = newText OrElse String.IsNullOrEmpty(doc.RawText) Then Return
            Dim first = doc.RawText.IndexOf(oldText, StringComparison.Ordinal)
            If first < 0 Then Return
            If doc.RawText.LastIndexOf(oldText, StringComparison.Ordinal) <> first Then Return
            doc.RawText = doc.RawText.Substring(0, first) & newText & doc.RawText.Substring(first + oldText.Length)
        End Sub

        ''' <summary>保证 doc.Elements 可用。之前多处直接 doc.Elements.Add，Nothing 时抛 NRE 被上层 Catch 吞掉。</summary>
        Private Shared Function EnsureElements(doc As OcrDocument) As List(Of OcrElement)
            If doc Is Nothing Then Return Nothing
            If doc.Elements Is Nothing Then doc.Elements = New List(Of OcrElement)()
            Return doc.Elements
        End Function

        Private Shared Sub AddElement(doc As OcrDocument, el As OcrElement)
            If doc Is Nothing OrElse el Is Nothing Then Return
            EnsureElements(doc)
            doc.Elements.Add(el)
        End Sub

#End Region

#Region "阶段 1.5：引线规则前置修整"

        ''' <summary>
        ''' 竖排标注被检测器切成上下两半时合成一条（「Ø44.6」+「F8」→「Ø44.6 F8」）。
        ''' 两个框都又高又窄、左右基本对齐、上下挨着，就按阅读方向把文本接起来。
        ''' 顺手清掉被整个包在文字框里的小碎片（实测竖排 Ø78.5 框内多检出一个「3」，
        ''' 两个盒子重叠着显示，就是截图里「被拆成两个盒子还重叠」的来源）。
        ''' 横排文字不碰——它们靠引线/分组逻辑收，这里只管竖排。
        ''' </summary>
        Private Shared Sub MergeSplitVerticalText(doc As OcrDocument, targets As List(Of OcrElement))
            If targets Is Nothing OrElse targets.Count < 2 Then Return
            Dim drop As New HashSet(Of OcrElement)()

            ' 1) 包在别人框里的小碎片：直接丢掉（误检），不并进文本
            Dim kept = targets.Where(Function(el) el IsNot Nothing AndAlso el.BoundingBox IsNot Nothing AndAlso el.Type = OcrElementType.Paragraph).ToList()
            For Each a In kept
                If drop.Contains(a) Then Continue For
                For Each b In kept
                    If a Is b OrElse drop.Contains(b) Then Continue For
                    Dim areaB = Math.Max(1, (b.BoundingBox.Right - b.BoundingBox.Left) * (b.BoundingBox.Bottom - b.BoundingBox.Top))
                    Dim inter = IntersectArea(a.BoundingBox, b.BoundingBox)
                    If inter < areaB * 0.8 Then Continue For   ' b 几乎整个在 a 里
                    ' 被包住的是碎片：字少、置信度低、框也小。球标/基准圆此时已定型，不会走到这
                    Dim txtB = If(b.Text, "").Trim()
                    Dim areaA = Math.Max(1, (a.BoundingBox.Right - a.BoundingBox.Left) * (a.BoundingBox.Bottom - a.BoundingBox.Top))
                    ' 带圈数字是球标（圆环没验出来的也会留在 Paragraph），永远不当碎片丢——
                    ' 实测球标压在竖排文字框上时会被这条规则吃掉，气泡整条消失。
                    If CircledText.IsMatch(txtB) Then Continue For
                    If txtB.Length <= 2 AndAlso b.Confidence < a.Confidence AndAlso areaB < areaA * 0.5 Then
                        drop.Add(b)
                        RemoveRawTextLines(doc, txtB)
                    End If
                Next
            Next

            ' 2) 竖排被切成上下两半：合成一条
            Dim vertical = targets.Where(Function(el) el IsNot Nothing AndAlso el.BoundingBox IsNot Nothing AndAlso
                                                        el.Type = OcrElementType.Paragraph AndAlso Not drop.Contains(el) AndAlso
                                                        IsVerticalText(el.BoundingBox)).ToList()
            For i = 0 To vertical.Count - 1
                Dim top = vertical(i)
                If drop.Contains(top) Then Continue For
                For j = i + 1 To vertical.Count - 1
                    Dim bottom = vertical(j)
                    If drop.Contains(bottom) OrElse bottom Is top Then Continue For
                    If Not VerticalStackPairs(top, bottom) Then Continue For
                    ' 带圈数字是球标，不能并进尺寸文本里
                    If CircledText.IsMatch(If(top.Text, "").Trim()) OrElse CircledText.IsMatch(If(bottom.Text, "").Trim()) Then Continue For
                    Dim dir = If(top.VerticalDir <> "", top.VerticalDir, bottom.VerticalDir)
                    If dir = "" Then Continue For   ' 不知道阅读方向就不乱并
                    ' up=从下往上读：下面的盒子是文本开头
                    Dim merged As OcrElement = If(String.Equals(dir, "up", StringComparison.OrdinalIgnoreCase), bottom, top)
                    Dim later As OcrElement = If(merged Is top, bottom, top)
                    merged.Text = (If(merged.Text, "").Trim() & " " & If(later.Text, "").Trim()).Trim()
                    merged.BoundingBox = UnionBox(merged.BoundingBox, later.BoundingBox)
                    merged.Confidence = Math.Max(merged.Confidence, later.Confidence)
                    merged.Quad = Nothing   ' 盒子被合并改过，原四边形失效
                    drop.Add(later)
                    RemoveRawTextLines(doc, If(later.Text, "").Trim())
                    WriteLog($"[竖排合并] {merged.Text}")
                    Exit For   ' 一块并完就退出，剩下的块各自再找
                Next
            Next

            If drop.Count > 0 Then
                targets.RemoveAll(Function(el) drop.Contains(el))
                Dim twin = If(Object.ReferenceEquals(targets, doc.RawCellElements), doc.Elements, doc.RawCellElements)
                If twin IsNot Nothing Then twin.RemoveAll(Function(el) drop.Contains(el))
            End If
        End Sub

        Private Shared Function IsVerticalText(box As OcrBoundingBox) As Boolean
            Dim w = box.Right - box.Left
            Dim h = box.Bottom - box.Top
            Return h > w * 1.15
        End Function

        Private Shared Function IntersectArea(a As OcrBoundingBox, b As OcrBoundingBox) As Integer
            Dim w = Math.Min(a.Right, b.Right) - Math.Max(a.Left, b.Left)
            Dim h = Math.Min(a.Bottom, b.Bottom) - Math.Max(a.Top, b.Top)
            If w <= 0 OrElse h <= 0 Then Return 0
            Return w * h
        End Function

        Private Shared Function UnionBox(a As OcrBoundingBox, b As OcrBoundingBox) As OcrBoundingBox
            Return New OcrBoundingBox(Math.Min(a.Left, b.Left), Math.Min(a.Top, b.Top), Math.Max(a.Right, b.Right), Math.Max(a.Bottom, b.Bottom))
        End Function

        ''' <summary>上下两半能不能并成一条竖排标注：左右对齐、上下挨着（间隙不大）。</summary>
        Private Shared Function VerticalStackPairs(top As OcrElement, bottom As OcrElement) As Boolean
            If top Is Nothing OrElse bottom Is Nothing OrElse top.BoundingBox Is Nothing OrElse bottom.BoundingBox Is Nothing Then Return False
            If bottom.BoundingBox.Top < top.BoundingBox.Top Then Return False   ' top 必须真的在上
            Dim a = top.BoundingBox
            Dim b = bottom.BoundingBox
            Dim overlapX = Math.Min(a.Right, b.Right) - Math.Max(a.Left, b.Left)
            Dim narrow = Math.Min(a.Right - a.Left, b.Right - b.Left)
            If narrow <= 0 OrElse overlapX < narrow * 0.5 Then Return False
            Dim gap = b.Top - a.Bottom
            Dim tall = Math.Max(a.Bottom - a.Top, b.Bottom - b.Top)
            ' 允许略微重叠（检测框常互相压住），但不许隔着老远——那是另一条标注
            Return gap <= tall * 0.6 AndAlso gap >= -tall * 0.3
        End Function

        ''' <summary>
        ''' 竖排尺寸的直径符号修回来：引擎把 Ø 读成 0（竖排 Ø78.5 → 078.5），意思全变。
        ''' Ø 是文本第一个字符，竖排时就在文字起点那一侧：从下往上读在下方，从上往下读在上方。
        ''' 检测框常有偏移（不一定刚好包住第一个字符），所以从起点往里滑窗探几格，
        ''' 再往框外探一格——哪格像斜线圈，就认定 Ø 被读成了 0。
        ''' </summary>
        Private Shared Sub RestoreVerticalDiameter(bmp As Bitmap, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse targets Is Nothing Then Return
            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "").Trim()
                If txt = "" Then Continue For
                If txt.Contains("Ø") OrElse txt.Contains("ø") OrElse txt.Contains("⌀") Then Continue For
                If Not Regex.IsMatch(txt, "^0\d") Then Continue For   ' 只修 0 开头的（0 就是被读错的 Ø）
                Dim dir = If(el.VerticalDir, "")
                If dir = "" Then Continue For
                Dim box = el.BoundingBox
                Dim side = Math.Max(8, Math.Min(box.Right - box.Left, box.Bottom - box.Top))
                Dim fromBottom = String.Equals(dir, "up", StringComparison.OrdinalIgnoreCase)
                Dim found = False
                ' 起点那一侧：框内滑 3 格 + 框外 1格，覆盖检测框偏移/没包住首字符的情况
                For k = -1 To 3
                    Dim anchor As Integer
                    If fromBottom Then
                        anchor = box.Bottom - side * k
                    Else
                        anchor = box.Top + side * k
                    End If
                    Dim top As Integer
                    Dim bottom As Integer
                    If fromBottom Then
                        bottom = Math.Min(bmp.Height - 1, anchor + 2)
                        top = Math.Max(0, anchor - side + 2)
                    Else
                        top = Math.Max(0, anchor - 2)
                        bottom = Math.Min(bmp.Height - 1, anchor + side - 2)
                    End If
                    If bottom - top < 6 Then Continue For
                    Dim cell = ReadInk(bmp, box.Left, top, box.Right, bottom)
                    If LooksLikeDiameterSymbol(cell) Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then Continue For
                el.Text = "Ø" & txt.Substring(1)
                el.Quad = Nothing
                WriteLog($"[竖排直径] {txt} -> {el.Text}")
            Next
        End Sub

        ''' <summary>
        ''' 单个汉字里确实是标注词的白名单：图纸上的深/均布/沉孔，以及标题栏里单独成格的
        ''' 重量/材料/比例/设计…。不在这里面的汉字，机械图纸上单独出现基本都是箭头尖、
        ''' 引线末端的误识（业、北、甲、由、申、电……）。
        ''' </summary>
        Private Shared ReadOnly LeaderJunkKeepChars As New HashSet(Of Char)() From {
            "深"c, "均"c, "布"c, "沉"c, "锪"c, "孔"c, "宽"c, "厚"c, "长"c, "高"c, "度"c,
            "重"c, "量"c, "材"c, "料"c, "比"c, "例"c, "设"c, "计"c, "校"c, "对"c,
            "审"c, "核"c, "批"c, "准"c, "日"c, "期"c, "阶"c, "段"c, "共"c, "张"c,
            "第"c, "名"c, "称"c, "图"c, "号"c, "代"c, "密"c, "级"c, "数"c, "质"c
        }

        ''' <summary>只可能是误识的符号碎片。GDT 符号、Ø、√、± 都不在里面。</summary>
        Private Shared ReadOnly LeaderJunkSymbolChars As New HashSet(Of Char)() From {
            "®"c, "©"c, "™"c, "§"c, "※"c, "†"c, "‡"c, "•"c, "¶"c
        }

        ''' <summary>
        ''' 公差栈里丢小数点的裸偏差修回来：+02 → +0.2、±02 → ±0.2。
        ''' 不修的话 DrawingDimensionGrouper 的 IsDeviation 不认，上下偏差永远叠不上名义值，
        ''' 一条公差会被拆成三行（截图里的「22.4-0 +02」就是这么来的）。
        ''' 只修「符号 + 0 + 一位数字」这种三段式：-0 是真零偏差、+20 可能是真公差，都不碰。
        ''' </summary>
        Private Shared Sub RepairBareDeviations(doc As OcrDocument, targets As List(Of OcrElement))
            If doc Is Nothing OrElse targets Is Nothing Then Return
            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "")
                If txt = "" Then Continue For
                Dim fixed_ = LostDecimalDeviation.Replace(txt, "${1}0.$2")
                ' 尾随符号的裸偏差：0- → -0。水平公差栈里引擎常把下偏差 -0 读成 0-，
                ' IsDeviation 不认，公差栈就叠不上名义值（RapidOCR 实测）。
                Dim tail = Regex.Match(txt, "^(\d+(?:\.\d+)?)([+-])$")
                If tail.Success Then fixed_ = tail.Groups(2).Value & tail.Groups(1).Value
                If fixed_ <> txt Then
                    Dim old = txt.Trim()
                    el.Text = fixed_
                    SyncTwinText(doc, targets, el, fixed_)
                    ReplaceUniqueText(doc, old, fixed_.Trim())
                End If
            Next
        End Sub

        ''' <summary>符号 + 0 + 一位数字，且自成一段（开头或空白之后，空白或结尾之前）。</summary>
        Private Shared ReadOnly LostDecimalDeviation As New Regex("(?<=^|\s)([+±-])0(\d)(?=\s|$)", RegexOptions.Compiled)

        ''' <summary>
        ''' 清掉箭头、引线末端的误识字：单个汉字（业、北、甲、由）和纯符号碎片（®©™§※）。
        ''' 它们不是标注文字，留在列表里既污染输出，又会在按行合并时把几段文字连成一块——
        ''' 用户看到的「识别盒子脱离引线规则」主要就是这些碎片和错位合并造成的。
        ''' 同步清掉另一份镜像列表和 RawText 里对应的整行。
        ''' </summary>
        Private Shared Sub DropLeaderJunk(doc As OcrDocument, targets As List(Of OcrElement))
            If doc Is Nothing OrElse targets Is Nothing Then Return
            Dim junk As New List(Of OcrElement)()
            For Each el In targets
                If IsLeaderJunk(el) Then junk.Add(el)
            Next
            If junk.Count = 0 Then Return
            targets.RemoveAll(Function(el) junk.Contains(el))
            ' RawCellElements / Elements 是两份等价列表，只清一份会让输出里留着垃圾框
            Dim other = If(Object.ReferenceEquals(targets, doc.Elements), doc.RawCellElements, doc.Elements)
            If other IsNot Nothing Then other.RemoveAll(Function(el) junk.Contains(el))
            For Each el In junk
                RemoveRawTextLines(doc, If(el.Text, "").Trim())
            Next
            WriteLog("引线规则：清掉 " & junk.Count & " 个箭头/引线误识字：" &
                     String.Join("、", junk.Select(Function(el) If(el.Text, "").Trim())))
        End Sub

        Private Shared Function IsLeaderJunk(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return False
            If el.Type = OcrElementType.Balloon OrElse el.Type = OcrElementType.DatumCircle OrElse el.Type = OcrElementType.GdtSymbol Then Return False
            Dim txt = If(el.Text, "").Trim()
            If txt = "" OrElse txt.Length > 2 Then Return False
            For Each ch In txt
                If ch >= "一"c AndAlso ch <= "龥"c Then
                    If Not LeaderJunkKeepChars.Contains(ch) Then Return True
                ElseIf LeaderJunkSymbolChars.Contains(ch) Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>RawText 里整行等于这段文字时删掉该行（碎片在原文里也只可能是独立一行）。</summary>
        Private Shared Sub RemoveRawTextLines(doc As OcrDocument, text As String, Optional maxCount As Integer = Integer.MaxValue)
            If doc Is Nothing OrElse String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(doc.RawText) Then Return
            Dim lines = doc.RawText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf)
            Dim kept As New List(Of String)()
            Dim removed = 0
            For Each line In lines
                If removed < maxCount AndAlso String.Equals(line.Trim(), text, StringComparison.Ordinal) Then
                    removed += 1
                    Continue For
                End If
                kept.Add(line)
            Next
            If removed = 0 Then Return
            doc.RawText = String.Join(vbCrLf, kept)
        End Sub

#End Region

#Region "阶段 2：球标圆圈检测"

        ''' <summary>
        ''' 球标检测：OCR 已经识别出孤立的数字（1-99），在它周围验证是否有圆形轮廓。
        ''' 圆形轮廓用径向 ink 覆盖率判断：以文字框中心为圆心，在半径 R 处的圆环上应该有较多墨迹。
        ''' </summary>
        Private Shared Sub DetectBalloons(bmp As Bitmap, doc As OcrDocument, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse doc Is Nothing OrElse targets Is Nothing Then Return

            Dim existingTexts As New HashSet(Of String)()
            If doc.Elements IsNot Nothing Then
                For Each el In doc.Elements
                    If el Is Nothing Then Continue For
                    If el.Type = OcrElementType.Balloon OrElse el.Type = OcrElementType.DatumCircle Then
                        If el.BoundingBox IsNot Nothing Then existingTexts.Add(el.BoundingBox.Left & "," & el.BoundingBox.Top)
                    End If
                Next
            End If

            SplitGluedBalloons(bmp, doc, targets)

            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "").Trim()
                If Not BalloonDigit.IsMatch(txt) Then Continue For
                If existingTexts.Contains(el.BoundingBox.Left & "," & el.BoundingBox.Top) Then Continue For

                ' 文字框有时只包住数字，有时已经包住整圈。两种都要能套到圆。
                Dim glyph = Math.Max(8, Math.Max(el.BoundingBox.Right - el.BoundingBox.Left, el.BoundingBox.Bottom - el.BoundingBox.Top))
                Dim box = ExpandAround(bmp, el.BoundingBox, 5.0, 5.0)
                If box Is Nothing Then Continue For
                Dim ink = ReadInkAll(bmp, box.Left, box.Top, box.Right, box.Bottom)
                If ink Is Nothing Then Continue For

                Dim cx = (el.BoundingBox.Left + el.BoundingBox.Right) \ 2 - box.Left
                Dim cy = (el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2 - box.Top
                Dim H = ink.GetLength(1)
                Dim W = ink.GetLength(0)

                Dim bestRadius As Integer = 0
                Dim bestScore As Double = 0
                Dim minR = Math.Max(5, CInt(glyph * 0.35))
                Dim maxR = Math.Min(Math.Min(cx, cy), Math.Min(W - 1 - cx, H - 1 - cy))
                maxR = Math.Min(maxR, CInt(glyph * 2.8))
                If maxR < minR Then Continue For

                For r = minR To maxR
                    Dim score = RingInkScore(ink, cx, cy, r)
                    If score > bestScore Then
                        bestScore = score
                        bestRadius = r
                    End If
                Next

                If bestScore >= Services.DomainPack.CorrectionProfile.ResolveRingScore() AndAlso bestRadius >= minR Then
                    el.Type = OcrElementType.Balloon
                    el.Text = txt
                    el.BoundingBox = New OcrBoundingBox(
                        box.Left + cx - bestRadius,
                        box.Top + cy - bestRadius,
                        box.Left + cx + bestRadius,
                        box.Top + cy + bestRadius)
                    el.Confidence = bestScore
                End If
            Next
            StripBalloonSuffix(targets)
        End Sub

        ''' <summary>
        ''' 以 (cx,cy) 为圆心，r 为半径的圆环上的 ink 覆盖率
        ''' </summary>
        Private Shared Function RingInkScore(ink As Boolean(,), cx As Integer, cy As Integer, r As Integer, Optional thickness As Integer = 2) As Double
            If ink Is Nothing OrElse r < 2 Then Return 0
            Dim H = ink.GetLength(1)
            Dim W = ink.GetLength(0)
            Dim samples = 0
            Dim hits = 0
            For a = 0 To 359 Step 3
                Dim rad = a * Math.PI / 180.0
                For t = -thickness To thickness
                    Dim x = CInt(cx + (r + t) * Math.Cos(rad))
                    Dim y = CInt(cy + (r + t) * Math.Sin(rad))
                    If x < 0 OrElse x >= W OrElse y < 0 OrElse y >= H Then Continue For
                    samples += 1
                    If ink(x, y) Then hits += 1
                Next
            Next
            If samples = 0 Then Return 0
            Return hits / CDbl(samples)
        End Function

        ''' <summary>
        ''' 56 和球标 19 被收成一个框时，原文是 5619。右侧圆圈对得上才把球标剥出去。
        ''' </summary>
        Private Shared Sub SplitGluedBalloons(bmp As Bitmap, doc As OcrDocument, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse doc Is Nothing OrElse targets Is Nothing Then Return
            Dim added As New List(Of OcrElement)()
            For Each el In New List(Of OcrElement)(targets)
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                If el.Type = OcrElementType.Balloon OrElse el.Type = OcrElementType.DatumCircle Then Continue For
                Dim txt = Regex.Replace(If(el.Text, ""), "\s+", "")
                If Not Regex.IsMatch(txt, "^\d{3,6}$") Then Continue For
                Dim box = el.BoundingBox
                Dim width = box.Right - box.Left
                If width < 8 Then Continue For
                Dim charW = width / CDbl(txt.Length)
                Dim bestLen = 0
                Dim bestScore = 0.0
                For nlen = 1 To Math.Min(2, txt.Length - 1)
                    Dim tail = txt.Substring(txt.Length - nlen)
                    If tail = "0" OrElse tail = "00" Then Continue For
                    Dim split = box.Left + CInt(Math.Round((txt.Length - nlen) * charW))
                    Dim cy = (box.Top + box.Bottom) \ 2
                    Dim expectR = Math.Max(8, (box.Bottom - box.Top) \ 2)
                    Dim cxIn = (split + box.Right) \ 2
                    Dim score = Math.Max(BalloonAt(bmp, cxIn, cy, expectR), BalloonAt(bmp, box.Right + expectR \ 3, cy, expectR))
                    If score > bestScore Then
                        bestScore = score
                        bestLen = nlen
                    End If
                Next
                If bestLen = 0 OrElse bestScore < 0.42 Then Continue For

                Dim body = txt.Substring(0, txt.Length - bestLen)
                Dim tailText = txt.Substring(txt.Length - bestLen)
                Dim splitX = box.Left + CInt(Math.Round(body.Length * charW))
                Dim bodyIsBalloon = False
                If Regex.IsMatch(body, "^\d{1,2}$") Then
                    Dim bcx = (box.Left + splitX) \ 2
                    Dim bcy = (box.Top + box.Bottom) \ 2
                    bodyIsBalloon = BalloonAt(bmp, bcx, bcy, Math.Max(8, (box.Bottom - box.Top) \ 2)) >= 0.42
                End If

                Dim rightBalloon As New OcrElement() With {
                    .Type = OcrElementType.Balloon,
                    .Text = tailText,
                    .BoundingBox = New OcrBoundingBox(splitX, box.Top, box.Right, box.Bottom),
                    .Confidence = bestScore
                }
                added.Add(rightBalloon)
                If bodyIsBalloon Then
                    el.Type = OcrElementType.Balloon
                    el.Text = body
                    el.BoundingBox = New OcrBoundingBox(box.Left, box.Top, splitX, box.Bottom)
                Else
                    el.Text = body
                    el.BoundingBox = New OcrBoundingBox(box.Left, box.Top, Math.Max(box.Left + 1, splitX), box.Bottom)
                End If
            Next
            For Each balloon In added
                AddElement(doc, balloon)
                targets.Add(balloon)
            Next
        End Sub

        Private Shared Sub StripBalloonSuffix(targets As List(Of OcrElement))
            If targets Is Nothing Then Return
            For Each el In targets
                If el Is Nothing OrElse el.Type = OcrElementType.Balloon OrElse el.BoundingBox Is Nothing Then Continue For
                Dim compact = Regex.Replace(If(el.Text, ""), "\s+", "")
                If Not Regex.IsMatch(compact, "^\d{3,6}$") Then Continue For
                For Each balloon In targets
                    If balloon Is Nothing OrElse balloon.Type <> OcrElementType.Balloon OrElse balloon.BoundingBox Is Nothing Then Continue For
                    Dim digits = Regex.Match(If(balloon.Text, ""), "\d{1,2}$").Value
                    If digits = "" OrElse compact.Length <= digits.Length OrElse Not compact.EndsWith(digits) Then Continue For
                    If Not BalloonTouches(el.BoundingBox, balloon.BoundingBox) Then Continue For
                    el.Text = compact.Substring(0, compact.Length - digits.Length)
                    Exit For
                Next
            Next
        End Sub

        Private Shared Function BalloonTouches(text As OcrBoundingBox, balloon As OcrBoundingBox) As Boolean
            Dim h = Math.Max(8, text.Bottom - text.Top)
            Dim cx = (balloon.Left + balloon.Right) \ 2
            Dim cy = (balloon.Top + balloon.Bottom) \ 2
            Return cx >= text.Left - h AndAlso cx <= text.Right + h AndAlso cy >= text.Top - h AndAlso cy <= text.Bottom + h
        End Function

        Private Shared Function BalloonAt(bmp As Bitmap, cx As Integer, cy As Integer, expectR As Integer) As Double
            If bmp Is Nothing Then Return 0
            Dim pad = Math.Max(expectR * 3, 28)
            Dim left = Math.Max(0, cx - pad)
            Dim top = Math.Max(0, cy - pad)
            Dim right = Math.Min(bmp.Width - 1, cx + pad)
            Dim bottom = Math.Min(bmp.Height - 1, cy + pad)
            Dim ink = ReadInkAll(bmp, left, top, right, bottom)
            If ink Is Nothing Then Return 0
            Dim lx = cx - left
            Dim ly = cy - top
            Dim minR = Math.Max(6, expectR \ 2)
            Dim maxR = Math.Min(Math.Min(lx, ly), Math.Min(ink.GetLength(0) - 1 - lx, ink.GetLength(1) - 1 - ly))
            If maxR < minR Then Return 0
            Dim best = 0.0
            For r = minR To maxR
                Dim score = RingInkScore(ink, lx, ly, r)
                If score > best Then best = score
            Next
            Return best
        End Function

#End Region

#Region "阶段 3：基准圆圈检测"

        ''' <summary>
        ''' 基准圆圈：
        ''' (a) OCR 识别出孤立的单个大写字母 A/B/C/D/E/F...（排除 XYZ）
        ''' (b) OCR 把多个基准字母粘成一组（如 AB / ABC）时，用垂直 ink 投影拆回独立字母
        '''     然后按每个字母各自的位置验证圆 + 内部空密度。
        ''' </summary>
        Private Shared Sub DetectDatumCircles(bmp As Bitmap, doc As OcrDocument, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse doc Is Nothing OrElse targets Is Nothing Then Return

            Dim consumed As New List(Of OcrElement)()
            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "").Trim()

                ' 候选判定：单字母 或 2-4 字母组（排除 XYZ 坐标轴）
                Dim isSingle = DatumLetter.IsMatch(txt) AndAlso Not "XYZ".Contains(txt)
                Dim isGroup = DatumLetterGroup.IsMatch(txt) AndAlso Not txt.Any(Function(c) "XYZ".Contains(c))
                If Not (isSingle OrElse isGroup) Then Continue For

                ' 扩展搜索框（组比单字母更宽，所以 pad 也要放大）
                Dim groupPad As Double = If(isGroup, 3.5, 2.8)
                Dim box = ExpandAround(bmp, el.BoundingBox, groupPad, groupPad)
                If box Is Nothing Then Continue For
                Dim ink = ReadInkAll(bmp, box.Left, box.Top, box.Right, box.Bottom)
                If ink Is Nothing Then Continue For

                ' 全局中心（用于先粗找圆）
                Dim gcx = (el.BoundingBox.Left + el.BoundingBox.Right) \ 2 - box.Left
                Dim gcy = (el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2 - box.Top

                If isSingle Then
                    ' --- 单字母：原有逻辑 ---
                    If TryAddDatumCircle(ink, box, gcx, gcy, txt, el, doc) Then consumed.Add(el)
                Else
                    ' --- 多字母组：用垂直投影拆分 ---
                    Dim letters = SplitDatumGroup(ink, el, box, txt)
                    For Each letter In letters
                        If TryAddDatumCircle(ink, box, letter.CenterX, letter.CenterY, letter.Chars, el, doc) Then
                            consumed.Add(el)
                            Exit For
                        End If
                    Next
                End If
            Next

            ' 字母框已经被基准圆收编，原框要删掉：否则输出里同时留着「ⓍA」和一个游离的「A」，
            ' 用户看到的就是识别结果里多出一行没有引线规则的字母。
            For Each el In consumed
                targets.Remove(el)
                Dim twin = If(Object.ReferenceEquals(targets, doc.Elements), doc.RawCellElements, doc.Elements)
                If twin IsNot Nothing Then twin.Remove(el)
                ' 只删一行：同名基准字母（另一个 A）在原文里可能是另一条标注
                RemoveRawTextLines(doc, If(el.Text, "").Trim(), 1)
            Next
        End Sub

        ''' <summary>
        ''' 用垂直 ink 投影把 "AB"/"ABC" 拆成独立字母，返回每个字母的中心位置和字符
        ''' </summary>
        Private Shared Function SplitDatumGroup(ink As Boolean(,), origEl As OcrElement, searchBox As OcrBoundingBox, groupText As String) As List(Of DatumLetterPiece)
            Dim result As New List(Of DatumLetterPiece)()
            If ink Is Nothing OrElse String.IsNullOrEmpty(groupText) Then Return result

            Dim W = ink.GetLength(0)
            Dim H = ink.GetLength(1)
            ' 字母相对搜索框的位置
            Dim lx0 = Math.Max(0, origEl.BoundingBox.Left - searchBox.Left)
            Dim lx1 = Math.Min(W - 1, origEl.BoundingBox.Right - searchBox.Left)
            Dim ly0 = Math.Max(0, origEl.BoundingBox.Top - searchBox.Top)
            Dim ly1 = Math.Min(H - 1, origEl.BoundingBox.Bottom - searchBox.Top)
            If lx1 - lx0 < 4 OrElse ly1 - ly0 < 4 Then Return result

            ' 垂直投影：每列的 ink 像素数
            Dim proj() As Integer = New Integer(W - 1) {}
            For x = lx0 To lx1
                Dim cnt = 0
                For y = ly0 To ly1
                    If ink(x, y) Then cnt += 1
                Next
                proj(x) = cnt
            Next

            ' 用阈值切分：找连续的 ink-run，空区即拆分点
            Dim threshold = Math.Max(2, CInt((ly1 - ly0 + 1) * 0.08))
            Dim runs As New List(Of (start As Integer, end_ As Integer))()
            Dim inRun = False
            Dim runStart = 0
            For x = lx0 To lx1
                If proj(x) >= threshold AndAlso Not inRun Then
                    inRun = True
                    runStart = x
                ElseIf proj(x) < threshold AndAlso inRun Then
                    inRun = False
                    runs.Add((runStart, x - 1))
                End If
            Next
            If inRun Then runs.Add((runStart, lx1))

            ' run 数量 ≥ 字符数才算拆成功，否则当作单字母基准处理
            If runs.Count < groupText.Length Then
                result.Add(New DatumLetterPiece With {
                    .Chars = groupText,
                    .CenterX = (lx0 + lx1) \ 2,
                    .CenterY = (ly0 + ly1) \ 2
                })
                Return result
            End If

            ' 取前 N 个 run（N = 字符数），对齐
            For i = 0 To Math.Min(runs.Count, groupText.Length) - 1
                Dim r = runs(i)
                Dim cx = (r.start + r.end_) \ 2
                Dim cy = (ly0 + ly1) \ 2
                result.Add(New DatumLetterPiece With {
                    .Chars = groupText(i).ToString(),
                    .CenterX = cx,
                    .CenterY = cy
                })
            Next
            Return result
        End Function

        Private Shared Function TryAddDatumCircle(ink As Boolean(,), searchBox As OcrBoundingBox,
                                               localCx As Integer, localCy As Integer,
                                               letterText As String, origEl As OcrElement,
                                               doc As OcrDocument) As Boolean
            If ink Is Nothing OrElse doc Is Nothing Then Return False
            Dim H = ink.GetLength(1)
            Dim W = ink.GetLength(0)

            ' 以传入中心搜索最优圆
            Dim bestRadius As Integer = 0
            Dim bestScore As Double = 0
            Dim minR = 8
            Dim maxR = Math.Min(Math.Min(localCx, localCy), Math.Min(W - 1 - localCx, H - 1 - localCy))
            If maxR < minR Then Return False

            For r = minR To maxR
                Dim score = RingInkScore(ink, localCx, localCy, r)
                If score > bestScore Then
                    bestScore = score
                    bestRadius = r
                End If
            Next

            If bestScore < 0.35 Then Return False

            ' 圆内空密度检查（排除原文字母框区域）
            Dim innerEmpty = CheckCircleInnerEmpty(ink, localCx, localCy, bestRadius, origEl.BoundingBox, searchBox)
            If Not innerEmpty Then Return False

            ' 圆框坐标回原图纸坐标
            Dim circleBox = New OcrBoundingBox(
                searchBox.Left + localCx - bestRadius,
                searchBox.Top + localCy - bestRadius,
                searchBox.Left + localCx + bestRadius,
                searchBox.Top + localCy + bestRadius)

            Dim datum As New OcrElement() With {
                .Type = OcrElementType.DatumCircle,
                .Text = "Ⓧ" & letterText,
                .BoundingBox = circleBox,
                .Confidence = bestScore
            }
            AddElement(doc, datum)
            Return True
        End Function

        Private Shared Function CheckCircleInnerEmpty(ink As Boolean(,), cx As Integer, cy As Integer, r As Integer,
                                                     letterBox As OcrBoundingBox, searchBox As OcrBoundingBox) As Boolean
            If ink Is Nothing Then Return False
            Dim H = ink.GetLength(1)
            Dim W = ink.GetLength(0)
            Dim samples = 0
            Dim hits = 0
            Dim lx0 = Math.Max(0, letterBox.Left - searchBox.Left)
            Dim lx1 = Math.Min(W - 1, letterBox.Right - searchBox.Left)
            Dim ly0 = Math.Max(0, letterBox.Top - searchBox.Top)
            Dim ly1 = Math.Min(H - 1, letterBox.Bottom - searchBox.Top)
            For y = cy - r + 2 To cy + r - 2
                If y < 0 OrElse y >= H Then Continue For
                For x = cx - r + 2 To cx + r - 2
                    If x < 0 OrElse x >= W Then Continue For
                    Dim dx = x - cx, dy = y - cy
                    If dx * dx + dy * dy > (r - 2) * (r - 2) Then Continue For
                    ' 跳过字母本身区域
                    If x >= lx0 AndAlso x <= lx1 AndAlso y >= ly0 AndAlso y <= ly1 Then Continue For
                    samples += 1
                    If ink(x, y) Then hits += 1
                Next
            Next
            If samples < 4 Then Return True
            Dim density = hits / CDbl(samples)
            Return density < 0.20
        End Function

        ''' <summary>
        ''' 全图圆扫描基准圆圈 fallback：不需要 OCR 锚点，直接在整图上用 Hough-like 思路找圆，
        ''' 再检查圆内空密度 + 中心是否有字母墨迹。补足 OCR 漏识别的基准圆圈。
        ''' </summary>
        Private Shared Sub ScanAllDatumCircles(bmp As Bitmap, doc As OcrDocument)
            If bmp Is Nothing OrElse doc Is Nothing Then Return
            Dim scanned = 0
            ' Use ReadInkAll to include blue fills
            Dim ink = ReadInkAll(bmp, 0, 0, bmp.Width - 1, bmp.Height - 1)
            Dim H = ink.GetLength(1)
            Dim W = ink.GetLength(0)

            ' 降采样：原写法是 36 个半径 × 每 4 像素一个中心 × 每圆 120 次采样，大图上要跑几十秒。
            ' 中心步长按图幅放大（4~12），半径步长 2，再用候选上限兜底。
            Dim centerStep = Math.Max(4, Math.Min(12, CInt(Math.Sqrt(CDbl(W) * H) / 220)))
            Dim maxScanR = Math.Min(40, Math.Min(W, H) \ 2 - 1)
            Dim candidates As New List(Of (cx As Integer, cy As Integer, r As Integer, score As Double))
            Const MaxCandidates As Integer = 4000

            For r = 5 To maxScanR Step 2
                For y = r To H - 1 - r Step centerStep
                    For x = r To W - 1 - r Step centerStep
                        Dim s = RingInkScore(ink, x, y, r)
                        If s >= 0.40 Then  ' lower threshold for full-image scan, we filter later
                            candidates.Add((x, y, r, s))
                        End If
                    Next
                Next
                If candidates.Count >= MaxCandidates Then Exit For
            Next
            If candidates.Count >= MaxCandidates Then
                WriteLog("datumScan 候选达到上限 " & MaxCandidates & "，结果可能不全（图 " & W & "x" & H & "）")
            End If

            ' Non-max suppression: keep circles that don't heavily overlap
            Dim sorted = candidates.OrderByDescending(Function(c) c.score).ToList()
            Dim kept As New List(Of (cx As Integer, cy As Integer, r As Integer, score As Double))
            For Each c In sorted
                Dim overlap = False
                For Each k In kept
                    Dim dx = c.cx - k.cx, dy = c.cy - k.cy
                    If Math.Sqrt(dx * dx + dy * dy) < c.r + k.r - 2 Then
                        overlap = True
                        Exit For
                    End If
                Next
                If Not overlap Then kept.Add(c)
            Next

            ' For each kept circle, check if interior (minus exact center area) has low ink AND center has letter-like ink
            For Each c In kept
                ' Interior density check
                Dim samples As Integer = 0, hits As Integer = 0
                For y = Math.Max(0, c.cy - c.r + 2) To Math.Min(H - 1, c.cy + c.r - 2)
                    For x = Math.Max(0, c.cx - c.r + 2) To Math.Min(W - 1, c.cx + c.r - 2)
                        Dim dx = x - c.cx, dy = y - c.cy
                        If dx * dx + dy * dy > (c.r - 2) * (c.r - 2) Then Continue For
                        samples += 1
                        If ink(x, y) Then hits += 1
                    Next
                Next
                Dim emptyRatio = If(samples > 0, hits / CDbl(samples), 1.0)
                If emptyRatio > 0.25 Then Continue For  ' too filled, not a datum circle

                ' Center 40% area should have SOME ink (the letter)
                Dim centerSamples As Integer = 0, centerHits As Integer = 0
                Dim cr = CInt(c.r * 0.4)
                For y = c.cy - cr To c.cy + cr
                    For x = c.cx - cr To c.cx + cr
                        If x < 0 Or x >= W Or y < 0 Or y >= H Then Continue For
                        Dim dx = x - c.cx, dy = y - c.cy
                        If dx * dx + dy * dy > cr * cr Then Continue For
                        centerSamples += 1
                        If ink(x, y) Then centerHits += 1
                    Next
                Next
                Dim centerRatio = If(centerSamples > 0, centerHits / CDbl(centerSamples), 0.0)
                If centerRatio < 0.05 OrElse centerRatio > 0.5 Then Continue For  ' no letter or too much noise

                ' Avoid duplicates with existing datum circles
                Dim alreadyFound = EnsureElements(doc).Any(Function(el) el IsNot Nothing AndAlso el.Type = OcrElementType.DatumCircle AndAlso
                    el.BoundingBox IsNot Nothing AndAlso
                    Math.Abs((el.BoundingBox.Left + el.BoundingBox.Right) \ 2 - c.cx) < c.r AndAlso
                    Math.Abs((el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2 - c.cy) < c.r)
                If alreadyFound Then Continue For

                ' Circle found! Letter unknown without OCR, use generic label
                Dim datumBox As New OcrBoundingBox(c.cx - c.r, c.cy - c.r, c.cx + c.r, c.cy + c.r)
                Dim datum As New OcrElement() With {
                    .Type = OcrElementType.DatumCircle,
                    .Text = "Ⓧ?",  ' letter unknown, marked for user verification
                    .BoundingBox = datumBox,
                    .Confidence = c.score
                }
                AddElement(doc, datum)
                scanned += 1
            Next
            If scanned > 0 Then WriteLog("datumScan 补到 " & scanned & " 个基准圆")
        End Sub

#End Region

#Region "阶段 4：粗糙度 √ 三角符号检测"

        ''' <summary>
        ''' 粗糙度符号由一个 √ 三角（两斜边 + 一条水平底边）加上上方的 Ra/Rz 数值组成。
        ''' OCR 通常能识别出 Ra/Rz/数值，用它作为锚点，在其上方寻找三角轮廓。
        ''' </summary>
        Private Shared Sub DetectRoughnessSymbols(bmp As Bitmap, doc As OcrDocument, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse doc Is Nothing OrElse targets Is Nothing Then Return

            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "").Trim()
                If Not RoughnessValue.IsMatch(txt) Then Continue For

                ' 在数值框上方扩展搜索，找三角轮廓
                Dim box = ExpandAbove(bmp, el.BoundingBox, 2.5)
                If box Is Nothing Then Continue For
                Dim ink = ReadInk(bmp, box.Left, box.Top, box.Right, box.Bottom)
                If ink Is Nothing Then Continue For

                Dim H = ink.GetLength(1)
                Dim W = ink.GetLength(0)
                If W < 12 OrElse H < 12 Then Continue For

                ' 粗糙度三角特征：
                ' 1) 底边（底部水平线）有连续的 ink
                ' 2) 左斜边（从底边左端斜向右上）有 ink
                ' 3) 右斜边（从底边右端斜向左上）有 ink
                ' 或者简化版：底边 + 两条斜线，构成 √ 形

                Dim bottom = FindHorizonalBottom(ink, W, H)
                If bottom < 0 Then Continue For

                Dim leftEnd = FindDiagonalLeft(ink, bottom, W, H)
                Dim rightEnd = FindDiagonalRight(ink, bottom, W, H)
                If leftEnd < 0 OrElse rightEnd < 0 Then Continue For

                Dim triBox = New OcrBoundingBox(
                    box.Left + leftEnd,
                    box.Top,
                    box.Left + rightEnd,
                    box.Top + bottom + 2)

                el.Type = OcrElementType.Roughness
                el.Text = "√ " & txt
                el.BoundingBox = triBox
                el.Confidence = 0.85
            Next
        End Sub

        ''' <summary>找底部水平线（底边）的 y 坐标</summary>
        Private Shared Function FindHorizonalBottom(ink As Boolean(,), W As Integer, H As Integer) As Integer
            For y = H - 1 To CInt(H * 0.4) Step -1
                Dim hits = 0
                Dim spans = 0
                Dim inRun = False
                For x = 0 To W - 1
                    If ink(x, y) Then
                        hits += 1
                        If Not inRun Then inRun = True
                    ElseIf inRun Then
                        spans += 1
                        inRun = False
                    End If
                Next
                If inRun Then spans += 1
                If hits >= W * 0.35 AndAlso spans <= 2 Then Return y
            Next
            Return -1
        End Function

        ''' <summary>从底边上的点向左上方找斜线终点 x（底边左端的 x）</summary>
        Private Shared Function FindDiagonalLeft(ink As Boolean(,), bottomY As Integer, W As Integer, H As Integer) As Integer
            ' 底边的起始 x（从左往右第一个 ink 像素）
            Dim startX = -1
            For x = 0 To W - 1
                If ink(x, bottomY) Then startX = x : Exit For
            Next
            If startX < 0 Then Return -1

            ' 从底边左端点，向上向左扫描（√ 的左半边：从底边左 → 顶点方向）
            Dim px = startX
            For dy = 1 To bottomY
                Dim cx = px - 1
                If cx < 0 Then Exit For
                Dim found = False
                For dx = -3 To 3
                    Dim tx = px + dx
                    Dim ty = bottomY - dy
                    If tx >= 0 AndAlso tx < W AndAlso ty >= 0 AndAlso ty < H AndAlso ink(tx, ty) Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then
                    ' 三角形斜线可以有间隙，检查附近 y 是否有 ink
                    For dx = -4 To 4
                        Dim tx = px + dx
                        If tx >= 0 AndAlso tx < W AndAlso bottomY - dy - 1 >= 0 AndAlso ink(tx, bottomY - dy - 1) Then
                            found = True
                            px = tx
                            Exit For
                        End If
                    Next
                End If
                If found Then px = cx
            Next
            Return Math.Max(0, px - 1)
        End Function

        ''' <summary>从底边上的点向右上方找斜线终点 x（底边右端的 x）</summary>
        Private Shared Function FindDiagonalRight(ink As Boolean(,), bottomY As Integer, W As Integer, H As Integer) As Integer
            Dim endX = -1
            For x = W - 1 To 0 Step -1
                If ink(x, bottomY) Then endX = x : Exit For
            Next
            If endX < 0 Then Return -1

            Dim px = endX
            For dy = 1 To bottomY
                Dim cx = px + 1
                If cx >= W Then Exit For
                Dim found = False
                For dx = -3 To 3
                    Dim tx = px + dx
                    Dim ty = bottomY - dy
                    If tx >= 0 AndAlso tx < W AndAlso ty >= 0 AndAlso ty < H AndAlso ink(tx, ty) Then
                        found = True
                        Exit For
                    End If
                Next
                If Not found Then
                    For dx = -4 To 4
                        Dim tx = px + dx
                        If tx >= 0 AndAlso tx < W AndAlso bottomY - dy - 1 >= 0 AndAlso ink(tx, bottomY - dy - 1) Then
                            found = True
                            px = tx
                            Exit For
                        End If
                    Next
                End If
                If found Then px = cx
            Next
            Return Math.Min(W - 1, px + 1)
        End Function

#End Region

#Region "辅助：扩展文字框搜索范围"

        Private Shared Function ExpandAround(bmp As Bitmap, box As OcrBoundingBox, hMult As Double, vMult As Double) As OcrBoundingBox
            If bmp Is Nothing OrElse box Is Nothing Then Return box
            Dim w = box.Right - box.Left
            Dim h = box.Bottom - box.Top
            Dim cx = (box.Left + box.Right) \ 2
            Dim cy = (box.Top + box.Bottom) \ 2
            Dim hw = CInt(w * hMult / 2)
            Dim hh = CInt(h * vMult / 2)
            Dim l = Math.Max(0, cx - hw)
            Dim t = Math.Max(0, cy - hh)
            Dim r = Math.Min(bmp.Width - 1, cx + hw)
            Dim b = Math.Min(bmp.Height - 1, cy + hh)
            If r - l < 8 OrElse b - t < 8 Then Return Nothing
            Return New OcrBoundingBox(l, t, r, b)
        End Function

        Private Shared Function ExpandAbove(bmp As Bitmap, box As OcrBoundingBox, topMult As Double) As OcrBoundingBox
            If bmp Is Nothing OrElse box Is Nothing Then Return box
            Dim h = box.Bottom - box.Top
            Dim padTop = CInt(h * topMult)
            Dim l = Math.Max(0, box.Left - CInt(h * 0.5))
            Dim t = Math.Max(0, box.Top - padTop)
            Dim r = Math.Min(bmp.Width - 1, box.Right + CInt(h * 0.5))
            Dim b = Math.Min(bmp.Height - 1, box.Bottom + CInt(h * 0.3))
            If r - l < 8 OrElse b - t < 8 Then Return Nothing
            Return New OcrBoundingBox(l, t, r, b)
        End Function

#End Region

#Region "阶段 5：标记 GDT 符号 + 尺寸 Element 类型"

        Private Shared Sub MarkGdtAndDimension(doc As OcrDocument, targets As List(Of OcrElement))
            If doc Is Nothing OrElse doc.Elements Is Nothing Then Return

            Dim symChars = GdtSymbolChars
            Dim diaChars As New HashSet(Of Char)() From {"Ø"c, "ø"c, "Φ"c, "φ"c, "⌀"c, "∅"c}

            For Each el In doc.Elements
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                If el.Type <> OcrElementType.Paragraph AndAlso el.Type <> OcrElementType.Image Then Continue For
                Dim txt = If(el.Text, "")
                If String.IsNullOrEmpty(txt) Then Continue For

                ' GDT 符号
                If txt.Length = 1 AndAlso symChars.Contains(txt(0)) Then
                    el.Type = OcrElementType.GdtSymbol
                    Continue For
                End If

                ' 尺寸标注
                If BalloonDigit.IsMatch(txt) Then Continue For ' 球标数字已经单独处理
                Dim hasDim = diaChars.Overlaps(txt) OrElse Regex.IsMatch(txt, "\d+\s*[×xX]\s*\d") OrElse Regex.IsMatch(txt, "±")
                If hasDim AndAlso Regex.IsMatch(txt, "\d") Then
                    el.Type = OcrElementType.Dimension
                    Continue For
                End If

                ' 粗糙度数值
                If RoughnessValue.IsMatch(txt) Then
                    el.Type = OcrElementType.Roughness
                    Continue For
                End If
            Next
        End Sub

#End Region

        Private Shared Sub RestoreRow(bmp As Bitmap, doc As OcrDocument, row As List(Of OcrElement))
            Dim ordered = row.OrderBy(Function(el) BoxLeft(el)).ToList()
            Dim oldJoined = JoinRow(ordered)
            If oldJoined = "" Then Return

            Dim rebuiltAny = False
            For Each el In ordered
                Dim own = Normalize(el.Text)
                If own = "" OrElse el.BoundingBox Is Nothing OrElse Not LooksLikeFrame(own) Then Continue For
                Dim rebuilt = RebuildFrame(bmp, el)
                If rebuilt = "" OrElse Normalize(rebuilt) = own Then Continue For
                el.Text = rebuilt
                ReplaceText(doc, own, rebuilt)
                rebuiltAny = True
            Next
            If rebuiltAny Then Return
            If Not LooksLikeFrame(oldJoined) Then Return

            Dim valueCell = ordered.FirstOrDefault(Function(el) Regex.IsMatch(If(el.Text, ""), "\d"))
            If valueCell Is Nothing Then valueCell = ordered(0)
            Dim changed = False

            If Not HasPosition.IsMatch(oldJoined) Then
                Dim symbol = ""
                For Each el In ordered
                    If el Is valueCell Then Exit For
                    ' 带圈数字是球标不是形位符号。行里只有「球标 + 一个数值」时
                    ' （② 11.3），旧逻辑会把球标改写成「位置度」，球标整条消失。
                    If IsReplaceableSymbol(el.Text) AndAlso Not CircledText.IsMatch(Normalize(el.Text)) AndAlso el.BoundingBox IsNot Nothing Then
                        symbol = ClassifyBox(bmp, el.BoundingBox, False)
                        Dim symName = SymbolName(symbol)
                        If symName <> "" Then
                            el.Text = symName
                            changed = True
                            Exit For
                        End If
                    End If
                Next
                If symbol = "" AndAlso valueCell.BoundingBox IsNot Nothing Then
                    symbol = ClassifyBox(bmp, valueCell.BoundingBox, True)
                    Dim symName2 = SymbolName(symbol)
                    If symName2 <> "" Then
                        If IsReplaceableSymbol(valueCell.Text) Then
                            valueCell.Text = symName2
                        Else
                            valueCell.Text = ApplySymbol(symbol, valueCell.Text)
                        End If
                        changed = True
                    End If
                End If
            End If

            If valueCell.BoundingBox IsNot Nothing AndAlso Not Regex.IsMatch(If(valueCell.Text, ""), "[Øø]") Then
                If HasDiameterGlyph(bmp, valueCell.BoundingBox) Then
                    valueCell.Text = WithDiameter(valueCell.Text)
                    changed = True
                End If
            End If

            If Not changed Then Return
            ReplaceText(doc, oldJoined, JoinRow(ordered))
        End Sub

        Private Shared Sub ReplaceText(doc As OcrDocument, oldText As String, newText As String)
            If String.IsNullOrEmpty(oldText) OrElse oldText = newText Then Return
            If doc.Elements IsNot Nothing Then
                For Each el In doc.Elements
                    If el Is Nothing OrElse String.IsNullOrEmpty(el.Text) Then Continue For
                    Dim flat = Normalize(el.Text)
                    If flat = oldText Then el.Text = newText
                Next
            End If
            If Not String.IsNullOrEmpty(doc.RawText) AndAlso doc.RawText.Contains(oldText) Then
                doc.RawText = ReplaceFirst(doc.RawText, oldText, newText)
            End If
        End Sub

        Private Shared Function ReplaceFirst(text As String, oldText As String, newText As String) As String
            Dim at = text.IndexOf(oldText, StringComparison.Ordinal)
            If at < 0 Then Return text
            Return text.Substring(0, at) & newText & text.Substring(at + oldText.Length)
        End Function

        Private Shared Function LooksLikeFrame(text As String) As Boolean
            Dim compact = CoreCompact(text)
            If SlashFrame.IsMatch(compact) OrElse LetterFrame.IsMatch(compact) Then Return True
            ' 单独的 11.3、56 是线性尺寸。没有基准字母或直径符号时，不把它当成公差框。
            If Regex.IsMatch(compact, "^\d+(?:\.\d+)?$") Then Return False
            If GluedFrame.IsMatch(compact) Then Return True
            Dim flat = Normalize(text)
            If ToleranceCore.IsMatch(flat) OrElse FrameValue.IsMatch(flat) Then Return True
            Dim parts = flat.Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
            If parts.Length >= 2 AndAlso IsReplaceableSymbol(parts(0)) Then
                Dim rest = String.Join(" ", parts.Skip(1))
                Return ToleranceCore.IsMatch(rest) OrElse FrameValue.IsMatch(rest)
            End If
            Return False
        End Function

        Private Shared Function CoreCompact(text As String) As String
            Dim compact = Regex.Replace(Normalize(text), "\s+", "")
            Return Regex.Replace(compact, "^[LlTtIi](?=\d)", "")
        End Function

        Private Shared Function RebuildFrame(bmp As Bitmap, el As OcrElement) As String
            Dim compact = CoreCompact(el.Text)
            Dim known = TryKnownGluedFrame(bmp, el, compact)
            If known <> "" Then Return known
            Dim m = GluedFrame.Match(compact)
            If Not m.Success OrElse el.BoundingBox Is Nothing Then Return ""

            Dim number = m.Groups("num").Value
            Dim letters = m.Groups("dat").Value
            Dim symbol = m.Groups("sym").Value
            Dim frameBox = SnapToControlFrame(bmp, el.BoundingBox)
            Dim cells = FrameCells(bmp, frameBox)
            Dim seen = SymbolFromCells(bmp, cells)
            Dim beside = SymbolBelongsToValue(el.BoundingBox, cells)
            If letters = "" AndAlso Not beside Then
                symbol = ""
                seen = ""
            End If
            If seen <> "" Then symbol = seen
            letters &= ExtraDatums(bmp, cells, letters.Length)
            Dim datums = SplitDatums(letters)

            Dim ocrDia = m.Groups("dia").Success AndAlso m.Groups("dia").Value <> ""
            Dim diameter = False
            If IsPositionSymbol(symbol) Then
                diameter = ocrDia OrElse DiameterInValueCell(bmp, cells)
            ElseIf Not IsNonDiameterSymbol(symbol) Then
                diameter = ocrDia OrElse DiameterInValueCell(bmp, cells)
            End If

            Dim value = number
            If diameter Then
                If number.StartsWith("00") Then
                    value = "Ø" & number.Substring(1)
                Else
                    value = "Ø" & number
                End If
            End If

            Dim parts As New List(Of String)()
            Dim symName = SymbolName(symbol)
            If symName <> "" Then parts.Add(symName)
            parts.Add(value)
            If datums <> "" Then parts.Add(datums)
            Return String.Join(" ", parts)
        End Function

        Private Shared Function TryKnownGluedFrame(bmp As Bitmap, el As OcrElement, compact As String) As String
            If el Is Nothing OrElse el.BoundingBox Is Nothing OrElse bmp Is Nothing Then Return ""
            Dim slash = SlashFrame.Match(compact)
            Dim letter = LetterFrame.Match(compact)
            If Not slash.Success AndAlso Not letter.Success Then Return ""
            ' //002B 本身就是平行度框，框线没扫到也按文字拆开。
            If slash.Success Then
                Return "平行度 " & ZoneNumber(slash.Groups("num").Value) & " " & SplitDatums(slash.Groups("dat").Value)
            End If

            Dim frame = SnapToControlFrame(bmp, el.BoundingBox)
            Dim textW = Math.Max(1, el.BoundingBox.Right - el.BoundingBox.Left)
            If frame.Right - frame.Left < CInt(textW * 1.15) Then Return ""
            Dim cells = FrameCells(bmp, frame)
            If cells Is Nothing OrElse cells.Count < 2 Then Return ""

            Dim symbol = SymbolFromCells(bmp, cells)
            Dim symName = SymbolName(symbol)
            If symName = "" Then Return ""
            Return symName & " " & letter.Groups("num").Value & " " & SplitDatums(letter.Groups("dat").Value)
        End Function

        Private Shared Function ApplySymbol(symbol As String, text As String) As String
            Dim symName = SymbolName(symbol)
            If symName = "" Then Return If(text, "").Trim()
            Dim letter = LetterFrame.Match(CoreCompact(text))
            If letter.Success Then
                Return symName & " " & letter.Groups("num").Value & " " & SplitDatums(letter.Groups("dat").Value)
            End If
            Return symName & " " & If(text, "").Trim()
        End Function

        Private Shared Function ZoneNumber(raw As String) As String
            If If(raw, "").Contains("."c) Then Return raw
            Dim match = Regex.Match(If(raw, ""), "^0{1,2}(\d)$")
            If match.Success Then Return "0.0" & match.Groups(1).Value
            Return raw
        End Function

        Private Shared Function SplitDatums(letters As String) As String
            If String.IsNullOrEmpty(letters) Then Return ""
            Dim parts As New List(Of String)()
            For Each ch In letters
                parts.Add(ch.ToString())
            Next
            Return String.Join(" ", parts)
        End Function

        Private Shared Function IsPositionSymbol(symbol As String) As Boolean
            Return Not String.IsNullOrEmpty(symbol) AndAlso symbol.Length = 1 AndAlso PositionSymbolChars.Contains(symbol(0))
        End Function

        Private Shared Function IsNonDiameterSymbol(symbol As String) As Boolean
            Return Not String.IsNullOrEmpty(symbol) AndAlso symbol.Length = 1 AndAlso NonDiameterSymbolChars.Contains(symbol(0))
        End Function

        Private Shared Function SymbolFromCells(bmp As Bitmap, cells As List(Of OcrBoundingBox)) As String
            If cells Is Nothing Then Return ""
            For Each cell In cells
                Dim h = Math.Max(1, cell.Bottom - cell.Top)
                Dim w = Math.Max(1, cell.Right - cell.Left)
                If w < h * 0.45 Then Continue For
                Dim box = cell
                If w > h * 1.7 Then
                    box = New OcrBoundingBox(cell.Left, cell.Top, Math.Min(cell.Right, cell.Left + CInt(h * 1.2)), cell.Bottom)
                End If
                Dim symbol = Classify(ReadInk(bmp, box.Left, box.Top, box.Right, box.Bottom))
                Dim separated = cells.Count >= 2 AndAlso w <= h * 1.8
                If symbol <> "" AndAlso (separated OrElse symbol = "⊕" OrElse symbol = "⌖" OrElse WallOnRight(bmp, box)) Then Return symbol
                Exit For
            Next
            Return ""
        End Function

        Private Shared Function SymbolBelongsToValue(textBox As OcrBoundingBox, cells As List(Of OcrBoundingBox)) As Boolean
            If textBox Is Nothing OrElse cells Is Nothing OrElse cells.Count = 0 Then Return False
            Dim cell As OcrBoundingBox = Nothing
            For Each item In cells
                Dim ch = Math.Max(1, item.Bottom - item.Top)
                Dim cw = Math.Max(1, item.Right - item.Left)
                If cw < ch * 0.45 Then Continue For
                cell = item
                Exit For
            Next
            If cell Is Nothing Then Return False
            Dim h = Math.Max(1, cell.Bottom - cell.Top)
            Dim textW = Math.Max(1, textBox.Right - textBox.Left)
            If cell.Right < textBox.Left - CInt(h * 0.4) Then Return False
            If cell.Left > textBox.Left + CInt(textW * 0.55) Then Return False
            If textW < h * 2.2 AndAlso cell.Left >= textBox.Left - 2 Then Return False
            Return True
        End Function

        Private Shared Function DiameterInValueCell(bmp As Bitmap, cells As List(Of OcrBoundingBox)) As Boolean
            If cells Is Nothing OrElse cells.Count < 2 Then Return False
            Return HasDiameterGlyph(bmp, cells(1), False)
        End Function

        Private Shared Function SnapToControlFrame(bmp As Bitmap, box As OcrBoundingBox) As OcrBoundingBox
            If bmp Is Nothing OrElse box Is Nothing Then Return box
            Dim textH = Math.Max(6, box.Bottom - box.Top)
            Dim left = Math.Max(0, box.Left - Math.Max(36, textH * 4))
            Dim top = Math.Max(0, box.Top - Math.Max(18, textH * 3))
            Dim right = Math.Min(bmp.Width - 1, box.Right + Math.Max(24, textH * 2))
            Dim bottom = Math.Min(bmp.Height - 1, box.Bottom + Math.Max(18, textH * 3))
            If right - left < 12 OrElse bottom - top < 12 Then Return box

            Dim ink = ReadInk(bmp, left, top, right, bottom, 210)
            Dim w = ink.GetLength(0)
            Dim h = ink.GetLength(1)
            Dim tx0 = Math.Max(0, Math.Min(w - 1, box.Left - left))
            Dim tx1 = Math.Max(tx0 + 1, Math.Min(w - 1, box.Right - left))
            Dim ty0 = Math.Max(0, Math.Min(h - 1, box.Top - top))
            Dim ty1 = Math.Max(ty0 + 1, Math.Min(h - 1, box.Bottom - top))

            Dim topRule = -1
            Dim botRule = -1
            For y = 0 To h - 1
                If Not RowCoversSpan(ink, y, tx0, tx1) Then Continue For
                If y <= ty0 + 1 AndAlso (topRule < 0 OrElse y > topRule) Then topRule = y
                If y >= ty1 - 1 AndAlso (botRule < 0 OrElse y < botRule) Then botRule = y
            Next
            If topRule < 0 OrElse botRule < 0 OrElse botRule - topRule < 8 Then Return box

            Dim lines = FindSpanDividers(ink, topRule, botRule)
            If lines.Count < 2 Then Return box
            Dim frameLeft = left + lines(0)
            Dim frameRight = left + lines(lines.Count - 1)
            Dim frameTop = top + topRule
            Dim frameBottom = top + botRule
            If frameRight - frameLeft < 12 Then Return box
            Return New OcrBoundingBox(frameLeft, frameTop, frameRight, frameBottom)
        End Function

        Private Shared Function RowCoversSpan(ink As Boolean(,), y As Integer, x0 As Integer, x1 As Integer) As Boolean
            Dim hits = 0
            Dim span = Math.Max(1, x1 - x0)
            For x = x0 To x1
                If ink(x, y) Then hits += 1
            Next
            Return hits >= span * 0.45
        End Function

        Private Shared Function FindSpanDividers(ink As Boolean(,), y0 As Integer, y1 As Integer) As List(Of Integer)
            Dim runs As New List(Of Integer)()
            Dim w = ink.GetLength(0)
            Dim span = Math.Max(1, y1 - y0)
            Dim runStart = -1
            For x = 0 To w - 1
                Dim hits = 0
                For y = y0 To y1
                    If ink(x, y) Then hits += 1
                Next
                Dim isLine = hits >= span * 0.82
                If isLine Then
                    If runStart < 0 Then runStart = x
                ElseIf runStart >= 0 Then
                    runs.Add((runStart + x - 1) \ 2)
                    runStart = -1
                End If
            Next
            If runStart >= 0 Then runs.Add((runStart + w - 1) \ 2)
            Return runs
        End Function

        Private Shared Function ExtraDatums(bmp As Bitmap, cells As List(Of OcrBoundingBox), knownCount As Integer) As String
            If cells Is Nothing OrElse cells.Count < 3 + knownCount Then Return ""
            Dim extra = ""
            For i = 2 + knownCount To cells.Count - 1
                Dim cell = cells(i)
                Dim h = Math.Max(1, cell.Bottom - cell.Top)
                Dim w = Math.Max(1, cell.Right - cell.Left)
                If w > h * 1.6 Then Exit For
                Dim letter = ClassifyDatum(ReadInk(bmp, cell.Left, cell.Top, cell.Right, cell.Bottom))
                If letter = "" Then Exit For
                extra &= letter
            Next
            Return extra
        End Function

        Private Shared Function ClassifyDatum(cell As Boolean(,)) As String
            If cell Is Nothing Then Return ""
            StripFrameRails(cell)
            StripEdgeWalls(cell)
            Dim sample = NormalizeInk(cell, 24)
            If sample Is Nothing Then Return ""
            Dim rightGap = 1.0 - ColCoverage(sample, 0.72, 0.98, 0.28, 0.72)
            Dim leftSide = ColCoverage(sample, 0.02, 0.3, 0.12, 0.88)
            Dim top = RowCoverage(sample, 0.0, 0.24)
            Dim bottom = RowCoverage(sample, 0.76, 1.0)
            Dim hollow = 1.0 - BoxCoverage(sample, 0.35, 0.65, 0.35, 0.65)
            If rightGap > 0.62 AndAlso leftSide > 0.35 AndAlso top > 0.28 AndAlso bottom > 0.28 AndAlso hollow > 0.4 Then
                Return "C"
            End If
            Return ""
        End Function

        Private Shared Function FrameCells(bmp As Bitmap, box As OcrBoundingBox) As List(Of OcrBoundingBox)
            Dim result As New List(Of OcrBoundingBox)()
            Dim ink = ReadInk(bmp, box.Left, box.Top, box.Right, box.Bottom, 210)
            If ink Is Nothing Then Return result
            Dim h = ink.GetLength(1)
            Dim w = ink.GetLength(0)
            Dim lines = FindFrameDividers(ink, w, h)
            If lines.Count = 0 Then
                Dim side = Math.Max(4, Math.Min(w - 1, CInt(h * 1.15)))
                result.Add(New OcrBoundingBox(box.Left, box.Top, box.Left + side, box.Bottom))
                Return result
            End If

            Dim cuts As New List(Of Integer) From {0}
            For Each line In lines
                If line - cuts(cuts.Count - 1) >= h * 0.78 Then cuts.Add(line)
            Next
            If w - 1 - cuts(cuts.Count - 1) >= h * 0.78 Then cuts.Add(w - 1)

            For i = 0 To cuts.Count - 2
                Dim x0 = cuts(i)
                Dim x1 = cuts(i + 1)
                If x1 - x0 < 4 Then Continue For
                result.Add(New OcrBoundingBox(box.Left + x0 + 1, box.Top, box.Left + x1 - 1, box.Bottom))
            Next
            If result.Count = 0 Then
                Dim side = Math.Max(4, Math.Min(w - 1, CInt(h * 1.15)))
                result.Add(New OcrBoundingBox(box.Left, box.Top, box.Left + side, box.Bottom))
            End If
            Return result
        End Function

        Private Shared Function FindFrameDividers(ink As Boolean(,), w As Integer, h As Integer) As List(Of Integer)
            Dim runs As New List(Of Integer)()
            Dim runStart = -1
            Dim y0 = Math.Min(h - 1, Math.Max(0, CInt(h * 0.12)))
            Dim y1 = Math.Max(y0, CInt(h * 0.88))
            Dim span = Math.Max(1, y1 - y0)
            For x = 0 To w - 1
                Dim hits = 0
                For y = y0 To y1
                    If ink(x, y) Then hits += 1
                Next
                Dim isLine = hits >= span * 0.78
                If isLine Then
                    If runStart < 0 Then runStart = x
                ElseIf runStart >= 0 Then
                    runs.Add((runStart + x - 1) \ 2)
                    runStart = -1
                End If
            Next
            If runStart >= 0 Then runs.Add((runStart + w - 1) \ 2)
            Return runs
        End Function

        Private Shared Function WithDiameter(text As String) As String
            Dim s = If(text, "").Trim()
            Dim m = Regex.Match(s, "^(.*?)([OØøΦφ⌀∅]?\s*\d[\d.]*)\s*$")
            If Not m.Success Then Return "Ø" & s
            Dim head = m.Groups(1).Value
            Dim num = m.Groups(2).Value.Trim()
            If num.StartsWith("Ø") OrElse num.StartsWith("ø") Then Return s
            If Regex.IsMatch(num, "^[OΦφ⌀∅]") Then
                num = "Ø" & num.Substring(1).TrimStart()
            ElseIf num.StartsWith("00") Then
                num = "Ø" & num.Substring(1)
            Else
                num = "Ø" & num
            End If
            Return (head & num).Trim()
        End Function

        Private Shared Function IsReplaceableSymbol(text As String) As Boolean
            Dim s = If(text, "").Trim()
            If s.Length = 0 OrElse s.Length > 2 Then Return False
            If Regex.IsMatch(s, "\d") Then Return False
            If Regex.IsMatch(s, "^[A-Z]$") Then Return False
            Return True
        End Function

        Private Shared Function Normalize(text As String) As String
            Dim flat = If(text, "").Replace(vbTab, " ").Replace(vbCr, " ").Replace(vbLf, " ")
            Return Regex.Replace(flat, "\s+", " ").Trim()
        End Function

        Private Shared Function JoinRow(row As List(Of OcrElement)) As String
            Return Normalize(String.Join(" ", row.Select(Function(el) If(el.Text, "").Trim()).Where(Function(t) t <> "")))
        End Function

        Private Shared Function GroupRows(items As List(Of OcrElement)) As List(Of List(Of OcrElement))
            Dim ordered = items.Where(Function(el) el IsNot Nothing AndAlso el.BoundingBox IsNot Nothing).
                OrderBy(Function(el) (el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2).
                ThenBy(Function(el) el.BoundingBox.Left).ToList()
            Dim rows As New List(Of List(Of OcrElement))()
            For Each el In ordered
                Dim cy = (el.BoundingBox.Top + el.BoundingBox.Bottom) \ 2
                Dim h = Math.Max(8, el.BoundingBox.Bottom - el.BoundingBox.Top)
                Dim placed = False
                For Each row In rows
                    Dim anchor = row(0)
                    Dim ay = (anchor.BoundingBox.Top + anchor.BoundingBox.Bottom) \ 2
                    Dim ah = Math.Max(h, anchor.BoundingBox.Bottom - anchor.BoundingBox.Top)
                    If Math.Abs(cy - ay) <= ah * 0.6 AndAlso Not IsStandaloneFrame(el) AndAlso Not IsStandaloneFrame(anchor) Then
                        row.Add(el)
                        placed = True
                        Exit For
                    End If
                Next
                If Not placed Then rows.Add(New List(Of OcrElement) From {el})
            Next
            Return rows
        End Function

        Private Shared Function IsStandaloneFrame(el As OcrElement) As Boolean
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return False
            Dim w = el.BoundingBox.Right - el.BoundingBox.Left
            Dim h = Math.Max(1, el.BoundingBox.Bottom - el.BoundingBox.Top)
            Return w > h * 1.8 AndAlso LooksLikeFrame(el.Text)
        End Function

        Private Shared Function BoxLeft(el As OcrElement) As Integer
            If el Is Nothing OrElse el.BoundingBox Is Nothing Then Return 0
            Return el.BoundingBox.Left
        End Function

        Private Shared Function ClassifyBox(bmp As Bitmap, box As OcrBoundingBox, expandLeft As Boolean) As String
            If box Is Nothing OrElse bmp Is Nothing Then Return ""
            Dim rowH = Math.Max(8, box.Bottom - box.Top)
            Dim left = box.Left
            If expandLeft Then left = box.Left - CInt(rowH * 2.4)
            left = Math.Max(0, left)
            Dim top = Math.Max(0, box.Top - 3)
            Dim right = Math.Min(bmp.Width - 1, Math.Max(box.Right, box.Left + 4))
            Dim bottom = Math.Min(bmp.Height - 1, box.Bottom + 3)
            If right - left < 8 OrElse bottom - top < 8 Then Return ""

            Dim w = right - left + 1
            Dim h = bottom - top + 1
            Dim ink(w - 1, h - 1) As Boolean
            Dim dark = 0
            For y = 0 To h - 1
                For x = 0 To w - 1
                    Dim color = bmp.GetPixel(left + x, top + y)
                    Dim lum = (CInt(color.R) * 299 + CInt(color.G) * 587 + CInt(color.B) * 114) \ 1000
                    Dim onInk = lum < Services.DomainPack.CorrectionProfile.ResolveInk(150)
                    ink(x, y) = onInk
                    If onInk Then dark += 1
                Next
            Next
            If dark < 16 Then Return ""

            Dim lines = FindVerticalLines(ink, w, h)
            For i = 0 To lines.Count - 2
                Dim gap = lines(i + 1) - lines(i)
                If gap < h * 0.45 OrElse gap > h * 2.2 Then Continue For
                Dim symbol = Classify(Crop(ink, lines(i) + 1, lines(i + 1) - 1, 0, h - 1))
                If symbol <> "" Then Return symbol
            Next

            If Not expandLeft Then Return Classify(Crop(ink, 0, w - 1, 0, h - 1))

            Dim textStart = Math.Max(0, box.Left - left)
            If textStart < 6 Then Return ""
            Return Classify(Crop(ink, 0, textStart - 1, 0, h - 1))
        End Function

        Private Shared Function HasDiameterGlyph(bmp As Bitmap, box As OcrBoundingBox, Optional padLeft As Boolean = True) As Boolean
            If bmp Is Nothing OrElse box Is Nothing Then Return False
            Dim rowH = Math.Max(8, box.Bottom - box.Top)
            Dim left = box.Left
            If padLeft Then left = Math.Max(0, box.Left - CInt(rowH * 0.3))
            Dim width = Math.Min(box.Right - left, CInt(rowH * 1.15))
            If width < 6 Then Return False
            Dim right = Math.Min(bmp.Width - 1, left + width)
            Dim top = Math.Max(0, box.Top - 2)
            Dim bottom = Math.Min(bmp.Height - 1, box.Bottom + 2)
            Dim cell = ReadInk(bmp, left, top, right, bottom)
            Return LooksLikeDiameterSymbol(cell)
        End Function

        ''' <summary>一小块墨迹是不是「圆圈加一斜杠」的直径符号。竖横排共用。</summary>
        Private Shared Function LooksLikeDiameterSymbol(cell As Boolean(,)) As Boolean
            If cell Is Nothing Then Return False
            StripEdgeWalls(cell)
            Dim sample = NormalizeInk(cell, 24)
            If sample Is Nothing Then Return False
            Dim d1 = DiagCoverage(sample, False)
            Dim d2 = DiagCoverage(sample, True)
            Dim slash = Math.Max(d1, d2)
            Dim other = Math.Min(d1, d2)
            Dim midH = RowCoverage(sample, 0.4, 0.6)
            Dim midV = ColCoverage(sample, 0.4, 0.6, 0.1, 0.9)
            Dim ring = (RowCoverage(sample, 0.08, 0.28) + RowCoverage(sample, 0.72, 0.94) + ColCoverage(sample, 0.06, 0.24, 0.2, 0.8) + ColCoverage(sample, 0.76, 0.94, 0.2, 0.8)) / 4.0
            ' 直径：一条斜线明显长过另一条。圆圈或数字 0 两条对角差不多，不能当成直径。
            Return slash > 0.55 AndAlso slash - other > 0.15 AndAlso ring > 0.16 AndAlso Not (midH > 0.5 AndAlso midV > 0.5)
        End Function

        Private Shared Function ReadInk(bmp As Bitmap, left As Integer, top As Integer, right As Integer, bottom As Integer, Optional threshold As Integer = 150) As Boolean(,)
            threshold = Services.DomainPack.CorrectionProfile.ResolveInk(threshold)
            Dim w = Math.Max(1, right - left + 1)
            Dim h = Math.Max(1, bottom - top + 1)
            Dim ink(w - 1, h - 1) As Boolean
            For y = 0 To h - 1
                For x = 0 To w - 1
                    Dim px = Math.Max(0, Math.Min(bmp.Width - 1, left + x))
                    Dim py = Math.Max(0, Math.Min(bmp.Height - 1, top + y))
                    Dim color = bmp.GetPixel(px, py)
                    Dim lum = (CInt(color.R) * 299 + CInt(color.G) * 587 + CInt(color.B) * 114) \ 1000
                    ink(x, y) = lum < threshold
                Next
            Next
            Return ink
        End Function

        Private Shared Function WallOnRight(bmp As Bitmap, box As OcrBoundingBox) As Boolean
            If bmp Is Nothing OrElse box Is Nothing Then Return False
            Dim x0 = Math.Max(0, box.Right - 2)
            Dim x1 = Math.Min(bmp.Width - 1, box.Right + 2)
            If x1 <= x0 Then Return False
            Dim ink = ReadInk(bmp, x0, box.Top, x1, box.Bottom, 210)
            Dim h = ink.GetLength(1)
            Dim w = ink.GetLength(0)
            For x = 0 To w - 1
                Dim hits = 0
                For y = 0 To h - 1
                    If ink(x, y) Then hits += 1
                Next
                If hits >= h * 0.72 Then Return True
            Next
            Return False
        End Function

        Private Shared Function DiagCoverage(sample As Boolean(,), anti As Boolean) As Double
            Dim n = sample.GetLength(0)
            Dim hits = 0
            For i = 0 To n - 1
                Dim y = If(anti, n - 1 - i, i)
                For d = -2 To 2
                    Dim x = i + d
                    If x >= 0 AndAlso x < n AndAlso sample(x, y) Then
                        hits += 1
                        Exit For
                    End If
                Next
            Next
            Return hits / CDbl(n)
        End Function

        Private Shared Function FindVerticalLines(ink As Boolean(,), w As Integer, h As Integer) As List(Of Integer)
            Dim runs As New List(Of Integer)()
            Dim runStart = -1
            For x = 0 To w - 1
                Dim hits = 0
                For y = 0 To h - 1
                    If ink(x, y) Then hits += 1
                Next
                Dim isLine = hits >= h * 0.72
                If isLine Then
                    If runStart < 0 Then runStart = x
                ElseIf runStart >= 0 Then
                    runs.Add((runStart + x - 1) \ 2)
                    runStart = -1
                End If
            Next
            If runStart >= 0 Then runs.Add((runStart + w - 1) \ 2)
            Return runs
        End Function

        Private Shared Function Crop(ink As Boolean(,), x0 As Integer, x1 As Integer, y0 As Integer, y1 As Integer) As Boolean(,)
            If ink Is Nothing Then Return Nothing
            Dim w = ink.GetLength(0)
            Dim h = ink.GetLength(1)
            x0 = Math.Max(0, x0)
            y0 = Math.Max(0, y0)
            x1 = Math.Min(w - 1, x1)
            y1 = Math.Min(h - 1, y1)
            If x1 - x0 < 4 OrElse y1 - y0 < 4 Then Return Nothing
            Dim cw = x1 - x0 + 1
            Dim ch = y1 - y0 + 1
            Dim cell(cw - 1, ch - 1) As Boolean
            For y = 0 To ch - 1
                For x = 0 To cw - 1
                    cell(x, y) = ink(x0 + x, y0 + y)
                Next
            Next
            Return cell
        End Function

        Private Shared Sub StripEdgeWalls(cell As Boolean(,))
            If cell Is Nothing Then Return
            Dim w = cell.GetLength(0)
            Dim h = cell.GetLength(1)
            For x = 0 To w - 1
                If x > w * 0.2 AndAlso x < w * 0.8 Then Continue For
                Dim hits = 0
                For y = 0 To h - 1
                    If cell(x, y) Then hits += 1
                Next
                If hits < h * 0.75 Then Continue For
                For y = 0 To h - 1
                    cell(x, y) = False
                Next
            Next
        End Sub

        Private Shared Sub StripFrameRails(cell As Boolean(,))
            If cell Is Nothing Then Return
            Dim w = cell.GetLength(0)
            Dim h = cell.GetLength(1)
            Dim band = Math.Min(2, Math.Max(1, h \ 6))
            For y = 0 To h - 1
                If y >= band AndAlso y < h - band Then Continue For
                Dim hits = 0
                For x = 0 To w - 1
                    If cell(x, y) Then hits += 1
                Next
                If hits < w * 0.75 Then Continue For
                For x = 0 To w - 1
                    cell(x, y) = False
                Next
            Next
        End Sub

        Private Shared Function Classify(cell As Boolean(,)) As String
            If cell Is Nothing Then Return ""
            StripFrameRails(cell)
            StripEdgeWalls(cell)
            Dim sample = NormalizeInk(cell, 24)
            If sample Is Nothing Then Return ""

            Dim bottom = RowCoverage(sample, 0.68, 0.98)
            Dim stem = ColCoverage(sample, 0.38, 0.62, 0.08, 0.92)
            Dim side = (ColCoverage(sample, 0.02, 0.24, 0.2, 0.62) + ColCoverage(sample, 0.76, 0.98, 0.2, 0.62)) / 2.0
            Dim crown = BoxCoverage(sample, 0.32, 0.68, 0.04, 0.32)
            Dim hollow = 1.0 - BoxCoverage(sample, 0.34, 0.66, 0.34, 0.66)

            Dim midH = RowCoverage(sample, 0.4, 0.6)
            Dim midV = ColCoverage(sample, 0.4, 0.6, 0.1, 0.9)
            Dim ring = (RowCoverage(sample, 0.0, 0.16) + RowCoverage(sample, 0.84, 1.0) + ColCoverage(sample, 0.0, 0.16, 0.12, 0.88) + ColCoverage(sample, 0.84, 1.0, 0.12, 0.88)) / 4.0
            Dim pocket = (BoxCoverage(sample, 0.18, 0.36, 0.18, 0.36) + BoxCoverage(sample, 0.64, 0.82, 0.18, 0.36) + BoxCoverage(sample, 0.18, 0.36, 0.64, 0.82) + BoxCoverage(sample, 0.64, 0.82, 0.64, 0.82)) / 4.0

            ' 位置度：外圈是圆，中间横竖十字。中心线只有十字、没有圆，不能当成位置度。
            If midH > 0.5 AndAlso midV > 0.5 AndAlso ring > 0.32 AndAlso pocket < 0.45 Then
                Return "⊕"
            End If
            Dim elongation = 0.0
            Dim degrees = 0.0
            InkAxis(sample, elongation, degrees)
            Dim tilt = Math.Abs(degrees)
            If tilt > 90 Then tilt = 180 - tilt
            Dim topBand = RowCoverage(sample, 0.0, 0.18)
            Dim botBand = RowCoverage(sample, 0.82, 1.0)
            Dim base = RowCoverage(sample, 0.9, 1.0)
            ' 圆跳动是一根细长斜箭头。先按主轴方向认，避免被横线规则收成圆柱度。
            If elongation > 0.58 AndAlso tilt > 18 AndAlso tilt < 78 Then
                Return "↗"
            End If
            ' 圆柱度的主轴是水平的：上下两条长横线夹着圆，不是斜线。
            If tilt < 15 AndAlso elongation < 0.72 AndAlso topBand > 0.75 AndAlso botBand > 0.75 AndAlso ring > 0.3 AndAlso hollow > 0.45 Then
                Return "⌭"
            End If
            ' 垂直度：最底下一条贯通横线，中间一条竖线。斜箭头的尾端不算底边。
            If base > 0.72 AndAlso stem > 0.4 AndAlso side < 0.35 AndAlso topBand < 0.5 AndAlso tilt > 70 Then
                Return "⊥"
            End If
            ' 面轮廓度：上拱加底边，拱内是空的。
            If bottom > 0.42 AndAlso side > 0.22 AndAlso crown > 0.12 AndAlso hollow > 0.45 AndAlso stem < 0.42 Then
                Return "⌓"
            End If
            ' 线轮廓度：只有上拱，没有底边。
            If bottom < 0.28 AndAlso side > 0.18 AndAlso crown > 0.12 AndAlso hollow > 0.45 Then
                Return "⌒"
            End If
            Return ""
        End Function

        Private Shared Sub InkAxis(sample As Boolean(,), ByRef elongation As Double, ByRef degrees As Double)
            elongation = 0
            degrees = 0
            If sample Is Nothing Then Return
            Dim n = sample.GetLength(0)
            Dim count = 0
            Dim sx = 0.0
            Dim sy = 0.0
            For y = 0 To n - 1
                For x = 0 To n - 1
                    If Not sample(x, y) Then Continue For
                    count += 1
                    sx += x
                    sy += y
                Next
            Next
            If count < 8 Then Return
            Dim mx = sx / count
            Dim my = sy / count
            Dim cxx = 0.0
            Dim cyy = 0.0
            Dim cxy = 0.0
            For y = 0 To n - 1
                For x = 0 To n - 1
                    If Not sample(x, y) Then Continue For
                    Dim dx = x - mx
                    Dim dy = y - my
                    cxx += dx * dx
                    cyy += dy * dy
                    cxy += dx * dy
                Next
            Next
            Dim trace = cxx + cyy
            Dim det = cxx * cyy - cxy * cxy
            Dim disc = Math.Sqrt(Math.Max(0, trace * trace / 4.0 - det))
            Dim major = trace / 2.0 + disc
            Dim minor = Math.Max(0, trace / 2.0 - disc)
            If major <= 0 Then Return
            elongation = 1 - minor / major
            degrees = 0.5 * Math.Atan2(2 * cxy, cxx - cyy) * 180.0 / Math.PI
        End Sub

        Private Shared Function NormalizeInk(cell As Boolean(,), n As Integer) As Boolean(,)
            Dim w = cell.GetLength(0)
            Dim h = cell.GetLength(1)
            Dim minX = w
            Dim minY = h
            Dim maxX = -1
            Dim maxY = -1
            Dim count = 0
            For y = 0 To h - 1
                For x = 0 To w - 1
                    If Not cell(x, y) Then Continue For
                    count += 1
                    If x < minX Then minX = x
                    If y < minY Then minY = y
                    If x > maxX Then maxX = x
                    If y > maxY Then maxY = y
                Next
            Next
            If count < 12 OrElse maxX - minX < 3 OrElse maxY - minY < 3 Then Return Nothing

            Dim sample(n - 1, n - 1) As Boolean
            Dim bw = maxX - minX + 1
            Dim bh = maxY - minY + 1
            For y = 0 To n - 1
                Dim sy = minY + Math.Min(bh - 1, CInt(y * bh / n))
                For x = 0 To n - 1
                    Dim sx = minX + Math.Min(bw - 1, CInt(x * bw / n))
                    sample(x, y) = cell(sx, sy)
                Next
            Next
            Return sample
        End Function

        Private Shared Function RowCoverage(sample As Boolean(,), y0 As Double, y1 As Double) As Double
            Dim n = sample.GetLength(0)
            Dim fromY = CInt(n * y0)
            Dim toY = Math.Min(n - 1, CInt(n * y1))
            Dim hitCols = 0
            For x = 0 To n - 1
                For y = fromY To toY
                    If sample(x, y) Then
                        hitCols += 1
                        Exit For
                    End If
                Next
            Next
            Return hitCols / CDbl(n)
        End Function

        Private Shared Function ColCoverage(sample As Boolean(,), x0 As Double, x1 As Double, y0 As Double, y1 As Double) As Double
            Dim n = sample.GetLength(0)
            Dim fromX = CInt(n * x0)
            Dim toX = Math.Min(n - 1, CInt(n * x1))
            Dim fromY = CInt(n * y0)
            Dim toY = Math.Min(n - 1, CInt(n * y1))
            Dim rows = Math.Max(1, toY - fromY + 1)
            Dim hitRows = 0
            For y = fromY To toY
                For x = fromX To toX
                    If sample(x, y) Then
                        hitRows += 1
                        Exit For
                    End If
                Next
            Next
            Return hitRows / CDbl(rows)
        End Function

        Private Shared Function BoxCoverage(sample As Boolean(,), x0 As Double, x1 As Double, y0 As Double, y1 As Double) As Double
            Dim n = sample.GetLength(0)
            Dim fromX = CInt(n * x0)
            Dim toX = Math.Min(n - 1, CInt(n * x1))
            Dim fromY = CInt(n * y0)
            Dim toY = Math.Min(n - 1, CInt(n * y1))
            Dim total = 0
            Dim hit = 0
            For y = fromY To toY
                For x = fromX To toX
                    total += 1
                    If sample(x, y) Then hit += 1
                Next
            Next
            If total = 0 Then Return 0
            Return hit / CDbl(total)
        End Function

#Region "阶段 6：FCF (Feature Control Frame) 形位公差框检测"

        ''' <summary>
        ''' 扩展版 ReadInk：既把深灰色像素（lum < 150）当作墨迹，也把蓝色填充像素当作墨迹。
        ''' 真实工程图中 GD&T 框和球标经常用实心蓝色填充，常规灰度阈值会漏掉。
        ''' 蓝色判定：R < 130 AND B > G + 30 AND B > 160。
        ''' </summary>
        Private Shared Function ReadInkAll(bmp As Bitmap, left As Integer, top As Integer, right As Integer, bottom As Integer, Optional threshold As Integer = 150) As Boolean(,)
            threshold = Services.DomainPack.CorrectionProfile.ResolveInk(threshold)
            Dim w = Math.Max(1, right - left + 1)
            Dim h = Math.Max(1, bottom - top + 1)
            Dim ink(w - 1, h - 1) As Boolean
            For y = 0 To h - 1
                For x = 0 To w - 1
                    Dim px = Math.Max(0, Math.Min(bmp.Width - 1, left + x))
                    Dim py = Math.Max(0, Math.Min(bmp.Height - 1, top + y))
                    Dim color = bmp.GetPixel(px, py)
                    Dim lum = (CInt(color.R) * 299 + CInt(color.G) * 587 + CInt(color.B) * 114) \ 1000
                    Dim isDark = lum < threshold
                    Dim isBlueFill = color.R < 130 AndAlso color.B > CInt(color.G) + 30 AndAlso color.B > 160
                    ink(x, y) = isDark OrElse isBlueFill
                Next
            Next
            Return ink
        End Function

        ''' <summary>
        ''' 在指定行 y 上，扫描 x_start..x_end，返回所有长度 ≥ min_len 的水平墨迹段。
        ''' 每段返回 (startX, endX, length)。
        ''' </summary>
        Private Shared Function FindHSegments(ink As Boolean(,), y As Integer, x_start As Integer, x_end As Integer, min_len As Integer) As List(Of (Integer, Integer, Integer))
            Dim result As New List(Of (Integer, Integer, Integer))()
            If ink Is Nothing Then Return result
            Dim W = ink.GetLength(0)
            Dim H = ink.GetLength(1)
            If y < 0 OrElse y >= H Then Return result
            x_start = Math.Max(0, x_start)
            x_end = Math.Min(W - 1, x_end)
            If x_end < x_start Then Return result

            Dim inRun = False
            Dim runStart = 0
            For x = x_start To x_end
                If ink(x, y) Then
                    If Not inRun Then
                        inRun = True
                        runStart = x
                    End If
                Else
                    If inRun Then
                        Dim len_ = x - runStart
                        If len_ >= min_len Then result.Add((runStart, x - 1, len_))
                        inRun = False
                    End If
                End If
            Next
            If inRun Then
                Dim len_ = x_end - runStart + 1
                If len_ >= min_len Then result.Add((runStart, x_end, len_))
            End If
            Return result
        End Function

        ''' <summary>
        ''' 在指定列 x 上，扫描 y_start..y_end，返回所有长度 ≥ min_len 的垂直墨迹段。
        ''' 每段返回 (startY, endY, length)。
        ''' </summary>
        Private Shared Function FindVSegments(ink As Boolean(,), x As Integer, y_start As Integer, y_end As Integer, min_len As Integer) As List(Of (Integer, Integer, Integer))
            Dim result As New List(Of (Integer, Integer, Integer))()
            If ink Is Nothing Then Return result
            Dim W = ink.GetLength(0)
            Dim H = ink.GetLength(1)
            If x < 0 OrElse x >= W Then Return result
            y_start = Math.Max(0, y_start)
            y_end = Math.Min(H - 1, y_end)
            If y_end < y_start Then Return result

            Dim inRun = False
            Dim runStart = 0
            For y = y_start To y_end
                If ink(x, y) Then
                    If Not inRun Then
                        inRun = True
                        runStart = y
                    End If
                Else
                    If inRun Then
                        Dim len_ = y - runStart
                        If len_ >= min_len Then result.Add((runStart, y - 1, len_))
                        inRun = False
                    End If
                End If
            Next
            If inRun Then
                Dim len_ = y_end - runStart + 1
                If len_ >= min_len Then result.Add((runStart, y_end, len_))
            End If
            Return result
        End Function

        ''' <summary>
        ''' 公差值候选文本的判定正则：纯小数 / 含 ± / 含 字母-数字组合（如 0.05A、0.02 B）。
        ''' </summary>
        Private Shared ReadOnly ToleranceNumeric As New Regex("^\d+\.?\d*$", RegexOptions.Compiled)
        Private Shared ReadOnly ToleranceWithPlusMinus As New Regex("±", RegexOptions.Compiled)
        Private Shared ReadOnly ToleranceWithDatum As New Regex("^\d+[\.\d]*\s*[A-Z]{1,3}$", RegexOptions.Compiled)

        ''' <summary>
        ''' FCF 形位公差框检测主入口。
        ''' 利用看起来像公差值的 OCR 元素作为锚点，向上/下/左/右扫描完整的矩形框线，
        ''' 然后在最左侧格子里用 InferFcfSymbol 推断形位公差的**中文名**，
        ''' 最后把原元素改写成「中文名 + 值 + 基准」（如 位置度 0.05 A B）——
        ''' 这就是解析脚本示例的写法，不再另建 FCF: 前缀元素（那条文本解析不出来，还和原值重复）。
        ''' </summary>
        Private Shared Sub DetectFcfFrames(bmp As Bitmap, doc As OcrDocument, targets As List(Of OcrElement))
            If bmp Is Nothing OrElse doc Is Nothing OrElse targets Is Nothing Then Return

            For Each el In targets
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                Dim txt = If(el.Text, "").Trim()

                ' 公差值候选判定
                Dim isTol = ToleranceNumeric.IsMatch(txt) OrElse _
                            ToleranceWithPlusMinus.IsMatch(txt) OrElse _
                            ToleranceWithDatum.IsMatch(txt)
                If Not isTol Then Continue For

                ' 扩展搜索区域（向四周各 pad 50px）
                Dim box = ExpandFcfSearchBox(bmp, el.BoundingBox, 50)
                If box Is Nothing Then Continue For

                Dim ink = ReadInkAll(bmp, box.Left, box.Top, box.Right, box.Bottom)
                If ink Is Nothing Then Continue For
                Dim W = ink.GetLength(0)
                Dim H = ink.GetLength(1)

                ' 把元素坐标转换成 ink 数组内的局部坐标
                Dim ex0 = Math.Max(0, Math.Min(W - 1, el.BoundingBox.Left - box.Left))
                Dim ex1 = Math.Max(ex0 + 1, Math.Min(W - 1, el.BoundingBox.Right - box.Left))
                Dim ey0 = Math.Max(0, Math.Min(H - 1, el.BoundingBox.Top - box.Top))
                Dim ey1 = Math.Max(ey0 + 1, Math.Min(H - 1, el.BoundingBox.Bottom - box.Top))

                Dim searchPad = 20
                Dim frameMinLen = ex1 - ex0 + 10
                Dim textH = Math.Max(12, ey1 - ey0)
                Dim textW = Math.Max(12, ex1 - ex0)
                ' 只在文字自己旁边找框。扫到整行会把 11.3、圆跳动、Rz 25 收成一个大框。
                Dim vertReach = CInt(textH * 1.3)
                Dim sideReach = Math.Max(CInt(textH * 1.2), CInt(textW * 0.4))

                ' --- 1) 向上扫描寻找顶部横线 ---
                Dim topY = -1
                Dim topLimit = Math.Max(0, ey0 - vertReach)
                For y = ey0 - 1 To topLimit Step -1
                    Dim hSegs = FindHSegments(ink, y, ex0 - searchPad, ex1 + searchPad, frameMinLen)
                    If hSegs.Count > 0 Then
                        topY = y
                        Exit For
                    End If
                Next

                ' --- 2) 向下扫描寻找底部横线 ---
                Dim botY = -1
                Dim botLimit = Math.Min(H - 1, ey1 + vertReach)
                For y = ey1 + 1 To botLimit
                    Dim hSegs = FindHSegments(ink, y, ex0 - searchPad, ex1 + searchPad, frameMinLen)
                    If hSegs.Count > 0 Then
                        botY = y
                        Exit For
                    End If
                Next

                If topY < 0 OrElse botY < 0 Then Continue For
                Dim frameHeight = botY - topY + 1
                If frameHeight < 10 Then Continue For

                ' --- 3) 在 topY..botY 范围内，找最左和最右的竖线 ---
                Dim leftX = -1
                Dim rightX = -1
                Dim minVLen = CInt(frameHeight * 0.8)

                ' 左侧：只扫文字左侧很短的一段，符号格就在那里
                For x = Math.Max(0, ex0 - searchPad) To Math.Max(0, ex0 - sideReach) Step -1
                    Dim vSegs = FindVSegments(ink, x, topY, botY, minVLen)
                    If vSegs.Count > 0 Then
                        leftX = x
                        Exit For
                    End If
                Next

                ' 右侧：同样不越过相邻的另一条尺寸
                For x = Math.Min(W - 1, ex1 + searchPad) To Math.Min(W - 1, ex1 + sideReach)
                    Dim vSegs = FindVSegments(ink, x, topY, botY, minVLen)
                    If vSegs.Count > 0 Then
                        rightX = x
                        Exit For
                    End If
                Next

                If leftX < 0 OrElse rightX < 0 Then Continue For
                If rightX - leftX < frameHeight Then Continue For
                If rightX - leftX > textW * 2.2 Then Continue For

                ' --- 4) 寻找内部竖分隔线 ---
                Dim dividers As New List(Of Integer)()
                For x = leftX + 10 To rightX - 10
                    Dim vSegs = FindVSegments(ink, x, topY, botY, CInt(frameHeight * 0.8))
                    If vSegs.Count > 0 Then dividers.Add(x)
                Next
                ' 合并相近的分隔线（取均值）
                dividers = MergeCloseDividers(dividers, 3)

                ' --- 5) 最左侧符号格子 ---
                Dim symX0 = leftX + 1
                Dim symX1 = If(dividers.Count > 0, dividers(0) - 1, rightX - 1)
                Dim symY0 = topY + 1
                Dim symY1 = botY - 1
                If symX1 - symX0 < 3 OrElse symY1 - symY0 < 3 Then Continue For

                ' --- 6) 对符号格区域独立调用 ReadInkAll，再推断 GD&T 符号 ---
                ' 转换为图纸绝对坐标，在符号格子区域重新提取 ink，避免整图 OCR 污染
                Dim symAbsLeft = box.Left + symX0
                Dim symAbsTop = box.Top + symY0
                Dim symAbsRight = box.Left + symX1
                Dim symAbsBottom = box.Top + symY1
                Dim symInk = ReadInkAll(bmp, symAbsLeft, symAbsTop, symAbsRight, symAbsBottom)
                If symInk Is Nothing Then Continue For
                Dim symW = symInk.GetLength(0)
                Dim symH = symInk.GetLength(1)
                If symW < 5 OrElse symH < 5 Then Continue For
                ' 局部坐标内缩 2px 去掉框线
                Dim gdtName = InferFcfSymbol(symInk, 2, 2, symW - 3, symH - 3)
                ' 符号认不出（空框）就别写：旧写法输出 "FCF:Unknown 0.05"，下游解析脚本一条都出不来，
                ' 而原 OCR 文本本来就在，不写只是少一条冗余，不会丢内容。
                If gdtName = "" Then Continue For

                ' --- 7) 把结果写回原数值元素：值 + 基准 → 中文形位名 + 值 + 基准 ---
                ' 输出即解析脚本示例的写法（"位置度 0.05 A B"），不写符号、不写英文、不写 FCF: 前缀。
                Dim fullBox = New OcrBoundingBox(
                    box.Left + leftX,
                    box.Top + topY,
                    box.Left + rightX,
                    box.Top + botY)

                ' 0.05A → 0.05 A，下游按空格切基准字母
                Dim value = GluedDatum.Replace(txt, "$1 $2")
                Dim newText = gdtName & " " & value
                Dim oldText = If(el.Text, "").Trim()
                If oldText <> newText AndAlso Not HasGdtName(oldText) Then
                    SyncTwinText(doc, targets, el, newText)
                    el.Text = newText
                    ReplaceUniqueText(doc, oldText, newText)
                End If
                ' 文字框扩到整框，面板里能看出这条标注占的范围
                el.BoundingBox = fullBox
            Next
        End Sub

        ''' <summary>合并间距 ≤ gap 的分隔线（取均值）</summary>
        Private Shared Function MergeCloseDividers(dividers As List(Of Integer), gap As Integer) As List(Of Integer)
            Dim result As New List(Of Integer)()
            If dividers Is Nothing OrElse dividers.Count = 0 Then Return result
            Dim ordered = dividers.OrderBy(Function(x) x).ToList()
            Dim groupStart = ordered(0)
            Dim groupSum = ordered(0)
            Dim groupCount = 1
            For i = 1 To ordered.Count - 1
                If ordered(i) - ordered(i - 1) <= gap Then
                    groupSum += ordered(i)
                    groupCount += 1
                Else
                    result.Add(groupSum \ groupCount)
                    groupStart = ordered(i)
                    groupSum = ordered(i)
                    groupCount = 1
                End If
            Next
            result.Add(groupSum \ groupCount)
            Return result
        End Function

        ''' <summary>为 FCF 框搜索扩展文字框（四周各 pad padding 像素）</summary>
        Private Shared Function ExpandFcfSearchBox(bmp As Bitmap, box As OcrBoundingBox, padding As Integer) As OcrBoundingBox
            If bmp Is Nothing OrElse box Is Nothing Then Return box
            Dim l = Math.Max(0, box.Left - padding)
            Dim t = Math.Max(0, box.Top - padding)
            Dim r = Math.Min(bmp.Width - 1, box.Right + padding)
            Dim b = Math.Min(bmp.Height - 1, box.Bottom + padding)
            If r - l < 12 OrElse b - t < 12 Then Return Nothing
            Return New OcrBoundingBox(l, t, r, b)
        End Function

        ''' <summary>
        ''' 推断最左侧格子里的形位公差，返回**中文名**（解析脚本示例同款：位置度/垂直度/平行度…）。
        ''' 先对格子进行 2px 内缩以去掉框线墨迹，然后统计水平/垂直 run、对角线像素、中心密度等特征，
        ''' 按固定优先级依次判定。认不出返回空串，调用方不写这条，避免 Unknown 污染文本。
        ''' </summary>
        Private Shared Function InferFcfSymbol(ink As Boolean(,), x0 As Integer, y0 As Integer, x1 As Integer, y1 As Integer) As String
            If ink Is Nothing Then Return ""
            Dim W = ink.GetLength(0)
            Dim H = ink.GetLength(1)

            ' 2px padding inward to avoid frame border
            Dim cx0 = Math.Max(0, x0 + 2)
            Dim cy0 = Math.Max(0, y0 + 2)
            Dim cx1 = Math.Min(W - 1, x1 - 2)
            Dim cy1 = Math.Min(H - 1, y1 - 2)
            If cx1 - cx0 < 3 OrElse cy1 - cy0 < 3 Then Return ""

            Dim cellW = cx1 - cx0 + 1
            Dim cellH = cy1 - cy0 + 1

            ' --- 水平 ink run 统计 ---
            ' 每行：True 像素数 > 25% cellW 算该行有水平墨迹
            Dim hRowFlags As New List(Of Integer)()  ' 含墨迹的行号（相对 cell）
            For y = cy0 To cy1
                Dim cnt = 0
                For x = cx0 To cx1
                    If ink(x, y) Then cnt += 1
                Next
                If cnt > cellW * 0.25 Then hRowFlags.Add(y - cy0)
            Next
            ' 合并连续行成 run
            Dim hRunCount = 0
            Dim hMaxGap = 0
            If hRowFlags.Count > 0 Then
                Dim prev = hRowFlags(0)
                hRunCount = 1
                Dim curGap = 0
                For i = 1 To hRowFlags.Count - 1
                    Dim diff = hRowFlags(i) - prev - 1
                    If diff > 0 Then
                        hRunCount += 1
                        If curGap > hMaxGap Then hMaxGap = curGap
                        curGap = diff
                    End If
                    prev = hRowFlags(i)
                Next
                If curGap > hMaxGap Then hMaxGap = curGap
            End If

            ' --- 垂直 ink run 统计 ---
            Dim vColFlags As New List(Of Integer)()
            For x = cx0 To cx1
                Dim cnt = 0
                For y = cy0 To cy1
                    If ink(x, y) Then cnt += 1
                Next
                If cnt > cellH * 0.25 Then vColFlags.Add(x - cx0)
            Next
            Dim vRunCount = 0
            Dim vMaxGap = 0
            If vColFlags.Count > 0 Then
                Dim prev = vColFlags(0)
                vRunCount = 1
                Dim curGap = 0
                For i = 1 To vColFlags.Count - 1
                    Dim diff = vColFlags(i) - prev - 1
                    If diff > 0 Then
                        vRunCount += 1
                        If curGap > vMaxGap Then vMaxGap = curGap
                        curGap = diff
                    End If
                    prev = vColFlags(i)
                Next
                If curGap > vMaxGap Then vMaxGap = curGap
            End If

            ' --- 对角线像素计数 ---
            Dim diagonalCount = 0
            Dim totalInk = 0
            For y = cy0 To cy1
                For x = cx0 To cx1
                    If ink(x, y) Then
                        totalInk += 1
                        Dim hasDiag = False
                        If x - 1 >= 0 AndAlso y - 1 >= 0 AndAlso ink(x - 1, y - 1) Then hasDiag = True
                        If x + 1 < W AndAlso y - 1 >= 0 AndAlso ink(x + 1, y - 1) Then hasDiag = True
                        If x - 1 >= 0 AndAlso y + 1 < H AndAlso ink(x - 1, y + 1) Then hasDiag = True
                        If x + 1 < W AndAlso y + 1 < H AndAlso ink(x + 1, y + 1) Then hasDiag = True
                        If hasDiag Then diagonalCount += 1
                    End If
                Next
            Next

            ' --- 中心 50% 区域密度 ---
            Dim cx0C = cx0 + cellW \ 4
            Dim cy0C = cy0 + cellH \ 4
            Dim cx1C = cx1 - cellW \ 4
            Dim cy1C = cy1 - cellH \ 4
            Dim centerTotal As Integer = 0
            Dim centerHits As Integer = 0
            If cx1C > cx0C AndAlso cy1C > cy0C Then
                For y = cy0C To cy1C
                    For x = cx0C To cx1C
                        centerTotal += 1
                        If ink(x, y) Then centerHits += 1
                    Next
                Next
            End If
            Dim centerDensity As Double = If(centerTotal > 0, centerHits / CDbl(centerTotal), 0)
            Dim inkDensity As Double = If(cellW * cellH > 0, totalInk / CDbl(cellW * cellH), 0)

            ' ============ 判定规则（严格按优先级） ============
            ' 直接返回中文形位名：这条文字会替换原 OCR 元素进结果，下游解析脚本的示例就是
            ' 「位置度 0.05 A B」这种写法（中文名 + 值 + 基准）。
            ' 旧写法返回 "∥ Parallelism"、"Unknown"，脚本一条都解不出来，所以认不出就返回空串。
            ' 值里带 Ø 没关系：脚本的中文名分支本来就认「位置度 Ø0.05 A」。

            ' 1. 平行度：多条横线，很少竖线，横线间有明显间距
            If hRunCount >= 2 AndAlso vRunCount <= 1 AndAlso hMaxGap > cellH * 0.15 Then
                Return "平行度"
            End If

            ' 2. 圆跳动：斜箭头，对角线比横竖多。要排在垂直度前面。
            If diagonalCount > 20 AndAlso hRunCount <= 2 AndAlso vRunCount <= 2 Then
                Return "圆跳动"
            End If

            ' 3. 位置度 / 垂直度（十字形）：横竖都多，中心密度高
            If vRunCount >= 2 AndAlso hRunCount >= 2 AndAlso centerDensity > 0.1 AndAlso diagonalCount < 20 Then
                If vMaxGap < 5 Then
                    Return "垂直度"
                Else
                    Return "位置度"
                End If
            End If

            ' 4. 垂直度：一条竖线加横线，而且对角线很少。斜箭头不算。
            If vRunCount = 1 AndAlso hRunCount >= 1 AndAlso diagonalCount < 12 Then
                Return "垂直度"
            End If

            ' 垂直度（额外竖线可能是分隔线）：多条竖线但对角线很少。斜箭头的对角线更多。
            If vRunCount >= 2 AndAlso hRunCount >= 1 AndAlso diagonalCount < 12 Then
                Return "垂直度"
            End If

            ' 5. 跳动 / 轮廓度：对角线多，横竖少
            If diagonalCount > 50 AndAlso hRunCount <= 1 AndAlso vRunCount <= 1 Then
                Return "圆跳动"
            End If

            ' 6. 轮廓度（上拱/下凹）：对角线中等。线/面轮廓度在这个粒度上分不出来，按面轮廓度出。
            If diagonalCount > 20 AndAlso vRunCount <= 1 AndAlso hRunCount <= 1 Then
                Return "面轮廓度"
            End If

            ' 7. 同轴度 / 对称度：多条竖线且间距大
            If vRunCount >= 2 AndAlso vMaxGap > cellW * 0.3 Then
                Return "同轴度"
            End If

            ' 8. 空框 / 认不出：都不写进文本
            Return ""
        End Function

#End Region

    End Class

End Namespace
