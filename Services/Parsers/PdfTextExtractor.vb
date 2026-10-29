Option Strict On
Option Explicit On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Text
Imports System.Diagnostics
Imports PdfSharp.Pdf
Imports PdfSharp.Pdf.Advanced
Imports PdfSharp.Pdf.Content
Imports PdfSharp.Pdf.Content.Objects

Namespace Services.Parsers

    ''' <summary>
    ''' PDF 内容流文本提取器（PdfSharp ContentReader 实现）
    ''' 
    ''' 背景：旧实现靠 page.Contents.Elements 遍历后对 ToString() 抠括号取文本，
    ''' 对真实 PDF（内容流为压缩字节、字符串是 Tj/TJ 算子操作数）几乎取不到任何文字，
    ''' 导致所有 PDF 都被判定为"扫描版"而拒绝解析。
    ''' 
    ''' 本提取器走 PdfSharp 自带的 ContentReader 语法树：
    '''   1. 从页面 Resources 构建字体映射（Type0/简单字体、Encoding、ToUnicode CMap）
    '''   2. 遍历内容流操作符，跟踪 Tf 字体选择与 Td/Tm/TD/T* 文本定位
    '''   3. 对 Tj/TJ/'/" 的字符串操作数按当前字体解码（ToUnicode → UTF-16BE 兜底 → 编码表）
    '''   4. 纵向位移超过阈值时换行，保留图纸/文档的行结构
    '''   5. 递归进入 Form XObject（图纸标题栏常藏在 Form 里），深度上限防爆炸
    ''' 
    ''' 自包含设计：只依赖 PdfSharp + BCL，可被 vbc 单独编译做样本验证。
    ''' </summary>
    Public Class PdfTextExtractor

        ' ==================== 常量 ====================

        ''' <summary>Form XObject 递归深度上限（防畸形 PDF 无限递归）</summary>
        Private Const MAX_XOBJECT_DEPTH As Integer = 4

        ''' <summary>TJ 阵列中负向 kerning 小于等于该值（千分比 em）时补一个空格</summary>
        Private Const TJ_GAP_THRESHOLD As Double = -120.0

        ''' <summary>纵向位移超过 字号*该比例 视为换行</summary>
        Private Const LINE_SHIFT_RATIO As Double = 0.35

        ' ==================== 字体信息 ====================

        ''' <summary>单个字体资源的解码信息</summary>
        Private Class FontInfo
            ''' <summary>Type0/CID 字体（字符串按双字节 CID 解读）</summary>
            Public Property IsCidFont As Boolean = False
            ''' <summary>Encoding 名称（WinAnsiEncoding/MacRomanEncoding/Identity-H/StandardEncoding…）</summary>
            Public Property EncodingName As String = ""
            ''' <summary>简单字体 + 自定义 Encoding 字典时的字节→字形名映射（/Differences）</summary>
            Public Property Differences As Dictionary(Of Integer, String) = Nothing
            ''' <summary>ToUnicode CMap：码位 → Unicode 文本</summary>
            Public Property ToUnicode As Dictionary(Of UInteger, String) = Nothing
            Public ReadOnly Property HasToUnicode As Boolean
                Get
                    Return ToUnicode IsNot Nothing AndAlso ToUnicode.Count > 0
                End Get
            End Property
        End Class

        ' ==================== 文本遍历状态 ====================

        ''' <summary>内容流遍历过程中的文本定位状态</summary>
        Private Class TextState
            Private ReadOnly _sb As New StringBuilder(4096)
            Private _lastY As Double = Double.NaN
            Private _lastSize As Double = 0.0
            Private _started As Boolean = False

            Public Property X As Double = 0.0
            Public Property Y As Double = 0.0
            Public Property Leading As Double = 0.0
            Public Property CurrentFont As FontInfo = Nothing
            Public Property CurrentFontSize As Double = 0.0

            ''' <summary>输出一段文本（按纵向位移决定是否换行）</summary>
            Public Sub AppendText(text As String)
                If String.IsNullOrEmpty(text) Then Return
                If _started Then
                    Dim threshold As Double = Math.Max(0.5, Math.Abs(_lastSize) * LINE_SHIFT_RATIO)
                    If Not Double.IsNaN(_lastY) AndAlso Math.Abs(Me.Y - _lastY) > threshold Then
                        _sb.AppendLine()
                    End If
                End If
                _sb.Append(text)
                _lastY = Me.Y
                _lastSize = If(CurrentFontSize > 0.0, CurrentFontSize, 10.0)
                _started = True
            End Sub

            ''' <summary>强制换行（如 BT 边界）</summary>
            Public Sub AppendLineBreak()
                If _started Then
                    _sb.AppendLine()
                    _lastY = Double.NaN
                End If
            End Sub

            Public ReadOnly Property Length As Integer
                Get
                    Return _sb.Length
                End Get
            End Property

            Public Overrides Function ToString() As String
                Return _sb.ToString()
            End Function
        End Class

        ' ==================== 公共入口 ====================

        ''' <summary>
        ''' 提取单页 PDF 的文本（含 Form XObject 内的文本）。
        ''' 任何异常都吞掉返回空串——由调用方按"无文本"走扫描件/OCR 兜底，
        ''' 绝不让单页解析失败炸掉整个文档解析。
        ''' </summary>
        Public Shared Function ExtractPageText(page As PdfPage) As String
            If page Is Nothing Then Return String.Empty
            Try
                Dim state As New TextState()
                Dim fonts As Dictionary(Of String, FontInfo) = BuildFontMap(page.Resources)
                Dim seq As CSequence = ContentReader.ReadContent(page)
                WalkSequence(seq, fonts, page, 0, state)
                Return CleanupText(state.ToString())
            Catch ex As Exception
                Debug.WriteLine("[PdfTextExtractor] 单页提取异常: " & ex.Message)
                Return String.Empty
            End Try
        End Function

        ' ==================== 内容流遍历 ====================

        ''' <summary>递归遍历内容流语法树，抽取文本算子</summary>
        Private Shared Sub WalkSequence(seq As CSequence,
                                        fonts As Dictionary(Of String, FontInfo),
                                        page As PdfPage,
                                        depth As Integer,
                                        state As TextState)
            If seq Is Nothing OrElse depth > MAX_XOBJECT_DEPTH Then Return

            For Each item As CObject In seq
                Dim op As COperator = TryCast(item, COperator)
                If op Is Nothing Then Continue For

                Try
                    Select Case op.Name
                        Case "BT"
                            state.AppendLineBreak()
                            state.CurrentFont = Nothing
                            state.X = 0.0
                            state.Y = 0.0
                            state.Leading = 0.0

                        Case "ET"
                            state.AppendLineBreak()

                        Case "Tf"
                            If op.Operands.Count >= 2 Then
                                Dim fontName As String = OperandName(op.Operands(0))
                                If fontName IsNot Nothing AndAlso fonts.ContainsKey(fontName) Then
                                    state.CurrentFont = fonts(fontName)
                                Else
                                    state.CurrentFont = Nothing
                                End If
                                state.CurrentFontSize = OperandNumber(op.Operands(1))
                            End If

                        Case "Td", "TD"
                            If op.Operands.Count >= 2 Then
                                state.X += OperandNumber(op.Operands(0))
                                state.Y += OperandNumber(op.Operands(1))
                                If op.Name = "TD" Then state.Leading = -OperandNumber(op.Operands(1))
                            End If

                        Case "Tm"
                            If op.Operands.Count >= 6 Then
                                state.X = OperandNumber(op.Operands(4))
                                state.Y = OperandNumber(op.Operands(5))
                            End If

                        Case "TL"
                            If op.Operands.Count >= 1 Then state.Leading = OperandNumber(op.Operands(0))

                        Case "T*"
                            state.Y -= state.Leading

                        Case "Tj", "'", "QuoteSingle"
                            If op.Name <> "Tj" Then state.Y -= state.Leading
                            Dim s As CString = FirstStringOperand(op)
                            If s IsNot Nothing Then ShowText(s.Value, state)

                        Case """", "QuoteDbl"
                            state.Y -= state.Leading
                            If op.Operands.Count >= 3 Then
                                Dim s As CString = TryCast(op.Operands(2), CString)
                                If s IsNot Nothing Then ShowText(s.Value, state)
                            End If

                        Case "TJ"
                            If op.Operands.Count >= 1 Then
                                Dim arr As CArray = TryCast(op.Operands(0), CArray)
                                If arr IsNot Nothing Then
                                    For Each el As CObject In arr
                                        Dim cs As CString = TryCast(el, CString)
                                        If cs IsNot Nothing Then
                                            ShowText(cs.Value, state)
                                        Else
                                            Dim num As CObject = TryCast(el, CNumber)
                                            If num IsNot Nothing Then
                                                Dim nv As Double = OperandNumber(el)
                                                If nv <= TJ_GAP_THRESHOLD Then
                                                    ' 明显的字间空隙 → 补空格（TJ kerning 为负表示右移）
                                                    ShowText(" ", state)
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If

                        Case "Do"
                            WalkXObject(op, fonts, page, depth, state)

                    End Select
                Catch exOp As Exception
                    ' 单算子异常不影响整页
                    Debug.WriteLine("[PdfTextExtractor] 算子 " & op.Name & " 处理异常: " & exOp.Message)
                End Try
            Next
        End Sub

        ''' <summary>Do 算子：递归进入 Form XObject 的内容流</summary>
        Private Shared Sub WalkXObject(op As COperator,
                                       fonts As Dictionary(Of String, FontInfo),
                                       page As PdfPage,
                                       depth As Integer,
                                       state As TextState)
            Try
                If depth >= MAX_XOBJECT_DEPTH Then Return
                Dim name As String = OperandName(op.Operands(0))
                If String.IsNullOrEmpty(name) OrElse page Is Nothing OrElse page.Resources Is Nothing Then Return

                Dim xobjects As PdfDictionary = page.Resources.Elements.GetDictionary("/XObject")
                If xobjects Is Nothing Then Return

                Dim target As PdfDictionary = Nothing
                For Each kv In xobjects.Elements
                    If String.Equals(kv.Key.TrimStart("/"c), name, StringComparison.OrdinalIgnoreCase) Then
                        target = ResolveDictionary(kv.Value)
                        Exit For
                    End If
                Next
                If target Is Nothing Then Return

                Dim subtype As String = target.Elements.GetName("/Subtype")
                If subtype Is Nothing OrElse Not subtype.EndsWith("Form", StringComparison.OrdinalIgnoreCase) Then Return

                Dim stm As PdfDictionary.PdfStream = target.Stream
                If stm Is Nothing Then Return
                Dim raw As Byte() = stm.UnfilteredValue
                If raw Is Nothing OrElse raw.Length = 0 Then Return

                ' Form 自带 Resources 时合并进字体映射（表单级字体优先）
                Dim merged As Dictionary(Of String, FontInfo) = fonts
                Dim formRes As PdfResources = TryCast(target.Elements.GetDictionary("/Resources"), PdfResources)
                If formRes IsNot Nothing Then
                    Dim formFonts As Dictionary(Of String, FontInfo) = BuildFontMap(formRes)
                    If formFonts.Count > 0 Then
                        merged = New Dictionary(Of String, FontInfo)(fonts, StringComparer.OrdinalIgnoreCase)
                        For Each kv In formFonts
                            merged(kv.Key) = kv.Value
                        Next
                    End If
                End If

                Dim formSeq As CSequence = ContentReader.ReadContent(raw)
                WalkSequence(formSeq, merged, page, depth + 1, state)
            Catch ex As Exception
                Debug.WriteLine("[PdfTextExtractor] XObject 递归异常: " & ex.Message)
            End Try
        End Sub

        ''' <summary>输出文本（当前字体解码）</summary>
        Private Shared Sub ShowText(rawValue As String, state As TextState)
            If String.IsNullOrEmpty(rawValue) Then Return
            Dim decoded As String = DecodeString(rawValue, state.CurrentFont)
            If Not String.IsNullOrEmpty(decoded) Then state.AppendText(decoded)
        End Sub

        ' ==================== 字符串解码 ====================

        ''' <summary>
        ''' 把内容流字符串（PdfSharp 已处理括号/八进制转义，char=字节值 0-255）
        ''' 按当前字体解码为 Unicode 文本。
        ''' </summary>
        Private Shared Function DecodeString(value As String, font As FontInfo) As String
            If String.IsNullOrEmpty(value) Then Return String.Empty
            If font Is Nothing Then Return SanitizeAscii(value)

            Dim sb As New StringBuilder(value.Length)

            If font.IsCidFont Then
                ' Type0/Identity：双字节 CID
                Dim i As Integer = 0
                While i + 1 < value.Length
                    Dim code As UInteger = CUInt((AscW(value(i)) << 8) Or AscW(value(i + 1)))
                    sb.Append(MapCode(code, font))
                    i += 2
                End While
                ' 末尾落单字节按 Latin-1 兜底
                If i < value.Length Then sb.Append(MapCode(CUInt(AscW(value(i))), font))
            Else
                For Each ch As Char In value
                    sb.Append(MapCode(CUInt(AscW(ch)), font))
                Next
            End If

            Return sb.ToString()
        End Function

        ''' <summary>单码位 → Unicode（ToUnicode 优先，其次编码表/Differences）</summary>
        Private Shared Function MapCode(code As UInteger, font As FontInfo) As String
            If font.HasToUnicode AndAlso font.ToUnicode.ContainsKey(code) Then
                Dim mapped As String = font.ToUnicode(code)
                ' CAD 图纸的 ToUnicode CMap 常把未映射字形字面写成 <FFFD>，直接吐给模型是噪声，
                ' 且 U+FFFD 是文化敏感比较里的可忽略字符（见 CleanupText 历史坑注释），统一过滤。
                If String.Equals(mapped, ChrW(&HFFFD).ToString(), StringComparison.Ordinal) Then
                    Return String.Empty
                End If
                Return mapped
            End If

            If font.IsCidFont Then
                ' CID 字体无 ToUnicode：尝试把双字节直接当 UTF-16BE（常见兜底，能救一部分 CAD PDF）
                If code >= 32 AndAlso code <> &HFFFE AndAlso code <> &HFFFF Then
                    ' 过滤 U+FFFD（替换字符）与代理区（D800-DFFF）：
                    ' CAD 图纸按 CID→UTF-16BE 兜底时会大量产出 U+FFFD（实测真实图纸），
                    ' 而 U+FFFD 在 Windows 文化敏感字符串比较中是"可忽略字符"，
                    ' 会让 CleanupText 的IndexOf("\n\n\n")跨过它误报匹配、与序号比较的Replace
                    ' 永远不一致 → 死循环（2026-10-18 真实 PLIER 图纸事故）。
                    If code <> &HFFFDUI AndAlso (code < &HD800UI OrElse code > &HDFFFUI) Then
                        Return SanitizeAscii(ChrW(CInt(code)))
                    End If
                End If
                Return String.Empty
            End If

            ' 简单字体：自定义 Differences → 字形名
            If font.Differences IsNot Nothing AndAlso font.Differences.Count > 0 Then
                Dim b As Integer = CInt(code And &HFFUI)
                If font.Differences.ContainsKey(b) Then
                    Dim g As String = GlyphNameToUnicode(font.Differences(b))
                    If Not String.IsNullOrEmpty(g) Then Return g
                End If
            End If

            ' 标准编码表
            Return MapByteToUnicode(CByte(code And &HFFUI), font.EncodingName)
        End Function

        ''' <summary>字节 → Unicode（WinAnsi=CP1252 / MacRoman=CP10000 / 其余按 Latin-1）</summary>
        Private Shared Function MapByteToUnicode(b As Byte, encodingName As String) As String
            Try
                Dim cp As Integer = 28591 ' ISO-8859-1 默认
                Dim enc As String = If(encodingName, String.Empty)
                If enc.IndexOf("WinAnsi", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cp = 1252
                ElseIf enc.IndexOf("MacRoman", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cp = 10000
                ElseIf enc.IndexOf("Standard", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                       enc.IndexOf("PDFDoc", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    cp = 28591
                End If
                Dim bytes As Byte() = New Byte() {b}
                Return Encoding.GetEncoding(cp).GetString(bytes)
            Catch
                If b >= 32 AndAlso b < 127 Then Return ChrW(CInt(b))
                Return String.Empty
            End Try
        End Function

        ''' <summary>无字体信息时的兜底：只保留可打印 ASCII，其余丢弃（避免乱码进上下文）</summary>
        Private Shared Function SanitizeAscii(value As String) As String
            Dim sb As New StringBuilder(value.Length)
            For Each ch As Char In value
                Dim c As Integer = AscW(ch)
                If (c >= 32 AndAlso c < 127) OrElse c = 10 OrElse c = 13 OrElse c = 9 Then
                    sb.Append(ch)
                ElseIf c > 255 Then
                    sb.Append(ch) ' 高位直接保留（UTF-16 兜底场景）
                End If
            Next
            Return sb.ToString()
        End Function

        ' ==================== 字体映射构建 ====================

        ''' <summary>从 Resources 构建 字体名 → FontInfo</summary>
        Private Shared Function BuildFontMap(resources As PdfResources) As Dictionary(Of String, FontInfo)
            Dim map As New Dictionary(Of String, FontInfo)(StringComparer.OrdinalIgnoreCase)
            If resources Is Nothing OrElse resources.Elements Is Nothing Then Return map

            Try
                Dim fontDict As PdfDictionary = resources.Elements.GetDictionary("/Font")
                If fontDict Is Nothing OrElse fontDict.Elements Is Nothing Then Return map

                For Each kv In fontDict.Elements
                    Try
                        Dim key As String = If(kv.Key, String.Empty).TrimStart("/"c)
                        If String.IsNullOrEmpty(key) Then Continue For
                        Dim fd As PdfDictionary = ResolveDictionary(kv.Value)
                        If fd Is Nothing OrElse fd.Elements Is Nothing Then Continue For
                        map(key) = BuildFontInfo(fd)
                    Catch exOne As Exception
                        Debug.WriteLine("[PdfTextExtractor] 单字体构建失败: " & exOne.Message)
                    End Try
                Next
            Catch ex As Exception
                Debug.WriteLine("[PdfTextExtractor] 字体映射构建异常: " & ex.Message)
            End Try

            Return map
        End Function

        ''' <summary>解析单个字体字典</summary>
        Private Shared Function BuildFontInfo(fd As PdfDictionary) As FontInfo
            Dim fi As New FontInfo()

            Dim subtype As String = fd.Elements.GetString("/Subtype")
            If subtype IsNot Nothing AndAlso subtype.TrimStart("/"c).Equals("Type0", StringComparison.OrdinalIgnoreCase) Then
                fi.IsCidFont = True
            End If

            ' Encoding：可能是名称，也可能是带 /BaseEncoding + /Differences 的字典
            Dim encItem As PdfItem = fd.Elements("/Encoding")
            Dim encName As String = fd.Elements.GetName("/Encoding")
            If encName IsNot Nothing Then
                fi.EncodingName = encName
            Else
                Dim encDict As PdfDictionary = TryCast(encItem, PdfDictionary)
                If encDict Is Nothing Then
                    Dim encRef As PdfReference = TryCast(encItem, PdfReference)
                    If encRef IsNot Nothing Then encDict = TryCast(encRef.Value, PdfDictionary)
                End If
                If encDict IsNot Nothing AndAlso encDict.Elements IsNot Nothing Then
                    Dim baseEnc As String = encDict.Elements.GetName("/BaseEncoding")
                    If baseEnc IsNot Nothing Then fi.EncodingName = baseEnc
                    fi.Differences = ParseDifferences(encDict)
                End If
            End If

            ' ToUnicode CMap
            Try
                Dim tu As PdfDictionary = fd.Elements.GetDictionary("/ToUnicode")
                If tu IsNot Nothing Then
                    Dim stm As PdfDictionary.PdfStream = tu.Stream
                    If stm IsNot Nothing Then
                        Dim raw As Byte() = stm.UnfilteredValue
                        If raw Is Nothing OrElse raw.Length = 0 Then raw = stm.Value
                        If raw IsNot Nothing AndAlso raw.Length > 0 Then
                            fi.ToUnicode = ParseToUnicodeCMap(raw)
                        End If
                    End If
                End If
            Catch exTu As Exception
                Debug.WriteLine("[PdfTextExtractor] ToUnicode 解析失败: " & exTu.Message)
            End Try

            Return fi
        End Function

        ''' <summary>解析 /Encoding 字典的 /Differences 数组 → 字节码→字形名</summary>
        Private Shared Function ParseDifferences(encDict As PdfDictionary) As Dictionary(Of Integer, String)
            Dim map As New Dictionary(Of Integer, String)()
            Try
                Dim diffs As PdfArray = encDict.Elements.GetArray("/Differences")
                If diffs Is Nothing Then Return map
                Dim currentCode As Integer = 0
                For Each el As PdfItem In diffs
                    Dim num As PdfInteger = TryCast(el, PdfInteger)
                    If num IsNot Nothing Then
                        currentCode = num.Value
                    Else
                        Dim nameItem As PdfName = TryCast(el, PdfName)
                        If nameItem IsNot Nothing Then
                            Dim glyphName As String = If(nameItem.Value, String.Empty).TrimStart("/"c)
                            If currentCode >= 0 AndAlso currentCode <= 255 Then
                                map(currentCode) = glyphName
                            End If
                            currentCode += 1
                        End If
                    End If
                Next
            Catch ex As Exception
                Debug.WriteLine("[PdfTextExtractor] Differences 解析异常: " & ex.Message)
            End Try
            Return map
        End Function

        ' ==================== ToUnicode CMap 解析 ====================

        ''' <summary>
        ''' 解析 ToUnicode CMap（bfchar/bfrange），返回 码位 → Unicode 文本。
        ''' CMap 文本已由 PdfSharp 解压（UnfilteredValue），按 Latin-1 读成 char=byte。
        ''' </summary>
        Private Shared Function ParseToUnicodeCMap(data As Byte()) As Dictionary(Of UInteger, String)
            Dim map As New Dictionary(Of UInteger, String)()
            If data Is Nothing OrElse data.Length = 0 Then Return map

            Dim tokens As List(Of String) = Nothing
            Try
                tokens = TokenizeCMap(data)
            Catch ex As Exception
                Debug.WriteLine("[PdfTextExtractor] CMap 分词异常: " & ex.Message)
                Return map
            End Try
            If tokens Is Nothing OrElse tokens.Count = 0 Then Return map

            Dim i As Integer = 0
            While i < tokens.Count
                Dim t As String = tokens(i)
                If t = "beginbfchar" Then
                    i += 1
                    While i + 1 < tokens.Count AndAlso tokens(i) <> "endbfchar"
                        If tokens(i).StartsWith("H:") AndAlso tokens(i + 1).StartsWith("H:") Then
                            Dim src As UInteger = CUInt(ParseHex(tokens(i).Substring(2)))
                            Dim dst As String = HexToUnicode(tokens(i + 1).Substring(2))
                            If Not map.ContainsKey(src) Then map(src) = dst
                            i += 2
                        Else
                            i += 1
                        End If
                    End While
                    i += 1 ' 跳过 endbfchar
                ElseIf t = "beginbfrange" Then
                    i += 1
                    While i < tokens.Count AndAlso tokens(i) <> "endbfrange"
                        ' 两种形态：lo hi dst 或 lo hi [dst1 dst2 ...]
                        If i + 2 < tokens.Count AndAlso tokens(i).StartsWith("H:") AndAlso
                           tokens(i + 1).StartsWith("H:") Then
                            Dim lo As UInteger = CUInt(ParseHex(tokens(i).Substring(2)))
                            Dim hi As UInteger = CUInt(ParseHex(tokens(i + 1).Substring(2)))
                            If tokens(i + 2).StartsWith("H:") Then
                                Dim dstStart As UInteger = ParseHex(tokens(i + 2).Substring(2))
                                If hi >= lo AndAlso hi - lo < 65536 Then
                                    For c As UInteger = lo To hi
                                        If Not map.ContainsKey(c) Then
                                            map(c) = HexToUnicode(FormatHex(dstStart + (c - lo)))
                                        End If
                                    Next
                                End If
                                i += 3
                            ElseIf tokens(i + 2) = "[" Then
                                ' 数组形态
                                Dim j As Integer = i + 3
                                Dim offset As UInteger = 0
                                While j < tokens.Count AndAlso tokens(j) <> "]"
                                    If tokens(j).StartsWith("H:") Then
                                        Dim code As UInteger = lo + offset
                                        If Not map.ContainsKey(code) Then
                                            map(code) = HexToUnicode(tokens(j).Substring(2))
                                        End If
                                        offset += 1UI
                                    End If
                                    j += 1
                                End While
                                i = j + 1
                            Else
                                i += 1
                            End If
                        Else
                            i += 1
                        End If
                    End While
                    i += 1 ' 跳过 endbfrange
                Else
                    i += 1
                End If
            End While

            Return map
        End Function

        ''' <summary>把 CMap 原文切成 token：十六进制串记为 "H:xxxx"，其余为裸词</summary>
        Private Shared Function TokenizeCMap(data As Byte()) As List(Of String)
            Dim tokens As New List(Of String)()
            Dim text As String = Encoding.GetEncoding(28591).GetString(data)
            Dim i As Integer = 0
            ' 防御性迭代上限：任何分支都必须保证 i 前进，此处兜底杜绝死循环
            Dim guard As Integer = text.Length + 16
            While i < text.Length AndAlso guard > 0
                guard -= 1
                Dim ch As Char = text(i)
                If ch = "<"c Then
                    If i + 1 < text.Length AndAlso text(i + 1) = "<"c Then
                        tokens.Add("<<") ' CMap 内嵌字典起始（如 /CIDSystemInfo << ... >>）
                        i += 2
                    Else
                        Dim j As Integer = text.IndexOf(">"c, i)
                        If j < 0 Then Exit While
                        tokens.Add("H:" & text.Substring(i + 1, j - i - 1).Trim())
                        i = j + 1
                    End If
                ElseIf ch = ">"c Then
                    If i + 1 < text.Length AndAlso text(i + 1) = ">"c Then
                        tokens.Add(">>") ' CMap 内嵌字典结束
                        i += 2
                    Else
                        tokens.Add(">")
                        i += 1
                    End If
                ElseIf ch = "%"c Then
                    ' 注释到行尾
                    Dim j As Integer = text.IndexOf(vbLf, i)
                    If j < 0 Then Exit While
                    i = j + 1
                ElseIf Char.IsWhiteSpace(ch) Then
                    i += 1
                Else
                    Dim j As Integer = i
                    While j < text.Length AndAlso Not Char.IsWhiteSpace(text(j)) AndAlso text(j) <> "<"c AndAlso text(j) <> ">"c
                        j += 1
                    End While
                    If j <= i Then
                        ' 理论上不可达（< 与 > 已由上方分支处理）；防御性强制前进
                        tokens.Add(ch.ToString())
                        i += 1
                    Else
                        Dim w As String = text.Substring(i, j - i)
                        If w.Length > 0 Then tokens.Add(w)
                        i = j
                    End If
                End If
            End While
            Return tokens
        End Function

        ''' <summary>CMap 十六进制串 → Unicode 文本（UTF-16BE；4 字节且前导零按 UTF-32 取尾部）</summary>
        Private Shared Function HexToUnicode(hex As String) As String
            If String.IsNullOrEmpty(hex) Then Return String.Empty
            Dim clean As String = New String(hex.Where(Function(c) Not Char.IsWhiteSpace(c)).ToArray())
            If clean.Length = 0 OrElse clean.Length Mod 2 <> 0 Then Return String.Empty

            Try
                Dim byteCount As Integer = clean.Length \ 2
                Dim bytes As Byte() = New Byte(byteCount - 1) {}
                For k As Integer = 0 To byteCount - 1
                    bytes(k) = Byte.Parse(clean.Substring(k * 2, 2), Globalization.NumberStyles.HexNumber)
                Next

                ' 4 字节形态：UTF-32（前两字节为零）或 UTF-16 代理对
                If byteCount = 4 AndAlso bytes(0) = 0 AndAlso bytes(1) = 0 Then
                    Return Encoding.BigEndianUnicode.GetString(New Byte() {bytes(2), bytes(3)})
                End If

                Return Encoding.BigEndianUnicode.GetString(bytes)
            Catch
                Return String.Empty
            End Try
        End Function

        ''' <summary>码位 → CMap 十六进制串（bfrange 递增值用）</summary>
        Private Shared Function FormatHex(code As UInteger) As String
            If code <= &HFFFFUI Then
                Return code.ToString("X4")
            End If
            Return code.ToString("X8")
        End Function

        Private Shared Function ParseHex(hex As String) As UInteger
            If String.IsNullOrEmpty(hex) Then Return 0UI
            Try
                Return Convert.ToUInt32(hex, 16)
            Catch
                Return 0UI
            End Try
        End Function

        ' ==================== 字形名 → Unicode（/Differences 兜底） ====================

        ''' <summary>
        ''' 常见字形名 → Unicode（覆盖 ASCII + 工程图纸常见符号 + uniXXXX）。
        ''' 未覆盖的字形名返回空串（调用方继续走编码表）。
        ''' </summary>
        Private Shared Function GlyphNameToUnicode(name As String) As String
            If String.IsNullOrEmpty(name) Then Return String.Empty
            Select Case name
                Case "space" : Return " "
                Case "exclam" : Return "!"
                Case "quotedbl" : Return """"
                Case "numbersign" : Return "#"
                Case "dollar" : Return "$"
                Case "percent" : Return "%"
                Case "ampersand" : Return "&"
                Case "quotesingle" : Return "'"
                Case "parenleft" : Return "("
                Case "parenright" : Return ")"
                Case "asterisk" : Return "*"
                Case "plus" : Return "+"
                Case "comma" : Return ","
                Case "hyphen", "minus" : Return "-"
                Case "period" : Return "."
                Case "slash" : Return "/"
                Case "zero" : Return "0"
                Case "one" : Return "1"
                Case "two" : Return "2"
                Case "three" : Return "3"
                Case "four" : Return "4"
                Case "five" : Return "5"
                Case "six" : Return "6"
                Case "seven" : Return "7"
                Case "eight" : Return "8"
                Case "nine" : Return "9"
                Case "colon" : Return ":"
                Case "semicolon" : Return ";"
                Case "less" : Return "<"
                Case "equal" : Return "="
                Case "greater" : Return ">"
                Case "question" : Return "?"
                Case "at" : Return "@"
                Case "bracketleft" : Return "["
                Case "backslash" : Return "\"
                Case "bracketright" : Return "]"
                Case "asciicircum" : Return "^"
                Case "underscore" : Return "_"
                Case "grave" : Return "`"
                Case "braceleft" : Return "{"
                Case "bar" : Return "|"
                Case "braceright" : Return "}"
                Case "asciitilde" : Return "~"
                Case "degree" : Return ChrW(&HB0)
                Case "plusminus" : Return ChrW(&HB1)
                Case "multiply" : Return ChrW(&HD7)
                Case "divide" : Return ChrW(&HF7)
                Case "lessequal" : Return ChrW(&H2264)
                Case "greaterequal" : Return ChrW(&H2265)
                Case "notequal" : Return ChrW(&H2260)
                Case "bullet" : Return ChrW(&H2022)
                Case "ellipsis" : Return ChrW(&H2026)
                Case "endash" : Return ChrW(&H2013)
                Case "emdash" : Return ChrW(&H2014)
                Case "quotedblleft", "quotedblbase" : Return ChrW(&H201C)
                Case "quotedblright" : Return ChrW(&H201D)
                Case "diameter" : Return ChrW(&H2300)
                Case Else
                    ' uniXXXX（4位）/ uXXXXXX（6位）
                    If name.StartsWith("uni") AndAlso name.Length >= 7 Then
                        Try
                            Dim hexPart As String = name.Substring(3, 4)
                            Dim cp As Integer = Convert.ToInt32(hexPart, 16)
                            If cp >= 0 AndAlso cp <= &HFFFF Then Return ChrW(cp)
                        Catch
                            ' 忽略
                        End Try
                    End If
                    ' 单字符字形名（A-Z a-z 0-9）
                    If name.Length = 1 Then Return name
                    Return String.Empty
            End Select
        End Function

        ' ==================== 杂项 ====================

        ''' <summary>解引用：PdfReference → PdfDictionary；本身是字典则直接返回</summary>
        Private Shared Function ResolveDictionary(item As PdfItem) As PdfDictionary
            If item Is Nothing Then Return Nothing
            Dim fd As PdfDictionary = TryCast(item, PdfDictionary)
            If fd IsNot Nothing Then Return fd
            Dim ref As PdfReference = TryCast(item, PdfReference)
            If ref IsNot Nothing Then Return TryCast(ref.Value, PdfDictionary)
            Return Nothing
        End Function

        ''' <summary>取算子第一个操作数的名称（CName，去掉前导斜杠）</summary>
        Private Shared Function OperandName(operand As CObject) As String
            If operand Is Nothing Then Return Nothing
            Dim nm As CName = TryCast(operand, CName)
            If nm Is Nothing OrElse String.IsNullOrEmpty(nm.Name) Then Return Nothing
            Return nm.Name.TrimStart("/"c)
        End Function

        ''' <summary>取算子第一个字符串操作数</summary>
        Private Shared Function FirstStringOperand(op As COperator) As CString
            If op Is Nothing OrElse op.Operands Is Nothing OrElse op.Operands.Count = 0 Then Return Nothing
            Return TryCast(op.Operands(0), CString)
        End Function

        ''' <summary>操作数 → Double（CInteger/CReal/CNumber）</summary>
        Private Shared Function OperandNumber(operand As CObject) As Double
            If operand Is Nothing Then Return 0.0
            Dim ci As CInteger = TryCast(operand, CInteger)
            If ci IsNot Nothing Then Return CDbl(ci.Value)
            Dim cr As CReal = TryCast(operand, CReal)
            If cr IsNot Nothing Then Return cr.Value
            Return 0.0
        End Function

        ''' <summary>文本收尾：压缩连续空行、去首尾空白</summary>
        Private Shared Function CleanupText(text As String) As String
            If String.IsNullOrEmpty(text) Then Return String.Empty
            Dim t As String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            ' 折叠 3 个以上连续换行（线性扫描，保留最多 2 个连续 LF）。
            ' 【历史坑】旧写法是 While t.IndexOf(vbLf & vbLf & vbLf) >= 0 : t = t.Replace(...)。
            ' String.IndexOf(String) 默认走【文化敏感比较】，U+FFFD 等可忽略字符会被跳过匹配，
            ' 而 String.Replace 是【序号比较】找不到字面目标 → 字符串永不改变 → 无限循环，
            ' 真实 CAD 图纸（CID→UTF-16BE 兜底产出 U+FFFD）一碰就挂死（2026-10-18 实测）。
            ' 线性改写同时消掉旧写法 O(n²) 的性能坑（50 万字符一页可跑数万次全串 Replace）。
            Dim sb As New StringBuilder(t.Length)
            Dim lfRun As Integer = 0
            For Each ch As Char In t
                If ch = vbLf Then
                    lfRun += 1
                    If lfRun <= 2 Then sb.Append(ch)
                Else
                    lfRun = 0
                    sb.Append(ch)
                End If
            Next
            t = sb.ToString()
            ' 行尾空白
            Dim lines As String() = t.Split(New String() {vbLf}, StringSplitOptions.None)
            For k As Integer = 0 To lines.Length - 1
                lines(k) = lines(k).TrimEnd()
            Next
            Return String.Join(vbLf, lines).Trim()
        End Function

    End Class

End Namespace
