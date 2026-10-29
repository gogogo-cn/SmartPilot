Option Strict On
Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Runtime.InteropServices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Core
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class ExcelMcpTools

        ''' <summary>
        ''' excel_add_shape — 在看板工作表上绘制形状（圆角卡片 / 底色块 / 细分隔条）。
        ''' 用 range 作锚点定位尺寸（range.Left/Top/Width/Height → 磅），可选 fill/line 颜色，
        ''' 可选 lines（JSON，每行独立字号/颜色/加粗/对齐）写入形状文字。
        ''' 配合"不合并单元格"策略，用形状实现专业卡片，规避合并丢值。
        ''' </summary>
        Private Function ExecuteAddShape(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result As New ToolsCallResult()
            Dim excelApp As Microsoft.Office.Interop.Excel.Application = Globals.ThisAddIn.Application
            If excelApp Is Nothing OrElse excelApp.ActiveWorkbook Is Nothing Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "当前没有打开任何 Excel 工作簿，请先打开一个 Excel 文件。"})
                Return result
            End If

            Dim sheetName As String = GetArg(Of String)(args, "sheet", "")
            Dim rangeAddr As String = GetArg(Of String)(args, "range", "")
            Dim shapeType As String = (GetArg(Of String)(args, "shapeType", "roundedRect") & "").ToLowerInvariant()
            Dim fill As String = GetArg(Of String)(args, "fill", "")
            Dim line As String = GetArg(Of String)(args, "line", "")
            Dim lineWeight As Double = GetArg(Of Double)(args, "lineWeight", 1.0)
            Dim noShadow As Boolean = GetArg(Of Boolean)(args, "noShadow", True)
            Dim cornerRadius As Double = GetArg(Of Double)(args, "cornerRadius", 0.0)
            Dim linesJson As String = GetArg(Of String)(args, "lines", "")
            Dim sendToBack As Boolean = GetArg(Of Boolean)(args, "sendToBack", False)

            ' 兼容 sheet!range 前缀：sheet 为空时从 range 前缀取表名
            If String.IsNullOrWhiteSpace(rangeAddr) Then
                result.IsError = True
                Dim _keys As String = If(args IsNot Nothing, String.Join(",", args.Keys), "null")
                result.Content.Add(New McpTextContent With {.Text = "[AddShape诊断] range 为空。收到参数 keys=[" & _keys & "]"})
                Return result
            End If
            Dim sheetFromRange As String = ""
            If rangeAddr.Contains("!") AndAlso Not rangeAddr.Contains("[") Then
                Dim i = rangeAddr.IndexOf("!")
                Dim sh = rangeAddr.Substring(0, i)
                Dim addr = rangeAddr.Substring(i + 1)
                If Not (addr.StartsWith("$") OrElse Char.IsLetter(addr(0)) OrElse addr.StartsWith("R") AndAlso addr.Contains("C")) Then
                    sheetFromRange = sh.Trim("'"c)
                    rangeAddr = addr
                End If
            End If
            If String.IsNullOrWhiteSpace(sheetName) Then sheetName = sheetFromRange

            Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing
            Try
                ws = If(String.IsNullOrWhiteSpace(sheetName), TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet), TryCast(excelApp.ActiveWorkbook.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet))
            Catch : End Try
            ' 【关键修复】传了 sheet 参数但找不到时，必须报错，不能静默回退到 ActiveSheet
            If ws Is Nothing AndAlso Not String.IsNullOrWhiteSpace(sheetName) Then
                result.IsError = True
                Dim curWs = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                Dim curName As String = If(curWs IsNot Nothing, curWs.Name, "未知")
                result.Content.Add(New McpTextContent With {.Text = $"找不到工作表: [{sheetName}]，请先创建该工作表。当前活动表: [{curName}]"})
                Return result
            End If
            ' 未传 sheet 参数时，才默认使用 ActiveSheet
            If ws Is Nothing Then
                ws = TryCast(excelApp.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
            End If
            ' 解析锚点范围获取磅坐标
            Dim rng As Microsoft.Office.Interop.Excel.Range = Nothing
            Try : rng = ws.Range(rangeAddr) : Catch : rng = Nothing : End Try
            If rng Is Nothing Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"无法解析锚点范围 range='{rangeAddr}'"})
                Return result
            End If
            Dim left As Double = 0, top As Double = 0, width As Double = 0, height As Double = 0
            Try
                left = CDbl(rng.Left)
                top = CDbl(rng.Top)
                width = CDbl(rng.Width)
                height = CDbl(rng.Height)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"读取锚点范围坐标失败：{ex.Message}"})
                Utils.ComReleaseHelper.Release(rng)
                Return result
            Finally
                Utils.ComReleaseHelper.Release(rng)
            End Try

            Dim shape As Microsoft.Office.Interop.Excel.Shape = Nothing
            Try
                Dim msoType As MsoAutoShapeType = MsoAutoShapeType.msoShapeRectangle
                Select Case shapeType
                    Case "roundedrect" : msoType = MsoAutoShapeType.msoShapeRoundedRectangle
                    Case "oval", "circle" : msoType = MsoAutoShapeType.msoShapeOval
                    Case Else : msoType = MsoAutoShapeType.msoShapeRectangle
                End Select
                shape = ws.Shapes.AddShape(msoType, CSng(left), CSng(top), CSng(width), CSng(height))

                ' 填充
                Try
                    If String.Equals(fill, "none", StringComparison.OrdinalIgnoreCase) Then
                        shape.Fill.Visible = MsoTriState.msoFalse
                    ElseIf Not String.IsNullOrWhiteSpace(fill) Then
                        shape.Fill.Solid()
                        shape.Fill.ForeColor.RGB = OLE(fill)
                        shape.Fill.Visible = MsoTriState.msoTrue
                    End If
                Catch : End Try
                ' 边框
                Try
                    If String.Equals(line, "none", StringComparison.OrdinalIgnoreCase) Then
                        shape.Line.Visible = MsoTriState.msoFalse
                    ElseIf Not String.IsNullOrWhiteSpace(line) Then
                        shape.Line.ForeColor.RGB = OLE(line)
                        shape.Line.Weight = CSng(lineWeight)
                        shape.Line.Visible = MsoTriState.msoTrue
                    End If
                Catch : End Try
                ' 圆角半径（roundedrect 专用）
                If shapeType = "roundedrect" AndAlso cornerRadius > 0 Then
                    Try
                        If shape.Adjustments.Count >= 1 Then shape.Adjustments.Item(1) = CSng(Math.Max(0, Math.Min(0.5, cornerRadius)))
                    Catch : End Try
                End If
                ' 关闭阴影（专业平面卡片）
                If noShadow Then
                    Try : shape.Shadow.Visible = MsoTriState.msoFalse : Catch : End Try
                End If
                If sendToBack Then
                    Try : shape.ZOrder(MsoZOrderCmd.msoSendToBack) : Catch : End Try
                End If

                ' 形状文字
                Dim lines As JArray = Nothing
                If Not String.IsNullOrWhiteSpace(linesJson) Then
                    Try : lines = JArray.Parse(linesJson) : Catch : lines = Nothing : End Try
                End If
                If lines IsNot Nothing AndAlso lines.Count > 0 Then
                    ApplyShapeText(shape, lines)
                End If
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = $"创建形状失败：{ex.Message}"})
                Return result
            Finally
                Utils.ComReleaseHelper.Release(shape)
            End Try

            Dim actualSheetName As String = "未知"
            Try
                Dim tlCell = shape.TopLeftCell
                actualSheetName = tlCell.Worksheet.Name
            Catch
            End Try
            result.Content.Add(New McpTextContent With {.Text = $"✅ 已绘制形状（{ShapeLabel(shapeType)}），目标工作表=[{ws.Name}]，实际归属工作表=[{actualSheetName}]，范围 {rangeAddr}"})
            Return result
        End Function

        ''' <summary>
        ''' hex(#RRGGBB / RRGGBB / #RGB) → OLE 长整型（Excel Interior.Color / Font.Color / Borders.Color 用）。
        ''' 必须 ToOle：直接赋 System.Drawing.Color 会在 COM 层类型不匹配并被空 Catch 吞掉 → 底色/边框色看似成功实则未生效。
        ''' 不用 ColorTranslator.FromHtml 解析 hex：无 # 的 "FFFFCC" 会被当成 Int32 解析而抛错；模型/硬短路常漏写 #。
        ''' </summary>
        Private Shared Function OLE(hex As String) As Integer
            Dim raw As String = If(hex, "").Trim()
            If raw.Length = 0 Then Throw New ArgumentException("颜色值为空")
            Dim s As String = raw
            If s.StartsWith("#", StringComparison.Ordinal) Then s = s.Substring(1)
            If s.StartsWith("0x", StringComparison.OrdinalIgnoreCase) Then s = s.Substring(2)
            ' #AARRGGBB → 取 RRGGBB
            If s.Length = 8 Then s = s.Substring(2)
            If s.Length = 3 Then
                s = New String(s(0), 2) & New String(s(1), 2) & New String(s(2), 2)
            End If
            If s.Length = 6 Then
                Dim r As Integer = Convert.ToInt32(s.Substring(0, 2), 16)
                Dim g As Integer = Convert.ToInt32(s.Substring(2, 2), 16)
                Dim b As Integer = Convert.ToInt32(s.Substring(4, 2), 16)
                Return ColorTranslator.ToOle(Color.FromArgb(r, g, b))
            End If
            ' 命名色（red / lightyellow 等）
            Return ColorTranslator.ToOle(ColorTranslator.FromHtml(raw))
        End Function

        Private Shared Function ShapeLabel(t As String) As String
            Select Case t
                Case "roundedrect" : Return "圆角卡片"
                Case "oval", "circle" : Return "圆形"
                Case Else : Return "矩形色块"
            End Select
        End Function

        ''' <summary>
        ''' 往形状写入多行文字：lines=[{text,size,color,bold,alignH,vAlignH}...]。
        ''' 优先 TextFrame2 逐段排版，异常时回退整段基础样式（宁可样式简化不可失败）。
        ''' </summary>
        Private Shared Sub ApplyShapeText(shape As Microsoft.Office.Interop.Excel.Shape, lines As JArray)
            Try
                Dim sb As New System.Text.StringBuilder()
                Dim lineObjs As New List(Of JObject)()
                For i = 0 To lines.Count - 1
                    Dim jo = TryCast(lines(i), JObject)
                    If jo Is Nothing Then Continue For
                    lineObjs.Add(jo)
                    If i > 0 Then sb.Append(vbCr)   ' TextRange2 用回车分段，便于逐行设字号/颜色
                    sb.Append(jo.Value(Of String)("text") & "")
                Next
                If lineObjs.Count = 0 Then Return

                ' 边距与换行
                Try
                    shape.TextFrame2.MarginLeft = CSng(3)
                    shape.TextFrame2.MarginRight = CSng(3)
                    shape.TextFrame2.MarginTop = CSng(1)
                    shape.TextFrame2.MarginBottom = CSng(1)
                Catch : End Try
                Try : shape.TextFrame2.WordWrap = MsoTriState.msoTrue : Catch : End Try
                Try : shape.TextFrame2.HorizontalAnchor = MsoHorizontalAnchor.msoAnchorNone : Catch : End Try
                Try : shape.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle : Catch : End Try

                Dim txt As TextRange2 = Nothing
                Try : txt = shape.TextFrame2.TextRange : Catch : txt = Nothing : End Try
                If txt Is Nothing Then
                    Try : shape.TextFrame.Characters().Insert(sb.ToString()) : Catch : End Try
                    Return
                End If
                txt.Text = sb.ToString()

                ' 自第一行派生基础样式，再逐行覆盖
                For p = 1 To lineObjs.Count
                    Dim jo = lineObjs(p - 1)
                    Dim sz As Double = If(jo("size") IsNot Nothing, jo.Value(Of Double)("size"), 10.0)
                    Dim col As String = If(jo("color") IsNot Nothing, jo.Value(Of String)("color"), "")
                    Dim bold As Boolean = If(jo("bold") IsNot Nothing, jo.Value(Of Boolean)("bold"), False)
                    Dim alignH As String = If(jo("alignH") IsNot Nothing, jo.Value(Of String)("alignH"), "center")
                    Try
                        Dim tr As TextRange2 = Nothing
                        If lineObjs.Count = 1 Then
                            tr = txt
                        Else
                            tr = txt.Paragraphs(p)
                        End If
                        If tr IsNot Nothing Then
                            Try : tr.Font.Name = "Microsoft YaHei" : Catch : End Try
                            If sz > 0 Then Try : tr.Font.Size = CSng(sz) : Catch : End Try
                            Try : tr.Font.Bold = If(bold, MsoTriState.msoTrue, MsoTriState.msoFalse) : Catch : End Try
                            If Not String.IsNullOrWhiteSpace(col) Then
                                Try : tr.Font.Fill.ForeColor.RGB = OLE(col) : Catch : End Try
                            End If
                            Try
                                Dim pg = tr.ParagraphFormat
                                Select Case alignH.ToLowerInvariant()
                                    Case "left" : pg.Alignment = MsoParagraphAlignment.msoAlignLeft
                                    Case "right" : pg.Alignment = MsoParagraphAlignment.msoAlignRight
                                    Case Else : pg.Alignment = MsoParagraphAlignment.msoAlignCenter
                                End Select
                            Catch : End Try
                        End If
                    Catch
                    End Try
                Next
            Catch
            End Try
        End Sub

    End Class

End Namespace






