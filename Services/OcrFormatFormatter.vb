Imports System.Text

Namespace Services

    ''' <summary>
    ''' OCR 格式化器
    ''' 将 OcrDocument 结构化结果转换为 Markdown / HTML / Excel 格式
    ''' </summary>
    Public Class OcrFormatFormatter

        ''' <summary>
        ''' 将 OCR 文档转换为指定格式的文本
        ''' </summary>
        Public Shared Function Format(doc As OcrDocument, outputFormat As OcrOutputFormat) As String
            Select Case outputFormat
                Case OcrOutputFormat.Markdown
                    Return FormatAsMarkdown(doc)
                Case OcrOutputFormat.Html
                    Return FormatAsHtml(doc)
                Case OcrOutputFormat.Excel
                    Return FormatAsExcelCsv(doc)
                Case Else
                    ' 纯文本默认输出 raw_ocr_data（text + box + confidence）
                    Return FormatAsRawOcrData(doc)
            End Select
        End Function

        ''' <summary>
        ''' 生成移动端使用的纯 Tab 分隔表格文本（无 BOM、无元数据、无注释）
        ''' 仅包含表格数据行，供前端直接解析为二维数组
        ''' </summary>
        Public Shared Function FormatAsMobileTable(doc As OcrDocument) As String
            Try
                Dim fullText = FormatAsExcelCsv(doc)

                ' 去掉 BOM（如果存在）
                If fullText.Length > 0 AndAlso AscW(fullText(0)) = &HFEFF Then
                    fullText = fullText.Substring(1)
                End If

                ' 按行分割，过滤掉元数据行(#开头)、HTML注释行、空行
                Dim lines = fullText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
                Dim tableLines As New List(Of String)()
                For Each line In lines
                    Dim trimmed = line.Trim()
                    If String.IsNullOrEmpty(trimmed) Then Continue For
                    If trimmed.StartsWith("#"c) Then Continue For
                    If trimmed.StartsWith("<!--") AndAlso trimmed.EndsWith("-->") Then Continue For
                    tableLines.Add(trimmed)
                Next

                ' 如果有有效的表格数据，返回
                If tableLines.Count > 0 Then
                    Return String.Join(vbLf, tableLines)
                End If
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine($"[OcrFormatFormatter] FormatAsMobileTable 失败: {ex.Message}")
            End Try

            ' 兜底：返回 RawText（纯文本换行），至少保证有内容
            Return If(doc?.RawText, "")
        End Function

        ''' <summary>
        ''' 输出原始 OCR 数据（text + confidence + box），便于调试和导出
        ''' 优先使用 RawCellElements（单元格级原始元素），没有时才用 Elements
        ''' </summary>
        Private Shared Function FormatAsRawOcrData(doc As OcrDocument) As String
            Dim sourceElements = If(doc.RawCellElements.Count > 0, doc.RawCellElements, doc.Elements)
            Dim items As New List(Of Object)()
            For Each element In sourceElements
                If element.BoundingBox Is Nothing Then Continue For
                Dim box = New Integer()() {
                    New Integer() {element.BoundingBox.Left, element.BoundingBox.Top},
                    New Integer() {element.BoundingBox.Right, element.BoundingBox.Top},
                    New Integer() {element.BoundingBox.Right, element.BoundingBox.Bottom},
                    New Integer() {element.BoundingBox.Left, element.BoundingBox.Bottom}
                }
                items.Add(New With {
                    .text = element.Text,
                    .confidence = element.Confidence,
                    .box = box
                })
            Next

            If items.Count > 0 Then
                Return Newtonsoft.Json.JsonConvert.SerializeObject(items, Newtonsoft.Json.Formatting.Indented)
            End If

            ' 兜底：没有元素时返回原始文本
            Return If(Not String.IsNullOrEmpty(doc.RawText), doc.RawText, "")
        End Function

        ''' <summary>
        ''' 转换为 Markdown 格式
        ''' </summary>
        Private Shared Function FormatAsMarkdown(doc As OcrDocument) As String
            ' 如果有 PP-StructureV3 直接输出的 Markdown，优先使用
            If Not String.IsNullOrEmpty(doc.MarkdownText) Then
                Return doc.MarkdownText
            End If

            ' 如果整篇只有段落，统一输出为行号/内容表格，避免 Markdown 看起来像纯文本
            Dim paragraphTable = FormatParagraphsAsMarkdownTable(doc)
            If paragraphTable IsNot Nothing Then
                Return paragraphTable
            End If

            Dim sb As New StringBuilder()

            For Each element In doc.Elements
                Select Case element.Type
                    Case OcrElementType.Heading
                        Dim prefix As String = New String("#"c, Math.Min(element.HeadingLevel, 6))
                        sb.AppendLine($"{prefix} {element.Text}")
                        sb.AppendLine()

                    Case OcrElementType.Paragraph
                        sb.AppendLine(element.Text)
                        sb.AppendLine()

                    Case OcrElementType.Dimension, OcrElementType.GdtSymbol, OcrElementType.Roughness, OcrElementType.Balloon, OcrElementType.DatumCircle
                        sb.AppendLine(FormatCalloutLine(element))
                        sb.AppendLine()

                    Case OcrElementType.ListItem
                        Dim indent As String = New String(" "c, element.ListLevel * 2)
                        Dim marker As String = If(String.IsNullOrEmpty(element.ListMarker), "-", element.ListMarker)
                        sb.AppendLine($"{indent}{marker} {element.Text}")

                    Case OcrElementType.Table
                        If element.TableData IsNot Nothing Then
                            ' 优先使用表格 HTML（来自 PP-StructureV3）
                            If Not String.IsNullOrEmpty(element.TableData.HtmlContent) Then
                                sb.AppendLine(element.TableData.HtmlContent)
                            Else
                                sb.AppendLine(FormatTableAsMarkdown(element.TableData))
                            End If
                            sb.AppendLine()
                        End If

                    Case OcrElementType.Formula
                        If element.FormulaData IsNot Nothing AndAlso Not String.IsNullOrEmpty(element.FormulaData.Latex) Then
                            sb.AppendLine($"$$ {element.FormulaData.Latex} $$")
                        Else
                            sb.AppendLine($"$$ {element.Text} $$")
                        End If
                        sb.AppendLine()

                    Case OcrElementType.LayoutRegion
                        sb.AppendLine($"<!-- {element.LayoutLabel}: {element.Text} -->")
                        sb.AppendLine()

                    Case OcrElementType.HorizontalRule
                        sb.AppendLine("---")
                        sb.AppendLine()

                    Case OcrElementType.PageBreak
                        sb.AppendLine("---")
                        sb.AppendLine()
                End Select
            Next

            ' 如果没有结构化元素，回退到原始文本
            If sb.Length = 0 Then
                Return doc.RawText
            End If

            Return sb.ToString().Trim()
        End Function

        Private Shared Function FormatCalloutLine(element As OcrElement) As String
            If element Is Nothing Then Return ""
            Return If(element.Text, "")
        End Function

        ''' <summary>
        ''' 如果文档只有段落，输出为 Markdown 表格（行号/内容），避免看起来像纯文本
        ''' </summary>
        Private Shared Function FormatParagraphsAsMarkdownTable(doc As OcrDocument) As String
            Dim paragraphs = doc.Elements.Where(Function(e) e.Type = OcrElementType.Paragraph).ToList()
            If paragraphs.Count = 0 OrElse paragraphs.Count <> doc.Elements.Count Then
                Return Nothing
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine("| 行号 | 内容 |")
            sb.AppendLine("| --- | --- |")
            For i = 0 To paragraphs.Count - 1
                sb.AppendLine($"| {i + 1} | {paragraphs(i).Text.Replace("|", "\|")} |")
            Next
            Return sb.ToString().Trim()
        End Function

        ''' <summary>
        ''' 将表格转换为 Markdown 表格
        ''' </summary>
        Private Shared Function FormatTableAsMarkdown(table As OcrTableData) As String
            Dim sb As New StringBuilder()

            ' 表头
            If table.Header.Count > 0 Then
                sb.AppendLine("| " & String.Join(" | ", table.Header) & " |")
                ' 分隔行
                sb.AppendLine("| " & String.Join(" | ", table.Header.Select(Function(c) New String("-"c, Math.Max(3, c.Length)))) & " |")
            End If

            ' 数据行
            For Each row In table.Rows
                ' 补齐到列数
                While row.Count < table.ColumnCount
                    row.Add("")
                End While
                sb.AppendLine("| " & String.Join(" | ", row.Take(table.ColumnCount)) & " |")
            Next

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 转换为 HTML 格式
        ''' </summary>
        Private Shared Function FormatAsHtml(doc As OcrDocument) As String
            ' 如果有 PP-StructureV3 直接输出的 Markdown，包装为 HTML
            If Not String.IsNullOrEmpty(doc.MarkdownText) Then
                Return $"<!DOCTYPE html><html lang=""zh-CN""><head><meta charset=""UTF-8""><title>OCR 识别结果</title></head><body><pre>{EscapeHtml(doc.MarkdownText)}</pre></body></html>"
            End If

            Dim sb As New StringBuilder()
            sb.AppendLine("<!DOCTYPE html>")
            sb.AppendLine("<html lang=""zh-CN"">")
            sb.AppendLine("<head><meta charset=""UTF-8""><title>OCR 识别结果</title>")
            sb.AppendLine("<style>")
            sb.AppendLine("body { font-family: -apple-system, sans-serif; max-width: 800px; margin: 20px auto; padding: 0 16px; }")
            sb.AppendLine("table { border-collapse: collapse; width: 100%; margin: 8px 0; }")
            sb.AppendLine("th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; }")
            sb.AppendLine("th { background: #f0f4ff; }")
            sb.AppendLine(".formula { background: #f8f9fa; padding: 8px 12px; border: 1px solid #e0e0e0; border-radius: 4px; font-family: 'Courier New', monospace; }")
            sb.AppendLine(".layout-label { display: inline-block; background: #e3f2fd; padding: 2px 8px; border-radius: 3px; font-size: 11px; color: #1565c0; }")
            sb.AppendLine("hr { border: none; border-top: 1px solid #ddd; margin: 16px 0; }")
            sb.AppendLine("</style></head><body>")
            sb.AppendLine($"<p><em>来源: {EscapeHtml(doc.SourceFileName)} | 引擎: {EscapeHtml(doc.EngineName)}</em></p>")

            ' 如果整篇只有段落，输出为表格，避免 HTML 看起来只是一堆 <p>
            Dim paragraphOnly = doc.Elements.All(Function(e) e.Type = OcrElementType.Paragraph)
            If paragraphOnly AndAlso doc.Elements.Count > 0 Then
                sb.AppendLine("<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-family:sans-serif;'>")
                sb.AppendLine("<thead><tr><th>行号</th><th>内容</th></tr></thead><tbody>")
                For i = 0 To doc.Elements.Count - 1
                    sb.AppendLine($"<tr><td>{i + 1}</td><td>{EscapeHtml(doc.Elements(i).Text)}</td></tr>")
                Next
                sb.AppendLine("</tbody></table>")
            Else
                For Each element In doc.Elements
                    Select Case element.Type
                        Case OcrElementType.Heading
                            Dim level As Integer = Math.Min(element.HeadingLevel, 6)
                            sb.AppendLine($"<h{level}>{EscapeHtml(element.Text)}</h{level}>")

                        Case OcrElementType.Paragraph
                            sb.AppendLine($"<p>{EscapeHtml(element.Text)}</p>")

                        Case OcrElementType.ListItem
                            Dim indent = element.ListLevel * 20
                            sb.AppendLine($"<p style=""padding-left:{indent}px"">{EscapeHtml(element.ListMarker)} {EscapeHtml(element.Text)}</p>")

                        Case OcrElementType.Table
                            If element.TableData IsNot Nothing Then
                                sb.AppendLine(FormatTableAsHtml(element.TableData))
                            End If

                        Case OcrElementType.Formula
                            Dim latex = If(element.FormulaData IsNot Nothing AndAlso Not String.IsNullOrEmpty(element.FormulaData.Latex),
                                          element.FormulaData.Latex, element.Text)
                            sb.AppendLine($"<div class=""formula"">📐 {EscapeHtml(latex)}</div>")

                        Case OcrElementType.LayoutRegion
                            sb.AppendLine($"<div><span class=""layout-label"">{EscapeHtml(element.LayoutLabel)}</span> {EscapeHtml(element.Text)}</div>")

                        Case OcrElementType.HorizontalRule
                            sb.AppendLine("<hr>")

                        Case OcrElementType.PageBreak
                            sb.AppendLine("<hr style=""border-top: 2px dashed #999;"">")
                    End Select
                Next
            End If

            ' 如果没有结构化元素，使用原始文本
            If doc.Elements.Count = 0 AndAlso Not String.IsNullOrEmpty(doc.RawText) Then
                sb.AppendLine($"<pre>{EscapeHtml(doc.RawText)}</pre>")
            End If

            sb.AppendLine("</body></html>")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 将表格转换为 HTML 表格
        ''' </summary>
        Private Shared Function FormatTableAsHtml(table As OcrTableData) As String
            Dim sb As New StringBuilder()
            sb.AppendLine("<table>")

            ' 表头
            If table.Header.Count > 0 Then
                sb.AppendLine("<thead><tr>")
                For Each cell In table.Header
                    sb.AppendLine($"<th>{EscapeHtml(cell)}</th>")
                Next
                sb.AppendLine("</tr></thead>")
            End If

            ' 数据行
            sb.AppendLine("<tbody>")
            For Each row In table.Rows
                sb.AppendLine("<tr>")
                Dim cols = Math.Max(1, table.ColumnCount)
                For i As Integer = 0 To cols - 1
                    Dim val As String = If(i < row.Count, row(i), "")
                    sb.AppendLine($"<td>{EscapeHtml(val)}</td>")
                Next
                sb.AppendLine("</tr>")
            Next
            sb.AppendLine("</tbody></table>")

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 转换为 Excel CSV 格式（UTF-8 BOM，制表符分隔）
        ''' </summary>
        Private Shared Function FormatAsExcelCsv(doc As OcrDocument) As String
            Dim sb As New StringBuilder()
            ' UTF-8 BOM
            sb.Append(ChrW(&HFEFF))

            ' 元数据行（简洁化，不占用表格区域）
            sb.AppendLine($"# 来源:{doc.SourceFileName} 引擎:{doc.EngineName} 页数:{doc.TotalPages} 耗时:{doc.ElapsedSeconds:F1}s")
            sb.AppendLine()

            ' 1) 优先输出真正的 Table 元素
            Dim tableElements = doc.Elements.Where(Function(e) e.Type = OcrElementType.Table AndAlso e.TableData IsNot Nothing).ToList()
            For Each tableEl In tableElements
                Dim tbl = tableEl.TableData
                If tbl.Header.Count > 0 Then
                    sb.AppendLine(String.Join(vbTab, tbl.Header.Select(Function(c) EscapeCsvField(c))))
                End If
                For Each row In tbl.Rows
                    While row.Count < tbl.ColumnCount
                        row.Add("")
                    End While
                    sb.AppendLine(String.Join(vbTab, row.Take(tbl.ColumnCount).Select(Function(c) EscapeCsvField(c))))
                Next
                sb.AppendLine()
            Next

            ' 1.5) 没有真正 Table 元素时，按优先级尝试多种表格重建策略
            If tableElements.Count = 0 Then
                ' 优先使用 RawCellElements（单元格级原始元素），避免 BuildDocumentFromPositionedResults 的按行合并丢失精度
                Dim sourceElements = If(doc.RawCellElements.Count > 0, doc.RawCellElements, doc.Elements)
                Dim positionedElements = sourceElements.Where(Function(e) e.BoundingBox IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(e.Text)).ToList()

                ' 策略 A：Python table_parser_v2 移植版（坐标聚类 + rowspan）
                If positionedElements.Count >= 4 Then
                    Dim parser = New OcrTableParser(positionedElements)
                    Dim parsedTable = parser.Parse()
                    If parsedTable IsNot Nothing AndAlso parsedTable.ColumnCount >= 2 AndAlso parsedTable.Rows.Count >= 1 Then
                        AppendTableData(sb, parsedTable)
                        AppendHtmlTableComment(sb, parser.ToHtml())
                        Return sb.ToString()
                    End If
                End If

                ' 策略 B：原有位置对齐重建（增量列中心 + 滑动平均）
                Dim fallbackTable = TryRebuildTableFromPositionedElements(sourceElements)
                If fallbackTable IsNot Nothing AndAlso fallbackTable.ColumnCount >= 2 AndAlso fallbackTable.Rows.Count >= 1 Then
                    AppendTableData(sb, fallbackTable)
                    Return sb.ToString()
                End If

                ' 策略 C：文本启发式拆分（按行分组 + 空格/制表符切列）
                Dim textTable = TryRebuildTableFromText(sourceElements)
                If textTable IsNot Nothing AndAlso textTable.ColumnCount >= 2 AndAlso textTable.Rows.Count >= 1 Then
                    AppendTableData(sb, textTable)
                    Return sb.ToString()
                End If

                ' 兜底：输出诊断信息，并附带原始 OCR 数据摘要
                sb.AppendLine("# 表格解析未成功")
                sb.AppendLine($"# OCR元素总数:{sourceElements.Count} 含坐标:{positionedElements.Count}")
                If positionedElements.Count > 0 Then
                    sb.AppendLine($"# 前5个元素文本:" + String.Join(" | ", positionedElements.Take(5).Select(Function(e) e.Text)))
                End If
                sb.AppendLine("# 请切换到「纯文本」格式查看完整 raw_ocr_data")
                Return sb.ToString()
            End If

            ' 2) 段落/标题/ListItem 统一按多列拆分输出，不再加"段落"前缀
            Dim rowLines As New List(Of String())()
            Dim maxCols As Integer = 0
            For Each element In doc.Elements
                If element.Type = OcrElementType.Paragraph OrElse element.Type = OcrElementType.Heading OrElse element.Type = OcrElementType.ListItem Then
                    Dim cells = SplitTextIntoColumns(element.Text)
                    If cells.Length > 0 Then
                        rowLines.Add(cells)
                        maxCols = Math.Max(maxCols, cells.Length)
                    End If
                End If
            Next

            If rowLines.Count > 0 Then
                For Each cells In rowLines
                    Dim padded = cells.ToList()
                    While padded.Count < maxCols
                        padded.Add("")
                    End While
                    sb.AppendLine(String.Join(vbTab, padded.Take(maxCols).Select(Function(c) EscapeCsvField(c))))
                Next
            End If

            ' 3) 兜底：没有任何元素时直接输出原始文本行
            If doc.Elements.Count = 0 AndAlso Not String.IsNullOrEmpty(doc.RawText) Then
                Dim lines = doc.RawText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.RemoveEmptyEntries)
                For Each line In lines
                    If String.IsNullOrWhiteSpace(line) Then Continue For
                    Dim cells = SplitTextIntoColumns(line)
                    sb.AppendLine(String.Join(vbTab, cells.Select(Function(c) EscapeCsvField(c))))
                Next
            End If

            Return sb.ToString()
        End Function

        ''' <summary>
        ''' 转义 HTML 文本
        ''' </summary>
        Private Shared Function EscapeHtml(text As String) As String
            If String.IsNullOrEmpty(text) Then Return ""
            Return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("""", "&quot;")
        End Function

        ''' <summary>
        ''' 转义 CSV 字段（包含制表符/引号时包裹）
        ''' </summary>
        Private Shared Function EscapeCsvField(value As String) As String
            If String.IsNullOrEmpty(value) Then Return String.Empty
            Dim dq As Char = ChrW(34)
            If value.Contains(vbTab) OrElse value.Contains(dq) OrElse value.Contains(vbLf) Then
                Return dq & value.Replace(dq.ToString(), dq.ToString() & dq.ToString()) & dq
            End If
            Return value
        End Function

        ''' <summary>
        ''' 将 OcrTableData 追加为制表符分隔行
        ''' </summary>
        Private Shared Sub AppendTableData(sb As StringBuilder, table As OcrTableData)
            If table.Header.Count > 0 Then
                sb.AppendLine(String.Join(vbTab, table.Header.Select(Function(c) EscapeCsvField(c))))
            End If
            Dim colCount = Math.Max(table.ColumnCount, If(table.Header.Count > 0, table.Header.Count, 0))
            For Each row In table.Rows
                While row.Count < colCount
                    row.Add("")
                End While
                sb.AppendLine(String.Join(vbTab, row.Take(colCount).Select(Function(c) EscapeCsvField(c))))
            Next
        End Sub

        ''' <summary>
        ''' 附加 Base64 HTML 表格预览注释
        ''' </summary>
        Private Shared Sub AppendHtmlTableComment(sb As StringBuilder, htmlTable As String)
            If String.IsNullOrEmpty(htmlTable) Then Return
            sb.AppendLine()
            Dim b64 = Convert.ToBase64String(System.Text.Encoding.UTF8.GetBytes(htmlTable))
            sb.AppendLine($"<!-- HTML_TABLE_B64:{b64} -->")
        End Sub

        ''' <summary>
        ''' 文本启发式表格重建（增强版）：
        ''' 1) 每个元素优先按 tab / 多空格拆分为候选单元格，并估算每个子单元格的 left/right
        ''' 2) 按 Y 坐标分组为行
        ''' 3) 从清晰分隔行（tab/多空格）确定目标列数
        ''' 4) 对仍为一整块的行尝试按单空格拆分，使其对齐目标列数
        ''' 5) 按 left 位置聚类列中心，跨行对齐列
        ''' </summary>
        Private Shared Function TryRebuildTableFromText(elements As List(Of OcrElement)) As OcrTableData
            If elements Is Nothing OrElse elements.Count < 4 Then Return Nothing

            Dim textElements = elements.Where(Function(e) Not String.IsNullOrWhiteSpace(e.Text)).OrderBy(Function(e) If(e.BoundingBox IsNot Nothing, e.BoundingBox.Top, 0)).ToList()
            If textElements.Count < 4 Then Return Nothing

            ' 按 Y 坐标分组（阈值 40px），无坐标则每元素一行
            Dim yGroups As New List(Of List(Of OcrElement))()
            Dim currentYGroup As New List(Of OcrElement)()
            Dim lastY As Integer? = Nothing
            Const yThreshold As Integer = 40

            For Each e In textElements
                Dim hasY = e.BoundingBox IsNot Nothing
                Dim cy = If(hasY, CInt((e.BoundingBox.Top + e.BoundingBox.Bottom) / 2), 0)

                If Not hasY Then
                    If currentYGroup.Count > 0 Then
                        yGroups.Add(currentYGroup)
                        currentYGroup = New List(Of OcrElement)()
                    End If
                    yGroups.Add(New List(Of OcrElement)() From {e})
                    lastY = Nothing
                Else
                    If lastY.HasValue AndAlso Math.Abs(cy - lastY.Value) <= yThreshold Then
                        currentYGroup.Add(e)
                    Else
                        If currentYGroup.Count > 0 Then yGroups.Add(currentYGroup)
                        currentYGroup = New List(Of OcrElement)() From {e}
                    End If
                    lastY = cy
                End If
            Next
            If currentYGroup.Count > 0 Then yGroups.Add(currentYGroup)

            If yGroups.Count < 2 Then Return Nothing

            ' 预拆分：tab / 多空格，保留估算坐标
            Dim preSplitRows As New List(Of List(Of CellCandidate))()
            For Each group In yGroups
                Dim rowCells As New List(Of CellCandidate)()
                For Each e In group.OrderBy(Function(x) If(x.BoundingBox IsNot Nothing, x.BoundingBox.Left, 0))
                    Dim parts = SplitTextIntoColumnsForTable(e.Text)
                    If parts.Length = 0 Then Continue For

                    If parts.Length = 1 OrElse e.BoundingBox Is Nothing Then
                        rowCells.Add(New CellCandidate With {
                            .Text = parts(0).Trim(),
                            .Left = If(e.BoundingBox IsNot Nothing, e.BoundingBox.Left, 0),
                            .Right = If(e.BoundingBox IsNot Nothing, e.BoundingBox.Right, 0),
                            .HasBox = e.BoundingBox IsNot Nothing
                        })
                    Else
                        Dim totalWidth = Math.Max(1, e.BoundingBox.Right - e.BoundingBox.Left)
                        Dim totalLen = Math.Max(1, parts.Sum(Function(p) p.Length))
                        Dim currentLeft = e.BoundingBox.Left
                        For i = 0 To parts.Length - 1
                            Dim w = CInt(totalWidth * parts(i).Length / totalLen)
                            Dim right = Math.Min(e.BoundingBox.Right, currentLeft + w)
                            rowCells.Add(New CellCandidate With {
                                .Text = parts(i).Trim(),
                                .Left = currentLeft,
                                .Right = right,
                                .HasBox = True
                            })
                            currentLeft = right
                        Next
                    End If
                Next
                preSplitRows.Add(rowCells)
            Next

            ' 从清晰分隔行确定目标列数
            Dim delimiterColCounts = preSplitRows.Where(Function(r) r.Count > 1).GroupBy(Function(r) r.Count).OrderByDescending(Function(g) g.Count).ThenByDescending(Function(g) g.Key).ToList()
            Dim targetCols As Integer = 0
            If delimiterColCounts.Count > 0 Then
                targetCols = delimiterColCounts(0).Key
            End If

            ' 兜底：没有 tab/多空格分隔时，看看单空格是否能产生多列
            If targetCols < 2 Then
                Dim maxSpaceCols = 0
                For i = 0 To yGroups.Count - 1
                    Dim e = yGroups(i)(0)
                    If e.Text.Contains(" ") AndAlso Not e.Text.Contains(vbTab) Then
                        Dim spaceParts = e.Text.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Where(Function(p) Not String.IsNullOrWhiteSpace(p)).ToList()
                        If spaceParts.Count > maxSpaceCols Then maxSpaceCols = spaceParts.Count
                    End If
                Next
                targetCols = maxSpaceCols
            End If

            If targetCols < 2 Then Return Nothing

            ' 对仍为一整块的行尝试单空格拆分，使其接近目标列数
            For i = 0 To preSplitRows.Count - 1
                Dim row = preSplitRows(i)
                If row.Count <> 1 Then Continue For

                Dim e = yGroups(i)(0)
                Dim text = e.Text.Trim()
                If Not text.Contains(" ") OrElse text.Contains(vbTab) Then Continue For

                Dim spaceParts = text.Split({" "c}, StringSplitOptions.RemoveEmptyEntries).Select(Function(p) p.Trim()).Where(Function(p) Not String.IsNullOrEmpty(p)).ToList()
                If spaceParts.Count < 2 Then Continue For

                ' 拆分后若超过目标列数，尾部合并；不足则补空（由后续对齐处理）
                If spaceParts.Count > targetCols Then
                    Dim keep = spaceParts.Take(targetCols - 1).ToList()
                    keep.Add(String.Join(" ", spaceParts.Skip(targetCols - 1)))
                    spaceParts = keep
                End If

                Dim totalWidth = If(e.BoundingBox IsNot Nothing, Math.Max(1, e.BoundingBox.Right - e.BoundingBox.Left), spaceParts.Sum(Function(p) p.Length) * 20)
                Dim totalLen = Math.Max(1, spaceParts.Sum(Function(p) p.Length))
                Dim currentLeft = If(e.BoundingBox IsNot Nothing, e.BoundingBox.Left, 0)
                Dim newCells As New List(Of CellCandidate)()
                For j = 0 To spaceParts.Count - 1
                    Dim w = CInt(totalWidth * spaceParts(j).Length / totalLen)
                    newCells.Add(New CellCandidate With {
                        .Text = spaceParts(j),
                        .Left = currentLeft,
                        .Right = currentLeft + w,
                        .HasBox = e.BoundingBox IsNot Nothing
                    })
                    currentLeft += w
                Next
                preSplitRows(i) = newCells
            Next

            ' 按 left 位置聚类列
            Dim boxedCells = preSplitRows.SelectMany(Function(r) r).Where(Function(c) c.HasBox).ToList()
            If boxedCells.Count < 4 Then Return Nothing

            Dim avgWidth = boxedCells.Average(Function(c) c.Right - c.Left)
            Dim colThreshold = CInt(Math.Max(15, avgWidth * 0.5))
            Dim colCenters As New List(Of Double)()

            For Each row In preSplitRows
                For Each cell In row
                    If Not cell.HasBox Then Continue For
                    Dim cx = (cell.Left + cell.Right) / 2.0
                    Dim bestCol = -1
                    Dim bestDist = Double.MaxValue
                    For j = 0 To colCenters.Count - 1
                        Dim dist = Math.Abs(cx - colCenters(j))
                        If dist < bestDist AndAlso dist <= colThreshold Then
                            bestDist = dist
                            bestCol = j
                        End If
                    Next
                    If bestCol < 0 Then
                        bestCol = colCenters.Count
                        colCenters.Add(cx)
                    Else
                        colCenters(bestCol) = colCenters(bestCol) * 0.7 + cx * 0.3
                    End If
                    cell.ColIndex = bestCol
                Next
            Next

            If colCenters.Count < 2 Then Return Nothing

            ' 列中心排序，得到最终列索引
            Dim sortedCols = colCenters.Select(Function(c, i) New With {.Index = i, .Center = c}).OrderBy(Function(x) x.Center).ToList()
            Dim finalColMap As New Dictionary(Of Integer, Integer)()
            For i = 0 To sortedCols.Count - 1
                finalColMap(sortedCols(i).Index) = i
            Next

            ' 构建表格
            Dim tableData As New OcrTableData()
            Dim finalColCount = sortedCols.Count

            For rowIdx = 0 To preSplitRows.Count - 1
                Dim row = preSplitRows(rowIdx)
                Dim cells(finalColCount - 1) As String
                For j = 0 To cells.Length - 1
                    cells(j) = ""
                Next
                For Each cell In row
                    If Not finalColMap.ContainsKey(cell.ColIndex) Then Continue For
                    Dim finalIdx = finalColMap(cell.ColIndex)
                    If String.IsNullOrEmpty(cells(finalIdx)) Then
                        cells(finalIdx) = cell.Text
                    Else
                        cells(finalIdx) = cells(finalIdx) & " " & cell.Text
                    End If
                Next
                If rowIdx = 0 Then
                    tableData.Header.AddRange(cells)
                Else
                    tableData.Rows.Add(cells.ToList())
                End If
            Next

            tableData.ColumnCount = finalColCount
            tableData.RowCount = tableData.Rows.Count
            Return tableData
        End Function

        ''' <summary>
        ''' 表格重建用的临时单元格候选
        ''' </summary>
        Private Class CellCandidate
            Public Property Text As String
            Public Property Left As Integer
            Public Property Right As Integer
            Public Property HasBox As Boolean
            Public Property ColIndex As Integer
        End Class

        ''' <summary>
        ''' 尝试从带位置信息的元素中重建表格（用于 OCR 降级/未触发 PP-Structure 的情况）
        ''' </summary>
        Private Shared Function TryRebuildTableFromPositionedElements(elements As List(Of OcrElement)) As OcrTableData
            If elements Is Nothing OrElse elements.Count < 2 Then Return Nothing

            ' 只取有位置信息且类型为文本类的元素
            Dim cellElements = elements.Where(Function(e)
                Return e.BoundingBox IsNot Nothing AndAlso
                       (e.Type = OcrElementType.Paragraph OrElse
                        e.Type = OcrElementType.Heading OrElse
                        e.Type = OcrElementType.ListItem) AndAlso
                       Not String.IsNullOrWhiteSpace(e.Text)
            End Function).ToList()
            If cellElements.Count < 4 Then Return Nothing

            ' 辅助函数
            Dim centerY = Function(e As OcrElement) CInt((e.BoundingBox.Top + e.BoundingBox.Bottom) / 2)
            Dim centerX = Function(e As OcrElement) CInt((e.BoundingBox.Left + e.BoundingBox.Right) / 2)
            Dim boxHeight = Function(e As OcrElement) e.BoundingBox.Bottom - e.BoundingBox.Top
            Dim boxWidth = Function(e As OcrElement) e.BoundingBox.Right - e.BoundingBox.Left

            Dim avgHeight = cellElements.Average(Function(e) boxHeight(e))
            Dim avgWidth = cellElements.Average(Function(e) boxWidth(e))
            If avgHeight <= 0 OrElse avgWidth <= 0 Then Return Nothing

            ' 按 Y 聚类成行（阈值宽松一些，适应多行单元格/高分辨率）
            Dim rowThreshold = Math.Max(12, avgHeight * 0.6)
            Dim sortedByY = cellElements.OrderBy(Function(e) centerY(e)).ThenBy(Function(e) centerX(e)).ToList()
            Dim rows As New List(Of List(Of OcrElement))()
            Dim currentRow As New List(Of OcrElement)()
            For Each e In sortedByY
                If currentRow.Count = 0 Then
                    currentRow.Add(e)
                Else
                    Dim lastCY = centerY(currentRow.Last())
                    Dim thisCY = centerY(e)
                    If Math.Abs(thisCY - lastCY) <= rowThreshold Then
                        currentRow.Add(e)
                    Else
                        rows.Add(currentRow)
                        currentRow = New List(Of OcrElement)() From {e}
                    End If
                End If
            Next
            If currentRow.Count > 0 Then rows.Add(currentRow)

            If rows.Count < 2 Then Return Nothing

            ' 对每行按 X 排序，并建立列中心（跨行对齐）
            For Each row In rows
                row.Sort(Function(a, b) centerX(a).CompareTo(centerX(b)))
            Next

            Dim colThreshold = Math.Max(15, avgWidth * 0.75)
            Dim colCenters As New List(Of Double)()
            Dim colAssignments As New Dictionary(Of OcrElement, Integer)()

            For Each row In rows
                For Each cell In row
                    Dim cx = centerX(cell)
                    Dim bestCol As Integer = -1
                    Dim bestDist As Double = Double.MaxValue
                    For i As Integer = 0 To colCenters.Count - 1
                        Dim dist = Math.Abs(cx - colCenters(i))
                        If dist < bestDist AndAlso dist <= colThreshold Then
                            bestDist = dist
                            bestCol = i
                        End If
                    Next
                    If bestCol < 0 Then
                        bestCol = colCenters.Count
                        colCenters.Add(cx)
                    Else
                        ' 更新列中心为滑动平均
                        colCenters(bestCol) = (colCenters(bestCol) * 0.7 + cx * 0.3)
                    End If
                    colAssignments(cell) = bestCol
                Next
            Next

            If colCenters.Count < 2 Then Return Nothing

            ' 按列中心排序，得到最终列索引
            Dim sortedColIndices = colCenters.Select(Function(c, i) New With {.Index = i, .Center = c}).OrderBy(Function(x) x.Center).ToList()
            Dim finalColIndex As New Dictionary(Of Integer, Integer)()
            For i As Integer = 0 To sortedColIndices.Count - 1
                finalColIndex(sortedColIndices(i).Index) = i
            Next

            ' 构建表格数据
            Dim tableData As New OcrTableData()
            Dim headerRow As List(Of String) = Nothing
            For rowIdx As Integer = 0 To rows.Count - 1
                Dim row = rows(rowIdx)
                Dim cells(sortedColIndices.Count - 1) As String
                For i As Integer = 0 To cells.Length - 1
                    cells(i) = ""
                Next
                For Each cell In row
                    Dim rawCol = colAssignments(cell)
                    Dim finalIdx = finalColIndex(rawCol)
                    cells(finalIdx) = cell.Text.Trim()
                Next
                If rowIdx = 0 Then
                    headerRow = cells.ToList()
                Else
                    tableData.Rows.Add(cells.ToList())
                End If
            Next

            ' 第一行作为表头（仅当它与后续行有明显差异，或所有行都是短文本）
            If headerRow IsNot Nothing Then
                Dim headerLooksLikeHeader = True
                If tableData.Rows.Count > 0 Then
                    Dim firstDataRow = tableData.Rows(0)
                    Dim numericHeaderCount = headerRow.Where(Function(c) IsNumeric(c.Replace(".", "").Replace("，", "").Replace(",", ""))).Count()
                    Dim numericDataCount = firstDataRow.Where(Function(c) IsNumeric(c.Replace(".", "").Replace("，", "").Replace(",", ""))).Count()
                    ' 如果第一行数字明显多于第二行，不把它当表头
                    If numericHeaderCount > numericDataCount Then headerLooksLikeHeader = False
                End If
                If headerLooksLikeHeader Then
                    tableData.Header.AddRange(headerRow)
                Else
                    tableData.Rows.Insert(0, headerRow)
                End If
            End If

            Return tableData
        End Function

        ''' <summary>
        ''' 将单行文本按可能的表格列拆分：优先按制表符，其次按2个及以上连续空格，再按逗号/顿号，最后兜底为单个单元格
        ''' </summary>
        Private Shared Function SplitTextIntoColumns(text As String) As String()
            If String.IsNullOrWhiteSpace(text) Then Return New String() {}

            Dim trimmed = text.Trim()

            ' 1) 已有制表符直接拆分
            If trimmed.Contains(vbTab) Then
                Return trimmed.Split({vbTab}, StringSplitOptions.RemoveEmptyEntries).Select(Function(c) c.Trim()).Where(Function(c) Not String.IsNullOrEmpty(c)).ToArray()
            End If

            ' 2) 按2个及以上连续空白拆分（OCR 表格常见对齐空白）
            Dim multiSpaceSplit = System.Text.RegularExpressions.Regex.Split(trimmed, "\s{2,}")
            Dim multiSpaceCandidates = multiSpaceSplit.Select(Function(c) c.Trim()).Where(Function(c) Not String.IsNullOrEmpty(c)).ToList()
            If multiSpaceCandidates.Count >= 2 Then
                Return multiSpaceCandidates.ToArray()
            End If

            ' 3) 按逗号/中文逗号拆分
            '    - 3列及以上直接视为表格
            '    - 2列时要求整行无空格，避免"你好，世界"这类普通句子被误拆
            Dim commaParts = System.Text.RegularExpressions.Regex.Split(trimmed, "[,，]")
            Dim commaCandidates = commaParts.Select(Function(c) c.Trim()).Where(Function(c) Not String.IsNullOrEmpty(c)).ToList()
            If commaCandidates.Count >= 3 OrElse (commaCandidates.Count = 2 AndAlso Not trimmed.Contains(" ")) Then
                Return commaCandidates.ToArray()
            End If

            ' 4) 兜底：单行文本作为一个单元格
            Return New String() {trimmed}
        End Function

        ''' <summary>
        ''' 表格重建专用拆分：只按制表符和 2 个及以上连续空格拆分，不处理单空格和普通逗号句子
        ''' 单空格拆分由 TryRebuildTableFromText 在上下文确定目标列数后单独决策
        ''' </summary>
        Private Shared Function SplitTextIntoColumnsForTable(text As String) As String()
            If String.IsNullOrWhiteSpace(text) Then Return New String() {}

            Dim trimmed = text.Trim()

            ' 1) 制表符直接拆分
            If trimmed.Contains(vbTab) Then
                Return trimmed.Split({vbTab}, StringSplitOptions.RemoveEmptyEntries).Select(Function(c) c.Trim()).Where(Function(c) Not String.IsNullOrEmpty(c)).ToArray()
            End If

            ' 2) 2 个及以上连续空白拆分
            Dim multiSpaceSplit = System.Text.RegularExpressions.Regex.Split(trimmed, "\s{2,}")
            Dim multiSpaceCandidates = multiSpaceSplit.Select(Function(c) c.Trim()).Where(Function(c) Not String.IsNullOrEmpty(c)).ToList()
            If multiSpaceCandidates.Count >= 2 Then
                Return multiSpaceCandidates.ToArray()
            End If

            ' 3) 兜底：单行文本作为一个单元格（单空格不在这里拆）
            Return New String() {trimmed}
        End Function

    End Class

End Namespace
