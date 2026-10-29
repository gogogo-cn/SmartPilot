Option Strict On
Imports Newtonsoft.Json
Imports Microsoft.Office.Interop.Excel

Namespace Utils

    ''' <summary>
    ''' Excel Markdown 渲染器 - 将 Markdown 文本渲染到 Excel 工作表
    ''' </summary>
    Public Class ExcelMarkdownRenderer

        Private ReadOnly _excelApp As Microsoft.Office.Interop.Excel.Application

        Public Sub New(excelApp As Microsoft.Office.Interop.Excel.Application)
            _excelApp = excelApp
        End Sub

        ''' <summary>
        ''' 将Markdown文本渲染到Excel，每行文本独立成行显示
        ''' </summary>
        ''' <param name="markdownText">Markdown格式文本</param>
        ''' <param name="startCell">起始单元格地址，如"A1"</param>
        ''' <returns>操作结果JSON</returns>
        Public Function RenderFormattedText(markdownText As String, Optional startCell As String = "") As String
            Try
                If _excelApp Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "Excel应用不可用"})
                End If

                If String.IsNullOrWhiteSpace(markdownText) Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "文本内容不能为空"})
                End If

                Dim targetCell As Range = Nothing
                If String.IsNullOrWhiteSpace(startCell) Then
                    targetCell = _excelApp.ActiveCell
                Else
                    targetCell = DirectCast(_excelApp.Range(startCell), Range)
                End If

                If targetCell Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {.success = False, .error = "无效的起始单元格"})
                End If

                Dim sheet As Worksheet = DirectCast(_excelApp.ActiveSheet, Worksheet)
                Dim startRow As Integer = CInt(targetCell.Row)
                Dim startCol As Integer = CInt(targetCell.Column)

                Dim colRange As Range = DirectCast(sheet.Columns(startCol), Range)
                colRange.ColumnWidth = 100

                Dim currentRow As Integer = startRow

                Dim lines As List(Of String) = markdownText.Split(New String() {Environment.NewLine}, StringSplitOptions.None).ToList()
                Dim lineIndex As Integer = 0

                While lineIndex < lines.Count
                    Dim line As String = lines(lineIndex)
                    Dim trimmedLine As String = line.TrimStart()

                    If String.IsNullOrEmpty(trimmedLine) Then
                        currentRow += 1
                        lineIndex += 1
                        Continue While
                    End If

                    If trimmedLine.StartsWith("|") AndAlso trimmedLine.EndsWith("|") Then
                        Dim tableLines As New List(Of String)()
                        tableLines.Add(trimmedLine)

                        Dim nextIndex As Integer = lineIndex + 1
                        While nextIndex < lines.Count
                            Dim nextLine As String = lines(nextIndex).TrimStart()
                            If nextLine.StartsWith("|") AndAlso nextLine.EndsWith("|") Then
                                tableLines.Add(nextLine)
                                nextIndex += 1
                            Else
                                Exit While
                            End If
                        End While

                        If tableLines.Count > 0 Then
                            Dim tableEndRow As Integer = RenderMarkdownTable(sheet, tableLines, currentRow, startCol)
                            currentRow = tableEndRow + 2
                            lineIndex = nextIndex
                            Continue While
                        End If
                    End If

                    If trimmedLine.StartsWith("### ") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        cell.Value = trimmedLine.Substring(4)
                        cell.Font.Bold = True
                        cell.Font.Size = 13
                        cell.Font.Color = CInt(RGB(0, 0, 0))
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 2
                        lineIndex += 1
                        Continue While
                    ElseIf trimmedLine.StartsWith("## ") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        cell.Value = trimmedLine.Substring(3)
                        cell.Font.Bold = True
                        cell.Font.Size = 14
                        cell.Font.Color = CInt(RGB(0, 0, 0))
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 2
                        lineIndex += 1
                        Continue While
                    ElseIf trimmedLine.StartsWith("# ") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        cell.Value = trimmedLine.Substring(2)
                        cell.Font.Bold = True
                        cell.Font.Size = 15
                        cell.Font.Color = CInt(RGB(0, 0, 0))
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 2
                        lineIndex += 1
                        Continue While
                    End If

                    If trimmedLine.StartsWith("---") OrElse trimmedLine.StartsWith("***") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        cell.Value = ""
                        Dim bottomBorder As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeBottom), Border)
                        bottomBorder.LineStyle = XlLineStyle.xlContinuous
                        bottomBorder.Weight = XlBorderWeight.xlThin
                        bottomBorder.Color = CInt(RGB(200, 200, 200))
                        currentRow += 2
                        lineIndex += 1
                        Continue While
                    End If

                    If trimmedLine.StartsWith("- ") OrElse trimmedLine.StartsWith("* ") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        Dim indent As String = New String(" "c, line.Length - line.TrimStart().Length)
                        cell.Value = indent & "• " & ProcessInlineFormatting(trimmedLine.Substring(2), cell)
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 1
                        lineIndex += 1
                        Continue While
                    End If

                    If System.Text.RegularExpressions.Regex.IsMatch(trimmedLine, "^\d+\.\s") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        Dim indent As String = New String(" "c, line.Length - line.TrimStart().Length)
                        Dim match As System.Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(trimmedLine, "^(\d+\.\s)(.*)$")
                        If match.Success Then
                            cell.Value = indent & match.Groups(1).Value & ProcessInlineFormatting(match.Groups(2).Value, cell)
                        End If
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 1
                        lineIndex += 1
                        Continue While
                    End If

                    If trimmedLine.StartsWith("> ") Then
                        Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                        Dim indent As String = New String(" "c, line.Length - line.TrimStart().Length)
                        cell.Value = indent & ProcessInlineFormatting(trimmedLine.Substring(2), cell)
                        cell.Font.Italic = True
                        cell.Font.Color = CInt(RGB(100, 100, 100))
                        Dim leftBorder As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeLeft), Border)
                        leftBorder.LineStyle = XlLineStyle.xlContinuous
                        leftBorder.Weight = XlBorderWeight.xlMedium
                        leftBorder.Color = CInt(RGB(0, 120, 212))
                        cell.VerticalAlignment = XlVAlign.xlVAlignTop
                        currentRow += 1
                        lineIndex += 1
                        Continue While
                    End If

                    Dim paraCell As Range = DirectCast(sheet.Cells(currentRow, startCol), Range)
                    Dim paraIndent As String = New String(" "c, line.Length - line.TrimStart().Length)
                    paraCell.Value = paraIndent & ProcessInlineFormatting(trimmedLine, paraCell)
                    paraCell.VerticalAlignment = XlVAlign.xlVAlignTop
                    currentRow += 1
                    lineIndex += 1
                End While

                For r As Integer = startRow To currentRow - 1
                    Try
                        Dim rowRange As Range = DirectCast(sheet.Rows(r), Range)
                        rowRange.AutoFit()
                    Catch ex As Exception  ' 
                    End Try
                Next

                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .message = "文本已渲染完成",
                    .rowsRendered = currentRow - startRow
                })

            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("RenderFormattedText 错误: " & ex.Message)
                System.Diagnostics.Debug.WriteLine("堆栈: " & ex.StackTrace)
                Return JsonConvert.SerializeObject(New With {.success = False, .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 渲染Markdown表格到Excel
        ''' </summary>
        ''' <param name="sheet">工作表</param>
        ''' <param name="tableLines">表格行列表</param>
        ''' <param name="startRow">起始行</param>
        ''' <param name="startCol">起始列</param>
        ''' <returns>表格结束行号</returns>
        Private Function RenderMarkdownTable(sheet As Worksheet, tableLines As List(Of String), startRow As Integer, startCol As Integer) As Integer
            If tableLines Is Nothing OrElse tableLines.Count = 0 Then Return startRow

            Dim currentRow As Integer = startRow

            Dim separatorIndex As Integer = -1
            For i As Integer = 0 To tableLines.Count - 1
                Dim line As String = tableLines(i).Trim()
                Dim separatorContent As String = line.Replace("|", "").Replace("-", "").Replace(":", "").Replace(" ", "")
                If String.IsNullOrEmpty(separatorContent) AndAlso line.Contains("-") Then
                    separatorIndex = i
                    Exit For
                End If
            Next

            For i As Integer = 0 To tableLines.Count - 1
                Dim line As String = tableLines(i).Trim()

                If i = separatorIndex Then
                    Continue For
                End If

                Dim cells As List(Of String) = line.Split("|"c).ToList()
                If cells.Count > 0 AndAlso String.IsNullOrWhiteSpace(cells(0)) Then
                    cells.RemoveAt(0)
                End If
                If cells.Count > 0 AndAlso String.IsNullOrWhiteSpace(cells(cells.Count - 1)) Then
                    cells.RemoveAt(cells.Count - 1)
                End If

                For j As Integer = 0 To cells.Count - 1
                    cells(j) = cells(j).Trim()
                Next

                If cells.Count = 0 Then Continue For

                For colIndex As Integer = 0 To cells.Count - 1
                    Dim cell As Range = DirectCast(sheet.Cells(currentRow, startCol + colIndex), Range)
                    cell.Value = ProcessInlineFormatting(cells(colIndex), cell)

                    If separatorIndex > 0 AndAlso i < separatorIndex Then
                        cell.Font.Bold = True
                        cell.Interior.Color = CInt(RGB(240, 240, 240))
                        cell.HorizontalAlignment = XlHAlign.xlHAlignCenter
                    End If

                    Dim leftB As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeLeft), Border)
                    leftB.LineStyle = XlLineStyle.xlContinuous
                    leftB.Weight = XlBorderWeight.xlThin
                    leftB.Color = CInt(RGB(200, 200, 200))

                    Dim rightB As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeRight), Border)
                    rightB.LineStyle = XlLineStyle.xlContinuous
                    rightB.Weight = XlBorderWeight.xlThin
                    rightB.Color = CInt(RGB(200, 200, 200))

                    Dim topB As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeTop), Border)
                    topB.LineStyle = XlLineStyle.xlContinuous
                    topB.Weight = XlBorderWeight.xlThin
                    topB.Color = CInt(RGB(200, 200, 200))

                    Dim bottomB As Border = DirectCast(cell.Borders(XlBordersIndex.xlEdgeBottom), Border)
                    bottomB.LineStyle = XlLineStyle.xlContinuous
                    bottomB.Weight = XlBorderWeight.xlThin
                    bottomB.Color = CInt(RGB(200, 200, 200))
                Next

                currentRow += 1
            Next

            If tableLines.Count > 0 Then
                Dim firstLine As String = tableLines(0).Trim()
                Dim firstCells As List(Of String) = firstLine.Split("|"c).ToList()
                If firstCells.Count > 0 AndAlso String.IsNullOrWhiteSpace(firstCells(0)) Then
                    firstCells.RemoveAt(0)
                End If
                If firstCells.Count > 0 AndAlso String.IsNullOrWhiteSpace(firstCells(firstCells.Count - 1)) Then
                    firstCells.RemoveAt(firstCells.Count - 1)
                End If

                For colIndex As Integer = 0 To firstCells.Count - 1
                    Dim colRange As Range = DirectCast(sheet.Columns(startCol + colIndex), Range)
                    colRange.AutoFit()
                Next
            End If

            Return currentRow - 1
        End Function

        Private Function ProcessInlineFormatting(text As String, cell As Range) As String
            If String.IsNullOrEmpty(text) Then Return text

            Dim result As String = text

            Dim boldItalicPattern As String = "\*\*\*(.+?)\*\*\*"
            Dim boldItalicMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, boldItalicPattern)
            For Each match As System.Text.RegularExpressions.Match In boldItalicMatches
                Dim formattedText As String = match.Groups(1).Value
                result = result.Replace(match.Value, formattedText)
                cell.Font.Bold = True
                cell.Font.Italic = True
            Next

            Dim boldPattern As String = "\*\*(.+?)\*\*"
            Dim boldMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, boldPattern)
            For Each match As System.Text.RegularExpressions.Match In boldMatches
                Dim formattedText As String = match.Groups(1).Value
                result = result.Replace(match.Value, formattedText)
                cell.Font.Bold = True
            Next

            Dim italicPattern As String = "(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|_(.+?)_"
            Dim italicMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, italicPattern)
            For Each match As System.Text.RegularExpressions.Match In italicMatches
                Dim formattedText As String = If(String.IsNullOrEmpty(match.Groups(1).Value), match.Groups(2).Value, match.Groups(1).Value)
                result = result.Replace(match.Value, formattedText)
                cell.Font.Italic = True
            Next

            Dim codePattern As String = "`(.+?)`"
            Dim codeMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, codePattern)
            For Each match As System.Text.RegularExpressions.Match In codeMatches
                Dim formattedText As String = match.Groups(1).Value
                result = result.Replace(match.Value, formattedText)
                cell.Font.Name = "Consolas"
                cell.Font.Color = CInt(RGB(220, 20, 60))
            Next

            Dim strikethroughPattern As String = "~~(.+?)~~"
            Dim strikethroughMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, strikethroughPattern)
            For Each match As System.Text.RegularExpressions.Match In strikethroughMatches
                Dim formattedText As String = match.Groups(1).Value
                result = result.Replace(match.Value, formattedText)
                cell.Font.Strikethrough = True
            Next

            Dim highlightPattern As String = "==(.+?)=="
            Dim highlightMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(result, highlightPattern)
            For Each match As System.Text.RegularExpressions.Match In highlightMatches
                Dim formattedText As String = match.Groups(1).Value
                result = result.Replace(match.Value, formattedText)
                cell.Interior.Color = CInt(RGB(255, 255, 0))
            Next

            Return result
        End Function

    End Class

End Namespace