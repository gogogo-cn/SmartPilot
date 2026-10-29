' ============================================================================
' WordBridge - Word 宿主桥：读取文档/选区、写入文本与表格
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports Newtonsoft.Json
' 注意：不要再加 Imports ExcelAIVSTOAddIn.Utils —— 本工程未编译那个命名空间，
' 历史上加过这一行导致整个宿主工程编译不过。路径净化改用 HostPathHelper。

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class WordBridge
    Inherits HostBridgeBase

    Private _wordApp As Object

    Public Sub New(app As Object)
        MyBase.New(OfficeHostKind.Word, app)
        _wordApp = app
    End Sub

    ' ---------------- 文档信息 ----------------
    ' ---------------- 路径（宿主导出文件默认落点）----------------
    ' 基类 HostBridgeBase.GetDesktopPath 返回 "operation not supported on this host"，
    ' 导致前端拿不到当前用户桌面。这里按**当前登录用户**动态返回裸路径。
    <ComVisible(True)>
    Public Overrides Function GetDesktopPath() As String
        Return HostPathHelper.GetRealDesktop()
    End Function

    ' 基类没有 GetDownloadsPath（前端会先 typeof 探测再调），这里作为新成员提供
    <ComVisible(True)>
    Public Function GetDownloadsPath() As String
        Return HostPathHelper.GetRealDownloads()
    End Function

    <ComVisible(True)>
    Public Function GetDocumentInfo() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim name = Safe(CStr(doc.Name))
            Dim fullName = Safe(CStr(doc.FullName))
            Dim words = 0, pages = 0, paras = 0
            Try
                words = CInt(doc.Words.Count)
            Catch
            End Try
            Try
                paras = CInt(doc.Paragraphs.Count)
            Catch
            End Try
            Try
                pages = CInt(doc.ComputeStatistics(2))  ' wdStatisticPages = 2
            Catch
            End Try
            Return "{""success"":true,""name"":""" & name & """,""fullName"":""" & fullName & """," &
                   """words"":" & words & ",""paragraphs"":" & paras & ",""pages"":" & pages & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetActiveDocumentName() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return ""
            Return CStr(doc.Name)
        Catch
            Return ""
        End Try
    End Function

    ' ---------------- 读取 ----------------
    ''' <summary>去掉 Word 单元格结束符(Chr7)/回车等，得到可喂给 LLM 的纯文本。</summary>
    Private Function CleanWordText(t As String) As String
        If t Is Nothing Then Return ""
        Return t.Replace(ChrW(7), "").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Trim()
    End Function

    Private Function IsMeaningfulWordText(t As String) As Boolean
        Return CleanWordText(t).Length > 0
    End Function

    <ComVisible(True)>
    Public Function GetSelectionText() As String
        Try
            If _wordApp Is Nothing Then Return ""
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then Return ""
            Dim t As String = ""
            Try : t = CStr(sel.Text) : Catch : End Try
            ' 光标折叠或仅单元格结束符时，Selection.Text 为空 → 扩到当前单元格/段落，
            ' 否则「将选中的内容翻译」会读不到任何字（前端 documentText 空 → 诚实失败提示）。
            If Not IsMeaningfulWordText(t) Then
                Try
                    If Convert.ToInt32(sel.Information(12)) <> 0 Then
                        t = CStr(sel.Cells(1).Range.Text)
                    End If
                Catch
                End Try
            End If
            If Not IsMeaningfulWordText(t) Then
                Try : t = CStr(sel.Paragraphs(1).Range.Text) : Catch : End Try
            End If
            Return CleanWordText(t)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[WordBridge] GetSelectionText: " & ex.Message)
            Return ""
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetDocumentText(maxChars As Integer) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return ""
            Dim text = CStr(doc.Content.Text)
            If maxChars > 0 AndAlso text.Length > maxChars Then text = text.Substring(0, maxChars)
            Return text
        Catch ex As Exception
            Return ""
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetDocumentTextJson(maxChars As Integer) As String
        Dim t = GetDocumentText(maxChars)
        Return "{""success"":true,""text"":""" & Safe(t) & """}"
    End Function

    <ComVisible(True)>
    Public Function GetSelectionTextJson() As String
        Return "{""success"":true,""text"":""" & Safe(GetSelectionText()) & """}"
    End Function

    ' ---------------- 写入 ----------------
    <ComVisible(True)>
    Public Function InsertText(text As String) As String
        Try
            If _wordApp Is Nothing Then Return Err("Word 未就绪")
            _wordApp.Selection.InsertAfter(text)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ReplaceSelectionText(text As String) As String
        Try
            If _wordApp Is Nothing Then Return Err("Word 未就绪")
            _wordApp.Selection.Text = text
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function AppendToDocument(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            doc.Content.InsertAfter(text)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function SaveDocument(filePath As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            If String.IsNullOrWhiteSpace(filePath) Then
                doc.Save()
            Else
                doc.SaveAs(filePath)
            End If
            Return "{""success"":true,""fullName"":""" & Safe(CStr(doc.FullName)) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FindReplace(findText As String, replaceText As String, matchCase As Boolean) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            If String.IsNullOrEmpty(findText) Then Return Err("findText 不能为空")
            Dim find = doc.Content.Find
            find.ClearFormatting()
            find.Text = findText
            find.Replacement.Text = replaceText
            find.MatchCase = matchCase
            find.MatchWholeWord = False
            find.Forward = True
            find.Wrap = 1
            Dim miss = System.Type.Missing
            CallByName(find, "Execute", CallType.Method, miss, matchCase, False, False, False, False, True, 1, False, 2)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FormatSelection(bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object, style As String) As String
        Try
            If _wordApp Is Nothing Then Return Err("Word 未就绪")
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then Return Err("没有选中内容")
            If bold IsNot Nothing Then sel.Font.Bold = If(CBool(bold), -1, 0)
            If italic IsNot Nothing Then sel.Font.Italic = If(CBool(italic), -1, 0)
            If fontSize IsNot Nothing Then sel.Font.Size = CSng(fontSize)
            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then sel.Font.Name = CStr(fontName)
            If fontColor IsNot Nothing Then sel.Font.Color = OleColorFromArgb(CStr(fontColor))
            If style IsNot Nothing AndAlso style <> "" Then sel.Style = style
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' OleColorFromArgb 已上提到 HostBridgeBase 共享。

    <ComVisible(True)>
    Public Function InsertPicture(filePath As String, width As Object, height As Object) As String
        Try
            If _wordApp Is Nothing Then Return Err("Word 未就绪")
            If String.IsNullOrEmpty(filePath) Then Return Err("filePath 不能为空")
            Dim sel = _wordApp.Selection
            Dim pic = sel.InlineShapes.AddPicture(filePath, False, True)
            If width IsNot Nothing Then pic.Width = CSng(width)
            If height IsNot Nothing Then pic.Height = CSng(height)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function DeleteSelection() As String
        Try
            If _wordApp Is Nothing Then Return Err("Word 未就绪")
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then Return Err("没有选中内容")
            sel.Delete()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ExportPdf(filePath As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")

            If String.IsNullOrWhiteSpace(filePath) Then
                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                If String.IsNullOrWhiteSpace(desktop) Then desktop = System.IO.Path.GetTempPath()
                Dim baseName As String = "document"
                Try
                    Dim n As String = CStr(doc.Name)
                    If Not String.IsNullOrWhiteSpace(n) Then baseName = System.IO.Path.GetFileNameWithoutExtension(n)
                Catch
                End Try
                baseName = HostPathHelper.GetSafeFileName(baseName)
                If String.IsNullOrWhiteSpace(baseName) Then baseName = "document"
                Dim ts As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                filePath = System.IO.Path.Combine(desktop, baseName & "_" & ts & ".pdf")
            End If

            filePath = HostPathHelper.SanitizeUserOutputPath(filePath)

            If Not System.IO.Path.IsPathRooted(filePath) Then
                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                If String.IsNullOrWhiteSpace(desktop) Then desktop = System.IO.Path.GetTempPath()
                filePath = System.IO.Path.Combine(desktop, System.IO.Path.GetFileName(filePath))
            End If

            Dim dir As String = System.IO.Path.GetDirectoryName(filePath)
            If Not String.IsNullOrEmpty(dir) AndAlso Not System.IO.Directory.Exists(dir) Then
                System.IO.Directory.CreateDirectory(dir)
            End If

            doc.SaveAs(filePath, CObj(17))
            ' message 字段与 excel_export_csv 同构：tool-card 通用成功摘要会优先展示 data.message
            Return "{""success"":true,""path"":""" & Safe(filePath) & """,""message"":""成功导出到 PDF: " & Safe(filePath) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' ==================================================================
    ' 选中对象信息采集：让 AI 知道用户光标所在的文档表格，而非"无选区"
    ' ==================================================================
    <ComVisible(True)>
    Public Overrides Function GetSelectionInfo() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then
                Return JsonConvert.SerializeObject(New With {.success = False, .selectionType = "none", .summary = "Word 未就绪"})
            End If
            Dim sel = _wordApp.Selection
            If sel Is Nothing Then
                Return JsonConvert.SerializeObject(New With {.success = True, .selectionType = "none", .summary = "未选中任何对象"})
            End If
            Dim selectedText As String = ""
            Try : selectedText = CStr(sel.Text) : Catch : End Try

            Dim inTable = False
            Dim r = 0, c = 0, tableIndex = -1
            Dim tSelected As Object = Nothing
            ' 首选 wdWithInTable(=12)：光标/选区落在表格内即判定为表格。
            ' 旧实现用 sel.Range.Tables.Count，对"部分选区/未全含整表"会返回 0 → 误判为 text，
            ' 导致前端确定性硬短路不触发、退回 LLM 反问行列号（正是本次 Word 的症状）。
            Try
                If Convert.ToInt32(sel.Information(12)) <> 0 Then inTable = True
            Catch
            End Try
            If Not inTable Then
                Try
                    If Convert.ToInt32(sel.Range.Tables.Count) > 0 Then inTable = True
                Catch
                End Try
            End If
            If inTable Then
                ' 取所在表格：优先 Range.Tables(1)，再按选区范围在 doc.Tables 中定位（部分选区/合并单元格下更稳）
                Try : tSelected = sel.Range.Tables(1) : Catch : End Try
                If tSelected Is Nothing Then
                    Try
                        Dim s0 = Convert.ToInt32(sel.Range.Start)
                        Dim e0 = Convert.ToInt32(sel.Range.End)
                        For Each tb As Object In doc.Tables
                            Try
                                If s0 >= Convert.ToInt32(tb.Range.Start) AndAlso e0 <= Convert.ToInt32(tb.Range.End) Then
                                    tSelected = tb : Exit For
                                End If
                            Catch
                            End Try
                        Next
                    Catch
                    End Try
                End If
                If tSelected IsNot Nothing Then
                    Try : r = Convert.ToInt32(tSelected.Rows.Count) : Catch : End Try
                    Try : c = Convert.ToInt32(tSelected.Columns.Count) : Catch : End Try
                    tableIndex = TableIndexInDoc(tSelected)
                End If
            End If

            If inTable Then
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .selectionType = "table", .tableIndex = tableIndex,
                    .rowCount = r, .colCount = c, .selectedText = selectedText,
                    .summary = "光标位于文档第 " & tableIndex & " 个表格内（共 " & r & " 行 " & c & " 列）"})
            End If

            Return JsonConvert.SerializeObject(New With {
                .success = True, .selectionType = "text", .tableIndex = -1, .selectedText = selectedText,
                .summary = If(selectedText <> "", "选中文本：" & selectedText, "未选中表格或具体对象")})
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    Private Function TableIndexInDoc(t As Object) As Integer
        Try
            Dim doc = GetActiveDocument()
            Dim idx As Integer = 0
            For Each tb As Object In doc.Tables
                idx += 1
                If tb Is t Then Return idx
            Next
        Catch
        End Try
        Return -1
    End Function

    ''' <summary>
    ''' 格式化「光标所在表格」的某个单元格（无需 AI 再传 tableIndex）。
    ''' 若光标不在表格内，返回明确错误提示。
    '''
    ''' 2026-10-07 新增排版维度（整表模式下生效）：
    '''   borderStyle      边框版式 all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
    '''   headerFontColor  表头字体色（深底白字时用 #FFFFFF）
    '''   headerItalic     表头斜体
    '''   zebraColor       隔行底纹色（不给则由表头底色派生）
    '''   firstColBold     首列加粗（左侧标题列）
    '''   lastRowBold      末行（合计行）加粗 + 上边线
    '''   bodyAlign        正文对齐（覆盖整表 align，用于"表头居中、正文左对齐/数字右对齐"）
    '''   rowHeight        行高磅值（紧凑 16 / 标准 22 / 宽松 30）
    '''   valign           垂直对齐 middle 居中 | top 顶端
    ''' </summary>
    <ComVisible(True)>
    Public Function FormatSelectedTableCell(row As Integer, col As Integer, bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object, align As Object, fillColor As Object, borderColor As Object, headerBold As Object, headerFillColor As Object, clearFill As Object, highlightColor As Object, zebra As Object, Optional borderStyle As Object = Nothing, Optional headerFontColor As Object = Nothing, Optional zebraColor As Object = Nothing, Optional firstColBold As Object = Nothing, Optional lastRowBold As Object = Nothing, Optional bodyAlign As Object = Nothing, Optional rowHeight As Object = Nothing, Optional valign As Object = Nothing, Optional headerItalic As Object = Nothing) As String
        ' 屏幕刷新原值：必须声明在 Try 之外，否则 Finally 里访问不到（VB 的 Try/Catch/Finally 各自独立作用域）
        Dim _suPrev As Boolean = True
        Try
            CrashStep("FormatSelectedTableCell ENTER align=" & If(align Is Nothing, "∅", CStr(align)) &
                      " headerFill=" & If(headerFillColor Is Nothing, "∅", CStr(headerFillColor)) &
                      " border=" & If(borderColor Is Nothing, "∅", CStr(borderColor)) &
                      " zebra=" & If(zebra Is Nothing, "∅", CStr(zebra)))
            Dim doc = GetActiveDocument()
            CrashStep("after GetActiveDocument")
            If doc Is Nothing Then Return Err("没有打开的文档")
            If _wordApp.Selection Is Nothing Then Return Err("当前没有选区")
            Dim selTbl As Object = Nothing
            Try
                CrashStep("before Selection.Information(12)")
                If Convert.ToInt32(_wordApp.Selection.Information(12)) <> 0 Then
                    CrashStep("before Selection.Range.Tables(1)")
                    selTbl = _wordApp.Selection.Range.Tables(1)
                End If
            Catch
            End Try
            If selTbl Is Nothing Then
                Try
                    CrashStep("before Selection.Range.Tables.Count")
                    If Convert.ToInt32(_wordApp.Selection.Range.Tables.Count) > 0 Then selTbl = _wordApp.Selection.Range.Tables(1)
                Catch
                End Try
            End If
            If selTbl Is Nothing Then Return Err("当前光标不在表格内，请先把光标放到要设置的表格中")
            Dim table = selTbl
            Dim rCount As Integer = 0
            Dim cCount As Integer = 0
            ' 【勿碰 Columns.Count】纵向合并表上会 0xc0000005 直接杀进程（Try 无效）
            CrashStep("before Rows.Count")
            Try : rCount = CInt(table.Rows.Count) : Catch : rCount = 0 : End Try
            CrashStep("after Rows.Count=" & rCount)
            Dim tblIndex As Integer = -1
            ' 整表美化不再扫 doc.Tables 做索引比对（减少 COM 面）

            ' row/col 任一省略(<=0) → 整表生效
            If row <= 0 OrElse col <= 0 Then
                ' ════════════════════════════════════════════════════════════
                ' 【防崩溃·内置样式】系统日志：入口后立刻 wwlib!0xc0000005。
                ' 手改 Borders/Shading/Range.Font 均会触发；改为只应用 Word 内置表格样式。
                ' 面包屑：%LOCALAPPDATA%\ExcelAIVSTOAddIn\crash_last_step.txt
                ' ════════════════════════════════════════════════════════════
                Dim _hasClear As Boolean = (clearFill IsNot Nothing AndAlso CBool(clearFill))
                Dim _hasAny As Boolean = (headerBold IsNot Nothing) OrElse (headerFillColor IsNot Nothing) OrElse
                    (fillColor IsNot Nothing) OrElse (borderColor IsNot Nothing) OrElse (align IsNot Nothing) OrElse
                    (fontSize IsNot Nothing) OrElse (fontName IsNot Nothing) OrElse (fontColor IsNot Nothing) OrElse
                    (highlightColor IsNot Nothing) OrElse _hasClear OrElse (borderStyle IsNot Nothing) OrElse
                    (zebra IsNot Nothing) OrElse (zebraColor IsNot Nothing) OrElse (headerFontColor IsNot Nothing) OrElse
                    (bodyAlign IsNot Nothing) OrElse (rowHeight IsNot Nothing) OrElse (valign IsNot Nothing) OrElse
                    (firstColBold IsNot Nothing) OrElse (lastRowBold IsNot Nothing) OrElse (headerItalic IsNot Nothing) OrElse
                    (bold IsNot Nothing)
                If Not _hasAny Then
                    Return Err("未指定任何格式项（对齐/颜色/底纹/字体/边框/高亮/版式等），请说明要设置什么")
                End If

                Dim styleApplied As String = ""
                Dim parts As New List(Of String)

                If _hasClear Then
                    ' 清除：尽量用无样式
                    CrashStep("clear: set Style=Normal Table / 普通表格")
                    For Each sn As String In New String() {"普通表格", "Normal Table", "Table Normal"}
                        Try
                            table.Style = sn
                            styleApplied = sn
                            Exit For
                        Catch
                        End Try
                    Next
                    If styleApplied <> "" Then parts.Add("已清除表格样式→" & styleApplied)
                Else
                    ' 美化：优先带表头+隔行的内置样式（中英文名都试）
                    CrashStep("beautify: try built-in table styles")
                    Dim candidates As String() = {
                        "网格表 4 - 着色 1", "Grid Table 4 - Accent 1",
                        "网格表 5 深色 - 着色 1", "Grid Table 5 Dark - Accent 1",
                        "列表表 3 - 着色 1", "List Table 3 - Accent 1",
                        "网格型", "Table Grid", "表格网格"
                    }
                    For Each sn As String In candidates
                        Try
                            CrashStep("beautify: Style=" & sn)
                            table.Style = sn
                            styleApplied = sn
                            Exit For
                        Catch
                        End Try
                    Next
                    ' 样式失败时：只开 Outside/Inside 边框（不再碰 Shading / Rows / Range.Font）
                    If styleApplied = "" AndAlso borderColor IsNot Nothing Then
                        CrashStep("beautify: fallback Outside/Inside borders only")
                        Try
                            Dim bc = OleColorFromArgb(CStr(borderColor))
                            ApplyTableBorderStyle(table, bc, "all", rCount)
                            parts.Add("边框全网格→" & CStr(borderColor))
                        Catch exB As Exception
                            CrashStep("beautify: border fallback EX " & exB.Message)
                        End Try
                    End If
                    If styleApplied <> "" Then
                        parts.Add("应用表格样式→" & styleApplied)
                        ' 尽量打开表头行/隔行（属性不存在则跳过）
                        Try : table.ApplyStyleHeadingRows = True : Catch : End Try
                        Try : table.ApplyStyleRowBands = True : Catch : End Try
                        CrashStep("beautify: ApplyStyleHeadingRows/RowBands done")
                    End If
                End If

                CrashStep("FormatSelectedTableCell SUCCESS style=" & styleApplied & " parts=" & parts.Count)
                Dim summary As String = If(parts.Count = 0,
                    "表格（" & rCount & " 行）未能应用样式（文档可能受保护或样式库缺失）",
                    "表格（" & rCount & " 行）已美化：" & String.Join("、", parts.ToArray()))
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .scope = "wholeTable", .tableIndex = tblIndex,
                    .rowCount = rCount, .colCount = cCount, .changed = (parts.Count > 0),
                    .applied = New With {
                        .tableStyle = styleApplied,
                        .borderColor = If(borderColor Is Nothing, "", CStr(borderColor)),
                        .clearFill = _hasClear},
                    .changes = New Object() {},
                    .summary = summary})
            End If

            CrashStep("single-cell path r=" & row & " c=" & col)
            If rCount > 0 AndAlso row > rCount Then Return Err("行超出范围")
            ' 单格才取 Columns.Count；整表路径禁止碰它
            Try : cCount = CInt(table.Columns.Count) : Catch : cCount = 0 : End Try
            If cCount > 0 AndAlso col > cCount Then Return Err("列超出范围")
            Dim cell = table.Cell(row, col)
            Dim rng = CellTextRange(cell)
            If rng Is Nothing Then rng = cell.Range
            ' 防御：只给了 row/col 却一个格式项都没给 → 明确报错，避免"什么都没改却报成功"的静默假完成
            If bold Is Nothing AndAlso italic Is Nothing AndAlso fontSize Is Nothing AndAlso fontName Is Nothing AndAlso fontColor Is Nothing AndAlso align Is Nothing AndAlso fillColor Is Nothing AndAlso borderColor Is Nothing AndAlso highlightColor Is Nothing AndAlso Not (clearFill IsNot Nothing AndAlso CBool(clearFill)) Then
                Return Err("未指定任何格式项（对齐/字号/字体/颜色/底纹/边框/高亮等），请说明要设置什么格式；若要对整张表格设置，请省略行号/列号")
            End If
            Dim chg As New List(Of Object)
            If bold IsNot Nothing Then
                rng.Font.Bold = If(CBool(bold), -1, 0)
                chg.Add(New With {.param = "bold", .value = CStr(CBool(bold)), .cells = 1})
            End If
            If italic IsNot Nothing Then
                rng.Font.Italic = If(CBool(italic), -1, 0)
                chg.Add(New With {.param = "italic", .value = CStr(CBool(italic)), .cells = 1})
            End If
            If fontSize IsNot Nothing Then
                rng.Font.Size = CSng(fontSize)
                chg.Add(New With {.param = "fontSize", .value = CStr(fontSize), .cells = 1})
            End If
            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then
                rng.Font.Name = CStr(fontName)
                chg.Add(New With {.param = "fontName", .value = CStr(fontName), .cells = 1})
            End If
            If fontColor IsNot Nothing Then
                rng.Font.Color = OleColorFromArgb(CStr(fontColor))
                chg.Add(New With {.param = "fontColor", .value = CStr(fontColor), .cells = 1})
            End If
            If highlightColor IsNot Nothing Then
                Dim _hIdx As Object = MapHighlightColor(CStr(highlightColor))
                If _hIdx IsNot Nothing Then
                    rng.HighlightColorIndex = _hIdx
                    chg.Add(New With {.param = "highlightColor", .value = CStr(highlightColor), .cells = 1})
                End If
            End If
            If align IsNot Nothing Then
                rng.ParagraphFormat.Alignment = MapAlign(CStr(align))
                chg.Add(New With {.param = "align", .value = CStr(align), .cells = 1})
            End If
            If fillColor IsNot Nothing Then
                SetCellFill(cell, OleColorFromArgb(CStr(fillColor)))
                chg.Add(New With {.param = "fillColor", .value = CStr(fillColor), .cells = 1})
            End If
            If clearFill IsNot Nothing AndAlso CBool(clearFill) Then
                Dim had As Boolean = False
                Try : had = ClearCellFill(cell.Range) : Catch : End Try
                chg.Add(New With {.param = "clearFill", .value = "True", .cells = If(had, 1, 0)})
            End If
            If borderColor IsNot Nothing Then
                Dim bc = OleColorFromArgb(CStr(borderColor))
                For Each b As Object In cell.Borders
                    Dim _bt As Integer = 0
                    Try : _bt = Convert.ToInt32(b.Type) : Catch : End Try
                    If _bt = 7 OrElse _bt = 8 Then Continue For
                    b.LineStyle = 1                              ' wdLineStyleSingle
                    b.LineWidth = 4                              ' wdLineWidth050pt
                    b.Color = bc                                 ' Word：Color=Long，勿写 Color.RGB
                Next
                chg.Add(New With {.param = "borderColor", .value = CStr(borderColor), .cells = 1})
            End If
            Dim summaryC As String = "表格#" & tblIndex & " 第" & row & "行第" & col & "列：已应用 " & chg.Count & " 项格式"
            LogBridge("FormatSelectedTableCell 成功(单格) r=" & row & " c=" & col & " clearFill=" & If(clearFill Is Nothing, "∅", CStr(clearFill)) & " items=" & chg.Count)
            Return JsonConvert.SerializeObject(New With {
                .success = True, .scope = "cell", .tableIndex = tblIndex,
                .row = row, .col = col,
                .changes = chg, .summary = summaryC})
        Catch ex As Exception
            Return Err(ex.Message)
        Finally
            ' 无论成功/异常/提前 Return，都恢复屏幕刷新，避免把 Word 卡在"界面不刷新"的状态
            Try : _wordApp.ScreenUpdating = _suPrev : Catch : End Try
        End Try
    End Function

    ''' <summary>
    ''' 按版式给整张表画边框（Word 后期绑定）。
    ''' style: all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
    ''' 【防崩溃】美化默认 all：用 Borders.Outside/Inside* 整表属性，禁止先清后逐边索引
    ''' （逐边 Borders(i)+先 LineStyle=0 在合并单元格/复杂表上会让 Word 进程强退）。
    ''' Color 必须赋 Long，绝不能 Color.RGB（PPT API）。
    ''' </summary>
    Private Sub ApplyTableBorderStyle(table As Object, bc As Object, style As Object, rCount As Integer)
        Dim st As String = If(style Is Nothing, "all", LCase(CStr(style).Trim()))
        If st = "" Then st = "all"
        Try
            Dim borders As Object = table.Borders
            If st = "none" Then
                Try : borders.OutsideLineStyle = 0 : Catch : End Try
                Try : borders.InsideLineStyle = 0 : Catch : End Try
                Return
            End If
            If st = "all" OrElse st = "outer" Then
                Try : borders.OutsideLineStyle = 1 : Catch : End Try
                Try : borders.OutsideLineWidth = 4 : Catch : End Try
                Try : borders.OutsideColor = bc : Catch : End Try
                If st = "all" Then
                    Try : borders.InsideLineStyle = 1 : Catch : End Try
                    Try : borders.InsideLineWidth = 4 : Catch : End Try
                    Try : borders.InsideColor = bc : Catch : End Try
                Else
                    Try : borders.InsideLineStyle = 0 : Catch : End Try
                End If
                Return
            End If
            If st = "horizontal" Then
                ' 仅横线：关竖线，开顶/底/内部横线（用 Outside + InsideHorizontal）
                Try : borders.OutsideLineStyle = 1 : Catch : End Try
                Try : borders.OutsideLineWidth = 4 : Catch : End Try
                Try : borders.OutsideColor = bc : Catch : End Try
                Try : borders.InsideLineStyle = 0 : Catch : End Try
                Try
                    Dim ih = borders(-5)   ' wdBorderHorizontal
                    ih.LineStyle = 1 : ih.LineWidth = 4 : ih.Color = bc
                Catch
                End Try
                ' 去掉左右外框，只留上下外框
                Try : borders(-2).LineStyle = 0 : Catch : End Try
                Try : borders(-4).LineStyle = 0 : Catch : End Try
                Return
            End If
            If st = "threeline" Then
                Try : borders.OutsideLineStyle = 0 : Catch : End Try
                Try : borders.InsideLineStyle = 0 : Catch : End Try
                Try
                    borders(-1).LineStyle = 1 : borders(-1).LineWidth = 6 : borders(-1).Color = bc
                Catch
                End Try
                Try
                    borders(-3).LineStyle = 1 : borders(-3).LineWidth = 6 : borders(-3).Color = bc
                Catch
                End Try
                If rCount >= 1 Then
                    Try
                        table.Rows(1).Borders(-3).LineStyle = 1
                        table.Rows(1).Borders(-3).LineWidth = 6
                        table.Rows(1).Borders(-3).Color = bc
                    Catch
                    End Try
                End If
                Return
            End If
            ' 未知 → 全网格
            Try : borders.OutsideLineStyle = 1 : borders.OutsideLineWidth = 4 : borders.OutsideColor = bc : Catch : End Try
            Try : borders.InsideLineStyle = 1 : borders.InsideLineWidth = 4 : borders.InsideColor = bc : Catch : End Try
        Catch
        End Try
    End Sub

    ''' <summary>将 #RRGGBB 或中文色名映射为 Word 高亮调色板(WdColorIndex)整数；无法识别返回 Nothing。</summary>
    Private Function MapHighlightColor(s As String) As Object
        If s Is Nothing Then Return Nothing
        Dim t = LCase(Trim(s))
        If t = "" OrElse t = "none" OrElse t = "无" Then Return Nothing
        Select Case t
            Case "黄", "黄色", "黄绿" : Return 7          ' wdYellow
            Case "亮绿", "绿", "绿色" : Return 4          ' wdBrightGreen
            Case "青", "青色", "蓝绿" : Return 3          ' wdTurquoise
            Case "粉", "粉色", "粉红" : Return 5          ' wdPink
            Case "蓝", "蓝色" : Return 2                  ' wdBlue
            Case "红", "红色" : Return 6                  ' wdRed
            Case "紫", "紫色" : Return 12                 ' wdViolet
            Case "深蓝" : Return 9                        ' wdDarkBlue
            Case "深红" : Return 13                       ' wdDarkRed
            Case "深黄" : Return 14                       ' wdDarkYellow
            Case "灰", "灰色", "灰50" : Return 15         ' wdGray50
            Case "浅灰", "灰25" : Return 16               ' wdGray25
            Case "黑", "黑色" : Return 1                  ' wdBlack
            Case "白", "白色" : Return 8                  ' wdWhite
        End Select
        ' #RRGGBB → 取最近高亮调色板色
        Dim m = System.Text.RegularExpressions.Regex.Match(t, "#([0-9a-f]{6})")
        If m.Success Then
            Dim hex = m.Groups(1).Value
            Dim r = Convert.ToInt32(hex.Substring(0, 2), 16)
            Dim g = Convert.ToInt32(hex.Substring(2, 2), 16)
            Dim b = Convert.ToInt32(hex.Substring(4, 2), 16)
            Dim palette = New Integer()() {
                New Integer() {7, 255, 255, 0}, New Integer() {4, 0, 255, 0}, New Integer() {3, 0, 255, 255},
                New Integer() {5, 255, 0, 255}, New Integer() {2, 0, 0, 255}, New Integer() {6, 255, 0, 0},
                New Integer() {12, 128, 0, 128}, New Integer() {9, 0, 0, 128}, New Integer() {13, 128, 0, 0},
                New Integer() {14, 128, 128, 0}, New Integer() {15, 128, 128, 128}, New Integer() {16, 192, 192, 192},
                New Integer() {1, 0, 0, 0}, New Integer() {8, 255, 255, 255}}
            Dim best = -1, bestD = Integer.MaxValue
            For Each p As Integer() In palette
                Dim d = (p(1) - r) * (p(1) - r) + (p(2) - g) * (p(2) - g) + (p(3) - b) * (p(3) - b)
                If d < bestD Then bestD = d : best = p(0)
            Next
            Return best
        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' 取单元格 Range（不再 SetRange/MoveEnd：收缩 Range 在合并格/复杂表上会让 Word 强退）。
    ''' </summary>
    Private Function CellTextRange(cell As Object) As Object
        If cell Is Nothing Then Return Nothing
        Try
            Return cell.Range
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 安全设置单元格底纹：只碰 Cell.Shading，异常则放弃（绝不因底纹拖崩 Word）。
    ''' </summary>
    Private Sub SetCellFill(cell As Object, oleColor As Integer)
        If cell Is Nothing Then Return
        Try
            cell.Shading.BackgroundPatternColor = oleColor
        Catch
            ' 合并单元格/受保护文档上 Shading 可能不可写——跳过，勿二次折腾 Range
        End Try
    End Sub

    ''' <summary>
    ''' 彻底清除单元格"底纹"：覆盖 Word 中三种可能的实现——表格底纹(Shading)、
    ''' 文字突出显示(HighlightColorIndex，即"荧光笔"高亮)、字体底纹(Font.Shading)。
    ''' 之前只清 Shading.BackgroundPatternColor，若用户用的是"突出显示"则清不掉（表现为"清了底纹还在"）。
    ''' </summary>
    Private Function ClearCellFill(rng As Object) As Boolean
        If rng Is Nothing Then Return False
        ' 先探测"原本是否有可见底纹/高亮"，用于向用户回传"实际影响了多少格"（原本就无底纹的不算变化）
        Dim had As Boolean = False
        Try
            Dim bg = Convert.ToInt32(rng.Shading.BackgroundPatternColor)
            If bg <> -16777216 AndAlso bg <> 9999999 Then had = True   ' 非 wdColorAutomatic / 非 wdUndefined
        Catch
        End Try
        Try
            Dim hl = Convert.ToInt32(rng.HighlightColorIndex)
            If hl <> 0 AndAlso hl <> 9999999 Then had = True           ' 非 wdNoHighlight
        Catch
        End Try
        ' 【性能】原本就没有任何可见底纹/高亮 → 直接返回，省掉每格最多 8 次 COM 写入。
        ' 整表"消除底纹"时绝大多数格子本就无底纹，此跳过能带来数量级的提速。
        If Not had Then Return False
        Try : rng.Shading.Texture = 0 : Catch : End Try                        ' wdTextureNone
        Try : rng.Shading.BackgroundPatternColor = -16777216 : Catch : End Try ' wdColorAutomatic（无填充）
        Try : rng.Shading.ForegroundPatternColor = -16777216 : Catch : End Try
        Try : rng.HighlightColorIndex = 0 : Catch : End Try                    ' wdNoHighlight
        Try : rng.Font.Shading.Texture = 0 : Catch : End Try
        Try : rng.Font.Shading.BackgroundPatternColor = -16777216 : Catch : End Try
        Try : rng.Font.Shading.ForegroundPatternColor = -16777216 : Catch : End Try
        Return had
    End Function

    Public Function FormatTableCell(tableIndex As Integer, row As Integer, col As Integer, bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object, align As Object, fillColor As Object, borderColor As Object, highlightColor As Object) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            If row <= 0 OrElse col <= 0 Then Return Err("row/col 必须从 1 开始")
            If CInt(doc.Tables.Count) = 0 Then Return Err("文档中没有表格")
            Dim idx = If(tableIndex <= 0, 1, tableIndex)
            If idx > CInt(doc.Tables.Count) Then Return Err("表格序号超出范围")
            Dim table = doc.Tables(idx)
            Dim cell = table.Cell(row, col)
            Dim rng = cell.Range
            If bold IsNot Nothing Then rng.Font.Bold = If(CBool(bold), -1, 0)
            If italic IsNot Nothing Then rng.Font.Italic = If(CBool(italic), -1, 0)
            If fontSize IsNot Nothing Then rng.Font.Size = CSng(fontSize)
            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then rng.Font.Name = CStr(fontName)
            If fontColor IsNot Nothing Then rng.Font.Color = OleColorFromArgb(CStr(fontColor))
            If highlightColor IsNot Nothing Then
                Dim _hIdx As Object = MapHighlightColor(CStr(highlightColor))
                If _hIdx IsNot Nothing Then rng.HighlightColorIndex = _hIdx
            End If
            If align IsNot Nothing Then rng.ParagraphFormat.Alignment = MapAlign(CStr(align))
            If fillColor IsNot Nothing Then SetCellFill(cell, OleColorFromArgb(CStr(fillColor)))
            If borderColor IsNot Nothing Then
                Dim bc = OleColorFromArgb(CStr(borderColor))
                For Each b As Object In cell.Borders
                    Dim _bt As Integer = 0
                    Try : _bt = Convert.ToInt32(b.Type) : Catch : End Try
                    If _bt = 7 OrElse _bt = 8 Then Continue For
                    Try
                        b.LineStyle = 1
                        b.LineWidth = 4
                        b.Color = bc    ' Word：Color=Long，勿写 Color.RGB（会强退）
                    Catch
                    End Try
                Next
            End If
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ''' <summary>Word 对齐枚举 WdParagraphAlignment：左=0 中=1 右=2（与 PowerPoint 不同，必须按本宿主枚举）。</summary>
    Protected Overrides Function MapAlign(s As String) As Integer
        If s Is Nothing Then Return 1
        Select Case LCase(s)
            Case "left" : Return 0
            Case "right" : Return 2
            Case Else : Return 1
        End Select
    End Function

    <ComVisible(True)>
    Public Function InsertPageBreak() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            doc.Application.Selection.InsertBreak(7)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertHyperlink(address As String, text As String, displayText As Object) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim anchor = doc.Application.Selection.Range
            Dim disp = If(displayText Is Nothing, text, CStr(displayText))
            doc.Hyperlinks.Add(anchor, address, Nothing, Nothing, disp)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertToc() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim range = doc.Application.Selection.Range
            doc.TablesOfContents.Add(range, True, 1, 3)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertSectionBreak(breakType As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim bt = 2
            If breakType IsNot Nothing AndAlso LCase(CStr(breakType)) = "continuous" Then bt = 3
            doc.Application.Selection.InsertBreak(bt)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertComment(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim range = doc.Application.Selection.Range
            doc.Comments.Add(range, text)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertBookmark(name As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            If String.IsNullOrWhiteSpace(name) Then Return Err("书签名必填")
            Dim range = doc.Application.Selection.Range
            doc.Bookmarks.Add(name, range)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertFootnote(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim range = doc.Application.Selection.Range
            doc.Footnotes.Add(range, Nothing, text)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertEndnote(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim range = doc.Application.Selection.Range
            doc.Endnotes.Add(range, Nothing, text)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ApplyNumbering() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            doc.Application.Selection.Range.ListFormat.ApplyNumberDefault()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ApplyBullets() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            doc.Application.Selection.Range.ListFormat.ApplyBulletDefault()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function InsertTableData(rowsJson As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim data = ParseRows(rowsJson)
            If data Is Nothing Then Return Err("表格数据为空或格式不正确")
            Dim rowCount = data.GetLength(1)
            Dim colCount = data.GetLength(0)
            Dim sel = _wordApp.Selection
            Dim range = sel.Range
            Dim table = doc.Tables.Add(range, rowCount, colCount)
            For r As Integer = 1 To rowCount
                For c As Integer = 1 To colCount
                    table.Cell(r, c).Range.Text = If(data(c - 1, r - 1), "")
                Next
            Next
            Return "{""success"":true,""rows"":" & rowCount & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function InsertParsedData(rowsJson As String) As String
        Return InsertTableData(rowsJson)
    End Function

    <ComVisible(True)>
    Public Overrides Function GetSelectedRangeData() As String
        Try
            ' 必须用 JsonConvert：旧手拼 JSON 未给字符串加引号，前端 JSON.parse 恒失败 → documentText 永不注入
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .host = "word",
                .document = GetActiveDocumentName(),
                .selectedText = GetSelectionText(),
                .documentText = GetDocumentText(20000)
            })
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function GetSelection() As String
        Try
            Return GetSelectionText()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function GetMultiSelection() As String
        Try
            Return GetSelectionText()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertTextLines(linesJson As String) As String
        Try
            Dim rows = ParseRows(linesJson)
            If rows Is Nothing Then Return Err("内容为空")
            Dim sb As New StringBuilder()
            For r As Integer = 0 To rows.GetLength(1) - 1
                For c As Integer = 0 To rows.GetLength(0) - 1
                    sb.Append(If(rows(c, r), ""))
                    If c < rows.GetLength(0) - 1 Then sb.Append(vbTab)
                Next
                sb.Append(vbCrLf)
            Next
            Return InsertText(sb.ToString())
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' ---------------- 页面 / 页眉页脚 / 域 / 查找 ----------------
    <ComVisible(True)>
    Public Function SetPageOrientation(orientation As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim o = 1  ' wdOrientPortrait = 1
            If orientation IsNot Nothing AndAlso LCase(CStr(orientation)) = "landscape" Then o = 2  ' wdOrientLandscape = 2
            doc.PageSetup.Orientation = o
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertHeaderText(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim sec = doc.Sections(1)
            Dim h = sec.Headers(1)  ' wdHeaderFooterPrimary = 1
            h.Range.Text = If(text, "")
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function InsertFooterText(text As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            Dim sec = doc.Sections(1)
            Dim f = sec.Footers(1)  ' wdHeaderFooterPrimary = 1
            f.Range.Text = If(text, "")
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function UpdateFields() As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            doc.Fields.Update()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FindText(findTextValue As String) As String
        Try
            Dim doc = GetActiveDocument()
            If doc Is Nothing Then Return Err("没有打开的文档")
            If String.IsNullOrEmpty(findTextValue) Then Return Err("findText 不能为空")
            Dim rng = doc.Content
            Dim find = rng.Find
            find.ClearFormatting()
            find.Text = findTextValue
            find.Forward = True
            find.Wrap = 0  ' wdFindStop = 0
            Dim count As Integer = 0
            Dim miss = System.Type.Missing
            While CBool(CallByName(find, "Execute", CallType.Method, miss, False, False, False, False, False, True, 1, False, 2))
                count += 1
                If count >= 1000 Then Exit While
            End While
            Return "{""success"":true,""count"":" & count & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' ---------------- 内部 ----------------
    Private Function GetActiveDocument() As Object
        Try
            If _wordApp Is Nothing Then Return Nothing
            Return _wordApp.ActiveDocument
        Catch
            Return Nothing
        End Try
    End Function
End Class