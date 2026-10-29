Option Strict Off
Option Explicit On

Imports System
Imports System.Runtime.InteropServices
Imports System.Text
Imports Newtonsoft.Json
' 注意：不要再加 Imports ExcelAIVSTOAddIn.Utils —— 本工程未编译那个命名空间，
' 历史上加过这一行导致整个宿主工程编译不过（源码改了但 DLL 停在旧时间戳）。
' 路径净化改用 HostPathHelper（Hosts\Shared\HostPathHelper.vb，随本工程编译）。

' PowerPoint 常用常量（晚期绑定用，避免 Interop 版本差异）
' msoTextOrientationHorizontal = 1
' ppLayoutTitle = 1 ; ppLayoutText = 2 ; ppLayoutBlank = 12
' msoShapeRectangle = 1 ; msoTrue = -1

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.AutoDual)>
Public Class PptBridge
    Inherits HostBridgeBase

    Private _pptApp As Object

    Public Sub New(app As Object)
        MyBase.New(OfficeHostKind.PowerPoint, app)
        _pptApp = app
    End Sub

    ' ---------------- 演示文稿信息 ----------------
    ' ---------------- 路径（宿主导出文件默认落点）----------------
    ' 基类 HostBridgeBase.GetDesktopPath 返回 "operation not supported on this host"，
    ' 导致前端拿不到当前用户桌面（只能退化到"由文档路径反推 profile"的兜底）。
    ' 这里按**当前登录用户**动态返回裸路径（与 Excel 主桥返回格式一致）。
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
    Public Function GetPresentationInfo() As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim count = 0
            Try
                count = CInt(pres.Slides.Count)
            Catch
            End Try
            Return "{""success"":true,""name"":""" & Safe(CStr(pres.Name)) & """," &
                   """fullName"":""" & Safe(CStr(pres.FullName)) & """," &
                   """slides"":" & count & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetSlideCount() As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return "0"
            Return CStr(pres.Slides.Count)
        Catch
            Return "0"
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetCurrentSlideIndex() As String
        Try
            Dim idx = GetCurrentSlideIndexValue()
            Return CStr(idx)
        Catch
            Return "0"
        End Try
    End Function

    ' ---------------- 读取 ----------------
    <ComVisible(True)>
    Public Function GetSlideText(slideIndex As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return ""
            ' 【关键修复】slideIndex 缺省为 0（HostMcpTools 未传参时给的默认值就是 0），
            ' 语义是"当前幻灯片"（见 ppt_tools.json 的工具描述）。但旧实现直接 pres.Slides(0)
            ' 会抛异常 → 被下面 Catch 吞掉 → 永远返回空串：只要调用方不显式传 slideIndex，
            ' "读取/总结当前幻灯片"就必然读不到任何内容（模型据此编造正文的直接诱因）。
            ' 与 DeleteSlide / DuplicateSlide / SetSlideNotes / GetSlideNotes 保持一致：
            ' <=0 回落当前页；越界（>总页数）同样回落，避免抛异常被吞。
            Dim idx = slideIndex
            Dim slideCount = CInt(pres.Slides.Count)
            If idx <= 0 OrElse idx > slideCount Then idx = GetCurrentSlideIndexValue()
            If idx <= 0 OrElse idx > slideCount Then Return ""
            Dim slide = pres.Slides(idx)
            Dim sb As New StringBuilder()
            For Each shp As Object In slide.Shapes
                Try
                    If shp.HasTextFrame Then
                        If shp.TextFrame.HasText Then
                            sb.AppendLine(CStr(shp.TextFrame.TextRange.Text))
                        End If
                    End If
                Catch
                End Try
            Next
            Return sb.ToString()
        Catch ex As Exception
            Return ""
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetSlideTextJson(slideIndex As Integer) As String
        Return "{""success"":true,""text"":""" & Safe(GetSlideText(slideIndex)) & """}"
    End Function

    <ComVisible(True)>
    Public Function GetAllSlidesTextJson(maxChars As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return "{""success"":false,""error"":""没有打开的演示文稿""}"
            Dim sb As New StringBuilder()
            For i As Integer = 1 To CInt(pres.Slides.Count)
                sb.AppendLine("[幻灯片 " & i & "]")
                sb.AppendLine(GetSlideText(i))
            Next
            Dim text = sb.ToString()
            If maxChars > 0 AndAlso text.Length > maxChars Then text = text.Substring(0, maxChars)
            Return "{""success"":true,""text"":""" & Safe(text) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetSelectedText() As String
        Try
            If _pptApp Is Nothing Then Return ""
            Dim sel = _pptApp.ActiveWindow.Selection
            If sel Is Nothing Then Return ""
            Return CStr(sel.TextRange.Text)
        Catch
            Return ""
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetSelectedTextJson() As String
        Return "{""success"":true,""text"":""" & Safe(GetSelectedText()) & """}"
    End Function

    ' ---------------- 写入 ----------------
    <ComVisible(True)>
    Public Function AddSlide(layoutIndex As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then idx = CInt(pres.Slides.Count)
            Dim layout = If(layoutIndex <= 0, 12, layoutIndex)
            Dim slide = pres.Slides.Add(idx + 1, layout)
            Return "{""success"":true,""index"":" & idx + 1 & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function AddTitleAndContentSlide(title As String, content As String) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then idx = CInt(pres.Slides.Count)
            Dim slide = pres.Slides.Add(idx + 1, 2)   ' ppLayoutText
            Try
                slide.Shapes(1).TextFrame.TextRange.Text = title
            Catch
            End Try
            Try
                slide.Shapes(2).TextFrame.TextRange.Text = content
            Catch
            End Try
            Return "{""success"":true,""index"":" & idx + 1 & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function AddTextbox(text As String) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片（请切换到普通视图）")
            Dim pres = GetActivePresentation()
            Dim w = 600.0F, h = 60.0F
            Dim left = 60.0F, top = 120.0F
            Try
                w = CSng(pres.PageSetup.SlideWidth) - 120.0F
            Catch
            End Try
            Dim shape = slide.Shapes.AddTextbox(1, left, top, w, h)
            shape.TextFrame.TextRange.Text = text
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function SavePresentation(filePath As String) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            If String.IsNullOrWhiteSpace(filePath) Then
                Try
                    pres.Save()
                Catch
                    Return Err("演示文稿尚未保存过，请提供 filePath 参数以另存为")
                End Try
            Else
                pres.SaveAs(filePath)
            End If
            Return "{""success"":true,""fullName"":""" & Safe(CStr(pres.FullName)) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FindReplace(findText As String, replaceText As String, matchCase As Boolean) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            If String.IsNullOrEmpty(findText) Then Return Err("findText 不能为空")
            Dim replaced As Integer = 0
            For i As Integer = 1 To CInt(pres.Slides.Count)
                Dim slide = pres.Slides(i)
                For Each shp As Object In slide.Shapes
                    Try
                        If shp.HasTextFrame Then
                            If shp.TextFrame.HasText Then
                                Dim tr = shp.TextFrame.TextRange
                                Dim afterRange = tr.Replace(findText, replaceText, , matchCase, False)
                                If afterRange IsNot Nothing Then replaced += 1
                            End If
                        End If
                    Catch
                    End Try
                Next
            Next
            Return "{""success"":true,""replacedShapes"":" & replaced & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FormatText(bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object) As String
        Try
            Dim targets As New System.Collections.ArrayList()
            Try
                Dim sel = _pptApp.ActiveWindow.Selection
                If sel IsNot Nothing AndAlso CInt(sel.Type) = 2 Then
                    targets.Add(sel.TextRange)
                End If
            Catch
            End Try
            If targets.Count = 0 Then
                Dim slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片（请切换到普通视图并选中文本）")
                For Each shp As Object In slide.Shapes
                    If shp.HasTextFrame AndAlso shp.TextFrame.HasText Then targets.Add(shp.TextFrame.TextRange)
                Next
            End If
            If targets.Count = 0 Then Return Err("没有可格式化的文本")
            For Each tr As Object In targets
                ApplyTextFormat(tr.Font, bold, italic, fontSize, fontName, fontColor)
            Next
            Return "{""success"":true,""targets"":" & targets.Count & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Sub ApplyTextFormat(font As Object, bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object)
        If bold IsNot Nothing Then font.Bold = If(CBool(bold), -1, 0)
        If italic IsNot Nothing Then font.Italic = If(CBool(italic), -1, 0)
        If fontSize IsNot Nothing Then font.Size = CSng(fontSize)
        If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then font.Name = CStr(fontName)
        If fontColor IsNot Nothing Then font.Color.RGB = OleColorFromArgb(CStr(fontColor))
    End Sub

    ' OleColorFromArgb 已上提到 HostBridgeBase 共享。

    <ComVisible(True)>
    Public Function InsertPicture(filePath As String, left As Object, top As Object, width As Object, height As Object) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片（请切换到普通视图）")
            If String.IsNullOrEmpty(filePath) Then Return Err("filePath 不能为空")
            Dim l = If(left Is Nothing, 60.0F, CSng(left))
            Dim t = If(top Is Nothing, 120.0F, CSng(top))
            Dim w = If(width Is Nothing, 300.0F, CSng(width))
            Dim h = If(height Is Nothing, 200.0F, CSng(height))
            slide.Shapes.AddPicture(filePath, False, True, l, t, w, h)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function DeleteSlide(slideIndex As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim idx = slideIndex
            If idx <= 0 Then idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then Return Err("无法确定要删除的幻灯片")
            pres.Slides(idx).Delete()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function DuplicateSlide(slideIndex As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim idx = slideIndex
            If idx <= 0 Then idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then Return Err("无法确定要复制的幻灯片")
            Dim newSlide = pres.Slides(idx).Duplicate()
            Return "{""success"":true,""index"":" & CInt(newSlide.SlideIndex) & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function DeleteSelectedShape() As String
        Try
            Dim sel = _pptApp.ActiveWindow.Selection
            If sel Is Nothing Then Return Err("没有选中形状")
            If CInt(sel.Type) = 1 Then
                sel.ShapeRange.Delete()
                Return "{""success"":true}"
            End If
            Return Err("当前选中的不是形状")
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ExportPdf(filePath As String) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")

            If String.IsNullOrWhiteSpace(filePath) Then
                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                If String.IsNullOrWhiteSpace(desktop) Then desktop = System.IO.Path.GetTempPath()
                Dim baseName As String = "presentation"
                Try
                    Dim n As String = CStr(pres.Name)
                    If Not String.IsNullOrWhiteSpace(n) Then baseName = System.IO.Path.GetFileNameWithoutExtension(n)
                Catch
                End Try
                baseName = HostPathHelper.GetSafeFileName(baseName)
                If String.IsNullOrWhiteSpace(baseName) Then baseName = "presentation"
                Dim ts As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                filePath = System.IO.Path.Combine(desktop, baseName & "_" & ts & ".pdf")
            End If

            Dim rawPath As String = filePath
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

            ' 【实测】ExportAsFixedFormat 的 IDispatch 参数在 PowerPoint 类型库里声明为
            ' ByRef VARIANT，.NET 后期绑定（Object 变量直调）必报 DISP_E_TYPEMISMATCH
            ' （VBA 能用是因为 VBA 全按 ByRef Variant 传参）。改用 SaveAs(ppSaveAsPDF=32)，
            ' 与 WordBridge.ExportPdf 的 doc.SaveAs(path, wdFormatPDF=17) 同构。
            pres.SaveAs(filePath, CObj(32))
            ' message 字段与 excel_export_csv 同构：tool-card 通用成功摘要会优先展示 data.message
            Return "{""success"":true,""path"":""" & Safe(filePath) & """,""message"":""成功导出到 PDF: " & Safe(filePath) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function FormatTableCell(shapeIndex As Integer, row As Integer, col As Integer, bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object, align As Object, fillColor As Object, borderColor As Object, highlightColor As Object) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            If row <= 0 OrElse col <= 0 Then Return Err("row/col 必须从 1 开始")
            Dim tableShape = FindTableShape(slide, shapeIndex)
            If tableShape Is Nothing Then Return Err("未找到表格形状")
            Dim table = tableShape.Table
            If row > CInt(table.Rows.Count) OrElse col > CInt(table.Columns.Count) Then Return Err("行列超出范围")
            Dim cell = table.Cell(row, col)
            Dim tr = cell.Shape.TextFrame.TextRange
            If bold IsNot Nothing Then tr.Font.Bold = If(CBool(bold), -1, 0)
            If italic IsNot Nothing Then tr.Font.Italic = If(CBool(italic), -1, 0)
            If fontSize IsNot Nothing Then tr.Font.Size = CSng(fontSize)
            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then tr.Font.Name = CStr(fontName)
            If fontColor IsNot Nothing Then tr.Font.Color.RGB = OleColorFromArgb(CStr(fontColor))
            If highlightColor IsNot Nothing Then
                Try : cell.Shape.TextFrame2.TextRange.Font.Highlight.RGB = OleColorFromArgb(CStr(highlightColor)) : Catch : End Try
            End If
            If align IsNot Nothing Then tr.ParagraphFormat.Alignment = MapAlign(CStr(align))
            If fillColor IsNot Nothing Then
                cell.Shape.Fill.Visible = -1
                cell.Shape.Fill.ForeColor.RGB = OleColorFromArgb(CStr(fillColor))
            End If
            If borderColor IsNot Nothing Then
                Dim bc = OleColorFromArgb(CStr(borderColor))
                For Each b As Object In cell.Borders
                    b.ForeColor.RGB = bc
                    b.Visible = -1
                Next
            End If
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' ==================================================================
    ' 选中对象信息采集：让 AI 知道用户选中的是幻灯片上的表格/形状，而非"无选区"
    ' ==================================================================
    <ComVisible(True)>
    Public Overrides Function GetSelectionInfo() As String
        Try
            If _pptApp Is Nothing Then
                Return JsonConvert.SerializeObject(New With {.success = False, .selectionType = "none", .summary = "PowerPoint 未运行"})
            End If
            Dim aw As Object = Nothing
            Try : aw = _pptApp.ActiveWindow : Catch : aw = Nothing : End Try
            If aw Is Nothing Then
                Return JsonConvert.SerializeObject(New With {.success = True, .selectionType = "none", .slideIndex = GetCurrentSlideIndexValue(), .summary = "没有活动窗口"})
            End If

            Dim sel As Object = Nothing
            Try : sel = aw.Selection : Catch : sel = Nothing : End Try
            Dim slideIndex = GetCurrentSlideIndexValue()

            If sel Is Nothing Then
                Return JsonConvert.SerializeObject(New With {.success = True, .selectionType = "none", .slideIndex = slideIndex, .summary = "未选中任何对象"})
            End If

            Dim selType As Integer = 0
            Try : selType = CInt(sel.Type) : Catch : selType = 0 : End Try

            Dim selectedText As String = ""
            Try
                If sel.TextRange IsNot Nothing Then selectedText = CStr(sel.TextRange.Text)
            Catch
            End Try

            ' 【修复·关键】PpSelectionType 枚举实际取值：ppSelectionNone=0、ppSelectionSlides=1、
            ' ppSelectionShapes=2、ppSelectionText=3。原代码写 `If selType = 2 Then` 并注释为
            ' "选中整张幻灯片（缩略图窗格）"——枚举值错位：2 其实是**选中了形状**，1 才是选中整张幻灯片。
            ' 后果：用户在幻灯片上【选中表格（形状）】时被提前返回 selectionType="slides"，
            ' 前端永远认不出"已选中表格"，选中表格格式化分支彻底失效（实测宿主日志：
            ' selectedObject={type=slides,rows=,cols=,summary=选中了第 12 张幻灯片}）。
            ' 改为不依赖枚举：先尝试取 ShapeRange，取不到才判定为"整张幻灯片/无具体对象"。
            Dim shp As Object = Nothing
            Try : shp = sel.ShapeRange(1) : Catch : shp = Nothing : End Try
            If shp Is Nothing Then
                If selType = 1 Then
                    ' 选中整张幻灯片（缩略图窗格）
                    Return JsonConvert.SerializeObject(New With {
                        .success = True, .selectionType = "slides", .slideIndex = slideIndex,
                        .summary = "选中了第 " & slideIndex & " 张幻灯片"})
                End If
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .selectionType = "text", .slideIndex = slideIndex, .selectedText = selectedText,
                    .summary = If(selectedText <> "", "选中文本：" & selectedText, "未选中具体对象")})
            End If

            Dim shapeName As String = ""
            Dim shapeType As Integer = 0
            Dim hasTable As Boolean = False
            Try : shapeName = CStr(shp.Name) : Catch : End Try
            Try : shapeType = CInt(shp.Type) : Catch : End Try
            Try : hasTable = CBool(shp.HasTable) : Catch : End Try

            If hasTable Then
                Dim tbl As Object = shp.Table
                Dim r = CInt(tbl.Rows.Count)
                Dim c = CInt(tbl.Columns.Count)
                Dim shapeIndex = TableShapeIndexOnSlide(shapeName)
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .selectionType = "table", .slideIndex = slideIndex, .shapeIndex = shapeIndex,
                    .shapeName = shapeName, .shapeType = shapeType, .hasTable = True,
                    .rowCount = r, .colCount = c, .selectedText = selectedText,
                    .summary = "在第 " & slideIndex & " 张幻灯片选中了一个表格（形状序号 " & shapeIndex & "，共 " & r & " 行 " & c & " 列）"})
            End If

            Return JsonConvert.SerializeObject(New With {
                .success = True, .selectionType = "shape", .slideIndex = slideIndex,
                .shapeName = shapeName, .shapeType = shapeType, .hasTable = False, .selectedText = selectedText,
                .summary = "在第 " & slideIndex & " 张幻灯片选中了一个形状（" & shapeName & "）"})
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    ' 在"当前幻灯片"的 HasTable 形状集合中，返回与 shapeName 同名的表格形状序号
    ' （与 FindTableShape 的计数口径一致，保证 AI 拿到的 shapeIndex 能直接传给 FormatTableCell）
    Private Function TableShapeIndexOnSlide(shapeName As String) As Integer
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return -1
            Dim idx As Integer = 0
            For Each shp As Object In slide.Shapes
                If CBool(shp.HasTable) Then
                    idx += 1
                    If CStr(shp.Name) = shapeName Then Return idx
                End If
            Next
        Catch
        End Try
        Return -1
    End Function

    ''' <summary>给单个单元格画一条边（后期绑定，失败静默）。idx: 1=Top 2=Left 3=Bottom 4=Right</summary>
    Private Sub PaintCellBorder(table As Object, r As Integer, c As Integer, idx As Integer, bc As Object)
        Try
            Dim b As Object = table.Cell(r, c).Borders(idx)
            b.Visible = -1                       ' msoTrue
            b.ForeColor.RGB = bc
            Try : b.LineStyle = 1 : Catch : End Try   ' msoLineSingle（个别版本不支持则忽略）
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' 按版式给表格画边框（PowerPoint 后期绑定；PPT 没有"整表边框"对象，只能逐格画四条边）。
    ''' style: all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
    ''' PpBorderType: Top=1 Left=2 Bottom=3 Right=4（对角线 5/6 不涉及）。
    ''' </summary>
    Private Sub ApplyTableBorderStyle(table As Object, bc As Object, style As Object, rCount As Integer, cCount As Integer)
        Dim st As String = If(style Is Nothing, "all", LCase(CStr(style).Trim()))
        If st = "" Then st = "all"
        Dim r As Integer, c As Integer
        Try
            ' 先全部擦掉（msoFalse=0），避免旧边框残留
            For r = 1 To rCount
                For c = 1 To cCount
                    Try
                        Dim cl As Object = table.Cell(r, c)
                        For i As Integer = 1 To 4
                            cl.Borders(i).Visible = 0
                        Next
                    Catch
                    End Try
                Next
            Next
            If st = "none" Then Exit Sub
            If st = "all" Then
                For r = 1 To rCount
                    For c = 1 To cCount
                        For i As Integer = 1 To 4
                            PaintCellBorder(table, r, c, i, bc)
                        Next
                    Next
                Next
            ElseIf st = "horizontal" Then
                For r = 1 To rCount
                    For c = 1 To cCount
                        If r = 1 Then PaintCellBorder(table, r, c, 1, bc)   ' 首行顶线
                        PaintCellBorder(table, r, c, 3, bc)                 ' 每行底线 → 形成横向分隔
                    Next
                Next
            ElseIf st = "outer" Then
                For r = 1 To rCount
                    For c = 1 To cCount
                        If r = 1 Then PaintCellBorder(table, r, c, 1, bc)
                        If r = rCount Then PaintCellBorder(table, r, c, 3, bc)
                        If c = 1 Then PaintCellBorder(table, r, c, 2, bc)
                        If c = cCount Then PaintCellBorder(table, r, c, 4, bc)
                    Next
                Next
            ElseIf st = "threeline" Then
                If rCount >= 1 Then
                    For c = 1 To cCount
                        PaintCellBorder(table, 1, c, 1, bc)     ' 顶线
                        PaintCellBorder(table, 1, c, 3, bc)     ' 表头下线
                    Next
                End If
                For c = 1 To cCount
                    PaintCellBorder(table, rCount, c, 3, bc)    ' 底线
                Next
            Else
                ' 未知版式 → 退回全网格
                For r = 1 To rCount
                    For c = 1 To cCount
                        For i As Integer = 1 To 4
                            PaintCellBorder(table, r, c, i, bc)
                        Next
                    Next
                Next
            End If
        Catch
        End Try
    End Sub

    Private Function GetSelectedTableShape() As Object
        Try
            If _pptApp Is Nothing OrElse _pptApp.ActiveWindow Is Nothing Then Return Nothing
            Dim aw = _pptApp.ActiveWindow
            Dim sel = aw.Selection
            If sel Is Nothing Then Return Nothing

            ' 1) 形状级选择（点表格边框 / 选择窗格选中整个表格）
            Dim shp As Object = Nothing
            Try : shp = sel.ShapeRange(1) : Catch : shp = Nothing : End Try
            If shp IsNot Nothing AndAlso SafeHasTable(shp) Then Return shp

            ' 2) 文本级选择（光标停在单元格内编辑）：Selection.TextRange 属于某个单元格，
            '    回溯 Parent 链取到 Table，再取回其 Shape。
            Try
                Dim tr = sel.TextRange
                If tr IsNot Nothing Then
                    Dim p1 = tr.Parent          ' 通常是 Cell
                    Dim t1 = TableOf(p1)
                    If t1 IsNot Nothing Then Return ShapeOfTable(t1)
                    Dim p2 = SafeParent(p1)     ' Cell.Parent = Table
                    Dim t2 = TableOf(p2)
                    If t2 IsNot Nothing Then Return ShapeOfTable(t2)
                End If
            Catch
            End Try

            ' 3) 再次尝试 ShapeRange（文本选择下有时仍含表格形状）
            Try
                If sel.ShapeRange IsNot Nothing AndAlso sel.ShapeRange.Count >= 1 Then
                    Dim s2 = sel.ShapeRange(1)
                    If SafeHasTable(s2) Then Return s2
                End If
            Catch
            End Try
        Catch
        End Try
        Return Nothing
    End Function

    Private Function SafeHasTable(shp As Object) As Boolean
        Try : Return CBool(shp.HasTable) : Catch : Return False : End Try
    End Function

    Private Function TableOf(o As Object) As Object
        Try : Return o.Table : Catch : Return Nothing : End Try
    End Function

    Private Function SafeParent(o As Object) As Object
        Try : Return o.Parent : Catch : Return Nothing : End Try
    End Function

    Private Function ShapeOfTable(tbl As Object) As Object
        Try : Return tbl.Parent : Catch : Return Nothing : End Try
    End Function

    ''' <summary>
    ''' 格式化「当前选中的表格」的某个单元格（无需 AI 再传 shapeIndex/幻灯片号）。
    ''' 若当前未选中表格，返回明确错误提示，而不是静默失败。
    '''
    ''' 2026-10-07 新增排版维度（整表模式下生效）：borderStyle / headerFontColor / headerItalic /
    ''' zebraColor / firstColBold / lastRowBold / bodyAlign / rowHeight / valign（含义同 WordBridge）。
    ''' </summary>
    <ComVisible(True)>
    Public Function FormatSelectedTableCell(row As Integer, col As Integer, bold As Object, italic As Object, fontSize As Object, fontName As Object, fontColor As Object, align As Object, fillColor As Object, borderColor As Object, headerBold As Object, headerFillColor As Object, clearFill As Object, highlightColor As Object, zebra As Object, Optional borderStyle As Object = Nothing, Optional headerFontColor As Object = Nothing, Optional zebraColor As Object = Nothing, Optional firstColBold As Object = Nothing, Optional lastRowBold As Object = Nothing, Optional bodyAlign As Object = Nothing, Optional rowHeight As Object = Nothing, Optional valign As Object = Nothing, Optional headerItalic As Object = Nothing) As String
        Try
            LogBridge("FormatSelectedTableCell 入口: align=" & If(align Is Nothing, "∅", CStr(align)) & " headerFillColor=" & If(headerFillColor Is Nothing, "∅", CStr(headerFillColor)) & " clearFill=" & If(clearFill Is Nothing, "∅", CStr(clearFill)) & " bold=" & If(bold Is Nothing, "∅", CStr(bold)) & " fillColor=" & If(fillColor Is Nothing, "∅", CStr(fillColor)))
            Dim shp = GetSelectedTableShape()
            If shp Is Nothing Then
                LogBridge("FormatSelectedTableCell 未选中表格（GetSelectedTableShape 返回 Nothing）")
                Return Err("当前没有选中表格，请先点选要设置的表格")
            End If
            Dim table = shp.Table
            Dim rCount As Integer = CInt(table.Rows.Count)
            Dim cCount As Integer = CInt(table.Columns.Count)

            ' row/col 任一省略(<=0) → 整表生效：用户只说"设置对齐/颜色"时无需再报幻灯片或行列。
            ' 语义：align/fontSize/fontName/fontColor/borderColor 作用于所有单元格；
            '       bold/fillColor（或显式 headerBold/headerFillColor）作用于表头行（第 1 行）。
            If row <= 0 OrElse col <= 0 Then
                Dim bcWhole As Object = Nothing
                If borderColor IsNot Nothing Then bcWhole = OleColorFromArgb(CStr(borderColor))
                Dim hbWhole As Object = If(headerBold IsNot Nothing, headerBold, bold)
                Dim hfWhole As Object = If(headerFillColor IsNot Nothing, headerFillColor, fillColor)
                Dim _hasClear As Boolean = (clearFill IsNot Nothing AndAlso CBool(clearFill))
                If zebraColor IsNot Nothing AndAlso zebra Is Nothing Then zebra = True
                Dim _hasLayoutItem As Boolean = (borderStyle IsNot Nothing) OrElse (headerFontColor IsNot Nothing) OrElse
                    (firstColBold IsNot Nothing) OrElse (lastRowBold IsNot Nothing) OrElse (bodyAlign IsNot Nothing) OrElse
                    (rowHeight IsNot Nothing) OrElse (valign IsNot Nothing) OrElse (headerItalic IsNot Nothing) OrElse
                    (zebra IsNot Nothing AndAlso CBool(zebra))
                ' 防御：整表模式却一个格式项都没给 → 明确报错，避免"什么都没改却报成功"的静默假完成（clearFill 也算有效项）
                If hbWhole Is Nothing AndAlso hfWhole Is Nothing AndAlso align Is Nothing AndAlso fontSize Is Nothing AndAlso fontName Is Nothing AndAlso fontColor Is Nothing AndAlso bcWhole Is Nothing AndAlso highlightColor Is Nothing AndAlso Not _hasClear AndAlso Not _hasLayoutItem Then
                    Return Err("未指定任何格式项（对齐/颜色/底纹/字体/边框/高亮/版式等），请说明要设置什么")
                End If
                ' 【表头默认区分】走到这里说明本次至少有一项格式要求；若用户既没指定表头样式、也没要求清除底纹，
                ' 默认给表头行加粗 + 浅灰填充，避免"整张表一个样、看不出哪行是表头"（实测反馈："默认的没有表头区分吗"）。
                ' 用户显式给的 headerBold/headerFillColor/bold/fillColor 及 clearFill 优先，均不受影响。
                ' 仅整表模式生效；指定 row/col 的单格调用不走此分支。
                If hfWhole Is Nothing AndAlso Not _hasClear Then hfWhole = "#D9D9D9"
                If hbWhole Is Nothing Then hbWhole = True
                ' 清除底色：Fill.Visible=-1(msoTrue) 表示原本有可见填充，仅这类格计入"实际变化"
                Dim clearedFillCells As Integer = 0
                If _hasClear Then
                    Dim cr As Integer, cc As Integer
                    For cr = 1 To rCount
                        For cc = 1 To cCount
                            Try
                                If Convert.ToInt32(table.Cell(cr, cc).Shape.Fill.Visible) = -1 Then clearedFillCells += 1
                                table.Cell(cr, cc).Shape.Fill.Visible = 0
                            Catch
                            End Try
                        Next
                    Next
                End If
                Dim hasAlign As Boolean = (align IsNot Nothing)
                Dim targetAlign As Integer = 0
                If hasAlign Then targetAlign = MapAlign(CStr(align))
                Dim fontCells As Integer = 0
                Dim highlightCells As Integer = 0
                Dim alignCells As Integer = 0
                Dim borderCells As Integer = 0
                Dim r As Integer
                Dim c As Integer
                For r = 1 To rCount
                    For c = 1 To cCount
                        Try
                            Dim cellObj = table.Cell(r, c)
                            Dim tw = cellObj.Shape.TextFrame.TextRange
                            If fontSize IsNot Nothing Then
                                tw.Font.Size = CSng(fontSize) : fontCells += 1
                            End If
                            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then
                                tw.Font.Name = CStr(fontName) : fontCells += 1
                            End If
                            If fontColor IsNot Nothing Then
                                tw.Font.Color.RGB = OleColorFromArgb(CStr(fontColor)) : fontCells += 1
                            End If
                            If highlightColor IsNot Nothing Then
                                Dim _hc = OleColorFromArgb(CStr(highlightColor))
                                Try : cellObj.Shape.TextFrame2.TextRange.Font.Highlight.RGB = _hc : highlightCells += 1 : Catch : End Try
                            End If
                            If hasAlign Then
                                tw.ParagraphFormat.Alignment = targetAlign : alignCells += 1
                            End If
                            If bcWhole IsNot Nothing Then borderCells += 1
                        Catch
                        End Try
                    Next
                Next
                ' 【边框版式】PPT 无"整表边框"对象，逐格按版式画/擦四条边
                If bcWhole IsNot Nothing Then
                    ApplyTableBorderStyle(table, bcWhole, borderStyle, rCount, cCount)
                End If
                Dim headerBoldCells As Integer = 0
                If hbWhole IsNot Nothing AndAlso CBool(hbWhole) Then
                    For c = 1 To cCount
                        Try
                            table.Cell(1, c).Shape.TextFrame.TextRange.Font.Bold = -1 : headerBoldCells += 1
                        Catch
                        End Try
                    Next
                End If
                ' 表头字色（深底白字 / 浅底深字）
                Dim headerFontCells As Integer = 0
                If headerFontColor IsNot Nothing AndAlso CStr(headerFontColor).Trim() <> "" Then
                    Dim hfc = OleColorFromArgb(CStr(headerFontColor))
                    For c = 1 To cCount
                        Try
                            table.Cell(1, c).Shape.TextFrame.TextRange.Font.Color.RGB = hfc : headerFontCells += 1
                        Catch
                        End Try
                    Next
                End If
                ' 斜体表头
                Dim headerItalicCells As Integer = 0
                If headerItalic IsNot Nothing AndAlso CBool(headerItalic) Then
                    For c = 1 To cCount
                        Try
                            table.Cell(1, c).Shape.TextFrame.TextRange.Font.Italic = -1 : headerItalicCells += 1
                        Catch
                        End Try
                    Next
                End If
                Dim headerFillCells As Integer = 0
                If hfWhole IsNot Nothing Then
                    If CStr(hfWhole).Trim() = "" Then
                        For c = 1 To cCount
                            Try
                                table.Cell(1, c).Shape.Fill.Visible = 0 : headerFillCells += 1   ' 明确不要表头底色
                            Catch
                            End Try
                        Next
                    Else
                        Dim hcWhole = OleColorFromArgb(CStr(hfWhole))
                        For c = 1 To cCount
                            Try
                                table.Cell(1, c).Shape.Fill.Visible = -1
                                table.Cell(1, c).Shape.Fill.ForeColor.RGB = hcWhole
                                headerFillCells += 1
                            Catch
                            End Try
                        Next
                    End If
                End If
                ' 隔行底纹（zebra）：正文偶数行（第2、4、6…行）着浅色；与 clearFill 互斥。
                ' 浅色由表头底色派生（MakeZebraTint 向白色混合 80%），主题色表得到同色系浅底纹。
                Dim zebraCells As Integer = 0
                If zebra IsNot Nothing AndAlso CBool(zebra) AndAlso Not _hasClear Then
                    Dim _zHex As String = If(zebraColor IsNot Nothing AndAlso CStr(zebraColor).Trim() <> "",
                                             CStr(zebraColor),
                                             MakeZebraTint(If(hfWhole IsNot Nothing AndAlso CStr(hfWhole).Trim() <> "", CStr(hfWhole), "#D9D9D9")))
                    Dim zc = OleColorFromArgb(_zHex)
                    For r = 2 To rCount
                        If (r Mod 2) = 0 Then
                            For c = 1 To cCount
                                Try
                                    table.Cell(r, c).Shape.Fill.Visible = -1
                                    table.Cell(r, c).Shape.Fill.ForeColor.RGB = zc
                                    zebraCells += 1
                                Catch
                                End Try
                            Next
                        End If
                    Next
                End If
                ' ── 行高（紧凑 16 / 标准 22 / 宽松 30 磅）──
                Dim rowHeightCells As Integer = 0
                If rowHeight IsNot Nothing Then
                    Dim _rh As Double = 0
                    Try : _rh = CDbl(rowHeight) : Catch : _rh = 0 : End Try
                    If _rh >= 10 AndAlso _rh <= 120 Then
                        For r = 1 To rCount
                            Try
                                table.Rows(r).Height = CSng(_rh) : rowHeightCells += 1
                            Catch
                            End Try
                        Next
                    End If
                End If
                ' ── 垂直对齐：middle 垂直居中(默认) / top 顶端 ──
                Dim valignCells As Integer = 0
                If valign IsNot Nothing AndAlso CStr(valign).Trim() <> "" Then
                    Dim _va As Integer = If(LCase(CStr(valign).Trim()) = "top", 1, 3)   ' msoAnchorTop=1 / msoAnchorMiddle=3
                    For r = 1 To rCount
                        For c = 1 To cCount
                            Try
                                table.Cell(r, c).Shape.TextFrame.VerticalAnchor = _va : valignCells += 1
                            Catch
                            End Try
                        Next
                    Next
                End If
                ' ── 正文对齐（覆盖整表 align）──
                Dim bodyAlignCells As Integer = 0
                If bodyAlign IsNot Nothing AndAlso CStr(bodyAlign).Trim() <> "" Then
                    Dim _ba As Integer = MapAlign(CStr(bodyAlign))
                    For r = 2 To rCount
                        For c = 1 To cCount
                            Try
                                table.Cell(r, c).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = _ba : bodyAlignCells += 1
                            Catch
                            End Try
                        Next
                    Next
                End If
                ' ── 首列强调（加粗 + 极浅底）──
                Dim firstColCells As Integer = 0
                If firstColBold IsNot Nothing AndAlso CBool(firstColBold) Then
                    Dim _fcTint As String = MakeZebraTint(If(hfWhole IsNot Nothing AndAlso CStr(hfWhole).Trim() <> "", CStr(hfWhole), "#D9D9D9"), 0.90)
                    Dim _fcColor = OleColorFromArgb(_fcTint)
                    For r = 1 To rCount
                        Try
                            table.Cell(r, 1).Shape.TextFrame.TextRange.Font.Bold = -1
                            table.Cell(r, 1).Shape.Fill.Visible = -1
                            table.Cell(r, 1).Shape.Fill.ForeColor.RGB = _fcColor
                            firstColCells += 1
                        Catch
                        End Try
                    Next
                End If
                ' ── 末行（合计行）强调：加粗 + 上边线 ──
                Dim lastRowCells As Integer = 0
                If lastRowBold IsNot Nothing AndAlso CBool(lastRowBold) Then
                    Dim _lrColor As Object = If(bcWhole IsNot Nothing, bcWhole,
                                                OleColorFromArgb(If(hfWhole IsNot Nothing AndAlso CStr(hfWhole).Trim() <> "", CStr(hfWhole), "#000000")))
                    For c = 1 To cCount
                        Try
                            table.Cell(rCount, c).Shape.TextFrame.TextRange.Font.Bold = -1
                            Dim _tb As Object = table.Cell(rCount, c).Borders(1)   ' 1 = ppBorderTop
                            _tb.Visible = -1
                            _tb.ForeColor.RGB = _lrColor
                            lastRowCells += 1
                        Catch
                        End Try
                    Next
                End If
                ' 组装"具体参数变化"清单：每项含参数名/目标值/影响格数
                Dim changes As New List(Of Object)
                If hasAlign Then changes.Add(New With {.param = "align", .value = CStr(align), .cells = alignCells})
                If fontCells > 0 Then changes.Add(New With {.param = "font", .value = ("字号=" & If(fontSize Is Nothing, "-", CStr(fontSize)) & " 字体=" & If(fontName Is Nothing, "-", CStr(fontName)) & " 颜色=" & If(fontColor Is Nothing, "-", CStr(fontColor))), .cells = fontCells})
                If highlightCells > 0 Then changes.Add(New With {.param = "highlightColor", .value = CStr(highlightColor), .cells = highlightCells})
                If headerBoldCells > 0 Then changes.Add(New With {.param = "headerBold", .value = "True", .cells = headerBoldCells})
                If headerFillCells > 0 Then changes.Add(New With {.param = "headerFillColor", .value = CStr(hfWhole), .cells = headerFillCells})
                If headerFontCells > 0 Then changes.Add(New With {.param = "headerFontColor", .value = CStr(headerFontColor), .cells = headerFontCells})
                If headerItalicCells > 0 Then changes.Add(New With {.param = "headerItalic", .value = "True", .cells = headerItalicCells})
                If zebraCells > 0 Then changes.Add(New With {.param = "zebra", .value = "True", .cells = zebraCells})
                If rowHeightCells > 0 Then changes.Add(New With {.param = "rowHeight", .value = CStr(rowHeight), .cells = rowHeightCells})
                If valignCells > 0 Then changes.Add(New With {.param = "valign", .value = CStr(valign), .cells = valignCells})
                If bodyAlignCells > 0 Then changes.Add(New With {.param = "bodyAlign", .value = CStr(bodyAlign), .cells = bodyAlignCells})
                If firstColCells > 0 Then changes.Add(New With {.param = "firstColBold", .value = "True", .cells = firstColCells})
                If lastRowCells > 0 Then changes.Add(New With {.param = "lastRowBold", .value = "True", .cells = lastRowCells})
                If _hasClear Then changes.Add(New With {.param = "clearFill", .value = "True", .cells = clearedFillCells})
                If borderCells > 0 Then changes.Add(New With {.param = "borderStyle", .value = (If(borderStyle Is Nothing, "all", CStr(borderStyle)) & " " & CStr(borderColor)), .cells = borderCells})
                Dim parts As New List(Of String)
                If alignCells > 0 Then parts.Add("对齐→" & CStr(align) & "(" & alignCells & "格)")
                If fontCells > 0 Then parts.Add("字体格式(" & fontCells & "格)")
                If highlightCells > 0 Then parts.Add("文本高亮→" & CStr(highlightColor) & "(" & highlightCells & "格)")
                If headerBoldCells > 0 Then parts.Add("表头加粗(" & headerBoldCells & "格)")
                If headerFillCells > 0 Then parts.Add("表头底色→" & CStr(hfWhole) & "(" & headerFillCells & "格)")
                If headerFontCells > 0 Then parts.Add("表头字色→" & CStr(headerFontColor) & "(" & headerFontCells & "格)")
                If headerItalicCells > 0 Then parts.Add("表头斜体(" & headerItalicCells & "格)")
                If zebraCells > 0 Then parts.Add("隔行底纹(" & zebraCells & "格)")
                If rowHeightCells > 0 Then parts.Add("行高→" & CStr(rowHeight) & "磅(" & rowHeightCells & "行)")
                If valignCells > 0 Then parts.Add("垂直对齐→" & CStr(valign) & "(" & valignCells & "格)")
                If bodyAlignCells > 0 Then parts.Add("正文对齐→" & CStr(bodyAlign) & "(" & bodyAlignCells & "格)")
                If firstColCells > 0 Then parts.Add("首列强调(" & firstColCells & "行)")
                If lastRowCells > 0 Then parts.Add("末行强调(" & lastRowCells & "格)")
                If _hasClear Then parts.Add("清除底纹(" & clearedFillCells & "格)")
                If borderCells > 0 Then parts.Add("边框" & (If(borderStyle Is Nothing, "全网格", CStr(borderStyle))) & "→" & CStr(borderColor) & "(" & borderCells & "格)")
                Dim summary As String
                If parts.Count = 0 Then
                    summary = "表格（" & rCount & "行×" & cCount & "列）：未发现需变化的单元格（原本即符合目标格式）"
                Else
                    summary = "表格（" & rCount & "行×" & cCount & "列）整表已修改：" & String.Join("、", parts.ToArray())
                End If
                LogBridge("FormatSelectedTableCell 成功(整表) r=" & rCount & " c=" & cCount & " clearFill=" & If(clearFill Is Nothing, "∅", CStr(clearFill)) & " changes=" & String.Join(";", parts.ToArray()))
                Return JsonConvert.SerializeObject(New With {
                    .success = True, .scope = "wholeTable",
                    .rowCount = rCount, .colCount = cCount, .changed = (parts.Count > 0),
                    .applied = New With {
                        .align = If(align Is Nothing, "", CStr(align)),
                        .fontSize = If(fontSize Is Nothing, "", CStr(fontSize)),
                        .fontName = If(fontName Is Nothing, "", CStr(fontName)),
                        .fontColor = If(fontColor Is Nothing, "", CStr(fontColor)),
                        .highlightColor = If(highlightColor Is Nothing, "", CStr(highlightColor)),
                        .headerBold = If(hbWhole Is Nothing, "", CStr(CBool(hbWhole))),
                        .headerFillColor = If(hfWhole Is Nothing, "", CStr(hfWhole)),
                        .headerFontColor = If(headerFontColor Is Nothing, "", CStr(headerFontColor)),
                        .borderColor = If(borderColor Is Nothing, "", CStr(borderColor)),
                        .borderStyle = If(borderStyle Is Nothing, "all", CStr(borderStyle)),
                        .rowHeight = If(rowHeight Is Nothing, "", CStr(rowHeight)),
                        .valign = If(valign Is Nothing, "", CStr(valign)),
                        .bodyAlign = If(bodyAlign Is Nothing, "", CStr(bodyAlign)),
                        .firstColBold = If(firstColBold Is Nothing, "", CStr(CBool(firstColBold))),
                        .lastRowBold = If(lastRowBold Is Nothing, "", CStr(CBool(lastRowBold))),
                        .clearFill = _hasClear,
                        .zebra = If(zebra Is Nothing, "", CStr(CBool(zebra)))},
                    .changes = changes,
                    .summary = summary})
            End If

            If row > rCount OrElse col > cCount Then Return Err("行列超出范围")
            Dim cell = table.Cell(row, col)
            Dim tr = cell.Shape.TextFrame.TextRange
            ' 防御：只给了 row/col 却一个格式项都没给 → 明确报错，避免"什么都没改却报成功"的静默假完成
            If bold Is Nothing AndAlso italic Is Nothing AndAlso fontSize Is Nothing AndAlso fontName Is Nothing AndAlso fontColor Is Nothing AndAlso align Is Nothing AndAlso fillColor Is Nothing AndAlso borderColor Is Nothing AndAlso highlightColor Is Nothing AndAlso Not (clearFill IsNot Nothing AndAlso CBool(clearFill)) Then
                Return Err("未指定任何格式项（对齐/字号/字体/颜色/底纹/边框/高亮等），请说明要设置什么格式；若要对整张表格设置，请省略行号/列号")
            End If
            Dim chg As New List(Of Object)
            If bold IsNot Nothing Then
                tr.Font.Bold = If(CBool(bold), -1, 0)
                chg.Add(New With {.param = "bold", .value = CStr(CBool(bold)), .cells = 1})
            End If
            If italic IsNot Nothing Then
                tr.Font.Italic = If(CBool(italic), -1, 0)
                chg.Add(New With {.param = "italic", .value = CStr(CBool(italic)), .cells = 1})
            End If
            If fontSize IsNot Nothing Then
                tr.Font.Size = CSng(fontSize)
                chg.Add(New With {.param = "fontSize", .value = CStr(fontSize), .cells = 1})
            End If
            If fontName IsNot Nothing AndAlso CStr(fontName) <> "" Then
                tr.Font.Name = CStr(fontName)
                chg.Add(New With {.param = "fontName", .value = CStr(fontName), .cells = 1})
            End If
            If fontColor IsNot Nothing Then
                tr.Font.Color.RGB = OleColorFromArgb(CStr(fontColor))
                chg.Add(New With {.param = "fontColor", .value = CStr(fontColor), .cells = 1})
            End If
            If highlightColor IsNot Nothing Then
                Dim _hc = OleColorFromArgb(CStr(highlightColor))
                Try : cell.Shape.TextFrame2.TextRange.Font.Highlight.RGB = _hc : chg.Add(New With {.param = "highlightColor", .value = CStr(highlightColor), .cells = 1}) : Catch : End Try
            End If
            If align IsNot Nothing Then
                tr.ParagraphFormat.Alignment = MapAlign(CStr(align))
                chg.Add(New With {.param = "align", .value = CStr(align), .cells = 1})
            End If
            If fillColor IsNot Nothing Then
                cell.Shape.Fill.Visible = -1
                cell.Shape.Fill.ForeColor.RGB = OleColorFromArgb(CStr(fillColor))
                chg.Add(New With {.param = "fillColor", .value = CStr(fillColor), .cells = 1})
            End If
            If clearFill IsNot Nothing AndAlso CBool(clearFill) Then
                Dim had As Boolean = False
                Try : had = (Convert.ToInt32(cell.Shape.Fill.Visible) = -1) : cell.Shape.Fill.Visible = 0 : Catch : End Try
                chg.Add(New With {.param = "clearFill", .value = "True", .cells = If(had, 1, 0)})
            End If
            If borderColor IsNot Nothing Then
                Dim bc = OleColorFromArgb(CStr(borderColor))
                For Each b As Object In cell.Borders
                    b.ForeColor.RGB = bc
                    b.Visible = -1
                Next
                chg.Add(New With {.param = "borderColor", .value = CStr(borderColor), .cells = 1})
            End If
            Dim summaryC As String = "第" & row & "行第" & col & "列：已应用 " & chg.Count & " 项格式"
            LogBridge("FormatSelectedTableCell 成功(单格) r=" & row & " c=" & col & " clearFill=" & If(clearFill Is Nothing, "∅", CStr(clearFill)) & " items=" & chg.Count)
            Return JsonConvert.SerializeObject(New With {
                .success = True, .scope = "cell",
                .row = row, .col = col,
                .changes = chg, .summary = summaryC})
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Function FindTableShape(slide As Object, shapeIndex As Integer) As Object
        Dim idx As Integer = 0
        For Each shp As Object In slide.Shapes
            If shp.HasTable Then
                idx += 1
                If shapeIndex <= 0 OrElse idx = shapeIndex Then Return shp
            End If
        Next
        Return Nothing
    End Function

    ''' <summary>PowerPoint 对齐枚举 PpParagraphAlignment：左=1 中=2 右=3（与 Word 不同，必须按本宿主枚举）。</summary>
    Protected Overrides Function MapAlign(s As String) As Integer
        If s Is Nothing Then Return 2
        Select Case LCase(s)
            Case "left" : Return 1
            Case "right" : Return 3
            Case Else : Return 2
        End Select
    End Function

    <ComVisible(True)>
    Public Function SetSlideBackground(color As String, index As Integer) As String
        Try
            Dim slide As Object
            If index > 0 Then
                Dim pres = GetActivePresentation()
                If pres Is Nothing Then Return Err("没有打开的演示文稿")
                If index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
                slide = pres.Slides(index)
            Else
                slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            End If
            slide.Background.Fill.Solid()
            slide.Background.Fill.ForeColor.RGB = OleColorFromArgb(color)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function AddChart(index As Integer, chartType As String, title As String, left As Object, top As Object, width As Object, height As Object) As String
        Try
            Dim slide As Object
            If index > 0 Then
                Dim pres = GetActivePresentation()
                If pres Is Nothing Then Return Err("没有打开的演示文稿")
                If index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
                slide = pres.Slides(index)
            Else
                slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            End If
            Dim ct = MapChartType(chartType)
            Dim L = If(left Is Nothing, 60.0F, CSng(left))
            Dim T = If(top Is Nothing, 120.0F, CSng(top))
            Dim W = If(width Is Nothing, 480.0F, CSng(width))
            Dim H = If(height Is Nothing, 300.0F, CSng(height))
            Dim shape = slide.Shapes.AddChart(ct, L, T, W, H)
            If title IsNot Nothing AndAlso CStr(title) <> "" Then
                Try
                    shape.Chart.HasTitle = True
                    shape.Chart.ChartTitle.Text = CStr(title)
                Catch
                End Try
            End If
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function ExportImage(index As Integer, filePath As String, format As String) As String
        Try
            Dim slide As Object
            If index > 0 Then
                Dim pres = GetActivePresentation()
                If pres Is Nothing Then Return Err("没有打开的演示文稿")
                If index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
                slide = pres.Slides(index)
            Else
                slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            End If
            If String.IsNullOrWhiteSpace(filePath) Then
                Dim desktop As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                If String.IsNullOrWhiteSpace(desktop) Then desktop = System.IO.Path.GetTempPath()
                Dim ts As String = DateTime.Now.ToString("yyyyMMdd_HHmmss")
                Dim ext = If(format Is Nothing, "png", LCase(format))
                If ext <> "jpg" AndAlso ext <> "jpeg" Then ext = "png"
                filePath = System.IO.Path.Combine(desktop, "slide_" & ts & "." & ext)
            End If
            Dim ext2 = If(format Is Nothing, "png", LCase(format))
            Dim filter = If(ext2 = "jpg" OrElse ext2 = "jpeg", "JPG", "PNG")
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
            slide.Export(filePath, filter)
            Return "{""success"":true,""path"":""" & Safe(filePath) & """,""message"":""成功导出到图片: " & Safe(filePath) & """}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Function MapChartType(s As String) As Integer
        If s Is Nothing Then Return 51
        Select Case LCase(s)
            Case "pie" : Return 5
            Case "bar" : Return 57
            Case "line" : Return 65
            Case "area" : Return 1
            Case "scatter" : Return -4169
            Case Else : Return 51
        End Select
    End Function

    <ComVisible(True)>
    Public Function SetSlideLayout(index As Integer, layout As String) As String
        Try
            Dim slide As Object
            If index > 0 Then
                Dim pres = GetActivePresentation()
                If pres Is Nothing Then Return Err("没有打开的演示文稿")
                If index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
                slide = pres.Slides(index)
            Else
                slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            End If
            slide.Layout = MapSlideLayout(layout)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Function MapSlideLayout(s As String) As Integer
        If s Is Nothing Then Return 2
        Select Case LCase(s)
            Case "title" : Return 1
            Case "section_header" : Return 3
            Case "two_column_text" : Return 4
            Case "title_only" : Return 11
            Case "blank" : Return 12
            Case "two_objects" : Return 7
            Case "picture_caption" : Return 10
            Case Else : Return 2
        End Select
    End Function

    <ComVisible(True)>
    Public Function SetChartData(index As Integer, shapeIndex As Integer, rowsJson As String) As String
        Try
            Dim slide As Object
            If index > 0 Then
                Dim pres = GetActivePresentation()
                If pres Is Nothing Then Return Err("没有打开的演示文稿")
                If index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
                slide = pres.Slides(index)
            Else
                slide = GetCurrentSlide()
                If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            End If
            Dim chartShape = FindChartShape(slide, shapeIndex)
            If chartShape Is Nothing Then Return Err("未找到图表形状")
            Dim chart = chartShape.Chart
            Dim data = Newtonsoft.Json.JsonConvert.DeserializeObject(Of List(Of List(Of Object)))(rowsJson)
            If data Is Nothing OrElse data.Count = 0 Then Return Err("rows 为空")
            chart.ChartData.Activate()
            Dim wb = chart.ChartData.Workbook
            Dim sheet = wb.Worksheets(1)
            sheet.Cells.Clear()
            For r As Integer = 0 To data.Count - 1
                For c As Integer = 0 To data(r).Count - 1
                    sheet.Cells(r + 1, c + 1).Value = data(r)(c)
                Next
            Next
            wb.Close()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Function FindChartShape(slide As Object, shapeIndex As Integer) As Object
        Dim idx As Integer = 0
        For Each shp As Object In slide.Shapes
            If shp.HasChart Then
                idx += 1
                If shapeIndex <= 0 OrElse idx = shapeIndex Then Return shp
            End If
        Next
        Return Nothing
    End Function

    <ComVisible(True)>
    Public Function InsertSmartArt(layoutName As String, left As Object, top As Object, width As Object, height As Object) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片")
            Dim layoutObj = GetSmartArtLayout(layoutName)
            Dim L = If(left Is Nothing, 60.0F, CSng(left))
            Dim T = If(top Is Nothing, 120.0F, CSng(top))
            Dim W = If(width Is Nothing, 400.0F, CSng(width))
            Dim H = If(height Is Nothing, 300.0F, CSng(height))
            slide.Shapes.AddSmartArt(layoutObj, L, T, W, H)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    Private Function GetSmartArtLayout(nameOrIndex As String) As Object
        Dim layouts = _pptApp.SmartArtLayouts
        If nameOrIndex IsNot Nothing Then
            Dim num As Integer
            If Integer.TryParse(nameOrIndex, num) AndAlso num > 0 AndAlso num <= CInt(layouts.Count) Then
                Return layouts(num)
            End If
            Dim key = LCase(nameOrIndex)
            For Each ly As Object In layouts
                If LCase(CStr(ly.Name)).Contains(key) Then Return ly
            Next
        End If
        Return layouts(1)
    End Function

    <ComVisible(True)>
    Public Overrides Function InsertTableData(rowsJson As String) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片（请切换到普通视图）")
            Dim data = ParseRows(rowsJson)
            If data Is Nothing Then Return Err("表格数据为空或格式不正确")
            Dim rowCount = data.GetLength(1)
            Dim colCount = data.GetLength(0)
            Dim shape = slide.Shapes.AddTable(rowCount, colCount, 60.0F, 120.0F, 600.0F, 30.0F * rowCount)
            For r As Integer = 1 To rowCount
                For c As Integer = 1 To colCount
                    shape.Table.Cell(r, c).Shape.TextFrame.TextRange.Text = If(data(c - 1, r - 1), "")
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
    Public Function SetCurrentSlideText(text As String) As String
        Return AddTextbox(text)
    End Function

    <ComVisible(True)>
    Public Overrides Function GetSelectedRangeData() As String
        Try
            Dim idx As Integer = 0
            Integer.TryParse(GetCurrentSlideIndex(), idx)
            Dim slideText = GetSlideText(idx)
            Dim sel = GetSelectedText()
            ' 必须用 JsonConvert：旧手拼 JSON 未给字符串加引号，前端 JSON.parse 恒失败 → 正文永不注入
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .host = "powerpoint",
                .slideIndex = idx,
                .slideText = slideText,
                .selectedText = sel
            })
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function GetSelection() As String
        Try
            Return GetSelectedText()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    <ComVisible(True)>
    Public Overrides Function GetMultiSelection() As String
        Try
            Return GetSelectedText()
        Catch ex As Exception
            Return "{""success"":false,""error"":""" & Safe(ex.Message) & """}"
        End Try
    End Function

    ' ---------------- 幻灯片导航 / 备注 / 顺序 ----------------
    <ComVisible(True)>
    Public Function GotoSlide(index As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            If index <= 0 OrElse index > CInt(pres.Slides.Count) Then Return Err("幻灯片序号超出范围")
            pres.Slides(index).Select()
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function SetSlideNotes(index As Integer, text As String) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            Dim idx = index
            If idx <= 0 Then idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then Return Err("无法确定幻灯片")
            Dim slide = pres.Slides(idx)
            Dim target As Object = Nothing
            For Each shp As Object In slide.NotesPage.Shapes
                If shp.HasTextFrame Then
                    target = shp.TextFrame.TextRange
                    Exit For
                End If
            Next
            If target Is Nothing Then Return Err("找不到备注占位符")
            target.Text = If(text, "")
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function GetSlideNotes(index As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return "{""success"":false,""error"":""没有打开的演示文稿""}"
            Dim idx = index
            If idx <= 0 Then idx = GetCurrentSlideIndexValue()
            If idx <= 0 Then Return "{""success"":false,""error"":""无法确定幻灯片""}"
            Dim slide = pres.Slides(idx)
            Dim sb As New StringBuilder()
            For Each shp As Object In slide.NotesPage.Shapes
                If shp.HasTextFrame AndAlso shp.TextFrame.HasText Then
                    sb.Append(CStr(shp.TextFrame.TextRange.Text))
                End If
            Next
            Return "{""success"":true,""notes"":" & Safe(sb.ToString()) & "}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function MoveSlide(fromIndex As Integer, toIndex As Integer) As String
        Try
            Dim pres = GetActivePresentation()
            If pres Is Nothing Then Return Err("没有打开的演示文稿")
            If fromIndex <= 0 OrElse fromIndex > CInt(pres.Slides.Count) Then Return Err("源幻灯片序号超出范围")
            If toIndex <= 0 OrElse toIndex > CInt(pres.Slides.Count) Then Return Err("目标幻灯片序号超出范围")
            pres.Slides(fromIndex).MoveTo(toIndex)
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    <ComVisible(True)>
    Public Function SetShapeText(shapeIndex As Integer, text As String) As String
        Try
            Dim slide = GetCurrentSlide()
            If slide Is Nothing Then Return Err("无法获取当前幻灯片（请切换到普通视图）")
            If shapeIndex <= 0 Then Return Err("shapeIndex 必须从 1 开始")
            If shapeIndex > CInt(slide.Shapes.Count) Then Return Err("形状序号超出范围")
            Dim shp = slide.Shapes(shapeIndex)
            If Not shp.HasTextFrame Then Return Err("该形状不是文本框，无法写入文本")
            shp.TextFrame.TextRange.Text = If(text, "")
            Return "{""success"":true}"
        Catch ex As Exception
            Return Err(ex.Message)
        End Try
    End Function

    ' ---------------- 内部 ----------------
    Private Function GetActivePresentation() As Object
        Try
            If _pptApp Is Nothing Then Return Nothing
            Return _pptApp.ActivePresentation
        Catch
            Return Nothing
        End Try
    End Function

    Private Function GetCurrentSlide() As Object
        Try
            If _pptApp Is Nothing Then Return Nothing
            ' 普通视图：ActiveWindow.View.Slide；幻灯片浏览/其他视图可能为空
            Dim slide = _pptApp.ActiveWindow.View.Slide
            If slide IsNot Nothing Then Return slide
        Catch
        End Try
        Return Nothing
    End Function

    Private Function GetCurrentSlideIndexValue() As Integer
        ' 主路径：普通视图且焦点在幻灯片窗格
        Try
            Dim slide = GetCurrentSlide()
            If slide IsNot Nothing Then Return CInt(slide.SlideIndex)
        Catch
        End Try
        ' 【新增回落】焦点不在幻灯片窗格 / 大纲或备注视图时 ActiveWindow.View.Slide 常为 Nothing，
        ' 旧实现直接返回 0 → GetSlideText 读不到当前页 → 上层被迫回落"全部幻灯片文本"，
        ' "总结当前幻灯片"退化成逐页罗列全部页（实测 11 页目录被整篇抄出）。
        ' 回落1：缩略图窗格中选中的幻灯片（Selection.Type=2 即 ppSelectionSlides）
        Try
            If _pptApp IsNot Nothing AndAlso _pptApp.ActiveWindow IsNot Nothing Then
                Dim sel = _pptApp.ActiveWindow.Selection
                If sel IsNot Nothing AndAlso CInt(sel.Type) = 2 Then
                    Return CInt(sel.SlideRange.SlideIndex)
                End If
            End If
        Catch
        End Try
        ' 回落2：正在放映时取放映窗口当前页
        Try
            If _pptApp IsNot Nothing AndAlso _pptApp.SlideShowWindows.Count >= 1 Then
                Return CInt(_pptApp.SlideShowWindows(1).View.Slide.SlideIndex)
            End If
        Catch
        End Try
        Return 0
    End Function
End Class