' ============================================================================
' HostMcpTools - Word / PowerPoint 的本地 MCP 工具集（source: word-local / ppt-local）
'
' · 工具定义来自 Config\word_tools.json / ppt_tools.json（与 Excel 版 excel_tools.json 同风格）
' · 执行时反射调用宿主 Bridge（WordBridge / PptBridge）的对应方法，与 WebView2 前端共用一套实现
' · 返回结构与 Excel 版一致：{success, result, taskCompleted}（前端 app-mcp.js 读取 result）
' ============================================================================
Option Strict Off
Option Explicit On

Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class HostMcpTools

    Private Class ParamDef
        Public Property name As String
        Public Property type As String
        Public Property required As Boolean
        Public Property description As String
    End Class

    Private Class ToolDef
        Public Property name As String
        Public Property layer As Integer
        Public Property description As String
        Public Property parameters As List(Of ParamDef)
        Public Property displayName As String
        Public Property intentKeywords As List(Of String)
        Public Property operationType As String
        Public Property isWrite As Boolean
        Public Property isDangerous As Boolean
    End Class

    Private Class ToolsFile
        Public Property version As String
        Public Property source As String
        Public Property tools As List(Of ToolDef)
    End Class

    Private ReadOnly _kind As OfficeHostKind
    Private ReadOnly _bridge As Object
    Private ReadOnly _source As String
    Private _cfg As ToolsFile

    Public Sub New(kind As OfficeHostKind, bridge As Object)
        _kind = kind
        _bridge = bridge
        _source = If(kind = OfficeHostKind.Word, "word-local", "ppt-local")
        LoadConfig()
    End Sub

    Private Sub LoadConfig()
        Try
            Dim fileName = If(_kind = OfficeHostKind.Word, "word_tools.json", "ppt_tools.json")
            Dim asmDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)
            Dim baseDir = AppDomain.CurrentDomain.BaseDirectory
            Dim candidates As New List(Of String)()
            For Each d As String In New String() {asmDir, baseDir, Path.GetDirectoryName(asmDir)}
                If Not String.IsNullOrEmpty(d) Then
                    candidates.Add(Path.Combine(d, "Config", fileName))
                    candidates.Add(Path.Combine(d, fileName))
                End If
            Next
            Dim found As String = Nothing
            For Each c As String In candidates
                If System.IO.File.Exists(c) Then found = c : Exit For
            Next
            Try
                Dim dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ExcelAIVSTOAddIn")
                If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
                System.IO.File.AppendAllText(Path.Combine(dir, "host_bridge.log"),
                    DateTime.Now.ToString("HH:mm:ss") & " [" & _kind.ToString() & "] HostMcpTools file=" & fileName & " found=" & (found IsNot Nothing) & If(found IsNot Nothing, " path=" & found, "") & vbCrLf)
            Catch
            End Try
            If found IsNot Nothing Then
                _cfg = JsonConvert.DeserializeObject(Of ToolsFile)(System.IO.File.ReadAllText(found))
            End If
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine("[HostMcpTools] 工具定义加载失败: " & ex.Message)
        End Try
    End Sub

    ' ------------------------------------------------------------------
    ' 工具清单
    ' ------------------------------------------------------------------
    Public Function GetToolsJson() As String
        Dim arr As New JArray()
        If _cfg IsNot Nothing AndAlso _cfg.tools IsNot Nothing Then
            For Each t As ToolDef In _cfg.tools
                Dim o As New JObject()
                o("name") = t.name
                o("displayName") = If(String.IsNullOrEmpty(t.displayName), t.name, t.displayName)
                o("source") = If(String.IsNullOrEmpty(_cfg.source), _source, _cfg.source)
                o("description") = "【" & If(String.IsNullOrEmpty(t.displayName), t.name, t.displayName) & "】" & If(t.description, "")
                Dim ps As New JArray()
                If t.parameters IsNot Nothing Then
                    For Each pd As ParamDef In t.parameters
                        Dim po As New JObject()
                        po("name") = pd.name
                        po("type") = If(pd.type, "string")
                        po("description") = If(pd.description, "")
                        po("required") = pd.required
                        ps.Add(po)
                    Next
                End If
                o("parameters") = ps
                Dim kws As New JArray()
                If t.intentKeywords IsNot Nothing Then
                    For Each k As String In t.intentKeywords
                        kws.Add(k)
                    Next
                End If
                o("intentKeywords") = kws
                o("isWrite") = t.isWrite
                o("isDangerous") = t.isDangerous
                arr.Add(o)
            Next
        End If
        Dim root As New JObject()
        root("success") = True
        root("tools") = arr
        root("count") = arr.Count
        Return root.ToString(Formatting.None)
    End Function

    ' ------------------------------------------------------------------
    ' 工具执行
    ' ------------------------------------------------------------------
    Public Function CallTool(toolName As String, argumentsJson As String) As String
        Try
            Dim args As JObject = Nothing
            If Not String.IsNullOrWhiteSpace(argumentsJson) Then
                Try
                    args = JObject.Parse(argumentsJson)
                Catch
                    args = Nothing
                End Try
            End If

            Dim result As Object = Nothing
            Select Case (If(toolName, String.Empty)).ToLowerInvariant()

                ' ---------------- Word ----------------
                Case "word_get_document_info"
                    result = CallBridge("GetDocumentInfo")
                Case "word_get_selection_text"
                    result = CallBridge("GetSelectionText")
                Case "word_get_document_text"
                    result = CallBridge("GetDocumentText", GetInt(args, "maxChars", 20000))
                Case "word_insert_text"
                    result = CallBridge("InsertText", GetStr(args, "text"))
                Case "word_replace_selection"
                    result = CallBridge("ReplaceSelectionText", GetStr(args, "text"))
                Case "word_append_text"
                    result = CallBridge("AppendToDocument", GetStr(args, "text"))
                Case "word_save"
                    result = CallBridge("SaveDocument", GetStr(args, "filePath"))
                Case "word_insert_table"
                    result = CallBridge("InsertTableData", GetStr(args, "rows"))
                Case "word_find_replace"
                    result = CallBridge("FindReplace", GetStr(args, "findText"), GetStr(args, "replaceText"), GetBool(args, "matchCase", False))
                Case "word_format_selection"
                    result = CallBridge("FormatSelection", GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"), GetStr(args, "style"))
                Case "word_insert_picture"
                    result = CallBridge("InsertPicture", GetStr(args, "filePath"), GetOpt(args, "width"), GetOpt(args, "height"))
                Case "word_delete_selection"
                    result = CallBridge("DeleteSelection")
                Case "word_export_pdf"
                    result = CallBridge("ExportPdf", GetStr(args, "filePath"))
                Case "word_format_table_cell"
                    result = CallBridge("FormatTableCell", GetInt(args, "tableIndex", 0), GetInt(args, "row", 0), GetInt(args, "col", 0), GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"), GetOpt(args, "align"), GetOpt(args, "fillColor"), GetOpt(args, "borderColor"), GetOpt(args, "highlightColor"))
                Case "word_format_selected_table_cell"
                    result = CallBridge("FormatSelectedTableCell", GetInt(args, "row", 0), GetInt(args, "col", 0), GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"), GetOpt(args, "align"), GetOpt(args, "fillColor"), GetOpt(args, "borderColor"), GetOpt(args, "headerBold"), GetOpt(args, "headerFillColor"), GetOpt(args, "clearFill"), GetOpt(args, "highlightColor"), GetOpt(args, "zebra"), GetOpt(args, "borderStyle"), GetOpt(args, "headerFontColor"), GetOpt(args, "zebraColor"), GetOpt(args, "firstColBold"), GetOpt(args, "lastRowBold"), GetOpt(args, "bodyAlign"), GetOpt(args, "rowHeight"), GetOpt(args, "valign"), GetOpt(args, "headerItalic"))
                Case "word_insert_page_break"
                    result = CallBridge("InsertPageBreak")
                Case "word_insert_hyperlink"
                    result = CallBridge("InsertHyperlink", GetStr(args, "address"), GetStr(args, "text"), GetOpt(args, "displayText"))
                Case "word_insert_toc"
                    result = CallBridge("InsertToc")
                Case "word_insert_section_break"
                    result = CallBridge("InsertSectionBreak", GetStr(args, "breakType"))
                Case "word_insert_comment"
                    result = CallBridge("InsertComment", GetStr(args, "text"))
                Case "word_insert_bookmark"
                    result = CallBridge("InsertBookmark", GetStr(args, "name"))
                Case "word_insert_footnote"
                    result = CallBridge("InsertFootnote", GetStr(args, "text"))
                Case "word_insert_endnote"
                    result = CallBridge("InsertEndnote", GetStr(args, "text"))
                Case "word_apply_numbering"
                    result = CallBridge("ApplyNumbering")
                Case "word_apply_bullets"
                    result = CallBridge("ApplyBullets")
                Case "word_set_page_orientation"
                    result = CallBridge("SetPageOrientation", GetStr(args, "orientation"))
                Case "word_insert_header_text"
                    result = CallBridge("InsertHeaderText", GetStr(args, "text"))
                Case "word_insert_footer_text"
                    result = CallBridge("InsertFooterText", GetStr(args, "text"))
                Case "word_update_fields"
                    result = CallBridge("UpdateFields")
                Case "word_find_text"
                    result = CallBridge("FindText", GetStr(args, "findText"))

                    ' ---------------- PowerPoint ----------------
                Case "ppt_get_presentation_info"
                    result = CallBridge("GetPresentationInfo")
                Case "ppt_get_slide_count"
                    result = CallBridge("GetSlideCount")
                Case "ppt_get_slide_text"
                    result = CallBridge("GetSlideText", GetInt(args, "slideIndex", 0))
                Case "ppt_get_all_slides_text"
                    result = CallBridge("GetAllSlidesTextJson", GetInt(args, "maxChars", 20000))
                Case "ppt_get_selected_text"
                    result = CallBridge("GetSelectedText")
                Case "ppt_add_slide"
                    result = CallBridge("AddSlide", GetInt(args, "layoutIndex", 12))
                Case "ppt_add_title_content_slide"
                    result = CallBridge("AddTitleAndContentSlide", GetStr(args, "title"), GetStr(args, "content"))
                Case "ppt_add_textbox"
                    result = CallBridge("AddTextbox", GetStr(args, "text"))
                Case "ppt_insert_table"
                    result = CallBridge("InsertTableData", GetStr(args, "rows"))
                Case "ppt_save"
                    result = CallBridge("SavePresentation", GetStr(args, "filePath"))
                Case "ppt_find_replace"
                    result = CallBridge("FindReplace", GetStr(args, "findText"), GetStr(args, "replaceText"), GetBool(args, "matchCase", False))
                Case "ppt_format_text"
                    result = CallBridge("FormatText", GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"))
                Case "ppt_insert_picture"
                    result = CallBridge("InsertPicture", GetStr(args, "filePath"), GetOpt(args, "left"), GetOpt(args, "top"), GetOpt(args, "width"), GetOpt(args, "height"))
                Case "ppt_delete_slide"
                    result = CallBridge("DeleteSlide", GetInt(args, "slideIndex", 0))
                Case "ppt_duplicate_slide"
                    result = CallBridge("DuplicateSlide", GetInt(args, "slideIndex", 0))
                Case "ppt_delete_shape"
                    result = CallBridge("DeleteSelectedShape")
                Case "ppt_export_pdf"
                    result = CallBridge("ExportPdf", GetStr(args, "filePath"))
                Case "ppt_format_table_cell"
                    result = CallBridge("FormatTableCell", GetInt(args, "shapeIndex", 0), GetInt(args, "row", 0), GetInt(args, "col", 0), GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"), GetOpt(args, "align"), GetOpt(args, "fillColor"), GetOpt(args, "borderColor"), GetOpt(args, "highlightColor"))
                Case "ppt_format_selected_table_cell"
                    result = CallBridge("FormatSelectedTableCell", GetInt(args, "row", 0), GetInt(args, "col", 0), GetOpt(args, "bold"), GetOpt(args, "italic"), GetOpt(args, "fontSize"), GetOpt(args, "fontName"), GetOpt(args, "fontColor"), GetOpt(args, "align"), GetOpt(args, "fillColor"), GetOpt(args, "borderColor"), GetOpt(args, "headerBold"), GetOpt(args, "headerFillColor"), GetOpt(args, "clearFill"), GetOpt(args, "highlightColor"), GetOpt(args, "zebra"), GetOpt(args, "borderStyle"), GetOpt(args, "headerFontColor"), GetOpt(args, "zebraColor"), GetOpt(args, "firstColBold"), GetOpt(args, "lastRowBold"), GetOpt(args, "bodyAlign"), GetOpt(args, "rowHeight"), GetOpt(args, "valign"), GetOpt(args, "headerItalic"))
                Case "ppt_set_slide_background"
                    result = CallBridge("SetSlideBackground", GetStr(args, "color"), GetInt(args, "index", 0))
                Case "ppt_add_chart"
                    result = CallBridge("AddChart", GetInt(args, "index", 0), GetStr(args, "chartType"), GetStr(args, "title"), GetOpt(args, "left"), GetOpt(args, "top"), GetOpt(args, "width"), GetOpt(args, "height"))
                Case "ppt_export_image"
                    result = CallBridge("ExportImage", GetInt(args, "index", 0), GetStr(args, "filePath"), GetStr(args, "format"))
                Case "ppt_set_slide_layout"
                    result = CallBridge("SetSlideLayout", GetInt(args, "index", 0), GetStr(args, "layout"))
                Case "ppt_set_chart_data"
                    result = CallBridge("SetChartData", GetInt(args, "index", 0), GetInt(args, "shapeIndex", 0), GetStr(args, "rows"))
                Case "ppt_insert_smartart"
                    result = CallBridge("InsertSmartArt", GetStr(args, "layoutName"), GetOpt(args, "left"), GetOpt(args, "top"), GetOpt(args, "width"), GetOpt(args, "height"))
                Case "ppt_goto_slide"
                    result = CallBridge("GotoSlide", GetInt(args, "index", 0))
                Case "ppt_set_slide_notes"
                    result = CallBridge("SetSlideNotes", GetInt(args, "index", 0), GetStr(args, "text"))
                Case "ppt_get_slide_notes"
                    result = CallBridge("GetSlideNotes", GetInt(args, "index", 0))
                Case "ppt_move_slide"
                    result = CallBridge("MoveSlide", GetInt(args, "fromIndex", 0), GetInt(args, "toIndex", 0))
                Case "ppt_set_shape_text"
                    result = CallBridge("SetShapeText", GetInt(args, "shapeIndex", 0), GetStr(args, "text"))

                Case Else
                    Return "{""success"":false,""error"":""未实现的工具: " & SafeStr(toolName) & """}"
            End Select

            Dim text = If(result Is Nothing, "", CStr(result))
            Dim payload As JToken
            Dim trimmed = text.Trim()
            If trimmed.StartsWith("{") OrElse trimmed.StartsWith("[") Then
                Try
                    payload = JToken.Parse(trimmed)
                Catch
                    payload = New JValue(text)
                End Try
            Else
                payload = New JValue(text)
            End If

            Dim root As New JObject()
            root("success") = True
            root("result") = payload
            root("taskCompleted") = True
            Return root.ToString(Formatting.None)

        Catch ex As Exception
            Dim root As New JObject()
            root("success") = False
            root("error") = ex.Message
            Return root.ToString(Formatting.None)
        End Try
    End Function

    ' ------------------------------------------------------------------
    ' 反射调用宿主 Bridge 方法（与前端 JS 调用同一批 ComVisible 方法）
    ' ------------------------------------------------------------------
    Private Function CallBridge(methodName As String, ParamArray args As Object()) As Object
        If _bridge Is Nothing Then Throw New Exception("宿主桥未初始化")
        Dim mi = _bridge.GetType().GetMethod(methodName)
        If mi Is Nothing Then Throw New Exception("当前宿主不支持操作: " & methodName)
        Return mi.Invoke(_bridge, args)
    End Function

    Private Function GetStr(args As JObject, key As String) As String
        If args Is Nothing Then Return ""
        Dim t As JToken = Nothing
        If args.TryGetValue(key, t) AndAlso t IsNot Nothing Then
            If t.Type = JTokenType.String Then Return CStr(t)
            Return t.ToString(Formatting.None)
        End If
        Return ""
    End Function

    Private Function GetInt(args As JObject, key As String, def As Integer) As Integer
        If args Is Nothing Then Return def
        Dim t As JToken = Nothing
        If args.TryGetValue(key, t) AndAlso t IsNot Nothing Then
            Try
                Return CInt(t)
            Catch
                Return def
            End Try
        End If
        Return def
    End Function

    Private Function GetBool(args As JObject, key As String, def As Boolean) As Boolean
        If args Is Nothing Then Return def
        Dim t As JToken = Nothing
        If args.TryGetValue(key, t) AndAlso t IsNot Nothing AndAlso t.Type = JTokenType.Boolean Then
            Return CBool(t)
        End If
        Return def
    End Function

    Private Function GetOpt(args As JObject, key As String) As Object
        If args Is Nothing Then Return Nothing
        Dim t As JToken = Nothing
        If args.TryGetValue(key, t) AndAlso t IsNot Nothing AndAlso t.Type <> JTokenType.Null Then
            Select Case t.Type
                Case JTokenType.Boolean : Return CBool(t)
                Case JTokenType.Integer : Return CInt(t)
                Case JTokenType.Float : Return CDbl(t)
                Case Else : Return CStr(t)
            End Select
        End If
        Return Nothing
    End Function

    Private Function SafeStr(s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace("\", "\\").Replace("""", "\""")
    End Function
End Class
