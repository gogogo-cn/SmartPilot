Option Explicit On
Imports System
Imports System.Collections.Concurrent
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Services
Imports ExcelAIVSTOAddIn.Utils

Namespace Services.Mcp

    Partial Public Class PptMcpTools

        ' ------------------------------------------------------------
        ' PPT 分步构建工具集：AI 逐页自由拼装 PPT
        ' 流程：ppt_project_start 创建项目 → ppt_add_*_slide 逐页添加
        '       → （可选）ppt_project_list/update/remove 调整
        '       → ppt_project_build 渲染生成 .pptx
        ' 页面规格与 ppt_generate 的 slides 数组元素完全同构，渲染复用 PptGenerationService。
        ' ------------------------------------------------------------

        Private Shared ReadOnly _pptProjects As New ConcurrentDictionary(Of String, PptProject)(StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly _pptProjectLock As New Object()

        ''' <summary>PPT 构建项目上下文（页面规格 + 项目级数据来源）</summary>
        Private Class PptProject
            Public Property ProjectId As String
            Public Property ReportTitle As String
            Public Property Theme As String
            Public Property Orientation As String
            Public Property Requirement As String
            Public Property RangeAddress As String
            Public Property Headers As List(Of String)
            Public Property Rows As List(Of Dictionary(Of String, Object))
            Public Property Slides As JArray
            Public Property CreatedAt As Date
            Public Property LastAccess As Date
        End Class

        ''' <summary>获取项目并按访问刷新时间戳；不存在返回 Nothing</summary>
        Private Shared Function GetProject(projectId As String) As PptProject
            If String.IsNullOrWhiteSpace(projectId) Then Return Nothing
            Dim p As PptProject = Nothing
            If _pptProjects.TryGetValue(projectId, p) Then
                p.LastAccess = Date.Now
                Return p
            End If
            Return Nothing
        End Function

        ''' <summary>清理超过 2 小时未访问的项目，防内存泄漏</summary>
        Private Shared Sub CleanupExpiredProjects()
            Dim cutoff = Date.Now.AddHours(-2)
            For Each kv In _pptProjects
                If kv.Value.LastAccess < cutoff Then
                    _pptProjects.TryRemove(kv.Key, Nothing)
                End If
            Next
        End Sub

        ''' <summary>校验项目存在；失败时在 result 写入错误并返回 Nothing</summary>
        Private Function RequireProject(args As Dictionary(Of String, Object), result As ToolsCallResult) As PptProject
            Dim projectId = GetArg(Of String)(args, "projectId", "")
            Dim p = GetProject(projectId)
            If p Is Nothing Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {
                    .Text = "PPT 项目不存在: " & projectId & "。请先用 ppt_project_start 创建项目（并拿到 projectId）"
                })
            End If
            Return p
        End Function

        ''' <summary>成功结果（text 为 JSON 字符串，与既有工具一致）</summary>
        Private Shared Function PptOk(obj As JObject) As ToolsCallResult
            Return New ToolsCallResult() With {
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = obj.ToString(Formatting.None)}}
            }
        End Function

        Private Shared Function PptErr(message As String) As ToolsCallResult
            Return New ToolsCallResult() With {
                .IsError = True,
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = message}}
            }
        End Function

        ''' <summary>
        ''' 把参数复制进页面规格：字符串尝试 JSON/bool 推断，其余原样保留
        ''' </summary>
        Private Shared Sub CopySpecField(args As Dictionary(Of String, Object), spec As JObject, key As String)
            If args Is Nothing OrElse Not args.ContainsKey(key) OrElse args(key) Is Nothing Then Return
            Dim v = args(key)
            If TypeOf v Is String Then
                Dim s = Convert.ToString(v).Trim()
                If s.Length = 0 Then Return
                If s.Equals("true", StringComparison.OrdinalIgnoreCase) Then
                    spec(key) = True
                ElseIf s.Equals("false", StringComparison.OrdinalIgnoreCase) Then
                    spec(key) = False
                ElseIf s.StartsWith("[") OrElse s.StartsWith("{") Then
                    Try
                        spec(key) = JToken.Parse(s)
                        Return
                    Catch
                        spec(key) = s
                    End Try
                Else
                    spec(key) = s
                End If
            Else
                spec(key) = JToken.FromObject(v)
            End If
        End Sub

        ''' <summary>页型白名单（对齐 SUPPORTED_SLIDE_TYPES）</summary>
        Private Shared Function IsSupportedSlideType(slideType As String) As Boolean
            Return PptGenerationService.IsSupportedSlideType(slideType)
        End Function

        ''' <summary>项目级数据不足时，允许本次调用补充；仍无则报错（数据页专用）</summary>
        Private Function EnsureProjectData(project As PptProject, args As Dictionary(Of String, Object),
                                           result As ToolsCallResult, slideType As String) As Boolean
            If String.IsNullOrWhiteSpace(project.RangeAddress) AndAlso
               (project.Headers Is Nothing OrElse project.Headers.Count = 0) Then
                Dim rangeAddress = GetArg(Of String)(args, "rangeAddress", "")
                Dim headers = ResolveStringList(GetArg(Of Object)(args, "headers", Nothing))
                If Not String.IsNullOrWhiteSpace(rangeAddress) Then project.RangeAddress = rangeAddress
                If headers.Count > 0 Then
                    project.Headers = headers
                    project.Rows = ResolveRows(GetArg(Of Object)(args, "rows", Nothing))
                End If
            End If
            If String.IsNullOrWhiteSpace(project.RangeAddress) AndAlso
               (project.Headers Is Nothing OrElse project.Headers.Count = 0) Then
                result.IsError = True
                result.Content.Add(New McpTextContent With {
                    .Text = slideType & " 页需要数据来源。请在 ppt_project_start 传 rangeAddress（Excel 选区）或 headers+rows（列名数组+数据行数组），或在本次调用中补充。"
                })
                Return False
            End If
            Return True
        End Function

        ''' <summary>通用动画字段复制</summary>
        Private Shared Sub CopyAnimationFields(args As Dictionary(Of String, Object), spec As JObject)
            CopySpecField(args, spec, "transition")
            CopySpecField(args, spec, "animation")
            CopySpecField(args, spec, "animate")
        End Sub

        ''' <summary>添加一页到项目（加锁保护 Slides）</summary>
        Private Shared Sub AddSlideToProject(project As PptProject, spec As JObject, index As Integer)
            SyncLock _pptProjectLock
                If index >= 0 AndAlso index < project.Slides.Count Then
                    project.Slides(index) = spec
                Else
                    project.Slides.Add(spec)
                End If
            End SyncLock
        End Sub

        ''' <summary>页面添加成功响应</summary>
        Private Shared Function SlideAddedResponse(project As PptProject, spec As JObject) As ToolsCallResult
            Dim resp As New JObject()
            resp("success") = True
            resp("projectId") = project.ProjectId
            resp("slideIndex") = project.Slides.Count
            resp("slideCount") = project.Slides.Count
            resp("message") = "已添加 " & spec.Value(Of String)("type") & " 页（第 " & project.Slides.Count.ToString() &
                              " 页）。继续用 ppt_add_*_slide 添加，或 ppt_project_build 生成 PPT。"
            Return PptOk(resp)
        End Function

        '#Region "项目生命周期"

        ''' <summary>
        ''' ppt_project_start：创建 PPT 构建项目，返回 projectId。之后用 ppt_add_*_slide 逐页构建。
        ''' </summary>
        Private Function ExecuteProjectStart(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                CleanupExpiredProjects()
                Dim p As New PptProject() With {
                    .ProjectId = Guid.NewGuid().ToString("N"),
                    .ReportTitle = GetArg(Of String)(args, "reportTitle", "数据分析报告"),
                    .Theme = GetArg(Of String)(args, "theme", ""),
                    .Orientation = GetArg(Of String)(args, "orientation", ""),
                    .Requirement = GetArg(Of String)(args, "requirement", ""),
                    .RangeAddress = GetArg(Of String)(args, "rangeAddress", ""),
                    .Headers = ResolveStringList(GetArg(Of Object)(args, "headers", Nothing)),
                    .Rows = ResolveRows(GetArg(Of Object)(args, "rows", Nothing)),
                    .Slides = New JArray(),
                    .CreatedAt = Date.Now,
                    .LastAccess = Date.Now
                }
                _pptProjects(p.ProjectId) = p

                Dim resp As New JObject()
                resp("success") = True
                resp("projectId") = p.ProjectId
                resp("reportTitle") = p.ReportTitle
                resp("slideCount") = 0
                resp("message") = "PPT 项目已创建（projectId=" & p.ProjectId & "）。接下来用 ppt_add_cover_slide / ppt_add_text_slide / ppt_add_chart_slide 等逐页添加内容，最后用 ppt_project_build 生成 .pptx。"
                result.Content.Add(New McpTextContent With {.Text = resp.ToString(Formatting.None)})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "创建 PPT 项目失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_project_list：查看项目当前页面结构（页序/类型/标题），供 AI 确认进度或定位要改的页。
        ''' </summary>
        Private Function ExecuteProjectList(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result

                Dim slidesArr As New JArray()
                For i As Integer = 0 To p.Slides.Count - 1
                    Dim s = TryCast(p.Slides(i), JObject)
                    Dim item As New JObject()
                    item("index") = i + 1
                    item("type") = If(s Is Nothing, "", s.Value(Of String)("type"))
                    item("title") = If(s Is Nothing, "", s.Value(Of String)("title"))
                    slidesArr.Add(item)
                Next

                Dim resp As New JObject()
                resp("success") = True
                resp("projectId") = p.ProjectId
                resp("reportTitle") = p.ReportTitle
                resp("slideCount") = p.Slides.Count
                resp("slides") = slidesArr
                resp("message") = "共 " & p.Slides.Count.ToString() & " 页。用 ppt_project_update 修改某页（index 从 1 起）、ppt_project_remove 删除某页。"
                result.Content.Add(New McpTextContent With {.Text = resp.ToString(Formatting.None)})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "查看项目失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_project_update：合并修改指定页（index 从 1 起）。传入字段覆盖该页对应字段，type 也可改。
        ''' </summary>
        Private Function ExecuteProjectUpdate(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim index = GetArg(Of Integer)(args, "index", 0)
                If index < 1 OrElse index > p.Slides.Count Then
                    Return PptErr("index 越界：当前共 " & p.Slides.Count.ToString() & " 页，index 应为 1~" & p.Slides.Count.ToString())
                End If
                Dim page = TryCast(p.Slides(index - 1), JObject)
                If page Is Nothing Then
                    Return PptErr("第 " & index.ToString() & " 页不是对象规格，无法修改")
                End If

                Dim changed As New JArray()
                SyncLock _pptProjectLock
                    For Each kv In args
                        Dim key = kv.Key
                        If key.Equals("projectId", StringComparison.OrdinalIgnoreCase) OrElse
                           key.Equals("index", StringComparison.OrdinalIgnoreCase) Then
                            Continue For
                        End If
                        CopySpecField(args, page, key)
                        changed.Add(key)
                    Next
                End SyncLock

                Dim resp As New JObject()
                resp("success") = True
                resp("projectId") = p.ProjectId
                resp("slideIndex") = index
                resp("changed") = changed
                resp("message") = "已更新第 " & index.ToString() & " 页字段：" & String.Join(", ", changed.ToObject(Of String())())
                result.Content.Add(New McpTextContent With {.Text = resp.ToString(Formatting.None)})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "修改页面失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_project_remove：删除指定页（index 从 1 起）。
        ''' </summary>
        Private Function ExecuteProjectRemove(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim index = GetArg(Of Integer)(args, "index", 0)
                If index < 1 OrElse index > p.Slides.Count Then
                    Return PptErr("index 越界：当前共 " & p.Slides.Count.ToString() & " 页，index 应为 1~" & p.Slides.Count.ToString())
                End If
                SyncLock _pptProjectLock
                    p.Slides.RemoveAt(index - 1)
                End SyncLock

                Dim resp As New JObject()
                resp("success") = True
                resp("projectId") = p.ProjectId
                resp("slideCount") = p.Slides.Count
                resp("message") = "已删除第 " & index.ToString() & " 页，剩余 " & p.Slides.Count.ToString() & " 页。"
                result.Content.Add(New McpTextContent With {.Text = resp.ToString(Formatting.None)})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "删除页面失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_project_build：按项目页面规格渲染生成 .pptx 并自动打开。
        ''' </summary>
        Private Function ExecuteProjectBuild(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                If p.Slides.Count = 0 Then
                    Return PptErr("项目还没有任何页面。请先用 ppt_add_cover_slide / ppt_add_text_slide 等工具逐页添加内容。")
                End If

                Dim payload As New JObject()
                payload("reportTitle") = p.ReportTitle
                payload("slides") = p.Slides.DeepClone()
                If Not String.IsNullOrWhiteSpace(p.Theme) Then payload("theme") = p.Theme
                If Not String.IsNullOrWhiteSpace(p.Orientation) Then payload("orientation") = p.Orientation
                If Not String.IsNullOrWhiteSpace(p.Requirement) Then payload("requirement") = p.Requirement
                If Not String.IsNullOrWhiteSpace(p.RangeAddress) Then payload("rangeAddress") = p.RangeAddress
                If p.Headers IsNot Nothing AndAlso p.Headers.Count > 0 Then payload("headers") = JArray.FromObject(p.Headers)
                If p.Rows IsNot Nothing AndAlso p.Rows.Count > 0 Then payload("rows") = JArray.FromObject(p.Rows)

                Dim resp = PptGenerationService.Generate(payload.ToString(Formatting.None))
                result.Content.Add(New McpTextContent With {.Text = resp})
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "PPT 生成失败: " & ex.Message})
            End Try
            Return result
        End Function

        '#End Region

        '#Region "页型添加工具"

        ''' <summary>
        ''' ppt_add_cover_slide：添加封面/结尾类页面（type: cover/title/end/closing）
        ''' </summary>
        Private Function ExecuteAddCoverSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim slideType = GetArg(Of String)(args, "type", "cover").ToLowerInvariant()
                If Not IsSupportedSlideType(slideType) Then Return PptErr("不支持的页型: " & slideType)
                Dim spec As New JObject()
                spec("type") = slideType
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "subtitle")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加封面页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_text_slide：添加文本/要点页（type: text/bullets）
        ''' </summary>
        Private Function ExecuteAddTextSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim slideType = GetArg(Of String)(args, "type", "text").ToLowerInvariant()
                If Not IsSupportedSlideType(slideType) Then Return PptErr("不支持的页型: " & slideType)
                Dim spec As New JObject()
                spec("type") = slideType
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "bullets")
                CopySpecField(args, spec, "content")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加文本页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_table_slide：添加数据表格页
        ''' </summary>
        Private Function ExecuteAddTableSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                If Not EnsureProjectData(p, args, result, "table") Then Return result
                Dim spec As New JObject()
                spec("type") = "table"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加表格页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_chart_slide：添加图表页（柱/条/折线/饼/环/面积/雷达/散点/组合图/气泡图）
        ''' </summary>
        Private Function ExecuteAddChartSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                If Not EnsureProjectData(p, args, result, "chart") Then Return result
                Dim spec As New JObject()
                spec("type") = "chart"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "chartType")
                CopySpecField(args, spec, "xField")
                CopySpecField(args, spec, "series")
                CopySpecField(args, spec, "secondaryAxis")
                CopySpecField(args, spec, "lineSeries")
                CopySpecField(args, spec, "insight")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加图表页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_kpi_slide：添加核心指标页（items: [{label,value,highlight}]）
        ''' </summary>
        Private Function ExecuteAddKpiSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                If Not EnsureProjectData(p, args, result, "kpi") Then Return result
                Dim spec As New JObject()
                spec("type") = "kpi"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "items")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加指标页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_structure_slide：添加结构页（type: agenda/section/timeline/comparison）
        ''' </summary>
        Private Function ExecuteAddStructureSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim slideType = GetArg(Of String)(args, "type", "agenda").ToLowerInvariant()
                If Not IsSupportedSlideType(slideType) Then Return PptErr("不支持的页型: " & slideType)
                Dim spec As New JObject()
                spec("type") = slideType
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "subtitle")
                CopySpecField(args, spec, "number")
                CopySpecField(args, spec, "items")
                CopySpecField(args, spec, "leftTitle")
                CopySpecField(args, spec, "rightTitle")
                CopySpecField(args, spec, "leftItems")
                CopySpecField(args, spec, "rightItems")
                CopySpecField(args, spec, "variant")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加结构页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_shapes_slide：添加自由形状页（矩形/椭圆/箭头/流程图/标注/线条等）
        ''' </summary>
        Private Function ExecuteAddShapesSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim spec As New JObject()
                spec("type") = "shapes"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "items")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加形状页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_picture_slide：添加图片页（本地路径或 http(s) URL，支持圆角/边框/阴影）
        ''' </summary>
        Private Function ExecuteAddPictureSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim spec As New JObject()
                spec("type") = "picture"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "images")
                CopySpecField(args, spec, "src")
                CopySpecField(args, spec, "caption")
                CopySpecField(args, spec, "x")
                CopySpecField(args, spec, "y")
                CopySpecField(args, spec, "w")
                CopySpecField(args, spec, "h")
                CopySpecField(args, spec, "rounded")
                CopySpecField(args, spec, "border")
                CopySpecField(args, spec, "shadow")
                CopySpecField(args, spec, "fit")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加图片页失败: " & ex.Message})
            End Try
            Return result
        End Function

        ''' <summary>
        ''' ppt_add_smartart_slide：添加 SmartArt 替代页（layout: flow/cycle/hierarchy/matrix）
        ''' </summary>
        Private Function ExecuteAddSmartartSlide(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim result = New ToolsCallResult()
            Try
                Dim p = RequireProject(args, result)
                If p Is Nothing Then Return result
                Dim spec As New JObject()
                spec("type") = "smartart"
                CopySpecField(args, spec, "title")
                CopySpecField(args, spec, "layout")
                CopySpecField(args, spec, "items")
                CopyAnimationFields(args, spec)
                AddSlideToProject(p, spec, -1)
                Return SlideAddedResponse(p, spec)
            Catch ex As Exception
                result.IsError = True
                result.Content.Add(New McpTextContent With {.Text = "添加 SmartArt 页失败: " & ex.Message})
            End Try
            Return result
        End Function

        '#End Region

    End Class

End Namespace
