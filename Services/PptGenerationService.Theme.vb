Option Explicit On
Option Infer On
Option Strict Off

Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Linq
Imports System.Text
Imports System.Xml.Linq
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Newtonsoft.Json.Serialization
Imports ExcelAIVSTOAddIn.Utils
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

Partial Public Class PptGenerationService

        ''' <summary>
        ''' 生成版式画像的人类可读摘要（用于导入响应与前端提示）
        ''' </summary>
        Private Shared Function LayoutSummary(layout As PptLayoutProfile) As String
            If layout Is Nothing Then Return ""
            Dim parts As New List(Of String)()
            If layout.Title IsNot Nothing Then
                parts.Add("标题(" & layout.Title.Left & "," & layout.Title.Top & "," & layout.Title.Width & "x" & layout.Title.Height & ")")
            End If
            If layout.Body IsNot Nothing Then
                parts.Add("正文(" & layout.Body.Left & "," & layout.Body.Top & "," & layout.Body.Width & "x" & layout.Body.Height & ")")
            End If
            If layout.Table IsNot Nothing Then
                Dim colInfo As String = If(layout.Table.Columns Is Nothing, "0", layout.Table.Columns.Count.ToString())
                parts.Add("表格(" & layout.Table.Left & "," & layout.Table.Top & "," & layout.Table.Width & "x" & layout.Table.Height & "," & colInfo & "列)")
            End If
            If layout.Shapes IsNot Nothing AndAlso layout.Shapes.Count > 0 Then
                parts.Add("形状" & layout.Shapes.Count.ToString() & "个")
            End If
            If layout.Background IsNot Nothing Then
                Dim bgInfo As String = "母版背景"
                Select Case layout.Background.Kind
                    Case "gradient"
                        bgInfo &= "(渐变)"
                    Case "picture"
                        bgInfo &= "(背景图)"
                End Select
                parts.Add(bgInfo)
            End If
            If layout.MasterShapes IsNot Nothing AndAlso layout.MasterShapes.Count > 0 Then
                parts.Add("母版图形" & layout.MasterShapes.Count.ToString() & "个")
            End If
            Dim cap As PptCapacityProfile = CapacityOf(layout)
            If cap IsNot Nothing Then
                parts.Add("容量(要点≤" & cap.MaxBullets & "/表格≤" & cap.MaxTableRows & "行" & cap.MaxTableCols & "列/图表系列≤" & cap.MaxChartSeries & ")")
            End If
            Return String.Join(" ", parts)
        End Function

        ' ── 母版读取常量（完整母版结构：背景 + 背景图形 + 占位符回退）──
        Private Shared ReadOnly MASTER_RELS_PATH As String = "ppt/slideMasters/_rels/slideMaster1.xml.rels"
        Private Shared ReadOnly PKG_REL_NS As XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        Private Shared ReadOnly OFFICE_REL_NS As XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        Private Shared ReadOnly DRAWING_NS As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"

        ''' <summary>
        ''' 从填充节点提取颜色（solidFill/ln 内 srgbClr/sysClr），失败返回 -1
        ''' </summary>
        Private Shared Function ExtractFillColor(fill As XElement, ans As XNamespace) As Long
            If fill Is Nothing Then Return -1
            Dim srgb As XElement = fill.Descendants(ans + "srgbClr").FirstOrDefault()
            If srgb IsNot Nothing Then
                Dim v As String = srgb.Attribute("val")?.Value
                If v IsNot Nothing AndAlso v.Length = 6 Then Return Convert.ToInt32(v, 16)
            End If
            Dim sys As XElement = fill.Descendants(ans + "sysClr").FirstOrDefault()
            If sys IsNot Nothing Then
                Dim v As String = sys.Attribute("lastClr")?.Value
                If v IsNot Nothing AndAlso v.Length = 6 Then Return Convert.ToInt32(v, 16)
            End If
            Return -1
        End Function

        ''' <summary>
        ''' bgRef idx → clrScheme 主题色映射（1~12 或 1001~1012 → dk1/lt1/dk2/lt2/accent1-6/hlink/folHlink）
        ''' </summary>
        Private Shared Function ResolveBgRefColor(clrScheme As XElement, themeNs As XNamespace, idxValue As String) As Long
            If clrScheme Is Nothing Then Return -1
            Dim n As Integer = 0
            If Not Integer.TryParse(idxValue, n) OrElse n <= 0 Then Return -1
            If n > 1000 Then n -= 1000
            If n > 12 Then n = 12
            Static slotNames As String() = {"dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3",
                                            "accent4", "accent5", "accent6", "hlink", "folHlink"}
            Return ExtractColorFromScheme(clrScheme, themeNs, slotNames(n - 1))
        End Function

        ''' <summary>
        ''' 解析 rels 得到 embedId 对应的 zip 内媒体路径（如 ppt/media/image1.png），失败返回 Nothing
        ''' </summary>
        Private Shared Function ResolveMediaTarget(zip As ZipArchive, relsPath As String, embedId As String) As String
            If String.IsNullOrEmpty(embedId) Then Return Nothing
            Dim relsEntry As ZipArchiveEntry = zip.GetEntry(relsPath)
            If relsEntry Is Nothing Then Return Nothing
            Dim relsXml As String
            Using r As New StreamReader(relsEntry.Open(), Encoding.UTF8)
                relsXml = r.ReadToEnd()
            End Using
            Dim rdoc As XDocument = XDocument.Parse(relsXml)
            Dim rel As XElement = rdoc.Descendants(PKG_REL_NS + "Relationship").FirstOrDefault(
                Function(x) String.Equals(x.Attribute("Id")?.Value, embedId, StringComparison.OrdinalIgnoreCase))
            If rel Is Nothing Then Return Nothing
            Dim target As String = rel.Attribute("Target")?.Value
            If String.IsNullOrWhiteSpace(target) Then Return Nothing

            Dim baseDir As String = ""
            If Not target.StartsWith("/") Then
                baseDir = Path.GetDirectoryName(relsPath.Replace("\"c, "/"c)) & "/"
            End If
            Dim combined As String = Uri.UnescapeDataString(baseDir & target)
            ' 规范化 ../
            Dim parts As New List(Of String)()
            For Each seg In combined.Split("/"c)
                If seg = ".." Then
                    If parts.Count > 0 Then parts.RemoveAt(parts.Count - 1)
                ElseIf seg <> "." AndAlso seg <> "" Then
                    parts.Add(seg)
                End If
            Next
            If parts.Count = 0 Then Return Nothing
            Return String.Join("/", parts)
        End Function

        ''' <summary>
        ''' 复制 zip 内媒体到 mediaSaveDir（Config/ppt_media/{themeId}），返回相对 Config 的路径；失败返回 Nothing
        ''' </summary>
        Private Shared Function SaveMediaToConfig(zip As ZipArchive, zipMediaPath As String, mediaSaveDir As String) As String
            If String.IsNullOrEmpty(mediaSaveDir) Then Return Nothing
            Dim entry As ZipArchiveEntry = zip.GetEntry(zipMediaPath)
            If entry Is Nothing Then Return Nothing
            Try
                Directory.CreateDirectory(mediaSaveDir)
                Dim fileName As String = Path.GetFileName(zipMediaPath)
                Dim fullPath As String = Path.Combine(mediaSaveDir, fileName)
                If File.Exists(fullPath) Then
                    Dim nameOnly As String = Path.GetFileNameWithoutExtension(fileName)
                    Dim ext As String = Path.GetExtension(fileName)
                    Dim i As Integer = 1
                    Do While File.Exists(fullPath)
                        fullPath = Path.Combine(mediaSaveDir, nameOnly & "_" & i.ToString() & ext)
                        i += 1
                    Loop
                End If
                Using src = entry.Open()
                    Using dst As New FileStream(fullPath, FileMode.Create, FileAccess.Write)
                        src.CopyTo(dst)
                    End Using
                End Using
                Dim cfgDir As String = PathHelper.ConfigDirectory.TrimEnd("\"c) & "\"
                If fullPath.StartsWith(cfgDir, StringComparison.OrdinalIgnoreCase) Then
                    Return fullPath.Substring(cfgDir.Length).Replace("\"c, "/"c)
                End If
                Return fullPath
            Catch
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 从 slideMaster1.xml 选择性提取母版结构：
        ''' 1) 背景（bgPr 纯色/渐变/图片，bgRef 主题色映射兜底）
        ''' 2) 背景图形（sp/pic：跳过占位符、纯文本框、透明无内容形状，上限 16 个）
        ''' 3) 占位符回退（版式未提取到标题/正文时用母版位置）
        ''' </summary>
        Private Shared Sub ExtractMasterProfile(zip As ZipArchive,
                                                profile As PptLayoutProfile,
                                                scaleX As Double, scaleY As Double,
                                                mediaSaveDir As String)
            Dim masterEntry As ZipArchiveEntry = zip.GetEntry("ppt/slideMasters/slideMaster1.xml")
            If masterEntry Is Nothing Then Return
            Dim xml As String
            Using r As New StreamReader(masterEntry.Open(), Encoding.UTF8)
                xml = r.ReadToEnd()
            End Using
            Dim mdoc As XDocument = XDocument.Parse(xml)
            Dim pns As XNamespace = mdoc.Root.Name.Namespace

            ' 主题 clrScheme（bgRef 映射用）
            Dim clrScheme As XElement = Nothing
            Dim themeNs As XNamespace = DRAWING_NS
            Dim themeEntry As ZipArchiveEntry = zip.GetEntry("ppt/theme/theme1.xml")
            If themeEntry IsNot Nothing Then
                Try
                    Dim txml As String
                    Using tr As New StreamReader(themeEntry.Open(), Encoding.UTF8)
                        txml = tr.ReadToEnd()
                    End Using
                    Dim tdoc As XDocument = XDocument.Parse(txml)
                    themeNs = tdoc.Root.Name.Namespace
                    clrScheme = tdoc.Descendants(themeNs + "clrScheme").FirstOrDefault()
                Catch
                End Try
            End If

            ' 1) 母版背景
            Dim bg As XElement = mdoc.Descendants(pns + "bg").FirstOrDefault()
            If bg IsNot Nothing Then
                Dim bgPr As XElement = bg.Element(pns + "bgPr")
                If bgPr IsNot Nothing Then
                    Dim solidFill As XElement = bgPr.Element(DRAWING_NS + "solidFill")
                    If solidFill IsNot Nothing Then
                        Dim c As Long = ExtractFillColor(solidFill, DRAWING_NS)
                        If c >= 0 Then profile.Background = New PptBackgroundProfile() With {.Kind = "color", .Color = c}
                    Else
                        Dim gradFill As XElement = bgPr.Element(DRAWING_NS + "gradFill")
                        If gradFill IsNot Nothing Then
                            Dim gsList As List(Of XElement) = gradFill.Descendants(DRAWING_NS + "gs").ToList()
                            Dim c1 As Long = If(gsList.Count > 0, ExtractFillColor(gsList(0), DRAWING_NS), -1)
                            Dim c2 As Long = If(gsList.Count > 1, ExtractFillColor(gsList(gsList.Count - 1), DRAWING_NS), -1)
                            If c1 >= 0 OrElse c2 >= 0 Then
                                Dim angle As Integer = 90
                                Dim lin As XElement = gradFill.Element(DRAWING_NS + "lin")
                                If lin IsNot Nothing Then
                                    Dim av As Integer = 0
                                    If Integer.TryParse(lin.Attribute("ang")?.Value, av) Then
                                        angle = ((av \ 60000) Mod 360 + 360) Mod 360
                                    End If
                                End If
                                profile.Background = New PptBackgroundProfile() With {
                                    .Kind = "gradient",
                                    .GradientColor1 = If(c1 >= 0, c1, If(c2 >= 0, c2, RGB(255, 255, 255))),
                                    .GradientColor2 = If(c2 >= 0, c2, If(c1 >= 0, c1, RGB(255, 255, 255))),
                                    .GradientAngle = angle
                                }
                            End If
                        Else
                            Dim blipFill As XElement = bgPr.Element(DRAWING_NS + "blipFill")
                            If blipFill IsNot Nothing Then
                                Dim blip As XElement = blipFill.Element(DRAWING_NS + "blip")
                                If blip IsNot Nothing Then
                                    Dim embed As String = blip.Attribute(OFFICE_REL_NS + "embed")?.Value
                                    If Not String.IsNullOrEmpty(embed) Then
                                        Dim mediaPath As String = ResolveMediaTarget(zip, MASTER_RELS_PATH, embed)
                                        If mediaPath IsNot Nothing Then
                                            Dim relPath As String = SaveMediaToConfig(zip, mediaPath, mediaSaveDir)
                                            If relPath IsNot Nothing Then
                                                profile.Background = New PptBackgroundProfile() With {.Kind = "picture", .ImagePath = relPath}
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                ' bgRef 主题色映射兜底（仅当 bgPr 未解析出背景时）
                If profile.Background Is Nothing Then
                    Dim bgRef As XElement = bg.Element(pns + "bgRef")
                    If bgRef IsNot Nothing AndAlso clrScheme IsNot Nothing Then
                        Dim c As Long = ResolveBgRefColor(clrScheme, themeNs, bgRef.Attribute("idx")?.Value)
                        If c >= 0 Then profile.Background = New PptBackgroundProfile() With {.Kind = "color", .Color = c}
                    End If
                End If
            End If

            ' 2) 母版背景图形（选择性提取）
            Dim spTree As XElement = mdoc.Descendants(pns + "spTree").FirstOrDefault()
            If spTree IsNot Nothing Then
                Dim masterShapes As New List(Of PptShapeProfile)()
                For Each child As XElement In spTree.Elements()
                    If masterShapes.Count >= 16 Then Exit For
                    Dim sp As XElement = Nothing
                    Dim isPic As Boolean = False
                    If child.Name = pns + "sp" Then
                        sp = child
                    ElseIf child.Name = pns + "pic" Then
                        sp = child
                        isPic = True
                    Else
                        Continue For   ' 跳过 grpSp/cxnSp 等复杂元素
                    End If
                    ' 跳过占位符
                    Dim nvPr As XElement = sp.Element(pns + "nvSpPr")?.Element(pns + "nvPr")
                    If nvPr IsNot Nothing AndAlso nvPr.Element(pns + "ph") IsNot Nothing Then Continue For
                    Dim spPr As XElement = sp.Element(pns + "spPr")
                    If spPr Is Nothing Then Continue For
                    Dim xfrm As XElement = spPr.Element(DRAWING_NS + "xfrm")
                    If xfrm Is Nothing Then Continue For
                    Dim off As XElement = xfrm.Element(DRAWING_NS + "off")
                    Dim ext As XElement = xfrm.Element(DRAWING_NS + "ext")
                    If off Is Nothing OrElse ext Is Nothing Then Continue For

                    Dim ps As New PptShapeProfile() With {
                        .Kind = "rect",
                        .Left = CInt(Math.Round(ParseEmu(off.Attribute("x")?.Value) * scaleX)),
                        .Top = CInt(Math.Round(ParseEmu(off.Attribute("y")?.Value) * scaleY)),
                        .Width = Math.Max(1, CInt(Math.Round(ParseEmu(ext.Attribute("cx")?.Value) * scaleX))),
                        .Height = Math.Max(1, CInt(Math.Round(ParseEmu(ext.Attribute("cy")?.Value) * scaleY)))
                    }
                    If isPic Then
                        ' 图片（logo 等）：blipFill 提取媒体
                        Dim blip As XElement = sp.Descendants(DRAWING_NS + "blip").FirstOrDefault()
                        If blip Is Nothing Then Continue For
                        Dim embed As String = blip.Attribute(OFFICE_REL_NS + "embed")?.Value
                        If String.IsNullOrEmpty(embed) Then Continue For
                        Dim mediaPath As String = ResolveMediaTarget(zip, MASTER_RELS_PATH, embed)
                        If mediaPath Is Nothing Then Continue For
                        ps.ImagePath = SaveMediaToConfig(zip, mediaPath, mediaSaveDir)
                        If String.IsNullOrEmpty(ps.ImagePath) Then Continue For
                    Else
                        ' 形状：类型 + 填充/线型；无视觉内容则跳过
                        Dim prstGeom As XElement = spPr.Element(DRAWING_NS + "prstGeom")
                        If prstGeom IsNot Nothing Then
                            ps.Kind = prstGeom.Attribute("prst")?.Value
                        End If
                        Dim solidFill As XElement = spPr.Descendants(DRAWING_NS + "solidFill").FirstOrDefault()
                        If solidFill IsNot Nothing Then
                            ps.Fill = ExtractFillColor(solidFill, DRAWING_NS)
                        Else
                            Dim ln As XElement = spPr.Element(DRAWING_NS + "ln")
                            If ln IsNot Nothing Then
                                Dim lnFill As XElement = ln.Element(DRAWING_NS + "solidFill")
                                If lnFill IsNot Nothing Then
                                    ps.LineColor = ExtractFillColor(lnFill, DRAWING_NS)
                                    ps.Fill = -2   ' 纯线条形状标记
                                End If
                            End If
                        End If
                        If ps.Fill < 0 AndAlso ps.LineColor < 0 Then Continue For
                    End If
                    masterShapes.Add(ps)
                Next
                If masterShapes.Count > 0 Then profile.MasterShapes = masterShapes
                ' 内容避让：计算顶部装饰区高度（顶部横幅/横贯形状的底部边界），内容区统一下移避免与装饰重叠。
                ' 判定：形状起点在页面上 30% 内，且横贯（宽≥1/2 页宽）或较高（高≥15% 页高）；上限 36% 页高防过度挤压。
                Dim inset As Integer = 0
                For Each ms As PptShapeProfile In masterShapes
                    If ms.Top <= 216 AndAlso (ms.Width >= 480 OrElse ms.Height >= 108) Then
                        Dim bottom As Integer = ms.Top + ms.Height
                        If bottom > inset Then inset = bottom
                    End If
                Next
                profile.ContentTopInset = Math.Min(inset, 260)
            End If

            ' 3) 母版占位符回退（版式未提取到标题/正文时）
            If profile.Title Is Nothing OrElse profile.Body Is Nothing Then
                For Each sp As XElement In mdoc.Descendants(pns + "sp")
                    Dim ph As XElement = sp.Descendants(pns + "ph").FirstOrDefault()
                    If ph Is Nothing Then Continue For
                    Dim phType As String = ph.Attribute("type")?.Value
                    If phType <> "title" AndAlso phType <> "body" Then Continue For
                    Dim xfrm As XElement = sp.Descendants(DRAWING_NS + "xfrm").FirstOrDefault()
                    If xfrm Is Nothing Then Continue For
                    Dim off As XElement = xfrm.Element(DRAWING_NS + "off")
                    Dim ext As XElement = xfrm.Element(DRAWING_NS + "ext")
                    If off Is Nothing OrElse ext Is Nothing Then Continue For
                    Dim rect As New PptRect() With {
                        .Left = CInt(Math.Round(ParseEmu(off.Attribute("x")?.Value) * scaleX)),
                        .Top = CInt(Math.Round(ParseEmu(off.Attribute("y")?.Value) * scaleY)),
                        .Width = CInt(Math.Round(ParseEmu(ext.Attribute("cx")?.Value) * scaleX)),
                        .Height = CInt(Math.Round(ParseEmu(ext.Attribute("cy")?.Value) * scaleY))
                    }
                    If phType = "title" AndAlso profile.Title Is Nothing Then profile.Title = rect
                    If phType = "body" AndAlso profile.Body Is Nothing Then profile.Body = rect
                Next
            End If
        End Sub
        Private Shared Function ExtractThemeFromPptx(pptxPath As String) As PptThemeJson
            Dim tj As New PptThemeJson()
            Using zip As ZipArchive = ZipFile.OpenRead(pptxPath)
                Dim entry As ZipArchiveEntry = zip.GetEntry("ppt/theme/theme1.xml")
                If entry Is Nothing Then
                    Throw New InvalidDataException("PPT 中未找到主题定义（ppt/theme/theme1.xml）")
                End If
                Dim xml As String
                Using reader As New StreamReader(entry.Open(), Encoding.UTF8)
                    xml = reader.ReadToEnd()
                End Using
                Dim doc As XDocument = XDocument.Parse(xml)
                Dim ns As XNamespace = doc.Root.Name.Namespace

                ' 配色方案：accent1 主色 / accent2 辅色 / accent3 强调 / lt1 背景 / lt2 浅色 / dk1 文字
                Dim clrScheme As XElement = doc.Descendants(ns + "clrScheme").FirstOrDefault()
                If clrScheme IsNot Nothing Then
                    Dim accent1 As Integer = ExtractColorFromScheme(clrScheme, ns, "accent1")
                    Dim accent2 As Integer = ExtractColorFromScheme(clrScheme, ns, "accent2")
                    Dim accent3 As Integer = ExtractColorFromScheme(clrScheme, ns, "accent3")
                    Dim accent6 As Integer = ExtractColorFromScheme(clrScheme, ns, "accent6")
                    Dim dk1 As Integer = ExtractColorFromScheme(clrScheme, ns, "dk1")
                    Dim dk2 As Integer = ExtractColorFromScheme(clrScheme, ns, "dk2")
                    Dim lt1 As Integer = ExtractColorFromScheme(clrScheme, ns, "lt1")
                    Dim lt2 As Integer = ExtractColorFromScheme(clrScheme, ns, "lt2")

                    Dim primary As Integer = If(accent1 >= 0, accent1, If(dk2 >= 0, dk2, 0))
                    Dim secondary As Integer = If(accent2 >= 0, accent2, primary)
                    Dim accent As Integer = If(accent3 >= 0, accent3, If(accent6 >= 0, accent6, primary))
                    Dim bg As Integer = If(lt1 >= 0, lt1, -1)
                    Dim light As Integer = If(lt2 >= 0, lt2, LightenColor(primary))
                    Dim text As Integer = If(dk1 >= 0, dk1, -1)

                    tj.Colors = New Dictionary(Of String, List(Of Integer))()
                    tj.Colors("primary") = RgbToList(primary)
                    tj.Colors("secondary") = RgbToList(secondary)
                    tj.Colors("accent") = RgbToList(accent)
                    tj.Colors("bg") = RgbToList(bg)
                    tj.Colors("light") = RgbToList(light)
                    tj.Colors("text") = RgbToList(text)
                    tj.Colors("gray") = New List(Of Integer)() From {150, 150, 150}
                End If

                ' 字体方案：优先 minorFont 东亚字体（中文模板），其次 latin，最后 majorFont
                Dim fontScheme As XElement = doc.Descendants(ns + "fontScheme").FirstOrDefault()
                If fontScheme IsNot Nothing Then
                    Dim minor As XElement = fontScheme.Element(ns + "minorFont")
                    Dim major As XElement = fontScheme.Element(ns + "majorFont")
                    Dim ea As String = ""
                    Dim latin As String = ""
                    If minor IsNot Nothing Then
                        ea = minor.Element(ns + "ea")?.Attribute("typeface")?.Value
                        latin = minor.Element(ns + "latin")?.Attribute("typeface")?.Value
                        If String.IsNullOrWhiteSpace(ea) Then
                            ' 中文模板常以 <a:font script="Hans"> 指定东亚字体
                            Dim hans As XElement = minor.Elements(ns + "font").FirstOrDefault(
                                Function(f) String.Equals(f.Attribute("script")?.Value, "Hans", StringComparison.OrdinalIgnoreCase))
                            If hans IsNot Nothing Then ea = hans.Attribute("typeface")?.Value
                        End If
                    End If
                    If String.IsNullOrWhiteSpace(ea) AndAlso major IsNot Nothing Then
                        ea = major.Element(ns + "ea")?.Attribute("typeface")?.Value
                    End If
                    If String.IsNullOrWhiteSpace(latin) AndAlso major IsNot Nothing Then
                        latin = major.Element(ns + "latin")?.Attribute("typeface")?.Value
                    End If
                    tj.Font = If(Not String.IsNullOrWhiteSpace(ea), ea,
                             If(Not String.IsNullOrWhiteSpace(latin), latin, "微软雅黑"))
                End If
            End Using
            Return tj
        End Function
        Private Shared Function ParseEmu(value As String) As Double
            Dim d As Double = 0
            Double.TryParse(value, d)
            Return d
        End Function
        Private Shared Function CapacityOf(layout As PptLayoutProfile) As PptCapacityProfile
            If layout Is Nothing Then Return Nothing
            If layout.Capacity IsNot Nothing Then Return layout.Capacity
            Dim cap As New PptCapacityProfile()
            If layout.Body IsNot Nothing Then
                cap.MaxBullets = Math.Max(3, CInt(Math.Floor(layout.Body.Height / 52)))
            End If
            If layout.Table IsNot Nothing Then
                cap.MaxTableRows = Math.Max(5, CInt(Math.Floor(layout.Table.Height / 14)))
                If layout.Table.Columns IsNot Nothing AndAlso layout.Table.Columns.Count > 0 Then
                    cap.MaxTableCols = layout.Table.Columns.Count
                End If
            End If
            Return cap
        End Function
        Private Shared Function ExtractColorFromScheme(clrScheme As XElement, ns As XNamespace, name As String) As Integer
            Dim el As XElement = clrScheme.Element(ns + name)
            If el Is Nothing Then Return -1
            Dim srgb As XElement = el.Element(ns + "srgbClr")
            If srgb IsNot Nothing Then
                Dim val As String = srgb.Attribute("val")?.Value
                If val IsNot Nothing AndAlso val.Length = 6 Then
                    Return Convert.ToInt32(val, 16)
                End If
            End If
            Dim sys As XElement = el.Element(ns + "sysClr")
            If sys IsNot Nothing Then
                Dim last As String = sys.Attribute("lastClr")?.Value
                If last IsNot Nothing AndAlso last.Length = 6 Then
                    Return Convert.ToInt32(last, 16)
                End If
            End If
            Return -1
        End Function
        Private Shared Function SanitizeThemeId(name As String) As String
            Dim sb As New StringBuilder()
            For Each ch As Char In name
                sb.Append(If(Char.IsLetterOrDigit(ch), ch, "_"c))
            Next
            Dim slug As String = sb.ToString().Trim("_"c)
            If String.IsNullOrWhiteSpace(slug) Then slug = "custom"
            Return "custom_" & slug
        End Function
        Private Shared Sub SaveCustomTheme(tj As PptThemeJson)
            Dim customPath As String = Path.Combine(PathHelper.ConfigDirectory, "custom_ppt_themes.json")
            Dim root As JObject = Nothing
            If File.Exists(customPath) Then
                Try
                    root = JObject.Parse(File.ReadAllText(customPath))
                Catch
                End Try
            End If
            If root Is Nothing Then
                root = New JObject()
                root("version") = 1
            End If
            If root("themes") Is Nothing Then root("themes") = New JArray()
            Dim themesArr As JArray = TryCast(root("themes"), JArray)

            Dim newTheme As New JObject()
            newTheme("id") = tj.Id
            newTheme("name") = tj.Name
            newTheme("font") = tj.Font
            Dim colors As New JObject()
            For Each kv As KeyValuePair(Of String, List(Of Integer)) In tj.Colors
                colors(kv.Key) = New JArray(kv.Value.Select(Function(v) JToken.FromObject(v)))
            Next
            newTheme("colors") = colors
            If tj.Layout IsNot Nothing Then
                ' camelCase 序列化（前端按 l.title/l.body/l.table 读取，与 PascalCase 不兼容）
                newTheme("layout") = JToken.FromObject(tj.Layout, JsonSerializer.Create(New JsonSerializerSettings With {
                    .ContractResolver = New CamelCasePropertyNamesContractResolver()
                }))
            End If

            Dim existing As JToken = themesArr.FirstOrDefault(
                Function(x) String.Equals(x.Value(Of String)("id"), tj.Id, StringComparison.OrdinalIgnoreCase))
            If existing IsNot Nothing Then
                existing.Replace(newTheme)
            Else
                themesArr.Add(newTheme)
            End If
            root("lastUpdated") = DateTime.Now.ToString("yyyy-MM-dd")
            File.WriteAllText(customPath, root.ToString(Formatting.Indented), Encoding.UTF8)
        End Sub
        Private Shared Function RgbToList(rgb As Integer) As List(Of Integer)
            If rgb < 0 Then Return New List(Of Integer)() From {245, 248, 252}
            Return New List(Of Integer)() From {(rgb >> 16) And &HFF, (rgb >> 8) And &HFF, rgb And &HFF}
        End Function
        Private Shared Function LightenColor(color As Integer) As Integer
            If color < 0 Then Return RGB(230, 240, 250)
            Dim r As Integer = (color >> 16) And &HFF
            Dim g As Integer = (color >> 8) And &HFF
            Dim b As Integer = color And &HFF
            r = CInt(r + (255 - r) * 0.85)
            g = CInt(g + (255 - g) * 0.85)
            b = CInt(b + (255 - b) * 0.85)
            Return (r << 16) Or (g << 8) Or b
        End Function
        Private Shared Function FindFirstLayoutEntry(zip As ZipArchive) As ZipArchiveEntry
            For Each entry As ZipArchiveEntry In zip.Entries
                If entry.FullName.StartsWith("ppt/slideLayouts/slideLayout", StringComparison.OrdinalIgnoreCase) AndAlso
                   entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) Then
                    Return entry
                End If
            Next
            Return Nothing
        End Function
        ''' <summary>
        ''' 使主题缓存失效（导入自定义主题后调用，下次生成/查询重新加载）。
        ''' </summary>
        Public Shared Sub InvalidateThemeCache()
            SyncLock _themeLoadLock
                _themeConfig = Nothing
                _themeLoadAttempted = False
            End SyncLock
        End Sub
        ''' <summary>
        ''' 解析主题：JSON 按 id 匹配 → 按显示名匹配 → 默认主题 → 内置兜底
        ''' （编辑器提示用户「用 [主题名] 生成」，须支持 name 与 id）
        ''' </summary>
        Private Shared Function ResolveTheme(themeId As String) As PptTheme
            Dim cfg As PptTemplatesConfig = LoadThemeConfig()
            If cfg IsNot Nothing Then
                Dim targetId As String = If(String.IsNullOrWhiteSpace(themeId), cfg.DefaultTheme, themeId)
                Dim tj As PptThemeJson = cfg.Themes.FirstOrDefault(
                    Function(t) String.Equals(t.Id, targetId, StringComparison.OrdinalIgnoreCase))
                If tj Is Nothing AndAlso Not String.IsNullOrWhiteSpace(themeId) Then
                    tj = cfg.Themes.FirstOrDefault(
                        Function(t) String.Equals(t.Name, themeId, StringComparison.OrdinalIgnoreCase))
                End If
                If tj Is Nothing AndAlso Not String.IsNullOrWhiteSpace(themeId) Then
                    ' 指定主题不存在 → 回退 JSON 默认主题
                    tj = cfg.Themes.FirstOrDefault(
                        Function(t) String.Equals(t.Id, cfg.DefaultTheme, StringComparison.OrdinalIgnoreCase))
                End If
                If tj IsNot Nothing Then
                    Return BuildTheme(tj)
                End If
            End If
            Return FALLBACK_THEME
        End Function
        ''' <summary>
        ''' 返回合并后的模板配置 JSON（内置 ppt_templates.json 完整结构 + 用户导入的 custom_ppt_themes.json 主题）。
        ''' 供前端 GetPptTemplates 与后端 LoadThemeConfig 共用，保证单一数据源一致。
        ''' </summary>
        Private Shared Function LoadThemeConfig() As PptTemplatesConfig
            If _themeLoadAttempted Then Return _themeConfig
            SyncLock _themeLoadLock
                If _themeLoadAttempted Then Return _themeConfig
                _themeLoadAttempted = True
                Try
                    Dim mergedJson As String = GetMergedTemplatesJson()
                    Dim cfg As PptTemplatesConfig = JsonConvert.DeserializeObject(Of PptTemplatesConfig)(mergedJson)
                    If cfg Is Nothing OrElse cfg.Themes Is Nothing OrElse cfg.Themes.Count = 0 Then
                        ChatLogger.Instance.LogWarning("PPT_GEN", "主题配置无有效主题，使用内置主题")
                        _themeConfig = Nothing
                        Return Nothing
                    End If
                    _themeConfig = cfg
                    ChatLogger.Instance.LogInfo("PPT_GEN", "已加载 " & cfg.Themes.Count.ToString() & " 个主题（含自定义）")
                Catch ex As Exception
                    ChatLogger.Instance.LogError("PPT_GEN", "主题配置加载失败: " & ex.Message)
                    _themeConfig = Nothing
                End Try
                Return _themeConfig
            End SyncLock
        End Function
        Public Shared Function GetMergedTemplatesJson() As String
            Dim root As JObject = Nothing
            Dim builtinPath As String = Path.Combine(PathHelper.ConfigDirectory, "ppt_templates.json")
            If File.Exists(builtinPath) Then
                Try
                    root = JObject.Parse(File.ReadAllText(builtinPath))
                Catch
                End Try
            End If
            If root Is Nothing Then root = New JObject()
            If root("themes") Is Nothing Then root("themes") = New JArray()

            Dim customPath As String = Path.Combine(PathHelper.ConfigDirectory, "custom_ppt_themes.json")
            If File.Exists(customPath) Then
                Try
                    Dim customRoot As JObject = JObject.Parse(File.ReadAllText(customPath))
                    Dim customThemes As JArray = TryCast(customRoot("themes"), JArray)
                    If customThemes IsNot Nothing Then
                        Dim themesArr As JArray = TryCast(root("themes"), JArray)
                        For Each t As JToken In customThemes
                            Dim id As String = t.Value(Of String)("id")
                            If String.IsNullOrWhiteSpace(id) Then Continue For
                            Dim existing As JToken = themesArr.FirstOrDefault(
                                Function(x) String.Equals(x.Value(Of String)("id"), id, StringComparison.OrdinalIgnoreCase))
                            If existing IsNot Nothing Then
                                existing.Replace(t)
                            Else
                                themesArr.Add(t)
                            End If
                        Next
                    End If
                Catch
                End Try
            End If
            Return root.ToString(Formatting.None)
        End Function

        ' ── 用户自定义主题管理（写 ConfigDirectory\custom_ppt_themes.json，与 GetMergedTemplatesJson 合并路径一致）──

        ''' <summary>
        ''' 保存用户自定义 PPT 主题。同 id 覆盖；成功后 InvalidateThemeCache 使 ppt_generate 立即生效。
        ''' themeJson: {"id":"my_theme","name":"我的主题","font":"微软雅黑","colors":{"primary":[r,g,b],...}}
        ''' </summary>
        Public Shared Function SaveCustomPptTheme(themeJson As String) As String
            Try
                If String.IsNullOrWhiteSpace(themeJson) Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "主题内容为空"})
                End If
                Dim t As JObject = JObject.Parse(themeJson)
                Dim themeId As String = If(t.Value(Of String)("id"), "").Trim()
                If String.IsNullOrWhiteSpace(themeId) Then
                    themeId = "custom_" & Guid.NewGuid().ToString("N").Substring(0, 8)
                End If
                t("id") = themeId
                If String.IsNullOrWhiteSpace(t.Value(Of String)("name")) Then t("name") = themeId
                If String.IsNullOrWhiteSpace(t.Value(Of String)("font")) Then t("font") = "微软雅黑"
                If Not (TypeOf t("colors") Is JObject) Then t("colors") = New JObject()
                Dim customPath As String = System.IO.Path.Combine(PathHelper.ConfigDirectory, "custom_ppt_themes.json")
                Dim root As JObject = New JObject()
                If System.IO.File.Exists(customPath) Then
                    Try
                        root = JObject.Parse(System.IO.File.ReadAllText(customPath))
                    Catch
                        root = New JObject()
                    End Try
                End If
                Dim arr As JArray = TryCast(root("themes"), JArray)
                If arr Is Nothing Then
                    arr = New JArray()
                    root("themes") = arr
                End If
                Dim existing As JToken = arr.FirstOrDefault(
                    Function(x) String.Equals(x.Value(Of String)("id"), themeId, StringComparison.OrdinalIgnoreCase))
                If existing IsNot Nothing Then
                    existing.Replace(t)
                Else
                    arr.Add(t)
                End If
                System.IO.File.WriteAllText(customPath, root.ToString(Formatting.Indented))
                InvalidateThemeCache()
                ChatLogger.Instance.LogInfo("PPT_GEN", "已保存自定义主题: " & themeId)
                Return JsonConvert.SerializeObject(New With {Key .success = True, Key .id = themeId, Key .message = "主题已保存并生效"})
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "保存自定义主题失败: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 删除用户自定义 PPT 主题（按 id），成功后清主题缓存
        ''' </summary>
        Public Shared Function DeleteCustomPptTheme(themeId As String) As String
            Try
                If String.IsNullOrWhiteSpace(themeId) Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "主题 id 为空"})
                End If
                Dim customPath As String = System.IO.Path.Combine(PathHelper.ConfigDirectory, "custom_ppt_themes.json")
                If Not System.IO.File.Exists(customPath) Then
                    Return JsonConvert.SerializeObject(New With {Key .success = True, Key .message = "无自定义主题"})
                End If
                Dim root As JObject = JObject.Parse(System.IO.File.ReadAllText(customPath))
                Dim arr As JArray = TryCast(root("themes"), JArray)
                If arr Is Nothing OrElse arr.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {Key .success = True, Key .message = "无自定义主题"})
                End If
                Dim kept As New JArray()
                For Each x As JToken In arr
                    If Not String.Equals(x.Value(Of String)("id"), themeId, StringComparison.OrdinalIgnoreCase) Then
                        kept.Add(x)
                    End If
                Next
                root("themes") = kept
                System.IO.File.WriteAllText(customPath, root.ToString(Formatting.Indented))
                InvalidateThemeCache()
                ChatLogger.Instance.LogInfo("PPT_GEN", "已删除自定义主题: " & themeId)
                Return JsonConvert.SerializeObject(New With {Key .success = True, Key .message = "已删除主题：" & themeId})
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "删除自定义主题失败: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message})
            End Try
        End Function

        ''' <summary>
        ''' 列出用户自定义 PPT 主题（custom_ppt_themes.json 的 themes 数组）
        ''' </summary>
        Public Shared Function ListCustomPptThemes() As String
            Try
                Dim customPath As String = System.IO.Path.Combine(PathHelper.ConfigDirectory, "custom_ppt_themes.json")
                Dim arr As New JArray()
                If System.IO.File.Exists(customPath) Then
                    Try
                        Dim root As JObject = JObject.Parse(System.IO.File.ReadAllText(customPath))
                        Dim t As JArray = TryCast(root("themes"), JArray)
                        If t IsNot Nothing Then arr = t
                    Catch
                    End Try
                End If
                Return JsonConvert.SerializeObject(New With {Key .success = True, Key .themes = arr})
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "读取自定义主题失败: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = ex.Message, Key .themes = New JArray()})
            End Try
        End Function

        Private Shared Function BuildTheme(tj As PptThemeJson) As PptTheme
            Dim theme As New PptTheme()
            theme.Id = If(String.IsNullOrEmpty(tj.Id), "business_blue", tj.Id)
            theme.Name = If(String.IsNullOrEmpty(tj.Name), theme.Id, tj.Name)
            theme.Font = If(String.IsNullOrEmpty(tj.Font), "微软雅黑", tj.Font)
            theme.Primary = ColorFromJson(tj.Colors, "primary", FALLBACK_THEME.Primary)
            theme.Secondary = ColorFromJson(tj.Colors, "secondary", FALLBACK_THEME.Secondary)
            theme.Accent = ColorFromJson(tj.Colors, "accent", theme.Secondary)
            theme.Bg = ColorFromJson(tj.Colors, "bg", FALLBACK_THEME.Bg)
            theme.Light = ColorFromJson(tj.Colors, "light", FALLBACK_THEME.Light)
            theme.Text = ColorFromJson(tj.Colors, "text", FALLBACK_THEME.Text)
            theme.Gray = ColorFromJson(tj.Colors, "gray", FALLBACK_THEME.Gray)
            theme.LayoutProfile = tj.Layout
            theme.DesignTokens = BuildDesignTokens(tj.DesignTokens)
            Return theme
        End Function

        ''' <summary>
        ''' 从主题 JSON 的 designTokens 构建运行时令牌对象。
        ''' - density 预设（compact/comfortable/spacious）展开为显式值；显式字段覆盖预设。
        ''' - 非法值回退对应默认（校验器已在创建路径做过硬校验，此处为手工 JSON 的兜底容错）。
        ''' - Nothing / 空对象 → Nothing（渲染走内置默认，行为与改造前一致）。
        ''' </summary>
        Private Shared Function BuildDesignTokens(tokenJson As JToken) As PptDesignTokens
            If tokenJson Is Nothing Then Return Nothing
            Dim obj As JObject = TryCast(tokenJson, JObject)
            If obj Is Nothing OrElse obj.Count = 0 Then Return Nothing
            Dim t As New PptDesignTokens()
            Dim density As String = If(obj.Value(Of String)("density"), "").Trim().ToLowerInvariant()
            Select Case density
                Case "compact"
                    t.TitleSize = 22 : t.HeadingSize = 18 : t.BodySize = 14 : t.SmallSize = 11 : t.FooterSize = 9
                    t.Margin = 28 : t.Grid = 16 : t.TitleTop = 16 : t.ContentTop = 84 : t.Dense = True
                Case "spacious"
                    t.TitleSize = 30 : t.HeadingSize = 24 : t.BodySize = 18 : t.SmallSize = 13 : t.FooterSize = 11
                    t.Margin = 48 : t.Grid = 24 : t.TitleTop = 24 : t.ContentTop = 116 : t.Dense = False
                Case Else
                    ' comfortable = 构造默认（与内置常量一致）
            End Select
            t.TitleSize = TokenInt(obj, "titleSize", t.TitleSize)
            t.HeadingSize = TokenInt(obj, "headingSize", t.HeadingSize)
            t.BodySize = TokenInt(obj, "bodySize", t.BodySize)
            t.SmallSize = TokenInt(obj, "smallSize", t.SmallSize)
            t.FooterSize = TokenInt(obj, "footerSize", t.FooterSize)
            t.Margin = TokenInt(obj, "margin", t.Margin)
            t.Grid = TokenInt(obj, "grid", t.Grid)
            t.TitleTop = TokenInt(obj, "titleTop", t.TitleTop)
            t.ContentTop = TokenInt(obj, "contentTop", t.ContentTop)
            t.RibbonWidth = TokenInt(obj, "ribbonWidth", t.RibbonWidth)
            Dim ribbon As String = If(obj.Value(Of String)("ribbonStyle"), "").Trim().ToLowerInvariant()
            If ribbon = "line" OrElse ribbon = "bar" OrElse ribbon = "none" Then t.RibbonStyle = ribbon

            ' pageVariants：主题级页型变体默认（spec 页级 variant 优先覆盖）。
            ' 只收合法键/值（PageVariantAllowed 与校验器共用单一事实源），非法忽略。
            Dim pv = obj("pageVariants")
            If pv IsNot Nothing Then
                Dim pvObj As JObject = TryCast(pv, JObject)
                If pvObj IsNot Nothing Then
                    Dim variants As New Dictionary(Of String, String)()
                    For Each prop As JProperty In pvObj.Properties()
                        Dim allowed As String() = PptThemeDesignValidator.PageVariantAllowed(prop.Name)
                        If allowed Is Nothing Then Continue For
                        Dim val As String = Convert.ToString(prop.Value).Trim().ToLowerInvariant()
                        If Array.IndexOf(allowed, val) >= 0 Then
                            variants(prop.Name.ToLowerInvariant()) = val
                        End If
                    Next
                    If variants.Count > 0 Then t.PageVariants = variants
                End If
            End If
            Return t
        End Function

        Private Shared Function TokenInt(obj As JObject, key As String, defaultValue As Integer) As Integer
            Dim raw = obj(key)
            If raw Is Nothing Then Return defaultValue
            Dim n As Integer
            If Integer.TryParse(Convert.ToString(raw), n) AndAlso n >= 4 AndAlso n <= 400 Then Return n
            Return defaultValue
        End Function
        ''' <summary>
        ''' 打开原生文件选择对话框，选择公司模板 .pptx。
        ''' 成功返回 {"success":true,"filePath":"..."}；用户取消返回 {"success":false,"message":"用户取消选择"}。
        ''' </summary>
        Public Shared Function BrowsePptFile() As String
            Try
                Using dlg As New System.Windows.Forms.OpenFileDialog()
                    dlg.Title = "选择公司模板 PPT"
                    dlg.Filter = "PowerPoint 演示文稿 (*.pptx)|*.pptx|所有文件 (*.*)|*.*"
                    dlg.CheckFileExists = True
                    dlg.Multiselect = False
                    If dlg.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                        Return JsonConvert.SerializeObject(New With {
                            Key .success = True,
                            Key .filePath = dlg.FileName
                        })
                    Else
                        Return JsonConvert.SerializeObject(New With {
                            Key .success = False,
                            Key .message = "用户取消选择"
                        })
                    End If
                End Using
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "打开文件对话框失败: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {
                    Key .success = False,
                    Key .error = "打开文件对话框失败：" & ex.Message
                })
            End Try
        End Function
        ''' <summary>
        ''' 从已有 PPT 提取版式画像（标题/正文占位符位置、表格位置与列宽、装饰形状位置、母版背景与背景图形）。
        ''' 坐标从 EMU 换算到设计空间（960x720）。提取失败时返回 Nothing（不影响配色/字体导入）。
        ''' </summary>
        Private Shared Function ExtractLayoutProfile(pptxPath As String, Optional mediaSaveDir As String = "") As PptLayoutProfile
            Try
                Dim profile As New PptLayoutProfile()
                Using zip As ZipArchive = ZipFile.OpenRead(pptxPath)
                    ' 幻灯片尺寸（EMU），默认 4:3
                    Dim slideW As Double = 9144000
                    Dim slideH As Double = 6858000
                    Dim presEntry As ZipArchiveEntry = zip.GetEntry("ppt/presentation.xml")
                    If presEntry IsNot Nothing Then
                        Dim presXml As String
                        Using r As New StreamReader(presEntry.Open(), Encoding.UTF8)
                            presXml = r.ReadToEnd()
                        End Using
                        Dim pdoc As XDocument = XDocument.Parse(presXml)
                        Dim pns As XNamespace = pdoc.Root.Name.Namespace
                        Dim sldSz As XElement = pdoc.Descendants(pns + "sldSz").FirstOrDefault()
                        If sldSz IsNot Nothing Then
                            Dim cx As Double = 0
                            Dim cy As Double = 0
                            Double.TryParse(sldSz.Attribute("cx")?.Value, cx)
                            Double.TryParse(sldSz.Attribute("cy")?.Value, cy)
                            If cx > 0 AndAlso cy > 0 Then
                                slideW = cx
                                slideH = cy
                            End If
                        End If
                    End If
                    Dim scaleX As Double = DESIGN_W / slideW
                    Dim scaleY As Double = DESIGN_H / slideH

                    ' 标题/正文占位符：从 slideLayout 提取
                    Dim layoutEntry As ZipArchiveEntry = FindFirstLayoutEntry(zip)
                    If layoutEntry IsNot Nothing Then
                        Dim lxml As String
                        Using r As New StreamReader(layoutEntry.Open(), Encoding.UTF8)
                            lxml = r.ReadToEnd()
                        End Using
                        Dim ldoc As XDocument = XDocument.Parse(lxml)
                        Dim lpns As XNamespace = ldoc.Root.Name.Namespace
                        Dim lans As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"
                        For Each sp As XElement In ldoc.Descendants(lpns + "sp")
                            Dim ph As XElement = sp.Descendants(lpns + "ph").FirstOrDefault()
                            If ph Is Nothing Then Continue For
                            Dim phType As String = ph.Attribute("type")?.Value
                            Dim xfrm As XElement = sp.Descendants(lans + "xfrm").FirstOrDefault()
                            If xfrm Is Nothing Then Continue For
                            Dim off As XElement = xfrm.Element(lans + "off")
                            Dim ext As XElement = xfrm.Element(lans + "ext")
                            If off Is Nothing OrElse ext Is Nothing Then Continue For
                            Dim rect As New PptRect() With {
                                .Left = CInt(Math.Round(ParseEmu(off.Attribute("x")?.Value) * scaleX)),
                                .Top = CInt(Math.Round(ParseEmu(off.Attribute("y")?.Value) * scaleY)),
                                .Width = CInt(Math.Round(ParseEmu(ext.Attribute("cx")?.Value) * scaleX)),
                                .Height = CInt(Math.Round(ParseEmu(ext.Attribute("cy")?.Value) * scaleY))
                            }
                            If phType = "title" AndAlso profile.Title Is Nothing Then
                                profile.Title = rect
                            ElseIf phType = "body" AndAlso profile.Body Is Nothing Then
                                profile.Body = rect
                            End If
                        Next
                    End If

                    ' 表格 + 形状：从第一张幻灯片提取
                    Dim slideEntry As ZipArchiveEntry = zip.GetEntry("ppt/slides/slide1.xml")
                    If slideEntry IsNot Nothing Then
                        Dim sxml As String
                        Using r As New StreamReader(slideEntry.Open(), Encoding.UTF8)
                            sxml = r.ReadToEnd()
                        End Using
                        Dim sdoc As XDocument = XDocument.Parse(sxml)
                        Dim spns As XNamespace = sdoc.Root.Name.Namespace
                        Dim sans As XNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main"

                        ' 表格：位置 + 列宽 + 表头填充色
                        For Each gf As XElement In sdoc.Descendants(spns + "graphicFrame")
                            Dim tbl As XElement = gf.Descendants(sans + "tbl").FirstOrDefault()
                            If tbl Is Nothing Then Continue For
                            Dim tp As New PptTableProfile()
                            ' graphicFrame 的 xfrm 位于 presentationml 命名空间（<p:xfrm>），off/ext 为 drawingml
                            Dim xfrm As XElement = gf.Descendants(spns + "xfrm").FirstOrDefault()
                            If xfrm IsNot Nothing Then
                                Dim off As XElement = xfrm.Element(sans + "off")
                                Dim ext As XElement = xfrm.Element(sans + "ext")
                                If off IsNot Nothing Then
                                    tp.Left = CInt(Math.Round(ParseEmu(off.Attribute("x")?.Value) * scaleX))
                                    tp.Top = CInt(Math.Round(ParseEmu(off.Attribute("y")?.Value) * scaleY))
                                End If
                                If ext IsNot Nothing Then
                                    tp.Width = CInt(Math.Round(ParseEmu(ext.Attribute("cx")?.Value) * scaleX))
                                    tp.Height = CInt(Math.Round(ParseEmu(ext.Attribute("cy")?.Value) * scaleY))
                                End If
                            End If
                            Dim tblGrid As XElement = tbl.Element(sans + "tblGrid")
                            If tblGrid IsNot Nothing Then
                                tp.Columns = New List(Of Integer)()
                                For Each gc As XElement In tblGrid.Elements(sans + "gridCol")
                                    tp.Columns.Add(CInt(Math.Round(ParseEmu(gc.Attribute("w")?.Value) * scaleX)))
                                Next
                            End If
                            Dim firstRow As XElement = tbl.Descendants(sans + "tr").FirstOrDefault()
                            If firstRow IsNot Nothing Then
                                Dim tc As XElement = firstRow.Descendants(sans + "tc").FirstOrDefault()
                                If tc IsNot Nothing Then
                                    Dim fill As XElement = tc.Descendants(sans + "srgbClr").FirstOrDefault()
                                    If fill IsNot Nothing Then
                                        Dim v As String = fill.Attribute("val")?.Value
                                        If v IsNot Nothing AndAlso v.Length = 6 Then
                                            tp.HeaderFill = Convert.ToInt32(v, 16)
                                        End If
                                    End If
                                End If
                            End If
                            profile.Table = tp
                            Exit For
                        Next

                        ' 装饰形状：位置 + 类型 + 填充色
                        ' 仅提取无文字、非占位符的形状（文本框/标题占位符不作为装饰形状套用）
                        profile.Shapes = New List(Of PptShapeProfile)()
                        For Each sp As XElement In sdoc.Descendants(spns + "sp")
                            Dim nvPr As XElement = sp.Element(spns + "nvSpPr")?.Element(spns + "nvPr")
                            If nvPr IsNot Nothing AndAlso nvPr.Element(spns + "ph") IsNot Nothing Then Continue For
                            Dim txBody As XElement = sp.Element(spns + "txBody")
                            If txBody IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(txBody.Value) Then Continue For
                            Dim spPr As XElement = sp.Element(spns + "spPr")
                            If spPr Is Nothing Then Continue For
                            Dim xfrm As XElement = spPr.Element(sans + "xfrm")
                            If xfrm Is Nothing Then Continue For
                            Dim off As XElement = xfrm.Element(sans + "off")
                            Dim ext As XElement = xfrm.Element(sans + "ext")
                            If off Is Nothing OrElse ext Is Nothing Then Continue For
                            Dim kind As String = "rect"
                            Dim prstGeom As XElement = spPr.Element(sans + "prstGeom")
                            If prstGeom IsNot Nothing Then
                                kind = prstGeom.Attribute("prst")?.Value
                            End If
                            Dim fill As Long = -1
                            Dim solidFill As XElement = spPr.Descendants(sans + "solidFill").FirstOrDefault()
                            If solidFill IsNot Nothing Then
                                Dim srgb As XElement = solidFill.Descendants(sans + "srgbClr").FirstOrDefault()
                                If srgb IsNot Nothing Then
                                    Dim v As String = srgb.Attribute("val")?.Value
                                    If v IsNot Nothing AndAlso v.Length = 6 Then fill = Convert.ToInt32(v, 16)
                                End If
                            End If
                            profile.Shapes.Add(New PptShapeProfile() With {
                                .Kind = kind,
                                .Left = CInt(Math.Round(ParseEmu(off.Attribute("x")?.Value) * scaleX)),
                                .Top = CInt(Math.Round(ParseEmu(off.Attribute("y")?.Value) * scaleY)),
                                .Width = CInt(Math.Round(ParseEmu(ext.Attribute("cx")?.Value) * scaleX)),
                                .Height = CInt(Math.Round(ParseEmu(ext.Attribute("cy")?.Value) * scaleY)),
                                .Fill = fill
                            })
                        Next
                        If profile.Shapes.Count = 0 Then profile.Shapes = Nothing
                    End If

                    ' 母版结构：背景 + 背景图形 + 占位符回退（选择性提取）
                    ExtractMasterProfile(zip, profile, scaleX, scaleY, mediaSaveDir)
                End Using

                ' 容量限制：由内容区尺寸自动推导（正文区高度→要点数，表格区高度→行数）
                profile.Capacity = New PptCapacityProfile()
                If profile.Body IsNot Nothing Then
                    profile.Capacity.MaxBullets = Math.Max(3, CInt(Math.Floor(profile.Body.Height / 52)))
                End If
                If profile.Table IsNot Nothing Then
                    profile.Capacity.MaxTableRows = Math.Max(5, CInt(Math.Floor(profile.Table.Height / 14)))
                    If profile.Table.Columns IsNot Nothing AndAlso profile.Table.Columns.Count > 0 Then
                        profile.Capacity.MaxTableCols = profile.Table.Columns.Count
                    End If
                End If

                ' 至少提取到一项才算成功
                If profile.Title Is Nothing AndAlso profile.Body Is Nothing AndAlso
                   profile.Table Is Nothing AndAlso profile.Shapes Is Nothing Then
                    Return Nothing
                End If
                Return profile
            Catch ex As Exception
                ChatLogger.Instance.LogWarning("PPT_GEN", "提取版式画像失败（忽略，仅配色/字体生效）: " & ex.Message)
                Return Nothing
            End Try
        End Function
        Private Shared Function ColorFromJson(colors As Dictionary(Of String, List(Of Integer)),
                                              key As String,
                                              fallback As Long) As Long
            If colors IsNot Nothing AndAlso colors.ContainsKey(key) Then
                Dim rgbValues As List(Of Integer) = colors(key)
                If rgbValues IsNot Nothing AndAlso rgbValues.Count >= 3 Then
                    Return RGB(rgbValues(0), rgbValues(1), rgbValues(2))
                End If
            End If
            Return fallback
        End Function
        ''' <summary>
        ''' 从已有 PPT 提取主题（配色 + 字体）并注册为自定义主题。
        ''' 成功返回 {"success":true,"themeId","themeName","font","colors"}；失败返回 {"success":false,"error"}。
        ''' </summary>
        Public Shared Function ImportThemeFromPptx(pptxPath As String) As String
            Try
                If String.IsNullOrWhiteSpace(pptxPath) Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "未提供 PPT 文件路径"})
                End If
                If Not File.Exists(pptxPath) Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "文件不存在：" & pptxPath})
                End If
                Dim ext As String = Path.GetExtension(pptxPath).ToLowerInvariant()
                If ext <> ".pptx" Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "仅支持 .pptx 文件（当前：" & ext & "）"})
                End If

                Dim tj As PptThemeJson = ExtractThemeFromPptx(pptxPath)
                Dim fileName As String = Path.GetFileNameWithoutExtension(pptxPath)
                tj.Id = SanitizeThemeId(fileName)
                tj.Name = fileName
                If String.IsNullOrWhiteSpace(tj.Font) Then tj.Font = "微软雅黑"
                If tj.Colors Is Nothing OrElse tj.Colors.Count = 0 Then
                    Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "未能从 PPT 提取到配色方案"})
                End If

                ' 版式画像：标题/正文/表格/形状位置 + 母版背景与背景图形（失败不影响配色字体导入）
                ' 母版媒体图片复制到 Config/ppt_media/{themeId}/，主题 JSON 存相对路径
                Dim mediaDir As String = Path.Combine(PathHelper.ConfigDirectory, "ppt_media", tj.Id)
                tj.Layout = ExtractLayoutProfile(pptxPath, mediaDir)

                SaveCustomTheme(tj)
                InvalidateThemeCache()

                ChatLogger.Instance.LogInfo("PPT_GEN", "已导入模板主题: " & tj.Name & " (" & tj.Id & ")" &
                    If(tj.Layout Is Nothing, "", "，含版式画像"))
                Return JsonConvert.SerializeObject(New With {
                    Key .success = True,
                    Key .themeId = tj.Id,
                    Key .themeName = tj.Name,
                    Key .font = tj.Font,
                    Key .colors = tj.Colors,
                    Key .layout = LayoutSummary(tj.Layout)
                })
            Catch ex As Exception
                ChatLogger.Instance.LogError("PPT_GEN", "导入 PPT 主题失败: " & ex.Message)
                Return JsonConvert.SerializeObject(New With {Key .success = False, Key .error = "导入失败：" & ex.Message})
            End Try
        End Function
End Class
End Namespace
