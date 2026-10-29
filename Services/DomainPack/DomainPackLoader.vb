Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Services.DomainPack

    ''' <summary>
    ''' 专项识别模块清单解析器。
    '''
    ''' 清单格式 = Markdown + YAML frontmatter + 一个 ```options 代码块（JSON 数组）：
    '''
    ''' ---
    ''' id: architectural_drawing
    ''' name: 建筑图
    ''' icon: 建筑
    ''' category: 工程图
    ''' version: 0.1.0
    ''' pipeline: python
    ''' module: arch_drawing.py
    ''' requires: opencv-python, numpy
    ''' ---
    '''
    ''' ```options
    ''' [{"key":"detect_grid","label":"识别轴网","type":"bool","default":true}]
    ''' ```
    '''
    ''' 约定：解析**永不抛异常**。坏清单返回 Nothing 或带 ErrorText 的对象，
    ''' 由注册表跳过并在界面上标红 —— 一个坏文件不能拖垮整台 OCR。
    ''' </summary>
    Public Class DomainPackLoader

        ''' <summary>从文本解析清单；不是合法清单返回 Nothing。</summary>
        Public Shared Function ParseMarkdown(content As String, Optional filePath As String = "") As DomainPackManifest
            If String.IsNullOrWhiteSpace(content) Then Return Nothing

            Dim lines = content.Replace(vbCrLf, vbLf).Split(vbLf)
            If lines.Length = 0 OrElse lines(0).Trim() <> "---" Then Return Nothing

            Dim m As New DomainPackManifest()
            m.FilePath = filePath

            Dim i = 1
            While i < lines.Length
                Dim line = lines(i)
                If line.Trim() = "---" Then Exit While

                If Not String.IsNullOrWhiteSpace(line) AndAlso Not line.TrimStart().StartsWith("#") Then
                    Dim idx = line.IndexOf(":"c)
                    If idx > 0 Then
                        Dim key = line.Substring(0, idx).Trim().ToLowerInvariant()
                        Dim value = line.Substring(idx + 1).Trim().Trim(""""c).Trim("'"c)
                        ApplyField(m, key, value)
                    End If
                End If
                i += 1
            End While

            ' options 代码块（JSON 数组）
            Dim optionsBlock = ExtractFencedBlock(content, "options")
            If Not String.IsNullOrWhiteSpace(optionsBlock) Then ParseOptions(m, optionsBlock)

            Return If(Validate(m), m, Nothing)
        End Function

        ''' <summary>读取清单文件；失败返回 Nothing。</summary>
        Public Shared Function ParseFile(fullPath As String) As DomainPackManifest
            Try
                If Not File.Exists(fullPath) Then Return Nothing
                Dim m = ParseMarkdown(File.ReadAllText(fullPath), fullPath)
                If m Is Nothing Then
                    Dim bad As New DomainPackManifest()
                    bad.Id = System.IO.Path.GetFileNameWithoutExtension(fullPath)
                    bad.Name = bad.Id
                    bad.FilePath = fullPath
                    bad.ErrorText = "清单缺少 frontmatter 或必填字段（id / name / pipeline）"
                    Return bad
                End If
                If String.IsNullOrWhiteSpace(m.Id) Then m.Id = System.IO.Path.GetFileNameWithoutExtension(fullPath)
                Return m
            Catch ex As Exception
                Return New DomainPackManifest() With {
                    .Id = System.IO.Path.GetFileNameWithoutExtension(fullPath),
                    .Name = System.IO.Path.GetFileNameWithoutExtension(fullPath),
                    .FilePath = fullPath,
                    .ErrorText = "读取失败：" & ex.Message
                }
            End Try
        End Function

        Private Shared Sub ApplyField(m As DomainPackManifest, key As String, value As String)
            Select Case key
                Case "id"
                    m.Id = NormalizeId(value)
                Case "name"
                    m.Name = value
                Case "description"
                    m.Description = value
                Case "icon"
                    m.Icon = value
                Case "category"
                    m.Category = value
                Case "version"
                    m.Version = value
                Case "author"
                    m.Author = value
                Case "pipeline"
                    Select Case value.ToLowerInvariant()
                        Case "python"
                            m.Pipeline = DomainPackPipeline.Python
                        Case "builtin", "built-in", "vb"
                            m.Pipeline = DomainPackPipeline.BuiltIn
                        Case Else
                            m.Pipeline = DomainPackPipeline.Text
                    End Select
                Case "module", "moduleref"
                    m.ModuleRef = value
                Case "requires", "deps"
                    m.Requires = SplitList(value)
                Case "parsescript", "parsescriptid"
                    m.ParseScriptId = value
            End Select
        End Sub

        Private Shared Sub ParseOptions(m As DomainPackManifest, json As String)
            Try
                Dim arr = JArray.Parse(json)
                For Each item In arr
                    If TypeOf item Is JObject Then
                        Dim o = DirectCast(item, JObject)
                        Dim opt As New DomainPackOption()
                        opt.Key = Str(o, "key")
                        opt.Label = Str(o, "label")
                        opt.Type = Str(o, "type")
                        If String.IsNullOrWhiteSpace(opt.Type) Then opt.Type = "bool"
                        opt.DefaultValue = RawDefault(o)

                        Dim vals = o.Value(Of JArray)("values")
                        If vals IsNot Nothing Then
                            For Each v In vals
                                opt.Values.Add(v.ToString())
                            Next
                        End If

                        If Not String.IsNullOrWhiteSpace(opt.Key) Then
                            If String.IsNullOrWhiteSpace(opt.Label) Then opt.Label = opt.Key
                            m.Options.Add(opt)
                        End If
                    End If
                Next
            Catch ex As Exception
                m.ErrorText = "options 代码块不是合法 JSON：" & ex.Message
            End Try
        End Sub

        Private Shared Function RawDefault(o As JObject) As String
            Dim d = o.Value(Of JToken)("default")
            If d Is Nothing Then Return ""
            If d.Type = JTokenType.Boolean Then Return If(d.Value(Of Boolean)(), "true", "false")
            Return d.ToString()
        End Function

        Private Shared Function Str(o As JObject, key As String) As String
            Dim v = o.Value(Of String)(key)
            Return If(v, "")
        End Function

        ''' <summary>取 ```lang 代码块内容（首个匹配）。</summary>
        Public Shared Function ExtractFencedBlock(content As String, lang As String) As String
            Try
                Dim marker = "```" & lang
                Dim start = content.IndexOf(marker, StringComparison.Ordinal)
                If start < 0 Then Return ""
                Dim bodyStart = content.IndexOf(vbLf, start)
                If bodyStart < 0 Then Return ""
                Dim close = content.IndexOf("```", bodyStart + 1, StringComparison.Ordinal)
                If close < 0 Then Return content.Substring(bodyStart + 1)
                Return content.Substring(bodyStart + 1, close - bodyStart - 1).Trim()
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Private Shared Function SplitList(value As String) As List(Of String)
            Dim list As New List(Of String)()
            If String.IsNullOrWhiteSpace(value) Then Return list
            For Each part In value.Split(","c)
                Dim p = part.Trim()
                If p.Length > 0 AndAlso Not list.Contains(p) Then list.Add(p)
            Next
            Return list
        End Function

        Private Shared Function NormalizeId(value As String) As String
            If String.IsNullOrWhiteSpace(value) Then Return ""
            Dim sb As New StringBuilder()
            For Each ch In value.Trim().ToLowerInvariant()
                If Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c Then sb.Append(ch)
            Next
            Return sb.ToString()
        End Function

        Private Shared Function Validate(m As DomainPackManifest) As Boolean
            If String.IsNullOrWhiteSpace(m.Id) Then Return False
            If m.Pipeline = DomainPackPipeline.Python AndAlso String.IsNullOrWhiteSpace(m.ModuleRef) Then
                m.ErrorText = "pipeline=python 必须指定 module（脚本文件名）"
                Return True   ' 仍然返回对象，让界面把错误显示出来
            End If
            If m.Pipeline = DomainPackPipeline.BuiltIn AndAlso String.IsNullOrWhiteSpace(m.ModuleRef) Then
                m.ErrorText = "pipeline=builtin 必须指定 module（识别器代号）"
                Return True
            End If
            Return True
        End Function
    End Class
End Namespace
