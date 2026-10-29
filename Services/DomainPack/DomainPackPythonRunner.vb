Imports System
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq

Namespace Services.DomainPack

    ''' <summary>
    ''' 识别改写入口。文字引擎先出框，已勾选的 python 模块再按原图改写文字。
    ''' 机械制图仍走内置 GdtSymbolRecognizer，不经过这里。
    ''' 脚本超时、缺失或返回失败时，保留原来的识别结果。
    ''' </summary>
    Public Class DomainPackPythonRunner

        Public Const TimeoutMs As Integer = 15000

        Private Shared _pythonPath As String
        Private Shared ReadOnly _pyLock As New Object()

        ''' <summary>跑全部已启用、且清单本身没有错误的 python 模块。</summary>
        Public Shared Sub ApplyEnabled(doc As OcrDocument)
            If doc Is Nothing OrElse Not String.IsNullOrEmpty(doc.ErrorMessage) Then Return
            Dim packs As List(Of DomainPackManifest) = Nothing
            Try
                packs = DomainPackRegistry.Instance.All()
            Catch ex As Exception
                Debug.WriteLine("[DomainPack] 读取模块列表失败：" & ex.Message)
                Return
            End Try
            If packs Is Nothing Then Return
            For Each pack In packs
                If pack Is Nothing OrElse Not pack.Enabled Then Continue For
                If pack.Pipeline <> DomainPackPipeline.Python Then Continue For
                If Not String.IsNullOrEmpty(pack.ErrorText) Then Continue For
                Try
                    ApplyOne(pack, doc)
                Catch ex As Exception
                    Debug.WriteLine("[DomainPack] 模块 " & pack.Id & " 执行失败：" & ex.Message)
                End Try
            Next
        End Sub

        Public Shared Sub ApplyOne(pack As DomainPackManifest, doc As OcrDocument)
            If pack Is Nothing OrElse doc Is Nothing Then Return
            Dim root = DomainPackRegistry.Instance.RootDirectory
            Dim script = ResolveScriptPath(root, pack.ModuleRef, pack.Id)
            If String.IsNullOrEmpty(script) OrElse Not File.Exists(script) Then
                Debug.WriteLine("[DomainPack] 脚本不存在，跳过 " & pack.Id)
                Return
            End If
            Dim targets = TargetElements(doc)
            If targets Is Nothing OrElse targets.Count = 0 Then Return

            Dim python = ResolvePython()
            If String.IsNullOrEmpty(python) Then
                Debug.WriteLine("[DomainPack] 未找到 Python，跳过 " & pack.Id)
                Return
            End If

            Dim request = BuildRequest(pack, doc, targets)
            Dim output = RunScript(python, script, root, request)
            If String.IsNullOrWhiteSpace(output) Then Return
            ApplyResponse(doc, targets, output)
        End Sub

        ''' <summary>只接受模块目录里的 .py 文件名，拒绝绝对路径和 .. 。</summary>
        Public Shared Function SafeScriptFileName(moduleRef As String, id As String) As String
            Dim name = If(String.IsNullOrWhiteSpace(moduleRef), If(id, "") & ".py", Path.GetFileName(moduleRef.Trim()))
            If String.IsNullOrWhiteSpace(name) Then Return ""
            If name.IndexOfAny(New Char() {Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar, ":"c, """"c}) >= 0 Then Return ""
            If name.Contains("..") Then Return ""
            If Not name.EndsWith(".py", StringComparison.OrdinalIgnoreCase) Then Return ""
            For Each ch In name
                If Char.IsLetterOrDigit(ch) OrElse ch = "_"c OrElse ch = "-"c OrElse ch = "."c Then Continue For
                Return ""
            Next
            Return name
        End Function

        Public Shared Function ResolveScriptPath(root As String, moduleRef As String, id As String) As String
            Dim name = SafeScriptFileName(moduleRef, id)
            If String.IsNullOrEmpty(name) OrElse String.IsNullOrWhiteSpace(root) Then Return ""
            Try
                Dim rootFull = Path.GetFullPath(root)
                Dim full = Path.GetFullPath(Path.Combine(rootFull, name))
                Dim prefix = rootFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) & Path.DirectorySeparatorChar
                If Not full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then Return ""
                Return full
            Catch ex As Exception
                Return ""
            End Try
        End Function

        ''' <summary>把脚本输出写回文字框。失败或 success=false 时不改文档。</summary>
        Public Shared Sub ApplyResponse(doc As OcrDocument, targets As List(Of OcrElement), output As String)
            If doc Is Nothing OrElse targets Is Nothing Then Return
            Dim json = ParsePayload(output)
            If json Is Nothing Then Return
            Dim ok = json.Value(Of JToken)("success")
            If ok IsNot Nothing AndAlso ok.Type = JTokenType.Boolean AndAlso Not ok.Value(Of Boolean)() Then Return
            Dim arr = json.Value(Of JArray)("elements")
            If arr Is Nothing Then Return

            Dim ambiguous As Boolean = False
            For Each item In arr
                Dim obj = TryCast(item, JObject)
                If obj Is Nothing Then Continue For
                Dim indexToken = obj.Value(Of JToken)("index")
                Dim textToken = obj.Value(Of JToken)("text")
                If indexToken Is Nothing OrElse textToken Is Nothing Then Continue For
                Dim index As Integer
                If Not Integer.TryParse(indexToken.ToString(), index) Then Continue For
                If index < 0 OrElse index >= targets.Count Then Continue For
                Dim el = targets(index)
                If el Is Nothing Then Continue For
                Dim newText = textToken.ToString()
                Dim oldText = If(el.Text, "")
                If oldText = newText Then Continue For
                el.Text = newText
                SyncTwin(doc, targets, el, newText)
                If Not Object.ReferenceEquals(targets, doc.Elements) Then ReplaceOnce(doc.Elements, oldText, newText)
                If Not ReplaceUnique(doc, oldText, newText) Then ambiguous = True
            Next
            If ambiguous Then RebuildRawText(doc, targets)
        End Sub

        Private Shared Function TargetElements(doc As OcrDocument) As List(Of OcrElement)
            If doc.RawCellElements IsNot Nothing AndAlso doc.RawCellElements.Count > 0 Then Return doc.RawCellElements
            Return doc.Elements
        End Function

        Private Shared Function BuildRequest(pack As DomainPackManifest, doc As OcrDocument, targets As List(Of OcrElement)) As String
            Dim elements As New JArray()
            For i = 0 To targets.Count - 1
                Dim el = targets(i)
                If el Is Nothing Then Continue For
                Dim box As JObject = Nothing
                If el.BoundingBox IsNot Nothing Then
                    box = New JObject(
                        New JProperty("left", el.BoundingBox.Left),
                        New JProperty("top", el.BoundingBox.Top),
                        New JProperty("right", el.BoundingBox.Right),
                        New JProperty("bottom", el.BoundingBox.Bottom))
                End If
                elements.Add(New JObject(
                    New JProperty("index", i),
                    New JProperty("text", If(el.Text, "")),
                    New JProperty("confidence", el.Confidence),
                    New JProperty("box", If(box Is Nothing, CType(JValue.CreateNull(), JToken), box))))
            Next
            Dim options As New JObject()
            If pack.Options IsNot Nothing Then
                For Each opt In pack.Options
                    If opt Is Nothing OrElse String.IsNullOrWhiteSpace(opt.Key) Then Continue For
                    options(opt.Key) = DomainPackRegistry.Instance.GetOptionValue(pack.Id, opt.Key)
                Next
            End If
            Dim root As New JObject(
                New JProperty("imagePath", If(doc.SourceFilePath, "")),
                New JProperty("packId", If(pack.Id, "")),
                New JProperty("options", options),
                New JProperty("elements", elements))
            Return root.ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function RunScript(python As String, script As String, workDir As String, request As String) As String
            Dim proc As Process = Nothing
            Try
                Dim psi As New ProcessStartInfo()
                psi.FileName = python
                psi.Arguments = "-u """ & script & """"
                psi.WorkingDirectory = workDir
                psi.UseShellExecute = False
                psi.CreateNoWindow = True
                psi.RedirectStandardInput = True
                psi.RedirectStandardOutput = True
                psi.RedirectStandardError = True
                psi.StandardOutputEncoding = Encoding.UTF8
                psi.StandardErrorEncoding = Encoding.UTF8
                psi.EnvironmentVariables("PYTHONIOENCODING") = "utf-8"
                psi.EnvironmentVariables("PYTHONUTF8") = "1"

                proc = Process.Start(psi)
                If proc Is Nothing Then Return ""
                Dim stdout = proc.StandardOutput.ReadToEndAsync()
                Dim stderr = proc.StandardError.ReadToEndAsync()
                Dim writer = New StreamWriter(proc.StandardInput.BaseStream, New UTF8Encoding(False)) With {.AutoFlush = True}
                writer.Write(request)
                writer.Close()
                If Not proc.WaitForExit(TimeoutMs) Then
                    Try
                        proc.Kill()
                    Catch ex As Exception
                    End Try
                    Debug.WriteLine("[DomainPack] 脚本超时：" & script)
                    Return ""
                End If
                proc.WaitForExit()
                If Not stdout.Wait(2000) Then Return ""
                Dim err = ""
                Try
                    If stderr.Wait(2000) Then err = stderr.Result
                Catch ex As Exception
                End Try
                If Not String.IsNullOrWhiteSpace(err) Then Debug.WriteLine("[DomainPack] 脚本 stderr：" & err)
                Return stdout.Result
            Catch ex As Exception
                Debug.WriteLine("[DomainPack] 启动脚本失败：" & ex.Message)
                Return ""
            Finally
                If proc IsNot Nothing Then proc.Dispose()
            End Try
        End Function

        Private Shared Function ParsePayload(output As String) As JObject
            Dim trimmed = output.Trim()
            Dim direct = TryParseObject(trimmed)
            If direct IsNot Nothing Then Return direct
            Dim lines = trimmed.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf)
            For i = lines.Length - 1 To 0 Step -1
                Dim parsed = TryParseObject(lines(i).Trim())
                If parsed IsNot Nothing Then Return parsed
            Next
            Return Nothing
        End Function

        Private Shared Function TryParseObject(text As String) As JObject
            If String.IsNullOrWhiteSpace(text) OrElse Not text.StartsWith("{") Then Return Nothing
            Try
                Return JObject.Parse(text)
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        ''' <summary>合并后的行元素没有相同坐标时，把这一处旧文字替换掉。</summary>
        Private Shared Sub ReplaceOnce(list As List(Of OcrElement), oldText As String, newText As String)
            If list Is Nothing OrElse String.IsNullOrEmpty(oldText) OrElse oldText = newText Then Return
            For Each el In list
                If el Is Nothing OrElse String.IsNullOrEmpty(el.Text) Then Continue For
                Dim at = el.Text.IndexOf(oldText, StringComparison.Ordinal)
                If at < 0 OrElse el.Text.LastIndexOf(oldText, StringComparison.Ordinal) <> at Then Continue For
                el.Text = el.Text.Substring(0, at) & newText & el.Text.Substring(at + oldText.Length)
                Return
            Next
        End Sub

        Private Shared Sub SyncTwin(doc As OcrDocument, targets As List(Of OcrElement), source As OcrElement, newText As String)
            Dim other = If(Object.ReferenceEquals(targets, doc.Elements), doc.RawCellElements, doc.Elements)
            If other Is Nothing OrElse source.BoundingBox Is Nothing Then Return
            For Each el In other
                If el Is Nothing OrElse el.BoundingBox Is Nothing Then Continue For
                If el.BoundingBox.Left = source.BoundingBox.Left AndAlso
                   el.BoundingBox.Top = source.BoundingBox.Top AndAlso
                   el.BoundingBox.Right = source.BoundingBox.Right AndAlso
                   el.BoundingBox.Bottom = source.BoundingBox.Bottom Then
                    el.Text = newText
                End If
            Next
        End Sub

        ''' <summary>原文里这段文字只出现一次时才替换。出现多次则交给整表重建，避免改错行。</summary>
        Private Shared Function ReplaceUnique(doc As OcrDocument, oldText As String, newText As String) As Boolean
            If String.IsNullOrEmpty(oldText) OrElse oldText = newText OrElse String.IsNullOrEmpty(doc.RawText) Then Return True
            Dim first = doc.RawText.IndexOf(oldText, StringComparison.Ordinal)
            If first < 0 Then Return True
            If doc.RawText.LastIndexOf(oldText, StringComparison.Ordinal) <> first Then Return False
            doc.RawText = doc.RawText.Substring(0, first) & newText & doc.RawText.Substring(first + oldText.Length)
            Return True
        End Function

        Private Shared Sub RebuildRawText(doc As OcrDocument, targets As List(Of OcrElement))
            Dim lines As New List(Of String)()
            For Each el In targets
                If el IsNot Nothing Then lines.Add(If(el.Text, ""))
            Next
            doc.RawText = String.Join(vbCrLf, lines)
        End Sub

        Private Shared Function ResolvePython() As String
            SyncLock _pyLock
                If Not String.IsNullOrEmpty(_pythonPath) AndAlso File.Exists(_pythonPath) Then Return _pythonPath
                Try
                    If File.Exists(PythonBootstrapper.EmbeddedPythonExe) Then
                        _pythonPath = PythonBootstrapper.EmbeddedPythonExe
                        Return _pythonPath
                    End If
                Catch ex As Exception
                End Try
                Try
                    Dim found = PythonBootstrapper.Instance.DetectSystemPython()
                    If Not String.IsNullOrEmpty(found) AndAlso File.Exists(found) Then
                        _pythonPath = found
                        Return _pythonPath
                    End If
                Catch ex As Exception
                End Try
                Return ""
            End SyncLock
        End Function
    End Class
End Namespace
