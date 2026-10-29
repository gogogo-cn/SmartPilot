Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Text
Imports System.Threading
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.DomainPack

    ''' <summary>
    ''' 机械制图矫正的配置。检测器仍是代码；这里只决定开哪些阶段、两个阈值，以及整框文字替换。
    ''' 缺文件或坏文件时用内置默认，行为和原来的硬编码一致。
    ''' 用户覆盖：%AppData%\ExcelAIVSTOAddIn\Correction\mechanical_drawing.json
    ''' 仓库模板：Config\Correction\mechanical_drawing.json
    ''' </summary>
    Public NotInheritable Class CorrectionProfile

        Public Const DefaultInkLuminance As Integer = 150
        Public Const DefaultRingScore As Double = 0.34

        Private Shared ReadOnly ActiveLocal As New ThreadLocal(Of CorrectionProfile)()
        Private ReadOnly _stages As New Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
        Private ReadOnly _rules As New List(Of CorrectionRule)()

        Public Property InkLuminance As Integer = DefaultInkLuminance
        Public Property RingScore As Double = DefaultRingScore

        Public Shared ReadOnly Property Active As CorrectionProfile
            Get
                Return ActiveLocal.Value
            End Get
        End Property

        Public Shared Function Enter(profile As CorrectionProfile) As IDisposable
            Dim scope As New Scope(ActiveLocal.Value)
            ActiveLocal.Value = profile
            Return scope
        End Function

        ''' <summary>显式传入 150 时改用当前配置；其它阈值（例如墙线 210）保持调用方的值。</summary>
        Public Shared Function ResolveInk(requested As Integer) As Integer
            If requested <> DefaultInkLuminance OrElse Active Is Nothing Then Return requested
            If Active.InkLuminance <= 0 OrElse Active.InkLuminance > 255 Then Return DefaultInkLuminance
            Return Active.InkLuminance
        End Function

        Public Shared Function ResolveRingScore() As Double
            If Active Is Nothing Then Return DefaultRingScore
            If Active.RingScore <= 0 OrElse Active.RingScore > 1 Then Return DefaultRingScore
            Return Active.RingScore
        End Function

        ''' <summary>把界面上的机械制图配置写到用户目录。返回空字符串表示成功，否则是错误说明。</summary>
        Public Shared Function SaveMechanical(json As String) As String
            If String.IsNullOrWhiteSpace(json) Then Return "内容为空"
            Dim root As JObject
            Try
                root = JObject.Parse(json)
            Catch ex As Exception
                Return "不是合法的 JSON"
            End Try
            If root("stages") Is Nothing OrElse root("stages").Type <> JTokenType.Object Then Return "缺少 stages"
            Try
                Dim target = UserPath()
                Directory.CreateDirectory(Path.GetDirectoryName(target))
                File.WriteAllText(target, root.ToString(Formatting.Indented), New UTF8Encoding(False))
                Return ""
            Catch ex As Exception
                Return "写入失败：" & ex.Message
            End Try
        End Function

        Public Shared Function LoadMechanical() As CorrectionProfile
            Dim profile = CreateDefault()
            Try
                OverlayFile(profile, FindTemplate())
                OverlayFile(profile, UserPath())
                ApplyGroupOption(profile)
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[Correction] 读取配置失败，使用默认：" & ex.Message)
            End Try
            Return profile
        End Function

        Public Function IsStageOn(name As String) As Boolean
            Dim value As Boolean
            If _stages.TryGetValue(name, value) Then Return value
            Return True
        End Function

        Public Sub ApplyRules(doc As OcrDocument)
            If doc Is Nothing OrElse _rules.Count = 0 Then Return
            For Each rule In _rules
                If String.IsNullOrEmpty(rule.Text) OrElse rule.Text = rule.Replace Then Continue For
                ReplaceBoxes(doc.RawCellElements, rule)
                ReplaceBoxes(doc.Elements, rule)
                ReplaceLines(doc, rule)
            Next
        End Sub

        Private Shared Function CreateDefault() As CorrectionProfile
            Dim profile As New CorrectionProfile()
            For Each name In New String() {"gdtFrame", "balloon", "vertical", "datum", "roughness", "mark", "fcf", "leader", "groupDimensions", "leaderRules"}
                profile._stages(name) = True
            Next
            ' datumScan 是全图扫圆的兜底，慢且有误报，默认关（IsStageOn 对未知名返回 True，必须显式登记）。
            profile._stages("datumScan") = False
            Return profile
        End Function

        Private Shared Sub OverlayFile(profile As CorrectionProfile, path As String)
            If String.IsNullOrEmpty(path) OrElse Not File.Exists(path) Then Return
            Dim root = JObject.Parse(File.ReadAllText(path))
            Dim ink = root.Value(Of JToken)("inkLuminance")
            Dim ring = root.Value(Of JToken)("ringScore")
            Dim inkValue As Integer
            Dim ringValue As Double
            If ink IsNot Nothing AndAlso Integer.TryParse(ink.ToString(), inkValue) Then profile.InkLuminance = inkValue
            If ring IsNot Nothing AndAlso Double.TryParse(ring.ToString(), Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, ringValue) Then
                profile.RingScore = ringValue
            End If

            Dim stages = TryCast(root("stages"), JObject)
            If stages IsNot Nothing Then
                For Each prop In stages.Properties()
                    If prop.Value IsNot Nothing AndAlso prop.Value.Type = JTokenType.Boolean Then
                        profile._stages(prop.Name) = prop.Value.Value(Of Boolean)()
                    End If
                Next
            End If

            Dim rules = TryCast(root("rules"), JArray)
            If rules Is Nothing Then Return
            profile._rules.Clear()
            For Each item In rules
                Dim obj = TryCast(item, JObject)
                If obj Is Nothing Then Continue For
                Dim fromText = obj.Value(Of String)("text")
                Dim toText = obj.Value(Of String)("replace")
                If String.IsNullOrEmpty(fromText) OrElse toText Is Nothing Then Continue For
                profile._rules.Add(New CorrectionRule(fromText, toText))
            Next
        End Sub

        ''' <summary>已有的 group_dimensions 开关为 false 时关闭尺寸归组。</summary>
        Private Shared Sub ApplyGroupOption(profile As CorrectionProfile)
            Dim value = DomainPackRegistry.Instance.GetOptionValue("mechanical_drawing", "group_dimensions")
            If String.Equals(value, "false", StringComparison.OrdinalIgnoreCase) Then
                profile._stages("groupDimensions") = False
            End If
        End Sub

        Private Shared Function UserPath() As String
            Return Path.Combine(Utils.PathHelper.AppDataRoot, "Correction", "mechanical_drawing.json")
        End Function

        Private Shared Function FindTemplate() As String
            Dim dir = AppDomain.CurrentDomain.BaseDirectory
            For i = 0 To 6
                If String.IsNullOrEmpty(dir) Then Exit For
                Dim candidate = Path.Combine(dir, "Config", "Correction", "mechanical_drawing.json")
                If File.Exists(candidate) Then Return candidate
                dir = Path.GetDirectoryName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar))
            Next
            Return ""
        End Function

        Private Shared Sub ReplaceBoxes(list As IList(Of OcrElement), rule As CorrectionRule)
            If list Is Nothing Then Return
            For Each el In list
                If el Is Nothing Then Continue For
                If String.Equals(If(el.Text, "").Trim(), rule.Text, StringComparison.Ordinal) Then el.Text = rule.Replace
            Next
        End Sub

        Private Shared Sub ReplaceLines(doc As OcrDocument, rule As CorrectionRule)
            If String.IsNullOrEmpty(doc.RawText) Then Return
            Dim lines = doc.RawText.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(vbLf)
            Dim changed = False
            For i = 0 To lines.Length - 1
                If String.Equals(lines(i).Trim(), rule.Text, StringComparison.Ordinal) Then
                    lines(i) = rule.Replace
                    changed = True
                End If
            Next
            If changed Then doc.RawText = String.Join(vbCrLf, lines)
        End Sub

        Private NotInheritable Class Scope
            Implements IDisposable

            Private ReadOnly _previous As CorrectionProfile
            Private _disposed As Boolean

            Public Sub New(previous As CorrectionProfile)
                _previous = previous
            End Sub

            Public Sub Dispose() Implements IDisposable.Dispose
                If _disposed Then Return
                _disposed = True
                ActiveLocal.Value = _previous
            End Sub
        End Class
    End Class

    Public NotInheritable Class CorrectionRule
        Public ReadOnly Text As String
        Public ReadOnly Replace As String

        Public Sub New(text As String, replaceWith As String)
            Me.Text = text
            Me.Replace = replaceWith
        End Sub
    End Class
End Namespace
