Imports System
Imports System.Collections.Generic
Imports System.IO
Imports System.Linq
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.DomainPack

    ''' <summary>
    ''' 专项识别模块注册表。
    '''
    ''' 职责：
    '''   1. 登记内置模块（机械制图 GDT —— 已有的 2105 行 VB 实现原地保留，不重写）
    '''   2. 扫描用户目录 %AppData%\ExcelAIVSTOAddIn\DomainPacks\*.md
    '''   3. 持久化启用/禁用与模块选项取值
    '''   4. 把旧的 recognitionMode（mechanical/gdt/drawing）映射到内置 GDT，保证向后兼容
    '''
    ''' 由此 OcrServiceManager 不再硬编码字符串判断，新增专项只需丢一个 .md。
    ''' </summary>
    Public Class DomainPackRegistry

        Private Shared ReadOnly _instance As New Lazy(Of DomainPackRegistry)(
            Function() New DomainPackRegistry(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As DomainPackRegistry
            Get
                Return _instance.Value
            End Get
        End Property

        ''' <summary>机械制图：既有实现代号。第三方模块不要占用。</summary>
        Public Const BUILTIN_MECHANICAL_ID As String = "mechanical_drawing"
        Public Const BUILTIN_MECHANICAL_MODULE As String = "GdtSymbolRecognizer"

        Private ReadOnly _lock As New Object()
        Private ReadOnly _rootDir As String
        Private ReadOnly _statePath As String
        Private _builtins As List(Of DomainPackManifest)
        Private _installed As List(Of DomainPackManifest)

        Private Sub New()
            _rootDir = Path.Combine(Utils.PathHelper.AppDataRoot, "DomainPacks")
            _statePath = Path.Combine(Utils.PathHelper.AppDataRoot, "domain_packs_state.json")
            _builtins = New List(Of DomainPackManifest)()
            RegisterBuiltinModules()
            Try
                If Not Directory.Exists(_rootDir) Then Directory.CreateDirectory(_rootDir)
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[DomainPack] 无法创建模块目录：" & ex.Message)
            End Try
            Reload()
        End Sub

        Private Sub RegisterBuiltinModules()
            Dim mech As New DomainPackManifest()
            mech.Id = BUILTIN_MECHANICAL_ID
            mech.Name = "机械制图"
            mech.Description = "GD&T 形位公差框补全、球标圆圈与引线还原（内置实现，精度最高）"
            mech.Category = "工程图"
            mech.Version = "1.0.0"
            mech.Pipeline = DomainPackPipeline.BuiltIn
            mech.ModuleRef = BUILTIN_MECHANICAL_MODULE
            mech.IsBuiltIn = True
            mech.Enabled = True
            mech.Options.Add(New DomainPackOption() With {
                .Key = "group_dimensions",
                .Label = "尺寸标注归类",
                .Type = "bool",
                .DefaultValue = "true"
            })
            _builtins.Add(mech)
        End Sub

        ''' <summary>重新扫描用户目录。</summary>
        Public Sub Reload()
            SyncLock _lock
                Dim state = LoadState()
                Dim disabled = DisabledSet(state)
                Dim list As New List(Of DomainPackManifest)()
                Try
                    If Directory.Exists(_rootDir) Then
                        For Each f In Directory.GetFiles(_rootDir, "*.md").OrderBy(Function(x) x)
                            Dim m = DomainPackLoader.ParseFile(f)
                            If m IsNot Nothing Then
                                m.IsBuiltIn = False
                                m.Enabled = Not disabled.Contains(m.Id)
                                list.Add(m)
                            End If
                        Next
                    End If
                Catch ex As Exception
                    Diagnostics.Debug.WriteLine("[DomainPack] 扫描失败：" & ex.Message)
                End Try

                For Each b In _builtins
                    b.Enabled = Not disabled.Contains(b.Id)
                Next

                _installed = list
                ApplyOptionDefaults(state)
            End SyncLock
        End Sub

        ''' <summary>模块目录（供安装/卸载定位）。</summary>
        Public ReadOnly Property RootDirectory As String
            Get
                Return _rootDir
            End Get
        End Property

        ''' <summary>内置 + 已安装的全量清单（供界面渲染）。</summary>
        Public Function All() As List(Of DomainPackManifest)
            SyncLock _lock
                Dim list As New List(Of DomainPackManifest)()
                list.AddRange(_builtins)
                list.AddRange(_installed)
                Return list
            End SyncLock
        End Function

        Public Function [Get](id As String) As DomainPackManifest
            If String.IsNullOrWhiteSpace(id) Then Return Nothing
            SyncLock _lock
                Return All().FirstOrDefault(Function(m) String.Equals(m.Id, id, StringComparison.OrdinalIgnoreCase))
            End SyncLock
        End Function

        ''' <summary>
        ''' 按旧的 recognitionMode 字符串解析模块（向后兼容）。
        ''' mechanical / gdt / drawing → 内置机械制图；否则按 id 精确匹配。
        ''' </summary>
        Public Function ResolveByMode(recognitionMode As String) As DomainPackManifest
            If String.IsNullOrWhiteSpace(recognitionMode) Then Return Nothing
            Dim mode = recognitionMode.Trim().ToLowerInvariant()
            If mode = "mechanical" OrElse mode = "gdt" OrElse mode = "drawing" Then
                Return [Get](BUILTIN_MECHANICAL_ID)
            End If
            Return [Get](mode)
        End Function

        Public Function IsEnabled(id As String) As Boolean
            Dim m = [Get](id)
            Return m IsNot Nothing AndAlso m.Enabled
        End Function

        Public Function SetEnabled(id As String, enabled As Boolean) As Boolean
            SyncLock _lock
                Dim state = LoadState()
                Dim disabled As New List(Of String)()
                Dim arr = state.Value(Of JArray)("disabled")
                If arr IsNot Nothing Then
                    For Each t In arr
                        Dim v = t.ToString()
                        If Not disabled.Contains(v) Then disabled.Add(v)
                    Next
                End If
                If enabled Then
                    disabled.RemoveAll(Function(x) String.Equals(x, id, StringComparison.OrdinalIgnoreCase))
                ElseIf Not disabled.Contains(id) Then
                    disabled.Add(id)
                End If

                state("disabled") = New JArray(disabled.ToArray())
                SaveState(state)
                Reload()
                Return True
            End SyncLock
        End Function

        ''' <summary>读取模块某开关的当前取值（无记录则返回清单默认值）。</summary>
        Public Function GetOptionValue(id As String, key As String) As String
            If String.IsNullOrWhiteSpace(id) OrElse String.IsNullOrWhiteSpace(key) Then Return ""
            Dim state = LoadState()
            Dim opts = state.Value(Of JObject)("options")
            If opts IsNot Nothing Then
                Dim pack = opts.Value(Of JObject)(id)
                If pack IsNot Nothing AndAlso pack.Value(Of JToken)(key) IsNot Nothing Then
                    Return pack.Value(Of JToken)(key).ToString()
                End If
            End If
            Dim m = [Get](id)
            If m Is Nothing Then Return ""
            Dim opt = m.Options.FirstOrDefault(Function(o) String.Equals(o.Key, key, StringComparison.OrdinalIgnoreCase))
            Return If(opt Is Nothing, "", opt.DefaultValue)
        End Function

        Public Sub SetOptionValue(id As String, key As String, value As String)
            If String.IsNullOrWhiteSpace(id) OrElse String.IsNullOrWhiteSpace(key) Then Return
            SyncLock _lock
                Dim state = LoadState()
                Dim opts = state.Value(Of JObject)("options")
                If opts Is Nothing Then
                    opts = New JObject()
                    state("options") = opts
                End If
                Dim pack = opts.Value(Of JObject)(id)
                If pack Is Nothing Then
                    pack = New JObject()
                    opts(id) = pack
                End If
                pack(key) = value
                SaveState(state)
            End SyncLock
        End Sub

        ''' <summary>
        ''' 安装模块：校验 markdown 清单 → 落盘到用户目录 → 重新加载。
        ''' 同名已存在时覆盖（等价于升级）。返回安装结果对象。
        ''' </summary>
        Public Function InstallFromMarkdown(content As String) As DomainPackInstallResult
            Dim res As New DomainPackInstallResult()
            If String.IsNullOrWhiteSpace(content) Then
                res.Error = "内容为空"
                Return res
            End If

            Dim m = DomainPackLoader.ParseMarkdown(content)
            If m Is Nothing Then
                res.Error = "不是合法的模块清单：缺少 frontmatter 或必填字段（id / name / pipeline）"
                Return res
            End If
            If String.IsNullOrWhiteSpace(m.Id) Then
                res.Error = "清单缺少 id"
                Return res
            End If
            If Not String.IsNullOrWhiteSpace(m.ErrorText) Then
                res.Error = m.ErrorText
                Return res
            End If

            Dim builtinIds = (From b In _builtins Select b.Id).ToList()
            If builtinIds.Contains(m.Id) Then
                res.Error = "id 与内置模块冲突（内置模块不可覆盖）：" & m.Id
                Return res
            End If

            Try
                Directory.CreateDirectory(_rootDir)
                Dim scriptBody = DomainPackLoader.ExtractFencedBlock(content, "python")
                Dim scriptName = DomainPackPythonRunner.SafeScriptFileName(m.ModuleRef, m.Id)
                If m.Pipeline = DomainPackPipeline.Python AndAlso Not String.IsNullOrWhiteSpace(scriptBody) AndAlso String.IsNullOrEmpty(scriptName) Then
                    res.Error = "python 脚本名无效。module 只能是当前目录下的 .py 文件名"
                    Return res
                End If
                Dim target = Path.Combine(_rootDir, m.Id & ".md")
                File.WriteAllText(target, content, New UTF8Encoding(False))
                If Not String.IsNullOrWhiteSpace(scriptBody) AndAlso Not String.IsNullOrEmpty(scriptName) Then
                    Dim scriptPath = Path.Combine(_rootDir, scriptName)
                    File.WriteAllText(scriptPath, scriptBody.Replace(vbCr, ""), New UTF8Encoding(False))
                    res.ScriptPath = scriptPath
                End If
                Reload()
                res.Success = True
                res.Id = m.Id
                res.Name = If(String.IsNullOrWhiteSpace(m.Name), m.Id, m.Name)
                res.Path = target
                Return res
            Catch ex As Exception
                res.Error = "写入失败：" & ex.Message
                Return res
            End Try
        End Function

        ''' <summary>卸载（删除清单文件）。内置模块拒绝卸载。</summary>
        Public Function Uninstall(id As String) As Boolean
            Dim m = [Get](id)
            If m Is Nothing Then Return False
            If m.IsBuiltIn Then Return False
            Try
                If Not String.IsNullOrEmpty(m.FilePath) AndAlso File.Exists(m.FilePath) Then File.Delete(m.FilePath)
                Dim scriptName = DomainPackPythonRunner.SafeScriptFileName(m.ModuleRef, m.Id)
                If Not String.IsNullOrEmpty(scriptName) Then
                    Dim scriptPath = Path.Combine(_rootDir, scriptName)
                    If File.Exists(scriptPath) Then File.Delete(scriptPath)
                End If
                Reload()
                Return True
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[DomainPack] 卸载失败：" & ex.Message)
                Return False
            End Try
        End Function

        ' ---------------- 状态文件 ----------------

        Private Function LoadState() As JObject
            Try
                If File.Exists(_statePath) Then
                    Return JObject.Parse(File.ReadAllText(_statePath))
                End If
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[DomainPack] 状态文件损坏，按空处理：" & ex.Message)
            End Try
            Return New JObject()
        End Function

        Private Function DisabledSet(state As JObject) As HashSet(Of String)
            Dim set_ As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            Dim arr = state.Value(Of JArray)("disabled")
            If arr IsNot Nothing Then
                For Each t In arr
                    set_.Add(t.ToString())
                Next
            End If
            Return set_
        End Function

        Private Sub ApplyOptionDefaults(state As JObject)
            ' 选项取值随用随读（GetOptionValue 兜默认值），这里仅保证 JSON 形状存在
            If state.Value(Of JObject)("options") Is Nothing Then state("options") = New JObject()
        End Sub

        Private Sub SaveState(state As JObject)
            Try
                File.WriteAllText(_statePath, state.ToString(Formatting.Indented), New UTF8Encoding(False))
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[DomainPack] 状态写入失败：" & ex.Message)
            End Try
        End Sub
    End Class

    Public Class DomainPackInstallResult
        Public Property Success As Boolean
        Public Property Id As String
        Public Property Name As String
        Public Property Path As String
        ''' <summary>安装时从 ```python 代码块写出的脚本路径。没有代码块则为空。</summary>
        Public Property ScriptPath As String
        Public Property [Error] As String
    End Class
End Namespace
