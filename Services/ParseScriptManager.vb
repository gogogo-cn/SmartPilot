Imports System.IO
Imports System.Threading
Imports System.Collections.Generic
Imports Newtonsoft.Json

Namespace Services

    ''' <summary>
    ''' 解析字段定义
    ''' </summary>
    Public Class ParseField
        Public Property Name As String       ' 字段显示名称（如"发票代码"）
        Public Property Key As String         ' 字段键名（如"invoice_code"）
        Public Property Pattern As String     ' 正则表达式模式（可选）
        Public Property Required As Boolean = False  ' 是否必填
        Public Property DefaultValue As String = ""   ' 默认值
        Public Property Visible As Boolean = True     ' 是否在结果中显示（可由用户在字段配置中勾选）
    End Class

    ''' <summary>
    ''' OCR解析脚本定义
    ''' </summary>
    Public Class ParseScript
        Public Property Id As String
        Public Property Name As String           ' 脚本名称
        Public Property Description As String    ' 描述
        Public Property Icon As String = "📝"    ' 图标
        Public Property Category As String = "通用" ' 分类：发票/证件/票据/通用/自定义
        Public Property Version As Integer = 1
        Public Property IsBuiltIn As Boolean = False  ' 是否内置模板
        Public Property CreatedAt As DateTime
        Public Property LastModified As DateTime
        Public Property Fields As List(Of ParseField)  ' 要提取的字段列表
        Public Property Script As String = ""          ' JavaScript 解析脚本
        Public Property OutputFormat As String = "table" ' 输出格式: table | json | key-value | excel
        Public Property SampleText As String = ""       ' 示例文本（用于测试）
        Public Property Tags As List(Of String)          ' 标签
        Public Property UsageCount As Integer = 0        ' 使用次数
    End Class

    ''' <summary>
    ''' OCR解析脚本管理器（单例模式）
    ''' 管理解析脚本的CRUD操作和执行
    ''' </summary>
    Public Class ParseScriptManager
        Private Shared ReadOnly _instance As New Lazy(Of ParseScriptManager)(Function() New ParseScriptManager(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
        Private _scriptsPath As String
        Private _defaultScriptsPath As String
        Private _scripts As Dictionary(Of String, ParseScript)
        Private ReadOnly _lock As New Object()

        Public Shared ReadOnly Property Instance As ParseScriptManager
            Get
                Return _instance.Value
            End Get
        End Property

        Private Sub New()
            Try
                ' 用户脚本放在本机 AppData，Excel 和单独运行的识图服务读写同一份。
                _scriptsPath = Path.Combine(Utils.PathHelper.AppDataRoot, "ParseScripts")
                
                If Not Directory.Exists(_scriptsPath) Then
                    Directory.CreateDirectory(_scriptsPath)
                End If
                
                ' 默认模板仍从插件自带的 Config 读取。
                _defaultScriptsPath = ResolveDefaultScriptsPath()
                ImportLegacyScripts()
                
                ' 初始化字典
                _scripts = New Dictionary(Of String, ParseScript)(StringComparer.OrdinalIgnoreCase)
                
                ' 先加载内置模板
                EnsureBuiltInTemplates()
                ' 再加载所有脚本（包括内置缓存和用户自定义）
                LoadAllScripts()
            Catch ex As Exception
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 初始化失败: {ex.Message}")
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 堆栈: {ex.StackTrace}")
                ' 确保即使失败也有空字典
                If _scripts Is Nothing Then
                    _scripts = New Dictionary(Of String, ParseScript)(StringComparer.OrdinalIgnoreCase)
                End If
            End Try
        End Sub

        ''' <summary>
        ''' 加载所有解析脚本（支持 .md 和 .json 格式）
        ''' </summary>
        Private Sub LoadAllScripts()
            SyncLock _lock
                _scripts = New Dictionary(Of String, ParseScript)(StringComparer.OrdinalIgnoreCase)
                
                ' 先加载内置模板缓存
                For Each scriptFile In Directory.GetFiles(_scriptsPath, "*.md")
                    Try
                        Dim markdown = File.ReadAllText(scriptFile, System.Text.Encoding.UTF8)
                        Dim script = ParseScriptFromMarkdown(markdown)
                        If script IsNot Nothing AndAlso Not String.IsNullOrEmpty(script.Id) Then
                            _scripts(script.Id) = script
                            Diagnostics.Debug.WriteLine($"[ParseScriptManager] 已加载: {script.Name} ({script.Id})")
                        End If
                    Catch ex As Exception
                        Diagnostics.Debug.WriteLine("[ParseScriptManager] 加载 MD 脚本失败: " & scriptFile & ", 错误: " & ex.Message)
                    End Try
                Next
                
                ' 兼容旧 JSON 格式脚本（自动转换为 MD）
                For Each scriptFile In Directory.GetFiles(_scriptsPath, "*.json")
                    Try
                        Dim json = File.ReadAllText(scriptFile, System.Text.Encoding.UTF8)
                        Dim script = JsonConvert.DeserializeObject(Of ParseScript)(json)
                        If script IsNot Nothing AndAlso Not String.IsNullOrEmpty(script.Id) Then
                            ' 转换为 MD 格式并保存
                            script.IsBuiltIn = False
                            SaveScriptAsMarkdown(script)
                            _scripts(script.Id) = script
                            Diagnostics.Debug.WriteLine($"[ParseScriptManager] 已转换并加载: {script.Name} ({script.Id})")
                        End If
                    Catch ex As Exception
                        Diagnostics.Debug.WriteLine("[ParseScriptManager] 加载 JSON 脚本失败: " & scriptFile & ", 错误: " & ex.Message)
                    End Try
                Next
                
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 总计加载 " & _scripts.Count & " 个解析脚本")
            End SyncLock
        End Sub

        Private Function ResolveDefaultScriptsPath() As String
            Dim candidates As New List(Of String)
            candidates.Add(Path.Combine(Utils.PathHelper.ConfigDirectory, "DefaultParseScripts"))
            Dim dir = Utils.PathHelper.PluginInstallPath
            For i = 1 To 6
                dir = Path.GetDirectoryName(If(dir, "").TrimEnd("\"c, "/"c))
                If String.IsNullOrEmpty(dir) Then Exit For
                candidates.Add(Path.Combine(dir, "Config", "DefaultParseScripts"))
            Next
            For Each candidate In candidates
                If Directory.Exists(candidate) AndAlso Directory.GetFiles(candidate, "*.md").Length > 0 Then
                    Return candidate
                End If
            Next
            Return Path.Combine(Utils.PathHelper.ConfigDirectory, "DefaultParseScripts")
        End Function

        ''' <summary>
        ''' 把以前放在各进程 Config\ParseScripts 里的脚本抄进共享目录，已有文件不覆盖。
        ''' </summary>
        Private Sub ImportLegacyScripts()
            Dim legacyDirs As New List(Of String)
            legacyDirs.Add(Path.Combine(Utils.PathHelper.ConfigDirectory, "ParseScripts"))
            Dim dir = Utils.PathHelper.PluginInstallPath
            For i = 1 To 6
                dir = Path.GetDirectoryName(If(dir, "").TrimEnd("\"c, "/"c))
                If String.IsNullOrEmpty(dir) Then Exit For
                legacyDirs.Add(Path.Combine(dir, "Config", "ParseScripts"))
            Next
            Dim sharedFull = Path.GetFullPath(_scriptsPath)
            For Each legacy In legacyDirs
                If Not Directory.Exists(legacy) Then Continue For
                Dim legacyFull = Path.GetFullPath(legacy)
                If String.Equals(legacyFull, sharedFull, StringComparison.OrdinalIgnoreCase) Then Continue For
                For Each scriptFile In Directory.GetFiles(legacy, "*.md")
                    Dim dest = Path.Combine(_scriptsPath, Path.GetFileName(scriptFile))
                    If Not File.Exists(dest) Then
                        Try
                            File.Copy(scriptFile, dest)
                        Catch ex As Exception
                            Diagnostics.Debug.WriteLine("[ParseScriptManager] 迁移脚本失败: " & ex.Message)
                        End Try
                    End If
                Next
            Next
        End Sub

        ''' <summary>
        ''' 确保内置模板存在（从 MD 文件加载或创建）
        ''' </summary>
        Private Sub EnsureBuiltInTemplates()
            ' 检查默认模板目录是否存在
            If Not Directory.Exists(_defaultScriptsPath) Then
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 默认模板目录不存在: " & _defaultScriptsPath)
                ' 使用备用方案创建基本模板
                CreateFallbackTemplates()
                Return
            End If

            Dim loadedCount = 0
            ' 从 MD 源文件加载内置脚本到缓存
            For Each mdFile In Directory.GetFiles(_defaultScriptsPath, "*.md")
                Try
                    Dim markdown = File.ReadAllText(mdFile, System.Text.Encoding.UTF8)
                    Dim template = ParseScriptFromMarkdown(markdown)

                    If template IsNot Nothing AndAlso Not String.IsNullOrEmpty(template.Id) Then
                        ' 总是从源文件更新缓存
                        template.IsBuiltIn = True
                        If template.CreatedAt = Nothing Then
                            template.CreatedAt = DateTime.Now
                        End If
                        template.LastModified = DateTime.Now

                        ' 设置示例文本
                        SetSampleText(template)

                        ' 保存到缓存目录（覆盖旧文件）
                        Dim cacheFile = Path.Combine(_scriptsPath, template.Id & ".md")
                        Dim cacheMarkdown = ConvertScriptToMarkdown(template)
                        ' 使用 UTF-8 with BOM 确保 Windows 系统正确识别
                        File.WriteAllText(cacheFile, cacheMarkdown, System.Text.Encoding.UTF8)
                        loadedCount += 1
                        Diagnostics.Debug.WriteLine($"[ParseScriptManager] 已缓存内置模板: {template.Name} ({template.Id})")
                    End If
                Catch ex As Exception
                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 加载默认 MD 模板失败: " & mdFile & ", 错误: " & ex.Message)
                End Try
            Next

            If loadedCount > 0 Then
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 已从源目录缓存 {loadedCount} 个内置模板")
            End If

            ' 清理：删除旧的 DefaultParseScripts.json 文件
            Dim oldJsonFile = Path.Combine(Utils.PathHelper.ConfigDirectory, "DefaultParseScripts.json")
            If File.Exists(oldJsonFile) Then
                Try
                    File.Delete(oldJsonFile)
                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 已删除旧的 JSON 模板文件")
                Catch ex As Exception
                    ' 忽略删除失败
                End Try
            End If
        End Sub

        ''' <summary>
        ''' 为每个模板设置示例文本
        ''' </summary>
        Private Sub SetSampleText(template As ParseScript)
            If String.IsNullOrEmpty(template.SampleText) Then
                Select Case template.Id
                    Case "invoice"
                        template.SampleText = "增值税普通发票" & vbCrLf & "发票代码: 142001911510" & vbCrLf & "发票号码: 12345678" & vbCrLf & "开票日期: 2024年01月15日" & vbCrLf & "购买方名称: XX科技有限公司" & vbCrLf & "纳税人识别号: 91110000XXXXXXXXXX"
                    Case "idcard"
                        template.SampleText = "姓名 张三" & vbCrLf & "性别 男" & vbCrLf & "民族 汉" & vbCrLf & "出生 1990年1月1日" & vbCrLf & "住址 XX省XX市XX区XX街道XX号" & vbCrLf & "公民身份号码 11010119900101123X"
                    Case "business_card"
                        template.SampleText = "张三" & vbCrLf & "总经理" & vbCrLf & "XX科技有限公司" & vbCrLf & "138-0000-0000" & vbCrLf & "010-88888888" & vbCrLf & "zhangsan@example.com"
                    Case "bank_slip"
                        template.SampleText = "中国工商银行电子回单" & vbCrLf & "交易日期: 2024-01-15 14:30:25" & vbCrLf & "付款人账号: 6222************1234" & vbCrLf & "付款人户名: 张三" & vbCrLf & "金额: ¥50,000.00" & vbCrLf & "摘要: 货款"
                    Case "express"
                        template.SampleText = "顺丰速运" & vbCrLf & "运单号: SF1234567890123" & vbCrLf & "寄件人: 张三" & vbCrLf & "电话: 138****1234" & vbCrLf & "收件人: 李四" & vbCrLf & "电话: 139****5678"
                    Case "business_license"
                        template.SampleText = "营业执照" & vbCrLf & "统一社会信用代码: 91110108MA01XXXXX2X" & vbCrLf & "名称: XX科技有限公司" & vbCrLf & "法定代表人: 张三" & vbCrLf & "注册资本: 1000万元人民币"
                    Case "train_ticket"
                        template.SampleText = "北京局售" & vbCrLf & "北京南 → 上海虹桥" & vbCrLf & "G102 次 2024年01月15日 08:00 开" & vbCrLf & "06车 12A号 二等座" & vbCrLf & "¥553.0元 网购票"
                    Case "driver_license"
                        template.SampleText = "中华人民共和国机动车驾驶证" & vbCrLf & "证号: 11010119900101123X" & vbCrLf & "姓名: 张三" & vbCrLf & "准驾车型: C1" & vbCrLf & "有效期限: 2015-06-15 至 2025-06-15"
                    Case "vehicle_license"
                        template.SampleText = "中华人民共和国机动车行驶证" & vbCrLf & "车牌号码: 京A·XXXXX" & vbCrLf & "车辆类型: 小型轿车" & vbCrLf & "所有人: 张三" & vbCrLf & "品牌型号: 大众SVW71412AR"
                    Case "date_extractor"
                        template.SampleText = "合同编号: HT202401150001" & vbCrLf & "签订日期: 2024年1月15日" & vbCrLf & "生效日期: 2024-02-01" & vbCrLf & "到期日期: 2025/12/31"
                    Case "amount_extractor"
                        template.SampleText = "采购订单 PO-2024001" & vbCrLf & "商品A 单价: ¥125.00 小计: ¥12,500.00" & vbCrLf & "商品B 单价: ￥88.50 小计: ￥4,425.00" & vbCrLf & "总计: ¥16,625.00"
                    Case "invoice_electronic"
                        template.SampleText = "增值税电子普通发票" & vbCrLf & "发票代码: 142001911510" & vbCrLf & "发票号码: 12345678" & vbCrLf & "开票日期: 2024年01月15日" & vbCrLf & "*办公用品*打印纸*A4*箱*10*150.00*1500.00*13%*195.00"
                    Case "mechanical_drawing"
                        template.SampleText = "技术要求" & vbCrLf & "1.未注倒角 C0.5" & vbCrLf & "2.去除毛刺" & vbCrLf & "零件名称：法兰盘" & vbCrLf & "图样代号：HT-01-002" & vbCrLf & "材料：45" & vbCrLf & "比例：1:2" & vbCrLf & "重量：0.86kg" & vbCrLf & "共 2 张  第 1 张" & vbCrLf & "设计：张三 2024.03.15" & vbCrLf & "校对：李四" & vbCrLf & "审核：王五" & vbCrLf & "某某机械有限公司"
                    Case "dimension_annotation"
                        template.SampleText = "Ø25±0.02" & vbCrLf & "R3" & vbCrLf & "C1" & vbCrLf & "M8×1.25-6H" & vbCrLf & "45°" & vbCrLf & "120±0.1" & vbCrLf & "位置度 Ø0.05 A B" & vbCrLf & "垂直度 0.03 A" & vbCrLf & "Ra 1.6" & vbCrLf & "未注公差按 GB/T 1804-m"
                    Case "dimension_balloon"
                        template.SampleText = "① Ø25±0.02" & vbCrLf & "⬡3" & vbCrLf & "50" & vbCrLf & "+0.02" & vbCrLf & "-0.01" & vbCrLf & "○4 R3" & vbCrLf & "⬢5 位置度 Ø0.05 A B"
                End Select
            End If
        End Sub

        ''' <summary>
        ''' 备用方案：如果JSON文件加载失败，创建最基本的模板
        ''' </summary>
        Private Sub CreateFallbackTemplates()
            ' 发票模板
            If Not _scripts.ContainsKey("invoice") Then
                Dim invoiceTemplate As New ParseScript With {
                    .Id = "invoice",
                    .Name = "增值税发票",
                    .Description = "解析增值税专用/普通发票",
                    .Icon = "🧾",
                    .Category = "票据",
                    .IsBuiltIn = True,
                    .CreatedAt = DateTime.Now,
                    .LastModified = DateTime.Now,
                    .OutputFormat = "table",
                    .SampleText = "发票代码: 142001911510" & vbCrLf & "发票号码: 12345678" & vbCrLf & "开票日期: 2024年01月15日",
                    .Tags = New List(Of String) From {"发票", "税务"},
                    .Fields = New List(Of ParseField) From {
                        New ParseField With {.Name = "发票代码", .Key = "invoice_code", .Pattern = "发票代码[：:\s]*(\d+)", .Required = True},
                        New ParseField With {.Name = "发票号码", .Key = "invoice_number", .Pattern = "发票号码[：:\s]*(\d+)", .Required = True},
                        New ParseField With {.Name = "开票日期", .Key = "issue_date", .Pattern = "开票日期[：:\s]*([\d\-年月日]+)", .Required = True},
                        New ParseField With {.Name = "合计金额", .key = "total_amount", .Pattern = "合计[：:\s]*[¥￥]?([\d,]+\.?\d*)", .Required = False}
                    },
                    .Script = LoadScriptFromResource("invoice.js")
                }
                SaveScript(invoiceTemplate)
            End If

            ' 身份证模板
            If Not _scripts.ContainsKey("idcard") Then
                Dim idCardTemplate As New ParseScript With {
                    .Id = "idcard",
                    .Name = "身份证",
                    .Description = "解析中国居民身份证信息",
                    .Icon = "🆔",
                    .Category = "证件",
                    .IsBuiltIn = True,
                    .CreatedAt = DateTime.Now,
                    .LastModified = DateTime.Now,
                    .OutputFormat = "key-value",
                    .SampleText = "姓名 张三" & vbCrLf & "性别 男" & vbCrLf & "身份证号 11010119900101123X",
                    .Tags = New List(Of String) From {"身份证", "证件"},
                    .Fields = New List(Of ParseField) From {
                        New ParseField With {.Name = "姓名", .Key = "name", .Pattern = "姓名[：:\s]*([^\n]+)", .Required = True},
                        New ParseField With {.Name = "性别", .Key = "gender", .Pattern = "性别[：:\s]*([男女])", .Required = False},
                        New ParseField With {.Name = "身份证号", .Key = "id_number", .Pattern = "(?:公民身份号码|身份证号)[：:\s]*([A-Z0-9]{18}[A-Z0-9xX]?)", .Required = True}
                    },
                    .Script = LoadScriptFromResource("idcard.js")
                }
                SaveScript(idCardTemplate)
            End If
        End Sub

        ''' <summary>
        ''' 从资源文件加载JS脚本（备用方案）
        ''' </summary>
        Private Function LoadScriptFromResource(resourceName As String) As String
            Try
                Dim resourcePath = Path.Combine(Path.GetDirectoryName(_defaultScriptsPath), "Scripts", resourceName)
                If File.Exists(resourcePath) Then
                    Return File.ReadAllText(resourcePath, System.Text.Encoding.UTF8)
                End If
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 加载资源脚本失败: " & ex.Message)
            End Try
            
            ' 返回基础脚本
            Return "function parse(text) { var result = {}; for (var i = 0; i < FIELDS.length; i++) { var field = FIELDS[i]; if (field.pattern) { try { var regex = new RegExp(field.pattern); var match = text.match(regex); if (match && match[1]) { result[field.key] = match[1].trim(); } } catch(e) {} } } return result; }"
        End Function

#Region "公共API"

        ''' <summary>
        ''' 获取所有解析脚本
        ''' </summary>
        Public Function GetAllScripts() As List(Of ParseScript)
            LoadAllScripts()
            SyncLock _lock
                If _scripts Is Nothing Then
                    Return New List(Of ParseScript)()
                End If
                Return _scripts.Values.ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 清空缓存并重新加载内置脚本
        ''' </summary>
        Public Function ClearCache() As Boolean
            Try
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 开始清空缓存...")
                
                SyncLock _lock
                    ' 删除缓存目录中的所有文件
                    If System.IO.Directory.Exists(_scriptsPath) Then
                        For Each file In System.IO.Directory.GetFiles(_scriptsPath)
                            Try
                                System.IO.File.Delete(file)
                                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 已删除: {System.IO.Path.GetFileName(file)}")
                            Catch ex As Exception
                                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 删除失败: {file}, {ex.Message}")
                            End Try
                        Next
                    End If
                    
                    ' 清空内存中的字典
                    _scripts.Clear()
                    
                    ' 重新加载内置脚本
                    EnsureBuiltInTemplates()
                    LoadAllScripts()
                End SyncLock
                
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 缓存已清空，重新加载 {_scripts.Count} 个脚本")
                Return True
            Catch ex As Exception
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 清空缓存失败: {ex.Message}")
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 堆栈: {ex.StackTrace}")
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 根据ID获取脚本
        ''' </summary>
        Public Function GetScriptById(id As String) As ParseScript
            LoadAllScripts()
            SyncLock _lock
                If _scripts.ContainsKey(id) Then
                    Return _scripts(id)
                End If
                Return Nothing
            End SyncLock
        End Function

        ''' <summary>
        ''' 按分类获取脚本
        ''' </summary>
        Public Function GetScriptsByCategory(category As String) As List(Of ParseScript)
            SyncLock _lock
                Return _scripts.Values.Where(Function(s) s.Category = category).ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 保存脚本（新增或更新） - 改为 MD 格式
        ''' </summary>
        Public Function SaveScript(script As ParseScript) As Boolean
            Try
                SyncLock _lock
                    If String.IsNullOrEmpty(script.Id) Then
                        script.Id = Guid.NewGuid().ToString("N").Substring(0, 12)
                    End If
                    
                    script.LastModified = DateTime.Now
                    If script.CreatedAt = Nothing Then
                        script.CreatedAt = DateTime.Now
                    End If
                    
                    _scripts(script.Id) = script
                    
                    ' 保存到 MD 文件
                    Dim markdown = ConvertScriptToMarkdown(script)
                    Dim filePath = Path.Combine(_scriptsPath, script.Id & ".md")
                    Dim utf8NoBom As New System.Text.UTF8Encoding(False)
                    File.WriteAllText(filePath, markdown, utf8NoBom)
                    
                    ' 删除旧的 JSON 文件
                    Dim oldJsonPath = Path.Combine(_scriptsPath, script.Id & ".json")
                    If File.Exists(oldJsonPath) Then
                        Try
                            File.Delete(oldJsonPath)
                        Catch ex As Exception  ' 
                        End Try
                    End If
                    
                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 已保存脚本: " & script.Name & " (" & script.Id & ")")
                    Return True
                End SyncLock
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 保存脚本失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 从JSON字符串保存脚本
        ''' </summary>
        Public Function SaveScriptFromJson(jsonStr As String) As Boolean
            Try
                Dim script = JsonConvert.DeserializeObject(Of ParseScript)(jsonStr)
                If script IsNot Nothing Then
                    Return SaveScript(script)
                End If
                Return False
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[ParseScriptManager] JSON解析失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 删除脚本（内置模板不可删除）
        ''' </summary>
        Public Function DeleteScript(id As String) As Boolean
            Try
                LoadAllScripts()
                SyncLock _lock
                    If Not _scripts.ContainsKey(id) Then Return False
                    
                    Dim script = _scripts(id)
                    If script.IsBuiltIn Then
                        Return False ' 内置模板不可删除
                    End If
                    
                    _scripts.Remove(id)
                    
                    ' 删除文件（.md 格式）
                    Dim filePath = Path.Combine(_scriptsPath, id & ".md")
                    If File.Exists(filePath) Then
                        File.Delete(filePath)
                    End If
                    ' 兼容旧版 .json 文件
                    Dim jsonPath = Path.Combine(_scriptsPath, id & ".json")
                    If File.Exists(jsonPath) Then
                        File.Delete(jsonPath)
                    End If
                    
                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 已删除脚本: " & id)
                    Return True
                End SyncLock
            Catch ex As Exception
                Diagnostics.Debug.WriteLine("[ParseScriptManager] 删除脚本失败: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' 复制脚本
        ''' </summary>
        Public Function DuplicateScript(sourceId As String, newName As String) As ParseScript
            SyncLock _lock
                If Not _scripts.ContainsKey(sourceId) Then
                    Throw New ArgumentException("找不到源脚本: " & sourceId)
                End If
                
                Dim source = _scripts(sourceId)
                Dim json = JsonConvert.SerializeObject(source)
                Dim copy = JsonConvert.DeserializeObject(Of ParseScript)(json)
                
                copy.Id = Guid.NewGuid().ToString("N").Substring(0, 12)
                copy.Name = newName
                copy.IsBuiltIn = False
                copy.CreatedAt = DateTime.Now
                copy.LastModified = DateTime.Now
                copy.UsageCount = 0
                
                _scripts(copy.Id) = copy

                Dim filePath = Path.Combine(_scriptsPath, copy.Id & ".json")
                Dim utf8NoBom As New System.Text.UTF8Encoding(False)
                File.WriteAllText(filePath, JsonConvert.SerializeObject(copy, Formatting.Indented), utf8NoBom)
                
                Return copy
            End SyncLock
        End Function

        ''' <summary>
        ''' 还原内置脚本到默认配置（从 MD 目录重新加载）
        ''' </summary>
        Public Function ResetBuiltinScript(id As String) As ParseScript
            SyncLock _lock
                ' 检查脚本是否存在
                If Not _scripts.ContainsKey(id) Then
                    Return Nothing
                End If

                Dim existing = _scripts(id)
                If Not existing.IsBuiltIn Then
                    Return Nothing ' 非内置脚本不可还原
                End If

                ' 从 MD 文件中查找默认定义
                Dim defaultFile = Path.Combine(_defaultScriptsPath, id & ".md")
                If Not File.Exists(defaultFile) Then
                    Return Nothing
                End If

                Try
                    Dim markdown = File.ReadAllText(defaultFile, System.Text.Encoding.UTF8)
                    Dim defaultTemplate = ParseScriptFromMarkdown(markdown)

                    If defaultTemplate Is Nothing Then
                        Return Nothing
                    End If

                    ' 删除旧的缓存文件
                    Dim filePath = Path.Combine(_scriptsPath, id & ".md")
                    If File.Exists(filePath) Then
                        File.Delete(filePath)
                    End If

                    ' 构建还原后的脚本（保持 isBuiltIn 标记）
                    defaultTemplate.IsBuiltIn = True
                    If defaultTemplate.CreatedAt = Nothing Then
                        defaultTemplate.CreatedAt = DateTime.Now
                    End If
                    defaultTemplate.LastModified = DateTime.Now
                    SetSampleText(defaultTemplate)

                    ' 更新内存中的脚本
                    _scripts(id) = defaultTemplate

                    ' 保存为新的缓存文件
                    SaveScript(defaultTemplate)

                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 已还原内置脚本: " & id)
                    Return defaultTemplate
                Catch ex As Exception
                    Diagnostics.Debug.WriteLine("[ParseScriptManager] 还原内置脚本失败: " & ex.Message)
                    Return Nothing
                End Try
            End SyncLock
        End Function

        ''' <summary>
        ''' 执行解析脚本（返回执行上下文，由前端执行JavaScript）
        ''' </summary>
        Public Function ExecuteScript(scriptId As String, ocrText As String) As String
            Try
                Dim script As ParseScript = Nothing
                SyncLock _lock
                    If _scripts.ContainsKey(scriptId) Then
                        script = _scripts(scriptId)
                    End If
                End SyncLock
                
                If script Is Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                        .success = False,
                        .error = "找不到解析脚本: " & scriptId
                    })
                End If
                
                ' 更新使用次数
                script.UsageCount += 1
                SaveScript(script)
                
                ' 构建执行上下文
                Dim fieldsJson = JsonConvert.SerializeObject(script.Fields)
                
                Return JsonConvert.SerializeObject(New With {
                    .success = True,
                    .scriptId = script.Id,
                    .scriptName = script.Name,
                    .fields = script.Fields,
                    .script = script.Script,
                    .outputFormat = script.OutputFormat,
                    .ocrText = ocrText,
                    .message = "请在前端执行JavaScript解析"
                })
            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
                    .success = False,
                    .error = ex.Message
                })
            End Try
        End Function

        ''' <summary>
        ''' 获取所有分类列表
        ''' </summary>
        Public Function GetCategories() As List(Of String)
            LoadAllScripts()
            SyncLock _lock
                Return _scripts.Values.Select(Function(s) s.Category).Distinct().OrderBy(Function(c) c).ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 搜索脚本（按名称、描述、标签）
        ''' </summary>
        Public Function SearchScripts(keyword As String) As List(Of ParseScript)
            If String.IsNullOrWhiteSpace(keyword) Then
                Return GetAllScripts()
            End If
            LoadAllScripts()
            SyncLock _lock
                If String.IsNullOrWhiteSpace(keyword) Then
                    Return _scripts.Values.ToList()
                End If
                
                keyword = keyword.ToLower()
                Return _scripts.Values.Where(Function(s) s.Name.ToLower().Contains(keyword) OrElse (s.Description IsNot Nothing AndAlso s.Description.ToLower().Contains(keyword)) OrElse (s.Tags IsNot Nothing AndAlso s.Tags.Any(Function(t) t.ToLower().Contains(keyword)))).ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 从 Markdown 文本解析 ParseScript 对象
        ''' </summary>
        Private Function ParseScriptFromMarkdown(markdown As String) As ParseScript
            If String.IsNullOrEmpty(markdown) Then Return Nothing
            
            Diagnostics.Debug.WriteLine($"[ParseMD] ========== 开始解析 ==========")
            Diagnostics.Debug.WriteLine($"[ParseMD] Markdown 长度: {markdown.Length}")
            Diagnostics.Debug.WriteLine($"[ParseMD] 前200字符: {markdown.Substring(0, Math.Min(200, markdown.Length))}")
            
            Dim script As New ParseScript()
            script.Fields = New List(Of ParseField)()
            script.Tags = New List(Of String)()
            script.Id = ""
            script.Name = ""
            script.Description = ""
            script.Icon = "📝"
            script.Category = "通用"
            script.OutputFormat = "table"
            script.Script = ""
            script.SampleText = ""

            Dim lines = markdown.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
            ' 处理可能的 BOM 标记（UTF-8 BOM: EF BB BF）
            If lines.Length > 0 AndAlso lines(0).Length > 0 AndAlso AscW(lines(0)(0)) = &HFEFF Then
                lines(0) = lines(0).Substring(1)
                Diagnostics.Debug.WriteLine("[ParseMD] 检测到 BOM 标记，已移除")
            End If
            Diagnostics.Debug.WriteLine($"[ParseMD] 总行数: {lines.Length}")
            Dim inFrontmatter = False
            Dim frontmatterEnded = False
            Dim inFieldsTable = False
            Dim inScriptBlock = False
            Dim inSampleBlock = False
            Dim scriptLines As New List(Of String)()
            Dim sampleLines As New List(Of String)()
            Dim lineIndex = 0

            For Each line In lines
                ' YAML frontmatter
                If line.Trim() = "---" Then
                    If Not frontmatterEnded AndAlso lineIndex = 0 Then
                        inFrontmatter = True
                        frontmatterEnded = False
                        lineIndex += 1
                        Continue For
                    ElseIf inFrontmatter Then
                        inFrontmatter = False
                        frontmatterEnded = True
                        lineIndex += 1
                        Continue For
                    End If
                End If

                If inFrontmatter Then
                    Dim parts = line.Split(New Char() {":"c}, 2)
                    If parts.Length = 2 Then
                        Dim key = parts(0).Trim().ToLower()
                        Dim value = parts(1).Trim().Trim(""""c).Trim("'"c)
                        Select Case key
                            Case "id"
                                script.Id = value
                            Case "name"
                                script.Name = value
                            Case "description"
                                script.Description = value
                            Case "icon"
                                script.Icon = value
                            Case "category"
                                script.Category = value
                            Case "outputformat"
                                script.OutputFormat = value
                            Case "tags"
                                If Not String.IsNullOrEmpty(value) Then
                                    script.Tags = value.Split(","c).Select(Function(t) t.Trim()).Where(Function(t) Not String.IsNullOrEmpty(t)).ToList()
                                Else
                                    script.Tags = New List(Of String)()
                                End If
                        End Select
                    End If
                    Continue For
                End If

                ' 检测字段表格
                If line.Trim().StartsWith("|") AndAlso line.Contains("字段名") AndAlso line.Contains("键名") Then
                    Diagnostics.Debug.WriteLine($"[ParseMD] ✅ 发现字段表格表头: {line.Trim()}")
                    inFieldsTable = True
                    Continue For
                End If

                If inFieldsTable Then
                    If Not line.Trim().StartsWith("|") Then
                        Diagnostics.Debug.WriteLine($"[ParseMD] ⚠️ 字段表格结束（遇到非|行）, 共解析 {script.Fields.Count} 个字段")
                        inFieldsTable = False
                    Else
                        ' 智能解析表格行，处理反引号内的 |
                        Dim cells = ParseMarkdownTableCells(line)
                        Diagnostics.Debug.WriteLine($"[ParseMD] 📋 表格行: {line.Trim()}")
                        Diagnostics.Debug.WriteLine($"[ParseMD] 🔍 解析到 {cells.Length} 个单元格")
                        For ci As Integer = 0 To cells.Length - 1
                            Diagnostics.Debug.WriteLine($"[ParseMD]   单元格[{ci}]: [{cells(ci)}]")
                        Next
                        If cells.Length >= 4 AndAlso Not cells(0).Contains("---") Then
                            Diagnostics.Debug.WriteLine($"[ParseMD] ✅ 添加字段：{cells(0)} - {cells(1)}")
                            Dim field As New ParseField()
                            field.Name = cells(0)
                            field.Key = cells(1)
                            ' 提取正则表达式（去除前后反引号）
                            field.Pattern = cells(2).Trim("`"c)
                            field.Required = cells(3).Contains("是") OrElse cells(3).Contains("true")
                            ' 解析可见字段（第5列，可选）
                            If cells.Length >= 5 Then
                                field.Visible = cells(4).Contains("是") OrElse cells(4).Contains("true") OrElse Not cells(4).Contains("否")
                            Else
                                field.Visible = True
                            End If
                            script.Fields.Add(field)
                        Else
                            Diagnostics.Debug.WriteLine($"[ParseMD] 跳过此行: cells.Length={cells.Length}")
                        End If
                    End If
                    Continue For
                End If

                ' ===== 代码块检测区域（表格和 frontmatter 之后的行都会到这里）=====
                Dim trimmedLine = line.Trim()

                ' 调试：打印每个到达这里的非空行（前100行）
                If lineIndex < 100 AndAlso Not String.IsNullOrWhiteSpace(trimmedLine) Then
                    Diagnostics.Debug.WriteLine($"[ParseMD] 行{lineIndex}: [{trimmedLine.Substring(0, Math.Min(60, trimmedLine.Length))}] inScript={inScriptBlock} inSample={inSampleBlock}")
                End If

                ' 检测代码块（支持 3+ 个反引号）
                If trimmedLine.StartsWith("```") Then
                    If inScriptBlock Then
                        ' JavaScript代码块结束
                        inScriptBlock = False
                        script.Script = String.Join(vbCrLf, scriptLines).Trim()
                        scriptLines.Clear()
                        Diagnostics.Debug.WriteLine("[ParseMD] ✅ JavaScript代码块结束")
                        Continue For
                    ElseIf inSampleBlock Then
                        ' 示例文本代码块结束
                        inSampleBlock = False
                        script.SampleText = String.Join(vbCrLf, sampleLines).Trim()
                        sampleLines.Clear()
                        Diagnostics.Debug.WriteLine("[ParseMD] ✅ 示例文本代码块结束")
                        Continue For
                    Else
                        ' 检测到新代码块开始，检查是否有语言标记
                        Dim langMarker = trimmedLine.ToLower()
                        If langMarker.Contains("javascript") OrElse langMarker.Contains("js") Then
                            ' JavaScript 代码块
                            inScriptBlock = True
                            scriptLines.Clear()
                            Diagnostics.Debug.WriteLine($"[ParseMD] 📝 JavaScript代码块开始: {trimmedLine}")
                        ElseIf langMarker.Contains("sample") OrElse langMarker.Contains("example") OrElse langMarker.Contains("text") _
                            OrElse langMarker.Contains("plaintext") OrElse langMarker.Contains("markdown") _
                            OrElse langMarker.Contains("示例") Then
                            ' 明确标记为示例文本的代码块
                            inSampleBlock = True
                            sampleLines.Clear()
                            Diagnostics.Debug.WriteLine($"[ParseMD] 📋 示例文本代码块开始(显式标记): {trimmedLine}")
                        ElseIf trimmedLine = "```" OrElse trimmedLine = "````" OrElse trimmedLine = "`````" Then
                            ' 无语言标记的代码块 → 根据上下文判断：
                            ' 如果已有脚本代码且没有示例文本，则视为示例文本；否则视为脚本
                            If Not String.IsNullOrEmpty(script.Script) AndAlso String.IsNullOrEmpty(script.SampleText) Then
                                inSampleBlock = True
                                sampleLines.Clear()
                                Diagnostics.Debug.WriteLine("[ParseMD] 📋 示例文本代码块开始(无标记,在脚本后)")
                            Else
                                inScriptBlock = True
                                scriptLines.Clear()
                                Diagnostics.Debug.WriteLine("[ParseMD] 📝 JavaScript代码块开始(无标记,默认脚本)")
                            End If
                        Else
                            ' 其他语言标记（如 python、json 等），视为示例文本
                            inSampleBlock = True
                            sampleLines.Clear()
                            Diagnostics.Debug.WriteLine($"[ParseMD] 📋 示例文本代码块开始(其他标记): {trimmedLine}")
                        End If
                        Continue For
                    End If
                End If

                If inScriptBlock Then
                    scriptLines.Add(line)
                    Continue For
                End If

                If inSampleBlock Then
                    ' 跳过空行
                    If Not String.IsNullOrWhiteSpace(line) Then
                        sampleLines.Add(line)
                    End If
                    Continue For
                End If
                
                lineIndex += 1
            Next

            ' 组装脚本（如果代码块没有正常关闭，在这里处理）
            If scriptLines.Count > 0 AndAlso String.IsNullOrEmpty(script.Script) Then
                script.Script = String.Join(vbCrLf, scriptLines).Trim()
            End If
            If sampleLines.Count > 0 AndAlso String.IsNullOrEmpty(script.SampleText) Then
                script.SampleText = String.Join(vbCrLf, sampleLines).Trim()
            End If

            ' 调试日志：输出解析结果
            Diagnostics.Debug.WriteLine($"[ParseMD] 解析完成: name={script.Name}, id={script.Id}")
            Diagnostics.Debug.WriteLine($"[ParseMD] Script长度: {If(script.Script, "").Length}, SampleText长度: {If(script.SampleText, "").Length}")
            Diagnostics.Debug.WriteLine($"[ParseMD] Fields数量: {If(script.Fields, New List(Of ParseField)()).Count}")
            If Not String.IsNullOrEmpty(script.Script) Then
                Diagnostics.Debug.WriteLine($"[ParseMD] Script前100字符: {script.Script.Substring(0, Math.Min(100, script.Script.Length))}")
            Else
                Diagnostics.Debug.WriteLine("[ParseMD] ⚠️ Script 为空！未检测到 ```javascript 代码块")
            End If
            If Not String.IsNullOrEmpty(script.SampleText) Then
                Diagnostics.Debug.WriteLine($"[ParseMD] SampleText前100字符: {script.SampleText.Substring(0, Math.Min(100, script.SampleText.Length))}")
            Else
                Diagnostics.Debug.WriteLine("[ParseMD] ⚠️ SampleText 为空！未检测到示例文本代码块")
            End If

            Return script
        End Function

        ''' <summary>
        ''' 智能解析 Markdown 表格单元格，正确处理反引号内的分隔符
        ''' </summary>
        Private Function ParseMarkdownTableCells(line As String) As String()
            Dim cells As New List(Of String)()
            Dim currentCell As New System.Text.StringBuilder()
            Dim inBacktick = False
            Dim trimmedLine = line.Trim()
            
            ' 去除首尾的 |
            If trimmedLine.StartsWith("|") Then trimmedLine = trimmedLine.Substring(1)
            If trimmedLine.EndsWith("|") Then trimmedLine = trimmedLine.Substring(0, trimmedLine.Length - 1)
            
            For i As Integer = 0 To trimmedLine.Length - 1
                Dim ch = trimmedLine(i)
                
                If ch = "`"c Then
                    inBacktick = Not inBacktick
                    currentCell.Append(ch)
                ElseIf ch = "|"c AndAlso Not inBacktick Then
                    ' 遇到表格分隔符且不在反引号内
                    cells.Add(currentCell.ToString().Trim())
                    currentCell.Clear()
                Else
                    currentCell.Append(ch)
                End If
            Next
            
            ' 添加最后一个单元格
            If currentCell.Length > 0 Then
                cells.Add(currentCell.ToString().Trim())
            End If
            
            Return cells.ToArray()
        End Function

        ''' <summary>
        ''' 将 ParseScript 转换为 Markdown 格式
        ''' </summary>
        Private Function ConvertScriptToMarkdown(script As ParseScript) As String
            Dim md As New System.Text.StringBuilder()

            ' YAML frontmatter
            md.AppendLine("---")
            md.AppendLine($"id: {script.Id}")
            md.AppendLine($"name: {script.Name}")
            md.AppendLine($"description: {script.Description}")
            md.AppendLine($"icon: {script.Icon}")
            md.AppendLine($"category: {script.Category}")
            md.AppendLine($"outputFormat: {script.OutputFormat}")
            If script.Tags IsNot Nothing AndAlso script.Tags.Count > 0 Then
                md.AppendLine($"tags: {String.Join(", ", script.Tags)}")
            End If
            md.AppendLine("---")
            md.AppendLine()

            ' 字段定义表格
            md.AppendLine("## 字段定义")
            md.AppendLine()
            md.AppendLine("| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |")
            md.AppendLine("|--------|------|-----------|------|------|")
            If script.Fields IsNot Nothing Then
                For Each field In script.Fields
                    Dim required = If(field.Required, "是", "否")
                    Dim visible = If(field.Visible, "是", "否")
                    Dim pattern = If(String.IsNullOrEmpty(field.Pattern), "-", $"`{field.Pattern}`")
                    md.AppendLine($"| {field.Name} | {field.Key} | {pattern} | {required} | {visible} |")
                Next
            End If
            md.AppendLine()

            ' JavaScript 代码
            md.AppendLine("## 解析脚本")
            md.AppendLine()
            md.AppendLine("```javascript")
            md.AppendLine(script.Script)
            md.AppendLine("```")
            md.AppendLine()

            ' 示例文本
            If Not String.IsNullOrEmpty(script.SampleText) Then
                md.AppendLine("## 示例文本")
                md.AppendLine()
                md.AppendLine("```")
                md.AppendLine(script.SampleText)
                md.AppendLine("```")
            End If

            Return md.ToString()
        End Function

        ''' <summary>
        ''' 将脚本保存为 Markdown 格式（辅助方法）
        ''' </summary>
        Private Sub SaveScriptAsMarkdown(script As ParseScript)
            Dim markdown = ConvertScriptToMarkdown(script)
            Dim filePath = Path.Combine(_scriptsPath, script.Id & ".md")
            ' 使用 UTF-8 with BOM 确保 Windows 系统正确识别
            File.WriteAllText(filePath, markdown, System.Text.Encoding.UTF8)
        End Sub

        ''' <summary>
        ''' 从 Markdown 文件内容导入解析脚本
        ''' </summary>
        Public Function ImportScriptFromMarkdown(markdownContent As String) As ParseScript
            Try
                If String.IsNullOrWhiteSpace(markdownContent) Then
                    Return Nothing
                End If

                ' 解析 Markdown 内容为 ParseScript 对象
                Dim script = ParseScriptFromMarkdown(markdownContent)
                
                If script Is Nothing OrElse String.IsNullOrEmpty(script.Id) Then
                    Return Nothing
                End If

                ' 生成新 ID（避免冲突）
                script.Id = Guid.NewGuid().ToString("N").Substring(0, 12)
                script.IsBuiltIn = False
                script.CreatedAt = DateTime.Now
                script.LastModified = DateTime.Now

                ' 保存到文件
                SyncLock _lock
                    _scripts(script.Id) = script
                    Dim markdown = ConvertScriptToMarkdown(script)
                    Dim filePath = Path.Combine(_scriptsPath, script.Id & ".md")
                    Dim utf8NoBom As New System.Text.UTF8Encoding(False)
                    File.WriteAllText(filePath, markdown, utf8NoBom)
                End SyncLock

                Return script
            Catch ex As Exception
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 导入脚本失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 【Skill 自动导入】把 Skill 生成的解析脚本写入脚本库：
        ''' 同名（非内置）脚本直接覆盖，避免每次重新生成都多出一张重复卡片。
        ''' </summary>
        ''' <param name="markdownContent">Skill 生成的解析脚本 Markdown</param>
        ''' <param name="replacedExisting">True 表示覆盖了同名旧脚本</param>
        Public Function ImportScriptFromMarkdownOrReplace(markdownContent As String, ByRef replacedExisting As Boolean) As ParseScript
            replacedExisting = False
            Try
                Dim md = ExtractParseScriptSection(markdownContent)
                If String.IsNullOrWhiteSpace(md) Then Return Nothing

                Dim script = ParseScriptFromMarkdown(md)
                If script Is Nothing Then Return Nothing
                If String.IsNullOrWhiteSpace(script.Name) Then Return Nothing
                If String.IsNullOrWhiteSpace(script.Script) Then Return Nothing
                If script.Fields Is Nothing OrElse script.Fields.Count = 0 Then Return Nothing

                ' frontmatter 漏写 id 时用名称兜底：不应因为 LLM 的小疏漏让「生成即入库」失败
                If String.IsNullOrWhiteSpace(script.Id) Then
                    script.Id = "auto_" & Guid.NewGuid().ToString("N").Substring(0, 8)
                End If
                script.IsBuiltIn = False

                SyncLock _lock
                    Dim existing As ParseScript = Nothing
                    For Each kvp In _scripts
                        Dim s = kvp.Value
                        If s IsNot Nothing AndAlso Not s.IsBuiltIn AndAlso
                           Not String.IsNullOrWhiteSpace(s.Name) AndAlso
                           String.Equals(s.Name.Trim(), script.Name.Trim(), StringComparison.OrdinalIgnoreCase) Then
                            existing = s
                            Exit For
                        End If
                    Next

                    If existing IsNot Nothing Then
                        ' 覆盖：保留原 Id/创建时间/使用统计，让面板里的卡片原地更新
                        script.Id = existing.Id
                        script.CreatedAt = existing.CreatedAt
                        script.UsageCount = existing.UsageCount
                        script.Version = existing.Version
                        replacedExisting = True
                    Else
                        script.CreatedAt = DateTime.Now
                    End If
                    script.LastModified = DateTime.Now

                    _scripts(script.Id) = script
                    Dim markdown = ConvertScriptToMarkdown(script)
                    Dim filePath = Path.Combine(_scriptsPath, script.Id & ".md")
                    Dim utf8NoBom As New System.Text.UTF8Encoding(False)
                    File.WriteAllText(filePath, markdown, utf8NoBom)
                End SyncLock

                Return script
            Catch ex As Exception
                Diagnostics.Debug.WriteLine($"[ParseScriptManager] 自动导入解析脚本失败: {ex.Message}")
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' 从 Skill 输出中提取解析脚本 Markdown：
        ''' 1) 去掉 LLM 习惯性加的外层 ```markdown 围栏；2) 砍掉 frontmatter 之前的客套话。
        ''' （解析器要求首行就是 ---，否则整份脚本会被判为无效 → 用户表现为「生成的脚本用不了」）
        ''' </summary>
        Public Shared Function ExtractParseScriptSection(rawContent As String) As String
            If String.IsNullOrWhiteSpace(rawContent) Then Return String.Empty

            Dim text = rawContent.Replace(vbCrLf, vbLf).Trim()

            ' 1) 整份脚本被 ```markdown ... ``` 包住的情况
            Dim fence = System.Text.RegularExpressions.Regex.Match(text, "^```[a-zA-Z]*\s*\n([\s\S]*?)\n?```\s*$")
            If fence.Success AndAlso fence.Groups(1).Value.Contains("---") Then
                text = fence.Groups(1).Value.Trim()
            End If

            ' 2) 砍掉 frontmatter 之前的说明文字（"好的，脚本如下：" 之类）
            Dim idx = text.IndexOf("---" & vbLf)
            If idx > 0 Then text = text.Substring(idx).Trim()

            If Not text.StartsWith("---") Then Return String.Empty
            Return text
        End Function

        ''' <summary>
        ''' 从 WorkflowResult.Outputs 里取 Skill 的最终文本输出（兼容大小写键名）
        ''' </summary>
        Public Shared Function GetSkillOutputText(outputs As IDictionary(Of String, Object)) As String
            If outputs Is Nothing Then Return String.Empty
            For Each key In New String() {"content", "Content", "markdown", "Markdown"}
                Dim value As Object = Nothing
                If outputs.TryGetValue(key, value) AndAlso value IsNot Nothing Then
                    Dim s = value.ToString()
                    If Not String.IsNullOrWhiteSpace(s) Then Return s
                End If
            Next
            Return String.Empty
        End Function

        ''' <summary>
        ''' 【Skill 自动导入】Skill 输出 → OCR 智能解析面板脚本库的统一入口。
        ''' 由 Skill 执行链路在 Skill 声明 autoImportParseScript: true 时调用（如 GenerateOCRParseScript）。
        ''' 结果写入 result.Outputs("parseScriptImport") 供前端提示，并推送给已打开的 OCR 面板刷新列表。
        ''' </summary>
        Public Shared Sub AutoImportSkillParseScript(skill As Skills.SkillDefinition,
                                                     result As Skills.WorkflowResult)
            If result Is Nothing OrElse result.Outputs Is Nothing Then Return

            Dim outcome As New ParseScriptImportResult()
            Try
                Dim content = GetSkillOutputText(result.Outputs)
                If String.IsNullOrWhiteSpace(content) Then
                    outcome.ErrorMessage = "Skill 输出为空"
                Else
                    Dim replaced As Boolean = False
                    Dim script = Instance.ImportScriptFromMarkdownOrReplace(content, replaced)
                    If script Is Nothing Then
                        outcome.ErrorMessage = "输出不是可导入的解析脚本（需 YAML frontmatter + 字段定义表格 + ```javascript 解析脚本）"
                    Else
                        outcome.Success = True
                        outcome.ScriptId = script.Id
                        outcome.ScriptName = script.Name
                        outcome.Replaced = replaced
                    End If
                End If
            Catch ex As Exception
                outcome.ErrorMessage = ex.Message
            End Try

            result.Outputs("parseScriptImport") = outcome.ToOutput()

            If outcome.Success Then
                Utils.AppLogger.LogInfo("SKILL_OCR_IMPORT",
                    $"✓ Skill '{skill?.Name}' 生成的解析脚本已导入 OCR 智能解析面板: 「{outcome.ScriptName}」(覆盖同名={outcome.Replaced}, id={outcome.ScriptId})")
#If Not SMARTPLOT_CORE Then
                ' 【Excel 专属】TaskPaneManager 依赖 Microsoft.Office.Tools / ExcelBridge，
                ' 无法编入与 Excel 解耦的公共核心库 Core；本段只是"通知 OCR 面板刷新"的次要功能，
                ' 且外层已有 Try/Catch 兜底，故在 Core 编译时跳过（SMARTPLOT_CORE）。
                Try
                    Dim paneManager = TaskPaneManager.Current
                    If paneManager IsNot Nothing Then
                        paneManager.NotifyParseScriptUpdated(outcome.ScriptId, outcome.ScriptName, outcome.Replaced)
                    End If
                Catch ex As Exception
                    Diagnostics.Debug.WriteLine($"[AutoImportSkillParseScript] 通知 OCR 面板失败: {ex.Message}")
                End Try
#End If
            Else
                Utils.AppLogger.LogWarning("SKILL_OCR_IMPORT",
                    $"Skill '{skill?.Name}' 声明了 autoImportParseScript，但自动导入未完成: {outcome.ErrorMessage}")
            End If
        End Sub

        ''' <summary>Skill 自动导入解析脚本的执行结果（转成前端可读的 outputs 字典）</summary>
        Public Class ParseScriptImportResult
            Public Property Success As Boolean
            Public Property ScriptId As String = ""
            Public Property ScriptName As String = ""
            Public Property Replaced As Boolean
            Public Property ErrorMessage As String = ""

            Public Function ToOutput() As Dictionary(Of String, Object)
                Dim dict As New Dictionary(Of String, Object)()
                dict("success") = Success
                dict("replaced") = Replaced
                If Not String.IsNullOrEmpty(ScriptId) Then dict("scriptId") = ScriptId
                If Not String.IsNullOrEmpty(ScriptName) Then dict("scriptName") = ScriptName
                If Not String.IsNullOrEmpty(ErrorMessage) Then dict("error") = ErrorMessage
                Return dict
            End Function
        End Class

#End Region

    End Class

    ''' <summary>
    ''' 默认脚本容器类（用于反序列化JSON模板文件）
    ''' </summary>
    Public Class DefaultScriptsContainer
        Public Property Scripts As List(Of ParseScript)
    End Class

End Namespace