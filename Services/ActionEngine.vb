Option Strict On
Option Infer On
Imports System.Threading
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Excel = Microsoft.Office.Interop.Excel
Imports ExcelAIVSTOAddIn
Imports ExcelAIVSTOAddIn.Bridges

Namespace Services

    ''' <summary>
    ''' 操作结果
    ''' </summary>
    Public Class OperationResult
        Public Property Success As Boolean
        Public Property Message As String
        Public Property Data As Object
        Public Property ErrorMessage As String
        ''' <summary>P3：动作类型（excel_operation/code_generation/...）</summary>
        Public Property ActionType As String
        ''' <summary>P3：原始动作描述（显示哪个动作失败）</summary>
        Public Property ActionDescription As String
        ''' <summary>P3：异常分类（Transient/Permanent/UserError/SystemCrash）</summary>
        Public Property ErrorCategory As String
        ''' <summary>P3：诊断元数据（retry_count/stack_trace/ui_thread_ms）</summary>
        Public Property Metadata As Dictionary(Of String, Object)

        Public Sub New()
            Metadata = New Dictionary(Of String, Object)()
        End Sub

        Public Shared Function CreateSuccess(message As String, Optional data As Object = Nothing) As OperationResult
            Return New OperationResult With {
                .Success = True,
                .Message = message,
                .Data = data
            }
        End Function

        Public Shared Function CreateFailure(errorMsg As String) As OperationResult
            Return New OperationResult With {
                .Success = False,
                .ErrorMessage = errorMsg,
                .Message = errorMsg
            }
        End Function

        ''' <summary>P3：带异常分类+上下文的失败构造器（对齐Skill Router分类）</summary>
        Public Shared Function CreateFailureEx(errorMsg As String, ex As Exception, Optional action As ActionItem = Nothing) As OperationResult
            Dim category As String = "Unknown"
            If ex IsNot Nothing Then
                If TypeOf ex Is ArgumentException OrElse TypeOf ex Is InvalidOperationException Then
                    category = "UserError"
                ElseIf TypeOf ex Is System.Runtime.InteropServices.COMException OrElse TypeOf ex Is IO.IOException OrElse TypeOf ex Is TimeoutException Then
                    category = "Transient"
                ElseIf TypeOf ex Is OutOfMemoryException OrElse TypeOf ex Is Runtime.InteropServices.SEHException Then
                    category = "SystemCrash"
                Else
                    category = "Permanent"
                End If
            End If
            Dim res = New OperationResult With {
                .Success = False,
                .ErrorMessage = errorMsg,
                .Message = errorMsg,
                .ErrorCategory = category
            }
            If action IsNot Nothing Then
                res.ActionType = action.ActionType
                res.ActionDescription = action.Description
            End If
            If ex IsNot Nothing Then
                res.Metadata("stack_trace") = ex.StackTrace
                res.Metadata("exception_type") = ex.GetType().Name
            End If
            Return res
        End Function
    End Class

    ''' <summary>
    ''' 执行策略枚举：决定任务走 VSTO 直接操作还是 VBA 代码生成
    ''' </summary>
    Public Enum ExecutionStrategy
        DirectVstoOperation    ' 优先走 ActionEngine 内置 Excel 操作
        GenerateAndRunVBA      ' 生成 VBA 并立即执行
        GenerateVBAOnly        ' 仅生成 VBA 代码（给用户保存为宏）
        Unknown                ' 无法判断，交给默认逻辑
    End Enum

    ''' <summary>
    ''' 动作执行引擎 - 执行AI识别出的各种动作
    ''' </summary>
    Public Class ActionEngine
        Private _excelBridge As ExcelAIVSTOAddIn.ExcelBridge

        Public Sub New()
        End Sub

        Public Sub New(bridge As ExcelAIVSTOAddIn.ExcelBridge)
            _excelBridge = bridge
        End Sub

        Private Shared ReadOnly _instance As New Lazy(Of ActionEngine)(Function() New ActionEngine(), System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)

        Public Shared ReadOnly Property Instance As ActionEngine
            Get
                Return _instance.Value
            End Get
        End Property

        ''' <summary>
        ''' 注入 ExcelBridge 引用（用于 MCP 工具调用）
        ''' </summary>
        Public Sub Initialize(bridge As ExcelAIVSTOAddIn.ExcelBridge)
            _excelBridge = bridge
        End Sub

        ''' <summary>
        ''' 执行动作列表
        ''' P3 增强：
        '''   ① 每个动作独立Try/Catch隔离（单个失败不影响已完成结果）
        '''   ② 每个结果带ActionType/ActionDescription，可直接展示谁失败
        '''   ③ Cancellation支持（中途取消可返回已成功部分+取消标记）
        '''   ④ 失败使用CreateFailureEx对齐Skill Router的错误分类
        ''' </summary>
        Public Async Function ExecuteActions(actions As List(Of ActionItem), Optional ct As CancellationToken = Nothing) As Task(Of List(Of OperationResult))
            Dim results = New List(Of OperationResult)()
            If actions Is Nothing Then Return results
            Dim total = actions.Count
            Dim idx = 0

            For Each action In actions
                idx += 1
                If ct.IsCancellationRequested Then
                    Dim cancelRes = OperationResult.CreateFailure($"动作已取消（{idx}/{total}），已返回前{idx - 1}条结果")
                    cancelRes.ErrorCategory = "Transient"
                    cancelRes.ActionType = If(action?.ActionType, "unknown")
                    cancelRes.ActionDescription = If(action?.Description, "")
                    cancelRes.Metadata("cancelled") = True
                    cancelRes.Metadata("progress") = $"{idx - 1}/{total}"
                    results.Add(cancelRes)
                    Exit For
                End If

                Dim sw = System.Diagnostics.Stopwatch.StartNew()
                Dim result As OperationResult = Nothing
                Try
                    result = Await ExecuteSingleAction(action)
                Catch ex As Exception
                    result = OperationResult.CreateFailureEx($"动作#{idx}执行异常: {ex.Message}", ex, action)
                End Try

                sw.Stop()
                If result IsNot Nothing Then
                    If String.IsNullOrEmpty(result.ActionType) AndAlso action IsNot Nothing Then
                        result.ActionType = action.ActionType
                    End If
                    If String.IsNullOrEmpty(result.ActionDescription) AndAlso action IsNot Nothing Then
                        result.ActionDescription = action.Description
                    End If
                    result.Metadata("index") = idx
                    result.Metadata("total") = total
                    result.Metadata("elapsed_ms") = sw.ElapsedMilliseconds
                    results.Add(result)
                End If
            Next

            Return results
        End Function

        ''' <summary>
        ''' 在 Excel 主线程（UI 线程）上执行同步操作。
        ''' Excel COM 对象必须在创建它们的 STA 线程上访问。
        ''' </summary>
        Private Shared Function ExecuteOnUIThread(Of T)(func As Func(Of T)) As T
            Dim result As T = Nothing
            ThisAddIn.RunOnUIThread(Sub() result = func())
            Return result
        End Function

        ''' <summary>
        ''' 执行单个动作
        ''' </summary>
        Public Async Function ExecuteSingleAction(action As ActionItem) As Task(Of OperationResult)
            Try
                ' 【优化】对 code_generation 动作进行策略路由：
                ' 如果描述明显可用内置 VSTO 操作完成，则优先走 DirectVstoOperation，
                ' 避免把简单任务都绕去生成 VBA。
                If action.ActionType.ToLower() = "code_generation" Then
                    Dim strategy = DetermineStrategy(action.Description)
                    If strategy = ExecutionStrategy.DirectVstoOperation Then
                        Dim routedAction = TryRouteToExcelOperation(action)
                        If routedAction IsNot Nothing Then
                            Return Await ExecuteExcelOperation(routedAction)
                        End If
                    ElseIf strategy = ExecutionStrategy.GenerateVBAOnly Then
                        ' 仅生成代码不执行：当前直接返回提示，未来可接代码展示 UI
                        Return OperationResult.CreateSuccess(
                            "该需求建议生成可复用的 VBA 宏代码，请使用生成代码模式。",
                            New With {.strategy = "GenerateVBAOnly"})
                    End If
                End If

                Select Case action.ActionType.ToLower()
                    Case "excel_operation"
                        Return Await ExecuteExcelOperation(action)

                    Case "code_generation"
                        Return Await GenerateAndExecuteCode(action)

                    Case "data_extraction"
                        Return Await ExtractAndFillData(action)

                    Case "image_recognition"
                        Return Await RecognizeImage(action)

                    Case "mcp_tool_call"
                        Return Await ExecuteMcpToolCall(action)

                    Case Else
                        Return OperationResult.CreateFailure($"未知的动作类型: {action.ActionType}")
                End Select

            Catch ex As Exception
                Return OperationResult.CreateFailureEx($"执行动作失败: {ex.Message}", ex, action)
            End Try
        End Function

        ''' <summary>
        ''' 根据动作描述判断执行策略
        ''' </summary>
        Public Function DetermineStrategy(description As String) As ExecutionStrategy
            If String.IsNullOrWhiteSpace(description) Then Return ExecutionStrategy.Unknown

            Dim lower = description.ToLower()

            ' 明确要“代码/宏/VBA/脚本”的，走仅生成路线（不自动执行）
            If lower.Contains("宏") OrElse lower.Contains("vba") OrElse lower.Contains("代码") OrElse
               lower.Contains("脚本") OrElse lower.Contains("保存为") OrElse lower.Contains("导出代码") OrElse
               lower.Contains("生成代码") OrElse lower.Contains("写代码") OrElse lower.Contains("写宏") Then
                Return ExecutionStrategy.GenerateVBAOnly
            End If

            ' 常见可直接内置操作完成的意图
            Dim directOperations = New String() {
                "创建工作表", "新建工作表", "新建sheet", "create sheet", "create worksheet",
                "填充数据", "写入数据", "fill data", "写入单元格",
                "格式化", "设置格式", "format cells", "设置字体", "设置颜色",
                "合并工作表", "合并sheet", "merge sheets", "合并数据",
                "拆分工作簿", "拆分sheet", "split workbook",
                "读取数据", "获取数据", "read data", "导出数据"
            }

            If directOperations.Any(Function(k) lower.Contains(k.ToLower())) Then
                Return ExecutionStrategy.DirectVstoOperation
            End If

            Return ExecutionStrategy.Unknown
        End Function

        ''' <summary>
        ''' 尝试把 code_generation 描述路由为 excel_operation 动作
        ''' 仅处理能简单解析参数的常见场景，复杂场景返回 Nothing 保持原流程
        ''' </summary>
        Private Function TryRouteToExcelOperation(action As ActionItem) As ActionItem
            If action Is Nothing OrElse String.IsNullOrWhiteSpace(action.Description) Then Return Nothing

            Dim lower = action.Description.ToLower()
            Dim routed = New ActionItem With {
                .ActionType = "excel_operation",
                .Description = action.Description,
                .Confidence = action.Confidence,
                .Parameters = New Dictionary(Of String, Object)()
            }

            ' 创建/新建工作表：提取工作表名
            If lower.Contains("创建") OrElse lower.Contains("新建") OrElse lower.Contains("create") Then
                If lower.Contains("工作表") OrElse lower.Contains("sheet") OrElse lower.Contains("worksheet") Then
                    routed.Parameters("operation") = "create_sheet"
                    ' 简单提取引号或“名为”后的名字，无法提取时留空让创建方法自动命名
                    Dim sheetName = ExtractSheetName(action.Description)
                    If Not String.IsNullOrEmpty(sheetName) Then
                        routed.Parameters("sheetName") = sheetName
                    End If
                    Return routed
                End If
            End If

            ' 格式化/设置格式
            If lower.Contains("格式") OrElse lower.Contains("format") OrElse lower.Contains("字体") OrElse lower.Contains("颜色") Then
                routed.Parameters("operation") = "format_cells"
                If action.Parameters.ContainsKey("range") Then
                    routed.Parameters("range") = action.Parameters("range")
                End If
                If action.Parameters.ContainsKey("format") Then
                    routed.Parameters("format") = action.Parameters("format")
                End If
                Return routed
            End If

            ' 合并工作表
            If (lower.Contains("合并") OrElse lower.Contains("merge")) AndAlso
               (lower.Contains("工作表") OrElse lower.Contains("sheet")) Then
                routed.Parameters("operation") = "merge_sheets"
                If action.Parameters.ContainsKey("sourceSheets") Then
                    routed.Parameters("sourceSheets") = action.Parameters("sourceSheets")
                End If
                If action.Parameters.ContainsKey("targetSheet") Then
                    routed.Parameters("targetSheet") = action.Parameters("targetSheet")
                End If
                Return routed
            End If

            ' 拆分工作簿
            If (lower.Contains("拆分") OrElse lower.Contains("split")) AndAlso
               (lower.Contains("工作簿") OrElse lower.Contains("workbook") OrElse lower.Contains("sheet")) Then
                routed.Parameters("operation") = "split_workbook"
                If action.Parameters.ContainsKey("splitMethod") Then
                    routed.Parameters("splitMethod") = action.Parameters("splitMethod")
                End If
                Return routed
            End If

            ' 读取/获取数据
            If lower.Contains("读取") OrElse lower.Contains("获取") OrElse lower.Contains("read") OrElse lower.Contains("导出数据") Then
                routed.Parameters("operation") = "read_data"
                If action.Parameters.ContainsKey("sheetName") Then
                    routed.Parameters("sheetName") = action.Parameters("sheetName")
                End If
                If action.Parameters.ContainsKey("range") Then
                    routed.Parameters("range") = action.Parameters("range")
                End If
                Return routed
            End If

            ' 默认不路由
            Return Nothing
        End Function

        ''' <summary>
        ''' 从描述中简单提取工作表名
        ''' </summary>
        Private Function ExtractSheetName(description As String) As String
            ' 匹配 "名为 XXX" / "name XXX" / 引号内容
            Dim patterns = New String() {
                "名为[""']?([^""'\s]+)",
                "name[""']?([^""'\s]+)",
                "[""']([^""']+)[""']"
            }
            For Each pattern In patterns
                Dim match = System.Text.RegularExpressions.Regex.Match(description, pattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                If match.Success AndAlso match.Groups.Count > 1 Then
                    Return match.Groups(1).Value.Trim()
                End If
            Next
            Return String.Empty
        End Function

        ''' <summary>
        ''' 执行Excel操作（异步包装，实际在 Excel 主线程同步执行）
        ''' </summary>
        Private Function ExecuteExcelOperation(action As ActionItem) As Task(Of OperationResult)
            Return Task.FromResult(ExecuteOnUIThread(Function() ExecuteExcelOperationCore(action)))
        End Function

        ''' <summary>
        ''' 执行Excel操作核心逻辑（必须在 Excel 主线程调用）
        ''' </summary>
        Private Function ExecuteExcelOperationCore(action As ActionItem) As OperationResult
            Try
                Dim validation = ValidateExcelOperationParameters(action)
                If validation IsNot Nothing Then Return validation

                Dim operation = action.Parameters("operation")?.ToString()

                Select Case operation?.ToLower()
                    Case "create_sheet"
                        Return CreateWorksheet(action.Parameters)

                    Case "fill_data"
                        Return FillDataToSheet(action.Parameters)

                    Case "format_cells"
                        Return FormatCells(action.Parameters)

                    Case "merge_sheets"
                        Return MergeWorksheets(action.Parameters)

                    Case "split_workbook"
                        Return SplitWorkbook(action.Parameters)

                    Case "read_data"
                        Return ReadDataFromSheet(action.Parameters)

                    Case Else
                        Return OperationResult.CreateFailure($"未知的Excel操作: {operation}")
                End Select

            Catch ex As Exception
                Return OperationResult.CreateFailure($"Excel操作失败: {ex.Message}")
            End Try
        End Function

        ''' <summary>
        ''' Excel 操作参数校验入口。
        ''' 对 operation 类型及对应必填参数、类型、边界进行校验，失败时返回 OperationResult，成功返回 Nothing。
        ''' </summary>
        Private Function ValidateExcelOperationParameters(action As ActionItem) As OperationResult
            If action Is Nothing Then
                Return OperationResult.CreateFailure("动作对象不能为空")
            End If

            If action.Parameters Is Nothing OrElse Not action.Parameters.ContainsKey("operation") Then
                Return OperationResult.CreateFailure("缺少 operation 参数")
            End If

            Dim operation = action.Parameters("operation")?.ToString()
            If String.IsNullOrWhiteSpace(operation) Then
                Return OperationResult.CreateFailure("operation 参数不能为空")
            End If

            Select Case operation.ToLower()
                Case "create_sheet"
                    If action.Parameters.ContainsKey("sheetName") Then
                        Dim sheetName = action.Parameters("sheetName")?.ToString()
                        If Not String.IsNullOrEmpty(sheetName) AndAlso Not IsValidSheetName(sheetName) Then
                            Return OperationResult.CreateFailure($"工作表名称非法或超出31字符限制: {sheetName}")
                        End If
                    End If

                Case "fill_data"
                    If Not action.Parameters.ContainsKey("data") OrElse action.Parameters("data") Is Nothing Then
                        Return OperationResult.CreateFailure("fill_data 缺少 data 参数")
                    End If
                    If action.Parameters.ContainsKey("startCell") Then
                        Dim startCell = action.Parameters("startCell")?.ToString()
                        If Not String.IsNullOrWhiteSpace(startCell) AndAlso Not IsValidRangeAddress(startCell) Then
                            Return OperationResult.CreateFailure($"startCell 地址格式非法: {startCell}")
                        End If
                    End If

                Case "format_cells"
                    If Not action.Parameters.ContainsKey("range") OrElse String.IsNullOrWhiteSpace(action.Parameters("range")?.ToString()) Then
                        Return OperationResult.CreateFailure("format_cells 缺少 range 参数")
                    End If
                    If Not action.Parameters.ContainsKey("format") OrElse action.Parameters("format") Is Nothing Then
                        Return OperationResult.CreateFailure("format_cells 缺少 format 参数")
                    End If
                    Dim formatObj = action.Parameters("format")
                    If Not (TypeOf formatObj Is String OrElse TypeOf formatObj Is Newtonsoft.Json.Linq.JObject) Then
                        Return OperationResult.CreateFailure("format 参数必须是 JSON 字符串或对象")
                    End If

                Case "merge_sheets"
                    If Not action.Parameters.ContainsKey("sourceSheets") OrElse action.Parameters("sourceSheets") Is Nothing Then
                        Return OperationResult.CreateFailure("merge_sheets 缺少 sourceSheets 参数")
                    End If
                    Dim sourceSheetsObj As Object = action.Parameters("sourceSheets")
                    Dim sourceSheetNames As New List(Of String)()
                    If TypeOf sourceSheetsObj Is Newtonsoft.Json.Linq.JArray Then
                        sourceSheetNames = CType(sourceSheetsObj, Newtonsoft.Json.Linq.JArray).
                            Select(Function(t) t.ToString()).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
                    ElseIf TypeOf sourceSheetsObj Is System.Collections.IEnumerable AndAlso Not TypeOf sourceSheetsObj Is String Then
                        sourceSheetNames = CType(sourceSheetsObj, System.Collections.IEnumerable).
                            Cast(Of Object)().Select(Function(o) o.ToString()).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
                    ElseIf sourceSheetsObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sourceSheetsObj.ToString()) Then
                        sourceSheetNames = sourceSheetsObj.ToString().Split(New Char() {","c, ";"c, "、"c}, StringSplitOptions.RemoveEmptyEntries).
                            Select(Function(s) s.Trim()).Where(Function(s) Not String.IsNullOrWhiteSpace(s)).ToList()
                    End If
                    If sourceSheetNames.Count = 0 Then
                        Return OperationResult.CreateFailure("sourceSheets 不能为空")
                    End If
                    If action.Parameters.ContainsKey("targetSheet") Then
                        Dim targetSheet = action.Parameters("targetSheet")?.ToString()
                        If Not String.IsNullOrEmpty(targetSheet) AndAlso Not IsValidSheetName(targetSheet) Then
                            Return OperationResult.CreateFailure($"targetSheet 名称非法或超出31字符限制: {targetSheet}")
                        End If
                    End If

                Case "split_workbook"
                    If action.Parameters.ContainsKey("splitMethod") Then
                        Dim splitMethod = action.Parameters("splitMethod")?.ToString()?.ToLower()
                        If Not String.IsNullOrWhiteSpace(splitMethod) AndAlso splitMethod <> "by_sheet" Then
                            Return OperationResult.CreateFailure($"不支持的 splitMethod: {splitMethod}，仅支持 by_sheet")
                        End If
                    End If
                    If action.Parameters.ContainsKey("outputFolder") Then
                        Dim outputFolder = action.Parameters("outputFolder")?.ToString()
                        If Not String.IsNullOrWhiteSpace(outputFolder) AndAlso Not System.IO.Directory.Exists(outputFolder) Then
                            Return OperationResult.CreateFailure($"输出目录不存在: {outputFolder}")
                        End If
                    End If

                Case "read_data"
                    If action.Parameters.ContainsKey("range") Then
                        Dim rangeAddr = action.Parameters("range")?.ToString()
                        If Not String.IsNullOrWhiteSpace(rangeAddr) AndAlso Not IsValidRangeAddress(rangeAddr) Then
                            Return OperationResult.CreateFailure($"range 地址格式非法: {rangeAddr}")
                        End If
                    End If

                Case Else
                    Return OperationResult.CreateFailure($"未知的Excel操作: {operation}")
            End Select

            Return Nothing
        End Function

        ''' <summary>
        ''' 校验工作表名称：非空、不含非法字符、长度不超过31。
        ''' </summary>
        Private Shared Function IsValidSheetName(sheetName As String) As Boolean
            If String.IsNullOrWhiteSpace(sheetName) Then Return False
            If sheetName.Length > 31 Then Return False
            Dim invalidChars = New Char() {":"c, "\"c, "/"c, "?"c, "*"c, "["c, "]"c}
            Return Not sheetName.Any(Function(c) invalidChars.Contains(c))
        End Function

        ''' <summary>
        ''' 简单校验 Excel 区域地址格式（如 A1、A1:B10、Sheet1!A1:B10）。
        ''' </summary>
        Private Shared Function IsValidRangeAddress(address As String) As Boolean
            If String.IsNullOrWhiteSpace(address) Then Return False
            ' 去除可能的工作表前缀
            Dim addr = address
            If address.Contains("!") Then
                Dim parts = address.Split(New Char() {"!"c}, 2)
                addr = parts(1)
            End If
            ' 支持 A1、A1:B10、1:1、A:A、$A$1:$B$10 等
            Dim pattern = "^\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$|^\$?[A-Za-z]+:\$?[A-Za-z]+$|^\$?\d+:\$?\d+$"
            Return System.Text.RegularExpressions.Regex.IsMatch(addr.Trim(), pattern)
        End Function

        ''' <summary>
        ''' 生成并执行代码 — 通过 LLM 生成VBA代码后经由 VBACodeSandbox 安全执行
        ''' 如果前端已经生成好代码并通过 generatedCode 传入，则直接执行，不再重复调用 LLM。
        ''' </summary>
        Private Async Function GenerateAndExecuteCode(action As ActionItem) As Task(Of OperationResult)
            Try
                Dim description = action.Parameters("description")?.ToString()
                Dim vbaCode As String = Nothing

                ' 优先执行前端传入的已生成代码
                If action.Parameters.ContainsKey("generatedCode") Then
                    vbaCode = action.Parameters("generatedCode")?.ToString()
                End If

                ' 没有现成代码时，通过 LLM 生成
                If String.IsNullOrWhiteSpace(vbaCode) Then
                    If String.IsNullOrEmpty(description) Then
                        Return OperationResult.CreateFailure("缺少代码生成描述")
                    End If

                    Dim codePrompt As String = "请根据以下需求生成VBA代码（仅输出代码，不要解释）：" & description & vbCrLf &
                        "要求：代码必须是可执行的Sub或Function，使用Excel原生对象模型。"

                    Dim llmEngine = MultimodalLLMEngine.SharedInstance
                    Dim result = Await llmEngine.ProcessTextAsync(codePrompt)

                    If Not result.Success OrElse String.IsNullOrWhiteSpace(result.Content) Then
                        Return OperationResult.CreateFailure("代码生成失败: " & result.ErrorMessage)
                    End If

                    ' 提取生成的VBA代码（去除可能的markdown代码块标记）
                    vbaCode = result.Content.Trim()
                    If vbaCode.StartsWith("```") Then
                        Dim firstNewLine = vbaCode.IndexOf(vbLf)
                        If firstNewLine > 0 Then vbaCode = vbaCode.Substring(firstNewLine + 1)
                        If vbaCode.EndsWith("```") Then vbaCode = vbaCode.Substring(0, vbaCode.Length - 3).Trim()
                    End If
                End If

                If String.IsNullOrWhiteSpace(vbaCode) Then
                    Return OperationResult.CreateFailure("生成的代码为空")
                End If

                ' 通过 VBACodeSandbox 安全执行生成的代码
                Dim execResult As Object = Security.VBACodeSandbox.ExecuteCode(vbaCode, requireConfirmation:=True)
                Dim execMethod As String = ""
                Try
                    execMethod = Microsoft.VisualBasic.Interaction.CallByName(execResult, "ExecutionMethod", CallType.Get)?.ToString()
                Catch ex As Exception  ' 
                End Try
                Return OperationResult.CreateSuccess(
                    "代码已生成并执行成功。执行方式: " & If(String.IsNullOrEmpty(execMethod), "Unknown", execMethod),
                    execResult)

            Catch ex As Exception
                Return OperationResult.CreateFailure("代码生成失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 数据提取并填充 — 将图片/表格数据解析后填入Excel工作表
        ''' </summary>
        Private Async Function ExtractAndFillData(action As ActionItem) As Task(Of OperationResult)
            Try
                Dim dataType = action.Parameters("dataType")?.ToString()
                Dim imageBase64 = action.Parameters("imageBase64")?.ToString()
                Dim targetSheet = action.Parameters("targetSheet")?.ToString()

                If String.IsNullOrEmpty(imageBase64) Then
                    Return OperationResult.CreateFailure("缺少图片数据")
                End If

                ' 【修复实现】根据数据类型路由到对应服务
                Select Case dataType?.ToLower()
                    Case "table", "ocr_table"
                        ' 使用 OCR 服务识别表格：先将base64保存为临时文件，再调用RecognizeText
                        Dim tempPath As String = System.IO.Path.GetTempFileName() & ".png"
                        Try
                            Dim imageBytes = Convert.FromBase64String(imageBase64)
                            System.IO.File.WriteAllBytes(tempPath, imageBytes)
                            Dim ocrMgr = OcrServiceManager.Instance
                            Dim ocrText = Await ocrMgr.RecognizeText(tempPath)
                            If Not String.IsNullOrWhiteSpace(ocrText) Then
                                Dim excelSvc = ExcelService.Instance
                                Dim ocrDict As New Dictionary(Of String, Object) From {{"content", ocrText}}
                                excelSvc.FillDataToExcel(ocrDict, dataType, If(targetSheet, ""))
                                Return OperationResult.CreateSuccess(
                                    "已识别并填入表格数据到工作表 " & If(targetSheet, "当前活动表"),
                                    New With {.ocrText = ocrText})
                            Else
                                Return OperationResult.CreateFailure("OCR识别未返回有效数据")
                            End If
                        Finally
                            Try : System.IO.File.Delete(tempPath) : Catch : End Try
                        End Try

                    Case "image_text", "text"
                        ' 使用多模态LLM提取图片中的文本（通过ProcessContentAsync）
                        Dim llmEngine = MultimodalLLMEngine.SharedInstance
                        Dim contents = New List(Of MultimodalContent) From {
                            MultimodalContent.CreateImage(imageBase64, "image/png")
                        }
                        Dim llmResult = Await llmEngine.ProcessContentAsync(contents,
                            "从图片中提取所有文本内容，保持原有格式和结构。只输出提取的文本内容。")
                        If llmResult.Success Then
                            Return OperationResult.CreateSuccess(
                                "已提取图片文本内容",
                                New With {.extractedText = llmResult.Content})
                        Else
                            Return OperationResult.CreateFailure("文本提取失败: " & llmResult.ErrorMessage)
                        End If

                    Case Else
                        Return OperationResult.CreateFailure(
                            "不支持的数据提取类型: " & dataType & "，支持类型: table, ocr_table, image_text, text")
                End Select

            Catch ex As Exception
                Return OperationResult.CreateFailure("数据提取失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 图片识别 — 调用多模态LLM或OCR引擎进行图片内容分析
        ''' </summary>
        Private Async Function RecognizeImage(action As ActionItem) As Task(Of OperationResult)
            Try
                Dim imageBase64 = action.Parameters("imageBase64")?.ToString()
                Dim recognitionType = action.Parameters("recognitionType")?.ToString()

                If String.IsNullOrEmpty(imageBase64) Then
                    Return OperationResult.CreateFailure("缺少图片数据")
                End If

                ' 【修复实现】根据识别类型路由到对应的服务
                Dim recognitionTypeLower = If(recognitionType, "general").ToLower()

                Select Case recognitionTypeLower
                    Case "ocr", "text"
                        ' 使用OCR引擎：先将base64保存为临时文件，再调用RecognizeText
                        Dim tempPath As String = System.IO.Path.GetTempFileName() & ".png"
                        Try
                            Dim imageBytes = Convert.FromBase64String(imageBase64)
                            System.IO.File.WriteAllBytes(tempPath, imageBytes)
                            Dim ocrMgr = OcrServiceManager.Instance
                            Dim ocrText = Await ocrMgr.RecognizeText(tempPath)
                            If Not String.IsNullOrWhiteSpace(ocrText) Then
                                Return OperationResult.CreateSuccess(
                                    "OCR识别完成",
                                    New With {.ocrText = ocrText})
                            Else
                                Return OperationResult.CreateFailure("OCR识别未返回结果")
                            End If
                        Finally
                            Try : System.IO.File.Delete(tempPath) : Catch : End Try
                        End Try

                    Case "invoice", "receipt"
                        ' 发票识别服务已弃用，改用多模态LLM进行发票识别
                        Dim llmEngineInv = MultimodalLLMEngine.SharedInstance
                        Dim contentsInv = New List(Of MultimodalContent) From {
                            MultimodalContent.CreateImage(imageBase64, "image/png")
                        }
                        Dim llmResultInv = Await llmEngineInv.ProcessContentAsync(contentsInv,
                            "你是一个专业的发票识别专家。请从图片中提取发票信息，以JSON格式输出包含以下字段：发票号码、开票日期、金额、销售方名称、购买方名称、税额。如果无法识别请说明原因。")
                        If llmResultInv.Success Then
                            Return OperationResult.CreateSuccess(
                                "发票识别完成",
                                New With {.invoiceData = llmResultInv.Content})
                        Else
                            Return OperationResult.CreateFailure("发票识别失败: " & llmResultInv.ErrorMessage)
                        End If

                    Case "general", "description", "analysis"
                        ' 使用多模态LLM进行通用图片理解（通过ProcessContentAsync）
                        Dim llmEngine = MultimodalLLMEngine.SharedInstance
                        Dim userPrompt = action.Parameters("prompt")?.ToString()
                        If String.IsNullOrEmpty(userPrompt) Then
                            userPrompt = "请详细描述这张图片的内容，包括文字、图表、布局等所有可见信息。"
                        End If

                        Dim contents = New List(Of MultimodalContent) From {
                            MultimodalContent.CreateImage(imageBase64, "image/png")
                        }
                        Dim llmResult = Await llmEngine.ProcessContentAsync(contents,
                            "你是一个专业的图像分析师。详细描述图片内容，包括文字、图表、数据、布局等。" & vbCrLf & userPrompt)
                        If llmResult.Success Then
                            Return OperationResult.CreateSuccess(
                                "图片分析完成",
                                New With {.analysis = llmResult.Content})
                        Else
                            Return OperationResult.CreateFailure("图片分析失败: " & llmResult.ErrorMessage)
                        End If

                    Case Else
                        Return OperationResult.CreateFailure(
                            "不支持的识别类型: " & recognitionTypeLower & "，支持: ocr, invoice, general, description")
                End Select

            Catch ex As Exception
                Return OperationResult.CreateFailure("图片识别失败: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' 执行 MCP 工具调用
        ''' 统一路由：
        '''   - 有 clientName → 走外部 MCP Client Pool
        '''   - 无 clientName → 自动识别本地工具（excel_/ppt_/computer_ 前缀 或 registry.Match 翻译）
        ''' </summary>
        Private Async Function ExecuteMcpToolCall(action As ActionItem) As Task(Of OperationResult)
            Try
                If _excelBridge Is Nothing Then
                    Return OperationResult.CreateFailure("ExcelBridge 未注入，无法调用 MCP 工具")
                End If

                Dim clientName = action.Parameters("clientName")?.ToString()
                Dim toolName = action.Parameters("toolName")?.ToString()
                Dim argumentsJson = action.Parameters("arguments")?.ToString()

                If String.IsNullOrEmpty(toolName) Then Return OperationResult.CreateFailure("缺少工具名称 (toolName)")
                If String.IsNullOrEmpty(argumentsJson) Then argumentsJson = "{}"

                ' 🔥 自动翻译：LLM 可能输出中文描述（如"截取当前屏幕"），用 registry.Match 转成标准工具名
                Dim originalToolName = toolName
                Dim registry = Services.Mcp.McpCapabilityRegistry.Instance
                Services.Mcp.McpCapabilityRegistry.RegisterAllBuiltInToolSets() ' 兜底注册
                ' 先查本地工具是否已叫这个名（前缀匹配）
                Dim allCaps = registry.GetAllCapabilities().Where(Function(c) c.Kind = Services.Mcp.CapabilityKind.McpTool).ToList()
                Dim directMatch = allCaps.FirstOrDefault(Function(c) String.Equals(c.Id, toolName, StringComparison.OrdinalIgnoreCase))
                If directMatch Is Nothing Then
                    ' 尝试关键词匹配（取较大窗口 top50，再按"同族"择优，避免泛化名跨族错译）
                    Dim matches = registry.Match(toolName, topN:=50)
                    If matches IsNot Nothing AndAlso matches.Count > 0 AndAlso matches(0).Score >= 20 Then
                        ' 🔥 同族守卫：LLM 可能输出泛化名（如 "chart"），能力匹配易误命中跨族工具（如 ppt_add_chart_slide）。
                        '   - 原名带家族前缀（excel_/ppt_/computer_/web_/time_）→ 只允许译到同族
                        '   - 裸名 → 默认 excel_ 系；含 ppt/幻灯片/演示 → ppt_ 系
                        Dim srcFamily = GetToolFamily(originalToolName)
                        Dim desired = srcFamily
                        If String.IsNullOrEmpty(desired) Then
                            Dim lower = originalToolName.ToLowerInvariant()
                            If lower.Contains("ppt") OrElse originalToolName.Contains("幻灯片") OrElse originalToolName.Contains("演示") Then
                                desired = "ppt_"
                            Else
                                desired = "excel_"
                            End If
                        End If
                        Dim pick = matches.FirstOrDefault(Function(m) String.Equals(GetToolFamily(m.Capability.Id), desired, StringComparison.OrdinalIgnoreCase))
                        If pick Is Nothing AndAlso String.Equals(desired, "excel_", StringComparison.OrdinalIgnoreCase) Then
                            ' 裸名兜底：取任意非 ppt_ 家族（仍避免误入 PPT）
                            pick = matches.FirstOrDefault(Function(m) GetToolFamily(m.Capability.Id) <> "ppt_")
                        End If
                        If pick IsNot Nothing Then
                            toolName = pick.Capability.Id
                            System.Diagnostics.Debug.WriteLine($"[ActionEngine] 工具名翻译: ""{originalToolName}"" → ""{toolName}"" (score={pick.Score})")
                        Else
                            System.Diagnostics.Debug.WriteLine($"[ActionEngine] 工具名翻译被同族守卫拦截: ""{originalToolName}""（top1={matches(0).Capability.Id}, score={matches(0).Score}）")
                        End If
                    End If
                End If

                ' 🔥 路由：有 clientName → 外部 Pool；无 → 本地 CallLocalTool
                If Not String.IsNullOrEmpty(clientName) Then
                    ' 外部 MCP Client Pool
                    System.Diagnostics.Debug.WriteLine(String.Format("[ActionEngine] 外部MCP: {0}.{1}", clientName, toolName))
                    Dim resultJson = Await _excelBridge.CallMcpToolAsync(clientName, toolName, argumentsJson)
                    Dim resultObj = Newtonsoft.Json.Linq.JObject.Parse(resultJson)
                    Dim successToken = resultObj("success")
                    Dim isSuccess = If(successToken IsNot Nothing, successToken.ToString() = "True", False)
                    If isSuccess Then
                        Dim toolResultToken = resultObj("result")
                        Dim toolResult = If(toolResultToken IsNot Nothing, toolResultToken.ToString(), "")
                        Return OperationResult.CreateSuccess(String.Format("MCP 工具 {0}.{1} 执行成功", clientName, toolName), toolResult)
                    Else
                        Dim errorToken = resultObj("error")
                        Dim errorMsg = If(errorToken IsNot Nothing, errorToken.ToString(), resultObj.ToString())
                        Return OperationResult.CreateFailure(String.Format("MCP 工具调用失败: {0}", errorMsg))
                    End If
                Else
                    ' 本地工具（Excel / PPT / Computer）
                    System.Diagnostics.Debug.WriteLine(String.Format("[ActionEngine] 本地MCP: {0}", toolName))
                    Dim resultJson = ExecuteOnUIThread(Function() Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IMcpBridge)().CallLocalTool(toolName, argumentsJson))
                    Dim resultObj = Newtonsoft.Json.Linq.JObject.Parse(resultJson)
                    Dim successToken = resultObj("success")
                    Dim isSuccess = If(successToken IsNot Nothing, successToken.ToString() = "True", False)
                    Dim isErrorToken = resultObj("isError")
                    Dim isError = If(isErrorToken IsNot Nothing, isErrorToken.ToString() = "True", False)
                    If isSuccess AndAlso Not isError Then
                        Dim toolResultToken = resultObj("result")
                        Dim toolResult = If(toolResultToken IsNot Nothing, toolResultToken.ToString(), "")
                        Return OperationResult.CreateSuccess(String.Format("✅ {0} 执行成功", toolName), toolResult)
                    Else
                        Dim errorToken = resultObj("error")
                        Dim errorMsg = If(errorToken IsNot Nothing, errorToken.ToString(),
                                          If(resultObj("result") IsNot Nothing, resultObj("result").ToString(), resultJson))
                        Return OperationResult.CreateFailure(String.Format("❌ {0} 执行失败: {1}", toolName, errorMsg))
                    End If
                End If

            Catch ex As Exception
                Return OperationResult.CreateFailure(String.Format("MCP 工具执行失败: {0}", ex.Message))
            End Try
        End Function

        ''' <summary>
        ''' 提取工具家族前缀（excel_/ppt_/computer_/web_/time_/skill__），无前缀返回空串。
        ''' 用于工具名翻译时做"同族一致性"守卫，避免泛化名跨族错译（如 chart → ppt_add_chart_slide）。
        ''' </summary>
        Private Shared Function GetToolFamily(toolId As String) As String
            If String.IsNullOrEmpty(toolId) Then Return ""
            Dim lower = toolId.ToLowerInvariant()
            For Each pfx In {"excel_", "ppt_", "computer_", "web_", "time_", "skill__"}
                If lower.StartsWith(pfx) Then Return pfx
            Next
            Return ""
        End Function

        #Region "Excel操作实现"

        ''' <summary>
        ''' 获取 Excel Application 实例
        ''' </summary>
        Private Shared Function GetExcelApp() As Excel.Application
            Return Globals.ThisAddIn.Application
        End Function

        ''' <summary>
        ''' 根据名称获取工作表；名称为空时使用活动工作表
        ''' </summary>
        Private Shared Function GetWorksheetByNameOrActive(sheetName As String) As Excel.Worksheet
            Dim app = GetExcelApp()
            If app.ActiveWorkbook Is Nothing Then
                Throw New InvalidOperationException("当前没有打开的工作簿")
            End If

            If String.IsNullOrWhiteSpace(sheetName) Then
                Return CType(app.ActiveSheet, Excel.Worksheet)
            End If

            For Each ws As Excel.Worksheet In app.ActiveWorkbook.Worksheets
                If ws.Name = sheetName Then
                    Return ws
                End If
            Next

            Throw New InvalidOperationException($"未找到工作表: {sheetName}")
        End Function

        ''' <summary>
        ''' 判断工作表是否已存在
        ''' </summary>
        Private Shared Function SheetExists(workbook As Excel.Workbook, sheetName As String) As Boolean
            For Each ws As Excel.Worksheet In workbook.Worksheets
                If ws.Name = sheetName Then
                    Return True
                End If
            Next
            Return False
        End Function

        ''' <summary>
        ''' 将列号转换为 Excel 列字母（1 -> A, 27 -> AA）
        ''' </summary>
        Private Shared Function GetColumnLetter(colIndex As Integer) As String
            Dim result = String.Empty
            Dim n = colIndex
            While n > 0
                Dim remainder = (n - 1) Mod 26
                result = ChrW(Asc("A") + remainder) & result
                n = (n - 1) \ 26
            End While
            Return result
        End Function

        ''' <summary>
        ''' 在目标表头列表中查找与给定表头最匹配的索引。
        ''' 匹配策略：精确匹配 -> 不区分大小写 -> 去空白 -> 相互包含 -> 编辑距离。
        ''' 返回匹配的索引，未找到返回 -1。
        ''' </summary>
        Private Shared Function FindBestHeaderMatch(commonHeaders As List(Of String), header As String) As Integer
            If commonHeaders Is Nothing OrElse commonHeaders.Count = 0 OrElse String.IsNullOrWhiteSpace(header) Then
                Return -1
            End If

            Dim normalizedHeader = header.Trim()

            ' 1. 精确匹配
            For i = 0 To commonHeaders.Count - 1
                If commonHeaders(i) = normalizedHeader Then Return i
            Next

            ' 2. 不区分大小写
            For i = 0 To commonHeaders.Count - 1
                If String.Equals(commonHeaders(i), normalizedHeader, StringComparison.OrdinalIgnoreCase) Then Return i
            Next

            ' 3. 去空白后匹配
            Dim noSpaceHeader = normalizedHeader.Replace(" ", "").Replace("　", "")
            For i = 0 To commonHeaders.Count - 1
                Dim noSpaceTarget = commonHeaders(i).Replace(" ", "").Replace("　", "")
                If String.Equals(noSpaceTarget, noSpaceHeader, StringComparison.OrdinalIgnoreCase) Then Return i
            Next

            ' 4. 相互包含（适用于"客户名"与"客户名称"）
            For i = 0 To commonHeaders.Count - 1
                Dim target = commonHeaders(i)
                If target.IndexOf(normalizedHeader, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                   normalizedHeader.IndexOf(target, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return i
                End If
            Next

            ' 5. 编辑距离（仅对较短表头启用，避免长句误匹配）
            If normalizedHeader.Length <= 12 Then
                Dim bestIndex = -1
                Dim bestDistance = Integer.MaxValue
                For i = 0 To commonHeaders.Count - 1
                    Dim target = commonHeaders(i)
                    If target.Length > 12 Then Continue For
                    Dim distance = ComputeLevenshteinDistance(normalizedHeader, target)
                    If distance < bestDistance AndAlso distance <= 2 Then
                        bestDistance = distance
                        bestIndex = i
                    End If
                Next
                Return bestIndex
            End If

            Return -1
        End Function

        ''' <summary>
        ''' 计算两个字符串的 Levenshtein 编辑距离。
        ''' </summary>
        Private Shared Function ComputeLevenshteinDistance(s As String, t As String) As Integer
            If String.IsNullOrEmpty(s) Then Return If(t Is Nothing, 0, t.Length)
            If String.IsNullOrEmpty(t) Then Return s.Length

            Dim n = s.Length
            Dim m = t.Length
            Dim previousRow(m) As Integer
            Dim currentRow(m) As Integer

            For j = 0 To m
                previousRow(j) = j
            Next

            For i = 1 To n
                currentRow(0) = i
                For j = 1 To m
                    Dim cost = If(s(i - 1) = t(j - 1), 0, 1)
                    currentRow(j) = Math.Min(Math.Min(currentRow(j - 1) + 1, previousRow(j) + 1), previousRow(j - 1) + cost)
                Next
                Dim temp = previousRow
                previousRow = currentRow
                currentRow = temp
            Next

            Return previousRow(m)
        End Function

        ''' <summary>
        ''' 如果文件路径已存在，则在文件名后追加递增序号，避免覆盖已有文件。
        ''' </summary>
        Private Shared Function GetUniqueFilePath(filePath As String) As String
            If Not System.IO.File.Exists(filePath) Then Return filePath

            Dim folder = System.IO.Path.GetDirectoryName(filePath)
            Dim fileNameWithoutExt = System.IO.Path.GetFileNameWithoutExtension(filePath)
            Dim extension = System.IO.Path.GetExtension(filePath)
            Dim index = 1

            Do
                Dim newPath = System.IO.Path.Combine(folder, $"{fileNameWithoutExt}_{index}{extension}")
                If Not System.IO.File.Exists(newPath) Then Return newPath
                index += 1
            Loop
        End Function

        ''' <summary>
        ''' 捕获指定区域的旧值，用于撤销数据填充
        ''' </summary>
        Private Shared Function CaptureRangeValues(ws As Excel.Worksheet, startRow As Integer, startCol As Integer, rowCount As Integer, colCount As Integer) As List(Of List(Of Object))
            Dim result As New List(Of List(Of Object))()
            Dim rng As Excel.Range = Nothing
            Try
                If rowCount <= 0 OrElse colCount <= 0 Then Return result
                Dim topLeftCell As Excel.Range = CType(ws.Cells(startRow, startCol), Excel.Range)
                Dim bottomRightCell As Excel.Range = CType(ws.Cells(startRow + rowCount - 1, startCol + colCount - 1), Excel.Range)
                rng = CType(ws.Range(topLeftCell, bottomRightCell), Excel.Range)
                Utils.ComReleaseHelper.Release(topLeftCell)
                Utils.ComReleaseHelper.Release(bottomRightCell)
                Dim values As Object(,) = CType(rng.Value, Object(,))
                If values Is Nothing Then
                    For r As Integer = 0 To rowCount - 1
                        Dim row As New List(Of Object)()
                        For c As Integer = 0 To colCount - 1
                            row.Add(Nothing)
                        Next
                        result.Add(row)
                    Next
                    Return result
                End If

                For r As Integer = 1 To values.GetLength(0)
                    Dim row As New List(Of Object)()
                    For c As Integer = 1 To values.GetLength(1)
                        row.Add(values(r, c))
                    Next
                    result.Add(row)
                Next
                Return result
            Catch ex As Exception  ' 
                Return result
            Finally
                Utils.ComReleaseHelper.Release(rng)
            End Try
        End Function

        ''' <summary>
        ''' 创建工作表
        ''' </summary>
        Private Function CreateWorksheet(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim wb As Excel.Workbook = Nothing
            Dim ws As Excel.Worksheet = Nothing
            Try
                Dim sheetName = parameters("sheetName")?.ToString()
                Dim app = GetExcelApp()
                wb = app.ActiveWorkbook
                If wb Is Nothing Then
                    Return OperationResult.CreateFailure("当前没有打开的工作簿")
                End If

                ' 未提供名称时自动生成 SheetN
                If String.IsNullOrWhiteSpace(sheetName) Then
                    Dim index = wb.Worksheets.Count + 1
                    Do
                        sheetName = $"Sheet{index}"
                        index += 1
                    Loop While SheetExists(wb, sheetName)
                ElseIf SheetExists(wb, sheetName) Then
                    Return OperationResult.CreateFailure($"工作表 '{sheetName}' 已存在")
                End If

                ws = CType(wb.Worksheets.Add(), Excel.Worksheet)
                ws.Name = sheetName

                ' 注册撤销动作：删除该工作表
                Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddUndoAction(New Bridges.UndoAction(
                    "create_sheet", "", Nothing, Nothing, sheetName))

                Return OperationResult.CreateSuccess($"工作表 '{sheetName}' 创建成功", New With {.sheetName = sheetName})

            Catch ex As Exception
                Return OperationResult.CreateFailure($"创建工作表失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' 填充数据到工作表
        ''' 支持两种 data 格式：
        ''' 1. List(Of Dictionary(Of String, Object))：按列标题+行数据写入
        ''' 2. 二维数组 / JArray of JArray：原样写入
        ''' </summary>
        Private Function FillDataToSheet(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim ws As Excel.Worksheet = Nothing
            Dim rng As Excel.Range = Nothing
            Try
                Dim sheetName = parameters("sheetName")?.ToString()
                Dim startCell = parameters("startCell")?.ToString()
                Dim data = parameters("data")

                If data Is Nothing Then
                    Return OperationResult.CreateFailure("缺少数据")
                End If

                ws = GetWorksheetByNameOrActive(sheetName)
                Dim app = GetExcelApp()
                Dim startRow = 1
                Dim startCol = 1

                If Not String.IsNullOrWhiteSpace(startCell) Then
                    Dim addrRng As Excel.Range = CType(ws.Range(startCell), Excel.Range)
                    startRow = CInt(addrRng.Row)
                    startCol = CInt(addrRng.Column)
                    Utils.ComReleaseHelper.Release(addrRng)
                Else
                    Dim activeCellRng As Excel.Range = CType(app.ActiveCell, Excel.Range)
                    startRow = CInt(activeCellRng.Row)
                    startCol = CInt(activeCellRng.Column)
                End If

                ' 计算数据维度并写入
                Dim rowCount As Integer = 0
                Dim colCount As Integer = 0
                Dim jtoken = TryCast(data, Newtonsoft.Json.Linq.JToken)
                If jtoken IsNot Nothing AndAlso jtoken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    Dim jarray = CType(jtoken, Newtonsoft.Json.Linq.JArray)
                    If jarray.Count > 0 AndAlso jarray(0).Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                        rowCount = jarray.Count + 1 ' 含表头
                        colCount = jarray(0).Children(Of Newtonsoft.Json.Linq.JProperty)().Count()
                    Else
                        rowCount = jarray.Count
                        colCount = jarray(0).Count()
                    End If
                ElseIf TypeOf data Is System.Collections.IList AndAlso Not TypeOf data Is Byte() Then
                    Dim list = CType(data, System.Collections.IList)
                    If list.Count > 0 AndAlso TypeOf list(0) Is System.Collections.IDictionary Then
                        Dim first = CType(list(0), System.Collections.IDictionary)
                        rowCount = list.Count + 1
                        colCount = first.Keys.Count
                    Else
                        Return OperationResult.CreateFailure("不支持的数据格式")
                    End If
                ElseIf TypeOf data Is Array Then
                    Dim arr = CType(data, Array)
                    rowCount = arr.GetLength(0)
                    colCount = If(arr.Rank > 1, arr.GetLength(1), 1)
                Else
                    Return OperationResult.CreateFailure("不支持的数据格式")
                End If

                ' 撤销记录：整表备份（必须在写入之前，撤销时整张表还原）
                Try
                    Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                Catch undoEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[FillData] 撤销记录失败: {undoEx.Message}")
                End Try

                ' 写入数据
                If jtoken IsNot Nothing AndAlso jtoken.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                    Dim jarray = CType(jtoken, Newtonsoft.Json.Linq.JArray)
                    If jarray.Count > 0 AndAlso jarray(0).Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                        FillObjectArrayToSheet(ws, startRow, startCol, jarray)
                    Else
                        Fill2DArrayToSheet(ws, startRow, startCol, jarray)
                    End If
                ElseIf TypeOf data Is System.Collections.IList AndAlso Not TypeOf data Is Byte() Then
                    Dim list = CType(data, System.Collections.IList)
                    FillDictionaryListToSheet(ws, startRow, startCol, list)
                ElseIf TypeOf data Is Array Then
                    FillArrayToSheet(ws, startRow, startCol, CType(data, Array))
                End If

                ' 捕获新值（用于重做）
                Dim newValues = CaptureRangeValues(ws, startRow, startCol, rowCount, colCount)


                Return OperationResult.CreateSuccess("数据填充完成", New With {.sheetName = ws.Name, .startRow = startRow, .startCol = startCol})

            Catch ex As Exception
                Return OperationResult.CreateFailure($"填充数据失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' 将对象数组（JArray of JObject）写入工作表
        ''' </summary>
        Private Shared Sub FillObjectArrayToSheet(ws As Excel.Worksheet, startRow As Integer, startCol As Integer, jarray As Newtonsoft.Json.Linq.JArray)
            Dim headers As List(Of String) = jarray(0).Children(Of Newtonsoft.Json.Linq.JProperty)().Select(Function(p) p.Name).ToList()
            Dim rowCount As Integer = jarray.Count
            Dim colCount As Integer = headers.Count

            ' 一次性写入表头+数据到数组，再赋给 Range
            Dim values(rowCount, colCount - 1) As Object
            For c As Integer = 0 To colCount - 1
                values(0, c) = headers(c)
            Next
            For r As Integer = 0 To rowCount - 1
                Dim obj As Newtonsoft.Json.Linq.JObject = CType(jarray(r), Newtonsoft.Json.Linq.JObject)
                For c As Integer = 0 To colCount - 1
                    values(r + 1, c) = obj(headers(c))?.ToObject(Of Object)()
                Next
            Next

            Dim topLeft As Excel.Range = CType(ws.Cells(startRow, startCol), Excel.Range)
            Dim bottomRight As Excel.Range = CType(ws.Cells(startRow + rowCount, startCol + colCount - 1), Excel.Range)
            Dim rng As Excel.Range = CType(ws.Range(topLeft, bottomRight), Excel.Range)
            rng.Value = values
            Utils.ComReleaseHelper.Release(topLeft)
            Utils.ComReleaseHelper.Release(bottomRight)
            Utils.ComReleaseHelper.Release(rng)
        End Sub

        ''' <summary>
        ''' 将二维 JArray 原样写入工作表
        ''' </summary>
        Private Shared Sub Fill2DArrayToSheet(ws As Excel.Worksheet, startRow As Integer, startCol As Integer, jarray As Newtonsoft.Json.Linq.JArray)
            Dim rowCount As Integer = jarray.Count
            If rowCount = 0 Then Return
            Dim colCount As Integer = jarray(0).Count()

            Dim values(rowCount - 1, colCount - 1) As Object
            For r As Integer = 0 To rowCount - 1
                For c As Integer = 0 To colCount - 1
                    values(r, c) = jarray(r)(c)?.ToObject(Of Object)()
                Next
            Next

            Dim topLeft As Excel.Range = CType(ws.Cells(startRow, startCol), Excel.Range)
            Dim bottomRight As Excel.Range = CType(ws.Cells(startRow + rowCount - 1, startCol + colCount - 1), Excel.Range)
            Dim rng As Excel.Range = CType(ws.Range(topLeft, bottomRight), Excel.Range)
            rng.Value = values
            Utils.ComReleaseHelper.Release(topLeft)
            Utils.ComReleaseHelper.Release(bottomRight)
            Utils.ComReleaseHelper.Release(rng)
        End Sub

        ''' <summary>
        ''' 将 Dictionary 列表写入工作表
        ''' </summary>
        Private Shared Sub FillDictionaryListToSheet(ws As Excel.Worksheet, startRow As Integer, startCol As Integer, list As System.Collections.IList)
            Dim first As System.Collections.IDictionary = CType(list(0), System.Collections.IDictionary)
            Dim headers As List(Of String) = first.Keys.Cast(Of Object)().Select(Function(k) k.ToString()).ToList()
            Dim rowCount As Integer = list.Count
            Dim colCount As Integer = headers.Count

            Dim values(rowCount, colCount - 1) As Object
            For c As Integer = 0 To colCount - 1
                values(0, c) = headers(c)
            Next
            For r As Integer = 0 To rowCount - 1
                Dim dict As System.Collections.IDictionary = CType(list(r), System.Collections.IDictionary)
                For c As Integer = 0 To colCount - 1
                    values(r + 1, c) = dict(headers(c))
                Next
            Next

            Dim topLeft As Excel.Range = CType(ws.Cells(startRow, startCol), Excel.Range)
            Dim bottomRight As Excel.Range = CType(ws.Cells(startRow + rowCount, startCol + colCount - 1), Excel.Range)
            Dim rng As Excel.Range = CType(ws.Range(topLeft, bottomRight), Excel.Range)
            rng.Value = values
            Utils.ComReleaseHelper.Release(topLeft)
            Utils.ComReleaseHelper.Release(bottomRight)
            Utils.ComReleaseHelper.Release(rng)
        End Sub

        ''' <summary>
        ''' 将 .NET 二维数组写入工作表
        ''' </summary>
        Private Shared Sub FillArrayToSheet(ws As Excel.Worksheet, startRow As Integer, startCol As Integer, arr As Array)
            Dim rowCount As Integer = arr.GetLength(0)
            Dim colCount As Integer = If(arr.Rank > 1, arr.GetLength(1), 1)
            Dim topLeft As Excel.Range = CType(ws.Cells(startRow, startCol), Excel.Range)
            Dim bottomRight As Excel.Range = CType(ws.Cells(startRow + rowCount - 1, startCol + colCount - 1), Excel.Range)
            Dim rng As Excel.Range = CType(ws.Range(topLeft, bottomRight), Excel.Range)
            rng.Value = arr
            Utils.ComReleaseHelper.Release(topLeft)
            Utils.ComReleaseHelper.Release(bottomRight)
            Utils.ComReleaseHelper.Release(rng)
        End Sub

        ''' <summary>
        ''' 格式化单元格
        ''' format 支持 JSON 字符串或字典，常见键：bold/italic/underline/color/numberFormat/horizontalAlignment
        ''' </summary>
        Private Function FormatCells(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim ws As Excel.Worksheet = Nothing
            Dim rng As Excel.Range = Nothing
            Try
                Dim sheetName As String = parameters("sheetName")?.ToString()
                Dim rangeAddr As String = parameters("range")?.ToString()
                Dim formatObj As Object = parameters("format")

                If String.IsNullOrWhiteSpace(rangeAddr) Then
                    Return OperationResult.CreateFailure("缺少目标单元格区域 range")
                End If

                ws = GetWorksheetByNameOrActive(sheetName)
                rng = CType(ws.Range(rangeAddr), Excel.Range)

                ' 撤销记录：整表备份（必须在格式化之前，撤销时整张表还原）
                Try
                    Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddSheetBackup(ws.Name)
                Catch undoEx As Exception
                    System.Diagnostics.Debug.WriteLine($"[FormatCells] 撤销记录失败: {undoEx.Message}")
                End Try

                ' 统一把 format 转成 JObject
                Dim formatJson As Newtonsoft.Json.Linq.JObject
                If TypeOf formatObj Is String Then
                    formatJson = Newtonsoft.Json.Linq.JObject.Parse(CStr(formatObj))
                ElseIf TypeOf formatObj Is Newtonsoft.Json.Linq.JObject Then
                    formatJson = CType(formatObj, Newtonsoft.Json.Linq.JObject)
                Else
                    Return OperationResult.CreateFailure("format 参数格式不正确")
                End If

                With rng
                    If formatJson.ContainsKey("bold") Then CType(.Font, Excel.Font).Bold = formatJson("bold").ToObject(Of Boolean)()
                    If formatJson.ContainsKey("italic") Then CType(.Font, Excel.Font).Italic = formatJson("italic").ToObject(Of Boolean)()
                    If formatJson.ContainsKey("underline") Then CType(.Font, Excel.Font).Underline = If(formatJson("underline").ToObject(Of Boolean)(), CObj(Excel.XlUnderlineStyle.xlUnderlineStyleSingle), CObj(Excel.XlUnderlineStyle.xlUnderlineStyleNone))
                    If formatJson.ContainsKey("color") Then CType(.Font, Excel.Font).Color = CObj(System.Drawing.ColorTranslator.FromHtml(formatJson("color").ToString()).ToArgb())
                    If formatJson.ContainsKey("bgColor") Then CType(.Interior, Excel.Interior).Color = CObj(System.Drawing.ColorTranslator.FromHtml(formatJson("bgColor").ToString()).ToArgb())
                    If formatJson.ContainsKey("numberFormat") Then .NumberFormat = formatJson("numberFormat").ToString()
                    If formatJson.ContainsKey("horizontalAlignment") Then
                        Dim align As String = formatJson("horizontalAlignment").ToString().ToLower()
                        Select Case align
                            Case "left"
                                .HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft
                            Case "center"
                                .HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                            Case "right"
                                .HorizontalAlignment = Excel.XlHAlign.xlHAlignRight
                        End Select
                    End If
                    If formatJson.ContainsKey("columnWidth") Then .ColumnWidth = formatJson("columnWidth").ToObject(Of Double)()
                    If formatJson.ContainsKey("rowHeight") Then .RowHeight = formatJson("rowHeight").ToObject(Of Double)()
                End With

                Return OperationResult.CreateSuccess("格式设置完成", New With {.range = rangeAddr})

            Catch ex As Exception
                Return OperationResult.CreateFailure($"格式化失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        ''' <summary>
        ''' 合并多个工作表到一个目标工作表
        ''' </summary>
        Private Function MergeWorksheets(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim wb As Excel.Workbook = Nothing
            Dim targetWs As Excel.Worksheet = Nothing
            Try
                Dim sourceSheetsObj As Object = parameters("sourceSheets")
                Dim targetSheetName As String = parameters("targetSheet")?.ToString()
                Dim includeHeaders As Boolean = True
                If parameters.ContainsKey("includeHeaders") Then
                    Dim ihStr As String = parameters("includeHeaders").ToString().Trim()
                    includeHeaders = ihStr.Equals("true", StringComparison.OrdinalIgnoreCase)
                End If

                If sourceSheetsObj Is Nothing Then
                    Return OperationResult.CreateFailure("缺少源工作表列表 sourceSheets")
                End If

                Dim sourceSheetNames As New List(Of String)()
                If TypeOf sourceSheetsObj Is Newtonsoft.Json.Linq.JArray Then
                    sourceSheetNames = CType(sourceSheetsObj, Newtonsoft.Json.Linq.JArray).
                        Select(Function(t) t.ToString()).ToList()
                ElseIf TypeOf sourceSheetsObj Is System.Collections.IEnumerable AndAlso Not TypeOf sourceSheetsObj Is String Then
                    sourceSheetNames = CType(sourceSheetsObj, System.Collections.IEnumerable).
                        Cast(Of Object)().Select(Function(o) o.ToString()).ToList()
                ElseIf sourceSheetsObj IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(sourceSheetsObj.ToString()) Then
                    sourceSheetNames = sourceSheetsObj.ToString().Split(New Char() {","c, ";"c, "、"c}, StringSplitOptions.RemoveEmptyEntries).
                        Select(Function(s) s.Trim()).ToList()
                End If

                If sourceSheetNames.Count = 0 Then
                    Return OperationResult.CreateFailure("源工作表列表为空")
                End If

                Dim app As Excel.Application = GetExcelApp()
                wb = CType(app.ActiveWorkbook, Excel.Workbook)
                If wb Is Nothing Then
                    Return OperationResult.CreateFailure("当前没有打开的工作簿")
                End If

                ' 目标表
                If String.IsNullOrWhiteSpace(targetSheetName) Then
                    targetSheetName = $"合并_{DateTime.Now:MMdd_HHmmss}"
                End If

                If SheetExists(wb, targetSheetName) Then
                    targetWs = CType(wb.Worksheets(targetSheetName), Excel.Worksheet)
                    CType(targetWs.Cells, Excel.Range).Clear()
                Else
                    targetWs = CType(wb.Worksheets.Add(), Excel.Worksheet)
                    targetWs.Name = targetSheetName
                End If

                Dim currentRow As Integer = 1
                Dim headersWritten As Boolean = False
                Dim commonHeaders As List(Of String) = Nothing

                For Each sheetName As String In sourceSheetNames
                    Dim srcWs As Excel.Worksheet = CType(wb.Worksheets(sheetName), Excel.Worksheet)
                    Dim usedRng As Excel.Range = CType(srcWs.UsedRange, Excel.Range)
                    Dim values As Object(,) = CType(usedRng.Value, Object(,))

                    If values IsNot Nothing Then
                        Dim rowCount As Integer = values.GetLength(0)
                        Dim colCount As Integer = values.GetLength(1)

                        If rowCount = 0 OrElse colCount = 0 Then
                            Utils.ComReleaseHelper.Release(usedRng)
                            Utils.ComReleaseHelper.Release(srcWs)
                            Continue For
                        End If

                        If includeHeaders AndAlso Not headersWritten Then
                            ' 写入第一个表的表头
                            Dim headerArray(0, colCount - 1) As Object
                            For c As Integer = 0 To colCount - 1
                                headerArray(0, c) = values(1, c + 1)
                            Next
                            Dim hTopLeft As Excel.Range = CType(targetWs.Cells(currentRow, 1), Excel.Range)
                            Dim hBottomRight As Excel.Range = CType(targetWs.Cells(currentRow, colCount), Excel.Range)
                            Dim headerRng As Excel.Range = CType(targetWs.Range(hTopLeft, hBottomRight), Excel.Range)
                            headerRng.Value = headerArray
                            Utils.ComReleaseHelper.Release(hTopLeft)
                            Utils.ComReleaseHelper.Release(hBottomRight)
                            Utils.ComReleaseHelper.Release(headerRng)
                            commonHeaders = Enumerable.Range(1, colCount).Select(Function(c) If(values(1, c)?.ToString(), $"Col{c}")).ToList()
                            currentRow += 1
                            headersWritten = True

                            ' 写入第一个表的数据（不含表头）
                            If rowCount > 1 Then
                                Dim dataArray(rowCount - 2, colCount - 1) As Object
                                For r As Integer = 2 To rowCount
                                    For c As Integer = 1 To colCount
                                        dataArray(r - 2, c - 1) = values(r, c)
                                    Next
                                Next
                                Dim dTopLeft As Excel.Range = CType(targetWs.Cells(currentRow, 1), Excel.Range)
                                Dim dBottomRight As Excel.Range = CType(targetWs.Cells(currentRow + rowCount - 2, colCount), Excel.Range)
                                Dim dataRng As Excel.Range = CType(targetWs.Range(dTopLeft, dBottomRight), Excel.Range)
                                dataRng.Value = dataArray
                                Utils.ComReleaseHelper.Release(dTopLeft)
                                Utils.ComReleaseHelper.Release(dBottomRight)
                                Utils.ComReleaseHelper.Release(dataRng)
                                currentRow += rowCount - 1
                            End If
                        Else
                            ' 后续表：如果已写表头，尝试按列标题对齐；否则直接追加
                            If includeHeaders AndAlso headersWritten AndAlso commonHeaders IsNot Nothing Then
                                Dim srcHeaders As List(Of String) = Enumerable.Range(1, colCount).Select(Function(c) If(values(1, c)?.ToString(), "")).ToList()
                                Dim dataArray(rowCount - 1, commonHeaders.Count - 1) As Object
                                For r As Integer = 2 To rowCount
                                    For c As Integer = 1 To colCount
                                        Dim header As String = srcHeaders(c - 1)
                                        Dim targetCol As Integer = FindBestHeaderMatch(commonHeaders, header) + 1
                                        If targetCol > 0 Then
                                            dataArray(r - 2, targetCol - 1) = values(r, c)
                                        End If
                                    Next
                                Next
                                Dim dTopLeft As Excel.Range = CType(targetWs.Cells(currentRow, 1), Excel.Range)
                                Dim dBottomRight As Excel.Range = CType(targetWs.Cells(currentRow + rowCount - 2, commonHeaders.Count), Excel.Range)
                                Dim dataRng As Excel.Range = CType(targetWs.Range(dTopLeft, dBottomRight), Excel.Range)
                                dataRng.Value = dataArray
                                Utils.ComReleaseHelper.Release(dTopLeft)
                                Utils.ComReleaseHelper.Release(dBottomRight)
                                Utils.ComReleaseHelper.Release(dataRng)
                                currentRow += rowCount - 1
                            Else
                                Dim dataArray(rowCount - 1, colCount - 1) As Object
                                For r As Integer = 1 To rowCount
                                    For c As Integer = 1 To colCount
                                        dataArray(r - 1, c - 1) = values(r, c)
                                    Next
                                Next
                                Dim dTopLeft As Excel.Range = CType(targetWs.Cells(currentRow, 1), Excel.Range)
                                Dim dBottomRight As Excel.Range = CType(targetWs.Cells(currentRow + rowCount - 1, colCount), Excel.Range)
                                Dim dataRng As Excel.Range = CType(targetWs.Range(dTopLeft, dBottomRight), Excel.Range)
                                dataRng.Value = dataArray
                                Utils.ComReleaseHelper.Release(dTopLeft)
                                Utils.ComReleaseHelper.Release(dBottomRight)
                                Utils.ComReleaseHelper.Release(dataRng)
                                currentRow += rowCount
                            End If
                        End If
                    End If

                    Utils.ComReleaseHelper.Release(usedRng)
                    Utils.ComReleaseHelper.Release(srcWs)
                Next

                ' 捕获合并后目标表的完整数据（用于重做）
                Dim targetUsedRng As Excel.Range = Nothing
                Dim targetData As List(Of List(Of Object)) = New List(Of List(Of Object))()
                Try
                    targetUsedRng = CType(targetWs.UsedRange, Excel.Range)
                    If targetUsedRng IsNot Nothing Then
                        Dim vals As Object(,) = CType(targetUsedRng.Value, Object(,))
                        If vals IsNot Nothing Then
                            Dim tRows As Integer = vals.GetLength(0)
                            Dim tCols As Integer = vals.GetLength(1)
                            For tr = 1 To tRows
                                Dim row As New List(Of Object)()
                                For tc = 1 To tCols
                                    row.Add(vals(tr, tc))
                                Next
                                targetData.Add(row)
                            Next
                        End If
                    End If
                Finally
                    Utils.ComReleaseHelper.Release(targetUsedRng)
                End Try

                ' 注册撤销动作：删除目标工作表，Payload 保存数据用于重做
                Dim mergePayload = JsonConvert.SerializeObject(New With {
                    .sheetName = targetSheetName,
                    .data = targetData
                })
                Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddUndoAction(New Bridges.UndoAction(
                    "merge_sheets", "", Nothing, Nothing, targetSheetName, mergePayload))

                Return OperationResult.CreateSuccess($"已合并 {sourceSheetNames.Count} 个工作表到 '{targetSheetName}'", New With {.targetSheet = targetSheetName})

            Catch ex As Exception
                Return OperationResult.CreateFailure($"合并工作表失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(targetWs)
            End Try
        End Function

        ''' <summary>
        ''' 拆分工作簿：按工作表拆分为多个独立工作簿
        ''' </summary>
        Private Function SplitWorkbook(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim wb As Excel.Workbook = Nothing
            Try
                Dim splitMethod As String = parameters("splitMethod")?.ToString()?.ToLower()
                Dim outputFolder As String = parameters("outputFolder")?.ToString()

                If String.IsNullOrWhiteSpace(splitMethod) Then splitMethod = "by_sheet"

                Dim app As Excel.Application = GetExcelApp()
                wb = CType(app.ActiveWorkbook, Excel.Workbook)
                If wb Is Nothing Then
                    Return OperationResult.CreateFailure("当前没有打开的工作簿")
                End If

                If String.IsNullOrWhiteSpace(outputFolder) OrElse Not System.IO.Directory.Exists(outputFolder) Then
                    outputFolder = System.IO.Path.GetDirectoryName(wb.FullName)
                End If
                If String.IsNullOrWhiteSpace(outputFolder) Then
                    outputFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                End If

                Dim generatedFiles As New List(Of String)()
                Dim baseName As String = System.IO.Path.GetFileNameWithoutExtension(wb.Name)
                Dim sheetCount As Integer = CInt(wb.Worksheets.Count)

                For i As Integer = sheetCount To 1 Step -1
                    Dim srcWs As Excel.Worksheet = CType(wb.Worksheets(i), Excel.Worksheet)
                    Dim sheetName As String = srcWs.Name
                    srcWs.Copy()
                    Dim newWb As Excel.Workbook = CType(app.ActiveWorkbook, Excel.Workbook)
                    Dim fileName As String = $"{baseName}_{sheetName}.xlsx"
                    Dim filePath As String = GetUniqueFilePath(System.IO.Path.Combine(outputFolder, fileName))
                    newWb.SaveAs(filePath)
                    newWb.Close(SaveChanges:=False)
                    generatedFiles.Add(filePath)
                    Utils.ComReleaseHelper.Release(srcWs)
                    Utils.ComReleaseHelper.Release(newWb)
                Next

                ' 注册撤销动作：删除生成的文件，Payload 同时保存参数用于重做
                Dim sourceWorkbookPath As String = ""
                Try
                    sourceWorkbookPath = wb.FullName
                Catch ex As Exception  ' 
                End Try
                Dim splitPayload As String = JsonConvert.SerializeObject(New With {
                    .files = generatedFiles,
                    .splitMethod = splitMethod,
                    .outputFolder = outputFolder,
                    .sourceWorkbook = sourceWorkbookPath,
                    .sheetCount = sheetCount
                })
                Bridges.BridgeRegistry.Instance.Resolve(Of Bridges.IUndoBridge)().AddUndoAction(New Bridges.UndoAction(
                    "split_workbook", "", Nothing, Nothing, "", splitPayload))

                Return OperationResult.CreateSuccess($"已拆分为 {generatedFiles.Count} 个工作簿", New With {.files = generatedFiles, .outputFolder = outputFolder})

            Catch ex As Exception
                Return OperationResult.CreateFailure($"拆分工作簿失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(wb)
            End Try
        End Function

        ''' <summary>
        ''' 读取工作表数据，返回 JSON 可序列化的结构
        ''' </summary>
        Private Function ReadDataFromSheet(parameters As Dictionary(Of String, Object)) As OperationResult
            Dim ws As Excel.Worksheet = Nothing
            Dim rng As Excel.Range = Nothing
            Try
                Dim sheetName As String = parameters("sheetName")?.ToString()
                Dim rangeAddr As String = parameters("range")?.ToString()
                Dim includeEmpty As Boolean = False
                If parameters.ContainsKey("includeEmpty") Then
                    Dim ieStr As String = parameters("includeEmpty").ToString().Trim()
                    includeEmpty = ieStr.Equals("true", StringComparison.OrdinalIgnoreCase)
                End If

                ws = GetWorksheetByNameOrActive(sheetName)

                If String.IsNullOrWhiteSpace(rangeAddr) Then
                    rng = CType(ws.UsedRange, Excel.Range)
                Else
                    rng = CType(ws.Range(rangeAddr), Excel.Range)
                End If

                If rng Is Nothing OrElse CLng(rng.Cells.Count) = 0 Then
                    Return OperationResult.CreateSuccess("未读取到数据", New With {.rows = New List(Of Object)()})
                End If

                Dim values As Object(,) = CType(rng.Value, Object(,))
                Dim rowCount As Integer = values.GetLength(0)
                Dim colCount As Integer = values.GetLength(1)
                Dim rows As New List(Of Object)()

                For r As Integer = 1 To rowCount
                    Dim rowValues As New List(Of Object)()
                    Dim hasValue As Boolean = False
                    For c As Integer = 1 To colCount
                        Dim v As Object = values(r, c)
                        If v IsNot Nothing Then hasValue = True
                        rowValues.Add(v)
                    Next
                    If includeEmpty OrElse hasValue Then
                        rows.Add(rowValues)
                    End If
                Next

                Dim result = New With {
                    .sheetName = ws.Name,
                    .range = rng.Address.ToString(),
                    .rowCount = rowCount,
                    .colCount = colCount,
                    .rows = rows
                }

                Return OperationResult.CreateSuccess($"已读取 {rows.Count} 行数据", result)

            Catch ex As Exception
                Return OperationResult.CreateFailure($"读取工作表数据失败: {ex.Message}")
            Finally
                Utils.ComReleaseHelper.Release(rng)
                Utils.ComReleaseHelper.Release(ws)
            End Try
        End Function

        #End Region

        ''' <summary>
        ''' 获取可用动作列表
        ''' 优先暴露可直接执行的 VSTO Excel 操作；code_generation 仅作为复杂/宏生成场景的兜底。
        ''' </summary>
        Public Function GetAvailableActions() As List(Of String)
            Return New List(Of String) From {
                "excel_operation",
                "code_generation",
                "data_extraction",
                "image_recognition",
                "mcp_tool_call"
            }
        End Function

    End Class

End Namespace