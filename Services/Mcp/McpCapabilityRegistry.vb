Option Explicit On
Option Strict On
Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json

Namespace Services.Mcp

    ''' <summary>
    ''' 能力类型：MCP 工具 或 Skill 组合
    ''' </summary>
    Public Enum CapabilityKind
        McpTool = 1
        Skill = 2
    End Enum

    ''' <summary>
    ''' 能力分类（用于意图匹配和 UI 展示）
    ''' </summary>
    Public Enum CapabilityCategory
        Read = 1        ' 读取/查看
        Write = 2       ' 写入/修改
        Format = 3      ' 格式化/美化
        Analyze = 4     ' 分析/统计
        Transform = 5   ' 转换/拆分/合并
        Export = 6      ' 导出
        Import = 7      ' 导入
        Manage = 8      ' 管理（工作表/工作簿操作）
        Visualize = 9   ' 可视化/图表
        Utility = 10    ' 工具（截图/剪贴板/OCR）
        Compute = 11    ' 纯计算/分析（仅返回结果，不落盘；落地由 targetSheet/写表工具负责）
        ComputerAgent = 12  ' 电脑操作 Agent（键鼠/OCR/截图）
    End Enum

    ''' <summary>
    ''' 统一能力描述符：MCP 工具 或 Skill 的单一注册项
    ''' </summary>
    Public Class CapabilityDescriptor
        Public Property Id As String                        ' 唯一标识，如 "excel_remove_duplicates"
        Public Property Kind As CapabilityKind = CapabilityKind.McpTool
        Public Property Source As String = "excel-local"    ' 工具集来源（excel-local/ppt-local/外部客户端名），前端分组依据
        Public Property DisplayName As String              ' 中文名，如 "删除重复行"
        Public Property Description As String              ' 功能描述，用于匹配
        Public Property Category As CapabilityCategory     ' 分类
        Public Property Keywords As List(Of String)        ' 强匹配关键词（"去重","删除重复"等）
        Public Property IsWriteOperation As Boolean        ' 是否为写操作（需要真实执行）
        Public Property InputSchema As McpToolInputSchema  ' 参数 Schema（MCP 工具用）
        Public Property SkillName As String                ' Skill 原名（Kind=Skill 时用）
        Public Property IsDangerous As Boolean             ' 是否危险操作
        Public Property DangerSeverity As String           ' high|medium（来自 excel_tools.json）
        Public Property DangerMessage As String            ' 安全守卫确认文案

        Public Function ToMatchText() As String
            Dim parts As New List(Of String)()
            If Not String.IsNullOrEmpty(DisplayName) Then parts.Add(DisplayName)
            If Not String.IsNullOrEmpty(Description) Then parts.Add(Description)
            If Keywords IsNot Nothing Then parts.AddRange(Keywords)
            Return String.Join(" ", parts).ToLowerInvariant()
        End Function
    End Class

    ''' <summary>
    ''' 匹配结果
    ''' </summary>
    Public Class CapabilityMatchResult
        Public Property Capability As CapabilityDescriptor
        Public Property Score As Double
        Public Property MatchedKeywords As List(Of String)
        Public Property Confidence As Double    ' 0~1，用于决策层
    End Class

    ''' <summary>
    ''' 统一能力注册中心（单例）
    ''' 所有 MCP 工具 + Skill 都注册到这里，提供统一的匹配和执行入口
    ''' </summary>
    Public Class McpCapabilityRegistry
        Private Shared _instance As McpCapabilityRegistry
        Private Shared _lock As New Object()
        Private _capabilities As New Dictionary(Of String, CapabilityDescriptor)(StringComparer.OrdinalIgnoreCase)
        Private _allMatchText As New List(Of (Descriptor As CapabilityDescriptor, Text As String))()
        Private _sourceNames As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        Public Shared ReadOnly Property Instance As McpCapabilityRegistry
            Get
                If _instance Is Nothing Then
                    SyncLock _lock
                        If _instance Is Nothing Then
                            _instance = New McpCapabilityRegistry()
                        End If
                    End SyncLock
                End If
                Return _instance
            End Get
        End Property

        Private Sub New()
        End Sub

        ''' <summary>
        ''' 注册一个 MCP 工具
        ''' </summary>
        Public Sub RegisterMcpTool(tool As McpTool,
                                    Optional displayName As String = "",
                                    Optional category As CapabilityCategory = CapabilityCategory.Utility,
                                    Optional keywords As List(Of String) = Nothing,
                                    Optional isWriteOperation As Boolean = False,
                                    Optional isDangerous As Boolean = False,
                                    Optional source As String = "excel-local",
                                    Optional dangerSeverity As String = Nothing,
                                    Optional dangerMessage As String = Nothing)
            If tool Is Nothing OrElse String.IsNullOrEmpty(tool.Name) Then Return

            Dim desc = New CapabilityDescriptor With {
                .Id = tool.Name,
                .Kind = CapabilityKind.McpTool,
                .Source = source,
                .DisplayName = If(String.IsNullOrEmpty(displayName), tool.Name, displayName),
                .Description = tool.Description,
                .Category = category,
                .Keywords = keywords,
                .IsWriteOperation = isWriteOperation,
                .InputSchema = tool.InputSchema,
                .IsDangerous = isDangerous,
                .DangerSeverity = dangerSeverity,
                .DangerMessage = dangerMessage
            }

            SyncLock _lock
                _capabilities(tool.Name) = desc
                _allMatchText.Add((desc, desc.ToMatchText()))
            End SyncLock
        End Sub

        ''' <summary>
        ''' 批量注册一整套 MCP 工具（推荐入口，加新工具集只用调这一个）
        ''' </summary>
        ''' <param name="tools">工具列表</param>
        ''' <param name="source">来源标识（如 excel-local / ppt-local / computer-local）</param>
        ''' <param name="displayName">来源显示名（如 "Excel本地操作"）</param>
        ''' <param name="category">默认分类</param>
        ''' <param name="isWritePrefix">写操作前缀（如 "excel_set_"，用于自动标记 isWriteOperation）</param>
        ''' <param name="isDangerousNames">危险工具名列表（如 execute_command）</param>
        Public Sub RegisterToolSet(tools As List(Of McpTool),
                                    source As String,
                                    displayName As String,
                                    Optional category As CapabilityCategory = CapabilityCategory.Utility,
                                    Optional isWritePrefix As String() = Nothing,
                                    Optional isDangerousNames As String() = Nothing)
            If tools Is Nothing OrElse tools.Count = 0 Then Return
            If Not String.IsNullOrEmpty(source) Then
                _sourceNames(source) = displayName
            End If
            Dim writeSet As New HashSet(Of String)(If(isWritePrefix, New String() {}))
            Dim dangerSet As New HashSet(Of String)(If(isDangerousNames, New String() {}))
            For Each tool In tools
                Dim isWrite = writeSet.Contains(tool.Name) OrElse tool.Name.EndsWith("_set") OrElse tool.Name.EndsWith("_update") OrElse tool.Name.EndsWith("_delete") OrElse tool.Name.EndsWith("_add") OrElse tool.Name.StartsWith("computer_mouse") OrElse tool.Name.StartsWith("computer_key") OrElse tool.Name = "computer_type_text"
                Dim isDanger = dangerSet.Contains(tool.Name)
                RegisterMcpTool(tool,
                    displayName:=tool.Name,
                    category:=category,
                    keywords:=New List(Of String) From {tool.Name},
                    isWriteOperation:=isWrite,
                    isDangerous:=isDanger,
                    source:=source)
            Next
        End Sub

        ''' <summary>
        ''' source → 显示名翻译（替代各处硬编码 Case）
        ''' </summary>
        Public Function TranslateSource(source As String) As String
            If String.IsNullOrEmpty(source) Then Return source
            Dim name As String = Nothing
            If _sourceNames.TryGetValue(source, name) Then Return name
            Return source
        End Function

        ''' <summary>
        ''' 注册所有内置本地工具集（真源 — 加新工具只改这一处）
        ''' </summary>
        Public Shared Sub RegisterAllBuiltInToolSets()
            Dim reg = Instance
            ' 先登记所有 source → 显示名映射（供 TranslateSource 统一查）
            reg._sourceNames("excel-local") = "Excel本地操作"
            reg._sourceNames("ppt-local") = "PPT演示文稿"
            reg._sourceNames("computer-local") = "电脑操作Agent"
            reg._sourceNames("web-local") = "联网搜索"
            reg._sourceNames("time-local") = "系统时间"
            ' 再注册各工具集（每个工具集自己带 keywords metadata）
            ExcelMcpTools.RegisterAllToolsToRegistry()
            PptMcpTools.RegisterAllToolsToRegistry()
            ComputerAgentTools.RegisterAllTools()
            WebSearchTools.RegisterAllTools()
            TimeTools.RegisterAllTools()
        End Sub

        ''' <summary>
        ''' 从注册中心直接输出前端需要的完整工具列表（McManager/McpBridge 统一走这里）
        ''' </summary>
        Public Function GetAllLocalToolsForUi() As List(Of Object)
            Dim result As New List(Of Object)()
            Dim caps = GetAllCapabilities().Where(Function(c) c.Kind = CapabilityKind.McpTool).ToList()
            For Each c In caps
                Dim params As New List(Of Object)()
                If c.InputSchema IsNot Nothing AndAlso c.InputSchema.Properties IsNot Nothing Then
                    Dim reqSet As New HashSet(Of String)(If(c.InputSchema.Required, Enumerable.Empty(Of String)()))
                    For Each kvp In c.InputSchema.Properties
                        params.Add(New With {
                            .name = kvp.Key,
                            .type = kvp.Value.Type,
                            .description = kvp.Value.Description,
                            .required = reqSet.Contains(kvp.Key)
                        })
                    Next
                End If
                result.Add(New With {
                    .name = c.Id,
                    .source = c.Source,
                    .sourceDisplayName = TranslateSource(c.Source),
                    .description = c.Description,
                    .category = c.Category.ToString(),
                    .parameters = params,
                    .isWrite = c.IsWriteOperation,
                    .enabled = True
                })
            Next
            Return result
        End Function
        Public Sub RegisterSkill(skillName As String,
                                  displayName As String,
                                  description As String,
                                  category As CapabilityCategory,
                                  Optional keywords As List(Of String) = Nothing,
                                  Optional isWriteOperation As Boolean = False)
            If String.IsNullOrEmpty(skillName) Then Return

            Dim desc = New CapabilityDescriptor With {
                .Id = "skill__" & skillName,
                .Kind = CapabilityKind.Skill,
                .DisplayName = If(String.IsNullOrEmpty(displayName), skillName, displayName),
                .Description = description,
                .Category = category,
                .Keywords = keywords,
                .IsWriteOperation = isWriteOperation,
                .SkillName = skillName
            }

            SyncLock _lock
                _capabilities(desc.Id) = desc
                _allMatchText.Add((desc, desc.ToMatchText()))
            End SyncLock
        End Sub

        ''' <summary>
        ''' 根据用户消息匹配最佳能力（后端统一匹配，不依赖前端 ngram）
        ''' 返回 top-N 匹配结果，按分数降序
        ''' </summary>
        Public Function Match(userMessage As String, Optional topN As Integer = 5) As List(Of CapabilityMatchResult)
            Dim results As New List(Of CapabilityMatchResult)()
            If String.IsNullOrWhiteSpace(userMessage) Then Return results

            Dim msg As String = userMessage.ToLowerInvariant().Trim()

            SyncLock _lock
                For Each entry In _allMatchText
                    Dim desc = entry.Descriptor
                    Dim matchText = entry.Text
                    If String.IsNullOrEmpty(matchText) Then Continue For

                    Dim score As Double = 0
                    Dim matchedKw As New List(Of String)()

                    ' 1) 关键词精确命中（权重最高：每个关键词 +20）
                    If desc.Keywords IsNot Nothing Then
                        For Each kw In desc.Keywords
                            If String.IsNullOrEmpty(kw) Then Continue For
                            If msg.Contains(kw.ToLowerInvariant()) Then
                                score += 20
                                matchedKw.Add(kw)
                            End If
                        Next
                    End If

                    ' 2) DisplayName 完全命中 +100
                    If Not String.IsNullOrEmpty(desc.DisplayName) Then
                        Dim dn = desc.DisplayName.ToLowerInvariant()
                        If msg.Contains(dn) Then
                            score += 100
                            matchedKw.Add(desc.DisplayName)
                        Else
                            ' N-gram 部分命中（2~4字滑动窗口）
                            score += NgramOverlap(dn, msg) * 3
                        End If
                    End If

                    ' 3) Description N-gram 命中
                    If Not String.IsNullOrEmpty(desc.Description) Then
                        score += NgramOverlap(desc.Description.ToLowerInvariant(), msg) * 1
                    End If

                    ' 4) 工具名拆分命中（computer_screen_capture → screen, capture；excel_remove_duplicates → remove, duplicates）
                    Dim namePart = desc.Id
                    For Each pfx In {"excel_", "ppt_", "computer_", "web_", "time_", "skill__"}
                        namePart = namePart.Replace(pfx, "")
                    Next
                    Dim nameParts = namePart.Split("_"c)
                    For Each part In nameParts
                        If part.Length >= 3 AndAlso msg.Contains(part.ToLowerInvariant()) Then
                            score += 2
                        End If
                    Next

                    If score > 0 Then
                        results.Add(New CapabilityMatchResult With {
                            .Capability = desc,
                            .Score = score,
                            .MatchedKeywords = matchedKw,
                            .Confidence = Math.Min(1.0, score / 150.0)  ' 归一化到 0~1
                        })
                    End If
                Next
            End SyncLock

            Return results.OrderByDescending(Function(r) r.Score).Take(topN).ToList()
        End Function

        ''' <summary>
        ''' 获取单个能力
        ''' </summary>
        Public Function GetCapability(id As String) As CapabilityDescriptor
            SyncLock _lock
                Dim result As CapabilityDescriptor = Nothing
                _capabilities.TryGetValue(id, result)
                Return result
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有能力
        ''' </summary>
        Public Function GetAllCapabilities() As List(Of CapabilityDescriptor)
            SyncLock _lock
                Return _capabilities.Values.ToList()
            End SyncLock
        End Function

        ''' <summary>
        ''' 获取所有能力的 JSON（供前端 toolRegistry 使用）
        ''' </summary>
        Public Function GetAllCapabilitiesJson() As String
            Dim list = GetAllCapabilities().Select(Function(c) New With {
                .name = c.Id,
                .displayName = c.DisplayName,
                .description = c.Description,
                .category = c.Category.ToString(),
                .kind = c.Kind.ToString(),
                .isWriteOperation = c.IsWriteOperation,
                .isDangerous = c.IsDangerous,
                .dangerSeverity = c.DangerSeverity,
                .dangerMessage = c.DangerMessage,
                .keywords = c.Keywords
            }).ToList()
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .capabilities = list,
                .count = list.Count
            })
        End Function

        ''' <summary>
        ''' 匹配 JSON API（供前端 HTTP 调用）
        ''' </summary>
        Public Function MatchJson(userMessage As String, Optional topN As Integer = 5) As String
            Dim matches = Match(userMessage, topN)
            Dim list = matches.Select(Function(m) New With {
                .id = m.Capability.Id,
                .displayName = m.Capability.DisplayName,
                .description = m.Capability.Description,
                .category = m.Capability.Category.ToString(),
                .kind = m.Capability.Kind.ToString(),
                .score = m.Score,
                .confidence = m.Confidence,
                .isWriteOperation = m.Capability.IsWriteOperation,
                .matchedKeywords = m.MatchedKeywords,
                .inputSchema = If(m.Capability.InputSchema IsNot Nothing,
                    New With {
                        .properties = If(m.Capability.InputSchema.Properties IsNot Nothing,
                            m.Capability.InputSchema.Properties.ToDictionary(
                                Function(kvp) kvp.Key,
                                Function(kvp) New With {
                                    .type = kvp.Value.Type,
                                    .description = kvp.Value.Description
                                }),
                            Nothing),
                        .required = m.Capability.InputSchema?.Required
                    },
                    Nothing)
            }).ToList()
            Return JsonConvert.SerializeObject(New With {
                .success = True,
                .matches = list,
                .query = userMessage
            })
        End Function

#Region "N-gram 辅助"

        Private Shared Function NgramOverlap(text As String, query As String) As Double
            If String.IsNullOrEmpty(text) OrElse String.IsNullOrEmpty(query) Then Return 0
            Dim score As Double = 0
            ' 从 text 中提取 2~4 字 ngram，检查是否在 query 中
            Dim cleaned = Regex.Replace(text, "[^\u4e00-\u9fa5a-zA-Z0-9]", "")
            Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
            For n = 2 To 4
                For i = 0 To cleaned.Length - n
                    Dim ng = cleaned.Substring(i, n)
                    If seen.Add(ng) AndAlso query.Contains(ng) Then
                        score += n
                    End If
                Next
            Next
            Return score
        End Function

#End Region

    End Class

End Namespace