Option Strict On
Option Explicit On
Imports System
Imports System.Collections.Generic
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Namespace Services.Mcp

    ''' <summary>
    ''' 进程内系统时间工具集（单一职责：当前时间/日期）
    ''' 提供可调用的 time_now，让模型在需要"现在几号/几点/周几/是否工作日"时主动取真实系统时钟，
    ''' 不依赖提示词注入，也不靠训练知识臆造。自带时区与工作日判定。
    ''' </summary>
    Partial Public Class TimeTools

        Public Shared Function GetToolDefinitions() As List(Of McpTool)
            Return New List(Of McpTool) From {
                New McpTool("time_now", "获取当前的系统日期、时间、星期、时区与工作日状态，用于回答今天几号/几点/周几等问题",
                    New McpToolInputSchema() With {
                        .Properties = New Dictionary(Of String, McpToolProperty) From {
                            {"withSeconds", New McpToolProperty("boolean", "是否包含秒，默认 false")}
                        },
                        .Required = New List(Of String)()
                    })
            }
        End Function

        Public Shared Sub RegisterAllTools()
            For Each tool In GetToolDefinitions()
                McpCapabilityRegistry.Instance.RegisterMcpTool(tool,
                    displayName:="当前时间",
                    category:=CapabilityCategory.Utility,
                    keywords:=New List(Of String) From {
                        "time_now", "时间", "几点", "几号", "今天", "今天几号", "今天是", "日期",
                        "周几", "星期几", "今天星期几", "现在几点", "当前时间", "工作日", "星期"
                    },
                    isWriteOperation:=False,
                    isDangerous:=False,
                    source:="time-local")
            Next
        End Sub

        Public Shared Function ExecuteTool(toolName As String, args As Dictionary(Of String, Object)) As ToolsCallResult
            Select Case toolName
                Case "time_now"
                    Return Now2(args)
                Case Else
                    Return Err("未找到执行方法: " & toolName)
            End Select
        End Function

        Private Shared Function Now2(args As Dictionary(Of String, Object)) As ToolsCallResult
            Dim withSeconds = G(args, "withSeconds", False)
            Dim now = DateTime.Now
            Dim weekNames = {"日", "一", "二", "三", "四", "五", "六"}
            Dim offset = TimeZoneInfo.Local.GetUtcOffset(now)
            Dim iso = now.ToString(If(withSeconds, "yyyy-MM-dd'T'HH:mm:sszzz", "yyyy-MM-dd'T'HH:mmzzz"))
            Dim tzSign = If(offset.TotalHours >= 0, "+", "-")
            Dim tzAbs = Math.Abs(offset.TotalHours)

            Dim payload = New JObject() From {
                {"iso", iso},
                {"date", now.ToString("yyyy-MM-dd")},
                {"time", If(withSeconds, now.ToString("HH:mm:ss"), now.ToString("HH:mm"))},
                {"weekday", CInt(now.DayOfWeek)},
                {"weekdayName", "星期" & weekNames(CInt(now.DayOfWeek))},
                {"isWeekend", now.DayOfWeek = DayOfWeek.Saturday OrElse now.DayOfWeek = DayOfWeek.Sunday},
                {"timezone", String.Format("UTC{0}{1:00}:{2:00}", tzSign, CInt(tzAbs), CInt((tzAbs - CInt(tzAbs)) * 60))}
            }

            Return OK(payload.ToString(Formatting.Indented))
        End Function

        ' ── 与其它工具集一致的取参/返回约定 ──

        Private Shared Function G(Of T)(a As Dictionary(Of String, Object), n As String, d As T) As T
            If a Is Nothing Then Return d
            Dim v As Object = Nothing
            If a.TryGetValue(n, v) AndAlso v IsNot Nothing Then
                Try
                    Return CType(v, T)
                Catch
                    Try
                        Return DirectCast(Convert.ChangeType(v, GetType(T)), T)
                    Catch
                    End Try
                End Try
            End If
            Return d
        End Function

        Private Shared Function OK(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}
            }
        End Function

        Private Shared Function Err(s As String) As ToolsCallResult
            Return New ToolsCallResult() With {
                .IsError = True,
                .Content = New List(Of McpContentItem) From {New McpTextContent With {.Text = s}}
            }
        End Function

    End Class

End Namespace