Imports System.Collections.Generic
Imports System.IO
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module ToolCoverage

    Sub RunAll()
        Dim base = AppContext.BaseDirectory
        TestJsonDefinitions(base)
        TestDispatchCoverage(base)
        TestBridgeMethods(base)
    End Sub

    Sub Assert(cond As Boolean, msg As String)
        If Not cond Then Throw New Exception("FAIL: " & msg)
        Console.WriteLine("  PASS: " & msg)
    End Sub

    Private Function LoadTools(path As String) As JArray
        Dim root = JsonConvert.DeserializeObject(Of JObject)(File.ReadAllText(path))
        Return CType(root("tools"), JArray)
    End Function

    ' 1) JSON 工具定义合法性：名称唯一、前缀正确、必填字段与参数齐全
    Sub TestJsonDefinitions(base As String)
        Console.WriteLine("[1] JSON 工具定义校验")
        For Each f In {"ppt_tools.json", "word_tools.json"}
            Dim arr = LoadTools(base & f)
            Dim names As New HashSet(Of String)
            For Each t As JToken In arr
                Dim nm = t("name").ToString()
                Assert(Not names.Contains(nm), f & " 名称唯一: " & nm)
                names.Add(nm)
                Assert(nm.StartsWith("ppt_") OrElse nm.StartsWith("word_"), f & " 前缀正确: " & nm)
                Assert(t("description") IsNot Nothing AndAlso t("description").ToString().Length > 0, nm & " 有 description")
                Assert(t("parameters") IsNot Nothing, nm & " 有 parameters")
                Assert(t("displayName") IsNot Nothing AndAlso t("displayName").ToString().Length > 0, nm & " 有 displayName")
                Assert(t("intentKeywords") IsNot Nothing AndAlso t("intentKeywords").Count() > 0, nm & " 有 intentKeywords")
                Assert(t("operationType") IsNot Nothing, nm & " 有 operationType")
                For Each p As JToken In t("parameters")
                    Assert(p("name") IsNot Nothing, nm & " 参数有 name")
                    Assert(p("type") IsNot Nothing, nm & " 参数有 type")
                    Assert(p("required") IsNot Nothing, nm & " 参数有 required")
                Next
            Next
            Console.WriteLine("  [OK] " & f & " 定义数=" & arr.Count)
        Next
    End Sub

    ' 2) 每个 JSON 工具在 HostMcpTools 都有对应 Case 分支
    Sub TestDispatchCoverage(base As String)
        Console.WriteLine("[2] dispatch 分支覆盖")
        Dim src = File.ReadAllText(base & "src/HostMcpTools.vb")
        Dim cases As New HashSet(Of String)
        For Each m As Match In Regex.Matches(src, "Case\s+""(ppt_\w+|word_\w+)""")
            cases.Add(m.Groups(1).Value)
        Next
        For Each f In {"ppt_tools.json", "word_tools.json"}
            Dim arr = LoadTools(base & f)
            For Each t As JToken In arr
                Dim nm = t("name").ToString()
                Assert(cases.Contains(nm), "dispatch 含 " & nm)
            Next
        Next
        Console.WriteLine("  [OK] HostMcpTools 所有 JSON 工具均有 dispatch 分支")
    End Sub

    ' 3) 每个 dispatch 分支调用的桥方法，在对应桥里都已实现（Public/Private/Overrides Function）
    Sub TestBridgeMethods(base As String)
        Console.WriteLine("[3] 桥方法实现覆盖")
        Dim dispatchSrc = File.ReadAllText(base & "src/HostMcpTools.vb")
        Dim pptBridge = File.ReadAllText(base & "src/PptBridge.vb")
        Dim wordBridge = File.ReadAllText(base & "src/WordBridge.vb")
        Dim lines = dispatchSrc.Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None)
        Dim curTool As String = Nothing
        For Each line In lines
            If line.Contains("Case Else") Then
                curTool = Nothing
                Continue For
            End If
            Dim mCase = Regex.Match(line, "Case\s+""(ppt_\w+|word_\w+)""")
            If mCase.Success Then
                curTool = mCase.Groups(1).Value
                Continue For
            End If
            Dim mCall = Regex.Match(line, "CallBridge\(""(\w+)"",")
            If mCall.Success AndAlso curTool IsNot Nothing Then
                Dim method = mCall.Groups(1).Value
                If curTool.StartsWith("ppt_") Then
                    Assert(Regex.IsMatch(pptBridge, "Function\s+" & Regex.Escape(method) & "\s*\("),
                           "PPT 桥实现 " & method & " (工具 " & curTool & ")")
                Else
                    Assert(Regex.IsMatch(wordBridge, "Function\s+" & Regex.Escape(method) & "\s*\("),
                           "Word 桥实现 " & method & " (工具 " & curTool & ")")
                End If
            End If
        Next
        Console.WriteLine("  [OK] 所有 dispatch 调用的桥方法均已实现")
    End Sub

End Module
