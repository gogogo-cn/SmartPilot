Imports System
Imports System.Collections.Generic
Imports System.Linq

Namespace Services.Mcp

    ''' <summary>
    ''' 操作型任务的统一"步骤规划"记录器。
    ''' 提供统一的中文步骤模板，任何 skill 复用同一套措辞，前端通过结果里的 steps 数组渲染"步骤规划"清单。
    ''' 只负责记录步骤文案，不执行任何业务逻辑。
    ''' </summary>
    Public Class StepRecorder

        Private ReadOnly _steps As New List(Of String)()

        ''' <summary>已记录的步骤（返回副本，便于序列化与不可变使用）</summary>
        Public ReadOnly Property Steps As List(Of String)
            Get
                Return New List(Of String)(_steps)
            End Get
        End Property

        ''' <summary>当前步骤数</summary>
        Public ReadOnly Property Count As Integer
            Get
                Return _steps.Count
            End Get
        End Property

        ''' <summary>追加一条已格式化好的步骤（保留完全自定义文案的场景）</summary>
        Public Sub Add(text As String)
            If Not String.IsNullOrWhiteSpace(text) Then _steps.Add(text.Trim())
        End Sub

        ''' <summary>读取源表</summary>
        Public Sub ReadSource(name As String, rows As Integer)
            Add("读取源工作表 '" & name & "'，共 " & rows & " 行数据")
        End Sub

        ''' <summary>读取并核对多张源表（合并用）</summary>
        Public Sub ReviewSources(count As Integer, names As IEnumerable(Of String))
            Add("读取并核对源表 " & count & " 张：" & String.Join("、", names))
        End Sub

        ''' <summary>准备目标表</summary>
        Public Sub PrepareTarget(name As String, Optional desc As String = "已存在则清空复用，不残留空白表")
            Add("准备目标表 '" & name & "'（" & desc & "）")
        End Sub

        ''' <summary>写入表头</summary>
        Public Sub WriteHeader(rows As Integer)
            Add("写入表头（" & rows & " 行）")
        End Sub

        ''' <summary>从源表向目标末尾插入数据行</summary>
        Public Sub InsertFrom(source As String, dataRows As Integer, Optional startRow As Integer = -1)
            If startRow > 0 Then
                Add("插入源表 '" & source & "' 数据 " & dataRows & " 行（追加到第 " & startRow & " 行）")
            Else
                Add("插入源表 '" & source & "' 数据 " & dataRows & " 行")
            End If
        End Sub

        ''' <summary>新建工作表并写入数据</summary>
        Public Sub NewSheet(name As String, dataRows As Integer, Optional hasHeader As Boolean = True)
            Dim prefix = If(hasHeader, "表头 + ", "")
            Add("新建工作表 '" & name & "'，写入" & prefix & dataRows & " 行数据")
        End Sub

        ''' <summary>生成工作表文件并写入数据</summary>
        Public Sub GenerateFile(fileName As String, dataRows As Integer, Optional hasHeader As Boolean = True)
            Dim prefix = If(hasHeader, "表头 + ", "")
            Add("生成文件 " & fileName & "，写入" & prefix & dataRows & " 行数据")
        End Sub

        ''' <summary>按某列拆分为若干组</summary>
        Public Sub SplitIntoGroups(column As String, groupCount As Integer, groupNames As IEnumerable(Of String))
            Add("按拆分列 '" & column & "' 将数据分为 " & groupCount & " 组：" & String.Join("、", groupNames))
        End Sub

        ''' <summary>按固定行数拆分</summary>
        Public Sub SplitByRowCount(rowsPerChunk As Integer, chunkCount As Integer, totalDataRows As Integer)
            Add("按每 " & rowsPerChunk & " 行拆为 " & chunkCount & " 份，共 " & totalDataRows & " 行数据行")
        End Sub

        ''' <summary>读取左右两张关联表</summary>
        Public Sub JoinRead(leftName As String, leftRows As Integer, rightName As String, rightRows As Integer)
            Add("读取左表 '" & leftName & "'（" & leftRows & " 行）与右表 '" & rightName & "'（" & rightRows & " 行）")
        End Sub

        ''' <summary>关联匹配结果</summary>
        Public Sub JoinMatch(joinType As String, leftKeys As String, rightKeys As String, matchedRows As Integer)
            Add("以 " & joinType & " 关联按 [" & leftKeys & "] = [" & rightKeys & "] 匹配，得到 " & matchedRows & " 行")
        End Sub

        ''' <summary>写入目标表</summary>
        Public Sub WriteToTarget(name As String, totalRows As Integer)
            Add("写入目标表 '" & name & "'，共 " & totalRows & " 行")
        End Sub

    End Class

End Namespace