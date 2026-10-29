---
hosts: [excel]
name: DataDiffAnalyzer
displayName: 数据差异对比分析器
description: 对比两个Sheet或区域，标记新增、删除、修改的单元格，生成差异摘要报告。不做数据合并，不做数据清洗，不做格式美化。
keywords:
  - 对比
  - 差异
  - 比较
  - 对账
  - 核对
  - 版本对比
examples:
  - 对比这两个表
  - 找出差异
  - 对账
  - 比较上周和本周数据
  - 合并这些数据（反例：应走DataMerger）
  - 清洗这些数据（反例：应走SmartDataCleaner）
isGeneric: true
category: 数据处理
needsConfig: true
configFields:
  - name: keyColumn
    label: 关键列
    type: text
    required: true
  - name: rangeA
    label: 对比区域A
    type: range
    required: true
  - name: rangeB
    label: 对比区域B
    type: range
    required: true
pipeline: skill
processingMode: Hybrid
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 数据对比
  - 差异分析
  - 对账
quickPrompts:
  - label: 换关键列对比
    prompt: 换一个关键列重新对比当前两个区域
  - label: 对比新增行
    prompt: 对比当前两个区域，重点标出新增的行
  - label: 对比修改项
    prompt: 对比当前两个区域，重点标出被修改的单元格
  - label: 导出差异报告
    prompt: 把当前两个区域的差异生成摘要报告
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 数据差异对比分析器

## System Prompt

你是数据对比专家。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际JSON对比结果输出，禁止只思考不输出。

【对比能力】
1. **行级对比**: 按关键列对齐，找出新增行、删除行
2. **单元格对比**: 标记具体修改过的单元格
3. **差异标记**: 用颜色区分差异类型（新增=绿色，删除=红色，修改=黄色）
4. **差异报告**: 生成摘要（共多少行变化、哪些字段被修改）

【输出格式】
输出 JSON 格式对比结果：
{
  "summary": {
    "totalRowsA": 100,
    "totalRowsB": 105,
    "addedRows": 8,
    "deletedRows": 3,
    "modifiedCells": 15
  },
  "keyColumn": "ID",
  "diffDetails": [
    {"rowKey": "001", "type": "modified", "changes": [{"column": "金额", "old": 100, "new": 120}]}
  ],
  "vbaCode": "应用条件格式的VBA代码"
}

【规则】
- 优先让用户指定关键列（如 ID、订单号）
- 若未指定，自动寻找最合适的唯一标识列
- 对比前检查两表结构是否一致
- 文本对比忽略前后空格，数字对比考虑精度

## Template Code

```vba
Sub DataDiffAnalysis()
    Dim wsA As Worksheet, wsB As Worksheet
    Dim wsResult As Worksheet
    Dim lastRowA As Long, lastRowB As Long
    Dim lastCol As Long
    Dim keyCol As Long
    Dim keyColLetter As String
    Dim i As Long, j As Long
    Dim keyA As String, keyB As String
    Dim found As Range
    Dim addedCount As Long, deletedCount As Long, modifiedCount As Long
    Dim resultRow As Long
    Dim bRow As Long
    
    On Error GoTo ErrorHandler
    
    ' 默认对比当前工作簿中的 Sheet1 和 Sheet2
    Set wsA = ThisWorkbook.Worksheets("Sheet1")
    Set wsB = ThisWorkbook.Worksheets("Sheet2")
    
    If wsA Is Nothing Or wsB Is Nothing Then
        MsgBox "未找到 Sheet1 或 Sheet2，请确保存在这两个工作表", vbExclamation
        Exit Sub
    End If
    
    keyColLetter = InputBox("请输入关键列列标（如 A、B）:", "差异对比", "A")
    If keyColLetter = "" Then Exit Sub
    keyCol = Range(keyColLetter & "1").Column
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 创建结果 Sheet
    On Error Resume Next
    Set wsResult = ThisWorkbook.Worksheets("差异报告")
    On Error GoTo ErrorHandler
    
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Worksheets.Add
        wsResult.Name = "差异报告"
    Else
        wsResult.Cells.Clear
    End If
    
    lastRowA = wsA.Cells(wsA.Rows.Count, keyCol).End(xlUp).Row
    lastRowB = wsB.Cells(wsB.Rows.Count, keyCol).End(xlUp).Row
    lastCol = wsA.UsedRange.Columns.Count
    
    ' 写入标题
    wsResult.Range("A1").Value = "差异对比报告"
    wsResult.Range("A1").Font.Size = 16
    wsResult.Range("A1").Font.Bold = True
    wsResult.Range("A3").Value = "类型"
    wsResult.Range("B3").Value = "关键值"
    wsResult.Range("C3").Value = "位置"
    wsResult.Range("D3").Value = "详情"
    wsResult.Range("A3:D3").Font.Bold = True
    resultRow = 4
    
    addedCount = 0
    deletedCount = 0
    modifiedCount = 0
    
    ' 检查 A 中新增到 B 的行
    For i = 2 To lastRowB
        keyB = CStr(wsB.Cells(i, keyCol).Value)
        If keyB <> "" Then
            Set found = wsA.Range(wsA.Cells(2, keyCol), wsA.Cells(lastRowA, keyCol)).Find(keyB, LookIn:=xlValues, LookAt:=xlWhole)
            If found Is Nothing Then
                wsResult.Cells(resultRow, 1).Value = "新增"
                wsResult.Cells(resultRow, 2).Value = keyB
                wsResult.Cells(resultRow, 3).Value = "Sheet2 第" & i & "行"
                resultRow = resultRow + 1
                addedCount = addedCount + 1
            End If
        End If
    Next i
    
    ' 检查 B 中删除的行
    For i = 2 To lastRowA
        keyA = CStr(wsA.Cells(i, keyCol).Value)
        If keyA <> "" Then
            Set found = wsB.Range(wsB.Cells(2, keyCol), wsB.Cells(lastRowB, keyCol)).Find(keyA, LookIn:=xlValues, LookAt:=xlWhole)
            If found Is Nothing Then
                wsResult.Cells(resultRow, 1).Value = "删除"
                wsResult.Cells(resultRow, 2).Value = keyA
                wsResult.Cells(resultRow, 3).Value = "Sheet1 第" & i & "行"
                resultRow = resultRow + 1
                deletedCount = deletedCount + 1
            End If
        End If
    Next i
    
    ' 检查修改的单元格
    For i = 2 To lastRowA
        keyA = CStr(wsA.Cells(i, keyCol).Value)
        If keyA <> "" Then
            Set found = wsB.Range(wsB.Cells(2, keyCol), wsB.Cells(lastRowB, keyCol)).Find(keyA, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                bRow = found.Row
                For j = 1 To lastCol
                    If j <> keyCol Then
                        If CStr(wsA.Cells(i, j).Value) <> CStr(wsB.Cells(bRow, j).Value) Then
                            wsResult.Cells(resultRow, 1).Value = "修改"
                            wsResult.Cells(resultRow, 2).Value = keyA
                            wsResult.Cells(resultRow, 3).Value = "列" & Split(wsA.Cells(1, j).Address, "$")(1)
                            wsResult.Cells(resultRow, 4).Value = "Sheet1: " & wsA.Cells(i, j).Value & " -> Sheet2: " & wsB.Cells(bRow, j).Value
                            resultRow = resultRow + 1
                            modifiedCount = modifiedCount + 1
                        End If
                    End If
                Next j
            End If
        End If
    Next i
    
    ' 写入摘要
    wsResult.Range("A2").Value = "新增: " & addedCount & " 行，删除: " & deletedCount & " 行，修改: " & modifiedCount & " 个单元格"
    wsResult.Columns.AutoFit
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "差异对比完成！" & vbCrLf & "新增: " & addedCount & " 行" & vbCrLf & _
           "删除: " & deletedCount & " 行" & vbCrLf & "修改: " & modifiedCount & " 个单元格", vbInformation
    Exit Sub
    
ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "对比失败: " & Err.Description, vbCritical
End Sub
```
