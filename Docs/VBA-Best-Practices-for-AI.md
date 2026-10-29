# 🤖 AI VBA 代码生成 - 安全编程规范

> **重要**: 本文档为 SmartPilot AI 系统提供 VBA 代码生成的安全标准
> 
> **目标**: 确保 AI 生成的代码**零错误、可执行、符合防御性编程原则**
>
> ⚠️ **2026-10-13 更新**：文中执行引擎原为 Node 桥 `local-bridge-server.js`（已删除），
> 现为 `Tools/py-excel-bridge.py`；`vba-validator.js` 不变。

---

## ✅ 核心原则

### 1️⃣ **永远不要信任用户输入**
```vba
' ❌ 错误做法 - 直接使用单元格值
ws.Name = ws.Range("A1").Value

' ✅ 正确做法 - 验证后使用
Dim newName As String
newName = GetSafeCellText(ws.Range("A1"))
If ValidateSheetName(newName, errorMsg) Then
    ws.Name = newName
Else
    Debug.Print "无效名称: " & errorMsg
End If
```

### 2️⃣ **必须包含错误处理**
```vba
Sub SafeOperation()
    On Error GoTo ErrorHandler
    
    ' ... 业务逻辑 ...
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "错误 [" & Err.Number & "]: " & Err.Description
    ' 可选：恢复操作或通知用户
End Sub
```

### 3️⃣ **所有变量必须声明类型**
```vba
Option Explicit  ' ← 必须在模块顶部

Sub Example()
    Dim ws As Worksheet        ' ✅ 明确类型
    Dim sheetName As String     ' ✅ 明确类型
    Dim counter As Integer      ' ✅ 明确类型
    
    ' ❌ 避免: Dim x, y, z（Variant 类型，性能差）
End Sub
```

---

## 🔒 危险操作安全清单

### ⚠️ 工作表重命名（最高频错误场景）

**必须使用的完整模板：**

```vba
Option Explicit

Sub BatchRenameSheets_Safe()
    Dim ws As Worksheet
    Dim cellValue As String
    Dim cleanName As String
    Dim errMsg As String
    Dim successCount As Integer
    Dim failCount As Integer
    
    successCount = 0
    failCount = 0
    
    Application.ScreenUpdating = False  ' 提高性能
    
    For Each ws In ThisWorkbook.Worksheets
        ' 步骤1: 安全获取单元格值
        cellValue = GetSafeCellText(ws.Range("A1"), "Sheet")
        
        ' 步骤2: 清理非法字符
        cleanName = CleanWorksheetName(cellValue)
        
        ' 步骤3: 执行安全重命名
        If SafeRenameWorksheet(ws, cleanName, errMsg) Then
            successCount = successCount + 1
            Debug.Print "✅ 成功: " & ws.Name & " -> " & cleanName
        Else
            failCount = failCount + 1
            Debug.Print "❌ 失败 (" & ws.Name & "): " & errMsg
        End If
    Next ws
    
    Application.ScreenUpdating = True
    
    MsgBox "完成！" & vbCrLf & _
           "成功: " & successCount & " 个" & vbCrLf & _
           "失败: " & failCount & " 个", _
           vbInformation, "批量重命名结果"
End Sub

' ============================================================
' 【必需辅助函数】以下函数必须在每个VBA模块中包含
' ============================================================

Private Function GetSafeCellText(ByVal rng As Range, Optional ByVal defaultValue As String = "") As String
    On Error Resume Next
    If rng Is Nothing Then
        GetSafeCellText = defaultValue
        Exit Function
    End If
    
    If IsEmpty(rng.Value) Or IsError(rng.Value) Or IsNull(rng.Value) Then
        GetSafeCellText = defaultValue
    Else
        GetSafeCellText = CStr(Trim(rng.Value))
    End If
    On Error GoTo 0
End Function

Private Function CleanWorksheetName(ByVal nameStr As String) As String
    Dim illegalChars As Variant
    Dim i As Integer
    
    nameStr = Trim(nameStr)
    
    If Len(nameStr) = 0 Then
        CleanWorksheetName = ""
        Exit Function
    End If
    
    If Len(nameStr) > 31 Then
        nameStr = Left(nameStr, 31)
        Debug.Print "⚠️ 名称已截断至31字符"
    End If
    
    illegalChars = Array("/", "\", "[", "]", ":", "*", "?", "|")
    
    For i = LBound(illegalChars) To UBound(illegalChars)
        nameStr = Replace(nameStr, illegalChars(i), "_")
    Next i
    
    CleanWorksheetName = nameStr
End Function

Private Function WorksheetExists(ByVal wsName As String, Optional ByVal wb As Workbook = Nothing) As Boolean
    Dim ws As Worksheet
    Dim targetWb As Workbook
    
    On Error Resume Next
    If wb Is Nothing Then Set targetWb = ThisWorkbook Else Set targetWb = wb
    Set ws = targetWb.Worksheets(wsName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function SafeRenameWorksheet(ByVal ws As Worksheet, ByVal newName As String, ByRef errorMsg As String) As Boolean
    Dim cleanName As String
    
    On Error GoTo ErrorHandler
    
    cleanName = CleanWorksheetName(newName)
    
    If Len(cleanName) = 0 Then
        errorMsg = "工作表名称不能为空"
        SafeRenameWorksheet = False
        Exit Function
    End If
    
    If StrComp(cleanName, ws.Name, vbTextCompare) = 0 Then
        errorMsg = ""
        SafeRenameWorksheet = True
        Exit Function
    End If
    
    If WorksheetExists(cleanName) Then
        errorMsg = "工作表 '" & cleanName & "' 已存在"
        SafeRenameWorksheet = False
        Exit Function
    End If
    
    ws.Name = cleanName
    errorMsg = ""
    SafeRenameWorksheet = True
    Exit Function
    
ErrorHandler:
    errorMsg = "重命名失败: " & Err.Description
    SafeRenameWorksheet = False
End Function
```

---

## 📋 常见错误模式及修复

### 错误 #1: 运行时错误 1004（工作表名称）
**触发条件**: `.Name =` 赋值时值不合法

**修复方案**: 使用上面的 `SafeRenameWorksheet()` 函数

---

### 错误 #2: 下标越界 (Error 9)
**触发条件**: 访问不存在的工作表/工作簿

**修复方案**:
```vba
' ❌ 直接访问
Set ws = Workbooks("数据.xlsx").Worksheets("报表")

' ✅ 安全访问
If WorkbookExists("数据.xlsx") Then
    If WorksheetExists("报表", Workbooks("数据.xlsx")) Then
        Set ws = Workbooks("数据.xlsx").Worksheets("报表")
    End If
End If
```

---

### 错误 #3: 类型不匹配 (Error 13)
**触发条件**: 单元格值为空或非文本

**修复方案**: 使用 `GetSafeCellText()` 或 `IsNumeric()` 检查

---

### 错误 #4: 对象变量或With块未设置 (Error 91)
**触发条件**: Range/Worksheet 对象为 Nothing

**修复方案**:
```vba
If Not rng Is Nothing Then
    ' 安全操作
End If
```

---

## 🛡️ 必须遵循的编码规范

### ✅ DO（必须做）

1. **始终使用 `Option Explicit`**
2. **每个 Sub/Function 都要有错误处理**
3. **循环内操作要检查状态**
4. **文件/对象操作前验证存在性**
5. **用户输入必须清理和验证**
6. **添加调试日志 (`Debug.Print`)**

### ❌ DON'T（禁止做）

1. **禁止 `On Error Resume Next` 全局屏蔽错误**
   - 允许临时使用，但必须紧跟恢复语句
2. **禁止直接使用 `.Value` 不做空值检查**
3. **禁止硬编码路径或名称**
4. **禁止不释放 COM 对象**
5. **禁止忽略返回值**

---

## 🎯 代码质量评分标准

系统会自动对 AI 生成的代码进行评分：

| 评分 | 等级 | 标准 |
|------|------|------|
| **A+** | 🟢 完美 | 包含所有安全检查、错误处理、辅助函数 |
| **A** | 🟢 优秀 | 有基本验证和错误处理 |
| **B** | 🟡 良好 | 有部分验证，但不够完善 |
| **C** | 🟠 及格 | 能运行但缺乏保护 |
| **F** | 🔴 不合格 | 存在严重安全隐患 |

**目标**: 所有生成代码必须达到 **A级以上**

---

## 📝 AI Prompt 模板

当用户请求生成 VBA 代码时，AI 应：

```
用户需求: [具体描述]

请生成满足以下要求的 VBA 代码：
1. ✅ Option Explicit 在顶部
2. ✅ 所有变量明确声明类型
3. ✅ 完整的错误处理（On Error GoTo）
4. ✅ 用户输入验证（空值、长度、非法字符）
5. ✅ 对象存在性检查（WorksheetExists 等）
6. ✅ 辅助函数完整实现
7. ✅ Debug.Print 日志输出
8. ✅ 性能优化（ScreenUpdating=False）
9. ✅ 用户友好提示（MsgBox 结果汇总）

代码风格要求：
- 缩进清晰（4空格）
- 注释充分（中文）
- 函数职责单一
- 变量命名语义化
```

---

## 🔧 集成到系统的说明

本规范已集成到以下组件中：

1. **[vba-validator.js](../Tools/vba-validator.js)** - 自动检测并修复问题代码
2. **[py-excel-bridge.py](../Tools/py-excel-bridge.py)** - 执行前强制验证
3. **执行日志** - `Logs/VBA-Execution/` 目录记录所有执行历史

### 使用流程

```
AI生成代码 → vba-validator自动验证 → 
  ├─ 通过 → 执行并记录日志
  ├─ 中风险 → 注入安全检查后执行
  └─ 高风险 → 阻止执行，提示用户确认
```

---

## 📊 监控与维护

### 执行日志分析
定期检查 `Logs/VBA-Execution/` 目录：
- 统计成功率/失败率
- 分析常见错误模式
- 优化验证规则

### 规则更新
当发现新的错误模式时：
1. 更新 [vba-validator.js](../Tools/vba-validator.js) 的规则库
2. 更新本文档的示例代码
3. 通知团队新规范

---

**版本**: 2.0.0  
**最后更新**: 2026-06-28  
**维护者**: SmartPilot 开发团队