---
hosts: [excel]
name: DataSplitter
displayName: 数据拆分助手
description: 按某一列的字段值或按固定行数，将数据拆分到多个 Sheet 或文件。不做数据合并，不做跨表关联，不做数据清洗，不做格式美化。
keywords:
  - 拆分
  - 数据拆分
  - 按列拆分
  - 按部门拆分
  - 按字段拆分
  - 分类拆分
  - 拆分工作表
  - 拆分工作簿
  - 拆分表格
  - 拆分成多个表
  - 拆分到不同sheet
  - 拆分到多个文件
  - 按行数拆分
  - 每n行拆分
  - 每1000行拆分
  - 每1000行
  - 按行拆分
  - 按行数
  - 拆分成多个文件
  - 拆成多份
  - 拆文件
  - 分批
  - 切片
  - 数据切片
  - 分割数据
  - 把一个表拆成多份
examples:
  - 按部门拆分成不同Sheet
  - 每1000行拆一个文件
  - 按城市拆分到多个文件
  - 前500行存成一个表
  - 按省份把一张表拆成多张
  - 把这个大表按日期拆成每天一个文件
  - 合并这些数据（反例：应走DataMerger）
  - 关联这两个表（反例：应走TableJoiner）
  - 清洗这些数据（反例：应走SmartDataCleaner）
  - 批量重命名工作表（反例：应走WorkbookBatchOps）
isGeneric: true
category: 数据处理
needsConfig: true
configFields:
  - name: operationType
    label: 操作类型
    type: select
    options:
      - split_by_column
      - split_by_rows
    required: true
  - name: keyColumn
    label: 拆分列
    type: text
    required: false
pipeline: skill
subSkills:
  - SmartDataCleaner
  - SmartTableBeautifier
processingMode: Hybrid
# 【执行模式】不声明 executionMode → 走确定性后端 MCP 执行（Hybrid Phase 2）：
#   CustomSkillExecutor 检测到 DataSplitter 后直接调用后端 MCP 工具
#   （excel_split_by_column / excel_split_by_rows）确定性完成任务。
#   注意：不要改成前端 agent 模式——前端 agent 工具集不包含 excel_split_by_rows 等工具，
#   强行 agent 会报"未知工具"且无法实际执行。
allowMcpTools: true
supportsImageInput: false
supportsFileInput: true
outputFormat: text
isCustom: false
tags:
  - 数据拆分
  - 批量处理
quickPrompts:
  - label: 按部门拆分
    prompt: 按部门把当前数据拆分到多个工作表
  - label: 按城市拆分
    prompt: 按城市把当前数据拆分到多个工作表
  - label: 按行数拆分
    prompt: 把当前数据按每1000行拆分成多个文件
  - label: 按日期拆分
    prompt: 按日期把当前数据拆分成每天一个工作表
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 数据拆分助手

【重要实现说明】
- 本 Skill **强制通过 MCP 工具完成**，绝不执行任何可能清空用户原始数据的 VBA 模板代码。
- 后端拦截逻辑：`CustomSkillExecutor.vb` 在 Hybrid Phase 2 识别到 DataSplitter 后，会跳过 VBA 模板注入直接调用 MCP 工具（excel_split_by_column / excel_split_by_rows）。
- MCP 工具的安全保证：**所有操作不会修改或清空源工作表或用户已选中的数据**。结果始终写入新建的独立工作表或独立文件。

## System Prompt

你是 Excel 数据拆分专家。帮助用户把一张表按字段值或按行数拆分成多份。

【支持操作】
1. **按列拆分** (operationType=split_by_column): 按某一列的唯一值，将数据拆分到不同 Sheet 或不同文件。对应 MCP 工具：excel_split_by_column
2. **按行数拆分** (operationType=split_by_rows): 按固定行数（如每1000行）拆分为多个 Sheet/文件。对应 MCP 工具：excel_split_by_rows

【阶段 1 输出规则 - 你只需输出结构化 JSON】
- 你的任务是根据用户输入和 Excel 数据上下文，输出一个合法 JSON，包含以下字段：
  {
    "operationType": "split_by_column | split_by_rows（二选一）",
    "userOriginalIntent": "<一句话概括用户原始意图>",
    "confidence": <0-1整数>,
    "parameters": {
      // split_by_column: column, outputMode("sheets"|"files"), outputDir(可选)
      // split_by_rows: rowsPerChunk(整数), outputMode("sheets"|"files"), outputDir(可选)
    }
  }
- 同时，为兼容旧版参数解析，顶层直接展开 parameters 的字段（如 keyColumn、rowsPerFile 等）。
- 对于 split_by_column，顶层必须包含 keyColumn 或 column。
- 对于 split_by_rows，顶层必须包含 rowsPerFile 或 rowsPerChunk。
- **不要输出 VBA 代码**，不要输出 markdown 代码块，不要输出多余解释性文字。
- 只输出 JSON。

## Template Code

```vba
' ============================================================
' 【安全锁定】DataSplitter 不再执行任何 VBA 模板代码
' ------------------------------------------------------------
' 本 Skill 已完全迁移至 MCP 工具：
'   excel_split_by_column / excel_split_by_rows
' CustomSkillExecutor.vb 在 Hybrid Phase 2 检测到
' DataSplitter 后会自动跳过此处，调用 MCP 工具完成任务。
' 以下代码是"安全空壳"，以防万一配置错误被调用时
' 不会执行任何破坏性操作（如 Cells.Clear）。
' ============================================================
' 此模板仅作参考，实际拆分逻辑由 MCP 工具 excel_split_by_column / excel_split_by_rows 完成
```
