---
hosts: [excel]
name: TableJoiner
displayName: 跨表关联助手
description: 类似 SQL JOIN，根据关键列把两个表的列关联到同一张结果表，支持 inner/left/right/full。不做数据合并，不做数据拆分，不做数据清洗。
keywords:
  - 关联
  - 表关联
  - 跨表
  - 跨表匹配
  - JOIN
  - join
  - 连接表
  - 两表匹配
  - 数据匹配
  - VLOOKUP
  - vlookup
  - 左连接
  - 右连接
  - 内连接
  - 外连接
  - 全连接
  - 合并匹配
  - 按ID关联
  - 按关键列匹配
examples:
  - 把两个表按ID关联起来
  - 用商品编码匹配两个表
  - 左连接这两张表
  - vlookup匹配数据
  - 把表B的列对照表A补进去
  - 合并这些数据（反例：应走DataMerger）
  - 拆分这张表（反例：应走DataSplitter）
  - 清洗这些数据（反例：应走SmartDataCleaner）
isGeneric: true
category: 数据处理
needsConfig: true
configFields:
  - name: joinType
    label: 关联类型
    type: select
    options:
      - inner
      - left
      - right
      - full
    required: false
  - name: keyColumn
    label: 关键列
    type: text
    required: false
pipeline: skill
subSkills:
  - SmartDataCleaner
  - SmartTableBeautifier
processingMode: Hybrid
# 【执行模式】不声明 executionMode → 走确定性后端 MCP 执行（Hybrid Phase 2）：
#   CustomSkillExecutor 检测到 TableJoiner 后直接调用后端 MCP 工具
#   （excel_join_tables）确定性完成任务。
#   注意：不要改成前端 agent 模式——前端 agent 工具集不包含 excel_join_tables 工具，
#   强行 agent 会报"未知工具"且无法实际执行。
allowMcpTools: true
supportsImageInput: false
supportsFileInput: true
outputFormat: text
isCustom: false
tags:
  - 跨表关联
  - 数据匹配
quickPrompts:
  - label: 左连接
    prompt: 用左连接关联当前两个表
  - label: 内连接
    prompt: 用内连接关联当前两个表
  - label: 全连接
    prompt: 用全连接关联当前两个表
  - label: 按ID关联
    prompt: 按ID列关联当前两个表
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 跨表关联助手

【重要实现说明】
- 本 Skill **强制通过 MCP 工具完成**，绝不执行任何可能清空用户原始数据的 VBA 模板代码。
- 后端拦截逻辑：`CustomSkillExecutor.vb` 在 Hybrid Phase 2 识别到 TableJoiner 后，会跳过 VBA 模板注入直接调用 MCP 工具（excel_join_tables）。
- MCP 工具的安全保证：**所有操作不会修改或清空源工作表或用户已选中的数据**。结果始终写入新建的独立工作表（默认"关联结果"）。

## System Prompt

你是 Excel 跨表关联专家。帮助用户根据关键列把两张表的字段关联到一张结果表，类似 SQL JOIN。

【支持操作】
1. **跨表匹配** (operationType=join_tables): 根据关键列把左表与右表按指定关联类型匹配到同一张结果表。支持 inner/left/right/full。对应 MCP 工具：excel_join_tables

【阶段 1 输出规则 - 你只需输出结构化 JSON】
- 你的任务是根据用户输入和 Excel 数据上下文，输出一个合法 JSON，包含以下字段：
  {
    "operationType": "join_tables",
    "userOriginalIntent": "<一句话概括用户原始意图>",
    "confidence": <0-1整数>,
    "parameters": {
      // join_tables: leftSheet, rightSheet, leftKeyColumn, rightKeyColumn(可选), joinType(inner|left|right|full), rightColumns(可选,数组)
    }
  }
- 同时，为兼容旧版参数解析，顶层直接展开 parameters 的字段（如 leftSheet, rightSheet, joinKey 等）。
- 顶层必须包含 leftSheet, rightSheet, joinKey。
- **不要输出 VBA 代码**，不要输出 markdown 代码块，不要输出多余解释性文字。
- 只输出 JSON。

## Template Code

```vba
' ============================================================
' 【安全锁定】TableJoiner 不再执行任何 VBA 模板代码
' ------------------------------------------------------------
' 本 Skill 已完全迁移至 MCP 工具：
'   excel_join_tables
' CustomSkillExecutor.vb 在 Hybrid Phase 2 检测到
' TableJoiner 后会自动跳过此处，调用 MCP 工具完成任务。
' 以下代码是"安全空壳"，以防万一配置错误被调用时
' 不会执行任何破坏性操作（如 Cells.Clear）。
' ============================================================
' 此模板仅作参考，实际跨表关联逻辑由 MCP 工具 excel_join_tables 完成
```
