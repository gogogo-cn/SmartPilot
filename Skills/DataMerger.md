---
hosts: [excel]
name: DataMerger
displayName: 数据合并助手
description: 将多个 Sheet 或多个 Excel/CSV 文件按行合并到一个工作表或工作簿，支持去重合并。不做数据拆分，不做跨表关联，不做数据清洗，不做格式美化。
keywords:
  - 合并
  - 数据合并
  - 多文件合并
  - 合并Sheet
  - 合并工作表
  - 合并工作簿
  - 合并表格
  - 合并数据
  - 去重合并
  - 合并去重
  - 汇总合并
  - 合并多表
  - 合并到一个表
  - 合并到一张表
  - 合并所有Sheet
  - 把多个文件合并到一起
  - 把所有Sheet合成一个
  - 横向汇总
  - 两张表放一起
examples:
  - 合并所有Sheet
  - 把多个文件合并到一起
  - 去重合并选中区域表格
  - 合并销售数据并去重
  - 合并多个Sheet到一个表并去重
  - 把多个CSV合并成一个工作簿
  - 按部门拆分这张表（反例：应走DataSplitter）
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
      - merge_sheets
      - merge_files
    required: true
  - name: keyColumn
    label: 去重关键列
    type: text
    required: false
pipeline: skill
subSkills:
  - SmartDataCleaner
  - SmartTableBeautifier
processingMode: Hybrid
# 【执行模式】不声明 executionMode → 走确定性后端 MCP 执行（Hybrid Phase 2）：
#   CustomSkillExecutor 检测到 DataMerger 后直接调用后端 MCP 工具
#   （excel_merge_sheets / excel_merge_workbooks）确定性完成任务。
#   注意：不要改成前端 agent 模式——前端 agent 工具集不包含 excel_merge_sheets 工具，
#   强行 agent 会报"未知工具: excel_merge_sheets"且无法实际执行。
allowMcpTools: true
supportsImageInput: false
supportsFileInput: true
outputFormat: text
isCustom: false
tags:
  - 数据合并
  - 批量处理
quickPrompts:
  - label: 清洗合并结果
    prompt: 对合并后的数据做清洗，处理空值和格式问题
  - label: 美化合并结果
    prompt: 美化合并后的表格样式
  - label: 分析合并结果
    prompt: 对合并后的数据做统计分析
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 数据合并助手

【重要实现说明】
- 本 Skill **强制通过 MCP 工具完成**，绝不执行任何可能清空用户原始数据的 VBA 模板代码。
- 后端拦截逻辑：`CustomSkillExecutor.vb` 在 Hybrid Phase 2 识别到 DataMerger 后，会跳过 VBA 模板注入直接调用 MCP 工具（excel_merge_sheets / excel_merge_workbooks）。
- MCP 工具的安全保证：**所有操作不会修改或清空源工作表或用户已选中的数据**。结果始终写入新建的独立工作表或独立文件。若目标表（如"合并结果"）已存在且有内容，会自动重命名旧表为备份，绝不覆盖。

## System Prompt

你是 Excel 数据合并专家。帮助用户把多份数据按行合并成一份。

【支持操作】
1. **多Sheet合并** (operationType=merge_sheets): 将同一工作簿中的多个 Sheet 按行合并到新工作表。对应 MCP 工具：excel_merge_sheets
2. **多文件合并** (operationType=merge_files): 将多个 Excel/CSV 文件合并到一个工作簿。对应 MCP 工具：excel_merge_workbooks

【阶段 1 输出规则 - 你只需输出结构化 JSON】
- 你的任务是根据用户输入和 Excel 数据上下文，输出一个合法 JSON，包含以下字段：
  {
    "operationType": "merge_sheets | merge_files（二选一）",
    "userOriginalIntent": "<一句话概括用户原始意图>",
    "confidence": <0-1整数>,
    "deduplicate": <true|false，用户要求"去重/去重复/去重合并/合并去重"时为 true，否则 false>,
    "mergeAll": <true|false，用户要求"合并所有/所有Sheet/全部表/所有工作表"时为 true，否则 false>,
    "parameters": {
      // merge_sheets: targetSheet, includeHeaders, sourceSheets(可选,数组)
      // merge_files: filePaths(数组), targetPath(可选), sheetPrefix(可选)
    }
  }
- 同时，为兼容旧版参数解析，顶层直接展开 parameters 的字段（如 keyColumn、targetSheet 等）。
- **不要输出 VBA 代码**，不要输出 markdown 代码块，不要输出多余解释性文字。
- 只输出 JSON。

## Template Code

```vba
' ============================================================
' 【安全锁定】DataMerger 不再执行任何 VBA 模板代码
' ------------------------------------------------------------
' 本 Skill 已完全迁移至 MCP 工具：
'   excel_merge_sheets / excel_merge_workbooks
' CustomSkillExecutor.vb 在 Hybrid Phase 2 检测到
' DataMerger 后会自动跳过此处，调用 MCP 工具完成任务。
' 以下代码是"安全空壳"，以防万一配置错误被调用时
' 不会执行任何破坏性操作（如 Cells.Clear）。
' ============================================================
' 此模板仅作参考，实际合并逻辑由 MCP 工具 excel_merge_sheets / excel_merge_workbooks 完成
```
