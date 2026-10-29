---
hosts: [excel]
name: SmartDataCleaner
displayName: 智能数据清洗大师
description: 自动检测并执行数据清洗：去重、空值填充/删除、日期/数字/文本格式统一、列拆分合并、去除空格、文本规范化。不分析异常原因，不生成统计报告，只做数据修复。
keywords:
  - 清洗
  - 清洗数据
  - 清洗表格
  - 数据清洗
  - 去重
  - 删除重复
  - 删除重复行
  - 去除重复
  - 去掉重复
  - 重复
  - 重复行
  - 高亮重复
  - 空值
  - 填充
  - 填充空值
  - 格式转换
  - 修复
  - 处理
  - 清理数据
  - 规范数据
  - 空行
  - 删除空行
  - 删空行
  - 清理空行
  - 去空行
  - 移除空行
  - 去掉空行
examples:
  - 清洗这份数据
  - 删除重复行
  - 填充空值
  - 统一日期格式
  - 分析异常值（反例：应走SmartDataAnalyzer）
  - 生成统计报告（反例：应走DescriptiveStatsReporter）
  - 校验数据质量、标出异常单元格（反例：应走DataQualityValidator，只检测不修复）
isGeneric: true
category: 数据处理
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
pipeline: skill
# SkillSpec 模式：LLM 输出结构化 spec（清洗操作数组：去重/空值/去空格/日期/数字/文本规范化），
# 由 skill-pipeline 执行：全部一次性交给 excel_clean_data（原生引擎）。removeDuplicates 默认【标记重复不删除】：
# 把重复行整行黄色高亮标记，不新增列、不删除行；如需真正删除重复行，请输出 spec：{"markDuplicate": false}。
# excel_clean_data 不可用/失败时回退 excel_read_range 读取 + JS 确定性变换 + excel_write_range 写回。
# 纯 MCP 工具链，不回退 VBA。LLM 负责选操作 → 灵活；MCP/模板负责正确 → 不报错。
# 注：原 subSkills(SmartTableBeautifier) 移除——清洗 spec 模式自带格式整理步骤，避免双重执行。
processingMode: SkillSpec
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 数据清洗
  - 数据质量
  - 自动化
quickPrompts:
  - label: 去重 · 仅高亮
    prompt: 对当前数据继续清洗，去重并只把重复行黄色高亮标记（不删除、不新增列）
  - label: 去重 · 删除重复行
    prompt: 对当前数据继续清洗，删除重复行（真正删除，仅保留首条唯一数据）
  - label: 填充空值
    prompt: 对当前数据继续清洗，填充空值
  - label: 去除空格
    prompt: 对当前数据继续清洗，去除文本前后空格
  - label: 统一日期格式
    prompt: 对当前数据继续清洗，统一日期格式
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 智能数据清洗大师

## System Prompt

你是数据清洗专家。

**职责边界**：本 Skill 只执行数据修复操作（去重、填充、格式转换、拆分列等）。不分析异常原因，不生成统计报告，不输出数据洞察。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际JSON清洗策略输出，禁止只思考不输出。

【分析维度】
1. **重复数据**: 检测完全重复或部分重复的行，执行去重。必须依据用户措辞区分去重模式：用户说“删除重复行/删除重复/去掉重复/清除重复/去重删除”等**明确要删除**时，`markDuplicate` 必须设为 `false`（真正删除重复行）；用户只说“去重/标记重复/标黄/高亮重复/找重复”等**仅标记**意图、或表述含糊时，`markDuplicate` 保持默认 `true`（只黄色高亮，不新增列、不删除），绝不擅自删除。
2. **空值/缺失值**: 统计空值分布，执行填充或删除
3. **格式不一致**: 日期格式混乱、数字存为文本、大小写不统一等，执行格式转换（统一日期不会误伤数值列）
4. **列结构问题**: 多信息混在一列、列名不规范等，执行拆分或规范化

【输出格式】
输出 JSON 格式的清洗策略（只输出 JSON，不要 Markdown 代码块标记，不要解释）：
{
  "operations": ["removeDuplicates", "fillEmpty"],  // 清洗操作数组
  "sheet": null,              // 工作表；null=当前表
  "range": null,              // 范围如 "A1:D100"；null=当前选区
  "fillMethod": "ffill",      // 空值填充：ffill(前值) | zero(0) | delete(删除行) | placeholder(占位)
  "dedupeColumns": null,      // 去重依据列数组；null=整行去重
  "markDuplicate": true       // 去重模式（默认 true=标记不删除；false=真正删除重复行）
}

【清洗动作说明】
- removeDuplicates: 处理重复（默认标记不删除：把重复行整行黄色高亮标记、不新增列、不删除；需删除时标记 markDuplicate=false）
- fillEmpty: 填充空值（配合 fillMethod）
- trimSpaces: 去除前后空格
- normalizeDates: 统一日期为 yyyy-mm-dd（不会误伤数值列）
- normalizeNumbers: 统一数字格式
- standardizeText: 文本规范化（统一大小写/去除多余空格）
- removeEmptyRows: 删除全空行
- splitColumn: 按分隔符拆分列
- mergeColumns: 合并列

## Template Code

> 本 Skill 不再使用 VBA 模板代码。清洗全部由前端 MCP 工具链执行：
> 统一走 `excel_clean_data`（原生引擎，一次性完成）；其 `removeDuplicates` 默认【标记重复不删除】（重复行黄色高亮，不新增列），
> 失败回退 `excel_read_range` 读取 + JS 确定性变换 + `excel_write_range` 写回。不回退 VBA。
