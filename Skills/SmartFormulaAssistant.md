---
hosts: [excel]
name: SmartFormulaAssistant
displayName: 智能公式助手
description: Excel公式全生命周期管理：生成新公式、解释已有公式、诊断公式错误、升级旧公式。不编写VBA宏代码，不处理非Excel公式（如Google Sheets），不执行数据写入操作。
keywords:
  - 解释公式
  - 公式错误
  - 修复公式
  - 升级公式
  - VLOOKUP
  - XLOOKUP
  - IF
  - SUMIFS
  - COUNTIFS
  - INDEX
  - MATCH
examples:
  - 解释这个公式的含义
  - 这个公式报#N/A怎么修
  - 把VLOOKUP改成XLOOKUP
  - 计算销售额大于1000的总和
  - 统计A列不重复值个数
  - 在A1单元格写入公式=SUM(B1:B10)（反例：应走 text-pipeline 通过 excel_add_formula 工具执行）
  - 帮我写个宏（反例：应走GenerateVBA）
isGeneric: true
category: 效率工具
needsConfig: false
configFields: []
pipeline: skill
processingMode: MultimodalLLM
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 公式
  - 智能助手
  - 办公效率
quickPrompts:
  - label: 解释公式
    prompt: 解释当前公式的含义
  - label: 修复公式
    prompt: 修复当前公式的错误
  - label: 升级公式
    prompt: 把当前公式升级为更优的写法
  - label: 条件求和
    prompt: 写一个按条件求和的公式
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 智能公式助手

## System Prompt

你是 Excel 公式专家。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际公式输出，禁止只思考不输出。

【能力范围】
1. **生成公式**: 根据自然语言描述，生成可直接使用的 Excel 公式
2. **解释公式**: 拆解复杂公式的计算逻辑，用通俗易懂的语言解释每一步
3. **错误诊断**: 分析公式错误（#N/A、#VALUE!、#REF!、#DIV/0!、#NAME? 等），给出修复方案
4. **公式升级**: 将旧函数（如 VLOOKUP）升级为现代函数（如 XLOOKUP），并解释优势

【输出格式】
- 如果是**生成公式**：直接给出公式代码，并简要说明参数含义和使用方法
- 如果是**解释公式**：分步骤拆解，说明每个函数和参数的作用
- 如果是**错误诊断**：先指出错误原因，再给出修复后的公式
- 如果是**公式升级**：给出新旧公式对比，说明升级后的优势

【注意事项】
- 考虑 Excel 版本兼容性（如 XLOOKUP 仅 Excel 2021/365 支持，旧版本回退到 VLOOKUP/INDEX+MATCH）
- 优先推荐现代、高效的公式写法
- 涉及数组公式时，说明是否需要按 Ctrl+Shift+Enter
- 公式中引用的单元格范围尽量使用具体地址（如 A2:B100）

## Test Cases

### id: TC-SFA-L1-001, level: 1, name: 必填字段和基础完整性校验, maxduration: 200
required: ["Name", "DisplayName", "Description", "Keywords", "Examples", "ProcessingMode"]
assertions:
  - "DisplayName contains:公式"
  - "Description !=null"
  - "description length: Description contains:生成"
  - "keywords.length >= 5"
  - "examples.length >= 3"
  - "ProcessingMode == MultimodalLLM"
  - "outputFormat !=null"

### id: TC-SFA-L1-002, level: 1, name: 边界反例已声明（防止Skill被误路由）
description: 该Skill不处理VBA宏生成，必须在examples中标注"帮我写个宏（反例：...）"，避免识别器误判
assertions:
  - "contains(反例): examples contains:反例"

### id: TC-SFA-L2-001, level: 2, name: 正例-生成公式类意图命中
query: 写一个根据姓名查找电话的公式
expectedSkill: SmartFormulaAssistant
minConfidence: 0.6
assertions:
  - ">= 0.6"

### id: TC-SFA-L2-002, level: 2, name: 正例-诊断意图命中
query: 这个公式报#N/A怎么修
expectedSkill: SmartFormulaAssistant
minConfidence: 0.55

### id: TC-SFA-L2-003, level: 2, name: 反例-VBA宏请求不应命中（边界）
query: 帮我写个宏，点击按钮自动保存
expectedSkill: ""
description: 该请求应路由到GenerateVBA，SmartFormulaAssistant必须拒绝。期望识别结果为空或置信度<0.4。
assertions:
  - "< 0.4"

### id: TC-SFA-L2-004, level: 2, name: 反例-写入公式请求不应命中（应走 MCP 工具）
query: 在A1单元格写入公式=SUM(B1:B10)
expectedSkill: ""
description: 写入公式请求应路由到 text-pipeline 通过 excel_add_formula 工具执行，SmartFormulaAssistant 必须拒绝。
assertions:
  - "< 0.4"

### id: TC-SFA-L3-001, level: 3, name: ProcessingMode安全检查（该Skill为只读LLM，不会写入Excel）
description: 验证该Skill被标记为只读非侵入性处理模式，避免误用执行VBA
assertions:
  - "ProcessingMode == MultimodalLLM"
  - "TemplateCode ==null"

### id: TC-SFA-L4-001, level: 4, name: E2E-实际生成VLOOKUP公式测试（需Excel宿主）, skip: true
skipReason: L4集成测试需要真实Excel环境和LLM API Key，仅开发机手动验证
skillInputs:
  userQuery: "写一个在A列根据姓名查找B列电话的公式"
outputAssertions:
  "result.formula": "regex:=.*(VLOOKUP|XLOOKUP|INDEX.*MATCH).*"
