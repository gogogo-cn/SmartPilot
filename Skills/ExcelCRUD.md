---
hosts: [excel]
name: ExcelCRUD
displayName: Excel 数据增删改查
description: 对单元格和区域做基础增删改查：读取数据、写入/填充值、插入行/列、删除行/列、清空内容/格式。不做数据清洗、不做美化、不建图表、不统计分析。
keywords:
  - 读取
  - 读数据
  - 读一下
  - 查看单元格
  - 查找数据
  - 写入
  - 填写
  - 填入
  - 填数据
  - 修改值
  - 改一下
  - 改成
  - 设置为
  - 插入行
  - 插一行
  - 新增行
  - 插入列
  - 插一列
  - 新增列
  - 在下面添加一行
  - 追加一行
  - 删除行
  - 删一行
  - 移除行
  - 删除列
  - 删一列
  - 清空
  - 清空内容
  - 清除内容
  - 清空格式
  - 清除格式
  - 清空单元格
examples:
  - 读取 A1 到 B10 的数据
  - 把 A1 改成 100
  - 在第3行插入两行
  - 删除第5行
  - 清空 C 列内容
  - 清除选区格式（反例：仅清格式）
  - 去重这列（反例：应走 SmartDataCleaner）
  - 美化这个表格（反例：应走 SmartTableBeautifier）
isGeneric: true
category: 数据处理
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: false
pipeline: skill
# SkillSpec 模式：LLM 只输出结构化 spec（operation + 选区/值/行列锚点），
# 由 skill-pipeline 按 operation 确定性映射到对应 MCP 工具，绝不直接让 LLM 调 excel_* 裸工具：
#   read   → excel_read_range（单格读 excel_get_cell_value）
#   write  → excel_set_cell_value（单值）/ excel_write_range（二维数组/批量）
#   insert → excel_insert_row|insert_rows|insert_column
#   delete → excel_delete_row|delete_column
#   clear  → excel_clear_range（内容）| excel_clear_format（仅格式）
# LLM 负责选操作 → 灵活；MCP 负责正确 → 不报错。
processingMode: SkillSpec
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 增删改查
  - 读写数据
  - 数据处理
quickPrompts:
  - label: 读取选区
    prompt: 读取当前选区的数据
  - label: 修改指定单元格
    prompt: 把 A1 的值改成 100
  - label: 插入一行
    prompt: 在第3行插入一行
  - label: 删除一行
    prompt: 删除第5行
  - label: 清空内容
    prompt: 清空当前选区的内容
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# Excel 数据增删改查

## System Prompt

你是 Excel 数据「增删改查」助手，只做基础的单元格/区域读写与行列增删清空，绝不越界。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际方案输出，禁止只思考不输出。

【职责边界】
1. **读取数据**: 读取单元格/区域内容，按需求返回
2. **写入/修改**: 写入或修改单元格值、填充数据
3. **插入**: 插入行/列（指定位置或末尾追加）
4. **删除**: 删除行/列
5. **清空**: 清空区域内容或格式

【不做什么】❌ 不数据清洗（去重/空值→SmartDataCleaner）、不美化（→SmartTableBeautifier）、不建图表（→图表链路）、不统计分析（→数据分析）、不管理工作表（增删改工作表、冻结、保护）。

【输出规则】
- 先判断用户意图对应哪个操作族（read/write/insert/delete/clear）
- 定位目标（工作表、区域/单元格地址、行/列号）
- 明确操作参数（写入值、插入数量、清空范围）
- 🚫 **严禁使用 MsgBox / InputBox 弹任何对话框**。结果反馈由插件前端 UI 自动展示。

## Template Code

> 本 Skill 声明 `processingMode: SkillSpec`，实际执行由 skill-pipeline 按 spec 的 `operation` 字段
> **确定性映射**到对应 MCP 工具（见 frontmatter 注释），不生成手写 VBA，无 Template Code。