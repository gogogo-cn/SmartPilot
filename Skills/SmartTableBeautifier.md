---
# ⚠️ 仅 Excel：本 Skill 是 SkillSpec 形态，靠前端 Spec 引擎渲染 Excel VBA（依赖 ExcelBridge / excel_execute_vba）。
# Word 宿主里无 ExcelBridge，注册到 Word 后 "美化表格" 会被路由进 Spec 引擎 → 生成 Excel VBA 在 Word 中执行 → Word 卡死（实测 2026-10-05）。
# Word/PPT 的选中表格美化改由非 Excel 硬短路直调 word_/ppt_format_selected_table_cell（Hosts/*/bin/Debug/Config 内工具），故此处不再声明 word。
hosts: [excel]
name: SmartTableBeautifier
displayName: 智能表格美化
description: 一键专业美化表格、智能格式识别、自动调整行列、主题色应用、智能汇总行。含格式整理能力：统一字体字号、标准边框、智能对齐、数字日期格式化、冻结首行。不做数据清洗，不做数据分析，不修改公式。
keywords:
  - 美化
  - 美化表格
  - 表格美化
  - 美化这个表格
  - 美化一下
  - 格式化
  - 表格样式
  - 自动调整
  - 一键美化
  - 主题
  - 汇总
  - 美化选区
  - 美化工作表
  - 调整格式
  - 设置样式
  - 格式整理
  - 整理表格
  - 统一格式
  - 清理格式
  - 清理冗余格式
  - 表头美化
  - 标准边框
  - 统一字体
  - 数字格式化
  - 日期格式化
  - 冻结首行
  - 自动换行
examples:
  - 美化这个表格
  - 一键格式化
  - 应用专业样式
  - 自动调整列宽
  - 分析这张表的数据（反例：应走数据分析）
  - 清洗这些数据（反例：应走SmartDataCleaner）
isGeneric: true
category: 数据处理
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
pipeline: skill
# SkillSpec 模式：LLM 输出结构化 spec（主题/表头/对齐/条件格式/冻结/斑马纹等），
# 由 skillSpecEngine 确定性渲染 VBA。LLM 负责规划样式 → 灵活多变；模板负责代码正确性 → 永不报错。
processingMode: SkillSpec
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
specFields:
  - name: stylePreset
    label: 样式方案
    type: select
    options:
      - classic_grid
      - deep_header
      - three_line
      - borderless
      - card_stripe
      - underline_header
      - first_col_focus
      - total_row
      - compact_report
      - magazine
    defaultValue: classic_grid
  - name: borderStyle
    label: 边框版式
    type: select
    options:
      - all
      - none
      - horizontal
      - outer
      - threeLine
    defaultValue: all
  - name: headerLayout
    label: 表头版式
    type: select
    options:
      - deep
      - light
      - underline
      - italic
    defaultValue: deep
  - name: rowHeight
    label: 行高(磅)
    type: number
    defaultValue: null
  - name: valign
    label: 垂直对齐
    type: select
    options:
      - middle
      - top
    defaultValue: middle
  - name: boldFirstColumn
    label: 首列加粗
    type: boolean
    defaultValue: false
  - name: emphasizeLastRow
    label: 末行(合计行)强调
    type: boolean
    defaultValue: false
  - name: bodyNumberRight
    label: 数字右对齐
    type: boolean
    defaultValue: false
  - name: theme
    label: 主题风格
    type: select
    options:
      - business_blue
      - tech_green
      - vibrant_orange
      - elegant_purple
      - minimalist_gray
      - finance_red
    defaultValue: business_blue
  - name: headerStyle
    label: 表头样式
    type: select
    options:
      - bold_colored
      - bold_only
      - colored_only
      - italic
    defaultValue: bold_colored
  - name: alignment
    label: 对齐方式
    type: select
    options:
      - center
      - left
      - right
    defaultValue: center
  - name: autofit
    label: 自动列宽
    type: boolean
    defaultValue: true
  - name: freezeHeader
    label: 冻结表头
    type: boolean
    defaultValue: false
  - name: addBorders
    label: 添加边框
    type: boolean
    defaultValue: true
  - name: bandedRows
    label: 斑马纹隔行
    type: boolean
    defaultValue: false
  - name: conditionalFormat
    label: 条件格式
    type: select
    options:
      - null
      - dataBar
      - colorScale
      - highlight_top
      - highlight_duplicates
    defaultValue: null
  - name: numberFormat
    label: 数字格式
    type: string
    defaultValue: null
  - name: title
    label: 标题行文字
    type: string
    defaultValue: null
  - name: fontName
    label: 字体
    type: string
    defaultValue: 微软雅黑
  - name: fontSize
    label: 字号
    type: number
    defaultValue: 11
  - name: dateFormat
    label: 日期格式
    type: string
    defaultValue: null
  - name: wrapText
    label: 自动换行
    type: boolean
    defaultValue: false
  - name: clearRedundant
    label: 清理冗余格式
    type: boolean
    defaultValue: false
tags:
  - 格式化
  - 表格美化
  - 格式整理
  - 汇报
# 任意 Skill 声明 quickPrompts 即自动生效，无需改渲染逻辑。
quickPrompts:
  - label: 三线表（论文）
    prompt: 把当前表格做成三线表（学术风：顶线+表头下线+底线，无竖线无底色）
  - label: 无边框留白
    prompt: 把当前表格做成无边框留白版式（去掉网格，靠行距分区，宽松行高）
  - label: 卡片条纹
    prompt: 把当前表格做成卡片条纹版式（隔行浅底，无竖线，宽松行距）
  - label: 深表头商务
    prompt: 把当前表格做成深表头版式（深色表头白字 + 全网格 + 隔行底纹）
  - label: 紧凑报表
    prompt: 把当前表格做成紧凑报表版式（行高压紧，数字右对齐）
  - label: 首列强调
    prompt: 把当前表格做成首列强调版式（首列加粗浅底，正文左对齐）
  - label: 合计行
    prompt: 把当前表格做成合计行版式（末行加粗 + 上边线，数字右对齐）
  - label: 杂志宽松
    prompt: 把当前表格做成杂志宽松版式（深表头 + 宽松行距 + 横向分隔线）
  - label: 商务蓝
    prompt: 用商务蓝主题美化表格
  - label: 科技绿
    prompt: 用科技绿主题美化表格
  - label: 活力橙
    prompt: 用活力橙主题美化表格
  - label: 优雅紫
    prompt: 用优雅紫主题美化表格
  - label: 极简灰
    prompt: 用极简灰主题美化表格
  - label: 金融红
    prompt: 用金融红主题美化表格
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 2
---

# 智能表格美化

## System Prompt

你是 Excel 表格美化专家。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际美化方案输出，禁止只思考不输出。

【美化能力】
1. **表头样式**: 应用专业的表头背景色、字体加粗、白色字体
2. **斑马纹**: 交替行背景色，提升可读性
3. **边框优化**: 应用精致的网格线或无边框设计
4. **格式识别**: 自动识别日期列→日期格式，金额列→货币格式，百分比→百分比格式
5. **列宽优化**: 根据内容自动调整列宽，避免显示不全或过度留白
6. **冻结首行**: 大数据表自动冻结首行
7. **汇总行**: 在表格底部添加 SUM/AVERAGE 汇总

【主题风格（配色）】
- 商务蓝 / 科技绿 / 活力橙 / 优雅紫 / 极简灰 / 金融红 / 墨黑金 / 晴空青

【样式方案（排版）— 主题不只是换颜色，版式也要变】
- classic_grid 经典网格：全网格边框 + 浅色表头 + 隔行底纹 + 居中
- deep_header 深表头：深底白字表头，视觉重心在上
- three_line 三线表：学术/论文规范，仅顶线 + 表头下线 + 底线，无竖线无底色
- borderless 无边框留白：去网格，靠行距分区，宽松行高
- card_stripe 卡片条纹：整行浅底交替，无竖线，宽松行距
- underline_header 下划线表头：表头无底色只留下边线，仅横线分隔
- first_col_focus 首列强调：首列加粗 + 浅底，正文左对齐
- total_row 合计行：末行加粗 + 上边线 + 浅底，数字右对齐
- compact_report 紧凑报表：行高压紧，数字右对齐
- magazine 杂志宽松：深表头 + 宽松行距 + 横向分隔

【排版维度】
- 边框版式 borderStyle：all 全网格 / none 无边框 / horizontal 仅横线 / outer 仅外框 / threeLine 三线表
- 表头版式 headerLayout：deep 深底白字 / light 浅底深字 / underline 仅底部粗线 / italic 斜体
- 行高 rowHeight：紧凑 15 / 标准 20 / 宽松 28 磅；valign 垂直对齐 middle|top
- 强调开关：boldFirstColumn 首列加粗、emphasizeLastRow 末行合计、bodyNumberRight 数字右对齐

【输出规则】
- 分析用户选区数据结构
- 推荐最佳美化方案
- 生成可直接执行的 VBA 代码
- 说明应用了哪些格式变化
- 🚫 **严禁使用 MsgBox / InputBox / MsgBox(...) 弹任何对话框**：否则沙箱预处理会把弹框文本（如"表格美化完成"）强行写入 A1 单元格，覆盖用户表头造成数据破坏。所有提示输出一律改为 `Debug.Print`。美化成功/失败的通知由插件前端 UI 自动显示，不需要你在 VBA 中再弹框。

## Template Code

> 本 Skill 声明 `processingMode: SkillSpec`，实际执行由前端 **Spec 引擎**（skillSpecEngine）根据 LLM 规划的结构化 spec（主题/表头/对齐/条件格式等）**动态生成 VBA** 并调用 `excel_execute_vba` 执行。
> 下方旧版手写 VBA 宏仅作参考，**不参与执行**。

```vba
' 说明：实际美化逻辑由 Spec 引擎按 plan 动态渲染，无需手写模板。
' 渲染方向参考：
'   - 主题色表头（Font.Bold / Interior.Color）
'   - 斑马纹（奇数/偶数行交替底色）
'   - 自动调整列宽（Columns.AutoFit）
'   - 边框（Borders.LineStyle / Weight）
'   - 冻结首行（FreezePanes / SplitRow）
```
