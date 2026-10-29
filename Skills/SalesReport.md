---
hosts: [excel]
name: SalesReport
displayName: 销售报表生成器
description: 自动生成销售报表：读取数据 → 按维度聚合 → 生成柱状图 + 写入汇总表。
keywords:
  - 销售报表
  - 销售分析
  - 月度报告
  - 区域销售
  - sales report
examples:
  - 用当前选区生成销售报表，按地区汇总
  - 生成月度销售报表，按产品分组
isGeneric: true
category: 报表生成
pipeline: skill
processingMode: SkillSpec
executionMode: steps
allowedTools:
  - excel_read_range
  - excel_aggregate
  - excel_create_chart
  - excel_write_range
  - excel_create_sheet
  - excel_format_range
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
version: 0.1
# ── 声明式 steps：按顺序依次执行，前序步骤产出自动暴露给后序步骤 ──
# params 里 {{ctx.xxx}} 引用 turn context，{{步骤名.result.field}} 引用前序步骤
steps:
  - name: readData
    tool: excel_read_range
    displayName: 读取数据
    params: { "sheet": "{{ctx.sheet || 'active'}}", "range": "{{ctx.range || 'selection'}}" }

  - name: aggregateByRegion
    tool: excel_aggregate
    displayName: 按地区聚合
    params: { "sourceSheet": "{{readData.result.sheet}}", "sourceRange": "{{readData.result.range}}", "groupBy": ["地区"], "funcs": [{"column": "销售额", "func": "sum", "alias": "总销售额"}], "createChart": false }

  - name: createReportSheet
    tool: excel_create_sheet
    displayName: 创建报表工作表
    params: { "name": "销售报表" }

  - name: writeSummary
    tool: excel_write_range
    displayName: 写入汇总表
    params: { "sheet": "销售报表", "range": "A1", "data": "{{aggregateByRegion.data}}", "headers": "{{aggregateByRegion.headers}}" }

  - name: createBarChart
    tool: excel_create_chart
    displayName: 生成柱状图
    params: { "sheet": "{{aggregateByRegion.resultSheet}}", "chartType": "xlColumnClustered", "anchorSheet": "销售报表", "anchorRange": "D2", "dataValues": "{{aggregateByRegion.data}}", "title": "各地区销售额" }
---
# 销售报表生成器 SalesReport

当用户说"生成销售报表/销售分析/月度报告"时，自动按以下步骤施工。

## 执行流程
1. **readData** — 读取用户选中区域或指定 sheet/range 的数据
2. **aggregateByRegion** — 按"地区"维度聚合销售额 sum（groupBy 可按需替换）
3. **createReportSheet** — 新建"销售报表"工作表
4. **writeSummary** — 把聚合结果写入新表 A1 起始位置
5. **createBarChart** — 在新表 D2 位置锚定柱状图（dataValues 内存绑定，不落地中间表）

## 变量模板
| 模板 | 含义 |
|------|------|
| `{{ctx.sheet}}` | turn context 里的 sheet（用户当前表） |
| `{{ctx.range}}` | turn context 里的 range（用户选中区域） |
| `{{readData.result.sheet}}` | 前序步骤 readData 产出的 sheet 字段 |
| `{{aggregateByRegion.data}}` | 聚合结果的 data 数组（对象数组） |
| `{{aggregateByRegion.headers}}` | 聚合结果的列头 |
