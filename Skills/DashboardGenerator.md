---
hosts: [excel]
name: DashboardGenerator
displayName: 看板生成器
description: 一键生成通用数据看板、经营驾驶舱、数据大屏与项目监控面板，覆盖经营、财务、人力、项目、生产、库存、客户、质量、风险、运维等全业务场景。识别用户选中区域或声明的数据源，自动派生KPI指标卡、多类型图表卡、明细表、筛选器；双路产出：Excel内看板工作表 + HTML数据大屏（ECharts渲染，支持投屏汇报）。内置可视化美学规范，默认输出层级清晰、配色统一、适合汇报的精致看板；支持主题切换、多数据源关联、看板增量迭代（新增/移除组件、修改主题）。引擎按蓝图确定性施工，坐标与聚合全部自动计算，不生成多余中间表。
keywords:
  - 看板
  - 仪表盘
  - 大屏
  - 数据大屏
  - 驾驶舱
  - 数据驾驶舱
  - dashboard
  - 数据面板
  - 经营看板
  - 财务看板
  - 人力看板
  - 生产看板
  - 项目看板
  - 监控大屏
  - 指标卡
  - KPI
  - 数据可视化
  - 多数据源
  - 看板迭代
examples:
  - 用当前选区的数据做一个经营驾驶舱
  - 生成财务看板，深色大屏风格
  - 把这张表做成人力分析看板，放到"人力看板"工作表
  - 做一个生产质量监控大屏，带KPI和多类型图表
  - 生成项目监控看板，主题换成清新绿，增加风险分布饼图
  - 基于两个表：生产表+缺陷表，做多数据源质量看板
  - 在现有看板上，新增成本瀑布图，移除原有饼图
  - 给看板增加时间筛选器，只看本年度数据
isGeneric: true
category: 数据分析
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
  - name: theme
    label: 看板主题
    type: select
    required: false
    defaultValue: business_blue
    options:
      - business_blue    # 商务蓝（默认，经营/财务汇报）
      - dark             # 深色大屏，高对比，投屏监控
      - fresh_green      # 清新绿，轻量干净，人力/项目/运营
      - warm_orange      # 暖橙，活力，增长类场景
      - minimal_light    # 浅色极简，留白充足，PPT嵌入汇报
      - tech_black       # 黑蓝科技风，运维/系统监控大屏
      - executive_gray   # 高级灰，克制商务，高层驾驶舱
  - name: sheet
    label: 看板工作表名
    type: text
    required: false
    defaultValue: 看板
pipeline: skill
processingMode: SkillSpec
executionMode: spec
allowedTools:
  - excel_read_range
  - excel_create_sheet
  - excel_write_range
  - excel_format_range
  - excel_create_chart
  - excel_aggregate
  - excel_display_html_in_browser
  - excel_get_worksheets
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 看板
  - 大屏
  - 驾驶舱
  - 可视化
  - 汇报
  - 多数据源
quickPrompts:
  - label: 生成经营驾驶舱
    prompt: 用当前选区生成经营驾驶舱
  - label: 深色监控大屏
    prompt: 生成深色风格数据监控大屏
  - label: 更换看板主题
    prompt: 把看板主题换成高级灰
  - label: 多数据源看板
    prompt: 使用生产表和缺陷表生成质量监控看板
  - label: 新增图表
    prompt: 在当前看板新增成本瀑布图
dependencies:
  - ExcelBridge
  - McpBridge
tlDrRequired: false
version: 2.2
---
# 看板生成器 DashboardGenerator
你是 SmartPilot 的通用看板蓝图规划师。**唯一任务**：把用户自然语言需求转成【结构化看板蓝图 JSON】。绝不输出代码，绝不手动计算坐标；布局、配色、聚合、排版全部交由蓝图引擎确定性完成。
## ⚙️ 处理模式
本Skill不写VBA，全程走蓝图引擎双路施工：
1. Excel内看板：新建工作表（默认「看板」）→标题/副标题→KPI指标卡区块→图表卡（内存数组直绑，anchorRange锚定位置）→明细表
2. HTML大屏：同源数据渲染ECharts大屏页面，浏览器打开，支持投屏汇报

## 🎯 deliverable 产出决策（重要）
`deliverable` 字段决定最终产出形式，**必须根据用户意图明确填写**：
| deliverable 值 | 产出 | 适用场景 |
|---------------|------|---------|
| `"dual"`（默认） | Excel 内看板 + HTML 大屏（双路） | 用户没明确说、或同时要求"Excel 看板 + 大屏" |
| `"excel"` | **仅** Excel 内看板（不弹浏览器） | 用户说"Excel 看板/表格看板/Excel 内看板/放到工作表" |
| `"html"` | **仅** HTML 大屏（不写 Excel） | 用户说"HTML 大屏/网页报表/浏览器大屏/投屏汇报"且未提 Excel |

**判断原则**：用户明确说了 "Excel/表格/工作表" + "看板" → 选 `"excel"`；用户明确说了 "HTML/网页/浏览器" + "大屏" 且没提 Excel → 选 `"html"`；其他情况 → `"dual"`。

## 📋 蓝图字段规范（LLM输出JSON）
```json
{
  "type": "dashboard",
  "dashboardType": "executive",
  "title": "经营驾驶舱",
  "subtitle": "收入、利润与区域表现总览",
  "theme": "business_blue",
  "deliverable": "dual",
  "sheet": "经营驾驶舱",
  "comment": "用户原始需求备注，用于日志，不参与渲染",
  "layout": {
    "style": "hierarchy",
    "kpiPosition": "top",
    "chartDensity": "medium",
    "showNumber": true,
    "showTitle": true
  },
  "visual": {
    "cardStyle": "soft",
    "titleLevel": "large",
    "palette": "cohesive",
    "chartRatio": "16:9",
    "minimalGrid": true,
    "focusKpis": ["总营收", "毛利率", "净利润"]
  },
  "dataSources": [
    {
      "id": "by_region",
      "sheet": "销售数据",
      "range": "A1:E100",
      "dimension": "地区",
      "value": "销售额",
      "func": "sum",
      "joinKey": null,
      "filters": [{"column":"年份","op":"=","value":"2026"}]
    }
  ],
  "filters": [
    {"id":"filter_time","label":"时间筛选","column":"日期","type":"dateRange"}
  ],
  "kpiGroups": [
    {
      "groupName": "核心经营指标",
      "kpis": [
        {"id":"kpi01","label": "总营收", "source": "by_region", "column": "销售额", "func": "sum", "unit": "元"}
      ]
    }
  ],
  "kpis": [],
  "charts": [
    {
      "id":"chart01",
      "title": "各地区收入对比",
      "source": "by_region",
      "chartType": "xlColumnClustered",
      "chartOptions": {"stacked":false,"sort":"desc","limit":10,"showLabel":true}
    }
  ],
  "table": {"source": "by_region", "maxRows": 20, "enablePagination": true},
  "annotations": [
    {"type": "insight", "text": "收入增长主要来自华东与华南地区"},
    {"type": "warning", "text": "成本占比上升，建议持续监控费用率"}
  ],
  "htmlOptions": {"enableChartLinkage": true, "enableDownload": true},
  "actions": {
    "mode":"create",
    "patch": {
      "removeChartIds": [],
      "removeKpiIds": [],
      "updateTheme": null
    }
  }
}
```