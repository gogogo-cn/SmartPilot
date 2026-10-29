# 第二优先级扩展功能使用指南

## 📋 概述

本文档详细介绍 SmartPilot 插件的三个核心体验升级功能：

1. **事件驱动与定时触发** - 让 Skill 自动响应 Excel 事件或定时执行
2. **增强输出格式（图表/透视表/条件格式）** - 直接生成 Excel 原生可视化对象
3. **权限与审批机制** - 安全防护层，敏感操作需用户确认

---

## 🎯 功能一：事件驱动与定时触发

### 1.1 功能说明

**作用**：允许 Skill 绑定到 Excel 事件（单元格变化、工作簿打开等）或定时器，实现真正的自动化监控和无人值守执行。

**适用场景**：
- 数据变化时自动计算/校验
- 定时生成报表
- 工作簿打开时自动加载数据
- 实时数据监控和告警

### 1.2 配置参数

在 Skill 定义 JSON 中添加 `triggerConfig` 属性：

```json
{
  "name": "YourSkill",
  "triggerConfig": {
    "triggerType": "Event",              // 触发类型：Event / Timer / Manual
    "eventType": "Change",               // 事件类型（Event模式）
    "filterCondition": "",               // 过滤条件
    "debounceMs": 1000,                 // 防抖时间（毫秒）
    "enabled": true,                     // 是否启用
    "timerIntervalSeconds": 0,           // 定时间隔秒数（Timer模式）
    "fireOnStart": false                // Timer模式：启动时是否立即执行一次
  }
}
```

#### TriggerType 枚举值

| 类型 | 说明 | 适用场景 |
|------|------|----------|
| `Manual` | 手动触发（默认） | 普通Skill，用户主动调用 |
| `Event` | Excel事件触发 | 自动化响应 |
| `Timer` | 定时器触发 | 定时任务、周期性检查 |

#### EventType 支持列表

| 事件类型 | 触发时机 | 可用上下文变量 |
|----------|----------|----------------|
| `Change` | 单元格内容修改后 | targetAddress, targetColumn, targetRow, targetSheet, cellValue |
| `Open` | 工作簿打开时 | 无特殊上下文 |
| `Close` | 工作簿关闭前 | 无特殊上下文 |
| `SelectionChange` | 选区变化时 | targetAddress, targetColumn, targetRow, targetSheet |

### 1.3 FilterCondition 语法

支持简单表达式，用于精确控制何时触发：

```javascript
// 格式: 变量名 操作符 值
"targetColumn = 'B'"                    // 仅B列变化时触发
"targetSheet = '订单表'"                // 仅指定工作表
"targetColumn = 'B' && targetSheet = '订单表'"  // 组合条件
"cellValue != ''"                       // 单元格非空时触发
"targetRow > 10"                        // 行号大于10时触发
```

**可用变量**：
- `targetColumn` - 变化的列字母（A, B, C...）
- `targetRow` - 变化的行号
- `targetSheet` - 工作表名称
- `targetAddress` - 单元格地址（如"B5"）
- `cellValue` - 单元格的新值

**支持的操作符**：`=`, `==`, `<>`, `!=`, `>`, `<`, `>=`, `<=`, `contains`

### 1.4 防抖机制（Debounce）

**问题**：用户快速输入或批量粘贴时，Change 事件会频繁触发
**解决方案**：`debounceMs` 参数控制最小触发间隔

```json
{
  "triggerConfig": {
    "debounceMs": 1000   // 1秒内只触发一次（适合输入场景）
  }
}
```

**推荐值**：
- 输入框自动保存：500-1000ms
- 数据校验：200-500ms
- 批量操作：1000-2000ms
- 即时响应：0-100ms

### 1.5 工作原理

```
Excel 事件发生 (如 Change)
    ↓
SkillTriggerEngine.HandleEventTrigger()
    ↓
遍历所有已注册的 Event 类型触发器
    ↓
检查 eventType 是否匹配? ──No──→ 跳过
    ↓ Yes
防抖检查: 距上次触发 < debounceMs? ──Yes──→ 跳过
    ↓ No (更新最后触发时间)
FilterCondition 评估? ──False──→ 跳过
    ↓ True
异步执行 Skill:
  ├─ 构造 inputs (包含事件上下文)
  ├─ 调用 CustomSkillExecutor.ExecuteCustomSkill()
  └─ 结果通知前端 WebView
```

### 1.6 示例配置

#### 示例 1：自动佣金计算（事件驱动）

```json
{
  "name": "AutoCommissionCalculator",
  "displayName": "自动佣金计算器(事件触发)",
  "description": "当订单表B列(金额)变化时，自动计算佣金写入C列",
  "processingMode": "Hybrid",
  "securityLevel": "Safe",                    // 安全操作，无需确认
  "outputFormat": "vba",
  "triggerConfig": {
    "triggerType": "Event",
    "eventType": "Change",
    "filterCondition": "targetColumn = 'B' && targetSheet = '订单表'",
    "debounceMs": 1000,
    "enabled": true
  },
  "templateCode": "Sub AutoCalculateCommission()\n    Dim targetCell As Range\n    Set targetCell = Application.Range(\"{{TARGET_ADDRESS}}\")\n    If IsNumeric(targetCell.Value) Then\n        targetCell.Offset(0, 1).Value = CDbl(targetCell.Value) * 0.1\n    End If\nEnd Sub"
}
```

**效果**：用户在"订单表"的B列输入金额 → 1秒后自动在C列计算10%佣金

---

#### 示例 2：定时数据备份（定时触发）

```json
{
  "name": "DailyDataBackup",
  "displayName": "每日数据备份(定时)",
  "description": "每天下午5点自动备份当前工作簿到指定目录",
  "processingMode": "VBACode",
  "securityLevel": "NeedsConfirmation",
  "triggerConfig": {
    "triggerType": "Timer",
    "timerIntervalSeconds": 86400,          // 24小时 = 86400秒
    "fireOnStart": false,
    "enabled": true
  },
  "templateCode": "Sub DailyBackup()\n    ThisWorkbook.SaveCopyAs \"C:\\Backup\\\" & Format(Date, \"yyyy-mm-dd\") & \".xlsx\"\nEnd Sub"
}
```

**效果**：每24小时自动执行一次备份（首次启动不执行）

---

#### 示例 3：工作簿打开时加载汇率（事件驱动）

```json
{
  "name": "AutoLoadExchangeRate",
  "displayName": "自动加载最新汇率",
  "description": "工作簿打开时从API获取最新汇率并写入指定区域",
  "processingMode": "MultimodalLLM",
  "securityLevel": "Safe",
  "triggerConfig": {
    "triggerType": "Event",
    "eventType": "Open",
    "enabled": true
  },
  "llmSystemPrompt": "获取今日美元兑人民币汇率并输出JSON格式"
}
```

**效果**：每次打开包含此Skill的工作簿时，自动加载最新汇率

---

## 📊 功能二：增强输出格式（图表/透视表/条件格式）

### 2.1 功能说明

**作用**：Skill 的输出不再局限于文本，可以直接生成 Excel 原生的图表、数据透视表、条件格式等可视化对象。

**核心优势**：
- 生成的图表是原生 Excel 对象（可编辑、可交互）
- 支持多种图表类型（柱状图、折线图、饼图、散点图等）
- 数据透视表支持动态筛选和钻取
- 条件格式实时响应数据变化

### 2.2 outputFormat 枚举值

| 格式 | 说明 | 调用方法 |
|------|------|----------|
| `text` | 纯文本（默认） | 直接返回字符串 |
| `json` | JSON结构化数据 | 解析为对象 |
| `markdown` | Markdown文本 | 渲染显示 |
| `vba` | VBA代码 | 在沙箱中执行 |
| **`chart`** | **Excel图表** | **FormatBridge.RenderChart()** |
| **`pivotTable`** | **数据透视表** | **FormatBridge.RenderPivotTable()** |
| **`conditionalFormat`** | **条件格式规则** | **FormatBridge.ApplyConditionalFormatting()** |

### 2.3 图表生成（Chart）

#### 配置参数

在 Skill 定义中添加 `chartConfig` 属性：

```json
{
  "outputFormat": "chart",
  "chartConfig": {
    "chartType": "bar",              // 图表类型
    "title": "销售数据分析",         // 图表标题
    "xAxisField": "",               // X轴字段（可选）
    "yAxisFields": "",              // Y轴字段（逗号分隔）
    "targetRange": "",              // 目标位置（空=自动）
    "width": 500,                   // 宽度（像素）
    "height": 300,                  // 高度（像素）
    "showLegend": true              // 显示图例
  }
}
```

#### 图表类型映射

| chartType 值 | Excel常量 | 说明 |
|--------------|-----------|------|
| `"bar"` 或 `"column"` | xlColumnClustered | 柱状图（默认）|
| `"line"` | xlLine | 折线图 |
| `"pie"` | xlPie | 饼图 |
| `"scatter"` | xlXYScatter | 散点图 |
| `"area"` | xlArea | 面积图 |
| `"radar"` | xlRadar | 雷达图 |

#### LLM 输出格式要求

当 `outputFormat="chart"` 时，LLM 必须输出以下 JSON 格式：

```json
{
  "labels": ["一月", "二月", "三月", "四月"],
  "datasets": [
    {
      "name": "销售额",
      "values": [120000, 150000, 135000, 180000]
    },
    {
      "name": "利润",
      "values": [30000, 45000, 35000, 55000]
    }
  ]
}
```

或简化格式：

```json
{
  "data": [
    {"key": "产品A", "value": 120},
    {"key": "产品B", "value": 150},
    {"key": "产品C", "value": 90}
  ]
}
```

#### 示例：销售数据图表生成器

```json
{
  "name": "SalesChartGenerator",
  "displayName": "销售数据图表生成器",
  "outputFormat": "chart",
  "processingMode": "Hybrid",
  "securityLevel": "NeedsConfirmation",      // 写入操作需确认
  "chartConfig": {
    "chartType": "bar",
    "title": "月度销售分析",
    "width": 600,
    "height": 400
  },
  "llmSystemPrompt": "分析选区数据并输出图表JSON:\n{\n  \"labels\": [...],\n  \"datasets\": [{\"name\": \"...\", \"values\": [...]}]\n}"
}
```

**使用流程**：
1. 用户选中 Excel 数据区域
2. 调用 SalesChartGenerator Skill
3. LLM 分析数据并生成 JSON
4. FormatBridge.RenderChart() 创建原生图表
5. 图表插入到工作表中（可编辑、可交互）

---

### 2.4 数据透视表（PivotTable）

#### 配置参数

```json
{
  "outputFormat": "pivotTable",
  "pivotConfig": {
    "sourceRange": "A1:D100",       // 数据源范围（空=当前选区）
    "targetRange": "F2",            // 目标位置（空=自动）
    "rowFields": ["部门", "产品"],   // 行字段
    "columnFields": ["季度"],         // 列字段
    "dataFields": [                  // 数据字段
      {"fieldName": "销售额", "aggregation": "Sum"}
    ],
    "filterFields": ["年份"]         // 筛选器字段
  }
}
```

#### 聚合函数支持

| aggregation 值 | Excel函数 | 说明 |
|----------------|-----------|------|
| `"Sum"` | xlSum | 求和（默认）|
| `"Average"` | xlAverage | 平均值 |
| `"Count"` | xlCount | 计数 |
| `"Max"` | xlMax | 最大值 |
| `"Min"` | xlMin | 最小值 |

#### 示例：智能透视表创建器

```json
{
  "name": "SmartPivotTableCreator",
  "displayName": "智能透视表创建器",
  "outputFormat": "pivotTable",
  "processingMode": "MultimodalLLM",
  "securityLevel": "NeedsConfirmation",
  "pivotConfig": {
    "rowFields": [],
    "columnFields": [],
    "dataFields": [],
    "filterFields": []
  },
  "llmSystemPrompt": "分析数据结构并输出透视表配置JSON"
}
```

**使用流程**：
1. 用户选中原始数据表
2. 调用 SmartPivotTableCreator
3. LLM 智能识别字段类型并推荐配置
4. 用户确认后生成交互式透视表
5. 可拖拽字段调整布局

---

### 2.5 条件格式（Conditional Formatting）

#### 配置参数

```json
{
  "outputFormat": "conditionalFormat"
  // 无需额外配置属性，由 LLM 动态生成规则
}
```

#### LLM 输出格式要求

LLM 必须输出 JSON 数组，每个元素是一条规则：

```json
[
  {
    "type": "cellvalue",
    "operator": "equal",
    "value": "",
    "format": "highlight_yellow",
    "description": "标记空值"
  },
  {
    "type": "colorscale",
    "minColor": "#FFFFFF",
    "midColor": "#FFFF00",
    "maxColor": "#FF0000",
    "description": "数值色阶"
  },
  {
    "type": "databar",
    "barColor": "#4472C4",
    "description": "数据条"
  },
  {
    "type": "iconset",
    "iconSet": "3arrows",
    "description": "图标集"
  }
]
```

#### 规则类型详解

##### 1) 单元格值规则 (cellvalue)

```json
{
  "type": "cellvalue",
  "operator": "greaterthan",     // 操作符
  "value": "1000",              // 比较值
  "format": "highlight_red"     // 样式
}
```

**operator 支持**：`greaterthan`(>), `lessthan`(<), `equal`(=), `between`, `notbetween`, `greaterorequal`(>=), `lessorequal`(<=), `notequal`(!=)

**format 样式**：

| 样式名 | 效果 |
|--------|------|
| `highlight_red` | 红色背景+深红文字 |
| `highlight_yellow` | 黄色背景+橙色文字 |
| `highlight_orange` | 橙色背景+深橙文字 |
| `highlight_green` | 绿色背景+深绿文字 |
| `bold` | 加粗 |
| `italic` | 斜体 |
| `strikethrough` | 删除线 |

##### 2) 色阶规则 (colorscale)

```json
{
  "type": "colorscale",
  "colorScaleType": 3,            // 3色阶（可选）
  "minColor": "#FFFFFF",          // 最小值颜色
  "midColor": "#FFFF00",          // 中间值颜色
  "maxColor": "#FF0000"           // 最大值颜色
}
```

##### 3) 数据条规则 (databar)

```json
{
  "type": "databar",
  "barColor": "#4472C4"           // 数据条颜色
}
```

##### 4) 图标集规则 (iconset)

```json
{
  "type": "iconset",
  "iconSet": "3arrows"            // 图标集类型
}
```

**图标集选项**：`3arrows`, `3flags`, `3trafficlights1`, `3trafficlights2`, `3signs`, `4arrows`, `4rating`, `5arrows`, `5quartiles`

#### 示例：数据质量高亮检查器

```json
{
  "name": "DataQualityHighlighter",
  "displayName": "数据质量高亮检查器",
  "outputFormat": "conditionalFormat",
  "processingMode": "Hybrid",
  "securityLevel": "Safe",               // 只读操作，安全
  "llmSystemPrompt": "检测数据质量问题并输出条件格式规则JSON数组"
}
```

**使用流程**：
1. 用户选中要检查的数据区域
2. 调用 DataQualityHighlighter
3. LLM 分析数据质量并生成规则
4. FormatBridge.ApplyConditionalFormatting() 应用规则
5. 问题数据立即被高亮标记

---

## 🔒 功能三：权限与审批机制

### 3.1 功能说明

**作用**：对敏感操作增加安全防护层，根据操作的危险程度要求不同级别的用户确认。

**为什么需要**：
- 防止误操作导致的数据丢失
- 高危操作（发送邮件、删除数据）需要二次确认
- 满足企业合规审计要求
- 提供完整的操作追溯日志

### 3.2 安全级别定义

在 Skill 定义中设置 `securityLevel` 属性：

```json
{
  "securityLevel": "Safe"  |  "NeedsConfirmation"  |  "Critical"
}
```

| 级别 | 行为 | 适用场景 | 用户体验 |
|------|------|----------|----------|
| **Safe** | 直接执行，仅记录审计日志 | 只读分析、格式高亮、条件标注 | 无感知，即时完成 |
| **NeedsConfirmation** | 弹窗确认一次后执行 | 写入Excel、创建图表、生成报表 | 点击"确认"即可 |
| **Critical** | 二次确认 + 详细审计日志 | MCP写操作、发邮件、删除数据 | 二次确认弹窗 + 风险提示 |

### 3.3 审批流程详解

#### Safe 级别流程

```
ExecuteCustomSkill() 调用
    ↓
ApprovalGate.CheckAndApproveAsync(skill)
    ↓
判断: securityLevel == Safe?
    ↓ Yes
✅ 直接通过
    ↓
记录审计日志: [AUTO_APPROVED] Safe级别自动放行
    ↓
继续执行 Skill...
```

#### NeedsConfirmation 级别流程

```
ExecuteCustomSkill() 调用
    ↓
ApprovalGate.CheckAndApproveAsync(skill)
    ↓
判断: securityLevel == NeedsConfirmation?
    ↓ Yes
📤 发送审批请求到前端 WebView:
   {
     type: "approvalRequest",
     requestId: "abc123",
     skillName: "销售图表生成器",
     description: "即将执行 Skill: 销售图表生成器\n\n- 操作类型: AI 分析 + VBA 执行...",
     securityLevel: "NeedsConfirmation",
     isCritical: false,
     singleConfirm: true
   }
    ↓
⏳ 等待前端响应（最多120秒）
    ↓
用户点击:
  ├─ ✅ [确认] → ApprovalGate.Approve("abc123")
  │   ↓
  │   记录审计: [APPROVED] 用户确认执行
  │   ↓
  │   返回 True → 继续 Skill 执行
  │
  └─ ❌ [取消] → ApprovalGate.Reject("abc123")
      ↓
      记录审计: [REJECTED] 用户拒绝执行
      ↓
      返回 False → 终止执行，返回错误信息
```

#### Critical 级别流程

```
ExecuteCustomSkill() 调用
    ↓
ApprovalGate.CheckAndApproveAsync(skill)
    ↓
判断: securityLevel == Critical?
    ↓ Yes
📝 先记录高危预警日志: [CRITICAL_REQUESTED] 高危操作已触发
    ↓
📤 发送审批请求到前端（带危险标识）:
   {
     type: "approvalRequest",
     isCritical: true,             // ← 关键差异
     singleConfirm: false          // ← 要求二次确认
   }
    ↓
⏳ 等待前端响应（最多120秒）
    ↓
用户看到强化版确认弹窗:
  ┌────────────────────────────────────┐
  │ ⚠️  高危操作确认                    │
  │                                    │
  │ 技能: 批量邮件发送器                │
  │                                    │
  │ [危险操作详情]                      │
  │ - 将调用外部 SMTP 服务             │
  │ - 发送邮件给 Selection.Rows.Count 个收件人 │
  │                                    │
  │ 此操作将被记录到审计日志            │
  │                                    │
  │  [第一次确认]    [取消]             │
  │                                    │
  │  [第二次确认]  ← 第一次点击后出现   │
  └────────────────────────────────────┘
    ↓
二次确认通过 → 记录详细审计 → 执行 Skill
```

### 3.4 审计日志

所有经过审批的 Skill 执行都会记录到审计日志文件：

**日志位置**：`%AppData%\SmartPilot\skill_audit.log`

**日志格式**：

```
[2026-06-09 14:30:15.123] [AUTO_APPROVED] Skill=DataQualityHighlighter | Safe级别自动放行
[2026-06-09 14:31:22.456] [APPROVED] Skill=SalesChartGenerator | 用户确认执行
[2026-06-09 14:32:10.789] [REJECTED] Skill=BatchEmailSender | 用户拒绝执行
[2026-06-09 14:33:05.012] [CRITICAL_REQUESTED] Skill=BatchEmailSender | 高危操作已触发: 根据Excel客户列表批量发送邮件
[2026-06-09 14:33:45.678] [APPROVED] Skill=BatchEmailSender | 用户二次确认通过
[2026-06-09 14:35:20.999] [TIMEOUT] Skill=SomeSkill | 用户未在时限内响应
[2026-06-09 14:36:11.111] [ERROR_FALLBACK] Skill=SomeSkill | 审批异常，降级放行: 连接超时
```

**查看最近日志**（调试用）：

```vb
Dim logs = SkillApprovalGate.Instance.GetRecentAuditLogs(50)
For Each log In logs
    Debug.WriteLine(log)
Next
```

### 3.5 示例配置

#### 示例 1：Safe 级别 - 数据质量检查

```json
{
  "name": "DataQualityHighlighter",
  "securityLevel": "Safe",              // 只读分析，直接执行
  "outputFormat": "conditionalFormat"
}
```

**用户体验**：无任何弹窗，即时完成高亮标记

---

#### 示例 2：NeedsConfirmation 级别 - 图表生成

```json
{
  "name": "SalesChartGenerator",
  "securityLevel": "NeedsConfirmation",  // 写入操作，需确认
  "outputFormat": "chart"
}
```

**用户体验**：弹出标准确认窗口，用户点击"确认"后继续

---

#### 示例 3：Critical 级别 - 批量邮件发送

```json
{
  "name": "BatchEmailSender",
  "securityLevel": "Critical",           // 高危操作，二次确认
  "dangerousOperations": ["SendEmail", "SmtpSend"],
  "processingMode": "VBACode"
}
```

**用户体验**：
1. 弹出醒目的危险警告窗口
2. 显示详细的操作描述（将影响多少收件人等）
3. 要求两次确认才能执行
4. 所有操作完整记录到审计日志

---

## 🚀 快速开始指南

### 步骤 1：选择合适的功能组合

根据您的需求选择功能组合：

| 场景 | 推荐功能组合 |
|------|-------------|
| 自动化数据处理 | 事件触发 + Safe权限 |
| 定时报表生成 | 定时触发 + 图表输出 + NeedsConfirmation |
| 数据质量监控 | 事件触发 + 条件格式 + Safe权限 |
| 敏感操作保护 | Critical权限 + 审计日志 |

### 步骤 2：编辑 Skill 配置

在 `Config/DefaultSkills.json` 中添加或修改 Skill：

**启用事件触发**：
```json
"triggerConfig": { "triggerType": "Event", "eventType": "Change", ... }
```

**启用增强输出**：
```json
"outputFormat": "chart"  // 或 pivotTable / conditionalFormat
```

**设置安全级别**：
```json
"securityLevel": "NeedsConfirmation"  // Safe / NeedsConfirmation / Critical
```

### 步骤 3：测试验证

#### 测试事件触发

1. 创建一个名为"订单表"的工作表
2. 在B列输入金额
3. 观察1秒后C列是否自动计算佣金
4. 查看 Debug 日志中的 `[SkillTriggerEngine]` 信息

#### 测试图表输出

1. 准备包含类别和数值的数据
2. 选中数据区域
3. 调用 SalesChartGenerator Skill
4. 确认后观察是否生成可交互的图表

#### 测试权限审批

1. 调用 BatchEmailSender（Critical级别）
2. 观察是否弹出二次确认窗口
3. 检查 `%AppData%\SmartPilot\skill_audit.log` 日志

---

## 📊 监控与调试

### 日志关键字

| 关键字 | 含义 | 出现位置 |
|--------|------|----------|
| `[SkillTriggerEngine]` | 触发引擎相关 | SkillTriggerEngine.vb |
| `已注册事件触发` | 成功注册事件监听 | RegisterEventTrigger() |
| `已注册定时触发` | 成功启动定时器 | RegisterTimerTrigger() |
| `步骤 X 条件不满足，跳过` | FilterCondition 为False | HandleEventTrigger() |
| `触发执行异常` | Skill执行出错 | HandleEventTrigger() |
| `[ApprovalGate]` | 审批相关 | SkillApprovalGate.vb |
| `AUTO_APPROVED` | Safe级别自动放行 | CheckAndApproveAsync() |
| `APPROVED` | 用户确认 | Approve()回调 |
| `REJECTED` | 用户拒绝 | Reject()回调 |
| `CRITICAL_REQUESTED` | 高危操作触发 | CheckAndApproveAsync() |
| `TIMEOUT` | 用户超时未响应 | RequestConfirmationAsync() |
| `已创建.*图表` | 图表渲染成功 | FormatBridge.RenderChart() |
| `已创建数据透视表` | 透视表创建成功 | FormatBridge.RenderPivotTable() |
| `已应用.*条条件格式` | 条件格式应用成功 | FormatBridge.ApplyConditionalFormatting() |

### 常见问题排查

#### 事件触发问题

**问题**：修改单元格但没有触发 Skill
- 检查 `triggerConfig.enabled` 是否为 `true`
- 检查 `eventType` 是否正确（Change/Open/Close/SelectionChange）
- 检查 `filterCondition` 是否过于严格
- 确认 SkillTriggerEngine 已初始化（ThisAddIn_Startup）
- 查看日志是否有 `已注册事件触发` 记录

**问题**：触发太频繁或太少
- 调整 `debounceMs` 值（单位：毫秒）
- 输入场景建议 500-1000ms
- 批量操作建议 1000-2000ms

#### 输出格式问题

**问题**：图表没有生成
- 检查 `outputFormat` 是否设置为 `"chart"`
- 确认 LLM 输出的是有效 JSON（labels + datasets）
- 验证 datasets 中的 values 都是数字类型
- 检查 chartConfig.chartType 是否正确

**问题**：透视表报错"无法确定数据源范围"
- 设置 `pivotConfig.sourceRange` 或确保已选中数据区域
- 确认源数据有表头行
- 检查字段名称是否与实际列标题匹配

#### 权限审批问题

**问题**：没有弹出确认窗口
- 检查 `securityLevel` 不是 `"Safe"`
- 确认前端 WebView 正常运行
- 查看日志是否有 `approvalRequest` 消息发送记录
- 检查是否超过120秒超时

**问题**：找不到审计日志
- 确认路径：`%AppData%\SmartPilot\skill_audit.log`
- 如果目录不存在，手动触发一次 Critical 操作会自动创建
- 检查是否有文件写入权限

---

## 💡 最佳实践

### 事件驱动最佳实践

1. **精准的条件过滤**
   ```json
   // ✅ 好：精确定位
   "filterCondition": "targetColumn = 'B' && targetSheet = '订单表'"

   // ❌ 差：过于宽泛
   "filterCondition": ""
   ```

2. **合理的防抖时间**
   - 用户输入：1000ms（避免每次按键都触发）
   - 数据导入：2000ms（等待批量操作完成）
   - 实时监控：100-300ms（需要即时反馈）

3. **幂等的 Skill 设计**
   - 事件触发的 Skill 应该可以重复执行而不产生副作用
   - 避免追加式写入（应该更新或覆盖）

### 输出格式最佳实践

1. **明确的 Prompt 引导**
   ```
   // ✅ 好：明确要求 JSON 格式
   "请输出严格的 JSON 格式，不要包含其他文字：\n{...}"

   // ❌ 差：模糊的要求
   "帮我分析一下数据"
   ```

2. **错误处理**
   - LLM 可能输出非标准 JSON，做好解析容错
   - 提供默认配置作为回退方案

3. **性能考虑**
   - 大数据量（>1000行）生成图表可能较慢
   - 透视表建议限制源数据范围
   - 条件格式规则数量建议 < 50 条

### 权限审批最佳实践

1. **合理的安全分级**
   - 只读操作 → Safe
   - 写入当前工作簿 → NeedsConfirmation
   - 外部系统交互/删除操作 → Critical

2. **清晰的描述文案**
   ```json
   // ✅ 好：具体说明影响
   "description": "将在 C2:C100 区域写入计算结果，覆盖现有数据"

   // ❌ 差：模糊不清
   "description": "执行某些操作"
   ```

3. **审计日志定期审查**
   - 建议每周检查审计日志
   - 关注 REJECTED 和 TIMEOUT 记录
   - Critical 操作应重点审核

---

## 🔄 与其他功能的协同

### 事件触发 + RAG

```json
{
  "triggerConfig": { "triggerType": "Event", ... },
  "ragEnabled": true,                    // 事件触发时也检索知识库
  "llmSystemPrompt": "基于知识库回答用户关于{{cellValue}}的问题"
}
```

**适用场景**：用户输入问题时自动查询知识库回答

### 事件触发 + Pipeline

```json
{
  "isPipeline": true,
  "pipelineSteps": [...],
  "triggerConfig": { "triggerType": "Event", ... }
}
```

**适用场景**：数据变化时自动执行多步处理流程

### 输出格式 + Pipeline

Pipeline 的某个步骤可以输出图表/透视表：

```json
{
  "isPipeline": true,
  "pipelineSteps": [
    {"stepId": "analyze", "outputFormat": "json"},
    {"stepId": "visualize", "outputFormat": "chart"}  // 最后一步生成图表
  ]
}
```

### 权限审批 + Pipeline

整个管道只需确认一次（在入口处）：

```json
{
  "isPipeline": true,
  "securityLevel": "NeedsConfirmation",  // 管道入口确认
  "pipelineSteps": [...]                   // 内部步骤不再重复确认
}
```

---

## 📝 版本历史

| 版本 | 日期 | 变更内容 |
|------|------|----------|
| v1.0 | 2026-06-09 | 初始版本，记录事件驱动、输出增强、权限审批的使用方法 |

---

## 🔗 相关文档

- [第一优先级扩展指南](./Priority1-Extension-Guide.md) - RAG 和 Pipeline 功能
- [Skill 系统架构](./SkillCreator.md) - Skill 定义和元编程
- [Skill 扩展路线图](../.trae/documents/skill-extension-roadmap.md) - 全部 6 个扩展方向

---

## 🎯 总结

第二优先级的三个功能将 SmartPilot 从"被动工具"提升为"智能自动化平台"：

| 维度 | 扩展前 | 扩展后 |
|------|--------|--------|
| **自动化程度** | 用户手动触发 | ✅ 事件驱动/定时执行 |
| **输出能力** | 文本/代码 | ✅ 原生图表/透视表/条件格式 |
| **安全性** | 无防护 | ✅ 三级权限审批+审计日志 |
| **适用场景** | 简单任务 | ✅ 企业级复杂业务流 |
| **用户体验** | 单次操作 | ✅ 全自动智能化工作流 |

**立即可用**：示例配置已在 DefaultSkills.json 中，编译后即可测试！
