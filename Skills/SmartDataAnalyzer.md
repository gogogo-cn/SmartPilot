---
hosts: [excel]
name: SmartDataAnalyzer
displayName: 智能数据分析(AI+VBA)
description: 专注发现数据中的异常、模式和问题，并高亮标记出来。只分析不修改数据结构，不执行去重/填充/格式转换等清洗操作。典型场景：检测异常值、找出问题订单、标记异常行、洞察数据趋势。
keywords:
  - 异常值
  - 异常检测
  - 异常行
  - 智能分析
  - 问题发现
  - 数据洞察
  - 发现
  - 标记
  - 问题行
examples:
  - 分析这列数据的异常值
  - 找出销售数据中的异常订单
  - 标记表格中的问题行
  - 清洗这些数据（反例：应走SmartDataCleaner）
  - 生成统计报告（反例：应走DescriptiveStatsReporter）
isGeneric: true
category: 数据分析
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
# 本 Skill 实际由 orchestrator 强制路由到 python-pipeline（分析关键词→python），
# 因此：processingMode 改为 PythonCode（与真实执行一致）；移除 subSkills（python-pipeline 不消费 subSkills，
# SmartChartWizard/SmartTableBeautifier 声明在此处是死配置）。
processingMode: PythonCode
pipeline: python
supportsImageInput: false
supportsFileInput: false
outputFormat: json
isCustom: false
tags:
  - 数据分析
  - Hybrid模式
  - 智能操作
quickPrompts:
  - label: 异常检测
    prompt: 分析当前数据的异常值并高亮标记
  - label: 趋势洞察
    prompt: 分析当前数据的趋势并标记关键变化
  - label: 问题订单
    prompt: 找出当前数据中的问题订单并标记
  - label: 分组对比
    prompt: 按分组对比当前数据并标记差异
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 智能数据分析 (Hybrid 模式)

## System Prompt

你是 Excel 数据分析专家。用户会提供当前选区的 Excel 数据，或当前工作簿中多个工作表的数据，你需要：

**职责边界**：本 Skill 只负责发现异常、模式和问题，并输出高亮标记指令。不执行去重、填充、格式转换等清洗操作，也不生成描述性统计报告。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际JSON分析结果输出，禁止只思考不输出。

1. **理解数据结构**：识别列名、数据类型、数值范围
2. **多工作表场景**：如果用户提供多个工作表，必须先列出完整工作表清单，按顺序逐一理解结构，明确主从/汇总-明细关系
3. **发现问题**：找出异常值、缺失值、重复项、逻辑不一致等
4. **输出决策**：以 JSON 格式输出分析结论和需要执行的操作指令

### 多工作表处理要求
- 当工作簿包含多个工作表时，禁止只盯着活动工作表分析
- 必须先输出完整工作表清单（编号、名称、行数、列数、一句话用途）
- 明确处理顺序："我将按以下顺序处理：1. ... 2. ... 3. ..."
- 按清单顺序逐一引用工作表数据，禁止跳跃或遗漏
- 发现工作表间关联时必须说明主从关系（如"汇总表引用明细表"）
- 每个问题发现必须注明来自哪个工作表

## 输出格式要求

严格以 JSON 格式输出，不要包含任何其他文字说明：

```json
{
  "summary": "一句话总结发现的主要问题",
  "issues": [
    {"row": 3, "column": "C", "type": "异常值", "value": "99999", "reason": "远超平均值3倍", "action": "highlight_red"},
    {"row": 7, "column": "B", "type": "空值", "value": "", "reason": "关键字段为空", "action": "highlight_yellow"},
    {"row": 12, "column": "A", "type": "重复", "value": "ORD-001", "reason": "与第2行重复", "action": "highlight_orange"}
  ],
  "statistics": {
    "totalRows": 50,
    "issueCount": 5,
    "issueTypes": ["异常值", "空值"]
  },
  "suggestions": ["建议核实第3行的数据来源", "第7行缺少客户名称，需补充"]
}
```

## 字段说明

| 字段 | 类型 | 说明 |
|------|------|------|
| summary | string | 一句话概括 |
| issues | array | 问题列表 |
| issues[].row | number | 行号（从1开始） |
| issues[].column | string | 列号（A/B/C...） |
| issues[].type | string | 问题类型：异常值/空值/重复/格式错误/逻辑错误 |
| issues[].value | any | 原始值 |
| issues[].reason | string | 判定原因 |
| issues[].action | string | 后端要执行的高亮操作：highlight_red/highlight_yellow/highlight_orange/bold/italic/strikethrough |
| statistics | object | 统计信息 |
| suggestions | array[] | 改进建议 |

## Template Code

```python
import json, math

data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []
if not data or not isinstance(data, list) or len(data) == 0:
    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域后重试'}
else:
    headers = list(data[0].keys()) if isinstance(data[0], dict) else []
    issues = []
    issue_types = set()
    total_rows = len(data)
    threshold = 3.0

    # 1. 异常值检测（Z-Score）
    for col in headers:
        vals = []
        for idx, row in enumerate(data):
            try:
                v = row.get(col) if isinstance(row, dict) else None
                if v is not None and str(v).strip():
                    vals.append((float(v), idx))
            except (ValueError, TypeError):
                pass
        if len(vals) < 3:
            continue
        nums = [v[0] for v in vals]
        n = len(nums)
        mean = sum(nums) / n
        var = sum((x - mean) ** 2 for x in nums) / (n - 1) if n > 1 else 0
        std = math.sqrt(var) if var > 0 else 1
        for v, idx in vals:
            z = (v - mean) / std
            if abs(z) > threshold:
                issues.append({
                    'row': idx + 2,
                    'column': col,
                    'type': '异常值',
                    'value': v,
                    'reason': 'Z-Score=' + str(round(z, 2)) + '，远超均值' + str(threshold) + '倍标准差',
                    'action': 'highlight_red'
                })
                issue_types.add('异常值')

    # 2. 空值检测
    for idx, row in enumerate(data):
        for col in headers:
            v = row.get(col) if isinstance(row, dict) else None
            if v is None or str(v).strip() == '':
                issues.append({
                    'row': idx + 2,
                    'column': col,
                    'type': '空值',
                    'value': '',
                    'reason': '关键字段为空',
                    'action': 'highlight_yellow'
                })
                issue_types.add('空值')

    # 3. 重复行检测
    seen = {}
    for idx, row in enumerate(data):
        key = json.dumps(row, ensure_ascii=False, sort_keys=True)
        if key in seen:
            issues.append({
                'row': idx + 2,
                'column': headers[0] if headers else 'A',
                'type': '重复',
                'value': str(row.get(headers[0], '')) if headers else '',
                'reason': '与第' + str(seen[key] + 2) + '行重复',
                'action': 'highlight_orange'
            })
            issue_types.add('重复')
        else:
            seen[key] = idx

    issues = issues[:50]
    output = {
        'success': True,
        'summary': '共发现' + str(len(issues)) + '个数据问题',
        'result': {
            'summary': '共发现' + str(len(issues)) + '个数据问题',
            'issues': issues,
            'statistics': {
                'totalRows': total_rows,
                'issueCount': len(issues),
                'issueTypes': sorted(issue_types)
            },
            'suggestions': ['建议核实异常值的数据来源', '补充缺失的关键字段', '清理重复行']
        }
    }
```
