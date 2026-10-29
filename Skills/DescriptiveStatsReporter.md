---
hosts: [excel]
name: DescriptiveStatsReporter
displayName: 描述性统计报告
description: 一键生成描述性统计报告：均值、中位数、标准差、分位数、频率分布，支持按组统计。只做统计计算和报告输出，不检测异常，不修复数据。
keywords:
  - 统计
  - 描述统计
  - 均值
  - 中位数
  - 标准差
  - 频率分布
  - 统计报告
examples:
  - 生成统计报告
  - 描述性统计
  - 计算均值和标准差
  - 频率分布表
  - 清洗这些数据（反例：应走SmartDataCleaner）
  - 检测异常值（反例：应走SmartDataAnalyzer）
  - 校验数据质量并标红异常单元格（反例：应走DataQualityValidator）
isGeneric: true
category: 数据分析
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
  - name: groupColumn
    label: 分组列
    type: string
    required: false
pipeline: python
subSkills:
  - SmartChartWizard
processingMode: Hybrid
supportsImageInput: false
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 统计分析
  - 数据报告
  - 办公效率
quickPrompts:
  - label: 分组统计
    prompt: 对当前数据按分组列生成描述性统计报告
  - label: 频率分布
    prompt: 对当前数据生成频率分布表
  - label: 分位数
    prompt: 对当前数据计算分位数统计
  - label: 生成图表
    prompt: 根据当前数据的统计结果生成图表
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 描述性统计报告

## System Prompt

你是统计分析专家。

**职责边界**：本 Skill 只做统计计算和报告输出。不检测异常，不修复数据，不执行去重/填充/格式转换等清洗操作。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际统计报告输出，禁止只思考不输出。

【统计指标】
1. **集中趋势**: 均值(Mean)、中位数(Median)、众数(Mode)
2. **离散程度**: 标准差(STDEV)、方差(VAR)、极差(Range)、四分位距(IQR)
3. **分布形态**: 最小值、最大值、25%分位数、75%分位数、偏度、峰度
4. **频率分布**: 自动分箱并计算每组的频数和频率
5. **分组统计**: 按分类列分组计算上述指标

【输出格式】
生成结构化的统计报告 JSON：
{
  "overall": {
    "count": 100,
    "mean": 50.5,
    "median": 51,
    "stdDev": 15.2,
    "min": 10,
    "max": 99
  },
  "byGroup": {
    "部门A": {"mean": 55, "stdDev": 12},
    "部门B": {"mean": 48, "stdDev": 18}
  },
  "frequencyDistribution": [
    {"bin": "10-20", "count": 5, "percentage": "5%"}
  ],
  "result": {"列名": {"count": 100, "mean": 50.5, "median": 51, "stdev": 15.2, "min": 10, "max": 99, "q1": 30, "q3": 75}}
}

【规则】
- 自动识别数值列和分类列
- 对非数值列跳过统计计算
- 空值不计入统计
- 结果输出到新的"统计摘要"工作表
- 附带简要的数据解读

## Template Code

```python
import json, statistics

data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []
if not data or not isinstance(data, list) or len(data) == 0:
    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域后重试'}
else:
    target_cols = [c for c in list(data[0].keys()) if c != '分组列']
    headers = list(data[0].keys()) if isinstance(data[0], dict) else []
    results = {}
    for col in target_cols:
        if col not in headers:
            continue
        vals = []
        for row in data:
            try:
                v = row.get(col) if isinstance(row, dict) else None
                if v is not None and str(v).strip():
                    vals.append(float(v))
            except (ValueError, TypeError):
                pass
        if vals:
            n = len(vals)
            mean = sum(vals) / n
            s = sorted(vals)
            median = s[n//2] if n % 2 else (s[n//2-1] + s[n//2]) / 2
            variance = sum((x - mean)**2 for x in vals) / (n - 1) if n > 1 else 0
            results[col] = {
                'count': n,
                'mean': round(mean, 4),
                'median': round(median, 4),
                'stdev': round(variance**0.5, 4),
                'min': round(min(vals), 4),
                'max': round(max(vals), 4),
                'q1': round(s[n//4], 4),
                'q3': round(s[3*n//4], 4)
            }
    if not results:
        output = {'success': False, 'message': '所选列均无法计算描述统计'}
    else:
        print(json.dumps(results, ensure_ascii=False, indent=2))
        output = {'success': True, 'summary': '完成' + str(len(results)) + '列的描述统计', 'result': results}
```
