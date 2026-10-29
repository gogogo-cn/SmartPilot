---
hosts: [excel]
name: DataNormalizer
displayName: 数据归一化
description: 对数值列做归一化/标准化：Min-Max 缩放到 [0,1]，或 Z-Score 标准化（均值0、标准差1）。输出与原始数据行对齐的结果并支持一键插入 Excel。只做数值变换，不检测异常，不修改原始数据。
keywords:
  - 归一化
  - 标准化
  - 正规化
  - min-max
  - minmax
  - z-score
  - zscore
  - z分数
  - 缩放到
  - 标准化处理
examples:
  - 对当前数据做归一化
  - 把销售额和利润标准化
  - 用 z-score 标准化这几列
  - 把数据缩放到 0 到 1
  - 描述性统计（反例：应走 DescriptiveStatsReporter）
  - 检测异常值（反例：应走 SmartDataAnalyzer）
isGeneric: true
category: 数据处理
needsConfig: false
configFields:
  - name: selection
    label: 数据区域
    type: range
    required: true
  - name: columns
    label: 归一化列
    type: string
    required: false
  - name: method
    label: 归一化方法
    type: string
    required: false
    defaultValue: minmax
pipeline: python
processingMode: Hybrid
supportsImageInput: false
supportsFileInput: false
outputFormat: json
isCustom: false
tags:
  - 数据变换
  - 数据预处理
quickPrompts:
  - label: Min-Max 归一化
    prompt: 对当前数值列做 Min-Max 归一化，缩放到 0 到 1
  - label: Z-Score 标准化
    prompt: 对当前数值列做 z-score 标准化
  - label: 指定列归一化
    prompt: 对当前数据的指定列做归一化
dependencies:
  - ExcelBridge
tlDrRequired: true
version: 1
---

# 数据归一化

## System Prompt

你是数据预处理专家。

**职责边界**：本 Skill 只做数值列的归一化/标准化变换，输出与原始数据行对齐的结果。不检测异常、不去重、不改数据类型、不修改原始区域。

【方法】
1. **Min-Max 归一化**：x' = (x - min) / (max - min)，缩放到 [0, 1]；常量列（max = min）输出 0。
2. **Z-Score 标准化**：x' = (x - mean) / std，得到均值 0、标准差 1；std = 0 时输出 0。
3. 非数值 / 空值 → 输出 null，不计入统计。
4. 保留原始列，新增 `<列名>_minmax` 或 `<列名>_zscore` 列。
5. 自动识别数值列；用户指定列时只处理指定列。

【执行说明】
本 Skill 声明 `pipeline: python`，确定性执行由 `python-pipeline.js` 的 `TEMPLATES.normalize` 模板承担（含参数自动兜底与结果表格渲染）。
下方 Template Code 为**等价参考实现**（自包含：自动识别数值列、默认 Min-Max），当前 `pipeline: python` 流程下不会被执行，仅作能力声明与降级路径参考。

## Template Code

```python
import json, math

data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []
if not data or not isinstance(data, list) or len(data) == 0:
    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域后重试'}
else:
    method = 'minmax'  # 可改为 'zscore'
    headers = list(data[0].keys()) if isinstance(data[0], dict) else []
    series = {}
    for col in headers:
        vals = []
        for row in data:
            try:
                v = float(row.get(col)) if isinstance(row, dict) else None
                if v is not None:
                    vals.append(v)
            except (ValueError, TypeError):
                pass
        if vals:
            series[col] = vals
    if not series:
        output = {'success': False, 'message': '未找到可归一化的数值列'}
    else:
        suffix = '_minmax' if method == 'minmax' else '_zscore'
        out = []
        for row in data:
            newrow = dict(row) if isinstance(row, dict) else {}
            for col in series:
                try:
                    v = float(row.get(col)) if isinstance(row, dict) else None
                except (ValueError, TypeError):
                    v = None
                if v is None:
                    newrow[col + suffix] = None
                    continue
                arr = series[col]
                if method == 'minmax':
                    mn, mx = min(arr), max(arr)
                    nv = (v - mn) / (mx - mn) if mx > mn else 0.0
                else:
                    n = len(arr)
                    mean = sum(arr) / n
                    var = sum((x - mean) ** 2 for x in arr) / (n - 1) if n > 1 else 0
                    std = var ** 0.5
                    nv = (v - mean) / std if std > 0 else 0.0
                newrow[col + suffix] = round(nv, 4)
            out.append(newrow)
        print(json.dumps(out, ensure_ascii=False, indent=2))
        output = {'success': True, 'summary': '完成' + str(len(series)) + '列归一化（方法=' + method + '），共' + str(len(out)) + '行（已新增' + suffix + '列）', 'result': out}
```
