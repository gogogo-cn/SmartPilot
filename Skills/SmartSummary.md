---
hosts: [excel, word]
name: SmartSummary
displayName: 智能总结(AI)
description: 使用 AI 视觉模型对单张图片内容进行智能总结，输出文本描述。不做OCR字段提取，不做数据分析，不处理多文件对比。
keywords:
  - 总结
  - 智能总结
  - 图片总结
  - 分析图片
  - 图片内容
  - 总结图片
  - 概括
  - 提炼
  - 内容概括
  - 图片分析
examples:
  - 总结图片内容
  - 分析图片
  - 图片智能分析
  - 识别发票信息（反例：应创建发票识别Skill）
  - 对比这几张图片（反例：应走MultiFileAnalysis）
isGeneric: true
category: 效率工具
needsConfig: false
configFields:
  - name: image
    label: 图片
    type: image
    required: true
pipeline: skill
processingMode: MultimodalLLM
supportsImageInput: true
supportsFileInput: false
outputFormat: text
isCustom: false
tags:
  - 智能分析
  - 图片处理
  - 内容总结
quickPrompts:
  - label: 更详细
    prompt: 对当前图片做更详细的总结
  - label: 提取要点
    prompt: 提取当前图片的核心要点
  - label: 结构化总结
    prompt: 用结构化清单总结当前图片内容
dependencies:
  - MultimodalLLM
tlDrRequired: true
version: 9
---

# 智能总结(AI)

## System Prompt

你是一个专业的图片分析助手。

### 思考框架
由 ThinkingEngine 根据任务复杂度动态决定思考步骤。简单任务直接输出结果，复杂任务按需进行理解、约束检查、方案规划和风险评估。
`</thinking>`标签后**必须**紧跟实际总结内容输出，禁止只思考不输出。

请对图片内容进行详细分析和总结。

## Template Code

```vba
' 无模板代码
```
