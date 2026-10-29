---
hosts: [excel]
name: GeneratePython
displayName: Python数据分析（执行）
description: 使用预置分析模板进行数据分析（描述统计、趋势分析、相关性分析、异常检测、数据分布、分组聚合）。支持用户明确要求时生成自定义 Python 脚本。模板代码经过测试，比 AI 自由生成更可靠。
keywords:
  - 数据分析
  - 分析数据
  - 分析一下
  - 分析当前
  - 分析工作表
  - 描述统计
  - 均值
  - 中位数
  - 标准差
  - 趋势分析
  - 趋势
  - 找趋势
  - 预测
  - 相关性
  - 相关性分析
  - 相关系数
  - 异常值
  - 异常检测
  - 找异常
  - 数据分布
  - 分布
  - 频率分布
  - 分组聚合
  - 分组
  - 聚合
  - 统计
  - Python
  - python
  - 机器学习
  - 波动
  - 假设检验
  - 方差分析
  - pandas
  - numpy
  - sklearn
  - 回归
  - 分类
  - 聚类
  - 时间序列
  - matplotlib
  - 可视化
  - 文本分析
  - NLP
  - python代码
  - python脚本
  - 图表
  - 图表展示
  - 业绩图表
  - 部门分析
  - 按部门
  - 按类别
  - 按维度
examples:
  - 分析选中数据的描述统计
  - 趋势分析并预测下季度销售额
  - 相关性分析
  - 检测异常值
  - 数据分布分析
  - 按部门分组聚合
  - 用Python预测下季度销售额
  - 用pandas分析一下这份数据
  - 生成一个时间序列预测的Python代码
  - 用Python做相关性分析
  - 按部门展示业绩图表
  - 按类别聚合数据
  - 生成部门业绩图表
  - 按维度展示数据图表
isGeneric: true
category: 数据分析工具
needsConfig: false
configFields: []
processingMode: PythonCode
supportsImageInput: false
supportsFileInput: true
outputFormat: json
isCustom: false
tags: Python代码生成,数据分析,描述统计,趋势分析,相关性分析,异常检测,数据分布,分组聚合
riskLevel: medium
riskDesc: Python脚本执行分析，可能消耗计算资源
inputHint: 描述数据分析需求，如：描述统计、趋势分析、相关性分析
maxOutputLength: 2000
executeTimeout: 120000
limitDesc: 使用预置模板，代码量可控
dependencies:
  - PythonRuntime
tlDrRequired: true
version: 2
pipeline: python
quickPrompts:
  - label: 趋势分析
    prompt: 用Python对当前数据做趋势分析
  - label: 相关性分析
    prompt: 用Python对当前数据做相关性分析
  - label: 异常检测
    prompt: 用Python对当前数据做异常检测
  - label: 分组聚合
    prompt: 用Python对当前数据做分组聚合统计
---
