# 第一优先级扩展功能使用指南

## 📋 概述

本文档详细介绍 SmartPilot 插件的两个核心扩展功能：

1. **RAG 知识库联动** - 让 Skill 基于企业知识库回答问题
2. **Skill 管道编排（Pipeline）** - 多个 Skill 串联执行，形成自动化流程

---

## 🎯 功能一：RAG 知识库联动

### 1.1 功能说明

**作用**：在 Skill 执行前，自动从企业知识库（Qdrant向量数据库）中检索相关文档内容，并将这些上下文信息注入到 LLM 的 System Prompt 中，使 AI 回答基于真实的企业知识。

**适用场景**：
- 企业内部知识问答（政策、流程、规范）
- 技术文档查询
- FAQ 自动应答
- 合规性检查
- 培训材料检索

### 1.2 配置参数

在 Skill 定义 JSON 中添加以下属性：

```json
{
  "name": "YourSkillName",
  "ragEnabled": true,                    // 【必填】启用 RAG 开关
  "ragKbId": "kb_enterprise_policy",     // 【可选】目标知识库 ID（空=使用默认）
  "ragTopK": 5,                          // 【可选】检索返回的文档数量（默认3）
  "ragScoreThreshold": 0.6,              // 【可选】相似度阈值（0-1，默认0.5）
  "ragFallbackToNormal": true            // 【可选】无结果时降级到普通模式（默认true）
}
```

### 1.3 工作原理

```
用户提问: "公司的报销流程是什么"
    ↓
检查 skill.ragEnabled == true?
    ↓ Yes
调用 RagOrchestrator.RetrieveAndEnhanceAsync(question, kbId)
    ↓
在 Qdrant 向量数据库中检索 Top-K 相关文档
    ↓
构建增强后的 System Prompt:
┌─────────────────────────────────────┐
│ [原始 System Prompt]                │
│ 你是企业知识助手...                  │
│                                     │
│ ========== 参考知识（来自企业知识库）======== │
│ --- 资料 1 (相关度: 92%, 来源: 财务制度.docx) --- │
│ 报销流程：                          │
│ 1. 填写报销单...                    │
│ 2. 部门经理审批...                  │
│ 3. 财务审核...                      │
│                                     │
│ 【来源】: 财务制度.docx              │
│ ======================================== │
└─────────────────────────────────────┘
    ↓
LLM 基于增强后的 Prompt 生成回答
    ↓
返回给用户（包含来源引用）
```

### 1.4 示例配置

#### 示例 1：企业知识问答 Skill

```json
{
  "name": "EnterpriseKnowledgeQA",
  "displayName": "企业知识问答(RAG)",
  "description": "基于企业知识库的智能问答系统",
  "processingMode": "MultimodalLLM",
  "llmSystemPrompt": "你是企业知识助手。请基于【参考知识】回答用户问题。\n\n【重要规则】\n1. 优先使用参考资料中的信息回答\n2. 如果资料中没有相关信息，明确告知用户\n3. 不要编造资料中没有的信息\n4. 回答时引用来源文档名称",
  "ragEnabled": true,
  "ragKbId": "",                    // 使用默认知识库
  "ragTopK": 5,
  "ragScoreThreshold": 0.6,
  "ragFallbackToNormal": true       // 无结果时降级
}
```

#### 示例 2：技术文档查询 Skill

```json
{
  "name": "TechDocSearch",
  "displayName": "技术文档查询",
  "description": "从技术知识库中查找API文档、开发指南等",
  "processingMode": "MultimodalLLM",
  "llmSystemPrompt": "你是技术文档助手。根据参考资料提供准确的技术解答，包含代码示例。",
  "ragEnabled": true,
  "ragKbId": "kb_tech_docs",        // 指定技术文档知识库
  "ragTopK": 3,
  "ragScoreThreshold": 0.7,          // 更高的阈值确保准确性
  "ragFallbackToNormal": false       // 无结果时明确提示
}
```

### 1.5 使用效果对比

| 场景 | 无 RAG | 有 RAG |
|------|--------|--------|
| 用户问："报销流程" | AI 可能编造通用流程 | 返回公司实际制度文件内容 |
| 回答准确性 | 取决于训练数据 | 基于企业真实文档 |
| 来源引用 | 无 | ✅ 显示文档名称和来源 |
| 信息时效性 | 取决于模型更新时间 | 实时反映知识库最新内容 |

---

## 🔗 功能二：Skill 管道编排（Pipeline）

### 2.1 功能说明

**作用**：将多个 Skill 串联成处理管道，上一个 Skill 的输出自动作为下一个 Skill 的输入。适用于复杂的多步骤业务流程。

**适用场景**：
- 发票/票据处理（OCR→提取→校验→录入）
- 数据分析流水线（采集→清洗→分析→报告）
- 文档处理（识别→分类→归档→索引）
- 审批流程（提交→初审→复审→归档）

### 2.2 配置参数

在 Skill 定义 JSON 中添加以下属性：

```json
{
  "name": "MyPipelineSkill",
  "isPipeline": true,                    // 【必填】标记为管道型 Skill
  "pipelineSteps": [                     // 【必填】管道步骤列表
    {
      "stepId": "step_1",               // 步骤唯一标识
      "stepName": "第一步名称",
      "skillName": "SubSkillA",         // 要调用的子 Skill 名称
      "inputMapping": {                 // 输入字段映射
        "param1": "上游输出字段名"
      },
      "outputFilter": "$.field",        // 输出过滤（JSONPath，可选）
      "nextStepOnSuccess": "step_2",   // 成功后的下一步骤ID
      "nextStepOnFailure": "",          // 失败后的下一步骤（空=终止管道）
      "condition": "变量 != null",      // 执行条件（可选）
      "retryCount": 1,                   // 重试次数
      "timeoutSeconds": 30              // 超时秒数
    },
    // ... 更多步骤
  ]
}
```

### 2.3 PipelineStepDefinition 完整参数说明

| 参数名 | 类型 | 必填 | 默认值 | 说明 |
|--------|------|------|--------|------|
| `stepId` | String | ✅ | - | 步骤唯一标识符 |
| `stepName` | String | ❌ | "" | 步骤显示名称（用于日志） |
| `skillName` | String | ✅ | - | 要调用的子 Skill 名称（必须已注册） |
| `inputMapping` | Dictionary | ❌ | {} | 输入映射：{"本步输入参数": "上游输出字段"} |
| `outputFilter` | String | ❌ | "" | 输出过滤表达式（JSONPath） |
| `nextStepOnSuccess` | String | ❌ | "" | 成功后执行的下一步骤 ID |
| `nextStepOnFailure` | String | ❌ | "" | 失败后执行的下一步骤 ID（空则终止） |
| `condition` | String | ❌ | "" | 条件表达式（满足才执行此步骤） |
| `retryCount` | Integer | ❌ | 0 | 失败重试次数 |
| `timeoutSeconds` | Integer | ❌ | 30 | 单步超时时间（秒） |

### 2.4 数据流转机制

```
用户输入: {imagePath: "C:\invoice.jpg"}
    ↓
[Step 1: ocr_recognition]
  输入: {imagePath: "input.imagePath", text: "..."}
  输出: {content: "OCR识别的文字内容..."}
  → 存储为: step_ocr_recognition.content = "..."
            pipeData["content"] = "..." (扁平化)
    ↓
[Step 2: data_extraction]  (条件: step_ocr_recognition.content != null)
  输入: {text: "step_ocr_recognition.content"}  ← 引用上一步输出
  输出: {content: "{invoiceNumber: '12345', ...}"}
  → 存储为: step_data_extraction.content = "{...}"
    ↓
[Step 3: validation]
  输入: {text: "step_data_extraction.content"}
  输出: {content: "{isValid: true, ...}"}
    ↓
管道完成 → 最终结果 = pipeData (包含所有中间数据)
```

### 2.5 InputMapping 映射规则

支持三种数据源：

```json
{
  "inputMapping": {
    "myParam": "input.originalField",           // 规则1: 引用用户原始输入
    "text": "step_ocr_recognition.content",     // 规则2: 引用上一步输出 (格式: step_{stepId}.{field})
    "fixedValue": "\"常量值\""                  // 规则3: 直接指定固定值（带引号）
  }
}
```

### 2.6 Condition 条件表达式语法

支持简单比较操作：

```javascript
// 格式: 变量 操作符 值
"variableName != null"          // 变量存在且非空
"step_1.amount > 1000"          // 数值比较
"step_2.status == \"success\""  // 字符串相等
"result.isValid == true"        // 布尔判断
```

支持的运算符：`==`, `!=`, `<>`, `>`, `<`, `>=`, `<=`

### 2.7 示例配置

#### 示例 1：发票处理流水线（完整版）

```json
{
  "name": "InvoiceProcessingPipeline",
  "displayName": "发票处理流水线(Pipeline)",
  "description": "自动化发票处理：OCR识别 → 数据提取 → 校验 → 写入Excel",
  "isPipeline": true,
  "processingMode": "Hybrid",
  "supportsImageInput": true,
  "outputFormat": "json",
  "pipelineSteps": [
    {
      "stepId": "ocr_recognition",
      "stepName": "OCR文字识别",
      "skillName": "SmartSummary",
      "inputMapping": {
        "imagePath": "input.imagePath",
        "text": "\"请识别这张发票图片中的所有文字内容\""
      },
      "nextStepOnSuccess": "data_extraction",
      "retryCount": 1,
      "timeoutSeconds": 30
    },
    {
      "stepId": "data_extraction",
      "stepName": "结构化数据提取",
      "skillName": "SmartSummary",
      "inputMapping": {
        "text": "step_ocr_recognition.content",
        "prompt": "\"请从以下文本中提取发票JSON数据\""
      },
      "condition": "step_ocr_recognition.content != null",
      "nextStepOnSuccess": "validation",
      "retryCount": 1
    },
    {
      "stepId": "validation",
      "stepName": "数据合法性校验",
      "skillName": "SmartSummary",
      "inputMapping": {
        "text": "step_data_extraction.content",
        "prompt": "\"请验证发票数据的合法性\""
      },
      "nextStepOnSuccess": "",
      "retryCount": 0
    }
  ]
}
```

#### 示例 2：数据分析流水线

```json
{
  "name": "SalesAnalysisPipeline",
  "displayName": "销售数据分析流水线",
  "description": "Excel数据 → 清洗 → 统计分析 → 生成报告",
  "isPipeline": true,
  "processingMode": "Hybrid",
  "supportsImageInput": false,
  "outputFormat": "markdown",
  "pipelineSteps": [
    {
      "stepId": "data_cleaning",
      "stepName": "数据清洗",
      "skillName": "GenerateVBA",
      "inputMapping": {
        "prompt": "\"编写VBA代码清洗当前选区数据：删除空行、去除首尾空格、统一日期格式\""
      },
      "nextStepOnSuccess": "analysis"
    },
    {
      "stepId": "analysis",
      "stepName": "统计分析",
      "skillName": "SmartSummary",
      "inputMapping": {
        "excelData": "input.excelData",
        "text": "\"请对以上Excel数据进行统计分析\""
      },
      "nextStepOnSuccess": "report"
    },
    {
      "stepId": "report",
      "stepName": "生成报告",
      "skillName": "SmartSummary",
      "inputMapping": {
        "text": "step_analysis.content",
        "prompt": "\"请将分析结果整理成Markdown格式的报告\""
      }
    }
  ]
}
```

---

## 🚀 快速开始指南

### 步骤 1：准备环境（仅 RAG 需要）

如果使用 RAG 功能，需先配置：

1. 安装并启动 Qdrant 向量数据库
2. 在插件设置中配置 RAG 连接参数
3. 导入企业文档到知识库

```bash
# 启动 Qdrant（Docker方式）
docker run -p 6333:6333 qdrant/qdrant
```

### 步骤 2：创建或修改 Skill

编辑 `Config/DefaultSkills.json`，添加新 Skill 或修改现有 Skill：

**启用 RAG**：添加 `"ragEnabled": true`
**创建 Pipeline**：添加 `"isPipeline": true` 和 `"pipelineSteps": [...]`

### 步骤 3：测试验证

#### 测试 RAG 功能

1. 确保知识库中有文档
2. 触发 EnterpriseKnowledgeQA Skill
3. 观察日志输出 `[RAG] 增强 Prompt 成功`
4. 检查回答是否包含来源引用

#### 测试 Pipeline 功能

1. 准备测试输入（如发票图片）
2. 触发 InvoiceProcessingPipeline Skill
3. 观察日志输出各步骤执行情况
4. 检查最终输出是否包含所有步骤的结果

---

## 📊 监控与调试

### 日志关键字

| 关键字 | 含义 | 出现位置 |
|--------|------|----------|
| `[RAG]` | RAG 相关操作 | CustomSkillExecutor |
| `[Pipeline]` | 管道执行过程 | CustomSkillExecutor |
| `增强 Prompt 成功` | RAG 检索成功 | EnrichWithRagContextAsync |
| `降级到普通模式` | RAG 无结果降级 | EnrichWithRagContextAsync |
| `步骤 X 执行成功` | Pipeline 步骤完成 | ExecutePipelineSkill |
| `条件不满足，跳过` | Condition 为 False | ExecutePipelineSkill |

### 常见问题排查

#### RAG 问题

**问题**：回答未包含知识库内容
- 检查 `ragEnabled` 是否为 `true`
- 检查 Qdrant 服务是否运行
- 检查知识库是否有文档
- 查看 Debug 日志中的 `[RAG]` 信息

**问题**：总是降级到普通模式
- 降低 `ragScoreThreshold`（如从 0.7 降到 0.5）
- 增加 `ragTopK`（如从 3 改为 5）
- 确认知识库文档与问题相关

#### Pipeline 问题

**问题**：管道在某个步骤失败
- 检查 `nextStepOnFailure` 配置
- 查看 ExecutionLog 中的错误信息
- 确认子 Skill 名称正确且已注册

**问题**：步骤间数据传递失败
- 检查 InputMapping 的 key 是否正确
- 确认使用正确的引用格式：`step_{stepId}.{fieldName}`
- 查看上一步是否成功执行并有输出

---

## 💡 最佳实践

### RAG 最佳实践

1. **精准的知识库分区**
   - 按主题建多个知识库（财务、技术、HR）
   - 通过 `ragKbId` 指定目标库，提高检索精度

2. **合理的阈值设置**
   - 严格场景（合规、法务）：`ragScoreThreshold >= 0.7`
   - 一般问答：`ragScoreThreshold = 0.5-0.6`
   - 探索性搜索：`ragScoreThreshold = 0.3-0.4`

3. **Prompt 工程配合**
   - 明确告诉 LLM 优先使用参考资料
   - 要求引用来源
   - 无相关知识时的兜底话术

### Pipeline 最佳实践

1. **原子化步骤设计**
   - 每个步骤只做一件事
   - 避免步骤过于复杂

2. **容错与重试**
   - 关键步骤设置 `retryCount >= 1`
   - 配置 `nextStepOnFailure` 实现回退逻辑

3. **条件分支**
   - 使用 `condition` 实现动态流程控制
   - 数据不满足条件时跳过非必要步骤

4. **日志与监控**
   - 每个 Step 设置有意义的 `stepName`
   - 定期检查 ExecutionLog 排查问题

---

## 🔄 与其他功能的协同

### RAG + Hybrid 模式

Hybrid 模式的 Phase 1（LLM 分析）也支持 RAG 增强：

```json
{
  "processingMode": "Hybrid",
  "ragEnabled": true,           // Phase 1 会注入知识库上下文
  "templateCode": "Sub ProcessData()\n  ' 使用 {{ANALYSIS_JSON.result}}\nEnd Sub"
}
```

**适用场景**：基于企业规范进行数据分析后，生成 VBA 代码操作 Excel

### Pipeline + RAG

管道中的某个步骤可以单独启用 RAG：

```json
{
  "isPipeline": true,
  "pipelineSteps": [
    {"stepId": "step1", "skillName": "NormalSkill"},
    {"stepId": "step2", "skillName": "EnterpriseKnowledgeQA"},  // 这个步骤本身启用了 RAG
    {"stepId": "step3", "skillName": "GenerateVBA"}
  ]
}
```

### Pipeline + 安全审批

Pipeline 中的每个子 Skill 都会经过安全审批检查：

```json
{
  "isPipeline": true,
  "securityLevel": "NeedsConfirmation",  // 整个管道需要确认一次
  "pipelineSteps": [...]
}
```

---

## 📝 版本历史

| 版本 | 日期 | 变更内容 |
|------|------|----------|
| v1.0 | 2026-06-09 | 初始版本，记录 RAG 和 Pipeline 功能的使用方法 |

---

## 🔗 相关文档

- [Skill 系统架构](./SkillCreator.md) - Skill 定义和元编程
- [RAG 知识库集成计划](../.trae/documents/RAG-Knowledge-Base-Integration-Plan.md) - RAG 架构设计
- [Skill 扩展路线图](../.trae/documents/skill-extension-roadmap.md) - 全部 6 个扩展方向
