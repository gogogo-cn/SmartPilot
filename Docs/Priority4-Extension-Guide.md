# 第四优先级扩展功能使用指南

## 📋 概述

本文档详细介绍 SmartPilot 插件的两个创新探索功能：

1. **Python 脚本执行引擎** - 利用 Python 生态进行复杂数学建模和机器学习
2. **插件生态系统** - 第三方开发者可创建、分享、安装插件，构建开放生态

---

## 🐍 功能一：Python 脚本执行引擎

### 1.1 功能说明

**作用**：允许 Skill 使用 Python 脚本处理数据，充分利用 Python 丰富的科学计算库（pandas, numpy, scipy, scikit-learn 等）。

**核心价值**：
- 弥补 VBA 在数学计算和数据分析方面的不足
- 接入 Python 丰富的机器学习生态
- 支持本地执行（隐私保护）和远程API调用（安全隔离）

### 1.2 架构设计

```
用户触发 PythonCode 模式的 Skill
    ↓
CustomSkillExecutor.ExecutePythonSkill()
    ↓
PythonExecutionService.ExecuteScriptAsync(scriptCode, inputData)
    ↓
┌─────────────────────────────────────┐
│         执行模式选择                 │
│                                     │
│  ┌─ 本地模式 (local) ─────────────┐ │
│  │ 1. 创建临时目录                  │ │
│  │ 2. 写入 input.json             │ │
│  │ 3. 包装脚本 (注入 I/O)          │ │
│  │ 4. subprocess 启动 python       │ │
│  │ 5. 捕获 stdout/stderr           │ │
│  │ 6. 读取 output.json            │ │
│  │ 7. 清理临时文件                 │ │
│  └─────────────────────────────────┘ │
│                                     │
│  ┌─ 远程模式 (remote) ────────────┐ │
│  │ 1. 打包 {script, data}          │ │
│  │ 2. POST 到远程端点              │ │
│  │ 3. 沙箱中执行 Python            │ │
│  │ 4. 返回 {stdout, stderr, data}  │ │
│  └─────────────────────────────────┘ │
│                                     │
└─────────────────────────────────────┘
    ↓
返回 PythonResult (Success/OutputData/Error)
    ↓
结果写入 Excel 或展示给用户
```

### 1.3 配置参数

在 `Config` 中配置 Python 执行环境：

```json
{
  "Python": {
    "ExecutionPath": "python",                    // 本地 Python 可执行文件路径
    "ExecutionMode": "local",                      // local | remote
    "RemoteEndpoint": "https://your-server/api/python", // 远程 API 端点
    "TimeoutSeconds": 60                           // 执行超时时间（秒）
  }
}
```

### 1.4 Skill 定义示例

#### 示例 A：高级数据分析

```json
{
  "name": "AdvancedDataAnalysis",
  "displayName": "高级数据分析(Python)",
  "processingMode": "PythonCode",
  "templateCode": "# 导入必要的库\nimport pandas as pd\nimport numpy as np\nfrom scipy import stats\n\n# 读取输入数据\ndf = pd.DataFrame(input_data.get('data', []))\n\n# 执行分析\nresult = {\n    'summary': df.describe().to_dict(),\n    'correlation': df.corr().to_dict(),\n    'row_count': len(df)\n}\n\noutput = result"
}
```

**使用流程**：
1. 用户选中 Excel 数据区域
2. 调用 AdvancedDataAnalysis Skill
3. Excel 数据自动转换为 `input_data` 变量传入 Python
4. Python 执行分析并输出 JSON 结果
5. 结果展示或写入 Excel

---

#### 示例 B：ML 预测引擎

```json
{
  "name": "MLPredictionEngine",
  "displayName": "ML预测引擎(Python+AI)",
  "processingMode": "PythonCode",
  "outputFormat": "chart",
  "templateCode": "# ML 预测代码\nfrom sklearn.ensemble import RandomForestRegressor\nfrom sklearn.model_selection import train_test_split\nfrom sklearn.metrics import mean_squared_error\n\n# 加载数据\ndf = pd.DataFrame(input_data.get('data', []))\nX = df.drop(columns=['target'])\ny = df['target']\n\n# 训练模型\nX_train, X_test, y_train, y_test = train_test_split(X, y)\nmodel = RandomForestRegressor(n_estimators=100)\nmodel.fit(X_train, y_train)\npredictions = model.predict(X_test)\n\n# 输出结果\noutput = {\n    'predictions': predictions.tolist(),\n    'feature_importance': dict(zip(X.columns, model.feature_importances_.tolist())),\n    'rmse': np.sqrt(mean_squared_error(y_test, predictions))\n}"
}
```

**能力**：
- 自动训练机器学习模型
- 特征重要性分析
- 预测精度评估
- 可视化输出（图表格式）

---

### 1.5 Python 脚本编写规范

#### 输入数据格式

SmartPilot 会自动将 Excel 数据转换为以下格式传入：

```python
# input_data 是一个字典，包含:
input_data = {
    'data': [                          # Excel 数据行列表
        {'列A': '值1', '列B': 123},   # 每行是一个字典
        {'列A': '值2', '列B': 456}
    ],
    'selectedRange': 'A1:B10',          # 选区地址
    'sheetName': 'Sheet1',              # 工作表名
    'targetColumn': 'B',                # 用户指定的目标列
    'featureColumns': ['A', 'B']        # 特征列
}
```

#### 输出数据格式

```python
# 必须设置 output 变量作为输出
output = {
    'data': [...],      # 要写回 Excel 的数据
    'message': '...',    # 说明信息
    'charts': [...]     # 图表定义（可选）
}

# 或者直接使用简单值
output = 42
output = "文本结果"
output = [1, 2, 3]
```

#### 可用的 Python 库

| 库名 | 用途 | 是否预装 |
|------|------|----------|
| pandas | 数据处理 | ✅ 推荐 |
| numpy | 数值计算 | ✅ 推荐 |
| scipy | 科学计算 | ✅ 推荐 |
| matplotlib | 绑图 | ✅ 可选 |
| scikit-learn | 机器学习 | ⚠️ 需安装 |
| statsmodels | 统计模型 | ⚠️ 需安装 |

#### 安全限制

- 执行时间限制：默认 60 秒（可配置）
- 文件系统访问：仅限临时目录
- 网络访问：默认禁止（需权限）
- 内存限制：受系统资源约束

### 1.6 两种执行模式对比

| 维度 | 本地模式 (Local) | 远程模式 (Remote) |
|------|------------------|------------------|
| **速度** | 快（无网络延迟） | 取决于网络 |
| **安全性** | 数据不出本机 | 数据发送到服务器 |
| **依赖管理** | 需本地安装库 | 服务端统一管理 |
| **适用场景** | 敏感数据、离线环境 | 团队共享、复杂依赖 |
| **配置要求** | 安装 Python | 配置远程端点 URL |
| **调试便利性** | 可查看完整日志 | 仅返回结果 |

### 1.7 环境检查

```vb
' 检查 Python 环境是否可用
Dim envInfo = Await PythonExecutionService.Instance.CheckEnvironmentAsync()

' 返回:
' {
'   "localAvailable": true/false,
'   "pythonVersion": "Python 3.11.0",
'   "executionMode": "local",
'   "remoteAvailable": true/false,
'   "ready": true/false
' }
```

---

## 🔌 功能二：插件生态系统

### 2.1 功能说明

**作用**：建立开放的插件体系，允许第三方开发者扩展 SmartPilot 的能力，形成丰富的应用市场。

**核心组件**：

| 组件 | 文件 | 作用 |
|------|------|------|
| **PluginSystem** | Plugins/PluginSystem.vb | 插件生命周期管理器 |
| **PluginSDK** | Plugins/PluginSDK.vb | 开发者接口定义 |
| **PluginMarketplaceBridge** | Bridges/PluginMarketplaceBridge.vb | 市场API桥梁 |

### 2.2 插件架构

```
┌──────────────────────────────────────────────────────┐
│                SmartPilot 宿主 (Host)                 │
│                                                      │
│  ┌────────────────────────────────────────────────┐  │
│  │              PluginSystem (核心)               │  │
│  │  • 加载/卸载                                   │  │
│  │  • 启用/禁用                                   │  │
│  │  • 沙箱执行                                    │  │
│  │  • 权限控制                                    │  │
│  └────────────────────────────────────────────────┘  │
│                                                      │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐          │
│  │ Plugin A │  │ Plugin B │  │ Plugin C │          │
│  │ (财务)   │  │ (HR)     │  │ (数据)   │          │
│  └──────────┘  └──────────┘  └──────────┘          │
│                                                      │
│  ┌────────────────────────────────────────────────┐  │
│  │         PluginMarketplaceBridge (市场)          │  │
│  │  • 浏览可用插件                                │  │
│  │  • 安装/卸载                                  │  │
│  │  • 版本更新                                    │  │
│  │  • 开发工具链                                  │  │
│  └────────────────────────────────────────────────┘  │
│                                                      │
└──────────────────────────────────────────────────────┘
```

### 2.3 插件开发指南

#### Step 1: 创建插件项目结构

```
MyFinancePlugin/
├── manifest.json          # 插件清单（必需）
├── src/
│   ├── PluginMain.cs      # C# 入口类（二选一）
│   └── PluginMain.vb      # VB.NET 入口类（二选一）
├── MyFinancePlugin.dll    # 编译产物
└── README.md              # 说明文档
```

#### Step 2: 编写 manifest.json

```json
{
  "id": "com.mycompany.finance",
  "name": "财务分析增强包",
  "version": "1.0.0",
  "description": "提供专业的财务分析功能：汇率转换、税务计算、报表生成",
  "author": "Your Name",
  "minPluginHostVersion": "2.0.0",
  "entryClass": "MyFinancePlugin.PluginMain",
  "category": "Finance",
  "tags": ["财务", "金融", "报表"],
  "permissions": {
    "allowFileRead": true,
    "allowFileWrite": true,
    "allowNetworkAccess": true,
    "allowMcpTools": false,
    "maxExecutionTimeMs": 30000,
    "allowedDomains": ["api.exchangerate-api.com"]
  }
}
```

#### Step 3: 实现 IPlugin 接口

**C# 示例**：

```csharp
using System.Collections.Generic;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using ExcelAIVSTOAddIn.Plugins;

namespace MyFinancePlugin
{
    public class PluginMain : IPlugin
    {
        private PluginInfo _info;

        public async Task<bool> OnInitializeAsync()
        {
            _info = new PluginInfo
            {
                Id = "com.mycompany.finance",
                Name = "财务分析增强包",
                Version = "1.0.0"
            };
            
            // 初始化资源...
            return true;
        }

        public async Task OnEnableAsync() 
        { 
            // 启动后台服务...
        }

        public async Task OnDisableAsync() 
        { 
            // 清理资源...
        }

        public async Task<JObject> ExecuteActionAsync(
            string actionName, 
            Dictionary<string, object> inputs)
        {
            var result = new JObject();
            
            switch (actionName)
            {
                case "convertCurrency":
                    result = await ConvertCurrencyAsync(inputs);
                    break;
                    
                case "calculateTax":
                    result = await CalculateTaxAsync(inputs);
                    break;
                    
                default:
                    result["error"] = $"未知动作: {actionName}";
                    break;
            }
            
            return result;
        }

        public List<SkillDefinition> GetProvidedSkills()
        {
            // 返回此插件提供的自定义 Skill
            return new List<SkillDefinition>
            {
                new SkillDefinition 
                { 
                    Name = "CurrencyConverter",
                    DisplayName = "汇率转换器"
                },
                new SkillDefinition 
                { 
                    Name = "TaxCalculator",
                    DisplayName = "个税计算器"
                }
            };
        }

        public PluginInfo GetPluginInfo() => _info;
        
        private async Task<JObject> ConvertCurrencyAsync(Dictionary<string, object> inputs)
        {
            // 实现汇率转换逻辑...
            var result = new JObject();
            result["convertedAmount"] = 1234.56m;
            result["rate"] = 7.24m;
            return result;
        }
        
        private async Task<JObject> CalculateTaxAsync(Dictionary<string, object> inputs)
        {
            // 实现税务计算逻辑...
            var result = new JObject();
            result["taxAmount"] = 500m;
            result["afterTaxIncome"] = 9500m;
            return result;
        }
    }
}
```

**VB.NET 示例**：

```vb
Imports System.Collections.Generic
Imports System.Threading.Tasks
Imports Newtonsoft.Json.Linq
Imports ExcelAIVSTOAddIn.Plugins

Namespace MyFinancePlugin

    Public Class PluginMain
        Implements IPlugin

        Private _info As PluginInfo

        Public Async Function OnInitializeAsync() As Task(Of Boolean) _
            Implements IPlugin.OnInitializeAsync
            
            _info = New PluginInfo With {
                .Id = "com.mycompany.finance",
                .Name = "财务分析增强包",
                .Version = "1.0.0"
            }
            Return True
        End Function

        Public Async Function OnEnableAsync() As Task Implements IPlugin.OnEnableAsync
        End Function

        Public Async Function OnDisableAsync() As Task Implements IPlugin.OnDisableAsync
        End Function

        Public Async Function ExecuteActionAsync(actionName As String,
                                                inputs As Dictionary(Of String, Object)) _
                                            As Task(Of JObject) Implements IPlugin.ExecuteActionAsync
            Dim result As New JObject()
            
            Select Case actionName
                Case "convertCurrency"
                    result("convertedAmount") = 1234.56D
                    result("rate") = 7.24D
                    
                Case "calculateTax"
                    result("taxAmount") = 500D
                    result("afterTaxIncome") = 9500D
                    
                Case Else
                    result("error") = "未知动作: " & actionName
            End Select
            
            Return result
        End Function

        Public Function GetProvidedSkills() As List(Of SkillDefinition) _
            Implements IPlugin.GetProvidedSkills
            Return New List(Of SkillDefinition)()
        End Function

        Public Function GetPluginInfo() As PluginInfo Implements IPlugin.GetPluginInfo
            Return _info
        End Function

    End Class

End Namespace
```

#### Step 4: 编译和部署

**编译要求**：
- 目标框架：.NET Framework 4.8
- 引用程序集：
  - `ExcelAIVSTOAddIn.dll` （宿主主程序集）
  - `NewtonSoft.Json.dll` （JSON 处理）

**部署步骤**：
1. 编译生成 DLL
2. 将整个插件文件夹复制到 SmartPilot 的 `Plugins` 目录
3. 重启 SmartPilot 或刷新插件列表
4. 在插件管理界面启用新插件

### 2.4 插件沙箱机制

**为什么需要沙箱？**
- 隔离第三方代码，防止影响宿主稳定性
- 权限最小化原则，降低安全风险
- 资源限制，防止单个插件占用过多资源

**沙箱能力**：

| 能力 | 默认权限 | 说明 |
|------|---------|------|
| 读取文件 | ✅ 允许 | 读取必要的数据文件 |
| 写入文件 | ❌ 禁止 | 防止恶意修改 |
| 网络访问 | ❌ 禁止 | 防止数据外泄 |
| MCP 工具 | ❌ 禁止 | 高危操作需显式授权 |
| Excel 写入 | ✅ 允许 | 核心功能需要 |
| 执行时间 | 30秒上限 | 防止死循环 |

**权限申请流程**：

```
开发者声明所需权限 (manifest.json)
    ↓
用户安装时审查权限
    ↓
PluginSystem 校验白名单
    ↓
运行时动态检查每次操作
    ↓
违规操作 → 抛出 SecurityException
```

### 2.5 插件市场 API

#### 浏览插件

```javascript
// 获取所有可用插件
const result = await pluginMarketplaceBridge.getAvailablePluginsAsync();
// 返回: { plugins: [...], count: 15, categories: [...] }

// 分类筛选
const financePlugins = await pluginMarketplaceBridge.getAvailablePluginsAsync("Finance");

// 搜索
const searchResult = await pluginMarketplaceBridge.getAvailablePluginsAsync("", "股票");
```

#### 安装/卸载

```javascript
// 从本地文件安装
await pluginMarketplace.installPluginAsync("C:\\Downloads\\MyPlugin.spplugin");

// 从URL安装
await pluginMarketplace.installPluginAsync("https://market.smartpilot.cloud/plugins/finance.zip");

// 卸载
await pluginMarketplace.uninstallPluginAsync("com.example.finance");

// 启用/禁用
await pluginMarketplace.togglePluginAsync("com.example.finance", true);  // 启用
await pluginMarketplace.togglePluginAsync("com.example.finance", false); // 禁用
```

#### 开发工具

```javascript
// 创建插件模板
const template = await pluginMarketplace.createPluginTemplateAsync(JSON.stringify({
    id: "com.mycompany.myplugin",
    name: "我的插件",
    author: "Developer"
}));
// 返回: { pluginDir: "...", nextSteps: [...] }

// 验证插件包
const validation = await pluginMarketplace.validatePluginPackageAsync("C:\\MyPlugin\\");
// 返回: { isValid: true, validationResults: [...] }
```

### 2.6 插件状态机

```
                    ┌─────────────┐
                    │  Unloaded   │ ← 未加载
                    └──────┬──────┘
                           │ Initialize()
                           ▼
                    ┌─────────────┐
                    │   Loaded    │ ← 已加载（未激活）
                    └──────┬──────┘
                           │ Enable()
                           ▼
                    ┌─────────────┐
           ┌────────►│   Active    │ ← 运行中
           │         └──────┬──────┘
           │ Disable()       │ ExecuteAction()
           │                 ▼
           │         ┌─────────────┐
           │         │  Disabled   │ ← 已禁用
           │         └──────┬──────┘
           │    Enable()       │
           │                 ▼
           └─────────────────┘
                           
                    Error (异常状态)
```

---

## 🚀 快速开始

### 场景 1：使用 Python 进行数据分析

1. 确保 Python 已安装（推荐 3.8+）且 pandas/numpy 可用
2. 准备 Excel 数据表
3. 选中数据区域
4. 调用 **AdvancedDataAnalysis** Skill
5. 查看 Python 输出的统计分析结果

### 场景 2：开发你的第一个插件

1. 调用 **PluginManager** Skill
2. 选择"创建新插件模板"
3. 填写插件基本信息
4. 编辑生成的 `src/PluginMain.cs` 或 `.vb`
5. 编译为 DLL
6. 将文件夹复制到 `Plugins` 目录
7. 重启 SmartPilot 并在插件管理界面查看

### 场景 3：从市场安装插件

1. 调用 **PluginManager** Skill
2. 选择"浏览插件市场"
3. 搜索或分类查找感兴趣的插件
4. 点击"安装"
5. 审查权限请求
6. 确认安装并启用

---

## 💡 最佳实践

### Python 脚本最佳实践

1. **输入验证**
   ```python
   if not input_data or 'data' not in input_data:
       output = {'error': '缺少输入数据'}
   ```

2. **错误处理**
   ```python
   try:
       # 分析逻辑
   except Exception as e:
       output = {'error': str(e), 'traceback': traceback.format_exc()}
   ```

3. **性能优化**
   - 大数据量时使用 chunk 处理
   - 避免循环操作，优先向量化
   - 及时释放大对象内存

4. **输出标准化**
   ```python
   output = {
       'success': True,
       'data': result_df.to_dict('records'),
       'summary': {...},
       'message': '分析完成'
   }
   ```

### 插件开发最佳实践

1. **权限最小化**
   - 只申请必要的权限
   - 在文档中说明每项权限的用途

2. **版本兼容性**
   - 明确标注最低宿主版本
   - 处理 API 不存在时的降级方案

3. **错误处理**
   - 所有公共方法捕获异常
   - 提供有意义的错误消息
   - 不要让异常传播到宿主

4. **资源管理**
   - OnDisable 时释放所有资源
   - 使用 IDisposable 模式管理非托管资源
   - 避免内存泄漏

5. **日志记录**
   - 使用 Debug.WriteLine 输出诊断信息
   - 记录关键操作和耗时
   - 方便排查问题

---

## 📊 四大优先级完整能力矩阵

```
┌─────────────────────────────────────────────────────────────────────┐
│              SmartPilot 智航 - 企业级智能平台                        │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  第一优先级 (智能化)                                                │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │ ✓ RAG知识库联动     ✓ Pipeline管道编排                       │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  第二优先级 (体验升级)                                              │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │ ✓ 事件驱动触发     ✓ 图表/透视表/条件格式                     │   │
│  │ ✓ 三级权限审批                                               │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  第三优先级 (战略扩展)                                              │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │ ✓ 多模态AI(表格识别/文档理解)  ✓ 协作分享(导出/导入/同步)    │   │
│  │ ✓ 性能监控(仪表盘/健康检查)                                 │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  第四优先级 (创新探索)                                              │
│  ┌─────────────────────────────────────────────────────────────┐   │
│  │ ✓ Python执行引擎(本地/远程)  ✓ 插件生态(SDK/沙箱/市场)      │   │
│  └─────────────────────────────────────────────────────────────┘   │
│                                                                     │
│  总计: 17个示例Skill | 5个新建Bridge | 4份详细文档                │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

---

## 🎉 总结

第四优先级的实施将 SmartPilot 提升为**完全开放的智能平台**：

| 维度 | 实施前 | 实施后（四大优先级全部完成）|
|------|--------|---------------------------|
| **编程语言支持** | VBA + LLM | ✅ + Python（完整生态）|
| **扩展方式** | 仅内置 Skill | ✅ + 第三方插件系统 |
| **数据处理** | 基础计算 | ✅ + pandas/numpy/scipy/sklearn |
| **AI 能力** | 对话式 AI | ✅ + 机器学习模型训练/推理 |
| **生态开放性** | 封闭系统 | ✅ + 开放 SDK + 插件市场 |
| **社区参与** | 无 | ✅ + 开发者可贡献插件 |
| **应用场景** | 办公辅助 | ✅ + 科学计算/数据分析/AI建模 |

**当前状态**：所有四大优先级已全部实施完成！🎊

---

## 🔗 相关文档

- [第一优先级扩展指南](./Priority1-Extension-Guide.md) - RAG 和 Pipeline
- [第二优先级扩展指南](./Priority2-Extension-Guide.md) - 事件驱动/输出格式/权限审批
- [第三优先级扩展指南](./Priority3-Extension-Guide.md) - 多模态/协作/监控
- [Skill 系统架构](./SkillCreator.md) - Skill 定义和元编程
- [Skill 扩展路线图](../.trae/documents/skill-extension-roadmap.md) - 全部扩展方向总览

---

## 📁 关键文件索引（第四优先级新增）

| 文件 | 类型 | 行数 | 说明 |
|------|------|------|------|
| [PythonExecutionService.vb](ExcelAIVSTOAddIn/Services/PythonExecutionService.vb) | 已有 | ~344行 | Python执行引擎（本地+远程）|
| [PluginSystem.vb](ExcelAIVSTOAddIn/Plugins/PluginSystem.vb) | 新建 | ~450行 | 插件生命周期管理器 |
| [PluginSDK.vb](ExcelAIVSTOAddIn/Plugins/PluginSDK.vb) | 新建 | ~120行 | 插件开发接口定义 |
| [PluginMarketplaceBridge.vb](ExcelAIVSTOAddIn/Bridges/PluginMarketplaceBridge.vb) | 新建 | ~550行 | 插件市场API桥梁 |
| [DefaultSkills.json](ExcelAIVSTOAddIn/Config/DefaultSkills.json) | 更新 | ~684行 | +3个新示例Skill |
| [Priority4-Extension-Guide.md](ExcelAIVSTOAddIn/Docs/Priority4-Extension-Guide.md) | 新建 | ~700行 | 完整使用文档 |
