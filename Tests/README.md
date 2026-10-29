# Tests — 自动化测试说明

ExcelAI VSTO Add-In 的自动化测试。前端 / 服务端 / 路由 / 模板逻辑使用 Node 内置 `node:test`（**零依赖**，不引入 node_modules / jest）；VB CoreUtils 逻辑使用独立 .NET 测试工程。

## 一键运行（推荐）

```powershell
Tools\run-tests.ps1
```

与 CI `test` job 完全一致，分两步：

| 步骤 | 命令 | 产物 |
|---|---|---|
| 1/2 JS 套件 | `node Tests/run-tests.js` | 聚合运行所有 `*.test.js`（进程内，零 spawn） |
| 2/2 VB Core | `dotnet run --project Tests\VbTests\ExcelMcpCoreUtils.Tests.vbproj -c Release` | `ExcelMcpCoreUtils` 纯逻辑断言 |

任一失败即非 0 退出。结尾输出 `RESULT: ALL TESTS PASSED` 或 `RESULT: FAIL`。

## 运行方式

### 全部 JS 套件

```powershell
cd Tests
node run-tests.js        # 聚合运行器（进程内加载所有套件）
```

或经 npm：

```powershell
cd Tests
npm test                 # 等效 node run-tests.js
```

### 单个套件 / 分组

```powershell
node Tests/agent-loop.test.js            # 单独跑某套件
npm run test:router                      # 按分组跑（见 package.json scripts）
npm run test:e2e-script                  # 脚本式套件（skill-chart-clean-e2e，单独运行）
```

> **为什么用 `node run-tests.js` 而不是 `node --test Tests/`？**
> node:test 的 `--test` 模式会 spawn 子进程，在部分沙箱 / CI 环境可能受限。
> `run-tests.js` 在**单进程内顺序加载**每个套件，规避该问题。

### VB Core 测试

```powershell
dotnet run --project Tests\VbTests\ExcelMcpCoreUtils.Tests.vbproj -c Release
```

> 该测试工程直接编译并测试**生产代码**（单一事实源）：`ExcelMcpCoreUtils.vb`
> 与 `PptDesignEngine.vb`。其中 `PptDesignEngine`（数据感知 PPT 设计引擎的
> 纯决策链：列分类 / 趋势 / ID 识别 / 版式方向 / 图表类型 / 洞察文案 / 需求偏置
> / column hints 覆盖）由 `Tests/VbTests/PptDesignEngineTests.vb` 覆盖行为断言，
> 输出依赖的 `PathHelper` 由 `Tests/VbTests/Utils/PathHelperStub.vb` 最小桩满足。
> 注意：`PptDesignEngine` 整条决策链（`AnalyzeData` → `BuildDesign`/`DesignFromData`）
> 是确定性规则引擎，不触碰 Excel COM，可离线单测。

## harness 原理

`Tests/helpers/harness.js` 在 Node 中**模拟浏览器环境**（`window` / `document` / `App` / `WebView2`），
按与 `index.html` 一致的加载顺序，把 `UI/js/chat-v2/` 下的真实模块按源文件加载进同一个 `App.ChatV2`
命名空间，供 `node:test` 直接驱动。

```js
// 测试文件开头
const { loadChatV2 } = require('./helpers/harness');
const ctx = loadChatV2();            // 返回 { App, ... }
const engine = ctx.App.ChatV2.PythonTemplateValidator;
```

- **共享源接缝（VB）**：`excel_*` 等工具的 COM-free 纯逻辑已抽取到
  `Services/Mcp/ExcelMcpCoreUtils.vb`，被插件工程与 `Tests/VbTests` 用
  `<Compile Link=...>` 编译同一源文件，实现"同一份逻辑、两端验证"。
- **e2e 套件**：`*-e2e.test.js` 加载更完整模块图（含 regex / 序列化 / 渲染），
  用于验证路由到执行的端到端链路。

## 测试矩阵

| 套件文件 | 覆盖 |
|---|---|
| `agent-loop.test.js` / `skill-agent-e2e.test.js` | 智能体循环、原生 tool_calls 契约、工具链调度 |
| `tool-registry.test.js` / `tools-consistency.test.js` | 工具注册、显示名无硬编码、JS=JSON=VB 元数据一致性 |
| `session-store.test.js` | 会话增量 / 持久化 |
| `pipeline-router.test.js` / `routing-constants.test.js` | 路由常量、pipeline 分配 |
| `skill-route.test.js` / `skill-routing-quickprompts.test.js` | skill 意图 → 后端工具路由、快捷提示词路由 |
| `trajectory-replay.test.js` | 轨迹回放 |
| `python-data-inject.test.js` / `python-pipeline-e2e.test.js` | Python 数据注入、模板执行链路 |
| `python-template-validator.test.js` | Python 模板库注册、参数校验、脚本生成、JSON 变体容错 |
| `skill-spec-engine.test.js` / `skill-spec-e2e.test.js` | 规范引擎、spec 驱动执行 |
| `skill-formatter.test.js` / `skill-chart-clean.test.js` | 结果格式化、清洗/图表 spec |
| `skill-chart-clean-e2e.test.js` | 脚本式 e2e（单独运行） |
| `remove-duplicates.test.js` | 去重工具多列 / columns 类型契约 |
| `vba-validator.test.js` | VBA 模板安全空壳静态校验 |
| `steps-plan-outline.test.js` | 步骤路线图顶部概览（P1-6，整句→短标题） |
| `test-suite-registration.test.js` | 护栏：套件自动发现完整、无漏跑、进程隔离与 harness 全局身份稳定 |
| `host-tools-config.test.js` | 宿主工具配置（Word/PPT）元数据 + 全仓关键 JSON 可解析（防静默加载为空） |
| `mobile-bridge-auth.test.js` | 护栏：本地遥控服务 9528 的 CORS 白名单与 sessionToken 不外泄 |
| `VbTests/` | `ExcelMcpCoreUtils`：`BuildChunkName` / `ColumnLetterToNumber` / 数字与列字母互转 |

## 新增测试套件的注意点

1. 套件文件以 `.test.js` 命名，文件顶层调用 `test()`。
2. **无需手工注册**：`run-tests.js` 自动扫描 `Tests/*.test.js`（2026-10-09 起）。
   此前手工维护 `SUITES` 数组曾导致 12 个护栏套件（含
   `host-export-pdf-dispatch`、`mcp-dispatcher-format-guard`、
   `skill-routing-domain-guard`）**从未被执行**，已修复的 bug 悄悄回归。
   护栏见 `Tests/test-suite-registration.test.js`。
3. 依赖 `App.ChatV2` 的套件通过 `require('./helpers/harness')` 的 `loadChatV2()` 取环境。
4. 脚本式套件（自行 `process.exit`，如 `skill-chart-clean-e2e`、`skill-crud-e2e`）
   必须在 `run-tests.js` 的 `EXCLUDES` 中**写明原因**，并在 package.json 用独立
   script（`test:e2e-script` / `test:crud`）运行；未登记但疑似脚本式的会被自动
   拦截并告警（不会静默漏跑）。
5. **每个套件在独立子进程中运行**（2026-10-09 起）。此前所有套件 `require` 进同一
   进程，共享 `window` / `App` 全局互相污染，表现为"单跑全绿、聚合必红"。
6. **子进程以 `node --test <file>` 启动**（2026-10-16 起）。此前直接 `node <file>`：
   node:test 的 TAP 汇总要等事件循环排空才打印，而部分套件会残留 ref'd 定时器
   （StepExecutor 的 60s AGENT_TIMEOUT、dashboard-html-renderer 的 1s setInterval 等），
   于是套件自带的兜底 `setTimeout(process.exit(0), 3~5s)` 先把进程杀掉 →
   汇总永远打不出来（6 个套件长期显示"用例全 ok 但未打印汇总"）。
   `--test` 由 runner 接管生命周期，用例跑完立即出汇总并带 exit code 退出。
7. **危险工具在所有思考模式（含 expert）都要用户确认**（见 `tool-registry.js` 守卫
   `dangerous-tools`）。凡是要断言"工具真的执行了"的用例，必须先用 harness 导出的
   `stubDangerConfirm(true)` 模拟用户点「仍要继续执行」，否则会被守卫拦成 `ok=false`。

## 相关工具

- `node Tools/lint.js` — 零依赖前端静态检查（递归语法 + 规则子集），CI `lint` job 硬闸门。
- `Tools/check-tools-consistency.ps1` — 构建前校验 MCP 工具元数据一致性。
- `node Tools/sync-ui-to-debug.js` — 运行 `npm`/`node` 之外的 UI 资源增量同步（见根 README）。