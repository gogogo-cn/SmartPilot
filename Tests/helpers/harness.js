/**
 * ============================================================
 * Tests/helpers/harness.js — chat-v2 前端模块测试环境
 * ============================================================
 * 用途：在 Node 中模拟浏览器环境（window/document/App/WebView2），
 *       按真实加载顺序加载 chat-v2 模块，供 node:test 测试使用。
 * 用法：
 *   const { loadChatV2, App } = require('./helpers/harness');
 * ============================================================
 */
'use strict';
const fs = require('fs');
const path = require('path');

/** chat-v2 模块目录 */
const BASE = path.join(__dirname, '..', '..', 'UI', 'js', 'chat-v2');

/**
 * 与 index.html 对齐的加载顺序（测试用精简集合）。
 *
 * ⚠️ 历史坑（2026-10-09）：这里曾漏掉 config/skill-domain.js —— 技能类型判定
 * 注册表。缺了它 detectSkillType() 返回 null → SKILL_SPEC_CONTRACT[null]
 * 静默回退 beautifier → SmartChartWizard 被当成"表格美化"渲染 VBA。
 * 线上 index.html 会加载它，所以**功能其实没坏，是测试在骗人**。已补上。
 *
 * ⚠️ 不要盲目把 index.html 的 65 个模块全搬进来：实测补到 65 个后，
 * pipeline-router / skill-route 这类"套件内重复 loadChatV2 9 次"的套件会
 * 触发 V8 OOM（exit=134）。若需新增模块，请逐个加并跑全量确认不 OOM。
 */
const DEFAULT_ORDER = [
    'services/session-store.js',
    // ★ 选区范围判断与降级阶梯的单一事实源（agent-loop 无选区兜底 / skill-spec-build 选区优先 /
    //   mcp-dispatcher 批量美化链共用）。必须早于 agent-loop.js 与 services/mcp-dispatcher.js。
    'services/range-resolver.js',
    'services/compaction.js',
    'services/trajectory-replay.js',
    'services/stream-processor.js',
    'services/llm-stream-receiver.js',
    'services/content-extractor.js',
    'services/llm-client.js',
    'services/agent-events.js',
    'agent-loop/chart-gate.js',
    'agent-loop/zero-tool-guards.js',
    'agent-loop/file-delivery-gate.js',
    'agent-loop/tool-execution.js',
    'agent-loop.js',
    'services/tool-registry.util.js',
    'services/tool-envelope.js',
    'services/turn-goal.js',
    'services/tool-registry.js',
    // ★ 必须加载：技能类型判定注册表。缺了它 detectSkillType() 返回 null →
    //   SKILL_SPEC_CONTRACT[null] 静默回退 beautifier → SmartChartWizard 被当成"表格美化"
    //   渲染 VBA（真实踩坑，见 Tests/chart-empty-source.test.js 的历史注释）。
    'config/skill-domain.js',
    'services/pipeline-registry.js',
    'services/prompt-builder.js',
    'services/vba-template-engine.js',
    'services/skill-spec-engine.js',
    'services/intent-config.js',
    'services/intent-detector.js',
    'services/task-planner.js',
    'services/render-html.js',
    'services/task-execution.js',
    'services/script-generator.js',
    'services/json-utils.js',
    'services/llm-io.js',
    'pipelines/text-pipeline.js',
    'pipelines/python-pipeline.js',
    'pipelines/html-display-pipeline.js',
    'pipelines/image-pipeline.js',
    'pipelines/skill/skill-shared-util.js',
    'pipelines/doc-pipeline.js',
    'pipelines/skill/skill-constants.js',
    'pipelines/skill/skill-clean-crud.js',
    'pipelines/skill/skill-spec-build.js',
    'pipelines/skill/skill-steps-panel.js',
    'pipelines/skill/skill-spec-execute.js',
    'pipelines/skill/skill-steps-execute.js',
    'pipelines/skill-pipeline.js',
    'pipelines/code-pipeline.js',
    'config/routing-constants.js',
    'chat-orchestrator.js',
    'ui/renderers/steps-plan.js',
    'ui/renderers/code-block.js',
    'ui/renderers/json.js',
    'ui/renderers/tool-card.js',
    // 说明：index.html 还加载 syntax-highlight / markdown / inline-chart / suggestions
    // 四个渲染器，但它们会在每次 loadChatV2() 重建大型正则/词法表，套件内多次加载
    // （pipeline-router 9 次）会触发 V8 OOM（实测 exit=134）。这四个是纯展示层，
    // 不参与 Agent 决策，故测试环境不加载。
    'ui/chat-renderer.js',
    'chat-entry.js',
    'index.js'
];

/** DOM 最小 stub（chat-renderer/chat-entry 需要） */
function createDomStub() {
    return {
        readyState: 'complete',
        addEventListener() {}, removeEventListener() {},
        querySelector() { return null; }, querySelectorAll() { return []; },
        getElementById() { return null; },
        createElement() { return {
            style: {}, classList: { add() {}, remove() {}, contains() { return false; } },
            appendChild() {}, addEventListener() {}, setAttribute() {}, getAttribute() { return null; }
        }; },
        body: { appendChild() {}, style: {} }, documentElement: { style: {} }, head: { appendChild() {} },
        title: '', cookie: '', location: { href: 'about:blank', search: '', hash: '' },
        visibilityState: 'visible', hidden: false
    };
}

/** 进程内共享的全局环境（见 loadChatV2 注释：对象身份必须稳定） */
let SHARED_WINDOW = null;
let SHARED_DOM = null;

/**
 * 创建 window / DOM stub 环境。
 * @returns {{ window: Object, domStub: Object }}
 */
function createWindowEnv() {
    const window = {};
    const domStub = createDomStub();
    window.document = domStub;
    window.navigator = { userAgent: 'node-test', platform: 'Win32' };
    window.location = domStub.location;
    window.ChatLog = { info() {}, warn() {}, error() {}, debug() {} };
    window.chrome = { webview: { hostObjects: { excelBridge: {} } } };
    window.App = { ChatV2: {}, config: { model: 'test-model' }, mcpTools: [], cachedSkills: [] };
    window.AbortController = AbortController;
    // session-store 的 saveToLegacyHistory 在异步定时器中引用裸标识符 AppHistory，
    // 若未声明会在测试结束后抛出 ReferenceError。提供与 app-history.js 一致的 stub。
    window.AppHistory = { addToHistory() {}, loadHistory() {}, renderHistory() {} };
    // 挂载到 App 便于测试直接访问（贴近 WebView2 环境的 window.chrome）
    window.App.window = window;
    window.App.chrome = window.chrome;
    return { window, domStub };
}

/**
 * 加载 chat-v2 模块（每次调用重置模块注册，但**保持 window/App 对象身份不变**）。
 *
 * ⚠️ 为什么必须共享同一个 window 对象（历史 Bug · 2026-10-09）：
 * 早期实现每次调用都 `const window = {}` 新建环境。聚合运行（run-tests.js）时
 * 所有套件先被 require（模块级 loadChatV2 全部执行完）才跑 test()，于是：
 *   · 套件 A 拿到的 window 引用是「A 加载时」的那个对象；
 *   · 套件 B 一加载，global.window 被换成 B 的对象；
 *   · 被 eval 的模块是后期绑定全局标识符 `window` 的，运行时读的是 global.window（B 的），
 *   · 于是 A 在测试里 `window.SmartPilotHost = WORD` 写进了孤儿对象 → 模块看不到 → A 变红。
 * 表现为「单跑全绿、聚合必红」，曾导致 host-export-pdf-dispatch 等护栏套件长期
 * 被误判为不稳定。修复：window/App 复用同一对象，只重置模块注册内容。
 *
 * @param {string[]} [order] - 自定义加载顺序（默认 DEFAULT_ORDER）
 * @returns {{ window: Object, App: Object, ChatLog: Object }}
 */
function loadChatV2(order) {
    if (!SHARED_WINDOW) {
        const env = createWindowEnv();
        SHARED_WINDOW = env.window;
        SHARED_DOM = env.domStub;
    }
    const window = SHARED_WINDOW;
    const domStub = SHARED_DOM;

    // 重置「模块注册内容」而非对象本身：保持身份稳定，同时给每个套件干净命名空间
    window.App.ChatV2 = {};
    window.App.mcpTools = [];
    window.App.cachedSkills = [];
    window.App.config = { model: 'test-model' };

    global.window = window;
    global.document = domStub;
    global.location = domStub.location;
    // node 22 全局 navigator 是只读 getter，用 defineProperty 覆盖
    try {
        Object.defineProperty(global, 'navigator', { value: window.navigator, configurable: true, writable: true });
    } catch (e) { /* 忽略（某些 node 版本允许直接赋值） */ }
    globalThis.App = window.App;
    globalThis.AppHistory = window.AppHistory;

    // 间接 eval 的模块拿不到 Node 的 require；把 excel_tools.json 挂到 globalThis，
    // 供 filterToolsForMessage / resolveDangerInfo 在测试里读 intentKeywords / 危险元数据。
    try {
        const toolsPath = path.join(__dirname, '..', '..', 'Config', 'excel_tools.json');
        if (fs.existsSync(toolsPath)) {
            globalThis.__SMARTPILOT_EXCEL_TOOLS__ = JSON.parse(
                fs.readFileSync(toolsPath, 'utf8').replace(/^\uFEFF/, '')
            );
            window.__SMARTPILOT_EXCEL_TOOLS__ = globalThis.__SMARTPILOT_EXCEL_TOOLS__;
        }
    } catch (e) { /* 缺配置时模块走空索引兜底 */ }

    const loadOrder = order || DEFAULT_ORDER;
    for (const rel of loadOrder) {
        const full = path.join(BASE, rel);
        if (!fs.existsSync(full)) throw new Error('MISSING module: ' + rel);
        // 在全局作用域执行 IIFE
        (0, eval)(fs.readFileSync(full, 'utf8'));
    }

    return { window, App: window.App, ChatLog: window.ChatLog };
}

// 测试结束强制退出：session-store 的 saveToLegacyHistory 定时器会挂起进程。
// 用 unref 的定时器做兜底（node:test runner 模式下由 runner 控制退出，不受影响）。
if (require.main === module) {
    setTimeout(() => process.exit(0), 4000).unref();
}

/**
 * 危险操作确认弹窗桩（模拟用户点击）。
 *
 * 背景（2026-10-16）：tool-registry 的【安全守卫】对危险工具（excel_execute_vba /
 * excel_run_python_script / excel_delete_sheet …）在**所有思考模式（含 expert）**都要求
 * 用户在弹窗上确认后才执行（见 tool-registry.js 守卫 dangerous-tools 的注释）。
 * 而 harness 的默认 DOM stub 无法渲染确认弹窗（createElement 返回的对象没有
 * querySelector）→ confirmDangerousAction 内部抛错被吞 → 等同"用户取消"。
 * 因此凡是要断言"工具真的被执行了"的用例，都必须显式模拟用户点「仍要继续执行」。
 *
 * @param {boolean} approve true=自动点「仍要继续执行」，false=自动点「取消」
 * @returns {Object} 已安装到 global.document 的 DOM stub
 */
function stubDangerConfirm(approve) {
    const fireSel = approve === false ? '[data-ac="cancel"]' : '[data-ac="ok"]';
    const makeBtn = (isFire) => ({
        onclick: null,
        set onclick(fn) { if (isFire && typeof fn === 'function') fn(); }
    });
    const el = {
        style: {}, classList: { add() {}, remove() {}, contains() { return false; } },
        appendChild() {}, addEventListener() {}, setAttribute() {},
        set innerHTML(v) { this.__html = v; }, get innerHTML() { return this.__html || ''; },
        querySelector(sel) { return makeBtn(sel === fireSel); },
        querySelectorAll() { return []; }
    };
    const doc = {
        createElement() { return el; },
        body: { appendChild() {} },
        addEventListener() {},
        removeEventListener() {}
    };
    global.document = doc;
    return doc;
}

module.exports = { loadChatV2, DEFAULT_ORDER, BASE, stubDangerConfirm };
