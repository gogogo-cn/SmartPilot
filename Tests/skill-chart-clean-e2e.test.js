'use strict';
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，但截断到 chat-orchestrator（不加载 chat-entry/index，避免入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

async function testChart() {
    const App = loadChatV2(ORDER).App;
    App.cachedSkills = [{ name: 'SmartChartWizard', displayName: '智能图表向导', keywords: ['画图', '图表', '柱状图', '饼图'], inputContext: [], pipeline: 'skill', processingMode: 'SkillSpec' }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({ name: 'skill.SmartChartWizard', displayName: '智能图表向导', description: '图表', inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: App.cachedSkills[0], async execute() { return { ok: true, result: 'x' }; } });
    let chartCalled = null;
    tr.register({ name: 'excel_create_chart', displayName: '创建图表', description: 'd', inputSchema: { type: 'object', properties: {} }, source: 'mcp', async execute(args) { chartCalled = args; return { ok: true, result: { chartName: 'Chart 1' } }; } });

    App.ChatV2.llmClient = { stream({ messages }) { return (async function* () { yield { type: 'model', model: 't' }; yield { type: 'token', delta: JSON.stringify({ chartType: '饼图', title: '销售占比' }) }; yield { type: 'done', finishReason: 'stop' }; })(); } };
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({ message: '画个饼图看占比', kind: 'text', context: { selection: { address: 'A1:B10', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] } });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    let failed = null;
    for await (const evt of gen) { events.push(evt.type); if (evt.type === 'error') failed = evt.error; }
    const ok = !failed && events.includes('pipeline_done');
    console.log(`图表: ${ok ? '✅' : '❌ ' + JSON.stringify(failed)} chartType=${chartCalled?.chartType} title=${chartCalled?.title}`);
    return ok && chartCalled?.chartType === 'xlPie';
}

async function testClean() {
    const App = loadChatV2(ORDER).App;
    App.cachedSkills = [{ name: 'SmartDataCleaner', displayName: '智能数据清洗大师', keywords: ['清洗', '去重', '空值'], inputContext: [], pipeline: 'skill', processingMode: 'SkillSpec' }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({ name: 'skill.SmartDataCleaner', displayName: '智能数据清洗大师', description: '清洗', inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: App.cachedSkills[0], async execute() { return { ok: true, result: 'x' }; } });
    let cleanCalled = null;
    let dedupCalled = null;
    // 清洗走 MCP 工具链（无 VBA）：executeCleanWithMcp 会检查去重/清洗工具是否注册
    // 【对齐生产】excel_read_range 在生产始终注册（预读校验用），此处注册桩覆盖真实链路
    tr.register({ name: 'excel_read_range', displayName: '读取区域', description: 'd', inputSchema: { type: 'object', properties: {} }, source: 'mcp', async execute() { return { ok: true, result: { data: [['名称', '数值'], ['a', 1], ['b', 2]] } }; } });
    // 【对齐现架构】去重走专用 excel_remove_duplicates（不再并入 excel_clean_data.operations）
    tr.register({ name: 'excel_remove_duplicates', displayName: '去重', description: 'd', inputSchema: { type: 'object', properties: {} }, source: 'mcp', async execute(args) { dedupCalled = args; return { ok: true, result: { success: true } }; } });
    tr.register({ name: 'excel_clean_data', displayName: '智能数据清洗', description: '清洗', inputSchema: { type: 'object', properties: {} }, source: 'mcp', async execute(args) { cleanCalled = args; return { ok: true, result: { success: true, message: '清洗完成', steps: ['清洗完成'] } }; } });

    App.ChatV2.llmClient = { stream({ messages }) { return (async function* () { yield { type: 'model', model: 't' }; yield { type: 'token', delta: JSON.stringify({ operations: ['removeDuplicates', 'fillEmpty'], fillMethod: 'zero' }) }; yield { type: 'done', finishReason: 'stop' }; })(); } };
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({ message: '清洗数据，去重并填充空值', kind: 'text', context: { selection: { address: 'A1:D10', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] } });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    let failed = null;
    for await (const evt of gen) { events.push(evt.type); if (evt.type === 'error') failed = evt.error; }
    const ok = !failed && events.includes('pipeline_done');
    const ops = (cleanCalled && cleanCalled.operations) || [];
    // 2026-10-21 行为修正：excel_remove_duplicates / excel_clean_data 均为危险工具
    // （excel_tools.json catalog，severity=medium），headless 环境 fail-safe 取消 →
    // spec 层优雅收尾（pipeline_done、无 error、清洗工具零执行），
    // 不再降级 fallbackToText 让 agent-loop 拿同一条消息二次执行。
    // 因此这里断言的是"守卫取消行为"而非"清洗链路跑通"（链路被守卫挡住是设计使然）。
    const hasDedup = !!dedupCalled;
    const hasFill = ops.includes('fillEmpty');
    const good = ok && !dedupCalled && !cleanCalled;
    console.log(`清洗: ${good ? '✅' : '❌ ' + JSON.stringify(failed)} 守卫取消→优雅收尾（工具零执行）去重=${hasDedup} 空值=${hasFill} events=${events.join(',')}`);
    return good;
}

async function main() {
    const c = await testChart();
    const cl = await testClean();
    console.log(c && cl ? '\nALL CHART+CLEAN E2E PASS' : '\nSOME FAILED');
    if (!(c && cl)) process.exit(1);
}

main().catch(e => { console.error('MAIN FAIL:', e.message); process.exit(1); });
if (require.main === module) { setTimeout(() => process.exit(0), 5000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
