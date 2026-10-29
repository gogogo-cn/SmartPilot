'use strict';
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，但截断到 chat-orchestrator（不加载 chat-entry/index，避免入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

function setupCrud() {
    const App = loadChatV2(ORDER).App;
    App.cachedSkills = [{ name: 'ExcelCRUD', displayName: 'Excel 数据增删改查', keywords: ['读取', '写入', '插入行', '删除行', '清空'], inputContext: [], pipeline: 'skill', processingMode: 'SkillSpec' }];
    const tr = App.ChatV2.toolRegistry;
    const called = {};
    const mock = { ok: true, result: { success: true } };
    [
        'excel_read_range', 'excel_set_cell_value', 'excel_write_range',
        'excel_insert_row', 'excel_insert_rows', 'excel_insert_column',
        'excel_delete_row', 'excel_delete_column',
        'excel_clear_range', 'excel_clear_format'
    ].forEach(n => tr.register({
        name: n, displayName: n, description: n,
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) { called[n] = args; return mock; }
    }));
    return { App, tr, called };
}

function streamSpec(specJson) {
    return { stream({ messages }) { return (async function* () {
        yield { type: 'model', model: 't' };
        yield { type: 'token', delta: JSON.stringify(specJson) };
        yield { type: 'done', finishReason: 'stop' };
    })(); } };
}

async function runCrud(specJson, message, ctx) {
    const { App, called } = setupCrud();
    App.ChatV2.llmClient = streamSpec(specJson);
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({
        message, kind: 'text',
        context: Object.assign({ selection: { address: 'A1:B10', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }, ctx)
    });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    let failed = null;
    for await (const evt of gen) { events.push(evt.type); if (evt.type === 'error') failed = evt.error; }
    return { ok: !failed && events.includes('pipeline_done'), failed, called, events };
}

async function testWriteSingle() {
    const r = await runCrud(
        { operation: 'write', sheet: 'Sheet1', range: 'A1', value: '张三' },
        '把A1改成张三', {});
    const args = r.called['excel_set_cell_value'];
    const good = r.ok && args && args.cell === 'A1' && args.value === '张三';
    console.log(`写单值: ${good ? '✅' : '❌ ' + JSON.stringify(r.failed)} args=${JSON.stringify(args)}`);
    return good;
}

async function testWriteBulk() {
    const r = await runCrud(
        { operation: 'write', sheet: 'Sheet1', range: 'A1:B2', values: [['a', 'b'], ['c', 'd']] },
        '把数据批量写入A1:B2', {});
    const args = r.called['excel_write_range'];
    let decoded = null;
    try { decoded = JSON.parse(args?.values); } catch (e) { /* 保持 null */ }
    const good = r.ok && args && Array.isArray(decoded) && decoded.length === 2 && args.range === 'A1:B2';
    console.log(`写批量: ${good ? '✅' : '❌ ' + JSON.stringify(r.failed)} range=${args?.range} rows=${Array.isArray(decoded) ? decoded.length : args?.values}`);
    return good;
}

// excel_clear_range 是危险工具：headless 环境 fail-safe 到拒绝（安全守卫生效）。
// 2026-10-21 行为修正：守卫取消后 spec 层优雅收尾（输出取消说明 + pipeline_done），
// 不再降级 fallbackToText 让 agent-loop 拿同一条消息二次执行（实测会又去调 read_range）。
// 因此断言：pipeline 正常终态、无 error、且没有任何工具被真正执行。
async function testClear() {
    const r = await runCrud(
        { operation: 'clear', sheet: 'Sheet1', range: 'A1:B10', clearContent: true, clearFormat: false },
        '清空A1:B10的内容但保留格式', {});
    const good = r.ok && !r.failed && Object.keys(r.called).length === 0;
    console.log(`清空[守卫拦截]: ${good ? '✅' : '❌ ' + JSON.stringify(r.failed)} events=${r.events.join(',')} called=${Object.keys(r.called).join(',') || '无'}`);
    return good;
}

// 守卫直连单测：绕过 pipeline，直接 invoke 危险工具，验证守卫本身。
// 这是"守卫生效"的权威断言点（pipeline 层面的 called 为空只是间接证据）。
async function testGuardDirect() {
    const { App, called } = setupCrud();
    const r = await App.ChatV2.toolRegistry.invoke('excel_clear_range', {
        operation: 'clear', sheet: 'Sheet1', range: 'A1:B10', clearContent: true, clearFormat: false
    }, {});
    const good = r.ok === false
        && (r.cancelled === true || r.guarded === true)
        && /excel_clear_range/.test(String(r.error || ''))
        && !called['excel_clear_range'];
    console.log(`守卫直连: ${good ? '✅' : '❌ ' + JSON.stringify({ ok: r.ok, cancelled: r.cancelled, guarded: r.guarded, err: String(r.error || '').slice(0, 80) })}`);
    return good;
}

async function testClearFormatOnly() {
    const r = await runCrud(
        { operation: 'clear', sheet: 'Sheet1', range: 'A1:B10', clearContent: false, clearFormat: true },
        '只清除A1:B10的格式', {});
    const args = r.called['excel_clear_format'];
    const good = r.ok && args;
    console.log(`清格式: ${good ? '✅' : '❌ ' + JSON.stringify(r.failed)} tool=${r.called['excel_clear_format'] ? 'excel_clear_format' : 'NONE'}`);
    return good;
}

async function testRead() {
    const r = await runCrud(
        { operation: 'read', sheet: 'Sheet1', range: 'A1:B10' },
        '读取A1:B10的数据', {});
    const args = r.called['excel_read_range'];
    const good = r.ok && args && args.range === 'A1:B10';
    console.log(`读取: ${good ? '✅' : '❌ ' + JSON.stringify(r.failed)} range=${args?.range}`);
    return good;
}

// spec 解析兜底：LLM 只吐自然语言 → crudSpecFromMessage 兜底识别
async function testSpecFallback() {
    const { App, called } = setupCrud();
    App.ChatV2.llmClient = { stream() { return (async function* () {
        yield { type: 'model', model: 't' };
        yield { type: 'token', delta: '好的，我来处理' };
        yield { type: 'done', finishReason: 'stop' };
    })(); } };
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({ message: '读取A1:B10的数据', kind: 'text', context: { selection: { address: 'A1:B10', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] } });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    let failed = null;
    for await (const evt of gen) { events.push(evt.type); if (evt.type === 'error') failed = evt.error; }
    const args = called['excel_read_range'];
    const good = !failed && args && args.range === 'A1:B10';
    console.log(`兜底路由: ${good ? '✅' : '❌ ' + JSON.stringify(failed)} range=${args?.range}`);
    return good;
}

async function main() {
    const w1 = await testWriteSingle();
    const wb = await testWriteBulk();
    const cl = await testClear();
    const gd = await testGuardDirect();
    const cf = await testClearFormatOnly();
    const rd = await testRead();
    const fb = await testSpecFallback();
    const all = w1 && wb && cl && gd && cf && rd && fb;
    console.log(all ? '\nALL CRUD E2E PASS' : '\nSOME FAILED');
    if (!all) process.exit(1);
}

main().catch(e => { console.error('MAIN FAIL:', e.message); process.exit(1); });