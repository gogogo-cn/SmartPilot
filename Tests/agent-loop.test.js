/**
 * ============================================================
 * Tests/agent-loop.test.js — agent-loop 状态机测试
 * ============================================================
 * 覆盖：turn/step 生命周期、事件广播、取消、异常兜底、假执行拦截
 * 运行：node Tests/agent-loop.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

/** 创建最小 llmClient stub：立即返回 done */
function stubLLM(behavior) {
    return {
        stream({ messages }) {
            return (async function* () {
                yield { type: 'model', model: 'test' };
                if (typeof behavior === 'function') {
                    yield* behavior(messages);
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}

test('agent-loop: turn/step 生命周期事件序列', async () => {
    const { App } = loadChatV2();
    App.ChatV2.llmClient = stubLLM();

    const ev = App.ChatV2.agentEvents;
    const eventsLog = [];
    const offs = [
        ev.on('turn/start', () => eventsLog.push('turn/start')),
        ev.on('agent/status', p => eventsLog.push('status:' + p.status)),
        ev.on('step/start', p => eventsLog.push('step/start:' + p.step)),
        ev.on('turn/end', p => eventsLog.push('turn/end:' + p.reason))
    ];

    const gen = App.ChatV2.agentLoop.run({
        turn: { id: 't1', kind: 'text', userMessage: 'hi' },
        messages: [{ role: 'user', content: 'hi' }],
        tools: [],
        signal: new AbortController().signal
    });
    const first = await gen.next();
    assert.strictEqual(first.value.type, 'round_start');

    // 允许额外的 status:running（phase/细粒度刷新），但必须按序包含核心生命周期
    const expected = ['turn/start', 'status:running', 'step/start:1'];
    let ei = 0;
    for (const e of eventsLog) {
        if (e === expected[ei]) ei++;
        if (ei >= expected.length) break;
    }
    assert.strictEqual(ei, expected.length,
        '生命周期应含 turn/start → status:running → step/start:1，实际: ' + JSON.stringify(eventsLog));
    offs.forEach(f => f());
});

test('agent-loop: 取消路径广播 turn/end 并清理状态', async () => {
    const { App } = loadChatV2();
    App.ChatV2.llmClient = stubLLM();

    const ev = App.ChatV2.agentEvents;
    const eventsLog = [];
    const off = ev.on('turn/end', p => eventsLog.push('turn/end:' + p.reason));

    const ctrl = new AbortController();
    ctrl.abort();
    const gen = App.ChatV2.agentLoop.run({
        turn: { id: 't2', kind: 'text', userMessage: 'hi' },
        messages: [{ role: 'user', content: 'hi' }],
        tools: [],
        signal: ctrl.signal
    });
    const evt = await gen.next();
    assert.strictEqual(evt.value.type, 'canceled');
    await gen.next(); // 触发 finally
    assert.ok(eventsLog.includes('turn/end:user_aborted'));
    assert.strictEqual(App.ChatV2.agentLoop.getRunStatus('t2'), null);
    off();
});

test('agent-loop: llm 异常时错误传播 + turn/end 兜底', async () => {
    const { App } = loadChatV2();
    App.ChatV2.llmClient = { stream() { throw new Error('llm boom'); } };

    const ev = App.ChatV2.agentEvents;
    const eventsLog = [];
    const off = ev.on('turn/end', p => eventsLog.push('turn/end:' + p.reason));

    const gen = App.ChatV2.agentLoop.run({
        turn: { id: 't3', kind: 'text', userMessage: 'hi' },
        messages: [{ role: 'user', content: 'hi' }],
        tools: [],
        signal: new AbortController().signal
    });
    await gen.next(); // round_start
    let threw = false;
    try { await gen.next(); } catch (e) { threw = true; }
    assert.ok(threw, 'llm 异常应传播');
    assert.ok(eventsLog.some(e => e.startsWith('turn/end')));
    off();
});

test('agent-loop: pre-step 钩子 rewrite/reject/inject', async () => {
    const { App } = loadChatV2();
    App.ChatV2.llmClient = stubLLM();
    const al = App.ChatV2.agentLoop;

    // rewrite 钩子
    let actions = [];
    const offEv = App.ChatV2.agentEvents.on('agent/pre-step', p => actions.push(p.action));
    const off = al.onPreStep('test', async (ctx) => ({ action: 'rewrite', messages: [{ role: 'user', content: 'rewritten' }] }));
    const gen = al.run({
        turn: { id: 't4', kind: 'text', userMessage: 'x' },
        messages: [{ role: 'user', content: 'x' }],
        tools: [],
        signal: new AbortController().signal
    });
    for await (const _e of gen) { /* drain */ }
    assert.ok(actions.includes('rewrite'));
    off(); offEv();
});

test('agent-loop: getRunStatus/listRunStates API', async () => {
    const { App } = loadChatV2();
    assert.strictEqual(typeof App.ChatV2.agentLoop.getRunStatus, 'function');
    assert.strictEqual(typeof App.ChatV2.agentLoop.listRunStates, 'function');
    assert.strictEqual(App.ChatV2.agentLoop.getRunStatus('nope'), null);
});

// ── 原生 tool_calls 路径工具名校验回归测试 ──

/** 收集 agentLoop.run 产生的全部事件类型 */
async function collectEvents(al, runOpts) {
    const types = [];
    for await (const e of al.run(runOpts)) {
        types.push(e);
    }
    return types;
}

test('agent-loop: 原生 tool_calls 未注册工具名 → force_retry(invalid_tool_name) 且不执行', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    // 注册一个真实工具，确认它不会被误执行
    let invoked = 0;
    tr.register({
        name: 'excel_read_range', displayName: '读取选区数据', description: '读取数据', source: 'mcp',
        async execute() { invoked++; return { ok: true, result: 'DATA' }; }
    });

    // 模型在原生 tool_calls 里编造了不存在的工具名
    App.ChatV2.llmClient = stubLLM(() => [
        { type: 'model', model: 'test' },
        { type: 'tool_calls', toolCalls: [{ id: 'c1', function: { name: 'fake_unregistered_tool', arguments: '{}' } }] },
        { type: 'done', finishReason: 'tool_calls' }
    ]);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '读取数据', kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '读取数据' }],
        tools: [],
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry');
    assert.ok(retry, '应触发 force_retry');
    assert.strictEqual(retry.reason, 'invalid_tool_name');
    // 关键断言：编造的工具与已注册工具都不得被执行
    assert.ok(!events.some(e => e.type === 'tool_call_result'), '未注册工具不得被实际执行');
    assert.strictEqual(invoked, 0, '已注册工具也不得被误执行');
});

test('agent-loop: 原生 tool_calls 已注册工具 → 正常执行 tool_call_result（加固不误伤）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let invoked = 0;
    tr.register({
        name: 'excel_read_range', displayName: '读取选区数据', description: '读取数据', source: 'mcp',
        async execute() { invoked++; return { ok: true, result: 'DATA' }; }
    });

    App.ChatV2.llmClient = stubLLM(() => [
        { type: 'model', model: 'test' },
        { type: 'tool_calls', toolCalls: [{ id: 'c1', function: { name: 'excel_read_range', arguments: '{}' } }] },
        { type: 'done', finishReason: 'tool_calls' }
    ]);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '读取数据', kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '读取数据' }],
        tools: [],
        signal: new AbortController().signal
    });

    const resultEvt = events.find(e => e.type === 'tool_call_result');
    assert.ok(resultEvt, '应正常执行已注册工具');
    assert.strictEqual(resultEvt.toolName, 'excel_read_range');
    assert.strictEqual(resultEvt.ok, true);
    assert.ok(invoked >= 1, '已注册工具应被调用');
    assert.ok(!events.some(e => e.type === 'force_retry'), '已注册工具不应触发 force_retry');
});

// ── 文本工具调用路径回归测试（stripTextToolCalls 缺参崩溃） ──

test('agent-loop: 文本 [tool] 标记 → 解析执行 + 正文清洗，不再 toolRegistry 未定义崩溃', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let invoked = 0;
    // 用非 OCR、非危险类工具，避开场景安全闸（OCR 场景闸/危险操作守卫），聚焦文本解析路径本身
    tr.register({
        name: 'excel_search', displayName: '搜索数据', description: '搜索数据', source: 'mcp',
        async execute() { invoked++; return { ok: true, result: 'SEARCH_DATA' }; }
    });

    // 复现真实故障：小模型在正文输出 [excel_search] 文本标记（qwen2.5:7b 实测行为）。
    // 有状态桩：首轮带标记，后续轮纯文本总结（配合 summary_after_tools 收尾，避免撞 max_rounds）
    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1) {
            return [
                { type: 'model', model: 'test' },
                { type: 'token', delta: '好的，我来搜索 ' },
                { type: 'token', delta: '[excel_search]' },
                { type: 'done', finishReason: 'stop' }
            ];
        }
        return [
            { type: 'token', delta: '搜索完成。' },
            { type: 'done', finishReason: 'stop' }
        ];
    });

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '搜索数据', kind: 'text' });
    // 修复前此处抛 ReferenceError: toolRegistry is not defined → 整轮 PIPELINE_ERROR
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '搜索数据' }],
        tools: [],
        signal: new AbortController().signal
    });

    assert.ok(events.length > 0, 'run 应正常完成而非抛异常');
    const resultEvt = events.find(e => e.type === 'tool_call_result');
    assert.ok(resultEvt, '文本标记应被解析为工具调用并执行');
    assert.strictEqual(resultEvt.toolName, 'excel_search');
    assert.strictEqual(resultEvt.ok, true, '工具应执行成功，实际: ' + JSON.stringify(resultEvt).slice(0, 200));
    assert.strictEqual(invoked, 1, '应恰好执行一次');
});

// ── 真实失败（9/9 凌晨）：agent-loop 中 LLM 自由调用 excel_aggregate 未传 targetSheet
//    → 后端纯聚合不落盘；随后 LLM 凭直觉写 dataRange='A1:B6' + sheet='销售数据'
//    → 后端按显式 sheet 读原表 A1:B6（B 列是文本分类）→ "没有可绘制的数值列"。
// 修复：① 聚合调用自动补 targetSheet='聚合结果'（落盘）；
//       ② 聚合成功后把 writtenTo/groupCount 记录到 turn._aggregateResult；
//       ③ 后续 excel_create_chart 未显式引用聚合结果表时，强制 sheet+dataRange 指向聚合结果表。
test('agent-loop: 聚合→图表数据源贯通（LLM 自由调用不再读错表，饼图落在聚合结果上）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let aggArgs = null;
    let chartArgs = null;
    const callOrder = [];
    tr.register({
        name: 'excel_aggregate', displayName: '聚合分析', description: '分组聚合', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            aggArgs = args; callOrder.push('aggregate');
            return {
                ok: true,
                result: JSON.stringify({
                    success: true, aggregated: true, dimColumn: '部门', valueColumn: '销售额',
                    aggFunc: 'sum', groupCount: 5, writtenTo: '聚合结果',
                    data: [
                        { 部门: '销售一部', 销售额: 4889788 },
                        { 部门: '销售二部', 销售额: 663721 },
                        { 部门: '市场部', 销售额: 547261 },
                        { 部门: '销售三部', 销售额: 519878 },
                        { 部门: '大客户部', 销售额: 467701 }
                    ]
                })
            };
        }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; callOrder.push('chart'); return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    // LLM 两轮工具调用 + 一轮总结（模拟真实模型行为：聚合不带 targetSheet，图表写错表）
    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c1', function: { name: 'excel_aggregate', arguments: JSON.stringify({ sheet: '销售数据', range: '$A$1:$I$101', dimColumn: '部门', valueColumn: '销售额', aggFunc: 'sum' }) } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        if (llmCall === 2) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c2', function: { name: 'excel_create_chart', arguments: JSON.stringify({ chartType: 'xlPie', title: '部门销售额占比', dataRange: 'A1:B6', sheet: '销售数据' }) } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        return [
            { type: 'token', delta: '已完成。' },
            { type: 'done', finishReason: 'stop' }
        ];
    });

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '使用饼图按部门展示销售额',
        kind: 'text',
        context: { selection: { address: 'A1:I101', sheetName: '销售数据' } }
    });
    // 「聚合 → 图表」是多步骤请求：真实链路（text-pipeline 的 _forceMultiStep / 多步骤管道）
    // 会以 persistTools=true 调用 agent-loop；persistTools=false（默认）是"单轮工具 + 总结"，
    // 第 1 个工具成功后不再下发 tools，第 2 个工具会被终态收尾闸门忽略（agent-loop.js:2393）。
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '使用饼图按部门展示销售额' }],
        tools: [],
        persistTools: true,
        signal: new AbortController().signal
    });

    assert.ok(aggArgs, '应调用 excel_aggregate');
    assert.strictEqual(aggArgs.targetSheet, '聚合结果', '聚合调用应自动补 targetSheet 落盘（后端纯聚合不落盘）');
    assert.strictEqual(aggArgs.dimColumn, '部门');
    assert.ok(chartArgs, '应调用 excel_create_chart');
    assert.deepStrictEqual(callOrder, ['aggregate', 'chart'], '应先聚合后画图');
    assert.strictEqual(chartArgs.sheet, '聚合结果',
        `图表应强制指向聚合结果表，实际 sheet=${chartArgs.sheet}, dataRange=${chartArgs.dataRange}`);
    assert.strictEqual(chartArgs.dataRange, '聚合结果!A1:B6',
        `图表应引用聚合结果表（5 组 + 表头 = B6），实际: ${chartArgs.dataRange}`);
    assert.strictEqual(chartArgs.chartType, 'xlPie', '保留 LLM 规划的图表类型');
    console.log('\n✅ agent-loop 聚合→图表贯通：excel_aggregate 自动落盘 → excel_create_chart 指向 聚合结果!A1:B6，不再读错表');
    assert.ok(!events.some(e => e.type === 'loop_done' && e.reason === 'max_rounds_reached'), '不应撞最大轮数');

    // 正文中的标记应被清洗掉，避免重复显示
    const stored = App.ChatV2.sessionStore.getTurns().find(t => t.id === stTurn.id);
    const assistantMsg = stored?.messages?.find(m => m.role === 'assistant');
    const content = typeof assistantMsg?.content === 'string' ? assistantMsg.content : JSON.stringify(stored?.messages || '');
    assert.ok(!content.includes('[excel_search]'), '正文不应残留工具调用标记，实际: ' + content.slice(0, 120));
});

// ── 边界加固（9/9）：LLM 已显式引用聚合结果表（dataRange='聚合结果!A1:B6'）但 sheet 仍写原表
//    → 后端显式 sheet 优先于 dataRange 前缀 → 读原表 A1:B6 文本列 → "没有可绘制的数值列"。
//    修复：图表调用即使已引用聚合结果表，sheet 仍须与 dataRange 前缀对齐。
test('agent-loop: 图表 sheet 与 dataRange 前缀对齐（LLM 引用聚合结果表但 sheet 写错）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let aggArgs = null;
    let chartArgs = null;
    tr.register({
        name: 'excel_aggregate', displayName: '聚合分析', description: '分组聚合', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            aggArgs = args;
            return {
                ok: true,
                result: JSON.stringify({
                    success: true, aggregated: true, dimColumn: '部门', valueColumn: '销售额',
                    aggFunc: 'sum', groupCount: 5, writtenTo: '聚合结果',
                    data: [{ 部门: '销售一部', 销售额: 100 }, { 部门: '销售二部', 销售额: 200 }]
                })
            };
        }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    // LLM：聚合带了 targetSheet（不触发自动补参），图表引用聚合结果表但 sheet 写错为原表
    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c1', function: { name: 'excel_aggregate', arguments: JSON.stringify({ sheet: '销售数据', range: '$A$1:$I$101', dimColumn: '部门', valueColumn: '销售额', aggFunc: 'sum', targetSheet: '聚合结果' }) } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        if (llmCall === 2) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c2', function: { name: 'excel_create_chart', arguments: JSON.stringify({ chartType: 'xlPie', title: '部门销售额占比', dataRange: '聚合结果!A1:B6', sheet: '销售数据' }) } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        return [
            { type: 'token', delta: '已完成。' },
            { type: 'done', finishReason: 'stop' }
        ];
    });

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '使用饼图按部门展示销售额',
        kind: 'text',
        context: { selection: { address: 'A1:I101', sheetName: '销售数据' } }
    });
    await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '使用饼图按部门展示销售额' }],
        tools: [],
        persistTools: true,   // 同上：多步骤（聚合→图表）真实调用方式
        signal: new AbortController().signal
    });

    assert.ok(aggArgs, '应调用 excel_aggregate');
    assert.ok(chartArgs, '应调用 excel_create_chart');
    // 已引用聚合结果表 → 不强制覆盖 dataRange，但 sheet 必须与 dataRange 前缀对齐
    assert.strictEqual(chartArgs.dataRange, '聚合结果!A1:B6', 'dataRange 保留聚合结果表地址');
    assert.strictEqual(chartArgs.sheet, '聚合结果',
        `sheet 必须与 dataRange 前缀对齐（否则后端读原表文本列报"没有可绘制的数值列"），实际 ${chartArgs.sheet}`);
    console.log('\n✅ agent-loop sheet 前缀对齐：LLM 引用聚合结果表但 sheet 写错时自动修正 → sheet=' + chartArgs.sheet);
});

test('agent-loop: 非空 tools 白名单拒绝名单外工具，且仍回喂失败结果', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let wrote = 0;
    tr.register({
        name: 'excel_write_range', displayName: '写入', description: '写入单元格', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { wrote++; return { ok: true, result: 'written' }; }
    });
    tr.register({
        name: 'excel_read_range', displayName: '读取', description: '读取单元格', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { return { ok: true, result: 'DATA' }; }
    });

    App.ChatV2.llmClient = stubLLM(() => [
        { type: 'model', model: 'test' },
        { type: 'tool_calls', toolCalls: [{ id: 'c_deny', function: { name: 'excel_write_range', arguments: '{}' } }] },
        { type: 'done', finishReason: 'tool_calls' }
    ]);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '只读一下', kind: 'text', context: {} });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '只读一下' }],
        tools: [{ name: 'excel_read_range' }],
        maxIterations: 1,
        signal: new AbortController().signal
    });

    const denied = events.find(e => e.type === 'tool_call_result' && e.toolName === 'excel_write_range');
    assert.ok(denied, '名单外工具应回一条 tool_call_result，避免下一轮缺少 tool 消息');
    assert.strictEqual(denied.ok, false);
    assert.match(String(denied.error || ''), /不在本次允许的工具列表中/);
    assert.strictEqual(wrote, 0, '名单外工具不得执行');
});

test('agent-loop: 总结含「然后」不算未完成信号 → llm_decided_stop', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    tr.register({
        name: 'excel_read_range', displayName: '读取', description: '读取', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { return { ok: true, result: { rows: 2 } }; }
    });

    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c1', function: { name: 'excel_read_range', arguments: '{"range":"A1:B2"}' } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        return [
            { type: 'token', delta: '已读取选区，然后汇总如下：共 2 行。' },
            { type: 'done', finishReason: 'stop' }
        ];
    });

    // 让 hasExcelOperationIntent 在无 routingConstants 时也不挡：直接给操作类问句
    if (!App.ChatV2.routingConstants) {
        App.ChatV2.routingConstants = {
            hasExcelOperationIntent: () => true
        };
    } else {
        App.ChatV2.routingConstants.hasExcelOperationIntent = () => true;
    }

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '读取选区数据', kind: 'text', context: {}
    });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '读取选区数据' }],
        tools: [{ name: 'excel_read_range', description: '读取', inputSchema: { type: 'object', properties: {} } }],
        signal: new AbortController().signal,
        persistTools: true,
        maxIterations: 5
    });

    const done = events.find(e => e.type === 'loop_done');
    assert.ok(done, '应正常收尾');
    assert.strictEqual(done.reason, 'llm_decided_stop',
        '含「然后」的正常总结不应被 INCOMPLETE_SIGNAL 拦住，实际 reason=' + done.reason);
});

// 兜底退出：避免 session-store 防抖定时器挂起进程
if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出