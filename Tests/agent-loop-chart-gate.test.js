/**
 * ============================================================
 * Tests/agent-loop-chart-gate.test.js — 图表交付闸门 + 图表补链
 * ============================================================
 * 背景（2026-10-10 07:46 实测，log_20261010.log 第 1698-1827 行）：
 *   全部内置 Skill 被停用 → Skill 路由失效 → "使用饼图按部门展示销售额" 落通用多步骤管道。
 *   弱模型第 1 轮【零工具调用】只写了 712 字自然语言分析；
 *   零工具调用守卫链补了一次 excel_read_range（$A$1:$I$100）；
 *   覆盖度自检判“未完成=0”（intentKeywords 反查看不出“缺一张图”）→ 直接收尾。
 *   结果：用户拿到一大段文字、零张图，界面却显示“任务完成”。
 *
 * 修复（agent-loop.js）：
 *   ① isChartIntent —— 用“图的存在”而非模型自称来完成来定义图表请求的完成；
 *   ② 图表交付闸门 —— 图表意向下未建图不许收尾，先劝一轮（提示里带具体聚合参数）；
 *   ③ 图表补链 —— 连聚合都没做时，同一轮内确定性合成 aggregate → create_chart。
 *
 * 运行：node Tests/agent-loop-chart-gate.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，截断到入口之前（不加载带副作用的入口）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

function loop() {
    const { App } = loadChatV2();
    return App.ChatV2.agentLoop;
}

/** 造一个“选中了销售明细表”的 turn */
function makeTurn(extra) {
    return Object.assign({
        id: 't-chart',
        kind: 'text',
        userMessage: '使用饼图按部门展示销售额',
        context: {
            selection: {
                address: '$A$1:$C$6',
                range: '$A$1:$C$6',
                sheetName: '销售数据',
                headers: ['部门', '产品', '销售额'],
                rows: [
                    ['销售一部', '产品A', 1024],
                    ['销售二部', '产品B', 2048],
                    ['销售一部', '产品C', 512]
                ],
                rowCount: 3,
                colCount: 3
            }
        },
        _executedToolCalls: [
            { name: 'excel_read_range', args: { range: '$A$1:$C$100', sheet: '销售数据' } }
        ]
    }, extra || {});
}

// ── isChartIntent ────────────────────────────────────────────────

test('闸门: context._forceChart 标记 → 判定为图表请求', () => {
    const L = loop();
    assert.strictEqual(L.isChartIntent({ context: { _forceChart: true }, userMessage: '随便' }, '随便'), true);
});

test('闸门: 消息含明确图表类型词 → 判定为图表请求', () => {
    const L = loop();
    assert.strictEqual(L.isChartIntent({ context: {} }, '使用饼图按部门展示销售额'), true);
    assert.strictEqual(L.isChartIntent({ context: {} }, '按部门画个柱状图'), true);
    assert.strictEqual(L.isChartIntent({ context: {} }, '给我做个折线图看趋势'), true);
});

test('闸门: 疑问句不算出图指令（不强制建图）', () => {
    const L = loop();
    assert.strictEqual(L.isChartIntent({ context: {} }, '饼图怎么做'), false);
    assert.strictEqual(L.isChartIntent({ context: {} }, '饼图和柱状图有什么区别'), false);
    assert.strictEqual(L.isChartIntent({ context: {} }, '什么是折线图？'), false);
});

test('闸门: 普通非图表请求不误命中', () => {
    const L = loop();
    assert.strictEqual(L.isChartIntent({ context: {} }, '把表头加粗'), false);
    assert.strictEqual(L.isChartIntent({ context: {} }, '把 A 列去重'), false);
});

// ── resolveAggregateSource ───────────────────────────────────────

test('补链: 数据区优先取最近一次成功 read_range 的实参，维度/数值列取自选区表头', () => {
    const L = loop();
    const src = L.resolveAggregateSource(makeTurn());
    assert.ok(src, '应能推断出数据源');
    assert.strictEqual(src.range, '$A$1:$C$100', 'range 应取模型自己读过的区域');
    assert.strictEqual(src.sheet, '销售数据');
    assert.strictEqual(src.dimColumn, '部门', '第一个非数值列 = 维度列');
    assert.strictEqual(src.valueColumn, '销售额', '第一个数值列 = 数值列');
});

test('补链: 无 read_range 记录时回退到当前选区地址', () => {
    const L = loop();
    const t = makeTurn({ _executedToolCalls: [] });
    const src = L.resolveAggregateSource(t);
    assert.strictEqual(src.range, '$A$1:$C$6');
    assert.strictEqual(src.sheet, '销售数据');
});

test('补链: 选区无表头 → dimColumn 留空交后端择优（不再直接放弃补链）', () => {
    // 2026-10-28 修改：旧逻辑这里 return null，导致"用户没手动选中 + 上下文无 headers"时补链整体跳过，
    // 模型只能空手收尾（实测"图表没有执行"）。后端已具备取值基数择优 + 1 组退化换列，
    // 因此改为 autoDim 交后端处理；后端仍只有 1 组时补链不会建废图（runChartBackfill 有 _gc<=1 拦截）。
    const L = loop();
    const t = makeTurn();
    t.context.selection = { address: '$A$1:$C$6', sheetName: '销售数据' };
    const src = L.resolveAggregateSource(t);
    assert.ok(src, '没有表头也要能补链');
    assert.strictEqual(src.dimColumn, '');
    assert.strictEqual(src.autoDim, true);
    assert.strictEqual(src.valueColumn, '');
});

test('补链: 完全没有数据区（无选区、无读取）→ 返回 null', () => {
    const L = loop();
    assert.strictEqual(L.resolveAggregateSource({ context: {} }), null);
    assert.strictEqual(L.resolveAggregateSource(null), null);
});

// ── buildChartNudgeHint ──────────────────────────────────────────

test('闸门提示: 含两步强指令 + 具体聚合参数 + 推断出的图表类型', () => {
    const L = loop();
    const h = L.buildChartNudgeHint(makeTurn(), '使用饼图按部门展示销售额');
    assert.strictEqual(h.chartType, 'xlPie', '“饼图”应推断为 xlPie');
    assert.ok(h.text.includes('excel_aggregate'), '要给出第 1 步工具');
    assert.ok(h.text.includes('dimColumn="部门"'), '要带具体维度列名');
    assert.ok(h.text.includes('valueColumn="销售额"'), '要带具体数值列名');
    assert.ok(h.text.includes('excel_create_chart'), '要给出第 2 步工具');
    assert.ok(h.text.includes('xlPie'), '要带推断出的图表类型');
    assert.ok(h.text.includes('严禁直接用原始明细数据建图'), '要显式禁止对明细建废图');
});

test('闸门提示: 推断不出表头时退化为通用占位，不报错', () => {
    const L = loop();
    const h = L.buildChartNudgeHint({ context: {} }, '使用饼图按部门展示销售额');
    assert.strictEqual(h.src, null);
    assert.ok(h.text.includes('dimColumn=分组维度列名'), '应给出通用列名占位');
    assert.ok(h.text.includes('xlPie'));
});

// ── 端到端：读了数据却只输出文字 → 必须补出图 ────────────────────

function stubRounds(plan) {
    let n = 0;
    return {
        stream() {
            const idx = n++;
            const out = plan(idx) || {};
            return (async function* () {
                yield { type: 'model', model: 'test' };
                if (out.toolCalls) yield { type: 'tool_calls', toolCalls: out.toolCalls };
                if (out.text) yield { type: 'token', delta: out.text };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}

function registerChartTools(tr, calls) {
    tr.register({
        name: 'excel_read_range', displayName: '读取数据', description: '读取区域', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            calls.push({ name: 'excel_read_range', args });
            return { ok: true, result: { range: args?.range || '$A$1:$C$6', rows: 3 } };
        }
    });
    tr.register({
        name: 'excel_aggregate', displayName: '分组聚合', description: '聚合', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            calls.push({ name: 'excel_aggregate', args });
            return {
                ok: true,
                result: {
                    writtenTo: args?.targetSheet || '部门汇总',
                    groupCount: 2,
                    data: [{ 部门: '销售一部', 销售额: 1536 }, { 部门: '销售二部', 销售额: 2048 }]
                }
            };
        }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '画图', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            calls.push({ name: 'excel_create_chart', args });
            return { ok: true, result: { chartName: 'AIChart_1' } };
        }
    });
}

const CHART_TOOLS = [
    { name: 'excel_read_range', description: '读取区域', inputSchema: { type: 'object', properties: {} }, intentKeywords: ['读取', '数据'] },
    { name: 'excel_aggregate', description: '分组聚合', inputSchema: { type: 'object', properties: {} }, intentKeywords: ['聚合', '汇总'] },
    { name: 'excel_create_chart', description: '创建图表', inputSchema: { type: 'object', properties: {} }, intentKeywords: ['图表', '饼图'] }
];

function startChartTurn(App) {
    return App.ChatV2.sessionStore.startTurn({
        message: '使用饼图按部门展示销售额',
        kind: 'text',
        context: {
            _forceChart: true,
            selection: {
                address: '$A$1:$C$6', range: '$A$1:$C$6', sheetName: '销售数据',
                headers: ['部门', '产品', '销售额'],
                rows: [['销售一部', '产品A', 1024], ['销售二部', '产品B', 2048], ['销售一部', '产品C', 512]],
                rowCount: 3, colCount: 3
            }
        }
    });
}

test('e2e: 读了数据却只输出文字 → nudge 后当场补链 aggregate→create_chart（不再零图收场）', async () => {
    const { App } = loadChatV2(ORDER);
    const tr = App.ChatV2.toolRegistry;
    const calls = [];
    registerChartTools(tr, calls);

    // 第 1 轮读数据（只读探测 → 轮末不抢跑补链）；随后两轮纯文字：先 nudge 一次，再用尽后当场补链
    App.ChatV2.llmClient = stubRounds(i => (i === 0)
        ? { toolCalls: [{ id: 'c1', name: 'excel_read_range', args: { range: '$A$1:$C$100', sheet: '销售数据' } }] }
        : { text: '总体来看，选中数据涵盖了从 2024 年 1 月 9 日至 2024 年 1 月 15 日的销售记录，销售额合计为 2657874 元。' });

    const turn = startChartTurn(App);
    const events = [];
    const gen = App.ChatV2.agentLoop.run({
        turn,
        messages: [{ role: 'user', content: '使用饼图按部门展示销售额' }],
        tools: CHART_TOOLS,
        signal: new AbortController().signal,
        persistTools: true,
        maxIterations: 8
    });
    for await (const ev of gen) { events.push(ev); }

    const names = calls.map(c => c.name).join(',');
    const agg = calls.find(c => c.name === 'excel_aggregate');
    const chart = calls.find(c => c.name === 'excel_create_chart');
    const nudged = events.some(e => e.type === 'force_retry' && e.reason === 'chart_not_generated');
    const done = events.find(e => e.type === 'loop_done');

    assert.ok(nudged, '应先触发一次图表补图引导');
    assert.ok(agg, 'nudge 用尽后应当场合成 excel_aggregate；实际调用=' + names);
    assert.ok(chart, 'nudge 用尽后应当场合成 excel_create_chart；实际调用=' + names);
    assert.strictEqual(agg.args.dimColumn, '部门');
    assert.strictEqual(agg.args.valueColumn, '销售额');
    assert.strictEqual(agg.args.aggFunc, 'sum');
    assert.strictEqual(chart.args.chartType, 'xlPie', '应按用户说的饼图建图');
    assert.ok(/A1:B3$/.test(String(chart.args.dataRange)), '数据源应为聚合结果区域，实际=' + chart.args.dataRange);
    assert.strictEqual(turn._chartDone, true, '补图成功应置位 _chartDone');
    assert.ok(done && done.reason === 'chart_backfill_done', '应收尾原因为 chart_backfill_done，实际=' + (done && done.reason));
});

test('e2e: nudge 用尽后不得以 no_tool_calls 零图收场', async () => {
    const { App } = loadChatV2(ORDER);
    const calls = [];
    registerChartTools(App.ChatV2.toolRegistry, calls);
    App.ChatV2.llmClient = stubRounds(i => (i === 0)
        ? { toolCalls: [{ id: 'c1', name: 'excel_read_range', args: { range: '$A$1:$C$100', sheet: '销售数据' } }] }
        : { text: '已完成分析，建议按部门查看销售额分布。' });

    const turn = startChartTurn(App);
    const events = [];
    for await (const ev of App.ChatV2.agentLoop.run({
        turn,
        messages: [{ role: 'user', content: '使用饼图按部门展示销售额' }],
        tools: CHART_TOOLS,
        signal: new AbortController().signal,
        persistTools: true,
        maxIterations: 8
    })) { events.push(ev); }

    const done = events.find(e => e.type === 'loop_done');
    assert.notStrictEqual(done && done.reason, 'no_tool_calls', '不得以 no_tool_calls 收尾');
    assert.ok(calls.some(c => c.name === 'excel_create_chart'), '必须产出图表');
    assert.strictEqual(turn._chartDone, true);
});
