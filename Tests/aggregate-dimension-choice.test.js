/**
 * ============================================================
 * Tests/aggregate-dimension-choice.test.js — 聚合维度「择优」回归
 * ============================================================
 * 背景（2026-10-28 实测截图）：源表「地区」列 20 行全是"华南"，模型按"地区"聚合 →
 *   只有 1 组 → 图表只有一根柱/一个扇区，而模型照常收尾说"已完成"。
 * 根因链：
 *   1) VB ExecuteAggregate：dimColumn 只要在表头里匹配上就照做，从不看数据分布；
 *      所谓"择优"只在「列匹配不上」时才跑（且只按表头关键词打分，不看基数）。
 *   2) chart-gate.resolveAggregateSource（图表补链）取"第一个非数值列"，
 *      本例会选中「日期」列，比选「地区」更糟。
 *   3) 结果退化无任何检测：groupCount=1 只被打印进产物摘要，模型不会换维度重试。
 * 修复：
 *   1) VB/py 桥：扫描各分类列 distinct → 基数分(权重2)+关键词分(权重1) 择优；
 *      聚合只有 1 组时自动改用评分最高的列重算，并回传 dimSwitched /
 *      dimCandidates / dimSwitchNote。
 *   2) chart-gate：按基数择优选维度，跳过日期列与 distinct<=1 的列。
 *   3) agent-loop 产物摘要：groupCount=1 时注入"必须换维度重算"的硬提示。
 * 运行：node Tests/aggregate-dimension-choice.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

const VB_SRC = fs.readFileSync(
    path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.Advanced.Aggregate.vb'), 'utf8');
const PY_SRC = fs.readFileSync(path.join(ROOT, 'Tools', 'py-excel-bridge.py'), 'utf8');
const AGENT_LOOP_SRC = fs.readFileSync(
    path.join(ROOT, 'UI', 'js', 'chat-v2', 'agent-loop.js'), 'utf8');

/**
 * 构造截图同款选区：地区列 20 行全是"华南"（distinct=1），
 * 部门 4 个取值、产品 5 个、客户类型 3 个、销售员 20 行全唯一（明细列）、日期 20 个。
 */
function buildSelection() {
    const headers = ['日期', '地区', '部门', '产品', '销售额', '成本', '利润', '销售员', '客户类型'];
    const depts = ['市场部', '销售部', '电商部', '渠道部'];
    const products = ['A-100', 'B-200', 'C-300', 'D-400', 'E-500'];
    const custTypes = ['企业客户', '个人客户', '政府客户'];
    const rows = [];
    for (let i = 0; i < 20; i++) {
        rows.push([
            '2024/1/' + (i + 1),
            '华南',
            depts[i % 4],
            products[i % 5],
            1000 + i * 37,
            600 + i * 12,
            400 + i * 25,
            '销售员' + (i + 1),
            custTypes[i % 3]
        ]);
    }
    return { address: 'A1:I21', sheetName: '华南市场部', headers, rows };
}

function buildTurn(sel) {
    return { id: 't1', userMessage: '对选中区域生成图表', context: { selection: sel } };
}

test('chart-gate: 维度列只有 1 个取值时不选它（"地区"全表"华南" → 改选"部门"）', () => {
    const App = loadChatV2(ORDER).App;
    const src = App.ChatV2.agentLoopChartGate.resolveAggregateSource(buildTurn(buildSelection()));
    assert.ok(src, '应能推断出聚合数据源');
    assert.notStrictEqual(src.dimColumn, '地区', '「地区」只有 1 个取值，不能当维度');
    assert.notStrictEqual(src.dimColumn, '日期', '「日期」不能当聚合维度');
    assert.strictEqual(src.dimColumn, '部门', '4 个取值 + 表头最典型，应为最优维度');
});

test('chart-gate: 数值列仍按第一个数值列取（销售额），不因维度择优受影响', () => {
    const App = loadChatV2(ORDER).App;
    const src = App.ChatV2.agentLoopChartGate.resolveAggregateSource(buildTurn(buildSelection()));
    assert.strictEqual(src.valueColumn, '销售额');
    assert.strictEqual(src.range, 'A1:I21');
    assert.strictEqual(src.sheet, '华南市场部');
});

test('chart-gate: 明细列（销售员 20 行全唯一）不优先于正常分类列', () => {
    const App = loadChatV2(ORDER).App;
    const sel = buildSelection();
    // 只留下"日期 + 销售员（全唯一）+ 销售额"，此时只能选销售员，不应返回 null
    const src = App.ChatV2.agentLoopChartGate.resolveAggregateSource(buildTurn({
        address: 'A1:B21',
        sheetName: 'S1',
        headers: ['日期', '销售员'],
        rows: sel.rows.map(r => [r[0], r[7]])
    }));
    assert.ok(src, '应有兜底维度，不能因为都是弱候选就放弃补链');
    assert.strictEqual(src.dimColumn, '销售员');
});

test('chart-gate: 所有分类列都只有 1 个取值时退回第一个非数值列（不返回 null）', () => {
    const App = loadChatV2(ORDER).App;
    const src = App.ChatV2.agentLoopChartGate.resolveAggregateSource(buildTurn({
        address: 'A1:C5',
        sheetName: 'S2',
        headers: ['日期', '地区', '销售额'],
        rows: [
            ['2024/1/1', '华南', 100],
            ['2024/1/2', '华南', 200],
            ['2024/1/3', '华南', 300]
        ]
    }));
    assert.ok(src);
    assert.strictEqual(src.dimColumn, '地区', '数据本身无区分度 → 退回旧行为，交回模型判断');
});

test('VB: ExecuteAggregate 具备维度择优（distinct 扫描 / 退化换列 / 结果回传）', () => {
    for (const needle of [
        'CountDistinctValues',      // 各列取值基数
        'ScoreDimCandidate',        // 基数分(权重2) + 关键词分(权重1)
        'AggregateByDimColumn',     // 抽成方法以便换列重算
        'dimCandidates',            // 候选维度回传，供模型继续择优
        'dimSwitched',
        'dimSwitchNote',
        'KeywordDimScore'           // 打分表单点维护（兜底选择与择优共用）
    ]) {
        assert.ok(VB_SRC.includes(needle), 'VB 聚合实现应包含 ' + needle);
    }
    // 退化判定必须基于真实分组结果，而不是只看列名
    const reSwitch = /If\s+aggResults\.Count\s*<=\s*1\s+Then/;
    assert.ok(reSwitch.test(VB_SRC), '应有"聚合结果只有 1 组"的退化检测');
});

test('py 桥: m_excel_aggregate 具备同款维度择优', () => {
    for (const needle of [
        '_agg_dim_candidates',
        '_agg_group_rows',
        '_agg_distinct',
        '_agg_is_numeric_col',
        'dimSwitchNote',
        'dimCandidates'
    ]) {
        assert.ok(PY_SRC.includes(needle), 'py 桥应包含 ' + needle);
    }
    assert.ok(/if\s+len\(out_rows\)\s*<=\s*1:/.test(PY_SRC), 'py 桥应有 1 组退化检测');
});

test('chart-gate: 没有选区表头时不再放弃补链（dimColumn 留空交后端择优）', () => {
    const App = loadChatV2(ORDER).App;
    const trn = {
        id: 't2',
        context: {},
        _executedToolCalls: [{ name: 'excel_read_range', args: { range: 'A1:21', sheet: '市场部' } }]
    };
    const src = App.ChatV2.agentLoopChartGate.resolveAggregateSource(trn);
    assert.ok(src, '拿不到选区 headers 也必须能补链（旧逻辑这里 return null → 模型空手收尾）');
    assert.strictEqual(src.dimColumn, '');
    assert.strictEqual(src.autoDim, true);
    assert.strictEqual(src.range, 'A1:21');
    assert.strictEqual(src.sheet, '市场部');
});

test('chart-gate: 补链在无表头时仍执行 aggregate（dimColumn 缺省、不传行简写 range）', async () => {
    const App = loadChatV2(ORDER).App;
    const calls = [];
    App.ChatV2.toolRegistry.invokeTracked = async (trn, name, args) => {
        calls.push({ name, args });
        if (name === 'excel_aggregate') {
            return {
                ok: true, durationMs: 1,
                result: { groupCount: 2, writtenTo: '聚合结果', dimColumn: '部门', data: [{ a: 1 }, { a: 2 }] }
            };
        }
        if (name === 'excel_create_chart') return { ok: true, durationMs: 1, result: { chartName: '图表 1' } };
        return { ok: false };
    };
    const gate = App.ChatV2.agentLoopChartGate;
    const p = {
        turn: {
            id: 't3', context: {},
            _executedToolCalls: [{ name: 'excel_read_range', args: { range: 'A1:21', sheet: '市场部' } }]
        },
        toolCalls: [], localResults: [], round: 1, stepNo: 1, rawMsg: '对选中区域生成图表'
    };
    const gen = gate.runChartBackfill(p);
    let r = await gen.next();
    while (!r.done) r = await gen.next();
    assert.strictEqual(r.value.chartDone, true);
    const agg = calls.find(c => c.name === 'excel_aggregate');
    assert.ok(agg, '应自动合成 excel_aggregate');
    assert.strictEqual(agg.args.dimColumn, undefined, '维度留空，由后端按取值基数择优');
    assert.strictEqual(agg.args.range, undefined, 'A1:21 这类行简写不直接传后端（会被解析成整行 16384 列）');
    assert.strictEqual(agg.args.sheet, '市场部');
    const chart = calls.find(c => c.name === 'excel_create_chart');
    assert.ok(chart);
    assert.strictEqual(chart.args.dataRange, '聚合结果!A1:B3');
});

test('chart-gate: 后端择优后仍只有 1 组时不建废图', async () => {
    const App = loadChatV2(ORDER).App;
    const calls = [];
    App.ChatV2.toolRegistry.invokeTracked = async (trn, name, args) => {
        calls.push({ name, args });
        if (name === 'excel_aggregate') {
            return { ok: true, durationMs: 1, result: { groupCount: 1, writtenTo: '聚合结果', data: [{ a: 1 }] } };
        }
        return { ok: true, durationMs: 1, result: { chartName: '图表 1' } };
    };
    const gate = App.ChatV2.agentLoopChartGate;
    const p = {
        turn: {
            id: 't4', context: {},
            _executedToolCalls: [{ name: 'excel_read_range', args: { range: 'A1:21', sheet: '市场部' } }]
        },
        toolCalls: [], localResults: [], round: 1, stepNo: 1, rawMsg: '对选中区域生成图表'
    };
    const gen = gate.runChartBackfill(p);
    let r = await gen.next();
    while (!r.done) r = await gen.next();
    assert.strictEqual(r.value.chartDone, false);
    assert.strictEqual(calls.filter(c => c.name === 'excel_create_chart').length, 0,
        '1 根柱的废图不该产出，应如实交回模型说明数据没有可分组维度');
});

test('agent-loop: 聚合只有 1 组时向模型注入"换维度重算"的硬提示', () => {
    assert.ok(AGENT_LOOP_SRC.includes('必须换维度重算 excel_aggregate'),
        '产物摘要应在 groupCount=1 时要求换维度，而不是让模型照常收尾');
    assert.ok(AGENT_LOOP_SRC.includes('没有区分度'), '应说明 1 组结果为什么是废图');
    assert.ok(AGENT_LOOP_SRC.includes('a.dimSwitched'), '应把后端自动换维度的事实告知模型');
    assert.ok(AGENT_LOOP_SRC.includes('a.dimCandidates'), '应把候选维度（含基数）交给模型继续择优');
});
