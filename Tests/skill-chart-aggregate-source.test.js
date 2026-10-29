/**
 * ============================================================
 * Tests/skill-chart-aggregate-source.test.js — 图表数据源贯通回归测试
 * ============================================================
 * 覆盖两个真实失败（用户 9/9 实测复现）：
 *   1. 「使用饼图按部门展示销售额」→ 聚合成功（第二步）但图表失败（第三步）
 *      根因：聚合结果写入「聚合结果」表后，chartArgs.sheet 仍是 plan.sheet（原表），
 *      后端 ResolveRangeSafe 显式 sheet 优先于 dataRange 前缀 → 读原表 A1:Bn 文本列
 *      → "没有可绘制的数值列"。
 *      断言：excel_create_chart 的 sheet/dataRange 必须同指「聚合结果」表。
 *   2. 「对选中区域生成图表」→ LLM 自拟整表 $A$1:$I$101 掩盖真实选区（纯文本分类列）
 *      断言：plan.dataRange 被选区覆盖 → 纯文本频次聚合 → 图表数据源表，
 *      excel_create_chart 的 sheet 同步为「图表数据源」。
 * 运行：node Tests/skill-chart-aggregate-source.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

function setupChartSkill(App) {
    App.cachedSkills = [{
        name: 'SmartChartWizard', displayName: '智能图表向导',
        keywords: ['画图', '图表', '饼图', '柱状图', '折线图'],
        inputContext: [], pipeline: 'skill', processingMode: 'SkillSpec'
    }];
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'skill.SmartChartWizard', displayName: '智能图表向导', description: '图表',
        inputSchema: { type: 'object', properties: {} }, source: 'skill', skillInfo: App.cachedSkills[0],
        async execute() { return { ok: true, result: 'x' }; }
    });
    return tr;
}

async function runSkillPipeline(App, message, context) {
    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({ message, kind: 'text', context });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    let failed = null;
    for await (const evt of gen) {
        events.push(evt.type);
        if (evt.type === 'error') failed = evt.error;
    }
    return { events, failed };
}

test('「使用饼图按部门展示销售额」：聚合成功后图表 sheet 必须贯通为「聚合结果」', async () => {
    const App = loadChatV2(ORDER).App;
    const tr = setupChartSkill(App);

    // 快照读取（图表开头探测）：表含数值列 → 不走纯文本频次路径
    tr.register({
        name: 'excel_read_range', displayName: '读取区域', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) {
            return {
                ok: true,
                result: {
                    headers: ['日期', '地区', '部门', '产品', '销售额', '成本'],
                    result: [
                        { 日期: '2026-01-01', 地区: '华东', 部门: '销售一部', 产品: 'A', 销售额: 100, 成本: 50 },
                        { 日期: '2026-01-02', 地区: '华南', 部门: '销售二部', 产品: 'B', 销售额: 200, 成本: 90 }
                    ]
                }
            };
        }
    });
    // 聚合：落盘「聚合结果」表并返回 writtenTo/groupCount
    let aggArgs = null;
    tr.register({
        name: 'excel_aggregate', displayName: '分组聚合', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) { aggArgs = args; return { ok: true, result: { writtenTo: '聚合结果', groupCount: 5, data: [] } }; }
    });
    // 图表：捕获收到的参数
    let chartArgs = null;
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) { chartArgs = args; return { ok: true, result: { chartName: 'Chart 1' } }; }
    });

    // LLM 规划：模拟真实模型看到上下文后自拟 sheet=销售数据 + 整表 dataRange
    App.ChatV2.llmClient = {
        stream() {
            return (async function* () {
                yield { type: 'token', delta: JSON.stringify({ chartType: '饼图', title: '按部门销售额占比', sheet: '销售数据', dataRange: '销售数据!A1:I101' }) };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const { events, failed } = await runSkillPipeline(App, '使用饼图按部门展示销售额', {
        selection: { address: 'A1:I101', sheetName: '销售数据', rowCount: 101, colCount: 9 },
        _matchedSkill: App.cachedSkills[0]
    });

    assert.ok(!failed, '管道不应报错: ' + JSON.stringify(failed));
    assert.ok(events.includes('pipeline_done'), '应有 pipeline_done');
    assert.ok(aggArgs, '应调用 excel_aggregate');
    assert.strictEqual(aggArgs.dimColumn, '部门', '聚合维度应为「部门」');
    assert.strictEqual(aggArgs.valueColumn, '销售额', '聚合指标应为「销售额」');
    // 新契约（2026-10-09）：skill 管道走"零落盘内存绑定"——永不强制 targetSheet，
    // 落盘与否交由后端决定（与 chart-empty-source.test.js 用例 4 同契约）。
    // 图表层级的贯通才是本回归真正要守护的（见下方 chartArgs.sheet/dataRange 断言）。
    assert.strictEqual(aggArgs.targetSheet, undefined, 'skill 管道不应强制 targetSheet（零落盘契约，落盘由后端决定）');
    assert.ok(chartArgs, '应调用 excel_create_chart');
    // 【核心回归】图表 sheet 必须同步为聚合结果表，而不是 plan.sheet（原表 销售数据）
    assert.strictEqual(chartArgs.sheet, '聚合结果',
        `图表 sheet 必须贯通为聚合结果表（后端显式 sheet 优先于 dataRange 前缀，若仍是销售数据 → 读原表文本列报错），实际 ${chartArgs.sheet}`);
    assert.strictEqual(chartArgs.dataRange, '聚合结果!A1:B6',
        `图表数据范围必须指向聚合结果表，实际 ${chartArgs.dataRange}`);
    console.log('✅ 聚合→图表数据源贯通：sheet=' + chartArgs.sheet + ' dataRange=' + chartArgs.dataRange);
});

test('「对选中区域生成图表」：选区优先于 LLM 整表规划 + 纯文本频次聚合 sheet 贯通', async () => {
    const App = loadChatV2(ORDER).App;
    const tr = setupChartSkill(App);

    // 快照读取：真实选区是纯文本分类列（客户类型）
    tr.register({
        name: 'excel_read_range', displayName: '读取区域', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) {
            return {
                ok: true,
                result: {
                    headers: ['客户类型'],
                    result: [
                        { 客户类型: '企业' }, { 客户类型: '个人' }, { 客户类型: '政府客户' },
                        { 客户类型: '企业' }, { 客户类型: '个人' }, { 客户类型: '企业' }
                    ]
                }
            };
        }
    });
    let aggCalled = false;
    tr.register({
        name: 'excel_aggregate', displayName: '分组聚合', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute() { aggCalled = true; return { ok: true, result: {} }; }
    });
    // 频次聚合 → 写入「图表数据源」表
    let writeArgs = null;
    tr.register({
        name: 'excel_write_range', displayName: '写入区域', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) { writeArgs = args; return { ok: true, result: { success: true } }; }
    });
    let chartArgs = null;
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: 'd',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute(args) { chartArgs = args; return { ok: true, result: { chartName: 'Chart 1' } }; }
    });

    // LLM 规划：忽略真实选区，自拟整表范围（历史 bug 源）
    App.ChatV2.llmClient = {
        stream() {
            return (async function* () {
                yield { type: 'token', delta: JSON.stringify({ chartType: '饼图', title: '客户类型占比', sheet: '销售数据', dataRange: '销售数据!A1:I101' }) };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const { events, failed } = await runSkillPipeline(App, '对选中区域生成图表', {
        selection: { address: 'I79:I101', sheetName: '销售数据', rowCount: 23, colCount: 1 },
        _matchedSkill: App.cachedSkills[0]
    });

    assert.ok(!failed, '管道不应报错: ' + JSON.stringify(failed));
    assert.ok(events.includes('pipeline_done'), '应有 pipeline_done');
    assert.strictEqual(aggCalled, false, '纯文本频次路径不应调用 excel_aggregate');
    assert.ok(writeArgs, '应写入「图表数据源」表（频次统计）');
    assert.strictEqual(writeArgs.sheet, '图表数据源');
    assert.ok(chartArgs, '应调用 excel_create_chart');
    // 【核心回归】图表数据源必须指向频次聚合表，sheet 同步（3 类 → A1:B4）
    assert.strictEqual(chartArgs.dataRange, '图表数据源!A1:B4',
        `图表数据范围必须指向频次统计结果表，实际 ${chartArgs.dataRange}`);
    assert.strictEqual(chartArgs.sheet, '图表数据源',
        `图表 sheet 必须同步为图表数据源表，实际 ${chartArgs.sheet}`);
    console.log('✅ 选区优先 + 频次聚合贯通：sheet=' + chartArgs.sheet + ' dataRange=' + chartArgs.dataRange);
});

if (require.main === module) { setTimeout(() => process.exit(0), 5000).unref(); }
