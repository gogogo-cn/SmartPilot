/**
 * Tests/dashboard-audit.test.js — 看板生成能力审核回归
 * 覆盖：一键工具 dimension/value 字段、交付物意图、图表别名兼容
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const DASH_ORDER = [
    ...DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('services/skill-spec-engine.js') + 1),
    'services/dashboard-html-renderer.js',
    'services/dashboard-excel-renderer.js',
    'services/dashboard-blueprint-engine.js',
    'services/dashboard-one-click.js',
    ...DEFAULT_ORDER.slice(DEFAULT_ORDER.indexOf('services/skill-spec-engine.js') + 1)
];

function loadDash() {
    return loadChatV2(DASH_ORDER).App;
}

test('dashboardSpecFromMessage: Excel 看板 → deliverable=excel', () => {
    const App = loadDash();
    const eng = App.ChatV2.services.skillSpecEngine;
    assert.ok(eng && eng.dashboardSpecFromMessage, 'skillSpecEngine 应导出 dashboardSpecFromMessage');
    const s1 = eng.dashboardSpecFromMessage('做个 Excel 看板');
    assert.strictEqual(s1.deliverable, 'excel', 'Excel 看板应仅出 Excel');
    const s2 = eng.dashboardSpecFromMessage('生成表格看板放到工作表');
    assert.strictEqual(s2.deliverable, 'excel');
    const s3 = eng.dashboardSpecFromMessage('只要网页大屏投屏');
    assert.strictEqual(s3.deliverable, 'html');
    const s4 = eng.dashboardSpecFromMessage('做个经营驾驶舱');
    assert.strictEqual(s4.deliverable, 'dual', '模糊看板默认 dual');
});

test('aggregateChartData: 兼容 dim/vals 别名（与 dimension/value 等价）', () => {
    const App = loadDash();
    const eng = App.ChatV2.dashboardModules.blueprintEngine;
    assert.ok(eng && eng.aggregateChartData, 'blueprintEngine.aggregateChartData');
    const source = {
        id: 's1',
        headers: ['部门', '销售额'],
        rows: [
            { '部门': 'A', '销售额': 100 },
            { '部门': 'B', '销售额': 200 },
            { '部门': 'A', '销售额': 50 }
        ]
    };
    const a = eng.aggregateChartData({ dimension: '部门', value: '销售额', func: 'sum', chartType: 'xlPie' }, source);
    const b = eng.aggregateChartData({ dim: '部门', vals: ['销售额'], func: 'sum', chartType: 'xlPie' }, source);
    assert.ok(a && a.data && a.data.length >= 2, 'dimension/value 应产出分组');
    assert.ok(b && b.data && b.data.length >= 2, 'dim/vals 别名应产出分组');
    assert.strictEqual(a.dim, b.dim);
    assert.deepStrictEqual(a.data.sort(), b.data.sort());
});

test('one-click dims 映射为 dimension/value（不再写 dim/vals）', () => {
    // 静态读源文件断言字段形态，避免依赖 Excel 桥
    const fs = require('fs');
    const path = require('path');
    const src = fs.readFileSync(path.join(__dirname, '../UI/js/chat-v2/services/dashboard-one-click.js'), 'utf8');
    assert.ok(/dimension:\s*dim/.test(src), 'dims 映射应写 dimension');
    assert.ok(/value:\s*val/.test(src), '指标应写 value');
    assert.ok(!/spec\.charts\s*=\s*a\.dims[\s\S]{0,200}\bdim,\s*\n\s*vals:/.test(src),
        '不应再使用 dim/vals 旧字段形态');
});

test('filterToolsForMessage: 看板消息保留 excel_generate_dashboard', () => {
    const App = loadDash();
    const util = App.ChatV2.toolRegistryUtil;
    const kept = util.filterToolsForMessage(
        [
            { name: 'excel_read_range' },
            { name: 'excel_generate_dashboard' },
            { name: 'excel_create_chart' },
            { name: 'excel_add_shape' }
        ],
        '帮我做个数据看板'
    ).map(t => t.name);
    assert.ok(kept.includes('excel_generate_dashboard'), '实际=' + kept.join(','));
    assert.ok(kept.includes('excel_create_chart'));
});

test('one-click: 销售数据蓝图含 excel_create_slicer 步骤', () => {
    const App = loadDash();
    const engine = App.ChatV2.dashboardModules.blueprintEngine;
    const excel = App.ChatV2.dashboardModules.excelRenderer;
    assert.ok(engine && excel && excel.buildExcelPlan, 'excelRenderer.buildExcelPlan 应可用');
    const headers = ['日期', '地区', '部门', '产品', '销售额', '成本', '利润', '销售员', '客户类型'];
    const rows = [];
    const regions = ['华南', '华北', '华东'];
    const depts = ['销售一部', '销售二部'];
    const products = ['产品A', '产品B'];
    const customers = ['政府客户', '企业客户'];
    for (let i = 0; i < 24; i++) {
        rows.push({
            '日期': '2024-01-' + String((i % 28) + 1).padStart(2, '0'),
            '地区': regions[i % regions.length],
            '部门': depts[i % depts.length],
            '产品': products[i % products.length],
            '销售额': 1000 + i * 100,
            '成本': 600 + i * 40,
            '利润': 400 + i * 60,
            '销售员': '销售' + (i % 5),
            '客户类型': customers[i % customers.length]
        });
    }
    const prepared = engine.computeFromSources([{
        id: 's1', sheet: '销售数据', range: 'A1:I25', headers, rows
    }], { title: '销售看板', sheet: '看板', deliverable: 'excel' });
    const steps = excel.buildExcelPlan(prepared, { title: '销售看板', sheet: '看板', deliverable: 'excel' });
    const slicerStep = steps.find(s => s && s.tool === 'excel_create_slicer');
    assert.ok(slicerStep, '应生成 excel_create_slicer 步骤；实际工具=' + steps.map(s => s.tool).join(','));
    assert.ok(slicerStep.args && slicerStep.args.sourceRange, '切片器需 sourceRange');
    assert.ok(String(slicerStep.args.sourceRange).includes('销售数据'), 'sourceRange 应含工作表名');
    assert.ok(slicerStep.args.slicerDimensions, '应有 slicerDimensions');
    assert.ok(slicerStep.args.pivotTables, '应有 pivotTables（透视图联动）');
});

test('one-click generateDashboard: 真正失败时降级出图', async () => {
    const App = loadDash();
    const oc = App.ChatV2.dashboardModules.oneClick;
    assert.ok(oc && oc.generateDashboard);
    if (typeof oc.registerTool === 'function') oc.registerTool();

    const calls = [];
    const tr = App.ChatV2.toolRegistry;
    tr.invokeTracked = async function (_turn, name, args) {
        calls.push({ name, args });
        if (name === 'excel_read_range') {
            return {
                ok: true,
                result: {
                    worksheet: '销售数据',
                    range: 'A1:I4',
                    headers: ['地区', '部门', '销售额'],
                    result: [
                        ['华南', '销售一部', 100],
                        ['华北', '销售二部', 200],
                        ['华东', '销售一部', 150]
                    ]
                }
            };
        }
        if (name === 'excel_create_slicer') {
            return { ok: false, error: 'slicer stub fail' };
        }
        return { ok: true, result: {} };
    };
    tr.invoke = tr.invokeTracked;

    const res = await oc.generateDashboard({
        sourceSheet: '销售数据',
        sourceRange: 'A1:I4',
        board: '看板',
        title: '销售看板'
    }, { turn: { id: 't1' } });

    assert.ok(res && res.ok, '降级后整体应成功: ' + JSON.stringify(res && res.error));
    assert.ok(calls.some(c => c.name === 'excel_create_slicer'), '应尝试创建切片器');
    assert.ok(calls.some(c => c.name === 'excel_create_chart'), '切片失败后应降级 create_chart');
    assert.ok(res.result && /切片失败|降级|原生切片/.test(String(res.result.message || '')),
        '消息应提及切片状态: ' + (res.result && res.result.message));
});

test('one-click: 上层误判失败但 payload 成功 → 不叠静态图', async () => {
    const App = loadDash();
    const oc = App.ChatV2.dashboardModules.oneClick;
    const calls = [];
    const tr = App.ChatV2.toolRegistry;
    tr.invokeTracked = async function (_turn, name) {
        calls.push(name);
        if (name === 'excel_read_range') {
            return {
                ok: true,
                result: {
                    worksheet: '销售数据', range: 'A1:C4',
                    headers: ['地区', '部门', '销售额'],
                    result: [['华南', '一部', 100], ['华北', '二部', 200], ['华东', '一部', 150]]
                }
            };
        }
        if (name === 'excel_create_slicer') {
            // 模拟：tool-registry 因 FALLBACK 文案误判 ok=false，但文末 payload 显示全成功
            return {
                ok: false,
                error: '=== excel_create_slicer 执行汇总 ===\n' +
                    'PivotTable: 2 张 | PivotChart: 2/2 成功 | Slicer: 2/2\n' +
                    '── Warnings ──\n[SLICER_DIAG] 8-arg FALLBACK: 参数错误\n[SLICER_DIAG] SUCCESS (no-Level)\n' +
                    '=== Payload JSON ===\n' +
                    JSON.stringify({ slicersCreated: 2, ptChartOkCount: 2, slicers: [{ header: '地区' }, { header: '部门' }] })
            };
        }
        return { ok: true, result: {} };
    };
    tr.invoke = tr.invokeTracked;

    const res = await oc.generateDashboard({
        sourceSheet: '销售数据', sourceRange: 'A1:C4', board: '看板', title: '销售看板'
    }, { turn: { id: 't2' } });

    assert.ok(res && res.ok, '应整体成功: ' + JSON.stringify(res && res.error));
    assert.strictEqual(res.result.slicers, 'ok');
    assert.ok(!calls.includes('excel_create_chart'), 'payload 已含透视图时不得再叠静态图，实际=' + calls.filter(n => n === 'excel_create_chart').length);
    assert.ok(/原生切片器联动/.test(String(res.result.message || '')), '消息应为联动成功: ' + res.result.message);
});

test('resolveThemeFromMessage / isThemeChangeIntent: 红/蓝/紫', () => {
    const App = loadDash();
    const eng = App.ChatV2.dashboardModules.blueprintEngine;
    assert.strictEqual(eng.resolveThemeFromMessage('修改为红色主题'), 'crimson_red');
    assert.strictEqual(eng.resolveThemeFromMessage('改为蓝色主题'), 'business_blue');
    assert.strictEqual(eng.resolveThemeFromMessage('改为紫色主题'), 'violet_purple');
    assert.strictEqual(eng.resolveThemeFromMessage('换成商务蓝'), 'business_blue');
    assert.ok(eng.isThemeChangeIntent('修改为红色主题'));
    assert.ok(eng.isThemeChangeIntent('改为紫色主题'));
    assert.ok(!eng.isThemeChangeIntent('生成Excel看板'));
    assert.ok(eng.THEMES.crimson_red, '应注册中国红主题');
    assert.ok(eng.THEMES.violet_purple, '应注册紫罗兰主题');
});

test('text-pipeline: 修改为红色主题 → 确定性换主题不空转', async () => {
    const App = loadDash();
    const oc = App.ChatV2.dashboardModules.oneClick;
    if (oc && typeof oc.registerTool === 'function') oc.registerTool();
    App.ChatV2._lastDashboardSource = { sheet: '销售数据', range: 'A1:I101', board: '看板' };

    let themeArgs = null;
    let htmlCalled = 0;
    const tr = App.ChatV2.toolRegistry;
    const origChange = oc.changeDashboardTheme.bind(oc);
    oc.changeDashboardTheme = async function (args, ctx) {
        themeArgs = args;
        return {
            ok: true,
            result: { board: '看板', theme: args.theme, message: '已将看板切换为「中国红」主题' }
        };
    };
    tr.register({
        name: 'excel_display_html_in_browser',
        displayName: 'HTML大屏',
        source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { htmlCalled++; return { ok: true, result: {} }; }
    });
    tr.invokeTracked = async function (turn, name, args, opts) {
        if (name === 'excel_display_html_in_browser') {
            htmlCalled++;
            return { ok: true, result: {} };
        }
        return { ok: true, result: {}, durationMs: 1 };
    };
    App.ChatV2.llmClient = {
        stream() {
            return (async function* () {
                yield { type: 'token', delta: '不应该走到这里' };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const turn = App.ChatV2.sessionStore.startTurn({
        message: '修改为红色主题',
        kind: 'text',
        context: { selection: { address: 'A1:I20', sheetName: '看板' } }
    });
    const events = [];
    for await (const evt of App.ChatV2.pipelines.text.execute(turn, new AbortController().signal)) {
        events.push(evt.type);
    }

    assert.ok(themeArgs, '应调用 changeDashboardTheme');
    assert.strictEqual(themeArgs.theme, 'crimson_red');
    assert.strictEqual(themeArgs.sourceSheet, '销售数据');
    assert.strictEqual(htmlCalled, 0);
    assert.ok(String(turn.content || '').includes('中国红') || String(turn.content || '').includes('主题'),
        '正文应说明主题已切换: ' + turn.content);
    assert.ok(!String(turn.content || '').includes('不应该走到这里'));
    assert.ok(events.includes('pipeline_done'));
    oc.changeDashboardTheme = origChange;
});

test('text-pipeline: 生成Excel看板 → 一键施工且不弹 HTML', async () => {
    const App = loadDash();
    // 确保 one-click 已注册
    const oc = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.oneClick;
    if (oc && typeof oc.registerTool === 'function') oc.registerTool();

    let oneClickArgs = null;
    let htmlCalled = 0;
    const tr = App.ChatV2.toolRegistry;
    // 桩：一键工具成功；display_html 若被调则计数
    tr.register({
        name: 'excel_generate_dashboard',
        displayName: '一键生成看板',
        source: 'builtin-dashboard',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            oneClickArgs = args;
            return {
                ok: true,
                result: { board: '看板', message: '已生成看板工作表「看板」：3 个 KPI + 2 张图表', charts: [{ ok: true }, { ok: true }] }
            };
        }
    });
    tr.register({
        name: 'excel_display_html_in_browser',
        displayName: 'HTML大屏',
        source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() { htmlCalled++; return { ok: true, result: { filePath: 'C:\\tmp\\x.html' } }; }
    });
    // 一键内部还会调 read/create 等——用通用桩吞掉未知工具
    const origInvoke = tr.invokeTracked.bind(tr);
    tr.invokeTracked = async function (turn, name, args, opts) {
        if (name === 'excel_generate_dashboard' || name === 'excel_display_html_in_browser') {
            return origInvoke(turn, name, args, opts);
        }
        return { ok: true, result: {}, toolCallId: null, durationMs: 1 };
    };

    App.ChatV2.llmClient = {
        stream() {
            return (async function* () {
                // 若仍走进 agent-loop，说明一键短路失败
                yield { type: 'token', delta: '不应该走到这里' };
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '生成Excel看板',
        kind: 'text',
        context: {
            _forceDashboard: true,
            _forceChart: true,
            _forceMultiStep: true,
            selection: { address: 'A1:I101', sheetName: '销售数据' }
        }
    });

    const pipeline = App.ChatV2.pipelines.text;
    const events = [];
    for await (const evt of pipeline.execute(turn, new AbortController().signal)) {
        events.push(evt.type);
    }

    assert.ok(oneClickArgs, '应确定性调用 excel_generate_dashboard');
    assert.strictEqual(htmlCalled, 0, 'Excel 看板意图不得再开 HTML 大屏');
    assert.ok(String(turn.content || '').includes('看板'), '正文应说明看板已生成');
    assert.ok(!String(turn.content || '').includes('不应该走到这里'), '不应落入 agent-loop 空转');
    assert.ok(events.includes('pipeline_done'));
});

if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); }
