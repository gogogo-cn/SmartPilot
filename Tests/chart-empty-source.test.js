/**
 * ============================================================
 * Tests/chart-empty-source.test.js — 图表空数据源防护测试
 * ============================================================
 * 背景（真实 bug）：选区只有"客户类型"这类纯文本分类列时，
 *   excel_create_chart 的 SetSourceData 不报错但饼图是空壳（只有标题没有扇区），
 *   工具却返回 ✅（假完成）。
 * 修复：
 *   1) 前端 skill-spec-execute：快照读取后检测纯文本选区 → 内存统计频次 →
 *      写入"图表数据源"表 → 在聚合结果上绘图（频次占比才是正确语义）
 *   2) 后端 ExecuteCreateChart：SetSourceData 后校验系列确有可绘数值，
 *      空壳图一律删除并明确报错（诚实失败，不做假完成）
 * 运行：node Tests/chart-empty-source.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');
// 与真实 index.html 一致的加载顺序，截断到 chat-orchestrator（不加载入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

// ── 模拟截图真实场景：选区 I79:I101（客户类型列，无表头行，23 个分类值）──
const RAW_VALUES = [
    '企业客户', '个人客户', '政府客户', '个人客户', '企业客户',
    '企业客户', '个人客户', '政府客户', '政府客户', '个人客户',
    '政府客户', '个人客户', '个人客户', '个人客户', '个人客户',
    '个人客户', '企业客户', '个人客户', '个人客户', '个人客户',
    '个人客户', '个人客户', '个人客户'
];
// hasHeaders=true 读取时：首行"企业客户"被当表头，其余 22 行作为数据；
// 伪表头修复把"企业客户"补回频次统计（企 3+1=4，个 15，政 4，共 23）
const EXPECTED_FREQ = [
    ['类别', '数量'],
    ['个人客户', 15],
    ['企业客户', 4],
    ['政府客户', 4]
];

test('skill-spec chart e2e: 纯文本选区 → 自动频次聚合后绘图（不再产出空壳饼图）', async () => {
    const App = loadChatV2(ORDER).App;

    // 1. 注册 SmartChartWizard（SkillSpec 模式）
    App.cachedSkills = [{
        name: 'SmartChartWizard', displayName: '智能图表向导',
        keywords: ['图表', '占比'], inputContext: ['selection'], pipeline: 'skill',
        processingMode: 'SkillSpec'
    }];

    const tr = App.ChatV2.toolRegistry;
    let writeArgs = null;
    let chartArgs = null;
    let readCount = 0;

    tr.register({
        name: 'excel_read_range', displayName: '读取数据', description: '读取区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            readCount++;
            // 模拟后端 hasHeaders=true 行为：首行"企业客户"成为表头，其余 22 行为 dict rows
            const headers = [RAW_VALUES[0]];
            const rows = RAW_VALUES.slice(1).map(v => ({ [headers[0]]: v }));
            return {
                ok: true,
                result: {
                    workbook: '测试.xlsx', worksheet: args?.sheet || 'Sheet1',
                    range: args?.range || 'I79:I101', rows: rows.length, columns: 1,
                    headers, result: rows
                }
            };
        }
    });
    tr.register({
        name: 'excel_write_range', displayName: '写入数据', description: '写入区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { writeArgs = args; return { ok: true, result: { success: true } }; }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    // 2. stub LLM：规划调用返回饼图方案（模拟"对选中区域生成图表"+LLM 规划 xlPie 占比分析）
    App.ChatV2.llmClient = {
        stream({ messages }) {
            return (async function* () {
                // 识别"规划调用"不能再用"规划"字样（图表规划 prompt 已改写，不含该词）。
// 改为识别图表 spec 模板特征字段 "chartType"，与 buildChartPlanningPrompt 契约绑定。
const isPlanning = messages.some(m => m.role === 'system' && String(m.content).includes('"chartType"'));
                if (isPlanning) {
                    yield { type: 'token', delta: JSON.stringify({ chartType: 'xlPie', title: '占比分析', showLegend: true }) };
                } else {
                    yield { type: 'token', delta: '好的。' };
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    // 3. 模拟 turn：选区 I79:I101（纯文本客户类型列）
    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '对选中区域生成图表',
        kind: 'text',
        context: { selection: { address: 'I79:I101', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }
    });

    // 4. 运行 skill-pipeline
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    const events = [];
    for await (const evt of gen) {
        events.push(evt.type);
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    // 5. 断言：频次聚合写入 + 图表落在聚合结果上
    assert.ok(writeArgs, '应调用 excel_write_range 写入频次聚合结果');
    assert.strictEqual(writeArgs.sheet, '图表数据源', '应写入"图表数据源"表（自动新建）');
    assert.deepStrictEqual(writeArgs.values, EXPECTED_FREQ,
        `频次统计应为 ${JSON.stringify(EXPECTED_FREQ)}，实际: ${JSON.stringify(writeArgs.values)}`);
    assert.ok(chartArgs, '应调用 excel_create_chart');
    assert.strictEqual(chartArgs.chartType, 'xlPie');
    assert.strictEqual(chartArgs.dataRange, '图表数据源!A1:B4', '图表应绘制在频次聚合结果上（3 类 + 表头 = B4）');
    assert.strictEqual(chartArgs.title, '占比分析');
    assert.ok(events.includes('pipeline_done'), 'pipeline 应正常完成');
    assert.ok(readCount >= 1, '应先读取选区数据做数值性判定');
    console.log('\n✅ 纯文本选区 → 频次聚合 → 图表数据源!A1:B4 绘图，全链路通过');
});

test('VB 源码: ExecuteCreateChart 含空数据源诚实校验（校验系列数值 + 删除空壳 + 明确报错）', () => {
    const vb = fs.readFileSync(path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.Core.Chart.vb'), 'utf8');

    assert.ok(vb.includes('诚实绘图校验'), '应包含诚实绘图校验注释块');
    assert.ok(/SeriesCollection/.test(vb), '应检查 SeriesCollection');
    assert.ok(/sChk\.Values/.test(vb), '应逐系列检查 Values 是否含可绘数值');
    assert.ok(/chartObj\.Delete\(\)/.test(vb), '空壳图应被删除（不留垃圾图表框）');
    assert.ok(vb.includes('没有可绘制的数值列'), '应有明确的"无数值列"报错文案');
    // 豁免旭日/树状（可按类别层级自聚合渲染）
    assert.ok(/<> 117 AndAlso CInt\(targetType\) <> 115/.test(vb), '应豁免 xlSunburst(117)/xlTreemap(115)');
    // 校验必须发生在 SetSourceData 之后（绑定数据源后才可能判定空壳）
    const setIdx = vb.indexOf('SetSourceData');
    const chkIdx = vb.indexOf('诚实绘图校验');
    assert.ok(setIdx > -1 && chkIdx > setIdx, '校验应在 SetSourceData 之后执行');
    console.log('  [ok] 后端诚实校验：空数据源拒绝绘图并清理空壳');
});

// ── 真实失败场景②（22:27）：用户选 I 列说"对选中区域生成图表"，LLM 却自拟整表
//    $A$1:$I$101 → 整表数值列掩盖纯文本判定 → 整表画饼图 → Excel 取第一列（日期
//    文本）当系列值 → 空壳图。修复：用户指明"选中区域"时选区优先于 plan.dataRange。
test('skill-spec chart e2e: 用户指明选中区域时，选区优先于 LLM 自拟的整表 dataRange', async () => {
    const App = loadChatV2(ORDER).App;
    App.cachedSkills = [{
        name: 'SmartChartWizard', displayName: '智能图表向导',
        keywords: ['图表', '占比'], inputContext: ['selection'], pipeline: 'skill',
        processingMode: 'SkillSpec'
    }];
    const tr = App.ChatV2.toolRegistry;
    let writeArgs = null;
    let chartArgs = null;
    tr.register({
        name: 'excel_read_range', displayName: '读取数据', description: '读取区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            // 无论读什么都返回 I 列纯文本数据（stub 语义：模拟选区快照）
            const headers = [RAW_VALUES[0]];
            const rows = RAW_VALUES.slice(1).map(v => ({ [headers[0]]: v }));
            return {
                ok: true,
                result: {
                    workbook: '测试.xlsx', worksheet: args?.sheet || 'Sheet1',
                    range: args?.range || 'I79:I101', rows: rows.length, columns: 1,
                    headers, result: rows
                }
            };
        }
    });
    tr.register({
        name: 'excel_write_range', displayName: '写入数据', description: '写入区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { writeArgs = args; return { ok: true, result: { success: true } }; }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    // LLM 规划：无视 I 列选区、按上下文整表自拟范围（真实失败样本行为）
    App.ChatV2.llmClient = {
        stream({ messages }) {
            return (async function* () {
                // 识别"规划调用"不能再用"规划"字样（图表规划 prompt 已改写，不含该词）。
// 改为识别图表 spec 模板特征字段 "chartType"，与 buildChartPlanningPrompt 契约绑定。
const isPlanning = messages.some(m => m.role === 'system' && String(m.content).includes('"chartType"'));
                if (isPlanning) {
                    yield { type: 'token', delta: JSON.stringify({ chartType: 'xlPie', title: '占比分析', dataRange: 'Sheet1!$A$1:$I$101', showLegend: true }) };
                } else {
                    yield { type: 'token', delta: '好的。' };
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '对选中区域生成图表',
        kind: 'text',
        context: { selection: { address: 'I79:I101', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }
    });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    for await (const evt of gen) {
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    assert.ok(writeArgs, '纯文本选区应触发频次聚合写入（而不是按整表画图）');
    assert.strictEqual(writeArgs.sheet, '图表数据源');
    assert.ok(chartArgs, '应调用 excel_create_chart');
    assert.strictEqual(chartArgs.dataRange, '图表数据源!A1:B4',
        '图表应绘制在选区频次聚合结果上，而不是 LLM 自拟的整表范围');
    console.log('\n✅ 选区优先：I79:I101（纯文本）→ 频次聚合 → 图表数据源!A1:B4，未被 $A$1:$I$101 覆盖');
});

// ── 真实失败场景③（22:28）："使用饼图按部门展示销售额" → 聚合步骤只做内存计算
//    （纯聚合不落盘），前端旧逻辑误判"旧后端已内置画图"跳过 excel_create_chart /
//    或图表步骤缺 dataRange 报"缺少必要参数"。修复：聚合传 targetSheet 落盘 +
//    按 writtenTo/groupCount 构造确定性 dataRange。
test('skill-spec chart e2e: 聚合+饼图 → 聚合结果落盘（targetSheet）后在其上绘图（不再漏画/缺参）', async () => {
    const App = loadChatV2(ORDER).App;
    App.cachedSkills = [{
        name: 'SmartChartWizard', displayName: '智能图表向导',
        keywords: ['饼图', '图表', '汇总'], inputContext: ['selection'], pipeline: 'skill',
        processingMode: 'SkillSpec'
    }];
    const tr = App.ChatV2.toolRegistry;
    let aggArgs = null;
    let chartArgs = null;
    const AGG_DATA = [
        { 部门: '销售一部', 销售额: 100 },
        { 部门: '销售二部', 销售额: 220 },
        { 部门: '销售三部', 销售额: 180 }
    ];
    tr.register({
        name: 'excel_read_range', displayName: '读取数据', description: '读取区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            // 整表快照：含数值列（与真实场景一致，不应触发纯文本频次聚合）
            const headers = ['部门', '销售额'];
            const rows = [
                { 部门: '销售一部', 销售额: 100 },
                { 部门: '销售二部', 销售额: 220 }
            ];
            return {
                ok: true,
                result: { workbook: '测试.xlsx', worksheet: args?.sheet || 'Sheet1', range: args?.range || 'A1:B3', rows: rows.length, columns: 2, headers, result: rows }
            };
        }
    });
    tr.register({
        name: 'excel_aggregate', displayName: '聚合分析', description: '分组聚合', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            aggArgs = args;
            // 模拟（已重建的）后端：targetSheet 落盘后返回 writtenTo + groupCount
            return {
                ok: true,
                result: JSON.stringify({
                    success: true, aggregated: true, dimColumn: '部门', valueColumn: '销售额',
                    aggFunc: 'sum', groupCount: AGG_DATA.length, writtenTo: '聚合结果', data: AGG_DATA
                })
            };
        }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    App.ChatV2.llmClient = {
        stream({ messages }) {
            return (async function* () {
                // 识别"规划调用"不能再用"规划"字样（图表规划 prompt 已改写，不含该词）。
// 改为识别图表 spec 模板特征字段 "chartType"，与 buildChartPlanningPrompt 契约绑定。
const isPlanning = messages.some(m => m.role === 'system' && String(m.content).includes('"chartType"'));
                if (isPlanning) {
                    yield { type: 'token', delta: JSON.stringify({ chartType: 'xlPie', title: '部门销售额占比', dataRange: 'Sheet1!$A$1:$B$3', showLegend: true }) };
                } else {
                    yield { type: 'token', delta: '好的。' };
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '按部门汇总销售额，用饼图展示',
        kind: 'text',
        context: { selection: { address: 'A1:B3', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }
    });
    const gen = App.ChatV2.pipelines.skill.execute(turn, new AbortController().signal);
    for await (const evt of gen) {
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    assert.ok(aggArgs, '应调用 excel_aggregate');
    // 契约变更（2026-10-09）：skill 管道的聚合改为**零落盘内存绑定** ——
    // 不传 targetSheet（后端纯内存聚合），返回的 data 由前端构造成 dataValues 直接绑定系列，
    // 彻底消除"聚合结果/图表数据源"残留工作表（见 skill-spec-execute.js:806-878）。
    // 落盘契约仍由 agent-loop 承担（同名 targetSheet 强制覆盖，见 agent-loop.test.js）。
    assert.strictEqual(aggArgs.targetSheet, undefined, 'skill 管道应走零落盘内存绑定（不传 targetSheet）');
    assert.strictEqual(aggArgs.dimColumn, '部门');
    assert.strictEqual(aggArgs.valueColumn, '销售额');
    assert.ok(chartArgs, '聚合后必须创建图表（旧逻辑会误判"旧后端已内置画图"而跳过）');
    assert.strictEqual(chartArgs.chartType, 'xlPie');
    const dv = (() => { try { return JSON.parse(chartArgs.dataValues); } catch { return null; } })();
    assert.deepStrictEqual(dv, [
        ['部门', '销售额'],
        ['销售一部', 100], ['销售二部', 220], ['销售三部', 180]
    ], `图表数据应来自聚合结果（dataValues 内存绑定），实际: ${chartArgs.dataValues}`);
    console.log('\n✅ 聚合（零落盘）→ dataValues 内存绑定饼图，全链路确定性、无残留工作表');
});

// ── 真实失败场景④（22:28 多步骤路径）：LLM 规划 3 步（读取→聚合→饼图），
//    聚合步骤只内存计算未落盘，图表步骤缺 dataRange → needsInput 硬失败
//    "缺少必要参数：数据范围"。修复：StepExecutor 图表步骤发现前置聚合仅有
//    内存结果时，兜底落盘为新表并引用。
test('StepExecutor: 图表步骤发现前置聚合仅有内存结果 → 兜底落盘后引用（不再缺 dataRange 硬失败）', async () => {
    const ORDER_SE = ORDER.concat(['services/StepExecutor.js']);
    const App = loadChatV2(ORDER_SE).App;
    const tr = App.ChatV2.toolRegistry;
    let landWrite = null;    // 兜底落盘写入
    let chartArgs = null;
    tr.register({
        name: 'excel_write_range', displayName: '写入数据', description: '写入区域数据', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { landWrite = args; return { ok: true, result: { success: true } }; }
    });
    tr.register({
        name: 'excel_create_chart', displayName: '创建图表', description: '创建图表', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) { chartArgs = args; return { ok: true, result: { success: true, name: '图表 1' } }; }
    });

    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '使用饼图按部门展示销售额', kind: 'text', context: {}
    });

    // 前置聚合步骤的内存结果（chartType 规划在图表步上 → 聚合步未落盘，无 resultSheet）
    const prevAggResult = {
        ok: true, error: null,
        result: {
            success: true, aggregated: true, dimColumn: '部门', valueColumn: '销售额',
            aggFunc: 'sum', groupCount: 3, headers: ['部门', '销售额'],
            data: [
                { 部门: '销售一部', 销售额: 100 },
                { 部门: '销售二部', 销售额: 220 },
                { 部门: '销售三部', 销售额: 180 }
            ]
        }
    };
    const step = {
        id: 'step-3',
        tool: 'excel_create_chart', toolName: 'excel_create_chart',
        strategy: 'direct',
        description: '创建按部门展示销售额的饼图',
        params: { chartType: 'xlPie', title: '部门销售额占比' }   // 缺 dataRange/sheet（真实失败样本）
    };
    const stepResult = await App.ChatV2.StepExecutor.executeStep(step, 2, {
        turn, sessionStore, toolRegistry: tr,
        results: [prevAggResult],
        stepContext: {},
        enableClientAggregate: false
    });

    assert.ok(stepResult.ok, `图表步骤应执行成功（不再 needsInput 硬失败）: ${stepResult.error || ''}`);
    assert.ok(landWrite, '应先兜底落盘聚合结果');
    // 表名规则：聚合_{维度}_{步骤索引}_{时间戳后 4 位}（StepExecutor.js:209），
    // 带步骤索引 + 时间戳是为了"同一轮多张图各自的落盘表互不覆盖"。
    assert.ok(/^聚合_部门_\d+_\d{1,5}$/.test(landWrite.sheet),
        `落盘表名应为 聚合_部门_{步骤}_{时间戳}，实际: ${landWrite.sheet}`);
    const written = typeof landWrite.values === 'string' ? JSON.parse(landWrite.values) : landWrite.values;
    assert.deepStrictEqual(written, [
        ['部门', '销售额'],
        ['销售一部', 100],
        ['销售二部', 220],
        ['销售三部', 180]
    ], `落盘内容应为聚合结果二维表，实际: ${JSON.stringify(written)}`);
    assert.ok(chartArgs, '应调用 excel_create_chart');
    assert.strictEqual(chartArgs.dataRange, 'A1:B4', '图表数据范围应指向落盘表 A1:B4（表头 + 3 组）');
    assert.ok(chartArgs.sheet && chartArgs.sheet.startsWith('聚合_部门_'), '图表应建在落盘表上');
    console.log('\n✅ StepExecutor 兜底落盘：聚合内存结果 → 聚合_部门_N!A1:B4 → 饼图，不再"缺少必要参数"');
});

// ── 契约归属护栏（2026-10-09）──
// 历史：text-pipeline 曾内置"聚合+图表"路径（__forceTextPipeline），旧后端不落盘时
//       前端还会补写"图表数据源"表。现在这条路径已从 text-pipeline 移除：
//         · 落盘契约   → agent-loop（excel_aggregate 强制补 targetSheet，图表强制对齐落盘表）
//         · 零落盘契约 → skill 管道（dataValues 内存绑定，见用例 4）
// 原用例 6/7 的 e2e 因此退役，改为这条静态护栏：防止同一契约在两处重复实现而漂移。
test('契约归属: 聚合→图表 只由 agent-loop（落盘）与 skill 管道（内存绑定）实现，text-pipeline 不再内置', () => {
    const tp = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'text-pipeline.js'), 'utf8');
    const al = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'agent-loop.js'), 'utf8');
    const se = fs.readFileSync(path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill', 'skill-spec-execute.js'), 'utf8');

    assert.ok(!tp.includes('excel_aggregate'), 'text-pipeline 不应再内置聚合调用（契约已迁出，避免两处实现漂移）');
    assert.ok(/targetSheet:\s*ts/.test(al), 'agent-loop 应强制覆盖 excel_aggregate 的 targetSheet（落盘）');
    assert.ok(/pickedAgg\.dataRange/.test(al), 'agent-loop 应把图表 sheet/dataRange 对齐到聚合结果表');
    assert.ok(/dataValues/.test(se), 'skill 管道应以 dataValues 内存绑定聚合结果（零落盘）');
    console.log('  [ok] 聚合→图表 契约归属单一：agent-loop(落盘) / skill 管道(内存绑定)');
});

// ── 规划层参数完整性：buildSingleAggregateTask 生成的聚合任务必须携带 targetSheet，
//    供 text-pipeline / prompt 构造等所有下游消费方复用（防止某条路径再次漏传落盘参数）。
test('task-planner buildSingleAggregateTask: 聚合任务参数含 targetSheet（落盘"聚合结果"）', () => {
    const App = loadChatV2(ORDER).App;
    const task = App.ChatV2.taskPlanner.buildSingleAggregateTask('按部门汇总销售额');
    assert.ok(task, '应构建出聚合任务');
    assert.strictEqual(task.tool, 'excel_aggregate');
    assert.ok(task.params, '聚合任务应有参数');
    assert.strictEqual(task.params.targetSheet, '聚合结果', '聚合参数必须显式落盘，供后续图表引用');
    const sysPrompt = App.ChatV2.taskPlanner.buildAggregateSystemPrompt(task);
    assert.ok(sysPrompt.includes('targetSheet="聚合结果"'), 'System Prompt 应携带 targetSheet 落盘要求');
    console.log('\n✅ task-planner 聚合任务参数含 targetSheet，prompt 同步携带落盘要求');
});
