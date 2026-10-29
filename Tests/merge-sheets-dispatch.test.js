/**
 * ============================================================
 * Tests/merge-sheets-dispatch.test.js — "合并工作表"必须真的合并，不得被演绎成画图
 * ============================================================
 * 背景（实测 2026-10-07 19:26，用户连续 4 次同一句）：
 *   用户："将选中工作表进行数据合并"（标签栏已多选 5 张表）
 *   实际：excel_read_range(大客户部) → excel_aggregate(sum) → excel_create_chart(饼图)，
 *         合并工具一次都没调用；界面最终写"成功创建图表…"。
 *
 * 根因两层（本测试各锁一条）：
 *   ① McpDispatcher 只有 清洗/筛选/检查/格式化 硬短路，没有"合并工作表" →
 *      请求整条交给 LLM，被自由发挥成"读数→聚合→画图"。
 *   ② 工具瘦身 filterToolsForMessage 按"合并"只保留了 merge_cells 族
 *      （词表里 merge_sheets 只挂在"多表/拆分/横向合并/纵向合并/join"上），
 *      用户最自然的语序"…工作表…数据合并"一个词都命中不了 →
 *      excel_merge_sheets 被剔除，模型手上根本没有合并工具。
 * 另外写类裁决闸门（EXEC_COMMAND_PATTERNS）也补了该语序，避免命中后又被 LLM 判成 other 放行。
 *
 * 运行：node --test Tests/merge-sheets-dispatch.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 需要 orchestrator（写类裁决闸门）+ mcp-dispatcher（硬短路），
// 所以用默认链路（已含 config/routing-constants.js、chat-orchestrator.js）再追加 dispatcher。
const { App } = loadChatV2(DEFAULT_ORDER.concat(['services/mcp-dispatcher.js']));
const dispatcher = App.ChatV2.mcpDispatcher;
const util = App.ChatV2.toolRegistryUtil;
const orch = App.ChatV2.orchestrator;

const opts = (message, ctx) => ({
    message, toolRegistry: null, sessionStore: null,
    turn: { context: ctx || {} },
    allowHardDispatch: true
});

// ══ ① dispatcher：合并工作表 → 确定性直调 excel_merge_sheets ══
test('"将选中工作表进行数据合并" → 硬短路 excel_merge_sheets（不再交 LLM 自由发挥）', async () => {
    const d = await dispatcher.tryDispatch(opts('将选中工作表进行数据合并', {
        selectedSheets: ['大客户部', '销售三部', '市场部', '销售二部', '销售一部']
    }));
    assert.strictEqual(d.hardMatched, true, '应走确定性硬短路');
    assert.strictEqual(d.toolName, 'excel_merge_sheets');
    assert.deepStrictEqual(JSON.parse(d.args.sourceSheets),
        ['大客户部', '销售三部', '市场部', '销售二部', '销售一部']);
    assert.notStrictEqual(d.args.mergeAll, true, '用户选中的是这几张表，不应强制合并全部可见表');
});

test('线上默认：Excel 任务不硬短路，交给模型规划', async () => {
    for (const msg of ['将选中工作表进行数据合并，并设置成中国红', '美化工作簿中的所有工作表，统一为商务蓝', '按部门拆分成不同Sheet']) {
        const d = await dispatcher.tryDispatch({
            message: msg, toolRegistry: null, sessionStore: null, turn: { context: {} }
        });
        assert.strictEqual(d.hit, false, msg);
        assert.strictEqual(d.deferToPlanner, true, msg);
        assert.strictEqual(d.reason, 'llm-plan', msg);
        if (/中国红/.test(msg)) assert.ok(/中国红/.test(d.styleTarget || ''), msg + ' ' + d.styleTarget);
    }
});

test('"将选中工作表进行数据合并，并设置成中国红" → 不硬短路，颜色交给模型', async () => {
    const d = await dispatcher.tryDispatch(opts('将选中工作表进行数据合并，并设置成中国红'));
    assert.strictEqual(d.hit, false, '带中国红的合并不能只合并就结束');
    assert.strictEqual(d.deferToPlanner, true);
    assert.ok(/中国红/.test(d.styleTarget || ''), d.styleTarget);
    assert.notStrictEqual(d.toolName, 'excel_merge_sheets');
});

test('"把这几张表合并成一张" 同样硬短路（动作词在句尾的自然语序）', async () => {
    const d = await dispatcher.tryDispatch(opts('把这几张表合并成一张'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_merge_sheets');
});

test('"合并所有工作表" → mergeAll=true（忽略标签选中态）', async () => {
    const d = await dispatcher.tryDispatch(opts('合并所有工作表到一个新表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_merge_sheets');
    assert.strictEqual(d.args.mergeAll, true);
});

test('"合并到「汇总数据」表" → 带引号的落表名透传为 targetSheet', async () => {
    const d = await dispatcher.tryDispatch(opts('把选中的工作表合并到「汇总数据」表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.args.targetSheet, '汇总数据');
});

// ══ ② 不抢其它工具族的活 ══
test('"合并单元格"仍归 excel_merge_cells，不被合并硬短路抢走', async () => {
    const d = await dispatcher.tryDispatch(opts('把 A1:B2 单元格合并'));
    assert.notStrictEqual(d.toolName, 'excel_merge_sheets');
});

test('"合并多个 xlsx 文件" 需路径 → 不在 dispatcher 里强调 excel_merge_sheets', async () => {
    const d = await dispatcher.tryDispatch(opts('把桌面上的多个 xlsx 文件合并成一个工作簿'));
    assert.notStrictEqual(d.toolName, 'excel_merge_sheets');
});

test('复合意图"合并工作表后画个饼图"放行给多步骤管道', async () => {
    const d = await dispatcher.tryDispatch(opts('把这几张表合并成一张，然后画个饼图'));
    assert.strictEqual(d.hit, false, '复合意图应放行，不在硬短路里独吞');
});

// ══ ③ 工具瘦身：合并意图必须看得见 excel_merge_sheets ══
test('工具瘦身：合并请求保留 excel_merge_sheets（否则模型手上没有合并工具）', () => {
    const tools = [
        { name: 'excel_read_range' }, { name: 'excel_merge_cells' }, { name: 'excel_merge_sheets' },
        { name: 'excel_merge_workbooks' }, { name: 'excel_create_chart' }, { name: 'excel_format_range' }
    ];
    for (const msg of ['将选中工作表进行数据合并', '合并工作表的销售数据', '把这几个表合并到一张', '汇总表数据合并']) {
        const kept = util.filterToolsForMessage(tools, msg).map(t => t.name);
        assert.ok(kept.includes('excel_merge_sheets'), `"${msg}" 瘦身后应保留 excel_merge_sheets，实际=[${kept.join(',')}]`);
    }
});

test('工具瘦身：纯绘图请求不因合并词表扩容而带上合并工具', () => {
    const tools = [{ name: 'excel_read_range' }, { name: 'excel_merge_sheets' }, { name: 'excel_create_chart' }];
    const kept = util.filterToolsForMessage(tools, '按部门画个饼图').map(t => t.name);
    assert.ok(kept.includes('excel_create_chart'));
    assert.ok(!kept.includes('excel_merge_sheets'), '绘图请求不应附带合并工具');
});

test('工具瘦身：合并并设置成中国红必须同时留着合并和格式化', () => {
    const tools = [
        { name: 'excel_read_range' }, { name: 'excel_merge_sheets' },
        { name: 'excel_format_range' }, { name: 'excel_create_chart' }
    ];
    const kept = util.filterToolsForMessage(tools, '将选中工作表进行数据合并，并设置成中国红').map(t => t.name);
    assert.ok(kept.includes('excel_merge_sheets'), '合并工具不能被裁掉，实际=' + kept.join(','));
    assert.ok(kept.includes('excel_format_range'), '点了中国红必须看得见 excel_format_range，实际=' + kept.join(','));
    assert.ok(!kept.includes('excel_create_chart'), '这句不是画图');
});

test('选中两张表时，模型列出全部表也改回选中的两张', () => {
    const args = App.ChatV2.turnGoal.constrainMergeSource(
        {
            sourceSheets: ['销售二部', '销售一部', '大客户部', '销售三部', '市场部'],
            mergeAll: false,
            targetSheet: '合并结果'
        },
        '将选中工作表进行数据合并，并设置成中国红',
        ['销售二部', '销售一部']
    );
    assert.deepStrictEqual(args.sourceSheets, ['销售二部', '销售一部']);
    assert.strictEqual(args.mergeAll, false);
    assert.strictEqual(args.targetSheet, '合并结果');
});

test('说了全部工作表时，不把范围改回标签选中', () => {
    const raw = { sourceSheets: ['销售二部', '销售一部', '市场部'], mergeAll: true };
    const args = App.ChatV2.turnGoal.constrainMergeSource(raw, '合并所有工作表', ['销售二部', '销售一部']);
    assert.strictEqual(args, raw);
});

test('没说选中时，保留模型自己列出的表', () => {
    const raw = { sourceSheets: ['销售二部', '市场部'] };
    const args = App.ChatV2.turnGoal.constrainMergeSource(raw, '把这几张表合并成一张', ['销售二部', '销售一部']);
    assert.strictEqual(args, raw);
});

test('turn-goal：合并并设置成中国红，只合并不算完成', () => {
    const goal = App.ChatV2.turnGoal.build('将选中工作表进行数据合并，并设置成中国红');
    assert.strictEqual(goal.styleTarget, '中国红');
    assert.ok(goal.doneWhen.some(d => d.kind === 'merge'));
    assert.ok(goal.doneWhen.some(d => d.kind === 'format'));
    const merged = {
        goal: goal,
        _executedToolCalls: [{ name: 'excel_merge_sheets', args: { targetSheet: '合并结果' } }],
        _executedTools: ['excel_merge_sheets']
    };
    const gap = App.ChatV2.turnGoal.unmet(merged);
    assert.ok(gap.some(d => d.kind === 'format'), '只合并时外观仍未完成');
    const painted = {
        goal: goal,
        _executedToolCalls: [
            { name: 'excel_merge_sheets', args: { targetSheet: '合并结果' } },
            { name: 'excel_format_range', args: { sheet: '合并结果', range: 'A1:I100', format: { bgColor: 'DE2910', bgScope: 'header', headerFontColor: 'FFFFFF' } } }
        ],
        _executedTools: ['excel_merge_sheets', 'excel_format_range']
    };
    assert.strictEqual(App.ChatV2.turnGoal.unmet(painted).length, 0);
});

// ══ ④ 写类裁决闸门：合并是执行命令，不得落到 LLM 判 other ══
test('写类裁决：合并工作表类措辞直接判 command（source=rule）', async () => {
    for (const msg of ['将选中工作表进行数据合并', '把这几张表合并成一张后再写入新表', '合并工作表里的销售明细']) {
        const v = await orch.adjudicateWriteHit({ message: msg, toolName: 'excel_merge_sheets' });
        assert.strictEqual(v.verdict, 'command', `"${msg}" 应判 command`);
        assert.strictEqual(v.source, 'rule', `"${msg}" 不得落到 LLM 二次裁决（小模型易判 other）`);
    }
});

test('写类裁决：美化全部工作表为商务蓝直接判 command', async () => {
    for (const msg of ['美化工作簿中的所有工作表，统一为商务蓝', '将所有工作表统一为商务蓝', '转为商务蓝']) {
        const v = await orch.adjudicateWriteHit({ message: msg, toolName: 'excel_format_range' });
        assert.strictEqual(v.verdict, 'command', `"${msg}" 应判 command`);
        assert.strictEqual(v.source, 'rule', `"${msg}" 不得落到 LLM 二次裁决`);
    }
});

test('写类裁决：要代码时仍判 code（不因合并规则而执行）', async () => {
    const v = await orch.adjudicateWriteHit({ message: '写一段 VBA 把多个工作表合并到一张表', toolName: 'excel_merge_sheets' });
    assert.strictEqual(v.verdict, 'code');
});
