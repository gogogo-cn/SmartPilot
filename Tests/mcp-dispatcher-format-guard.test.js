/**
 * ============================================================
 * Tests/mcp-dispatcher-format-guard.test.js
 * 验证"多要求格式化不再被确定性硬短路静默丢弃"
 * 运行：node Tests/mcp-dispatcher-format-guard.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

// dispatcher 的关键词常量来自 config/routing-constants.js（单一事实源），必须一并加载
const { App } = loadChatV2(['config/routing-constants.js', 'services/mcp-dispatcher.js']);
const dispatcher = App.ChatV2.mcpDispatcher;

const opts = (message) => ({ message, toolRegistry: null, sessionStore: null, turn: { context: {} }, allowHardDispatch: true });

test('多项要求(底色/边框/对齐/列宽) → 放弃确定性硬短路，放行 agent-loop', async () => {
    const d = await dispatcher.tryDispatch(opts('美化当前选中的表格区域：设置表头底色、边框、对齐方式与列宽，使其清晰美观'));
    assert.strictEqual(d.hit, false, '不应直调 excel_format_range');
    assert.strictEqual(d.formatMultiGuard, true, '应标记为多要求放行');
    assert.ok(!d.hardMatched, '不应走确定性硬短路');
});

test('"底色+边框"混合(底色颜色解析不出) → 放行，不静默丢弃底色', async () => {
    const d = await dispatcher.tryDispatch(opts('设置表头底色和边框'));
    assert.strictEqual(d.hit, false);
    assert.strictEqual(d.formatMultiGuard, true);
});

test('单一且可覆盖的要求(加边框) → 仍走确定性硬短路', async () => {
    const d = await dispatcher.tryDispatch(opts('把A1:D1加边框'));
    // 测试环境无 CallLocalTool → 硬短路命中后返回 bridge-unavailable，但 hardMatched=true 证明短路已触发
    assert.strictEqual(d.hardMatched, true, '简单可覆盖的格式化应仍走硬短路');
    assert.strictEqual(d.toolName, 'excel_format_range');
});

test('多项且全部可覆盖(边框+字体颜色) → 仍走确定性硬短路', async () => {
    const d = await dispatcher.tryDispatch(opts('把A1:D1加边框，字体颜色设为红色'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_format_range');
});

test('加粗+明确字号 → 仍走确定性硬短路', async () => {
    const d = await dispatcher.tryDispatch(opts('把A1:B2加粗，字号12'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_format_range');
});

test('「清洗这张表」硬短路必须就地写回，不能只回报预览', async () => {
    const d = await dispatcher.tryDispatch(opts('清洗这张表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_clean_data');
    assert.strictEqual(d.args.dryRun, false, '默认 dryRun 会只说清洗完成、表格不动');
    assert.strictEqual(d.args.writeBack, true);
    const ops = JSON.parse(d.args.operations);
    assert.ok(ops.includes('removeDuplicates'), '清洗整表必须查整行重复；实际=' + d.args.operations);
});

test('明确只要预览时不写回', async () => {
    const d = await dispatcher.tryDispatch(opts('先看看清洗效果，别改表'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_clean_data');
    assert.ok(d.args.writeBack !== true);
});

test('「清洗这份表格，移除重复行」必须写回，不能只计算', async () => {
    const d = await dispatcher.tryDispatch(opts('清洗这份表格，移除重复行'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_clean_data', 'excel_remove_duplicates 只算不写');
    assert.strictEqual(d.args.dryRun, false);
    assert.strictEqual(d.args.writeBack, true);
    const ops = JSON.parse(d.args.operations);
    assert.ok(ops.includes('removeDuplicates'), '实际=' + d.args.operations);
});

test('单独说去重也要删行写回', async () => {
    const d = await dispatcher.tryDispatch(opts('去重'));
    assert.strictEqual(d.toolName, 'excel_clean_data');
    assert.strictEqual(d.args.writeBack, true);
});

test('只高亮不删除时不去改表', async () => {
    const d = await dispatcher.tryDispatch(opts('去重并只把重复行黄色高亮标记，不删除'));
    assert.strictEqual(d.toolName, 'excel_remove_duplicates');
    assert.ok(d.args.writeBack !== true);
    assert.strictEqual(d.args.markOnly, true);
});

test('「标记重复行」只标黄，不能删行', async () => {
    const d = await dispatcher.tryDispatch(opts('标记重复行'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_remove_duplicates', '标记不能走 excel_clean_data 的 removeDuplicates');
    assert.ok(d.args.writeBack !== true);
    assert.strictEqual(d.args.markOnly, true);
});

test('「标记王五」走条件格式标黄，不追问', async () => {
    const d = await dispatcher.tryDispatch(opts('标记王五'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_apply_conditional_format');
    assert.strictEqual(d.args.formatType, 'highlight');
    assert.strictEqual(d.args.markLiteral, '王五');
    const opt = JSON.parse(d.args.options);
    assert.strictEqual(opt.operator, 'equal');
    assert.strictEqual(opt.value, '王五');
    assert.strictEqual(opt.color, 'FFFF00');
});

test('「将王五标红」必须真的套红色，不能只读表', async () => {
    const d = await dispatcher.tryDispatch(opts('将王五标红'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_apply_conditional_format');
    assert.strictEqual(d.args.markLiteral, '王五');
    assert.strictEqual(d.args.markColorName, '红色');
    const opt = JSON.parse(d.args.options);
    assert.strictEqual(opt.color, 'FF0000');
    assert.strictEqual(opt.value, '王五');
});

test('换成别的颜色也要套上对应色，不能掉回只读', async () => {
    const cases = [
        ['将王五标成绿色', '王五', '00B050', '绿色'],
        ['把王五涂成橙色', '王五', 'FFC000', '橙色'],
        ['将王五换成蓝色', '王五', '5B9BD5', '蓝色'],
        ['标紫王五', '王五', '7030A0', '紫色'],
        ['把李四标成浅蓝色', '李四', 'BDD7EE', '浅蓝色']
    ];
    for (const [msg, value, hex, name] of cases) {
        const d = await dispatcher.tryDispatch(opts(msg));
        assert.strictEqual(d.toolName, 'excel_apply_conditional_format', msg);
        assert.strictEqual(d.args.markLiteral, value, msg);
        assert.strictEqual(d.args.markColorName, name, msg);
        assert.strictEqual(JSON.parse(d.args.options).color, hex, msg);
    }
});

test('「移除销售一部数据」必须按值删行，不能只读表', async () => {
    const d = await dispatcher.tryDispatch(opts('移除销售一部数据'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_delete_row');
    assert.strictEqual(d.args.matchValue, '销售一部');
    assert.strictEqual(d.args.deleteMatchingRows, true);
});

test('「将销售一部数据从单元格中移除」只匹配销售一部', async () => {
    const d = await dispatcher.tryDispatch(opts('将销售一部数据从单元格中移除'));
    assert.strictEqual(d.toolName, 'excel_delete_row');
    assert.strictEqual(d.args.matchValue, '销售一部');
    assert.strictEqual(d.args.deleteMatchingRows, true);
});

test('「把王五删掉」同样按值删行', async () => {
    const d = await dispatcher.tryDispatch(opts('把王五删掉'));
    assert.strictEqual(d.toolName, 'excel_delete_row');
    assert.strictEqual(d.args.matchValue, '王五');
});

test('「将销售一部黄色」必须套黄色，不能只读表', async () => {
    const d = await dispatcher.tryDispatch(opts('将销售一部黄色'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_apply_conditional_format');
    assert.strictEqual(d.args.markLiteral, '销售一部');
    assert.strictEqual(d.args.markColorName, '黄色');
    const opt = JSON.parse(d.args.options);
    assert.strictEqual(opt.operator, 'equal');
    assert.strictEqual(opt.value, '销售一部');
    assert.strictEqual(opt.color, 'FFFF00');
});

test('「把李四标黄」同样标出该值', async () => {
    const d = await dispatcher.tryDispatch(opts('把李四标黄'));
    assert.strictEqual(d.toolName, 'excel_apply_conditional_format');
    assert.strictEqual(d.args.markLiteral, '李四');
});

test('模型按原话改参数时覆盖种子值，没给的字段保留', () => {
    const merged = dispatcher._mergeSemanticArgs(
        {
            formatType: 'highlight',
            options: JSON.stringify({ operator: 'equal', value: '王五', color: 'FFFF00' }),
            markLiteral: '王五',
            markColorName: '黄色',
            sheet: 'Sheet1',
            writeBack: true
        },
        {
            options: { operator: 'equal', value: '王五', color: 'C6EFCE' },
            markColorName: '薄荷绿'
        },
        '把王五标成薄荷绿'
    );
    const opt = JSON.parse(merged.options);
    assert.strictEqual(opt.color, 'C6EFCE');
    assert.strictEqual(opt.value, '王五');
    assert.strictEqual(merged.markColorName, '薄荷绿');
    assert.strictEqual(merged.sheet, 'Sheet1');
    assert.strictEqual(merged.writeBack, true);
});

test('模型照抄示例列 A、默认黄、中文色名时不覆盖已解析参数', () => {
    const sort = dispatcher._mergeSemanticArgs(
        { sortBy: JSON.stringify([{ column: '销售额', order: 'descending' }]) },
        { column: 'A', sortBy: JSON.stringify([{ column: 'A', order: 'descending' }]) },
        '按销售额从高到低'
    );
    assert.strictEqual(sort.column, undefined);
    assert.strictEqual(JSON.parse(sort.sortBy)[0].column, '销售额');
    const named = dispatcher._mergeSemanticArgs(
        { sortBy: JSON.stringify([{ column: 'A', order: 'descending' }]) },
        { sortBy: [{ column: '销售额', order: 'descending' }] },
        '按销售额从高到低'
    );
    assert.strictEqual(JSON.parse(named.sortBy)[0].column, '销售额');

    const color = dispatcher._mergeSemanticArgs(
        { options: JSON.stringify({ operator: 'equal', value: '王五', color: '00B050' }) },
        { options: { color: '绿色', operator: 'equal' } },
        '将王五标成绿色'
    );
    assert.strictEqual(JSON.parse(color.options).color, '00B050');

    const yellow = dispatcher._mergeSemanticArgs(
        { options: JSON.stringify({ color: '00B050', value: '王五' }) },
        { options: { color: 'FFFF00' } },
        '将王五标成绿色'
    );
    assert.strictEqual(JSON.parse(yellow.options).color, '00B050');
});

test('Q3 是第三季度：标题用原话，行只留 7–9 月', () => {
    const period = dispatcher._parseQuarterPeriod('根据这份数据生成一份 Q3 经营分析 PPT');
    assert.strictEqual(period.quarter, 3);
    assert.strictEqual(period.startMonth, 7);
    assert.strictEqual(period.endMonth, 9);
    assert.strictEqual(period.year, null);
    const shaped = dispatcher._shapePptArgs('根据这份数据生成一份 Q3 经营分析 PPT', ['日期', '销售额'], [
        { 日期: '2026-01-15', 销售额: 1 },
        { 日期: '2026-08-02', 销售额: 2 },
        { 日期: '2026-11-01', 销售额: 3 }
    ], '销售');
    assert.strictEqual(shaped.reportTitle, 'Q3 经营分析');
    assert.strictEqual(shaped.rows.length, 1);
    assert.strictEqual(shaped.rows[0].销售额, 2);
    assert.ok(shaped.sourceLabel.indexOf('Q3') >= 0);
});

test('2026年第三季度只留该年 7–9 月', () => {
    const shaped = dispatcher._shapePptArgs('做一份2026年第三季度经营分析PPT', ['日期', '销售额'], [
        { 日期: '2025-08-01', 销售额: 1 },
        { 日期: '2026-09-30', 销售额: 2 }
    ], '');
    assert.strictEqual(shaped.reportTitle, '2026年第三季度经营分析');
    assert.strictEqual(shaped.rows.length, 1);
    assert.strictEqual(shaped.rows[0].销售额, 2);
});

test('按部门分组可以再交给图表，不绑在 PPT 上', () => {
    const headers = ['部门', '销售额'];
    const rows = [
        { 部门: '销售一部', 销售额: 10 },
        { 部门: '销售一部', 销售额: 5 },
        { 部门: '销售二部', 销售额: 3 }
    ];
    const plan = dispatcher._parseGroupPlan('按部门分组做柱状图', headers);
    assert.strictEqual(plan.dimColumn, '部门');
    assert.strictEqual(plan.follow, 'chart');
    assert.strictEqual(plan.chartType, 'xlColumnClustered');
    const view = dispatcher._applyGroupView(headers, rows, plan);
    assert.strictEqual(view.rows.length, 2);
    assert.strictEqual(view.rows.find(r => r.部门 === '销售一部').销售额, 15);
});

test('按季度分组把同一季的行收成一组', () => {
    const plan = dispatcher._parseGroupPlan('按季度分组汇总销售额', ['日期', '销售额']);
    assert.strictEqual(plan.grain, 'quarter');
    assert.strictEqual(plan.follow, '');
    const view = dispatcher._applyGroupView(['日期', '销售额'], [
        { 日期: '2026-01-15', 销售额: 1 },
        { 日期: '2026-02-02', 销售额: 4 },
        { 日期: '2026-08-02', 销售额: 9 }
    ], plan);
    assert.strictEqual(view.rows.length, 2);
    assert.strictEqual(view.rows.find(r => r.季度 === '2026年Q1').销售额, 5);
});

test('Q3 再按部门：先留 7–9 月，再按部门汇总', () => {
    const headers = ['日期', '部门', '销售额'];
    const plan = dispatcher._parseGroupPlan('把 Q3 按部门分组', headers);
    assert.strictEqual(plan.period.quarter, 3);
    assert.strictEqual(plan.dimColumn, '部门');
    const view = dispatcher._applyGroupView(headers, [
        { 日期: '2026-01-15', 部门: '销售一部', 销售额: 10 },
        { 日期: '2026-08-02', 部门: '销售一部', 销售额: 4 },
        { 日期: '2026-09-01', 部门: '销售一部', 销售额: 6 },
        { 日期: '2026-11-01', 部门: '销售二部', 销售额: 8 }
    ], plan);
    assert.strictEqual(view.rows.length, 1);
    assert.strictEqual(view.rows[0].部门, '销售一部');
    assert.strictEqual(view.rows[0].销售额, 10);
});

test('按部门分组生成 PPT 用的是汇总行', () => {
    const shaped = dispatcher._shapePptArgs('按部门分组生成一份PPT', ['部门', '销售额'], [
        { 部门: '销售一部', 销售额: 10 },
        { 部门: '销售一部', 销售额: 5 },
        { 部门: '销售二部', 销售额: 3 }
    ], '销售');
    assert.strictEqual(shaped.reportTitle, '按部门分组');
    assert.strictEqual(shaped.rows.length, 2);
    assert.strictEqual(shaped.rows.find(r => r.部门 === '销售一部').销售额, 15);
});

test('只说按部门分组时直接汇总，不停在读表', async () => {
    const d = await dispatcher.tryDispatch({
        message: '按部门分组汇总销售额',
        toolRegistry: null,
        sessionStore: null,
        allowHardDispatch: true,
        turn: { context: { selection: { headers: ['部门', '销售额'], rows: [{ 部门: '销售一部', 销售额: 1 }], sheetName: '销售' } } }
    });
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_aggregate');
    assert.strictEqual(d.args.dimColumn, '部门');
    assert.strictEqual(d.args.valueColumn, '销售额');
    assert.strictEqual(d.args.targetSheet, '分组结果');
    assert.strictEqual(d.args.sheet, '销售');
});

test('「美化表格」走预设，不把整表涂成一个浅色', async () => {
    const d = await dispatcher.tryDispatch(opts('美化表格'));
    assert.strictEqual(d.hardMatched, true);
    assert.strictEqual(d.toolName, 'excel_format_range');
    assert.strictEqual(d.args.format, '{}');
    assert.strictEqual(d.args.autoRange, true);
    assert.strictEqual(d.args.beautifyPreset, true);
});

test('指明颜色的美化不走整表预设', () => {
    assert.strictEqual(dispatcher.isBareTableBeautify('美化表格，表头改成红色'), false);
    assert.strictEqual(dispatcher.isBareTableBeautify('把当前表格美化一下'), true);
});

test('「高亮重复行」同样只标黄', async () => {
    const d = await dispatcher.tryDispatch(opts('高亮重复行'));
    assert.strictEqual(d.toolName, 'excel_remove_duplicates');
    assert.strictEqual(d.args.markOnly, true);
    assert.ok(d.args.writeBack !== true);
});
