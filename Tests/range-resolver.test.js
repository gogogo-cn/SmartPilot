/**
 * ============================================================
 * Tests/range-resolver.test.js — 选区范围判断与降级阶梯（单一事实源）
 * ============================================================
 * 背景（实测 2026-10-22 用户截图）：
 *   "将工作簿中的所有工作表转为商务蓝"跑通后，7 张表里 4 张只格式化了 G30 单格、
 *   3 张没有范围。根因：Excel 的 UsedRange 维度缓存对【从未激活过的工作表】不刷新，
 *   拆分/新建的表返回光标单格（G30）或空串，前端原样传给 excel_format_range →
 *   "只美化 1 个格子还报成功"。
 *
 * 本测试锁死三件事：
 *   ① isReliableTableRange / isSingleCellAddr 的边界（单格/空/畸形/单行/$前缀/sheet前缀）；
 *   ② usedRangeAddr / readRangeAddr 对各种返回形态的解析（含 maxRows 截断注释污染 JSON）；
 *   ③ createSheetRangeLadder 三级降级的每一级条件、调用参数与"全不可靠→null"结局。
 *
 * 运行：node --test Tests/range-resolver.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const { window, App } = loadChatV2(DEFAULT_ORDER);
const RR = App.ChatV2.rangeResolver;

// ══ 第一组：模块注册与加载顺序 ══
test('rangeResolver 已随 DEFAULT_ORDER 注册（agent-loop/mcp-dispatcher 之前加载）', () => {
    assert.ok(RR, 'App.ChatV2.rangeResolver 应存在（harness DEFAULT_ORDER 与 index.html 同序）');
    for (const fn of ['colToNum', 'normRangeAddr', 'isSingleCellAddr', 'isReliableTableRange',
        'usedRangeAddr', 'readRangeAddr', 'createSheetRangeLadder',
        'classifySelection', 'extractSpokenRange', 'decideFormatTarget', 'headerRowOf']) {
        assert.strictEqual(typeof RR[fn], 'function', '缺少导出: ' + fn);
    }
});

// ══ 第二组：isReliableTableRange（能否当"整张数据表"目标）══
test('isReliableTableRange：多行多格才算可靠；单格/空/畸形/单行一律不可靠', () => {
    assert.strictEqual(RR.isReliableTableRange('A1:I21'), true, '表头+20 行数据');
    assert.strictEqual(RR.isReliableTableRange('$A$1:$I$21'), true, '带 $ 也应认');
    assert.strictEqual(RR.isReliableTableRange('Sheet1!A1:I21'), true, '带 sheet 前缀也应认');
    assert.strictEqual(RR.isReliableTableRange('AA10:AB20'), true, '多字母列');
    assert.strictEqual(RR.isReliableTableRange('A1:B2'), true, '最小 2 行');
    // 不可靠：单格（未激活表的陈旧 UsedRange 就长这样）、空、畸形、只有表头一行
    assert.strictEqual(RR.isReliableTableRange('G30'), false, '单格=光标，不是数据表');
    assert.strictEqual(RR.isReliableTableRange('$G$30'), false, '带 $ 的单格');
    assert.strictEqual(RR.isReliableTableRange('A1'), false, '单格');
    assert.strictEqual(RR.isReliableTableRange(''), false, '空');
    assert.strictEqual(RR.isReliableTableRange(null), false, 'null');
    assert.strictEqual(RR.isReliableTableRange('A1:I'), false, '畸形（缺行号）');
    assert.strictEqual(RR.isReliableTableRange('A1:I1'), false, '只有表头一行');
    assert.strictEqual(RR.isReliableTableRange('A1:A1'), false, '单行单列');
});

// ══ 第三组：isSingleCellAddr（是否"光标单格"）══
test('isSingleCellAddr：单格/空为 true；任何多格区域（含单行）为 false', () => {
    assert.strictEqual(RR.isSingleCellAddr('A1'), true);
    assert.strictEqual(RR.isSingleCellAddr('$G$30'), true, '⚠ 必须归一化 $ 再判：GetSelection 返回带 $ 地址');
    assert.strictEqual(RR.isSingleCellAddr('Sheet1!A1'), true);
    assert.strictEqual(RR.isSingleCellAddr(''), true, '空=没有可用区域');
    assert.strictEqual(RR.isSingleCellAddr(null), true);
    assert.strictEqual(RR.isSingleCellAddr('A1:B2'), false);
    assert.strictEqual(RR.isSingleCellAddr('A5:G5'), false, '单行也是用户圈定的区域（agent-loop 语义）');
    assert.strictEqual(RR.isSingleCellAddr('A1,B2'), false, '多区联合');
    assert.strictEqual(RR.isSingleCellAddr('$A$1:$B$2'), false);
});

test('normRangeAddr / colToNum', () => {
    assert.strictEqual(RR.normRangeAddr('$A$1:$B$2'), 'A1:B2');
    assert.strictEqual(RR.normRangeAddr('Sheet1!$G$30'), 'G30');
    assert.strictEqual(RR.normRangeAddr('  A1  '), 'A1');
    assert.strictEqual(RR.normRangeAddr(''), '');
    assert.strictEqual(RR.colToNum('A'), 1);
    assert.strictEqual(RR.colToNum('Z'), 26);
    assert.strictEqual(RR.colToNum('AA'), 27);
});

// ══ 第四组：usedRangeAddr / readRangeAddr 返回形态解析 ══
test('usedRangeAddr：认 {range}/{usedRange}/JSON 字符串/result 下钻/content[0].text', () => {
    assert.strictEqual(RR.usedRangeAddr({ range: '$A$1:$G$21' }), 'A1:G21');
    assert.strictEqual(RR.usedRangeAddr({ usedRange: 'A1:D9' }), 'A1:D9');
    assert.strictEqual(RR.usedRangeAddr(JSON.stringify({ range: 'B2:E10' })), 'B2:E10');
    assert.strictEqual(RR.usedRangeAddr({ result: JSON.stringify({ range: 'A1:C5' }) }), 'A1:C5');
    assert.strictEqual(RR.usedRangeAddr({ content: [{ text: JSON.stringify({ range: 'A1:C7' }) }] }), 'A1:C7');
    assert.strictEqual(RR.usedRangeAddr({ success: true }), null, '无 range 字段');
    assert.strictEqual(RR.usedRangeAddr('not json'), null);
    assert.strictEqual(RR.usedRangeAddr(null), null);
});

test('readRangeAddr：maxRows 截断注记污染 JSON 时切到最后一个 } 仍能解析', () => {
    // 后端 maxRows=1 截断会在 JSON 尾部追加"/* ⚠️ 数据已截断… */"注释，JSON.parse 直接失败
    const truncated = JSON.stringify({ worksheet: '销售三部', range: '$A$1:$G$21', rows: 0, columns: 7, headers: [], result: [] })
        + '\n/* ⚠️ 数据已截断：行数超过 maxRows=1。请缩小 range 范围后重新读取 */';
    assert.strictEqual(RR.readRangeAddr(truncated), 'A1:G21');
    assert.strictEqual(RR.readRangeAddr({ range: 'A1:B2' }), 'A1:B2');
    assert.strictEqual(RR.readRangeAddr({ result: { range: 'C3:F9' } }), 'C3:F9');
    assert.strictEqual(RR.readRangeAddr({ content: [{ text: JSON.stringify({ range: 'A1:D4' }) }] }), 'A1:D4');
    assert.strictEqual(RR.readRangeAddr({ success: true }), null);
    assert.strictEqual(RR.readRangeAddr(null), null);
});

// ══ 第五组：三级降级阶梯 ══
/** 构造记录调用的桩 callTool */
function makeStub(handlers) {
    const calls = [];
    const call = async (tool, argsJson) => {
        const args = (() => { try { return JSON.parse(argsJson); } catch (_) { return {}; } })();
        calls.push({ tool, args });
        const h = handlers[tool];
        if (typeof h === 'function') return { raw: h(args) };
        if (h === 'throw') throw new Error('桥不可用');
        return { raw: null };
    };
    return { call, calls };
}
const okUr = { range: '$A$1:$G$21' };
const okRr = { worksheet: 'x', range: '$A$1:$G$21' };

test('阶梯①：种子范围可靠 → 直接用，零工具调用', async () => {
    const stub = makeStub({});
    const ladder = RR.createSheetRangeLadder(stub.call);
    const r = await ladder('市场部', 'A1:I21');
    assert.deepStrictEqual(r, { range: 'A1:I21', via: 'get_sheets' });
    assert.strictEqual(stub.calls.length, 0, '可靠种子不应再调任何工具');
});

test('阶梯②：种子是单格/空（未激活表陈旧值）→ get_used_range 兜底', async () => {
    for (const seed of ['G30', '']) {
        const stub = makeStub({ 'excel_get_used_range': () => okUr });
        const ladder = RR.createSheetRangeLadder(stub.call);
        const r = await ladder('销售三部', seed);
        assert.deepStrictEqual(r, { range: 'A1:G21', via: 'get_used_range' }, 'seed=' + JSON.stringify(seed));
        assert.strictEqual(stub.calls.length, 1);
        assert.strictEqual(stub.calls[0].tool, 'excel_get_used_range');
        assert.strictEqual(stub.calls[0].args.sheet, '销售三部');
    }
});

test('阶梯③：get_used_range 仍不可靠 → read_range + autoProbe（CurrentRegion，不依赖激活态）', async () => {
    const stub = makeStub({
        'excel_get_used_range': () => ({ range: 'G30' }),          // 还是陈旧单格
        'excel_read_range': () => okRr
    });
    const ladder = RR.createSheetRangeLadder(stub.call);
    const r = await ladder('大客户部', 'G30');
    assert.deepStrictEqual(r, { range: 'A1:G21', via: 'read_range:CurrentRegion' });
    assert.strictEqual(stub.calls.length, 2);
    const rr = stub.calls[1];
    assert.strictEqual(rr.tool, 'excel_read_range');
    assert.deepStrictEqual(rr.args,
        { sheet: '大客户部', range: 'A1', autoProbe: true, maxRows: 1 },
        '必须显式 autoProbe + maxRows=1（只取地址不搬数据）');
});

test('阶梯③：get_used_range 抛异常不阻断，继续 read_range', async () => {
    const stub = makeStub({
        'excel_get_used_range': 'throw',
        'excel_read_range': () => okRr
    });
    const ladder = RR.createSheetRangeLadder(stub.call);
    const r = await ladder('表X', 'G30');
    assert.deepStrictEqual(r, { range: 'A1:G21', via: 'read_range:CurrentRegion' });
    assert.strictEqual(stub.calls.length, 2);
});

test('阶梯④：三级全不可靠 → null（调用方必须跳过该表并说明原因，不得猜范围）', async () => {
    const stub = makeStub({
        'excel_get_used_range': () => ({ range: 'G30' }),
        'excel_read_range': () => ({ range: 'G30' })
    });
    const ladder = RR.createSheetRangeLadder(stub.call);
    const r = await ladder('空表', '');
    assert.strictEqual(r, null);
    assert.strictEqual(stub.calls.length, 2, '两级都试过才放弃');
});

test('阶梯④：全部抛异常 → null 且不抛出', async () => {
    const stub = makeStub({
        'excel_get_used_range': 'throw',
        'excel_read_range': 'throw'
    });
    const ladder = RR.createSheetRangeLadder(stub.call);
    const r = await ladder('表Y', 'G30');
    assert.strictEqual(r, null);
    assert.strictEqual(stub.calls.length, 2);
});

test('createSheetRangeLadder 不传 callTool 时不炸（返回 null 走安全侧）', async () => {
    const ladder = RR.createSheetRangeLadder(null);
    const r = await ladder('表Z', 'G30');
    assert.strictEqual(r, null);
});

function plan(message, selection, extra) {
    return RR.decideFormatTarget(message, Object.assign({
        selection: selection || null,
        activeSheet: { sheet: '市场部' },
        sheets: ['销售二部', '销售一部', '大客户部', '销售三部', '市场部']
    }, extra || {}));
}

test('口头美化每张表：没选区、光标单格、自动捕获、只圈在一张表上，都用各表自己的数据区', () => {
    const msg = '美化工作簿中的所有工作表，统一为科技绿';
    assert.strictEqual(plan(msg, null).rangeMode, 'table');
    assert.strictEqual(plan(msg, null).sheetScope, 'all');
    const cursor = plan(msg, { address: '$G$30', sheetName: '市场部' });
    assert.strictEqual(cursor.rangeMode, 'table');
    assert.strictEqual(cursor.reason, 'cursor-ignored-for-all-sheets');
    const auto = plan(msg, { address: 'A1:I21', sheetName: '市场部', isAutoCaptured: true });
    assert.strictEqual(auto.rangeMode, 'table');
    assert.strictEqual(auto.selectionKind, 'auto');
    const oneBlock = plan(msg, { address: 'A1:I21', sheetName: '市场部' });
    assert.strictEqual(oneBlock.rangeMode, 'table', '没说按这个范围，不把一张表的选区套到其他表');
});

test('圈了一行用这一行；只点一个格子不算圈定，改用该表数据区', () => {
    const row = plan('美化选中的这一行', { address: 'A5:I5', sheetName: '市场部' });
    assert.strictEqual(row.rangeMode, 'selection');
    assert.strictEqual(row.selectionRange, 'A5:I5');
    assert.strictEqual(row.sheetScope, 'active');
    const cell = plan('美化这个表', { address: 'G30', sheetName: '市场部' });
    assert.strictEqual(cell.rangeMode, 'table');
    assert.strictEqual(cell.reason, 'cursor-cell-use-table');
    const namedCell = plan('把这个单元格涂红', { address: 'C3', sheetName: '市场部' });
    assert.strictEqual(namedCell.rangeMode, 'cell');
});

test('嘴里指定的地址原样用；点名的表不用别的表上的选区', () => {
    const spoken = plan('把所有工作表的 A1:D10 统一成中国红', null);
    assert.strictEqual(spoken.rangeMode, 'spoken');
    assert.strictEqual(spoken.spokenRange, 'A1:D10');
    assert.strictEqual(spoken.sheetScope, 'all');
    const named = plan('美化市场部', { address: 'A1:I21', sheetName: '销售二部' });
    assert.strictEqual(named.sheetScope, 'named');
    assert.deepStrictEqual(named.namedSheets, ['市场部']);
    assert.strictEqual(named.rangeMode, 'table');
    assert.strictEqual(named.reason, 'selection-sheet-mismatch');
});

test('其他表按同一块才复制选区；只要表头、第 N 行、整列、整表误选、否定，各走各的', () => {
    const copy = plan('其他工作表也按这个区域美化', { address: 'B2:F20', sheetName: '市场部' });
    assert.strictEqual(copy.sheetScope, 'others');
    assert.strictEqual(copy.rangeMode, 'copy-selection');
    const copyCell = plan('其他工作表也按这个区域美化', { address: 'G30', sheetName: '市场部' });
    assert.strictEqual(copyCell.rangeMode, 'table');
    const header = plan('美化当前表的表头', null);
    assert.strictEqual(header.rangeMode, 'header');
    assert.strictEqual(RR.headerRowOf('A1:I21'), 'A1:I1');
    const rowN = plan('把第3行加粗', null);
    assert.strictEqual(rowN.rangeMode, 'row');
    assert.strictEqual(rowN.spokenRange, '3');
    const col = plan('把 B列 居中', null);
    assert.strictEqual(col.rangeMode, 'col');
    assert.strictEqual(col.spokenRange, 'B');
    const whole = plan('美化这个表', { address: 'A1:XFD1048576', sheetName: '市场部' });
    assert.strictEqual(whole.selectionKind, 'whole');
    assert.strictEqual(whole.rangeMode, 'table');
    const no = plan('不要美化所有工作表', null);
    assert.strictEqual(no.rangeMode, 'skip');
    assert.strictEqual(plan('保持原样', null).rangeMode, 'skip');
    assert.strictEqual(plan('保持原样', null).sheetScope, 'none');
});

test('口头单格、整列、整行原样保留；图表和超宽选区不算圈定；阶梯定不到就跳过', () => {
    for (const addr of ['A1', 'A:C', '1:3']) {
        const p = plan('美化所有工作表的 ' + addr, { address: 'G30', sheetName: '市场部' });
        assert.strictEqual(p.rangeMode, 'spoken', addr);
        assert.strictEqual(p.spokenRange, addr, addr);
        assert.strictEqual(p.sheetScope, 'all', addr);
    }
    const chart = plan('美化这个表', { address: 'A1:I21', sheetName: '市场部', objectType: 'chart' });
    assert.strictEqual(chart.selectionKind, 'object');
    assert.strictEqual(chart.rangeMode, 'table');
    const wide = plan('美化这个表', { address: 'A1:AZ20', sheetName: '市场部' });
    assert.strictEqual(wide.selectionKind, 'whole');
    assert.strictEqual(wide.rangeMode, 'table');
});
