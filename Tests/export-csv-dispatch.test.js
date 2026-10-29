/**
 * ============================================================
 * Tests/export-csv-dispatch.test.js — Excel"导出 CSV"确定性硬短路回归
 * ============================================================
 * 实测失败（截图 2026-10-28 19:04）：用户说"将当前工作表导出为 CSV 并保存到桌面"，
 * Excel 侧没有 CSV 导出硬短路 → 交给 agent-loop，弱模型第 1 轮调了
 * excel_save_workbook（存了个 xlsx），第 2 轮（总结轮，tools 已锁死不下发）只写下
 * "调用 `excel_export_csv` 工具导出 CSV 文件"的永动机计划 —— CSV 永远没导出。
 *
 * 修复：Excel 宿主命中"导出/另存 + CSV"→ 直调 excel_export_csv（sheet 取活动表，
 *      filePath 留空由后端落盘桌面）。多表批量导出仍交规划（一次只导一张表）。
 * 运行：node Tests/export-csv-dispatch.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const { window, App } = loadChatV2(['config/routing-constants.js', 'services/mcp-dispatcher.js']);
const dispatcher = App.ChatV2.mcpDispatcher;

const opts = (message, context) => ({ message, toolRegistry: null, sessionStore: null, turn: { context: context || {} } });
const setHost = (host) => { window.SmartPilotHost = host; };
const EXCEL = { isExcel: true, isPowerPoint: false, kind: 'excel' };
const WORD = { isExcel: false, isPowerPoint: false, kind: 'word' };

test('Excel：将当前工作表导出为 CSV 并保存到桌面 → 硬短路 excel_export_csv', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('将当前工作表导出为 CSV 并保存到桌面'));
    assert.strictEqual(d.hardMatched, true, '应确定性直调，不能交给弱模型（实测它先存 xlsx 再干说）');
    assert.strictEqual(d.toolName, 'excel_export_csv');
});

test('Excel：把数据另存为 csv 文件 → excel_export_csv', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('把数据另存为 csv 文件'));
    assert.strictEqual(d.toolName, 'excel_export_csv');
});

test('Excel：导出为 CSV 格式 → excel_export_csv（"格式"二字不得被当成排版意图排除）', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('把这张表导出为 CSV 格式'));
    assert.strictEqual(d.toolName, 'excel_export_csv');
});

test('Excel：筛选后导出 CSV → 交规划（要先筛选，不能只导出）', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('筛选销售一部后导出为 CSV'));
    assert.notStrictEqual(d.toolName, 'excel_export_csv');
});

test('Excel：导出数据 → 不猜 CSV（未提目标时放行）', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('把数据导出'));
    assert.notStrictEqual(d.toolName, 'excel_export_csv', '未提 CSV 时不得擅自决定产物格式');
});

test('Excel：所有工作表都导出成 CSV → 不硬短路（一次只能导一张表）', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('把所有工作表都导出成 CSV'));
    assert.notStrictEqual(d.toolName, 'excel_export_csv', '多表导出应交规划，避免只导当前表还宣称完成');
});

test('Excel：选中表名/活动表会作为 sheet 传给 excel_export_csv', async () => {
    setHost(EXCEL);
    const d = await dispatcher.tryDispatch(opts('把当前工作表导出为CSV', {
        activeSheet: { sheet: '销售数据' }
    }));
    assert.strictEqual(d.toolName, 'excel_export_csv');
    assert.strictEqual(d.args.sheet, '销售数据');
    assert.deepStrictEqual(Object.keys(d.args).filter(k => k === 'filePath'), [],
        'filePath 必须留空，由后端落盘桌面（_ 时间戳文件名）');
});

test('Word 宿主：导出 CSV 请求不得派发到 Excel 工具', async () => {
    setHost(WORD);
    const d = await dispatcher.tryDispatch(opts('将当前文档内容导出为 CSV'));
    assert.notStrictEqual(d.toolName, 'excel_export_csv');
});
