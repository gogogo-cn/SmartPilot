/**
 * ============================================================
 * Tests/export-intent.test.js — "导出/保存"意图的提示词路由回归
 * ============================================================
 * 背景（实测失败，日志取证）：用户说"将当前工作表导出为CSV并保存到桌面"，
 *   detectAnalysisIntent        → false（无分析词）
 *   detectOperationOnlyIntent   → false（旧词表只覆盖 格式化|美化|样式|主题|合并单元格|拆分|转置）
 * ⇒ 落进"分析类"分支，被注入"1. 先调用 excel_read_range 读取活动工作表的实际数据"，
 *   模型照做：读完整表($A$1:$G$201) → 输出一段数据分析；导出工具一次都没被调用，
 *   任务卡显示"任务完成"，磁盘上却没有 CSV。
 *
 * 本套件锁定：文件产出意图必须走"直接调导出工具"，且不得再出现"先调用 excel_read_range"；
 * 同时守住分析类/格式操作类的既有行为不被改写。
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadChatV2 } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');
const EXPORT_MSG = '将当前工作表导出为CSV并保存到桌面';

/**
 * 构造真实 turn 并产出实际下发给 LLM 的 messages。
 * @param {string} message 用户消息
 */
function buildContext(message) {
    const { App } = loadChatV2();
    const pb = App.ChatV2.promptBuilder;
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({
        message,
        kind: 'text',
        context: { activeSheet: { sheet: '全流程综合测试表' } }
    });
    const messages = pb.buildMessages(turn, [], [], {});
    const join = role => messages
        .filter(m => m.role === role)
        .map(m => String(m.content || ''))
        .join('\n');
    return { App, pb, systemText: join('system'), userText: join('user') };
}

test('detectFileOutputIntent: 识别导出/另存/保存类请求', () => {
    const { pb } = buildContext('占位');
    assert.strictEqual(pb.detectFileOutputIntent(EXPORT_MSG), true, '导出为CSV并保存到桌面');
    assert.strictEqual(pb.detectFileOutputIntent('导出CSV到桌面'), true);
    assert.strictEqual(pb.detectFileOutputIntent('把这张表另存为CSV'), true);
    assert.strictEqual(pb.detectFileOutputIntent('保存工作簿'), true);
    assert.strictEqual(pb.detectFileOutputIntent('把这张表导出成pdf放到桌面'), true);

    // 反例：纯分析 / 纯格式操作不得被判成文件产出（否则会抢走各自的分支）
    assert.strictEqual(pb.detectFileOutputIntent('分析当前选中的销售数据并给出总结'), false, '分析请求');
    assert.strictEqual(pb.detectFileOutputIntent('把表头加粗并居中'), false, '格式操作请求');
    assert.strictEqual(pb.detectFileOutputIntent('你好'), false, '闲聊');
});

test('导出请求：不得再让模型"先调用 excel_read_range"，且必须点名 excel_export_csv', () => {
    const { systemText, userText } = buildContext(EXPORT_MSG);

    assert.ok(!/先调用 excel_read_range/.test(userText),
        `不得再注入"先调用 excel_read_range"（这正是模型跑去分析数据的直接原因）\n实际:\n${userText}`);
    assert.ok(/不需要先调 excel_read_range/.test(userText), '应显式说明这类请求无需先读整表');
    assert.ok(/excel_export_csv/.test(userText), '用户消息提示里应点名导出工具');
    assert.ok(/必须真正调用工具完成落盘/.test(userText), '应强调必须真正落盘，禁止无工具调用就宣称完成');

    assert.ok(/文件产出/.test(systemText), '应注入"文件产出"专用任务提示');
    assert.ok(/excel_export_csv/.test(systemText), '系统提示里应给出导出工具名');
});

test('分析请求保持原行为：仍先读数据，且不被文件产出分支抢走', () => {
    const { systemText, userText } = buildContext('按销售区域汇总订单金额并给出数据结论');
    assert.ok(/先调用 excel_read_range/.test(userText), '分析类仍应先读数据');
    assert.ok(/数据分析与总结/.test(systemText), '仍注入分析专用提示');
    assert.ok(!/文件产出/.test(systemText), '不得注入文件产出提示');
});

test('格式操作请求保持原行为：仍注入格式与结构操作提示', () => {
    const { systemText, userText } = buildContext('把数据区域美化一下，加边框和居中');
    assert.ok(/格式与结构操作/.test(systemText), '仍注入格式操作专用提示');
    assert.ok(!/先调用 excel_read_range/.test(userText), '格式操作不需要先读数据');
});

test('extractTargetSheet: "保存到桌面"不得被解析成工作表名"桌面"', () => {
    const { App } = buildContext('占位');
    const det = App.ChatV2.intentDetector;

    assert.strictEqual(det.extractTargetSheet('保存到桌面'), null, '桌面是文件位置，不是工作表');
    assert.strictEqual(det.extractTargetSheet(EXPORT_MSG), null, '导出到桌面不得派生出"写入到桌面"任务');
    // 真正的"写入到某工作表"仍要能提取，不能把功能改坏
    assert.strictEqual(det.extractTargetSheet('把结果写入工作表汇总表'), '汇总表');
});

test('配置镜像一致性：内嵌 DEFAULT_CONFIG 与 Config/intent_config.json 的 targetSheet.exclude 必须一致', () => {
    const { App } = buildContext('占位');
    const embedded = App.ChatV2.intentConfig.defs().targetSheet.exclude || [];
    const json = JSON.parse(
        fs.readFileSync(path.join(ROOT, 'Config/intent_config.json'), 'utf8').replace(/^\uFEFF/, ''));
    assert.deepStrictEqual(embedded, json.targetSheet.exclude || [],
        'intent-config.js 的内嵌兜底必须与 Config/intent_config.json 同步（离线时生效的是内嵌那份）');
});
