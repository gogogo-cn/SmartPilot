/**
 * ============================================================
 * Tests/merge-truncate.test.js — 合并结果回喂 LLM 前的确定性摘要（治本）
 * ============================================================
 * 背景（2026-10-09 用户反馈 + 本次"治本"需求）：
 *   excel_merge_sheets 后端返回的是标准 JSON（含 sourceSheets/mergedRows/data 矩阵），
 *   agent-loop 通过 truncateResult 把它作为 role:'tool' 消息回喂 LLM。
 *   模型看到完整 payload 后会"复述"成无逗号裸 JSON 塞进正文（治标层已拦截，但仍会泄漏）。
 *
 *   治本：在 truncateResult 中为 excel_merge_sheets 增加特化分支，
 *   回喂前强制把合并结果压成确定性纯文本摘要（与 read_range/clean_data 特化同范式），
 *   模型既看不到完整 JSON，也无从复读。这是真正的"源头截断"。
 *
 * 运行：node --test Tests/merge-truncate.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

// 默认加载集已含 agent-loop / tool-registry.util / chat-orchestrator
const { App } = loadChatV2();
const al = App.ChatV2.agentLoop;

/** 后端 ExecuteMergeSheets 的真实返回体（Services/Mcp/ExcelMcpTools.SheetOps.SplitMerge.vb 用 JsonConvert 带逗号+缩进） */
function mergePayload() {
    return {
        success: true,
        sourceSheets: ['产品A', '产品B'],
        mergedRows: 2,
        dataRows: 2,
        headerPreserved: true,
        usedSelectedSheets: true,
        data: [
            ['日期', '地区', '部门', '产品', '销售额', '客户类型'],
            ['2024-01', '华南', '销售一部', '产品A', 45300, '政府客户'],
            ['2024-02', '华南', '销售一部', '产品B', 31500, '政府客户']
        ],
        steps: ['合并 2 张表'],
        writtenTo: '合并结果',
        targetSheet: '合并结果',
        cleanedEmptySheets: 0
    };
}

test('truncateResult 已导出（agent-loop 回喂截断层）', () => {
    assert.strictEqual(typeof al.truncateResult, 'function', 'truncateResult 应可访问');
});

test('治本：合并结果（对象形态）回喂前被压成确定性纯文本摘要，杜绝 LLM 复读裸 JSON', () => {
    const fed = al.truncateResult(mergePayload(), 'excel_merge_sheets', false);
    assert.ok(typeof fed === 'string' && fed.length > 0, '应返回非空字符串');
    // ① 不得再是合并 JSON 回声（模型正是从它学成复读）
    assert.ok(!/^\s*\{[\s\S]*"sourceSheets"/.test(fed), '回喂内容不应是合并 JSON 回声');
    // ② 不得携带完整 data 矩阵（诱发复读的元凶）
    assert.ok(!fed.includes('"data"') && !fed.includes('政府客户'), '回喂内容不应含 data 矩阵原始数据');
    // ③ 摘要应保留关键信息（来源表 + 合并/行数），让模型仍能如实总结
    assert.ok(fed.includes('产品A') && fed.includes('产品B'), '摘要应保留来源表信息');
    assert.ok(/合并|\d+\s*行/.test(fed), '摘要应含合并/行数信息');
    // ④ 不得残留渲染层 HTML（回喂 LLM 须纯文本，与 read_range/clean_data 特化一致）
    assert.ok(!/<\s*br\s*\/?>|<span/.test(fed), '回喂内容不应残留 HTML 标签');
});

test('治本：合并结果（字符串形态，如已序列化的 payload）同样被摘要化', () => {
    const fed = al.truncateResult(JSON.stringify(mergePayload()), 'excel_merge_sheets', false);
    assert.ok(typeof fed === 'string' && fed.length > 0);
    assert.ok(!/^\s*\{[\s\S]*"sourceSheets"/.test(fed), '字符串形态也应被摘要化，不得是 JSON 回声');
    assert.ok(fed.includes('产品A'), '摘要应保留来源表');
});

test('不误伤：非合并工具结果原样回喂（行为与改动前一致）', () => {
    const other = { success: true, value: 42 };
    const fed = al.truncateResult(other, 'excel_read_range', false);
    // 非合并 → 不走摘要分支，回落通用截断；此处确认仍返回字符串且不抛错
    assert.ok(typeof fed === 'string', '非合并工具应仍返回字符串');
    // 通用截断会把对象 stringify，不含合并文案
    assert.ok(!/已合并/.test(fed), '非合并工具不应出现合并摘要文案');
});

test('边界：merge 工具返回非合并结构时不强行替换（回落通用截断）', () => {
    // toolName 命中但 content 不是合并结果 → replaceJsonEchoWithFriendly 返回 null → 回落原截断
    const fed = al.truncateResult({ ok: true, note: 'nothing about merge' }, 'excel_merge_sheets', false);
    assert.ok(typeof fed === 'string', '极端形态不应抛错');
});
