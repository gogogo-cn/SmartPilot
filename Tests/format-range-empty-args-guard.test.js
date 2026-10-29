/**
 * ============================================================
 * Tests/format-range-empty-args-guard.test.js —
 *   空参/无效参 excel_format_range 必须在 agent-loop 被拦截（22:21 回归）
 * ============================================================
 * 背景（实测 2026-10-22 22:21）：批量美化链成功后，turn-goal 覆盖度层误判 format 未完成
 * → 追加一轮 → 弱模型回空参 excel_format_range（{}）→ 后端"当前选区"兜底把用户选区
 * 涂预设色，1ms 假成功，面板出现莫名 ❌ "excel_format_range {} - 销售二部"。
 * （覆盖度误引导本身已由 mcp-dispatcher 的 _hardExecCalls 执行登记修复，本测试锁
 *   agent-loop 侧的兜底闸门：即使再有空参调用漏进来，也绝不真正执行。）
 *
 * 判定规则：调用必须带 format/alignment/markLiteral/conditionalFormat/autoRange 之一；
 * 全都没有（含纯 {}、纯 {range}/{sheet}）即无效。批量链的 format='{}'（字符串）有
 * payload，不算空参，不得误拦。
 *
 * 运行：node --test Tests/format-range-empty-args-guard.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

const { App } = loadChatV2(DEFAULT_ORDER);
const util = App.ChatV2.toolRegistryUtil;

// ══ 第一组：isMeaninglessFormatRangeArgs 单元判定 ══
test('util 导出 isMeaninglessFormatRangeArgs', () => {
    assert.strictEqual(typeof util.isMeaninglessFormatRangeArgs, 'function');
});

test('空对象 / null / 非对象 → 无效', () => {
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({}), true);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs(null), true);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs(undefined), true);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs('{}'), true);
});

test('纯 {range} / {sheet} / {range,sheet} → 无效（不构成任何格式化操作）', () => {
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ range: 'A1:G21' }), true);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ sheet: '销售二部' }), true);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ range: 'A1:G21', sheet: '销售二部' }), true);
});

test('带 format（含批量链的 format="{}" 字符串）→ 有效，不得误拦', () => {
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ format: '{}' }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ format: { bold: true } }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ range: 'A1:G21', format: '{}' }), false);
});

test('带 alignment / markLiteral / conditionalFormat / autoRange → 有效', () => {
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ alignment: { horizontal: 'center' } }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ markLiteral: true }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ conditionalFormat: { type: 'dataBar' } }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ autoRange: true }), false);
    assert.strictEqual(util.isMeaninglessFormatRangeArgs({ autoRange: false }), true, 'autoRange=false 不等于 true，仍视为无效');
});

// ══ 第二组：agent-loop 确实接上了闸门（结构性断言） ══
test('agent-loop 用共享判定拦截空参 format_range，且拦截后返回失败不再执行', () => {
    const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');
    const src = fs.readFileSync(path.join(BASE, 'agent-loop.js'), 'utf8');

    // 闸门条件：excel_format_range + 共享判定（含回退分支）
    assert.ok(/tc\.name === 'excel_format_range'/.test(src), 'agent-loop 缺少 format_range 闸门条件');
    assert.ok(src.includes('isMeaninglessFormatRangeArgs'), 'agent-loop 未接入共享判定 isMeaninglessFormatRangeArgs');
    assert.ok(/!\s*App\.ChatV2\.toolRegistryUtil\.isMeaninglessFormatRangeArgs\(tc\.args\)/.test(src),
        '闸门未对共享判定取反（有效调用应放行）');

    // 拦截分支：手工构造失败卡片 + result.ok=false（与 OCR 闸门同构，不进执行分支）
    const gateIdx = src.indexOf("tc.name === 'excel_format_range'");
    const slice = src.slice(gateIdx, gateIdx + 4000);
    assert.ok(slice.includes("status: 'failed'"), '拦截分支必须落失败卡片');
    assert.ok(/ok:\s*false/.test(slice), '拦截结果必须 ok:false，禁止继续执行');
    assert.ok(slice.includes('errorReason'), '拦截卡片必须有 errorReason 说明');
});
