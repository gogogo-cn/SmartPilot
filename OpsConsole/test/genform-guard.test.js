'use strict';
/**
 * 生成激活码表单守栏。
 *
 * 背景（2026-10-20 实测）：用户截图「方案」下拉空白。根因是 loadGenForm 与初始化
 * 的 /api/pricing 请求都没有 .catch，服务没起（或用 file:// 直接打开页面）时
 * 静默失败 → 下拉永远空白，且无任何提示。
 *
 * 修复：方案下拉用本地兜底（PLAN_FALLBACK，与 lib/pricing.js DEFAULTS 一致）
 * 立即填充；请求失败 toast 提示；file:// 打开时页面顶部出红色警告条。
 * 本测试防止改回"静默空白"。
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'public', 'index.html'), 'utf8');

test('方案下拉有本地兜底（请求失败也不空白）', () => {
    assert.ok(html.indexOf('PLAN_FALLBACK') >= 0, '应定义 PLAN_FALLBACK 兜底方案');
    assert.ok(html.indexOf('fillPlanSelect(PLAN_FALLBACK)') >= 0, '初始化时应立即用兜底填充下拉');
    // 兜底键必须与 lib/pricing.js 的 DEFAULTS 键一致
    const pricing = fs.readFileSync(path.join(__dirname, '..', 'lib', 'pricing.js'), 'utf8');
    for (const plan of ['pro_yearly', 'pro_monthly', 'team']) {
        assert.ok(pricing.indexOf(plan) >= 0, 'pricing.js 应含 ' + plan);
        assert.ok(html.indexOf(plan) >= 0, '前端兜底应含 ' + plan);
    }
});

test('定价请求失败必须有提示，不允许静默', () => {
    const gen = html.slice(html.indexOf('function loadGenForm'), html.indexOf('function submitGen'));
    assert.ok(gen.indexOf('.catch(') >= 0, 'loadGenForm 必须有 .catch');
    assert.ok(gen.indexOf('toast(') >= 0, '失败时必须 toast 提示');

    const submit = html.slice(html.indexOf('function submitGen'), html.indexOf('function loadGenForm') > 0 && html.indexOf('// ---------- 事件') > 0 ? html.indexOf('// ---------- 事件') : html.length);
    assert.ok(submit.indexOf('.catch(') >= 0, 'submitGen 必须有 .catch');
});

test('下拉文案写全：金额 + 计费方式 + 周期，并实时算金额', () => {
    assert.ok(html.indexOf('function priceText') >= 0, '应有 priceText 生成单价文案');
    const fn = html.slice(html.indexOf('function priceText'), html.indexOf('function fillPlanSelect'));
    assert.ok(fn.indexOf('perSeat') >= 0, '单价文案应区分是否按席位');
    assert.ok(fn.indexOf('period') >= 0, '单价文案应带周期（月/年）');

    // 定价表必须提供 period，否则显示不出来
    const cfg = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'config.json'), 'utf8'));
    for (const p of Object.keys(cfg)) {
        assert.ok(cfg[p].period, p + ' 应配置 period（月/年）');
    }

    // 下方说明行 + 实时刷新
    assert.ok(html.indexOf('g-plan-hint') >= 0, '下拉下方应有说明行');
    assert.ok(html.indexOf('function updatePlanHint') >= 0, '应有 updatePlanHint');
    assert.ok(/\['g-plan','g-seats','g-days'\]/.test(html), '方案/席位/天数变化都应实时刷新说明');
});

test('按席位方案（团队版）必须拆成「每人一码」', () => {
    // 背景（2026-10-21 用户拍板）：多席位订单发 N 张独立激活码而非一张共用码。
    // 共用码时任一人退款会连坐全队（吊销唯一那张 key）；一客一码则退款天然隔离且可追溯。
    // 两边总价相同（129 × 人数）。本测试防止被改回"一张码写 seats=N"。
    const hint = html.slice(html.indexOf('function updatePlanHint'), html.indexOf('function loadGenForm'));
    assert.ok(hint.indexOf('c.perSeat') >= 0, '说明行应区分按席位方案');
    assert.ok(/cntEl\.value = seats/.test(hint) || /g-count'\)\.value = seats/.test(hint),
        '按席位时「生成数量」应联动为人数');
    assert.ok(hint.indexOf('个独立激活码') >= 0, '应明确告知会生成 N 个独立激活码');

    const submit = html.slice(html.indexOf('function submitGen'), html.indexOf('// ---------- 事件'));
    assert.ok(/seats: c\.perSeat \? 1 : seatsInput/.test(submit),
        '提交时必须把 Teams 码的 seats 归一为 1（每人一码）');
    assert.ok(/var count = c\.perSeat \? seatsInput/.test(submit), 'count 应由人数推导');
    assert.ok(/count: count/.test(submit), 'body.count 必须使用推导出的数量');

    // 席位数列的意义要让用户看懂
    assert.ok(html.indexOf('席位 / 人数') >= 0, '席位输入框应标明即人数');
});

test('团队版 5 人起订：前端兜底表带 minSeats，提交前拦截', () => {
    // 2026-10-26 用户拍板：团队版（5人以上）¥129/席位/年。
    // 前端兜底表（/api/pricing 失败时用）必须带 minSeats，否则 file:// 直开时起订量失效。
    assert.ok(/team:\s*\{\s*amount:\s*129/.test(html), '兜底表团队版单价应为 129');
    assert.ok(/minSeats:\s*5/.test(html), '兜底表团队版应带 minSeats=5');
    assert.ok(/pro_monthly:\s*\{\s*amount:\s*19/.test(html), '兜底表月付应为 19');
    // 提交前先给一句人话，别让用户点了提交才看红色 toast
    const submit = html.slice(html.indexOf('function submitGen'), html.indexOf('// ---------- 事件'));
    assert.ok(/minSeats/.test(submit), 'submitGen 应读 minSeats');
    assert.ok(/人起订/.test(submit), '拦截提示要写「人起订」');
    // 说明行也要把起订量说清楚
    const hint = html.slice(html.indexOf('function updatePlanHint'), html.indexOf('function loadGenForm'));
    assert.ok(hint.indexOf('minSeats') >= 0, 'updatePlanHint 应按起订量调整输入框下限');
});

test('file:// 直接打开页面时给出醒目警告', () => {
    assert.ok(html.indexOf("location.protocol === 'file:'") >= 0, '应检测 file:// 协议');
    assert.ok(html.indexOf('127.0.0.1:9600') >= 0, '警告里应给出正确的打开方式');
});
