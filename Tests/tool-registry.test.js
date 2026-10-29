/**
 * ============================================================
 * Tests/tool-registry.test.js — 工具契约测试
 * ============================================================
 * 覆盖：outputSchema 校验、finalizeContent、presentCall/presentResult、
 *       list() 白名单隔离、守卫链、静态卡片回退
 * 运行：node Tests/tool-registry.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

test('tool-registry: outputSchema 契约校验', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;

    tr.register({
        name: 'test_sum', displayName: '求和', description: 'd',
        source: 'builtin',
        outputSchema: { type: 'object', properties: { sum: { type: 'number' } }, required: ['sum'] },
        async execute() { return { ok: true, result: { sum: 42 } }; }
    });

    const v1 = tr.validateToolOutput('test_sum', { sum: 42 });
    assert.strictEqual(v1.valid, true);
    const v2 = tr.validateToolOutput('test_sum', { sum: 'x' });
    assert.strictEqual(v2.valid, false);
});

test('tool-registry: finalizeContent 归一化', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;

    tr.register({
        name: 'test_final', displayName: 'f', description: 'd', source: 'builtin',
        finalizeContent() { return 'finalized'; },
        async execute() { return { ok: true, result: { a: 1 } }; }
    });
    const inv = await tr.invoke('test_final', {}, {});
    assert.strictEqual(inv.content, 'finalized');
});

test('tool-registry: list() 白名单隔离（不泄漏契约字段）', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'secret_tool', displayName: 's', description: 'd', source: 'mcp',
        outputSchema: { type: 'object' },
        finalizeContent() { return 'x'; },
        presentCall() { return {}; },
        async execute() { return { ok: true, result: 1 }; }
    });
    const listed = tr.list().find(t => t.name === 'secret_tool');
    assert.deepStrictEqual(Object.keys(listed).sort(), ['description', 'inputSchema', 'name']);
    assert.ok(!JSON.stringify(listed).includes('outputSchema'));
});

test('tool-registry: presentCall/presentResult 动态契约 + 静态表回退', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;

    // 动态契约优先
    tr.register({
        name: 'test_card', displayName: 'c', description: 'd', source: 'builtin',
        presentCall() { return { card: 'generic', title: '动态调用' }; },
        presentResult() { return { card: 'generic', title: '动态完成' }; },
        async execute() { return { ok: true, result: 1 }; }
    });
    const dynamic = tr.renderCardIntent('test_card', 'call', {});
    assert.strictEqual(dynamic.title, '动态调用');

    // 静态表回退（excel_create_chart 在 TOOL_CARD_RENDERERS 中）
    tr.register({
        name: 'excel_create_chart', displayName: '图表', description: 'c', source: 'mcp',
        async execute() { return { ok: true, result: {} }; }
    });
    const fallback = tr.renderCardIntent('excel_create_chart', 'call', { chartType: 'xlPie' });
    assert.strictEqual(fallback.card, 'chart');
});

test('tool-registry: 守卫链拒绝危险操作', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    let executed = false;
    tr.register({
        name: 'excel_delete_sheet', displayName: '删除', description: 'd', source: 'mcp',
        async execute() { executed = true; return { ok: true, result: 1 }; }
    });
    const res = await tr.invoke('excel_delete_sheet', {}, { turn: { mode: 'standard' } });
    assert.strictEqual(res.ok, false);
    assert.strictEqual(res.guarded, true);
    assert.strictEqual(executed, false, '危险工具不应执行');
});

test('tool-registry: 守卫模式接线 — 全模式确认 + _tpAutoBypass 放行', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;

    // 1) resolveGuardMode 纯函数映射（可离线断言）
    assert.strictEqual(tr.resolveGuardMode(null, 'AUTO'), 'standard');
    assert.strictEqual(tr.resolveGuardMode(null, 'Standard'), 'standard');
    assert.strictEqual(tr.resolveGuardMode(null, 'Deep'), 'deep');
    assert.strictEqual(tr.resolveGuardMode(null, 'expert'), 'expert');
    assert.strictEqual(tr.resolveGuardMode('expert', 'standard'), 'expert', 'turn.mode 优先');
    assert.strictEqual(tr.resolveGuardMode(null, '未知值'), 'standard', '未知回退安全默认');

    // 复用的执行标志工具（刻意用真实危险工具名以触发 dangerous-tools 守卫）
    let ran = false;
    tr.register({
        name: 'excel_run_python_script', displayName: 'Python', description: 'd', source: 'mcp',
        async execute() { ran = true; return { ok: true, result: {} }; }
    });

    // 2) 商业默认：deep/expert 仍须确认（不再凭模式自动放行）
    const viaTurn = await tr.invoke('excel_run_python_script', { _templateMode: false }, { turn: { mode: 'expert' } });
    assert.strictEqual(viaTurn.ok, false, 'turn.mode=expert 仍应拦截待确认');
    assert.strictEqual(viaTurn.guarded, true);
    assert.strictEqual(ran, false);

    // 3) 界面选择器 expert → 同样拦截
    if (!App.config) App.config = {};
    App.config.thinkingMode = 'expert';
    ran = false;
    const viaSelector = await tr.invoke('excel_run_python_script', { _templateMode: false }, { turn: {} });
    assert.strictEqual(viaSelector.ok, false, 'thinkingMode=expert 仍应拦截待确认');
    assert.strictEqual(viaSelector.guarded, true);
    assert.strictEqual(ran, false);

    // 4) 用户显式开启「自动化免确认」→ 放行
    window._tpAutoBypass = true;
    ran = false;
    const viaAuto = await tr.invoke('excel_run_python_script', { _templateMode: false }, { turn: {} });
    assert.strictEqual(viaAuto.ok, true, '_tpAutoBypass 应放行');
    assert.strictEqual(ran, true);
    window._tpAutoBypass = false;

    // 5) 默认 AUTO → 仍拦截（安全兜底未破坏）
    App.config.thinkingMode = 'AUTO';
    ran = false;
    const blocked = await tr.invoke('excel_run_python_script', { _templateMode: false }, { turn: {} });
    assert.strictEqual(blocked.ok, false);
    assert.strictEqual(blocked.guarded, true);
    assert.strictEqual(ran, false, 'AUTO 应仍拦截');
});

test('tool-registry: 弹窗确认后执行 — 确认放行 / 取消拦截 / 无 DOM 兜底拒绝', async () => {
    // 构造"可自动点击指定按钮"的 document stub，驱动 confirmDangerousAction
    // fireSel：匹配的 [data-ac] 值（'[data-ac="ok"]' 或 '[data-ac="cancel"]'）将被"自动点击"；
    //           其他按钮赋值 onclick 时静默，不触发。
    function stubDoc(fireSel) {
        const makeBtn = (isFire) => ({
            onclick: null,
            set onclick(fn) { if (isFire && typeof fn === 'function') fn(); }
        });
        const el = {
            style: {}, classList: { add() {}, remove() {}, contains() { return false; } },
            appendChild() {}, addEventListener() {}, setAttribute() {},
            set innerHTML(v) { this.__html = v; }, get innerHTML() { return this.__html || ''; },
            querySelector(sel) { return makeBtn(!!fireSel && sel === fireSel); },
            querySelectorAll() { return []; }
        };
        return {
            createElement() { return el; },
            body: { appendChild() {} },
            addEventListener() {},
            removeEventListener() {}
        };
    }

    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;

    let executed = 0;
    tr.register({
        name: 'excel_clear_range', displayName: '清空', description: 'd', source: 'mcp',
        async execute() { executed++; return { ok: true, result: 'cleared' }; }
    });
    App.config.thinkingMode = 'AUTO'; // standard，命中守卫

    // 1) 确认 → 放行执行
    global.document = stubDoc('[data-ac="ok"]');
    const confirmed = await tr.invoke('excel_clear_range', {}, { turn: {} });
    assert.strictEqual(confirmed.ok, true, '确认后应放行执行');
    assert.strictEqual(executed, 1, '确认后工具应真正执行');

    // 2) 取消 → 拦截，不执行
    executed = 0;
    global.document = stubDoc('[data-ac="cancel"]');
    const cancelled = await tr.invoke('excel_clear_range', {}, { turn: {} });
    assert.strictEqual(cancelled.ok, false, '取消后不应执行');
    assert.strictEqual(cancelled.guarded, true);
    assert.strictEqual(cancelled.cancelled, true, '取消应标记 cancelled');
    assert.strictEqual(executed, 0);

    // 3) 无 DOM（headless 环境）→ fail-safe 拒绝
    executed = 0;
    delete global.document;
    const noDom = await tr.invoke('excel_clear_range', {}, { turn: {} });
    assert.strictEqual(noDom.ok, false);
    assert.strictEqual(noDom.guarded, true);
    assert.strictEqual(executed, 0, '无 DOM 时应安全拒绝');
});

test('tool-registry: dangerAlwaysAsk 开启时本轮 allow 不复用', async () => {
    function stubDoc(fireSel) {
        const makeBtn = (isFire) => ({
            onclick: null,
            set onclick(fn) { if (isFire && typeof fn === 'function') fn(); }
        });
        const el = {
            style: {}, classList: { add() {}, remove() {}, contains() { return false; } },
            appendChild() {}, addEventListener() {}, setAttribute() {},
            set innerHTML(v) { this.__html = v; }, get innerHTML() { return this.__html || ''; },
            querySelector(sel) { return makeBtn(!!fireSel && sel === fireSel); },
            querySelectorAll() { return []; }
        };
        return {
            createElement() { return el; },
            body: { appendChild() {} },
            addEventListener() {},
            removeEventListener() {}
        };
    }

    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    let executed = 0;
    tr.register({
        name: 'excel_clear_range', displayName: '清空', description: 'd', source: 'mcp',
        async execute() { executed++; return { ok: true, result: 'cleared' }; }
    });
    if (!App.config) App.config = {};
    App.config.thinkingMode = 'AUTO';
    App.config.dangerAlwaysAsk = true;

    const turn = {};
    global.document = stubDoc('[data-ac="ok"]');
    const first = await tr.invoke('excel_clear_range', {}, { turn });
    assert.strictEqual(first.ok, true);
    assert.strictEqual(executed, 1);

    // 第二次：始终询问 → 无 DOM 时 fail-safe 拒绝（说明仍走了确认，未直接沿用 allow）
    executed = 0;
    delete global.document;
    const second = await tr.invoke('excel_clear_range', {}, { turn });
    assert.strictEqual(second.ok, false, '始终询问时不应沿用 allow 直接放行');
    assert.strictEqual(second.guarded, true);
    assert.strictEqual(executed, 0);

    // 关闭始终询问：确认后第二次应沿用 allow
    App.config.dangerAlwaysAsk = false;
    const turn2 = {};
    executed = 0;
    global.document = stubDoc('[data-ac="ok"]');
    await tr.invoke('excel_clear_range', {}, { turn: turn2 });
    executed = 0;
    delete global.document;
    const reused = await tr.invoke('excel_clear_range', {}, { turn: turn2 });
    assert.strictEqual(reused.ok, true, '关闭始终询问后应沿用本轮 allow');
    assert.strictEqual(executed, 1);
});

// 兜底退出：避免 session-store 防抖定时器挂起进程
if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
