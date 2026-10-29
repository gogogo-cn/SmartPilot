/**
 * ============================================================
 * Tests/history-sanitize.test.js — 跨轮历史正文清洗回归
 * ============================================================
 * 背景（2026-10-10 00:32 实测）：
 *   第 1 轮「分析选中数据的描述统计」走 python-pipeline，assistant 正文里带了
 *   ① 完整 Python 脚本的围栏代码块 ②「✅ 分析完成！共完成 1 项分析」收尾套话。
 *   第 2 轮「使用饼图按部门展示销售额」时，这些文本被 getHistoryForLLM 原样喂回模型，
 *   弱模型照抄上一轮套路 → 又跑一遍脚本分析、又写一段"关键发现/结论"，饼图走偏。
 *
 * 清洗原则（store.sanitizeHistoryContent）：
 *   只删"上一轮的实现细节"（脚本代码块 / 管道收尾套话 / 超长行内代码），
 *   保留"上一轮的结论摘要"（关键发现、结论、汇总表格都是有效上下文，下一轮可能真要引用）。
 *   只作用于喂 LLM 的历史副本，不动 STATE.history；user 原文一字不动。
 *
 * 运行：node Tests/history-sanitize.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadStore() {
    const window = {};
    window.App = { ChatV2: {} };
    const ls = {};
    const localStorage = {
        getItem: k => (Object.prototype.hasOwnProperty.call(ls, k) ? ls[k] : null),
        setItem: (k, v) => { ls[k] = String(v); },
        removeItem: k => { delete ls[k]; }
    };
    window.localStorage = localStorage;
    global.localStorage = localStorage;
    // 生产里 app-history.js 提供全局 AppHistory；harness 不加载它时仍要保证
    // saveToLegacyHistory 的防抖回调不会因未声明标识符抛 ReferenceError。
    global.AppHistory = { addToHistory() {} };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/session-store.js'), 'utf8'));
    return App.ChatV2.sessionStore;
}

const PY_ANALYSIS_REPLY = [
    '### 分组统计结果',
    '| 分组 | 求和 |',
    '|------|------|',
    '| 销售一部 | 480788 |',
    '| 销售二部 | 392341 |',
    '',
    '```python',
    'import pandas as pd',
    'df = pd.DataFrame(rows)',
    'print(df.groupby("部门")["销售额"].sum())',
    '```',
    '',
    '### 关键发现',
    '- **最高求和**：销售一部 480788',
    '',
    '### 结论',
    '销售一部贡献最大。',
    '',
    '✅ **分析完成！** 共完成 1 项分析'
].join('\n');

test('清洗: 剥离围栏代码块（含 ```python 语言标注）', () => {
    const s = loadStore().sanitizeHistoryContent(PY_ANALYSIS_REPLY);
    assert.ok(!s.includes('import pandas'), '不应残留脚本内容；实际=' + JSON.stringify(s));
    assert.ok(!s.includes('```'), '不应残留代码围栏；实际=' + JSON.stringify(s));
    assert.ok(s.includes('[代码块已省略]'), '应留下占位提示（避免模型以为内容凭空消失）');
});

test('清洗: 剥离管道收尾套话「✅ 分析完成！共完成 1 项分析」', () => {
    const s = loadStore().sanitizeHistoryContent(PY_ANALYSIS_REPLY);
    assert.ok(!s.includes('分析完成'), '不应残留"分析完成"套话；实际=' + JSON.stringify(s));
    assert.ok(!s.includes('共完成 1 项分析'), '不应残留计数套话；实际=' + JSON.stringify(s));
});

test('清洗: 保留结论摘要（分组统计结果 / 关键发现 / 结论）', () => {
    const s = loadStore().sanitizeHistoryContent(PY_ANALYSIS_REPLY);
    assert.ok(s.includes('分组统计结果'), '汇总表格是有效上下文，应保留；实际=' + JSON.stringify(s));
    assert.ok(s.includes('关键发现'), '关键发现是有效上下文，应保留；实际=' + JSON.stringify(s));
    assert.ok(s.includes('销售一部贡献最大'), '结论应保留；实际=' + JSON.stringify(s));
});

test('清洗: 超长行内代码剥除，短标识符（列名）不受影响', () => {
    const s = loadStore().sanitizeHistoryContent(
        '说明：`df.groupby("部门")["销售额"].sum().reset_index()` 用于分组，列 `部门` 与 `销售额` 参与。'
    );
    assert.ok(s.includes('[代码片段已省略]'), '超长行内代码应剥除；实际=' + JSON.stringify(s));
    assert.ok(s.includes('`部门`'), '短标识符（列名）必须保留；实际=' + JSON.stringify(s));
});

test('清洗: 连续空行压实 + 首尾空白裁剪', () => {
    const s = loadStore().sanitizeHistoryContent('结论一。\n\n\n\n\n结论二。   ');
    assert.strictEqual(s, '结论一。\n\n结论二。', '实际=' + JSON.stringify(s));
});

test('清洗: 仅代码块时留占位符，不再喂回脚本', () => {
    const s = loadStore().sanitizeHistoryContent('```python\nprint(1)\n```');
    assert.strictEqual(s, '[代码块已省略]', '实际=' + JSON.stringify(s));
});

test('清洗: 掏空安全阀 —— 只剩套话时回退原文，绝不丢上下文', () => {
    const src = '✅ **分析完成！** 共完成 1 项分析\n';
    const s = loadStore().sanitizeHistoryContent(src);
    assert.strictEqual(s, src, '整段被掏空时应回退原文；实际=' + JSON.stringify(s));
});

test('清洗: 非字符串输入原样返回', () => {
    const f = loadStore().sanitizeHistoryContent;
    assert.strictEqual(f(null), null);
    assert.strictEqual(f(undefined), undefined);
    assert.strictEqual(f(''), '');
});

test('e2e: getHistoryForLLM 的 assistant 正文已清洗，user 原文一字不动', () => {
    const ss = loadStore();

    const t1 = ss.startTurn({ message: '分析选中数据的描述统计', kind: 'text', context: {} });
    ss.appendDelta(t1.id, 'content', PY_ANALYSIS_REPLY);
    ss.completeTurn(t1.id, {});

    // 第 2 轮：独立的新命令
    const t2 = ss.startTurn({ message: '使用饼图按部门展示销售额', kind: 'text', context: {} });
    const h = ss.getHistoryForLLM();

    // 当前进行中的 turn 不进历史
    assert.ok(!h.some(m => m.content === '使用饼图按部门展示销售额'),
        '当前 turn 不应出现在历史里；实际=' + JSON.stringify(h));

    const userMsg = h.find(m => m.role === 'user');
    assert.strictEqual(userMsg && userMsg.content, '分析选中数据的描述统计',
        'user 原文必须一字不动；实际=' + JSON.stringify(userMsg));

    const assistant = h.find(m => m.role === 'assistant');
    assert.ok(assistant, '应有上一轮 assistant 消息；实际=' + JSON.stringify(h));
    assert.ok(!assistant.content.includes('import pandas'),
        '不应把上一轮脚本喂回模型；实际=' + JSON.stringify(assistant.content));
    assert.ok(!assistant.content.includes('共完成 1 项分析'),
        '不应把收尾套话喂回模型；实际=' + JSON.stringify(assistant.content));
    assert.ok(assistant.content.includes('分组统计结果'),
        '结论摘要必须保留；实际=' + JSON.stringify(assistant.content));
});
