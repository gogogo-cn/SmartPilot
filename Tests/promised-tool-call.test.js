/**
 * ============================================================
 * Tests/promised-tool-call.test.js — "只说不做"两类失败的回归
 * ============================================================
 * 实测失败（截图）：用户对上一轮格式化的结果说"不够美观"，模型回了一段完全正确的计划
 *   "…设置表头行底色为浅灰色/添加外框和内阁边框/表头居中/自适应列宽…
 *    接下来我将使用 excel_format_range 工具进行美化处理。"
 * 但一个工具都没调，UI 却显示"任务完成 · 未调用工具"，Excel 毫无变化。
 *
 * 两条独立缺陷（任一条都能单独造成该现象）：
 *  ① 工具可见性（主因）：filterToolsForMessage 只吃当前这一句用户消息，而"不够美观"
 *     一个格式化词都不命中（词表只有"美化"，没有"美观"）→ 72→12 个工具，
 *     excel_format_range 被整体剔除，模型物理上无法调用，只能写文字计划。
 *     跟进型消息（"继续"/"再来一次"）同样会丢掉上一轮正在用的工具体系。
 *  ② 承诺未执行（兜底）：agent-loop 的"光说不做"引导链全部以 userOpIntent 为前提，
 *     而 EXCEL_OPERATION_INTENT_RE 同样没有"美观"→ userOpIntent=false，
 *     整套引导/兜底被"无关问答护栏"设计性跳过，直接判定 no_tool_calls 收尾；
 *     且 FAKE_EXECUTION_PATTERN 只认过去式"已X"，不认未来式"我将X"。
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('node:fs');
const path = require('node:path');
const { loadChatV2 } = require('./helpers/harness');

const ROOT = path.join(__dirname, '..');

/** 真实工具清单（与运行时下发的一致） */
function realTools() {
    const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, 'Config/excel_tools.json'), 'utf8').replace(/^\uFEFF/, ''));
    return cfg.tools;
}

/** 最小 llmClient 桩；seen 用于回看每一轮真正发给模型的消息 */
function stubLLM(behavior, seen) {
    return {
        stream({ messages }) {
            if (seen) seen.push(messages);
            return (async function* () {
                yield { type: 'model', model: 'test' };
                if (typeof behavior === 'function') yield* behavior(messages);
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}

async function collectEvents(al, runOpts) {
    const out = [];
    for await (const e of al.run(runOpts)) out.push(e);
    return out;
}

/** 截图里模型那段原文（分片流式输出，模拟真实 token 流） */
function planTextBehavior() {
    return [
        { type: 'model', model: 'test' },
        { type: 'token', delta: '为了确保当前选区的表格更加美观，我将执行以下操作：\n' },
        { type: 'token', delta: '设置表头行（第1行）的底色为浅灰色。\n添加外框和内阁边框到所有单元格。\n' },
        { type: 'token', delta: '将表头文字水平居中、垂直居中对齐。\n根据数据内容自动调整列宽。\n' },
        { type: 'token', delta: '接下来我将使用 `excel_format_range` 工具进行美化处理。' },
        { type: 'done', finishReason: 'stop' }
    ];
}

// ══════════════ ① 工具可见性 ══════════════

test('可见性①: "不够美观"必须保留格式化工具族（词表补齐"美观"）', () => {
    const { App } = loadChatV2();
    const all = realTools();
    const kept = App.ChatV2.toolRegistryUtil.filterToolsForMessage(all, '不够美观');
    const names = kept.map(t => String(t.name || t.originalName || '').toLowerCase());

    assert.ok(names.includes('excel_format_range'),
        `"不够美观"下 excel_format_range 必须可见，否则模型无法调用它。实际可见(${kept.length}): ${names.join(', ')}`);
    assert.ok(names.includes('excel_autofit'), '自 adaptive 列宽对应的 autofit 也应可见');
});

test('可见性②: 跟进型消息（"继续"）必须粘性保留上一轮用过的工具', () => {
    const { App } = loadChatV2();
    const util = App.ChatV2.toolRegistryUtil;
    const all = realTools();

    // 不带粘性：这是个纯跟进词，工具体系会丢
    const bare = util.filterToolsForMessage(all, '继续').map(t => String(t.name || ''));
    assert.ok(!bare.includes('excel_format_range'), '前提：裸"继续"本身不含格式化意图（用于衬托粘性必要性）');

    // 带粘性：上一轮调过 excel_format_range → 本轮继续可见
    const sticky = util.filterToolsForMessage(all, '继续', { preferTools: ['excel_format_range'] })
        .map(t => String(t.name || ''));
    assert.ok(sticky.includes('excel_format_range'), '上一轮用过的工具必须粘性保留');

    // 大小写/前缀容错
    const sticky2 = util.filterToolsForMessage(all, '继续', { preferTools: ['EXCEL_FORMAT_RANGE'] })
        .map(t => String(t.name || ''));
    assert.ok(sticky2.includes('excel_format_range'), 'preferTools 应大小写不敏感');
});

// ══════════════ ② 承诺未执行 ══════════════

test('承诺未执行: 计划文字点名工具但零调用 → 引导重试，重试后真正执行', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let invoked = 0;
    tr.register({
        name: 'excel_format_range', displayName: '格式化区域', description: '设置边框底色对齐', source: 'mcp',
        inputSchema: { type: 'object', required: ['range'], properties: { range: { type: 'string' }, format: { type: 'object' } } },
        async execute() { invoked++; return { ok: true, result: '已格式化' }; }
    });
    const fmtTool = tr.get('excel_format_range');

    const seen = [];
    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1) return planTextBehavior();
        if (llmCall === 2) {
            // 第二轮：被守卫纠正后真正发起调用
            return [
                { type: 'model', model: 'test' },
                {
                    type: 'tool_calls',
                    toolCalls: [{ id: 'c1', function: { name: 'excel_format_range', arguments: JSON.stringify({ range: '$A$1:$G$201', format: { borderStyle: 'thin' } }) } }]
                },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        // 第三轮起：正常收尾（避免无限重调工具）
        return [
            { type: 'model', model: 'test' },
            { type: 'token', delta: '已经完成美化。' },
            { type: 'done', finishReason: 'stop' }
        ];
    }, seen);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '不够美观', kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '不够美观' }],
        tools: [fmtTool],   // 关键：工具本轮可见
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry');
    assert.ok(retry, `应触发引导重试，实际事件: ${events.map(e => e.type + (e.reason ? ':' + e.reason : '')).join(' | ')}`);
    assert.strictEqual(retry.reason, 'named_tool_commitment_without_call');

    // 关键：不得在重试之前就判定"无工具调用"收尾（那正是"任务完成·未调用工具"的来源）
    const firstLoopDone = events.findIndex(e => e.type === 'loop_done' && e.reason === 'no_tool_calls');
    const retryIdx = events.indexOf(retry);
    assert.ok(firstLoopDone === -1 || retryIdx < firstLoopDone, '不得先以 no_tool_calls 收尾');

    // 重试提示必须点名工具 + 明说"没有产生任何工具调用"
    const retryMsgs = seen[seen.length - 1] || [];
    const hint = retryMsgs.filter(m => m.role === 'system').map(m => String(m.content || '')).join('\n');
    assert.ok(/excel_format_range/.test(hint), '重试提示应点名该工具');
    assert.ok(/没有产生任何工具调用/.test(hint), '重试提示应明确指出未调用工具');

    // 重试后工具真的被执行
    const resultEvt = events.find(e => e.type === 'tool_call_result' && e.toolName === 'excel_format_range');
    assert.ok(resultEvt, '重试后必须真正执行 excel_format_range');
    assert.ok(invoked >= 1, '工具 execute 必须被调用');
});

test('承诺未执行（工具不可见）: 提示模型改用真实可用工具，不诱导调用不存在的工具', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    // 注册但【不】放进本轮 tools → 模拟被可见性过滤掉的情形
    tr.register({
        name: 'excel_autofit', displayName: '自适应列宽', description: '自适应列宽', source: 'mcp',
        async execute() { return { ok: true, result: 'ok' }; }
    });
    const readTool = { name: 'excel_read_range', originalName: 'excel_read_range', description: '读取数据', inputSchema: { type: 'object', properties: {} } };

    const seen = [];
    App.ChatV2.llmClient = stubLLM(planTextBehavior, seen);

    const stTurn = App.ChatV2.sessionStore.startTurn({ message: '继续调整', kind: 'text' });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '继续调整' }],
        tools: [readTool],   // excel_format_range 不在其中
        signal: new AbortController().signal
    });

    const retry = events.find(e => e.type === 'force_retry' && e.reason === 'named_tool_commitment_without_call');
    assert.ok(retry, '工具不可见时同样要引导重试（不能静默收尾）');
    const hint = (seen[seen.length - 1] || []).filter(m => m.role === 'system').map(m => String(m.content || '')).join('\n');
    assert.ok(/不在你的可用工具列表中/.test(hint), '应告知该工具本轮不可用');
    assert.ok(/excel_read_range/.test(hint), '应列出真实可用工具');
});

test('反向护栏: 建议语气/过去式陈述不得被误判成"承诺未执行"', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;
    tr.register({
        name: 'excel_format_range', displayName: '格式化区域', description: '格式化', source: 'mcp',
        async execute() { return { ok: true, result: 'ok' }; }
    });
    const fmtTool = tr.get('excel_format_range');

    const CASES = [
        // 建议语气（第二人称，不是自己承诺）
        '你可以用 excel_format_range 来设置边框和底色，也可以直接用菜单里的"设置单元格格式"。',
        // 过去式陈述（假执行另有一套 FAKE_EXECUTION_PATTERN 处理，不该走本守卫）
        '我已调用 excel_read_range 读取了数据，结果如下：A1:G1 共 7 个单元格，内容为表头。'
    ];

    for (const text of CASES) {
        App.ChatV2.llmClient = stubLLM(() => [
            { type: 'model', model: 'test' },
            { type: 'token', delta: text },
            { type: 'done', finishReason: 'stop' }
        ]);
        const stTurn = App.ChatV2.sessionStore.startTurn({ message: '这个怎么弄', kind: 'text' });
        const events = await collectEvents(al, {
            turn: stTurn,
            messages: [{ role: 'user', content: '这个怎么弄' }],
            tools: [fmtTool],
            signal: new AbortController().signal
        });
        assert.ok(!events.some(e => e.type === 'force_retry' && e.reason === 'named_tool_commitment_without_call'),
            `不得误判为承诺未执行：${text.slice(0, 30)}…`);
    }
});
