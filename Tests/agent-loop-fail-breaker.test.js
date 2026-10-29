/**
 * ============================================================
 * Tests/agent-loop-fail-breaker.test.js — 同参数失败熔断回归
 * ============================================================
 * 背景（2026-10-05 22:44 实测）：
 *   PPT 宿主里 LLM 调用 ppt_export_pdf 未传 filePath → 后端报 "filePath 不能为空"，
 *   而 agent-loop 的幂等台账（turn._toolCallLedger / sessionLedger）只记录"成功"的调用，
 *   失败调用没有任何记录 → 弱模型把同一个错误一路重放到 MAX_TOOL_ROUNDS(5)，
 *   UI 弹出 5 张一模一样的失败卡片，模型最后还照着自己的想象编了个文件名说"已导出"。
 *
 * 修复：SAME_CALL_FAIL_LIMIT 连续失败熔断 —— 执行前检查"工具名+参数"签名的连续失败数，
 *      达阈值就跳过重复执行，并回喂"不要再重复/严禁声称已导出"的强指令。
 *
 * 覆盖：① 同参数只真执行 2 次；② 第 3+ 次事件带"重复失败已熔断"；
 *      ③ tool 回喂消息含禁止虚假成功的措辞；④ 成功即清零失败计数。
 * 运行：node Tests/agent-loop-fail-breaker.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

/** 最小 llmClient stub：按调用次数返回 behavior */
function stubLLM(behavior) {
    return {
        stream({ messages }) {
            return (async function* () {
                yield { type: 'model', model: 'test' };
                if (typeof behavior === 'function') yield* behavior(messages);
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };
}

async function collectEvents(al, runOpts) {
    const types = [];
    for await (const e of al.run(runOpts)) types.push(e);
    return types;
}

const FAIL_ARG = JSON.stringify({ filePath: '' });

test('agent-loop: 同一工具同一参数连续失败达阈值后被熔断（不再无限重试）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let executed = 0;
    tr.register({
        name: 'ppt_export_pdf', displayName: '导出 PDF', description: '导出 PDF', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() {
            executed++;
            return { ok: false, error: 'filePath 不能为空' };
        }
    });

    // 弱模型行为：每一轮都重复同一个失败调用，直到耗尽轮数
    // 通过 messages 捕获"回喂给模型的最后一轮上下文"，用于验证熔断指令真的进了 prompt
    let llmCall = 0;
    let lastMessages = [];
    App.ChatV2.llmClient = stubLLM((messages) => {
        llmCall++;
        if (messages && messages.length) lastMessages = messages;
        if (llmCall <= 5) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c' + llmCall, function: { name: 'ppt_export_pdf', arguments: FAIL_ARG } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        return [{ type: 'token', delta: '已导出到桌面。' }, { type: 'done', finishReason: 'stop' }];
    });

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '将演示文稿导出为 PDF',
        kind: 'text',
        context: {}
    });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '将演示文稿导出为 PDF' }],
        tools: [],
        signal: new AbortController().signal
    });

    // 核心断言：真执行不超过 SAME_CALL_FAIL_LIMIT(2) 次
    assert.ok(executed <= 2,
        `同一失败调用不应无限重试，实际执行 ${executed} 次（期望 <= 2）`);

    // 熔断事件必须带明确标记
    const breakerEvents = events.filter(e =>
        e.type === 'tool_call_result' && String(e.error || '').includes('重复失败已熔断'));
    assert.ok(breakerEvents.length > 0,
        '应至少产生一条"重复失败已熔断"的结果事件，实际: ' +
        JSON.stringify(events.filter(e => e.type === 'tool_call_result').map(e => e.error)));

    // 熔断回喂消息必须禁止模型虚假声称成功（从回喂给模型的 tool 消息里查）
    const dump = JSON.stringify(lastMessages || []);
    assert.ok(dump.includes('严禁在未成功的情况下声称'),
        '熔断回喂应包含禁止虚假声称成功的指令');

    console.log(`\n✅ 失败熔断生效：ppt_export_pdf 同参数实际执行 ${executed} 次，后续 ${breakerEvents.length} 次被拦截`);
});

test('agent-loop: 同参数失败后改用不同参数可继续执行（熔断不误伤）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    const seen = [];
    tr.register({
        name: 'ppt_export_pdf', displayName: '导出 PDF', description: '导出 PDF', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute(args) {
            const p = String((args && args.filePath) || '');
            seen.push(p);
            if (!p) return { ok: false, error: 'filePath 不能为空' };
            return { ok: true, result: { success: true, path: p } };
        }
    });

    let llmCall = 0;
    App.ChatV2.llmClient = stubLLM(() => {
        llmCall++;
        if (llmCall === 1 || llmCall === 2) {
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c' + llmCall, function: { name: 'ppt_export_pdf', arguments: FAIL_ARG } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        if (llmCall === 3) {
            // 换成有效绝对路径 —— 签名不同，不应被熔断
            return [
                { type: 'model', model: 'test' },
                { type: 'tool_calls', toolCalls: [{ id: 'c3', function: { name: 'ppt_export_pdf', arguments: JSON.stringify({ filePath: 'C:\\Users\\Administrator\\Desktop\\a.pdf' }) } }] },
                { type: 'done', finishReason: 'tool_calls' }
            ];
        }
        return [{ type: 'token', delta: '已导出。' }, { type: 'done', finishReason: 'stop' }];
    });

    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '导出 PDF', kind: 'text', context: {}
    });
    const events = await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '导出 PDF' }],
        tools: [],
        signal: new AbortController().signal
    });

    const okEvent = events.find(e => e.type === 'tool_call_result' && e.ok === true);
    assert.ok(okEvent, '更换参数后应有一次成功执行（熔断按"工具名+参数"签名，不误伤）');
    assert.ok(seen.includes('C:\\Users\\Administrator\\Desktop\\a.pdf'),
        '有效路径应被真正执行，实际: ' + JSON.stringify(seen));
    console.log(`\n✅ 熔断按参数签名隔离：无效路径失败 ${seen.filter(p => !p).length} 次后，带路径的调用正常执行成功`);
});

test('agent-loop: 熔断后非 persistTools 路径下一轮仍下发 tools（不得误标 toolsExecuted）', async () => {
    const { App } = loadChatV2();
    const al = App.ChatV2.agentLoop;
    const tr = App.ChatV2.toolRegistry;

    let executed = 0;
    tr.register({
        name: 'ppt_export_pdf', displayName: '导出 PDF', description: '导出 PDF', source: 'mcp',
        inputSchema: { type: 'object', properties: {} },
        async execute() {
            executed++;
            return { ok: false, error: 'filePath 不能为空' };
        }
    });

    const toolsOffered = [];
    let llmCall = 0;
    App.ChatV2.llmClient = {
        stream({ tools }) {
            toolsOffered.push(Array.isArray(tools) ? tools.length : -1);
            return (async function* () {
                llmCall++;
                yield { type: 'model', model: 'test' };
                if (llmCall <= 4) {
                    yield {
                        type: 'tool_calls',
                        toolCalls: [{ id: 'c' + llmCall, function: { name: 'ppt_export_pdf', arguments: FAIL_ARG } }]
                    };
                    yield { type: 'done', finishReason: 'tool_calls' };
                } else {
                    yield { type: 'token', delta: '导出失败，请检查路径。' };
                    yield { type: 'done', finishReason: 'stop' };
                }
            })();
        }
    };

    const toolDesc = [{
        name: 'ppt_export_pdf', description: '导出 PDF',
        inputSchema: { type: 'object', properties: {} }
    }];
    const stTurn = App.ChatV2.sessionStore.startTurn({
        message: '将演示文稿导出为 PDF', kind: 'text', context: {}
    });
    await collectEvents(al, {
        turn: stTurn,
        messages: [{ role: 'user', content: '将演示文稿导出为 PDF' }],
        tools: toolDesc,
        signal: new AbortController().signal,
        persistTools: false,
        maxIterations: 5
    });

    assert.ok(executed <= 2, `真执行应被熔断，实际 ${executed}`);
    // 熔断发生后至少还有一轮 LLM 调用仍带 tools（长度>0）；若误标 toolsExecuted 会变成 0
    const afterBreaker = toolsOffered.slice(2); // 第 3 次起多为熔断 skip 轮
    assert.ok(afterBreaker.some(n => n > 0),
        '熔断后下一轮仍应下发 tools，实际 tools 长度序列=' + JSON.stringify(toolsOffered));
    console.log(`\n✅ 熔断后 tools 仍可见：序列=${JSON.stringify(toolsOffered)}，真执行=${executed}`);
});