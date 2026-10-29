/**
 * ============================================================
 * Tests/llm-client-runaway-guard.test.js — 模型失控输出防护守卫
 * ============================================================
 * 背景（2026-10-18 实测）：用户发"读取图纸中的尺寸"+PDF 附件，qwen2.5:7b 进入
 * 退化循环，reasoning 里 "1,1,2,2,3,3…"无限列举，生成 14397 token、耗时 3分10秒、
 * truncated=1（Ollama 无 num_predict 时默认生成到 num_ctx 上限）。期间前端只有
 * "正在分析你的需求"，用户感知"无限输出/没有过程渲染"，最终只能手动中止。
 *
 * 本套件守三件事：
 *   1. buildBody 对 Ollama 必须显式设置 num_predict（默认取 maxTokens，可覆盖）
 *   2. stream() 的 chunk 数硬上限：即使服务端忽略 num_predict，客户端也强制收尾
 *      （finishReason=runaway_truncated），且绝不会明显超出上限
 *   3. stream() 的 stall 看门狗：读流连续 N 秒无数据即报 STREAM_STALL，
 *      不再无限静默（复现"6分钟零反馈"场景的修复）
 *
 * ⚠️ 桩必须复现真实失败形态：reader.read() 用永不 settle 的 promise 模拟
 * Ollama 冷加载/被占满时的静默；cancel() 必须让挂起的 read() 以 {done:true}
 * 返回——这正是 Chromium 的真实语义，否则测试会挂死而不是验证看门狗。
 * ============================================================
 */
'use strict';

const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const ENC = new TextEncoder();

/** Ollama /api/chat 流式 NDJSON 行（真实格式：无 data: 前缀） */
function ollamaChunk(text, model) {
    return JSON.stringify({ model: model || 'qwen2.5:7b', message: { role: 'assistant', content: text }, done: false }) + '\n';
}

/** Response-like 桩：body.reader 依次吐出 chunks，然后 {done:true} */
function makeStreamResponse(chunkTexts) {
    let i = 0;
    const reader = {
        read() {
            if (i < chunkTexts.length) {
                return Promise.resolve({ done: false, value: ENC.encode(chunkTexts[i++]) });
            }
            return Promise.resolve({ done: true, value: undefined });
        },
        cancel() { i = chunkTexts.length; return Promise.resolve(); }
    };
    return {
        ok: true,
        status: 200,
        headers: { get: (k) => (String(k).toLowerCase() === 'content-type' ? 'text/event-stream' : null) },
        body: { getReader: () => reader }
    };
}

/** Response-like 桩：read() 永不返回；cancel() 时以 {done:true} 解脱（Chromium 语义） */
function makeStalledResponse() {
    let cancelResolve = null;
    const pending = new Promise((resolve) => { cancelResolve = resolve; });
    const reader = {
        read() { return pending.then(() => ({ done: true, value: undefined })); },
        cancel() { cancelResolve(); return Promise.resolve(); }
    };
    return {
        ok: true,
        status: 200,
        headers: { get: (k) => (String(k).toLowerCase() === 'content-type' ? 'text/event-stream' : null) },
        body: { getReader: () => reader }
    };
}

test('buildBody: Ollama 必须显式设置 num_predict（失控根因防护）', () => {
    const { App } = loadChatV2();
    const { buildBody } = App.ChatV2.llmClient;

    // 1) 默认取 config.maxTokens（用户配置里通常 4096）
    const body1 = buildBody({
        messages: [{ role: 'user', content: '读取图纸中的尺寸' }],
        config: { apiType: 'ollama', model: 'qwen2.5:7b', maxTokens: 4096 }
    });
    assert.strictEqual(body1.options.num_predict, 4096, 'num_predict 必须显式设置，不能留给 Ollama 默认 -1（无限）');
    assert.strictEqual(body1.options.num_ctx, 16384, 'num_ctx 保持 16K 不回退');

    // 2) config.options.num_predict 可覆盖
    const body2 = buildBody({
        messages: [{ role: 'user', content: 'hi' }],
        config: { apiType: 'ollama', model: 'qwen2.5:7b', maxTokens: 4096, options: { num_predict: 2048 } }
    });
    assert.strictEqual(body2.options.num_predict, 2048);

    // 3) 连 maxTokens 都没配时也要有兜底值（不能是 undefined/-1）
    const body3 = buildBody({
        messages: [{ role: 'user', content: 'hi' }],
        config: { apiType: 'ollama', model: 'qwen2.5:7b' }
    });
    assert.strictEqual(body3.options.num_predict, 4096);
    assert.ok(body3.options.num_predict > 0 && Number.isFinite(body3.options.num_predict));

    // 4) OpenAI 兼容分支不受影响（走 max_tokens）
    const body4 = buildBody({
        messages: [{ role: 'user', content: 'hi' }],
        config: { apiType: 'openai', model: 'gpt-4o', maxTokens: 2048 }
    });
    assert.strictEqual(body4.max_tokens, 2048);
    assert.strictEqual(body4.options, undefined, 'OpenAI 分支不应出现 Ollama options');
});

test('stream: chunk 数失控硬上限——服务端忽略 num_predict 时客户端强制收尾', async () => {
    const { App } = loadChatV2();
    const originalFetch = global.fetch;
    // 模拟失控：500 个 chunk 还在继续吐（真实案例是 14397 token）
    const runawayChunks = [];
    for (let i = 0; i < 500; i++) runawayChunks.push(ollamaChunk('尺寸数据 ' + i + ' '));
    let fetchCalls = 0;
    global.fetch = async () => { fetchCalls++; return makeStreamResponse(runawayChunks); };

    try {
        const events = [];
        const CAP = 20; // 用 config.options.maxStreamChunks 压小上限，加快测试
        for await (const evt of App.ChatV2.llmClient.stream({
            messages: [{ role: 'user', content: '读取图纸中的尺寸' }],
            config: { apiType: 'ollama', model: 'qwen2.5:7b', maxTokens: 4096, options: { maxStreamChunks: CAP } }
        })) {
            events.push(evt);
        }

        const tokenEvents = events.filter(e => e.type === 'token');
        const doneEvents = events.filter(e => e.type === 'done');
        assert.strictEqual(doneEvents.length, 1, '必须以 done 收尾');
        assert.strictEqual(doneEvents[0].finishReason, 'runaway_truncated', '截断原因必须是 runaway_truncated');
        assert.ok(tokenEvents.length <= CAP, `token 事件数(${tokenEvents.length})不得超过上限 ${CAP}`);
        assert.ok(tokenEvents.length >= CAP - 1, '应贴在上限处收尾，而不是提前误伤');
        assert.ok(!events.some(e => e.type === 'error'), '失控截断不是错误，是正常收尾');
        assert.strictEqual(fetchCalls, 1, '不应触发非流式回退（已有部分输出）');
    } finally {
        global.fetch = originalFetch;
    }
});

test('stream: 正常短流不受熔断影响（防误伤守卫）', async () => {
    const { App } = loadChatV2();
    const originalFetch = global.fetch;
    global.fetch = async () => makeStreamResponse([
        ollamaChunk('图纸尺寸'),
        ollamaChunk('：10.45mm'),
        JSON.stringify({ model: 'qwen2.5:7b', message: { role: 'assistant', content: '' }, done: true }) + '\n'
    ]);

    try {
        const events = [];
        for await (const evt of App.ChatV2.llmClient.stream({
            messages: [{ role: 'user', content: '读取图纸中的尺寸' }],
            config: { apiType: 'ollama', model: 'qwen2.5:7b', maxTokens: 4096, options: { maxStreamChunks: 20 } }
        })) {
            events.push(evt);
        }
        const done = events.find(e => e.type === 'done');
        assert.ok(done, '应正常收到 done');
        assert.strictEqual(done.finishReason, 'stop', '正常流的 finishReason 不应被改成截断');
        const text = events.filter(e => e.type === 'token').map(e => e.delta).join('');
        assert.strictEqual(text, '图纸尺寸：10.45mm');
    } finally {
        global.fetch = originalFetch;
    }
});

test('stream: stall 看门狗——连续无数据不再无限静默', async () => {
    const { App } = loadChatV2();
    const originalFetch = global.fetch;
    global.fetch = async () => makeStalledResponse();

    try {
        const events = [];
        for await (const evt of App.ChatV2.llmClient.stream({
            messages: [{ role: 'user', content: '读取图纸中的尺寸' }],
            // 80ms 看门狗：复现"Ollama 静默数分钟"场景（真实默认 180s）
            config: { apiType: 'ollama', model: 'qwen2.5:7b', options: { streamStallTimeoutMs: 80 } }
        })) {
            events.push(evt);
        }
        const err = events.find(e => e.type === 'error');
        assert.ok(err, 'stall 必须产出 error 事件，不能永远挂住');
        assert.strictEqual(err.error.code, 'STREAM_STALL');
        assert.ok(/模型响应中断/.test(err.error.message), '错误文案要可操作（提示检查模型服务）');
        assert.strictEqual(err.error.recoverable, true, 'stall 标记为可恢复，用户可重试');
    } finally {
        global.fetch = originalFetch;
    }
});

/** 最小 llmClient stub：吐一段正文后以指定 finishReason 收尾 */
function stubLLMFinish(delta, finishReason) {
    return {
        stream() {
            return (async function* () {
                yield { type: 'model', model: 'test' };
                if (delta) yield { type: 'token', delta };
                yield { type: 'done', finishReason };
            })();
        }
    };
}

for (const reason of ['runaway_truncated', 'length']) {
    test(`agent-loop: finishReason=${reason} 时向正文注入"已自动截断"可见提示`, async () => {
        const { App } = loadChatV2();
        App.ChatV2.llmClient = stubLLMFinish('图纸主要尺寸：10.45mm、9.27±0.15', reason);

        const turn = App.ChatV2.sessionStore.startTurn({
            message: '读取图纸中的尺寸', kind: 'text', context: {}
        });
        const events = [];
        for await (const e of App.ChatV2.agentLoop.run({
            turn,
            messages: [{ role: 'user', content: '读取图纸中的尺寸' }],
            tools: [],
            signal: new AbortController().signal
        })) {
            events.push(e);
        }

        const content = App.ChatV2.sessionStore.getCurrentTurn().content || '';
        assert.ok(/已自动截断/.test(content), '截断提示必须注入正文，不能让用户面对半句话/空白气泡');
        assert.ok(/10\.45mm/.test(content), '原有正文内容必须保留');
        assert.ok(!events.some(e => e.type === 'error'), '截断不是错误，不应再弹错误');
    });
}
