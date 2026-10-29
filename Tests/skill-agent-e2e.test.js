/**
 * ============================================================
 * Tests/skill-agent-e2e.test.js — Skill Agent 模式端到端集成测试
 * ============================================================
 * 验证执行模式声明化：
 *   - Skill 声明 executionMode=agent + agentConfig 时，skill-pipeline
 *     分发到自主 Agent 循环，并把 maxIterations / tokenGuard 透传给 agentLoop.run
 *   - agent-loop 每轮落步骤日志（thought/tool_call/tool_return/finish）到
 *     sessionStore（独立于上下文压缩）
 * 运行：node Tests/skill-agent-e2e.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, DEFAULT_ORDER } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，但截断到 skill-pipeline（不加载 chat-entry/index，避免入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('pipelines/code-pipeline.js') + 1);

test('skill-agent e2e: executionMode=agent → 透传 agentLoop.run(maxIterations, tokenGuard) + 落步骤日志', async () => {
    const App = loadChatV2(ORDER).App;

    // 1. 注册 agent 模式 Skill（模拟某 Skill 声明 executionMode=agent）
    App.cachedSkills = [{
        name: 'SmartReportAgent', displayName: '报表智能代理',
        keywords: ['报表', '汇总', '报告'],
        inputContext: ['selection'], pipeline: 'skill',
        executionMode: 'agent',
        agentConfig: { maxIterations: 3, maxContextTokens: 4096, tokenWarnRatio: 0.7, tokenCriticalRatio: 0.9 }
    }];
    const tr = App.ChatV2.toolRegistry;

    // 2. 用桩替换 agentLoop.run：捕获透传参数，并模拟 agent 过程（写步骤日志 + 产出内容）
    let runParams = null;
    App.ChatV2.agentLoop = {
        MAX_TOOL_ROUNDS: 8,
        async *run(params) {
            runParams = params;
            const sessionStore = App.ChatV2.sessionStore;
            yield { type: 'status', text: 'agent 模拟运行中' };
            // 模拟 agent 内部：思考 → 调用工具 → 返回 → 完成
            sessionStore.appendStepLog(params.turn.id, 'thought', '分析用户报表需求', 1);
            sessionStore.appendStepLog(params.turn.id, 'tool_call', '调用工具 excel_read_range', 1);
            sessionStore.appendStepLog(params.turn.id, 'tool_return', '工具 excel_read_range 成功', 1);
            sessionStore.appendStepLog(params.turn.id, 'finish', 'Agent 循环结束：completed', 2);
            sessionStore.updateTurn(params.turn.id, { content: '报表已生成，共 3 个区域汇总完成。' });
        }
    };

    // 3. 模拟 turn
    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '帮我把这个月的报表汇总一下',
        kind: 'text',
        context: { selection: { address: 'A1:F30', sheetName: 'Sheet1' }, _matchedSkill: App.cachedSkills[0] }
    });

    // 4. 运行 skill-pipeline
    const pipeline = App.ChatV2.pipelines.skill;
    const gen = pipeline.execute(turn, new AbortController().signal);
    const events = [];
    for await (const evt of gen) {
        events.push(evt.type);
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    // 5. 断言：分发到 agentLoop.run，且透传 agentConfig 参数
    assert.ok(runParams, 'agent 模式应调用 agentLoop.run');
    assert.strictEqual(runParams.maxIterations, 3, 'maxIterations 应透传 agentConfig.maxIterations');
    assert.strictEqual(runParams.persistTools, true, '自主 Agent 应每轮保留工具，才能多步调用');
    assert.strictEqual(runParams.tokenGuard.maxTokens, 4096, 'maxTokens 应透传 agentConfig.maxContextTokens');
    assert.strictEqual(runParams.tokenGuard.warnRatio, 0.7, 'warnRatio 应透传');
    assert.strictEqual(runParams.tokenGuard.criticalRatio, 0.9, 'criticalRatio 应透传');

    // 6. 断言：messages 含 system + user
    const msgs = runParams.messages;
    assert.strictEqual(msgs.length, 2, '初始 messages 应含 system + user');
    assert.strictEqual(msgs[0].role, 'system');
    assert.ok(String(msgs[0].content).includes('报表智能代理'), 'system prompt 应引用 Skill 显示名');
    assert.strictEqual(msgs[1].role, 'user');

    // 7. 断言：步骤日志已持久化到 sessionStore（独立于上下文压缩）
    const logs = sessionStore.queryStepLogs({ taskId: turn.id });
    const types = logs.map(l => l.stepType);
    assert.ok(types.includes('thought'), '应含 thought 步骤日志');
    assert.ok(types.includes('tool_call'), '应含 tool_call 步骤日志');
    assert.ok(types.includes('tool_return'), '应含 tool_return 步骤日志');
    assert.ok(types.includes('finish'), '应含 finish 步骤日志');
    console.log('  步骤日志类型:', types.join(', '));

    // 8. 断言：turn 正常完成
    // 注意：completeTurn 会把 STATE.currentTurn 置空，因此用测试持有的 turn 引用验证。
    assert.ok(events.includes('pipeline_done'), 'agent 模式应完成 pipeline');
    assert.strictEqual(turn.status, 'completed', 'turn 应标记为 completed');
    assert.ok(String(turn.content).includes('报表已生成'), 'turn 应保存 agent 产出内容');
});

test('agent-loop token 水位: 真实循环触发 token_watermark 事件 + 落 warning 步骤日志', async () => {
    const { App } = loadChatV2();  // 完整加载（真实 agent-loop）
    const sessionStore = App.ChatV2.sessionStore;
    const toolRegistry = App.ChatV2.toolRegistry;

    // 1. 注册一个 MCP 测试工具（agent 循环第一轮会调用它）
    toolRegistry.register({
        name: 'excel_test_read',
        source: 'mcp',
        displayName: '测试读取',
        description: '测试用读取工具',
        execute: async () => ({ ok: true, result: '[[1,2],[3,4]]' })
    });

    // 2. llmClient stub：第 1 轮调用工具，第 2 轮总结文本
    let llmCalls = 0;
    App.ChatV2.llmClient = {
        async *stream() {
            llmCalls++;
            yield { type: 'model', model: 'test' };
            if (llmCalls === 1) {
                yield {
                    type: 'tool_calls',
                    toolCalls: [{ id: 'call_1', function: { name: 'excel_test_read', arguments: '{}' } }]
                };
            } else {
                yield { type: 'token', delta: '数据已读取完成，共 2 行 2 列。' };
            }
            yield { type: 'done', finishReason: 'stop' };
        }
    };

    // 3. 构造 turn + 运行真实 agent-loop
    const turn = sessionStore.startTurn({
        message: '读取数据',
        kind: 'text',
        context: { selection: { address: 'A1:B2', sheetName: 'Sheet1' } }
    });

    const gen = App.ChatV2.agentLoop.run({
        turn,
        messages: [
            { role: 'system', content: '你是测试助手，请自主调用工具完成任务。' },
            { role: 'user', content: '读取数据' }
        ],
        tools: toolRegistry.listBySource('mcp'),
        signal: new AbortController().signal,
        maxIterations: 3,
        // 【P1 水位护栏】故意设小阈值：maxTokens=200 + warnRatio=0.1 → 告警水位 ≈ 20 token
        tokenGuard: { maxTokens: 200, warnRatio: 0.1, criticalRatio: 0.9 },
        keepLatestRound: 1
    });

    const eventTypes = [];
    let watermark = null;
    for await (const evt of gen) {
        eventTypes.push(evt.type);
        if (evt.type === 'token_watermark') watermark = evt;
    }

    // 4. 断言：水位事件触发（第 2 轮 round>0 时检查）
    assert.ok(watermark, '第 2 轮应触发 token_watermark 事件');
    assert.strictEqual(watermark.level, 'warn', '应处于告警水位（非危险）');
    assert.ok(watermark.total > 0, '应估算出 token 用量');
    assert.strictEqual(watermark.windowMax, 200, '窗口上限应取 tokenGuard.maxTokens');
    console.log(`  [watermark] level=${watermark.level} total=${watermark.total} windowMax=${watermark.windowMax} ratio=${watermark.ratio.toFixed(3)}`);

    // 5. 断言：步骤日志持久化（含 warning）
    const logs = sessionStore.queryStepLogs({ taskId: turn.id });
    const types = logs.map(l => l.stepType);
    assert.ok(types.includes('thought'), '应含 thought 步骤日志');
    assert.ok(types.includes('tool_call'), '应含 tool_call 步骤日志');
    assert.ok(types.includes('tool_return'), '应含 tool_return 步骤日志');
    assert.ok(types.includes('finish'), '应含 finish 步骤日志');
    assert.ok(types.includes('warning'), 'token 水位应落 warning 步骤日志');
    const warnLog = logs.find(l => l.stepType === 'warning');
    assert.ok(String(warnLog.content).includes('token 水位告警'), 'warning 日志应含水位描述');
    // 压缩结果按实际触发情况记录：事件 < 60 条时 compactNow 正确返回"压缩未生效"
    const hasCompactionResult = String(warnLog.content).includes('已触发上下文压缩')
        || String(warnLog.content).includes('压缩未生效');
    assert.ok(hasCompactionResult, 'warning 日志应记录压缩触发结果（成功或未生效）');
    assert.strictEqual(typeof watermark.compacted, 'boolean', 'token_watermark 事件应携带 compacted 字段');
    console.log('  步骤日志类型:', types.join(', '));
    console.log('  warning 日志:', warnLog.content);

    // 6. 断言：循环正常结束
    assert.ok(eventTypes.includes('loop_done'), '循环应以 loop_done 结束');
    assert.ok(eventTypes.includes('round_start'), '应经过多轮 round_start');
});
