/**
 * ============================================================
 * Tests/session-store.test.js — 事件溯源测试
 * ============================================================
 * 覆盖：事件序列、deriveMessages 投影、replayTurn 重放、compaction
 * 运行：node --test Tests/
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

test('session-store: 完整 turn 的事件序列', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: '美化表格', kind: 'text' });
    ss.appendDelta(turn.id, 'content', '执行中');
    const tc = ss.addToolCall(turn.id, { name: 'excel_read_range', args: { range: 'A1' } });
    ss.updateToolCall(turn.id, tc.id, { status: 'completed', result: { rows: 5 } });
    ss.completeTurn(turn.id, { content: '完成' });

    const types = ss.getEvents(turn.id).map(e => e.type);
    assert.deepStrictEqual(types, [
        'turn/start', 'user/message', 'assistant/chunk', 'tool/call', 'tool/result',
        'assistant/message', 'turn/end'
    ]);
});

test('session-store: deriveMessages 投影 chunk 累积', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: 'hi', kind: 'text' });
    ss.appendDelta(turn.id, 'content', '你');
    ss.appendDelta(turn.id, 'content', '好');
    ss.completeTurn(turn.id, { content: '你好' });

    const msgs = ss.deriveMessages(turn.id);
    const assistant = msgs.filter(m => m.role === 'assistant');
    // assistant/message 事件带最终 content '你好'（chunk 累积 + completeTurn 的 content）
    assert.ok(assistant.some(m => (m.content || '').includes('你好')));
});

test('session-store: replayTurn 断点重放', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: '清洗数据', kind: 'text', context: { selection: { range: 'A1:B2' } } });
    ss.appendDelta(turn.id, 'content', '去重中');
    const tc = ss.addToolCall(turn.id, { name: 'skill_SmartDataCleaner', args: {} });
    ss.updateToolCall(turn.id, tc.id, { status: 'completed', result: { ok: true } });
    ss.completeTurn(turn.id, { content: '已清洗' });

    const replayed = ss.replayTurn(turn.id);
    assert.strictEqual(replayed.userMessage, '清洗数据');
    assert.strictEqual(replayed.status, 'completed');
    assert.strictEqual(replayed.toolCalls.length, 1);
    assert.ok(replayed.content.includes('去重中'));
});

test('session-store: compaction 摘要注入投影', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: 'a', kind: 'text' });
    ss.completeTurn(turn.id, { content: 'b' });
    ss.appendCompactionEvent(turn.id, '【摘要】前文已处理');

    const msgs = ss.deriveMessages();
    const comp = msgs.find(m => m._compaction);
    assert.ok(comp, 'compaction 摘要应注入投影');
    assert.strictEqual(comp.content, '【摘要】前文已处理');
});

test('session-store: clearHistory 清空事件日志', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;
    ss.startTurn({ message: 'x', kind: 'text' });
    assert.ok(ss.getEvents().length > 0);
    ss.clearHistory();
    assert.strictEqual(ss.getEvents().length, 0);
});

test('compaction 防误压缩: 不重复折叠已排除事件、摘要事件永不被折叠、有效事件收敛', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;
    const comp = App.ChatV2.compaction;
    assert.ok(comp && typeof comp.compactNow === 'function', 'compaction 模块应已加载');

    // 造一轮完整 turn（含用户请求 + 工具读表数据 + 结论），模拟高负载关键数据
    function addTurn(i) {
        const t = ss.startTurn({ message: `第${i}次请求：读取并处理销售数据`, kind: 'text' });
        ss.appendDelta(t.id, 'content', `正在处理第${i}轮`);
        const tc = ss.addToolCall(t.id, { name: 'excel_read_range', args: { range: 'A1:C50' } });
        ss.updateToolCall(t.id, tc.id, { status: 'completed', result: { rows: `第${i}轮读取到 50 行数据，含金额列，合计 10000 元` } });
        ss.completeTurn(t.id, { content: `第${i}轮处理完成，金额合计 10000 元` });
    }

    // 35 轮 → 245 有效事件（超过 compressToEvents=200，触发真实折叠）
    for (let i = 0; i < 35; i++) addTurn(i);

    const r1 = await comp.compactNow(ss);
    assert.strictEqual(r1.compacted, true, '首次压缩应生效');
    assert.ok(r1.removedEvents > 0, `首次压缩应折叠事件，实际 ${r1.removedEvents}`);

    const firstExcluded = new Set(ss.getCompactedSeqs());
    assert.strictEqual(firstExcluded.size, r1.removedEvents, '被排除集合应等于本次折叠事件数');

    // 摘要事件绝不应被折叠/排除（防"摘要套娃"）
    const summaryEvents = ss.getEvents().filter(e => e.type === 'compaction/summary');
    assert.ok(summaryEvents.length === 1, '首次压缩应产生 1 条摘要事件');
    for (const s of summaryEvents) {
        assert.ok(!firstExcluded.has(s.seq), '摘要事件不应被排除');
    }

    // 再补 35 轮 → 有效事件回升到 ~445，触发第二次压缩
    for (let i = 35; i < 70; i++) addTurn(i);
    const r2 = await comp.compactNow(ss);
    if (r2.compacted) {
        const secondExcluded = new Set(ss.getCompactedSeqs());
        // 关键保护：已排除 seq 只增不减（每份摘要只覆盖一段新事件，而非重排后误删旧块）
        for (const seq of firstExcluded) {
            assert.ok(secondExcluded.has(seq), `已排除 seq=${seq} 不应丢失`);
        }
        for (const s of ss.getEvents().filter(e => e.type === 'compaction/summary')) {
            assert.ok(!secondExcluded.has(s.seq), '多次压缩后摘要事件仍不应被折叠');
        }
    }

    // 压缩后"有效事件"（非排除、非摘要）应收敛到保留区附近，防止重复压缩套娃
    const effective = ss.getEvents().filter(e =>
        !ss.getCompactedSeqs().has(e.seq) && e.type !== 'compaction/summary');
    assert.ok(effective.length <= comp.DEFAULTS.compressToEvents + 1,
        `压缩后有效事件应收敛到保留区附近，实际 ${effective.length}`);

    // 压力检测不再把已排除事件计入 → 不会每轮都误报超限
    const pressure = comp.checkPressure(ss);
    assert.strictEqual(pressure.needed, false, '有效事件收敛后不应再触发压缩');
});

test('session-store: 编辑重发退回到该消息发送前，工具结果和压缩摘要都不再可见', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const earlier = ss.startTurn({ message: '先看销售一部', kind: 'text' });
    ss.completeTurn(earlier.id, { content: '销售一部已查看' });
    const before = ss.getHistoryForLLM().map(m => m.content);

    const edited = ss.startTurn({ message: '将选中工作表进行数据合并，并设置成中国红', kind: 'text' });
    ss.addToolCall(edited.id, { name: 'excel_get_sheets', args: {} });
    ss.completeTurn(edited.id, {
        content: '[{"name":"销售一部","usedRange":"A1:I21"},{"name":"合并结果","usedRange":"A1:I40"}]'
    });
    const earlierSeqs = ss.getEvents(earlier.id).map(e => e.seq);
    ss.appendCompactionEvent(edited.id, '已读取工作表清单并合并成合并结果', { excludedSeqs: earlierSeqs });
    ss.appendStepLog(edited.id, 'tool_return', '{"name":"销售一部","usedRange":"A1:I21"}', 1);

    assert.strictEqual(ss.removeTurnsSince(edited.id), true);

    const after = ss.getHistoryForLLM().map(m => m.content);
    const blob = after.join('\n');
    assert.deepStrictEqual(after, before);
    assert.ok(!blob.includes('usedRange'));
    assert.ok(!blob.includes('合并结果'));
    assert.ok(!blob.includes('工作表清单'));
    assert.ok(ss.getHistory().every(h => h.turnId !== edited.id));
    assert.ok(!ss.getHistory().some(h => String(h.content || '').includes('usedRange')));
    assert.strictEqual(ss.getTurns().some(t => t.id === edited.id), false);
    assert.ok(ss.queryStepLogs({ taskId: edited.id }).length === 0);
});

// 测试结束后强制退出（session-store 的 saveToLegacyHistory 定时器会挂起进程）
// 仅在直接运行（非 --test runner）时生效
if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
