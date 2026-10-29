/**
 * ============================================================
 * Tests/trajectory-replay.test.js — 轨迹回放测试
 * ============================================================
 * 覆盖：事件序列回放、快速回放、导出文本、调速
 * 运行：node Tests/trajectory-replay.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

test('trajectory-replay: 完整 turn 回放帧序列', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: '你好', kind: 'text' });
    ss.appendDelta(turn.id, 'content', '你');
    ss.appendDelta(turn.id, 'content', '好');
    const tc = ss.addToolCall(turn.id, { name: 'excel_read_range', args: { range: 'A1' } });
    ss.updateToolCall(turn.id, tc.id, { status: 'completed', result: { rows: 3 } });
    ss.completeTurn(turn.id, { content: '你好！' });

    const replay = App.ChatV2.trajectoryReplay;
    const frames = replay.quickReplay(ss, turn.id);
    const types = frames.map(f => f.type);

    // 快速回放应包含核心事件
    assert.ok(types.includes('turn/start'));
    assert.ok(types.includes('user/message'));
    assert.ok(types.includes('assistant/chunk'));
    assert.ok(types.includes('tool/call'));
    assert.ok(types.includes('tool/result'));
    assert.ok(types.includes('turn/end'));
    // 帧应有 meta 标签
    assert.ok(frames[0].meta && frames[0].meta.label);
    console.log('帧序列:', types.join(' → '));
});

test('trajectory-replay: replayTurn 异步逐帧产出', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({ message: 'hi', kind: 'text' });
    ss.appendDelta(turn.id, 'content', 'a');
    ss.completeTurn(turn.id, { content: 'a' });

    const replay = App.ChatV2.trajectoryReplay;
    const frames = [];
    for await (const f of replay.replayTurn(ss, turn.id, { speed: 0 })) {
        frames.push(f);
    }
    // speed=0 立即产出全部
    assert.ok(frames.length >= 4, '应产出所有帧，实际 ' + frames.length);
    assert.strictEqual(frames[frames.length - 1].type, 'replay_done');
});

test('trajectory-replay: 导出文本包含轨迹信息', async () => {
    const { App } = loadChatV2();
    const ss = App.ChatV2.sessionStore;
    const turn = ss.startTurn({ message: '测试', kind: 'text' });
    ss.completeTurn(turn.id, { content: '完成' });

    const text = App.ChatV2.trajectoryReplay.exportTrajectoryText(ss, turn.id);
    assert.ok(text.includes('轨迹回放'));
    assert.ok(text.includes('用户消息'));
    assert.ok(text.includes('最终回复'));
    console.log(text.split('\n')[0]);
});

test('trajectory-replay: 不存在 turn 返回错误帧', async () => {
    const { App } = loadChatV2();
    const replay = App.ChatV2.trajectoryReplay;
    const gen = replay.replayTurn(App.ChatV2.sessionStore, 'no_such_turn', { speed: 0 });
    const first = await gen.next();
    assert.strictEqual(first.value.type, 'error');
});

// 兜底退出：避免 session-store 防抖定时器挂起进程
if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
