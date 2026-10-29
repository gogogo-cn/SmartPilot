/**
 * ============================================================
 * chat-v2/services/trajectory-replay.js — 轨迹回放引擎
 * ============================================================
 * 职责：基于事件溯源日志（sessionStore.getEvents）按时间顺序重放
 *       一个 turn 的完整执行轨迹，供 UI 回放/调试/审计。
 *   - replayTurn(turnId, { speed }) → AsyncGenerator<ReplayFrame>
 *   - ReplayFrame: { type, ts, offsetMs, data } 按事件序产出
 *   - 事件经 deriveMessages 投影可还原模型视角；原始事件保证回放保真
 * 设计（对齐 DSH：session/event 是唯一真源，回放派生自事件流）：
 *   - 非破坏性：只读事件日志，不修改任何状态
 *   - 支持调速（speed 倍速）与暂停信号（AbortSignal）
 *   - 帧类型覆盖：turn/start、user/message、assistant/chunk、tool/call、
 *     tool/result、assistant/message、turn/end、compaction
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /** 帧类型 → 展示语义（供 UI 渲染层消费） */
    const FRAME_META = {
        'turn/start':        { label: '轮次开始', icon: '🚀' },
        'user/message':      { label: '用户消息', icon: '👤' },
        'assistant/chunk':   { label: '模型输出', icon: '🤖' },
        'tool/call':         { label: '工具调用', icon: '🔧' },
        'tool/result':       { label: '工具结果', icon: '✅' },
        'assistant/message': { label: '最终回复', icon: '💬' },
        'step/start':        { label: '步骤开始', icon: '▶️' },
        'step/end':          { label: '步骤结束', icon: '⏹️' },
        'turn/end':          { label: '轮次结束', icon: '🏁' },
        'compaction/summary':{ label: '上下文压缩', icon: '📦' },
        'turn/canceled':     { label: '已取消', icon: '⛔' },
        'turn/failed':       { label: '执行失败', icon: '❌' }
    };

    /** 暂停控制：返回一个可 abort 的 signal（供回放循环检查） */
    function createPauseController() {
        const ctrl = new AbortController();
        return {
            signal: ctrl.signal,
            pause: () => ctrl.abort(),
            resume: () => {
                // AbortSignal 不可复用，暂停后重新创建（由调用方持有新 controller）
                return createPauseController();
            }
        };
    }

    /**
     * 按时间顺序重放指定 turn 的轨迹。
     * @param {Object} sessionStore
     * @param {string} turnId
     * @param {Object} [opts]
     * @param {number} [opts.speed=1] - 倍速（>1 快进，<1 慢放；0 = 不等，立即产出全部）
     * @param {AbortSignal} [opts.signal] - 中止信号
     * @returns {AsyncGenerator<Object>} ReplayFrame
     */
    async function* replayTurn(sessionStore, turnId, opts = {}) {
        if (!sessionStore || typeof sessionStore.getEvents !== 'function') {
            yield { type: 'error', error: 'sessionStore 不可用' };
            return;
        }
        const events = sessionStore.getEvents(turnId);
        if (!events || events.length === 0) {
            yield { type: 'error', error: '未找到该 turn 的事件日志: ' + turnId };
            return;
        }

        const speed = opts.speed === undefined ? 1 : Math.max(0, opts.speed);
        const signal = opts.signal || null;
        const startTs = events[0].ts;
        let prevTs = startTs;

        for (const evt of events) {
            if (signal?.aborted) {
                yield { type: 'aborted', turnId };
                return;
            }

            // 计算该帧相对偏移（原始事件时间差 × 1/speed）
            const rawOffset = evt.ts - prevTs;
            prevTs = evt.ts;

            const frame = {
                type: evt.type,
                turnId: evt.turnId,
                ts: evt.ts,
                offsetMs: Math.round(rawOffset / Math.max(speed, 0.001)),
                seq: evt.seq,
                data: evt.data || {},
                meta: FRAME_META[evt.type] || { label: evt.type, icon: '•' }
            };
            yield frame;

            // 调速（speed=0 时不等待）
            if (speed > 0 && rawOffset > 0) {
                const waitMs = Math.round(rawOffset / speed);
                if (waitMs > 0) {
                    await new Promise(resolve => {
                        const timer = setTimeout(resolve, Math.min(waitMs, 5000));
                        if (signal) {
                            signal.addEventListener('abort', () => {
                                clearTimeout(timer);
                                resolve();
                            }, { once: true });
                        }
                    });
                }
            }
        }

        yield { type: 'replay_done', turnId, frameCount: events.length, durationMs: prevTs - startTs };
    }

    /**
     * 快速回放：不等待，立即产出全部帧（用于审计/导出）。
     * @param {Object} sessionStore
     * @param {string} turnId
     * @returns {Array<Object>} ReplayFrame[]
     */
    function quickReplay(sessionStore, turnId) {
        const frames = [];
        const events = sessionStore?.getEvents?.(turnId) || [];
        const startTs = events[0]?.ts || Date.now();
        for (const evt of events) {
            frames.push({
                type: evt.type,
                turnId: evt.turnId,
                ts: evt.ts,
                offsetMs: evt.ts - startTs,
                seq: evt.seq,
                data: evt.data || {},
                meta: FRAME_META[evt.type] || { label: evt.type, icon: '•' }
            });
        }
        return frames;
    }

    /**
     * 导出轨迹为可读文本（审计/调试）。
     * @param {Object} sessionStore
     * @param {string} turnId
     * @returns {string}
     */
    function exportTrajectoryText(sessionStore, turnId) {
        const frames = quickReplay(sessionStore, turnId);
        const lines = frames.map(f => {
            const d = f.data || {};
            let detail = '';
            if (f.type === 'user/message') detail = String(d.content || '').slice(0, 80);
            else if (f.type === 'assistant/chunk') detail = String(d.delta || '').slice(0, 80);
            else if (f.type === 'tool/call') detail = `${d.toolName || ''} ${JSON.stringify(d.args || {}).slice(0, 60)}`;
            else if (f.type === 'tool/result') detail = `${d.toolName || ''} status=${d.status || '?'} ${String(d.content || d.error || '').slice(0, 60)}`;
            else if (f.type === 'compaction/summary') detail = String(d.summary || '').slice(0, 80);
            else if (f.type === 'turn/start' || f.type === 'assistant/message' || f.type === 'turn/end') detail = JSON.stringify(d).slice(0, 80);
            return `[${(f.offsetMs).toString().padStart(6)}ms] ${f.meta.icon} ${f.meta.label}${detail ? ' | ' + detail : ''}`;
        });
        return `=== 轨迹回放: ${turnId} (${frames.length} 帧) ===\n` + lines.join('\n');
    }

    App.ChatV2.trajectoryReplay = {
        replayTurn,
        quickReplay,
        exportTrajectoryText,
        createPauseController,
        FRAME_META
    };
    ChatLog.info?.('CHAT-V2', '✅ trajectoryReplay loaded');
})();
