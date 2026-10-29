/**
 * ============================================================
 * chat-v2/services/compaction.js — 上下文压缩引擎
 * ============================================================
 * 职责：参考 DSH `compaction` 子系统，防止长对话上下文溢出：
 *   - compactIfNeeded()  — 检测压力（事件数/字符数超阈值）触发压缩
 *   - compactNow()       — 立即压缩
 *   - 压缩 = 把较早的对话折叠成摘要（带 provenance 的 compaction 事件），
 *     摘要经 sessionStore.appendCompactionEvent 写入事件流，
 *     deriveMessages 投影时摘要以 user 消息注入（_compaction 标记）。
 * 设计：
 *   - 可插拔：summarize 回调可替换（默认走 LLM；无 LLM 时降级为"截断+计数摘要"）
 *   - 非破坏性：原始事件保留（回放/审计仍完整），只影响模型可见投影
 *   - 阈值可配置：maxEvents / maxChars
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    const DEFAULTS = {
        maxEvents: 400,        // 事件数超过该值触发压缩（约 50 轮对话）
        maxChars: 60000,       // 累计字符超过该值触发压缩
        compressToEvents: 200, // 压缩后保留最近事件数
        minEventsToCompress: 60, // 事件太少不压缩（避免小会话误压缩）
        maxFoldPerPass: 800    // 单批最多折叠条数：积压巨大时分批渐进折叠，
                               // 避免一次性把上千条事件压成一份小摘要（严重失真）
    };

    /** 自定义摘要函数（可替换）。默认降级策略 */
    let _summarizer = null;

    /** 压缩摘要事件类型（与 session-store EVENT_TYPES.COMPACTION 一致） */
    const COMPACTION_EVENT_TYPE = 'compaction/summary';

    /**
     * 有效事件：尚未被压缩折叠（不在 _excludedSeqs）、且非摘要本身的事件。
     * 供压力检测与压缩选区共用：
     *   - 过滤已折叠事件 → 每个摘要只覆盖一段"新"原始事件，杜绝"摘要套娃"
     *   - 过滤摘要事件 → 永不把摘要折进摘要（避免信息逐级降质，直至 2000 截断后丢失）
     */
    function effectiveEvents(sessionStore, events) {
        const excluded = sessionStore?.getCompactedSeqs?.();
        return events.filter(e => {
            if (!e) return false;
            if (excluded && excluded.has(e.seq)) return false;
            if (e.type === COMPACTION_EVENT_TYPE) return false;
            return true;
        });
    }

    /**
     * 注册自定义摘要函数。
     * @param {Function} fn - async (events, turnId) => string 摘要文本
     */
    function setSummarizer(fn) {
        if (typeof fn === 'function') _summarizer = fn;
    }

    /**
     * 默认摘要：无 LLM 时降级为统计摘要。
     * 记录用户请求概览、工具调用概览、工具返回数据、最终结论。
     * 新增"工具结果"：旧实现只记工具名，一旦投影真正剔除早期事件，
     * 读取的表格/计算结果等实际数据会丢失——这是"高负载误删重要上下文"的最大风险点。
     */
    async function defaultSummarize(events, turnId) {
        const turnMsgs = events.filter(e => e.type === 'user/message');
        const toolCalls = events.filter(e => e.type === 'tool/call');
        const toolResults = events.filter(e => e.type === 'tool/result');
        const finalContent = events.filter(e => e.type === 'assistant/message')
            .map(e => e.data?.content || '').filter(Boolean).join(' ').slice(0, 500);

        const parts = [];
        if (turnMsgs.length > 0) {
            const reqs = turnMsgs.slice(-5).map(e => String(e.data?.content || '').slice(0, 60)).join('；');
            parts.push(`用户请求（近${Math.min(turnMsgs.length, 5)}条）: ${reqs}`);
        }
        // 工具返回的实际数据（压缩后模型不再看到原始 tool 消息，这里必须补偿）
        // 预算提升：按工具名去重保留最新结果（防同类调用挤掉不同数据），
        // 最多 5 个 × 160 字符。早期关键工具数据（读表/计算结果）折叠后仍能在摘要中找到。
        if (toolResults.length > 0) {
            const seen = new Map();
            for (const e of toolResults) {
                const name = e.data?.toolName || '工具';
                let content = e.data?.content;
                if (typeof content !== 'string') {
                    try { content = JSON.stringify(content ?? ''); } catch (err) { content = ''; }
                }
                const short = String(content || '').replace(/\s+/g, ' ').slice(0, 160);
                seen.set(name, `${name} → ${short || '(空)'}`);
            }
            const resParts = [...seen.values()].slice(-5);
            parts.push(`工具结果（近${resParts.length}个不同工具）: ${resParts.join('；')}`);
        }
        if (toolCalls.length > 0) {
            const tools = toolCalls.map(e => e.data?.toolName).filter(Boolean);
            const uniq = [...new Set(tools)].slice(0, 8);
            parts.push(`已执行工具: ${uniq.join(', ')}（共${tools.length}次调用）`);
        }
        if (finalContent) {
            parts.push(`最终结论: ${finalContent.slice(0, 200)}`);
        }
        const summary = `【上下文摘要】${parts.join('；')}`;
        return summary.length > 900 ? summary.slice(0, 900) + '…' : summary;
    }

    /**
     * 检查是否需要压缩。
     * @param {Object} sessionStore
     * @returns {{needed: boolean, reason: string, events: number, chars: number}}
     */
    function checkPressure(sessionStore) {
        if (!sessionStore || typeof sessionStore.getEvents !== 'function') {
            return { needed: false, reason: 'no-session-store', events: 0, chars: 0 };
        }
        // 压力只看"有效事件"（尚未折叠的原始事件，摘要不计入），
        // 避免把已排除事件算进阈值导致每轮都过度触发压缩（放大"摘要套娃"）。
        const effective = effectiveEvents(sessionStore, sessionStore.getEvents());
        if (effective.length < DEFAULTS.minEventsToCompress) {
            return { needed: false, reason: 'too-few-events', events: effective.length, chars: 0 };
        }
        // 字符统计（仅统计 content 相关事件）
        let chars = 0;
        for (const e of effective) {
            const d = e.data || {};
            const text = d.delta || d.content || d.summary || '';
            if (typeof text === 'string') chars += text.length;
        }
        if (effective.length > DEFAULTS.maxEvents) {
            return { needed: true, reason: 'events-over-limit', events: effective.length, chars };
        }
        if (chars > DEFAULTS.maxChars) {
            return { needed: true, reason: 'chars-over-limit', events: effective.length, chars };
        }
        return { needed: false, reason: 'ok', events: effective.length, chars };
    }

    /**
     * 执行压缩：折叠较早事件为摘要，保留最近事件。
     * @param {Object} sessionStore
     * @param {Object} [opts]
     * @param {string} [opts.turnId] - 指定 turn（默认全部）
     * @returns {Promise<{compacted: boolean, summary: string, removedEvents: number}>}
     */
    async function compactNow(sessionStore, opts = {}) {
        if (!sessionStore || typeof sessionStore.getEvents !== 'function') {
            return { compacted: false, summary: '', removedEvents: 0 };
        }
        const events = sessionStore.getEvents(opts.turnId || undefined);
        // 只对"有效事件"压缩：
        //   - 过滤已折叠（_excludedSeqs）的事件 → 每个摘要只覆盖一段新原始事件，杜绝"摘要套娃"
        //   - 过滤 compaction 摘要事件 → 永不把摘要折进摘要（避免信息逐级降质）
        const effective = effectiveEvents(sessionStore, events);
        if (effective.length < DEFAULTS.minEventsToCompress) {
            return { compacted: false, summary: '', removedEvents: 0 };
        }

        // 待压缩：除最近 compressToEvents 条外的早期【有效】事件。
        // 单批最多折叠 maxFoldPerPass 条：压缩积压巨大时避免一次性把
        // 上千条事件压成一份小摘要（严重失真），改为分批渐进折叠，保证每份摘要质量。
        const keepCount = Math.min(DEFAULTS.compressToEvents, effective.length - 1);
        const foldEnd = effective.length - keepCount;
        const foldStart = Math.max(0, foldEnd - DEFAULTS.maxFoldPerPass);
        const toCompress = effective.slice(foldStart, foldEnd);
        if (toCompress.length === 0) {
            return { compacted: false, summary: '', removedEvents: 0 };
        }

        // 生成摘要（可插拔）
        let summary;
        try {
            summary = _summarizer
                ? await _summarizer(toCompress, opts.turnId || null)
                : await defaultSummarize(toCompress, opts.turnId || null);
        } catch (e) {
            ChatLog.warn?.('COMPACT', '[compaction] 摘要生成失败，降级为统计摘要:', e?.message);
            summary = await defaultSummarize(toCompress, opts.turnId || null);
        }

        // 写入 compaction 事件（带 provenance），deriveMessages 会以 user 消息注入，
        // 并把 excludedSeqs 中被折叠的事件从模型投影中剔除（真正的上下文收缩）
        const turnIdForEvent = opts.turnId || toCompress[0]?.turnId || null;
        if (sessionStore.appendCompactionEvent) {
            sessionStore.appendCompactionEvent(turnIdForEvent, summary, {
                removedEvents: toCompress.length,
                keptEvents: keepCount,
                excludedSeqs: toCompress.map(e => e && Number.isFinite(e.seq) ? e.seq : null).filter(s => s != null)
            });
        }

        ChatLog.info?.('COMPACT', `[compaction] 已压缩 ${toCompress.length} 条早期事件 → 摘要(${summary.length}字)，保留 ${keepCount} 条`);
        return { compacted: true, summary, removedEvents: toCompress.length };
    }

    /**
     * 按需压缩（供 agent-loop / orchestrator 在合适时机调用）。
     * @param {Object} sessionStore
     * @returns {Promise<{needed: boolean, compacted: boolean, reason?: string, summary?: string}>}
     */
    async function compactIfNeeded(sessionStore) {
        const pressure = checkPressure(sessionStore);
        if (!pressure.needed) return { needed: false, compacted: false, reason: pressure.reason };
        const result = await compactNow(sessionStore);
        return { needed: true, compacted: result.compacted, reason: pressure.reason, summary: result.summary };
    }

    App.ChatV2.compaction = {
        compactIfNeeded,
        compactNow,
        checkPressure,
        setSummarizer,
        DEFAULTS
    };
    ChatLog.info?.('CHAT-V2', '✅ compaction loaded');
})();
