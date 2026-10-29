/**
 * ============================================================
 * chat-v2/services/session-store.js — 单一状态源
 * ============================================================
 * 职责：维护 Turn lifecycle + 历史记录 + 当前会话状态
 * 设计：唯一状态 owner，其他模块通过 subscribe/get 访问，不直接修改
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // ═══════════════════════════════════════════════════════════════
    // ExecSpec Schema 声明式（v20260924 新增）
    // 统一 setExecSpec 写入契约，消除 5 个 pipeline 各自组装 data 导致的字段漂移。
    //
    // 设计原则：
    //   1. 公共字段（selection / referenceRegions）走 buildCommonData 工厂，
    //      保证每个 pipeline 写入的 selection 都有 address/sheet/headers/rows 全套字段。
    //   2. pipeline 专有字段（如 PPT 的 slides、DOC 的 attachments、Python 的 analysis）
    //      统一挂在 data.ext[pipeline] 命名空间下，不污染公共 data 字段。
    //   3. setExecSpec 内部做 schema 校验（开发模式 warn，生产静默），
    //      写入方忘了带公共字段时显式告警。
    // ═══════════════════════════════════════════════════════════════

    /** 公共 selection 字段 schema —— 所有 Excel 类 pipeline 必须按此结构写入 */
    const SELECTION_SCHEMA = [
        'address', 'range', 'workbook', 'workbookName',
        'sheet', 'worksheet', 'sheetName',
        'headers', 'rows', 'rowCount', 'colCount'
    ];

    /** 公共 ExecSpec 顶层字段 —— 所有 pipeline 必须包含 */
    const EXEC_SPEC_REQUIRED_TOP = ['pipeline'];
    /** 公共 ExecSpec 可选顶层字段 */
    const EXEC_SPEC_OPTIONAL_TOP = ['data', 'params', 'result', 'taskStructure', 'ts'];

    /** 公共 data 字段 —— 大部分 pipeline 应写入这些 */
    const EXEC_SPEC_DATA_COMMON = ['selection', 'referenceRegions', 'ext'];

    /**
     * 工厂：从 turn.context 提取公共 selection 快照（统一 schema）
     * 所有 Excel 类 pipeline（text/skill/python/code）应通过此函数构建 data.selection，
     * 不要再各写一套 selection 组装逻辑，避免字段漂移。
     * @param {Object} turn - turn 对象
     * @returns {Object|null} 规范化后的 selection 或 null（turn.context 无 selection）
     */
    function buildCommonSelection(turn) {
        const ctxSel = turn?.context?.selection;
        if (!ctxSel) return null;
        const headers = Array.isArray(ctxSel.headers) ? ctxSel.headers : [];
        const rows = Array.isArray(ctxSel.rows) ? ctxSel.rows.slice(0, 200) : [];
        return {
            address: ctxSel.address || ctxSel.range || '',
            range: ctxSel.range || ctxSel.address || '',
            workbook: ctxSel.workbook || ctxSel.workbookName || '',
            workbookName: ctxSel.workbookName || ctxSel.workbook || '',
            sheet: ctxSel.sheet || ctxSel.sheetName || ctxSel.worksheet || '',
            worksheet: ctxSel.worksheet || ctxSel.sheet || ctxSel.sheetName || '',
            sheetName: ctxSel.sheetName || ctxSel.sheet || ctxSel.worksheet || '',
            headers: headers,
            rows: rows,
            rowCount: ctxSel.rowCount || rows.length || 0,
            colCount: ctxSel.colCount || headers.length || 0
        };
    }

    /**
     * 工厂：从 turn.context 提取公共 referenceRegions 快照（统一 schema）
     * @param {Object} turn - turn 对象
     * @returns {Array|null} 规范化后的 referenceRegions 或 null
     */
    function buildCommonReferenceRegions(turn) {
        const refs = turn?.context?.referenceRegions;
        if (!Array.isArray(refs) || refs.length === 0) return null;
        return refs.map(function (r) {
            const headers = Array.isArray(r.headers) ? r.headers : [];
            const rows = Array.isArray(r.rows) ? r.rows.slice(0, 200) : [];
            return {
                address: r.address || r.range || '',
                range: r.range || r.address || '',
                workbook: r.workbook || '',
                sheet: r.sheet || r.sheetName || r.worksheet || '',
                worksheet: r.worksheet || r.sheet || r.sheetName || '',
                sheetName: r.sheetName || r.sheet || r.worksheet || '',
                headers: headers,
                rows: rows,
                rowCount: r.rowCount || rows.length || 0,
                colCount: r.colCount || headers.length || 0
            };
        });
    }

    /**
     * 工厂：构建公共 data 对象 —— 所有 Excel 类 pipeline 的 setExecSpec data 都应从此起步
     * @param {Object} turn
     * @param {Object} [ext] - pipeline 专有扩展数据，挂在 data.ext[pipeline] 下
     * @returns {Object} 规范化后的 data
     */
    function buildCommonData(turn, ext) {
        const data = {};
        const sel = buildCommonSelection(turn);
        const refs = buildCommonReferenceRegions(turn);
        if (sel) data.selection = sel;
        if (refs) data.referenceRegions = refs;
        if (ext && typeof ext === 'object') {
            data.ext = ext;
        }
        return data;
    }

    /**
     * 校验 ExecSpec 顶层 + 公共 data 字段（开发模式 warn，生产静默）
     * @param {Object} spec
     * @returns {boolean} true=通过校验
     */
    function validateExecSpec(spec) {
        if (!spec || typeof spec !== 'object') {
            ChatLog.warn?.('SESSION', '[SessionStore.setExecSpec] spec 不是有效对象');
            return false;
        }
        // 必填顶层
        for (const f of EXEC_SPEC_REQUIRED_TOP) {
            if (spec[f] === undefined || spec[f] === null) {
                ChatLog.warn?.('SESSION', `[SessionStore.setExecSpec] 缺少必填顶层字段: ${f}`);
            }
        }
        // data.selection schema 校验（pipeline 有 data 时）
        if (spec.data && spec.pipeline !== 'doc') {
            const sel = spec.data.selection;
            if (sel && typeof sel === 'object') {
                if (!sel.address && !sel.range) {
                    ChatLog.warn?.('SESSION', `[SessionStore.setExecSpec] pipeline=${spec.pipeline} data.selection 缺少 address/range`);
                }
                if (!sel.sheet && !sel.sheetName && !sel.worksheet) {
                    ChatLog.warn?.('SESSION', `[SessionStore.setExecSpec] pipeline=${spec.pipeline} data.selection 缺少 sheet/sheetName/worksheet`);
                }
            }
        }
        return true;
    }

    // ═══════════════════════════════════════════════════════════════
    // append-only SessionEvent 日志（参考 DSH core/session）
    // 原则：模型可见 ⟺ 已记录。所有持久事实都追加到事件流，
    // 状态（turns/history/currentTurn）只是事件流的投影，可从事件重放重建。
    // 现有 STATE 结构保留作"投影缓存"，新增 _events 为唯一持久事实源。
    // ═══════════════════════════════════════════════════════════════
    /** @type {Array<Object>} append-only 事件日志：{seq, type, turnId, ts, data} */
    const _events = [];
    /** 编辑重发已撤回的 turnId。投影历史时跳过，效果等同这条消息没发过。 */
    const _droppedTurnIds = new Set();
    /** 撤回点之后才写上的压缩摘要。只藏这一条，不连更早的轮次一起藏。 */
    const _rewoundSeqs = new Set();
    let _eventSeq = 0;

    // ═══════════════════════════════════════════════════════════════
    // 已压缩折叠事件的 seq 集合（投影层排除）
    // 背景：旧实现 compactNow 只追加摘要、从不剔除早期事件，导致"压缩空转"——
    //   deriveMessages 注入摘要但原事件仍全部投影，模型可见 token 永不减少。
    // 修复：compactNow 把被折叠事件的 seq 记入该集合，deriveMessages 投影时跳过，
    //   实现真正的上下文收缩。原始事件仍保留在 _events（回放/审计完整），仅模型投影排除。
    // 原则：只跳转"被最新压缩折叠的旧事件"，摘要事件本身与保留区不受影响。
    // ═══════════════════════════════════════════════════════════════
    const _excludedSeqs = new Set();
    /** 事件类型映射：新增持久状态时扩展此表 */
    const EVENT_TYPES = {
        TURN_START: 'turn/start',
        TURN_END: 'turn/end',
        USER_MESSAGE: 'user/message',
        ASSISTANT_CHUNK: 'assistant/chunk',
        ASSISTANT_MESSAGE: 'assistant/message',
        TOOL_CALL: 'tool/call',
        TOOL_RESULT: 'tool/result',
        STEP_START: 'step/start',
        STEP_END: 'step/end',
        COMPACTION: 'compaction/summary',
        TURN_CANCELED: 'turn/canceled',
        TURN_FAILED: 'turn/failed'
    };

    // 诊断：appendDelta 因“尾部重复”被跳过的次数（定位表格/思考被多次渲染）
    let _deltaSkip = 0;

    /** 追加一条持久事件（内部） */
    function appendEvent(type, turnId, data) {
        _eventSeq++;
        const evt = { seq: _eventSeq, type, turnId: turnId || null, ts: Date.now(), data: data || {} };
        _events.push(evt);
        // 上限保护：保留最近 2000 条事件（防内存无限增长）。
        // 物理截断会移除最早原始事件（回放/审计不再完整，仅剩摘要），
        // 同步从 _excludedSeqs 剪枝，避免已删除 seq 残留导致集合无限膨胀。
        if (_events.length > 2000) {
            const removed = _events.splice(0, _events.length - 2000);
            for (const r of removed) {
                _excludedSeqs.delete(r.seq);
                _rewoundSeqs.delete(r.seq);
            }
            ChatLog.warn?.('SESSION', `[SessionStore] 事件日志超上限，物理截断 ${removed.length} 条最早事件（该段原始记录仅剩摘要，回放/审计不再完整）`);
        }
        return evt;
    }

    // ═══════════════════════════════════════════════════════════════
    // 独立于上下文压缩的调试日志（对齐 multi-step-agent-mcp）
    // 原则：与 LLM 上下文完全隔离——compaction 只折叠 _events 投影，绝不销毁 _stepLogs。
    // 每轮 thought/tool_call/tool_return/warning/error/finish 都记录；可按 taskId 查询。
    // 持久化：内存 + localStorage 镜像（防抖），跨会话可查；存储不可用（如测试环境）自动降级为纯内存。
    // ═══════════════════════════════════════════════════════════════
    /** 步骤类型（对齐 multi-step-agent-mcp step_record_append 的 step_type 枚举） */
    const STEP_TYPES = ['thought', 'tool_call', 'tool_return', 'warning', 'error', 'finish'];
    /** @type {Array<Object>} 步骤日志：{seq, taskId, stepType, content, iteration, ts} */
    const _stepLogs = [];
    let _stepLogSeq = 0;
    /** 每任务持久化防抖定时器 */
    const _stepLogSaveTimers = new Map();
    /** 单任务持久化上限（防 localStorage 无限膨胀） */
    const STEP_LOG_PERSIST_MAX = 500;

    function stepLogStorageKey(taskId) {
        return 'smartpilot.stepLogs.' + String(taskId);
    }

    function loadStepLogsFromStorage(taskId) {
        try {
            const raw = localStorage.getItem(stepLogStorageKey(taskId));
            if (!raw) return null;
            const parsed = JSON.parse(raw);
            return Array.isArray(parsed) ? parsed : null;
        } catch (e) {
            return null;
        }
    }

    function scheduleStepLogSave(taskId) {
        if (_stepLogSaveTimers.has(taskId)) clearTimeout(_stepLogSaveTimers.get(taskId));
        const timer = setTimeout(() => {
            _stepLogSaveTimers.delete(taskId);
            try {
                const logs = _stepLogs.filter(l => l.taskId === taskId);
                localStorage.setItem(stepLogStorageKey(taskId), JSON.stringify(logs.slice(-STEP_LOG_PERSIST_MAX)));
            } catch (e) {
                // localStorage 不可用（测试/隐私模式）→ 纯内存降级
            }
        }, 500);
        _stepLogSaveTimers.set(taskId, timer);
    }

    function appendStepLog(taskId, stepType, content, iteration) {
        if (!taskId) return null;
        const type = STEP_TYPES.indexOf(stepType) === -1 ? 'thought' : stepType; // 未知类型兜底
        _stepLogSeq++;
        const entry = {
            seq: _stepLogSeq,
            taskId,
            stepType: type,
            content: content || null,
            iteration: iteration == null ? null : Number(iteration),
            ts: Date.now()
        };
        _stepLogs.push(entry);
        // 上限保护：保留最近 2000 条（防内存无限增长）
        if (_stepLogs.length > 2000) _stepLogs.splice(0, _stepLogs.length - 2000);
        scheduleStepLogSave(taskId);
        return entry;
    }

    /**
     * 【事件溯源】从事件日志投影出模型可见的 messages（对齐 DSH deriveMessages）。
     * 规则：
     *   - user/message → { role: 'user', content }
     *   - assistant/message → { role: 'assistant', content }（chunk 累积）
     *   - assistant/chunk 累积到最近一个 assistant/message（或合并到 user 之后）
     *   - tool/call + tool/result 保留原始块（供回放）
     *   - compaction 摘要 → { role: 'user', _compaction: true }
     * 
     *   - 已被压缩折叠的事件（_excludedSeqs）从投影中剔除 → 压缩真正生效
     *   - tool/result 前先 flushAssistant，保证 assistant(tool_calls) 先于 tool 消息（LLM 合法顺序）
     *   - 每条消息附加 _turnId（供 getHistoryForLLM 排除当前进行中的 turn）
     * @param {string|null} [turnId] - 只投影指定 turn；不传则投影全部
     * @returns {Array<Object>} LLM 可直接使用的 messages
     */
    // ── 跨轮历史正文清洗（喂 LLM 之前瘦身，详见 store.sanitizeHistoryContent 说明）──
    // 围栏代码块（```lang … ```）：上一轮的脚本/代码对下一轮只有误导作用。
    const HISTORY_FENCE_RE = /```[\s\S]*?```/g;
    // 超长行内代码（`…`）：模型常把脚本碎片塞进单反引号，短标识符（列名）不受影响。
    const HISTORY_INLINE_CODE_RE = /`[^`\n]{40,}`/g;
    // 管道收尾套话行：python-pipeline 的「✅ 分析完成！共完成 N 项分析」「⚠️ 部分分析未能完成…」。
    // 这类句子本体没有信息量，却会把下一轮带成"又要出一版分析"的语气。
    const HISTORY_PIPELINE_TAIL_RE = /^[ \t]*(?:✅|⚠️|❌)?[ \t]*\*{0,2}(?:分析完成！?|部分分析未能完成)[^\n]*\*{0,2}[ \t]*$/gm;

    function sanitizeHistoryContent(content) {
        if (!content || typeof content !== 'string') return content;
        const src = content;
        let s = src.replace(HISTORY_FENCE_RE, '[代码块已省略]');
        s = s.replace(HISTORY_INLINE_CODE_RE, '[代码片段已省略]');
        s = s.replace(HISTORY_PIPELINE_TAIL_RE, '');
        s = s.replace(/\n{3,}/g, '\n\n').trim();
        // 安全阀：清洗把整段都掏空时回退原文 —— 宁可啰嗦，也不能丢掉全部上下文
        return s || src;
    }

    function deriveMessages(turnId) {
        const messages = [];
        let pendingChunks = '';
        let pendingToolCalls = [];

        function flushAssistant(force) {
            // 当前累积块或工具调用，压成 assistant 消息
            const hasChunk = pendingChunks.trim().length > 0;
            const hasTools = pendingToolCalls.length > 0;
            if (!hasChunk && !hasTools) return;
            const msg = { role: 'assistant', content: hasChunk ? pendingChunks : null };
            if (hasTools) msg.tool_calls = pendingToolCalls;
            messages.push(msg);
            pendingChunks = '';
            pendingToolCalls = [];
        }

        for (const evt of _events) {
            if (turnId && evt.turnId !== turnId) continue;
            // 跳过已被压缩折叠的事件（投影层收缩，原始事件仍保留审计）
            if (_excludedSeqs.has(evt.seq)) continue;
            // 编辑重发已撤回的轮次：事件留作审计，但不再喂给模型。
            if (evt.turnId && _droppedTurnIds.has(evt.turnId)) continue;
            if (_rewoundSeqs.has(evt.seq)) continue;
            switch (evt.type) {
                case EVENT_TYPES.USER_MESSAGE:
                    flushAssistant();
                    messages.push({ role: 'user', content: evt.data.content || '', _turnId: evt.turnId });
                    break;
                case EVENT_TYPES.ASSISTANT_CHUNK:
                    pendingChunks += evt.data.delta || '';
                    break;
                case EVENT_TYPES.ASSISTANT_MESSAGE:
                    // 最终组装消息取代流式累积：先清掉 chunk/tool_calls，
                    // 否则 flushAssistant() 会把累积内容压成一条、再 push 最终内容，
                    // 造成喂给 LLM 的历史里每轮 assistant 出现两次（小模型会据此复刻重复输出）
                    pendingChunks = '';
                    pendingToolCalls = [];
                    if (evt.data.content !== undefined && evt.data.content !== null) {
                        messages.push({
                            role: 'assistant', content: evt.data.content,
                            tool_calls: evt.data.tool_calls || undefined,
                            _turnId: evt.turnId
                        });
                    }
                    break;
                case EVENT_TYPES.TOOL_CALL:
                    pendingToolCalls.push({
                        id: evt.data.callId,
                        type: 'function',
                        function: { name: evt.data.toolName, arguments: JSON.stringify(evt.data.args || {}) }
                    });
                    break;
                case EVENT_TYPES.TOOL_RESULT:
                    // 先 flush 挂起的 tool_calls，保证
                    // assistant(tool_calls) 位于 tool 消息之前（LLM 合法顺序）
                    flushAssistant();
                    // tool 消息必须在 assistant tool_calls 之后
                    messages.push({
                        role: 'tool',
                        tool_call_id: evt.data.callId,
                        name: evt.data.toolName,
                        content: evt.data.content || '',
                        _turnId: evt.turnId
                    });
                    break;
                case EVENT_TYPES.COMPACTION:
                    // 压缩摘要：作为 user 消息注入（带 provenance 标记）
                    if (evt.data.summary) {
                        flushAssistant();
                        messages.push({
                            role: 'user',
                            content: evt.data.summary,
                            _compaction: true,
                            _provenanceSeq: evt.seq,
                            _turnId: evt.turnId
                        });
                    }
                    break;
                default:
                    break;
            }
        }
        flushAssistant();
        return messages;
    }

    /**
     * 【事件溯源】导出事件日志（调试/回放/持久化）
     * @param {string|null} [turnId]
     * @returns {Array<Object>}
     */
    function getEvents(turnId) {
        if (!turnId) return [..._events];
        return _events.filter(e => e.turnId === turnId);
    }

    /**
     * 【事件溯源】从事件日志重放重建指定 turn 的状态（断点恢复用）
     * @param {string} turnId
     * @returns {Object|null} turn 状态对象或 null
     */
    function replayTurn(turnId) {
        const evts = _events.filter(e => e.turnId === turnId);
        if (evts.length === 0) return null;
        const turn = {
            id: turnId,
            status: 'completed',
            kind: evts[0].data?.kind || 'text',
            userMessage: '',
            attachments: evts[0].data?.attachments || [],
            context: evts[0].data?.context || {},
            model: '',
            reasoning: '',
            content: '',
            toolCalls: [],
            startedAt: evts[0].ts,
            completedAt: null
        };
        for (const evt of evts) {
            switch (evt.type) {
                case EVENT_TYPES.TURN_START:
                    Object.assign(turn, evt.data);
                    break;
                case EVENT_TYPES.USER_MESSAGE:
                    turn.userMessage = evt.data.content || turn.userMessage;
                    break;
                case EVENT_TYPES.ASSISTANT_CHUNK:
                    turn.content += evt.data.delta || '';
                    break;
                case EVENT_TYPES.TOOL_CALL:
                    turn.toolCalls.push({
                        id: evt.data.callId,
                        name: evt.data.toolName,
                        args: evt.data.args || {},
                        status: evt.data.status || 'completed'
                    });
                    break;
                case EVENT_TYPES.TURN_END:
                    turn.status = 'completed';
                    turn.completedAt = evt.ts;
                    if (evt.data?.model) turn.model = evt.data.model;
                    break;
                case EVENT_TYPES.TURN_CANCELED:
                    turn.status = 'canceled';
                    turn.cancelReason = evt.data?.reason;
                    turn.completedAt = evt.ts;
                    break;
                case EVENT_TYPES.TURN_FAILED:
                    turn.status = 'failed';
                    turn.error = evt.data?.error;
                    turn.completedAt = evt.ts;
                    break;
                default:
                    break;
            }
        }
        return turn;
    }

    const STATE = {
        turns: [],          // 最近 N 个 turn 的完整记录
        currentTurn: null,  // 正在进行的 turn
        history: [],       // 持久化对话历史（user/assistant 交替）
        config: {
            maxTurns: 100,
            maxHistoryItems: 100
        }
    };

    // 内部标志：抑制自动保存到旧系统历史（用于加载历史记录时）
    let _suppressSave = false;
    // 内部标志：当前会话是从历史记录加载的（用于 startNewChat 避免重复保存）
    let _loadedFromHistory = false;

    const listeners = new Set();

    function notify(change) {
        listeners.forEach(fn => {
            try { fn(change); } catch(e) {
                ChatLog.error?.('SESSION', '[SessionStore] listener error', e);
            }
        });
    }

    function generateTurnId() {
        return 'turn_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8);
    }

    // ---- 历史富内容持久化的尺寸封顶 ----
    // 目的：历史存储走 localStorage 且 addToHistory 有 100KB 单项上限，
    // 富内容（reasoning / 步骤面板 HTML / 工具卡结果）若不封顶会把整个历史项挤爆而整体跳过保存。
    function capStr(s, max) {
        return typeof s === 'string' ? s.slice(0, max) : s;
    }

    function capToolCalls(arr) {
        const MAX_CALLS = 25;
        const MAX_RESULT = 6000;
        // Skill 卡片结果不能按 6000 截断：skill_card 里内嵌完整 Skill 定义（含 steps 的 VBA code
        // 转义后很长），截断会让它变成非法 JSON → 历史记录恢复时 extractSkillCardFromResult 解析失败
        // → 卡片消失（用户反馈"保存了，但打开历史记录卡片没了"）。
        // 这里对 skill_card 单独放宽上限；100KB 的历史项硬上限仍在，不会撑爆存储。
        const MAX_RESULT_SKILL_CARD = 40000;
        return arr.slice(0, MAX_CALLS).map(tc => {
            const out = { id: tc.id, name: tc.name, args: tc.args, status: tc.status };
            let limit = MAX_RESULT;
            try {
                const preview = typeof tc.result === 'string'
                    ? tc.result
                    : (tc.result != null ? JSON.stringify(tc.result) : '');
                if (preview && preview.indexOf('"skill_card"') >= 0) limit = MAX_RESULT_SKILL_CARD;
            } catch (e) { /* 保持默认上限 */ }
            if (typeof tc.result === 'string') out.result = capStr(tc.result, limit);
            else if (tc.result != null) out.result = capStr(JSON.stringify(tc.result), limit);
            if (tc.error) out.error = capStr(String(tc.error), 2000);
            if (tc.inputParams != null) out.inputParams = tc.inputParams;
            return out;
        });
    }

    /** 封顶 thoughtList，避免超大过程列表把历史项撑爆（与 capToolCalls 同思路）。
     *  注意：app-history 对单条历史项有 100KB 硬上限（超限整条跳过保存），
     *  故块数与单块内容都收得较紧。 */
    /** 按 turnId 定位 turn：优先 currentTurn，再回退到 STATE.turns 全量查找。
     *  【单例 Bug 规避】skill-pipeline 等路径拿到的 turn 对象可能不是 STATE.currentTurn，
     *  只认单例会静默 return（技能步骤/思考块全部丢失），故必须回退全量查找。 */
    function resolveTurnById(turnId) {
        if (STATE.currentTurn && STATE.currentTurn.id === turnId) return STATE.currentTurn;
        for (let i = STATE.turns.length - 1; i >= 0; i--) {
            if (STATE.turns[i] && STATE.turns[i].id === turnId) return STATE.turns[i];
        }
        return null;
    }

    function capThoughtList(arr) {
        const MAX_BLOCKS = 50;
        const MAX_CONTENT = 800;
        return arr.slice(0, MAX_BLOCKS).map(t => {
            if (!t || typeof t !== 'object') return t;
            const out = {
                id: t.id,
                round: t.round,
                kind: t.kind,
                title: capStr(t.title, 200),
                status: t.status,
                toolName: t.toolName,
                durationMs: t.durationMs,
                resultHint: t.resultHint != null ? capStr(String(t.resultHint), 200) : undefined,
                startTs: t.startTs,
                endTs: t.endTs,
                anchor: t.anchor
            };
            if (t.content != null) out.content = capStr(String(t.content), MAX_CONTENT);
            return out;
        });
    }

    const store = {
        // ===== 订阅 =====
        subscribe(listener) {
            listeners.add(listener);
            return () => listeners.delete(listener);
        },

        // ===== 查询 =====
        get(path) {
            if (!path) return STATE;
            return path.split('.').reduce((obj, key) => (obj == null ? obj : obj[key]), STATE);
        },

        getCurrentTurn() { return STATE.currentTurn; },
        getTurns() { return [...STATE.turns]; },

        /** 设置/获取 抑制自动保存标志（用于加载历史记录时） */
        setSuppressSave(val) { _suppressSave = !!val; },
        /** 获取 抑制自动保存标志 */
        getSuppressSave() { return _suppressSave; },
        /** 设置/获取 从历史记录加载标志（用于 startNewChat 避免重复保存） */
        setLoadedFromHistory(val) { _loadedFromHistory = !!val; },
        isLoadedFromHistory() { return _loadedFromHistory; },

        // ===== Turn lifecycle =====
        startTurn({ message, kind, attachments, context }) {
            const turn = {
                id: generateTurnId(),
                status: 'working',     // working | completed | failed | canceled
                kind: kind || 'text',  // text | image | doc
                userMessage: message,
                attachments: attachments || [],
                context: context || {},
                model: App.config?.model || '',  // 初始值为当前配置模型，LLM 响应后覆盖为实际模型
                reasoning: '',
                content: '',
                toolCalls: [],
                startedAt: Date.now(),
                completedAt: null
            };
            STATE.currentTurn = turn;
            STATE.turns.push(turn);
            if (STATE.turns.length > STATE.config.maxTurns) STATE.turns.shift();
            // 记录 turn/start + 首条 user/message
            appendEvent(EVENT_TYPES.TURN_START, turn.id, { kind, attachments, context });
            appendEvent(EVENT_TYPES.USER_MESSAGE, turn.id, { content: message || '' });
            notify({ type: 'turn_start', turn });
            return turn;
        },

        // 设置 LLM 实际使用的模型名（从流式响应中提取）
        setTurnModel(turnId, modelName) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId || !modelName) return;
            turn.model = modelName;
            ChatLog.info?.('SESSION', `[SessionStore] setTurnModel turn=${turnId} model=${modelName}`);
            notify({ type: 'model_set', turnId, model: modelName });
        },

        /**
         * 累计 turn 的 token 用量（每个 LLM 调用结束时由 agent-loop 调用）。
         * tokens: { in, out, estimated? } —— 同一 turn 多轮调用求和。
         * 不 notify：usage 只在终态汇总盒展示，无需触发流式重渲。
         */
        addTurnUsage(turnId, tokens) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId || !tokens || typeof tokens !== 'object') return;
            const _in = Number(tokens.in) || 0;
            const _out = Number(tokens.out) || 0;
            if (!_in && !_out) return;
            if (!turn.usage) turn.usage = { in: 0, out: 0, estimated: false };
            turn.usage.in += _in;
            turn.usage.out += _out;
            // 估算值不覆盖真实统计；一旦有真实数据即视为准确
            if (tokens.estimated && !turn.usage._hasReal) turn.usage.estimated = true;
            if (!tokens.estimated) { turn.usage.estimated = false; turn.usage._hasReal = true; }
        },

        // 设置 turn 状态文本（显示在思考区/状态栏）
        setTurnStatus(turnId, text) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            turn.statusText = text;
            notify({ type: 'status_update', turnId, text });
            // 同一句也写到输入区进度条。技能/代码/Python 的 status 以前只进气泡，
            // 进度条会一直停在「正在准备…」，用户看不到当前在执行什么。
            try {
                const entry = window.App && App.ChatV2 && App.ChatV2.chatEntry;
                if (entry && typeof entry.setPhaseStatus === 'function' && text) {
                    entry.setPhaseStatus(text);
                }
            } catch (_phase) { /* 进度条失败不影响状态写入 */ }
        },

        // 设置 PPT 生成配置（主题 + 版式），触发渲染器显示配置卡
        setPptConfig(turnId, config) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            turn._pptConfig = config;
            notify({ type: 'ppt_config', turnId, config });
        },

        updateTurn(turnId, patch) {
            if (STATE.currentTurn?.id !== turnId) return;
            Object.assign(STATE.currentTurn, patch);
            notify({ type: 'turn_update', turnId, patch });
        },

        // ===== 通用执行上下文（多轮消息继承）=====
        // 执行类管道在 turn 上存 _execSpec（data 快照 + params + result），
        // 供后续消息继承并路由回原管道重新执行。
        // 设计参考：OpenAI Thread（持久会话状态）/ LangGraph checkpointer（状态持久化可恢复）。
        // v20260924：增加 schema 校验 + 工厂函数（buildCommonData 等），
        // 写入方应通过工厂构建 data，保证 selection/referenceRegions 字段统一。
        setExecSpec(turnId, spec) {
            validateExecSpec(spec);
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            turn._execSpec = spec;
            notify({ type: 'turn_update', turnId, patch: { _execSpec: spec } });
        },

        /** 取最近一个有执行上下文的 turn（同会话内倒序查找） */
        getLastExecSpec() {
            for (let i = STATE.turns.length - 1; i >= 0; i--) {
                const t = STATE.turns[i];
                if (t && t._execSpec) {
                    return { spec: t._execSpec, turnId: t.id };
                }
            }
            return null;
        },

        /** 清理执行上下文（用户开新任务时调用，防止跨任务误继承）。
         *  传 turnId 清指定 turn 的 _execSpec；不传则清最近一个有 _execSpec 的 turn。 */
        clearExecSpec(turnId) {
            if (turnId) {
                for (let i = 0; i < STATE.turns.length; i++) {
                    if (STATE.turns[i] && STATE.turns[i].id === turnId) {
                        delete STATE.turns[i]._execSpec;
                        return;
                    }
                }
                return;
            }
            for (let i = STATE.turns.length - 1; i >= 0; i--) {
                if (STATE.turns[i] && STATE.turns[i]._execSpec) {
                    delete STATE.turns[i]._execSpec;
                    return;
                }
            }
        },

        // ===== 流式增量 =====
        appendDelta(turnId, kind, delta) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId || typeof delta !== 'string') return;
            // 【防御】引擎重复推送同一批流式增量时，turn.content/reasoning 会被 += 累积成多份相同副本。
            // 若增量已整体包含在累计串尾部（说明该批内容刚追加过），跳过，避免同份内容被渲染多次。
            if (kind === 'reasoning' || kind === 'content') {
                const cur = kind === 'reasoning' ? turn.reasoning : turn.content;
                if (delta && cur && cur.endsWith(delta)) {
                    _deltaSkip++;
                    return;
                }
            }
            // 【迭代替换模型】reasoning 不再累积到 turn.reasoning 字符串，
            // 而是追加到 thoughtList 里最后一个 running 的 thought 块（每轮一个独立块）
            if (kind === 'reasoning') {
                this.appendThoughtDelta(turnId, delta);
            } else if (kind === 'content') {
                turn.content += delta;
            }
            // 记录 assistant/chunk（逐块追加）
            if (kind === 'content') {
                appendEvent(EVENT_TYPES.ASSISTANT_CHUNK, turnId, { delta });
            }
            notify({ type: 'delta', turnId, kind, delta });
        },

        // ===== 迭代 thoughtList（迭代替换模型） =====
        // 每轮 agent-loop 迭代 → 一个独立 thought 块，状态 running→completed/failed
        // 同 round 的 thought 块原地更新状态和标题，不是 append 新行
        // chat-renderer 渲染时：thoughtList + content 按时间顺序混排到同一个 contentEl
        _ensureThoughtList(turn) {
            if (!turn) return null;
            if (!Array.isArray(turn.thoughtList)) {
                turn.thoughtList = [];
            }
            return turn.thoughtList;
        },

        /** 开始一个新 thought 块（LLM 推理开始 / 工具调用开始都调） */
        startThought(turnId, round, title, opts = {}) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return null;
            const list = this._ensureThoughtList(turn);
            // 若当前已有 running 的 thought 块，先把它标记为 completed（防止并发）
            // 同时替换 title 为完成态措辞：
            //   thinking  "🔍 正在思考..."   → "🔍 思考过程"
            //   routing   "🔍 正在理解..."   → "🔍 意图理解"
            //   tool 块   保持 startToolThought 默认标题不变（completeThought 会精修成 "✅ **tool** ✓(Nms)"）
            for (const t of list) {
                if (t.status === 'running') {
                    t.status = 'completed';
                    t.endTs = Date.now();
                    // thinking 块统一按"标题含'正在'"收敛（agent-loop 建的标题是「正在分析你的需求…」，
                    // 旧逻辑只认'正在思考'导致完成态仍显示"正在..."）
                    if (t.kind === 'thinking' && t.title && t.title.indexOf('正在') >= 0) {
                        t.title = '🔍 思考过程';
                    } else if (t.kind === 'routing' && t.title && t.title.indexOf('正在理解') >= 0) {
                        t.title = '🔍 意图理解';
                    }
                }
            }
            const block = {
                id: 'th_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
                round: round || (list.length + 1),
                title: title || '思考中...',
                content: '',
                status: 'running',
                kind: opts.kind || 'thinking', // 'thinking' | 'tool'
                startTs: Date.now(),
                endTs: null,
                toolName: opts.toolName || null,
                // 锚点：用于迭代替换时定位同一个块（同 round + 同 kind 视为同一块）
                anchor: (round || '') + ':' + (opts.kind || 'thinking') + ':' + (opts.toolName || '')
            };
            list.push(block);
            ChatLog.info?.('THOUGHT', `[startThought] round=${round} kind=${opts.kind} tool=${opts.toolName || '-'} title="${title.slice(0, 30)}" → list size=${list.length}`);
            notify({ type: 'thought_update', turnId });
            return block;
        },

        /** 向最后一个 running 的 thought 块追加 delta（LLM reasoning 流） */
        appendThoughtDelta(turnId, delta) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId || !delta) return;
            const list = this._ensureThoughtList(turn);
            const last = list[list.length - 1];
            if (!last || last.status !== 'running') {
                // 没有 running 块，创建一个默认 thinking 块
                this.startThought(turnId, list.length + 1, '正在思考...', { kind: 'thinking' });
                const t2 = list[list.length - 1];
                if (t2) t2.content += delta;
            } else {
                last.content += delta;
            }
            notify({ type: 'thought_update', turnId });
        },

        /**
         * 替换路由块里的某一句。意图分析先写「正在等模型…」，
         * 模型返回后必须把这句换成结果，否则完成后状态仍停在等待。
         */
        replaceThoughtText(turnId, search, replacement) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId || !search) return false;
            const list = this._ensureThoughtList(turn);
            let target = null;
            for (let i = list.length - 1; i >= 0; i--) {
                const t = list[i];
                if (t && typeof t.content === 'string' && t.content.indexOf(search) >= 0) {
                    target = t;
                    break;
                }
            }
            if (!target) return false;
            let next = target.content.split(search).join(replacement == null ? '' : String(replacement));
            next = next.replace(/\n{2,}/g, '\n').replace(/^\n+|\n+$/g, '');
            target.content = next;
            notify({ type: 'thought_update', turnId });
            return true;
        },

        /** 完成最后一个 running 的 thought 块，可更新标题 */
        completeThought(turnId, title, opts = {}) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            const list = this._ensureThoughtList(turn);
            // 如果传了 round+kind+toolName，精准定位；否则取最后一个 running
            let target = null;
            if (opts.round != null && opts.kind) {
                target = list.find(t =>
                    t.round === opts.round && t.kind === opts.kind &&
                    (opts.toolName ? t.toolName === opts.toolName : true) &&
                    t.status === 'running'
                );
            }
            if (!target) target = [...list].reverse().find(t => t.status === 'running');
            if (!target) return;
            target.status = opts.status || 'completed';
            target.endTs = Date.now();
            if (title) target.title = title;
            ChatLog.info?.('THOUGHT', `[completeThought] round=${target.round} kind=${target.kind} tool=${target.toolName || '-'} → status=${target.status} title="${(title || target.title || '').slice(0, 40)}"`);
            if (opts.durationMs) target.durationMs = opts.durationMs;
            if (opts.resultHint) target.resultHint = opts.resultHint;
            notify({ type: 'thought_update', turnId });
        },

        /** 快速工具 thought：开始一个 kind='tool' 的 running 块，标题默认「🔧 ${toolName} — 执行中」 */
        startToolThought(turnId, round, toolName, paramSummary) {
            const title = paramSummary
                ? `🔧 **${toolName}** — ${paramSummary}`
                : `🔧 **${toolName}** — 执行中...`;
            return this.startThought(turnId, round, title, { kind: 'tool', toolName });
        },

        /**
         * 【步骤统一到 cbx v20261003】把"多步骤 / Skill / 计算机链"的一步写入 thoughtList，
         * 让所有步骤展示都走 contentEl 顶部那套 CodeBuddy 风格 cbx 扁平列表（唯一过程视觉），
         * 不再往 toolsEl 插旧的 .exec-plan-card 绿面板。
         *
         * 以 stepId 为锚点做幂等 upsert：同一步反复更新只改状态/标题/详情，不新增行，
         * 配合渲染层的 thought_update → 全量重渲，实现"边做边展示"的实时进度。
         *
         * @param {string} turnId
         * @param {Object} step - { id, title|name, status, detail|description, toolName?, index? }
         *        status 兼容 done/success/ok → completed，error/err → failed
         * @param {Object} [opts] - { durationMs }
         */
        upsertStepThought(turnId, step, opts = {}) {
            const turn = resolveTurnById(turnId);
            if (!turn || !step) return null;
            const list = this._ensureThoughtList(turn);
            const sid = String(step.id != null ? step.id : (step.title || step.name || ''));
            const raw = String(step.status || 'pending').toLowerCase();
            const status = (raw === 'done' || raw === 'success' || raw === 'ok' || raw === 'completed') ? 'completed'
                : (raw === 'error' || raw === 'err' || raw === 'failed') ? 'failed'
                : (raw === 'running') ? 'running' : 'pending';
            const title = String(step.title || step.name || step.description || '步骤');
            const detail = String(step.detail || step.description || '').trim();

            let block = sid ? list.find(t => t.kind === 'step' && String(t.stepId || '') === sid) : null;
            if (!block) {
                block = {
                    id: 'th_step_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
                    round: (step.index != null ? Number(step.index) + 1 : list.length + 1),
                    title: title,
                    content: detail,
                    status: status,
                    kind: 'step',
                    startTs: Date.now(),
                    endTs: null,
                    stepId: sid,
                    toolName: step.toolName || null,
                    anchor: 'step:' + sid
                };
                list.push(block);
            } else {
                block.title = title;
                // 详情留空时保留上一次内容，避免"已完成但无内容"被渲染层过滤
                block.content = detail || block.content || '';
                block.status = status;
                if (step.toolName) block.toolName = step.toolName;
            }
            if (status === 'completed' || status === 'failed') {
                if (!block.endTs) block.endTs = Date.now();
            }
            if (opts.durationMs != null) block.durationMs = opts.durationMs;
            notify({ type: 'thought_update', turnId });
            return block;
        },

        /** 清空某轮的所有步骤 thought（新一轮开始前调用，避免上一轮步骤残留） */
        clearStepThoughts(turnId) {
            const turn = resolveTurnById(turnId);
            if (!turn) return;
            const list = this._ensureThoughtList(turn);
            for (let i = list.length - 1; i >= 0; i--) {
                if (list[i] && list[i].kind === 'step') list.splice(i, 1);
            }
            notify({ type: 'thought_update', turnId });
        },

        // ===== 工具调用 =====
        addToolCall(turnId, toolCall, opts = {}) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            const tc = {
                id: 'tc_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6),
                status: 'pending',
                startedAt: Date.now(),
                ...toolCall
            };
            turn.toolCalls.push(tc);
            // 记录 tool/call
            appendEvent(EVENT_TYPES.TOOL_CALL, turnId, { callId: tc.id, toolName: tc.name, args: tc.args || {}, status: 'pending' });
            if (!opts.silent) {
                notify({ type: 'tool_call', turnId, toolCall: tc });
            }
            return tc;
        },

        updateToolCall(turnId, toolCallId, patch, opts = {}) {
            const turn = STATE.currentTurn;
            if (!turn?.toolCalls) return;
            const tc = turn.toolCalls.find(t => t.id === toolCallId);
            if (!tc) return;
            Object.assign(tc, patch);
            // 记录 tool/result（完成/失败时）
            if (patch.status === 'completed' || patch.status === 'failed') {
                appendEvent(EVENT_TYPES.TOOL_RESULT, turnId, {
                    callId: toolCallId,
                    toolName: tc.name,
                    status: patch.status,
                    content: patch.status === 'completed'
                        ? (typeof patch.result === 'string' ? patch.result : JSON.stringify(patch.result ?? ''))
                        : (patch.error || ''),
                    error: patch.error || null,
                    durationMs: patch.durationMs || 0
                });
            }
            if (!opts.silent) {
                notify({ type: 'tool_call_update', turnId, toolCallId, patch });
            }
        },

        // ===== 完成 / 失败 / 取消 =====
        completeTurn(turnId, finalState = {}) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            // 【thoughtList 收尾】turn 完成时把所有 remaining running 的 thought 块都置为 completed
            // 并替换标题为完成态措辞（防止最后一个 thinking 块永远停在 running 状态）
            const list = this._ensureThoughtList(turn);
            if (list && list.length) {
                for (const t of list) {
                    if (t.status === 'running') {
                        t.status = 'completed';
                        t.endTs = Date.now();
                        // 同 startThought 收尾：thinking 块按"标题含'正在'"统一收敛
                        if (t.kind === 'thinking' && t.title && t.title.indexOf('正在') >= 0) {
                            t.title = '🔍 思考过程';
                        } else if (t.kind === 'routing' && t.title && t.title.indexOf('正在理解') >= 0) {
                            t.title = '🔍 意图理解';
                        } else if (t.kind === 'tool' && t.title && t.title.indexOf('执行中') >= 0) {
                            t.title = '🔧 ' + (t.toolName || '工具') + ' — 已结束';
                        }
                    }
                }
                notify({ type: 'thought_update', turnId });
            }
            Object.assign(turn, finalState, {
                status: 'completed',
                completedAt: Date.now()
            });
            console.info('[SS] completeTurn id=' + turnId + ' contentLen=' + (turn.content?.length || 0) +
                ' reasoningLen=' + (turn.reasoning?.length || 0) + ' deltaSkip=' + _deltaSkip);
            // 推入持久化历史
            STATE.history.push({
                role: 'user',
                content: turn.userMessage,
                originalContent: turn._originalUserMessage ?? turn.userMessage,
                kind: turn.kind,
                attachments: turn.attachments || [],
                timestamp: turn.startedAt,
                turnId: turn.id
            });
            STATE.history.push({
                role: 'assistant',
                content: turn.content,
                reasoning: turn.reasoning,
                toolCalls: turn.toolCalls,
                thoughtList: turn.thoughtList,
                kind: turn.kind,
                usage: turn.usage ? { in: turn.usage.in, out: turn.usage.out, estimated: !!turn.usage.estimated } : undefined,
                _enableSteps: !!(turn._enableSteps || turn.context?._enableSteps),
                steps: Array.isArray(turn._steps) ? turn._steps : null,
                planHtml: typeof turn._planHtml === 'string' ? turn._planHtml : null,
                timestamp: turn.completedAt,
                turnId: turn.id
            });
            const max = STATE.config.maxHistoryItems * 2;
            if (STATE.history.length > max) {
                STATE.history.splice(0, STATE.history.length - max);
            }
            // 记录 assistant/message（最终组装）+ turn/end
            appendEvent(EVENT_TYPES.ASSISTANT_MESSAGE, turnId, { content: turn.content, tool_calls: undefined, model: turn.model });
            appendEvent(EVENT_TYPES.TURN_END, turnId, { model: turn.model });
            // 在 notify 前清洗 turn.content 中残留的"excel_xxx: "工具名前缀，
            // 确保渲染层（finalizeAIBubble）拿到的是纯结果文本，不再包含工具名。
            // 历史记录中保留原始内容供调试，不影响渲染。
            if (turn.content && typeof turn.content === 'string') {
                turn.content = turn.content.replace(/^\s*excel_[a-zA-Z0-9_]+:\s*/g, '');
            }
            notify({ type: 'turn_done', turnId, turn });
            STATE.currentTurn = null;

            // 自动保存到旧系统历史记录（防抖）
            if (!_suppressSave) {
                try { this.saveToLegacyHistory(); } catch (e) {
                    ChatLog.warn?.('SESSION', '[SessionStore] saveToLegacyHistory 失败', e);
                }
            }

            return turn;
        },

        failTurn(turnId, error) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            turn.status = 'failed';
            turn.error = error;
            turn.completedAt = Date.now();
            // 记录 turn/failed
            appendEvent(EVENT_TYPES.TURN_FAILED, turnId, { error });
            notify({ type: 'turn_failed', turnId, error });
            STATE.currentTurn = null;
        },

        cancelTurn(turnId, reason) {
            const turn = STATE.currentTurn;
            if (!turn || turn.id !== turnId) return;
            turn.status = 'canceled';
            turn.cancelReason = reason;
            turn.completedAt = Date.now();
            // 记录 turn/canceled
            appendEvent(EVENT_TYPES.TURN_CANCELED, turnId, { reason });
            notify({ type: 'turn_canceled', turnId, reason });
            STATE.currentTurn = null;
        },

        // ===== 历史管理 =====
        getHistory() { return [...STATE.history]; },

        /**
         * 获取已完成 turn 的消息列表（为旧系统 addToHistory 提供数据）。
         * 从 STATE.history 中提取 role/content 对，并合并连续的 assistant 消息。
         * @returns {Array<{role: string, content: string}>}
         */
        getCompletedTurnsMessages() {
            // 从 STATE.history 中提取，过滤掉无内容的记录
            const raw = STATE.history.filter(h => h.content && h.content.trim());
            // 合并相邻的 assistant 消息（同一个 AI 回复可能被拆成多条）
            const merged = [];
            for (const item of raw) {
                if (merged.length > 0 && merged[merged.length - 1].role === item.role) {
                    merged[merged.length - 1].content += '\n\n' + item.content;
                    // 合并 assistant：reasoning 串联，工具卡/步骤面板等富字段以首个为准
                    const last = merged[merged.length - 1];
                    if (item.reasoning) last.reasoning = (last.reasoning || '') + '\n\n' + item.reasoning;
                } else {
                    const entry = { role: item.role, content: item.content };
                    if (item.attachments && item.attachments.length > 0) {
                        entry.attachments = item.attachments;
                    }
                    if (item.originalContent) {
                        entry.originalContent = item.originalContent;
                    }
                    // 保留富内容，供历史回放时还原"思考区 / 工具卡 / 步骤面板"（尺寸封顶）
                    if (item.role === 'assistant' || item.role === 'ai') {
                        if (item.reasoning) entry.reasoning = capStr(item.reasoning, 8000);
                        if (Array.isArray(item.toolCalls) && item.toolCalls.length) entry.toolCalls = capToolCalls(item.toolCalls);
                        if (Array.isArray(item.thoughtList) && item.thoughtList.length) entry.thoughtList = capThoughtList(item.thoughtList);
                        if (item.usage && (Number(item.usage.in) || Number(item.usage.out))) entry.usage = item.usage;
                        if (item.kind) entry.kind = item.kind;
                        if (item._enableSteps) entry._enableSteps = true;
                        if (Array.isArray(item.steps) && item.steps.length) entry.steps = item.steps;
                        if (typeof item.planHtml === 'string' && item.planHtml) entry.planHtml = capStr(item.planHtml, 12000);
                    }
                    merged.push(entry);
                }
            }
            return merged;
        },

        /**
         * 跨轮历史正文清洗：只作用于"喂给 LLM 的历史副本"，不动 STATE.history（画面回放不受影响）。
         *
         * 动机（实测 2026-10-10 00:32）：上一轮若是 python 分析管道，assistant 正文里会带
         *   ① 完整脚本的围栏代码块 ②「✅ 分析完成！共完成 N 项分析」这类管道收尾套话。
         * 它们经本函数之前的 getHistoryForLLM 原样喂给下一轮，弱模型（qwen2.5）会照抄上一轮的
         * 套路（"再跑一遍脚本 + 再写一段关键发现/结论"），把一条独立的新命令带偏。
         *
         * 原则：只删"上一轮的实现细节"，保留"上一轮的结论摘要"——
         *   关键发现/结论/汇总表格都是有效上下文，下一轮可能真要引用（"把它做成图表"）。
         *
         * @param {string} content 原始正文
         * @returns {string} 清洗后的正文（清洗掏空时回退原文）
         */
        sanitizeHistoryContent(content) {
            return sanitizeHistoryContent(content);
        },

        getHistoryForLLM() {
            // 改用 deriveMessages 事件投影（含压缩摘要 + 已剔除折叠事件），
            // 让"上下文压缩"真正作用于模型可见历史。
            // 规则：
            //   - 排除当前进行中的 turn（其 USER_MESSAGE 已在 buildMessages 末尾追加，避免重复）
            //   - 只保留 user/assistant/摘要 消息（tool 消息与空 content 的 tool_calls 声明
            //     不进入跨轮历史——与旧行为一致的干净配对，避免孤儿 tool 消息）
            const currentId = STATE.currentTurn?.id || null;
            const projected = deriveMessages();
            const msgs = [];
            for (const m of projected) {
                if (currentId && m._turnId === currentId) continue;
                if (!m || !m.role) continue;
                if (m.role === 'tool') continue;
                if (m.role === 'assistant' && !m.content && Array.isArray(m.tool_calls)) continue;
                // assistant 正文做轻量清洗（剥代码块/管道套话），user 原文一字不动
                const isAssistant = (m.role === 'assistant' || m.role === 'ai');
                msgs.push({ role: m.role, content: isAssistant ? sanitizeHistoryContent(m.content) : m.content });
            }
            return msgs;
        },

        /**
         * 保存当前对话到旧系统历史记录（App.state.chatHistory + 持久化）。
         * 由 completeTurn 自动调用（防抖），也可由 startNewChat 手动调用。
         * @param {boolean} [force=false] - 是否强制保存（忽略防抖）
         */
        saveToLegacyHistory(force) {
            if (!this._saveTimer) this._saveTimer = null;
            if (this._saveTimer && !force) {
                // 已有定时器，重置
                clearTimeout(this._saveTimer);
            }
            const doSave = () => {
                this._saveTimer = null;
                const messages = this.getCompletedTurnsMessages();
                if (messages.length === 0) return;
                // 从第一条用户消息取标题
                const firstUser = messages.find(m => m.role === 'user');
                const title = firstUser
                    ? firstUser.content.replace(/<[^>]+>/g, '').substring(0, 30) || '对话'
                    : '对话';
                // 调用旧系统的 addToHistory（如果存在）
                if (typeof App.addToHistory === 'function') {
                    App.addToHistory(title, messages);
                    ChatLog.info?.('SESSION', `[SessionStore] 已保存 ${messages.length} 条消息到历史记录`);
                } else if (typeof AppHistory !== 'undefined' && typeof AppHistory.addToHistory === 'function') {
                    AppHistory.addToHistory(title, messages);
                    ChatLog.info?.('SESSION', `[SessionStore] 已通过 AppHistory 保存 ${messages.length} 条消息`);
                }
            };
            if (force) {
                // 清除定时器后立即执行
                if (this._saveTimer) clearTimeout(this._saveTimer);
                this._saveTimer = null;
                doSave();
            } else {
                // 防抖 2 秒，避免频繁保存
                this._saveTimer = setTimeout(doSave, 2000);
            }
        },

        /**
         * 删除指定 turnId 及其之后的所有 turn 和对应的 history 记录。
         * 用于编辑重发：用户编辑某条消息后，移除该消息及之后的所有对话。
         * @param {string} turnId - 要删除的起始 turnId（含）
         * @returns {boolean} 是否找到并删除了记录
         */
        removeTurnsSince(turnId) {
            const idx = STATE.turns.findIndex(t => t.id === turnId);
            if (idx === -1) {
                ChatLog.warn?.('SESSION', `[SessionStore] removeTurnsSince: turn ${turnId} not found`);
                return false;
            }

            // 收集被删除的 turnIds
            const removedTurns = STATE.turns.slice(idx);
            const removedIds = removedTurns.map(t => t.id);
            const removedIdSet = new Set(removedIds);
            let cutTs = Infinity;
            for (const t of removedTurns) {
                if (t && t.startedAt && t.startedAt < cutTs) cutTs = t.startedAt;
            }

            // 从这条消息的第一条事件起，整段不再进入模型。
            // 压缩摘要如果包进了这段，摘要本身也撤掉，并把摘要盖住的更早事件放回来。
            const removedSeqs = new Set();
            let cutSeq = Infinity;
            for (const e of _events) {
                if (e && e.turnId && removedIdSet.has(e.turnId)) {
                    removedSeqs.add(e.seq);
                    if (e.seq < cutSeq) cutSeq = e.seq;
                }
            }
            for (const e of _events) {
                if (!e || e.type !== EVENT_TYPES.COMPACTION) continue;
                const meta = (e.data && e.data.meta) || {};
                const excluded = Array.isArray(meta.excludedSeqs) ? meta.excludedSeqs : [];
                const touches = (e.turnId && removedIdSet.has(e.turnId))
                    || (cutSeq < Infinity && e.seq >= cutSeq)
                    || excluded.some(function (s) { return removedSeqs.has(Number(s)); });
                if (!touches) continue;
                _rewoundSeqs.add(e.seq);
                for (const s of excluded) {
                    const seq = Number(s);
                    const src = _events.find(function (ev) { return ev && ev.seq === seq; });
                    if (src && src.turnId && removedIdSet.has(src.turnId)) continue;
                    _excludedSeqs.delete(seq);
                }
            }
            for (const id of removedIds) _droppedTurnIds.add(id);
            STATE.lastRemovedSteps = null;
            for (let i = _stepLogs.length - 1; i >= 0; i--) {
                if (_stepLogs[i] && removedIdSet.has(_stepLogs[i].taskId)) _stepLogs.splice(i, 1);
            }

            STATE.turns.splice(idx);

            const beforeHist = STATE.history.length;
            STATE.history = STATE.history.filter(function (h) {
                if (!h) return false;
                if (h.turnId && removedIdSet.has(h.turnId)) return false;
                if (!h.turnId && cutTs < Infinity && h.timestamp >= cutTs) return false;
                return true;
            });
            const removedHistoryCount = beforeHist - STATE.history.length;

            if (STATE.currentTurn && removedIdSet.has(STATE.currentTurn.id)) {
                STATE.currentTurn = null;
            }

            ChatLog.info?.('SESSION', `[SessionStore] removeTurnsSince: removed ${removedTurns.length} turns, ${removedHistoryCount} history items, from turn ${turnId}`);
            notify({ type: 'turns_removed', turnIds: removedIds });
            return true;
        },

        /**
         * 取出（并清空）编辑/重试时抢救下来的上一轮步骤结果元数据。
         * 供新任务继承注入上下文——步骤读取内容是上下文元数据，不随 turn 移除而丢弃。
         * @returns {{at:number,userText:string,steps:Array}|null}
         */
        consumeLastRemovedSteps() {
            const s = STATE.lastRemovedSteps || null;
            STATE.lastRemovedSteps = null;
            return s;
        },

        clearHistory() {
            STATE.history = [];
            STATE.turns = [];
            _loadedFromHistory = false;
            // 清空事件日志（新会话）
            _events.length = 0;
            _eventSeq = 0;
            _droppedTurnIds.clear();
            _rewoundSeqs.clear();
            // 同步清空压缩排除集合（新会话无历史可排除）
            _excludedSeqs.clear();
            notify({ type: 'history_cleared' });
        },

        // ===== 事件日志 API =====
        /** 追加一条事件（供外部如 compaction 使用） */
        appendEvent(type, turnId, data) {
            if (!type) return null;
            return appendEvent(type, turnId, data);
        },
        /** 追加 compaction 摘要事件（带 provenance） */
        appendCompactionEvent(turnId, summary, meta = {}) {
            const evt = appendEvent(EVENT_TYPES.COMPACTION, turnId, { summary, meta });
            // 把被压缩折叠的事件 seq 标记为"投影排除"，
            // 让 deriveMessages 真正收缩模型可见上下文（原始事件仍保留审计）。
            if (Array.isArray(meta.excludedSeqs)) {
                for (const s of meta.excludedSeqs) {
                    if (Number(s) > 0) _excludedSeqs.add(Number(s));
                }
            }
            notify({ type: 'compaction', turnId, summary, meta });
            return evt;
        },
        /** 从事件日志投影 LLM 消息（DSH deriveMessages 对齐） */
        deriveMessages(turnId) { return deriveMessages(turnId); },
        /** 导出事件日志（调试/回放/持久化） */
        getEvents(turnId) { return getEvents(turnId); },
        /** 获取已压缩折叠事件的 seq 集合（供 compaction 计算待压缩区时过滤，防止重复折叠） */
        getCompactedSeqs() { return _excludedSeqs; },
        /** 重放指定 turn（断点恢复） */
        replayTurn(turnId) { return replayTurn(turnId); },
        /** 事件类型常量 */
        EVENT_TYPES: EVENT_TYPES,

        // ===== 调试日志 API（独立于上下文压缩）=====
        /** 追加步骤日志（thought/tool_call/tool_return/warning/error/finish） */
        appendStepLog(taskId, stepType, content, iteration) {
            return appendStepLog(taskId, stepType, content, iteration);
        },
        /** 查询步骤日志（可按 taskId / stepType 过滤；内存+localStorage 合并，默认返回最近 200 条） */
        queryStepLogs({ taskId, stepType, limit } = {}) {
            let logs = _stepLogs;
            if (taskId) logs = logs.filter(l => l.taskId === String(taskId));
            if (stepType) logs = logs.filter(l => l.stepType === stepType);
            const num = Number(limit) > 0 ? Number(limit) : 200;
            return logs.slice(-num);
        },

        // ===== 调试 =====
        getDebugInfo() {
            return {
                currentTurnId: STATE.currentTurn?.id,
                currentTurnStatus: STATE.currentTurn?.status,
                turnsCount: STATE.turns.length,
                historyCount: STATE.history.length,
                listenersCount: listeners.size,
                // 
                eventCount: _events.length,
                lastEventSeq: _eventSeq
            };
        },

        // ===== ExecSpec 工厂（v20260924 新增）=====
        // 供 pipeline 统一构建 setExecSpec 的 data 字段，消除字段漂移。
        execSpecFactory: {
            buildCommonData: buildCommonData,
            buildCommonSelection: buildCommonSelection,
            buildCommonReferenceRegions: buildCommonReferenceRegions,
            validateExecSpec: validateExecSpec
        }
    };

    App.ChatV2.sessionStore = store;
    ChatLog.info?.('CHAT-V2', '✅ sessionStore loaded');
})();
