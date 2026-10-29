/**
 * ============================================================
 * chat-v2/agent-loop.js — LLM 驱动的工具调用循环
 * ============================================================
 * 职责：执行 LLM → 工具调用 → LLM 的多轮循环（Function Calling）
 * 设计：
 *   - 最多 5 轮工具调用
 *   - 用户可通过 AbortController 中断
 *   - 每个中间状态作为 typed event 流转
 *   - 取消/错误作为事件而非异常
 *   - 内部按 turn/step 状态机运行（参考 DSH agent-loop 架构）：
 *       turn/start → step/start → agent/request → llm/stream → tool/call → tool/result
 *       → step/end → (下一 step 或 turn/end)
 *     所有生命周期经 App.ChatV2.agentEvents 广播（agent/* 事件），UI/调试面板可订阅；
 *     同时保留原有 typed events（round_start/token/...）向后兼容。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 域别名兜底：模型/规划器可能输出泛化工具名（如 "chart"），在此归一为真实 Excel 工具，
    // 避免后端"工具名翻译"把它误译到跨族工具（如 ppt_add_chart_slide）。
    const TOOL_NAME_ALIASES = { chart: 'excel_create_chart' };
    function aliasToolName(name) {
        if (!name) return name;
        return TOOL_NAME_ALIASES[name] || name;
    }

    // 【#6 收敛】错误文本化已统一到 App.ChatV2.toolRegistryUtil.safeErrorText（见 tool-registry.util.js），
    // 不再在此重复实现。运行时经 App.ChatV2.toolRegistryUtil.safeErrorText 引用。

    // 【单一来源】agent 主循环轮次上限。所有 pipeline（text/skill/code...）应引用 agentLoop.MAX_TOOL_ROUNDS，
    // 切勿在别处硬编码不同数值，否则会导致各管道轮次漂移（历史 bug：text-pipeline 曾硬编码 8）。
    const MAX_TOOL_ROUNDS = 5;
    // 同一"工具名+参数"连续失败多少次后熔断（不再重复执行）。
    // 取 2 的理由：第 1 次失败给模型一次自适应机会（它可能换参数/换工具，那签名就变了）；
    // 第 2 次仍用**完全相同**的参数失败，基本可判定为环境/校验问题，重放只会刷屏。
    // 用户明确说"重新/再来"时（_reprobe）不熔断。
    // 单一来源：agent-loop/tool-execution.js（未加载时回退 2）
    const SAME_CALL_FAIL_LIMIT = (App.ChatV2 && App.ChatV2.agentLoopToolExecution
        && Number(App.ChatV2.agentLoopToolExecution.SAME_CALL_FAIL_LIMIT) > 0)
        ? Number(App.ChatV2.agentLoopToolExecution.SAME_CALL_FAIL_LIMIT) : 2;
    // 跨轮冗余防护：把"工具名+参数"压成签名，用于检测模型重复请求同一操作（死循环/反复重刷同一张表）
    function _toolSig(name, args) {
        return _bareToolName(name) + '|' + JSON.stringify(args && typeof args === 'object' ? args : {});
    }
    const MAX_TOOL_RESULT_CHARS = 4000;
    const MAX_FORCE_TOOL_RETRIES = 2;
    const TOOL_MATCH_THRESHOLD = 20;
    // 【要求覆盖度自检】用户一次列多条要求、模型只做了第一条就收尾时的"续做"引导次数上限。
    // 只给 1 次：提示里会一次列全所有未完成项，够用；再多轮只是在推高 token 水位。
    const MAX_COVERAGE_NUDGES = 1;
    // 【图表交付闸门】图表请求"只给文字没出图"时的补图引导次数上限。
    // 同样只给 1 次：弱模型第 2 次仍可能只写文字，此时由循环末尾的兜底合成工具链接管
    // （先补 excel_aggregate 再补 excel_create_chart），不能无限劝。
    const MAX_CHART_NUDGES = 1;
    // 【文件产物交付闸门】"导出 CSV/PDF 却只存了工作簿、还在正文编造产物路径"的补做引导次数上限。
    // 只给 1 次：给了提示仍不调产物工具，就由闸门当场确定性补执行（或如实说明失败），
    // 绝不允许把"没产出"当成"已完成"交付给用户。
    const MAX_FILE_NUDGES = 1;

    // ── 运行状态跟踪（agent/status + turn/step 生命周期）─────────────
    /** @type {Map<string, {phase: string, step: number, status: string}>} */
    const _runStates = new Map();

    // ── agent/pre-step 拦截钩子（参考 DSH agent/pre-step waterfall）───
    /** @type {Array<{name: string, fn: Function, priority: number}>} */
    const _preStepHooks = [];

    /**
     * 注册 agent/pre-step 钩子。
     * 在每次 step 开始后、LLM 请求前运行，可改写消息 / 拒绝 step / 注入上下文。
     * 钩子签名：async (ctx) => { action: 'pass'|'rewrite'|'reject'|'inject', messages?, reason?, injectMessages? }
     *   - pass        — 不干预，继续
     *   - rewrite     — 用 ctx.messages 替换当前消息
     *   - reject      — 拒绝本 step（closeReason 记录 reason）
     *   - inject      — 追加 ctx.injectMessages 到消息（不阻断）
     * 多个钩子按 priority 升序执行，返回的 action 优先级：reject > rewrite > inject > pass（后者覆盖前者时合并 inject）
     * @param {string} name - 钩子名称
     * @param {Function} fn - 钩子函数
     * @param {number} [priority=100]
     * @returns {Function} disposer
     */
    function onPreStep(name, fn, priority = 100) {
        const entry = { name, fn, priority };
        let insertAt = _preStepHooks.length;
        for (let i = 0; i < _preStepHooks.length; i++) {
            if (_preStepHooks[i].priority > priority) { insertAt = i; break; }
        }
        _preStepHooks.splice(insertAt, 0, entry);
        ChatLog.info?.('AGENT', `[agent-loop] 注册 pre-step 钩子: ${name} (priority=${priority})`);
        return () => {
            const idx = _preStepHooks.indexOf(entry);
            if (idx >= 0) _preStepHooks.splice(idx, 1);
        };
    }

    // ═══════════════════════════════════════════════════════════════
    // 通用框架层：布局快照注入 pre-step hook
    // 多步骤场景（persistTools=true）自动调用 excel_get_layout_snapshot，
    // 把当前工作表的已占用区域/图表/KPI/建议锚点注入 system prompt，
    // 让 LLM 知道自己已经放了什么、还能放哪里。
    // 通用能力——任何多步骤复合操作（看板/多图表/复合格式化）受益。
    // ═══════════════════════════════════════════════════════════════
    onPreStep('layout-snapshot', async (ctx) => {
        const { turn, step } = ctx;
        // 1) 条件检查：turn 存在 + 有已执行工具（多步骤场景才需要快照）
        if (!turn) return { action: 'pass' };
        // 看板一键 / _forceDashboard：布局由编排工具锚定，不需要每轮刷 layout_snapshot
        // （实测弱模型会把快照当工具连调 4 次，烧光轮次且不出 Excel 看板）。
        if (turn._oneClickDashboardDone || (turn.context && turn.context._forceDashboard)) {
            return { action: 'pass' };
        }
        const executedCount = Array.isArray(turn._executedTools) ? turn._executedTools.length : 0;
        if (executedCount === 0) return { action: 'pass' };
        // 2) 频率控制：每 2 轮调一次（节省 token + VB 调用时间）
        //    step 从 1 开始，第 2、4、6... 轮调
        if (step % 2 !== 0) return { action: 'pass' };
        // 3) 调 VB 端 excel_get_layout_snapshot（通过 CallLocalTool 路由）
        try {
            const bridge = App.excelBridge;
            if (!bridge || typeof bridge.CallLocalTool !== 'function') return { action: 'pass' };
            const argsJson = JSON.stringify({ sheet: '', includeHeader: true, includeKpiCells: true });
            const callJson = await bridge.CallLocalTool('excel_get_layout_snapshot', argsJson);
            const call = JSON.parse(callJson);
            if (!call || !call.success) return { action: 'pass' };
            // 从 MCP content[0].text 提取真正的 snapshot JSON
            const text = Array.isArray(call.content) && call.content.length > 0 ? call.content[0].text : '';
            if (!text) return { action: 'pass' };
            const snap = JSON.parse(text);
            // 4) 格式化成简洁文字（控制 token 量，只保留关键信息）
            const lines = [];
            lines.push('=== 当前工作表布局快照 ===');
            lines.push(`工作表: ${snap.sheet || '未知'} | UsedRange: ${snap.usedRange || '空'} (${snap.usedRows || 0}行 × ${snap.usedCols || 0}列)`);
            if (Array.isArray(snap.charts) && snap.charts.length > 0) {
                lines.push(`已创建图表(${snap.charts.length}个):`);
                for (const c of snap.charts) {
                    lines.push(`  • ${c.name} [${c.type || '?'}] "${c.title || ''}" @${c.anchor || '?'} (${c.widthPt || '?'}×${c.heightPt || '?'}) 数据源:${c.dataSource ? c.dataSource.slice(0, 60) : '?'}`);
                }
            }
            if (Array.isArray(snap.headerRow) && snap.headerRow.length > 0) {
                lines.push(`表头: ${snap.headerRow.slice(0, 10).join(' | ')}${snap.headerRow.length > 10 ? ' ...' : ''}`);
            }
            if (Array.isArray(snap.kpiCells) && snap.kpiCells.length > 0) {
                lines.push(`已写单元格(${snap.kpiCells.length}个):`);
                for (const k of snap.kpiCells.slice(0, 10)) {
                    lines.push(`  ${k.cell} = ${k.value ? String(k.value).slice(0, 30) : ''}`);
                }
            }
            if (snap.suggestedNextAnchor) {
                lines.push(`建议下一个锚点: ${snap.suggestedNextAnchor}（在 ${snap.usedRangeBottom || 'UsedRange末尾'} 之后）`);
            }
            lines.push('=== 快照结束 ===');
            const injectText = lines.join('\n');
            ChatLog.info?.('AGENT', `[layout-snapshot] step ${step}: ${snap.charts?.length || 0}图表, ${snap.kpiCells?.length || 0}单元格, 建议锚点 ${snap.suggestedNextAnchor || '?'}`);
            return {
                action: 'inject',
                injectMessages: [{ role: 'system', content: injectText }]
            };
        } catch (ex) {
            ChatLog.warn?.('AGENT', `[layout-snapshot] 失败（已忽略）: ${ex.message}`);
            return { action: 'pass' };
        }
    }, 50);  // priority=50，在其他 hook 之前执行（先注入事实）

    // ═══════════════════════════════════════════════════════════════
    // 通用框架层：结构化产物摘要注入 pre-step hook
    // 每轮自动把 turn._artifacts（VB 端工具返回的 Artifact 字段）注入 system prompt
    // 让 LLM 精确知道自己已经创建了什么（图表名称/位置、聚合结果在哪张表、写了哪些 KPI）
    // 通用能力——任何多步骤复合操作（看板/多图表/复合格式化）受益。
    // 和 layout-snapshot（VB 端实时快照）互补——artifact 是工具返回时自带的事实，
    // snapshot 是 VB 端实时扫描当前状态。
    // ═══════════════════════════════════════════════════════════════
    onPreStep('artifacts-summary', async (ctx) => {
        const { turn, step } = ctx;
        if (!turn) return { action: 'pass' };
        if (!Array.isArray(turn._artifacts) || turn._artifacts.length === 0) return { action: 'pass' };
        // 每轮都注入（artifact 是累积事实，不费多少 token）
        const arts = turn._artifacts;
        const lines = [];
        lines.push(`=== 本轮已创建的结构化产物（共 ${arts.length} 个）===`);
        for (const a of arts) {
            if (a.type === 'chart') {
                lines.push(`  📊 图表: "${a.chartName}" [${a.chartType || '?'}] 标题="${a.title || ''}" @${a.anchor || '?'} 数据源=${a.dataRange ? String(a.dataRange).slice(0, 50) : '?'}`);
            } else if (a.type === 'aggregate') {
                const _gc = Number(a.groupCount || 0);
                lines.push(`  📈 聚合: 维度=${a.dimColumn || '?'} 数值=${a.valueColumn || '?'} 函数=${a.aggFunc || '?'} 组数=${_gc || '?'} 落地表=${a.targetSheet || '(未落地)'}`);
                if (a.dimSwitched) {
                    lines.push(`     ⚙️ 维度已自动择优：请求「${a.requestedDimColumn || '?'}」在本次数据中无区分度，实际按「${a.dimColumn || '?'}」聚合，总结里必须向用户说明这一点。`);
                }
                if (_gc === 1) {
                    // 【低区分度闸门】1 组的聚合建图 = 一根柱/一个扇区的废图，不能让模型照常收尾
                    lines.push('  ⚠️ 该聚合只有 1 组：该维度列在本次数据里没有区分度，基于它建的图只有一根柱/一个扇区，属于废图。');
                    const _cands = Array.isArray(a.dimCandidates) ? a.dimCandidates : [];
                    const _good = _cands
                        .filter(c => c && Number(c.distinct) >= 2)
                        .sort((x, y) => Number(y.distinct) - Number(x.distinct))
                        .slice(0, 5);
                    if (_good.length > 0) {
                        lines.push('     必须换维度重算 excel_aggregate，可选维度（列名/取值个数）：' + _good.map(c => `${c.column}(${c.distinct})`).join('、'));
                    } else {
                        lines.push('     必须换一个有多个取值的维度列重算 excel_aggregate 后再建图。');
                    }
                    lines.push('     严禁把这组只有 1 行的结果直接建图，或直接当作完成结果总结给用户。');
                }
            } else {
                // 通用兜底：输出所有 key
                const pairs = Object.entries(a).filter(([k]) => !k.startsWith('_')).map(([k, v]) => `${k}=${typeof v === 'object' ? JSON.stringify(v) : v}`);
                lines.push(`  🔧 ${a.type || 'tool'}: ${pairs.join(', ')}`);
            }
        }
        // 【数据血缘追踪 v2】注入血缘链
        const lineage = Array.isArray(turn._lineage) ? turn._lineage : [];
        if (lineage.length > 0) {
            lines.push(`--- 数据血缘链（${lineage.length} 条依赖）---`);
            for (const e of lineage) {
                lines.push(`  🔗 aggregate@${e.fromSheet}(${e.fromDataRange}) → 📊 chart "${e.toChartName}" @${e.toAnchor || '?'}`);
            }
            // 未消费的 aggregate 提示
            const unconsumed = arts.filter(a => a.type === 'aggregate' && !a._consumed);
            if (unconsumed.length > 0) {
                lines.push(`  ⏳ 还有 ${unconsumed.length} 个聚合结果未被图表消费:`);
                for (const u of unconsumed) lines.push(`    • ${u.targetSheet || '?'} (${u.dimColumn || '?'}) → 可创建新图表`);
            }
        }
        lines.push('=== 产物清单结束 ===');
        const injectText = lines.join('\n');
        ChatLog.info?.('AGENT', `[artifacts-summary] step ${step}: ${arts.length} 个产物`);
        return {
            action: 'inject',
            injectMessages: [{ role: 'system', content: injectText }]
        };
    }, 60);  // priority=60，在 layout-snapshot(50) 之后

    /**
     * 运行 agent/pre-step 钩子链。
     * @param {Object} ctx - { turnId, turn, messages, step, signal }
     * @returns {Promise<{action: string, messages?: Array, reason?: string, injectMessages?: Array}>}
     */
    async function runPreStepHooks(ctx) {
        if (_preStepHooks.length === 0) return { action: 'pass' };
        let result = { action: 'pass', messages: ctx.messages };
        const injects = [];
        for (const h of _preStepHooks) {
            try {
                const r = await h.fn(ctx);
                if (!r || typeof r !== 'object') continue;
                // reject 最高优先级，一旦 reject 立即终止（单调）
                if (r.action === 'reject') {
                    ChatLog.warn?.('AGENT', `[agent-loop] pre-step 钩子 ${h.name} 拒绝了 step: ${r.reason || 'no reason'}`);
                    return { action: 'reject', reason: r.reason || 'rejected_by_pre_step' };
                }
                if (r.action === 'rewrite' && Array.isArray(r.messages)) {
                    result = { action: 'rewrite', messages: r.messages };
                } else if (r.action === 'inject' && Array.isArray(r.injectMessages)) {
                    injects.push(...r.injectMessages);
                }
            } catch (e) {
                ChatLog.warn?.('AGENT', `[agent-loop] pre-step 钩子 ${h.name} 异常（已忽略）:`, e);
            }
        }
        if (result.action === 'rewrite') {
            // rewrite + inject 合并
            if (injects.length > 0) result.messages = [...result.messages, ...injects];
            return result;
        }
        if (injects.length > 0) {
            return { action: 'inject', messages: ctx.messages, injectMessages: injects };
        }
        return { action: 'pass' };
    }

    /** 获取 agentEvents（延迟解析，避免模块加载顺序问题） */
    function getAgentEvents() {
        return App.ChatV2?.agentEvents || null;
    }

    /** 广播 agent/* 事件（agentEvents 未加载时静默跳过） */
    function emitAgent(event, payload) {
        try {
            getAgentEvents()?.emit(event, payload);
        } catch (e) {
            ChatLog.warn?.('AGENT', `[agent-loop] emit ${event} 失败:`, e);
        }
    }

    /** 更新运行状态并广播 agent/status（status / phase / 细粒度字段变化时） */
    function setRunStatus(turnId, status, extra = {}) {
        const st = _runStates.get(turnId);
        if (!st) return;
        const prev = {
            status: st.status,
            phase: st.phase,
            step: st.step,
            toolName: st.toolName,
            detail: st.detail,
            displayToolName: st.displayToolName,
            stepTotal: st.stepTotal
        };
        st.status = status;
        Object.assign(st, extra);
        const changed = prev.status !== status
            || (extra.phase !== undefined && st.phase !== prev.phase)
            || (extra.step !== undefined && st.step !== prev.step)
            || (extra.toolName !== undefined && st.toolName !== prev.toolName)
            || (extra.detail !== undefined && st.detail !== prev.detail)
            || (extra.displayToolName !== undefined && st.displayToolName !== prev.displayToolName)
            || (extra.stepTotal !== undefined && st.stepTotal !== prev.stepTotal);
        if (changed) {
            emitAgent('agent/status', {
                turnId,
                status,
                phase: st.phase,
                step: st.step,
                stepTotal: st.stepTotal || null,
                toolName: st.toolName || null,
                displayToolName: st.displayToolName || null,
                detail: st.detail || null
            });
        }
    }

    function resolveToolDisplayName(toolName) {
        if (!toolName) return '';
        try {
            const tool = App.ChatV2 && App.ChatV2.toolRegistry && typeof App.ChatV2.toolRegistry.get === 'function'
                ? App.ChatV2.toolRegistry.get(toolName)
                : null;
            if (tool && tool.displayName) return String(tool.displayName);
        } catch (_e) { /* ignore */ }
        return String(toolName).replace(/^excel[_-]?/i, '').replace(/_/g, ' ');
    }

    /** 查询指定 turn 的实时运行状态（供外部订阅/调试） */
    function getRunStatus(turnId) {
        const st = _runStates.get(turnId);
        if (!st) return null;
        return { turnId, ...st };
    }

    /** 查询所有活跃运行（调试面板用） */
    function listRunStates() {
        return Array.from(_runStates.entries()).map(([turnId, st]) => ({ turnId, ...st }));
    }

    // ═══════════════════════════════════════════════════════════════
    // 对齐 multi-step-agent-mcp 的
    // token_estimate + context_compress 水位模型。
    // 原则：只做估算与告警，触发压缩下沉到 compaction 服务（不在 LLM 内部做 heavy 逻辑）。
    // ═══════════════════════════════════════════════════════════════

    /**
     * 估算 messages(+tools) 的 token 用量（借助 promptBuilder.estimateTokens）。
     * @param {Array} messages - LLM 消息数组
     * @param {Array} [tools] - 本轮可见工具列表
     * @returns {number|null} 估算 token 数；估算器不可用时返回 null（跳过水位监控）
     */
    function estimateMessagesTokens(messages, tools) {
        const est = App.ChatV2?.promptBuilder?.estimateTokens;
        if (typeof est !== 'function') return null;
        let total = 0;
        for (const m of messages || []) {
            const c = m?.content;
            if (c != null) total += est(typeof c === 'string' ? c : JSON.stringify(c));
            if (Array.isArray(m?.tool_calls)) total += est(JSON.stringify(m.tool_calls));
        }
        if (Array.isArray(tools) && tools.length > 0) total += est(JSON.stringify(tools));
        return total;
    }

    /**
     * 每轮 LLM 调用前执行 token 水位检查：
     *   - ratio >= criticalRatio → 强制压缩（level=critical）
     *   - ratio >= warnRatio     → 触发压缩（level=warn）
     *   - 否则跳过（level=none）
     * 触发压缩时通过 sessionStore.appendStepLog 落 warning 步骤日志（独立于 LLM 上下文）。
     * @returns {Promise<{level: 'none'|'warn'|'critical', total?: number, windowMax?: number, ratio?: number, compacted?: boolean}>}
     */
    async function checkTokenWatermark({ turnId, sessionStore, messages, tools, tokenGuard, iteration }) {
        const compaction = App.ChatV2?.compaction;
        if (!sessionStore || !compaction || typeof compaction.compactNow !== 'function') {
            return { level: 'none' };
        }
        const total = estimateMessagesTokens(messages, tools);
        if (total == null) return { level: 'none' };

        // 窗口上限：tokenGuard.maxTokens > 上下文窗口 > 默认 8192
        let windowMax = 8192;
        try {
            const cw = App.ChatV2?.promptBuilder?.getContextWindow?.();
            if (Number(cw) > 0) windowMax = Number(cw);
        } catch (e) { /* 忽略 */ }
        if (Number(tokenGuard?.maxTokens) > 0) windowMax = Number(tokenGuard.maxTokens);

        const ratio = windowMax > 0 ? total / windowMax : 0;
        const warnRatio = Number(tokenGuard?.warnRatio) > 0 ? Number(tokenGuard.warnRatio) : 0.7;
        const criticalRatio = Number(tokenGuard?.criticalRatio) > 0 ? Number(tokenGuard.criticalRatio) : 0.9;

        if (ratio < warnRatio) return { level: 'none', total, windowMax, ratio };

        // 达到告警/危险阈值 → 触发会话上下文压缩
        let compacted = false;
        try {
            const res = await compaction.compactNow(sessionStore, { turnId });
            compacted = !!res?.compacted;
        } catch (e) {
            ChatLog.warn?.('AGENT', '[agent-loop] token 水位触发压缩异常:', e?.message);
        }

        const level = ratio >= criticalRatio ? 'critical' : 'warn';
        const levelText = level === 'critical' ? '危险' : '告警';
        ChatLog.warn?.('AGENT', `[agent-loop] token 水位${levelText}: ${total}/${windowMax} (${(ratio * 100).toFixed(1)}%) 已压缩=${compacted}`);
        // 落步骤日志（warning 类型，独立于上下文压缩）
        sessionStore.appendStepLog?.(turnId, 'warning',
            `token 水位${levelText}：${total}/${windowMax}（${(ratio * 100).toFixed(1)}%）${compacted ? '，已触发上下文压缩' : '，压缩未生效'}`, iteration);

        return { level, total, windowMax, ratio, compacted };
    }

    /**
     * 危险水位时的本地上下文折叠：保留 system + 最近 keepRounds 轮，更早历史折叠为一条 system 摘要。
     * 与 compaction（会话级）互补——compaction 管会话历史，这里管当前 loop 的 currentMessages。
     * @param {Array} messages
     * @param {number} keepRounds
     * @returns {Array} 折叠后的消息数组（无变化时返回原引用）
     */
    function _localCompressMessages(messages, keepRounds) {
        if (!Array.isArray(messages) || messages.length < 6) return messages;
        const keepCount = Math.max(4, (Number(keepRounds) > 0 ? Number(keepRounds) : 2) * 2);
        if (messages.length <= keepCount) return messages;
        const systemPart = messages[0]?.role === 'system' ? [messages[0]] : [];
        const keepPart = messages.slice(-keepCount);
        const toFold = messages.slice(systemPart.length, messages.length - keepPart.length);
        const lines = [];
        for (const m of toFold) {
            const role = m?.role;
            const c = typeof m?.content === 'string' ? m.content.slice(0, 120) : '';
            if (role === 'user') lines.push(`[用户] ${c}`);
            else if (role === 'assistant' && Array.isArray(m.tool_calls)) lines.push(`[助手调用工具] ${m.tool_calls.map(t => t?.function?.name).filter(Boolean).join(', ')}`);
            else if (role === 'assistant') lines.push(`[助手] ${c}`);
            else if (role === 'tool') lines.push(`[工具返回] ${c.slice(0, 80)}`);
        }
        const folded = lines.filter(Boolean).join('\n').slice(0, 1500);
        if (!folded) return messages;
        return [...systemPart, { role: 'system', content: `【早期上下文摘要】更早的对话/工具记录已压缩，请基于以下摘要继续：\n${folded}` }, ...keepPart];
    }

    //  KEYWORD_SCORE_BOOST + FORCE_RETRY_SKIP_TOOLS + scoreTool 已全部移除
    // 意图匹配统一走后端 McpCapabilityRegistry.Match()，不再依赖前端硬编码词表

    /**
     * 【承诺未执行检测】模型正文里是否"点名了某个 excel_* 工具 + 未来时承诺调用"。
     * 只在"本轮零工具调用"时作为判据使用——它比"用用户消息打分"精确得多：
     * 跟进型消息可能极短（"不够美观"），既不含 EXCEL_OPERATION_INTENT_RE 任何词
     * （userOpIntent=false，整套引导/兜底被设计性跳过），打分也过不了弱匹配 5 分线，
     * 于是直接判定 no_tool_calls，UI 却显示"任务完成 · 未调用工具"。
     * 实测：模型写下"接下来我将使用 excel_format_range 工具进行美化处理。"却什么都没调。
     *
     * @param {string} content - 本轮助手正文
     * @param {Array} tools - 本轮实际下发给模型的工具列表
     * @returns {{name:string, visible:boolean}|null} 点名的工具及它本轮是否可见
     */
    function detectNamedToolCommitment(content, tools) {
        if (!content || typeof content !== 'string') return null;
        // 只认"第一人称/未来时"的承诺（"我将/我会/我们将/接下来我/马上/准备…"），
        // 不认泛化的"调用/执行"——否则"你可以用 excel_format_range 来设置边框"这类
        // 建议型回答会被误判成承诺，白跑一轮重试。
        // 【2026-10-28】原正则只认单数"我将/我会"，漏掉复数句式：实测模型写
        // "现在我们将当前工作表导出为 CSV 文件，并保存到桌面。\n\n调用 `excel_export_csv` 工具导出 CSV 文件。"
        // ——"我们将"不匹配 "我将" → 整句被判成非承诺 → 静默收尾，用户只看到永不执行的计划。
        // 主语扩到 我/我们/咱们/咱，并补"接下来/下一步/随后/下面我"等时间副词句式。
        const COMMIT_RE = /((?:我|我们|咱们|咱)\s*(?:现在|接下来|下一步|下面)?\s*(?:将|会|要|来|去|打算|计划)|接下来|下一步|随后|下面我|马上|这就|准备|开始|继续(?:调用|使用|执行))/;
        // 过去式陈述（"我已调用 excel_read_range 得到结果"）不是承诺，跳过——假执行另有 FAKE_EXECUTION_PATTERN 处理
        const PAST_RE = /(已(?:经)?\s*(?:调用|执行|使用)|调用过|用过了)/;
        // 必须覆盖全部宿主工具前缀：此前只认 excel_，导致 Word/PPT 里模型"说要调用 word_/ppt_ 工具却没调"
        // 无法被识别，兜底重试失效（用户截图里模型干说"将使用 word_set_table_format"就是此因）。
        const TOOL_RE = /\b((?:excel|word|ppt)_[a-z0-9_]+)\b/gi;
        const visible = new Set(
            (tools || []).map(t => String((t && (t.name || t.originalName)) || '').toLowerCase()).filter(Boolean));
        let m;
        while ((m = TOOL_RE.exec(content)) !== null) {
            const name = m[1].toLowerCase();
            // 承诺必须出现在工具名"附近"（前 80 字符内），避免正文里顺带提及工具名被误判成承诺
            const win = content.slice(Math.max(0, m.index - 80), m.index);
            if (PAST_RE.test(win)) continue;
            if (!COMMIT_RE.test(win)) continue;
            return { name, visible: visible.has(name) };
        }
        return null;
    }

    /**
     * 【可观测性】把关键决策写进宿主日志文件（%LOCALAPPDATA%\ExcelAIVSTOAddIn\Logs\chat_log_*.txt）。
     *
     * 背景（这是一次真实的排查教训）：前端 ChatLog 只写 console（app-bootstrap.js 里
     * `App.ChatLog = { ... console.* }`），WebView2 里根本看不到，JS 侧的任何决策
     * 在生产环境都不可观测。上一轮"要求覆盖度自检"因为 intentKeywords 在工具注册时被
     * 丢弃而完全静默空转，却没有任何日志能暴露它——只能靠翻代码反复猜测。
     * ExcelBridge.LogMessage(level, message) 会经 C# 侧 ChatLogger 落盘（tag [FRONTEND]），
     * 因此守卫一类的"关键判定"必须走这条路。
     */
    // 供 chart-gate 等子模块复用同一宿主日志通道
    function _hostLog(level, message) {
        try {
            const b = (typeof App !== 'undefined' && App.excelBridge)
                || (typeof window !== 'undefined' && window.excelBridge);
            if (b && typeof b.LogMessage === 'function') {
                const r = b.LogMessage(level || 'INFO', `[agent-loop] ${message}`);
                if (r && typeof r.catch === 'function') r.catch(function () { /* 宿主日志失败可忽略 */ });
            }
        } catch (_e) { /* 宿主日志失败绝不影响主流程 */ }
    }
    App.ChatV2.agentLoopHostLog = _hostLog;

    // 【工作表名清单·2026-10-01】宿主真实工作表清单缓存。
    // 供 sheet 参数确定性纠偏使用；3 秒 TTL + 取不到时沿用上一次结果，
    // 避免每次工具调用都打一次 COM（GetSheets 本身很轻，但没必要重复取）。
    let _sheetListCache = { at: 0, names: [] };
    async function getRealSheetNames() {
        const now = Date.now();
        if (_sheetListCache.names.length > 0 && now - _sheetListCache.at < 3000) return _sheetListCache.names;
        try {
            const b = (typeof App !== 'undefined' && App.excelBridge)
                || (typeof window !== 'undefined' && window.excelBridge)
                || (typeof window !== 'undefined' && window.chrome && window.chrome.webview
                    && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
            if (!b || typeof b.GetSheets !== 'function') return _sheetListCache.names;
            const raw = await b.GetSheets();
            let arr = [];
            try { arr = JSON.parse(raw); } catch (_e) { arr = []; }
            if (Array.isArray(arr) && arr.length > 0) _sheetListCache = { at: now, names: arr };
        } catch (_e) { /* 取不到清单 → 沿用缓存（空则跳过纠偏，不误伤） */ }
        return _sheetListCache.names;
    }

    /** 去掉 <source>-local__ 前缀，用于跨来源比较工具名 */
    function _bareToolName(n) {
        const s = String(n || '');
        const i = s.indexOf('__');
        return i > 0 ? s.slice(i + 2) : s;
    }

    // ── 【等价完成】覆盖度自检的"字面词陷阱"修正 ─────────────────────────
    // 实测误报（截图 20:37，Docs/公式能力确认与复杂公式测试提示词.md 实测案例 5）：
    // 用户说"在H2写入公式…订单金额合计"，模型用 excel_add_formula 写入 =SUMIFS(...)
    // 一步完成了"写入 + 聚合"两件事；但自检按关键词字面映射
    // （"写入"→excel_write_range、"合计"→excel_aggregate）只认"这两个工具名是否被调用过"，
    // 于是误判"没执行"→ 续做引导 → 收尾正文追加"⚠️ 你要求的以下内容本次【没有执行】"——
    // 任务明明已完成，用户却看到"还有下一步"。
    // 这里给自检加一层等价完成：某些工具执行后，语义上已覆盖另一工具的意图。
    // 保守策略：只列高置信规则；"分组/每个/排名"类多行结果一条公式装不下，仍算缺项。
    const _AGG_FORMULA_RE = /^\s*=?\s*(?:SUMIFS?|COUNTIFS?|COUNTA?|AVERAGEIFS?|AVERAGE|SUBTOTAL|AGGREGATE|MAXIFS?|MINIFS?|MAX|MIN|SUMPRODUCT)\s*\(/i;
    // 用户要的是"按X分组/每个X/排名"的多行结果 → 单条公式无法承载，不认等价
    const _MULTI_GROUP_RE = /分组|每个|各类|各项|各月|各地区|各部门|各区域|分布|排名|排行/;
    /** 解析 format_range 的 format 实参（可能是对象或 JSON 字符串） */
    function _parseFormatArgs(call) {
        let fmt = call && call.args && call.args.format;
        if (typeof fmt === 'string') {
            try { fmt = JSON.parse(fmt); } catch (_e) { fmt = null; }
        }
        return (fmt && typeof fmt === 'object') ? fmt : null;
    }

    /** 仅边框的残缺 format（实测"只做了边框就宣称完成"）——不豁免 excel_autofit */
    function _isBorderOnlyFormat(fmt) {
        if (!fmt) return false;
        const keys = Object.keys(fmt);
        if (keys.length === 0) return false;
        return keys.every(k => k === 'borderStyle' || k === 'borderColor' || k === 'border');
    }

    const _TOOL_EQUIVALENCE = [
        {
            // 写公式进单元格 = 内容已经"写入"了工作表
            covering: 'excel_add_formula',
            covers: ['excel_write_range'],
            // 公式本身是聚合函数时，"聚合统计"的意图也已由公式完成（单值聚合）
            coversIf: {
                excel_aggregate: (call, msg) =>
                    _AGG_FORMULA_RE.test(String((call && call.args && call.args.formula) || ''))
                    && !_MULTI_GROUP_RE.test(msg)
            }
        },
        {
            // 写单元格值同样是"写入"
            covering: 'excel_set_cell_value',
            covers: ['excel_write_range']
        },
        {
            // 区域填充数据本身就是"写入"
            covering: 'excel_fill_data',
            covers: ['excel_write_range']
        },
        {
            // excel_format_range 后端会 AutoFit 列宽（空 format→商务蓝预设；显式 format→末尾统一 AutoFit）。
            // 工具本身 intentKeywords 也含「列宽」。实质美化后不得再误报 excel_autofit 未执行。
            // 例外：仅 borderStyle/borderColor 的残缺调用（多步美化回归用例）仍要求补 excel_autofit。
            covering: 'excel_format_range',
            coversIf: {
                excel_autofit: (call) => {
                    const fmt = _parseFormatArgs(call);
                    if (!fmt || Object.keys(fmt).length === 0) return true; // 预设路径
                    if (_isBorderOnlyFormat(fmt)) return false;
                    return true;
                }
            }
        }
    ];

    /** 判定"缺项工具 name"是否已被已执行工具等价完成（见 _TOOL_EQUIVALENCE 注释） */
    function _isCoveredByEquivalence(name, executedCalls, msg) {
        for (const rule of _TOOL_EQUIVALENCE) {
            const covering = _bareToolName(rule.covering);
            const call = (executedCalls || []).find(c => c && c.name === covering);
            if (!call) continue;
            if ((rule.covers || []).some(c => _bareToolName(c) === name)) return true;
            const cond = rule.coversIf && rule.coversIf[name];
            if (typeof cond === 'function' && cond(call, msg)) return true;
        }
        return false;
    }

    /**
     * 【要求覆盖度自检】把用户消息拆成"要求项"，用工具自带的 intentKeywords
     * （excel_tools.json 维护，单一数据源）反查每条要求该由哪个工具完成，
     * 再与本题已执行过的工具比对，返回"用户明确要求、但本题还没做"的工具。
     *
     * 实测失败（截图）：用户一条消息列了 5 条要求——
     *   "美化当前选中的表格区域: 设置表头底色、边框、对齐、居中、自动列宽，使其清晰美观"
     * 模型只调了 1 次 excel_format_range（format 里只有 borderStyle/borderColor），
     * 这在 schema 层面完全合法（required 只有 range/format）→ 直接"任务完成"。
     * 已有的三道门（工具可见性 / 承诺未执行 / 假执行文字）都只管"有没有调工具"，
     * 没有一道管"调够没调够"，所以"只做第一条"正好从缝里漏过去。
     *
     * 保守策略：要求项少于 3 条直接放弃判定（单步请求不该被当成多步），
     * 只认工具自己声明的 intentKeywords，不硬编码任何工具名。
     *
     * @param {Object} turn - 当前 turn（读 userMessage 与 _executedTools）
     * @param {Array} tools - 本轮下发的可见工具
     * @returns {Array<{name:string, displayName:string, keywords:string[]}>} 未覆盖的工具
     */
    function detectUncoveredRequirements(turn, tools) {
        try {
            const msg = String((turn && (turn.userMessage || turn.message)) || '');
            if (!msg) return [];
            // 拆要求项：顿号/逗号/分号/换行/冒号（"美化表格: 底色、边框" → 三项）
            const items = msg.split(/[、，,;；:：\n]+/).map(s => s.trim()).filter(s => s.length >= 2);
            if (items.length < 3) return [];
            // 优先读 _executedToolCalls（含参数，供等价完成按公式实参判定）；旧会话只有名字列表时降级
            const executedCalls = Array.isArray(turn._executedToolCalls)
                ? turn._executedToolCalls.map(c => ({ name: _bareToolName(c && c.name), args: c && c.args }))
                : (Array.isArray(turn._executedTools) ? turn._executedTools.map(n => ({ name: _bareToolName(n), args: null })) : []);
            const executed = new Set(executedCalls.map(c => c.name));
            const uncovered = [];
            for (const t of (Array.isArray(tools) ? tools : [])) {
                const name = _bareToolName(t && (t.name || t.originalName));
                if (!name) continue;
                const kws = Array.isArray(t.intentKeywords) ? t.intentKeywords : [];
                if (kws.length === 0) continue;
                const hitKws = kws.filter(k => k && items.some(it => it.includes(k)));
                if (hitKws.length === 0) continue;   // 用户压根没提这个工具族
                if (executed.has(name)) continue;    // 本题已经做过了
                // 【等价完成】已有工具在语义上覆盖了该意图（如 add_formula 写 =SUMIFS
                // 同时完成"写入"与"合计"），不再判缺项 —— 取代旧的关键词白名单豁免
                //（旧版按 hitKws 全员是否在白名单判定，"看看/分析"等宽泛词会让豁免失效，
                //  且把"分组"列为公式可满足，多维分组会被漏放）。
                if (_isCoveredByEquivalence(name, executedCalls, msg)) continue;
                uncovered.push({ name, displayName: t.displayName || name, keywords: hitKws });
            }
            return uncovered;
        } catch (_e) { return []; }
    }

    /**
     * 计算工具调用标签名（与 tool-registry 注册一致）：
     * MCP 本地工具 → 完整 originalName（excel-local__excel_read_range / ppt-local__pptx_generate）；
     * 其余 → 规范化名。
     * 不硬编码来源前缀，新工具集（如 PPT）自动适配。
     */
    function _toolCallTag(tool) {
        if (!tool) return '';
        const orig = String(tool.originalName || '');
        return (orig && orig.includes('__')) ? orig : tool.name;
    }

    /**
     * 构造"完整可拷贝"工具调用标签示例
     * 用于重试提示：把必填参数（尤其是工作表名参数）用真实活动表名填好，让 LLM 照抄就行
     * 减少小模型"自由发挥"编造工具名或漏参数的概率
     */
    function _buildToolCallExample(turn, tool) {
        if (!tool || !tool.name) return '';
        const prefix = '[' + _toolCallTag(tool) + ']';
        const schema = tool.inputSchema?.properties || {};
        const required = tool.inputSchema?.required || [];
        const activeSheet = String(turn?.context?.activeSheet?.sheet || '').trim();
        const activeWorkbook = String(turn?.context?.activeSheet?.workbook || '').trim();
        const lastUserText = String(turn?.userContent || turn?.message || '');

        const rawName = tool.name.includes('__') ? tool.name.split('__').pop() : tool.name;
        const args = {};

        // 【修复"只说不做" · 智能预填】自动从用户最后一条消息中提取实际参数，而不是输出占位符
        //   这样 LLM 复制示例标签就能直接运行，效果与用户话语精准对应
        const allKeys = Object.keys(schema);
        for (const key of allKeys) {
            const prop = schema[key] || {};
            const desc = String(prop.description || '');
            let val = null;

            // 工作表 / 工作簿 → 预填当前活动
            if (/工作表|sheet/.test(desc) && activeSheet) {
                val = activeSheet;
            } else if (/工作簿|workbook/.test(desc) && activeWorkbook) {
                val = activeWorkbook;
            }
            // 去重列 columns：从"根据AB列 / 按A列 / 按A、B列 / 按 部门、产品 列 / 根据第1列和第3列"里抽
            else if (key === 'columns' && rawName === 'excel_remove_duplicates') {
                const cols = [];
                // 先把全文中所有 "按...列 / 根据...列 / 基于...列" 的候选段取出来（避免匹配到句子中其他地方的字母）
                const colSegRe = /(?:按|根据|基于)\s*([\s\S]{0,60}?)\s*列/g;
                const segments = [];
                let mSeg;
                while ((mSeg = colSegRe.exec(lastUserText)) !== null) {
                    segments.push(mSeg[1]);
                }
                colSegRe.lastIndex = 0;

                for (const seg of segments) {
                    // (a) 最常见：列字母**连写**（根据AB列 / 按AB列）→ 每个 1~3 位连续字母视为一个列号
                    const runLetters = seg.match(/[A-Za-z]{1,3}/g) || [];
                    if (runLetters.length > 0) {
                        // 过滤：若候选里包含"部门/产品/姓名/日期"等中文表头词，应该跳过字母规则
                        const hasChineseHeader = /[\u4e00-\u9fa5]{2,}/.test(seg);
                        if (!hasChineseHeader) {
                            runLetters.forEach(x => cols.push(x.toUpperCase()));
                            continue;
                        }
                    }
                    // (b) 分隔写法（按A、B、C列 / 按A B列）
                    if (cols.length === 0) {
                        const sepLetters = seg.split(/\s*[、，,;；\s]\s*/).filter(s => /^[A-Za-z]{1,3}$/.test(s));
                        if (sepLetters.length > 0) {
                            sepLetters.forEach(x => cols.push(x.toUpperCase()));
                            continue;
                        }
                    }
                    // (c) 列序号（第 1、3 列 / 根据第1列和第2列）
                    if (cols.length === 0) {
                        const numRe = /(\d+(?:\s*[、，,;；和及\s]\s*\d+)*)/;
                        const mN = seg.match(numRe);
                        if (mN && mN[1]) {
                            const tokens = mN[1].split(/\s*[、，,;；和及\s]+\s*/).filter(x => /^\d+$/.test(x));
                            if (tokens.length > 0) {
                                tokens.forEach(n => cols.push(parseInt(n, 10)));
                                continue;
                            }
                        }
                    }
                    // (d) 中文表头词（按 客户、产品 列）→ 直接原样字符串数组即可（后端 ExecuteRemoveDuplicates 会用表头值匹配）
                    if (cols.length === 0) {
                        const tokens = seg.split(/\s*[、，,;；\s]+\s*/).filter(s => s && s.trim());
                        if (tokens.length >= 1) {
                            tokens.forEach(s => cols.push(s));
                        }
                    }
                }
                if (cols.length > 0) val = cols;
            }

            if (val !== null) args[key] = val;
        }

        // 必填：把剩下的 required 用占位符补齐
        for (const key of required) {
            if (key in args) continue;
            const prop = schema[key] || {};
            const desc = String(prop.description || '');
            if (/工作表|sheet/i.test(desc)) args[key] = activeSheet || '<当前活动工作表名>';
            else if (/工作簿|workbook/i.test(desc)) args[key] = activeWorkbook || '<当前工作簿名>';
            else args[key] = `<${key}的实际值>`;
        }

        const argKeys = Object.keys(args);
        // 没有任何参数 → 直接输出"无需参数"，LLM 可以按两种格式复制
        if (argKeys.length === 0) {
            return `  ${prefix}（无需参数）`;
        }
        return `  ${prefix}${JSON.stringify(args)}`;
    }

    /**
     * 通用N-gram分词：从文本中提取2~4字的片段用于匹配
     * 不依赖任何硬编码词表，纯算法
     */
    function extractNgrams(text) {
        if (!text || typeof text !== 'string') return [];
        const cleaned = text.replace(/[^\u4e00-\u9fa5a-zA-Z0-9]/g, '');
        const ngrams = new Set();
        for (let n = 2; n <= 4; n++) {
            for (let i = 0; i <= cleaned.length - n; i++) {
                ngrams.add(cleaned.slice(i, i + n));
            }
        }
        return Array.from(ngrams);
    }

    //  scoreTool 已移除。意图匹配统一走后端 McpCapabilityRegistry.Match()。
    // extractNgrams 保留，用于后端不可用时的纯前端兜底匹配。

    /**
     *  调用后端统一能力匹配 API，替代前端硬编码词表
     * 后端 McpCapabilityRegistry.Match() 会基于 DisplayName + Keywords + Description + Ngram 四层权重匹配
     * 返回格式与旧 findCandidateTools 兼容：{tool, score} 数组
     * @param {string} userMessage
     * @param {Object} toolRegistry
     * @param {number} topN
     * @returns {Promise<Array<{tool: Object, score: number}>>}
     */
    async function findCandidateTools(userMessage, toolRegistry, topN = 3, toolWhitelist = null) {
        if (!userMessage || typeof userMessage !== 'string') return [];
        if (!toolRegistry || typeof toolRegistry.listBySource !== 'function') return [];

        // 优先：调用后端统一匹配 API
        try {
            if (App.excelBridge && App.excelBridge.MatchCapability) {
                // WebView2 COM 不支持可选参数，后端已改为单参数（内部固定 topN=5）
                const resultJson = await App.excelBridge.MatchCapability(userMessage);
                const parsed = typeof resultJson === 'string' ? JSON.parse(resultJson) : resultJson;
                if (parsed && parsed.success && Array.isArray(parsed.matches)) {
                    // 能力 id → 真实工具：优先精确/后缀匹配，再做"同族优先"模糊匹配。
                    // 关键：excel-local / ppt-local 工具都以 source='mcp'、去前缀名注册，
                    // 所以 name.includes(id) 会跨族命中（如 id='chart' 命中 ppt_add_chart_slide），
                    // 导致"画柱状图"被误路由到 PPT（后端报 "PPT 项目不存在"）。
                    // 同族优先：id 无族前缀时排除 PPT 族，仅当无同族候选才回退异族。
                    const _allTools = () => [
                        ...toolRegistry.listBySource('mcp'),
                        ...toolRegistry.listBySource('skill')
                    ];
                    const _familyOf = (n) => (String(n || '').match(/^([a-z]+)_/) || [])[1] || '';
                    const _resolveCapabilityTool = (id) => {
                        const direct = toolRegistry.get(id);
                        if (direct) return direct;
                        const list = _allTools();
                        const exact = list.find(t => (t.name || '') === id);
                        if (exact) return exact;
                        const suffix = list.find(t => (t.name || '').endsWith('_' + id));
                        if (suffix) return suffix;
                        const idFam = _familyOf(id);
                        const sameFam = list.filter(t => idFam ? _familyOf(t.name) === idFam : _familyOf(t.name) !== 'ppt');
                        return sameFam.find(t => (t.name || '').includes(id))
                            || list.find(t => (t.name || '').includes(id))
                            || null;
                    };
                    const results = parsed.matches
                        .filter(m => m && m.id)
                        .map(m => {
                            const tool = _resolveCapabilityTool(m.id);
                            return {
                                tool: tool || { name: m.id, displayName: m.displayName, description: m.description },
                                score: m.score || 0,
                                confidence: m.confidence || 0,
                                matchedKeywords: m.matchedKeywords || []
                            };
                        });
                    // 后端 API 返回后也按白名单过滤
                    if (Array.isArray(toolWhitelist) && toolWhitelist.length > 0) {
                        const whitelistNames = new Set(toolWhitelist.map(t => t?.name || t));
                        return results.filter(r => whitelistNames.has(r.tool.name));
                    }
                    return results;
                }
            }
        } catch (e) {
            ChatLog.warn?.('AGENT', '[agent-loop] 后端 MatchCapability 不可用，降级到前端 ngram 兜底:', e.message);
        }

        // 兜底：纯前端 ngram 匹配（后端不可用时，同时搜索 MCP 工具和 Skill）
        const msg = userMessage.toLowerCase();
        let allTools = [
            ...toolRegistry.listBySource('mcp'),
            ...toolRegistry.listBySource('skill')
        ];
        // 如果调用方指定了工具白名单，只在白名单内搜索
        if (Array.isArray(toolWhitelist) && toolWhitelist.length > 0) {
            const whitelistNames = new Set(toolWhitelist.map(t => t?.name || t));
            allTools = allTools.filter(t => whitelistNames.has(t.name));
            if (allTools.length === 0) return [];
        }
        return allTools
            .map(tool => {
                let score = 0;
                const dn = (tool.displayName || '').toLowerCase();
                if (dn && msg.includes(dn)) score += 100;
                for (const ng of extractNgrams(dn)) {
                    if (ng.length >= 2 && msg.includes(ng.toLowerCase())) score += ng.length * 5;
                }
                // skill keywords 匹配（对齐后端四层权重的 Keywords 层）：
                // 纯 ngram 对"美化表格"匹配"智能表格美化"得分低（4分），加上 keywords 后
                // "美化表格"命中"美化表格"(4字) 大幅加分，前端兜底也能可靠识别 skill 意图。
                const skillInfo = tool.skillInfo;
                if (skillInfo) {
                    const kws = Array.isArray(skillInfo.keywords) ? skillInfo.keywords
                        : Array.isArray(skillInfo.Keywords) ? skillInfo.Keywords
                        : [];
                    for (const kw of kws) {
                        if (typeof kw !== 'string' || kw.length < 2) continue;
                        if (msg.includes(kw.toLowerCase())) {
                            score += kw.length >= 4 ? 15 : (kw.length >= 3 ? 8 : 3);
                        }
                    }
                }
                const desc = (tool.description || '').toLowerCase();
                for (const ng of extractNgrams(desc)) {
                    if (ng.length >= 2 && msg.includes(ng.toLowerCase())) score += ng.length;
                }
                const name = (tool.name || '').replace(/^.*__/, '').replace(/^(excel|ppt)_/, '');
                const parts = name.split('_');
                for (const part of parts) {
                    if (part.length >= 3 && msg.includes(part.toLowerCase())) score += 2;
                }
                return { tool, score };
            })
            .filter(t => t.score > 0)
            .sort((a, b) => b.score - a.score)
            .slice(0, topN);
    }

    /**
     * 从 MCP / Skill 工具结果中提取"给 LLM 看的纯文本摘要"，避免把内部嵌套 JSON / CSS 字典 /
     * McpContentItem 包装结构整个喂给 LLM（小模型经常会直接把这些 JSON 片段抄进用户回复，即"代码溢出"）
     */
    function extractToolSummaryForLLM(result) {
        if (result == null) return '';
        if (typeof result === 'string') return result;

        const collected = [];

        // 1. 标准 MCP: { content: [{type:"text", text:"..."}, ...] }
        if (Array.isArray(result.content)) {
            for (const item of result.content) {
                if (!item) continue;
                if (typeof item === 'string') { collected.push(item); continue; }
                const t = item.text ?? item.Text ?? item.content ?? item.Content;
                if (typeof t === 'string' && t.trim()) collected.push(t);
            }
        }
        // 兼容 VB 首字母大写: { Content: [ {Text:"..."} ] }
        if (Array.isArray(result.Content)) {
            for (const item of result.Content) {
                if (!item) continue;
                if (typeof item === 'string') { collected.push(item); continue; }
                const t = item.text ?? item.Text ?? item.content ?? item.Content;
                if (typeof t === 'string' && t.trim()) collected.push(t);
            }
        }

        // 2. Skill 结果: { message, Message, outputs:{content,message,...} }
        const msg = result.message || result.Message;
        if (typeof msg === 'string' && msg.trim() && !collected.includes(msg.trim())) collected.push(msg.trim());

        const outs = result.outputs || result.Outputs;
        if (outs && typeof outs === 'object') {
            const oc = outs.content || outs.Content || outs.message || outs.Message;
            if (typeof oc === 'string' && oc.trim() && !collected.includes(oc.trim())) collected.push(oc.trim());
        }

        // 3. 纯数组/表格 → 简短 shape 描述
        if (Array.isArray(result)) {
            const rows = result.length;
            const cols = (rows > 0 && result[0] && typeof result[0] === 'object') ? Object.keys(result[0]).length : 0;
            if (rows > 0 || cols > 0) {
                collected.push(`[数组数据 ${rows} 行 × ${cols} 列]`);
            }
        } else if (result && typeof result === 'object' && Array.isArray(result.rows) && Array.isArray(result.headers)) {
            collected.push(`[表格数据 ${result.rows.length} 行 × ${result.headers.length} 列]`);
        }

        if (collected.length === 0) {
            // 兜底：极简化 JSON，只保留顶层字符串字段，剥掉内部字典
            try {
                const flat = {};
                for (const k of Object.keys(result)) {
                    const v = result[k];
                    if (typeof v === 'string' && v.length < 500) flat[k] = v;
                    else if (typeof v === 'number' || typeof v === 'boolean') flat[k] = v;
                }
                const s = JSON.stringify(flat);
                return s.length > 400 ? s.slice(0, 400) + '\n...(已截断)' : s;
            } catch (e) {
                return '';
            }
        }

        return collected.join('\n');
    }

    /**
     * 截断工具结果防止上下文爆炸
     * 【关键修复】先 extractToolSummaryForLLM 剥掉内部包装 JSON/CSS 字典，再截断，
     *            从源头上避免小模型总结时把 JSON 片段直接抄进对用户的回复（"代码溢出"）。
     */
    /**
     * 【read_range 特化】解析完整工具结果，生成"统计摘要 + 前5行预览"。
     * 背景：executeLocalExcelTool 返回的 read_range 结果是 11KB+ 的完整 JSON，
     * 原 truncateResult 会截断到 MAX_TOOL_RESULT_CHARS=4000，模型只看到
     * "头部元数据 + previewRows 第一行"，于是只能回显第一行而非总结。
     * 此处改为：前端解析全量 rows，计算行数/数值列合计均值极值/分类列唯一值数，
     * 附上前 5 行样例，让模型有真实统计依据可总结，且 token 受控。
     * @param {string} resultStr - executeLocalExcelTool 返回的 JSON 字符串
     * @returns {string|null} 摘要文本；解析失败返回 null（走原截断逻辑）
     */
    /**
     * 从 read_range 解析结果里提取"数据行数组 + 列名 + 行数"。
     * 真实后端结构（ExcelMcpTools.Core.vb ExecuteReadRange）：
     *   output{ workbook, worksheet, range, rows:<行数·数字>, columns, headers:[...], previewRows:[...], result:[{列名:值}...] }
     * 兼容旧假设的嵌套写法 { result:{ rows:[], headers:[] } }。
     * @returns {{rows:Array, headers:Array, count:number}|null}
     */
    function extractReadRowsHeaders(parsed) {
        // 裸二维数组（模型/中间层只回传了数据体，如 [["日期",...],[...]]）
        if (Array.isArray(parsed)) {
            if (parsed.length === 0) return null;
            return { rows: parsed, headers: [], count: parsed.length };
        }
        if (!parsed || typeof parsed !== 'object') return null;
        // 前端先剥掉后端在截断时追加的 "/* ⚠️ 数据已截断 */" 尾注（合法 JSON 解析才会成功）
        if (Array.isArray(parsed.result) && Array.isArray(parsed.headers)) {
            const rows = parsed.result;
            return {
                rows,
                headers: parsed.headers,
                count: (typeof parsed.rows === 'number' && parsed.rows >= 0) ? parsed.rows : rows.length
            };
        }
        // { data:[[...]], headers:[...] }：excel_clean_data 及部分批量读取工具的返回形态
        if (Array.isArray(parsed.data) && parsed.data.length > 0) {
            return {
                rows: parsed.data,
                headers: Array.isArray(parsed.headers) ? parsed.headers : [],
                count: (typeof parsed.rows === 'number' && parsed.rows >= 0) ? parsed.rows : parsed.data.length
            };
        }
        // 兼容嵌套容器形态
        const inner = parsed.result && typeof parsed.result === 'object' ? parsed.result : parsed;
        if (Array.isArray(inner.rows) && Array.isArray(inner.headers)) {
            return {
                rows: inner.rows,
                headers: inner.headers,
                count: (typeof inner.rowsCount === 'number') ? inner.rowsCount : inner.rows.length
            };
        }
        return null;
    }

    // 数值判定：空串/空白不算 0（Number('') === 0 会把空列误统计成 0），支持千分位与百分号
    function toNumericValue(v) {
        if (typeof v === 'number') return Number.isFinite(v) ? v : NaN;
        if (typeof v !== 'string') return NaN;
        const s = v.trim().replace(/,/g, '');
        if (!s) return NaN;
        const body = s.endsWith('%') ? s.slice(0, -1) : s;
        if (!/^[+-]?(\d+(\.\d+)?|\.\d+)$/.test(body)) return NaN;
        const num = Number(body);
        return Number.isFinite(num) ? num : NaN;
    }

    // 预览用：单值裁剪，避免把整段长文本塞回上下文
    function clipCellValue(v) {
        if (v == null) return '';
        const s = String(v);
        return s.length > 30 ? s.slice(0, 30) + '…' : s;
    }

    /**
     * 把 read_range 的数据行结构格式化为"统计摘要 + 前3行预览"文本。
     * 供字符串入口与对象入口共用。
     * 支持两种行形态：对象行 [{列名:值}]（returnFormat=object）与二维数组行 [[值,...]]（默认 array）。
     * 【关键】预览必须是"列名=值"文本，绝不 JSON.stringify 原始数据行——
     * 否则弱模型（本地 7B）会把这坨可复制的数组直接当答案回显（实测："清洗这份数据"）。
     * @returns {string|null}
     */
    /**
     * 【仅展示意图】用户只想"看到"数据（展示/显示/查看 + 表格/json/数据/结果），
     * 且不含分析/聚合/计算语义。用于决定回喂模型的工具摘要是否要求"给结论"。
     * 与 prompt-builder.detectDisplayIntent 保持同一判定口径。
     * @param {string} msg
     * @returns {boolean}
     */
    function isDisplayOnlyMessage(msg) {
        if (!msg || typeof msg !== 'string') return false;
        if (!/(展示|显示|呈现|列出|查看|看看|输出|渲染|做成)/.test(msg)) return false;
        if (!/(表格|数据|结果|明细|列表|清单|行列|内容)/.test(msg)) return false;
        if (/(分析|总结|汇总|统计|对比|比较|排名|趋势|占比|比例|计算|评估|洞察|解读|结论|归纳|概括|按|各|分组|每个|分别|透视|聚合|同比|环比)/.test(msg)) return false;
        return true;
    }

    function formatReadSummary(rows, headers, count, displayOnly) {
        if (!rows || rows.length === 0) return null;
        const hd = Array.isArray(headers) ? headers : [];
        const rowsAreArrays = Array.isArray(rows[0]);
        // 数据体：若 headers 全为空白串（hasHeaders=false 的数组格式），退化为"列N"占位，统计才能按列号生效
        let dataRows = rows;
        let colCount = hd.length;
        for (const r of rows) {
            const len = rowsAreArrays ? (Array.isArray(r) ? r.length : 0) : Object.keys(r || {}).length;
            if (len > colCount) colCount = len;
        }
        const namedCount = hd.filter(h => h != null && String(h).trim() !== '').length;
        let firstRowNames = null;
        if (rowsAreArrays && namedCount === 0 && rows.length > 1) {
            // 首行是列名（后端 hasHeaders=false 时把表头当数据行返回）
            const first = rows[0];
            const looksLikeHeader = Array.isArray(first) && first.length > 0 &&
                first.every(v => v != null && String(v).trim() !== '' && !Number.isFinite(toNumericValue(v)));
            if (looksLikeHeader) {
                dataRows = rows.slice(1);
                colCount = Math.max(colCount, first.length);
                firstRowNames = first.map(v => String(v).trim());
            }
        }
        if (colCount === 0 || dataRows.length === 0) return null;
        const cols = [];
        for (let i = 0; i < colCount; i++) {
            const fromHeader = (i < hd.length && hd[i] != null) ? String(hd[i]).trim() : '';
            const fromFirstRow = (firstRowNames && i < firstRowNames.length) ? firstRowNames[i] : '';
            cols.push(fromHeader !== '' ? fromHeader : (fromFirstRow !== '' ? fromFirstRow : ('列' + (i + 1))));
        }
        const cellAt = (row, i) => rowsAreArrays
            ? (Array.isArray(row) ? row[i] : undefined)
            : (row ? row[cols[i]] : undefined);

        // 逐列：数值列出合计/均值/极值，文本列出唯一值数
        const numericLines = [];
        const catLines = [];
        for (let i = 0; i < colCount; i++) {
            let sum = 0, max = -Infinity, min = Infinity, n = 0;
            const uniq = new Set();
            for (const r of dataRows) {
                const v = cellAt(r, i);
                const num = toNumericValue(v);
                if (Number.isFinite(num)) {
                    sum += num;
                    if (num > max) max = num;
                    if (num < min) min = num;
                    n++;
                } else if (v != null && String(v).trim() !== '') {
                    uniq.add(String(v));
                }
            }
            if (n > 0) {
                numericLines.push('- ' + cols[i] + ': 合计=' + (Math.round(sum * 100) / 100) + ', 平均=' + (Math.round((sum / n) * 100) / 100) + ', 最大=' + max + ', 最小=' + min);
            } else if (uniq.size >= 2 && uniq.size <= 50) {
                catLines.push('- ' + cols[i] + ': ' + uniq.size + ' 种（' + Array.from(uniq).slice(0, 5).join('、') + '）');
            }
        }

        const lines = [];
        lines.push('【选区数据】共 ' + count + ' 行 × ' + cols.length + ' 列，列名：' + cols.join('、'));
        if (numericLines.length > 0) {
            lines.push('【数值统计】');
            lines.push(numericLines.join('\n'));
        }
        if (catLines.length > 0) {
            lines.push('【分类维度（唯一值数）】');
            lines.push(catLines.join('\n'));
        }
        const previewRows = dataRows.slice(0, 3);
        if (previewRows.length > 0) {
            lines.push('【数据预览（前3行，仅供理解字段含义）】');
            for (const r of previewRows) {
                lines.push('  ' + cols.map((c, i) => c + '=' + clipCellValue(cellAt(r, i))).join(' | '));
            }
        }
        if (displayOnly) {
            // 【仅展示意图】用户只想看到数据，界面已把工具结果渲染成表格/JSON 展示。
            // 若此处仍写"请给出结论"，模型必然强行输出一段分析
            // （实测投诉："似乎每次都必须给个数据分析结果"）。
            lines.push('以上是本次读取到的真实数据，界面已将其渲染为表格直接展示给用户。你只需回复"已展示"或一句极简说明即可；严禁输出任何分析、总结、结论、趋势解读或指标计算（如合计/均值/最大/最小/占比/排名），也不要复述或粘贴任何数据行。');
        } else {
            lines.push('请基于以上统计给出结论：禁止以 JSON、数组或表格形式回显原始数据，禁止复制任何数据行，用自然语言输出。');
        }
        return lines.join('\n');
    }

    function buildReadRangeSummary(resultStr, displayOnly) {
        try {
            const clean = String(resultStr).replace(/\s*\/\*[\s\S]*$/, '').trim();
            const parsed = JSON.parse(clean);
            if (!parsed || typeof parsed !== 'object') return null;
            const shape = extractReadRowsHeaders(parsed);
            if (!shape) return null;
            return formatReadSummary(shape.rows, shape.headers, shape.count, displayOnly);
        } catch (e) {
            return null;
        }
    }

    // 兼容 result.result 已是解析对象（非字符串）的情况
    function buildReadRangeSummaryFromObject(obj, displayOnly) {
        if (!obj || typeof obj !== 'object') return null;
        try {
            const shape = extractReadRowsHeaders(obj);
            if (!shape) return null;
            return formatReadSummary(shape.rows, shape.headers, shape.count, displayOnly);
        } catch (e) {
            return null;
        }
    }

    /**
     * excel_clean_data 结果摘要：清洗结果里 data 字段带全量数据数组，原文可达 15KB+，
     * 截断后模型只剩"抄数组"（实测：'清洗这份数据' 的回复是一坨 [["日期","地区",...]]）。
     * 这里压成"统计 + 前 3 行预览"，完整数据仍留在工具卡/表格卡片里给用户看。
     * 与 skill 管道的 buildCleanOpSummary 对齐，但走通用管道时也生效。
     * @returns {string|null} 命中清洗结果形态时返回摘要，否则 null
     */
    function buildCleanResultSummary(resultStr) {
        try {
            const clean = String(resultStr).replace(/\s*\/\*[\s\S]*$/, '').trim();
            const parsed = JSON.parse(clean);
            if (!parsed || typeof parsed !== 'object' || parsed.cleaned !== true) return null;
            const data = Array.isArray(parsed.data) ? parsed.data : null;
            const headers = Array.isArray(parsed.headers) ? parsed.headers : (data && Array.isArray(data[0]) ? data[0] : []);
            const headOffset = parsed.hasHeader === false ? 0 : 1;
            const rowCount = data ? Math.max(0, data.length - headOffset) : (parsed.cleanedRows || 0);
            const ops = [];
            if (parsed.trimmedCells) ops.push(`去首尾空格 ${parsed.trimmedCells} 处`);
            if (parsed.normalizedDates) ops.push(`日期标准化 ${parsed.normalizedDates} 处`);
            if (parsed.normalizedNumbers) ops.push(`数字标准化 ${parsed.normalizedNumbers} 处`);
            if (parsed.removedEmptyRows) ops.push(`删除空行 ${parsed.removedEmptyRows} 行`);
            if (parsed.removedRows) ops.push(`删除重复行 ${parsed.removedRows} 行`);
            const lines = [];
            lines.push(`清洗已完成：工作表『${parsed.worksheet || ''}』${parsed.range || ''}，共 ${rowCount} 行 × ${headers.length} 列`);
            lines.push(ops.length ? `实际改动：${ops.join('、')}` : '实际改动：无（数据已符合规范）');
            if (parsed.writtenTo) lines.push(`已写回工作表：${parsed.writtenTo}`);
            if (parsed.note) lines.push(String(parsed.note));
            const preview = data ? data.slice(headOffset, headOffset + 3) : [];
            if (preview.length > 0) {
                lines.push(`前 ${preview.length} 行预览：`);
                for (const r of preview) {
                    lines.push('  ' + (Array.isArray(r)
                        ? r.map((v, i) => `${headers[i] || ('列' + (i + 1))}=${v == null ? '' : v}`).join(' | ')
                        : String(r)));
                }
            }
            lines.push('（完整清洗结果已在上方表格卡片中；回复里请直接给文字结论，不要粘贴原始数组）');
            return lines.join('\n');
        } catch (e) {
            return null;
        }
    }

    /**
     * 【新建工作表名回喂】写入/落地类工具的结果里带工作表名（writtenTo/worksheet/targetSheet）。
     * 摘要若已含表名则不重复；否则补一行"工作表：X"，让 LLM 下一轮能引用真实表名
     * （替代旧 MultiStepEngine.extractStepContext 的上下文传递）。
     * @param {string} toolName
     * @param {*} result - 工具原始返回（对象或字符串）
     * @returns {string} 需要补的表名行；无需补充时返回空串
     */
    /**
     * 【幻觉工具名自纠】找出与（不存在的）工具名最接近的真实工具名。
     * 背景：弱模型会编造不存在的工具名（实测：excel_set_pivot_table / excel_set_pivot_range /
     * excel_move_pivot_table / excel_chart_slice / excel_set_chart_property 连环 10+ 次失败），
     * 旧报错只说"未注册，请改用可用工具"，模型无从纠正 → 继续编下一个名字烧光轮次。
     * 这里按"下划线词元重叠 + 公共前缀"打分，把最接近的真实工具名直接喂回给模型。
     * @param {Object} registry - toolRegistry（需有 list()）
     * @param {string} hallucinated - 模型编造的工具名
     * @param {number} [limit=3] - 最多返回几个
     * @returns {string[]}
     */
    function findClosestToolNames(registry, hallucinated, limit) {
        limit = limit || 3;
        try {
            const hRaw = String(hallucinated || '').toLowerCase();
            if (!hRaw) return [];
            const names = (registry && typeof registry.list === 'function')
                ? registry.list().map(t => t && t.name).filter(Boolean) : [];
            const hTokens = hRaw.split('_').filter(s => s.length >= 3);
            const scored = [];
            for (const n of names) {
                const nl = String(n).toLowerCase();
                if (nl === hRaw) continue;
                let score = 0;
                for (const tok of hTokens) {
                    if (tok.length >= 4 && nl.indexOf(tok) >= 0) score += 2;
                    else if (nl.indexOf(tok) >= 0) score += 1;
                }
                // 公共前缀（字符级）：excel_set_chart_property vs excel_set_chart_properties 前缀极长
                let p = 0;
                const minLen = Math.min(nl.length, hRaw.length);
                while (p < minLen && nl[p] === hRaw[p]) p++;
                score += Math.min(p, 12) * 0.1;
                if (score > 0) scored.push({ name: n, score: score });
            }
            scored.sort((a, b) => b.score - a.score);
            return scored.slice(0, limit).map(s => s.name);
        } catch (e) {
            return [];
        }
    }

    function buildSheetNameHint(toolName, result) {
        if (!toolName || !/write|create|split|merge|clean|dedup|aggregat|copy|group|sheet/i.test(toolName)) return '';
        if (result == null) return '';
        const text = typeof result === 'string' ? result : (() => { try { return JSON.stringify(result); } catch (e) { return ''; } })();
        if (!text) return '';
        const m = text.match(/"(?:writtenTo|worksheet|targetSheet|sheetName|sheet)"\s*:\s*"([^"\\]{1,60})"/i);
        if (!m) return '';
        const sheet = m[1].trim();
        if (!sheet) return '';
        return `工作表：${sheet}`;
    }

    // 图表闸门 / 补链：实现见 agent-loop/chart-gate.js（须先于本文件加载）
    function _chartGate() {
        const g = App.ChatV2 && App.ChatV2.agentLoopChartGate;
        if (!g) throw new Error('agentLoopChartGate not loaded — check script order');
        return g;
    }
    function inferChartType(userMsg, override) { return _chartGate().inferChartType(userMsg, override); }
    function isChartIntent(trn, rawMsg) { return _chartGate().isChartIntent(trn, rawMsg); }
    function resolveAggregateSource(trn) { return _chartGate().resolveAggregateSource(trn); }
    function buildChartNudgeHint(trn, rawMsg, chartTypeOverride) {
        return _chartGate().buildChartNudgeHint(trn, rawMsg, chartTypeOverride);
    }

    /** 工具参数摘要（显示在 thought 块标题里，最多 60 字符） */
    function summarizeToolParamsForThought(toolName, args) {
        if (!args || typeof args !== 'object') return '';
        const pickMap = {
            'excel_read_range': ['sheet', 'range'],
            'excel_aggregate': ['range', 'dimColumn', 'valueColumn', 'func', 'targetSheet', 'groupBy'],
            'excel_create_chart': ['chartType', 'chartEnum', 'dataRange', 'sheet', 'title'],
            'excel_create_chart_bar': ['chartType', 'chartEnum', 'dataRange', 'sheet', 'title'],
            'excel_write_range': ['sheet', 'range'],
            'excel_clean_data': ['sheet', 'range', 'cleanType'],
            'excel_split_workbook': ['sheetName', 'splitBy'],
            'excel_merge_workbooks': ['sheetName'],
            'excel_dedup_range': ['sheet', 'range', 'column']
        };
        const keys = pickMap[toolName] || Object.keys(args).slice(0, 4);
        const parts = [];
        for (const k of keys) {
            let v = args[k];
            if (v == null) continue;
            if (typeof v === 'string') {
                if (v.length > 30) v = v.slice(0, 28) + '...';
            } else if (typeof v === 'object') {
                try { v = JSON.stringify(v).slice(0, 20); } catch { continue; }
            }
            parts.push(`${k}=${v}`);
        }
        let s = parts.join(', ');
        if (s.length > 55) s = s.slice(0, 52) + '...';
        return s;
    }

    /**
     * 从 excel_aggregate 的结果里稳健提取 labels（分类名）/ values（数值）。
     * 后端 data 可能是对象数组、二维数组，也可能只在 Artifact 里带结构化 labels/values；
     * 三者任一命中即可，避免"labels 为空 → 内联图表被跳过 → 聊天里看不到可视化图"。
     * @param {Object} meta - 解析后的工具返回对象（含 data / dimColumn / valueColumn / metrics / artifact）
     * @param {Object} [args] - 工具调用参数（含 dimColumn / valueColumn，用于对象数组取键）
     * @returns {{labels: string[], values: number[]}}
     */
    function extractAggregateLabelsValues(meta, args) {
        const out = { labels: [], values: [] };
        const m = (meta && typeof meta === 'object') ? meta : {};
        const a = (args && typeof args === 'object') ? args : {};

        // ① 后端结构化 Artifact（ExcelMcpTools.Advanced.Aggregate.vb 已补 labels/values）
        const art = (m.artifact && typeof m.artifact === 'object') ? m.artifact : null;
        if (art && Array.isArray(art.labels) && art.labels.length) {
            out.labels = art.labels.map(v => v == null ? '' : String(v));
            out.values = Array.isArray(art.values) ? art.values.map(v => Number(v) || 0) : [];
        }

        // ② meta.data：对象数组 [{维度列:值, 数值列:值}] / ③ 二维数组 [["部门A",100]]
        const rows = Array.isArray(m.data) ? m.data : [];
        if (!out.labels.length && rows.length) {
            const dimKey = String(m.dimColumn || a.dimColumn || '');
            const valKey = String(m.valueColumn || a.valueColumn ||
                (Array.isArray(m.metrics) && m.metrics[0] && m.metrics[0].name) || '');
            const labels = [];
            const values = [];
            for (const r of rows) {
                if (Array.isArray(r)) {
                    labels.push(r.length ? String(r[0] == null ? '' : r[0]) : '');
                    values.push(r.length > 1 ? (Number(r[1]) || 0) : 0);
                    continue;
                }
                if (r && typeof r === 'object') {
                    const keys = Object.keys(r);
                    const dk = (dimKey && r[dimKey] !== undefined) ? dimKey : keys[0];
                    const vk = (valKey && r[valKey] !== undefined) ? valKey
                        : (keys.length > 1 ? keys[1] : keys[0]);
                    labels.push(r[dk] == null ? '' : String(r[dk]));
                    values.push(Number(r[vk]) || 0);
                    continue;
                }
                labels.push('');
                values.push(0);
            }
            out.labels = labels;
            out.values = values;
        }

        // 过滤空标签并对齐两列长度
        const pairs = out.labels.map((l, i) => [String(l || '').trim(), Number(out.values[i]) || 0])
            .filter(p => p[0] !== '');
        return { labels: pairs.map(p => p[0]), values: pairs.map(p => p[1]) };
    }

    function detachVisionImages(result) {
        const payload = result && result.result;
        if (!payload || !Array.isArray(payload.visionImages) || !payload.visionImages.length) return [];
        const images = payload.visionImages;
        delete payload.visionImages;
        return images;
    }

    function pushVisionPages(messages, images) {
        if (!messages || !images || !images.length) return;
        const content = [{ type: 'text', text: '下面是图纸页面（最长边不超过 1600 像素）。请直接读取图中的尺寸、公差和标题栏，只写图上有的内容。' }];
        images.forEach(function (img) {
            const b64 = img && (img.base64 || img.data);
            if (!b64) return;
            const mime = (img && img.mime) || 'image/png';
            content.push({ type: 'image_url', image_url: { url: 'data:' + mime + ';base64,' + b64 } });
        });
        if (content.length > 1) messages.push({ role: 'user', content: content });
    }

    function truncateResult(result, toolName, displayOnly) {
        if (result == null) return '';
        // 【excel_merge_sheets 特化】合并结果回喂前强制压成确定性摘要。
        // 根因（2026-10-09 用户反馈）：合并结果含 sourceSheets/mergedRows/data 矩阵，
        // 原样回喂后模型会"复述"成无逗号裸 JSON 塞进正文。把完整 payload 换成
        // 友好摘要，模型既看不到完整 JSON 也无从复读，从源头杜绝该问题。
        // 与下方 excel_clean_data / read_range 特化并列，走同一"压摘要回喂"范式。
        if (toolName === 'excel_merge_sheets') {
            const _mergeStr = typeof result === 'string' ? result : JSON.stringify(result);
            try {
                const _friendly = App.ChatV2?.toolRegistryUtil?.replaceJsonEchoWithFriendly?.(_mergeStr);
                if (_friendly) {
                    // 摘要脱去渲染层 HTML（与 read_range/clean_data 特化一致：回喂 LLM 的是纯文本，
                    // 而非带 <br>/<span> 的 HTML，避免模型把标签也学进复述）
                    const _plain = String(_friendly).replace(/<[^>]+>/g, '').replace(/✅/g, '').trim();
                    if (_plain) {
                        ChatLog.info?.('TOOLS', '[agent-loop] merge_sheets 结果已压成确定性纯文本摘要回喂模型');
                        return _plain;
                    }
                }
            } catch (e) { /* 守卫失败回落通用截断，不影响主路径 */ }
        }
        // 【excel_clean_data 特化】清洗结果内含全量 data 数组，必须先压成统计摘要，
        // 否则模型拿到被截断的裸 JSON 只会照抄数组（实测："清洗这份数据"）。
        if (typeof result === 'string') {
            const cleanSummary = buildCleanResultSummary(result);
            if (cleanSummary) {
                ChatLog.info?.('TOOLS', `[agent-loop] clean_data 结果已压成统计摘要（${cleanSummary.length} chars）回喂模型`);
                return cleanSummary;
            }
        }
        // 【read_range 特化】字符串或已解析对象都尝试生成统计摘要，避免大数据量下弱模型被裸 JSON
        // 淹没后只回显首行。命中则返回摘要，并打日志便于排查（未命中说明形态不匹配，走通用截断）。
        const readSummary = typeof result === 'string'
            ? buildReadRangeSummary(result, displayOnly)
            : (typeof result === 'object' ? buildReadRangeSummaryFromObject(result, displayOnly) : null);
        if (readSummary) {
            ChatLog.info?.('TOOLS', `[agent-loop] read_range 数据已聚合为统计摘要（${readSummary.length} chars）回喂模型`);
            return readSummary;
        }
        if (typeof result === 'string') {
            ChatLog.warn?.('TOOLS', `[agent-loop] read_range 摘要未命中（strlen=${result.length}），走通用截断；原文前 120 字：${result.slice(0, 120)}`);
        }
        const summary = extractToolSummaryForLLM(result);
        const str = (typeof summary === 'string' && summary.length > 0) ? summary
            : (typeof result === 'string' ? result : JSON.stringify(result));
        const body = str.length <= MAX_TOOL_RESULT_CHARS
            ? str
            : str.slice(0, MAX_TOOL_RESULT_CHARS) + '\n...(结果已截断)';
        // 表名回喂：摘要未含表名时补一行，保证后续步骤能引用新建的工作表
        const hint = buildSheetNameHint(toolName, result);
        return (hint && body.indexOf(hint) < 0) ? (hint + '\n' + body) : body;
    }

    // 【空结果防幻觉】读取类工具「执行成功但没有任何内容」时的识别与拦截。
    // 实测故障：ppt_get_slide_text 因 slideIndex 缺省 0 抛异常返回空串，弱模型拿到空结果后
    // 不承认"没读到"，反而凭页面标题编出一份看起来合理的正文
    // （"项目进度更新 / 项目A: 已完成30%，预计10月完成…"），用户看到的就是"假内容"。
    // 只靠模型自觉不够——框架必须把"空"这件事显式写进回喂内容。
    const READ_TOOL_NAME_RE = /(?:^|_)(get|read|list|fetch|query|info|count|select|view|extract)(?:_|$)/i;

    /** 状态字段不算"内容"，只有它们时视为空结果 */
    const _STATUS_ONLY_KEYS_RE = /^(success|ok|taskCompleted|method|durationMs|elapsed|tool|toolName|host|error|message)$/i;

    /** @returns {boolean} 对象是否为"只有状态字段、没有任何内容"的壳 */
    function _isShellObjectEmpty(obj) {
        if (obj == null) return true;
        if (Array.isArray(obj)) return obj.length === 0;
        if (typeof obj !== 'object') return false;
        const keys = Object.keys(obj);
        if (keys.length === 0) return true;
        for (const k of keys) {
            // error/message 非空说明是失败信息，不算"空结果"（交由失败分支处理）
            if (/^(error|message)$/i.test(k) && obj[k] != null && String(obj[k]).trim() !== '') return false;
            if (_STATUS_ONLY_KEYS_RE.test(k)) continue;
            const v = obj[k];
            if (v == null) continue;
            if (typeof v === 'string' && v.trim() === '') continue;
            if (Array.isArray(v) && v.length === 0) continue;
            if (typeof v === 'object' && Object.keys(v).length === 0) continue;
            return false;
        }
        return true;
    }

    /**
     * 判断工具返回是否等价于"空"（空串 / 空数组 / 仅含 success 等状态字段的壳 JSON）。
     * @param {*} result
     * @returns {boolean}
     */
    function isEffectivelyEmptyToolResult(result) {
        if (result == null) return true;
        if (typeof result === 'string') {
            const s = result.trim();
            if (!s) return true;
            const head = s[0];
            if (head !== '{' && head !== '[') return false;   // 纯文本结果（如 "0"、"2026-10-04"）不算空
            let obj;
            try { obj = JSON.parse(s); } catch (e) { return false; }
            return _isShellObjectEmpty(obj);
        }
        if (typeof result === 'object') return _isShellObjectEmpty(result);
        return false;
    }

    /**
     * 解析 LLM 的 tool_calls 事件为标准格式
     * 与文本路径对齐：原生 tool_calls 路径同样校验工具名是否已注册。
     * 模型（尤其本地小模型）可能直接在原生 tool_calls 里编造不存在的工具名，
     * 必须在执行前拦截并触发强引导重试，而不是等到 invokeTracked 运行时才失败。
     * @param {Array} toolCallsRaw - LLM 返回的原始 tool_calls
     * @param {Object} [toolRegistry] - 工具注册表（缺省则不校验，保持向后兼容）
     * @returns {{valid: Array, invalid: Array}}
     */
    function parseToolCalls(toolCallsRaw, toolRegistry) {
        if (!toolCallsRaw) return { valid: [], invalid: [] };
        if (Array.isArray(toolCallsRaw)) {
            const valid = [];
            const invalid = [];
            toolCallsRaw.forEach((tc, idx) => {
                let name = tc.function?.name || tc.name || '';
                name = aliasToolName(name);
                const args = tc.function?.arguments
                    ? (typeof tc.function.arguments === 'string'
                        ? safeParseJSON(tc.function.arguments)
                        : tc.function.arguments)
                    : (tc.args || tc.arguments || {});
                // 【幻觉加固】空名 / 未注册工具 → 拦截并标记 invalid
                // 之前 name='' 时 `name && !toolRegistry.get(name)` 短路不触发 → 空名直接进 valid
                if (!name || (toolRegistry && !toolRegistry.get(name))) {
                    invalid.push({ rawName: name, attemptedName: name || '(空名)' });
                    return;
                }
                valid.push({
                    id: tc.id || `call_${Date.now()}_${idx}`,
                    name,
                    args
                });
            });
            return { valid, invalid };
        }
        return { valid: [], invalid: [] };
    }

    function safeParseJSON(str) {
        try { return JSON.parse(str); } catch (e) { return null; }
    }

    /**
     * 稳定序列化：对象键排序后序列化，使 {a:1,b:2} 与 {b:2,a:1} 得到同一字符串。
     * 用于"同参数"判定（工具调用去重 / 重复执行拦截）。
     */
    function stableStringify(v) {
        if (v === null || typeof v !== 'object') {
            try { return JSON.stringify(v); } catch (e) { return String(v); }
        }
        if (Array.isArray(v)) return '[' + v.map(stableStringify).join(',') + ']';
        const keys = Object.keys(v).sort();
        return '{' + keys.map(k => JSON.stringify(k) + ':' + stableStringify(v[k])).join(',') + '}';
    }

    /**
     * 计算工具调用的"幂等键"（工具名 + 规范化参数）。
     * 参数值会被截断到合理长度，避免超大 rows 参与比较时开销过大。
     */
    function toolCallKey(tc) {
        if (!tc || !tc.name) return '';
        let argsStr;
        try { argsStr = stableStringify(tc.args || {}); } catch (e) { argsStr = String(tc.args); }
        return tc.name + '|' + argsStr;
    }

    /**
     * 【通用框架·同轮工具调用去重】
     * LLM（尤其 chat-only 小模型）经常在同一轮把同一个工具调用重复输出多次：
     *   - 原生 tool_calls 里重复；
     *   - 正文里把同一段 [tool]{args} / JSON 复读多遍。
     * 不去重会导致同一个重型工具（如 ppt_generate）被执行 N 次：
     * UI 连弹 N 张相同卡片、重复生成文件、"已完成"被复读多遍。
     * 这里按"工具名 + 规范化参数"去重，只保留首次出现。
     * @param {Array} calls - 工具调用数组
     * @returns {Array} 去重后的数组
     */
    function dedupeToolCalls(calls) {
        if (!Array.isArray(calls) || calls.length <= 1) return calls || [];
        const seen = new Set();
        const out = [];
        let dropped = 0;
        for (const tc of calls) {
            if (!tc) continue;
            const key = toolCallKey(tc);
            if (key && seen.has(key)) { dropped++; continue; }
            if (key) seen.add(key);
            out.push(tc);
        }
        if (dropped > 0) {
            ChatLog.warn?.('AGENT', `[agent-loop] 🔁 工具调用去重：同轮丢弃 ${dropped} 个重复调用（${calls.length} → ${out.length}）`);
        }
        return out;
    }

    /**
     * 【通用框架·终态工具】生成类工具：一旦成功即产出最终文件，
     * 属"任务终点"。此类工具成功后应及时收尾，避免模型在后续轮次反复重复调用/复读结论。
     * 通过工具元数据（isTerminal / terminal）或内置名表识别，新工具可在 ppt_tools.json/excel_tools.json
     * 中声明 "isTerminal": true 自动纳入，无需改动本表。
     */
    // 终态 = 产出最终交付物后应收尾的工具。
    // 注意：excel_create_chart 故意不在此表——多图/看板需连续建图；单图收尾由图表闸门/_chartDone 负责。
    // 真实工具名是 excel_export_csv（无 excel_export_to_csv / excel_export_to_pdf）。
    const BUILTIN_TERMINAL_TOOLS = new Set([
        'ppt_generate',
        'ppt_project_build',
        'excel_export_csv',
        'excel_split_workbook',
        // 按列拆分是重型写操作：同 turn 内参数一致的成功调用须幂等拦截，避免叠出 销售一部_2
        'excel_split_by_column',
        'excel_split_by_rows'
    ]);

    function isTerminalTool(toolName, toolRegistry) {
        if (!toolName) return false;
        const t = toolRegistry && typeof toolRegistry.get === 'function' ? toolRegistry.get(toolName) : null;
        if (t && (t.isTerminal === true || t.IsTerminal === true || (t.meta && t.meta.isTerminal === true))) return true;
        return BUILTIN_TERMINAL_TOOLS.has(toolName);
    }

    // ── 【通用框架·按 turn 隔离的幂等台账】─────────────────────────────
    // 单轮内的幂等台账 turn._toolCallLedger 只能拦"同一 run() 内复读"；
    // 这里额外提供跨轮幂等：同一 turn 内的多轮（多步骤复合任务）若重复调用同一终态工具，
    // 直接复用上次结果、跳过执行，避免重复生成文件、连弹相同卡片。
    // 关键修复：原实现为"会话级"单例（Map<idemKey, ...>），从不在 turn 边界清空，
    // 导致同一会话内用户两条**无关**消息若参数相同会被静默复用、跳过执行（误挡真实重做）。
    // 现改为按 turnId 分桶（Map<turnId, Map<idemKey, ...>>）：跨 turn 天然隔离，
    // 新消息/新 run 用新 turnId，不会命中历史台账，仅"同一 turn 内的多轮"才复用。
    // 安全护栏：
    //   - 仅对终态/重型工具（isTerminalTool）生效，读类工具可放心重复；
    //   - TTL 15min，过期自动失效（避免长期误挡）；
    //   - 容量上限 200（跨所有 turn 的 entry 总数），超量近似 FIFO 淘汰；
    //   - 用户消息含"重新/再来/重新生成/regenerate"等再生意图时绕过，放行真正的重做。
    const _SESSION_TOOL_LEDGER = new Map(); // turnId -> Map<idemKey, { ok, result, at }>
    const SESSION_LEDGER_TTL_MS = 15 * 60 * 1000;
    const SESSION_LEDGER_MAX = 200;
    const REPROBE_RE = /重新|再来|重做|重画|再画|换个?\s*(版|份|个|次|张)?\s*(ppt|图表|报告|文件)|重新生成|再来一份|重新做|regenerate|redo|again|make (a|another)/i;

    // 按 turnId 取桶：跨 turn 天然隔离，仅"同一 turn 内的多轮"复用结果。
    function _sessionLedgerBucket(turnId) {
        let bucket = _SESSION_TOOL_LEDGER.get(turnId);
        if (!bucket) { bucket = new Map(); _SESSION_TOOL_LEDGER.set(turnId, bucket); }
        return bucket;
    }
    function sessionLedgerGet(turnId, key) {
        const bucket = _SESSION_TOOL_LEDGER.get(turnId);
        if (!bucket) return null;
        const e = bucket.get(key);
        if (!e) return null;
        if (Date.now() - e.at > SESSION_LEDGER_TTL_MS) { bucket.delete(key); return null; }
        return e;
    }
    function sessionLedgerSet(turnId, key, val) {
        const bucket = _sessionLedgerBucket(turnId);
        // 总容量保护：跨所有 turn 的 entry 数超限时，淘汰首个非空 bucket 的首个 key（近似 FIFO）
        let total = 0;
        _SESSION_TOOL_LEDGER.forEach(function (b) { total += b.size; });
        if (total >= SESSION_LEDGER_MAX && !bucket.has(key)) {
            for (const b of _SESSION_TOOL_LEDGER.values()) {
                if (b.size > 0) { const fk = b.keys().next().value; if (fk !== undefined) b.delete(fk); break; }
            }
        }
        bucket.set(key, val);
    }
    function resetSessionToolLedger() { _SESSION_TOOL_LEDGER.clear(); }
    function latestUserMsgText(messages) {
        if (!Array.isArray(messages)) return '';
        for (let i = messages.length - 1; i >= 0; i--) {
            const m = messages[i];
            if (m && m.role === 'user') {
                const c = m.content;
                return typeof c === 'string' ? c : (c && typeof c === 'object' ? JSON.stringify(c) : '');
            }
        }
        return '';
    }

    /**
     * 解析文本形式的工具调用
     * 支持格式：
     *   [excel-local__tool_name]{"arg":"value"}
     *   [tool_name]{"arg":"value"}
     *   [tool_name]
     * 先匹配已注册工具名；再检测伪造/无效的工具调用格式（如[excel-local__不存在的工具名]），防止模型编造工具和假结果
     */
    function parseTextToolCalls(text, toolRegistry, turn) {
        if (!text || typeof text !== 'string' || !toolRegistry) return { valid: [], invalid: [] };

        const registeredNames = toolRegistry.names();
        if (registeredNames.length === 0) return { valid: [], invalid: [] };

        // 收集原始名和规范化名，支持 [excel-local__xxx] 和 [xxx] 两种前缀
        // 额外支持模型常省略的 excel_ 前缀，如 [excel-local__clear_sheet] → excel_clear_sheet
        // 同时支持连字符变体 [excel-create-chart] → excel_create_chart
        // 动态收集本地 MCP 工具的来源前缀（excel-local__ / ppt-local__ 等）：
        // 从已注册工具的 originalName 推导，新工具集自动纳入，不硬编码
        const localPrefixes = new Set(['excel-local__', 'excel-local_', 'word-local__', 'ppt-local__']);
        for (const n of registeredNames) {
            const t = typeof toolRegistry.get === 'function' ? toolRegistry.get(n) : null;
            const orig = String(t?.originalName || '');
            const sep = orig.indexOf('__');
            if (sep > 0) {
                localPrefixes.add(orig.slice(0, sep + 2));
                localPrefixes.add(orig.slice(0, sep + 1));
            }
        }
        const localPrefixList = Array.from(localPrefixes);

        const nameLookup = new Map();
        for (const name of registeredNames) {
            // 小模型常把 <source>-local__（双下划线）错写成 <source>-local_（单下划线），
            // 两种前缀变体都注册到 nameLookup，保证真实工具无论哪种拼法都能被解析执行。
            nameLookup.set(name, name);
            for (const p of localPrefixList) nameLookup.set(p + name, name);
            // 连字符变体：excel_create_chart → excel-create-chart
            const hyphenated = name.replace(/_/g, '-');
            nameLookup.set(hyphenated, name);
            for (const p of localPrefixList) nameLookup.set(p + hyphenated, name);
            // 短名变体：小模型常省略工具名前缀（excel_/ppt_/word_ 等）
            const knownPrefixes = ['excel_', 'ppt_', 'word_'];
            for (const kp of knownPrefixes) {
                if (name.startsWith(kp)) {
                    const shortName = name.slice(kp.length);
                    for (const p of localPrefixList) nameLookup.set(p + shortName, name);
                    nameLookup.set(shortName, name);
                    const shortHyphenated = shortName.replace(/_/g, '-');
                    for (const p of localPrefixList) nameLookup.set(p + shortHyphenated, name);
                    nameLookup.set(shortHyphenated, name);
                }
            }
        }
        // 域别名兜底：把泛化/简写工具名归一为真实工具（如 chart → excel_create_chart），
        // 与 TOOL_NAME_ALIASES 保持一致，避免后端"工具名翻译"误译到跨族工具（如 ppt_add_chart_slide）。
        for (const [alias, real] of Object.entries(TOOL_NAME_ALIASES)) {
            if (!nameLookup.has(alias)) nameLookup.set(alias, real);
        }
        const escapedNames = Array.from(nameLookup.keys()).map(n => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
        // 小模型常把"["... "]" 与"{"... "}"顺序写反（[name{json}] 而非 [name]{json}），
        // 标准正则 \[name\](?:\{...\})? 匹配不到，导致"明明有调用意图"却被当成"未调用工具"反复重试。
        // 改为先匹配"整段 [name...]"（name 之后任意非 ] 字符），再从整段里抠出 args 段。
        const pattern = new RegExp('\\[(' + escapedNames.join('|') + ')([^\\]]*)\\]', 'g');

        // 从冒号分隔的 key:value 字符串解析参数
        // 例如：chart_type: "pie" data_source: "$E$2:$E$74" include_legend: true title: "按部门展示利润"
        function parseColonKV(text) {
            const args = {};
            // 匹配 key: value 对，value 可以是双引号字符串、单引号字符串、裸值（true/false/数字/null/变量）
            // 用贪婪分割保证 key 名里的下划线/连字符都能被识别
            const kvRe = /([a-zA-Z_][a-zA-Z0-9_]*)\s*:\s*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)'|([^\s\[\],][^\s\[\],]*)?)/g;
            let m;
            while ((m = kvRe.exec(text)) !== null) {
                const key = m[1];
                let val;
                if (m[2] !== undefined) val = m[2];
                else if (m[3] !== undefined) val = m[3];
                else if (m[4] !== undefined) {
                    const raw = m[4].trim();
                    if (raw === 'true') val = true;
                    else if (raw === 'false') val = false;
                    else if (raw === 'null' || raw === 'undefined') val = null;
                    else if (/^-?\d+(\.\d+)?$/.test(raw)) val = Number(raw);
                    else val = raw;
                }
                // 跳过冒号前的英文介词/冗余词（只在 value 未识别时兜底）
                if (!val) val = true;
                args[key] = val;
            }
            return args;
        }

        /**
         * 【修复·方括号外参数丢失】从闭括号 "]" 之后的文本中提取参数。
         * 模型常把参数写在括号外：
         *   [excel_format_range] {"range":"A1:C3","bold":true}
         *   [excel_format_range] range:A1:C3 bold:true
         * 旧逻辑只扫 name 与 ] 之间的 tail（此时为空），args 被整段丢弃 → 工具以空参执行/失败。
         * 这里补抓 ] 之后、同行（或下一个 "[" 之前）的参数段：JSON 对象 或 key:value/key=value 序列。
         * 仅当文本确实像参数才返回，绝不吞掉后续自然语言（ prose 不含 "标识符:值" 形态）。
         * @param {string} afterText - 紧接闭括号 "]" 之后的文本
         * @returns {Object|null} 解析出的参数对象；非参数形态返回 null
         */
        function extractTrailingParams(afterText) {
            if (!afterText) return null;
            // 1) JSON 对象：{ ... }（非贪婪，遇到第一个 } 即止，不会越界吞掉下一个工具调用）
            const jsonM = afterText.match(/^\s*(\{[\s\S]*?\})/);
            if (jsonM) {
                const parsed = safeParseJSON(jsonM[1]);
                if (parsed && typeof parsed === 'object' && !Array.isArray(parsed)) return parsed;
            }
            // 2) key:value / key=value 序列：限同行、遇换行或下一个 "[" 即止
            const line = afterText.split('\n')[0].split('[')[0].trim();
            if (!line) return null;
            if (!/[:=]/.test(line)) return null; // 没有键值分隔符 → 不是参数，放过（避免吞自然语言）
            const kv = parseColonKV(line);
            if (Object.keys(kv).length > 0) return kv;
            return null;
        }

        const validCalls = [];
        let match;
        while ((match = pattern.exec(text)) !== null) {
            const rawName = match[1];
            const tail = match[2] || '';  // name 之后、] 之前的所有字符，例如 "" 或 "{\"source\":\"Sheet1\"}"
            const normalizedName = nameLookup.get(rawName) || rawName;
            // 抠 args：优先从 tail 抽 {…}，再处理 [name]{json}（tail=""）→ 跟原解析逻辑保持一致
            let args = {};
            const braceStart = tail.indexOf('{');
            if (braceStart >= 0) {
                const braceEnd = tail.lastIndexOf('}');
                if (braceEnd > braceStart) {
                    const argsStr = tail.substring(braceStart, braceEnd + 1);
                    const parsed = safeParseJSON(argsStr);
                    if (parsed && typeof parsed === 'object') {
                        args = parsed;
                    }
                    // 解析失败 → 继续 fallback 到 colonKV（下面的 if 分支）
                } else {
                    // 有 { 没有 }（半花括号）→ 清除残留，fallback 到 colonKV
                    // 不把不完整 JSON 泄漏到渲染层
                }
            }
            // 没有有效 JSON 时，尝试解析冒号分隔的 key:value 格式
            // 例如 LLM 输出 [excel_create_chart chart_type: "pie" data_source: "$A$1:$B$10" title: "按地区"]
            if (Object.keys(args).length === 0 && tail.trim().length > 0) {
                args = parseColonKV(tail);
            }
            // 【修复·方括号外参数】tail 内仍无参数时，再补抓闭括号 "]" 之后写在外侧的参数段
            // （[tool] {json} 或 [tool] key:val）。不补抓的话这些参数会被整段丢弃。
            if (Object.keys(args).length === 0) {
                const after = text.slice(match.index + match[0].length);
                const trailing = extractTrailingParams(after);
                if (trailing) args = trailing;
            }
            validCalls.push({
                id: `textcall_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                name: normalizedName,
                args
            });
        }

        // 【假调用转真调用 2026-10-09"只说不做"根治】模型有时不用约定的 [tool]{args} 方括号格式，
        // 而是用函数调用样式文本冒充已执行："excel_split_by_column(column="部门", sheet="", namingPattern="{value}") → 成功: ✅ 已按列拆分…"
        // （实测 23:05：turn 显示"未调用工具"，正文却是伪造的调用+成功结果，用户诉求完全没执行）。
        // 与其拦截重试（弱模型可能反复犯），不如顺势而为：名字在注册表里 + 紧邻"成功/完成"声明
        // （确认这是"冒充已执行"而非普通提及）→ 解析括号内 k=v 参数 → 当真实调用执行。
        // 防误伤：普通散文提及（无成功声明）不转换；同名工具已由方括号解析捕获时不重复执行。
        {
            const fnCallRe = /\b((?:excel|ppt|word|computer)_[a-z][a-z0-9_]{2,40})\s*\(([^()]{0,400})\)([\s\S]{0,90})/g;
            let fnMatch;
            while ((fnMatch = fnCallRe.exec(text)) !== null) {
                const fnName = nameLookup.get(fnMatch[1]);
                if (!fnName) continue; // 未注册名 → 交给下方伪造检测/未注册提示
                if (validCalls.some(c => c.name === fnName)) continue; // 同名已捕获，防重复执行
                // 必须紧邻"冒充已执行"声明（箭头/成功/完成/✅），普通散文提及不转换
                if (!/(?:→|->|=>|成功|完成|✅|done)/i.test(fnMatch[3])) continue;
                const argStr = fnMatch[2] || '';
                let fnArgs = {};
                const bm = argStr.match(/\{[\s\S]*\}/);
                if (bm) { const p = safeParseJSON(bm[0]); if (p && typeof p === 'object' && !Array.isArray(p)) fnArgs = p; }
                if (Object.keys(fnArgs).length === 0 && /[:=]/.test(argStr)) {
                    const kvRe = /([a-zA-Z_][a-zA-Z0-9_]*)\s*[:=]\s*("([^"\\]|\\.)*"|'[^']*'|[^,，;；\s)][^,，;；]*)/g;
                    let kvm;
                    while ((kvm = kvRe.exec(argStr)) !== null) {
                        let v = kvm[2];
                        if (v && /^(".*"|'.*')$/.test(v)) v = v.slice(1, -1);
                        else if (v === 'true') v = true;
                        else if (v === 'false') v = false;
                        else if (v && /^-?\d+(\.\d+)?$/.test(v)) v = Number(v);
                        fnArgs[kvm[1]] = v;
                    }
                }
                validCalls.push({
                    id: `fncall_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                    name: fnName,
                    args: fnArgs
                });
                ChatLog.info?.('TOOLS', `[agent-loop] 函数样式假调用已转真实调用: ${fnName} args=${JSON.stringify(fnArgs)}`);
            }
        }

        // 【重要防御】检测伪造的工具调用：匹配 [<source>-local__xxx] / [excel_xxx] / [ppt_xxx] / 颠倒格式 [name{json}]
        // 但 xxx 不在注册列表中。本地小模型经常会编造工具名并输出假的"已完成"结果，必须拦截。
        // 前缀集合动态收集（excel-local__/ppt-local__ 等），新工具集自动纳入，不硬编码。
        const invalidCalls = [];
        const fakePrefixList = Array.from(localPrefixes).sort((a, b) => b.length - a.length);
        const fakePrefixRe = fakePrefixList.map(p => p.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|');
        const fakeToolPattern = new RegExp('\\[(' + fakePrefixRe + '[a-zA-Z0-9_-]+|excel[-_][a-zA-Z0-9_-]+|ppt[-_][a-zA-Z0-9_-]+)\\s*(?:\\{[\\s\\S]*?\\})?\\s*(?:\\]|\\}\\])', 'g');
        const matchedFakeSpans = [];
        while ((match = fakeToolPattern.exec(text)) !== null) {
            const rawName = match[1];
            // 提取实际工具名（去掉 <source>-local__ / <source>-local_ 等前缀，长前缀优先）
            let toolName = rawName;
            for (const p of fakePrefixList) {
                if (rawName.startsWith(p)) {
                    toolName = rawName.slice(p.length);
                    break;
                }
            }
            // 兼容 [excel_xxx] / [ppt_xxx] 假工具（去掉 excel_/ppt_ 前缀）
            const barePrefix = toolName.match(/^(excel|ppt)[-_]/);
            if (barePrefix) toolName = toolName.slice(barePrefix[0].length);
            // 检查是否是已知工具（nameLookup里已经包含了所有变体，所以直接查原始rawName即可）
            if (!nameLookup.has(rawName)) {
                invalidCalls.push({
                    rawName: rawName,
                    attemptedName: toolName
                });
                matchedFakeSpans.push({ start: match.index, end: match.index + match[0].length });
            }
        }

        // 兜底：处理 AI 违规输出的 [EXECUTE:SkillName] 标记
        // 虽然系统提示词禁止，但本地模型经常不听话，这里自动转成 skill 调用
        const executePattern = /\[EXECUTE:([a-zA-Z0-9_\-.]+)\]/g;
        while ((match = executePattern.exec(text)) !== null) {
            const skillName = match[1];
            const skillToolName = 'skill_' + skillName.replace(/[.\s]/g, '_');
            if (registeredNames.includes(skillToolName)) {
                validCalls.push({
                    id: `textcall_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                    name: skillToolName,
                    args: { prompt: turn?.userMessage || '', text: turn?.userMessage || '' }
                });
            }
        }

        return { valid: validCalls, invalid: invalidCalls };
    }

    /**
     * 从 assistantContent 中移除已解析的文本工具调用标记和伪造的工具调用标记
     * @param {Object} [toolRegistry] - 工具注册表（缺省时跳过来源前缀推导，仍做基础 strip）
     */
    function stripTextToolCalls(text, toolCalls, invalidCalls, toolRegistry) {
        if (!text || typeof text !== 'string') return text;
        let cleaned = text;
        const allCalls = [...(toolCalls || [])];
        // 把invalidCalls也加入待移除列表，构造伪tc对象用于strip
        (invalidCalls || []).forEach(ic => {
            allCalls.push({ name: ic.attemptedName, rawName: ic.rawName });
        });
        for (const tc of allCalls) {
            const rawName = tc.name;
            // 动态推导该工具的来源前缀（excel-local / ppt-local），新工具集自动适配
            const tObj = (toolRegistry && typeof toolRegistry.get === 'function') ? toolRegistry.get(rawName) : null;
            const orig = String(tObj?.originalName || '');
            const sep = orig.indexOf('__');
            const localBase = sep > 0 ? orig.slice(0, sep) : '';
            const namesToStrip = [rawName];
            if (localBase) {
                namesToStrip.push(localBase + '__' + rawName);
                namesToStrip.push(localBase + '_' + rawName);
            }
            // 连字符变体：excel_create_chart → excel-create-chart
            const hyphenated = rawName.replace(/_/g, '-');
            namesToStrip.push(hyphenated);
            if (localBase) {
                namesToStrip.push(localBase + '__' + hyphenated);
                namesToStrip.push(localBase + '_' + hyphenated);
            }
            if (tc.rawName) namesToStrip.push(tc.rawName);
            if (rawName.startsWith('excel_') || rawName.startsWith('ppt_') || rawName.startsWith('word_')) {
                const kp = rawName.startsWith('ppt_') ? 'ppt_' : (rawName.startsWith('word_') ? 'word_' : 'excel_');
                const shortName = rawName.slice(kp.length);
                namesToStrip.push(shortName);
                if (localBase) {
                    namesToStrip.push(localBase + '__' + shortName);
                    namesToStrip.push(localBase + '_' + shortName);
                }
                const shortHyphenated = shortName.replace(/_/g, '-');
                namesToStrip.push(shortHyphenated);
                if (localBase) {
                    namesToStrip.push(localBase + '__' + shortHyphenated);
                    namesToStrip.push(localBase + '_' + shortHyphenated);
                }
            }
            for (const name of namesToStrip) {
                const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                // 同时移除：[name]{json}、[name{json}]、[name key:val key:val...]、[name]
                // 以及参数写在 ] 外侧时泄漏的尾随 JSON： [name] {"range":"A1:C3"}（见上方 extractTrailingParams）
                const regex = new RegExp('\\[' + escaped + '(?:\\{[\\s\\S]*?\\}(?:\\])|\\](?:\\{[\\s\\S]*?\\})?|[^\\]]*\\])(?:\\s*\\{[\\s\\S]*?\\})?', 'g');
                cleaned = cleaned.replace(regex, '');
            }
            // 同时移除 [EXECUTE:SkillName] 标记
            if (rawName.startsWith('skill_')) {
                const skillDisplayName = rawName.replace(/^skill_/, '');
                const executeRegex = new RegExp('\\[EXECUTE:' + skillDisplayName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\]', 'g');
                cleaned = cleaned.replace(executeRegex, '');
            }
        }
        // 清理工具标记残留的说明文字（如"（无需参数）"、"（需要参数: xxx）"等）
        cleaned = cleaned
            .replace(/[（(]\s*无需参数\s*[）)]/g, '')
            .replace(/[（(]\s*需要参数[^）)]*[）)]/g, '')
            .replace(/[（(]\s*必填参数[^）)]*[）)]/g, '')
            .replace(/[（(]\s*参数[^）)]*[）)]/g, '');
        // 小模型可能用 ```markdown\n  excel_create_chart(...) \n```
        // 格式输出工具调用，这段代码块应被清理掉，避免作为自然语言文本渲染
        for (const tc of allCalls) {
            const rawName = tc.name;
            const markdownRe = new RegExp(
                '```[\\w]*\\s*' + rawName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
                '\\([\\s\\S]*?\\)\\s*```', 'gi'
            );
            cleaned = cleaned.replace(markdownRe, '');
            // 无代码围栏的裸 markdown 调用：excel_create_chart(...)
            const bareRe = new RegExp(
                rawName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') +
                '\\([^)]*\\)', 'gi'
            );
            cleaned = cleaned.replace(bareRe, '');
        }
        // ── 清理残留的"JSON 工具调用"代码块 ──
        // 模型有时不写 [tool]{args} 标签，而把工具调用写成 fenced JSON 代码块：
        //   ```json
        //   {"name":"ppt_generate","arguments":{...}}
        //   ```
        // 上面的标记清理只处理 [name]{args} / 裸 markdown 调用，没清理 fenced JSON，
        // 导致这类块原样渲染成"JSON 卡片"（截图实测：ppt_generate 参数块被重复渲染多次）。
        // 这里移除内容形如工具调用表示的 fenced 代码块：name 为已注册工具，或含 arguments/args/tool_calls 键。
        try {
            const fenceRe = /```[a-zA-Z]*\s*\n([\s\S]*?)```/g;
            let fm;
            const toRemove = [];
            const regNames = (toolRegistry && typeof toolRegistry.names === 'function') ? toolRegistry.names() : [];
            const regNameSet = new Set(regNames);
            while ((fm = fenceRe.exec(cleaned)) !== null) {
                const block = fm[1].trim();
                let parsed = null;
                try { parsed = JSON.parse(block); } catch (e) { parsed = null; }
                if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) continue;
                const candidateName = parsed.name || parsed.tool || (typeof parsed.function === 'object' ? parsed.function.name : null);
                const nameIsTool = !!candidateName && regNameSet.has(candidateName);
                const hasArgs = ('arguments' in parsed && typeof parsed.arguments === 'object')
                    || ('args' in parsed && typeof parsed.args === 'object')
                    || ('tool_calls' in parsed && Array.isArray(parsed.tool_calls));
                const isGenerateCall = /ppt_generate|ppt_project_build/i.test(String(candidateName || ''));
                if (nameIsTool || hasArgs || isGenerateCall) {
                    toRemove.push([fm.index, fm.index + fm[0].length]);
                }
            }
            // 从后往前删除，避免索引偏移
            for (let i = toRemove.length - 1; i >= 0; i--) {
                const [s, e] = toRemove[i];
                cleaned = cleaned.slice(0, s) + cleaned.slice(e);
            }
            if (toRemove.length > 0) {
                cleaned = cleaned.replace(/\n{3,}/g, '\n\n').trim();
                ChatLog.info?.('AGENT', `[agent-loop] 清理 ${toRemove.length} 个残留 JSON 工具调用代码块`);
            }
        } catch (e) { /* 无关紧要，保留原文本 */ }
        return cleaned.replace(/\n{3,}/g, '\n\n').trim();
    }

    /**
     * 主入口：执行 agent loop
     * @param {Object} params
     *   - turn: 当前 turn 对象
     *   - messages: 初始 messages 数组
     *   - tools: 可用工具列表（toolRegistry.list() 的结果）
     *   - signal: AbortSignal
     *   - maxIterations: 最大循环轮数（默认 MAX_TOOL_ROUNDS；Skill agentConfig.maxIterations 可覆盖）
     *   - tokenGuard: token 水位护栏 { maxTokens, warnRatio, criticalRatio }（默认 8192 / 0.7 / 0.9）
     *   - keepLatestRound: 危险水位本地折叠时保留的最近轮数（默认 2）
     *   - persistTools: 是否每轮都下发 tools（默认 false）。true 时循环不再"执行一次工具即
     *                   锁死为总结轮"，而由 LLM 每轮自主决定继续调工具或收尾，循环动力来自 LLM
     *                   决策；安全护栏（轮次上限 / token 水位 / abort）不变。
     * @returns {AsyncGenerator} typed events
     */
    async function* run({ turn, messages, tools, signal, maxIterations, tokenGuard, keepLatestRound, chartTypeOverride, silentContent, paramOverrides, persistTools }) {
        const llmClient = App.ChatV2.llmClient;
        const toolRegistry = App.ChatV2.toolRegistry;
        const sessionStore = App.ChatV2.sessionStore;
        const ContentExtractor = App.ChatV2.ContentExtractor;

        // 【仅展示意图】用户只想看到数据（"用表格/json 展示"）→ 回喂工具结果时不要求"给结论"，
        // 从根上避免模型每轮都追加一段分析；含分析/聚合词的请求不受影响。
        const _displayOnlyIntent = isDisplayOnlyMessage((turn && (turn._originalUserMessage || turn.userMessage)) || '');

        // agentConfig.maxIterations 覆盖默认循环上限；
        // keepLatestRound 用于危险水位本地折叠的保留轮数。
        const effectiveMaxRounds = Number(maxIterations) > 0 ? Number(maxIterations) : MAX_TOOL_ROUNDS;
        const effectiveKeepLatest = Number(keepLatestRound) > 0 ? Number(keepLatestRound) : 2;
        // persistTools=true：每轮都下发 tools，让 LLM 自主决定"继续调工具 or 收尾总结"，
        // 循环动力交由 LLM 决策；默认 false 保持单轮工具 + 收尾总结的既有行为。
        const keepToolsEveryRound = persistTools === true;

        // 【可观测性】入口留痕。这一行能一次性回答"这轮到底有没有走 agent-loop、
        // 手上多少工具、工具有没有带 intentKeywords（覆盖度自检的数据源）"——
        // 上一次排查卡住正是因为 JS 侧 ChatLog 只写 console（WebView2 看不到），
        // 生产环境完全无法区分"守卫没走到"和"走到了但判定为空"。
        _hostLog('INFO', `agent-loop 启动: persistTools=${keepToolsEveryRound} 轮次上限=${effectiveMaxRounds}`
            + ` 工具=${Array.isArray(tools) ? tools.length : 0}`
            + `(带intentKeywords=${(Array.isArray(tools) ? tools : [])
                .filter(t => Array.isArray(t && t.intentKeywords) && t.intentKeywords.length > 0).length})`
            // 【可见工具清单】排查"模型没调工具"时最需要知道的就是"该工具到底在不在手里"。
            // 例如格式化族被 filterToolsForMessage 瘦身剔除后，模型只能干说"我将格式化"。
            + ` 工具名=[${(Array.isArray(tools) ? tools : []).map(t => t && t.name).filter(Boolean).slice(0, 60).join(',')}]`
            + ` 用户=${String((turn && (turn.userMessage || turn.message)) || '').slice(0, 50)}`);

        if (!llmClient) {
            yield { type: 'error', error: { code: 'NO_LLM_CLIENT', message: 'llmClient not loaded', recoverable: false } };
            return;
        }

        if (!Array.isArray(messages) || messages.length === 0) {
            yield { type: 'error', error: { code: 'INVALID_MESSAGES', message: 'messages 为空或格式错误', recoverable: false } };
            return;
        }

        // turn 关闭时广播 finish 步骤日志
        const finishStepLog = (reason) => {
            try {
                sessionStore?.appendStepLog?.(turnId, 'finish', `Agent 循环结束：${reason || 'completed'}`, round + 1);
            } catch (e) { /* 忽略 */ }
        };

        // 注册运行状态 + 广播 turn/start
        const turnId = turn?.id || `turn_${Date.now()}`;
        let closeReason = '';
        let round = 0;  // 提升到 try 外，finally 收尾需要读取
        _runStates.set(turnId, { phase: 'idle', step: 0, status: 'idle' });
        emitAgent('turn/start', { turnId, kind: turn?.kind || '', message: turn?.userMessage || turn?.message || '' });
        setRunStatus(turnId, 'running', { phase: 'turn', step: 0 });

        try {
        let currentMessages = [...messages];
        try {
            const goalLine = App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.goalPrompt === 'function'
                ? App.ChatV2.turnGoal.goalPrompt(turn && turn.goal)
                : '';
            if (goalLine) currentMessages.push({ role: 'system', content: goalLine, _goal: true });
        } catch (_goalInj) { /* 目标注入失败不影响循环 */ }
        let forceRetryCount = 0;
        let coverageNudgeCount = 0;  // 【要求覆盖度自检】"多条要求只做了一条"的续做引导次数
        let chartNudgeCount = 0;     // 【图表交付闸门】"要出图但只写了文字"的补图引导次数
        let fileNudgeCount = 0;      // 【文件产物交付闸门】"要产物文件却没执行导出工具"的补做引导次数
        let pendingUncovered = [];   // 【如实收尾】最近一次自检里"用户要求、但还没执行"的工具

        /**
         * 【如实收尾】把"没做完"写进正文。
         * 续做引导只是"再劝一次"，模型完全可以不听；但框架不能替它把没做的说成做完了——
         * 用户投诉的核心正是"只做了 1 条，UI 却显示任务完成"。
         */
        function withCoverageNote(content) {
            if (!Array.isArray(pendingUncovered) || pendingUncovered.length === 0) return content;
            const miss = pendingUncovered.map(u => `${u.displayName || u.name}(${u.name})`).join('、');
            _hostLog('WARN', `收尾时仍有 ${pendingUncovered.length} 项未完成（${pendingUncovered.map(u => u.name).join(',')}）→ 正文追加未完成说明`);
            return `${content || ''}\n\n> ⚠️ 你要求的以下内容本次【没有执行】：${miss}。请让我继续，我不会把没做的说成已完成。`;
        }
        let toolsExecuted = false;  // 标记是否已成功执行过工具调用
        let anyToolFailed = false;  // 本轮是否有任何工具失败 —— 失败时 toolsExecuted 不标记，保留下轮重试机会
        let terminalToolDone = false;  // 终态工具（生成类）已成功执行 → 不再下发 tools，强制收尾
        // 重试期间内容只在本地累积，不推给 sessionStore，UI 不刷新
        // 等最终结果出来再一次性写入，避免用户看到消息"闪三遍"
        let pendingContent = '';
        let pendingReasoning = '';
        let isRetryRound = false;
        // 多步骤通道（persistTools=true：runMultiStepPipeline / skill-pipeline 等）下：
        // 把每轮 LLM 的叙述（计划/过渡词/思路）当作"思考"写进本轮 thinking 块，而非正文；
        // 正文只留给最终总结轮。这样"计划的字"归入思考/进度栏，填空"（无推理文字）"，
        // 也不再实时刷在正文里被"先答后做治理"清掉。
        const routeToThinking = !!persistTools;

        const _executedSignatures = new Set();   // 跨轮冗余防护：已成功执行的"工具名+参数"签名集合

        // ===== 关注点拆离：从主循环抽出的纯函数（便于单测、消除块级作用域耦合）=====
        // ① 本轮工具执行进度追踪：把成功工具名/参数/artifact 累加到 turn，并把"成功执行"的
        //    调用签名写入 _executedSignatures（供 ② 的跨轮重复判定使用）。
        /** 图表补链：委托 agent-loop/chart-gate.js */
        async function* runChartBackfill(p) {
            return yield* _chartGate().runChartBackfill(p);
        }


        function recordToolExecutions(trn, calls, results, rnd) {
            if (!trn) return;
            if (!Array.isArray(trn._executedTools)) trn._executedTools = [];
            if (!Array.isArray(trn._executedToolCalls)) trn._executedToolCalls = [];
            if (!Array.isArray(trn._artifacts)) trn._artifacts = [];
            for (let i = 0; i < calls.length; i++) {
                const tc = calls[i];
                const r = results?.[i];
                const succeeded = r && (r.envelope ? r.envelope.ok === true : !!r.ok);
                if (tc && tc.name) {
                    trn._executedTools.push(tc.name);
                    // 【等价完成的数据源】同时记录参数（覆盖度自检要用 formula 实参
                    // 判定"写 =SUMIFS 公式 = 已完成写入+聚合"），_executedTools 只存名字不够用
                    trn._executedToolCalls.push({
                        name: tc.name,
                        args: (tc.args && typeof tc.args === 'object') ? tc.args : null
                    });
                    // 跨轮冗余防护：只把信封判定为成功的调用记入签名集合
                    if (succeeded) _executedSignatures.add(_toolSig(tc.name, tc.args));
                }
                if (r && r.envelope) {
                    if (!Array.isArray(trn._envelopes)) trn._envelopes = [];
                    trn._envelopes.push(r.envelope);
                    if (succeeded && r.envelope.context && r.envelope.context.artifact
                        && !(r.result && r.result.artifact)) {
                        const artIdx = trn._artifacts.length;
                        trn._artifacts.push(Object.assign({}, r.envelope.context.artifact, {
                            _id: `art_${Date.now()}_${artIdx}`,
                            _toolName: tc?.name || '',
                            _round: rnd + 1,
                            _consumed: false
                        }));
                    }
                }
                // 从 result 提取 artifact（每个 tc 对应一个 result，按顺序）
                if (r && r.ok && r.result && r.result.artifact && typeof r.result.artifact === 'object') {
                    const artIdx = trn._artifacts.length;
                    const art = { ...r.result.artifact, _id: `art_${Date.now()}_${artIdx}`, _toolName: tc?.name || '', _round: rnd + 1, _consumed: false };
                    trn._artifacts.push(art);
                    ChatLog.info?.('AGENT', `[agent-loop] artifact 捕获 [${art._id}]: ${art.type || '?'} ${art.chartName || art.targetSheet || ''}`);
                    // 【血缘边补全】chart artifact 捕获时，匹配 pendingLineage 里对应的项
                    if (art.type === 'chart' && Array.isArray(trn._pendingLineage) && trn._pendingLineage.length > 0) {
                        const pending = trn._pendingLineage.find(p => p.toChartCallId === tc?.id || p.toTool === tc?.name);
                        if (pending) {
                            if (!Array.isArray(trn._lineage)) trn._lineage = [];
                            trn._lineage.push({
                                fromId: pending.fromId,
                                toId: art._id,
                                fromType: pending.fromType,
                                toType: 'chart',
                                fromSheet: pending.fromSheet,
                                fromDataRange: pending.fromDataRange,
                                toChartName: art.chartName,
                                toAnchor: art.anchor,
                                round: rnd + 1
                            });
                            ChatLog.info?.('AGENT', `[lineage] 血缘边补全: ${pending.fromId} → ${art._id} (${pending.fromSheet} → chart@${art.anchor || '?'})`);
                            // 清除 pending
                            const idx = trn._pendingLineage.indexOf(pending);
                            if (idx >= 0) trn._pendingLineage.splice(idx, 1);
                        }
                    }
                }
            }
        }

        // ② 跨轮重复判定：本轮每一个工具调用是否都已在"本轮之前"（beforeSigs）成功执行过，
        //    是则判定模型在死循环式重刷同一操作，直接收尾。beforeSigs 须在本轮写入签名前快照传入。
        function detectRedundantRepeat(calls, beforeSigs) {
            const sigs = calls.filter(tc => tc && tc.name).map(tc => _toolSig(tc.name, tc.args));
            return sigs.length > 0 && sigs.every(s => beforeSigs.has(s));
        }

        while (round <= effectiveMaxRounds) {
            // 本轮工具失败标记重置（每轮重新判定）
            anyToolFailed = false;

            // 检查取消
            if (signal?.aborted) {
                closeReason = 'user_aborted';
                yield { type: 'canceled', reason: 'user_aborted' };
                return;
            }

            // step/start
            const stepNo = round + 1;
            const stepTotal = effectiveMaxRounds + 1;
            setRunStatus(turnId, 'running', { phase: 'step', step: stepNo, stepTotal });
            emitAgent('step/start', { turnId, step: stepNo, maxSteps: stepTotal });

            // 【迭代替换模型】每轮 LLM 推理前创建一个 thinking thought 块
            // 标题用中性描述，不暴露轮次号——agent-loop 是 LLM 自主循环的，轮次是运行时动态决定的，
            // 不是预先规划好的固定步骤，显示轮次号会让用户误以为是硬编码的"第几步"。
            // 捕获本轮 thinking 块引用：多步骤通道下轮内叙述会流式写进它，收尾时再决定是否移回正文
            let currentThought = null;
            if (sessionStore && typeof sessionStore.startThought === 'function') {
                currentThought = sessionStore.startThought(turn.id, stepNo, '正在分析你的需求…', { kind: 'thinking' });
            }
            setRunStatus(turnId, 'running', { phase: 'thinking', step: stepNo, stepTotal });

            yield { type: 'round_start', round: stepNo, maxRounds: effectiveMaxRounds };

            // ── 每轮 LLM 调用前检查 ──
            // 对齐 multi-step-agent-mcp：token_estimate → 水位判断 → 可选压缩。
            if (round > 0) {  // 首轮消息通常是刚构造的，不需要检查
                const wm = await checkTokenWatermark({ turnId, sessionStore, messages: currentMessages, tools, tokenGuard, iteration: stepNo });
                // 【P2 修复】本地折叠条件扩展：压缩未生效（compactNow 失败/无变化）时，
                // warn 级水位同样回退本地折叠，避免上下文超窗导致下一轮 LLM 调用硬失败。
                // critical 级保持既有无条件折叠行为不变。
                const needLocalFold = currentMessages.length > 4 && (
                    wm.level === 'critical' ||
                    (wm.level === 'warn' && wm.compacted === false)
                );
                if (needLocalFold) {
                    currentMessages = _localCompressMessages(currentMessages, effectiveKeepLatest);
                    ChatLog.warn?.('AGENT', `[agent-loop] 本地折叠(${wm.level}${wm.compacted === false ? '/压缩未生效' : ''}): ${currentMessages.length} 条消息剩余`);
                }
                if (wm.level !== 'none') {
                    yield { type: 'token_watermark', level: wm.level, total: wm.total, windowMax: wm.windowMax, ratio: wm.ratio, compacted: wm.compacted };
                }
            }

            // ── 每轮开始记录 thought ──
            const stepThought = `第 ${stepNo}/${effectiveMaxRounds + 1} 轮：${
                keepToolsEveryRound
                    ? (toolsExecuted ? 'LLM 决定是否继续调用工具或收尾' : 'LLM 推理并选择工具')
                    : (toolsExecuted ? '工具已执行，等待模型总结' : 'LLM 推理并选择工具')
            }`;
            try { sessionStore?.appendStepLog?.(turnId, 'thought', stepThought, stepNo); } catch (e) { /* 忽略 */ }

            // agent/pre-step：改写 / 拒绝 / 注入（参考 DSH waterfall）
            const preStepResult = await runPreStepHooks({ turnId, turn, messages: currentMessages, step: stepNo, signal });
            if (preStepResult.action === 'reject') {
                closeReason = preStepResult.reason || 'pre_step_rejected';
                ChatLog.warn?.('AGENT', `[agent-loop] step ${stepNo} 被 pre-step 拒绝: ${closeReason}`);
                // 广播 agent/error（非致命，recoverable=true）
                emitAgent('agent/error', { turnId, step: stepNo, error: { code: 'PRE_STEP_REJECTED', message: closeReason, recoverable: true } });
                try { sessionStore?.appendStepLog?.(turnId, 'error', `步骤 ${stepNo} 被 pre-step 拒绝：${closeReason}`, stepNo); } catch (e) { /* 忽略 */ }
                yield { type: 'error', error: { code: 'PRE_STEP_REJECTED', message: closeReason, recoverable: true } };
                return;
            }
            if (preStepResult.action === 'rewrite' && Array.isArray(preStepResult.messages)) {
                currentMessages = preStepResult.messages;
                ChatLog.info?.('AGENT', `[agent-loop] step ${stepNo} pre-step 改写消息: ${preStepResult.messages.length} 条`);
            } else if (preStepResult.action === 'inject' && Array.isArray(preStepResult.injectMessages)) {
                currentMessages = [...currentMessages, ...preStepResult.injectMessages];
                ChatLog.info?.('AGENT', `[agent-loop] step ${stepNo} pre-step 注入上下文: ${preStepResult.injectMessages.length} 条`);
            }
            // 广播 pre-step 结果（供外部观察）
            emitAgent('agent/pre-step', { turnId, step: stepNo, action: preStepResult.action });

            // 调用 LLM
            // 只要还没有成功执行过工具，就必须传入tools让模型可以调用
            // 只有在工具执行完成后，下一轮总结回复时才传空tools
            // persistTools=true 时每轮都保留 tools，由 LLM 自行决定继续行动或收尾。
            // 【终态工具收尾】终态（生成类）工具已成功 → 不再下发 tools，
            // 模型只剩"给出最终总结"这一件事，避免在 PPT 生成后无限复读"已完成"。
            const shouldProvideTools = terminalToolDone ? false : (keepToolsEveryRound ? true : !toolsExecuted);
            emitAgent('agent/request', {
                turnId, step: stepNo,
                toolsProvided: shouldProvideTools ? (tools?.length || 0) : 0,
                model: App.config?.model || ''
            });
            // 【能力适配】按模型能力档案决定是否下发原生 tools：
            // Ollama 本地通道 / chat-only 档案（如 qwen2.5）→ 不下发，走文本标记 [tool_name]{args} 模式；
            // function-calling 档案 → 正常下发（与旧行为一致）。决策点：capability-profile.js resolveToolMode
            const toolMode = (App.CapabilityProfile && typeof App.CapabilityProfile.resolveToolMode === 'function')
                ? App.CapabilityProfile.resolveToolMode(App.config?.model, App.config)
                : { nativeTools: true };
            ChatLog.debug?.('AGENT', `[agent-loop] 工具模式: nativeTools=${toolMode.nativeTools} profile=${toolMode.profileName || 'default'}`);
            ChatLog.info?.('AGENT', `[agent-loop] ⚙️ LLM 请求前: keepToolsEveryRound=${keepToolsEveryRound} shouldProvideTools=${shouldProvideTools} tools.length=${tools?.length || 0} nativeTools=${toolMode.nativeTools}`);
            setRunStatus(turnId, 'running', {
                phase: toolsExecuted ? 'summarizing' : 'connecting',
                step: stepNo,
                stepTotal
            });
            const stream = llmClient.stream({
                messages: currentMessages,
                model: App.config?.model,
                config: App.config,
                tools: shouldProvideTools ? tools : [],
                nativeTools: toolMode.nativeTools,
                signal
            });

            let assistantContent = '';
            let assistantReasoning = '';
            let base64Cut = false;  // base64 失控熔断标记
            let _gotRealUsage = false;  // 本轮是否收到后端真实 token 用量
            let toolCalls = [];
            const nativeInvalidCalls = [];  // 原生 tool_calls 中的未注册工具名
            let finishReason = null;
            let streamError = null;

            // 流式消费
            // 重试期间内容只在本地累积，不推给 sessionStore，UI 不刷新
            for await (const evt of stream) {
                if (signal?.aborted) {
                    closeReason = 'user_aborted';
                    yield { type: 'canceled', reason: 'user_aborted' };
                    return;
                }

                if (evt.type === 'model') {
                    ChatLog.info?.('AGENT', `[agent-loop] model event: ${evt.model}`);
                    sessionStore.setTurnModel(turn.id, evt.model);
                    yield evt;
                } else if (evt.type === 'reasoning') {
                    assistantReasoning += evt.delta;
                    // 【防闪屏】重试期间不推给 sessionStore
                    if (!isRetryRound) {
                        sessionStore.appendDelta(turn.id, 'reasoning', evt.delta);
                    }
                    yield evt;
                } else if (evt.type === 'token') {
                    assistantContent += evt.delta;
                    // 【防闪屏】重试期间不推给 sessionStore
                    // 子任务模式下 LLM 文本不是面向用户的，不推给 sessionStore
                    if (!isRetryRound && !silentContent) {
                        if (routeToThinking && !assistantReasoning) {
                            // 多步骤通道：轮内叙述（计划/过渡词/思路）属"思考"，流式写入本轮 thinking 块，
                            // 而非正文。正文只留给最终总结轮（toolCalls.length===0 收尾时再 move 回正文）。
                            // 这样"计划的字"归入思考/进度栏，填空"（无推理文字）"，也不再实时刷在正文里。
                            // !assistantReasoning 确保模型自带独立 reasoning 通道时不重复写入（避免双份）。
                            sessionStore.appendDelta(turn.id, 'reasoning', evt.delta);
                        } else {
                            sessionStore.appendDelta(turn.id, 'content', evt.delta);
                        }
                    }
                    // 【base64 失控熔断】模型幻觉"内嵌图片"时会无限编造 base64（真实案例：
                    // "![](data:image/png;base64,IVBORw0KGgo..."——连 PNG 头都是拼错的），
                    // 一直生成到 max_tokens 才停，表现为"任务无法关闭"、"正在思考..."永远执行中。
                    // 检测到超长 base64/data URL 立即截断正文并跳出流式循环，走正常无工具收尾
                    // （completeTurn 会把所有 running 的 thought 块统一收尾）。
                    if (!base64Cut) {
                        const _tail = assistantContent.length > 1400 ? assistantContent.slice(-1400) : assistantContent;
                        if (/data:[a-z]+\/[a-z0-9.+-]+;base64,[A-Za-z0-9+\/=]{64,}/i.test(_tail) ||
                            /[A-Za-z0-9+\/=]{512,}/.test(_tail)) {
                            base64Cut = true;
                            const _dm = assistantContent.match(/data:[a-z]+\/[a-z0-9.+-]+;base64,/i);
                            let _cutAt = _dm ? _dm.index : Math.max(0, assistantContent.length - 800);
                            // 连同前面残缺的 Markdown 图片语法（![alt]( ）一并去掉
                            const _before = assistantContent.slice(0, _cutAt);
                            const _imgFrag = _before.match(/!\[[^\]\n]{0,80}\]?\s*\(?$/);
                            if (_imgFrag) _cutAt -= _imgFrag[0].length;
                            assistantContent = assistantContent.slice(0, _cutAt).trimEnd() +
                                '\n\n⚠️ 已自动截断：模型试图内嵌图片数据（无真实图像来源，属幻觉输出，会无限生成 base64）。';
                            if (!isRetryRound && !silentContent) {
                                sessionStore.updateTurn(turn.id, { content: assistantContent });
                            }
                            ChatLog.warn?.('AGENT', `[agent-loop] 🛑 base64 失控熔断：已截断正文并终止本轮流式（contentLen=${assistantContent.length}）`);
                            break;
                        }
                    }
                    yield evt;
                } else if (evt.type === 'tool_calls') {
                    // 原生 tool_calls 也校验工具名，未注册的并入 invalid 触发强引导重试
                    const parsed = parseToolCalls(evt.toolCalls, toolRegistry);
                    nativeInvalidCalls.push(...parsed.invalid);
                    toolCalls.push(...parsed.valid);
                    yield evt;
                } else if (evt.type === 'usage') {
                    // 【token 统计】每轮 LLM 调用的真实用量累计到 turn（汇总盒展示）
                    _gotRealUsage = true;
                    sessionStore.addTurnUsage?.(turn.id, evt.tokens);
                    yield evt;
                } else if (evt.type === 'done') {
                    finishReason = evt.finishReason;
                    // 【估算兜底】后端不回传 usage 时（部分 Ollama/中转 API），按字符数粗估输出 token
                    if (!_gotRealUsage && assistantContent && !silentContent) {
                        sessionStore.addTurnUsage?.(turn.id, {
                            out: Math.ceil((assistantContent.length + assistantReasoning.length) / 2),
                            estimated: true
                        });
                    }
                    yield evt;
                    break;
                } else if (evt.type === 'error') {
                    streamError = evt.error;
                    closeReason = 'stream_error';
                    yield evt;
                    return;
                }
            }

            // 【输出截断提示·2026-10-18】模型输出达到长度上限（num_predict/上下文打满，
            // finishReason=length）或被失控熔断强制收尾（runaway_truncated）时，把可见
            // 说明追加到正文。否则用户看到的是"半句话突然没了"——纯思考轮甚至是一个
            // 空白气泡，正是"附件处理没有过程渲染"观感的来源之一。
            // 注意：base64 熔断自己已追加提示且不走 finishReason，这里不会重复。
            if (!streamError && (finishReason === 'length' || finishReason === 'runaway_truncated')) {
                const _truncNotice = finishReason === 'runaway_truncated'
                    ? '\n\n⚠️ 已自动截断：模型输出失控（无意义循环或超长），已被强制停止。建议缩短输入内容或更换模型后重试。'
                    : '\n\n⚠️ 已自动截断：模型输出达到长度上限，回答可能不完整。可缩短输入或调大输出上限后重试。';
                assistantContent = (assistantContent || '') + _truncNotice;
                if (!isRetryRound && !silentContent) {
                    sessionStore.appendDelta(turn.id, 'content', _truncNotice);
                }
                ChatLog.warn?.('AGENT', `[agent-loop] 输出截断提示已注入正文: finishReason=${finishReason} contentLen=${assistantContent.length}`);
            }

            // 如果没有原生 tool_calls，尝试从文本输出中解析 [tool_name]{args} 格式
            // 即使有原生 tool_calls，也解析文本输出用于交叉验证：
            // 原生 tool_calls 常使用默认值（如 xlColumnClustered + A1:G74），而文本输出往往体现用户真实意图（pie/bar）。
            // 当文本里的参数比原生 tool_calls 更具体时，用文本参数覆盖。
            let textParseResult = { valid: [], invalid: [] };
            if (assistantContent) {
                textParseResult = parseTextToolCalls(assistantContent, toolRegistry, turn);
            }
            if (textParseResult.valid.length > 0) {
                ChatLog.info?.('AGENT', `[agent-loop] 从文本中解析到 ${textParseResult.valid.length} 个工具调用`, textParseResult.valid.map(tc => tc.name));
                if (toolCalls.length === 0) {
                    // 无原生 tool_calls：直接用文本解析结果（原逻辑，兼容纯文本模式）
                    toolCalls = textParseResult.valid;
                } else {
                    // 有原生 tool_calls：补充文本中额外规划的工具调用。
                    // 场景：LLM 输出原生 tool_calls（1个）+ 文本中【执行】[tool2]【执行】[tool3] 规划，
                    // 原逻辑只处理原生 tool_calls 导致"规划了3步，只执行了第1步"。
                    // 此处把文本中除原生已调用外的其余工具调用补充到执行列表。
                    // 【关键修正】原实现用 Set(工具名) 判重，对"多个**不同**工具"是对的；
                    // 但当用户的多条要求**全由同一个工具完成**时（实测："设置表头底色、边框、
                    // 对齐、居中"全都是 excel_format_range），原生已有 1 个同名调用 →
                    // 文本里第 2、3、4 个同名调用被整体丢弃：用户列了 5 条要求，只做了 1 条。
                    // 改为按"同名计数"：原生已有 k 个同名调用，就只吸收文本里第 k+1 个起的部分。
                    const nativeNameCount = new Map();
                    for (const tc of toolCalls) {
                        nativeNameCount.set(tc.name, (nativeNameCount.get(tc.name) || 0) + 1);
                    }
                    const textNameSeen = new Map();
                    for (const textTc of textParseResult.valid) {
                        const seenIdx = textNameSeen.get(textTc.name) || 0;
                        textNameSeen.set(textTc.name, seenIdx + 1);
                        // 该序号位置由原生调用占据（其参数会在下方"交叉验证"里被文本补全）
                        if (seenIdx < (nativeNameCount.get(textTc.name) || 0)) continue;
                        toolCalls.push(textTc);
                        ChatLog.info?.('AGENT', `[agent-loop] 文本解析补充同名工具调用 #${seenIdx + 1}: ${textTc.name}`);
                    }
                }
                // 从 assistantContent 中移除工具调用标记，避免重复显示
                const cleanedContent = stripTextToolCalls(assistantContent, toolCalls, textParseResult.invalid, toolRegistry);
                assistantContent = cleanedContent;
                // 【防闪屏】重试期间不推给 sessionStore
                if (!isRetryRound) {
                    sessionStore.updateTurn(turn.id, { content: cleanedContent });
                }
            }

            // 原生 tool_calls 与文本输出的参数对齐
            // 小模型常"两张皮"：原生 tool_calls 用默认值（xlColumnClustered + A1:G74），
            // 但文本里的 [tool_name chart_type: pie data_source: ...] 体现了真实意图。
            // 逐个匹配同名工具调用：若文本参数比原生更具体，则用文本参数"补全+覆盖"
            // （保留原生的 camelCase 参数名，补入文本中的 snake_case 值）。
            if (toolCalls.length > 0 && textParseResult.valid.length > 0) {
                const nameCounter = new Map();
                for (const tc of toolCalls) {
                    const name = tc.name;
                    const idx = (nameCounter.get(name) || 0);
                    nameCounter.set(name, idx + 1);
                    // 找到第 idx 个同名文本解析调用
                    const textMatch = textParseResult.valid.filter(t => t.name === name)[idx];
                    if (!textMatch) continue;
                    // 合并参数：文本中的 snake_case 值补到原生的 camelCase args 上
                    const merged = { ...(tc.args || {}) };
                    const textArgs = textMatch.args || {};
                    // 图表工具：原生 chartType 可能被模型赋予了具体但错误的枚举（如把"柱状图"误当 xlBarClustered/bar），
                    // 此时文本里的意图（chart_type）才是权威。只要文本明确给出了类型，就以文本为准覆盖，
                    // 类型归一化（bar/柱状图/横向柱状图→对应枚举）统一交由 normalizeChartArgs 处理。
                    if (name === 'excel_create_chart' || name === 'excel_create_chart_bar') {
                        if (textArgs.chart_type) {
                            merged.chartType = textArgs.chart_type;
                        }
                        if (textArgs.data_source && !merged.dataRange) {
                            merged.dataRange = textArgs.data_source;
                        } else if (textArgs.data_source && merged.dataRange === 'A1:G74') {
                            // 原生用了整个选区，文本有具体数据范围 → 用文本的
                            merged.dataRange = textArgs.data_source;
                        }
                        if (textArgs.category_source && !merged.categoryRange) {
                            merged.categoryRange = textArgs.category_source;
                        }
                        if (textArgs.title && !merged.title) merged.title = textArgs.title;
                        if (textArgs.include_legend !== undefined && merged.hasLegend === undefined) {
                            merged.hasLegend = textArgs.include_legend;
                        }
                    } else {
                        // 通用合并：文本中非空的参数补到原生 args 里
                        for (const [k, v] of Object.entries(textArgs)) {
                            if ((merged[k] === undefined || merged[k] === null || merged[k] === '') && v !== undefined && v !== null) {
                                merged[k] = v;
                            }
                        }
                    }
                    tc.args = merged;
                    ChatLog.info?.('AGENT', `[agent-loop] 交叉验证：${name} 调用 ${idx + 1} 参数已用文本补全`, merged);
                }
            }

            // 【终态工具收尾】本轮未下发 tools 时，模型输出的任何工具调用都不应执行：
            // 终态工具成功后 we 不再下发 tools，模型只剩"总结"任务；若它仍输出工具调用
            //（小模型常复读历史格式），执行后会导致重复生成文件/重复建图。
            if (!shouldProvideTools && toolCalls.length > 0) {
                ChatLog.warn?.('AGENT', `[agent-loop] 本轮未下发 tools（终态工具已执行或等待总结），忽略模型输出的 ${toolCalls.length} 个工具调用`);
                toolCalls = [];
            }

            // 【高优先级拦截】检测到模型编造了不存在的工具名
            // 合并文本路径与原生 tool_calls 路径的未注册工具名，统一走强引导重试
            const allInvalidCalls = [...(textParseResult.invalid || []), ...nativeInvalidCalls];
            const hasInvalidToolCalls = allInvalidCalls.length > 0;
            if (hasInvalidToolCalls && forceRetryCount < MAX_FORCE_TOOL_RETRIES) {
                forceRetryCount++;
                const invalidNames = allInvalidCalls.map(ic => ic.attemptedName).join(', ');
                ChatLog.warn?.('AGENT', `[agent-loop] ⛔ 检测到模型使用了不存在的工具（重试 ${forceRetryCount}/${MAX_FORCE_TOOL_RETRIES}）: ${invalidNames}`);

                // 【防闪屏】重试期间不清空 UI，只标记 isRetryRound，下一轮流式输出时不推给 sessionStore
                // sessionStore.updateTurn(turn.id, { content: '', reasoning: '' });  // 注释掉，避免 UI 闪屏
                yield { type: 'force_retry', reason: 'invalid_tool_name', round: round + 1 };

                // 把模型的错误回复加入 messages
                currentMessages.push({ role: 'assistant', content: assistantContent });

                // 根据用户意图找出正确的工具建议
                const lastUserMsg = [...currentMessages].reverse().find(m => m.role === 'user');
                const userText = lastUserMsg?.content || '';
                const userContentStr = typeof userText === 'string' ? userText : JSON.stringify(userText);
                const candidates = await findCandidateTools(userContentStr, toolRegistry);
                const suggestionLines = candidates.slice(0, 3).map((c, i) => {
                    const req = c.tool.inputSchema?.required || [];
                    const reqHint = req.length > 0 ? `（需要参数: ${req.join(', ')}）` : '（无需参数）';
                    return `  ${i + 1}. ${c.tool.name} — ${c.tool.displayName || ''}${reqHint}: ${(c.tool.description || '').slice(0, 60)}`;
                }).join('\n');

                // 给出"top-1 工具名 + 完整可拷贝标签 + 唯一答案"硬指令
                // qwen2.5:7b 这种小模型容易"自由发挥"换工具名，必须明确"就这一个、照抄"
                const top1 = candidates[0]?.tool;
                const top1Example = top1 ? _buildToolCallExample(turn, top1) : '（无可用工具）';
                const top1Line = top1 ? `  唯一答案：${top1.name}（${top1.displayName || top1.name}）` : '';
                const retryHint = `⛔ 严重错误！你使用了不存在的工具名：${invalidNames}

这些工具不在你的可用工具列表中！你必须从下面"真实存在的工具"里选，绝不能编造！

用户请求是："${userContentStr.slice(0, 80)}"

${top1Line}

与该意图最匹配的可用工具是：
${suggestionLines}

🔒 硬约束（违反即错误）：
1. 必须使用上面列出的工具名，禁止改写、禁止"近似名"（excel_export_to_csv / excel_copy_worksheet 等都不是真实工具）
2. 完整可拷贝的标签（直接复制到你的回复中）：
${top1Example}
3. 只输出这一个工具标签，不要输出"已完成/正在执行/已复制"等任何假执行文字
4. 中括号顺序：] 必须在 { 之前  →  [name]{"key":"value"}  ✅  不要写成 [name{"key":"value"}]  ❌

请立即重新输出上面"完整可拷贝的标签"那一行来调用工具！`;
                currentMessages.push({ role: 'system', content: retryHint });

                assistantContent = '';
                assistantReasoning = '';
                isRetryRound = true;  // 【防闪屏】标记下一轮为重试轮，不推给 sessionStore
                continue;
            }

            // 如果没有工具调用，检查是否需要引导重试
            // 当 tools 参数为空时（如 image-pipeline 的纯文本分析 / 视觉识别），
            // 表示调用方明确"不需要工具调用"，本轮有正文就直接结束，不要多轮累积重复内容。
            if (toolCalls.length === 0 && (!tools || tools.length === 0)) {
                closeReason = 'no_tools_mode_text_reply';
                // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                }
                yield { type: 'loop_done', reason: 'no_tools_mode_text_reply', round: round + 1 };
                return;
            }
            // ── 提前计算 userOpIntent（多步骤自检 + 光说不做检测链都要用）──
            // 必须用【用户原始问句】判定，不能用 userContentStr（后者会被 prompt-builder 的
            // 【当前选区】元信息块污染，导致天气等无关问句被误判成有操作意图）。
            const _rawUserMsg = (turn && (typeof turn.userMessage === 'string' && turn.userMessage))
                || (turn && turn.message) || '';
            const userOpIntent = !!(App.ChatV2?.routingConstants?.hasExcelOperationIntent?.(_rawUserMsg));

            // 【水位修复】工具已成功执行过 → 本轮无工具调用的文本输出即最终总结，直接收尾。
            // 不能落入下方"光说不做"重试引导：此时 shouldProvideTools=false，模型根本没有
            // 工具可调，注入重试提示会浪费一轮、把 token 水位推高（实测 71→190 critical），
            // 且假执行告警会误报"实际没有调用任何工具"（工具明明已执行成功）。
            // persistTools=true（多轮工具）下同样在此收尾：模型每轮都拿到了 tools，
            // 本轮无工具调用即表示它自主决定停止行动、给出最终结论，故收尾原因标记为
            // llm_decided_stop。两种情况都不应落入下方"光说不做"引导重试。
            //
            // 【多步骤自检】但 persistTools=true + userOpIntent=true 时，不直接信任 LLM 的"总结"。
            // 有些模型在多步骤场景（看板/多图表）下只调了第 1 个工具就说"完成了"——框架不能信。
            // 加一个轻量检查：如果 LLM 输出里包含明显的"未完成信号"，落回光说不做检测链。
            // 未完成信号：见 agent-loop/zero-tool-guards.js（收窄后的未来式短语）
            const _ztg = App.ChatV2 && App.ChatV2.agentLoopZeroToolGuards;
            const isIncomplete = _ztg && typeof _ztg.isIncompleteReply === 'function'
                ? _ztg.isIncompleteReply(assistantContent, { keepToolsEveryRound, userOpIntent })
                : (keepToolsEveryRound && userOpIntent && assistantContent
                    && /(下一步|接下来|还需要|待完成|TODO|要做|我将|我会|next step|to do)/i.test(assistantContent));

            // ── 【要求覆盖度自检】用户列了多条要求，模型只做了第一条就宣称完成 ──
            // 比"未完成信号"更根本：模型写的是"已格式化…成功格式化 A1:G1"（完成语气，
            // 不含"接下来/我将"），INCOMPLETE_SIGNAL 抓不到，于是直接落 llm_decided_stop。
            // 用工具自带的 intentKeywords 反查"用户还要求了哪些工具族、本题还没执行"。
            if (toolCalls.length === 0 && toolsExecuted) {
                const uncovered = detectUncoveredRequirements(turn, tools);
                // 【可观测性·必留】无论有无缺项都落一行宿主日志。
                // 否则"守卫没被走到"与"走到了但没判定出缺项"在生产环境无法区分——
                // 上一轮正是因为 intentKeywords 在工具注册时被丢弃而静默空转，
                // 却没有任何日志能暴露它，排查只能靠反复读代码猜。
                try {
                    const toolCount = Array.isArray(tools) ? tools.length : 0;
                    const kwCovered = (Array.isArray(tools) ? tools : [])
                        .filter(t => Array.isArray(t && t.intentKeywords) && t.intentKeywords.length > 0).length;
                    _hostLog(uncovered.length > 0 ? 'WARN' : 'INFO',
                        `要求覆盖度自检: 未完成=${uncovered.length}`
                        + `${uncovered.length > 0 ? '(' + uncovered.map(u => u.name).join(',') + ')' : ''}`
                        + ` 已执行=${(turn._executedTools || []).join(',') || '无'}`
                        + ` 可见工具=${toolCount}(带intentKeywords=${kwCovered})`
                        + ` 续做引导=${coverageNudgeCount}/${MAX_COVERAGE_NUDGES}`);
                } catch (_e) { /* 日志失败不影响判定 */ }
                pendingUncovered = uncovered;
                let goalUnmet = [];
                try {
                    goalUnmet = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.unmet === 'function')
                        ? (App.ChatV2.turnGoal.unmet(turn) || []) : [];
                } catch (_gu) { goalUnmet = []; }
                if (goalUnmet.length) {
                    const extra = goalUnmet
                        .filter(u => !pendingUncovered.some(p => p && p.displayName === u.label))
                        .map(u => ({ name: u.kind || 'goal', displayName: u.label || u.kind, keywords: [] }));
                    pendingUncovered = pendingUncovered.concat(extra);
                }
                const missingItems = pendingUncovered;
                if (missingItems.length > 0 && coverageNudgeCount < MAX_COVERAGE_NUDGES) {
                    coverageNudgeCount++;
                    const missingLines = missingItems.slice(0, 6).map(u =>
                        `  • ${u.displayName}（${u.name}）`
                        + (u.keywords && u.keywords.length > 0 ? ' —— 用户提到: ' + u.keywords.slice(0, 4).join('、') : '')
                    ).join('\n');
                    const coverageHint = `⏳ 任务还没完成。用户这一条消息里列了多项要求，你只做了其中一部分。\n\n`
                        + `用户明确要求、但本题【还没有执行】的内容：\n${missingLines}\n\n`
                        + `请立刻继续调用工具，只把上面缺的项做完，全部做完后再给最终总结。\n`
                        + `没做完时不要使用"已完成/已成功"这类措辞。`;
                    ChatLog.warn?.('AGENT', `[agent-loop] ⏳ 要求覆盖度自检：还有 ${missingItems.length} 项未完成（${missingItems.map(u => u.name).join(', ')}）→ 续做引导 ${coverageNudgeCount}/${MAX_COVERAGE_NUDGES}`);
                    _hostLog('WARN', `⏳ 要求覆盖度自检触发续做引导: 缺 ${missingItems.map(u => u.name).join(',')}`);
                    yield { type: 'force_retry', reason: 'requirements_not_covered', round: round + 1, missing: missingItems.map(u => u.name) };

                    // 把模型的"已完成"结论喂回去，让它自己看到漏了什么
                    if (assistantContent) {
                        currentMessages.push({ role: 'assistant', content: assistantContent });
                    }
                    currentMessages.push({ role: 'system', content: coverageHint });

                    // 下一轮必须重新带上 tools，否则模型手里没工具、只能再把"已完成"复读一遍。
                    // 沿用"工具失败时保持 tools 可见"的既有机制（shouldProvideTools = !toolsExecuted）。
                    toolsExecuted = false;
                    assistantContent = '';
                    assistantReasoning = '';
                    isRetryRound = true;   // 【防闪屏】续做轮不推给 sessionStore
                    continue;
                }
            }

            // ── 【文件产物交付闸门】"导出 X"必须真的产出文件，不能只做周边动作 ──
            // 实测两次（2026-10-28 19:04 / 19:19，用户两次截图）：
            //   "将当前工作表导出为 CSV 并保存到桌面" → 唯一执行的是 excel_save_workbook（存了 .xlsx），
            //   正文却写"工作表已导出为 CSV 文件…文件路径为：C:\...\xxx.csv"——那个文件根本不存在。
            // 根因：模型把"导出 CSV"理解成"先保存一下工作簿"，而框架从不用"产物工具是否真的执行过"
            // 这一硬事实校验结论。这里补上：没执行 → 劝一轮（重新下发 tools）；劝完仍不执行 →
            // 当场确定性补执行；补执行也失败 → 删掉编造的路径、如实说明（不再放行假结论）。
            const _fdg = App.ChatV2 && App.ChatV2.agentLoopFileDeliveryGate;
            // 【与"承诺未执行"守卫的分工】模型写了"我将调用 excel_export_csv…"这种点名承诺时，
            // 交给下方 2913 起的既有守卫链（它有专门的死锁修复与重试预算），这里不抢跑，
            // 避免同一次失败被两条守卫各劝一轮、白烧 token。只有"没有点名工具"的假结论
            //（如"已导出，路径为 C:\...\x.csv"）才由本闸门兜住。
            let _fNamedResolvable = false;
            if (toolCalls.length === 0) {
                const _fNamed = detectNamedToolCommitment(assistantContent, tools);
                if (_fNamed) {
                    try {
                        const _fReg = App.ChatV2 && App.ChatV2.toolRegistry;
                        const _fTool = (_fReg && typeof _fReg.get === 'function') ? _fReg.get(_fNamed.name) : null;
                        _fNamedResolvable = !!_fTool && !_fTool.isDangerous && !_fTool.IsDangerous;
                    } catch (_fe) { _fNamedResolvable = false; }
                }
            }
            if (toolCalls.length === 0 && !_fNamedResolvable && _fdg && typeof _fdg.detect === 'function') {
                const _fHost = (typeof window !== 'undefined' && window.SmartPilotHost) || App.SmartPilotHost || null;
                const _fSpec = _fdg.detect(_rawUserMsg, _fHost);
                if (_fSpec && !_fdg.isDelivered(turn, _fSpec)) {
                    if (fileNudgeCount < MAX_FILE_NUDGES) {
                        fileNudgeCount++;
                        const _fHint = _fdg.buildNudge(_fSpec, turn);
                        ChatLog.warn?.('AGENT', `[agent-loop] ⏳ 文件产物交付闸门：要 ${_fSpec.label} 但 ${_fSpec.tool} 未执行 → 补做引导 ${fileNudgeCount}/${MAX_FILE_NUDGES}`);
                        _hostLog('WARN', `⏳ 文件产物交付闸门：${_fSpec.tool} 未执行（已执行=${(turn._executedTools || []).join(',') || '无'}）`
                            + ` → 补做引导 ${fileNudgeCount}/${MAX_FILE_NUDGES}`);
                        yield { type: 'force_retry', reason: 'file_product_not_delivered', round: round + 1, missing: [_fSpec.tool] };
                        if (assistantContent) currentMessages.push({ role: 'assistant', content: assistantContent });
                        currentMessages.push({ role: 'system', content: _fHint.text });
                        toolsExecuted = false;
                        assistantContent = '';
                        assistantReasoning = '';
                        isRetryRound = true;   // 【防闪屏】补做轮不推给 sessionStore
                        continue;
                    }
                    // 引导已用尽 → 当场确定性补执行，产物必须真的落盘（或诚实报错）
                    const _fArgs = _fdg.buildArgs(_fSpec, turn);
                    const _fTc = { id: 'syn_file_' + Math.random().toString(36).slice(2, 8), name: _fSpec.tool, args: _fArgs };
                    _hostLog('WARN', `文件产物交付闸门：补做引导用尽仍未执行 → 确定性补执行 ${_fSpec.tool}`);
                    let _fRes = null;
                    try {
                        _fRes = await toolRegistry.invokeTracked(turn, _fSpec.tool, _fArgs, {
                            round: round + 1, ctx: { turn, signal }
                        });
                    } catch (_fErr) {
                        _fRes = { ok: false, error: (_fErr && _fErr.message) || String(_fErr), durationMs: 0, toolCallId: null };
                    }
                    if (_fRes) {
                        const _fOk = !!(_fRes.envelope ? _fRes.envelope.ok === true : _fRes.ok);
                        try { recordToolExecutions(turn, [_fTc], [_fRes], round); } catch (_rf) { /* 记录失败不影响正文 */ }
                        emitAgent('tool/result', {
                            turnId: turn.id, step: stepNo, toolName: _fTc.name, callId: _fTc.id,
                            ok: _fOk, error: _fRes.error || null,
                            durationMs: _fRes.durationMs || 0, toolCallId: _fRes.toolCallId || null
                        });
                        try {
                            sessionStore?.appendStepLog?.(turn.id, 'tool_return',
                                '工具 ' + _fTc.name + ' ' + (_fOk ? '成功' : '失败') + '（文件产物补执行）', stepNo);
                        } catch (_se) { /* 忽略 */ }
                        yield {
                            type: 'tool_call_result', toolCallId: _fRes.toolCallId, toolName: _fTc.name,
                            ok: _fOk, result: _fRes.result, error: _fRes.error, durationMs: _fRes.durationMs
                        };
                        // 先删掉没有事实支撑的"已导出/已保存 + 路径"句子，再补真实结论
                        assistantContent = _fdg.scrubClaims(assistantContent || '', _fSpec);
                        if (_fOk) {
                            const _line = _fdg.describeSuccess(_fSpec, _fRes);
                            assistantContent = (assistantContent ? assistantContent.replace(/\s+$/, '') + '\n\n' : '') + _line;
                            _hostLog('INFO', `文件产物交付闸门：模型未产出 → 确定性补执行 ${_fSpec.tool} 成功，正文已用真实结果覆盖`);
                        } else {
                            pendingUncovered = (Array.isArray(pendingUncovered) ? pendingUncovered : []).concat([
                                { name: _fSpec.tool, displayName: _fSpec.displayName, keywords: [] }
                            ]);
                            assistantContent = (assistantContent ? assistantContent.replace(/\s+$/, '') + '\n\n' : '')
                                + _fdg.describeFailure(_fSpec, _fRes.error || '');
                            _hostLog('WARN', `文件产物交付闸门：自动补导出失败（${_fRes.error || '未知原因'}）→ 正文如实说明`);
                        }
                    }
                    // 补执行过（无论成败）都算"已跑过工具"：让下面的收尾逻辑按总结轮处理，
                    // 否则会落入"光说不做"守卫链，把刚写好的真实结论当成空话再劝一轮 / 清空正文。
                    toolsExecuted = true;
                }
            }

            // ── 【图表交付闸门】图表请求不能"只给文字"就交付 ──
            // 为什么必须在收尾之前：本轮无工具调用 + 整轮已执行过工具（多数是守卫链补的
            // excel_read_range）时，下面会直接 closeReason='llm_decided_stop' 收尾。
            // 覆盖度自检靠 intentKeywords 反查，判不出"缺一张图"（实测 07:46 未完成=0）。
            // 这里用"图是否存在"这一硬事实补上：没建图就不许收尾，先劝一轮；
            // 劝完仍无图 → 当场确定性补链（勿再靠 _chartBackfillPending 穿透到轮末，
            // 否则会落入零工具守卫并以 no_tool_calls 零图收场）。
            if (toolCalls.length === 0 && !turn._chartDone && isChartIntent(turn, _rawUserMsg)) {
                if (toolsExecuted && chartNudgeCount < MAX_CHART_NUDGES) {
                    chartNudgeCount++;
                    const _hint = buildChartNudgeHint(turn, _rawUserMsg, chartTypeOverride);
                    ChatLog.warn?.('AGENT', `[agent-loop] ⏳ 图表交付闸门：本轮无工具调用且尚未建图（图表意图）→ 补图引导 ${chartNudgeCount}/${MAX_CHART_NUDGES}`);
                    _hostLog('WARN', `⏳ 图表交付闸门触发补图引导 ${chartNudgeCount}/${MAX_CHART_NUDGES}: 未建图，已执行=${(turn._executedTools || []).join(',') || '无'}`
                        + ` 建议聚合=${_hint.src ? _hint.src.dimColumn + (_hint.src.valueColumn ? '×' + _hint.src.valueColumn : '') + '@' + _hint.src.range : '无表头可推断'}`);
                    yield { type: 'force_retry', reason: 'chart_not_generated', round: round + 1, missing: ['excel_create_chart'] };
                    if (assistantContent) currentMessages.push({ role: 'assistant', content: assistantContent });
                    currentMessages.push({ role: 'system', content: _hint.text });
                    toolsExecuted = false;
                    assistantContent = '';
                    assistantReasoning = '';
                    isRetryRound = true;   // 【防闪屏】补图轮不推给 sessionStore
                    continue;
                }
                // 引导已用尽（或引导后本轮仍只写字）→ 当场补链，不再落入零工具守卫。
                if (chartNudgeCount >= MAX_CHART_NUDGES && !turn._aggFallbackTried && !anyToolFailed) {
                    assistantContent = '';
                    assistantReasoning = '';
                    _hostLog('WARN', '图表交付闸门：补图引导已用尽仍未建图 → 当场兜底补链（aggregate→create_chart）');
                    const _gateBfToolCalls = [];
                    const _gateBfResults = [];
                    const _bfOut = yield* runChartBackfill({
                        turn, toolCalls: _gateBfToolCalls, localResults: _gateBfResults,
                        round, stepNo, signal, rawMsg: _rawUserMsg, chartTypeOverride
                    });
                    if (_bfOut.anyToolFailed) anyToolFailed = true;
                    if (_bfOut.chartDone) {
                        terminalToolDone = true;
                        toolsExecuted = true;
                        if (_gateBfToolCalls.length > 0) {
                            recordToolExecutions(turn, _gateBfToolCalls, _gateBfResults, round);
                        }
                        closeReason = 'chart_backfill_done';
                        yield { type: 'loop_done', reason: closeReason, round: round + 1 };
                        return;
                    }
                    // 补链未能建图（无表头等）→ 交回后续守卫/收尾，勿再死循环劝模型
                }
            }

            // 【死锁修复·关键口】总结轮收尾前先查"点名承诺"：模型在总结轮（手里没工具）
            // 写下"我将调用 excel_export_csv…"时，模型自己已判定任务没做完。
            // 若该工具真实可解析（本轮可见或注册表有、非危险），不得在此静默收尾——
            // 放行到下方守卫链的承诺重试（重试轮会重新下发 tools，见承诺分支的死锁修复），
            // 否则"永不执行的计划"原样展示给用户（实测 19:34 导出 CSV）。
            // 注册表里也没有的工具（幻觉点名）不拦截，正常收尾、不烧重试预算。
            const _namedCommitEarly = (toolCalls.length === 0 && toolsExecuted)
                ? detectNamedToolCommitment(assistantContent, tools) : null;
            let _namedCommitResolvable = false;
            if (_namedCommitEarly) {
                try {
                    const _regEarly = App.ChatV2 && App.ChatV2.toolRegistry;
                    const _tEarly = (_regEarly && typeof _regEarly.get === 'function')
                        ? _regEarly.get(_namedCommitEarly.name) : null;
                    _namedCommitResolvable = !!_tEarly && !_tEarly.isDangerous && !_tEarly.IsDangerous;
                } catch (_e) { _namedCommitResolvable = false; }
            }
            if (toolCalls.length === 0 && toolsExecuted && !isIncomplete && !_namedCommitResolvable) {
                closeReason = keepToolsEveryRound ? 'llm_decided_stop' : 'summary_after_tools';
                // 【JSON 回声守卫】弱模型常把工具返回的 JSON 原样当"最终总结"复读
                // （实测 excel_merge_sheets 成功后正文就是整坨 {success,sourceSheets,mergedRows,data}）。
                // 命中已知工具结果结构时替换为确定性友好摘要；自然语言总结不受影响。
                let _echoFixed = null;
                try {
                    const _util = App.ChatV2 && App.ChatV2.toolRegistryUtil;
                    _echoFixed = _util && typeof _util.replaceJsonEchoWithFriendly === 'function'
                        ? _util.replaceJsonEchoWithFriendly(assistantContent) : null;
                    if (!_echoFixed && _util && typeof _util.replaceToolDumpEcho === 'function') {
                        _echoFixed = _util.replaceToolDumpEcho(assistantContent, turn);
                    }
                } catch (_je) { /* 守卫失败不影响原路径 */ }
                if (_echoFixed) {
                    _hostLog('WARN', '收尾正文是工具 JSON 回声 → 已替换为友好摘要');
                    ChatLog.warn?.('AGENT', '[agent-loop] 收尾正文是工具 JSON 回声 → 已替换为友好摘要');
                }
                // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                // 【过程语言清洗】模型任务没做完时会用"接下来，将继续…请稍候。"这类悬挂承诺句收尾
                // （实测 2026-10-22：批量美化只做 1 张表，正文却是过程语言+任务完成）。
                // 覆盖度/图表闸门有引导预算，兜底时至少把承诺句从最终总结里剔掉。
                let finalContent = withCoverageNote(_echoFixed || assistantContent);
                try {
                    const _ztgStrip = App.ChatV2 && App.ChatV2.agentLoopZeroToolGuards;
                    if (_ztgStrip && typeof _ztgStrip.stripDanglingPromises === 'function') {
                        const _beforeStrip = finalContent;
                        finalContent = _ztgStrip.stripDanglingPromises(finalContent);
                        if (_beforeStrip && finalContent !== _beforeStrip) {
                            _hostLog('WARN', '收尾正文含悬挂承诺句（"…请稍候/将继续…"）→ 已从最终总结中剔除');
                            ChatLog.warn?.('AGENT', '[agent-loop] 收尾正文含过程语言承诺句 → 已清洗');
                        }
                    }
                } catch (_sdErr) { /* 清洗失败不影响原路径 */ }
                try {
                    const envApi = App.ChatV2 && App.ChatV2.toolEnvelope;
                    if (envApi && typeof envApi.buildVerdict === 'function') {
                        const verdict = envApi.buildVerdict(turn);
                        turn._verdict = verdict;
                        finalContent = envApi.applyVerdict(finalContent, verdict);
                    }
                } catch (_vd) { /* 验收失败不影响原文 */ }
                if (typeof finalContent !== 'undefined' && finalContent && finalContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: finalContent });
                }
                yield { type: 'loop_done', reason: closeReason, round: round + 1 };
                return;
            }
            // 多步骤场景 LLM 说"下一步我将..." → 不落 loop_done，让下面的光说不做检测链接管
            if (isIncomplete) {
                ChatLog.info?.('AGENT', `[agent-loop] 多步骤自检: LLM 输出包含"未完成信号"("${assistantContent?.slice(0, 80)}") → 不落 llm_decided_stop，引导继续调工具`);
            }
            // 【内容处理·免工具放行】总结 / 润色 / 翻译 / 提炼 等只读任务，当前文档/幻灯片文本
            // 已随上下文注入（turn.context.documentText 非空），模型无需再调读取工具即可作答。
            // 原"零工具调用→假执行"守卫会误伤这类正常文字回复（实测：用户要"总结当前幻灯片"，
            // 正文已在 system prompt 的【当前幻灯片内容】里，模型直接总结却被追加"未执行"警告，
            // 用户误以为在编造）。仅当内容确实没注入（读不到）时才保留守卫、引导去读。
            let _contentTaskExempt = false;
            let _contentTaskNoContent = false;
            try {
                const _docTextInCtx = (turn && turn.context)
                    ? String(turn.context.documentText || turn.context.selectionText || '').trim() : '';
                const _pb = App.ChatV2 && App.ChatV2.promptBuilder;
                // 【修复】意图判定必须用【原始】用户消息，不能用 currentMessages 里的——
                // 那份已被 prompt-builder 追加"若用户要求总结/提炼/润色…严禁编造"的反编造后缀，
                // 其中的"总结/提炼/润色"会污染 detectContentTaskIntent：用户问"你有多少工具可用"
                // 这类元问题被误判成内容任务 + documentText 为空 → 替换守卫把模型的正确回答
                // 整条替换成"未能读取到当前文档的文字内容"（实测 Word 16:09 复现）。
                const _rawUserForIntent = String(
                    (turn && (turn._originalUserMessage || turn.userMessage || turn.message))
                    || latestUserMsgText(currentMessages)
                    || '');
                const _isContentTask = !!_pb && typeof _pb.detectContentTaskIntent === 'function'
                    && _pb.detectContentTaskIntent(_rawUserForIntent);
                _contentTaskExempt = _docTextInCtx.length > 0 && _isContentTask;
                // 【内容任务·无素材】总结/润色/翻译类任务，但上下文里没有任何可读文本
                // → 模型手上没有素材，只会凭标题编造（实测 qwen2.5 编出"项目A: 已完成30%…"）。
                // 仅对非 Excel 宿主生效（Excel 的素材走选区/range，不适用本判定），
                // 且必须排除"素材来自附件/引用区域/选区"的正常情形，避免误伤。
                const _host = (typeof window !== 'undefined' && window.SmartPilotHost) || null;
                const _isNonExcel = !!(_host && !_host.isExcel);
                // 【修复】非 Excel 宿主下 selection/referenceRegions 不会被注入提示词
                // （prompt-builder 非 Excel 分支只注入 documentText），把它们当"其他素材来源"
                // 会让 _contentTaskNoContent 恒为 false → 编造正文逃过替换守卫
                // （实测：PPT 总结任务编出"项目A:已完成30%…"仅追加弱警告）。故仅附件算真实来源。
                // selectionText 已并入 _docTextInCtx，不再单独当“伪来源”。
                const _hasOtherSource = (Array.isArray(turn && turn.attachments) && turn.attachments.length > 0);
                _contentTaskNoContent = _isContentTask && _docTextInCtx.length === 0
                    && _isNonExcel && !_hasOtherSource;
            } catch (_e) { _contentTaskExempt = false; _contentTaskNoContent = false; }
            if (_contentTaskExempt) {
                ChatLog.info?.('AGENT', `[agent-loop] 内容处理任务且上下文已注入文本 → 免工具放行，跳过假执行守卫`);
                // 【修复·终止空转】纯只读内容任务（总结/润色/翻译/提炼）已注入正文且本轮未调工具，
                // 模型已给出文本 → 必须立即收尾。原实现只打日志、无收尾分支，导致
                // while(round<=maxRounds) 一直重问 LLM（多步骤通道下弱模型每轮仅回 13~18 字、
                // finishReason=stop），卡到上限；而偶数轮是思考轮(0 字)，末轮恰落 0 字 →
                // 最终 content 为空 → 触发"模型返回空内容"兜底，用户看到废话（实测 qwen2.5:7b）。
                if (toolCalls.length === 0) {
                    closeReason = 'content_task_exempt_done';
                    if (assistantContent && assistantContent.trim()) {
                        sessionStore.updateTurn(turn.id, { content: assistantContent });
                    }
                    yield { type: 'loop_done', reason: closeReason, round: round + 1 };
                    return;
                }
            }

            // 当 tools 参数非空时，才检查是否需要引导重试（模型应调用工具但没调用）
            if (toolCalls.length === 0 && tools && tools.length > 0 && !_contentTaskExempt) {
                // 【假执行清洗标志】本轮内任一分支已处理假执行文字（替换/追加警告）后置位，
                // 防止下方放行口二次处理。
                let fakeTextHandled = false;

                // 【可观测性·必留】零工具调用进入守卫链时，先把各项判定依据落盘。
                // 为什么必须：整条链（弱匹配引导 / 假执行清洗 / 承诺未执行 / 诚实收尾·绝对兜底）
                // 全部嵌在上行 `tools && tools.length > 0` 里——"守卫压根没被走到"与
                // "走到了但条件不成立"，在生产环境无法区分。实测出现"0 工具调用却在正文里写
                // 已经完成…"且没有任何守卫日志，只能靠反复读代码猜是哪一项断掉的。
                try {
                    _hostLog('WARN', `零工具调用进入守卫链: tools=${Array.isArray(tools) ? tools.length : -1}`
                        + ` userOpIntent=${userOpIntent} forceRetry=${forceRetryCount}/${MAX_FORCE_TOOL_RETRIES}`
                        + ` 已执行=${(turn._executedTools || []).join(',') || '无'}`
                        + ` 正文=${(assistantContent || '').length}字`);
                } catch (_e) { /* 日志失败不影响判定 */ }
                const lastUserMsg = [...currentMessages].reverse().find(m => m.role === 'user');
                const userText = lastUserMsg?.content || '';
                const userContentStr = typeof userText === 'string' ? userText : JSON.stringify(userText);

                // 【无关问答护栏】是否有 Excel 处理意图（操作动词/选区引用）。
                // 对天气、新闻、闲聊等无关问句，以下整套"模型该调工具却没调"的
                // 引导重试/强制兜底都不应启动——否则会把 clear_range 等破坏性工具
                // 强推到用户数据上（如"东莞今天天气"被误判成清空 A2:G22）。
                // 注意：必须用【用户原始问句】判定，不能用 userContentStr——后者是发给模型的
                // user content，会被 prompt-builder 的【当前选区】元信息块污染（"当前选区/选中"
                // 字样命中正则），导致天气等无关问句被误判成有操作意图而强执破坏性工具。
                const rawUserMsg = _rawUserMsg;  // 复用上方已计算的值，避免重复逻辑

                // 用工具注册表动态检测：用户消息与可用工具的匹配度是否足够高
                // 如果工具注册表中有高匹配度（score >= 20）的工具但模型没调用，说明模型漏掉了工具调用
                // 如果 run() 调用方指定了 tools 白名单（如 text-pipeline 多图表循环只允许 excel_create_chart），
                // 则只在白名单内搜索候选工具，不要全量搜索 toolRegistry 中的 skill 工具，
                // 否则重试时会引入不在白名单里的 skill 工具（如 skill_SmartChartWizard），导致模型尝试调用但实际不可用。
                // 用用户原话打分。发给模型的 user 消息后面拼了「数据区 / 没有手选」等提示，
                // 那些字会把「美化所有表」打成「写入Excel数据」，兜底因缺 values 被跳过。
                const candidates = await findCandidateTools(String(rawUserMsg || userContentStr || ''), toolRegistry, 3, tools);

                // ── 【通用框架·执行去重】用已执行次数对 candidates 降分排序 ──
                // 通用框架级逻辑：任何多步骤场景（看板/多图表/复合操作）LLM 卡壳时，
                // 强制兜底不应该盲目选静态打分第一名（如"读数据"），而应该优先选还没执行过的工具。
                //
                // 示例：看板场景已执行 aggregate×2 + create_chart×1
                //   → findCandidateTools 的静态分数: read_range=180, aggregate=160, create_chart=140
                //   → 执行去重后降分: read_range=180-100=80, aggregate=160-200=-40, create_chart=140-0=140
                //   → 排序后第一名变成 create_chart（正确！下一步应该继续制图）
                //
                // 逻辑：已执行 count 次 → score 减去 count×50（50 分权重够让第二名跃过第一名）
                // 不排除任何候选（以防确实需要重复调用，如多次 aggregate 不同维度），只降分排序。
                if (turn && Array.isArray(turn._executedTools) && turn._executedTools.length > 0 && candidates.length > 0) {
                    const execCount = {};
                    for (const n of turn._executedTools) execCount[n] = (execCount[n] || 0) + 1;
                    const wasExecuted = Object.keys(execCount).length;
                    if (wasExecuted > 0) {
                        for (const c of candidates) {
                            const cName = c.tool?.name || c.tool?.Name || '';
                            const cnt = execCount[cName] || 0;
                            if (cnt > 0) c.score = (c.score || 0) - cnt * 50;
                        }
                        // 按新分数降序重排
                        candidates.sort((a, b) => (b.score || 0) - (a.score || 0));
                        // 恢复 hasStrongMatch 标志（因为 topCandidate 变了）
                        ChatLog.info?.('AGENT', `[agent-loop] 执行去重: 已执行 ${turn._executedTools.length} 次(${wasExecuted}种) → candidates: ${candidates.map(c => `${c.tool?.name}(${(c.score||0).toFixed(0)})`).join(', ')}`);
                    }
                }

                // 全表格式化还没取过清单时，零工具不要去写单元格。先补 excel_get_sheets，后面才有表名和范围。
                if (!(turn && Array.isArray(turn._executedTools) && turn._executedTools.some(function (n) { return /get_sheets$/i.test(String(n || '')); }))) {
                    const _plan = (App.ChatV2.rangeResolver && typeof App.ChatV2.rangeResolver.decideFormatTarget === 'function')
                        ? App.ChatV2.rangeResolver.decideFormatTarget(String(rawUserMsg || ''), turn && turn.context)
                        : null;
                    if (_plan && (_plan.sheetScope === 'all' || _plan.sheetScope === 'others') && _plan.rangeMode !== 'skip'
                        && /(美化|格式化|主题|配色|样式|统一)/.test(String(rawUserMsg || ''))) {
                        const _sheetsTool = (tools || []).find(function (t) {
                            return /get_sheets$/i.test(String((t && (t.name || t.originalName)) || '').replace(/^excel-local__/, '').replace(/^mcp_/, ''));
                        });
                        if (_sheetsTool) {
                            const _rest = candidates.filter(function (c) { return c && c.tool && c.tool.name !== _sheetsTool.name; });
                            candidates.length = 0;
                            candidates.push({ tool: _sheetsTool, score: 1000, suggestedArgs: {} });
                            _rest.forEach(function (c) { candidates.push(c); });
                            _hostLog('WARN', '零工具且要格式化全部工作表 → 候选改为先取工作表清单');
                        }
                    }
                }

                const topCandidate = candidates[0];
                let hasStrongMatch = candidates.length > 0 && candidates[0].score >= TOOL_MATCH_THRESHOLD;

                // 假执行文字检测：模型声称"已调用/已完成/已美化/已处理"等
                // 但实际没有输出任何工具标签 → 视为幻觉执行，强制重试（即使候选分数略低于阈值 20）
                // 【按用户推荐修复】"经"字可省：模型常写"已**经**完成/已**经**调用"，
                // 原正则要求"已"与动词紧邻 → "已经完成"漏判（实测：0 工具调用却正文写"已经完成了…美化操作"）。
                // 【2026-10-10】再补"成功"：实测写"已成功美化选中区域…背景颜色 #FFFFCC"，
                // 旧式 已(?:经)?美化 中间插了"成功"整段漏判 → 任务完成·未调用工具 + 表仍未动。
                // 另认"成功美化/操作完成/结果摘要"等无"已"字的完成腔。
                const FAKE_EXECUTION_PATTERN = /已(?:经)?(?:成功)?(?:调用|执行|完成|美化|格式化|处理|清洗|排序|筛选|删除|插入|合并|拆分|生成|创建|设置|更新|导出|导入|保存|调整|填充|转换|添加|修改|应用|整理|汇总|计算|分析|生成图表|应用格式)|成功(?:美化|格式化|完成)|操作完成|结果摘要/;
                // 【守卫修正·关键】假执行判定必须看【整轮】已执行工具，不能只看本轮 toolCalls。
                // 实测 bug（Word 08:56）：模型先成功调用 word_get_selection_text /
                // word_format_selected_table_cell，随后输出"已完成"总结（该轮无工具调用），
                // 旧逻辑把这条正确总结判成"假执行"→ 先烧掉 2 次强制重试，耗尽后把总结整条
                // 替换成"实际没有调用任何工具，Excel 未被修改"，用户被反向误导（工具明明成功）。
                // turn._executedTools 是框架维护的整轮执行记录，是"是否真的干过活"的唯一可靠依据。
                const _turnHasExecutedTools = !!(turn && Array.isArray(turn._executedTools)
                    && turn._executedTools.length > 0);
                // 告警文案的宿主感知称谓：在 Word/PPT 里说"Excel 未被修改"会误导用户（实测）。
                const _docLabel = (function () {
                    try {
                        const h = (typeof window !== 'undefined' && window.SmartPilotHost) || null;
                        if (h && h.isPowerPoint) return '演示文稿未被修改';
                        if (h && h.isExcel) return 'Excel 未被修改';
                        if (h && (h.isWord || h.kind === 'word')) return 'Word 文档未被修改';
                    } catch (_e) { /* 宿主判定失败则用通用称谓 */ }
                    return '文件未被修改';
                })();
                // 【守卫修正·引导型回复】模型若是在引导用户操作（而非声称自己已完成），
                // 不应判为假执行：例如"未选中具体对象 / 请先选择表格 / 请按照以下步骤…"，
                // 这类文字常被 FAKE_EXECUTION_PATTERN 误命中（含"已…设置"等），导致整段引导
                // 被替换成"声称已执行但未调用工具"，用户被反向误导（实测 PPT：未选中表格时模型
                // 给出正确引导却被整条替换）。命中引导标记 → isFakeExecution 强制置 false，
                // 且下方兜底替换/追加警告分支一律跳过，保留模型原始引导文字。
                const USER_GUIDANCE_RE = /(未选中(?:具体)?对象|请按照以下步骤|按照以下步骤|请先选择|先选择表格|先选中表格|请先选中|请先确定|请先确认)/;
                const isUserGuidance = !!(assistantContent && USER_GUIDANCE_RE.test(assistantContent));
                let isFakeExecution = !!assistantContent
                    && FAKE_EXECUTION_PATTERN.test(assistantContent)
                    && (hasStrongMatch || (candidates.length > 0 && candidates[0].score >= 10))
                    && !_turnHasExecutedTools   // 整轮已执行过工具 → 这是执行后的总结，不是假执行
                    && !isUserGuidance;         // 引导型回复（请用户操作）不算假执行

                // ── 【承诺未执行】模型点名了具体工具 + 未来时承诺，但本轮零工具调用 ──
                // 这条守卫【刻意不依赖 userOpIntent】：跟进型消息（"不够美观"）往往一个操作动词都没有，
                // userOpIntent=false → 上面整套引导/兜底按"无关问答护栏"被设计性跳过，
                // 用户最终看到"任务完成 · 未调用工具"＋一段永远不会执行的计划（实测失败）。
                // 模型自己点名要调哪个工具，本身就是最强的"确有 Excel 操作意图"证据。
                const namedCommitment = detectNamedToolCommitment(assistantContent, tools);
                if (namedCommitment && forceRetryCount < MAX_FORCE_TOOL_RETRIES) {
                    forceRetryCount++;
                    ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 承诺未执行：模型点名 ${namedCommitment.name}`
                        + `（本轮${namedCommitment.visible ? '可见' : '不可见'}）但零工具调用 → 引导重试 ${forceRetryCount}/${MAX_FORCE_TOOL_RETRIES}`);
                    yield { type: 'force_retry', reason: 'named_tool_commitment_without_call', round: round + 1 };

                    if (assistantContent) {
                        currentMessages.push({ role: 'assistant', content: assistantContent });
                    }

                    let commitHint;
                    if (namedCommitment.visible) {
                        const tgt = (tools || []).find(t => String((t && (t.name || t.originalName)) || '').toLowerCase() === namedCommitment.name);
                        const req = (tgt && tgt.inputSchema && Array.isArray(tgt.inputSchema.required)) ? tgt.inputSchema.required : [];
                        commitHint = `你刚才明确表示要调用 ${namedCommitment.name}，但本轮没有产生任何工具调用——用户不会看到任何变化。\n\n`
                            + `现在立即调用它，不要再输出任何计划或说明文字：${namedCommitment.name}`
                            + (req.length > 0 ? `（必填参数: ${req.join(', ')}）` : '（无需参数）');
                    } else {
                        // 【死锁修复】点名工具"本轮不可见"≠注册表里没有。原提示只叫模型
                        // "别调不存在的工具"，模型永远完不成任务（实测：导出 CSV 卡死在这）。
                        // 从注册表找回并注入本轮工具列表，让重试轮有真工具可调。
                        let recovered = null;
                        try {
                            const _reg = App.ChatV2 && App.ChatV2.toolRegistry;
                            const _cand = (_reg && typeof _reg.get === 'function') ? _reg.get(namedCommitment.name) : null;
                            const _candDangerous = !!(_cand && (_cand.isDangerous || _cand.IsDangerous || (_cand.meta && _cand.meta.isDangerous)));
                            if (_cand && !_candDangerous) {
                                if (!(tools || []).some(t => String((t && (t.name || t.originalName)) || '').toLowerCase() === namedCommitment.name)) {
                                    (tools = tools || []).push(_cand);
                                }
                                recovered = _cand;
                            }
                        } catch (_e) { /* 注册表不可用 → 保持原行为 */ }
                        if (recovered) {
                            const req = (recovered.inputSchema && Array.isArray(recovered.inputSchema.required)) ? recovered.inputSchema.required : [];
                            commitHint = `你刚才明确表示要调用 ${namedCommitment.name}，它现在已加入你的可用工具列表。立即调用它，不要再输出任何计划或说明文字：${namedCommitment.name}`
                                + (req.length > 0 ? `（必填参数: ${req.join(', ')}，不确定时可留空由系统自动补全）` : '（无需参数）');
                            ChatLog.warn?.('AGENT', `[agent-loop] 承诺未执行：点名 ${namedCommitment.name} 本轮不可见 → 已从注册表注入工具列表后重试`);
                        } else {
                            const avail = (tools || []).map(t => t && (t.name || t.originalName)).filter(Boolean);
                            commitHint = `你打算使用 ${namedCommitment.name}，但它本轮不在你的可用工具列表中——不要调用不存在的工具。\n\n`
                                + `你当前真实可用的工具是：${avail.slice(0, 20).join(', ')}\n\n`
                                + `请从中选择能完成用户请求的工具并立即调用；若确实没有可用工具，如实说明缺少什么能力，不要只写"我将要做什么"。`;
                        }
                    }
                    currentMessages.push({ role: 'system', content: commitHint });

                    assistantContent = '';
                    assistantReasoning = '';
                    // 【死锁修复·关键】重试轮必须重新下发 tools。原逻辑 toolsExecuted 仍为 true
                    // → shouldProvideTools=false → 模型手里没工具，重试提示喊"立即调用"也调不了，
                    // 2 次重试全部空转后把"永不执行的计划"展示给用户（实测 19:34 导出 CSV）。
                    // 与 2752 行覆盖度续做轮的既有做法一致。
                    toolsExecuted = false;
                    isRetryRound = true;  // 【防闪屏】重试轮不推给 sessionStore
                    continue;
                }
                if (namedCommitment) {
                    ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 承诺未执行且重试预算耗尽（${namedCommitment.name}），将在收尾处补诚实提示`);
                }

                // 【光说不做·静默放行口修复】当模型只回复了文字、没调用任何工具，
                // 且候选工具存在"弱匹配"（分数未达强匹配阈值但并非全无关联）时，
                // 原逻辑会直接落到下方 loop_done 静默放行，用户看到一段"假完成/无关文字"。
                // 这里给一次带候选工具建议的引导重试，促使模型明确输出 [工具名] 标签。
                // 分数过低（<5）视为纯咨询无关，仍静默放行避免打断正常问答。
                const WEAK_MATCH_GUIDE_THRESHOLD = 5;
                const isWeakMatchSilentPass = !hasStrongMatch && !isFakeExecution
                    && userOpIntent
                    && candidates.length > 0 && candidates[0].score >= WEAK_MATCH_GUIDE_THRESHOLD
                    && forceRetryCount < MAX_FORCE_TOOL_RETRIES;

                if (isWeakMatchSilentPass) {
                    forceRetryCount++;
                    const weakTop = candidates[0].tool;
                    ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 静默放行口命中：模型回复文字但未调用工具，弱匹配 ${weakTop?.name || '无'}(${weakTop?.displayName || ''}) score=${candidates[0].score}，引导重试 ${forceRetryCount}/${MAX_FORCE_TOOL_RETRIES}`);
                    yield { type: 'force_retry', reason: 'weak_match_without_call', round: round + 1 };

                    // 把模型的错误回复加入 messages
                    if (assistantContent) {
                        currentMessages.push({ role: 'assistant', content: assistantContent });
                    }

                    const suggestionLines = candidates.slice(0, 3).map((c, i) => {
                        const req = c.tool.inputSchema?.required || [];
                        const reqHint = req.length > 0 ? `（需要参数: ${req.join(', ')}）` : '（无需参数）';
                        return `  ${i + 1}. ${c.tool.name} — ${c.tool.displayName || ''}${reqHint}: ${(c.tool.description || '').slice(0, 60)}`;
                    }).join('\n');

                    const weakTag = _toolCallTag(weakTop) || '工具名';
                    // 【守卫修正】整轮已执行过工具时不得再说"没有调用任何工具"——事实错误会把模型
                    // 引向重复执行（实测 Word：已成功调用 word_format_selected_table_cell 后，
                    // 系统仍告知其"未调用任何工具"）。此处按"续做"口径表述。
                    const _ranPrefix = _turnHasExecutedTools
                        ? `本任务【已执行过】：${(turn._executedTools || []).join('、')}。你刚才只输出了文字，没有继续调用工具。`
                        : `你刚才只输出了文字，但实际没有调用任何工具。`;
                    const retryHint = `${_ranPrefix}如果用户请求需要执行操作，必须通过输出 [${weakTag}]{"参数":"值"} 标签来调用 MCP 工具，单纯的文字回复无法实际完成操作。\n\n检测到与用户意图可能相关的工具（匹配度不高但存在关联）：\n${suggestionLines}\n\n请判断：\n1. 若用户请求还有未完成的步骤，立即只输出最匹配工具的 [${weakTag}] 标签；\n2. 若已全部完成或是纯咨询问题，直接如实作答即可（已完成的可正常说明"已完成"），但严禁编造没有做过的操作。`;
                    currentMessages.push({ role: 'system', content: retryHint });

                    assistantContent = '';
                    assistantReasoning = '';
                    // 【死锁修复】同承诺重试：重试轮必须重新下发 tools，否则模型没工具可调
                    toolsExecuted = false;
                    isRetryRound = true;  // 【防闪屏】标记下一轮为重试轮，不推给 sessionStore
                    continue;
                }

                if ((hasStrongMatch || isFakeExecution) && userOpIntent && forceRetryCount < MAX_FORCE_TOOL_RETRIES) {
                    // 【chat-only 快速通道】Ollama 等本地小模型根本不会输出工具标签（无论 nativeTools 怎么配），
                    // 重试 2 轮也是空回复或假文字 → 直接强制兜底，省 token 省时间。
                    // 条件：Ollama apiType + 有强匹配 + 工具非危险
                    const apiType = (App.config && App.config.apiType) || '';
                    const isOllama = apiType === 'ollama';
                    const topTool = candidates[0].tool;
                    const topIsDangerous = !!(topTool && (topTool.isDangerous || topTool.IsDangerous || (topTool.meta && topTool.meta.isDangerous)));
                    ChatLog.info?.('AGENT', `[agent-loop] 快速通道判定: apiType=${apiType} isOllama=${isOllama} nativeTools=${toolMode.nativeTools} dangerous=${topIsDangerous} contentLen=${assistantContent?.length || 0}`);
                    if (isOllama && hasStrongMatch && !topIsDangerous) {
                        ChatLog.warn?.('AGENT', `[agent-loop] 🛡️ Ollama 快速通道触发: 强匹配(${candidates[0].score}) → 跳过重试，直接强制兜底 ${topTool.name}`);
                        forceRetryCount = MAX_FORCE_TOOL_RETRIES;
                    } else {
                        forceRetryCount++;
                        ChatLog.warn?.('AGENT', `[agent-loop] ️ 检测到匹配工具但模型未调用（重试 ${forceRetryCount}/${MAX_FORCE_TOOL_RETRIES}），最佳匹配: ${topTool.name}(${topTool.displayName}) score=${candidates[0].score}`);

                        // 【防闪屏】重试期间不清空 UI，只标记 isRetryRound，下一轮流式输出时不推给 sessionStore
                        // sessionStore.updateTurn(turn.id, { content: '', reasoning: '' });  // 注释掉，避免 UI 闪屏
                        yield { type: 'force_retry', reason: 'tool_match_without_call', round: round + 1 };

                        // 把模型的错误回复加入 messages
                        if (assistantContent) {
                            currentMessages.push({ role: 'assistant', content: assistantContent });
                        }

                        // 构造候选工具建议列表（从工具注册表动态生成，无硬编码）
                        const suggestionLines = candidates.map((c, i) => {
                            const req = c.tool.inputSchema?.required || [];
                            const reqHint = req.length > 0 ? `（需要参数: ${req.join(', ')}）` : '（无需参数）';
                            return `  ${i + 1}. ${c.tool.name} — ${c.tool.displayName || ''}${reqHint}: ${(c.tool.description || '').slice(0, 60)}`;
                        }).join('\n');

                        const topTag = _toolCallTag(topTool) || '工具名';
                        let retryHint;
                        if (_turnHasExecutedTools) {
                            // 【守卫修正】整轮已执行过工具 → 本轮是"总结轮/续做轮"，文案必须与事实一致。
                            // 原分支谎称"你连续多次回复了文字但没有调用任何工具"，会让模型把已完成的
                            // 操作再做一遍（实测 Word：格式化已成功，仍被系统告知未调用）。
                            retryHint = `你刚才只输出了文字，没有继续调用工具。\n\n本任务【已执行过】：${(turn._executedTools || []).join('、')}。\n\n请判断：\n1. 若用户请求还有未完成的步骤，立即只输出 [${topTag}]{"参数":"值"} 标签继续调用工具；\n2. 若已全部完成，直接给出最终总结即可（可以说"已完成"），不要重复执行已完成的步骤。\n\n与用户意图最匹配的工具是：\n${suggestionLines}`;
                        } else if (forceRetryCount >= MAX_FORCE_TOOL_RETRIES) {
                            retryHint = `你连续多次回复了文字但没有调用任何工具！用户说"${userContentStr.slice(0, 50)}"，不调用工具什么都不会发生。\n\n根据你可用的工具列表，与用户意图最匹配的工具是：\n${suggestionLines}\n\n你必须立即输出 [${topTag}] 标签来调用最匹配的工具！不要只回复文字。`;
                        } else if (isFakeExecution) {
                            // 检测到"已调用/已完成"等假执行文字
                            retryHint = `⛔ 严重错误：你刚才回复了"已调用/已完成"等文字，但实际上你并没有输出任何工具调用标签！不调用工具，目标文件/表格什么都不会改变。\n\n用户请求："${userContentStr.slice(0, 60)}"\n\n你必须通过输出 [${topTag}]{"参数":"值"} 标签来真正调用工具，工具才会执行。\n\n与用户意图最匹配的工具是：\n${suggestionLines}\n\n请立即只输出工具调用标签，不要输出任何"已完成"之类的文字！`;
                        } else {
                            retryHint = `你刚才回复了文字但没有调用任何工具。实际修改/生成的操作必须通过输出 [${topTag}] 标签来调用 MCP 工具，文字回复无法真正执行。\n\n以下是与用户意图相关的工具，你应该从中选择最合适的输出 [${topTag}] 标签：\n${suggestionLines}\n\n请立即输出工具标签！`;
                        }
                        currentMessages.push({ role: 'system', content: retryHint });

                        assistantContent = '';
                        assistantReasoning = '';
                        // 【死锁修复】同承诺重试：重试轮必须重新下发 tools，否则模型没工具可调
                        toolsExecuted = false;
                        isRetryRound = true;  // 【防闪屏】标记下一轮为重试轮，不推给 sessionStore
                        continue;
                    }
                }

                // 重试耗尽或无强匹配：判定是否需要强制兜底执行
                // 【门槛修复·死代码】原条件只有 (hasStrongMatch || isFakeExecution)，导致下面注释里写的
                // multiStepSafeMatch（"只要有 userOpIntent + 候选 + 安全工具，不管 tools 有多少，都兜底
                // 执行第一名"）永远不可达：弱匹配（score 5~19）既不满足 hasStrongMatch(>=20)，又不是
                // 假执行文字，于是整个兜底块被跳过 → 最终零工具调用却静默收尾。
                // 这里把"弱但安全"的匹配也纳入入口条件（安全 = 非危险工具、分数过弱匹配线）。
                const _preTopTool = (candidates[0] && candidates[0].tool) || null;
                const _preTopDangerous = !!(_preTopTool && (_preTopTool.isDangerous || _preTopTool.IsDangerous
                    || (_preTopTool.meta && _preTopTool.meta.isDangerous)));
                const weakSafeMatch = userOpIntent && candidates.length > 0 && !!_preTopTool
                    && !_preTopDangerous && (Number(candidates[0].score) || 0) >= WEAK_MATCH_GUIDE_THRESHOLD;
                if ((hasStrongMatch || isFakeExecution || weakSafeMatch) && userOpIntent && forceRetryCount >= MAX_FORCE_TOOL_RETRIES) {
                    ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 模型${MAX_FORCE_TOOL_RETRIES}次重试后仍未调用工具，放弃引导。最佳匹配: ${candidates[0]?.tool?.name || '无'} fakeExecution=${isFakeExecution}`);

                    // ── 光说不做终极修复：强制工具调用兜底 ──
                    // 当模型反复假执行（说"已完成"但不调工具）时，不再用"未执行"文字打发用户，
                    // 而是直接合成 toolCall 强制执行——从"引导模型"转向"保障执行"。
                    //
                    // 触发条件（满足任一）：
                    //   ① hasDeterministicOverrides + (singleToolMode | singleCandidate)
                    //      调用方已确定性解析出参数 + 只有一个可用工具或候选（旧路径，保留）
                    //   ② safeSingleMatch
                    //      hasStrongMatch + 非危险工具（旧路径，保留）
                    //   ③ multiStepSafeMatch   ← 新增：多步骤场景通用兜底
                    //      有 userOpIntent + candidates 非空 + 第一名安全
                    //      多工具场景（tools.length>1）LLM 反复卡壳时，兜底执行最匹配的那个工具，
                    //      让框架继续推进，避免死循环在"下一步我将..."上
                    const hasDeterministicOverrides = paramOverrides && typeof paramOverrides === 'object' && Object.keys(paramOverrides).length > 0;
                    const singleToolMode = Array.isArray(tools) && tools.length === 1;
                    const singleCandidate = candidates.length === 1;
                    const topTool = (candidates[0] && candidates[0].tool) || null;
                    const topIsDangerous = !!(topTool && (topTool.isDangerous || topTool.IsDangerous || (topTool.meta && topTool.meta.isDangerous)));
                    const safeSingleMatch = hasStrongMatch && !!topTool && !topIsDangerous;
                    // 【多步骤通用兜底】只要有 userOpIntent + 候选 + 安全工具，不管 tools 有多少，都兜底执行第一名
                    const multiStepSafeMatch = userOpIntent && candidates.length > 0 && !!topTool && !topIsDangerous;
                    const canForceInvoke = (hasDeterministicOverrides && (singleToolMode || singleCandidate)) || safeSingleMatch || multiStepSafeMatch;

                    if (canForceInvoke) {
                        // singleToolMode 旧路径用 tools[0]，其余用 candidates[0].tool（多步骤场景通用）
                        // 【优先补覆盖度缺项】若自检已点名还没做的工具（如美化后缺 excel_autofit），
                        // 兜底应执行缺项而非再调一次已做过的 format_range。
                        let forcedTool = singleToolMode ? tools[0] : topTool;
                        if (!singleToolMode && Array.isArray(pendingUncovered) && pendingUncovered.length > 0) {
                            const missBare = _bareToolName(pendingUncovered[0].name);
                            const fromTools = (tools || []).find(t => _bareToolName(t && t.name) === missBare);
                            const fromCand = (candidates || []).find(c => _bareToolName(c && c.tool && c.tool.name) === missBare);
                            if (fromTools || (fromCand && fromCand.tool)) forcedTool = fromTools || fromCand.tool;
                        }
                        // 【优先模型点名工具】模型自己点名要调的工具 > 静态词表打分第一名——
                        // 模型比关键词打分更懂上下文（实测：导出 CSV 任务兜底选了别的工具）。
                        // 点名工具从 tools/candidates/注册表三处解析；危险工具不自动执行。
                        if (!singleToolMode && namedCommitment
                            && _bareToolName(forcedTool && forcedTool.name) !== namedCommitment.name) {
                            const _resolveNamed = (n) => {
                                const fromTools = (tools || []).find(t => _bareToolName(t && (t.name || t.originalName)) === n);
                                if (fromTools) return fromTools;
                                const fromCand = (candidates || []).find(c => _bareToolName(c && c.tool && c.tool.name) === n);
                                if (fromCand && fromCand.tool) return fromCand.tool;
                                try {
                                    const _reg = App.ChatV2 && App.ChatV2.toolRegistry;
                                    return (_reg && typeof _reg.get === 'function') ? _reg.get(n) : null;
                                } catch (_e) { return null; }
                            };
                            const _namedTool = _resolveNamed(namedCommitment.name);
                            const _namedDangerous = !!(_namedTool && (_namedTool.isDangerous || _namedTool.IsDangerous
                                || (_namedTool.meta && _namedTool.meta.isDangerous)));
                            if (_namedTool && !_namedDangerous) forcedTool = _namedTool;
                        }
                        const forcedToolName = forcedTool?.name;
                        let forcedArgs = {};
                        if (hasDeterministicOverrides) {
                            forcedArgs = { ...paramOverrides };
                        } else {
                            const _candForForced = (candidates || []).find(c => c && c.tool && _bareToolName(c.tool.name) === _bareToolName(forcedToolName))
                                || candidates[0];
                            forcedArgs = { ...((_candForForced && _candForForced.suggestedArgs) || {}) };
                            // 【失败重试·参数回填】强制兜底参数优先复用"本轮失败调用"的完整实参——
                            // 模型上一轮明确传过的参数比词表打分猜的 suggestedArgs 可信得多。
                            // （实测 21:47：excel_create_chart 因瞬时类型错误失败后，下一轮模型
                            // 只输出文字"重新生成图表"，兜底因缺 dataRange 被跳过 → 不再重试 →
                            // 直接收尾，用户看到"错了不会重新规划"。）失败实参覆盖候选猜测；
                            // 若错误是确定性的（如源区间无数值列），桥会返回明确错误文案，
                            // 模型可在下一轮自行修正参数。
                            try {
                                const _failedCall = (Array.isArray(turn && turn._executedToolCalls) ? turn._executedToolCalls : [])
                                    .slice().reverse()
                                    .find(c => c && c.args && _bareToolName(c.name) === _bareToolName(forcedToolName));
                                if (_failedCall) forcedArgs = { ...forcedArgs, ..._failedCall.args };
                            } catch (_e) { /* 回填失败不阻断兜底 */ }
                            // safeSingleMatch：把前端已知的当前选区注入给带 range 参数的选区类工具，
                            // 使兜底执行精确作用于用户所选区域（如清空 A1:H6）。
                            const sel = turn?.context?.selection;
                            const selAddr = (sel && typeof sel.address === 'string' && sel.address) ? sel.address
                                : (sel && typeof sel.range === 'string' && sel.range) ? sel.range : null;
                            let paramNames = [];
                            if (forcedTool && forcedTool.inputSchema && forcedTool.inputSchema.properties) {
                                paramNames = Object.keys(forcedTool.inputSchema.properties);
                            } else if (forcedTool && Array.isArray(forcedTool.inputSchema)) {
                                paramNames = forcedTool.inputSchema.map(p => p && p.name).filter(Boolean);
                            } else if (forcedTool && Array.isArray(forcedTool.parameters)) {
                                paramNames = forcedTool.parameters.map(p => p && p.name).filter(Boolean);
                            }
                            if (paramNames.includes('range') && selAddr && _bareToolName(forcedToolName) !== 'excel_format_range') forcedArgs.range = selAddr;
                        }
                        // chartTypeOverride 补充兜底（excel_create_chart 场景）
                        if (chartTypeOverride && forcedToolName && /excel_create_chart/i.test(forcedToolName) && !forcedArgs.chartType) {
                            forcedArgs.chartType = chartTypeOverride;
                        }
                        // 【美化兜底缺参补全】excel_format_range 必填 format；空 {} 后端走商务蓝预设+AutoFit。
                        // 实测：候选命中 format_range 但 suggestedArgs 只有 range → 下方缺参检查跳过兜底
                        // → 零工具收尾，模型"已成功美化…#FFFFCC"幻觉原样展示（用户截图 23:20）。
                        const _forceBare = _bareToolName(forcedToolName || '');
                        // Word/PPT 表格格式化：仅当用户话含美化/格式语义才允许强制；禁止对「翻译选中内容」等空参硬推
                        // （实测：翻译请求 → 零工具 → 强推 word_format_selected_table_cell {} → 连续失败）。
                        const _hostFmtTool = /^(word|ppt)_format_selected_table_cell$/.test(_forceBare);
                        const _userFmtIntent = /(美化|格式化|样式|边框|底色|底纹|表头|加粗|居中|对齐|主题|配色|隔行|斑马|版式|三线表)/.test(
                            String(rawUserMsg || userContentStr || ''));
                        let _skipHostFmtForce = false;
                        if (_hostFmtTool && !_userFmtIntent && !hasDeterministicOverrides) {
                            _skipHostFmtForce = true;
                            ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 跳过兜底执行 ${forcedToolName}：用户意图非表格格式化（如翻译/润色）`);
                        }
                        if (_forceBare === 'excel_format_range') {
                            if (forcedArgs.format == null || forcedArgs.format === '') forcedArgs.format = {};
                            if (!forcedArgs.range && App.ChatV2.rangeResolver && typeof App.ChatV2.rangeResolver.decideFormatTarget === 'function') {
                                const _fp = App.ChatV2.rangeResolver.decideFormatTarget(String(rawUserMsg || userContentStr || ''), turn && turn.context);
                                if ((_fp.rangeMode === 'selection' || _fp.rangeMode === 'cell' || _fp.rangeMode === 'copy-selection') && _fp.selectionRange) {
                                    forcedArgs.range = _fp.selectionRange;
                                } else if (_fp.rangeMode === 'spoken' && _fp.spokenRange) {
                                    forcedArgs.range = _fp.spokenRange;
                                }
                            }
                        } else if (_forceBare === 'excel_autofit') {
                            if (!forcedArgs.type) forcedArgs.type = 'column';
                            if (!forcedArgs.range) {
                                const _sel = turn?.context?.selection;
                                const _addr = (_sel && (typeof _sel.address === 'string' ? _sel.address : _sel.range)) || '';
                                if (_addr) forcedArgs.range = _addr;
                            }
                        } else if (_hostFmtTool && _userFmtIntent && Object.keys(forcedArgs).length === 0) {
                            // 美化意图但无参：补经典网格默认，避免空 {} 被后端拒成「未指定任何格式项」
                            forcedArgs = {
                                borderStyle: 'all', borderColor: '#4472C4', zebra: true,
                                zebraColor: '#EDF2FA', headerBold: true, headerFillColor: '#D9E2F3',
                                headerFontColor: '#1F3864', align: 'center'
                            };
                        }
                        // 【防止必然失败的兜底调用】自动兜底（非调用方确定性传参）时，若目标工具的
                        // 必填参数在有效参数中缺失/为空，则不再合成调用——否则每轮都会产生一张
                        // "0ms 失败"卡片（如 excel_write_range 缺 values、excel_create_chart 缺 dataRange），
                        // 并在"重试→兜底→失败→重试"之间形成死循环。
                        // 空对象 {} 对 format 合法（触发后端预设），不得当缺失。
                        const _autoForce = !hasDeterministicOverrides;
                        const _effArgs = { ...forcedArgs, ...(paramOverrides && typeof paramOverrides === 'object' ? paramOverrides : {}) };
                        const _reqParams = Array.isArray(forcedTool?.inputSchema?.required) ? forcedTool.inputSchema.required : [];
                        const _isEmptyArg = (v) => v === undefined || v === null || v === ''
                            || (Array.isArray(v) && v.length === 0);
                        // 【有安全默认 ≠ 缺失】注册表里登记了安全默认生成器的参数（如
                        // excel_export_csv.filePath 自动落桌面）缺了也能被 invokeTracked 自动补全，
                        // 不应因缺参放弃兜底——否则"导出 CSV"永远到不了执行那一步（实测 19:34）。
                        const _hasSafeDefault = (() => {
                            try {
                                const _reg = App.ChatV2 && App.ChatV2.toolRegistry;
                                return (_reg && typeof _reg.hasSafeDefault === 'function') ? _reg.hasSafeDefault : null;
                            } catch (_e) { return null; }
                        })();
                        const _missingReq = _autoForce ? _reqParams.filter(p => {
                            const v = _effArgs[p];
                            if (p === 'format' && v && typeof v === 'object' && !Array.isArray(v)) return false;
                            if (_isEmptyArg(v) && _hasSafeDefault && _hasSafeDefault(_forceBare, p)) return false;
                            return _isEmptyArg(v);
                        }) : [];
                        if (forcedToolName && (topIsDangerous || _missingReq.length > 0 || _skipHostFmtForce)) {
                            ChatLog.warn?.('AGENT', `[agent-loop] ⚠️ 跳过兜底执行 ${forcedToolName}：` +
                                (topIsDangerous ? '危险工具不自动执行'
                                    : (_skipHostFmtForce ? '非表格格式化意图'
                                        : `缺少必填参数 [${_missingReq.join(', ')}]`)));
                        } else if (forcedToolName) {
                            ChatLog.warn?.('AGENT', `[agent-loop] 🛡️ 强制执行兜底：合成 toolCall ${forcedToolName} args=${JSON.stringify(forcedArgs).slice(0, 200)}`);
                            try { sessionStore?.appendStepLog?.(turnId, 'tool_call', `【兜底执行】强制调用 ${forcedToolName} 参数：${JSON.stringify(forcedArgs).slice(0, 150)}`, stepNo); } catch (e) { /* ignore */ }

                            const forcedCallId = `call_force_${Date.now()}_${Math.random().toString(36).slice(2, 7)}`;
                            // 合成假 tool_calls 数组，交由后续正常工具执行流程（含卡片渲染/paramOverrides再覆盖/日志等）
                            // 注意：后续 1251 行有 paramOverrides 二次覆盖，这里只传 args 基线即可
                            // 元技能补调（skill_xxx）：模型自由发挥轮没写参数，用用户原话兜底，
                            // 确保 SkillCreator 拿到明确输入（而非空 {}，仅靠 buildSkillInputs 的
                            // userMessage 兜底才不丢）。userText 在上方 2339 行已取（最后一条 user 消息）。
                            if (/^skill_/i.test(forcedToolName)) {
                                forcedArgs = Object.assign({ prompt: userText || '', text: userText || '' }, forcedArgs);
                            }
                            const forcedTc = [{ id: forcedCallId, name: forcedToolName, args: forcedArgs }];
                            toolCalls = forcedTc;
                            // 清除假执行文字，避免与真实执行结果混淆
                            assistantContent = '';
                            assistantReasoning = '';
                            // 无条件清空本轮「假执行」文字与正文：原逻辑依赖
                            // getCurrentTurn()?.id === turn.id，多步骤/历史重载等场景下
                            // getCurrentTurn 可能指向别处 → 判断失败 → 模型自由发挥的废文本
                            // （如 agent-loop 首轮生成的无关 VBA/说明）残留在正文里，
                            // 用户看到"先生成了 VBA，但 Skill 正文没调用它"。
                            // turn.id 即本循环正处理的 turn，直接更新即可；turn 已取消时
                            // updateTurn 抛错由 try 吞掉，不影响其他流程。
                            try { sessionStore.updateTurn(turn.id, { content: '', reasoning: '' }); } catch (e) { /* turn 已取消则忽略 */ }
                            // 不再走 return/loop_done；跳出 if，让后续正常的 for (const tc of toolCalls) 工具执行流程接管。
                            // 【勿 slice 历史】旧逻辑 currentMessages.slice(0, 2) 会清掉已成功的 tool 结果
                            // 与多轮上下文，恰好在模型已卡壳时毁掉多步记忆；假执行正文已在上方清空。
                            hasStrongMatch = false;  // 防止误触发后续分支
                            isFakeExecution = false;
                        }
                    }

                    // 兜底失败（无可强制工具）：降级为原"未执行"文字提示。
                    // 【守卫修正】整轮已成功执行过工具时必须整段跳过：本轮文字即执行后的最终总结，
                    // 替换/追加"实际没有调用任何工具"会反向误导用户（实测 Word 08:56：模型已成功
                    // 调用 word_format_selected_table_cell 并返回 success:true，总结却被整条替换成
                    // "实际没有调用任何工具，Excel 未被修改"）。
                    if (toolCalls.length === 0 && !_turnHasExecutedTools) {
                        if (isFakeExecution) {
                            const replacementMsg = `⚠️ 抱歉，刚才的回复声称已执行操作，但**实际没有调用任何工具，${_docLabel}**。

您请求的是："${String(rawUserMsg || '').slice(0, 80)}"

${candidates.length > 0 ? `与您意图最匹配的可用操作是 **${candidates[0].tool.displayName || candidates[0].tool.name}**。` : '当前没有匹配到可执行的操作工具。'}

请重新描述您的需求（例如"执行XX"、"调用工具"），或确认您选择的区域/参数是否正确。`;
                            if (sessionStore.getCurrentTurn()?.id === turn.id) {
                                sessionStore.updateTurn(turn.id, { content: replacementMsg });
                            }
                            assistantContent = replacementMsg;
                            ChatLog.warn?.('AGENT', '[agent-loop] 已用未执行提示替换假执行文字');
                            fakeTextHandled = true;
                        } else if (isUserGuidance) {
                            // 【守卫修正】引导型回复（如"未选中具体对象，请先选择表格"）：
                            // 模型本意是指导用户操作，并非声称自己已完成。保留其原始引导文字，
                            // 不整段替换、也不追加"未执行"误导警告（避免反向误导用户）。
                            fakeTextHandled = true;
                        } else {
                            // 非假执行但匹配失败：保留原文字并追加警告
                            const warnMsg = `\n\n⚠️ 注意：以上回复声称已执行操作，但实际未调用任何工具，${_docLabel}。如需执行该操作，请明确说"执行/调用工具"或重新描述需求。`;
                            if (sessionStore.getCurrentTurn()?.id === turn.id) {
                                sessionStore.appendDelta(turn.id, 'content', warnMsg);
                            }
                            fakeTextHandled = true;
                        }
                    }
                }
                // 兜底合成了 toolCalls 后（强制调用路径），toolCalls 已经非空，
                // 不能再走 no_tool_calls return；只有 toolCalls 仍为空时才按原逻辑结束
                if (toolCalls.length === 0) {
                    // 只读了表，正文却写「已标红 / 已涂成红色」。read_range 不算改色。
                    if (_turnHasExecutedTools && assistantContent
                        && /已标[红黄蓝绿橙紫粉灰青]|涂成|醒目红|已标记|标成|换成.{0,4}色/.test(assistantContent)) {
                        const _ran = (turn && Array.isArray(turn._executedTools)) ? turn._executedTools.map(String) : [];
                        const _painted = _ran.some(n => /format|conditional|clean_data|write_range|set_cell/.test(n));
                        if (!_painted) {
                            const _ask = String(userContentStr || '').slice(0, 40);
                            assistantContent = '⚠️ 只读取了数据，单元格颜色没有改。'
                                + (_ask ? `「${_ask}」没有作用到工作表上。` : '工作表没有被修改。');
                            try { sessionStore.updateTurn(turn.id, { content: assistantContent }); } catch (_e) { /* ignore */ }
                            fakeTextHandled = true;
                            ChatLog.warn?.('AGENT', '[agent-loop] 只读却声称已标色 → 已改成诚实说明');
                        }
                    }
                    // 【内容任务·无素材拦截】总结/润色/翻译类任务，但上下文里没有注入任何文本、
                    // 整轮也零工具调用 → 模型输出必然是凭标题编造（"项目A: 已完成30%…"）。
                    // 只追加"未执行"小字警告不够（用户看到的仍是满屏假数据），必须用诚实结论
                    // 替换掉编造正文，直接向用户报"读不到内容"（= 你要的"失败要报"）。
                    if (_contentTaskNoContent && assistantContent && String(assistantContent).trim()) {
                        const _h = (typeof window !== 'undefined' && window.SmartPilotHost) || null;
                        const _isPpt = !!(_h && _h.isPowerPoint);
                        const _what = _isPpt ? '当前幻灯片' : '当前文档';
                        const _where = _isPpt ? '幻灯片' : '文档';
                        const noContentMsg = `未能读取到${_what}的文字内容，因此无法完成该处理。`
                            + `（当前${_where}没有可读取的文字，或内容尚未同步。）\n\n`
                            + `请确认${_what}中已有文字后重试，或直接在消息里粘贴需要处理的文本。`;
                        if (sessionStore.getCurrentTurn()?.id === turn.id) {
                            sessionStore.updateTurn(turn.id, { content: noContentMsg });
                        }
                        assistantContent = noContentMsg;
                        fakeTextHandled = true;
                        ChatLog.warn?.('AGENT', '[agent-loop] 内容任务但上下文无任何文本且零工具调用 → 已用"读不到内容"替换编造正文');
                        try { _hostLog('WARN', '[agent-loop] 内容任务无素材 → 已替换幻觉正文为诚实提示'); } catch (_e) { /* 日志失败不影响判定 */ }
                    }
                    // 【P1 修复】放行口假执行清洗：弱匹配（score 5-9，不足 isFakeExecution 的 10 分线）
                    // + 假执行文字 + 重试预算耗尽时，文本会从这里原样放行给用户。
                    // 此处统一追加诚实警告（保留模型原文，不替换，score<5 的纯咨询回答不受影响）。
                    if (!fakeTextHandled && !_contentTaskExempt && !_turnHasExecutedTools && assistantContent && topCandidate
                        && topCandidate.score >= WEAK_MATCH_GUIDE_THRESHOLD
                        && FAKE_EXECUTION_PATTERN.test(assistantContent) && !isUserGuidance) {
                        const fakeWarn = `\n\n⚠️ 注意：以上回复声称已执行操作，但本次会话未能调用任何工具，${_docLabel}。请点击重试，或换个说法明确要求"执行"该操作。`;
                        if (sessionStore.getCurrentTurn()?.id === turn.id) {
                            sessionStore.appendDelta(turn.id, 'content', fakeWarn);
                        }
                        assistantContent += fakeWarn;
                        ChatLog.warn?.('AGENT', '[agent-loop] 放行口检测到假执行文字（弱匹配+预算耗尽），已追加未执行警告');
                    }
                    // 【诚实收尾·绝对兜底】用户明确有 Excel 操作意图、但整轮【一个工具都没执行】时，
                    // 无论候选分数高低、模型是否点名工具，都不能让"操作完成 / 结果摘要：改了…"这类
                    // 幻觉顶着"任务完成 · 未调用工具"过去。
                    // 为什么需要这条：上面两个分支都带门槛——FAKE_EXECUTION_PATTERN 需"已X"、
                    // namedCommitment 需模型点名真实工具名、弱匹配需 score≥5。
                    // 实测模型写"操作完成。结果摘要：改了表头…调整了列宽以适应内容。"三者全不命中，
                    // 于是零工具调用却静默显示成功（用户看到的正是这张图）。
                    // 【2026-10-10】若正文已是"已成功美化…背景颜色 #FFFFCC"完成腔，只追加小字警告
                    // 不够（用户仍以为做完了）；改为整段替换为诚实失败说明。
                    if (!fakeTextHandled && !_contentTaskExempt && assistantContent && userOpIntent
                        && !_turnHasExecutedTools && !(turn && turn._attachmentAnswerOnly)) {
                        const _looksDone = FAKE_EXECUTION_PATTERN.test(assistantContent)
                            || /背景颜色|底色|#(?:FF|ff)[0-9A-Fa-f]{4}|Border Style/i.test(assistantContent);
                        if (_looksDone && !isUserGuidance) {
                            const replacementMsg = `⚠️ 刚才的回复声称已完成美化/格式化，但**实际没有调用任何工具，${_docLabel}**。\n\n`
                                + `您请求的是："${String(rawUserMsg || '').slice(0, 80)}"\n\n`
                                + `${candidates.length > 0 ? `与意图最匹配的工具是 **${candidates[0].tool.displayName || candidates[0].tool.name}**，请点击重试让我真正执行。` : '请点击重试，或明确说"执行/调用工具"。'}`;
                            try { sessionStore.updateTurn(turn.id, { content: replacementMsg }); } catch (_e) { /* ignore */ }
                            assistantContent = replacementMsg;
                            ChatLog.warn?.('AGENT', '[agent-loop] 零工具+完成腔幻觉 → 已整段替换为诚实失败说明');
                        } else {
                            const neverRanWarn = `\n\n⚠️ 注意：以上内容只是计划/描述，**本次没有调用任何工具，${_docLabel}**。`
                                + `请点击重试，或明确说"执行/调用工具"。`;
                            try { sessionStore.updateTurn(turn.id, { content: assistantContent + neverRanWarn }); } catch (_e) {
                                try { sessionStore.appendDelta(turn.id, 'content', neverRanWarn); } catch (_e2) { /* ignore */ }
                            }
                            assistantContent += neverRanWarn;
                            try {
                                _hostLog('WARN', `[agent-loop] 零工具执行（用户有操作意图）→ 已追加"未执行"诚实告警`
                                    + ` candidates=${candidates.length}`
                                    + ` top=${candidates[0] && candidates[0].tool ? candidates[0].tool.name : '无'}`
                                    + ` score=${candidates[0] ? candidates[0].score : 0}`);
                            } catch (_e) { /* 日志失败不影响判定 */ }
                        }
                        fakeTextHandled = true;
                    }
                    // 【承诺未执行·诚实兜底】重试预算耗尽后仍只有"我将要…"的计划文字时，
                    // 必须明确告诉用户"什么都没执行"，不能让它顶着"任务完成 · 未调用工具"过去。
                    // 与上面假执行清洗的区别：那是过去式（"已完成"），这是未来式（"我将要"）。
                    // 注意：`!fakeTextHandled` —— 若上方"内容任务·无素材"已用诚实提示替换正文，
                    // 这里不再追加"执行计划"警告，避免两段提示叠在一起。
                    if (!fakeTextHandled && namedCommitment && !_turnHasExecutedTools) {
                        const commitWarn = `\n\n⚠️ 注意：以上是执行计划，**实际没有调用任何工具，${_docLabel}**。`
                            + `请点击重试，或直接说"执行/调用工具"。`;
                        if (sessionStore.getCurrentTurn()?.id === turn.id) {
                            sessionStore.appendDelta(turn.id, 'content', commitWarn);
                        }
                        assistantContent += commitWarn;
                        ChatLog.warn?.('AGENT', `[agent-loop] 放行口检测到"承诺未执行"（${namedCommitment.name}，预算耗尽），已追加未执行警告`);
                    }
                    // 【防闪屏】重试耗尽前，把本地累积的内容写入 sessionStore
                    // 多步骤通道(routeToThinking)下正文本就留空，重试轮内容已在 thinking 块，不回灌正文
                    if (isRetryRound && assistantContent && !routeToThinking) {
                        sessionStore.updateTurn(turn.id, { content: assistantContent });
                    }
                    if (isRetryRound && assistantReasoning && !routeToThinking) {
                        sessionStore.updateTurn(turn.id, { reasoning: assistantReasoning });
                    }
                    closeReason = 'no_tool_calls';
                    // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: withCoverageNote(assistantContent) });
                    // 多步骤通道：本轮是"最终总结轮"（无工具调用收尾），其叙述此前已被流式写入
                    // 本轮 thinking 块；此处 move 回正文，避免同一段总结同时出现在思考栏与正文。
                    //
                    // 【不要留空】此前直接置 '' → 渲染成"（无推理文字）"，而这段字是用户
                    // "盯着思考栏一个字一个字蹦出来"的，收尾瞬间凭空消失（实测反馈：
                    // "开始我看到正在思考中有内容，后来有消失了"）。正文已拿到同一段文字，
                    // 这里只留一行去向说明：既不重复展示，也不制造"内容丢了"的错觉。
                    if (routeToThinking && currentThought) {
                        currentThought.content = '（本段叙述已归入下方正文）';
                        // 【完成态措辞】不回传 running 态标题（"正在分析你的需求…"），否则完成 ✓ 后仍显示"正在..."
                        const _doneTitle = (currentThought.title && currentThought.title.indexOf('正在') >= 0)
                            ? '🔍 思考过程' : currentThought.title;
                        try { sessionStore.completeThought(turn.id, _doneTitle, { status: 'completed' }); } catch (_e) { /* 思考块已收尾，忽略 */ }
                    }
                }
                yield { type: 'loop_done', reason: 'no_tool_calls', round: round + 1 };
                    return;
                }
                // 兜底强制调用分支：toolCalls 已合成，跳出父 if(toolCalls.length===0)，进入后续正常工具执行流程
                // 重置 isRetryRound，确保工具结果能正常写入 sessionStore
                isRetryRound = false;
            }

            // 【防闪屏】如果本轮是重试轮且成功了（有工具调用或正常结束），把本地累积的内容一次性写入 sessionStore
            // 多步骤通道(routeToThinking)下正文留空，重试轮内容已在 thinking 块，不回灌正文
            if (isRetryRound && assistantContent && !routeToThinking) {
                sessionStore.updateTurn(turn.id, { content: assistantContent });
                isRetryRound = false;  // 重置，后续轮恢复正常流式推送
            }
            if (isRetryRound && assistantReasoning && !routeToThinking) {
                sessionStore.updateTurn(turn.id, { reasoning: assistantReasoning });
            }

            // 【通用框架·工具名别名归一】在执行前统一归一，覆盖原生 tool_calls、文本标记、
            // 兜底合成三条来源。模型或"强制执行兜底"可能产出泛化名（如 chart），
            // 在此归一到真实工具，避免后端"工具名翻译"误译到跨族工具（如 chart → ppt_add_chart_slide）。
            for (const tc of toolCalls) {
                if (tc && tc.name) {
                    const resolved = aliasToolName(tc.name);
                    if (resolved !== tc.name) {
                        ChatLog.info?.('AGENT', `[agent-loop] 工具名别名归一: ${tc.name} → ${resolved}`);
                        tc.name = resolved;
                    }
                }
            }

            // 【通用框架·工具调用去重】在执行前统一去重（覆盖原生 tool_calls、文本标记、
            // 兜底合成三条来源），避免同名同参调用被重复执行（如 ppt_generate 连生成多份）。
            toolCalls = dedupeToolCalls(toolCalls);

            // 把 assistant 的工具调用加入 messages（OpenAI function calling 格式）
            currentMessages.push({
                role: 'assistant',
                content: assistantContent || null,
                tool_calls: toolCalls.map(tc => ({
                    id: tc.id,
                    type: 'function',
                    function: {
                        name: tc.name,
                        arguments: typeof tc.args === 'string' ? tc.args : JSON.stringify(tc.args || {})
                    }
                }))
            });

            // 附件正文已经写好时，模型若再去调 Excel 工具，会先清掉这段分析，再按工具失败重写。
            // 图纸尺寸这类问题没有要求改表，保留已经写出的分析。
            if (turn && turn._attachmentAnswerOnly && toolCalls.length > 0
                && toolCalls.every(function (tc) { return /^excel_/i.test(String((tc && tc.name) || '')); })) {
                ChatLog.info?.('AGENT', '[agent-loop] 附件分析已写好，忽略后续 Excel 工具，保留原文');
                if (assistantContent && String(assistantContent).trim()) {
                    try { sessionStore.updateTurn(turn.id, { content: assistantContent }); } catch (_keepErr) { /* 忽略 */ }
                }
                closeReason = 'attachment_answer_kept';
                yield { type: 'loop_done', reason: closeReason, round: round + 1 };
                return;
            }

            // 【先答后做治理】模型常在工具调用前先流式输出"已成功美化/已生成…"等假完成总结，
            // 文本实时上屏（见上方 token 事件处理），随后才真正执行工具 → 用户第一时间看到假结论。
            // 本轮最终确定要执行工具 → 执行前清掉这批前置正文；真实总结由工具执行完的下一轮流式产出。
            // 纯文本回答（无工具调用）不受影响，保持打字机效果；重试轮/静默轮本就未推 UI。
            // 注意：只清 UI 展示（turn.content），assistantContent 变量保留，下一轮 LLM 上下文不受影响。
            if (toolCalls.length > 0 && !isRetryRound && !silentContent && assistantContent) {
                try {
                    sessionStore.updateTurn(turn.id, { content: '' });
                    ChatLog.info?.('AGENT', `[agent-loop] 已清除工具执行前的模型前置正文（${assistantContent.length} 字），防"先答后做"假总结`);
                } catch (e) { /* 忽略 */ }
            }

            // 执行每个工具调用
            const localResults = [];  // 收集本轮所有工具执行结果（供 artifact 提取用）
            let _needAutoRecognize = false;
            // 非空 tools 即本轮白名单（Skill allowedTools / 管道过滤后的列表）。
            // 空数组保持旧契约：不限制，测试与「未下发工具」路径仍可执行已注册工具。
            const _offeredToolNames = (Array.isArray(tools) && tools.length > 0)
                ? new Set(tools.map(t => String((t && (t.name || t.originalName)) || '').trim().toLowerCase()).filter(Boolean))
                : null;
            for (const tc of toolCalls) {
                if (signal?.aborted) {
                    closeReason = 'user_aborted';
                    yield { type: 'canceled', reason: 'user_aborted' };
                    return;
                }

                // 【致命修复】过滤 LLM 幻觉输出的空名/空 id tool_call
                // 之前不跳过 → 后面 registeredTool=null → "Tool not found" → 当 failed 存进 messages
                // → 下一轮发给 LLM API → HTTP 400 "tool call id :N is duplicated" → 整个 turn 崩溃！
                // 典型症状：LLM round 1 输出 24 个 tool_calls，23 个 name="" 1 个正常
                if (!tc.name || tc.name.trim() === '' || !tc.id) {
                    ChatLog.warn?.('AGENT', `[agent-loop] 跳过无效 tool_call: name="${tc.name || ''}" id="${tc.id || ''}" (LLM 幻觉输出)`);
                    continue;
                }

                if (_offeredToolNames && !_offeredToolNames.has(String(tc.name).toLowerCase())) {
                    const deny = `工具 ${tc.name} 不在本次允许的工具列表中，已拒绝执行。请只调用提供的工具。`;
                    ChatLog.warn?.('AGENT', `[agent-loop] 白名单拒绝: ${tc.name}`);
                    anyToolFailed = true;
                    emitAgent('tool/result', { turnId, step: stepNo, toolName: tc.name, callId: tc.id, ok: false, error: deny, durationMs: 0, toolCallId: null });
                    yield { type: 'tool_call_result', toolCallId: null, toolName: tc.name, ok: false, result: null, error: deny, durationMs: 0 };
                    currentMessages.push({ role: 'tool', tool_call_id: tc.id, name: tc.name, content: deny });
                    localResults.push({ ok: false, error: deny, durationMs: 0 });
                    continue;
                }

                let result;
                const currentRound = round + 1;

                // 提前取出注册表中的工具定义，供后续未注册工具判定及跳过分支复用。
                const registeredTool = toolRegistry.get(tc.name);
                // 提前计算幂等键，供下方 else 内拦截及循环末尾台账记录复用。
                const _idemKey = toolCallKey(tc);

                // 【图表单请求单实例】单图请求（非显式多图表/看板）一个 turn 内只允许一次【成功】的 excel_create_chart。
                // 关键：用 turn._chartDone（仅成功时置位，跨轮持久）而非"是否被调用过"，
                // 否则 LLM 调图但失败时，recordToolExecutions 仍会记名 → fallback 误判"已建图" → 出图数为 0。
                // 注：_forceMultiStep 在普通单图请求里也会被守卫 D 置 true，不能作为"允许多图"信号；
                // 只有显式看板/仪表盘意图（_forceDashboard）才放行多图。
                const _isSingleChartBlock = tc.name === 'excel_create_chart'
                    && !turn.context?._forceDashboard
                    && !!turn._chartDone;
                if (_isSingleChartBlock) {
                    ChatLog.warn?.('AGENT', `[agent-loop] 跳过重复图表创建：单图请求本 turn 已建过图 ${tc.name}`);
                    result = { ok: true, skipped: true, result: '已跳过：同一请求内已创建图表' };
                } else {
                // 【看板/多图放行】显式看板/仪表盘意图（_forceDashboard）生成的多个图表需保留，
                // 向后端传 multiChart=true 关闭默认的单图自清理。
                // 普通单图请求无需注入：后端 m_excel_create_chart 对 non-multiChart 建图前会自动清理历史 AI 图。
                if (tc.name === 'excel_create_chart' && turn.context?._forceDashboard) {
                    tc.args = { ...(tc.args || {}), multiChart: true };
                }
                // 在缓冲渲染卡片之前就应用 paramOverrides / chartTypeOverride，
                // 使卡片顶部图表类型标签（来自 args.chartType）显示用户真实意图，而非默认的柱状图。
                // 原覆盖逻辑在 invokeTracked 前执行，但卡片早已用未覆盖 args 渲染，导致标签恒为"柱状图"。
                if (paramOverrides && typeof paramOverrides === 'object' && Object.keys(paramOverrides).length > 0) {
                    tc.args = { ...(tc.args || {}), ...paramOverrides };
                } else if (chartTypeOverride && (tc.name === 'excel_create_chart' || tc.name === 'excel_create_chart_bar')) {
                    tc.args = { ...(tc.args || {}), chartType: chartTypeOverride };
                }

                // ── 【通用框架·重复调用拦截（turn 级 + 同轮级）】────────────────
                // 终态/重型工具额外走"按 turn 隔离"的台账：仅同一 turn 内的多轮复读才拦住复用，
                // 避免重复生成文件、连弹相同卡片；跨 turn（新消息）自然放行真实重做。
                // 仅终态工具启用此台账；读类工具可放心重复。
                // 命中幂等结果（会话级 / 同轮级）且非用户明确"再生/重做"意图 → 复用，不重跑；
                // 用户说"重新/再来/重新生成"等 → 绕过拦截，放行真正的重做。
                if (turn && _idemKey) {
                    if (!(turn._toolCallLedger instanceof Map)) turn._toolCallLedger = new Map();
                    const _prevRun = turn._toolCallLedger.get(_idemKey);
                    const _isTerminal = isTerminalTool(tc.name, toolRegistry);
                    const _prevSession = _isTerminal ? sessionLedgerGet(turnId, _idemKey) : null;
                    const _reprobe = REPROBE_RE.test(latestUserMsgText(currentMessages));
                    const _cached = (_prevSession && _prevSession.ok) ? _prevSession
                        : (_prevRun && _prevRun.ok) ? _prevRun : null;
                    if (_cached) {
                        if (_reprobe) {
                            ChatLog.info?.('AGENT', `[agent-loop] 🔄 用户要求再生/重做 → 跳过幂等拦截，重新执行：${tc.name}`);
                        } else {
                            ChatLog.warn?.('AGENT', `[agent-loop] ♻️ 拦截重复调用（${_prevSession && _prevSession.ok ? '会话级' : '同轮级'}，参数一致且已成功）：${tc.name}`);
                            try { sessionStore?.appendStepLog?.(turnId, 'warning', `工具 ${tc.name} 与之前调用参数完全一致且已成功执行，跳过重复执行`, stepNo); } catch (e) { /* 忽略 */ }
                            emitAgent('tool/result', { turnId, step: stepNo, toolName: tc.name, callId: tc.id, ok: true, error: null, durationMs: 0, toolCallId: null });
                            yield { type: 'tool_call_result', toolCallId: null, toolName: tc.name, ok: true, result: _cached.result, error: null, durationMs: 0 };
                            currentMessages.push({
                                role: 'tool',
                                tool_call_id: tc.id,
                                name: tc.name,
                                content: '（该调用与之前完全一致，已执行成功，结果同上，本次未重复执行。请勿重复调用同一工具，直接给出最终总结。）'
                            });
                            localResults.push({ ok: true, result: _cached.result, durationMs: 0 });
                            continue;
                        }
                    }

                    // 【通用框架·失败熔断】同一"工具名+参数"已连续失败 N 次 → 不再执行。
                    // 与上面的 ♻️ 成功幂等拦截互补：成功台账只记成功，失败调用此前完全无记录，
                    // 于是弱模型会把同一个错误一路重放到 maxRounds——实测 2026-10-05 22:44 的
                    // ppt_export_pdf 连失败 5 次（后端缺参校验），UI 弹出 5 张一样的失败卡片，
                    // 模型最后还照着自己的想象编了个文件名告诉用户"已导出"。
                    // 熔断后回喂明确指令：不要再重复同一调用，改用其他方式或如实说明失败原因。
                    if (!turn._failStreak) turn._failStreak = new Map();
                    const _fb = (App.ChatV2.agentLoopToolExecution
                        && typeof App.ChatV2.agentLoopToolExecution.checkFailBreaker === 'function')
                        ? App.ChatV2.agentLoopToolExecution.checkFailBreaker(
                            turn._failStreak, _idemKey, _reprobe, SAME_CALL_FAIL_LIMIT)
                        : { break: false, failN: 0, lastErr: '' };
                    const _failN = _fb.failN;
                    if (_fb.break) {
                        const _lastErr = _fb.lastErr;
                        ChatLog.warn?.('AGENT', `[agent-loop] 🛑 失败熔断：${tc.name} 同参数已连续失败 ${_failN} 次，跳过重复执行`);
                        try { sessionStore?.appendStepLog?.(turnId, 'warning', `工具 ${tc.name} 同参数已连续失败 ${_failN} 次，已阻止重复执行`, stepNo); } catch (e) { /* 忽略 */ }
                        emitAgent('tool/result', { turnId, step: stepNo, toolName: tc.name, callId: tc.id, ok: false, error: '重复失败已熔断', durationMs: 0, toolCallId: null });
                        yield { type: 'tool_call_result', toolCallId: null, toolName: tc.name, ok: false, result: null, error: `重复失败已熔断（${_failN} 次）：${_lastErr}`, durationMs: 0 };
                        currentMessages.push({
                            role: 'tool',
                            tool_call_id: tc.id,
                            name: tc.name,
                            content: `（该调用已连续失败 ${_failN} 次，最后一次错误：${_lastErr}。系统已停止重复执行。`
                                + '请不要再用相同参数调用它；可以：①改用其他工具或补齐/更换参数；'
                                + '②如认为这是环境问题，如实告知用户失败原因与建议。'
                                + '严禁在未成功的情况下声称"已导出/已保存/已完成"。）'
                        });
                        localResults.push({ ok: false, error: `重复失败已熔断（${_failN} 次）：${_lastErr}`, durationMs: 0 });
                        // 熔断也是失败：必须置 anyToolFailed，否则本轮若无真实 invoke，
                        // 下方会把 toolsExecuted=true → 下一轮非 persistTools 路径拿不到 tools。
                        anyToolFailed = true;
                        continue;
                    }
                }

                // 【无选区降级】用户没选任何单元格（或选区只是光标单格）→ selection 为 null/空/单格，
                // LLM 大概率瞎编 range / sheet 参数。兜底策略：
                // 1. 先用 turn.context.selection（可能是 context-engine 自动捕获的 UsedRange；
                //    单格一律当"光标停留"弃用，与 dispatcher 的 _selIsRange 同一语义）
                // 2. 如果也没有，实时调 GetSelection/GetUsedRangeData 拿真实范围
                // 3. 用真实范围覆盖 LLM 瞎编的 range 参数（但保留 sheet 名替换逻辑）
                const _sel = turn?.context?.selection;
                const _hasSelRange = !!(
                    _sel && (
                        (_sel.range && String(_sel.range).length > 0) ||
                        (_sel.address && String(_sel.address).length > 0)
                    )
                );
                const _excelTool = /^excel_(read_range|aggregate|create_chart|clean_data|write_range|copy_range|sort_range|filter_range|create_pivot|format_range|extract|split|merge|remove_duplicates|create_chart_bar|create_chart_line|create_chart_pie)/.test(tc.name);
                const _rangeParamMap = { excel_read_range: 'range', excel_aggregate: 'range', excel_create_chart: 'dataRange', excel_create_chart_bar: 'dataRange', excel_create_chart_line: 'dataRange', excel_create_chart_pie: 'dataRange', excel_clean_data: 'range', excel_write_range: 'range', excel_copy_range: 'sourceRange', excel_sort_range: 'range', excel_filter_range: 'range', excel_create_pivot: 'range', excel_format_range: 'range' };
                const _rangeField = _excelTool ? (_rangeParamMap[tc.name] || 'range') : null;

                // 【瞎编 range 拦截】当 selection 为空/是单格时，用 GetSelection 或 GetUsedRangeData 兜底
                if (_excelTool && _rangeField) {
                    const _existingRange = String(tc.args?.[_rangeField] || '').trim();
                    let _realRangeAddr = '';
                    let _realSheetName = '';
                    // 地址归一化 + 单格判定：单一事实源见 services/range-resolver.js
                    // （与 dispatcher 的 _selIsRange、reference-area-manager 的单格忽略同一语义）。
                    // ⚠ 必须先归一化再判单格：GetSelection 返回的 address 默认带 $（"$G$30"），
                    //   旧实现用 /^[A-Za-z]+\d+$/ 直接测带 $ 的地址 → 单格全被漏判成"有效范围"，
                    //   拿 G30 覆盖 LLM range，最后只美化 1 个格子还报成功（实测 2026-10-22）。
                    const _rr = (typeof App !== 'undefined' && App.ChatV2 && App.ChatV2.rangeResolver) || null;
                    const _stripAddr = (a) => String(a || '').replace(/\$/g, '').trim().replace(/^[^!]*!/, '').trim();
                    const _cursorOnly = (a) => {
                        if (_rr && typeof _rr.isSingleCellAddr === 'function') return _rr.isSingleCellAddr(a);
                        const s = _stripAddr(a);
                        return !s || (s.indexOf(':') < 0 && s.indexOf(',') < 0);
                    };

                    if (_hasSelRange && !_cursorOnly(_sel.range || _sel.address)) {
                        // 有（自动捕获的）多格选区 → 用它的地址
                        _realRangeAddr = _stripAddr(_sel.range || _sel.address);
                        _realSheetName = String(_sel.sheetName || _sel.worksheet || '').trim();
                    } else {
                        // 真·无选区 → 实时从 Excel 拿。
                        // ⚠ 多来源取桥：只用 App.excelBridge 时，WebView2 面板里该引用可能没挂上
                        // （window.excelBridge / BridgeAccessor 却是可用的），导致整个兜底静默跳过、
                        // LLM 瞎编的 range（如把"A1"当整张表）原样直通后端——实测 2026-10-22
                        // "美化工作簿中的所有工作表"只美化了 A1 一个单元格还报成功。
                        const _br = App.excelBridge
                            || (typeof window !== 'undefined' && window.excelBridge)
                            || (typeof window !== 'undefined' && window.BridgeAccessor && typeof window.BridgeAccessor.getBridge === 'function' && window.BridgeAccessor.getBridge())
                            || null;
                        if (_br && typeof _br.GetSelection === 'function') {
                            try {
                                const _selRaw = await _br.GetSelection();
                                const _selLive = JSON.parse(_selRaw);
                                // 有效范围 = 非空且不是光标单格（归一化后再判：address 默认带 $）
                                if (_selLive && _selLive.address && String(_selLive.address).trim() && !_cursorOnly(_selLive.address)) {
                                    // 有效范围（不是单格 A1）
                                    _realRangeAddr = _stripAddr(_selLive.address);
                                    _realSheetName = String(_selLive.sheetName || '').trim();
                                } else {
                                    // 单格或无效 → 兜底调 UsedRange
                                    try {
                                        const _urRaw = await _br.GetUsedRangeData();
                                        const _urLive = JSON.parse(_urRaw);
                                        if (_urLive && _urLive.address && String(_urLive.address).trim()) {
                                            _realRangeAddr = _stripAddr(_urLive.address);
                                            _realSheetName = String(_urLive.sheetName || _urLive.worksheet || '').trim();
                                        }
                                    } catch (_u) { /* ignore */ }
                                }
                            } catch (_e) { /* GetSelection 失败不阻塞 */ }
                        }
                    }

                    // 有真实地址且 LLM 要么没传 range，要么传的 range 看起来瞎编的（过短或无意义）→ 覆盖
                    if (_realRangeAddr) {
                        const _llmRangeLooksFake = !_existingRange
                            || /^[A-Za-z]{1,2}\d+$/.test(_existingRange)                        // "A1" — 单格
                            || /^[A-Za-z]{1,3}\d+:[A-Za-z]{1,3}\d+$/.test(_existingRange) === false  // 非有效范围格式
                            || /^[A-Za-z]+\d+:[A-Za-z]+\d+$/.test(_existingRange) && /^[A-Za-z]{1,3}1:[A-Za-z]{1,3}\d{1,2}$/.test(_existingRange);  // 过于规整的 "A1:C10" 式范围
                        if (_llmRangeLooksFake) {
                            ChatLog.info?.('AGENT', `[agent-loop] 范围兜底: LLM range="${_existingRange || '(未传)'}" → 真实="${_realRangeAddr}"（${tc.name}）`);
                            const _overrideRange = _rangeField === 'sourceRange' ? _realRangeAddr : _realRangeAddr;
                            tc.args = { ...(tc.args || {}), [_rangeField]: _overrideRange };
                            if (_realSheetName && !tc.args.sheet && _rangeField !== 'sourceRange') {
                                tc.args.sheet = _realSheetName;
                                ChatLog.info?.('AGENT', `[agent-loop] 同时补 sheet=${_realSheetName}`);
                            }
                            if (_rangeField === 'sourceRange' && _realSheetName && !tc.args.sourceSheet) {
                                tc.args.sourceSheet = _realSheetName;
                            }
                        }
                    }
                }

                // 工作表名兜底：LLM 瞎编 "Sheet1" → 替换成真实活动工作表名
                const _knownDefaultSheets = /^Sheet\d+$|^工作表\d+$|^Worksheet\d+$/i;
                if (tc.args && typeof tc.args === 'object' && tc.args.sheet && typeof tc.args.sheet === 'string') {
                    const _llmSheet = String(tc.args.sheet).trim();
                    if (_knownDefaultSheets.test(_llmSheet)) {
                        const _selSheet = turn?.context?.selection?.sheetName
                            || turn?.context?.selection?.worksheet
                            || turn?.context?.activeSheet?.sheet
                            || turn?.context?.activeSheet?.worksheet
                            || null;
                        if (_selSheet) {
                            ChatLog.warn?.('AGENT', `[agent-loop] Sheet 名兜底: LLM="${_llmSheet}" → 实际="${_selSheet}"（${tc.name}）`);
                            tc.args = { ...tc.args, sheet: _selSheet };
                        }
                    }
                }

                // tool/call
                const displayToolName = resolveToolDisplayName(tc.name);
                setRunStatus(turnId, 'running', {
                    phase: 'tool',
                    step: stepNo,
                    stepTotal,
                    toolName: tc.name,
                    displayToolName,
                    detail: displayToolName
                });
                emitAgent('tool/call', { turnId, step: stepNo, toolName: tc.name, displayToolName, args: tc.args || {}, callId: tc.id });

                // ── 工具调用记录 ──
                let callArgsSummary = '';
                try {
                    const argsStr = JSON.stringify(tc.args || {});
                    callArgsSummary = argsStr && argsStr.length > 150 ? argsStr.slice(0, 150) + '…' : (argsStr || '{}');
                } catch (e) { callArgsSummary = '{}'; }
                try { sessionStore?.appendStepLog?.(turnId, 'tool_call', `调用工具 ${tc.name} 参数：${callArgsSummary}`, stepNo); } catch (e) { /* 忽略 */ }

                // 【最高优先级安全闸】拦截 OCR 工具在非图片场景下的错误调用
                // 只要不是 image-pipeline，就不允许调用 OCR。OCR 仅用于图片文件识别，Excel 数据必须用 excel_read_range 直接读取。
                const isOcrTool = /ocr|OCR/.test(tc.name);
                const isImagePipeline = (turn?.kind === 'image');
                const hasImageArgs = !!(tc.args?.imagePath || tc.args?.imageData || tc.args?.imageBase64);
                const hasExcelSelection = !!(turn?.context?.selection?.range || turn?.context?.selection?.address ||
                    (Array.isArray(turn?.context?.referenceRegions) && turn.context.referenceRegions.length > 0));

                // 【聚合→图表数据源贯通】（真实失败：LLM 自由调用 excel_aggregate 未传 targetSheet，
                // 后端纯聚合不落盘；随后图表步骤 LLM 凭直觉写 "A1:B6" 且带 sheet='销售数据'，
                // 后端按显式 sheet 读原表 A1:B6 → B 列是文本分类 → "没有可绘制的数值列"。
                // 修复：聚合调用自动补 targetSheet 落盘，writtenTo/groupCount 记录到 turn；
                // 图表调用未显式引用聚合结果表时，强制指向最近聚合结果表（sheet+dataRange 一起覆盖，
                // 后端显式 sheet 优先于 dataRange 前缀，两者必须同指聚合结果表）。
                // 【聚合→图表数据源贯通】核心修复：
                // 问题：LLM 不懂"多个 aggregate 写同一 targetSheet 会互相覆盖"，三次都写"聚合结果"，
                // 后两次覆盖前一次 → aggList 只剩最后一条 → 后续 chart 全部指向同一份数据。
                // 修复：targetSheet 无条件强制覆盖（不管 LLM 传没传），用本轮已落盘次数错开表名：
                //   第 1 次 aggregate → 聚合结果
                //   第 2 次 aggregate → 聚合结果2
                //   第 3 次 aggregate → 聚合结果3
                // 这样每个 aggregate 结果都独立落盘，chart 可以一对一配对。
                if (tc.name === 'excel_aggregate') {
                    const used = Array.isArray(turn && turn._aggregateResults) ? turn._aggregateResults.length : 0;
                    const ts = used === 0 ? '聚合结果' : `聚合结果${used + 1}`;
                    const oldTs = tc.args?.targetSheet || '(未传)';
                    tc.args = { ...(tc.args || {}), targetSheet: ts };
                    if (oldTs !== ts) {
                        ChatLog.info?.('AGENT', `[agent-loop] 强制覆盖 aggregate targetSheet: LLM="${oldTs}" → "${ts}"（防止多写覆盖）`);
                    } else {
                        ChatLog.info?.('AGENT', `[agent-loop] aggregate targetSheet=${ts}`);
                    }
                } else if (tc.name === 'excel_create_chart' || tc.name === 'excel_create_chart_bar') {
                    // 【数据血缘追踪 v2】基于 artifacts 体系 + 显式血缘边
                    // 从 turn._artifacts 找未消费的 aggregate artifact（优先），
                    // 若 artifacts 暂无则回退旧 _aggregateResults（兼容同轮先 aggregate 后 chart）。
                    const aggArts = Array.isArray(turn && turn._artifacts)
                        ? turn._artifacts.filter(a => a.type === 'aggregate' && !a._consumed && a.targetSheet)
                        : [];
                    const aggList = aggArts.length
                        ? aggArts.map(a => ({ sheet: a.targetSheet, dataRange: a.dataRange, _artifact: a }))
                        : ((Array.isArray(turn && turn._aggregateResults) && turn._aggregateResults.length)
                            ? turn._aggregateResults
                            : (turn && turn._aggregateResult ? [turn._aggregateResult] : []));
                    const ds = String((tc.args && (tc.args.dataRange || tc.args.data_source || tc.args.range)) || '');
                    const exclIdx = ds.lastIndexOf('!');
                    let prefix = exclIdx > 0 ? ds.slice(0, exclIdx).trim() : '';
                    if (prefix.startsWith("'") && prefix.endsWith("'")) prefix = prefix.slice(1, -1);
                    const explicitAgg = prefix && aggList.find(a => a.sheet === prefix);
                    let pickedAgg = null;
                    if (explicitAgg && !(turn._aggregateClaimed && turn._aggregateClaimed.has(explicitAgg.dataRange))) {
                        (turn._aggregateClaimed || (turn._aggregateClaimed = new Set())).add(explicitAgg.dataRange);
                        pickedAgg = explicitAgg;
                        ChatLog.info?.('AGENT', `[lineage] 显式认领 aggregate → chart: ${explicitAgg.sheet}`);
                    } else if (aggList.length) {
                        const claimed = turn._aggregateClaimed || (turn._aggregateClaimed = new Set());
                        let pick = aggList.find(a => !claimed.has(a.dataRange));
                        if (!pick) pick = aggList[aggList.length - 1];
                        claimed.add(pick.dataRange);
                        pickedAgg = pick;
                        ChatLog.info?.('AGENT', `[lineage] 顺序认领 aggregate → chart: ${pick.sheet}`);
                    }
                    // 应用对齐 + 构建血缘边
                    if (pickedAgg) {
                        const oldDs = ds || '(未传)';
                        tc.args = {
                            ...(tc.args || {}),
                            sheet: pickedAgg.sheet,
                            dataRange: pickedAgg.dataRange
                        };
                        ChatLog.info?.('AGENT', `[agent-loop] 图表数据源强制对齐: LLM dataRange="${oldDs}" → "${pickedAgg.dataRange}" (sheet=${pickedAgg.sheet})`);
                        // 【看板模式】数据取自聚合结果表，但图表必须落到"看板"表并按序锚定布局，
                        // 否则 3 张图散落在聚合结果N 表上、看板表空空如也。放在血缘对齐之后覆盖生效。
                        if (turn && turn.context && turn.context._forceDashboard) {
                            const seq = Number(turn._chartOnDashboard || 0);
                            // 布局约定：A1:I2 标题横幅（美化阶段绘制）、A3:C4 KPI 卡、图表从 B6 起依次排布
                            const anchors = ['B6:H18', 'B20:H32', 'B34:H46', 'B48:H60', 'B62:H74'];
                            const anchor = anchors[seq % anchors.length];
                            tc.args = { ...(tc.args || {}), sheet: '看板', position: JSON.stringify({ anchorRange: anchor }) };
                            turn._chartOnDashboard = seq + 1;
                            ChatLog.info?.('AGENT', `[agent-loop] 看板模式：图表放置 sheet=看板 anchor=${anchor}（第 ${seq + 1} 张）`);
                        }
                        // 【血缘边构建】把 aggregate artifact 标记为 _consumed
                        if (pickedAgg._artifact) {
                            pickedAgg._artifact._consumed = true;
                            pickedAgg._artifact._consumedBy = `pending_chart_${tc.id || Math.random().toString(36).slice(2, 6)}`;
                        }
                        // 血缘边先挂在 tc 上，等 chart artifact 捕获时补全 toId
                        if (!Array.isArray(turn._lineage)) turn._lineage = [];
                        const fromId = pickedAgg._artifact?._id || 'legacy_' + pickedAgg.sheet;
                        turn._pendingLineage = turn._pendingLineage || [];
                        turn._pendingLineage.push({
                            fromId, fromType: 'aggregate', fromSheet: pickedAgg.sheet, fromDataRange: pickedAgg.dataRange,
                            toTool: tc.name, toChartCallId: tc.id
                        });
                    } else if (prefix && !/^(undefined|null|\[object Object\]|none|nan)$/i.test(prefix)) {
                        if (String(tc.args?.sheet || '') !== prefix) {
                            tc.args = { ...(tc.args || {}), sheet: prefix };
                            ChatLog.info?.('AGENT', `[agent-loop] 图表 sheet 对齐（无聚合结果）：${prefix}`);
                        }
                    }
                }

                // 【迭代替换模型】工具执行前，先 complete 上一个 running 的 thought（可能是上一个工具的 tool 块或 LLM 的 thinking 块），
                // 然后创建一个新的 tool thought 块（title 显示工具名 + 参数摘要，状态 running）
                if (sessionStore && typeof sessionStore.startToolThought === 'function') {
                    const ps = summarizeToolParamsForThought(tc.name, tc.args || {});
                    sessionStore.startToolThought(turn.id, currentRound, tc.name, ps);
                }
                // 一旦本轮真正调用工具，开启步骤/工具卡实时渲染（与多步骤通道一致）：
                // 否则单步文本 turn 的 _enableSteps 为 false → 工具卡被抑制、正文仅终态渲染（中途只剩 3 个点）。
                if (!turn._enableSteps && sessionStore && typeof sessionStore.updateTurn === 'function') {
                    sessionStore.updateTurn(turn.id, { _enableSteps: true });
                }

                // 【空参 format_range 拦截】实测 2026-10-22 22:21：覆盖度误引导后模型回空参 format_range（{}），
                // 后端"当前选区"兜底把用户选区涂预设色（1ms 假成功 ❌ 卡片）。合法的空 range 用法必带
                // format/alignment/markLiteral/conditionalFormat/autoRange 之一；全都没有即视为无效调用。
                const _fmtHasPayload = (window.App && App.ChatV2 && App.ChatV2.toolRegistryUtil
                    && typeof App.ChatV2.toolRegistryUtil.isMeaninglessFormatRangeArgs === 'function')
                    ? !App.ChatV2.toolRegistryUtil.isMeaninglessFormatRangeArgs(tc.args)
                    : !(tc.args && (tc.args.format || tc.args.alignment || tc.args.markLiteral
                        || tc.args.conditionalFormat || tc.args.autoRange === true));
                if (tc.name === 'excel_format_range' && tc.args && typeof tc.args === 'object'
                    && !_fmtHasPayload) {
                    const _emptyArgsSummary = JSON.stringify(tc.args);
                    ChatLog.warn?.('AGENT', `[agent-loop] ⛔ 拦截空参/无效参 format_range: args=${_emptyArgsSummary}`);
                    const _emptyErr = `⛔ 拦截了无效的 excel_format_range 调用（${_emptyArgsSummary}）：缺少 format/alignment 等任何格式化参数。合法的格式化必须至少指定一种操作（如 format.bold 或 alignment.horizontal）。已跳过此调用，避免误格式化当前选区。`;
                    const _tcRec = sessionStore.addToolCall(turn.id, {
                        name: tc.name,
                        displayName: registeredTool?.displayName,
                        args: tc.args,
                        round: currentRound
                    });
                    sessionStore.updateToolCall(turn.id, _tcRec.id, {
                        status: 'failed',
                        result: null,
                        error: _emptyErr,
                        errorType: '参数无效',
                        errorReason: 'format_range 调用缺少任何格式化参数（format/alignment/markLiteral/conditionalFormat/autoRange）。',
                        errorSuggestion: '请重新指定要执行的具体格式化操作，例如 format={"bold":true} 或 alignment={"horizontal":"center"}。',
                        durationMs: 0,
                        completedAt: Date.now()
                    });
                    result = { ok: false, error: _emptyErr, durationMs: 0, toolCallId: _tcRec.id };
                } else if (isOcrTool && !isImagePipeline && (!hasImageArgs || hasExcelSelection)) {
                    // 拦截 OCR 工具错误调用：手动创建失败卡片，不实际执行工具
                    ChatLog.warn?.('AGENT', `[agent-loop] ⛔ 拦截 OCR 工具在非图片场景的错误调用: ${tc.name} kind=${turn?.kind} selection=${hasExcelSelection} args=`, tc.args);
                    const fixHint = hasExcelSelection
                        ? `检测到当前有Excel选区，您应该使用 excel_read_range 工具直接读取单元格数据（不是OCR识别图片）！请立即改为调用 excel_read_range 获取选中区域的内容。`
                        : `excel_ocr_recognize 只能用于识别图片文件（需要imagePath或imageData参数）！如果您想读取Excel工作表中的单元格数据，请使用 excel_read_range 工具。如果确实需要识别图片文字，请上传图片后再试。`;
                    const errMsg = `⛔ 工具调用错误：${tc.name} 严禁在此场景使用！${fixHint}`;
                    const tcRecord = sessionStore.addToolCall(turn.id, {
                        name: tc.name,
                        displayName: registeredTool?.displayName,
                        args: tc.args,
                        round: currentRound
                    });
                    sessionStore.updateToolCall(turn.id, tcRecord.id, {
                        status: 'failed',
                        result: null,
                        error: errMsg,
                        errorType: '工具误用',
                        errorReason: 'OCR 工具仅用于图片识别，当前场景不应调用。',
                        errorSuggestion: fixHint,
                        durationMs: 0,
                        completedAt: Date.now()
                    });
                    result = { ok: false, error: errMsg, durationMs: 0, toolCallId: tcRecord.id };
                } else {
                    // 【参数覆盖已在 emit 前应用】此处仅保留注释说明，避免与前置覆盖重复产生歧义。
                    // 正常路径（含未注册工具的错误提示）：使用通用 invokeTracked，自动处理 UI 渲染
                    // 🔧 关键修复：未在前端 toolRegistry 注册的本地 MCP 工具（computer_*/未注册 excel_*）
                    //    直接走后端 CallLocalTool，不依赖前端注册表（前端 toolRegistry 只有 excel/ppt 工具，
                    //    computer 工具集合注册在后端 McpCapabilityRegistry，前端拿不到）。
                    // 【诊断·工具参数落盘】此前工具实参从不进宿主日志，导致"sheet=订单 → 找不到工作表"
                    // 这类问题无法复盘（究竟是模型编的表名、还是带了空格/不可见字符、还是补全注入的）。
                    // 每次调用一行，成本极低、排障价值极高。
                    // 中国红被模型写成浅粉、或用 fontColor 把整段涂白（表身字消失）时，只纠正这两个参数。
                    try {
                        if (tc.name === 'excel_format_range' && tc.args && App.ChatV2.rangeResolver
                            && typeof App.ChatV2.rangeResolver.decideFormatTarget === 'function') {
                            const RR = App.ChatV2.rangeResolver;
                            const _plan = RR.decideFormatTarget(turn.message || turn.userMessage || '', turn.context);
                            const _addr = RR.normRangeAddr(tc.args.range || '');
                            if (_plan.rangeMode === 'spoken' && _plan.spokenRange) {
                                tc.args.range = _plan.spokenRange;
                            } else if ((_plan.rangeMode === 'selection' || _plan.rangeMode === 'cell' || _plan.rangeMode === 'copy-selection')
                                && _plan.selectionRange && (!_addr || RR.isSingleCellAddr(_addr))) {
                                tc.args.range = _plan.selectionRange;
                                if (!tc.args.sheet && _plan.selectionSheet) tc.args.sheet = _plan.selectionSheet;
                            } else if ((_plan.rangeMode === 'table' || _plan.rangeMode === 'header' || _plan.rangeMode === 'row' || _plan.rangeMode === 'col')
                                && tc.args.sheet && (!_addr || RR.isSingleCellAddr(_addr))) {
                                const _call = App.ChatV2.callLocalToolWithFallback;
                                if (typeof _call === 'function') {
                                    const _ladder = RR.createSheetRangeLadder(_call);
                                    const _got = await _ladder(String(tc.args.sheet), '');
                                    if (_got && _got.range) {
                                        let _r = _got.range;
                                        if (_plan.rangeMode === 'header') _r = RR.headerRowOf(_r);
                                        else if (_plan.rangeMode === 'row' && _plan.spokenRange) {
                                            const _hm = String(_r).match(/^([A-Za-z]+)\d+:([A-Za-z]+)\d+$/);
                                            if (_hm) _r = _hm[1] + _plan.spokenRange + ':' + _hm[2] + _plan.spokenRange;
                                        } else if (_plan.rangeMode === 'col' && _plan.spokenRange) {
                                            const _cm = String(_r).match(/^[A-Za-z]+(\d+):[A-Za-z]+(\d+)$/);
                                            if (_cm) _r = _plan.spokenRange + _cm[1] + ':' + _plan.spokenRange + _cm[2];
                                        }
                                        tc.args.range = _r;
                                        delete tc.args.autoRange;
                                        _hostLog('INFO', `[agent-loop] 范围按「${_plan.reason}」定为 ${tc.args.sheet}!${_r}（${_got.via}）`);
                                    }
                                }
                            } else if ((_plan.sheetScope === 'all' || _plan.sheetScope === 'others')
                                && !String(tc.args.sheet || '').trim() && RR.isSingleCellAddr(_addr)) {
                                _hostLog('WARN', `[agent-loop] 口头要多张表，这次调用没带表名和数据区（${_plan.reason}），不拿当前格去套`);
                            }
                        }
                    } catch (_rangePlanErr) { /* 范围判断失败不影响后面的颜色纠偏 */ }
                    try {
                        if (tc.name === 'excel_format_range' && tc.args && App.ChatV2.turnGoal
                            && typeof App.ChatV2.turnGoal.prepareFormatArgs === 'function') {
                            const _style = (turn && turn.goal && turn.goal.styleTarget)
                                || (turn && turn.context && turn.context._semanticStyleTarget) || '';
                            const _fixed = App.ChatV2.turnGoal.prepareFormatArgs(tc.args, _style);
                            if (_fixed && _fixed !== tc.args) {
                                _hostLog('WARN', `[agent-loop] format 纠偏 ${_style}: ${JSON.stringify(tc.args.format)} → ${JSON.stringify(_fixed.format)}`);
                                tc.args = _fixed;
                            }
                        }
                    } catch (_fmtFix) { /* 纠偏失败不影响执行 */ }
                    try {
                        if (/merge_sheets/i.test(String(tc.name || '')) && tc.args && App.ChatV2.turnGoal
                            && typeof App.ChatV2.turnGoal.constrainMergeSource === 'function') {
                            const _msg = (turn && (turn.userMessage || turn.message)) || '';
                            const _sel = turn && turn.context && turn.context.selectedSheets;
                            const _fixedMerge = App.ChatV2.turnGoal.constrainMergeSource(tc.args, _msg, _sel);
                            if (_fixedMerge && _fixedMerge !== tc.args) {
                                _hostLog('WARN', '[agent-loop] 合并范围改回选中表: '
                                    + JSON.stringify(tc.args.sourceSheets) + ' → ' + JSON.stringify(_fixedMerge.sourceSheets));
                                tc.args = _fixedMerge;
                            }
                        }
                    } catch (_mergeFix) { /* 范围纠偏失败不影响执行 */ }
                    try {
                        if (/format_range/i.test(String(tc.name || '')) && tc.args) {
                            const _calls = Array.isArray(turn && turn._executedToolCalls) ? turn._executedToolCalls : [];
                            const _merge = _calls.slice().reverse().find(function (c) {
                                return c && /merge_sheets/i.test(String(c.name || ''));
                            });
                            const _target = _merge && _merge.args && String(_merge.args.targetSheet || '合并结果').trim();
                            let _src = _merge && _merge.args ? _merge.args.sourceSheets : null;
                            if (typeof _src === 'string') {
                                try { _src = JSON.parse(_src); } catch (_) { _src = []; }
                            }
                            _src = Array.isArray(_src) ? _src.map(function (x) { return String(x || '').trim(); }) : [];
                            const _sheet = String(tc.args.sheet || '').trim();
                            if (_target && _src.length && (_sheet === '' || _src.indexOf(_sheet) >= 0) && _sheet !== _target) {
                                _hostLog('WARN', '[agent-loop] 合并之后的格式化改到结果表: ' + (_sheet || '(未指定)') + ' → ' + _target);
                                tc.args.sheet = _target;
                            }
                        }
                    } catch (_fmtTarget) { /* 改表失败不影响执行 */ }
                    try {
                        _hostLog('INFO', `[agent-loop] 调用 ${tc.name} 参数=${JSON.stringify(tc.args || {}).slice(0, 500)}`);
                    } catch (_e) { /* 日志失败不影响执行 */ }

                    // 【工作表名确定性纠偏·2026-10-01】
                    // 背景：LLM（尤其 qwen2.5）会从数据列名编造 sheet（实测造出"订单明细""销售数据"），
                    // 而 CSV 在 Excel 里打开后真实工作表名 = 文件名（如 全流程综合测试表_20261001_080927）。
                    // 后果：excel_autofit 直接抛 DISP_E_BADINDEX(0x8002000B) 整轮失败、format_range 找不到表。
                    // 只靠提示词"注入真实清单"属于劝导，模型可以无视（实测仍然编造）；这里在调用前
                    // 用宿主真实清单做确定性纠偏：sheet 不在清单内 → 清空该参数，让后端默认作用于
                    // 当前活动工作表，并把纠偏写进宿主日志留痕（GUI 的 sheet 单元格由工具结果回显）。
                    try {
                        if (tc.args && typeof tc.args.sheet === 'string' && tc.args.sheet.trim()) {
                            const realSheets = await getRealSheetNames();
                            if (realSheets.length > 0 && realSheets.indexOf(tc.args.sheet) < 0) {
                                const _keptMergeTarget = (Array.isArray(turn && turn._executedToolCalls) ? turn._executedToolCalls : [])
                                    .some(function (c) {
                                        return c && /merge_sheets/i.test(String(c.name || ''))
                                            && c.args && String(c.args.targetSheet || '').trim() === tc.args.sheet;
                                    });
                                if (_keptMergeTarget) {
                                    _hostLog('INFO', '[agent-loop] sheet「' + tc.args.sheet + '」是本轮刚合并出来的结果表，清单里还没有，保留');
                                } else {
                                    _hostLog('WARN', `[agent-loop] ⚠️ sheet 纠偏: "${tc.args.sheet}" 不在真实工作表清单 [${realSheets.join(', ')}] → 已清空该参数（改用当前活动表）工具=${tc.name}`);
                                    tc.args.sheet = '';
                                }
                            }
                        }
                    } catch (_e) { /* 纠偏失败不影响执行 */ }
                    // 已在前端注册的工具走 invokeTracked（卡片、参数补全、危险标记）。
                    // 只有注册表没有、但后端按前缀能分发的本地工具才直连 CallLocalTool。
                    const localOrmPrefix = !registeredTool && (tc.name.startsWith('computer_') || tc.name.startsWith('excel_') || tc.name.startsWith('ppt_') || tc.name.startsWith('word_'));
                    if (localOrmPrefix) {
                        try {
                            ChatLog.warn?.('AGENT', `[agent-loop] 未注册本地工具 ${tc.name} → 直接调后端 CallLocalTool`);
                            let raw = null;
                            // 【修复】优先用统一的多来源回退调用器：某些环境下 BridgeAccessor 包装桥
                            // 对 CallLocalTool 不可调用（报 "e is not a function"），而裸宿主对象可用。
                            const _fallbackCaller = (window.App && App.ChatV2 && App.ChatV2.callLocalToolWithFallback)
                                || (typeof McpDispatcher !== 'undefined' && McpDispatcher.callLocalToolWithFallback);
                            if (typeof _fallbackCaller === 'function') {
                                const _r = await _fallbackCaller(tc.name, JSON.stringify(tc.args || {}));
                                raw = _r && _r.raw;
                            } else {
                                const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (App.excelBridge || (window.chrome?.webview?.hostObjects?.excelBridge));
                                if (bridge && typeof bridge.CallLocalTool === 'function') {
                                    // 【修复】hostObjects defaultSyncMethod='async'：必须 await 拿到真实结果
                                    raw = await bridge.CallLocalTool(tc.name, JSON.stringify(tc.args || {}));
                                }
                            }
                            let parsed = null;
                            try { parsed = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch { parsed = raw; }
                            const ok = !!(parsed?.success === true || parsed?.isError === false);
                            const errText = ok ? '' : (App.ChatV2.toolRegistryUtil.safeErrorText(parsed?.error || parsed?.result || '后端调用失败'));
                            result = {
                                ok,
                                result: ok ? parsed : undefined,
                                error: ok ? null : errText,
                                durationMs: 0,
                                toolCallId: null,
                                _rawMcp: parsed
                            };
                            if (sessionStore) sessionStore.updateToolCall(turn.id, result.toolCallId, {
                                status: ok ? 'success' : 'error',
                                error: result.error || '',
                                result: ok ? JSON.stringify(parsed).slice(0, 2000) : null
                            });
                        } catch (invokeErr) {
                            result = { ok: false, error: '后端调用异常: ' + (invokeErr?.message || invokeErr), durationMs: 0, toolCallId: null };
                        }
                    } else {
                        result = await toolRegistry.invokeTracked(turn, tc.name, tc.args || {}, {
                            round: currentRound,
                            ctx: { turn, signal }  // 【修复】透传 turn，让 executeMcpToolCall 能拿到 turn.context.activeSheet 做工作表名兜底
                        });
                    }
                    // 【工作面跟随·修复活动表漂移】写类/建表类工具成功后，把 turn 上下文的
                    // activeSheet 更新为该工具实际写入/创建的工作表。此后省略 sheet 参数的调用
                    // 会被 executeMcpToolCall 的 activeSheet 兜底接到【最新工作面】，而不是
                    // turn 开始时的旧表——多步流程"中途建新表（聚合结果/透视表/清洗输出）→
                    // 活动表漂移 → 后续 range 落错表连环失败"（实测 2026-10-06）由此断根。
                    // 只认"明确改变工作面"的参数（targetSheet/outputSheet/name），不误伤普通读写。
                    if (result && result.ok && turn && turn.context) {
                        try {
                            const _a = tc.args || {};
                            const _wsByTool = {
                                excel_create_sheet: () => _a.name,
                                excel_aggregate: () => _a.targetSheet,
                                excel_create_pivot_table: () => _a.targetSheet,
                                excel_clean_data: () => _a.targetSheet || _a.outputSheet,
                                excel_date_group: () => _a.outputSheet || _a.targetSheet,
                                excel_split_by_column: () => _a.targetSheet || _a.outputSheet
                            };
                            const _getter = _wsByTool[String(tc.name || '')];
                            const _ws = _getter ? String(_getter() || '').trim() : '';
                            if (_ws) {
                                turn.context.activeSheet = Object.assign({}, turn.context.activeSheet || {}, { sheet: _ws });
                                ChatLog.info?.('AGENT', `[agent-loop] 工作面跟随: ${tc.name} → activeSheet="${_ws}"（后续缺省 sheet 调用将落在此表）`);
                            }
                        } catch (_wsErr) { /* 跟随失败不影响主流程 */ }
                    }
                    // 未注册工具的友好提示（覆盖 invokeTracked 生成的分类建议，加上具体上下文）
                    // 直连后端的本地工具失败时保留后端原文，不要改写成「工具不存在」。
                    if (!registeredTool && !localOrmPrefix && result && result.ok === false) {
                        let hint;
                        if (tc.name.startsWith('skill_')) {
                            hint = 'Skill 由系统根据用户意图自动路由，请勿直接调用 skill_xxx 工具。请在 Skill 面板确认该 Skill 已启用，或重新描述您的意图让系统自动匹配 Skill。';
                        } else if (/aggregate|聚合/i.test(tc.name)) {
                            hint = '聚合工具未注册。请确认已选中包含完整数据（含表头）的区域后重试，或改用其他可用的数据分析工具。';
                        } else {
                            // 【修复·幻觉工具名连环失败】旧提示只说"未注册"，不给出任何真实工具名，
                            // 模型只能继续编下一个不存在的名字（实测 10+ 连败烧光轮次）。
                            // 现按词元相似度给出最接近的真实工具，模型第一轮即可自纠；
                            // 连近似工具都没有时才回退到泛化提示。
                            const closeTools = findClosestToolNames(toolRegistry, tc.name, 3);
                            hint = closeTools.length > 0
                                ? `该工具不存在。最接近的真实工具：${closeTools.join('、')}。请改用其中之一（若都无法实现用户需求，请如实告知，严禁编造工具名）。`
                                : `该工具未注册，请改用可用的 MCP 工具。如需执行特定操作，请重新描述您的意图。`;
                        }
                        result.error = `Tool not found: ${tc.name}。${hint}`;
                        sessionStore.updateToolCall(turn.id, result.toolCallId, {
                            error: result.error,
                            errorType: '工具未注册',
                            errorReason: `工具 ${tc.name} 未在工具注册表中找到`,
                            errorSuggestion: hint
                        });
                        ChatLog.warn?.('AGENT', `[agent-loop] 阻止未注册工具调用: ${tc.name}`);
                    }
                }
            }

                // 空名/空 id 已在循环开头丢弃（那才会拼出重复的 tool_call id 并让接口 400）。
                // 留在列表里的每一次调用都必须回一条 tool 消息，否则下一轮请求缺少对应结果。
                // 未注册但已执行的本地工具（computer_* 等）同样要把真实结果交回模型。

                // 【迭代替换模型】工具执行完毕 → complete 刚才 startToolThought 创建的 tool 块
                // 用 result.ok 决定 status=completed/failed，durationMs + resultHint（表名等）进标题
                if (sessionStore && typeof sessionStore.completeThought === 'function') {
                    const durationMs = result && result.durationMs ? result.durationMs : 0;
                    const resultHint = buildSheetNameHint(tc.name, result?.ok === false ? null : result?.result);
                    const baseTitle = result?.ok === false
                        ? `❌ **${tc.name}** ✗（${durationMs}ms）— ${(App.ChatV2.toolRegistryUtil.safeErrorText(result?.error || '执行失败')).slice(0, 40)}`
                        : `✅ **${tc.name}** ✓（${durationMs}ms）`;
                    const fullTitle = resultHint && result?.ok !== false
                        ? baseTitle + ` → ${resultHint}`
                        : baseTitle;
                    sessionStore.completeThought(turn.id, fullTitle, {
                        round: currentRound, kind: 'tool', toolName: tc.name,
                        durationMs, status: result?.ok === false ? 'failed' : 'completed',
                        resultHint
                    });
                }

                // 【聚合结果记录】聚合调用成功后，把 writtenTo/groupCount 落到 turn，
                // 供同轮后续 excel_create_chart 调用贯通数据源（见上方调用前拦截）。
                if (tc.name === 'excel_aggregate' && result && result.ok !== false && !result.error) {
                    try {
                        let meta = result.result;
                        if (Array.isArray(meta)) {
                            const textParts = meta
                                .filter(c => c && typeof c === 'object' && typeof c.text === 'string')
                                .map(c => c.text).join('\n');
                            if (textParts) { try { meta = JSON.parse(textParts); } catch { meta = null; } }
                        } else if (typeof meta === 'string') {
                            try { meta = JSON.parse(meta); } catch { meta = null; }
                        }
                        if (meta && typeof meta === 'object' && !Array.isArray(meta)) {
                            const writtenTo = meta.writtenTo || meta.resultSheet || null;
                            const groupCount = Number(meta.groupCount) ||
                                (Array.isArray(meta.data) ? meta.data.length : 0);
                            if (writtenTo && groupCount > 0) {
                                // 【图表数据源提取】labels/values 供 finalContent 用 buildInlineChartHtml
                                // 生成内联 ECharts 图（可视化看板的图源）。
                                // 三种来源按优先级：
                                //   ① 后端 Artifact（已结构化 labels/values，最可靠）；
                                //   ② meta.data 为对象数组：[{维度列名:值, 数值列名:值}]（当前后端实现）；
                                //   ③ meta.data 为二维数组：[["部门A",100],...]（历史/其它实现）。
                                const lv = extractAggregateLabelsValues(meta, tc.args);
                                const labels = lv.labels;
                                const values = lv.values;
                                const rec = {
                                    sheet: writtenTo,
                                    dataRange: `${writtenTo}!A1:B${groupCount + 1}`,
                                    dim: meta.dimColumn || '',
                                    val: meta.valueColumn || '',
                                    func: meta.aggFunc || '',
                                    labels,
                                    values,
                                    title: `${meta.dimColumn || ''}按${meta.valueColumn || ''}${meta.aggFunc === 'count' ? '计数' : '统计'}`
                                };
                                // 最新一项写 _aggregateResult（兼容旧读取方）；
                                // 同时累积进 _aggregateResults 队列，供多图一对一认领配对。
                                turn._aggregateResult = rec;
                                if (!Array.isArray(turn._aggregateResults)) turn._aggregateResults = [];
                                const dupIdx = turn._aggregateResults.findIndex(a => a.dataRange === rec.dataRange);
                                if (dupIdx >= 0) turn._aggregateResults[dupIdx] = rec;
                                else turn._aggregateResults.push(rec);
                                ChatLog.info?.('AGENT', `[agent-loop] 聚合落盘记录：${rec.dataRange}（${groupCount} 组，${rec.dim}×${rec.val}；队列 ${turn._aggregateResults.length}）`);
                            }
                        }
                    } catch (e) {
                        ChatLog.warn?.('AGENT', '[agent-loop] 聚合结果记录异常:', e?.message);
                    }
                }

                // 【图表单请求单实例·成功标记】excel_create_chart 真正成功执行（非 skipped）后，
                // 在 turn 上置位 _chartDone，跨轮持久。后续无论同轮还是跨轮、无论参数是否不同，
                // 再出现的 create_chart 都会被 _isSingleChartBlock 拦截 → 保证单图请求只出一张图。
                if (tc.name === 'excel_create_chart' && result && result.ok && !result.skipped) {
                    turn._chartDone = true;
                }

                // 【必填参数反问】通用对话路径下工具必填参数无法兜底时，
                // 向用户反问补齐参数，而不是硬失败；同时终止本轮其余工具调用，避免级联。
                if (result.needsInput && Array.isArray(result.needsInput) && result.needsInput.length) {
                    // 【无关调用静默降级】先判断这次调用是否与用户诉求相关：
                    // 工具属于某确定性 Skill 域、而用户消息完全没命中该域关键词 →
                    // 是模型自行发挥的多余调用（实测：出图成功后幻觉调 excel_clean_data），
                    // 此时不应反问用户、更不应终止本轮——按普通失败回传 LLM 继续收尾。
                    const _sdNI = App.ChatV2?.skillDomain;
                    let _irrelevantCall = false;
                    try {
                        const _userMsg = String(turn.message || turn.userMessage || '');
                        if (_sdNI && typeof _sdNI.getAll === 'function') {
                            // excel_format_range 同时挂在 formatter / beautifier / dashboard 上。
                            // 只看第一个域会把「美化…科技绿」判成没命中 formatter，再打成「无关的多余调用」。
                            // 任一守护域的关键词命中，这次调用就和用户请求有关。
                            const _guardsNI = (_sdNI.getAll() || []).filter(d =>
                                Array.isArray(d.guardedTools) && d.guardedTools.indexOf(tc.name) >= 0
                                && Array.isArray(d.keywords) && d.keywords.length);
                            if (_guardsNI.length) {
                                _irrelevantCall = !_guardsNI.some(d => d.keywords.some(kw => kw && _userMsg.indexOf(kw) >= 0));
                            }
                        }
                        // 【显式动作优先】用户消息本身就要求该工具的动作时（明确要删/建/改名/复制工作表等），
                        // 不得仅因"没命中某个守护域关键词"就判成无关调用。
                        // 实测 `删除工作表"异常场景测试表"`：模型调 excel_delete_sheet 缺 name，而该工具
                        // 只被 dashboard 域 guardedTools 收录（关键词=看板/仪表盘…）→ 消息不命中该域
                        // → 被静默降级为失败：用户既没被执行、也没被反问（"缺参数"提示被吞掉）。
                        if (_irrelevantCall) {
                            const _EXPLICIT_ACTION_RE = {
                                excel_delete_sheet: /删除|删掉|移除|去掉|清掉/,
                                excel_create_sheet: /新建|创建|增加|添加/,
                                excel_rename_sheet: /重命名|改名|更名/,
                                excel_copy_sheet: /复制|拷贝|克隆/,
                                excel_clear_sheet: /清空|清除|清掉/,
                                excel_clear_range: /清空|清除/,
                                excel_delete_row: /删除|删掉|移除|去掉/,
                                excel_delete_column: /删除|删掉|移除|去掉/,
                                excel_format_range: /美化|格式化|主题|配色|样式|统一/
                            };
                            const _actRe = _EXPLICIT_ACTION_RE[tc.name];
                            if (_actRe && _actRe.test(_userMsg)) {
                                _irrelevantCall = false;
                                ChatLog.info?.('AGENT', `[agent-loop] 工具 ${tc.name} 缺参但用户消息显式要求该动作 → 视为相关，正常反问补齐`);
                            }
                        }
                    } catch (e) { /* 判定失败保守视为相关，走原反问流程 */ }
                    const _beautifyMissRange = tc.name === 'excel_format_range'
                        && /美化|格式化|主题|配色|样式|统一/.test(String(turn.message || turn.userMessage || ''));
                    if (_irrelevantCall || _beautifyMissRange) {
                        // 美化请求里缺 range 的调用不是「和这次无关」。同表稍后带上范围就会成功，
                        // 不能把这张表标成失败。不中断后面的调用。
                        const _skipErr = _beautifyMissRange
                            ? '已跳过：这次没带单元格范围，不代表该表美化失败'
                            : '已跳过：与本次请求无关的多余调用（缺参数）';
                        ChatLog.info?.('AGENT', `[agent-loop] 工具 ${tc.name} 缺参${_beautifyMissRange ? '（美化请求，缺范围，继续）' : '但与用户诉求无关'} → 不反问，回传 LLM 继续`);
                        try {
                            sessionStore.updateToolCall?.(turn.id, result.toolCallId, {
                                status: 'failed',
                                error: _skipErr,
                                errorSuggestion: null,
                                needsInput: null
                            });
                        } catch (e) { /* 忽略 */ }
                        try {
                            const _skipNote = _beautifyMissRange
                                ? `工具 ${tc.name} 这次没带单元格范围，已跳过，不代表该表美化失败`
                                : `工具 ${tc.name} 为多余调用（与请求无关且缺参数），已跳过反问`;
                            sessionStore?.appendStepLog?.(turnId, 'warning', _skipNote, stepNo);
                        } catch (e) { /* 忽略 */ }
                        // 不设置 closeReason、不 return —— 落入下方普通 tool/result 分支，错误回给模型
                    } else {
                    const q = String(result.error || '缺少必要参数，请补充后重试。');
                    assistantContent = (assistantContent ? assistantContent + '\n\n' : '') + '⚠️ ' + q;
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                    emitAgent('tool/needs-input', {
                        turnId, step: stepNo, toolName: tc.name, callId: tc.id,
                        missing: result.needsInput, question: q, toolCallId: result.toolCallId
                    });
                    try { sessionStore?.appendStepLog?.(turnId, 'needs_input', `工具 ${tc.name} 需补充参数：${q}`, stepNo); } catch (e) { /* 忽略 */ }
                    closeReason = 'needs_input';
                    yield { type: 'tool/needs-input', toolName: tc.name, missing: result.needsInput, question: q, toolCallId: result.toolCallId };
                    // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                }
                yield { type: 'loop_done', reason: 'needs_input', round: round + 1 };
                    return;
                    }
                }

                // tool/result
                emitAgent('tool/result', {
                    turnId, step: stepNo, toolName: tc.name, callId: tc.id,
                    ok: !!result.ok,
                    error: result.error || null,
                    durationMs: result.durationMs || 0,
                    toolCallId: result.toolCallId || null
                });

                // ── 工具返回记录（独立于上下文，压缩不清除）──
                const retContent = result.ok
                    ? `成功（耗时 ${result.durationMs || 0}ms）`
                    : `失败：${String(result.error || '未知错误').slice(0, 200)}`;
                try { sessionStore?.appendStepLog?.(turnId, 'tool_return', `工具 ${tc.name} ${retContent}`, stepNo); } catch (e) { /* 忽略 */ }

                yield {
                    type: 'tool_call_result',
                    toolCallId: result.toolCallId,
                    toolName: tc.name,
                    ok: result.ok,
                    result: result.result,
                    error: result.error,
                    durationMs: result.durationMs
                };

                // 把工具结果加入下一轮 messages。视觉页图不要进文字截断。
                const _visionImages = detachVisionImages(result);
                const _rawToolResult = result.ok ? result.result : result.error;
                let _toolContent = truncateResult(_rawToolResult, tc.name, _displayOnlyIntent);
                if (result.envelope && App.ChatV2.toolEnvelope && typeof App.ChatV2.toolEnvelope.prefixFeed === 'function') {
                    _toolContent = App.ChatV2.toolEnvelope.prefixFeed(result.envelope, _toolContent);
                }
                // 【空结果防幻觉】读取类工具执行成功却什么都没读到（当前页为空、参数指向不存在的位置等）时，
                // 若把空串原样回喂，弱模型会凭页面标题"补全"出一份看起来合理的内容（实测"假内容"）。
                // 显式改写回喂内容，要求模型如实说明"没有读取到内容"。
                if (result.ok && READ_TOOL_NAME_RE.test(String(tc.name || '')) && isEffectivelyEmptyToolResult(_rawToolResult)) {
                    try { _hostLog('WARN', `读取类工具 ${tc.name} 返回空结果 → 回喂"无内容"提示以阻断编造`); } catch (_e) { /* 日志失败不影响主流程 */ }
                    _toolContent = `⚠️ ${tc.name} 执行成功，但返回内容为空：没有读取到任何内容。`
                        + `请如实告知用户"没有读取到内容"（例如当前页 / 当前区域没有任何文字），`
                        + `严禁编造、猜测或凭页面标题虚构任何正文。`;
                }
                currentMessages.push({
                    role: 'tool',
                    tool_call_id: tc.id,
                    name: tc.name,
                    content: _toolContent
                });
                pushVisionPages(currentMessages, _visionImages);
                if (String(tc.name) === 'read_attachment' && result && result.ok && result.result && result.result.needsRecognition) {
                    _needAutoRecognize = true;
                }
                if (String(tc.name) === 'recognize_pages') _needAutoRecognize = false;

                // 收集本轮结果（供 artifact 提取用）
                localResults.push(result);

                // 【通用框架·幂等台账】记录已成功执行的"工具名+参数"，
                // 供后续轮次的重复调用拦截（见上方 ♻️ 分支）与终态工具收尾判断。
                if (turn && _idemKey) {
                    const succeeded = result && (result.envelope ? result.envelope.ok === true : !!result.ok);
                    const failText = (result && result.envelope && result.envelope.error) || (result && result.error) || '未知错误';
                    if (succeeded) {
                        if (!(turn._toolCallLedger instanceof Map)) turn._toolCallLedger = new Map();
                        turn._toolCallLedger.set(_idemKey, { ok: true, result: result.result, at: Date.now() });
                        // 成功即清零该签名的连续失败计数
                        if (turn._failStreak instanceof Map) turn._failStreak.delete(_idemKey);
                    } else {
                        // 【失败台账】成功台账只记成功，失败的没有任何记录 → 弱模型会把同一个错误
                        // 一路重放到 maxRounds（实测 2026-10-05 22:44：ppt_export_pdf 连失败 5 次，
                        // UI 弹出 5 张一模一样的失败卡片，模型最后还照着编了个文件名）。
                        // 这里按"工具名+参数"签名计连续失败次数，供执行前的 🛑 熔断分支判定。
                        if (!(turn._failStreak instanceof Map)) turn._failStreak = new Map();
                        const _fs = turn._failStreak.get(_idemKey);
                        turn._failStreak.set(_idemKey, {
                            n: ((_fs && _fs.n) || 0) + 1,
                            error: String(failText).slice(0, 200),
                            at: Date.now()
                        });
                    }
                    // 终态/重型工具 → 同步写入会话级台账（跨轮幂等，防重复生成文件/连弹卡片）
                    if (isTerminalTool(tc.name, toolRegistry)) {
                        sessionLedgerSet(turnId, _idemKey, { ok: succeeded, result: result.result, error: failText, at: Date.now() });
                    }
                }
                const _execOk = result && (result.envelope ? result.envelope.ok === true : !!result.ok);
                // 【通用框架·终态工具】生成类工具成功 = 任务已产出最终文件，标记本轮收尾
                if (_execOk && isTerminalTool(tc.name, toolRegistry)) {
                    terminalToolDone = true;
                    ChatLog.info?.('AGENT', `[agent-loop] 🏁 终态工具执行成功：${tc.name} → 本轮结束即收尾`);
                }

                // 【修复】工具失败时，清除 LLM 之前生成的"成功文字"
                // LLM 在调用工具前往往会先输出"已完成XX"等预测文字，
                // 如果工具实际执行失败，这些文字会误导用户
                // 同时标记 anyToolFailed → 本轮 toolsExecuted 不标记 → 下一轮 tools 保持可见（弱模型会重试）
                if (!_execOk) {
                    assistantContent = '';
                    sessionStore.updateTurn(turn.id, { content: '' });
                    anyToolFailed = true;
                }
            }

            // 读取附件没有文字层时，不等模型自己接上下一步，直接识别页面。
            // 过程列表仍是单独的「识别页面」卡片。模型本轮已经调用过则不重复。
            if (_needAutoRecognize && toolRegistry && !signal?.aborted) {
                const callId = 'auto_recognize_' + Date.now();
                currentMessages.push({
                    role: 'assistant',
                    content: null,
                    tool_calls: [{
                        id: callId,
                        type: 'function',
                        function: { name: 'recognize_pages', arguments: '{}' }
                    }]
                });
                let follow;
                try {
                    follow = await toolRegistry.invokeTracked(turn, 'recognize_pages', {}, {
                        round: round + 1,
                        ctx: { turn, signal }
                    });
                } catch (autoErr) {
                    follow = { ok: false, error: (autoErr && autoErr.message) || String(autoErr), durationMs: 0 };
                }
                const followImages = detachVisionImages(follow);
                const followRaw = follow && follow.ok ? follow.result : (follow && follow.error);
                const followText = truncateResult(followRaw, 'recognize_pages', false);
                currentMessages.push({
                    role: 'tool',
                    tool_call_id: callId,
                    name: 'recognize_pages',
                    content: followText
                });
                pushVisionPages(currentMessages, followImages);
                localResults.push(follow);
                yield {
                    type: 'tool_call_result',
                    toolCallId: follow && follow.toolCallId,
                    toolName: 'recognize_pages',
                    ok: !!(follow && follow.ok),
                    result: follow && follow.result,
                    error: follow && follow.error,
                    durationMs: (follow && follow.durationMs) || 0
                };
                ChatLog.info?.('AGENT', `[agent-loop] 读取附件无文字层 → 已自动识别页面 ok=${!!(follow && follow.ok)}`);
            }

            // 【强制图表兜底】图表意图 + 本轮已聚合但未建图 → 自动补一次 excel_create_chart。
            // 背景：弱模型在多步骤管道里常"只聚合不建图"或反复调用 excel_aggregate，把图表请求做成文字分析。
            // 此处用兜底强制生成图表，保证用户"按部门展示饼图"这类请求一定有图产出。
            // 条件用 isChartIntent（含 _forceChart / 消息强图表词）而非只认 context._forceChart：
            // 全部 Skill 停用时 _forceChart 未必被设上，但"用户要图"这一事实不变（实测 07:46）。
            if (turn && !anyToolFailed && isChartIntent(turn, _rawUserMsg)) {
                const _aggResults = Array.isArray(turn._aggregateResults) ? turn._aggregateResults : [];
                // 兜底触发条件：本 turn 尚未【成功】建过图（turn._chartDone 为假）且已有聚合结果。
                // 用 _chartDone 而非"是否被调用过"，这样 LLM 调图失败时不误判，fallback 仍能补出一张图。
                if (!turn._chartDone && _aggResults.length > 0) {
                    const _lastAgg = _aggResults[_aggResults.length - 1];
                    const _chartType = inferChartType((turn.userMessage || turn.message || ''), chartTypeOverride);
                    // 注：此处原传 clearExistingCharts:true，但 excel_create_chart 的后端 schema
                    // （Services/Mcp/ExcelMcpTools.ToolDefinitions.Data.vb）中并无该属性——
                    // 它是个不被后端解析的死参数，既不生效也无副作用。
                    // "单图自清理"本就是后端 create_chart 的默认行为，且已收敛为
                    // 【只清 AI 自建的 AIChart_ 前缀图、不碰用户手工图表】，
                    // 无需前端传参。移除该参数，避免误导后来者以为清图开关在此处。
                    const _chartArgs = {
                        chartType: _chartType,
                        dataRange: _lastAgg.dataRange,
                        sheet: _lastAgg.sheet,
                        title: _lastAgg.title || '图表'
                    };
                    const _chartTc = {
                        id: 'syn_chart_' + Math.random().toString(36).slice(2, 8),
                        name: 'excel_create_chart',
                        args: _chartArgs
                    };
                    ChatLog.warn?.('AGENT', `[agent-loop] _forceChart 兜底：LLM 未调用 create_chart，自动合成调用，数据源=${_lastAgg.dataRange}`);
                    let _chartResult;
                    try {
                        _chartResult = await toolRegistry.invokeTracked(turn, 'excel_create_chart', _chartArgs, {
                            round: round + 1,
                            ctx: { turn, signal }
                        });
                    } catch (_chartErr) {
                        _chartResult = { ok: false, error: _chartErr?.message || String(_chartErr), durationMs: 0, toolCallId: null };
                    }
                    if (_chartResult) {
                        toolCalls.push(_chartTc);
                        localResults.push(_chartResult);
                        emitAgent('tool/result', {
                            turnId, step: stepNo, toolName: _chartTc.name, callId: _chartTc.id,
                            ok: !!_chartResult.ok,
                            error: _chartResult.error || null,
                            durationMs: _chartResult.durationMs || 0,
                            toolCallId: _chartResult.toolCallId || null
                        });
                        try { sessionStore?.appendStepLog?.(turnId, 'tool_return', `工具 ${_chartTc.name} ${_chartResult.ok ? '成功' : '失败'}（兜底合成）`, stepNo); } catch (e) { /* 忽略 */ }
                        yield {
                            type: 'tool_call_result',
                            toolCallId: _chartResult.toolCallId,
                            toolName: _chartTc.name,
                            ok: _chartResult.ok,
                            result: _chartResult.result,
                            error: _chartResult.error,
                            durationMs: _chartResult.durationMs
                        };
                        const _rawToolResult = _chartResult.ok ? _chartResult.result : _chartResult.error;
                        let _toolContent = truncateResult(_rawToolResult, _chartTc.name, false);
                        currentMessages.push({
                            role: 'tool',
                            tool_call_id: _chartTc.id,
                            name: _chartTc.name,
                            content: _toolContent
                        });
                        if (_chartResult.ok) {
                            terminalToolDone = true;
                            turn._chartDone = true;  // 兜底合成成功 → 标记，防止后续轮次再补或重复建图
                        } else {
                            anyToolFailed = true;
                        }
                    }
                } else if (!turn._chartDone && !turn._aggFallbackTried) {
                    // 【图表补链】模型连聚合都没做 → 确定性补齐 aggregate → create_chart。
                    // 若本轮工具全是只读探测（read_range 等），先不抢跑：留给下一轮零工具时的
                    // 图表闸门（先 nudge 一次，用尽后再当场补链），避免"读完立刻建图"抢走模型一轮机会。
                    // 非只读轮（已聚合/写过表等）仍同轮补链。
                    const _onlyProbe = (App.ChatV2.agentLoopToolExecution
                        && typeof App.ChatV2.agentLoopToolExecution.isOnlyProbeToolRound === 'function')
                        ? App.ChatV2.agentLoopToolExecution.isOnlyProbeToolRound(toolCalls)
                        : (toolCalls.length > 0 && toolCalls.every(tc =>
                            /^(excel_)?(read_range|get_|list_|describe_)/i.test(String(tc && tc.name || '').replace(/^mcp_/, ''))));
                    if (_onlyProbe) {
                        _hostLog('INFO', '图表补链延后：本轮仅只读探测 → 交图表闸门（nudge→当场补链）');
                    } else {
                        const _bfOut = yield* runChartBackfill({
                            turn, toolCalls, localResults, round, stepNo, signal,
                            rawMsg: _rawUserMsg, chartTypeOverride
                        });
                        if (_bfOut.anyToolFailed) anyToolFailed = true;
                        if (_bfOut.chartDone) terminalToolDone = true;
                    }
                }
            }

            // 【关键修复】工具失败时不标记 toolsExecuted
            // 原逻辑：只要本轮调过工具 → 不管成功失败 → toolsExecuted=true → 下一轮 shouldProvideTools=false
            // 导致弱模型工具瞎编参数失败后，下一轮拿不到 tools → 只能输出文字总结（假装成功）
            // 新逻辑：本轮至少一次成功且无硬失败（含熔断 skip）才标记 → 下一轮 tools 仍可见
            // 多步骤管道（persistTools=true）不受此影响，但 fallback 管道依赖 toolsExecuted 锁死逻辑
            // 本轮之前的已执行签名（用于"全重复"判定）——必须在 recordToolExecutions
            // 写入本轮签名之前快照，否则本轮调用会被自己判成"已执行过的重复调用"。
            const _beforeSigs = new Set(_executedSignatures);
            const _merged = (App.ChatV2.agentLoopToolExecution
                && typeof App.ChatV2.agentLoopToolExecution.mergeToolRoundFlags === 'function')
                ? App.ChatV2.agentLoopToolExecution.mergeToolRoundFlags(localResults, anyToolFailed)
                : {
                    anyOk: Array.isArray(localResults) && localResults.some(r => r && r.ok === true),
                    anyToolFailed: anyToolFailed || (Array.isArray(localResults) && localResults.some(r => r && r.ok === false)),
                    toolsExecuted: false
                };
            anyToolFailed = _merged.anyToolFailed;
            const _anyOkThisRound = _merged.anyOk;
            if (toolCalls.length > 0 && _merged.toolsExecuted) {
                toolsExecuted = true;
                recordToolExecutions(turn, toolCalls, localResults, round);
            } else if (toolCalls.length > 0 && _anyOkThisRound) {
                // 部分成功：记进度，但不锁死 tools（仍有失败/熔断需重试）
                recordToolExecutions(turn, toolCalls, localResults, round);
            }

            // 模型一次输出往往写不完「所有工作表」（实测 00:37 只发出 4 张，市场部被截在外面）。
            // 清单里还缺的表，按本轮已经成功的 format 原样补上，不再等下一轮模型自己想起来。
            if (turn && !turn._formatGapFilled && toolRegistry && typeof toolRegistry.invokeTracked === 'function') {
                const _tg = App.ChatV2.turnGoal;
                const _missing = _tg && typeof _tg.missingFormatSheets === 'function' ? _tg.missingFormatSheets(turn) : [];
                const _rr = App.ChatV2.rangeResolver;
                const _fillable = _missing.filter(function (s) {
                    return s && s.usedRange && _rr && typeof _rr.isReliableTableRange === 'function' && _rr.isReliableTableRange(s.usedRange);
                });
                const _sample = (Array.isArray(turn.toolCalls) ? turn.toolCalls : []).slice().reverse().find(function (c) {
                    return c && /format_range/i.test(String(c.name || '')) && c.status === 'completed'
                        && c.args && c.args.format && (typeof c.args.format !== 'object' || Object.keys(c.args.format).length > 0);
                });
                const _style = turn.goal && turn.goal.styleTarget;
                let _sampleFormat = _sample && _sample.args && _sample.args.format;
                // 模型只取了清单就停（编辑重发常见）时，没有可抄的 format。
                // 目标外观能落成具体底色（中国红）才补；落不成的不发明一套商务蓝。
                if (!_sampleFormat && _style && _tg && typeof _tg.prepareFormatArgs === 'function') {
                    const _prepared = _tg.prepareFormatArgs({ format: {} }, _style);
                    const _fmt = _prepared && _prepared.format;
                    if (_fmt && typeof _fmt === 'object' && (_fmt.bgColor || _fmt.themeHeader)) _sampleFormat = _fmt;
                }
                if (_fillable.length && _sampleFormat) {
                    turn._formatGapFilled = true;
                    for (const _sheet of _fillable) {
                        let _args = { sheet: _sheet.name, range: _sheet.usedRange, format: _sampleFormat };
                        if (_tg && typeof _tg.prepareFormatArgs === 'function') _args = _tg.prepareFormatArgs(_args, _style);
                        _hostLog('WARN', '全表格式化补齐遗漏表: ' + _sheet.name + ' ' + _sheet.usedRange);
                        try {
                            const _r = await toolRegistry.invokeTracked(turn, 'excel_format_range', _args, {
                                round: round + 1,
                                ctx: { turn: turn, signal: signal }
                            });
                            if (_r && _r.ok) {
                                toolsExecuted = true;
                                recordToolExecutions(turn, [{ name: 'excel_format_range', args: _args }], [{ ok: true }], round);
                            }
                        } catch (_gapErr) {
                            _hostLog('WARN', '全表格式化补齐失败: ' + _sheet.name + ' ' + ((_gapErr && _gapErr.message) || _gapErr));
                        }
                    }
                }
            }

            // 跨轮冗余防护：若本轮每一个工具调用都与之前某轮"完全相同"（同名+同参），
            // 说明模型在重复请求同一操作（典型表现：把同一张表反复整表格式化 → 界面底色循环刷新、
            // 聊天栏却一直只有 3 个点，因为中途内容被缓冲）。直接收尾，避免死循环式重刷文档。
            if (detectRedundantRepeat(toolCalls, _beforeSigs)) {
                closeReason = 'redundant_repeat_calls';
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                }
                yield { type: 'loop_done', reason: 'redundant_repeat_calls', round };
                return;
            }

            // step/end（本轮完成）
            setRunStatus(turnId, 'running', { phase: 'turn', step: round + 1 });
            emitAgent('step/end', { turnId, step: round + 1, reason: 'step_completed' });

            // ── 本轮完成 ──
            try { sessionStore?.appendStepLog?.(turnId, 'finish', `第 ${round + 1} 轮完成（${toolCalls.length} 个工具调用）`, round + 1); } catch (e) { /* 忽略 */ }

            round++;
            if (round >= effectiveMaxRounds) {
                closeReason = 'max_rounds_reached';
                // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                }
                yield { type: 'loop_done', reason: 'max_rounds_reached', round };
                return;
            }

            yield { type: 'round_end', round };
        }

        closeReason = 'exited_loop';
        // 【防累积】loop 收尾前强制用本轮 assistantContent replace turn.content，
                // 避免前几轮的中间规划文字残留在 appendDelta 累积结果里
                if (typeof assistantContent !== 'undefined' && assistantContent && assistantContent.trim()) {
                    sessionStore.updateTurn(turn.id, { content: assistantContent });
                }
                yield { type: 'loop_done', reason: 'exited_loop', round };
        } finally {
            // 确保轮次正确关闭（含取消/错误/提前 return）
            try {
                emitAgent('step/end', { turnId, step: round + 1, reason: closeReason || 'turn_closed' });
            } catch (e) { /* 忽略 */ }
            const endPhase = (closeReason === 'user_aborted' || closeReason === 'replaced_by_new_turn')
                ? 'canceled'
                : (closeReason && closeReason !== 'completed' && closeReason !== 'exited_loop' ? 'idle' : 'done');
            setRunStatus(turnId, 'idle', { phase: endPhase, step: 0, toolName: null });
            emitAgent('turn/end', { turnId, reason: closeReason || 'completed' });
            // ── turn 收尾（含取消/错误）──
            finishStepLog(closeReason || 'completed');
            _runStates.delete(turnId);

            // turn 关闭后按需压缩（异步，不阻塞返回）
            // 对齐 DSH：compaction 在轮次边界处理压力，防止长对话上下文溢出
            try {
                const compaction = App.ChatV2?.compaction;
                if (compaction && typeof compaction.compactIfNeeded === 'function') {
                    const sessionStore = App.ChatV2?.sessionStore;
                    if (sessionStore) {
                        // fire-and-forget，异常隔离
                        compaction.compactIfNeeded(sessionStore).then(res => {
                            if (res.needed) {
                                ChatLog.info?.('AGENT', `[agent-loop] turn 关闭后自动压缩: ${res.reason} compacted=${res.compacted}`);
                            }
                        }).catch(e => {
                            ChatLog.warn?.('AGENT', '[agent-loop] 自动压缩失败（不影响主流程）:', e?.message);
                        });
                    }
                }
            } catch (e) { /* 忽略 */ }
        }
    }

    App.ChatV2.agentLoop = {
        run,
        MAX_TOOL_ROUNDS,
        // 运行状态查询 API
        getRunStatus,
        listRunStates,
        // agent/pre-step 拦截注册 API
        onPreStep,
        // 导出工具匹配器（供 text-pipeline 把强匹配 skill 加入 LLM 可见工具）
        findCandidateTools,
        // 导出估算/水位检查（调试面板与测试用）
        estimateMessagesTokens,
        checkTokenWatermark,
        // 会话级幂等台账重置（新对话/切换会话时调用，避免跨对话误挡）
        resetSessionLedger: resetSessionToolLedger,
        // 文本工具调用解析（含"方括号外参数"修复）；导出供单元测试
        parseTextToolCalls,
        // 工具结果回喂 LLM 前的截断/摘要（excel_merge_sheets 等特化在此）；导出供回归测试
        truncateResult,
        // 【图表交付闸门】判定与数据源推断；导出供回归测试
        // （实测 07:46：图表请求只出文字不出图 → 见 Tests/agent-loop-chart-gate.test.js）
        isChartIntent,
        resolveAggregateSource,
        buildChartNudgeHint
    };
    ChatLog.info?.('CHAT-V2', '✅ agentLoop loaded');
})();

