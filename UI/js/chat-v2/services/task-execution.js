/**
 * ============================================================
 * chat-v2/services/task-execution.js — 子任务执行/验证/汇总/徽章
 * ============================================================
 * 职责：从 text-pipeline 抽离的公共执行辅助工具。
 *   - makeTaskCounters：统一记录 toolCalls 基线 + 验证 + 计数
 *   - buildTasksSummary：按 全部失败/全部成功/部分成功 三分支构造 summary
 *   - buildLinkBadge / applyLinkBadge：统一"来路徽章"HTML 生成与写入
 *   - invokeToolDirectly：确定性直接调用工具（跳过 LLM agent-loop）
 * 依赖：sessionStore、toolRegistry（运行时传入，不硬编码）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    function makeTaskCounters() {
        const succeeded = [];
        const failed = [];
        return {
            succeeded,
            failed,
            begin(sessionStore) {
                const cur = sessionStore.getCurrentTurn();
                const baseline = Array.isArray(cur?.toolCalls) ? cur.toolCalls.length : 0;
                return {
                    baseline,
                    forcedOK: false,
                    forcedFail: false,
                    markOK() { this.forcedOK = true; this.forcedFail = false; },
                    markFail() { this.forcedFail = true; this.forcedOK = false; },
                    finalize(sessionStore, label) {
                        let ok = this.forcedOK && !this.forcedFail;
                        if (!ok && !this.forcedFail) {
                            const after = sessionStore.getCurrentTurn();
                            const newTC = Array.isArray(after?.toolCalls) ? after.toolCalls.slice(this.baseline) : [];
                            ok = newTC.some(tc => tc && (tc.status === 'success' || tc.status === 'completed'));
                            ChatLog.info?.('PIPELINE', `[task-execution] 子任务"${label}"验证：新增 toolCalls=${newTC.length}, success=${ok}` +
                                (newTC.length > 0 ? `, statuses=[${newTC.map(t => t.status || '?').join(',')}]` : ''));
                        }
                        if (ok) succeeded.push(label); else failed.push(label);
                        return ok;
                    }
                };
            },
            record(label, ctx, sessionStore) { return ctx.finalize(sessionStore, label); },
            okCount() { return succeeded.length; },
            failCount() { return failed.length; },
            totalCount() { return succeeded.length + failed.length; }
        };
    }

    function buildTasksSummary(counters, kind, itemName, plural) {
        const ok = counters.okCount();
        const fail = counters.failCount();
        const labels = counters.succeeded;
        const fails = counters.failed;
        const plr = plural || (kind === 'aggregate' ? '个分类汇总' : kind === 'chart' ? '个图表' : '项任务');
        const typeLabel = kind === 'aggregate' ? '分类汇总' : kind === 'chart' ? '图表' : '任务';
        if (ok === 0) {
            if (fail >= 2) return `⚠️ ${fail} ${plr}均未能生效：${fails.join('、')}。请确认已在 Excel 中选中包含表头的完整数据区域，然后重试。`;
            const one = fails[0] || itemName || typeLabel;
            return `⚠️ 未能${kind === 'aggregate' ? '执行' : '创建'}${one}。请确认已在 Excel 中选中包含表头的完整数据区域，然后重试。`;
        }
        if (fail === 0) {
            if (ok >= 2) return `✅ 已完成 ${ok} ${plr}：${labels.join('、')}`;
            const one = labels[0] || itemName || typeLabel;
            return `✅ 已${kind === 'aggregate' ? '完成' : '创建'}${one}。`;
        }
        return `✅ 已完成 ${ok} 项：${labels.join('、')}；⚠️ ${fail} 项未生效：${fails.join('、')}。`;
    }

    function buildLinkBadge(source, taskName, mode) {
        const safeName = String(taskName || '').replace(/[<>"']/g, ' ').trim();
        const safeMode = String(mode || '').replace(/[<>"']/g, ' ').trim();
        const safeSource = String(source || '通用对话').replace(/[<>"']/g, ' ').trim();
        const parts = [safeSource];
        if (safeName) parts.push(safeName);
        if (safeMode) parts.push(safeMode);
        return `<div data-link-badge="1" style="margin:4px 0 8px 0;padding:6px 12px;background:#eef4ff;border-left:3px solid #5b8def;border-radius:4px;font-size:12px;color:#334155;line-height:1.6;">🛠️ ${parts.join(' · ')}</div>`;
    }

    function applyLinkBadge(turn, sessionStore, source, taskName, mode) {
        const badge = buildLinkBadge(source, taskName, mode);
        try { sessionStore.updateTurn(turn.id, { _linkBadge: badge }); } catch (_) { /* UI 字段未知时忽略 */ }
        return badge;
    }

    async function invokeToolDirectly(turn, toolRegistry, sessionStore, toolName, args, opts) {
        const tool = toolRegistry.get(toolName);
        if (!tool) {
            ChatLog.warn?.('PIPELINE', `[task-execution] 直接调用失败：工具 ${toolName} 未注册`);
            return { ok: false, error: `工具 ${toolName} 未注册` };
        }
        ChatLog.info?.('PIPELINE', `[task-execution] 🛡️ 确定性直接调用 ${toolName}，跳过 LLM agent-loop`);
        try {
            const invokeOpts = { round: 1, ...(opts || {}) };
            const result = await toolRegistry.invokeTracked(turn, toolName, args || {}, invokeOpts);
            return result || { ok: false };
        } catch (e) {
            ChatLog.warn?.('PIPELINE', `[task-execution] 直接调用 ${toolName} 异常:`, e?.message);
            return { ok: false, error: e?.message || 'unknown' };
        }
    }

    /**
     * 【步骤参数不足 → 模型补全】当某聚合步骤因参数缺失/不正确而无法执行时，
     * 用 LLM 分析补全 dimColumn/valueColumn/aggFunc/chartType，再确定性重调 excel_aggregate。
     * 顺序遵循"先确定性补全，仍缺/仍失败才走模型"：本函数只负责最后的模型兜底。
     * @returns {Promise<Object|null>} 成功返回 {args, result}；失败/参数仍不足返回 null
     */
    async function refillAggregateParams(opts) {
        const { turn, toolRegistry, sessionStore, signal, hintText, forcedChartType, toolName, displayName, round } = opts || {};
        const llmIo = App.ChatV2.llmIo;
        if (!llmIo || typeof llmIo.askJson !== 'function') return null;
        let json = null;
        try {
            const res = await llmIo.askJson({
                messages: [
                    {
                        role: 'system',
                        content: '你是 Excel 分类汇总参数补全助手。用户希望对某一维度按某方式汇总某一指标并生成图表。请只返回一个 JSON 对象，包含以下全部字段：' +
                            '{"dimColumn":"维度列名(如 部门/地区/产品/日期)","valueColumn":"指标列名(如 利润/销售额/数量)","aggFunc":"sum|avg|count|max|min",' +
                            '"chartType":"xlPie|xlColumnClustered|xlBarClustered|xlLine|xlDoughnut|xlArea|xlXYScatter|xlRadar|xlFunnel|xlWaterfall"}。' +
                            '字段必须完整、可直接使用，不要任何多余解释或代码块。'
                    },
                    {
                        role: 'user',
                        content: '需求：' + String(hintText || '').slice(0, 200) +
                            (forcedChartType ? `\n图表类型必须为 ${forcedChartType}，保持不变。` : '')
                    }
                ],
                signal
            });
            json = res && res.json;
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[task-execution] 模型补全参数·流式请求失败:', e?.message);
        }
        if (!json || typeof json !== 'object') return null;
        const dim = json.dimColumn;
        const val = json.valueColumn;
        if (!dim || !val) {
            ChatLog.warn?.('PIPELINE', '[task-execution] 模型补全参数·缺少维度/指标，放弃:', JSON.stringify(json));
            return null;
        }
        const aggRaw = String(json.aggFunc || 'sum').toLowerCase();
        const agg = ['sum', 'avg', 'count', 'max', 'min'].indexOf(aggRaw) >= 0 ? aggRaw : 'sum';
        const chart = forcedChartType || json.chartType || 'xlColumnClustered';
        const args = { dimColumn: dim, valueColumn: val, aggFunc: agg, chartType: chart, targetSheet: '聚合结果' };
        ChatLog.info?.('PIPELINE', `[task-execution] 模型补全参数 → excel_aggregate(${JSON.stringify(args)})`);
        try {
            const result = await invokeToolDirectly(turn, toolRegistry, sessionStore, toolName || 'excel_aggregate', args, {
                round: round || 99,
                displayName: displayName || '模型补全参数',
                ctx: { turn, sessionStore, signal }
            });
            if (result && result.ok === false) return null;
            return { args, result };
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[task-execution] 模型补全参数·执行失败:', e?.message);
            return null;
        }
    }

    /**
     * 【统一套壳】"执行规划"卡：各规划路径（多图表/多聚合/单聚合/多步骤）只传一条描述数组，
     * 由这里统一生成并持久化 _planHtml，避免散落多处维护同一份 HTML。
     * @param {Object} turn
     * @param {Object} sessionStore
     * @param {string[]} itemDescs 每个子任务的一行描述（如 "饼图 · 按部门展示利润"）
     * @param {Object} [options] 可选参数
     * @param {string} [options.title] 面板标题（默认"📋 执行规划 (N 个子任务)"）
     * @param {string} [options.prefix] 步骤前缀（默认数字序号）
     * @param {Object} [options.style] 样式覆盖
     * @returns {string} 生成的 HTML（空数组时返回 ''）
     */
    function persistPlanHtml(turn, sessionStore, itemDescs, options) {
        const list = Array.isArray(itemDescs) ? itemDescs : [];
        if (!list.length) return '';
        const opts = options || {};
        const title = opts.title || (list.length > 1 ? `📋 执行规划 (${list.length} 个子任务)` : '📋 执行规划');
        const rows = list.map((d, i) => {
            const prefix = opts.prefix !== undefined ? opts.prefix : `${i + 1}. `;
            const label = typeof d === 'object' ? (d.description || d.name || '') : d;
            return `<div style="padding:6px 10px;font-size:13px;color:#475569;">${prefix}${label}</div>`;
        }).join('');
        const html = `<div style="margin:8px 0;padding:8px 12px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:6px;"><div style="font-size:13px;font-weight:600;color:#334155;margin-bottom:4px;">${title}</div>${rows}</div>`;
        sessionStore.updateTurn(turn.id, { _planHtml: html });
        // 【步骤统一到 cbx v20261003】旧的 _planHtml 卡已下线，把规划条目也写进 thoughtList 保证可见
        if (typeof sessionStore.upsertStepThought === 'function') {
            list.forEach(function (d, i) {
                const label = (typeof d === 'object') ? (d.description || d.name || '') : String(d);
                sessionStore.upsertStepThought(turn.id, {
                    id: 'task_exec_' + i,
                    title: label || ('步骤 ' + (i + 1)),
                    detail: '',
                    status: 'pending',
                    index: i
                });
            });
        }
        ChatLog.info?.('PIPELINE', '[task-execution] 执行规划已持久化, 长度=' + html.length);
        return html;
    }

    App.ChatV2.taskExecution = {
        makeTaskCounters,
        buildTasksSummary,
        buildLinkBadge,
        applyLinkBadge,
        invokeToolDirectly,
        refillAggregateParams,
        persistPlanHtml
    };

    ChatLog.info?.('CHAT-V2', '✅ taskExecution loaded');
})();
