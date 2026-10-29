/**
 * ============================================================
 * chat-v2/pipelines/text-pipeline.js — 文本对话流程（重构 v20260924）
 * ============================================================
 * 职责：纯调度器 — 消费 intentDetector / taskPlanner /
 *   renderHtml / taskExecution / scriptGenerator 五个 service，
 *   只保留路由决策 + 执行编排 + 结果汇总。
 *
 * 重构要点（v20260924）：
 *   - execute() 瘦身成纯调度器：初始化 → 工具准备 → 意图归一化 → 分发 → 兜底
 *   - 多图表专用内联编排（原 L341-932）已按用户要求彻底删除：
 *     含图表/聚合意图的消息统一走多步骤管道或 agent-loop 兜底
 *   - 多步骤引擎（原 L933-1135）提取为 runMultiStepPipeline
 *   - 公共数据源三段降级（原 L362-437）提取为 resolveTopDataSource
 *   - 保留骨架：所有辅助函数 / IIFE / pipelineRegistry.register 完全不动
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    // ====== service 依赖抽取（顶层闭包，所有私有函数共享）======
    // intentDetector      → 意图识别正则 + detect* 函数
    // taskPlanner          → splitChartTasks / splitAggregateTasks / buildSingleAggregateTask / prompt 构造
    // renderHtml           → buildAggChartHtml / buildInlineChartHtml / buildAggTableHtml / buildCombinedHtml / esc
    // taskExecution        → makeTaskCounters / buildTasksSummary / buildLinkBadge / applyLinkBadge / invokeToolDirectly
    // promptBuilder        → buildMessages / detectAnalysisIntent
    // toolRegistry         → listBySource / invokeTracked / isImageOnlyTool / get / init
    // agentLoop            → run / findCandidateTools
    // sessionStore         → setExecSpec / getHistoryForLLM / clearExecSpec / completeTurn / failTurn / appendDelta / updateTurn

    /**
     * 统一依赖访问入口：从 App.ChatV2 提取所有 service 引用。
     * 所有私有函数通过 getDeps() 获取依赖，避免在每个函数内重复解构。
     * @returns {Object} App.ChatV2 命名空间引用
     */
    function getDeps() {
        return App.ChatV2;
    }

    /**
     * 从 toolCall 中提取可读的工具结果摘要（取前 300 字）。
     * 优先走 renderers.toolCard.unwrapToolResult 解包，兼容不同工具返回结构。
     * 失败态（success/ok=false）时返回"失败：xxx"明确告知用户。
     * @param {Object|null} tc - toolCall 对象（含 result 字段）
     * @returns {string} 可读摘要；tc 为 null 或无法提取时返回空串
     */
    function summarizeToolResult(tc) {
        if (!tc) return '';
        const unwrap = App.ChatV2?.renderers?.toolCard?.unwrapToolResult;
        try {
            let data = tc.result;
            if (unwrap && typeof unwrap === 'function') data = unwrap(data);
            if (typeof data === 'string') return data.slice(0, 300);
            if (data && typeof data === 'object') {
                if (data.success === false || data.ok === false) {
                    return '失败：' + String(data.error || data.message || '未知原因');
                }
                for (const k of ['summary', 'message', 'resultText', 'msg', 'text']) {
                    if (typeof data[k] === 'string' && data[k].trim()) return data[k].slice(0, 300);
                }
                const s = JSON.stringify(data);
                return s.length > 300 ? s.slice(0, 300) + '…' : s;
            }
            if (data != null) return String(data);
        } catch (e) { /* 忽略 */ }
        return '';
    }

    /**
     * 提取工具返回的原始结果文本（不含工具名标签）。
     * 用户要求直接展示工具"返回的结果"字段的值，而非 LLM 重新组织的回复。
     */
    function buildRawToolResultText(turn) {
        const tcs = turn && Array.isArray(turn.toolCalls) ? turn.toolCalls : [];
        const done = tcs.filter(t => t && t.status === 'completed');
        if (!done.length) return '';
        const lines = [];
        for (const tc of done) {
            const summary = summarizeToolResult(tc);
            if (summary) lines.push(summary);
        }
        return lines.join('\n');
    }

    /** 模型没写总结时：全表格式化用覆盖事实，其它任务仍回退工具原文。 */
    function textWhenModelSilent(turn, llmText, rawToolResult) {
        const coverageText = (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.formatCoverageText === 'function')
            ? App.ChatV2.turnGoal.formatCoverageText(turn) : '';
        return (llmText && String(llmText).trim().length > 0) ? llmText : (coverageText || rawToolResult || llmText);
    }

    // ============================================================
    // 【骨架保留】后续消息继承（Phase 5）：读取上一轮 skill/text _execSpec，注入上轮图表结果
    // ============================================================
    /**
     * 判定当前 turn 是否应该基于上一轮 skill/text 管道结果继续执行。
     * 条件：turn.context._execFollowUp 存在 + sessionStore.getLastExecSpec 返回
     *       上一轮 pipeline ∈ { skill, text }。
     * @param {Object} turn
     * @returns {Object|null} 上一轮 ExecSpec（含 pipeline/data/result/taskStructure），
     *   不满足条件或无有效历史时返回 null
     */
    function getChartFollowUp(turn) {
        if (!turn || !turn.context || !turn.context._execFollowUp) return null;
        const ss = App.ChatV2.sessionStore;
        if (!ss || typeof ss.getLastExecSpec !== 'function') return null;
        const prev = ss.getLastExecSpec();
        if (!prev || !prev.spec) return null;
        const p = prev.spec.pipeline;
        if (p !== 'skill' && p !== 'text') return null;
        return prev.spec;
    }

    /**
     * 后续消息继承 —— 把上一轮图表结果 + 数据快照注入当前 turn。
     * 两阶段：
     *   1) 优先用 intentAnalyzer.buildFollowUpPlanningPrompt 构造规划 prompt，
     *      注入上轮 result + taskStructure，让 LLM 基于上轮规划继续而非只靠关键词。
     *   2) 兜底（旧逻辑）：手动拼接"上一轮结果/任务/规划"三段文本，
     *      兼容 intentAnalyzer 未就绪的场景。
     * 数据快照继承：校验 selection 必需字段（address + sheet），
     * 不全时跳过并 warn，防止静默继承坏数据（v20260924 schema 守卫）。
     * @param {Object} turn - 当前 turn（原地修改 userMessage + context.selection/referenceRegions）
     */
    function applyChartFollowUp(turn) {
        const prevSpec = getChartFollowUp(turn);
        if (!prevSpec) return;
        if (turn._originalUserMessage === undefined) {
            turn._originalUserMessage = turn.userMessage || '';
        }
        const intent = (turn.context && turn.context._execFollowUp && turn.context._execFollowUp.intent) || null;
        // 优先用意图分析器构造规划 prompt：注入任务结构 + 用户新需求 + 意图 + 规划指令
        const analyzer = App.ChatV2.intentAnalyzer;
        let planningPrompt = null;
        if (analyzer && typeof analyzer.buildFollowUpPlanningPrompt === 'function') {
            planningPrompt = analyzer.buildFollowUpPlanningPrompt({
                prevSpec: { spec: prevSpec },
                userMessage: turn.userMessage || '',
                intent: intent
            });
        }
        if (planningPrompt) {
            turn.userMessage = (turn.userMessage || '') + '\n\n' + planningPrompt;
        } else {
            // 兜底：旧逻辑
            const res = prevSpec.result || {};
            const parts = [];
            if (res.summary) parts.push('【上一轮图表结果】\n' + String(res.summary).slice(0, 2000));
            if (res.chartTasks && res.chartTasks.length) {
                parts.push('【上一轮图表任务】\n' + res.chartTasks.map(function (t) {
                    return '- ' + (t.chartLabel || t.title || t.type || '图表') + (t.text ? '：' + t.text : '');
                }).join('\n'));
            }
            // 注入上一轮任务结构（规划步骤/任务清单），让后续消息基于上轮规划继续而非只靠关键词。
            // 兼容两种结构：text/python 分析管道存 ts.tasks；skill 管道存 ts.steps。
            const ts = prevSpec.taskStructure;
            if (ts) {
                const planTasks = Array.isArray(ts.tasks) ? ts.tasks : (Array.isArray(ts.steps) ? ts.steps : null);
                if (planTasks && planTasks.length) {
                    parts.push('【上一轮任务规划】\n' + planTasks.map(function (t, i) {
                        if (typeof t === 'string') return (i + 1) + '. ' + t;
                        return (i + 1) + '. ' + (t.chartLabel || t.displayName || t.title || '任务') + (t.text ? '：' + t.text : '');
                    }).join('\n'));
                }
            }
            if (parts.length) {
                parts.push('请基于以上上一轮结果，按我的新需求调整。只输出调整后的完整结果。');
                turn.userMessage = (turn.userMessage || '') + '\n\n' + parts.join('\n\n');
            }
        }
        // v20260924：读取端加 schema 守卫，统一字段名兜底（address→range / sheet→sheetName→worksheet）
        // 防止某 pipeline 写入的 selection 字段不全（如原 code-pipeline 精简版）导致静默继承坏数据
        const pd = prevSpec.data || {};
        const ctx = turn.context || {};
        const sel = pd.selection;
        if (!ctx.selection && sel && typeof sel === 'object') {
            const hasAddr = sel.address || sel.range;
            const hasSheet = sel.sheet || sel.sheetName || sel.worksheet;
            if (hasAddr && hasSheet) {
                ctx.selection = sel;
            } else {
                ChatLog.warn?.('PIPELINE', `[text-pipeline] 后续消息继承跳过 selection: 缺少 ${hasAddr ? '' : 'address/range'}${hasSheet ? '' : 'sheet'}`);
            }
        }
        if (!ctx.referenceRegions && Array.isArray(pd.referenceRegions) && pd.referenceRegions.length) {
            ctx.referenceRegions = pd.referenceRegions;
        }
        ChatLog.info?.('PIPELINE', '[text-pipeline] 后续消息继承: 复用上一轮图表结果');
    }

    /**
     * text-pipeline 完成后写入 ExecSpec（Phase 4 生命周期）。
     * 供后续消息继承 —— turn.context.selection/referenceRegions 快照 +
     * taskStructure（chartTasks 数组）让下一轮能定位上轮图表结果继续执行。
     * 注意：写入 pipeline='skill' 而非 'text'（历史原因：text-pipeline 多图表
     * 场景最终产出是图表，语义上归 Skill 域管理）。
     * v20260924：data 统一走 sessionStore.execSpecFactory.buildCommonData(turn)，
     * 保证 selection/referenceRegions schema 与其他 pipeline 一致。
     * @param {Object} turn - 已完成的 turn
     * @param {Object} sessionStore
     * @param {Object} result - 执行结果（含 summary/chartTasks/taskStructure）
     */
    function writeChartExecSpec(turn, sessionStore, result) {
        if (!sessionStore || typeof sessionStore.setExecSpec !== 'function') return;
        // v20260924：用工厂统一构建 data，消除 selection/referenceRegions 字段漂移
        const fact = sessionStore.execSpecFactory;
        const commonData = fact && typeof fact.buildCommonData === 'function'
            ? fact.buildCommonData(turn)
            : {}; // 兜底（旧 session-store 版本）
        const skill = turn.context?._matchedSkill || null;
        sessionStore.setExecSpec(turn.id, {
            pipeline: 'skill',
            data: commonData,
            params: {
                skillName: (skill && (skill.name || skill.Name)) || '',
                displayName: (skill && (skill.displayName || skill.DisplayName)) || '',
                mode: 'text'
            },
            result: result || {},
            // 任务结构（规划步骤/任务清单）：供后续消息继承时注入规划输入
            taskStructure: (result && result.taskStructure) || null,
            ts: Date.now()
        });
    }

    /**
     * 公共数据源确立 —— 三段降级，规划阶段一次全程共享。
     * 为何要"锁顶层数据源"：多图表子任务循环中聚合会新建"聚合结果"工作表，
     * 如果后续建图时不锁数据源会读到"聚合结果"而不是原始数据 → 所有图都基于错误数据。
     * 降级顺序：turn.context.selection（上下文收集）→ bridge.GetSelection → GetSelectedRangeData
     * 三段都失败时 resolve=false，子任务显式看到"数据源未明"而非静默读错表。
     * @param {Object} turn - turn 对象（含 context.selection）
     * @returns {Promise<{sheetName:string, rangeAddr:string, resolve:boolean}>} 数据源对象；
     *   resolve=true 时已同步写入 turn.context.publicSource 供跨管道复用
     */
    async function resolveTopDataSource(turn) {
        const sourceData = { sheetName: '', rangeAddr: '', resolve: false };
        const ctxSel = turn?.context?.selection;
        if (ctxSel?.sheetName && ctxSel?.address) {
            sourceData.sheetName = String(ctxSel.sheetName).trim();
            sourceData.rangeAddr = String(ctxSel.address).trim();
        }
        // 降级1：bridge.GetSelection()
        if (!sourceData.sheetName || !sourceData.rangeAddr) {
            try {
                const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() ||
                    (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
                if (bridge && typeof bridge.GetSelection === 'function') {
                    const selRaw = await bridge.GetSelection();
                    const sel = typeof selRaw === 'string' ? JSON.parse(selRaw) : selRaw;
                    if (sel && sel.address) {
                        if (!sourceData.sheetName) sourceData.sheetName = String(sel.sheetName || sel.worksheetName || '').trim();
                        if (!sourceData.rangeAddr) sourceData.rangeAddr = String(sel.address || '').trim();
                    }
                }
            } catch (_) { /* bridge 不可用时静默——下一级 GetSelectedRangeData 接管 */ }
        }
        // 降级2：bridge.GetSelectedRangeData()
        const bridge2 = window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge;
        if ((!sourceData.sheetName || !sourceData.rangeAddr) && bridge2 && typeof bridge2.GetSelectedRangeData === 'function') {
            try {
                const raw = await bridge2.GetSelectedRangeData();
                const data = typeof raw === 'string' ? JSON.parse(raw) : raw;
                if (data && data.success) {
                    if (!sourceData.sheetName) sourceData.sheetName = String(data.sheetName || '').trim();
                    if (!sourceData.rangeAddr) sourceData.rangeAddr = String(data.address || '').trim();
                }
            } catch (_) { /* 最后一级也失败则 sourceData.resolve=false */ }
        }
        // 规范化 rangeAddr：去除工作表限定与绝对引用
        if (sourceData.rangeAddr) {
            let ra = String(sourceData.rangeAddr).trim();
            if (ra.includes('!')) ra = ra.split('!').pop().trim();
            ra = ra.replace(/\$/g, '');
            sourceData.rangeAddr = ra;
        }
        sourceData.resolve = !!(sourceData.sheetName && sourceData.rangeAddr);
        // 落库到 turn.context.publicSource，供本 turn 后续所有子任务及跨管道复用
        if (sourceData.resolve) {
            if (!turn.context) turn.context = {};
            turn.context.publicSource = {
                sheetName: sourceData.sheetName,
                rangeAddr: sourceData.rangeAddr,
                resolve: true,
                scope: 'turn'
            };
        }
        ChatLog.info?.('PIPELINE',
            `[text-pipeline] 顶层公共数据源: resolve=${sourceData.resolve}, sheet=${sourceData.sheetName || '(空)'}, range=${sourceData.rangeAddr || '(空)'}`);
        return sourceData;
    }

    /**
     * 执行前工具准备 —— 三步流水线：过滤 → 瘦身 → messages。
     * 1) 剔除 OCR 工具（text-pipeline 只管纯文本，走 image-pipeline）
     * 2) 工具注入瘦身（按消息意图预筛，123 → ~40，解决 tool tokens 7391 打爆预算）
     * 3) 构建 LLM 历史 messages
     * @param {Object} turn
     * @param {Object} deps - 含 promptBuilder/toolRegistry/sessionStore/taskPlanner
     * @returns {{tools:Array, messages:Array, userMsg:string, matchedSkill:Object|null,
     *            promptOptions:Object}}
     */
    /**
     * 收集最近若干轮真实执行过的工具名（供"上一轮工具粘性"使用）。
     * 只读 sessionStore 历史，不依赖任何业务字段；任何异常都忽略——瘦身失败不应影响主流程。
     * @param {Object} sessionStore
     * @param {Object} turn - 当前 turn（跳过自身）
     * @param {number} maxTurns - 最多回溯多少轮
     * @returns {string[]} 工具名（小写、去重、最近优先）
     */
    function collectRecentToolNames(sessionStore, turn, maxTurns) {
        const out = [];
        try {
            const turns = (sessionStore && typeof sessionStore.getTurns === 'function') ? sessionStore.getTurns() : [];
            let scanned = 0;
            for (let i = turns.length - 1; i >= 0 && scanned < maxTurns; i--) {
                const t = turns[i];
                if (!t || (turn && t.id === turn.id)) continue;
                scanned++;
                const calls = Array.isArray(t.toolCalls) ? t.toolCalls : [];
                for (const c of calls) {
                    const n = String((c && (c.name || c.toolName)) || '').toLowerCase();
                    if (n && out.indexOf(n) === -1) out.push(n);
                }
            }
        } catch (_e) { /* 忽略：仅为优化可见性，失败则退回按当前消息裁剪 */ }
        return out;
    }

    /** 看板交付物：excel | html | dual（与 skillSpecEngine.dashboardSpecFromMessage 对齐） */
    function resolveDashboardDeliverable(turn) {
        const msg = String((turn && (turn._originalUserMessage || turn.userMessage || turn.message)) || '');
        try {
            const eng = App.ChatV2 && App.ChatV2.services && App.ChatV2.services.skillSpecEngine;
            if (eng && typeof eng.dashboardSpecFromMessage === 'function') {
                const d = String((eng.dashboardSpecFromMessage(msg) || {}).deliverable || 'dual').toLowerCase();
                if (d === 'excel' || d === 'html' || d === 'dual') return d;
            }
        } catch (_e) { /* ignore */ }
        if (/(excel\s*(内)?\s*(看板|仪表盘|驾驶舱)|表格看板|工作表(里的)?看板)/i.test(msg)) return 'excel';
        if (/(网页|浏览器|html|投屏|只要大屏|仅大屏)/i.test(msg) && !/excel/i.test(msg)) return 'html';
        return 'dual';
    }

    /**
     * 无 Skill 时的看板确定性施工：直接调 excel_generate_dashboard，不交给弱模型拆步。
     * 实测「生成Excel看板」走 agent-loop 时模型只读数 + 重复 layout_snapshot + display_html，
     * 从不建 Excel 看板表——一键工具必须在循环前硬跑。
     * @returns {Promise<{ok:boolean, result?:Object, error?:string}|null>} 非看板意图返回 null
     */
    async function maybeRunOneClickDashboard(turn, deps, signal) {
        const dashFlag = !!(turn && ((turn.context && turn.context._forceDashboard) || turn._forceDashboard));
        if (!dashFlag) return null;
        if (turn._oneClickDashboardDone) return turn._oneClickDashboardResult || { ok: true };
        const deliverable = resolveDashboardDeliverable(turn);
        turn.context = turn.context || {};
        turn.context._dashboardDeliverable = deliverable;
        // 只要网页大屏 → 不跑 Excel 一键，留给 maybeRenderHtmlDashboard
        if (deliverable === 'html') {
            ChatLog.info?.('PIPELINE', '[one-click-dash] deliverable=html → 跳过 Excel 一键，交给 HTML 大屏');
            return null;
        }
        const toolRegistry = (deps && deps.toolRegistry) || (App.ChatV2 && App.ChatV2.toolRegistry);
        const oneClick = App.ChatV2 && App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.oneClick;
        if (!toolRegistry || typeof toolRegistry.invokeTracked !== 'function') {
            ChatLog.warn?.('PIPELINE', '[one-click-dash] toolRegistry 不可用');
            return null;
        }
        // 优先走已注册工具（有卡片/进度）；回退直接调 oneClick.generateDashboard
        const hasTool = !!(toolRegistry.get && toolRegistry.get('excel_generate_dashboard'));
        if (!hasTool && !(oneClick && typeof oneClick.generateDashboard === 'function')) {
            ChatLog.warn?.('PIPELINE', '[one-click-dash] excel_generate_dashboard 未注册且 oneClick 不可用');
            return null;
        }
        const sel = turn.context && turn.context.selection;
        const args = {
            board: '看板',
            title: '数据看板'
        };
        if (sel && (sel.sheetName || sel.worksheet)) args.sourceSheet = sel.sheetName || sel.worksheet;
        if (sel && (sel.address || sel.range)) args.sourceRange = String(sel.address || sel.range).replace(/\$/g, '');
        // 生成话术里带主题色（如「生成红色看板」）→ 直接用对应主题
        try {
            const msg0 = String(turn.originalUserMessage || turn.message || '');
            const eng0 = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.blueprintEngine;
            const tid = eng0 && typeof eng0.resolveThemeFromMessage === 'function' ? eng0.resolveThemeFromMessage(msg0) : null;
            if (tid) args.theme = tid;
        } catch (_e) { /* noop */ }
        ChatLog.info?.('PIPELINE', `[one-click-dash] 确定性执行 excel_generate_dashboard deliverable=${deliverable} sheet=${args.sourceSheet || ''} range=${args.sourceRange || ''} theme=${args.theme || ''}`);
        let res = null;
        try {
            if (hasTool) {
                res = await toolRegistry.invokeTracked(turn, 'excel_generate_dashboard', args, {
                    round: 1,
                    displayName: '一键生成看板',
                    ctx: { turn, signal }
                });
            } else {
                res = await oneClick.generateDashboard(args, { turn, signal, sessionStore: deps && deps.sessionStore });
            }
        } catch (e) {
            res = { ok: false, error: (e && e.message) || String(e) };
        }
        turn._oneClickDashboardDone = true;
        turn._oneClickDashboardResult = res;
        if (res && res.ok) {
            turn._chartDone = true; // 防止图表闸门再补链
            if (!Array.isArray(turn._executedTools)) turn._executedTools = [];
            turn._executedTools.push('excel_generate_dashboard');
        }
        ChatLog.info?.('PIPELINE', `[one-click-dash] 结果 ok=${!!(res && res.ok)} ${res && res.error ? res.error : ''}`);
        return res;
    }

    /**
     * 「修改为红色主题」等改配色意图：确定性换主题重建看板，不经弱模型空转。
     */
    async function maybeRunDashboardThemeChange(turn, deps, signal) {
        if (turn._dashboardThemeChangeDone) return turn._dashboardThemeChangeResult || null;
        const msg = String(
            (turn && (turn.originalUserMessage || turn._originalUserMessage || turn.userText || turn.userMessage || turn.message)) || ''
        );
        const engine = App.ChatV2 && App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.blueprintEngine;
        const oneClick = App.ChatV2 && App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.oneClick;
        if (!engine || typeof engine.isThemeChangeIntent !== 'function' || !engine.isThemeChangeIntent(msg)) {
            return null;
        }
        const themeId = engine.resolveThemeFromMessage(msg);
        if (!themeId) {
            ChatLog.warn?.('PIPELINE', '[dash-theme] 未能从话术解析主题: ' + msg.slice(0, 40));
            return null;
        }
        const toolRegistry = (deps && deps.toolRegistry) || (App.ChatV2 && App.ChatV2.toolRegistry);
        if (!oneClick || typeof oneClick.changeDashboardTheme !== 'function') {
            ChatLog.warn?.('PIPELINE', '[dash-theme] changeDashboardTheme 不可用');
            return null;
        }
        const sel = turn.context && turn.context.selection;
        const last = App.ChatV2._lastDashboardSource || {};
        const args = {
            theme: themeId,
            board: last.board || '看板',
            title: '数据看板'
        };
        if (sel && (sel.sheetName || sel.worksheet) && String(sel.sheetName || sel.worksheet) !== '看板') {
            args.sourceSheet = sel.sheetName || sel.worksheet;
            if (sel.address || sel.range) args.sourceRange = String(sel.address || sel.range).replace(/\$/g, '');
        } else if (last.sheet) {
            args.sourceSheet = last.sheet;
            if (last.range) args.sourceRange = last.range;
        }
        // 跨轮无记忆时：从会话历史挖「销售数据!A1:I101」类来源
        if (!args.sourceSheet) {
            try {
                const hist = (deps && deps.sessionStore && typeof deps.sessionStore.getHistoryForLLM === 'function')
                    ? deps.sessionStore.getHistoryForLLM() : [];
                for (let hi = hist.length - 1; hi >= 0; hi--) {
                    const blob = String((hist[hi] && hist[hi].content) || '');
                    const m = blob.match(/([^\s!！「」【】]{1,40})!(A\d{1,7}:[A-Z]{1,3}\d{1,7})/i);
                    if (m && m[1] && m[1] !== '看板' && m[1] !== '_dash_pivot') {
                        args.sourceSheet = m[1];
                        args.sourceRange = m[2];
                        ChatLog.info?.('PIPELINE', `[dash-theme] 从历史恢复源 ${args.sourceSheet}!${args.sourceRange}`);
                        break;
                    }
                }
            } catch (_e) { /* noop */ }
        }
        // 仍无源：试探工作簿里非看板表
        if (!args.sourceSheet && toolRegistry && typeof toolRegistry.invokeTracked === 'function') {
            try {
                const info = await toolRegistry.invokeTracked(turn, 'excel_get_workbook_info', {}, {
                    round: 0, displayName: '探测工作簿', ctx: { turn, signal }
                });
                const text = typeof info?.result === 'string' ? info.result : JSON.stringify(info?.result || '');
                const names = [];
                const re = /"name"\s*:\s*"([^"]+)"/g;
                let mm;
                while ((mm = re.exec(text))) names.push(mm[1]);
                const cand = names.find(n => n && n !== '看板' && n !== '_dash_pivot' && n !== '透视缓存');
                if (cand) {
                    args.sourceSheet = cand;
                    ChatLog.info?.('PIPELINE', `[dash-theme] 探测到数据表 ${cand}`);
                }
            } catch (_e) { /* noop */ }
        }
        if (!args.sourceSheet && !args.sourceRange) {
            ChatLog.warn?.('PIPELINE', '[dash-theme] 无数据源（请先生成看板或选中源数据区）');
            return null;
        }
        ChatLog.info?.('PIPELINE', `[dash-theme] 确定性换主题 theme=${themeId} src=${args.sourceSheet || ''}!${args.sourceRange || ''}`);
        let res = null;
        try {
            res = await oneClick.changeDashboardTheme(args, {
                turn, signal, sessionStore: deps && deps.sessionStore
            });
        } catch (e) {
            res = { ok: false, error: (e && e.message) || String(e) };
        }
        turn._dashboardThemeChangeDone = true;
        turn._dashboardThemeChangeResult = res;
        if (res && res.ok) {
            turn._chartDone = true;
            turn._oneClickDashboardDone = true; // 阻止后续重复一键/美化
            if (!Array.isArray(turn._executedTools)) turn._executedTools = [];
            turn._executedTools.push('excel_generate_dashboard');
        }
        return res;
    }

    function prepareTools(turn, deps) {
        const { promptBuilder, toolRegistry, sessionStore } = deps;
        let tools = toolRegistry.listBySource('mcp');
        // 看板一键编排（source=builtin-dashboard）必须并入可见工具，否则 prompt 里写了
        // excel_generate_dashboard 但模型手里没有 → 只能乱调 read/layout/display_html。
        try {
            const dashTools = typeof toolRegistry.listBySource === 'function'
                ? (toolRegistry.listBySource('builtin-dashboard') || []) : [];
            if (dashTools.length) {
                const seen = new Set(tools.map(t => String(t && t.name || '').toLowerCase()));
                for (const t of dashTools) {
                    const n = String(t && t.name || '').toLowerCase();
                    if (n && !seen.has(n)) { tools.push(t); seen.add(n); }
                }
            }
        } catch (_e) { /* ignore */ }

        // 剔除图片类（OCR）工具——text-pipeline 只处理纯文本
        const originalCount = tools.length;
        tools = tools.filter(t => !toolRegistry.isImageOnlyTool(t));
        if (tools.length !== originalCount) {
            ChatLog.info?.('PIPELINE', `[text-pipeline] 已从工具列表剔除图片类（OCR）工具（text-pipeline只处理纯文本，图片请走image-pipeline）`);
        }

        const userMsg = turn._originalUserMessage || turn.userMessage || turn.message || '';

        // 工具注入瘦身：按用户消息意图预筛工具，只保留核心基本盘 + 意图命中的工具族
        const toolsBeforeFilter = tools.length;
        try {
            const filterFn = App.ChatV2.toolRegistryUtil?.filterToolsForMessage;
            if (typeof filterFn === 'function') {
                // 【宿主感知·按宿主喂工具】filterToolsForMessage 的核心基本盘（CORE_TOOLS）与
                // 意图词表（INTENT_RULES）几乎全是 excel_* 专属——Word/PPT 宿主下消息不带
                // Excel 关键词时 kept=0 只能靠"回退全量"侥幸活命，一旦撞上某个 Excel 意图词
                // 反而会把 word_*/ppt_* 工具错误裁剪掉（实测 Word 问"列出工具名称"，模型
                // 手里只剩 2 个读取工具可汇报）。而 word/ppt 工具总量小（20-30 个），
                // 瘦身本无收益——非 Excel 宿主直接保持全量。
                const _hHost = (typeof window !== 'undefined' && window.SmartPilotHost) || null;
                if (_hHost && !_hHost.isExcel) {
                    ChatLog.info?.('PIPELINE', `[text-pipeline] 非 Excel 宿主（${_hHost.kind}）跳过工具瘦身，保持全量 ${tools.length} 个`);
                } else {
                    // 【上一轮工具粘性】瘦身原先只吃当前这一句用户消息，而跟进型消息（"不够美观"/"继续"）
                    // 往往一个意图词都不含 → 用户正在用的工具体系（如格式化族）被整体剔除，
                    // 模型只能回"接下来我将使用 excel_format_range…"的文字计划（实测失败）。
                    const preferTools = collectRecentToolNames(sessionStore, turn, 3);
                    tools = filterFn(tools, userMsg, { preferTools });
                    if (tools.length !== toolsBeforeFilter) {
                        const sticky = preferTools.length > 0 ? `｜粘性保留 ${preferTools.length} 个: ${preferTools.join(',')}` : '';
                        ChatLog.info?.('PIPELINE', `[text-pipeline] 工具注入瘦身: ${toolsBeforeFilter} → ${tools.length}（消息=${userMsg.slice(0, 30)}）${sticky}`);
                    }
                }
            }
        } catch (filterErr) {
            ChatLog.warn?.('PIPELINE', '[text-pipeline] 工具瘦身失败（不影响主流程）:', filterErr?.message);
        }

        // 看板意图：剔除 layout_snapshot（一键编排已锚定布局；留给模型只会连刷烧轮次）
        if (turn && turn.context && turn.context._forceDashboard) {
            const beforeSnap = tools.length;
            tools = tools.filter(t => !/get_layout_snapshot/i.test(String((t && (t.name || t.originalName)) || '')));
            if (tools.length !== beforeSnap) {
                ChatLog.info?.('PIPELINE', `[text-pipeline] 看板意图 → 剔除 layout_snapshot: ${beforeSnap} → ${tools.length}`);
            }
        }

        // 【纯格式化意图：剔除读数工具】格式化/美化只需要"目标范围"，不需要知道单元格里的值。
        // excel_read_range 一直被 filterToolsForMessage 的 CORE_TOOLS 无条件保留，
        // 于是即便 OPERATION_ONLY_TASK_PROMPT 明写"不要先调 excel_read_range"，弱模型仍会先读整表，
        // 再顺手输出一段"根据提供的数据统计…"——把格式化请求做成数据分析总结（实测截图：
        // 用户只要"设置表头底色/边框/对齐/列宽"，回的却是全表维度统计 + 数据总结）。
        // 提示词约束对小模型不可靠，这里直接从工具表剔除读数族，从源头杜绝。
        // 范围解析不受影响：OPERATION_ONLY_TASK_PROMPT 已指示"优先用选区，否则不传 range 走 UsedRange"，
        // 且 excel_get_used_range / 选区表头仍在，足以确定作用范围。
        try {
            if (promptBuilder && typeof promptBuilder.detectOperationOnlyIntent === 'function'
                && promptBuilder.detectOperationOnlyIntent(userMsg)) {
                const beforeNoRead = tools.length;
                const noRead = tools.filter(t => {
                    const n = String((t && (t.name || t.originalName)) || '').toLowerCase();
                    return !/read_range|read_sheet|ocr/i.test(n);
                });
                // 防御：剔除后不应为空（格式化族由 filterToolsForMessage 的 INTENT_RULES 保证保留），
                // 万一为空则放弃本次剔除，绝不让模型陷入"无工具可用"。
                if (noRead.length > 0) tools = noRead;
                if (tools.length !== beforeNoRead) {
                    ChatLog.info?.('PIPELINE', `[text-pipeline] 纯格式化意图 → 剔除读数工具: ${beforeNoRead} → ${tools.length}（消息=${userMsg.slice(0, 30)}）`);
                }
            }
        } catch (noReadErr) {
            ChatLog.warn?.('PIPELINE', '[text-pipeline] 剔除读数工具失败（不影响主流程）:', noReadErr?.message);
        }

        // 有文档附件时，read_attachment 与 excel_read_range 一样是可读步骤，瘦身不能把它裁掉。
        try {
            const pin = App.ChatV2.attachmentReadTool && App.ChatV2.attachmentReadTool.pinOn;
            if (typeof pin === 'function') tools = pin(tools, turn);
        } catch (_pinErr) { /* 附件工具缺失时仍用原工具表 */ }

        const matchedSkill = turn.context?._matchedSkill || null;
        let promptOptions = {};
        if (matchedSkill) {
            const persona = matchedSkill.persona || matchedSkill.Persona;
            if (persona) {
                promptOptions.persona = persona;
                ChatLog.info?.('PIPELINE', `[text-pipeline] 使用 Skill 身份: ${persona}`);
            }
        }

        const history = sessionStore.getHistoryForLLM();
        const messages = promptBuilder.buildMessages(turn, history, tools, promptOptions);

        ChatLog.info?.('PIPELINE', `[text-pipeline] turn=${turn.id} tools=${tools.length} history=${history.length}`);

        return { tools, messages, userMsg, matchedSkill, promptOptions };
    }

    /**
     * 异步注入"强匹配 Skill 工具"到 LLM 可见工具列表。
     * 场景：orchestrator 未路由到 Skill 但消息里强匹配（score>=20）某个 skill_* 工具名，
     * 把它塞回 tools 数组防止 LLM 因找不到工具而幻觉。
     * 独立成 async 是因为 agentLoop.findCandidateTools 在不同版本签名可能是 async。
     * @param {Object} turn
     * @param {Array} tools - 上一步 prepareTools 产出的工具数组
     * @param {Object} deps - 含 agentLoop/toolRegistry
     * @param {Object|null} matchedSkill - 已匹配 Skill 时跳过（Skill 自己有完整工具集）
     * @returns {Promise<Array>} 注入后的 tools 数组（可能 +1 skill_* 工具）
     */
    async function injectStrongSkillCandidates(turn, tools, deps, matchedSkill) {
        const { agentLoop, toolRegistry } = deps;
        if (matchedSkill || !agentLoop || typeof agentLoop.findCandidateTools !== 'function' || !tools.length) {
            return tools;
        }
        try {
            const userMsgForMatch = turn?.userMessage || turn?.message || '';
            const skillCandidates = await agentLoop.findCandidateTools(userMsgForMatch, toolRegistry, 3);
            const strongSkill = skillCandidates.find(c =>
                c.score >= 20 && c.tool && String(c.tool.name).startsWith('skill_'));
            if (strongSkill?.tool) {
                const skillTool = toolRegistry.get(strongSkill.tool.name);
                if (skillTool) {
                    // 【优先级修复】原来是 append → 排在第 17 位，与 16 个通用工具完全平权。
                    // 实测（09:36/09:37 两轮宿主日志）skill_SmartTableBeautifier 确实被注入了，
                    // 但模型照样挑 excel_format_range，Skill 从未被调用。
                    // 强匹配（score>=20）意味着"这个 Skill 就是用户要的东西"→ 置顶，让它在工具表里第一眼可见。
                    tools = [skillTool, ...tools];
                    ChatLog.info?.('PIPELINE', `[text-pipeline] 强匹配 skill 已置顶到 LLM 工具列表首位: ${strongSkill.tool.name} (score=${strongSkill.score})`);
                }
            }
        } catch (matchErr) {
            ChatLog.warn?.('PIPELINE', '[text-pipeline] skill 匹配增强失败（不影响主流程）:', matchErr?.message);
        }
        return tools;
    }

    /**
     * 判定是否应该走多步骤管道（agent-loop 循环 + 步骤骨架实时填色）。
     * 判定顺序（优先级从高到低）：
     *   1) turn.context._forceMultiStep=true → 直接 true（LLM 意图分析已判定）
     *   2) turn.context._intentSingleOp=true → 直接 false（单操作/问答场景）
     *   3) TaskPlanner.needPlanning(userMsg) → 同步正则判定
     *      - true  → 硬判定走多步骤
     *      - 'needs_llm_check' → 委托 needPlanningLLM 让 LLM 再判一次（regex 边界模糊的场景）
     *      - false → 跳过
     * 为何不合并到 prepareTools：prepareTools 在 execute 阶段 2 跑，decideMultiStep
     *   在 execute 阶段 4 分发决策前跑，是"策略层"，prepareTools 是"准备层"。
     * @param {Object} turn
     * @param {string} userMsg
     * @param {Object} deps - 透传到 App.ChatV2 取 TaskPlanner/ProgressTracker/llmIo
     * @returns {Promise<boolean>} true=进入 runMultiStepPipeline，false=跳过
     */
    async function decideMultiStep(turn, userMsg, deps) {
        const { TaskPlanner, ProgressTracker, llmIo } = App.ChatV2;
        if (!TaskPlanner || !ProgressTracker) return false;

        if (turn.context?._forceMultiStep) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] LLM 意图分析已判定需多步骤，直接进入多步骤管道');
            return true;
        }
        if (turn.context?._forceDashboard) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] 强制看板意图 → 直接进入多步骤管道（一键施工）');
            return true;
        }
        if (turn.context?._forceChart) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] 强制图表意图 → 直接进入多步骤管道');
            return true;
        }
        if (turn.context?._intentSingleOp) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] LLM 意图分析已判定单操作/问答，跳过多步骤管道');
            return false;
        }

        // 纯格式化/样式命令（把表头加粗、字体改红、加边框、设置对齐…）不应进入多步骤规划器，
        // 否则规划器会把"加粗表头"过度编排成"聚合+画图"等脏动作（见 2026-10 单格选区 999 格问题）。
        // 交给 dispatcher 确定性直调或单步 agent 即可；含分析/复合操作意图时不拦截。
        const PURE_FMT_RE = /(加粗|加斜|倾斜|斜体|居中|边框|字体颜色|字体|字号|填充|背景色|底色|高亮|行高|列宽|自动换行|清除格式|设置格式|格式化|单元格格式|对齐|表头|首行|标题行)/;
        const ANALYSIS_RE = /(汇总|聚合|统计|分析|图表|画图|绘制|透视|对比|环比|同比|占比|趋势|分组|分类汇总|明细|洞察)/;
        const COMPOSITE_RE = /(复制|写回|写入|新建|粘贴|导出|移动|追加|拆分|合并|转换|生成|插入|删除行|删除列)/;
        if (PURE_FMT_RE.test(userMsg) && !ANALYSIS_RE.test(userMsg) && !COMPOSITE_RE.test(userMsg)) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] 纯格式化命令，跳过多步骤管道');
            return false;
        }

        const planCheck = TaskPlanner.needPlanning(userMsg);
        if (planCheck === true) return true;
        if (planCheck === 'needs_llm_check') {
            const askLLM = async (prompt) => {
                try {
                    if (!llmIo || typeof llmIo.askText !== 'function') return 'NO';
                    const res = await llmIo.askText({
                        messages: [{ role: 'user', content: prompt }],
                        config: App.config,
                        signal: turn.context?._execFollowUp ? undefined : undefined
                    });
                    return res?.text || 'NO';
                } catch (_) { return 'NO'; }
            };
            return await TaskPlanner.needPlanningLLM(userMsg, askLLM);
        }
        return false;
    }

    /**
     * 多步骤管道（async generator）—— agent-loop 事件驱动的步骤面板。
     *
     * 设计原则：
     *   - 零骨架：面板内容完全由 agent-loop 事件自然生长，不预先规划、不预占位
     *   - 每个工具调用 = 一个步骤卡片：标题 = 真实工具名，detail = 结果摘要
     *   - round_start / round_end 静默处理（迭代是循环内部逻辑，不暴露给用户）
     *   - 每次 push 新步骤时整块 reset（简单可靠，面板不会闪烁）
     *
     * @param {Object} turn - context 需已含 _llmTools + _llmMessages
     * @param {AbortSignal} signal
     * @param {Object} deps - App.ChatV2
     * @yields {Object} agent-loop 事件流 + {type:'pipeline_done'}
     */
    async function* runMultiStepPipeline(turn, signal, deps) {
        const { agentLoop, sessionStore } = deps;

        const messages = turn.context?._llmMessages || null;
        if (!messages || !agentLoop || typeof agentLoop.run !== 'function') {
            ChatLog.warn?.('PIPELINE', '[text-pipeline] runMultiStepPipeline 缺少 messages/agentLoop，走兜底 agent-loop');
            yield* runAgentLoopFallback(turn, signal, deps, messages);
            return;
        }

        ChatLog.info?.('PIPELINE', '[text-pipeline] 命中多步骤管道（事件驱动 · 零骨架）');
        sessionStore.updateTurn(turn.id, { _enableSteps: true });

        const stepsPanel = App.ChatV2.stepsPanel
            || (App.ChatV2.skillPipelineModules && App.ChatV2.skillPipelineModules.stepsPanel)
            || null;
        if (stepsPanel && !App.ChatV2.stepsPanel) App.ChatV2.stepsPanel = stepsPanel;

        // 【看板·改主题】「修改为红色主题」等 → 确定性换主题，避免弱模型只规划不调工具
        const themeChange = await maybeRunDashboardThemeChange(turn, deps, signal);
        if (themeChange && themeChange.ok) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] 看板改主题成功 → 跳过 agent-loop');
            const msgT = (themeChange.result && themeChange.result.message) || '已切换看板主题。';
            sessionStore.completeTurn(turn.id, { content: '✅ ' + msgT });
            yield { type: 'pipeline_done', pipeline: 'text' };
            return;
        }

        // 【看板·确定性一键】_forceDashboard 时先硬跑 excel_generate_dashboard，成功则跳过 agent-loop
        // （避免弱模型只读数 + 刷 layout_snapshot + 只开 HTML、从不建 Excel 看板表）。
        const oneClick = await maybeRunOneClickDashboard(turn, deps, signal);
        if (oneClick && oneClick.ok) {
            ChatLog.info?.('PIPELINE', '[text-pipeline] 看板一键成功 → 跳过 agent-loop，进入收尾合成');
            const finalTurn = sessionStore.getCurrentTurn() || turn;
            const deliverable = (turn.context && turn.context._dashboardDeliverable) || resolveDashboardDeliverable(turn);
            let htmlDash = null;
            if (deliverable !== 'excel') {
                htmlDash = await maybeRenderHtmlDashboard(turn, finalTurn, deps, signal);
            }
            const msg = (oneClick.result && oneClick.result.message)
                || `已生成 Excel 看板工作表「${(oneClick.result && oneClick.result.board) || '看板'}」。`;
            let finalContent = '✅ ' + msg;
            if (htmlDash && htmlDash.ok) {
                finalContent += '\n\n🖥️ **HTML 大屏已生成并在浏览器打开**' +
                    (htmlDash.filePath ? `（${htmlDash.filePath}）` : '') + '。';
            } else if (deliverable === 'excel') {
                finalContent += '\n\n（按你的要求仅产出 Excel 内看板，未打开网页大屏。）';
            }
            sessionStore.completeTurn(turn.id, { content: finalContent });
            yield { type: 'pipeline_done', pipeline: 'text' };
            return;
        }

        const tools = turn.context?._llmTools || [];
        const MAX_ROUNDS = (agentLoop && agentLoop.MAX_TOOL_ROUNDS) || 5; // 引用 agent-loop 单一来源，消除轮次漂移（原硬编码 8）

        // 初始骨架为空——面板等第一个 tool_call_result 出现时再 lazy 出
        let firstEvent = true;

        // ── 事件拦截 ──
        for await (const evt of agentLoop.run({
            turn, messages, tools, signal,
            maxIterations: MAX_ROUNDS,
            persistTools: true
        })) {
            // round_start：静默——迭代是循环内部逻辑，不做 UI 变化
            if (evt.type === 'round_start') {
                ChatLog.info?.('PIPELINE', `[text-pipeline] round_start ${evt.round}（静默）`);
            }

            // tool_call_result：每来一个工具结果 → push 一个新步骤卡片
            if (evt.type === 'tool_call_result') {
                const toolName = evt.toolName || '工具';
                const ok = !!evt.ok;
                const detail = ok
                    ? `${toolName} ✓${evt.durationMs ? `（${evt.durationMs}ms）` : ''}`
                    : `${toolName} ✕${evt.error ? '：' + String(evt.error).slice(0, 80) : ''}`;
                const step = {
                    id: `tp_${Date.now()}_${Math.random().toString(36).slice(2, 6)}`,
                    title: toolName,
                    description: toolName,
                    detail,
                    status: ok ? 'completed' : 'failed',
                    strategy: 'agent',
                    toolName,
                    result: evt.result || null
                };
                if (!Array.isArray(turn._skillSteps)) turn._skillSteps = [];
                turn._skillSteps.push(step);
                if (stepsPanel && typeof stepsPanel.initSkillStepsPanel === 'function') {
                    if (firstEvent) {
                        // 第一次：整块 reset 出面板
                        stepsPanel.initSkillStepsPanel(turn, sessionStore, turn._skillSteps);
                        firstEvent = false;
                    } else {
                        // 后续：整块 reset（简单可靠，面板不会闪烁）
                        stepsPanel.renderSkillStepsPatch(turn, sessionStore, [], 'reset');
                    }
                }
            }

            // round_end：静默
            if (evt.type === 'round_end') {
                // noop
            }

            yield evt;
        }

        // ── 收尾 ──
        if (stepsPanel && typeof stepsPanel.finalizeSkillSteps === 'function') {
            stepsPanel.finalizeSkillSteps(turn, sessionStore);
        }

        // ── 最终内容合成（与 runAgentLoopFallback 同规则）──
        const finalTurn = sessionStore.getCurrentTurn();
        if (finalTurn && finalTurn.id === turn.id) {
            // 【无 Skill 看板】确定性合成 HTML 大屏并打开浏览器（不依赖 LLM，放在内容合成前）
            const htmlDash = await maybeRenderHtmlDashboard(turn, finalTurn, deps, signal);
            // 【看板美化】标题横幅 + KPI 卡片配色 + 图表区白底细边框 + 统一列宽
            const beautify = await maybeBeautifyDashboard(turn, finalTurn, deps, signal);
            const rawUserMsg = (turn && (turn._originalUserMessage || turn.userMessage)) || '';
            const preferSummary = typeof deps.promptBuilder?.detectAnalysisIntent === 'function'
                ? deps.promptBuilder.detectAnalysisIntent(rawUserMsg) : false;
            const llmText = finalTurn?.content || '';
            const rawToolResult = buildRawToolResultText(finalTurn);
            // 永远优先 LLM 总结文本（比原始工具结果 JSON 可读 100 倍）；仅当 LLM 总结为空时才 fallback rawToolResult
            let finalContent = textWhenModelSilent(finalTurn, llmText, rawToolResult);
            finalContent = _decorateFinalContent(finalContent, turn, finalTurn);
            // 【空回复兜底】模型返回空内容（流中断/模型冷启动未就绪/纯空白输出）时，
            // 决不能让用户面对空气泡——给诚实报错 + 重试指引。
            if (!String(finalContent || '').trim()) {
                finalContent = '⚠️ 模型本次未返回任何内容（常见原因：模型服务刚启动还在加载模型、请求超时或返回了空白输出）。\n\n'
                    + '请点击重试；若连续失败，请确认模型服务（如 Ollama）已启动并已加载对应模型。';
                (typeof window !== 'undefined' && window.ChatLog || console).warn?.('PIPELINE', `[text-pipeline] 模型返回空内容 → 已用兜底提示替换（turn=${turn.id}）`);
            }
            if (htmlDash && htmlDash.ok) {
                finalContent += '\n\n🖥️ **HTML 大屏已生成并在浏览器打开**' +
                    (htmlDash.filePath ? `（${htmlDash.filePath}）` : '') +
                    '：含 KPI 卡片、交互图表与维度筛选器，与 Excel「看板」工作表并行产出。';
            }
            if (beautify && beautify.ok) {
                finalContent += `\n\n🎨 **看板已美化**：标题横幅 + KPI 卡片配色 + 图表区边框（${beautify.steps} 项样式已应用）。`;
            }
            sessionStore.completeTurn(turn.id, { content: finalContent, reasoning: finalTurn?.reasoning || '' });
        }
        yield { type: 'pipeline_done', pipeline: 'text' };
    }

    /**
     * 在 finalContent 末尾追加：①产物文件 banner；②inline chart HTML（<div data-echarts>）。
     * 两处 finalContent 合成（主管道 + runAgentLoopFallback）共用此函数。
     * @param {string} content 初始 finalContent
     * @param {object} turn 当前 turn（含 _aggregateResults / _aggregateClaimed）
     * @param {object} finalTurn sessionStore.getCurrentTurn() 结果（含 toolCalls）
     * @returns {string} 追加后的 content
     */
    function _decorateFinalContent(content, turn, finalTurn) {
        let finalContent = content || '';
        const ChatLog = (typeof window !== 'undefined' ? window.ChatLog : null) || console;

        // ── 产物文件 banner ──
        const filePath = scanFilePathInResults(Array.isArray(finalTurn.toolCalls) ? finalTurn.toolCalls : []);
        if (filePath && finalContent && finalContent.indexOf(filePath) < 0) {
            const h = App.ChatV2?.renderHtml?.escapeHtml
                || (s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'));
            finalContent += `\n\n<div class="multi-step-file-banner" style="margin-top:10px;padding:10px 12px;background:#eef6ff;border:1px solid #bfdbfe;border-radius:8px;font-size:13px;line-height:1.6;color:#1e40af;">📄 <strong>已生成文件：</strong><span style="word-break:break-all;font-family:Consolas,Menlo,monospace;">${h(filePath)}</span></div>`;
        }

        // ── 图表 inline 渲染 ──
        // 从 create_chart 工具调用 + aggregate 结果里取数据，用 buildInlineChartHtml 生成 <div data-echarts>
        const _chartToolNames = /^excel_create_chart(_(pie|bar|line|doughnut|column))?$/;
        const _allAggs = Array.isArray(turn._aggregateResults) && turn._aggregateResults.length
            ? turn._aggregateResults
            : (turn._aggregateResult ? [turn._aggregateResult] : []);
        const _claimed = turn._aggregateClaimed || new Set();
        // DEBUG: 全链路诊断
        ChatLog.info?.('PIPELINE', `[text-pipeline _decorateFinalContent] DEBUG: toolCalls=${Array.isArray(finalTurn.toolCalls)?finalTurn.toolCalls.length:'NO ARRAY'}, aggResults=${turn._aggregateResults?.length||0}, aggResult=${turn._aggregateResult?'YES':'NO'}`);
        if (Array.isArray(finalTurn.toolCalls)) finalTurn.toolCalls.forEach(tc => ChatLog.info?.('PIPELINE', `  tc: name=${tc.name} status=${tc.status}`));
        _allAggs.forEach((a, i) => ChatLog.info?.('PIPELINE', `  agg[${i}]: sheet=${a.sheet} labels=${a.labels?.length||0} values=${a.values?.length||0}`));
        const _chartCalls = Array.isArray(finalTurn.toolCalls)
            ? finalTurn.toolCalls.filter(tc => _chartToolNames.test(tc.name) && tc.status === 'completed')
            : [];

        ChatLog.info?.('PIPELINE', `  chartCalls matched: ${_chartCalls.length}`);        let _chartHtmlAccum = '';
        for (const _cc of _chartCalls) {
            const _args = _cc.args || {};
            const _dsRaw = String(_args.dataRange || _args.range || _args.data_source || '');
            const _dsPrefix = _dsRaw.lastIndexOf('!') > 0
                ? _dsRaw.slice(0, _dsRaw.lastIndexOf('!')).replace(/^['"]|['"]$/g, '')
                : '';
            let _agg = _allAggs.find(a => _dsPrefix && a.sheet === _dsPrefix);
            if (!_agg) _agg = _allAggs.find(a => _claimed.has(a.dataRange));
            if (!_agg) _agg = _allAggs.find(a => a.sheet === String(_args.sheet || ''));
            if (!_agg) _agg = _allAggs[0];
            if (!_agg || !Array.isArray(_agg.labels) || _agg.labels.length < 2) continue;

            const _chartType = _args.chartType || _args.chartEnum || 'PieChart';
            const _title = String(_args.title || _agg.title || '数据图表');
            const _aggData = {
                labels: _agg.labels, values: _agg.values,
                chartType: _chartType, title: _title,
                dimColumn: _agg.dim, valueColumn: _agg.val
            };
            const _html = App.ChatV2?.renderHtml?.buildInlineChartHtml
                ? App.ChatV2.renderHtml.buildInlineChartHtml(_aggData) : '';
            if (_html) _chartHtmlAccum += '\n' + _html;
        }
        if (_chartHtmlAccum) {
            ChatLog.info?.('PIPELINE', `[text-pipeline] 追加 inline chart HTML：${_chartCalls.length} 张图`);
            finalContent += '\n' + _chartHtmlAccum;
        }

        return finalContent;
    }

    // ============================================================
    // 【无 Skill 看板 · HTML 大屏】确定性合成（不经过 LLM）
    // 触发：turn.context._forceDashboard（看板/大屏意图且 DashboardGenerator Skill 不可用）
    // 链路：本轮 excel_read_range 的真实数据 → dashboardBlueprintEngine.computeFromSources
    //      → htmlRenderer.renderHtml（深色主题 KPI 卡 + ECharts 图 + 内置 JS 筛选器）
    //      → excel_display_html_in_browser 在浏览器打开
    // 目的：Excel 看板工作表之外，额外产出真正的"大屏"，且不依赖小模型是否记得调工具。
    // ============================================================

    /** 解析 excel_read_range 的返回（string JSON / {result} / MCP content 数组 多形态兼容） */
    function parseReadRangeResult(raw) {
        function unwrap(o, d) {
            if (!o || d > 6) return null;
            if (typeof o === 'string') {
                try { const p = JSON.parse(o); return unwrap(p, d + 1); } catch (e) { return null; }
            }
            if (typeof o !== 'object') return null;
            if (Array.isArray(o.content)) {
                const texts = o.content.filter(c => c && typeof c.text === 'string').map(c => c.text).join('\n');
                if (texts) { try { const p = JSON.parse(texts); return unwrap(p, d + 1); } catch (e) { return null; } }
            }
            if (Array.isArray(o) && o.length && o[0] && typeof o[0] === 'object' && 'headers' in o[0]) {
                return unwrap(o[0], d + 1);
            }
            if (Array.isArray(o.headers)) return o;
            if (o.result !== undefined && o.result !== null) return unwrap(o.result, d + 1);
            if (o.data !== undefined && o.data !== null) return unwrap(o.data, d + 1);
            return null;
        }
        const o = unwrap(raw, 0);
        if (!o || !Array.isArray(o.headers) || !o.headers.length) return null;
        const rows = Array.isArray(o.result) ? o.result : (Array.isArray(o.rows) ? o.rows : null);
        if (!Array.isArray(rows) || !rows.length) return null;
        return {
            worksheet: o.worksheet || o.sheet || '',
            range: o.range || '',
            headers: o.headers.map(String),
            rows
        };
    }

    /**
     * 看板模式下合成 HTML 大屏并打开浏览器。
     * @returns {Promise<{ok:boolean, filePath:string}|null>} 未触发/跳过时返回 null
     */
    async function maybeRenderHtmlDashboard(turn, finalTurn, deps, signal) {
        const L = (msg) => { try { ChatLog.info?.('PIPELINE', '[html-dash] ' + msg); } catch (e) { /* noop */ } };
        try {
            // 【关键修复】仅 _forceDashboard（真看板意图）才合成 HTML 大屏。
            // 之前把 _forceChart 也算进 dashFlag：补救守卫 D 在 chart skill 不可用时只设 _forceChart，
            // 导致"使用饼图按部门展示销售额"这类简单图表请求除了原生图引导外，
            // 还被额外合成 HTML 大屏并在浏览器打开（"简单图都要走网页"的回归来源）。
            // 看板流程不受影响：守卫 E 检测到看板意图时会同时设 _forceDashboard + _forceChart。
            const dashFlag = !!(turn && ((turn.context && turn.context._forceDashboard) || turn._forceDashboard));
            if (!dashFlag) { L('跳过：非看板意图（无 _forceDashboard）'); return null; }
            // 「生成 Excel 看板」等仅 Excel 意图 → 禁止再弹浏览器（用户明确不要网页）
            const deliverable = (turn.context && turn.context._dashboardDeliverable) || resolveDashboardDeliverable(turn);
            if (deliverable === 'excel') { L('跳过：deliverable=excel（仅 Excel 内看板）'); return null; }
            if (turn._htmlDashboardDone) { L('跳过：本轮已生成过'); return null; }
            // 【修复·双页面】_htmlDashboardDone 只防本函数自身重复，不防"LLM 本轮已经自己调过
            // excel_display_html_in_browser"。实测"帮我生成一个excel看板"：弱模型在工具链末尾
            // 自行调用 display_html（✓ 5861ms，浏览器已开一个），随后本确定性合成又开一个
            // → 用户看到两个内容相同的大屏页面。这里检查本轮 toolCalls，LLM 已成功展示即跳过。
            {
                const calls0 = Array.isArray(finalTurn && finalTurn.toolCalls) ? finalTurn.toolCalls : [];
                const llmDisplayed = calls0.some(c => c && String(c.name || '') === 'excel_display_html_in_browser'
                    && (c.status === 'completed' || c.status === 'success'));
                if (llmDisplayed) { L('跳过：LLM 本轮已成功调用 excel_display_html_in_browser（避免双页面）'); return null; }
            }
            const htmlRenderer = App.ChatV2 && App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.htmlRenderer;
            const engine = App.ChatV2 && App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.blueprintEngine;
            if (!htmlRenderer || typeof htmlRenderer.renderHtml !== 'function') { L('跳过：htmlRenderer 未就绪'); return null; }
            if (!engine || typeof engine.computeFromSources !== 'function') { L('跳过：blueprintEngine 未就绪'); return null; }
            const toolRegistry = deps && deps.toolRegistry;
            if (!toolRegistry || !toolRegistry.get || !toolRegistry.get('excel_display_html_in_browser')) { L('跳过：display_html 工具未注册'); return null; }

            // ── 数据源 ──
            // ① 优先用 toolCalls 里 read_range 的原始返回（可能是摘要，解析不出就用 ②）
            // ② 回退：用本轮 read_range 的调用参数重新读一次全量数据（与 html-display-pipeline 同做法）
            const calls = Array.isArray(finalTurn && finalTurn.toolCalls) ? finalTurn.toolCalls : [];
            let readCall = null;
            for (let i = calls.length - 1; i >= 0; i--) {
                if (calls[i] && calls[i].name === 'excel_read_range' && calls[i].status === 'completed') { readCall = calls[i]; break; }
            }
            L('readCall=' + (readCall ? 'found' : 'none'));
            let data = readCall ? parseReadRangeResult(readCall.result) : null;
            if (!data && readCall && readCall.args) {
                L('toolCalls 结果无法解析为结构化数据 → 按原参数重读一次');
                const reRead = await toolRegistry.invokeTracked(turn, 'excel_read_range',
                    { sheet: readCall.args.sheet, range: readCall.args.range }, { round: 98, ctx: { turn, signal } });
                data = reRead && reRead.ok ? parseReadRangeResult(reRead.result) : null;
            }
            if (!data) {
                // 连 read_range 记录都没有 → 直接用当前选区读一次
                L('无 read_range 记录 → 读取当前选区');
                const reRead = await toolRegistry.invokeTracked(turn, 'excel_read_range', {}, { round: 98, ctx: { turn, signal } });
                data = reRead && reRead.ok ? parseReadRangeResult(reRead.result) : null;
            }
            if (!data || !data.rows.length) { L('跳过：未拿到有效数据（headers/rows 为空）'); return null; }
            L('数据源就绪：sheet=' + (data.worksheet || '-') + ' range=' + (data.range || '-') + ' headers=' + data.headers.length + ' rows=' + data.rows.length);

            const rawMsg = String((turn && (turn._originalUserMessage || turn.userMessage)) || '数据看板').trim();
            // 标题不得用用户原话（"帮我生成excel大看板"）：剥命令动词，空则数据驱动回退
            const _fbTitle = htmlRenderer.dashboardFallbackTitle ? htmlRenderer.dashboardFallbackTitle({ primary: { headers: data.headers, rows: data.rows } }) : '数据看板';
            const boardTitle = htmlRenderer.sanitizeBoardTitle ? htmlRenderer.sanitizeBoardTitle(rawMsg, _fbTitle) : '数据看板';
            const spec = { title: boardTitle, theme: 'business_blue' };
            const sources = [{
                id: 's1',
                sheet: data.worksheet,
                range: data.range,
                headers: data.headers,
                rows: data.rows
            }];
            let prepared = null;
            try {
                prepared = engine.computeFromSources(sources, spec);
            } catch (e) {
                L('蓝图计算失败: ' + (e && e.message));
                return null;
            }
            L('蓝图就绪：kpis=' + ((prepared && prepared.kpis || []).length) + ' charts=' + ((prepared && prepared.charts || []).length));
            let html = '';
            try {
                html = htmlRenderer.renderHtml(prepared, spec);
            } catch (e) {
                L('HTML 渲染失败: ' + (e && e.message));
                return null;
            }
            if (!html || !html.length) { L('跳过：renderHtml 返回空'); return null; }
            L('HTML 生成成功，长度=' + html.length);
            turn._htmlDashboardDone = true;

            const title = (spec.title || '数据看板') + '大屏';
            const sessionStore = deps && deps.sessionStore;
            const dashRound = 99; // 与大屏工具 invokeTracked 的 round 对齐，保证过程列表顺序与精准定位一致
            // 过程列表补一步（统一过程展示走 thoughtList）
            try {
                if (sessionStore && typeof sessionStore.startThought === 'function') {
                    sessionStore.startThought(turn.id, dashRound, '🎨 **excel_display_html_in_browser** — 生成 HTML 大屏',
                        { kind: 'tool', toolName: 'excel_display_html_in_browser' });
                }
            } catch (e) { /* 忽略 */ }

            const res = await toolRegistry.invokeTracked(turn, 'excel_display_html_in_browser',
                { htmlContent: html, title }, { round: dashRound, ctx: { turn, signal } });

            try {
                if (sessionStore && typeof sessionStore.completeThought === 'function') {
                    const ok = !!(res && res.ok);
                    sessionStore.completeThought(turn.id,
                        ok ? '✅ **excel_display_html_in_browser** ✓ → HTML 大屏已打开'
                           : '❌ **excel_display_html_in_browser** ✗（生成失败）',
                        { round: dashRound, kind: 'tool', toolName: 'excel_display_html_in_browser', status: ok ? 'completed' : 'failed' });
                }
            } catch (e) { /* 忽略 */ }

            const filePath = (res && res.ok) ? (extractHtmlFilePath(res) || '') : '';
            ChatLog.info?.('PIPELINE', `[text-pipeline] HTML 大屏合成${res && res.ok ? '成功' : '失败'}${filePath ? ' → ' + filePath : ''}`);
            return { ok: !!(res && res.ok), filePath };
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[text-pipeline] HTML 大屏合成异常: ' + (e && e.message));
            return null;
        }
    }

    // ============================================================
    // 【看板美化】确定性执行（不依赖 LLM）：标题横幅 + KPI 卡片配色 +
    // 图表区白底细边框 + 统一列宽。先探测 KPI 实际所在行，避免与数据错位。
    // ============================================================
    const DASH_THEME = {
        banner: '#1F4E79',      // 标题横幅深蓝
        bannerText: '#FFFFFF',
        kpiLabelBg: '#DEEBF7',  // KPI 标签浅蓝底
        kpiLabelFg: '#0F3D6E',
        kpiValueFg: '#123A6B',
        panelBg: '#FFFFFF',
        border: '#BDD7EE'
    };

    /**
     * 美化「看板」工作表。
     * @returns {Promise<{ok:boolean, steps:number}|null>}
     */
    async function maybeBeautifyDashboard(turn, finalTurn, deps, signal) {
        const L = (msg) => { try { ChatLog.info?.('PIPELINE', '[dash-beautify] ' + msg); } catch (e) { /* noop */ } };
        try {
            const dashFlag = !!(turn && ((turn.context && turn.context._forceDashboard) || turn._forceDashboard));
            if (!dashFlag) { L('跳过：非看板意图'); return null; }
            if (turn._dashBeautified) { L('跳过：本轮已美化'); return null; }
            const toolRegistry = deps && deps.toolRegistry;
            if (!toolRegistry || !toolRegistry.get || !toolRegistry.get('excel_format_range')) { L('跳过：format_range 未注册'); return null; }
            const calls = Array.isArray(finalTurn && finalTurn.toolCalls) ? finalTurn.toolCalls : [];
            // 一键编排工具 excel_generate_dashboard 内部已含美化 → 不重复做
            if (turn._oneClickDashboardDone
                || calls.some(c => c && String(c.name || '').indexOf('generate_dashboard') >= 0 && c.status === 'completed')) {
                L('跳过：本轮已由 excel_generate_dashboard 一键生成（内含美化）');
                return null;
            }
            // 本轮确实建过看板表才美化
            const hasBoard = calls.some(c => c && c.name === 'excel_create_sheet' && c.status === 'completed'
                && String((c.args && (c.args.name || c.args.sheet)) || '') === '看板');
            if (!hasBoard) { L('跳过：本轮未创建看板工作表'); return null; }

            const ctx = { round: 100, ctx: { turn, signal } };
            const invoke = async (name, args) => {
                try {
                    const r = await toolRegistry.invokeTracked(turn, name, args, ctx);
                    L(`${name} → ${r && r.ok ? 'ok' : ('fail: ' + String((r && r.error) || '').slice(0, 80))}`);
                    return r;
                } catch (e) {
                    L(`${name} 异常: ${e && e.message}`);
                    return null;
                }
            };

            // ① 探测 KPI 实际起始行（模型可能写 A1 或 A3）
            let kpiRow = 1;
            const probe = await invoke('excel_read_range', { sheet: '看板', range: 'A1:C8' });
            const parsed = probe && probe.ok ? parseReadRangeResult(probe.result) : null;
            if (parsed && Array.isArray(parsed.rows)) {
                const firstNonEmpty = parsed.rows.findIndex(r => Array.isArray(r)
                    && r.some(v => v !== null && v !== undefined && String(v).trim() !== ''));
                if (firstNonEmpty >= 0) kpiRow = firstNonEmpty + 1; // 0-based → 1-based 行号
            }
            L('KPI 起始行 = ' + kpiRow);

            let steps = 0;

            // ② 标题横幅（仅当 KPI 起始行 >= 3，前面有空行可放横幅）
            if (kpiRow >= 3 && toolRegistry.get('excel_add_shape')) {
                const titleText = String((turn && (turn._originalUserMessage || turn.userMessage)) || '数据看板').trim().slice(0, 24);
                const lines = JSON.stringify([{ text: titleText, size: 18, color: DASH_THEME.bannerText, bold: true, alignH: 'center' }]);
                const r = await invoke('excel_add_shape', {
                    sheet: '看板', range: 'A1:I2', shapeType: 'roundedRect',
                    fill: DASH_THEME.banner, line: 'none', cornerRadius: 0.18, noShadow: true, lines
                });
                if (r && r.ok) steps++;
            }

            // ③ KPI 标签行
            const labelRow = kpiRow;
            const valueRow = kpiRow + 1;
            const r1 = await invoke('excel_format_range', {
                sheet: '看板', range: `A${labelRow}:C${labelRow}`,
                format: { bold: true, fontSize: 11, fontColor: DASH_THEME.kpiLabelFg, bgColor: DASH_THEME.kpiLabelBg, rowHeight: 22 },
                alignment: JSON.stringify({ horizontal: 'center', vertical: 'middle', wrapText: false })
            });
            if (r1 && r1.ok) steps++;

            // ④ KPI 数值行：大字号 + 千分位
            const r2 = await invoke('excel_format_range', {
                sheet: '看板', range: `A${valueRow}:C${valueRow}`,
                format: { bold: true, fontSize: 16, fontColor: DASH_THEME.kpiValueFg, numberFormat: '#,##0.00', rowHeight: 30 },
                alignment: JSON.stringify({ horizontal: 'center', vertical: 'middle' })
            });
            if (r2 && r2.ok) steps++;

            // ⑤ 图表/留白区：白底 + 细边框
            const panelFrom = Math.max(valueRow + 2, 4);
            const r3 = await invoke('excel_format_range', {
                sheet: '看板', range: `A${panelFrom}:I46`,
                format: { bgColor: DASH_THEME.panelBg, bgScope: 'all', borderStyle: 'thin', borderColor: DASH_THEME.border }
            });
            if (r3 && r3.ok) steps++;

            // ⑥ 统一列宽（看板列宽一致更整齐；A 列略宽放标签）
            const r4 = await invoke('excel_format_range', {
                sheet: '看板', range: 'A1:I1', format: { columnWidth: 13 }
            });
            if (r4 && r4.ok) steps++;

            turn._dashBeautified = true;
            L('美化完成，成功步骤 = ' + steps);
            return { ok: steps > 0, steps };
        } catch (e) {
            L('美化异常: ' + (e && e.message));
            return null;
        }
    }

    /** 从 display_html 工具返回里挖 .html 绝对路径 */
    function extractHtmlFilePath(result) {
        const PATH_RE = /[A-Za-z]:\\[^\s"]+\.html/i;
        function scan(o, d) {
            if (!o || d > 6) return '';
            if (typeof o === 'string') {
                const m = o.match(PATH_RE);
                if (m) return m[0];
                try { return scan(JSON.parse(o), d + 1); } catch (e) { return ''; }
            }
            if (typeof o === 'object') {
                if (typeof o.filePath === 'string' && PATH_RE.test(o.filePath)) return o.filePath;
                if (typeof o.path === 'string' && PATH_RE.test(o.path)) return o.path;
                for (const k of ['result', 'data', 'content']) {
                    if (o[k] !== undefined && o[k] !== null) { const f = scan(o[k], d + 1); if (f) return f; }
                }
            }
            return '';
        }
        return scan(result, 0);
    }

    /**
     * 从工具调用结果里递归扫描 Windows 风格产物文件路径（.html 后缀）。
     * 场景：python-pipeline / code-pipeline 执行完后可能生成 .html 报告（pyecharts 图表/plotly dashboard），
     *   本函数从 tool result 中挖 `C:\...\report.html` 路径给 UI 做"产物文件 banner"。
     * 递归扫描路径：result → result.data → result.content（string 或 Array）→ JSON.parse 兜底。
     * 匹配正则：`[A-Za-z]:\\[^\s"<>']+\.html`，支持大小写盘符和正斜杠。
     * @param {Array} results - tool call 结果数组（每元素 {ok, result}，取自 turn.toolCalls）
     * @returns {string} 第一个匹配到的文件路径；找不到返回空串
     */
    function scanFilePathInResults(results) {
        const RE = /[A-Za-z]:\\[^\s"<>']+\.html/i;
        function scan(o, d) {
            if (!o || d > 6) return '';
            if (typeof o === 'string') {
                const m = o.match(RE);
                if (m) return m[0];
                try { return scan(JSON.parse(o), d + 1); } catch (e) { return ''; }
            }
            if (typeof o === 'object') {
                if (typeof o.filePath === 'string' && RE.test(o.filePath)) return o.filePath;
                if (typeof o.path === 'string' && RE.test(o.path)) return o.path;
                for (const k of ['result', 'data', 'content']) {
                    if (o[k] !== undefined && o[k] !== null) { const f = scan(o[k], d + 1); if (f) return f; }
                }
                if (Array.isArray(o.content)) {
                    for (const c of o.content) { const f = scan(c, d + 1); if (f) return f; }
                }
            }
            return '';
        }
        for (const rr of results) {
            if (rr && typeof rr === 'object') { const f = scan(rr, 0); if (f) return f; }
        }
        return '';
    }

    /**
     * agent-loop 通用对话兜底（async generator）—— execute() 的默认分发路径。
     * 场景：不需要多步骤规划（单操作/问答），也不命中 builtinChart 内置路径，
     * 让 LLM 通过 agentLoop.run 自由对话 + 按需调用 MCP 工具。
     *
     * 最终内容合成（preferSummary 判定）：
     *   - preferSummary=true  → 分析意图（"总结/统计/帮我看看"类）→ 优先 LLM 文字总结
     *   - preferSummary=false → 执行意图（"做图表/写表"类）→ 优先工具返回值（buildRawToolResultText）
     *   两者任一有值就 completeTurn，保证气泡里一定有可读结果。
     *
     * @param {Object} turn
     * @param {AbortSignal} signal
     * @param {Object} deps - 含 agentLoop/sessionStore/promptBuilder
     * @param {Array} messages - LLM 历史消息（prepareTools 透传产物）
     * @yields {Object} 事件流 —— {type:'pipeline_done'}
     */
    async function* runAgentLoopFallback(turn, signal, deps, messages) {
        const { agentLoop, sessionStore, promptBuilder } = deps;
        // 【兜底】多步骤管道降级进来时调用方可能未透传 messages，从 turn.context 取回
        // （prepareTools 已写入 _llmMessages）。否则 agent-loop 直接 INVALID_MESSAGES → 空白回复。
        const agentMessages = (Array.isArray(messages) && messages.length > 0)
            ? messages
            : (turn.context?._llmMessages || messages);
        // 非 Excel 宿主（PPT/Word）：若只给一轮工具，模型"先读选区、再调格式化工具"的两步链会断在
        // 第二步（下一轮不再下发工具）→ 只能在正文里干说（实测 Word：调 word_get_selection_text 后
        // 写下"我们将使用 word_set_table_format 工具"却什么都没执行）。故非 Excel 开启 persistTools
        // 多轮保留工具；Excel 维持既有单轮语义，避免行为回归。
        const _tpHost = (typeof window !== 'undefined' && window.SmartPilotHost) || null;
        const _tpPersist = !!(_tpHost && !_tpHost.isExcel);

        // 兜底路径：改主题 → 一键看板
        const themeFb = await maybeRunDashboardThemeChange(turn, deps, signal);
        if (themeFb && themeFb.ok) {
            sessionStore.completeTurn(turn.id, {
                content: '✅ ' + ((themeFb.result && themeFb.result.message) || '已切换看板主题。')
            });
            yield { type: 'pipeline_done', pipeline: 'text' };
            return;
        }
        const oneClickFb = await maybeRunOneClickDashboard(turn, deps, signal);
        if (oneClickFb && oneClickFb.ok) {
            const finalTurn0 = sessionStore.getCurrentTurn() || turn;
            const deliverable = (turn.context && turn.context._dashboardDeliverable) || resolveDashboardDeliverable(turn);
            let htmlDash0 = null;
            if (deliverable !== 'excel') {
                htmlDash0 = await maybeRenderHtmlDashboard(turn, finalTurn0, deps, signal);
            }
            let finalContent0 = '✅ ' + ((oneClickFb.result && oneClickFb.result.message)
                || `已生成 Excel 看板工作表「${(oneClickFb.result && oneClickFb.result.board) || '看板'}」。`);
            if (htmlDash0 && htmlDash0.ok) {
                finalContent0 += '\n\n🖥️ **HTML 大屏已生成并在浏览器打开**' +
                    (htmlDash0.filePath ? `（${htmlDash0.filePath}）` : '') + '。';
            } else if (deliverable === 'excel') {
                finalContent0 += '\n\n（按你的要求仅产出 Excel 内看板，未打开网页大屏。）';
            }
            sessionStore.completeTurn(turn.id, { content: finalContent0 });
            yield { type: 'pipeline_done', pipeline: 'text' };
            return;
        }

        yield* agentLoop.run({ turn, messages: agentMessages, tools: turn.context?._llmTools || [], signal, persistTools: _tpPersist });

        const finalTurn = sessionStore.getCurrentTurn();
        if (finalTurn && finalTurn.id === turn.id) {
            // 【无 Skill 看板】兜底路径同样尝试合成 HTML 大屏（与主管道同逻辑）
            const dashDeps = Object.assign({}, deps, {
                toolRegistry: (deps && deps.toolRegistry) || (App.ChatV2 && App.ChatV2.toolRegistry)
            });
            const htmlDash = await maybeRenderHtmlDashboard(turn, finalTurn, dashDeps, signal);
            const rawUserMsg = (turn && (turn._originalUserMessage || turn.userMessage)) || '';
            const preferSummary = typeof promptBuilder?.detectAnalysisIntent === 'function'
                ? promptBuilder.detectAnalysisIntent(rawUserMsg) : false;
            const llmText = finalTurn?.content || '';
            const rawToolResult = buildRawToolResultText(finalTurn);
            // 永远优先 LLM 总结文本（比原始工具结果 JSON 可读 100 倍）；仅当 LLM 总结为空时才 fallback rawToolResult
            let finalContent = textWhenModelSilent(finalTurn, llmText, rawToolResult);
            // 追加产物文件 banner + inline 图表 HTML（与主管道共用同一装饰函数）
            finalContent = _decorateFinalContent(finalContent, turn, finalTurn);
            // 【空回复兜底】同主管道：空内容绝不能原样透传成空气泡。
            if (!String(finalContent || '').trim()) {
                finalContent = '⚠️ 模型本次未返回任何内容（常见原因：模型服务刚启动还在加载模型、请求超时或返回了空白输出）。\n\n'
                    + '请点击重试；若连续失败，请确认模型服务（如 Ollama）已启动并已加载对应模型。';
                (typeof window !== 'undefined' && window.ChatLog || console).warn?.('PIPELINE', `[text-pipeline] 模型返回空内容 → 已用兜底提示替换（turn=${turn.id}）`);
            }
            if (htmlDash && htmlDash.ok) {
                finalContent += '\n\n🖥️ **HTML 大屏已生成并在浏览器打开**' +
                    (htmlDash.filePath ? `（${htmlDash.filePath}）` : '') + '。';
            }
            sessionStore.completeTurn(turn.id, { content: finalContent, reasoning: finalTurn?.reasoning || '' });
        }
        yield { type: 'pipeline_done', pipeline: 'text' };
    }

    // ============================================================
    // 主入口：execute —— 纯调度器（v20260924 重构后）
    //   阶段 1: 初始化（init + 生命周期 + 后续消息继承）
    //   阶段 2: 工具准备（瘦身 + 过滤 + skill 候选注入 + messages 构建）
    //   阶段 3: 分发决策（多步骤管道 / agent-loop 兜底）
    // ============================================================
    /**
     * 执行文本对话
     * @param {Object} turn - turn 对象
     * @param {AbortSignal} signal
     * @returns {AsyncGenerator} typed events
     */
    async function* execute(turn, signal) {
        const deps = getDeps();
        const { toolRegistry, sessionStore } = deps;

        try {
            // === 阶段 1: 初始化 ===
            await toolRegistry.init();

            // 【看板·改主题短路】「修改为红色主题」等在分发前硬跑，避免落入 agent-loop 只规划不调工具
            {
                const themeEarly = await maybeRunDashboardThemeChange(turn, deps, signal);
                if (themeEarly && themeEarly.ok) {
                    const msgT = (themeEarly.result && themeEarly.result.message) || '已切换看板主题。';
                    sessionStore.completeTurn(turn.id, { content: '✅ ' + msgT });
                    yield { type: 'pipeline_done', pipeline: 'text' };
                    return;
                }
            }

            // 生命周期：新任务清空上一轮执行上下文（子管道不破坏父上下文）
            const isNewTask = !turn.context?._execFollowUp && !turn.context?._isSubPipeline;
            if (isNewTask && sessionStore.clearExecSpec) {
                sessionStore.clearExecSpec();
            }
            // 编辑重发已经把被撤回的轮次从模型历史里拿掉，这里不再把那一轮的工具返回注回来。
            // 后续消息继承：注入上轮图表结果 + 继承数据快照
            applyChartFollowUp(turn);

            // === 阶段 2: 工具准备 ===
            const prepared = prepareTools(turn, deps);
            let tools = await injectStrongSkillCandidates(turn, prepared.tools, deps, prepared.matchedSkill);

            // === 阶段 2.5: RAG 知识库增强 ===
            // 用户在"RAG 设置"里配置了默认知识库且开启对话增强时，先检索知识库，
            // 命中则把上下文作为系统消息插到系统提示词之后（模型可见、用户无感）。
            // 未启用/未命中/异常一律跳过，绝不影响正常聊天（enhanceForChat 内部有 5s 超时）。
            try {
                const ragContext = await (window.App && App.rag && typeof App.rag.enhanceForChat === 'function'
                    ? App.rag.enhanceForChat(prepared.userMsg)
                    : null);
                if (ragContext && Array.isArray(prepared.messages)) {
                    const sysIdx = prepared.messages.findIndex(m => m.role === 'system');
                    prepared.messages.splice(sysIdx >= 0 ? sysIdx + 1 : 0, 0, { role: 'system', content: ragContext });
                    ChatLog.info?.('PIPELINE', `[text-pipeline] RAG 知识库上下文已注入 (${ragContext.length} 字符)`);
                }
            } catch (ragErr) {
                ChatLog.warn?.('PIPELINE', '[text-pipeline] RAG 增强失败（不影响主流程）:', ragErr?.message);
            }

            // 透传到 turn.context，供后续 runXxx 函数使用
            if (!turn.context) turn.context = {};
            turn.context._llmMessages = prepared.messages;
            turn.context._llmTools = tools;

            // === 阶段 3: 分发决策 ===
            // 路径 B: 非后续消息 + (无已匹配 Skill 或被强制多步骤) → 多步骤管道判定
            // 放行条件说明：多图表请求即使命中了 SmartChartWizard，orchestrator 也会强制
            // kind='text' 并打 _forceMultiStep（每张图各需"聚合→建图"，单 Skill 只会做第一张），
            // 此时不能因为 _matchedSkill 存在就跳过多步骤管道。
            const isUserFollowUp = turn.context?._execFollowUp && !turn.context?._isSubPipeline;
            const forceMultiStep = !!turn.context?._forceMultiStep;
            const forceChart = !!turn.context?._forceChart;
            if ((!turn.context?._matchedSkill || forceMultiStep || forceChart) && !isUserFollowUp) {
                const needsPlanning = await decideMultiStep(turn, prepared.userMsg, deps);
                if (needsPlanning) {
                    try {
                        if (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.markMultiStep === 'function') {
                            turn.goal = App.ChatV2.turnGoal.markMultiStep(turn.goal, prepared.userMsg || '');
                            if (sessionStore.updateTurn) sessionStore.updateTurn(turn.id, { goal: turn.goal });
                        }
                    } catch (_msGoal) { /* 多步目标标记失败不影响管道 */ }
                    yield* runMultiStepPipeline(turn, signal, deps);
                    return;
                }
            }

            // 路径 C: agent-loop 通用对话兜底
            yield* runAgentLoopFallback(turn, signal, deps, prepared.messages);

        } catch (err) {
            ChatLog.error?.('PIPELINE', '[text-pipeline] error', err);
            sessionStore.failTurn(turn.id, { message: err?.message || String(err) });
            yield {
                type: 'error',
                error: { code: 'PIPELINE_ERROR', message: err?.message || String(err), recoverable: false }
            };
        }
    }

    // 优先走 pipelineRegistry 声明式注册；注册表未加载时回退旧挂载
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('text', {
            execute,
            name: '文本对话',
            description: '处理纯文本用户消息（含 MCP 工具调用 / Skill 路由）',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.text = { execute };
    }
    ChatLog.info?.('CHAT-V2', '✅ text-pipeline loaded (v20260924 refactor)');
})();
