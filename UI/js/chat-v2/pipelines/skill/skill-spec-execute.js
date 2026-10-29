/**
 * ============================================================
 * skill-spec-execute.js — Skill 管道 spec 模式执行器（从 skill-pipeline.js 拆出）
 * ------------------------------------------------------------
 * 职责：executeSkillSpecMode——LLM 规划 spec → skillSpecEngine 确定性渲染 VBA →
 *       excel_execute_vba 执行 → 步骤面板渐进填色 → 聚合结果快照继承（aggregateSnapshot）。
 * 挂载：App.ChatV2.skillPipelineModules.specExecute
 * 行为零变更重构：函数体自 skill-pipeline.js 逐字迁移，未改动任何逻辑/正则/文案。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.skillPipelineModules) App.ChatV2.skillPipelineModules = {};
    const MODS = App.ChatV2.skillPipelineModules;
    const SP = MODS;
    const C = MODS.constants;
    const ChatLog = window.ChatLog || console;
    /**
     * 
     * 流程：LLM 输出结构化 spec（主题/表头/对齐/条件格式等）→ skillSpecEngine
     * 确定性渲染 VBA → 调用 excel_execute_vba 执行。
     * 关键：LLM 只负责"规划想要什么效果"（JSON），代码由人工验证的模板生成 →
     * 既灵活（样式参数化）又不报错（确定性渲染）。
     * @param {Object} turn
     * @param {Object} skill
     * @param {string} displayName
     * @param {AbortSignal} signal
     * @returns {Promise<{outputText: string, spec: Object, code: string}>}
     */
    async function* executeSkillSpecMode(turn, skill, displayName, signal) {
        const sessionStore = App.ChatV2.sessionStore;
        const toolRegistry = App.ChatV2.toolRegistry;
        const llmClient = App.ChatV2.llmClient;
        const specEngine = App.ChatV2.services?.skillSpecEngine;
        const llmIo = App.ChatV2.llmIo;

        // ── 方向 A 辅助：让 LLM 根据 anomalies 修正看板 spec ──
        async function _reviseDashboardSpec(originalSpec, anomalies, meta) {
            if (!llmIo || typeof llmIo.askJson !== 'function') return null;
            const CHART_TYPES = 'xlColumnClustered,xlColumnStacked,xlBarClustered,xlBarStacked,xlLine,xlLineMarkers,xlArea,xlPie,xlDoughnut,xlXYScatter,xlRadar,xlFunnel,xlWaterfall';
            const msgs = [
                { role: 'system', content:
                    `你是看板蓝图修正器。根据数据摘要和当前方案的问题，微调字段；KPI 少于 2 个或图表少于 1 张时可以补充数量，其余情况不要改动数量。
返回严格 JSON，字段：{ charts: [...], kpis: [...], title?, theme? }。
图表可选类型：${CHART_TYPES}。常用：分组太少→柱状图(xlColumnClustered)；占比→饼图(xlPie)/环形(xlDoughnut)；趋势→折线(xlLine)/面积(xlArea)；长标签→条形图(xlBarClustered)。
聚合函数 func 可选：sum, avg, max, min, count。` },
                { role: 'user', content:
                    `数据摘要：${JSON.stringify(meta)}
当前方案：${JSON.stringify({ title: originalSpec.title, theme: originalSpec.theme, charts: originalSpec.charts, kpis: originalSpec.kpis })}
需要修正的问题：${JSON.stringify(anomalies, null, 2)}

请返回修正后的 charts 和 kpis（数量只能增不能减；不足下限时补齐到 KPI≥2、图表≥1）。只返回 JSON，不要其它解释。` }
            ];
            const result = await llmIo.askJson({ messages: msgs, config: App.config, signal });
            if (result.error || !result.json) {
                ChatLog.warn?.('PIPELINE', `[_reviseDashboardSpec] LLM 返回异常:`, result.error?.message || '无 JSON');
                return null;
            }
            return result.json;
        }
        // 复用 text-pipeline 的意图检测（按维聚合/数值列/聚合函数）
        const intentDetector = App.ChatV2.intentDetector || {};

        if (!specEngine || typeof specEngine.planAndRender !== 'function') {
            throw new Error('skillSpecEngine 未就绪');
        }
        // 先同步 MCP 工具（excel_execute_vba 是 MCP 工具，异步注册），
        // 避免首次触发时工具未注册。注意：即使工具未注册，执行阶段会优先走
        // excelBridge.ExecuteVBA 直接接口，此处检查仅用于日志提示，不阻断。
        if (typeof toolRegistry?.init === 'function') {
            await toolRegistry.init();
        }
        if (!toolRegistry.get('excel_execute_vba')) {
            ChatLog.warn?.('PIPELINE', '[skill-spec] excel_execute_vba 工具未注册（将回退 excelBridge.ExecuteVBA 直接执行）');
        }

        // 按 Skill 类型选择规划 prompt / 规范化器 / 执行策略：
        //   - SmartTableBeautifier / SmartTableFormatter → VBA 渲染（美化+格式整理）
        //   - SmartChartWizard → 工具链模式（excel_create_chart）
        //   - SmartDataCleaner → 工具链模式（清洗工具）
        const skillName = String(skill.name || '');
        // 类型判定收敛在 config/skill-domain.js（detectSkillType 经 detectTypeByName 返回唯一类型）
        const skillType = C.detectSkillType(skillName);
        const isBeautifier = skillType === 'beautifier';
        const isFormatter = skillType === 'formatter';
        const isChart = skillType === 'chart';
        const isCleaner = skillType === 'cleaner';
        const isCrud = skillType === 'crud';
        const isDashboard = skillType === 'dashboard';
        // 规划 prompt / 规范化器 / 方案名 / 骨架 kind 由 constants.SKILL_SPEC_CONTRACT 按类型解析
        // （promptBuilder/normalizer 是 skillSpecEngine 方法名，延迟解析保持常量模块零依赖）
        const specContract = C.SKILL_SPEC_CONTRACT[skillType] || C.SKILL_SPEC_CONTRACT.beautifier;

        let planningPrompt, normalizeSpecFn, planLabel = '';
        // 聚合结果表快照（地址+参数），随 result 写入 _execSpec 供下一轮复用。
        // 声明在函数顶层：美化/清洗/图表所有路径的 result 都会引用它，避免图表路径块内声明导致的 ReferenceError。
        let aggregateSnapshot = null;
        planningPrompt = specEngine[specContract.promptBuilder] ? specEngine[specContract.promptBuilder](skill) : specEngine.buildPlanningPrompt(skill);
        normalizeSpecFn = specEngine[specContract.normalizer] || specEngine.normalizeSpec;
        planLabel = specContract.planLabel;

        yield { type: 'status', text: `正在规划 ${displayName} ${planLabel}...` };

        // 1. LLM 规划 spec（结构化 JSON，不写代码）
        const userMessage = turn.userMessage || '';
        let specRaw = null;
        let plannedByLLM = false;

        if (llmClient && typeof llmClient.stream === 'function') {
            const msgs = [
                { role: 'system', content: planningPrompt },
                { role: 'user', content: userMessage }
            ];
            let llmOut = '';
            try {
                for await (const evt of llmClient.stream({ messages: msgs, config: App.config, signal })) {
                    if (signal?.aborted) break;
                    if (evt.type === 'token') {
                        llmOut += evt.delta || '';
                        yield evt;
                    }
                }
                specRaw = null;
                // 从 LLM 输出提取 spec JSON
                const m = llmOut.match(/```(?:json)?\s*([\s\S]*?)```/);
                const text = m ? m[1].trim() : llmOut.trim();
                try {
                    const parsed = JSON.parse(text);
                    if (parsed && typeof parsed === 'object') specRaw = parsed;
                } catch (e) { /* 非 JSON，用消息兜底 */ }
            } catch (llmErr) {
                ChatLog.warn?.('PIPELINE', '[skill-spec] LLM 规划失败，用消息关键词兜底:', llmErr?.message);
            }
        }

        // 2. 无 LLM 输出或解析失败 → 从用户消息关键词确定性兜底
        let plan;
        // 安全守卫取消标志（2026-10-21）：用户拒绝危险操作后必须优雅收尾，
        // 绝不 throw——否则上层 skill-pipeline 会 fallbackToText 降级到 agent-loop
        // 二次执行，出现"取消了清空却又跑去读数据"的诡异行为（skill-crud-e2e 实测）。
        let specCancelled = false;
        let cancelMessage = '';
        if (specRaw) {
            plan = normalizeSpecFn(specRaw);
            plannedByLLM = true;
        } else {
            plan = isCrud
                ? specEngine.crudSpecFromMessage(userMessage)
                : isDashboard
                    ? specEngine.dashboardSpecFromMessage(userMessage)
                    : isFormatter
                        ? specEngine.normalizeFormatSpec(specEngine.specFromMessage(userMessage))
                        : isCleaner
                            ? specEngine.cleanSpecFromMessage(userMessage)
                            : specEngine.specFromMessage(userMessage);
        }
        // CRUD 兜底误判为 null（无法判定操作族）→ 明确报错提示，避免静默失败
        if (isCrud && !plan) {
            throw new Error('未识别出明确的增删改查操作，请说明要读取/写入/插入/删除/清空的目标与内容');
        }

        // ── Dashboard deliverable：用户消息强意图覆盖 LLM 默认 dual ──
        // LLM 常无脑写 dual → "做个 Excel 看板"仍弹浏览器。关键词兜底只在用户明确
        // 只要 excel/html 时覆盖；模糊说法（仅"看板"）仍尊重 LLM / 默认 dual。
        if (isDashboard && plan && typeof specEngine.dashboardSpecFromMessage === 'function') {
            const fromMsg = specEngine.dashboardSpecFromMessage(userMessage);
            const msgDeliv = String(fromMsg && fromMsg.deliverable || '').toLowerCase();
            const planDeliv = String(plan.deliverable || 'dual').toLowerCase();
            if ((msgDeliv === 'excel' || msgDeliv === 'html') && planDeliv !== msgDeliv) {
                ChatLog.info?.('PIPELINE', `[skill-dashboard] deliverable 用户意图覆盖: ${planDeliv} → ${msgDeliv}`);
                plan.deliverable = msgDeliv;
            } else {
                ChatLog.info?.('PIPELINE', `[skill-dashboard] deliverable: ${planDeliv}（LLM/默认）`);
            }
        }

        // 清洗且 operations 为空（小模型输出空/非法操作名被白名单过滤）→ 关键词兜底补全，
        // 否则后端 excel_clean_data 会报"请指定清洗操作 operations"。
        if (isCleaner && (!Array.isArray(plan?.operations) || plan.operations.length === 0)) {
            const fb = specEngine.cleanSpecFromMessage(userMessage);
            if (fb.operations.length) {
                plan.operations = fb.operations;
                ChatLog.info?.('PIPELINE', `[skill-cleaner] LLM 未给出清洗操作，关键词兜底: ${fb.operations.join(', ')}`);
            }
        }

        // ✨ 模糊请求扩展：用户只说"清洗数据"没说具体做什么 且 【完全没识别出任何具体操作】(operations 为空)
// → 自动扩展成全量清洗 5 步，让面板有东西展示，用户也能看到"我真的做了清洗"。
// 识别依据：userMessage 含"清洗/清理/优化/整理"但完全不含具体操作关键词。
// ⚠️ 修复：条件是 operations.length === 0（而非 ≤1）。
//   此前用 ≤1，导致"去除文本前后空格"这类已识别到单操作(trimSpaces)的请求，
//   因 HAS_SPECIFIC_OP 正则缺"空格"而误判为模糊请求，被强行扩成全量清洗（还误删了重复/空行）。
//   只要识别出明确单操作，就应尊重用户只做那一步。
if (isCleaner && Array.isArray(plan?.operations) && plan.operations.length === 0) {
            const msg = String(userMessage || '').toLowerCase();
            const HAS_CLEAN_INTENT = C.CLEAN_INTENT_RE.test(msg);
            const HAS_SPECIFIC_OP = C.CLEAN_SPECIFIC_OP_RE.test(msg);
            if (HAS_CLEAN_INTENT && !HAS_SPECIFIC_OP) {
                // 模糊清洗意图 → 全量 5 步。
                // 全量操作与顺序收敛在 constants.FULL_CLEAN_OPS（与 CLEAN_OPS 注册表相对顺序一致，
                // 骨架/排序/执行共用同一数据源，杜绝 stepIndex 错位）
                plan.operations = [...C.FULL_CLEAN_OPS];
                ChatLog.info?.('PIPELINE', `[skill-cleaner] 模糊请求"${userMessage.slice(0,20)}" → 自动扩展全量清洗 ${plan.operations.length} 步: ${plan.operations.join(', ')}`);
            }
        }

        // ⚠️ 纯清洗意图二次覆盖：用户消息只表达"清洗/清理/整理数据"这类整体意图、且完全没点具体子操作时，
        // 即使 LLM 已输出具体 operations 也强制全量清洗。否则弱模型（如 qwen:7b / qwen2.5）常把"清理数据"
        // 错判成只剩一件子操作（如 removeEmptyRows），导致"清理数据变成只去空行"。
        if (isCleaner) {
            const __msgC = String(userMessage || '').toLowerCase();
            const PURE_CLEAN = C.CLEAN_INTENT_RE.test(__msgC);
            const HAS_SUB_OP = C.CLEAN_SPECIFIC_OP_EXT_RE.test(__msgC);
            if (PURE_CLEAN && !HAS_SUB_OP && Array.isArray(plan?.operations)) {
                const FULL_OPS = C.FULL_CLEAN_OPS;
                const curOps = plan.operations;
                const diff = curOps.length !== FULL_OPS.length || FULL_OPS.some((o, i) => o !== curOps[i]);
                if (diff) {
                    plan.operations = [...FULL_OPS];
                    ChatLog.info?.('PIPELINE', `[skill-cleaner] 纯清洗意图覆盖 → FULL_OPS ${FULL_OPS.join(', ')} (原 ops=${JSON.stringify(curOps)})`);
                }
            }
        }

        // ⚠️ 空行清理确定性锁定：消息明确只要"删空行/清理空行"、且没提其它清洗动作时，
        // 即使 LLM 给出了别的 operations 也强制为 removeEmptyRows 单步，
        // 杜绝"清理空行被规划成去重/清空/其它"。
        if (isCleaner) {
            const __msgR = String(userMessage || '').toLowerCase();
            // ⚠️ 词级判定：只要出现"空行"二字就是删空行意图（不要求"清理空行/删除空行"连成串，
            // 否则"清理一下空行""把空行清理掉"会被漏掉、误落到填充空值/去重/清空）。"空值/空白"属填充、
            // "空格"属去空格，均不在此列。
            const ONLY_EMPTY_ROW = C.CLEAN_EMPTY_ROW_RE.test(__msgR)
                && !C.CLEAN_EMPTY_ROW_CONFLICT_RE.test(__msgR);
            if (ONLY_EMPTY_ROW && Array.isArray(plan?.operations)
                && !(plan.operations.length === 1 && plan.operations[0] === 'removeEmptyRows')) {
                plan.operations = ['removeEmptyRows'];
                ChatLog.info?.('PIPELINE', `[skill-cleaner] 空行清理锁定 → removeEmptyRows (消息含"空行"，原 ops=${JSON.stringify(plan.operations)})`);
            }
        }
        // 覆盖小模型在 spec 中可能丢失的 markDuplicate。
        // ✨ 无条件跑 detectCleanDedupeMode——不管 plan.operations 里有没有 removeDuplicates
        // （LLM spec 偶尔漏 operations 字段，但 specEngine 内部还是会加 removeDuplicates）。
        if (isCleaner) {
            const dedupeMode = SP.util.detectCleanDedupeMode(userMessage);
            ChatLog.info?.('PIPELINE', `[skill-cleaner] detectCleanDedupeMode("${userMessage.slice(0,30)}")=${dedupeMode} (type=${typeof dedupeMode})  plan.markDuplicate before=${plan.markDuplicate}`);
            if (dedupeMode !== null) {
                plan.markDuplicate = dedupeMode;
                if (Array.isArray(plan.operations) && !plan.operations.includes('removeDuplicates')) {
                    plan.operations.push('removeDuplicates');
                    ChatLog.info?.('PIPELINE', `[skill-cleaner] 兜底补 operations.removeDuplicates（markDuplicate=${dedupeMode}）`);
                }
            } else {
                // ✨ 默认去重模式：仅高亮标记（安全），不直接删除。
                // 清洗措辞含糊（如"清洗数据"没说"去重/高亮/删除"）时，
                // 先只高亮重复行让用户确认，而非直接删行，避免误删数据。
                const hasDedupeOp = Array.isArray(plan?.operations) && plan.operations.includes('removeDuplicates');
                if (hasDedupeOp && typeof plan.markDuplicate !== 'boolean') {
                    plan.markDuplicate = true;
                    ChatLog.info?.('PIPELINE', `[skill-cleaner] 默认 markDuplicate=true（plan 有 removeDuplicates 但用户未指定模式，按安全意图默认仅高亮）`);
                }
            }
        }

        // ✨ 强制排序 plan.operations 对齐骨架顺序（buildSpecSteps 的 if 出现顺序），
        // 确保骨架 stepIndex 0,1,2... 和循环执行 stepIndex 完全一致。
        if (isCleaner && Array.isArray(plan?.operations) && plan.operations.length > 1) {
            // 骨架顺序收敛在 constants.CLEAN_OP_NAMES（buildSpecSteps 按同一注册表渲染）
            const SKELETON_ORDER = C.CLEAN_OP_NAMES;
            plan.operations.sort((a, b) => {
                const ai = SKELETON_ORDER.indexOf(a);
                const bi = SKELETON_ORDER.indexOf(b);
                return (ai === -1 ? 99 : ai) - (bi === -1 ? 99 : bi);
            });
            ChatLog.info?.('PIPELINE', `[skill-cleaner] operations 排序后: ${plan.operations.join(', ')}`);
        }

        // ✨ 关键改动：提前生成 spec 骨架并 yield skeleton 事件。
        // 原 buildSpecSteps 在 VBA 执行完后才调用（L774），导致面板是"一次性全完成"。
        // 现在 LLM 规划 + normalizeSpec 完成后立即 buildSpecSteps → 骨架事件，
        // 外层 execute() 收到后用 ProgressTracker 初始化完整骨架面板（N 步 pending），
        // 后续 generator 中间事件逐步填色。
        // 注意：_kind 必须和 L773 的最终值一致，否则骨架标题不匹配。
        const _kindForSkeleton = specContract.skeletonKind;

        // ════════════════════════════════════════════════════════════════
        // 看板方向 A + C 预修正：骨架生成前先读数据 + computeFromSources，
        // 让方向 C 自动修已知问题，方向 A 用 LLM 修模糊问题。
        // 这样骨架基于最终 spec，不会出现"先闪一次骨架再修"的体验。
        // ════════════════════════════════════════════════════════════════
        let _dashPrecomputed = null;
        if (isDashboard) {
            try {
                const _dashEngine = App.ChatV2.dashboardModules?.blueprintEngine;
                if (_dashEngine && typeof _dashEngine.computeFromSources === 'function') {
                    const _preCtx = {
                        selection: turn.context?.selection || null,
                        activeSheet: turn.context?.activeSheet?.sheet || ''
                    };
                    const _preReadSteps = _dashEngine.buildReadSteps(plan, _preCtx);
                    const _preSources = [];
                    for (const _st of _preReadSteps) {
                        if (signal?.aborted) throw new Error('user_aborted');
                        let _r = null;
                        try {
                            _r = await toolRegistry.invokeTracked(turn, _st.tool, _st.args, {
                                round: 1, displayName: _st.displayName, displayArgs: _st.displayArgs,
                                ctx: { turn, sessionStore, signal }
                            });
                        } catch (e) {
                            if (e?.message === "user_aborted") throw e;
                    _stepOk = false;
                    _stepErr = String(e?.message || "执行失败").slice(0, 200);
                            ChatLog.warn?.('PIPELINE', `[dashboard-pre] 读取数据源失败（跳过）:`, e?.message);
                        }
                        let _meta = _r && (_r.result != null ? _r.result : _r);
                        if (typeof _meta === 'string') { try { _meta = JSON.parse(_meta); } catch { _meta = null; } }
                        const _headers = _meta && Array.isArray(_meta.headers) ? _meta.headers.map(String) : [];
                        const _rows = _meta && (Array.isArray(_meta.result) ? _meta.result : Array.isArray(_meta.rows) ? _meta.rows : []);
                        ChatLog.info?.('PIPELINE', `[dashboard-pre] excel_read_range 返回: keys=${_meta ? Object.keys(_meta).join(',') : 'null'} headers(${_headers.length})=[${_headers.slice(0,6).join(',')}] rows(${_rows.length}) row0_type=${_rows[0] ? Array.isArray(_rows[0]) ? 'array('+_rows[0].length+')' : typeof _rows[0] : 'N/A'} row0_sample=${_rows[0] ? (Array.isArray(_rows[0]) ? JSON.stringify(_rows[0].slice(0,6)) : JSON.stringify(_rows[0]).slice(0,200)) : 'N/A'}`);
                        if (_headers.length && _rows.length) {
                            _preSources.push({
                                id: (plan.dataSources && plan.dataSources[_preSources.length] && plan.dataSources[_preSources.length].id) || ('src_' + _preSources.length),
                                headers: _headers, rows: _rows,
                                range: _meta.range || (_st.args && _st.args.range) || '',
                                worksheet: _meta.worksheet || (_st.args && _st.args.sheet) || ''
                            });
                        }
                    }
                    if (_preSources.length > 0) {
                        let _prep = _dashEngine.computeFromSources(_preSources, plan);
                        const _anoms = _prep._anomalies || [];
                        const _severeAnoms = _anoms.filter(a => a.type === 'chart_all_zero' || a.type === 'chart_too_few_groups' || a.type === 'chart_regenerated');

                        // 方向 A：有严重问题且 LLM 可用 → 让 AI 调整方案
                        if (_severeAnoms.length > 0 && llmIo && typeof llmIo.askJson === 'function') {
                            yield { type: 'status', text: `检测到看板方案有 ${_severeAnoms.length} 处不理想，AI 正在自动调整...` };
                            try {
                                const _rev = await _reviseDashboardSpec(plan, _severeAnoms, _prep._meta || {});
                                if (_rev && typeof _rev === 'object') {
                                    // 只允许改 charts/kpis 内部字段（类型/维度/值列/函数/标题），数量只能增（低于下限时补齐）
                                    if (Array.isArray(_rev.charts)) {
                                        // 修正输出是 LLM 裸 JSON，先过白名单归一（label/占位符回退），再合并
                                        const _cleanedCharts = _rev.charts.slice(0, 10)
                                            .map((c, i) => (typeof specEngine.normalizeDashboardChart === 'function' ? specEngine.normalizeDashboardChart(c, i) : c))
                                            .filter(Boolean);
                                        for (let i = 0; i < _cleanedCharts.length; i++) {
                                            if (i < plan.charts.length) Object.assign(plan.charts[i], _cleanedCharts[i]);
                                            else plan.charts.push(_cleanedCharts[i]);
                                        }
                                    }
                                    if (Array.isArray(_rev.kpis)) {
                                        // 修正输出是 LLM 裸 JSON，先过白名单归一（label 空/占位 → 「指标 N」），再合并，杜绝无标题卡片
                                        const _cleanedKpis = _rev.kpis.slice(0, 12)
                                            .map((k, i) => (typeof specEngine.normalizeDashboardKpi === 'function' ? specEngine.normalizeDashboardKpi(k, i) : k))
                                            .filter(Boolean);
                                        for (let i = 0; i < _cleanedKpis.length; i++) {
                                            if (i < plan.kpis.length) Object.assign(plan.kpis[i], _cleanedKpis[i]);
                                            else plan.kpis.push(_cleanedKpis[i]);
                                        }
                                    }
                                    if (_rev.title && typeof _rev.title === 'string') plan.title = _rev.title;
                                    if (_rev.theme && typeof _rev.theme === 'string') plan.theme = _rev.theme;
                                    ChatLog.info?.('PIPELINE', `[dashboard-pre] 方向 A LLM 修正完成`);
                                }
                            } catch (_le) {
                                ChatLog.warn?.('PIPELINE', `[dashboard-pre] 方向 A LLM 修正失败（不阻断）:`, _le?.message);
                            }
                            // 修正后重新 compute（方向 C 会再次自动修正新 spec）
                            _prep = _dashEngine.computeFromSources(_preSources, plan);
                        }

                        _dashPrecomputed = { sources: _preSources, prepared: _prep };
                        ChatLog.info?.('PIPELINE', `[dashboard-pre] 预计算完成: rows=${_prep._meta?.sourceRows} kpis=${_prep.kpis.length} charts=${_prep.charts.length} anomalies=${(_prep._anomalies || []).length}`);
                    }
                }
            } catch (e) {
                if (e?.message === 'user_aborted') throw e;
                ChatLog.warn?.('PIPELINE', '[dashboard-pre] 预计算失败（不阻断，正常流程继续）:', e?.message);
                _dashPrecomputed = null;
            }
        }

        const rawSkeleton = SP.specBuild.buildSpecSteps(plan, _kindForSkeleton, isDashboard && _dashPrecomputed ? { prepared: _dashPrecomputed.prepared } : null);
        if (!rawSkeleton || rawSkeleton.length === 0) {
            rawSkeleton.push({ title: '执行操作', detail: '', status: 'pending', strategy: 'direct' });
        }

        // ✨ 包装：Skill 本身是大步骤，N 个 op 是 subSteps
        let earlySkeleton;
        if (isCleaner || isCrud) {
            // Skill 大步骤 + 小步骤
            earlySkeleton = [{
                title: displayName || 'Skill 执行',
                description: displayName || 'Skill 执行',
                strategy: 'pipeline',
                status: 'pending',
                subSteps: rawSkeleton.map(s => ({
                    title: s.title,
                    description: s.description || s.title,
                    detail: s.detail || '',
                    status: 'pending',
                    strategy: 'direct'
                }))
            }];
        } else {
            earlySkeleton = rawSkeleton.map(s => ({ ...s, status: 'pending' }));
        }

        ChatLog.info?.('PIPELINE', `[skill-spec] skeleton 生成: kind=${_kindForSkeleton} ${isCleaner || isCrud ? '嵌套' : '平铺'} 共 ${earlySkeleton.length} 步 (subSteps=${earlySkeleton[0]?.subSteps?.length || 0})`);
        yield { type: 'skeleton', steps: earlySkeleton };
        // 大步骤标记为 running
        yield { type: 'progress', stepIndex: 0, status: 'running' };

        // 3. 按 Skill 类型分发执行：
        //    - Beautifier → 渲染确定性 VBA → ExecuteVBA
        //    - Chart → 调用 excel_create_chart 工具链
        //    - Cleaner → 调用清洗工具链
        const ctx = {
            selection: turn.context?.selection || null,
            activeSheet: turn.context?.activeSheet?.sheet || '',
            // 后续消息继承透传：applySkillFollowUp 已把上一轮图表聚合快照/数据源范围写入 turn.context，
            // 这里一并带入，供"换成柱状图/折线图"等类型切换复用
            _aggregateResult: turn.context?._aggregateResult || null,
            _chartDataRange: turn.context?._chartDataRange || null
        };
        let probedInfo = null;
        // ✨ 三级区域定位：用户显式写死区域 > 已选选区 > 自动探测。
        // 优先从用户输入解析显式写死的目标区域（如"A1:E31"），防止其被选区/自动探测覆盖。
        const _userInput = (typeof turn.userMessage === 'string' ? turn.userMessage : '')
            || turn?.text || turn?.query || turn?.prompt || '';
        const explicitRange = _userInput ? SP.specBuild.parseExplicitRange(_userInput) : null;
        if (explicitRange) {
            const _tgtSheet = turn.context?.activeSheet?.sheet || ctx.selection?.sheetName || '';
            ctx.selection = Object.assign({}, ctx.selection || {}, {
                address: explicitRange.address,
                range: explicitRange.address,
                sheetName: _tgtSheet || '',
                rowCount: explicitRange.rowCount,
                colCount: explicitRange.colCount,
                explicit: true
            });
        }
        // ✨ 自动探测（前端先取地址，供写型/绘图操作构造工具调用）。
        // 美化/格式/图表没有预置"读取"步骤，须在开头探测一次 excel_read_range；
        // 清洗的"读取数据"步骤会兼任探测（在读取处只调一次 excel_read_range），故不在开头重复探测。
        // 用户已显式写死区域时无需探测。
        if ((isBeautifier || isFormatter || isChart) && !explicitRange) {
            try {
                const probed = await SP.specBuild.detectDataSelection(ctx, turn, toolRegistry, sessionStore, signal);
                if (probed) {
                    ctx.selection = probed;
                    probedInfo = probed;
                }
            } catch (e) {
                ChatLog.warn?.('PIPELINE', '[skill-spec] 选区自动探测失败，按原选区/既有回退执行:', e && e.message ? e.message : e);
            }
        }
        let execResult = null;
        let planSummary = '';

        // ── 图表特有变量声明到函数最外层
        //    旧代码把 effectiveTitle/dataRange/chartMeta 放在 if (isChart) 内部声明，
        //    导致（1）总结表格 if(isChart) 分支二次访问时出现作用域穿透
        //         （2）工具兜底 catch 块 / 第三方渲染器引用时直接 ReferenceError
        //    解决方案：在函数入口统一初始化为默认值，isChart 分支内再填充
        let effectiveTitle = '';             // 推理后的图表标题（用户说"占比"→"占比分析"）
        let dataRange = null;                // 最终传给后端的 dataRange
        let rangeForDisplay = '';            // UI 显示用的范围文案
        let chartMeta = { ...C.UI_DEFAULTS.chartMeta };
        let displayChartType = '';           // displayArgs 用
        let displayTitle = C.UI_DEFAULTS.chartTitle;       // displayArgs 用
        let displayRange = '';               // displayArgs 用
        let chartSnapMeta = null;            // 快照读取结果（headers+rows），供"纯文本选区→频次聚合"判定
        // renderedCode 提升到函数级，供总结输出引用（块内 let 声明会作用域穿透报错）
        let renderedCode = '';               // 美化/格式整理渲染出的 VBA 代码

        if (isChart) {
            // ── 图表工具链 ──
            // 枚举→中文名+图标，避免截图里"Column Clustered（柱状图功能却显示英文）"
            chartMeta = (typeof specEngine.humanizeChartType === 'function')
                ? specEngine.humanizeChartType(plan.chartType)
                : chartMeta;
            // 用户说"展示占比"→"占比分析"，避免硬编码"(自动)"
            effectiveTitle = plan.title || '';
            if (!effectiveTitle && typeof specEngine.inferDefaultChartTitle === 'function') {
                effectiveTitle = specEngine.inferDefaultChartTitle(userMessage, plan.chartType);
            }
            planSummary = `${chartMeta.icon || '📊'} ${chartMeta.name}${effectiveTitle ? '《' + effectiveTitle + '》' : ''}`;
            yield { type: 'status', text: `已按「${planSummary}」规划图表，正在创建...` };

            // 数据范围：spec 优先，缺省用当前选区
            // selection.sheetName 可能为 undefined，绝不能拼接成 "undefined!A1:B10" 发给后端
            // 【选区优先】用户明确说"对选中区域生成图表"类语义时，消息发送时的选区优先于
            // LLM 自拟的 dataRange。真实 bug：选区是纯文本分类列（如 I79:I101 客户类型），
            // LLM 却按上下文整表自拟 $A$1:$I$101，整表数值列掩盖了纯文本判定 →
            // 直接整表画饼图 → Excel 取第一列（日期文本）当系列值 → 空壳图。
            if (plan.dataRange && ctx.selection?.address &&
                C.USER_SELECTION_RE.test(userMessage)) {
                const normRange = (s) => String(s).replace(/\$/g, '').toUpperCase().split(':').sort().join(':');
                let planPart = String(plan.dataRange);
                if (planPart.includes('!')) planPart = planPart.split('!').pop();
                if (normRange(ctx.selection.address) !== normRange(planPart)) {
                    ChatLog.info?.('PIPELINE', `[skill-chart] 用户指明"选中区域"：选区 ${normRange(ctx.selection.address)} 优先于规划范围 ${normRange(planPart)}`);
                    plan.dataRange = null;   // 走下方选区分支
                }
            }
            dataRange = plan.dataRange || null;
            rangeForDisplay = '';  // displayArgs 用，显示更简洁
            if (!dataRange && ctx.selection?.address) {
                const sName = ctx.selection.sheetName;
                const hasValidSheet = C.isValidSheetName(sName);
                dataRange = hasValidSheet ? (sName + '!' + ctx.selection.address) : ctx.selection.address;
                rangeForDisplay = `当前选区 (${ctx.selection.address})`;
            } else if (dataRange) {
                rangeForDisplay = dataRange;
            }
            // dataRange 空值校验推迟到"图表类型切换后续继承"（见下）之后：
            // 纯类型切换的第二条消息（如"换成柱状图"）通常无新选区也无 plan 范围，
            // 需先复用上一轮数据源/聚合快照，若此时提前 throw 会阻断继承导致第二条消息必失败。
            // 记录原始数据范围：后续分类型"自动频次统计"会覆盖 dataRange 指向图表数据源表，
            // 聚合/回退必须以原始范围为准，避免把聚合/建图绑定到频次表导致"找不到销售额列/错误图表"。
            const originalDataRange = dataRange;

            // 快照选区到后端 RangeContext：显式调用 excel_read_range 读取快照地址，
            // 后端成功读取后会把地址写入 RangeContext（ExcelMcpTools.Core.vb），
            // 后续聚合/图表等工具即使未显式传 range 也能命中快照地址，不依赖用户实时选区。
            // 快照失败不阻断主流程（仅优化，后续工具仍走显式 range / 实时选区回退）。
            // 纯类型切换继承时 dataRange 可能尚为空（待下方继承填充），跳过快照与自动频次统计。
            // ★ 无选区（dataRange 为空）时也要快照：让后端走 UsedRange/选区兜底读取并返回真实范围，
            //   供"读取数据"步骤展示真实地址（而非"当前选定区域"占位文本）。
            //   仅当上一次图表已有聚合结果（纯类型切换继承）时才跳过，避免重复快照覆盖继承数据源。
            if ((dataRange || !turn.context?._aggregateResult) && toolRegistry.get('excel_read_range')) {
                try {
                    const snapArgs = { returnFormat: 'object', hasHeaders: true };
                    if (dataRange) {
                        const snapMatch = String(dataRange).match(/^(.*)!(.+)$/);
                        if (snapMatch) {
                            snapArgs.sheet = snapMatch[1];
                            snapArgs.range = snapMatch[2];
                        } else {
                            snapArgs.range = String(dataRange);
                        }
                    }
                    const snapResult = await toolRegistry.invokeTracked(turn, 'excel_read_range', snapArgs, {
                        displayName: '读取选区数据',
                        displayArgs: { '数据范围': dataRange },
                        ctx: { turn, sessionStore, signal }
                    });
                    if (signal?.aborted) throw new Error('user_aborted');
                    if (snapResult?.ok === false || snapResult?.error) {
                        ChatLog.warn?.('PIPELINE', '[skill-chart] 快照选区到 RangeContext 失败（不阻断）:', snapResult?.error);
                    } else {
                        let meta = snapResult?.result ?? snapResult;
                        if (typeof meta === 'string') { try { meta = JSON.parse(meta); } catch { meta = null; } }
                        // 须是 read_range 的结构化返回（headers+result），排除 invokeTracked 包装对象
                        if (meta && typeof meta === 'object' && Array.isArray(meta.headers)) chartSnapMeta = meta;
                    }
                } catch (e) {
                    if (e?.message === 'user_aborted') throw e;
                    ChatLog.warn?.('PIPELINE', '[skill-chart] 快照选区异常（不阻断）:', e?.message);
                }
            }
            // ──【智能识别】纯文本选区 → 频次聚合后再绘图 ──
            // 选区只有分类列（如"客户类型"：企业/个人/政府客户）且无任何数值列时，
            // "对选中区域生成图表/占比分析"的正确语义是按类别统计频次占比。
            // 直接 SetSourceData 纯文本列不会报错但图表是空壳（历史 bug：饼图只有标题没有扇区）。
            // 故先在内存统计频次、写入"图表数据源"表（自动新建/覆盖），再在该结果上绘图。
            try {
                const snapHeaders = Array.isArray(chartSnapMeta?.headers) ? chartSnapMeta.headers : null;
                const snapRows = Array.isArray(chartSnapMeta?.result) ? chartSnapMeta.result : null;
                if (snapHeaders && snapHeaders.length > 0 && snapRows && snapRows.length > 0) {
                    const isNumVal = (v) => v !== null && v !== undefined && v !== '' &&
                        typeof v !== 'object' && Number.isFinite(Number(v));
                    const colNumeric = snapHeaders.map(h => snapRows.some(r => r && isNumVal(r[h])));
                    // ──【智能识别】分类型图表 × 多列明细 → 自动选分类列做频次聚合 ──
                    // 用户在明细表（日期/区域/产品/金额等多列）上直接"生成图表"且 LLM 规划为
                    // 饼图/环形这类"分类占比"图时，SetSourceData 用整块区域会把日期/序号列误当
                    // 数值序列 → 空壳图（只出标题/图例，没有扇区）。正确语义是挑一个分类列统计频次
                    // 再画占比。此分支扩展触发条件：不限于"全文本列"，分类型图表 + 存在可作维度的
                    // 文本分类列（排除日期/纯序号）也自动聚合，避免硬画空壳。
                    const ctRaw = String(plan?.chartType || '').toLowerCase();
                    // 分类型图表（饼/环/瀑布/旭日）或用户消息含"占比/构成"语意 → 归类为占比意图
                    const isCategoryChart = /xlpie|xldoughnut|xlwaterfall|xlsunburst/.test(ctRaw) ||
                        /占比|比例|构成|饼图|环形/.test(userMessage);
                    // "分类×数值"标准两列结构（首列表头文本、次列数值）不需要频次聚合——SetSourceData
                    // 后端已有双列修复。仅当出现 ≥3 列或多列明细时才走自动频次。
                    const twoColCatVal =
                        snapHeaders.length >= 2 &&
                        (colNumeric[1] === true) &&            // 有数值列
                        (colNumeric.length <= 2 || snapHeaders.filter((_, i) => colNumeric[i]).length === 1);
                    // 若用户已显式表达"按XX汇总XX"的聚合意图，则频次统计是误导性派生（会额外创建
                    // "图表数据源"表），且与下方 excel_aggregate 的聚合结果表语义重叠/"污染"选区。
                    // 这种情况数据源判定交给聚合块，跳过频次统计，避免残留冗余工作表。
                    const hasAggIntent = (typeof intentDetector.detectAggregateIntent === 'function' &&
                        intentDetector.detectAggregateIntent(userMessage));
                    const needAutoFreq = !hasAggIntent && ((!colNumeric.some(Boolean)) ||
                        (isCategoryChart && !twoColCatVal && toolRegistry.get('excel_write_range')));
                    if (needAutoFreq && toolRegistry.get('excel_write_range')) {
                        // 选维度列：在"非数值、非日期"的分类列里，选唯一值种类最多的那列做占比维度
                        //（信息量最大、饼图扇区更有意义）。日期/序号/纯数字列都会产生无意义占比，跳过。
                        const DATE_RE = /^\s*\d{4}[-/年]\d{1,2}[-/月]/;
                        const DATE_HDR_RE = /日期|时间|date|time/i;
                        const candidates = [];
                        for (let hi = 0; hi < snapHeaders.length; hi++) {
                            const h = String(snapHeaders[hi] || '').trim();
                            const isNumericCol = colNumeric[hi] === true;
                            const sample = snapRows.map(r => r && r[h]).find(v => v !== null && v !== undefined && v !== '');
                            const looksDate = DATE_HDR_RE.test(h) ||
                                (sample !== undefined && DATE_RE.test(String(sample)));
                            if (!isNumericCol && !looksDate) {
                                const uniq = new Set(
                                    snapRows.map(r => (r && r[h] !== undefined && r[h] !== null && String(r[h]) !== '') ? String(r[h]) : null)
                                        .filter(v => v !== null)
                                ).size;
                                candidates.push({ h, uniq });
                            }
                        }
                        // 优先唯一值种类最多者；无合适候选则退回首个可用列
                        let dimHeader = '';
                        if (candidates.length) {
                            candidates.sort((a, b) => b.uniq - a.uniq);
                            dimHeader = candidates[0].h;
                        } else {
                            dimHeader = String(snapHeaders[0] || C.UI_DEFAULTS.dimHeader);
                        }
                        let vals = snapRows
                            .map(r => (r && r[dimHeader] !== undefined) ? r[dimHeader] : null)
                            .filter(v => v !== null && v !== undefined && v !== '');
                        // 伪表头修复：hasHeaders 读取把首行数据当表头时，该"表头值"会在数据中重复出现
                        //（分类列特征）→ 它其实是数据，补回频次统计，维度名改用通用"类别"。
                        if (vals.includes(dimHeader)) {
                            vals.unshift(dimHeader);
                            dimHeader = C.UI_DEFAULTS.dimHeader;
                        }
                        const freq = new Map();
                        for (const v of vals) {
                            const k = String(v);
                            freq.set(k, (freq.get(k) || 0) + 1);
                        }
                        const entries = [...freq.entries()].sort((a, b) => b[1] - a[1]);
                        const pureText = !colNumeric.some(Boolean);
                        yield { type: 'status', text: pureText
                            ? `选区为纯文本列，按「${dimHeader}」统计频次（${entries.length} 类）后绘图...`
                            : `识别为分类型图表，自动按「${dimHeader}」统计频次（${entries.length} 类），基于 ${dataRange || '当前选区'} 绘图...` };
                        const wres = await toolRegistry.invokeTracked(turn, 'excel_write_range', {
                            sheet: C.SHEET_CHART_DATA,
                            range: 'A1',
                            values: [[dimHeader, C.UI_DEFAULTS.valueHeader], ...entries]
                        }, {
                            displayName: '写入图表数据源',
                            displayArgs: { '工作表': C.SHEET_CHART_DATA, '内容': `${dimHeader}×${C.UI_DEFAULTS.valueHeader}，共 ${entries.length} 类` },
                            ctx: { turn, sessionStore, signal }
                        });
                        if (wres && wres.ok !== false && !wres.error) {
                            dataRange = `${C.SHEET_CHART_DATA}!A1:B${entries.length + 1}`;
                            displayRange = `${dataRange}（${dimHeader} 频次统计）`;
                        } else {
                            ChatLog.warn?.('PIPELINE', '[skill-chart] 频次聚合写入失败，按原范围绘图（后端将诚实校验）:', wres?.error || wres);
                        }
                    }
                }
            } catch (e) {
                if (e?.message === 'user_aborted') throw e;
                ChatLog.warn?.('PIPELINE', '[skill-chart] 自动频次聚合判定异常（不阻断）:', e?.message);
            }
            const chartArgs = {
                chartType: plan.chartType,
                // 给后端也传入推理后的默认标题，让 Excel 图表真正显示标题（空串=不设置）
                title: effectiveTitle || undefined,
                dataRange: dataRange,
                sheet: plan.sheet || undefined,
                showLegend: plan.showLegend,
                showDataLabels: plan.showDataLabels
            };
            if (plan.comboChart && plan.secondaryValueColumn) {
                chartArgs.comboChart = true;
                chartArgs.secondaryValueColumn = plan.secondaryValueColumn;
            }

            // displayArgs 中文显示，替换死板的 "(自动)"
            displayChartType = `${chartMeta.name} (${plan.chartType || chartMeta.name})`;
            displayTitle = effectiveTitle || C.UI_DEFAULTS.chartTitle;
            displayRange = rangeForDisplay || dataRange;

            // 用户同时表达了"按XX汇总XX + 图表"时：
            //   1. 先 excel_aggregate(createChart=false) 分组聚合 → 得聚合结果表 dataRange
            //   2. 再 excel_create_chart 在聚合结果上应用 LLM 规划的美化参数画图
            // 这样"LLM 规划图表类型+美化参数"由 Skill 层负责，"聚合"由 excel_aggregate 负责，
            // "画图"由 excel_create_chart 负责 —— 各司其职，且多图表时不互相污染选区。
            let aggregateReady = false;
            let aggregateDataRange = null;
            let aggregateDim = '';
            let aggregateVal = '';
            let aggregateFunc = '';
            let aggregateHandled = false;  // true=聚合已自带画图（旧后端），跳过 excel_create_chart

            // ── 图表类型切换后续继承：用户说"换成柱状图/改成折线图"等仅换类型的后续消息时，
            //    复用上一轮图表的数据口径（聚合快照或数据源范围），只换 chartType，不重读选区/重新聚合。
            //    判定用原始消息（turn._originalUserMessage），避免被 applySkillFollowUp 追加的上一轮结果
            //    （含"按部门求和"等字样）误触发 REDIMENSION 阻却。
            const followUpMsg = turn._originalUserMessage || userMessage;
            const inheritedAgg = ctx._aggregateResult || null;
            const chartSrcRange = ctx._chartDataRange || null;
            const isChartTypeSwitch = C.CHART_TYPE_SWITCH_RE.test(followUpMsg) &&
                !C.AGGREGATE_REDIMENSION_RE.test(followUpMsg) &&
                (!!inheritedAgg || !!chartSrcRange);
            if (isChartTypeSwitch) {
                if (inheritedAgg && (inheritedAgg.dataValues || inheritedAgg.dataRange)) {
                    // 上一轮图表是聚合结果：直接继承聚合快照（内存数组或落盘表），仅改类型
                    aggregateDim = inheritedAgg.dim || '';
                    aggregateVal = inheritedAgg.val || '';
                    aggregateFunc = inheritedAgg.func || 'sum';
                    aggregateReady = true;
                    aggregateHandled = false;
                    if (Array.isArray(inheritedAgg.dataValues)) {
                        aggregateSnapshot = {
                            dataValues: inheritedAgg.dataValues,
                            dim: aggregateDim, val: aggregateVal, func: aggregateFunc
                        };
                    } else if (inheritedAgg.dataRange) {
                        aggregateDataRange = inheritedAgg.dataRange;
                        dataRange = inheritedAgg.dataRange;
                        displayRange = inheritedAgg.dataRange;
                        aggregateSnapshot = { dataRange: inheritedAgg.dataRange, dim: aggregateDim, val: aggregateVal, func: aggregateFunc };
                    }
                    if (plan && typeof plan === 'object') {
                        // ★ 必须把 dataValues 透传给 _resolvedAgg：
                        //    finalize 气泡预览只认 _aggInfo.dataValues（finalize 闭包作用域安全，
                        //    无法访问非 finalize 声明的 aggregateSnapshot）。继承场景若漏传，
                        //    则"换折线图"等类型切换不产生气泡预览。
                        plan._resolvedAgg = {
                            ready: true, dim: aggregateDim, val: aggregateVal, func: aggregateFunc,
                            dataRange: aggregateDataRange || null,
                            dataValues: Array.isArray(aggregateSnapshot?.dataValues) ? aggregateSnapshot.dataValues : undefined
                        };
                        // 继承场景没有经过聚合/快照回填，读取数据步骤仍需真实来源范围：
                        // 用 chartSrcRange 补进 chartSnapMeta，让"读取数据"展示真实地址而非"当前选定区域"。
                        if (chartSrcRange) {
                            if (!chartSnapMeta || typeof chartSnapMeta !== 'object') chartSnapMeta = {};
                            if (!chartSnapMeta.range) chartSnapMeta.range = chartSrcRange.replace(/\$/g, '');
                            if (!chartSnapMeta.worksheet && plan && plan.sheetName) chartSnapMeta.worksheet = plan.sheetName;
                        }
                    }
                    // 展示用来源范围：优先数据源快照，便于"读取数据"步骤显示真实地址
                    if (chartSrcRange && (!displayRange || String(displayRange).includes('当前'))) {
                        displayRange = chartSrcRange;
                        if (!dataRange) dataRange = chartSrcRange;
                    }
                    const aggLabel = aggregateFunc === 'count' ? '计数' : (aggregateFunc === 'avg' || aggregateFunc === 'average') ? '求平均' : (aggregateFunc === 'max' ? '取最大' : (aggregateFunc === 'min' ? '取最小' : '求和'));
                    yield { type: 'status', text: `继承上一轮「按「${aggregateDim}」${aggLabel}「${aggregateVal}」」的数据口径，仅切换图表类型...` };
                } else if (chartSrcRange) {
                    // 上一轮图表无聚合：直接复用其数据源范围，仅改类型
                    dataRange = chartSrcRange;
                    displayRange = chartSrcRange;
                    yield { type: 'status', text: `复用上一张图的数据源（${chartSrcRange}），仅切换图表类型...` };
                }
            }

            // dataRange 兜底校验：类型切换继承已尝试补充数据源/聚合快照，
            // 仍未取得任何可绘制数据（无范围且无内存数组）才报错。必须在继承块之后执行。
            if (!dataRange && !(aggregateSnapshot && aggregateSnapshot.dataValues)) {
                throw new Error('未指定图表数据范围。请先在 Excel 中选中要绘制的数据区域（含表头），或明确说"用 XX 表 A1:B10 的数据"');
            }

            // 聚合分支判定用"用户原始消息"（followUpMsg），而非 applySkillFollowUp 富化后的
            // userMessage：富化消息夹带上轮聚合文本（如"按部门求和销售额"），会让纯类型切换的
            // "换成柱状图"误触发再聚合。且纯类型切换（isChartTypeSwitch）已在继承块补充数据，
            // 必须跳过此处重复聚合。
            if (!isChartTypeSwitch && typeof intentDetector.detectAggregateIntent === 'function' &&
                intentDetector.detectAggregateIntent(followUpMsg)) {
                // 优先复用上一轮聚合结果表：用户说"基于统计结果/上轮结果生成图表"时，
                // 若上一轮已聚合出结果表，直接引用其地址，避免重新按原数据聚合。
                const prevAgg = ctx._aggregateResult || null;
                const reuseAgg = prevAgg && prevAgg.dataRange &&
                    C.REUSE_AGGREGATE_RE.test(followUpMsg) &&
                    !C.AGGREGATE_REDIMENSION_RE.test(followUpMsg);
                if (reuseAgg) {
                    aggregateDataRange = prevAgg.dataRange;
                    dataRange = prevAgg.dataRange;
                    displayRange = prevAgg.dataRange;
                    aggregateReady = true;
                    yield { type: 'status', text: `复用上一轮聚合结果表（${prevAgg.dataRange}），直接生成图表...` };
                } else {
                    const dims = (intentDetector.detectDimensionColumns || (() => []))(followUpMsg) || [];
                    const vals = (intentDetector.detectValueColumns || (() => []))(followUpMsg) || [];
                    const aggDef = (intentDetector.detectAggFunc || (() => ({ func: 'sum', label: '求和' })))(followUpMsg);
                    if (dims.length >= 1 && vals.length >= 1 && toolRegistry.get('excel_aggregate')) {
                    aggregateDim = dims[0];
                    aggregateVal = vals[0];
                    aggregateFunc = (aggDef && aggDef.func) || 'sum';
                    aggregateReady = true;
                    // 把解析出的聚合口径持久到 plan（全函数可见），供 finalize 阶段透明回写骨架步骤。
                    // 不能用闭包变量 aggregate*：它们声明在 chart 内层块，finalize 在更外层块，块外访问 let 会 ReferenceError。
                    if (plan && typeof plan === 'object') {
                        plan._resolvedAgg = { ready: true, dim: aggregateDim, val: aggregateVal, func: aggregateFunc };
                    }
                    yield { type: 'status', text: `检测到聚合需求：按「${aggregateDim}」${aggDef?.label || '汇总'}「${aggregateVal}」，先分组聚合再创建图表...` };

                    // 第 1 步：分组聚合（只写聚合结果表，不内置画图）
                    // 聚合与图表必须使用同一数据源（消息发送时的选区快照 originalDataRange），
                    // 否则用户移开选区后 excel_aggregate 按实时 Selection 读到空区域而失败，
                    // 而 excel_create_chart 按快照地址仍能画图 → "聚合失败但图表成功"的矛盾。
                    // 不能用 dataRange：分类型"自动频次统计"已将其覆盖为"图表数据源"表（销售员×数量），
                    // 会导致聚合 range 落在频次表上、"销售额"列缺失而失败。
                    // range 只传地址部分（后端用 srcWs.Range 定位，sheet 由 srcWs 决定）。
                    let aggRangeAddr = originalDataRange || null;
                    if (aggRangeAddr && String(aggRangeAddr).includes('!')) {
                        aggRangeAddr = String(aggRangeAddr).split('!').slice(1).join('!');
                    }
                    const aggArgs = {
                        dimColumn: aggregateDim,
                        valueColumn: aggregateVal,
                        aggFunc: aggregateFunc,
                        chartType: plan.chartType,
                        createChart: false,
                        range: aggRangeAddr || undefined
                        // 内存绑定：不传 targetSheet → 后端走纯内存聚合（不落任何工作表），
                        // 返回 JSON 中的 data 由前端构造成 dataValues 直接内存绑定系列，
                        // 彻底消除"聚合结果/图表数据源"残留工作表。
                    };
                    // 解析聚合结果（主路径与兜底重试共用）：
                    //   1. 后端返回聚合 data（对象行/数组行）→ 构造成 agg2d → dataValues 内存绑定（主路径，零落盘）
                    //   2. 旧后端已把结果落地（writtenTo）→ dataRange 指向落盘表（兼容分支，仍优先内存数组）
                    //   3. 完全不可解析（真·旧后端纯文本且已内置画图）→ aggregateHandled 跳过二次画图
                    const applyAggregateResult = async (res) => {
                        const aggPayload = res?.result || res;
                        let meta = aggPayload;
                        if (typeof meta === 'string') { try { meta = JSON.parse(meta); } catch { meta = null; } }
                        if (!meta || typeof meta !== 'object') {
                            ChatLog.info?.('PIPELINE', '[skill-chart] 后端未返回可解析聚合结果（旧后端），聚合已内置画图，跳过二次画图');
                            aggregateHandled = true;
                            return;
                        }
                        // ★ 聚合返回真实源范围（ExcelMcpTools.Advanced.Aggregate output("range") = srcRange.Address），
                        // 回填 chartSnapMeta 供"读取数据"步骤展示真实地址（无选区/快照缺 range 时尤为重要——
                        // read_range 历史上可能无 range 字段，见 project_memory P0，而聚合一定有真实 UsedRange 范围）。
                        try {
                            const _aggRange = meta.range && String(meta.range).length ? String(meta.range) : '';
                            if (_aggRange && (!chartSnapMeta || typeof chartSnapMeta !== 'object' || !chartSnapMeta.range)) {
                                if (!chartSnapMeta || typeof chartSnapMeta !== 'object') chartSnapMeta = {};
                                if (!chartSnapMeta.range) chartSnapMeta.range = _aggRange;
                                if (!chartSnapMeta.worksheet && meta.worksheet) chartSnapMeta.worksheet = meta.worksheet;
                                if (chartSnapMeta.rows == null && typeof meta.rows === 'number') chartSnapMeta.rows = meta.rows;
                                if (chartSnapMeta.columns == null && typeof meta.columns === 'number') chartSnapMeta.columns = meta.columns;
                            }
                        } catch (e) { /* 展示优化，不阻断 */ }
                        const writtenTo = meta.writtenTo || meta.resultSheet || null;
                        // 区分两种 data 形态：对象行（维度聚合，无表头行）vs 数组行（日期分组，首行即表头）
                        const rawData = Array.isArray(meta.data) ? meta.data : null;
                        const isArrayRows = Array.isArray(rawData) && Array.isArray(rawData[0]);
                        const groupCount = Number(meta.groupCount) || (Array.isArray(rawData)
                            ? (isArrayRows ? Math.max(0, rawData.length - 1) : rawData.length) : 0);
                        // 内存二维数组：后端返回的聚合 data 直接构造成 [表头, ...数据]，
                        // 供 excel_create_chart 的 dataValues 内存绑定，避免把聚合结果落地成工作表。
                        let agg2d = null;
                        if (Array.isArray(rawData) && rawData.length > 0) {
                            try {
                                const firstRow = rawData[0] || {};
                                if (isArrayRows) {
                                    // 数组行（日期分组）：首行即表头，其后为数据；仅取分类|数值两列
                                    const dimKey = (firstRow[0] != null ? firstRow[0] : C.UI_DEFAULTS.dimHeader);
                                    const valKey = (firstRow[1] != null ? firstRow[1] : C.UI_DEFAULTS.valueHeader);
                                    agg2d = [dimKey, valKey].concat(
                                        rawData.slice(1).map(r => [
                                            (r && r[0] !== undefined && r[0] !== null) ? r[0] : '',
                                            (r && r[1] !== undefined && r[1] !== null) ? r[1] : ''
                                        ])
                                    );
                                } else {
                                    const keys = Object.keys(firstRow);
                                    const dimKey = (meta.dimColumn && keys.includes(meta.dimColumn)) ? meta.dimColumn : (keys[0] || C.UI_DEFAULTS.dimHeader);
                                    const valKey = (meta.valueColumn && keys.includes(meta.valueColumn) && meta.valueColumn !== dimKey)
                                        ? meta.valueColumn : (keys.find(k => k !== dimKey) || C.UI_DEFAULTS.valueHeader);
                                    agg2d = [[dimKey, valKey]].concat(rawData.map(r => [
                                        (r && r[dimKey] !== undefined && r[dimKey] !== null) ? r[dimKey] : '',
                                        (r && r[valKey] !== undefined && r[valKey] !== null) ? r[valKey] : ''
                                    ]));
                                }
                            } catch (e) { agg2d = null; }
                        }
                        if (writtenTo && groupCount > 0) {
                            // 后端落盘了聚合结果表：dataRange 指向它（用于展示），但画图仍优先内存数组
                            aggregateDataRange = `${writtenTo}!A1:B${groupCount + 1}`;
                        } else if (agg2d && groupCount > 0) {
                            // 后端未落盘但返回了内存数据：无需补写任一工作表，直接内存绑定
                            aggregateDataRange = null;
                        }
                        // 用聚合结果表覆盖 dataRange，让 excel_create_chart 在聚合结果上画图
                        if (agg2d) {
                            // 内存数据模式：聚合结果直接以数组交给 excel_create_chart，不落地任何工作表
                            aggregateSnapshot = {
                                dataValues: agg2d,
                                dim: aggregateDim,
                                val: aggregateVal,
                                func: aggregateFunc
                            };
                            if (plan && typeof plan === 'object' && plan._resolvedAgg) {
                                plan._resolvedAgg.dim = aggregateDim;
                                plan._resolvedAgg.val = aggregateVal;
                                plan._resolvedAgg.func = aggregateFunc;
                                plan._resolvedAgg.dataValues = agg2d;
                            }
                        } else if (aggregateDataRange) {
                            dataRange = aggregateDataRange;
                            displayRange = aggregateDataRange;
                            aggregateSnapshot = {
                                dataRange: aggregateDataRange,
                                dim: aggregateDim,
                                val: aggregateVal,
                                func: aggregateFunc
                            };
                            if (plan && typeof plan === 'object' && plan._resolvedAgg) {
                                plan._resolvedAgg.dataRange = aggregateDataRange;
                            }
                        } else {
                            // 无落盘地址且无内存数据 → 旧后端聚合已自带画图，跳过 excel_create_chart
                            ChatLog.info?.('PIPELINE', '[skill-chart] 聚合无可用数据（旧后端已内置画图），跳过二次画图');
                            aggregateHandled = true;
                        }
                    };
                    const aggResult = await toolRegistry.invokeTracked(turn, 'excel_aggregate', aggArgs, {
                        round: 1,
                        displayName: `按${aggregateDim}${aggDef?.label || '汇总'}${aggregateVal}`,
                        displayArgs: { '聚合维度': aggregateDim, '聚合指标': aggregateVal, '聚合方式': aggDef?.label || aggregateFunc },
                        ctx: { turn, sessionStore, signal }
                    });
                    if (signal?.aborted) throw new Error('user_aborted');
                    if (aggResult?.ok !== false && !aggResult?.error) {
                        await applyAggregateResult(aggResult);
                    } else {
                        const aggErrStr = String(aggResult?.error || aggResult?.result || '');
                        // 选区无效类错误：快照地址失效（数据被删/移走）时，先无 range 兜底重试一次，
                        // 让后端走 RangeContext → Selection → UsedRange 回退链（数据兜底）；
                        // 兜底也失败才中止并提示重新选择，避免"聚合失败但图表按旧数据成功"的矛盾。
                        if (C.AGGREGATE_INVALID_RANGE_RE.test(aggErrStr)) {
                            ChatLog.warn?.('PIPELINE', '[skill-chart] 快照地址失效，尝试无 range 兜底重试（UsedRange）...');
                            const retryArgs = {
                                dimColumn: aggregateDim,
                                valueColumn: aggregateVal,
                                aggFunc: aggregateFunc,
                                chartType: plan.chartType,
                                createChart: false
                            };
                            const retryResult = await toolRegistry.invokeTracked(turn, 'excel_aggregate', retryArgs, {
                                round: 1,
                                displayName: `按${aggregateDim}${aggDef?.label || '汇总'}${aggregateVal}（兜底重试）`,
                                displayArgs: { '聚合维度': aggregateDim, '聚合指标': aggregateVal, '聚合方式': aggDef?.label || aggregateFunc },
                                ctx: { turn, sessionStore, signal }
                            });
                            if (signal?.aborted) throw new Error('user_aborted');
                            if (retryResult?.ok !== false && !retryResult?.error) {
                                await applyAggregateResult(retryResult);
                            } else {
                                throw new Error(`聚合分析失败：${aggErrStr || '未选中有效数据区域'}。请重新选中包含表头的数据区域后再试。`);
                            }
                        } else {
                            ChatLog.warn?.('PIPELINE', '[skill-chart] excel_aggregate 聚合失败，回退到直接在当前选区画图:', aggResult?.error);
                            // 聚合失败回退时，绘制范围必须复位到原始数据范围，而非被"频次统计"
                            // 覆盖的图表数据源表（销售员×数量），否则回退图表仍是错误口径。
                            // 注：此时 dataRange 可能指向聚合结果表（聚合部分成功）或图表数据源表，
                            // 统一复位为原始选区，保证回退图表与用户意图的数据源一致。
                            dataRange = originalDataRange;
                        }
                    }
                    }
                }
            }
            // 内存数据模式：聚合结果直接在内存（dataValues），不发 dataRange、不落任何工作表
            if (aggregateSnapshot && aggregateSnapshot.dataValues) {
                chartArgs.dataValues = JSON.stringify(aggregateSnapshot.dataValues);
                chartArgs.dataRange = undefined;
                delete chartArgs.dataRange;
            } else {
                chartArgs.dataRange = dataRange;
            }
            // 【图表数据源贯通 · skill-pipeline】dataRange 被覆盖为派生表
            // （聚合结果/图表数据源，带 "!" 前缀）时，sheet 必须同步为前缀表名：
            // 后端 ResolveRangeSafe 显式 sheet 优先于 dataRange 前缀，若 chartArgs.sheet
            // 仍是 plan.sheet（原表），后端会读原表 A1:Bn → 文本列 → "没有可绘制的数值列"。
            // 实测复现："使用饼图按部门展示销售额" 第二步聚合成功（聚合结果!A1:B6），
            // 第三步图表却按 plan.sheet=销售数据 读原表 → 报错（显式 sheet 压过前缀）。
            // 内存数据模式不涉及派生表，跳过该派生逻辑。
            if (typeof dataRange === 'string' && dataRange.includes('!') &&
                !(aggregateSnapshot && aggregateSnapshot.dataValues)) {
                let prefix = dataRange.split('!')[0].trim();
                if (prefix.startsWith("'") && prefix.endsWith("'")) prefix = prefix.slice(1, -1);
                if (prefix && C.isValidSheetName(prefix)) {
                    chartArgs.sheet = prefix;
                }
            }

            // 图表工具未注册兜底三级保障：
            //   1. 先重新确保 MCP 工具已同步（首次打开页面时 App.mcpTools 可能还未填充）
            //   2. 工具仍未注册 → 直接调用 App.executeMcpToolCall（绕过注册表）
            //   3. executeMcpToolCall 也不可用 → 直接底层 bridge CallMcpToolAsync
            try {
                if (typeof toolRegistry?.init === 'function' && !toolRegistry.get('excel_create_chart')) {
                    ChatLog.info?.('PIPELINE', '[skill-chart] excel_create_chart 未注册，尝试重新同步 MCP 工具...');
                    await toolRegistry.init();
                }
            } catch (e) {
                ChatLog.warn?.('PIPELINE', '[skill-chart] 重新同步 MCP 工具异常:', e?.message);
            }

            // 旧后端聚合已内置画图（aggregateHandled=true），跳过 excel_create_chart
            if (aggregateHandled) {
                execResult = { ok: true, result: { success: true, message: '聚合已内置创建图表' }, toolCallId: null };
                ChatLog.info?.('PIPELINE', '[skill-chart] 聚合已内置画图，跳过 excel_create_chart');
            } else if (toolRegistry.get('excel_create_chart')) {
                execResult = await toolRegistry.invokeTracked(turn, 'excel_create_chart', chartArgs, {
                    round: 1,
                    displayName: `${chartMeta.icon || '📊'} ${chartMeta.name}  ${displayName}`,
                    displayArgs: { '图表类型': displayChartType, '标题': displayTitle, '数据范围': displayRange },
                    ctx: { turn, sessionStore, signal }
                });
            } else {
                // 兜底：工具未注册，直接调用 executeMcpToolCall 或 bridge
                ChatLog.warn?.('PIPELINE', '[skill-chart] excel_create_chart 工具未注册，尝试绕过注册表直接调用后端 MCP...');
                let addedToolCallId = null;
                try {
                    const added = sessionStore.addToolCall?.(turn.id, {
                        round: 1,
                        name: 'excel_create_chart',
                        displayName: `${chartMeta.icon || '📊'} ${chartMeta.name}  ${displayName}（直连）`,
                        displayArgs: { '图表类型': displayChartType, '标题': displayTitle, '数据范围': displayRange },
                        args: chartArgs,
                        status: 'running',
                        startedAt: Date.now()
                    });
                    addedToolCallId = added?.id || null;
                } catch (e) { /* 忽略 UI 注册失败 */ }
                const rawStart = Date.now();
                let rawResult = null;
                try {
                    if (typeof App.executeMcpToolCall === 'function') {
                        const raw = await App.executeMcpToolCall({
                            function: { name: C.MCP_CHART_TOOL_FQN, arguments: chartArgs }
                        }, { turn, sessionStore, signal, tool: { name: 'excel_create_chart' } });
                        rawResult = (typeof raw === 'string') ? (() => { try { return JSON.parse(raw); } catch { return raw; } })() : raw;
                    } else if (typeof App.excelBridge?.CallMcpToolAsync === 'function') {
                        const rawStr = await App.excelBridge.CallMcpToolAsync(C.MCP_SERVER, 'excel_create_chart', JSON.stringify(chartArgs));
                        rawResult = (typeof rawStr === 'string') ? (() => { try { return JSON.parse(rawStr); } catch { return rawStr; } })() : rawStr;
                    } else {
                        throw new Error('Tool not found: excel_create_chart（且 MCP 直连接口 App.executeMcpToolCall / CallMcpToolAsync 均不可用）');
                    }
                    const dur = Date.now() - rawStart;
                    const ok = rawResult && (rawResult.ok !== false) && !rawResult.error;
                    const normResult = {
                        ok,
                        result: rawResult?.result !== undefined ? rawResult.result : rawResult,
                        error: rawResult?.error || null,
                        durationMs: dur,
                        toolCallId: addedToolCallId
                    };
                    const patch = {
                        status: ok ? 'completed' : 'failed',
                        result: normResult.result,
                        error: normResult.error,
                        durationMs: dur,
                        completedAt: Date.now()
                    };
                    if (addedToolCallId && typeof sessionStore.updateToolCall === 'function') {
                        try { sessionStore.updateToolCall(turn.id, addedToolCallId, patch, { silent: false }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec-execute] updateToolCall 完成态更新失败:', e); }
                    }
                    execResult = normResult;
                } catch (callErr) {
                    const dur = Date.now() - rawStart;
                    execResult = { ok: false, error: callErr?.message || String(callErr), durationMs: dur, toolCallId: addedToolCallId };
                    if (addedToolCallId && typeof sessionStore.updateToolCall === 'function') {
                        try { sessionStore.updateToolCall(turn.id, addedToolCallId, { status: 'failed', error: execResult.error, durationMs: dur, completedAt: Date.now() }, { silent: false }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec-execute] updateToolCall 失败态更新失败:', e); }
                    }
                }
            }
            if (signal?.aborted) throw new Error('user_aborted');
            if (!execResult?.ok) {
                throw new Error(execResult?.error || '图表创建失败');
            }
        } else if (isCleaner) {
            // ── 清洗工具链（纯 MCP，无 VBA）──
            // 🛡️ 强制收敛：结构变换 op（splitByColumn / mergeSheets 等 5 个）一旦出现就独占执行，
            // 丢弃 LLM 自由脑补的清洗 op（removeDuplicates / fillEmpty）或分析/聚合/图表工具。
            // 根因：LLM 看到"合并/拆分"常脑补"先聚合→画图"，导致 3 步无关工具 + 0 步真正工具。
            const STRUCTURE_OPS_SET = new Set(['splitByColumn', 'splitWorkbook', 'splitByRows', 'mergeSheets', 'mergeWorkbooks']);
            const rawOps = plan.operations || [];
            const structureHit = rawOps.find(op => STRUCTURE_OPS_SET.has(op));
            const CLEAN_OPS = structureHit ? [structureHit] : rawOps;
            if (structureHit && rawOps.length > 1) {
                ChatLog.warn?.('PIPELINE', `[skill-cleaner] 规划含结构变换 op "${structureHit}"，强制收敛丢弃混排 op: [${rawOps.filter(o => !STRUCTURE_OPS_SET.has(o)).join(', ')}]`);
            }
            // 🛡️ 提示词也显示收敛后的 plan，避免面板出现"已按「聚合、图表」规划"这种误导
            const ops = CLEAN_OPS.map(op => C.cleanOpLabel(op) || op);
            planSummary = (ops.length > 0 ? ops.join('、') : '格式整理');
            yield { type: 'status', text: `已按「${planSummary}」规划，正在执行...` };

            const SUM_FIELDS = new Set(C.CLEAN_SUM_FIELDS);
            const KEEP_FIRST_FIELDS = new Set(C.CLEAN_KEEP_FIRST_FIELDS);
            const isNum = v => typeof v === 'number' && !Number.isNaN(v);
            const aggregatedMeta = {};
            let allSteps = [];
            let lastToolCallId = null;

            // ✨ 维护 subSteps 数组（初始值来自骨架 earlySkeleton[0].subSteps）
            const subSteps = (earlySkeleton[0]?.subSteps || []).map(s => ({ ...s }));

            // buildSpecSteps('clean') 已在骨架第 0 步放入"读取数据"（steps[0]），
            // 这里只引用它并对数据区做一次真实读取（读取兼探测，仅一次 excel_read_range），
            // 验证行进到小步骤第一步。
            // 注意：不可再 unshift 一个"读取数据"，否则会出现两个读取步骤、且后续
            // subSteps[i+1] 的下标对齐会整体错位（每一步结果显示到上一步）。
            // ✨ 读取兼探测：仅当没有清晰数据区（未选中/单格）且用户未显式写死区域时，
            //    在读取步骤内先探测真实数据区并回填 ctx.selection 供后续 excel_clean_data 使用。
            if (!explicitRange && !(ctx.selection?.address && (Number(ctx.selection.rowCount) > 1 || Number(ctx.selection.colCount) > 1))) {
                try {
                    const probed = await SP.specBuild.detectDataSelection(ctx, turn, toolRegistry, sessionStore, signal);
                    if (probed) {
                        ctx.selection = probed;
                        probedInfo = probed;
                    }
                } catch (e) {
                    ChatLog.warn?.('PIPELINE', '[clean] 读取兼自动探测异常:', e && e.message ? e.message : e);
                }
            }
            const selAddr0 = (ctx.selection?.address || '').replace(/\$/g, '');
            const selSheet0 = ctx.selection?.sheetName || '';
            if (!subSteps.length || subSteps[0]?.title !== '读取数据') {
                subSteps.unshift({ id: 'read-data', title: '读取数据', detail: `选区 ${selAddr0 || '当前选中区域'}${selSheet0 ? ' · ' + selSheet0 : ''}`, status: 'pending' });
            }
            subSteps[0] = { ...subSteps[0], title: '读取数据', status: 'running' };
            yield { type: 'progress', stepIndex: 0, status: 'running', subSteps: [...subSteps] };
            // ✨ 读取兼探测：未选中/单格时，上面的探测(796)已用一次 excel_read_range 读到整表数据(rawResult)，
            // 这里直接复用该结果作为"读取数据"步骤的内容，不再二次调用工具 → 全程仅一次 excel_read_range；
            // 已有清晰选区则由前端回传 address 正常构造读取（无 range，后端按选区注入）。
            const probeReused = !!(probedInfo && probedInfo.address && probedInfo.rawResult != null);
            let readPre;
            if (probeReused) {
                readPre = { ok: true, result: probedInfo.rawResult };
            } else {
                const readArgs = { returnFormat: 'array', hasHeaders: false };
                if (selSheet0) readArgs.sheet = selSheet0;
                if (explicitRange && explicitRange.address) readArgs.range = explicitRange.address;
                readPre = await toolRegistry.invokeTracked(turn, 'excel_read_range', readArgs, {
                    displayName: '读取数据',
                    displayArgs: { '范围': explicitRange && explicitRange.address ? explicitRange.address : (selAddr0 || '当前选区'), '工作表': selSheet0 || '活动表' },
                    ctx: { turn, sessionStore, signal }
                });
            }
            if (signal?.aborted) throw new Error('user_aborted');
            try {
                const rawT = typeof readPre?.result === 'string' ? readPre.result : JSON.stringify(readPre?.result);
                console.log(`[READ-DIAG] selAddr="${selAddr0}" selSheet="${selSheet0}" ok=${readPre?.ok!==false && !readPre?.error} error=${readPre?.error||''} raw=${(rawT||'').slice(0,500)}`);
            } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-spec-execute] READ-DIAG 调试日志输出失败:', e); }
            const readRows = (readPre?.ok === false || readPre?.error) ? [] : (SP.cleanCrud.parseReadRows(readPre?.result) || []);
            const rowArr = Array.isArray(readRows) ? readRows : [];
            const colN = rowArr.length ? Math.max(0, ...rowArr.map(r => (Array.isArray(r) ? r.length : 0))) : 0;
            if (rowArr.length === 0 && (readPre?.ok !== false && !readPre?.error)) {
                // 显式配置读到空但后端未报错：改用"无 range"的实时选区兜底再读一次
                const readPre2 = await toolRegistry.invokeTracked(turn, 'excel_read_range', { returnFormat: 'array', hasHeaders: false }, {
                    displayName: '读取数据(回退)-', displayArgs: { '方式': '实时选区' },
                    ctx: { turn, sessionStore, signal }
                });
                const rowArr2 = (readPre2?.ok === false || readPre2?.error) ? [] : (SP.cleanCrud.parseReadRows(readPre2?.result) || []);
                if (readPre2 && readPre2.ok !== false && !readPre2.error && rowArr2.length > 0) {
                    const c2 = Math.max(0, ...rowArr2.map(r => (Array.isArray(r) ? r.length : 0)));
                    subSteps[0] = { ...subSteps[0], status: 'completed', detail: `已读取 ${rowArr2.length} 行 × ${c2} 列（实时选区）` };
                } else {
                    subSteps[0] = { ...subSteps[0], status: 'failed', detail: `读取为空（${rowArr.length} 行）；已试实时选区兜底仍空，请确认已选中数据区域且活动工作表正确` };
                }
            } else if (readPre?.ok === false || readPre?.error) {
                subSteps[0] = { ...subSteps[0], status: 'failed', detail: `读取失败：${String(readPre?.error || '未知错误').slice(0, 80)}` };
            } else {
                subSteps[0] = { ...subSteps[0], status: 'completed', detail: `已读取 ${rowArr.length} 行 × ${colN} 列${probedInfo && probedInfo.address ? `（自动探测 ${probedInfo.address}）` : ''}` };
            }
            // 读取/探测步骤带工具行，展开可见 excel_read_range + 真实范围
            subSteps[0] = { ...subSteps[0], toolName: subSteps[0].toolName || 'excel_read_range', params: probeReused ? { '方式': '读取兼自动探测', '范围': probedInfo?.address || selAddr0 } : (subSteps[0].params || null) };
            yield { type: 'progress', stepIndex: 0, status: 'running', subSteps: [...subSteps] };

            for (let i = 0; i < CLEAN_OPS.length; i++) {
                const op = CLEAN_OPS[i];
                // 标记第 i 个小步骤 running（+1 因前插了"读取数据"步骤）
                subSteps[i + 1] = { ...(subSteps[i + 1] || {}), status: 'running', title: subSteps[i + 1]?.title || op };
                yield { type: 'progress', stepIndex: 0, status: 'running', subSteps: [...subSteps] };

                const single = await SP.cleanCrud.executeCleanSingleOp(turn, op, plan, ctx, toolRegistry, sessionStore, signal);
                if (signal?.aborted) throw new Error('user_aborted');

                if (!single.ok) {
                    subSteps[i + 1] = { ...subSteps[i + 1], status: 'failed', detail: single.error || '执行失败' };
                    yield { type: 'progress', stepIndex: 0, status: 'running', subSteps: [...subSteps] };
                    if (single.cancelled) {
                        // 同 CRUD 分支：安全守卫拦截/用户取消 → 优雅收尾，不降级 agent-loop 二次执行。
                        // excel_remove_duplicates / excel_clean_data 均为危险工具（catalog medium），
                        // headless 或用户点取消都会走到这里；此前直接 throw → fallbackToText → ok=true
                        // 但清洗工具从未执行（skill-chart-clean-e2e 实测挂点）。
                        specCancelled = true;
                        cancelMessage = single.error || '已取消，未执行任何操作。';
                        ChatLog.info?.('PIPELINE', `[skill-spec] 清洗操作 ${op} 被安全守卫拦截/用户取消，优雅收尾（不降级）`);
                        break;
                    }
                    throw new Error(single.error || `清洗操作 ${op} 执行失败`);
                }

                // meta 累加
                for (const [key, val] of Object.entries(single.meta || {})) {
                    if (!isNum(val)) continue;
                    if (SUM_FIELDS.has(key)) aggregatedMeta[key] = (aggregatedMeta[key] || 0) + val;
                    else if (KEEP_FIRST_FIELDS.has(key) && aggregatedMeta[key] === undefined) aggregatedMeta[key] = val;
                }

                allSteps = allSteps.concat(single.steps || []);
                lastToolCallId = single.toolCallId || lastToolCallId;

                const summary = SP.cleanCrud.buildCleanOpSummary(op, single.meta || {}, plan, single.writeOk);
                try { console.log(`[CLEAN-STEP#${i}] op="${op}" meta=`, single.meta, `writeOk=${single.writeOk} → summary="${summary}"`); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-spec-execute] CLEAN-STEP 调试日志输出失败:', e); }

                const done = !!(single.writeOk);
                subSteps[i + 1] = { ...subSteps[i + 1], status: done ? 'completed' : 'failed', detail: done ? summary : (single.writeError || summary || '写回未成功') };
                yield { type: 'progress', stepIndex: 0, status: 'running', subSteps: [...subSteps] };
            }

            // 任一子步骤写回失败 → 大步骤也置 failed，整体 ok=false（如实反映，不再"假完成"）
            const hasFailed = subSteps.some(s => s.status === 'failed');
            yield { type: 'progress', stepIndex: 0, status: hasFailed ? 'failed' : 'completed', subSteps: [...subSteps] };

            // 构造 execResult
            execResult = {
                ok: !hasFailed,
                result: {
                    success: !hasFailed,
                    message: hasFailed ? '清洗部分步骤未写回成功' : (allSteps.length ? allSteps[allSteps.length - 1] : '清洗完成'),
                    steps: allSteps,
                    // 优先 removedRows，其次 deletedRows（不同后端可能返回不同字段名）
                    removedRows: aggregatedMeta.removedRows ?? null,
                    deletedRows: aggregatedMeta.deletedRows ?? null,
                    markedRows: aggregatedMeta.markedRows ?? null,
                    cleanedRows: aggregatedMeta.cleanedRows ?? null,
                    filledCells: aggregatedMeta.filledCells ?? null,
                    trimmedCells: aggregatedMeta.trimmedCells ?? null,
                    normalizedDates: aggregatedMeta.normalizedDates ?? null,
                    normalizedNumbers: aggregatedMeta.normalizedNumbers ?? null,
                    removedEmptyRows: aggregatedMeta.removedEmptyRows ?? null
                },
                // ✨ 把执行循环里更新好的 subSteps 带出去，finalize 用
                subSteps: subSteps.map(s => ({ ...s })),
                toolCallId: lastToolCallId
            };
        } else if (isCrud) {
            // ── 增删改查工具链（纯 MCP，无 VBA）──
            planSummary = SP.cleanCrud.opLabel(plan.operation) + (SP.cleanCrud.crudWhere(plan) ? ' · ' + SP.cleanCrud.crudWhere(plan) : '');
            yield { type: 'status', text: `已按「${planSummary}」规划，正在执行...` };
            const mcpCrud = await SP.cleanCrud.executeCrudWithMcp(turn, plan, ctx, toolRegistry, sessionStore, signal);
            if (signal?.aborted) throw new Error('user_aborted');
            if (mcpCrud.ok) {
                execResult = mcpCrud;
            } else if (mcpCrud.cancelled) {
                // 安全守卫拦截 / 用户在危险确认框点"取消"：优雅收尾，不降级 fallbackToText
                // （降级会让 agent-loop 拿同一条消息重跑，实测又去调 excel_read_range）。
                specCancelled = true;
                cancelMessage = mcpCrud.error || '已取消，未执行任何操作。';
                execResult = mcpCrud;
                ChatLog.info?.('PIPELINE', '[skill-spec] CRUD 被安全守卫拦截/用户取消，优雅收尾（不降级）');
            } else {
                throw new Error(mcpCrud.error || '增删改查执行失败');
            }
        } else if (isDashboard) {
            // ── 看板工具链（纯 MCP，无 VBA）──
            // 蓝图引擎：excel_read_range 读取 → 内存计算 KPI/图表聚合 → 原子工具施工
            // （create_sheet / write_range / format_range / create_chart dataValues+anchorRange）→
            // HTML 大屏（模板渲染 + excel_display_html_in_browser）。
            // ⭐️ 方向 A+C：骨架生成前已做预计算+预修正（_dashPrecomputed），
            //    有则直接复用跳过重复读数据，无则正常流程。
            const dashEngine = App.ChatV2.dashboardModules?.blueprintEngine || null;
            if (!dashEngine || typeof dashEngine.computeFromSources !== 'function' || typeof dashEngine.buildPlan !== 'function') {
                throw new Error('看板引擎未就绪（dashboard-blueprint-engine.js 未加载）');
            }
            // plan.title 可能是 LLM 抄的用户原话（"帮我生成excel大看板"）→ 统一净化
            {
                const _hr = App.ChatV2.dashboardModules?.htmlRenderer;
                if (_hr && typeof _hr.sanitizeBoardTitle === 'function') {
                    plan.title = _hr.sanitizeBoardTitle(plan.title, '数据看板');
                }
            }
            planSummary = `「${plan.title || '数据看板'}」· ${plan.theme || 'business_blue'} 主题`;
            yield { type: 'status', text: `已按「${plan.title || '数据看板'}」规划看板蓝图，正在施工...` };
            const _dashRecs = [];
            let sources = null;
            let prepared = null;
            let readSteps = null;

            if (_dashPrecomputed && _dashPrecomputed.sources && _dashPrecomputed.sources.length > 0) {
                // ✅ 复用预计算结果（方向 A+C 已完成读数据 + 自动修正）
                sources = _dashPrecomputed.sources;
                prepared = _dashPrecomputed.prepared;
                readSteps = dashEngine.buildReadSteps(plan, ctx);
                // 骨架里的"读数据"步骤标记为已完成（预计算阶段已读过）
                for (let ri = 0; ri < readSteps.length; ri++) {
                    yield { type: 'progress', stepIndex: ri, status: 'completed', detail: `${readSteps[ri].detail || readSteps[ri].title}（已预读）` };
                }
                ChatLog.info?.('PIPELINE', `[skill-dashboard] 复用预计算: sources=${sources.length} prepared.kpis=${prepared?.kpis?.length} charts=${prepared?.charts?.length}`);
            } else {
                // ❌ 无预计算 → 正常流程：读数据 → computeFromSources
                readSteps = dashEngine.buildReadSteps(plan, ctx);
                sources = [];
                for (let ri = 0; ri < readSteps.length; ri++) {
                    const st = readSteps[ri];
                    yield { type: 'progress', stepIndex: ri, status: 'running', detail: st.detail || st.title };
                    let r = null;
                    try {
                        r = await toolRegistry.invokeTracked(turn, st.tool, st.args, {
                            round: 1, displayName: st.displayName, displayArgs: st.displayArgs,
                            ctx: { turn, sessionStore, signal }
                        });
                    } catch (e) {
                        if (e?.message === 'user_aborted') throw e;
                        ChatLog.warn?.('PIPELINE', `[skill-dashboard] 读取数据源失败（跳过）:`, e?.message);
                    }
                    if (signal?.aborted) throw new Error('user_aborted');
                    _dashRecs.push({ tool: st.tool, showArgs: st.displayArgs || {}, raw: r });
                    let meta = r && (r.result != null ? r.result : r);
                    if (typeof meta === 'string') { try { meta = JSON.parse(meta); } catch { meta = null; } }
                    const headers = meta && Array.isArray(meta.headers) ? meta.headers.map(String) : [];
                    const rows = meta && (Array.isArray(meta.result) ? meta.result : Array.isArray(meta.rows) ? meta.rows : []);
                    if (headers.length && rows.length) {
                        sources.push({ id: (plan.dataSources && plan.dataSources[ri] && plan.dataSources[ri].id) || ('src_' + ri), headers, rows, range: meta.range || (st.args && st.args.range) || '', worksheet: meta.worksheet || (st.args && st.args.sheet) || '' });
                    }
                }
                prepared = dashEngine.computeFromSources(sources, plan);
            }

            // 3) 构造 Action 计划（读取步骤 + 渲染步骤，与骨架逐一对齐）
            const planSteps = dashEngine.buildPlan(prepared, plan, ctx);
            if (!Array.isArray(planSteps) || planSteps.length === 0) {
                throw new Error('看板蓝图未能生成施工步骤，请检查数据源是否有效');
            }
            const renderOffset = readSteps.length;
            // 4) 逐项执行渲染步骤；单步失败不阻断后续，错误汇总进总结
            let _dashHtmlErr = null;
            let _dashFailedTools = new Map();   // tool->errorStr，记录失败渲染步骤的工具名及其报错原文（不静默，总结里点名并透出异常）
            let _dashActualSheet = null;   // CreateSheet 返回的真实工作表名
            let slicerOk = true;   // excel_create_slicer 执行失败时置 false，后续图表据此降级回退 dataValues
            for (let di = renderOffset; di < planSteps.length; di++) {
                const st = planSteps[di];
                // ⭐️ 确保 thoughtList[di] 存在且 title 正确（skeleton 和 planSteps 数量可能不一致 → 数组越界乱序）
                if (Array.isArray(turn.thoughtList)) {
                    while (turn.thoughtList.length <= di) {
                        turn.thoughtList.push({ id: `skill_dash_${turn.id}_${turn.thoughtList.length}`, round: di+1, kind: 'skill_step', title: '', status: 'pending', startTs: Date.now() });
                    }
                    if (!turn.thoughtList[di].title || (st.title && turn.thoughtList[di].title !== st.title)) {
                        turn.thoughtList[di].title = st.title || st.detail || `步骤 ${di+1}`;
                    }
                }
                yield { type: 'progress', stepIndex: di, status: 'running', detail: st.detail || st.title };
                if (st.tool === '__html__') {
                    let _htmlOk = true;
                    try {
                        const html = dashEngine.renderHtml(prepared, plan);
                        const r = await toolRegistry.invokeTracked(turn, 'excel_display_html_in_browser', {
                            htmlContent: html, title: (plan.title || '数据看板') + '大屏'
                        }, {
                            round: 1, displayName: '生成 HTML 大屏',
                            displayArgs: { '页面标题': (plan.title || '数据看板') + '大屏' },
                            ctx: { turn, sessionStore, signal }
                        });
                        if (signal?.aborted) throw new Error('user_aborted');
                        _dashRecs.push({ tool: 'excel_display_html_in_browser', showArgs: { 页面: (plan.title || '数据看板') + '大屏' }, raw: r });
                        if (r?.ok === false || r?.error) { _htmlOk = false; _dashHtmlErr = String(r?.error || r?.result || 'HTML 大屏打开失败'); }
                    } catch (e) {
                        if (e?.message === 'user_aborted') throw e;
                        _htmlOk = false;
                        _dashHtmlErr = e?.message || 'HTML 大屏打开失败';
                        ChatLog.warn?.('PIPELINE', `[skill-dashboard] HTML 大屏失败（不影响 Excel 内看板）:`, _dashHtmlErr);
                    }
                    yield { type: 'progress', stepIndex: di, status: _htmlOk ? 'completed' : 'failed', detail: _htmlOk ? 'HTML 大屏已生成' : _dashHtmlErr };
                    continue;
                }
                // ── create_slicer 失败降级：本步骤是图表且此前切片器已失败 → 改回 dataValues 渲染
                //    （useSlicer 时图表 dataRange 指向"透视缓存"透视表区域，切片器/透表未建成则该区域为空 → 图表空白；
                //     回退用蓝图预存的聚合数据数组直绑，保底出图，代价是失去切片器联动。）
                if (st.tool === 'excel_create_chart' && !slicerOk && st._dashFallback && st._dashFallback.dataValues) {
                    st.args = Object.assign({}, st.args, { dataValues: st._dashFallback.dataValues });
                    delete st.args.dataRange;
                    delete st.args.pivotSheet;
                    delete st.args.pivotTable;
                    ChatLog.warn?.('PIPELINE', `[skill-dashboard] 切片器创建失败，图表「${st.title}」已回退为内存数据渲染（无切片器联动）`);
                }
                // 失败工具名收集：不静默——若新增工具在后端未加载（如未重启 Excel），
                // 会在总结中明确提示"工具未就绪，请重启 Excel"，避免顶部整体空白却无解释。
                let _stepOk = true, _stepErr = '';
                try {
                    const r = await toolRegistry.invokeTracked(turn, st.tool, st.args, {
                        displayArgs: st.displayArgs || {},
                        ctx: { turn, sessionStore, signal }
                    });
                    if (signal?.aborted) throw new Error('user_aborted');
                    _dashRecs.push({ tool: st.tool, showArgs: st.displayArgs || {}, raw: r });
                    if (r?.ok === false || r?.error) {
                        _stepOk = false;
                        _stepErr = String(r?.error || r?.result || "").slice(0, 200);
                        if (st.tool === 'excel_create_slicer') {
                            slicerOk = false;
                            // ── 🔑 CreateSlicer 失败降级：动态追加 create_chart 步骤（用内存数据）──
                            //    原降级逻辑只处理 useSlicer=false 时原本就存在的 create_chart 步骤，
                            //    但 useSlicer=true 时 buildExcelPlan 不会生成 create_chart 步骤
                            //    （CreateSlicer 一步到位），导致切片器失败后看板完全没图表。
                            //    现在 CreateSlicer 失败 → 动态向 planSteps 尾部追加 create_chart 步骤，
                            //    复用 prepared.charts 里 blueprint 引擎预聚合好的 data/header，
                            //    用 dataValues 直接绑内存数据出图，代价是失去切片器联动。
                            try {
                                const _charts = (prepared && Array.isArray(prepared.charts)) ? prepared.charts : [];
                                const _xl = planSteps.find(s => s.tool === 'excel_create_sheet' && s.args && s.args.name && s.args.name !== '_dash_pivot');
                                const _sheet = _xl ? _xl.args.name : (plan.sheet || '看板');
                                const _layout = { dashboardCols: 22, chartStartRow: 8, centerStartCol: 1, chartSlotRows: 12, chartRowGap: 12, chartColGap: 0, chartSlotCols: 6 };
                                const _slotCols = 8, _slotRows = 11, _rowGap = 14, _gridBase = 5;
                                const _colLett = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
                                for (let ci = 0; ci < _charts.length; ci++) {
                                    const cd = _charts[ci];
                                    if (!cd || !cd.data || !cd.data.length) continue;
                                    const slot = Math.floor(ci / 2);
                                    const colOff = (ci % 2) * _slotCols;
                                    const row = _layout.chartStartRow + slot * _rowGap;
                                    const colA = _colLett(_gridBase + colOff);
                                    const colB = (ci % 2) === 1 ? _colLett(_layout.dashboardCols) : _colLett(_gridBase + _slotCols + colOff - 1);
                                    const anchor = _sheet + '!' + colA + row + ':' + colB + (row + _slotRows - 1);
                                    const isPie = /xlPie|xlDoughnut|xlFunnel|xlWaterfall|xlTreemap|xlSunburst/.test(String(cd.chartType || ''));
                                    const _memData = JSON.stringify([cd.header, ...cd.data]);
                                    const fallbackStep = {
                                        title: `生成图表「${cd.title || '图表 ' + (ci + 1)}」（降级）`,
                                        detail: '内存数据渲染（切片器失败回退）',
                                        tool: 'excel_create_chart',
                                        args: {
                                            chartType: cd.chartType || 'xlColumnClustered',
                                            title: cd.title || '图表 ' + (ci + 1),
                                            sheet: _sheet,
                                            position: JSON.stringify({ anchorRange: anchor }),
                                            seriesBy: 'columns',
                                            chartStyle: 201,
                                            hasLegend: !isPie,
                                            showDataLabels: isPie,
                                            dataLabelPosition: isPie ? 'xlLabelPositionBestFit' : (isPie ? '' : /xlLine/.test(String(cd.chartType || '')) ? 'xlLabelPositionAbove' : 'xlLabelPositionOutsideEnd'),
                                            dataValues: _memData,
                                            // 【看板多图必须显式声明】降级路径同样是多张图，
                                            // 不传 multiChart 会让后建的图删掉先建的，最终只剩一张。
                                            multiChart: true
                                        },
                                        displayName: `[降级] 图表「${cd.title || ci + 1}」`,
                                        displayArgs: { '位置': anchor }
                                    };
                                    // 插入到后续步骤之前（push 到 planSteps 尾部，循环自然会执行）
                                    planSteps.push(fallbackStep);
                                    ChatLog.warn?.('PIPELINE', `[skill-dashboard] 追加降级图表步骤「${cd.title || ci + 1}」→ dataValues ${cd.data.length} 行`);
                                }
                            } catch (_fallbackErr) {
                                ChatLog.warn?.('PIPELINE', `[skill-dashboard] 降级步骤生成失败: ${_fallbackErr?.message}`);
                            }
                        }
                        _dashFailedTools.set(st.tool, String(r?.error || r?.result || '').slice(0, 200));
                        ChatLog.warn?.('PIPELINE', `[skill-dashboard] 步骤「${st.title}」失败（继续后续步骤）:`, String(r?.error || r?.result || ''));
                    }
                } catch (e) {
                    if (e?.message === 'user_aborted') throw e;
                    _stepOk = false;
                    _stepErr = String(e?.message || '执行失败').slice(0, 200);
                    if (st.tool === 'excel_create_slicer') slicerOk = false;
                    _dashFailedTools.set(st.tool, String(e?.message || '执行失败').slice(0, 200));
                    ChatLog.warn?.('PIPELINE', `[skill-dashboard] 步骤「${st.title}」异常（继续后续步骤）:`, e?.message);
                }
                yield { type: 'progress', stepIndex: di, status: _stepOk ? 'completed' : 'failed', detail: _stepOk ? (st.detail || st.title || '完成') : (_stepErr || '执行失败') };
            }
            // 总结里明确暴露未就绪工具，避免"顶部空白却无解释"
            let _dashMsg = dashEngine.summarize(prepared, plan, { htmlError: _dashHtmlErr });
            if (_dashFailedTools && _dashFailedTools.size > 0) {
                {
                    const errText = Array.from(_dashFailedTools.entries()).map(([t, e]) => t + '→' + (e || '无错误信息')).join('；');
                    _dashMsg += ' ⚠️ ' + _dashFailedTools.size + ' 个步骤失败（' + errText.slice(0, 600) + '）';
                    console.warn('[skill-dashboard] 失败步骤明细:', errText);
                }
            }
            execResult = {
                ok: true,
                result: { message: _dashMsg },
                stepResults: _dashRecs
            };
            planSummary = `「${plan.title || '数据看板'}」· ${_dashRecs.length} 步施工完成`;
            const _finalAnoms = dashEngine.collectAnomalies?.(prepared, execResult) || [];
            if (_finalAnoms.length > 0) {
                ChatLog.info?.('PIPELINE', `[skill-dashboard] 执行后仍有 ${_finalAnoms.length} 处 anomalies（已尽力修正）`);
            }
        } else {
            // ── 美化/格式整理：统一框架"每步=真实工具调用"；工具表达不了的选项整体回退原 VBA ──
            // 工具路径也要"自动知道表头在哪"：内容识别作为兜底/补足，但不再无条件覆盖 LLM 的 headerRows。
            // 取并集(更大者)：LLM 判了多行表头则采纳(它懂"合并标题+列头"语义)；内容识别只负责补 LLM 漏判的行，
            // 绝不把表头块算小(算小=漏掉真列头行，实测标题美化但列头白底黑字)。
            if (typeof SP.specBuild.detectBeautifyHeader === 'function' && plan && typeof plan === 'object') {
                try {
                    // 用户明确"移除/去内容底色" → 执行层强制 noContentBg，不再被 LLM 规划的浅底色覆盖。
                    // 实测 LLM 规划加了内容区底色，直接无视了用户"移除内容区底色"这一命令。
                    const _msg = String(turn?.originalUserMessage || turn?._originalUserMessage || turn?.userText || turn?.userMessage || turn?.message || '');
                    if (/(移除|去掉|去除|清除|取消|删除|不要|无需|不用)/.test(_msg) && /(内容)?(区)?(底色|填充色|背景|底纹|填充)/.test(_msg)) {
                        plan = Object.assign({}, plan, { noContentBg: true, bodyColor: null, bandedRows: false });
                    }
                    const llmRows = Number(plan.headerRows) >= 1 ? Number(plan.headerRows) : 0;
                    const _hdrOverlay = await SP.specBuild.detectBeautifyHeader(ctx, turn, toolRegistry, sessionStore, signal);
                    const detectRows = (_hdrOverlay && Number(_hdrOverlay.headerRows) >= 1) ? Number(_hdrOverlay.headerRows) : 0;
                    const usedRows = Math.max(llmRows, detectRows);
                    if (usedRows >= 1) {
                        plan = Object.assign({}, plan, { headerRows: Math.max(1, Math.min(5, usedRows)) });
                    }
                    ChatLog.info?.('PIPELINE', `[skill-spec] 表头协调: plan.headerRows=${llmRows} detectBeautifyHeader=${_hdrOverlay ? JSON.stringify(_hdrOverlay) : 'null'} detectRows=${detectRows} usedRows=${usedRows} 移除底色=${plan.noContentBg === true}`);
                } catch (e) {
                    if (e && e.message === 'user_aborted') throw e;
                    ChatLog.warn?.('PIPELINE', '[skill-spec] 表头内容识别失败，沿用规划 headerRows:', e && e.message ? e.message : e);
                }
            }
            let beautifyOpsDone = false;
            const beautifyOps = SP.specBuild.buildBeautifyOps(plan, ctx);
            const _beautifyToolsOk = beautifyOps && beautifyOps.length > 0 &&
                beautifyOps.every(o => o.tool === 'excel_execute_vba' || (toolRegistry && toolRegistry.get(o.tool)));
            if (_beautifyToolsOk) {
                planSummary = beautifyOps.map(o => o.label).join('、');
                yield { type: 'status', text: `已按「${planSummary}」规划，逐项应用工具...` };
                // 整次美化合并为单一撤销项：美化前先捕获整表快照登记一次，
                // 并临时抑制后续每个 format_range 子调用的独立登记（循环结束 finally 再关闭）
                try {
                    if (App && App.excelBridge && App.excelBridge.CaptureAndRegisterUndo) {
                        const reg = App.excelBridge.CaptureAndRegisterUndo('智能表格美化');
                        if (reg !== 'ok' && reg !== 'already') ChatLog.warn?.('PIPELINE', '[skill-spec] 撤销快照登记:', reg);
                    }
                    if (App && App.excelBridge && App.excelBridge.SetUndoSuppressed) {
                        App.excelBridge.SetUndoSuppressed(true);
                    }
                } catch (e) {
                    ChatLog.warn?.('PIPELINE', '[skill-spec] 撤销快照登记失败:', e && e.message ? e.message : e);
                }
                let _beautifyFailed = null;
                const _applied = [];
                const _stepRecs = [];
                try {
                for (let bi = 0; bi < beautifyOps.length; bi++) {
                    const op = beautifyOps[bi];
                    if (earlySkeleton && earlySkeleton[bi]) {
                        yield { type: 'progress', stepIndex: bi, status: 'running', detail: '正在执行…' };
                    }
                    const r = await toolRegistry.invokeTracked(turn, op.tool, op.args, {
                        displayName: op.label,
                        displayArgs: op.showArgs || {},
                        ctx: { turn, sessionStore, signal }
                    });
                    if (signal?.aborted) throw new Error('user_aborted');
                    const okB = !(r?.ok === false || r?.error);
                    const outcomeB = okB ? SP.util.summarizeToolResult(r) : (r?.error || '执行失败');
                    if (earlySkeleton && earlySkeleton[bi]) {
                        yield { type: 'progress', stepIndex: bi, status: okB ? 'completed' : 'failed', detail: `${op.label}：${outcomeB}` };
                    }
                    if (okB) _applied.push(op.label);
                    if (!okB) _beautifyFailed = outcomeB;
                    _stepRecs.push({ tool: op.tool, args: op.args, showArgs: op.showArgs || {}, label: op.label, ok: okB, outcome: outcomeB, raw: r });
                }
                } finally {
                    try { if (App && App.excelBridge && App.excelBridge.SetUndoSuppressed) App.excelBridge.SetUndoSuppressed(false); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec-execute] 恢复 SetUndoSuppressed 失败:', e); }
                }
                execResult = {
                    ok: !_beautifyFailed,
                    result: {
                        success: !_beautifyFailed,
                        message: _beautifyFailed ? ('部分步骤未应用：' + _beautifyFailed) : ('已应用：' + _applied.join('、')),
                        appliedSteps: _applied
                    },
                    stepResults: _stepRecs,
                    subSteps: (earlySkeleton || []).map(s => ({ ...s })),
                    toolCallId: null
                };
                beautifyOpsDone = true;
            }
            // ──（兜底）美化/格式整理 VBA 渲染：仅当工具框架不可用时走 ──
            if (!beautifyOpsDone) {
            // plan 已在上方通过 normalizeSpecFn 规范化，
            // 直接调用 renderBeautifyVBA/renderFormatVBA，避免 planAndRender 内部二次规范化
            // renderedCode 已在函数级声明（作用域提升），此处直接赋值
            if (isFormatter && typeof specEngine.renderFormatVBA === 'function') {
                renderedCode = specEngine.renderFormatVBA(plan);
            } else if (typeof specEngine.renderBeautifyVBA === 'function') {
                renderedCode = specEngine.renderBeautifyVBA(plan, ctx);
            } else if (typeof specEngine.planAndRender === 'function') {
                // 兜底：使用 planAndRender（会二次规范化，但兼容性好）
                const rendered = specEngine.planAndRender(plan, ctx);
                if (!rendered.ok || !rendered.code) {
                    throw new Error('Spec 渲染失败: ' + (rendered.error || 'unknown'));
                }
                renderedCode = rendered.code;
            } else {
                throw new Error('VBA 渲染器未就绪');
            }
            if (!renderedCode) {
                throw new Error('VBA 代码生成失败');
            }
            // planSummary 按 Skill 类型分别生成：
            //   - Beautifier：主题 + 字体
            //   - Formatter：字体 + 对齐 + 边框
            if (isFormatter) {
                planSummary = `${plan.fontName || '微软雅黑'} ${plan.fontSize || 11}号` +
                    (plan.alignment ? ` · ${plan.alignment === 'center' ? '居中' : plan.alignment === 'left' ? '左对齐' : '右对齐'}` : '') +
                    (plan.addBorders ? ' · 标准边框' : '') +
                    (plan.freezeHeader ? ' · 冻结表头' : '');
            } else {
                planSummary = `主题 ${specEngine.listThemes().find(t => t.id === plan.theme)?.name || plan.theme}` +
                    (plan.fontName ? ` · 字体 ${plan.fontName} ${plan.fontSize}号` : '');
            }
            yield { type: 'status', text: `已按「${planSummary}」渲染样式，正在应用...` };

            // ✨ 渲染 VBA 完成 → 前面 N-1 步标记 completed，最后 1 步 running（整体执行中）
            // 这样用户看到：先 6 步全 pending → 第 1 步 running → 前 5 步 completed + 最后 1 步 running
            // → 执行完最后 1 步也 completed（由外层 specResult 整体替换完成）
            if (earlySkeleton && earlySkeleton.length > 0) {
                const n = earlySkeleton.length;
                // 先把前 n-1 步标记为 completed（渲染已准备好所有步骤的配置）
                for (let i = 0; i < n - 1; i++) {
                    yield { type: 'progress', stepIndex: i, status: 'completed' };
                }
                // 最后 1 步 running（正在执行）
                yield { type: 'progress', stepIndex: n - 1, status: 'running', detail: '正在应用到工作表...' };
            }

            // 4. 执行渲染后的 VBA（优先 excelBridge.ExecuteVBA，回退 MCP 工具）
            // 直连 COM 不经过工具守卫：破坏性 VBA（删除/清空类操作）必须先经用户确认
            let vbaCancelled = false;
            if (toolRegistry.isDestructiveVba && toolRegistry.isDestructiveVba(renderedCode)) {
                const turnDec = toolRegistry.getTurnDangerousDecision ? toolRegistry.getTurnDangerousDecision(turn) : null;
                if (turnDec === 'deny') {
                    vbaCancelled = true;
                    execResult = { ok: false, error: '用户已取消本轮危险操作，技能 VBA 含删除/清空操作被自动拦截' };
                    ChatLog.info?.('PIPELINE', '[skill-spec] 本轮用户已取消危险操作，破坏性 VBA 自动拒绝');
                } else if (turnDec !== 'allow') {
                    const reason = '【安全守卫】该技能生成的 VBA 包含删除/清空数据的操作，执行后可能不可逆。'
                        + '\n\n技能: ' + (displayName || 'Spec 规划')
                        + '\n代码预览: ' + String(renderedCode).slice(0, 160).replace(/\s+/g, ' ') + '…';
                    let confirmedSpec = false;
                    try { confirmedSpec = await toolRegistry.confirmDangerous(displayName || 'excel_execute_vba', reason); } catch (_e) { confirmedSpec = false; }
                    if (!confirmedSpec) {
                        vbaCancelled = true;
                        if (toolRegistry.setTurnDangerousDecision) toolRegistry.setTurnDangerousDecision(turn, 'deny');
                        execResult = { ok: false, error: '用户取消执行危险操作（技能 VBA 含删除/清空操作）' };
                        ChatLog.info?.('PIPELINE', '[skill-spec] 用户取消破坏性 VBA 执行');
                    }
                }
            }
            const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
            if (!vbaCancelled && bridge && typeof bridge.ExecuteVBA === 'function') {
                try {
                    const resultJson = await bridge.ExecuteVBA(renderedCode, displayName);
                    const parsed = typeof resultJson === 'string' ? JSON.parse(resultJson) : resultJson;
                    const ok = parsed?.success === true || parsed?.ok === true || (parsed && !parsed.error);
                    execResult = ok
                        ? { ok: true, result: parsed, toolCallId: null }
                        : { ok: false, error: parsed?.error || parsed?.message || 'VBA 执行失败', toolCallId: null };
                    if (execResult.ok) {
                        ChatLog.info?.('PIPELINE', `[skill-spec] 通过 excelBridge.ExecuteVBA 执行成功 (${displayName})`);
                    } else {
                        ChatLog.warn?.('PIPELINE', '[skill-spec] excelBridge.ExecuteVBA 返回失败，尝试 MCP 工具回退:', execResult.error);
                    }
                } catch (bridgeErr) {
                    ChatLog.warn?.('PIPELINE', '[skill-spec] excelBridge.ExecuteVBA 异常，尝试 MCP 工具回退:', bridgeErr?.message);
                }
            }
            if (!vbaCancelled && (!execResult || !execResult.ok)) {
                const origMode = turn.mode;
                turn.mode = 'expert';
                execResult = await toolRegistry.invokeTracked(turn, 'excel_execute_vba', {
                    code: renderedCode
                }, {
                    round: 1,
                    displayName: displayName + '（Spec 规划）',
                    displayArgs: {
                        '规划模式': plannedByLLM ? 'LLM 结构化 Spec' : '关键词兜底',
                        '主题': plan.theme,
                        '表头': plan.headerStyle,
                        '字体': plan.fontName + ' ' + plan.fontSize + '号',
                        '对齐': plan.alignment,
                        '条件格式': plan.conditionalFormat || '无',
                        '冻结表头': plan.freezeHeader ? '是' : '否',
                        '斑马纹': plan.bandedRows ? '是' : '否'
                    },
                    ctx: { turn, sessionStore, signal }
                });
                turn.mode = origMode;
            }
            if (signal?.aborted) throw new Error('user_aborted');
            if (vbaCancelled) {
                // 同 CRUD 分支：守卫取消不 throw，走 finalize 输出取消说明（不降级 agent-loop）
                specCancelled = true;
                cancelMessage = execResult?.error || '用户取消执行危险操作';
            } else if (!execResult?.ok) {
                throw new Error(execResult?.error || 'VBA 执行失败');
            }
            } // ── 兜底 VBA 结束（仅工具框架不可用时进入）──
        }
        // 5. 输出总结（按 Skill 类型生成，包含实际执行结果）
        const _r = (execResult && execResult.result != null) ? execResult.result : execResult || {};
        let _er = (_r && typeof _r === "object") ? _r : {};
        if (typeof _er.message === "string" && _er.message.trim().startsWith("{")) {
            try { _er = { ..._er, ...JSON.parse(_er.message) }; } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-spec-execute] 输出 message 非 JSON，保留原文:', e); }
        }
        const _modeTag = plannedByLLM ? "LLM 智能规划" : "智能识别";
        let outputText;
        if (specCancelled) {
            // 取消态摘要：不用"✅ 完成"，明确告知未执行任何操作及原因
            outputText = `**⛔ ${displayName}已取消**\n\n${cancelMessage || '未执行任何操作。'}`;
        } else if (isCleaner) {
            const parts = [`**✅ ${displayName}完成**（${_modeTag}）`];
            // 真实重复数 = 总删除行数 − 空行数，空行独立统计，避免把"空行"错标成"重复数据"。
            const _dupTot = typeof _er.removedRows === "number" ? _er.removedRows : (typeof _er.deletedRows === "number" ? _er.deletedRows : 0);
            const _empN = typeof _er.removedEmptyRows === "number" ? _er.removedEmptyRows : 0;
            const _netDup = Math.max(0, _dupTot - _empN);
            if (_netDup > 0) parts.push(`已删除 ${_netDup} 行重复数据`);
            if (_empN > 0) parts.push(`已删除 ${_empN} 行空行`);
            else if (typeof _er.markedRows === "number" && _er.markedRows > 0) parts.push(`已标记 ${_er.markedRows} 行重复数据（仅高亮）`);
            if (typeof _er.cleanedRows === "number") parts.push(`共处理 ${_er.cleanedRows} 行`);
            outputText = parts.join(" · ");
        } else if (isCrud) {
            outputText = `**✅ ${displayName}完成**（${_modeTag}）`;
            if (typeof _er.affectedRows === "number") outputText += ` · 影响 ${_er.affectedRows} 行`;
            else if (typeof _er.deletedRows === "number") outputText += ` · 删除 ${_er.deletedRows} 行`;
        } else if (isChart) {
            outputText = `**✅ ${displayName}完成**（${_modeTag}）`;
            if (plan.chartType) outputText += ` · ${plan.chartType}图`;
            if (plan.title) outputText += `「${plan.title}」`;
        } else if (isDashboard) {
            outputText = `**✅ ${displayName}完成**（${_modeTag}） · 「${plan.title || '数据看板'}」`;
            if (_er && typeof _er.message === 'string' && _er.message.trim()) {
                outputText += '\n\n' + _er.message.trim();
            }
        } else {
            outputText = `**✅ ${displayName}完成**（${_modeTag}）`;
            if (planSummary) outputText += ` · ${planSummary}`;
        }
        // ✨ 收尾：无论清洗/增删改查/美化，统一把供给侧步骤推到 final 状态，
        // 驱动绿面板（ProgressTracker + renderSkillSteps）把"执行中"翻转为"已完成"。
        // 之前 CRUD / 美化分支缺失这一步，导致工作已做完但最后一步一直卡在"执行中"。
        if (Array.isArray(earlySkeleton) && earlySkeleton.length > 0) {
            const _finalStatus = execResult && execResult.ok ? 'completed' : 'failed';
            for (let _si = 0; _si < earlySkeleton.length; _si++) {
                yield { type: 'progress', stepIndex: _si, status: _finalStatus };
            }
        } else {
            yield { type: 'progress', stepIndex: 0, status: execResult && execResult.ok ? 'completed' : 'failed' };
        }
        // 面板步骤：规划骨架 + 真实执行步骤 + 执行结果合并。
        let _stepsArr = [];
        const _kind = specContract.skeletonKind;
        const finalizeSkeleton = SP.specBuild.buildSpecSteps(plan, _kind, isDashboard && _dashPrecomputed ? { prepared: _dashPrecomputed.prepared } : null);
        const _execTarget = execResult && (execResult.result != null ? execResult.result : execResult);
        let realSteps = [];
        if (execResult && execResult.ok) {
            realSteps = SP.util.findStepsDeep(_execTarget);
        }
        // ✨ 从 execResult.result 顶层提取摘要字段
        // 特别注意：后端 MCP 经常把整个结果 JSON.stringify 后塞进 message 字段，
        // 直接 push 会导致 detail 里全是原始 JSON。需要先尝试解析 message。
        const topSummary = [];
        let _execObj = _execTarget;
        if (_execTarget && typeof _execTarget === 'string') {
            try { _execObj = JSON.parse(_execTarget); } catch { _execObj = null; }
        }
        if (_execObj && typeof _execObj === 'object') {
            // 解析可能是 JSON 字符串的 message 字段
            let msg = _execObj.message;
            if (typeof msg === 'string' && msg.trim()) {
                const trimmed = msg.trim();
                if ((trimmed.startsWith('{') || trimmed.startsWith('[')) && trimmed.length < 4000) {
                    // 看起来像 JSON → 尝试解析
                    try {
                        const parsed = JSON.parse(trimmed);
                        // 从解析后的 JSON 里抽摘要
                        if (typeof parsed.message === 'string' && parsed.message !== trimmed) {
                            topSummary.push(parsed.message);
                        }
                        if (typeof parsed.deletedRows === 'number') topSummary.push(`删除 ${parsed.deletedRows} 行`);
                        if (typeof parsed.removedRows === 'number') topSummary.push(`删除 ${parsed.removedRows} 行重复数据`);
                        if (typeof parsed.cleanedRows === 'number') topSummary.push(`处理 ${parsed.cleanedRows} 行`);
                        if (typeof parsed.markedRows === 'number') topSummary.push(`标记 ${parsed.markedRows} 行`);
                        if (typeof parsed.success === 'boolean' && parsed.success === true && topSummary.length === 0) {
                            topSummary.push('执行成功');
                        }
                    } catch {
                        topSummary.push(trimmed); // 不是有效 JSON → 当纯文本摘要
                    }
                } else {
                    topSummary.push(trimmed); // 普通文本 message
                }
            }
            // 顶层直接有这些字段（message 不是 JSON 时）
            if (typeof _execObj.deletedRows === 'number') topSummary.push(`删除 ${_execObj.deletedRows} 行`);
            if (typeof _execObj.removedRows === 'number') topSummary.push(`删除 ${_execObj.removedRows} 行重复数据`);
            if (typeof _execObj.cleanedRows === 'number') topSummary.push(`处理 ${_execObj.cleanedRows} 行`);
            if (typeof _execObj.markedRows === 'number') topSummary.push(`标记 ${_execObj.markedRows} 行`);
            if (typeof _execObj.createdAt === 'string' || typeof _execObj.createdAt === 'number') topSummary.push(`新增 ${_execObj.createdAt} 行`);
            if (typeof _execObj.updatedAt === 'number') topSummary.push(`更新 ${_execObj.updatedAt} 处`);
        }
        // 去重 + 合并 realSteps 和 topSummary，过滤无意义的 generic 摘要
        const mergedDetails = [...realSteps, ...topSummary]
            .filter(t => t && t.trim() && !C.GENERIC_SUMMARY_RE.test(t.trim()))
            .filter((t, i, arr) => arr.indexOf(t) === i); // 去重
        const allDetails = mergedDetails.filter(t => !mergedDetails.some(r => r !== t && (r.includes(t) || t.includes(r))));

        // ✨ 包装成嵌套结构：Skill 大步骤 + subSteps（和骨架 yield 处对齐）
        if (isCleaner || isCrud) {
            // 优先用执行循环里更新好的 subSteps（有真实 detail/summary）
            const updatedSubs = execResult?.subSteps || [];
            const subSteps = updatedSubs.length > 0 ? updatedSubs : finalizeSkeleton.map(s => ({
                title: s.title,
                description: s.description || s.title,
                detail: s.detail || '',
                status: 'completed',
                strategy: 'direct'
            }));
            _stepsArr = [{
                title: displayName || 'Skill 执行',
                description: displayName || 'Skill 执行',
                strategy: 'pipeline',
                status: 'completed',
                subSteps
            }];
        } else if (finalizeSkeleton.length > 0) {
            // 末步展开只应显示自己的内容：若该步已有自身 detail（buildSpecSteps 已给），
            // 就不再合并 findStepsDeep 的整体摘要，否则会把"已应用：A、B、C…"累计列表拼进末步。
            const last = finalizeSkeleton[finalizeSkeleton.length - 1];
            const needStepDetail = !(last.detail && String(last.detail).trim());
            if (needStepDetail) {
                const stepOnly = realSteps
                    .filter(t => t && t.trim())
                    // 排除整体累计摘要（美化/格式的分步 message）
                    .filter(t => !/^已应用\s*[:：]|^部分步骤未应用/.test(t.trim()))
                    .filter((t, i, arr) => arr.indexOf(t) === i);
                if (stepOnly.length > 0) {
                    const merged = stepOnly.join('；');
                    last.detail = last.detail ? (last.detail + ' · ' + merged) : merged;
                }
            }
            const _stepRecs = (execResult && Array.isArray(execResult.stepResults)) ? execResult.stepResults : [];
            _stepsArr = _stepRecs.length > 0
                ? finalizeSkeleton.map((s, i) => {
                    const rec = _stepRecs[i];
                    return {
                        ...s,
                        status: 'completed',
                        strategy: 'direct',
                        toolName: rec ? rec.tool : null,
                        params: rec ? rec.showArgs : null,
                        result: rec ? rec.raw : null
                    };
                })
                : finalizeSkeleton.map(s => ({ ...s, status: s.status || 'completed' }));
            // ✨【透明回写 · chart】顶层骨架步骤的 detail 是 buildSpecSteps 生成的静态快照，
            // 不包含 LLM 解析出的聚合维度/指标（它们在 execute 内运行时才确定）。这里把
            // "按XX汇总XX"的聚合信息回写进"生成图表/图例"骨架步，让展开面板能看到真实口径，
            // 而非只显示"xlPie 图 / 显示"。
            // 聚合信息从 plan._resolvedAgg 读取（execute 聚合块已持久化到此），
            // 不能引用闭包变量 aggregate*（声明在 chart 内层块，finalize 在更外层块 → ReferenceError）。
            const resolvedAgg = (plan && typeof plan === 'object' && plan._resolvedAgg && plan._resolvedAgg.ready) ? plan._resolvedAgg : null;
            if (_kind === 'chart' && resolvedAgg && resolvedAgg.dim) {
                const _aggFunc = resolvedAgg.func || 'sum';
                const aggLabel = _aggFunc === 'count' ? '计数' :
                                 _aggFunc === 'avg' || _aggFunc === 'average' ? '求平均' :
                                 _aggFunc === 'max' ? '取最大' :
                                 _aggFunc === 'min' ? '取最小' : '求和';
                const aggText = `按「${resolvedAgg.dim}」${aggLabel}「${resolvedAgg.val || '数值'}」`;
                for (const step of _stepsArr) {
                    const key = String(step.description || step.title || '');
                    if (key === '生成图表') {
                        // 聚合成功：图表实际绘制在聚合结果上，detail 覆盖为聚合口径 + 派生范围
                        step.detail = resolvedAgg.dataRange
                            ? `${step.detail || ''} · ${aggText}（绘制于 ${resolvedAgg.dataRange}）`
                            : `${step.detail || ''} · ${aggText}`;
                    } else if (key === '图例' && step.detail === '显示') {
                        step.detail = '显示（按类别着色）';
                    } else if (key === '读取数据' && (!step.detail || step.detail.includes('当前'))) {
                        // 读取数据步骤展示"真正读取的数据源"，而非把"按XX聚合XX"挂在读取步上——
                        // 原实现 append aggText 导致"读取数据"却显示聚合口径，语义错位、用户看不懂。
                        // 聚合信息只留在"生成图表"步；这里用快照读取的真实范围+行列数替代占位符。
                        const psrc = chartSnapMeta || null;
                        const srcAddr = psrc && psrc.range ? String(psrc.range).replace(/\$/g, '') : '';
                        const srcSheet = psrc && psrc.worksheet ? String(psrc.worksheet).replace(/[\[\]']/g, '') : '';
                        const srcDims = (psrc && psrc.rows != null && psrc.columns != null)
                            ? `（${psrc.rows} 行 × ${psrc.columns} 列）` : '';
                        // 兜底也优先真实地址，绝不用"当前选定区域"占位：选区地址 > displayRange(非占位) > dataRange，
                        // 仅当连选区地址都没有时才对占位文本。真实地址均来自快照/聚合回填/选区映射。
                        const _selAddr = ctx.selection && ctx.selection.address ? String(ctx.selection.address).replace(/\$/g, '') : '';
                        let fallbackSrc = (ctx.selection && ctx.selection.sheetName && _selAddr)
                            ? `${String(ctx.selection.sheetName).replace(/[\[\]']/g, '')}!${_selAddr}` : '';
                        fallbackSrc = fallbackSrc || (displayRange && !String(displayRange).includes('当前') ? displayRange : '');
                        fallbackSrc = fallbackSrc || (dataRange || (_selAddr || ''));
                        if (!fallbackSrc) fallbackSrc = '当前选定区域';
                        step.detail = srcAddr
                            ? `读取 ${srcSheet ? srcSheet + '!' : ''}${srcAddr}${srcDims}`
                            : `读取 ${fallbackSrc}${srcDims || ''}`;
                    }
                }
            }
        } else if (allDetails.length > 0) {
            _stepsArr = allDetails.map((d, i) => ({ title: `执行 ${i + 1}`, detail: d, status: 'completed' }));
        }

        // ✨ 未框选触发自动探测时，把"自动探测数据区域"作为首个步骤展示（真实工具调用 excel_read_range）。
// 仅美化/格式 prepend；清洗的"读取数据"子步骤已兼职探测（读到探测地址即整表数据区），
// 不再单独多加一步，避免同一 excel_read_range 出现两次。
if (probedInfo && probedInfo.address && (isBeautifier || isFormatter)) {
    const probeSheet = probedInfo.sheetName || '';
    _stepsArr.unshift({
        title: '自动探测数据区域',
        description: '自动探测数据区域',
        detail: `探测到 ${probedInfo.address}（${probedInfo.rowCount || '?'}行 × ${probedInfo.colCount || '?'}列${probeSheet ? ' · ' + probeSheet : ''}）`,
        status: 'completed',
        strategy: 'direct',
        toolName: 'excel_read_range',
        params: { 方式: '未框选→自动探测数据区', 工作表: probeSheet || '活动表' },
        result: null
    });
}

        // DEBUG: trace spec steps construction
        ChatLog.info?.('PIPELINE', `[skill-spec] _kind=${_kind} finalizeSkeleton=${finalizeSkeleton.length} realSteps=${realSteps.length} → _stepsArr=${_stepsArr.length} (hasSub=${!!_stepsArr[0]?.subSteps})`);

        // ✨【图表预览 · skill-chart】聚合/内存绑定时，把聚合结果渲染成前端内嵌 ECharts 预览，
        // 让气泡里直接看到图表（而非仅"✅ 智能图表向导完成"文字）。Excel 里仍生成真实 Chart 对象。
        // 数据源用 plan._resolvedAgg.dataValues（execute 聚合块已持久化，finalize 作用域安全），
        // 避免引用 chart 内层闭包变量 aggregateSnapshot 触发 ReferenceError。
        if (_kind === 'chart') {
            // 自包含解析：不能依赖 L1658 的 resolvedAgg（它在 finalizeSkeleton 非空块内声明，
            // 当 finalizeSkeleton.length===0 走 else 分支时该 const 不执行 → 访问即 ReferenceError）。
            const _aggInfo = (plan && typeof plan === 'object' && plan._resolvedAgg && plan._resolvedAgg.ready) ? plan._resolvedAgg : null;
            if (_aggInfo && Array.isArray(_aggInfo.dataValues)) {
                try {
                    const _dv = _aggInfo.dataValues;
                    const _hasBody = Array.isArray(_dv) && _dv.length >= 2 && Array.isArray(_dv[0]);
                    if (_hasBody) {
                        const _labels = [], _values = [];
                        for (let di = 1; di < _dv.length; di++) {
                            const row = _dv[di];
                            if (Array.isArray(row) && row[0] != null && row[0] !== '') {
                                _labels.push(String(row[0]));
                                _values.push(Number(row[1]) || 0);
                            }
                        }
                        if (_labels.length > 0) {
                            const _aggData = {
                                chartType: plan && plan.chartType ? plan.chartType : 'xlPie',
                                title: (plan && (plan.title || plan.chartTitle)) || `按${_aggInfo.dim}${_aggInfo.val || ''}`,
                                dimColumn: _aggInfo.dim || '维度',
                                valueColumn: _aggInfo.val || '数值',
                                labels: _labels,
                                values: _values
                            };
                            const previewHtml = App.ChatV2.renderHtml && App.ChatV2.renderHtml.buildInlineChartHtml
                                ? App.ChatV2.renderHtml.buildInlineChartHtml(_aggData) : '';
                            if (previewHtml && String(previewHtml).includes('data-echarts')) {
                                outputText = String(outputText || '') + '\n\n' + previewHtml;
                                ChatLog.info?.('PIPELINE', `[skill-chart] 已生成图表内嵌预览：${_labels.length} 个分类 · ${plan && plan.chartType ? plan.chartType : 'xlPie'}`);
                            }
                        }
                    }
                } catch (e) {
                    ChatLog.warn?.('PIPELINE', '[skill-chart] 图表内嵌预览生成失败（不阻断）:', e && e.message ? e.message : e);
                }
            }
        }
        yield { type: 'result', result: { outputText, cancelled: specCancelled, spec: plan, steps: _stepsArr, code: typeof renderedCode !== 'undefined' ? renderedCode : null, aggregate: aggregateSnapshot, dataRange: dataRange || null } };
    }
    MODS.specExecute = {
        executeSkillSpecMode
    };
})();
