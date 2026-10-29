/**
 * ============================================================
 * chat-v2/services/MultiStepEngine.js — 多步骤任务编排（已收敛）
 * ============================================================
 * v20260925 收敛：通用管道已统一到 agent-loop 循环驱动
 *   （见 text-pipeline.runMultiStepPipeline），本模块不再参与路由：
 *   - planSkeleton()：供调用方一次产出 3–8 步骨架（薄封装 TaskPlanner.plan）
 *   - run()：@deprecated，保留一版供回滚；已移除全部"替 LLM 做决定"的硬编码规则
 *     （操作类型去重 / 汇总+产出早停 / 拒绝重复操作 / 读取保护 / 依赖失败跳过 / 任务调度），
 *     仅保留最小护栏：步骤上限、abort、组件未就绪时的降级通道。
 *   旧设计（供理解历史）：渐进增强、顺序执行 MVP、集成 StepExecutor/TaskScheduler/ResultAggregator。
 * 依赖：App.ChatV2.{toolRegistry, sessionStore, agentLoop, StepExecutor, ResultAggregator}
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 最小护栏：单次规划的步骤上限（不再限制"某类操作次数"——是否继续由 LLM 决定）
    const MAX_STEPS = 10;

    /**
     * 解析具备 plan/replan 能力的规划器。
     * 注意：App.ChatV2.taskPlanner（小写 t）只有 detectBuiltinIntent 等拆分工具，没有 plan；
     * 真正具备规划能力的是 App.ChatV2.TaskPlanner（大写 T）。
     */
    function resolvePlanner(taskPlanner) {
        return (taskPlanner && typeof taskPlanner.plan === 'function')
            ? taskPlanner
            : App.ChatV2?.TaskPlanner;
    }

    /**
     * 产出步骤骨架（原始 step 数组，由调用方决定如何转成面板骨架）。
     * @param {Object} opts - { turn, tools, messages, signal, maxSteps, taskPlanner }
     * @returns {Promise<Array<Object>>} 步骤数组；planner 不可用 / 规划失败 / 无需多步时返回 []
     */
    async function planSkeleton(opts) {
        const { turn, tools, messages, signal, maxSteps, taskPlanner } = opts || {};
        const planner = resolvePlanner(taskPlanner);
        if (!planner || typeof planner.plan !== 'function') {
            ChatLog.warn?.('MULTI-STEP', '[planSkeleton] TaskPlanner.plan 不可用，返回空骨架');
            return [];
        }
        const userMsg = (turn && (turn._originalUserMessage || turn.userMessage || turn.message)) || '';
        if (!userMsg) return [];
        try {
            const steps = await planner.plan({
                userMessage: userMsg,
                tools,
                messages,
                signal,
                maxSteps: maxSteps || MAX_STEPS,
                intent: turn?.context?._intent,
                context: turn?.context
            });
            return Array.isArray(steps) ? steps.slice(0, maxSteps || MAX_STEPS) : [];
        } catch (e) {
            ChatLog.warn?.('MULTI-STEP', '[planSkeleton] 规划失败:', e?.message);
            return [];
        }
    }

    /**
     * 运行多步骤管道。
     * @deprecated v20260925 起通用管道统一走 agent-loop（text-pipeline.runMultiStepPipeline），
     *   本函数不再被路由，保留一版供回滚。
     * @param {Object} opts
     * @param {Object} opts.turn - turn 对象
     * @param {AbortSignal} opts.signal
     * @param {Object} opts.sessionStore
     * @param {Object} opts.toolRegistry
     * @param {Object} opts.agentLoop
     * @param {Array} opts.tools - 可用工具列表
     * @param {Array} opts.messages - 已构建的 messages
     * @param {Object} opts.taskPlanner - TaskPlanner 实例
     * @param {Object} opts.progressTracker - ProgressTracker 实例
     * @param {Object} [opts.stepExecutor] - StepExecutor 实例（可选，默认使用 App.ChatV2.StepExecutor）
     * @param {Object} [opts.resultAggregator] - ResultAggregator 实例（可选，默认使用 App.ChatV2.ResultAggregator）
     * @returns {AsyncGenerator} typed events
     */
    async function* run(opts) {
        const {
            turn, signal, sessionStore, toolRegistry, agentLoop,
            tools, messages, taskPlanner
        } = opts || {};

        // ProgressTracker 注册在 App.ChatV2.ProgressTracker（大写 P），而 text-pipeline 传的是
        // deps.progressTracker（小写键，恒为 undefined）→ 必须兜底取全局实例，
        // 否则下面的 progressTracker.init() 直接抛 "Cannot read properties of undefined"（实测崩溃点）。
        const progressTracker = (opts && opts.progressTracker) || App.ChatV2.ProgressTracker;

        if (!turn) {
            yield { type: 'error', error: { code: 'NO_TURN', message: '缺少 turn 对象', recoverable: false } };
            return;
        }

        // 引用新组件（优先使用传入实例，否则从全局获取）
        const stepExecutor = opts.stepExecutor || App.ChatV2.StepExecutor;
        const resultAggregator = opts.resultAggregator || App.ChatV2.ResultAggregator;

        if (!stepExecutor || !resultAggregator || !progressTracker) {
            ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 新组件未就绪，降级为通用对话');
            yield { type: 'fallback_to_agent', reason: 'deps_missing',
                detail: 'StepExecutor/ResultAggregator 未加载' };
            return;
        }

        // 【修复】taskPlanner 能力解析（唯一收敛点）：调用方可能传入无 plan 的小写 taskPlanner，
        // 真正具备规划能力的是 App.ChatV2.TaskPlanner（大写 T）。
        const planner = resolvePlanner(taskPlanner);
        if (!planner || typeof planner.plan !== 'function') {
            ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] TaskPlanner 未就绪（plan 不可用），降级为通用对话');
            yield { type: 'fallback_to_agent', reason: 'planner_missing', detail: 'App.ChatV2.TaskPlanner.plan 不可用' };
            return;
        }

        // 通知渲染层：此 turn 需要显示步骤
        turn._enableSteps = true;

        const userMsg = turn._originalUserMessage || turn.userMessage || turn.message || '';
        ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 开始多步骤规划: "${userMsg.slice(0, 80)}..."`);

        // 判断是否需要动态规划（ReAct 模式）：分析/汇总类请求需要先看到数据才能确定参数
        const intent = turn?.context?._intent || null;
        const needsDynamicPlanning = (intent && Array.isArray(intent.operations) &&
            intent.operations.some(op => /analyze|aggregate|汇总|分析|统计|对比|排名/.test(op))) ||
            // 兜底：如果 intent 不可用，根据用户消息关键词判断
            /分析|汇总|聚合|统计|对比|排名|卖得|最好|最差|趋势|分布/.test(userMsg);

        let orderedSteps = [];
        let steps;

        if (needsDynamicPlanning) {
            // 【ReAct 动态模式】先只规划第1步（读取数据），执行后再动态规划后续步骤
            ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 启用动态规划模式（ReAct）：先执行读取步骤，观察数据后再规划后续');
            yield { type: 'status', text: '📋 动态规划模式：先读取数据...' };

            try {
                steps = await planner.plan({
                    userMessage: userMsg,
                    tools,
                    messages,
                    signal,
                    maxSteps: 1,
                    intent,
                    context: turn?.context
                });
            } catch (planErr) {
                ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 初始规划失败，降级为通用对话:', planErr?.message);
                yield { type: 'status', text: '⚠️ 规划步骤失败，转为通用对话模式' };
                yield { type: 'fallback_to_agent', reason: 'plan_failed', error: planErr?.message };
                return;
            }

            if (!Array.isArray(steps) || steps.length === 0) {
                ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 无需多步骤，降级为通用对话');
                yield { type: 'fallback_to_agent', reason: 'no_steps_needed' };
                return;
            }

            orderedSteps = steps;
            ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 初始规划: ${orderedSteps.length} 个步骤`);

            // 【条件读取】动态模式下，只有当 turn.context.selection.headers 不可用时才强制先读取
            // 如果 headers 已从 UI 层预取，直接规划汇总/分析步骤，不操作 Excel
            const sel = turn?.context?.selection || {};
            const hasSelectionHeaders = Array.isArray(sel.headers) && sel.headers.length > 0;

            if (orderedSteps.length > 0 && !hasSelectionHeaders) {
                const firstStep = orderedSteps[0];
                const firstTool = String(firstStep.toolName || firstStep.tool || '');
                const isReadStep = /^(excel_)?read/.test(firstTool) || /读取/.test(firstTool);
                if (!isReadStep) {
                    const readRange = sel.range || sel.address || 'A1';
                    const readSheet = sel.sheetName || sel.worksheet || '';

                    orderedSteps[0] = {
                        id: `step_read_${Date.now()}`,
                        name: '读取数据',
                        description: `读取工作表数据 (${readSheet || '当前表'} ${readRange})`,
                        strategy: 'direct',
                        toolName: 'excel_read_range',
                        params: {
                            range: readRange,
                            ...(readSheet ? { sheet: readSheet } : {})
                        },
                        displayArgs: { 工作表: readSheet || '当前表', 范围: readRange },
                        dependsOn: [],
                        status: 'pending'
                    };

                    ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 动态模式替换第1步为读取（selection.headers 为空）: ${readSheet || '当前表'} ${readRange}`);
                    yield { type: 'status', text: '📋 先读取数据...' };
                }
            } else if (hasSelectionHeaders) {
                ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 动态模式跳过读取（selection.headers 已就绪: ${sel.headers.join(', ')}）`);
            }
        } else {
            // 【静态模式】一次性规划所有步骤（简单/明确的请求）
            yield { type: 'status', text: '📋 正在规划步骤...' };
            try {
                steps = await planner.plan({
                    userMessage: userMsg,
                    tools,
                    messages,
                    signal,
                    maxSteps: MAX_STEPS,
                    intent,
                    context: turn?.context
                });
            } catch (planErr) {
                ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 规划步骤失败，降级为通用对话:', planErr?.message);
                yield { type: 'status', text: '⚠️ 规划步骤失败，转为通用对话模式' };
                yield { type: 'fallback_to_agent', reason: 'plan_failed', error: planErr?.message };
                return;
            }

            if (!Array.isArray(steps) || steps.length === 0) {
                ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 无需多步骤，降级为通用对话');
                yield { type: 'fallback_to_agent', reason: 'no_steps_needed' };
                return;
            }

            if (steps.length > MAX_STEPS) {
                ChatLog.warn?.('MULTI-STEP', `[MultiStepEngine] 步骤数 ${steps.length} 超过上限 ${MAX_STEPS}，截断`);
                steps = steps.slice(0, MAX_STEPS);
            }

            ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 规划完成: ${steps.length} 个步骤`);
            yield { type: 'status', text: `📋 规划完成: ${steps.length} 个步骤` };

            // 步骤顺序完全按 LLM 规划结果执行（不再经 TaskScheduler 调度/重排）
            orderedSteps = steps;
        }

        // 【条件插入读取】静态模式下，只有当 selection.headers 不可用时才插入读取步骤
        // 如果 headers 已从 UI 层预取，直接执行汇总/分析，不先操作 Excel
        const selStatic = turn?.context?.selection || {};
        const hasSelHeaders = Array.isArray(selStatic.headers) && selStatic.headers.length > 0;

        if (!needsDynamicPlanning && orderedSteps.length > 0 && !hasSelHeaders) {
            const firstStep = orderedSteps[0];
            const firstTool = String(firstStep.toolName || firstStep.tool || '');
            const isReadStep = /^(excel_)?read/.test(firstTool) || /读取/.test(firstTool);
            const needsData = orderedSteps.some(s => /aggregate|chart|汇总|图表|分析|统计/.test(String(s.toolName || s.tool || '')));

            if (!isReadStep && needsData) {
                const readRange = selStatic.range || selStatic.address || 'A1';
                const readSheet = selStatic.sheetName || selStatic.worksheet || '';

                const readStep = {
                    id: `step_read_${Date.now()}`,
                    name: '读取数据',
                    description: `读取工作表数据 (${readSheet || '当前表'} ${readRange})`,
                    strategy: 'direct',
                    toolName: 'excel_read_range',
                    params: {
                        range: readRange,
                        ...(readSheet ? { sheet: readSheet } : {})
                    },
                    displayArgs: { 工作表: readSheet || '当前表', 范围: readRange },
                    dependsOn: [],
                    status: 'pending'
                };

                orderedSteps.unshift(readStep);
                for (let i = 1; i < orderedSteps.length; i++) {
                    if (Array.isArray(orderedSteps[i].dependsOn)) {
                        orderedSteps[i].dependsOn = orderedSteps[i].dependsOn.map(depId =>
                            depId === firstStep.id ? readStep.id : depId
                        );
                    }
                }

                ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 静态模式插入读取步骤（selection.headers 为空）: ${readSheet || '当前表'} ${readRange}`);
                yield { type: 'status', text: '📋 自动添加数据读取步骤...' };
            }
        } else if (!needsDynamicPlanning && hasSelHeaders) {
            ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 静态模式跳过读取（selection.headers 已就绪: ${selStatic.headers.join(', ')}）`);
        }

        // 展示规划面板
        progressTracker.init(orderedSteps, turn, sessionStore);
        // 【步骤统一到 cbx v20261003】不再写 _planHtml 旧绿面板（.exec-plan-card），
        // 改把每一步写进 thoughtList（kind='step'），由 contentEl 顶部 cbx 列表统一展示。
        // 下面所有"刷新规划面板"的位置都复用它（幂等 upsert：只更新状态，不新增行）。
        function syncStepsToThoughts() {
            if (!sessionStore || typeof sessionStore.upsertStepThought !== 'function') return;
            orderedSteps.forEach(function (s, i) {
                sessionStore.upsertStepThought(turn.id, {
                    id: s.id || s.stepId || ('ms_step_' + i),
                    title: s.description || s.name || ('步骤 ' + (i + 1)),
                    detail: s.detail || '',
                    status: s.status || 'pending',
                    toolName: s.toolName || null,
                    index: i
                });
            });
        }
        syncStepsToThoughts();

        // 4. 执行阶段：顺序执行各步骤（MVP）
        const results = [];
        const stepContext = {};  // 步骤间上下文传递（如新创建的工作表名）
        let totalIterations = 0;

        // 前端聚合始终启用：聚合是纯计算操作，走内存 clientAggregate，不碰 Excel。
        // 当有图表步骤时，聚合结果会写入新表供图表引用（这是写入操作，不是聚合操作）。
        // 此标志现仅影响日期分组（clientDateGroup）的前端计算偏好。
        const hasChartStep = orderedSteps.some(s => /chart|图表|柱状图|折线图|饼图|面积图|散点图|雷达图|环形图/i.test(String((s && (s.toolName || s.tool)) || '')));
        const enableClientAggregate = !hasChartStep;
        if (enableClientAggregate) {
            ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 无图表步骤，聚合和日期分组结果在面板显示');
        }

        // 【任务套任务】监听子步骤状态更新事件，实时刷新规划面板
        const subStepListener = (evt) => {
            if (!evt || evt.turnId !== turn.id) return;
            const targetStep = orderedSteps.find(s => s.id === evt.stepId);
            if (targetStep && evt.subSteps) {
                targetStep.subSteps = evt.subSteps;
                try {
                    syncStepsToThoughts();
                } catch (e) {
                    ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 子步骤面板刷新失败:', e?.message);
                }
            }
        };
        if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.on === 'function') {
            App.ChatV2.agentEvents.on('multi_step:substep_update', subStepListener);
        }

        for (let i = 0; i < orderedSteps.length; i++) {
            if (signal?.aborted) {
                yield { type: 'status', text: '⏹️ 用户取消执行' };
                break;
            }

            const step = orderedSteps[i];
            step.status = 'running';
            progressTracker.updateStep(i, step, turn, sessionStore);

            // 步骤开始时同步刷新（显示"执行中"状态），实现实时进度
            syncStepsToThoughts();

            yield {
                type: 'status',
                text: `▶️ 步骤 ${i + 1}/${orderedSteps.length}: ${step.description || step.name || ''}`
            };

            // 通过 StepExecutor 执行步骤（含重试逻辑）
            // 注：不再做"前置失败即跳过"的规则判断——每一步都按 LLM 规划执行，
            // 失败信息回喂后由 LLM 自行决定继续/收尾。
            const stepResult = await stepExecutor.executeStep(step, i, {
                turn, signal, sessionStore, toolRegistry, agentLoop,
                tools, messages, results, totalIterations, stepContext,
                // 传完整计划步骤：StepExecutor 需按索引把"图表步骤"与"它自己的聚合步骤"配对取源（多图表场景）
                steps: orderedSteps,
                enableClientAggregate
            });

            totalIterations += stepResult.iterations || 1;
            results.push(stepResult);

            // 步骤执行后提取上下文信息（如新创建的工作表名），供后续步骤使用
            if (stepResult.ok && stepContext) {
                stepExecutor.extractStepContext(step, stepResult, stepContext);
            }

            // 更新步骤状态
            step.status = stepResult.ok ? 'completed' : 'failed';
            step.result = stepResult;
            progressTracker.updateStep(i, step, turn, sessionStore);

            // 重新渲染规划面板（步骤完成/失败状态同步到 cbx 行）
            syncStepsToThoughts();

            if (stepResult.ok) {
                yield { type: 'status', text: `✅ 步骤 ${i + 1} 完成: ${step.description || step.name || ''}` };
            } else {
                const errMsg = stepResult.error
                    ? (typeof stepResult.error === 'string' ? stepResult.error : stepResult.error.message || '执行失败')
                    : '执行失败';
                yield { type: 'status', text: `⚠️ 步骤 ${i + 1} 失败: ${errMsg}（继续下一步）` };
            }

            // 【ReAct 动态规划】动态模式下，基于观察结果让 LLM 决定下一步。
            // 注：已移除"操作类型去重 / 汇总+产出早停 / 拒绝重复操作 / 读取保护"等
            // 替 LLM 做决定的规则；是否继续、是否重复某类操作，完全由 LLM 判断。
            // 仅保留最小护栏：步骤上限（orderedSteps.length < MAX_STEPS）+ abort。
            if (needsDynamicPlanning && stepResult.ok && orderedSteps.length < MAX_STEPS) {

                const executedSteps = orderedSteps.slice(0, i + 1).map((s, idx) => ({
                    step: s,
                    result: results[idx]
                }));

                ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 动态重规划: 基于 ${executedSteps.length} 个已执行步骤的观察结果`);
                yield { type: 'status', text: '🔍 观察执行结果，动态规划下一步...' };

                try {
                    const nextSteps = await planner.replan({
                        userMessage: userMsg,
                        executedSteps,
                        tools,
                        messages,
                        signal,
                        intent
                    });

                    if (Array.isArray(nextSteps) && nextSteps.length > 0) {
                        const nextStep = nextSteps[0];
                        const nextToolName = String(nextStep.toolName || nextStep.tool || '');
                        nextStep.id = nextStep.id || `step_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
                        orderedSteps.push(nextStep);
                        ChatLog.info?.('MULTI-STEP', `[MultiStepEngine] 动态规划下一步: ${nextStep.name || nextStep.description} (${nextToolName})`);
                        yield { type: 'status', text: `📋 动态规划: ${nextStep.name || nextStep.description}` };

                        syncStepsToThoughts();
                    } else {
                        ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 动态规划返回空，任务完成');
                        yield { type: 'status', text: '✅ 任务已完成' };
                    }
                } catch (replanErr) {
                    ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 动态重规划异常:', replanErr?.message);
                    yield { type: 'status', text: '⚠️ 动态规划异常，继续执行剩余步骤' };
                }
            }
        }

        // 清理子步骤事件监听
        if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.off === 'function') {
            App.ChatV2.agentEvents.off('multi_step:substep_update', subStepListener);
        }

        // 5. 汇总阶段：使用 ResultAggregator
        yield { type: 'status', text: '📝 正在汇总结果...' };
        let summary;
        try {
            const llmIo = App.ChatV2.llmIo;
            const aggResult = await resultAggregator.aggregate({
                steps: orderedSteps,
                results,
                llmIo: llmIo?.askText ? llmIo : null,
                messages,
                config: App.config,
                signal
            });
            summary = aggResult.summary;
        } catch (aggErr) {
            ChatLog.warn?.('MULTI-STEP', '[MultiStepEngine] 汇总异常，使用内置摘要:', aggErr?.message);
            summary = buildFallbackSummary(orderedSteps, results);
        }

        yield { type: 'multi_step_done', steps: orderedSteps, results, summary };
        ChatLog.info?.('MULTI-STEP', '[MultiStepEngine] 多步骤管道执行完成');
    }

    /**
     * 兜底摘要（ResultAggregator 不可用时）
     */
    function buildFallbackSummary(steps, results) {
        const total = steps.length;
        const succeeded = results.filter(r => r && r.ok).length;
        const failed = results.filter(r => r && !r.ok).length;

        const parts = [];
        if (succeeded === total) {
            parts.push(`✅ 已完成全部 ${total} 个步骤`);
        } else if (failed === total) {
            parts.push(`⚠️ 所有 ${total} 个步骤均未能完成`);
        } else {
            parts.push(`✅ 已完成 ${succeeded}/${total} 个步骤`);
            if (failed > 0) {
                const failedNames = steps
                    .map((s, i) => (!results[i]?.ok ? s.description || s.name || `步骤${i + 1}` : null))
                    .filter(Boolean);
                if (failedNames.length > 0) {
                    parts.push(`⚠️ 未完成: ${failedNames.join('、')}`);
                }
            }
        }
        return parts.join('\n');
    }

    App.ChatV2.MultiStepEngine = { run, planSkeleton };
    ChatLog.info?.('CHAT-V2', '✅ MultiStepEngine loaded (已收敛：仅 planSkeleton 生效，run 为回滚保留)');
})();