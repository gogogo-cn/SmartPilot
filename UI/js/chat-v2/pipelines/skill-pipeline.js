/**
 * ============================================================
 * chat-v2/pipelines/skill-pipeline.js — 通用 Skill 执行流程（主文件/编排内核）
 * ============================================================
 * 职责：
 *   - 执行由 chat-orchestrator 通用匹配到的 Skill
 *   - 不硬编码任何 Skill 名称，只负责把 turn 上下文交给 tool-registry 执行
 *   - 使用 invokeTracked 通用方法，Skill 执行自动显示为工具卡片
 *
 * 模块化结构（2026-09 拆分，行为零变更）：
 *   - pipelines/skill/skill-shared-util.js    → SP.util        共享纯函数工具
 *   - pipelines/skill/skill-clean-crud.js     → SP.cleanCrud   清洗/CRUD 执行器
 *   - pipelines/skill/skill-spec-build.js     → SP.specBuild   spec 步骤构建
 *   - pipelines/skill/skill-steps-panel.js    → SP.stepsPanel  步骤面板渲染
 *   - pipelines/skill/skill-spec-execute.js   → SP.specExecute spec 模式执行器
 *   - 本文件保留：follow-up 继承、agentMode、execute 主流程、args/结果组装、注册挂载
 *   模块间经 App.ChatV2.skillPipelineModules 通信，对外契约（execute 事件流）不变。
 *
 * 触发条件（由 chat-orchestrator 决定）：
 *   - 用户消息与某个已注册 Skill 的 keywords / inputContext 匹配
 *   - 且该 Skill 所需的上下文（选区/引用区域/附件）已满足
 *
 * 与主流 Agent 框架对齐：
 *   - OpenAI GPTs / Anthropic MCP / LangChain Toolkits：
 *     能力通过注册表声明，系统按意图路由到对应执行器。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    // 拆分模块命名空间（pipelines/skill/ 目录下各模块挂载于此），调用时经 SP.<模块>.<函数> 解析
    const SP = App.ChatV2.skillPipelineModules || {};

    // ── Spec 引擎执行超时守卫 ──────────────────────────────────────────────
    // 背景：executeSkillSpecMode 是 async generator，若其内部某个 await 永挂
    // （如非 Excel 宿主 type=text 选区下桥调用无响应）或逻辑死循环不 yield 终态事件，
    // 外层 `for await` 永不结束、catch 永不触发 → 步骤面板永久停在"执行中"。
    // 给生成器加【绝对截止时间】：从进入起算，超过 SPEC_EXEC_TIMEOUT_MS 仍未结束则强制
    // gen.return() 终止并抛错，由 spec 分支 catch 降级为文字管道，保证面板一定能收尾。
    // 既能兜住"单个 await 永挂"（超时在等待 next() 时触发），
    // 也能兜住"密集死循环不 yield 终态"（每次 next() 立即返回、逐个检查 deadline）。
    const SPEC_EXEC_TIMEOUT_MS = (window.App && App.ChatV2 && App.ChatV2.skillSpecTimeoutMs)
        ? App.ChatV2.skillSpecTimeoutMs
        : 120000; // 默认 2 分钟

    function withSkillSpecDeadline(gen, timeoutMs, onTimeoutMsg) {
        const _ms = (typeof timeoutMs === 'number' && timeoutMs > 0) ? timeoutMs : SPEC_EXEC_TIMEOUT_MS;
        const deadline = Date.now() + _ms;
        let pending;
        try { pending = gen.next(); }
        catch (e) { const _e = e; return (async function* () { throw _e; })(); }
        if (!pending || typeof pending.then !== 'function') {
            const _sync = pending;
            return (async function* () { if (_sync && !_sync.done && _sync.value !== undefined) yield _sync.value; })();
        }
        return (async function* () {
            let current = pending;
            // 吞掉超时后悬挂的 next() 拒绝，避免未处理 Promise 拒绝刷屏
            current.catch(function (err) { if (typeof ChatLog !== 'undefined') ChatLog.debug?.('PIPELINE', '[skill-pipeline] 悬挂 next() 拒绝（超时后预期发生）:', err && err.message); });
            while (true) {
                const remaining = deadline - Date.now();
                if (remaining <= 0) {
                    try { gen.return(); } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('SKILL', '技能管道操作失败:', _ && _.message); }
                    throw new Error(onTimeoutMsg || 'SkillSpec 执行超时（已自动终止）');
                }
                const timeout = new Promise(function (_, rej) {
                    setTimeout(function () { rej(new Error('__SKILL_SPEC_TIMEOUT__')); }, remaining);
                });
                let nx;
                try {
                    nx = await Promise.race([current, timeout]);
                } catch (_) {
                    try { gen.return(); } catch (__) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PIPELINE', 'gen.return 失败:', __ && __.message); }
                    throw new Error(onTimeoutMsg || 'SkillSpec 执行超时（已自动终止）');
                }
                if (nx.done) return;
                yield nx.value;
                try { current = gen.next(); }
                catch (e) { throw e; }
                if (!current || typeof current.then !== 'function') {
                    if (current && !current.done && current.value !== undefined) yield current.value;
                    return;
                }
                current.catch(function (err) { if (typeof ChatLog !== 'undefined') ChatLog.debug?.('PIPELINE', '[skill-pipeline] 悬挂 next() 拒绝（超时后预期发生）:', err && err.message); });
            }
        })();
    }

    // ============================================================
    // 后续消息继承（Phase 5）：读取上一轮 _execSpec，注入上轮结果
    // 支持跨管道：上一轮是 skill/text/python 分析类管道时同样继承数据快照，
    // 避免"生成图表"等后续诉求因无新选区而失败。
    // ============================================================
    function getSkillFollowUp(turn) {
        if (!turn || !turn.context || !turn.context._execFollowUp) return null;
        const ss = App.ChatV2.sessionStore;
        if (!ss || typeof ss.getLastExecSpec !== 'function') return null;
        const prev = ss.getLastExecSpec();
        if (!prev || !prev.spec) return null;
        const p = prev.spec.pipeline;
        if (p !== 'skill' && p !== 'text' && p !== 'python') return null;
        return prev.spec;
    }

    function buildSkillFollowUpPrompt(spec, userMessage, intent) {
        const parts = [];
        const res = (spec && spec.result) || {};
        if (res.spec) {
            parts.push('【上一轮方案规格】\n' + JSON.stringify(res.spec));
        }
        if (res.outputText) {
            parts.push('【上一轮执行结果】\n' + String(res.outputText).slice(0, 2000));
        }
        // 意图分析器构造规划 prompt（任务结构 + 用户新需求 + 意图 + 规划指令）
        const analyzer = App.ChatV2.intentAnalyzer;
        if (analyzer && typeof analyzer.buildFollowUpPlanningPrompt === 'function') {
            const planning = analyzer.buildFollowUpPlanningPrompt({
                prevSpec: { spec: spec },
                userMessage: userMessage || '',
                intent: intent || null
            });
            if (planning) parts.push(planning);
        } else {
            // 兜底：旧逻辑（意图分析器不可用时）
            if (res.summary) {
                parts.push('【上一轮分析结论】\n' + String(res.summary).slice(0, 2000));
            }
            // 注入上一轮任务结构（任务清单/执行步骤），让后续调整基于上轮规划而非只靠关键词。
            // 兼容两种结构：skill 管道存 ts.steps；python/text 分析管道存 ts.tasks。
            const ts = spec && spec.taskStructure;
            if (ts) {
                const steps = Array.isArray(ts.steps) ? ts.steps : (Array.isArray(ts.tasks) ? ts.tasks : null);
                if (steps && steps.length) {
                    const lines = steps.map(function (s, i) {
                        if (typeof s === 'string') return (i + 1) + '. ' + s;
                        return (i + 1) + '. ' + (s.displayName || s.title || s.chartLabel || s.detail || '步骤');
                    });
                    parts.push('【上一轮任务规划】\n' + lines.join('\n'));
                }
            }
        }
        if (parts.length) {
            parts.push('请基于以上上一轮结果，按我的新需求调整。只输出调整后的完整结果，不要只描述改动。');
            return parts.join('\n\n');
        }
        return '';
    }

    /** 后续消息继承：富化 userMessage（保留 _originalUserMessage 供渲染）并继承数据快照 */
    function applySkillFollowUp(turn) {
        const prevSpec = getSkillFollowUp(turn);
        if (!prevSpec) return;
        if (turn._originalUserMessage === undefined) {
            turn._originalUserMessage = turn.userMessage || '';
        }
        const intent = (turn.context && turn.context._execFollowUp && turn.context._execFollowUp.intent) || null;
        const prevCtx = buildSkillFollowUpPrompt(prevSpec, turn.userMessage || '', intent);
        if (prevCtx) {
            turn.userMessage = (turn.userMessage || '') + '\n\n' + prevCtx;
        }
        // v20260924：读取端加 schema 守卫 + 专有字段兼容 ext.skill 和旧根字段
        const pd = prevSpec.data || {};
        const ctx = turn.context || {};
        // selection 继承：校验 address/range + sheet 必需字段
        const sel = pd.selection;
        if (!ctx.selection && sel && typeof sel === 'object') {
            const hasAddr = sel.address || sel.range;
            const hasSheet = sel.sheet || sel.sheetName || sel.worksheet;
            if (hasAddr && hasSheet) ctx.selection = sel;
        }
        if (!ctx.referenceRegions && Array.isArray(pd.referenceRegions) && pd.referenceRegions.length) {
            ctx.referenceRegions = pd.referenceRegions;
        }
        // 专有字段：优先从 ext.skill（v20260924 新写入位置），兼容旧根字段
        const extSkill = pd.ext && pd.ext.skill ? pd.ext.skill : {};
        const aggResult = extSkill.aggregateResult !== undefined ? extSkill.aggregateResult : pd.aggregateResult;
        const dataRange = extSkill.chartDataRange !== undefined ? extSkill.chartDataRange : pd.chartDataRange;
        if (aggResult && !ctx._aggregateResult) {
            ctx._aggregateResult = aggResult;
            ChatLog.info?.('PIPELINE', '[skill-pipeline] 后续消息继承: 复用上一轮聚合结果表 ' + (aggResult.dataRange || ''));
        }
        if (dataRange && !ctx._chartDataRange) {
            ctx._chartDataRange = dataRange;
            ChatLog.info?.('PIPELINE', '[skill-pipeline] 后续消息继承: 复用上一轮图表数据源 ' + dataRange);
        }
        ChatLog.info?.('PIPELINE', '[skill-pipeline] 后续消息继承: 复用上一轮 skill 结果');
    }

    /** 写入通用执行上下文（多轮继承）：供后续消息继承 */
    function writeSkillExecSpec(turn, sessionStore, result) {
        if (!sessionStore || typeof sessionStore.setExecSpec !== 'function') return;
        // v20260924：用工厂统一构建 data，selection/referenceRegions schema 不再手写
        const fact = sessionStore.execSpecFactory;
        const commonData = fact && typeof fact.buildCommonData === 'function'
            ? fact.buildCommonData(turn, {
                // Skill pipeline 专有扩展（挂 data.ext.skill）
                skill: {
                    aggregateResult: result?.aggregate || null,
                    chartDataRange: result?.dataRange || null
                }
            })
            : { // 兜底（旧 session-store 版本）
                aggregateResult: result?.aggregate || null,
                chartDataRange: result?.dataRange || null
              };
        const skill = turn.context?._matchedSkill || null;
        sessionStore.setExecSpec(turn.id, {
            pipeline: 'skill',
            data: commonData,
            params: {
                skillName: (skill && (skill.name || skill.Name)) || '',
                displayName: (skill && (skill.displayName || skill.DisplayName)) || '',
                executionMode: result?.executionMode || ''
            },
            result: {
                // 取消态不算 success：用户拒绝执行，execSpec 应如实记录
                success: !!(result && result.success !== false && !result.cancelled),
                spec: result?.spec || null,
                outputText: result?.outputText || '',
                steps: Array.isArray(result?.steps) ? result.steps : null
            },
            // 任务结构（方案规格/执行步骤）：供后续消息继承时注入规划输入
            taskStructure: {
                spec: result?.spec || null,
                steps: Array.isArray(result?.steps) ? result.steps : null
            },
            ts: Date.now()
        });
    }
    /**
     * 自主 Agent 模式执行。
     * 流程：Skill 以"LLM System Prompt + 业务 MCP 工具"驱动 agent-loop 的 ReAct 多步循环，
     * 自动获得 token 水位监控 + 步骤日志持久化（agent-loop 内置护栏），直到完成/超限/取消。
     * 设计对齐 multi-step-agent-mcp：业务思考/终止判断由 Skill 约束，计数/护栏下沉到 agent-loop。
     * @param {Object} turn
     * @param {Object} skill
     * @param {string} displayName
     * @param {AbortSignal} signal
     * @returns {AsyncGenerator} agent-loop 的 typed events
     */
    async function* executeSkillAgentMode(turn, skill, displayName, signal) {
        const sessionStore = App.ChatV2.sessionStore;
        const toolRegistry = App.ChatV2.toolRegistry;
        const agentLoop = App.ChatV2.agentLoop;

        if (!agentLoop || typeof agentLoop.run !== 'function') {
            throw new Error('agentLoop 未就绪，无法执行 Agent 模式 Skill');
        }
        if (typeof toolRegistry?.init === 'function') {
            await toolRegistry.init();
        }

        const agentCfg = skill?.agentConfig || {};
        // 技能可声明更多轮次，但封顶，避免 agentConfig.maxIterations 被写成极大值后空转。
        const SKILL_AGENT_ROUND_CAP = 12;
        const requestedRounds = Number(agentCfg.maxIterations) > 0
            ? Number(agentCfg.maxIterations)
            : (agentLoop.MAX_TOOL_ROUNDS || 5);
        const maxIterations = Math.min(requestedRounds, SKILL_AGENT_ROUND_CAP);

        const userMessage = turn.userMessage || '';
        const systemPrompt = skill?.llmSystemPrompt || skill?.description ||
            `你是 ${displayName}，一位专业的 Excel 办公助手。请根据用户目标，自主调用可用工具逐步完成，并在完成时给出清晰结论。`;

        // 给 LLM 的可用工具：业务 MCP 工具（agent 模式在循环内按需迭代调用）。
        // 【Skill 级工具白名单】skill.allowedTools 声明后只保留白名单内工具（支持尾缀通配符），
        // 未声明/空则保持现状——全部业务 MCP 工具可见。白名单直接约束 run() 的 tools 参数，
        // agent-loop 内部 findCandidateTools 亦以该 tools 作为候选白名单，两端口径一致。
        let tools = [];
        try {
            const allTools = toolRegistry.listBySource('mcp') || [];
            const allowed = (Array.isArray(skill?.allowedTools) && skill.allowedTools.length > 0)
                ? skill.allowedTools : null;
            if (!allowed) {
                tools = allTools;
            } else {
                const patterns = allowed.map(n => String(n).trim().toLowerCase()).filter(Boolean);
                tools = allTools.filter(t => {
                    const name = String(t?.name || '').toLowerCase();
                    return patterns.some(p => p.endsWith('*') ? name.startsWith(p.slice(0, -1)) : name === p);
                });
                ChatLog.info?.('PIPELINE',
                    `[skill-agent] Skill ${skill?.name || ''} 工具白名单生效: ${tools.length}/${allTools.length} 个工具可见`);
            }
            // 🛡️ 领域硬约束（skill 未声明 allowedTools 时也生效，双保险）：
            // cleaner / crud / ocr 等领域 skill 在 agent 模式下不得看到 aggregate/chart/pivot 等分析展示工具，
            // 防止 LLM 看到"合并/拆分"就脑补"先聚合→画图"。
            // 根因：ExcelMcp 平台有 50+ 工具，agent 模式全暴露后 LLM 幻觉概率极高。
            const DOMAIN_BLOCKLIST = {
                cleaner: ['excel_aggregate', 'excel_create_chart', 'excel_create_pivot', 'excel_create_pivot_table', 'excel_summary'],
                crud: ['excel_aggregate', 'excel_create_chart', 'excel_create_pivot', 'excel_create_pivot_table', 'excel_summary', 'excel_merge_sheets', 'excel_split_by_column', 'excel_split_workbook', 'excel_split_by_rows', 'excel_merge_workbooks']
            };
            const domainKey = String(skill?.skillType || '').toLowerCase();
            if (DOMAIN_BLOCKLIST[domainKey]) {
                const blockSet = new Set(DOMAIN_BLOCKLIST[domainKey]);
                const before = tools.length;
                tools = tools.filter(t => !blockSet.has(String(t?.name || '').toLowerCase()));
                if (tools.length < before) {
                    ChatLog.info?.('PIPELINE', `[skill-agent] 领域硬约束 domain=${domainKey} 拦截 ${before - tools.length} 个工具，剩余 ${tools.length} 个`);
                }
            }
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[skill-agent] 获取 MCP 工具列表失败:', e?.message);
        }

        const messages = [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userMessage }
        ];

        yield { type: 'status', text: `正在以自主 Agent 模式执行 ${displayName}（最多 ${maxIterations} 步）...` };

        // 骨架已由 execute() 入口统一 init，此处直接进入循环
        let currentRoundLabel = '';

        // 【P1 护栏】token 水位监控配置：Skill 声明的 agentConfig 优先，缺失时回退 agent-loop 默认
        const tokenGuard = {
            maxTokens: Number(agentCfg.maxContextTokens) > 0 ? Number(agentCfg.maxContextTokens) : undefined,
            warnRatio: Number(agentCfg.tokenWarnRatio) > 0 ? Number(agentCfg.tokenWarnRatio) : undefined,
            criticalRatio: Number(agentCfg.tokenCriticalRatio) > 0 ? Number(agentCfg.tokenCriticalRatio) : undefined
        };

        // 拦截 agent-loop 的事件流：按 round 号匹配骨架步骤并填色，
        // 骨架从 pending → running → completed 逐步点亮，detail 聚合本轮所有工具调用摘要。
        // 原事件继续 yield 给外层消费（保持原行为不变）。
        let currentRound = 0;
        // persistTools：每轮都保留工具，由模型决定继续调用还是收尾。
        // 不传时第一轮成功后就不再下发工具，自主 Agent 会停在第一步。
        for await (const evt of agentLoop.run({ turn, messages, tools, signal, maxIterations, tokenGuard, persistTools: true })) {
            // round_start：匹配骨架步骤（_round === evt.round），update 为 running
            if (evt.type === 'round_start') {
                currentRound = evt.round;
                const roundIdx = turn._skillSteps ? turn._skillSteps.findIndex(s => s._round === evt.round) : -1;
                if (roundIdx >= 0) {
                    const prev = turn._skillSteps[roundIdx];
                    turn._skillSteps[roundIdx] = { ...prev, status: 'running', detail: `第 ${evt.round} 轮执行中...` };
                    SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                }
                ChatLog.info?.('PIPELINE', `[skill-agent] 骨架填色: 第 ${evt.round} 轮 → running`);
            }

            // tool_call_result：把工具调用摘要 merge 到当前轮骨架步骤的 detail 里
            if (evt.type === 'tool_call_result') {
                const toolName = evt.toolName || '未知工具';
                const ok = !!evt.ok;
                const short = ok ? `${toolName}✓(${evt.durationMs || 0}ms)` : `${toolName}✕`;

                // 优先按 round 匹配骨架步骤；没有骨架时回退到 append
                let matched = -1;
                if (currentRound > 0) {
                    matched = turn._skillSteps ? turn._skillSteps.findIndex(s => s._round === currentRound) : -1;
                }
                if (matched < 0) {
                    // 回退：无骨架 / round 不匹配 → 追加为独立步骤
                    SP.stepsPanel.appendSkillStep(turn, sessionStore, {
                        title: toolName,
                        detail: ok ? (currentRound ? `第 ${currentRound} 轮 · 执行成功` : '执行成功')
                                   : (currentRound ? `第 ${currentRound} 轮 · 失败：${(evt.error || '').slice(0,80)}` : `失败：${(evt.error || '').slice(0,80)}`),
                        status: ok ? 'completed' : 'failed',
                        strategy: 'agent',
                        toolName: toolName,
                        result: evt.result || null
                    });
                } else {
                    // 骨架匹配成功 → merge 到骨架步骤
                    const prev = turn._skillSteps[matched];
                    const prevDetail = prev.detail || '';
                    const newDetail = prevDetail.includes('执行中')
                        ? `第 ${currentRound} 轮 · ${short}`
                        : `${prevDetail}；${short}`;
                    const merged = {
                        ...prev,
                        detail: newDetail,
                        strategy: 'agent'
                    };
                    // 如果这个 round 已经有失败，整个 round 标 failed
                    if (!ok) merged.status = 'failed';
                    turn._skillSteps[matched] = merged;
                    SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                }
            }

            yield evt;
        }

        // agent-loop 不负责 completeTurn（与 text-pipeline 一致），由调用方完成
        const finalTurn = sessionStore.getCurrentTurn();
        if (finalTurn && finalTurn.id === turn.id) {
            sessionStore.completeTurn(turn.id, {
                content: finalTurn.content || '',
                reasoning: finalTurn.reasoning || ''
            });
        }
        // 最终收尾：所有未失败的骨架步骤标 completed；双写 turn._steps
        // 把仍处于 pending/running 的骨架步骤（LLM 没有走满所有 round）标 completed 或保留
        if (Array.isArray(turn._skillSteps)) {
            turn._skillSteps.forEach(step => {
                if (step.status === 'pending' || step.status === 'running') {
                    step.status = 'completed';
                }
            });
        }
        SP.stepsPanel.finalizeSkillSteps(turn, sessionStore);
        ChatLog.info?.('PIPELINE', `[skill-agent] 骨架填色完成: ${turn._skillSteps?.length || 0} 步`);
    }

    /**
     * 判断一次 Skill 执行结果是否"什么都没做成"。
     * 仅在异常之外的路径上作为兜底判据使用（保守策略：宁可放过，不误判）。
     *   - 有 outputText → 认为有效（正常路径一定会生成"✅ xx完成"摘要）
     *   - 无 outputText 且无 steps → 视为空结果
     *   - 无 outputText 且 steps（含 subSteps）全部 failed → 视为失败结果
     */
    function isSkillResultEmpty(res) {
        if (!res || typeof res !== 'object') return true;
        if (String(res.outputText || '').trim()) return false;
        const steps = Array.isArray(res.steps) ? res.steps : [];
        const flat = [];
        const walk = (arr) => {
            for (const s of arr) {
                if (!s || typeof s !== 'object') continue;
                flat.push(s);
                if (Array.isArray(s.subSteps)) walk(s.subSteps);
            }
        };
        walk(steps);
        if (flat.length === 0) return true;
        return flat.every(s => s.status === 'failed');
    }

    /**
     * Skill 执行失败 → 降级到通用 text-pipeline 重试（静默降级，不向用户暴露错误）。
     *
     * 用途：Skill 本身不可用 / 执行抛异常 / 执行结果为空时，不再直接把 turn 置 failed，
     * 而是清掉 Skill 骨架面板、复位路由标记，改由通用管道（LLM 自主选 MCP 工具）再跑一次。
     *
     * 注意：
     *   - 必须在 sessionStore.failTurn 之前调用（failTurn 会把 currentTurn 置 null，
     *     降级后 text-pipeline 的 completeTurn 将无处落笔）。
     *   - `__skillFallbackTried` 单次标记，防止降级后再次回到 Skill 形成死循环。
     */
    async function* fallbackToText(turn, signal, reason) {
        const sessionStore = App.ChatV2.sessionStore;
        const reasonMsg = reason || 'Skill 执行失败';

        // 用户主动取消 → 不做降级重试
        if (signal?.aborted) {
            yield { type: 'canceled', reason: 'user_aborted' };
            return;
        }

        const ctx = (turn.context = turn.context || {});
        if (ctx.__skillFallbackTried) {
            // 已经降级过一次仍失败 → 如实上报，不再循环
            sessionStore?.failTurn?.(turn.id, { message: reasonMsg });
            yield { type: 'error', error: { code: 'PIPELINE_SKILL_EXECUTION_FAILED', message: reasonMsg, recoverable: false } };
            return;
        }
        ctx.__skillFallbackTried = true;

        // 清掉 Skill 路由痕迹，避免通用管道又被 Skill 上下文带偏
        ctx.__forceSkillPipeline = false;
        ctx._matchedSkill = null;
        if (turn.kind !== 'text') turn.kind = 'text';

        // 清空 Skill 骨架面板（含失败步骤），避免与新回复并存造成"假失败"观感
        turn._skillSteps = [];
        try { sessionStore?.updateTurn?.(turn.id, { _skillSteps: [], _planHtml: '' }); } catch (e) { /* 渲染清理失败不阻断降级 */ }

        ChatLog.warn?.('PIPELINE', `[skill-pipeline] Skill 执行失败，静默降级到通用管道：${reasonMsg}`);

        const textPipe = (App.ChatV2?.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.pick === 'function')
            ? App.ChatV2.pipelineRegistry.pick('text')
            : App.ChatV2?.pipelines?.text;
        const textExec = textPipe && typeof textPipe.execute === 'function' ? textPipe.execute : null;
        if (!textExec) {
            sessionStore?.failTurn?.(turn.id, { message: reasonMsg });
            yield { type: 'error', error: { code: 'PIPELINE_SKILL_FALLBACK_UNAVAILABLE', message: reasonMsg, recoverable: false } };
            return;
        }

        for await (const evt of textExec(turn, signal)) {
            yield evt;
        }
    }

    async function* execute(turn, signal) {
        const sessionStore = App.ChatV2.sessionStore;
        const toolRegistry = App.ChatV2.toolRegistry;

        if (!sessionStore || !toolRegistry) {
            yield { type: 'error', error: { code: 'PIPELINE_SKILL_DEPENDENCY_MISSING', message: 'skill-pipeline 依赖未就绪' } };
            return;
        }

        const skill = turn.context?._matchedSkill;
        if (!skill || !skill.name) {
            for await (const evt of fallbackToText(turn, signal, '未找到要执行的 Skill')) yield evt;
            return;
        }

        const skillName = skill.name;
        const displayName = skill.displayName || skill.name;

        // 生命周期（Phase 4）：新任务（非后续消息）时清空上一轮执行上下文
        if (!turn.context?._execFollowUp && sessionStore.clearExecSpec) {
            sessionStore.clearExecSpec();
        }
        // 后续消息继承（Phase 5）：注入上轮 skill 结果 + 继承数据快照
        applySkillFollowUp(turn);

        ChatLog.info?.('PIPELINE', `[skill-pipeline] turn=${turn.id} 开始执行 Skill: ${skillName}`);
        yield { type: 'status', text: `正在执行 ${displayName}...` };

        // 【确定性兜底】技能声明 pipeline=html 时一律交给 html-display-pipeline；
        // 绝不落进后端 skill 执行（后者只返回"分析完成/无VBA执行阶段"空壳，不真正生成网页）。
        // 该拦截与 orchestrator 把 kind 判成 'skill' 还是 'html' 完全解耦 —— 无论从哪条路进入
        // 本 pipeline，只要技能声明 html，即由前端确定性生成 HTML 报表并自动收尾。
        {
            const htmlPipe = (window.App?.ChatV2?.pipelineRegistry
                    && typeof window.App.ChatV2.pipelineRegistry.pick === 'function'
                    && window.App.ChatV2.pipelineRegistry.pick('html'))
                || window.App?.ChatV2?.pipelines?.html;
            const declaresHtml = String(skill.pipeline || skill.Pipeline || '').toLowerCase() === 'html';
            if (declaresHtml && htmlPipe && typeof htmlPipe.execute === 'function') {
                ChatLog.info?.('PIPELINE', `[skill-pipeline] 技能 ${skill.name} 声明 pipeline=html → 转交 html-display-pipeline`);
                for await (const evt of htmlPipe.execute(turn, signal)) {
                    yield evt;
                }
                return;
            }
        }

        // 统一走声明式 executionMode 分发：
        //   - spec  → 确定性 Spec 引擎（等价 processingMode=SkillSpec，v20260818 能力）
        //   - agent → 自主 Agent 循环（ReAct 多步 + token 水位 + 步骤日志）
        //   - fixed → 传统模式（按 processingMode 走后端 SkillExecutionRouter）
        // 未声明 executionMode 时向后兼容：processingMode=SkillSpec(5) → spec。
        // 后端 GetAllSkills 将 ProcessingMode 序列化为 CInt 枚举值：
        //   VBACode=0, MultimodalLLM=1, Hybrid=2, LocalOCR=3, PythonCode=4, SkillSpec=5
        // 前端需同时兼容数字 5 与字符串 'SkillSpec'（skill-shared-loader 解析时是字符串）。
        const processingMode = skill.processingMode !== undefined && skill.processingMode !== null
            ? skill.processingMode
            : (skill.ProcessingMode !== undefined && skill.ProcessingMode !== null
                ? skill.ProcessingMode
                : 'VBACode');
        const pmStr = String(processingMode).toLowerCase();
        const executionMode = SP.util.resolveExecutionMode(skill, processingMode, pmStr);

        // ✨ 统一预规划骨架：三条执行路径共享同一套绿面板渲染
        // buildSkillSkeleton 基于 skill 元数据（subSkills / agentConfig.maxIterations / executionMode）
        // 生成全 pending 步骤；initSkillStepsPanel 立即把骨架渲染到 sessionStore._planHtml，
        // chat-renderer 监听 patch._planHtml 会把完整骨架面板插到气泡最上方。
        // 执行中 appendSkillStep 按 id/title/toolName 匹配骨架步骤 → update 填色。
        SP.stepsPanel.initSkillStepsPanel(turn, sessionStore, SP.stepsPanel.buildSkillSkeleton(skill, executionMode));

        if (executionMode === 'spec') {
            try {
                let specResult = null;
                // executeSkillSpecMode 是 async generator，需消费其事件流。
                // ⚠️ 用绝对截止时间包裹：内部 await 永挂/死循环不 yield 终态时强制终止，
                // 抛错后由下方 catch 降级为文字管道，避免步骤面板永久停在"执行中"。
                const _specGen = withSkillSpecDeadline(
                    SP.specExecute.executeSkillSpecMode(turn, skill, displayName, signal),
                    SPEC_EXEC_TIMEOUT_MS,
                    'SkillSpec 执行超时（已自动终止，请重试；若反复出现请确认插件为最新版并完全退出 Office 后重开）'
                );
                for await (const evt of _specGen) {
                    // ✨ skeleton 事件：LLM 规划完成，提前初始化多步骨架面板
                    if (evt.type === 'skeleton' && Array.isArray(evt.steps)) {
                        // 归一化为 { id, title, detail, status, strategy, subSteps }
                        const normSteps = evt.steps.map((s, i) => ({
                            id: s.id || `sk_spec_${skill.name || 'skill'}_${i}`,
                            title: s.title || s.name || '步骤 ' + (i + 1),
                            description: s.description || s.title || s.name,
                            detail: s.detail || s.description || '',
                            status: s.status || 'pending',
                            strategy: s.strategy || 'direct',
                            // ✨ 保留嵌套 subSteps（Skill 大步骤嵌 op 小步骤）
                            subSteps: Array.isArray(s.subSteps) ? s.subSteps.map((sub, j) => ({
                                id: sub.id || `sk_spec_sub_${i}_${j}`,
                                title: sub.title || sub.name || `步骤 ${j + 1}`,
                                description: sub.description || sub.title || sub.name,
                                detail: sub.detail || sub.description || '',
                                status: sub.status || 'pending',
                                strategy: sub.strategy || 'direct'
                            })) : undefined
                        }));
                        // 替换掉 initSkillStepsPanel 塞的 1 步占位骨架
                        if (App.ChatV2.ProgressTracker && typeof App.ChatV2.ProgressTracker.init === 'function') {
                            App.ChatV2.ProgressTracker.init(normSteps);
                        }
                        turn._skillSteps = normSteps;
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);

                        // ✅ 【thoughtList 对接 v20260930 — 绕过 sessionStore 单例 Bug】
                        // sessionStore.completeThought 内部强依赖 STATE.currentTurn 单例，
                        // skill-pipeline 的 turn 对象可能不是 STATE.currentTurn → completeThought 静默 return！
                        // 直接操作 turn.thoughtList + 手动 notify 绕过这个限制。
                        if (!Array.isArray(turn.thoughtList)) turn.thoughtList = [];
                        turn.thoughtList.length = 0;
                        for (let i = 0; i < normSteps.length; i++) {
                            const ns = normSteps[i];
                            const now = Date.now();
                            turn.thoughtList.push({
                                id: `skill_step_${turn.id}_${i}`,
                                round: i + 1,
                                kind: 'skill_step',
                                toolName: (ns.title || '').match(/[^\s]+/)?.[0] || '',
                                title: (ns.title || '步骤 ' + (i + 1)).slice(0, 80),
                                content: ns.detail || '',
                                status: 'pending',
                                startTs: now,
                                endTs: null
                            });
                        }
                        // 手动触发渲染（通知失败仅记录，不阻断管线）
                        try { sessionStore?.notify?.({ type: 'thought_update', turnId: turn.id }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec] notify thought_update 失败（skeleton thoughtList 注入后）:', e); }
                        ChatLog.info?.('PIPELINE', `[skill-spec] thoughtList 直接注入: ${turn.thoughtList.length} 步 (绕过 sessionStore 单例)`);

                        ChatLog.info?.('PIPELINE', `[skill-spec] skeleton 面板初始化: ${normSteps.length} 步 pending → 已渲染 (hasSub=${!!normSteps[0]?.subSteps})`);
                    }
                    // ✨ progress 事件：执行中更新骨架步骤状态（逐步填色）
                    else if (evt.type === 'progress' && typeof evt.stepIndex === 'number') {
                        const pt = App.ChatV2.ProgressTracker;
                        const updatePayload = { status: evt.status || 'pending', detail: evt.detail };
                        // ✨ 执行循环会带 subSteps（嵌套 Skill 逐步填色）
                        if (Array.isArray(evt.subSteps)) updatePayload.subSteps = evt.subSteps;
                        if (pt && typeof pt.updateStep === 'function') {
                            pt.updateStep(evt.stepIndex, updatePayload, turn, sessionStore);
                        }
                        // 同步到 turn._skillSteps（兜底）
                        if (Array.isArray(turn._skillSteps) && turn._skillSteps[evt.stepIndex]) {
                            if (evt.status) turn._skillSteps[evt.stepIndex].status = evt.status;
                            if (evt.detail) turn._skillSteps[evt.stepIndex].detail = evt.detail;
                            if (Array.isArray(evt.subSteps)) turn._skillSteps[evt.stepIndex].subSteps = evt.subSteps;
                        }
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);

                        // ✅ 【thoughtList 逐步填色 — 绕过 sessionStore 单例】
                        const _tl = Array.isArray(turn.thoughtList) ? turn.thoughtList : null;
                        const _targetIdx = evt.stepIndex;
                        // ⭐️ 防御：thoughtList 越界时动态 push（skeleton 和 planSteps 数量可能不一致）
                        if (_tl && _targetIdx >= _tl.length) {
                            while (_tl.length <= _targetIdx) {
                                _tl.push({ id: 'skill_dyn_' + _tl.length, round: _tl.length + 1, kind: 'skill_step', title: '', status: 'pending', startTs: Date.now() });
                            }
                        }
                        if (_tl && _tl[_targetIdx]) {
                            _tl[_targetIdx].status = evt.status || 'pending';
                            _tl[_targetIdx].content = evt.detail || _tl[_targetIdx].content || '';
                            if (evt.status === 'completed' || evt.status === 'failed') {
                                _tl[_targetIdx].endTs = Date.now();
                            }
                            try { sessionStore?.notify?.({ type: 'thought_update', turnId: turn.id }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec] notify thought_update 失败（thoughtList 填色）:', e); }
                        }
                    }
                    // 转发内部事件（status/token/result 等）
                    else if (evt.type === 'result') {
                        specResult = evt.result;
                    } else {
                        yield evt;
                    }
                }
                // 非异常路径兜底判据：执行"完成"但实际什么都没产出 → 同样视为失败，走降级
                if (specResult && isSkillResultEmpty(specResult)) {
                    throw new Error('SkillSpec 执行未产生有效结果');
                }
                if (specResult) {
                    // 把 spec 结果里的 steps 挂到 turn，
                    // 气泡最终化渲染（flushContentRender）用同一 renderStepsPlanHtml 置顶展示，
                    // 与合并/拆分/关联的绿面板保持一致。
                    if (Array.isArray(specResult.steps)) {
                        turn._steps = (specResult.steps || []).filter(s =>
                (typeof s === 'string' && s.trim())
                || (s && typeof s === 'object' && !!s.title));
                        // 把 spec steps 也喂给统一的绿面板管线（ProgressTracker.buildPlanHtml），
                        // 与 agent/fixed 两种模式的面板视觉统一。
                        // specResult.steps 可能来自 findStepsDeep（string[]）或 buildSpecSteps（object[]），
                        // 统一归一化为 { title, detail, status, strategy } 四元组。
                        turn._skillSteps = (specResult.steps || []).map(s => {
                            if (typeof s === 'string') {
                                const trimmed = s.trim();
                                if (!trimmed) return null;
                                // "设置主题 · 成功（120ms）" → title="设置主题", detail="成功（120ms）"
                                const sepIdx = trimmed.search(/[·:：]/);
                                if (sepIdx > 0) {
                                    return {
                                        title: trimmed.slice(0, sepIdx).trim(),
                                        description: trimmed.slice(0, sepIdx).trim(),
                                        detail: trimmed.slice(sepIdx + 1).trim(),
                                        status: 'completed',
                                        strategy: 'direct'
                                    };
                                }
                                return { title: trimmed, description: trimmed, detail: '', status: 'completed', strategy: 'direct' };
                            }
                            if (s && typeof s === 'object') {
                                const title = s.title || s.name;
                                if (!title) return null;
                                return {
                                    title: String(title),
                                    description: s.description || String(title),
                                    detail: s.detail || s.description || '',
                                    status: s.status || 'completed',
                                    strategy: s.strategy || 'direct',
                                    // ✨ 保留真实工具调用信息（否则展开区只显示空的步骤名）
                                    toolName: s.toolName || null,
                                    params: s.params || null,
                                    result: s.result || null,
                                    // ✨ 保留嵌套 subSteps
                                    subSteps: Array.isArray(s.subSteps) ? s.subSteps.map((sub, j) => ({
                                        id: sub.id || `res_sub_${j}`,
                                        title: sub.title || sub.name || `步骤 ${j + 1}`,
                                        description: sub.description || sub.title || sub.name,
                                        detail: sub.detail || sub.description || '',
                                        status: sub.status || 'completed',
                                        strategy: sub.strategy || 'direct'
                                    })) : undefined
                                };
                            }
                            return null;
                        }).filter(Boolean);
                        // DEBUG: trace what got into turn._skillSteps
                        ChatLog.info?.('PIPELINE', `[skill-spec] execute() 消费结果: specResult.steps=${specResult.steps.length} → turn._skillSteps=${turn._skillSteps.length}`);
                        // ✨【终态兜底】progress 的 stepIndex 基于 generator 内部 earlySkeleton 编号，
                        // 而 UI 面板/thoughtList 可能含动态插入的真实工具步骤（如"识别表头行数"），
                        // 二者索引错位时，收尾的 completed 事件覆盖不到最后一条 → 永远卡在"正在..."。
                        // specResult 已到达 = 整轮执行成功，这里统一把所有未终态步骤翻转为 completed。
                        if (Array.isArray(turn._skillSteps)) {
                            turn._skillSteps.forEach(s => {
                                if (s && s.status && s.status !== 'completed' && s.status !== 'failed' && s.status !== 'done') {
                                    s.status = 'completed';
                                }
                                if (s && typeof s.detail === 'string' && /^正在应用到/.test(s.detail.trim())) {
                                    s.detail = '已应用到工作表';
                                }
                            });
                        }
                        if (Array.isArray(turn.thoughtList)) {
                            const _finalizeTs = Date.now();
                            turn.thoughtList.forEach(t => {
                                if (t && t.status !== 'completed' && t.status !== 'failed') {
                                    t.status = 'completed';
                                    if (!t.endTs) t.endTs = _finalizeTs;
                                    if (typeof t.content === 'string' && /^正在应用到/.test(t.content.trim())) {
                                        t.content = '已应用到工作表';
                                    }
                                }
                            });
                            try { sessionStore?.notify?.({ type: 'thought_update', turnId: turn.id }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec] notify thought_update 失败（thoughtList 终态收尾）:', e); }
                        }
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                    }
                    // 操作类技能：步骤已由绿面板（ProgressTracker + turn._skillSteps）渲染，
                    // 但气泡内容仍需保留 specResult.outputText（执行结果摘要）——面板是结构化视觉，
                    // 气泡内容是简短摘要文字，两者互补而非互斥。
                    // 旧逻辑 hasSpecSteps ? '' : outputText 会把气泡内容清空 → contentLen=0 → 用户看不到任何文字回复。
                    const specOutput = specResult.outputText || '';
                    sessionStore.completeTurn(turn.id, { content: specOutput });
                    writeSkillExecSpec(turn, sessionStore, {
                        executionMode: 'spec',
                        spec: specResult.spec || null,
                        outputText: specOutput || specResult.outputText || '',
                        steps: specResult.steps || null,
                        // 持久化图表数据上下文：数据源范围 + 聚合快照，供后续"换成柱状图/折线图"复用
                        dataRange: specResult.dataRange || null,
                        aggregate: specResult.aggregate || null
                    });
                    if (specOutput) {
                        yield { type: 'content', delta: specOutput };
                    }
                    yield { type: 'pipeline_done', pipeline: 'skill' };
                    return;
                }
                // specResult 为空 → 抛错（走下方 catch）
                throw new Error('SkillSpec 未产生结果');
            } catch (specErr) {
                // Spec 模式失败时的降级策略：
                //   - 原始 processingMode 是 SkillSpec(5) → 不降级（传统 executeSkill 不识别 SkillSpec）
                //   - 原始 processingMode 不是 SkillSpec（如 VBACode=0 被按名称强制进 spec）→ 降级到 fixed 执行
                const specMsg = specErr?.message || '未知错误';
                ChatLog.error?.('PIPELINE', `[skill-pipeline] SkillSpec 模式执行失败 (原 processingMode=${processingMode}):`, specErr);

                // ✨ 终态收尾：Spec 失败/超时后，把仍处 running/pending 的步骤与 thoughtList 翻成 failed，
                // 否则降级到文字管道前步骤面板会残留"正在执行中"假象（尤其 type=text 选区误路由导致超时）。
                if (Array.isArray(turn._skillSteps)) {
                    turn._skillSteps.forEach(function (s) {
                        if (s && s.status && s.status !== 'completed' && s.status !== 'failed' && s.status !== 'done') {
                            s.status = 'failed';
                        }
                    });
                    SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                }
                if (Array.isArray(turn.thoughtList)) {
                    const _ft = Date.now();
                    turn.thoughtList.forEach(function (t) {
                        if (t && t.status !== 'completed' && t.status !== 'failed') {
                            t.status = 'failed';
                            if (!t.endTs) t.endTs = _ft;
                        }
                    });
                    try { sessionStore?.notify?.({ type: 'thought_update', turnId: turn.id }); } catch (e) { ChatLog.warn?.('PIPELINE', '[skill-spec] notify thought_update 失败（thoughtList 失败收尾）:', e); }
                }

                const wasForced = pmStr !== 'skillspec' && Number(processingMode) !== 5;
                if (wasForced) {
                    ChatLog.warn?.('PIPELINE', `[skill-pipeline] 降级到 fixed 模式重试 ${skillName}`);
                    // 降级后 executionMode 变 fixed，重新进 fixed 路径（复用下方 invokeTracked 逻辑）
                    // 把 turn._skillSteps 重置为 fixed 骨架
                    const subSkills = skill.subSkills || [];
                    const fixedSkeleton = [{
                        title: displayName, detail: `${displayName} 主执行`,
                        status: 'pending', strategy: 'pipeline',
                        toolName: 'skill.' + skillName, id: `sk_${skillName}_main`
                    }, ...subSkills.map((s, i) => ({
                        title: typeof s === 'string' ? s : (s.displayName || s.name || '子 Skill'),
                        detail: `子流程 ${i+1}/${subSkills.length}`,
                        status: 'pending', strategy: 'pipeline',
                        toolName: 'skill.' + (typeof s === 'string' ? s : (s.name || '')),
                        id: `sk_${skillName}_sub_${i}`
                    }))];
                    turn._skillSteps = fixedSkeleton;
                    sessionStore?.updateTurn?.(turn.id, { _planHtml: '' });
                    SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                    // 不 return，继续走下方 fixed 路径
                } else {
                    // 原始就是 SkillSpec 且失败 → 无法在 Skill 内自愈，降级到通用管道
                    for await (const evt of fallbackToText(turn, signal, `${displayName} 执行失败：${specMsg}`)) yield evt;
                    return;
                }
            }
        }

        if (executionMode === 'steps') {
            try {
                let stepsResult = null;
                // 守卫前置：stepsExecute 未挂载时先报清晰错误，避免在调用点 TypeError
                if (!SP.stepsExecute || typeof SP.stepsExecute.executeSkillStepsMode !== 'function') {
                    throw new Error('skill-steps-execute 模块未加载，无法执行声明式步骤（请确认 index.html 引入 skill-steps-execute.js 且在 skill-pipeline.js 之前）');
                }
                // executeSkillStepsMode 是 async generator，消费其事件流（契约与 spec 分支同构）
                for await (const evt of SP.stepsExecute.executeSkillStepsMode(turn, skill, displayName, signal)) {
                    // skeleton：声明式步骤已由执行器产出全 pending 骨架，直接初始化绿面板
                    if (evt.type === 'skeleton' && Array.isArray(evt.steps)) {
                        const normSteps = evt.steps.map((s, i) => ({
                            id: s.id || `sk_steps_${skill.name || 'skill'}_${i}`,
                            title: s.title || s.name || '步骤 ' + (i + 1),
                            description: s.description || s.title || s.name,
                            detail: s.detail || s.description || '',
                            status: s.status || 'pending',
                            strategy: s.strategy || 'direct',
                            toolName: s.toolName || s.tool || null
                        }));
                        if (App.ChatV2.ProgressTracker && typeof App.ChatV2.ProgressTracker.init === 'function') {
                            App.ChatV2.ProgressTracker.init(normSteps);
                        }
                        turn._skillSteps = normSteps;
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                        ChatLog.info?.('PIPELINE', `[skill-steps] skeleton 面板初始化: ${normSteps.length} 步 pending`);
                    }
                    // progress：执行中逐步填色
                    else if (evt.type === 'progress' && typeof evt.stepIndex === 'number') {
                        const pt = App.ChatV2.ProgressTracker;
                        if (pt && typeof pt.updateStep === 'function') {
                            pt.updateStep(evt.stepIndex, { status: evt.status || 'pending', detail: evt.detail }, turn, sessionStore);
                        }
                        if (Array.isArray(turn._skillSteps) && turn._skillSteps[evt.stepIndex]) {
                            if (evt.status) turn._skillSteps[evt.stepIndex].status = evt.status;
                            if (evt.detail) turn._skillSteps[evt.stepIndex].detail = evt.detail;
                        }
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                    }
                    // result：执行完成，收尾（面板渲染 + 气泡摘要 + 结果快照继承）
                    else if (evt.type === 'result') {
                        stepsResult = evt.result;
                    } else {
                        yield evt;
                    }
                }
                if (stepsResult) {
                    if (Array.isArray(stepsResult.steps)) {
                        turn._skillSteps = stepsResult.steps.map(s => {
                            if (!s || typeof s !== 'object') return null;
                            const title = s.title || s.name || s.tool;
                            if (!title) return null;
                            return {
                                title: String(title),
                                description: s.description || String(title),
                                detail: s.detail || s.tool || '',
                                status: s.status || 'completed',
                                strategy: s.strategy || 'direct',
                                toolName: s.toolName || s.tool || null,
                                params: s.params || null,
                                result: s.result || null
                            };
                        }).filter(Boolean);
                        SP.stepsPanel.renderSkillSteps(turn, sessionStore);
                    }
                    const stepsOutput = stepsResult.outputText || '';
                    sessionStore.completeTurn(turn.id, { content: stepsOutput });
                    writeSkillExecSpec(turn, sessionStore, {
                        executionMode: 'steps',
                        steps: stepsResult.steps || null,
                        outputText: stepsOutput || ''
                    });
                    if (stepsOutput) {
                        yield { type: 'content', delta: stepsOutput };
                    }
                    yield { type: 'pipeline_done', pipeline: 'skill' };
                    return;
                }
                throw new Error('声明式 steps 未产生结果');
            } catch (stepsErr) {
                const stepsMsg = stepsErr?.message || '未知错误';
                ChatLog.error?.('PIPELINE', `[skill-steps] 声明式 steps 模式执行失败 (原 processingMode=${processingMode}):`, stepsErr);
                for await (const evt of fallbackToText(turn, signal, `${displayName} 执行失败：${stepsMsg}`)) yield evt;
                return;
            }
        }

        if (executionMode === 'agent') {
            try {
                // executeSkillAgentMode 内部已 completeTurn；此处只转发事件并收尾
                for await (const evt of executeSkillAgentMode(turn, skill, displayName, signal)) {
                    yield evt;
                }
                const finalTurn = sessionStore.getCurrentTurn();
                writeSkillExecSpec(turn, sessionStore, {
                    executionMode: 'agent',
                    outputText: (finalTurn && finalTurn.id === turn.id) ? (finalTurn.content || '') : '',
                    steps: Array.isArray(turn._steps) ? turn._steps : null
                });
                yield { type: 'pipeline_done', pipeline: 'skill' };
                return;
            } catch (agentErr) {
                const agentMsg = agentErr?.message || '未知错误';
                ChatLog.error?.('PIPELINE', `[skill-pipeline] Agent 模式执行失败:`, agentErr);
                for await (const evt of fallbackToText(turn, signal, `${displayName} 执行失败：${agentMsg}`)) yield evt;
                return;
            }
        }

        try {
            await toolRegistry.init();

            const args = buildSkillArgs(turn);

            // 骨架已由 execute() 入口统一 init，此处直接进入执行

            // 使用通用 invokeTracked 执行 Skill——自动显示工具卡片，UI 渲染完全通用
            const result = await toolRegistry.invokeTracked(turn, 'skill.' + skillName, args, {
                round: 1,
                displayName: displayName,
                displayArgs: buildDisplayArgs(turn, skill),
                ctx: { turn, sessionStore, signal }
            });

            if (signal?.aborted) {
                yield { type: 'canceled', reason: 'user_aborted' };
                return;
            }

            // 主 Skill 结果写入步骤面板
            SP.stepsPanel.appendSkillStep(turn, sessionStore, {
                title: displayName,
                detail: (result?.ok ? '执行成功' : '执行失败：' + (result?.error || '')),
                status: result?.ok ? 'completed' : 'failed',
                strategy: 'pipeline',
                toolName: 'skill.' + skillName,
                result: result?.result || null
            });
            if (!result?.ok) {
                const errMsg = result?.error || 'Skill 执行失败';
                ChatLog.error?.('PIPELINE', `[skill-pipeline] Skill ${skillName} 执行返回失败:`, result);
                for await (const evt of fallbackToText(turn, signal, errMsg)) yield evt;
                return;
            }

            // ── 检查是否有 subSkills 需要顺序执行 ──────────────────────────
            const subSkills = skill.subSkills || [];
            let accumulatedResult = result.result;
            let subSkillOutputs = [];

            if (subSkills.length > 0) {
                ChatLog.info?.('PIPELINE', `[skill-pipeline] ${skillName} 声明了 ${subSkills.length} 个子 Skill，开始顺序执行`);
                yield { type: 'status', text: `${displayName} 分析完成，开始执行 ${subSkills.length} 个子流程...` };

                for (let i = 0; i < subSkills.length; i++) {
                    if (signal?.aborted) {
                        yield { type: 'canceled', reason: 'user_aborted' };
                        return;
                    }

                    const subSkillName = subSkills[i];
                    const subDisplayName = subSkillName;
                    const round = i + 2; // 主 Skill 占 round=1，子 Skill 从 round=2 开始

                    yield { type: 'status', text: `正在执行子流程 ${i + 1}/${subSkills.length}: ${subDisplayName}...` };

                    // 构建子 Skill 参数：继承主 Skill 上下文 + 父 Skill 结果
                    const subArgs = buildSubSkillArgs(turn, accumulatedResult, subSkillOutputs);

                    const subResult = await toolRegistry.invokeTracked(turn, 'skill.' + subSkillName, subArgs, {
                        round: round,
                        displayName: subDisplayName,
                        displayArgs: { source: skillName, parentResult: '已继承' },
                        ctx: { turn, sessionStore, signal }
                    });

                    if (signal?.aborted) {
                        yield { type: 'canceled', reason: 'user_aborted' };
                        return;
                    }

                    if (subResult?.ok) {
                        const subOutput = formatSkillResult(subResult.result, subDisplayName);
                        subSkillOutputs.push({ name: subSkillName, output: subOutput });
                        // 合并子 Skill 结果到累计上下文，供后续子 Skill 使用
                        accumulatedResult = mergeResults(accumulatedResult, subResult.result);
                        ChatLog.info?.('PIPELINE', `[skill-pipeline] 子 Skill ${subSkillName} 执行成功`);
                        // 子 Skill 成功 → 写入步骤面板
                        SP.stepsPanel.appendSkillStep(turn, sessionStore, {
                            title: subDisplayName,
                            detail: `子流程 ${i + 1}/${subSkills.length} · 执行成功`,
                            status: 'completed',
                            strategy: 'pipeline',
                            toolName: 'skill.' + subSkillName,
                            result: subResult?.result || null
                        });
                    } else {
                        const errMsg = subResult?.error || `${subSkillName} 执行失败`;
                        ChatLog.warn?.('PIPELINE', `[skill-pipeline] 子 Skill ${subSkillName} 执行失败:`, errMsg);
                        // 子 Skill 失败不影响主流程，记录警告继续执行
                        subSkillOutputs.push({ name: subSkillName, error: errMsg });
                        yield { type: 'status', text: `子流程 ${subSkillName} 执行失败: ${errMsg}` };
                        // 子 Skill 失败 → 步骤面板也要记录（绿面板应能看到失败）
                        SP.stepsPanel.appendSkillStep(turn, sessionStore, {
                            title: subDisplayName,
                            detail: `子流程 ${i + 1}/${subSkills.length} · 失败：${errMsg}`,
                            status: 'failed',
                            strategy: 'pipeline',
                            toolName: 'skill.' + subSkillName,
                            result: null
                        });
                    }
                }
            }

            // ── 组装最终输出 ──────────────────────────────────────────────
            // 操作类技能（合并/拆分/关联/清洗/美化/图表）的结果已由工具卡片"常显区"
            // 统一渲染步骤规划绿面板 + 摘要，此处不再重复生成消息气泡，避免同一份信息出现两次。
            // 例外：SkillCreator 生成的自定义 Skill 结果（skill_card）**必须**显式展示——
            // 用户要看生成的 Skill 本体（名称/描述/分类 + 保存/编辑/复制）。findStepsDeep 会把
            // 后端返回壳的 message 摘要误判为 steps，导致 hasSteps=true 而把 skill_card 吞掉，
            // 最终气泡只剩"步骤规划"面板、看不到生成的 Skill。
            const candidateOut = buildFinalOutput(accumulatedResult, displayName, subSkills, subSkillOutputs);
            const isSkillCardText = typeof candidateOut === 'string'
                && /"type"\s*:\s*"skill_card"/.test(candidateOut);
            const hasSteps = SP.util.findStepsDeep(accumulatedResult).length > 0;
            let outputText = (hasSteps && !isSkillCardText) ? '' : candidateOut;

            // ── 可保存脚本卡片（uiActions 含 showCopy/showSaveMd/showRegenerate，如 GenerateOCRParseScript） ──
            // 这类技能产出的是「可保存的解析脚本」，即使后端带了 steps 也不要被 hasSteps 清空。
            // 脚本内容挂到 turn._skillResultCard，由 chat-renderer 在气泡底部「继续操作」区渲染专门的盒子
            // （复用 App.ChatV2._buildSaveableScriptCard）+ 复制/保存/重新生成 按钮，取代与技能无关的通用 chips；
            // 执行详情里仍保留原始 JSON 供排查，正文仅留简短成功提示，避免同一份脚本在两处重复展示。
            const _cached = (App.cachedSkills || []).find(function (s) { return s.name === skillName || s.Name === skillName; });
            const _sk = _cached || skill || null;
            const _ui = normalizeUiActions(_sk) || normalizeUiActions(skill);
            const saveableScript = _ui ? extractSaveableScript(accumulatedResult) : null;
            if (saveableScript) {
                outputText = '✅ 已生成解析脚本「' + _extractScriptName(saveableScript, ((_sk && (_sk.displayName || _sk.DisplayName)) || skillName || '解析脚本')) + '」，可在下方查看、复制、保存或重新生成。';
                // 把可保存脚本挂到 turn，供 chat-renderer 在底部「继续操作」区渲染专门的盒子 + 操作按钮
                // （取代与技能无关的通用 chips，如"调整图表/做成PPT"）。render 时复用 _buildSaveableScriptCard。
                turn._skillResultCard = {
                    script: saveableScript,
                    skillName: ((_sk && (_sk.name || _sk.Name)) || skillName || ''),
                    displayName: ((_sk && (_sk.displayName || _sk.DisplayName)) || skillName || '解析脚本'),
                    prompt: (turn && turn.userMessage) || '',
                    uiActions: _ui
                };
                ChatLog.info?.('PIPELINE', '[skill-pipeline] 可保存脚本已挂到 turn._skillResultCard（' + skillName + '）');
            } else {
                ChatLog.info?.('PIPELINE', '[skill-pipeline] 未生成可保存脚本卡：uiActions=' +
                    JSON.stringify(_ui) + ', scriptFound=' + (extractSaveableScript(accumulatedResult) != null));
            }

            // ── 自动导入提示（仅当技能仍声明 autoImportParseScript 且后端真的导入时） ──
            const parseScriptImportNotice = buildParseScriptImportNotice(accumulatedResult);
            if (parseScriptImportNotice) {
                yield { type: 'status', text: parseScriptImportNotice };
            }

            if (outputText) {
                sessionStore.appendDelta(turn.id, 'content', outputText);
                yield { type: 'content', delta: outputText };
            }

            sessionStore.completeTurn(turn.id, {
                content: outputText
            });
            writeSkillExecSpec(turn, sessionStore, {
                executionMode: 'fixed',
                outputText: outputText,
                steps: SP.util.findStepsDeep(accumulatedResult)
            });
            // 最终收尾：刷新绿面板到最终状态 + 双写 turn._steps
            SP.stepsPanel.finalizeSkillSteps(turn, sessionStore);
            yield { type: 'pipeline_done', pipeline: 'skill' };

        } catch (err) {
            ChatLog.error?.('PIPELINE', `[skill-pipeline] Skill ${skillName} 执行异常:`, err);
            const msg = err?.message || String(err);
            for await (const evt of fallbackToText(turn, signal, msg)) yield evt;
        }
    }

    /**
     * 把 turn 上下文组装成 Skill 可消费的参数
     */
    function buildSkillArgs(turn) {
        const ctx = turn.context || {};
        const args = {
            prompt: turn.userMessage || '',
            text: turn.userMessage || '',
            message: turn.userMessage || ''
        };

        if (ctx.selection) {
            args.selection = ctx.selection;
            args.excelData = JSON.stringify(ctx.selection, null, 2);
        }

        if (Array.isArray(ctx.referenceRegions) && ctx.referenceRegions.length > 0) {
            args.referenceRegions = ctx.referenceRegions;
        }

        if (Array.isArray(ctx.attachments) && ctx.attachments.length > 0) {
            args.attachments = ctx.attachments;
        }

        return args;
    }

    /**
     * 构建 UI 展示用的参数摘要（不传入实际大数据，避免参数面板撑爆）
     */
    function buildDisplayArgs(turn, skill) {
        const ctx = turn.context || {};
        const display = {};
        if (ctx.selection) {
            const sel = ctx.selection;
            display.selection = sel.address || sel.range || '已选中区域';
            if (sel.rowCount || sel.colCount) {
                display.selection += ` (${sel.rowCount || '?'}行 × ${sel.colCount || '?'}列)`;
            }
        }
        if (Array.isArray(ctx.referenceRegions) && ctx.referenceRegions.length > 0) {
            display.referenceRegions = `${ctx.referenceRegions.length} 个引用区域`;
        }
        if (Array.isArray(ctx.attachments) && ctx.attachments.length > 0) {
            display.attachments = `${ctx.attachments.length} 个附件`;
        }
        if (Object.keys(display).length === 0) {
            display.input = turn.userMessage?.substring(0, 100) || '';
        }
        return display;
    }

    /**
     * 格式化 Skill 执行结果用于展示
     * 对于 skill_card 结构化 JSON：透传，并把 generatedSkillJson 合入 skillCardJson 字段，供渲染器使用
     * 其他情况：返回文本/HTML，由渲染器继续处理
     */
    function formatSkillResult(result, displayName) {
        if (!result) return `${displayName} 执行完成。`;

        if (typeof result === 'string') {
            // 【JSON 回声守卫】后端工具的 content 本身就是 payload JSON 字符串
            // （如 excel_merge_sheets 的 {success,sourceSheets,mergedRows,data,steps}），
            // 原样返回会被当作技能总结写进气泡（2026-10-09 用户反馈"合并成功后返回的内容
            // 被渲染成了 json"）。命中已知工具结果结构时替换为可读摘要；纯文本不动。
            try {
                const _fx = App.ChatV2?.toolRegistryUtil?.replaceJsonEchoWithFriendly?.(result);
                if (_fx) return _fx;
            } catch (e) { /* 守卫失败不影响原路径 */ }
            return result;
        }

        const outs = result.outputs || result.Outputs || {};
        if (!outs || typeof outs !== 'object') {
            if (result.message && typeof result.message === 'string') return result.message;
            if (result.content && typeof result.content === 'string') return result.content;
            return `${displayName} 执行完成。`;
        }

        const contentText = outs.content;
        const skillJson = outs.generatedSkillJson;

        // 如果 content 是 skill_card JSON，把 generatedSkillJson 合入 skillCardJson 字段，让渲染器统一处理
        if (typeof contentText === 'string') {
            const trimmed = contentText.trim();
            if (trimmed.startsWith('{') && trimmed.includes('"type"') && trimmed.includes('skill_card')) {
                try {
                    const card = JSON.parse(trimmed);
                    if (card && card.type === 'skill_card') {
                        if (skillJson && typeof skillJson === 'string') {
                            card.skillCardJson = skillJson;
                        }
                        return JSON.stringify(card);
                    }
                } catch (e) {
                    // 不是合法 JSON，按普通文本处理
                }
            }
        }

        // 如果 content 是普通 HTML（后端直接生成的带按钮卡片），直接透传
        if (typeof contentText === 'string' && /<[a-z][\s\S]*>/i.test(contentText)) {
            return contentText;
        }

        // 其他情况优先 message / content
        if (outs.message && typeof outs.message === 'string') return outs.message;
        if (typeof contentText === 'string') return contentText;

        if (result.message && typeof result.message === 'string') return result.message;
        if (result.content && typeof result.content === 'string') return result.content;

        return `${displayName} 执行完成。`;
    }

    /**
     * 构建子 Skill 调用参数：继承主 Skill 的原始上下文 + 累计结果
     */
    function buildSubSkillArgs(turn, parentResult, previousOutputs) {
        const ctx = turn.context || {};
        const args = {
            prompt: turn.userMessage || '',
            text: turn.userMessage || '',
            message: turn.userMessage || '',
            parentResult: parentResult,           // 父 Skill 累计结果
            previousOutputs: previousOutputs       // 已完成子 Skill 的输出列表
        };

        // 继承原始上下文
        if (ctx.selection) {
            args.selection = ctx.selection;
            args.excelData = JSON.stringify(ctx.selection, null, 2);
        }
        if (Array.isArray(ctx.referenceRegions) && ctx.referenceRegions.length > 0) {
            args.referenceRegions = ctx.referenceRegions;
        }
        if (Array.isArray(ctx.attachments) && ctx.attachments.length > 0) {
            args.attachments = ctx.attachments;
        }

        return args;
    }

    /**
     * 合并两个结果对象（浅合并，子 Skill 结果优先）
     */
    function mergeResults(base, additional) {
        if (!additional) return base;
        if (!base) return additional;
        if (typeof base !== 'object' || typeof additional !== 'object') return additional;

        return { ...base, ...additional };
    }

    /**
     * Skill 输出被自动导入到其他面板时的提示文案。
     * 目前用于 GenerateOCRParseScript：生成的解析脚本会被后端自动写入 OCR 智能解析面板，
     * 结果放在 outputs.parseScriptImport（success/scriptName/replaced/error）。
     */
    function buildParseScriptImportNotice(result) {
        try {
            const outputs = result && (result.outputs || result.Outputs);
            const info = outputs && (outputs.parseScriptImport || outputs.ParseScriptImport);
            if (!info) return '';
            if (info.success) {
                const name = info.scriptName || '解析脚本';
                return info.replaced
                    ? `✅ 已自动更新 OCR 智能解析面板中的同名脚本「${name}」`
                    : `✅ 已自动导入 OCR 智能解析面板：「${name}」（可直接在面板中使用）`;
            }
            return `⚠️ 自动导入 OCR 解析脚本未完成：${info.error || '未知原因'}`;
        } catch (e) {
            return '';
        }
    }

    // ── 可保存脚本结果卡片（技能声明 uiActions.showSaveMd: true，如 GenerateOCRParseScript） ──
    let _ocrScriptSeq = 0;
    App.__ocrScriptStore = App.__ocrScriptStore || {};

    function _escapeHtmlForCard(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    function _extractScriptName(mdText, fallback) {
        const m = String(mdText || '').match(/^\s*name:\s*([^\n]+)/m)
            || String(mdText || '').match(/^\s*id:\s*([^\n]+)/m);
        const name = m ? String(m[1]).trim().replace(/^["']|["']$/g, '') : '';
        return name || fallback || '解析脚本';
    }

    // 提取可保存的脚本正文（outputs.content）；无内容返回 null
    function extractSaveableScript(result) {
        const outs = result && (result.outputs || result.Outputs);
        const content = outs && (outs.content || outs.Content);
        if (typeof content === 'string' && content.trim()) return content.trim();
        if (typeof result === 'string' && result.trim()) return result.trim();
        return null;
    }

    /**
     * 归一化 uiActions：兼容后端序列化的大小写差异
     * （外键 uiActions/UIActions，内键 showCopy/showSaveMd/showRegenerate 或 PascalCase）。
     * 三个开关全为假或缺省时返回 null。
     */
    function normalizeUiActions(skill) {
        if (!skill) return null;
        const ui = skill.uiActions || skill.UIActions;
        if (!ui) return null;
        const pick = function (a, b) { return (a === undefined || a === null) ? b : a; };
        const norm = {
            showCopy: !!pick(ui.showCopy, ui.ShowCopy),
            showSaveMd: !!pick(ui.showSaveMd, ui.ShowSaveMd),
            showRegenerate: !!pick(ui.showRegenerate, ui.ShowRegenerate)
        };
        if (!norm.showCopy && !norm.showSaveMd && !norm.showRegenerate) return null;
        return norm;
    }

    // 把可保存脚本渲染成友好卡片：预览 + 复制/保存/重新生成（按 skill.uiActions 动态组装按钮）
    function buildSaveableScriptCard(scriptText, skill, result, prompt) {
        const ui = normalizeUiActions(skill) || {};
        const displayName = (skill && (skill.displayName || skill.name)) || '解析脚本';
        const name = _extractScriptName(scriptText, displayName);
        const token = 'ocr' + (++_ocrScriptSeq);
        // 连同重跑所需信息一起存：复制/保存取 text，重新生成取 prompt+skillName
        App.__ocrScriptStore[token] = {
            text: scriptText,
            prompt: prompt || (typeof turn !== 'undefined' && turn && turn.userMessage) || '',
            skillName: skill && (skill.name || skill.Name) || ''
        };
        const escName = _escapeHtmlForCard(name);
        const escaped = _escapeHtmlForCard(scriptText);

        let actions = '';
        if (ui.showCopy) {
            actions += '<button type="button" class="code-action-btn copy-btn" onclick="App.copyGeneratedMd(\'' + token + '\')" ' +
                'style="background:#2d6cdf;color:#fff;border:none;border-radius:4px;padding:6px 12px;cursor:pointer;font-size:12px;margin-right:6px;">📋 复制</button>';
        }
        if (ui.showSaveMd) {
            actions += '<button type="button" class="code-action-btn save-btn" onclick="App.saveGeneratedMd(\'' + token + '\')" ' +
                'style="background:#1f9d55;color:#fff;border:none;border-radius:4px;padding:6px 12px;cursor:pointer;font-size:12px;margin-right:6px;">💾 保存为 .md</button>';
        }
        if (ui.showRegenerate) {
            actions += '<button type="button" class="code-action-btn regen-btn" onclick="App.regenerateSkill(\'' + token + '\')" ' +
                'style="background:#6c5ce7;color:#fff;border:none;border-radius:4px;padding:6px 12px;cursor:pointer;font-size:12px;">🔄 重新生成</button>';
        }

        return (
            '<div class="ocr-script-card" style="border:1px solid rgba(45,108,223,.35);border-radius:8px;padding:12px 14px;margin:10px 0;background:rgba(13,27,42,.6);">' +
                '<div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">' +
                    '<span style="font-size:16px;">📄</span>' +
                    '<strong style="color:#e6edf3;font-size:14px;">已生成解析脚本：' + escName + '</strong>' +
                '</div>' +
                (actions ? '<div style="margin-bottom:10px;">' + actions + '</div>' : '') +
                '<div style="max-height:320px;overflow:auto;border:1px solid #1f2d3d;border-radius:6px;background:#0a1622;">' +
                    '<div style="margin:0;padding:10px;white-space:pre-wrap;word-break:break-word;font-size:12px;line-height:1.5;color:#c9d1d9;font-family:Consolas,Menlo,monospace;">' + escaped + '</div>' +
                '</div>' +
            '</div>'
        );
    }

    // 全局处理函数：卡片按钮 onclick 调用
    App.saveGeneratedMd = function (token) {
        try {
            const item = (App.__ocrScriptStore || {})[token];
            const md = item && (item.text != null ? item.text : item);
            if (md == null) { if (App.showToast) App.showToast('未找到脚本内容', 'error'); return; }
            const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'ocr-parse-script.md';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
            if (App.showToast) App.showToast('已保存为 ocr-parse-script.md', 'success');
        } catch (e) {
            if (App.showToast) App.showToast('保存失败: ' + (e && e.message || e), 'error');
        }
    };

    App.copyGeneratedMd = function (token) {
        try {
            const item = (App.__ocrScriptStore || {})[token];
            const md = item && (item.text != null ? item.text : item);
            if (md == null) { if (App.showToast) App.showToast('未找到脚本内容', 'error'); return; }
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(md).then(function () {
                    if (App.showToast) App.showToast('已复制到剪贴板', 'success');
                });
            } else {
                const ta = document.createElement('textarea');
                ta.value = md; document.body.appendChild(ta); ta.select();
                document.execCommand('copy'); document.body.removeChild(ta);
                if (App.showToast) App.showToast('已复制到剪贴板', 'success');
            }
        } catch (e) {
            if (App.showToast) App.showToast('复制失败: ' + (e && e.message || e), 'error');
        }
    };

    App.regenerateSkill = function (token) {
        try {
            const item = (App.__ocrScriptStore || {})[token];
            const skillName = item && (item.skillName || '');
            const prompt = item && (item.prompt || '');
            if (!skillName || !prompt) { if (App.showToast) App.showToast('缺少重跑信息，请重新发起', 'error'); return; }
            const orch = App.ChatV2 && App.ChatV2.orchestrator;
            if (!orch || typeof orch.dispatch !== 'function') { if (App.showToast) App.showToast('无法重新生成', 'error'); return; }
            if (App.showToast) App.showToast('正在重新生成…', 'info');
            orch.dispatch({
                message: prompt,
                context: { _forcedSkillName: skillName, _forcedSkillSource: 'regenerate' }
            });
        } catch (e) {
            if (App.showToast) App.showToast('重新生成失败: ' + (e && e.message || e), 'error');
        }
    };

    /**
     * 组装最终输出：有子 Skill 时汇总输出，无子 Skill 时返回父 Skill 结果
     */
    function buildFinalOutput(accumulatedResult, displayName, subSkills, subSkillOutputs) {
        // 无子 Skill 时，直接返回父 Skill 结果
        if (!subSkills || subSkills.length === 0) {
            return formatSkillResult(accumulatedResult, displayName);
        }

        // 有子 Skill 时，构建汇总输出
        const parts = [];

        // 父 Skill 结果
        const parentText = formatSkillResult(accumulatedResult, displayName);
        if (parentText) {
            parts.push(`## ${displayName}\n\n${parentText}`);
        }

        // 各子 Skill 结果
        for (const output of subSkillOutputs) {
            if (output.error) {
                parts.push(`### ${output.name}（失败）\n\n${output.error}`);
            } else if (output.output) {
                parts.push(`### ${output.name}\n\n${output.output}`);
            }
        }

        return parts.join('\n\n---\n\n');
    }



    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('skill', {
            execute,
            name: 'Skill 执行',
            description: '通用 Skill 执行（VB 后端路由 / subSkills 顺序编排）',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.skill = { execute };
    }
    // 暴露「可保存脚本友好卡」构造器：由 chat-renderer 在工具卡"返回结果"处渲染，取代原始 JSON 转储。
    // 无论走哪条注册分支都必须挂载，否则工具卡仍会回退到原始 JSON。
    App.ChatV2._buildSaveableScriptCard = buildSaveableScriptCard;
    // 供 chat-renderer 判断/渲染 uiActions 时复用同一套大小写兼容逻辑
    App.ChatV2._normalizeUiActions = normalizeUiActions;
    // 供 chat-renderer 的"最终结果"区渲染器从工具结果里取 outputs.content（Markdown 脚本）
    App.ChatV2._extractSaveableScript = extractSaveableScript;
    ChatLog.info?.('CHAT-V2', '✅ skill-pipeline loaded');
})();