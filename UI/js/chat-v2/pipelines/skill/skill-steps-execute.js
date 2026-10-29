/**
 * ============================================================
 * skill-steps-execute.js — Skill 管道 steps 模式执行器（声明式工具序列）
 * ------------------------------------------------------------
 * 职责：executeSkillStepsMode——读取 Skill 声明式 steps（由 SkillCreator 生成的
 *       默认产物），按 STEPS_CONTRACT 归一化/校验，逐条分派到真实执行通道：
 *         - strategy='direct' → toolRegistry.invokeTracked（MCP 工具 / skill 委托）
 *         - strategy='code'   → 显式 VBA（优先 excelBridge.ExecuteVBA，回退 excel_execute_vba）
 *       执行中间步经 yield 事件流（skeleton/progress/result/status）驱动步骤面板渐进填色。
 *
 * 挂载：App.ChatV2.skillPipelineModules.stepsExecute
 * 契约：与 skill-spec-execute.js 事件流对齐（skeleton/progress/result），
 *       skill-pipeline 消费方式与 spec 分支一致。
 *
 * 约束（STEPS_CONTRACT 唯一事实源）：
 *   - 默认产物是声明式工具序列，每个步骤=真实 MCP 工具，可审计/可预览/可分步填色；
 *   - VBA 仅当某步 strategy='code' 且提供 code 时执行，绝不默认生成（不产生不可审计黑盒）；
 *   - 校验与入库共用同一份 STEPS_CONTRACT.validateSteps，杜绝"产物与执行器各写一套"。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.skillPipelineModules) App.ChatV2.skillPipelineModules = {};
    const MODS = App.ChatV2.skillPipelineModules;
    const ChatLog = window.ChatLog || console;

    /** 取 STEPS_CONTRACT（声明式契约，来自 skill-domain.js 的 App.ChatV2.skillDomain） */
    function getContract() {
        const dom = App.ChatV2.skillDomain;
        if (!dom || !dom.STEPS_CONTRACT) {
            throw new Error('skillDomain.STEPS_CONTRACT 未就绪，声明式 steps 契约缺失');
        }
        return dom.STEPS_CONTRACT;
    }

    /** 从 Skill 元数据读取原始 steps（兼容 Steps 大写字段） */
    function readRawSteps(skill) {
        const raw = skill && (skill.steps ?? skill.Steps);
        return Array.isArray(raw) ? raw : [];
    }

    /** 取当前选区上下文参数（传给工具，便于无选区参数时也能定位默认范围） */
    function buildTurnParams(turn) {
        const ctx = turn?.context || {};
        const p = {};
        if (ctx.selection) p.selection = ctx.selection;
        if (Array.isArray(ctx.referenceRegions) && ctx.referenceRegions.length) {
            p.referenceRegions = ctx.referenceRegions;
        }
        if (Array.isArray(ctx.attachments) && ctx.attachments.length) {
            p.attachments = ctx.attachments;
        }
        return p;
    }

    // ============================================================
    // 模板变量渲染：{{ctx.xxx}} / {{stepName.field}} / {{stepName.result.xxx}}
    // ------------------------------------------------------------
    // 让声明式 Skill 的 params 能引用前序步骤产出和 turn context，
    // 不再需要 LLM 动态规划、也不需要 VBA 硬编码中间数据传递。
    // ============================================================

    /** 安全取嵌套属性：getter(obj, 'a.b.c') → obj?.a?.b?.c */
    function _gt(obj, path) {
        if (!obj || !path) return undefined;
        let cur = obj;
        for (const p of String(path).split('.').filter(Boolean)) {
            if (cur == null) return undefined;
            cur = typeof cur === 'object' ? cur[p] : undefined;
        }
        return cur;
    }

    /**
     * 渲染单个模板路径 → 真实值。
     * 命名空间：
     *   {{ctx.field}}                  → execCtx.ctx.field（turn context）
     *   {{stepName.field}}             → execCtx[stepName].field（归一化后顶层快捷字段）
     *   {{stepName.data}}              → 快捷：execCtx[stepName].data（或 .result.data）
     *   {{stepName.headers}}           → 快捷：execCtx[stepName].headers（或 .result.headers）
     *   {{stepName.result.xxx}}        → execCtx[stepName].result.xxx（invokeTracked 原始返回）
     */
    function _resolveTemplate(path, execCtx) {
        const p = String(path).trim();
        if (!p || !execCtx) return undefined;

        const parts = p.split('.');
        const first = parts[0];

        if (first === 'ctx') {
            return _gt(execCtx.ctx, parts.slice(1).join('.'));
        }

        const sc = execCtx[first];
        if (!sc) return undefined;

        // stepName.data / stepName.headers 快捷
        if (parts[1] === 'data' && parts.length === 2) {
            if (sc.data !== undefined) return sc.data;
            return _gt(sc.result, 'data');
        }
        if (parts[1] === 'headers' && parts.length === 2) {
            if (sc.headers !== undefined) return sc.headers;
            return _gt(sc.result, 'headers');
        }
        if (parts[1] === 'result') {
            return _gt(sc.result, parts.slice(2).join('.'));
        }

        return _gt(sc, parts.slice(1).join('.'));
    }

    /**
     * 递归遍历 params，把所有 {{xxx}} 占位符替换成真实值。
     * 单个模板且整个 string 就是模板 → 返回原始类型（object/array/bool/number 保持）。
     * 多个模板嵌在文本中间 → 字符串拼接。
     * 数字/布尔/日期等非 string 值原样返回。
     */
    function renderTemplates(params, execCtx) {
        if (params == null) return params;

        if (typeof params === 'string') {
            if (!/\{\{[^}]+\}\}/.test(params)) return params;

            // 整个 string 就是单个模板 → 返回真实类型
            const whole = params.match(/^\{\{([^}]+)\}\}$/);
            if (whole) {
                const v = _resolveTemplate(whole[1], execCtx);
                return v !== undefined ? v : params;
            }

            // 嵌在文本中间 → 全部转 string 拼接
            return params.replace(/\{\{([^}]+)\}\}/g, (_, path) => {
                const v = _resolveTemplate(path, execCtx);
                if (v === undefined || v === null) return '';
                if (typeof v === 'object') { try { return JSON.stringify(v); } catch { return String(v); } }
                return String(v);
            });
        }

        if (Array.isArray(params)) return params.map(item => renderTemplates(item, execCtx));
        if (typeof params === 'object') {
            const out = {};
            for (const k of Object.keys(params)) out[k] = renderTemplates(params[k], execCtx);
            return out;
        }
        return params; // number / bool 原样
    }

    /**
     * 把 invokeTracked 返回值归一化成扁平对象（挂到 execCtx[stepName]）。
     * 因为不同工具返回结构不一样，统一归一化后模板引用不管工具差异都能 work。
     */
    function _normalizeResult(toolName, raw) {
        const inner = (raw && raw.result != null) ? raw.result : raw;
        return {
            ok: !!(raw && raw.ok !== false && !raw.error),
            raw,
            result: inner,
            error: raw?.error || inner?.error || null,
            // 顶层快捷字段
            sheet: inner?.sheet || inner?.worksheet || inner?.resultSheet || null,
            range: inner?.range || inner?.dataRange || null,
            headers: Array.isArray(inner?.headers) ? inner.headers : null,
            data: Array.isArray(inner?.data) ? inner.data
                 : Array.isArray(inner?.rows) ? inner.rows
                 : Array.isArray(inner?.result) ? inner.result
                 : null
        };
    }

    /** 构建 UI 展示用参数摘要（避免参数面板撑爆） */
    function buildDisplayArgs(step, turn) {
        const disp = { 策略: getContract().displayStrategy[step.strategy] || step.strategy };
        const params = step.params || {};
        const keys = Object.keys(params);
        if (keys.length) {
            disp.参数 = keys.slice(0, 4).map(k => {
                const v = params[k];
                const s = typeof v === 'object' && v !== null ? JSON.stringify(v) : String(v);
                return k + '=' + (s.length > 30 ? s.slice(0, 30) + '…' : s);
            }).join('，' + (keys.length > 4 ? ' …' : ''));
        }
        return disp;
    }

    /**
     * 执行一个 code 步（显式 VBA）：优先 excelBridge.ExecuteVBA 直接通道，回退 MCP excel_execute_vba。
     * @returns {Promise<{ok:boolean,result?:any,error?:string}>}
     */
    function executeVbaStep(code, displayName, turn, toolRegistry, signal, sessionStore) {
        const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
        const attempt = async () => {
            if (code) {
                // 直连 COM 不经过工具守卫：破坏性 VBA（删除/清空类操作）必须先经用户确认，
                // 且遵守本轮对危险操作族的已有决定（拒绝 → 不再执行；同意 → 不重复询问）
                if (toolRegistry.isDestructiveVba && toolRegistry.isDestructiveVba(code)) {
                    const turnDec = toolRegistry.getTurnDangerousDecision ? toolRegistry.getTurnDangerousDecision(turn) : null;
                    if (turnDec === 'deny') {
                        ChatLog.info?.('PIPELINE', '[skill-steps] 本轮用户已取消危险操作，破坏性 VBA 自动拒绝');
                        return { ok: false, error: '用户已取消本轮危险操作，技能 VBA 含删除/清空操作被自动拦截', toolCallId: null, cancelled: true };
                    }
                    if (turnDec !== 'allow') {
                        const reason = '【安全守卫】该技能生成的 VBA 包含删除/清空数据的操作，执行后可能不可逆。'
                            + '\n\n技能步骤: ' + (displayName || 'VBA')
                            + '\n代码预览: ' + String(code).slice(0, 160).replace(/\s+/g, ' ') + '…';
                        let confirmedVba = false;
                        try { confirmedVba = await toolRegistry.confirmDangerous(displayName || 'excel_execute_vba', reason); } catch (_e) { confirmedVba = false; }
                        if (!confirmedVba) {
                            ChatLog.info?.('PIPELINE', '[skill-steps] 用户取消破坏性 VBA 执行');
                            if (toolRegistry.setTurnDangerousDecision) toolRegistry.setTurnDangerousDecision(turn, 'deny');
                            return { ok: false, error: '用户取消执行危险操作（技能 VBA 含删除/清空操作）', toolCallId: null, cancelled: true };
                        }
                    }
                }
                try {
                    const resultJson = await bridge.ExecuteVBA(code, displayName);
                    const parsed = typeof resultJson === 'string' ? JSON.parse(resultJson) : resultJson;
                    const ok = parsed?.success === true || parsed?.ok === true || (parsed && !parsed.error);
                    return ok
                        ? { ok: true, result: parsed, toolCallId: null }
                        : { ok: false, error: parsed?.error || parsed?.message || 'VBA 执行失败', toolCallId: null };
                } catch (bridgeErr) {
                    ChatLog.warn?.('PIPELINE', '[skill-steps] excelBridge.ExecuteVBA 异常，尝试 MCP 工具回退:', bridgeErr?.message);
                }
            }
            const origMode = turn.mode;
            turn.mode = 'expert';
            try {
                return await toolRegistry.invokeTracked(turn, 'excel_execute_vba', { code }, {
                    round: 1,
                    displayName: displayName + '（VBA）',
                    displayArgs: { '策略': 'VBA', '行数': String(code.split('\n').length) },
                    ctx: { turn, sessionStore, signal }
                });
            } finally {
                turn.mode = origMode;
            }
        };
        return attempt();
    }

    /**
     * 声明式 steps 执行器。
     * 事件流：
     *   { type:'status', text }
     *   { type:'skeleton', steps:[{id,title,detail,status,strategy,tool,...}] }
     *   { type:'progress', stepIndex, status, detail }
     *   { type:'result', result:{ outputText, steps } }
     * @param {Object} turn
     * @param {Object} skill
     * @param {string} displayName
     * @param {AbortSignal} signal
     */
    async function* executeSkillStepsMode(turn, skill, displayName, signal) {
        const sessionStore = App.ChatV2.sessionStore;
        const toolRegistry = App.ChatV2.toolRegistry;
        const contract = getContract();

        if (typeof toolRegistry?.init === 'function') await toolRegistry.init();
        if (!toolRegistry || typeof toolRegistry.invokeTracked !== 'function') {
            throw new Error('skill-steps: toolRegistry 未就绪，无法执行声明式步骤');
        }

        // 1. 归一化 + 校验（复用唯一契约）
        const raw = readRawSteps(skill);
        const { ok, steps, errors } = contract.validateSteps(raw);
        yield { type: 'status', text: `正在执行 ${displayName}（${steps.length} 步声明式流程）...` };

        if (!ok) {
            const detail = (errors || []).join('；');
            throw new Error(`声明式步骤校验失败（${steps.length} 步可用）：${detail}`);
        }

        // 2. 骨架：全 pending 步骤（与 spec 分支 skeleton 事件同构）
        const skeleton = steps.map((s, i) => ({
            id: s.tool || ('sk_step_' + i),
            title: s.name || (contract.displayStrategy[s.strategy] === 'VBA' ? 'VBA 步骤' : s.tool),
            description: s.description || s.name || s.tool,
            detail: s.description || s.tool || '步骤 ' + (i + 1),
            status: 'pending',
            strategy: s.strategy || 'direct',
            toolName: s.tool || null,
            index: i
        }));
        yield { type: 'skeleton', steps: skeleton };

        // 3. 保留每步真实结果，供最终摘要/后续引用
        const stepOutputs = [];
        // execCtx：模板变量命名空间（{{ctx.xxx}} + 每步产出 {{stepName.field}}）
        const execCtx = { ctx: turn?.context || {} };
        if (typeof toolRegistry?.init === 'function') await toolRegistry.init();

        for (let i = 0; i < steps.length; i++) {
            const step = steps[i];
            if (signal?.aborted) throw new Error('user_aborted');

            const stepName = step.name || (step.tool || '步骤 ' + (i + 1));
            yield { type: 'progress', stepIndex: i, status: 'running', detail: `执行：${stepName}` };

            let result;
            try {
                if (step.strategy === 'code') {
                    if (!step.code.trim()) {
                        throw new Error(`步骤[${i}] strategy='code' 但缺少 code（VBA）`);
                    }
                    result = await executeVbaStep(
                        step.code, stepName, turn, toolRegistry, signal, sessionStore
                    );
                } else {
                    // direct：MCP 工具 / skill.<name> 委托，走通用 toolRegistry（自动出工具卡片）
                    const baseParams = Object.assign({}, buildTurnParams(turn), step.params || {});
                    // 模板变量渲染：把 {{ctx.xxx}} / {{前序步骤名.field}} 替换成真实值
                    const mergedParams = renderTemplates(baseParams, execCtx);
                    result = await toolRegistry.invokeTracked(turn, step.tool, mergedParams, {
                        round: i + 1,
                        displayName: stepName,
                        displayArgs: buildDisplayArgs(step, turn),
                        ctx: { turn, sessionStore, signal }
                    });
                }
                if (signal?.aborted) throw new Error('user_aborted');
            } catch (stepErr) {
                if ((stepErr && stepErr.message) === 'user_aborted') throw stepErr;
                const msg = stepErr?.message || String(stepErr);
                stepOutputs.push({ name: stepName, error: msg });
                yield { type: 'progress', stepIndex: i, status: 'failed', detail: `失败：${msg}` };
                // 声明式步骤失败：中断流程并明确报错（可审计，不让黑盒静默吞掉）
                throw new Error(`步骤 ${i + 1}「${stepName}」执行失败：${msg}`);
            }

            const okStep = !!result && (result.ok === true || result.success === true);
            const resultMsg = (result && (result.message || result.result?.message)) || (okStep ? '已执行' : '');
            stepOutputs.push({ name: stepName, ok: okStep, result: result?.result ?? result });
            // 归一化产出 → 挂到 execCtx[stepName]，供后续步骤模板引用
            execCtx[stepName] = _normalizeResult(step.tool, result);
            yield { type: 'progress', stepIndex: i, status: okStep ? 'completed' : 'failed',
                detail: okStep ? (resultMsg || '已完成') : ('失败：' + (result?.error || '')) };
            if (!okStep) {
                throw new Error(`步骤 ${i + 1}「${stepName}」返回失败：${result?.error || '未知原因'}`);
            }
        }

        // 4. 归一化骨架（status 已由 progress 填色），产出 result 供 pipeline 收尾
        const finalSteps = steps.map((s, i) => {
            const out = stepOutputs[i] || {};
            return {
                title: s.name || s.tool || '步骤 ' + (i + 1),
                description: s.description || s.name || s.tool,
                detail: s.tool || '',
                status: out.error ? 'failed' : 'completed',
                strategy: s.strategy || 'direct',
                toolName: s.tool || null,
                params: s.params || null,
                result: out.result || null
            };
        });
        const outputText = `${displayName}：声明式流程执行完成（${steps.length} 步）。`;
        yield { type: 'result', result: { outputText, steps: finalSteps, stepOutputs } };
    }

    MODS.stepsExecute = {
        executeSkillStepsMode,
        readRawSteps,
        // 工具函数（方便测试 / debug）
        renderTemplates,
        normalizeResult: _normalizeResult
    };
})();