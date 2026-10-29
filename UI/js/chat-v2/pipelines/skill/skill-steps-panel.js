/**
 * ============================================================
 * skill-steps-panel.js — Skill 管道步骤面板辅助函数（从 skill-pipeline.js 拆出）
 * ------------------------------------------------------------
 * 让 agent/fixed/spec 三种执行模式统一产出 turn._planHtml + turn._steps，
 * 复用 ProgressTracker.buildPlanHtml 渲染绿面板，与 text-pipeline 多任务体验一致。
 * 挂载：App.ChatV2.skillPipelineModules.stepsPanel
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

    // ============================================================
    // 【进度统一】Skill 管道步骤面板辅助函数
    // 让 agent/fixed/spec 三种执行模式统一产出 turn._planHtml + turn._steps，
    // 复用 ProgressTracker.buildPlanHtml 渲染绿面板，与 text-pipeline 多任务体验一致。
    //
    // 设计：init 可带【预规划骨架】，执行中 appendSkillStep 按 title 匹配→update 而非追加，
    // 实现"骨架先出 → 逐步填色"的渐进式体验。没有骨架时 appendSkillStep 正常追加。
    // ============================================================

    /**
     * 基于 Skill 元数据生成【预规划骨架】（全 pending）。
     * 不同 executionMode 策略：
     *   spec  → 1 步（确定性 SkillSpec 执行，步骤已由 Spec 引擎产出，这里只占位）
     *   agent → agentConfig.maxIterations 步（ReAct 轮次骨架，默认 8）
     *   fixed → 1 主 Skill + subSkills 个子 Skill 骨架
     * @param {Object} skill - YAML 解析后的 skill manifest
     * @param {string} executionMode - 'spec' | 'agent' | 'fixed'
     * @returns {Array<Object>} skeleton steps (status='pending')，为空表示无法预判
     */
    function buildSkillSkeleton(skill, executionMode) {
        if (!skill) return [];
        const displayName = skill.displayName || skill.name || 'Skill';

        if (executionMode === 'fixed') {
            // fixed 模式：主 Skill + subSkills 顺序执行
            const subSkills = skill.subSkills || [];
            const skeleton = [];
            skeleton.push({
                title: displayName,
                detail: `${displayName} 主执行`,
                status: 'pending',
                strategy: 'pipeline',
                toolName: 'skill.' + skill.name,
                id: `sk_${skill.name}_main`
            });
            subSkills.forEach((sub, i) => {
                skeleton.push({
                    title: typeof sub === 'string' ? sub : (sub.displayName || sub.name || '子 Skill'),
                    detail: `子流程 ${i + 1}/${subSkills.length}`,
                    status: 'pending',
                    strategy: 'pipeline',
                    toolName: 'skill.' + (typeof sub === 'string' ? sub : (sub.name || '')),
                    id: `sk_${skill.name}_sub_${i}`
                });
            });
            return skeleton;
        }

        if (executionMode === 'agent') {
            // agent 模式：ReAct 轮次骨架
            const agentCfg = skill.agentConfig || {};
            const maxIter = Number(agentCfg.maxIterations) > 0
                ? Number(agentCfg.maxIterations) : 8;
            const skeleton = [];
            for (let i = 1; i <= maxIter; i++) {
                skeleton.push({
                    _round: i,
                    title: `第 ${i} 轮推理`,
                    detail: `自主 Agent 执行第 ${i}/${maxIter} 轮`,
                    status: 'pending',
                    strategy: 'agent',
                    toolName: null,   // agent 轮次不确定调用哪个工具
                    id: `sk_${skill.name}_round_${i}`
                });
            }
            return skeleton;
        }

        if (executionMode === 'spec') {
            // spec 模式：1 步占位（Spec 引擎产出的步骤更细，这里只先出总骨架）
            return [{
                title: displayName + '（Spec 引擎）',
                detail: '确定性执行',
                status: 'pending',
                strategy: 'direct',
                toolName: 'skill.' + skill.name,
                id: `sk_${skill.name}_spec`
            }];
        }

        if (executionMode === 'steps') {
            // steps 模式：SkillCreator 默认产物，直接以声明式 steps 作为预规划骨架（全 pending）
            const raw = skill && (skill.steps ?? skill.Steps);
            if (!Array.isArray(raw) || raw.length === 0) return [];
            return raw.map((s, i) => {
                const step = (s && typeof s === 'object') ? s : {};
                const tool = typeof step.tool === 'string' ? step.tool : '';
                const strategy = (typeof step.strategy === 'string' && step.strategy.trim()) || 'direct';
                const name = (typeof step.name === 'string' && step.name.trim())
                    ? step.name.trim() : (tool || ('步骤 ' + (i + 1)));
                return {
                    title: name,
                    description: (typeof step.description === 'string' && step.description.trim())
                        ? step.description.trim() : name,
                    detail: name,
                    status: 'pending',
                    strategy: strategy,
                    toolName: tool || null,
                    id: tool || `sk_${skill.name}_step_${i}`
                };
            });
        }

        // 未知 executionMode：返回空数组，走 fallback（执行中逐步追加）
        return [];
    }

    /**
     * 初始化 Skill 步骤面板：清空旧状态；若有 skeleton 则直接塞进去并渲染 pending 面板。
     * @param {Object} turn
     * @param {Object} sessionStore
     * @param {Array<Object>} [skeleton] - 可选预规划骨架（全 pending 步骤）
     */
    function initSkillStepsPanel(turn, sessionStore, skeleton) {
        if (!turn) return;
        turn._skillSteps = (Array.isArray(skeleton) && skeleton.length > 0)
            ? skeleton.map(s => ({ ...s, status: s.status || 'pending' }))
            : [];
        // 清空可能的旧面板（来自上一轮、或 fixed→agent 降级等）
        // 【步骤统一到 cbx】旧 _planHtml 绿面板已下线，改为清空 thoughtList 里的步骤行
        try { if (sessionStore && typeof sessionStore.clearStepThoughts === 'function') sessionStore.clearStepThoughts(turn.id); } catch (e) { /* 忽略 */ }
        // 有骨架立即渲染 pending 面板（让用户先看到完整规划再开始执行）
        if (turn._skillSteps.length > 0) {
            renderSkillSteps(turn, sessionStore);
        }
    }

    /**
     * 向 Skill 步骤面板写入一步：若同名（title/id/toolName 任一匹配）则 update，否则 push。
     * 这样配合 initSkillStepsPanel(skeleton) 实现"骨架先出 → 执行中逐步填色"。
     * @param {Object} turn
     * @param {Object} sessionStore
     * @param {Object} step - { title, detail, status, toolName, params?, result?, id? }
     */
    function appendSkillStep(turn, sessionStore, step) {
        if (!turn) return;
        if (!Array.isArray(turn._skillSteps)) turn._skillSteps = [];

        // 匹配规则：id 精确匹配优先；否则 title 精确匹配；再退化到 toolName 匹配
        let idx = -1;
        if (step.id) idx = turn._skillSteps.findIndex(s => s.id === step.id);
        if (idx < 0 && step.title) idx = turn._skillSteps.findIndex(s => s.title === step.title);
        if (idx < 0 && step.toolName) idx = turn._skillSteps.findIndex(s => s.toolName === step.toolName);

        if (idx >= 0) {
            // 骨架已有同名步骤 → update 填色（保留骨架 id 不变）
            const merged = { ...turn._skillSteps[idx], ...step, id: turn._skillSteps[idx].id || step.id };
            turn._skillSteps[idx] = merged;
        } else {
            // 无骨架 / 骨架没这步 → 追加
            turn._skillSteps.push({ ...step });
        }
        renderSkillSteps(turn, sessionStore);
    }

    /**
     * turn._skillSteps 格式 → ProgressTracker 步骤格式。
     * 状态归一化：done/success/ok→completed，error/err/aborted/canceled→failed；并透传 subSteps。
     * @param {Object} turn
     * @returns {Array<Object>}
     */
    function toPtSteps(turn) {
        // 我们的: { title, detail, status, strategy, toolName, params?, result? }
        // PT 要的: { name, description, status, strategy, toolName, params, result, id }
        return turn._skillSteps.map((s, i) => {
            // 状态归一化：把外部的 done/success/error 等统一映射到 PT 识别的四态
            const raw = s.status || 'pending';
            let status = raw;
            if (raw === 'done' || raw === 'success' || raw === 'ok') status = 'completed';
            else if (raw === 'error' || raw === 'err' || raw === 'aborted' || raw === 'canceled') status = 'failed';
            else if (raw !== 'pending' && raw !== 'running' && raw !== 'completed' && raw !== 'failed') status = 'pending';

            return {
                id: s.id || `skill_${turn.id}_step_${i}`,
                name: s.title || s.name || '',
                description: s.title || s.description || `步骤 ${i + 1}`,
                detail: s.detail || s.description || '',
                status: status,
                strategy: s.strategy || 'direct',
                toolName: s.toolName || null,
                params: s.params || null,
                result: s.result || null,
                // ✨ 透传嵌套子步骤（清洗/增删改查等大步骤内的实际动作），否则子步骤永远不会进 ProgressTracker
                subSteps: Array.isArray(s.subSteps) ? s.subSteps.map((sub, j) => {
                    const subRaw = sub.status || sub.state || 'pending';
                    let subStatus = subRaw;
                    if (subRaw === 'done' || subRaw === 'success' || subRaw === 'ok') subStatus = 'completed';
                    else if (subRaw === 'error' || subRaw === 'err') subStatus = 'failed';
                    else if (subStatus !== 'pending' && subStatus !== 'running' && subStatus !== 'completed' && subStatus !== 'failed') subStatus = 'pending';
                    return {
                        id: sub.id || `skill_${turn.id}_sub_${i}_${j}`,
                        name: sub.title || sub.name || `子步骤 ${j + 1}`,
                        description: sub.description || sub.title || sub.name || `子步骤 ${j + 1}`,
                        detail: sub.detail || sub.description || '',
                        status: subStatus,
                        strategy: sub.strategy || 'direct',
                        toolName: sub.toolName || null,
                        params: sub.params || null
                    };
                }) : undefined
            };
        });
    }

    /**
     * 整块渲染：基于 turn._skillSteps 重新生成 _planHtml 并持久化。
     * 使用 ProgressTracker.buildPlanHtml（带进度条、✓/✕ 徽章、可折叠详情的完整视觉），
     * 与 text-pipeline 多任务面板完全对齐。用于首次出面板 / 需要整体重建时；
     * 执行过程中的"逐步填色"走 renderSkillStepsPatch 的增量通道。
     */
    function renderSkillSteps(turn, sessionStore) {
        if (!turn || !Array.isArray(turn._skillSteps) || turn._skillSteps.length === 0) return;
        const ptSteps = toPtSteps(turn);
        // 【步骤统一到 cbx v20261003】不再生成 _planHtml 旧绿面板（.exec-plan-card），
        // 改为把每一步写入 thoughtList（kind='step'），由渲染层 contentEl 顶部的
        // cbx 扁平列表统一展示；upsert 幂等，重复调用只更新不新增行。
        if (sessionStore && typeof sessionStore.upsertStepThought === 'function') {
            ptSteps.forEach(function (s, i) {
                sessionStore.upsertStepThought(turn.id, {
                    id: s.id || ('skill_step_' + i),
                    title: s.description || s.name || ('步骤 ' + (i + 1)),
                    detail: s.detail || '',
                    status: s.status || 'pending',
                    toolName: s.toolName || null,
                    index: i
                });
            });
        }
        // ProgressTracker 仍保持内步数据同步（统计 / 子步骤等其它逻辑依赖）
        const pt = window.App?.ChatV2?.ProgressTracker;
        if (pt && typeof pt.init === 'function') {
            try { pt.init(ptSteps, turn, sessionStore); } catch (e) { /* 忽略 */ }
        }
    }

    /**
     * 【增量渲染】只把"变化的那一步"作为 patch 推给渲染层，由 chat-renderer 按 data-step-id
     * 原位替换单行、并按 stats 只更新进度条，避免整块 outerHTML 替换导致的闪烁、进度条跳变、
     * <details> 展开态丢失。
     * patch 契约：turn._stepPatch = { seq, op, id?, index?, html?, planHtml, stats }
     *   - op='reset'  → 面板首次出现 / 骨架重建；planHtml 为完整面板 HTML，渲染层整块插入
     *   - op='upsert' → html 为单步行 HTML；planHtml 仅在"面板尚不存在"时兜底整块插入
     * @param {Object} turn
     * @param {Object} sessionStore
     * @param {Array<string|number>} [changedIds] - 本次变化的步骤 id 列表
     * @param {string} [op] - 'upsert'（默认）| 'reset'
     */
    function renderSkillStepsPatch(turn, sessionStore, changedIds, op) {
        if (!turn || !Array.isArray(turn._skillSteps) || turn._skillSteps.length === 0) return;
        if (!sessionStore) return;

        const ptSteps = toPtSteps(turn);
        // 【步骤统一到 cbx v20261003】不再推 _stepPatch（那是往旧 .exec-plan-card 打补丁的通道），
        // 改为只把变化的步骤 upsert 进 thoughtList，渲染层按 thought_update 重渲 cbx 行。
        const ids = Array.isArray(changedIds) ? changedIds.filter(id => id != null) : [];
        const targets = (op === 'reset' || ids.length === 0)
            ? ptSteps
            : ptSteps.filter(s => ids.some(id => String(id) === String(s.id)));

        if (sessionStore.upsertStepThought) {
            targets.forEach(function (s) {
                const index = ptSteps.indexOf(s);
                sessionStore.upsertStepThought(turn.id, {
                    id: s.id || ('skill_step_' + index),
                    title: s.description || s.name || ('步骤 ' + (index + 1)),
                    detail: s.detail || '',
                    status: s.status || 'pending',
                    toolName: s.toolName || null,
                    index: index
                });
            });
        }
    }

    /** 组装并写入单个 patch（seq 单调递增，渲染层据此丢弃乱序 patch） */
    function pushPatch(turn, sessionStore, patch) {
        turn._stepPatchSeq = (turn._stepPatchSeq || 0) + 1;
        const payload = Object.assign({ seq: turn._stepPatchSeq }, patch);
        try { sessionStore.updateTurn(turn.id, { _stepPatch: payload }); } catch (e) { /* 忽略 */ }
    }

    /** patch 携带的进度统计（渲染层据此只改进度条宽度与文案，不重建 DOM） */
    function computePatchStats(ptSteps) {
        const total = ptSteps.length;
        const completed = ptSteps.filter(s => s.status === 'completed').length;
        const failed = ptSteps.filter(s => s.status === 'failed').length;
        const pct = total > 0 ? Math.round((completed / total) * 100) : 0;
        return { total, completed, failed, pct };
    }

    /**
     * 最终收尾：把 turn._skillSteps 折叠进 turn._steps，
     * 与现有 spec 模式（用 turn._steps 渲染）保持双写兼容。
     */
    function finalizeSkillSteps(turn, sessionStore) {
        if (!turn) return;
        if (Array.isArray(turn._skillSteps) && turn._skillSteps.length > 0) {
            // ProgressTracker 需要最终的 completed/failed 状态刷新
            renderSkillSteps(turn, sessionStore);
            // 双写：保持旧链路（spec 模式的 flushContentRender 里 turn._steps 消费）
            turn._steps = turn._skillSteps
                .filter(s => s && typeof s === 'object' && (s.title || s.name))
                .map(s => ({
                    title: s.title || s.name,
                    detail: s.detail || s.description || '',
                    status: s.status || 'done'
                }));
        }
    }

    MODS.stepsPanel = {
        buildSkillSkeleton,
        initSkillStepsPanel,
        appendSkillStep,
        renderSkillSteps,
        renderSkillStepsPatch,
        finalizeSkillSteps
    };
})();
