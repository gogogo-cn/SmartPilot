/**
 * ============================================================
 * chat-v2/services/TaskScheduler.js — 步骤调度器
 * ============================================================
 * 职责：将步骤列表转化为可执行的执行序列。
 *   - MVP：顺序执行（保持步骤顺序，支持依赖解析准备）
 *   - 后续：DAG 并行调度（依赖解析 -> 拓扑排序 -> 并行执行）
 *
 * 设计：
 *   - 单一职责：只负责调度决策（"谁先执行"），不负责执行
 *   - 依赖解析：支持 dependsOn 字段定义步骤间依赖关系
 *   - 可扩展：新增调度策略只需替换 schedule() 方法
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 调度步骤（MVP：顺序执行，附带依赖验证）
     * @param {Object[]} steps - 步骤定义列表
     * @param {Object} [opts] - 可选参数
     * @param {boolean} [opts.strictDeps=false] - 是否严格校验依赖（默认宽松）
     * @returns {Object} { schedule: Object[], errors: string[] }
     *   - schedule: 调度后的执行队列（每个元素是 {step, level}）
     *   - errors: 依赖解析错误列表
     *
     * 步骤格式：
     *   { id: 'step1', tool: 'excel_read_sheet', dependsOn: ['step0'] }
     *   dependsOn 为可选字段，指向其他步骤的 id
     */
    function schedule(steps, opts) {
        const options = opts || {};
        const strictDeps = options.strictDeps === true;

        if (!Array.isArray(steps) || steps.length === 0) {
            return { schedule: [], errors: ['步骤列表为空'] };
        }

        // 为每个步骤分配 id（如果未指定）
        const indexedSteps = steps.map((step, i) => ({
            ...step,
            _index: i,
            _id: step.id || step.name || step.tool || `step_${i}`
        }));

        const errors = [];
        const stepIds = new Set(indexedSteps.map(s => s._id));

        // 验证依赖关系
        for (const step of indexedSteps) {
            const deps = step.dependsOn;
            if (Array.isArray(deps) && deps.length > 0) {
                for (const depId of deps) {
                    if (!stepIds.has(depId)) {
                        const msg = `步骤 "${step._id}" 依赖的 "${depId}" 不存在`;
                        errors.push(msg);
                        if (strictDeps) {
                            ChatLog.error?.('MULTI-STEP', `[TaskScheduler] ${msg}`);
                        } else {
                            ChatLog.warn?.('MULTI-STEP', `[TaskScheduler] ${msg}（宽松模式，忽略）`);
                        }
                    }
                }
            }
        }

        // 检测循环依赖（如果存在 dependsOn）
        try {
            detectCyclicDependencies(indexedSteps);
        } catch (e) {
            errors.push(e.message);
            if (strictDeps) {
                // 循环依赖在严格模式下阻止执行
                return { schedule: [], errors };
            }
            ChatLog.warn?.('MULTI-STEP', `[TaskScheduler] 检测到循环依赖（宽松模式，按原始顺序执行）`);
        }

        // MVP：顺序执行，每个步骤分配 level=0（无并行）
        const schedule = indexedSteps.map(step => ({
            step,
            level: 0,
            order: step._index
        }));

        // 按原始顺序排序
        schedule.sort((a, b) => a.order - b.order);

        return { schedule, errors };
    }

    /**
     * 检测循环依赖（DFS）
     * 在循环依赖存在时抛出错误
     */
    function detectCyclicDependencies(steps) {
        const depMap = new Map();

        for (const step of steps) {
            const id = step._id;
            const deps = step.dependsOn || [];
            depMap.set(id, { deps, visited: false, inStack: false });
        }

        for (const [id] of depMap) {
            if (!depMap.get(id).visited) {
                dfs(id, depMap, []);
            }
        }

        function dfs(nodeId, map, path) {
            const node = map.get(nodeId);
            if (!node) return;
            if (node.inStack) {
                // 发现循环
                const cycleStart = path.indexOf(nodeId);
                const cycle = path.slice(cycleStart).concat(nodeId);
                throw new Error(`循环依赖检测: ${cycle.join(' -> ')}`);
            }
            if (node.visited) return;

            node.visited = true;
            node.inStack = true;
            path.push(nodeId);

            for (const depId of node.deps) {
                if (map.has(depId)) {
                    dfs(depId, map, path);
                }
            }

            path.pop();
            node.inStack = false;
        }
    }

    /**
     * 计算每个步骤的层级（用于后续DAG并行时确定并行级别）
     * @param {Object[]} steps - 带 dependsOn 的步骤列表
     * @returns {Map<string, number>} 步骤id -> 层级
     */
    function computeLevels(steps) {
        const levels = new Map();

        // 拓扑排序思想：没有依赖的步骤 level=0
        // 依赖的步骤 level = max(dep levels) + 1
        function getLevel(stepId) {
            if (levels.has(stepId)) return levels.get(stepId);

            const step = steps.find(s => (s.id || s.name || s.tool || `step_${s._index}`) === stepId);
            if (!step) return 0;

            const deps = step.dependsOn || [];
            if (deps.length === 0) {
                levels.set(stepId, 0);
                return 0;
            }

            let maxDepLevel = 0;
            for (const depId of deps) {
                const depLevel = getLevel(depId);
                maxDepLevel = Math.max(maxDepLevel, depLevel);
            }
            const level = maxDepLevel + 1;
            levels.set(stepId, level);
            return level;
        }

        for (const step of steps) {
            const id = step.id || step.name || step.tool || `step_${step._index}`;
            getLevel(id);
        }

        return levels;
    }

    // 导出
    App.ChatV2.TaskScheduler = {
        schedule,
        computeLevels
    };

})();