/**
 * ============================================================
 * chat-v2/services/pipeline-registry.js — Pipeline 注册表
 * ============================================================
 * 职责：统一注册 / 发现 / 路由 7 条 pipeline（参考 DSH 插件化设计：
 *       "新能力 = 注册，不修改主循环"）
 *   - register(kind, definition)    — 声明式注册（execute/name/description/priority）
 *   - get(kind) / list()            — 查询
 *   - pick(kind)                    — 路由（未知 kind 降级到 text，优先级可配 fallback）
 *   - 自动同步 App.ChatV2.pipelines — 向后兼容旧引用（orchestrator / verifyModules）
 * 设计：
 *   - pipeline 文件只做"注册"，路由决策由本注册表 + orchestrator 完成
 *   - 支持优先级排序（priority 越小越优先，默认 100）
 *   - 支持显式 fallbackKind（未知 kind 时的降级目标，默认 'text'）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /** @type {Map<string, Object>} kind -> { kind, execute, name, description, priority, fallbackKind } */
    const _pipelines = new Map();

    /**
     * 注册一条 pipeline。
     * @param {string} kind - pipeline 标识（text/image/doc/python/html/skill/code）
     * @param {Object} definition
     * @param {Function} definition.execute - (turn, signal) => AsyncGenerator<typed event>
     * @param {string} [definition.name] - 显示名
     * @param {string} [definition.description] - 描述
     * @param {number} [definition.priority=100] - 优先级（越小越优先）
     * @param {string} [definition.fallbackKind] - 未知 kind 降级目标（默认 'text'）
     * @returns {boolean}
     */
    function register(kind, definition) {
        if (!kind || typeof kind !== 'string') {
            ChatLog.warn?.('PIPE-REG', '[pipeline-registry] register 失败：缺少 kind');
            return false;
        }
        if (!definition || typeof definition.execute !== 'function') {
            ChatLog.warn?.('PIPE-REG', `[pipeline-registry] register(${kind}) 失败：缺少 execute`);
            return false;
        }
        const entry = {
            kind,
            execute: definition.execute,
            name: definition.name || kind,
            description: definition.description || '',
            priority: typeof definition.priority === 'number' ? definition.priority : 100,
            fallbackKind: definition.fallbackKind || 'text'
        };
        _pipelines.set(kind, entry);
        // 同步到旧命名空间，向后兼容（App.ChatV2.pipelines.<kind> = { execute }）
        if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};
        App.ChatV2.pipelines[kind] = { execute: entry.execute };
        ChatLog.info?.('PIPE-REG', `[pipeline-registry] 注册 pipeline: ${kind} (${entry.name}, priority=${entry.priority})`);
        return true;
    }

    /**
     * 批量注册（registerAll）
     * @param {Object} defs - { kind: definition, ... }
     * @returns {number} 成功注册数
     */
    function registerAll(defs) {
        if (!defs || typeof defs !== 'object') return 0;
        let count = 0;
        for (const [kind, def] of Object.entries(defs)) {
            if (register(kind, def)) count++;
        }
        return count;
    }

    /**
     * 获取 pipeline
     * @param {string} kind
     * @returns {Object|null}
     */
    function get(kind) {
        return _pipelines.get(kind) || null;
    }

    /**
     * 列出所有已注册 pipeline（按 priority 升序）
     * @returns {Array<Object>}
     */
    function list() {
        return Array.from(_pipelines.values()).sort((a, b) => a.priority - b.priority);
    }

    /**
     * 路由：按 kind 选取 pipeline。
     *   - kind 已注册 → 返回该 pipeline
     *   - kind 未注册 → 降级到 'text'（产品默认管线）
     * 说明：fallbackKind 是注册时的元数据（供 orchestrator 决策参考），
     *       pick 本身对"已注册"的 kind 不做跳过。
     * @param {string} kind
     * @returns {Object|null} 选中的 pipeline 定义（含 execute）
     */
    function pick(kind) {
        const direct = _pipelines.get(kind);
        if (direct) return direct;
        if (!kind) {
            ChatLog.debug?.('PIPE-REG', '[pipeline-registry] pick 收到空 kind，降级到 text');
            return _pipelines.get('text') || null;
        }
        // 未知 kind：降级到默认 text
        const fallback = _pipelines.get('text');
        if (!fallback) {
            ChatLog.warn?.('PIPE-REG', `[pipeline-registry] pick(${kind}) 无可用 pipeline（含降级）`);
            return null;
        }
        ChatLog.info?.('PIPE-REG', `[pipeline-registry] pick(${kind}) 未注册，降级 → text`);
        return fallback;
    }

    /**
     * 取消注册
     * @param {string} kind
     * @returns {boolean}
     */
    function unregister(kind) {
        const removed = _pipelines.delete(kind);
        if (removed && App.ChatV2.pipelines) {
            delete App.ChatV2.pipelines[kind];
        }
        return removed;
    }

    /**
     * 调试信息
     */
    function getDebugInfo() {
        return {
            total: _pipelines.size,
            pipelines: list().map(p => ({ kind: p.kind, name: p.name, priority: p.priority, fallbackKind: p.fallbackKind }))
        };
    }

    App.ChatV2.pipelineRegistry = { register, registerAll, get, list, pick, unregister, getDebugInfo };
    ChatLog.info?.('CHAT-V2', '✅ pipelineRegistry loaded');
})();
