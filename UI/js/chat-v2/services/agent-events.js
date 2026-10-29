/**
 * ============================================================
 * chat-v2/services/agent-events.js — Agent 事件总线
 * ============================================================
 * 职责：提供 agent/* 命名空间的轻量事件广播（参考 DSH 架构）
 *   - turn/start, turn/end          — 轮次生命周期
 *   - step/start, step/end          — 步骤生命周期（一次模型请求 + 其工具调用）
 *   - agent/status                  — 运行状态变化（running/idle）
 *   - agent/request                 — 模型请求发起（waterfall 预留）
 *   - agent/error                   — 循环级错误
 *   - tool/call, tool/result        — 工具调用与结果
 *   - agent/inject                  — 上下文注入（预留）
 * 设计：
 *   - 发布/订阅模式，任意模块可订阅，不依赖 sessionStore
 *   - 与 sessionStore 通知并存：sessionStore 负责状态持久化，
 *     agentEvents 负责实时生命周期广播（UI 进度 / 调试面板可订阅）
 *   - 事件名与 DSH 的 agent/* 事件对齐，便于后续迁移
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /** @type {Map<string, Array<Function>>} 事件名 -> 监听器列表 */
    const _listeners = new Map();

    /**
     * 订阅事件
     * @param {string} event - 事件名，如 'turn/start'、'step/end'、'agent/status'
     * @param {Function} fn - 监听器 (payload) => void
     * @returns {Function} 取消订阅的 disposer
     */
    function on(event, fn) {
        if (!event || typeof fn !== 'function') return () => {};
        if (!_listeners.has(event)) _listeners.set(event, []);
        const list = _listeners.get(event);
        list.push(fn);
        return () => off(event, fn);
    }

    /**
     * 一次性订阅
     * @param {string} event - 事件名
     * @param {Function} fn - 监听器
     * @returns {Function} 取消订阅的 disposer
     */
    function once(event, fn) {
        const wrapper = (payload) => {
            off(event, wrapper);
            fn(payload);
        };
        return on(event, wrapper);
    }

    /**
     * 取消订阅
     * @param {string} event - 事件名
     * @param {Function} fn - 监听器（不传则移除该事件全部监听器）
     */
    function off(event, fn) {
        if (!_listeners.has(event)) return;
        if (!fn) {
            _listeners.delete(event);
            return;
        }
        const list = _listeners.get(event);
        const idx = list.indexOf(fn);
        if (idx >= 0) list.splice(idx, 1);
        if (list.length === 0) _listeners.delete(event);
    }

    /**
     * 广播事件（同步调用所有监听器，异常隔离）
     * @param {string} event - 事件名
     * @param {Object} payload - 事件负载
     */
    function emit(event, payload) {
        const list = _listeners.get(event);
        if (!list || list.length === 0) return;
        // 复制一份，避免监听器内部增删影响本次广播
        for (const fn of [...list]) {
            try {
                fn(payload);
            } catch (e) {
                ChatLog.warn?.('AGENT-EVENTS', `[agent-events] ${event} 监听器异常:`, e);
            }
        }
    }

    /**
     * 当前订阅统计（调试用）
     */
    function getDebugInfo() {
        const info = {};
        for (const [k, v] of _listeners) info[k] = v.length;
        return { total: _listeners.size, listeners: info };
    }

    App.ChatV2.agentEvents = { on, once, off, emit, getDebugInfo };
    ChatLog.info?.('CHAT-V2', '✅ agentEvents loaded');
})();
