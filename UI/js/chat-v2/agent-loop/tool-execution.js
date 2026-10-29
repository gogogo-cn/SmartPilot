/**
 * chat-v2/agent-loop/tool-execution.js — 工具执行管线辅助
 * ============================================================
 * 从 agent-loop 抽离：失败熔断判定、执行结果归并等纯逻辑。
 * 主循环仍编排 for (tc of toolCalls) 的 VB/COM 副作用；
 * 本模块提供可单测的决策函数，避免熔断/toolsExecuted 判定漂移。
 * 技术债：continue 不计 round、forceRetryCount 跨守卫共用 —— 见主循环注释。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /** 同一"工具名+参数"连续失败多少次后熔断 */
    const SAME_CALL_FAIL_LIMIT = 2;

    /**
     * 是否应对本调用做失败熔断（跳过真实 invoke）。
     * @param {Map|undefined} failStreak turn._failStreak
     * @param {string} idemKey
     * @param {boolean} reprobe 用户要求重做时不熔断
     * @param {number} [limit]
     * @returns {{ break: boolean, failN: number, lastErr: string }}
     */
    function checkFailBreaker(failStreak, idemKey, reprobe, limit) {
        const lim = Number(limit) > 0 ? Number(limit) : SAME_CALL_FAIL_LIMIT;
        if (reprobe || !idemKey) return { break: false, failN: 0, lastErr: '' };
        const fs = failStreak instanceof Map ? failStreak.get(idemKey) : null;
        const failN = (fs && fs.n) || 0;
        if (failN < lim) return { break: false, failN, lastErr: '' };
        return {
            break: true,
            failN,
            lastErr: String((fs && fs.error) || '未知错误').slice(0, 200)
        };
    }

    /**
     * 根据本轮 localResults 归并 toolsExecuted / anyToolFailed。
     * 规则：至少一次 ok 且无硬失败 → toolsExecuted；任一 ok===false → anyToolFailed。
     * @returns {{ anyToolFailed: boolean, toolsExecuted: boolean, anyOk: boolean }}
     */
    function mergeToolRoundFlags(localResults, prevAnyToolFailed) {
        const list = Array.isArray(localResults) ? localResults : [];
        const anyOk = list.some(r => r && r.ok === true);
        let anyToolFailed = !!prevAnyToolFailed;
        if (!anyToolFailed && list.some(r => r && r.ok === false)) anyToolFailed = true;
        return {
            anyOk,
            anyToolFailed,
            toolsExecuted: anyOk && !anyToolFailed
        };
    }

    /**
     * 本轮工具是否全是只读探测（图表补链应延后给闸门）。
     * @param {Array<{name?:string}>} toolCalls
     */
    function isOnlyProbeToolRound(toolCalls) {
        if (!Array.isArray(toolCalls) || toolCalls.length === 0) return false;
        return toolCalls.every(tc =>
            /^(excel_)?(read_range|get_|list_|describe_)/i.test(
                String(tc && tc.name || '').replace(/^mcp_/, '')
            )
        );
    }

    App.ChatV2.agentLoopToolExecution = {
        SAME_CALL_FAIL_LIMIT,
        checkFailBreaker,
        mergeToolRoundFlags,
        isOnlyProbeToolRound,
        /** @type {null|Function} 可选：完整执行管线注入点（主循环暂未迁完） */
        executeToolCalls: null
    };
    ChatLog.info?.('CHAT-V2', '✅ agentLoopToolExecution loaded');
})();
