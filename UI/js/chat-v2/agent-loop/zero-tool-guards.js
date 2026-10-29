/**
 * chat-v2/agent-loop/zero-tool-guards.js — 零工具守卫决策辅助
 * ============================================================
 * 从 agent-loop 抽离：未完成信号、收尾/重试决策的纯判定。
 * 假执行清洗 / 强匹配强制调用等仍在主循环（依赖 findCandidateTools 闭包）；
 * 本模块提供可单测的信号与高层 action 建议，避免 incomplete 误伤回归。
 * 技术债：forceRetryCount 与 invalid-name / fake-exec 共用（本轮不拆预算）。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 只匹配未来式/续做导向；裸词「然后/还有/计划」易误伤正常总结
    // 2026-10-22 扩充：实测"美化工作簿中的所有工作表"被弱模型用"正在依次应用相同预设至其他工作表，
    // 请稍候。"收尾——这些过程语言一条都不在旧表里，回合以"任务完成"闭环，实际只做了 1 张表。
    // 新增项均为"最终总结里不该出现"的续做/过程措辞（请稍候/正在依次/将继续…）。
    const INCOMPLETE_SIGNAL = /(下一步|接下来|还需要|待完成|TODO|要做|我将|我会|请稍候|请稍等|稍候|正在依次|将继续|会继续|继续处理|继续执行|还未完成|尚未完成|next step|to do)/i;

    /**
     * 收尾正文里的"悬挂承诺"句：模型在任务没做完时用过程语言收尾
     * （"接下来，将继续美化工作簿中的其他工作表，请稍候。"）。
     * 这些句子属于过程叙述，永远不该出现在最终总结里；覆盖度/图表闸门兜不住时
     * （引导预算用尽），至少要把承诺句从正文里剔掉，避免用户以为还有后续在执行。
     * 三类模式（同一句内"未来式指引 + 续做/过程词"才删，避免误伤"接下来你可以查看图表"这类建议句）：
     *   ① 接下来/下一步/随后… + 继续/将/会/依次/处理/执行/美化/格式化…
     *   ② "正在依次…"过程叙述
     *   ③ 独立的"…，请稍候/请稍等。"过程句
     * @param {string} content
     * @returns {string} 清洗后的正文（无悬挂承诺句时原样返回）
     */
    function stripDanglingPromises(content) {
        const text = String(content || '');
        if (!text) return text;
        const PATTERNS = [
            /[^。！？!?\n]*(?:接下来|下一步|随后|稍后|即将)[^。！？!?\n]*(?:继续|将|会|请稍[候等]|稍[候等]|依次|处理|执行|操作|美化|格式化|更改|修改)[^。！？!?\n]*[。！？!?]?/gi,
            /[^。！？!?\n]*正在依次[^。！？!?\n]*[。！？!?]?/gi,
            /[^。！？!?\n]*请稍[候等][^。！？!?\n]*[。！？!?]?/gi
        ];
        let out = text;
        for (const re of PATTERNS) out = out.replace(re, '');
        return out.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    }

    /**
     * @param {string} assistantContent
     * @param {{ keepToolsEveryRound:boolean, userOpIntent:boolean }} opts
     */
    function isIncompleteReply(assistantContent, opts) {
        const o = opts || {};
        return !!(o.keepToolsEveryRound && o.userOpIntent && assistantContent
            && INCOMPLETE_SIGNAL.test(assistantContent));
    }

    /**
     * 零工具轮次高层决策（在覆盖度/图表闸门之后调用）。
     * @param {object} s
     * @returns {'stop'|'retry_guards'|'exempt_stop'}
     */
    function suggestZeroToolAction(s) {
        const st = s || {};
        if (st.contentTaskExempt && (!st.toolCalls || st.toolCalls.length === 0)) {
            return 'exempt_stop';
        }
        if (st.toolsExecuted && !st.isIncomplete && (!st.toolCalls || st.toolCalls.length === 0)) {
            return 'stop';
        }
        if ((!st.toolCalls || st.toolCalls.length === 0) && st.tools && st.tools.length > 0
            && !st.contentTaskExempt) {
            return 'retry_guards';
        }
        return 'stop';
    }

    App.ChatV2.agentLoopZeroToolGuards = {
        INCOMPLETE_SIGNAL,
        isIncompleteReply,
        suggestZeroToolAction,
        stripDanglingPromises,
        /** @type {null|Function} 可选：完整守卫链注入点（主循环暂未迁完） */
        decideZeroTool: null
    };
    ChatLog.info?.('CHAT-V2', '✅ agentLoopZeroToolGuards loaded');
})();
