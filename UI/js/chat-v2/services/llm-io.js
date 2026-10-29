/**
 * ============================================================
 * chat-v2/services/llm-io.js — LLM 流式消费封装
 * ============================================================
 * 职责：收敛散落在各 pipeline 的 llmClient.stream 样板（收 token、
 *   拼文本、判断 done/error）。提供统一入口：
 *   - streamText(messages, opts)  收集为纯文本
 *   - askJson(messages, opts)     自动提取 JSON 对象（strictJson）
 * 依赖：llm-client（App.ChatV2.llmClient）、json-utils
 * 收敛不改行为：事件语义与底层完全一致。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 流式对话，把 llmClient.stream 收敛为一次文本收集
     * @param {Object} opts
     *   - messages  : [{role, content}]
     *   - config    : App.config
     *   - signal    : AbortSignal
     *   - model     : 模型名（可选）
     *   - tools     : 工具 schema（可选）
     *   - nativeTools : 是否下发原生 tools（可选；缺省时由 CapabilityProfile.resolveToolMode
     *                   按通道/配置/模型能力自动决策）
     *   - onToken   : (delta) => void 可选回调
     *   - onReasoning : (delta) => void 可选回调
     * @returns {Promise<{text:string, reasoning:string, usage:object, error:object|null, aborted:boolean}>}
     */
    async function streamText(opts) {
        const llmClient = App.ChatV2.llmClient;
        const { messages, config, signal, model, tools, onToken, onReasoning } = opts || {};
        if (!llmClient || typeof llmClient.stream !== 'function') {
            return { text: '', reasoning: '', usage: null, error: { code: 'NO_LLM_CLIENT', message: 'llmClient 未就绪', recoverable: false }, aborted: false };
        }
        // 【能力适配】nativeTools 未显式指定时，按模型能力档案自动决策（单一决策点见 capability-profile.js）
        let nativeTools = opts?.nativeTools;
        if (nativeTools === undefined) {
            nativeTools = (App.CapabilityProfile && typeof App.CapabilityProfile.resolveToolMode === 'function')
                ? App.CapabilityProfile.resolveToolMode(model || config?.model, config).nativeTools
                : true;
        }
        let text = '';
        let reasoning = '';
        let usage = null;
        try {
            for await (const evt of llmClient.stream({ messages, config, model, tools, signal, nativeTools })) {
                switch (evt?.type) {
                    case 'token':
                        if (evt.delta) {
                            text += evt.delta;
                            onToken?.(evt.delta);
                        }
                        break;
                    case 'reasoning':
                        if (evt.delta) {
                            reasoning += evt.delta;
                            onReasoning?.(evt.delta);
                        }
                        break;
                    case 'usage':
                        usage = evt.tokens || evt.usage || null;
                        break;
                    case 'error':
                        return { text, reasoning, usage, error: evt.error || { code: 'LLM_ERROR', message: String(evt.error), recoverable: false }, aborted: signal?.aborted || false };
                    default:
                        break; // done / 其他事件忽略
                }
            }
        } catch (e) {
            return { text, reasoning, usage, error: { code: 'STREAM_ERROR', message: e?.message || String(e), recoverable: true }, aborted: signal?.aborted || false };
        }
        return { text, reasoning, usage, error: null, aborted: signal?.aborted || false };
    }

    /**
     * 流式对话并自动提取 JSON 对象（strictJson 模式）
     * @returns {Promise<{json:object|null, raw:string, error:object|null}>}
     */
    async function askJson(opts) {
            const result = await streamText(opts);
            const json = App.ChatV2.jsonUtils
                ? App.ChatV2.jsonUtils.extractJson(result.text)
                : (() => {
                    try {
                        const t = result.text || '';
                        const s = t.indexOf('{'); const e = t.lastIndexOf('}');
                        if (s >= 0 && e > s) try { return JSON.parse(t.substring(s, e + 1)); } catch (_) { void 0; }
                        const sa = t.indexOf('['); const ea = t.lastIndexOf(']');
                        if (sa >= 0 && ea > sa) try { return JSON.parse(t.substring(sa, ea + 1)); } catch (_) { void 0; }
                        return null;
                    } catch (_) { return null; }
                })();
            return { json, raw: result.text, error: result.error };
        }

    /**
     * 流式对话，返回纯文本回答
     * @returns {Promise<{text:string, reasoning:string, usage:object, error:object|null}>}
     */
    async function askText(opts) {
        const result = await streamText(opts);
        return { text: result.text, reasoning: result.reasoning, usage: result.usage, error: result.error };
    }

    App.ChatV2.llmIo = { streamText, askText, askJson };

    ChatLog.info?.('CHAT-V2', '✅ llmIo loaded');
})();