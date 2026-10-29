/**
 * ============================================================
 * chat-v2/services/json-utils.js — 统一 JSON 解析工具
 * ============================================================
 * 职责：收敛散落在各 pipeline 的 JSON 解析实现。
 *   - safeParse      严格容错解析（对应旧 python-pipeline safeParse）
 *   - extractJson    从任意文本提取首个 JSON 对象/数组（对应旧 extractJson/extractJSON）
 *   - extractJsonObject  强制只提取对象（对应旧 extractJSON）
 *   - toRecords      把后端结果 data 规约为可消费记录数组
 * 语义与各 pipeline 原实现保持一致，收敛不改行为。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 严格容错解析：能解析则返回对象/数组，否则返回 null
     */
    function safeParse(str) {
        try { return JSON.parse(str); } catch (e) { return null; }
    }

    /**
     * 从执行结果文本中提取 JSON（支持对象或数组，兼容前缀文本）
     * 若输入本身已是对象/数组，直接返回。
     */
    function extractJson(text) {
        if (!text) return null;
        if (typeof text === 'object') return text;
        let t = String(text).trim();
        // 判断内容是否以 [ 开头（考虑代码块包裹）
        let inner = t;
        let blockMatch = t.match(/```(?:json)?\s*([\s\S]*?)```/);
        if (blockMatch) inner = blockMatch[1].trim();
        const startsWithArray = inner.startsWith('[');

        if (startsWithArray) {
            // 优先尝试数组
            let sa = t.indexOf('[');
            let ea = t.lastIndexOf(']');
            if (sa >= 0 && ea > sa) {
                try { return JSON.parse(t.substring(sa, ea + 1)); } catch(e) { void 0; }
            }
        }
        // 尝试对象
        let start = t.indexOf('{');
        let end = t.lastIndexOf('}');
        if (start >= 0 && end > start) {
            let candidate = t.substring(start, end + 1);
            try { return JSON.parse(candidate); } catch(e) { void 0; }
        }
        // 尝试数组（非开头的情况）
        if (!startsWithArray) {
            let sa = t.indexOf('[');
            let ea = t.lastIndexOf(']');
            if (sa >= 0 && ea > sa) {
                try { return JSON.parse(t.substring(sa, ea + 1)); } catch(e) { void 0; }
            }
        }
        return null;
    }

    /**
     * 只提取 JSON 对象（支持 ```json 代码块包裹）；提取不到返回 null
     */
    function extractJsonObject(text) {
        if (!text) return null;
        let blockMatch = String(text).match(/```(?:json)?\s*([\s\S]*?)```/);
        let content = blockMatch ? blockMatch[1] : String(text);
        let start = content.indexOf('{');
        let end = content.lastIndexOf('}');
        if (start === -1 || end === -1) return null;
        try {
            return JSON.parse(content.substring(start, end + 1));
        } catch (e) {
            return null;
        }
    }

    /**
     * 把后端返回的 data / 执行结果 result 规约为可消费的记录数组
     * @param {*} data - 后端 data（数组 / {rows} / 已 JSON.stringify 的字符串等）
     * @returns {Array|null} 规范化后的记录数组
     */
    function toRecords(data) {
        if (data == null) return null;
        let d = data;
        if (typeof d === 'string') {
            d = safeParse(d);
            if (d == null) return null;
        }
        if (Array.isArray(d)) return d;
        if (Array.isArray(d.data)) return d.data;
        if (Array.isArray(d.rows)) return d.rows;
        if (Array.isArray(d.result)) return d.result;
        return null;
    }

    App.ChatV2.jsonUtils = { safeParse, extractJson, extractJsonObject, toRecords };

    ChatLog.info?.('CHAT-V2', '✅ jsonUtils loaded');
})();