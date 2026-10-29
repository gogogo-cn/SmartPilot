/**
 * ============================================================
 * chat-v2/services/tool-envelope.js — 工具结果信封
 * ============================================================
 * 出口只认三样：value（给人看的结果）、context（下一轮用的事实）、error（非空即失败）。
 * invoke / dispatcher 在边界上收成这一个结构；回喂模型、熔断、收尾验收都读它。
 * 原始结果仍留在 result 上，表格和图表卡片继续用原始数据。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const CONTEXT_KEYS = [
        'address', 'range', 'dataRange', 'sheet', 'sheetName', 'targetSheet',
        'outputSheet', 'resultSheet', 'newSheet', 'chartName', 'anchor',
        'suggestedNextAnchor', 'rowCount', 'colCount', 'headers'
    ];

    function unwrap(raw) {
        let data = raw;
        let depth = 0;
        while (data && typeof data === 'object' && depth < 5) {
            depth++;
            if (typeof data.result === 'string') {
                try {
                    const parsed = JSON.parse(data.result);
                    if (parsed && typeof parsed === 'object') { data = parsed; continue; }
                } catch (_) { data = data.result; break; }
            }
            if (data.result && typeof data.result === 'object') { data = data.result; continue; }
            if (Array.isArray(data.content)) {
                const texts = data.content.filter(c => c && c.type === 'text' && typeof c.text === 'string').map(c => c.text);
                if (texts.length) {
                    const combined = texts.join('\n');
                    try { data = JSON.parse(combined); } catch (_) { data = combined; }
                    continue;
                }
            }
            break;
        }
        if (typeof data === 'string') {
            try {
                const parsed = JSON.parse(data);
                if (parsed && typeof parsed === 'object') data = parsed;
            } catch (_) { /* 保持原文 */ }
        }
        return data;
    }

    function pickError(data) {
        if (!data || typeof data !== 'object') return '';
        for (const k of ['error', 'Error', 'errMsg', 'errorMessage', 'message']) {
            const v = data[k];
            if (typeof v === 'string' && v.trim() && /错误|失败|不能为空|invalid|error|fail/i.test(v)) return v.trim();
        }
        if (typeof data.error === 'string' && data.error.trim()) return data.error.trim();
        return '';
    }

    function pickValue(data) {
        if (data == null) return '';
        if (typeof data === 'string') return data;
        if (typeof data !== 'object') return data;
        for (const k of ['message', 'summary', 'resultText', 'text']) {
            if (typeof data[k] === 'string' && data[k].trim()) return data[k].trim();
        }
        return data;
    }

    function collectContext(data, into) {
        if (!data || typeof data !== 'object' || Array.isArray(data)) return;
        for (const k of CONTEXT_KEYS) {
            if (into[k] == null && data[k] != null && data[k] !== '') into[k] = data[k];
        }
        if (!into.artifact && data.artifact && typeof data.artifact === 'object') into.artifact = data.artifact;
    }

    function bodyFailed(data) {
        return !!(data && typeof data === 'object' && !Array.isArray(data)
            && (data.success === false || data.ok === false || data.isError === true));
    }

    /**
     * @param {string} toolName
     * @param {{ok?:boolean, result?:*, error?:string}} invokeResult
     * @returns {{ok:boolean, value:*, context:object, error:string}}
     */
    function fromInvoke(toolName, invokeResult) {
        const src = invokeResult || {};
        const data = unwrap(src.result);
        const context = {};
        collectContext(data, context);
        collectContext(src.result, context);
        if (Array.isArray(context.headers)) context.headers = context.headers.slice(0, 12);
        let error = '';
        const explicit = src.error ? String(src.error).trim() : '';
        const failed = src.ok === false || bodyFailed(data);
        if (failed) error = explicit || pickError(data) || pickError(src.result) || '执行失败';
        return {
            ok: !failed && !error,
            value: pickValue(data),
            context,
            error,
            tool: toolName || ''
        };
    }

    function fromRaw(toolName, raw, ok) {
        const data = unwrap(raw);
        const failed = ok === false || bodyFailed(data);
        return fromInvoke(toolName, {
            ok: !failed,
            result: raw,
            error: failed ? (pickError(data) || pickError(raw) || '') : ''
        });
    }

    function slimContext(ctx) {
        if (!ctx || typeof ctx !== 'object') return '';
        const o = {};
        for (const k of CONTEXT_KEYS) {
            if (ctx[k] == null || ctx[k] === '') continue;
            o[k] = k === 'headers' && Array.isArray(ctx[k]) ? ctx[k].slice(0, 12) : ctx[k];
        }
        const s = JSON.stringify(o);
        return s === '{}' ? '' : s;
    }

    /** 回喂模型时加在原文前面，成败和事实与原始 JSON 分开。 */
    function prefixFeed(envelope, body) {
        if (!envelope) return body == null ? '' : String(body);
        const lines = [envelope.ok ? '[验收] 成功' : '[验收] 失败'];
        if (envelope.error) lines.push('[错误] ' + String(envelope.error).slice(0, 300));
        const ctx = slimContext(envelope.context);
        if (ctx) lines.push('[上下文] ' + ctx);
        return lines.join('\n') + '\n' + (body == null ? '' : String(body));
    }

    function looksLikeRawJson(text) {
        const t = String(text || '').trim();
        if (!/^[{[]/.test(t)) return false;
        return /"success"|"ok"|mergedRows|sourceSheets|"rows"|"headers"|"error"/i.test(t);
    }

    function formatForUser(envelope, toolLabel) {
        if (!envelope) return '';
        if (!envelope.ok) return '未完成：' + (envelope.error || '执行失败');
        if (typeof envelope.value === 'string' && envelope.value && !looksLikeRawJson(envelope.value)) {
            return envelope.value.length > 500 ? envelope.value.slice(0, 500) : envelope.value;
        }
        const bits = [];
        const ctx = envelope.context || {};
        const where = ctx.address || ctx.range || ctx.dataRange || '';
        if (where) bits.push(String(where));
        if (ctx.sheet || ctx.sheetName || ctx.targetSheet) bits.push(String(ctx.sheet || ctx.sheetName || ctx.targetSheet));
        if (ctx.chartName) bits.push('图表 ' + ctx.chartName);
        const name = toolLabel || envelope.tool || '操作';
        return '已执行 ' + name + (bits.length ? '（' + bits.join('，') + '）' : '') + '。';
    }

    function buildVerdict(turn) {
        const envelopes = Array.isArray(turn && turn._envelopes) ? turn._envelopes : [];
        const failed = envelopes.filter(e => e && e.ok === false);
        const done = envelopes.filter(e => e && e.ok === true);
        let unmet = [];
        try {
            const goalApi = App.ChatV2.turnGoal;
            if (goalApi && typeof goalApi.unmet === 'function') unmet = goalApi.unmet(turn) || [];
        } catch (_) { unmet = []; }
        const lines = [];
        if (done.length) {
            lines.push('已完成：' + done.map(e => e.tool || '操作').join('、'));
        }
        if (failed.length) {
            lines.push('失败：' + failed.map(e => (e.tool || '工具') + ' — ' + (e.error || '失败')).join('；'));
        }
        if (unmet.length) {
            lines.push('尚未完成：' + unmet.map(u => u.label || u.kind).join('；'));
        }
        if (!lines.length) lines.push('没有新的工具结果。');
        return {
            ok: failed.length === 0 && unmet.length === 0,
            done,
            failed,
            unmet,
            text: lines.join('\n'),
            failureNote: failed.length ? ('失败：' + failed.map(e => e.error || e.tool).join('；')) : '',
            unmetNote: unmet.length
                ? ('尚未完成：' + unmet.map(u => u.label || u.kind).join('；') + '。以上内容本次没有执行。')
                : ''
        };
    }

    /**
     * 人话只能复述验收结论。模型把失败说成已完成、或把工具 JSON 当正文时，换成确定性文本。
     */
    function applyVerdict(narration, verdict) {
        if (!verdict) return narration == null ? '' : String(narration);
        let text = String(narration || '').trim();
        if (looksLikeRawJson(text)) text = verdict.text;
        const bad = (verdict.failed && verdict.failed.length) || (verdict.unmet && verdict.unmet.length);
        if (bad && /已完成|已成功|全部完成|任务完成|成功完成/.test(text)) text = verdict.text;
        if (verdict.failed && verdict.failed.length && !/失败|未完成|没有执行/.test(text)) {
            text = (text ? text + '\n\n' : '') + verdict.failureNote;
        }
        if (verdict.unmet && verdict.unmet.length && !/没有执行|尚未完成/.test(text)) {
            text = (text ? text + '\n\n' : '') + verdict.unmetNote;
        }
        if (!text) text = verdict.text;
        return text;
    }

    App.ChatV2.toolEnvelope = {
        fromInvoke,
        fromRaw,
        prefixFeed,
        formatForUser,
        buildVerdict,
        applyVerdict,
        looksLikeRawJson,
        unwrap
    };
})();
