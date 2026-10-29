/**
 * ============================================================
 * chat-v2/services/turn-goal.js — Turn 目标与完成条件
 * ============================================================
 * 开 Turn 时用规则写一行目标和完成条件，不调用模型，也不停下来等确认。
 * 多步管道只追加约束。收尾时用已成功的工具信封验收，缺项再续一轮。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const FAMILIES = [
        { kind: 'format', re: /加粗|倾斜|斜体|居中|边框|字号|字体|底色|填充|背景|美化|格式|对齐|换行|列宽|行高/, tools: 'format|style|beautif|font|align|autofit|border' },
        { kind: 'chart', re: /图表|柱状图|折线图|饼图|条形图|组合图|画图|做图|可视化/, tools: 'create_chart|chart' },
        { kind: 'sort', re: /排序|从高到低|从低到高/, tools: 'sort' },
        { kind: 'filter', re: /筛选|过滤/, tools: 'filter' },
        { kind: 'aggregate', re: /求和|汇总|聚合|统计|合计/, tools: 'aggregate|sum|formula|subtotal' },
        { kind: 'merge', re: /合并/, tools: 'merge' },
        { kind: 'split', re: /拆分|拆成|分成/, tools: 'split' },
        { kind: 'clean', re: /清洗|去重|重复值/, tools: 'clean|dedup|duplicate' }
    ];

    // 【多目标批次】"美化工作簿中的所有工作表 / 把每个sheet都加粗"——一条消息要对 N 张表做同一操作。
    // 旧实现只有 1 个 format 完成条件：LLM 调 1 次 excel_format_range 即判"已满足"，剩下 N-1 张表
    // 被过程语言充数（实测 2026-10-22：只美化 1 张表甚至 1 个单元格就宣称任务完成）。
    // 命中多目标措辞时给完成条件打 multiTarget 标记，unmet() 改为统计【不同工作表】数量。
    const MULTI_TARGET_RE = /(所有|全部|各个|每个|每一张|每张|批量|其他|其余|剩余|剩下).{0,6}(工作表|sheet|表格|数据表)|(整个|整本|全部).{0,4}工作簿/i;

    /**
     * 用户有没有点名一套关键词预设盖不住的外观。
     * 盖得住的只有「没点名」和「商务蓝/深蓝」（后端空 format 预设本身就是它）。
     * 「统一为中国红」「改成科技绿」这类必须留给模型规划，不能在这里换成默认蓝。
     * @returns {string} 点名的外观；没有或已由预设覆盖时返回空串
     */
    function semanticStyleTarget(message) {
        const m = String(message || '');
        const verb = m.match(/(?:统一(?:为|成|到|用)|改成|改为|换成|换为|调成|调为|设置成|设为|设成|变成|涂成|刷成|转为|转成)\s*([^，。；,.\n]{1,24})/);
        let target = '';
        if (verb) target = verb[1].replace(/^(?:一个|一种|成|为)\s*/, '').trim();
        if (target && !isOwnedPresetName(target)) return target.slice(0, 24);
        const stripped = m.replace(/商务蓝|深蓝/g, '');
        if (/[红橙黄绿青紫粉金]/.test(stripped)) {
            const hit = stripped.match(/[^，。；\s]{0,10}[红橙黄绿青紫粉金]色?/);
            return (hit && hit[0]) || '指定颜色';
        }
        return '';
    }

    function isOwnedPresetName(target) {
        const t = String(target || '').trim();
        if (/^(?:商务蓝|深蓝|默认|预设|专业预设|默认预设)(?:主题|配色|风格|色系|样式)?$/.test(t)) return true;
        return /商务蓝|深蓝/.test(t) && !/[红橙黄绿青紫粉金]/.test(t.replace(/商务蓝|深蓝/g, ''));
    }

    /** 这次格式化有没有把颜色写进参数。空 {} 会被后端改成商务蓝，不能算实现了用户点名的外观。 */
    function callCarriesExplicitStyle(args) {
        if (!args || typeof args !== 'object') return false;
        let fmt = args.format;
        if (fmt == null || fmt === '') return false;
        if (typeof fmt === 'string') {
            const s = fmt.trim();
            if (!s || s === '{}' || s === 'null') return false;
            try { fmt = JSON.parse(s); } catch (_) { return true; }
        }
        if (!fmt || typeof fmt !== 'object' || Array.isArray(fmt)) return false;
        if (Object.keys(fmt).length === 0) return false;
        return !!(fmt.bgColor || fmt.fontColor || fmt.headerFontColor || fmt.borderColor
            || fmt.themeHeader || fmt.themeName || fmt.fillColor || fmt.headerFillColor);
    }

    const CHINA_RED = 'DE2910';

    function parseFormatObject(fmt) {
        if (fmt == null || fmt === '') return null;
        if (typeof fmt === 'string') {
            const s = fmt.trim();
            if (!s || s === '{}' || s === 'null') return null;
            try {
                const o = JSON.parse(s);
                return (o && typeof o === 'object' && !Array.isArray(o)) ? o : null;
            } catch (_) { return null; }
        }
        if (typeof fmt === 'object' && !Array.isArray(fmt)) return fmt;
        return null;
    }

    function hexOf(color) {
        const s = String(color || '').replace(/^#/, '').trim();
        return /^[0-9a-fA-F]{6}$/.test(s) ? s.toUpperCase() : '';
    }

    function isNearWhite(color) {
        const h = hexOf(color);
        if (!h) return false;
        return parseInt(h.slice(0, 2), 16) >= 230
            && parseInt(h.slice(2, 4), 16) >= 230
            && parseInt(h.slice(4, 6), 16) >= 230;
    }

    /** 中国红是饱和正红。F4C2C4 这种浅粉不算。 */
    function isDeepChinaRed(color) {
        const h = hexOf(color);
        if (!h) return false;
        const r = parseInt(h.slice(0, 2), 16);
        const g = parseInt(h.slice(2, 4), 16);
        const b = parseInt(h.slice(4, 6), 16);
        return r >= 160 && g <= 90 && b <= 90 && r > g + 80 && r > b + 80;
    }

    function wantsChinaRed(styleTarget) {
        return /中国红|朱红|国旗红/.test(String(styleTarget || ''));
    }

    /**
     * 模型自己规划出来的 format 里，两处会把语意做坏：
     *   浅粉冒充中国红；fontColor 写成白色（作用整段，表身字消失，看起来像数据被清空）。
     * 这里只改这两个参数，不改它选了哪张表、格式化哪一段。
     */
    function prepareFormatArgs(args, styleTarget) {
        if (!args || typeof args !== 'object') return args;
        const fmt = parseFormatObject(args.format);
        if (!fmt) return args;
        const nextFmt = Object.assign({}, fmt);
        let changed = false;
        if (wantsChinaRed(styleTarget) && !isDeepChinaRed(nextFmt.bgColor || nextFmt.themeHeader)) {
            nextFmt.bgColor = CHINA_RED;
            nextFmt.bgScope = 'header';
            delete nextFmt.themeHeader;
            changed = true;
        }
        if (isNearWhite(nextFmt.fontColor) && (nextFmt.bgScope === 'header' || wantsChinaRed(styleTarget) || !nextFmt.bgColor)) {
            nextFmt.headerFontColor = 'FFFFFF';
            nextFmt.fontColor = '000000';
            if (nextFmt.bgColor && !nextFmt.bgScope) nextFmt.bgScope = 'header';
            changed = true;
        }
        if (!changed) return args;
        const out = Object.assign({}, args);
        out.format = typeof args.format === 'string' ? JSON.stringify(nextFmt) : nextFmt;
        return out;
    }

    function callMatchesStyleTarget(args, styleTarget) {
        if (!callCarriesExplicitStyle(args)) return false;
        if (!wantsChinaRed(styleTarget)) return true;
        const fmt = parseFormatObject(args && args.format);
        return !!(fmt && isDeepChinaRed(fmt.bgColor || fmt.themeHeader));
    }

    function sheetNameList(value) {
        if (Array.isArray(value)) {
            return value.map(function (s) { return String(s || '').trim(); }).filter(Boolean);
        }
        if (typeof value === 'string' && value.trim()) {
            try {
                const parsed = JSON.parse(value);
                if (Array.isArray(parsed)) return sheetNameList(parsed);
            } catch (_) { /* 不是 JSON 就按分隔符拆 */ }
            return value.split(/[,，、]/).map(function (s) { return s.trim(); }).filter(Boolean);
        }
        return [];
    }

    /**
     * 用户说的是选中的工作表。模型经常先 get_sheets，再把全部可见表写进 sourceSheets。
     * 这里只改合并范围，不改它选的目标表名。说了「所有/全部」时不动。
     */
    function constrainMergeSource(args, message, selectedSheets) {
        if (!args || typeof args !== 'object') return args;
        const msg = String(message || '');
        if (/(所有|全部|每张|各个|各张)/.test(msg)) return args;
        if (!/选中|选定|勾选/.test(msg)) return args;
        const sel = sheetNameList(selectedSheets);
        if (sel.length < 2) return args;
        const current = sheetNameList(args.sourceSheets);
        const same = current.length === sel.length && sel.every(function (name, i) { return current[i] === name; });
        const mergeAllOn = args.mergeAll === true || args.mergeAll === 'true';
        if (same && !mergeAllOn) return args;
        const out = Object.assign({}, args);
        out.sourceSheets = Array.isArray(args.sourceSheets) || args.sourceSheets == null ? sel.slice() : JSON.stringify(sel);
        out.mergeAll = false;
        return out;
    }

    function extractConstraints(text) {
        const out = [];
        const re = /(?:不要|别|禁止|勿|不能|保留)[^，。；\n]{0,24}/g;
        let m;
        while ((m = re.exec(text))) {
            const s = m[0].trim();
            if (s.length >= 2 && s.length <= 30 && out.indexOf(s) < 0) out.push(s);
            if (out.length >= 4) break;
        }
        return out;
    }

    function splitClauses(text) {
        return String(text || '')
            .split(/[；;]|(?:，(?:并且|同时|还要|另外|以及))|(?:\d+[.、．])/)
            .map(s => s.trim())
            .filter(s => s.length >= 2);
    }

    function build(message) {
        const text = String(message || '').trim();
        const constraints = extractConstraints(text);
        const clauses = splitClauses(text);
        const pieces = clauses.length >= 2 ? clauses : [text];
        const doneWhen = [];
        const seen = new Set();
        for (const piece of pieces) {
            if (!piece) continue;
            for (const fam of FAMILIES) {
                const key = fam.kind + '\n' + piece;
                if (!fam.re.test(piece) || seen.has(key)) continue;
                seen.add(key);
                doneWhen.push({
                    kind: fam.kind,
                    label: piece.slice(0, 40),
                    tools: fam.tools
                });
            }
        }
        const rangeMatch = text.match(/\$?[A-Za-z]{1,3}\$?\d{1,7}(?:\s*[:：]\s*\$?[A-Za-z]{1,3}\$?\d{1,7})?/);
        const range = rangeMatch ? rangeMatch[0] : '';
        if (range) {
            for (const item of doneWhen) {
                if (item.kind === 'format' && item.label.indexOf(range) < 0) {
                    item.label = (item.label + ' ' + range).slice(0, 48);
                }
            }
        }
        if (!doneWhen.length && text) {
            doneWhen.push({ kind: 'reply', label: text.slice(0, 80), tools: '' });
        }
        const styleTarget = semanticStyleTarget(text);
        if (styleTarget) {
            const line = '外观按「' + styleTarget + '」实现，不要换成默认商务蓝';
            if (constraints.indexOf(line) < 0) constraints.push(line);
        }
        // 「合并，并设置成中国红」只命中合并族，格式族词表没有「中国红」。
        // 不补这一条的话，合并做完覆盖度就是 0，颜色停在口头承诺上。
        // 必须写在 multiTarget 标记之前，否则「所有工作表统一为商务蓝」不会按每张表验收。
        const namedLook = styleTarget || (/商务蓝|深蓝/.test(text) ? '商务蓝' : '');
        if (namedLook && !doneWhen.some(d => d && d.kind === 'format')) {
            doneWhen.push({
                kind: 'format',
                label: '外观按「' + namedLook + '」设置',
                tools: 'format|style|beautif|font|align|autofit|border'
            });
        }
        // 多目标批次：完成条件升级为"每张表各做一次"（unmet 按不同工作表计数）
        if (MULTI_TARGET_RE.test(text)) {
            doneWhen.forEach(d => { if (d && d.kind !== 'reply') d.multiTarget = true; });
        }
        return {
            text: text.slice(0, 120),
            range,
            constraints,
            doneWhen,
            styleTarget,
            multiStep: false,
            source: 'rule'
        };
    }

    function markMultiStep(goal, message) {
        const g = (goal && typeof goal === 'object') ? goal : build(message);
        g.multiStep = true;
        g.source = String(g.source || 'rule').indexOf('multistep') >= 0 ? g.source : (g.source || 'rule') + '+multistep';
        if (!Array.isArray(g.constraints)) g.constraints = [];
        const line = '未完成下列各项不得声称已完成';
        if (g.constraints.indexOf(line) < 0) g.constraints.push(line);
        return g;
    }

    /** 从 excel_get_sheets 的返回里取出可见表。没有清单时返回空数组。 */
    function sheetRowsFromToolResult(raw) {
        let data = raw;
        for (let d = 0; d < 6; d++) {
            if (data == null) return [];
            if (typeof data === 'string') {
                const s = data.trim();
                if (!s || (s.charAt(0) !== '[' && s.charAt(0) !== '{')) return [];
                try { data = JSON.parse(s); } catch (e) { return []; }
                continue;
            }
            if (Array.isArray(data)) {
                const rows = [];
                for (const x of data) {
                    if (!x || typeof x !== 'object') continue;
                    const name = String(x.name || x.sheet || '').trim();
                    if (!name) continue;
                    rows.push({
                        name: name,
                        usedRange: String(x.usedRange || x.range || '').replace(/\$/g, '').trim()
                    });
                }
                return rows;
            }
            if (typeof data === 'object') {
                if (Array.isArray(data.content)) {
                    data = data.content.map(function (c) { return c && c.text; }).filter(Boolean).join('\n');
                    continue;
                }
                const nxt = data.result != null ? data.result : (data.data != null ? data.data : null);
                if (nxt == null) return [];
                data = nxt;
                continue;
            }
            return [];
        }
        return [];
    }

    function listedSheets(turn) {
        const calls = Array.isArray(turn && turn.toolCalls) ? turn.toolCalls : [];
        for (let i = calls.length - 1; i >= 0; i--) {
            const c = calls[i];
            const name = String((c && c.name) || '').replace(/^excel-local__/, '').replace(/^mcp_/, '');
            if (!c || !/get_sheets$/i.test(name)) continue;
            if (c.status && c.status !== 'completed') continue;
            const rows = sheetRowsFromToolResult(c.result);
            if (rows.length) return rows;
        }
        return [];
    }

    function isMultiSheetFormatGoal(turn) {
        const goal = turn && turn.goal;
        return !!(goal && Array.isArray(goal.doneWhen)
            && goal.doneWhen.some(function (d) { return d && d.multiTarget && d.kind === 'format'; }));
    }

    /** 已成功套上目标外观的工作表名。失败调用不算。 */
    function formattedSheetSet(turn) {
        const goal = turn && turn.goal;
        const needStyle = !!(goal && goal.styleTarget);
        const calls = Array.isArray(turn && turn.toolCalls) ? turn.toolCalls : [];
        const set = new Set();
        for (const c of calls) {
            if (!c || c.status !== 'completed' || !/format_range/i.test(String(c.name || ''))) continue;
            if (needStyle && !callMatchesStyleTarget(c.args, goal.styleTarget)) continue;
            const sh = c.args && typeof c.args.sheet === 'string' ? c.args.sheet.trim() : '';
            if (sh) set.add(sh);
        }
        return set;
    }

    /**
     * 「所有工作表」且本轮已经拿到表清单时，还没格式化的表。
     * 没有清单时返回空数组，完成条件仍按「至少两张不同的表」判断。
     */
    function missingFormatSheets(turn) {
        if (!isMultiSheetFormatGoal(turn)) return [];
        const listed = listedSheets(turn);
        if (!listed.length) return [];
        const done = formattedSheetSet(turn);
        return listed.filter(function (s) { return s.name && !done.has(s.name); });
    }

    /** 模型没写总结时用的一句事实，不把工具原始 JSON 摊进正文。 */
    function formatCoverageText(turn) {
        if (!isMultiSheetFormatGoal(turn)) return '';
        const listed = listedSheets(turn);
        if (!listed.length) return '';
        const missing = missingFormatSheets(turn);
        const miss = {};
        missing.forEach(function (s) { miss[s.name] = true; });
        const done = listed.filter(function (s) { return !miss[s.name]; }).map(function (s) { return s.name; });
        const parts = [];
        if (done.length) parts.push('已格式化 ' + done.length + ' 张工作表：' + done.join('、') + '。');
        if (missing.length) parts.push('以下工作表本次没有格式化：' + missing.map(function (s) { return s.name; }).join('、') + '。');
        return parts.join('\n');
    }

    /** 只把「还没成功执行」的操作类完成条件算作缺项。纯问答不缺。 */
    function unmet(turn) {
        const goal = turn && turn.goal;
        if (!goal || !Array.isArray(goal.doneWhen)) return [];
        const envelopes = Array.isArray(turn._envelopes) ? turn._envelopes : null;
        // 工具名池：信封 > _executedTools > _executedToolCalls（后者带参数，是多目标计数的依据；
        // 旧会话/部分路径只有 calls 时也要能判"同类工具执行过"）
        const pool = envelopes
            ? envelopes.filter(e => e && e.ok).map(e => String(e.tool || '')).join(' ')
            : (Array.isArray(turn._executedTools) ? turn._executedTools.map(n => String(n || '')).join(' ')
                : (Array.isArray(turn._executedToolCalls) ? turn._executedToolCalls.map(c => String((c && c.name) || '')).join(' ') : ''));
        const out = [];
        for (const item of goal.doneWhen) {
            if (!item || item.kind === 'reply') continue;
            if (item.kind === 'chart' && turn._chartDone) continue;
            // 【多目标批次】"美化所有工作表"：只有 1 次同工具调用不算做完——
            // 统计已执行的同类工具覆盖了多少张【不同工作表】，<2 即未完成
            // （弱模型典型行为：调 1 次 format_range 就用过程语言收尾）。
            if (item.multiTarget && item.tools) {
                const _listedMissing = (item.kind === 'format') ? missingFormatSheets(turn) : [];
                if (_listedMissing.length) {
                    out.push(Object.assign({}, item, {
                        label: (item.label || '格式化') + '（还差 ' + _listedMissing.map(function (s) { return s.name; }).join('、') + '）'
                    }));
                    continue;
                }
                const _calls = Array.isArray(turn._executedToolCalls) ? turn._executedToolCalls : [];
                const _sheets = new Set();
                let _familyRe = null;
                try { _familyRe = new RegExp(item.tools, 'i'); } catch (_re) { _familyRe = null; }
                const _needStyle = !!(goal.styleTarget);
                for (const c of _calls) {
                    if (!c || !_familyRe || !_familyRe.test(String(c.name || ''))) continue;
                    if (_needStyle && !callMatchesStyleTarget(c.args, goal.styleTarget)) continue;
                    const _sh = c && c.args && typeof c.args.sheet === 'string' ? c.args.sheet.trim() : '';
                    if (_sh) _sheets.add(_sh);
                }
                // 本轮没有表清单时，仍只要求至少两张不同的表。有清单时上面已经按全表核对。
                if (!listedSheets(turn).length && _sheets.size < 2) { out.push(item); continue; }
                continue;
            }
            let hit = false;
            if (item.tools) {
                try { hit = new RegExp(item.tools, 'i').test(pool); } catch (_) { hit = false; }
            }
            // 点了颜色时，光有一次 format 名字不够：空 {} 会被执行成商务蓝。
            // 多目标批次上面已经按 callMatchesStyleTarget 计过，这里只补单次。
            if (hit && item.kind === 'format' && goal.styleTarget && !item.multiTarget) {
                const calls = Array.isArray(turn._executedToolCalls) ? turn._executedToolCalls : [];
                if (calls.length) {
                    let re = null;
                    try { re = new RegExp(item.tools, 'i'); } catch (_) { re = null; }
                    hit = calls.some(function (c) {
                        return c && re && re.test(String(c.name || ''))
                            && callMatchesStyleTarget(c.args, goal.styleTarget);
                    });
                }
            }
            if (!hit) out.push(item);
        }
        return out;
    }

    function goalPrompt(goal) {
        if (!goal) return '';
        const lines = ['[本轮目标] ' + (goal.text || '')];
        if (goal.constraints && goal.constraints.length) lines.push('[约束] ' + goal.constraints.join('；'));
        const checks = (goal.doneWhen || []).filter(d => d && d.kind !== 'reply');
        if (checks.length) lines.push('[完成条件] ' + checks.map(d => d.label).join('；'));
        if (goal.styleTarget) {
            lines.push('[外观] 用户指定「' + goal.styleTarget + '」。每次 excel_format_range 的 format 都要写出这个颜色。'
                + '禁止 format 使用 {}：空对象会被执行成商务蓝。先列出全部工作表，再逐表做。');
            if (wantsChinaRed(goal.styleTarget)) {
                lines.push('[中国红] 用饱和正红 #DE2910 做表头底色（bgColor=DE2910，bgScope=header）。'
                    + '不要用浅红、粉色或「商务红」（例如 #F4C2C4）。'
                    + '表头白字写 headerFontColor=FFFFFF。禁止写 fontColor=FFFFFF，那会把整段文字涂白，表身数据会看不见。');
            }
        }
        lines.push('收尾时只根据工具结果里的 [验收] 行描述成败，不得把失败或未完成说成已完成。');
        return lines.join('\n');
    }

    App.ChatV2.turnGoal = {
        build, markMultiStep, unmet, goalPrompt, semanticStyleTarget,
        callCarriesExplicitStyle, prepareFormatArgs, callMatchesStyleTarget, constrainMergeSource, CHINA_RED,
        listedSheets, missingFormatSheets, formatCoverageText
    };
})();
