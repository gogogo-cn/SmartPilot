/**
 * ============================================================
 * chat-v2/services/range-resolver.js — 选区范围判断与降级解析（单一事实源）
 * ============================================================
 * 背景（实测 2026-10-22 用户截图）：
 *   批量美化跑通后暴露真正的问题：7 张表里 4 张只格式化了 G30 单格、3 张没有范围。
 *   根因是 Excel COM 的 UsedRange 维度缓存【对从未激活过的工作表不刷新】——
 *   拆分/新建出来的表会返回陈旧值（光标所在单格 G30 或空串；用户点开过的表是准的）。
 *   前端若原样传给 excel_format_range → "只美化 1 个格子还报成功"；
 *   若改传空 range + autoRange，后端又会用"当前选区"兜底，把用户停在别的表上的
 *   选区格式化掉（silently 做错表，还报表名）。
 *
 * 本模块统一回答三个问题（agent-loop / skill / dispatcher 共用，杜绝各写一套）：
 *   ① 一个地址能不能当"整张数据表"的目标？  → isReliableTableRange
 *   ② 一个地址是不是"光标单格"？            → isSingleCellAddr
 *   ③ 拿不到可靠范围时怎么逐级降级解析？    → createSheetRangeLadder
 *
 * 降级阶梯（每级都过 isReliableTableRange 校验，不可靠才进下一级）：
 *   ① seedRange（excel_get_sheets 的 usedRange）
 *   ② excel_get_used_range {sheet}（仍是 UsedRange，但对已激活表是准的）
 *   ③ excel_read_range {sheet, range:'A1', autoProbe:true, maxRows:1}
 *      —— 后端对单格锚点做 CurrentRegion 连续区扩展，【不依赖工作表激活态】，
 *         是破解 UsedRange 缓存不刷新的关键一跳；maxRows=1 只取地址不搬数据。
 *   ④ 三级全不可靠 → 返回 null：调用方必须【跳过该表并说明原因】——
 *      不猜范围、不拿单格硬调、不空 range 直调。
 *
 * 注册名：App.ChatV2.rangeResolver（IIFE，无依赖，须早于 agent-loop /
 *         mcp-dispatcher / skill-spec-build 加载）。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    /** 列字母 → 列号（A=1…Z=26、AA=27…） */
    function colToNum(col) {
        const s = String(col || '').toUpperCase().replace(/[^A-Z]/g, '');
        let n = 0;
        for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
        return n;
    }

    /**
     * 归一化范围地址：去 $、去空白、去 sheet 前缀。
     * "Sheet1!$A$1:$B$2" → "A1:B2"；"$G$30" → "G30"。
     */
    function normRangeAddr(addr) {
        return String(addr || '').replace(/\$/g, '').trim().replace(/^[^!]*!/, '').trim();
    }

    /**
     * 是否"光标单格"（单格/空/畸形——即不能当作数据区使用）。
     * Excel 任何时刻都有活动单元格，所以单格选区一律视为"光标停留"而非用户圈定的数据表
     * （与 dispatcher 的 _selIsRange、reference-area-manager 的单格忽略同一语义）。
     * ⚠ GetSelection 返回的 address 默认带 $（"$G$30"），必须先归一化再判——
     *   早期 agent-loop 用 /^[A-Za-z]+\d+$/ 直接测带 $ 的地址，单格全被漏判成"有效范围"。
     * @param {string} addr
     * @returns {boolean} true=单格/空/畸形（不是多格区域）
     */
    function isSingleCellAddr(addr) {
        const s = normRangeAddr(addr);
        if (!s) return true;
        return s.indexOf(':') < 0 && s.indexOf(',') < 0;
    }

    /**
     * 【选区范围判断·第一跳】判断一个"已用范围"能否作为"美化/格式化整张表"的目标。
     * 要求：X#:X# 形式（单格/空/畸形不可靠）且行数 ≥ 2（至少"表头 + 1 行数据"）。
     * @param {string} addr 形如 "A1:I21" / "G30" / ""
     * @returns {boolean} true=可作为整表美化目标
     */
    function isReliableTableRange(addr) {
        const s = normRangeAddr(addr);
        if (!s) return false;
        const m = s.match(/^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/);
        if (!m) return false;                                    // 单格（无冒号）或畸形 → 不可靠
        const rows = (+m[4]) - (+m[2]) + 1;
        const cols = colToNum(m[3]) - colToNum(m[1]) + 1;
        return rows >= 2 && cols >= 1;
    }

    /**
     * 从 excel_get_used_range 的返回里取 range/usedRange 地址（去 $）。
     * 逐层下钻：字符串 JSON → {range|usedRange} → result → content[0].text。
     */
    function usedRangeAddr(raw) {
        let o = raw;
        for (let i = 0; i < 4; i++) {
            if (o == null) return null;
            if (typeof o === 'string') {
                const s = o.trim();
                if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return null;
                try { o = JSON.parse(s); } catch (_e) { return null; }
            }
            if (typeof o !== 'object' || Array.isArray(o)) return null;
            if (typeof o.range === 'string' && /[A-Za-z]/.test(o.range)) return o.range.replace(/\$/g, '');
            if (typeof o.usedRange === 'string' && /[A-Za-z]/.test(o.usedRange)) return o.usedRange.replace(/\$/g, '');
            if (o.result != null) { o = o.result; continue; }
            if (Array.isArray(o.content)) {
                const t = o.content.find(function (c) { return c && typeof c.text === 'string'; });
                if (t) { o = t.text; continue; }
            }
            return null;
        }
        return null;
    }

    /**
     * 从 excel_read_range 的返回里取解析后的 range 地址。
     * 与 usedRangeAddr 的差别：read_range 数据被 maxRows 截断时会在 JSON 尾部追加
     * "/* ⚠️ 数据已截断… * /" 注释（JSON.parse 直接失败），这里切到最后一个 '}' 再解析。
     */
    function readRangeAddr(raw) {
        let o = raw;
        for (let i = 0; i < 4; i++) {
            if (o == null) return null;
            if (typeof o === 'string') {
                const s = o.trim();
                if (!s || s.charAt(0) !== '{') return null;
                let parsed = null;
                try { parsed = JSON.parse(s); } catch (_e1) {
                    const end = s.lastIndexOf('}');
                    if (end > 0) { try { parsed = JSON.parse(s.slice(0, end + 1)); } catch (_e2) { return null; } }
                    if (!parsed) return null;
                }
                o = parsed;
            }
            if (typeof o !== 'object' || Array.isArray(o)) return null;
            if (typeof o.range === 'string' && /[A-Za-z]/.test(o.range)) return o.range.replace(/\$/g, '');
            if (o.result != null) { o = o.result; continue; }
            if (Array.isArray(o.content)) {
                const t = o.content.find(function (c) { return c && typeof c.text === 'string'; });
                if (t) { o = t.text; continue; }
            }
            return null;
        }
        return null;
    }

    /**
     * 创建"逐级降级解析一张表数据范围"的解析器。
     * @param {Function} callTool (toolName, argsJson) => Promise<{raw}> 工具调用器
     *        （dispatcher 传 callLocalToolWithFallback；测试可直接传桩）
     * @returns {Function} async (sheetName, seedRange) => {range, via} | null
     *          via ∈ 'get_sheets' | 'get_used_range' | 'read_range:CurrentRegion'
     */
    function createSheetRangeLadder(callTool) {
        const call = (typeof callTool === 'function') ? callTool : (async () => ({ raw: null }));
        return async function resolveSheetRangeLadder(sheetName, seedRange) {
            // ① seedRange（get_sheets 的 usedRange）
            if (isReliableTableRange(seedRange)) {
                return { range: normRangeAddr(seedRange), via: 'get_sheets' };
            }
            // ② get_used_range
            try {
                const got = await call('excel_get_used_range', JSON.stringify({ sheet: sheetName }));
                const a = usedRangeAddr(got && got.raw);
                if (isReliableTableRange(a)) {
                    return { range: normRangeAddr(a), via: 'get_used_range' };
                }
            } catch (_e2) { /* 进入下一级 */ }
            // ③ read_range + CurrentRegion（autoProbe 显式传，免得对端改默认）
            try {
                const got = await call('excel_read_range',
                    JSON.stringify({ sheet: sheetName, range: 'A1', autoProbe: true, maxRows: 1 }));
                const a = readRangeAddr(got && got.raw);
                if (isReliableTableRange(a)) {
                    return { range: normRangeAddr(a), via: 'read_range:CurrentRegion' };
                }
            } catch (_e3) { /* 进入跳过分支 */ }
            // ④ 三级全不可靠 → null（调用方跳过该表并说明原因）
            return null;
        };
    }

    /**
     * 选区形状。Excel 永远有活动单元格，所以单格是光标，不是用户圈定的区域。
     * 自动捕获的已用区域也不是手选。整表被选中（几百行/几十列）同样不当成圈定。
     */
    function classifySelection(sel) {
        if (!sel || typeof sel !== 'object') return { kind: 'none', address: '', sheet: '' };
        const sheet = String(sel.sheetName || sel.sheet || sel.worksheet || '');
        if (sel.objectType && sel.objectType !== 'range' && sel.objectType !== 'cell') {
            return { kind: 'object', address: '', sheet: sheet, objectType: String(sel.objectType) };
        }
        const addr = normRangeAddr(sel.address || sel.range || '');
        if (sel.isAutoCaptured === true) return { kind: 'auto', address: addr, sheet: sheet };
        if (!addr) return { kind: 'none', address: '', sheet: sheet };
        if (addr.indexOf(',') >= 0) return { kind: 'areas', address: addr, sheet: sheet };
        if (isSingleCellAddr(addr)) return { kind: 'cell', address: addr, sheet: sheet };
        if (/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(addr)) return { kind: 'col', address: addr, sheet: sheet };
        if (/^\d+:\d+$/.test(addr)) return { kind: 'row', address: addr, sheet: sheet };
        const m = addr.match(/^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})(\d+)$/);
        if (!m) return { kind: 'other', address: addr, sheet: sheet };
        const rows = (+m[4]) - (+m[2]) + 1;
        const cols = colToNum(m[3]) - colToNum(m[1]) + 1;
        if (rows === 1 && cols === 1) return { kind: 'cell', address: addr, sheet: sheet };
        if (rows > 500 || cols > 40) return { kind: 'whole', address: addr, sheet: sheet };
        if (rows === 1) return { kind: 'row', address: addr, sheet: sheet };
        if (cols === 1) return { kind: 'col', address: addr, sheet: sheet };
        return { kind: 'block', address: addr, sheet: sheet };
    }

    /** 用户嘴里说出的单元格地址。没有就返回空串，绝不拿光标格充数。 */
    function extractSpokenRange(message) {
        const m = String(message || '');
        const hit = m.match(/(?:^|[^A-Za-z0-9])((?:'[^']+'!|[A-Za-z0-9_\u4e00-\u9fa5]+!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)/);
        return hit ? normRangeAddr(hit[1]) : '';
    }

    function _sheetNamesIn(message, context) {
        const m = String(message || '');
        const named = [];
        const q = m.match(/工作表\s*[「“"']([^」”"']+)[」”"']/);
        if (q && q[1].trim()) named.push(q[1].trim());
        const list = context && Array.isArray(context.sheets) ? context.sheets : [];
        list.forEach(function (s) {
            const name = typeof s === 'string' ? s : (s && s.name);
            if (name && String(name).length >= 2 && m.indexOf(name) >= 0 && named.indexOf(name) < 0) named.push(name);
        });
        return named;
    }

    function _realRegion(kind) {
        return kind === 'row' || kind === 'block' || kind === 'col' || kind === 'areas';
    }

    /**
     * 格式化该打在哪。颜色由模型决定，范围只按这句话和选区决定。
     *
     * sheetScope: all 每张可见表 / others 除活动表外 / named 点了名的表 / active 当前表 / none 否定
     * rangeMode:
     *   spoken          嘴里的地址，原样用，不换成数据区
     *   selection       用户圈定的一行、一列或一块（单格不算）
     *   cell            明确说了「这个单元格」
     *   copy-selection  「其他表也按这个区域」，且圈定的是一块而不是光标
     *   header          只要表头行
     *   row / col       第 N 行或某一列，列宽跟该表数据区走
     *   table           每张目标表自己的数据区（至少两行）。光标格、自动捕获、整表误选都不算
     *   skip            用户说了不要改格式
     */
    function decideFormatTarget(message, context) {
        const m = String(message || '');
        const ctx = context || {};
        const sel = classifySelection(ctx.selection);
        const active = (ctx.activeSheet && (ctx.activeSheet.sheet || ctx.activeSheet.Name))
            || (typeof ctx.activeSheet === 'string' ? ctx.activeSheet : '')
            || sel.sheet || '';
        const named = _sheetNamesIn(m, ctx);
        const base = {
            sheetScope: 'active', namedSheets: named, rangeMode: 'table', spokenRange: '',
            selectionRange: sel.address || '', selectionKind: sel.kind, selectionSheet: sel.sheet || '',
            activeSheet: active, reason: ''
        };
        if (!m) return Object.assign(base, { reason: 'empty-message' });
        if (/(不要|无需|不用|不必|禁止|请勿).{0,8}(美化|格式化)|保持原样|格式不变|维持原样/.test(m)) {
            return Object.assign(base, { sheetScope: 'none', rangeMode: 'skip', reason: 'negated' });
        }

        const batch = /(所有|全部|各个|每个|每一张|每张|批量)/.test(m) && /(工作表|sheet|表格|数据表)/i.test(m);
        const book = /(整个|整本|全部).{0,4}工作簿/.test(m);
        const others = /(其他|其余|剩余|剩下)/.test(m) && /(工作表|sheet|表格)/i.test(m);
        let sheetScope = 'active';
        if (others && !batch && !book) sheetScope = 'others';
        else if (batch || book) sheetScope = 'all';
        else if (named.length) sheetScope = 'named';

        const spoken = extractSpokenRange(m);
        const wantsCell = /这个单元格|这个格子|当前单元格|活动单元格|这一格/.test(m);
        const wantsSelection = /选中|所选|选区|这块|这段|这一行|这行|当前选|我选/.test(m);
        const wantsSameArea = /同样的区域|同一个区域|照这个范围|按这个范围|按这个区域|按选区|按这个选区/.test(m);
        const wantsWhole = /整表|整个表|全表|整个工作表|整张工作表|所有数据|全部数据|这个表|这张表|当前表/.test(m);
        const wantsHeader = /表头|标题行|首行|第一行/.test(m) && !/整表|整个表|全表|所有数据|全部数据/.test(m);
        const rowN = m.match(/第\s*(\d+)\s*行/);
        const colLetter = m.match(/([A-Za-z]{1,3})\s*列/);

        let rangeMode = 'table';
        let spokenRange = '';
        let reason = 'per-sheet-data';

        if (spoken) {
            rangeMode = 'spoken';
            spokenRange = spoken;
            reason = 'spoken-address';
        } else if (wantsCell) {
            rangeMode = 'cell';
            reason = 'named-cell';
        } else if (wantsSameArea && (sheetScope === 'all' || sheetScope === 'others')) {
            if (_realRegion(sel.kind)) {
                rangeMode = 'copy-selection';
                reason = 'same-area-on-each-sheet';
            } else {
                reason = 'same-area-but-no-real-selection';
            }
        } else if (wantsSelection && _realRegion(sel.kind)) {
            const foreign = sheetScope === 'named' && sel.sheet && named.indexOf(sel.sheet) < 0;
            if (sheetScope === 'all' || sheetScope === 'others' || foreign) {
                reason = foreign ? 'selection-sheet-mismatch' : 'all-sheets-keep-own-data';
            } else {
                rangeMode = 'selection';
                reason = sel.kind === 'row' ? 'selected-row' : 'selected-region';
            }
        } else if (wantsHeader) {
            rangeMode = 'header';
            reason = 'header-row-of-table';
        } else if (rowN) {
            rangeMode = 'row';
            spokenRange = rowN[1];
            reason = 'spoken-row';
        } else if (colLetter) {
            rangeMode = 'col';
            spokenRange = colLetter[1].toUpperCase();
            reason = 'spoken-column';
        } else if (sheetScope === 'all' || sheetScope === 'others') {
            reason = sel.kind === 'cell' ? 'cursor-ignored-for-all-sheets'
                : (sel.kind === 'auto' ? 'auto-capture-ignored' : 'each-sheet-own-data');
        } else if (_realRegion(sel.kind) && sel.sheet && sheetScope === 'named' && named.indexOf(sel.sheet) < 0) {
            reason = 'selection-sheet-mismatch';
        } else if (_realRegion(sel.kind)) {
            if (sel.kind === 'row' && wantsWhole) {
                reason = 'row-selected-but-asked-whole-table';
            } else {
                rangeMode = 'selection';
                reason = sel.kind === 'row' ? 'selected-row' : 'selected-region';
            }
        } else if (sel.kind === 'cell') {
            reason = 'cursor-cell-use-table';
        } else if (sel.kind === 'auto') {
            reason = 'auto-capture-not-a-selection';
        } else if (sel.kind === 'whole') {
            reason = 'whole-sheet-selection-use-table';
        } else if (sel.kind === 'object') {
            reason = 'object-selected-not-cells';
        } else {
            reason = 'no-selection-use-table';
        }

        return Object.assign(base, {
            sheetScope: sheetScope, rangeMode: rangeMode, spokenRange: spokenRange, reason: reason
        });
    }

    /** 数据区的第一行，供「只改表头」。A1:I21 → A1:I1。 */
    function headerRowOf(addr) {
        const s = normRangeAddr(addr);
        const m = s.match(/^([A-Za-z]{1,3})(\d+):([A-Za-z]{1,3})\d+$/);
        return m ? (m[1] + m[2] + ':' + m[3] + m[2]) : s;
    }

    App.ChatV2.rangeResolver = {
        MIN_TABLE_ROWS: 2,
        colToNum: colToNum,
        normRangeAddr: normRangeAddr,
        isSingleCellAddr: isSingleCellAddr,
        isReliableTableRange: isReliableTableRange,
        usedRangeAddr: usedRangeAddr,
        readRangeAddr: readRangeAddr,
        createSheetRangeLadder: createSheetRangeLadder,
        classifySelection: classifySelection,
        extractSpokenRange: extractSpokenRange,
        decideFormatTarget: decideFormatTarget,
        headerRowOf: headerRowOf
    };
})();
