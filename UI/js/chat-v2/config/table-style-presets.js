/**
 * TableStylePresets — 表格「样式方案」预设（Excel / Word / PPT 三宿主共享的单一事实源）
 *
 * 背景（用户反馈）：以前的"主题"只有颜色不同（商务蓝/科技绿/…），版式一模一样，
 * 换主题等于换个配色，看不出"样式版本"的差异。
 *
 * 解决方案：把样式拆成【三个正交维度】+【命名预设】，任一维度都能单独组合：
 *   ① 配色 palette ：商务蓝 / 科技绿 / 活力橙 / 优雅紫 / 极简灰 / 中国红 / 金融红 / 墨黑金 / 晴空青
 *   ② 版式 layout  ：经典网格 / 深表头 / 三线表(学术) / 无边框留白 / 卡片条纹 /
 *                    下划线表头 / 首列强调 / 合计行 / 紧凑报表 / 杂志宽松
 *   ③ 密度 density ：compact 紧凑 / normal 标准 / loose 宽松（行高 + 留白）
 *   另加开关：隔行底纹 zebra、首列加粗、末行(合计行)强调、数字右对齐、垂直居中。
 *
 * 使用方：
 *   - Word/PPT：mcp-dispatcher.js 的 buildHostTableFormatArgs → toHostArgs()
 *   - Excel   ：skill-spec-engine.js 的 specFromMessage / normalizeSpec → toExcelSpec()
 *
 * ⚠️ 本文件必须是自包含的（不依赖任何其它模块），因为 skill-spec-engine 会被
 *    Tests 以 eval 方式单独加载，缺模块时只会导致 preset 不生效、不能报错。
 */

(function (global) {
    'use strict';
    var App = global.App || (global.App = {});
    var ChatV2 = App.ChatV2 || (App.ChatV2 = {});

    // ============================================================
    // ① 配色（palette）
    //   headerFill 浅色表头底；headerDeep 深色表头底（深底白字用）；
    //   headerFontOnDeep 深底上的字色；border 网格线色；zebra 隔行底纹色
    // ============================================================
    var PALETTES = {
        business_blue: {
            id: 'business_blue', name: '商务蓝',
            aliases: ['商务蓝', '深蓝', '蓝色', '蓝'],
            headerFill: '#D9E2F3', headerDeep: '#1F4E79', headerFontOnDeep: '#FFFFFF',
            border: '#4472C4', zebra: '#EDF2FA', headerFontOnLight: '#1F3864'
        },
        tech_green: {
            id: 'tech_green', name: '科技绿',
            aliases: ['科技绿', '绿色', '绿'],
            headerFill: '#D9F2EA', headerDeep: '#1E6F5C', headerFontOnDeep: '#FFFFFF',
            border: '#2A9D8F', zebra: '#EAF7F3', headerFontOnLight: '#14594A'
        },
        vibrant_orange: {
            id: 'vibrant_orange', name: '活力橙',
            aliases: ['活力橙', '橙色', '橙'],
            headerFill: '#FCE4D6', headerDeep: '#C1502E', headerFontOnDeep: '#FFFFFF',
            border: '#E76F51', zebra: '#FDF1EA', headerFontOnLight: '#8C3A22'
        },
        elegant_purple: {
            id: 'elegant_purple', name: '优雅紫',
            aliases: ['优雅紫', '紫色', '紫'],
            headerFill: '#E5D9F2', headerDeep: '#5B3A8E', headerFontOnDeep: '#FFFFFF',
            border: '#7030A0', zebra: '#F3EDF9', headerFontOnLight: '#4A2F75'
        },
        minimalist_gray: {
            id: 'minimalist_gray', name: '极简灰',
            aliases: ['极简灰', '灰色', '灰'],
            headerFill: '#EDEDED', headerDeep: '#595959', headerFontOnDeep: '#FFFFFF',
            border: '#BFBFBF', zebra: '#F5F5F5', headerFontOnLight: '#333333'
        },
        finance_red: {
            id: 'finance_red', name: '金融红',
            aliases: ['金融红', '红色', '红'],
            headerFill: '#F8D7DA', headerDeep: '#9C1C24', headerFontOnDeep: '#FFFFFF',
            border: '#C00000', zebra: '#FBEDEF', headerFontOnLight: '#7A1620'
        },
        ink_gold: {
            id: 'ink_gold', name: '墨黑金',
            aliases: ['墨黑金', '黑金', '金色', '金'],
            headerFill: '#EFE7D6', headerDeep: '#2B2B2B', headerFontOnDeep: '#E8C874',
            border: '#B99A4B', zebra: '#F7F3EA', headerFontOnLight: '#3B3222'
        },
        sky_cyan: {
            id: 'sky_cyan', name: '晴空青',
            aliases: ['晴空青', '青色', '青'],
            headerFill: '#D8EEF8', headerDeep: '#0F6C8C', headerFontOnDeep: '#FFFFFF',
            border: '#2E9BC4', zebra: '#EBF6FB', headerFontOnLight: '#0B5268'
        }
    };
    var DEFAULT_PALETTE = 'business_blue';

    // ============================================================
    // ② 版式（layout）：真正的"排版"差异所在
    //   borderStyle : all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
    //   headerMode  : deep 深底白字 | light 浅底深字 | underline 仅底部粗线(无底) | italic 斜体
    //   density     : compact / normal / loose（行高与留白）
    // ============================================================
    var LAYOUTS = {
        classic_grid: {
            id: 'classic_grid', name: '经典网格',
            aliases: ['经典网格', '全网格', '全边框', '经典', '网格', '标准表格'],
            desc: '全网格边框 + 浅色表头 + 隔行底纹 + 居中',
            borderStyle: 'all', headerMode: 'light', density: 'normal',
            zebra: true, headerAlign: 'center', bodyAlign: 'center',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        deep_header: {
            id: 'deep_header', name: '深表头',
            aliases: ['深表头', '深色表头', '深底白字', '深底', '深色底'],
            desc: '深色表头 + 白字，视觉重心在上，适合汇报',
            borderStyle: 'all', headerMode: 'deep', density: 'normal',
            zebra: true, headerAlign: 'center', bodyAlign: 'center',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        three_line: {
            id: 'three_line', name: '三线表',
            aliases: ['三线表', '三线', '学术表', '论文表', '期刊表', '学术风'],
            desc: '论文/报告规范：仅顶线、表头下线与底线，无竖线无底色',
            borderStyle: 'threeLine', headerMode: 'underline', density: 'normal',
            zebra: false, headerAlign: 'center', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        borderless: {
            id: 'borderless', name: '无边框留白',
            aliases: ['无边框留白', '极简留白', '无边框', '无框', '留白'],
            desc: '去网格、靠留白与行距分区，最克制',
            borderStyle: 'none', headerMode: 'underline', density: 'loose',
            zebra: false, headerAlign: 'left', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        card_stripe: {
            id: 'card_stripe', name: '卡片条纹',
            aliases: ['卡片条纹', '条纹卡片', '卡片风', '卡片'],
            desc: '整行浅底交替如卡片，无竖线，行距宽松',
            borderStyle: 'none', headerMode: 'light', density: 'loose',
            zebra: true, headerAlign: 'left', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        underline_header: {
            id: 'underline_header', name: '下划线表头',
            aliases: ['下划线表头', '表头下划线', '极简下划线', '简约下划线'],
            desc: '表头无底色、只留一条下边线，仅保留横向分隔',
            borderStyle: 'horizontal', headerMode: 'underline', density: 'normal',
            zebra: false, headerAlign: 'left', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        first_col_focus: {
            id: 'first_col_focus', name: '首列强调',
            aliases: ['首列强调', '强调首列', '首列加粗', '左侧标题列'],
            desc: '左侧第一列当行标题：加粗 + 浅底，正文左对齐',
            borderStyle: 'all', headerMode: 'light', density: 'normal',
            zebra: false, headerAlign: 'center', bodyAlign: 'left',
            firstColBold: true, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        },
        total_row: {
            id: 'total_row', name: '合计行',
            aliases: ['合计行', '总计行', '汇总行', '末行强调', '末行加粗'],
            desc: '末行做合计：加粗 + 上边线 + 浅底，数字右对齐',
            borderStyle: 'all', headerMode: 'light', density: 'normal',
            zebra: true, headerAlign: 'center', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: true, bodyNumberRight: true, valign: 'middle'
        },
        compact_report: {
            id: 'compact_report', name: '紧凑报表',
            aliases: ['紧凑报表', '紧凑', '紧密', '压缩', '密集'],
            desc: '行高压到最紧、数字右对齐，一屏看更多行',
            borderStyle: 'all', headerMode: 'light', density: 'compact',
            zebra: true, headerAlign: 'center', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: true, valign: 'middle'
        },
        magazine: {
            id: 'magazine', name: '杂志宽松',
            aliases: ['杂志宽松', '杂志风', '宽松', '大留白', '宽松风'],
            desc: '深表头 + 宽松行距 + 横向分隔，偏画册/杂志',
            borderStyle: 'horizontal', headerMode: 'deep', density: 'loose',
            zebra: true, headerAlign: 'left', bodyAlign: 'left',
            firstColBold: false, lastRowEmphasis: false, bodyNumberRight: false, valign: 'middle'
        }
    };
    var DEFAULT_LAYOUT = 'classic_grid';

    // ============================================================
    // ③ 密度（行高/留白），按宿主分别给磅值（Word/PPT 与 Excel 的可读行高不同）
    // ============================================================
    var ROW_HEIGHT = {
        host: { compact: 16, normal: 22, loose: 30 },
        excel: { compact: 15, normal: 20, loose: 28 }
    };

    function paletteOf(id) { return PALETTES[id] || PALETTES[DEFAULT_PALETTE]; }
    function layoutOf(id) { return LAYOUTS[id] || LAYOUTS[DEFAULT_LAYOUT]; }

    /** 全部版式清单（供 UI / 提示词枚举） */
    function listLayouts() {
        return Object.keys(LAYOUTS).map(function (k) {
            var l = LAYOUTS[k];
            return { id: l.id, name: l.name, desc: l.desc, aliases: l.aliases.slice() };
        });
    }

    /** 全部配色清单（供 UI / 提示词枚举） */
    function listPalettes() {
        return Object.keys(PALETTES).map(function (k) {
            var p = PALETTES[k];
            return { id: p.id, name: p.name, headerFill: p.headerFill, headerDeep: p.headerDeep, border: p.border };
        });
    }

    // 别名索引（长词优先，避免"留白"抢掉"无边框留白"）
    var _aliasIndex = (function () {
        var items = [];
        Object.keys(LAYOUTS).forEach(function (k) {
            (LAYOUTS[k].aliases || []).forEach(function (a) { items.push({ kind: 'layout', id: k, alias: a }); });
        });
        Object.keys(PALETTES).forEach(function (k) {
            (PALETTES[k].aliases || []).forEach(function (a) { items.push({ kind: 'palette', id: k, alias: a }); });
        });
        items.sort(function (x, y) { return y.alias.length - x.alias.length; });
        return items;
    })();

    /**
     * 从自然语言里识别【版式 + 配色 + 密度 + 开关】。
     * 未识别到的维度返回 null（调用方决定是否套默认值）。
     * @param {string} message
     * @returns {{layoutId:string|null, paletteId:string|null, density:string|null,
     *            zebra:boolean|null, firstColBold:boolean|null, lastRowEmphasis:boolean|null,
     *            bodyNumberRight:boolean|null, headerMode:string|null, valign:string|null}}
     */
    function detect(message) {
        var m = String(message || '');
        var out = {
            layoutId: null, paletteId: null, density: null, zebra: null,
            firstColBold: null, lastRowEmphasis: null, bodyNumberRight: null,
            headerMode: null, valign: null
        };
        if (!m) return out;
        for (var i = 0; i < _aliasIndex.length; i++) {
            var it = _aliasIndex[i];
            if (m.indexOf(it.alias) >= 0) {
                if (it.kind === 'layout' && !out.layoutId) out.layoutId = it.id;
                else if (it.kind === 'palette' && !out.paletteId) out.paletteId = it.id;
            }
        }
        // 只说了"美化/主题/样式/格式"没说具体版式 → 用经典网格兜底（保证有完整排版，而不只是换颜色）
        if (!out.layoutId && /美化|主题|样式|表格格式|设置格式|格式化|风格/.test(m)) out.layoutId = DEFAULT_LAYOUT;
        // 密度
        if (/紧凑|紧密|压缩|密集|小行高/.test(m)) out.density = 'compact';
        else if (/宽松|大行高|稀疏|大留白/.test(m)) out.density = 'loose';
        // 表头形态（显式措辞优先于版式自带）
        if (/深表头|深色表头|深底白字|深底|深色底|白字表头/.test(m)) out.headerMode = 'deep';
        else if (/浅色表头|浅底表头|浅色底/.test(m)) out.headerMode = 'light';
        else if (/下划线表头|表头下划线|表头下边线/.test(m)) out.headerMode = 'underline';
        else if (/斜体表头|表头斜体/.test(m)) out.headerMode = 'italic';
        // 开关
        if (/隔行|斑马|条纹|斑马纹/.test(m)) out.zebra = true;
        if (/不要斑马|无斑马|取消隔行|取消斑马|不要隔行|不要条纹/.test(m)) out.zebra = false;
        if (/首列加粗|首列强调|第一列加粗|左侧列加粗|强调首列/.test(m)) out.firstColBold = true;
        if (/合计行|总计行|汇总行|末行加粗|末行强调|最后一行加粗|合计行强调/.test(m)) out.lastRowEmphasis = true;
        if (/数字右对齐|数值右对齐|数字靠右|金额右对齐|数字右/.test(m)) out.bodyNumberRight = true;
        if (/垂直居中|上下居中|垂直方向居中/.test(m)) out.valign = 'middle';
        else if (/顶部对齐|靠上对齐|顶端对齐/.test(m)) out.valign = 'top';
        return out;
    }

    /**
     * 直接按 id 组装样式描述符（不经过自然语言识别）。
     * 供 skill-spec-engine / dispatcher 在已拿到结构化 id 时展开预设。
     * @param {string} layoutId
     * @param {string} [paletteId]
     */
    function describe(layoutId, paletteId) {
        var layout = LAYOUTS[layoutId] || layoutOf(DEFAULT_LAYOUT);
        var palette = PALETTES[paletteId] || paletteOf(DEFAULT_PALETTE);
        return {
            layoutId: layout.id,
            layoutName: layout.name,
            paletteId: palette.id,
            paletteName: palette.name,
            palette: palette,
            borderStyle: layout.borderStyle,
            headerMode: layout.headerMode,
            density: layout.density,
            zebra: !!layout.zebra,
            firstColBold: !!layout.firstColBold,
            lastRowEmphasis: !!layout.lastRowEmphasis,
            bodyNumberRight: !!layout.bodyNumberRight,
            valign: layout.valign || 'middle',
            headerAlign: layout.headerAlign || 'center',
            bodyAlign: layout.bodyAlign || 'center',
            overrides: null
        };
    }

    /**
     * 把识别结果补全成完整样式描述符（未指定的维度取版式/默认配色）。
     * @param {string} message
     * @param {Object} [base] 额外的显式覆盖（如用户直接给了颜色/字号）
     * @returns {Object} descriptor
     */
    function resolve(message, base) {
        var d = detect(message);
        var layout = layoutOf(d.layoutId || DEFAULT_LAYOUT);
        var palette = paletteOf(d.paletteId || DEFAULT_PALETTE);
        var density = d.density || layout.density || 'normal';
        var headerMode = d.headerMode || layout.headerMode || 'light';
        return {
            layoutId: layout.id,
            layoutName: layout.name,
            paletteId: palette.id,
            paletteName: palette.name,
            palette: palette,
            borderStyle: layout.borderStyle,
            headerMode: headerMode,
            density: density,
            zebra: d.zebra === null ? !!layout.zebra : d.zebra,
            firstColBold: d.firstColBold === null ? !!layout.firstColBold : d.firstColBold,
            lastRowEmphasis: d.lastRowEmphasis === null ? !!layout.lastRowEmphasis : d.lastRowEmphasis,
            bodyNumberRight: d.bodyNumberRight === null ? !!layout.bodyNumberRight : d.bodyNumberRight,
            valign: d.valign || layout.valign || 'middle',
            headerAlign: layout.headerAlign || 'center',
            bodyAlign: layout.bodyAlign || 'center',
            overrides: base || null
        };
    }

    /**
     * 样式描述符 → Word/PPT 的 {word|ppt}_format_selected_table_cell 参数。
     * @param {Object} desc resolve() 的产物
     * @param {Object} [extra] 调用方已有的显式参数（优先级最高）
     */
    function toHostArgs(desc, extra) {
        var p = desc.palette || paletteOf(DEFAULT_PALETTE);
        var deep = desc.headerMode === 'deep';
        var underline = desc.headerMode === 'underline';
        var a = {
            borderStyle: desc.borderStyle,
            align: desc.headerAlign,          // 表头对齐（正文另由 bodyAlign 覆盖）
            bodyAlign: desc.bodyAlign,
            borderColor: p.border,
            zebra: !!desc.zebra,
            zebraColor: p.zebra,
            firstColBold: !!desc.firstColBold,
            lastRowBold: !!desc.lastRowEmphasis,
            valign: desc.valign,
            rowHeight: (ROW_HEIGHT.host[desc.density] || ROW_HEIGHT.host.normal)
        };
        if (deep) {
            a.headerFillColor = p.headerDeep;
            a.headerFontColor = p.headerFontOnDeep;
            a.headerBold = true;
        } else if (underline) {
            a.headerFillColor = '';           // 无底色：仅底部粗线
            a.headerFontColor = p.headerFontOnLight;
            a.headerBold = true;
        } else if (desc.headerMode === 'italic') {
            a.headerFillColor = '';
            a.headerFontColor = p.headerFontOnLight;
            a.headerBold = false;
            a.headerItalic = true;
        } else {
            a.headerFillColor = p.headerFill;
            a.headerFontColor = p.headerFontOnLight;
            a.headerBold = true;
        }
        if (extra) {
            Object.keys(extra).forEach(function (k) {
                if (extra[k] !== undefined && extra[k] !== null && extra[k] !== '') a[k] = extra[k];
            });
        }
        return a;
    }

    /**
     * 样式描述符 → Excel 表格美化（SmartTableBeautifier）的 spec 字段。
     * 字段名与 skill-spec-engine.js 的 DEFAULT_SPEC 对齐。
     */
    function toExcelSpec(desc, extra) {
        var p = desc.palette || paletteOf(DEFAULT_PALETTE);
        var s = {
            stylePreset: desc.layoutId,
            theme: desc.paletteId,
            borderStyle: desc.borderStyle,
            headerLayout: desc.headerMode,
            rowHeight: (ROW_HEIGHT.excel[desc.density] || ROW_HEIGHT.excel.normal),
            valign: desc.valign,
            bandedRows: !!desc.zebra,
            bandColor: null,
            boldFirstColumn: !!desc.firstColBold,
            emphasizeLastRow: !!desc.lastRowEmphasis,
            bodyNumberRight: !!desc.bodyNumberRight,
            alignment: desc.bodyAlign,
            addBorders: desc.borderStyle !== 'none'
        };
        if (extra) {
            Object.keys(extra).forEach(function (k) {
                if (extra[k] !== undefined && extra[k] !== null && extra[k] !== '') s[k] = extra[k];
            });
        }
        return s;
    }

    /** 把 #RRGGBB 转成 [R,G,B]（Excel 侧渲染 VBA 用） */
    function hexToRgb(hex) {
        var h = String(hex || '').replace('#', '');
        if (h.length !== 6) return [0, 0, 0];
        return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)];
    }

    ChatV2.tableStylePresets = {
        PALETTES: PALETTES,
        LAYOUTS: LAYOUTS,
        DEFAULT_PALETTE: DEFAULT_PALETTE,
        DEFAULT_LAYOUT: DEFAULT_LAYOUT,
        ROW_HEIGHT: ROW_HEIGHT,
        listLayouts: listLayouts,
        listPalettes: listPalettes,
        detect: detect,
        describe: describe,
        resolve: resolve,
        toHostArgs: toHostArgs,
        toExcelSpec: toExcelSpec,
        hexToRgb: hexToRgb
    };
    App.tableStylePresets = ChatV2.tableStylePresets;
})(typeof window !== 'undefined' ? window : this);
