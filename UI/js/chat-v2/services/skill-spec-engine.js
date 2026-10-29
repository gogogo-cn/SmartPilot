/**
 * ============================================================
 * chat-v2/services/skill-spec-engine.js — Skill Spec 规划引擎
 * ============================================================
 * 架构原则（解决"固定模板 + LLM 只识别不规划"的根本问题）：
 *   1. LLM 不再输出固定操作 —— 它输出【结构化 Spec JSON】描述"想要什么效果"
 *      （主题、表头样式、对齐、列宽、条件格式、冻结窗格等）
 *   2. 真正的 VBA/操作由人工验证过的模板 + 参数化渲染生成
 *   3. 确定性：相同 Spec 永远产出相同代码 → 永不因 LLM 细节错误报错
 *   4. 灵活性：LLM 通过 Spec 字段控制样式变化，模板负责"代码怎么组织"
 *
 * 示例 Spec（美化表格）：
 *   {
 *     "theme": "business_blue",              // 主题（引擎内置 6 种）
 *     "headerStyle": "bold_colored",          // 表头样式
 *     "alignment": "center",                  // 对齐
 *     "autofit": true,                        // 自动列宽
 *     "freezeHeader": true,                   // 冻结表头
 *     "addBorders": true,                     // 边框
 *     "bandedRows": true,                     // 斑马纹
 *     "conditionalFormat": "dataBar",         // 条件格式（可选）
 *     "numberFormat": "#,##0",                // 数字格式（可选）
 *     "title": "2023年度销售报表"              // 首行标题（可选）
 *   }
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.services) App.ChatV2.services = {};

    const ChatLog = window.ChatLog || console;

    // 图表映射唯一事实源：App.ChatV2.intentConfig（Config/intent_config.json chartSpec 段）
    function defs() {
        const ic = App.ChatV2.intentConfig;
        return (ic && typeof ic.defs === 'function') ? ic.defs() : {};
    }

    // ============================================================
    // 动态颜色支持（解决"换颜色词没变化"）：
    // LLM 可以输出任意颜色词（红/黄/蓝/粉/金/自定义...）或 RGB 数组，
    // 引擎映射为真实颜色。主题(theme)是预设组合，自定义颜色字段优先。
    // ============================================================

    /** 常见颜色词 → RGB（中文 + 英文） */
    const COLOR_WORDS = {
        // 中文
        '中国红': [222, 41, 16], '朱红': [222, 41, 16], '国旗红': [222, 41, 16],
        '红': [220, 60, 60], '红色': [220, 60, 60], '大红': [200, 30, 30], '深红': [160, 40, 40],
        '橙': [235, 130, 40], '橙色': [235, 130, 40], '橘': [235, 130, 40], '橘色': [235, 130, 40],
        '黄': [240, 190, 50], '黄色': [240, 190, 50], '金黄': [212, 175, 55], '金色': [212, 175, 55],
        '绿': [60, 150, 90], '绿色': [60, 150, 90], '深绿': [40, 110, 70], '浅绿': [120, 200, 140],
        '青': [30, 140, 140], '青色': [30, 140, 140], '蓝绿': [30, 140, 140],
        '蓝': [50, 110, 190], '蓝色': [50, 110, 190], '深蓝': [30, 70, 140], '浅蓝': [130, 180, 230],
        '天蓝': [80, 160, 220], '天蓝': [80, 160, 220],
        '紫': [130, 90, 180], '紫色': [130, 90, 180], '深紫': [100, 60, 150], '浅紫': [180, 150, 220],
        '粉': [230, 130, 170], '粉色': [230, 130, 170], '桃红': [235, 100, 140],
        '棕': [140, 95, 55], '棕色': [140, 95, 55], '褐色': [140, 95, 55],
        '黑': [50, 50, 50], '黑色': [50, 50, 50], '灰': [128, 128, 128], '灰色': [128, 128, 128],
        '银': [190, 195, 200], '银色': [190, 195, 200], '白': [245, 245, 245], '白色': [245, 245, 245],
        '薄荷': [100, 190, 160], '薄荷绿': [100, 190, 160], '翡翠': [50, 160, 120],
        '酒红': [140, 30, 50], '玫红': [220, 80, 120], '珊瑚': [240, 128, 110],
        // 英文
        'red': [220, 60, 60], 'orange': [235, 130, 40], 'yellow': [240, 190, 50],
        'green': [60, 150, 90], 'teal': [30, 140, 140], 'cyan': [30, 150, 170],
        'blue': [50, 110, 190], 'purple': [130, 90, 180], 'pink': [230, 130, 170],
        'brown': [140, 95, 55], 'black': [50, 50, 50], 'gray': [128, 128, 128], 'grey': [128, 128, 128],
        'silver': [190, 195, 200], 'white': [245, 245, 245], 'gold': [212, 175, 55],
        'navy': [30, 50, 110], 'maroon': [140, 30, 50], 'olive': [130, 140, 60],
        'violet': [140, 110, 200], 'magenta': [210, 60, 150], 'coral': [240, 128, 110],
        'mint': [100, 190, 160], 'sky': [80, 160, 220], 'indigo': [75, 60, 150],
        // 补充常见英文色（LLM 常用）
        'lavender': [140, 110, 200], 'emerald': [16, 140, 80], 'rust': [160, 80, 40],
        'cerulean': [60, 140, 190], 'mauve': [170, 110, 150], 'ocean': [30, 90, 160],
        'salmon': [250, 128, 114], 'turquoise': [64, 224, 208], 'amber': [255, 191, 0],
        'crimson': [220, 20, 60], 'beige': [245, 245, 220], 'ivory': [255, 255, 240],
        'khaki': [195, 176, 145], 'tan': [210, 180, 140], 'chocolate': [123, 63, 0],
        'forest': [34, 139, 34], 'lime': [50, 205, 50], 'aqua': [0, 255, 255],
        'fuchsia': [255, 0, 255], 'plum': [221, 160, 221], 'orchid': [218, 112, 214],
        'steel': [70, 130, 180], 'slate': [112, 128, 144], 'periwinkle': [204, 204, 255],
        'scarlet': [255, 36, 0], 'burgundy': [128, 0, 32], 'sage': [155, 179, 143],
        'charcoal': [54, 69, 79], 'denim': [21, 96, 189], 'cobalt': [0, 71, 171]
    };

    /** 颜色词 → 主题关键词别名（中英混合规范化） */
    function toColorKey(text) {
        return String(text || '').toLowerCase().trim();
    }

    /**
     * 把颜色词/RGB/主题名解析为 RGB 数组。
     * @param {*} color - 颜色词字符串 / RGB 数组 / 主题 id
     * @returns {Array<number>|null} [r,g,b] 或 null
     */
    function resolveColor(color) {
        if (color == null) return null;
        // RGB 数组 [r,g,b]
        if (Array.isArray(color) && color.length >= 3) {
            const r = Number(color[0]), g = Number(color[1]), b = Number(color[2]);
            if ([r, g, b].every(v => Number.isFinite(v) && v >= 0 && v <= 255)) return [Math.round(r), Math.round(g), Math.round(b)];
            return null;
        }
        const key = toColorKey(color);
        // 直接颜色词
        if (COLOR_WORDS[key]) return COLOR_WORDS[key];
        // 主题 id → 主题 header 色
        if (THEMES[key]) return THEMES[key].header;
        // 包含颜色词的短语（如"黄色主题"、"深蓝色"）
        for (const [word, rgb] of Object.entries(COLOR_WORDS)) {
            if (key.includes(word)) return rgb;
        }
        // 数字形式 "255,200,0"
        const nums = key.match(/\d+/g);
        if (nums && nums.length >= 3) {
            const r = Number(nums[0]), g = Number(nums[1]), b = Number(nums[2]);
            if ([r, g, b].every(v => Number.isFinite(v) && v <= 255)) return [r, g, b];
        }
        // 任意未知颜色词（lavender/sunset/薄荷绿/任意描述）→
        // 用词哈希生成稳定色相（HSV→RGB），保证任何词都有对应颜色，绝不回退默认。
        // 这是"去硬编码"的最后一环：颜色完全由 LLM/用户输入驱动，引擎不限制词汇。
        if (key.length > 0) {
            let hash = 0;
            for (let i = 0; i < key.length; i++) {
                hash = (hash * 31 + key.charCodeAt(i)) >>> 0;
            }
            // 色相 0-360（哈希决定），饱和度/亮度取较高值保证颜色明显
            const hue = hash % 360;
            const sat = 0.55 + ((hash >>> 8) % 30) / 100;   // 0.55-0.84
            const light = 0.35 + ((hash >>> 16) % 20) / 100; // 0.35-0.54（深色系表头，白字可读）
            return hslToRgb(hue, sat, light);
        }
        return null;
    }

    /** HSL → RGB（用于未知颜色词兜底） */
    function hslToRgb(h, s, l) {
        const c = (1 - Math.abs(2 * l - 1)) * s;
        const hp = h / 60;
        const x = c * (1 - Math.abs(hp % 2 - 1));
        let r = 0, g = 0, b = 0;
        if (hp < 1) { r = c; g = x; }
        else if (hp < 2) { r = x; g = c; }
        else if (hp < 3) { g = c; b = x; }
        else if (hp < 4) { g = x; b = c; }
        else if (hp < 5) { r = x; b = c; }
        else { r = c; b = x; }
        const m = l - c / 2;
        return [
            Math.round((r + m) * 255),
            Math.round((g + m) * 255),
            Math.round((b + m) * 255)
        ];
    }

    /** 从字符串提取颜色词（如"红色"、"yellow"、"深蓝色"） */
    function extractColorWords(text) {
        if (!text) return [];
        const t = toColorKey(text);
        const found = [];
        for (const [word, rgb] of Object.entries(COLOR_WORDS)) {
            if (word.length >= 1 && t.includes(word)) found.push({ word, rgb });
        }
        return found;
    }

    // ============================================================
    // 主题定义（RGB 数组，与 vba-template-engine 的 COLOR_THEMES 一致风格）
    // ============================================================
    const THEMES = {
        'business_blue':    { name: '商务蓝',   header: [0, 90, 156],  accent: [255, 165, 0],  band: [226, 238, 250], border: [180, 200, 220], body: [242, 248, 255] },
        'tech_green':       { name: '科技绿',   header: [0, 128, 128], accent: [255, 105, 97], band: [216, 240, 240], border: [170, 210, 210], body: [238, 248, 248] },
        'vibrant_orange':   { name: '活力橙',   header: [230, 100, 50], accent: [50, 100, 180], band: [255, 234, 220], border: [230, 190, 160], body: [255, 246, 240] },
        'elegant_purple':   { name: '优雅紫',   header: [100, 60, 150], accent: [220, 180, 50], band: [235, 225, 245], border: [200, 185, 220], body: [247, 242, 252] },
        'minimalist_gray':  { name: '极简灰',   header: [80, 80, 80],  accent: [0, 120, 200],  band: [238, 238, 238], border: [200, 200, 200], body: [248, 248, 248] },
        'finance_red':      { name: '金融红',   header: [160, 40, 40], accent: [0, 110, 180],  band: [250, 235, 235], border: [220, 180, 180], body: [252, 243, 243] }
    };

    /** 默认 Spec（确定性兜底：LLM 未提供任何字段时使用） */
    const DEFAULT_SPEC = {
        theme: 'business_blue',
        headerColor: null,      // 自定义表头颜色（优先于 theme）
        accentColor: null,      // 自定义强调色（数据条/高亮）
        bandColor: null,        // 自定义斑马纹色
        bodyColor: null,        // 自定义内容(数据)区底色，优先于 theme.body
        noContentBg: false,     // 用户明确"移除内容底色/不要底色" → 清空数据区填充，不加主题浅底、不加斑马纹色
        headerStyle: 'bold_colored',
        alignment: 'center',
        autofit: true,
        freezeHeader: false,
        addBorders: true,
        bandedRows: false,
        conditionalFormat: null,
        numberFormat: null,
        title: null,
        headerRows: 1,              // 多行表头总行数（含合并标题行），LLM 依据选中区数据自动判定
        // ── 排版维度（2026-10-07 新增：主题不再只是换颜色，版式/表头/行高/强调都能变）──
        // 一键样式方案（classic_grid/deep_header/three_line/borderless/card_stripe/
        // underline_header/first_col_focus/total_row/compact_report/magazine）。
        // 只给 stylePreset 时，normalizeSpec 会自动展开成下面这些具体字段。
        // 注意：borderStyle / headerLayout 默认必须是 null —— 否则默认值会盖住旧字段
        // （addBorders=false / headerStyle=bold_only）的兼容映射。实际取值在 renderBeautifyVBA 里派生。
        stylePreset: null,
        borderStyle: null,              // all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
        headerLayout: null,             // deep 深底白字 | light 浅底深字 | underline 仅底部粗线 | italic 斜体表头
        rowHeight: null,                // 行高磅值：紧凑 15 / 标准 20 / 宽松 28；null = 不改行高
        valign: 'middle',           // middle 垂直居中 | top 顶端对齐
        boldFirstColumn: false,     // 首列加粗（左侧标题列）+ 极浅底
        emphasizeLastRow: false,    // 末行（合计行）加粗 + 上边线 + 浅底
        bodyNumberRight: false,     // 正文数值单元格右对齐（文本仍按 alignment）
        // 格式整理字段
        fontName: '微软雅黑',
        fontSize: 11,
        dateFormat: null,
        wrapText: false,
        clearRedundant: false
    };

    // ============================================================
    // Spec 规范化：合并默认值 + 字段白名单（防止 LLM 注入非法字段）
    // ============================================================
    const ALLOWED_FIELDS = new Set(Object.keys(DEFAULT_SPEC));
    const VALID_THEMES = new Set(Object.keys(THEMES));
    const VALID_HEADER_STYLES = new Set(['bold_colored', 'bold_only', 'colored_only', 'italic']);
    const VALID_ALIGNMENTS = new Set(['left', 'center', 'right']);
    const VALID_CONDITIONAL = new Set(['dataBar', 'colorScale', 'highlight_top', 'highlight_duplicates']);
    // 排版维度白名单（防注入：只允许下列取值）
    const VALID_BORDER_STYLES = new Set(['all', 'none', 'horizontal', 'outer', 'threeLine']);
    const VALID_HEADER_LAYOUTS = new Set(['deep', 'light', 'underline', 'italic']);
    const VALID_VALIGN = new Set(['middle', 'top']);
    const VALID_STYLE_PRESETS = new Set([
        'classic_grid', 'deep_header', 'three_line', 'borderless', 'card_stripe',
        'underline_header', 'first_col_focus', 'total_row', 'compact_report', 'magazine'
    ]);

    /**
     * 规范化颜色值（颜色词 / RGB 数组 / 主题 id → RGB 数组或 null）
     */
    function normalizeColorValue(val) {
        if (val == null) return null;
        return resolveColor(val);
    }

    /**
     * 规范化 LLM 输出的 Spec（白名单校验 + 默认值合并）
     * @param {Object} raw - LLM 输出的 JSON
     * @returns {Object} 规范化后的 spec
     */
    /** 取共享样式预设模块（未加载时返回 null，功能降级但不报错） */
    function presetsModule() {
        try {
            const w = (typeof window !== 'undefined') ? window : null;
            return (w && w.App && w.App.ChatV2 && w.App.ChatV2.tableStylePresets) || null;
        } catch { return null; }
    }

    function normalizeSpec(raw) {
        const spec = { ...DEFAULT_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        // 【样式方案一键展开】只给了 stylePreset（如 "three_line"）时，
        // 先把预设展开成 borderStyle/headerLayout/rowHeight/... 具体字段，
        // 再由下方 switch 用显式字段覆盖（显式优先于预设）。
        const _P = presetsModule();
        const _presetId = raw.stylePreset || raw.preset || raw.layoutStyle;
        if (_P && VALID_STYLE_PRESETS.has(_presetId)) {
            try {
                const _desc = _P.describe(_presetId, VALID_THEMES.has(raw.theme) ? raw.theme : null);
                const _expanded = _P.toExcelSpec(_desc);
                Object.keys(_expanded).forEach(k => {
                    if (Object.prototype.hasOwnProperty.call(DEFAULT_SPEC, k)) spec[k] = _expanded[k];
                });
            } catch { /* 预设模块异常 → 忽略，走普通字段 */ }
        }
        for (const key of Object.keys(raw)) {
            if (!ALLOWED_FIELDS.has(key)) continue;  // 白名单：忽略未知字段
            const val = raw[key];
            switch (key) {
                case 'theme':
                    // theme 接受任意颜色词/RGB/主题id。
                    // 是合法主题 id → 保留；否则解析为自定义 headerColor（不再回退默认！）
                    if (VALID_THEMES.has(val)) {
                        spec.theme = val;
                    } else {
                        const rgb = resolveColor(val);
                        if (rgb) {
                            // 未知主题名但能解析出颜色 → 用该颜色作为表头色
                            spec.theme = 'custom';
                            spec.headerColor = rgb;
                        }
                        // 解析不出 → 保留默认 theme（白名单防护）
                    }
                    break;
                case 'headerColor':
                case 'accentColor':
                case 'bandColor':
                case 'bodyColor': {
                    const rgb = normalizeColorValue(val);
                    if (rgb) spec[key] = rgb;
                    break;
                }
                case 'headerStyle':
                    if (VALID_HEADER_STYLES.has(val)) spec.headerStyle = val;
                    break;
                case 'noContentBg':
                    spec.noContentBg = !!val;
                    break;
                case 'alignment':
                    if (VALID_ALIGNMENTS.has(val)) spec.alignment = val;
                    break;
                // ── 排版维度 ──
                case 'stylePreset':
                    if (val === null) spec.stylePreset = null;
                    else if (VALID_STYLE_PRESETS.has(val)) spec.stylePreset = val;
                    break;
                case 'borderStyle':
                    if (VALID_BORDER_STYLES.has(val)) spec.borderStyle = val;
                    break;
                case 'headerLayout':
                    if (VALID_HEADER_LAYOUTS.has(val)) spec.headerLayout = val;
                    break;
                case 'valign':
                    if (VALID_VALIGN.has(val)) spec.valign = val;
                    break;
                case 'rowHeight': {
                    const n = Number(val);
                    if (val === null) spec.rowHeight = null;
                    else if (Number.isFinite(n) && n >= 10 && n <= 60) spec.rowHeight = Math.round(n);
                    break;
                }
                case 'boldFirstColumn':
                case 'emphasizeLastRow':
                case 'bodyNumberRight':
                    spec[key] = val === true || val === 'true' || val === 1;
                    break;
                case 'conditionalFormat':
                    if (val === null || VALID_CONDITIONAL.has(val)) spec.conditionalFormat = val;
                    break;
                case 'numberFormat':
                    // 数字格式白名单（防注入）：只允许常见格式串
                    if (val === null) spec.numberFormat = null;
                    else if (typeof val === 'string' && /^[0#.,%$€¥\-\s()]*$/.test(val) && val.length <= 20) {
                        spec.numberFormat = val;
                    }
                    break;
                // 格式整理字段
                case 'fontName':
                    if (typeof val === 'string' && val.length <= 30 && /^[\u4e00-\u9fa5A-Za-z0-9\s]+$/.test(val)) spec.fontName = val;
                    break;
                case 'fontSize':
                    if (typeof val === 'number' && val >= 8 && val <= 24) spec.fontSize = Math.round(val);
                    break;
                case 'dateFormat':
                    if (val === null) spec.dateFormat = null;
                    else if (typeof val === 'string' && /^[ymd\-/.\s]*$/i.test(val) && val.length <= 20) spec.dateFormat = val;
                    break;
                case 'wrapText':
                case 'clearRedundant':
                    spec[key] = val === true || val === 'true' || val === 1;
                    break;
                case 'title':
                    if (val === null) spec.title = null;
                    else if (typeof val === 'string' && val.length <= 100) spec.title = val;
                    break;
                case 'headerRows':
                    if (val === null) spec.headerRows = 1;
                    else {
                        const n = Math.round(Number(val));
                        if (Number.isFinite(n) && n >= 1 && n <= 5) spec.headerRows = n;
                    }
                    break;
                case 'autofit':
                case 'freezeHeader':
                case 'addBorders':
                case 'bandedRows':
                    spec[key] = val === true || val === 'true' || val === 1;
                    break;
                default:
                    break;
            }
        }
        return spec;
    }

    // ============================================================
    // 渲染辅助：VBA 安全
    // ============================================================
    function rgb(arr) {
        if (!Array.isArray(arr) || arr.length < 3) return 'RGB(0,0,0)';
        return `RGB(${arr[0]}, ${arr[1]}, ${arr[2]})`;
    }

    function vbaStr(s) {
        if (s == null) return '""';
        return `"${String(s).replace(/"/g, '""')}"`;
    }

    function alignConst(alignment) {
        switch (alignment) {
            case 'left': return 'xlLeft';
            case 'right': return 'xlRight';
            case 'center':
            default: return 'xlCenter';
        }
    }

    // ============================================================
    // 核心：根据 Spec 渲染确定性的 VBA 美化代码
    // ============================================================
    /**
     * @param {Object} spec - 规范化后的 spec
     * @param {Object} ctx - { selection: {address, sheetName}, activeSheet }
     * @returns {string} 完整 VBA 代码
     */
    function renderBeautifyVBA(spec, ctx) {
        const theme = THEMES[spec.theme] || THEMES[DEFAULT_SPEC.theme];
        // 自定义颜色优先：headerColor/accentColor/bandColor/bodyColor 覆盖主题色。
        // 这样 LLM 输出任意颜色词（红/黄/粉/金...）或 RGB 时，样式真正变化，而不是固定主题。
        const palette = {
            header: spec.headerColor || theme.header,
            accent: spec.accentColor || theme.accent,
            band: spec.bandColor || theme.band,
            body: spec.bodyColor || theme.body,
            border: theme.border
        };
        const themeLabel = spec.theme === 'custom'
            ? `自定义 RGB(${palette.header.join(',')})`
            : theme.name;
        // 边框版式：spec.borderStyle 优先；未给时按旧 addBorders 布尔兼容
        const borderStyle = spec.borderStyle || (spec.addBorders === false ? 'none' : 'all');
        // 表头版式：spec.headerLayout 优先；未给时按旧 headerStyle 兼容
        const hdrLayout = spec.headerLayout || (spec.headerStyle === 'italic' ? 'italic'
            : (spec.headerStyle === 'colored_only' ? 'light'
                : (spec.headerStyle === 'bold_only' ? 'underline' : 'deep')));
        const lines = [];
        lines.push(`' ============================================================`);
        lines.push(`' 智能表格美化（Skill Spec 确定性渲染）`);
        lines.push(`' 主题: ${themeLabel} · 表头: ${spec.headerStyle} · 对齐: ${spec.alignment}`);
        lines.push(`' ============================================================`);
        lines.push(`Sub SmartBeautifyTable()`);
        lines.push(`    Dim selRng As Range`);
        lines.push(`    Dim hdrRng As Range`);
        lines.push(`    Dim dataRng As Range`);
        lines.push(`    Dim hr As Long`);
        lines.push(`    Dim scrUpd As Boolean`);
        lines.push(`    Dim calcMode As Long`);
        lines.push(``);
        lines.push(`    scrUpd = Application.ScreenUpdating`);
        lines.push(`    calcMode = Application.Calculation`);
        lines.push(`    Application.ScreenUpdating = False`);
        lines.push(`    Application.Calculation = xlCalculationManual`);
        lines.push(`    On Error GoTo ErrHandler`);
        lines.push(``);
        lines.push(`    Set selRng = Selection`);
        lines.push(`    If selRng Is Nothing Then`);
        lines.push(`        Debug.Print "[SmartBeautifier] 请先选中需要美化的数据区域（含表头）"`);
        lines.push(`        GoTo CleanExit`);
        lines.push(`    End If`);
        lines.push(``);
        // ============ 表头自动定位（内容识别 + 向上扩展）============
        // 不依赖规划层猜测：直接读选中区及其上方的连续非空行，依单元格内容判断哪些行是"表头"。
        // 覆盖三种场景：① 选区顶部即表头（单/多行）；② 表头在选区上方（向上扩展并入）；
        // ③ 多行表头（合并标题行 + 列头行）。彻底解决"不知道表头在哪"。
        lines.push(`    Dim r0 As Long, r1 As Long, c0 As Long, c1 As Long`);
        lines.push(`    Dim rp As Long, hh As Long, extRows As Long`);
        lines.push(`    r0 = selRng.Row`);
        lines.push(`    r1 = selRng.Row + selRng.Rows.Count - 1`);
        lines.push(`    c0 = selRng.Column`);
        lines.push(`    c1 = selRng.Column + selRng.Columns.Count - 1`);
        lines.push(`    ' 向上扩展：选区上方连续非空行并入表范围（最多 3 行）`);
        lines.push(`    extRows = 0`);
        lines.push(`    rp = r0 - 1`);
        lines.push(`    Do While rp >= 1 And extRows < 3`);
        lines.push(`        If Application.WorksheetFunction.CountA(Range(Cells(rp, c0), Cells(rp, c1))) > 0 Then`);
        lines.push(`            extRows = extRows + 1`);
        lines.push(`            rp = rp - 1`);
        lines.push(`        Else`);
        lines.push(`            Exit Do`);
        lines.push(`        End If`);
        lines.push(`    Loop`);
        lines.push(`    If extRows > 0 Then`);
        lines.push(`        Set selRng = Range(Cells(r0 - extRows, c0), Cells(r1, c1))`);
        lines.push(`        r0 = r0 - extRows`);
        lines.push(`    End If`);
        lines.push(`    ' 依内容计算表头高度（合并标题行 / 文本列头行计入表头）`);
        lines.push(`    hh = 0`);
        lines.push(`    rp = r0`);
        lines.push(`    Do While rp <= r1 And hh < 5`);
        lines.push(`        If IsHeaderRow(rp, c0, c1) Then`);
        lines.push(`            hh = hh + 1`);
        lines.push(`            rp = rp + 1`);
        lines.push(`        Else`);
        lines.push(`            Exit Do`);
        lines.push(`        End If`);
        lines.push(`    Loop`);
        lines.push(`    If hh < 1 Then hh = 1`);
        lines.push(`    hr = hh`);
        lines.push(``);
        // 清理冗余格式（在全部样式之前执行）
        if (spec.clearRedundant) {
            lines.push(`    ' ---- 清理冗余格式 ----`);
            lines.push(`    selRng.ClearFormats`);
            lines.push(``);
        }
        // 标题插行：仅当表头仅 1 行（无多行表头/无合并标题）且确需新增标题时才插行，
        // 避免把已有标题重复叠加、把列头/数据整体顶乱（实测"多行表头看起来和之前一样"）。
        if (spec.title) {
            lines.push(`    ' ---- 插入标题行（仅当表头为单行且非合并标题）----`);
            lines.push(`    If hr = 1 And Not selRng.Rows(1).MergeCells Then`);
            lines.push(`        Rows(selRng.Row).Insert Shift:=xlDown`);
            lines.push(`        With selRng.Rows(1)`);
            lines.push(`            .Merge`);
            lines.push(`            .Value = ${vbaStr(spec.title)}`);
            lines.push(`            .Font.Size = 14`);
            lines.push(`            .Font.Bold = True`);
            lines.push(`            .Font.Color = ${rgb(palette.header)}`);
            lines.push(`            .HorizontalAlignment = xlCenter`);
            lines.push(`            .RowHeight = 30`);
            lines.push(`        End With`);
            lines.push(`    End If`);
            lines.push(``);
        }
        lines.push(`    ' ---- 表头样式（${hdrLayout}）----`);
        lines.push(`    Set hdrRng = selRng.Rows("1:" & hr)`);
        lines.push(`    With hdrRng`);
        if (hdrLayout === 'deep') {
            lines.push(`        .Interior.Color = ${rgb(palette.header)}`);
            lines.push(`        .Font.Color = RGB(255, 255, 255)`);
            lines.push(`        .Font.Bold = True`);
        } else if (hdrLayout === 'light') {
            lines.push(`        .Interior.Color = ${rgb(lightenColor(palette.header, 0.78))}`);
            lines.push(`        .Font.Color = ${rgb(palette.header)}`);
            lines.push(`        .Font.Bold = True`);
        } else if (hdrLayout === 'underline') {
            // 无底色：只靠加粗 + 底部粗线区分表头（学术/极简场景）
            lines.push(`        .Interior.Pattern = xlNone`);
            lines.push(`        .Font.Color = ${rgb(palette.header)}`);
            lines.push(`        .Font.Bold = True`);
            if (!spec.title) {
                lines.push(`        .Borders(xlEdgeBottom).LineStyle = xlContinuous`);
                lines.push(`        .Borders(xlEdgeBottom).Weight = xlMedium`);
                lines.push(`        .Borders(xlEdgeBottom).Color = ${rgb(palette.header)}`);
            }
        } else if (hdrLayout === 'italic') {
            lines.push(`        .Interior.Pattern = xlNone`);
            lines.push(`        .Font.Color = ${rgb(palette.header)}`);
            lines.push(`        .Font.Italic = True`);
        }
        lines.push(`        .Font.Size = 11`);
        lines.push(`        .VerticalAlignment = ${spec.valign === 'top' ? 'xlTop' : 'xlCenter'}`);
        lines.push(`        .HorizontalAlignment = xlCenter`);
        if (spec.rowHeight) lines.push(`        .RowHeight = ${spec.rowHeight}`);
        // 三线表：表头顶线（第一条线）
        if (borderStyle === 'threeLine') {
            lines.push(`        .Borders.LineStyle = xlNone`);
            lines.push(`        .Borders(xlEdgeTop).LineStyle = xlContinuous`);
            lines.push(`        .Borders(xlEdgeTop).Weight = xlMedium`);
            lines.push(`        .Borders(xlEdgeTop).Color = ${rgb(palette.header)}`);
        }
        lines.push(`    End With`);
        lines.push(``);
        // 数据区域
        lines.push(`    ' ---- 数据区域样式 ----`);
        lines.push(`    If selRng.Rows.Count > hr Then`);
        lines.push(`        Set dataRng = selRng.Offset(hr, 0).Resize(selRng.Rows.Count - hr, selRng.Columns.Count)`);
        lines.push(`        With dataRng`);
        // 字体用 spec（默认微软雅黑）
        lines.push(`            .Font.Name = ${vbaStr(spec.fontName || '微软雅黑')}`);
        lines.push(`            .Font.Size = ${spec.fontSize || 10}`);
        // 正文数据区强制常规（取消原表自带的大字/加粗），
        // 避免"有些单元加粗、有些不加粗"混排；表头加粗由表头区块单独控制。
        lines.push(`            .Font.Bold = False`);
        lines.push(`            .HorizontalAlignment = ${alignConst(spec.alignment)}`);
        lines.push(`            .VerticalAlignment = ${spec.valign === 'top' ? 'xlTop' : 'xlCenter'}`);
        // 数据区默认不上底色。只有用户点名表身/隔行并写出 bodyColor 时才填。
        if (spec.noContentBg || !spec.bodyColor || borderStyle === 'none' || borderStyle === 'threeLine') {
            lines.push(`            .Interior.Pattern = xlNone`);
        } else {
            lines.push(`            .Interior.Color = ${rgb(spec.bodyColor)}`);
        }
        if (spec.wrapText) {
            lines.push(`            .WrapText = True`);
        }
        if (spec.rowHeight) lines.push(`            .RowHeight = ${spec.rowHeight}`);
        // ── 边框版式：all 全网格 / horizontal 仅横线 / outer 仅外框 / threeLine 三线表 / none 无 ──
        if (borderStyle === 'all') {
            lines.push(`            .Borders.LineStyle = xlContinuous`);
            lines.push(`            .Borders.Color = ${rgb(palette.border)}`);
            lines.push(`            .Borders.Weight = xlThin`);
        } else if (borderStyle === 'none') {
            lines.push(`            .Borders.LineStyle = xlNone`);
        } else if (borderStyle === 'horizontal') {
            lines.push(`            .Borders.LineStyle = xlNone`);
            lines.push(`            .Borders(xlInsideHorizontal).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlInsideHorizontal).Color = ${rgb(palette.border)}`);
            lines.push(`            .Borders(xlInsideHorizontal).Weight = xlThin`);
        } else if (borderStyle === 'outer') {
            lines.push(`            .Borders.LineStyle = xlNone`);
            lines.push(`            .Borders(xlEdgeTop).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeTop).Color = ${rgb(palette.border)}`);
            lines.push(`            .Borders(xlEdgeBottom).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeBottom).Color = ${rgb(palette.border)}`);
            lines.push(`            .Borders(xlEdgeLeft).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeLeft).Color = ${rgb(palette.border)}`);
            lines.push(`            .Borders(xlEdgeRight).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeRight).Color = ${rgb(palette.border)}`);
        } else if (borderStyle === 'threeLine') {
            // 学术三线表：顶线 + 表头下线 + 底线（中粗），无竖线、无内部横线
            lines.push(`            .Borders.LineStyle = xlNone`);
            lines.push(`            .Borders(xlEdgeTop).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeTop).Weight = xlMedium`);
            lines.push(`            .Borders(xlEdgeTop).Color = ${rgb(palette.header)}`);
            lines.push(`            .Borders(xlEdgeBottom).LineStyle = xlContinuous`);
            lines.push(`            .Borders(xlEdgeBottom).Weight = xlMedium`);
            lines.push(`            .Borders(xlEdgeBottom).Color = ${rgb(palette.header)}`);
        }
        if (spec.bandedRows && !spec.noContentBg) {
            lines.push(`            .FormatConditions.Delete`);
            lines.push(`            .FormatConditions.Add Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0"`);
            lines.push(`            .FormatConditions(1).Interior.Color = ${rgb(palette.band)}`);
        }
        if (spec.numberFormat) {
            lines.push(`            .NumberFormat = ${vbaStr(spec.numberFormat)}`);
        }
        if (spec.dateFormat) {
            lines.push(`            .NumberFormat = ${vbaStr(spec.dateFormat)}`);
        }
        lines.push(`        End With`);
        lines.push(`    End If`);
        lines.push(``);
        // ── 首列强调：左侧第一列当行标题（加粗 + 极浅底），整表首列一起处理 ──
        if (spec.boldFirstColumn) {
            lines.push(`    ' ---- 首列强调 ----`);
            lines.push(`    With selRng.Columns(1)`);
            lines.push(`        .Font.Bold = True`);
            lines.push(`        .Interior.Color = ${rgb(lightenColor(palette.header, 0.90))}`);
            lines.push(`    End With`);
            lines.push(``);
        }
        // ── 末行（合计行）强调：加粗 + 顶部中粗线 + 浅底 ──
        if (spec.emphasizeLastRow) {
            lines.push(`    ' ---- 末行(合计行)强调 ----`);
            lines.push(`    With selRng.Rows(selRng.Rows.Count)`);
            lines.push(`        .Font.Bold = True`);
            lines.push(`        .Interior.Color = ${rgb(palette.band)}`);
            lines.push(`        .Borders(xlEdgeTop).LineStyle = xlContinuous`);
            lines.push(`        .Borders(xlEdgeTop).Weight = xlMedium`);
            lines.push(`        .Borders(xlEdgeTop).Color = ${rgb(palette.header)}`);
            lines.push(`    End With`);
            lines.push(``);
        }
        // ── 正文数字右对齐：数值单元格右对齐、文本仍按 alignment ──
        if (spec.bodyNumberRight) {
            lines.push(`    ' ---- 数值单元格右对齐 ----`);
            lines.push(`    If Not dataRng Is Nothing Then`);
            lines.push(`        Dim cl As Range`);
            lines.push(`        For Each cl In dataRng.Cells`);
            lines.push(`            If Len(cl.Text) > 0 Then`);
            lines.push(`                If IsNumeric(cl.Value) Then`);
            lines.push(`                    cl.HorizontalAlignment = xlRight`);
            lines.push(`                Else`);
            lines.push(`                    cl.HorizontalAlignment = ${alignConst(spec.alignment)}`);
            lines.push(`                End If`);
            lines.push(`            End If`);
            lines.push(`        Next cl`);
            lines.push(`    End If`);
            lines.push(``);
        }
        // 条件格式
        if (spec.conditionalFormat) {
            lines.push(`    ' ---- 条件格式 ----`);
            lines.push(`    If Not dataRng Is Nothing Then`);
            lines.push(`        dataRng.FormatConditions.Delete`);
            if (spec.conditionalFormat === 'dataBar') {
                lines.push(`        dataRng.FormatConditions.AddDatabar`);
                lines.push(`        dataRng.FormatConditions(dataRng.FormatConditions.Count).BarColor.Color = ${rgb(palette.accent)}`);
            } else if (spec.conditionalFormat === 'colorScale') {
                lines.push(`        dataRng.FormatConditions.AddColorScale ColorScaleType:=3`);
            } else if (spec.conditionalFormat === 'highlight_top') {
                lines.push(`        dataRng.FormatConditions.AddTop10`);
                lines.push(`        dataRng.FormatConditions(dataRng.FormatConditions.Count).Interior.Color = ${rgb(palette.accent)}`);
            } else if (spec.conditionalFormat === 'highlight_duplicates') {
                lines.push(`        dataRng.FormatConditions.AddUniqueValues`);
                lines.push(`        dataRng.FormatConditions(dataRng.FormatConditions.Count).DupeUnique = xlDuplicate`);
                lines.push(`        dataRng.FormatConditions(dataRng.FormatConditions.Count).Interior.Color = RGB(255, 199, 206)`);
            }
            lines.push(`    End If`);
            lines.push(``);
        }
        // 冻结表头
        if (spec.freezeHeader) {
            lines.push(`    ' ---- 冻结表头 ----`);
            lines.push(`    With ActiveWindow`);
            lines.push(`        .SplitRow = selRng.Row + hr - 1`);
            lines.push(`        .FreezePanes = True`);
            lines.push(`    End With`);
            lines.push(``);
        }
        // 自动列宽
        if (spec.autofit) {
            lines.push(`    ' ---- 自动列宽 ----`);
            lines.push(`    selRng.EntireColumn.AutoFit`);
            lines.push(``);
        }
        lines.push(`CleanExit:`);
        lines.push(`    Application.ScreenUpdating = scrUpd`);
        lines.push(`    Application.Calculation = calcMode`);
        lines.push(`    Exit Sub`);
        lines.push(`ErrHandler:`);
        lines.push(`    Application.ScreenUpdating = scrUpd`);
        lines.push(`    Application.Calculation = calcMode`);
        lines.push(`    Debug.Print "[SmartBeautifier] 运行时错误: " & Err.Description`);
        lines.push(`End Sub`);
        // 内容识别辅助：判断某行是否为表头行（合并标题行 / 整行均为文本的列头行）
        lines.push(``);
        lines.push(`Private Function IsHeaderRow(ByVal r As Long, ByVal c0 As Long, ByVal c1 As Long) As Boolean`);
        lines.push(`    Dim cc As Long, nc As Long, t As Long, m As Long`);
        lines.push(`    Dim v As Variant`);
        lines.push(`    nc = 0: t = 0: m = 0`);
        lines.push(`    For cc = c0 To c1`);
        lines.push(`        If Cells(r, cc).MergeCells Then m = m + 1`);
        lines.push(`        v = Cells(r, cc).Value`);
        lines.push(`        If Not IsEmpty(v) Then`);
        lines.push(`            nc = nc + 1`);
        lines.push(`            If VarType(v) = vbString Then t = t + 1`);
        lines.push(`        End If`);
        lines.push(`    Next cc`);
        lines.push(`    If nc = 0 Then IsHeaderRow = False: Exit Function`);
        lines.push(`    If m > 0 And nc <= 1 Then IsHeaderRow = True: Exit Function   ' 合并标题行`);
        lines.push(`    IsHeaderRow = (t / nc) >= 0.999   ' 整行无数值单元格（全文本）→ 列头行；数据行含数值列则停`);
        lines.push(`End Function`);
        return lines.join('\n');
    }

    // ============================================================
    // 格式整理渲染器（SmartTableFormatter）
    // 与 Beautify 互补：Beautify 偏"样式美化"（主题/色彩/条件格式），
    // Format 偏"格式整理"（字体/对齐/边框/数字日期/冻结/换行/清理冗余）。
    // ============================================================

    /** 格式整理默认 spec */
    const DEFAULT_FORMAT_SPEC = {
        fontName: '微软雅黑',
        fontSize: 11,
        headerBold: true,
        alignment: 'center',
        addBorders: true,
        autofit: true,
        freezeHeader: false,
        numberFormat: null,
        dateFormat: null,
        wrapText: false,
        clearRedundant: false,
        headerColor: null,     // 可选：表头背景色（颜色词/RGB）
        accentColor: null
    };

    /** 格式整理字段白名单 */
    const FORMAT_ALLOWED_FIELDS = new Set(Object.keys(DEFAULT_FORMAT_SPEC));

    /**
     * 规范化格式整理 spec（白名单 + 默认合并）
     */
    function normalizeFormatSpec(raw) {
        const spec = { ...DEFAULT_FORMAT_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        for (const key of Object.keys(raw)) {
            if (!FORMAT_ALLOWED_FIELDS.has(key)) continue;
            const val = raw[key];
            switch (key) {
                case 'fontName':
                    if (typeof val === 'string' && val.length <= 30 && /^[\u4e00-\u9fa5A-Za-z0-9\s]+$/.test(val)) spec.fontName = val;
                    break;
                case 'fontSize':
                    if (typeof val === 'number' && val >= 8 && val <= 24) spec.fontSize = Math.round(val);
                    break;
                case 'headerBold':
                case 'addBorders':
                case 'autofit':
                case 'freezeHeader':
                case 'wrapText':
                case 'clearRedundant':
                    spec[key] = val === true || val === 'true' || val === 1;
                    break;
                case 'alignment':
                    if (['left', 'center', 'right'].includes(val)) spec.alignment = val;
                    break;
                case 'numberFormat':
                    if (val === null) spec.numberFormat = null;
                    else if (typeof val === 'string' && /^[0#.,%$€¥\-\s()]*$/.test(val) && val.length <= 20) spec.numberFormat = val;
                    break;
                case 'dateFormat':
                    if (val === null) spec.dateFormat = null;
                    else if (typeof val === 'string' && /^[ymd\-/.\s]*$/i.test(val) && val.length <= 20) spec.dateFormat = val;
                    break;
                case 'headerColor': {
                    const rgb = resolveColor(val);
                    if (rgb) spec.headerColor = rgb;
                    break;
                }
                default:
                    break;
            }
        }
        return spec;
    }

    /**
     * 渲染格式整理的确定性 VBA。
     * @param {Object} spec - 规范化后的格式 spec
     * @returns {string} VBA 代码
     */
    function renderFormatVBA(spec) {
        const lines = [];
        lines.push(`' ============================================================`);
        lines.push(`' 智能表格格式整理（Skill Spec 确定性渲染）`);
        lines.push(`' 字体: ${spec.fontName} ${spec.fontSize}号 · 对齐: ${spec.alignment} · 边框: ${spec.addBorders}`);
        lines.push(`' ============================================================`);
        lines.push(`Sub SmartFormatTable()`);
        lines.push(`    Dim selRng As Range`);
        lines.push(`    Dim dataRng As Range`);
        lines.push(`    Dim scrUpd As Boolean`);
        lines.push(`    Dim calcMode As Long`);
        lines.push(``);
        lines.push(`    scrUpd = Application.ScreenUpdating`);
        lines.push(`    calcMode = Application.Calculation`);
        lines.push(`    Application.ScreenUpdating = False`);
        lines.push(`    Application.Calculation = xlCalculationManual`);
        lines.push(`    On Error GoTo ErrHandler`);
        lines.push(``);
        lines.push(`    Set selRng = Selection`);
        lines.push(`    If selRng Is Nothing Then`);
        lines.push(`        Debug.Print "[SmartFormatter] 请先选中需要整理的表格区域（含表头）"`);
        lines.push(`        GoTo CleanExit`);
        lines.push(`    End If`);
        lines.push(``);
        // 清理冗余格式（可选）
        if (spec.clearRedundant) {
            lines.push(`    ' ---- 清理冗余格式 ----`);
            lines.push(`    selRng.ClearFormats`);
            lines.push(``);
        }
        // 统一字体字号
        lines.push(`    ' ---- 统一字体字号 ----`);
        lines.push(`    With selRng.Font`);
        lines.push(`        .Name = ${vbaStr(spec.fontName)}`);
        lines.push(`        .Size = ${spec.fontSize}`);
        lines.push(`    End With`);
        lines.push(``);
        // 表头：加粗 + 可选背景色 + 底部边框
        lines.push(`    ' ---- 表头 ----`);
        lines.push(`    With selRng.Rows(1)`);
        if (spec.headerBold) lines.push(`        .Font.Bold = True`);
        if (spec.headerColor) {
            lines.push(`        .Interior.Color = ${rgb(spec.headerColor)}`);
            lines.push(`        .Font.Color = RGB(255, 255, 255)`);
        }
        lines.push(`        .HorizontalAlignment = xlCenter`);
        lines.push(`        .Borders(xlEdgeBottom).LineStyle = xlContinuous`);
        lines.push(`        .Borders(xlEdgeBottom).Weight = xlMedium`);
        if (spec.headerColor) lines.push(`        .Borders(xlEdgeBottom).Color = ${rgb(spec.headerColor)}`);
        lines.push(`    End With`);
        lines.push(``);
        // 数据区域：对齐 + 换行 + 边框
        lines.push(`    ' ---- 数据区域 ----`);
        lines.push(`    If selRng.Rows.Count > 1 Then`);
        lines.push(`        Set dataRng = selRng.Offset(1, 0).Resize(selRng.Rows.Count - 1, selRng.Columns.Count)`);
        lines.push(`        With dataRng`);
        lines.push(`            .HorizontalAlignment = ${alignConst(spec.alignment)}`);
        lines.push(`            .VerticalAlignment = xlCenter`);
        if (spec.wrapText) lines.push(`            .WrapText = True`);
        if (spec.addBorders) {
            lines.push(`            .Borders.LineStyle = xlContinuous`);
            lines.push(`            .Borders.Color = RGB(180, 180, 180)`);
            lines.push(`            .Borders.Weight = xlThin`);
        }
        if (spec.numberFormat) lines.push(`            .NumberFormat = ${vbaStr(spec.numberFormat)}`);
        if (spec.dateFormat) lines.push(`            .NumberFormat = ${vbaStr(spec.dateFormat)}`);
        lines.push(`        End With`);
        lines.push(`    End If`);
        lines.push(``);
        // 冻结首行
        if (spec.freezeHeader) {
            lines.push(`    ' ---- 冻结首行 ----`);
            lines.push(`    With ActiveWindow`);
            lines.push(`        .SplitRow = selRng.Row`);
            lines.push(`        .FreezePanes = True`);
            lines.push(`    End With`);
            lines.push(``);
        }
        // 自适应列宽
        if (spec.autofit) {
            lines.push(`    ' ---- 自适应列宽 ----`);
            lines.push(`    selRng.EntireColumn.AutoFit`);
            lines.push(``);
        }
        lines.push(`CleanExit:`);
        lines.push(`    Application.ScreenUpdating = scrUpd`);
        lines.push(`    Application.Calculation = calcMode`);
        lines.push(`    Exit Sub`);
        lines.push(`ErrHandler:`);
        lines.push(`    Application.ScreenUpdating = scrUpd`);
        lines.push(`    Application.Calculation = calcMode`);
        lines.push(`    Debug.Print "[SmartFormatter] 运行时错误: " & Err.Description`);
        lines.push(`End Sub`);
        return lines.join('\n');
    }

    /** 格式整理的规划 prompt */
    function buildFormatPlanningPrompt(skill) {
        return `你是 SmartPilot 的表格格式整理规划师。你的唯一任务：把用户的自然语言需求转成【结构化 JSON spec】，绝不输出代码。

【用户需求】将从下方 user 消息提供。

【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块标记，不要解释），字段如下：
{
  "fontName": "微软雅黑",               // 字体（用户说字体时用；否则默认）
  "fontSize": 11,                       // 字号（8-24）
  "headerBold": true,                   // 表头加粗
  "headerColor": null,                  // 表头背景色：RGB 数组 [R,G,B] 或颜色词（"深蓝"、"gold"）；用户说颜色就用
  "alignment": "center",                // left | center | right
  "addBorders": true,                   // 标准化边框
  "autofit": true,                      // 自适应列宽
  "freezeHeader": false,                // 冻结首行
  "numberFormat": null,                 // null 或 "#,##0"、"0.00%"（数字格式）
  "dateFormat": null,                   // null 或 "yyyy-mm-dd"（日期格式）
  "wrapText": false,                    // 自动换行
  "clearRedundant": false               // 清理冗余格式
}

【颜色规则】
- 用户说任何颜色（任意中文/英文词），用颜色知识转 RGB 数组，如 "深蓝"→headerColor:[30,70,140]
- 用户没说颜色 → 省略 headerColor

【规则】
- 所有字段可省略，省略用默认
- 只输出 JSON，不要任何其他文字`;
    }

    /**
     * 格式整理主入口：LLM spec → 规范化 → 确定性 VBA
     */
    function planFormat(llmOutput) {
        const rawSpec = extractSpec(llmOutput);
        const spec = normalizeFormatSpec(rawSpec);
        try {
            const code = renderFormatVBA(spec);
            return { ok: true, spec, code };
        } catch (e) {
            return { ok: false, spec, code: '', error: String(e?.message || e) };
        }
    }

    // ============================================================
    // 主入口
    // ============================================================

    /**
     * 从 LLM 输出解析 Spec JSON（支持 ```json 代码块 / 裸 JSON）
     * @param {string|Object} llmOutput
     * @returns {Object|null}
     */
    function extractSpec(llmOutput) {
        if (!llmOutput) return null;
        if (typeof llmOutput === 'object') return llmOutput;
        const m = String(llmOutput).match(/```(?:json)?\s*([\s\S]*?)```/);
        const text = m ? m[1].trim() : String(llmOutput).trim();
        try {
            const parsed = JSON.parse(text);
            return parsed && typeof parsed === 'object' ? parsed : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 规划并渲染：LLM spec → 规范化 → 确定性 VBA
     * @param {string|Object} llmOutput - LLM 输出的 spec（JSON 或对象）
     * @param {Object} [ctx] - { selection, activeSheet }
     * @returns {{ok: boolean, spec: Object, code: string, error?: string}}
     */
    function planAndRender(llmOutput, ctx = {}) {
        const rawSpec = extractSpec(llmOutput);
        const spec = normalizeSpec(rawSpec);
        try {
            const code = renderBeautifyVBA(spec, ctx);
            return { ok: true, spec, code };
        } catch (e) {
            return { ok: false, spec, code: '', error: String(e?.message || e) };
        }
    }

    /**
     * 从用户消息直接生成 Spec（无 LLM 时的确定性兜底：动态颜色 + 关键词映射）
     * @param {string} userMessage
     * @returns {Object}
     */
    function specFromMessage(userMessage) {
        const spec = { ...DEFAULT_SPEC };
        const msg = String(userMessage || '').toLowerCase();
        // 【样式方案】先按"版式 + 配色 + 密度 + 开关"展开预设（如"三线表""紧凑""首列加粗""深表头"），
        // 只有在识别出明确样式信号时才套用；后续显式字段（颜色词/对齐/斑马…）再覆盖。
        const _P = presetsModule();
        if (_P) {
            try {
                const det = _P.detect(msg);
                const hasSignal = !!(det.layoutId || det.paletteId || det.density ||
                    det.zebra !== null || det.firstColBold !== null || det.lastRowEmphasis !== null ||
                    det.bodyNumberRight !== null || det.headerMode || det.valign);
                if (hasSignal) {
                    const ex = _P.toExcelSpec(_P.resolve(msg));
                    Object.keys(ex).forEach(k => {
                        if (Object.prototype.hasOwnProperty.call(DEFAULT_SPEC, k)) spec[k] = ex[k];
                    });
                }
            } catch { /* 预设模块异常 → 走原有兜底 */ }
        }
        // 识别任意颜色词（红/黄/蓝/粉/金/绿...）→ 自定义表头色
        const colorHits = extractColorWords(msg);
        if (colorHits.length > 0) {
            // 取最具体的颜色词（优先长的，如"深蓝色"胜过"蓝色"）
            colorHits.sort((a, b) => b.word.length - a.word.length);
            const best = colorHits[0];
            spec.theme = 'custom';
            spec.headerColor = best.rgb;
            spec.accentColor = deriveAccent(best.rgb);
            // 颜色词只决定表头。表身底色、隔行色仅在用户点名时才派生。
            if (/表身|数据区|斑马|隔行|条纹/.test(msg)) {
                spec.bandColor = lightenColor(best.rgb, 0.82);
                spec.bodyColor = lightenColor(best.rgb, 0.92);
            }
        } else {
            // 无明确颜色词 → 保持默认
            spec.theme = 'business_blue';
        }
        // 用户明确"移除内容底色/不要底色" → 清空数据区填充，杜绝再加主题浅底
        if (/(移除|去掉|去掉|去除|清除|取消|删去|不要).*(内容底色|内容区底色|数据区底色|底色|填充色|背景色|背景)/.test(msg)) {
            spec.noContentBg = true;
            spec.bodyColor = null;
        }
        // 表头
        if (/加粗/.test(msg) && !/颜色/.test(msg)) spec.headerStyle = 'bold_only';
        // 多行表头（标题行+字段行的原生结构在选中区内，勿与"插入标题行"混淆）
        if (/(\d+)\s*行表头|多行表头|双行表头/.test(msg)) {
            const m = msg.match(/(\d+)\s*行表头/);
            spec.headerRows = Math.max(1, Math.min(5, m ? Number(m[1]) : 2));
        }
        // 对齐
        if (/左对齐/.test(msg)) spec.alignment = 'left';
        else if (/右对齐/.test(msg)) spec.alignment = 'right';
        // 斑马纹
        if (/斑马|隔行|条纹/.test(msg)) spec.bandedRows = true;
        // 冻结表头
        if (/冻结/.test(msg)) spec.freezeHeader = true;
        // 数据条
        if (/数据条/.test(msg)) spec.conditionalFormat = 'dataBar';
        // 色阶
        if (/色阶/.test(msg)) spec.conditionalFormat = 'colorScale';
        // 标题（"标题：XXX"）
        const titleMatch = msg.match(/标题[:：]\s*(.{1,40})/);
        if (titleMatch) spec.title = titleMatch[1].trim();
        return spec;
    }

    /** 从主色派生强调色（偏暖/深一档） */
    function deriveAccent(rgb) {
        return [Math.min(255, rgb[0] + 60), Math.min(255, rgb[1] - 20), Math.min(255, rgb[2] - 20)]
            .map(v => Math.max(0, Math.round(v)));
    }

    /** 颜色变浅（用于斑马纹/浅色带） */
    function lightenColor(rgb, ratio) {
        return [
            Math.round(rgb[0] + (255 - rgb[0]) * ratio),
            Math.round(rgb[1] + (255 - rgb[1]) * ratio),
            Math.round(rgb[2] + (255 - rgb[2]) * ratio)
        ];
    }

    // ============================================================
    // 规划 Prompt：让 LLM 输出结构化 Spec（而非代码）
    // ============================================================

    /**
     * 生成"LLM 规划 spec"的 system prompt。
     * 核心：LLM 只负责理解用户意图 → 输出 JSON spec，不写任何代码。
     * 主题选项优先取自 skill.specFields 中声明的 theme.options，
     * 未声明时回退到内置 THEMES。
     * @param {Object} skill - Skill 定义（含 displayName/description/configFields/specFields）
     * @returns {string}
     */
    function buildPlanningPrompt(skill) {
        const themeOptions = resolveThemeOptions(skill);
        return `你是 SmartPilot 的表格美化规划师。你的唯一任务：把用户的自然语言需求转成【结构化 JSON spec】，绝不输出代码。

【用户需求】将从下方 user 消息提供。

【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块标记，不要解释），字段如下：
{
  "theme": null,                      // 可选：${themeOptions} 之一；或用任意颜色词（如 "lavender"、"sunset orange"、"薄荷绿"）
  "headerColor": null,                // 【推荐】表头颜色 = RGB 数组 [R,G,B]，0-255 整数。用户说任何颜色，你都用颜色知识转换成 RGB！
  "accentColor": null,                // 强调色 RGB 数组（数据条/高亮）
  "bandColor": null,                  // 斑马纹浅色 RGB 数组
  "headerStyle": "bold_colored | bold_only | colored_only | italic",
  "alignment": "left | center | right",
  "autofit": true,
  "freezeHeader": false,
  "addBorders": true,
  "bandedRows": false,
  "conditionalFormat": null,          // null | dataBar | colorScale | highlight_top | highlight_duplicates
  "numberFormat": null,               // null 或 "#,##0"（只能数字格式）
  "title": null,                      // null 或首行标题文字。【重要】仅当选中区【没有】标题、用户要"新增标题"时填。
                                      //   若选中区顶部已自带标题/表头块，一律填 null，改用 headerRows 表达，绝不再设 title，
                                      //   否则会多插一行标题、把列头和数据整体顶乱。
  "headerRows": 1,                    // 表头总行数（1 基）。依据选中区数据自动判定：
                                      //   单行表头=1；多行表头（含合并标题行）则填入实际表头行数，
                                      //   如"第1行合并大标题 + 第2行列标题" → 2；纯数据行不计入
  // ── 排版维度（主题不只是颜色！版式/表头/行高/强调都要按用户意图选）──
  "stylePreset": null,                // 【推荐】一键样式方案，填了它其余排版字段会自动派生；也可再单独覆盖
                                      //   classic_grid 经典网格(全边框+浅表头+隔行+居中)
                                      //   deep_header 深表头(深底白字+全边框)
                                      //   three_line 三线表(学术/论文：仅顶线+表头下线+底线，无竖线无底色)
                                      //   borderless 无边框留白(去网格，靠行距分区，宽松)
                                      //   card_stripe 卡片条纹(整行浅底交替，无竖线，宽松)
                                      //   underline_header 下划线表头(表头无底色只留下边线，仅横线)
                                      //   first_col_focus 首列强调(首列加粗+浅底，正文左对齐)
                                      //   total_row 合计行(末行加粗+上边线+浅底，数字右对齐)
                                      //   compact_report 紧凑报表(行高压紧，数字右对齐)
                                      //   magazine 杂志宽松(深表头+宽松行距+横向分隔)
  "borderStyle": "all",               // all 全网格 | none 无边框 | horizontal 仅横线 | outer 仅外框 | threeLine 三线表
  "headerLayout": "deep",             // deep 深底白字 | light 浅底深字 | underline 仅底部粗线(无底色) | italic 斜体表头
  "rowHeight": null,                  // 行高磅值：紧凑 15 / 标准 20 / 宽松 28；null = 不改行高
  "valign": "middle",                 // middle 垂直居中 | top 顶端对齐
  "boldFirstColumn": false,           // 首列加粗（左侧标题列）
  "emphasizeLastRow": false,          // 末行（合计行）加粗 + 上边线 + 浅底
  "bodyNumberRight": false            // 正文数值单元格右对齐（文本列仍按 alignment）
}

【颜色规则 - 最重要】
- 用户说任何颜色（任意中文/英文颜色词、色系、风格色），你必须用颜色知识自行转换为 RGB 数组，如：
  - "深蓝色" → headerColor:[30,70,140]
  - "薄荷绿" → headerColor:[100,190,160]
  - "lavender" → headerColor:[140,110,200]
  - "夕阳橙" → headerColor:[255,120,60]
  - "商务黑金" → headerColor:[40,40,40], accentColor:[212,175,55]
- 优先用 headerColor/accentColor/bandColor 输出 RGB；theme 只在你认为适合预设主题时用
- 用户没说颜色 → 全部颜色字段省略（默认商务蓝）
- 只输出 JSON，不要任何其他文字`;
    }

    /**
     * 从 skill.specFields 中解析主题选项列表。
     * 优先取 specFields 里 name=theme 的 options；未声明时回退到内置 THEMES。
     * @param {Object} skill
     * @returns {string} 如 "business_blue(商务蓝)、tech_green(科技绿)"
     */
    function resolveThemeOptions(skill) {
        const specFields = skill?.specFields || skill?.SpecFields || [];
        if (Array.isArray(specFields) && specFields.length > 0) {
            const themeField = specFields.find(f => String(f.name || f.Name || '') === 'theme');
            if (themeField && Array.isArray(themeField.options) && themeField.options.length > 0) {
                return themeField.options
                    .map(opt => {
                        const id = opt && opt.value !== undefined && opt.value !== null ? opt.value : opt;
                        const name = opt && opt.label ? opt.label : id;
                        return String(id) + (name && String(name) !== String(id) ? `(${name})` : '');
                    })
                    .join('、');
            }
        }
        return Object.entries(THEMES).map(([id, t]) => `${id}(${t.name})`).join('、');
    }

    // ============================================================
    // 
    // SmartChartWizard（图表）/ SmartDataCleaner（清洗）等 skill 不走 VBA 渲染，
    // 而是 LLM 输出【操作 spec】→ 直接调用 MCP 工具链（excel_create_chart 等）。
    // 引擎负责：spec 规范化（白名单）+ 生成可执行的工具调用计划。
    // ============================================================

    /** 图表 spec 默认值 */
    const DEFAULT_CHART_SPEC = {
        chartType: 'xlColumnClustered',  // Excel 图表类型枚举
        title: null,
        dataRange: null,                 // 如 "Sheet1!A1:B10"（缺省用当前选区）
        sheet: null,                     // 图表放置工作表
        style: 'business_blue',          // 主题样式（复用 THEMES）
        showLegend: true,
        showDataLabels: false,
        comboChart: false,               // 组合图（主柱+次折）
        secondaryValueColumn: null       // 组合图次轴列
    };

    /** 图表类型白名单（对应 excel_create_chart 的 Excel 枚举） */
    const CHART_TYPES = new Set([
        'xlColumnClustered', 'xlColumnStacked', 'xlColumnStacked100',
        'xlBarClustered', 'xlBarStacked', 'xlBarStacked100',
        'xlLine', 'xlLineMarkers', 'xlArea',
        'xlPie', 'xlDoughnut', 'xlXYScatter',
        'xlRadar', 'xlRadarFilled', 'xlBubble',
        'xlSunburst', 'xlTreemap', 'xlHistogram',
        'xlBoxwhisker', 'xlFunnel', 'xlWaterfall'
    ]);

    /** 清洗 spec 默认值 */
    const DEFAULT_CLEAN_SPEC = {
        operations: [],       // ['removeDuplicates','fillEmpty','trimSpaces','normalizeDates','normalizeNumbers','standardizeText']
        sheet: null,
        range: null,
        fillMethod: 'ffill',  // 空值填充方式: ffill(前值) | zero(0) | delete(删除行) | placeholder(占位)
        dedupeColumns: null   // 去重依据列（null=整行去重）
    };

    /** 清洗+表结构变换操作白名单
     *  —— 前 9 个走 excel_clean_data 统一引擎；后 5 个是独立原子工具（split/merge 系列），
     *     skill-clean-crud.js 的 executeCleanSingleOp 里每个 op 都有独立 case 直接调对应 MCP 工具。 */
    const CLEAN_OPERATIONS = new Set([
        'removeDuplicates', 'fillEmpty', 'trimSpaces',
        'normalizeDates', 'normalizeNumbers', 'standardizeText',
        'removeEmptyRows', 'splitColumn', 'mergeColumns',
        // —— 表结构变换（独立原子工具，非 excel_clean_data）——
        'splitByColumn', 'splitWorkbook', 'splitByRows',
        'mergeSheets', 'mergeWorkbooks'
    ]);

    /** 增删改查（CRUD）操作族白名单 */
    const CRUD_OPERATIONS = new Set([
        'read', 'write', 'insert', 'delete', 'clear'
    ]);

    /** CRUD spec 默认值 */
    const DEFAULT_CRUD_SPEC = {
        operation: 'read',        // read | write | insert | delete | clear
        sheet: null,
        range: null,              // read/write/clear 目标，如 "A1:B10" / "第3行" / "C列"
        value: null,              // write 单值辅助
        values: null,             // write 二维数组
        row: null,                // insert/delete 行锚点（1 基）
        column: null,             // insert/delete 列锚点（1 基）
        count: 1,                 // insert/delete 数量
        append: false,            // insert：true=追加到表尾
        clearContent: true,       // clear：清内容
        clearFormat: true         // clear：清格式
    };

    /**
     * 规范化图表 spec（白名单 + 默认合并）
     */
    function normalizeChartSpec(raw) {
        const spec = { ...DEFAULT_CHART_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        const styleColor = resolveColor(raw.style || raw.theme);
        if (styleColor) spec.style = raw.style || 'custom';
        // chartType 归一化：兼容 LLM 的各种输出形态
        //   - 标准枚举：xlPie / xlColumnClustered
        //   - 英文无前缀：Pie / Column Clustered / columnclustered
        //   - 中文：饼图 / 柱状图 / 折线图
        spec.chartType = normalizeChartTypeValue(raw.chartType) || spec.chartType;
        if (typeof raw.title === 'string' && raw.title.length <= 100) spec.title = raw.title;
        if (typeof raw.dataRange === 'string' && raw.dataRange.length <= 100) spec.dataRange = raw.dataRange;
        if (typeof raw.sheet === 'string' && raw.sheet.length <= 50) spec.sheet = raw.sheet;
        if (typeof raw.secondaryValueColumn === 'string' && raw.secondaryValueColumn.length <= 50) spec.secondaryValueColumn = raw.secondaryValueColumn;
        if (raw.showLegend !== undefined) spec.showLegend = raw.showLegend === true || raw.showLegend === 'true';
        if (raw.showDataLabels !== undefined) spec.showDataLabels = raw.showDataLabels === true || raw.showDataLabels === 'true';
        if (raw.comboChart !== undefined) spec.comboChart = raw.comboChart === true || raw.comboChart === 'true';
        return spec;
    }

    /**
     * 图表类型值归一化：处理 LLM 的各种输出形态。
     * 支持：
     *   - 标准枚举：xlPie / xlColumnClustered（大小写不敏感）
     *   - 英文无前缀：Pie / Column Clustered / columnclustered
     *   - 中文：饼图 / 柱状图 / 折线图 / 占比 / 趋势
     * @param {*} val - LLM 输出的 chartType
     * @returns {string|null} 标准枚举名或 null（无法识别）
     */
    function normalizeChartTypeValue(val) {
        if (!val) return null;
        if (typeof val === 'string') {
            // 1. 精确枚举（忽略大小写 + 去空格）
            const cleaned = val.trim().replace(/\s+/g, '');
            for (const t of CHART_TYPES) {
                if (t.toLowerCase() === cleaned.toLowerCase()) return t;
            }
            // 2. 英文无 xl 前缀 → 补前缀
            const prefixed = 'xl' + cleaned;
            for (const t of CHART_TYPES) {
                if (t.toLowerCase() === prefixed.toLowerCase()) return t;
            }
            // 3. 中文映射
            const zh = mapChineseChartType(val);
            if (zh) return zh;
            // 4. 语义映射（配置 chartSpec.semanticMap）："占比/构成"→饼图、"趋势/走势"→折线、"对比"→柱状
            const lower = val.toLowerCase();
            for (const m of defs().chartSpec.semanticMap || []) {
                if (m.re && m.re.test(lower)) return m.enum;
            }
        }
        return null;
    }

    /** 中文图表类型 → Excel 枚举
     * 说明（R1 收敛后）：这是"LLM 规格规范化路径"的独立且更完整的税onomy（含 箱线/树状/旭日/直方/气泡 +
     * 语义映射），与"检测侧" intent-detector.CHART_TYPE_DEFS/CHART_SPLIT_WORDS 有意分离：
     *   检测侧 折线→xlLine（快捷图表）；本处 折线→xlLineMarkers（折线带点）。请勿合并两者枚举，避免改行为。
     * 映射数据唯一事实源：Config/intent_config.json chartSpec.zhMap。
     */
    function mapChineseChartType(t) {
        const lower = String(t || '').toLowerCase();
        for (const m of defs().chartSpec.zhMap || []) {
            if (m.re && m.re.test(lower)) return m.enum;
        }
        return null;
    }

    /**
     * 枚举名 → 人类可读中文名（用于 UI 卡片显示，避免截图里出现 "Column Clustered" 这种英文）
     * 返回 { name: string, icon: string, category: string }
     * @param {string} enumName - xlPie / xlColumnClustered 等标准枚举
     * @returns {{name:string, icon:string, category:string}}
     */
    function humanizeChartType(enumName) {
        const clean = String(enumName || '').trim();
        const DEFAULT = { name: '图表', icon: '📊', category: '通用图表' };
        if (!clean) return DEFAULT;
        // 显示映射唯一事实源：Config/intent_config.json chartSpec.displayMap
        const displayMap = defs().chartSpec.displayMap || [];
        const byEnum = {};
        for (const d of displayMap) byEnum[d.enum] = d;
        // 先精确匹配
        if (byEnum[clean]) return byEnum[clean];
        // 宽松匹配（忽略大小写+空格+xl前缀）
        const norm = clean.replace(/^xl/i, '').replace(/\s+/g, '').toLowerCase();
        for (const d of displayMap) {
            if (d.enum.replace(/^xl/i, '').toLowerCase() === norm) return d;
        }
        // 关键词部分匹配（keywords 来自配置）
        const lower = clean.toLowerCase();
        for (const d of displayMap) {
            if ((d.keywords || []).some(kw => lower.includes(String(kw).toLowerCase()))) return d;
        }
        return DEFAULT;
    }

    /**
     * 根据用户意图/图表类型推断默认图表标题（避免显示死板的"(自动)"）
     * @param {string} userMsg - 用户原始消息（用于抽取语义关键词）
     * @param {string} chartType - xlPie / xlColumnClustered 枚举
     * @returns {string} 自然语言标题或空串（空串 = 让 Excel 自己默认）
     */
    function inferDefaultChartTitle(userMsg, chartType) {
        const msg = String(userMsg || '');
        const h = humanizeChartType(chartType);
        const cat = h.category;
        // 1) 从用户消息中抽取"展示/比较/分析/看"+关键词 结构
        const kwMatch = msg.match(/(?:展示|比较|对比|分析|看看|看下|看一下|画(?:一|个|张)?|做(?:一|个|张)?)\s*([\u4e00-\u9fa5A-Za-z0-9]{1,10})/);
        if (kwMatch && kwMatch[1]) {
            const kw = kwMatch[1].trim();
            if (kw.length >= 2) {
                // 用户说"占比/构成/趋势/对比"
                if (/占比|构成|比例|份额/i.test(kw)) return '占比分析';
                if (/趋势|走势|变化|时序/i.test(kw)) return '趋势分析';
                if (/对比|比较|差异/i.test(kw)) return '对比分析';
                if (/分布/i.test(kw)) return '分布分析';
                if (/转化|漏斗/i.test(kw)) return '转化漏斗';
                return kw;
            }
        }
        // 2) 按图表类别给通用默认标题
        if (cat === '占比构成') return '占比分析';
        if (cat === '对比分析') return '对比分析';
        if (cat === '趋势时序') return '趋势分析';
        if (cat === '累积构成') return '累积构成分析';
        if (cat === '相关性') return '相关性分析';
        if (cat === '多维评估') return '多维评估';
        if (cat === '层级构成') return '层级构成分析';
        if (cat === '分布分析') return '分布分析';
        if (cat === '转化分析') return '转化分析';
        if (cat === '增减分析') return '增减分析';
        return '';
    }

    /**
     * 规范化清洗 spec（白名单 + 默认合并）
     */
    function normalizeCleanSpec(raw) {
        const spec = { ...DEFAULT_CLEAN_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        if (Array.isArray(raw.operations)) {
            spec.operations = raw.operations.filter(op => CLEAN_OPERATIONS.has(op));
        }
        // 宽松：字符串操作（逗号分隔）
        else if (typeof raw.operations === 'string') {
            spec.operations = raw.operations.split(/[,，\s]+/).filter(op => CLEAN_OPERATIONS.has(op));
        }
        if (typeof raw.sheet === 'string' && raw.sheet.length <= 50) spec.sheet = raw.sheet;
        // range 容错：字符串 / 数字（Excel 行列索引）→ 统一字符串
        if (typeof raw.range === 'string' && raw.range.length <= 100) spec.range = raw.range;
        else if (typeof raw.range === 'number' && isFinite(raw.range)) spec.range = String(raw.range);
        if (['ffill', 'zero', 'delete', 'placeholder'].includes(raw.fillMethod)) spec.fillMethod = raw.fillMethod;
        // dedupeColumns 容错：数组（逐项过滤非字符串/空串+截断）/ 单值字符串（包装成数组）
        if (Array.isArray(raw.dedupeColumns)) {
            spec.dedupeColumns = raw.dedupeColumns.filter(c => typeof c === 'string' && c.trim() !== '').slice(0, 10);
        } else if (typeof raw.dedupeColumns === 'string' && raw.dedupeColumns.trim() !== '') {
            spec.dedupeColumns = [raw.dedupeColumns.trim()];
        }
        return spec;
    }

    /**
     * 清洗 spec 关键词兜底：LLM 规划失败/输出空 operations 时，
     * 从用户原话确定性生成清洗操作，保证小模型也能执行清洗。
     */
    function cleanSpecFromMessage(userMessage) {
        const spec = { ...DEFAULT_CLEAN_SPEC };
        const msg = String(userMessage || '').toLowerCase();
        const ops = [];
        if (/(去重|重复|重复行|去重合并|合并去重|重复数据|找重复|标重复)/.test(msg)) ops.push('removeDuplicates');
        if (/(空值|填充|补空|缺值|空白|fill)/.test(msg)) ops.push('fillEmpty');
        if (/(去空格|去除空格|清理空格|trim|删空格|空格)/.test(msg)) ops.push('trimSpaces');
        if (/(统一日期|日期格式|日期统一|日期规范|日期标准化|日期转换)/.test(msg)) ops.push('normalizeDates');
        if (/(统一数字|数字格式|金额|数值规范|数字统一|数字标准化|数值格式)/.test(msg)) ops.push('normalizeNumbers');
        if (/(文本规范|统一文本|规范化|标准文本|大小写|文本标准化)/.test(msg)) ops.push('standardizeText');
        if (/(空行)/.test(msg)) ops.push('removeEmptyRows');   // ⚠️ 词级判定：只要出现"空行"就确定是删空行。不能要求"清理空行/删除空行"连成串，否则"清理一下空行""把空行清理掉"会落空被错判成填充空值；"空值/空白"属填充、"空格"属去空格，均不在此列
        if (/(拆分列|拆列|拆分.*列)/.test(msg)) ops.push('splitColumn');
        if (/(合并列|并列|合并.*列)/.test(msg)) ops.push('mergeColumns');
        // —— 表结构变换（独立原子工具）——
        // ⚠️ splitByColumn（按字段值拆数据到多 Sheet）vs splitColumn（按分隔符拆列）：
        //    用户说"按XX拆到不同Sheet/按字段拆分数据"→ splitByColumn
        //    用户说"按逗号拆列/拆分B列"→ splitColumn（上面已处理）
        if (/(按.*拆.*(不同|多|各|分别|多个).*(表|sheet|Sheet)|拆.*(不同|多|各).*(表|sheet|Sheet)|按字段.*拆分|数据.*拆分.*表|拆分.*数据.*表)/.test(msg)) ops.push('splitByColumn');
        // 自动从"按 XX 拆分"里提取 dimColumn
        const dimMatch = msg.match(/按(.+?)拆(?:分到|分成|为|成)(?:不同的|多个|各)?(?:表|sheet|Sheet)/i);
        if (dimMatch && ops.includes('splitByColumn')) spec.dimColumn = (dimMatch[1] || '').trim();
        if (/(拆.*工作簿|工作簿.*拆|拆分.*xlsx|按.*行.*拆.*(表|sheet|Sheet))/.test(msg) && !ops.includes('splitByColumn')) {
            if (/(工作簿|xlsx|文件)/.test(msg)) ops.push('splitWorkbook');
            else { ops.push('splitByRows'); const rowM = msg.match(/(\d+)\s*行/); if (rowM) spec.splitRows = parseInt(rowM[1], 10); }
        }
        // ⚠️ mergeSheets（合并同一工作簿里的多张 Sheet）vs mergeWorkbooks（合并多个 xlsx 文件）
        if (/(合并.*(工作表|sheet|Sheet|多表|所有表|选中表|工作簿内)|合到.*(一张|一|单).*(表|sheet|Sheet)|合表|汇总表|合并数据).*/.test(msg)) ops.push('mergeSheets');
        if (/(合并.*(文件|xlsx|多个工作簿)|合并.*(目录|文件夹|批量).*(文件|xlsx|工作簿))/.test(msg)) ops.push('mergeWorkbooks');
        spec.operations = ops;
        return spec;
    }

    /** 图表规划 prompt */
    function buildChartPlanningPrompt(skill) {
        return `你是 SmartPilot 的图表规划师。你的唯一任务：把用户的自然语言需求转成【结构化 JSON spec】，绝不输出代码。

【用户需求】将从下方 user 消息提供。

【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块标记，不要解释），字段如下：
{
  "chartType": "xlPie",
  "title": null,                // 图表标题（用户说的标题）
  "dataRange": null,            // 数据范围如 "Sheet1!A1:B10"；用户没说则 null（用当前选区）
  "sheet": null,                // 图表放置的工作表；null=当前表
  "style": "business_blue",     // 主题样式（可选）
  "showLegend": true,
  "showDataLabels": false,
  "comboChart": false,          // 是否组合图（主柱+次折）
  "secondaryValueColumn": null  // 组合图次轴数据列
}

【chartType 映射规则 - 最重要，必须严格按此选】
- 用户说"饼图/占比/构成/比例/份额" → chartType: "xlPie"
- 用户说"柱状图/对比各部门/柱形" → "xlColumnClustered"
- 用户说"条形图/排名对比" → "xlBarClustered"
- 用户说"折线图/趋势/走势/随时间变化" → "xlLineMarkers"
- 用户说"面积图" → "xlArea"
- 用户说"环形图/圆环" → "xlDoughnut"
- 用户说"散点图/相关性" → "xlXYScatter"
- 用户说"漏斗图/转化" → "xlFunnel"
- 用户说"瀑布图/收支变化" → "xlWaterfall"
- 用户说"雷达图/多维评估" → "xlRadar"
- 用户说"直方图/分布/频率" → "xlHistogram"
- 用户说"树状图/层级" → "xlTreemap"

【规则】
- 用户说"画个饼图看占比" → chartType 必须是 "xlPie"，绝不能输出其他值
- 所有字段可省略
- 只输出 JSON，不要任何其他文字`;
    }

    /** 清洗规划 prompt */
    function buildCleanPlanningPrompt(skill) {
        return `你是 SmartPilot 的数据清洗与表结构变换规划师。把用户的自然语言需求转成【结构化 JSON spec】，绝不输出代码或代码块标记。

【用户需求】将从下方 user 消息提供。

【输出要求】只输出一个 JSON 对象，字段如下：
{
  "operations": ["removeDuplicates"],
  // ===== 数据清洗（走 excel_clean_data）=====
  //   removeDuplicates 去重 | fillEmpty 填充空值 | trimSpaces 去空格 |
  //   normalizeDates 日期统一 | normalizeNumbers 数字统一 | standardizeText 文本规范化 |
  //   removeEmptyRows 删空行 | splitColumn 按分隔符拆列 | mergeColumns 合并列（文本拼接）
  // ===== 表结构变换（独立原子工具，非 excel_clean_data）=====
  //   splitByColumn   按某列的值把数据拆到不同 Sheet（如"按部门拆分"）→ 必须填 dimColumn
  //   splitWorkbook   把每个工作表拆成独立 xlsx 文件
  //   splitByRows     按固定行数拆分成多个 Sheet → 必须填 splitRows
  //   mergeSheets     同一工作簿里多张表合并到一张（合并 Sheet，不是合并工作簿文件）
  //   mergeWorkbooks  多个 xlsx 文件合并到一个工作簿（需 filePaths）
  // ⚠️ 用户说"按 XX 拆分成不同 Sheet / 按字段拆分数据" → 选 splitByColumn，不是 splitColumn！
  // ⚠️ 用户说"合并工作表 / 合并多表" → 选 mergeSheets，不是 mergeWorkbooks！
  "sheet": null,
  "range": null,
  "fillMethod": "ffill",
  "dedupeColumns": null,
  // ====== splitByColumn 专用 ======
  "dimColumn": "",       // 按哪一列的唯一值拆分（如 "部门"）
  "mergeAll": false,     // true=强制合并全部可见表
  // ====== splitByRows 专用 ======
  "splitRows": 100,      // 每多少行拆一张
  // ====== splitWorkbook 专用 ======
  "outputDir": "",       // 拆分后的保存目录
  // ====== mergeWorkbooks 专用 ======
  "filePaths": [],       // 要合并的 xlsx 文件路径数组
  "targetPath": ""       // 合并后保存路径
}

【规则】
1. 选最匹配的 operation：如"去重"→removeDuplicates、"按部门拆到不同表"→splitByColumn + dimColumn="部门"、"合并多张工作表"→mergeSheets
2. 只输出 JSON，不要任何其他文字

🔒🔒🔒 结构变换强制规则（splitByColumn / splitWorkbook / splitByRows / mergeSheets / mergeWorkbooks）——违反会导致错误执行：
3. 结构变换 op **只能单独出现**（operations 数组里只有它一个），禁止和 removeDuplicates / fillEmpty / trimSpaces / normalize* 等清洗 op 混排，禁止和任何分析/聚合/图表 op 混排
4. 用户说"合并工作表/合并多表/合到一张表" → operations=["mergeSheets"]，**绝对禁止**脑补"先调 excel_aggregate 聚合 → 再调 excel_create_chart 画图"——合并是原子操作，一步到位
5. 用户说"按部门拆分不同 Sheet" → operations=["splitByColumn"], dimColumn="部门"，**绝对禁止**脑补聚合/画图——拆分是原子操作
6. 结构变换 op **不需要** fillMethod / dedupeColumns 这些清洗字段，保持它们为 null / 默认值即可`;
    }

    /**
     * 规范化增删改查（CRUD）spec（操作族白名单 + 默认合并）
     */
    function normalizeCrudSpec(raw) {
        const spec = { ...DEFAULT_CRUD_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        if (CRUD_OPERATIONS.has(String(raw.operation || '').toLowerCase())) {
            spec.operation = String(raw.operation).toLowerCase();
        }
        // 兼容英文 write/update、insert/add、delete/remove、clear/erase 多种说法归一
        const lowOp = String(raw.operation || '').toLowerCase();
        if (['update', 'set', 'edit', 'replace', 'fill', 'modify'].includes(lowOp)) spec.operation = 'write';
        else if (['add', 'append'].includes(lowOp)) spec.operation = 'insert';
        else if (['remove', 'erase'].includes(lowOp) && !/format|格式/.test(String(raw.range || ''))) spec.operation = 'delete';
        else if (/clear|cancel/i.test(lowOp) && /format|格式/.test(String(raw.range || ''))) spec.operation = 'clear';

        if (typeof raw.sheet === 'string' && raw.sheet.trim().length <= 50) spec.sheet = raw.sheet.trim();
        if (typeof raw.range === 'string' && raw.range.trim().length <= 100) spec.range = raw.range.trim();
        // value：字符串 / 数字 / 布尔 → 统一字符串（数字保留原样写法）
        if (raw.value !== undefined && raw.value !== null) {
            if (typeof raw.value === 'string') spec.value = raw.value;
            else if (typeof raw.value === 'number' || typeof raw.value === 'boolean') spec.value = String(raw.value);
        }
        // values：二维数组（逐项容错为非数组→单元素行）+ 长度上限
        if (Array.isArray(raw.values)) {
            spec.values = raw.values.slice(0, 1000).map(row =>
                (Array.isArray(row) ? row : [row]).slice(0, 500).map(v =>
                    (v === null || v === undefined) ? null : String(v)));
        } else if (Array.isArray(raw.container) && Array.isArray(raw.container[0])) {
            spec.values = raw.container.slice(0, 1000).map(row => row.map(v => (v == null ? null : String(v))));
        }
        // row / column（1 基）
        if (typeof raw.row === 'number' && Number.isInteger(raw.row) && raw.row > 0) spec.row = raw.row;
        else if (typeof raw.row === 'string' && /^\d+$/.test(raw.row.trim())) spec.row = parseInt(raw.row, 10);
        if (typeof raw.column === 'number' && Number.isInteger(raw.column) && raw.column > 0) spec.column = raw.column;
        else if (typeof raw.column === 'string' && /^\d+$/.test(raw.column.trim())) spec.column = parseInt(raw.column, 10);
        // count（1~500）
        if (typeof raw.count === 'number' && Number.isInteger(raw.count) && raw.count >= 1) spec.count = Math.min(raw.count, 500);
        else if (typeof raw.count === 'string' && /^\d+$/.test(raw.count.trim())) spec.count = Math.min(parseInt(raw.count, 10), 500);
        if (typeof raw.append === 'boolean') spec.append = raw.append;
        if (typeof raw.clearContent === 'boolean') spec.clearContent = raw.clearContent;
        if (typeof raw.clearFormat === 'boolean') spec.clearFormat = raw.clearFormat;
        return spec;
    }

    /** CRUD 操作默认值（供 normalizeCrudSpec 兜底，与 DEFAULT_CRUD_SPEC 同源） */
    function defaultCrudSpec() { return { ...DEFAULT_CRUD_SPEC }; }

    /**
     * CRUD spec 关键词兜底：LLM 规划失败/输出非法 operation 时，
     * 从用户原话确定性推导操作族与行列锚点，保证小模型也能执行增删改查。
     */
    function crudSpecFromMessage(userMessage) {
        const spec = defaultCrudSpec();
        const msg = String(userMessage || '').replace(/\s+/g, '');
        const lower = msg.toLowerCase();
        // 多步骤/聚合/图表请求不属于单一 CRUD 操作，交回多步骤管道统一规划。
        // 用 needPlanning 类别计数 + detectBuiltinIntent 判定（而非关键词枚举），
        // 避免"换个词就失效"——例如"读取销售数据，按部门分类汇总，然后创建柱状图"
        // 含"读取+汇总+建图"三类操作，若被 CRUD 拦截只会执行"读取"而丢失后续步骤。
        const chatV2 = (typeof window !== 'undefined' && window.App && window.App.ChatV2) ? window.App.ChatV2 : null;
        if (chatV2) {
            const TP = chatV2.TaskPlanner;
            if (TP && typeof TP.needPlanning === 'function' && TP.needPlanning(msg) === true) return null;
            const planner = chatV2.taskPlanner;
            if (planner && typeof planner.detectBuiltinIntent === 'function' && planner.detectBuiltinIntent(msg).builtin) return null;
        }
        // 排除明确属于其它 Skill 的意图，避免误入 CRUD（这些词出现在消息中时交回关键词匹配路由）
        // 含"工作表/工作簿/重命名/冻结/保护"等表管理类意图 → 走 excel_*_sheet 工作表管理工具，不归 CRUD
        if (/(去重|重复|清洗|美化|图表|分析|统计|报告|合并|拆分|排序|筛选|工作表|工作簿|重命名|冻结|保护)/.test(msg)) return null;

        // 行列锚点提取（含中文数字）
        let row = null, col = null, count = 1;
        // 中文数字映射
        const cnNum = (s) => {
            if (s == null) return null;
            const map = { '一': 1, '两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9, '十': 10 };
            if (/^十[一二三四五六七八九]$/.test(s)) return 10 + (map[s[1]] || 0);
            if (s === '十') return 10;
            if (/^[一二三四五六七八九]十$/.test(s)) return map[s[0]] * 10;
            if (/^[一二三四五六七八九]十[一二三四五六七八九]$/.test(s)) return map[s[0]] * 10 + map[s[2]];
            return map[s] || null;
        };
        // "第X行" / "删除第5行" → row
        const rowMatch = msg.match(/第\s*([一二三四五六七八九十两0-9]+)\s*行/);
        if (rowMatch) row = /^\d+$/.test(rowMatch[1]) ? parseInt(rowMatch[1], 10) : cnNum(rowMatch[1]);
        // "第X列"/"C列"/"第3列" → col
        const colCn = msg.match(/第\s*([一二三四五六七八九十两0-9]+)\s*列/);
        if (colCn) col = /^\d+$/.test(colCn[1]) ? parseInt(colCn[1], 10) : cnNum(colCn[1]);
        // 字母列，"C列" / "D:E"
        const colLetter = msg.match(/([A-Za-z]{1,2})\s*列/);
        if (colLetter && col == null) {
            const letters = colLetter[1].toUpperCase(); let c = 0;
            for (let i = 0; i < letters.length; i++) c = c * 26 + (letters.charCodeAt(i) - 64);
            col = c;
        }
        // 显式范围 "A1:B10" / "A1" / "A:A" / "C:F" / "5:8"（读/写/清空场景，用户消息明示范围时优先生效）
        let rangeExtract = null;
        const rangeMatch = msg.match(/\b[A-Za-z]{1,3}\d+\s*:\s*[A-Za-z]{1,3}\d+\b/) ||  // A1:B10
                           msg.match(/\b[A-Za-z]:[A-Za-z]\b/) ||                       // A:A / C:F
                           msg.match(/\b\d+:\d+\b/);                                   // 5:8
        if (rangeMatch) rangeExtract = rangeMatch[0].replace(/\s+/g, '');
        // 单格地址 "A1"（插入/写入场景）
        const cellAddr = msg.match(/\b([A-Za-z]{1,2})(\d{1,6})\b/);
        // 插入/删除数量："插入两行"→2。需跳过"第X行/第X列"锚点（其数字是锚点而非数量）
        const cntCandidates = msg.matchAll(/([一二两三四五六七八九十0-9]+)\s*(?:行|列)/g);
        for (const m of cntCandidates) {
            const prev = msg.slice(0, m.index);
            // 锚点是"第N行/第N列/在第3行"：最后一个字符是"第/在/于"→跳过
            if (/[第在于]$/.test(prev)) continue;
            const c = /^\d+$/.test(m[1]) ? parseInt(m[1], 10) : cnNum(m[1]);
            if (c && c > count && count === 1) { count = c; }
        }
        if (count < 1) count = 1;

        // 操作族判定
        spec.row = row; spec.column = col; spec.count = count;
        if (/(读取|读一下|读数据|查一下|查看|查询|读出|返回|列出|看看.*数据|读取.*数据|\bread\b)/.test(lower)) {
            spec.operation = 'read';
            if (rangeExtract) spec.range = rangeExtract;
        } else if (/(插入|新增行|新增列|插|添加行|添加列|追加|下面添加|加一行|加一列|insert|\badd\b)/.test(msg)) {
            spec.operation = 'insert';
            // 追加只在明确"末尾/追加"时置位
            if (/(末尾|最后|追加|底部|尾部)/.test(msg)) spec.append = true;
            if (col != null && row == null) spec.column = col;
        } else if (/(删除|删掉|移除|删去|删除行|删除列|清除行|delete|remove)/.test(lower)) {
            spec.operation = 'delete';
        } else if (/(清空|清除|清掉|清理|清备注|抹掉|clear)/.test(lower)) {
            spec.operation = 'clear';
            if (rangeExtract) spec.range = rangeExtract;
            // 仅清格式 vs 清内容：明确出现"格式"且无"内容/数据/值"时判定为仅清格式
            if (/(格式)/.test(msg) && !/(内容|数据|值|单元格)/.test(msg)) {
                spec.clearContent = false; spec.clearFormat = true;
            } else if (/(内容|数据|值)/.test(msg)) {
                spec.clearContent = true; spec.clearFormat = false;
            }
        } else if (/(写入|填写|填|改为|改成|设置|设为|修改|替换|换成|更新|填上|write|set|update|replace)/.test(lower) || cellAddr) {
            spec.operation = 'write';
            if (cellAddr && spec.row == null && spec.column == null) {
                // "把A1改成100"：A1 → row=1, column=1，供后端单格写
                const L = cellAddr[1].toUpperCase(); let c = 0;
                for (let i = 0; i < L.length; i++) c = c * 26 + (L.charCodeAt(i) - 64);
                spec.column = c; spec.row = parseInt(cellAddr[2], 10);
                spec.range = cellAddr[0].toUpperCase();
                spec.value = extractWriteValue(userMessage);
            } else if (rangeExtract) {
                spec.range = rangeExtract;
            }
        } else {
            return null; // 无法判定操作族 → 交回通用对话
        }
        return spec;
    }

    /** 从用户消息中提取写入目标值（"把A1改成100"→"100"） */
    function extractWriteValue(userMessage) {
        const msg = String(userMessage || '').trim();
        for (const m of msg.matchAll(/(?:改成|改为|设置为|设为|调成|变成|替换成|更新为|更新成|填入|写入|填成)\s*"?(?:[^\s"，。；]+|(?:"[^"]*"))/g)) {
            const v = m[0].replace(/^(?:改成|改为|设置为|设为|调成|变成|替换成|更新为|更新成|填入|写入|填成)\s*"?/, '').replace(/"$/, '').trim();
            if (v) return v;
        }
        // 兜底："= TOKEN" 后面的值（排除表名/保留字）
        const eq = msg.match(/[=＝]\s*"?([^"，。；\s"=]+)"?$/);
        if (eq) return eq[1].replace(/"/g, '').trim();
        return null;
    }

    /** CRUD 规划 prompt */
    function buildCrudPlanningPrompt(skill) {
        return `你是 SmartPilot 的 Excel 增删改查规划师。你的唯一任务：把用户的自然语言需求转成【结构化 JSON spec】，绝不输出代码、绝不调用任何函数。

【用户需求】将从下方 user 消息提供。

【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块标记，不要解释），字段如下：
{
  "operation": "write",        // 操作族：read(读) | write(写/改) | insert(插入) | delete(删除) | clear(清空)
  "sheet": null,               // 工作表；null=当前活动表
  "range": null,               // 目标区域，如 "A1:B10"；单格 "A1"；行 "5:8"；列 "C:C"
  "value": null,               // write 单值：如改"张三"→"张三"
  "values": null,              // write 批量：二维数组 [["a","b"],["c","d"]]
  "row": null,                 // insert/delete 行锚点（1 基），如第3行→3
  "column": null,              // insert/delete 列锚点（1 基），如C列→3
  "count": 1,                  // insert/delete 数量，插入两行→2
  "append": false,             // insert：是否追加到表尾（用户说"末尾/追加"时 true）
  "clearContent": true,        // clear：是否清内容
  "clearFormat": true          // clear：是否清格式
}

【规则】
- operation 判定示例：看/读/查询→read；改成/写入/填写→write；插入/添加行列→insert；删除/删掉行列→delete；清空/清除内容或格式→clear
- insert/delete 若用户给出位置（第X行/第Y列/X列/A3）填入 row 或 column；追加到末尾时 append=true
- write 单格用 value，批量用 values；value/values 至少给一个
- clear 若明确"只清格式"→clearContent:false, clearFormat:true；明确"只清内容"→clearContent:true, clearFormat:false
- 所有字段可省略；拿不准的操作族给最接近的一个
- 只输出 JSON，不要任何其他文字`;
    }

    // ============================================================
    // 看板蓝图（DashboardGenerator）
    // ------------------------------------------------------------
    // 设计：LLM 只输出高层蓝图（标题/主题/KPI/图表/数据源），蓝图引擎把它翻译为
    //       有序的原子工具调用序列（Action 计划），双路生成 Excel 内看板 + HTML 大屏。
    // 蓝图是独立产物：可持久化、可 diff 修正（后续消息"换个主题/加个KPI"基于上轮蓝图调整）。
    // ============================================================

    /** 看板主题白名单（与 dashboard-blueprint-engine 的调色板对齐） */
    const DASHBOARD_THEMES = new Set(['business_blue', 'dark', 'fresh_green', 'warm_orange']);

    /** 看板交付物白名单：dual(双路,默认) | excel(仅 Excel 内工作表) | html(仅 HTML 大屏) */
    const DASHBOARD_DELIVERABLES = new Set(['dual', 'excel', 'html']);

    /** 看板蓝图默认值 */
    const DEFAULT_DASHBOARD_SPEC = {
        type: 'dashboard',
        title: '数据看板',
        subtitle: '',
        theme: 'business_blue',     // business_blue | dark | fresh_green | warm_orange
        deliverable: 'dual',        // dual(默认) | excel(仅 Excel 工作表) | html(仅 HTML 大屏)
        sheet: null,                // Excel 看板工作表名（null → 默认 "看板"）
        dataSources: [],            // [{ id, sheet?, range?, dimension?, value?, func? }]
        kpis: [],                   // [{ label, value?, unit?, source?, column?, func? }]
        charts: [],                 // [{ title?, source?, chartType, width?, height? }]
        slicers: [],                // Excel 原生切片器维度列名（最多 4 个，联动透视表与图表）
        table: null                 // { source?, maxRows? } 可选明细表
    };

    /** 去掉 LLM 原样照抄的 <占位符>（整体被尖括号包裹的演示值视为无效，返回空串） */
    function _cleanPlaceholder(v) {
        if (typeof v !== 'string') return '';
        const s = v.trim();
        return /^<.*>$/.test(s) ? '' : s;
    }

    /** 归一化单个数据源声明 */
    function normalizeDashboardDataSource(raw, index) {
        if (!raw || typeof raw !== 'object') return null;
        const ds = {};
        ds.id = (typeof raw.id === 'string' && raw.id.trim()) ? raw.id.trim() : ('src_' + index);
        const sheet = _cleanPlaceholder(raw.sheet);
        const range = _cleanPlaceholder(raw.range);
        const dimension = _cleanPlaceholder(raw.dimension);
        const value = _cleanPlaceholder(raw.value);
        if (sheet && sheet.length <= 50) ds.sheet = sheet;
        if (range && range.length <= 100) ds.range = range;
        if (dimension && dimension.length <= 50) ds.dimension = dimension;
        if (value && value.length <= 50) ds.value = value;
        if (typeof raw.func === 'string' && /^(sum|avg|average|count|max|min)$/i.test(raw.func.trim())) {
            ds.func = raw.func.trim().toLowerCase();
            if (ds.func === 'average') ds.func = 'avg';
        }
        return ds;
    }

    /** 归一化单个 KPI 卡声明 */
    function normalizeDashboardKpi(raw, index) {
        if (!raw || typeof raw !== 'object') return null;
        const kpi = {};
        const label = _cleanPlaceholder(raw.label);
        const labelIsPlaceholder = !label || /^<.+>$/.test(label) || /^指标\s*\d*$/.test(String(label).trim());
        if (label && !labelIsPlaceholder) kpi.label = label;
        else {
            // 无标题或标题是"指标 N"类占位 → 从计算方式自动生成有意义的指标名（LLM 常只给 column/func 不写 label）
            const col = _cleanPlaceholder(raw.column);
            const f = typeof raw.func === 'string' ? raw.func.trim().toLowerCase() : '';
            if (f === 'count') kpi.label = '记录数';
            else if (f === 'avg' || f === 'average') kpi.label = col ? '平均' + col : '平均值';
            else if (f === 'max') kpi.label = col ? '最大' + col : '最大值';
            else if (f === 'min') kpi.label = col ? '最小' + col : '最小值';
            else if (col) kpi.label = col;
            else if (raw.value !== undefined && raw.value !== null && raw.value !== '') kpi.label = '指标 ' + (index + 1);
            else kpi.label = '合计';
        }
        if (raw.value !== undefined && raw.value !== null && raw.value !== '') kpi.value = raw.value;
        const unit = _cleanPlaceholder(raw.unit);
        if (unit && unit.length <= 20) kpi.unit = unit;
        if (typeof raw.trend === 'string' && raw.trend.trim().length <= 20) kpi.trend = raw.trend.trim();
        const source = _cleanPlaceholder(raw.source);
        if (source) kpi.source = source;
        const column = _cleanPlaceholder(raw.column);
        if (column && column.length <= 50) kpi.column = column;
        if (typeof raw.func === 'string' && /^(sum|avg|average|count|max|min)$/i.test(raw.func.trim())) {
            kpi.func = raw.func.trim().toLowerCase();
            if (kpi.func === 'average') kpi.func = 'avg';
        }
        return kpi;
    }

    /** 归一化单个图表卡声明 */
    function normalizeDashboardChart(raw, index) {
        if (!raw || typeof raw !== 'object') return null;
        const chart = {};
        const title = _cleanPlaceholder(raw.title);
        if (title && title.length <= 60) chart.title = title;
        else chart.title = '图表 ' + (index + 1);
        chart.chartType = normalizeChartTypeValue(raw.chartType) || 'xlColumnClustered';
        const source = _cleanPlaceholder(raw.source);
        if (source) chart.source = source;
        const sheet = _cleanPlaceholder(raw.sheet);
        if (sheet && sheet.length <= 50) chart.sheet = sheet;
        const range = _cleanPlaceholder(raw.range);
        if (range && range.length <= 100) chart.range = range;
        const dimension = _cleanPlaceholder(raw.dimension || raw.dim);
        if (dimension && dimension.length <= 50) chart.dimension = dimension;
        const value = _cleanPlaceholder(raw.value
            || (Array.isArray(raw.vals) ? raw.vals[0] : '')
            || (Array.isArray(raw.values) ? raw.values[0] : ''));
        if (value && value.length <= 50) chart.value = value;
        if (typeof raw.func === 'string' && /^(sum|avg|average|count|max|min)$/i.test(raw.func.trim())) {
            chart.func = raw.func.trim().toLowerCase();
            if (chart.func === 'average') chart.func = 'avg';
        }
        const w = Number(raw.width), h = Number(raw.height);
        if (Number.isFinite(w) && w >= 1 && w <= 4) chart.width = Math.round(w);
        if (Number.isFinite(h) && h >= 1 && h <= 4) chart.height = Math.round(h);
        return chart;
    }

    /**
     * 规范化看板蓝图 spec（白名单 + 默认合并）
     * @param {Object} raw - LLM 输出的蓝图 JSON
     * @returns {Object} 规范化后的蓝图
     */
    function normalizeDashboardSpec(raw) {
        const spec = { ...DEFAULT_DASHBOARD_SPEC };
        if (!raw || typeof raw !== 'object') return spec;
        const _title = _cleanPlaceholder(raw.title);
        if (_title && _title.length <= 60) spec.title = _title;
        const _subtitle = _cleanPlaceholder(raw.subtitle);
        if (_subtitle && _subtitle.length <= 120) spec.subtitle = _subtitle;
        if (DASHBOARD_THEMES.has(String(raw.theme || '').toLowerCase())) spec.theme = String(raw.theme).toLowerCase();
        if (DASHBOARD_DELIVERABLES.has(String(raw.deliverable || '').toLowerCase())) spec.deliverable = String(raw.deliverable).toLowerCase();
        const _sheet = _cleanPlaceholder(raw.sheet);
        if (_sheet && _sheet.length <= 30) spec.sheet = _sheet;
        if (Array.isArray(raw.dataSources)) {
            spec.dataSources = raw.dataSources.slice(0, 12)
                .map(normalizeDashboardDataSource)
                .filter(Boolean);
        }
        if (Array.isArray(raw.kpis)) {
            spec.kpis = raw.kpis.slice(0, 12).map(normalizeDashboardKpi).filter(Boolean);
        }
        if (Array.isArray(raw.charts)) {
            spec.charts = raw.charts.slice(0, 10).map(normalizeDashboardChart).filter(Boolean);
        }
        if (Array.isArray(raw.slicers)) {
            // 原生切片器维度：列名白名单（1~50 字符，最多 4 个）
            spec.slicers = raw.slicers.slice(0, 4)
                .map(s => typeof s === 'string' ? _cleanPlaceholder(s.trim()) : '')
                .filter(s => s && s.length >= 1 && s.length <= 50);
        }
        if (raw.table && typeof raw.table === 'object') {
            const t = {};
            if (typeof raw.table.source === 'string' && raw.table.source.trim()) t.source = raw.table.source.trim();
            if (typeof raw.table.sheet === 'string' && raw.table.sheet.trim()) t.sheet = raw.table.sheet.trim();
            if (typeof raw.table.range === 'string' && raw.table.range.trim()) t.range = raw.table.range.trim();
            if (typeof raw.table.maxRows === 'number' && raw.table.maxRows >= 1) t.maxRows = Math.min(Math.round(raw.table.maxRows), 200);
            spec.table = t;
        }
        return spec;
    }

    /** 看板规划 prompt：LLM 只规划蓝图，不写代码、不算坐标 */
    function buildDashboardPlanningPrompt(skill) {
        return `你是 SmartPilot 的看板（Dashboard）规划师。你的唯一任务：把用户的自然语言需求转成【结构化看板蓝图 JSON】，绝不输出代码，绝不计算坐标。

【用户需求】将从下方 user 消息提供（可能附带当前选区的列名/样本数据，供你规划 KPI 与图表）。

【输出要求】只输出一个 JSON 对象（不要 Markdown 代码块标记，不要解释），字段如下
（注意：<尖括号> 是占位符，输出时必须替换成真实值；没有真实值就整字段省略，禁止原样输出 <...>）：
{
  "type": "dashboard",
  "title": "<按用户需求或数据内容概括的看板名>",   // 看板标题（1~30 字），严禁使用本示例中的任何标题字样
  "subtitle": "<可选副标题>",                              // 可省略
  "theme": "business_blue",           // 主题：business_blue(商务蓝) | dark(深色大屏) | fresh_green(清新绿) | warm_orange(暖橙)
  "deliverable": "<excel|html|dual>", // 🔴 必须填！按下方"首要决策"规则自主选择，禁止无脑写 "dual"
  "sheet": "<可选，Excel 看板工作表名，省略默认“看板”>",
  "dataSources": [                    // 数据源声明（图表/KPI 引用它；可选，缺省时引擎按选区自动派生）
    { "id": "src_1", "sheet": "<实际工作表名，可省略>", "dimension": "<实际分类列名>", "value": "<实际数值列名>", "func": "sum" }
  ],
  "kpis": [                           // KPI 指标卡（建议 3~5 个，最多 8 个）
    { "label": "<总额类指标名>", "source": "src_1", "column": "<实际数值列名>", "func": "sum", "unit": "<可选单位>" },
    { "label": "<均值类指标名>", "source": "src_1", "column": "<实际数值列名>", "func": "avg" },
    { "label": "<计数类指标名>", "source": "src_1", "func": "count" }
  ],
  "charts": [                         // 图表卡：覆盖每个有分析价值的分类维度，通常 2~4 张（最多 6 张）
    { "title": "<按维度A的图表标题>", "source": "src_1", "dimension": "<实际分类列名A>", "value": "<实际数值列名或计数列>", "func": "sum", "chartType": "xlColumnClustered" },
    { "title": "<按维度B的图表标题>", "source": "src_1", "dimension": "<实际分类列名B>", "value": "<实际数值列名或计数列>", "func": "sum", "chartType": "xlColumnClustered" }
  ],
  "slicers": [ "<分类列名A>", "<分类列名B>" ],  // 可选：Excel 原生切片器维度（实际分类列名，1~4 个）；省略时引擎自动取图表维度前 3 个
  "table": { "source": "src_1", "maxRows": 20 }   // 可选：明细表
}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  🔴 首要决策：deliverable（交付物类型）—— 必须先判断再填 JSON
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
根据用户消息里的关键词，选择 deliverable（excel/html/dual 三选一）：
✅ 选 "excel"：用户提到 Excel、表格、工作表、sheet、表内、单元格、"只要 Excel 里"、"不要大屏/不要网页"
   例："生成Excel看板"、"做个经营仪表盘放工作表里"、"用表格数据做看板"
✅ 选 "html"：用户只提 HTML、网页、浏览器、投屏、大屏、"不要 Excel"
   例："HTML大屏"、"做个网页版驾驶舱"
✅ 选 "dual"：同时提到 Excel + HTML/投屏 需求，或用户没说、模糊的"看板/仪表盘/驾驶舱"
⚠️  禁止无脑写 "dual"！先读用户消息再决定。

【数据源说明】
- dimension：分组维度列（表头名）；value：数值列（表头名）；func：sum/avg/count/max/min
- 图表/KPI 通过 "source" 引用 dataSources 里的 id；直接引用时也可在图表里内联 sheet/range/dimension/value
- 若数据已是明细（如"直接用选区做看板"），可省略 dataSources，引擎会自动把数值列做成 KPI、分类列×数值列做成图表

【chartType 映射规则】
- 占比/构成 → "xlPie"；对比大小 → "xlColumnClustered"；排名 → "xlBarClustered"
- 趋势/走势/随时间 → "xlLineMarkers"；面积 → "xlArea"；环形 → "xlDoughnut"
- 散点/相关 → "xlXYScatter"；漏斗/转化 → "xlFunnel"；瀑布/收支 → "xlWaterfall"

【规则】
- 严禁照抄示例：标题、id、工作表名、列名、指标名必须来自用户需求或当前选区的真实列名/样本数据；示例值仅演示字段格式
- 用户没给业务名时，title 按数据内容概括命名（如数据是商品类别订单，就叫“商品类别分析看板”），不要套用示例标题
- 拿不准数据源或列名时直接省略 dataSources / kpis.column / charts.dimension、value，引擎会按选区自动派生，禁止编造列名
- KPI 值不要手写数字：一律不输出 value 字段，交给引擎按 func + 数据计算（防占位 0 污染看板）
- KPI 建议 3~5 个（总额/均值/记录数/主要维度数），至少 2 个；拿不准的图表类型用 xlColumnClustered
- 图表要覆盖数据中每个有分析价值的分类维度（如销售区域、商品类别、业务员各一张），通常 2~4 张，不要只出一张图；省略 charts 时引擎会自动枚举全部合格维度生成
- 所有字段可省略
- 只输出 JSON，不要任何其他文字`;
    }

    /** 关键词兜底：LLM 规划失败时构造最小可用蓝图（引擎按选区自动派生 KPI/图表） */
    function dashboardSpecFromMessage(msg) {
        const text = String(msg || '');
        const spec = { ...DEFAULT_DASHBOARD_SPEC };
        const titleMatch = text.match(/(?:生成|做|建|创建|搭|搞)?(?:一个|张|块)?([\u4e00-\u9fa5A-Za-z0-9]{1,12})?(?:看板|仪表盘|驾驶舱|大屏)/);
        if (titleMatch && titleMatch[1] && /看板|仪表盘|驾驶舱|大屏/.test(titleMatch[1]) === false) {
            spec.title = titleMatch[1] + '看板';
        }
        if (/深色|dark|暗色|夜间/.test(text)) spec.theme = 'dark';
        else if (/清新|绿|green/.test(text)) spec.theme = 'fresh_green';
        else if (/暖橙|橙色|orange/.test(text)) spec.theme = 'warm_orange';
        // ── 交付物意图识别（与 Skill MD / 规划 prompt 对齐）──
        // 明确只要网页/投屏 → html；明确 Excel/工作表看板且未提网页 → excel；其余 dual。
        if (/(网页|浏览器|html|web|投屏|大屏就够|只要大屏|仅大屏|单独大屏|纯大屏)/i.test(text)
            && !/(excel|工作表|表格看板|excel\s*看板)/i.test(text)) {
            spec.deliverable = 'html';
        } else if (
            /(只要|仅|就|只)(excel|表里|工作表里|表格里|放到excel|不要大屏|不要网页|别做大屏|别开网页)/i.test(text)
            || /(不?开网页|不?打开浏览器|不要大屏|不要网页)/.test(text)
            || /(excel\s*(内)?\s*(看板|仪表盘|驾驶舱|大屏)|表格看板|工作表(里的)?看板|放到(工作表|excel)|excel\s*里)/i.test(text)
        ) {
            spec.deliverable = 'excel';
        }
        const sheetMatch = text.match(/(?:放在|写到|写入|建在|命名叫?|表名)\s*["'“”]?([\u4e00-\u9fa5A-Za-z0-9_]{1,20})["'”]?/);
        if (sheetMatch && sheetMatch[1]) spec.sheet = sheetMatch[1];
        return spec;
    }

    App.ChatV2.services.skillSpecEngine = {
        planAndRender,
        planFormat,
        specFromMessage,
        cleanSpecFromMessage,
        normalizeSpec,
        normalizeFormatSpec,
        normalizeChartSpec,
        normalizeCleanSpec,
        normalizeCrudSpec,
        normalizeDashboardSpec,
        normalizeDashboardKpi,
        normalizeDashboardChart,
        dashboardSpecFromMessage,
        buildDashboardPlanningPrompt,
        defaultCrudSpec,
        crudSpecFromMessage,
        extractWriteValue,
        renderBeautifyVBA,
        renderFormatVBA,
        buildPlanningPrompt,
        buildFormatPlanningPrompt,
        buildChartPlanningPrompt,
        buildCleanPlanningPrompt,
        buildCrudPlanningPrompt,
        normalizeChartTypeValue,
        mapChineseChartType,
        humanizeChartType,
        inferDefaultChartTitle,
        CHART_TYPES,
        CLEAN_OPERATIONS,
        CRUD_OPERATIONS,
        DASHBOARD_THEMES,
        DEFAULT_CHART_SPEC,
        DEFAULT_CLEAN_SPEC,
        DEFAULT_CRUD_SPEC,
        DEFAULT_DASHBOARD_SPEC,
        resolveColor,
        extractColorWords,
        COLOR_WORDS,
        THEMES,
        DEFAULT_SPEC,
        DEFAULT_FORMAT_SPEC,
        listThemes: () => Object.entries(THEMES).map(([id, t]) => ({ id, name: t.name })),
        // 样式方案（版式 × 配色 × 密度）：供 UI/提示词枚举；预设模块未加载时返回空数组
        listStylePresets: () => {
            const P = presetsModule();
            return P ? P.listLayouts() : [];
        },
        listStylePalettes: () => {
            const P = presetsModule();
            return P ? P.listPalettes() : [];
        }
    };
    ChatLog.info?.('CHAT-V2', '✅ skillSpecEngine loaded (' + Object.keys(THEMES).length + ' themes)');
})();