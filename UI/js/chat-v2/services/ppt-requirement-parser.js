/**
 * ============================================================
 * chat-v2/services/ppt-requirement-parser.js — PPT 用户需求解析（规则层）
 * ============================================================
 * 职责：
 *   - 从用户消息中规则化提取结构化需求（图表类型 / 强调重点 / 页面类型 / 章节）
 *   - 把解析出的需求应用到推荐 slides（applyRequirementsToDesign）
 *   - 判断需求是否复杂到需要降级 LLM（needsLlm）
 *
 * 设计原则：确定性规则，不依赖 LLM，符合"少报错"哲学。
 * 复杂需求（叙事结构 / 自定义内容）由调用方（code-pipeline）降级到 LLM 生成规格。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.services) App.ChatV2.services = {};

    const ChatLog = window.ChatLog || console;

    // ------------------------------------------------------------
    // 规则表
    // ------------------------------------------------------------
    const CHART_TYPE_RULES = [
        { type: 'line', keywords: ['折线', '趋势图', '走势', '变化曲线', '增长曲线', 'line'] },
        { type: 'column', keywords: ['柱状', '柱形', '条形', '对比图', '柱图', 'bar', 'column'] },
        { type: 'pie', keywords: ['饼图', '占比', '构成', '份额', 'pie'] }
    ];

    // 图表形式：图片（稳定不可编辑）vs 带数据的原生图（可编辑）；不命中默认 auto 自动回退
    const CHART_MODE_RULES = [
        { mode: 'picture', keywords: ['用图片', '图片图表', '图片形式', '静态图', '贴图', '图表用图片', '用图不要数据'] },
        { mode: 'native', keywords: ['带数据的图', '带数据', '可编辑', '原生图表', '能编辑', '可修改', '图表带数据'] }
    ];

    // 聚合方式（用户驱动模型可显式指定）：求和 / 均值 / 计数；不命中默认交给数据驱动判别
    const AGG_RULES = [
        { agg: 'sum', label: '按列求和', keywords: ['求和', '合计', '加总', '汇总', '累加', '总计', 'sum'] },
        { agg: 'avg', label: '取均值', keywords: ['平均', '均值', '取平均', '算平均', 'avg', 'mean'] },
        { agg: 'count', label: '计数', keywords: ['计数', '数量统计', '条数', 'occurrences', 'count'] }
    ];

    // 版式比例：16:9（宽屏，默认）vs 4:3（标准屏）；竖版经 orientation=portrait 处理（此处不兜图表）
    const RATIO_RULES = [
        { ratio: '16x9', keywords: ['16:9', '16比9', '宽屏', '宽银幕', 'widescreen', '16x9', '169'] },
        { ratio: '4x3', keywords: ['4:3', '4比3', '标准屏', '4x3', '43'] }
    ];

    const RATIO_ORIENTATION_RULES = [
        { orientation: 'portrait', keywords: ['竖版', '纵向', '竖屏', '竖排', 'portrait'] }
    ];

    const PAGE_RULES = [
        { type: 'comparison', keywords: ['对比', '比较', '竞品', '对照', '方案比'] },
        { type: 'timeline', keywords: ['时间线', '时间轴', '里程碑', '发展历程', '演进', '大事记'] },
        { type: 'table', keywords: ['数据明细', '明细表', '完整数据', '原始数据', '表格'] },
        { type: 'kpi', keywords: ['关键指标', 'KPI', '核心指标', '指标总览', '指标卡'] },
        { type: 'text', keywords: ['总结', '建议', '结论', '要点', '概述', '汇报'] }
    ];

    const EMPHASIS_PATTERNS = [
        /重点(?:突出|关注|强调)?[：:]?\s*["“《「]?([^"”》」,，。；;、\s]+)["”》」]?/,
        /突出\s*["“《「]?([^"”》」,，。；;、\s]+)["”》」]?/,
        /强调\s*["“《「]?([^"”》」,，。；;、\s]+)["”》」]?/,
        /重点关注\s*["“《「]?([^"”》」,，。；;、\s]+)["”》」]?/
    ];

    const SECTION_PATTERNS = [
        /(?:章节|部分|板块)[：:]?\s*["“《「]?([^"”》」,，。；;]+)["”》」]?/,
        /(?:包括|包含)\s*["“《「]?([^"”》」,，。；;]+)["”》」]?/
    ];

    const NARRATIVE_PATTERNS = [
        /(先|首先).{0,20}(再|然后).{0,20}(最后|最后)/,
        /第[一二三四五六七八九十百]+部分/,
        /(背景|引言).{0,20}(问题|现状).{0,20}(建议|总结|展望)/,
        /(结构|顺序|流程)[：:]?\s*(先|首先)/
    ];

    // ------------------------------------------------------------
    // 解析：用户消息 → 结构化需求
    // ------------------------------------------------------------
    function parsePptRequirements(userMessage) {
        const msg = String(userMessage || '');
        const req = {
            chartType: null,
            chartMode: null,
            agg: null,
            ratio: null,
            orientation: null,
            emphasis: [],
            extraPages: [],
            sections: [],
            narrative: false,
            needsLlm: false,
            recognized: []
        };
        if (!msg.trim()) return req;

        // 0a. 版式比例 / 页面方向（竖版）
        for (const rule of RATIO_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.ratio = rule.ratio;
                req.recognized.push(rule.ratio === '16x9' ? '宽屏 16:9' : '标准屏 4:3');
                break;
            }
        }
        for (const rule of RATIO_ORIENTATION_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.orientation = rule.orientation;
                req.recognized.push('竖版');
                break;
            }
        }

        // 0b. 图表形式（图片 / 带数据的原生图）
        for (const rule of CHART_MODE_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.chartMode = rule.mode;
                req.recognized.push(rule.mode === 'picture' ? '图表用图片' : '图表带数据（可编辑）');
                break;
            }
        }

        // 1. 图表类型
        for (const rule of CHART_TYPE_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.chartType = rule.type;
                req.recognized.push(CHART_TYPE_LABEL[rule.type]);
                break;
            }
        }

        // 1b. 聚合方式（用户驱动模型显式指定 sum/avg/count；不命中默认数据驱动判别）
        for (const rule of AGG_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.agg = rule.agg;
                req.recognized.push(rule.label);
                break;
            }
        }

        // 2. 页面类型
        for (const rule of PAGE_RULES) {
            if (rule.keywords.some(k => msg.includes(k))) {
                req.extraPages.push(rule.type);
                req.recognized.push(PAGE_LABEL[rule.type]);
            }
        }

        // 3. 强调重点
        for (const p of EMPHASIS_PATTERNS) {
            const m = msg.match(p);
            if (m && m[1] && m[1].length <= 12) {
                req.emphasis.push(m[1]);
                req.recognized.push(`重点：${m[1]}`);
            }
        }
        req.emphasis = [...new Set(req.emphasis)];

        // 4. 自定义章节
        for (const p of SECTION_PATTERNS) {
            const m = msg.match(p);
            if (m && m[1] && m[1].length <= 20) {
                req.sections.push(m[1]);
            }
        }
        req.sections = [...new Set(req.sections)];

        // 5. 叙事结构
        req.narrative = NARRATIVE_PATTERNS.some(p => p.test(msg));

        // 6. 复杂度判定：叙事结构 / 长消息多要求 → 降级 LLM
        req.needsLlm = req.narrative || detectComplex(msg, req);

        return req;
    }

    function detectComplex(msg, req) {
        if (msg.length > 100) return true;
        const distinctCount = (req.chartType ? 1 : 0) + req.extraPages.length + req.emphasis.length + req.sections.length;
        if (distinctCount >= 4) return true;
        return false;
    }

    // ------------------------------------------------------------
    // 应用：把需求应用到推荐 slides
    // ------------------------------------------------------------
    function applyRequirementsToDesign(slides, req) {
        if (!Array.isArray(slides)) return slides;
        if (!req) return slides;
        const result = slides.map(s => ({ ...s }));

        // 1. 图表类型覆盖：把第一个 chart 页的 chartType 改为用户指定
        if (req.chartType) {
            const chartSlide = result.find(s => String(s.type || '').toLowerCase() === 'chart');
            if (chartSlide) chartSlide.chartType = req.chartType;
        }

        // 1b. 聚合方式覆盖（用户驱动模型）：把全部 chart 页的 agg 改为用户指定
        if (req.agg) {
            for (const s of result) {
                if (String(s.type || '').toLowerCase() === 'chart') s.agg = req.agg;
            }
        }

        // 2. 增加页面类型（去重：已有同类型则跳过）
        for (const pageType of req.extraPages) {
            const exists = result.some(s => String(s.type || '').toLowerCase() === pageType);
            if (exists) continue;
            const slide = buildPageSlide(pageType, req);
            if (slide) result.push(slide);
        }

        // 3. 强调重点：注入 KPI / text 页
        if (req.emphasis.length > 0) {
            const kpiSlide = result.find(s => String(s.type || '').toLowerCase() === 'kpi');
            if (kpiSlide && Array.isArray(kpiSlide.items)) {
                for (const em of req.emphasis) {
                    if (!kpiSlide.items.some(it => String(it.label || '').includes(em))) {
                        kpiSlide.items.push({ label: em, value: '—' });
                    }
                }
            }
            const textSlide = result.find(s => String(s.type || '').toLowerCase() === 'text');
            if (textSlide && Array.isArray(textSlide.bullets)) {
                textSlide.bullets.unshift(`重点：${req.emphasis.join('、')}`);
            }
        }

        // 4. 自定义章节：插入 section 页（放在目录之后）
        if (req.sections.length > 0) {
            const agendaIdx = result.findIndex(s => String(s.type || '').toLowerCase() === 'agenda');
            const insertAt = agendaIdx >= 0 ? agendaIdx + 1 : 1;
            const sectionSlides = req.sections.map((sec, i) => ({
                type: 'section',
                number: String(i + 1).padStart(2, '0'),
                title: sec,
                subtitle: ''
            }));
            result.splice(insertAt, 0, ...sectionSlides);
        }

        return result;
    }

    function buildPageSlide(pageType, req) {
        switch (pageType) {
            case 'comparison':
                return { type: 'comparison', title: '对比分析', leftTitle: '方案 A', rightTitle: '方案 B', leftItems: [], rightItems: [] };
            case 'timeline':
                return { type: 'timeline', title: '时间线', items: [] };
            case 'table':
                return { type: 'table', title: '数据明细', columns: [], rowCount: 0, startRow: 2 };
            case 'kpi':
                return { type: 'kpi', title: '关键指标', items: [] };
            case 'text':
                return { type: 'text', title: '总结与建议', bullets: [] };
            default:
                return null;
        }
    }

    // ------------------------------------------------------------
    // 展示标签
    // ------------------------------------------------------------
    const CHART_TYPE_LABEL = { line: '折线图', column: '柱状图', pie: '饼图' };
    const PAGE_LABEL = {
        comparison: '对比页', timeline: '时间线页', table: '数据明细页',
        kpi: '指标页', text: '总结页'
    };

    // ------------------------------------------------------------
    // 导出
    // ------------------------------------------------------------
    App.ChatV2.services.pptRequirementParser = {
        parsePptRequirements,
        applyRequirementsToDesign,
        needsLlm: (userMessage) => parsePptRequirements(userMessage).needsLlm,
        CHART_TYPE_LABEL,
        PAGE_LABEL
    };
    ChatLog.info?.('CHAT-V2', '✅ pptRequirementParser loaded');
})();
