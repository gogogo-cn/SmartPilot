/**
 * ============================================================
 * dashboard-blueprint-engine.js — 看板蓝图引擎 · 纯计算层 + 组合器（方案 B）
 * ============================================================
 * 职责：
 *   1. 纯数据计算：KPI 聚合、图表分组/求和、明细表派生（computeFromSources）
 *   2. 方向 C 确定性自检：饼图分组少→柱状图、分组多→top10+其他（validateAndAutoFix）
 *   3. buildPlan 组合器：按 deliverable 调 excelRenderer + htmlRenderer，保持旧导出名不变
 *
 * 渲染层已分离为 dashboard-excel-renderer.js / dashboard-html-renderer.js
 * 执行器（skill-spec-execute.js）零改动：所有旧函数名仍在 blueprintEngine 上
 *
 * 挂载：App.ChatV2.dashboardModules.blueprintEngine
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.dashboardModules) App.ChatV2.dashboardModules = {};

    const ChatLog = window.ChatLog || console;

    // ============================================================
    // 主题调色板（唯一事实源）
    // ============================================================
    const THEMES = {
        business_blue: {
            name: '商务蓝',
            excel: {
                title: '#1F4E79', label: '#6B7B8D', value: '#1F4E79',
                accent: '#2E75B6', border: '#D6E4F0', headerBg: '#F2F6FB',
                bodyBg: '#FFFFFF', headerText: '#1F4E79',
                kpiFill: '#F2F6FB', kpiLine: '#B8CCE4'
            },
            html: {
                bg: '#0B1C30', panel: '#102A44', panelBorder: '#2C5F8F',
                title: '#EAF2FB', sub: '#9DB8D8', kpiValue: '#4DA3FF',
                kpiLabel: '#B7CCE4', accent: '#2E75B6', grid: '#1F4468'
            }
        },
        dark: {
            name: '深色大屏',
            excel: {
                title: '#1F2937', label: '#4B5563', value: '#111827',
                accent: '#6D28D9', border: '#D1D5DB', headerBg: '#EDE9FE',
                bodyBg: '#FFFFFF', headerText: '#1F2937'
            },
            html: {
                bg: '#0B0F19', panel: '#141B2D', panelBorder: '#2A3550',
                title: '#E6E9F2', sub: '#8B93A7', kpiValue: '#8B7CF6',
                kpiLabel: '#A9B2C7', accent: '#6D28D9', grid: '#232D45'
            }
        },
        fresh_green: {
            name: '清新绿',
            excel: {
                title: '#1B5E20', label: '#4B5563', value: '#1B4332',
                accent: '#2E7D32', border: '#C8E6C9', headerBg: '#E8F5E9',
                bodyBg: '#FFFFFF', headerText: '#1B5E20'
            },
            html: {
                bg: '#0A1F14', panel: '#0F2E1E', panelBorder: '#2A5A3C',
                title: '#E8F5E9', sub: '#9DC6A8', kpiValue: '#43D17C',
                kpiLabel: '#A9D3B6', accent: '#2E7D32', grid: '#1E4A30'
            }
        },
        warm_orange: {
            name: '暖橙',
            excel: {
                title: '#BF360C', label: '#4B5563', value: '#8C2F0A',
                accent: '#E65100', border: '#FFCCBC', headerBg: '#FBE9E7',
                bodyBg: '#FFFFFF', headerText: '#BF360C'
            },
            html: {
                bg: '#1F1006', panel: '#2E1A0B', panelBorder: '#5A3A20',
                title: '#FBE9E7', sub: '#CBA98F', kpiValue: '#FF8A3D',
                kpiLabel: '#D9BFA8', accent: '#E65100', grid: '#4A2E16'
            }
        },
        minimal_light: {
            name: '浅色极简',
            excel: {
                title: '#2F3A45', label: '#6B7280', value: '#374151',
                accent: '#64748B', border: '#E5E7EB', headerBg: '#F1F5F9',
                bodyBg: '#FFFFFF', headerText: '#374151'
            },
            html: {
                bg: '#F5F7FA', panel: '#FFFFFF', panelBorder: '#E5E7EB',
                title: '#1F2937', sub: '#94A3B8', kpiValue: '#475569',
                kpiLabel: '#64748B', accent: '#64748B', grid: '#E5E9F0'
            }
        },
        tech_black: {
            name: '黑蓝科技',
            excel: {
                title: '#0B1220', label: '#475569', value: '#0F172A',
                accent: '#0284C7', border: '#CBD5E1', headerBg: '#E0F2FE',
                bodyBg: '#FFFFFF', headerText: '#0B1220'
            },
            html: {
                bg: '#05070D', panel: '#0B1220', panelBorder: '#1E3A5F',
                title: '#E6F1FF', sub: '#7B8CA6', kpiValue: '#38BDF8',
                kpiLabel: '#93A5BE', accent: '#38BDF8', grid: '#12233B'
            }
        },
        executive_gray: {
            name: '高级灰',
            excel: {
                title: '#1F2937', label: '#6B7280', value: '#111827',
                accent: '#6B7280', border: '#D1D5DB', headerBg: '#E5E7EB',
                bodyBg: '#FFFFFF', headerText: '#1F2937'
            },
            html: {
                bg: '#101418', panel: '#181D24', panelBorder: '#2E3540',
                title: '#E5E9F0', sub: '#8A93A3', kpiValue: '#B6C0CE',
                kpiLabel: '#9BA7B8', accent: '#8A93A3', grid: '#232B36'
            }
        },
        crimson_red: {
            name: '中国红',
            excel: {
                title: '#9B1B1B', label: '#6B7280', value: '#7F1D1D',
                accent: '#DC2626', border: '#FECACA', headerBg: '#FEE2E2',
                bodyBg: '#FFFFFF', headerText: '#9B1B1B', kpiFill: '#FECACA'
            },
            html: {
                bg: '#1A0808', panel: '#2A1010', panelBorder: '#7F1D1D',
                title: '#FEE2E2', sub: '#F5A5A5', kpiValue: '#F87171',
                kpiLabel: '#FECACA', accent: '#DC2626', grid: '#4A1515'
            }
        },
        violet_purple: {
            name: '紫罗兰',
            excel: {
                title: '#5B21B6', label: '#6B7280', value: '#4C1D95',
                accent: '#7C3AED', border: '#DDD6FE', headerBg: '#EDE9FE',
                bodyBg: '#FFFFFF', headerText: '#5B21B6', kpiFill: '#DDD6FE'
            },
            html: {
                bg: '#12081F', panel: '#1E1033', panelBorder: '#5B21B6',
                title: '#EDE9FE', sub: '#C4B5FD', kpiValue: '#A78BFA',
                kpiLabel: '#DDD6FE', accent: '#8B5CF6', grid: '#2E1065'
            }
        }
    };
    const DEFAULT_THEME = 'business_blue';

    /** 从用户话术解析主题 id；未命中返回 null（绝不静默回落中国红） */
    function resolveThemeFromMessage(msg) {
        const s = String(msg || '');
        if (!s.trim()) return null;
        // 显式 id / 中文名
        for (const id of Object.keys(THEMES)) {
            if (s.indexOf(id) >= 0) return id;
            if (s.indexOf(THEMES[id].name) >= 0) return id;
        }
        // 按色相匹配（顺序：先特殊组合，再单色；用「蓝/红/紫」字面，避免 改.*红 误伤）
        if (/(黑蓝|科技黑|科技蓝)/.test(s)) return 'tech_black';
        if (/(中国红|朱红|大红|猩红|crimson)/i.test(s)) return 'crimson_red';
        if (/(紫色|紫罗兰|紫主题|紫配色|violet|purple|改成紫|改为紫|换成紫|调成紫|改.*紫|换.*紫)/i.test(s)) return 'violet_purple';
        if (/(红色|红主题|红配色|改成红|改为红|换成红|调成红)/.test(s) && !/紫/.test(s)) return 'crimson_red';
        if (/(商务蓝|蓝色|蓝主题|蓝配色|改成蓝|改为蓝|换成蓝|调成蓝)/.test(s) && !/黑蓝|紫/.test(s)) return 'business_blue';
        if (/(清新绿|绿色|绿主题|绿配色|改成绿|改为绿|换成绿)/.test(s)) return 'fresh_green';
        if (/(暖橙|橙色|橘色|橘黄|橙主题|改成橙|改为橙|换成橙)/.test(s)) return 'warm_orange';
        if (/(高级灰|灰色|灰主题|改成灰|改为灰)/.test(s)) return 'executive_gray';
        if (/(深色|暗色|黑夜|dark)/i.test(s) && /主题|配色|大屏/.test(s)) return 'dark';
        if (/(极简|浅色|素雅)/.test(s)) return 'minimal_light';
        return null;
    }

    /** 是否「改主题/换配色」意图（不含首次生成看板） */
    function isThemeChangeIntent(msg) {
        const s = String(msg || '');
        if (!s.trim()) return false;
        if (!/(主题|配色|换色|改色|颜色|色调)/.test(s)) return false;
        // 「生成红色看板」算生成而非改主题
        if (/(生成|创建|做个|制作).{0,8}看板/.test(s) && !/(改|换|修改|调整|变为|换成|调成)/.test(s)) return false;
        return !!(resolveThemeFromMessage(s) || /(改|换|修改|调整|变为|换成|调成).{0,8}(主题|配色|颜色)/.test(s));
    }

    function themeOf(spec) {
        let id = spec && spec.theme;
        if (id && !THEMES[id] && typeof id === 'string') {
            id = resolveThemeFromMessage(id) || id;
        }
        id = THEMES[id] ? id : DEFAULT_THEME;
        return THEMES[id];
    }
    function hexToRgb(hex) {
        const m = String(hex || '').replace('#', '').match(/^([0-9a-f]{6})$/i);
        if (!m) return null;
        const n = parseInt(m[1], 16);
        return [n >> 16 & 255, n >> 8 & 255, n & 255];
    }
    function rgbToHex(rgb) {
        if (!rgb) return null;
        const c = rgb.map(v => Math.max(0, Math.min(255, Math.round(Number(v) || 0))));
        return '#' + c.map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
    }

    // ============================================================
    // 工具：数值/列识别
    // ============================================================
    function num(v) {
        if (v === null || v === undefined || v === '' || typeof v === 'object') return null;
        const n = Number(v);
        return Number.isFinite(n) ? n : null;
    }
    /** 把"2025-01" / "2025/1/5" / "2025年1月" / "1月" / "1月5日" 解析为可比较的时间序号，不可解析返回 NaN */
    function parseTimeKey(s) {
        s = String(s == null ? '' : s).trim();
        let m = s.match(/^(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?$/);
        if (m) return Number(m[1]) * 10000 + Number(m[2]) * 100 + Number(m[3] || 1);
        m = s.match(/^(\d{4})年(\d{1,2})月(?:(\d{1,2})日?)?$/);
        if (m) return Number(m[1]) * 10000 + Number(m[2]) * 100 + Number(m[3] || 1);
        m = s.match(/^(\d{1,2})月(?:(\d{1,2})日?)?$/);
        if (m) return Number(m[1]) * 100 + Number(m[2] || 1);
        m = s.match(/^(\d{1,2})[-/](\d{1,2})$/);
        if (m) return Number(m[1]) * 100 + Number(m[2]);
        return NaN;
    }
    /** 判断分组 key 集合是否是时间/有序类别（>=60% 可解析为日期），是则按时间升序而非按值降序 */
    function looksLikeTimeDim(keys) {
        if (!keys || keys.length < 2) return false;
        let hit = 0;
        for (const k of keys) { if (!Number.isNaN(parseTimeKey(k))) hit++; }
        return hit / keys.length >= 0.6;
    }
    function isDateHeader(h) { return /日期|时间|date|time/i.test(String(h || '')); }
    function isIdHeader(h) {
        const s = String(h || '');
        // 中英文 ID 类表头：订单编号/商品编号/序号/流水号/编码/id（"编号"前是汉字也命中，如"订单编号"）
        return /(^|[^a-z0-9])id($|[^a-z0-9])/i.test(s) || /编号|单号|序号|编码|流水号|单id|id号/i.test(s);
    }
    function numericColumns(source) {
        const headers = source.headers || [];
        // ID/日期列即使是数字也不是度量
        return headers.map((h, i) => ({
            h, i,
            numeric: !isIdHeader(h) && !isDateHeader(h) && (source.rows || []).some(r => num(r[h]) !== null)
        }));
    }
    function firstNumericHeader(source) {
        const cols = numericColumns(source);
        const hit = cols.find(c => c.numeric);
        return hit ? hit.h : null;
    }
    function allNumericHeaders(source) {
        return numericColumns(source).filter(c => c.numeric).map(c => c.h);
    }
    function guessDimensionHeader(source) {
        const headers = source.headers || [];
        const rows = source.rows || [];
        const cols = headers.map((h, i) => {
            const values = rows.map(r => (r[h] !== undefined && r[h] !== null && String(r[h]) !== '') ? String(r[h]) : null).filter(Boolean);
            return { h, i, uniq: new Set(values).size, total: values.length };
        });
        const textCols = cols.filter(c => num(c.h) === null && !isDateHeader(c.h) && !isIdHeader(c.h) && !rows.some(r => num(r[c.h]) !== null));
        if (textCols.length) {
            const hit = textCols
                .filter(c => c.uniq >= 2 && c.uniq <= 24)
                .sort((a, b) => b.uniq - a.uniq)[0];
            return hit ? hit.h : textCols[0].h;
        }
        return cols.length ? cols[0].h : null;
    }

    /**
     * 枚举全部合格分类维度：文本列、非 ID/日期、唯一值 2~上限。
     * 上限取 min(24, max(8, 行数×0.5))——基数接近行数（如"客户名称"几乎一人一值）
     * 的实体列没有聚合分析价值，排除。按基数升序，最多 max 个。
     */
    function enumerateDimensions(source, max) {
        const headers = source.headers || [];
        const rows = source.rows || [];
        const uniqCap = Math.min(24, Math.max(8, Math.floor(rows.length * 0.5)));
        const out = [];
        for (const h of headers) {
            if (isDateHeader(h) || isIdHeader(h)) continue;
            // 数值列不作为分类维度
            if (rows.some(r => num(r[h]) !== null)) continue;
            const vals = rows.map(r => (r[h] !== undefined && r[h] !== null && String(r[h]) !== '') ? String(r[h]) : null).filter(Boolean);
            const uniq = new Set(vals).size;
            if (uniq >= 2 && uniq <= uniqCap) out.push({ dim: h, uniq });
        }
        out.sort((a, b) => a.uniq - b.uniq);
        return out.slice(0, max || 4).map(x => x.dim);
    }

    /**
     * 智能图表类型推荐：根据数据形状选最合适的图表。
     * 时间维→折线/面积图；多数值列对比→瀑布图；层级型→树状/旭日；
     * 分布型→直方图；单一数值对比→柱状/条形；占比→饼/环形/漏斗。
     */
    function recommendChartType(source, dim, value) {
        const rows = source.rows || [];
        const vals = rows.map(r => r[dim] == null ? '' : String(r[dim])).filter(Boolean);
        const isDateLike = looksLikeTimeDim(vals);
        const uniqueVals = new Set(vals).size;
        const numCols = allNumericHeaders(source).length;

        // 时间维度 → 折线图（趋势展示）
        if (isDateLike) return 'xlLine';
        // 分组数少（2-6）且占比场景 → 环形图（百分比一眼看出）
        if (uniqueVals >= 2 && uniqueVals <= 8 && numCols === 1) return 'xlDoughnut';
        // 分组数中等 → 漏斗图（转化率/占比场景）
        if (uniqueVals >= 3 && uniqueVals <= 12 && numCols === 1) return 'xlFunnel';
        // 分类维度多值对比 → 条形图（横向，长标签不拥挤）
        if (uniqueVals >= 8) return 'xlBarClustered';
        // 默认柱状图
        return 'xlColumnClustered';
    }

    /**
     * 自动图表方案：智能推荐多种图表类型，每个合格维度一张图。
     * 上限 cap 张（默认 6，3×2 网格，比原来 4 张更丰富）。
     */
    function buildAutoChartSpecs(source, cap) {
        if (!source) return [];
        const max = cap || 6;
        const dims = enumerateDimensions(source, Math.min(max, 6));
        const val = firstNumericHeader(source);
        const specs = [];
        // 按图表序号轮换类型，避免连续同类型（全是甜甜圈）
        // 0=条形图(区域对比) 1=环形图(类别占比) 2=柱状图(人员对比) 3=折线/面积
        const _typeRotation = ['xlBarClustered', 'xlDoughnut', 'xlColumnClustered', 'xlLineMarkers'];
        dims.forEach((dim, idx) => {
            let ct = recommendChartType(source, dim, val);
            // 时间维度保持折线；其余按序号轮换，避免视觉单调
            if (!/xlLine|xlArea/.test(ct)) {
                ct = _typeRotation[idx % _typeRotation.length];
            }
            if (val) {
                specs.push({ title: '各' + dim + val, dimension: dim, value: val, func: 'sum', chartType: ct, source: source.id });
            } else {
                const countCol = (source.headers || []).find(h => h !== dim && !isDateHeader(h) && !isIdHeader(h)) || dim;
                specs.push({ title: '各' + dim + '分布', dimension: dim, value: countCol, func: 'count', chartType: ct, source: source.id });
            }
        });
        return specs.slice(0, max);
    }

    /**
     * 排行榜（Top N）组件数据：按维度分组聚合，取 Top N + 其他。
     * @param {object} source - 数据源
     * @param {string} dim - 维度列名
     * @param {string} value - 数值列名
     * @param {string} func - 聚合函数 sum/avg/count
     * @param {number} topN - 取前 N 名，默认 5
     */
    function computeTopNRanking(source, dim, value, func, topN) {
        if (!source) return null;
        const rows = source.rows || [];
        const val = value || firstNumericHeader(source);
        const n = topN || 5;
        const groups = new Map();
        for (const r of rows) {
            const key = (r[dim] !== undefined && r[dim] !== null && String(r[dim]) !== '') ? String(r[dim]) : '未分类';
            if (!groups.has(key)) groups.set(key, { sum: 0, count: 0 });
            const g = groups.get(key);
            const v = num(r[val]);
            if (v !== null) {
                g.sum += v;
                g.count += 1;
            }
        }
        const arr = [...groups.entries()].map(([k, g]) => ({
            name: k,
            value: func === 'avg' && g.count > 0 ? Math.round((g.sum / g.count) * 100) / 100 : Math.round(g.sum * 100) / 100
        })).sort((a, b) => b.value - a.value);
        const total = arr.reduce((s, x) => s + x.value, 0);
        const top = arr.slice(0, n);
        if (arr.length > n) {
            const restSum = arr.slice(n).reduce((s, x) => s + x.value, 0);
            top.push({ name: '其他', value: Math.round(restSum * 100) / 100 });
        }
        return { dim, val, func, topN: n, total, items: top };
    }

    /**
     * 仪表盘（Gauge）组件数据：基于目标值的完成率 / 百分比指标。
     * 适用于"完成率""达标率""覆盖率"等百分比场景。
     * @param {number} value - 当前值
     * @param {number} target - 目标值，默认 100（百分比场景）
     * @param {string} label - 指标标签
     */
    function computeGaugeData(value, target, label) {
        const t = target || 100;
        const v = Math.max(0, Math.min(Number.isFinite(Number(value)) ? Number(value) : 0, t));
        const pct = Math.round((v / t) * 1000) / 10;
        // 根据完成率选色段：<60 红，60-80 黄，>=80 绿
        const color = pct >= 80 ? '#43D17C' : (pct >= 60 ? '#F5C044' : '#F2728F');
        return { label: label || '完成率', value: v, target: t, percent: pct, color };
    }

    /**
     * KPI 同比/环比计算：自动识别时间列，对比"本期 vs 上期/去年同期"。
     * 返回 { label, current, previous, change, changePct, trend: 'up'|'down'|'flat' }。
     * @param {object} source - 数据源
     * @param {string} dateCol - 日期列名（自动推断时可留空）
     * @param {string} valueCol - 数值列名
     * @param {string} func - 聚合函数 sum/avg
     */
    function computeKpiTrend(source, dateCol, valueCol, func) {
        if (!source) return null;
        const rows = source.rows || [];
        const headers = source.headers || [];
        // 自动推断日期列
        let dc = dateCol;
        if (!dc) {
            dc = headers.find(h => /日期|时间|date|time|年|月|季度/i.test(h));
        }
        if (!dc) return null;
        const vc = valueCol || firstNumericHeader(source);
        if (!vc) return null;
        const f = (func || 'sum').toLowerCase();
        // 按年-月分组
        const periodAgg = new Map();
        for (const r of rows) {
            const dv = r[dc];
            if (dv === null || dv === undefined || String(dv).trim() === '') continue;
            const p = String(dv).match(/^(\d{4})[-年](\d{1,2})/) || String(dv).match(/^(\d{4})/);
            if (!p) continue;
            const key = p[1] + '-' + (p[2] || '01').padStart(2, '0');
            if (!periodAgg.has(key)) periodAgg.set(key, { sum: 0, count: 0 });
            const g = periodAgg.get(key);
            const v = num(r[vc]);
            if (v !== null) { g.sum += v; g.count += 1; }
        }
        const sortedKeys = [...periodAgg.keys()].sort();
        if (sortedKeys.length < 2) return null;
        const latest = sortedKeys[sortedKeys.length - 1];
        const previous = sortedKeys[sortedKeys.length - 2];
        const curG = periodAgg.get(latest);
        const prevG = periodAgg.get(previous);
        const cur = f === 'avg' && curG.count > 0 ? curG.sum / curG.count : curG.sum;
        const prev = f === 'avg' && prevG.count > 0 ? prevG.sum / prevG.count : prevG.sum;
        let change = cur - prev;
        let changePct = prev !== 0 ? (change / Math.abs(prev)) * 100 : 0;
        change = Math.round(change * 100) / 100;
        changePct = Math.round(changePct * 10) / 10;
        const trend = change > 0 ? 'up' : (change < 0 ? 'down' : 'flat');
        return {
            label: vc,
            periodCurrent: latest, periodPrevious: previous,
            current: Math.round(cur * 100) / 100,
            previous: Math.round(prev * 100) / 100,
            change, changePct, trend
        };
    }

    // ============================================================
    // KPI / 图表聚合（纯内存计算）
    // ============================================================
    function computeKpi(kpi, source) {
        const hasFunc = typeof kpi.func === 'string' && kpi.func.trim() !== '';
        const hasCol = typeof kpi.column === 'string' && kpi.column.trim() !== '';
        const hasLit = kpi.value !== undefined && kpi.value !== null && kpi.value !== '';
        // 字面值仅当未声明计算方式、且不是 0 占位时采用；否则一律以数据计算为准
        // （防 LLM 规划时输出 value:0 占位符覆盖真实计算，导致看板全是 0）
        if (!hasFunc && !hasCol && hasLit && Number(kpi.value) !== 0) {
            return { label: kpi.label, unit: kpi.unit || '', value: kpi.value, numeric: typeof kpi.value === 'number' && Number.isFinite(kpi.value) };
        }
        const rows = source.rows || [];
        const func = (hasFunc ? String(kpi.func) : 'sum').toLowerCase();
        if (func === 'count') return { label: kpi.label, unit: kpi.unit || '', value: rows.length, numeric: true, computed: true, func: 'count' };
        if (func === 'distinct') {
            const dCol = hasCol ? kpi.column.trim() : (source.headers || []).find(h => !/金额|额|价|费|成本|利润|编号|单号|序号|编码|流水号|日期|时间/i.test(h));
            if (dCol) {
                const distinct = new Set(rows.map(r => (r[dCol] !== undefined && r[dCol] !== null && String(r[dCol]) !== '') ? String(r[dCol]) : '未分类')).size;
                return { label: kpi.label, unit: kpi.unit || '', value: distinct, numeric: true, computed: true, func: 'distinct', column: dCol };
            }
            return { label: kpi.label, unit: kpi.unit || '', value: rows.length, numeric: true, computed: true, func: 'count' };
        }

        // 纯标签声明（无 func/column/value，如 LLM 只写 label:"销售区域数"）→ 智能推断：
        //   1) label 以「维度列名+数」结尾 → 该维度 distinct 计数（修复"销售区域数=总金额"的误算）
        //   2) label 含 量/单/笔/条/行 且以"数"结尾 → 记录数
        if (!hasFunc && !hasCol && !hasLit) {
            const lbl = String(kpi.label || '').trim();
            const dimHit = (source.headers || []).find(h => lbl === h + '数' || lbl === h);
            if (dimHit && !/金额|额|价|费|成本|利润/.test(lbl)) {
                const distinct = new Set(rows.map(r => (r[dimHit] !== undefined && r[dimHit] !== null && String(r[dimHit]) !== '') ? String(r[dimHit]) : '未分类')).size;
                return { label: lbl, unit: kpi.unit || '', value: distinct, numeric: true, computed: true, func: 'distinct', column: dimHit };
            }
            if (/数$/.test(lbl) && /量|单|笔|条|行|人|个|家/.test(lbl)) {
                return { label: lbl, unit: kpi.unit || '', value: rows.length, numeric: true, computed: true, func: 'count' };
            }
        }

        const col = hasCol ? kpi.column.trim() : firstNumericHeader(source);
        const vals = col ? rows.map(r => num(r[col])).filter(v => v !== null) : [];
        if (vals.length === 0) {
            // 指定列不存在或无数值 → 标记 noNumeric，由 validateAndAutoFix 统一降级为计数口径，杜绝误导性 0
            return { label: kpi.label, unit: kpi.unit || '', value: 0, numeric: false, func, column: col || undefined, noNumeric: true };
        }
        let v;
        if (func === 'avg') v = vals.reduce((a, b) => a + b, 0) / vals.length;
        else if (func === 'max') v = Math.max.apply(null, vals);
        else if (func === 'min') v = Math.min.apply(null, vals);
        else v = vals.reduce((a, b) => a + b, 0);
        // label 以括号结尾（如"订单金额(元)"/"平均订单金额（元）"）且未声明单位 → 拆成「标题 + 单位」，卡片展示更清晰
        let label = kpi.label, unit = kpi.unit || '';
        if (!unit) {
            const lm = String(label).match(/^(.+?)[（(]([^（()]{1,6})[）)]$/);
            if (lm && lm[1].trim() && lm[2].trim()) { label = lm[1].trim(); unit = lm[2].trim(); }
        }
        // column 保留给 HTML 大屏前端筛选重算使用
        return { label, unit, value: Math.round(v * 100) / 100, numeric: true, computed: true, func: func === 'avg' ? 'avg' : func, column: col };
    }

    function aggregateChartData(chart, source) {
        const rows = source.rows || [];
        const headers = source.headers || [];
        // ── 维度字段名校验 + 模糊匹配纠错 ──
        // 兼容 dim / dimension、value / vals（一键工具与弱模型蓝图常见写法）
        const _dimRaw = (typeof chart.dimension === 'string' && chart.dimension.trim())
            ? chart.dimension.trim()
            : (typeof chart.dim === 'string' && chart.dim.trim()) ? chart.dim.trim() : '';
        let dim = _dimRaw || guessDimensionHeader(source);
        if (dim && headers.length && headers.indexOf(dim) < 0) {
            // 精确找不到 → 模糊匹配："销售" → "销售额"；"地区" → "所在地区"；"产品" → "产品名称"
            let best = null; let bestScore = 0;
            const tgt = dim.toLowerCase();
            for (const h of headers) {
                const hl = String(h).toLowerCase();
                let score = 0;
                if (hl === tgt) score = 100;
                else if (hl.includes(tgt) || tgt.includes(hl)) score = Math.max(hl.length, tgt.length);
                if (score > bestScore) { bestScore = score; best = h; }
            }
            if (best && bestScore >= 2) {
                ChatLog.info?.('BLUEPRINT', `[dim-auto-fix] 维度字段「${dim}」不存在 → 自动纠正为「${best}」`);
                dim = best;
            } else {
                // 完全匹配不到 → 回退到自动推断的维度列
                const guessed = guessDimensionHeader(source);
                if (guessed && guessed !== dim) {
                    ChatLog.info?.('BLUEPRINT', `[dim-auto-fix] 维度字段「${dim}」完全匹配不到 → 自动推断为「${guessed}」`);
                    dim = guessed;
                }
            }
        }
        let valCols;
        if (typeof chart.value === 'string' && chart.value.trim()) {
            valCols = [chart.value.trim()];
        } else if (Array.isArray(chart.vals) && chart.vals.length) {
            valCols = chart.vals.map(v => String(v || '').trim()).filter(Boolean);
        } else if (Array.isArray(chart.values) && chart.values.length) {
            valCols = chart.values.map(v => String(v || '').trim()).filter(Boolean);
        } else {
            valCols = allNumericHeaders(source).slice(0, 3) || [];
        }
        // ── 数值字段名校验：逐个检查，不在 headers 的剔除 ──
        valCols = valCols.filter(v => headers.indexOf(v) >= 0);
        if (valCols.length === 0) valCols = allNumericHeaders(source).slice(0, 1);
        if (!dim || valCols.length === 0 || rows.length === 0) return null;
        const func = String(chart.func || 'sum').toLowerCase();
        // count 型只画一个系列（行数），多列会导致完全重叠的柱子
        if (func === 'count') valCols.splice(1);
        const groups = new Map();
        for (const r of rows) {
            const key = (r[dim] !== undefined && r[dim] !== null && String(r[dim]) !== '') ? String(r[dim]) : '未分类';
            if (!groups.has(key)) groups.set(key, { vals: valCols.map(() => 0), count: 0 });
            const g = groups.get(key);
            g.count += 1;
            valCols.forEach((c, i) => { const v = num(r[c]); if (v !== null) g.vals[i] += v; });
        }
        const entries = [...groups.entries()];
        if (func === 'count') {
            entries.forEach(([, g]) => { g.vals.forEach((_, i) => { g.vals[i] = g.count; }); });
        }
        const isTrend = /xlLine|xlArea/.test(chart.chartType);
        // 时间维度（月份/季度/日期）保持时间升序，不按值降序打乱时间轴；
        // 非时间维度的柱状/饼图仍按首列值降序（Top 项在左）。
        const timeDim = looksLikeTimeDim(entries.map(e => e[0]));
        if (timeDim) {
            entries.sort((a, b) => parseTimeKey(a[0]) - parseTimeKey(b[0]));
        } else if (!isTrend) {
            entries.sort((a, b) => b[1].vals[0] - a[1].vals[0]);
        }
        const header = [dim, ...valCols];
        const data = entries.map(([k, g]) => [k, ...g.vals.map(v => Math.round(v * 100) / 100)]);
        // func 保留给 HTML 大屏前端筛选重算使用（count 图与 sum 图重算逻辑不同）
        return { header, data, dim, vals: valCols, chartType: chart.chartType, func };
    }

    // ============================================================
    // 布局常量
    // ============================================================
    const LAYOUT = {
        titleRow: 1, subtitleRow: 3,
        kpiLabelRow: 4, kpiValueRow: 5, kpiCardEndRow: 7,
        chartStartRow: 8, centerStartCol: 1, chartSlotCols: 6, chartSlotRows: 12, chartRowGap: 12, chartColGap: 1,
        kpiColWidth: 16, tableMaxRows: 20, dashboardCols: 16,
        // Excel 原生切片器布局（磅坐标；透视表统一放"透视缓存"工作表）
        slicerTopPx: 99.75, slicerLeftStart: 760, slicerLeftStep: 160, slicerGap: 0, slicerWidth: 180, slicerHeight: 114,
        pivotSheet: '透视缓存', pivotRowStep: 30
    };
    function colLetter(n) {
        let s = '';
        while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
        return s;
    }

    // ============================================================
    // 骨架（规划步骤）
    // ============================================================
    function readSourceCount(spec) {
        const declared = Array.isArray(spec.dataSources) ? spec.dataSources.length : 0;
        return Math.max(1, Math.min(3, declared || 1));
    }
    function readShells(spec) {
        const n = readSourceCount(spec);
        const shells = [];
        for (let i = 0; i < n; i++) {
            const ds = (Array.isArray(spec.dataSources) && spec.dataSources[i]) ? spec.dataSources[i] : null;
            const label = ds && ds.id ? ds.id : '数据源 ' + (i + 1);
            const loc = ds && ds.range ? ds.range : (ds && ds.sheet ? ds.sheet + '（整表）' : (i === 0 ? '当前选区' : '当前选区'));
            shells.push({ title: `读取数据源「${label}」`, detail: loc });
        }
        return shells;
    }
    function renderShells(spec, prepared) {
        const kpis = (prepared && Array.isArray(prepared.kpis)) ? prepared.kpis
            : (Array.isArray(spec.kpis) ? spec.kpis : []);
        const kpiCount = kpis.length;
        const charts = (prepared && Array.isArray(prepared.charts)) ? prepared.charts
            : (Array.isArray(spec.charts) ? spec.charts : []);
        const deliverable = String(spec.deliverable || 'dual').toLowerCase();
        const excelOnly = deliverable === 'excel';
        const htmlOnly = deliverable === 'html';
        const shells = [];
        if (!htmlOnly) {
            shells.push({ title: '创建看板工作表', detail: spec.sheet || '看板' });
            shells.push({ title: '标题横幅（形状）', detail: '主题底色圆角卡片（含副标题）' });
            for (let i = 0; i < kpiCount; i++) {
                const k = kpis[i] || {};
                const name = k.label || ('指标 ' + (i + 1));
                shells.push({ title: `卡片「${name}」形状`, detail: '圆角卡片 · 数值+标签' });
            }
            shells.push({ title: 'KPI 卡片列宽', detail: '固定列宽 ' + LAYOUT.kpiColWidth });
            const slicerShells = (Array.isArray(spec.slicers) && spec.slicers.length)
                ? spec.slicers.slice(0, 4)
                : charts.slice(0, 3).map(c => c.dim).filter((d, i, a) => d && a.indexOf(d) === i);
            if (slicerShells.length && charts.length) {
                shells.push({ title: '创建原生切片器', detail: slicerShells.join('、') + '（联动 ' + charts.length + ' 张图表）' });
            }
            for (let i = 0; i < charts.length; i++) {
                const c = charts[i];
                if (!c) continue;
                shells.push({ title: `生成图表「${c.title || '图表 ' + (i + 1)}」`, detail: humanizeChartType(c.chartType) });
            }
            shells.push({ title: '写入明细表', detail: (spec.table && spec.table.maxRows ? spec.table.maxRows : LAYOUT.tableMaxRows) + ' 行' });
            shells.push({ title: '明细表表头样式', detail: '加粗 · 主题色' });
            shells.push({ title: '明细表数据样式', detail: '边框 · 浅色底' });
        }
        if (!excelOnly) {
            shells.push({ title: '生成 HTML 大屏', detail: (spec.title || '数据看板') + '大屏' });
        }
        return shells;
    }
    function buildSkeletonSteps(spec, prepared) {
        return [...readShells(spec), ...renderShells(spec, prepared)];
    }

    function humanizeChartType(ct) {
        const map = {
            xlColumnClustered: '柱状图', xlColumnStacked: '堆叠柱状图', xlColumnStacked100: '100%堆叠柱状图',
            xlBarClustered: '条形图', xlBarStacked: '堆叠条形图', xlBarStacked100: '100%堆叠条形图',
            xlLine: '折线图', xlLineMarkers: '折线图（带点）', xlArea: '面积图',
            xlPie: '饼图', xlDoughnut: '环形图', xlXYScatter: '散点图',
            xlRadar: '雷达图', xlRadarFilled: '填充雷达图', xlBubble: '气泡图',
            xlSunburst: '旭日图', xlTreemap: '树状图', xlHistogram: '直方图',
            xlBoxwhisker: '箱线图', xlFunnel: '漏斗图', xlWaterfall: '瀑布图'
        };
        return map[ct] || String(ct || '图表');
    }
    function isCategoryChart(ct) {
        return /xlPie|xlDoughnut|xlFunnel|xlWaterfall|xlTreemap|xlSunburst/.test(String(ct || ''));
    }
    /** 图表类型分类：用于 HTML 渲染器路由到正确的 ECharts 系列类型 */
    function chartFamily(ct) {
        const s = String(ct || '').toLowerCase();
        if (s === 'xlline' || s === 'xllinemarkers') return 'line';
        if (s === 'xlarea') return 'area';
        if (s === 'xlpie') return 'pie';
        if (s === 'xldoughnut') return 'doughnut';
        if (s === 'xlbarclustered' || s === 'xlbarstacked' || s === 'xlbarstacked100') return 'hbar';
        if (s === 'xlxyscatter' || s === 'xlbubble') return 'scatter';
        if (s === 'xlradar' || s === 'xlradarfilled') return 'radar';
        if (s === 'xlfunnel') return 'funnel';
        if (s === 'xlwaterfall') return 'waterfall';
        if (s === 'xltreemap') return 'treemap';
        if (s === 'xlsunburst') return 'sunburst';
        if (s === 'xlhistogram') return 'histogram';
        return 'column';
    }

    // ============================================================
    // 读取步骤（真实工具参数）
    // ============================================================
    function buildReadSteps(spec, ctx) {
        const n = readSourceCount(spec);
        const sel = ctx && ctx.selection ? ctx.selection : {};
        const selAddr = String(sel.address || sel.range || '').replace(/\$/g, '');
        const selSheet = sel.sheetName || '';
        const steps = [];
        for (let i = 0; i < n; i++) {
            const ds = (Array.isArray(spec.dataSources) && spec.dataSources[i]) ? spec.dataSources[i] : null;
            const label = ds && ds.id ? ds.id : '数据源 ' + (i + 1);
            const args = { returnFormat: 'object', hasHeaders: true };
            let loc = '当前选区';
            if (ds && (ds.range || ds.sheet)) {
                if (ds.range && ds.range.includes('!')) {
                    const m = String(ds.range).match(/^(.*)!(.+)$/);
                    args.sheet = m[1].trim().replace(/^['"]|['"]$/g, '');
                    args.range = m[2];
                } else if (ds.range) {
                    args.range = ds.range;
                }
                if (ds.sheet && !args.sheet) args.sheet = ds.sheet;
                loc = ds.range || ds.sheet + '（整表）';
            } else if (i === 0 && selAddr) {
                args.range = selAddr;
                if (selSheet) args.sheet = selSheet;
                loc = selAddr + (selSheet ? '（' + selSheet + '）' : '');
            }
            steps.push({
                title: `读取数据源「${label}」`,
                detail: loc,
                tool: 'excel_read_range',
                args,
                displayName: `读取数据源「${label}」`,
                displayArgs: { '数据源': label, '范围': loc }
            });
        }
        return steps;
    }

    // ============================================================
    // rows 格式归一化：后端 excel_read_range 返回二维数组，
    // 蓝图引擎全部用 r[h] 对象索引，这里统一转为对象数组
    // ============================================================
    function _normalizeSources(sources) {
        if (!Array.isArray(sources)) return sources;
        for (const src of sources) {
            if (!src || !Array.isArray(src.rows) || src.rows.length === 0) continue;
            const headers = (src.headers || []).map(String);
            const dedup = [];
            const seen = new Map();
            for (const h of headers) {
                if (seen.has(h)) { seen.set(h, seen.get(h) + 1); dedup.push(h + '_' + seen.get(h)); }
                else { seen.set(h, 0); dedup.push(h); }
            }
            if (dedup.some((h, i) => h !== headers[i])) {
                ChatLog.warn?.('BLUEPRINT', `[_normalizeSources] 发现重复列名，已加后缀去重: ${headers.filter((h, i) => dedup[i] !== h).join(',')} → ${dedup.filter((h, i) => h !== headers[i]).join(',')}`);
            }
            src.headers = dedup;
            if (Array.isArray(src.rows[0])) {
                ChatLog.info?.('BLUEPRINT', `[_normalizeSources] rows 是二维数组 → 转换为对象数组（${src.rows.length} 行 × ${dedup.length} 列）`);
                src.rows = src.rows.map(row => {
                    const obj = {};
                    dedup.forEach((h, i) => { obj[h] = row[i]; });
                    return obj;
                });
            } else if (typeof src.rows[0] === 'object' && src.rows[0] !== null) {
                const keys = Object.keys(src.rows[0]);
                if (keys.length !== dedup.length || keys.some((k, i) => k !== dedup[i])) {
                    ChatLog.warn?.('BLUEPRINT', `[_normalizeSources] rows 已是对象数组但 keys 与 headers 不匹配 → 强制重建对象（headers=[${dedup.join(',')}] keys=[${keys.join(',')}]）`);
                    src.rows = src.rows.map(row => {
                        const obj = {};
                        dedup.forEach((h, i) => { obj[h] = row[keys[i]]; });
                        return obj;
                    });
                }
            }
            // 统一日期显示：`2025-02-04 00:00:00` / `2025-02-04T00:00:00`（零点时间戳）→ 纯日期；
            // 非零时间（如 08:30:00）保留完整，避免丢信息。
            for (const r of src.rows) {
                for (const h of dedup) {
                    const v = r[h];
                    if (typeof v === 'string' && v.length >= 10 && v.indexOf(':') !== -1) {
                        const m = v.match(/^(\d{4}[-/]\d{2}[-/]\d{2})[T ]00:00:00(?:\.\d+)?$/);
                        if (m) r[h] = m[1];
                    }
                }
            }
            // 反汇总 / 反残影：VB read_range 的 UsedRange 可能被图表锚点、
            // 自动扩展的表结构等撑大，读到两类伪行：
            //   (a) "汇总/总计/合计/Total/Grand Total" 这类分类汇总行
            //   (b) 和前一行完全相同的"残影行"（Excel 自动填充残留）
            //   (c) 所有分类列值相同但数值列被算出来是总和的行（也是汇总）
            const _isSummaryCell = (v) => {
                if (v === null || v === undefined) return false;
                const s = String(v).trim().toLowerCase();
                return /^(汇总|总计|合计|小计|合計|總計|sum|total|grand total|求和|总计项)$/.test(s);
            };
            // 分类列 = 非日期、非数值的列（维度列）
            const _isDimCol = (h) => {
                const svals = src.rows.map(r => r[h]).filter(v => v !== null && v !== undefined && v !== '');
                if (svals.length === 0) return false;
                // 如果所有非空值都是 number 或可解析为 number → 数值列
                const allNum = svals.every(v => typeof v === 'number' || !isNaN(Number(v)));
                return !allNum;
            };
            const _dimCols = dedup.filter(_isDimCol);
            const _beforeRows = src.rows.length;
            const _filteredRows = [];
            for (let ri = 0; ri < src.rows.length; ri++) {
                const r = src.rows[ri];
                // (a) 维度列里有"汇总/总计" → 跳过
                let _hitSummary = false;
                for (const h of _dimCols) {
                    if (_isSummaryCell(r[h])) { _hitSummary = true; break; }
                }
                if (_hitSummary) {
                    ChatLog.info?.('BLUEPRINT', `[_normalizeSources] 过滤汇总行 row#${ri}: ${JSON.stringify(r).slice(0,150)}`);
                    continue;
                }
                // (b) 和前一行完全相同 → 残影，跳过
                if (_filteredRows.length > 0) {
                    const prev = _filteredRows[_filteredRows.length - 1];
                    let _allSame = true;
                    for (const h of dedup) {
                        if (String(r[h]) !== String(prev[h])) { _allSame = false; break; }
                    }
                    if (_allSame) {
                        ChatLog.info?.('BLUEPRINT', `[_normalizeSources] 过滤残影行 row#${ri}: 和前一行完全相同`);
                        continue;
                    }
                }
                _filteredRows.push(r);
            }
            if (_filteredRows.length !== _beforeRows) {
                src.rows = _filteredRows;
                ChatLog.info?.('BLUEPRINT', `[_normalizeSources] 过滤汇总/残影行: ${_beforeRows} → ${src.rows.length}`);
            }
            ChatLog.info?.('BLUEPRINT', `[_normalizeSources] 归一化完成: headers(${src.headers.length})=[${src.headers.slice(0,8).join(',')}] rows(${src.rows.length}) row0=${JSON.stringify(src.rows[0]).slice(0,200)}`);
        }
        return sources;
    }

    // ============================================================
    // computeFromSources：纯函数——sources → 内存计算 KPI/图表/明细
    // ============================================================
    function computeFromSources(sources, spec) {
        if (!Array.isArray(sources) || sources.length === 0) {
            throw new Error('未读取到有效数据：请先选中包含表头的数据区域，或在看板需求中明确数据源（sheet/range）。');
        }
        sources = _normalizeSources(sources);
        const primary = sources[0];
        ChatLog.info?.('BLUEPRINT', `[computeFromSources] primary: headers(${primary.headers?.length || 0})=[${(primary.headers||[]).slice(0,8).join(',')}] rows(${primary.rows?.length || 0}) row0_keys=${primary.rows?.[0] ? Object.keys(primary.rows[0]).join(',') : 'N/A'} row0=${primary.rows?.[0] ? JSON.stringify(primary.rows[0]).slice(0,300) : 'N/A'}`);
        ChatLog.info?.('BLUEPRINT', `[computeFromSources] spec.kpis=${JSON.stringify(spec.kpis||[]).slice(0,500)} spec.charts=${JSON.stringify(spec.charts||[]).slice(0,500)}`);
        let kpiSpecs = Array.isArray(spec.kpis) && spec.kpis.length ? spec.kpis : null;
        if (!kpiSpecs) {
            const nums = allNumericHeaders(primary);
            if (nums.length) {
                // 自动派生：前 2 个数值列总额 + 主数值列均值 + 记录数 + 主维度数（上限 5 个）
                kpiSpecs = [];
                nums.slice(0, 2).forEach(h => kpiSpecs.push({ label: h, column: h, func: 'sum', source: primary.id }));
                kpiSpecs.push({ label: '平均' + nums[0], column: nums[0], func: 'avg', source: primary.id });
                kpiSpecs.push({ label: '记录数', func: 'count', source: primary.id });
                const dim0 = enumerateDimensions(primary, 1)[0];
                if (dim0) {
                    const distinct = new Set((primary.rows || []).map(r => (r[dim0] !== undefined && r[dim0] !== null && String(r[dim0]) !== '') ? String(r[dim0]) : null).filter(Boolean)).size;
                    kpiSpecs.push({ label: dim0 + '数', value: distinct, source: primary.id });
                }
            } else {
                kpiSpecs = [{ label: '记录数', func: 'count', source: primary.id }];
            }
        }
        // 引擎层兜底：任何路径（含 LLM 修正绕过 normalize）进来的 KPI 规格，
        // 1) label 空/占位（含"指标 N"类偷懒标题）→ 从计算方式自动生成有意义的指标名
        //    （sum→列名、avg→平均+列名、count→记录数…；字面值 KPI 尝试按值匹配数值列 sum/avg/max/min/count 反推列名）
        // 2) 同计算（func+column 或相同字面值）重复 → 只保留第一个，杜绝重复卡片
        const inferKpiFromValue = (value, primary) => {
            const v = Number(value);
            if (!Number.isFinite(v)) return null;
            if (primary && Array.isArray(primary.rows)) {
                for (const h of allNumericHeaders(primary)) {
                    const vals = primary.rows.map(r => num(r[h])).filter(x => x !== null);
                    if (!vals.length) continue;
                    const s = vals.reduce((a, b) => a + b, 0);
                    const a = s / vals.length;
                    const mx = Math.max.apply(null, vals), mn = Math.min.apply(null, vals);
                    const near = (x, y) => Math.abs(x - y) < Math.max(1, Math.abs(y) * 0.001);
                    if (near(v, s)) return { label: h, func: 'sum', column: h };
                    if (near(v, a)) return { label: '平均' + h, func: 'avg', column: h };
                    if (near(v, mx)) return { label: '最大' + h, func: 'max', column: h };
                    if (near(v, mn)) return { label: '最小' + h, func: 'min', column: h };
                }
                if (v === primary.rows.length) return { label: '记录数', func: 'count' };
            }
            return null;
        };
        kpiSpecs = kpiSpecs.slice(0, 12).map((k, i) => {
            const c = { ...k };
            const lbl = String(c.label == null ? '' : c.label).trim();
            if (!lbl || /^<.+>$/.test(lbl) || /^指标\s*\d*$/.test(lbl)) {
                const col = String(c.column == null ? '' : c.column).trim();
                const f = String(c.func || '').toLowerCase();
                if (f === 'count') c.label = '记录数';
                else if (f === 'distinct') c.label = col ? col + '数' : '维度数';
                else if (f === 'avg') c.label = col ? '平均' + col : '平均值';
                else if (f === 'max') c.label = col ? '最大' + col : '最大值';
                else if (f === 'min') c.label = col ? '最小' + col : '最小值';
                else if (col) c.label = col;
                else if (!c.func && !c.column && c.value !== undefined && c.value !== null && c.value !== '') {
                    const infer = inferKpiFromValue(c.value, primary);
                    if (infer) { c.label = infer.label; c.func = infer.func; c.column = infer.column; }
                    else c.label = '指标 ' + (i + 1);
                }
                else c.label = '合计';
            }
            return c;
        });
        {
            const seen = new Set();
            const deduped = [];
            for (const k of kpiSpecs) {
                const hasLit = k.value !== undefined && k.value !== null && k.value !== '' && !k.func && !k.column;
                const key = hasLit ? 'lit|' + String(k.value) : String(k.func || 'sum').toLowerCase() + '|' + String(k.column || '');
                if (seen.has(key)) {
                    ChatLog.info?.('BLUEPRINT', `[computeFromSources] KPI 去重移除（同计算重复）: ${JSON.stringify(k).slice(0,120)}`);
                    continue;
                }
                seen.add(key);
                deduped.push(k);
            }
            if (deduped.length !== kpiSpecs.length) kpiSpecs = deduped;
        }
        const kpis = kpiSpecs.slice(0, 12).map(k => {
            const src = (k.source && sources.find(s => s.id === k.source)) || primary;
            const r = computeKpi(k, src);
            r._src = src; // 记录实际数据源，供 validateAndAutoFix 做降级判断
            return r;
        });
        let chartSpecs = Array.isArray(spec.charts) && spec.charts.length ? spec.charts : null;
        if (!chartSpecs) {
            // 无显式图表规划 → 枚举全部合格分类维度，每维度一张图（上限 4 张，2×2 网格）
            chartSpecs = buildAutoChartSpecs(primary, 4);
        }
        // 图表维度去重: 同 dimension+value+func 只保留第一张（兼容 dim/vals 别名）
        {
            const _seenChartDim = new Set();
            const _dedupedCharts = [];
            for (const _c of (chartSpecs || [])) {
                const _d = _c.dimension || _c.dim || '';
                const _v = _c.value || (Array.isArray(_c.vals) ? _c.vals[0] : '') || '';
                const _key = String(_d) + '|' + String(_v) + '|' + String(_c.func || 'sum');
                if (_seenChartDim.has(_key)) continue;
                _seenChartDim.add(_key);
                _dedupedCharts.push(_c);
            }
            chartSpecs = _dedupedCharts;
        }
        const charts = [];
        for (const c of (chartSpecs || []).slice(0, 10)) {
            const src = (c.source && sources.find(s => s.id === c.source)) || primary;
            const cd = aggregateChartData(c, src);
            if (cd) charts.push({ ...cd, title: c.title || '图表 ' + (charts.length + 1) });
        }
        const maxRows = spec.table && spec.table.maxRows ? spec.table.maxRows : LAYOUT.tableMaxRows;
        const tSrc = (spec.table && spec.table.source && sources.find(s => s.id === spec.table.source)) || primary;
        const table = {
            headers: tSrc.headers,
            rows: tSrc.rows.slice(0, Math.max(1, Math.min(200, maxRows))).map(r => tSrc.headers.map(h => r[h] !== undefined && r[h] !== null ? r[h] : ''))
        };
        const rawPrepared = { kpis, charts, table, primary, sources };
        return validateAndAutoFix(rawPrepared, sources, spec);
    }

    // ============================================================
    // 方向 C：确定性启发式自检 + 自动修正
    // ============================================================
    function validateAndAutoFix(prepared, sources, spec) {
        const anomalies = [];
        const primary = sources && sources[0] ? sources[0] : null;
        const numericCols = primary ? allNumericHeaders(primary) : [];
        const totalRows = primary ? (primary.rows || []).length : 0;

        const kpis = prepared.kpis || [];
        if (kpis.length === 0 && primary) {
            kpis.push(computeKpi({ label: '记录数', func: 'count' }, primary));
        }
        const allZeroCount = kpis.filter(k => k.value === 0).length;
        if (allZeroCount > 0 && allZeroCount === kpis.length && numericCols.length >= 2) {
            kpis.forEach((k, i) => {
                if (k.value === 0) {
                    const altCol = numericCols[i % numericCols.length];
                    if (altCol && altCol !== k.column) {
                        const alt = computeKpi({ label: k.label, column: altCol, func: k.func || 'sum' }, primary);
                        if (alt.value !== 0) { k.value = alt.value; k.column = altCol; k.func = alt.func; }
                    }
                }
            });
        }

        // ── KPI 降级：聚合 KPI 找不到可算数值列 → 降级为计数口径，杜绝误导性 0 ──
        if (primary) {
            kpis.forEach(k => {
                if (!k.noNumeric) return;
                const src = k._src || primary;
                const srcNumeric = allNumericHeaders(src);
                if (srcNumeric.length > 0) {
                    // 有数值列但指定列无效 → 换用第一个数值列重算
                    const altCol = srcNumeric[0];
                    const alt = computeKpi({ label: k.label, column: altCol, func: (k.func && k.func !== 'count') ? k.func : 'sum' }, src);
                    k.value = alt.value; k.func = alt.func; k.column = altCol; k.numeric = true; k.noNumeric = false; k.downgraded = true;
                    anomalies.push({ type: 'kpi_downgraded', message: `指标「${k.label}」无可算数值列，已改按「${altCol}」${alt.func} 计算` });
                } else {
                    // 数据源完全没有数值列 → 降级为记录数
                    const cnt = computeKpi({ label: k.label, func: 'count' }, src);
                    k.value = cnt.value; k.func = 'count'; k.column = undefined; k.unit = ''; k.numeric = true; k.noNumeric = false; k.downgraded = true;
                    anomalies.push({ type: 'kpi_downgraded', message: `指标「${k.label}」无数值列，已按记录数口径 ${cnt.value} 展示` });
                }
            });
        }

        // ── KPI 数量下限：至少 2 个指标，避免看板是单指标空壳 ──
        if (kpis.length < 2 && primary) {
            const dim = guessDimensionHeader(primary);
            if (dim) {
                const distinct = new Set((primary.rows || []).map(r => (r[dim] !== undefined && r[dim] !== null && String(r[dim]) !== '') ? String(r[dim]) : '未分类')).size;
                kpis.push({ label: dim + '数', unit: '', value: distinct, numeric: true, computed: true });
                anomalies.push({ type: 'kpi_added', message: `自动补充指标「${dim}数」=${distinct}` });
            } else {
                kpis.push(computeKpi({ label: '记录数', func: 'count' }, primary));
            }
        }

        const chartsIn = prepared.charts || [];
        const chartsOut = [];
        for (const cd of chartsIn) {
            // ── 最终维度字段校验：dim 不在 headers 里 → 模糊匹配或自动推断 ──
            if (primary && cd.dim && primary.headers.indexOf(cd.dim) < 0) {
                const headers = primary.headers || [];
                let best = null; let bestScore = 0;
                const tgt = cd.dim.toLowerCase();
                for (const h of headers) {
                    const hl = String(h).toLowerCase();
                    let score = 0;
                    if (hl === tgt) score = 100;
                    else if (hl.includes(tgt) || tgt.includes(hl)) score = Math.max(hl.length, tgt.length);
                    if (score > bestScore) { bestScore = score; best = h; }
                }
                const finalDim = (best && bestScore >= 2) ? best : (guessDimensionHeader(primary) || cd.dim);
                if (finalDim !== cd.dim) {
                    ChatLog.info?.('BLUEPRINT', `[validateAndAutoFix] chart.dim「${cd.dim}」→ 纠正为「${finalDim}」`);
                    cd.dim = finalDim;
                    // 维度变了，data 也得重算
                    const freshData = aggregateChartData({ dimension: finalDim, value: cd.vals && cd.vals[0], func: cd.func, chartType: cd.chartType }, primary);
                    if (freshData && freshData.data && freshData.data.length > 1) {
                        cd.data = freshData.data;
                        cd.header = freshData.header;
                    }
                }
            }
            const groups = cd.data ? cd.data.length : 0;
            const isPieType = /xlPie|xlDoughnut/.test(cd.chartType);
            if (isPieType && groups <= 2) {
                ChatLog.info?.('BLUEPRINT', `[auto-fix] 饼图只有 ${groups} 组 → 强制柱状图`);
                cd.chartType = 'xlColumnClustered';
            }
            if (groups > 20) {
                ChatLog.info?.('BLUEPRINT', `[auto-fix] 图表 ${groups} 组 → top10 + 其他`);
                const top = cd.data.slice(0, 10);
                const rest = cd.data.slice(10);
                const otherVals = rest.reduce((acc, row) => {
                    for (let i = 1; i < row.length; i++) acc[i - 1] = (acc[i - 1] || 0) + (row[i] || 0);
                    return acc;
                }, cd.vals.map(() => 0));
                const otherRow = ['其他', ...otherVals.map(v => Math.round(v * 100) / 100)];
                cd.data = [...top, otherRow];
            }
            if (groups <= 1) {
                anomalies.push({ type: 'chart_too_few_groups', message: `图表「${cd.title || cd.dim}」只有 ${groups} 组数据，跳过`, chartIndex: chartsOut.length });
                continue;
            }
            const allZero = cd.data.every(row => row.slice(1).every(v => v === 0 || v === null || v === undefined));
            if (allZero) {
                anomalies.push({ type: 'chart_all_zero', message: `图表「${cd.title || cd.dim}」数值全为 0`, chartIndex: chartsOut.length, dimension: cd.dim });
                continue;
            }
            chartsOut.push(cd);
        }
        // ── 维度冲突自修复：多张图 data 完全相同 → LLM 重复填了同一个 dimension
        //   根因：LLM 在 spec.charts 里把所有 chart 的 dimension 都写成了"客户类型"，
        //   导致 aggregateChartData 算出完全一样的 data。
        //   修复：检测 data 指纹重复 → 重新枚举正确维度分配给重复图。
        if (chartsOut.length >= 2 && primary) {
            const _dataFp = (cd) => JSON.stringify((cd.data || []).map(r => r[0] + '=' + r.slice(1).join(',')));
            const _fpMap = new Map();
            for (let i = 0; i < chartsOut.length; i++) {
                const fp = _dataFp(chartsOut[i]);
                if (!_fpMap.has(fp)) _fpMap.set(fp, []);
                _fpMap.get(fp).push(i);
            }
            const _dupFps = [..._fpMap.entries()].filter(([, idxs]) => idxs.length > 1);
            if (_dupFps.length > 0) {
                // 收集所有被重复占用的 dimension（来自唯一 data 图）
                const _usedDims = new Set(chartsOut.map(c => c.dim).filter(Boolean));
                const _allDims = enumerateDimensions(primary, 10);
                const _availDims = _allDims.filter(d => !_usedDims.has(d));
                ChatLog.info?.('BLUEPRINT', `[auto-fix] 检测到图表维度冲突: ${_dupFps.map(([fp, idxs]) => `data相同于idx=${idxs.join(',')}`).join('; ')}，可用维度=[${_availDims.join(',')}]`);
                // 对每个重复组，idxs[0] 保留原 dim，idxs[1..] 分配新维度
                for (const [, idxs] of _dupFps) {
                    for (let k = 1; k < idxs.length; k++) {
                        if (_availDims.length === 0) break;
                        const _newDim = _availDims.shift();
                        const _chart = chartsOut[idxs[k]];
                        const _src = primary;
                        const _val = firstNumericHeader(primary) || _chart.vals[0];
                        const _fixed = aggregateChartData({
                            dimension: _newDim, value: _val, func: _chart.func,
                            chartType: _chart.chartType, source: 'primary'
                        }, _src);
                        if (_fixed && _fixed.data.length > 1) {
                            chartsOut[idxs[k]] = { ..._fixed, title: _chart.title.replace(_chart.dim, _newDim) || ('各' + _newDim + '销售额') };
                            anomalies.push({ type: 'chart_dimension_fixed', message: `图表维度冲突修复: 「${_chart.title}」→ 「${chartsOut[idxs[k]].title}」` });
                            ChatLog.info?.('BLUEPRINT', `[auto-fix] 维度修复 idx=${idxs[k]}: ${_chart.dim} → ${_newDim}，新 data=${_fixed.data.length}组`);
                        }
                    }
                }
            }
        }
        // ── 维度覆盖兜底：为每个未覆盖的合格分类维度补图（上限 4 张，2×2 网格）──
        if (primary) {
            const covered = new Set(chartsOut.map(c => c.dim).filter(Boolean));
            const autoSpecs = buildAutoChartSpecs(primary, 4);
            for (const sc of autoSpecs) {
                if (chartsOut.length >= 4) break;
                if (covered.has(sc.dimension)) continue;
                const src = (sc.source && sources.find(s => s.id === sc.source)) || primary;
                const cd = aggregateChartData(sc, src);
                if (!cd || !Array.isArray(cd.data) || cd.data.length <= 1) continue;
                const allZeroAuto = cd.data.every(row => row.slice(1).every(v => v === 0 || v === null || v === undefined));
                if (allZeroAuto) continue;
                // 原图表全部失效时的首张补图属于"重新生成"（严重，触发 LLM 修正）；
                // 已有有效图时的补图属于"维度覆盖"（正常增强，仅入口径说明）
                const wasEmpty = chartsOut.length === 0;
                chartsOut.push({ ...cd, title: sc.title });
                covered.add(sc.dimension);
                anomalies.push(wasEmpty
                    ? { type: 'chart_regenerated', message: `原图表全部无效，自动生成「${sc.title}」` }
                    : { type: 'chart_dimension_added', message: `自动补充「${sc.title}」图表（维度覆盖）` });
                ChatLog.info?.('BLUEPRINT', `[auto-fix] ${wasEmpty ? '图表兜底' : '维度覆盖'} → ${sc.title}`);
            }
        }

        kpis.forEach(k => { delete k._src; });
        const preparedOut = { ...prepared, kpis, charts: chartsOut, _anomalies: anomalies };
        preparedOut._meta = { sourceRows: totalRows, sourceCols: primary ? (primary.headers || []).length : 0, kpiCount: kpis.length, chartCount: chartsOut.length, numericColCount: numericCols.length };
        return preparedOut;
    }

    function collectAnomalies(prepared, execResult) {
        const list = [];
        if (prepared && Array.isArray(prepared._anomalies)) {
            for (const a of prepared._anomalies) list.push(a);
        }
        if (execResult && Array.isArray(execResult.stepResults)) {
            for (const sr of execResult.stepResults) {
                if (sr && (sr.ok === false || sr.raw?.ok === false || sr.raw?.error)) {
                    list.push({ type: 'tool_failed', message: `步骤「${sr.tool || sr.title}」执行失败`, detail: String(sr.raw?.error || sr.error || '未知错误').slice(0, 200) });
                }
            }
        }
        return list;
    }

    // ============================================================
    // buildPlan：组合器 —— 调 excelRenderer + htmlRenderer
    // ============================================================
    function buildPlan(prepared, spec, ctx) {
        const readSteps = buildReadSteps(spec, ctx);
        const deliverable = String(spec.deliverable || 'dual').toLowerCase();
        const steps = [...readSteps];
        if (deliverable !== 'html') {
            const xl = App.ChatV2.dashboardModules.excelRenderer;
            if (xl && typeof xl.buildExcelPlan === 'function') {
                const xlSteps = xl.buildExcelPlan(prepared, spec);
                steps.push(...xlSteps);
            }
        }
        if (deliverable !== 'excel') {
            const hl = App.ChatV2.dashboardModules.htmlRenderer;
            const htmlStep = hl && typeof hl.buildHtmlStep === 'function' ? hl.buildHtmlStep(prepared, spec) : null;
            if (htmlStep) steps.push(htmlStep);
        }
        return steps;
    }

    // ============================================================
    // renderHtml：代理到 htmlRenderer.renderHtml（保持旧函数名）
    // ============================================================
    function renderHtml(prepared, spec) {
        const hl = App.ChatV2.dashboardModules.htmlRenderer;
        if (hl && typeof hl.renderHtml === 'function') return hl.renderHtml(prepared, spec);
        return '';
    }

    // ============================================================
    // summarize：最终总结
    // ============================================================
    function summarize(prepared, spec, opts) {
        const o = opts || {};
        const deliverable = String(spec.deliverable || 'dual').toLowerCase();
        const excelOnly = deliverable === 'excel';
        const htmlOnly = deliverable === 'html';
        const lines = [];
        const targets = [];
        if (!htmlOnly) targets.push(`Excel 内看板${spec.sheet ? '（工作表「' + spec.sheet + '」）' : ''}`);
        if (!excelOnly) targets.push('HTML 大屏');
        lines.push(`- 看板「${spec.title || '数据看板'}」已生成 → ${targets.join(' + ')}`);
        lines.push(`- 主题「${themeOf(spec).name}」 · KPI ${(prepared.kpis || []).length} 个 · 图表 ${(prepared.charts || []).length} 张 · 明细表 ${(prepared.table && prepared.table.rows ? prepared.table.rows.length : 0)} 行`);
        const anoms = (prepared && Array.isArray(prepared._anomalies)) ? prepared._anomalies : [];
        const notes = anoms.filter(a => a.type === 'kpi_downgraded' || a.type === 'chart_regenerated' || a.type === 'kpi_added' || a.type === 'chart_dimension_added');
        if (notes.length) {
            lines.push(`- 口径说明：${notes.map(n => n.message).slice(0, 5).join('；')}`);
        }
        if (htmlOnly) {
            if (o.htmlError) lines.push(`- ⚠️ HTML 大屏打开失败：${o.htmlError}（可稍后手动访问）`);
            else lines.push(`- HTML 大屏已在浏览器打开（支持缩放全屏，适合投屏汇报）`);
        } else if (!excelOnly) {
            if (o.htmlError) lines.push(`- ⚠️ HTML 大屏打开失败：${o.htmlError}（Excel 内看板已生成，可稍后手动刷新浏览器）`);
            else lines.push(`- HTML 大屏已在浏览器打开（支持缩放全屏，适合投屏汇报）`);
        }
        return lines.join('\n');
    }

    // ============================================================
    // 导出（保持所有旧导出名 + 删除已废弃的 buildRenderSteps）
    // ============================================================
    App.ChatV2.dashboardModules.blueprintEngine = {
        THEMES, DEFAULT_THEME, themeOf, resolveThemeFromMessage, isThemeChangeIntent, hexToRgb, rgbToHex,
        LAYOUT, colLetter,
        buildSkeletonSteps, buildReadSteps,
        buildPlan, computeFromSources,
        aggregateChartData, firstNumericHeader,
        renderHtml, summarize,
        humanizeChartType, isCategoryChart, chartFamily,
        enumerateDimensions, buildAutoChartSpecs, recommendChartType,
        validateAndAutoFix, collectAnomalies,
        _normalizeSources,
        computeTopNRanking, computeGaugeData, computeKpiTrend
    };
    ChatLog.info?.('CHAT-V2', '✅ dashboardBlueprintEngine loaded (方案 B · 纯计算层 + 组合器, ' + Object.keys(THEMES).length + ' themes)');
})();
