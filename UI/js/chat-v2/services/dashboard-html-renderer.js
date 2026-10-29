/**
 * ============================================================
 * dashboard-html-renderer.js — HTML 看板大屏渲染（方案 B · 渲染层分离）
 * ============================================================
 * 职责：把 compute 层产出的 prepared{ kpis, charts, table, primary } 转成完整 HTML 大屏。
 *       只管浏览器里长啥样，不碰 Excel COM 原子工具，不碰数据计算。
 *
 * 交互（阶段3）：HTML 大屏内置 JS 筛选器（零后端依赖）。
 *   - 嵌入原始明细数据（prepared.primary，截断保护）+ 各可筛选维度
 *   - 点击维度值 chip → 全局过滤 → 重算 KPI / 图表 / 明细表 → 局部重绘
 *   - 重算纯函数与 ECharts option 构造函数通过 toString() 内联进页面，
 *     与 node 端共用同一份实现，逻辑不漂移
 *
 * 依赖：themeOf 从 dashboard-blueprint-engine 共享引入（通过 App.ChatV2.dashboardModules.blueprintEngine）
 * 挂载：App.ChatV2.dashboardModules.htmlRenderer
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.dashboardModules) App.ChatV2.dashboardModules = {};

    const ChatLog = window.ChatLog || console;
    const MAX_FILTER_ROWS = 5000; // 前端筛选基于的原始行上限（防 HTML 过大）

    function themeOf(spec) {
        const eng = App.ChatV2.dashboardModules.blueprintEngine || {};
        if (typeof eng.themeOf === 'function') return eng.themeOf(spec);
        const t = (eng.THEMES && eng.THEMES[(spec && spec.theme) || 'business_blue']) || null;
        if (t) return t;
        return { name: '默认', html: { accent: '#4DA3FF', bg: '#0F1624', panel: '#192235', panelBorder: '#2B3A55', title: '#E6E9F2', sub: '#8B93A7', grid: '#2B3A55', kpiLabel: '#8B93A7', kpiValue: '#E6E9F2' } };
    }

    function esc(s) {
        return String(s === undefined || s === null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    function fmtNum(v) {
        const n = Number(v);
        if (!Number.isFinite(n)) return String(v);
        return new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 }).format(n);
    }

    /**
     * 看板标题净化：LLM/兜底链路常把用户原话（"帮我生成excel大看板"）直接当标题。
     * 剥掉命令动词与泛指词，剩余有效内容太短则回退到数据驱动标题（如"销售额数据分析"）。
     */
    function sanitizeBoardTitle(raw, fallback) {
        var s = String(raw == null ? '' : raw).trim();
        if (s) {
            var cleaned = s
                .replace(/帮我|帮忙|请|麻烦|给我|替我|帮我一下/g, '')
                .replace(/生成|制作|创建|搭建|搭|弄|画|输出|做一|做个|做份|一份|一个|一张|一下|做|出|排|看看|瞧瞧/g, '')
                .replace(/excel|Excel|EXCEL|表格|dashboard|Dashboard/g, '')
                .replace(/数据面板|驾驶舱|大看板|看板|大屏|仪表盘|的|了|吧|呗/g, '')
                .replace(/[\s，。,.\-—_·、!！?？]+/g, '');
            // 剥完只剩"数据/图表/报表/分析"这类无信息量残留 → 视为无效，走回退
            if (cleaned.length >= 2 && !/^(数据|图表|报表|分析)+$/.test(cleaned)) return cleaned.slice(0, 30);
        }
        return String(fallback || '数据看板');
    }

    /** 数据驱动回退标题：首个数值列名 + "数据分析"（如"销售额数据分析"），无数值列则"数据看板" */
    function dashboardFallbackTitle(prepared) {
        try {
            var p = prepared && prepared.primary;
            var headers = p && p.headers;
            var rows = p && p.rows;
            if (headers && headers.length && rows && rows.length) {
                for (var i = 0; i < headers.length; i++) {
                    var h = String(headers[i]);
                    if (!h) continue;
                    var hasNum = rows.some(function (r) { return Number.isFinite(Number(r[h])) && r[h] !== '' && r[h] != null; });
                    if (hasNum) return h + '数据分析';
                }
            }
        } catch (e) { /* 忽略，走默认 */ }
        return '数据看板';
    }

    function echartsOptionFor(cd, theme) {
        const accent = theme.html.accent;
        const textColor = theme.html.sub;
        const gridColor = theme.html.grid;
        const palette = [accent, '#43D17C', '#FF8A3D', '#8B7CF6', '#F5C044', '#4FD8C6', '#F2728F', '#A78BFA'];
        const cats = cd.data.map(r => String(r[0]));
        const chartType = String(cd.chartType || '').toLowerCase();
        const isPie = /xlpie/.test(chartType);
        const isDoughnut = /xldoughnut/.test(chartType);
        const isBarH = /xlbar/.test(chartType);

        // ── 漏斗图 ──
        if (/xlfunnel/.test(chartType)) {
            return {
                tooltip: { trigger: 'item', formatter: '{b}: {c}' },
                series: [{
                    type: 'funnel',
                    left: '10%', top: 10, bottom: 10, width: '80%',
                    maxSize: '100%', minSize: '10%',
                    sort: 'descending', gap: 2,
                    label: { color: '#E6E9F2', fontSize: 11, formatter: '{b}' },
                    labelLine: { lineStyle: { color: '#5A6B85' } },
                    itemStyle: { borderColor: gridColor, borderWidth: 1, borderRadius: 2 },
                    data: cd.data.map((r, i) => ({ name: String(r[0]), value: r[1], itemStyle: { color: palette[i % palette.length] } }))
                }]
            };
        }

        // ── 瀑布图 ──
        if (/xlwaterfall/.test(chartType)) {
            let base = 0;
            const data = cd.data.map((r, i) => {
                const v = r[1];
                const isFirst = i === 0;
                const isLast = i === cd.data.length - 1;
                if (isFirst) { base = 0; return { name: String(r[0]), value: v, itemStyle: { color: accent } }; }
                if (isLast) { return { name: String(r[0]), value: v, itemStyle: { color: '#43D17C' } }; }
                const rawValue = v;
                const up = v >= 0;
                base += v;
                return { name: String(r[0]), value: rawValue, itemStyle: { color: up ? '#43D17C' : '#F2728F' } };
            });
            // ECharts 瀑布图用 series 的 type:'bar' + stack 实现
            return {
                tooltip: { trigger: 'axis' },
                grid: { left: 40, right: 20, top: 20, bottom: 30 },
                xAxis: { type: 'category', data: cats, axisLabel: { color: textColor, fontSize: 10 }, axisLine: { lineStyle: { color: gridColor } } },
                yAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                series: [
                    { name: '占位', type: 'bar', stack: 'waterfall', itemStyle: { color: 'transparent', borderColor: 'transparent' }, data: cd.data.map((_, i) => { if (i === 0 || i === cd.data.length - 1) return 0; const v = cd.data[i-1] ? Number(cd.data[i-1][1]) : 0; return v; }) },
                    { name: '值', type: 'bar', stack: 'waterfall', data: data.map(d => d.value), itemStyle: { color: (p) => data[p.dataIndex].itemStyle.color }, label: { show: true, position: 'top', color: '#E6E9F2', fontSize: 10 } }
                ]
            };
        }

        // ── 雷达图 ──
        if (/xlradar/.test(chartType)) {
            const values = cd.vals || [];
            const series = values.map((v, i) => ({
                name: String(v),
                type: 'radar',
                smooth: /xlradarfilled/.test(chartType),
                areaStyle: /xlradarfilled/.test(chartType) ? { opacity: 0.2, color: palette[i % palette.length] } : undefined,
                lineStyle: { color: palette[i % palette.length], width: 2 },
                itemStyle: { color: palette[i % palette.length] },
                data: [{ value: cd.data.map(r => Number(r[i + 1]) || 0), name: String(v) }]
            }));
            // 计算各维度最大值做 indicator
            const maxVals = cats.map((_, ci) => {
                let m = 0;
                cd.data.forEach(r => { for (let vi = 1; vi < r.length; vi++) { const n = Number(r[vi]); if (n > m) m = n; } });
                return Math.max(m * 1.2, 1);
            });
            return {
                tooltip: {},
                legend: { top: 0, textStyle: { color: textColor, fontSize: 11 } },
                radar: {
                    indicator: cats.map((c, i) => ({ name: c, max: maxVals[i] || 100 })),
                    axisName: { color: textColor, fontSize: 10 },
                    splitLine: { lineStyle: { color: gridColor } },
                    splitArea: { areaStyle: { color: ['rgba(255,255,255,.02)', 'rgba(255,255,255,.05)'] } },
                    axisLine: { lineStyle: { color: gridColor } }
                },
                series
            };
        }

        // ── 散点图 ──
        if (/xlxyscatter|xlbubble/.test(chartType)) {
            const isBubble = /xlbubble/.test(chartType);
            return {
                tooltip: { trigger: 'item', formatter: isBubble ? '{b}<br/>X: {c[0]}<br/>Y: {c[1]}<br/>大小: {c[2]}' : '{b}<br/>X: {c[0]}<br/>Y: {c[1]}' },
                grid: { left: 40, right: 20, top: 20, bottom: 30 },
                xAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                yAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                series: [{
                    name: cd.vals ? String(cd.vals[1] || cd.vals[0]) : '数据点',
                    type: isBubble ? 'scatter' : 'scatter',
                    symbolSize: isBubble ? function (d) { return Math.max(6, Math.min(30, Number(d[2]) || 10)); } : 12,
                    itemStyle: { color: accent, borderRadius: 2 },
                    data: cd.data.map(r => isBubble ? [Number(r[0]) || 0, Number(r[1]) || 0, Number(r[2]) || 10] : [Number(r[0]) || 0, Number(r[1]) || 0])
                }]
            };
        }

        // ── 旭日图 ──
        if (/xlsunburst/.test(chartType)) {
            return {
                tooltip: { formatter: '{b}: {c}' },
                series: [{
                    type: 'sunburst',
                    radius: ['15%', '85%'],
                    center: ['50%', '50%'],
                    itemStyle: { borderColor: gridColor, borderWidth: 1 },
                    label: { color: '#E6E9F2', fontSize: 10 },
                    data: cd.data.map((r, i) => ({ name: String(r[0]), value: Number(r[1]) || 0, itemStyle: { color: palette[i % palette.length] } }))
                }]
            };
        }

        // ── 树状图 ──
        if (/xltreemap/.test(chartType)) {
            return {
                tooltip: { formatter: '{b}: {c}' },
                series: [{
                    type: 'treemap',
                    width: '90%', height: '85%', top: '5%', left: '5%',
                    itemStyle: { borderColor: gridColor, borderWidth: 1 },
                    label: { color: '#E6E9F2', fontSize: 11 },
                    data: cd.data.map((r, i) => ({ name: String(r[0]), value: Number(r[1]) || 0, itemStyle: { color: palette[i % palette.length] } }))
                }]
            };
        }

        // ── 直方图 ──
        if (/xlhistogram/.test(chartType)) {
            return {
                tooltip: { trigger: 'axis' },
                grid: { left: 40, right: 20, top: 20, bottom: 30 },
                xAxis: { type: 'category', data: cats, axisLabel: { color: textColor, fontSize: 11 }, axisLine: { lineStyle: { color: gridColor } } },
                yAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                series: [{
                    type: 'bar',
                    barWidth: '80%',
                    itemStyle: { color: accent, borderRadius: [3, 3, 0, 0] },
                    data: cd.data.map(r => Number(r[1]) || 0)
                }]
            };
        }

        // ── 饼图 / 环形图（原有逻辑增强）──
        if (isPie || isDoughnut) {
            return {
                tooltip: { trigger: 'item', formatter: '{b}<br/>{c}（{d}%）' },
                legend: { orient: 'horizontal', bottom: 0, textStyle: { color: textColor, fontSize: 11 } },
                color: palette,
                series: [{
                    type: 'pie',
                    radius: isDoughnut ? ['42%', '68%'] : '68%',
                    center: ['50%', '46%'],
                    itemStyle: { borderColor: gridColor, borderWidth: 1 },
                    data: cd.data.map((r, i) => ({ name: String(r[0]), value: Number(r[1]) || 0, itemStyle: { color: palette[i % palette.length] } })),
                    label: { color: '#E6E9F2', fontSize: 11, formatter: '{b}\n{d}%' },
                    labelLine: { lineStyle: { color: '#5A6B85' } }
                }]
            };
        }

        // ── 条形图（增强：圆角 + 渐变色）──
        if (isBarH) {
            const series = cd.vals.map((v, i) => ({
                name: String(v),
                type: 'bar',
                itemStyle: { color: palette[i % palette.length], borderRadius: [0, 3, 3, 0] },
                data: cd.data.map(r => Number(r[i + 1]) || 0)
            }));
            return {
                tooltip: { trigger: 'axis' },
                grid: { left: 8, right: 24, top: 24, bottom: 8, containLabel: true },
                legend: { top: 0, textStyle: { color: textColor, fontSize: 11 } },
                xAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                yAxis: { type: 'category', data: cats.slice().reverse(), axisLabel: { color: textColor, fontSize: 11 }, axisLine: { lineStyle: { color: gridColor } } },
                series: series.map(s => ({ ...s, data: s.data.slice().reverse() }))
            };
        }

        // ── 柱状图（xlColumn*：竖向柱，分类对比默认形态）──
        // 未知 chartType 也兜底为柱状图，避免分类数据被误画成折线
        if (/xlcolumn|xl3dcolumn/.test(chartType) || !/xlline|xlarea|xlstock|xlxy|xlpie|xldoughnut|xlbar|xlhistogram|xlradar|xlsxscatter|xlbubble|xlsunburst|xltreemap|xlwaterfall|xlfunnel/.test(chartType)) {
            const series = cd.vals.map((v, i) => ({
                name: String(v),
                type: 'bar',
                barWidth: '60%',
                itemStyle: { color: palette[i % palette.length], borderRadius: [3, 3, 0, 0] },
                data: cd.data.map(r => Number(r[i + 1]) || 0)
            }));
            return {
                tooltip: { trigger: 'axis' },
                grid: { left: 8, right: 24, top: 30, bottom: 8, containLabel: true },
                legend: { top: 0, textStyle: { color: textColor, fontSize: 11 } },
                xAxis: { type: 'category', data: cats, axisLabel: { color: textColor, fontSize: 11 }, axisLine: { lineStyle: { color: gridColor } } },
                yAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
                series: series
            };
        }

        // ── 折线图 / 面积图（仅时间/趋势类：xlLine、xlArea）──
        const area = /xlArea/.test(chartType);
        const lineSeries = cd.vals.map((v, i) => ({
            name: String(v),
            type: 'line',
            smooth: true,
            areaStyle: area ? { opacity: 0.2, color: palette[i % palette.length] } : undefined,
            lineStyle: { width: 2.5, color: palette[i % palette.length] },
            itemStyle: { color: palette[i % palette.length] },
            data: cd.data.map(r => Number(r[i + 1]) || 0)
        }));
        return {
            tooltip: { trigger: 'axis' },
            legend: { top: 0, textStyle: { color: textColor, fontSize: 11 } },
            grid: { left: 8, right: 24, top: 30, bottom: 8, containLabel: true },
            xAxis: { type: 'category', data: cats, axisLabel: { color: textColor, fontSize: 11, interval: Math.max(0, Math.floor(cats.length / 8) - 1) }, axisLine: { lineStyle: { color: gridColor } } },
            yAxis: { type: 'value', axisLabel: { color: textColor, fontSize: 11 }, splitLine: { lineStyle: { color: gridColor } } },
            series: lineSeries
        };
    }

    // ════════════════════════════════════════════════════════════════
    // 前端筛选重算纯函数（浏览器内联 + node 测试共用）
    // ════════════════════════════════════════════════════════════════

    /** 按当前筛选条件过滤原始行（filters: {维度名: 值数组}，同维度多值取 OR，跨维度 AND；空数组 = 不限） */
    function dashFilterRows(rows, filters) {
        const keys = Object.keys(filters || {});
        if (!keys.length) return rows;
        return rows.filter(r => {
            for (const d of keys) {
                const sel = filters[d];
                if (!Array.isArray(sel) || sel.length === 0) continue;
                if (sel.indexOf(String(r[d] ?? '')) === -1) return false;
            }
            return true;
        });
    }

    /** 列去重值（按出现频率降序） */
    function dashDistinct(rows, col) {
        const m = new Map();
        for (const r of rows) {
            const v = r[col] == null || r[col] === '' ? '未分类' : String(r[col]);
            m.set(v, (m.get(v) || 0) + 1);
        }
        return [...m.entries()].sort((a, b) => b[1] - a[1]).map(e => e[0]);
    }

    /** 聚合数值列（sum/avg/max/min），返回 number 或 null */
    function dashAggNum(rows, col, func) {
        let sum = 0, cnt = 0, max = -Infinity, min = Infinity;
        for (const r of rows) {
            const raw = r[col];
            // 与引擎 num() 一致：空值不计入，避免空单元格当 0 拉低 avg
            if (raw === null || raw === undefined || raw === '') continue;
            const v = Number(raw);
            if (!Number.isFinite(v)) continue;
            sum += v; cnt++;
            if (v > max) max = v;
            if (v < min) min = v;
        }
        if (cnt === 0) return null;
        if (func === 'avg') return Math.round((sum / cnt) * 100) / 100;
        if (func === 'max') return max;
        if (func === 'min') return min;
        return Math.round(sum * 100) / 100;
    }

    /** 把"2025-01" / "2025/1/5" / "2025年1月" / "1月" / "1月5日" 解析为可比较时间序号，不可解析返回 NaN */
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
    /** 判断分组 key 是否为时间/有序类别（>=60% 可解析为日期） */
    function looksLikeTimeDim(keys) {
        if (!keys || keys.length < 2) return false;
        let hit = 0;
        for (const k of keys) { if (!Number.isNaN(parseTimeKey(k))) hit++; }
        return hit / keys.length >= 0.6;
    }

    /** 重算 KPI 展示数组（返回与初始 kpis 等长的 {label,unit,value}） */
    function dashRecalcKpis(kpis, rows, dims) {
        return kpis.map(k => {
            const func = String(k.func || '').toLowerCase();
            if (func === 'count') {
                return { label: k.label, unit: k.unit || '', value: rows.length, numeric: true };
            }
            if (func === 'distinct' && k.column) {
                return { label: k.label, unit: k.unit || '', value: new Set(rows.map(r => String(r[k.column] ?? '未分类'))).size, numeric: true };
            }
            if (k.column) {
                const v = dashAggNum(rows, k.column, func || 'sum');
                return { label: k.label, unit: k.unit || '', value: v == null ? 0 : v, numeric: v != null };
            }
            // 字面值 KPI（引擎自动补充的「维度数」类）：按 label 以「维度名+数」结尾重算 distinct
            const m = String(k.label || '').match(/^(.+?)数$/);
            if (m && dims.includes(m[1])) {
                return { label: k.label, unit: k.unit || '', value: new Set(rows.map(r => String(r[m[1]] ?? '未分类'))).size, numeric: true };
            }
            return { label: k.label, unit: k.unit || '', value: k.value, numeric: k.numeric };
        });
    }

    /** 图表聚合：按维度分组聚合（func=count 时各组值=行数），返回 {header, data} */
    function dashChartData(rows, dim, vals, func) {
        // count 型只画一个系列（行数），多列会导致完全重叠的柱子
        if (func === 'count') vals = vals.slice(0, 1);
        const groups = new Map();
        for (const r of rows) {
            const key = (r[dim] == null || r[dim] === '') ? '未分类' : String(r[dim]);
            if (!groups.has(key)) groups.set(key, { vals: vals.map(() => 0), count: 0 });
            const g = groups.get(key);
            g.count += 1;
            vals.forEach((c, i) => {
                const v = Number(r[c]);
                if (Number.isFinite(v)) g.vals[i] += v;
            });
        }
        let entries = [...groups.entries()].map(([k, g]) => {
            if (func === 'count') g.vals = g.vals.map(() => g.count);
            return [k, g.vals.map(v => Math.round(v * 100) / 100)];
        });
        // 时间维度（月份/日期）保持时间升序，不按值降序打乱时间轴；非时间维度按首列值降序
        const _keys = entries.map(e => e[0]);
        if (looksLikeTimeDim(_keys)) {
            entries.sort((a, b) => parseTimeKey(a[0]) - parseTimeKey(b[0]));
        } else {
            entries.sort((a, b) => b[1][0] - a[1][0]);
        }
        // >20 组 → top10 + 其他（与引擎规则一致）
        if (entries.length > 20) {
            const top = entries.slice(0, 10);
            const rest = entries.slice(10);
            const otherVals = vals.map((_, i) => Math.round(rest.reduce((s, e) => s + (e[1][i] || 0), 0) * 100) / 100);
            entries = [...top, ['其他', otherVals]];
        }
        // 与引擎 aggregateChartData 一致：行 = [key, v1, v2, ...]（数值展开，ECharts series 直接取 r[i+1]）
        return { header: [dim, ...vals], data: entries.map(([k, vals]) => [k, ...vals]) };
    }

    // ════════════════════════════════════════════════════════════════
    // renderHtml：生成完整大屏 HTML（含筛选器）
    // ════════════════════════════════════════════════════════════════
    function renderHtml(prepared, spec) {
        if (spec) spec.title = sanitizeBoardTitle(spec.title, dashboardFallbackTitle(prepared));
        const theme = themeOf(spec);
        const t = theme.html;
        // ← 补上 gauge 组件需要的变量（之前只在 echartsOptionFor 里有）
        const gridColor = t.grid;
        const textColor = t.sub;
        const kpis = prepared.kpis || [];
        const charts = prepared.charts || [];
        const table = prepared.table || null;
        const primary = prepared.primary || null;
        const dims = primary ? (function () {
            const eng = App.ChatV2.dashboardModules.blueprintEngine || {};
            if (typeof eng.enumerateDimensions === 'function') {
                return eng.enumerateDimensions(primary, 6);
            }
            // 兜底：本地简化枚举（文本列、非 ID/日期、唯一值 2~24）
            const headers = primary.headers || [];
            const rows = primary.rows || [];
            const isId = h => /(^|[^a-z0-9])id($|[^a-z0-9])/i.test(h) || /编号|单号|序号|编码|流水号|单id|id号/i.test(h);
            const isDate = h => /日期|时间|date|time|year|month|day/i.test(h);
            return headers.filter(h => !isId(h) && !isDate(h) && !rows.some(r => Number.isFinite(Number(r[h]))) && new Set(rows.map(r => String(r[h] ?? '未分类'))).size >= 2);
        })() : [];

        // ── 筛选器 UI：每维度一个 chip 组 ──
        let filterHtml = '';
        if (dims.length && primary) {
            const groups = dims.map(d => {
                const vals = dashDistinct(primary.rows || [], d);
                const chips = ['<button type="button" class="f-chip active" data-dim="' + esc(d) + '" data-val="">全部</button>']
                    .concat(vals.map(v => '<button type="button" class="f-chip" data-dim="' + esc(d) + '" data-val="' + esc(v) + '">' + esc(v) + '</button>'))
                    .join('');
                return '<div class="f-group"><span class="f-label">' + esc(d) + '</span><span class="f-chips">' + chips + '</span></div>';
            }).join('');
            filterHtml = '<div class="panel filter-panel"><div class="panel-title">筛选器<span class="table-count">（同一维度可多选）</span></div>' + groups + '</div>';
        }

        const chartCards = charts.map((cd, i) => {
            const opt = JSON.stringify(echartsOptionFor(cd, theme)).replace(/</g, '\\u003c').replace(/'/g, '&#39;');
            return `<div class="chart-card"><div class="chart-title">${esc(cd.title || '图表 ' + (i + 1))}</div><div class="chart" id="chart-${i}" data-opt='${opt}'></div></div>`;
        }).join('\n');

        // ── 排行榜（TopN）组件：从引擎调用 computeTopNRanking 自动生成 ──
        let rankHtml = '';
        const _engTopN = App.ChatV2.dashboardModules.blueprintEngine || {};
        if (typeof _engTopN.computeTopNRanking === 'function' && primary && primary.headers) {
            const _rankDim = dims[0] || (function () {
                const h = primary.headers || [];
                return h.find(x => !/日期|时间|date|time|编号|ID|单号/i.test(x) && !rows.some(r => Number.isFinite(Number(r[x]))));
            })();
            const _rankVal = _engTopN.firstNumericHeader ? _engTopN.firstNumericHeader(primary) : null;
            if (_rankDim && _rankVal) {
                const rank = _engTopN.computeTopNRanking(primary, _rankDim, _rankVal, 'sum', 5);
                if (rank && rank.items && rank.items.length) {
                    const _rankColors = ['#F5C044', '#F5C044', '#F5C044', '#8B93A7', '#8B93A7'];
                    rankHtml = `<div class="panel rank-panel"><div class="panel-title">🏆 ${rank.topN} Top ${esc(_rankDim)}排行</div><div class="rank-list">${rank.items.map((item, i) => {
                        const pct = rank.total > 0 ? Math.round(item.value / rank.total * 100) : 0;
                        return `<div class="rank-item"><span class="rank-num" style="color:${_rankColors[i] || '#8B93A7'}">${i + 1}</span><span class="rank-name">${esc(item.name)}</span><div class="rank-bar-wrap"><div class="rank-bar" style="width:${pct}%;background:${_rankColors[i] || '#8B93A7'}"></div></div><span class="rank-val">${esc(fmtNum(item.value))}</span></div>`;
                    }).join('')}</div></div>`;
                }
            }
        }

        // ── 仪表盘组件：完成率/达标率展示 ──
        let gaugeHtml = '';
        if (kpis && kpis.length >= 2 && typeof _engTopN.computeGaugeData === 'function') {
            // 用 KPI 里最后两个做完成率（第一个 KPI / 假设的目标 = 仪表盘）
            const kpiSum = kpis.find(k => String(k.func || '').toLowerCase() === 'sum' && k.numeric);
            if (kpiSum && Number.isFinite(Number(kpiSum.value))) {
                // 目标值 = 当前值 × 1.2（假设增长20%为目标，或用KPI自身×1.3做目标）
                const target = Math.round(Number(kpiSum.value) * 1.3);
                const gd = _engTopN.computeGaugeData(Number(kpiSum.value), target, (kpiSum.label || '目标完成率') + ' 达成度');
                const gaugeOpt = JSON.stringify({
                    series: [{
                        type: 'gauge', startAngle: 200, endAngle: -20,
                        min: 0, max: 100,
                        radius: '90%', center: ['50%', '55%'],
                        progress: { show: true, width: 14, itemStyle: { color: gd.color } },
                        axisLine: { lineStyle: { width: 14, color: [[1, gridColor]] } },
                        axisTick: { show: false }, splitLine: { show: false }, axisLabel: { show: false },
                        pointer: { show: false },
                        anchor: { show: false },
                        title: { show: true, offsetCenter: [0, '-15%'], color: textColor, fontSize: 11 },
                        detail: { show: true, offsetCenter: [0, '20%'], color: gd.color, fontSize: 24, fontWeight: 'bold', formatter: '{value}%' },
                        data: [{ value: gd.percent, name: gd.label }]
                    }]
                }).replace(/</g, '\\u003c').replace(/'/g, '&#39;');
                gaugeHtml = `<div class="panel gauge-panel"><div class="chart" id="gauge-0" data-opt='${gaugeOpt}' style="height:160px"></div></div>`;
            }
        }

        // ── 实时时钟组件（大屏装饰）──
        const clockHtml = `<div class="clock-widget"><div class="clock-date" id="clock-date">——</div><div class="clock-time" id="clock-time">--:--:--</div></div>`;

        const kpiCards = kpis.map((k, i) => `
            <div class="kpi-card" style="animation-delay:${i * 60}ms">
                <div class="kpi-label">${esc(k.label)}${k.unit ? '<span class="kpi-unit">（' + esc(k.unit) + '）</span>' : ''}</div>
                <div class="kpi-value">${esc(fmtNum(k.value))}</div>
            </div>`).join('\n');
        // 初始明细行数 = 引擎实际渲染的行数（引擎已按 LAYOUT/spec 截断），作为行数按钮的默认高亮
        const tableMaxRows = (table && table.rows) ? table.rows.length : 20;
        const tableLimitBtns = [20, 50, 100, 0].map(n => `<button type="button" class="limit-btn${tableMaxRows === n ? ' active' : ''}" data-n="${n}">${n === 0 ? '全部' : n}</button>`).join('');
        const tableHtml = table && table.headers.length
            ? `<div class="panel table-panel">
                <div class="panel-title">数据明细 <span class="table-count" id="table-count"></span><span class="table-limit">${tableLimitBtns}</span></div>
                <div class="table-wrap"><table><thead><tr>${table.headers.map(h => `<th>${esc(h)}</th>`).join('')}</tr></thead>
                <tbody id="table-body">${table.rows.map(r => `<tr>${r.map(c => `<td>${esc(fmtNum(c))}</td>`).join('')}</tr>`).join('')}</tbody></table></div>
               </div>`
            : '';

        // ── 嵌入初始数据（浏览器端筛选重算的原料）──
        const rawRows = primary && Array.isArray(primary.rows) ? primary.rows : null;
        const initRows = rawRows ? rawRows.slice(0, MAX_FILTER_ROWS) : [];
        const truncated = rawRows ? rawRows.length > MAX_FILTER_ROWS : false;
        const init = {
            headers: primary && primary.headers ? primary.headers : [],
            rows: initRows,
            truncated,
            dims,
            theme: theme,
            kpis: kpis.map(k => ({ label: k.label, unit: k.unit || '', func: k.func, column: k.column, value: k.value, numeric: !!k.numeric })),
            charts: charts.map(cd => ({ title: cd.title || '', chartType: cd.chartType, dim: cd.dim, vals: cd.vals || [], func: cd.func || 'sum' })),
            tableHeaders: table && table.headers ? table.headers : [],
            tableMaxRows: (spec.table && spec.table.maxRows) ? spec.table.maxRows : 20
        };
        const initJson = JSON.stringify(init).replace(/</g, '\\u003c');

        // ── 内联源码（node 端函数 toString，浏览器与 node 共用同一实现）──
        const pureSrc = [
            dashFilterRows.toString(),
            dashDistinct.toString(),
            dashAggNum.toString(),
            parseTimeKey.toString(),
            looksLikeTimeDim.toString(),
            dashRecalcKpis.toString(),
            dashChartData.toString(),
            echartsOptionFor.toString()
        ].join('\n\n');

        const bootSrc = `(function () {
    var INIT = window.__DASH_INIT__ || { rows: [], dims: [], kpis: [], charts: [], theme: null, tableHeaders: [], tableMaxRows: 20 };
    var filters = {};
    var tableLimit = INIT.tableMaxRows || 20;
    var instances = [];
    var fmt = function (v) {
        var n = Number(v);
        if (!isFinite(n)) return String(v);
        return new Intl.NumberFormat('zh-CN', { maximumFractionDigits: 2 }).format(n);
    };
    function initCharts() {
        var els = document.querySelectorAll('.chart');
        Array.prototype.forEach.call(els, function (el) {
            var opt = null;
            try { opt = JSON.parse(el.getAttribute('data-opt')); } catch (e) { ChatLog.debug?.('UI', '[dashboard] 图表 data-opt 非 JSON，跳过初始化:', e); }
            if (opt && window.echarts) {
                var idx = String(el.id).replace('chart-', '');
                instances[idx] = echarts.init(el);
                instances[idx].setOption(opt);
            } else {
                el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#8B93A7;font-size:12px;">离线模式：图表无法加载（ECharts 需要联网）</div>';
            }
        });
    }
    function renderKpis() {
        var rows = dashFilterRows(INIT.rows, filters);
        var ks = dashRecalcKpis(INIT.kpis, rows, INIT.dims);
        var cards = document.querySelectorAll('.kpi-card');
        for (var i = 0; i < cards.length; i++) {
            var v = cards[i].querySelector('.kpi-value');
            if (v && ks[i]) v.textContent = fmt(ks[i].value);
        }
        return rows;
    }
    function renderCharts(rows) {
        for (var i = 0; i < INIT.charts.length; i++) {
            var cd = INIT.charts[i];
            var el = document.getElementById('chart-' + i);
            if (!el) continue;
            var res = dashChartData(rows, cd.dim, cd.vals, cd.func);
            if (!res.data.length) {
                if (instances[i]) { instances[i].dispose(); instances[i] = null; }
                el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#8B93A7;font-size:12px;">当前筛选下无数据</div>';
                continue;
            }
            var cd2 = { chartType: cd.chartType, dim: cd.dim, vals: cd.vals, data: res.data };
            var opt = echartsOptionFor(cd2, INIT.theme);
            if (!instances[i]) {
                if (!window.echarts) {
                    el.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;color:#8B93A7;font-size:12px;">离线模式：图表无法加载（ECharts 需要联网）</div>';
                    continue;
                }
                instances[i] = echarts.init(el);
            }
            instances[i].setOption(opt, true);
        }
    }
    function renderTable(rows) {
        var body = document.getElementById('table-body');
        if (!body) return;
        var show = tableLimit > 0 ? rows.slice(0, tableLimit) : rows;
        var h = INIT.tableHeaders;
        body.innerHTML = show.map(function (r) {
            return '<tr>' + h.map(function (c) {
                var v = r[c] == null ? '' : r[c];
                var txt;
                // 空单元格显示空白，不是 0（Number('')=0 会把空格渲染成"0"）
                if (v === '' || v === null || v === undefined) { txt = ''; }
                else { var n = Number(v); txt = isFinite(n) ? fmt(n) : String(v); }
                return '<td>' + String(txt).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;') + '</td>';
            }).join('') + '</tr>';
        }).join('');
        var cnt = document.getElementById('table-count');
        if (cnt) cnt.textContent = tableLimit > 0
            ? '（当前筛选 ' + rows.length + ' 行，显示前 ' + show.length + ' 行）'
            : '（当前筛选 ' + rows.length + ' 行，全部显示）';
    }
    function bindTableLimit() {
        var btns = document.querySelectorAll('.limit-btn');
        Array.prototype.forEach.call(btns, function (b) {
            b.addEventListener('click', function () {
                tableLimit = parseInt(b.getAttribute('data-n'), 10) || 0;
                var all = document.querySelectorAll('.limit-btn');
                Array.prototype.forEach.call(all, function (x) { x.classList.toggle('active', x === b); });
                renderTable(dashFilterRows(INIT.rows, filters));
            });
        });
    }
    function applyFilters() {
        var rows = renderKpis();
        renderCharts(rows);
        renderTable(rows);
    }
    function bindFilters() {
        var chips = document.querySelectorAll('.f-chip');
        Array.prototype.forEach.call(chips, function (ch) {
            ch.addEventListener('click', function () {
                var dim = ch.getAttribute('data-dim');
                var val = ch.getAttribute('data-val');
                if (!Array.isArray(filters[dim])) filters[dim] = [];
                if (val === '') {
                    filters[dim] = [];   // 「全部」：清空该维度选择
                } else {
                    var i = filters[dim].indexOf(val);
                    if (i >= 0) filters[dim].splice(i, 1); else filters[dim].push(val);
                }
                syncChips();
                applyFilters();
            });
        });
    }
    function syncChips() {
        var chips = document.querySelectorAll('.f-chip');
        Array.prototype.forEach.call(chips, function (ch) {
            var dim = ch.getAttribute('data-dim');
            var val = ch.getAttribute('data-val');
            var sel = filters[dim];
            var on = val === '' ? (!Array.isArray(sel) || sel.length === 0) : (Array.isArray(sel) && sel.indexOf(val) !== -1);
            ch.classList.toggle('active', on);
        });
    }
        // 初始图表始终用服务端渲染的初始 option（保证 prepared 缺失时也不白屏）
    initCharts();
    // Gauge 仪表盘初始化
    (function initGauge() {
        var gaugeEls = document.querySelectorAll('.gauge-panel .chart');
        for (var i = 0; i < gaugeEls.length; i++) {
            var opt = null;
            try { opt = JSON.parse(gaugeEls[i].getAttribute('data-opt')); } catch (e) { ChatLog.debug?.('UI', '[dashboard] 仪表盘 data-opt 非 JSON，跳过初始化:', e); }
            if (opt && window.echarts) {
                var inst = echarts.init(gaugeEls[i]);
                inst.setOption(opt);
                instances['gauge-' + i] = inst;
            }
        }
    })();
    // 实时时钟更新（每秒）
    (function startClock() {
        function tick() {
            var now = new Date();
            var pad = function(n) { return n < 10 ? '0' + n : '' + n; };
            document.getElementById('clock-time').textContent = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
            var weeks = ['周日','周一','周二','周三','周四','周五','周六'];
            document.getElementById('clock-date').textContent = now.getFullYear() + '-' + pad(now.getMonth()+1) + '-' + pad(now.getDate()) + ' ' + weeks[now.getDay()];
        }
        tick(); setInterval(tick, 1000);
    })();
    // 浏览器窗口缩放 → 图表跟随容器自适应重绘（ECharts 不会自动监听 window resize）
    window.addEventListener('resize', function () {
        for (var i in instances) {
            if (instances[i] && typeof instances[i].resize === 'function') instances[i].resize();
        }
    });
    // 有原始数据且可筛选维度 → 初始按全量重算一次（与 prepared 一致）+ 绑定筛选
    if (INIT.rows && INIT.rows.length && INIT.dims && INIT.dims.length) {
        applyFilters();
    }
    if (INIT.dims && INIT.dims.length) {
        bindFilters();
    }
    bindTableLimit();
    if (INIT.truncated) {
        var bar = document.createElement('div');
        bar.style.cssText = 'margin:10px 0;padding:8px 14px;font-size:12px;color:#F5C044;background:rgba(245,192,68,.12);border:1px solid rgba(245,192,68,.35);border-radius:6px;';
        bar.textContent = '数据量较大，筛选基于前 ' + INIT.rows.length + ' 行样本计算';
        var fp = document.querySelector('.filter-panel');
        if (fp) fp.appendChild(bar);
    }
})();`;

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>${esc(spec.title || '数据看板')}大屏</title>
<script src="__ECHARTS_LOCAL__"></script>
<style>
* { margin:0; padding:0; box-sizing:border-box; }
body {
    font-family: 'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC', 'Helvetica Neue', Arial, sans-serif;
    background: radial-gradient(1200px 600px at 20% -10%, ${t.panel}, transparent 60%),
                radial-gradient(1000px 500px at 90% 0%, ${t.panel}, transparent 55%),
                ${t.bg};
    color: ${t.title};
    min-height: 100vh; padding: 22px 28px 40px;
}
.header { text-align:center; margin-bottom: 20px; position:relative; padding: 30px 0 24px; }
.header::before {
    content:''; position:absolute; left:50%; top:0; transform:translateX(-50%);
    width:200px; height:3px; background: linear-gradient(90deg, transparent, ${t.accent}, transparent); border-radius:2px;
}
.header::after {
    content:''; position:absolute; left:50%; bottom:0; transform:translateX(-50%);
    width:120px; height:2px; background: linear-gradient(90deg, transparent, ${t.accent}, transparent); border-radius:2px;
}
.header h1 { font-size: 30px; letter-spacing: 4px; font-weight: 600; text-shadow: 0 0 20px ${t.accent}40; }
.header .sub { font-size: 13px; color: ${t.sub}; margin-top: 8px; letter-spacing: 2px; }
.clock-widget { position:absolute; right:10px; top:10px; text-align:right; padding:10px 16px; background:rgba(255,255,255,.04); border:1px solid ${t.panelBorder}; border-radius:8px; }
.clock-date { font-size:11px; color:${t.kpiLabel}; letter-spacing:1px; }
.clock-time { font-size:20px; font-weight:700; color:${t.kpiValue}; font-variant-numeric: tabular-nums; letter-spacing:1px; margin-top:2px; }
.filter-panel { margin-bottom: 20px; }
.f-group { display:flex; align-items:flex-start; gap:8px; margin-bottom:8px; }
.f-group:last-child { margin-bottom:0; }
.f-label { font-size:12px; color:${t.kpiLabel}; width:74px; flex-shrink:0; line-height:26px; letter-spacing:1px; }
.f-chips { display:flex; flex-wrap:wrap; gap:6px; }
.f-chip {
    background: rgba(255,255,255,.06); border:1px solid ${t.panelBorder}; color:${t.title};
    font-size:12px; border-radius:14px; padding:4px 12px; cursor:pointer; transition: all .15s; line-height:1.4;
}
.f-chip:hover { border-color:${t.accent}; color:${t.accent}; }
.f-chip.active { background:${t.accent}; border-color:${t.accent}; color:#fff; font-weight:600; }
.kpi-row { display:grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 14px; margin-bottom: 20px; }
.kpi-card {
    background: ${t.panel}; border:1px solid ${t.panelBorder}; border-radius:8px;
    padding: 14px 16px; text-align:center; box-shadow: 0 4px 18px rgba(0,0,0,.25);
    animation: rise .5s ease both; position:relative; overflow:hidden;
}
.kpi-card::before {
    content:''; position:absolute; top:0; left:0; right:0; height:3px;
    background: linear-gradient(90deg, transparent, ${t.accent}, transparent); opacity:.8;
}
@keyframes rise { from { opacity:0; transform: translateY(10px); } to { opacity:1; transform:none; } }
.kpi-label { font-size: 13px; color: ${t.kpiLabel}; letter-spacing: 1px; }
.kpi-unit { font-size: 10px; opacity:.75; margin-left:2px; }
.kpi-value { font-size: 26px; font-weight: 700; color: ${t.kpiValue}; margin-top: 6px; font-variant-numeric: tabular-nums; }
.chart-grid { display:grid; grid-template-columns: repeat(auto-fit, minmax(420px, 1fr)); gap: 14px; margin-bottom: 20px; }
.chart-card { background: ${t.panel}; border:1px solid ${t.panelBorder}; border-radius:8px; padding: 12px 14px; position:relative; overflow:hidden; }
.chart-card::before { content:''; position:absolute; top:0; left:0; width:3px; height:100%; background:${t.accent}; opacity:.6; }
.chart-title { font-size: 13px; color: ${t.title}; margin-bottom: 8px; letter-spacing: 1px; padding-left: 6px; border-left: 3px solid ${t.accent}; }
.chart { width: 100%; height: 280px; }
.panel { background: ${t.panel}; border:1px solid ${t.panelBorder}; border-radius:8px; padding: 14px 16px; }
.panel-title { font-size: 13px; color: ${t.title}; margin-bottom: 10px; padding-left: 6px; border-left: 3px solid ${t.accent}; letter-spacing: 1px; }
.table-count { font-size: 11px; color:${t.kpiLabel}; margin-left:8px; font-weight:400; }
.table-limit { display:inline-flex; gap:4px; margin-left:10px; vertical-align:middle; }
.limit-btn { background:rgba(255,255,255,.05); border:1px solid ${t.panelBorder}; color:${t.kpiLabel}; border-radius:4px; padding:1px 8px; font-size:11px; cursor:pointer; line-height:1.6; }
.limit-btn.active { background:${t.accent}; color:#fff; border-color:${t.accent}; }
.table-wrap { overflow-x:auto; max-height: 320px; overflow-y:auto; }
table { width:100%; border-collapse: collapse; font-size: 12px; }
th { background: ${t.accent}; color:#fff; padding: 7px 10px; text-align:left; font-weight: 600; position: sticky; top: 0; }
td { padding: 6px 10px; border-bottom: 1px solid ${t.grid}; color: ${t.title}; }
tr:nth-child(even) td { background: rgba(255,255,255,.03); }
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-thumb { background: ${t.grid}; border-radius: 3px; }
/* ── 排行榜组件 ── */
.rank-panel { padding:12px 16px; }
.rank-list { display:flex; flex-direction:column; gap:8px; }
.rank-item { display:grid; grid-template-columns:24px 90px 1fr auto; gap:8px; align-items:center; font-size:12px; }
.rank-num { font-size:16px; font-weight:700; text-align:center; font-variant-numeric: tabular-nums; }
.rank-name { color:${t.title}; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
.rank-bar-wrap { height:10px; background:rgba(255,255,255,.06); border-radius:5px; overflow:hidden; }
.rank-bar { height:100%; border-radius:5px; transition:width .6s ease; }
.rank-val { color:${t.kpiValue}; font-weight:600; font-variant-numeric: tabular-nums; }
/* ── 左右布局容器（KPI+仪表盘 并排，图表+排行榜 并排）── */
.side-by-side { display:grid; grid-template-columns: 1fr 280px; gap:14px; margin-bottom:20px; }
.main-area { margin-bottom:20px; }
/* 响应式：小屏变单列 */
@media (max-width: 1024px) { .side-by-side { grid-template-columns: 1fr; } .clock-widget { position:static; margin:10px auto; } }
</style>
</head>
<body>
<div class="header">
    ${clockHtml}
    <h1>${esc(spec.title || '数据看板')}</h1>
    ${spec.subtitle ? '<div class="sub">' + esc(spec.subtitle) + '</div>' : ''}
</div>
${filterHtml}
<div class="side-by-side">
    <div class="main-area">
        <div class="kpi-row">${kpiCards}</div>
        ${charts.length ? '<div class="chart-grid">' + chartCards + '</div>' : ''}
        ${tableHtml}
    </div>
    <div class="side-panel">
        ${gaugeHtml}
        ${rankHtml}
    </div>
</div>
<script>
window.__DASH_INIT__ = ${initJson};
</script>
<script>
${pureSrc}
</script>
<script>
${bootSrc}
</script>
</body>
</html>`;
    }

    function buildHtmlStep(prepared, spec) {
        if (spec) spec.title = sanitizeBoardTitle(spec.title, dashboardFallbackTitle(prepared));
        const deliverable = String(spec.deliverable || 'dual').toLowerCase();
        if (deliverable === 'excel') return null;
        return {
            title: '生成 HTML 大屏', detail: (spec.title || '数据看板') + '大屏',
            tool: '__html__',
            displayName: '生成 HTML 大屏', displayArgs: { '页面标题': (spec.title || '数据看板') + '大屏' }
        };
    }

    ChatLog.info?.('CHAT-V2', '✅ dashboardHtmlRenderer loaded (含 JS 筛选器)');
    App.ChatV2.dashboardModules.htmlRenderer = {
        renderHtml, buildHtmlStep, echartsOptionFor, fmtNum,
        sanitizeBoardTitle, dashboardFallbackTitle,
        MAX_FILTER_ROWS,
        dashFilterRows, dashDistinct, dashAggNum, dashRecalcKpis, dashChartData
    };
})();
