/**
 * ============================================================
 * chat-v2/services/dashboard-one-click.js — 一键生成看板（编排工具）
 * ============================================================
 * 职责：
 *   - 注册编排型工具 excel_generate_dashboard：一次调用完成
 *     「读数据 → 蓝图计算 → excelRenderer 施工（KPI/原生切片器/透视图）
 *      → 切片失败则回退内存图表」。
 *   - 确定性执行：完全不经 LLM 决策，LLM 只负责传参数（源数据/维度/指标/主题）。
 *   - 与 DashboardGenerator Skill 解耦：Skill 不可用时也能产出完整看板 + 切片器。
 *
 * 挂载：App.ChatV2.dashboardModules.oneClick
 * 注册：App.ChatV2.toolRegistry（source: 'builtin-dashboard'）
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.dashboardModules) App.ChatV2.dashboardModules = {};

    const ChatLog = window.ChatLog || console;
    const L = (msg) => { try { ChatLog.info?.('DASH-1CLICK', msg); } catch (e) { /* noop */ } };

    const THEME = {
        banner: '#1F4E79',
        bannerText: '#FFFFFF',
        kpiLabelBg: '#DEEBF7',
        kpiLabelFg: '#0F3D6E',
        kpiValueFg: '#123A6B',
        panelBg: '#FFFFFF',
        border: '#BDD7EE'
    };
    // 图表锚点：A1:I2 标题横幅、A3:C4 KPI 卡、图表从 B6 起依次排布
    const CHART_ANCHORS = ['B6:H18', 'B20:H32', 'B34:H46', 'B48:H60'];

    /** 把 read_range 的返回解析成 {headers, rows(二维数组), sheet, range} */
    function parseRangeResult(raw) {
        function unwrap(o, d) {
            if (!o || d > 6) return null;
            if (typeof o === 'string') {
                try { return unwrap(JSON.parse(o), d + 1); } catch (e) { return null; }
            }
            if (typeof o !== 'object') return null;
            if (Array.isArray(o.content)) {
                const texts = o.content.filter(c => c && typeof c.text === 'string').map(c => c.text).join('\n');
                if (texts) { try { return unwrap(JSON.parse(texts), d + 1); } catch (e) { return null; } }
            }
            if (Array.isArray(o) && o.length && o[0] && typeof o[0] === 'object' && 'headers' in o[0]) return unwrap(o[0], d + 1);
            if (Array.isArray(o.headers)) return o;
            if (o.result !== undefined && o.result !== null) return unwrap(o.result, d + 1);
            if (o.data !== undefined && o.data !== null) return unwrap(o.data, d + 1);
            return null;
        }
        const o = unwrap(raw, 0);
        if (!o || !Array.isArray(o.headers) || !o.headers.length) return null;
        const rows = Array.isArray(o.result) ? o.result : (Array.isArray(o.rows) ? o.rows : null);
        if (!Array.isArray(rows) || !rows.length) return null;
        return { sheet: o.worksheet || o.sheet || '', range: o.range || '', headers: o.headers.map(String), rows };
    }

    /**
     * 一键生成看板。
     * @param {Object} args
     *   - sourceSheet / sourceRange：源数据（缺省当前选区）
     *   - board：看板表名（默认"看板"）
     *   - dims：维度列数组（缺省引擎自动枚举，最多 4 个）
     *   - value：数值/指标列（缺省取首个数值列）
     *   - slicers：切片器维度列数组（缺省自动取分类维度前 3 个）
     *   - enableSlicer：是否启用原生切片器（默认 true；失败自动回退内存图表）
     *   - chartTypes：可选，与各维度一一对应的图表类型数组
     *   - title / theme：标题与主题
     * @param {Object} ctx - { turn, sessionStore, signal }
     */
    async function generateDashboard(args, ctx) {
        const reg = App.ChatV2 && App.ChatV2.toolRegistry;
        if (!reg) return { ok: false, error: 'toolRegistry 不可用' };
        const engine = App.ChatV2.dashboardModules.blueprintEngine;
        if (!engine || typeof engine.computeFromSources !== 'function') {
            return { ok: false, error: 'dashboardBlueprintEngine 不可用' };
        }
        const turn = ctx && ctx.turn;
        const board = String((args && (args.board || args.sheet)) || '看板').trim() || '看板';
        const a = args || {};
        const enableSlicer = a.enableSlicer !== false && a.slicers !== false;

        const call = async (name, callArgs) => {
            try {
                if (turn && typeof reg.invokeTracked === 'function') {
                    return await reg.invokeTracked(turn, name, callArgs, { ctx: { turn, signal: ctx && ctx.signal } });
                }
                return await reg.invoke(name, callArgs, ctx || {});
            } catch (e) {
                return { ok: false, error: String((e && e.message) || e) };
            }
        };

        // ① 读源数据
        const readArgs = {};
        if (a.sourceSheet || a.sheet) readArgs.sheet = a.sourceSheet || a.sheet;
        if (a.sourceRange || a.range) readArgs.range = a.sourceRange || a.range;
        const read = await call('excel_read_range', readArgs);
        const data = read && read.ok ? parseRangeResult(read.result) : null;
        if (!data || !data.rows.length) {
            return { ok: false, error: '读取源数据失败：请选中含表头的数据区域，或传 sourceSheet/sourceRange' };
        }
        // 切片器依赖 sheet!range；read_range 偶发不回填 → 用入参兜底
        const srcSheet = String(data.sheet || a.sourceSheet || a.sheet || '').trim();
        const srcRange = String(data.range || a.sourceRange || a.range || '').replace(/\$/g, '').trim();
        data.sheet = srcSheet;
        data.range = srcRange;
        L(`源数据：${srcSheet || '(活动表)'} ${srcRange || ''} headers=${data.headers.length} rows=${data.rows.length}`);

        // ② 蓝图计算
        const _hrOc = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.htmlRenderer;
        const _ocTitle = (_hrOc && typeof _hrOc.sanitizeBoardTitle === 'function')
            ? _hrOc.sanitizeBoardTitle(a.title || (srcSheet || '数据看板') + '数据分析', '数据看板')
            : String(a.title || '数据看板').slice(0, 30);
        const spec = {
            title: _ocTitle,
            theme: a.theme || 'business_blue',
            sheet: board,
            deliverable: 'excel'
        };
        if (Array.isArray(a.dims) && a.dims.length) {
            const val = a.value || (engine.firstNumericHeader ? engine.firstNumericHeader({ headers: data.headers, rows: data.rows }) : null);
            const types = Array.isArray(a.chartTypes) ? a.chartTypes : [];
            spec.charts = a.dims.slice(0, 4).map((dim, i) => ({
                dimension: dim,
                value: val || undefined,
                func: a.func || 'sum',
                chartType: types[i] || (i === 1 ? 'xlPie' : (i === 2 ? 'xlBarClustered' : 'xlColumnClustered'))
            }));
        }
        if (Array.isArray(a.slicers) && a.slicers.length) {
            spec.slicers = a.slicers.slice(0, 4).map(String);
        }
        let prepared;
        try {
            prepared = engine.computeFromSources([{
                id: 's1', sheet: srcSheet, range: srcRange, headers: data.headers, rows: data.rows
            }], spec);
        } catch (e) {
            return { ok: false, error: '蓝图计算失败: ' + String((e && e.message) || e) };
        }
        // 回填源定位，供 excelRenderer 判定 useSlicer
        if (prepared && Array.isArray(prepared.sources) && prepared.sources[0]) {
            if (!prepared.sources[0].range) prepared.sources[0].range = srcRange;
            if (!prepared.sources[0].sheet && !prepared.sources[0].worksheet) prepared.sources[0].sheet = srcSheet;
        }
        if (!enableSlicer && prepared && Array.isArray(prepared.sources) && prepared.sources[0]) {
            // 显式关闭切片：抹掉 range，强制走 dataValues 图表路径
            prepared.sources[0].range = '';
        }
        L(`蓝图：kpis=${(prepared.kpis || []).length} charts=${(prepared.charts || []).length} slicer=${enableSlicer && !!srcRange}`);

        rememberDashboardSource(srcSheet, srcRange, board);
        const excelRenderer = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.excelRenderer;
        if (excelRenderer && typeof excelRenderer.buildExcelPlan === 'function') {
            return await executeExcelPlan(prepared, spec, board, srcSheet, srcRange, call, enableSlicer);
        }
        // excelRenderer 未加载 → 旧版简化路径（无切片器）
        return await executeLegacyDashboard(prepared, spec, board, data, engine, call, a);
    }

    /** 从 excel_create_slicer 返回文本中解析 payload（VB 会在文末附带 JSON） */
    function parseSlicerPayload(r) {
        if (!r) return null;
        const bag = [];
        const push = (x) => { if (x != null) bag.push(x); };
        push(r.result);
        push(r.error);
        if (r.result && typeof r.result === 'object') {
            push(r.result.text);
            push(r.result.Text);
            if (Array.isArray(r.result.content)) {
                for (const c of r.result.content) push(c && (c.text || c.Text));
            }
        }
        for (const item of bag) {
            if (item && typeof item === 'object' && (item.slicersCreated != null || item.ptChartOkCount != null || Array.isArray(item.slicers))) {
                return item;
            }
            if (typeof item !== 'string' || !item) continue;
            // 优先取文末 Payload JSON
            const marker = item.lastIndexOf('=== Payload JSON ===');
            const slice = marker >= 0 ? item.slice(marker + '=== Payload JSON ==='.length) : item;
            const brace = slice.lastIndexOf('{');
            if (brace < 0) continue;
            try {
                const obj = JSON.parse(slice.slice(brace).trim());
                if (obj && typeof obj === 'object') return obj;
            } catch (_e) { /* try next */ }
            const m = item.match(/"slicersCreated"\s*:\s*(\d+)[\s\S]*?"ptChartOkCount"\s*:\s*(\d+)/)
                || item.match(/"ptChartOkCount"\s*:\s*(\d+)[\s\S]*?"slicersCreated"\s*:\s*(\d+)/);
            if (m) {
                const a = Number(m[1]) || 0, b = Number(m[2]) || 0;
                // 两组捕获顺序不定，用字段名再扫一次
                const sc = /"slicersCreated"\s*:\s*(\d+)/.exec(item);
                const pc = /"ptChartOkCount"\s*:\s*(\d+)/.exec(item);
                return {
                    slicersCreated: sc ? Number(sc[1]) || 0 : a,
                    ptChartOkCount: pc ? Number(pc[1]) || 0 : b
                };
            }
            // 人类可读汇总行：Slicer: 3/3 | PivotChart: 4/4
            const human = /Slicer:\s*(\d+)\s*\/\s*(\d+)[\s\S]*?PivotChart:\s*(\d+)\s*\/\s*(\d+)/i.exec(item)
                || /PivotChart:\s*(\d+)\s*\/\s*(\d+)[\s\S]*?Slicer:\s*(\d+)\s*\/\s*(\d+)/i.exec(item);
            if (human) {
                const nums = human.slice(1).map(n => Number(n) || 0);
                // 归一：取成对的 created/requested
                if (/^Slicer:/i.test(human[0])) {
                    return { slicersCreated: nums[0], ptChartOkCount: nums[2] };
                }
                return { ptChartOkCount: nums[0], slicersCreated: nums[2] };
            }
        }
        return null;
    }

    /**
     * 解读切片器调用结果。
     * VB 多层回退会在 warnings 里写「参数错误/FALLBACK」，上层常误判 ok=false；
     * 以 payload.slicersCreated / ptChartOkCount 为准。
     */
    function interpretSlicerResult(r) {
        const payload = parseSlicerPayload(r);
        const slicersCreated = payload
            ? Number(payload.slicersCreated != null ? payload.slicersCreated
                : (Array.isArray(payload.slicers) ? payload.slicers.length : 0)) || 0
            : 0;
        const chartsOk = payload ? Number(payload.ptChartOkCount || 0) || 0 : 0;
        if (slicersCreated > 0 || chartsOk > 0) {
            return { ok: true, slicersCreated, chartsOk, payload, softFail: !(r && r.ok) };
        }
        return { ok: !!(r && r.ok), slicersCreated: 0, chartsOk: 0, payload, softFail: false };
    }

    /** 用 excelRenderer 计划施工：含原生切片器；仅当透视图也未建成时才追加内存图表降级 */
    async function executeExcelPlan(prepared, spec, board, srcSheet, srcRange, call, enableSlicer) {
        const excelRenderer = App.ChatV2.dashboardModules.excelRenderer;
        const planSteps = excelRenderer.buildExcelPlan(prepared, spec) || [];
        if (!planSteps.length) {
            return { ok: false, error: '看板施工计划为空' };
        }
        let slicerOk = true;
        let slicerTried = false;
        let slicersCreated = 0;
        let pivotChartsOk = 0;
        const created = [];
        const failed = [];

        const appendFallbackCharts = () => {
            const charts = (prepared.charts || []).slice(0, CHART_ANCHORS.length);
            for (let i = 0; i < charts.length; i++) {
                const cd = charts[i];
                if (!cd || !cd.data || !cd.data.length) continue;
                const isPie = /xlPie|xlDoughnut/.test(String(cd.chartType || ''));
                planSteps.push({
                    title: `生成图表「${cd.title || cd.dim || (i + 1)}」（降级）`,
                    tool: 'excel_create_chart',
                    args: {
                        chartType: cd.chartType || 'xlColumnClustered',
                        title: cd.title || (cd.dim + '分布'),
                        dataValues: JSON.stringify([cd.header].concat(cd.data)),
                        sheet: board,
                        position: JSON.stringify({ anchorRange: CHART_ANCHORS[i] }),
                        seriesBy: 'columns',
                        hasLegend: !isPie,
                        showDataLabels: isPie,
                        chartStyle: 201,
                        multiChart: true
                    }
                });
            }
        };

        for (let di = 0; di < planSteps.length; di++) {
            const st = planSteps[di];
            if (!st || !st.tool || st.tool === '__html__') continue;
            if (st.tool === 'excel_create_slicer') slicerTried = true;
            if (st.tool === 'excel_create_chart' && !slicerOk && st._dashFallback && st._dashFallback.dataValues) {
                st.args = Object.assign({}, st.args, { dataValues: st._dashFallback.dataValues });
                delete st.args.dataRange;
                delete st.args.pivotSheet;
                delete st.args.pivotTable;
            }
            const r = await call(st.tool, st.args || {});
            if (st.tool === 'excel_create_slicer') {
                const judged = interpretSlicerResult(r);
                slicersCreated = judged.slicersCreated;
                pivotChartsOk = judged.chartsOk;
                if (judged.ok) {
                    slicerOk = true;
                    created.push({ title: st.title || st.tool, ok: true, tool: st.tool });
                    if (judged.softFail) {
                        L(`切片器上层误判失败但 payload 成功：slicers=${slicersCreated} charts=${pivotChartsOk} → 不降级`);
                    }
                    // 切片器在、透视图 0 → 才需要内存图补洞（此时切片无法联动静态图，文案会说明）
                    if (pivotChartsOk === 0 && (prepared.charts || []).length > 0) {
                        L('透视图未建成，追加内存图表保底');
                        appendFallbackCharts();
                    }
                } else {
                    slicerOk = false;
                    const err = String((r && r.error) || '失败').slice(0, 200);
                    failed.push({ tool: st.tool, error: err });
                    L('切片器失败，追加内存图表降级: ' + err);
                    appendFallbackCharts();
                }
                continue;
            }
            if (r && r.ok) {
                if (st.tool === 'excel_create_chart') {
                    created.push({ title: st.title || st.tool, ok: true, tool: st.tool });
                }
            } else {
                const err = String((r && r.error) || '失败').slice(0, 200);
                failed.push({ tool: st.tool, error: err });
            }
        }

        const kpis = (prepared.kpis || []).slice(0, 3);
        const chartsPlanned = (prepared.charts || []).length;
        const staticChartOk = created.filter(c => c.tool === 'excel_create_chart' && c.ok).length;
        const effectiveChartOk = (slicerTried && slicerOk && pivotChartsOk > 0)
            ? pivotChartsOk
            : (slicerTried && slicerOk && chartsPlanned > 0 && pivotChartsOk === 0)
                ? staticChartOk
                : (slicerTried && slicerOk && chartsPlanned > 0)
                    ? chartsPlanned
                    : staticChartOk;

        if (chartsPlanned > 0 && effectiveChartOk === 0 && !(slicerTried && slicerOk && slicersCreated > 0)) {
            return {
                ok: false,
                error: `看板工作表「${board}」已建 KPI，但图表/切片器全部失败`,
                result: {
                    board,
                    source: (srcSheet || '(活动表)') + (srcRange ? '!' + srcRange : ''),
                    kpis: kpis.map(k => ({ label: k.label, value: k.value })),
                    charts: created,
                    slicers: slicerTried ? (slicerOk ? 'ok' : 'failed') : 'skipped',
                    failed
                }
            };
        }
        let slicerNote = '（无可用切片维度）';
        if (!enableSlicer) slicerNote = '（未启用切片）';
        else if (slicerTried && slicerOk && slicersCreated > 0 && pivotChartsOk > 0) {
            slicerNote = ` + ${slicersCreated} 个原生切片器联动`;
        } else if (slicerTried && slicerOk && slicersCreated > 0 && pivotChartsOk === 0) {
            slicerNote = `（已建 ${slicersCreated} 个切片器，透视图失败故图表为静态、筛选不联动）`;
        } else if (slicerTried && !slicerOk) {
            slicerNote = '（切片失败已降级为静态图）';
        }
        return {
            ok: true,
            result: {
                board,
                source: (srcSheet || '(活动表)') + (srcRange ? '!' + srcRange : ''),
                kpis: kpis.map(k => ({ label: k.label, value: k.value })),
                charts: created,
                slicers: slicerTried ? (slicerOk ? 'ok' : 'failed') : 'skipped',
                slicersCreated,
                pivotChartsOk,
                message: `已生成看板工作表「${board}」：${kpis.length} 个 KPI + ${effectiveChartOk || staticChartOk} 张图表${slicerNote}`
            }
        };
    }

    /** 旧版简化路径：无切片器，内存 dataValues 建图（excelRenderer 不可用时） */
    async function executeLegacyDashboard(prepared, spec, board, data, engine, call, a) {
        await call('excel_create_sheet', { name: board });
        await call('excel_clear_sheet', { sheet: board });

        const kpis = (prepared.kpis || []).slice(0, 3);
        if (kpis.length) {
            const values = [
                kpis.map(k => String(k.label || '')),
                kpis.map(k => (k.numeric ? Number(k.value) || 0 : String(k.value == null ? '' : k.value)))
            ];
            await call('excel_write_range', { sheet: board, range: 'A3', values });
        }

        const charts = (prepared.charts || []).slice(0, CHART_ANCHORS.length);
        const created = [];
        for (let i = 0; i < charts.length; i++) {
            const cd = charts[i];
            let header = cd.header, rows = cd.data;
            if ((!header || !rows) && typeof engine.aggregateChartData === 'function') {
                const agg = engine.aggregateChartData(cd, { headers: data.headers, rows: data.rows });
                header = agg.header; rows = agg.data;
            }
            if (!header || !rows || !rows.length) continue;
            const dataValues = JSON.stringify([header].concat(rows));
            const r = await call('excel_create_chart', {
                chartType: cd.chartType || 'xlColumnClustered',
                title: cd.title || (cd.dim + '分布'),
                dataValues,
                sheet: board,
                position: JSON.stringify({ anchorRange: CHART_ANCHORS[i] }),
                seriesBy: 'columns',
                hasLegend: !/xlPie|xlDoughnut/.test(String(cd.chartType || '')),
                showDataLabels: /xlPie|xlDoughnut/.test(String(cd.chartType || '')),
                chartStyle: 201,
                multiChart: true
            });
            created.push({ title: cd.title || cd.dim, ok: !!(r && r.ok) });
        }

        const titleText = String(a.title || spec.title || '数据看板').slice(0, 24);
        await call('excel_add_shape', {
            sheet: board, range: 'A1:I2', shapeType: 'roundedRect',
            fill: THEME.banner, line: 'none', cornerRadius: 0.18, noShadow: true,
            lines: JSON.stringify([{ text: titleText, size: 18, color: THEME.bannerText, bold: true, alignH: 'center' }])
        });
        await call('excel_format_range', {
            sheet: board, range: 'A3:C3',
            format: { bold: true, fontSize: 11, fontColor: THEME.kpiLabelFg, bgColor: THEME.kpiLabelBg, rowHeight: 22 },
            alignment: JSON.stringify({ horizontal: 'center', vertical: 'middle' })
        });
        await call('excel_format_range', {
            sheet: board, range: 'A4:C4',
            format: { bold: true, fontSize: 16, fontColor: THEME.kpiValueFg, numberFormat: '#,##0.00', rowHeight: 30 },
            alignment: JSON.stringify({ horizontal: 'center', vertical: 'middle' })
        });
        await call('excel_format_range', {
            sheet: board, range: 'A6:I46',
            format: { bgColor: THEME.panelBg, bgScope: 'all', borderStyle: 'thin', borderColor: THEME.border }
        });
        await call('excel_format_range', { sheet: board, range: 'A1:I1', format: { columnWidth: 13 } });

        const chartOk = created.filter(c => c.ok).length;
        const chartTried = created.length;
        if (chartTried > 0 && chartOk === 0) {
            return {
                ok: false,
                error: `看板工作表「${board}」已建 KPI，但 ${chartTried} 张图表全部创建失败`,
                result: {
                    board,
                    source: (data.sheet || '(活动表)') + (data.range ? '!' + data.range : ''),
                    kpis: kpis.map(k => ({ label: k.label, value: k.value })),
                    charts: created
                }
            };
        }
        const partialNote = (chartTried > 0 && chartOk < chartTried)
            ? `（其中 ${chartTried - chartOk} 张图失败）` : '';
        return {
            ok: true,
            result: {
                board,
                source: (data.sheet || '(活动表)') + (data.range ? '!' + data.range : ''),
                kpis: kpis.map(k => ({ label: k.label, value: k.value })),
                charts: created,
                slicers: 'skipped',
                message: `已生成看板工作表「${board}」：${kpis.length} 个 KPI + ${chartOk} 张图表（含标题横幅与配色）${partialNote}`
            }
        };
    }

    /** 注册为 LLM 可见的编排工具 */
    function registerTool() {
        const reg = App.ChatV2 && App.ChatV2.toolRegistry;
        if (!reg || typeof reg.register !== 'function') {
            ChatLog.warn?.('DASH-1CLICK', 'toolRegistry 未就绪，跳过注册');
            return false;
        }
        reg.register({
            name: 'excel_generate_dashboard',
            displayName: '一键生成看板',
            description: '一键生成完整数据看板（确定性编排，无需多步调用）：读取源数据 → 计算 KPI → 新建看板工作表 → 写入 KPI 卡片 → 创建 Excel 原生切片器（共享透视缓存，筛选联动全部图表）→ 按维度创建透视图/图表 → 标题横幅与配色。看板/仪表盘/大屏/驾驶舱类需求【优先直接调用本工具一次完成】，不要拆成多个 excel_aggregate/excel_create_chart/excel_create_slicer 调用。',
            inputSchema: {
                type: 'object',
                properties: {
                    sourceSheet: { type: 'string', description: '【可选】源数据工作表名；不传则用当前活动表' },
                    sourceRange: { type: 'string', description: '【可选】源数据范围，如 A1:I101；不传则用当前选区/已用区域（切片器需要明确范围）' },
                    board: { type: 'string', description: '【可选】看板工作表名，默认「看板」' },
                    dims: { type: 'string', description: '【可选】维度列 JSON 数组，如 ["部门","地区","产品"]；不传则由引擎自动枚举最多 4 个分类维度' },
                    slicers: { type: 'string', description: '【可选】切片器维度 JSON 数组，如 ["地区","部门","客户类型"]；不传则自动取合格分类维度前 3 个' },
                    enableSlicer: { type: 'boolean', description: '【可选】是否启用原生切片器，默认 true；失败时自动降级为静态图表' },
                    value: { type: 'string', description: '【可选】指标/数值列名，如「销售额」；不传取首个数值列' },
                    func: { type: 'string', description: '【可选】聚合方式：sum(默认)/avg/count/max/min' },
                    chartTypes: { type: 'string', description: '【可选】与各维度一一对应的图表类型 JSON 数组，如 ["xlColumnClustered","xlPie","xlBarClustered"]' },
                    title: { type: 'string', description: '【可选】看板标题' },
                    theme: { type: 'string', description: '【可选】主题：business_blue/crimson_red/violet_purple/fresh_green/warm_orange/dark/tech_black/executive_gray/minimal_light；也可传「红色/紫色/蓝色主题」等中文' }
                },
                required: []
            },
            source: 'builtin-dashboard',
            async execute(args, ctx) {
                // 兼容 dims/chartTypes/slicers 以 JSON 字符串传入（LLM 常见写法）
                const norm = (v) => {
                    if (Array.isArray(v)) return v;
                    if (typeof v === 'string' && v.trim()) {
                        try { const p = JSON.parse(v); if (Array.isArray(p)) return p; } catch (e) { /* noop */ }
                        return v.split(/[,，]/).map(s => s.trim()).filter(Boolean);
                    }
                    return undefined;
                };
                const a = Object.assign({}, args || {});
                if (a.dims !== undefined) a.dims = norm(a.dims);
                if (a.chartTypes !== undefined) a.chartTypes = norm(a.chartTypes);
                if (a.slicers !== undefined) {
                    if (a.slicers === false || a.slicers === 'false') a.slicers = false;
                    else a.slicers = norm(a.slicers);
                }
                if (typeof a.enableSlicer === 'string') {
                    a.enableSlicer = !/^(0|false|no|off)$/i.test(a.enableSlicer.trim());
                }
                const eng = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.blueprintEngine;
                if (a.theme && eng && typeof eng.resolveThemeFromMessage === 'function' && !(eng.THEMES && eng.THEMES[a.theme])) {
                    a.theme = eng.resolveThemeFromMessage(a.theme) || a.theme;
                }
                return await generateDashboard(a, ctx);
            }
        });
        L('已注册编排工具 excel_generate_dashboard');
        return true;
    }

    /** 记住最近一次看板数据源，供「改主题」跨轮复用 */
    function rememberDashboardSource(sheet, range, board) {
        try {
            App.ChatV2._lastDashboardSource = {
                sheet: String(sheet || '').trim(),
                range: String(range || '').replace(/\$/g, '').trim(),
                board: String(board || '看板').trim() || '看板',
                ts: Date.now()
            };
        } catch (_e) { /* noop */ }
    }

    /**
     * 确定性改主题：用新主题整表重建看板（含切片器），不经 LLM。
     * @param {Object} args - { theme|themeId, board, sourceSheet, sourceRange, title }
     */
    async function changeDashboardTheme(args, ctx) {
        const engine = App.ChatV2.dashboardModules && App.ChatV2.dashboardModules.blueprintEngine;
        const a = Object.assign({}, args || {});
        let themeId = a.theme || a.themeId || '';
        if (themeId && engine && engine.THEMES && !engine.THEMES[themeId] && typeof engine.resolveThemeFromMessage === 'function') {
            themeId = engine.resolveThemeFromMessage(themeId) || themeId;
        }
        if (!themeId || (engine && engine.THEMES && !engine.THEMES[themeId])) {
            return { ok: false, error: '无法识别主题，请说「红色主题」「商务蓝」等' };
        }
        const last = App.ChatV2._lastDashboardSource || {};
        if (!a.sourceSheet && last.sheet) a.sourceSheet = last.sheet;
        if (!a.sourceRange && last.range) a.sourceRange = last.range;
        if (!a.board && last.board) a.board = last.board;
        a.theme = themeId;
        a.board = a.board || '看板';
        a.title = a.title || '数据看板';
        L(`改主题 → ${themeId} board=${a.board} src=${a.sourceSheet || ''}!${a.sourceRange || ''}`);
        const res = await generateDashboard(a, ctx);
        if (res && res.ok && res.result) {
            const themeName = (engine && engine.THEMES && engine.THEMES[themeId] && engine.THEMES[themeId].name) || themeId;
            res.result.theme = themeId;
            res.result.message = `已将看板切换为「${themeName}」主题` +
                (res.result.message ? `（${String(res.result.message).replace(/^已生成看板工作表[^：]*：/, '')}）` : '');
        }
        return res;
    }

    App.ChatV2.dashboardModules.oneClick = {
        generateDashboard, changeDashboardTheme, registerTool, parseRangeResult, rememberDashboardSource
    };

    // 延迟注册：等 toolRegistry 就绪（脚本加载顺序不确定）
    try {
        if (App.ChatV2.toolRegistry && typeof App.ChatV2.toolRegistry.register === 'function') {
            registerTool();
        } else {
            let tries = 0;
            const timer = setInterval(() => {
                tries++;
                if (App.ChatV2.toolRegistry && typeof App.ChatV2.toolRegistry.register === 'function') {
                    registerTool();
                    clearInterval(timer);
                } else if (tries > 40) {
                    clearInterval(timer);
                    L('注册超时：toolRegistry 始终未就绪');
                }
            }, 250);
        }
    } catch (e) {
        L('注册失败: ' + String((e && e.message) || e));
    }

    ChatLog.info?.('CHAT-V2', '✅ dashboard-one-click loaded（一键生成看板）');
})();
