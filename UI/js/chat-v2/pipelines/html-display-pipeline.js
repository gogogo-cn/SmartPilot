/**
 * ============================================================
 * chat-v2/pipelines/html-display-pipeline.js — HTML 浏览器展示 Pipeline
 * ============================================================
 * 职责：
 *   - 处理 HtmlBrowserDisplay Skill 路由来的"用 HTML/浏览器展示选中数据"请求
 *   - 确定性执行：读取选区 → 生成 HTML 表格 → 调用浏览器展示 MCP 工具
 *   - 不经过 LLM，零延迟确定性执行，提供步骤反馈卡片
 *
 * 触发方式：
 *   - 由 HtmlBrowserDisplay Skill 声明 pipeline: html，经 shouldRouteToSkill 匹配后路由到此
 *   - Skill 定义在 Skills/HtmlBrowserDisplay.md，通过 keywords 决定何时命中
 *   - 新增类似的前端确定性快捷路径：创建 Skill .md 文件声明 pipeline 字段即可，无需改此文件
 *
 * 与后端工具的关系：
 *   - 此 pipeline 在前端生成带样式的 HTML 表格，并通过 MCP 工具打开浏览器
 *   - 提供 UI 步骤反馈（显示"读取选区数据"等步骤卡片）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    function escapeHtml(text) {
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(text) : '';
    }

    function extractFilePath(result, depth) {
        if (!result) return '';
        depth = depth || 0;
        if (depth > 6) return '';
        // 仅接受真实的绝对 Windows html 路径，避免把标题/任意字符串误判为文件
        const PATH_RE = /[A-Za-z]:\\[^\s"]+\.html/i;
        if (typeof result === 'string') {
            const parsed = tryParseJson(result);
            if (parsed) return extractFilePath(parsed, depth + 1);
            const match = result.match(PATH_RE);
            return match ? match[0] : '';
        }
        if (typeof result === 'object') {
            if (typeof result.filePath === 'string' && PATH_RE.test(result.filePath)) return result.filePath;
            if (typeof result.path === 'string' && PATH_RE.test(result.path)) return result.path;
            if (typeof result.result === 'string') {
                const match = result.result.match(PATH_RE);
                if (match) return match[0];
            }
            // 递归解包 result/data/content 等承载键
            for (const key of ['result', 'data']) {
                const v = result[key];
                if (v === undefined || v === null) continue;
                const found = extractFilePath(v, depth + 1);
                if (found) return found;
            }
            // MCP 标准 content 数组：拼接 text 再解包
            if (Array.isArray(result.content)) {
                for (const c of result.content) {
                    const t = c && c.text !== undefined ? c.text : c;
                    const found = extractFilePath(t, depth + 1);
                    if (found) return found;
                }
            }
        }
        return '';
    }

    function tryParseJson(val) {
        if (typeof val !== 'string') return null;
        return App.ChatV2.jsonUtils ? App.ChatV2.jsonUtils.safeParse(val) : null;
    }

    function normalizeRangeData(raw) {
        if (!raw) return null;

        // 递归解包，找到实际数据
        function deepUnwrap(obj, depth) {
            depth = depth || 0;
            if (depth > 6) return obj;
            if (typeof obj === 'string') {
                const parsed = tryParseJson(obj);
                return parsed ? deepUnwrap(parsed, depth + 1) : obj;
            }
            if (!obj || typeof obj !== 'object') return obj;
            // 解包 { result: ..., taskCompleted: true } / { success: true, result: ... }
            if (obj.result !== undefined && obj.result !== null) {
                return deepUnwrap(obj.result, depth + 1);
            }
            if (obj.data !== undefined && typeof obj.data === 'object') {
                return deepUnwrap(obj.data, depth + 1);
            }
            // MCP 标准格式: { content: [{type:"text", text:"..."}] }
            if (Array.isArray(obj.content)) {
                const texts = obj.content.filter(c => c && c.type === 'text' && typeof c.text === 'string').map(c => c.text);
                if (texts.length > 0) {
                    const combined = texts.join('\n');
                    const parsed = tryParseJson(combined);
                    return deepUnwrap(parsed !== null ? parsed : combined, depth + 1);
                }
            }
            // ExcelBridge GetSelectedRangeData 返回 { success:true, headers:[], rows:[], ... }
            if (obj.success === true && (obj.headers || obj.rows)) {
                return obj;
            }
            return obj;
        }

        const data = deepUnwrap(raw);
        if (!data) return null;

        // 格式1: 直接是数组（MCP 返回行对象数组 [{列名:值,...}, ...]，keys 即列名）
        if (Array.isArray(data)) {
            if (data.length === 0) return { headers: [], rows: [] };
            const headers = Object.keys(data[0] || {});
            return { headers, rows: data };
        }

        // 格式2: {headers:[], rows:[]} 或 {success:true, headers:[], rows:[]}
        if (typeof data === 'object') {
            if (Array.isArray(data.rows)) {
                if (!Array.isArray(data.headers) && data.rows.length > 0) {
                    return { ...data, headers: Object.keys(data.rows[0] || {}) };
                }
                return data;
            }
            if (Array.isArray(data.headers) && data.rowCount != null) {
                return data;
            }
        }

        // 尝试在候选字段中找数组
        const candidates = [data, data.result, data.data, data.content];
        for (const candidate of candidates) {
            const unwrapped = deepUnwrap(candidate);
            if (Array.isArray(unwrapped) && unwrapped.length > 0) {
                const headers = Object.keys(unwrapped[0] || {});
                return { headers, rows: unwrapped };
            }
            if (unwrapped && typeof unwrapped === 'object' && Array.isArray(unwrapped.rows)) {
                if (!Array.isArray(unwrapped.headers) && unwrapped.rows.length > 0) {
                    return { ...unwrapped, headers: Object.keys(unwrapped.rows[0] || {}) };
                }
                return unwrapped;
            }
        }

        return null;
    }

    function buildHtmlTable(data, title) {
        const headers = Array.isArray(data?.headers) ? data.headers : [];
        const rows = Array.isArray(data?.rows) ? data.rows : [];
        const address = data?.address || data?.range || '';
        const sheetName = data?.sheetName || data?.worksheet || '';
        const workbookName = data?.workbookName || data?.workbook || '';

        // 单元格取值：兼容「行=数组」与「行=对象{列名:值}」两种返回
        const cell = (row, idx, header) => Array.isArray(row) ? row[idx] : row?.[header];
        const toNum = (v) => {
            if (v === null || v === undefined || v === '') return null;
            const n = Number(String(v).replace(/[$,%¥￥元]/g, ''));
            return isNaN(n) ? null : n;
        };
        // 列是否数值型（多数行可解析为数字）
        const colIsNumeric = (idx) => {
            if (rows.length === 0) return false;
            let numeric = 0;
            for (const r of rows) {
                if (toNum(cell(r, idx, headers[idx])) !== null) numeric++;
            }
            return numeric / rows.length >= 0.6;
        };
        // 分类列：第一个非数值列；全是数值则回退到第一列
        let catIdx = headers.findIndex((_, i) => !colIsNumeric(i));
        if (catIdx < 0) catIdx = 0;
        // 数值列：剔除分类列后可解析的列
        const numIdxs = headers.map((_, i) => i).filter(i => i !== catIdx && colIsNumeric(i));

        // ---- 生成表格 ----
        const headerHtml = headers.map(h => `<th>${escapeHtml(h)}</th>`).join('');
        const rowsHtml = rows.map(row => {
            const cells = headers.map((h, i) => `<td>${escapeHtml(cell(row, i, h))}</td>`).join('');
            return `<tr>${cells}</tr>`;
        }).join('');

        // ---- KPI 汇总卡片 ----
        const kpis = [];
        kpis.push({ label: '数据行', value: rows.length.toLocaleString('zh-CN') });
        const sumCard = (idx, label) => {
            const sum = rows.reduce((acc, r) => { const n = toNum(cell(r, idx, headers[idx])); return acc + (n === null ? 0 : n); }, 0);
            kpis.push({ label, value: sum.toLocaleString('zh-CN', { maximumFractionDigits: 2 }) });
        };
        if (numIdxs.length > 0) sumCard(numIdxs[0], `${headers[numIdxs[0]]} 合计`);
        if (rows.length > 0 && numIdxs.length > 0) {
            const avg = kpis[1] && typeof kpis[1].value === 'string'
                ? (rows.reduce((a, r) => { const n = toNum(cell(r, numIdxs[0], headers[numIdxs[0]])); return a + (n === null ? 0 : n); }, 0) / rows.length)
                : 0;
            kpis.push({ label: `${headers[numIdxs[0]]} 均值`, value: avg.toLocaleString('zh-CN', { maximumFractionDigits: 2 }) });
        }
        if (numIdxs.length > 1) sumCard(numIdxs[1], `${headers[numIdxs[1]]} 合计`);

        // ---- 图表：注入完整数据矩阵，让用户在浏览器里自行选择 X 轴列 / 数值列 / 图表类型 ----
        const matrix = rows.map(r => headers.map((h, i) => { const v = cell(r, i, h); return (v === null || v === undefined) ? '' : v; }));
        const dataJson = JSON.stringify({ headers, matrix })
            .replace(/</g, '\\u003c').replace(/\u2028/g, '\\u2028').replace(/\u2029/g, '\\u2029');

        const subtitle = [workbookName, sheetName, address].filter(Boolean).join(' / ');
        // 预选默认：X 轴=第一个非数值列（全数值则用第一列），数值=首个数值列
        const selX = headers[catIdx] || headers[0] || '';
        const selY = numIdxs.length > 0 ? headers[numIdxs[0]] : '';
        const xOptions = headers.map((h, i) => `<option value="${i}"${h === selX ? ' selected' : ''}>${escapeHtml(h)}</option>`).join('');
        const yOptions = numIdxs.map((i, ki) => `<option value="${i}"${ki === 0 ? ' selected' : ''}>${escapeHtml(headers[i])}</option>`).join('');
        const hasChart = numIdxs.length > 0 && rows.length > 0;

        return `<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>${escapeHtml(title)}</title>
    <style>
        body { font-family:'Microsoft YaHei','PingFang SC','Noto Sans CJK SC',-apple-system,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif; margin:20px; background:#f5f6f8; color:#1f2937; }
        .container { max-width:1200px; margin:0 auto; background:#fff; border:1px solid #e2e4e8; border-radius:8px; padding:20px; }
        h1 { font-size:18px; margin:0 0 6px; color:#111827; }
        .subtitle { color:#6b7280; font-size:13px; margin-bottom:16px; }
        .kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(160px,1fr)); gap:12px; margin-bottom:18px; }
        .kpi-card { border:1px solid #e5e7eb; border-radius:8px; padding:12px 14px; background:#fafafa; }
        .kpi-card .v { font-size:20px; font-weight:600; color:#2563eb; font-variant-numeric:tabular-nums; }
        .kpi-card .l { font-size:12px; color:#6b7280; margin-top:2px; }
        .ctl { display:flex; flex-wrap:wrap; align-items:center; gap:14px; border:1px solid #e5e7eb; border-radius:8px; padding:10px 12px; margin-bottom:12px; background:#fafbfc; }
        .ctl-item { display:flex; align-items:center; gap:6px; font-size:13px; color:#374151; }
        .ctl-item select { font-family:inherit; font-size:13px; border:1px solid #d1d5db; border-radius:6px; padding:4px 8px; background:#fff; color:#111827; }
        .types { display:flex; gap:6px; }
        .tbtn { font-family:inherit; font-size:12px; border:1px solid #d1d5db; border-radius:6px; padding:4px 10px; background:#fff; color:#374151; cursor:pointer; }
        .tbtn.on { background:#2563eb; border-color:#2563eb; color:#fff; }
        .chart-box { border:1px solid #e5e7eb; border-radius:8px; padding:14px; margin-bottom:18px; }
        #chart { width:100%; height:340px; }
        .table-wrap { max-height:460px; overflow:auto; border:1px solid #e2e4e8; border-radius:8px; }
        table { width:100%; border-collapse:separate; border-spacing:0; font-size:13px; }
        th, td { border-bottom:1px solid #eceef1; padding:7px 11px; text-align:left; white-space:nowrap; }
        th { background:#f3f4f6; font-weight:600; position:sticky; top:0; color:#374151; }
        tr:nth-child(even) td { background:#fafbfc; }
        tr:hover td { background:#f0f6ff; }
        tfoot td { font-weight:600; background:#f8fafc; }
        .footer { margin-top:16px; color:#9ca3af; font-size:12px; }
    </style>
</head>
<body>
    <div class="container">
        <h1>${escapeHtml(title)}</h1>
        <div class="subtitle">${escapeHtml(subtitle)} · 共 ${rows.length} 行 · 支持按分类聚合 / 自由切换 X 轴与图表类型</div>
        <div class="kpi-grid">
            ${kpis.map(k => `<div class="kpi-card"><div class="v">${escapeHtml(k.value)}</div><div class="l">${escapeHtml(k.label)}</div></div>`).join('')}
        </div>
        ${hasChart ? `
        <div class="ctl">
            <span class="ctl-item"><label>X 轴</label><select id="xSel">${xOptions}</select></span>
            <span class="ctl-item"><label>数值</label><select id="ySel">${yOptions}</select></span>
            <span class="ctl-item types">
                类型：
                <button class="tbtn on" data-type="bar">柱状</button>
                <button class="tbtn" data-type="line">折线</button>
                <button class="tbtn" data-type="area">面积</button>
                <button class="tbtn" data-type="pie">饼图</button>
            </span>
            <span class="ctl-item types">
                口径：
                <button class="tbtn abtn on" data-aggr="1">分类聚合</button>
                <button class="tbtn abtn" data-aggr="0">显示明细</button>
            </span>
        </div>
        <div class="chart-box"><div id="chart"></div></div>
        ` : ''}
        <div class="table-wrap">
            <table>
                <thead><tr>${headerHtml}</tr></thead>
                <tbody>${rowsHtml}</tbody>
            </table>
        </div>
        <div class="footer">Generated by Excel AI</div>
    </div>
    ${hasChart ? `
    <script src="https://cdn.jsdelivr.net/npm/echarts@5.4.3/dist/echarts.min.js"><\/script>
    <script>
        (function(){
            var D = ${dataJson};
            var chart = echarts.init(document.getElementById('chart'));
            var xIn = document.getElementById('xSel');
            var yIn = document.getElementById('ySel');
            var btns = Array.prototype.slice.call(document.querySelectorAll('.tbtn[data-type]'));
            var aBtns = Array.prototype.slice.call(document.querySelectorAll('.abtn'));
            var type = 'bar';
            var aggr = true; // 数据口径：默认按分类聚合求和
            var PIE_MAX = 30;
            function toNum(v){
                if (v === null || v === undefined || v === '') return null;
                var n = Number(String(v).replace(/[$,%\u00a5\u5143]/g, ''));
                return isNaN(n) ? null : n;
            }
            function draw(){
                var x = parseInt(xIn.value, 10) || 0;
                var y = parseInt(yIn.value, 10) || 0;
                var rawCat = D.matrix.map(function(r){ return String(r[x]); });
                var rawVals = D.matrix.map(function(r){ var n = toNum(r[y]); return n === null ? 0 : n; });
                var cat = rawCat, vals = rawVals;
                // 数据口径：分类聚合时按 X 分类值合并求和；饼图必须聚合（否则同名扇区重复）
                if (aggr || type === 'pie') {
                    var m = {}; var order = [];
                    for (var i = 0; i < rawCat.length; i++) {
                        var c = rawCat[i];
                        if (!(c in m)) { m[c] = 0; order.push(c); }
                        m[c] = m[c] + rawVals[i];
                    }
                    cat = order; vals = order.map(function(c){ return m[c]; });
                }
                var yName = D.headers[y] || '数值';
                var opt;
                if (type === 'pie') {
                    var data = cat.map(function(c, i){ return { name: c, value: vals[i] }; }).slice(0, PIE_MAX);
                    opt = {
                        tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
                        legend: { type: 'scroll', top: 28, bottom: 8 },
                        series: [{ type: 'pie', radius: '62%', center: ['50%', '54%'], data: data }]
                    };
                } else {
                    var lineT = (type === 'area') ? 'line' : type;
                    var s = { name: yName, type: lineT, data: vals };
                    if (type === 'bar') s.itemStyle = { borderRadius: [3, 3, 0, 0] };
                    if (type === 'line' || type === 'area') { s.smooth = true; s.lineStyle = { width: 2 }; }
                    if (type === 'area') s.areaStyle = {};
                    opt = {
                        tooltip: { trigger: 'axis' },
                        legend: { top: 24 },
                        grid: { left: '3%', right: '4%', bottom: '8%', containLabel: true },
                        xAxis: { type: 'category', data: cat, axisLabel: { interval: function(i){ return (i % Math.max(1, Math.ceil(cat.length / 12)) === 0); }, rotate: cat.length > 8 ? 30 : 0 } },
                        yAxis: { type: 'value', scale: true },
                        series: [s]
                    };
                }
                chart.setOption(opt, true);
            }
            xIn.onchange = draw;
            yIn.onchange = draw;
            aBtns.forEach(function(b){
                b.onclick = function(){
                    aggr = b.getAttribute('data-aggr') === '1';
                    aBtns.forEach(function(x){ x.classList.remove('on'); });
                    b.classList.add('on');
                    draw();
                };
            });
            btns.forEach(function(b){
                b.onclick = function(){
                    type = b.getAttribute('data-type');
                    btns.forEach(function(x){ x.classList.remove('on'); });
                    b.classList.add('on');
                    draw();
                };
            });
            window.addEventListener('resize', function(){ chart.resize(); });
            draw();
        })();
    <\/script>
    ` : ''}
</body>
</html>`;
    }

    async function* execute(turn, signal) {
        const { sessionStore, toolRegistry } = App.ChatV2;
        if (!sessionStore || !toolRegistry) {
            throw new Error('html-display-pipeline 依赖未就绪');
        }

        try {
            await toolRegistry.init();

            ChatLog.info?.('PIPELINE', `[html-display-pipeline] turn=${turn.id} 开始执行`);

            // 步骤面板：与 skill/text 通道一致，先渲染"读取→生成"骨架，执行中逐步填色
            const stepsPanelRef = App.ChatV2?.skillPipelineModules?.stepsPanel;
            if (stepsPanelRef?.initSkillStepsPanel) {
                stepsPanelRef.initSkillStepsPanel(turn, sessionStore, [
                    { id: 'html_p1', title: '读取数据', detail: '读取工作表数据并解析', status: 'pending', strategy: 'direct', toolName: 'excel_read_range' },
                    { id: 'html_p2', title: '生成 HTML 报表', detail: '生成含 KPI 与交互图表的报表并浏览器展示', status: 'pending', strategy: 'direct', toolName: 'excel_display_html_in_browser' }
                ]);
            }
            if (stepsPanelRef?.appendSkillStep) {
                stepsPanelRef.appendSkillStep(turn, sessionStore, { id: 'html_p1', status: 'running', detail: '正在读取数据...' });
            }

            // ========== 步骤 1：读取数据 ==========
            // 【公共数据源优先】若本 turn 的 text-pipeline（或上层）已确立 turn.context.publicSource，
            // 则锁定该表名+范围读取，避免用户切换活动表后读到错误数据；否则回退到当前选区。
            const publicSrc = turn?.context?.publicSource;
            const readArgs = {};
            if (publicSrc?.resolve && publicSrc.sheetName && publicSrc.rangeAddr) {
                readArgs.sheet = publicSrc.sheetName;
                readArgs.range = publicSrc.rangeAddr;
                ChatLog.info?.('PIPELINE',
                    `[html-display-pipeline] 使用公共数据源读取: ${publicSrc.sheetName}!${publicSrc.rangeAddr}`);
            }
            const readRes = await toolRegistry.invokeTracked(turn, 'excel_read_range', readArgs, {
                round: 1,
                ctx: { turn, signal }
            });

            if (signal?.aborted) {
                yield { type: 'canceled', reason: 'user_aborted' };
                return;
            }

            // 解析数据（支持多种返回格式）
            const rangeData = normalizeRangeData(readRes.result);
            if (!readRes.ok || !rangeData) {
                throw new Error(readRes.error || '读取选区数据失败');
            }

            const rowCount = Array.isArray(rangeData.rows) ? rangeData.rows.length : 0;

            if (stepsPanelRef?.appendSkillStep) {
                stepsPanelRef.appendSkillStep(turn, sessionStore, { id: 'html_p1', status: 'completed', detail: `读取成功 · ${rowCount} 行` });
                stepsPanelRef.appendSkillStep(turn, sessionStore, { id: 'html_p2', status: 'running', detail: '正在生成报表并打开展示...' });
            }

            if (signal?.aborted) {
                yield { type: 'canceled', reason: 'user_aborted' };
                return;
            }

            // ========== 步骤 2：生成 HTML 并在浏览器中展示 ==========
            const title = '选中数据报表';
            const htmlContent = buildHtmlTable(rangeData, title);

            // 优先通过注册工具展示，带 fallback
            let displayRes;
            const displayTool = toolRegistry.get('excel_display_html_in_browser');
            if (displayTool) {
                displayRes = await toolRegistry.invokeTracked(turn, 'excel_display_html_in_browser',
                    { htmlContent, title },
                    { round: 2, ctx: { turn, signal } }
                );
            } else {
                // Fallback：使用通用方法直接展示
                displayRes = await displayHtmlFallback(turn, htmlContent, title);
            }

            if (signal?.aborted) {
                yield { type: 'canceled', reason: 'user_aborted' };
                return;
            }

            if (!displayRes.ok) {
                throw new Error(displayRes.error || '浏览器展示失败');
            }

            // 真实保存路径从工具返回中提取，提示与实际文件位置保持一致
            const filePath = extractFilePath(displayRes?.result) || extractFilePath(displayRes);

            // 步骤面板收尾：标记"生成"完成 + 终结面板
            if (stepsPanelRef?.appendSkillStep) {
                stepsPanelRef.appendSkillStep(turn, sessionStore, { id: 'html_p2', status: 'completed', detail: filePath ? `已生成 → ${filePath}` : '已生成并在浏览器打开' });
            }
            if (stepsPanelRef?.finalizeSkillSteps) {
                stepsPanelRef.finalizeSkillSteps(turn, sessionStore);
            }

            // 数据驱动的"继续操作"建议数据源：把本 turn 的实际产出概要挂到 turn，
            // 供 suggestions 扫描后动态生成贴合内容的后续建议（而不是按管道类型写死）。
            const hasNumericCol = (() => {
                const hdrs = Array.isArray(rangeData?.headers) ? rangeData.headers : [];
                const rws = Array.isArray(rangeData?.rows) ? rangeData.rows : [];
                if (!hdrs.length || !rws.length) return false;
                for (let ci = 0; ci < hdrs.length; ci++) {
                    let n = 0;
                    for (const r of rws) {
                        const v = Array.isArray(r) ? r[ci] : r?.[hdrs[ci]];
                        const num = Number(String(v == null ? '' : v).replace(/[$,%¥￥元]/g, ''));
                        if (!isNaN(num)) n++;
                    }
                    if (n / rws.length >= 0.6) return true;
                }
                return false;
            })();
            turn._produce = {
                kind: 'html_report',
                filePath,
                rowCount,
                hasChart: rowCount > 0 && hasNumericCol,
                headers: Array.isArray(rangeData?.headers) ? rangeData.headers : []
            };

            const msg = filePath
                ? `🔗 通道：html 数据报表（前端确定性生成）\n✅ 已生成 HTML 报表《${title}》，共 ${rowCount} 行。\n文件已保存至：${filePath}`
                : `🔗 通道：html 数据报表（前端确定性生成）\n✅ 已打开 HTML 报表《${title}》，共 ${rowCount} 行。`;
            sessionStore.appendDelta(turn.id, 'content', msg);
            yield { type: 'content', delta: msg };

            sessionStore.completeTurn(turn.id);
            yield { type: 'pipeline_done', pipeline: 'html' };

        } catch (err) {
            ChatLog.error?.('PIPELINE', '[html-display-pipeline] 执行异常:', err);
            sessionStore.failTurn(turn.id, { message: err?.message || String(err) });
            yield { type: 'error', error: { code: 'PIPELINE_HTML_DISPLAY_FAILED', message: err?.message || String(err) } };
        }
    }

    /**
     * Fallback 展示 HTML：当 MCP 工具不可用时，直接调用 Bridge 或 Blob URL
     * 返回与 invokeTracked 兼容的结果格式
     */
    async function displayHtmlFallback(turn, htmlContent, title) {
        const sessionStore = App.ChatV2.sessionStore;
        const start = Date.now();

        // 创建一个 toolCall 卡片记录 fallback 操作
        const tc = sessionStore.addToolCall(turn.id, {
            name: 'excel_display_html_in_browser',
            displayName: '在浏览器中展示数据',
            args: { title },
            round: 2
        });

        let result;
        try {
            // Fallback 1：ExcelBridge
            const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
            if (bridge && typeof bridge.OpenHtmlInBrowser === 'function') {
                await bridge.OpenHtmlInBrowser(htmlContent, title);
                result = { ok: true, result: { opened: true }, durationMs: Date.now() - start };
            } else if (typeof window !== 'undefined' && window.URL && window.URL.createObjectURL) {
                // Fallback 2：Web blob URL
                const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
                const url = window.URL.createObjectURL(blob);
                window.open(url, '_blank');
                result = { ok: true, result: { opened: true, via: 'blob' }, durationMs: Date.now() - start };
            } else {
                result = { ok: false, error: '浏览器展示工具不可用', durationMs: Date.now() - start };
            }
        } catch (e) {
            result = { ok: false, error: e?.message || String(e), durationMs: Date.now() - start };
        }

        sessionStore.updateToolCall(turn.id, tc.id, {
            status: result.ok ? 'completed' : 'failed',
            result: result.result,
            error: result.error,
            durationMs: result.durationMs,
            completedAt: Date.now()
        });

        return { ...result, toolCallId: tc.id };
    }

    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('html', {
            execute,
            name: '浏览器展示',
            description: '读选区→生成 HTML 表格→浏览器展示（确定性执行，不走 LLM）',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.html = { execute };
    }
    ChatLog.info?.('CHAT-V2', '✅ html-display-pipeline loaded');
})();
