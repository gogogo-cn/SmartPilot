/**
 * ============================================================
 * dashboard-excel-renderer.js — Excel 看板原子工具链生成（方案 B · 渲染层分离）
 * ============================================================
 * 职责：把 compute 层产出的 prepared{ kpis, charts, table } 转成 Excel COM 原子工具步骤数组。
 *       只管"生成 Excel 步骤"，不碰数据计算，不碰 HTML 渲染。
 *
 * 布局（形状卡片化）：
 *   行 1-2  标题横幅（圆角形状卡片：主题底色 + 标题/副标题写入形状文本）
 *   行 4-7  KPI 卡片带（每卡=圆角形状卡片，数值/标签写入形状文本，不合并单元格）
 *   行 8 起 图表 2 列网格
 *   图表后  明细表
 *
 * 依赖：从 dashboard-blueprint-engine 共享引入 themeOf / LAYOUT / colLetter / humanizeChartType / isCategoryChart
 * 挂载：App.ChatV2.dashboardModules.excelRenderer
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.dashboardModules) App.ChatV2.dashboardModules = {};

    const ChatLog = window.ChatLog || console;

    function _eng() { return App.ChatV2.dashboardModules.blueprintEngine || {}; }
    function _themeOf(spec) { const e = _eng(); return e.themeOf ? e.themeOf(spec) : { name: '默认', excel: { title: '#4DA3FF', label: '#8B93A7', value: '#E6E9F2', headerText: '#E6E9F2', headerBg: '#4DA3FF', bodyBg: '#192235', border: '#2B3A55' } }; }
    function _LAYOUT() { const e = _eng(); return e.LAYOUT || { kpiLabelRow: 4, kpiValueRow: 5, kpiCardEndRow: 7, kpiColWidth: 16, chartStartRow: 8, centerStartCol: 1, chartSlotCols: 6, chartSlotRows: 12, chartRowGap: 12, chartColGap: 0, slicerHeight: 114, slicerGap: 6, slicerLeftStart: 0, slicerWidth: 184, tableMaxRows: 200, dashboardCols: 22 }; }
    function _colLetter(n) { const e = _eng(); return e.colLetter ? e.colLetter(n) : (() => { let s = ''; n = Math.max(1, n); while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; })(); }
    function _humanize(ct) { const e = _eng(); return e.humanizeChartType ? e.humanizeChartType(ct) : String(ct || '图表'); }
    function _isCategory(ct) { const e = _eng(); return e.isCategoryChart ? e.isCategoryChart(ct) : /xlPie|xlDoughnut|xlFunnel|xlWaterfall|xlTreemap|xlSunburst/.test(String(ct || '')); }

    /** 形状方案强制启用（后端 Data.vb 已注册 excel_add_shape，不再运行时探测，避免 toolRegistry 时序问题导致回退单元格）。
     *  保留探测仅作日志诊断；执行失败时 skill-spec-execute.js 会自动提示"工具未就绪请重启"。 */
    function _canUseShapes() {
        try {
            const reg = App?.ChatV2?.toolRegistry;
            const hasTool = !!(reg && typeof reg.get === 'function' && reg.get('excel_add_shape'));
            console.log('[dash-excel] _canUseShapes → true (强制) | toolRegistry.has(excel_add_shape)=' + hasTool);
            return true;
        } catch (_e) {
            console.log('[dash-excel] _canUseShapes exception:', _e && _e.message);
            return true;
        }
    }

    /** 数值千分位格式化（用于写入 KPI 卡片形状文字，避免超长数字溢出） */
    function _thousands(n, dec) {
        if (n === null || n === undefined || n === '') return '';
        const d = (dec === undefined) ? 0 : dec;
        const num = (typeof n === 'number') ? n : parseFloat(n);
        if (isNaN(num)) return String(n);
        return num.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: d });
    }

    /** 数据源范围（sheet!range），供原生切片器/透视表引用 */
    function _sourceRange(prepared, spec) {
        // spec.dataSources[0] 优先，但其 range 无效/缺失时回退 prepared.sources[0]（运行时后端只回填 prepared，
        // spec.dataSources 可能只有 id 无 range，若直接取 spec 会误判无源 → 切片器被禁用）
        const specSrc = (spec.dataSources && spec.dataSources[0]) || null;
        const prepSrc = (prepared.sources && prepared.sources[0]) || null;
        let src = null;
        if (specSrc && specSrc.range && typeof specSrc.range === 'string' && specSrc.range.trim()) src = specSrc;
        else if (prepSrc && prepSrc.range && typeof prepSrc.range === 'string' && prepSrc.range.trim()) src = prepSrc;
        if (!src) return null;
        const r = src.range.trim();
        // 兼容 sheet / worksheet 两种字段名（blueprint 管道回填用 worksheet）
        const sh = (src.sheet && String(src.sheet).trim()) || (src.worksheet && String(src.worksheet).trim()) || '';
        if (sh) return sh + '!' + r;
        return r;
    }

    /** 切片器维度：spec.slicers 优先，否则取图表维度去重前 3 个（须是分类维度且存在于数据列） */
    function _resolveSlicerDims(prepared, spec) {
        const src = (prepared.sources && prepared.sources[0]) || null;
        const headers = src && Array.isArray(src.headers) ? src.headers : [];
        const headSet = new Set(headers);
        const _charts = prepared.charts || [];
        console.log('[dash-excel] _resolveSlicerDims START: headers.len=' + headers.length + ' charts.len=' + _charts.length + ' spec.slicers=' + JSON.stringify(spec.slicers));
        _charts.forEach((c, i) => console.log('  chart[' + i + '] keys=' + Object.keys(c).join(',') + ' dim=' + JSON.stringify(c.dim) + ' dimension=' + JSON.stringify(c.dimension)));
        if (Array.isArray(spec.slicers) && spec.slicers.length) {
            const r = spec.slicers.filter(h => headSet.has(h)).slice(0, 4);
            console.log('[dash-excel] _resolveSlicerDims → spec.slicers 路径 返回 ' + JSON.stringify(r));
            return r;
        }
        // 复用与兜底路径一致的维度合格性判断：排除日期/ID/数值列与高基数列
        const rows0 = src && Array.isArray(src.rows) ? src.rows : [];
        const isNum0 = (h) => rows0.some(r => { const v = r[h]; return v !== undefined && v !== null && v !== '' && Number.isFinite(Number(v)); });
        const isDate0 = (h) => rows0.some(r => r[h] !== undefined && r[h] !== null && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(String(r[h])));
        const isId0 = (h) => /(编号|ID|id|单号|订单号)$/.test(String(h));
        const dimCap = Math.min(50, Math.max(8, Math.floor(rows0.length * 0.5)));
        console.log('[dash-excel] _resolveSlicerDims dimCap=' + dimCap + ' rows0.len=' + rows0.length);
        const dimOk = (h) => {
            if (!h || isDate0(h) || isId0(h) || isNum0(h)) { console.log('    dimOk REJECT ['+h+'] date='+isDate0(h)+' id='+isId0(h)+' num='+isNum0(h)); return false; }
            const vals = rows0.map(r => (r[h] !== undefined && r[h] !== null && String(r[h]) !== '') ? String(r[h]) : null).filter(Boolean);
            const uniq = new Set(vals).size;
            const ok = uniq >= 2 && uniq <= dimCap;
            console.log('    dimOk ['+h+'] uniq='+uniq+' → '+ok);
            return ok;
        };
        const dims = [];
        for (const cd of _charts) {
            console.log('  路径1检查 cd.dim=' + JSON.stringify(cd.dim) + ' headSet.has=' + (cd.dim ? headSet.has(cd.dim) : 'N/A'));
            if (cd.dim && headSet.has(cd.dim) && dims.indexOf(cd.dim) < 0 && dimOk(cd.dim)) dims.push(cd.dim);
            if (dims.length >= 3) break;
        }
        console.log('[dash-excel] _resolveSlicerDims 路径1 dims=' + JSON.stringify(dims));
        if (dims.length) return dims.slice(0, 4);
        // 兜底 2：从数据列直接枚举合格分类维度（文本列、唯一值 2~50、排除日期/ID/数值），
        // 保证 spec.slicers 空且图表 dim 缺失时切片器仍能生成
        const rows = src && Array.isArray(src.rows) ? src.rows : [];
        const uniqCap = Math.min(50, Math.max(8, Math.floor(rows.length * 0.5)));
        // 宽松数值判断：与 blueprint-engine 的 num() 一致，兼容 excel_read_range 返回的字符串数字列
        const isNum = (h) => rows.some(r => { const v = r[h]; return v !== undefined && v !== null && v !== '' && Number.isFinite(Number(v)); });
        const isDate = (h) => rows.some(r => r[h] !== undefined && r[h] !== null && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(String(r[h])));
        const isId = (h) => /(编号|ID|id|单号|订单号)$/.test(String(h));
        for (const h of headers) {
            if (isDate(h) || isId(h) || isNum(h)) { console.log('  路径2 SKIP ['+h+']'); continue; }
            const vals = rows.map(r => (r[h] !== undefined && r[h] !== null && String(r[h]) !== '') ? String(r[h]) : null).filter(Boolean);
            const uniq = new Set(vals).size;
            console.log('  路径2 ['+h+'] uniq='+uniq+' cap='+uniqCap);
            if (uniq >= 2 && uniq <= uniqCap) dims.push(h);
            if (dims.length >= 3) break;
        }
        console.log('[dash-excel] _resolveSlicerDims 路径2 dims=' + JSON.stringify(dims));
        return dims.slice(0, 4);
    }

    /** 图表对应的透视表数值列：聚合列名不存在时退回首个数值列（count 型图表） */
    function _pivotValueFor(cd, prepared) {
        const src = (prepared.sources && prepared.sources[0]) || null;
        const headers = src && Array.isArray(src.headers) ? src.headers : [];
        const rows = src && Array.isArray(src.rows) ? src.rows : [];
        const vals = cd && Array.isArray(cd.vals) ? cd.vals : [];
        for (const v of vals) if (headers.indexOf(v) >= 0) return v;
        // 宽松数值判断：兼容字符串数字列（与 blueprint-engine num() 一致）
        const num = headers.find(h => rows.some(r => { const v = r[h]; return v !== undefined && v !== null && v !== '' && Number.isFinite(Number(v)); }));
        return num || (headers[0] || '');
    }

    /** 表头是否真实（非"字段1/Column1"占位）：占位表头下透视缓存字段名无意义，禁用切片器路线 */
    function _hasRealHeaders(prepared) {
        const src = (prepared.sources && prepared.sources[0]) || null;
        const headers = src && Array.isArray(src.headers) ? src.headers : [];
        if (!headers.length) return false;
        const placeholder = /^(字段|列|Column|col|field)\s*\d+$/i;
        return headers.some(h => !placeholder.test(String(h).trim()));
    }

    /** 图表类型确定性兜底：分类维度→柱状、日期维度→折线（不被 LLM 带偏成折线/雷达） */
    function _deterministicChartType(cd, prepared) {
        let t = cd && cd.chartType ? String(cd.chartType) : '';
        const src = (prepared.sources && prepared.sources[0]) || null;
        const dim = cd && cd.dim;
        const rows = src && Array.isArray(src.rows) ? src.rows : [];
        const isDateDim = !!(dim && rows.some(r => r[dim] !== undefined && r[dim] !== null && /^\d{4}[-/]\d{1,2}[-/]\d{1,2}/.test(String(r[dim]))));
        const isCategory = !!(dim && !isDateDim);
        if (/xlLine|xlArea|xlRadar/.test(t) && isCategory) return 'xlColumnClustered';
        if (/xlColumn/.test(t) && isDateDim) return 'xlLine';
        if (!t) return isDateDim ? 'xlLine' : 'xlColumnClustered';
        return t;
    }

    /** 数值列集合（透视表值字段白名单）：宽松判断，兼容字符串数字列 */
    function _numericHeaders(prepared) {
        const src = (prepared.sources && prepared.sources[0]) || null;
        if (!src || !Array.isArray(src.rows) || !src.rows.length) return [];
        const rows = src.rows;
        return (Array.isArray(src.headers) ? src.headers : []).filter(h => rows.some(r => {
            const v = r[h];
            return v !== undefined && v !== null && v !== '' && Number.isFinite(Number(v));
        }));
    }

    function buildExcelPlan(prepared, spec) {
        console.log('[dash-excel] buildExcelPlan v20260919-shapes-force | kpis=' + (prepared.kpis||[]).length + ' charts=' + (prepared.charts||[]).length);
        const theme = _themeOf(spec);
        const layout = _LAYOUT();
        const sheet = spec.sheet || '看板';
        const steps = [];
        const DASH_COLS = layout.dashboardCols || 20;
        const lastCol = _colLetter(DASH_COLS);
        const title = spec.title || '数据看板';
        const subtitle = spec.subtitle || '';
        const xl = theme.excel;
        const shapes = _canUseShapes();   // 运行时探测形状工具；不可用则回退纯单元格

        // ── 清理残留：删除旧透视缓存表（含 _dash_pivot）
        //   二次生成若不删，残留 Pivot1 会让 CreatePivotTable 抛 0x800A03EC，
        //   且看板 Cells.Clear 清不掉旧切片器形状 → 界面像"假切片"。
        steps.push({
            title: '清理残留', detail: '删除 _dash_pivot / 透视缓存',
            tool: 'excel_delete_sheet', args: { name: '_dash_pivot', ignoreNotExists: true },
            displayName: '清理残留', displayArgs: { '目标': '_dash_pivot' }
        });
        steps.push({
            title: '清理残留', detail: '删除旧版透视缓存表',
            tool: 'excel_delete_sheet', args: { name: '透视缓存', ignoreNotExists: true },
            displayName: '清理残留', displayArgs: { '目标': '透视缓存' }
        });

        // ── 创建看板工作表（后端同名自动删重建，旧图表浮层/内容一并清掉）──
steps.push({
title: '创建看板工作表', detail: sheet,
tool: 'excel_create_sheet', args: { name: sheet },
displayName: '创建看板工作表', displayArgs: { '工作表名': sheet }
});

// ── 统一列宽 9 + 隐藏尾部空列 ──
steps.push({
title: '看板列宽优化', detail: 'A-' + lastCol + ' 统一 9，尾部隐藏',
tool: 'excel_format_range', args: {
sheet, range: 'A:' + lastCol,
format: JSON.stringify({ columnWidth: 9 })
},
displayName: '看板列宽', displayArgs: { '宽度': '9' }
});
// ── KPI 行高：只设看板区域 B4:Q6，避免整行遍历 16000 列导致卡顿 ──
steps.push({
title: 'KPI 行高', detail: 'B4:Q6 每行22磅',
tool: 'excel_format_range', args: {
sheet, range: 'B4:Q6',
format: JSON.stringify({ rowHeight: 22 })
},
displayName: 'KPI 行高', displayArgs: { '行高': '22pt' }
});
// ── 图表区行高：只设 B8:Q32，避免整行遍历卡顿 ──
steps.push({
title: '图表区行高', detail: 'B8:Q32 每行18磅',
tool: 'excel_format_range', args: {
sheet, range: 'B8:Q32',
format: JSON.stringify({ rowHeight: 18 })
},
displayName: '图表区行高', displayArgs: { '行高': '18pt' }
});
// 隐藏尾部空列/空行: VB 终极清理已用 EntireColumn.Hidden/EntireRow.Hidden 完成
// 禁止用 excel_format_range 处理 N1:XFD (16000+ 列) 或 A80:A1048576 (100万+ 行) → DISP_E_OVERFLOW

        // ── 标题横幅：shapes=true 纯形状画布（不写单元格文字）；shapes=false 单元格回退 ──
        const bannerRange = 'A1:' + lastCol + '3';
        if (shapes) {
            const _titleLines = [{ text: title || '', size: 20, color: '#FFFFFF', bold: true, alignH: 'center' }];
            if (subtitle) _titleLines.push({ text: subtitle, size: 10, color: '#E8EDF6', bold: false, alignH: 'center' });
            steps.push({
                title: '标题横幅', detail: '圆角形状 · 主题底色白字',
                tool: 'excel_add_shape', args: {
                    sheet, range: bannerRange,
                    shapeType: 'roundedRect', cornerRadius: 0.08,
                    fill: xl.title, line: 'none', noShadow: true,
                    lines: JSON.stringify(_titleLines)
                },
                displayName: '标题横幅', displayArgs: { '底色': xl.title, '范围': bannerRange }
            });
        } else {
            steps.push({
                title: '写入看板标题', detail: title,
                tool: 'excel_write_range', args: {
                    sheet, range: 'A1',
                    values: JSON.stringify([[title], [''], [subtitle || '']]), writeMode: 'overwrite'
                },
                displayName: '写入看板标题', displayArgs: { '标题': title, '副标题': subtitle || '无' }
            });
            steps.push({
                title: '标题横幅样式', detail: '跨列居中 · 主题底色白字',
                tool: 'excel_format_range', args: {
                    sheet, range: 'A1:' + lastCol + '2',
                    format: JSON.stringify({ bold: true, fontSize: 20, fontColor: '#FFFFFF', bgColor: xl.title, bgScope: 'all', rowHeight: 30 }),
                    alignment: JSON.stringify({ horizontal: 'centerAcross', vertical: 'middle' })
                },
                displayName: '标题横幅样式', displayArgs: { '底色': xl.title, '字号': 20, '对齐': '跨列居中' }
            });
            steps.push({
                title: '副标题样式', detail: subtitle || '无副标题',
                tool: 'excel_format_range', args: {
                    sheet, range: 'A3:' + lastCol + '3',
                    format: JSON.stringify({ fontSize: 10, fontColor: xl.label, bold: false, rowHeight: 18 }),
                    alignment: JSON.stringify({ horizontal: 'left', vertical: 'middle' })
                },
                displayName: '副标题样式', displayArgs: subtitle ? { '副标题': subtitle } : { '副标题': '无' }
            });
        }

        // ── KPI 指标卡：shapes=true 纯形状画布（每卡一个圆角形状，不写单元格文字）；shapes=false 单元格回退 ──
        const kpis = prepared.kpis || [];
        const kpiCount = kpis.length;
        const labelRow = layout.kpiLabelRow;
        const valueRow = layout.kpiValueRow;
        const cardEndRow = layout.kpiCardEndRow || (valueRow + 1);
        const cardN = Math.max(1, kpiCount);
        const cardW = Math.max(1, Math.ceil(DASH_COLS / cardN));
        const bandOffset = 0; // KPI 卡贴左边（零间隙）
        if (shapes) {
            // KPI 卡片配色：浅灰蓝底 + 细边框 + 扁平无阴影（现代 BI 卡片风格）
            const _kpiFill = xl.kpiFill || '#F2F6FB';
            const _kpiLine = xl.kpiLine || '#D6E4F0';
            const _kpiValueColor = xl.value || '#1F4E79';
            const _kpiLabelColor = xl.label || '#6B7B8D';
            kpis.forEach((k, i) => {
                const name = k.label || ('指标 ' + (i + 1));
                const c1 = _colLetter(bandOffset + i * cardW + 1);
                const c2Idx = (i === cardN - 1) ? DASH_COLS : (bandOffset + (i + 1) * cardW);
                const blockRange = c1 + labelRow + ':' + _colLetter(c2Idx) + cardEndRow;
                const _val = k.numeric ? _thousands(k.value, k.func === 'avg' ? 2 : 0) : String(k.value);
                steps.push({
                    title: `卡片「${name}」`, detail: '圆角形状 · 数值+标签',
                    tool: 'excel_add_shape', args: {
                        sheet, range: blockRange,
                        shapeType: 'roundedRect', cornerRadius: 0.08,
                        fill: _kpiFill, line: _kpiLine, lineWeight: 1.0, noShadow: true,
                        lines: JSON.stringify([
                            { text: _val, size: 16, color: _kpiValueColor, bold: true, alignH: 'left' },
                            { text: (k.label || '') + (k.unit ? '（' + k.unit + '）' : ''), size: 9, color: _kpiLabelColor, bold: false, alignH: 'left' }
                        ])
                    },
                    displayName: `卡片「${name}」`, displayArgs: { '范围': blockRange, '数值': _val }
                });
            });
        } else {
            const labelMatrix = new Array(DASH_COLS).fill('');
            const valueMatrix = new Array(DASH_COLS).fill('');
            kpis.forEach((k, i) => {
                const c0 = bandOffset + i * cardW;
                labelMatrix[c0] = k.label + (k.unit ? '（' + k.unit + '）' : '');
                valueMatrix[c0] = k.numeric ? k.value : String(k.value);
            });
            steps.push({
                title: '写入 KPI 指标卡', detail: kpiCount + ' 个指标',
                tool: 'excel_write_range', args: {
                    sheet, range: 'A' + labelRow,
                    values: JSON.stringify([labelMatrix, valueMatrix]), writeMode: 'overwrite'
                },
                displayName: '写入 KPI 指标卡', displayArgs: { '指标数': kpiCount, '指标': kpis.map(k => k.label + (k.unit ? '(' + k.unit + ')' : '')).join('、') }
            });
            kpis.forEach((k, i) => {
                const name = k.label || ('指标 ' + (i + 1));
                const c1 = _colLetter(bandOffset + i * cardW + 1);
                const c2Idx = (i === cardN - 1) ? DASH_COLS : (bandOffset + (i + 1) * cardW);
                const c2 = _colLetter(c2Idx);
                const labRange = c1 + labelRow + ':' + c2 + labelRow;
                const valRange = c1 + valueRow + ':' + c2 + cardEndRow;
                const blockRange = c1 + labelRow + ':' + c2 + cardEndRow;
                steps.push({
                    title: `卡片「${name}」标签样式`, detail: '小号灰字 · 跨列居中',
                    tool: 'excel_format_range', args: {
                        sheet, range: labRange,
                        format: JSON.stringify({ fontSize: 10, fontColor: xl.label, bold: false }),
                        alignment: JSON.stringify({ horizontal: 'centerAcross', vertical: 'middle' })
                    },
                    displayName: `卡片「${name}」标签样式`, displayArgs: { '范围': labRange, '对齐': '跨列居中' }
                });
                const valFmt = { bold: true, fontSize: 18, fontColor: xl.value, rowHeight: 32 };
                if (k.numeric) valFmt.numberFormat = k.func === 'avg' ? '#,##0.00' : '#,##0';
                steps.push({
                    title: `卡片「${name}」数值样式`, detail: '大号主题色加粗 · 跨列居中',
                    tool: 'excel_format_range', args: {
                        sheet, range: valRange,
                        format: JSON.stringify(valFmt),
                        alignment: JSON.stringify({ horizontal: 'centerAcross', vertical: 'middle' })
                    },
                    displayName: `卡片「${name}」数值样式`, displayArgs: { '范围': valRange, '数字格式': valFmt.numberFormat || '文本' }
                });
                steps.push({
                    title: `卡片「${name}」容器样式`, detail: '底色 · 边框',
                    tool: 'excel_format_range', args: {
                        sheet, range: blockRange,
                        format: JSON.stringify({ bgColor: xl.headerBg, bgScope: 'all', borderStyle: 'thin', borderColor: xl.border }),
                        alignment: JSON.stringify({ horizontal: 'center', vertical: 'middle' })
                    },
                    displayName: `卡片「${name}」容器样式`, displayArgs: { '底色': xl.headerBg, '范围': blockRange }
                });
            });
        }
        // 列宽设置跳过（形状浮层不依赖单元格列宽，避免 format_range 对单行范围报错）

        // ── 图表与切片器前置判定（决定左栏是否让位给筛选栏、图表/表格起始列）──
        const charts = prepared.charts || [];
        const srcRange = _sourceRange(prepared, spec);
        const slicerDims = _resolveSlicerDims(prepared, spec);
        const numHeaders = _numericHeaders(prepared);
        console.log('[dash-excel] useSlicer 诊断: srcRange=' + JSON.stringify(srcRange) +
            ' | slicerDims=' + JSON.stringify(slicerDims) +
            ' | slicerDims.length=' + (slicerDims ? slicerDims.length : 0) +
            ' | _hasRealHeaders=' + _hasRealHeaders(prepared) +
            ' | prepared.sources[0]=' + JSON.stringify(prepared.sources && prepared.sources[0] ? { range: prepared.sources[0].range, sheet: prepared.sources[0].sheet, worksheet: prepared.sources[0].worksheet, headersLen: (prepared.sources[0].headers||[]).length } : 'null') +
            ' | spec.dataSources[0]=' + JSON.stringify(spec.dataSources && spec.dataSources[0] ? spec.dataSources[0] : 'null'));
        // pivotSheet 指向独立的隐藏工作表——
        // CreateSlicer 需要一个干净的地方放 4 个透视表（不干扰看板上的标题、KPI、图表），
        // 同时 chartArgs.pivotSheet 也必须精确指向这个 sheet，否则 CreateChart 找不到透视表
        // 就会静默回退到 dataValues，切片器不联动。
        const pivotSheet = spec.pivotSheet || '_dash_pivot';
        const pivotRowStep = layout.pivotRowStep || 20;

        // 先创建/清空 pivotSheet（CreateSlicer 会在上面建 4 个透视表）
        steps.push({
            title: '创建透视缓存', detail: pivotSheet + '（独立工作表，避免覆盖看板布局）',
            tool: 'excel_create_sheet', args: { name: pivotSheet },
            displayName: '创建透视缓存', displayArgs: { '工作表': pivotSheet }
        });
        // 启用原生切片器：有数据源范围 + 有切片器维度 + 真实表头时，走透视表+透视图（可刷新、可联动筛选）。
        // UI 垃圾清理（"求和项"副标题/绿色字段按钮）已下放到 VB 端 PivotChart 创建后处理。
        const useSlicer = !!(srcRange && slicerDims && slicerDims.length && _hasRealHeaders(prepared));
        // 左栏宽度：仅在真正启用切片器时为筛选栏留白（好放切片器），否则全宽铺满消除大片空档
        const _gridBase = useSlicer ? (layout.centerStartCol || 5) : 2;
        // ⚠ 必须读 LAYOUT.chartSlotCols！别用 dashboardCols 硬算（会覆盖 Slicer 位置）
        const _slotCols = layout.chartSlotCols !== undefined ? layout.chartSlotCols : Math.max(4, Math.floor(((layout.dashboardCols || 20) - _gridBase + 1) / 2));
        const _slotRows = layout.chartSlotRows || 11;   // 槽位行数（直接用 LAYOUT 原值，不硬截断）
        const _rowGap = layout.chartRowGap || (_slotRows + 1);  // 默认槽高+1 行间距，避免硬截断

        // ── 分节标题已删！零间隙：KPI(行4-7) 直接贴紧 图表(行8-16) ──

        // ── Excel 原生切片器（共享透视缓存 + 透视表 + SlicerCache/Slicers）──
        // 一步到位：CreateSlicer 创建每张透视表后立刻 CreatePivotChart，
        // 没有中间"pivotSheet + pivotTable 名字字符串对接"的脆弱环节。
        if (useSlicer) {
            // 先算好 2×2 网格 anchor（提前算，后面 pivotTables map 需要）
            const _dashCols = layout.dashboardCols || 20;
            charts.forEach((cd, i) => {
                const slot = Math.floor(i / 2);
                const _colGap = layout.chartColGap !== undefined ? layout.chartColGap : 1;
                const colOff = (i % 2) * (_slotCols + _colGap);
                const row = layout.chartStartRow + slot * _rowGap;
                const colA = _colLetter(_gridBase + colOff);
                const isRightSlot = (i % 2) === 1;
                const colB = _colLetter(_gridBase + _slotCols + colOff - 1);
                cd._anchor = sheet + '!' + colA + row + ':' + colB + (row + _slotRows - 1);
            });

            const pivotTables = charts.map((cd, i) => {
                const isPie = _isCategory(cd.chartType);
                // ── 维度字段最终兜底：cd.dim 不在 headers → 模糊匹配或自动推断 ──
                let dim = cd.dim || '';
                const hdr = (prepared.sources && prepared.sources[0] && prepared.sources[0].headers) || [];
                if (dim && hdr.length && hdr.indexOf(dim) < 0) {
                    let best = null; let bestScore = 0;
                    const tgt = dim.toLowerCase();
                    for (const h of hdr) {
                        const hl = String(h).toLowerCase();
                        let score = 0;
                        if (hl === tgt) score = 100;
                        else if (hl.includes(tgt) || tgt.includes(hl)) score = Math.max(hl.length, tgt.length);
                        if (score > bestScore) { bestScore = score; best = h; }
                    }
                    if (best && bestScore >= 2) dim = best;
                }
                if (!dim && hdr.length) dim = hdr[0];
                return {
                    dimension: dim,
                    value: _pivotValueFor(cd, prepared),
                    func: cd.func || 'sum',
                    targetCell: 'A' + (1 + i * pivotRowStep),
                    name: 'Pivot' + (i + 1),
                    // 一步到位：每张透视表 entry 直接带 PivotChart 配置
                    chartType: _deterministicChartType(cd, prepared),
                    chartTitle: cd.title || '图表 ' + (i + 1),
                    anchor: cd._anchor,
                    anchorSheet: sheet,
                    chartStyle: 201,
                    hasLegend: !isPie,
                    showDataLabels: isPie,
                    dataLabelPosition: isPie ? 'xlLabelPositionBestFit' : /xlLine/.test(cd.chartType) ? 'xlLabelPositionAbove' : 'xlLabelPositionOutsideEnd'
                };
            });
            // 切片器浮层：纵排于图表区左侧，垂直堆叠，与图表行对齐
            // ⚠ 必须用 !==undefined 而不是 || ！0 是 falsy → || 会把 0 变成 fallback
            const slicerBaseTop = layout.slicerTopPx !== undefined ? layout.slicerTopPx : ((layout.chartStartRow || 9) * 15 + 6);
            const slicerH = layout.slicerHeight !== undefined ? layout.slicerHeight : 120;
            const slicerGap = layout.slicerGap !== undefined ? layout.slicerGap : 4;
            const slicerDimensions = slicerDims.map((h, i) => ({
                header: h,
                top: slicerBaseTop + i * (slicerH + slicerGap),
                left: layout.slicerLeftStart !== undefined ? layout.slicerLeftStart : 693,
                width: layout.slicerWidth !== undefined ? layout.slicerWidth : 288.75,
                height: slicerH
            }));
            steps.push({
                title: '创建原生切片器', detail: slicerDims.join('、') + '（联动 ' + charts.length + ' 张图表）',
                tool: 'excel_create_slicer', args: {
                    sourceRange: srcRange, sheet, pivotSheet,
                    slicerDimensions: JSON.stringify(slicerDimensions),
                    pivotTables: JSON.stringify(pivotTables)
                },
                displayName: '创建原生切片器', displayArgs: { '维度': slicerDims.join('、'), '联动图表数': charts.length, '数据源': srcRange }
            });
        }

        // ── 图表卡（2 列网格）──
        // useSlicer=true 时跳过——CreateSlicer 一步到位已创建透视表 + PivotChart，
        // 不再单独调 excel_create_chart（避免重复创建，且彻底消除 pivotSheet/pivotTable 名字对接的脆弱环节）
        // useSlicer=false 时正常走 create_chart + dataValues（不落地中间表）
        if (!useSlicer) {
            const _dashCols2 = layout.dashboardCols || 20;
            charts.forEach((cd, i) => {
                const slot = Math.floor(i / 2);
                const _colGap = layout.chartColGap !== undefined ? layout.chartColGap : 1;
                const colOff = (i % 2) * (_slotCols + _colGap);
                const row = layout.chartStartRow + slot * _rowGap;
                const _base = _gridBase;
                const colA = _colLetter(_base + colOff);
                const isRightSlot = (i % 2) === 1;
                const colB = isRightSlot
                    ? _colLetter(_dashCols2)
                    : _colLetter(_base + _slotCols + colOff - 1);
                const anchor = sheet + '!' + colA + row + ':' + colB + (row + _slotRows - 1);
                const isPie = _isCategory(cd.chartType);
                const chartArgs = {
                    chartType: _deterministicChartType(cd, prepared),
                    title: cd.title || '图表 ' + (i + 1),
                    sheet,
                    position: JSON.stringify({ anchorRange: anchor }),
                    seriesBy: 'columns',
                    chartStyle: 201,
                    hasLegend: !isPie,
                    showDataLabels: isPie,
                    dataLabelPosition: isPie ? 'xlLabelPositionBestFit' : /xlLine/.test(cd.chartType) ? 'xlLabelPositionAbove' : 'xlLabelPositionOutsideEnd',
                    // 【看板多图必须显式声明】后端 multiChart 默认 False 时，每次建图前会清理
                    // 本表所有 AIChart_ 前缀图 → 多张图互相覆盖，最终只剩最后一张。
                    multiChart: true
                };
                const _memData = JSON.stringify([cd.header, ...cd.data]);
                chartArgs.dataValues = _memData;
                if (!isPie) { chartArgs.xAxisTitle = cd.dim; chartArgs.yAxisTitle = cd.vals[0]; }
                steps.push({
                    title: `生成图表「${cd.title || '图表 ' + (i + 1)}」`,
                    detail: _humanize(cd.chartType),
                    tool: 'excel_create_chart', args: chartArgs,
                    _dashFallback: { dataValues: _memData },
                    displayName: `生成图表「${cd.title || '图表 ' + (i + 1)}」`,
                    displayArgs: { '类型': _humanize(cd.chartType), '维度': cd.dim, '指标': cd.vals.join('、'), '位置': anchor }
                });
            });
        }

        // ── 明细表（紧随中央图表区下方，同样从中央栏起始列起，避开左侧筛选栏）──
        // 不套 shape 面板：Excel Shape 永远浮在单元格之上，无法做真实表格的装饰框，
        // 且透明边框在几十行的真实数据表上视觉上就是一条丑线。
// 明细表已移除: 看板只保留横幅+KPI+图表+Slicer, 底部不再写源数据表
        // 把工作表名拼到范围地址里（看板!A1:T3），后端解析范围时直接拿到工作表，不会 fallback 到原始数据表

        return steps;
    }

    ChatLog.info?.('CHAT-V2', '✅ dashboardExcelRenderer loaded');
    App.ChatV2.dashboardModules.excelRenderer = { buildExcelPlan };
})();
