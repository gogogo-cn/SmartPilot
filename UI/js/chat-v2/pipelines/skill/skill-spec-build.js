/**
 * ============================================================
 * skill-spec-build.js — Skill 管道 spec 构建函数（从 skill-pipeline.js 拆出）
 * ------------------------------------------------------------
 * 职责：从用户输入/LLM plan 构造 skill 执行步骤（buildSpecSteps）、显式区域解析、
 *       选区自动探测、美化操作映射（buildBeautifyOps）。
 * 挂载：App.ChatV2.skillPipelineModules.specBuild
 * 行为零变更重构：函数体自 skill-pipeline.js 逐字迁移，未改动任何逻辑/正则/文案。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.skillPipelineModules) App.ChatV2.skillPipelineModules = {};
    const MODS = App.ChatV2.skillPipelineModules;
    const SP = MODS;
    const C = MODS.constants;
    const ChatLog = window.ChatLog || console;
    /**
     * Word 宿主的"美化 plan → 有序 Word 工具操作"。
     * 背景：Word 宿主既没有 excel_format_range，也没有 VBA 执行能力。
     * 若沿用 Excel 的 op 映射，skill-spec-execute 的"工具可用性"校验会失败 →
     * 回退 Excel VBA 兜底 → Word 侧 ExecuteVBA 恒返回 not supported →
     * skill-pipeline 判定 wasForced 后降级 agent-loop 自由调用细粒度 word_* 工具，
     * 表现为"步骤重复、流程缓慢、与 Excel 效果不一致"。
     * 这里直接产出 Word 自己的整表工具 op，让 Word 也走"每步=真实工具"的确定性骨架。
     * 注意：不依赖 A1 区域（Word 无 range 概念），全部作用于"光标当前所在表格"。
     */
    function buildBeautifyOpsForWord(plan, se) {
        if (!plan || typeof plan !== 'object') return null;
        const theme = (se && se.THEMES ? (se.THEMES[plan.theme] || se.THEMES[se.DEFAULT_SPEC?.theme]) : null) || {};
        const headerRGB = plan.headerColor || theme.header || [0, 90, 156];
        const bodyRGB = plan.bodyColor || theme.body || theme.header || null;
        const borderRGB = theme.border || [180, 200, 220];
        const themeName = (plan.headerColor ? '自定义色彩' : ((se && se.listThemes) ? (se.listThemes().find(t => t.id === plan.theme)?.name || plan.theme) : plan.theme));
        const headerHex = SP.util.rgbToHex(headerRGB);
        const borderHex = SP.util.rgbToHex(borderRGB);
        const bodyHex = bodyRGB ? SP.util.rgbToHex(bodyRGB) : null;
        const headerRows = Math.max(1, Math.min(5, Number(plan.headerRows) || 1));
        const WT = 'word_format_selected_table_cell';
        const wops = [];
        // 主题：Word 无主题概念，落成整表边框色（与 Excel 的"主题"op 语义对齐）
        if (plan.theme) {
            wops.push({
                label: `主题「${themeName}」`, tool: WT,
                args: { borderColor: borderHex || '#B4C8DC' },
                showArgs: { 主题: themeName }
            });
        }
        // 表头：加粗 + 主题色底纹（Word 整表工具的 headerFillColor 作用于首行）
        if (plan.headerStyle) {
            const ha = { headerBold: true };
            if (headerHex) ha.headerFillColor = headerHex;
            wops.push({
                label: '设置表头', tool: WT, args: ha,
                showArgs: { 表头: headerRows > 1 ? `前${headerRows}行（Word 工具作用于首行）` : '首行' }
            });
        }
        // 内容区底色：Word 整表工具没有"仅内容区 range"的概念；清除底纹可做，
        // 上底色若要落到内容区需宿主侧扩展（暂不在此误把底色刷到表头）。
        if (plan.noContentBg) {
            wops.push({ label: '清除底纹', tool: WT, args: { clearFill: true }, showArgs: { 内容底色: '清除' } });
        } else if (bodyHex) {
            ChatLog.info?.('PIPELINE', '[spec-build] Word 整表工具暂不支持"仅内容区"底色，已跳过该步（避免误刷表头）');
        }
        // 字体：字号 + 字体名（作用于整表所有单元格）
        if (plan.fontName || plan.fontSize) {
            const fa = {};
            if (plan.fontName) fa.fontName = plan.fontName;
            if (plan.fontSize) fa.fontSize = plan.fontSize;
            wops.push({
                label: `统一字体 ${plan.fontName || '默认'} ${plan.fontSize || 11}号`, tool: WT, args: fa,
                showArgs: { 字体: plan.fontName || '默认', 字号: plan.fontSize || 11 }
            });
        }
        // 对齐：Word 整表工具的 align 作用于所有单元格（与 Excel 拆表头/内容两 op 略不同）
        if (plan.alignment) {
            wops.push({
                label: `对齐方式 ${plan.alignment}`, tool: WT, args: { align: plan.alignment },
                showArgs: { 对齐: plan.alignment }
            });
        }
        // 边框：整表边框色
        if (plan.addBorders) {
            wops.push({
                label: '添加表格边框', tool: WT, args: { borderColor: borderHex || '#B4C8DC' },
                showArgs: { 边框: '全边框' }
            });
        }
        return wops.length ? wops : null;
    }

    /** 把美化 plan 映射成"有序真实工具操作"。仅支持核心公式化操作，含其它选项则返回空 → 回退 VBA。 */
    function buildBeautifyOps(plan, ctx) {
        if (!plan || typeof plan !== 'object' || !ctx) return null;
        const se = (typeof specEngine !== 'undefined' && specEngine) ? specEngine : (App?.ChatV2?.services?.skillSpecEngine || null);
        if (!se) return null;
        // Word 宿主必须在此分流：Word 没有 A1 区域概念（下方 addr 解析大概率拿不到值会直接 return null），
        // 也没有 excel_format_range / VBA 能力。放在最前面可避免整条流程回退 Excel VBA 后降级。
        if (typeof window !== 'undefined' && window.SmartPilotHost && window.SmartPilotHost.isWord) {
            return buildBeautifyOpsForWord(plan, se);
        }
        const sel = ctx.selection || {};
        // 【选区优先】用户已框选多格数据区时，以选中区域为准（表头=其首行），
        // 不能用 plan.dataRange（LLM 自拟）覆盖——LLM 常按上下文自拟整表/首行范围，
        // 导致"按选中区域美化"时把表头格式错设到非选中区域（实测第一条消息错误）。
        // 仅在无真实多格选区时才回退 plan.dataRange / 单格选区。
        const selAddr = String(sel.address || sel.range || '').replace(/\$/g, '');
        const selRows = Number(sel.rowCount) || 0;
        const selCols = Number(sel.colCount) || 0;
        const hasRegion = !!(selAddr) && (selRows > 1 || selCols > 1);
        const addr = hasRegion
            ? selAddr
            : (plan.dataRange ? String(plan.dataRange).replace(/\$/g, '') : selAddr);
        if (!addr) return null;
        // 【选区范围判断·降级】最终目标只是"光标单格"（A1/G30）时，工具路径不下场：
        // rowRangeOf(单格) 只能格式化 1 个格子，属于"假成功"（实测 2026-10-22 批量美化
        // 4 张表只美化了 G30 单格）。返回 null 让调用方走 VBA 路径——它以 Selection 为基准
        // 向上扩展连续非空行 + 识别表头高度，单格也能定位到真实数据区。
        // 单一事实源：services/range-resolver.js 的 isSingleCellAddr（模块缺失时本地兜底）。
        const _rr = (typeof App !== 'undefined' && App.ChatV2 && App.ChatV2.rangeResolver) || null;
        const _addrNorm = addr.replace(/\$/g, '').trim().replace(/^[^!]*!/, '').trim();
        const _cursorOnly = _rr && typeof _rr.isSingleCellAddr === 'function'
            ? _rr.isSingleCellAddr(addr)
            : (!_addrNorm || (_addrNorm.indexOf(':') < 0 && _addrNorm.indexOf(',') < 0));
        if (_cursorOnly) {
            ChatLog.info?.('PIPELINE', `[skill-spec] 美化目标只是光标单格（${_addrNorm || '空'}）→ 工具路径弃权，回退 VBA 选区扩展`);
            return null;
        }
        // 仅核心公式化操作走工具框架，保持骨架与操作完全对齐；其余选项回退 VBA 保证正确
        // （排版维度：版式/表头/行高/首列/末行/数字右对齐 都只能由 VBA 渲染，必须回退）
        const _layoutOnly = !!(plan.stylePreset || plan.emphasizeLastRow || plan.boldFirstColumn ||
            plan.bodyNumberRight || plan.rowHeight ||
            (plan.borderStyle && plan.borderStyle !== 'all') ||
            (plan.headerLayout && plan.headerLayout !== 'deep'));
        const coreOnly = !_layoutOnly && !plan.freezeHeader && !plan.bandedRows && !plan.wrapText && !plan.conditionalFormat && !plan.numberFormat && !plan.dateFormat && !plan.title;
        if (!coreOnly) return null;
        const sheet = sel.sheetName || '';
        const theme = (se.THEMES ? (se.THEMES[plan.theme] || se.THEMES[se.DEFAULT_SPEC?.theme]) : null) || {};
        // 表头颜色：自定义优先，否则用主题。表身底色只在用户点名时才有 plan.bodyColor，不从主题回填。
        const headerRGB = plan.headerColor || theme.header || [0, 90, 156];
        const bodyRGB = plan.bodyColor || null;
        const borderRGB = theme.border || [180, 200, 220];
        const themeName = (plan.headerColor ? '自定义色彩' : (se.listThemes ? (se.listThemes().find(t => t.id === plan.theme)?.name || plan.theme) : plan.theme));
        const headerHex = SP.util.rgbToHex(headerRGB);
        const borderHex = SP.util.rgbToHex(borderRGB);
        const bodyHex = bodyRGB ? SP.util.rgbToHex(bodyRGB) : null;
        // 多行表头：headerRows 由 LLM 依据选中区数据自动判定（含合并标题行），
        // 表头样式统一套到前 headerRows 行；数据区自 headerRows+1 行起。
        const headerRows = Math.max(1, Math.min(5, Number(plan.headerRows) || 1));
        const headerRange = SP.util.rowRangeOf(addr, 1, headerRows);
        // 数据区（表头以下）：算出 addr 总行数，取 headerRows+1..末尾；为内容区补主题浅底色（对齐 VBA 路径 palette.body）。
        const am = addr.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
        const totalRows = am ? (Number(am[4]) - Number(am[2]) + 1) : 1;
        const dataRange = totalRows > headerRows ? SP.util.rowRangeOf(addr, headerRows + 1, totalRows) : null;
        const fmtMerge = o => Object.assign({}, o.sheet ? { sheet: o.sheet } : {}, o.range ? { range: o.range } : {});
        const base = fmtMerge({ sheet, range: addr });
        const ops = [];
        if (plan.theme) ops.push({
            label: `主题「${themeName}」`, tool: 'excel_format_range',
            args: Object.assign({}, base, { format: JSON.stringify({ borderStyle: 'xlContinuous', borderColor: borderHex || '#B4C8DC' }) }),
            showArgs: { 主题: themeName }
        });
        if (plan.headerStyle) {
            const fmt = { bold: true };
            if (headerHex) { fmt.bgColor = headerHex; fmt.fontColor = '#FFFFFF'; fmt.bgScope = 'all'; }
            ops.push({ label: '设置表头', tool: 'excel_format_range', args: Object.assign({}, fmtMerge({ sheet, range: headerRange }), { format: JSON.stringify(fmt) }), showArgs: { 表头: headerRange } });
        }
        if (dataRange) {
            // 用户明确"移除内容底色" → 清为白底；否则按主题/自定义 body 上浅底色（对齐 VBA 路径）。
            if (plan.noContentBg) {
                ops.push({ label: '去除内容底色', tool: 'excel_format_range', args: Object.assign({}, fmtMerge({ sheet, range: dataRange }), { format: JSON.stringify({ bgColor: '#FFFFFF', bgScope: 'all' }) }), showArgs: { 内容底色: '清除(白底)' } });
            } else if (bodyHex) {
                ops.push({ label: '内容区底色', tool: 'excel_format_range', args: Object.assign({}, fmtMerge({ sheet, range: dataRange }), { format: JSON.stringify({ bgColor: bodyHex, bgScope: 'all' }) }), showArgs: { 内容底色: dataRange } });
            }
        }
        if (plan.fontName || plan.fontSize) {
            const f = {};
            if (plan.fontName) f.fontName = plan.fontName;
            if (plan.fontSize) f.fontSize = plan.fontSize;
            ops.push({ label: `字体 ${plan.fontName || '默认'} ${plan.fontSize || 11}号`, tool: 'excel_format_range', args: Object.assign({}, base, { format: JSON.stringify(f) }), showArgs: { 字体: plan.fontName || '默认', 字号: plan.fontSize || 11 } });
        }
        if (plan.alignment) {
            // 对齐分表头/内容两块（与 VBA 路径一致：表头固定居中，内容用用户指定的对齐）。
            // 后端 excel_format_range 的 alignment 作用于传入整个 range，故必须拆两个不同 range 调用。
            if (headerRange) {
                ops.push({ label: '表头对齐', tool: 'excel_format_range', args: Object.assign({}, fmtMerge({ sheet, range: headerRange }), { format: '{}', alignment: JSON.stringify({ horizontal: 'center' }) }), showArgs: { 表头对齐: '居中' } });
            }
            if (dataRange) {
                ops.push({ label: '内容对齐', tool: 'excel_format_range', args: Object.assign({}, fmtMerge({ sheet, range: dataRange }), { format: '{}', alignment: JSON.stringify({ horizontal: plan.alignment }) }), showArgs: { 内容对齐: plan.alignment } });
            }
        }
        if (plan.addBorders) {
            ops.push({ label: '添加表格边框', tool: 'excel_format_range', args: Object.assign({}, base, { format: JSON.stringify({ borderStyle: 'xlContinuous', borderColor: borderHex || '#B4C8DC' }) }), showArgs: { 边框: '全边框' } });
        }
        if (plan.autofit) {
            ops.push({ label: '自动调整列宽', tool: 'excel_autofit', args: Object.assign({}, fmtMerge({ sheet, range: addr }), { type: 'column' }), showArgs: { 方式: '按内容' } });
        }
        return ops.length ? ops : null;
    }

    /**
     * 选区自动探测：写型操作（美化/格式）在用户未框选数据区、或选区疑似单格时，
     * 调用无 range 的 excel_read_range，让后端按 RangeContext→Selection→UsedRange 回退自动定位真实数据区，
     * 并回填 ctx.selection 供工具内执行（而非直接回退 VBA）。
     * @returns {Object|null} 探测到的选区 { address, range, sheetName, rowCount, colCount, autoDetected:true }；无需探测 / 失败 / 探测无效返回 null
     */

    /**
     * 从用户输入中软编码解析"显式写死"的目标区域（最高优先级）。
     * 支持 A1:E31 / $A$1:$E$31 / a1:e31 / 单格 A1 等；表格类 skill 的规划阶段也可能把区域放进 plan.dataRange，
     * 但用户有时仅用自然语言描述区域，这里做兜底，保证写死区域不被选区/自动探测覆盖。
     * @param {string|null|undefined} input
     * @returns {{address:string,rowCount:number,colCount:number,explicit:boolean}|null}
     */
    function parseExplicitRange(input) {
        if (typeof input !== 'string' || !input.trim()) return null;
        const t = input.replace(/\s+/g, '');
        // 区域：A1:E31（可带 $，大小写不敏感）
        const m = t.match(/\$?([A-Za-z]{1,3})\$?(\d+):\$?([A-Za-z]{1,3})\$?(\d+)/);
        if (m) {
            const c1 = m[1].toUpperCase(), r1 = Number(m[2]), c2 = m[3].toUpperCase(), r2 = Number(m[4]);
            if (c1 && r1 > 0 && c2 && r2 > 0) {
                return { address: `${c1}${r1}:${c2}${r2}`, rowCount: Math.abs(r2 - r1) + 1, colCount: Math.abs(SP.util._colToNum(c2) - SP.util._colToNum(c1)) + 1, explicit: true };
            }
        }
        // 单格：A1（避免把"第5行"这类误当列+行，需紧跟字母且在数字后有非字母或到头）
        const m1 = t.match(/\$?([A-Za-z]{1,3})\$?(\d+)/);
        const after1 = m1 ? (t[m1.index + m1[0].length] || '') : '';
        if (m1 && (!after1 || !/[A-Za-z]/.test(after1))) {
            const c = m1[1].toUpperCase(), r = Number(m1[2]);
            if (c && r > 0) return { address: `${c}${r}`, rowCount: 1, colCount: 1, explicit: true };
        }
        return null;
    }

    async function detectDataSelection(ctx, turn, toolRegistry, sessionStore, signal) {
        const sel = (ctx && ctx.selection) || {};
        const addr = String(sel.address || sel.range || '').replace(/\$/g, '').trim();
        const r = Number(sel.rowCount) || 0;
        const c = Number(sel.colCount) || 0;
        // 已有明确的非单格数据区 → 不必探测
        if (addr && (r > 1 || c > 1)) return null;
        if (!toolRegistry || typeof toolRegistry.invokeTracked !== 'function') return null;
        const sheet = ctx?.activeSheet || sel.sheetName || '';
        const readArgs = { returnFormat: 'object', hasHeaders: true, autoProbe: true };
        if (sheet) readArgs.sheet = sheet;
        let rr;
        try {
            rr = await toolRegistry.invokeTracked(turn, 'excel_read_range', readArgs, {
                displayName: '自动探测数据区域',
                displayArgs: { '方式': '未框选→自动探测数据区', '工作表': sheet || '活动表' },
                ctx: { turn, sessionStore, signal }
            });
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[probe-selection] 探测读取异常:', e && e.message ? e.message : e);
            return null;
        }
        if (signal?.aborted) throw new Error('user_aborted');
        if (!rr || rr.ok === false || rr.error) return null;
        let meta = rr.result;
        if (typeof meta === 'string') { try { meta = JSON.parse(meta); } catch { meta = null; } }
        const probedAddress = meta && meta.range;
        if (!probedAddress) return null;
        const pRows = Number(meta && meta.rows) || 0;
        const pCols = Number(meta && meta.columns) || 0;
        // 探测结果仍非有效数据区（空/单格）→ 不强用，交给既有回退（美化仍走 VBA UsedRange）
        if (pRows <= 1 && pCols <= 1) return null;
        return Object.assign({}, sel, {
            address: String(probedAddress).replace(/\$/g, ''),
            range: String(probedAddress).replace(/\$/g, ''),
            sheetName: (meta && meta.worksheet) || sel.sheetName || '',
            rowCount: pRows,
            colCount: pCols,
            autoDetected: true,
            // 探测即读取：把本次读取到的数据带回，供调用方（清洗）复用，避免再调一次 excel_read_range
            rawResult: rr.result
        });
    }

    function buildSpecSteps(plan, kind, extra) {
        if (!plan || typeof plan !== 'object') return [];
        // 每步含 { title: 规划步骤名, detail: 实际执行参数 }。
        // 渲染器据此展示"规划步骤"路线 + "步骤N：名称：详情"。
        const steps = [];
        if (kind === 'beautify') {
            if (plan.theme) {
                const _se = (typeof specEngine !== 'undefined' && specEngine) ? specEngine : (App?.ChatV2?.services?.skillSpecEngine || null);
                const t = (_se && typeof _se.listThemes === 'function')
                    ? (_se.listThemes().find(x => x.id === plan.theme)?.name || plan.theme)
                    : plan.theme;
                steps.push({ title: '设置主题', detail: t + '边框', status: 'done' });
            }
            const hs = plan.headerStyle;
            if (hs) {
                const hr = Math.max(1, Math.min(5, Number(plan.headerRows) || 1));
                steps.push({ title: '设置表头', detail: (hs === 'bold_colored' ? '加粗 + 白字 + 主题色填充' : hs) + (hr > 1 ? `（${hr} 行表头）` : ''), status: 'done' });
            }
            if (plan.theme) steps.push({ title: plan.noContentBg ? '清除内容底色' : '内容区底色', detail: plan.noContentBg ? '移除数据区填充(白底)' : '主题浅色调填充·覆盖原遗留底色', status: 'done' });
            if (plan.fontName || plan.fontSize) steps.push({ title: '统一字体', detail: (plan.fontName || '默认字体') + ' ' + (plan.fontSize || '11') + ' 号', status: 'done' });
            if (plan.alignment) steps.push({ title: '对齐方式', detail: '表头居中 · 内容' + (plan.alignment === 'center' ? '居中' : plan.alignment === 'left' ? '左对齐' : '右对齐'), status: 'done' });
            // ── 排版维度（2026-10-07）：让"版式/表头/行高/强调"在步骤面板可见 ──
            if (plan.stylePreset || plan.borderStyle) {
                const _P = App?.ChatV2?.tableStylePresets || null;
                const _bs = plan.borderStyle || 'all';
                const _bsName = { all: '全网格边框', none: '无边框(留白)', horizontal: '仅横向分隔线', outer: '仅外框线', threeLine: '三线表(顶线+表头下线+底线)' }[_bs] || _bs;
                const _presetName = (_P && plan.stylePreset && _P.LAYOUTS && _P.LAYOUTS[plan.stylePreset])
                    ? _P.LAYOUTS[plan.stylePreset].name : '';
                steps.push({
                    title: '应用版式',
                    detail: (_presetName ? _presetName + ' · ' : '') + _bsName,
                    status: 'done'
                });
            }
            const _hl = plan.headerLayout || null;
            if (_hl && _hl !== 'deep') {
                steps.push({
                    title: '表头版式',
                    detail: { deep: '深底白字', light: '浅底深字', underline: '仅底部粗线(无底色)', italic: '斜体表头' }[_hl] || _hl,
                    status: 'done'
                });
            }
            if (plan.rowHeight) steps.push({ title: '行高', detail: plan.rowHeight + ' 磅（' + (plan.rowHeight <= 16 ? '紧凑' : plan.rowHeight >= 26 ? '宽松' : '标准') + '）', status: 'done' });
            if (plan.boldFirstColumn) steps.push({ title: '首列强调', detail: '首列加粗 + 浅底', status: 'done' });
            if (plan.emphasizeLastRow) steps.push({ title: '末行强调', detail: '合计行加粗 + 上边线', status: 'done' });
            if (plan.bodyNumberRight) steps.push({ title: '数字右对齐', detail: '数值单元格靠右', status: 'done' });
            if (plan.addBorders && (plan.borderStyle || 'all') === 'all') steps.push({ title: '添加表格边框', detail: '全边框', status: 'done' });
            if (plan.autofit) steps.push({ title: '自动调整列宽', detail: '按内容自适应', status: 'done' });
            if (plan.freezeHeader) steps.push({ title: '冻结表头', detail: '锁定首行', status: 'done' });
            if (plan.bandedRows) steps.push({ title: '启用斑马纹', detail: '隔行底色', status: 'done' });
            if (plan.wrapText) steps.push({ title: '自动换行', detail: '开启', status: 'done' });
            if (plan.conditionalFormat) steps.push({ title: '条件格式', detail: plan.conditionalFormat, status: 'done' });
            if (plan.numberFormat) steps.push({ title: '数字格式', detail: plan.numberFormat, status: 'done' });
            if (plan.dateFormat) steps.push({ title: '日期格式', detail: plan.dateFormat, status: 'done' });
            if (plan.title) steps.push({ title: '添加标题行', detail: plan.title, status: 'done' });
        } else if (kind === 'clean') {
            const ops = plan.operations || [];
            // 清洗流程第 1 步即为读取数据
            steps.push({ title: '读取数据', detail: plan.dataRange || '当前选中区域', status: 'done' });
            // 如果 ops 为空但 detectCleanDedupeMode 判定了 markDuplicate → 兜底加 removeDuplicates
            // （LLM spec 偶尔会漏 operations 字段）
            if (ops.length === 0 && typeof plan.markDuplicate === 'boolean') {
                ops.push('removeDuplicates');
            }
            // 骨架顺序/标题来自 constants.CLEAN_OPS 注册表（与 executor 排序/标签同一数据源，
            // 原"if 顺序必须一致"的注释约束变成构造约束，顺序/标签改注册表即全局生效）
            const CLEAN_DETAILS = {
                removeDuplicates: (p) => {
                    const cols = Array.isArray(p.dedupeColumns) && p.dedupeColumns.length ? '按 ' + p.dedupeColumns.join('、') + ' 列去重' : '全行去重';
                    const mode = p.markDuplicate === false ? '删除重复行' : p.markDuplicate === true ? '仅高亮标记' : '自动判定';
                    return cols + '（' + mode + '）';
                },
                fillEmpty: (p) => '方法：' + (p.fillMethod === 'zero' ? '填 0' : p.fillMethod === 'delete' ? '删除该行' : p.fillMethod === 'placeholder' ? '占位' : '前值填充'),
                trimSpaces: () => '全部单元格',
                normalizeDates: () => '识别日期列并统一为 yyyy-mm-dd',
                normalizeNumbers: () => '识别数字列并统一为千分位格式',
                standardizeText: () => '统一大小写/全半角',
                removeEmptyRows: () => '整行空白',
                splitColumn: () => '按分隔符拆分',
                mergeColumns: () => '指定列合并'
            };
            for (const opDef of C.CLEAN_OPS) {
                if (ops.includes(opDef.name)) {
                    steps.push({ title: opDef.stepTitle, detail: CLEAN_DETAILS[opDef.name](plan), status: 'done' });
                }
            }
            // ops 仍为空但 plan 不是空 → 兜底"执行清洗"，至少有一个骨架步
            if (steps.length === 0) {
                steps.push({ title: '执行清洗', detail: '智能识别并处理数据质量问题', status: 'done' });
            }
        } else if (kind === 'chart') {
            steps.push({ title: '读取数据', detail: plan.dataRange || '当前选定区域', status: 'done' });
            if (plan.chartType) steps.push({ title: '生成图表', detail: plan.chartType + ' 图', status: 'done' });
            if (plan.title) steps.push({ title: '图表标题', detail: plan.title, status: 'done' });
            if (plan.showLegend) steps.push({ title: '图例', detail: '显示', status: 'done' });
            if (plan.showDataLabels) steps.push({ title: '数据标签', detail: '显示', status: 'done' });
            if (plan.chartStyle) steps.push({ title: '图表样式', detail: plan.chartStyle.toString(), status: 'done' });
        } else if (kind === 'crud') {
            steps.push({ title: SP.cleanCrud.opLabel(plan.operation), detail: SP.cleanCrud.crudWhere(plan), status: 'done' });
        } else if (kind === 'format') {
            // SmartTableFormatter 的 plan（来自 normalizeFormatSpec）：
            // fontName, fontSize, headerBold, headerColor, alignment, addBorders,
            // autofit, freezeHeader, wrapText, numberFormat, dateFormat
            if (plan.fontName || plan.fontSize) {
                steps.push({ title: '设置字体', detail: (plan.fontName || '默认字体') + ' ' + (plan.fontSize || 11) + ' 号', status: 'done' });
            }
            if (plan.headerBold || plan.headerColor) {
                const parts = [];
                if (plan.headerBold) parts.push('加粗');
                if (plan.headerColor) parts.push('主题色填充');
                steps.push({ title: '设置表头', detail: parts.join(' + ') || '表头样式', status: 'done' });
            }
            if (plan.alignment) {
                const alignLabel = plan.alignment === 'center' ? '居中对齐' : plan.alignment === 'left' ? '左对齐' : '右对齐';
                steps.push({ title: '对齐方式', detail: alignLabel, status: 'done' });
            }
            if (plan.addBorders) steps.push({ title: '添加表格边框', detail: '全边框', status: 'done' });
            if (plan.freezeHeader) steps.push({ title: '冻结表头', detail: '锁定首行', status: 'done' });
            if (plan.autofit) steps.push({ title: '自动调整列宽', detail: '按内容自适应', status: 'done' });
            if (plan.wrapText) steps.push({ title: '自动换行', detail: '开启', status: 'done' });
            if (plan.numberFormat) steps.push({ title: '数字格式', detail: plan.numberFormat, status: 'done' });
            if (plan.dateFormat) steps.push({ title: '日期格式', detail: plan.dateFormat, status: 'done' });
            // 智能列类型识别（Formtter 独有）：如果 numberFormat/dateFormat 是 null 但执行时有自动识别，
            // 也加一个条目；没有字段时至少给个"应用格式"兜底步
            if (steps.length === 0) {
                steps.push({ title: '应用格式', detail: '智能列类型识别 + 格式化', status: 'done' });
            }
        } else if (kind === 'dashboard') {
            // 看板骨架：委托 dashboard-blueprint-engine 的 buildSkeletonSteps（与真实施工步骤共用
            // 同一构造器，从构造上保证"规划步数 == 实际步数"，杜绝骨架与执行错位）。
            const engine = App?.ChatV2?.dashboardModules?.blueprintEngine || null;
            const prepared = extra && extra.prepared ? extra.prepared : null;
            if (engine && typeof engine.buildSkeletonSteps === 'function') {
                for (const s of engine.buildSkeletonSteps(plan, prepared)) {
                    steps.push({ title: s.title, detail: s.detail, status: 'done' });
                }
            } else {
                steps.push({ title: '生成看板', detail: plan.title || '数据看板', status: 'done' });
            }
        }
        // ✨ 统一补字段：description（供 ProgressTracker.buildPlanHtml 渲染步骤名）、
        // strategy（供渲染层显示徽章"工具/推理/管道"）。
        // buildSpecSteps 里各处 push 只写了 title/detail/status，ProgressTracker 只读 description/name。
        return steps.map(s => ({
            description: s.title,   // 双保险：渲染层优先读 description
            strategy: 'direct',     // 默认都是工具调用
            ...s
        }));
    }
    /**
     * 美化工具路径的表头内容识别：读取待美化范围真实数据，依内容判出表头行数（含合并标题行+列头行）。
     * 与 VBA 渲染路径的 IsHeaderRow 语义一致——让纯工具路径也能"自动知道表头在哪"，
     * 不再只看 LLM 给的 headerRows（实测列标题行被当数据）。
     * @returns {Promise<{headerRows:number}|null>} 命中的表头覆盖；无法读取/无范围返回 null
     */
    async function detectBeautifyHeader(ctx, turn, toolRegistry, sessionStore, signal) {
        if (!ctx || !toolRegistry || typeof toolRegistry.invokeTracked !== 'function') return null;
        const sel = (ctx && ctx.selection) || {};
        const addr = String(sel.address || sel.range || '').replace(/\$/g, '').trim();
        if (!addr) return null;
        const sheet = sel.sheetName || ctx.activeSheet || '';
        // 读"原始全量行序"(hasHeaders:false)：必须拿到 标题行+列头行+数据行 的完整顺序。
    // 实测 hasHeaders:true 会把列标签行(7个)单拎成 headers、把上方的大标题行丢掉，
    // 使 detectHeaderRows 只见 1 行表头、只能数到列头行，表头总数永远取不到"标题+列头"=2。
    const readArgs = { returnFormat: 'array', hasHeaders: false, range: addr };
    if (sheet) readArgs.sheet = sheet;
    let rr;
    try {
      rr = await toolRegistry.invokeTracked(turn, 'excel_read_range', readArgs, {
        displayName: '识别表头行数',
        displayArgs: { '表头识别范围': addr, '工作表': sheet || '活动表' },
        ctx: { turn, sessionStore, signal }
      });
    } catch (e) {
      ChatLog.warn?.('PIPELINE', '[spec-build] 表头识别读取异常:', e && e.message ? e.message : e);
      return null;
    }
    if (signal?.aborted) throw new Error('user_aborted');
    if (!rr || rr.ok === false || rr.error) return null;
    // 收敛到二维行矩阵（兼容字符串 / 嵌套对象 .data/.rows/.result）。
    const toRows = (raw) => {
      let d = raw;
      if (typeof d === 'string') { try { d = JSON.parse(d); } catch { return []; } }
      let depth = 0;
      while (d && typeof d === 'object' && !Array.isArray(d) && depth < 5) {
        depth++;
        if (Array.isArray(d.data)) { d = d.data; break; }
        if (Array.isArray(d.rows)) { d = d.rows; break; }
        if (Array.isArray(d.result)) { d = d.result; break; }
        if (d.result !== undefined) { let inner = d.result; if (typeof inner === 'string') { try { inner = JSON.parse(inner); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-spec-build] 解包 result 非 JSON，保留原串:', e); } } d = inner; continue; }
        break;
      }
      return Array.isArray(d) ? d : [];
    };
    const allRows = toRows(rr.result);
    if (allRows.length === 0) return null;
        const hr = SP.util.detectHeaderRows(allRows);
        if (!hr || hr < 1 || hr > 5) { ChatLog.warn?.('PIPELINE', `[spec-build] 表头识别 hr 无效=${hr} allRows=${allRows.length}`); return null; }
        ChatLog.info?.('PIPELINE', `[spec-build] detectBeautifyHeader hr=${hr}（共读取 ${allRows.length} 行）range=${addr}`);
        return { headerRows: hr };
    }
    MODS.specBuild = {
        buildBeautifyOps,
        parseExplicitRange,
        detectDataSelection,
        detectBeautifyHeader,
        buildSpecSteps
    };
})();