/**
 * ============================================================
 * skill-clean-crud.js — Skill 管道清洗/CRUD 执行器（从 skill-pipeline.js 拆出）
 * ------------------------------------------------------------
 * 职责：数据清洗（去重/空值/空格/格式标准化/空行）与 CRUD 读写的 MCP 工具执行路径。
 * 挂载：App.ChatV2.skillPipelineModules.cleanCrud
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
    // 智能数据清洗的 MCP 工具执行路径。
    // 去重 → excel_remove_duplicates（原生引擎，原地删除）
    // 其余操作 → 优先 excel_clean_data（原生引擎，一次性完成），失败回退 JS 变换路径
    // 任一环节失败返回 {ok:false, error}，由调用方直接报错（纯 MCP，不回退 VBA）。
    // ---- 读取校验采样尺寸：各步骤执行前只读首尾几行做数据存在性确认，不拉全量 ----
    const CLEAN_HEAD_ROWS = 5;
    const CLEAN_TAIL_ROWS = 5;

    /** 解析选区 "A1:C100"（兼容单格），返回行列边界；无法解析返回 null */
    function splitRangeAddr(addr) {
        // Excel rng.Address 返回绝对地址（如 $A$1:$G$27），剔除 $ 再解析，
        // 否则正则会因首字符是 $ 匹配失败，退回整块全量读取。
        const m = String(addr || '').replace(/\$/g, '').split(':');
        const A = (m[0] || '').match(/^([A-Za-z]{1,3})(\d+)$/);
        const B = (m[1] || '').match(/^([A-Za-z]{1,3})(\d+)$/);
        if (!A) return null;
        const colL = A[1].toUpperCase();
        const rowT = parseInt(A[2], 10);
        let colR = colL, rowB = rowT;
        if (B) { colR = B[1].toUpperCase(); rowB = parseInt(B[2], 10); }
        return { colL, rowT, colR, rowB };
    }

    /** 由选区生成首块+尾块采样范围，避免首尾重叠；返回 {ranges, covered} */
    function buildHeadTailRanges(addr, headCnt, tailCnt) {
        const s = splitRangeAddr(addr);
        if (!s) return { ranges: [addr], covered: true };
        const rowT = Math.min(s.rowT, s.rowB);
        const rowB = Math.max(s.rowT, s.rowB);
        const ranges = [];
        const headEnd = Math.min(rowT + headCnt - 1, rowB);
        const tailStart = Math.max(rowB - tailCnt + 1, rowT);
        if (headEnd >= rowT) ranges.push(`${s.colL}${rowT}:${s.colR}${headEnd}`);
        const tailBegin = Math.max(tailStart, headEnd + 1);
        if (tailBegin <= rowB) ranges.push(`${s.colL}${tailBegin}:${s.colR}${rowB}`);
        const coveredHead = (headEnd - rowT + 1) >= headCnt || headEnd >= rowB;
        const coveredTail = tailStart <= rowT || (rowB - tailStart + 1) >= tailCnt;
        return { ranges, covered: coveredHead && coveredTail };
    }

    /** 从 excel_read_range 的多层包装中抽取二维行数组 */
    function parseReadRows(raw) {
        let data = raw;
        if (typeof data === 'string') { try { data = JSON.parse(data); } catch { return []; } }
        let depth = 0;
        while (data && typeof data === 'object' && !Array.isArray(data) && depth < 5) {
            depth++;
            if (Array.isArray(data.data)) { data = data.data; break; }
            if (Array.isArray(data.rows)) { data = data.rows; break; }
            if (Array.isArray(data.result)) { data = data.result; break; }
            if (data.result !== undefined) {
                let inner = data.result;
                if (typeof inner === 'string') { try { inner = JSON.parse(inner); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] result 内层非 JSON，保留原串:', e); } }
                data = inner;
                continue;
            }
            break;
        }
        return Array.isArray(data) ? data : [];
    }
    /**
     * 单次清洗操作 helper（per-op）。
     * execute() 里循环调它，每次 yield progress 让面板逐步填色。
     * 返回: { ok, meta, summary, steps, toolCallId }
     */
    async function executeCleanSingleOp(turn, opRaw, plan, ctx, toolRegistry, sessionStore, signal) {
        // normalizeOp: LLM 意图分析 prompt 教用驼峰，但有时候还是会输出下划线/小写变体
        // → 统一 normalize 成内部 STRUCTURE_OPS 认识的驼峰名，前置分流才能命中
        const _opMap = {
            'merge_sheets': 'mergeSheets', 'merge_workbooks': 'mergeWorkbooks',
            'split_by_column': 'splitByColumn', 'split_workbook': 'splitWorkbook', 'split_by_rows': 'splitByRows',
            'mergesheets': 'mergeSheets', 'mergeworkbooks': 'mergeWorkbooks',
            'splitbycolumn': 'splitByColumn', 'splitworkbook': 'splitWorkbook', 'splitbyrows': 'splitByRows'
        };
        const op = _opMap[String(opRaw || '').toLowerCase()] || opRaw;
        // ===== 前置分流：表结构变换 op（split/merge 系列）→ 直接调原子工具 =====
        // 这 5 个 op 不走 excel_clean_data 的"预读→计算→写回"路径，
        // 因为它们各自有独立原子工具（excel_merge_sheets / excel_split_by_column 等），
        // 而且参数形态也完全不同（mergeSheets 要 sourceSheets/sheetList，不需要选区 range）。
        const STRUCTURE_OPS = new Set(['splitByColumn', 'splitWorkbook', 'splitByRows', 'mergeSheets', 'mergeWorkbooks']);
        if (STRUCTURE_OPS.has(op)) {
            return executeStructureTransform(turn, op, plan, ctx, toolRegistry, sessionStore, signal);
        }

        // 工具可用性
        if (typeof toolRegistry?.get !== 'function' || !toolRegistry.get('excel_clean_data')) {
            return { ok: false, error: 'MCP 工具 excel_clean_data 未注册' };
        }
        const selAddr = (ctx.selection?.address || '').replace(/\$/g, '');
        const selSheet = ctx.selection?.sheetName || '';
        if (!selAddr) {
            ChatLog.error?.('PIPELINE', '[clean] executeCleanSingleOp 无选中区域: op=' + String(op) + ' ctx.selection=' + (ctx.selection ? JSON.stringify({ address: ctx.selection.address, rowCount: ctx.selection.rowCount, colCount: ctx.selection.colCount, explicit: !!ctx.selection.explicit, autoDetected: !!ctx.selection.autoDetected }) : 'null'));
            return { ok: false, error: '未检测到选区，请先选中数据区域' };
        }

        // 清洗操作中文标签收敛在 constants.CLEAN_OPS 注册表（cleanOpLabel）
        const opLabel = C.cleanOpLabel(op) || op;

        try {
            // ---- 读取校验：每步清洗前先 excel_read_range 首尾采样，确认选区确有数据。
            // ---- 只读首尾几行（不拉全量），由 invokeTracked 自动生成可见的"读取数据"步骤卡。
            // 【降级】excel_read_range 未注册（环境性缺失，如后端工具清单未加载）时跳过预读校验：
            // 预读仅作参考，清洗真正执行走后端原生工具，由其给出真实错误，不因校验工具缺失卡死流程。
            if (typeof toolRegistry.get === 'function' && !toolRegistry.get('excel_read_range')) {
                ChatLog.warn?.('SKILL', '[clean] excel_read_range 未注册，跳过预读校验，直接执行后端清洗工具');
            } else {
            {
                const { ranges, covered } = buildHeadTailRanges(selAddr, CLEAN_HEAD_ROWS, CLEAN_TAIL_ROWS);
                let sampled = 0;
                let readErr = null;
                let backendInfo = null;
                let lastRawResult = null;
                // sheet 兜底：选区表名为空时回退活动工作表名，避免"有 range 无 sheet"
                // 导致后端读到错误工作表而误报"无可读取数据"。
                const realSheet = selSheet
                    || turn.context?.activeSheet?.sheet
                    || turn.context?.activeSheet?.Name
                    || '';
                for (const r of ranges) {
                    const ra = { returnFormat: 'array', hasHeaders: false };
                    if (realSheet) ra.sheet = realSheet;
                    if (r) ra.range = r;
                    let rr = await toolRegistry.invokeTracked(turn, 'excel_read_range', ra, {
                        displayName: `读取数据校验·${opLabel}`,
                        displayArgs: { '范围': r },
                        ctx: { turn, sessionStore, signal }
                    });
                    if (signal?.aborted) throw new Error('user_aborted');
                    if (rr?.ok === false || rr?.error) { readErr = rr?.error || '读取数据校验失败'; break; }
                    let rowsArr = parseReadRows(rr?.result);
                    // 显式 range 读到空（后端返回 "[]"）时，多半是地址含 $ 前缀、或 sheet/地址与真实选区错位，
                    // 引发后端 rng 为 Nothing。此时回退：仅传 sheet（不带 range），
                    // 让后端走"当前 Selection / UsedRange"读取用户真实选中的区域。
                    if (rowsArr.length === 0 && ra.range) {
                        const fbArgs = { returnFormat: 'array', hasHeaders: false };
                        if (realSheet) fbArgs.sheet = realSheet;
                        const fb = await toolRegistry.invokeTracked(turn, 'excel_read_range', fbArgs, {
                            displayName: `读取数据校验·${opLabel}(回退选区)`,
                            displayArgs: {},
                            ctx: { turn, sessionStore, signal }
                        });
                        if (signal?.aborted) throw new Error('user_aborted');
                        if (!(fb?.ok === false || fb?.error)) {
                            const fbRows = parseReadRows(fb?.result);
                            if (fbRows.length > 0) { rr = fb; rowsArr = fbRows; }
                        }
                    }
                    sampled += rowsArr.length;
                    // rr.result 可能是 JSON 字符串或对象，统一解析后再取后端字段，
                    // 避免上一版"rr.result.rows 直接访问字符串属性 → n/a"。
                    let __meta = null;
                    if (typeof rr?.result === 'string') { try { __meta = JSON.parse(rr.result); } catch { __meta = null; } }
                    if (__meta && typeof __meta === 'object' && !Array.isArray(__meta)) {
                        backendInfo = {
                            range: __meta.range || r,
                            sheet: __meta.worksheet || realSheet,
                            backendRows: __meta.rows,
                            backendResultCount: Array.isArray(__meta.result) ? __meta.result.length : (__meta.result != null ? 1 : 0),
                            parsedRows: rowsArr.length
                        };
                    } else {
                        backendInfo = { range: r, sheet: realSheet, backendRows: undefined, backendResultCount: undefined, parsedRows: rowsArr.length };
                    }
                    lastRawResult = rr?.result;
                }
                if (readErr) return { ok: false, error: readErr };
                if (sampled === 0) {
                    // 预读采样为空时【不再硬中断整个清洗】：
                    // 清洗真正执行走后端原生 excel_clean_data / excel_remove_duplicates 原位写回，
                    // 前端预读仅作参考；若选区确无数据，由后端清洗工具给出真实、具体的错误，
                    // 避免因预读/解析返回异常导致前端误报"无可读取数据"而卡死流程（美化同样不预读却正常）。
                    let rawPreview = '';
                    try {
                        const s = typeof lastRawResult === 'string' ? lastRawResult : JSON.stringify(lastRawResult);
                        if (s) rawPreview = s.length > 300 ? s.slice(0, 300) + '…' : s;
                    } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] 预读采样序列化失败:', e); }
                    ChatLog.warn?.('PIPELINE', `[skill-cleaner] 预读采样为空(sampled=0)，降级继续执行（交后端判定选区）: 表=${backendInfo?.sheet || realSheet || '活动表'} 原始返回=${rawPreview || '(空)'}`);
                }
                try { console.log(`[CLEAN-READ#${op}] sample=${sampled} rows covered=${covered} ranges=${JSON.stringify(ranges)} backend=${JSON.stringify(backendInfo)}`); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] CLEAN-READ 调试日志输出失败:', e); }
            }
            } // end else（excel_read_range 已注册时才预读校验）

            // 去重走原生 excel_remove_duplicates（只算不写），其余走 excel_clean_data。
            // 两个工具都只返回结果不落盘，统一由下方 writeCleanBack 原位写回 —— 修复"光完成不动 Excel"。
            const isDedup = (op === 'removeDuplicates');
            const markOnly = isDedup && plan?.markDuplicate === true;   // "仅高亮" → 只标黄重复行，不删除
            const cleanTool = isDedup ? 'excel_remove_duplicates' : 'excel_clean_data';
            if (!toolRegistry.get(cleanTool)) {
                return { ok: false, error: `MCP 工具 ${cleanTool} 未注册` };
            }
            const cleanArgs = isDedup ? { hasHeaderRow: true } : { operations: [op], hasHeaderRow: true };
            // dedupeColumns 只对去重有意义
            if (isDedup) {
                ChatLog.info?.('PIPELINE', `[skill-cleaner] 单步去重: plan.markDuplicate=${plan.markDuplicate} (type=${typeof plan.markDuplicate})`);
                if (Array.isArray(plan.dedupeColumns) && plan.dedupeColumns.length > 0) {
                    cleanArgs.columns = plan.dedupeColumns;
                }
            } else if (op === 'fillEmpty' && plan.fillMethod) {
                cleanArgs.fillMethod = plan.fillMethod;
            }
            if (selSheet) cleanArgs.sheet = selSheet;
            if (selAddr) cleanArgs.range = selAddr;

            const displayArgs = {
                '操作': opLabel,
                '数据范围': selAddr,
                ...(isDedup ? { '去重模式': markOnly ? '仅高亮标记' : '删除重复行' } : {})
            };
            const result = await toolRegistry.invokeTracked(turn, cleanTool, cleanArgs, {
                displayName: `清洗·${opLabel}`,
                displayArgs,
                ctx: { turn, sessionStore, signal }
            });
            if (signal?.aborted) throw new Error('user_aborted');
            if (result?.ok === false || result?.error) {
                // cancelled 透传：守卫拦截/用户取消时上层据此优雅收尾，不降级 agent-loop
                return { ok: false, cancelled: !!result?.cancelled, error: result?.error || `${opLabel}执行失败` };
            }

            // ✨ 调试日志
            try {
                console.log(`[CLEAN-RAW#${op}] result=`, result);
                console.log(`[CLEAN-RAW#${op}] result.result typeof=`, typeof result?.result, `value=`, typeof result?.result === 'string' ? result?.result?.substring?.(0, 300) : result?.result);
            } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] CLEAN-RAW 调试日志输出失败:', e); }

            // 解析返回：invokeTracked → executeLocalExcelTool 有多层 { result: ... } / content  / JSON 字符串 包装，
            // 用健壮的多层解包下钻，取到含工具业务字段（data/removedRows/removedCount 等）的最深层对象。
            // 修复：此前只下钻一层，导致 meta.data 永远是 undefined，写回步骤拿不到数据 → 全部判"失败"。
            let meta = {};
            {
                const isBizNode = (o) => o && typeof o === 'object' && !Array.isArray(o) &&
                    (Array.isArray(o.data) || o.removedRows !== undefined || o.removedCount !== undefined ||
                        o.cleanedRows !== undefined || o.uniqueCount !== undefined || o.rawCount !== undefined ||
                        o.success !== undefined || Array.isArray(o.headers));
                const parseIfString = (v) => {
                    if (typeof v !== 'string') return v;
                    const s = v.trim();
                    if (s.startsWith('{') || s.startsWith('[')) { try { return JSON.parse(s); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] 解包 parseIfString 非 JSON:', e); } }
                    return v;
                };
                const chain = [];
                const seenNodes = new Set();
                // result.result 是 executeLocalExcelTool 序列化的 JSON 字符串
                //（{"result":<toolObj>,"taskCompleted":true}）→ 必须先 parse，支持对象也兼容
                let curNode = parseIfString(result?.result);
                while (curNode && typeof curNode === 'object' && !Array.isArray(curNode) && !seenNodes.has(curNode)) {
                    seenNodes.add(curNode);
                    chain.push(curNode);
                    let next = null;
                    // { result: <obj> } 包装
                    if (curNode.result !== undefined) {
                        const rv = parseIfString(curNode.result);
                        if (rv && typeof rv === 'object' && !Array.isArray(rv)) next = rv;
                    } else if (Array.isArray(curNode.content) && curNode.content.length > 0) {
                        const txt = curNode.content.find(c => c && c.type === 'text' && typeof c.text === 'string' && c.text.trim());
                        if (txt) next = parseIfString(txt.text);
                    }
                    if (next === null || next === undefined) break;
                    curNode = next;
                }
                for (let i = chain.length - 1; i >= 0; i--) {
                    if (isBizNode(chain[i])) { meta = chain[i]; break; }
                }
                if (!meta || typeof meta !== 'object' || Array.isArray(meta)) meta = {};
            };
            const steps = SP.util.findStepsDeep(result);

            ChatLog.info?.('PIPELINE', `[skill-cleaner] per-op "${op}" meta=${JSON.stringify(meta)} keys=${Object.keys(meta).join(',')} rawType=${typeof result?.result}`);

            // ---- 写回 / 高亮 ----
            // "仅高亮"（markDuplicate=true）：不删除行，只对重复行标黄（用后端返回的 duplicateRows 绝对行号）。
            // 其余：把清洗/去重结果原位写回当前选区（此前只算不写，Excel 从未被改动）。
            let writeOk = true;
            let writeError = null;
            if (isDedup && markOnly) {
                const hl = await highlightDuplicateRows(turn, meta, selAddr, selSheet, opLabel, toolRegistry, sessionStore, signal);
                if (!hl.ok) { writeOk = false; writeError = hl.error || '高亮重复行失败'; }
            } else {
                const writeBack = await writeCleanBack(turn, meta, selAddr, selSheet, opLabel, toolRegistry, sessionStore, signal);
                writeOk = writeBack.ok;
                writeError = writeBack.error;
            }

            // [诊断] 单步清洗失败时输出决定性线索：后端原样返回 + 解析出的 meta + 写回失败原因。
            try {
                const metaKeys = Object.keys(meta || {});
                console.log(`[CLEAN-DIAG#${op}] writeOk=${writeOk} writeError=${writeError || '(none)'} metaKeys=[${metaKeys.join(',')}] hasData=Array.isArray(${JSON.stringify(Array.isArray(meta?.data))}) metaDataLen=${Array.isArray(meta?.data) ? meta.data.length : 'n/a'} rawType=${typeof result?.result}`);
            } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] CLEAN-DIAG 调试日志输出失败:', e); }

            return { ok: true, meta, steps, toolCallId: result?.toolCallId ?? null, writeOk, writeError };
        } catch (e) {
            if (e?.message === 'user_aborted') throw e;
            return { ok: false, error: e?.message || String(e) };
        }
    }

    /**
     * 表结构变换独立执行器（splitByColumn / mergeSheets 等 5 个 op）。
     * 不走 excel_clean_data 的"预读→计算→写回"路径，因为每个 op 对应独立原子 MCP 工具，
     * 且参数形态完全不同（mergeSheets 要 sourceSheets/sheetList，不需要选区 range）。
     * 返回值与 executeCleanSingleOp 一致：{ ok, meta, summary?, steps?, toolCallId? }
     */
    async function executeStructureTransform(turn, op, plan, ctx, toolRegistry, sessionStore, signal) {
        // op → { tool名, 工具注册检查 }
        const OP_REGISTRY = {
            splitByColumn:   { tool: 'excel_split_by_column',   need: ['excel_split_by_column'] },
            splitWorkbook:   { tool: 'excel_split_workbook',    need: ['excel_split_workbook'] },
            splitByRows:     { tool: 'excel_split_by_rows',      need: ['excel_split_by_rows'] },
            mergeSheets:     { tool: 'excel_merge_sheets',      need: ['excel_merge_sheets'] },
            mergeWorkbooks:  { tool: 'excel_merge_workbooks',   need: ['excel_merge_workbooks'] }
        };
        const entry = OP_REGISTRY[op];
        if (!entry) return { ok: false, error: `未知的表结构变换 op: ${op}` };

        // 工具可用性检查
        for (const t of entry.need) {
            if (typeof toolRegistry?.get !== 'function' || !toolRegistry.get(t)) {
                return { ok: false, error: `MCP 工具 ${t} 未注册` };
            }
        }

        // ==== 构造工具参数 ====
        const args = {};
        const realSheet = ctx.selection?.sheetName
            || turn.context?.activeSheet?.sheet
            || '';

        try {
            switch (op) {
                case 'splitByColumn': {
                    const dimCol = plan.dimColumn || plan.column || '';
                    if (!dimCol) return { ok: false, error: '请指定按哪一列拆分（dimColumn）' };
                    args.column = dimCol;
                    args.outputMode = plan.outputMode || 'sheets';
                    // 默认后端 onConflict=skip（同名表不叠 _2）；用户说重新/覆盖时覆写
                    const umsg = String(turn?.message || turn?.context?.userMessage || '');
                    if (plan.onConflict) args.onConflict = plan.onConflict;
                    else if (/(重新|再次|重做|覆盖|再拆|重新拆)/.test(umsg)) args.onConflict = 'replace';
                    const selSheets = turn.context?.selectedSheets;
                    if (Array.isArray(selSheets) && selSheets.length > 0) {
                        args.sourceSheets = JSON.stringify(selSheets);
                    }
                    if (plan.targetSheet) args.targetSheet = plan.targetSheet;
                    if (plan.mergeAll) args.mergeAll = true;
                    break;
                }
                case 'splitByRows': {
                    args.rowsPerSheet = plan.splitRows || plan.rowsPerSheet || 100;
                    if (realSheet) args.sheet = realSheet;
                    break;
                }
                case 'splitWorkbook': {
                    if (plan.outputDir) args.outputDir = plan.outputDir;
                    break;
                }
                case 'mergeSheets': {
                    const selSheets = turn.context?.selectedSheets;
                    if (Array.isArray(selSheets) && selSheets.length > 0) {
                        args.sourceSheets = JSON.stringify(selSheets);
                    } else if (plan.sourceSheets) {
                        const s = typeof plan.sourceSheets === 'string' ? plan.sourceSheets : JSON.stringify(plan.sourceSheets);
                        args.sourceSheets = s;
                    }
                    if (plan.targetSheet) args.targetSheet = plan.targetSheet;
                    else if (plan.mergeAll) args.mergeAll = true;
                    break;
                }
                case 'mergeWorkbooks': {
                    if (Array.isArray(plan.filePaths) && plan.filePaths.length > 0) {
                        args.filePaths = plan.filePaths;
                    } else if (typeof plan.filePaths === 'string' && plan.filePaths.length > 0) {
                        try { args.filePaths = JSON.parse(plan.filePaths); } catch { /* 忽略 */ }
                    }
                    if (!args.filePaths || args.filePaths.length === 0) {
                        return { ok: false, error: '请指定要合并的 xlsx 文件路径列表（filePaths）' };
                    }
                    if (plan.targetPath) args.targetPath = plan.targetPath;
                    break;
                }
            }

            const displayArgs = {};
            if (args.column) displayArgs['拆分列'] = args.column;
            if (args.rowsPerSheet) displayArgs['每表行数'] = args.rowsPerSheet;
            if (args.sourceSheets) displayArgs['源工作表'] = '已指定多个';
            if (args.filePaths) displayArgs['文件数'] = args.filePaths.length;
            if (args.targetSheet) displayArgs['目标表'] = args.targetSheet;

            ChatLog.info?.('PIPELINE', `[skill-cleaner] 表结构变换 "${op}" 调用 ${entry.tool}: args=${JSON.stringify(args)}`);

            const result = await toolRegistry.invokeTracked(turn, entry.tool, args, {
                displayName: `表结构·${op}`,
                displayArgs,
                ctx: { turn, sessionStore, signal }
            });
            if (signal?.aborted) throw new Error('user_aborted');
            if (result?.ok === false || result?.error) {
                return { ok: false, error: result?.error || `${op}执行失败` };
            }

            let meta = {};
            try {
                let curNode = result?.result;
                if (typeof curNode === 'string') { try { curNode = JSON.parse(curNode); } catch { /* keep */ } }
                while (curNode && typeof curNode === 'object' && !Array.isArray(curNode)) {
                    if (curNode.success !== undefined || curNode.mergedRows !== undefined || curNode.writtenTo !== undefined
                        || curNode.sourceSheets !== undefined || curNode.splitBy !== undefined
                        || curNode.newSheets !== undefined || curNode.outputFiles !== undefined) { meta = curNode; break; }
                    if (curNode.result !== undefined) { curNode = curNode.result; continue; }
                    if (Array.isArray(curNode.content) && curNode.content[0]?.type === 'text') {
                        try { curNode = JSON.parse(curNode.content[0].text); } catch { break; }
                        continue;
                    }
                    break;
                }
                if (!meta || typeof meta !== 'object') meta = {};
            } catch { meta = {}; }

            const steps = SP.util.findStepsDeep(result);
            ChatLog.info?.('PIPELINE', `[skill-cleaner] 表结构变换 "${op}" 完成: meta=${JSON.stringify(meta)}`);

            return { ok: true, meta, steps, toolCallId: result?.toolCallId ?? null, writeOk: true };
        } catch (e) {
            if (e?.message === 'user_aborted') throw e;
            ChatLog.error?.('PIPELINE', `[skill-cleaner] 表结构变换 "${op}" 异常:`, e);
            return { ok: false, error: e?.message || String(e) };
        }
    }

    /**
     * "仅高亮"模式：不删除行，把后端上报的重复行（绝对行号）用 excel_format_range 标黄。
     * @returns {{ok:boolean,error?:string,highlighted?:number,total?:number}}
     */
    async function highlightDuplicateRows(turn, meta, selAddr, selSheet, opLabel, toolRegistry, sessionStore, signal) {
        const rows = Array.isArray(meta.duplicateRows)
            ? meta.duplicateRows.filter(r => Number.isFinite(Number(r)) && Number(r) >= 1).map(Number)
            : [];
        if (rows.length === 0) {
            return { ok: true, okEmpty: true, highlighted: 0, total: 0 };
        }
        if (!toolRegistry.get('excel_format_range')) {
            return { ok: false, error: 'excel_format_range 未注册，无法高亮重复行' };
        }
        const startCell = getTopLeftCell(selAddr);
        const startColL = getColLetter(startCell);
        const colCount = (Array.isArray(meta.headers) && meta.headers.length > 0)
            ? meta.headers.length
            : (Array.isArray(meta.data?.[0]) ? meta.data[0].length : 1);
        const endColL = colCount >= 1 ? getEndColLetter(startColL, colCount) : startColL;
        let highlighted = 0;
        for (const r of rows) {
            const fmtArgs = { range: `${startColL}${r}:${endColL}${r}`, format: JSON.stringify({ bgColor: '#FFFF00', bgScope: 'all' }) };
            if (selSheet) fmtArgs.sheet = selSheet;
            try {
                await toolRegistry.invokeTracked(turn, 'excel_format_range', fmtArgs, {
                    displayName: `高亮重复·${opLabel}`,
                    displayArgs: { '行': r, '颜色': '黄色' },
                    ctx: { turn, sessionStore, signal }
                });
                highlighted++;
            } catch (e) {
                ChatLog.warn?.('PIPELINE', `[skill-cleaner] 高亮重复行 ${r} 失败: ${e?.message || String(e)}`);
            }
        }
        return { ok: true, highlighted, total: rows.length };
    }

    /**
     * 把单个清洗步骤的结果 meta.data（表头+数据，二维数组）原位写回当前选区：
     *   excel_write_range（clearFirst）写入 → 若行数减少，清掉原选区下方残留行。
     * 仅当写回成功才返回 ok=true，供 execute 判定步骤"完成"。
     */
    async function writeCleanBack(turn, meta, selAddr, selSheet, opLabel, toolRegistry, sessionStore, signal) {
        const dataArr = (meta && Array.isArray(meta.data)) ? meta.data : null;
        if (!dataArr || dataArr.length === 0) {
            return { ok: false, error: '清洗结果为空，未执行写回' };
        }
        const startCell = getTopLeftCell(selAddr);
        const originalHeight = getRangeHeight(meta.range || selAddr);
        const dataHeight = dataArr.length;
        const firstRow = dataArr[0];
        const colCount = Array.isArray(firstRow) ? firstRow.length : 0;
        try {
            if (!toolRegistry.get('excel_write_range')) return { ok: false, error: 'excel_write_range 未注册' };
            const writeArgs = { range: startCell, values: JSON.stringify(dataArr), clearFirst: true, writeMode: 'overwrite' };
            if (selSheet) writeArgs.sheet = selSheet;
            const wres = await toolRegistry.invokeTracked(turn, 'excel_write_range', writeArgs, {
                displayName: `写回清洗结果·${opLabel}`,
                displayArgs: { '目标': startCell, '数据行数': Math.max(0, dataHeight - (meta.hasHeader ? 1 : 0)), '列数': colCount },
                ctx: { turn, sessionStore, signal }
            });
            if (signal?.aborted) throw new Error('user_aborted');
            if (wres?.ok === false || wres?.error) {
                return { ok: false, error: wres?.error || '写回失败' };
            }

            // 行数减少时，物理删除原选区下方残留行（去重删行 / 删除空行后）。
            // 用 excel_delete_row 删除整行而非 excel_clear_range 只清内容，否则残留行仍在（只是变空白），
            // 用户会看到"空行还在 / 只删了一行"。删除是从 fromRow 起连删 (originalHeight - dataHeight) 行。
            if (dataHeight < originalHeight && colCount > 0) {
                const startRow = getRowNum(startCell);
                const fromRow = startRow + dataHeight;
                const toRow = startRow + originalHeight - 1;
                if (fromRow <= toRow) {
                    const deleteCount = toRow - fromRow + 1;
                    const deleteArgs = { row: fromRow, count: deleteCount };
                    if (selSheet) deleteArgs.sheet = selSheet;
                    try {
                        if (toolRegistry.get('excel_delete_row')) {
                            await toolRegistry.invokeTracked(turn, 'excel_delete_row', deleteArgs, {
                                displayName: `删除残留行·${opLabel}`,
                                displayArgs: { '起始行': fromRow, '行数': deleteCount },
                                ctx: { turn, sessionStore, signal }
                            });
                        }
                    } catch (ce) {
                        ChatLog.warn?.('PIPELINE', `[skill-cleaner] 删除残留行失败: ${ce?.message || String(ce)}`);
                    }
                }
            }

            // 写回后为日期列设置日期显示格式：normalizeDates 确认的列（1基相对列号）
            // 即使底层值是 Excel 序列号，显式设 yyyy/m/d 也让日期以可读格式显示，杜绝"仍是 45454"的伪完成。
            if (Array.isArray(meta.dateCols) && meta.dateCols.length > 0 && toolRegistry.get('excel_format_range')) {
                const startColL = getColLetter(startCell);
                const startRowNum = getRowNum(startCell);
                const endRowNum = startRowNum + dataHeight - 1;
                for (const dcRaw of meta.dateCols) {
                    const dc = Number(dcRaw);
                    if (!Number.isInteger(dc) || dc < 1) continue;
                    const colL = getEndColLetter(startColL, dc);
                    const fmtArgs = { range: `${colL}${startRowNum}:${colL}${endRowNum}`, format: JSON.stringify({ numberFormat: 'yyyy/m/d' }) };
                    if (selSheet) fmtArgs.sheet = selSheet;
                    try {
                        await toolRegistry.invokeTracked(turn, 'excel_format_range', fmtArgs, {
                            displayName: `日期格式·${opLabel}`,
                            displayArgs: { '列': colL, '格式': 'yyyy/m/d' },
                            ctx: { turn, sessionStore, signal }
                        });
                    } catch (fe) {
                        ChatLog.warn?.('PIPELINE', `[skill-cleaner] 设置日期列 ${colL} 格式失败: ${fe?.message || String(fe)}`);
                    }
                }
            }
            return { ok: true };
        } catch (e) {
            if (e?.message === 'user_aborted') throw e;
            return { ok: false, error: e?.message || String(e) };
        }
    }

    /**
     * 从后端真实 meta 生成单步 summary。
     * 只展示有证据（数字>0）的结果，不编造任何"已XX"。
     */
    function buildCleanOpSummary(op, meta, plan, writeOk) {
        if (!meta || typeof meta !== 'object') return '';
        const pickNum = (...keys) => {
            for (const k of keys) {
                const v = meta[k];
                if (typeof v === 'number' && v > 0) return v;
            }
            return 0;
        };
        // 是否存在记录数量的字段（即使 ==0），用于区分"确实无该问题"与"后端没返回字段"
        const hasCount = (...keys) => keys.some(k => typeof meta[k] === 'number');
        // 诚实 fallback：把 meta 里有用的数字字段列出来（过滤掉 success/cleanedRows 这种噪音）
        const NOISE_FIELDS = new Set(['success', 'cleanedRows', 'hasHeaderRow', 'sheet', 'range', 'operations', 'fillMethod', 'markDuplicate', 'message', 'data', 'dataRows']);
        const dumpMeta = () => {
            const parts = [];
            for (const [k, v] of Object.entries(meta)) {
                if (NOISE_FIELDS.has(k)) continue;
                if (typeof v === 'number' || typeof v === 'boolean') {
                    parts.push(`${k}=${v}`);
                }
            }
            if (parts.length === 0) return '';
            return `(${parts.join(', ')})`;
        };
        const notWritten = writeOk === false ? '（未写回）' : '';
        // 调试明细只进控制台，不上屏——避免把 (cleaned=true, hasHeader=true, ...) 这类原始字段漏给用户。
        const traceDump = (rawDump) => {
            try { if (rawDump) console.log('[CLEAN-SUMMARY]', op, rawDump); } catch (e) { ChatLog.debug?.('PIPELINE', '[skill-cleaner] CLEAN-SUMMARY 调试日志输出失败:', e); }
            return `${notWritten}`;
        };

        switch (op) {
            case 'removeDuplicates': {
                const removed = pickNum('removedCount', 'duplicateCount', 'removedRows', 'deletedRows', 'removedRowCount', 'rowsRemoved');
                const markOnly = plan?.markDuplicate === true;
                // 去重步骤只呈现"重复行"结果；空行统一交由 removeEmptyRows 单独呈现，避免空行被计进重复数而概念混同。
                if (removed > 0) return markOnly ? `标记 ${removed} 行重复（仅高亮未删除）${notWritten}` : `删除 ${removed} 行重复数据${notWritten}`;
                if (hasCount('removedCount', 'duplicateCount', 'removedRows', 'deletedRows', 'removedRowCount', 'rowsRemoved')) return `未发现重复行${notWritten}`;
                return `去重完成${traceDump(dumpMeta())}`.trim();
            }
            case 'trimSpaces': {
                const n = pickNum('trimmedCells', 'trimmed', 'cellsTrimmed', 'trimmedCount', 'spacesRemoved', 'trimmedSpaces');
                if (n > 0) return `清理 ${n} 个单元格空格${notWritten}`;
                if (hasCount('trimmedCells', 'trimmed', 'cellsTrimmed', 'trimmedCount', 'spacesRemoved', 'trimmedSpaces')) return `未发现多余空格${notWritten}`;
                return `去空格完成${traceDump(dumpMeta())}`.trim();
            }
            case 'removeEmptyRows': {
                const n = pickNum('removedRows', 'removedEmptyRows', 'emptyRowsRemoved', 'deletedRows', 'rowsRemoved');
                if (n > 0) return `删除 ${n} 行空行${notWritten}`;
                if (hasCount('removedRows', 'removedEmptyRows', 'emptyRowsRemoved', 'deletedRows', 'rowsRemoved')) return `未发现空行${notWritten}`;
                return `删空行完成${traceDump(dumpMeta())}`.trim();
            }
            case 'normalizeDates': {
                const n = pickNum('normalizedDates', 'dateCellsNormalized', 'datesFixed', 'datesConverted', 'formattedDates');
                if (n > 0) return `统一 ${n} 处日期格式${notWritten}`;
                if (hasCount('normalizedDates', 'dateCellsNormalized', 'datesFixed', 'datesConverted', 'formattedDates')) return `日期无需调整${notWritten}`;
                return `日期统一完成${traceDump(dumpMeta())}`.trim();
            }
            case 'normalizeNumbers': {
                const n = pickNum('normalizedNumbers', 'numberCellsNormalized', 'numbersFixed', 'numbersConverted', 'formattedNumbers');
                if (n > 0) return `统一 ${n} 处数字格式${notWritten}`;
                if (hasCount('normalizedNumbers', 'numberCellsNormalized', 'numbersFixed', 'numbersConverted', 'formattedNumbers')) return `数字无需调整${notWritten}`;
                return `数字统一完成${traceDump(dumpMeta())}`.trim();
            }
            case 'fillEmpty': {
                const n = pickNum('filledCells', 'cellsFilled', 'emptyFilled', 'fillCount');
                if (n > 0) return `填充 ${n} 个空值${notWritten}`;
                if (hasCount('filledCells', 'cellsFilled', 'emptyFilled', 'fillCount')) return `未发现空值${notWritten}`;
                return `填充完成${traceDump(dumpMeta())}`.trim();
            }
            case 'standardizeText': {
                const n = pickNum('standardizedCells', 'cellsStandardized', 'textStandardized');
                if (n > 0) return `规范化 ${n} 处文本${notWritten}`;
                if (hasCount('standardizedCells', 'cellsStandardized', 'textStandardized')) return `文本无需调整${notWritten}`;
                return `文本规范化完成${traceDump(dumpMeta())}`.trim();
            }
            // —— 表结构变换：每个 op 对应独立原子工具 ——
            case 'splitByColumn': {
                const sheets = pickNum('newSheets', 'createdSheets', 'sheetCount');
                const col = plan?.dimColumn ? `（按「${plan.dimColumn}」列）` : '';
                if (sheets > 0) return `拆出 ${sheets} 张新 Sheet${col}`;
                if (meta?.writtenTo) return `拆分完成 → ${meta.writtenTo}${col}`;
                return `按列拆分完成${col}${traceDump(dumpMeta())}`.trim();
            }
            case 'splitWorkbook': {
                const files = pickNum('outputFiles', 'createdFiles', 'fileCount');
                if (files > 0) return `拆出 ${files} 个 xlsx 文件`;
                return `工作簿拆分完成${traceDump(dumpMeta())}`.trim();
            }
            case 'splitByRows': {
                const sheets = pickNum('newSheets', 'createdSheets', 'sheetCount');
                const rows = plan?.splitRows ? `（每 ${plan.splitRows} 行）` : '';
                if (sheets > 0) return `拆出 ${sheets} 张 Sheet${rows}`;
                return `按行拆分完成${rows}${traceDump(dumpMeta())}`.trim();
            }
            case 'mergeSheets': {
                const rows = pickNum('mergedRows', 'mergedDataRows', 'totalRows');
                const srcs = meta?.sourceSheets ? meta.sourceSheets.length : 0;
                const to = meta?.writtenTo ? ` → ${meta.writtenTo}` : '';
                if (rows > 0 && srcs > 0) return `合并 ${srcs} 张表（${rows} 行）${to}`;
                if (rows > 0) return `合并 ${rows} 行数据${to}`;
                if (srcs > 0) return `合并 ${srcs} 张表${to}`;
                return `工作表合并完成${traceDump(dumpMeta())}`.trim();
            }
            case 'mergeWorkbooks': {
                const books = meta?.sourceSheets ? meta.sourceSheets.length : pickNum('mergedWorkbooks', 'fileCount');
                if (books > 0) return `合并 ${books} 个工作簿${meta?.targetPath ? ` → ${meta.targetPath}` : ''}`;
                return `工作簿合并完成${traceDump(dumpMeta())}`.trim();
            }
            default:
                traceDump(dumpMeta());
                return `清洗完成${notWritten}`.trim();
        }
    }

    // ===== 增删改查（CRUD）确定性工具链 =====
    // 契约表（operation → 中文标签 / MCP 工具名）收敛在 skill-constants.js（CRUD_OPS / CRUD_OP_TOOL）
    const CRUD_OP_LABEL = C.CRUD_OPS;
    const CRUD_OP_TOOL = C.CRUD_OP_TOOL;

    /** CRUD 操作族中文标签 */
    function opLabel(op) {
        return CRUD_OP_LABEL[op] || String(op || '');
    }

    /** CRUD 目标位置中文描述（步骤详情用） */
    function crudWhere(plan) {
        const parts = [];
        if (plan.sheet) parts.push(plan.sheet);
        const rc = resolveRowCol(plan);
        if (plan.range) parts.push(plan.range);
        else if (rc.column != null) parts.push('第' + rc.column + '列');
        else if (rc.row != null) parts.push('第' + rc.row + '行');
        if (plan.operation === 'write' && plan.value != null) parts.push('值：「' + plan.value + '」');
        if (plan.operation === 'write' && Array.isArray(plan.values) && plan.values.length) parts.push('共 ' + plan.values.length + ' 行数据');
        if ((plan.operation === 'insert' || plan.operation === 'delete') && rc.count > 1) parts.push(rc.count + (rc.column != null ? '列' : '行'));
        if (plan.operation === 'clear') {
            if (plan.clearContent === false) parts.push('仅清格式');
            else if (plan.clearFormat === false) parts.push('仅清内容');
        }
        return parts.filter(Boolean).join('·') || '当前选区';
    }

    /**
     * 从 spec 的 row/column 字段 + range 字符串解析 1 基行列锚点与数量。
     * 支持：显式 row/column；range 为 "5:8"（行）→ row=5,count=4；"C:C"/"C" → column=3；
     *       单格 "A3" → 视为锚点（默认行）。
     */
    function resolveRowCol(plan) {
        let row = (typeof plan.row === 'number') ? plan.row : null;
        let column = (typeof plan.column === 'number') ? plan.column : null;
        let count = (typeof plan.count === 'number' && plan.count >= 1) ? plan.count : 1;
        let range = (typeof plan.range === 'string' && plan.range.trim()) ? plan.range.trim() : '';
        // 行范围 "5:8"
        const rowRange = range.match(/^\s*(\d{1,7})\s*:\s*(\d{1,7})\s*$/);
        if (rowRange && row == null) {
            row = parseInt(rowRange[1], 10);
            count = Math.max(1, parseInt(rowRange[2], 10) - row + 1);
        }
        // 列范围 "C:C" / "D:E"
        const colRange = range.match(/^\s*([A-Za-z]{1,2})\s*:\s*([A-Za-z]{1,2})\s*$/);
        if (colRange && column == null) {
            const c1 = excelColToNum(colRange[1]), c2 = excelColToNum(colRange[2]);
            column = c1 != null ? c1 : null;
            if (c1 != null && c2 != null) count = Math.max(1, c2 - c1 + 1);
        }
        // 列 "C" 或 "C列"
        if (column == null && range && /^[A-Za-z]{1,2}$/.test(range)) column = excelColToNum(range);
        // 单元格锚点 "A3"（写入/插入定位用）
        const cell = range.match(/^([A-Za-z]{1,2})(\d{1,6})$/);
        if (cell && row == null && column == null) {
            column = excelColToNum(cell[1]);
            row = parseInt(cell[2], 10);
        }
        return { row, column, count };
    }

    /** 字母列 → 1 基列号（A=1） */
    function excelColToNum(letters) {
        if (!letters) return null;
        let n = 0;
        const L = String(letters).toUpperCase();
        for (let i = 0; i < L.length; i++) {
            const code = L.charCodeAt(i) - 64;
            if (code < 1 || code > 26) return null;
            n = n * 26 + code;
        }
        return n;
    }

    /** 单格写入目标地址（"A1"）判定；非单格或为空返回 null */
    function singleCellAddress(plan) {
        if (plan.range && /^[A-Za-z]{1,2}\d{1,6}$/.test(plan.range.trim())) return plan.range.trim();
        if (plan.row != null && plan.column != null) {
            // 用列号反推字母
            let n = plan.column; let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
            return s + plan.row;
        }
        return null;
    }

    /**
     * 执行 CRUD（纯 MCP 确定性工具链，不回退 VBA）。
     * @returns {Promise<{ok:boolean, error?:string, result?:Object, toolCallId?}>}
     */
    async function executeCrudWithMcp(turn, plan, ctx, toolRegistry, sessionStore, signal) {
        const op = plan.operation || 'read';
        const displayName = opLabel(op);
        const where = crudWhere(plan);
        // 写入/删除/清空属写操作：标准模式下危险确认会由 tool-registry 兜底，这里不额外弹框
        try {
            if (op === 'read') {
                const args = { returnFormat: 'object', hasHeaders: true };
                if (plan.sheet) args.sheet = plan.sheet;
                if (plan.range) args.range = plan.range;
                const res = await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.read, args, {
                    displayName, displayArgs: { '数据范围': where },
                    ctx: { turn, sessionStore, signal }
                });
                return res;
            }
            if (op === 'write') {
                const hasBulk = Array.isArray(plan.values) && plan.values.length > 0;
                if (hasBulk) {
                    // 批量二维数组 → excel_write_range
                    const args = { values: JSON.stringify(plan.values), clearFirst: false, writeMode: 'overwrite' };
                    if (plan.sheet) args.sheet = plan.sheet;
                    if (plan.range) args.range = plan.range;
                    const res = await toolRegistry.invokeTracked(turn, 'excel_write_range', args, {
                        displayName, displayArgs: { '目标': where, '行数': plan.values.length },
                        ctx: { turn, sessionStore, signal }
                    });
                    return res;
                }
                // 单值 → excel_set_cell_value
                const cell = singleCellAddress(plan) || (plan.range || '');
                if (!cell) {
                    return { ok: false, error: '写入目标未指定：请给出单元格（如 A1）或选择数据区域' };
                }
                const args = { cell, value: String(plan.value != null ? plan.value : '') };
                if (plan.sheet) args.sheet = plan.sheet;
                const res = await toolRegistry.invokeTracked(turn, 'excel_set_cell_value', args, {
                    displayName, displayArgs: { '单元格': cell, '值': args.value },
                    ctx: { turn, sessionStore, signal }
                });
                return res;
            }
            if (op === 'insert') {
                const rc = resolveRowCol(plan);
                if (plan.append || (rc.row == null && rc.column == null)) {
                    const args = { append: true };
                    if (plan.sheet) args.sheet = plan.sheet;
                    return await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.insertAppend, args, {
                        displayName, displayArgs: { '位置': '表尾追加' },
                        ctx: { turn, sessionStore, signal }
                    });
                }
                if (rc.column != null && rc.row == null) {
                    const args = { column: rc.column, count: rc.count };
                    if (plan.sheet) args.sheet = plan.sheet;
                    return await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.insertCol, args, {
                        displayName, displayArgs: { '位置': '第' + rc.column + '列', '数量': rc.count },
                        ctx: { turn, sessionStore, signal }
                    });
                }
                const args = { row: rc.row || 1, count: rc.count };
                if (plan.sheet) args.sheet = plan.sheet;
                return await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.insertRow, args, {
                    displayName, displayArgs: { '位置': '第' + (rc.row || 1) + '行', '数量': rc.count },
                    ctx: { turn, sessionStore, signal }
                });
            }
            if (op === 'delete') {
                const rc = resolveRowCol(plan);
                if (rc.column != null && rc.row == null) {
                    const args = { column: rc.column, count: rc.count };
                    if (plan.sheet) args.sheet = plan.sheet;
                    return await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.deleteCol, args, {
                        displayName, displayArgs: { '位置': '第' + rc.column + '列', '数量': rc.count },
                        ctx: { turn, sessionStore, signal }
                    });
                }
                if (rc.row != null) {
                    const args = { row: rc.row, count: rc.count };
                    if (plan.sheet) args.sheet = plan.sheet;
                    return await toolRegistry.invokeTracked(turn, CRUD_OP_TOOL.deleteRow, args, {
                        displayName, displayArgs: { '位置': '第' + rc.row + '行', '数量': rc.count },
                        ctx: { turn, sessionStore, signal }
                    });
                }
                return { ok: false, error: '删除目标未指定：请给出要删除的行号或列（如“第5行”/“C列”）' };
            }
            if (op === 'clear') {
                const formatOnly = (plan.clearFormat === true && plan.clearContent === false);
                const args = {};
                if (plan.sheet) args.sheet = plan.sheet;
                if (plan.range) args.range = plan.range;
                const toolName = formatOnly ? CRUD_OP_TOOL.clearFormat : CRUD_OP_TOOL.clearRange;
                if (!formatOnly) {
                    args.clearContent = plan.clearContent !== false;
                    args.clearFormat = plan.clearFormat !== false;
                }
                return await toolRegistry.invokeTracked(turn, toolName, args, {
                    displayName, displayArgs: { '范围': where || '当前选区' },
                    ctx: { turn, sessionStore, signal }
                });
            }
            return { ok: false, error: '未知的增删改查操作: ' + op };
        } catch (e) {
            if (e?.message === 'user_aborted') throw e;
            return { ok: false, error: e?.message || String(e) };
        }
    }

    // 清洗 JS 变换路径（excel_clean_data 不可用/失败时的备选）：
    // excel_read_range 读取 → JS 确定性变换 → excel_write_range 写回（clearFirst）→ 清理残留行
    // 步骤号由 invokeTracked 统一自动分配，无需手动接续。
    async function executeCleanWithJs(turn, plan, ctx, otherOps, toolRegistry, sessionStore, signal) {
        const selAddr = (ctx.selection?.address || '').replace(/\$/g, '');
        const selSheet = ctx.selection?.sheetName || '';
        try {
            // 读取当前选区数据（对象数组，含表头）
            const readArgs = { returnFormat: 'object', hasHeaders: true };
            if (selSheet) readArgs.sheet = selSheet;
            if (selAddr) readArgs.range = selAddr;
            const readResult = await toolRegistry.invokeTracked(turn, 'excel_read_range', readArgs, {
                displayName: '读取清洗数据',
                displayArgs: { '数据范围': selAddr || '当前选区' },
                ctx: { turn, sessionStore, signal }
            });
            if (signal?.aborted) throw new Error('user_aborted');
            if (readResult?.ok === false || readResult?.error) {
                return { ok: false, error: '读取数据失败: ' + (readResult?.error || '未知错误'), fallback: true };
            }

            // 解析对象数组（数据超 maxRows 截断时 JSON 带注释无法解析 → 直接报错，不回退 VBA）
            let rows = [];
            const raw = readResult?.result;
            if (typeof raw === 'string') {
                try { rows = JSON.parse(raw); } catch (e) { return { ok: false, error: '读取数据解析失败（可能超出最大行数）', fallback: true }; }
            } else if (Array.isArray(raw)) {
                rows = raw;
            } else if (raw && Array.isArray(raw.rows)) {
                rows = raw.rows;
            }
            if (!Array.isArray(rows) || rows.length === 0) {
                return { ok: false, error: '未读取到数据，请先选中包含表头和数据的区域', fallback: true };
            }

            // JS 确定性变换（对齐原 VBA 模板语义）
            const cleaned = applyCleanOps(rows, plan, otherOps);

            // 写回（clearFirst 清空目标区域后覆盖）
            const headers = Object.keys(cleaned[0] || {});
            const values2d = [headers, ...cleaned.map(r => headers.map(h => (r[h] === undefined ? '' : r[h])))];
            const startCell = getTopLeftCell(selAddr);
            const writeArgs = {
                range: startCell,
                values: JSON.stringify(values2d),
                clearFirst: true,
                writeMode: 'overwrite'
            };
            if (selSheet) writeArgs.sheet = selSheet;
            const writeResult = await toolRegistry.invokeTracked(turn, 'excel_write_range', writeArgs, {
                displayName: '写回清洗结果',
                displayArgs: {
                    '目标': startCell,
                    '数据行数': cleaned.length,
                    '列数': headers.length
                },
                ctx: { turn, sessionStore, signal }
            });
            if (signal?.aborted) throw new Error('user_aborted');
            if (writeResult?.ok === false || writeResult?.error) {
                return { ok: false, error: '写回失败: ' + (writeResult?.error || '未知错误'), fallback: true };
            }

            // 清洗后行数减少时，物理删除原选区下方残留行（如 removeEmptyRows / fillEmpty=delete）。
            // 用 excel_delete_row 删除整行，否则残留行只是被清空内容、仍在表中（用户会说"空行没删干净"）。
            if (cleaned.length < rows.length) {
                const startRow = getRowNum(startCell);
                const deleteRow = startRow + cleaned.length + 1;
                const deleteCount = rows.length - cleaned.length;
                if (deleteCount > 0) {
                    const deleteArgs = { row: deleteRow, count: deleteCount };
                    if (selSheet) deleteArgs.sheet = selSheet;
                    await toolRegistry.invokeTracked(turn, 'excel_delete_row', deleteArgs, {
                        displayName: '删除残留行',
                        displayArgs: { '起始行': deleteRow, '行数': deleteCount },
                        ctx: { turn, sessionStore, signal }
                    });
                }
            }

            return { ok: true, result: { success: true, message: '清洗完成' }, toolCallId: null };
        } catch (e) {
            if (e?.message === 'user_aborted') throw e;
            return { ok: false, error: e?.message || String(e), fallback: true };
        }
    }

    /**
     * 清洗操作的 JS 确定性变换（对齐原 VBA 模板语义）。
     * 输入为对象数组（含表头），返回清洗后的对象数组。
     */
    function applyCleanOps(rows, plan, ops) {
        let data = rows.map(r => ({ ...r }));
        const fillMethod = plan.fillMethod || 'ffill';

        for (const op of ops) {
            switch (op) {
                case 'trimSpaces': {
                    data = data.map(row => {
                        const out = {};
                        for (const k of Object.keys(row)) {
                            let v = row[k];
                            if (typeof v === 'string') v = v.trim();
                            out[k] = v;
                        }
                        return out;
                    });
                    break;
                }
                case 'standardizeText': {
                    data = data.map(row => {
                        const out = {};
                        for (const k of Object.keys(row)) {
                            let v = row[k];
                            if (v !== null && v !== undefined && v !== '') v = String(v).trim();
                            out[k] = v;
                        }
                        return out;
                    });
                    break;
                }
                case 'fillEmpty': {
                    if (fillMethod === 'delete') {
                        // 删除任一列为空的整行
                        data = data.filter(row => Object.values(row).every(v => v !== null && v !== undefined && v !== ''));
                    } else if (fillMethod === 'zero') {
                        data = data.map(row => {
                            const out = {};
                            for (const k of Object.keys(row)) {
                                let v = row[k];
                                if (v === null || v === undefined || v === '') v = 0;
                                out[k] = v;
                            }
                            return out;
                        });
                    } else if (fillMethod === 'placeholder') {
                        // placeholder：占位填充，空值统一填"待补充"（对齐后端默认 fillValue）
                        data = data.map(row => {
                            const out = {};
                            for (const k of Object.keys(row)) {
                                let v = row[k];
                                if (v === null || v === undefined || v === '') v = '待补充';
                                out[k] = v;
                            }
                            return out;
                        });
                    } else {
                        // ffill：前值填充，首行空值用占位
                        const last = {};
                        data = data.map(row => {
                            const out = {};
                            for (const k of Object.keys(row)) {
                                let v = row[k];
                                if (v === null || v === undefined || v === '') {
                                    v = (k in last) ? last[k] : '待补充';
                                } else {
                                    last[k] = v;
                                }
                                out[k] = v;
                            }
                            return out;
                        });
                    }
                    break;
                }
                case 'normalizeDates': {
                    data = data.map(row => {
                        const out = {};
                        for (const k of Object.keys(row)) {
                            let v = row[k];
                            if (v !== null && v !== undefined && v !== '' && typeof v !== 'number') {
                                const s = String(v).trim();
                                // 仅对"日期样"值做归一化，避免误伤纯数字/普通文本
                                if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(s) || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4}/.test(s)) {
                                    const d = new Date(s.replace(/[./]/g, '-'));
                                    if (!isNaN(d.getTime())) {
                                        const y = d.getFullYear();
                                        const m = String(d.getMonth() + 1).padStart(2, '0');
                                        const day = String(d.getDate()).padStart(2, '0');
                                        v = `${y}-${m}-${day}`;
                                    }
                                }
                            }
                            out[k] = v;
                        }
                        return out;
                    });
                    break;
                }
                case 'normalizeNumbers': {
                    data = data.map(row => {
                        const out = {};
                        for (const k of Object.keys(row)) {
                            let v = row[k];
                            if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) {
                                v = Number(v);
                            }
                            out[k] = v;
                        }
                        return out;
                    });
                    break;
                }
                case 'removeEmptyRows': {
                    data = data.filter(row => Object.values(row).some(v => v !== null && v !== undefined && v !== ''));
                    break;
                }
                default:
                    break;
            }
        }
        return data;
    }

    /** 从选区地址（如 "A1:G100" / "$A$1:$G$100"）提取左上角单元格（如 "A1"） */
    function getTopLeftCell(address) {
        if (!address) return 'A1';
        const clean = String(address).split('!').pop().replace(/\$/g, '');
        const m = clean.match(/([A-Za-z]+)(\d+)/);
        return m ? (m[1] + m[2]) : 'A1';
    }

    /** 从单元格（如 "A1"）提取行号（如 1） */
    function getRowNum(cell) {
        const m = String(cell || '').match(/(\d+)/);
        return m ? parseInt(m[1], 10) : 1;
    }

    /** 从单元格（如 "A1"）提取列字母（如 "A"） */
    function getColLetter(cell) {
        const m = String(cell || '').match(/([A-Za-z]+)/);
        return m ? m[1].toUpperCase() : 'A';
    }

    /** 从起始列字母 + 列数计算结束列字母（如 "A", 7 → "G"） */
    function getEndColLetter(colLetter, colCount) {
        let num = 0;
        for (const ch of String(colLetter || 'A').toUpperCase()) {
            num = num * 26 + (ch.charCodeAt(0) - 64);
        }
        num = num + Math.max(colCount, 1) - 1;
        let result = '';
        while (num > 0) {
            num -= 1;
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26);
        }
        return result || 'A';
    }

    /** 从范围地址（如 "A1:G30" / "$A$1:$G$30"）解析总行数（含表头） */
    function getRangeHeight(address) {
        if (!address) return 0;
        const clean = String(address).split('!').pop().replace(/\$/g, '');
        const cells = clean.split(':');
        const nums = [];
        for (const c of cells) {
            const m = String(c).match(/(\d+)/);
            if (m) nums.push(parseInt(m[1], 10));
        }
        if (nums.length === 0) return 0;
        return Math.max(1, Math.max(...nums) - Math.min(...nums) + 1);
    }
    MODS.cleanCrud = {
        splitRangeAddr,
        buildHeadTailRanges,
        parseReadRows,
        executeCleanSingleOp,
        highlightDuplicateRows,
        writeCleanBack,
        buildCleanOpSummary,
        opLabel,
        crudWhere,
        resolveRowCol,
        excelColToNum,
        singleCellAddress,
        executeCrudWithMcp,
        executeCleanWithJs,
        applyCleanOps,
        getTopLeftCell,
        getRowNum,
        getColLetter,
        getEndColLetter,
        getRangeHeight
    };
})();