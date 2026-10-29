/**
 * ============================================================
 * chat-v2/ui/renderers/tool-card.js — 工具卡片纯函数子模块
 * ============================================================
 * 从 chat-renderer.js 拆出的独立子模块，全部 DOM-free：
 *   - unwrapToolResult(raw)             ：深度解包工具返回结果
 *   - extractSkillCardFromResult(raw)   ：检测 outputs.content 里的 skill_card
 *   - autoSummarizeToolResult(name,raw) ：按工具名生成友好摘要
 *   - sanitizeArgsForDisplay(obj)       ：截断过长的参数值用于展示
 *   - truncateForDisplay(result)        ：结果文本截断
 *   - formatToolErrorHtml(err, meta)    ：格式化错误展示 HTML
 *   - _generateGenericSuggestion(err)   ：通用错误兜底建议
 *   - extractPlaceholderMessage(data)   ：后端占位消息检测
 *   - extractSkillResultSummary(raw)    ：操作类技能可读摘要
 *
 * DOM 相关（renderToolCard/updateToolCard 等）仍留在 chat-renderer.js。
 * 自包含 escapeHtml（委托到 App.ChatV2.renderHtml.escapeHtml）。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;

    function escapeHtml(str) {
        if (str == null) return '';
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(str) : '';
    }

    /** 参数截断阈值：超过则截断为 "...(N chars)" */
    const MAX_ARG_VALUE_LEN = 200;

    /**
     * 清理工具参数用于显示：截断过长的字符串值（如 HTML 内容、大段代码），
     * 避免参数展开时撑爆界面。
     */
    function sanitizeArgsForDisplay(obj, depth) {
        depth = depth || 0;
        if (depth > 3) return '...';
        if (obj == null) return obj;
        if (typeof obj === 'string') {
            if (obj.length > MAX_ARG_VALUE_LEN) {
                return obj.slice(0, MAX_ARG_VALUE_LEN) + `...(${obj.length} chars)`;
            }
            return obj;
        }
        if (Array.isArray(obj)) {
            return obj.map(v => sanitizeArgsForDisplay(v, depth + 1));
        }
        if (typeof obj === 'object') {
            const out = {};
            for (const [k, v] of Object.entries(obj)) {
                out[k] = sanitizeArgsForDisplay(v, depth + 1);
            }
            return out;
        }
        return obj;
    }

    /**
     * 深度解包工具返回结果，处理各种嵌套包裹格式：
     * - { result: {...}, taskCompleted: true }  (executeLocalExcelTool 包装)
     * - { success: true, result: {...} }       (CallLocalTool 包装)
     * - { ok: true, result: {...} }            (toolRegistry.invoke 包装)
     * - "JSON string"                          (字符串形式的 JSON)
     * - 多层嵌套：{result: {result: {...}, taskCompleted: true}}
     * 返回解包后的实际数据（可能是对象、数组、字符串）
     */
    function unwrapToolResult(raw) {
        let data = raw;
        let depth = 0;
        const MAX_DEPTH = 5;

        // 如果是字符串，先尝试解析 JSON
        if (typeof data === 'string') {
            try { data = JSON.parse(data); } catch (e) { return data; }
        }

        while (data && typeof data === 'object' && depth < MAX_DEPTH) {
            depth++;
            let unwrapped = false;

            // 解包 { result: ..., taskCompleted: true/success: true } 格式
            if (data.result !== undefined) {
                const inner = data.result;
                // 如果 result 是字符串，尝试 JSON 解析
                if (typeof inner === 'string') {
                    try {
                        const parsed = JSON.parse(inner);
                        if (parsed && typeof parsed === 'object') {
                            data = parsed;
                            unwrapped = true;
                        }
                    } catch (e) {
                        // result 是非 JSON 字符串（如 "HTML 页面已在浏览器中打开: ..."），保留
                        data = inner;
                        unwrapped = true;
                    }
                } else if (inner !== null && typeof inner === 'object') {
                    data = inner;
                    unwrapped = true;
                } else {
                    data = inner;
                    unwrapped = true;
                }
            } else if (data.data !== undefined && typeof data.data === 'object' && !Array.isArray(data.data)) {
                // 注意：data 是数组（如合并结果的 data 矩阵）时绝不能继续往下剥 ——
                // 否则 {success, sourceSheets, mergedRows, data:[[…]]} 会被摘成裸矩阵，
                // 统计字段（来源表/行数/表头）全丢（2026-10-07 合并工作表渲染成满屏裸 JSON）。
                // 仅当 data 是对象（{rows,headers} / {result:[…]} 等）时才视为包裹层继续解包。
                data = data.data;
                unwrapped = true;
            } else if (data.content !== undefined) {
                // MCP 标准格式: { content: [{type: "text", text: "..."}] }
                if (Array.isArray(data.content)) {
                    const texts = data.content.filter(c => c.type === 'text').map(c => c.text);
                    if (texts.length > 0) {
                        const combined = texts.join('\n');
                        try { data = JSON.parse(combined); } catch (e) { data = combined; }
                        unwrapped = true;
                    }
                }
            }

            if (!unwrapped) break;
        }

        // 最后如果还是字符串，再尝试一次解析
        if (typeof data === 'string') {
            try { data = JSON.parse(data); } catch (e) { /* keep as string */ }
        }

        return data;
    }

    /**
     * 检测工具结果中是否包含 skill_card（在 outputs.content 中）。
     * 如果是，返回 skill 信息对象；否则返回 null。
     */
    function extractSkillCardFromResult(rawResult) {
        try {
            const data = unwrapToolResult(rawResult);
            if (!data || typeof data !== 'object') return null;

            const outs = data.outputs || data.Outputs || {};
            if (!outs || typeof outs !== 'object') return null;

            // 检查 outputs.content 是否为 skill_card JSON
            const content = outs.content;
            if (typeof content === 'string') {
                try {
                    const obj = JSON.parse(content.trim());
                    if (obj && obj.type === 'skill_card') return obj;
                } catch (e) { /* not JSON */ }
            } else if (content && typeof content === 'object' && content.type === 'skill_card') {
                return content;
            }

            // 直接在顶层检查
            if (data.type === 'skill_card') return data;

            return null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 合并工作表结果（excel_merge_sheets）→ 可读统计。
     *
     * 后端回传形态（实测 2026-10-07）：
     *   { success:true, sourceSheets:[…], mergedRows:100, dataRows:100,
     *     headerPreserved:true, usedSelectedSheets:false, data:[[表头],[行],…] }
     * 直接把这坨 JSON（含 100 行 × 9 列的 data 矩阵）铺在气泡里既长又难读
     * （用户反馈"渲染有点不好"），这里抽出关键统计交给渲染层做"摘要 + 表格 + 折叠原始 JSON"。
     *
     * @returns {{text:string, chips:Array<{label:string,value:string}>}|null} 非合并结果返回 null
     */
    function summarizeMergeResult(data) {
        if (typeof data === 'string') { try { data = JSON.parse(data); } catch (e) { return null; } }
        if (!data || typeof data !== 'object' || Array.isArray(data)) return null;
        const num = (v) => (typeof v === 'number' && isFinite(v) && v > 0) ? v : 0;
        // 命中判据：只认"合并工作表"专属标记（mergedRows/dataRows/sourceSheets）。
        // 不能用 rowCount/sheetCount 这类泛化字段判定 —— 清洗、读取、拆分等结果也会带这些字段，
        // 会被误标成"已合并 N 张工作表"（formatToolResultHtml 会对所有对象结果调用本函数）。
        const rows = num(data.mergedRows) || num(data.dataRows);
        const sheets = Array.isArray(data.sourceSheets) ? data.sourceSheets.length : num(data.mergedSheets);
        if (!rows && !sheets) return null;

        const chips = [];
        if (sheets) chips.push({ label: '来源工作表', value: sheets + ' 张' });
        if (rows) chips.push({ label: '合并行数', value: rows + ' 行' });
        if (data.headerPreserved === true) chips.push({ label: '表头', value: '已保留' });
        else if (data.headerPreserved === false) chips.push({ label: '表头', value: '未保留（整表当数据）' });
        const target = data.targetSheet || data.outputSheet || data.resultSheet || data.newSheet || data.sheetName;
        if (target) chips.push({ label: '目标工作表', value: String(target) });

        const text = '已合并 ' + (sheets ? sheets + ' 张工作表' : '工作表')
            + (rows ? ' → 共 ' + rows + ' 行' : '')
            + (data.headerPreserved === true ? '（含表头）' : '');
        return { text, chips };
    }

    /**
     * 从"原始工具结果"识别合并结果（渲染层与摘要层都用它，避免各写一遍）。
     *
     * 为什么必须"多层扫描"：工具结果常被包两层（实测 excel_merge_sheets）：
     *   { result: { success, sourceSheets:[…], mergedRows:100, data:[[表头],[100 行]] } }
     * 而 unwrapToolResult 会 result → data 一路摘到底，最终只剩数据矩阵，
     * 统计字段（来源表/行数/表头）全丢 → 气泡里只能摊一坨裸矩阵
     * （2026-10-07 用户截图：满屏 "政府客户" / 45508.0 / "华南" 的 JSON）。
     * 因此这里按包裹键（result/data/outputs/content/text）做有界广度遍历，
     * 命中"像合并结果"的对象即返回，并把该对象作为 source 一并给出，供渲染层做表格与原始 JSON。
     *
     * @returns {{text:string, chips:Array<{label:string,value:string}>, source:object}|null}
     */
    function summarizeMergeFromRaw(rawResult) {
        const queue = [rawResult];
        const visited = new Set();
        let steps = 0;
        while (queue.length && steps++ < 24) {
            let cur = queue.shift();
            if (cur == null) continue;
            if (typeof cur === 'string') {
                const s = cur.trim();
                // 也接受被引号包住的 JSON 字符串（双层编码：'"{\"success\":…}"'）
                if (s[0] !== '{' && s[0] !== '[' && s[0] !== '"') continue;   // 纯文本结果（如"已在浏览器中打开…"）跳过
                try {
                    const p = JSON.parse(s);
                    // 【双层编码】剥一层后仍是 JSON 字符串（如 "\"{\\\"success\\\":…}\""）→ 回队再剥。
                    // 旧逻辑在此处把解析结果当"非对象"直接丢弃，是识别失败的静默缺口（2026-10-09）。
                    if (typeof p === 'string') { queue.push(p); continue; }
                    cur = p;
                } catch (e) { continue; }
            }
            if (typeof cur !== 'object' || visited.has(cur)) continue;
            visited.add(cur);
            const hit = summarizeMergeResult(cur);
            if (hit) return Object.assign({}, hit, { source: cur });
            const push = (v) => { if (v != null) queue.push(v); };
            if (Array.isArray(cur)) {
                // 单元素包裹 [{…}] / content:[{type,text}]；
                // 行矩阵（100 行）不展开，避免无谓遍历
                if (cur.length <= 5) cur.forEach(push);
                continue;
            }
            // 包装键扩展到常见别名，避免 payload 藏在非标准键下被漏掉
            push(cur.result); push(cur.data); push(cur.outputs); push(cur.content); push(cur.text);
            push(cur.payload); push(cur.output); push(cur.response); push(cur.value);
        }
        return null;
    }

    /**
     * 根据工具名和返回结果自动生成友好的摘要文本。
     * 返回 null 表示无法生成摘要，回退到原始 JSON 展示。
     */
    function autoSummarizeToolResult(toolName, rawResult) {
        try {
            const data = unwrapToolResult(rawResult);
            const name = (toolName || '').replace(/^excel-local__/, '').replace(/^excel_/, '');

            // Skill 执行结果检测：如果包含 skill_card，生成 Skill 摘要
            // 工具名兼容两种形态：'skill.SkillCreator'（注册名）与 'skill_SkillCreator'（normalizeName 后的下划线形态）
            if (/^skill[._]/.test(name) || name === 'skill_exec' || name === 'execute_skill') {
                const skillCard = extractSkillCardFromResult(rawResult);
                if (skillCard) {
                    const displayName = skillCard.skillDisplayName || skillCard.skillName || 'Skill';
                    const category = skillCard.skillCategory || '';
                    const catStr = category && category !== '-' ? `（${category}）` : '';
                    return `已创建 Skill：${displayName}${catStr}`;
                }
                if (data && data.success === true) {
                    return data.message || 'Skill 执行成功';
                }
            }

            // 表格格式化（PPT/Word/Excel 选中表格）：优先展示宿主回传的"具体参数变化"摘要
            if (/^(?:ppt|word|excel)_format_(?:selected_)?table_cell$/.test(name) || name === 'format_selected_table_cell' || name === 'format_table_cell') {
                if (data && typeof data === 'object') {
                    if (data.summary) return data.summary;
                    const ch = Array.isArray(data.changes) ? data.changes : [];
                    if (ch.length) {
                        return '已修改：' + ch.map(c => `${c.param}${c.cells != null ? '(' + c.cells + '格)' : ''}`).join('、');
                    }
                    if (data.success === true) return '格式已应用';
                }
                return null;
            }

            // 读取选区/范围数据
            if (name === 'read_range' || name === 'read_selection') {
                // 格式1: 数组（MCP直接返回行对象数组 [{col1:val1}, ...]）
                if (Array.isArray(data)) {
                    const rowCount = data.length;
                    const colCount = data.length > 0 ? Object.keys(data[0]).length : 0;
                    if (rowCount > 0) {
                        return `已读取 ${rowCount} 行 × ${colCount} 列数据`;
                    }
                    return '选区为空';
                }
                // 格式2: {headers:[], rows:[]} 或 {success:true, headers:[], rows:[], rowCount:N, colCount:M, address:"..."}
                if (data && typeof data === 'object') {
                    const rows = data.rows;
                    const headers = data.headers;
                    if (Array.isArray(rows)) {
                        const colCount = Array.isArray(headers) ? headers.length : (rows.length > 0 ? Object.keys(rows[0] || {}).length : 0);
                        const range = data.address || data.range || '';
                        const rangeStr = range ? `（${range}）` : '';
                        if (rows.length > 0) {
                            return `已读取 ${rows.length} 行 × ${colCount} 列数据${rangeStr}`;
                        }
                        return '选区为空';
                    }
                    // 格式3: {success:true, result: "[...]"} (bridge string result)
                    if (data.rowCount != null || data.colCount != null) {
                        const rc = data.rowCount || 0;
                        const cc = data.colCount || 0;
                        const range = data.address || '';
                        if (rc > 0) {
                            return `已读取 ${rc} 行 × ${cc} 列数据${range ? '（' + range + '）' : ''}`;
                        }
                        return '选区为空';
                    }
                }
                return '选区为空';
            }

            // 合并工作表：只给"合并了几张表、多少行"，绝不把 data 矩阵回退成原始 JSON
            // （实测 2026-10-07：合并成功后气泡里摊着 100 行 × 9 列 JSON，用户反馈"渲染不好"）。
            if (name === 'merge_sheets' || name === 'merge_workbooks' || name === 'merge_range') {
                const ms = summarizeMergeFromRaw(rawResult);
                if (ms) return ms.text;
                if (data && typeof data === 'object' && data.success === true) return '工作表已合并';
                return null;
            }

            // 数据清洗/去重：只给"重点步骤"结论，不要把清洗结果整份数组回退成原始 JSON。

            if (name === 'clean_data' || name === 'remove_duplicates' || name === 'clean_range') {
                // 计数对象形态 {removedCount/filledCells/...}
                if (data && typeof data === 'object' && !Array.isArray(data)) {
                    const get = (...ks) => { for (const k of ks) { const v = data[k]; if (typeof v === 'number' && v > 0) return v; } return 0; };
                    const dup = get('removedCount', 'duplicateCount', 'removedRows', 'deletedRows', 'removedRowCount', 'rowsRemoved');
                    const emp = get('removedEmptyRows', 'emptyRowsRemoved', 'removedEmpty');
                    const fill = get('filledCells', 'cellsFilled', 'emptyFilled', 'fillCount');
                    const trim = get('trimmedCells', 'trimmedCount', 'spacesRemoved', 'trimmed');
                    const nd = get('normalizedDates', 'dateCellsNormalized', 'datesFixed');
                    const nn = get('normalizedNumbers', 'numberCellsNormalized', 'numbersFixed');
                    const parts = [];
                    if (dup) parts.push(`删除重复 ${dup} 行`);
                    if (emp) parts.push(`删除空行 ${emp} 行`);
                    if (fill) parts.push(`填充空值 ${fill} 个`);
                    if (trim) parts.push(`清理空格 ${trim} 处`);
                    if (nd) parts.push(`统一日期 ${nd} 处`);
                    if (nn) parts.push(`统一数字 ${nn} 处`);
                    if (parts.length) return '已清洗：' + parts.join('、');
                    if (typeof data.cleanedRows === 'number' || typeof data.uniqueCount === 'number') {
                        return `清洗完成，共 ${data.cleanedRows ?? data.uniqueCount} 行`;
                    }
                    if (data.success !== undefined || data.markedRows !== undefined) return '清洗完成（已写回当前选区）';
                    return '清洗完成';
                }
                // 数组形态（清洗结果数据）：只给行列数，不铺整份数据
                if (Array.isArray(data)) {
                    const rowCount = data.length;
                    const colCount = rowCount > 0 ? (Array.isArray(data[0]) ? data[0].length : Object.keys(data[0] || {}).length) : 0;
                    return rowCount > 0 ? `清洗完成，共 ${rowCount} 行 × ${colCount} 列数据` : '选区为空';
                }
                return '清洗完成';
            }

            // 清空选区/工作表
            if (name === 'clear_range' || name === 'clear_selection') {
                return '已清空选区数据';
            }
            if (name === 'clear_sheet') {
                const sheet = data && data.sheet ? data.sheet : '';
                return `已清空工作表${sheet ? '（' + sheet + '）' : ''}`;
            }

            // 写入数据
            if (name === 'write_range' || name === 'write_data') {
                const rows = data && data.rows ? data.rows : (Array.isArray(data) ? data.length : '');
                return rows ? `已写入 ${rows} 行数据` : '数据已写入';
            }

            // HTML浏览器展示
            if (name === 'display_html_in_browser' || name === 'open_html_in_browser') {
                return '已在浏览器中打开';
            }

            // 导出 / 保存文件：优先展示落盘路径（配合卡片「打开文件夹 / 复制路径」）
            if (/export_(pdf|csv|chart_image)|save_workbook|save_as|save_presentation|save_document/i.test(name)
                || /_export_(pdf|csv|chart_image)$/i.test(toolName || '')) {
                const pathInfo = extractExportFilePath(toolName, rawResult, null);
                if (pathInfo && pathInfo.path) {
                    const base = pathInfo.path.replace(/^.*[\\/]/, '') || pathInfo.path;
                    if (/export_pdf|_export_pdf/i.test(toolName || name)) return `已导出 PDF：${base}`;
                    if (/export_csv|_export_csv/i.test(toolName || name)) return `已导出 CSV：${base}`;
                    if (/chart_image/i.test(toolName || name)) return `已导出图表图片：${base}`;
                    return `已保存：${base}`;
                }
                if (data && data.success === true) {
                    if (/export_pdf|_export_pdf/i.test(toolName || name)) return '已导出 PDF';
                    if (/export_csv|_export_csv/i.test(toolName || name)) return '已导出 CSV';
                    return data.message || '文件已保存';
                }
            }

            // 创建/打开/保存工作簿
            if (name === 'create_workbook') {
                const fn = data && data.fileName ? data.fileName : '';
                return fn ? `已创建工作簿：${fn}` : '工作簿已创建';
            }
            if (name === 'open_workbook') {
                return '工作簿已打开';
            }
            if (name === 'save_workbook') {
                const pathInfo = extractExportFilePath(toolName, rawResult, null);
                if (pathInfo && pathInfo.path) {
                    const base = pathInfo.path.replace(/^.*[\\/]/, '') || pathInfo.path;
                    return `工作簿已保存：${base}`;
                }
                return '工作簿已保存';
            }

            // 创建/删除/重命名工作表
            if (name === 'create_sheet') {
                const sn = data && data.sheetName ? data.sheetName : '';
                return sn ? `已创建工作表：${sn}` : '工作表已创建';
            }
            if (name === 'delete_sheet') {
                return '工作表已删除';
            }
            if (name === 'rename_sheet') {
                return '工作表已重命名';
            }

            // 合并/拆分单元格
            if (name === 'merge_cells') {
                return '单元格已合并';
            }
            if (name === 'unmerge_cells') {
                return '已取消合并';
            }

            // 格式设置
            if (name && (name.startsWith('set_') || name.includes('format') || name.includes('style'))) {
                return '格式已应用';
            }

            // 排序/筛选
            if (name === 'sort_range' || name === 'apply_sort') {
                return '数据已排序';
            }
            if (name === 'auto_filter' || name === 'apply_filter') {
                return '筛选已应用';
            }

            // 插入/删除行列
            if (name && (name.startsWith('insert_') || name.startsWith('delete_')) && (name.includes('row') || name.includes('column'))) {
                return '操作完成';
            }

            // 通用成功判断
            if (data && typeof data === 'object') {
                if (data.success === true) {
                    return data.summary || data.message || '操作成功';
                }
                if (data.ok === true) {
                    return data.message || '操作成功';
                }
            }

            return null; // 无法生成摘要，回退到原始 JSON
        } catch (e) {
            return null;
        }
    }

    function truncateForDisplay(result) {
        if (result == null) return '(无结果)';
        const str = typeof result === 'string' ? result : JSON.stringify(result, null, 2);
        if (str.length <= 2000) return str;
        return str.slice(0, 2000) + '\n...(结果已截断，共 ' + str.length + ' 字符)';
    }

    /**
     * 格式化工具错误，展示：错误类型徽章 + 原始错误 + 原因 + 修复建议
     * @param {string} errorText - 原始错误文本
     * @param {Object} [meta] - 可选的错误分类信息
     */
    function formatToolErrorHtml(errorText, meta) {
        const rawError = errorText == null ? '未知错误' : String(errorText);
        const escaped = escapeHtml(rawError);
        const toolName = meta?.toolName || '';

        // 构造分类信息块（若有）
        const hasMeta = meta && (meta.errorType || meta.errorReason || meta.errorSuggestion);
        let metaHtml = '';
        if (hasMeta) {
            const parts = [];
            if (meta.errorType) {
                parts.push('<div class="err-meta-row err-type-row"><span class="err-label">错误类型</span><span class="err-type-badge">' + escapeHtml(meta.errorType) + '</span></div>');
            }
            if (meta.errorReason) {
                parts.push('<div class="err-meta-row"><span class="err-label">原因</span><span class="err-reason">' + escapeHtml(meta.errorReason) + '</span></div>');
            }
            if (meta.errorSuggestion) {
                parts.push('<div class="err-meta-row"><span class="err-label">建议</span><span class="err-suggestion">' + escapeHtml(meta.errorSuggestion) + '</span></div>');
            }
            metaHtml = '<div class="err-meta-panel">' + parts.join('') + '</div>';
        }

        // Python 堆栈跟踪特殊处理
        if (/Traceback.*most recent call last/i.test(escaped) || /File.*line \d+/.test(escaped)) {
            const lines = escaped.split('\n');
            const errorLines = [];
            let inTraceback = false;
            let errorType = '';
            let errorMsg = '';

            lines.forEach(function(line) {
                if (/Traceback/.test(line)) {
                    inTraceback = true;
                    errorLines.push('<span class="err-traceback">' + line + '</span>');
                } else if (/File\s+.*line\s+\d+/.test(line)) {
                    errorLines.push('<span class="err-file">' + line + '</span>');
                } else if (/^\w+\s*(Error|Exception|Warning|Type|Key|Value|Index|Name|Attribute|Import|Syntax|Indentation|Tab|Recursion|Overflow|Underflow|Memory|System|Arithmetic|NotImplemented|Stop|Generator|Iterator|StopIteration|Async|Await|Cancelled|InvalidState|Operation|Aggregate|TargetInvocation|Runtime|Serialization).*/.test(line.trim())) {
                    inTraceback = false;
                    const match = line.match(/^(\w+(?:Error|Exception|Warning|Type|Key|Value|Index|Name|Attribute|Import|Syntax|Indentation|Tab|Recursion|Overflow|Underflow|Memory|System|Arithmetic|NotImplemented|Stop|Generator|Iterator))(.*)/);
                    if (match) {
                        errorType = match[1];
                        errorMsg = match[2] || '';
                    }
                    errorLines.push('<span class="err-type">' + line + '</span>');
                } else if (inTraceback || /^\s+/.test(line)) {
                    errorLines.push('<span class="err-fragment">' + line + '</span>');
                } else {
                    errorLines.push(line);
                }
            });

            const preview = errorType
                ? '<div class="err-summary"><span class="err-badge">⚠️</span> <span class="err-name">' + escapeHtml(errorType) + '</span><span class="err-msg">' + escapeHtml(errorMsg) + '</span></div>'
                : '';

            const fullTrace = '<pre class="err-traceback-block">' + errorLines.join('\n') + '</pre>';
            const truncated = rawError.length > 1500
                ? '<div class="err-truncation">错误过长，仅显示前 1500 字符</div>'
                : '';

            return '<div class="exec-tool-error python-error">' + metaHtml + preview + fullTrace + truncated + '</div>';
        }

        // MCP 工具标准错误格式（{"success": false, "message": "..."}）
        const isToolError = /^\{"success":\s*false/i.test(rawError);
        if (isToolError) {
            try {
                const obj = JSON.parse(rawError);
                const msg = obj.result || obj.message || rawError;
                const truncatedText = msg.length > 500
                    ? msg.slice(0, 500) + '...(错误已截断)'
                    : msg;
                return '<div class="exec-tool-error">' + metaHtml + '<div class="err-raw"><span class="err-badge">⚠️</span> <span class="err-raw-text">' + escapeHtml(truncatedText) + '</span></div></div>';
            } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOL', '工具卡片操作失败:', e && e.message); }
        }

        // 通用错误
        const truncatedText = rawError.length > 500
            ? rawError.slice(0, 500) + '...(错误已截断)'
            : rawError;

        // 未提供 meta 时，尝试对常见错误给出通用建议
        let genericSuggestion = '';
        if (!hasMeta) {
            const suggestion = _generateGenericSuggestion(rawError, toolName);
            if (suggestion) {
                genericSuggestion = '<div class="err-meta-panel"><div class="err-meta-row"><span class="err-label">建议</span><span class="err-suggestion">' + escapeHtml(suggestion) + '</span></div></div>';
            }
        }

        return '<div class="exec-tool-error">' + metaHtml + genericSuggestion + '<div class="err-raw"><span class="err-badge">⚠️</span> <span class="err-raw-text">' + escapeHtml(truncatedText) + '</span></div></div>';
    }

    /**
     * 为未被 classifyError 覆盖的通用错误生成兜底建议
     */
    function _generateGenericSuggestion(errText, toolName) {
        if (!errText) return '';
        const t = toolName || '';
        if (/tool.*not.*found|Tool not found/i.test(errText)) {
            if (/skill/i.test(t)) {
                return 'Skill 由系统根据用户意图自动路由，请勿直接调用 skill_xxx 工具。请在 Skill 面板确认该 Skill 已启用。';
            }
            return '该工具未注册或已被移除，请在工具面板刷新工具列表，或改用其他可用工具。';
        }
        if (/executor.*not.*available|Skill executor/i.test(errText)) {
            const host = (window.SmartPilotHost && window.SmartPilotHost.label) || 'Office';
            return 'Skill 执行器未就绪，请确认 ' + host + ' 插件已完全加载，尝试重启 ' + host + '。';
        }
        if (/excelBridge.*not|WebView2|hostObjects/i.test(errText)) {
            const host = (window.SmartPilotHost && window.SmartPilotHost.label) || 'Office';
            return '宿主通信桥接未建立，请确认已打开 ' + host + ' 且插件已加载，尝试重启 ' + host + '。';
        }
        if (/不能为空|缺少参数|missing.*param/i.test(errText)) {
            return '请检查必填参数是否完整（如工作表名、单元格范围、文件名等），确认选区已选中。';
        }
        if (/不存在|not.*found|DISP_E_BADINDEX/i.test(errText)) {
            return '请确认工作表/单元格范围存在，检查地址格式是否正确。';
        }
        if (/empty|数据为空|选区.*空/i.test(errText)) {
            return '当前选区为空，请先选中包含数据的单元格区域后重试。';
        }
        if (/protected|保护|只读/i.test(errText)) {
            const host = (window.SmartPilotHost && window.SmartPilotHost.label) || 'Office';
            return '目标文档受保护，请在 ' + host + ' 中取消保护后重试。';
        }
        if (/timeout|超时/i.test(errText)) {
            return '请求超时，请检查网络或稍后重试，必要时拆分任务。';
        }
        return '';
    }

    /**
     * 检测后端占位消息（Skill 未真正执行，仅返回"已识别/由前端处理"提示）。
     * 命中时返回友好提示文案；非占位返回 null。
     */
    function extractPlaceholderMessage(data) {
        if (!data || typeof data !== 'object') return null;
        const msg = (typeof data.message === 'string' && data.message)
            || (data.outputs && typeof data.outputs.message === 'string' && data.outputs.message)
            || '';
        if (!msg) return null;
        if (msg.includes('已识别，等待 AI 生成响应') || msg.includes('由前端 Spec 引擎处理')) {
            return msg;
        }
        return null;
    }

    /**
     * 从操作类技能结果中提取可读摘要（用于工具卡片常显区）。
     * 剥离 "Skill 'xxx' Hybrid 执行成功：" 前缀与 "执行步骤：..." 列表，
     * 只保留结果摘要。
     */
    function extractSkillResultSummary(rawResult) {
        try {
            const data = unwrapToolResult(rawResult);
            if (!data || typeof data !== 'object') return '';
            const outs = data.outputs || data.Outputs || {};
            const msg = (typeof outs.message === 'string' && outs.message)
                || (typeof data.message === 'string' && data.message)
                || '';
            if (!msg) return '';
            // 剥离 "Skill 'xxx' Hybrid 执行成功：" 前缀
            let text = msg.replace(/^Skill\s+['"][^'"]*['"]\s+Hybrid\s+执行成功[：:]\s*/i, '');
            // 剥离 "执行步骤：..." 列表（绿面板已展示步骤，避免重复）
            const stepIdx = text.search(/执行步骤[：:]/);
            if (stepIdx >= 0) text = text.slice(0, stepIdx);
            return text.replace(/[\r\n]+$/, '').trim();
        } catch (e) {
            return '';
        }
    }

    /**
     * 从工具结果 / 参数中提取导出或保存的文件路径。
     * 兼容 COM 斜杠丢失：统一走 normalizePathSeparators（若可用）。
     * @returns {{ path: string, source: string }|null}
     */
    function extractExportFilePath(toolName, rawResult, args) {
        try {
            const util = App.ChatV2 && App.ChatV2.toolRegistryUtil;
            const norm = (util && typeof util.normalizePathSeparators === 'function')
                ? util.normalizePathSeparators.bind(util)
                : function (p) { return String(p || '').replace(/\//g, '\\').replace(/\\+/g, '\\'); };

            const pickPath = function (v) {
                if (v == null) return '';
                if (typeof v !== 'string') return '';
                const s = v.trim();
                if (!s) return '';
                // Windows 盘符或 UNC
                if (!/^[A-Za-z]:[\\/]/.test(s) && !/^\\\\/.test(s)) return '';
                return norm(s);
            };

            const fromArgs = args && (pickPath(args.filePath) || pickPath(args.path)
                || pickPath(args.outputPath) || pickPath(args.savePath) || pickPath(args.fullPath));
            const data = unwrapToolResult(rawResult);
            let fromResult = '';
            if (typeof data === 'string') {
                fromResult = pickPath(data);
            } else if (data && typeof data === 'object') {
                fromResult = pickPath(data.filePath) || pickPath(data.path) || pickPath(data.savedPath)
                    || pickPath(data.outputPath) || pickPath(data.fullPath) || pickPath(data.FilePath);
                if (!fromResult && typeof data.result === 'string') fromResult = pickPath(data.result);
                if (!fromResult && data.result && typeof data.result === 'object') {
                    const r = data.result;
                    fromResult = pickPath(r.filePath) || pickPath(r.path) || pickPath(r.savedPath);
                }
            }
            const path = fromResult || fromArgs;
            if (!path) return null;
            return { path: path, source: fromResult ? 'result' : 'args' };
        } catch (e) {
            return null;
        }
    }

    App.ChatV2.renderers.toolCard = {
        sanitizeArgsForDisplay,
        unwrapToolResult,
        extractSkillCardFromResult,
        summarizeMergeResult,
        summarizeMergeFromRaw,
        autoSummarizeToolResult,
        truncateForDisplay,
        formatToolErrorHtml,
        _generateGenericSuggestion,
        extractPlaceholderMessage,
        extractSkillResultSummary,
        extractExportFilePath
    };
    ChatLog.info?.('CHAT-V2', '✅ renderers/tool-card 已加载');
})();