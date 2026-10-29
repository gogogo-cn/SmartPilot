/**
 * ============================================================
 * context-engine.js — Context Engine
 * ============================================================
 * 统一运行时上下文收集、格式化与注入引擎。
 *
 * 核心类：
 *   - ContextEngine
 *     - gather(contextHints) — 收集 referenceRegions / selection / attachments / history
 *     - render(blockNames, ctx) — 按名称渲染上下文块
 *     - renderReferenceRegions / renderSelection / renderAttachments / renderHistory
 *
 * 依赖：
 *   - App 命名空间（由 app-bootstrap.js / shared-utils.js / app-chat-ui.js 等提供）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    /**
     * 安全解析 JSON 字符串
     * @param {*} raw - 原始值
     * @returns {*} 解析后的对象或原始值
     */
    function safeJsonParse(raw) {
        if (raw === null || raw === undefined) return raw;
        if (typeof raw === 'string') {
            try {
                return JSON.parse(raw);
            } catch (e) {
                return raw;
            }
        }
        return raw;
    }

    /**
     * 统一上下文引擎
     */
    class ContextEngine {
        constructor() {
            this.blockRenderers = {
                referenceRegions: this.renderReferenceRegions.bind(this),
                selection: this.renderSelection.bind(this),
                attachments: this.renderAttachments.bind(this),
                history: this.renderHistory.bind(this)
            };
        }

        /**
         * 收集当前运行时上下文
         * @param {Object} contextHints - 调用方提供的额外提示
         * @returns {Promise<Object>} 上下文对象（含 _degradations 降级记录）
         */
        async gather(contextHints) {
            contextHints = contextHints || {};

            const ctx = {
                referenceRegions: [],
                selection: null,
                attachments: [],
                history: [],
                // 【降级提示】收集过程中的降级记录：[{block, message}]。
                // 结构化产出而非 UI 耦合——引擎只负责如实记录，展示层（chat-entry）决定提示方式与时机。
                _degradations: []
            };

            try {
                ctx.referenceRegions = await this._gatherReferenceRegions(contextHints);
            } catch (e) {
                ChatLog.warn('ContextEngine', '收集引用区域失败:', e && e.message ? e.message : e);
                ctx.referenceRegions = [];
                ctx._degradations.push({ block: 'referenceRegions', message: '引用区域读取失败' });
            }

            try {
                ctx.selection = await this._gatherSelection(contextHints);
            } catch (e) {
                ChatLog.warn('ContextEngine', '收集当前选区失败:', e && e.message ? e.message : e);
                ctx.selection = null;
                ctx._degradations.push({ block: 'selection', message: '选区读取失败' });
            }

            // 标记选区来源，供上层路由判断
            if (ctx.selection) {
                ctx.selection.isAutoCaptured = ctx.selection._autoCaptured === true;
                delete ctx.selection._autoCaptured;
            }

            // 【降级提示】用户消息明确指代工作表/数据，但显式选区与自动捕获均未得到有效数据
            // ——模型将拿不到任何表格内容，需如实告知（空表/单格选区也适用此提示）。
            if (!ctx.selection && this._shouldAutoCaptureWorksheet(contextHints)) {
                ctx._degradations.push({ block: 'selection', message: '未捕获到表格数据（当前选区为空或数据读取失败）' });
            }

            try {
                ctx.attachments = this._gatherAttachments(contextHints);
            } catch (e) {
                ChatLog.warn('ContextEngine', '收集附件失败:', e && e.message ? e.message : e);
                ctx.attachments = [];
                ctx._degradations.push({ block: 'attachments', message: '附件列表读取失败' });
            }

            try {
                ctx.history = await this._gatherHistory(contextHints);
            } catch (e) {
                ChatLog.warn('ContextEngine', '收集历史记录失败:', e && e.message ? e.message : e);
                ctx.history = [];
                ctx._degradations.push({ block: 'history', message: '对话历史读取失败' });
            }

            return ctx;
        }

        /**
         * 收集引用区域
         */
        async _gatherReferenceRegions(contextHints) {
            // 1. 优先使用 ReferenceAreaManager（真实数据：工作簿、工作表、地址、表头、行）
            if (App.ReferenceAreaManager && typeof App.ReferenceAreaManager.getReferenceRegionsContext === 'function') {
                try {
                    const regions = await App.ReferenceAreaManager.getReferenceRegionsContext();
                    if (regions && regions.length > 0) return regions;
                } catch (e) {
                    ChatLog.warn('ContextEngine', 'ReferenceAreaManager 获取引用区域失败:', e && e.message ? e.message : e);
                }
            }

            // 【清理】原步骤2 曾调用 excelBridge.GetRangeDataForReferences，但 VB 桥从未实现该方法：
            // WebView2 宿主对象对任意成员访问都返回 function 桩，typeof 检查形同虚设，
            // 每次调用必报 0x80020006（DISP_E_UNKNOWNNAME）并漏出 unhandled promise rejection。
            // 真实数据源是 ReferenceAreaManager（步骤1），无数据时走 DOM 兜底即可。
            return this._readReferenceRegionsFromDom();
        }

        /**
         * 从 DOM 读取引用区域地址（无 Excel 数据时的降级）
         */
        _readReferenceRegionsFromDom() {
            const list = document.getElementById('reference-list');
            if (!list) return [];

            const items = list.querySelectorAll('.reference-item');
            if (!items.length) return [];

            const regions = [];
            items.forEach(function(item, idx) {
                const addrEl = item.querySelector('.reference-address');
                if (!addrEl) return;

                const address = addrEl.textContent.trim();
                regions.push({
                    index: idx + 1,
                    address: address,
                    workbookName: '',
                    sheetName: '',
                    headers: [],
                    rows: [],
                    sampleRows: [],
                    rowCount: 0,
                    colCount: 0
                });
            });

            return regions;
        }

        /**
         * 收集当前选区
         *
         * 策略：
         * 1. 优先使用用户显式选区（GetSelectedRangeData）。
         * 2. 若显式选区为空/无意义，且用户消息有 Excel 操作意图（图表/汇总/分析/排序/筛选...），
         *    自动捕获当前活动工作表的已用范围（GetUsedRangeData）作为"默认数据源"。
         *    这使"帮我做个饼图""按部门汇总"等请求无需用户手动框选数据即可进入管道。
         */
        async _gatherSelection(contextHints) {
            // 1. 优先获取用户显式选区
            if (App.excelBridge && typeof App.excelBridge.GetSelectedRangeData === 'function') {
                try {
                    const raw = await App.excelBridge.GetSelectedRangeData();
                    const data = safeJsonParse(raw);
                    if (this._isMeaningfulSelection(data)) return data;
                } catch (e) {
                    ChatLog.warn('ContextEngine', 'GetSelectedRangeData 失败:', e && e.message ? e.message : e);
                }
            }

            // 2. 自动兜底：消息有 Excel 操作意图时，捕获当前活动工作表已用范围
            //    但要排除"纯写操作"（导出/保存/打印/重命名工作表/拆分工作簿等）——
            //    这些操作工具自己会管工作表，不需要先读 UsedRange 数据。
            if (this._shouldAutoCaptureUsedRange(contextHints)) {
                if (App.excelBridge && typeof App.excelBridge.GetUsedRangeData === 'function') {
                    try {
                        const raw = await App.excelBridge.GetUsedRangeData();
                        const data = safeJsonParse(raw);
                        if (this._isMeaningfulSelection(data)) {
                            data._autoCaptured = true;
                            ChatLog.info('ContextEngine', '已自动捕获当前工作表已用范围:', data.address);
                            return data;
                        }
                    } catch (e) {
                        ChatLog.warn('ContextEngine', 'GetUsedRangeData 失败:', e && e.message ? e.message : e);
                    }
                }
            }

            return null;
        }

        /**
         * 判断用户消息是否有 Excel 操作意图（用于自动捕获 UsedRange 作为默认数据源）。
         *
         * 覆盖范围远大于原 _shouldAutoCaptureWorksheet 的"当前工作表"关键词：
         * 图表/汇总/分析/排序/筛选/透视/数据清洗/格式化/拆分/合并/查找/定位等全部覆盖。
         * 对纯闲聊/天气/新闻/翻译等无关问句返回 false，避免误伤。
         */
        _hasExcelOpIntent(contextHints) {
            const msg = String((contextHints && contextHints.message) || '').trim();
            if (!msg) return false;
            // 宽泛但安全的 Excel 操作意图词：覆盖"做/画/建/创建/生成/汇总/统计/分析/排序/筛选/透视/合并/拆分/清洗/格式化/美化/高亮/填充/查找/定位/转换/转置/去重/删除/插入/导出/导入/打印"等动词
            // 及 "图表/饼图/柱图/折线图/柱状图/条形图/散点图/面积图/雷达图/直方图/趋势图/仪表盘/图/表格/数据/工作表/列/行/单元格/区域/范围/selection/range/chart/graph/pivot/sort/filter/aggregate/summary/clean/format/beautify" 等名词
            const re = /(饼图|柱图|折线图|柱状图|条形图|散点图|面积图|雷达图|直方图|趋势图|仪表盘|图[表形]|chart|graph|visuali[sz]e|plot|diagram|图形化|画图|制图|可视化|汇总|统计|分析|总结|解读|洞察|评估|计算|求和|平均|最大|最小|计数|排名|百分比|比例|分布|趋势|对比|比较|差异|相关性|回归|预测|pivot|透视|交叉表|筛选|排序|过滤|条件格式|高亮|颜色标记|删除重复|去重|remove.?duplicate|drop.?duplicate|unique|合并|拆分|转置|transpose|清洗|clean|脏数据|异常值|缺失值|填充|fill|格式化|format|beautif|美化|样式|主题|数字格式|日期格式|查找|search|find|定位|替换|替换为|导出|export|导入|import|print|打印|邮件|email|PDF|pdf|CSV|csv|Excel|工作表|sheet|worksheet|表格|数据|selection|range|区域|范围|单元格|cell|列|column|行|row|公式|formula|函数|function|VBA|宏|macro|power query|power.?pivot|power.?bi)/i;
            return re.test(msg);
        }

        /**
         * 判断是否应该自动捕获 UsedRange 作为默认数据源。
         *
         * 与 _hasExcelOpIntent 的区别：后者"宽泛覆盖"（任何含 Excel 关键词的消息都命中），
         * 但这里需要精准区分——哪些操作"必须先看数据"（读/分析/图表/汇总），哪些"自己会管"（纯写）。
         *
         * 规则：
         *   1) 有"读/分析意图"且无"纯写操作覆盖" → 捕获 ✅
         *   2) 只有"纯写操作"（导出/保存/打印/重命名工作表/拆分工作簿…） → 跳过 ❌
         *      ——这些工具自己知道怎么操作工作表，不需要预先读 UsedRange
         *   3) 两者都没有 → 跳过
         *
         * 典型误伤例："将当前工作表导出为 CSV"
         *   含"导出"+"CSV"+"工作表" → _hasExcelOpIntent=true → 旧逻辑必调 GetUsedRangeData
         *   但导出工具自己会调 ws.UsedRange，完全不需要 JS 层先读 → 纯浪费 + 报错风险
         */
        _shouldAutoCaptureUsedRange(contextHints) {
            const msg = String((contextHints && contextHints.message) || '').trim();
            if (!msg) return false;

            // 这些操作"必须先知道数据长什么样"才能做——需要 UsedRange 数据
            const READ_ANALYSIS_RE = /饼图|柱图|折线图|可视化|汇总|统计|分析|解读|洞察|评估|计算|求和|平均|均值|最大|最小|计数|排名|百分比|比例|分布|趋势|对比|比较|差异|相关性|回归|预测|pivot|透视|交叉表|筛选|排序|过滤|条件格式|高亮|颜色标记|删除重复|去重|remove.?duplicate|drop.?duplicate|unique|合并单元格|拆分|转置|transpose|清洗|clean|脏数据|异常值|缺失值|填充|fill|查找|search|find|定位|替换|公式|formula|函数计算/i;

            // 这些操作"不需要预先看数据内容"——工具自己会管工作表对象
            const WRITE_ONLY_RE = /(导出|export|保存为|存为|另存为|save.?as|打印|print|重命名|rename|删除工作表|删除.?sheet|delete.?sheet|新建工作表|创建.?sheet|insert.?sheet|加密码|保护.?工作表|拆分工作簿|合并工作簿|邮件|email|PDF|pdf|CSV|csv|写入.*工作表|写入.*文件|保存.*文件|刷新|refresh|关闭|close)/i;

            const needRead = READ_ANALYSIS_RE.test(msg);
            const writeOnly = WRITE_ONLY_RE.test(msg);

            // 读/分析优先，但如果消息明确只做纯写操作 → 短路跳过
            return needRead && !writeOnly;
        }

        /**
         * 判断选区是否有意义（至少包含有效表头，且不止单个单元格）
         */
        _isMeaningfulSelection(data) {
            if (!data || data.error) return false;
            const headers = Array.isArray(data.headers) ? data.headers : [];
            const rows = Array.isArray(data.rows) ? data.rows : [];
            return headers.length > 0 && (headers.length > 1 || rows.length > 0);
        }

        /**
         * 判断是否应自动捕获当前工作表（基于用户消息意图）
         */
        _shouldAutoCaptureWorksheet(contextHints) {
            const message = (contextHints && contextHints.message) || '';
            if (!message) return false;
            const worksheetPattern = /(当前工作表|这个工作表|此工作表|当前sheet|当前表格|当前数据|活动工作表|active sheet|current worksheet|this sheet|分析.*数据|查看.*数据)/i;
            return worksheetPattern.test(message);
        }

        /**
         * 收集文件附件
         */
        _gatherAttachments(contextHints) {
            const files = [];
            const seen = new Set();

            function addFile(file) {
                if (!file) return;
                const key = (file.name || '') + '|' + (file.size || 0) + '|' + (file.lastModified || 0);
                if (seen.has(key)) return;
                seen.add(key);
                files.push(file);
            }

            // 1. DOM 元素上存储的附件
            const preview = document.getElementById('attachment-preview');
            if (preview && preview._storedFiles && preview._storedFiles.length > 0) {
                preview._storedFiles.forEach(addFile);
            }

            // 2. App 命名空间
            if (App.selectedFiles && App.selectedFiles.length > 0) {
                App.selectedFiles.forEach(addFile);
            } else if (App.selectedFile) {
                addFile(App.selectedFile);
            }

            // 3. 全局备份
            if (window.__backupSelectedFiles && window.__backupSelectedFiles.length > 0) {
                window.__backupSelectedFiles.forEach(addFile);
            } else if (window.__backupSelectedFile) {
                addFile(window.__backupSelectedFile);
            }

            // 4. file input 兜底
            const fileInput = document.getElementById('file-input');
            if (fileInput && fileInput.files && fileInput.files.length > 0) {
                Array.prototype.forEach.call(fileInput.files, addFile);
            }

            return files;
        }

        /**
         * 收集最近对话历史
         */
        async _gatherHistory(contextHints) {
            // 若不存在则挂载轻量适配器
            this._ensureHistoryManager();

            let turns = null;
            if (App.historyManager && typeof App.historyManager.getRecentTurns === 'function') {
                try {
                    turns = App.historyManager.getRecentTurns(3);
                    if (turns && typeof turns.then === 'function') {
                        turns = await turns;
                    }
                } catch (e) {
                    ChatLog.warn('ContextEngine', 'historyManager.getRecentTurns 失败:', e && e.message ? e.message : e);
                }
            }

            if (turns && turns.length > 0) return turns;

            // 兜底：使用 chatContextManager
            if (window.chatContextManager && typeof window.chatContextManager.getRecentHistory === 'function') {
                try {
                    return window.chatContextManager.getRecentHistory(3);
                } catch (e) {
                    ChatLog.warn('ContextEngine', 'chatContextManager.getRecentHistory 失败:', e && e.message ? e.message : e);
                }
            }

            return [];
        }

        /**
         * 确保 App.historyManager 存在（适配 chatContextManager）
         */
        _ensureHistoryManager() {
            if (App.historyManager) return;
            if (window.chatContextManager && typeof window.chatContextManager.getRecentHistory === 'function') {
                App.historyManager = {
                    getRecentTurns: function(count) {
                        return window.chatContextManager.getRecentHistory(count);
                    }
                };
                ChatLog.info('ContextEngine', '已挂载 historyManager 适配器到 chatContextManager');
            }
        }

        /**
         * 根据 blockNames 按顺序渲染上下文块
         * @param {string[]} blockNames - 块名称数组
         * @param {Object} ctx - 上下文对象
         * @returns {string} 拼接后的上下文文本
         */
        render(blockNames, ctx) {
            if (!Array.isArray(blockNames)) return '';
            ctx = ctx || {};

            const blocks = [];
            for (let i = 0; i < blockNames.length; i++) {
                const name = blockNames[i];
                const renderer = this.blockRenderers[name];
                if (typeof renderer !== 'function') {
                    ChatLog.warn('ContextEngine', '未知上下文块:', name);
                    continue;
                }

                const value = ctx[name];
                const block = renderer(value);
                if (block && block.trim().length > 0) {
                    blocks.push(block);
                }
            }

            return blocks.join('\n\n');
        }

        /**
         * 格式化引用区域
         */
        renderReferenceRegions(regions) {
            if (!regions || regions.length === 0) return '';

            const lines = ['【引用区域】'];

            regions.forEach(function(region, idx) {
                const index = region.index || (idx + 1);
                const address = region.address || '';
                const workbookName = region.workbookName || '';
                const sheetName = region.sheetName || '';
                const headers = Array.isArray(region.headers) ? region.headers : [];
                const rows = Array.isArray(region.rows) ? region.rows : [];
                const sampleRows = Array.isArray(region.sampleRows)
                    ? region.sampleRows
                    : rows.slice(0, 5);
                const rowCount = region.rowCount || rows.length;
                const colCount = region.colCount || headers.length;

                lines.push('区域 ' + index + ':');
                if (address) lines.push('  地址: ' + address);
                if (workbookName) lines.push('  工作簿: ' + workbookName);
                if (sheetName) lines.push('  工作表: ' + sheetName);
                if (headers.length > 0) lines.push('  表头: ' + headers.join(', '));
                if (sampleRows.length > 0) {
                    lines.push('  样本数据 (' + sampleRows.length + '/' + rowCount + ' 行 × ' + colCount + ' 列):');
                    sampleRows.forEach(function(row) {
                        const cells = Array.isArray(row) ? row : [row];
                        lines.push('    ' + cells.join(' | '));
                    });
                }
                lines.push('');
            });

            return lines.join('\n').trim();
        }

        /**
         * 格式化当前选区
         */
        renderSelection(selection) {
            if (!selection) return '';

            const address = selection.address || '';
            const workbookName = selection.workbookName || '';
            const sheetName = selection.sheetName || '';
            const headers = Array.isArray(selection.headers) ? selection.headers : [];
            const rows = Array.isArray(selection.rows) ? selection.rows : [];

            const lines = ['【当前选区】'];
            if (address) lines.push('地址: ' + address);
            if (workbookName) lines.push('工作簿: ' + workbookName);
            if (sheetName) lines.push('工作表: ' + sheetName);
            if (headers.length > 0) lines.push('表头: ' + headers.join(', '));
            if (rows.length > 0) {
                lines.push('样本数据 (' + Math.min(rows.length, 5) + '/' + rows.length + ' 行):');
                rows.slice(0, 5).forEach(function(row) {
                    const cells = Array.isArray(row) ? row : [row];
                    lines.push('  ' + cells.join(' | '));
                });
            }

            return lines.join('\n').trim();
        }

        /**
         * 格式化文件附件
         */
        renderAttachments(attachments) {
            if (!attachments || attachments.length === 0) return '';

            const lines = ['【文件附件】'];
            attachments.forEach(function(file, idx) {
                const name = file.name || '';
                const size = file.size || 0;
                const type = file.type || '';

                lines.push('附件 ' + (idx + 1) + ': ' + name);
                if (size > 0) lines.push('  大小: ' + ContextEngine.prototype._formatFileSize(size));
                if (type) lines.push('  类型: ' + type);
            });

            return lines.join('\n').trim();
        }

        /**
         * 格式化历史记录
         */
        renderHistory(history) {
            if (!history || history.length === 0) return '';

            const lines = ['【最近对话历史】'];
            history.forEach(function(turn, idx) {
                let role = '未知';
                if (turn.role === 'user') role = '用户';
                else if (turn.role === 'assistant') role = 'AI';
                else if (turn.role) role = String(turn.role);

                const content = (turn.content !== undefined && turn.content !== null) ? String(turn.content) : '';
                lines.push('回合 ' + (idx + 1) + ' [' + role + ']:');
                if (content) lines.push(content);
            });

            return lines.join('\n').trim();
        }

        /**
         * 格式化文件大小
         */
        _formatFileSize(bytes) {
            bytes = bytes || 0;
            if (bytes === 0) return '0 B';
            const k = 1024;
            const sizes = ['B', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
        }
    }

    // 注册到 App 命名空间
    if (typeof App !== 'undefined') {
        App.ContextEngine = new ContextEngine();
        ChatLog.info('ContextEngine', '上下文引擎已挂载到 App.ContextEngine');
    } else {
        ChatLog.warn('ContextEngine', 'App 命名空间未找到，ContextEngine 未挂载');
    }
})();
