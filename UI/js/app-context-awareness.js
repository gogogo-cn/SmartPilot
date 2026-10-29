/**
 * ============================================================
 * app-context-awareness.js — Context Awareness
 * ============================================================
 * 功能：上下文感知模块，根据当前 Excel 选区、工作表、数据量等动态调整 AI 行为与提示词。
 *
 * 核心类/函数：
 *   - detectContext() — 检测当前上下文
 *   - getSelectedRangeInfo() — 获取选区信息
 *   - enhancePromptWithContext() — 结合上下文增强提示词
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function () {
    'use strict';

    const ChatLog = window.ChatLog || console;

    // ==================== 数据结构定义 ====================

    /**
     * 操作上下文对象
     */
    class OperationContext {
        constructor() {
            this.timestamp = Date.now();
            this.workbook = {
                name: '',
                path: '',
                isActive: true
            };
            this.worksheet = {
                name: '',
                index: 0,
                isActive: true,
                rowCount: 0,
                colCount: 0
            };
            this.selection = {
                range: '',          // "A1:D15"
                startRow: 0,
                startCol: 0,
                endRow: 0,
                endCol: 0,
                rowCount: 0,
                colCount: 0,
                data: null,        // 完整数据（大数据量时为 null）
                headers: [],       // 表头
                dataSample: null,  // 数据预览（前 10 行）
                areaCount: 0,      // 选区数量
                areas: []          // 多选区详情 [{address, sheetName, workbookName, rowCount, colCount}]
            };
            this.activeCell = '';
            this.openWorkbooks = [];
            this.availableTools = [];
            this.metadata = {
                dataSource: '',    // 'vsto' | 'local-bridge' | 'mock'
                latency: 0,
                error: null
            };
        }
    }

    // ==================== 上下文感知层 ====================

    class ContextAwarenessLayer {
        constructor() {
            this.cache = null;
            this.cacheTime = 0;
            this.TTL = 500; // 缓存 500ms
        }

        /**
         * 获取完整操作上下文
         * @param {TimelineRenderer} timeline - 时间线渲染器
         * @returns {Promise<OperationContext>}
         */
        async getOperationContext(timeline) {
            // 【静默执行】上下文获取不渲染UI，只提供数据给AI
            // 渲染只在AI思考层和工具执行层进行

            try {
                // 检查缓存
                if (this.cache && (Date.now() - this.cacheTime) < this.TTL) {
                    ChatLog.info('Context', '使用缓存的上下文数据');
                    return this.cache;
                }

                const startTime = Date.now();
                const context = new OperationContext();
                const bridge = App.excelBridge;

                if (!bridge) {
                    throw new Error('ExcelBridge 不可用');
                }

                // 并行获取各维度数据
                const results = await this._fetchAllContext(bridge);
                
                // 填充 workbook 信息
                if (results.workbookName) {
                    context.workbook.name = results.workbookName;
                }
                
                // 填充 worksheet 信息
                if (results.sheetName) {
                    context.worksheet.name = results.sheetName;
                }

                // 填充 selection 信息（支持多选区）
                let rangeStr = '';
                if (results.multiSelection && results.multiSelection.areas && results.multiSelection.areas.length > 0) {
                    // 多选区模式
                    context.selection.areaCount = results.multiSelection.areaCount || results.multiSelection.areas.length;
                    context.selection.areas = results.multiSelection.areas;
                    
                    if (context.selection.areaCount === 1) {
                        // 单选区：保持向后兼容
                        rangeStr = results.multiSelection.areas[0].address;
                        context.selection.sheetName = results.multiSelection.areas[0].sheetName;
                        context.selection.workbookName = results.multiSelection.areas[0].workbookName;
                    } else {
                        // 多选区：拼接显示
                        rangeStr = results.multiSelection.areas.map(a => {
                            const wbPrefix = a.workbookName && a.workbookName !== context.workbook.name ? `[${a.workbookName}]` : '';
                            return `${wbPrefix}${a.sheetName}!${a.address}`;
                        }).join(', ');
                    }
                } else if (results.selection) {
                    // 旧逻辑回退
                    rangeStr = results.selection;
                    if (typeof results.selection === 'object') {
                        rangeStr = results.selection.address || '';
                    } else if (typeof results.selection === 'string') {
                        try {
                            const parsed = JSON.parse(results.selection);
                            if (parsed.address) {
                                rangeStr = parsed.address;
                                context.selection.sheetName = parsed.sheetName;
                                context.selection.workbookName = parsed.workbookName;
                            }
                        } catch (e) { ChatLog.debug?.('UI', '[context-awareness] selection 非 JSON，使用原始字符串:', e); }
                    }
                }

                context.selection.range = rangeStr;
                this._parseRange(context.selection, rangeStr);

                // 计算行数和列数（处理多选区时取第一个选区作为默认，或汇总？这里保持默认逻辑）
                if (context.selection.range) {
                    context.selection.rowCount = context.selection.endRow - context.selection.startRow + 1;
                    context.selection.colCount = context.selection.endCol - context.selection.startCol + 1;
                }
                
                // 多选区时计算总行列数
                if (context.selection.areaCount > 1 && context.selection.areas.length > 0) {
                    let totalRows = 0;
                    let maxCols = 0;
                    context.selection.areas.forEach(area => {
                        totalRows += area.rowCount || 0;
                        if (area.colCount > maxCols) maxCols = area.colCount;
                    });
                    // 汇总到 selection 顶层，方便快速判断
                    context.selection.totalRows = totalRows;
                    context.selection.maxCols = maxCols;
                }

                // 填充 activeCell
                if (results.activeCell) {
                    context.activeCell = results.activeCell;
                }

                // 填充 openWorkbooks
                if (results.openWorkbooks) {
                    context.openWorkbooks = results.openWorkbooks;
                }

                // 填充 availableTools
                context.availableTools = window._mcpTools || [];

                context.metadata.dataSource = this._detectDataSource();
                context.metadata.latency = Date.now() - startTime;

                // 缓存结果
                this.cache = context;
                this.cacheTime = Date.now();

                ChatLog.info('Context', '上下文获取完成:', this._renderContextSummary(context));

                return context;

            } catch (error) {
                ChatLog.warn('Context', '获取上下文失败:', error.message);

                // 返回空上下文，不阻塞流程
                const fallbackContext = new OperationContext();
                fallbackContext.metadata.error = error.message;
                return fallbackContext;
            }
        }

        /**
         * 并行获取所有上下文数据
         */
        async _fetchAllContext(bridge) {
            const results = {
                workbookName: null,
                sheetName: null,
                selection: null,
                activeCell: null,
                openWorkbooks: []
            };

            // 尝试并行获取
            const promises = [];

            // 工作簿名称
            if (typeof bridge.GetActiveWorkbookName === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetActiveWorkbookName(), 'workbookName')
                        .then(v => results.workbookName = v)
                );
            }

            // 工作表名称
            if (typeof bridge.GetActiveSheetName === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetActiveSheetName(), 'sheetName')
                        .then(v => results.sheetName = v)
                );
            }

                // 选区（仅获取地址，不获取实际数据）
            // 实际数据由 AI 在执行层通过工具调用获取
            if (typeof bridge.GetSelection === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetSelection(), 'selection')
                        .then(v => results.selection = v)
                );
            }

            // 多选区信息（用于跨表比较/合并等场景）
            if (typeof bridge.GetMultiSelection === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetMultiSelection(), 'multiSelection')
                        .then(v => results.multiSelection = v)
                );
            }

            // 打开的工作簿
            if (typeof bridge.GetOpenWorkbooks === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetOpenWorkbooks(), 'openWorkbooks')
                        .then(v => {
                            try {
                                const parsed = JSON.parse(v);
                                results.openWorkbooks = parsed?.workbooks || [];
                            } catch (e) {
                                ChatLog.warn('Context', '解析 openWorkbooks 失败:', e.message);
                            }
                        })
                );
            }

            // 活动单元格
            if (typeof bridge.GetActiveCell === 'function') {
                promises.push(
                    this._safeCall(() => bridge.GetActiveCell(), 'activeCell')
                        .then(v => results.activeCell = v)
                );
            }

            // 等待所有请求完成（带超时）
            await Promise.allSettled(promises);

            return results;
        }

        /**
         * 安全调用（带错误处理和超时）
         * @param {Function} fn - 要调用的函数
         * @param {string} name - 调用名称（用于日志）
         * @param {number} timeout - 超时时间（毫秒），默认 10000（从5秒提升到10秒）
         */
        async _safeCall(fn, name, timeout = 10000) {
            try {
                // 检查bridge是否可用
                if (!App.excelBridge) {
                    ChatLog.warn('Context', `[${name}] ExcelBridge不可用，跳过调用`);
                    return null;
                }

                // 添加超时保护（提升到10秒，避免频繁超时）
                const timeoutPromise = new Promise((resolve) => {
                    setTimeout(() => {
                        ChatLog.warn('Context', `[${name}] 调用超时(${timeout}ms)，Excel可能正在忙碌或响应缓慢`);
                        
                        // 通知用户（仅首次超时时显示）
                        if (typeof App.showToast === 'function' && !this._hasNotifiedTimeout) {
                            App.showToast(`⏱️ Excel响应稍慢，请稍候...`, 'warning', 3000);
                            this._hasNotifiedTimeout = true;
                            
                            // 5秒后重置通知标记，避免重复提示
                            setTimeout(() => { this._hasNotifiedTimeout = false; }, 5000);
                        }
                        
                        resolve(null);
                    }, timeout);
                });

                const result = await Promise.race([fn(), timeoutPromise]);
                return result;
            } catch (e) {
                ChatLog.warn('Context', `[${name}] 调用失败:`, e.message);
                
                // 特殊处理COM异常
                if (e.message.includes('0x800') || e.message.includes('未知名称')) {
                    ChatLog.error('Context', `[${name}] COM/Excel异常，建议重启Excel或插件`);
                    if (typeof App.showToast === 'function') {
                        App.showToast(`❌ Excel连接异常，请检查Excel是否正常运行`, 'error', 5000);
                    }
                }
                
                return null;
            }
        }

        /**
         * 渲染上下文摘要（用于时间线显示）
         */
        _renderContextSummary(ctx) {
            const lines = [];
            if (ctx.workbook.name) {
                lines.push('工作簿: ' + ctx.workbook.name);
            }
            if (ctx.worksheet.name) {
                lines.push('工作表: ' + ctx.worksheet.name);
            }
            if (ctx.selection.range) {
                lines.push('选区: ' + ctx.selection.range + ' (' + ctx.selection.rowCount + '行×' + ctx.selection.colCount + '列)');
            }
            if (ctx.selection.headers.length > 0) {
                const headersPreview = ctx.selection.headers.slice(0, 5).join(', ');
                const suffix = ctx.selection.headers.length > 5 ? '...' : '';
                lines.push('表头: ' + headersPreview + suffix);
            }
            return lines.join('\n');
        }

        /**
         * 解析范围字符串（如 "A1:D15"）
         */
        _parseRange(selection, rangeStr) {
            if (!rangeStr) return;
            // 移除 Excel 绝对引用的 $ 符号
            const cleanRange = rangeStr.replace(/\$/g, '');
            const match = cleanRange.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
            if (match) {
                selection.startCol = this._colToIndex(match[1]);
                selection.startRow = parseInt(match[2]) - 1;
                selection.endCol = this._colToIndex(match[3]);
                selection.endRow = parseInt(match[4]) - 1;
            }
        }

        /**
         * 列名转索引（A -> 0, B -> 1, ...）
         */
        _colToIndex(col) {
            let index = 0;
            for (let i = 0; i < col.length; i++) {
                index = index * 26 + (col.charCodeAt(i) - 64);
            }
            return index - 1;
        }

        /**
         * 解析选区数据 JSON
         */
        _parseSelectionData(jsonStr) {
            try {
                const data = JSON.parse(jsonStr);
                
                // 支持对象格式（如 {"address":"$A$2:$G$22","rows":[...],"headers":[...]}）
                if (data && typeof data === 'object' && !Array.isArray(data)) {
                    const rows = data.rows || data.data || [];
                    const headers = data.headers || (Array.isArray(rows[0]) ? rows[0].map(h => String(h || '').trim()) : []);
                    const rowCount = rows.length;
                    const colCount = Array.isArray(rows[0]) ? rows[0].length : 0;
                    return { rowCount, colCount, headers, data: rows };
                }
                
                // 支持数组格式（直接是二维数组）
                if (Array.isArray(data) && data.length > 0) {
                    const rowCount = data.length;
                    const colCount = Array.isArray(data[0]) ? data[0].length : 1;
                    const headers = Array.isArray(data[0]) 
                        ? data[0].map(h => String(h || '').trim()) 
                        : [String(data[0] || '').trim()];
                    return { rowCount, colCount, headers, data };
                }
            } catch (e) {
                ChatLog.warn('Context', '解析选区数据失败:', e.message);
            }
            return { rowCount: 0, colCount: 0, headers: [], data: [] };
        }

        /**
         * 检测数据来源
         */
        _detectDataSource() {
            if (window.chrome?.webview?.hostObjects?.excelBridge) {
                return 'vsto';
            } else if (App.excelBridge?.baseUrl) {
                return 'local-bridge';
            }
            return 'mock';
        }

        /**
         * 清除缓存
         */
        clearCache() {
            this.cache = null;
            this.cacheTime = 0;
        }
    }

    // ==================== 注册到全局 ====================

    if (typeof App !== 'undefined') {
        App.ContextAwarenessLayer = ContextAwarenessLayer;
        App.OperationContext = OperationContext;
        App.contextAwareness = new ContextAwarenessLayer();
    }

    // 也暴露到 window（兼容非 App 模式）
    window.ContextAwarenessLayer = ContextAwarenessLayer;
    window.OperationContext = OperationContext;

})();