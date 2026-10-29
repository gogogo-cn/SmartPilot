/**
 * ============================================================
 * app-recognition.js — Recognition Result Operations
 * ============================================================
 * 功能：识别结果操作模块，处理 OCR/AI 识别结果的数据解析、格式化、写入 Excel 与后续操作。
 *
 * 核心函数：
 *   - parseRecognitionData() — 解析识别结果数据
 *   - formatRecognitionResultUniversal() — 统一结果渲染框架
 *   - formatRecognitionResult() — 卡片表格 HTML 渲染
 *   - processImageGeneric() — 通用图片处理流程
 *   - insertRecognitionKvToExcel() — 键值对智能插入 Excel
 *
 * 依赖：
 *   - app.js
 *   - app-utils.js
 *   - app-ocr-parser.js
 * ============================================================
 */

(function() {
    'use strict';

    // HTML 转义（防止 XSS，此文件大量使用但原 app.js 闭包内定义无法跨模块访问）
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.appendChild(document.createTextNode(text));
        return div.innerHTML;
    }

    // ==================== 区域1: 识别结果基础操作 ====================
    
    // --- 来自 app.js:L1199-L1217 ---
    /**
     * 复制识别结果到剪贴板
     */
    function copyRecognitionResult(button) {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        let text = '';
        
        table.querySelectorAll('tr').forEach(row => {
            const th = row.querySelector('th');
            const td = row.querySelector('td');
            if (th && td) {
                text += `${th.textContent.trim()}: ${td.textContent.trim()}\n`;
            }
        });
        
        navigator.clipboard.writeText(text).then(() => {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已复制到剪贴板</span>', true) : App.showToast('✅ 已复制到剪贴板');
        }).catch(() => {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败，请手动复制</span>', true) : App.showToast('复制失败，请手动复制');
        });
    }

    // --- 新增：将键值对识别结果插入Excel（字段/内容表格形式）---
    /**
     * 将键值对识别结果智能插入Excel
     * 用于身份证/名片/发票之外的通用识别结果（如火车票、营业执照、文档等）
     * 调用 InsertParsedData，自动匹配Excel现有表头
     */
    async function insertRecognitionKvToExcel(button) {
        try {
            const card = button.closest('.recognition-result-card');
            const table = card.querySelector('.recognition-table');
            const record = {};

            table.querySelectorAll('tr').forEach(row => {
                const th = row.querySelector('th');
                const td = row.querySelector('td');
                if (th && td) {
                    record[th.textContent.trim()] = td.textContent.trim();
                }
            });

            if (Object.keys(record).length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">⚠️ 没有可插入的数据</span>', true) : App.showToast('⚠️ 没有可插入的数据', 'warning');
                return;
            }

            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                button.disabled = true;
                button.style.opacity = '0.6';

                const result = await window.chrome.webview.hostObjects.excelBridge.InsertParsedData(JSON.stringify([record]));
                const response = JSON.parse(result);

                if (response.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + (response.message || '数据已智能插入Excel') + '</span>', true) : App.showToast('✅ ' + (response.message || '数据已智能插入Excel'), 'success');
                } else {
                    throw new Error(response.error || '插入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('RECOGNITION', '智能插入Excel失败:', error);
            window.ErrorHandler.showError(error, '插入Excel');
        } finally {
            button.disabled = false;
            button.style.opacity = '1';
        }
    }

    // --- 来自 app.js:L1222-L1251 ---
    /**
     * 下载识别结果为CSV文件
     */
    function downloadRecognitionResult(button) {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        
        // 使用BOM头确保Excel正确识别UTF-8编码
        let csv = '\uFEFF项目,内容\n';
        table.querySelectorAll('tr').forEach(row => {
            const th = row.querySelector('th');
            const td = row.querySelector('td');
            if (th && td) {
                const header = th.textContent.trim().replace(/"/g, '""');
                const value = td.textContent.trim().replace(/"/g, '""');
                csv += `"${header}","${value}"\n`;
            }
        });
        
        // 使用UTF-8编码
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', `recognition_result_${Date.now()}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已下载CSV文件</span>', true) : App.showToast('✅ 已下载CSV文件');
    }

    // --- 来自 app.js:L1257-L1290 ---
    /**
     * 将识别结果填入Excel（按钮触发 - 从 data 属性读取数据）
     * @param {HTMLElement} button - 点击的按钮元素
     */
    async function fillRecognitionDataToExcel(button) {
        try {
            // 从 data 属性读取 JSON 和类型，避免 HTML onclick 属性中的转义问题
            const dataJson = button.getAttribute('data-json');
            const typeLabel = button.getAttribute('data-type');

            if (!dataJson) {
                throw new Error('找不到识别数据');
            }

            ChatLog.info('RECOGNITION', `📊 手动触发填入Excel: ${typeLabel}`);
            const data = JSON.parse(dataJson);

            // ★ 状态已迁移到顶部固定状态栏，不再显示底部Toast
            // App.showToast(`📊 正在填入${typeLabel}数据...`, 'info');

            // 同步更新固定状态栏（强制执行）
            try {
                if (typeof App !== 'undefined' && App.fixedStatusBar) {
                    if (!App.fixedStatusBar.isInitialized) {
                        App.fixedStatusBar.init();
                    }
                    App.fixedStatusBar.setProcessing(`正在填入${typeLabel}数据...`);
                }
            } catch(e) {
                ChatLog.warn('RECOGNITION', '[FixedStatusBar] 更新失败:', e);
            }

            // 调用后端填入
            const result = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(
                JSON.stringify(data),
                typeLabel
            );
            const response = JSON.parse(result);

            if (response.success) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + (response.message || typeLabel + '数据已成功填入Excel') + '</span>', true) : App.showToast(`✅ ${response.message || `${typeLabel}数据已成功填入Excel`}`, 'success');
            } else {
                throw new Error(response.error || '填入失败');
            }
        } catch (error) {
            // 使用统一错误处理
            window.ErrorHandler.showError(error, '填入Excel');
            ChatLog.error('RECOGNITION', '填入Excel失败:', error);
        }
    }
    
    // --- 来自 app.js:L1295-L1349 ---
    /**
     * 将JSON识别结果智能填入Excel
     */
    async function fillJsonToExcel(button, imageType = '身份证') {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        
        // 从表格中提取数据
        const record = {};
        table.querySelectorAll('tr').forEach(row => {
            const th = row.querySelector('th');
            const td = row.querySelector('td');
            if (th && td) {
                const key = th.textContent.trim();
                const value = td.textContent.trim();
                if (key && value) {
                    record[key] = value;
                }
            }
        });
        
        if (Object.keys(record).length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">⚠️ 没有可填入的数据</span>', true) : App.showToast('⚠️ 没有可填入的数据', 'warning');
            return;
        }
        
        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 显示加载提示
                const originalIcon = button.innerHTML;
                button.disabled = true;
                button.innerHTML = '';
                button.style.opacity = '0.6';
                
                // 使用 FillDataToExcel 方法
                const dataJson = JSON.stringify(record);
                ChatLog.info('RECOGNITION', '📊 智能填入Excel:', dataJson, '类型:', imageType);
                
                const result = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(dataJson, imageType);
                const response = JSON.parse(result);
                
                if (response.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + (response.message || '数据已成功填入Excel') + '</span>', true) : App.showToast(`✅ ${response.message || '数据已成功填入Excel'}`, 'success');
                } else {
                    throw new Error(response.error || '填入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('RECOGNITION', '填入Excel失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 填入失败: ' + error.message + '</span>', true) : App.showToast('❌ 填入失败: ' + error.message, 'error');
        } finally {
            button.disabled = false;
            button.innerHTML = '📊';
            button.style.opacity = '1';
        }
    }

    // --- 来自 app.js:L1354-L1416 ---
    /**
     * 将多维表格填入Excel
     */
    async function fillMultiTableToExcel(button) {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        
        // 从表格中提取所有行数据
        const rows = [];
        const headers = [];
        
        // 获取表头
        table.querySelectorAll('thead th').forEach(th => {
            headers.push(th.textContent.trim());
        });
        
        // 获取数据行
        table.querySelectorAll('tbody tr').forEach(tr => {
            const row = {};
            const cells = tr.querySelectorAll('td');
            cells.forEach((td, index) => {
                if (headers[index]) {
                    row[headers[index]] = td.textContent.trim();
                }
            });
            rows.push(row);
        });
        
        if (rows.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">⚠️ 没有可填入的数据</span>', true) : App.showToast('⚠️ 没有可填入的数据', 'warning');
            return;
        }
        
        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                button.disabled = true;
                button.innerHTML = '';
                button.style.opacity = '0.6';
                
                const payload = {
                    records: rows,
                    mode: 'smart'
                };
                
                ChatLog.info('RECOGNITION', '📊 多维表格填入Excel:', JSON.stringify(payload, null, 2).substring(0, 200) + '...');
                
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertDataSmart(JSON.stringify(payload));
                const response = JSON.parse(result);
                
                if (response.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + (response.message || '数据已成功填入Excel') + '</span>', true) : App.showToast(`✅ ${response.message || '数据已成功填入Excel'}`, 'success');
                } else {
                    throw new Error(response.error || '填入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('RECOGNITION', '填入Excel失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 填入失败: ' + error.message + '</span>', true) : App.showToast('❌ 填入失败: ' + error.message, 'error');
        } finally {
            button.disabled = false;
            button.innerHTML = '📊';
            button.style.opacity = '1';
        }
    }

    // --- 来自 app.js:L1421-L1447 ---
    /**
     * 复制多维表格
     */
    function copyMultiTable(button) {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        let text = '';
        
        // 复制表头
        const headers = [];
        table.querySelectorAll('thead th').forEach(th => {
            headers.push(th.textContent.trim());
        });
        text += headers.join('\t') + '\n';
        
        // 复制数据行
        table.querySelectorAll('tbody tr').forEach(tr => {
            const cells = [];
            tr.querySelectorAll('td').forEach(td => {
                cells.push(td.textContent.trim());
            });
            text += cells.join('\t') + '\n';
        });
        
        navigator.clipboard.writeText(text).then(() => {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已复制到剪贴板</span>', true) : App.showToast('✅ 已复制到剪贴板');
        }).catch(() => {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败，请手动复制</span>', true) : App.showToast('复制失败，请手动复制');
        });
    }

    // --- 来自 app.js:L1452-L1487 ---
    /**
     * 下载多维表格为CSV
     */
    function downloadMultiTable(button) {
        const card = button.closest('.recognition-result-card');
        const table = card.querySelector('.recognition-table');
        
        // 使用BOM头确保Excel正确识别UTF-8编码
        let csv = '\uFEFF';
        
        // 写入表头
        const headers = [];
        table.querySelectorAll('thead th').forEach(th => {
            headers.push('"' + th.textContent.trim().replace(/"/g, '""') + '"');
        });
        csv += headers.join(',') + '\n';
        
        // 写入数据行
        table.querySelectorAll('tbody tr').forEach(tr => {
            const cells = [];
            tr.querySelectorAll('td').forEach(td => {
                cells.push('"' + td.textContent.trim().replace(/"/g, '""') + '"');
            });
            csv += cells.join(',') + '\n';
        });
        
        const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', `table_data_${Date.now()}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已下载CSV文件</span>', true) : App.showToast('✅ 已下载CSV文件');
    }


    // ==================== 区域2: 数据解析与渲染（最大块）====================

    // --- 来自 app.js:L14199-L14450 ---
    /**
     * 解析识别输出数据
     * @param {object} outputs - 后端返回的outputs对象
     * @param {object} fieldConfig - 字段配置（包含dataParser, metaFields等）
     * @returns {object} 解析后的结构化数据
     */
    function parseRecognitionData(outputs, fieldConfig) {
        if (!outputs) return {};
        
        let data = {};
        
        switch(fieldConfig.dataParser) {
            case 'custom':
                // 自定义解析（用于表格数据）
                // 表格数据结构: { headers: [...], rows: [...] }
                let headers = [];
                let rows = [];
                
                // 尝试从多个位置获取数据
                let source = outputs.structuredData || outputs;
                
                // 新增：支持 analysis.table_structure.columns 和 analysis.data_extraction 格式
                if (!source.headers && !source.rows && outputs.analysis) {
                    ChatLog.info('RECOGNITION', ' 检测到 analysis 格式数据，开始解析...');
                    
                    // 尝试从 analysis.table_structure.columns 获取表头
                    if (outputs.analysis.table_structure && outputs.analysis.table_structure.columns) {
                        headers = outputs.analysis.table_structure.columns;
                        ChatLog.info('RECOGNITION', '🔍 从 analysis.table_structure.columns 获取到表头:', headers);
                    }
                    
                    // 尝试从 analysis.data_extraction 获取数据行
                    if (outputs.analysis.data_extraction && Array.isArray(outputs.analysis.data_extraction)) {
                        rows = outputs.analysis.data_extraction.map(item => {
                            // 每个 item 应该是一个对象，转换为数组
                            if (typeof item === 'object' && item !== null) {
                                return headers.map(h => item[h] || '');
                            }
                            return item;
                        });
                        ChatLog.info('RECOGNITION', '🔍 从 analysis.data_extraction 获取到数据行:', rows.length, '行');
                    }
                    
                    // 如果成功解析，使用这些数据
                    if (headers.length > 0 && rows.length > 0) {
                        source = { headers, rows };
                    }
                }
                
                // 处理 headers
                if (source.headers) {
                    if (Array.isArray(source.headers)) {
                        headers = source.headers;
                    } else if (typeof source.headers === 'string') {
                        // 如果是字符串，优先按 | 分割，其次按逗号分割
                        let headerStr = source.headers.trim();
                        if (headerStr.includes('|')) {
                            headers = headerStr.split('|').map(h => h.trim()).filter(h => h);
                        } else if (headerStr.includes(',')) {
                            headers = headerStr.split(',').map(h => h.trim()).filter(h => h);
                        } else {
                            headers = [headerStr];
                        }
                    }
                }
                
                // 处理 rows
                if (source.rows) {
                    if (Array.isArray(source.rows)) {
                        rows = source.rows.map(row => {
                            if (typeof row === 'string') {
                                // 如果行是字符串，优先按 | 分割，其次按逗号分割
                                let rowStr = row.trim();
                                if (rowStr.includes('|')) {
                                    return rowStr.split('|').map(c => c.trim());
                                } else if (rowStr.includes(',')) {
                                    return rowStr.split(',').map(c => c.trim());
                                } else {
                                    return [rowStr];
                                }
                            }
                            return row;
                        });
                    } else if (typeof source.rows === 'string') {
                        // 如果是字符串，尝试按换行分割行
                        const rowStrings = source.rows.split('\n').map(r => r.trim()).filter(r => r);
                        rows = rowStrings.map(row => {
                            // 每行优先按 | 分割，其次按逗号分割
                            if (row.includes('|')) {
                                return row.split('|').map(c => c.trim());
                            } else if (row.includes(',')) {
                                return row.split(',').map(c => c.trim());
                            } else {
                                return [row];
                            }
                        });
                    }
                }
                
                // 如果解析到有效数据
                if (headers.length > 0 && rows.length > 0) {
                    ChatLog.info('RECOGNITION', '🔍 开始检查重复表头...');
                    ChatLog.info('RECOGNITION', '🔍 headers:', headers);
                    ChatLog.info('RECOGNITION', '🔍 rows[0]:', rows[0]);
                    
                    // 检查第一行是否与表头重复，如果是则跳过第一行
                    let finalRows = rows;
                    
                    // 更全面的重复检测
                    let isDuplicate = false;
                    if (rows.length > 0) {
                        const firstRow = rows[0];
                        
                        if (Array.isArray(firstRow) && firstRow.length === headers.length) {
                            // 数组形式的行
                            isDuplicate = true;
                            for (let i = 0; i < headers.length; i++) {
                                if (String(firstRow[i]).trim() !== String(headers[i]).trim()) {
                                    isDuplicate = false;
                                    ChatLog.info('RECOGNITION', '❌ 不匹配索引', i, ':', String(firstRow[i]).trim(), 'vs', String(headers[i]).trim());
                                    break;
                                }
                            }
                        } else if (typeof firstRow === 'object' && firstRow !== null) {
                            // 对象形式的行
                            isDuplicate = true;
                            for (let i = 0; i < headers.length; i++) {
                                const headerKey = String(headers[i]);
                                const cellValue = firstRow[headerKey] !== undefined ? firstRow[headerKey] : firstRow[i];
                                if (String(cellValue).trim() !== String(headers[i]).trim()) {
                                    isDuplicate = false;
                                    ChatLog.info('RECOGNITION', '❌ 对象不匹配索引', i, ':', String(cellValue).trim(), 'vs', String(headers[i]).trim());
                                    break;
                                }
                            }
                        }
                    }
                    
                    if (isDuplicate) {
                        ChatLog.info('RECOGNITION', '⚠️ 检测到第一行与表头重复，跳过第一行');
                        finalRows = rows.slice(1);
                        ChatLog.info('RECOGNITION', '⚠️ 去重后行数:', finalRows.length);
                    } else {
                        ChatLog.info('RECOGNITION', '✅ 第一行不是表头，保留');
                    }
                    
                    data = {
                        headers: headers,
                        rows: finalRows,
                        rowCount: finalRows.length,
                        colCount: headers.length,
                        description: outputs.description
                    };
                    ChatLog.info('RECOGNITION', '✅ 成功解析表格数据:', finalRows.length, '行', headers.length, '列');
                    ChatLog.info('RECOGNITION', 'headers:', headers);
                    ChatLog.info('RECOGNITION', 'rows:', finalRows);
                } else if (outputs.tableData) {
                    // 兼容其他表格数据格式
                    data = outputs.tableData;
                } else {
                    // 如果没有表格结构，返回原始数据
                    data = { ...outputs };
                }
                break;
                
            case 'extractedText':
                // 优先从 structuredData 获取（后端直接返回的结构化数据）
                if (outputs.structuredData && typeof outputs.structuredData === 'object' &&
                    Object.keys(outputs.structuredData).length > 0) {
                    data = { ...outputs.structuredData };
                    ChatLog.info('IMAGE', `📦 从structuredData获取: ${Object.keys(data).length}个字段`);
                } else if (outputs.extractedText) {
                    try {
                        const parsed = JSON.parse(outputs.extractedText);
                        if (typeof parsed === 'object' && !Array.isArray(parsed)) {
                            data = parsed;
                            ChatLog.info('IMAGE', `📦 从extractedText解析JSON: ${Object.keys(data).length}个字段`);
                        } else {
                            data = { rawText: outputs.extractedText };
                            ChatLog.info('IMAGE', `📄 使用原始文本(非对象)`);
                        }
                    } catch (e) {
                        data = { rawText: outputs.extractedText };
                        ChatLog.info('IMAGE', `📄 使用原始文本(JSON解析失败)`);
                    }
                }
                
                // 如果没有从上述方式解析到数据，使用原始字段
                if (Object.keys(data).length === 0 || (data.rawText && Object.keys(data).length === 1)) {
                    Object.entries(outputs).forEach(([key, value]) => {
                        if (!fieldConfig.metaFields.includes(key) &&
                            value !== undefined && value !== null) {
                            // 支持字符串和对象类型的值
                            if (typeof value === 'string' || typeof value === 'object') {
                                data[key] = value;
                            }
                        }
                    });
                    if (data.rawText) delete data.rawText;
                }

                // 显示解析后的数据内容
                ChatLog.info('IMAGE', `📋 解析后数据: ${JSON.stringify(data).substring(0, 200)}...`);
                break;
                
            case 'direct':
            default:
                // 直接使用outputs的所有字段（后端已经把结构化数据直接展开了）
                // 同时支持 structuredData 格式
                if (outputs.structuredData && typeof outputs.structuredData === 'object') {
                    data = { ...outputs.structuredData };
                } else {
                    data = { ...outputs };
                }
                
                // 如果 structuredData 只有 description 字段，尝试从 extractedText 中提取更多信息
                const hasOnlyDescription = Object.keys(data).length === 1 && data.description;
                
                if (outputs.extractedText && typeof outputs.extractedText === 'string') {
                    // 先尝试解析JSON
                    let parsedJson = null;
                    try {
                        parsedJson = JSON.parse(outputs.extractedText);
                    } catch (e) {
                        // 不是有效的JSON，继续处理文本
                    }
                    
                    if (parsedJson && typeof parsedJson === 'object' && !Array.isArray(parsedJson) && Object.keys(parsedJson).length > 0) {
                        // 合并解析后的JSON数据
                        data = { ...data, ...parsedJson };
                        ChatLog.info('RECOGNITION', '✅ 从 extractedText 解析JSON数据:', Object.keys(parsedJson));
                    } else if (hasOnlyDescription || Object.keys(data).length === 0) {
                        // 如果没有结构化数据或只有description，将extractedText作为原始文本
                        data.rawText = outputs.extractedText;
                        ChatLog.info('RECOGNITION', '✅ 使用 extractedText 作为原始文本');
                    }
                }
                break;
        }
        
        // 过滤元数据字段，但保留 rawText 和 extractedText（LocalOCR 模式需要）
        if (fieldConfig.metaFields && fieldConfig.metaFields.length > 0) {
            fieldConfig.metaFields.forEach(key => {
                // 保留 rawText 和 extractedText，删除其他 metaFields
                if (key !== 'rawText' && key !== 'extractedText' && key !== 'ocrEngine') {
                    delete data[key];
                }
            });
        }
        
        // 过滤原始文本字段（raw/raw_text），但保留 rawText 和 extractedText
        delete data['raw'];
        delete data['raw_text'];

        return data;
    }

    // --- 来自 app.js:L14458-L14485 ---
    /**
     * 收集要显示的字段（标准字段优先）
     * @param {object} data - 数据对象
     * @param {string[]} standardFields - 标准字段顺序
     * @returns {string[]} 要显示的字段名数组
     */
    function collectDisplayFields(data, standardFields) {
        ChatLog.info('RECOGNITION', 'collectDisplayFields: data=', data, 'standardFields=', standardFields);
        const displayFields = [];
        const addedFields = new Set();
        
        // 先添加标准字段（按定义顺序）
        if (standardFields && standardFields.length > 0) {
            for (const field of standardFields) {
                if (data[field] && String(data[field]).trim()) {
                    displayFields.push(field);
                    addedFields.add(field);
                }
            }
        }
        
        // 再添加其他非空字段
        Object.keys(data).forEach(key => {
            if (!addedFields.has(key)) {
                const value = data[key];
                if (value !== null && value !== undefined && String(value).trim()) {
                    displayFields.push(key);
                    addedFields.add(key);
                }
            }
        });
        
        return displayFields;
    }

    // --- 来自 app.js:L14494-L14528 ---
    /**
     * 根据布局构建Excel行数据
     * @param {object} data - 数据对象
     * @param {string} layout - 布局方式: 'horizontal' | 'vertical' | 'grid'
     * @param {string[]} [fields] - 指定字段（可选，不传则自动收集）
     * @returns {Array<Array<string>>} 二维数组，每行是一个数组
     */
    function buildExcelRows(data, layout, fields = null) {
        const displayFields = fields || collectDisplayFields(data, []);
        
        if (displayFields.length === 0) return [];
        
        switch(layout) {
            case 'horizontal':
                // 横向布局：第一行表头，第二行数据（用于发票、身份证）
                return [
                    displayFields,
                    displayFields.map(f => String(data[f] || ''))
                ];
                
            case 'vertical':
                // 纵向布局：每行两列（字段名|值）
                return [
                    ['字段', '值'],
                    ...displayFields.map(f => [f, String(data[f] || '')])
                ];
                
            case 'grid':
                // 网格布局：表格数据（多行多列）
                if (data.rows && Array.isArray(data.rows)) {
                    return data.rows;
                }
                // 如果没有rows属性，回退到横向
                return [
                    displayFields,
                    displayFields.map(f => String(data[f] || ''))
                ];
                
            default:
                return [];
        }
    }

    // --- 来自 app.js:L14536-L14563 ---
    /**
     * 渲染数据为HTML表格
     * @param {object} data - 数据对象
     * @param {string[]} [standardFields] - 标准字段顺序
     * @returns {string} HTML字符串
     */
    function renderDataTable(data, standardFields) {
        const displayFields = collectDisplayFields(data, standardFields);
        
        if (displayFields.length === 0) {
            return '<div style="color:#999; font-size:13px; padding:8px 0;">无有效数据</div>';
        }
        
        let html = '<table style="width:100%; border-collapse:collapse; font-size:13px;">';
        
        // 表头行
        html += '<tr>';
        displayFields.forEach(field => {
            html += `<th style="border:1px solid #e0e0e0; padding:6px 10px; background:#f5f5f5; color:#666; text-align:left; font-weight:500;">${escapeHtml(field)}</th>`;
        });
        html += '</tr>';
        
        // 数据行
        html += '<tr>';
        displayFields.forEach(field => {
            const value = String(data[field] || '');
            html += `<td style="border:1px solid #e0e0e0; padding:6px 10px; color:#333;">${escapeHtml(value)}</td>`;
        });
        html += '</tr>';
        
        html += '</table>';
        
        return html;
    }

    // --- 来自 app.js:L14572-L14615 ---
    /**
     * 渲染操作按钮
     * @param {string} resultId - 结果ID
     * @param {string[]} buttonTypes - 按钮类型数组
     * @param {string} typeName - 类型名称（用于提示文字）
     * @returns {string} HTML字符串
     */
    function renderActionButtons(resultId, buttonTypes, typeName) {
        const isExcelHost = !(window.SmartPilotHost) || !!window.SmartPilotHost.isExcel;
        const excelOnly = { insertVertical: 1, fillExcel: 1, insertToExcel: 1 };
        const buttonConfigs = {
            'copy': {
                label: '📋 复制',
                onclick: `copyRecognitionData('${resultId}')`,
                color: '#2196f3',
                hoverColor: '#1976d2'
            },
            'insertVertical': {
                label: '📊 插入 Excel',
                onclick: `insertVerticalToExcel('${resultId}')`,
                color: '#4caf50'
            },
            'fillExcel': {
                label: '📊 填入 Excel',
                onclick: `fillRecognitionDataToExcelByResultId('${resultId}')`,  // 使用旧函数名
                color: '#4caf50',
                hoverColor: '#388e3c'
            },
            'insertToExcel': {
                label: '插入到 Excel',
                onclick: `insertRecognitionTableToExcel('${resultId}')`,
                color: '#4caf50'
            },
            'saveJson': {
                label: '保存',
                onclick: `saveRecognitionDataAsJson('${resultId}')`,
                color: '#ff9800'
            }
        };
        
        let html = `<div style="display:flex; gap:8px; margin-top:12px; padding-top:10px; border-top:1px solid #eee;">`;
        
        buttonTypes.forEach(btnType => {
            if (!isExcelHost && excelOnly[btnType]) return; // Word/PPT 不展示 Excel 填表按钮
            const btn = buttonConfigs[btnType];
            if (btn) {
                const hoverColor = btn.hoverColor || btn.color;
                html += `<button onclick="${btn.onclick}" style="padding:7px 16px; background:${btn.color}; color:white; border:none; border-radius:5px; cursor:pointer; font-size:12px; transition:all 0.15s;" onmouseover="this.style.background='${hoverColor}'" onmouseout="this.style.background='${btn.color}'">${btn.label}</button>`;
            }
        });
        
        html += '</div>';
        return html;
    }

    // --- 来自 app.js:L14624-L14632 ---
    /**
     * 存储识别结果到全局变量
     * @param {string} resultId - 结果ID
     * @param {object} data - 解析后的数据
     * @param {string} type - 识别类型
     * @param {object} rawOutputs - 原始outputs数据
     */
    function storeRecognitionResult(resultId, data, type, rawOutputs) {
        window.recognitionResults = window.recognitionResults || {};
        window.recognitionResults[resultId] = {
            data: data,
            type: type,
            rawOutputs: rawOutputs,
            timestamp: Date.now()
        };
    }

    /**
     * 为 SmartTieredRenderer 渲染的 HTML 包装插入 Excel 按钮
     */
    function wrapSmartTieredHtmlWithInsertButton(html, type, data) {
        var dataJson = JSON.stringify(data);
        var actionsHtml = '<div class="recognition-actions" style="margin-top:8px;padding-top:8px;border-top:1px solid #e0e0e0;display:flex;gap:8px;align-items:center;">' +
            '<button class="recognition-action-btn" onclick="App.recognition.insertSmartTieredResultToExcel(this)" data-type="' + escapeHtml(type) + '" data-json="' + escapeHtml(dataJson) + '" title="插入到Excel" style="background:#4caf50;color:white;font-size:13px;">📊 插入到Excel</button>' +
            '</div>';
        return '<div class="recognition-result-container smart-tiered-wrapper">' + html + actionsHtml + '</div>';
    }

    /**
     * 将 SmartTieredRenderer 渲染的识别结果智能插入 Excel
     * 调用 InsertParsedData，自动匹配Excel现有表头
     */
    async function insertSmartTieredResultToExcel(button) {
        try {
            var type = button.dataset.type || 'unknown';
            var data = JSON.parse(button.dataset.json || '{}');
            var records = [];

            if (data.headers && Array.isArray(data.headers) && data.rows && Array.isArray(data.rows)) {
                // 标准表格数据
                data.rows.forEach(function(row) {
                    var record = {};
                    data.headers.forEach(function(header, idx) {
                        var value = Array.isArray(row) ? row[idx] : row[header];
                        record[header] = value !== undefined && value !== null ? String(value) : '';
                    });
                    records.push(record);
                });
            } else {
                // 键值对数据
                var skipKeys = ['imageType','image_type','description','structuredData','suggestedAction','suggested_action','extractedText','extracted_text','type','success','error','executionLog','totalExecutionTime','rawText','text','content','ocrEngine','imageData','fileName','timestamp','model','version','recognitionType'];
                var record = {};
                Object.entries(data).forEach(function(entry) {
                    var key = entry[0], value = entry[1];
                    if (skipKeys.indexOf(key) !== -1) return;
                    if (value === undefined || value === null || value === '') return;
                    record[key] = String(value);
                });
                records.push(record);
            }

            if (records.length === 0 || (records.length === 1 && Object.keys(records[0]).length === 0)) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">⚠️ 没有可插入的数据</span>', true) : App.showToast('⚠️ 没有可插入的数据', 'warning');
                return;
            }

            button.disabled = true;
            button.style.opacity = '0.6';

            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                var result = await window.chrome.webview.hostObjects.excelBridge.InsertParsedData(JSON.stringify(records));
                var response = JSON.parse(result);
                if (response.success) {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + (response.message || '数据已智能插入Excel') + '</span>', true) : App.showToast('✅ ' + (response.message || '数据已智能插入Excel'), 'success');
                } else {
                    throw new Error(response.error || '插入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('RECOGNITION', 'SmartTieredRenderer 智能插入Excel失败:', error);
            window.ErrorHandler.showError(error, '插入Excel');
        } finally {
            button.disabled = false;
            button.style.opacity = '1';
        }
    }

    // --- 来自 app.js:L14641-L14983 ---
    /**
     * 统一的识别结果处理主函数（简化版）
     * @param {string} type - 识别类型（invoice, card, table, general）
     * @param {object} response - 后端完整响应
     * @param {object} extraInputs - 额外参数（如 fillMode）
     * @returns {Promise<string>} HTML字符串
     */
    async function formatRecognitionResultUniversal(type, response, extraInputs = {}) {
                // ============================================================
        // 【SmartTieredRenderer 集成】智能分层渲染器 v1.0.0
        // ============================================================
        if (window.SmartTieredRenderer || (window.App && window.App.SmartTieredRenderer)) {
            try {
                ChatLog.info('RECOGNITION', '✅ [Universal] 使用 SmartTieredRenderer 渲染');

                const outputs = response.outputs || {};
                const rawText = outputs.rawText || outputs.description || outputs.text || 
                               response.rawText || response.description || '';
                
                const isTableData = outputs.headers && outputs.rows &&
                                   Array.isArray(outputs.headers) && Array.isArray(outputs.rows) &&
                                   outputs.headers.length > 0;

                let thinkingContent = '';
                if (rawText && typeof rawText === 'string') {
                    const lines = rawText.split('\n').filter(function(line) { return line.trim(); });
                    thinkingContent = lines.slice(0, Math.min(lines.length, 10)).join('\n');
                    if (lines.length > 10) {
                        thinkingContent += '\n\n...（完整思考过程见下方分析总结）';
                    }
                }

                const responseData = {
                    thinking: thinkingContent || undefined,
                    
                    toolCalls: [{
                        name: type === 'table' ? 'table_recognition' : 
                              type === 'invoice' ? 'invoice_ocr' :
                              type === 'idcard' ? 'idcard_recognition' :
                              type === 'card' ? 'card_recognition' : 'image_analysis',
                        parameters: { 
                            imageType: type, 
                            fillMode: extraInputs.fillMode,
                            timestamp: new Date().toISOString() 
                        },
                        result: { 
                            success: true,
                            dataType: isTableData ? 'table' : Object.keys(outputs).length > 3 ? 'structured' : 'text',
                            fieldsCount: isTableData ? outputs.headers.length : Object.keys(outputs).length,
                            recordsCount: isTableData ? outputs.rows.length : 1,
                            modelUsed: currentChatModel || 'unknown'
                        },
                        status: 'success',
                        duration: null
                    }],
                    
                    result: isTableData ? {
                        headers: outputs.headers,
                        rows: outputs.rows,
                        title: (type === 'table' ? '表格数据提取' : 
                               type === 'invoice' ? '发票信息识别' :
                               type === 'idcard' ? '身份证信息识别' :
                               type === 'card' ? '名片信息识别' : '图片文本提取结果')
                    } : outputs,
                    
                    analysis: rawText || undefined,
                    
                    metadata: { 
                        model: currentChatModel || 'AI Model',
                        version: '1.0.0', 
                        recognitionType: type,
                        timestamp: new Date().toISOString()
                    }
                };

                const renderer = window.App && window.App.SmartTieredRenderer ? window.App.SmartTieredRenderer : window.SmartTieredRenderer;
                const html = await renderer.render(responseData);
                const wrappedHtml = wrapSmartTieredHtmlWithInsertButton(html, type, responseData.result);
                
                ChatLog.info('RECOGNITION', '🎨 [Universal] SmartTieredRenderer 渲染完成');
                return wrappedHtml;

            } catch (error) {
                ChatLog.error('RECOGNITION', '❌ [Universal] SmartTieredRenderer 渲染失败，回退到传统模式:', error);
            }
        } else {
            ChatLog.warn('RECOGNITION', '⚠️ [Universal] SmartTieredRenderer 未加载，使用传统渲染模式');
        }
        let config = IMAGE_PROCESSING_CONFIG[type];
        if (!config) {
            return `<div style="color:#f44336; padding:12px; background:#ffebee; border-radius:6px;">❌ 未知的识别类型: ${type}</div>`;
        }
        
        if (typeof config === 'string') {
            config = IMAGE_PROCESSING_CONFIG[config];
            if (!config) {
                return `<div style="color:#f44336; padding:12px; background:#ffebee; border-radius:6px;">❌ 配置解析失败: ${type}</div>`;
            }
        }
        
        const { fieldConfig, displayConfig } = config;
        const outputs = response.outputs || {};
        const fillMode = extraInputs.fillMode;
        
        // 获取当前AI提供商和模型信息
        const _rawModel = currentChatModel || config.model || App.config.model || '';
        const _providerConfig = (typeof App !== 'undefined' && App.PROVIDER_CONFIGS) ? App.PROVIDER_CONFIGS[App.config.apiType] : null;
        const _actualModel = (_rawModel === '默认模型' || _rawModel === '当前模型' || _rawModel === 'AUTO' || !_rawModel)
            ? (_providerConfig ? (_providerConfig.model || '') : '')
            : _rawModel;
        const currentModel = _actualModel || '默认模型';

        // ========== 0. 提取结构化数据（AI处理：解析表头和字段）==========
        // 注意：extract 步骤已经在 handleImageRecognitionResponse 中标记为 running
        // 这里只是更新完成状态

        // ========== 1. 数据解析 ==========
        const parsedData = parseRecognitionData(outputs, fieldConfig);

        // 【已修改】检查是否为通用文本分析结果（description/text字段包含自由文本）
        const rawText = outputs.description || outputs.text || outputs.content || '';
        // 通用文本判断：type为general，且有rawText，且parsedData为空或只有元数据字段
        const metaFields = ['description', 'text', 'content', 'rawText', 'extractedText', 'ocrEngine', 'imageData', 'fileName', 'imageType', 'image_type', 'suggestedAction', 'suggested_action'];
        const parsedDataKeys = Object.keys(parsedData).filter(k => !metaFields.includes(k));
        // 放宽判断：只要有rawText且type为general，就认为是通用文本（优先使用通用文本渲染）
        const isGeneralText = type === 'general' && rawText;

        // 更新提取结构化数据步骤完成状态（显示解析到的表头和字段）
        if (currentTracker) {
            if (isGeneralText) {
                // 通用文本分析：直接标记完成
                currentTracker.updateStep('extract', 'complete', '分析完成');
            } else {
                // 检查是否有表格数据（有 headers 和 rows）
                const hasTableData = parsedData.headers && parsedData.rows &&
                                     Array.isArray(parsedData.headers) && Array.isArray(parsedData.rows);
                // 检查是否有结构化字段数据
                const hasStructuredData = parsedData && Object.keys(parsedData).length > 0 && !hasTableData;

                if (hasTableData) {
                    // 表格数据：显示识别结果
                    const headers = parsedData.headers;
                    const headerDisplay = headers.slice(0, 3).join('，');
                    const moreHeaders = headers.length > 3 ? `等${headers.length}列` : '';
                    const rowCount = parsedData.rows.length;
                    currentTracker.updateStep('extract', 'complete', `识别完成: ${headerDisplay}${moreHeaders}，共${rowCount}行数据`);
                } else if (hasStructuredData) {
                    // 结构化数据：显示提取的字段
                    const fields = Object.keys(parsedData);
                    const fieldDisplay = fields.slice(0, 3).join('，');
                    const moreFields = fields.length > 3 ? `等${fields.length}个字段` : '';
                    currentTracker.updateStep('extract', 'complete', `提取完成: ${fieldDisplay}${moreFields}`);
                } else {
                    currentTracker.updateStep('extract', 'complete', '未识别到结构化数据');
                }
            }
        }
        
        // ========== 2. Excel填入（统一策略）==========
        let fillStatus = '';
        const wantsToFill = fieldConfig.excelFill.enabled && 
                          fillMode && 
                          fillMode !== 'recognize_only';
        
        if (wantsToFill) {
            if (config.backendAutoFill) {
                const hasExecLog = response.executionLog && response.executionLog.length > 0;
                const hasActualData = parsedData && Object.keys(parsedData).length > 0;
                const hasStructuredData = outputs.structuredData && typeof outputs.structuredData === 'object' && Object.keys(outputs.structuredData).length > 0;
                fillStatus = (hasExecLog && (hasActualData || hasStructuredData)) ? "数据已写入到工作表\"***\"" : '';
            } else {
                try {
                    if (parsedData && Object.keys(parsedData).length > 0) {
                        const payload = {
                            records: [parsedData],
                            mode: 'smart'
                        };
                        await window.chrome.webview.hostObjects.excelBridge.InsertDataSmart(JSON.stringify(payload));
                        fillStatus = "数据已写入到工作表\"***\"";
                    }
                } catch (error) {
                    ChatLog.error('RECOGNITION', '❌ 填入失败:', error);
                    fillStatus = `（填入失败: ${error.message}）`;
                }
            }
        }
        
        // ========== 3. 生成HTML ==========
        const resultId = `${type}_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;

        // 获取类型标签
        const typeLabel = config.name.replace('识别', '');  // "发票"、"身份证/名片" 等
        
        // 存储识别结果到全局变量（用于后续操作）
        storeRecognitionResult(resultId, parsedData, type, response.outputs);
        
        // 获取字段映射
        const fieldMapping = fieldConfig.fieldMapping || {};
        const getFieldName = (key) => fieldMapping[key] || key;
        
        // 检测是否为表格数据（有 headers 和 rows）
        const isTableData = parsedData.headers && parsedData.rows && 
                           Array.isArray(parsedData.headers) && Array.isArray(parsedData.rows);
        
        let html = '';

        // 【稳定性修复】✨ 优先级调整：表格数据 > 通用文本 > 键值对
        // 原因：AI 返回格式不稳定，有时同时包含 rawText 和结构化表格数据
        // 解决方案：强制优先使用结构化表格数据，确保输出100%一致性和可操作性
        
        // 调试日志：帮助定位渲染模式选择原因（生产环境可移除）
        ChatLog.info('RECOGNITION', '🎯 渲染模式判断:', {
            type: type,
            isTableData: isTableData,
            isGeneralText: isGeneralText,
            hasRawText: !!rawText,
            rawTextLength: rawText?.length || 0,
            parsedDataKeys: Object.keys(parsedData),
            hasHeaders: !!(parsedData.headers && parsedData.headers.length),
            hasRows: !!(parsedData.rows && parsedData.rows.length),
            headersCount: parsedData.headers?.length || 0,
            rowsCount: parsedData.rows?.length || 0
        });
        
        if (isTableData) {
            ChatLog.info('RECOGNITION', '✅ 选择模式: 表格渲染 (优先级1)');
            // ========== 优先级1：有表格数据 → 强制使用表格渲染（最稳定）==========
            // ========== 表格数据显示（卡片样式，参考身份证/名片设计）==========
            
            // 剔除被模型当作首行数据重复输出的表头
            let finalDataRows = parsedData.rows;
            if (parsedData.rows.length > 0 && parsedData.headers.length > 0) {
                const firstRow = parsedData.rows[0];
                if (Array.isArray(firstRow) && firstRow.length === parsedData.headers.length &&
                    parsedData.headers.every((h, i) => String(firstRow[i]).trim() === String(h).trim())) {
                    finalDataRows = parsedData.rows.slice(1);
                }
            }

            const tableRows = [parsedData.headers, ...finalDataRows];
            const rowsJson = JSON.stringify(tableRows).replace(/"/g, '&quot;');
            
            // 使用卡片样式包装表格
            html += `<div class="recognition-result-container">`;
            html += `<h3>表格数据提取</h3>`;
            html += `<div class="recognition-result-card">`;
            
            // 卡片头部：图标 + 类型 + 操作按钮
            html += `<div class="recognition-header">`;
            html += `<div class="recognition-title">`;
            html += `<span class="recognition-title-icon">📊</span>`;
            html += `<span>表格</span>`;
            html += `</div>`;
            html += `<div class="recognition-actions">`;
            html += `<button class="recognition-action-btn" onclick="copyTableToClipboard(${rowsJson})" title="复制表格">📋</button>`;
            html += `<button class="recognition-action-btn" onclick="insertTableToExcel(${rowsJson})" title="插入到Excel" style="background: #4caf50; color: white;">📊</button>`;
            html += `<button class="recognition-action-btn" onclick="downloadTableAsCsv(${rowsJson})" title="下载CSV">⬇️</button>`;
            html += `</div>`;
            html += `</div>`;
            
            // 统计信息
            if (parsedData.rowCount && parsedData.colCount) {
                html += `<div style="padding: 8px 16px; background: #f5f5f5; border-bottom: 1px solid #e8e8e8; color: #666; font-size: 13px;">`;
                html += `识别到表格：${parsedData.colCount} 列 × ${parsedData.rowCount} 行`;
                html += `</div>`;
            }
            
            // 表格容器
            html += `<div style="overflow:auto; max-width:100%; max-height:400px; -webkit-overflow-scrolling:touch; box-sizing:border-box; scroll-behavior:smooth; overscroll-behavior:contain;">`;
            
            html += `<table class="recognition-table" style="border-collapse:collapse; font-size:13px; width:auto; min-width:100%;">`;
            
            // 表头（使用CSS类样式）
            html += `<thead><tr>`;
            parsedData.headers.forEach(header => {
                html += `<th>${escapeHtml(String(header))}</th>`;
            });
            html += `</tr></thead>`;
            
            // 表体
            html += `<tbody>`;
            finalDataRows.forEach((row, rowIndex) => {
                html += `<tr>`;
                if (Array.isArray(row)) {
                    row.forEach((cell, cellIndex) => {
                        // 直接渲染HTML,支持<br>等标签
                        const cellContent = String(cell || '').replace(/\n/g, '<br>');
                        html += `<td>${cellContent}</td>`;
                    });
                } else {
                    // 如果行不是数组，尝试作为对象处理
                    parsedData.headers.forEach(header => {
                        const cellValue = row[header] || '';
                        const cellContent = String(cellValue).replace(/\n/g, '<br>');
                        html += `<td>${cellContent}</td>`;
                    });
                }
                html += `</tr>`;
            });
            html += `</tbody>`;
            html += `</table>`;
            html += `</div>`;

            // 卡片底部说明
            html += `<div class="recognition-footer">`;
            html += `<p>说明：该表格数据为系统自动识别提取，仅供参考，请核对原始信息</p>`;
            html += `</div>`;

            html += `</div>`;
            html += `</div>`;

            // 【增强】如果同时存在文本分析（rawText），在表格下方显示分析总结
            // 这样既保证表格的稳定性，又保留AI的分析价值
            if (rawText && rawText.length > 50) {
                ChatLog.info('RECOGNITION', '📊 检测到附加分析文本，将在表格下方显示总结');
                html += `<div class="recognition-result-container" style="margin-top:12px;">`;
                html += `<div class="recognition-result-card" style="background:#fff; border-radius:8px; border:1px solid #e8e8e8; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.04);">`;

                // 分析总结标题
                html += `<div class="recognition-header" style="background:#f8f9fa; padding:10px 14px; border-bottom:1px solid #e8e8e8;">`;
                html += `<span style="font-weight:500; color:#333; font-size:13px;">📝 分析总结</span>`;
                html += `</div>`;

                // Markdown内容渲染
                html += `<div style="padding:14px; line-height:1.6;">`;
                let analysisHtml = '';
                try {
                    let processedAnalysis = rawText
                        .replace(/## /g, '\n## ')
                        .replace(/# /g, '\n# ')
                        .replace(/\*\* /g, '\n** ')
                        .replace(/- /g, '\n- ');

                    if (typeof marked !== 'undefined' && marked.parse) {
                        analysisHtml = marked.parse(processedAnalysis);
                    } else {
                        analysisHtml = processedAnalysis
                            .replace(/&/g, '&amp;')
                            .replace(/</g, '&lt;')
                            .replace(/>/g, '&gt;')
                            .replace(/\n\n/g, '</p><p>')
                            .replace(/\n/g, '<br>');
                        analysisHtml = '<p>' + analysisHtml + '</p>';
                    }
                } catch (e) {
                    ChatLog.error('RECOGNITION', '分析总结渲染失败:', e);
                    analysisHtml = rawText.replace(/\n/g, '<br>');
                }
                html += analysisHtml;
                html += `</div>`;

                html += `</div>`;
                html += `</div>`;
            }

        // ========== 优先级2：通用文本分析 → 渲染Markdown ==========
        } else if (isGeneralText) {
            ChatLog.info('RECOGNITION', '📝 选择模式: Markdown渲染 (优先级2)');
            html += `<div class="recognition-result-container">`;
            html += `<div class="recognition-result-card" style="background:#fff; border-radius:8px; border:1px solid #e8e8e8; overflow:hidden; margin-bottom:12px; box-shadow:0 1px 3px rgba(0,0,0,0.04);">`;

            // 卡片头部
            html += `<div class="recognition-header" style="background:#f8f9fa; padding:10px 14px; border-bottom:1px solid #e8e8e8; display:flex; align-items:center; gap:10px;">`;
            html += `<div class="status-badge" style="display:flex; align-items:center; gap:5px; background:#e8f5e9; color:#2e7d32; padding:4px 10px; border-radius:4px; font-size:12px; font-weight:500;">`;
            html += `<span style="font-size:11px;">✓</span>`;
            html += `<span>分析完成</span>`;
            html += `</div>`;
            html += `</div>`;

            // Markdown内容渲染
            html += `<div style="padding:16px; line-height:1.6;">`;
            let markdownHtml = '';
            try {
                let processedText = rawText
                    .replace(/## /g, '\n## ')
                    .replace(/# /g, '\n# ')
                    .replace(/\*\* /g, '\n** ')
                    .replace(/- /g, '\n- ');

                if (typeof marked !== 'undefined' && marked.parse) {
                    markdownHtml = marked.parse(processedText);
                } else {
                    markdownHtml = processedText
                        .replace(/&/g, '&amp;')
                        .replace(/</g, '&lt;')
                        .replace(/>/g, '&gt;')
                        .replace(/\n\n/g, '</p><p>')
                        .replace(/\n/g, '<br>');
                    markdownHtml = '<p>' + markdownHtml + '</p>';
                }
            } catch (e) {
                ChatLog.error('RECOGNITION', 'Markdown渲染失败:', e);
                markdownHtml = rawText.replace(/\n/g, '<br>');
            }
            html += markdownHtml;
            html += `</div>`;

            html += `</div>`;
            html += `</div>`;

        // ========== 优先级3：有结构化数据但不是表格 → 键值对显示 ==========
        } else if (Object.keys(parsedData).length > 0) {
            ChatLog.info('RECOGNITION', '🔑 选择模式: 键值对渲染 (优先级3)');
            // ========== 普通键值对数据显示（新卡片表格样式）==========
                
                // 转换type为中文名称，用于字段名映射
                let mappingType = type;
                if (type === 'invoice') mappingType = '发票';
                if (type === 'card') mappingType = '身份证';
                if (type === 'table') mappingType = '表格';
                
                // 使用新的卡片表格格式化函数
                html += formatRecognitionResult(parsedData, mappingType);
        } else if (config.backendAutoFill && response.executionLog && response.executionLog.length > 0 && !(config.displayConfig && config.displayConfig.hideDetailedSuccess)) {
            html += `<div class="recognition-result-card" style="background:#fff; border-radius:8px; border:1px solid #e8e8e8; overflow:hidden; margin-bottom:12px; box-shadow:0 1px 3px rgba(0,0,0,0.04);">
                <div class="recognition-header" style="background:#f8f9fa; padding:10px 14px; border-bottom:1px solid #e8e8e8; display:flex; align-items:center; gap:10px;">
                    <div class="status-badge" style="display:flex; align-items:center; gap:5px; background:#e8f5e9; color:#2e7d32; padding:4px 10px; border-radius:4px; font-size:12px; font-weight:500;">
                        <span style="font-size:11px;">✓</span>
                        <span>分析完成</span>
                    </div>
                    <div class="type-info" style="color:#666; font-size:12px;">
                        识别类型：<span style="color:#333; font-weight:500;">${typeLabel}</span>
                    </div>
                </div>
                <div style="padding:16px 14px; text-align:center;">
                    <div style="color:#4caf50; font-size:14px; font-weight:500;">✅ 数据已自动填入 Excel</div>
                    <div style="color:#999; font-size:12px; margin-top:4px;">请检查工作表确认数据已正确填入</div>
                </div>
            </div>`;
        } else {
            html += '<div style="color:#999; font-size:13px;">未识别到有效信息</div>';
        }
        
        // ========== 4. 存储数据 ==========
        // 把extraInputs中的imageData和fileName合并到outputs中
        const extendedOutputs = {
            ...outputs,
            imageData: extraInputs?.imageData,
            fileName: extraInputs?.fileName
        };
        storeRecognitionResult(resultId, parsedData, type, extendedOutputs);
        
        return html;
    }

    // ==================== 统一操作函数 ====================
    
    // --- 来自 app.js:L14990-L15019 ---
    /**
     * 统一复制功能
     */
    async function copyRecognitionData(resultId) {
        try {
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }
            
            const data = result.data;
            let text = '';
            
            if (result.type === 'table' && data.rows) {
                // 表格类型：复制为制表符分隔
                text = data.rows.map(row => row.join('\t')).join('\n');
            } else {
                // 其他类型：复制为键值对
                text = Object.entries(data)
                    .filter(([k,v]) => v)
                    .map(([k,v]) => `${k}: ${v}`)
                    .join('\n');
            }
            
            await navigator.clipboard.writeText(text);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已复制到剪贴板</span>', true) : App.showToast('✅ 已复制到剪贴板');
        } catch (error) {
            ChatLog.error('RECOGNITION', '复制失败:', error);
            // 使用统一错误处理
            window.ErrorHandler.showError(error, '复制');
        }
    }

    // --- 来自 app.js:L15028-L15053 ---
    /**
     * 纵向插入到Excel（卡片格式：每行一个字段）
     * 用于名片的"插入Excel"按钮
     */
    async function insertVerticalToExcel(resultId) {
        try {
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }

            const config = IMAGE_PROCESSING_CONFIG[result.type];
            const verticalConfig = config?.fieldConfig?.verticalFill || { layout: 'vertical' };

            // 构建纵向卡片格式的数据
            const rows = buildExcelRows(result.data, verticalConfig.layout);

            if (rows.length > 0 && rows[0].length > 0) {
                // 纵向模式：从当前选中位置或A1开始
                await insertTableToExcel(rows, { startCell: 'A1' });
                App.showToast('✅ 已纵向插入 Excel（卡片格式）');
            } else {
                App.showToast('⚠️ 没有有效数据可插入');
            }
        } catch (error) {
            ChatLog.error('RECOGNITION', '纵向插入Excel失败:', error);
            App.showToast('❌ 插入失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L15056-L15106 ---
    // 注释：此函数已被fillRecognitionDataToExcel(dataJson, typeLabel)替代
    async function fillRecognitionDataToExcelByResultId(resultId) {
        try {
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }

            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 获取结构化数据和类型
                const structuredData = result.data?.structuredData || result.data;
                const imageType = result.data?.imageType || result.rawOutputs?.imageType || 'general';
                
                // 获取字段映射配置
                let fieldMapping = null;
                const typeConfig = IMAGE_TYPE_CONFIG[imageType] || IMAGE_TYPE_CONFIG['默认'];
                if (typeConfig && typeConfig.fieldMapping && Object.keys(typeConfig.fieldMapping).length > 0) {
                    fieldMapping = typeConfig.fieldMapping;
                    ChatLog.info('RECOGNITION', '🎯 使用字段映射配置:', fieldMapping);
                }
                
                ChatLog.info('RECOGNITION', '🎯 填入Excel - 类型:', imageType);
                ChatLog.info('RECOGNITION', '🎯 填入Excel - 数据:', structuredData);
                
                // 构建请求数据
                const requestData = {
                    data: structuredData,
                    imageType: imageType,
                    fieldMapping: fieldMapping
                };
                
                // 使用 FillDataToExcel 方法，它支持按类型横向填入
                const dataJson = JSON.stringify(requestData);
                const response = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(dataJson, imageType);
                const resultObj = JSON.parse(response);
                
                if (resultObj.success) {
                    ChatLog.info('RECOGNITION', '✅ 填入成功:', resultObj);
                    App.showToast('✅ ' + (resultObj.message || '数据已成功填入Excel'));
                } else {
                    throw new Error(resultObj.error || '填入失败');
                }
            } else {
                App.showToast('❌ Excel 连接不可用');
            }
        } catch (error) {
            // 使用统一错误处理
            window.ErrorHandler.showError(error, '数据填入');
            ChatLog.error('RECOGNITION', '填入Excel失败:', error);
        }
    }

    // --- 来自 app.js:L15116-L15154 ---
    /**
     * 将识别结果填入Excel（自动触发版本）
     * @param {Object} structuredData - 结构化数据
     * @param {string} imageType - 图片类型（身份证/名片/发票）
     */
    async function fillToExcel(structuredData, imageType) {
        try {
            ChatLog.info('RECOGNITION', '========== 自动填入Excel ==========');
            ChatLog.info('RECOGNITION', '数据类型:', imageType);
            ChatLog.info('RECOGNITION', '结构化数据:', structuredData);
            
            if (!structuredData || typeof structuredData !== 'object' || Object.keys(structuredData).length === 0) {
                throw new Error('没有可用的数据');
            }
            
            // 添加AI消息提示正在填入（临时消息，不进入上下文历史）
            addMessage('ai', `📊 正在将${imageType}信息填入Excel...`, false, null, true);
            
            // 调用FillDataToExcel接口
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const dataJson = JSON.stringify(structuredData);
                ChatLog.info('RECOGNITION', '准备发送的数据JSON长度:', dataJson.length);
                
                const fillResult = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(dataJson, imageType);
                ChatLog.info('RECOGNITION', 'FillDataToExcel返回结果:', fillResult);
                
                const response = JSON.parse(fillResult);
                if (response.success) {
                    addMessage('ai', `✅ ${response.message}`);
                    App.showToast('✅ ' + response.message);
                } else {
                    throw new Error(response.error || '填入失败');
                }
            } else {
                throw new Error('无法连接到Excel');
            }
            
        } catch (error) {
            ChatLog.error('RECOGNITION', '========== fillToExcel失败 ==========');
            ChatLog.error('RECOGNITION', error);
            addMessage('ai', `❌ 填入Excel失败: ${error.message}`);
            throw error;
        }
    }

    // --- 来自 app.js:L15159-L15161 ---
    /**
     * 重试填入Excel（用于按钮点击）
     */
    async function retryFillToExcel() {
        App.showToast('🔄 请重新识别图片后手动填入');
    }

    // --- 来自 app.js:L15163-L15217 ---
    /**
     * 处理填入Excel按钮点击（新设计）
     */
    async function handleFillToExcel(resultId) {
        try {
            ChatLog.info('RECOGNITION', '========== 点击了填入Excel按钮 ==========');
            
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }

            ChatLog.info('RECOGNITION', '获取到存储的识别结果:', result);
            ChatLog.info('RECOGNITION', 'result.rawOutputs:', result.rawOutputs);
            ChatLog.info('RECOGNITION', 'result.data:', result.data);
            
            // 获取已识别的结构化数据和类型
            const structuredData = result.data?.structuredData || result.rawOutputs?.structuredData;
            const imageType = result.data?.imageType || result.rawOutputs?.imageType || '发票';
            
            ChatLog.info('RECOGNITION', 'structuredData:', structuredData);
            ChatLog.info('RECOGNITION', 'imageType:', imageType);
            
            if (!structuredData || typeof structuredData !== 'object' || Object.keys(structuredData).length === 0) {
                App.showToast('❌ 没有可用的数据，请重新识别');
                return;
            }
            
            // 添加用户消息
            addMessage('user', '📊 将' + imageType + '信息填入Excel');
            
            ChatLog.info('RECOGNITION', '开始直接填入Excel，不再重新识别');
            
            // 直接调用FillDataToExcel接口，不重新识别图片
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const dataJson = JSON.stringify(structuredData);
                ChatLog.info('RECOGNITION', '准备发送的数据JSON长度:', dataJson.length);
                
                const fillResult = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(dataJson, imageType);
                ChatLog.info('RECOGNITION', 'FillDataToExcel返回结果:', fillResult);
                
                const response = JSON.parse(fillResult);
                if (response.success) {
                    App.showToast('✅ ' + response.message);
                } else {
                    App.showToast('❌ 填入失败: ' + response.error);
                }
            } else {
                App.showToast('❌ 无法连接到Excel');
            }
            
        } catch (error) {
            ChatLog.error('RECOGNITION', '========== handleFillToExcel失败 ==========');
            ChatLog.error('RECOGNITION', error);
            App.showToast('❌ 填入Excel失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L15224-L15300 ---
    /**
     * 尝试从最近的识别结果直接填入Excel（跳过重新识别步骤）
     * @param {string} expectedType - 期望的数据类型（发票/身份证/名片）
     * @returns {boolean} - 是否成功填入
     */
    async function tryFillFromRecentResult(expectedType) {
        try {
            ChatLog.info('RECOGNITION', '========== 尝试从最近识别结果填入Excel ==========');
            ChatLog.info('RECOGNITION', '期望类型:', expectedType);
            
            // 查找最近的识别结果
            if (!window.recognitionResults || Object.keys(window.recognitionResults).length === 0) {
                ChatLog.info('RECOGNITION', '没有找到最近的识别结果');
                return false;
            }
            
            // 获取最近的结果ID（按时间排序）
            const resultIds = Object.keys(window.recognitionResults).sort((a, b) => b.localeCompare(a));
            const recentResultId = resultIds[0];
            const result = window.recognitionResults[recentResultId];
            
            if (!result) {
                ChatLog.info('RECOGNITION', '最近的识别结果为空');
                return false;
            }
            
            ChatLog.info('RECOGNITION', '找到最近的识别结果:', result);
            
            // 获取结构化数据和类型
            const structuredData = result.data?.structuredData || result.rawOutputs?.structuredData;
            const imageType = result.data?.imageType || result.rawOutputs?.imageType;
            
            ChatLog.info('RECOGNITION', 'structuredData:', structuredData);
            ChatLog.info('RECOGNITION', 'imageType:', imageType);
            
            // 检查数据是否有效
            if (!structuredData || typeof structuredData !== 'object' || Object.keys(structuredData).length === 0) {
                ChatLog.info('RECOGNITION', '结构化数据为空或无效');
                return false;
            }
            
            // ✨ 更灵活的类型匹配：只要是期望类型或相关类型都可以
            const isTypeMatch = 
                (expectedType === '身份证' && (imageType === '身份证' || imageType === '名片')) ||
                (expectedType === '发票' && imageType === '发票') ||
                (expectedType === '表格' && imageType === '表格');
            
            if (!isTypeMatch) {
                ChatLog.info('RECOGNITION', `类型不匹配: 期望 ${expectedType}，实际 ${imageType}`);
                // 即使类型不匹配，如果有数据也可以尝试填入
                ChatLog.info('RECOGNITION', '不过有数据，尝试使用通用方式填入');
            }
            
            ChatLog.info('RECOGNITION', '开始直接填入Excel');
            
            // 直接调用FillDataToExcel接口
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const dataJson = JSON.stringify(structuredData);
                ChatLog.info('RECOGNITION', '准备发送的数据JSON长度:', dataJson.length);
                
                const fillResult = await window.chrome.webview.hostObjects.excelBridge.FillDataToExcel(dataJson, imageType || expectedType);
                ChatLog.info('RECOGNITION', 'FillDataToExcel返回结果:', fillResult);
                
                const response = JSON.parse(fillResult);
                if (response.success) {
                    App.showToast('✅ ' + response.message);
                    return true;
                } else {
                    App.showToast('❌ 填入失败: ' + response.error);
                    return false;
                }
            } else {
                App.showToast('❌ 无法连接到Excel');
                return false;
            }
            
        } catch (error) {
            ChatLog.error('RECOGNITION', '========== tryFillFromRecentResult失败 ==========');
            ChatLog.error('RECOGNITION', error);
            return false;
        }
    }

    // --- 来自 app.js:L15305-L15331 ---
    /**
     * 统一插入表格到Excel（追加模式）
     */
    async function insertRecognitionTableToExcel(resultId) {
        try {
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }
            
            if (result.data.rows && Array.isArray(result.data.rows)) {
                await insertTableToExcel(result.data.rows);
                App.showToast('✅ 已插入到 Excel');
            } else {
                // 如果没有rows，使用buildExcelRows构建
                const rows = buildExcelRows(result.data, 'grid');
                if (rows.length > 0) {
                    await insertTableToExcel(rows);
                    App.showToast('✅ 已插入到 Excel');
                } else {
                    App.showToast('⚠️ 没有有效的表格数据');
                }
            }
        } catch (error) {
            // 使用统一错误处理
            window.ErrorHandler.showError(error, 'Excel插入');
            ChatLog.error('RECOGNITION', '插入Excel失败:', error);
        }
    }

    // --- 来自 app.js:L15336-L15360 ---
    /**
     * 统一保存JSON文件功能
     */
    function saveRecognitionDataAsJson(resultId) {
        try {
            const result = window.recognitionResults?.[resultId];
            if (!result) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 数据已过期，请重新识别</span>', true) : App.showToast('❌ 数据已过期，请重新识别');
                return;
            }
            
            const data = result.data;
            const fileName = result.rawOutputs?.fileName || `recognition_${result.type}`;
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName.replace(/\.[^/.]+$/, '') + '_识别结果.json';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
            App.showToast('✅ 已保存JSON文件');
        } catch (error) {
            ChatLog.error('RECOGNITION', '保存失败:', error);
            App.showToast('❌ 保存失败');
        }
    }


    // --- 来自 app.js:L15756-L15857 ---
    /**
     * 通用图片处理函数
     * @param {string} base64Image - 图片的base64数据
     * @param {string} fileName - 文件名
     * @param {string} type - 处理类型（invoice, card, table, general）
     * @param {object} extraInputs - 额外的输入参数
     */
    // 识别结果"有效内容空洞"判定：skill 成功后既无提取字段、也无实际文本/表格。
    // 识别面板路径后端不透传 OCR 置信度，故用"结果空洞"作为纯视觉兜底信号。
    function isRecognitionResultEmpty(response) {
        const outputs = (response && response.outputs) || response || {};
        const metaKeys = ['rawText', 'description', 'text', 'headers', 'rows', 'type', 'success',
            'error', 'executionLog', 'totalExecutionTime', 'extractedText', 'extracted_text',
            'structuredData', 'suggestedAction', 'suggested_action', 'imageType', 'image_type'];
        const userKeys = Object.keys(outputs).filter(function(k) { return metaKeys.indexOf(k) === -1; });
        const hasTable = Array.isArray(outputs.headers) && outputs.headers.length > 0 && Array.isArray(outputs.rows);
        const raw = outputs.rawText || outputs.description || outputs.text || response.rawText || response.description || '';
        if (hasTable || (typeof raw === 'string' && raw.trim().length > 0) || userKeys.length > 0) {
            return false;
        }
        return true;
    }

    // 结果空洞时调用 ImageFieldExtractor skill 纯视觉兜底（与 image-pipeline.js 的 fallbackToVisionSkill 行为一致）
    async function runImageFieldExtractor(base64Image, fileName) {
        const bridge = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge)
            || App.excelBridge || window.excelBridge;
        if (!bridge) return null;
        try {
            const result = await bridge.ExecuteSkillAsync('ImageFieldExtractor', JSON.stringify({ imageData: base64Image, fileName: fileName || 'image' }));
            const resp = JSON.parse(result);
            if (!resp || !resp.success) {
                ChatLog.warn('RECOGNITION', '[Recognition] ImageFieldExtractor 调用失败: ' + ((resp && resp.error) || '未知错误'));
                return null;
            }
            const out = resp.outputs || {};
            const texts = [resp.text, resp.rawText, out.result, out.content, out.markdown, resp.rawText];
            const txt = texts.find(function(t) { return typeof t === 'string' && t.trim(); });
            const vision = (typeof txt === 'string' && txt.trim()) ? txt : (Object.keys(out).length ? JSON.stringify(out) : '');
            if (!vision || /^\[?(无法识别|未识别|\[OCR)/.test(vision.trim())) {
                return null;
            }
            return vision;
        } catch (err) {
            ChatLog.warn('RECOGNITION', '[Recognition] 纯视觉兜底失败', err);
            return null;
        }
    }

    async function processImageGeneric(base64Image, fileName, type = 'general', extraInputs = {}) {
        let config = IMAGE_PROCESSING_CONFIG[type];
        if (!config) {
            ChatLog.error('RECOGNITION', '未知的图片处理类型:', type);
            addMessage('ai', '❌ 未知的处理类型: ' + type);
            return;
        }
        if (typeof config === 'string') {
            config = IMAGE_PROCESSING_CONFIG[config];
            if (!config) {
                ChatLog.error('RECOGNITION', '别名解析失败:', type);
                addMessage('ai', '❌ 配置解析失败: ' + type);
                return;
            }
        }

        const requestId = type + '-' + Date.now();
        
        // 动态获取步骤（支持填入模式）
        const fillMode = extraInputs.fillMode;
        const steps = (typeof config.getSteps === 'function') ? config.getSteps(fillMode) : config.steps;
        
        // 构建初始消息（带动态加载动画）- 使用 addMessage 添加新消息作为进度气泡（临时消息，不进入上下文历史）
        const firstStep = steps[0];
        addMessage('ai', `${config.icon} 正在${config.name}: ${fileName}<br><div style="display: flex; align-items: center; gap: 8px; margin-top: 8px;"><div class="loading-spinner" style="width: 14px; height: 14px; border: 2px solid #e3f2fd; border-top: 2px solid #4a90d9; border-radius: 50%; animation: spin 1s linear infinite;"></div><span>${firstStep.message}</span></div>`, false, null, true);

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 使用 ExecuteSkillAsync 同步调用后端
                const inputs = {
                    imageData: base64Image,
                    fileName: fileName,
                    ...extraInputs
                };
                
                // 依次更新同一个气泡显示步骤进度
                for (let i = 0; i < steps.length; i++) {
                    const step = steps[i];
                    updateLastMessage('ai', `${config.icon} 正在${config.name}: ${fileName}<br><div style="display: flex; align-items: center; gap: 8px; margin-top: 8px;"><div class="loading-spinner" style="width: 14px; height: 14px; border: 2px solid #e3f2fd; border-top: 2px solid #4a90d9; border-radius: 50%; animation: spin 1s linear infinite;"></div><span>${step.message}</span></div>`);
                    
                    // 模拟步骤延迟
                    if (i < steps.length - 1) {
                        await new Promise(resolve => setTimeout(resolve, step.duration));
                    }
                }
                
                // 调用后端
                const result = await window.chrome.webview.hostObjects.excelBridge.ExecuteSkillAsync(
                    config.skillName,
                    JSON.stringify(inputs)
                );
                
                const response = JSON.parse(result);

                // OCR → 纯视觉兜底门控：backend skill 成功但结果空洞时降级 ImageFieldExtractor
                let usedVisionFallback = false;
                if (response.success && isRecognitionResultEmpty(response)) {
                    ChatLog.warn('RECOGNITION', '[Recognition] 后端提取结果空洞，尝试 ImageFieldExtractor 纯视觉兜底');
                    const vision = await runImageFieldExtractor(base64Image, fileName);
                    if (vision) {
                        response.rawText = vision;
                        response.outputs = Object.assign({}, response.outputs || {}, { extractedText: vision, rawText: vision });
                        usedVisionFallback = true;
                        ChatLog.info('RECOGNITION', '[Recognition] 已使用 ImageFieldExtractor 兜底结果');
                    } else {
                        ChatLog.warn('RECOGNITION', '[Recognition] 纯视觉兜底失败，回落原提取结果');
                    }
                }

                if (response.success) {
                    // 使用统一的结果处理框架 ✨
                    // 添加imageData和fileName到extraInputs，供后续填入Excel使用
                    const extendedExtraInputs = {
                        ...extraInputs,
                        imageData: base64Image,
                        fileName: fileName
                    };
                    const resultHtml = await formatRecognitionResultUniversal(type, response, extendedExtraInputs);

                    // 降级发生在视觉兜底时，在结果前标注执行方式（与聊天 skill 徽章观感一致）
                    let finalHtml = resultHtml;
                    if (usedVisionFallback) {
                        const badgeHtml = '<div class="skill-badge">' +
                            '<span class="skill-badge-icon">⚙️</span>' +
                            '<span class="skill-badge-label">Skill: ImageFieldExtractor</span>' +
                            ' · ' +
                            '<span class="skill-pipeline-tag pipeline-mcp">纯视觉兜底</span>' +
                            '</div>';
                        finalHtml = badgeHtml + resultHtml;
                        ChatLog.info('RECOGNITION', '[Recognition] 识别结果已标注「纯视觉兜底」徽章');
                    }
                    updateLastMessage('ai', finalHtml);

                    // ✨ 只有在填入模式下才清理附件（与发票完全一致）
                    // 只识别模式（recognize_only）：保留附件和按钮，允许用户继续操作
                    if (fillMode && fillMode !== 'recognize_only') {
                        clearAttachment();
                    }
                    
                    // 保存到历史记录（更新当前对话）
                    const messages = getChatMessages();
                    addToHistory(`${config.name}: ${fileName}`, messages);
                } else {
                    updateLastMessage('ai', `❌ ${config.name}失败: ` + (response.error || '未知错误'));
                    clearAttachment();  // 失败也要清理
                    
                    // 保存到历史记录（更新当前对话）
                    const messages = getChatMessages();
                    addToHistory(`${config.name}失败: ${fileName}`, messages);
                }

            } else if (window.chrome && window.chrome.webview) {
                // 有 WebView2 但没有 hostObjects，使用 postMessage 方式
                addMessage('ai', `❌ ${config.name}失败：无法访问 Excel 后端接口`);
                clearAttachment();
            } else {
                // 非 WebView2 环境，显示错误
                addMessage('ai', `❌ ${config.name}失败：需要在 Excel 插件环境中运行`);
                clearAttachment();
            }

        } catch (error) {
            addMessage('ai', `❌ ${config.name}失败: ` + error.message);
            clearAttachment();  // 异常也要清理
            
            // 保存到历史记录（更新当前对话）
            const messages = getChatMessages();
            addToHistory(`${config.name}异常: ${fileName}`, messages);
        }
    }

    // 格式化识别结果为卡片表格HTML（来自原app.js:L1065）
        function formatRecognitionResult(data, imageType) {
                // ============================================================
                // 【SmartTieredRenderer 集成】智能分层渲染器 v1.0.0
                // ============================================================
                if (window.SmartTieredRenderer || (window.App && window.App.SmartTieredRenderer)) {
                    try {
                        ChatLog.info('RECOGNITION', '✅ 使用 SmartTieredRenderer 渲染');
                        
                        var rawText = data.rawText || data.description || data.text || '';
                        var isTableData = data.headers && data.rows &&
                                        Array.isArray(data.headers) && Array.isArray(data.rows) &&
                                        data.headers.length > 0;
                        
                        var thinkingContent = '';
                        if (rawText && typeof rawText === 'string') {
                            var lines = rawText.split('\n').filter(function(line) { return line.trim(); });
                            thinkingContent = lines.slice(0, Math.min(lines.length, 8)).join('\n');
                            if (lines.length > 8) {
                                thinkingContent += '\n...（更多内容见下方分析总结）';
                            }
                        }
                        
                        var responseData = {
                            thinking: thinkingContent || undefined,
                            toolCalls: [{
                                name: imageType === 'table' ? 'table_recognition' : 
                                    imageType === 'invoice' ? 'invoice_ocr' :
                                    imageType === 'idcard' ? 'idcard_recognition' : 'image_analysis',
                                parameters: { imageType: imageType || 'unknown', timestamp: new Date().toISOString() },
                                result: { 
                                    success: true,
                                    dataType: isTableData ? 'table' : Object.keys(data).length > 3 ? 'structured' : 'text',
                                    fieldsCount: isTableData ? data.headers.length : Object.keys(data).length,
                                    recordsCount: isTableData ? data.rows.length : 1
                                },
                                status: 'success',
                                duration: null
                            }],
                            result: isTableData ? {
                                headers: data.headers,
                                rows: data.rows,
                                title: (imageType === 'table' ? '表格数据提取' : 
                                    imageType === 'invoice' ? '发票信息识别' : 
                                    imageType === 'idcard' ? '身份证信息识别' : '图片文本提取结果')
                            } : data,
                            analysis: rawText || undefined,
                            metadata: { model: 'AI Vision Model', version: '1.0.0', recognitionType: imageType }
                        };
                        
                        var renderer = window.App && window.App.SmartTieredRenderer ? window.App.SmartTieredRenderer : window.SmartTieredRenderer;
                        var html = renderer.render(responseData);
                        var wrappedHtml = wrapSmartTieredHtmlWithInsertButton(html, imageType, data);
                        ChatLog.info('RECOGNITION', '🎨 SmartTieredRenderer 渲染完成');
                        return wrappedHtml;
                        
                    } catch (error) {
                        ChatLog.error('RECOGNITION', '❌ SmartTieredRenderer 渲染失败，回退到传统模式:', error);
                    }
                } else {
                    ChatLog.warn('RECOGNITION', '⚠️ SmartTieredRenderer 未加载，使用传统渲染模式');
                }
                // 【SmartTieredRenderer 集成结束】

            ChatLog.info('RECOGNITION', '🔍 formatRecognitionResult 被调用:', { imageType: imageType, dataKeys: Object.keys(data) });
        if (!data || typeof data !== 'object') {
            ChatLog.info('RECOGNITION', '⚠️ 数据无效:', data);
            return '<div style="color: #999; padding: 16px;">暂无识别数据</div>';
        }

        var transformedData = transformFieldNames(data, imageType);
        var headerOrder = getHeaderOrder(imageType);
        var orderedData = {};
        headerOrder.forEach(function(field) {
            if (transformedData.hasOwnProperty(field)) {
                orderedData[field] = transformedData[field];
                delete transformedData[field];
            }
        });
        Object.assign(orderedData, transformedData);

        var rowsHtml = '';
        Object.entries(orderedData).forEach(function(entry) {
            var key = entry[0], value = entry[1];
            if (value === undefined || value === null || value === '' ||
                ['imageType','image_type','description','structuredData',
                 'suggestedAction','suggested_action','extractedText',
                 'extracted_text','type','success','error','executionLog',
                 'totalExecutionTime'].indexOf(key) !== -1) {
                return;
            }
            var processedValue = String(value);
            if (key === 'rawText' || key === 'extractedText' || key === 'description') {
                if (processedValue.indexOf('<div') !== -1 || processedValue.indexOf('<br') !== -1) {
                    processedValue = processedValue
                        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
                        .replace(/javascript:/gi, '');
                } else {
                    processedValue = processedValue
                        .replace(/\\r\\n/g, '<br>').replace(/\\n/g, '<br>')
                        .replace(/\r\n/g, '<br>').replace(/\r/g, '<br>')
                        .replace(/\n/g, '<br>');
                }
            }
            var escapedValue = (key === 'rawText' && (String(value).indexOf('<div') !== -1 || String(value).indexOf('<br') !== -1))
                ? processedValue : escapeHtml(processedValue);
            rowsHtml += '<tr><th>' + escapeHtml(key) + '</th><td>' + escapedValue + '</td></tr>';
        });

        var typeIcons = {'身份证':'','名片':'📇','发票':'🧾','表格':'📊','文档':'📝','图片':'🖼️'};
        var icon = typeIcons[imageType] || '🔍';
        var typeNames = {'身份证':'身份证','名片':'名片','发票':'发票','表格':'表格数据','文档':'文档内容','图片':'图片分析'};
        var typeName = typeNames[imageType] || '识别结果';
        var dataJson = JSON.stringify(data);
        var html = '<div class="recognition-result-container"><h3>' + typeName + '提取</h3><div class="recognition-result-card"><div class="recognition-header"><div class="recognition-title"><span class="recognition-title-icon">' + icon + '</span><span>表格</span></div><div class="recognition-actions">' +
            ((imageType === '身份证' || imageType === '名片' || imageType === '发票')
                ? '<button class="recognition-action-btn" data-json="' + escapeHtml(dataJson) + '" data-type="' + escapeHtml(imageType) + '" onclick="fillRecognitionDataToExcel(this)" title="填入Excel" style="background:#4caf50;color:white;">📊</button>'
                : '<button class="recognition-action-btn" onclick="App.recognition.insertRecognitionKvToExcel(this)" title="插入到Excel" style="background:#4caf50;color:white;">📊</button>') +
            '<button class="recognition-action-btn" onclick="copyRecognitionResult(this)" title="复制数据">📋</button>' +
            '<button class="recognition-action-btn" onclick="downloadRecognitionResult(this)" title="下载数据">⬇️</button>' +
            '</div></div><table class="recognition-table"><tbody>' + rowsHtml + '</tbody></table>' +
            '<div class="recognition-footer"><p>说明：该' + (imageType === '身份证' ? '身份证' : '识别结果') + '为系统自动识别提取，仅供参考，请核对原始信息</p></div>' +
            '</div></div>';
        ChatLog.info('RECOGNITION', '🔍 生成的HTML:', html.substring(0, 200) + '...');
        return html;
    }


    // ============================================================
    // 导出公共API（挂载到 App 对象）
    // ============================================================
    window.App = window.App || {};
    window.App.recognition = {
        // 区域1: 识别结果基础操作
        copyRecognitionResult: copyRecognitionResult,
        downloadRecognitionResult: downloadRecognitionResult,
        fillRecognitionDataToExcel: fillRecognitionDataToExcel,
        insertRecognitionKvToExcel: insertRecognitionKvToExcel,
        insertSmartTieredResultToExcel: insertSmartTieredResultToExcel,
        fillJsonToExcel: fillJsonToExcel,
        fillMultiTableToExcel: fillMultiTableToExcel,
        copyMultiTable: copyMultiTable,
        downloadMultiTable: downloadMultiTable,

        // 区域2: 数据解析与渲染
        parseRecognitionData: parseRecognitionData,
        collectDisplayFields: collectDisplayFields,
        buildExcelRows: buildExcelRows,
        renderDataTable: renderDataTable,
        renderActionButtons: renderActionButtons,
        storeRecognitionResult: storeRecognitionResult,
        formatRecognitionResultUniversal: formatRecognitionResultUniversal,
        formatRecognitionResult: formatRecognitionResult,
        copyRecognitionData: copyRecognitionData,
        isRecognitionResultEmpty: isRecognitionResultEmpty,
        runImageFieldExtractor: runImageFieldExtractor,
        insertVerticalToExcel: insertVerticalToExcel,
        fillRecognitionDataToExcelByResultId: fillRecognitionDataToExcelByResultId,
        fillToExcel: fillToExcel,
        retryFillToExcel: retryFillToExcel,
        handleFillToExcel: handleFillToExcel,
        tryFillFromRecentResult: tryFillFromRecentResult,
        insertRecognitionTableToExcel: insertRecognitionTableToExcel,
        saveRecognitionDataAsJson: saveRecognitionDataAsJson,

        // 区域3: 处理流程
        processImageGeneric: processImageGeneric
    };

})();