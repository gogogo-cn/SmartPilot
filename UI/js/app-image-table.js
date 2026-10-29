/**
 * ============================================================
 * app-image-table.js — Image Table OCR
 * ============================================================
 * 功能：图片表格识别模块，专门处理含表格图片的结构化提取，支持表格线检测与单元格内容识别。
 *
 * 核心类/函数：
 *   - recognizeTableImage() — 识别图片表格
 *   - extractTableStructure() — 提取表格结构
 *   - convertTableToExcel() — 转换表格为 Excel 数据
 *
 * 依赖：
 *   - app-ocr.js
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    // --- 来自 app.js:L7184-L7268 ---
    async function handlePaste(e) {
        ChatLog.info('IMAGE', '粘贴事件触发');

        // 首先检查是否是表格数据（优先处理表格）
        const htmlData = e.clipboardData.getData('text/html');
        const textData = e.clipboardData.getData('text/plain');

        // 检测是否包含表格（优先于图片处理）
        if (htmlData && (htmlData.includes('<table') || htmlData.includes('<TABLE'))) {
            ChatLog.info('IMAGE', '检测到 HTML 表格数据，优先处理表格');
            e.preventDefault();
            e.stopPropagation();
            await handlePastedTable(htmlData, textData);
            return;
        }

        // 检测文本数据是否可能是表格（制表符分隔）
        if (textData && textData.includes('\t')) {
            ChatLog.info('IMAGE', '检测到制表符分隔的表格数据，优先处理表格');
            e.preventDefault();
            e.stopPropagation();
            await handlePastedTableText(textData);
            return;
        }

        // 如果不是表格，再检查图片
        let hasImage = false;

        // 检查 files（某些浏览器/系统直接提供文件）
        if (e.clipboardData.files && e.clipboardData.files.length > 0) {
            const files = e.clipboardData.files;
            for (let i = 0; i < files.length; i++) {
                if (files[i].type.indexOf('image') !== -1) {
                    hasImage = true;
                    e.preventDefault();
                    e.stopPropagation();
                    ChatLog.info('IMAGE', '从 files 中获取到图片:', files[i].name, files[i].type, files[i].size);
                    await handlePastedFile(files[i]);
                    return;
                }
            }
        }

        // 检查 items
        const items = e.clipboardData.items;
        if (items && !hasImage) {
            for (let i = 0; i < items.length; i++) {
                const item = items[i];
                ChatLog.info('IMAGE', '粘贴项类型:', item.type);

                // 检查是否是图片
                if (item.type.indexOf('image') !== -1) {
                    hasImage = true;
                    e.preventDefault();
                    e.stopPropagation();

                    // 使用 getAsFile 获取文件
                    const file = item.getAsFile();
                    if (file) {
                        ChatLog.info('IMAGE', '从 items 中获取到图片:', file.name, file.type, file.size);
                        await handlePastedFile(file);
                        return;
                    }
                }
            }
        }

        // 检查 text/html（某些截图工具会提供 HTML 格式的图片）
        if (!hasImage) {
            if (htmlData && htmlData.includes('<img')) {
                ChatLog.info('IMAGE', '检测到 HTML 格式的图片');
                // 尝试从 HTML 中提取图片
                const imgMatch = htmlData.match(/src="([^"]+)"/);
                if (imgMatch && imgMatch[1].startsWith('data:image')) {
                    hasImage = true;
                    e.preventDefault();
                    e.stopPropagation();
                    await handlePastedDataUrl(imgMatch[1]);
                    return;
                }
            }
        }

        ChatLog.info('IMAGE', '未检测到图片或表格，允许默认粘贴行为');
    }

    // --- 来自 app.js:L7273-L7322 ---
    async function handlePastedTable(htmlData, textData) {
        ChatLog.info('IMAGE', '处理粘贴的 HTML 表格');

        try {
            // 解析 HTML 表格
            const parser = new DOMParser();
            const doc = parser.parseFromString(htmlData, 'text/html');
            const table = doc.querySelector('table');

            if (!table) {
                ChatLog.info('IMAGE', '未找到表格元素，尝试文本格式');
                if (textData) {
                    await handlePastedTableText(textData);
                }
                return;
            }

            // 提取表格数据
            const rows = [];
            const trElements = table.querySelectorAll('tr');

            trElements.forEach(tr => {
                const row = [];
                const cells = tr.querySelectorAll('td, th');
                cells.forEach(cell => {
                    row.push(cell.textContent.trim());
                });
                if (row.length > 0) {
                    rows.push(row);
                }
            });

            ChatLog.info('IMAGE', '提取到表格数据:', rows.length, '行');

            if (rows.length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">未能识别表格数据</span>', true) : App.showToast('未能识别表格数据');
                return;
            }

            // 显示表格预览
            await showTablePreview(rows, 'HTML');

        } catch (error) {
            ChatLog.error('IMAGE', '处理 HTML 表格失败:', error);
            // 尝试文本格式
            if (textData) {
                await handlePastedTableText(textData);
            }
        }
    }

    // --- 来自 app.js:L7325-L7357 ---
    async function handlePastedTableText(textData) {
        ChatLog.info('IMAGE', '处理粘贴的文本表格');

        try {
            // 解析制表符分隔的数据
            const lines = textData.split(/\r?\n/);
            const rows = [];

            lines.forEach(line => {
                if (line.trim()) {
                    // 按制表符分割
                    const cells = line.split('\t');
                    if (cells.length > 0) {
                        rows.push(cells.map(cell => cell.trim()));
                    }
                }
            });

            ChatLog.info('IMAGE', '提取到文本表格数据:', rows.length, '行');

            if (rows.length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">未能识别表格数据</span>', true) : App.showToast('未能识别表格数据');
                return;
            }

            // 显示表格预览
            await showTablePreview(rows, '文本');

        } catch (error) {
            ChatLog.error('IMAGE', '处理文本表格失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">表格识别失败</span>', true) : App.showToast('表格识别失败');
        }
    }

    // --- 来自 app.js:L7360-L7409 ---
    async function showTablePreview(rows, sourceType) {
        ChatLog.info('IMAGE', '显示表格预览:', rows.length, '行, 来源:', sourceType);

        // 构建表格 HTML
        let tableHtml = '<div style="margin: 10px 0; max-height: 300px; overflow: auto;">';
        tableHtml += '<table style="border-collapse: collapse; width: 100%; font-size: 13px;">';

        rows.forEach((row, rowIndex) => {
            tableHtml += '<tr>';
            row.forEach((cell, cellIndex) => {
                const isHeader = rowIndex === 0;
                const tag = isHeader ? 'th' : 'td';
                const bgColor = isHeader ? '#f0f0f0' : 'white';
                tableHtml += `<${tag} style="border: 1px solid #ddd; padding: 8px; background: ${bgColor}; text-align: left;">${App.escapeHtml(cell)}</${tag}>`;
            });
            tableHtml += '</tr>';
        });

        tableHtml += '</table></div>';

        // 添加操作按钮
        const actionsHtml = `
            <div style="margin-top: 10px; display: flex; gap: 8px;">
                <button onclick="App.ImageTable.insertTableToExcel(${JSON.stringify(rows).replace(/"/g, '&quot;')})"
                        style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer;">
                    📊 插入到 Excel
                </button>
                <button onclick="App.ImageTable.copyTableToClipboard(${JSON.stringify(rows).replace(/"/g, '&quot;')})"
                        style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">
                    📋 复制表格
                </button>
                <button onclick="this.closest('.message-wrapper').remove()"
                        style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">
                    ❌ 关闭
                </button>
            </div>
        `;

        // 显示在聊天窗口
        const messageContent = `
            <div style="background: #f9f9f9; padding: 12px; border-radius: 8px; border: 1px solid #e0e0e0;">
                <div style="font-weight: 500; margin-bottom: 8px;">📋 检测到 ${sourceType} 表格 (${rows.length} 行 × ${rows[0]?.length || 0} 列)</div>
                ${tableHtml}
                ${actionsHtml}
            </div>
        `;

        App.addMessage('ai', messageContent);
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已识别 ' + sourceType + ' 表格: ' + rows.length + ' 行</span>', true) : App.showToast(`已识别 ${sourceType} 表格: ${rows.length} 行`);
    }

    // --- 来自 app.js:L7412-L7419 ---
    async function insertLastTableToExcel() {
        const rows = App.lastRecognizedTable;
        if (!rows) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 没有可插入的表格数据</span>', true) : App.showToast('❌ 没有可插入的表格数据');
            return;
        }
        await insertTableToExcel(rows);
    }

    // --- 来自 app.js:L7422-L7429 ---
    async function copyLastTableToClipboard() {
        const rows = App.lastRecognizedTable;
        if (!rows) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可复制的表格数据</span>', true) : App.showToast('❌ 没有可复制的表格数据');
            return;
        }
        await copyTableToClipboard(rows);
    }

    // --- 来自 app.js:L7432-L7439 ---
    function editLastTableData() {
        const rows = App.lastRecognizedTable;
        if (!rows) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可编辑的表格数据</span>', true) : App.showToast('❌ 没有可编辑的表格数据');
            return;
        }
        editTableData(rows);
    }

    // --- 来自 app.js:L7442-L7484 ---
    async function insertTableToExcel(rows, options = {}) {
        // 清理默认列名
        const cleanRows = cleanTableHeaders(rows);
        ChatLog.info('IMAGE', '插入表格到 Excel:', cleanRows);

        // 默认选项：从A1开始
        const { startCell = 'A1' } = options;

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {

                // 步骤1：选中目标单元格
                try {
                    await window.chrome.webview.hostObjects.excelBridge.ExecuteSkillAsync(
                        'ExcelOperations',
                        JSON.stringify({
                            operation: 'selectRange',
                            range: startCell
                        })
                    );
                    ChatLog.info('IMAGE', `✅ 已选中 ${startCell}`);
                } catch (selectError) {
                    ChatLog.warn('IMAGE', '⚠️ 无法选中单元格:', selectError);
                }

                // 步骤2：插入表格数据（后端 CardProcessing/InvoiceProcessing 会自动格式化）
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows));
                const resultObj = JSON.parse(result);

                if (resultObj.success) {
                    ChatLog.info('IMAGE', '✅ 表格已成功插入到 Excel');
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 数据已填入 Excel</span>', true) : App.showToast('✅ 数据已填入 Excel');
                } else {
                    ChatLog.error('IMAGE', '❌ 插入失败:', resultObj.error);
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 插入失败: ' + (resultObj.error || '未知错误') + '</span>', true) : App.showToast('❌ 插入失败: ' + (resultObj.error || '未知错误'));
                }
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> Excel 连接不可用</span>', true) : App.showToast(' Excel 连接不可用');
            }
        } catch (error) {
            ChatLog.error('IMAGE', '插入表格失败:', error);
            // 使用统一错误处理
            window.ErrorHandler.showError(error, '表格插入');
        }
    }

    // --- 来自 app.js:L7487-L7499 ---
    async function copyTableToClipboard(rows) {
        ChatLog.info('IMAGE', '复制表格到剪贴板');

        try {
            // 转换为制表符分隔的文本
            const text = rows.map(row => row.join('\t')).join('\n');
            await navigator.clipboard.writeText(text);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 表格已复制到剪贴板</span>', true) : App.showToast('✅ 表格已复制到剪贴板');
        } catch (error) {
            ChatLog.error('IMAGE', '复制表格失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败</span>', true) : App.showToast('❌ 复制失败');
        }
    }

    // --- 来自 app.js:L7502-L7545 ---
    async function downloadTableAsCsv(rows) {
        ChatLog.info('IMAGE', '下载表格为CSV:', rows);

        try {
            if (!rows || rows.length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有数据可下载</span>', true) : App.showToast('❌ 没有数据可下载');
                return;
            }

            // 使用BOM头确保Excel正确识别UTF-8编码
            let csv = '\uFEFF';

            // 转换每一行为CSV格式
            rows.forEach(row => {
                const csvRow = row.map(cell => {
                    // 处理包含逗号、引号或换行符的单元格
                    const cellText = String(cell || '').replace(/"/g, '""');
                    if (cellText.includes(',') || cellText.includes('"') || cellText.includes('\n')) {
                        return `"${cellText}"`;
                    }
                    return cellText;
                }).join(',');
                csv += csvRow + '\n';
            });

            // 创建Blob并下载
            const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const url = URL.createObjectURL(blob);
            link.setAttribute('href', url);
            link.setAttribute('download', `表格数据_${Date.now()}.csv`);
            link.style.visibility = 'hidden';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(url);

            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ CSV文件已下载</span>', true) : App.showToast('✅ CSV文件已下载');
        } catch (error) {
            ChatLog.error('IMAGE', '下载CSV失败:', error);
            // 使用统一错误处理
            window.ErrorHandler.showError(error, 'CSV下载');
        }
    }

    // --- 来自 app.js:L7575-L7845 ---
    async function processImageTableRecognition(base64Image, fileName) {
        ChatLog.info('IMAGE', 'processImageTableRecognition 被调用:', fileName);

        // ===== 首先检查是否有有效的 API 配置 =====
        // Ollama 本地模型不需要 apiKey，只需要 apiUrl 和 apiType
        const config = App.config;
        const hasValidApiConfig = config && config.apiUrl && config.apiType &&
            (config.apiType === 'ollama' || config.apiKey); // Ollama 不需要 apiKey，其他需要

        if (!hasValidApiConfig) {
            ChatLog.info('IMAGE', '⚠️ 未配置有效的 API，跳过 AI 表格识别，直接使用本地 OCR');
            App.addMessage('ai', `<div style="background: #e3f2fd; padding: 12px; border-radius: 8px; border-left: 4px solid #2196f3;">
                <div style="font-weight: 500; margin-bottom: 8px;">ℹ️ 未配置AI模型</div>
                <div style="color: #666; font-size: 12px;">将使用本地 OCR 识别表格...</div>
            </div>`);

            // 直接调用本地 OCR
            try {
                const localResult = await recognizeTableWithLocalOCR(base64Image, fileName);
                if (localResult) {
                    return;
                }
            } catch (localError) {
                ChatLog.error('IMAGE', '本地 OCR 识别失败:', localError);
            }

            App.addMessage('ai', '❌ 表格识别失败：未配置 AI 模型且本地 OCR 识别失败');
            return;
        }

        // ===== 确定使用的模型 =====
        ChatLog.info('IMAGE', '========== 确定表格识别使用的模型 ==========');

        let currentModelForAI;
        let switchedModelName = null;

        // 使用当前模型（图片已走 OCR 路径，无需视觉模型）
        if (App.isAutoMode()) {
            ChatLog.info('IMAGE', '🤖 [AUTO] 表格识别 - 使用 AUTO 智能选择');
            currentModelForAI = App.getEffectiveModel('表格识别图片', false);
            switchedModelName = currentModelForAI;
            ChatLog.info('IMAGE', '🤖 [AUTO] 选择的模型:', currentModelForAI);
        } else {
            currentModelForAI = App.currentChatModel || config.model || '';
            ChatLog.info('IMAGE', '当前模型:', currentModelForAI);
        }

        // 添加带有进度的消息
        const messageId = 'table-recognition-' + Date.now();
        const modelSwitchStep = switchedModelName ? `
            <div id="${messageId}-step0" style="color: #4caf50;">✅ 使用模型: ${switchedModelName}</div>
        ` : '';

        App.addMessage('ai', `<div id="${messageId}">
            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                <div class="loading-spinner" style="
                    width: 16px;
                    height: 16px;
                    border: 2px solid #e3f2fd;
                    border-top: 2px solid #4a90d9;
                    border-radius: 50%;
                    animation: spin 1s linear infinite;
                "></div>
                <span style="color: #4a90d9; font-weight: 500;">正在提取表格内容...</span>
            </div>
            <div style="color: #666; font-size: 12px; padding-left: 26px;">
                ${modelSwitchStep}
                <div id="${messageId}-step1">📤 正在上传图片...</div>
                <div id="${messageId}-step2" style="color: #ccc;">⏳ 正在识别表格结构...</div>
                <div id="${messageId}-step3" style="color: #ccc;">📊 正在解析表格数据...</div>
            </div>
        </div>`);

        // 更新步骤状态
        const updateStep = (step) => {
            for (let i = 1; i <= 3; i++) {
                const stepEl = document.getElementById(`${messageId}-step${i}`);
                if (stepEl) {
                    if (i < step) {
                        stepEl.style.color = '#4caf50';
                        stepEl.innerHTML = stepEl.innerHTML.replace('⏳', '✅').replace('📤', '✅').replace('📊', '✅');
                    } else if (i === step) {
                        stepEl.style.color = '#4a90d9';
                    } else {
                        stepEl.style.color = '#ccc';
                    }
                }
            }
        };

        // 步骤1完成
        setTimeout(() => updateStep(2), 500);

        try {
            // 首先尝试使用 AI 分析图片中的表格
            const result = await analyzeImageTableWithAI(base64Image, fileName, messageId, updateStep);

            // 检查是否是非视觉模型错误或原始文本
            if (result && result.isNonVisionModel) {
                ChatLog.info('IMAGE', '检测到非视觉模型或原始文本:', result);

                // 如果是原始文本（AI 返回了内容但格式不对），显示给用户
                if (result.isRawText && result.response) {
                    ChatLog.info('IMAGE', '显示 AI 原始文本结果');

                    // 自动检测并包装JSON内容
                    const processedResponse = App.autoDetectAndWrapJSON(result.response);

                    // 检查是否是JSON格式
                    const isJsonFormat = processedResponse.includes('```json');

                    if (isJsonFormat) {
                        // 使用代码块样式显示JSON
                        App.addMessage('ai', `
                            <div style="background: #e3f2fd; padding: 12px; border-radius: 8px; border-left: 4px solid #2196f3;">
                                <div style="font-weight: 500; margin-bottom: 8px;">📝 AI 识别结果</div>
                                <div style="color: #666; font-size: 12px; margin-bottom: 8px;">
                                    ⚠️ AI 返回的内容不是标准表格格式，已显示原始识别结果
                                </div>
                                ${App.parseCodeBlocks(processedResponse)}
                            </div>
                        `);
                    } else {
                        // 使用原始pre标签显示非JSON内容
                        App.addMessage('ai', `
                            <div style="background: #e3f2fd; padding: 12px; border-radius: 8px; border-left: 4px solid #2196f3;">
                                <div style="font-weight: 500; margin-bottom: 8px;">📝 AI 识别结果</div>
                                <div style="color: #666; font-size: 12px; margin-bottom: 8px;">
                                    ⚠️ AI 返回的内容不是标准表格格式，已显示原始识别结果
                                </div>
                                <pre style="background: white; padding: 10px; border-radius: 4px; overflow: auto; max-height: 300px; font-size: 12px; white-space: pre-wrap; border: 1px solid #e0e0e0;">${App.escapeHtml(result.response.substring(0, 2000))}${result.response.length > 2000 ? '\n... (已截断)' : ''}</pre>
                                <div style="margin-top: 10px;">
                                    <button onclick="App.ImageTable.copyOcrResultToClipboard(); App.lastOcrResult = \`${App.escapeHtml(result.replace(/`/g, '\\`'))}\`" style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px; margin-right: 8px;">
                                        📋 复制文字
                                    </button>
                                </div>
                            </div>
                        `);
                    }

                    // 保存 OCR 结果供后续使用
                    App.lastOcrResult = result.response;
                } else {
                    // 否则切换到本地 OCR
                    ChatLog.info('IMAGE', '切换到本地 OCR 识别');

                    // 更新步骤状态
                    if (updateStep) {
                        updateStep(2);
                    }

                    // 更新消息显示正在使用本地 OCR
                    const messageDiv = document.getElementById(messageId);
                    if (messageDiv) {
                        messageDiv.innerHTML = `
                            <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 8px;">
                                <div class="loading-spinner" style="
                                    width: 16px;
                                    height: 16px;
                                    border: 2px solid #e3f2fd;
                                    border-top: 2px solid #4a90d9;
                                    border-radius: 50%;
                                    animation: spin 1s linear infinite;
                                "></div>
                                <span style="color: #4a90d9; font-weight: 500;">正在使用本地 OCR 识别表格...</span>
                            </div>
                            <div style="color: #666; font-size: 12px; padding-left: 26px;">
                                <div style="color: #4caf50;">✅ 检测到非视觉模型，自动切换</div>
                                <div style="color: #4a90d9;">⏳ 正在使用本地 OCR 引擎识别...</div>
                                <div style="color: #ccc;">📊 正在解析表格结构...</div>
                            </div>
                        `;
                    }

                    // 使用本地 OCR 识别
                    await processImageTableWithLocalOCR(base64Image, fileName, messageId);
                }
            }

            // 识别成功，移除进度消息（因为结果已经以新消息形式显示）
            const progressMessage = document.getElementById(messageId);
            if (progressMessage) {
                const wrapper = progressMessage.closest('.message-wrapper');
                if (wrapper) {
                    wrapper.remove();
                }
            }

            // 更新原来的loading消息为完成状态（修复loading不消失的问题）
            const messages = document.querySelectorAll('.message-wrapper.ai .message-content');
            if (messages.length > 0) {
                const lastAiMessage = messages[messages.length - 1];
                if (lastAiMessage.innerHTML.includes('正在提取表格内容')) {
                    const originalContent = lastAiMessage.innerHTML;
                    const updatedContent = originalContent.replace(
                        /<div[^>]*>[\s]*🔄 正在提取表格内容\.\.\.<\/div>/,
                        '<div style="display: flex; align-items: center; gap: 8px; margin-top: 8px;"><span style="color: #4caf50;">✅</span><span>表格内容提取完成</span></div>'
                    ).replace(
                        /<div class="loading-spinner"[^>]*><\/div>\s*<span>正在提取表格内容\.\.\.<\/span>/,
                        '<span style="color: #4caf50;">✅ 表格内容提取完成</span>'
                    );
                    lastAiMessage.innerHTML = updatedContent;
                }
            }

            // 识别成功，清除附件预览
            App.clearAttachment();

        } catch (error) {
            ChatLog.error('IMAGE', '图片表格识别失败:', error);

            // 显示错误状态
            const messageDiv = document.getElementById(messageId);
            if (messageDiv) {
                messageDiv.innerHTML = `
                    <div style="color: #f44336; display: flex; align-items: center; gap: 8px;">
                        <span>❌</span>
                        <span>表格识别失败: ${error.message}</span>
                    </div>
                `;
            }

            const attachmentActions = document.getElementById('attachment-actions');
            if (attachmentActions) {
                attachmentActions.innerHTML = '<span style="color: #f44336; font-size: 12px;">❌ 识别失败</span>';
            }
        }
    }

    // --- 来自 app.js:L7848-L7934 ---
    async function processImageTableWithLocalOCR(base64Image, fileName, messageId = null) {
        ChatLog.info('IMAGE', '使用本地 OCR 识别表格:', fileName);

        try {
            // 将 base64 转换为文件路径（通过 WebView2 桥接）
            let imagePath = null;
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 保存 base64 图片到临时文件
                const result = await window.chrome.webview.hostObjects.excelBridge.SaveBase64Image(base64Image, fileName);
                const resultObj = JSON.parse(result);
                if (resultObj.success) {
                    imagePath = resultObj.path;
                }
            }

            if (!imagePath) {
                throw new Error('无法保存图片文件');
            }

            ChatLog.info('IMAGE', '图片已保存到:', imagePath);

            // 调用本地 OCR 服务
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const ocrResult = await window.chrome.webview.hostObjects.excelBridge.RecognizeTextWithOCR(imagePath);
                ChatLog.info('IMAGE', 'OCR 识别结果:', ocrResult);

                if (ocrResult.startsWith('【错误：')) {
                    throw new Error(ocrResult.replace('【错误：', '').replace('】', ''));
                }

                // 更新步骤状态
                if (messageId) {
                    const step3El = document.getElementById(`${messageId}-step3`);
                    if (step3El) {
                        step3El.style.color = '#4a90d9';
                    }
                }

                // 显示原始 OCR 文字结果
                ChatLog.info('IMAGE', '本地 OCR 识别完成，显示原始文字');
                App.addMessage('ai', `
                    <div style="background: #f5f5f5; padding: 12px; border-radius: 8px;">
                        <div style="font-weight: 500; margin-bottom: 8px; display: flex; justify-content: space-between; align-items: center;">
                            <span>📝 OCR 文字识别结果</span>
                            <span style="font-size: 11px; color: #999; font-weight: normal;">本地 OCR 引擎</span>
                        </div>
                        <pre id="ocr-result-text" style="background: white; padding: 10px; border-radius: 4px; overflow: auto; max-height: 300px; font-size: 12px; white-space: pre-wrap; border: 1px solid #e0e0e0;">${App.escapeHtml(ocrResult)}</pre>
                        <div style="margin-top: 10px; display: flex; gap: 8px; flex-wrap: wrap;">
                            <button onclick="App.ImageTable.copyOcrResultToClipboard()" style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                📋 复制全部文字
                            </button>
                            <button onclick="App.ImageTable.splitOcrResultByLine()" style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                📄 按行分割
                            </button>
                            <button onclick="App.ImageTable.extractOcrKeywords()" style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer; font-size: 12px;">
                                🔍 提取关键词
                            </button>
                        </div>
                        <div style="margin-top: 8px; color: #666; font-size: 11px; line-height: 1.5;">
                            💡 <strong>OCR 能识别：</strong>文字、数字、换行、空格、标点符号<br>
                            ❌ <strong>OCR 不能识别：</strong>表格边框、单元格结构、文字颜色、图片<br>
                            <br>
                            <strong>🎯 推荐方案（无脑使用）：</strong><br>
                            1. 使用本地 OCR 引擎自动识别表格<br>
                            2. 使用在线表格识别工具（如白描网页版、诚华OCR）<br>
                            3. 复制 OCR 文字后在 Excel 中手动整理（当前方案）<br>
                            <br>
                            <strong>⚠️ 关于 PaddleOCR 集成：</strong><br>
                            PaddleOCR 需要 Python + 模型文件（约 1GB+），集成到插件会导致：<br>
                            • 插件体积过大（可能几百 MB）<br>
                            • 首次安装下载时间过长<br>
                            • 可能超出 Office 插件限制<br>
                            当前使用 PaddleOCR，需安装 Python 环境。
                        </div>
                    </div>
                `);

                // 保存 OCR 结果供后续处理
                App.lastOcrResult = ocrResult;
            } else {
                throw new Error('本地 OCR 服务不可用');
            }
        } catch (error) {
            ChatLog.error('IMAGE', '本地 OCR 识别失败:', error);
            throw error;
        }
    }

    // --- 来自 app.js:L7937-L8044 ---
    function parseOcrTextToTable(ocrText) {
        ChatLog.info('IMAGE', '解析 OCR 文本为表格:', ocrText.substring(0, 200));

        if (!ocrText || ocrText.trim().length === 0) {
            return { rows: [], confidence: 0 };
        }

        const lines = ocrText.split('\n').filter(line => line.trim().length > 0);

        if (lines.length === 0) {
            return { rows: [], confidence: 0 };
        }

        // 尝试检测表格结构
        // 策略1：检测是否有明显的分隔符（制表符、多个空格、竖线等）
        const rows = [];
        let maxCols = 0;

        for (const line of lines) {
            // 尝试多种分隔方式
            let cells = [];

            // 方法1：制表符分隔
            if (line.includes('\t')) {
                cells = line.split('\t').map(c => c.trim());
            }
            // 方法2：两个以上空格分隔
            else if (line.match(/\s{2,}/)) {
                cells = line.split(/\s{2,}/).map(c => c.trim()).filter(c => c.length > 0);
            }
            // 方法3：竖线分隔
            else if (line.includes('|')) {
                cells = line.split('|').map(c => c.trim()).filter(c => c.length > 0);
            }
            // 方法4：逗号分隔（如果是CSV风格）
            else if (line.includes(',') && line.split(',').length > 1) {
                cells = line.split(',').map(c => c.trim());
            }
            // 方法5：中文标点分隔（如"、"）
            else if (line.includes('、') && line.split('、').length > 1) {
                cells = line.split('、').map(c => c.trim());
            }
            // 方法6：整行作为一个单元格
            else {
                cells = [line.trim()];
            }

            if (cells.length > 0) {
                rows.push(cells);
                maxCols = Math.max(maxCols, cells.length);
            }
        }

        // 标准化行长度（填充短行）
        const normalizedRows = rows.map(row => {
            while (row.length < maxCols) {
                row.push('');
            }
            return row;
        });

        // 评估置信度
        let confidence = 0.5; // 基础置信度

        // 如果所有行都有相同列数，增加置信度
        if (normalizedRows.length > 1) {
            const allSameLength = normalizedRows.every(row => row.length === maxCols);
            if (allSameLength) {
                confidence += 0.2;
            }

            // 如果第一行看起来像表头（包含常见表头关键词），增加置信度
            const firstRow = normalizedRows[0].join(' ');
            const headerKeywords = ['名称', '编号', '日期', '数量', '价格', '金额', '序号', '项目', '内容', '说明'];
            const hasHeaderKeywords = headerKeywords.some(kw => firstRow.includes(kw));
            if (hasHeaderKeywords) {
                confidence += 0.15;
            }

            // 如果行数较多，增加置信度
            if (normalizedRows.length >= 3) {
                confidence += 0.1;
            }
        }

        ChatLog.info('IMAGE', '解析结果:', normalizedRows.length, '行,', maxCols, '列, 置信度:', confidence);

        // 如果只有1行1列，尝试智能分栏
        if (normalizedRows.length === 1 && maxCols === 1 && normalizedRows[0] && normalizedRows[0][0] && normalizedRows[0][0].length > 50) {
            ChatLog.info('IMAGE', '检测到长文本，尝试智能分栏...');
            const smartRows = trySmartSplit(normalizedRows[0][0]);
            if (smartRows.length > 1) {
                ChatLog.info('IMAGE', '智能分栏成功:', smartRows.length, '行');
                return {
                    rows: smartRows,
                    confidence: 0.6,
                    rawText: ocrText,
                    isSmartSplit: true
                };
            }
        }

        return {
            rows: normalizedRows,
            confidence: Math.min(confidence, 0.95),
            rawText: ocrText
        };
    }

    // --- 来自 app.js:L8047-L8114 ---
    function trySmartSplit(text) {
        const rows = [];

        // 常见的字段关键词模式
        const fieldPatterns = [
            /卡片名称[:：]/,
            /核心定位[:：]/,
            /关联[:：]/,
            /ID[:：]/,
            /记录[:：]/,
            /管理[:：]/,
            /设定[:：]/,
            /规则[:：]/,
            /文化[:：]/,
            /经济[:：]/,
            /政治[:：]/,
            /军事[:：]/,
            /事件[:：]/,
            /人物[:：]/,
            /种族[:：]/,
            /势力[:：]/,
            /世界[:：]/,
            /背景[:：]/,
            /特征[:：]/
        ];

        // 尝试按关键词分割
        let currentRow = [];
        let currentField = '';

        // 首先尝试按"（"分割，因为很多字段格式是"字段名（内容）"
        const parts = text.split(/(?=[（(])/);

        for (let i = 0; i < parts.length; i++) {
            const part = parts[i].trim();
            if (part.length === 0) continue;

            // 检查是否是字段开始
            const isFieldStart = fieldPatterns.some(p => p.test(part));

            if (isFieldStart || currentRow.length === 0) {
                if (currentRow.length > 0) {
                    rows.push(currentRow);
                }
                currentRow = [part];
            } else {
                currentRow.push(part);
            }
        }

        if (currentRow.length > 0) {
            rows.push(currentRow);
        }

        // 如果分栏效果不好（只有1行），返回原始文本作为单行
        if (rows.length <= 1) {
            return [[text]];
        }

        // 标准化列数
        const maxCols = Math.max(...rows.map(r => r.length));
        return rows.map(row => {
            while (row.length < maxCols) {
                row.push('');
            }
            return row;
        });
    }

    // ==================== OCR 结果处理工具函数 ====================

    // --- 来自 app.js:L8119-L8132 ---
    async function copyOcrResultToClipboard() {
        const ocrText = App.lastOcrResult;
        if (!ocrText) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> 没有可复制的 OCR 结果</span>', true) : App.showToast('❌ 没有可复制的 OCR 结果');
            return;
        }
        try {
            await navigator.clipboard.writeText(ocrText);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ OCR 文字已复制到剪贴板</span>', true) : App.showToast('✅ OCR 文字已复制到剪贴板');
        } catch (error) {
            ChatLog.error('IMAGE', '复制失败:', error);
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败</span>', true) : App.showToast('❌ 复制失败');
        }
    }

    // --- 来自 app.js:L8135-L8171 ---
    function splitOcrResultByLine() {
        const ocrText = App.lastOcrResult;
        if (!ocrText) {
            App.showToast('❌ 没有可处理的 OCR 结果');
            return;
        }

        const lines = ocrText.split('\n').filter(line => line.trim().length > 0);

        // 显示分割后的结果
        let tableHtml = '<table style="border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 10px;">';
        lines.forEach((line, index) => {
            tableHtml += `<tr>
                <td style="border: 1px solid #ddd; padding: 6px; width: 40px; text-align: center; background: #f5f5f5;">${index + 1}</td>
                <td style="border: 1px solid #ddd; padding: 6px;">${App.escapeHtml(line.trim())}</td>
            </tr>`;
        });
        tableHtml += '</table>';

        App.addMessage('ai', `
            <div style="background: #f9f9f9; padding: 12px; border-radius: 8px;">
                <div style="font-weight: 500; margin-bottom: 8px;">📄 按行分割结果 (${lines.length} 行)</div>
                ${tableHtml}
                <div style="margin-top: 10px; display: flex; gap: 8px;">
                    <button onclick="App.ImageTable.copyLinesToClipboard()" style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        📋 复制所有行
                    </button>
                    <button onclick="App.ImageTable.insertLinesToExcel()" style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        📊 插入到 Excel
                    </button>
                </div>
            </div>
        `);

        // 保存行数据供后续使用
        App.lastOcrLines = lines;
    }

    // --- 来自 app.js:L8174-L8186 ---
    async function copyLinesToClipboard() {
        const lines = App.lastOcrLines || [];
        if (lines.length === 0) {
            App.showToast('❌ 没有可复制的行');
            return;
        }
        try {
            await navigator.clipboard.writeText(lines.join('\n'));
            App.showToast(`✅ ${lines.length} 行已复制到剪贴板`);
        } catch (error) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败</span>', true) : App.showToast('❌ 复制失败');
        }
    }

    // --- 来自 app.js:L8189-L8215 ---
    async function insertLinesToExcel() {
        const lines = App.lastOcrLines || [];
        if (lines.length === 0) {
            App.showToast('❌ 没有可插入的数据');
            return;
        }

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                // 将行数据转换为二维数组
                const rows = lines.map(line => [line]);
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows));
                const resultObj = JSON.parse(result);

                if (resultObj.success) {
                    App.showToast(`✅ ${lines.length} 行已插入到 Excel`);
                } else {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 插入失败: ' + (resultObj.error || '未知错误') + '</span>', true) : App.showToast('❌ 插入失败: ' + (resultObj.error || '未知错误'));
                }
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> Excel 连接不可用</span>', true) : App.showToast(' Excel 连接不可用');
            }
        } catch (error) {
            ChatLog.error('IMAGE', '插入失败:', error);
            App.showToast('❌ 插入失败');
        }
    }

    // --- 来自 app.js:L8218-L8300 ---
    function extractOcrKeywords() {
        const ocrText = App.lastOcrResult;
        if (!ocrText) {
            App.showToast('❌ 没有可处理的 OCR 结果');
            return;
        }

        // 常见的关键词模式
        const keywordPatterns = [
            /卡片名称[:：]\s*([^\n（(]+)/g,
            /核心定位[:：]\s*([^\n（(]+)/g,
            /关联([^[:：]\n]+)[:：]\s*([^\n]+)/g,
            /([^\s]{2,10})[:：]\s*([^\n]+)/g
        ];

        const keywords = [];

        // 提取匹配的关键词
        keywordPatterns.forEach(pattern => {
            let match;
            while ((match = pattern.exec(ocrText)) !== null) {
                if (match[1] && match[2]) {
                    keywords.push({
                        key: match[1].trim(),
                        value: match[2].trim()
                    });
                }
            }
        });

        // 如果没有提取到关键词，尝试简单分割
        if (keywords.length === 0) {
            const lines = ocrText.split(/[\n。；;]/).filter(l => l.trim().length > 0);
            lines.forEach(line => {
                const parts = line.split(/[:：]/);
                if (parts.length >= 2) {
                    keywords.push({
                        key: parts[0].trim(),
                        value: parts.slice(1).join(':').trim()
                    });
                }
            });
        }

        if (keywords.length === 0) {
            App.addMessage('ai', `
                <div style="background: #fff3e0; padding: 12px; border-radius: 8px;">
                    <div style="font-weight: 500; margin-bottom: 8px;">⚠️ 未提取到关键词</div>
                    <div style="color: #666; font-size: 12px;">
                        未能自动识别关键词模式。您可以使用"按行分割"功能查看所有内容。
                    </div>
                </div>
            `);
            return;
        }

        // 显示提取的关键词
        let tableHtml = '<table style="border-collapse: collapse; width: 100%; font-size: 13px; margin-top: 10px;">';
        keywords.forEach((item, index) => {
            tableHtml += `<tr>
                <td style="border: 1px solid #ddd; padding: 6px; width: 30%; background: #f5f5f5; font-weight: 500;">${App.escapeHtml(item.key)}</td>
                <td style="border: 1px solid #ddd; padding: 6px;">${App.escapeHtml(item.value)}</td>
            </tr>`;
        });
        tableHtml += '</table>';

        App.addMessage('ai', `
            <div style="background: #f9f9f9; padding: 12px; border-radius: 8px;">
                <div style="font-weight: 500; margin-bottom: 8px;">🔍 提取的关键词 (${keywords.length} 个)</div>
                ${tableHtml}
                <div style="margin-top: 10px; display: flex; gap: 8px;">
                    <button onclick="App.ImageTable.copyKeywordsToClipboard()" style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        📋 复制关键词
                    </button>
                    <button onclick="App.ImageTable.insertKeywordsToExcel()" style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer; font-size: 12px;">
                        📊 插入到 Excel
                    </button>
                </div>
            </div>
        `);

        App.lastOcrKeywords = keywords;
    }

    // --- 来自 app.js:L8303-L8316 ---
    async function copyKeywordsToClipboard() {
        const keywords = App.lastOcrKeywords || [];
        if (keywords.length === 0) {
            App.showToast('❌ 没有可复制的关键词');
            return;
        }
        try {
            const text = keywords.map(k => `${k.key}: ${k.value}`).join('\n');
            await navigator.clipboard.writeText(text);
            App.showToast(`✅ ${keywords.length} 个关键词已复制`);
        } catch (error) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">复制失败</span>', true) : App.showToast('❌ 复制失败');
        }
    }

    // --- 来自 app.js:L8319-L8344 ---
    async function insertKeywordsToExcel() {
        const keywords = App.lastOcrKeywords || [];
        if (keywords.length === 0) {
            App.showToast('❌ 没有可插入的数据');
            return;
        }

        try {
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const rows = keywords.map(k => [k.key, k.value]);
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows));
                const resultObj = JSON.parse(result);

                if (resultObj.success) {
                    App.showToast(`✅ ${keywords.length} 个关键词已插入到 Excel`);
                } else {
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 插入失败: ' + (resultObj.error || '未知错误') + '</span>', true) : App.showToast('❌ 插入失败: ' + (resultObj.error || '未知错误'));
                }
            } else {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;"> Excel 连接不可用</span>', true) : App.showToast(' Excel 连接不可用');
            }
        } catch (error) {
            ChatLog.error('IMAGE', '插入失败:', error);
            App.showToast('❌ 插入失败');
        }
    }

    /**
     * 清理表格中的默认列名（如"列1"、"列2"、"Column1"等）
     * 如果第一行是默认占位符列名，则删除该行或将表格转换为无表头格式
     */
    function cleanTableHeaders(rows) {
        if (!rows || rows.length === 0) return rows;

        // 默认列名的正则表达式（支持中英文）
        const defaultHeaderPatterns = [
            /^列\d+$/,           // 列1, 列2, 列10
            /^Column\d+$/i,      // Column1, Column2
            /^Col\d+$/i,         // Col1, Col2
            /^表头\d+$/,          // 表头1, 表头2
            /^Header\d+$/i,      // Header1, Header2
            /^字段\d+$/,          // 字段1, 字段2
            /^Field\d+$/i        // Field1, Field2
        ];

        const isDefaultHeader = (text) => {
            if (!text || typeof text !== 'string') return false;
            const trimmed = text.trim();
            return defaultHeaderPatterns.some(pattern => pattern.test(trimmed));
        };

        // 检查第一行是否全部是默认列名
        const firstRow = rows[0];
        if (firstRow && Array.isArray(firstRow)) {
            const allDefaultHeaders = firstRow.every(cell => isDefaultHeader(cell));

            if (allDefaultHeaders) {
                ChatLog.info('IMAGE', '⚠️ 检测到默认列名（列1、列2等），自动删除表头行');
                // 删除第一行（默认列名行），返回剩余的数据行
                return rows.slice(1);
            }

            // 如果只有部分是默认列名，记录警告但不处理
            const defaultCount = firstRow.filter(cell => isDefaultHeader(cell)).length;
            if (defaultCount > 0) {
                ChatLog.warn('IMAGE', `⚠️ 第一行包含 ${defaultCount} 个默认列名，请检查数据`);
            }
        }

        return rows;
    }

    // --- 来自 app.js:L8347-L8413 ---
    async function showImageTableResult(tableData, fileName) {
        ChatLog.info('IMAGE', '显示图片表格识别结果:', tableData);

        const rows = cleanTableHeaders(tableData.rows);
        const confidence = tableData.confidence || 0;

        // 构建表格 HTML
        let tableHtml = '<div style="margin: 10px 0; max-height: 400px; overflow: auto;">';
        tableHtml += '<table style="border-collapse: collapse; width: 100%; font-size: 13px;">';

        rows.forEach((row, rowIndex) => {
            tableHtml += '<tr>';
            row.forEach((cell, cellIndex) => {
                const isHeader = rowIndex === 0;
                const tag = isHeader ? 'th' : 'td';
                const bgColor = isHeader ? '#e3f2fd' : 'white';
                tableHtml += `<${tag} style="border: 1px solid #ddd; padding: 8px; background: ${bgColor}; text-align: left;">${App.escapeHtml(cell || '')}</${tag}>`;
            });
            tableHtml += '</tr>';
        });

        tableHtml += '</table></div>';

        // 置信度显示
        const confidenceText = confidence > 0 ? ` (置信度: ${(confidence * 100).toFixed(1)}%)` : '';

        // 完整性警告
        const completenessWarning = tableData.note && tableData.note.includes('不完整')
            ? `<div style="color: #ff9800; font-size: 12px; margin-bottom: 8px; padding: 6px; background: #fff3e0; border-radius: 4px;">⚠️ ${tableData.note}</div>`
            : '';

        // 将表格数据存储在全局变量中，供按钮使用
        App.lastRecognizedTable = rows;

        // 添加操作按钮
        const actionsHtml = `
            <div style="margin-top: 10px; display: flex; gap: 8px; flex-wrap: wrap;">
                <button onclick="App.ImageTable.insertLastTableToExcel()"
                        style="padding: 6px 12px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer;">
                    📊 插入到 Excel
                </button>
                <button onclick="App.ImageTable.copyLastTableToClipboard()"
                        style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">
                    📋 复制表格
                </button>
                <button onclick="App.ImageTable.editLastTableData()"
                        style="padding: 6px 12px; background: #f0f0f0; border: 1px solid #ddd; border-radius: 4px; cursor: pointer;">
                    ✏️ 编辑数据
                </button>
            </div>
        `;

        // 显示在聊天窗口
        const messageContent = `
            <div style="background: #f9f9f9; padding: 12px; border-radius: 8px; border: 1px solid #e0e0e0;">
                <div style="font-weight: 500; margin-bottom: 8px;">📊 识别到表格 ${confidenceText}</div>
                <div style="color: #666; font-size: 12px; margin-bottom: 8px;">${rows.length} 行 × ${rows[0]?.length || 0} 列</div>
                ${completenessWarning}
                ${tableHtml}
                ${actionsHtml}
            </div>
        `;

        // 使用 addMessage 而不是 updateLastMessage，确保 HTML 正确渲染
        App.addMessage('ai', messageContent);
        App.showToast(`✅ 成功识别表格: ${rows.length} 行`);
    }

    // --- 来自 app.js:L8416-L8872 ---
    async function analyzeImageTableWithAI(base64Image, fileName, messageId = null, updateStep = null) {
        ChatLog.info('IMAGE', '========== analyzeImageTableWithAI 开始 ==========');
        ChatLog.info('IMAGE', '使用 AI 分析图片表格:', fileName);
        ChatLog.info('IMAGE', '当前选择的模型 (currentChatModel):', App.currentChatModel);

        const config = App.config;
        ChatLog.info('IMAGE', '设置中的模型 (config.model):', config.model);

        let prompt = `你是一个专业的表格数据提取工具。请仔细分析图片中的表格内容，提取所有数据。

【严格要求】
1. 只输出一个JSON对象，不要任何解释、说明、前言、后记、markdown标记
2. JSON格式：{"rows":[["真实的列标题1","真实的列标题2"],["实际数据1","实际数据2"]]}
3. 第一行必须是图片中真实的表头/列标题（如"姓名"、"日期"、"金额"等业务字段名），绝对禁止使用"列1"、"列2"、"Column1"、"Column2"、"表头1"、"表头2"这类占位符
4. 如果图片中没有明确的表头行，则第一行直接就是数据，不需要生成虚拟表头
5. 保持原始文字，不要翻译、修改或猜测内容
6. 确保每行的列数一致

【输出示例】
正确的：{"rows":[["姓名","部门","工资"],["张三","技术部",8000],["李四","销售部",7000]]}
错误的：{"rows":[["列1","列2","列3"],["张三","技术部",8000]]}

直接开始输出JSON：`;

        // 添加完整性检查
        const checkCompleteness = (tableData) => {
            if (!tableData.rows || tableData.rows.length === 0) return false;

            // 检查最后一行是否完整（所有列都有值）
            const lastRow = tableData.rows[tableData.rows.length - 1];
            const firstRow = tableData.rows[0];

            // 如果最后一行明显不完整（列数少于第一行的一半），认为是不完整的
            if (lastRow.length < firstRow.length / 2) {
                return false;
            }

            // 检查最后一行的最后一个单元格是否为空或明显不完整
            const lastCell = lastRow[lastRow.length - 1];
            if (!lastCell || lastCell.length < 2) {
                return false;
            }

            return true;
        };

        // 添加重试机制
        let attempts = 0;
        const maxAttempts = 4;
        let lastError = null;

        while (attempts < maxAttempts) {
            attempts++;
            ChatLog.info('IMAGE', `第 ${attempts} 次尝试识别表格...`);

            try {
                // 调用 AI 分析图片
                const response = await callAIWithImage(prompt, base64Image, fileName);
                ChatLog.info('IMAGE', 'AI 完整响应:', response);
                ChatLog.info('IMAGE', 'AI 响应前500字符:', response.substring(0, 500));

                // 检测空响应或无效响应
                if (!response || response.trim().length === 0) {
                    ChatLog.warn('IMAGE', '⚠️ AI 返回空响应');

                    // 如果是第一次尝试，尝试切换到 llava:latest
                    if (attempts === 1 && App.currentChatModel.toLowerCase().includes('qwen')) {
                        ChatLog.info('IMAGE', '🔄 qwen 视觉模型返回空响应，尝试切换到 llava...');

                        // 临时切换到 llava
                        const modelSelect = document.getElementById('chat-model-select');
                        if (modelSelect) {
                            for (let i = 0; i < modelSelect.options.length; i++) {
                                if (modelSelect.options[i].value.toLowerCase().includes('llava')) {
                                    App.currentChatModel = modelSelect.options[i].value;
                                    modelSelect.value = App.currentChatModel;
                                    const headerModelBadge = document.querySelector('.model-badge');
                                    if (headerModelBadge) {
                                        headerModelBadge.textContent = App.currentChatModel + ' (视觉)';
                                    }
                                    ChatLog.info('IMAGE', '✅ 已切换到:', App.currentChatModel);

                                    // 实时显示第二次模型切换
                                    const messageDiv = document.getElementById(messageId);
                                    if (messageDiv) {
                                        const stepsContainer = messageDiv.querySelector('div[style*="padding-left: 26px"]');
                                        if (stepsContainer) {
                                            const retryStep = document.createElement('div');
                                            retryStep.style.color = '#ff9800';
                                            retryStep.style.marginTop = '4px';
                                            retryStep.innerHTML = `🔄 qwen 返回空响应，已切换到: ${App.currentChatModel}`;
                                            stepsContainer.appendChild(retryStep);
                                        }
                                    }
                                    break;
                                }
                            }
                        }

                        // 继续下一次循环重试（attempts 会正常递增）
                        continue;
                    }

                    // 如果已经重试过，返回错误
                    throw new Error('EMPTY_RESPONSE');
                }

                // 检测垃圾/重复响应（模型输出大量废话而非JSON）
                const isGarbageResponse = (text) => {
                    if (!text || text.length < 50) return false;
                    const trimmed = text.trim();

                    // 检测1: 超长响应但不含任何JSON特征
                    if (trimmed.length > 3000 && !trimmed.includes('"rows"') && !trimmed.includes('[[') && !trimmed.includes('{')) {
                        ChatLog.warn('IMAGE', '⚠️ 响应过长且无JSON特征');
                        return true;
                    }

                    // 检测2: 高重复率（取前200字符，看它在全文中出现次数）
                    const sample = trimmed.substring(0, 200);
                    if (sample.length >= 50) {
                        const occurrences = (trimmed.match(new RegExp(sample.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g')) || []).length;
                        if (occurrences >= 3) {
                            ChatLog.warn('IMAGE', `⚠️ 检测到高重复文本 (重复${occurrences}次)`);
                            return true;
                        }
                    }

                    // 检测3: 连续相同段落重复超过3次
                    const chunkSize = 100;
                    let repeatCount = 0;
                    for (let i = 0; i < Math.min(trimmed.length - chunkSize * 2, 1000); i += chunkSize) {
                        const chunk = trimmed.substring(i, i + chunkSize);
                        const nextChunk = trimmed.substring(i + chunkSize, i + chunkSize * 2);
                        if (chunk === nextChunk) repeatCount++;
                        if (repeatCount >= 3) {
                            ChatLog.warn('IMAGE', '⚠️ 检测到连续重复段落');
                            return true;
                        }
                    }

                    // 检测4: 常见垃圾模式 - 大量"用户可能提供"等模板文字
                    const junkPatterns = [/用户可能提供[^，。]{0,20}(?:表格|图片)[^，。]{0,20}/g];
                    for (const p of junkPatterns) {
                        const matches = trimmed.match(p) || [];
                        if (matches.length >= 5) {
                            ChatLog.warn('IMAGE', `⚠️ 检测到垃圾模式 (${matches.length}次匹配)`);
                            return true;
                        }
                    }

                    return false;
                };

                if (isGarbageResponse(response)) {
                    ChatLog.warn('IMAGE', '⚠️ AI 返回垃圾响应，触发重试');

                    if (attempts === 1 && App.currentChatModel.toLowerCase().includes('qwen')) {
                        ChatLog.info('IMAGE', '🔄 qwen 模型返回垃圾响应，尝试切换到 llava...');

                        const modelSelect = document.getElementById('chat-model-select');
                        if (modelSelect) {
                            for (let i = 0; i < modelSelect.options.length; i++) {
                                if (modelSelect.options[i].value.toLowerCase().includes('llava')) {
                                    App.currentChatModel = modelSelect.options[i].value;
                                    modelSelect.value = App.currentChatModel;
                                    const headerModelBadge = document.querySelector('.model-badge');
                                    if (headerModelBadge) {
                                        headerModelBadge.textContent = App.currentChatModel + ' (视觉)';
                                    }

                                    const messageDiv = document.getElementById(messageId);
                                    if (messageDiv) {
                                        const stepsContainer = messageDiv.querySelector('div[style*="padding-left: 26px"]');
                                        if (stepsContainer) {
                                            const retryStep = document.createElement('div');
                                            retryStep.style.color = '#ff9800';
                                            retryStep.style.marginTop = '4px';
                                            retryStep.innerHTML = `🔄 qwen 返回无效内容，已切换到: ${App.currentChatModel}`;
                                            stepsContainer.appendChild(retryStep);
                                        }
                                    }
                                    break;
                                }
                            }
                        }
                        continue;
                    }

                    throw new Error('GARBAGE_RESPONSE');
                }

                // 更新步骤到第3步
                if (updateStep) {
                    updateStep(3);
                }

                // ===== AI 输出的 JSON 常见错误修复 =====
                const fixAIJson = (text) => {
                    if (!text) return text;
                    let fixed = text;

                    // 1. 中文引号 → 英文引号
                    fixed = fixed.replace(/[""]/g, '"');

                    // 2. 中文单引号 → 英文单引号
                    fixed = fixed.replace(/['']/g, "'");

                    // 3. 中文括号 → 英文括号
                    fixed = fixed.replace(/（/g, '(').replace(/）/g, ')');

                    // 4. 常见的尾部截断修复：找到最后一个 ] 或 } 后补全
                    const openBrackets = (fixed.match(/\[/g) || []).length;
                    const closeBrackets = (fixed.match(/\]/g) || []).length;
                    const openBraces = (fixed.match(/\{/g) || []).length;
                    const closeBraces = (fixed.match(/\}/g) || []).length;

                    if (openBrackets > closeBrackets) {
                        fixed += ']'.repeat(openBrackets - closeBrackets);
                        ChatLog.info('IMAGE', `🔧 补全 ${openBrackets - closeBrackets} 个 ]`);
                    }
                    if (openBraces > closeBraces) {
                        fixed += '}'.repeat(openBraces - closeBraces);
                        ChatLog.info('IMAGE', `🔧 补全 ${openBraces - closeBraces} 个 }`);
                    }

                    // 5. 移除末尾可能的垃圾字符（非JSON部分）
                    fixed = fixed.replace(/[\s,]*[^}\]]*$/, (match) => {
                        if (/^[\s,]*$/.test(match)) return match;
                        return '';
                    });

                    // 6. 处理末尾多余的内容（如 "）]" 后面还有文字）
                    const lastBracket = Math.max(fixed.lastIndexOf(']'), fixed.lastIndexOf('}'));
                    if (lastBracket > 0 && lastBracket < fixed.length - 1) {
                        const afterLast = fixed.substring(lastBracket + 1).trim();
                        if (afterLast.length > 0 && !/^[,\s\]}]*$/.test(afterLast)) {
                            fixed = fixed.substring(0, lastBracket + 1);
                            ChatLog.info('IMAGE', '🔧 截断末尾多余内容');
                        }
                    }

                    return fixed.trim();
                };

                const responseForParse = fixAIJson(response);
                if (responseForParse !== response) {
                    ChatLog.info('IMAGE', '🔧 JSON 已修复，原始长度:', response.length, '修复后:', responseForParse.length);
                }

                // 尝试解析 JSON 结果 - 支持对象格式 {"rows": [...]} 或数组格式 [[...]]
                let tableData = null;

                // 策略1: 优先查找包含 "rows" 键的JSON对象（最精确）
                const rowsMatch = responseForParse.match(/\{[^{}]*"rows"\s*:\s*\[[\s\S]*?\]\s*\}/);
                if (rowsMatch) {
                    try {
                        const parsed = JSON.parse(rowsMatch[0]);
                        if (parsed.rows && Array.isArray(parsed.rows)) {
                            tableData = parsed;
                            ChatLog.info('IMAGE', '解析到rows格式:', tableData);
                        }
                    } catch (e) {
                        ChatLog.info('IMAGE', 'rows格式解析失败:', e.message);
                    }
                }

                // 策略2: 尝试匹配完整的 { "rows": [...] } 对象（允许嵌套）
                if (!tableData) {
                    const deepRowsMatch = responseForParse.match(/(\{[^]*?"rows"\s*:\s*\[(?:[^[\]]*(?:\[[^[\]]*\][^[\]]*)*)\][^}]*?\})/);
                    if (deepRowsMatch) {
                        try {
                            const cleaned = deepRowsMatch[0].replace(/,[\s]*$/, '');
                            const parsed = JSON.parse(cleaned);
                            if (parsed.rows && Array.isArray(parsed.rows)) {
                                tableData = parsed;
                                ChatLog.info('IMAGE', '深度解析到rows格式:', tableData);
                            }
                        } catch (e) {
                            ChatLog.info('IMAGE', '深度rows格式解析失败');
                        }
                    }
                }

                // 策略3: 匹配任何可解析的JSON对象
                if (!tableData) {
                    const objectMatches = responseForParse.matchAll(/\{[^{}]*(?:\{[^{}]*\}[^{}]*)*\}/g);
                    for (const m of objectMatches) {
                        try {
                            const parsed = JSON.parse(m[0]);
                            if ((parsed.rows || parsed.data || parsed.table) && Array.isArray(parsed.rows || parsed.data || parsed.table)) {
                                tableData = parsed;
                                if (!tableData.rows) tableData.rows = parsed.data || parsed.table;
                                ChatLog.info('IMAGE', '从候选对象中解析成功:', tableData);
                                break;
                            }
                        } catch (e) { (window.ChatLog || console).debug?.('IMAGE', '[image-table] 候选对象非表格 JSON，尝试下一策略:', e); }
                    }
                }

                // 策略4: 尝试二维数组格式 [[...],[...]]
                if (!tableData) {
                    const arrayMatch = responseForParse.match(/\[\s*\[.*?\]\s*(?:,\s*\[.*?\]\s*)*\]/);
                    if (arrayMatch) {
                        try {
                            const parsed = JSON.parse(arrayMatch[0]);
                            if (Array.isArray(parsed) && parsed.length > 0 && Array.isArray(parsed[0])) {
                                tableData = { rows: parsed };
                                ChatLog.info('IMAGE', '解析到数组格式:', tableData);
                            }
                        } catch (e) {
                            ChatLog.info('IMAGE', '数组格式解析失败');
                        }
                    }
                }

                // 如果标准 JSON 解析失败，显示原始文本供用户参考
                if (!tableData && response.length > 20) {
                    ChatLog.info('IMAGE', '⚠️ 无法解析为标准表格格式，显示原始文本');

                    // 直接返回标记，让上层显示原始文本
                    return { isNonVisionModel: true, response: response, isRawText: true };
                }

                if (tableData) {
                    // 检查是否有错误
                    if (tableData.error) {
                        ChatLog.info('IMAGE', 'AI 返回错误:', tableData.error);
                        App.updateLastMessage('ai', '⚠️ ' + tableData.error);
                        return;
                    }

                    // 检查是否识别到表格
                    if (tableData.rows && tableData.rows.length > 0) {
                        ChatLog.info('IMAGE', '识别到表格:', tableData.rows.length, '行');
                        ChatLog.info('IMAGE', '第一行内容:', tableData.rows[0]);

                        // 验证表格内容是否合理（检查是否包含示例数据）
                        const firstRowText = tableData.rows[0].join('');
                        const allText = tableData.rows.map(r => r.join('')).join('');

                        // 检测常见的示例数据模式
                        const samplePatterns = [
                            { pattern: /张三|李四|王五/, name: '常见人名' },
                            { pattern: /姓名.*年龄.*性别/, name: '人员信息表头' },
                            { pattern: /产品名称/, name: '产品名称' },
                            { pattern: /苹果|香蕉|橘子|橙子|葡萄|西瓜/, name: '水果名称' },
                            { pattern: /^产品[ABC12]/, name: '产品编号示例' },
                            { pattern: /商品/, name: '商品' },
                            { pattern: /单价/, name: '单价' }
                        ];

                        const matchedPatterns = samplePatterns.filter(p => p.pattern.test(allText));
                        const hasSamplePattern = matchedPatterns.length > 0;
                        // 放宽表格大小限制，只要能匹配到示例模式就认为是示例数据
                        const isSampleData = hasSamplePattern;

                        ChatLog.info('IMAGE', '=== 示例数据检测 ===');
                        ChatLog.info('IMAGE', '表格大小:', tableData.rows.length, '行 ×', (tableData.rows[0] ? tableData.rows[0].length : 0), '列');
                        ChatLog.info('IMAGE', '匹配的模式:', matchedPatterns.map(p => p.name));
                        ChatLog.info('IMAGE', '是否包含示例模式:', hasSamplePattern);
                        ChatLog.info('IMAGE', '是否示例数据:', isSampleData);
                        ChatLog.info('IMAGE', '检测到的文本:', allText.substring(0, 100));
                        ChatLog.info('IMAGE', '当前尝试次数:', attempts, '/', maxAttempts);
                        ChatLog.info('IMAGE', '第一行内容:', tableData.rows[0]);

                        // 如果检测到示例数据，直接切换到本地 OCR（不再重试 AI）
                        if (isSampleData) {
                            ChatLog.info('IMAGE', '检测到示例数据，可能是非视觉模型在编造数据，切换到本地 OCR');
                            // 返回标记，让上层使用本地 OCR
                            return { isNonVisionModel: true, response: '模型返回示例数据而非真实内容' };
                        }

                        // 检查识别结果是否完整
                        const isComplete = checkCompleteness(tableData);
                        ChatLog.info('IMAGE', '识别结果是否完整:', isComplete);

                        if (!isComplete) {
                            ChatLog.warn('IMAGE', '识别结果可能不完整，最后一行被截断');
                            tableData.note = '注意：识别结果可能不完整，建议检查原始图片';
                        }

                        // 显示识别结果
                        ChatLog.info('IMAGE', '显示识别结果...');
                        await showImageTableResult(tableData, fileName);
                        return;
                    } else {
                        ChatLog.info('IMAGE', '未识别到表格数据');
                        App.updateLastMessage('ai', '⚠️ 未能在图片中识别到表格');
                        return;
                    }
                } else {
                    ChatLog.info('IMAGE', '未找到 JSON 格式的响应');

                    // 检查是否是模型不支持图片识别的情况
                    const noImageKeywords = [
                        '没有提供图片',
                        '未提供图片',
                        '无法识别图片',
                        '不支持图片',
                        '没有图像',
                        '未提供图像',
                        '无法查看图片',
                        '无法查看图像'
                    ];

                    const isNonVisionModel = noImageKeywords.some(keyword =>
                        response.toLowerCase().includes(keyword)
                    );

                    if (isNonVisionModel) {
                        // 返回标记对象，让上层函数决定使用本地 OCR
                        ChatLog.info('IMAGE', '检测到非视觉模型，返回标记让上层处理');
                        return { isNonVisionModel: true, response: response };
                    } else {
                        // 如果解析失败，显示原始结果
                        ChatLog.info('IMAGE', 'AI 原始响应:', response);
                        const displayText = response && response.trim() ? response.substring(0, 1000) : '（空响应）';
                        App.updateLastMessage('ai', `
                            <div style="background: #fff3e0; padding: 12px; border-radius: 8px; border-left: 4px solid #ff9800;">
                                <div style="font-weight: 500; margin-bottom: 8px;">⚠️ AI 分析结果（未识别到表格格式）</div>
                                <pre style="background: white; padding: 10px; border-radius: 4px; overflow: auto; max-height: 300px; font-size: 12px; white-space: pre-wrap; border: 1px solid #e0e0e0;">${App.escapeHtml(displayText)}${response.length > 1000 ? '\n... (已截断)' : ''}</pre>
                                <div style="margin-top: 8px; color: #666; font-size: 11px;">
                                    💡 提示：AI 返回了内容，但不符合表格格式。您可以复制上面的文字手动整理，或使用本地 OCR。
                                </div>
                            </div>
                        `);
                    }
                    return;
                }

            } catch (error) {
                ChatLog.error('IMAGE', `第 ${attempts} 次尝试失败:`, error);
                ChatLog.error('IMAGE', '错误详情:', error.stack);
                lastError = error;

                // 如果是没有视觉模型的错误，直接返回标记让上层处理
                if (error.message === 'NO_VISION_MODEL') {
                    ChatLog.info('IMAGE', '检测到无视觉模型错误，返回标记让上层切换到本地 OCR');
                    return { isNonVisionModel: true, response: '未找到可用的视觉模型' };
                }

                // 如果是空响应错误，返回标记让上层切换到本地 OCR
                if (error.message === 'EMPTY_RESPONSE') {
                    ChatLog.info('IMAGE', '检测到空响应错误，返回标记让上层切换到本地 OCR');
                    return { isNonVisionModel: true, response: '视觉模型返回空响应' };
                }

                // 如果是垃圾响应错误，返回标记让上层切换到本地 OCR
                if (error.message === 'GARBAGE_RESPONSE') {
                    ChatLog.info('IMAGE', '检测到垃圾响应错误，返回标记让上层切换到本地 OCR');
                    return { isNonVisionModel: true, response: '视觉模型返回无效内容（重复/废话）' };
                }
                if (attempts >= maxAttempts) {
                    App.updateLastMessage('ai', '❌ 表格识别失败: ' + error.message);
                    return;
                }
            }
        }

        // 如果所有尝试都失败了
        ChatLog.info('IMAGE', '所有尝试都失败了');
        App.updateLastMessage('ai', '❌ 表格识别失败: 无法从图片中识别出有效的表格数据');
    }

    // --- 来自 app.js:L9280-L9300 ---
    async function handlePastedDataUrl(dataUrl) {
        try {
            ChatLog.info('IMAGE', '处理 Data URL 图片');
            // 从 Data URL 创建 Blob
            const byteString = atob(dataUrl.split(',')[1]);
            const mimeString = dataUrl.split(',')[0].split(':')[1].split(';')[0];
            const ab = new ArrayBuffer(byteString.length);
            const ia = new Uint8Array(ab);
            for (let i = 0; i < byteString.length; i++) {
                ia[i] = byteString.charCodeAt(i);
            }
            const blob = new Blob([ab], { type: mimeString });
            const file = new File([blob], `粘贴图片_${Date.now()}.png`, {
                type: mimeString
            });
            await handlePastedFile(file);
        } catch (error) {
            ChatLog.error('IMAGE', '处理 Data URL 图片失败:', error);
            App.showToast('粘贴图片失败');
        }
    }

    // --- 来自 app.js:L9303-L9342 ---
    async function handlePastedFile(file) {
        ChatLog.info('IMAGE', '处理粘贴的文件:', file.name, file.type, file.size);
        App.selectedFile = file;
        App.selectedFiles = []; // 清空多文件列表

        // 【关键修复】全局备份粘贴的文件
        window.__backupSelectedFile = file;
        window.__backupSelectedFiles = [];
        ChatLog.info('IMAGE', '💾 [文件备份] 粘贴文件已保存到全局:', file.name);

        // 显示附件预览区域（与文件上传统一）
        const preview = document.getElementById('attachment-preview');

        // 隐藏旧版单文件 header 元素（已改用 attachment-list）
        const legacyHeader = preview.querySelector('.attachment-header');
        if (legacyHeader) legacyHeader.style.display = 'none';

        // 渲染统一附件列表（含删除按钮）
        if (typeof App.renderAttachmentList === 'function') {
            App.renderAttachmentList([file]);
        }

        // 如果是图片,显示缩略图
        if (file.type.startsWith('image/')) {
            const thumbnail = document.getElementById('attachment-thumbnail');
            const thumbnailImg = document.getElementById('thumbnail-img');
            if (thumbnailImg) {
                try { URL.revokeObjectURL(thumbnailImg.src); } catch(e) { (window.ChatLog || console).warn?.('UI', '[image-table] 撤销缩略图 objectURL 失败:', e); }
                thumbnailImg.src = URL.createObjectURL(file);
            }
            if (thumbnail) thumbnail.style.display = 'flex';
        } else {
            document.getElementById('attachment-thumbnail').style.display = 'none';
        }

        preview.style.display = 'flex';

        // 【关键】将File对象存储到DOM元素上（粘贴文件模式）
        preview.dataset.fileCount = '1';
        preview._storedFiles = [file];
        ChatLog.info('IMAGE', '💾 [handlePastedFile] 粘贴文件已存储到DOM:', file.name);

        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已粘贴图片: ' + file.name + '</span>', true) : App.showToast('已粘贴图片: ' + file.name);

        // 图片不再自动分析，分析过程移到发送后
        document.getElementById('attachment-actions').innerHTML =
            '<span style="color: #666; font-size: 12px;">💡 点击"发送"按钮开始分析图片</span>';
    }

    // 编辑表格数据（内部辅助函数）
    function editTableData(rows) {
        ChatLog.info('IMAGE', '编辑表格数据:', rows);

        // 创建编辑对话框
        const dialog = document.createElement('div');
        dialog.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0,0,0,0.5);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 10000;
        `;

        // 构建编辑表格
        let tableHtml = '<table style="border-collapse: collapse; width: 100%;">';
        rows.forEach((row, rowIndex) => {
            tableHtml += '<tr>';
            row.forEach((cell, cellIndex) => {
                tableHtml += `<td style="border: 1px solid #ddd; padding: 4px;">
                    <input type="text" value="${App.escapeHtml(cell || '')}"
                           data-row="${rowIndex}" data-col="${cellIndex}"
                           style="width: 100%; border: none; padding: 4px; font-size: 13px;">
                </td>`;
            });
            tableHtml += '</tr>';
        });
        tableHtml += '</table>';

        dialog.innerHTML = `
            <div style="background: white; padding: 20px; border-radius: 8px; max-width: 90%; max-height: 80%; overflow: auto;">
                <h3 style="margin-top: 0;">编辑表格数据</h3>
                <div style="margin-bottom: 15px;">${tableHtml}</div>
                <div style="display: flex; gap: 10px; justify-content: flex-end;">
                    <button id="cancelEdit" style="padding: 8px 16px; border: 1px solid #ddd; background: #f0f0f0; border-radius: 4px; cursor: pointer;">取消</button>
                    <button id="saveEdit" style="padding: 8px 16px; background: #4a90d9; color: white; border: none; border-radius: 4px; cursor: pointer;">保存</button>
                </div>
            </div>
        `;

        document.body.appendChild(dialog);

        // 绑定事件
        dialog.querySelector('#cancelEdit').onclick = () => dialog.remove();
        dialog.querySelector('#saveEdit').onclick = () => {
            const inputs = dialog.querySelectorAll('input');
            const newRows = [];

            inputs.forEach(input => {
                const rowIndex = parseInt(input.dataset.row);
                const colIndex = parseInt(input.dataset.col);

                if (!newRows[rowIndex]) {
                    newRows[rowIndex] = [];
                }
                newRows[rowIndex][colIndex] = input.value;
            });

            dialog.remove();

            // 显示更新后的表格
            showImageTableResult({ rows: newRows }, '编辑后的表格');
            App.showToast('✅ 表格数据已更新');
        };
    }

    // ==================== 模块导出 ====================
    
    // 暴露公共API到 App.ImageTable 命名空间
    App.ImageTable = {
        // 图片表格识别
        processImageTableRecognition: processImageTableRecognition,
        processImageTableWithLocalOCR: processImageTableWithLocalOCR,

        // OCR 文本解析
        parseOcrTextToTable: parseOcrTextToTable,
        trySmartSplit: trySmartSplit,

        // 结果展示
        showImageTableResult: showImageTableResult,
        showTablePreview: showTablePreview,

        // AI 分析
        analyzeImageTableWithAI: analyzeImageTableWithAI,

        // OCR 操作
        copyOcrResultToClipboard: copyOcrResultToClipboard,
        splitOcrResultByLine: splitOcrResultByLine,
        copyLinesToClipboard: copyLinesToClipboard,
        insertLinesToExcel: insertLinesToExcel,
        extractOcrKeywords: extractOcrKeywords,
        copyKeywordsToClipboard: copyKeywordsToClipboard,
        insertKeywordsToExcel: insertKeywordsToExcel,

        // 表格操作
        insertLastTableToExcel: insertLastTableToExcel,
        copyLastTableToClipboard: copyLastTableToClipboard,
        editLastTableData: editLastTableData,
        insertTableToExcel: insertTableToExcel,
        copyTableToClipboard: copyTableToClipboard,
        downloadTableAsCsv: downloadTableAsCsv,

        // 粘贴处理
        handlePaste: handlePaste,
        handlePastedTable: handlePastedTable,
        handlePastedTableText: handlePastedTableText,
        handlePastedDataUrl: handlePastedDataUrl,
        handlePastedFile: handlePastedFile
    };

})();
