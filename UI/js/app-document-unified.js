/**
 * ============================================================
 * app-document-unified.js — Document Unified Processor
 * ============================================================
 * 功能：文档统一处理模块，统一解析 Word/PDF/TXT 等文档格式并提取结构化文本。
 *
 * 核心类/函数：
 *   - parseDocument() — 解析文档
 *   - extractText() — 提取文本
 *   - getDocumentMetadata() — 获取文档元数据
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    // ============================================================
    // 支持的文档格式配置
    // ============================================================
    var SUPPORTED_DOC_FORMATS = {
        pdf:   { ext: ['.pdf'],           name: 'PDF 文档',        icon: '📄', maxSize: 200 },
        docx:  { ext: ['.docx'],          name: 'Word 文档',       icon: '📝', maxSize: 50 },
        doc:   { ext: ['.doc'],           name: 'Word 97-03',     icon: '📝', maxSize: 50 },
        xlsx:  { ext: ['.xlsx'],          name: 'Excel 工作簿',    icon: '📊', maxSize: 50 },
        xls:   { ext: ['.xls'],           name: 'Excel 97-03',    icon: '📊', maxSize: 50 }
    };

    // 格式到类型名的映射（用于文件名检测）
    var EXT_TYPE_MAP = {};
    Object.keys(SUPPORTED_DOC_FORMATS).forEach(function(type) {
        SUPPORTED_DOC_FORMATS[type].ext.forEach(function(e) {
            EXT_TYPE_MAP[e.toLowerCase()] = type;
        });
    });

    // ============================================================
    // DocumentProcessor - 统一文档处理器
    // ============================================================
    var DocumentProcessor = {

        // 状态管理
        _state: {
            currentFile: null,
            parseResult: null,      // 解析结果 JSON
            chunks: [],             // 分块结果
            isProcessing: false,
            abortController: null
        },

        // ===== 公共 API =====

        /**
         * 检测是否为支持的文档格式
         * @param {File} file - 文件对象
         * @returns {{ isDoc: boolean, type: string, config: object }}
         */
        isDocumentFile: function(file) {
            if (!file || !file.name) return { isDoc: false, type: null, config: null };
            var ext = this._getFileExt(file.name);
            var type = EXT_TYPE_MAP[ext];
            if (!type) return { isDoc: false, type: null, config: null };
            return {
                isDoc: true,
                type: type,
                config: SUPPORTED_DOC_FORMATS[type]
            };
        },

        /**
         * 主入口：处理文档文件（解析 + 分块）
         * @param {File} file - 文件对象
         * @param {object} options - 可选参数 { onProgress, chunkSize, overlap, mode }
         * @returns {Promise<object>} 处理结果
         */
        processDocument: async function(file, options) {
            options = options || {};
            var self = this;
            var detection = this.isDocumentFile(file);

            if (!detection.isDoc) {
                throw new Error('不支持的文档格式: ' + (file ? file.name : 'null'));
            }

            // 大小检查
            var maxSizeMB = detection.config.maxSize;
            if (file.size > maxSizeMB * 1024 * 1024) {
                throw new Error(
                    detection.config.name + ' 文件过大 (' +
                    (file.size / 1024 / 1024).toFixed(1) + 'MB > ' + maxSizeMB + 'MB)'
                );
            }

            this._state.currentFile = file;
            this._state.isProcessing = true;

            try {
                // Step 1: 读取文件为 Base64
                if (options.onProgress) options.onProgress('reading', '📄 本地读取文件中');
                var base64Data = await this._readFileAsBase64(file);

                // Step 2: 调用后端解析
                if (options.onProgress) options.onProgress('parsing', '📄 本地解析' + detection.config.name + '中');
                
                var parseResult;
                var ext = file.name.toLowerCase().split('.').pop();
                
                // Excel 文件：先存临时文件，再用 Excel 对象模型直接读取（无需 NPOI）
                if (ext === 'xls' || ext === 'xlsx') {
                    parseResult = await this._callParseFromFile(file, base64Data);
                } else {
                    // PDF/DOCX 等：使用 Base64 方式
                    parseResult = await this._callParseApi(file.name, base64Data);
                }
                
                this._state.parseResult = parseResult;

                if (!parseResult.success) {
                    // 扫描版 PDF 提示走 OCR
                    if (parseResult.isScanned) {
                        throw new Error(
                            '检测到扫描版 PDF，文字内容无法直接提取。建议使用 OCR 功能识别图片中的文字。'
                        );
                    }
                    throw new Error(parseResult.error || '文档解析失败');
                }

                ChatLog.info('DOCUMENT', '[DocumentProcessor] 解析完成:', {
                    type: parseResult.fileType,
                    chars: parseResult.totalChars,
                    pages: parseResult.totalPages,
                    time: parseResult.parseTimeMs + 'ms'
                });

                // Step 3: 语义分块（仅当文本超过阈值时）
                var shouldChunk = parseResult.totalChars > 3000;
                if (shouldChunk) {
                    if (options.onProgress) options.onProgress('chunking', '正在智能分块...');
                    var chunkResult = await this._callChunkApi(
                        parseResult.fullText,
                        parseResult.structure,
                        options.chunkSize || 1500,
                        options.overlap || 200,
                        options.mode || 'auto'
                    );
                    this._state.chunks = chunkResult.chunks || [];

                    ChatLog.info('DOCUMENT', '[DocumentProcessor] 分块完成:', {
                        totalChunks: chunkResult.totalChunks,
                        summary: (chunkResult.summary || '').substring(0, 100) + '...'
                    });
                } else {
                    this._state.chunks = [{
                        id: 1,
                        text: parseResult.fullText,
                        charCount: parseResult.totalChars,
                        startPosition: 0,
                        endPosition: parseResult.totalChars,
                        sourceSection: '全文',
                        level: 1
                    }];
                }

                // Step 4: 构建返回结果
                var result = {
                    success: true,
                    fileName: file.name,
                    fileType: parseResult.fileType,
                    fileIcon: detection.config.icon,
                    totalChars: parseResult.totalChars,
                    totalPages: parseResult.totalPages,
                    isScanned: parseResult.isScanned,
                    metadata: parseResult.metadata,
                    structure: parseResult.structure,
                    pages: parseResult.pages,
                    fullText: parseResult.fullText,
                    chunks: this._state.chunks,
                    chunkSummary: shouldChunk ? (await this._getChunkSummary()) : null,
                    parseTimeMs: parseResult.parseTimeMs,
                    // 便捷方法：获取用于 AI 对话的上下文文本
                    getContextForAI: function(question) {
                        return self._buildAiContext(question);
                    }
                };

                this._state.isProcessing = false;
                return result;

            } catch (error) {
                this._state.isProcessing = false;
                ChatLog.error('DOCUMENT', '[DocumentProcessor] 处理失败:', error);
                throw error;
            }
        },

        /**
         * 基于已加载的文档回答问题 (RAG 模式)
         * @param {string} question - 用户问题
         * @returns {Promise<string>} 组装好的上下文文本（供 LLM 使用）
         */
        queryDocument: async function(question) {
            if (!this._state.chunks || this._state.chunks.length === 0) {
                throw new Error('请先上传并解析文档');
            }

            var chunksJson = JSON.stringify(this._state.chunks);

            try {
                var response = await window.excelBridge.QueryDocumentAsync(question, chunksJson);
                var data = JSON.parse(response);

                if (!data.success) {
                    throw new Error(data.error || '查询失败');
                }

                ChatLog.info('DOCUMENT', '[DocumentProcessor] RAG 检索完成，相关块数:', (data.relevantChunks || []).length);
                return data.context;

            } catch (error) {
                ChatLog.warn('DOCUMENT', '[DocumentProcessor] RAG 调用失败，降级为全文模式:', error.message);
                // 降级：返回全文（截断以避免超限）
                var text = this._state.parseResult ? this._state.parseResult.fullText : '';
                return text.substring(0, 50000);
            }
        },

        /**
         * 渲染文档分析面板 UI
         * @param {object} docResult - processDocument 的返回值
         * @param {HTMLElement} container - 容器元素
         */
        renderAnalysisPanel: function(docResult, container) {
            if (!container || !docResult) return;

            var html = '';

            // 标题栏
            html += '<div class="doc-analysis-header">';
            html += '<span class="doc-icon">' + (docResult.fileIcon || '📄') + '</span>';
            html += '<span class="doc-name">' + this._escapeHtml(docResult.fileName) + '</span>';
            html += '<span class="doc-meta">' +
                this._formatNumber(docResult.totalChars) + ' 字 · ' +
                docResult.totalPages + ' 页 · ' +
                (docResult.parseTimeMs || 0) + 'ms</span>';
            html += '</div>';

            // 元数据摘要
            if (docResult.metadata) {
                html += '<div class="doc-metadata">';
                if (docResult.metadata.title) html += '<div class="meta-item"><b>标题:</b> ' + this._escapeHtml(docResult.metadata.title) + '</div>';
                if (docResult.metadata.author) html += '<div class="meta-item"><b>作者:</b> ' + this._escapeHtml(docResult.metadata.author) + '</div>';
                html += '</div>';
            }

            // 分块摘要面板
            if (docResult.chunks && docResult.chunks.length > 1) {
                html += '<div class="doc-chunks-panel">';
                html += '<div class="chunks-title">📑 文档结构 (' + docResult.chunks.length + ' 个片段)</div>';
                html += '<details open><summary>展开查看所有片段</summary><div class="chunks-list">';

                for (var i = 0; i < Math.min(docResult.chunks.length, 50); i++) {
                    var chunk = docResult.chunks[i];
                    var preview = (chunk.text || '').substring(0, 120).replace(/\n/g, ' ');
                    var section = chunk.sourceSection || ('片段 ' + (i + 1));
                    html += '<div class="chunk-item" data-chunk-id="' + chunk.id + '">' +
                        '<span class="chunk-section">' + this._escapeHtml(section) + '</span>' +
                        '<span class="chunk-preview">' + this._escapeHtml(preview) + '...</span>' +
                        '<span class="chunk-size">' + this._formatNumber(chunk.charCount) + '字</span>' +
                        '</div>';
                }
                if (docResult.chunks.length > 50) {
                    html += '<div class="chunk-more">... 还有 ' + (docResult.chunks.length - 50) + ' 个片段</div>';
                }

                html += '</div></details></div>';
            }

            // 全局摘要
            if (docResult.chunkSummary) {
                html += '<div class="doc-summary"><b>内容概要:</b><p>' +
                    this._escapeHtml(docResult.chunkSummary.substring(0, 500)) + '</p></div>';
            }

            container.innerHTML = html;
        },

        /**
         * 重置状态
         */
        reset: function() {
            this._state = {
                currentFile: null,
                parseResult: null,
                chunks: [],
                isProcessing: false,
                abortController: null
            };
        },

        // ===== 内部方法 =====

        /** 读取文件为 Base64 */
        _readFileAsBase64: function(file) {
            return new Promise(function(resolve, reject) {
                var reader = new FileReader();
                reader.onload = function(e) {
                    // 移除 data:image/...;base64, 前缀，只保留纯 base64 数据
                    var result = e.target.result;
                    var commaIdx = result.indexOf(',');
                    resolve(commaIdx > -1 ? result.substring(commaIdx + 1) : result);
                };
                reader.onerror = function() { reject(new Error('文件读取失败')); };
                reader.readAsDataURL(file);
            });
        },

        /** 调用后端 ParseDocumentAsync */
        _callParseApi: async function(fileName, base64Data) {
            var bridge = window.excelBridge;
            if (!bridge || typeof bridge.ParseDocumentAsync !== 'function') {
                throw new Error('ExcelBridge.ParseDocumentAsync 方法不可用');
            }
            ChatLog.info('DOCUMENT', '[DocumentProcessor] 发送解析请求:', fileName, 'Base64长度:', base64Data.length);
            var response = await bridge.ParseDocumentAsync(fileName, base64Data);
            ChatLog.info('DOCUMENT', '[DocumentProcessor] 收到解析响应:', response ? response.substring(0, 200) : 'null');
            return JSON.parse(response);
        },

        /** 从文件路径解析 Excel（先存临时文件，再用 Excel 对象模型读取） */
        _callParseFromFile: async function(file, base64Data) {
            ChatLog.info('DOCUMENT', '[DocumentProcessor] _callParseFromFile 被调用, file:', file ? file.name : 'null', 'base64Len:', base64Data ? base64Data.length : 0);
            var bridge = window.excelBridge;
            ChatLog.info('DOCUMENT', '[DocumentProcessor] bridge 存在:', !!bridge, 'bridge类型:', typeof bridge);
            var bridgeResult = null;
            
            if (bridge) {
                try {
                    var fileName = file.name;
                    ChatLog.info('DOCUMENT', '[DocumentProcessor] 保存临时文件:', fileName, 'Base64长度:', base64Data.length);
                    
                    var filePath = '';
                    if (typeof bridge.SaveUploadedFile === 'function') {
                        filePath = await bridge.SaveUploadedFile(fileName, base64Data);
                    }
                    
                    if (filePath && typeof bridge.ParseDocumentFromFile === 'function') {
                        var result = await bridge.ParseDocumentFromFile(filePath);
                        bridgeResult = JSON.parse(result);
                    }
                } catch (bridgeErr) {
                    ChatLog.warn('DOCUMENT', '[DocumentProcessor] Bridge 调用失败，降级到 SheetJS:', bridgeErr.message);
                }
            }
            
            if (bridgeResult && bridgeResult.success) {
                ChatLog.info('DOCUMENT', '[DocumentProcessor] Bridge 解析成功，直接返回');
                return bridgeResult;
            }
            
            ChatLog.info('DOCUMENT', '[DocumentProcessor] bridgeResult:', bridgeResult ? '有结果但success=false' : 'null', 'XLSX:', typeof XLSX);
            // 【Web版 fallback】SheetJS 直接读 File.arrayBuffer()
            ChatLog.info('DOCUMENT', '[DocumentProcessor] 使用 SheetJS 客户端解析');
            if (typeof XLSX !== 'undefined') {
                try {
                    ChatLog.info('DOCUMENT', '[DocumentProcessor] 开始 file.arrayBuffer(中');
                    var buf = await file.arrayBuffer();
                    ChatLog.info('DOCUMENT', '[DocumentProcessor] arrayBuffer 完成, size:', buf.byteLength);
                    var bytes = new Uint8Array(buf);
                    var workbook = XLSX.read(bytes, { type: 'array' });
                    ChatLog.info('DOCUMENT', '[DocumentProcessor] XLSX.read 完成, sheets:', workbook.SheetNames);

                    var pages = [];
                    var fullTextParts = [];
                    var headings = [];
                    var tables = [];
                    var startPos = 0;

                    workbook.SheetNames.forEach(function(sheetName, sheetIdx) {
                        var worksheet = workbook.Sheets[sheetName];
                        var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        var sheetLines = [];
                        sheetLines.push('【工作表: ' + sheetName + '】');
                        var headerRow = null;
                        jsonData.forEach(function(row, rowIdx) {
                            var rowText = row.map(function(cell) { return String(cell != null ? cell : ''); }).join('\t');
                            if (rowIdx === 0) {
                                sheetLines.push('【表头】' + rowText);
                                headerRow = row.map(function(cell) { return String(cell != null ? cell : ''); });
                            } else {
                                sheetLines.push(rowText);
                            }
                        });
                        var sheetText = sheetLines.join('\n');
                        pages.push({ title: sheetName, text: sheetText });
                        fullTextParts.push(sheetText);

                        headings.push({ Level: 1, Title: sheetName + ' (' + jsonData.length + '行)', StartPosition: startPos });
                        if (headerRow && headerRow.some(function(c) { return c.trim() !== ''; })) {
                            tables.push({ StartPosition: startPos, HeaderRow: headerRow, RowCount: jsonData.length - 1 });
                        }
                        startPos += sheetText.length + 2;
                    });

                    var fullText = fullTextParts.join('\n\n');
                    ChatLog.info('DOCUMENT', '[DocumentProcessor] SheetJS 解析成功:', workbook.SheetNames.length, '个工作表, 总字符:', fullText.length);

                    return {
                        success: true,
                        fileType: 'excel',
                        fullText: fullText,
                        totalChars: fullText.length,
                        totalPages: pages.length,
                        isScanned: false,
                        pages: pages,
                        structure: { Headings: headings, Tables: tables },
                        metadata: {
                            fileFormat: (file.name || '').split('.').pop() || 'xlsx',
                            properties: { '工作表列表': workbook.SheetNames.join(', ') }
                        },
                        parseTimeMs: 0
                    };
                } catch (sheetjsErr) {
                    ChatLog.error('DOCUMENT', '[DocumentProcessor] SheetJS 解析失败:', sheetjsErr);
                    throw new Error('Excel 解析失败: ' + sheetjsErr.message);
                }
            }
            throw new Error('Excel 解析不可用');
        },

        /** 调用后端 ChunkDocumentAsync */
        _callChunkApi: async function(fullText, structureJson, chunkSize, overlap, mode) {
            var bridge = window.excelBridge;
            
            // 优先尝试后端语义分块（已加30秒超时保护）
            if (bridge && typeof bridge.ChunkDocumentAsync === 'function') {
                try {
                    var optionsJson = JSON.stringify({ chunkSize: chunkSize, overlap: overlap, mode: mode });
                    var structureParam = structureJson ? JSON.stringify(structureJson.headings || []) : '';
                    
                    // 前端也加超时保护（35秒，略大于后端30秒）
                    var timeoutPromise = new Promise(function(_, reject) {
                        setTimeout(function() { reject(new Error('分块超时')); }, 35000);
                    });
                    var chunkPromise = bridge.ChunkDocumentAsync(fullText, structureParam, optionsJson);
                    
                    var response = await Promise.race([chunkPromise, timeoutPromise]);
                    var result = JSON.parse(response);
                    
                    if (result.success) {
                        ChatLog.info('DOCUMENT', '[DocumentProcessor] 后端语义分块成功:', result.totalChunks, '块');
                        return result;
                    } else {
                        ChatLog.warn('DOCUMENT', '[DocumentProcessor] 后端分块返回失败:', result.error);
                    }
                } catch (e) {
                    ChatLog.warn('DOCUMENT', '[DocumentProcessor] 后端分块失败，使用前端降级:', e.message);
                }
            }
            
            // 降级：前端简单分块
            ChatLog.info('DOCUMENT', '[DocumentProcessor] 使用前端分块');
            return this._fallbackChunk(fullText, chunkSize, overlap);
        },

        /** 前端降级分块（后端不可用时） */
        _fallbackChunk: function(text, chunkSize, overlap) {
            var chunks = [];
            var pos = 0;
            var id = 1;

            while (pos < text.length) {
                var end = Math.min(pos + chunkSize, text.length);
                // 尝试在句子边界断开
                if (end < text.length) {
                    var lastPeriod = text.lastIndexOf('。', end);
                    var lastDot = text.lastIndexOf('. ', end);
                    var lastBreak = Math.max(lastPeriod, lastDot);
                    if (lastBreak > pos + chunkSize * 0.5) end = lastBreak + 1;
                }

                chunks.push({
                    id: id++,
                    text: text.substring(pos, end),
                    charCount: end - pos,
                    startPosition: pos,
                    endPosition: end,
                    sourceSection: '片段 ' + (id - 1),
                    level: 1
                });

                pos = end - (overlap > 0 && end < text.length ? overlap : 0);
            }

            return { totalChunks: chunks.length, chunks: chunks, summary: '' };
        },

        /** 构建分块摘要文本 */
        _getChunkSummary: async function() {
            var chunks = this._state.chunks;
            if (!chunks || chunks.length === 0) return '';

            var parts = [];
            var totalLen = 0;
            var maxLen = 800;

            for (var i = 0; i < chunks.length && totalLen < maxLen; i++) {
                var prefix = chunks[i].sourceSection ? '[' + chunks[i].sourceSection + '] ' : '';
                var snippet = (chunks[i].text || '').substring(0, 80);
                parts.push(prefix + snippet.replace(/\n/g, ' '));
                totalLen += parts[parts.length - 1].length;
            }

            return parts.join('\n') + (chunks.length > parts.length ? '\n... 共 ' + chunks.length + ' 个片段' : '');
        },

        /** 构建 AI 上下文（RAG 或全文） */
        _buildAiContext: function(question) {
            var parseResult = this._state.parseResult;
            var chunks = this._state.chunks;

            if (!parseResult) return '';

            // 短文档：直接返回全文
            if ((chunks && chunks.length <= 3) || (parseResult.totalChars <= 8000)) {
                return '[文档内容]\n文件名: ' + (this._state.currentFile ? this._state.currentFile.name : '') +
                    '\n\n' + parseResult.fullText;
            }

            // 长文档：返回前几个块的预览 + 提示用户可提问
            var context = '[文档内容]\n';
            context += '文件名: ' + (this._state.currentFile ? this._state.currentFile.name : '') + '\n';
            context += '总字符数: ' + parseResult.totalChars + ', 共 ' + (chunks ? chunks.length : 0) + ' 个片段\n\n';

            // 取前3个块作为预览
            var previewCount = Math.min(3, chunks ? chunks.length : 0);
            for (var j = 0; j < previewCount; j++) {
                context += '--- ' + (chunks[j].sourceSection || '片段' + (j+1)) + ' ---\n';
                context += chunks[j].text.substring(0, 500) + '\n\n';
            }

            if ((chunks ? chunks.length : 0) > previewCount) {
                context += '... (还有 ' + ((chunks ? chunks.length : 0) - previewCount) + ' 个片段)\n';
            }

            return context;
        },

        /** 获取文件扩展名 */
        _getFileExt: function(fileName) {
            if (!fileName) return '';
            var dotIdx = fileName.lastIndexOf('.');
            return dotIdx > -1 ? fileName.substring(dotIdx).toLowerCase() : '';
        },

        /** 格式化数字 */
        _formatNumber: function(n) {
            if (n >= 10000) return (n / 10000).toFixed(1) + '万';
            if (n >= 1000) return (n / 1000).toFixed(1) + 'k';
            return String(n);
        },

        /** HTML 转义 */
        _escapeHtml: function(str) {
            if (!str) return '';
            return str.replace(/&/g, '&amp;')
                      .replace(/</g, '&lt;')
                      .replace(/>/g, '&gt;')
                      .replace(/"/g, '&quot;');
        }
    };

    // ============================================================
    // 全局挂载
    // ============================================================
    App.DocumentProcessor = DocumentProcessor;

})();
