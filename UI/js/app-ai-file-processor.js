/**
 * ============================================================
 * app-ai-file-processor.js — File Processor
 * ============================================================
 * 功能：文件处理模块，负责单文件与多文件的识别、OCR、内容提取与上下文构建。
 *
 * 核心类/函数：
 *   - SingleFileProcessor — 单文件处理器
 *   - MultiFileProcessor — 多文件处理器
 *   - processFile() — 处理单个文件
 *
 * 依赖：
 *   - app-ai-context.js
 *   - app-ai-calls.js
 * ============================================================
 */

(function() {
    'use strict';

    if (!window.App) window.App = {};

    // ============================================================
    // 【单文件处理器】处理单个文件的完整流程
    // ============================================================
    class SingleFileProcessor {
        constructor(context) {
            this.ctx = context;
            this.documentTextContent = null;
            this.fileContext = null;
            this.isDocument = false;
            this.isImage = false;
        }

        /**
         * 处理文件并返回结果
         * @returns {Promise<{success: boolean, content?: string, error?: string}>}
         */
        async process() {
            const { file, loadingMessageId } = this.ctx;
            
            if (!file) {
                return { success: false, error: '没有文件可处理' };
            }

            const fileName = file.name || '未知文件';
            const fileExt = (fileName.match(/\.(\w+)$/) || [])[1]?.toLowerCase();
            
            // 判断文件类型
            this.isDocument = this._isDocumentType(fileExt);
            this.isImage = this._isImageType(fileExt);

            ChatLog.info('FILE', `处理文件: ${fileName}, 类型: ${fileExt}, 文档: ${this.isDocument}, 图片: ${this.isImage}`);

            if (loadingMessageId) {
                App.updateAIThinkingStatus(loadingMessageId, `📄 本地预处理 ${fileName}...`);
            }

            try {
                if (this.isDocument) {
                    return await this._processDocument();
                } else if (this.isImage) {
                    return await this._processImage();
                } else {
                    return await this._processOther();
                }
            } catch (error) {
                ChatLog.error('FILE', `文件处理失败: ${fileName}`, error);
                return { success: false, error: error.message || '文件处理异常' };
            }
        }

        /**
         * 判断是否为文档类型
         */
        _isDocumentType(ext) {
            return ['xlsx', 'xls', 'docx', 'doc', 'pdf', 'txt', 'csv', 'md'].includes(ext);
        }

        /**
         * 判断是否为图片类型
         */
        _isImageType(ext) {
            return ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'].includes(ext);
        }

        /**
         * 处理文档文件
         */
        async _processDocument() {
            const { file, loadingMessageId } = this.ctx;
            const fileName = file.name || '未知文件';
            const fileExt = (fileName.match(/\.(\w+)$/) || [])[1]?.toLowerCase();

            if (loadingMessageId) {
                App.updateAIThinkingStatus(loadingMessageId, `📖 本地解析文档 ${fileName}...`);
            }

            // 读取文件内容
            const fileContent = await this._readFileAsBase64(file);
            if (!fileContent) {
                return { success: false, error: '无法读取文件内容' };
            }

            // 解析文档
            let documentText = null;
            
            if (fileExt === 'xlsx' || fileExt === 'xls') {
                documentText = await this._parseExcel(fileContent, fileName);
            } else if (fileExt === 'docx' || fileExt === 'doc') {
                documentText = await this._parseWord(fileContent, fileName);
            } else if (fileExt === 'pdf') {
                documentText = await this._parsePdf(fileContent, fileName);
            } else if (fileExt === 'txt' || fileExt === 'csv' || fileExt === 'md') {
                documentText = await this._readAsText(file);
            }

            if (documentText) {
                this.documentTextContent = documentText;
                this.fileContext = {
                    file: file,
                    base64Data: fileContent,
                    fileName: fileName
                };
                return { success: true, content: documentText };
            }

            return { success: false, error: '文档解析失败' };
        }

        /**
         * 处理图片文件
         */
        async _processImage() {
            const { file, loadingMessageId } = this.ctx;
            const fileName = file.name || '未知文件';

            if (loadingMessageId) {
                App.updateAIThinkingStatus(loadingMessageId, `🖼️ 本地读取图片 ${fileName}...`);
            }

            const base64Data = await this._readFileAsBase64(file);
            if (!base64Data) {
                return { success: false, error: '无法读取图片' };
            }

            this.fileContext = {
                file: file,
                base64Data: base64Data,
                fileName: fileName
            };

            return { success: true, content: null };
        }

        /**
         * 处理其他类型文件
         */
        async _processOther() {
            const { file } = this.ctx;
            
            // 尝试作为文本读取
            try {
                const text = await this._readAsText(file);
                if (text && text.trim().length > 0) {
                    this.documentTextContent = text;
                    return { success: true, content: text };
                }
            } catch (e) {
                // 忽略
            }

            return { success: false, error: '不支持的文件类型' };
        }

        /**
         * 读取文件为 Base64
         */
        async _readFileAsBase64(file) {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const result = reader.result;
                    const base64 = result.includes(',') ? result.split(',')[1] : result;
                    resolve(base64);
                };
                reader.onerror = () => resolve(null);
                reader.readAsDataURL(file);
            });
        }

        /**
         * 读取文件为文本
         */
        async _readAsText(file) {
            return new Promise((resolve) => {
                const reader = new FileReader();
                reader.onload = () => resolve(reader.result);
                reader.onerror = () => resolve(null);
                reader.readAsText(file);
            });
        }

        /**
         * 统一通过 DocumentBridge 解析文档（PDF/DOCX/XLSX）
         * ExcelBridge 暴露的 ParseDocumentAsync 会自动路由到对应 Parser
         */
        async _parseDocumentViaBridge(base64Data, fileName) {
            try {
                if (App.excelBridge && App.excelBridge.ParseDocumentAsync) {
                    const rawResult = await App.excelBridge.ParseDocumentAsync(fileName, base64Data);
                    if (!rawResult) return null;

                    let result = rawResult;
                    if (typeof rawResult === 'string') {
                        try {
                            result = JSON.parse(rawResult);
                        } catch (parseErr) {
                            ChatLog.warn('FILE', '文档解析结果 JSON 解析失败:', parseErr.message);
                            return null;
                        }
                    }

                    if (result && result.success) {
                        // DocumentBridge 返回 fullText，兼容旧接口的 content 字段
                        return result.fullText || result.content || null;
                    } else {
                        ChatLog.warn('FILE', `文档解析未成功: ${result?.error || '未知错误'}`);
                    }
                } else {
                    ChatLog.warn('FILE', 'ExcelBridge.ParseDocumentAsync 不可用，无法解析文档');
                }
            } catch (e) {
                ChatLog.warn('FILE', `${fileName} 文档解析异常:`, e.message);
            }
            return null;
        }

        /**
         * 解析 Excel 文件
         */
        async _parseExcel(base64Data, fileName) {
            return this._parseDocumentViaBridge(base64Data, fileName);
        }

        /**
         * 解析 Word 文件
         */
        async _parseWord(base64Data, fileName) {
            return this._parseDocumentViaBridge(base64Data, fileName);
        }

        /**
         * 解析 PDF 文件
         */
        async _parsePdf(base64Data, fileName) {
            return this._parseDocumentViaBridge(base64Data, fileName);
        }
    }

    // ============================================================
    // 【多文件处理器】处理多个文件的完整流程
    // ============================================================
    class MultiFileProcessor {
        constructor(context) {
            this.ctx = context;
            this.results = [];
        }

        /**
         * 处理多个文件
         * @returns {Promise<{success: boolean, contents?: Array, error?: string}>}
         */
        async process() {
            const { files, loadingMessageId } = this.ctx;
            
            if (!files || files.length === 0) {
                return { success: false, error: '没有文件可处理' };
            }

            ChatLog.info('FILE', `批量处理 ${files.length} 个文件`);

            if (loadingMessageId) {
                App.updateAIThinkingStatus(loadingMessageId, `📁 本地批量处理 ${files.length} 个文件...`);
            }

            // 逐个处理文件
            for (let i = 0; i < files.length; i++) {
                const file = files[i];
                const progress = `[${i + 1}/${files.length}]`;
                
                if (loadingMessageId) {
                    App.updateAIThinkingStatus(loadingMessageId, `${progress} 本地预处理 ${file.name}...`);
                }

                const processor = new SingleFileProcessor({
                    ...this.ctx,
                    file: file
                });

                const result = await processor.process();
                this.results.push({
                    fileName: file.name,
                    ...result
                });
            }

            const successCount = this.results.filter(r => r.success).length;
            ChatLog.info('FILE', `批量处理完成: ${successCount}/${files.length} 成功`);

            return {
                success: successCount > 0,
                contents: this.results,
                error: successCount === 0 ? '所有文件处理失败' : null
            };
        }
    }

    // ============================================================
    // 【OCR处理器】处理图片OCR识别
    // ============================================================
    class OcrProcessor {
        constructor(context) {
            this.ctx = context;
            this.engines = ['paddleocr', 'rapidocr'];
            this.usedEngine = null;
        }

        /**
         * 执行OCR识别（带Fallback链）
         * @returns {Promise<{success: boolean, result?: string, engine?: string, error?: string}>}
         */
        async process() {
            const { fileContext, loadingMessageId } = this.ctx;
            
            if (!fileContext || !fileContext.base64Data) {
                return { success: false, error: '没有图片数据' };
            }

            for (const engine of this.engines) {
                ChatLog.info('OCR', `尝试引擎: ${engine}`);
                
                if (loadingMessageId) {
                    App.updateAIThinkingStatus(loadingMessageId, `📝 OCR识别中 (${engine}中`);
                }

                try {
                    const result = await this._tryEngine(engine);
                    if (result.success) {
                        this.usedEngine = engine;
                        return {
                            success: true,
                            result: result.content,
                            engine: engine
                        };
                    }
                } catch (e) {
                    ChatLog.warn('OCR', `${engine} 失败:`, e.message);
                }
            }

            return { success: false, error: '所有OCR引擎均失败' };
        }

        /**
         * 尝试单个OCR引擎
         */
        async _tryEngine(engine) {
            const { fileContext } = this.ctx;
            const base64Data = fileContext.base64Data;
            const fileName = fileContext.fileName || 'image.png';

            // 切换引擎
            if (App.excelBridge && App.excelBridge.SetOcrEngine) {
                await App.excelBridge.SetOcrEngine(engine);
            }

            // 保存图片
            let imagePath = await this._saveImage(base64Data, fileName);
            if (!imagePath) {
                return { success: false, error: '图片保存失败' };
            }

            // 执行OCR
            const ocrToolCall = {
                id: 'call_ocr_' + Date.now(),
                type: 'function',
                function: {
                    name: 'excel-local__excel_ocr_recognize',
                    arguments: JSON.stringify({ imagePath })
                }
            };

            const ocrResult = await App.executeMcpToolCall(ocrToolCall);
            const ocrResultStr = typeof ocrResult === 'string' ? ocrResult : JSON.stringify(ocrResult);

            // 检查结果
            if (this._isValidResult(ocrResultStr)) {
                return { success: true, content: ocrResultStr };
            }

            return { success: false, error: 'OCR结果为空或无效' };
        }

        /**
         * 保存图片到临时目录
         */
        async _saveImage(base64Data, fileName) {
            if (App.excelBridge && App.excelBridge.SaveBase64Image && !window.chrome?.webview?._mockMode) {
                const saveResult = await App.excelBridge.SaveBase64Image(base64Data, fileName);
                const saveResultObj = JSON.parse(saveResult);
                if (saveResultObj.success) {
                    return saveResultObj.path;
                }
            }

            // Fallback: HTTP API
            try {
                const bridgeUrl = window.ExcelBridgeConfig ? window.ExcelBridgeConfig.baseUrl : 'http://localhost:9527/api/excel';
                const response = await fetch(bridgeUrl + '/invoke', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ method: 'SaveBase64Image', arguments: [base64Data, fileName] })
                });
                const data = await response.json();
                if (data.success && data.data?.path) {
                    return data.data.path;
                }
            } catch (e) {
                ChatLog.error('OCR', '图片保存失败:', e);
            }

            return null;
        }

        /**
         * 验证OCR结果是否有效
         */
        _isValidResult(resultStr) {
            if (!resultStr || /^\s*\{?\s*"?mock/i.test(resultStr) || resultStr === '{}' || resultStr === '""') {
                return false;
            }
            if (resultStr.startsWith('【错误：')) {
                return false;
            }
            try {
                const json = JSON.parse(resultStr);
                if (json.success === false || json.isError === true || json.error) {
                    return false;
                }
            } catch (e) {
                // 纯文本结果
            }
            return resultStr.trim().length > 10;
        }
    }

    // ============================================================
    // 【工厂函数】创建文件处理器
    // ============================================================
    function createFileProcessor(context) {
        const { files } = context;
        
        if (files && files.length > 1) {
            return new MultiFileProcessor(context);
        } else if (files && files.length === 1) {
            return new SingleFileProcessor({ ...context, file: files[0] });
        } else if (context.file) {
            return new SingleFileProcessor(context);
        }
        
        return null;
    }

    function createOcrProcessor(context) {
        return new OcrProcessor(context);
    }

    // ============================================================
    // Exports
    // ============================================================
    App.FileProcessor = {
        SingleFileProcessor,
        MultiFileProcessor,
        OcrProcessor,
        createFileProcessor,
        createOcrProcessor
    };

})();
