/**
 * ============================================================
 * app-file-batch.js — File Batch Processor
 * ============================================================
 * 功能：文件批处理模块，支持多文件批量识别、汇总分析与批量导出。
 *
 * 核心类/函数：
 *   - processBatch() — 批量处理文件
 *   - addBatchItem() — 添加批处理项
 *   - getBatchResults() — 获取批量结果
 *   - exportBatchResults() — 导出批量结果
 *
 * 依赖：
 *   - app-ai-file-processor.js
 *   - app-ocr.js
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    // 文件类型图标映射
    const FILE_ICONS = {
        image: '🖼️',
        pdf: '📕',
        excel: '📊',
        word: '📝',
        text: '📄',
        default: '📎'
    };

    // --- 来自 app.js:L5712-L5719 ---
    function getFileIcon(file) {
        if (file.type.startsWith('image/')) return FILE_ICONS.image;
        if (file.type === 'application/pdf') return FILE_ICONS.pdf;
        if (file.type.includes('excel') || file.type.includes('spreadsheet')) return FILE_ICONS.excel;
        if (file.type.includes('word') || file.type.includes('document')) return FILE_ICONS.word;
        if (file.type.startsWith('text/')) return FILE_ICONS.text;
        return FILE_ICONS.default;
    }

    // --- 获取文件类型分类标识 ---
    function getFileType(file) {
        if (!file || !file.type) return 'other';
        if (file.type.startsWith('image/')) return 'image';
        if (file.type === 'application/pdf') return 'pdf';
        if (file.type.includes('excel') || file.type.includes('spreadsheet') || file.type.includes('csv')) return 'excel';
        if (file.type.includes('word') || file.type.includes('document')) return 'word';
        if (file.type.startsWith('text/') || file.name?.endsWith('.txt') || file.name?.endsWith('.md') || file.name?.endsWith('.csv')) return 'text';
        return 'other';
    }

    // --- 获取文件类型显示标签 ---
    function getFileTypeLabel(fileType) {
        const labels = {
            image: '图片',
            pdf: 'PDF',
            excel: '表格',
            word: '文档',
            text: '文本',
            other: '附件'
        };
        return labels[fileType] || '附件';
    }

    // --- 来自 app.js:L5722-L5728 ---
    // 格式化文件大小
    function formatFileSize(bytes) {
        if (bytes === 0) return '0 B';
        const k = 1024;
        const sizes = ['B', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    // HTML 转义（附件列表渲染用）
    function escapeHtml(str) {
        if (str == null) return '';
        return String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    // --- 来自 app.js:L6985-L6999 ---
    /**
     * 检测图片类型（简化版）
     */
    function detectImageType(file) {
        const name = file.name.toLowerCase();

        if (name.includes('发票') || name.includes('invoice') || name.includes('fp')) {
            return 'invoice';
        } else if (name.includes('身份证') || name.includes('idcard') || name.includes('sfz')) {
            return 'card';
        } else if (name.includes('名片') || name.includes('card') || name.includes('mp')) {
            return 'card';
        } else if (name.includes('表格') || name.includes('table') || name.includes('bg')) {
            return 'table';
        }

        return 'general';  // 默认通用类型
    }

    // --- 来自 app.js:L7005-L7012 ---
    // 将文件转换为 Base64（带缓存，避免重复读取）
    function fileToBase64(file) {
        if (file && file._cachedBase64) {
            ChatLog.info('FILEBATCH', `[fileToBase64] 使用缓存的 base64: ${file.name}`);
            return Promise.resolve(file._cachedBase64);
        }
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => {
                if (file) file._cachedBase64 = reader.result;
                resolve(reader.result);
            };
            reader.onerror = (e) => reject(e);
            reader.readAsDataURL(file);
        });
    }

    // ============================================================
    // 【OCR结果清理】移除明显的垃圾文本和UI元素
    // ============================================================
    function cleanOcrResult(text) {
        if (!text) return '';
        
        let cleaned = text;
        
        // 1. 移除常见的UI元素/按钮文本
        const uiPatterns = [
            /引擎\s*\|\s*PaddleOCR/gi,
            /中文推荐/gi,
            /新金词库/gi,
            /限票当当次车/gi,
            /保存|复制|下载|上传|选择文件/gi,
            /\[.*?\]/g,  // 移除方括号内容（通常是UI标签）
        ];
        
        for (const pattern of uiPatterns) {
            cleaned = cleaned.replace(pattern, '');
        }
        
        // 2. 移除过短的行（可能是OCR噪声）
        const lines = cleaned.split('\n');
        const filteredLines = lines.filter(line => {
            const trimmed = line.trim();
            // 保留非空行，且长度大于2或包含中文字符
            return trimmed.length === 0 || trimmed.length > 2 || /[\u4e00-\u9fa5]/.test(trimmed);
        });
        cleaned = filteredLines.join('\n');
        
        // 3. 清理多余空行
        cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
        
        // 4. 移除首尾空白
        cleaned = cleaned.trim();
        
        return cleaned;
    }

    // --- 来自 app.js:L7014-L7028 ---
    // ==================== 辅助函数：映射强制类型到内部类型 ====================
    function mapForcedType(forcedType) {
        const typeMap = {
            '身份证': 'card',      // 映射到 card 配置
            'idcard': 'card',
            '名片': 'card',
            'businesscard': 'card',
            '发票': 'invoice',     // 映射到 invoice 配置
            'invoice': 'invoice',
            '表格': 'table',       // 映射到 table 配置
            'table': 'table'
        };
        const mapped = typeMap[forcedType.toLowerCase()] || 'general';
        ChatLog.info('FILEBATCH', `[mapForcedType] ${forcedType} -> ${mapped}`);
        return mapped;
    }

    // ==================== Excel 结构化摘要构建器（P2 增强版） ====================
    // 将 Excel 解析结果转换为 AI 友好的结构化摘要
    // 每个 Sheet 展示：表名、行列数、表头+列类型、数值列统计、样本数据、空值检测

    // --- 辅助：判断值是否为数字 ---
    function _isNumericVal(v) {
        if (!v || v.trim() === '') return false;
        var n = Number(v.trim());
        return !isNaN(n) && isFinite(n);
    }

    // --- 辅助：判断值是否为日期格式 ---
    function _isDateVal(v) {
        if (!v) return false;
        var s = v.trim();
        return /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}/.test(s) || /^\d{1,2}[-\/]\d{1,2}[-\/]\d{4}/.test(s);
    }

    // --- 辅助：推断列类型 ---
    function _inferColumnType(values) {
        var numericCount = 0, dateCount = 0, emptyCount = 0, textCount = 0;
        var total = values.length;
        for (var i = 0; i < total; i++) {
            var v = values[i];
            if (!v || v.trim() === '') { emptyCount++; continue; }
            if (_isDateVal(v)) { dateCount++; continue; }
            if (_isNumericVal(v)) { numericCount++; continue; }
            textCount++;
        }
        var nonEmpty = total - emptyCount;
        if (nonEmpty === 0) return 'empty';
        if (numericCount / nonEmpty > 0.8) return 'numeric';
        if (dateCount / nonEmpty > 0.8) return 'date';
        if (textCount / nonEmpty > 0.8) return 'text';
        return 'mixed';
    }

    // --- 辅助：计算数值列统计 ---
    function _numericStats(values) {
        var nums = [];
        for (var i = 0; i < values.length; i++) {
            if (_isNumericVal(values[i])) nums.push(Number(values[i].trim()));
        }
        if (nums.length === 0) return null;
        var min = nums[0], max = nums[0], sum = 0;
        for (var j = 0; j < nums.length; j++) {
            if (nums[j] < min) min = nums[j];
            if (nums[j] > max) max = nums[j];
            sum += nums[j];
        }
        return { min: min, max: max, avg: sum / nums.length, count: nums.length };
    }

    // --- 辅助：从 structure.headings 提取 Sheet 行数 ---
    function _getSheetRowCounts(parseResult) {
        var counts = {};
        if (!parseResult.structure) return counts;
        var headings = parseResult.structure.headings || parseResult.structure.Headings || [];
        for (var i = 0; i < headings.length; i++) {
            var h = headings[i];
            var title = h.Title || h.title || '';
            var m = title.match(/\((\d+)\s*行/);
            if (m) counts[i] = parseInt(m[1]);
        }
        return counts;
    }

    // --- 辅助：解析 Sheet 文本，返回 {header, allDataRows} ---
    function _parseSheetText(sheetText) {
        var sheetLines = sheetText.split('\n');
        var headerRow = null;
        var allDataRows = [];
        for (var si = 0; si < sheetLines.length; si++) {
            var line = sheetLines[si].trim();
            if (!line) continue;
            if (line.indexOf('【表头】') === 0) {
                var rawHeader = line.substring(4).trim();
                // 如果【表头】行不含制表符，说明是标题行（如"2025年上半年财务报表"），
                // 真正的字段名在下一行。跳过标题，取下一行作为表头。
                if (rawHeader.indexOf('\t') < 0) {
                    var nextSi = si + 1;
                    while (nextSi < sheetLines.length && !sheetLines[nextSi].trim()) nextSi++;
                    if (nextSi < sheetLines.length) {
                        headerRow = sheetLines[nextSi].trim();
                        si = nextSi; // 跳过真正的表头行，不加入 allDataRows
                    } else {
                        headerRow = rawHeader;
                    }
                } else {
                    headerRow = rawHeader;
                }
                continue;
            }
            if (line.indexOf('【工作表:') === 0) continue;
            allDataRows.push(line);
        }
        return { header: headerRow, rows: allDataRows };
    }

    // 自动检测分隔符（| 或 \t），解决后端返回 | 分隔但前端按 \t 分割导致列数永远为1的问题
    function _splitCells(text) {
        if (!text) return [];
        if (text.indexOf('|') >= 0) { var r = text.split('|').map(function(s) { return s.trim(); }); ChatLog.info('FILEBATCH', '[splitCells] pipe分隔, 列数:', r.length, '| 前60字:', text.substring(0,60)); return r; }
        var r2 = text.split('\t'); ChatLog.info('FILEBATCH', '[splitCells] tab分隔, 列数:', r2.length, '| 前60字:', text.substring(0,60), '| charCodes:', Array.prototype.slice.call(text.substring(0,20)).map(function(c){return c.charCodeAt(0)}).join(',')); return r2;
    }

    // --- 辅助：构建单个 Sheet 的摘要行 ---
    function _buildSheetSummary(sheetName, sheetIndex, headerRow, allDataRows, totalRowCount, colCount) {
        var out = [];
        var rowCountLabel = totalRowCount > 0 ? (totalRowCount + '行') : (allDataRows.length + '行');
        out.push('━━━ 工作表 ' + (sheetIndex + 1) + ': ' + sheetName + ' (' + rowCountLabel + (colCount ? ' x ' + colCount + '列' : '') + ') ━━━');

        if (!headerRow || allDataRows.length === 0) {
            out.push('  (空工作表)');
            out.push('');
            return out;
        }

        var headerCols = _splitCells(headerRow);
        var numCols = headerCols.length;

        // 收集每列的所有样本值
        var colValues = [];
        for (var c = 0; c < numCols; c++) colValues[c] = [];
        for (var ri = 0; ri < allDataRows.length; ri++) {
            var cells = _splitCells(allDataRows[ri]);
            for (var c2 = 0; c2 < numCols; c2++) {
                colValues[c2].push(c2 < cells.length ? cells[c2] : '');
            }
        }

        // 表头 + 列类型
        var typeLabels = [];
        for (var c3 = 0; c3 < numCols; c3++) {
            var colType = _inferColumnType(colValues[c3]);
            var typeMark = colType === 'numeric' ? '🔢' : colType === 'date' ? '📅' : colType === 'empty' ? '⚪' : '';
            typeLabels.push(headerCols[c3] + (typeMark ? ' ' + typeMark : ''));
        }
        out.push('  表头: ' + typeLabels.join(' | '));

        // 数值列统计
        var hasNumericStats = false;
        var statsParts = [];
        for (var c4 = 0; c4 < numCols; c4++) {
            var stats = _numericStats(colValues[c4]);
            if (stats) {
                hasNumericStats = true;
                var avgStr = stats.avg === Math.floor(stats.avg) ? stats.avg.toString() : stats.avg.toFixed(2);
                statsParts.push(headerCols[c4] + ': ' + stats.min + '~' + stats.max + '(avg ' + avgStr + ')');
            }
        }
        if (hasNumericStats) {
            out.push('  数值范围: ' + statsParts.join('; '));
        }

        // 空值检测
        var emptyCols = [];
        for (var c5 = 0; c5 < numCols; c5++) {
            var emptyCount = 0;
            for (var ei = 0; ei < colValues[c5].length; ei++) {
                if (!colValues[c5][ei] || colValues[c5][ei].trim() === '') emptyCount++;
            }
            if (emptyCount > 0 && colValues[c5].length > 0) {
                var ratio = (emptyCount / colValues[c5].length * 100).toFixed(0);
                emptyCols.push(headerCols[c5] + ' ' + ratio + '%空');
            }
        }
        if (emptyCols.length > 0) {
            out.push('  ⚠ 空值列: ' + emptyCols.join(', '));
        }

        // 智能采样：前3行 + 最后1行（超过4行时）
        var sampleRows = [];
        if (allDataRows.length <= 4) {
            sampleRows = allDataRows.slice();
        } else {
            sampleRows = allDataRows.slice(0, 3);
            sampleRows.push(allDataRows[allDataRows.length - 1]);
        }
        out.push('  样本 (' + sampleRows.length + '/' + allDataRows.length + '行):');
        for (var di = 0; di < sampleRows.length; di++) {
            var isLast = (allDataRows.length > 4 && di === sampleRows.length - 1);
            var prefix = isLast ? '[末]' : '[' + (di + 1) + ']';
            var sCells = _splitCells(sampleRows[di]);
            out.push('    ' + prefix + ' ' + sCells.join(' | '));
        }

        out.push('');
        return out;
    }

    function buildExcelStructuredSummary(parseResult, fileName) {
        if (!parseResult) return '[Excel文件] ' + (fileName || '') + '\n注：解析结果为空';

        var lines = [];
        var sheetCount = parseResult.totalPages || (parseResult.structure && parseResult.structure.sectionCount) || 0;
        var totalChars = parseResult.totalChars || 0;

        lines.push('[Excel文件] ' + (fileName || parseResult.fileName || '未知'));
        lines.push('工作表数量: ' + sheetCount);
        lines.push('总字符数: ' + totalChars);
        lines.push('');

        // 提取 Sheet 行数映射
        var rowCountMap = _getSheetRowCounts(parseResult);

        // 优先使用 pages 数组
        var pages = parseResult.pages || [];
        if (pages.length > 0) {
            for (var pi = 0; pi < pages.length; pi++) {
                var page = pages[pi];
                var sheetName = page.title || page.Title || ('Sheet' + (pi + 1));
                var sheetText = page.text || page.Text || '';
                var parsed = _parseSheetText(sheetText);
                ChatLog.info('FILEBATCH', '[buildExcel] Sheet', pi, ':', sheetName, '| header:', JSON.stringify(parsed.header), '| headerLen:', (parsed.header||'').length, '| sheetText前100字:', sheetText.substring(0,100));
                var totalRows = rowCountMap[pi] || 0;
                var colCount = parsed.header ? _splitCells(parsed.header).length : 0;
                var sheetLines = _buildSheetSummary(sheetName, pi, parsed.header, parsed.rows, totalRows, colCount);
                for (var li = 0; li < sheetLines.length; li++) lines.push(sheetLines[li]);
            }
        } else if (parseResult.fullText) {
            // 降级：从 fullText 提取
            var fullText = parseResult.fullText;
            var sheetMarkers = fullText.split(/【工作表:/);
            for (var mi = 1; mi < sheetMarkers.length; mi++) {
                var section = sheetMarkers[mi];
                var nameEnd = section.indexOf('】');
                var sName = nameEnd > 0 ? section.substring(0, nameEnd).trim() : ('Sheet' + mi);
                var sectionText = '【工作表: ' + section;
                var parsed2 = _parseSheetText(sectionText);
                var totalRows2 = rowCountMap[mi - 1] || 0;
                var colCount2 = parsed2.header ? _splitCells(parsed2.header).length : 0;
                var sheetLines2 = _buildSheetSummary(sName, mi - 1, parsed2.header, parsed2.rows, totalRows2, colCount2);
                for (var li2 = 0; li2 < sheetLines2.length; li2++) lines.push(sheetLines2[li2]);
            }
            if (sheetMarkers.length <= 1) {
                lines.push('(无法识别工作表结构，展示原文前2000字符)');
                lines.push(fullText.substring(0, 2000));
            }
        }

        // 元数据
        if (parseResult.metadata && parseResult.metadata.properties) {
            var props = parseResult.metadata.properties;
            if (props['工作表列表']) lines.push('【元数据】工作表列表: ' + props['工作表列表']);
            if (props['格式']) lines.push('【元数据】格式: ' + props['格式']);
        }

        var summary = lines.join('\n');
        if (summary.length > 8000) {
            summary = summary.substring(0, 8000) + '\n...(摘要已截断，完整数据请通过 MCP 工具按需读取)';
        }
        return summary;
    }

    // ==================== 来自原app.js:L5120-L5574 ====================
    // 统一文件处理器（验证、类型检测、内容提取、批量处理）
    var UnifiedFileProcessor = {
        // 统一入口：处理文件上传（自动识别单/多文件）
        handleFiles: async function(files, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 开始处理 ' + files.length + ' 个文件');
            var fileArray = Array.from(files);
            if (fileArray.length === 0) { ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] 文件列表为空'); return; }
            try {
                if (fileArray.length === 1) {
                    ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 检测到单文件，使用优化流程');
                    await this._processSingleFile(fileArray[0], options);
                } else {
                    ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 检测到多文件(' + fileArray.length + '个)，使用批量流程');
                    await this._processMultipleFiles(fileArray, options);
                }
            } catch (error) {
                ChatLog.error('FILEBATCH', '[UnifiedFileProcessor] 处理失败:', error); throw error;
            }
        },

        // 验证文件是否有效
        validateFile: function(file) {
            if (!file) return { valid: false, error: '文件对象为空', type: null };
            var ftype = (file.type || '').toLowerCase();
            var fname = (file.name || '').toLowerCase();
            var isPdf = (ftype === 'application/pdf' || fname.indexOf('.pdf') !== -1);
            // PDF 最大 200MB，其他文档最大 50MB，图片保持 50MB
            var maxSize = isPdf ? (200 * 1024 * 1024) : (50 * 1024 * 1024);
            if (file.size > maxSize) return { valid: false, error: '文件过大 (' + (file.size / 1024 / 1024).toFixed(1) + 'MB > ' + (maxSize / 1024 / 1024) + 'MB)', type: null };
            var type = this.detectFileType(file);
            if (type === 'unknown') return { valid: false, error: '不支持的文件类型: ' + file.type, type: null };
            return { valid: true, error: null, type: type };
        },

        // 检测文件类型
        detectFileType: function(file) {
            var ftype = (file.type || '').toLowerCase();
            var fname = (file.name || '').toLowerCase();
            if (ftype.indexOf('image/') === 0) return 'image';
            if (ftype === 'application/pdf' || fname.indexOf('.pdf') !== -1) return 'pdf';
            if (ftype.indexOf('excel') !== -1 || ftype.indexOf('spreadsheet') !== -1 || fname.indexOf('.xls') !== -1 || fname.indexOf('.xlsx') !== -1) return 'excel';
            if (ftype.indexOf('word') !== -1 || ftype.indexOf('document') !== -1 || fname.indexOf('.doc') !== -1 || fname.indexOf('.docx') !== -1) return 'document';
            if (ftype.indexOf('text/') === 0 || fname.indexOf('.txt') !== -1 || fname.indexOf('.csv') !== -1) return 'text';
            return 'unknown';
        },

        // 提取文件内容（统一接口）
        extractContent: async function(file, base64Data, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 提取内容: ' + file.name);
            var fileType = this.detectFileType(file);
            var result = { name: file.name, type: fileType, base64Data: base64Data, size: file.size, content: null, metadata: {} };
            switch (fileType) {
                case 'image': result.content = await this._extractImageContent(file, base64Data, options); break;
                case 'pdf':
                    // 使用统一文档解析器处理 PDF
                    try {
                        var pdfResult = await App.DocumentProcessor.processDocument(file, {
                            onProgress: function(step, msg) { ChatLog.info('FILEBATCH', '[PDF]', step, msg); }
                        });
                        result.content = '[PDF文档] ' + file.name +
                            '\n字符数: ' + pdfResult.totalChars +
                            '\n页数: ' + pdfResult.totalPages +
                            '\n\n' + (pdfResult.fullText || '').substring(0, 50000);
                        result.metadata.docParseResult = pdfResult;
                        result.metadata.isLongDoc = pdfResult.totalChars > 3000;
                    } catch (e) {
                        ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] PDF 统一解析失败，使用降级模式:', e.message);
                        result.content = '[PDF文件] ' + file.name + '\n大小: ' + (file.size / 1024).toFixed(1) + ' KB\n注: PDF内容需要专门解析';
                    }
                    break;
                case 'excel':
                    // Excel 文件优先使用统一文档解析器（支持多 Sheet 分析）
                    if (window.excelBridge && typeof window.excelBridge.ParseDocumentAsync === 'function') {
                        try {
                            var xlsResult = await App.DocumentProcessor.processDocument(file);
                            // 【P0修复】使用结构化摘要替代 fullText 粗暴截断
                            result.content = buildExcelStructuredSummary(xlsResult, file.name);
                            result.metadata.docParseResult = xlsResult;
                        } catch (e) {
                            ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] Excel 统一解析失败，使用降级模式:', e.message);
                            result.content = await this._extractExcelContent(base64Data, file.name);
                        }
                    } else {
                        result.content = await this._extractExcelContent(base64Data, file.name);
                    }
                    break;
                case 'document':
                    // Word 文档使用统一文档解析器
                    if (window.excelBridge && typeof window.excelBridge.ParseDocumentAsync === 'function') {
                        try {
                            var docResult = await App.DocumentProcessor.processDocument(file);
                            result.content = '[Word文档] ' + file.name +
                                '\n字符数: ' + docResult.totalChars +
                                '\n页数: ' + docResult.totalPages +
                                '\n\n' + (docResult.fullText || '').substring(0, 50000);
                            result.metadata.docParseResult = docResult;
                            result.metadata.isLongDoc = docResult.totalChars > 3000;
                        } catch (e) {
                            ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] Word 解析失败:', e.message);
                            result.content = '[Word文件] ' + file.name + '\n大小: ' + (file.size / 1024).toFixed(1) + ' KB';
                        }
                    } else {
                        result.content = '[Word文件] ' + file.name + '\n大小: ' + (file.size / 1024).toFixed(1) + ' KB';
                    }
                    break;
                case 'text': result.content = await this._extractTextContent(base64Data); break;
                default: result.content = '[' + fileType.toUpperCase() + '文件] ' + file.name + '\n大小: ' + (file.size / 1024).toFixed(1) + ' KB';
            }
            return result;
        },

        // 智能检测图片类型（基于文件名）
        detectImageType: function(fileName) {
            var lowerName = (fileName || '').toLowerCase();
            var typePatterns = [
                { pattern: ['身份证','idcard','id_card'], type: '身份证', keyword: '身份证' },
                { pattern: ['名片','card'], type: '名片', keyword: '名片' },
                { pattern: ['发票','invoice','fapiao'], type: '发票', keyword: '发票' },
                { pattern: ['银行卡','bank'], type: '银行卡', keyword: '银行卡' },
                { pattern: ['营业执照','license','bizlicense'], type: '营业执照', keyword: '营业执照' },
                { pattern: ['表格','excel','sheet','table'], type: '表格', keyword: '表格' }
            ];
            for (var pi = 0; pi < typePatterns.length; pi++) {
                var p = typePatterns[pi];
                var matched = false;
                for (var pj = 0; pj < p.pattern.length; pj++) {
                    if (lowerName.indexOf(p.pattern[pj]) !== -1) { matched = true; break; }
                }
                if (matched) { ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 图片类型检测: ' + p.type); return { type: p.type, keyword: p.keyword }; }
            }
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 使用默认类型: 通用图片');
            return { type: '通用', keyword: '' };
        },

        // 处理单文件（优化流程）
        _processSingleFile: async function(file, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 开始单文件优化流程: ' + file.name);
            var validation = this.validateFile(file);
            if (!validation.valid) throw new Error(validation.error);
            var fileType = validation.type;
            if (fileType === 'image') await this._processSingleImage(file, options);
            else await this._processSingleNonImage(file, options);
        },

        // 处理单张图片
        _processSingleImage: async function(file, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 处理单张图片: ' + file.name);
            var updateProgress = options.showProgress ? this._createProgressUpdater() : null;
            try {
                if (updateProgress) updateProgress(1, '文件名检测完成', '检查文件名关键词: ' + file.name);
                var imageType = this.detectImageType(file.name);
                if (updateProgress) updateProgress(2, '准备图片数据...', '正在编码图片...');
                var base64Data = await fileToBase64(file);
                if (updateProgress) updateProgress(3, 'OCR文字识别中', '使用OCR引擎提取文字...');
                var analysisResult = await quickAnalyzeImage(base64Data, file.name);
                if (updateProgress) updateProgress(4, '分析完成', '检测类型: ' + analysisResult.imageType);
                file.aiThinking = analysisResult.description;
                file.detectedType = analysisResult.imageType;
                var fileContext = { file: file, base64Data: base64Data, fileType: file.type, fileName: file.name, detectedType: analysisResult.imageType, analysisResult: analysisResult };
                if (options.onComplete) options.onComplete(fileContext);
                return fileContext;
            } catch (error) {
                ChatLog.error('FILEBATCH', '[UnifiedFileProcessor] 单张图片处理失败:', error);
                if (updateProgress) updateProgress(4, '分析失败', error.message);
                throw error;
            }
        },

        // 处理单个非图片文件
        _processSingleNonImage: async function(file, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 处理非图片文件: ' + file.name);
            var base64Data = await fileToBase64(file);
            var content = await this.extractContent(file, base64Data);
            var fileContext = { file: file, base64Data: base64Data, fileType: file.type, fileName: file.name, content: content.content };
            if (options.onComplete) options.onComplete(fileContext);
            return fileContext;
        },

        // 处理多个文件（批量流程）
        _processMultipleFiles: async function(files, options) {
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 开始批量处理 ' + files.length + ' 个文件');
            var maxConcurrent = options.maxConcurrent || 3;
            var results = [];
            var failedFiles = [];
            var imageCount = 0;
            for (var fi = 0; fi < files.length; fi++) { if (files[fi].type.indexOf('image/') === 0) imageCount++; }
            var otherCount = files.length - imageCount;
            if (options.onBatchStart) options.onBatchStart({ total: files.length, imageCount: imageCount, otherCount: otherCount });
            for (var i = 0; i < files.length; i++) {
                var file = files[i];
                if (options.onProgress) options.onProgress(i + 1, files.length, file.name, 'processing');
                try {
                    var validation = this.validateFile(file);
                    if (!validation.valid) throw new Error(validation.error);
                    var base64Data = await fileToBase64(file);
                    var content = await this.extractContent(file, base64Data, options);
                    results.push({ success: true, data: content, fileName: file.name });
                    if (options.onProgress) options.onProgress(i + 1, files.length, file.name, 'completed');
                } catch (error) {
                    ChatLog.error('FILEBATCH', '[UnifiedFileProcessor] 文件处理失败 [' + file.name + ']:', error);
                    failedFiles.push({ name: file.name, error: error.message });
                    results.push({ success: false, error: error.message, fileName: file.name });
                    if (options.onProgress) options.onProgress(i + 1, files.length, file.name, 'failed');
                }
            }
            var summary = { total: files.length, successful: 0, failed: failedFiles.length, results: results, failedFiles: failedFiles };
            for (var ri = 0; ri < results.length; ri++) { if (results[ri].success) summary.successful++; }
            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 批量处理完成: 成功=' + summary.successful + ', 失败=' + summary.failed);
            if (options.onBatchComplete) options.onBatchComplete(summary);
            return summary;
        },

        // 创建进度更新器
        _createProgressUpdater: function() {
            var wrapperId = null;
            return function(step, status, detail) {
                detail = detail || '';
                var progressHtml =
                    '<div style="background:#f8f9fa;padding:12px;border-radius:8px;border:1px solid #e9ecef;">' +
                    '<div style="font-size:13px;color:#495057;font-weight:600;margin-bottom:8px;">图片分析进度</div>' +
                    '<div style="display:flex;flex-direction:column;gap:6px;">' +
                    '<div style="display:flex;align-items:center;gap:8px;font-size:12px;"><span style="color:' + (step>=1?'#28a745':'#6c757d') + ';">' + (step>=1?'✓':'○') + '</span><span style="color:' + (step>=1||step===0?'#28a745':step===0?'#007bff':'#6c757d') + ';font-weight:' + (step===0?'600':'normal') + ';">1. 文件名检测</span>' + (step===0?'<span style="color:#007bff;">进行中</span>':'') + '</div>' +
                    '<div style="display:flex;align-items:center;gap:8px;font-size:12px;"><span style="color:' + (step>=2?'#28a745':'#6c757d') + ';">' + (step>=2?'✓':'○') + '</span><span style="color:' + (step>=2||step===1?'#28a745':step===1?'#007bff':'#6c757d') + ';font-weight:' + (step===1?'600':'normal') + ';">2. OCR 文字识别</span>' + (step===1?'<span style="color:#007bff;">进行中</span>':'') + '</div>' +
                    '<div style="display:flex;align-items:center;gap:8px;font-size:12px;"><span style="color:' + (step>=3?'#28a745':'#6c757d') + ';">' + (step>=3?'✓':'○') + '</span><span style="color:' + (step>=3||step===2?'#28a745':step===2?'#007bff':'#6c757d') + ';font-weight:' + (step===2?'600':'normal') + ';">3. AI 视觉分析</span>' + (step===2?'<span style="color:#007bff;">进行中</span>':'') + '</div>' +
                    '<div style="display:flex;align-items:center;gap:8px;font-size:12px;"><span style="color:' + (step>=4?'#28a745':'#6c757d') + ';">' + (step>=4?'✓':'○') + '</span><span style="color:' + (step>=4||step===3?'#28a745':step===3?'#007bff':'#6c757d') + ';font-weight:' + (step===3?'600':'normal') + ';">4. 结果汇总</span>' + (step===3?'<span style="color:#007bff;">进行中</span>':'') + '</div>' +
                    '</div>' + (detail ? '<div style="margin-top:8px;padding:8px;background:#e7f3ff;border-radius:4px;font-size:11px;color:#004085;line-height:1.6;">' + detail + '</div>' : '') +
                    '</div>';
                if (!wrapperId) { wrapperId = addMessage('ai', progressHtml); }
                else { updateMessage(wrapperId, progressHtml); }
            };
        },

        // 提取图片内容（内部实现）
        _extractImageContent: async function(file, base64Data, options) {
            var imageType = this.detectImageType(file.name);
            if (options.forcedType) {
                try {
                    var operationType = mapForcedType(options.forcedType);
                    var result = await recognizeImageOnly(base64Data, file.name, operationType);
                    if (result && result.data) {
                        var simplifiedData = {};
                        var data = result.data;
                        if (data.rawText) simplifiedData.rawText = data.rawText.substring(0, 800) + (data.rawText.length > 800 ? '...' : '');
                        else if (data.description) simplifiedData.description = data.description.substring(0, 800);
                        var keyFields = ['姓名','性别','民族','出生日期','住址','公民身份号码','身份证号码','签发机关','有效期限'];
                        for (var ki = 0; ki < keyFields.length; ki++) { if (data[keyFields[ki]]) simplifiedData[keyFields[ki]] = data[keyFields[ki]]; }
                        var ocrEngine = data.ocrEngine || (result.rawResponse && result.rawResponse.ocrEngine) || 'PaddleOCR';
                        var engineLabel = result.usedVisionFallback ? 'ImageFieldExtractor 纯视觉兜底' : (ocrEngine + '识别');
                        var formattedRawText = simplifiedData.rawText || simplifiedData.description || '';
                        formattedRawText = formattedRawText.replace(/\\r\\n/g, '<br>').replace(/\\n/g, '<br>').replace(/\\r/g, '<br>');
                        var resultHtml = '<div style="margin-bottom:8px;"><strong>[' + engineLabel + ']</strong> ' + file.name + '</div>';
                        resultHtml += '<div style="margin-bottom:4px;"><span style="color:#666;">类型:</span> ' + options.forcedType + '</div>';
                        resultHtml += '<div style="margin-top:8px;padding:10px;background:#f5f5f5;border-radius:6px;font-size:13px;line-height:1.6;">' + formattedRawText + '</div>';
                        return resultHtml;
                    }
                } catch (recognizeError) {
                    return '[图片文件] ' + file.name + '\n类型: ' + options.forcedType + '\n注：OCR识别失败，建议使用多模态模型直接分析';
                }
            }
            return '[图片文件] ' + file.name + '\n大小: ' + (file.size / 1024).toFixed(1) + ' KB\n格式: ' + file.name.split('.').pop().toUpperCase() + '\n检测类型: ' + imageType.type + '\n注：图片内容已通过 OCR 识别';
        },

        // 提取Excel内容（内部实现）
        _extractExcelContent: async function(base64Data, fileName) {
            try {
                // 剥离 DataURL 前缀（data:...;base64,），只保留纯 Base64 数据
                var pureBase64 = base64Data.indexOf(',') > -1 ? base64Data.substring(base64Data.indexOf(',') + 1) : base64Data;
                var result = null;
                // 优先使用 ProcessExcelFileAsync（ExcelBridge 专用入口）
                if (window.excelBridge) {
                    try {
                        if (typeof window.excelBridge.ProcessExcelFileAsync === 'function') {
                            result = await window.excelBridge.ProcessExcelFileAsync(pureBase64, null);
                        }
                    } catch (e) {
                        ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] ProcessExcelFileAsync 调用失败，尝试 fallback:', e.message);
                    }
                }
                // Fallback: 使用 ParseDocumentAsync 通用文档解析入口
                if (!result && window.excelBridge) {
                    try {
                        if (typeof window.excelBridge.ParseDocumentAsync === 'function') {
                            result = await window.excelBridge.ParseDocumentAsync(fileName || 'uploaded.xlsx', pureBase64);
                        }
                    } catch (e) {
                        ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] ParseDocumentAsync fallback 也失败:', e.message);
                    }
                }
                // 检查 bridge 结果是否有效
                var bridgeSuccess = false;
                if (result) {
                    try {
                        var parsed = JSON.parse(result);
                        if (parsed.success) {
                            bridgeSuccess = true;
                            return buildExcelStructuredSummary(parsed, fileName);
                        } else {
                            ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] Bridge 返回失败:', parsed.error);
                        }
                    } catch (parseErr) {
                        ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] Bridge 结果解析失败:', parseErr.message);
                    }
                }
                
                // 【Web版 fallback】Bridge 不可用或失败，使用 SheetJS 客户端解析
                ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] 使用 SheetJS 客户端解析 Excel');
                if (typeof XLSX !== 'undefined') {
                    try {
                        // 从 App.selectedFiles / App.selectedFile 中找 File 对象
                        var fileObj = null;
                        var allFiles = App.selectedFiles || [];
                        if (allFiles.length > 0) {
                            fileObj = allFiles.find(function(f) { return f && f.name === fileName; });
                        }
                        if (!fileObj && App.selectedFile && App.selectedFile.name === fileName) {
                            fileObj = App.selectedFile;
                        }
                        
                        if (fileObj) {
                            var buf = await fileObj.arrayBuffer();
                            var bytes = new Uint8Array(buf);
                            var workbook = XLSX.read(bytes, { type: 'array' });
                            
                            var pages = [];
                            var fullTextParts = [];
                            workbook.SheetNames.forEach(function(sheetName) {
                                var worksheet = workbook.Sheets[sheetName];
                                var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                                var sheetLines = [];
                                sheetLines.push('【工作表: ' + sheetName + '】');
                                jsonData.forEach(function(row, rowIdx) {
                                    var rowText = row.map(function(cell) { return String(cell != null ? cell : ''); }).join('\t');
                                    if (rowIdx === 0) {
                                        sheetLines.push('【表头】' + rowText);
                                    } else {
                                        sheetLines.push(rowText);
                                    }
                                });
                                var sheetText = sheetLines.join('\n');
                                pages.push({ title: sheetName, text: sheetText });
                                fullTextParts.push(sheetText);
                            });
                            
                            var resultObj = {
                                success: true,
                                fileType: 'excel',
                                fullText: fullTextParts.join('\n\n'),
                                totalChars: fullTextParts.join('\n\n').length,
                                totalPages: pages.length,
                                isScanned: false,
                                pages: pages,
                                structure: { Headings: [], Tables: [] },
                                metadata: {
                                    fileFormat: (fileName || '').split('.').pop() || 'xlsx',
                                    properties: { '工作表列表': workbook.SheetNames.join(', ') }
                                }
                            };
                            
                            ChatLog.info('FILEBATCH', '[UnifiedFileProcessor] SheetJS 解析成功:', workbook.SheetNames.length, '个工作表');
                            return buildExcelStructuredSummary(resultObj, fileName);
                        } else {
                            ChatLog.warn('FILEBATCH', '[UnifiedFileProcessor] 未找到 File 对象:', fileName);
                        }
                    } catch (sheetjsErr) {
                        ChatLog.error('FILEBATCH', '[UnifiedFileProcessor] SheetJS 解析失败:', sheetjsErr);
                    }
                }
                return '[Excel文件] ' + fileName + '\n注：Excel 解析接口不可用，请检查插件是否正常加载';
            } catch (error) {
                return '[Excel文件] ' + fileName + '\n提取失败: ' + error.message;
            }
        },

        // 提取文本内容（内部实现）
        _extractTextContent: async function(base64Data) {
            try {
                var text = atob(base64Data.split(',')[1]);
                return text.substring(0, 1000) + (text.length > 1000 ? '...(截断)' : '');
            } catch (error) {
                return '[文本文件]\n读取失败: ' + error.message;
            }
        }
    };

    // 挂载到全局和 App 命名空间
    window.UnifiedFileProcessor = UnifiedFileProcessor;
    App.UnifiedFileProcessor = UnifiedFileProcessor;
    // 暴露 Excel 结构化摘要构建器，供 sendMessage 等模块使用
    window.buildExcelStructuredSummary = buildExcelStructuredSummary;
    if (typeof App !== 'undefined') App.buildExcelStructuredSummary = buildExcelStructuredSummary;

    // ============================================================
    // 附件列表渲染与删除
    // ============================================================

    /**
     * 统一更新附件状态变量（单文件/多文件模式自动切换）
     * @param {File[]} files - 剩余文件数组
     */
    function updateAttachmentState(files) {
        if (!files || files.length === 0) {
            selectedFile = null;
            selectedFiles = [];
            App.selectedFile = null;
            App.selectedFiles = [];
            window.__backupSelectedFile = null;
            window.__backupSelectedFiles = [];
        } else if (files.length === 1) {
            selectedFile = files[0];
            selectedFiles = [];
            App.selectedFile = files[0];
            App.selectedFiles = [];
            window.__backupSelectedFile = files[0];
            window.__backupSelectedFiles = [];
        } else {
            selectedFile = null;
            selectedFiles = files;
            App.selectedFile = null;
            App.selectedFiles = files;
            window.__backupSelectedFile = null;
            window.__backupSelectedFiles = files;
        }
    }

    /**
     * 渲染附件列表到 #attachment-list
     * @param {File[]} files - 文件数组
     */
    function renderAttachmentList(files) {
        const list = document.getElementById('attachment-list');
        if (!list) {
            ChatLog.warn('FILEBATCH', 'renderAttachmentList: attachment-list 元素不存在');
            return;
        }

        list.innerHTML = '';

        if (!files || files.length === 0) {
            list.style.display = 'none';
            return;
        }

        files.forEach((file, index) => {
            const fileType = getFileType(file);
            const icon = getFileIcon(file);
            const label = getFileTypeLabel(fileType);
            const item = document.createElement('div');
            item.className = 'attachment-item';
            item.innerHTML = `
                <span class="attachment-icon">${icon}</span>
                <div class="attachment-info">
                    <span class="attachment-name" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</span>
                    <div class="attachment-meta">
                        <span class="attachment-size">${formatFileSize(file.size)}</span>
                        <span class="attachment-type-badge type-${fileType}">${label}</span>
                    </div>
                </div>
                <button class="attachment-remove" onclick="App.removeAttachment(${index})" title="移除">✕</button>
            `;
            list.appendChild(item);
        });

        list.style.display = 'flex';
    }

    /**
     * 删除指定索引的附件
     * @param {number} index - 附件索引
     */
    function removeAttachment(index) {
        ChatLog.info('FILEBATCH', `[removeAttachment] 请求删除索引: ${index}`);

        // 收集当前所有文件：优先从 App 命名空间读取（chat-entry / app-image-table 设置此处），
        // 再回退到全局 selectedFile / selectedFiles
        let allFiles = [];
        if (App.selectedFiles && App.selectedFiles.length > 0) {
            allFiles = Array.from(App.selectedFiles);
        } else if (App.selectedFile) {
            allFiles = [App.selectedFile];
        } else if (selectedFiles && selectedFiles.length > 0) {
            allFiles = Array.from(selectedFiles);
        } else if (selectedFile) {
            allFiles = [selectedFile];
        }

        if (index < 0 || index >= allFiles.length) {
            ChatLog.warn('FILEBATCH', `[removeAttachment] 索引越界: ${index}/${allFiles.length}`);
            return;
        }

        allFiles.splice(index, 1);

        if (allFiles.length === 0) {
            clearAttachment();
            return;
        }

        updateAttachmentState(allFiles);

        const preview = document.getElementById('attachment-preview');
        if (preview) {
            preview.dataset.fileCount = String(allFiles.length);
            preview._storedFiles = allFiles;
        }

        // 重新渲染列表
        renderAttachmentList(allFiles);

        // 单文件时显示对应操作按钮；多文件时显示汇总提示
        const attachmentActions = document.getElementById('attachment-actions');
        if (allFiles.length === 1) {
            const remaining = allFiles[0];
            if (remaining.type && remaining.type.startsWith('image/')) {
                if (attachmentActions) {
                    attachmentActions.innerHTML = '<span style="color: #666; font-size: 12px;">💡 点击"发送"按钮开始分析图片</span>';
                }
            } else {
                renderAttachmentActions(remaining);
            }
        } else {
            if (attachmentActions) {
                attachmentActions.innerHTML = `<div style="color: #666; font-size: 12px; padding: 8px; background: #f5f5f5; border-radius: 6px; margin-top: 8px;">
                    💡 点击"发送"将逐个处理这些文件
                </div>`;
            }
        }

        ChatLog.info('FILEBATCH', `[removeAttachment] 已删除索引 ${index}，剩余 ${allFiles.length} 个文件`);
    }

    // --- 来自 app.js:L6040-L6193 ---
    async function handleFileSelect(event) {
        ChatLog.info('FILEBATCH', '🎯 [handleFileSelect] 函数被调用！', {
            event: event,
            target: event?.target,
            files: event?.target?.files,
            filesLength: event?.target?.files?.length
        });

        const files = Array.from(event.target.files);
        ChatLog.info('FILEBATCH', '📁 [handleFileSelect] 解析到文件列表:', files.length, '个文件', files);

        if (files.length === 0) {
            ChatLog.warn('FILEBATCH', '⚠️ [handleFileSelect] 文件列表为空，退出');
            return;
        }

        // 🚀 【方案A重构】使用统一文件处理器进行预验证
        ChatLog.info('FILEBATCH', `[handleFileSelect] 🔍 使用 UnifiedFileProcessor 预验证 ${files.length} 个文件`);

        const validationResults = [];
        let hasInvalidFile = false;

        for (const file of files) {
            const validation = window.UnifiedFileProcessor.validateFile(file);
            validationResults.push({ file, validation });

            if (!validation.valid) {
                ChatLog.error('FILEBATCH', `[handleFileSelect] ❌ 文件验证失败: ${file.name}`, validation.error);
                hasInvalidFile = true;
            }
        }

        if (hasInvalidFile && files.length === 1) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + validationResults[0].validation.error + '</span>', true) : showToast(`❌ ${validationResults[0].validation.error}`);
            return;
        } else if (hasInvalidFile) {
            const invalidNames = validationResults
                .filter(r => !r.validation.valid)
                .map(r => r.file.name)
                .join(', ');
            ChatLog.warn('FILEBATCH', `[handleFileSelect] ⚠️ 部分文件验证失败: ${invalidNames}`);
            // 多文件情况下只警告，不阻止（后续处理时会跳过无效文件）
        }

        if (files.length === 1) {
            // 单文件：智能检测后显示操作按钮（与粘贴方式统一）
            const file = files[0];
            const validation = validationResults[0].validation;

            selectedFile = file;
            selectedFiles = []; // 清空多文件列表
            App.selectedFile = file;   // 同步到App命名空间（sendMessage读取此处）
            App.selectedFiles = [];    // 同步到App命名空间

            // 【关键修复】全局备份，防止变量丢失
            window.__backupSelectedFile = file;
            window.__backupSelectedFiles = [];
            ChatLog.info('FILEBATCH', '💾 [文件备份] 已保存到全局:', file.name);

            ChatLog.info('FILEBATCH', '✅ [handleFileSelect] 设置单文件:', file.name,
                'selectedFile已更新:', !!selectedFile,
                '检测类型:', validation.type);

            // 显示附件预览区域
            const preview = document.getElementById('attachment-preview');
            if (!preview) {
                ChatLog.warn('FILEBATCH', '[handleFileSelect] attachment-preview 元素不存在，跳过预览显示');
                return;
            }

            // 隐藏旧版单文件 header 元素（已改用 attachment-list）
            const legacyHeader = preview.querySelector('.attachment-header');
            if (legacyHeader) legacyHeader.style.display = 'none';

            // 渲染统一附件列表（始终显示，含删除按钮）
            renderAttachmentList([file]);

            // 图片类型检测
            if (file.type.startsWith('image/')) {
                const thumbnail = document.getElementById('attachment-thumbnail');
                const thumbnailImg = document.getElementById('thumbnail-img');
                const analysisStatus = document.getElementById('attachment-analysis-status');

                // 单图片显示缩略图
                if (thumbnail && thumbnailImg) {
                    try { URL.revokeObjectURL(thumbnailImg.src); } catch(e) { (window.ChatLog || console).warn?.('UI', '[file-batch] 撤销缩略图 objectURL 失败:', e); }
                    thumbnailImg.src = URL.createObjectURL(file);
                    thumbnail.style.display = 'flex';
                }
                if (analysisStatus) analysisStatus.style.display = 'none';

                const imageTypeInfo = window.UnifiedFileProcessor.detectImageType(file.name);
                file.detectedType = imageTypeInfo.type;
                file.detectedKeyword = imageTypeInfo.keyword;
            } else {
                const thumbnail = document.getElementById('attachment-thumbnail');
                const analysisStatus = document.getElementById('attachment-analysis-status');
                if (thumbnail) thumbnail.style.display = 'none';
                if (analysisStatus) analysisStatus.style.display = 'none';
            }

            preview.style.display = 'flex';

            // 【关键】将File对象引用存储到DOM元素上（单文件模式）
            preview.dataset.fileCount = '1';
            preview._storedFiles = [file];
            ChatLog.info('FILEBATCH', '💾 [handleFileSelect] 单文件已存储到DOM: _storedFiles[0].name =', file.name);

            //
            // 最高优先级：全局锁定上下文（triggerQuickSkill 设置，不被 clearAttachment 影响）
            var triggerCtx = window.__TRIGGER_QUICK_SKILL_CTX || null;
            ChatLog.info('FILEBATCH', '[handleFileSelect] 🔍 调试: __TRIGGER_QUICK_SKILL_CTX =', JSON.stringify(triggerCtx), ', App._pendingImageSkill =', App._pendingImageSkill ? App._pendingImageSkill.name : null, ', App.activeMode =', App.activeMode);
            var pendingSkill = null;
            var pendingMsg = '';

            if (triggerCtx && triggerCtx.skill && (Date.now() - triggerCtx.timestamp < 120000)) {
                ChatLog.info('FILEBATCH', '[handleFileSelect] 🔄 全局锁定命中:', triggerCtx.skill.name);
                pendingSkill = triggerCtx.skill;
                pendingMsg = triggerCtx.message;
                window.__TRIGGER_QUICK_SKILL_CTX = null;
            } else if (App.activeMode === 'image' && App._pendingImageSkill) {
                ChatLog.info('FILEBATCH', '[handleFileSelect] 🔄 旧路径待处理技能:', App._pendingImageSkill.name);
                pendingSkill = App._pendingImageSkill;
                pendingMsg = App._pendingImageSkillMessage || '';
                App._pendingImageSkill = null;
                App._pendingImageSkillMessage = '';
                App._pendingImageSkillDirectTrigger = false;
                App.activeMode = null;
            }

            if (pendingSkill) {
                ChatLog.info('FILEBATCH', '[handleFileSelect] ✅ 技能直连模式生效, skill=', pendingSkill.name);
                window.__DIRECT_IMAGE_FILE = file;
                window.__DIRECT_IMAGE_SKILL = pendingSkill;
                window.__DIRECT_IMAGE_MSG = pendingMsg;

                if (window.App && typeof window.App.executeImageProcessingSkill === 'function') {
                    window.App.executeImageProcessingSkill(pendingSkill, pendingMsg);
                    return;
                }
                if (window.App && typeof window.App.executeCustomSkill === 'function') {
                    window.App.executeCustomSkill(pendingSkill, pendingMsg);
                    return;
                }
                // 【修复】上面两个"直连执行"函数在前端并未定义（历史遗留），原先这里直接 return：
                // 结果是"选了文件但预览/状态都不更新"，且技能身份也丢了。
                // 改为：绑定技能身份后继续走正常流程 —— 用户点发送时由
                // App._quickSkillBinding → context._forcedSkillName → orchestrator【快捷入口直连钉死】
                // 把这次请求钉在该技能上，不再被别的技能匹配走。
                ChatLog.warn('FILEBATCH', '[handleFileSelect] 直连执行函数不存在 → 转为绑定技能身份，走正常发送流程');
                if (window.App && typeof window.App.bindQuickSkill === 'function') {
                    window.App.bindQuickSkill(pendingSkill, pendingMsg || '');
                }
            }


            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已选择: ' + file.name + '</span>', true) : showToast('已选择: ' + file.name);

            // 图片不再自动分析，分析过程移到发送后
            if (file.type.startsWith('image/')) {
                // 只显示简单提示，等待用户发送后分析
                document.getElementById('attachment-actions').innerHTML =
                    '<span style="color: #666; font-size: 12px;">💡 点击"发送"按钮开始分析图片</span>';
            } else {
                // 非图片文件，显示操作按钮等待用户处理
                renderAttachmentActions(file);
            }
        } else {
            // 多文件：像单文件一样处理，但显示文件数量
            selectedFiles = files;  // 存储多个文件
            selectedFile = null;    // 清空单文件
            App.selectedFiles = files;   // 同步到App命名空间
            App.selectedFile = null;     // 同步到App命名空间

            // 【关键修复】全局备份多文件列表
            window.__backupSelectedFile = null;
            window.__backupSelectedFiles = files;
            ChatLog.info('FILEBATCH', '💾 [文件备份] 已保存多文件到全局:', files.length, '个');

            // 显示附件预览区域（简化版）
            const preview = document.getElementById('attachment-preview');

            // 隐藏旧版单文件 header 元素
            const legacyHeader = preview.querySelector('.attachment-header');
            if (legacyHeader) legacyHeader.style.display = 'none';

            // 渲染统一附件列表（显示全部文件，每项可删除）
            renderAttachmentList(files);

            preview.style.display = 'flex';

            // 【关键】将File对象引用存储到DOM元素上（终极可靠数据源）
            preview.dataset.fileCount = files.length;
            preview._storedFiles = files;
            ChatLog.info('FILEBATCH', '💾 [handleFileSelect] 多文件已存储到DOM: _storedFiles.length =', files.length);

            // 多文件模式下，隐藏检测状态区域和缩略图
            document.getElementById('attachment-analysis-status').style.display = 'none';
            document.getElementById('attachment-thumbnail').style.display = 'none';

            // 多文件模式下，显示简洁的操作提示
            const attachmentActions = document.getElementById('attachment-actions');
            attachmentActions.innerHTML = `<div style="color: #666; font-size: 12px; padding: 8px; background: #f5f5f5; border-radius: 6px; margin-top: 8px;">
                💡 点击"发送"将逐个处理这些文件${hasInvalidFile ? '<br>⚠️ 部分文件可能无法处理' : ''}
            </div>`;

            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已选择 ' + files.length + ' 个文件</span>', true) : showToast(`已选择 ${files.length} 个文件`);
        }
    }

    // --- 来自 app.js:L6199-L6210 ---
    function clearAttachment() {
        ChatLog.info('FILEBATCH', '🗑️ [clearAttachment] 被调用！清空附件变量');
        ChatLog.info('FILEBATCH', '   清空前 - selectedFile:', !!selectedFile, 'selectedFiles:', selectedFiles?.length);

        selectedFile = null;
        selectedFiles = [];  // 清空批量文件列表
        App.selectedFile = null;
        App.selectedFiles = [];
        window.selectedFile = null;
        window.selectedFiles = [];

        // 彻底清除所有文件备份来源，防止旧图片在后续消息中恢复
        window.__backupSelectedFile = null;
        window.__backupSelectedFiles = [];
        window.__DIRECT_IMAGE_FILE = null;
        window.__editingImageBase64 = null;

        const fileInput = document.getElementById('file-input');
        if (fileInput) fileInput.value = '';

        const previewEl = document.getElementById('attachment-preview');
        if (previewEl) {
            previewEl.style.display = 'none';
            previewEl._storedFiles = [];
        }

        // 清空附件列表
        const attachmentList = document.getElementById('attachment-list');
        if (attachmentList) attachmentList.innerHTML = '';

        // 释放单图片缩略图对象 URL
        const thumbnailImg = document.getElementById('thumbnail-img');
        if (thumbnailImg && thumbnailImg.src) {
            try { URL.revokeObjectURL(thumbnailImg.src); } catch(e) { (window.ChatLog || console).warn?.('UI', '[file-batch] 清理缩略图 objectURL 失败:', e); }
            thumbnailImg.src = '';
        }
        const thumbnail = document.getElementById('attachment-thumbnail');
        if (thumbnail) thumbnail.style.display = 'none';

        const attachmentName = document.getElementById('attachment-name');
        const attachmentSize = document.getElementById('attachment-size');
        const attachmentActions = document.getElementById('attachment-actions');
        if (attachmentName) attachmentName.textContent = '';
        if (attachmentSize) attachmentSize.textContent = '';
        if (attachmentActions) attachmentActions.innerHTML = '';

        // 清除待处理技能上下文
        if (App._pendingImageSkill) {
            App._pendingImageSkill = null;
            App._pendingImageSkillMessage = '';
        }

        ChatLog.info('FILEBATCH', '✅ [clearAttachment] 附件及所有备份来源已清空');
    }

    // --- 来自 app.js:L5876-L5891 ---
    // 获取文件快捷操作按钮配置（非图片文件）
    function getFileActionButtons(file) {
        const actions = [];

        if (file.type.includes('excel') || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
            actions.push({ id: 'analyze-excel', label: '📊 分析表格', action: 'analyzeExcel', primary: true });
        } else if (file.type === 'application/pdf' || file.name.endsWith('.pdf')) {
            actions.push({ id: 'extract-pdf', label: '📄 提取PDF', action: 'extractPdf', primary: true });
        } else if (file.type.startsWith('text/') || file.name.endsWith('.txt') || file.name.endsWith('.csv')) {
            actions.push({ id: 'read-text', label: '📃 读取内容', action: 'readText', primary: true });
        }

        // 统一添加 AI 总结
        actions.push({ id: 'summarize', label: '📝 AI总结', action: 'summarize' });

        return actions;
    }

    // --- 来自 app.js:L5894-L5926 ---
    // 渲染附件操作按钮
    function renderAttachmentActions(file, detectedType, description, thinking) {
        const actionsContainer = document.getElementById('attachment-actions');

        // 检查 actionsContainer 是否存在
        if (!actionsContainer) {
            ChatLog.error('FILEBATCH', 'renderAttachmentActions: attachment-actions 元素不存在');
            return;
        }

        // 检查 file 参数
        if (!file) {
            ChatLog.warn('FILEBATCH', 'renderAttachmentActions: file 参数为 null');
            actionsContainer.innerHTML = '';
            return;
        }

        if (file.type && file.type.startsWith('image/')) {
            actionsContainer.innerHTML = `
                <div style="display: flex; flex-direction: column; gap: 8px;">
                    <div style="display: flex; align-items: center; gap: 8px; color: #4a90d9; font-size: 12px; font-weight: 500;">
                        <span>🖼️</span>
                        <span>图片已就绪，点击下方"发送"按钮开始分析</span>
                    </div>
                </div>
            `;
        } else {
            const buttons = getFileActionButtons(file);
            actionsContainer.innerHTML = buttons.map(btn =>
                `<button class="attachment-action-btn ${btn.primary ? 'primary' : ''}" ` +
                `onclick="handleAttachmentAction('${btn.action}')">${btn.label}</button>`
            ).join('');
        }
    }

    // --- 来自 app.js:L5929-L6000 ---

    // 快速分析图片内容（用于显示推荐按钮）
    async function quickAnalyzeImage(base64Image, fileName) {
        // 附件预览区域直接显示文件元信息，不再展示分析进度
        ChatLog.info('IMAGE', `🖼️ 开始智能识别图片: ${fileName}`);

        // 基于文件名的快速检测
        let detectedType = '通用';
        let detectedKeyword = '';

        const lowerFileName = (fileName || '').toLowerCase();

        // 身份证检测
        if (lowerFileName.includes('身份证') || lowerFileName.includes('idcard') || lowerFileName.includes('id_card')) {
            detectedType = '身份证';
            detectedKeyword = '身份证';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 身份证');
        }
        // 名片检测
        else if (lowerFileName.includes('名片') || lowerFileName.includes('card')) {
            detectedType = '名片';
            detectedKeyword = '名片';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 名片');
        }
        // 发票检测
        else if (lowerFileName.includes('发票') || lowerFileName.includes('invoice') || lowerFileName.includes('fapiao')) {
            detectedType = '发票';
            detectedKeyword = '发票';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 发票');
        }
        // 银行卡检测
        else if (lowerFileName.includes('银行卡') || lowerFileName.includes('bank') || lowerFileName.includes('card')) {
            detectedType = '银行卡';
            detectedKeyword = '银行卡';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 银行卡');
        }
        // 营业执照检测
        else if (lowerFileName.includes('营业执照') || lowerFileName.includes('license') || lowerFileName.includes('bizlicense')) {
            detectedType = '营业执照';
            detectedKeyword = '营业执照';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 营业执照');
        }
        // 表格/Excel截图检测
        else if (lowerFileName.includes('表格') || lowerFileName.includes('excel') || lowerFileName.includes('sheet') || lowerFileName.includes('table')) {
            detectedType = '表格';
            detectedKeyword = '表格';
            ChatLog.info('FILEBATCH', 'IMAGE ✅ 文件名检测到: 表格');
        }
        // 通用图片（默认）
        else {
            ChatLog.info('FILEBATCH', 'IMAGE ℹ️ 未从文件名识别出特定类型，使用默认: 通用图片');
        }

        ChatLog.info('IMAGE', `✅ 图片识别完成: ${detectedType} (${fileName})`);

        return {
            imageType: detectedType,
            description: `${detectedType}图片`,
            detectedKeyword: detectedKeyword,
            fileName: fileName
        };
    }

    // --- 来自 app.js:L6012-L6029 ---
    // 取消当前操作
    function cancelCurrentOperation() {
        if (currentAbortController) {
            currentAbortController.abort();
            ChatLog.info('FILEBATCH', '用户取消了操作');

            const progressEl = document.getElementById('file-progress');
            if (progressEl) {
                progressEl.innerHTML = `<span style="color: #f44336; font-weight: 500;">❌ 操作已取消</span>`;
            }

            addMessage('ai', '⏹️ 已停止处理');

            // 恢复按钮
            setTimeout(() => {
                restoreAttachmentButtons();
            }, 500);
        }
    }

    // --- 来自 app.js:L6032-L6038 ---
    // 恢复附件操作按钮
    function restoreAttachmentButtons() {
        if (selectedFile) {
            renderAttachmentActions(selectedFile);
        } else {
            document.getElementById('attachment-actions').innerHTML = '';
        }
    }

    // ==================== 批量处理功能 ====================

    // --- 来自 app.js:L6218-L6381 ---
    /**
     * 显示批量处理界面（简洁版）
     * @param {File[]} files - 文件数组
     */
    function showBatchProcessingUI(files) {
        const attachmentActions = document.getElementById('attachment-actions');
        const imageFiles = files.filter(f => f.type.startsWith('image/'));

        // 简洁的文件列表
        let fileListHtml = `<div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 12px; border-radius: 8px; margin-bottom: 12px;">
            <div style="font-size: 15px; font-weight: 600; margin-bottom: 6px;">📦 已选择 ${files.length} 个文件</div>
            <div style="font-size: 12px; opacity: 0.9;">
                🖼️ ${imageFiles.length} 张图片可识别 ${files.length - imageFiles.length > 0 ? '· ' + (files.length - imageFiles.length) + ' 个其他文件' : ''}
            </div>
        </div>`;

        // 只显示前5个文件，超出部分显示"+N个"
        const displayFiles = files.slice(0, 5);
        if (files.length > 5) {
            fileListHtml += `<div style="max-height: 120px; overflow-y: auto; background: #f8f9fa; padding: 8px; border-radius: 6px; font-size: 12px;">`;
        } else {
            fileListHtml += `<div style="max-height: 120px; overflow-y: auto; background: #f8f9fa; padding: 8px; border-radius: 6px; font-size: 12px;">`;
        }

        displayFiles.forEach((file) => {
            const icon = file.type.startsWith('image/') ? '🖼️' : '📎';
            fileListHtml += `<div style="padding: 3px 0; color: #333;">
                ${icon} ${escapeHtml(file.name)}
                <span style="color: #888; float: right;">${formatFileSize(file.size)}</span>
            </div>`;
        });

        if (files.length > 5) {
            fileListHtml += `<div style="padding: 4px 0; color: #666; text-align: center; font-style: italic;">
                ... 还有 ${files.length - 5} 个文件
            </div>`;
        }

        fileListHtml += '</div>';

        // 简洁的操作区：只有一个主按钮
        let buttonsHtml = `<div style="display: flex; gap: 8px; margin-top: 12px;">`;

        if (imageFiles.length > 0) {
            // 主要按钮：一键批量处理（识别+填入）
            buttonsHtml += `<button class="attachment-action-btn primary" onclick="startBatchProcessing()" style="
                flex: 1;
                padding: 10px 20px;
                background: linear-gradient(135deg, #4caf50 0%, #45a049 100%);
                color: white;
                border: none;
                border-radius: 8px;
                cursor: pointer;
                font-size: 14px;
                font-weight: 600;
                box-shadow: 0 2px 8px rgba(76, 175, 80, 0.3);
                transition: all 0.2s;
            " onmouseover="this.style.transform='translateY(-1px)'; this.style.boxShadow='0 4px 12px rgba(76, 175, 80, 0.4)'" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(76, 175, 80, 0.3)'">
                ✨ 开始批量处理 (${imageFiles.length})
            </button>`;

            // 次要按钮：取消
            buttonsHtml += `<button onclick="clearBatchSelection()" style="
                padding: 10px 16px;
                background: white;
                color: #666;
                border: 1px solid #ddd;
                border-radius: 8px;
                cursor: pointer;
                font-size: 13px;
                transition: all 0.2s;
            " onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='white'">
                ✕
            </button>`;
        } else {
            // 非图片文件：根据文件类型提供相应操作
            const excelFiles = files.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return ext === 'xlsx' || ext === 'xls';
            });
            const otherFiles = files.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return !f.type.startsWith('image/') && ext !== 'xlsx' && ext !== 'xls';
            });

            let actionButtons = '';

            if (excelFiles.length > 0) {
                actionButtons += `<button onclick="startBatchFileProcessing('excel')" style="
                    flex: 1;
                    padding: 10px 20px;
                    background: linear-gradient(135deg, #2196F3 0%, #1976D2 100%);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 600;
                    box-shadow: 0 2px 8px rgba(33, 150, 243, 0.3);
                    transition: all 0.2s;
                " onmouseover="this.style.transform='translateY(-1px)'; this.style.boxShadow='0 4px 12px rgba(33, 150, 243, 0.4)'" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(33, 150, 243, 0.3)'">
                    📊 分析 Excel 文件 (${excelFiles.length})
                </button>`;
            }

            if (otherFiles.length > 0) {
                actionButtons += `<button onclick="startBatchFileProcessing('other')" style="
                    flex: 1;
                    padding: 10px 20px;
                    background: linear-gradient(135deg, #9C27B0 0%, #7B1FA2 100%);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 14px;
                    font-weight: 600;
                    box-shadow: 0 2px 8px rgba(156, 39, 176, 0.3);
                    transition: all 0.2s;
                " onmouseover="this.style.transform='translateY(-1px)'; this.style.boxShadow='0 4px 12px rgba(156, 39, 176, 0.4)'" onmouseout="this.style.transform='translateY(0)'; this.style.boxShadow='0 2px 8px rgba(156, 39, 176, 0.3)'">
                    📝 分析文件 (${otherFiles.length})
                </button>`;
            }

            buttonsHtml = `<div style="display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap;">
                ${actionButtons}
                <button onclick="clearBatchSelection()" style="
                    padding: 10px 16px;
                    background: white;
                    color: #666;
                    border: 1px solid #ddd;
                    border-radius: 8px;
                    cursor: pointer;
                    font-size: 13px;
                    transition: all 0.2s;
                " onmouseover="this.style.background='#f5f5f5'" onmouseout="this.style.background='white'">
                    ✕ 取消
                </button>
            </div>`;
        }

        buttonsHtml += '</div>';

        // 更新附件预览区域
        var batchPreview = document.getElementById('attachment-preview');
        batchPreview.style.display = 'flex';

        // 隐藏旧版单文件 header
        const legacyHeader = batchPreview.querySelector('.attachment-header');
        if (legacyHeader) legacyHeader.style.display = 'none';

        // 渲染统一附件列表（显示全部文件，每项可删除）
        renderAttachmentList(files);

        // 【关键】将File对象存储到DOM元素上（批量处理模式）
        batchPreview.dataset.fileCount = files.length;
        batchPreview._storedFiles = Array.prototype.slice.call(files);
        ChatLog.info('FILEBATCH', '💾 [showBatchProcessingUI] 批量文件已存储到DOM:', files.length, '个');

        // 隐藏单图片缩略图（批量模式不显示）
        const thumbnail = document.getElementById('attachment-thumbnail');
        if (thumbnail) thumbnail.style.display = 'none';

        // 显示界面
        attachmentActions.innerHTML = fileListHtml + buttonsHtml;

        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 已选择 ' + files.length + ' 个文件</span>', true) : showToast(`✅ 已选择 ${files.length} 个文件`);
    }

    // --- 来自 app.js:L6386-L6424 ---
    /**
     * 一键开始批量处理（默认：智能插入模式）
     */
    async function startBatchProcessing(userMessage = '') {
        if (!selectedFiles || selectedFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有选择文件</span>', true) : showToast('❌ 没有选择文件');
            return;
        }

        const imageFiles = selectedFiles.filter(f => f.type.startsWith('image/'));
        if (imageFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可识别的图片文件</span>', true) : showToast('❌ 没有可识别的图片文件');
            return;
        }

        // 构建批量处理的消息，显示文件列表
        let fileListHtml = imageFiles.slice(0, 5).map((file, index) => {
            return `<div style="display: flex; align-items: center; gap: 8px; margin: 4px 0; font-size: 13px;">
                <span>🖼️</span>
                <span>${App.escapeHtml(file.name)}</span>
                <span style="color: #999; font-size: 11px;">${(file.size / 1024).toFixed(1)} KB</span>
            </div>`;
        }).join('');

        if (imageFiles.length > 5) {
            fileListHtml += `<div style="color: #999; font-size: 12px; margin-top: 4px;">...还有 ${imageFiles.length - 5} 个文件</div>`;
        }

        // 如果有用户消息，显示在顶部
        const messageHeader = userMessage ? `<div style="margin-bottom: 10px; font-weight: 500;">${userMessage}</div>` : '';

        const batchMessage = `📊 开始批量处理 ${imageFiles.length} 张图片
${messageHeader}
<div style="background: #f5f5f5; border-radius: 8px; padding: 10px; margin-top: 8px;">
    <div style="font-size: 12px; color: #666; margin-bottom: 6px;">文件列表：</div>
    ${fileListHtml}
</div>`;

        addMessage('user', batchMessage);

        await executeBatchProcessing(imageFiles, 'auto', userMessage);  // 默认智能插入，传递用户消息
    }

    // --- 来自 app.js:L6429-L6491 ---
    /**
     * 批量处理非图片文件（Excel、PDF、文本等）
     */
    async function startBatchFileProcessing(fileType) {
        if (!selectedFiles || selectedFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有选择文件</span>', true) : showToast('❌ 没有选择文件');
            return;
        }

        let filesToProcess = [];

        if (fileType === 'excel') {
            filesToProcess = selectedFiles.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return ext === 'xlsx' || ext === 'xls';
            });

            if (filesToProcess.length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有 Excel 文件</span>', true) : showToast('❌ 没有 Excel 文件');
                return;
            }

            addMessage('user', `📊 开始分析 ${filesToProcess.length} 个 Excel 文件...`);

            for (let i = 0; i < filesToProcess.length; i++) {
                const file = filesToProcess[i];
                try {
                    const base64Data = await fileToBase64(file);
                    await processFileWithAI(base64Data, file.name, 'summarize');
                } catch (error) {
                    ChatLog.error('FILEBATCH', `处理文件 ${file.name} 失败:`, error);
                    addMessage('ai', `❌ 处理文件 ${file.name} 失败: ${error.message}`);
                }
            }

            addMessage('ai', `✅ 已完成 ${filesToProcess.length} 个 Excel 文件的分析`);

        } else if (fileType === 'other') {
            filesToProcess = selectedFiles.filter(f => {
                const ext = f.name.split('.').pop().toLowerCase();
                return !f.type.startsWith('image/') && ext !== 'xlsx' && ext !== 'xls';
            });

            if (filesToProcess.length === 0) {
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可处理的文件</span>', true) : showToast('❌ 没有可处理的文件');
                return;
            }

            addMessage('user', `📝 开始分析 ${filesToProcess.length} 个文件...`);

            for (let i = 0; i < filesToProcess.length; i++) {
                const file = filesToProcess[i];
                try {
                    const base64Data = await fileToBase64(file);
                    await processFileWithAI(base64Data, file.name, 'summarize');
                } catch (error) {
                    ChatLog.error('FILEBATCH', `处理文件 ${file.name} 失败:`, error);
                    addMessage('ai', `❌ 处理文件 ${file.name} 失败: ${error.message}`);
                }
            }

            addMessage('ai', `✅ 已完成 ${filesToProcess.length} 个文件的分析`);
        }

        clearBatchSelection();
    }

    // --- 来自 app.js:L6496-L6499 ---
    /**
     * 清空批量选择
     */
    function clearBatchSelection() {
        clearAttachment();
        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已取消选择</span>', true) : showToast('已取消选择');
    }

    // --- 来自 app.js:L6504-L6573 ---
    /**
     * 批量识别图片（只识别，不填入）
     */
    async function batchRecognizeImages() {
        if (!selectedFiles || selectedFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有选择文件</span>', true) : showToast('❌ 没有选择文件');
            return;
        }

        const imageFiles = selectedFiles.filter(f => f.type.startsWith('image/'));
        if (imageFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可识别的图片文件</span>', true) : showToast('❌ 没有可识别的图片文件');
            return;
        }

        addMessage('user', `📦 开始批量识别 ${imageFiles.length} 张图片...`);

        // 显示批量进度消息（临时消息，不进入上下文历史）
        const progressMsgId = addMessage('ai', `⏳ 正在批量识别图片 (0/${imageFiles.length})<br><div style="font-size: 12px; color: #666; margin-top: 5px;">准备中</div>`, false, null, true);

        let successCount = 0;
        let failCount = 0;
        const results = [];

        for (let i = 0; i < imageFiles.length; i++) {
            const file = imageFiles[i];

            // 更新进度
            updateLastMessage('ai', `⏳ 正在批量识别图片 (${i + 1}/${imageFiles.length})<br>
                <div style="font-size: 12px; color: #666; margin-top: 5px;">
                    📄 处理: <strong>${escapeHtml(file.name)}</strong>
                </div>
                <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 6px; margin-top: 8px;">
                    <div style="width: ${((i + 1) / imageFiles.length * 100)}%; height: 100%; background: #4caf50; border-radius: 4px; transition: width 0.3s;"></div>
                </div>`);

            try {
                const base64Image = await fileToBase64(file);

                // 使用通用识别函数（只识别模式）
                await processImageGeneric(base64Image, file.name, detectImageType(file), {
                    fillMode: 'recognize_only'
                });

                successCount++;
                results.push({ name: file.name, status: 'success' });
            } catch (error) {
                ChatLog.error('FILEBATCH', `批量识别失败 [${file.name}]:`, error);
                failCount++;
                results.push({ name: file.name, status: 'failed', error: error.message });
            }

            // 短暂延迟，避免请求过快
            if (i < imageFiles.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }

        // 完成提示
        updateLastMessage('ai', `✅ 批量识别完成！<br>
            <div style="margin-top: 10px; padding: 10px; background: #e8f5e9; border-radius: 6px;">
                ✅ 成功: ${successCount} 个<br>
                ❌ 失败: ${failCount} 个<br>
                📊 总计: ${imageFiles.length} 个
            </div>`);

        // 清理附件状态
        clearAttachment();

        // 保存到历史
        const messages = getChatMessages();
        addToHistory(`批量识别: ${imageFiles.length} 张图片`, messages);
    }

    // --- 来自 app.js:L6578-L6594 ---
    /**
     * 批量识别并填入Excel
     */
    async function batchRecognizeAndFillExcel() {
        if (!selectedFiles || selectedFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有选择文件</span>', true) : showToast('❌ 没有选择文件');
            return;
        }

        const imageFiles = selectedFiles.filter(f => f.type.startsWith('image/'));
        if (imageFiles.length === 0) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 没有可识别的图片文件</span>', true) : showToast('❌ 没有可识别的图片文件');
            return;
        }

        addMessage('user', `📊 开始批量识别并填入 Excel (${imageFiles.length} 张图片中`);

        // 显示填入模式选择对话框
        showBatchFillModeDialog(imageFiles);
    }

    // --- 来自 app.js:L6599-L6699 ---
    /**
     * 显示批量填入模式对话框
     */
    function showBatchFillModeDialog(imageFiles) {
        currentBatchFiles = imageFiles;

        const dialogHtml = `
            <div id="batch-fill-mode-dialog" style="
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
            ">
                <div style="
                    background: white;
                    border-radius: 12px;
                    padding: 24px;
                    max-width: 480px;
                    width: 90%;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                    max-height: 80vh;
                    overflow-y: auto;
                ">
                    <h3 style="margin: 0 0 16px 0; font-size: 18px;">📊 批量填入模式</h3>

                    <div style="background: #f5f5f5; padding: 12px; border-radius: 6px; margin-bottom: 16px; font-size: 13px;">
                        <strong>待处理文件 (${imageFiles.length} 个):</strong><br>
                        <div style="max-height: 120px; overflow-y: auto; margin-top: 8px;">
                            ${imageFiles.map(f => `<div style="padding: 2px 0;">🖼️ ${escapeHtml(f.name)}</div>`).join('')}
                        </div>
                    </div>

                    <p style="color: #666; margin-bottom: 20px; font-size: 14px;">
                        请选择填入模式：
                    </p>

                    <div style="display: flex; flex-direction: column; gap: 12px; margin-bottom: 20px;">
                        <label style="
                            display: flex;
                            align-items: flex-start;
                            padding: 12px;
                            border: 2px solid #4caf50;
                            border-radius: 8px;
                            cursor: pointer;
                        ">
                            <input type="radio" name="batchFillMode" value="auto" checked style="margin-right: 12px; margin-top: 2px;">
                            <div>
                                <div style="font-weight: 500; margin-bottom: 4px;">🤖 智能插入（推荐）</div>
                                <div style="font-size: 13px; color: #666;">
                                    自动检测类型，横向表格格式，智能追加不重复表头
                                </div>
                            </div>
                        </label>

                        <label style="
                            display: flex;
                            align-items: flex-start;
                            padding: 12px;
                            border: 2px solid #e0e0e0;
                            border-radius: 8px;
                            cursor: pointer;
                        " onmouseover="this.style.borderColor='#4caf50'" onmouseout="this.style.borderColor='#e0e0e0'">
                            <input type="radio" name="batchFillMode" value="vertical" style="margin-right: 12px; margin-top: 2px;">
                            <div>
                                <div style="font-weight: 500; margin-bottom: 4px;">📋 纵向卡片</div>
                                <div style="font-size: 13px; color: #666;">
                                    纵向卡片格式（每行一个字段），适合查看详情
                                </div>
                            </div>
                        </label>
                    </div>

                    <div style="display: flex; gap: 10px; justify-content: flex-end;">
                        <button onclick="closeBatchFillModeDialog()" style="
                            padding: 8px 16px;
                            border: 1px solid #ddd;
                            background: white;
                            border-radius: 6px;
                            cursor: pointer;
                            font-size: 14px;
                        ">取消</button>
                        <button onclick="confirmBatchFillMode()" style="
                            padding: 8px 24px;
                            border: none;
                            background: #4caf50;
                            color: white;
                            border-radius: 6px;
                            cursor: pointer;
                            font-size: 14px;
                            font-weight: 500;
                        ">开始批量处理</button>
                    </div>
                </div>
            </div>
        `;

        document.body.insertAdjacentHTML('beforeend', dialogHtml);
    }

    // --- 来自 app.js:L6701-L6704 ---
    function closeBatchFillModeDialog() {
        const dialog = document.getElementById('batch-fill-mode-dialog');
        if (dialog) dialog.remove();
    }

    // --- 来自 app.js:L6706-L6716 ---
    async function confirmBatchFillMode() {
        const selected = document.querySelector('input[name="batchFillMode"]:checked');
        const fillMode = selected ? selected.value : 'auto';

        closeBatchFillModeDialog();

        if (currentBatchFiles && currentBatchFiles.length > 0) {
            await executeBatchProcessing(currentBatchFiles, fillMode);
            currentBatchFiles = null;
        }
    }

    // --- 来自 app.js:L6722-L6804 ---
    /**
     * 执行批量处理
     * 改进：先收集所有数据，再一次性插入Excel
     */
    async function executeBatchProcessing(imageFiles, fillMode, userMessage = '') {
        const messageContext = userMessage ? `<div style="margin-bottom: 8px; color: #4a90d9;">📝 ${userMessage}</div>` : '';
        addMessage('ai', `${messageContext}⏳ 开始批量处理 (0/${imageFiles.length})<br>
            <div style="font-size: 12px; color: #666; margin-top: 5px;">模式: ${fillMode === 'auto' ? '智能插入（横向）' : '纵向卡片'}</div>`);

        // 同步更新固定状态栏（强制执行）
        try {
            if (typeof App !== 'undefined' && App.fixedStatusBar) {
                if (!App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.init();
                }
                App.fixedStatusBar.setLoading(`批量处理: 0/${imageFiles.length}`, 0);
            }
        } catch(e) {
            ChatLog.warn('FILEBATCH', '[FixedStatusBar] 更新失败:', e);
        }

        let successCount = 0;
        let failCount = 0;
        const allData = [];  // 收集所有识别的数据

        // ========== 第一阶段：识别所有图片 ==========
        for (let i = 0; i < imageFiles.length; i++) {
            const file = imageFiles[i];
            const imageType = detectImageType(file);

            // 更新进度
            updateLastMessage('ai', `🔍 批量识别中 (${i + 1}/${imageFiles.length})<br>
                <div style="font-size: 12px; color: #666; margin-top: 5px;">
                    📄 <strong>${escapeHtml(file.name)}</strong> [${imageType}]
                </div>
                <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 6px; margin-top: 8px;">
                    <div style="width: ${((i + 1) / imageFiles.length * 50)}%; height: 100%; background: #2196f3; border-radius: 4px; transition: width 0.3s;"></div>
                </div>`);

            // 同步更新固定状态栏 - 进度（强制执行）
            try {
                if (typeof App !== 'undefined' && App.fixedStatusBar) {
                    const percent = Math.round(((i + 1) / imageFiles.length) * 50);
                    App.fixedStatusBar.updateProgress(percent, `正在识别: ${file.name}`);
                }
            } catch(e) {
                ChatLog.warn('FILEBATCH', '[FixedStatusBar] 更新失败:', e);
            }

            try {
                const base64Image = await fileToBase64(file);

                // 调用后端识别（只识别，不填入）
                const result = await recognizeImageOnly(base64Image, file.name, imageType);

                if (result && result.data && Object.keys(result.data).length > 0) {
                    allData.push({
                        fileName: file.name,
                        type: imageType,
                        data: result.data
                    });
                    successCount++;
                } else {
                    ChatLog.warn('FILEBATCH', `[${file.name}] 未识别到有效数据`);
                    failCount++;
                }
            } catch (error) {
                ChatLog.error('FILEBATCH', `批量处理失败 [${file.name}]:`, error);
                failCount++;
            }

            // 延迟避免过快
            if (i < imageFiles.length - 1) {
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }

        // ========== 第二阶段：一次性填入Excel ==========
        if (allData.length > 0 && fillMode !== 'recognize_only') {
            // 同步更新固定状态栏 - 填入Excel（强制执行）
            try {
                if (typeof App !== 'undefined' && App.fixedStatusBar) {
                    App.fixedStatusBar.setProcessing(`正在填入 ${allData.length} 条数据到 Excel...`);
                }
            } catch(e) {
                ChatLog.warn('FILEBATCH', '[FixedStatusBar] 更新失败:', e);
            }

            updateLastMessage('ai', `📊 正在将 ${allData.length} 条数据填入 Excel...<br>
                <div style="font-size: 12px; color: #666; margin-top: 5px;">请稍候...</div>
                <div style="width: 100%; background: #e0e0e0; border-radius: 4px; height: 6px; margin-top: 8px;">
                    <div style="width: 75%; height: 100%; background: #ff9800; border-radius: 4px; transition: width 0.3s;"></div>
                </div>`);

            try {
                await batchInsertToExcel(allData, fillMode);
            } catch (error) {
                ChatLog.error('FILEBATCH', '批量填入Excel失败:', error);
                updateLastMessage('ai', `⚠️ 填入Excel失败: ${error.message}<br>但已成功识别 ${allData.length} 个文件`);
            }
        }

        // 完成
        // 同步更新固定状态栏 - 完成状态（强制执行）
        try {
            if (typeof App !== 'undefined' && App.fixedStatusBar) {
                if (failCount > 0) {
                    App.fixedStatusBar.setSuccess(`批量完成: ${successCount} 成功, ${failCount} 失败`, 3000);
                } else {
                    App.fixedStatusBar.setSuccess(`全部 ${imageFiles.length} 个文件处理完成`, 3000);
                }
            }
        } catch(e) {
            ChatLog.warn('FILEBATCH', '[FixedStatusBar] 更新失败:', e);
        }

        updateLastMessage('ai', `✅ 批量处理完成！<br>
            <div style="margin-top: 10px; padding: 12px; background: #fff3e0; border-radius: 6px; border-left: 3px solid #ff9800;">
                <div style="font-weight: 600; color: #e65100; margin-bottom: 8px;">📊 处理统计</div>
                ✅ 成功: <strong>${successCount}</strong> 个<br>
                ❌ 失败: <strong>${failCount}</strong> 个<br>
                📈 总计: <strong>${imageFiles.length}</strong> 个<br>
                🎯 模式: ${fillMode === 'auto' ? '智能插入（横向表格）' : '纵向卡片'}<br>
                📥 已填入: <strong>${allData.length}</strong> 条记录
            </div>`);

        clearAttachment();

        const messages = getChatMessages();
        addToHistory(`批量处理: ${imageFiles.length} 张图片 (${fillMode})`, messages);
    }

    // --- 来自 app.js:L6809-L6873 ---
    /**
     * 只识别图片,不填入Excel(用于批量处理的第一阶段)
     */
    async function recognizeImageOnly(base64Image, fileName, type) {
        const config = IMAGE_PROCESSING_CONFIG[type] || IMAGE_PROCESSING_CONFIG['general'];

        try {
            // 检查后端是否可用
            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const inputs = {
                    imageData: base64Image,
                    fileName: fileName
                };

                const result = await window.chrome.webview.hostObjects.excelBridge.ExecuteSkillAsync(
                    config.skillName,
                    JSON.stringify(inputs)
                );

                const response = JSON.parse(result);

                if (response.success) {
                    const parsedData = parseRecognitionData(response.outputs || {}, config.fieldConfig);
                    let usedVisionFallback = false;

                    // 结果空洞 → 纯视觉兜底（ImageFieldExtractor）
                    const fb = window.App && window.App.recognition;
                    if (fb && typeof fb.isRecognitionResultEmpty === 'function' && fb.isRecognitionResultEmpty(response) &&
                        typeof fb.runImageFieldExtractor === 'function') {
                        const vision = await fb.runImageFieldExtractor(base64Image, fileName);
                        if (vision) {
                            parsedData.rawText = vision;
                            if (!parsedData.description) parsedData.description = '视觉识别结果';
                            usedVisionFallback = true;
                        }
                    }

                    return { data: parsedData, rawResponse: response, usedVisionFallback: usedVisionFallback };
                } else {
                    throw new Error(response.error || '识别失败');
                }
            } else {
                // 后端不可用时,使用AI作为备用方案
                ChatLog.warn('FILEBATCH', '⚠️ 后端不可用,使用AI识别作为备用方案');
                const aiPrompt = `请识别这张${type || '图片'}中的关键信息,返回JSON格式:
{
  "识别类型": "${type || '未知'}",
  "关键信息": "提取的主要文本内容",
  "详细信息": "其他重要信息"
}`;
                const aiResult = await callAIWithImage(aiPrompt, base64Image, fileName);
                return {
                    data: {
                        rawText: aiResult.content || aiResult || '',
                        description: `AI识别结果: ${type || '通用图片'}`
                    },
                    rawResponse: { success: true, source: 'ai_fallback' }
                };
            }
        } catch (error) {
            ChatLog.error('FILEBATCH', '识别失败:', error);
            // 如果Skill调用失败,尝试使用AI
            if (error.message !== '无法访问后端') {
                ChatLog.warn('FILEBATCH', '⚠️ Skill识别失败,尝试使用AI作为备用方案');
                try {
                    const aiPrompt = `请识别这张${type || '图片'}中的关键信息`;
                    const aiResult = await callAIWithImage(aiPrompt, base64Image, fileName);
                    return {
                        data: {
                            rawText: aiResult.content || aiResult || '',
                            description: `AI识别结果: ${type || '通用图片'}`
                        },
                        rawResponse: { success: true, source: 'ai_fallback' }
                    };
                } catch (aiError) {
                    ChatLog.error('FILEBATCH', 'AI识别也失败:', aiError);
                    throw new Error(`识别失败: ${error.message}`);
                }
            }
            throw error;
        }
    }

    // --- 来自 app.js:L6878-L6980 ---
    /**
     * 批量插入数据到Excel（使用智能表头匹配 - 与发票完全一致）
     */
    async function batchInsertToExcel(allData, fillMode) {
        if (!allData || allData.length === 0) return;

        // 元数据字段黑名单
        const META_FIELDS = [
            'imageType', 'image_type',
            'description',
            'structuredData',
            'suggestedAction', 'suggested_action',
            'extractedText', 'extracted_text',
            'type', 'success', 'error',
            'executionLog', 'totalExecutionTime'
        ];

        const isMetaField = (key) => META_FIELDS.includes(key);

        // 构建记录列表（字典格式）
        const records = [];

        allData.forEach(item => {
            const record = {};

            Object.entries(item.data).forEach(([key, value]) => {
                if (value !== undefined && value !== null && value !== '' && !isMetaField(key)) {
                    record[key] = String(value);
                }
            });

            if (Object.keys(record).length > 0) {
                records.push(record);
            }
        });

        if (records.length === 0) {
            ChatLog.warn('FILEBATCH', '没有有效数据可插入');
            return;
        }

        try {
            // ✨ 调用新的智能表头匹配方法
            const payload = {
                records: records,
                mode: fillMode === 'vertical' ? 'new' : 'smart'  // smart=智能匹配, new=新建表头
            };

            ChatLog.info('FILEBATCH', '🧠 智能批量插入数据:', JSON.stringify(payload, null, 2).substring(0, 200) + '...');

            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertDataSmart(JSON.stringify(payload));
                const response = JSON.parse(result);

                if (response.success) {
                    ChatLog.info('FILEBATCH', '✅ 智能插入成功:', response);

                    // 显示详细结果
                    let modeText = response.modeUsed === 'SmartMatch' ? '智能表头匹配' : '创建新表头';
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ ' + response.message + ' (' + modeText + ', 匹配' + response.headersMatched + '个字段)</span>', true) : showToast(`✅ ${response.message} (${modeText}, 匹配${response.headersMatched}个字段)`);
                } else {
                    throw new Error(response.error || '智能插入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('FILEBATCH', '智能批量插入失败:', error);
            // 回退到旧方法
            ChatLog.info('FILEBATCH', '⚠️ 回退到简单插入模式...');

            // 构建二维数组作为回退方案
            let rowsToInsert = [];

            if (fillMode === 'vertical') {
                allData.forEach((item, index) => {
                    if (index > 0) rowsToInsert.push([]);
                    rowsToInsert.push(['字段', '值']);
                    Object.entries(item.data).forEach(([key, value]) => {
                        if (!isMetaField(key) && value) {
                            rowsToInsert.push([key, String(value)]);
                        }
                    });
                });
            } else {
                const allFields = new Map();
                allData.forEach(item => {
                    Object.keys(item.data).forEach(key => {
                        if (!isMetaField(key) && !allFields.has(key)) {
                            allFields.set(key, true);
                        }
                    });
                });

                const fieldNames = Array.from(allFields.keys());
                rowsToInsert.push(fieldNames);

                allData.forEach(item => {
                    const row = fieldNames.map(field => String(item.data[field] || ''));
                    rowsToInsert.push(row);
                });
            }

            await insertTableToExcel(rowsToInsert);
        }
    }

    // 全局变量存储批量处理的文件
    let currentBatchFiles = null;

    // --- 来自 app.js:L11965-L12044 ---
    async function extractFileContent(file, base64Data, forcedType = null) {
        const fileExtension = file.name.split('.').pop().toLowerCase();

        try {
            // 图片文件：尝试OCR识别或返回描述
            if (file.type.startsWith('image/')) {
                // 如果有强制类型，尝试调用后端识别
                if (forcedType) {
                    try {
                        const operationType = mapForcedType(forcedType);
                        const result = await recognizeImageOnly(base64Data, file.name, operationType);
                        if (result && result.data) {
                            // 简化识别结果，只保留关键信息
                            let simplifiedData = {};
                            const data = result.data;

                            // 提取关键字段
                            if (data.rawText) {
                                simplifiedData.rawText = data.rawText.substring(0, 800) + (data.rawText.length > 800 ? '...' : '');
                            } else if (data.description) {
                                simplifiedData.description = data.description.substring(0, 800);
                            }

                            // 如果有结构化字段也保留
                            const keyFields = ['姓名', '性别', '民族', '出生日期', '住址', '公民身份号码', '身份证号码', '签发机关', '有效期限'];
                            keyFields.forEach(field => {
                                if (data[field]) simplifiedData[field] = data[field];
                            });

                            // 检测使用的OCR引擎
                            const ocrEngine = data.ocrEngine || result.rawResponse?.ocrEngine || 'PaddleOCR';
                            const engineLabel = result.usedVisionFallback ? 'ImageFieldExtractor 纯视觉兜底' : `${ocrEngine}识别`;
                            ChatLog.info('FILEBATCH', `[OCR结果] ${ocrEngine}返回结果:`, simplifiedData);

                            // 格式化 rawText,将换行符转换为 <br>
                            let formattedRawText = simplifiedData.rawText || simplifiedData.description || '';
                            formattedRawText = formattedRawText.replace(/\\r\n/g, '<br>').replace(/\n/g, '<br>').replace(/\r/g, '<br>');

                            // 构建格式化的识别结果
                            let resultHtml = `<div style="margin-bottom: 8px;"><strong>[${engineLabel}]</strong> ${file.name}</div>`;
                            resultHtml += `<div style="margin-bottom: 4px;"><span style="color: #666;">类型:</span> ${forcedType}</div>`;
                            resultHtml += `<div style="margin-top: 8px; padding: 10px; background: #f5f5f5; border-radius: 6px; font-size: 13px; line-height: 1.6;">${formattedRawText}</div>`;

                            return resultHtml;
                        }
                    } catch (recognizeError) {
                        // 识别失败时，返回简化描述
                        return `[图片文件] ${file.name}\n类型: ${forcedType}\n注：OCR识别失败，建议使用多模态模型直接分析`;
                    }
                }
                return `[图片文件] ${file.name}\n大小: ${(file.size / 1024).toFixed(1)} KB\n格式: ${fileExtension.toUpperCase()}\n注：图片内容已通过 OCR 识别`;
            }

            // Excel 文件
            if (fileExtension === 'xlsx' || fileExtension === 'xls') {
                return await extractExcelContent(file, base64Data);
            }

            // 文本文件
            if (fileExtension === 'txt' || fileExtension === 'csv') {
                return await extractTextContent(base64Data);
            }

            // Word 文档
            if (fileExtension === 'docx' || fileExtension === 'doc') {
                if (typeof App !== 'undefined' && App.DocumentProcessor) {
                    try {
                        var docR = await App.DocumentProcessor.processDocument(file);
                        return `[Word文档] ${file.name}
字符数: ${docR.totalChars}
页数: ${docR.totalPages}

${(docR.fullText || '').substring(0, 50000)}`;
                    } catch (e) {
                        return `[Word文档] ${file.name}\n大小: ${(file.size / 1024).toFixed(1)} KB\n注：解析失败 - ${e.message}`;
                    }
                }
                return `[Word文档] ${file.name}\n大小: ${(file.size / 1024).toFixed(1)} KB\n注：Word文档需要专门解析，当前仅支持文本提取`;
            }

            // PDF 文件
            if (fileExtension === 'pdf') {
                if (typeof App !== 'undefined' && App.DocumentProcessor) {
                    try {
                        var pdfR = await App.DocumentProcessor.processDocument(file);
                        return `[PDF文档] ${file.name}
字符数: ${pdfR.totalChars}
页数: ${pdfR.totalPages}

${(pdfR.fullText || '').substring(0, 50000)}`;
                    } catch (e) {
                        return `[PDF文件] ${file.name}\n解析失败: ${e.message}`;
                    }
                }
                return `[PDF文件] ${file.name}\n大小: ${(file.size / 1024).toFixed(1)} KB\n注：PDF内容需要专门解析`;
            }

            // 其他文件类型
            return `[文件] ${file.name}\n类型: ${fileExtension.toUpperCase()}\n大小: ${(file.size / 1024).toFixed(1)} KB\n注：该文件类型暂不支持内容提取`;

        } catch (error) {
            ChatLog.error('FILEBATCH', `提取文件内容失败 [${file.name}]:`, error);
            return `[文件] ${file.name}\n提取失败: ${error.message}`;
        }
    }

    // --- 来自 app.js:L13917-L13939 ---
    async function extractExcelContent(file, base64Data) {
        var fileName = file.name || '';
        ChatLog.info('FILEBATCH', '[Excel] 提取 Excel 内容:', fileName);

        try {
            // 剥离 DataURL 前缀（data:...;base64,），只保留纯 Base64 数据
            var pureBase64 = base64Data.indexOf(',') > -1 ? base64Data.substring(base64Data.indexOf(',') + 1) : base64Data;

            // 使用后端解析 Excel（统一走 ParseDocumentAsync 入口）
            var bridgeResult = null;
            if (window.excelBridge && typeof window.excelBridge.ParseDocumentAsync === 'function') {
                try {
                    const result = await window.excelBridge.ParseDocumentAsync(fileName || 'uploaded.xlsx', pureBase64);
                    const resultObj = JSON.parse(result);
                    if (resultObj.success) {
                        bridgeResult = resultObj;
                    } else {
                        ChatLog.warn('FILEBATCH', '[Excel] ParseDocumentAsync 返回失败:', resultObj.error);
                    }
                } catch (bridgeErr) {
                    ChatLog.warn('FILEBATCH', '[Excel] ParseDocumentAsync 调用异常，降级到 SheetJS:', bridgeErr.message);
                }
            }
            
            if (bridgeResult) {
                // ★ 诊断日志
                ChatLog.info('FILEBATCH', '[Excel] resultObj keys:', Object.keys(bridgeResult));
                ChatLog.info('FILEBATCH', '[Excel] totalPages:', bridgeResult.totalPages, '| totalChars:', bridgeResult.totalChars);
                ChatLog.info('FILEBATCH', '[Excel] pages:', bridgeResult.pages ? bridgeResult.pages.length : 'NULL/UNDEFINED');

                // 【P0修复】使用结构化摘要替代旧的 structure.Headings 概览 + fullText 30000 截断
                var finalContent = buildExcelStructuredSummary(bridgeResult, fileName);
                ChatLog.info('FILEBATCH', '[Excel] 结构化摘要长度:', finalContent.length);
                ChatLog.info('FILEBATCH', '[Excel] 内容预览(前500字):', finalContent.substring(0, 500));
                return finalContent;
            }
            
            // 【Web版 fallback】SheetJS 直接读 File.arrayBuffer()，无需 Base64 中转
            ChatLog.info('FILEBATCH', '[Excel] 使用 SheetJS 客户端解析');
            if (typeof XLSX !== 'undefined') {
                try {
                    var buf = await file.arrayBuffer();
                    var bytes = new Uint8Array(buf);
                    var workbook = XLSX.read(bytes, { type: 'array' });

                    // 构造与 NPOI/Interop 路径兼容的 resultObj 结构
                    var pages = [];
                    var fullTextParts = [];
                    workbook.SheetNames.forEach(function(sheetName) {
                        var worksheet = workbook.Sheets[sheetName];
                        var jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
                        var sheetLines = [];
                        sheetLines.push('【工作表: ' + sheetName + '】');
                        jsonData.forEach(function(row, rowIdx) {
                            var rowText = row.map(function(cell) { return String(cell != null ? cell : ''); }).join('\t');
                            if (rowIdx === 0) {
                                sheetLines.push('【表头】' + rowText);
                            } else {
                                sheetLines.push(rowText);
                            }
                        });
                        var sheetText = sheetLines.join('\n');
                        pages.push({ title: sheetName, text: sheetText });
                        fullTextParts.push(sheetText);
                    });

                    var resultObj = {
                        success: true,
                        fileType: 'excel',
                        fullText: fullTextParts.join('\n\n'),
                        totalChars: fullTextParts.join('\n\n').length,
                        totalPages: pages.length,
                        isScanned: false,
                        pages: pages,
                        structure: { Headings: [], Tables: [] },
                        metadata: {
                            fileFormat: (fileName || '').split('.').pop() || 'xlsx',
                            properties: { '工作表列表': workbook.SheetNames.join(', ') }
                        }
                    };

                    ChatLog.info('FILEBATCH', '[Excel] SheetJS 解析成功:', workbook.SheetNames.length, '个工作表');
                    var finalContent = buildExcelStructuredSummary(resultObj, fileName);
                    ChatLog.info('FILEBATCH', '[Excel] Web版结构化摘要长度:', finalContent.length);
                    return finalContent;
                } catch (sheetjsErr) {
                    ChatLog.error('FILEBATCH', '[Excel] SheetJS 解析失败:', sheetjsErr);
                }
            }
            // 最终降级
            return 'Excel 文件: ' + fileName + '\n注意：需要后端支持才能解析 Excel 内容';
        } catch (error) {
            ChatLog.error('FILEBATCH', '[Excel] 提取 Excel 内容失败:', error);
            return 'Excel 文件: ' + fileName + '\n解析错误: ' + error.message;
        }
    }

    // --- 来自 app.js:L13942-L13953 ---
    // 提取文本文件内容
    async function extractTextContent(base64Data) {
        try {
            // 移除 base64 前缀
            const base64Content = base64Data.includes(',') ? base64Data.split(',')[1] : base64Data;
            // 解码 base64
            const byteString = atob(base64Content);
            return byteString;
        } catch (error) {
            ChatLog.error('FILEBATCH', '提取文本内容失败:', error);
            return '文本解码失败: ' + error.message;
        }
    }

    // ==================== 辅助函数：根据Skill名称执行对应识别 ====================

    // --- 来自 app.js:L7031-L7119 ---
    async function executeSkillByName(skillName, base64Data, fileName, messageElements = null) {
        const _emit = (content) => {
            if (messageElements) {
                App.updateStreamingMessage(messageElements, content);
            } else {
                addMessage('ai', content);
            }
        };

        // 从 cachedSkills 中获取 skill 详情
        const skills = App.cachedSkills || await App.loadSkillsForLLM();
        const skill = skills.find(s => s.name === skillName || s.Name === skillName);

        if (!skill) {
            ChatLog.error('FILEBATCH', '找不到 Skill:', skillName);
            _emit(`❌ 找不到 Skill: ${skillName}`);
            return;
        }

        // 获取 processingMode（处理可能的大小写差异）
        let processingMode = skill.processingMode || skill.ProcessingMode || '';
        const supportsImageInput = skill.supportsImageInput || skill.SupportsImageInput || false;

        // 将数字类型的 processingMode 转换为字符串（后端枚举值：0=VBACode, 1=MultimodalLLM, 2=Hybrid, 3=LocalOCR, 4=PythonCode）
        const processingModeMap = {
            0: 'VBACode',
            1: 'MultimodalLLM',
            2: 'Hybrid',
            3: 'LocalOCR',
            4: 'PythonCode'
        };
        if (typeof processingMode === 'number') {
            processingMode = processingModeMap[processingMode] || 'VBACode';
        }

        ChatLog.info('FILEBATCH', '执行 Skill:', skillName, '处理模式:', processingMode, '支持图片:', supportsImageInput);

        // 根据 processingMode 选择执行方式
        if (processingMode === 'MultimodalLLM' && supportsImageInput) {
            // 需要图片输入的 skill
            if (!base64Data) {
                // 检查是否有已选择的附件
                if (selectedFile) {
                    ChatLog.info('FILEBATCH', '📎 使用已选择的附件:', selectedFile.name);
                    base64Data = await fileToBase64(selectedFile);
                    fileName = selectedFile.name;
                    // 保存附件的检测结果（在清除附件之前）
                    const savedDetectedType = selectedFile.detectedType;
                    ChatLog.info('FILEBATCH', '💾 保存附件检测结果:', savedDetectedType || '无');
                    // 清除已选择的附件，避免重复处理
                    clearAttachment();
                    // 将保存的检测结果附加到文件名上，供后续使用
                    if (savedDetectedType) {
                        window._lastFileDetectedType = savedDetectedType;
                    }
                } else {
                    // 没有图片数据，提示用户上传图片
                    _emit(`💡 请上传图片以执行 "${skill.displayName || skillName}" 功能`);
                    // 触发文件选择
                    setTimeout(() => {
                        document.getElementById('file-input').click();
                    }, 100);
                    return;
                }
            }

            // 将技能名称映射为对应的 backendType
            let operationType = skillName;

            // 技能名称到 backendType 的映射
            const skillToBackendType = {
                'ExtractText': '文档',
                'ExtractDriverLicense': '驾驶证',
                'ExtractBusinessLicense': '营业执照',
                'ExtractBankCard': '银行卡'
            };

            if (skillToBackendType[skillName]) {
                operationType = skillToBackendType[skillName];
                ChatLog.info('FILEBATCH', '🎯 技能名称映射为 backendType:', skillName, '->', operationType);
            } else if (window._lastFileDetectedType &&
                (window._lastFileDetectedType === '身份证' || window._lastFileDetectedType === '名片')) {
                // 使用之前保存的附件检测结果
                operationType = window._lastFileDetectedType;
                ChatLog.info('FILEBATCH', '🎯 使用保存的附件检测结果:', operationType);
                // 清除已使用的检测结果
                window._lastFileDetectedType = null;
            }

            // 有图片数据，执行图片识别
            await processImageRecognition(base64Data, fileName, operationType);
        } else if (processingMode === 'VBACode') {
            // VBA 代码生成 skill
            await executeVBASkill(skill);
        } else if (processingMode === 'PythonCode') {
            // 【P0修复】Python 脚本 skill：复用 VBA 执行流程（填示例→发送→后端 ExecuteSkillAsync 分发）
            // 后端 SkillManager 已修复，会将 PythonCode 正确路由到 PythonExecutionService
            ChatLog.info('FILEBATCH', '[PythonCode Skill] 复用 executeVBASkill 流程，后端将路由到 PythonExecutionService');
            if (typeof App.showToast === 'function') {
                App.showToast('🐍 执行 Python 脚本 Skill...', 1500);
            }
            await executeVBASkill(skill);
        } else {
            // 其他类型的 skill，使用通用处理
            await executeGenericSkill(skill);
        }
    }

    // --- 来自 app.js:L7122-L7166 ---
    // 执行 VBA 类型的 Skill
    async function executeVBASkill(skill) {
        ChatLog.info('FILEBATCH', '执行 VBA Skill:', skill.name);

        // 设置输入框提示
        const input = document.getElementById('message-input');
        const displayName = skill.displayName || skill.Name || skill.name;

        // 根据 skill 名称执行不同操作
        const skillActions = {
            'GenerateVBA': () => {
                input.placeholder = `描述您需要的 VBA 功能，例如：${skill.examples ? skill.examples[0] : '写一个排序宏'}`;
                input.focus();
            },
            'SplitWorkbook': () => {
                input.value = '请帮我拆分当前工作簿，将每个工作表保存为单独的文件';
                sendMessage();
            },
            'MergeSheets': () => {
                input.value = '请帮我合并所有工作表到一个工作表';
                sendMessage();
            },
            'FormatCells': () => {
                input.value = '请帮我格式化选中的单元格为文本格式';
                sendMessage();
            },
            'RemoveDuplicates': () => {
                input.value = '请帮我删除 A 列的重复值';
                sendMessage();
            },
            'ClearChat': () => {
                clearChat();
            },
            'NewChat': () => {
                newChat();
            }
        };

        if (skillActions[skill.name]) {
            skillActions[skill.name]();
        } else {
            // 默认行为：显示 skill 描述作为提示
            input.placeholder = skill.description || `执行 ${displayName}`;
            input.focus();
        }
    }

    // --- 来自 app.js:L7169-L7181 ---
    // 执行通用 Skill
    // 恢复原名，确保AI核心模块能正确调用
    async function executeGenericSkill(skill) {
        ChatLog.info('FILEBATCH', '[executeGenericSkill] UI准备:', skill.name);
        const input = document.getElementById('message-input');
        const displayName = skill.displayName || skill.Name || skill.name;

        // 【修复】如果当前处于发送流程中（isSending=true），不回写输入框
        // 避免 sendMessage 已清空输入框后，又把 skill.examples 旧内容写回
        if (App.state && App.state.isSending) {
            ChatLog.info('FILEBATCH', '[executeGenericSkill] 跳过输入框回写（当前处于发送流程中）');
            return;
        }

        // 设置输入框为 skill 的示例或描述（仅在用户主动点击技能卡片、非发送路径时生效）
        if (skill.examples && skill.examples.length > 0) {
            input.value = skill.examples[0];
        } else {
            input.placeholder = skill.description || `执行 ${displayName}`;
        }
        input.focus();
    }

    // ==================== 多文件AI分析（sendMessage 多文件路径） ====================

    /**
     * 多文件AI分析：逐个提取文件内容，合并后发送给AI
     * 由 app-ai-send.js:sendMessage 在多文件场景下调用
     * @param {string} message - 用户消息
     * @param {File[]} files - 文件数组
     * @param {Object} ctx - 当前发送上下文（SendContext），用于保存 thinking 气泡
     * @param {string} loadingMessageId - 思考气泡ID，用于复用（旧接口兼容）
     */
    async function processMultipleFiles(message, files, ctx = null, loadingMessageId = null) {
        ChatLog.info('FILEBATCH', '[processMultipleFiles] 开始，文件数:', files.length);

        // ========== 初始化：创建思考气泡 ==========
        // 根据文件类型显示更明确的状态：图片走 OCR，文档走本地提取
        const initialImageCount = files.filter(f => UnifiedFileProcessor.validateFile(f).type === 'image').length;
        const initialDocCount = files.length - initialImageCount;
        let initialStatus = '📁 本地处理文件中';
        if (initialImageCount > 0 && initialDocCount === 0) {
            initialStatus = initialImageCount === 1 ? '🖼️ 本地 OCR 识别图片中' : `🖼️ 本地 OCR 识别 ${initialImageCount} 张图片中`;
        } else if (initialDocCount > 0 && initialImageCount === 0) {
            initialStatus = initialDocCount === 1 ? '📄 本地提取文档中' : `📄 本地提取 ${initialDocCount} 个文档中`;
        }

        // 优先使用外部传入的 thinking 气泡（第四阶段状态管理统一）
        let thinkingElements = ctx && typeof ctx.getThinkingElements === 'function'
            ? ctx.getThinkingElements()
            : null;
        if (thinkingElements && thinkingElements.wrapper) {
            loadingMessageId = thinkingElements.wrapper.id;
        }

        // 确保思考气泡被创建并正确保存到当前上下文
        if (!loadingMessageId) {
            loadingMessageId = App.showAIThinkingStatus(initialStatus);
            if (ctx && typeof ctx.setLoadingMessageId === 'function') {
                ctx.setLoadingMessageId(loadingMessageId);
            }
            if (loadingMessageId) {
                const thinkingWrapper = document.getElementById(loadingMessageId);
                if (thinkingWrapper) {
                    thinkingElements = {
                        wrapper: thinkingWrapper,
                        messageDiv: thinkingWrapper.querySelector('.message.ai')
                    };
                    if (ctx && typeof ctx.setThinkingElements === 'function') {
                        ctx.setThinkingElements(thinkingElements);
                    }
                }
            }
        } else {
            App.updateAIThinkingStatus(loadingMessageId, initialStatus);
            // 外部传入 loadingMessageId 时，同步补齐 ctx.thinkingElements
            if (!thinkingElements && loadingMessageId) {
                const thinkingWrapper = document.getElementById(loadingMessageId);
                if (thinkingWrapper) {
                    thinkingElements = {
                        wrapper: thinkingWrapper,
                        messageDiv: thinkingWrapper.querySelector('.message.ai')
                    };
                    if (ctx && typeof ctx.setThinkingElements === 'function') {
                        ctx.setThinkingElements(thinkingElements);
                    }
                }
            }
        }

        // ========== 第一阶段：逐个提取文件内容 ==========
        const fileContents = [];
        const failedFiles = [];

        for (let i = 0; i < files.length; i++) {
            const file = files[i];

            try {
                const validation = UnifiedFileProcessor.validateFile(file);
                if (!validation.valid) throw new Error(validation.error);

                const base64Data = await fileToBase64(file);
                let content = '';

                if (validation.type === 'image') {
                    if (loadingMessageId) {
                        App.updateAIThinkingStatus(loadingMessageId, `📝 OCR 识别 [${i + 1}/${files.length}] ${file.name} 中`);
                    }
                    try {
                        const ocrProcessor = App.FileProcessor.createOcrProcessor({
                            fileContext: {
                                base64Data: base64Data,
                                fileName: file.name
                            },
                            loadingMessageId: null
                        });
                        const ocrResult = await ocrProcessor.process();

                        if (ocrResult.success && ocrResult.result) {
                            content = cleanOcrResult(ocrResult.result);
                            if (loadingMessageId) {
                                App.updateAIThinkingStatus(loadingMessageId, `✅ [${i + 1}/${files.length}] ${file.name} OCR 识别完成`);
                            }
                        } else {
                            content = `[图片 OCR 失败: ${ocrResult.error || '未知错误'}]`;
                            if (loadingMessageId) {
                                App.updateAIThinkingStatus(loadingMessageId, `❌ [${i + 1}/${files.length}] ${file.name} OCR 识别失败`);
                            }
                        }
                    } catch (ocrErr) {
                        content = `[图片 OCR 异常: ${ocrErr.message}]`;
                        if (loadingMessageId) {
                            App.updateAIThinkingStatus(loadingMessageId, `❌ [${i + 1}/${files.length}] ${file.name} OCR 识别异常`);
                        }
                    }
                } else {
                    if (loadingMessageId) {
                        App.updateAIThinkingStatus(loadingMessageId, `📖 [${i + 1}/${files.length}] 本地提取 ${file.name} 中`);
                    }
                    const extracted = await UnifiedFileProcessor.extractContent(file, base64Data);
                    content = extracted.content || '[无法提取内容]';
                    if (loadingMessageId) {
                        App.updateAIThinkingStatus(loadingMessageId, `✅ [${i + 1}/${files.length}] ${file.name} 提取完成`);
                    }
                }

                fileContents.push({
                    fileName: file.name,
                    fileType: validation.type,
                    content: content,
                    metadata: {}
                });
                
            } catch (error) {
                failedFiles.push({ name: file.name, error: error.message });
                fileContents.push({
                    fileName: file.name,
                    fileType: 'unknown',
                    content: `[文件提取失败: ${error.message}]`,
                    metadata: {}
                });
            }
        }

        // 第一阶段结束：文件提取完成，更新状态避免停留在"OCR识别中"
        if (loadingMessageId) {
            const _pendingModel2 = (typeof App.getEffectiveModel === 'function'
                ? App.getEffectiveModel('', false)
                : (App.currentChatModel || App.config.model || ''));
            const _isReasoningModel2 = typeof App.shouldSkipForcedThinkingTags === 'function'
                ? App.shouldSkipForcedThinkingTags(_pendingModel2)
                : /deepseek[-_]?r1|o1|o3|qwq|qwen2\.5|qwen2-5|qwen2_5|qwen-2\.5/i.test(_pendingModel2 || '');
            const _waitingText2 = _isReasoningModel2
                ? '🧠 模型正在深度思考中'
                : '🤖 AI 正在分析识别结果...';
            App.updateAIThinkingStatus(loadingMessageId, _waitingText2);
        }

        // ========== 第二阶段：构建多文件分析 Prompt ==========
        const imageCount = fileContents.filter(f => f.fileType === 'image').length;
        const excelCount = fileContents.filter(f => f.fileType === 'excel').length;
        const pdfCount = fileContents.filter(f => f.fileType === 'pdf').length;
        const docCount = fileContents.filter(f => f.fileType === 'document').length;
        const textCount = fileContents.filter(f => f.fileType === 'text').length;

        const typeDescParts = [];
        if (imageCount > 0) typeDescParts.push(`${imageCount} 张图片`);
        if (excelCount > 0) typeDescParts.push(`${excelCount} 个Excel`);
        if (pdfCount > 0) typeDescParts.push(`${pdfCount} 个PDF`);
        if (docCount > 0) typeDescParts.push(`${docCount} 个Word`);
        if (textCount > 0) typeDescParts.push(`${textCount} 个文本`);
        const typeDesc = typeDescParts.join('、') || `${files.length} 个文件`;

        // 构建文件内容块
        let fileSummaryLines = [];
        fileContents.forEach((fc, idx) => {
            const preview = (fc.content || '').substring(0, 3000);
            fileSummaryLines.push(`===== ${fc.fileName} =====\n${preview}${(fc.content || '').length > 3000 ? '\n...(内容已截断)' : ''}`);
        });
        const fileSummaryBlock = fileSummaryLines.join('\n\n');

        // 失败文件提示
        let failedHint = '';
        if (failedFiles.length > 0) {
            failedHint = `\n\n【注意】${failedFiles.length} 个文件处理失败`;
        }

        // 用户消息
        const userPrompt = (message && message.trim().length > 0)
            ? message
            : `请分析以上文件内容，给出综合总结。`;

        // 判断当前模型是否应跳过强制 <thinking> 标签要求：
        // 1) 原生就会输出 <think> 的推理模型（deepseek-r1 / o1 / o3 / qwq 等）
        // 2) 已知不会/不擅长遵守 <thinking> 标签指令的模型（qwen2.5 系列等）
        // 对这类模型只给出口头顺序提示，不再强制标签格式，避免污染输出。
        const currentModel = (typeof App.getEffectiveModel === 'function')
            ? App.getEffectiveModel('', false)
            : (App.currentChatModel || App.config.model || '');
        const shouldSkipForcedThinking = typeof App.shouldSkipForcedThinkingTags === 'function'
            ? App.shouldSkipForcedThinkingTags(currentModel)
            : /deepseek[-_]?r1|o1|o3|qwq|qwen2\.5|qwen2-5|qwen2_5|qwen-2\.5|claude[-_]?3[-_]?5[-_]?sonnet[-_]?thinking/i.test(currentModel || '');
        const outputOrderHint = shouldSkipForcedThinking
            ? '请先输出分析推理过程，再输出最终分析结果。'
            : '首先用 <thinking> 和 </thinking> 标签输出你的分析推理过程（这部分强制要求，不可跳过），然后在 </thinking> 标签之后输出最终分析结果。';

        // 构建简洁高效的系统提示
        const systemPrompt = `你是文件分析助手。用户上传了 ${files.length} 个文件（${typeDesc}），以下是每个文件的提取内容。

请根据用户问题，基于文件内容直接给出分析结果。${outputOrderHint}

【输出格式要求】
1. 每个文件用二级标题 (##) 开头
2. 关键信息用**粗体**标注
3. 结构化数据用表格呈现（见下方【结构化数据要求】）
4. 不同文件之间用分隔线 (---) 分隔
5. 最后给出综合总结（用段落文字）

【结构化数据要求 - 重要】
- "结构化数据"必须提取文件内容中的业务关键字段，而不是文件元数据
- 例如车票/票据类图片，表格字段应包括：票号、出发地、目的地、日期、时间、车次/航班、座位类型、票价、乘客信息等实际业务信息
- 例如营业执照/证件类图片，表格字段应包括：企业名称、统一社会信用代码、法定代表人、注册资本、成立日期、经营范围、地址等实际业务信息
- 例如 Excel/Word/文本文件，表格字段应为文档中的关键列或关键属性
- 严禁在结构化数据表格中放置以下技术性元数据：文件类型、处理工具、OCR引擎、识别字符数、原始图片路径、临时文件路径

【禁止】
- 不要使用 bullet points（•、-、* 开头的列表项）
- 不要输出任务计划或步骤列表
- 不要在结尾突然添加新的列表项
- 不要把 JSON 字段名（如 result.text、result.engine、charCount、imageSource）当作结构化数据表格的字段

${failedHint}`;

        // 构建 customMessages
        const customMessages = [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: `${fileSummaryBlock}

---

用户问题：${userPrompt}` }
        ];

        // ========== 第三阶段：调用 AI ==========
        ChatLog.info('FILEBATCH', '[processMultipleFiles] 调用 callAIStreamAndGetFullResponse, loadingMessageId:', loadingMessageId);
        try {
            await App.callAIStreamAndGetFullResponse(
                userPrompt,
                null,    // imageBase64
                loadingMessageId,    // 复用思考气泡
                false,   // returnRawText
                true,    // skipLengthLimit
                customMessages,
                null,    // sendTime
                null,    // fileContext
                null,    // resolvedModel
                true,    // skipExecutionFlow：多文件分析不渲染四阶段卡片
                thinkingElements  // 复用外部传入的 thinking 气泡
            );
        } catch (error) {
            App.addMessage('ai', `❌ 多文件分析失败：${error.message}`);
        } finally {
            // 【清理】重置当前上下文的 thinking 状态，避免影响后续消息
            if (ctx && typeof ctx.setLoadingMessageId === 'function' && ctx.getLoadingMessageId() === loadingMessageId) {
                ctx.setLoadingMessageId(null);
            }
            if (App.state && App.state.loadingMessageId === loadingMessageId) {
                App.state.loadingMessageId = null;
            }
        }

        // 清理状态栏
        if (App.fixedStatusBar) {
            try {
                App.fixedStatusBar.setSuccess('分析完成', 2000);
            } catch(e) { (window.ChatLog || console).warn?.('UI', '[file-batch] 状态栏提示失败:', e); }
        }
    }

    // ==================== 模块导出 ====================
    App.getFileIcon = getFileIcon;
    App.getFileType = getFileType;
    App.getFileTypeLabel = getFileTypeLabel;
    App.formatFileSize = formatFileSize;
    App.escapeHtml = escapeHtml;
    App.detectImageType = detectImageType;
    App.fileToBase64 = fileToBase64;
    App.mapForcedType = mapForcedType;
    App.handleFileSelect = handleFileSelect;
    App.clearAttachment = clearAttachment;
    App.removeAttachment = removeAttachment;
    App.renderAttachmentList = renderAttachmentList;
    App.updateAttachmentState = updateAttachmentState;
    App.getFileActionButtons = getFileActionButtons;
    App.renderAttachmentActions = renderAttachmentActions;
    App.quickAnalyzeImage = quickAnalyzeImage;
    App.cancelCurrentOperation = cancelCurrentOperation;
    App.restoreAttachmentButtons = restoreAttachmentButtons;
    App.showBatchProcessingUI = showBatchProcessingUI;
    App.startBatchProcessing = startBatchProcessing;
    App.startBatchFileProcessing = startBatchFileProcessing;
    App.clearBatchSelection = clearBatchSelection;
    App.batchRecognizeImages = batchRecognizeImages;
    App.batchRecognizeAndFillExcel = batchRecognizeAndFillExcel;
    App.executeBatchProcessing = executeBatchProcessing;
    App.recognizeImageOnly = recognizeImageOnly;
    App.batchInsertToExcel = batchInsertToExcel;
    App.extractFileContent = extractFileContent;
    App.extractExcelContent = extractExcelContent;
    App.extractTextContent = extractTextContent;
    App.executeSkillByName = executeSkillByName;
    App.executeVBASkill = executeVBASkill;
    App.executeGenericSkill = executeGenericSkill;
    App.processMultipleFiles = processMultipleFiles;

})();