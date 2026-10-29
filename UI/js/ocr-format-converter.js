/**
 * ============================================================
 * ocr-format-converter.js — OCR Format Converter
 * ============================================================
 * 功能：OCR 格式转换器，将 OCR 结果在不同格式（JSON/CSV/Excel/Markdown）之间转换。
 *
 * 核心类/函数：
 *   - convertToJson() — 转为 JSON
 *   - convertToCsv() — 转为 CSV
 *   - convertToExcel() — 转为 Excel
 *   - convertToMarkdown() — 转为 Markdown
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';

    /**
     * 格式类型映射
     */
    const FORMAT_TYPES = {
        markdown:  { id: 'markdown',  name: 'Markdown', icon: '📝' },
        plaintext: { id: 'plaintext', name: '纯文本', icon: '📄' },
        raw:       { id: 'raw',       name: 'OCR原文', icon: '📄' },
        html:      { id: 'html',      name: 'HTML',     icon: '🌐' },
        excel:     { id: 'excel',     name: 'Excel 表格', icon: '📊' },
        preview:   { id: 'preview',   name: '预览', icon: '👁' }
    };

    /**
     * 简单的 Markdown 转 HTML（含标题、列表、表格、粗体、代码块）
     */
    function markdownToHtml(mdText) {
        if (!mdText) return '';

        // ---- 预处理：提取特殊块为占位符，避免后续处理干扰 ----
        var placeholders = {};
        var phIdx = 0;

        function storePlaceholder(replacement) {
            var key = '###_PH_' + (phIdx++) + '_PH_###';
            placeholders[key] = replacement;
            return key;
        }

        var html = mdText;

        // ============================================================
        // 关键：先提取所有需要特殊处理的块（$$公式、代码块），再做 HTML 转义
        // 否则转义后的文本无法被正则匹配
        // ============================================================

        // 0️⃣ 检测未被 $$ 包裹的裸 LaTeX 公式并自动包裹
        var latexPattern = /\\(frac|sqrt|sum|prod|int|lim|sin|cos|tan|log|ln|exp|alpha|beta|gamma|delta|theta|lambda|pi|sigma|omega|infty|partial|nabla|cdot|times|div|pm|mp|leq|geq|neq|approx)/;
        var hasDollarDollar = /\$\$/.test(html);
        if (!hasDollarDollar) {
            var lines = html.split('\n');
            for (var i = 0; i < lines.length; i++) {
                if (latexPattern.test(lines[i])) {
                    ChatLog.info('FORMAT', '[FormatConverter] 检测到裸LaTeX:', lines[i].substring(0, 80));
                    lines[i] = '$$\n' + lines[i] + '\n$$';
                }
            }
            html = lines.join('\n');
        }

        // 1️⃣ 【优先】提取 LaTeX 块级公式 $$...$$（必须在 HTML 转义之前！）
        html = html.replace(/\$\$([\s\S]*?)\$\$/g, function(match, formula) {
            formula = formula.trim();
            if (!formula) return match;
            ChatLog.info('FORMAT', '[FormatConverter] 块级公式:', formula.substring(0, 100), '| KaTeX可用:', typeof katex !== 'undefined');
            if (typeof katex !== 'undefined') {
                try {
                    var rendered = katex.renderToString(formula, {
                        throwOnError: false,
                        displayMode: true,
                        output: 'html'
                    });
                    return storePlaceholder(
                        '<div class="ocr-formula-block" style="background:#f8f9fa;border-left:3px solid #4a90d9;' +
                        'padding:12px 16px;margin:10px 0;border-radius:4px;overflow-x:auto;">' +
                        rendered + '</div>'
                    );
                } catch(e) {
                    ChatLog.warn('FORMAT', '[KaTeX] 渲染块级公式失败:', e.message);
                }
            }
            // 降级：保护反斜杠
            var safeFormula = formula.replace(/\\/g, '&#92;');
            return storePlaceholder(
                '<div class="ocr-formula-block" style="background:#f8f9fa;border-left:3px solid #4a90d9;' +
                'padding:10px 14px;margin:8px 0;border-radius:4px;overflow-x:auto;' +
                'font-family:\'Cambria Math\',\'Times New Roman\',serif;font-style:italic;' +
                'font-size:14px;line-height:1.6;">' + safeFormula.replace(/\n/g, '<br>') + '</div>'
            );
        });

        // 2️⃣ 【优先】提取 LaTeX 行内公式 $...$
        html = html.replace(/(?<!\d)\$([^$\n]+?)\$(?!\d)/g, function(match, formula) {
            if (typeof katex !== 'undefined') {
                try {
                    var rendered = katex.renderToString(formula.trim(), {
                        throwOnError: false,
                        displayMode: false,
                        output: 'html'
                    });
                    return storePlaceholder('<span class="ocr-formula-inline" style="padding:0 2px;">' + rendered + '</span>');
                } catch(e) {
                    ChatLog.warn('FORMAT', '[KaTeX] 渲染行内公式失败:', e.message);
                }
            }
            return storePlaceholder('<span class="ocr-formula-inline" style="font-family:\'Cambria Math\',\'Times New Roman\',serif;font-style:italic;padding:0 2px;background:#f5f5f5;border-radius:2px;">' +
                formula.trim().replace(/\\/g, '&#92;') + '</span>');
        });

        // 3️⃣ 提取代码块 ```...``` 为占位符
        html = html.replace(/```(\w*)\n?([\s\S]*?)```/g, function(match, lang, code) {
            return storePlaceholder('<pre><code class="language-' + (lang || '') + '">' +
                   code.trim() + '</code></pre>');
        });

        // 4️⃣ 最后做 HTML 转义（此时 $$ 和代码块已提取为占位符，不会被误转义）
        html = html.replace(/&/g, '&amp;')
                   .replace(/</g, '&lt;')
                   .replace(/>/g, '&gt;');

        // 5. 标题
        html = html.replace(/^### (.+)$/gm, '<h4>$1</h4>');
        html = html.replace(/^## (.+)$/gm, '<h3>$1</h3>');
        html = html.replace(/^# (.+)$/gm, '<h2>$1</h2>');

        // 6. 粗体
        html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

        // 7. 斜体
        html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

        // 8. 分隔线
        html = html.replace(/^---$/gm, '<hr>');

        // 9. 表格：连续的 | 行合成一张表
        html = html.replace(/(?:^\|.+\|\s*\n?)+/gm, function (match) {
            var lines = match.split('\n').map(function (line) { return line.trim(); }).filter(Boolean);
            var rows = lines.filter(function (line) {
                var cells = line.split('|').map(function (cell) { return cell.trim(); }).filter(function (cell) { return cell.length; });
                if (!cells.length) return false;
                return !cells.every(function (cell) { return /^:?-{2,}:?$/.test(cell); });
            });
            if (rows.length < 1) return match;
            var tableHtml = '<table class="ocr-excel-table"><thead><tr>';
            rows[0].split('|').map(function (cell) { return cell.trim(); }).filter(function (cell) { return cell.length; }).forEach(function (cell) {
                tableHtml += '<th>' + cell + '</th>';
            });
            tableHtml += '</tr></thead><tbody>';
            for (var i = 1; i < rows.length; i++) {
                var cells = rows[i].split('|').map(function (cell) { return cell.trim(); }).filter(function (cell) { return cell.length; });
                if (!cells.length) continue;
                tableHtml += '<tr>';
                cells.forEach(function (cell) { tableHtml += '<td>' + cell + '</td>'; });
                tableHtml += '</tr>';
            }
            tableHtml += '</tbody></table>';
            return tableHtml;
        });

        // 10. 无序列表
        html = html.replace(/^(\s*)[-*] (.+)$/gm, function(match, indent, item) {
            return '<p style="padding-left:' + (Math.floor(indent.length / 2) * 20) + 'px">• ' + item + '</p>';
        });

        // 11. 有序列表
        html = html.replace(/^(\s*)\d+\. (.+)$/gm, function(match, indent, item) {
            return '<p style="padding-left:' + (Math.floor(indent.length / 2) * 20) + 'px">' +
                   match.trim().split('.')[0] + '. ' + item + '</p>';
        });

        // 12. 段落
        html = html.replace(/\n\n/g, '</p><p>');

        // 13. 换行
        html = html.replace(/\n/g, '<br>');

        // 14. 已有标题或表格时不再外套段落，避免表格被拆开
        if (!/<(table|h2|h3|h4|hr|pre)\b/i.test(html)) {
            html = '<p>' + html + '</p>';
        }

        // 15. 恢复占位符
        for (var key in placeholders) {
            // 使用 split+join 避免正则特殊字符问题
            html = html.split(key).join(placeholders[key]);
        }

        return html;
    }

    /**
     * 在指定容器中渲染 OCR 结果
     * @param {string} containerId - 容器元素 ID
     * @param {string} text - OCR 识别文本
     * @param {string} format - 格式类型 (plaintext/markdown/html/excel)
     */
    function renderOcrResult(containerId, text, format) {
        const container = document.getElementById(containerId);
        if (!container) return;

        format = format || 'plaintext';

        switch (format) {
            case 'plaintext':
            case 'raw':
                container.innerHTML = '<pre style="white-space:pre-wrap;font-family:Consolas,monospace;font-size:13px;line-height:1.6;margin:0;">' +
                    escapeHtml(text) + '</pre>';
                break;

            case 'preview':
            case 'markdown':
                container.innerHTML = '<div class="ocr-markdown-render" style="font-size:13px;line-height:1.7;padding:8px;">' +
                    markdownToHtml(text) + '</div>';
                break;

            case 'html':
                // HTML 格式：先将纯文本转换为 HTML，再渲染（清理 XSS 风险）
                container.innerHTML = '<div class="ocr-html-render" style="font-size:13px;">' +
                    sanitizeHtml(convert(text, 'html')) + '</div>';
                break;

            case 'excel':
                // Excel 格式显示为表格预览
                container.innerHTML = renderExcelPreview(text);
                break;

            default:
                container.textContent = text;
        }
    }

    /**
     * 渲染 Excel 表格预览（制表符分隔转 HTML 表格）
     */
    function renderExcelPreview(text) {
        const tables = [];
        const re = /<!--\s*HTML_TABLE_B64:([A-Za-z0-9+/=]+)\s*-->/g;
        let match;
        while ((match = re.exec(String(text || '')))) {
            try {
                const bytes = Uint8Array.from(atob(match[1]), function (c) { return c.charCodeAt(0); });
                const html = new TextDecoder('utf-8').decode(bytes);
                if (html) tables.push('<div class="ocr-excel-shot">' + html + '</div>');
            } catch (e) {
                console.warn('[renderExcelPreview] Base64 decode failed:', e);
            }
        }
        if (tables.length) return tables.join('');

        text = String(text || '').replace(/<!--\s*HTML_TABLE_B64:[A-Za-z0-9+/=]+\s*-->/g, '').trim();
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length === 0) return '<p style="color:#999">无数据</p>';

        // 跳过注释/元数据行（以 # 开头或只包含键值对的说明行）
        const dataLines = lines.filter(l => {
            const trim = l.trim();
            if (trim.startsWith('#')) return false;
            // 允许包含制表符的数据行；跳过类似"来源文件: xxx"的说明行
            if (!trim.includes('\t') && /^[^\t]+[：:]\s*/.test(trim)) return false;
            return true;
        });

        // 如果数据行太少，直接显示
        if (dataLines.length < 1 || !dataLines.some(l => l.includes('\t'))) {
            return '<pre style="white-space:pre-wrap;font-family:Consolas,monospace;font-size:12px;">' +
                   escapeHtml(text) + '</pre>';
        }

        let html = '<table class="ocr-excel-table" style="border-collapse:collapse;width:100%;font-size:12px;">';

        let startIdx = 0;
        // 第一行可能是表头
        html += '<thead><tr>';
        const headers = dataLines[0].split('\t');
        headers.forEach(h => {
            html += '<th style="border:1px solid #ccc;padding:6px 10px;background:#f0f4ff;text-align:left;font-weight:600;">' +
                    escapeHtml(h.trim()) + '</th>';
        });
        html += '</tr></thead><tbody>';

        for (let i = 1; i < dataLines.length; i++) {
            const cells = dataLines[i].split('\t');
            if (cells.length === 0 || (cells.length === 1 && !cells[0].trim())) continue;
            html += '<tr>';
            cells.forEach(c => {
                html += '<td style="border:1px solid #e0e0e0;padding:5px 10px;">' +
                        escapeHtml(c.trim()) + '</td>';
            });
            html += '</tr>';
        }

        html += '</tbody></table>';
        return html;
    }

    /**
     * 获取格式选择器的 HTML
     * @param {string} selectedFormat - 当前选中的格式
     * @returns {string} HTML 字符串
     */
    function getFormatSelectorHtml(selectedFormat) {
        selectedFormat = selectedFormat || 'plaintext';
        let html = '<div class="ocr-format-selector" style="display:flex;align-items:center;gap:6px;">';
        html += '<span class="toolbar-group-label" style="font-size:11px;color:#999;">输出格式</span>';
        html += '<select id="ocr-format-select" class="form-select" style="max-width:90px;">';

        for (const key in FORMAT_TYPES) {
            const fmt = FORMAT_TYPES[key];
            const selected = key === selectedFormat ? 'selected' : '';
            html += '<option value="' + key + '" ' + selected + '>' + fmt.icon + ' ' + fmt.name + '</option>';
        }

        html += '</select></div>';
        return html;
    }

    /**
     * 转义 HTML
     */
    function escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * 简单 HTML 清洗（仅允许安全的标签）
     */
    function sanitizeHtml(html) {
        if (!html) return '';
        // 允许常用标签
        const allowedTags = ['p', 'br', 'b', 'i', 'u', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6',
                             'table', 'thead', 'tbody', 'tr', 'th', 'td', 'hr', 'pre', 'code',
                             'ul', 'ol', 'li', 'div', 'span', 'strong', 'em', 'img'];
        // 移除 script, iframe, style 等危险标签
        html = html.replace(/<script[\s\S]*?<\/script>/gi, '');
        html = html.replace(/<iframe[\s\S]*?<\/iframe>/gi, '');
        html = html.replace(/<style[\s\S]*?<\/style>/gi, '');
        html = html.replace(/on\w+\s*=\s*["'][^"']*["']/gi, '');
        return html;
    }

    /**
     * 获取格式化文本的 MIME 类型（用于导出）
     */
    function getFormatMimeType(format) {
        const mimes = {
            plaintext: 'text/plain',
            markdown:  'text/markdown',
            html:      'text/html',
            excel:     'text/csv'
        };
        return mimes[format] || 'text/plain';
    }

    /**
     * 获取格式化文本的文件扩展名
     */
    function getFormatExtension(format) {
        const exts = {
            plaintext: '.txt',
            markdown:  '.md',
            html:      '.html',
            excel:     '.csv'
        };
        return exts[format] || '.txt';
    }

    /**
     * 将原始文本转换为指定格式
     * @param {string} text - 原始文本（可能是多行文本或制表符分隔数据）
     * @param {string} format - 目标格式 (plaintext/markdown/html/excel)
     * @returns {string} 转换后的文本
     */
    function convert(text, format) {
        if (!text) return '';
        format = format || 'plaintext';

        switch (format) {
            case 'plaintext':
            case 'raw':
                return text;

            case 'markdown':
                return textToMarkdown(text);

            case 'html':
                return markdownToHtml(textToMarkdown(text));

            case 'excel':
                return text; // Excel 格式保持原始文本（制表符分隔）

            default:
                return text;
        }
    }

    /**
     * 将原始文本转换为 Markdown 表格格式
     */
    function textToMarkdown(text) {
        const lines = text.split('\n').filter(l => l.trim());
        if (lines.length === 0) return text;

        // 检测是否包含制表符（表格数据）
        const hasTabs = lines.some(l => l.includes('\t'));
        if (!hasTabs) return text; // 纯文本无表格结构，原样返回

        // 将制表符分隔的数据转换为 Markdown 表格
        let md = '';
        lines.forEach((line, i) => {
            const cells = line.split('\t').map(c => c.trim());
            md += '| ' + cells.join(' | ') + ' |\n';
            if (i === 0) {
                // 添加分隔行
                md += '| ' + cells.map(() => '---').join(' | ') + ' |\n';
            }
        });
        return md;
    }

    // 导出公共 API
    window.OcrFormatConverter = {
        FORMAT_TYPES: FORMAT_TYPES,
        renderOcrResult: renderOcrResult,
        renderExcelPreview: renderExcelPreview,
        markdownToHtml: markdownToHtml,
        convert: convert,
        textToMarkdown: textToMarkdown,
        getFormatSelectorHtml: getFormatSelectorHtml,
        escapeHtml: escapeHtml,
        sanitizeHtml: sanitizeHtml,
        getFormatMimeType: getFormatMimeType,
        getFormatExtension: getFormatExtension,
        isLatexGarbage: function(text) { return false; },
        cleanLatexGarbage: function(text) { return text; }
    };

})();
