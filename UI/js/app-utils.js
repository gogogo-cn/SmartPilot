/**
 * ============================================================
 * app-utils.js — Utility Functions
 * ============================================================
 * 功能：通用工具函数集合，提供字符串处理、日期格式化、DOM 操作、Excel 地址解析等基础能力。
 *
 * 核心类/函数：
 *   - escapeHtml() — HTML 转义
 *   - formatDate() — 日期格式化
 *   - debounce() / throttle() — 防抖节流
 *   - parseExcelRange() — 解析 Excel 地址
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    // --- 来自 app.js:L530-L553 ---
    function ensureMarkedLoaded() {
        return new Promise((resolve, reject) => {
            // marked v15+ 是对象 { parse, setOptions... }，不是函数，用 !== 'undefined' 检测
            if (typeof marked !== 'undefined') return resolve();

            // 检查是否已有 script 引用 marked.min.js
            const existing = Array.from(document.getElementsByTagName('script')).find(s => s.src && s.src.indexOf('marked.min.js') !== -1);
            if (existing) {
                if (existing.getAttribute('data-loaded') === '1' || typeof marked !== 'undefined') return resolve();
                existing.addEventListener('load', () => { existing.setAttribute('data-loaded', '1'); resolve(); });
                existing.addEventListener('error', () => reject(new Error('marked load error')));
                setTimeout(() => { if (typeof marked !== 'undefined') resolve(); }, 50);
                return;
            }

            // 动态注入
            const s = document.createElement('script');
            s.src = 'js/marked.min.js';
            s.async = false;
            s.onload = () => { s.setAttribute('data-loaded', '1'); resolve(); };
            s.onerror = () => reject(new Error('marked load error'));
            document.head.appendChild(s);
        });
    }
    App.ensureMarkedLoaded = ensureMarkedLoaded;

    // --- 来自 app.js:L560-L576 ---
    function sanitizeHtml(html) {
        try {
            const dp = new DOMParser();
            const doc = dp.parseFromString(html, 'text/html');
            // 移除 script
            doc.querySelectorAll('script').forEach(n => n.remove());
            // 移除事件处理器等不安全属性
            doc.querySelectorAll('*').forEach(el => {
                for (const attr of Array.from(el.attributes)) {
                    if (/^on/i.test(attr.name)) el.removeAttribute(attr.name);
                }
            });
            return doc.body.innerHTML;
        } catch (e) {
            return '';
        }
    }

    // --- 来自 app.js:L11073-L11100 ---
    // HTML 转义函数
    // preserveBr: 是否保留 <br> 标签不转义
    function escapeHtml(text, preserveBr = true) {
        if (!text) return '';

        if (preserveBr) {
            // 先保存 <br> 标签
            const brPlaceholder = '\x00BR\x00';
            let processed = text.replace(/<br>/gi, brPlaceholder);

            // 转义 HTML
            processed = processed
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');

            // 恢复 <br> 标签
            return processed.replace(new RegExp(brPlaceholder, 'g'), '<br>');
        } else {
            // 完全转义所有 HTML
            return text
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#039;');
        }
    }
    App.escapeHtml = escapeHtml;

    // --- 来自 app.js:L288-L290 ---
    function escapeRegExp(string) {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    // --- 来自 app.js:L9649-L9678 ---
    function safeJsonParse(content) {
        if (!content || typeof content !== 'string') {
            throw new Error('内容为空或不是字符串');
        }

        try {
            // 尝试直接解析
            return JSON.parse(content);
        } catch (e) {
            // 尝试提取JSON（处理markdown代码块）
            const patterns = [
                /```(?:json)?\s*([\s\S]*?)```/,  // markdown代码块: ```json ... ```
                /\{[\s\S]*\}/                       // 裸JSON对象
            ];

            for (const pattern of patterns) {
                const match = content.match(pattern);
                if (match) {
                    try {
                        const jsonStr = match[1] || match[0];
                        return JSON.parse(jsonStr.trim());
                    } catch (innerError) {
                        continue;  // 尝试下一个模式
                    }
                }
            }

            throw new Error('无法从响应内容中提取有效的JSON格式');
        }
    }


    // --- 来自 app.js:L4625-L4647 ---
    function resizeBase64Image(base64Src, maxWidth, maxHeight) {
        try {
            const canvas = document.createElement('canvas');
            const ctx = canvas.getContext('2d');
            const img = new Image();
            img.src = base64Src;
            // 同步获取尺寸（base64 已加载完成）
            let w = img.naturalWidth || img.width || 200;
            let h = img.naturalHeight || img.height || 150;
            if (w > maxWidth || h > maxHeight) {
                const ratio = Math.min(maxWidth / w, maxHeight / h);
                w = Math.round(w * ratio);
                h = Math.round(h * ratio);
            }
            canvas.width = w;
            canvas.height = h;
            ctx.drawImage(img, 0, 0, w, h);
            return canvas.toDataURL('image/jpeg', 0.7);
        } catch (e) {
            ChatLog.warn('UTILS', '[缩略图] 生成失败:', e.message);
            return null;
        }
    }

    // --- 来自 app.js:L18999-L19004 ---
    function extractCodeBlock(text, language = 'vba') {
        // 支持提取多种语言的代码块
        const codeRegex = new RegExp('```' + language + '\\s*([\\s\\S]*?)```');
        const match = text.match(codeRegex);
        return match ? match[1].trim() : null;
    }

    // --- 来自 app.js:L19011-L19114 ---
    // 注意：此函数依赖 convertJsonToTableCard()
    /**
     * 自动检测纯JSON内容并包装成表格样式或代码块
     * @param {string} text - 文本内容
     * @returns {string} 处理后的内容
     */
    function autoDetectAndWrapJSON(text) {
        // 先解码HTML实体
        const decodedText = text.replace(/&quot;/g, '"')
                                .replace(/&amp;/g, '&')
                                .replace(/&lt;/g, '<')
                                .replace(/&gt;/g, '>')
                                .replace(/&#39;/g, "'");
        const trimmed = decodedText.trim();

        // 检查是否包含HTML标签
        if (/<[^>]+>/.test(trimmed)) {
            return text; // 已解码的内容包含HTML，返回原文本
        }

        // 情况1：完整的JSON对象或数组（文本以{或[开头，可能有括号）
        const cleanText = trimmed.replace(/^[\(\[\{\s]+/, '').replace(/[\)\]\}\s]+$/, '');
        if ((cleanText.startsWith('{') && cleanText.endsWith('}')) ||
            (cleanText.startsWith('[') && cleanText.endsWith(']'))) {
            try {
                const parsed = JSON.parse(cleanText);
                // 数组：多维表格
                if (Array.isArray(parsed) && parsed.length > 0) {
                    return convertJsonToTableCard(parsed); // 直接渲染整个数组
                }
                // 对象：键值对表格
                if (typeof parsed === 'object' && !Array.isArray(parsed)) {
                    return convertJsonToTableCard(parsed); // 只返回表格，不保留原文本
                }
                return '```json\n' + cleanText + '\n```';
            } catch (e) {
                // 不是有效的JSON，继续检查其他模式
            }
        }

        // 情况2：从文本中提取JSON对象（处理混合内容）
        // 优先尝试匹配JSON数组（处理[{...}, {...}]格式）
        const jsonArrayMatch = trimmed.match(/([\s\S]*?)\[([\s\S]*?)\]([\s\S]*)/);
        if (jsonArrayMatch) {
            let beforeText = jsonArrayMatch[1].trim();
            const jsonArrayContent = '[' + jsonArrayMatch[2] + ']';
            let afterText = jsonArrayMatch[3].trim();

            try {
                const parsed = JSON.parse(jsonArrayContent);
                if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object') {
                    // 清空JSON数组包装符号
                    if (/^\[\s*\{?\s*$/.test(beforeText) || /^\[\s*\(\s*$/.test(beforeText)) {
                        beforeText = '';
                    }
                    if (/^\}\s*\]\s*$/.test(afterText) || /^\)\s*\]\s*$/.test(afterText)) {
                        afterText = '';
                    }
                    return convertJsonToTableCard(parsed, beforeText || null, afterText || null);
                }
            } catch (e) {
                // 提取的内容不是有效的JSON数组
            }
        }

        // 情况2b：提取单个JSON对象（处理{...}格式）
        // 查找 { 和 } 之间的内容，支持前面有括号等符号
        const jsonMatch = trimmed.match(/([\s\S]*?)\{([\s\S]*?)\}([\s\S]*)/);
        if (jsonMatch) {
            let beforeText = jsonMatch[1].trim();
            const jsonContent = '{' + jsonMatch[2] + '}';
            let afterText = jsonMatch[3].trim();

            try {
                const parsed = JSON.parse(jsonContent);
                // 如果是对象类型且有多个键值对，转换为表格卡片样式
                if (typeof parsed === 'object' && !Array.isArray(parsed) && Object.keys(parsed).length >= 2) {
                    // 如果前后文本是JSON数组包装，则去掉
                    if (/^\[\s*\{?\s*$/.test(beforeText) || /^\[\s*\(\s*$/.test(beforeText)) {
                        beforeText = '';
                    }
                    if (/^\}\s*\]\s*$/.test(afterText) || /^\)\s*\]\s*$/.test(afterText)) {
                        afterText = '';
                    }
                    return convertJsonToTableCard(parsed, beforeText || null, afterText || null);
                }
            } catch (e) {
                // 提取的内容不是有效的JSON
            }
        }

        // 情况3：不完整的JSON（缺少外层{}），如："key":"value","key2":"value2"
        // 检测是否包含多个"key":"value"模式
        const keyValuePattern = /"[^"]*"\s*:\s*"[^"]*"/g;
        const matches = trimmed.match(keyValuePattern);

        if (matches && matches.length >= 2) {
            // 尝试包装成完整的JSON对象
            const wrappedJSON = '{' + trimmed + '}';
            try {
                const parsed = JSON.parse(wrappedJSON);
                // 转换为表格卡片样式
                return convertJsonToTableCard(parsed);
            } catch (e) {
                // 包装后仍不是有效的JSON
            }
        }

        return text;
    }

    // --- 来自 app.js:L21861-L21892 ---
    // 注意：此函数依赖 escapeHtml()
    function showToast(message, type = 'info', duration) {
        const toast = document.getElementById('toast');

        // 根据类型设置样式
        const typeStyles = {
            success: { bg: '#10b981', icon: '✅' },
            error: { bg: '#ef4444', icon: '❌' },
            warning: { bg: '#f59e0b', icon: '⚠️' },
            info: { bg: '#3b82f6', icon: 'ℹ️' }
        };

        const style = typeStyles[type] || typeStyles.info;
        toast.style.background = style.bg;

        // 处理多行文本
        const lines = message.split('\n').filter(line => line.trim() !== '');
        if (lines.length > 1) {
            toast.innerHTML = lines.map((line, index) => {
                const escapedLine = escapeHtml(line);
                return index === 0 ?
                    `<div style="font-weight: 600; margin-bottom: 4px;">${style.icon} ${escapedLine}</div>` :
                    `<div style="font-size: 13px; opacity: 0.95;">${escapedLine}</div>`;
            }).join('');
        } else {
            toast.textContent = (style.icon ? style.icon + ' ' : '') + message;
        }

        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, duration || 3000);
    }

    // --- 来自 app.js:L21906-L21917 ---
    // 注意：此函数依赖 showToast() 和 fallbackCopyToClipboard()
    function copyToClipboard(text, showToastMsg) {
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(() => {
                showToast(showToastMsg || '已复制到剪贴板');
            }).catch(() => {
                // 降级方案
                fallbackCopyToClipboard(text, showToastMsg);
            });
        } else {
            fallbackCopyToClipboard(text, showToastMsg);
        }
    }

    // --- 来自 app.js:L21919-L21933 ---
    // 注意：此函数依赖 showToast()
    function fallbackCopyToClipboard(text, showToastMsg) {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.style.position = 'fixed';
        textarea.style.opacity = '0';
        document.body.appendChild(textarea);
        textarea.select();
        try {
            document.execCommand('copy');
            showToast(showToastMsg || '已复制到剪贴板');
        } catch (err) {
            showToast(showToastMsg || '复制失败，请手动复制');
        }
        document.body.removeChild(textarea);
    }


    // --- 来自 app.js:L18477-L18490 ---
    function formatMessageTime(date) {
        const now = new Date();
        const isToday = date.toDateString() === now.toDateString();
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');

        if (isToday) {
            return `${hours}:${minutes}`;
        } else {
            const month = (date.getMonth() + 1).toString().padStart(2, '0');
            const day = date.getDate().toString().padStart(2, '0');
            return `${month}-${day} ${hours}:${minutes}`;
        }
    }

    // 导出到 window.App 命名空间
    window.App = window.App || {};
    window.App.utils = {
        ensureMarkedLoaded,
        sanitizeHtml,
        escapeHtml,
        escapeRegExp,
        safeJsonParse,
        resizeBase64Image,
        extractCodeBlock,
        autoDetectAndWrapJSON,
        copyToClipboard,
        fallbackCopyToClipboard,
        showToast,
        formatMessageTime
    };

})();
