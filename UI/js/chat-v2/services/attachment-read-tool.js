/**
 * read_attachment — 读取本轮上传附件的正文（文本层）。
 * recognize_pages — 文本层为空时，识别已经渲染好的页面图片。
 * 两者和 excel_read_range 一样，都是规划器能选的一步。整轮留在文档管道里。
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    var TOOL_NAME = 'read_attachment';
    var RECOGNIZE_NAME = 'recognize_pages';
    var MAX_PAGES = 3;
    var OCR_BUDGET_MS = 60000;
    var DIM_LINE_RE = /(\d+(?:\.\d+)?)\s*(?:mm|cm|inch|in|°|±)|公差|尺寸|标题栏|\bDIA\b|Ø|φ|TOLERANCE|\bSCALE\b|\bMATERIAL\b/i;

    function fileLabel(att, index) {
        return (att && (att.fileName || att.name)) || ('附件' + (index + 1));
    }

    /** 扫描件：即使页面图还没渲，也要进入识别。 */
    function textNeedsRecognition(content) {
        var s = String(content || '');
        return s.indexOf('[扫描版PDF]') >= 0 || s.indexOf('没有可提取的文本层') >= 0;
    }

    /**
     * 尺寸和标题栏排到前面，避免 4000 字回喂截断时把关键数字切掉。
     */
    function promoteDrawingText(text) {
        var lines = String(text || '').split('\n');
        var hit = [];
        var rest = [];
        for (var i = 0; i < lines.length; i++) {
            if (DIM_LINE_RE.test(lines[i])) hit.push(lines[i]);
            else rest.push(lines[i]);
        }
        if (!hit.length) return String(text || '');
        return hit.join('\n') + '\n\n' + rest.join('\n');
    }

    function shouldOfferRecognize(attachments) {
        if (hasPageImages(attachments)) return true;
        return (Array.isArray(attachments) ? attachments : []).some(function (a) {
            return a && a.type !== 'image' && textNeedsRecognition(a.content || a.text || '');
        });
    }

    /** 问题在问图纸尺寸，但文字层里没有尺寸行时，仍要识别页面。 */
    function drawingAskNeedsOcr(ask, text, fileName) {
        var q = String(ask || '');
        if (!/尺寸|公差|标题栏|图纸/.test(q)) return false;
        if (!/\.pdf$/i.test(String(fileName || ''))) return false;
        if (textNeedsRecognition(text)) return true;
        return !DIM_LINE_RE.test(String(text || ''));
    }

    function shouldAutoRecognize(toolName, result, alsoCalled) {
        if (String(toolName || '') !== TOOL_NAME) return false;
        var queued = Array.isArray(alsoCalled) ? alsoCalled : [];
        for (var i = 0; i < queued.length; i++) {
            if (String(queued[i] || '').toLowerCase() === RECOGNIZE_NAME) return false;
        }
        return !!(result && result.ok && result.result && result.result.needsRecognition);
    }

    function pageImages(attachments) {
        return (Array.isArray(attachments) ? attachments : []).filter(function (a) {
            return a && a.type === 'image' && (a.base64 || a.data);
        });
    }

    function hasPageImages(attachments) {
        return pageImages(attachments).length > 0;
    }

    /**
     * 点名时按文件名包含关系过滤；最多前 3 页。不改原附件。
     */
    function selectPages(attachments, args) {
        var images = pageImages(attachments);
        var want = args && (args.fileName || args.name) ? String(args.fileName || args.name).trim() : '';
        if (want && images.length) {
            var hit = images.filter(function (a) {
                return fileLabel(a, 0).indexOf(want) >= 0;
            });
            if (hit.length) images = hit;
        }
        return images.slice(0, MAX_PAGES);
    }

    /**
     * 识别步骤才渲染扫描页。已有页面图则直接用。失败只记在文档说明上。
     */
    async function ensurePageImages(turn) {
        if (!turn) return [];
        var atts = Array.isArray(turn.attachments) ? turn.attachments : [];
        if (pageImages(atts).length) return atts;
        var render = (App.ChatV2 && App.ChatV2.pdfPageRender) || (typeof window !== 'undefined' && window.PdfPageRender);
        var entry = App.ChatV2 && App.ChatV2.chatEntry;
        for (var i = 0; i < atts.length; i++) {
            var doc = atts[i];
            if (!doc || !doc.sourceFile || !textNeedsRecognition(doc.content || '')) continue;
            if (!render || typeof render.renderPdfPages !== 'function') {
                doc.content += '\n补充：页面图片未能生成（页面渲染器未就绪）。';
                continue;
            }
            var imgs = [];
            try {
                var job = render.renderPdfPages(doc.sourceFile, {
                    maxPages: MAX_PAGES,
                    onProgress: function (info) {
                        var page = info && info.page;
                        var total = info && info.total;
                        if (page && total && entry && typeof entry.setPhaseStatus === 'function') {
                            entry.setPhaseStatus('正在渲染「' + (doc.fileName || '附件') + '」第 ' + page + '/' + total + ' 页…');
                        }
                    }
                });
                if (entry && typeof entry.withCallTimeout === 'function') {
                    imgs = await entry.withCallTimeout(job, entry.PDF_RENDER_TIMEOUT_MS || 45000, '扫描版 PDF 页面渲染超时');
                } else {
                    imgs = await job;
                }
            } catch (e) {
                doc.content += '\n补充：页面图片未能生成（' + ((e && e.message) || '渲染失败') + '）。';
                continue;
            }
            if (!imgs || !imgs.length) {
                doc.content += '\n补充：页面图片未能生成。';
                continue;
            }
            for (var p = 0; p < imgs.length; p++) {
                var img = imgs[p];
                atts.push({
                    type: 'image',
                    mime: img.mime || 'image/png',
                    fileName: (doc.fileName || '附件') + ' 第' + (img.pageNumber || (p + 1)) + '页',
                    base64: img.base64,
                    size: Math.floor(String(img.base64 || '').length * 0.75)
                });
            }
        }
        turn.attachments = atts;
        return atts;
    }

    /**
     * @param {Array} attachments
     * @param {{fileName?: string}} [args]
     * @returns {{ok:boolean, result?:object, error?:string}}
     */
    function readContent(attachments, args) {
        var list = Array.isArray(attachments) ? attachments.filter(function (a) { return a && a.type !== 'image'; }) : [];
        var images = Array.isArray(attachments) ? attachments.filter(function (a) { return a && a.type === 'image'; }) : [];
        if (!list.length && !images.length) {
            return { ok: false, error: '本轮没有上传附件。读取当前表格请调用 excel_read_range。' };
        }
        var want = args && (args.fileName || args.name) ? String(args.fileName || args.name).trim() : '';
        var chosen = list;
        if (want && list.length) {
            var hit = list.filter(function (a) {
                return fileLabel(a, 0).indexOf(want) >= 0;
            });
            if (hit.length) chosen = hit;
        }
        if (!chosen.length) {
            return {
                ok: true,
                result: {
                    success: true,
                    fileCount: 0,
                    imageCount: images.length,
                    needsRecognition: images.length > 0,
                    message: images.length
                        ? ('本轮只有 ' + images.length + ' 张图片，没有可提取的文档正文。请调用 recognize_pages。')
                        : '没有匹配到附件。'
                }
            };
        }
        var files = chosen.map(function (a, i) {
            return {
                fileName: fileLabel(a, i),
                mime: a.mime || '',
                content: promoteDrawingText(String(a.content || a.text || ''))
            };
        });
        var message = files.map(function (f) {
            return '【' + f.fileName + '】\n' + (f.content || '（该文件没有提取到文字）');
        }).join('\n\n');
        var needsRecognition = files.some(function (f) { return textNeedsRecognition(f.content); });
        if (needsRecognition) {
            message += '\n\n正文没有文字层。系统会接着识别前 ' + MAX_PAGES + ' 页，不要根据文件名编造内容。';
        }
        return {
            ok: true,
            result: {
                success: true,
                fileCount: files.length,
                imageCount: images.length,
                needsRecognition: needsRecognition,
                files: files,
                message: message
            }
        };
    }

    function stripBase64Prefix(data) {
        if (!data) return '';
        var commaIdx = data.indexOf(',');
        if (commaIdx !== -1 && String(data).indexOf('data:') === 0) {
            return data.substring(commaIdx + 1);
        }
        return data;
    }

    /**
     * @param {Array} attachments
     * @param {{fileName?: string, question?: string}} [args]
     * @param {{isVision?: boolean, askVision?: Function, recognizeWithOcr?: Function, signal?: AbortSignal}} [deps]
     */
    async function recognizeContent(attachments, args, deps) {
        var images = selectPages(attachments, args);
        if (!images.length) {
            return { ok: false, error: '本轮没有可识别的页面图片。有文字层时请直接使用 read_attachment 返回的正文。' };
        }
        var options = deps || {};
        var isVision = typeof options.isVision === 'boolean'
            ? options.isVision
            : (typeof App.isVisionModel === 'function' && App.isVisionModel((App.config && App.config.model) || ''));
        if (isVision) {
            return {
                ok: true,
                result: {
                    success: true,
                    mode: 'vision',
                    pageCount: images.length,
                    message: '已附上 ' + images.length + ' 页图纸（最长边不超过 1600 像素）。请直接读图中的尺寸、公差和标题栏，只写图上有的内容。',
                    visionImages: images.map(function (img, i) {
                        return {
                            fileName: fileLabel(img, i),
                            mime: img.mime || 'image/png',
                            base64: stripBase64Prefix(img.base64 || img.data || '')
                        };
                    })
                }
            };
        }
        var ocrFn = options.recognizeWithOcr
            || (App.ChatV2.pageRecognition && App.ChatV2.pageRecognition.recognizeWithOcr);
        if (typeof ocrFn !== 'function') {
            return { ok: false, error: '页面识别未就绪，无法对扫描页做 OCR。' };
        }
        var pack = await ocrFn(images, { maxPages: MAX_PAGES, budgetMs: OCR_BUDGET_MS });
        var pages = (pack && pack.pages) || [];
        var body = pages.map(function (p) {
            return '【' + (p.name || '页面') + '】\n' + (p.text || '（未识别到文字）');
        }).join('\n\n');
        return {
            ok: true,
            result: {
                success: true,
                mode: pack && pack.usedVision ? 'ocr+vision' : 'ocr',
                pageCount: pages.length,
                message: promoteDrawingText(body) || 'OCR 没有从图中读出文字。请如实说明没有识别到内容，不要编造尺寸。'
            }
        };
    }

    function readDefinition() {
        return {
            name: TOOL_NAME,
            displayName: '读取附件',
            description: '读取用户本轮上传的附件正文（PDF、Word、Excel 文件、文本、图纸的文字层）。问图纸、上传文件、附件里的尺寸或文字时先调用本工具。没有文字层时再调用 recognize_pages。读取当前打开的 Excel 工作表或选区时用 excel_read_range。',
            intentKeywords: ['附件', '图纸', 'pdf', '上传', '文件内容', '文档'],
            inputSchema: {
                type: 'object',
                properties: {
                    fileName: {
                        type: 'string',
                        description: '要读取的附件文件名。不传则返回本轮全部文档附件。'
                    }
                },
                required: []
            },
            source: 'builtin-attachment',
            presentCall: function (args) {
                var name = args && (args.fileName || args.name);
                return { card: 'generic', title: '读取附件', kind: 'read', content: name ? String(name) : '本轮上传的文件' };
            },
            presentResult: function (args, result) {
                var n = 0;
                try {
                    var r = result && result.result ? result.result : result;
                    n = (r && (r.fileCount || (r.files && r.files.length))) || 0;
                } catch (_e) { n = 0; }
                return {
                    card: 'generic',
                    title: '读取附件',
                    kind: 'read',
                    content: n ? ('已读取 ' + n + ' 个附件的正文') : '已返回附件内容'
                };
            },
            execute: async function (args, ctx) {
                var turn = ctx && ctx.turn;
                var atts = (turn && turn.attachments) || [];
                return readContent(atts, args || {});
            }
        };
    }

    function recognizeDefinition() {
        return {
            name: RECOGNIZE_NAME,
            displayName: '识别页面',
            description: '识别本轮已渲染的附件页面图片（扫描版 PDF、没有文字层的图纸）。视觉模型直接看图；普通模型对原图做 OCR，识别太差时再交视觉兜底。只在 read_attachment 写明没有文字层时调用。读取当前 Excel 用 excel_read_range。',
            intentKeywords: ['识别', '扫描', '图纸', '页面', '尺寸', 'OCR'],
            inputSchema: {
                type: 'object',
                properties: {
                    fileName: {
                        type: 'string',
                        description: '要识别的附件文件名。不传则识别本轮已渲染的前 3 页。'
                    }
                },
                required: []
            },
            source: 'builtin-attachment',
            presentCall: function (args) {
                var name = args && (args.fileName || args.name);
                return { card: 'generic', title: '识别页面', kind: 'read', content: name ? String(name) : '已渲染的页面' };
            },
            presentResult: function (args, result) {
                var n = 0;
                var mode = '';
                try {
                    var r = result && result.result ? result.result : result;
                    n = (r && r.pageCount) || 0;
                    mode = (r && r.mode) || '';
                } catch (_e) { n = 0; }
                var how = mode === 'vision' ? '视觉' : (mode === 'ocr+vision' ? 'OCR 后视觉兜底' : 'OCR');
                return {
                    card: 'generic',
                    title: '识别页面',
                    kind: 'read',
                    content: n ? ('已用' + how + '识别 ' + n + ' 页') : '已返回识别结果'
                };
            },
            execute: async function (args, ctx) {
                var turn = ctx && ctx.turn;
                if (turn && turn._pageRecognitionCache) {
                    return { ok: true, result: turn._pageRecognitionCache };
                }
                var atts = await ensurePageImages(turn);
                var out = await recognizeContent(atts, args || {}, { signal: ctx && ctx.signal });
                if (out && out.ok && out.result && turn) {
                    var cached = {};
                    Object.keys(out.result).forEach(function (key) {
                        if (key !== 'visionImages') cached[key] = out.result[key];
                    });
                    turn._pageRecognitionCache = cached;
                }
                return out;
            }
        };
    }

    var _waitTimer = null;
    function register() {
        var reg = App.ChatV2 && App.ChatV2.toolRegistry;
        if (!reg || typeof reg.register !== 'function') return false;
        reg.register(readDefinition());
        reg.register(recognizeDefinition());
        if (_waitTimer) { clearInterval(_waitTimer); _waitTimer = null; }
        return true;
    }

    /** 规划步骤：先读附件；有页面图时加上识别；再读当前 Excel。 */
    function planningTools(reg, turn) {
        register();
        var registry = reg || (App.ChatV2 && App.ChatV2.toolRegistry);
        if (!registry || typeof registry.get !== 'function') return [];
        var names = [TOOL_NAME];
        if (shouldOfferRecognize(turn && turn.attachments)) names.push(RECOGNIZE_NAME);
        names.push('excel_read_range');
        var out = [];
        for (var i = 0; i < names.length; i++) {
            var t = registry.get(names[i]);
            if (t) out.push(t);
        }
        return out;
    }

    /** 本轮有文档附件时，把读取/识别放进模型可见工具表的前面。 */
    function pinOn(tools, turn) {
        var atts = turn && turn.attachments;
        var hasDoc = Array.isArray(atts) && atts.some(function (a) { return a && a.type !== 'image'; });
        if (!hasDoc) return tools || [];
        register();
        var reg = App.ChatV2 && App.ChatV2.toolRegistry;
        var frontNames = [TOOL_NAME];
        if (shouldOfferRecognize(atts)) frontNames.push(RECOGNIZE_NAME);
        var front = [];
        var skip = {};
        for (var i = 0; i < frontNames.length; i++) {
            var tool = reg && typeof reg.get === 'function' ? reg.get(frontNames[i]) : null;
            if (!tool) continue;
            front.push(tool);
            skip[String(tool.name || '').toLowerCase()] = true;
        }
        if (!front.length) return tools || [];
        var rest = (tools || []).filter(function (t) {
            return t && !skip[String(t.name || '').toLowerCase()];
        });
        return front.concat(rest);
    }

    App.ChatV2.attachmentReadTool = {
        name: TOOL_NAME,
        recognizeName: RECOGNIZE_NAME,
        readContent: readContent,
        recognizeContent: recognizeContent,
        selectPages: selectPages,
        ensurePageImages: ensurePageImages,
        promoteDrawingText: promoteDrawingText,
        shouldAutoRecognize: shouldAutoRecognize,
        drawingAskNeedsOcr: drawingAskNeedsOcr,
        register: register,
        planningTools: planningTools,
        pinOn: pinOn
    };

    if (!register()) {
        var tries = 0;
        _waitTimer = setInterval(function () {
            tries++;
            if (register() || tries > 40) {
                if (_waitTimer) { clearInterval(_waitTimer); _waitTimer = null; }
            }
        }, 200);
    }
})();
