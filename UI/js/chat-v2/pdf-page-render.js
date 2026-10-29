/**
 * ============================================================
 * chat-v2/pdf-page-render.js — 扫描版 PDF 的页面渲染兜底
 * ============================================================
 * 职责：
 *   把 PDF 前 N 页渲染为 PNG base64，作为 image 附件进入 chat-v2 附件链，
 *   复用既有 image-pipeline（视觉模型直传图片 / 非视觉模型走后端 OCR）。
 *
 * 背景（2026-10-18 用户截图）：
 *   CAD 图纸 PDF 上传后 AI 只会答"需转换为文本/表格后才能解析"。
 *   文本层提取由后端 PdfTextExtractor 负责（Services/Parsers/PdfTextExtractor.vb）；
 *   对**没有文本层的扫描件**，前端渲染是最后一道兜底。
 *
 * 设计约束：
 *   - pdf.js 3.11.174 放在 vendor/pdfjs，不访问 CDN。库未加载时
 *     pdfjsLib 未定义 → 所有入口安全降级返回 []，绝不影响文字链路。
 *   - 页数/边长双上限：防大 PDF 把 base64 撑爆上下文。
 *   - 白底填充：透明背景的 PDF 页直接 toDataURL 会得到黑图。
 * ============================================================
 */
(function() {
    'use strict';

    var PDFJS_WORKER = 'vendor/pdfjs/pdf.worker.min.js';

    /** 默认渲染参数 */
    var DEFAULTS = {
        maxPages: 3,        // 最多渲染页数（扫描件兜底，3 页通常够覆盖首页+关键页）
        maxSidePx: 1600,    // 最长边像素上限（视觉模型甜点区，也控制 base64 体积）
        mime: 'image/png'
    };

    function isAvailable() {
        return typeof pdfjsLib !== 'undefined';
    }

    function ensureWorker() {
        if (!isAvailable()) return false;
        try {
            if (!pdfjsLib.GlobalWorkerOptions.workerSrc) {
                pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER;
            }
        } catch (_) { /* worker 配置失败不致命：主线程解析仍可尝试 */ }
        return true;
    }

    /**
     * 渲染 PDF 前 N 页为图片
     * @param {File} file PDF 文件对象
     * @param {{maxPages?:number, maxSidePx?:number, mime?:string}} [options]
     * @returns {Promise<Array<{pageNumber:number, base64:string, mime:string, width:number, height:number}>>}
     *          任何失败（pdf.js 未加载/文件损坏/渲染异常）都返回 []
     */
    async function renderPdfPages(file, options) {
        var opts = Object.assign({}, DEFAULTS, options || {});
        try {
            if (!file || !ensureWorker()) return [];
            if (typeof file.arrayBuffer !== 'function') return [];

            var buf = await file.arrayBuffer();
            var pdfDoc = await pdfjsLib.getDocument({ data: new Uint8Array(buf) }).promise;
            var out = [];
            var n = Math.min(pdfDoc.numPages || 0, opts.maxPages);

            for (var p = 1; p <= n; p++) {
                if (typeof opts.onProgress === 'function') {
                    try { opts.onProgress({ page: p, total: n }); } catch (_pg) { /* 进度回调失败不影响渲染 */ }
                }
                var page = await pdfDoc.getPage(p);
                var baseVp = page.getViewport({ scale: 1 });
                var longest = Math.max(baseVp.width, baseVp.height);
                var scale = longest > 0 ? Math.min(2.0, opts.maxSidePx / longest) : 1.0;
                var vp = page.getViewport({ scale: scale });

                var canvas = document.createElement('canvas');
                canvas.width = Math.max(1, Math.floor(vp.width));
                canvas.height = Math.max(1, Math.floor(vp.height));
                var ctx = canvas.getContext('2d');
                if (!ctx) { try { page.cleanup(); } catch (_c) {} continue; }
                ctx.fillStyle = '#ffffff';
                ctx.fillRect(0, 0, canvas.width, canvas.height);

                await page.render({ canvasContext: ctx, viewport: vp }).promise;

                var dataUrl = canvas.toDataURL(opts.mime);
                var comma = dataUrl.indexOf(',');
                out.push({
                    pageNumber: p,
                    base64: comma >= 0 ? dataUrl.slice(comma + 1) : dataUrl,
                    mime: opts.mime,
                    width: canvas.width,
                    height: canvas.height
                });
                try { page.cleanup(); } catch (_pc) {}
            }

            try { pdfDoc.destroy(); } catch (_d) {}
            return out;
        } catch (e) {
            console.warn('[pdf-page-render] 渲染失败:', e && e.message);
            return [];
        }
    }

    function preload() {
        return ensureWorker();
    }

    var api = { renderPdfPages, isAvailable, preload, DEFAULTS };
    window.PdfPageRender = api;
    if (window.App && window.App.ChatV2) {
        window.App.ChatV2.pdfPageRender = api;
    }
    console.log('[pdf-page-render] loaded, pdfjsLib available:', isAvailable());
})();
