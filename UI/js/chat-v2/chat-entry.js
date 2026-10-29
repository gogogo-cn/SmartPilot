/**
 * ============================================================
 * chat-v2/chat-entry.js — L1 UI 事件入口
 * ============================================================
 * 职责：
 *   - 接收 UI 事件（发送按钮、回车、文件选择、清除附件、停止、新建对话）
 *   - 从 DOM / App.* 读取用户输入、附件、引用区域
 *   - 构造 Turn 输入对象 → 派发到 Orchestrator
 *
 * 不做的事：
 *   - 不含业务逻辑（不调 LLM、不组装 prompt）
 *   - 不直接渲染回复（由 chat-renderer 做）
 *
 * 暴露的全局函数（供 index.html onclick 调用）：
 *   - window.sendMessage(text?)
 *   - window.handleFileSelect(event)
 *   - window.clearAttachment()
 *   - window.cancelCurrentRequest()
 *   - window.startNewChat()
 *
 * 依赖：App.ChatV2.orchestrator, App.contextEngine, App.attachmentManager
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 扫描版 PDF 的 doc 附件内容前缀标记（fileToText 产出，filesToAttachments 消费：
    // 命中时用 pdf.js 渲染前几页为 image 附件，走 image-pipeline 兜底）
    const SCANNED_PDF_MARKER = '[扫描版PDF]';

    // ============ 降级提示（统一节流） ============

    // 【P2 修复】上下文/附件收集降级不再静默：数据类降级上浮为用户可见的 toast 提示。
    // 8s 节流：同一轮发送中多个降级（如 ContextEngine 总失败 + 选区兜底失败）只提示第一条，
    // 避免连续发送时刷屏；提示失败不影响主流程。
    let _lastCtxWarnAt = 0;
    function notifyContextDegradation(message) {
        try {
            const now = Date.now();
            if (now - _lastCtxWarnAt < 8000) return;
            _lastCtxWarnAt = now;
            App.showToast?.(message, 'warning');
        } catch (e) { /* 提示失败不影响主流程 */ }
    }

    // ============ 宿主感知上下文辅助 ============

    // 安全 JSON 解析（容错：非字符串/非法 JSON 返回 null）
    function _safeJson(raw) {
        if (raw == null) return null;
        if (typeof raw !== 'string') return (raw && typeof raw === 'object') ? raw : null;
        try { return JSON.parse(raw); } catch (_) { return null; }
    }

    // 安全 JSON 解析（容错 COM 斜杠丢失）
    // WebView2 hostObjects 代理会把字符串里的反斜杠吃掉/转成正斜杠：后端 JSON 中
    // Windows 路径的 \\ 变 //、单个 \ 直接消失，残留的 \X（X 不是合法转义符）会让
    // JSON.parse 抛 SyntaxError。先按原文 parse，失败则把非法转义 \X 修成 \\X 再 parse。
    // 注意：合法 JSON 中每个 \ 后面必跟 " \ / b f n r t u，故该修复对合法 JSON 是恒等变换。
    function _safeJsonLoose(raw) {
        if (raw == null) return null;
        if (typeof raw !== 'string') return (raw && typeof raw === 'object') ? raw : null;
        try { return JSON.parse(raw); } catch (_) { /* 继续尝试修复 */ }
        try {
            const repaired = raw.replace(/\\(?!["\\\/bfnrtu])/g, '\\\\');
            return JSON.parse(repaired);
        } catch (_2) { return null; }
    }

    // 桥字符串归一化（去掉可能的多余引号包裹）
    function _asStrSafe(v) {
        if (v == null) return '';
        let s = String(v).trim();
        if (s.startsWith('"') && s.endsWith('"')) { try { s = JSON.parse(s); } catch { /* keep */ } }
        return s;
    }

    // 识别桥返回的 {"success":false,"error":"...not supported / 尚未实现"} 错误串（不应作为文档/工作簿名注入上下文）
    function _isBridgeErrorString(s) {
        return /"success"\s*:\s*false|not supported|尚未实现|当前宿主不支持/i.test(s);
    }

    // 【#3 修复】统一宿主日志落盘封装：把 chat-entry 的关键判定/异常写进宿主日志文件，
    // 避免仅停在 WebView2 console（生产环境不可见）。复用裸桥 App.excelBridge / window.excelBridge 的 LogMessage。
    function _entryHostLog(level, message) {
        try {
            const b = (typeof App !== 'undefined' && App.excelBridge)
                || (typeof window !== 'undefined' && window.excelBridge);
            if (b && typeof b.LogMessage === 'function') {
                const r = b.LogMessage(level || 'INFO', '[chat-entry] ' + message);
                if (r && typeof r.catch === 'function') r.catch(function () { /* 宿主日志失败可忽略 */ });
            }
        } catch (_e) { /* 宿主日志失败绝不影响主流程 */ }
    }

    /** Word 单元格结束符 / 纯空白 → 不算可读正文 */
    function _cleanHostText(t) {
        return String(t || '').replace(/\u0007/g, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
    }

    // 非 Excel 宿主：抓取当前文档/幻灯片文本（供 LLM 总结/润色/翻译）
    // 返回 { text, scope }：scope='current'=仅当前页/选区；'all'=回落到全部幻灯片文本。
    // 【修复】旧实现把全量文本也塞进【当前幻灯片内容】，"总结当前幻灯片"被弱模型
    // 理解成"罗列全部页"（实测 11 页目录被整篇抄出）。必须让 prompt-builder 知道范围。
    // 【修复 2026-10-11】GetSelectedRangeData 旧 DLL 手拼 JSON 缺引号 → parse 失败；
    // 且只取 selectedText、忽略 documentText → 折叠光标时永远空。多路兜底 + 全文回退。
    async function _gatherDocumentText(bridge, host) {
        if (!bridge) return { text: '', scope: '' };
        // 优先：GetSelectedRangeData（PPT/Word 已重写，返回 slideText/selectedText/documentText/host）
        if (typeof bridge.GetSelectedRangeData === 'function') {
            try {
                const data = _safeJson(await bridge.GetSelectedRangeData());
                if (data && data.success) {
                    const sel = _cleanHostText(data.slideText || data.selectedText || data.text || '');
                    const doc = _cleanHostText(data.documentText || '');
                    if (sel) return { text: sel, scope: 'current' };
                    if (doc) return { text: doc, scope: 'current' };
                }
            } catch (ge1) { _entryHostLog('WARN', 'GetSelectedRangeData 失败（走兜底）: ' + (ge1 && ge1.message)); }
        }
        // Word：专用 JSON API（旧 GetSelectedRangeData 破损时仍可用）
        if (host && host.isWord) {
            const tryJsonText = async (fnName, arg) => {
                if (typeof bridge[fnName] !== 'function') return '';
                try {
                    const raw = (arg === undefined) ? await bridge[fnName]() : await bridge[fnName](arg);
                    const data = _safeJson(raw);
                    return _cleanHostText(data && data.text);
                } catch (e) {
                    _entryHostLog('WARN', fnName + ' 失败: ' + (e && e.message));
                    return '';
                }
            };
            const tryPlainText = async (fnName, arg) => {
                if (typeof bridge[fnName] !== 'function') return '';
                try {
                    const raw = (arg === undefined) ? await bridge[fnName]() : await bridge[fnName](arg);
                    // 可能是纯字符串，或误返回 JSON
                    const asObj = _safeJson(raw);
                    if (asObj && typeof asObj.text === 'string') return _cleanHostText(asObj.text);
                    return _cleanHostText(raw);
                } catch (e) {
                    _entryHostLog('WARN', fnName + ' 失败: ' + (e && e.message));
                    return '';
                }
            };
            let t = await tryJsonText('GetSelectionTextJson');
            if (!t) t = await tryPlainText('GetSelectionText');
            if (t) return { text: t, scope: 'current' };
            t = await tryJsonText('GetDocumentTextJson', 20000);
            if (!t) t = await tryPlainText('GetDocumentText', 20000);
            if (t) return { text: t, scope: 'current' };
        }
        // GetSelectionInfo 里常带 selectedText（合法 JSON），旧路径从未采用
        if (typeof bridge.GetSelectionInfo === 'function') {
            try {
                const data = _safeJson(await bridge.GetSelectionInfo());
                const t = _cleanHostText(data && data.selectedText);
                if (t) return { text: t, scope: 'current' };
            } catch (geInfo) { _entryHostLog('WARN', 'GetSelectionInfo 文本兜底失败: ' + (geInfo && geInfo.message)); }
        }
        // PPT 兜底：当前页定位失败时，拉全部幻灯片文本（scope='all'，由提示词明确告知模型）
        if (host && host.isPowerPoint && typeof bridge.GetAllSlidesTextJson === 'function') {
            try {
                const data = _safeJson(await bridge.GetAllSlidesTextJson(4000));
                if (data && data.success && data.text) return { text: String(data.text).trim(), scope: 'all' };
            } catch (ge2b) { _entryHostLog('WARN', 'GetAllSlidesTextJson 失败（走兜底）: ' + (ge2b && ge2b.message)); }
        }
        return { text: '', scope: '' };
    }

    async function _currentSlideIndexLabel(bridge) {
        if (!bridge || typeof bridge.GetCurrentSlideIndex !== 'function') return '';
        try {
            const i = parseInt(_asStrSafe(await bridge.GetCurrentSlideIndex()), 10);
            return (isNaN(i) || i <= 0) ? '' : ('#' + i);
        } catch (_) { _entryHostLog('WARN', 'GetCurrentSlideIndex 失败: ' + (_.message)); return ''; }
    }

    // ============ 文件 → Turn.attachments 转换 ============

    /**
     * 后端文档解析超时（毫秒）。
     * CAD 图纸走 PdfSharp 实测 <1s；MinerU 增强解析最慢 5 分钟。
     * 超过该时间前端必须放弃等待并降级为占位说明——否则 UI 会永久停在
     * "正在准备…"且没有任何反馈（2026-10-18 真实事故：后端 CleanupText 死循环
     * → ParseDocumentAsync 永不返回 → 整条发送链路挂死，用户只能手动点停止）。
     */
    const DOC_PARSE_TIMEOUT_MS = 90000;

    /** 扫描版 PDF 页面渲染超时（worker 卡住时同样不能挂死识别步骤） */
    const PDF_RENDER_TIMEOUT_MS = 45000;

    /**
     * 给 Promise 加超时：超时即 reject。
     * 底层调用照常进行（COM/HTTP 无法取消），但调用方已不再等待，UI 恢复可操作。
     */
    function withCallTimeout(promise, ms, message) {
        return new Promise((resolve, reject) => {
            let done = false;
            const timer = setTimeout(() => {
                if (done) return;
                done = true;
                reject(new Error(message || ('调用超时（>' + Math.round(ms / 1000) + 's）')));
            }, ms);
            Promise.resolve(promise).then(
                v => { if (done) return; done = true; clearTimeout(timer); resolve(v); },
                e => { if (done) return; done = true; clearTimeout(timer); reject(e); }
            );
        });
    }

    function fileToBase64(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = e => {
                const dataUrl = e.target.result || '';
                const commaIdx = dataUrl.indexOf(',');
                resolve(commaIdx >= 0 ? dataUrl.slice(commaIdx + 1) : dataUrl);
            };
            reader.onerror = () => reject(new Error('读取图片失败: ' + file.name));
            reader.readAsDataURL(file);
        });
    }

    async function fileToText(file) {
        const name = (file.name || '').toLowerCase();
        const isExcel = /\.xlsx?$/.test(name) || (file.type || '').includes('excel') || (file.type || '').includes('spreadsheet');
        const isPdf = /\.pdf$/.test(name) || (file.type || '').includes('pdf');
        const isWord = /\.docx?$/.test(name) || (file.type || '').includes('word');
        const isText = /\.(txt|md|json|csv|xml|js|py|html|css|log)$/.test(name) || (file.type || '').startsWith('text/');

        // Excel/PDF/Word：优先走后端解析（DocumentRouter：MinerU/PdfSharp/DocxParser），
        // 避免把二进制乱码喂给模型。CAD 图纸 PDF 有文本层时可提取尺寸/标注/标题栏文字。
        if (isExcel || isPdf || isWord) {
            try {
                const base64 = await fileToBase64(file);
                const pureBase64 = base64.includes(',') ? base64.split(',').pop() : base64;
                const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || window.excelBridge);
                if (bridge && typeof bridge.ParseDocumentAsync === 'function') {
                    // 超时保护：后端解析死循环/极慢时，前端降级而不是无限期挂起"正在准备…"
                    const resultJson = await withCallTimeout(
                        bridge.ParseDocumentAsync(file.name || 'uploaded', pureBase64),
                        DOC_PARSE_TIMEOUT_MS,
                        `文档解析超时（>${Math.round(DOC_PARSE_TIMEOUT_MS / 1000)}s），已放弃等待后端`
                    );
                    // COM 斜杠丢失容错：先原文 parse，失败再修复非法转义
                    const resultObj = _safeJsonLoose(resultJson);
                    if (resultObj && resultObj.success) {
                        const chars = resultObj.totalChars || (resultObj.fullText ? resultObj.fullText.length : 0);
                        if (chars > 0) {
                            return formatParsedDocument(resultObj, file.name);
                        }
                        // 后端成功但无文本（扫描版 PDF / 空文档）：明确告知，不静默丢内容
                        const pages = resultObj.totalPages || 0;
                        if (isPdf) {
                            _entryHostLog('INFO', `PDF ${file.name} 无文本层（${pages}页），判定扫描版`);
                            // 先不提示"只能看到基本信息"：页面图片还没渲染，成功时这句话是假的。
                            // 结果由 filesToAttachments 在渲染结束后补上，并据此决定是否 toast。
                            return `${SCANNED_PDF_MARKER} ${file.name}\n页数: ${pages}\n注：该 PDF 没有可提取的文本层（扫描件或纯图形）。`;
                        }
                        _entryHostLog('INFO', `文档 ${file.name} 解析成功但无文本内容`);
                        return formatParsedDocument(resultObj, file.name);
                    }
                    _entryHostLog('WARN', `后端解析未成功 ${file.name}: ${(resultObj && resultObj.errorMessage) || 'unknown'}`);
                } else {
                    _entryHostLog('WARN', `宿主不支持 ParseDocumentAsync，${file.name} 走占位说明`);
                }
            } catch (e) {
                ChatLog.warn?.('CHAT-V2', `[chat-entry] 文档后端解析失败 ${file.name}`, e);
                _entryHostLog('WARN', `文档后端解析异常 ${file.name}: ${e && e.message}`);
                notifyContextDegradation(`文件 ${file.name} 解析失败，仅向 AI 提供文件基本信息`);
                // 超时/异常时给准确占位（不要落到"不支持此格式"的误导文案——格式是支持的，只是这次没解析出来）
                return `[附件] ${file.name}\n类型: ${file.type || 'unknown'}\n大小: ${formatSize(file.size)}\n注：后端解析失败（${(e && e.message) || '未知错误'}），已仅向 AI 提供文件基本信息。`;
            }
        }

        // 纯文本：直接读取
        if (isText) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = e => resolve(e.target.result || '');
                reader.onerror = () => reject(new Error('读取文件失败: ' + file.name));
                reader.readAsText(file);
            });
        }

        // 其他二进制文件（PPT/图片外的未知格式）：不读内容，避免模型收到乱码后误判文件损坏
        return `[附件] ${file.name}\n类型: ${file.type || 'unknown'}\n大小: ${formatSize(file.size)}\n注：当前不支持直接解析此格式内容，如需分析请转换为文本/表格后上传。`;
    }

    /**
     * 把后端 ParseDocumentAsync 返回的结果格式化为模型可读的文本
     */
    function formatParsedDocument(resultObj, fileName) {
        const lines = [];
        const lower = (fileName || '').toLowerCase();
        let kindLabel = '[文档]';
        if (/\.xlsx?$/.test(lower)) kindLabel = '[Excel文件]';
        else if (/\.pdf$/.test(lower)) kindLabel = '[PDF文档]';
        else if (/\.docx?$/.test(lower)) kindLabel = '[Word文档]';
        lines.push(`${kindLabel} ${fileName}`);
        lines.push(`工作表/页数: ${resultObj.totalPages || 0}`);
        lines.push(`总字符数: ${resultObj.totalChars || 0}`);
        lines.push('');

        const pages = resultObj.pages || [];
        if (pages.length > 0) {
            pages.forEach((page, idx) => {
                const title = page.title || page.Title || `Sheet${idx + 1}`;
                const text = page.text || page.Text || '';
                lines.push(`--- ${title} ---`);
                lines.push(text);
                lines.push('');
            });
        } else if (resultObj.fullText) {
            lines.push(resultObj.fullText);
        }

        let content = lines.join('\n');
        if (content.length > 12000) {
            content = content.substring(0, 12000) + '\n...（内容已截断，如需完整数据请分批询问）';
        }
        return content;
    }

    function isImageFile(file) {
        if (App.attachmentManager?.isImageFile) return App.attachmentManager.isImageFile(file);
        return /^image\//.test(file?.type || '');
    }

    /**
     * 把 File[] 转成 Turn 用的 attachments 结构
     * @returns {Promise<Array>} attachments
     */
    async function filesToAttachments(files) {
        const list = [];
        for (const file of files || []) {
            if (file) list.push(file);
        }
        const result = [];
        const total = list.length;
        for (let i = 0; i < list.length; i++) {
            const file = list[i];
            const nth = total > 1 ? `（${i + 1}/${total}）` : '';
            try {
                if (isImageFile(file)) {
                    reportPrep(`正在读取图片${nth}：${file.name}`);
                    const base64 = await fileToBase64(file);
                    result.push({
                        type: 'image',
                        mime: file.type || 'image/png',
                        fileName: file.name,
                        size: file.size,
                        base64
                    });
                } else {
                    reportPrep(`正在解析附件${nth}：${file.name}`);
                    const text = await fileToText(file);
                    const docAtt = {
                        type: 'doc',
                        mime: file.type || 'application/octet-stream',
                        fileName: file.name,
                        size: file.size,
                        content: text
                    };
                    result.push(docAtt);
                    // 扫描版 PDF 不在发送时渲染。文件留在附件上，识别页面这一步再渲前 3 页。
                    // 有文字层的 PDF 不会走到这里，因此不付渲染时间。
                    if (typeof text === 'string' && text.startsWith(SCANNED_PDF_MARKER)) {
                        docAtt.sourceFile = file;
                        docAtt.content += '\n补充：正文没有文字层。识别页面时会渲染前 3 页，不要根据文件名编造尺寸。';
                        _entryHostLog('INFO', `扫描版 PDF ${file.name} 留到识别页面时再渲染`);
                    }
                }
            } catch (err) {
                ChatLog.warn?.('CHAT-V2', `[chat-entry] 转换附件失败 ${file.name}`, err);
                // 【P2 修复】附件转换失败=该文件被静默丢弃，用户必须知道
                notifyContextDegradation(`附件 ${file.name} 处理失败，已跳过（未发送给 AI）`);
            }
        }
        return result;
    }

    // ============ 上下文收集 ============

    /**
     * 收集 turn 上下文（引用区域等）
     * 委托给 App.contextEngine.gather（已封装 ReferenceAreaManager / excelBridge 等）
     *
     * @param {string} message - 用户消息文本，用于上下文引擎判断是否需要自动捕获工作表
     */
    async function gatherContext(message, opts) {
        reportPrep((opts && opts.hasFiles)
            ? '附件已准备好，正在读取工作表上下文…'
            : '正在读取当前表格…');
        const ctx = { referenceRegions: [] };

        // 1. 优先使用 ContextEngine（注意大小写：App.ContextEngine 而非 contextEngine）
        let contextEngine = App.ContextEngine || App.contextEngine;
        try {
            if (contextEngine?.gather) {
                const gathered = await contextEngine.gather({ message: message || '' });
                ctx.referenceRegions = gathered.referenceRegions || [];
                if (gathered.selection) ctx.selection = gathered.selection;
                // 【P2 修复】数据类降级（选区/引用区域读取失败）上浮为用户可见提示，不再静默。
                // 历史/附件类降级影响小，保留日志即可，避免打扰。
                const dataDegradations = (gathered._degradations || [])
                    .filter(d => d.block === 'selection' || d.block === 'referenceRegions');
                if (dataDegradations.length > 0) {
                    notifyContextDegradation('表格数据读取异常：' + dataDegradations[0].message + '，AI 可能无法看到最新数据');
                }
            }
        } catch (e) {
            ChatLog.warn?.('CHAT-V2', '[chat-entry] gather context failed', e);
            notifyContextDegradation('上下文收集失败：AI 可能无法获取选区与引用区域数据');
        }

        // 2. 兜底：直接通过桥接层获取当前选中/幻灯片/文档数据
        // 【修复】非 Excel 宿主下 ContextEngine 可能返回 selection（且它不会被注入提示词——
        // prompt-builder 非 Excel 分支只注入 documentText），旧条件 !ctx.selection 会因此
        // 整体跳过本兜底 → documentText 依旧为空 → 模型无米下锅凭标题编造假正文。
        // 改为：Excel 仍按 !ctx.selection；非 Excel 只要 documentText 未取到就执行兜底。
        const _hostEarly = App.SmartPilotHost || window.SmartPilotHost;
        const _isNonExcelHost = !!(_hostEarly && !_hostEarly.isExcel);
        if (!ctx.selection || (_isNonExcelHost && !ctx.documentText)) {
            try {
                const bridge = App.excelBridge;
                if (bridge && typeof bridge.GetSelectedRangeData === 'function') {
                    const raw = await bridge.GetSelectedRangeData();
                    if (raw) {
                        const data = _safeJson(raw);
                        if (data && data.success && Array.isArray(data.headers) && data.headers.length > 0) {
                            // Excel：标准选区（含 headers/rows）
                            const sel = {
                                address: data.address || '',
                                workbookName: data.workbookName || '',
                                sheetName: data.sheetName || '',
                                range: data.address || '',
                                headers: data.headers || [],
                                rows: data.rows || [],
                                rowCount: data.rowCount || 0,
                                colCount: data.colCount || 0
                            };
                            if (sel.headers.length > 1 || sel.rows.length > 0) {
                                ctx.selection = sel;
                            }
                        } else if (data && data.success) {
                            // PowerPoint / Word：优先选区/当前页，空则回退全文（旧逻辑忽略 documentText）
                            const docText = _cleanHostText(
                                data.slideText || data.selectedText || data.text || data.documentText || '');
                            if (docText) ctx.documentText = docText;
                        }
                    }
                }
            } catch (e) {
                ChatLog.warn?.('CHAT-V2', '[chat-entry] 直接获取选区数据失败', e);
                // 【P2 修复】ContextEngine 与桥接兜底双双失败，选区数据彻底不可用，必须告知用户
                notifyContextDegradation('选区数据读取失败：AI 将无法获取当前选中区域的内容');
            }
        }

        // 3. 活动上下文（按宿主注入 LLM 上下文）
        try {
            const bridge = App.excelBridge;
            // 【修复】宿主对象挂在 window.SmartPilotHost，全仓未给 App.SmartPilotHost 赋值，
            // 若只取 App.SmartPilotHost 会恒为 undefined，导致非 Excel 分支（PPT/Word 活动上下文采集）
            // 被整体跳过 → documentText 永不注入 → 弱模型无米下锅编出"项目A…"等假内容。
            const host = App.SmartPilotHost || window.SmartPilotHost;

            if (host && host.isExcel) {
                // ===== Excel：活动工作表名 / 工作簿名 / 工作表清单 =====
                // 这样 AI 调用 excel_copy_sheet(source)/excel_delete_sheet(name)/excel_rename_sheet(oldName)
                // 等工具时能直接用真实表名填参，而不是写出"当前工作表名称"这类占位符。
                const asStr = (v) => {
                    if (v == null) return '';
                    const s = String(v).trim();
                    // 桥方法可能返回 JSON 字符串（如 "\"Sheet1\""）或纯字符串，统一归一化
                    if (s.startsWith('"') && s.endsWith('"')) { try { return JSON.parse(s); } catch { /* keep */ } }
                    return s;
                };
                let sheet = (typeof bridge?.GetActiveSheetName === 'function') ? asStr(await bridge.GetActiveSheetName()) : '';
                let workbook = (typeof bridge?.GetActiveWorkbookName === 'function') ? asStr(await bridge.GetActiveWorkbookName()) : '';
                if (!sheet && ctx.selection?.sheetName) sheet = String(ctx.selection.sheetName).trim();
                if (!workbook && ctx.selection?.workbookName) workbook = String(ctx.selection.workbookName).trim();
                if (sheet || workbook) ctx.activeSheet = { sheet, workbook };

                // 【真实工作表清单】顺手取一次全部可见工作表名，
                // 供 prompt-builder 注入系统提示——从根上避免模型编造表名（如"销售数据"）。
                let getSheetsFn = null;
                try { getSheetsFn = bridge?.GetSheets; } catch { getSheetsFn = null; }
                if (typeof getSheetsFn === 'function') {
                    try {
                        const raw = await bridge.GetSheets();
                        let arr = [];
                        try { arr = JSON.parse(raw); } catch { arr = []; }
                        if (Array.isArray(arr) && arr.length) ctx.workbookSheets = arr;
                        const summary = (Array.isArray(arr) && arr.length) ? arr.join(' | ') : '(空——未注入)';
                        ChatLog.info?.('CHAT-V2', `[chat-entry] 工作表清单(${Array.isArray(arr) ? arr.length : 0}): ${summary}`);
                        try { bridge.LogMessage?.('INFO', `[chat-entry] 工作表清单(${Array.isArray(arr) ? arr.length : 0}): ${summary}`); } catch (_e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'bridge.LogMessage 失败:', _e && _e.message); }
                    } catch (se) {
                        ChatLog.warn?.('CHAT-V2', '[chat-entry] 获取工作表清单失败（未注入）:', se?.message || se);
                        try { bridge.LogMessage?.('WARN', `[chat-entry] 获取工作表清单失败（未注入）: ${se?.message || se}`); } catch (_e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'bridge.LogMessage 失败:', _e && _e.message); }
                    }
                } else {
                    ChatLog.warn?.('CHAT-V2', '[chat-entry] 桥未提供 GetSheets（DLL 可能未更新到含该方法版本），跳过工作表清单注入');
                    try { bridge.LogMessage?.('WARN', '[chat-entry] 桥未提供 GetSheets（DLL 可能未更新），跳过工作表清单注入'); } catch (_e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'bridge.LogMessage 失败:', _e && _e.message); }
                }

                // 【用户选中的多个工作表标签】用户在 Excel 标签栏按 Ctrl/Shift 多选了几张表 → 注入这个"选中态"，
                // 让模型知道"合并/拆分操作的作用对象就是这几张表"，而不是只读活动工作表自己做分析。
                let getSelSheetsFn = null;
                try { getSelSheetsFn = bridge?.GetSelectedSheetNames; } catch { getSelSheetsFn = null; }
                if (typeof getSelSheetsFn === 'function') {
                    try {
                        const rawSel = await bridge.GetSelectedSheetNames();
                        let selArr = [];
                        try { selArr = JSON.parse(rawSel); } catch { selArr = []; }
                        if (Array.isArray(selArr) && selArr.length > 1) {
                            ctx.selectedSheets = selArr; // 只有多选（>1）才注入，单选=活动工作表，前端 prompt-builder 有 else 兜底
                            ChatLog.info?.('CHAT-V2', `[chat-entry] 用户选中工作表(${selArr.length}): ${selArr.join('、')}`);
                            try { bridge.LogMessage?.('INFO', `[chat-entry] 用户选中工作表(${selArr.length}): ${selArr.join('、')}`); } catch (_e) { /* ignore */ }
                        }
                    } catch (_e) { /* bridge 方法异常时静默降级，不阻塞主流程 */ }
                }
            } else if (host && (host.isWord || host.isPowerPoint)) {
                // ===== Word / PowerPoint：捕获当前文档/幻灯片文本，避免把 Excel 专属错误串当名称污染上下文 =====
                const _doc = await _gatherDocumentText(bridge, host);
                if (_doc && _doc.text) {
                    ctx.documentText = _doc.text;
                    ctx.documentScope = _doc.scope; // 'current' | 'all'（prompt-builder 据此改标签与指令）
                }

                // 文档/演示文稿名（替代 Excel 的 workbook 名；桥在非 Excel 宿主下 GetActiveWorkbookName
                // 会返回 {"success":false,"error":"operation not supported..."}，必须过滤，否则污染 LLM 上下文）
                let docName = '';
                if (typeof bridge?.GetPresentationInfo === 'function') {
                    try {
                        const info = _safeJson(await bridge.GetPresentationInfo());
                        if (info && info.success && info.name) docName = String(info.name);
                    } catch (_) { _entryHostLog('WARN', 'GetPresentationInfo 失败（忽略）: ' + (_.message)); }
                }
                if (!docName && typeof bridge?.GetActiveWorkbookName === 'function') {
                    const n = _asStrSafe(await bridge.GetActiveWorkbookName());
                    if (n && !_isBridgeErrorString(n)) docName = n;
                }
                const slideTag = host.isPowerPoint ? await _currentSlideIndexLabel(bridge) : '';
                if (docName || slideTag) {
                    ctx.activeSheet = { sheet: slideTag, workbook: docName };
                }

                // 选中对象采集（表格/形状）：让 AI 知道用户选中的是具体对象，而非“无选区”。
                // 此前 PPT/Word 只采选中文本、从不报选中形状/表格，导致意图分析判“未指定操作对象”。
                if (typeof bridge.GetSelectionInfo === 'function') {
                    try {
                        const sraw = await bridge.GetSelectionInfo();
                        const sinfo = _safeJson(sraw);
                        if (sinfo && sinfo.success && sinfo.selectionType && sinfo.selectionType !== 'none') {
                            ctx.selectedObject = sinfo;
                            ctx.selection = ctx.selection || {};
                            ctx.selection.objectType = sinfo.selectionType;
                            ctx.selection.objectSummary = sinfo.summary || '';
                            if (sinfo.selectionType === 'table') {
                                ctx.selection.slideIndex = sinfo.slideIndex || 0;
                                ctx.selection.shapeIndex = (sinfo.shapeIndex != null ? sinfo.shapeIndex : 0);
                                ctx.selection.tableRowCount = sinfo.rowCount || 0;
                                ctx.selection.tableColCount = sinfo.colCount || 0;
                                ctx.selection.address = sinfo.summary; // 让下游识别为有目标选区
                            }
                            // 【修复】GetSelectionInfo 已带 selectedText，但旧逻辑从不写入 documentText
                            // → 翻译/润色时 prompt 显示「没有可读取的文字」（实测 Word 01:08）。
                            const selTxt = _cleanHostText(sinfo.selectedText);
                            if (selTxt) {
                                ctx.selectionText = selTxt;
                                if (!ctx.documentText) ctx.documentText = selTxt;
                            }
                            ChatLog.info?.('CHAT-V2', `[chat-entry] 选中对象(${sinfo.selectionType}): ${sinfo.summary}`);
                        }
                    } catch (_e) {
                        ChatLog.warn?.('CHAT-V2', '[chat-entry] GetSelectionInfo 失败', _e);
                        _entryHostLog('WARN', 'GetSelectionInfo 失败: ' + (_e && _e.message));
                    }
                }
                ChatLog.info?.('CHAT-V2', `[chat-entry] 非Excel上下文(${host.kind}): docText=${(ctx.documentText || '').length}字, doc=${docName || '(空)'}`);
                _entryHostLog('INFO', `非Excel上下文(${host.kind}): docText=${(ctx.documentText || '').length}字`);
            }
        } catch (e) {
            ChatLog.warn?.('CHAT-V2', '[chat-entry] 捕获活动上下文失败', e);
        }

        return ctx;
    }

    // ============ 附件收集 ============

    function collectFiles() {
        // 1. attachmentManager.recoverFiles（最权威）
        if (App.attachmentManager?.recoverFiles) {
            const recovered = App.attachmentManager.recoverFiles();
            if (recovered && recovered.length > 0) return recovered;
        }
        // 2. App.selectedFiles / App.selectedFile
        if (Array.isArray(App.selectedFiles) && App.selectedFiles.length > 0) {
            return [...App.selectedFiles];
        }
        if (App.selectedFile) return [App.selectedFile];
        // 3. 兜底：直接读 file-input DOM
        const fi = document.getElementById('file-input');
        if (fi && fi.files && fi.files.length > 0) {
            const out = [];
            for (let i = 0; i < fi.files.length; i++) out.push(fi.files[i]);
            return out;
        }
        return [];
    }

    // ============ 编辑重发：恢复已保存的附件 ============

    /**
     * 编辑模式下，从 send-btn 的 dataset 中恢复上一轮的附件。
     * 图片附件带有 base64 数据，可直接恢复；文档附件仅保留元数据，需提示用户重新上传。
     * @returns {Promise<Array>} 可恢复的 attachments 数组
     */
    async function recoverEditingAttachments() {
        const sendBtn = document.getElementById('send-btn');
        if (!sendBtn || !sendBtn.dataset.editingTurnId) return [];

        const attachments = [];

        // 1. 恢复图片附件（有 base64 数据）
        const imageJson = sendBtn.dataset.editingImageBase64;
        if (imageJson) {
            try {
                const images = JSON.parse(imageJson);
                for (const img of images) {
                    attachments.push({
                        type: 'image',
                        mime: img.mime || 'image/png',
                        fileName: img.fileName || 'image.png',
                        base64: img.base64
                    });
                }
            } catch (e) {
                ChatLog.warn?.('CHAT-V2', '[recoverEditingAttachments] 解析图片附件失败', e);
            }
        }

        // 2. 恢复文档附件元数据（仅展示用，实际文件需重新上传）
        const attachJson = sendBtn.dataset.editingAttachments;
        if (attachJson) {
            try {
                const meta = JSON.parse(attachJson);
                for (const m of meta) {
                    if (m.type !== 'image') {
                        attachments.push({
                            type: 'doc',
                            mime: m.mime || 'application/octet-stream',
                            fileName: m.fileName || 'unknown',
                            size: m.size || 0,
                            content: `[附件] ${m.fileName}\n类型: ${m.mime || 'unknown'}\n注：编辑模式下附件元数据已保留，如需重新上传请使用文件上传按钮。`
                        });
                    }
                }
            } catch (e) {
                ChatLog.warn?.('CHAT-V2', '[recoverEditingAttachments] 解析文档附件元数据失败', e);
            }
        }

        return attachments;
    }

    /**
     * 清除编辑模式标记
     */
    function clearEditingMode() {
        const sendBtn = document.getElementById('send-btn');
        if (sendBtn) {
            delete sendBtn.dataset.editingTurnId;
            delete sendBtn.dataset.editingAttachments;
            delete sendBtn.dataset.editingImageBase64;
        }
        // 清除 preview._storedFiles（如果有编辑模式留下的）
        const preview = document.getElementById('attachment-preview');
        if (preview?._editingFiles) {
            delete preview._editingFiles;
        }
    }

    // ============ 主入口：sendMessage ============

    async function sendMessage() {
        const orchestrator = App.ChatV2.orchestrator;
        if (!orchestrator) {
            console.error('[chat-entry] orchestrator not loaded');
            App.showToast?.('系统初始化中，请稍候…', 'info');
            return;
        }

        // 1. 读取用户输入
        const textarea = document.getElementById('message-input');
        const text = (textarea?.value || '').trim();
        const files = collectFiles();

        // 检查是否处于编辑重发模式
        const sendBtn = document.getElementById('send-btn');
        const isEditing = sendBtn && sendBtn.dataset.editingTurnId;

        if (!text && files.length === 0 && !isEditing) {
            App.showToast?.('请输入消息或上传文件', 'info');
            return;
        }

        // 1.5 未配置模型/API Key：软拦截并引导去设置（避免直接报 HTTP/网络错）
        if (typeof App.ensureLlmConfiguredForSend === 'function') {
            if (!App.ensureLlmConfiguredForSend()) return;
        } else if (typeof App.isLlmConfigured === 'function' && !App.isLlmConfigured()) {
            App.showToast?.('请先在设置中配置 AI 模型或 API 密钥', 'warning');
            try { if (typeof window.openAiSettings === 'function') window.openAiSettings(); } catch (_) { /* ignore */ }
            return;
        }

        // 2. 若正在执行，先取消（新消息接管）。
        // 取消会同步触发上一轮 turn_done → resetProcessingUI，所以必须发生在下面
        // 把按钮切成「停止」之前，否则停止态会被立刻盖回「发送」。
        if (orchestrator.isRunning()) {
            orchestrator.cancelCurrent('replaced_by_new_turn');
        }

        // 3. 点击后立刻切按钮、清空输入框，再去做席位校验和读表。
        // 席位门禁走宿主桥，桥里是一次最长 20 秒的授权 HTTP。以前这步排在改按钮之前，
        // 授权服务慢或连不上时，按钮会一直停在「发送」，像是没点上。
        setProcessingUI(true);
        let draftText = '';
        if (textarea) {
            draftText = textarea.value;
            textarea.value = '';
            textarea.style.height = 'auto';
        }
        const restoreDraft = () => {
            try {
                if (textarea && draftText && !textarea.value) {
                    textarea.value = draftText;
                    textarea.style.height = 'auto';
                    textarea.style.height = Math.min(textarea.scrollHeight, 200) + 'px';
                }
            } catch (_) { /* 回填失败忽略 */ }
        };
        // 让出一帧，把「停止」画上去，再发起会占住 Excel 界面线程的桥调用。
        await new Promise(resolve => setTimeout(resolve, 0));

        // 首次发送前确认：云端模型会把输入和选区发到用户填写的 API 地址
        if (typeof window.ensureDataDisclosureAccepted === 'function') {
            const accepted = await window.ensureDataDisclosureAccepted();
            if (!accepted) {
                restoreDraft();
                setProcessingUI(false);
                return;
            }
        }

        // 席位订阅门禁（BYOK：不计量 Token，只校验试用/订阅是否有效）
        if (typeof window.ensureSubscriptionActive === 'function') {
            const subOk = await window.ensureSubscriptionActive();
            if (!subOk) {
                restoreDraft();
                setProcessingUI(false);
                return;
            }
        }

        try {
            if (typeof window.showErrorBanner === 'function') window.showErrorBanner('');
        } catch (_) { /* ignore */ }

        try {
            // 4. 转换附件（正常文件 或 编辑模式恢复的附件）
            let attachments = files.length > 0 ? await filesToAttachments(files) : [];

            // 编辑模式：如果用户没有重新选择文件，从存储中恢复上一轮的附件
            if (isEditing && attachments.length === 0) {
                const recovered = await recoverEditingAttachments();
                if (recovered.length > 0) {
                    attachments = recovered;
                    ChatLog.info?.('CHAT-V2', `[sendMessage] 编辑重发恢复 ${recovered.length} 个附件`);
                }
            }

            // 5. 清除编辑模式标记（必须在 dispatch 之前，因为 dispatch 会触发 turn_start 事件）
            if (isEditing) {
                clearEditingMode();
            }

            // 6. 收集上下文（传入用户消息，用于自动判断是否需要捕获当前工作表）
            const context = await gatherContext(text, { hasFiles: attachments.length > 0 });

            // 6.5 快捷入口直连：把技能身份带进 context，由 orchestrator 钉死该技能。
            // 用户在快捷入口里显式选了技能，这条消息就不该再被意图路由改判到别的技能
            // （详见 app-history.js 的"快捷技能绑定"说明与 orchestrator 的【快捷入口直连钉死】）。
            const _qsBind = App._quickSkillBinding;
            if (_qsBind && _qsBind.name && (Date.now() - _qsBind.ts) < 30 * 60 * 1000) {
                context._forcedSkillName = _qsBind.name;
                context._forcedSkillSource = 'quickShortcut';
                ChatLog.info?.('CHAT-V2', `[sendMessage] 快捷入口技能绑定 → context._forcedSkillName=${_qsBind.name}`);
            }

            // 7. 派发到 orchestrator（异步执行，立即返回 turnId）
            const turnId = await orchestrator.dispatch({
                message: text,
                attachments,
                context
            });

            ChatLog.info?.('CHAT-V2', `[chat-entry] dispatched turn=${turnId}`);

            // 8. 清空附件预览（文本已在 3.5 步即时清空；用户消息由 renderer 在 turn_start 事件中渲染）
            clearAttachment();

            // 8.5 绑定只对本次发送有效：发出即解绑，不影响后续普通对话（用户也可用 × 提前取消）
            if (_qsBind) App.clearQuickSkillBinding?.();
        } catch (err) {
            // 发送失败：回填草稿，避免"失败还把用户写的内容清没了"
            restoreDraft();
            ChatLog.error?.('CHAT-V2', '[chat-entry] dispatch failed', err);
            if (window.ErrorHandler && typeof ErrorHandler.showError === 'function') {
                ErrorHandler.showError(err, '发送消息', 'error');
            } else {
                App.showToast?.('发送失败: ' + (err?.message || err), 'error');
            }
            setProcessingUI(false);
        }
    }

    // ============ 取消当前请求 ============

    function cancelCurrentRequest() {
        const orchestrator = App.ChatV2.orchestrator;
        if (!orchestrator) return;
        const canceled = orchestrator.cancelCurrent('user_aborted');
        if (!canceled) {
            App.showToast?.('当前没有正在进行的请求', 'info');
        } else {
            setAgentPhaseStatus('已停止', { canceled: true });
            App.showToast?.('已停止当前任务', 'info');
        }
    }

    // ============ 文件选择处理 ============

    async function handleFileSelect(event) {
        const input = event?.target;
        const files = input?.files ? Array.from(input.files) : [];
        if (files.length === 0) return;

        // 同步到 App.selectedFiles（旧系统兼容）
        App.selectedFiles = files;
        App.selectedFile = files[0];

        // 在 DOM 上也存一份（attachment-preview._storedFiles）
        const preview = document.getElementById('attachment-preview');
        if (preview) preview._storedFiles = files;

        renderAttachmentPreview(files);

        if (files.some(f => /\.pdf$/i.test(f && f.name || ''))) {
            const render = (App.ChatV2 && App.ChatV2.pdfPageRender) || window.PdfPageRender;
            if (render && typeof render.preload === 'function') render.preload();
        }

        ChatLog.info?.('CHAT-V2', `[chat-entry] selected ${files.length} file(s): ${files.map(f => f.name).join(', ')}`);
    }

    // ============ 清除附件 ============

    function clearAttachment() {
        App.selectedFiles = [];
        App.selectedFile = null;

        const fi = document.getElementById('file-input');
        if (fi) fi.value = '';

        const preview = document.getElementById('attachment-preview');
        if (preview) {
            preview.style.display = 'none';
            preview._storedFiles = null;
            const thumbnail = document.getElementById('attachment-thumbnail');
            const thumbnailImg = document.getElementById('thumbnail-img');
            if (thumbnailImg && thumbnailImg.src) {
                try { URL.revokeObjectURL(thumbnailImg.src); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'revokeObjectURL 失败:', e && e.message); }
                thumbnailImg.src = '';
            }
            if (thumbnail) thumbnail.style.display = 'none';
            const header = preview.querySelector('.attachment-icon-only');
            if (header) header.style.display = 'none';
            const status = document.getElementById('attachment-analysis-status');
            if (status) status.style.display = 'none';
            const list = document.getElementById('attachment-list');
            if (list) list.innerHTML = '';
        }

        // 通知 attachmentManager 清理
        try { App.attachmentManager?.clear?.(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'attachmentManager.clear 失败:', e && e.message); }
    }

    // ============ 附件预览渲染 ============

    function renderAttachmentPreview(files) {
        if (!files || files.length === 0) return;
        const first = files[0];

        const preview = document.getElementById('attachment-preview');
        if (!preview) return;
        preview.style.display = 'flex';

        // 隐藏旧版单文件 header
        const header = preview.querySelector('.attachment-icon-only');
        if (header) header.style.display = 'none';

        // 使用统一附件列表渲染（支持单/多文件和独立删除）
        if (typeof App.renderAttachmentList === 'function') {
            App.renderAttachmentList(files);
        }

        const thumb = document.getElementById('attachment-thumbnail');
        if (files.length === 1 && isImageFile(first)) {
            const img = document.getElementById('thumbnail-img');
            if (thumb && img) {
                try { URL.revokeObjectURL(img.src); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'revokeObjectURL 失败:', e && e.message); }
                img.src = URL.createObjectURL(first);
                thumb.style.display = 'block';
            }
        } else {
            if (thumb) thumb.style.display = 'none';
        }
    }

    function fileIcon(file) {
        const name = (file.name || '').toLowerCase();
        if (/\.xlsx?$/.test(name)) return '📊';
        if (/\.pdf$/.test(name)) return '📕';
        if (/\.docx?$/.test(name)) return '📝';
        if (/\.txt$/.test(name)) return '📄';
        return '📎';
    }

    function extOf(name) {
        const i = (name || '').lastIndexOf('.');
        return i >= 0 ? name.slice(i + 1).toUpperCase() : '文件';
    }

    function formatSize(bytes) {
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.formatSize(bytes) : '0 B';
    }

    // ============ 发送按钮状态切换 ============

    /** 通俗阶段文案（给新手看的进度，不是调试日志） */
    function phaseT(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
        } catch (_e) { /* ignore */ }
        return key;
    }

    function phaseBaseLabel(phase) {
        const map = {
            turn: 'phase.turn',
            step: 'phase.step',
            thinking: 'phase.thinking',
            connecting: 'phase.connecting',
            tool: 'phase.tool',
            summarizing: 'phase.summarizing',
            done: 'phase.done',
            canceled: 'phase.canceled'
        };
        return map[phase] ? phaseT(map[phase]) : '';
    }

    /**
     * 发送前的附件解析 / 读表还没有 turn，气泡不存在。
     * 这段时间只写输入区进度条，避免整段停在「正在准备…」。
     * explicit：随后 1.5s 内，笼统的 turn/step 阶段不要把这句盖掉。
     */
    function reportPrep(text) {
        setAgentPhaseStatus(text, { explicit: true });
    }

    function setAgentPhaseStatus(text, opts) {
        const el = document.getElementById('agent-phase-status');
        if (!el) return;
        const msg = (text == null ? '' : String(text)).trim();
        if (!msg) {
            setAgentPhaseStatus._explicitAt = 0;
            el.style.display = 'none';
            el.textContent = '';
            el.classList.remove('is-active', 'is-done', 'is-canceled', 'is-tool', 'is-thinking');
            return;
        }
        if (opts && opts.explicit) setAgentPhaseStatus._explicitAt = Date.now();
        el.textContent = msg;
        el.style.display = 'block';
        el.classList.add('is-active');
        el.classList.toggle('is-done', !!(opts && opts.done));
        el.classList.toggle('is-canceled', !!(opts && opts.canceled));
        el.classList.toggle('is-tool', !!(opts && opts.tool));
        el.classList.toggle('is-thinking', !!(opts && opts.thinking));
        if (opts && (opts.done || opts.canceled)) {
            el.classList.remove('is-active', 'is-tool', 'is-thinking');
            // 短暂保留结果文案后清空
            clearTimeout(setAgentPhaseStatus._hideTimer);
            setAgentPhaseStatus._hideTimer = setTimeout(function () {
                if (el.textContent === msg) setAgentPhaseStatus('');
            }, 1800);
        }
    }

    function phaseLabelFromEvent(payload) {
        const phase = (payload && payload.phase) || '';
        const step = payload && payload.step;
        const stepTotal = payload && payload.stepTotal;
        let label = '';
        if (phase === 'tool') {
            const name = payload.displayToolName || payload.detail || payload.toolName || '';
            const raw = payload.toolName || '';
            if (name) label = phaseT('phase.runningTool', { name: name });
            else if (/write|insert|fill|format|clear|delete|merge/i.test(raw)) label = phaseT('phase.toolWrite');
            else if (/export|save/i.test(raw)) label = phaseT('phase.toolExport');
            else if (/read|get_|list_/i.test(raw)) label = phaseT('phase.toolRead');
            else label = phaseT('phase.tool');
        } else {
            label = phaseBaseLabel(phase) || (payload && payload.status === 'running' ? phaseT('phase.processing') : '');
        }
        if (label && step && stepTotal && stepTotal > 1 && (phase === 'thinking' || phase === 'connecting' || phase === 'summarizing' || phase === 'tool' || phase === 'step')) {
            label = phaseT('phase.stepOf', { n: step, total: stepTotal }) + ' · ' + label;
        }
        return label;
    }

    function bindAgentPhaseStatus() {
        const ae = App.ChatV2 && App.ChatV2.agentEvents;
        if (!ae || typeof ae.on !== 'function' || bindAgentPhaseStatus._bound) return;
        bindAgentPhaseStatus._bound = true;
        ae.on('agent/status', function (payload) {
            const label = phaseLabelFromEvent(payload || {});
            const pEarly = (payload && payload.phase) || '';
            // 附件解析、读表、管线状态句是具体的；紧接着的 turn/step 往往只是「正在准备/正在执行步骤」，
            // 会把刚写上的句子盖掉。思考、工具、连接这些更具体的阶段仍然立即覆盖。
            if (label && (pEarly === 'turn' || pEarly === 'step')
                && setAgentPhaseStatus._explicitAt
                && (Date.now() - setAgentPhaseStatus._explicitAt) < 1500) {
                return;
            }
            if (!label) {
                if (payload && payload.status === 'idle') {
                    const p = payload.phase;
                    if (p === 'canceled') setAgentPhaseStatus(phaseT('phase.canceled'), { canceled: true });
                    else if (p === 'done') setAgentPhaseStatus(phaseT('phase.done'), { done: true });
                    else setAgentPhaseStatus('');
                }
                return;
            }
            const p = (payload && payload.phase) || '';
            setAgentPhaseStatus(label, { tool: p === 'tool', thinking: p === 'thinking' || p === 'connecting' });
        });
        ae.on('turn/end', function (payload) {
            const reason = (payload && payload.reason) || '';
            if (reason === 'user_aborted' || reason === 'replaced_by_new_turn') {
                setAgentPhaseStatus(phaseT('phase.canceled'), { canceled: true });
            } else if (reason) {
                setAgentPhaseStatus(phaseT('phase.done'), { done: true });
            }
        });
        // 多步计划：把当前子步骤镜像到输入区进度条
        ae.on('multi_step:substep_update', function (payload) {
            if (!payload) return;
            const name = payload.name || payload.title || payload.detail || '';
            if (!name) return;
            const n = payload.index != null ? (payload.index + 1) : payload.step;
            const total = payload.total || payload.stepTotal;
            let label = String(name);
            if (n && total) label = phaseT('phase.stepOf', { n: n, total: total }) + ' · ' + label;
            setAgentPhaseStatus(label, { tool: true });
        });
    }

    function setProcessingUI(processing) {
        const sendBtn = document.getElementById('send-btn');
        if (sendBtn) {
            sendBtn.textContent = processing ? phaseT('input.stop') : phaseT('input.send');
            sendBtn.classList.toggle('processing', processing);
            sendBtn.onclick = processing ? cancelCurrentRequest : sendMessage;
        }
        if (App.state) App.state.isProcessing = processing;
        if (processing) {
            setAgentPhaseStatus(phaseT('phase.turn'));
        }
    }

    // 当 turn 结束时由 renderer 调用，把按钮恢复为「发送」
    function resetProcessingUI() {
        setProcessingUI(false);
        // 轮次结束必须清掉进度条。dispatcher 短路不会发 turn/end，
        // 否则会一直停在「正在读取当前表格…」。
        setAgentPhaseStatus('');
    }

    // ============ 新建对话 ============

    function startNewChat() {
        // 取消正在进行的 turn
        App.ChatV2.orchestrator?.cancelCurrent('new_chat');

        // 保存当前对话到旧系统历史记录
        // 注意：如果当前会话是从历史记录加载的，不重复保存
        try {
            const sessionStore = App.ChatV2.sessionStore;
            if (sessionStore && typeof sessionStore.saveToLegacyHistory === 'function') {
                if (sessionStore.isLoadedFromHistory()) {
                    ChatLog.info?.('CHAT-V2', '[startNewChat] 当前会话来自历史记录，跳过保存');
                } else {
                    const messages = sessionStore.getCompletedTurnsMessages();
                    if (messages.length > 0) {
                        sessionStore.saveToLegacyHistory(true); // force 模式
                        ChatLog.info?.('CHAT-V2', '[startNewChat] 已保存当前对话到历史记录');
                    }
                }
            }
        } catch (e) {
            ChatLog.warn?.('CHAT-V2', '[startNewChat] 保存历史失败', e);
        }

        // 清空会话状态
        App.ChatV2.sessionStore?.clearHistory();
        // 清空会话级幂等台账（避免新对话被上一对话的终态工具结果误挡）
        try { App.ChatV2.agentLoop?.resetSessionLedger?.(); } catch (e) { /* 忽略 */ }
        // 清空 UI
        const messagesEl = document.getElementById('chat-messages');
        if (messagesEl) {
            messagesEl.innerHTML = '';
            try { messagesEl.classList.remove('has-pomodoro'); } catch (_) { /* ignore */ }
        }
        // 新建对话后展示产品欢迎面板
        if (typeof window.ensurePixelWelcome === 'function') {
            try { window.ensurePixelWelcome(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'ensurePixelWelcome 失败:', e && e.message); }
        } else if (typeof window.showWelcomePanel === 'function') {
            try { window.showWelcomePanel(); } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('CHAT-V2', 'showWelcomePanel 失败:', e && e.message); }
        }
        clearAttachment();
        resetProcessingUI();
        // 重置后端选区上下文（防止跨会话选区污染）
        try {
            if (App.excelBridge && App.excelBridge.ResetRangeContext) {
                App.excelBridge.ResetRangeContext();
            }
        } catch(e) { /* 忽略COM调用异常 */ }
        ChatLog.info?.('CHAT-V2', '[chat-entry] new chat started');
    }

    // ============ 键盘绑定：Ctrl/Cmd + Enter 发送 ============

    function bindKeyboard() {
        const textarea = document.getElementById('message-input');
        if (!textarea || textarea._chatV2KeyBound) return;
        textarea._chatV2KeyBound = true;
        textarea.addEventListener('keydown', e => {
            if (e.key === 'Enter' && (e.ctrlKey || e.metaKey) && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        });
        // 自适应高度
        textarea.addEventListener('input', () => {
            textarea.style.height = 'auto';
            textarea.style.height = Math.min(textarea.scrollHeight, 200) + 'px';
        });
    }

    function bindNewChatBtn() {
        const btn = document.getElementById('new-chat-btn');
        if (btn && !btn._chatV2Bound) {
            btn.addEventListener('click', startNewChat);
            btn._chatV2Bound = true;
        }
    }

    // ============ 初始化 ============

    function init() {
        // 暴露给 index.html 的 onclick 全局函数
        window.sendMessage = sendMessage;
        window.handleFileSelect = handleFileSelect;
        window.clearAttachment = clearAttachment;
        window.cancelCurrentRequest = cancelCurrentRequest;
        window.startNewChat = startNewChat;

        bindKeyboard();
        bindNewChatBtn();
        // agentEvents 可能稍后挂载：立即绑一次，并延迟再绑一次
        bindAgentPhaseStatus();
        setTimeout(bindAgentPhaseStatus, 300);

        ChatLog.info?.('CHAT-V2', '✅ chat-entry initialized and bound to UI');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    App.ChatV2.chatEntry = {
        sendMessage,
        handleFileSelect,
        clearAttachment,
        cancelCurrentRequest,
        startNewChat,
        setProcessingUI,
        resetProcessingUI,
        setPhaseStatus: function (text) { setAgentPhaseStatus(text, { explicit: true }); },
        init,
        // 暴露给测试/复用：附件转换链路（Excel/PDF/Word → 后端解析，文本直接读）
        fileToText,
        filesToAttachments,
        formatParsedDocument,
        _safeJsonLoose,
        withCallTimeout,
        DOC_PARSE_TIMEOUT_MS,
        PDF_RENDER_TIMEOUT_MS
    };
    ChatLog.info?.('CHAT-V2', '✅ chat-entry loaded');
})();