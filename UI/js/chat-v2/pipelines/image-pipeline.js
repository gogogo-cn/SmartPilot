/**
 * ============================================================
 * chat-v2/pipelines/image-pipeline.js — 图片分析流程
 * ============================================================
 * 职责：图片分析流程
 * 设计：视觉模型直接传图片 base64；非视觉模型走 OCR 文本降级
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    // ============ OCR 桥调用：多来源回退 ============
    // 背景：BridgeAccessor 的包装层对个别宿主方法会抛 "e is not a function"（裸宿主对象反而可用），
    // 代码库内 chat-renderer / mcp-dispatcher 均因此改用裸桥优先。OCR 方法（SaveBase64Image /
    // StartOcrRecognition / GetOcrResult / GetOcrStructuredResult）正好踩中此坑，且包装层 .then
    // 兜底未捕获到该异步拒绝 → 泄漏为未处理 Promise 拒绝。这里统一走「裸宿主对象优先 + 回退包装层」，
    // 任一来源成功即返回；全部失败才抛出，并由调用方翻译为可读提示。
    function getOcrBridgeCandidates() {
        const out = [];
        const seen = new Set();
        const push = function (b, label) {
            if (!b || seen.has(b)) return;
            seen.add(b);
            out.push({ b: b, label: label });
        };
        try { push(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge, 'hostObjects.excelBridge'); } catch (_) {}
        try { if (typeof App !== 'undefined') push(App.excelBridge, 'App.excelBridge'); } catch (_) {}
        try { push(window.excelBridge, 'window.excelBridge'); } catch (_) {}
        try { if (window.BridgeAccessor && typeof window.BridgeAccessor.getBridge === 'function') push(window.BridgeAccessor.getBridge(), 'BridgeAccessor.getBridge()'); } catch (_) {}
        return out;
    }

    async function callOcrBridge(method, argsArray) {
        // 本地 OCR 每日免费额度（云通道 / 已激活不计数）。详情见 js/ocr-quota.js
        var quota = window.SmartPilotOcrQuota;
        if (quota && quota.gate({ method: method, engine: window.__currentOcrEngine || '' }) === false) {
            throw new Error('OCR_QUOTA_EXCEEDED');
        }
        const cands = getOcrBridgeCandidates();
        if (cands.length === 0) throw new Error('OCR 需要 Excel 宿主环境支持');
        let lastErr = null;
        for (let i = 0; i < cands.length; i++) {
            const c = cands[i];
            try {
                if (!c.b || typeof c.b[method] !== 'function') continue;
                const raw = await c.b[method].apply(c.b, argsArray || []);
                return raw;
            } catch (e) {
                lastErr = e;
                ChatLog.warn?.('PIPELINE', `[image-pipeline] OCR 桥 ${method} 经 ${c.label} 失败: ${(e && e.message) || e}`);
            }
        }
        // 所有来源都失败：退还这次额度（只为真正启动的识别计费）
        if (quota) quota.refund({ method: method, engine: window.__currentOcrEngine || '' });
        if (lastErr) throw lastErr;
        throw new Error('桥方法 ' + method + ' 在所有来源均不可用');
    }

    // 把宿主桥的"方法未暴露"类原始错误翻译成用户可读、可执行的提示，
    // 不再把 "e is not a function" / "Object doesn't support this property" 之类的内部错误直接甩给用户。
    function friendlyOcrError(err) {
        const msg = (err && err.message) ? err.message : (err == null ? '' : String(err));
        if (/is not a function|doesn't support this property|object doesn't support|0x80020006|未知名称|undefined is not/i.test(msg)) {
            return '当前加载的插件未暴露 OCR 桥接方法（疑似旧版插件仍在运行，或 WebView2 投影尚未就绪）。请：① 完全退出 Excel 进程后重新打开；② 在「文件 → 选项 → 加载项」确认本插件已启用且为最新版本；③ 若仍失败，清掉 %LOCALAPPDATA%\\Microsoft\\Edge\\User Data\\WebView2Cache 后重试。';
        }
        return msg;
    }


    // OCR → 纯视觉兜底门控参数（代码内默认值；可由 App.config.ocrVisionFallbackThreshold / ocrFallbackNoTextMin 覆盖）
    const OCR_LOW_CONF_THRESHOLD = 0.6;   // 结构化平均置信度低于此值判定为低质
    const OCR_NO_TEXT_MIN_LENGTH = 4;     // 无置信度时，文本长度低于此值判定为空洞

    // 读取门控参数：优先 App.config 配置源，缺失/非法则回落代码内默认值
    function getFallbackGates() {
        const c = App.config || {};
        return {
            lowConf: (typeof c.ocrVisionFallbackThreshold === 'number' && c.ocrVisionFallbackThreshold > 0 && c.ocrVisionFallbackThreshold <= 1)
                ? c.ocrVisionFallbackThreshold
                : OCR_LOW_CONF_THRESHOLD,
            noTextMin: (Number.isInteger(c.ocrFallbackNoTextMin) && c.ocrFallbackNoTextMin >= 0)
                ? c.ocrFallbackNoTextMin
                : OCR_NO_TEXT_MIN_LENGTH
        };
    }

    function sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    /**
     * 调用 OCR 识别图片文字并附带质量信息
     * @param {string} base64 - 图片 base64（不含 data: 前缀）
     * @param {string} fileName - 文件名（用于保存临时路径）
     * @returns {Promise<{text:string, avgConfidence:number|null, itemCount:number}>}
     *   avgConfidence 为 OCR 结构化结果的文本平均置信度；取不到置信度时为 null（交由调用方按文本空洞判定）
     */
    async function recognizeImageViaOcr(base64, fileName, deadlineAt) {
        // 1. 保存 base64 图片到临时文件
        const saveResult = await callOcrBridge('SaveBase64Image', [base64, fileName]);
        const saveObj = JSON.parse(saveResult);
        if (!saveObj.success) {
            throw new Error('OCR 保存图片失败: ' + (saveObj.error || '未知错误'));
        }
        const filePath = saveObj.path;

        // 2. 启动 OCR 识别
        const startResult = await callOcrBridge('StartOcrRecognition', [filePath]);
        const startObj = JSON.parse(startResult);
        if (!startObj.started) {
            // 启动失败退还额度（只为真正启动的识别计费）
            if (window.SmartPilotOcrQuota) {
                window.SmartPilotOcrQuota.refund({ method: 'StartOcrRecognition', engine: window.__currentOcrEngine || '' });
            }
            throw new Error('OCR 启动失败: ' + (startObj.error || '未知错误'));
        }

        // 3. 轮询结果。未指定截止时间时单页最多 60 秒；
        // 指定 deadlineAt 时与同批其他页共用这一截止时间。
        const deadline = deadlineAt || (Date.now() + 60000);

        while (true) {
            await sleep(500);
            if (Date.now() > deadline) {
                throw new Error(deadlineAt
                    ? 'OCR 识别超时（本批页面共用 60 秒）'
                    : 'OCR 识别超时（超过 60 秒）');
            }

            const pollResult = await callOcrBridge('GetOcrResult', []);
            const pollObj = JSON.parse(pollResult);
            if (pollObj.ready) {
                if (pollObj.success) {
                    return {
                        text: pollObj.text || '(未识别到文字内容)',
                        avgConfidence: await readOcrConfidence(),
                        itemCount: 0
                    };
                }
                throw new Error('OCR 识别失败: ' + (pollObj.error || '未知错误'));
            }
        }
    }

    /**
     * 读取 OCR 结构化结果的平均置信度。
     * 后端 GetOcrStructuredResult() 返回 items:[{text,confidence,box}]；
     * 当无 items 或置信度全为 0（引擎未填充）时返回 null，表示"无法用置信度判定"。
     */
    async function readOcrConfidence() {
        try {
            const raw = await callOcrBridge('GetOcrStructuredResult', []);
            const obj = JSON.parse(raw);
            const items = (obj && obj.items) || [];
            const confs = items
                .map(function(it) { return Number(it && it.confidence); })
                .filter(function(c) { return Number.isFinite(c) && c > 0; });
            if (confs.length === 0) {
                return null;
            }
            return confs.reduce(function(a, b) { return a + b; }, 0) / confs.length;
        } catch (err) {
            ChatLog.warn?.('PIPELINE', '[image-pipeline] 读取 OCR 置信度失败，按未知处理', err);
            return null;
        }
    }

    /**
     * 判定 OCR 质量是否低到需要纯视觉兜底。
     * 优先级：有置信度 → 用阈值；无置信度 → 用文本空洞启发式。
     */
    function isOcrTooLowQuality(ocr) {
        const gates = getFallbackGates();
        if (ocr.avgConfidence !== null && ocr.avgConfidence < gates.lowConf) {
            return true;
        }
        const t = (ocr.text || '').trim();
        if (!t || /未识别到文字内容/.test(t) || t.length < gates.noTextMin) {
            return true;
        }
        return false;
    }

    /**
     * 纯视觉兜底：调用 ImageFieldExtractor skill 直接看图提取结构化结果。
     * 失败或不满足条件时返回 null，由调用方回落到 OCR 文本，保证主链路不丢。
     */
    async function fallbackToVisionSkill(base64, fileName) {
        try {
            const result = await callOcrBridge('ExecuteSkillAsync', [
                'ImageFieldExtractor',
                JSON.stringify({ imageData: base64, fileName: fileName || 'image' })
            ]);
            const resp = JSON.parse(result);
            if (!resp || !resp.success) {
                ChatLog.warn?.('PIPELINE', '[image-pipeline] ImageFieldExtractor 调用失败: ' + (resp && resp.error || '未知错误'));
                return null;
            }
            const out = resp.outputs || {};
            const texts = [resp.text, resp.rawText, out.result, out.content, out.markdown, resp.rawText];
            const txt = texts.find(function(t) { return typeof t === 'string' && t.trim(); });
            const fallback = Object.keys(out).length ? JSON.stringify(out) : '';
            const vision = (typeof txt === 'string' && txt.trim()) ? txt : fallback;
            if (!vision || /^\[?(无法识别|未识别|\[OCR)/.test(vision.trim())) {
                return null;
            }
            return vision;
        } catch (err) {
            ChatLog.warn?.('PIPELINE', '[image-pipeline] 纯视觉兜底失败，回落 OCR 文本', err);
            return null;
        }
    }

    /**
     * 从 base64 数据中提取纯 base64（去掉 data:xxx;base64, 前缀）
     */
    function stripBase64Prefix(data) {
        if (!data) return '';
        const commaIdx = data.indexOf(',');
        if (commaIdx !== -1 && data.startsWith('data:')) {
            return data.substring(commaIdx + 1);
        }
        return data;
    }

    /**
     * 用原图做 OCR，质量不够再交 ImageFieldExtractor。
     * 不压缩、不改附件、不改 turn.kind。图纸页的小字需要渲染分辨率。
     * OCR 桥是全局单任务：新的 StartOcrRecognition 会作废上一页的结果，所以按页排队。
     * budgetMs 让这一批共用一个截止时间，避免 3 页各等 60 秒。
     * @param {Array} imageAttachments
     * @param {{maxPages?: number, budgetMs?: number}} [options]
     */
    async function recognizeWithOcr(imageAttachments, options) {
        const maxPages = (options && options.maxPages) > 0 ? options.maxPages : 3;
        const budgetMs = (options && options.budgetMs) > 0 ? options.budgetMs : 0;
        const deadlineAt = budgetMs ? Date.now() + budgetMs : 0;
        const list = (imageAttachments || []).filter(function (att) { return att; }).slice(0, maxPages);
        const pages = [];
        for (let i = 0; i < list.length; i++) {
            const att = list[i];
            const base64Data = stripBase64Prefix(att.base64 || att.data || '');
            const attName = att.name || att.fileName || ('图片' + (i + 1));
            if (deadlineAt && Date.now() > deadlineAt) {
                pages.push({
                    index: i + 1,
                    name: attName,
                    text: '[OCR 识别超时（本批页面共用 60 秒）]',
                    usedVision: false
                });
                break;
            }
            if (!base64Data) {
                ChatLog.warn?.('PIPELINE', `[image-pipeline] attachment ${i} 无 base64 数据，跳过 OCR`);
                continue;
            }
            ChatLog.info?.('PIPELINE', `[image-pipeline] OCR 识别图片 ${i + 1}: ${attName}`);
            try {
                const ocr = await recognizeImageViaOcr(base64Data, attName, deadlineAt || undefined);
                let ocrText = ocr.text;
                let usedVision = false;
                if (isOcrTooLowQuality(ocr) && (!deadlineAt || Date.now() < deadlineAt)) {
                    const confStr = ocr.avgConfidence === null ? '未知' : ocr.avgConfidence;
                    ChatLog.warn?.('PIPELINE', `[image-pipeline] 图片 ${i + 1} OCR 质量低(conf=${confStr}, len=${(ocr.text || '').length})，尝试纯视觉兜底`);
                    const vision = await fallbackToVisionSkill(base64Data, attName);
                    if (vision) {
                        ocrText = vision;
                        usedVision = true;
                    }
                }
                pages.push({ index: i + 1, name: attName, text: ocrText, usedVision: usedVision });
                ChatLog.info?.('PIPELINE', `[image-pipeline] ${usedVision ? '视觉兜底' : 'OCR'} 完成: ${ocrText.substring(0, 100)}${ocrText.length > 100 ? '...' : ''}`);
            } catch (ocrErr) {
                const friendly = friendlyOcrError(ocrErr);
                ChatLog.error?.('PIPELINE', `[image-pipeline] OCR 失败: ${ocrErr.message}`);
                pages.push({
                    index: i + 1,
                    name: attName,
                    text: `[OCR 识别失败: ${friendly}]`,
                    usedVision: false
                });
            }
        }
        return {
            ok: pages.length > 0,
            pages: pages,
            pageCount: pages.length,
            usedVision: pages.some(function (r) { return r.usedVision; })
        };
    }

    async function* execute(turn, signal) {
        const { promptBuilder, toolRegistry, agentLoop, sessionStore } = App.ChatV2;

        try {
            await toolRegistry.init();

            const attachments = turn.attachments || [];
            const imageAttachments = attachments.filter(att => att.type === 'image');

            // 0. 判断当前模型是否支持视觉输入
            const currentModel = App.config?.model || '';
            const useVision = typeof App.isVisionModel === 'function' && App.isVisionModel(currentModel);

            // 保存原始用户消息，供渲染和编辑重发使用（OCR 会修改 userMessage）
            if (turn._originalUserMessage === undefined) {
                turn._originalUserMessage = turn.userMessage || '';
            }

            if (useVision && imageAttachments.length > 0) {
                // ========== 视觉模型：直接传图片 base64，不走 OCR ==========
                ChatLog.info?.('PIPELINE', `[image-pipeline] ${currentModel} 为视觉模型，使用图片直接输入`);

                // 清理 base64 前缀，确保 prompt-builder 构造 data URI 时不会重复
                for (const att of imageAttachments) {
                    const raw = att.base64 || att.data || '';
                    const clean = stripBase64Prefix(raw);
                    // 源头压缩图片，避免超长 base64 撑爆本地模型 context（超窗会被截断、进而同 turn 反复重试成多份表格）
                    try {
                        const maxDim = Number(App.config?.imageMaxDimension) || 1280;
                        const comp = await (App.ChatV2.imageCompress?.compressBase64 ? App.ChatV2.imageCompress.compressBase64(clean, maxDim) : clean);
                        att.base64 = comp || clean;
                        if (comp && comp !== clean) {
                            ChatLog.info?.('PIPELINE', `[image-pipeline] 视觉图片源头压缩: ${clean.length} → ${comp.length} chars`);
                        }
                    } catch (_) {
                        att.base64 = clean;
                    }
                    att.data = att.base64;
                }

                // 保持 kind='image'，让 prompt-builder 走 vision 格式构造
                turn.kind = 'image';
                turn.recognitionSource = currentModel + ' 视觉识别';
                if (!turn.userMessage?.trim()) {
                    turn.userMessage = '请识别图片内容并输出结构化结果。';
                    turn._originalUserMessage = ''; // 视觉模型路径：无用户输入，原始消息为空
                }
            } else {
                // ========== 非视觉模型：走 OCR 文本降级（原图，不压缩） ==========
                const ocrPack = await recognizeWithOcr(imageAttachments, { maxPages: imageAttachments.length || 3 });
                const ocrResults = ocrPack.pages || [];

                if (ocrResults.length > 0) {
                    const ocrBlock = ocrResults.map(r =>
                        `【图片${r.index}: ${r.name} 的OCR识别结果】\n${r.text}`
                    ).join('\n\n');

                    // 保存原始用户消息（如果还没保存）
                    if (turn._originalUserMessage === undefined) {
                        turn._originalUserMessage = turn.userMessage || '';
                    }

                    const originalMessage = turn._originalUserMessage || '';
                    if (originalMessage.trim()) {
                        turn.userMessage = `${originalMessage}\n\n${ocrBlock}`;
                    } else {
                        turn.userMessage = `请根据以下 OCR 识别结果分析图片内容：\n\n${ocrBlock}`;
                    }
                    turn._ocrDone = true;

                    // 任一图片走了纯视觉兜底 → 标记给渲染层显示徽章
                    const usedVisionFallback = ocrResults.some(function(r) { return r.usedVision; });
                    turn.usedVisionFallback = usedVisionFallback;
                    turn.recognitionSource = usedVisionFallback
                        ? 'ImageFieldExtractor 纯视觉兜底'
                        : 'OCR + 大模型识别';
                } else if (!turn.userMessage?.trim()) {
                    if (turn._originalUserMessage === undefined) {
                        turn._originalUserMessage = '';
                    }
                    turn.userMessage = '请上传图片或输入您的问题。';
                }
            }

            // 图片分析任务不需要工具调用
            const tools = [];

            // 构造 messages
            const history = sessionStore.getHistoryForLLM();
            let messages;
            try {
                messages = promptBuilder.buildMessages(turn, history, tools);
            } catch (buildErr) {
                ChatLog.error?.('PIPELINE', '[image-pipeline] buildMessages failed', buildErr);
                throw new Error(`构建消息失败: ${buildErr?.message || String(buildErr)}`);
            }

            ChatLog.info?.('PIPELINE', `[image-pipeline] turn=${turn.id} model=${currentModel} vision=${useVision} images=${imageAttachments.length} tools=${tools.length} messages=${messages?.length || 0}`);

            // 执行 agent loop
            try {
                yield* agentLoop.run({ turn, messages, tools, signal });
            } catch (loopErr) {
                ChatLog.error?.('PIPELINE', '[image-pipeline] agentLoop.run failed', loopErr);
                throw new Error(`模型调用失败: ${loopErr?.message || String(loopErr)}`);
            }

            // 完成 turn
            const finalTurn = sessionStore.getCurrentTurn();
            if (finalTurn && finalTurn.id === turn.id) {
                sessionStore.completeTurn(turn.id, {
                    content: finalTurn?.content || '',
                    reasoning: finalTurn?.reasoning || ''
                });
            }
            yield { type: 'pipeline_done', pipeline: 'image' };
        } catch (err) {
            ChatLog.error?.('PIPELINE', '[image-pipeline] error', err);
            sessionStore.failTurn(turn.id, { message: err?.message || String(err) });
            yield {
                type: 'error',
                error: { code: 'PIPELINE_ERROR', message: err?.message || String(err), recoverable: false }
            };
        }
    }

    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('image', {
            execute,
            name: '图片分析',
            description: '处理图片识别/分析（视觉模型直传或 OCR 文本降级）',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.image = { execute };
    }
    App.ChatV2.pageRecognition = { recognizeWithOcr: recognizeWithOcr, maxPages: 3 };
    ChatLog.info?.('CHAT-V2', '✅ image-pipeline loaded');
})();