/**
 * ============================================================
 * chat-v2/pipelines/doc-pipeline.js — 文档附件分析流程
 * ============================================================
 * 职责：处理 Excel/Word/PDF/Text 附件上传 + 分析请求
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    // ============================================================
    // 后续消息继承（Phase 5）：读取上一轮 doc _execSpec，注入上轮文档内容
    // ============================================================
    function getDocFollowUp(turn) {
        if (!turn || !turn.context || !turn.context._execFollowUp) return null;
        const ss = App.ChatV2.sessionStore;
        if (!ss || typeof ss.getLastExecSpec !== 'function') return null;
        const prev = ss.getLastExecSpec();
        if (!prev || !prev.spec || prev.spec.pipeline !== 'doc') return null;
        return prev.spec;
    }

    function applyDocFollowUp(turn) {
        const prevSpec = getDocFollowUp(turn);
        if (!prevSpec) return;
        if (turn._originalUserMessage === undefined) {
            turn._originalUserMessage = turn.userMessage || '';
        }
        const prevContent = prevSpec.result?.content || '';
        if (prevContent) {
            turn.userMessage = (turn.userMessage || '') +
                '\n\n【上一轮文档内容】\n' + String(prevContent).slice(0, 3000) +
                '\n\n请基于以上文档内容，按我的新需求修改/补充。只输出修改后的完整内容。';
        }
        ChatLog.info?.('PIPELINE', '[doc-pipeline] 后续消息继承: 复用上一轮文档内容');
    }

    function writeDocExecSpec(turn, sessionStore, content) {
        if (!sessionStore || typeof sessionStore.setExecSpec !== 'function') return;
        const atts = Array.isArray(turn.attachments) ? turn.attachments.map(function (a) {
            return { name: a.name || a.fileName || '', type: a.type || a.mimeType || '' };
        }) : [];
        // v20260924：用工厂统一构建 data，doc 专用 attachments 挂 data.ext.doc 命名空间
        const fact = sessionStore.execSpecFactory;
        const commonData = fact && typeof fact.buildCommonData === 'function'
            ? fact.buildCommonData(turn, { doc: { attachments: atts } })
            : { attachments: atts }; // 兜底
        sessionStore.setExecSpec(turn.id, {
            pipeline: 'doc',
            data: commonData,
            params: {},
            result: { content: content || '' },
            ts: Date.now()
        });
    }

    function _feedToolResult(result) {
        if (!result) return '';
        if (!result.ok) return String(result.error || '读取失败');
        const payload = result.result;
        if (!payload) return '';
        if (payload.visionImages) delete payload.visionImages;
        const msg = payload.message || payload.content || '';
        return String(msg).slice(0, 8000);
    }

    function _wantsCurrentSheet(msg) {
        return /工作表|当前表|这个表|活动表|excel|写入|写回|写到|对照|单元格|选区|表格/i.test(String(msg || ''));
    }

    function _toolName(tool) {
        return String((tool && (tool.name || tool.originalName)) || '')
            .toLowerCase()
            .replace(/^excel-local__/, '');
    }

    function _putDocument(messages, text) {
        if (!messages || !text) return;
        for (let i = messages.length - 1; i >= 0; i--) {
            const m = messages[i];
            if (m && m.role === 'user' && typeof m.content === 'string') {
                m.content += '\n<document>\n' + text + '\n</document>';
                return;
            }
        }
    }

    function _pushVisionPages(messages, images) {
        if (!messages || !images || !images.length) return;
        const content = [{ type: 'text', text: '下面是图纸页面（最长边不超过 1600 像素）。请直接读取图中的尺寸、公差和标题栏，只写图上有的内容。' }];
        images.forEach(function (img) {
            const b64 = img && (img.base64 || img.data);
            if (!b64) return;
            const mime = (img && img.mime) || 'image/png';
            content.push({ type: 'image_url', image_url: { url: 'data:' + mime + ';base64,' + b64 } });
        });
        if (content.length > 1) messages.push({ role: 'user', content: content });
    }

    async function* execute(turn, signal) {
        const { promptBuilder, toolRegistry, agentLoop, sessionStore } = App.ChatV2;

        try {
            await toolRegistry.init();

            // 生命周期（Phase 4）：新任务（非后续消息）时清空上一轮执行上下文
            if (!turn.context?._execFollowUp && sessionStore.clearExecSpec) {
                sessionStore.clearExecSpec();
            }
            // 后续消息继承（Phase 5）：注入上轮文档内容
            applyDocFollowUp(turn);

            // 只上传文件、没有问题时，不要把空消息交给工具瘦身（空消息会放行全部工具，小模型直接空回复）。
            const _rawAsk = String(turn._originalUserMessage || turn.userMessage || turn.message || '').trim();
            if (!_rawAsk) {
                const _fallbackQ = '概述这份附件的正文。只根据读到的内容说，不要编造。';
                turn._originalUserMessage = '';
                turn.userMessage = _fallbackQ;
                turn.message = _fallbackQ;
            }

            // 附件正文不预塞进提示词。系统先读，写入类 Excel 工具仍留给同一次规划。
            const reader = App.ChatV2.attachmentReadTool;
            let tools = [];
            try {
                const mcp = typeof toolRegistry.listBySource === 'function' ? (toolRegistry.listBySource('mcp') || []) : [];
                tools = mcp.filter(function (t) { return !(toolRegistry.isImageOnlyTool && toolRegistry.isImageOnlyTool(t)); });
                const filterFn = App.ChatV2.toolRegistryUtil && App.ChatV2.toolRegistryUtil.filterToolsForMessage;
                const msg = turn._originalUserMessage || turn.userMessage || turn.message || '';
                if (typeof filterFn === 'function' && tools.length) {
                    const filtered = filterFn(tools, msg, {});
                    if (filtered && filtered.length) tools = filtered;
                }
            } catch (_toolErr) {
                tools = [];
            }
            if (reader && typeof reader.pinOn === 'function') tools = reader.pinOn(tools, turn);
            if (!tools.length && reader && typeof reader.planningTools === 'function') {
                tools = reader.planningTools(toolRegistry, turn);
            }

            // 文档消息构造在 promptBuilder.buildUserMessage 中处理（XML 标签包裹）
            const history = [];  // 文档分析通常不带历史
            const messages = promptBuilder.buildMessages(turn, history, tools);

            ChatLog.info?.('PIPELINE', `[doc-pipeline] turn=${turn.id} attachments=${turn.attachments?.length || 0}`);
            const _docName = (turn.attachments && turn.attachments[0] && (turn.attachments[0].fileName || turn.attachments[0].name)) || '附件';
            let _primedText = '';
            if (toolRegistry && typeof toolRegistry.invokeTracked === 'function'
                && Array.isArray(turn.attachments) && turn.attachments.some(function (a) { return a && a.type !== 'image'; })) {
                yield { type: 'status', text: `正在读取「${_docName}」…` };
                try {
                    const _read = await toolRegistry.invokeTracked(turn, 'read_attachment', {}, {
                        round: 0, ctx: { turn, signal }
                    });
                    const _readText = _feedToolResult(_read);
                    _primedText = _readText;
                    const _ask = String(turn.userMessage || turn.message || '');
                    const _marked = !!(_read && _read.ok && _read.result && _read.result.needsRecognition);
                    const _ocrDraw = !!(reader && typeof reader.drawingAskNeedsOcr === 'function'
                        && reader.drawingAskNeedsOcr(_ask, _readText, _docName));
                    if (_marked || _ocrDraw) {
                        yield { type: 'status', text: `正在识别「${_docName}」的页面…` };
                        const _rec = await toolRegistry.invokeTracked(turn, 'recognize_pages', {}, {
                            round: 0, ctx: { turn, signal }
                        });
                        const _imgs = _rec && _rec.result && _rec.result.visionImages;
                        const _recText = _feedToolResult(_rec);
                        if (_recText) _primedText = _recText;
                        _pushVisionPages(messages, _imgs);
                    }
                    _putDocument(messages, _primedText);
                    const _sheet = _wantsCurrentSheet(_ask);
                    if (!_sheet) turn._attachmentAnswerOnly = true;
                    tools = tools.filter(function (t) {
                        const n = _toolName(t);
                        if (n === 'read_attachment' || n === 'recognize_pages') return false;
                        if (!_sheet && n.indexOf('excel_') === 0) return false;
                        return true;
                    });
                } catch (primeErr) {
                    ChatLog.warn?.('PIPELINE', '[doc-pipeline] 预读附件失败，改由模型自己调用工具', primeErr);
                }
            }
            yield { type: 'status', text: `正在根据「${_docName}」组织回答…` };

            yield* agentLoop.run({ turn, messages, tools, signal });

            const finalTurn = sessionStore.getCurrentTurn();
            if (finalTurn && finalTurn.id === turn.id) {
                let _content = finalTurn?.content || '';
                if (!_content.trim() && _primedText) {
                    _content = _primedText;
                    sessionStore.updateTurn(turn.id, { content: _content });
                }
                sessionStore.completeTurn(turn.id, {
                    content: _content,
                    reasoning: finalTurn?.reasoning || ''
                });
                writeDocExecSpec(turn, sessionStore, _content);
            }
            yield { type: 'pipeline_done', pipeline: 'doc' };
        } catch (err) {
            ChatLog.error?.('PIPELINE', '[doc-pipeline] error', err);
            sessionStore.failTurn(turn.id, { message: err?.message || String(err) });
            yield {
                type: 'error',
                error: { code: 'PIPELINE_ERROR', message: err?.message || String(err), recoverable: false }
            };
        }
    }

    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('doc', {
            execute,
            name: '文档分析',
            description: '处理 Excel/Word/PDF/Text 附件上传 + 分析请求',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.doc = { execute };
    }
    ChatLog.info?.('CHAT-V2', '✅ doc-pipeline loaded');
})();
