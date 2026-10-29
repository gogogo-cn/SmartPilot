/**
 * ============================================================
 * chat-v2/services/llm-client.js — 模型抽象层
 * ============================================================
 * 职责：统一封装各家 LLM API（OpenAI/Moonshot/Qwen/Ollama 等）
 * 设计：所有 API 调用经过此模块，返回 typed event stream
 * 输出：async generator，yield 标准化事件
 *   { type: 'reasoning', delta }
 *   { type: 'token', delta }
 *   { type: 'tool_call', toolName, args, toolCallId }
 *   { type: 'usage', tokens: {in, out} }
 *   { type: 'done', finishReason }
 *   { type: 'error', error: { code, message, recoverable } }
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;
    const StreamProcessor = App.ChatV2.StreamProcessor;

    // ── 渲染提示注入（让 AI 能"选渲染器"）─────────────────────────────
    // 把可选的 _view 参数注入每个工具的 JSON Schema，模型才有可能输出它。
    // 不注入的话模型不会传 schema 之外的参数 → _view 通道形同虚设（实测 2026-10-03）。
    // 该参数是纯前端渲染提示：tool-registry.invoke 执行前会剥离，后端永远收不到。
    const VIEW_HINT_PROP = {
        type: 'string',
        enum: ['table', 'json', 'code', 'markdown', 'text', 'steps'],
        description: '结果渲染提示(可选)：指定本次工具结果在前端的展示方式——table 表格 / json 可折叠JSON / code 代码块 / markdown / text 纯文本 / steps 分步清单。仅当用户明确要求某种展示格式时才填写；不填则系统自动选择。'
    };
    /** 参数数组（[{name,type,description,required}]）→ JSON Schema 对象 */
    function _paramsArrayToSchema(arr) {
        const properties = {};
        const required = [];
        (Array.isArray(arr) ? arr : []).forEach(p => {
            if (!p || !p.name) return;
            properties[p.name] = { type: p.type || 'string', description: p.description || '' };
            if (p.required) required.push(p.name);
        });
        const schema = { type: 'object', properties: properties };
        if (required.length > 0) schema.required = required;
        return schema;
    }
    /** 在工具 schema 上追加可选 _view（返回新对象，避免污染注册表/其它消费方） */
    function withViewHint(schema) {
        const base = (schema && typeof schema === 'object' && !Array.isArray(schema))
            ? schema
            : (Array.isArray(schema) ? _paramsArrayToSchema(schema) : { type: 'object', properties: {} });
        const props = Object.assign({}, base.properties || {});
        if (!props._view) props._view = VIEW_HINT_PROP;
        return Object.assign({}, base, { properties: props });
    }

    /**
     * 根据 apiType 构造请求 URL
     * 处理配置中可能包含尾部路径的情况（如用户填了 http://localhost:11434/api/chat）
     * 若配置为空，回退到 PROVIDER_CONFIGS 中的默认值
     */
    function resolveApiUrl(config) {
        const apiType = config.apiType || 'ollama';
        let baseUrl = (config.apiUrl || '').trim().replace(/\/$/, '');

        // 兜底：apiUrl 为空时用 PROVIDER_CONFIGS 中的默认值
        if (!baseUrl && App.PROVIDER_CONFIGS?.[apiType]?.apiUrl) {
            baseUrl = App.PROVIDER_CONFIGS[apiType].apiUrl.replace(/\/$/, '');
        }

        if (apiType === 'ollama') {
            // 允许用户填 http://localhost:11434 或 http://localhost:11434/api/chat
            if (/\/api\/chat$/i.test(baseUrl)) {
                return baseUrl;
            }
            return (baseUrl || 'http://localhost:11434') + '/api/chat';
        }
        // OpenAI 兼容（OpenAI/Moonshot/Qwen/智谱/通义/讯飞/自定义）
        // 允许用户填 https://api.xxx.com/v1 或 https://api.xxx.com/v1/chat/completions
        if (/\/chat\/completions$/i.test(baseUrl)) {
            return baseUrl;
        }
        return (baseUrl || 'https://api.openai.com/v1') + '/chat/completions';
    }

    /**
     * 构造请求 headers
     */
    function buildHeaders(config) {
        const headers = { 'Content-Type': 'application/json' };
        if (config.apiKey) {
            headers['Authorization'] = 'Bearer ' + config.apiKey;
        }
        return headers;
    }

    /**
     * 从 OpenAI vision 内容数组中提取纯 base64 图片列表
     * OpenAI: [{type:'image_url', image_url:{url:'data:image/png;base64,XXX'}}, {type:'text', text:'...'}]
     * 返回: { text: '...', images: ['XXX'] }
     */
    function extractImagesFromVisionContent(contentArray) {
        const images = [];
        const textParts = [];
        for (const item of (contentArray || [])) {
            if (!item) continue;
            if (item.type === 'image_url' && item.image_url?.url) {
                const url = item.image_url.url;
                const base64 = url.replace(/^data:image\/\w+;base64,/, '');
                if (base64) images.push(base64);
            } else if (item.type === 'text') {
                textParts.push(item.text || '');
            }
        }
        return { text: textParts.join('\n'), images };
    }

    /**
     * 将消息转换为 Ollama 兼容的 vision 格式
     * Ollama /api/chat 官方推荐：message.images = ['base64', ...]，content 为纯文本
     * 经验证 qwen3-vl 对 content 数组中的 type:image 支持不稳定，images 字段更可靠
     *
     * 额外处理：
     *   - Ollama /api/chat 原生支持 tool_calls / role=tool 消息（0.4+），原样保留；
     *   - 仅清理空 tool_calls（Ollama 不接受空数组）与多余字段；
     *   - content 数组 (vision) 统一转文本 + images 字段。
     */
    function toOllamaMessages(messages) {
        const result = [];

        for (const raw of messages) {
            if (!raw) continue;

            // 1. role=tool：Ollama 原生 tool 消息（保留 tool_call_id）
            if (raw.role === 'tool') {
                const content = typeof raw.content === 'string'
                    ? raw.content
                    : JSON.stringify(raw.content == null ? '' : raw.content);
                result.push({
                    role: 'tool',
                    tool_call_id: raw.tool_call_id || '',
                    content: content
                });
                continue;
            }

            // 2. 处理 content 数组格式 (vision)
            let m = raw;
            if (Array.isArray(m.content)) {
                const { text, images } = extractImagesFromVisionContent(m.content);
                m = { ...m, content: text || '' };
                if (images.length > 0) {
                    m.images = images;
                }
            }

            // 3. assistant 的 tool_calls：保留原生结构（Ollama 原生 function calling）。
            //    【关键】Ollama 要求 function.arguments 为对象（实测：传 JSON 字符串会 400
            //    "Value looks like object, but can't find closing '}' symbol"，传对象正常）。
            //    空数组移除，避免 Ollama 报错。
            if (Array.isArray(m.tool_calls) && m.tool_calls.length > 0) {
                m.tool_calls = m.tool_calls.map(tc => {
                    const fn = tc && tc.function ? tc.function : tc;
                    if (fn && typeof fn.arguments === 'string') {
                        try {
                            fn.arguments = JSON.parse(fn.arguments);
                        } catch (parseErr) { /* 解析失败保留原字符串，由 Ollama 容错 */ }
                    }
                    return tc;
                });
            } else if (m.tool_calls) {
                delete m.tool_calls;
            }
            delete m.tool_call_id;

            result.push(m);
        }

        return result;
    }

    /**
     * 构造请求 body
     * Ollama /api/chat 参考: https://github.com/ollama/ollama/blob/main/docs/api.md
     *   - 只接受 model / messages / stream / options / format / keep_alive
     *   - options: { temperature, num_predict, num_ctx, ... }
     *   - vision 格式: message.images = ['base64', ...]（最稳定）
     */
    function buildBody({ messages, model, config, tools, stream = true, nativeTools }) {
        const apiType = config.apiType || 'ollama';

        // 归一化 messages（过滤空条目并确保 content 存在）
        let normalizedMessages = (messages || []).filter(m => m != null).map(m => {
            if (m.content == null) {
                return { ...m, content: '' };
            }
            // 纯字符串
            if (typeof m.content === 'string') return m;
            // 其他类型统一转字符串（理论上不应出现）
            if (!Array.isArray(m.content)) {
                return { ...m, content: String(m.content) };
            }
            // OpenAI 兼容：保留数组格式（vision），Ollama 在下面单独转换
            return { ...m, content: m.content };
        });

        if (apiType === 'ollama') {
            normalizedMessages = toOllamaMessages(normalizedMessages);
            const body = {
                model: model || config.model,
                stream: stream,
                messages: normalizedMessages,
                options: {
                    temperature: config.temperature ?? 0.7,
                    // 默认 16K 上下文：excel_read_range 返回的选区数据可能很大
                    // （101 行 ≈ 6000~8000 tokens），Ollama 默认 num_ctx 仅 4096，
                    // 会把输入截断，模型只能看到前几行导致“回显第一行”而非总结。
                    // 可在插件设置 config.options.num_ctx 显式覆盖。
                    num_ctx: config.options?.num_ctx ?? 16384,
                    // 【失控熔断·2026-10-18】必须显式设置 num_predict：
                    // 不设时 Ollama 默认 -1（无限生成），模型会一路写到 num_ctx 上限。
                    // 实测 qwen2.5:7b 解读 PDF 附件时进入退化循环（"1,1,2,2,3,3…"无限列举），
                    // 生成 14397 token、耗时 3分10秒、truncated=1——期间前端只有
                    // "正在分析你的需求"，用户感知"无限输出/没有过程渲染"。
                    // 默认取 maxTokens（设置里的输出上限，通常 4096），
                    // 可在 config.options.num_predict 覆盖。
                    num_predict: config.options?.num_predict ?? config.maxTokens ?? 4096
                }
            };
            ChatLog.info?.('LLM', `[buildBody Ollama] 输入: nativeTools=${nativeTools} toolsArr.length=${tools?.length || 0} → body.tools.length=${(tools && tools.length > 0 && nativeTools !== false) ? tools.length : 0}`);
            if (nativeTools !== false && tools && tools.length > 0) {
                body.tools = tools.map(t => ({
                    type: 'function',
                    function: {
                        name: t.name,
                        description: t.description,
                        parameters: withViewHint(t.inputSchema)
                    }
                }));
                body.tool_choice = 'auto';
            }
            try {
                const _msgChars = JSON.stringify(body.messages || []).length;
                // 【修时序】Ollama 诊断挪到 tools 赋值之后，之前永远打印 0
                ChatLog.info?.('LLM', `[Ollama] num_ctx=${body.options?.num_ctx ?? '未设置(走默认)'} messages=${body.messages?.length || 0}条 总字符=${_msgChars} tools=${body.tools?.length || 0} 含tool消息=${(body.messages || []).some(m => m.role === 'tool')}`);
            } catch (e) { ChatLog.debug?.('LLM', '[llm-client] Ollama 诊断日志输出失败:', e); }
            return body;
        }

        // OpenAI 兼容
        const resolvedModel = model || config.model;
        const isKimiThinking = /^kimi-k[2-9]\b/i.test(resolvedModel || '');
        const isDeepSeekV4 = /^deepseek-v4\b/i.test(resolvedModel || '');
        // DeepSeek V4 系列与 Moonshot 思考模型都只允许 temperature=1，否则 API 返回 400
        const temperature = (isKimiThinking || isDeepSeekV4)
            ? 1
            : (config.temperature ?? 0.7);

        if ((isKimiThinking || isDeepSeekV4) && config.temperature != null && config.temperature !== 1) {
            ChatLog.warn?.('LLM', `[llm-client] ${resolvedModel} 只允许 temperature=1，已强制（原设置 ${config.temperature} 已忽略）`);
        }

        // 思考模型需要更多输出 token（reasoning + content）
        const reasoningOutputMultiplier = isKimiThinking ? 2 : 1;
        const baseMaxTokens = config.maxTokens ?? 4096;
        const max_tokens = baseMaxTokens * reasoningOutputMultiplier;

        const body = {
            model: resolvedModel,
            stream: stream,
            messages: normalizedMessages,
            temperature: temperature,
            max_tokens: max_tokens
        };
        // nativeTools === false 时不下发原生 tools（能力档案 chat-only/本地通道决策，
        // 由 CapabilityProfile.resolveToolMode 统一决策，见 capability-profile.js）
        if (nativeTools !== false && tools && tools.length > 0) {
            body.tools = tools.map(t => ({
                type: 'function',
                function: {
                    name: t.name,
                    description: t.description,
                    parameters: withViewHint(t.inputSchema)
                }
            }));
            body.tool_choice = 'auto';
        }
        return body;
    }

    /**
     * 处理 fetch 错误，转换为标准 error event
     */
    function toErrorEvent(error, recoverable = false) {
        if (error?.name === 'AbortError') {
            return { type: 'error', error: { code: 'ABORTED', message: 'User canceled', recoverable: true } };
        }
        if (error?.message?.includes('Failed to fetch') || error?.message?.includes('NetworkError')) {
            return { type: 'error', error: { code: 'NETWORK', message: '网络连接失败', recoverable } };
        }
        return {
            type: 'error',
            error: {
                code: error?.code || 'UNKNOWN',
                message: error?.message || String(error),
                recoverable
            }
        };
    }

    // ============ 视觉图像压缩（防止超长 base64 撑爆本地模型 context，避免被截断/反复重试） ============
    function _loadImage(src) {
        return new Promise((resolve, reject) => {
            const im = new Image();
            im.onload = () => resolve(im);
            im.onerror = reject;
            im.src = src;
        });
    }

    async function compressDataUrl(dataUrl, maxDim) {
        try {
            const img = await _loadImage(dataUrl);
            let w = img.naturalWidth, h = img.naturalHeight;
            if (!w || !h) return dataUrl;
            if (w <= maxDim && h <= maxDim && !/data:image\/png/i.test(dataUrl)) {
                return dataUrl; // 尺寸未超限且非 png，保持原样
            }
            const scale = Math.min(1, maxDim / Math.max(w, h));
            w = Math.max(1, Math.round(w * scale));
            h = Math.max(1, Math.round(h * scale));
            const canvas = document.createElement('canvas');
            canvas.width = w; canvas.height = h;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, w, h); // 透明底转白（JPEG 不支持透明）
            ctx.drawImage(img, 0, 0, w, h);
            const out = canvas.toDataURL('image/jpeg', 0.7);
            return (out.length < dataUrl.length) ? out : dataUrl; // 仅当确实变小才替换
        } catch (_) {
            return dataUrl;
        }
    }

    function detectImageMime(b64) {
        if (/^iVBORw0KGgo/i.test(b64)) return 'png';
        if (/^\/9j\//i.test(b64)) return 'jpeg';
        if (/^R0lGOD/i.test(b64)) return 'gif';
        if (/^UklGR/i.test(b64)) return 'webp';
        return 'png';
    }

    async function compressVisionImages(messages, config) {
        const maxDim = (config && Number(config.imageMaxDimension)) || 1280;
        if (!messages) return messages;
        let before = 0, after = 0, compressed = 0, scanned = 0;
        for (const m of messages) {
            if (!m) continue;
            // 结构 A：content 数组里的 image_url（data: 前缀）
            const content = m.content;
            if (Array.isArray(content)) {
                for (const part of content) {
                    if (part && part.type === 'image_url' && part.image_url && typeof part.image_url.url === 'string') {
                        const url = part.image_url.url;
                        scanned++;
                        if (/^(https?:)?\/\//i.test(url) || /^blob:/i.test(url)) continue; // 远程/blob 不压
                        const isData = /^data:image\//i.test(url);
                        const full = isData ? url : 'data:image/' + detectImageMime(url) + ';base64,' + url;
                        before += url.length;
                        const next = await compressDataUrl(full, maxDim);
                        if (next !== full) {
                            part.image_url.url = isData ? next : next.replace(/^data:[^;]*;base64,/, '');
                            compressed++;
                        }
                        after += part.image_url.url.length;
                    }
                }
            }
            // 结构 B：Ollama 原生 message.images = [base64]（无 data: 前缀）
            if (Array.isArray(m.images)) {
                for (let i = 0; i < m.images.length; i++) {
                    const b64 = m.images[i];
                    if (typeof b64 !== 'string' || !b64) continue;
                    scanned++;
                    const dataUrl = 'data:image/' + detectImageMime(b64) + ';base64,' + b64;
                    before += b64.length;
                    const next = await compressDataUrl(dataUrl, maxDim);
                    if (next !== dataUrl) {
                        m.images[i] = next.replace(/^data:[^;]*;base64,/, '');
                        compressed++;
                    }
                    after += String(m.images[i]).length;
                }
            }
        }
        if (compressed) {
            const ratio = before > 0 ? Math.round((1 - after / before) * 100) : 0;
            ChatLog.info?.('LLM', `[llm-client] 图像压缩: ${compressed} 张, ${before} → ${after} chars (省 ${ratio}%)`);
        } else if (scanned === 0) {
            const overview = (messages || []).map((m, idx) => {
                const cArr = Array.isArray(m && m.content);
                const imgCount = cArr ? m.content.filter(p => p && p.type === 'image_url').length : 0;
                const longUrl = cArr ? m.content.filter(p => p && typeof p.image_url?.url === 'string' && p.image_url.url.length > 500).length : 0;
                const top = typeof m.visualData === 'string' ? 0 : (m ? Object.keys(m).join(',') : '');
                return `#${idx}:role=${m && m.role} content=${cArr ? 'array' : (m && typeof m.content)} img_url=${imgCount} longUrl=${longUrl} images=${Array.isArray(m && m.images) ? 'array[' + m.images.length + ']' : (m && typeof m.images)} keys=${top}`;
            });
            ChatLog.info?.('LLM', `[llm-client] 图像压缩未命中任何图片: ${overview.join(' | ')}`);
        }
        return messages;
    }

    /**
     * 真正的非流式调用（stream:false 一次性请求）
     * 与 complete()（流式收集包装）不同：本方法发 stream=false 的 HTTP 请求，
     * 用于流式通道失败时的降级路径，返回统一提取结果。
     * @returns {Promise<{model:string, reasoning:string, content:string, toolCalls:Array|null}>}
     * @throws {Error} HTTP 非 200 / 响应解析失败时抛出
     */
    async function completeNonStream({ messages, model, config, tools, signal, nativeTools }) {
        await compressVisionImages(messages, config);
        const url = resolveApiUrl(config);
        const headers = buildHeaders(config);
        const body = buildBody({ messages, model, config, tools, stream: false, nativeTools });
        ChatLog.info?.('LLM', `[llm-client] POST ${url} model=${body.model} stream=false`);
        ChatLog.debug?.('LLM', `[llm-client] request body:`, body);

        const response = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body), signal });
        if (!response.ok) {
            let errorMsg = `HTTP ${response.status} ${response.statusText}`;
            try {
                const text = await response.text();
                try {
                    const errBody = JSON.parse(text);
                    errorMsg = errBody?.error?.message || String(errBody?.error || errBody?.message || text);
                } catch (_) {
                    if (text) errorMsg = text.substring(0, 500);
                }
            } catch (e) { /* 忽略 body 读取失败 */ }
            const err = new Error(errorMsg);
            err.code = 'HTTP_' + response.status;
            throw err;
        }
        const json = await response.json();
        // 复用 StreamProcessor 的多协议提取（兼容 Ollama / OpenAI / 中转格式）
        const evt = new StreamProcessor()._extractFromParsed(json);
        if (evt.isError) {
            const err = new Error(evt.errorMessage || 'LLM 返回错误');
            err.code = 'LLM_ERROR';
            throw err;
        }
        return { model: evt.model || '', reasoning: evt.reasoning || '', content: evt.content || '', toolCalls: evt.toolCalls || null, usage: evt.usage || null };
    }

    /**
     * 非流式回退生成器：yield 降级请求的事件序列
     * @returns {AsyncGenerator} 返回值 true=回退成功（调用方应终止），false=回退失败（调用方继续原错误路径）
     */
    async function* _tryNonStreamFallback({ messages, model, config, tools, signal, nativeTools, reason }) {
        try {
            ChatLog.warn?.('LLM', `[llm-client] 流式通道${reason}，自动回退非流式请求（1次）`);
            const evt = await completeNonStream({ messages, model, config, tools, signal, nativeTools });
            if (evt.model) yield { type: 'model', model: evt.model };
            if (evt.reasoning) yield { type: 'reasoning', delta: evt.reasoning };
            if (evt.content) yield { type: 'token', delta: evt.content };
            if (evt.toolCalls && evt.toolCalls.length > 0) yield { type: 'tool_calls', toolCalls: evt.toolCalls };
            if (evt.usage) yield { type: 'usage', tokens: evt.usage };
            yield { type: 'done', finishReason: 'stop' };
            return true;
        } catch (fbErr) {
            if (fbErr?.name === 'AbortError') {
                yield { type: 'error', error: { code: 'ABORTED', message: 'User canceled', recoverable: true } };
                return true;
            }
            ChatLog.error?.('LLM', `[llm-client] 非流式回退失败: ${fbErr?.message || fbErr}`);
            return false;
        }
    }

    /**
     * 主调用入口 — 返回 typed event stream
     * @param {Object} params
     *   - messages: [{role, content}] 已组装好的消息数组（含 system）
     *   - model: 模型名（可选，默认 config.model）
     *   - config: App.config
     *   - tools: 工具 schema 数组（可选）
     *   - nativeTools: 是否下发原生 tools（可选，false 时即使传入 tools 也不进 body，
     *     由 CapabilityProfile.resolveToolMode 决策；缺省视为 true 保持旧行为）
     *   - signal: AbortSignal（可选，用于取消）
     * @returns {AsyncGenerator} typed events
     */
    async function* stream({ messages, model, config, tools, signal, nativeTools }) {
        // 视觉图像先压缩，避免超长 base64 撑爆本地模型 context（超窗会被截断、进而同 turn 反复重试）
        await compressVisionImages(messages, config);
        const url = resolveApiUrl(config);
        const headers = buildHeaders(config);
        const body = buildBody({ messages, model, config, tools, stream: true, nativeTools });

        // 视觉请求诊断日志（统计图片数量/大小，不输出完整 base64）
        const visionInfo = (body.messages || []).map((m, idx) => {
            if (!m.images || !Array.isArray(m.images)) return null;
            return `msg#${idx}: ${m.images.length} image(s), total ${m.images.reduce((sum, b64) => sum + (b64?.length || 0), 0)} chars`;
        }).filter(Boolean);
        if (visionInfo.length > 0) {
            ChatLog.info?.('LLM', `[llm-client] vision request: ${visionInfo.join('; ')}`);
        }

        ChatLog.info?.('LLM', `[llm-client] POST ${url} model=${body.model} stream=true`);
        ChatLog.debug?.('LLM', `[llm-client] request body:`, body);

        // 【双熔断·2026-10-18】provider 无关的读流保护：
        // 1) stall 看门狗：模型服务冷加载/显存不足/被其他请求占满时可能数分钟不吐字节，
        //    而 fetch 只带外部 abort signal、没有任何超时 → 前端无限静默
        //    （实测 20:43→20:49 六分钟零反馈，用户只能反复重发）。
        //    连续 STALL_TIMEOUT_MS 未收到数据即取消 reader 并报错，交还给上层。
        // 2) chunk 数硬上限：模型退化循环时（reasoning 里 "1,1,2,2,3,3…" 无限列举、
        //    正文反复复读），即使服务端忽略 num_predict，客户端也能强制收尾，
        //    UI 立即可见"已自动截断"提示，而不是无限滚屏/空白气泡。
        // 两者都可用 config.options 覆盖（测试用小值验证）。
        const STALL_TIMEOUT_MS = Number(config?.options?.streamStallTimeoutMs) > 0
            ? Number(config.options.streamStallTimeoutMs) : 180000;
        const MAX_STREAM_CHUNKS = Number(config?.options?.maxStreamChunks) > 0
            ? Number(config.options.maxStreamChunks) : 8192;

        let response;
        try {
            response = await fetch(url, {
                method: 'POST',
                headers,
                body: JSON.stringify(body),
                signal
            });
        } catch (err) {
            // 用户取消 / 已中止：不回退
            if (err?.name === 'AbortError' || signal?.aborted) {
                yield toErrorEvent(err, true);
                return;
            }
            // 【非流式回退】连接层失败（网络错误/代理不支持流式）→ 自动回退非流式重试一次
            const fbOk = yield* _tryNonStreamFallback({ messages, model, config, tools, signal, nativeTools, reason: `连接失败(${err?.message || err})` });
            if (fbOk) return;
            yield toErrorEvent(err, true);
            return;
        }

        if (!response.ok) {
            let errorMsg = `HTTP ${response.status} ${response.statusText}`;
            let recoverable = false;
            try {
                const text = await response.text();
                try {
                    const errBody = JSON.parse(text);
                    if (errBody?.error?.message) errorMsg = errBody.error.message;
                    else if (errBody?.error) errorMsg = String(errBody.error);
                    else if (errBody?.message) errorMsg = String(errBody.message);
                    if (errBody?.error?.code === 'invalid_api_key') recoverable = false;
                } catch (_) {
                    if (text) errorMsg = text.substring(0, 500);
                }
                // 识别常见的视觉不支持错误，给出更友好的提示
                let visionUnsupported = false;
                const lowerMsg = errorMsg.toLowerCase();
                if (lowerMsg.includes('image') && lowerMsg.includes('not supported') ||
                    lowerMsg.includes('vision') && lowerMsg.includes('not supported')) {
                    errorMsg = '当前模型不支持图片分析，请切换到视觉模型（如 kimi-k3、llava、qwen2.5-vl 等）后再试。';
                    recoverable = false;
                    visionUnsupported = true;
                }
                if (response.status === 429) recoverable = true;
                ChatLog.error?.('LLM', `[llm-client] HTTP ${response.status}: ${errorMsg}`);
                // 【非流式回退】流式请求被网关/服务端拒绝时（部分中转 API 不支持 stream:true），
                // 自动回退非流式重试一次。鉴权/限流/视觉不支持等明确无意义场景跳过。
                if (!visionUnsupported && response.status !== 401 && response.status !== 403 &&
                    response.status !== 429 && !signal?.aborted) {
                    const fbOk = yield* _tryNonStreamFallback({ messages, model, config, tools, signal, nativeTools, reason: `被拒绝(HTTP ${response.status})` });
                    if (fbOk) return;
                }
            } catch (e) {
                ChatLog.error?.('LLM', `[llm-client] HTTP ${response.status} (无法读取 body)`);
            }
            yield {
                type: 'error',
                error: { code: 'HTTP_' + response.status, message: errorMsg, recoverable }
            };
            return;
        }

        // 处理非流式响应（某些 API 可能返回完整 JSON）
        const contentType = response.headers.get('content-type') || '';
        if (!contentType.includes('text/event-stream') && !response.body) {
            try {
                const json = await response.json();
                const sp = new StreamProcessor();
                const evt = sp._extractFromParsed(json);
                if (evt.model) yield { type: 'model', model: evt.model };
                if (evt.reasoning) yield { type: 'reasoning', delta: evt.reasoning };
                if (evt.content) yield { type: 'token', delta: evt.content };
                if (evt.toolCalls) yield { type: 'tool_calls', toolCalls: evt.toolCalls };
                if (evt.usage) yield { type: 'usage', tokens: evt.usage };
                yield { type: 'done', finishReason: 'stop' };
            } catch (e) {
                yield toErrorEvent(e, false);
            }
            return;
        }

        // 流式读取
        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        const processor = new StreamProcessor();
        let buffer = '';
        let modelEmitted = false;
        let _tN = 0, _tChars = 0, _rN = 0, _rChars = 0;
        let gotToolCalls = false;
        // stall 看门狗状态：cancel() 会让挂起的 read() 以 {done:true} 返回，
        // 用这个标记区分"流正常结束"与"被看门狗掐断"。
        let stallTimer = null;
        let streamStalled = false;

        try {
            while (true) {
                // 每次读流前重置 stall 计时（收到数据即算活跃）
                if (stallTimer) clearTimeout(stallTimer);
                stallTimer = setTimeout(() => {
                    streamStalled = true;
                    try { reader.cancel(); } catch (_e) { /* cancel 失败由下方 read 结果分支处理 */ }
                }, STALL_TIMEOUT_MS);
                const { done, value } = await reader.read();
                if (stallTimer) { clearTimeout(stallTimer); stallTimer = null; }
                if (streamStalled) {
                    ChatLog.error?.('LLM', `[llm-client] 🛑 读流 stall：连续 ${STALL_TIMEOUT_MS}ms 无数据，已中止（content=${_tChars}字 reasoning=${_rChars}字）`);
                    yield {
                        type: 'error',
                        error: {
                            code: 'STREAM_STALL',
                            message: `模型响应中断：连续 ${Math.round(STALL_TIMEOUT_MS / 1000)} 秒没有收到新内容，已停止等待。请检查模型服务（如 Ollama）是否正常运行，然后重试。`,
                            recoverable: true
                        }
                    };
                    return;
                }
                if (done) break;

                buffer += decoder.decode(value, { stream: true });
                // 按行处理
                const lines = buffer.split('\n');
                buffer = lines.pop() || '';

                for (const line of lines) {
                    const evt = processor.processText(line + '\n');
                    for (const e of evt) {
                        // 【失控熔断】chunk 总数（正文+思考）硬上限，到达即强制收尾。
                        // 放在计数之前，保证绝不会超出上限太多。
                        if (_tN + _rN >= MAX_STREAM_CHUNKS) {
                            ChatLog.warn?.('LLM', `[llm-client] 🛑 失控输出熔断：chunk 数达上限 ${MAX_STREAM_CHUNKS}（content=${_tChars}字 reasoning=${_rChars}字），强制收尾`);
                            yield { type: 'done', finishReason: 'runaway_truncated' };
                            return;
                        }
                        if (e.reasoning) { _rN++; _rChars += String(e.reasoning).length; }
                        if (e.content) { _tN++; _tChars += String(e.content).length; }
                        if (e.model && !modelEmitted) {
                            yield { type: 'model', model: e.model };
                            modelEmitted = true;
                        }
                        if (e.reasoning) yield { type: 'reasoning', delta: e.reasoning };
                        if (e.content) yield { type: 'token', delta: e.content };
                        if (e.toolCalls) { gotToolCalls = true; yield { type: 'tool_calls', toolCalls: e.toolCalls }; }
                        if (e.usage) yield { type: 'usage', tokens: e.usage };
                        if (e.isError) {
                            yield { type: 'error', error: { code: 'STREAM', message: e.errorMessage, recoverable: false } };
                            return;
                        }
                        if (e.isDone) {
                            const reason = e.finishReason || 'stop';
                            if (reason === 'length') {
                                ChatLog.warn?.('LLM', `[llm-client] 模型输出因长度限制被截断，请尝试减小输入或增加 max_tokens`);
                            }
                            console.info('[LLM] stream done fR=' + reason + ' tokenN=' + _tN + ' tokenChars=' + _tChars + ' reasonN=' + _rN + ' reasonChars=' + _rChars);
                            yield { type: 'done', finishReason: reason };
                            return;
                        }
                    }
                }
            }
            // 处理剩余 buffer（流末尾用 processFinalBuffer，不修复不完整 JSON）
            if (buffer.trim()) {
                const evt = processor.processFinalBuffer(buffer + '\n');
                for (const e of evt) {
                    // 同主循环的失控熔断（尾包同样可能携带大量事件）
                    if (_tN + _rN >= MAX_STREAM_CHUNKS) {
                        ChatLog.warn?.('LLM', `[llm-client] 🛑 失控输出熔断（尾包）：chunk 数达上限 ${MAX_STREAM_CHUNKS}，强制收尾`);
                        yield { type: 'done', finishReason: 'runaway_truncated' };
                        return;
                    }
                    if (e.reasoning) { _rN++; _rChars += String(e.reasoning).length; }
                    if (e.content) { _tN++; _tChars += String(e.content).length; }
                    if (e.model && !modelEmitted) {
                        yield { type: 'model', model: e.model };
                        modelEmitted = true;
                    }
                    if (e.reasoning) yield { type: 'reasoning', delta: e.reasoning };
                    if (e.content) yield { type: 'token', delta: e.content };
                    if (e.toolCalls) { gotToolCalls = true; yield { type: 'tool_calls', toolCalls: e.toolCalls }; }
                    if (e.usage) yield { type: 'usage', tokens: e.usage };
                }
            }
            // 流自然结束（未收到 [DONE]）
            if (!processor.isDone()) {
                yield { type: 'done', finishReason: 'stream_ended' };
            }
        } catch (err) {
            // 【非流式回退】尚未向调用方输出任何内容时（连接中断/SSE 协议解析失败），
            // 回退非流式重试一次；已有部分输出则不回退（避免内容重复）
            if (_tN === 0 && _rN === 0 && !gotToolCalls && !signal?.aborted) {
                const fbOk = yield* _tryNonStreamFallback({ messages, model, config, tools, signal, nativeTools, reason: `中断(${err?.message || err})` });
                if (fbOk) return;
            }
            yield toErrorEvent(err, false);
        } finally {
            if (stallTimer) { clearTimeout(stallTimer); stallTimer = null; }
            try { reader.cancel(); } catch (e) { ChatLog.warn?.('LLM', '[llm-client] 关闭流式 reader 失败:', e); }
        }
    }

    /**
     * 非流式调用（用于简单场景或工具调用的第二轮）
     */
    async function complete({ messages, model, config, tools, signal }) {
        const events = [];
        for await (const evt of stream({ messages, model, config, tools, signal })) {
            events.push(evt);
            if (evt.type === 'error') throw new Error(evt.error.message);
            if (evt.type === 'done') break;
        }
        return events;
    }

    // 可视图像压缩工具：供 pipeline 在源头直接压缩 base64（最可靠，不依赖消息内部结构）
    async function compressBase64(base64, maxDim) {
        if (!base64 || typeof base64 !== 'string') return base64;
        try {
            const dataUrl = 'data:image/' + detectImageMime(base64) + ';base64,' + base64;
            const out = await compressDataUrl(dataUrl, maxDim ? Number(maxDim) : 1280);
            return (out === dataUrl) ? base64 : String(out).replace(/^data:[^;]*;base64,/, '');
        } catch (_) {
            return base64;
        }
    }

    App.ChatV2.llmClient = { stream, complete, completeNonStream, resolveApiUrl, buildBody };
    App.ChatV2.imageCompress = { compressBase64, compressDataUrl, detectImageMime };
    ChatLog.info?.('CHAT-V2', '✅ llmClient loaded');
})();