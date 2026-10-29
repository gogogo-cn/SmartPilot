/**
 * ============================================================
 * chat-v2/services/stream-processor.js — SSE 流式响应解析器
 * ============================================================
 * 职责：解析 SSE 协议 + 提取 JSON 中的 content/reasoning_content
 * 设计：单一来源，所有 LLM 流式响应必须经过此模块
 * 输出：{ content, reasoning, isDone, isError, errorMessage }
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * StreamProcessor — 流式响应处理器
     * 处理 SSE 协议 + 多家 API 格式（OpenAI 兼容 / Ollama / Anthropic 风格）
     */
    class StreamProcessor {
        constructor() {
            this._buffer = '';
            this._isDone = false;
            this._isError = false;
            this._errorMessage = '';
            // 【致命修复】tool_calls 流式增量累积缓存
            // Moonshot/Kimi/OpenAI 的 tool_calls 是分 chunk 增量传输的：
            //   chunk1: {index:0, function:{name:"excel_read_range", arguments:""}}
            //   chunk2: {index:0, function:{arguments:"{\\"range\\":..."}}  ← 无 name！
            //   chunk3: {index:1, function:{name:"excel_aggregate", arguments:""}}
            // 如果不合并直接覆盖 → 最后一个 chunk (只有 arguments 无 name) 覆盖前面的 → 空名 tool_call！
            this._accumulatedToolCalls = null;
        }

        /**
         * 处理原始文本 chunk（来自 fetch reader.read()）
         * @param {string} text - 文本片段
         * @returns {Array<{content, reasoning, isDone, isError, errorMessage}>} 解析出的事件列表
         */
        processText(text) {
            if (this._isDone) return [];
            this._buffer += text;

            const events = [];
            const lines = this._buffer.split('\n');
            // 最后一行可能不完整，保留在 buffer
            this._buffer = lines.pop() || '';

            for (const line of lines) {
                const evt = this._processLine(line);
                if (evt) events.push(evt);
                if (evt?.isDone || evt?.isError) {
                    this._isDone = evt.isDone;
                    this._isError = evt.isError || false;
                    this._errorMessage = evt.errorMessage || '';
                    break;
                }
            }
            return events;
        }

        /**
         * 处理流末尾的剩余 buffer（不修复不完整 JSON，避免误判为错误）
         * @param {string} text - 剩余文本片段
         * @returns {Array<{content, reasoning, isDone, isError, errorMessage}>} 解析出的事件列表
         */
        processFinalBuffer(text) {
            if (this._isDone) return [];
            this._buffer += text;

            const events = [];
            const lines = this._buffer.split('\n');
            this._buffer = lines.pop() || '';

            for (const line of lines) {
                const evt = this._processLine(line, true);
                if (evt) events.push(evt);
                if (evt?.isDone || evt?.isError) {
                    this._isDone = evt.isDone;
                    this._isError = evt.isError || false;
                    this._errorMessage = evt.errorMessage || '';
                    break;
                }
            }
            return events;
        }

        _processLine(line, skipFix = false) {
            const trimmed = line.trim();
            if (!trimmed) return null;

            // 结束标记
            if (trimmed === 'data: [DONE]' || trimmed === '[DONE]') {
                return { content: '', reasoning: '', isDone: true };
            }

            // error 行
            if (trimmed.startsWith('error:')) {
                return {
                    content: '', reasoning: '',
                    isDone: true, isError: true,
                    errorMessage: trimmed.slice(6).trim()
                };
            }

            // data: 前缀
            let jsonData = null;
            if (trimmed.startsWith('data:')) {
                jsonData = trimmed.slice(5).trim();
            } else if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
                // 某些 API 不遵循标准 SSE，直接给 JSON
                jsonData = trimmed;
            } else {
                // 忽略 id:/event:/comment 行
                return null;
            }

            if (!jsonData) return null;

            const parsed = this._parseJSON(jsonData, skipFix);
            if (!parsed) {
                // 不完整的 JSON 在流式场景下常见，跳过等待后续 chunk
                return null;
            }

            return this._extractFromParsed(parsed);
        }

        _parseJSON(data, skipFix = false) {
            try {
                return JSON.parse(data);
            } catch (e) {
                // 流末尾不修复不完整 JSON，避免误判为错误
                if (skipFix) return null;
                // 尝试修复未闭合的 JSON
                const fixed = this._fixIncompleteJSON(data);
                if (fixed) {
                    try { return JSON.parse(fixed); }
                    catch (e2) { return null; }
                }
                return null;
            }
        }

        _fixIncompleteJSON(data) {
            if (!data) return null;
            let fixed = data.trim();
            // 修复未闭合的引号
            const quoteCount = (fixed.match(/"/g) || []).length;
            if (quoteCount % 2 !== 0) fixed += '"';
            // 修复未闭合的 {
            const openBrace = (fixed.match(/\{/g) || []).length;
            const closeBrace = (fixed.match(/\}/g) || []).length;
            for (let i = closeBrace; i < openBrace; i++) fixed += '}';
            // 修复未闭合的 [
            const openBracket = (fixed.match(/\[/g) || []).length;
            const closeBracket = (fixed.match(/\]/g) || []).length;
            for (let i = closeBracket; i < openBracket; i++) fixed += ']';
            return fixed;
        }

        /**
         * 从 content 中提取嵌入的 thinking 标签内容
         * 某些模型把思考嵌在 <thinking>...\n</thinking> 标签里
         */
        _extractThinkingFromContent(text) {
            if (!text || typeof text !== 'string') return { content: text, reasoning: '' };

            const tags = [
                { open: '<thinking>',  close: '</thinking>' },
                { open: '<think>',     close: '</think>' },
                { open: '[thinking]',  close: '[/thinking]' },
                { open: '[think]',     close: '[/think]' },
                { open: '<antThinking>', close: '</antThinking>' }
            ];

            let reasoning = '';
            let remaining = text;
            let changed = true;

            while (changed) {
                changed = false;
                for (const tag of tags) {
                    const openIdx = remaining.indexOf(tag.open);
                    if (openIdx !== -1) {
                        const closeSearchStart = openIdx + tag.open.length;
                        const closeIdx = remaining.indexOf(tag.close, closeSearchStart);
                        if (closeIdx !== -1) {
                            const thinkingPart = remaining.substring(openIdx + tag.open.length, closeIdx);
                            reasoning += (reasoning ? '\n' : '') + thinkingPart.trim();
                            remaining = remaining.substring(0, openIdx) + remaining.substring(closeIdx + tag.close.length);
                            changed = true;
                            break;
                        }
                    }
                }
            }

            return { content: remaining.trim(), reasoning: reasoning.trim() };
        }

        /**
         * 仅剥离 content 中的 thinking 标签，不提取 reasoning
         * 用于已有独立 reasoning_content 时，避免思考内容重复显示
         */
        _stripThinkingTags(text) {
            if (!text || typeof text !== 'string') return text;

            const tags = [
                { open: '<thinking>',  close: '</thinking>' },
                { open: '<think>',     close: '</think>' },
                { open: '[thinking]',  close: '[/thinking]' },
                { open: '[think]',     close: '[/think]' },
                { open: '<antThinking>', close: '</antThinking>' }
            ];

            let remaining = text;
            let changed = true;
            while (changed) {
                changed = false;
                for (const tag of tags) {
                    const openIdx = remaining.indexOf(tag.open);
                    if (openIdx !== -1) {
                        const closeSearchStart = openIdx + tag.open.length;
                        const closeIdx = remaining.indexOf(tag.close, closeSearchStart);
                        if (closeIdx !== -1) {
                            remaining = remaining.substring(0, openIdx) + remaining.substring(closeIdx + tag.close.length);
                            changed = true;
                            break;
                        }
                    }
                }
            }
            return remaining.trim();
        }

        /**
         * 从已解析的 JSON 中提取 content/reasoning/model
         * 兼容多家 API 格式
         */
        _extractFromParsed(parsed) {
            let content = '';
            let reasoning = '';
            let toolCalls = null;
            let isDone = false;
            let isError = false;
            let errorMessage = '';
            let model = '';
            let finishReason = '';

            if (!parsed || typeof parsed !== 'object') {
                return { content, reasoning, isDone, isError, errorMessage };
            }

            // 错误响应
            if (parsed.error) {
                return {
                    content: '',
                    reasoning: '',
                    isDone: true,
                    isError: true,
                    errorMessage: typeof parsed.error === 'string' ? parsed.error : (parsed.error.message || JSON.stringify(parsed.error)),
                    finishReason: 'error'
                };
            }

            // 提取 model 信息（所有 API 格式都可能包含）
            if (parsed.model) {
                model = String(parsed.model);
            } else if (parsed.message?.model) {
                model = String(parsed.message.model);
            }

            // OpenAI 兼容格式: { choices: [{ delta: { content, reasoning_content }, finish_reason }] }
            if (Array.isArray(parsed.choices) && parsed.choices.length > 0) {
                const choice = parsed.choices[0];
                const delta = choice.delta || choice.message || {};

                if (delta.content !== undefined && delta.content !== null) {
                    content = String(delta.content);
                }
                // Kimi/Moonshot/OpenAI o1 等使用 reasoning_content 字段
                if (delta.reasoning_content !== undefined) {
                    reasoning = String(delta.reasoning_content);
                } else if (delta.reasoning !== undefined) {
                    reasoning = String(delta.reasoning);  // 部分 API
                } else if (delta.thinking !== undefined) {
                    reasoning = String(delta.thinking);  // Anthropic 风格
                }

                // 处理 content 中的 thinking 标签：
                // - 若已有独立 reasoning_content，仅剥离 content 中的 thinking 标签，避免重复显示
                // - 若没有独立 reasoning_content，尝试从 content 提取 thinking
                if (reasoning && content) {
                    content = this._stripThinkingTags(content);
                } else if (!reasoning && content) {
                    const extracted = this._extractThinkingFromContent(content);
                    if (extracted.reasoning) {
                        reasoning = extracted.reasoning;
                        content = extracted.content;
                    }
                }

                // 工具调用（function calling）—— OpenAI/Moonshot/Kimi 流式增量格式
                // 【致命修复】tool_calls 是增量传输的：每个 chunk 可能只带 index:0 的 arguments 追加
                // 必须按 index 累积 name + arguments，不能直接用最后一个 chunk 覆盖！
                if (delta.tool_calls || delta.toolCalls) {
                    const incoming = delta.tool_calls || delta.toolCalls;
                    if (Array.isArray(incoming) && incoming.length > 0) {
                        if (!this._accumulatedToolCalls) this._accumulatedToolCalls = [];
                        for (const inc of incoming) {
                            const idx = typeof inc.index === 'number' ? inc.index : this._accumulatedToolCalls.length;
                            // 确保数组足够长
                            while (this._accumulatedToolCalls.length <= idx) {
                                this._accumulatedToolCalls.push({ index: this._accumulatedToolCalls.length, id: '', type: 'function', function: { name: '', arguments: '' } });
                            }
                            const existing = this._accumulatedToolCalls[idx];
                            // id: 首次出现才设（后续 chunk 可能不重复）
                            if (inc.id && !existing.id) existing.id = inc.id;
                            // name: 首次出现才设（后续增量 chunk 通常不再带 name）
                            const incFn = inc.function || {};
                            if (incFn.name && !existing.function.name) existing.function.name = incFn.name;
                            // arguments: 增量追加（流式传输是字符串片段拼接）
                            if (typeof incFn.arguments === 'string' && incFn.arguments.length > 0) {
                                existing.function.arguments += incFn.arguments;
                            }
                        }
                        // 始终返回累积好的完整数组（去掉尾部空 name 的未完成项）
                        toolCalls = this._accumulatedToolCalls.filter(tc => tc && tc.function && tc.function.name);
                    }
                }

                if (choice.finish_reason === 'stop' || choice.finish_reason === 'length' ||
                    choice.finish_reason === 'tool_calls' || choice.finish_reason === 'function_call') {
                    isDone = true;
                    finishReason = choice.finish_reason;
                }
            }
            // Ollama 格式: { model, message: { content, thinking, tool_calls }, done }
            else if (parsed.message && typeof parsed.message === 'object') {
                if (parsed.message.content !== undefined) {
                    content = String(parsed.message.content);
                }
                if (parsed.message.thinking !== undefined) {
                    reasoning = String(parsed.message.thinking);
                } else if (parsed.message.reasoning !== undefined) {
                    reasoning = String(parsed.message.reasoning);
                }
                // 如果没有独立的 thinking 字段，尝试从 content 中提取
                if (!reasoning && content) {
                    const extracted = this._extractThinkingFromContent(content);
                    if (extracted.reasoning) {
                        reasoning = extracted.reasoning;
                        content = extracted.content;
                    }
                }
                // Ollama tool_calls 格式: [{ function: { name, arguments } }]
                if (parsed.message.tool_calls && Array.isArray(parsed.message.tool_calls)) {
                    toolCalls = parsed.message.tool_calls.map(tc => ({
                        id: tc.name ? `ollama_${tc.name}` : undefined,
                        type: 'function',
                        function: {
                            name: tc.function?.name || tc.name || '',
                            arguments: tc.function?.arguments
                                ? (typeof tc.function.arguments === 'string'
                                    ? tc.function.arguments
                                    : JSON.stringify(tc.function.arguments))
                                : '{}'
                        }
                    }));
                }
                if (parsed.done === true) isDone = true;
            }
            // 直接 content 字段（部分中转 API）
            else if (parsed.content !== undefined) {
                content = String(parsed.content);
                if (parsed.thinking !== undefined) reasoning = String(parsed.thinking);
                else if (parsed.reasoning !== undefined) reasoning = String(parsed.reasoning);
                // 从 content 提取嵌入的 thinking
                if (!reasoning && content) {
                    const extracted = this._extractThinkingFromContent(content);
                    if (extracted.reasoning) {
                        reasoning = extracted.reasoning;
                        content = extracted.content;
                    }
                }
                if (parsed.done === true) isDone = true;
            }
            // Anthropic 原生格式: { type: 'content_block_delta', delta: { type, text } }
            else if (parsed.type === 'content_block_delta' && parsed.delta) {
                if (parsed.delta.type === 'text_delta' && parsed.delta.text) {
                    content = parsed.delta.text;
                } else if (parsed.delta.type === 'thinking_delta' && parsed.delta.thinking) {
                    reasoning = parsed.delta.thinking;
                }
            } else if (parsed.type === 'message_stop') {
                isDone = true;
            }
            // 通义/智谱等使用 response 字段
            else if (parsed.response !== undefined) {
                content = String(parsed.response);
            }

            const result = { content, reasoning, isDone, isError, errorMessage };
            if (toolCalls) result.toolCalls = toolCalls;
            if (model) result.model = model;
            if (finishReason) result.finishReason = finishReason;
            // 【Token 用量】OpenAI 兼容(parsed.usage: prompt_tokens/completion_tokens)、
            // Anthropic(parsed.usage: input_tokens/output_tokens)、
            // Ollama 原生(prompt_eval_count/eval_count，done 帧携带)
            let usage = null;
            if (parsed.usage && typeof parsed.usage === 'object') {
                const _in = Number(parsed.usage.prompt_tokens ?? parsed.usage.input_tokens) || 0;
                const _out = Number(parsed.usage.completion_tokens ?? parsed.usage.output_tokens) || 0;
                if (_in || _out) usage = { in: _in, out: _out };
            } else if (parsed.eval_count !== undefined || parsed.prompt_eval_count !== undefined) {
                const _in = Number(parsed.prompt_eval_count) || 0;
                const _out = Number(parsed.eval_count) || 0;
                if (_in || _out) usage = { in: _in, out: _out };
            }
            if (usage) result.usage = usage;
            return result;
        }

        // ===== 状态查询 =====
        isDone() { return this._isDone; }
        isError() { return this._isError; }
        getErrorMessage() { return this._errorMessage; }
        getBuffer() { return this._buffer; }

        reset() {
            this._buffer = '';
            this._isDone = false;
            this._isError = false;
            this._errorMessage = '';
            this._accumulatedToolCalls = null;  // 清空 tool_calls 累积缓存
        }
    }

    App.ChatV2.StreamProcessor = StreamProcessor;
    ChatLog.info?.('CHAT-V2', '✅ StreamProcessor loaded');
})();