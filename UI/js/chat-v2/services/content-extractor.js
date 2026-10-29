/**
 * ============================================================
 * chat-v2/services/content-extractor.js — 思考/内容分离器
 * ============================================================
 * 职责：从模型输出中分离 <thinking> 标签包裹的思考过程与正文内容
 * 设计：状态机驱动，支持流式输入；单一来源，不在其他模块重复实现
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    const STATES = { IDLE: 0, IN_THINKING: 1, IN_CONTENT: 2 };

    // 支持的 thinking 标签格式（按优先级排序）
    const THINKING_TAGS = [
        { open: '<thinking>',  close: '</thinking>' },
        { open: '<think>',     close: '</think>' },          // DeepSeek-R1 风格
        { open: '[thinking]',  close: '[/thinking]' },
        { open: '[think]',     close: '[/think]' },
        { open: 'antThinking>',     close: '</antThinking>' }   // Anthropic 风格
    ];

    const MAX_TAG_LEN = Math.max(...THINKING_TAGS.map(t => t.open.length));

    class ContentExtractor {
        constructor() {
            this.reset();
        }

        reset() {
            this._state = STATES.IDLE;
            this._thinkingBuf = '';
            this._contentBuf = '';
            this._pendingTag = null;
            this._tagMatchBuf = '';
            this._closeMatchIdx = 0;
        }

        /**
         * 喂入文本片段（流式）
         * @param {string} text
         */
        feed(text) {
            if (!text || typeof text !== 'string') return;
            for (let i = 0; i < text.length; i++) {
                this._feedChar(text[i]);
            }
        }

        _feedChar(char) {
            switch (this._state) {
                case STATES.IDLE: this._handleIdle(char); break;
                case STATES.IN_THINKING: this._handleInThinking(char); break;
                case STATES.IN_CONTENT: this._handleInContent(char); break;
            }
        }

        _handleIdle(char) {
            this._tagMatchBuf += char;
            const matched = this._matchOpenTag(this._tagMatchBuf);
            if (matched) {
                this._state = STATES.IN_THINKING;
                this._pendingTag = matched;
                this._tagMatchBuf = '';
                this._closeMatchIdx = 0;
                return;
            }
            // 缓冲区超过最长标签长度仍未匹配，写入内容区
            if (this._tagMatchBuf.length > MAX_TAG_LEN) {
                this._contentBuf += this._tagMatchBuf;
                this._tagMatchBuf = '';
            }
        }

        _handleInThinking(char) {
            const closeTag = this._pendingTag.close;
            if (char === closeTag[this._closeMatchIdx]) {
                this._closeMatchIdx++;
                if (this._closeMatchIdx === closeTag.length) {
                    // 标签结束，进入内容状态
                    this._state = STATES.IN_CONTENT;
                    this._pendingTag = null;
                    this._closeMatchIdx = 0;
                    return;
                }
                // 还在匹配关闭标签中，不写入
                return;
            }
            // 不匹配，把已匹配的部分写回思考内容
            if (this._closeMatchIdx > 0) {
                this._thinkingBuf += closeTag.substring(0, this._closeMatchIdx);
                this._closeMatchIdx = 0;
            }
            this._thinkingBuf += char;
        }

        _handleInContent(char) {
            // 内容状态下也可能再次出现 thinking 标签（多段思考）
            this._tagMatchBuf += char;
            const matched = this._matchOpenTag(this._tagMatchBuf);
            if (matched) {
                this._state = STATES.IN_THINKING;
                this._pendingTag = matched;
                this._tagMatchBuf = '';
                this._closeMatchIdx = 0;
                return;
            }
            if (this._tagMatchBuf.length > MAX_TAG_LEN) {
                this._contentBuf += this._tagMatchBuf;
                this._tagMatchBuf = '';
            }
        }

        _matchOpenTag(text) {
            for (const tag of THINKING_TAGS) {
                if (text.endsWith(tag.open)) return tag;
            }
            return null;
        }

        // ===== 输出 =====
        getThinking() {
            return (this._thinkingBuf + (this._state === STATES.IN_THINKING ? this._tagMatchBuf : '')).trim();
        }

        getContent() {
            return (this._contentBuf + (this._state !== STATES.IN_THINKING ? this._tagMatchBuf : '')).trim();
        }

        getFullContent() {
            return this.getThinking() + '\n' + this.getContent();
        }

        isInThinking() { return this._state === STATES.IN_THINKING; }
        getState() { return this._state; }

        // ===== 静态方法：一次性提取（非流式场景） =====
        static extract(text) {
            if (!text) return { thinking: '', content: '' };
            const extractor = new ContentExtractor();
            extractor.feed(text);
            return {
                thinking: extractor.getThinking(),
                content: extractor.getContent()
            };
        }

        /**
         * 清理思考内容中的模板化废话
         * 用于过滤模型把 prompt 模板回显到思考中的情况
         */
        static cleanThinking(text) {
            if (!text) return '';
            let cleaned = text;
            const patterns = [
                /^\s*【[^】]+】[\s\S]*?(?=\n|$)/gm,    // 【...】开头的指令块
                /^\s*请.*?按照.*?格式.*?输出/gm,
                /^\s*你的回复必须包含/gm,
                /^\s*第一部分：/gm,
                /^\s*第二部分：/gm,
                /^\s*禁止输出/gm,
                /^\s*请立即开始输出/gm
            ];
            patterns.forEach(p => { cleaned = cleaned.replace(p, ''); });
            cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
            return cleaned.trim();
        }
    }

    App.ChatV2.ContentExtractor = ContentExtractor;
    ChatLog.info?.('CHAT-V2', '✅ ContentExtractor loaded');
})();
