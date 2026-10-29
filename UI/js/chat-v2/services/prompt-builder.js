/**
 * ============================================================
 * chat-v2/services/prompt-builder.js — 系统提示词组装器
 * ============================================================
 * 职责：PromptBuilder 类管理有序 section 注册 + 变量插值 + 最终渲染
 * 设计：**无硬码** — 本文件不含任何 prompt 内容字符串
 *       所有 prompt 内容来自配置（App.DEFAULT_PROMPTS / prompts.md / App.promptConfig）
 * 架构：借鉴 dsh-system-prompt 的 PromptAssembly 设计
 *       - addSection(name, text, order): 按序注册 section（同名替换）
 *       - setVariable(key, value): 注册变量（支持函数延迟求值）
 *       - assemble(): 合并所有 section，解析动态文本，按 order 排序
 *       - render(): 变量插值 + 空 section 过滤 + 最终拼接
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // ── 宿主感知（Excel / Word / PowerPoint）─────────────────────
    // Excel 宿主下所有分支保持原行为不变（isExcel 为 true 时走原逻辑）。
    const EXCEL_HOST = { kind: 'excel', label: 'Excel', isExcel: true, isWord: false, isPowerPoint: false };

    function getHost() {
        try {
            const h = window.SmartPilotHost;
            if (h && typeof h === 'object') {
                return {
                    kind: h.kind || 'excel',
                    label: h.label || 'Excel',
                    isExcel: h.isExcel !== false,
                    isWord: !!h.isWord,
                    isPowerPoint: !!h.isPowerPoint
                };
            }
        } catch (e) { /* 解析失败按 Excel 处理 */ }
        return EXCEL_HOST;
    }

    /**
     * 非 Excel 宿主的"运行环境"提示：明确当前宿主，禁止套用 Excel（工作表/选区/工作簿）心智。
     * @param {{kind:string,label:string,isWord:boolean,isPowerPoint:boolean}} host
     * @returns {string}
     */
    function buildHostEnvPrompt(host) {
        if (host.isWord) {
            return '【运行环境】你运行在 Microsoft Word 中，是一个 Word 智能助手。\n' +
                '1. 用户操作的对象是当前打开的 Word 文档，不存在 Excel 的"工作表 / 工作簿 / 单元格选区"等概念；禁止要求用户"去 Excel 中选择单元格或数据区域"。\n' +
                '2. 读取文档内容用 word_* 工具（如 word_get_document_text）；写入、排版、表格、替换、页眉页脚等也使用对应的 word_* 工具。\n' +
                '3. 不要再引用任何 Excel（excel_*）工具或说明，直接基于文档内容完成用户请求。\n' +
                '4. 全程使用与用户提问相同的语言，输出使用 Markdown 排版。\n' +
                '5. 严格基于真实内容作答：读取类工具返回为空、或没有读到任何内容时，必须如实说明"没有读取到内容"，严禁编造、猜测或凭标题虚构正文。';
        }
        return '【运行环境】你运行在 Microsoft PowerPoint 中，是一个 PowerPoint 智能助手。\n' +
            '1. 用户操作的对象是当前打开的演示文稿，不存在 Excel 的"工作表 / 工作簿 / 单元格选区"等概念；禁止要求用户"去 Excel 中选择单元格或数据区域"。\n' +
            '2. 读取幻灯片内容用 ppt_* 工具（如 ppt_get_slide_text / ppt_get_all_slides_text）；新增或编辑幻灯片、文本框、图表、备注等也使用对应的 ppt_* 工具。\n' +
            '3. 不要再引用任何 Excel（excel_*）工具或说明，直接基于演示文稿内容完成用户请求。\n' +
            '4. 全程使用与用户提问相同的语言，输出使用 Markdown 排版。\n' +
            '5. 严格基于真实内容作答：读取类工具返回为空、或没有读到任何内容时，必须如实说明"没有读取到内容"，严禁编造、猜测或凭标题虚构正文。';
    }

    // ── 内容处理任务（总结 / 润色 / 改写 / 翻译 / 提炼要点）───────
    // 非 Excel 宿主缺失这一层约束时的实测故障：用户说"通读全部幻灯片文本并润色表述"，
    // 模型调 ppt_get_all_slides_text 读完全文后，总结轮只回了一句
    // "我已经读取了全部幻灯片的文本。接下来，请提供您希望对表述进行的具体润色建议…"
    // 就把任务丢回给用户，任务卡却显示"任务完成"。
    // 根因：Excel 有 ANALYSIS_TASK_PROMPT 强制"读完直接给结论"，而 Word/PowerPoint
    // 只注入了 buildHostEnvPrompt（讲清宿主与工具族），没有任何"必须交付成品、禁止反问"的约束。
    const CONTENT_TASK_INTENT_RE = /总结|概括|归纳|摘要|要点|大纲|提纲|提炼|精简|扩写|润色|措辞|语气|改写|重写|优化表述|优化文字|通读|梳理|翻译|校对|纠错|错别字/;

    /**
     * 判断用户消息是否属于"内容处理"意图（总结/润色/改写/翻译/提炼）。
     * @param {string} text - 用户原始消息
     * @returns {boolean}
     */
    function detectContentTaskIntent(text) {
        if (!text || typeof text !== 'string') return false;
        return CONTENT_TASK_INTENT_RE.test(text);
    }

    /**
     * 非 Excel 宿主的"内容处理"任务提示：强制直接交付成品，禁止把任务反问回用户。
     * @param {{isPowerPoint:boolean}} host
     * @returns {string}
     */
    function buildHostContentTaskPrompt(host, rawUser) {
        const isPpt = !!host.isPowerPoint;
        const readTool = isPpt ? 'ppt_get_all_slides_text' : 'word_get_document_text';
        const readToolCurrent = isPpt ? 'ppt_get_slide_text' : '';
        const readScope = isPpt ? '按页全文' : '全文';
        const unit = isPpt ? '页' : '段';
        const targetName = isPpt ? '演示文稿' : '文档';
        const writeHint = isPpt
            ? 'ppt_set_shape_text（按形状改写）/ ppt_find_replace（批量替换）'
            : 'word_* 的写入 / 替换工具';
        // 【范围感知】用户说"当前幻灯片/这一页/本页"时只处理当前页，禁止罗列全部页
        // （实测弱模型把"总结当前幻灯片"做成 11 页逐页目录抄写）。
        const wantCurrent = /当前|这一页|本页|此页|现在这/.test(String(rawUser || ''));
        const scopeRule = wantCurrent
            ? `用户要处理的是【当前${isPpt ? '一页幻灯片' : '文档'}】：优先直接基于上下文已注入的"${isPpt ? '幻灯片' : '文档'}内容"作答，不要再调读取工具；` +
              `若上下文标注为"全部幻灯片文本"，说明当前页定位失败，此时给出全文的精炼整体总结即可。\n` +
              `   ❌ 严禁把全部${unit}逐${unit}罗列（"幻灯片1/幻灯片2…"清单式抄写）来冒充当前${unit}的内容。`
            : `用户未限定页码：可覆盖全部内容，但输出必须是跨${unit}提炼的归纳结论，不是逐${unit}目录。`;
        return `【任务类型：内容处理（总结 / 润色 / 改写 / 翻译 / 提炼要点）】\n` +
            `用户要对当前${targetName}的内容做处理（${isPpt ? '幻灯片' : '文档'}）：\n` +
            `1. ✅ 必须【直接交付成品】：读完内容后立刻给出完整结果，不要只汇报"我已经读取了…"。\n` +
            `2. 【处理范围】${scopeRule}\n` +
            `3. ❌ 禁止复述原文：总结/提炼只输出提炼后的要点与结论（每条一行、不超过 6 条）；\n` +
            `   只有润色/改写任务才给"${isPpt ? '原文' : '原句'} → 优化后"对照。\n` +
            `4. ❌ 禁止反问收尾：回复的最后一句话必须是任务成果本身（结论/结果），\n` +
            `   绝不以"请告诉我您希望…/是否需要…/您可以回复…/请详细说明您的需求"等结尾。\n` +
            `5. 内容读取一次即可：${readTool}（缺省返回${readScope}${readToolCurrent ? `；只读当前${unit}用 ${readToolCurrent}（不传参=当前${unit}）` : ''}），不要重复调用同一个读取工具。\n` +
            `6. 保持原意：数字、专有名词、专业术语一律不改动，只优化表达（更简洁、更专业、逻辑更顺）。\n` +
            `7. 默认只在对话里给出结果；仅当用户明确要求"直接改到${isPpt ? '幻灯片' : '文档'}里 / 替换原文"时，\n` +
            `   才调用写回工具（${writeHint}），并在完成后报告改了哪些${unit}。\n` +
            `8. 全程使用与用户提问相同的语言，输出使用 Markdown 排版。`;
    }

    /**
     * 非 Excel 宿主（PPT / Word）的"文件产出"任务提示。
     * 与 Excel 的 FILE_OUTPUT_TASK_PROMPT 对齐：告诉模型直接调导出工具、filePath 留空即可，
     * 严禁编造占位符路径（如 C:\Users\Yourusername\Documents\report.pdf）。
     */
    function buildHostFileOutputTaskPrompt(host) {
        const isPpt = !!host.isPowerPoint;
        const targetName = isPpt ? '演示文稿' : '文档';
        const exportTool = isPpt ? 'ppt_export_pdf' : 'word_export_pdf';
        const saveTool = isPpt ? 'ppt_save' : 'word_save';
        return `【任务类型：文件产出（导出 / 另存为 / 保存到文件）】\n` +
            `用户要把当前${targetName}导出或保存为文件。务必遵循：\n` +
            `1. 直接调用对应工具完成落盘：导出 PDF → ${exportTool}；保存 → ${saveTool}。\n` +
            `1.1 导出/保存是"只读落盘"动作，与内容、格式无关：空白${targetName}同样可以导出。\n` +
            `    严禁在此之前插入任何多余步骤——不要先做格式美化（如调 ${isPpt ? 'ppt' : 'word'}_format_selected_table_cell）、\n` +
            `    不要先改文字/插图/换主题，也不要先读取全文。一步调用即完成。\n` +
            `2. filePath 可以留空：系统会自动保存到真实桌面（文件名 = ${targetName}名_时间戳.pdf）。\n` +
            `3. 严禁编造路径（如 C:\\Users\\Yourusername\\Documents\\report.pdf）——这类占位符路径会被系统拒绝或重定向。\n` +
            `4. 完成后用工具返回的真实路径报告结果；不要在没有工具调用的情况下声称"已完成"。\n` +
            `5. 全程使用与用户提问相同的语言。`;
    }

    // ── PromptBuilder 类 ────────────────────────────────────────

    /**
     * PromptBuilder — 有序 section 注册 + 变量插值 + 最终渲染
     *
     * 用法:
     *   const pb = new PromptBuilder()
     *     .addSection('identity', '你是SmartPilot', -100)
     *     .addSection('rules', () => buildRules(), 50)
     *     .addSection('tools', '可用工具:\n{toolInfo}', 200)
     *     .setVariable('toolInfo', () => collectTools())
     *   const prompt = pb.render()
     *
     * 设计原则（借鉴 dsh-system-prompt）:
     *   - section.name 唯一（同名注册替换旧 section）
     *   - section.order 决定渲染顺序（小数字在前）
     *   - section.text 支持静态字符串或延迟求值函数
     *   - 变量通过 setVariable 注册，render 阶段统一插值
     *   - 空 section 自动过滤
     */
    class PromptBuilder {
        constructor() {
            /** @type {Array<{name: string, text: string|Function, order: number}>} */
            this._sections = [];
            /** @type {Object<string, string|Function>} */
            this._variables = {};
            /** @type {boolean} */
            this._modified = true;
            /** @type {Array<{name: string, text: string, order: number}>|null} */
            this._assembled = null;
            /** @type {Array<Function>} assemble 瀑布事件监听器 */
            this._assembleListeners = [];
        }

        /**
         * 设置助手身份（persona），替换默认的 identity section（order=-100）。
         * 支持链式调用。
         * @param {string} text - 身份描述，如 "你是数据分析专家"
         * @returns {PromptBuilder} 链式调用
         */
        setPersona(text) {
            if (!text || typeof text !== 'string') return this;
            return this.addSection('identity', text, -100);
        }

        /**
         * 注册 assemble 瀑布事件监听器。
         * 监听器在 assemble() 阶段被调用，可修改 sections 或 variables。
         * 借鉴 dsh-system-prompt 的 system-prompt/assemble 事件设计。
         * @param {Function} listener - (assembly) => void，assembly = { sections, variables }
         * @returns {Function} 取消注册的 disposer 函数
         */
        onAssemble(listener) {
            if (typeof listener !== 'function') return () => {};
            this._assembleListeners.push(listener);
            return () => {
                const idx = this._assembleListeners.indexOf(listener);
                if (idx >= 0) this._assembleListeners.splice(idx, 1);
            };
        }

        /**
         * 注册一个有序 section。
         * 同名 section 会替换旧值（保持原 order），实现"覆盖式"组合。
         * @param {string} name - 唯一 section 名（如 'identity', 'rules', 'tools'）
         * @param {string|Function} text - 静态文本或返回字符串的函数
         * @param {number} [order=100] - 排序权重（小数字在前）。约定:
         *   -100: 助手身份（"你是SmartPilot"）
         *   0:    核心规则
         *   100:  工具/Skill 列表
         *   200:  路由规则
         *   300:  安全约束
         * @returns {PromptBuilder} 链式调用
         */
        addSection(name, text, order = 100) {
            if (!name || typeof name !== 'string') {
                ChatLog.warn?.('PROMPT', `[PromptBuilder] addSection 跳过: 无效 name`);
                return this;
            }
            // 同名替换：保留原 order 或使用新 order
            const existing = this._sections.find(s => s.name === name);
            if (existing) {
                existing.text = text;
                if (order !== 100) existing.order = order;
                ChatLog.debug?.('PROMPT', `[PromptBuilder] 替换 section: ${name} (order=${existing.order})`);
            } else {
                this._sections.push({ name, text, order });
                ChatLog.debug?.('PROMPT', `[PromptBuilder] 注册 section: ${name} (order=${order})`);
            }
            this._sections.sort((a, b) => a.order - b.order);
            this._modified = true;
            return this;
        }

        /**
         * 移除指定 section。
         * @param {string} name - section 名
         * @returns {PromptBuilder} 链式调用
         */
        removeSection(name) {
            if (!name) return this;
            const len = this._sections.length;
            this._sections = this._sections.filter(s => s.name !== name);
            if (this._sections.length !== len) {
                this._modified = true;
                ChatLog.debug?.('PROMPT', `[PromptBuilder] 移除 section: ${name}`);
            }
            return this;
        }

        /**
         * 注册变量（用于 render 阶段的 {{variable}} 插值）。
         * @param {string} key - 变量名
         * @param {string|Function} value - 静态值或返回字符串的函数
         * @returns {PromptBuilder} 链式调用
         */
        setVariable(key, value) {
            if (!key) return this;
            this._variables[key] = value;
            return this;
        }

        /**
         * 批量注册变量。
         * @param {Object<string, string|Function>} vars - 变量映射
         * @returns {PromptBuilder} 链式调用
         */
        setVariables(vars) {
            if (vars && typeof vars === 'object') {
                for (const [k, v] of Object.entries(vars)) {
                    this.setVariable(k, v);
                }
            }
            return this;
        }

        /**
         * 获取所有已注册的 section 名列表。
         * @returns {string[]}
         */
        listSections() {
            return this._sections.map(s => s.name);
        }

        /**
         * 获取指定 section 的文本（已解析）。
         * @param {string} name - section 名
         * @returns {string|null}
         */
        getSectionText(name) {
            const s = this._sections.find(s => s.name === name);
            if (!s) return null;
            return typeof s.text === 'function' ? s.text() : s.text;
        }

        /**
         * 组装所有 section：解析动态文本，按 order 排序，过滤空 section。
         * 触发瀑布事件（onAssemble 监听器），允许外部修改 assembly。
         * @returns {Array<{name: string, text: string, order: number}>}
         */
        assemble() {
            if (!this._modified && this._assembled) return this._assembled;

            // 1. 解析所有 section 文本
            const resolved = this._sections.map(s => {
                let text;
                try {
                    text = typeof s.text === 'function' ? s.text() : s.text;
                } catch (e) {
                    ChatLog.warn?.('PROMPT', `[PromptBuilder] section ${s.name} 求值异常:`, e);
                    text = '';
                }
                return { name: s.name, text: String(text ?? ''), order: s.order };
            });

            // 2. 触发瀑布事件（waterfall）：允许监听器修改 sections
            if (this._assembleListeners.length > 0) {
                const assembly = {
                    sections: resolved,
                    variables: this._variables
                };
                for (const listener of this._assembleListeners) {
                    try {
                        listener(assembly);
                    } catch (e) {
                        ChatLog.warn?.('PROMPT', `[PromptBuilder] assemble 监听器异常:`, e);
                    }
                }
                // 监听器可能修改了 sections
                this._sections = assembly.sections.map(s => ({
                    name: s.name,
                    text: s.text,
                    order: s.order
                }));
                this._sections.sort((a, b) => a.order - b.order);
            }

            // 3. 过滤空 section
            this._assembled = resolved.filter(s => s.text.trim());
            this._modified = false;
            return this._assembled;
        }

        /**
         * 渲染最终系统提示词。
         * 流程: assemble() → 变量插值 → 空 section 过滤 → 换行拼接
         * @returns {string} 渲染后的系统提示词
         */
        render() {
            const sections = this.assemble();
            if (sections.length === 0) return '';

            // 先拼接，再统一插值
            let prompt = sections.map(s => s.text).join('\n\n');

            // 变量插值：替换 {{variable}} 或 {variable} 格式
            for (const [key, value] of Object.entries(this._variables)) {
                let resolved;
                try {
                    resolved = typeof value === 'function' ? value() : value;
                } catch (e) {
                    ChatLog.warn?.('PROMPT', `[PromptBuilder] 变量 ${key} 求值异常:`, e);
                    resolved = '';
                }
                const strVal = String(resolved ?? '');
                // 支持 {{key}} 和 {key} 两种格式
                prompt = prompt.replace(new RegExp(`\\{\\{${key}\\}\\}`, 'g'), strVal);
                prompt = prompt.replace(new RegExp(`\\{${key}\\}`, 'g'), strVal);
            }

            return prompt.trim();
        }

        /**
         * 重置所有 section 和变量。
         * @returns {PromptBuilder} 链式调用
         */
        clear() {
            this._sections = [];
            this._variables = {};
            this._assembled = null;
            this._modified = true;
            this._assembleListeners = [];
            return this;
        }

        /**
         * 从现有 PromptBuilder 复制所有 section 和变量。
         * @param {PromptBuilder} other - 源 builder
         * @returns {PromptBuilder} 链式调用
         */
        copyFrom(other) {
            if (!(other instanceof PromptBuilder)) return this;
            this._sections = other._sections.map(s => ({ ...s }));
            this._variables = { ...other._variables };
            this._modified = true;
            return this;
        }
    }

    // ── 全局默认实例 ─────────────────────────────────────────────

    /** @type {PromptBuilder} 全局默认 PromptBuilder 实例 */
    let _defaultBuilder = null;

    /**
     * 获取或创建全局默认 PromptBuilder 实例。
     * @returns {PromptBuilder}
     */
    function getDefaultBuilder() {
        if (!_defaultBuilder) {
            _defaultBuilder = new PromptBuilder();
        }
        return _defaultBuilder;
    }

    /**
     * 重置全局默认实例（用于热重载）。
     */
    function resetDefaultBuilder() {
        _defaultBuilder = null;
    }

    // ── Token 估算工具（保留，与现有逻辑兼容）────────────────────

    function estimateTokens(text) {
        if (!text) return 0;
        const str = typeof text === 'string' ? text : JSON.stringify(text);
        const chineseChars = (str.match(/[\u4e00-\u9fff]/g) || []).length;
        const totalChars = str.length;
        const chineseRatio = totalChars > 0 ? chineseChars / totalChars : 0;
        const tokens = chineseChars * 0.5 + (totalChars - chineseChars) * 0.25;
        return Math.ceil(tokens);
    }

    function getContextWindow() {
        try {
            const config = App.config || {};
            const apiType = config.apiType || 'ollama';
            const model = config.model || '';
            const providerConfig = App.PROVIDER_CONFIGS?.[apiType];
            if (!providerConfig) return 8192;
            if (providerConfig.models) {
                for (const m of providerConfig.models) {
                    if (typeof m === 'object' && m.name === model && m.contextWindow) {
                        return m.contextWindow;
                    }
                }
            }
            if (providerConfig.contextWindow) return providerConfig.contextWindow;
            return 8192;
        } catch (e) {
            return 8192;
        }
    }

    function applyTokenBudget(messages, contextWindow, toolTokens = 0) {
        const availableForMessages = Math.floor(contextWindow * 0.5);
        const safetyMargin = Math.floor(contextWindow * 0.05);
        const hardLimit = Math.max(1000, availableForMessages - safetyMargin - toolTokens);

        if (toolTokens > 0) {
            ChatLog.info?.('PROMPT', `[prompt-builder] tool tokens: ${toolTokens}, message budget: ${hardLimit}`);
        }

        let totalTokens = 0;
        const messageTokens = messages.map(m => {
            const content = typeof m.content === 'string' ? m.content : JSON.stringify(m.content);
            const tokens = estimateTokens(content);
            totalTokens += tokens;
            return tokens;
        });

        if (totalTokens <= hardLimit) {
            return { messages, truncated: false, totalTokens };
        }

        ChatLog.info?.('PROMPT', `[prompt-builder] token overflow: ${totalTokens}/${hardLimit}, applying truncation`);

        let systemIdx = -1;
        let lastUserIdx = -1;
        for (let i = messages.length - 1; i >= 0; i--) {
            if (messages[i].role === 'user' && lastUserIdx === -1) lastUserIdx = i;
            if (messages[i].role === 'system' && systemIdx === -1) { systemIdx = i; break; }
        }

        const newMessages = [...messages];
        let historyStart = systemIdx + 1;
        let historyEnd = lastUserIdx - 1;

        while (totalTokens > hardLimit && historyStart <= historyEnd && historyStart >= 0 && historyEnd < newMessages.length) {
            if (historyStart < historyEnd) {
                const removed = newMessages.splice(historyStart, 1)[0];
                if (!removed) break;
                totalTokens -= estimateTokens(typeof removed.content === 'string' ? removed.content : JSON.stringify(removed.content));
                historyEnd--;
                if (lastUserIdx > historyStart) lastUserIdx--;
                ChatLog.info?.('PROMPT', `[prompt-builder] removed history, remaining tokens: ${totalTokens}`);
            } else {
                break;
            }
        }

        systemIdx = newMessages.findIndex(m => m?.role === 'system');
        lastUserIdx = -1;
        for (let i = newMessages.length - 1; i >= 0; i--) {
            if (newMessages[i]?.role === 'user') { lastUserIdx = i; break; }
        }

        if (totalTokens > hardLimit && lastUserIdx >= 0) {
            const userMsg = newMessages[lastUserIdx];
            if (userMsg && typeof userMsg.content === 'string') {
                let content = userMsg.content;
                const ocrPattern = /【图片\d+:[^】]+的OCR识别结果】\n([\s\S]*?)(?=\n\n【图片|$)/g;
                const truncatedContent = content.replace(ocrPattern, (match, p1) => {
                    const ocrText = p1.trim();
                    if (ocrText.length > 500) {
                        return match.substring(0, match.indexOf(ocrText)) + ocrText.substring(0, 500) + '\n... (OCR结果已截断)';
                    }
                    return match;
                });
                if (truncatedContent !== content) {
                    newMessages[lastUserIdx] = { ...userMsg, content: truncatedContent };
                    totalTokens -= (estimateTokens(content) - estimateTokens(truncatedContent));
                    ChatLog.info?.('PROMPT', `[prompt-builder] truncated OCR results, remaining tokens: ${totalTokens}`);
                }
            }
        }

        if (totalTokens > hardLimit && lastUserIdx >= 0) {
            const userMsg = newMessages[lastUserIdx];
            if (userMsg && typeof userMsg.content === 'string') {
                let content = userMsg.content;
                const ocrPattern = /【图片\d+:[^】]+的OCR识别结果】\n([\s\S]*?)(?=\n\n【图片|$)/g;
                content = content.replace(ocrPattern, (match, p1) => {
                    const ocrText = p1.trim();
                    return match.substring(0, match.indexOf(ocrText)) + ocrText.substring(0, 200) + '\n... (OCR结果已截断)';
                });
                newMessages[lastUserIdx] = { ...userMsg, content };
                totalTokens = estimateTokens(content);
                ChatLog.info?.('PROMPT', `[prompt-builder] aggressively truncated OCR, remaining tokens: ${totalTokens}`);
            }
        }

        if (totalTokens > hardLimit && systemIdx >= 0) {
            const sysMsg = newMessages[systemIdx];
            if (sysMsg) {
                let sysContent = typeof sysMsg.content === 'string' ? sysMsg.content : '';
                const skillPattern = /(技能列表|Skill List|## 可用技能)[\s\S]*?(?=\n\n|$)/;
                sysContent = sysContent.replace(skillPattern, (match) => {
                    if (match.length > 500) return match.substring(0, 500) + '\n... (技能列表已截断)';
                    return match;
                });
                newMessages[systemIdx] = { ...sysMsg, content: sysContent };
                ChatLog.info?.('PROMPT', `[prompt-builder] truncated skill list`);
            }
        }

        totalTokens = newMessages.reduce((sum, m) => {
            const c = typeof m.content === 'string' ? m.content : JSON.stringify(m.content);
            return sum + estimateTokens(c);
        }, 0);

        return { messages: newMessages, truncated: true, totalTokens };
    }

    // ── Prompt 选择与变量收集 ─────────────────────────────────────

    function pickPromptId(turn) {
        const kind = turn?.kind || 'text';
        // 问候检测：简短问候语用 greeting 模板，避免加载通用数据处理提示导致"你好"反问"请选择数据处理方式"
        if (kind === 'text') {
            const msg = (turn?.userMessage || turn?.message || '').trim();
            if (msg && /^(你好[啊呀]?|您好[啊呀]?|早[上啊]好|下午好|晚上好|(hi|hello|hey)\b|good\s*(morning|afternoon|evening))$/i.test(msg)) {
                return 'greeting';
            }
        }
        switch (kind) {
            case 'image': return 'imageAnalysisUnified';
            case 'doc':   return 'documentAttachment';
            case 'text':
            default:      return 'general';
        }
    }

    function collectVariables(turn) {
        const vars = {};

        try {
            if (App.getSkillInfo) {
                vars.skillInfo = App.getSkillInfo();
            } else if (App.cachedSkills) {
                vars.skillInfo = App.cachedSkills.map(s =>
                    `- ${s.displayName || s.name}: ${s.description || ''}`
                ).join('\n');
            } else {
                vars.skillInfo = '（技能列表加载中）';
            }
        } catch (e) {
            vars.skillInfo = '（技能列表不可用）';
        }

        try {
            // 惰性构建：仅当最终渲染的 system prompt 模板真的引用 {toolInfo} 时才耗费
            // CPU 构造完整工具 schema（properties/required/enum）。
            //   正常路径（js/prompts.md 加载成功）模板只引用 {skillInfo}，不含 {toolInfo}——
            //   此时原生 body.tools 已把完整 schema 交给函数调用机制，system prompt 里再无
            //   必要堆一份同样的 schema 文本（避免"工具的完整参数两面写"）。
            //   仅当 prompts.md / prompts.json 均加载失败、回退到含 {toolInfo} 的
            //   App.DEFAULT_PROMPTS（app-bootstrap）时，才需要完整字符串化。
            const needsToolInfoText = (() => {
                const cfg = App.promptConfig;
                if (cfg && cfg.prompts) {
                    for (const p of Object.values(cfg.prompts)) {
                        if (p && typeof p.content === 'string' && p.content.includes('{toolInfo}')) return true;
                    }
                    // 正常路径已加载且无任何模板引用 {toolInfo} → 不需要构造
                    if (App.promptConfigLoaded) return false;
                }
                // 兜底：无法判定时保守构造（维持旧行为）
                return true;
            })();

            if (needsToolInfoText) {
                const toolsArg = Array.isArray(tools) ? tools : null;
                if (toolsArg && toolsArg.length > 0) {
                    vars.toolInfo = toolsArg
                        .map(t => {
                            const head = `- ${t.name}: ${t.description || ''}`;
                            const schema = t.inputSchema;
                            if (!schema) return head;
                            const props = (schema.properties && typeof schema.properties === 'object') ? schema.properties : null;
                            if (!props || Object.keys(props).length === 0) return head;
                            const required = Array.isArray(schema.required) ? schema.required : [];
                            const descKeys = Object.keys(props);
                            // 已展示过的 key 去重
                            const shown = new Set();
                            const lines = [];
                            // 必需的参数在前，说明其必填与枚举
                            for (const key of required) {
                                const p = props[key];
                                if (!p) continue; shown.add(key);
                                const enumInfo = Array.isArray(p.enum) && p.enum.length > 0
                                    ? ` [可选值: ${p.enum.join('/')}]`
                                    : '';
                                const reqTag = '（必填）';
                                lines.push(`    · ${key}${reqTag}${enumInfo}${p.description ? ` ${p.description}` : ''}`);
                            }
                            // 其余参数（含 title/hasLegend 等展示选项）按定义顺序补充
                            for (const key of descKeys) {
                                if (shown.has(key)) continue;
                                const p = props[key];
                                if (!p) continue;
                                const enumInfo = Array.isArray(p.enum) && p.enum.length > 0
                                    ? ` [可选值: ${p.enum.join('/')}]`
                                    : '';
                                lines.push(`    · ${key}${enumInfo}${p.description ? ` ${p.description}` : ''}`);
                            }
                            return lines.length > 0 ? `${head}\n${lines.join('\n')}` : head;
                        })
                        .join('\n') || '（暂无可用 MCP 工具）';
                } else if (App.mcpTools && Array.isArray(App.mcpTools) && App.mcpTools.length > 0) {
                    vars.toolInfo = App.mcpTools.map(t =>
                        `- ${t.name || t.Name}: ${t.description || t.Description || ''}`
                    ).join('\n');
                } else {
                    vars.toolInfo = '（工具列表加载中）';
                }
            } else {
                // 正常路径：模板不引用 {toolInfo}，无需构造（避免每次调用字符串化全部工具 schema 的开销）。
                // 仍置一个兜底值，防止其它独立消费方读到 undefined。
                vars.toolInfo = '';
            }
        } catch (e) {
            vars.toolInfo = '（工具列表不可用）';
        }

        if (turn?.context?.referenceRegions?.length > 0) {
            vars.referenceRegions = turn.context.referenceRegions
                .filter(r => r != null)
                .map((r, i) => `区域${i+1}: [${r.workbook || ''}]${r.worksheet || ''}!${r.address || ''}`)
                .join('\n');
        }

        if (turn?.kind === 'doc' && turn?.attachments?.length > 0) {
            vars.attachmentInfo = turn.attachments
                .filter(a => a != null)
                .map(a => `${a.fileName || a.name || '未命名'} (${a.type || 'unknown'})`)
                .join(', ');
        }

        try {
            if (!getHost().isExcel) {
                // 非 Excel 宿主（Word / PowerPoint）没有"工作簿 / 工作表"概念，保持为空，避免污染提示词
                vars.excelContext = '';
            } else {
                const as = turn?.context?.activeSheet;
                const sel = turn?.context?.selection;
                const sheet = as?.sheet || sel?.sheetName || '';
                const workbook = as?.workbook || sel?.workbookName || '';
                // 【真实工作表清单】来自 chat-entry 调 GetSheets() 取得的全部可见表名。
                // 列在这里能让模型直接抄真实名称，而不是编造"订单""Sheet1"之类。
                let sheetLine = '';
                const sheetList = turn?.context?.workbookSheets;
                if (Array.isArray(sheetList) && sheetList.length) {
                    sheetLine = `\n工作簿内全部工作表（调用带 sheet 参数的工具只能用这些真实名称，严禁编造其他表名）: ${sheetList.join('、')}`;
                }
                // 【用户在标签栏显式多选的工作表】—— 合并/拆分类操作的"作用对象"应该是这几张表，
                // 而不是模型自己凭活动工作表猜测。只有多选(>1)才注入，单选等同于活动工作表语义。
                let selectedLine = '';
                const selectedSheets = turn?.context?.selectedSheets;
                if (Array.isArray(selectedSheets) && selectedSheets.length > 1) {
                    selectedLine = `\n🔔 用户在 Excel 标签栏显式选中了 ${selectedSheets.length} 张工作表（合并/拆分操作必须作用于这些表，sourceSheets 参数就传这些名字）: ${selectedSheets.join('、')}`;
                }
                vars.excelContext = `工作簿: ${workbook || '(未获取)'} | 活动工作表: ${sheet || '(未获取)'}${sheetLine}${selectedLine}`;
            }
        } catch (e) {
            vars.excelContext = '';
        }

        return vars;
    }

    function substituteVars(template, vars) {
        if (!template || !vars) return template || '';
        let result = template;
        for (const [key, value] of Object.entries(vars)) {
            const placeholder = new RegExp(`\\{${key}\\}`, 'g');
            result = result.replace(placeholder, String(value ?? ''));
        }
        return result;
    }

    // ── 意图路由·分析 ─────────────────────────────────────────────
    // 在路由层把"分析/总结"用户请求识别出来，注入任务专用提示（写死意图），
    // 而不是让弱模型在通用系统提示下自行领悟。detectAnalysisIntent 也导出，
    // 供后续更完整的路由器复用。

    // 【扩展 v20260926】覆盖所有 Excel 操作意图（图表/可视化/汇总/透视/排序/清洗/格式化等），
    // 让 finalContent 合成逻辑一律优先 LLM 总结文本（而非原始工具结果 JSON）。
    // 与 context-engine._hasExcelOpIntent 的意图覆盖保持一致。
    const ANALYSIS_INTENT_RE = /分析|总结|汇总|解读|结论|评估|洞察|点评|概述|归纳|概括|研判|趋势|饼图|柱图|折线图|柱状图|条形图|散点图|面积图|雷达图|直方图|趋势图|仪表盘|图形化|画图|制图|可视化|图表|chart|graph|visuali[sz]e|plot|diagram|生成.*图|做.*图|出.*图|创建.*图|透视|pivot|交叉表|筛选|排序|过滤|清洗|去重|remove.?duplicate|按.*分组|按.*展示|按.*统计|按.*汇总|按.*计算|计算.*排名|排名|百分比|占比|分布|对比|比较|差异|相关性|回归|预测|clean|sort|filter|aggregate|summary/;

    // 纯操作意图（不需要先读数据，直接调工具）
    const OPERATION_ONLY_INTENT_RE = /格式化|美化|样式|主题|合并单元格|拆分|转置|transpose|format_range|beautify|apply_style/;

    function detectOperationOnlyIntent(text) {
        if (!text || typeof text !== 'string') return false;
        // 分析类优先命中（避免"按销售额汇总并美化"被误判）
        if (ANALYSIS_INTENT_RE.test(text)) return false;
        return OPERATION_ONLY_INTENT_RE.test(text);
    }

    // 【v20261001 文件产出意图】导出／导入／另存／保存到文件。
    // 与"格式结构操作"共用同一个前提（不需要先把数据读进上下文：导出/保存工具自己会读取源表并落盘），
    // 但工具族完全不同（导出/保存工具 vs 格式化工具），所以单独成类——若直接并进
    // OPERATION_ONLY_INTENT_RE，buildSystemPrompt 会注入只列 format_range/merge_cells/transpose
    // 的 OPERATION_ONLY_TASK_PROMPT，反而把模型带偏。
    //
    // 实测根因（日志取证）：用户说"将当前工作表导出为CSV并保存到桌面"，
    //   detectAnalysisIntent → false（无分析词）
    //   detectOperationOnlyIntent → false（旧词表只有 格式化|美化|样式|主题|合并单元格|拆分|转置）
    // ⇒ 落进"分析类"分支，被注入"先调用 excel_read_range 读取活动工作表的实际数据"，
    //   模型照做：读完整张表($A$1:$G$201) → 输出一段数据分析，导出工具一次都没被调用，
    //   任务卡却显示"任务完成"，磁盘上也没有 CSV。
    const FILE_OUTPUT_INTENT_RE = /导出|导入|另存为|另存|存为|保存为|保存成|保存到|保存至|保存|输出为|输出到|输出成|导出成|批量导出|落盘|写入文件|写文件|生成\s*(?:csv|excel|xlsx|pdf)|转\s*(?:成|为)\s*(?:csv|excel|xlsx|pdf|文件)|export|save\s*as|save\s*to|to\s*csv/i;

    /**
     * 判断原始用户消息是否属于"文件产出"意图（导出/另存为/保存到文件）。
     * @param {string} text - 原始用户消息（优先用户真实输入）
     * @returns {boolean}
     */
    function detectFileOutputIntent(text) {
        if (!text || typeof text !== 'string') return false;
        return FILE_OUTPUT_INTENT_RE.test(text);
    }

    const FILE_OUTPUT_TASK_PROMPT =
        '【任务类型：文件产出（导出 / 另存为 / 保存到文件）】\n' +
        '用户要把数据、图表或工作簿导出/保存成文件。务必遵循：\n' +
        '1. 直接调用"导出/保存"类工具完成落盘，【不要先调 excel_read_range 把整表读进上下文】——这类工具自己会读取源数据。\n' +
        '2. 工具选择：\n' +
        '   - 导出工作表为 CSV → excel_export_csv（sheet 留空即活动工作表）\n' +
        '   - 保存工作簿 → excel_save_workbook\n' +
        '   - 导出图表图片 → excel_export_chart_image（图表名可用 excel_list_charts 获取）\n' +
        '   - 拆分工作簿/工作表 → excel_split_workbook / excel_split_by_column / excel_split_by_rows\n' +
        '3. filePath / outputDir 可以留空：系统会自动落到真实桌面；严禁编造路径（如 C:\\Users\\Yourusername\\...）。\n' +
        '4. sheet/range 不传即作用于活动工作表；严禁编造 sheet 名或 range（如 "Sheet1"、"A1:C10"）。\n' +
        '5. 完成后用工具返回的真实路径报告结果；不要只给数据分析，也不要在没有工具调用的情况下声称"已完成"。\n' +
        '6. 全程使用与用户提问相同的语言。\n';

    const OPERATION_ONLY_TASK_PROMPT =
        '【任务类型：格式与结构操作】\n' +
        '用户希望对 Excel 做格式美化或结构调整。务必遵循：\n' +
        '1. 直接调用工具执行，【不要先调 excel_read_range 读取数据内容】——这类操作只需要目标范围，不需要知道单元格里写了什么值。\n' +
        '2. 范围解析：优先用选区；没选区时用活动工作表 UsedRange（不传 range 参数让后端自动识别）。\n' +
        '3. 工具选择：\n' +
        '   - 美化/格式化/样式/主题 → excel_format_range\n' +
        '   - 合并单元格 → excel_merge_cells\n' +
        '   - 拆分/转置 → excel_transpose\n' +
        '4. 操作完成后输出简短结果摘要（Markdown 格式），列出：改了什么、影响范围、结果如何。\n' +
        '5. 全程使用与用户提问相同的语言。\n';


    /**
     * 判断原始用户消息是否属于"仅展示数据"意图（只想看到数据，明确不需要分析）。
     * 与 detectAnalysisIntent 互斥，优先级更高。
     * @param {string} text
     * @returns {boolean}
     */
    function detectDisplayIntent(text) {
        if (!text || typeof text !== 'string') return false;
        // 必须有"展示类动词" + "数据类名词"
        if (!/(展示|显示|呈现|列出|查看|看看|输出|渲染|做成)/.test(text)) return false;
        // 【补"条/行/名"】"仅显示前5条""只看前10行"这类请求不含 表格/数据 等词，
        // 旧词表漏判 → 落到通用分支读全表（实测：100 条全量回显）。补量词后命中本分支。
        if (!/(表格|数据|结果|明细|列表|清单|行列|内容|条|行|名)/.test(text)) return false;
        // 一旦含分析/聚合/计算语义，就不算"仅展示"（交给 analysis 分支）
        if (/(分析|总结|汇总|统计|对比|比较|排名|趋势|占比|比例|计算|评估|洞察|解读|结论|归纳|概括|按|各|分组|每个|分别|透视|聚合|同比|环比)/.test(text)) return false;
        return true;
    }

    /**
     * 【仅展示数据】任务提示：明确禁止附加分析。
     */
    const DISPLAY_ONLY_TASK_PROMPT =
        '【任务类型：仅展示数据（禁止分析）】\n' +
        '用户只想"看到"数据，并未要求分析。务必严格遵循：\n' +
        '1. 只调用 excel_read_range 读取数据，并把结果以表格展示（参数带 _view:"table"）。\n' +
        '2. ❌ 严禁附加任何分析、总结、结论、趋势解读、指标计算（合计/均值/最大/最小/占比/排名等）。\n' +
        '3. ❌ 严禁输出"以下是分析结果""数据总览/销售总览"之类的段落。\n' +
        '4. 最多用一句话说明"已展示/已读取"，或直接给结果，不要长篇大论。\n' +
        '5. 全程使用与用户提问相同的语言。\n' +
        '6. 若用户要求"仅显示/只看前N条(行)"：range 只取表头+前N行（如前5条=第1~6行，即 A1:I6 这种），【严禁读取整个已用区域】；只有当用户要的是"前N名"（按某指标排序）时，才改用 TopN 汇总。\n';

    /**
     * 判断原始用户消息是否属于"数据分析/总结"意图。
     * @param {string} text - 原始用户消息（优先用户真实输入）
     * @returns {boolean}
     */
    function detectAnalysisIntent(text) {
        if (!text || typeof text !== 'string') return false;
        return ANALYSIS_INTENT_RE.test(text);
    }

    const ANALYSIS_TASK_PROMPT =
        '【任务类型：数据分析与总结】\n' +
        '用户希望对 Excel 中的数据做分析/总结/聚合/出图。务必遵循：\n' +
        '1. 定数据范围：目标是当前选中区域对应的整张表，不要只取单行单格。\n' +
        '   ⚠️ 关于 dimColumn（分组维度列名）的选择——【严禁瞎猜 "地区"、"部门"、"产品" 等关键词】：\n' +
        '      • 如果用户可见【当前选区】→ 直接用 prompt 里已经展示的 "表头: A, B, C, ..." 中的**真实列名**作为 dimColumn\n' +
        '      • 如果 prompt 里没展示表头 → 调 excel_read_range 看一眼表头再决定\n' +
        '      • 拿不准 → 干脆不传 dimColumn（留空），后端会自动选最佳维度列（优先级：部门>地区>产品>渠道>分类）\n' +
        '2. 输出结论，禁止抄数据：不要把 Excel 的原始行或整表原样复述回来。必须基于数据给出分析结论。\n' +
        '3. 建议回复结构：总体情况 → 关键发现 → 明显问题或趋势 → 结论与建议。\n' +
        '4. 数据较大（几十行以上）时，先做统计汇总（合计/平均/最大/最小、分类占比或排名）再给结论。\n' +
        '5. 全程使用与用户提问相同的语言。\n' +
        '6. 【格式规范·强制】所有面向用户的输出必须使用 Markdown 排版：\n' +
        '   - 结构化数据（汇总表、透视结果、预览行等）→ 使用 Markdown 表格（|列1|列2|列3| 格式）\n' +
        '   - 操作清单、变更列表、数据要点 → 使用有序/无序列表（1. / - 格式）\n' +
        '   - 关键数字、结论、建议 → 使用 **加粗** 突出\n' +
        '   - ❌ 禁止 使用 "键=值 | 键=值 | 键=值" 这种原始管道拼接格式，必须转成 Markdown 表格\n' +
        '   - ❌ 禁止 把工具返回的原始 JSON 直接展示给用户\n' +
        '   - ❌ 禁止 输出空表格/占位表格：表格每一行都必须填真实数据；只有表头、或数据行全是空白/"$"/"待填充"等占位符的表格一律不许输出。没有真实数据（如工具未返回明细）时不要画表格，改用文字说明；表格数值必须来自工具真实返回，严禁编造';

    /**
     * 通用输出格式规范（所有 agent-loop 任务都应遵循）。
     * 注入到 system prompt 的通用规范 section。
     */
    const GENERAL_FORMAT_PROMPT =
        '【通用输出格式规范·强制】\n' +
        '所有面向用户的回复必须使用 Markdown 排版，确保输出工整易读：\n' +
        '- 结构化数据（清洗结果预览、汇总表、按类别/部门/地区的统计等）→ 必须使用 Markdown 表格\n' +
        '- 操作变更列表（清洗做了什么、美化改了什么）→ 使用有序列表 1. 2. 3.\n' +
        '- 关键数字和结论 → **加粗**\n' +
        '- ❌ 绝对禁止：键=值 | 键=值 的管道拼接行（用 Markdown 表格替代）、原始 JSON 直接展示、无分段的长段落\n' +
        '- ❌ 绝对禁止：把 Excel 原始数据整行整列抄回来（用统计汇总/表格/列表呈现结论）';

    // ── 主入口：使用 PromptBuilder 组装 system prompt ────────────

    /**
     * 使用 PromptBuilder 组装 system prompt。
     * 从 prompts.md 加载主模板，构造 sections 并渲染。
     * @param {Object} turn - turn 对象
     * @param {Object} [options] - 可选参数
     * @param {string} [options.persona] - 替换默认助手身份（如 "你是数据分析专家"）
     * @param {Array<Function>} [options.assembleListeners] - 额外 assemble 瀑布监听器
     * @returns {string} 完整的 system prompt
     */
    function buildSystemPrompt(turn, options = {}) {
        const promptId = pickPromptId(turn);
        const vars = collectVariables(turn);

        // 加载主模板
        let mainTemplate = '';
        try {
            if (App.getPrompt) {
                mainTemplate = App.getPrompt(promptId, vars);
            } else if (App.DEFAULT_PROMPTS?.[promptId]?.content) {
                mainTemplate = substituteVars(App.DEFAULT_PROMPTS[promptId].content, vars);
            }
        } catch (e) {
            ChatLog.error?.('PROMPT', `[prompt-builder] 加载 ${promptId} 失败`, e);
            mainTemplate = 'You are a helpful AI assistant.';
        }

        if (!mainTemplate) {
            ChatLog.warn?.('PROMPT', `[prompt-builder] ${promptId} 返回空，使用 default`);
            mainTemplate = App.getPrompt?.('default', vars) || '';
        }

        // 使用 PromptBuilder 构建
        const pb = new PromptBuilder()
            .setVariables(vars);

        // 设置 persona（如果提供了则替换默认身份）
        // 宿主感知：非 Excel 宿主下默认身份改为 Word / PowerPoint（Excel 下文本完全不变）
        const _host = getHost();
        if (options.persona) {
            pb.setPersona(options.persona);
        } else {
            pb.addSection('identity', `你是SmartPilot，一个智能的${_host.label}辅助工具。`, -100);
        }

        // 注入真实系统日期时间（live system data，非训练知识）。
        // 让模型在回答"今天几号/日期/时间"类问题时有权威时钟依据，避免凭训练数据臆造日期。
        pb.addSection('now', () => {
            const d = new Date();
            const pad = n => String(n).padStart(2, '0');
            const week = ['日','一','二','三','四','五','六'][d.getDay()];
            return `当前日期与时间：${d.getFullYear()}年${d.getMonth()+1}月${d.getDate()}日 ${pad(d.getHours())}:${pad(d.getMinutes())}（星期${week}）。涉及"今天/当前日期/时间"的问题，一律以此处的真实时间为准，不要凭训练知识推算或臆造日期。`;
        }, -50);

        // 【步骤元数据】编辑/重试前被移除轮次的工具步骤结果（真实读取内容）——
        // 用户原则：步骤/思考返回的内容是上下文元数据，不得直接丢弃。此处把它注入，
        // 供模型直接引用（避免重试后把已读到的真实数据当不存在）。
        pb.addSection('prevAttemptSteps', () => {
            const steps = turn && turn.context && turn.context._prevAttemptSteps;
            if (!Array.isArray(steps) || steps.length === 0) return '';
            const lines = ['【上一轮尝试的步骤结果（编辑/重试前真实执行的工具返回，可放心引用，非编造）】'];
            for (const s of steps.slice(-6)) {
                const argStr = Object.keys(s.args || {}).slice(0, 3)
                    .map(k => `${k}=${String(s.args[k]).slice(0, 40)}`).join(', ');
                lines.push(`- ${s.name}(${argStr}) → ${s.ok ? '成功' : '失败'}: ${String(s.result || '').slice(0, 300)}`);
            }
            return lines.join('\n');
        }, -20);

        // 【意图路由·分析】命中分析/总结意图 → 注入任务专用提示（写死意图，
        // 强调"读整表+输出结论+禁止抄数据"），置于主模板之前。
        const rawUser = (turn && (turn._originalUserMessage || turn.userMessage)) || '';
        // 下列任务专用提示均为 Excel 专属（涉及 excel_* 工具、工作表/选区/工作簿）。
        // 非 Excel 宿主不注入这些提示，改注入"运行环境"提示，避免把 PowerPoint/Word 当成 Excel 环境
        // 而回答"当前工作簿中没有活动选区，请先在 Excel 中框选"之类。
        if (!_host.isExcel) {
            pb.addSection('hostEnv', buildHostEnvPrompt(_host), -10);
            ChatLog.info?.('PROMPT', `[prompt-builder] 非 Excel 宿主（${_host.kind}），注入宿主环境提示，跳过 Excel 专属任务提示`);
            // 内容处理（总结/润色/改写/翻译/提炼）：必须显式要求"直接产出成品、禁止反问"，
            // 否则弱模型读完全文后只回一句"请告诉我您想怎么润色"就收尾（实测，任务卡却显示完成）。
            if (detectContentTaskIntent(rawUser)) {
                pb.addSection('hostContentTask', buildHostContentTaskPrompt(_host, rawUser), -8);
                ChatLog.info?.('PROMPT', '[prompt-builder] 非 Excel 宿主命中"内容处理"意图，注入内容处理专用提示（禁止反问 / 必须直接产出）');
            } else if (detectFileOutputIntent(rawUser)) {
                // 导出/保存类（PPT/Word 导出 PDF 等）：告诉模型直接调导出工具、filePath 留空，
                // 严禁编造占位符路径（如 C:\Users\Yourusername\Documents\report.pdf）。
                pb.addSection('hostFileOutput', buildHostFileOutputTaskPrompt(_host), -8);
                ChatLog.info?.('PROMPT', '[prompt-builder] 非 Excel 宿主命中"文件产出"意图，注入文件产出专用提示（filePath 留空 / 禁止编造路径）');
            }
        } else if (detectDisplayIntent(rawUser)) {
            pb.addSection('displayOnly', DISPLAY_ONLY_TASK_PROMPT, -10);
            ChatLog.info?.('PROMPT', '[prompt-builder] 命中"仅展示数据"意图，注入展示专用提示（禁止附加分析）');
        } else if (detectAnalysisIntent(rawUser)) {
            pb.addSection('analysisIntent', ANALYSIS_TASK_PROMPT, -10);
            ChatLog.info?.('PROMPT', '[prompt-builder] 命中"分析/总结"意图，注入分析专用提示');
        } else if (detectOperationOnlyIntent(rawUser)) {
            pb.addSection('operationOnly', OPERATION_ONLY_TASK_PROMPT, -10);
            ChatLog.info?.('PROMPT', '[prompt-builder] 命中"格式/结构操作"意图，注入操作专用提示（不要求先读数据）');
        } else if (detectFileOutputIntent(rawUser)) {
            // 导出/保存类：必须显式告诉模型"直接调导出工具"，否则弱模型会先读整表再输出分析（实测失败）。
            // 放在 opOnly 之后：保留"美化并保存"这类请求原有的格式操作提示行为。
            pb.addSection('fileOutput', FILE_OUTPUT_TASK_PROMPT, -10);
            ChatLog.info?.('PROMPT', '[prompt-builder] 命中"文件产出（导出/保存）"意图，注入文件产出专用提示（不要求先读数据）');
        }
        // 关键词硬短路主动让出的外观语意（统一为中国红等）。必须写在操作提示之后，
        // 否则模型会按「没指定格式 → 空 format → 商务蓝」把用户点名的颜色丢掉。
        if (_host.isExcel && turn && turn.context && turn.context._semanticStyleTarget) {
            const _styleTarget = String(turn.context._semanticStyleTarget);
            pb.addSection('semanticStylePlan',
                '【本轮由你规划，按原话逐项做完】\n' +
                '用户指定了外观：「' + _styleTarget + '」。这句话关键词快捷路径盖不住，不要套默认美化。\n' +
                '1. 先调用 excel_get_sheets，拿到每一张可见工作表的名字。\n' +
                '2. 再对每一张表调用 excel_format_range。范围按这句话定，不要因为没手选区域就停下来：\n' +
                '   - 只说了美化每张表：每张表用它自己的数据区（至少表头加一行）。不要留空 range，不要用当前光标那一个格子，也不要把某一张表上的选区套到其他表。\n' +
                '   - 嘴里说了地址（如 A1:D10）：每张目标表都用这个地址。\n' +
                '   - 圈了一行或一块，并且说的是当前表或「选中的」：就用这一行或这块。只点中一个格子不算圈定，改用该表数据区。\n' +
                '   - 说「其他表也按这个区域」并且确实圈了一块：其他表用同一地址。只停在一个格子上则每张表仍用自己的数据区。\n' +
                '   - 点名了工作表：只做那几张，选区若在别的表上就不用。\n' +
                '   - 只要表头：用数据区的第一行。某张表找不到数据区就跳过并说明，不要拿单格硬做。\n' +
                '3. format 必须写出「' + _styleTarget + '」的表头底色（bgScope=header）。'
                + (/中国红|朱红|国旗红/.test(_styleTarget)
                    ? '中国红是饱和正红 #DE2910，不是浅红、粉色或商务红，禁止使用 #F4C2C4。'
                    : '')
                + '表头白字只能写 headerFontColor=FFFFFF。禁止写 fontColor=FFFFFF，否则整段文字变白，表身数据会看不见。禁止把 format 传成 {}。\n' +
                '4. 原话里的两项都要留下：每一张可见表都要做，指定的颜色也要留下。只做了一张表，或做成了别的颜色，都还没完成。\n',
                -8);
            ChatLog.info?.('PROMPT', '[prompt-builder] 外观语意交给模型规划: ' + _styleTarget);
        }

        // 【图表强意图补救】orchestrator 检测到用户明确要图表（饼图/柱图/折线图等），
        // 但 chart skill 不可用 → 回落 text-pipeline。必须注入图表引导，
        // 否则 LLM 只会调 excel_aggregate 做文字分析，忘了调 excel_create_chart 真正画图。
        //
        // ⚠️ 这里是"生成了一个没用图"的直接根因修复点：
        // 旧文案写"先调用 excel_aggregate 或 excel_read_range 拿到聚合/原始数据"，
        // 等于给了模型绕过聚合的漏洞——实测 qwen2.5 直接 excel_read_range($A$1:$I$101)
        // → excel_create_chart(dataRange=$A$1:$I$101)，对明细整表建图，图例是原始行、毫无意义。
        // 修复：① 消息含分组维度意图时，强制"先聚合再建图"，删掉"或 excel_read_range"的歧义；
        //       ② 直接用 splitAggregateTasks 解析出的**具体参数**写进引导——
        //          具体值（dimColumn="部门"…）远比抽象规则更易被弱模型逐字遵循；
        //       ③ 明确禁止"先读明细再拿明细范围建图"。
        if (turn && turn.context && turn.context._forceChart) {
            const _chartUserMsg = rawUser || '';
            // 分组维度意图：按X / 各X / 分组 / 汇总 / 聚合 / 占比 / 分布 / 排名 / 对比…
            const _needAggRe = /(按(?!照|时|钮|下|上)|各(部门|地区|区域|类|项|产品|渠道|城市|省份|月|季度|年)|分组|汇总|聚合|统计|占比|分布|排名|分别|每个|对比)/;
            const _needAgg = _needAggRe.test(_chartUserMsg);
            // 用确定性解析器取聚合参数，作为"直接采用"的硬约束
            let _aggParamLine = '';
            let _aggTargetSheet = '聚合结果';
            try {
                const _tp = App.ChatV2 && App.ChatV2.taskPlanner;
                const _tasks = (_tp && typeof _tp.splitAggregateTasks === 'function')
                    ? _tp.splitAggregateTasks(_chartUserMsg) : null;
                const _t = Array.isArray(_tasks) ? _tasks[0] : null;
                if (_t && _t.tool === 'excel_aggregate' && _t.params) {
                    const _p = _t.params;
                    _aggTargetSheet = _p.targetSheet || '聚合结果';
                    _aggParamLine = `dimColumn="${_p.dimColumn || ''}"、valueColumn="${_p.valueColumn || ''}"、aggFunc="${_p.aggFunc || 'sum'}"、targetSheet="${_aggTargetSheet}"`;
                }
            } catch (_e) { /* 解析失败 → 走通用强约束文案 */ }

            const CHART_ACTION_GUIDE = _needAgg
                ? ('【⚠️ 强制：先聚合、再建图 —— 严禁对明细整表直接建图】\n' +
                   '用户要"按某个维度展示/对比某个指标"，明细表里同一类别有多行，直接对明细建图必然是错图。必须严格按顺序执行，不得跳步：\n' +
                   '第1步 调用 excel_aggregate（严禁用 excel_read_range 替代！）：\n' +
                   (_aggParamLine ? ('   参数直接采用（不要改动）：' + _aggParamLine + '\n')
                                  : '   参数：dimColumn=分组维度列名、valueColumn=数值列名、aggFunc=sum\n') +
                   '   它返回"类别+汇总值"两列及其所在工作表范围（例如 ' + _aggTargetSheet + '!A1:B6）。\n' +
                   '第2步 调用 excel_create_chart，dataRange 必须设为第1步返回的聚合结果范围（例如 "' + _aggTargetSheet + '!A1:B6"），\n' +
                   '   绝不能用 A1:I101 这类明细整表范围；title 用"各<维度><指标>对比"（如"各部门销售额对比"）。\n' +
                   '绝对禁止：先用 excel_read_range 读明细、再拿明细范围建图。即使你已经读过明细，也必须再调用 excel_aggregate 得到聚合结果。\n' +
                   '建图完成后输出一句简短结论即可。')
                : ('【强制：图表生成】用户要求生成图表，你必须：\n' +
                   '1. 先调用 excel_aggregate 或 excel_read_range 拿到聚合/原始数据\n' +
                   '2. 紧接着调用 excel_create_chart 创建图表！不要只做文字分析就停。\n' +
                   '   excel_create_chart 的关键参数：\n' +
                   '   - chartType: 图表类型（pie=饼图, column=柱图, bar=条形图, line=折线图, scatter=散点图）\n' +
                   '   - dataRange: 数据范围（聚合结果所在的工作表区域，如"汇总!A1:B10"）\n' +
                   '   - title: 图表标题（如"各部门销售额对比"）\n' +
                   '3. 创建完图表后再输出分析结论。\n' +
                   '4. 禁止说"接下来我将使用饼图..."但实际不调用图表工具——必须真正生成图表！');
            pb.addSection('chartAction', CHART_ACTION_GUIDE, -5);
            ChatLog.info?.('PROMPT', `[prompt-builder] 检测到 _forceChart，注入图表引导（needAgg=${_needAgg}${_aggParamLine ? '，已注入确定性聚合参数' : ''}）`);
        }

        const _groupPlan = turn && turn.context && turn.context.groupPlan;
        if (_host.isExcel && _groupPlan && _groupPlan.collapse && (_groupPlan.dimColumn || _groupPlan.dateColumn)) {
            const _gStep = _groupPlan.grain
                ? ('excel_date_group({ dateColumn: "' + (_groupPlan.dateColumn || '') + '", groupBy: "' + _groupPlan.grain + '"'
                    + (_groupPlan.valueColumn ? ', aggColumn: "' + _groupPlan.valueColumn + '"' : '')
                    + ', targetSheet: "分组结果" })')
                : ('excel_aggregate({ dimColumn: "' + _groupPlan.dimColumn + '"'
                    + (_groupPlan.valueColumn ? ', valueColumn: "' + _groupPlan.valueColumn + '"' : '')
                    + ', aggFunc: "' + (_groupPlan.aggFunc || 'sum') + '", targetSheet: "分组结果" })');
            pb.addSection('groupFirst',
                '【分组是独立的一步，不代替用户要的功能】\n' +
                '先调用 ' + _gStep + '，把结果写到「分组结果」。源表不要改。\n' +
                '分组完成后继续做用户剩下的要求：画图就用分组结果的范围调 excel_create_chart；要写入、筛选、格式或别的操作，也在这一步之后做。\n' +
                '用户只要求分组时，做完这一步就停。禁止只读表明细后声称已经分组。',
                -4);
            ChatLog.info?.('PROMPT', '[prompt-builder] 注入分组步骤 ' + _groupPlan.label);
        }

        // 【看板/仪表盘强引导】重写为短指令清单（弱模型适配）。
        // 之前的问题：长 5 步 essay → LLM 把计划当文字输出 + 只调最后一个工具 + 失败后抄模板文字假装完成。
        // 修复：祈使句 + 参数示例 + 失败处理 + 反抄模板约束。
        if (_host.isExcel && turn && turn.context && turn.context._forceDashboard) {
            const DASHBOARD_ACTION_GUIDE =
                '【⚠️ 必须生成看板 — 不是分析报告】\n\n' +
                '⏩ 首选方案：一次调用 excel_generate_dashboard 完成全部工作！\n' +
                '   该工具内部确定性完成：读数据 → 算 KPI → 建/清看板表 → 写 KPI 卡 → 按维度建图 → 标题横幅+美化。\n' +
                '   只需要传参数，例如：\n' +
                '   excel_generate_dashboard({ sourceSheet: "销售数据", dims: ["部门","地区","产品"], value: "销售额", title: "销售看板" })\n' +
                '   dims 不传则由引擎自动枚举维度。调完它即完成，不要再拆成多步。\n\n' +
                '只有当 excel_generate_dashboard 不存在或调用失败时，才改用下面的手工多步流程：\n\n' +
                '1. excel_create_sheet(name="看板") 然后 excel_clear_sheet(sheet="看板")\n' +
                '   创建看板专用工作表后，【无条件】清空一次（上次运行可能残留 #N/A 公式/旧数据），\n' +
                '   保证干净起点。两步都必须调，不要跳过清空。\n\n' +
                '2. excel_aggregate（调 3 次）\n' +
                '   分别按部门 / 地区 / 产品各汇总一次。dimColumn 不填也可，后端自动选。\n' +
                '   ⚠️ 必须传 sheet="源数据表名"（如"销售数据"）！create_sheet 之后活动表会变成空的看板表，\n' +
                '   不传 sheet 会聚合空表而失败。若报"未选中有效数据区域"，就是忘了传 sheet——带上重试。\n\n' +
                '3. excel_create_chart（调 3 次）\n' +
                '   部门汇总→柱图(chartType="xlColumnClustered")，地区→饼图(chartType="xlPie")，产品→条形图(chartType="xlBarClustered")。\n' +
                '   dataAddress 用上次 aggregate 的结果表，如"聚合结果1!A1:B10"。\n' +
                '   必须传 sheet="看板" 让图表放到看板表上，并用 position 锚定布局，如 position="{\\"anchorRange\\":\\"B4:H16\\"}"（第二张 B18:H30，第三张 B32:H44），避免图表散落在聚合表上。\n\n' +
                '4. excel_write_range（只写 KPI 汇总，不要复制原始数据！）\n' +
                '   写一个小的 KPI 块到看板表：sheet="看板", range="A3", values=[["销售总额","总销售量","平均单价"],["26793","...","..."]]（≤10 行）。\n' +
                '   固定写在 A3 起（A1:I2 留给系统绘制标题横幅，不要占用第 1~2 行）。\n' +
                '   ⚠️ 禁止把原始数据整表（如 $A$1:$I$101 的 101 行）写进看板——values 参数装不下会失败；\n' +
                '   若失败信息为"数据为空或格式不正确"，说明 values 太大或格式错，改写成 ≤10 行的小数据块重试。\n\n' +
                '5. excel_format_range\n' +
                '   美化看板表的标题和数据区域。\n\n' +
                '【四条铁律】\n' +
                '① 工具失败必须重试或报告——禁止假装成功后输出模板里的文字。\n' +
                '② 禁止写"下一步我将..."但实际不调工具。\n' +
                '③ 最终总结必须只引用实际成功调用的工具结果，禁止从本提示词复制内容作为"完成"说明。\n' +
                '④ 严禁往看板表写 VLOOKUP/公式来"拼"数据——公式引用不存在的聚合结果会整表 #N/A。\n' +
                '   聚合失败时先修正参数（补 sheet=源数据表名）重试聚合，成功后再用 write_range 写静态值。';
            pb.addSection('dashboardAction', DASHBOARD_ACTION_GUIDE, -3);
            ChatLog.info?.('PROMPT', '[prompt-builder] 检测到 _forceDashboard，注入看板强引导（短指令版）');
        }

        // 主模板
        pb.addSection('main', mainTemplate, 0);

        // 注册瀑布监听器（如果有）
        if (Array.isArray(options.assembleListeners)) {
            for (const listener of options.assembleListeners) {
                if (typeof listener === 'function') {
                    pb.onAssemble(listener);
                }
            }
        }

        // 如果有额外注册的 section（由 pipeline/hook 注入），合并到全局默认实例
        const globalPb = getDefaultBuilder();
        for (const gs of globalPb._sections) {
            if (gs.name !== 'identity' && gs.name !== 'main') {
                pb.addSection(gs.name, gs.text, gs.order);
            }
        }
        for (const [k, v] of Object.entries(globalPb._variables)) {
            if (!(k in vars)) {
                pb.setVariable(k, v);
            }
        }

        const result = pb.render();
        ChatLog.debug?.('PROMPT', `[prompt-builder] 已渲染 prompt: ${promptId}, ${estimateTokens(result)} tokens`);
        return result;
    }

    /**
     * 组装完整的 messages 数组（system + history + user）
     * @param {Object} turn
     * @param {Array} history
     * @param {Array} tools
     * @param {Object} [options] - 透传给 buildSystemPrompt 的选项（persona, assembleListeners 等）
     * @returns {Array<{role, content}>}
     */
    function buildMessages(turn, history = [], tools = [], options = {}) {
        const systemPrompt = buildSystemPrompt(turn, options);
        const messages = [];

        if (systemPrompt) {
            messages.push({ role: 'system', content: systemPrompt });
        }

        const MAX_HISTORY = 20;
        const trimmedHistory = (history || []).slice(-MAX_HISTORY);
        for (const h of trimmedHistory) {
            if (!h || !h.role || !h.content) continue;
            messages.push({
                role: h.role,
                content: typeof h.content === 'string' ? h.content : JSON.stringify(h.content)
            });
        }

        const userContent = buildUserMessage(turn);
        messages.push({ role: 'user', content: userContent });

        let toolTokens = 0;
        if (tools && tools.length > 0) {
            const toolsJson = JSON.stringify(tools.map(t => ({
                type: 'function',
                function: { name: t.name, description: t.description }
            })));
            toolTokens = estimateTokens(toolsJson);
        }

        const contextWindow = getContextWindow();
        const result = applyTokenBudget(messages, contextWindow, toolTokens);

        if (result.truncated) {
            ChatLog.info?.('PROMPT', `[prompt-builder] applied token budget: ${result.totalTokens}/${contextWindow} (tools: ${toolTokens})`);
        }

        return result.messages;
    }

    // ── 能力／工具清单类问题 ──────────────────────────────────────
    // 正常路径下 system prompt 不含 {toolInfo}（工具仅以 function-calling schema 形式提供），
    // 弱模型无法枚举，只能凭印象报出 1~2 个（实测 PPT 16:21：实际 29 个工具只报出 1 个）。
    // 框架确切知道清单 → 这类问题直接把权威清单注入，模型照实回答，不猜不编。
    const CAPABILITY_QUESTION_RE = /多少(个|种)?工具|几个工具|有哪些工具|有什么工具|工具(列表|清单|数量|名)|可用工具|支持(哪些|什么)功能|能(做|干)什么|你会做什么|能力(清单|范围)/;

    function buildToolInventoryText() {
        const raw = (Array.isArray(window._mcpTools) && window._mcpTools.length > 0)
            ? window._mcpTools
            : (Array.isArray(App.mcpTools) ? App.mcpTools : []);
        const list = raw.filter(t => t && (t.name || t.Name));
        if (list.length === 0) return '';
        const bySource = {};
        for (const t of list) {
            const full = String(t.name || t.Name);
            const idx = full.indexOf('__');
            const source = idx > 0 ? full.slice(0, idx) : (t.source || 'mcp');
            const shortName = idx > 0 ? full.slice(idx + 2) : full;
            const dn = t.displayName || t.DisplayName || '';
            (bySource[source] = bySource[source] || []).push(dn ? `${shortName}（${dn}）` : shortName);
        }
        let txt = `【你当前可用的工具清单（系统注入的权威数据，回答工具/能力类问题时必须以此为准）】\n` +
            `工具总数：${list.length} 个\n`;
        for (const src of Object.keys(bySource)) {
            txt += `\n▸ ${src}（${bySource[src].length} 个）：\n  - ${bySource[src].join('\n  - ')}\n`;
        }
        txt += `\n回答"你有多少工具 / 有哪些工具 / 你能做什么"时，直接按上述清单如实回答（可按类别分组），` +
            `严禁遗漏或编造清单之外的工具。`;
        return txt;
    }

    function buildUserMessage(turn) {
        if (!turn) return '';
        const kind = turn.kind || 'text';
        const userMessage = turn.userMessage || '';

        if (turn._ocrDone) return userMessage;

        if (kind === 'image' && turn.attachments?.length > 0) {
            const content = [{ type: 'text', text: userMessage || '请分析这些图片' }];
            for (const imageAttachment of turn.attachments) {
                if (!imageAttachment || imageAttachment.type !== 'image') continue;
                const base64 = imageAttachment.base64 || imageAttachment.data || '';
                if (!base64) continue;
                const mime = imageAttachment.mime || 'image/png';
                content.push({
                    type: 'image_url',
                    image_url: { url: `data:${mime};base64,${base64}` }
                });
            }
            return content.length > 1 ? content : (userMessage || '');
        }

        if (kind === 'doc' && turn.attachments?.length > 0) {
            const userQ = userMessage || '[文件已上传]';
            const docs = turn.attachments.filter(a => a && a.type !== 'image');
            const images = turn.attachments.filter(a => a && a.type === 'image');
            const lines = docs.map((a, i) => `${i + 1}. ${a.fileName || a.name || '未命名'}`);
            let text = `<user_question>${userQ}</user_question>\n`;
            text += '附件正文在随后的 <document> 里。只根据那份正文回答，不要读取当前 Excel 工作表。\n';
            if (lines.length) text += `<attachments>\n${lines.join('\n')}\n</attachments>\n`;
            if (images.length) text += `已有 ${images.length} 张页面图片。\n`;
            text += '用户没写具体问题时，概述附件正文。只根据读到的内容说，不要编造，也不要复述本段说明。';
            return text;
        }

        let extraContext = '';

        // 【能力清单】"你有多少工具 / 有哪些工具 / 你能做什么" → 注入权威工具清单（所有宿主通用）
        {
            const _capRaw = String(turn._originalUserMessage || turn.userMessage || turn.message || '');
            if (CAPABILITY_QUESTION_RE.test(_capRaw)) {
                const _inv = buildToolInventoryText();
                if (_inv) extraContext += '\n\n' + _inv;
            }
        }

        // 添加引用区域
        if (turn.context?.referenceRegions?.length > 0) {
            const refText = turn.context.referenceRegions
                .filter(r => r != null)
                .map((r, i) => `区域${i+1}: [${r.workbook || ''}]${r.worksheet || ''}!${r.address || ''}${r.preview ? '\n数据预览:\n' + r.preview : ''}`)
                .join('\n\n');
            extraContext += `\n\n【引用区域】\n${refText}`;
        }

        /**
     * 判断当前选区的全量数据（表头 + 单元格样本）是否应进入 LLM 上下文。
     * 默认只给轻量元信息（地址/表名），避免无关数据每轮泄漏进提示词、诱发模型臆造。
     * 仅当 消息明显引用选中区域 / 命中数据操作意图 / 已路由到作用于选区的技能域 时才带全量数据。
     */
    function shouldIncludeSelectionFull(turn) {
        if (!turn) return false;
        const msg = turn.userMessage || '';

        // 显式引用选中区域
        if (/(选中的|当前选|我选|所选|选区|选好的|这块|这段|这个区域|这里的|这个表|这张表|当前区域)/.test(msg)) return true;

        // 数据操作/分析类意图（作用于当前选中数据）
        if (/(计算|汇总|求和|合计|统计|平均|均值|图表|柱状|折线|饼图|美化|格式化|清洗|去重|拆分|合并|排序|筛选|透视|生成报告|分析)/.test(msg)) return true;

        // 已明确路由到技能域（技能作用于选中区域）
        if (turn.kind === 'skill') return true;
        const sk = turn.skillName || (turn.skill && (turn.skill.name || (typeof turn.skill === 'string' ? turn.skill : '')));
        if (sk) return true;

        return false;
    }

   // 添加当前选区（如果存在且不在引用区域中）
        const _hasExcelIntent = /(图表|饼图|柱图|折线图|柱状图|条形图|散点图|面积图|雷达图|直方图|趋势图|仪表盘|图|chart|graph|visuali[sz]e|汇总|统计|分析|总结|解读|洞察|评估|计算|求和|平均|最大|最小|计数|排名|百分比|分布|趋势|pivot|透视|筛选|排序|清洗|格式化|美化|合并|拆分|转置|去重|查找|替换|导出|导入|表格|数据|工作表|sheet|worksheet|单元格|列|行|公式|VBA|宏)/i.test(String(turn.userMessage || ''));

        if (getHost().isExcel && turn.context?.selection && (!turn.context?.referenceRegions || turn.context.referenceRegions.length === 0)) {
            const sel = turn.context.selection;
            const address = sel.address || sel.range || '';
            const workbookName = sel.workbookName || sel.workbook || '';
            const sheetName = sel.sheetName || sel.worksheet || '';
            const isAuto = sel.isAutoCaptured === true;

            // 【相关性门控】无关查询只带元信息；仅引用选区/数据类意图/技能域才带表头与样本数据
            const includeFull = shouldIncludeSelectionFull(turn);
            const headers = includeFull && Array.isArray(sel.headers) ? sel.headers : [];
            const rows = includeFull && Array.isArray(sel.rows) ? sel.rows : [];

            let selText = `【当前${isAuto ? '默认数据源（自动捕获）' : '选区'}】\n地址: ${address}`;
            if (workbookName) selText += `\n工作簿: ${workbookName}`;
            if (sheetName) selText += `\n工作表: ${sheetName}`;
            if (headers.length > 0) selText += `\n表头: ${headers.join(', ')}`;
            if (!includeFull) selText += '\n数据: (未加载完整数据，如需读取请调用 excel_read_range)';
            if (rows.length > 0) {
                selText += `\n样本数据 (${Math.min(rows.length, 5)}/${rows.length} 行):`;
                rows.slice(0, 5).forEach(function(row) {
                    const cells = Array.isArray(row) ? row : [row];
                    selText += '\n  ' + cells.join(' | ');
                });
            }
            if (isAuto) {
                selText += '\n⚠️  该选区是从活动工作表已用范围自动捕获的——用户没有手动选中任何单元格。';
                selText += '\n如果数据范围不对，请先在 Excel 中手动框选数据区域再告诉我。';
            }
            extraContext += '\n\n' + selText;
        } else if (getHost().isExcel && !turn.context?.selection && _hasExcelIntent) {
            // 【降级感知】有 Excel 操作意图但没选区 —— 告诉 LLM 真实环境，防止瞎编 range / sheet
            const _as = turn.context?.activeSheet;
            const asSheet = _as?.sheet || '';
            const asWb = _as?.workbook || '';

            const _rawUser = (turn && (turn._originalUserMessage || turn.userMessage)) || '';
            const _rangePlan = (App.ChatV2.rangeResolver && typeof App.ChatV2.rangeResolver.decideFormatTarget === 'function')
                ? App.ChatV2.rangeResolver.decideFormatTarget(_rawUser, turn.context)
                : null;
            if (_rangePlan && (_rangePlan.sheetScope === 'all' || _rangePlan.sheetScope === 'others') && _rangePlan.rangeMode !== 'skip') {
                extraContext += '\n\n【多张表，没有手选区域是正常的】\n'
                    + '不要要求用户先框选，也不要回答「当前无选区所以没法做」。\n'
                    + '先取工作表清单，再逐表 excel_format_range：每张表用自己的数据区。\n'
                    + '用户嘴里给了地址就用那个地址；只点中一个格子、或系统自动捕获的范围，都不是他圈定的区域。';
            }
            let downgradeText = `⚠️ 【当前无显式选区】\n用户没有在 Excel 中手动选中任何单元格区域。`;
            if (asSheet) downgradeText += `\n活动工作表: ${asSheet}`;
            if (asWb) downgradeText += `\n活动工作簿: ${asWb}`;
            downgradeText += `\n\n👉 处理建议：`;
            // 根据意图类型给出不同建议：分析类要先读，操作类直接用范围
            // 三态分流：分析类 → 先读数据；文件产出类 → 直接调导出/保存工具；其余纯操作 → 直接调工具。
            // 【v20261001 修复】旧实现只有"分析 / 纯操作(格式化…)"两态，导出/保存既不命中分析词也不命中
            // 操作词表 → 被当成分析类注入"先调用 excel_read_range"，模型读完整表就去做数据分析了，
            // 导出工具一次都没调用（"任务完成"却拿不到文件）。故补第三态。
            const _isAnalysis = detectAnalysisIntent(_rawUser);
            const _isOpOnly = !_isAnalysis && detectOperationOnlyIntent(_rawUser);
            const _isFileOut = !_isAnalysis && !_isOpOnly && detectFileOutputIntent(_rawUser);
            if (_isFileOut) {
                downgradeText += `\n1. 这是"文件产出"请求：导出/保存类工具会自己读取源表并落盘，【不需要先调 excel_read_range 把整表读进上下文】——直接调用对应工具即可（sheet/range 不传即作用于活动工作表）。`;
                downgradeText += `\n2. 工具选择：导出 CSV → excel_export_csv；保存工作簿 → excel_save_workbook；导出图表图片 → excel_export_chart_image；按工作表/行数拆分 → excel_split_workbook / excel_split_by_rows。`;
                downgradeText += `\n3. filePath / outputDir 留空即可，系统会自动落到真实桌面；严禁编造路径（如 C:\\Users\\Yourusername\\Desktop\\...）。`;
                downgradeText += `\n4. 必须真正调用工具完成落盘：不要只输出数据分析，也不要在没有工具调用的情况下声称"已完成"。`;
            } else if (_isOpOnly) {
                downgradeText += `\n1. 这类操作只需要目标范围，不需要先读数据内容——直接调用对应工具（不传 range 参数让后端自动使用 UsedRange，或用 Selection）即可。`;
                downgradeText += `\n2. 严禁编造 range/sheet 参数（如 "A1:C10"、"Sheet1"）——这会导致工具作用在错误位置。`;
            } else {
                downgradeText += `\n1. 先调用 excel_read_range 读取活动工作表的实际数据（不传 range 参数让后端自动使用 UsedRange，或你也可以先调一次读取看实际范围）。`;
                downgradeText += `\n2. 读完数据后，再基于真实数据决定是否需要聚合、出图、导出或其他操作。`;
                downgradeText += `\n3. 严禁在没有读过真实数据的情况下编造 range/sheet 参数（如 "A1:C10"、"Sheet1"）——这会导致工具读不到数据或读错数据。`;
            }
            downgradeText += `\n如果活动工作表完全是空的，直接告知用户"活动工作表中没有数据，请先填入或选择含有数据的工作表"，不要硬调用工具。`;

            const _multiSheet = _rangePlan && (_rangePlan.sheetScope === 'all' || _rangePlan.sheetScope === 'others') && _rangePlan.rangeMode !== 'skip';
            if (!_multiSheet) extraContext += '\n\n' + downgradeText;
        } else if (!getHost().isExcel) {
            // ===== Word / PowerPoint：把当前文档/幻灯片文本喂给 LLM，使其能直接总结/润色 =====
            const isPptHost = getHost().isPowerPoint;
            // 【范围标注】scope='all' 表示当前页定位失败、注入的是全部幻灯片文本——
            // 必须如实标注，否则"总结当前幻灯片"会被弱模型做成逐页罗列（实测 11 页被整篇抄出）。
            const scopeAll = turn.context?.documentScope === 'all';
            const label = isPptHost ? (scopeAll ? '全部幻灯片文本（当前页定位失败）' : '当前幻灯片') : '当前文档';
            // 选区文本优先（「将选中的内容翻译」）；无选区再退全文
            const _selTxt = (turn.context?.selectionText ? String(turn.context.selectionText).trim() : '');
            const _docRaw = (turn.context?.documentText ? String(turn.context.documentText).trim() : '');
            const _rawUserNonExcel = String(
                (turn && (turn._originalUserMessage || turn.userMessage || turn.message)) || '');
            const _wantsSelection = /(选中|所选|选区|这段|这块|当前选)/.test(_rawUserNonExcel);
            const docText = (_wantsSelection && _selTxt) ? _selTxt : (_selTxt || _docRaw);
            const contentLabel = (_wantsSelection && (_selTxt || docText))
                ? (isPptHost ? '当前选中文本' : '当前选中文本')
                : label;
            // 【元问题豁免】只有当用户消息确属"内容处理"任务（总结/润色/改写/翻译/提炼…）时，
            // 才注入"文档内容 / 读不到内容"的声明。否则纯问答类消息（"你有多少工具可用"
            // "列出工具名称"）也会被塞进"必须如实告知用户'当前文档中没有可读取的内容'"，
            // 弱模型会把这句指令直接照抄成回答（实测 Word 16:12：问工具数量 → 答
            // "当前文档中没有可读取的内容。请确认文档中有文字后再次尝试…"）。
            const _isContentTaskNonExcel = detectContentTaskIntent(_rawUserNonExcel);
            // 【素材按需注入】只有"用户这条消息确实要用到正文"时才注入幻灯片/文档文本：
            // ① 内容处理任务（总结/润色/改写/翻译/提炼…）；或
            // ② 消息本身指涉了宿主对象（幻灯片 / 文档 / 这一页 / 内容 / 讲了什么…）。
            // 否则纯元问题（"你有多少工具""你能做什么""怎么用"）也会被塞进整篇正文，
            // 弱模型会把素材当成任务直接总结掉（实测 PPT 16:16：问"你有多少工具" →
            // 答整份"2024年华南地区销售报表"总结，既不回答用户问题，又白烧 240 输出 token）。
            const _refersDocNonExcel = /幻灯片|演示文稿|ppt|pptx|文档|正文|全文|这一页|本页|这页|当前页|当前这|内容|讲了什么|讲了啥|写的是|包含什么|有哪些页/.test(_rawUserNonExcel.toLowerCase());
            // 【文件产出豁免】导出/另存为类请求（"将当前文档导出为 PDF"）只做落盘，不需要正文。
            // 若仍注入"当前文档没有可读取的文字内容"警告，弱模型会照抄成"没有内容所以无法导出"
            // （实测 2026-10-06 00:55：空白文档导出 PDF 被拒——导出本来就不需要任何可读内容）。
            const _isFileOutputNonExcel = detectFileOutputIntent(_rawUserNonExcel);
            const _needDocCtx = !_isFileOutputNonExcel && (_isContentTaskNonExcel || _refersDocNonExcel);
            let docCtx = '';
            if (_needDocCtx && docText) {
                docCtx = `【${contentLabel}内容（参考素材）】\n${docText}`;
                // 【防劫持】弱模型看到大段正文会默认"任务就是总结它"，必须显式声明素材的从属地位。
                docCtx += `\n（以上只是供参考的素材，仅当用户要求总结/润色/改写/翻译/提炼它时才使用；` +
                    `用户问的是其它事情（如可用工具、功能、操作方式）时请忽略素材，直接回答用户的问题。）`;
                if (_wantsSelection) {
                    docCtx += `\n⚡ 用户要求处理「选中内容」：请基于上方素材直接完成翻译/润色/改写，并可用 word_replace_selection / ppt 对应写入工具回填（若有）；不要声称读不到内容。`;
                }
                if (scopeAll) {
                    docCtx += `\n（注意：以上是全部${isPptHost ? '幻灯片' : '文档'}文本。若用户要求的是"当前页"，` +
                        `请给出全文的精炼整体总结，严禁逐${isPptHost ? '页罗列原文' : '段照抄'}。）`;
                }
            } else if (_needDocCtx && !docText && _isContentTaskNonExcel) {
                // 【防幻觉】仅在"确实要用正文"（内容任务：总结/润色/改写/翻译/提炼）且采集为空时声明，
                // 并对弱模型下"禁止编造"的硬约束。否则空上下文会让 qwen2.5 这类小模型
                // 凭标题（如"项目进度更新"）编出"项目A:已完成30%…"假正文。
                // 【范围收紧·实测 2026-10-06 00:55】此前条件是 _refersDocNonExcel，于是"将当前文档导出
                // 为 PDF"（仅因含"文档"二字命中）也被塞进本警告，弱模型原样照抄成
                // "当前文档中没有可读取的内容，因此无法导出 PDF"——把确定性导出操作说成做不到。
                // 现只有真·内容处理任务才允许声明"读不到内容"；其余指涉宿主对象的操作类请求
                // （导出/保存/美化等）一律不注入，避免把"无需正文"的任务误标为失败。
                // 另：label 本身已带"当前"（'当前文档'/'当前幻灯片'），此处不再重复拼"当前"。
                docCtx = `【${contentLabel}内容】\n（${contentLabel}没有可读取的文字内容）\n` +
                    `⚠️ 若用户要求总结/润色/改写/翻译/提炼${contentLabel}的内容，必须如实告知"${contentLabel}中没有可读取的内容"，` +
                    `严禁编造、猜测或凭标题虚构任何正文。`;
            }
            if (docCtx) {
                if (turn.context?.activeSheet?.workbook) {
                    docCtx += `\n\n文档: ${turn.context.activeSheet.workbook}`;
                }
                extraContext += '\n\n' + docCtx;
            }

            // 选中对象（表格/形状）：让 LLM 知道用户选中的是具体对象，可直接操作，
            // 无需再追问幻灯片/表格序号（这是“无选区”误判的根因）。
            const selObj = turn.context?.selectedObject;
            if (selObj && selObj.summary) {
                let objCtx = `【已选中对象】\n${selObj.summary}`;
                if (selObj.selectionType === 'table') {
                    const fmtTool = isPptHost ? 'ppt_format_selected_table_cell' : 'word_format_selected_table_cell';
                    objCtx += `\n可用工具：${fmtTool} —— 省略 row/col 即作用于整张表格（推荐）；也可传 row/col（从 1 开始）只改某一格。`;
                    if (selObj.rowCount) objCtx += `\n表格共 ${selObj.rowCount} 行 ${selObj.colCount} 列，第 1 行为表头。`;
                    objCtx += `\n⚡ 当用户要求“设置对齐/颜色/字体/美化表格/调整格式”等整体格式、且未指明具体单元格时，必须直接调用 ${fmtTool}（省略 row/col 令其整表生效），并自行填入合理默认值（例如 align=center、headerBold=true、headerFillColor=#D9D9D9）。严禁反问用户行号/列号/颜色值——用户已选中表格，应直接操作并简报结果。`;
                }
                extraContext += '\n\n' + objCtx;
            }
        }

        if (extraContext) {
            return `${userMessage}${extraContext}\n\n请根据以上上下文处理我的请求。`;
        }

        return userMessage;
    }

    // ── 暴露 API ──────────────────────────────────────────────────

    App.ChatV2.PromptBuilder = PromptBuilder;
    App.ChatV2.promptBuilder = {
        // PromptBuilder 类
        PromptBuilder,
        getDefaultBuilder,
        resetDefaultBuilder,

        // 向后兼容的原始 API
        buildSystemPrompt,
        buildMessages,
        buildUserMessage,
        pickPromptId,
        collectVariables,
        substituteVars,
        estimateTokens,
        getContextWindow,
        applyTokenBudget,
        // 意图路由·分析：供路由器复用
        detectAnalysisIntent,
        ANALYSIS_TASK_PROMPT,
        // 意图路由·文件产出（导出/保存）：供测试与路由器复用
        detectFileOutputIntent,
        FILE_OUTPUT_TASK_PROMPT,
        // 意图路由·内容处理（总结/润色/改写/翻译）：非 Excel 宿主专用，供测试与路由器复用
        detectContentTaskIntent,
        buildHostContentTaskPrompt,
        // 意图路由·文件产出（导出/保存）：非 Excel 宿主专用，供测试与路由器复用
        buildHostFileOutputTaskPrompt,
        // 意图路由·纯操作（格式化/样式/结构，无需先读数据）：供 orchestrator 复用
        detectOperationOnlyIntent,
        OPERATION_ONLY_TASK_PROMPT
    };
    ChatLog.info?.('CHAT-V2', '✅ promptBuilder loaded (PromptAssembly)');
})();