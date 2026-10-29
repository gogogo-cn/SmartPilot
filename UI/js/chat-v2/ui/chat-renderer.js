/**
 * ============================================================
 * chat-v2/ui/chat-renderer.js — L5 渲染器
 * ============================================================
 * 职责：
 *   - 订阅 SessionStore 状态变化
 *   - 渲染：用户消息气泡 + AI 消息气泡（思考 / 正文 / 工具调用 / 错误 / 取消）
 *   - 流式增量更新（thinking / content delta）
 *   - 工具调用卡片（按轮次分组，状态：pending / completed / failed）
 *   - 取消 / 错误作为状态而非崩溃
 *
 * 不做的事：
 *   - 不直接调 LLM、不组装 prompt、不管理 turn 生命周期
 *
 * 依赖：App.ChatV2.sessionStore, App.ChatV2.chatEntry.resetProcessingUI
 * 复用 CSS：message-wrapper, message-content, thinking-*, exec-tool-*, cancel-btn
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 【破图兜底】聊天气泡内任何 <img> 加载失败 → 退化成文字提示，避免浏览器"破图"占位符。
    // 资源加载错误不冒泡，必须在捕获阶段拦截；只处理聊天气泡内的图片。
    document.addEventListener('error', function (e) {
        const t = e && e.target;
        if (!t || t.tagName !== 'IMG') return;
        const bubble = t.closest && t.closest('.chat-bubble, .message, .message-content');
        if (!bubble) return;
        const note = (t.getAttribute('alt') || t.getAttribute('title') || '图片').trim();
        const span = document.createElement('span');
        span.className = 'broken-img-note';
        span.textContent = '🖼️ ' + note + '（图片加载失败）';
        if (t.parentNode) t.parentNode.replaceChild(span, t);
    }, true);

    // 每个活动的 turn 渲染句柄：turnId -> RenderHandle
    // RenderHandle = { aiBubble, thinkingEl, thinkingBodyEl, contentEl, toolsEl, dotsEl, rounds: Map }
    const _handles = new Map();

    // markdown 渲染节流
    const _pendingRender = new Map(); // turnId -> timer
    const RENDER_THROTTLE_MS = 80;

    // ============ 工具：转义 / DOM ============

    function escapeHtml(str) {
        if (str == null) return '';
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(str) : '';
    }

    function el(tag, className, html) {
        const e = document.createElement(tag);
        if (className) e.className = className;
        if (html != null) e.innerHTML = html;
        return e;
    }

    function getMessagesContainer() {
        return document.getElementById('chat-messages');
    }

    function autoScroll() {
        const c = getMessagesContainer();
        if (c) c.scrollTop = c.scrollHeight;
    }

    function hideWelcomePanel() {
        const c = getMessagesContainer();
        if (!c) return;
        const welcomePanel = c.querySelector('.welcome-panel');
        if (welcomePanel) {
            welcomePanel.remove();
            ChatLog.info?.('RENDERER', '[chat-renderer] 欢迎面板已隐藏');
        }
        const welcomeMessage = c.querySelector('.welcome-message');
        if (welcomeMessage) {
            welcomeMessage.remove();
        }
        // 番茄钟欢迎页（真实消息出现时一并移除）
        const pomodoroWelcome = c.querySelector('.pomodoro-welcome');
        if (pomodoroWelcome) {
            pomodoroWelcome.remove();
        }
        try { c.classList.remove('has-pomodoro'); } catch (_) { /* ignore */ }
    }

    // ============ Markdown 渲染（节流）============

    // 代码块检测/清洗逻辑拆至 renderers/code-block 纯函数子模块
    function detectAndFormatCodeBlocks(text) {
        return App.ChatV2.renderers.codeBlock.detectAndFormatCodeBlocks(text);
    }

    function renderMarkdown(text) {
        return App.ChatV2.renderers.markdown ? App.ChatV2.renderers.markdown.renderMarkdown(text) : '';
    }
    function renderTextPreservingCharts(text) {
        return App.ChatV2.renderers.markdown ? App.ChatV2.renderers.markdown.renderTextPreservingCharts(text) : text;
    }
    function sanitizeLeakedJsonSnippets(text) {
        return App.ChatV2.renderers.markdown ? App.ChatV2.renderers.markdown.sanitizeLeakedJsonSnippets(text) : text;
    }
    /**
     * 【形状无关兜底 · 2026-10-09】正文里只要嵌了"合并结果 JSON"（writtenTo / mergedRows+sourceSheets），
     * 不管是纯回声、带说明文字的回声、双层编码还是未知包装键，都整块换成可读摘要。
     * 只命中"后端合并工具专属字段"，用户点名要的普通 JSON 输出不受影响。
     */
    function sanitizeToolJsonEcho(text) {
        if (typeof text !== 'string' || text.indexOf('{') < 0) return text;
        try {
            const _tu = App.ChatV2 && App.ChatV2.toolRegistryUtil;
            const replaced = _tu && typeof _tu.replaceMergeJsonInText === 'function'
                ? _tu.replaceMergeJsonInText(text) : null;
            return replaced || text;
        } catch (e) {
            return text;
        }
    }
    function sanitizeFakeToolTags(text) {
        return App.ChatV2.renderers.markdown ? App.ChatV2.renderers.markdown.sanitizeFakeToolTags(text) : text;
    }

    /**
     * 把 LLM 输出的 "key=value | key=value | ..." 行自动转成 Markdown 表格。
     * 兜底：Prompt 要求用 Markdown 表格但 LLM 没听时的后处理。
     * 识别条件：连续 2+ 行，每行含 2+ 个 key=value 用 " | " 分隔。
     * 已有 Markdown 表格的行（| --- |）不会被误转。
     */
    function enhanceKeyValueText(text) {
        if (!text || typeof text !== 'string') return text;
        const lines = text.split(/[\r\n]+/);
        // 找出所有"可转换行"（连续 2+ 行，每行含 2+ 个 key=value 用 | 分隔）
        const blocks = [];
        let i = 0;
        while (i < lines.length) {
            const row = parseKvLine(lines[i]);
            if (!row) { i++; continue; }
            let start = i, count = 1;
            while (i + count < lines.length && parseKvLine(lines[i + count])) count++;
            if (count >= 2) { blocks.push({ start, count }); i += count; } else { i++; }
        }
        if (blocks.length === 0) return text;
        // 从后往前替换，避免索引偏移
        const out = lines.slice();
        for (let b = blocks.length - 1; b >= 0; b--) {
            const { start, count } = blocks[b];
            const rows = [];
            for (let j = start; j < start + count; j++) rows.push(parseKvLine(lines[j]));
            const keys = rows[0].map(r => r.key);
            let md = '| ' + keys.join(' | ') + ' |';
            md += '\n| ' + keys.map(() => '---').join(' | ') + ' |';
            for (const row of rows) {
                md += '\n| ' + row.map(r => r.val.replace(/\|/g, '\\|')).join(' | ') + ' |';
            }
            out.splice(start, count, md);
        }
        return out.join('\n');

        function parseKvLine(line) {
            const trimmed = (line || '').trim();
            if (!trimmed) return null;
            // 跳过 Markdown 表格头行
            if (/^\s*\|?\s*:?-+/.test(trimmed)) return null;
            const eqCount = (trimmed.match(/=/g) || []).length;
            const pipeCount = (trimmed.match(/\|/g) || []).length;
            if (eqCount < 2 || pipeCount < 1) return null;
            const parts = trimmed.split(/\s*\|\s*/).filter(p => p.trim());
            if (parts.length < 2) return null;
            const result = [];
            for (const p of parts) {
                const eqIdx = p.indexOf('=');
                if (eqIdx <= 0) return null;
                const key = p.substring(0, eqIdx).trim();
                const val = p.substring(eqIdx + 1).trim();
                if (!key) return null;
                result.push({ key, val });
            }
            return result;
        }
    }
    function isCodeTurn(turn) {
        return App.ChatV2.renderers.markdown ? App.ChatV2.renderers.markdown.isCodeTurn(turn) : !!(turn && turn._codeLang);
    }

    function scheduleContentRender(turnId) {
        if (_pendingRender.has(turnId)) return;
        const t = setTimeout(() => {
            _pendingRender.delete(turnId);
            flushContentRender(turnId);
        }, RENDER_THROTTLE_MS);
        _pendingRender.set(turnId, t);
    }



    function flushContentRender(turnId) {
        const handle = _handles.get(turnId);
        if (!handle) return;
        const turn = App.ChatV2.sessionStore?.getCurrentTurn();
        if (!turn || turn.id !== turnId) return;
        if (handle.contentEl) {
            // 重渲前保存过程列表各行的展开态（流式全量重渲会重置 <details>），渲染后恢复
            const _cbxStates = collectCbxOpenStates(handle.contentEl);
            // thoughtList 迭代替换块独立于 content/_codeLang 分支 —— 永远先渲
            // 纯对话（无工具/步骤执行，如寒暄、知识问答）不渲染过程列表
            let thoughtHtml = hasRealExecution(turn) ? renderThoughtListHtml(turn) : '';
            // 过程列表一旦有可见内容（意图分析行/工具行），移除 loading 三个点——
            // 三个点只在"完全没有任何可见进展"时才有意义，有活动行后继续挂着会误导用户
            if (thoughtHtml && handle.dotsEl && handle.dotsEl.parentNode) {
                handle.dotsEl.remove();
                handle.dotsEl = null;
            }
            // 【三态渲染统一·流式】thoughtList 为空但有真实工具调用时（快速路径不产生 thought 块），
            // 复用 cbx 列表让过程在流式期也渐进可见；toolsEl 已有实时工具卡时跳过，避免同一过程双份展示
            if (!thoughtHtml && hasRealExecution(turn)) {
                const _streamCalls = Array.isArray(turn.toolCalls)
                    ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
                const _hasDomCards = handle.toolsEl && handle.toolsEl.querySelector('.exec-round, .cbx-item');
                if (!_hasDomCards && _streamCalls.some(tc =>
                    tc.status === 'completed' || tc.status === 'failed' || tc.status === 'running')) {
                    thoughtHtml = buildCbExecListHtml(_streamCalls, turnUserText(turn));
                }
            }
            // 【终态折叠】turn 已结束时把过程列表折叠为汇总盒（流式期原样平铺）
            thoughtHtml = wrapProcessSummaryBox(turn, thoughtHtml);

            if (turn._codeLang) {
                renderCodeOnly(handle.contentEl, turn.content || '', turn._codeLang, thoughtHtml);
                // 代码 turn 同样恢复过程列表展开态；running 行推理内容贴底跟随最新内容
                restoreCbxOpenStates(handle.contentEl, _cbxStates);
                // 【thoughtList + 图表 + JSON 表格】即使是代码 turn，thoughtList 也渲出来；
                // 如果代码块里有 data-echarts 或 JSON，也走后处理
                enhanceJsonTables(handle.contentEl);
                initInlineCharts(handle.contentEl);
            } else {
                const rawContent = turn.content || '';
                // 【关键修复】先清洗 LLM 总结里误抄进来的 tool_result JSON / CSS 字典尾巴（"代码溢出"）
                // 再清除残留的"仿造工具标签"（虚报没以该名执行的工具，如 [excel-operation__xxx]）
                let cleanContent = sanitizeFakeToolTags(sanitizeLeakedJsonSnippets(rawContent));
                // 【格式兜底】把 LLM 输出的 "key=value | key=value | ..." 行自动转成 Markdown 表格
                cleanContent = enhanceKeyValueText(cleanContent);
                // 通用对话里模型可能"只写计划不真执行"：把正文中【执行】[工具] 标注提取为"规划态"步骤，
                // 并把原始标注清洗成干净文本；无真实执行步骤时用解析出的规划步骤兜底渲染绿面板
                const prepared = prepareConversationContent(turn, cleanContent);
                // 分离可折叠面板内容与普通内容，仅对普通内容执行 Markdown
                const processed = processCollapsibleContent(prepared.text);
                // 将面板占位符替换为可折叠面板 HTML
                let renderedHtml = injectCollapsiblePanels(processed.html, processed.panels);
                // 旧版绿面板（exec-plan-card / renderStepsPlanHtml / _planHtml / _steps）已全量废弃 ——
                // thoughtList 的迭代替换模型统一承载所有步骤展示。

                // 【迭代替换模型 v20260925】thoughtList + 正文 Markdown 混排
                if (thoughtHtml) renderedHtml = thoughtHtml + renderedHtml;

                // 【Skill 卡片落地到"最终结果"】agent-loop 直调 skill.SkillCreator 时，卡片数据只在
                // 工具结果里（turn.toolCalls[].result），正文只有 LLM 的总结文本——卡片被埋在中途的工具区。
                // 这里在正文末尾统一补渲染卡片；若正文本身已带卡片 JSON（skill-pipeline 路径，
                // findSkillCardJson 已渲染出 .skill-card），检测到即跳过，避免同一卡片出现两次。
                try {
                    if (renderedHtml.indexOf('class="skill-card"') < 0) {
                        const _cardHtml = buildSkillCardFromTurnTools(turn);
                        if (_cardHtml) renderedHtml += _cardHtml;
                    }
                } catch (e) { /* 卡片兜底失败不影响正文渲染 */ }

                // 【Skill 脚本卡落地到"最终结果"】GenerateOCRParseScript 等技能：outputs.content 是
                // 原始 Markdown 解析脚本（不是 skill_card JSON，上面的渲染器不认）。这里用同一处钩子、
                // 由技能自身 uiActions 驱动渲染脚本盒子 + 复制/保存/重新生成 按钮，与 skill.SkillCreator
                // 的卡片走同一渲染链路，取代底部与本技能无关的通用 chips（调整图表/做成PPT）。
                try {
                    if (renderedHtml.indexOf('ocr-script-card') < 0) {
                        const _scriptHtml = buildSkillScriptCardFromTurnTools(turn);
                        if (_scriptHtml) renderedHtml += _scriptHtml;
                    }
                } catch (e) { /* 脚本卡兜底失败不影响正文渲染 */ }

                handle.contentEl.innerHTML = renderedHtml;
                restoreCbxOpenStates(handle.contentEl, _cbxStates);
                beautifyCodeBlocks(handle.contentEl);
                // 绑定折叠面板交互
                bindCollapsibleEvents(handle.contentEl);
                // 尝试将 JSON 结果转换为表格，并为所有表格添加独立的插入按钮
                enhanceJsonTables(handle.contentEl);
                // 扫描 content 内 data-echarts 节点并初始化为 ECharts
                initInlineCharts(handle.contentEl);
            }
            // 过程列表已经画在正文顶部时，收起工具区，避免同一步骤出现两份。
            if (thoughtHtml && handle.toolsEl) handle.toolsEl.style.display = 'none';
            // 终态兜底：命中展示诉求时强制展开过程盒与命中行
            enforceViewHintOpen(handle.contentEl, turn);
            autoScroll();
        }
    }


    // 剥代码块围栏逻辑拆至 renderers/code-block 纯函数子模块
    function stripCodeFences(code) {
        return App.ChatV2.renderers.codeBlock.stripCodeFences(code);
    }

    function renderCodeOnly(container, code, lang, thoughtHtml) {
        if (!container) return;
        const cleanCode = stripCodeFences(code);
        const safeCode = escapeHtml(cleanCode);
        const safeLang = escapeHtml(lang || 'text');
        let html = thoughtHtml || '';
        html += `<pre><code class="language-${safeLang}">${safeCode}</code></pre>`;
        container.innerHTML = html;
        beautifyCodeBlocks(container);
    }

    // ============ 代码块美化渲染 ============

    const LANG_DISPLAY = {
        vba: 'VBA',
        vb: 'VBA',
        visualbasic: 'VBA',
        python: 'Python',
        py: 'Python',
        javascript: 'JavaScript',
        js: 'JavaScript',
        typescript: 'TypeScript',
        ts: 'TypeScript',
        json: 'JSON',
        html: 'HTML',
        css: 'CSS',
        sql: 'SQL',
        bash: 'Bash',
        shell: 'Bash',
        powershell: 'PowerShell',
        csharp: 'C#',
        cs: 'C#',
        java: 'Java',
        cpp: 'C++',
        cplusplus: 'C++',
        ruby: 'Ruby',
        php: 'PHP',
        go: 'Go',
        rust: 'Rust',
        c: 'C',
        plaintext: '文本',
        text: '文本',
        txt: 'TXT',
        xml: 'XML',
        yaml: 'YAML',
        yml: 'YAML',
        markdown: 'Markdown',
        md: 'Markdown'
    };

    const EXECUTABLE_LANGS = { vba: true, vb: true, python: true, py: true, javascript: true, js: true, bash: true, shell: true, powershell: true };

    function beautifyCodeBlocks(container) {
        if (!container) return;
        const pres = container.querySelectorAll('pre');
        pres.forEach(function(pre) {
            if (pre.closest('.code-block-wrapper')) return;
            const codeEl = pre.querySelector('code');
            if (!codeEl) return;

            const lang = detectLanguage(codeEl, pre);
            const rawCode = codeEl.textContent;

            // 流式期间未闭合的代码围栏会渲染成"只有工具条没有内容"的空代码卡，
            // 直接隐藏；后续 flush 内容到位后重新渲染自然恢复显示
            if (!rawCode || !rawCode.trim()) {
                pre.style.display = 'none';
                return;
            }

            const wrapper = document.createElement('div');
            wrapper.className = 'code-block-wrapper';

            // Toolbar
            const toolbar = document.createElement('div');
            toolbar.className = 'code-block-toolbar';

            const langLabel = LANG_DISPLAY[lang] || (lang ? lang.toUpperCase() : '代码');
            const langClass = 'lang-' + (lang || 'unknown');
            const langSpan = document.createElement('span');
            langSpan.className = 'code-lang ' + langClass;
            langSpan.innerHTML = '<span class="lang-dot"></span>' + escapeHtml(langLabel);
            toolbar.appendChild(langSpan);

            // Actions
            const actions = document.createElement('div');
            actions.className = 'code-actions';

            // Execute button
            if (EXECUTABLE_LANGS[lang]) {
                const execBtn = document.createElement('button');
                execBtn.className = 'code-action-btn execute-btn';
                execBtn.innerHTML = '▶ 执行';
                execBtn.title = '执行此代码';
                execBtn.addEventListener('click', function() {
                    executeCodeBlock(rawCode, lang, execBtn, wrapper);
                });
                actions.appendChild(execBtn);
            }

            // Copy button
            const copyBtn = document.createElement('button');
            copyBtn.className = 'code-action-btn copy-btn';
            copyBtn.innerHTML = '📋 复制';
            copyBtn.title = '复制代码到剪贴板';
            copyBtn.addEventListener('click', function() {
                copyCode(rawCode, copyBtn);
            });
            actions.appendChild(copyBtn);

            // Download button（纯前端 Blob 下载：代码块/JSON 块存成文件）
            const LANG_EXT = { python: 'py', javascript: 'js', typescript: 'ts', json: 'json', html: 'html', sql: 'sql', csharp: 'cs', java: 'java', powershell: 'ps1', bash: 'sh', vba: 'vba', markdown: 'md', yaml: 'yml', xml: 'xml', css: 'css', csv: 'csv' };
            const dlBtn = document.createElement('button');
            dlBtn.className = 'code-action-btn download-btn';
            dlBtn.innerHTML = '⬇ 下载';
            dlBtn.title = '下载代码为文件';
            dlBtn.addEventListener('click', function () {
                const ext = LANG_EXT[lang] || 'txt';
                const ts = (function () { const d = new Date(); const p = function (n) { return String(n).padStart(2, '0'); }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); })();
                downloadText('snippet_' + ts + '.' + ext, rawCode, 'text/plain;charset=utf-8');
            });
            actions.appendChild(dlBtn);

            // Edit button
            const editBtn = document.createElement('button');
            editBtn.className = 'code-action-btn edit-btn';
            editBtn.innerHTML = '✏️ 编辑';
            editBtn.title = '在代码编辑器中编辑';
            editBtn.addEventListener('click', function() {
                sendToEditor(rawCode, lang, langLabel);
            });
            actions.appendChild(editBtn);

            // Collapse button (after edit)
            const collapseBtn = document.createElement('button');
            collapseBtn.className = 'code-action-btn collapse-btn';
            collapseBtn.innerHTML = '▾';
            collapseBtn.title = '折叠/展开';
            collapseBtn.addEventListener('click', function() {
                const isCollapsed = wrapper.classList.toggle('collapsed');
                collapseBtn.innerHTML = isCollapsed ? '▸' : '▾';
            });
            actions.appendChild(collapseBtn);

            toolbar.appendChild(actions);
            wrapper.appendChild(toolbar);

            // Code content
            const content = document.createElement('div');
            content.className = 'code-block-content';

            const codePre = document.createElement('pre');
            const codeTag = document.createElement('code');
            codeTag.className = 'language-' + (lang || '');
            codeTag.textContent = rawCode;

            // Apply syntax highlighting
            const highlighted = highlightSyntax(rawCode, lang);
            if (highlighted !== null) {
                codeTag.innerHTML = highlighted;
            }

            codePre.appendChild(codeTag);
            content.appendChild(codePre);
            wrapper.appendChild(content);

            // Replace original pre
            pre.parentNode.replaceChild(wrapper, pre);
        });
        // JSON 块（renderCollapsibleJson 产出的 <pre class="json-block">，无 <code> 子节点）
        // 不会被上面的循环接管，这里单独给它补"复制 / 下载"工具栏
        try { enhanceJsonBlocks(container); } catch (e) { /* 工具栏失败不影响代码块 */ }
    }

    /** 【JSON 块工具栏】给"查看原始结果（JSON）"那块裸 <pre class="json-block"> 补一行
     *  复制 / 下载 按钮。复用 .code-block-toolbar 的深色样式（与代码块工具栏同一套 CSS，
     *  无需新增样式），插在 <pre> 之前、<summary> 之后，随 <details> 展开时才可见。 */
    function enhanceJsonBlocks(container) {
        if (!container) return;
        container.querySelectorAll('pre.json-block').forEach(function (pre) {
            if (pre.closest('.code-block-wrapper')) return;   // 已按代码块渲染过，跳过
            const box = pre.parentElement;
            if (box && box.querySelector(':scope > .code-block-toolbar[data-json-tools]')) return;  // 幂等
            const raw = pre.textContent || '';

            const bar = document.createElement('div');
            bar.className = 'code-block-toolbar';
            bar.setAttribute('data-json-tools', '1');
            bar.style.marginTop = '6px';

            const langSpan = document.createElement('span');
            langSpan.className = 'code-lang lang-json';
            langSpan.innerHTML = '<span class="lang-dot"></span>JSON';
            bar.appendChild(langSpan);

            const actions = document.createElement('div');
            actions.className = 'code-actions';

            const copyBtn = document.createElement('button');
            copyBtn.className = 'code-action-btn copy-btn';
            copyBtn.innerHTML = '📋 复制';
            copyBtn.title = '复制 JSON 到剪贴板';
            copyBtn.addEventListener('click', function () { copyCode(raw, copyBtn); });
            actions.appendChild(copyBtn);

            const dlBtn = document.createElement('button');
            dlBtn.className = 'code-action-btn download-btn';
            dlBtn.innerHTML = '⬇ 下载';
            dlBtn.title = '下载为 .json 文件';
            dlBtn.addEventListener('click', function () {
                const ts = (function () { const d = new Date(); const p = function (n) { return String(n).padStart(2, '0'); }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); })();
                downloadText('json_' + ts + '.json', raw, 'application/json;charset=utf-8');
            });
            actions.appendChild(dlBtn);

            bar.appendChild(actions);
            pre.parentNode.insertBefore(bar, pre);
        });
    }

    function detectLanguage(codeEl, pre) {
        // Check class on code element
        const classes = (codeEl.className || pre.className || '').toLowerCase();
        const langMatch = classes.match(/language-([\w-]+)/);
        if (langMatch) return langMatch[1].toLowerCase();

        // Check data attribute
        const dataLang = codeEl.getAttribute('data-language') || pre.getAttribute('data-language');
        if (dataLang) return dataLang.toLowerCase();

        // Auto-detect
        const code = codeEl.textContent;
        return autoDetectLanguage(code);
    }

    function autoDetectLanguage(code) {
        if (!code) return '';
        const firstLine = code.trimStart().split('\n')[0] || '';

        // Python 检测优先级最高（避免被误判为 VBA）
        if (/\b(def |import |from |print\(|@\w+|pip |pandas|numpy|DataFrame|plt\.|fig,|ax=)\b/.test(code)) return 'python';
        if (/^\s*(import |from |def |class |if __name__|print\()/.test(firstLine)) return 'python';

        // HTML
        if (/^\s*<!DOCTYPE html|<html[\s>]/i.test(code)) return 'html';

        // JSON
        if (/^\s*\{[\s\S]*\}\s*$/.test(code)) {
            try { JSON.parse(code); return 'json'; } catch(e) { ChatLog.debug?.('UI', '[chat-renderer] 语言探测非 JSON，继续匹配:', e); }
        }

        // JavaScript
        if (/\b(function |const |let |var |=>|require\(|module\.exports|console\.log)\b/.test(code)) return 'javascript';

        // TypeScript
        if (/\b(async |await |interface |implements)\b/.test(code)) return 'typescript';

        // SQL
        if (/\b(SELECT\s+.*FROM|INSERT\s+INTO|UPDATE\s+.*SET|DELETE\s+FROM|CREATE\s+TABLE)\b/i.test(code)) return 'sql';

        // C#
        if (/^\s*(#include|using\s+System|namespace\s+\w+|public\s+class)/.test(code)) return 'csharp';
        if (/\b(using\s+System|namespace|class\s+\w+)/.test(code)) return 'csharp';

        // Java
        if (/^\s*(public|private|protected).*class\s+\w+/.test(code)) return 'java';

        // PowerShell
        if (/\b(Get-\w+|Set-\w+|Invoke-\w+)\b/.test(code)) return 'powershell';

        // Bash
        if (/^\s*(echo|ls|cd|mkdir|rm|cat|grep|awk|sed)\b/.test(firstLine)) return 'bash';

        // VBA 检测放在最后（避免误判 Python 代码）
        if (/^(Sub|Function|Dim|Public|Private|Const|Type)\s/i.test(firstLine) ||
            /\b(Sub|Function)\s+\w+\s*\(/.test(code)) {
            return 'vba';
        }
        if (/\b(SUB|FUNCTION|DIM|END\s+Sub|END\s+Function|NEXT|FOR)\b/i.test(code)) return 'vba';

        return '';
    }

    function asRunnableVba(code, lang) {
        const text = code == null ? '' : String(code);
        const looksVba = lang === 'vba' || lang === 'vb' ||
            (/\b(?:Sub|Function)\s+\S+/i.test(text) && /\bEnd\s+(?:Sub|Function)\b/i.test(text));
        if (!looksVba) return text;
        const engine = window.App && App.ChatV2 && App.ChatV2.services && App.ChatV2.services.vbaTemplateEngine;
        if (!engine || typeof engine.toRunnableVba !== 'function') return text;
        try {
            const cleaned = engine.toRunnableVba(text);
            return cleaned || text;
        } catch (_) {
            return text;
        }
    }

    function copyCode(code, btn) {
        try {
            code = asRunnableVba(code);
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(code).then(function() {
                    const orig = btn.innerHTML;
                    btn.innerHTML = '✅ 已复制';
                    btn.style.color = '#a6e3a1';
                    setTimeout(function() {
                        btn.innerHTML = orig;
                        btn.style.color = '';
                    }, 1500);
                });
            } else {
                const ta = document.createElement('textarea');
                ta.value = code;
                document.body.appendChild(ta);
                ta.select();
                document.execCommand('copy');
                document.body.removeChild(ta);
                const orig = btn.innerHTML;
                btn.innerHTML = '✅ 已复制';
                setTimeout(function() { btn.innerHTML = orig; }, 1500);
            }
        } catch (e) {
            App.showToast?.('复制失败: ' + e.message, 'error');
        }
    }

    /** 纯前端下载文本文件（Blob + 临时 <a download>）。WebView2 下落到默认下载目录。 */
    function downloadText(filename, text, mime) {
        try {
            const blob = new Blob([text == null ? '' : String(text)], { type: mime || 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = filename;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
            App.showToast?.('已开始下载：' + filename, 'info');
        } catch (e) {
            App.showToast?.('下载失败：' + (e && e.message ? e.message : e), 'error');
        }
    }

    function tableRowsToCsv(rows) {
        if (!rows || !rows.length) return '';
        return rows.map(function (row) {
            return (row || []).map(function (cell) {
                const s = cell == null ? '' : String(cell);
                if (/[",\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
                return s;
            }).join(',');
        }).join('\r\n');
    }

    function sendToEditor(code, lang, langLabel) {
        const editor = document.getElementById('code-editor');
        if (!editor) {
            App.showToast?.('代码编辑器未就绪', 'error');
            return;
        }

        const langMap = { vba: 'vba', vb: 'vba', python: 'python', py: 'python', javascript: 'javascript', js: 'javascript', typescript: 'javascript', ts: 'javascript' };
        const mappedLang = langMap[lang] || 'vba';
        if (lang === 'vba' || lang === 'vb') code = asRunnableVba(code, lang);

        // Expand editor if collapsed
        const leftPanel = document.getElementById('left-panel');
        if (leftPanel && !leftPanel.classList.contains('expanded')) {
            if (typeof window.toggleEditor === 'function') {
                window.toggleEditor();
            }
        }

        // 切换编辑器语言（同步内部 editorLanguage 状态并刷新 UI）
        if (typeof window.switchEditorLanguage === 'function') {
            window.switchEditorLanguage(mappedLang);
        } else {
            const langSelect = document.getElementById('editor-language');
            if (langSelect) langSelect.value = mappedLang;
        }

        // Set code
        editor.value = code;
        editor.placeholder = `在此编辑 ${langLabel || 'VBA'} 代码...`;

        // 立即刷新行号 / 语法高亮 / 折叠按钮
        // 关键：textarea 文字为透明色（color: transparent），完全依赖 #code-highlight-content
        // 高亮层显示内容；若不刷新高亮层，代码就会"隐形"
        App.updateLineNumbers?.();
        App.updateSyntaxHighlight?.();
        App.updateFoldButtons?.();

        // Focus editor
        setTimeout(() => {
            editor.focus();
            editor.setSelectionRange(editor.value.length, editor.value.length);
            // 面板展开动画后可能需要再次同步尺寸/滚动
            App.updateLineNumbers?.();
            App.updateSyntaxHighlight?.();
            if (typeof App.initLineNumbers === 'function') {
                App.initLineNumbers();
            }
        }, 150);

        App.showToast?.(`已载入${langLabel || '代码'}到编辑器`, 'success');
    }

    async function executeCodeBlock(code, lang, btn, wrapper) {
        if (!App.excelBridge) {
            App.showToast?.('Excel 未连接', 'error');
            return;
        }

        btn.classList.add('executing');
        btn.disabled = true;
        const origText = btn.innerHTML;
        btn.innerHTML = '⏳ 执行中...';

        try {
            let result;
            const startTime = performance.now();

            if (lang === 'vba' || lang === 'vb') {
                const macroName = 'GenMacro_' + Date.now().toString().slice(-6);
                result = await App.excelBridge.ExecuteVBA(asRunnableVba(code, lang), macroName);
            } else if (lang === 'python' || lang === 'py') {
                result = await App.excelBridge.ExecutePython(code);
            } else if (lang === 'javascript' || lang === 'js') {
                result = await App.excelBridge.ExecuteVBA(
                    'Sub RunJS()\n' +
                    '  Dim scriptControl As Object\n' +
                    '  Set scriptControl = CreateObject("MSScriptControl.ScriptControl")\n' +
                    '  scriptControl.Language = "JScript"\n' +
                    '  scriptControl.AddCode "' + code.replace(/"/g, '""').replace(/\n/g, '" & Chr(10) & "') + '"\n' +
                    'End Sub',
                    'RunJS'
                );
            } else {
                App.showToast?.('不支持执行此语言', 'warning');
                btn.classList.remove('executing');
                btn.disabled = false;
                btn.innerHTML = origText;
                return;
            }

            const elapsed = ((performance.now() - startTime) / 1000).toFixed(2);

            // Display result
            showExecutionResult(wrapper, result, elapsed, true);

            btn.classList.remove('executing');
            btn.disabled = false;
            btn.innerHTML = origText;

        } catch (e) {
            showExecutionResult(wrapper, e.message || String(e), '?', false);
            btn.classList.remove('executing');
            btn.disabled = false;
            btn.innerHTML = origText;
        }
    }

    function showExecutionResult(wrapper, result, elapsed, success) {
        const existing = wrapper.querySelector('.code-execution-result');
        if (existing) existing.remove();

        const resultEl = document.createElement('div');
        resultEl.className = 'code-execution-result ' + (success ? 'success' : 'error');

        const label = document.createElement('div');
        label.className = 'exec-label';
        label.textContent = (success ? '✅ 执行成功' : '❌ 执行失败') + ' · 耗时 ' + elapsed + 's';
        resultEl.appendChild(label);

        let displayText = '';
        try {
            if (typeof result === 'string') {
                try {
                    const parsed = JSON.parse(result);
                    displayText = JSON.stringify(parsed, null, 2);
                } catch(e) {
                    displayText = result;
                }
            } else if (typeof result === 'object') {
                displayText = JSON.stringify(result, null, 2);
            } else {
                displayText = String(result);
            }
        } catch(e) {
            displayText = String(result);
        }

        resultEl.appendChild(document.createTextNode(displayText));
        wrapper.appendChild(resultEl);
    }

    // Syntax highlighting modularized to renderers/syntax-highlight.js
    function highlightSyntax(code, lang) {
        if (App.ChatV2.renderers && App.ChatV2.renderers.syntaxHighlighter) {
            return App.ChatV2.renderers.syntaxHighlighter.highlightSyntax(code, lang);
        }
        return null;
    }

    // ============ 用户消息气泡 ============

    function renderUserBubble(turn) {
        const container = getMessagesContainer();
        if (!container) return;

        const wrapper = el('div', 'message-wrapper user');
        wrapper.dataset.turnId = turn.id;

        const avatar = el('div', 'message-avatar', '👤');
        const main = el('div', 'message-main');

        const header = el('div', 'message-header');
        header.appendChild(el('span', 'message-role', '我'));
        header.appendChild(el('span', 'message-time', formatTime(turn.startedAt)));
        main.appendChild(header);

        // 正文（优先使用原始消息，避免 OCR/富文本 pipeline 注入的提示词污染）
        const displayMessage = turn._originalUserMessage ?? turn.userMessage;
        if (displayMessage) {
            main.appendChild(el('div', 'message-content', escapeHtml(displayMessage)));
        }

        // 附件展示
        if (turn.attachments && turn.attachments.length > 0) {
            const attachDiv = el('div', 'attachment-list');
            for (const a of turn.attachments) {
                if (a.type === 'image' && a.base64) {
                    const img = el('div', 'image-container');
                    img.innerHTML = `<img src="data:${a.mime || 'image/png'};base64,${a.base64}" alt="${escapeHtml(a.fileName || '')}" style="max-width:200px;max-height:160px;border-radius:6px;">`;
                    attachDiv.appendChild(img);
                } else {
                    const icon = fileIcon(a.fileName);
                    const file = el('div', 'attachment-file-chip', `<span>${icon}</span><span>${escapeHtml(a.fileName || '')}</span>`);
                    attachDiv.appendChild(file);
                }
            }
            main.appendChild(attachDiv);
        }

        // 引用区域展示
        if (turn.context?.referenceRegions?.length > 0) {
            const refDiv = el('div', 'reference-chips');
            turn.context.referenceRegions.forEach((r, i) => {
                const addr = r.address || `[${r.workbook || r.workbookName || ''}]${r.worksheet || r.sheetName || ''}!${r.address || ''}`;
                refDiv.appendChild(el('span', 'reference-chip', `🔗 ${escapeHtml(addr)}`));
            });
            main.appendChild(refDiv);
        }

        // 用户消息操作按钮（编辑）
        const actions = el('div', 'message-actions user-actions');
        const editBtn = el('button', 'message-action-btn edit-btn', '✏️ 编辑');
        editBtn.addEventListener('click', () => editUserMessage(turn.id));
        actions.appendChild(editBtn);
        main.appendChild(actions);

        wrapper.appendChild(avatar);
        wrapper.appendChild(main);
        container.appendChild(wrapper);
        autoScroll();
    }

    // ============ 编辑用户消息 ============

    function editUserMessage(turnId) {
        const sessionStore = App.ChatV2.sessionStore;
        const turn = sessionStore.getTurns().find(t => t.id === turnId);
        if (!turn) {
            ChatLog.warn?.('RENDERER', '[editUserMessage] turn not found', turnId);
            return;
        }

        // 1. 如果有正在进行的 turn，先取消
        if (App.ChatV2.orchestrator?.isRunning()) {
            App.ChatV2.orchestrator.cancelCurrent('edit_message');
        }

        // 2. 恢复文本到输入框（优先使用原始用户消息）
        const textarea = document.getElementById('message-input');
        const originalMsg = turn._originalUserMessage ?? turn.userMessage;
        if (textarea && originalMsg) {
            textarea.value = originalMsg;
            textarea.style.height = 'auto';
            textarea.style.height = Math.min(textarea.scrollHeight, 200) + 'px';
            textarea.focus();
        }

        // 3. 在 send-btn 上标记编辑模式
        const sendBtn = document.getElementById('send-btn');
        if (sendBtn) {
            sendBtn.dataset.editingTurnId = turnId;
        }

        // 4. 恢复附件（如果有）— 同时渲染预览 UI
        if (turn.attachments && turn.attachments.length > 0) {
            const attachData = turn.attachments.map(a => ({
                fileName: a.fileName,
                type: a.type,
                mime: a.mime,
                size: a.size
            }));
            sendBtn.dataset.editingAttachments = JSON.stringify(attachData);

            const images = turn.attachments.filter(a => a.type === 'image' && a.base64);
            if (images.length > 0) {
                sendBtn.dataset.editingImageBase64 = JSON.stringify(images.map(img => ({
                    fileName: img.fileName,
                    base64: img.base64,
                    mime: img.mime
                })));
            }

            // 4.1 在预览区渲染附件 UI
            renderEditingAttachmentPreview(turn.attachments);
        } else {
            // 清除预览区
            const preview = document.getElementById('attachment-preview');
            if (preview) preview.style.display = 'none';
        }

        // 5. 从 session store 中移除该 turn 及之后的所有记录
        sessionStore.removeTurnsSince(turnId);

        // 6. 移除 DOM 中该 turn 及之后的所有消息 wrapper
        const container = getMessagesContainer();
        if (container) {
            const wrapper = container.querySelector(`.message-wrapper.user[data-turn-id="${turnId}"]`);
            if (wrapper) {
                let current = wrapper;
                while (current) {
                    const next = current.nextElementSibling;
                    current.remove();
                    current = next;
                }
            }
        }

        // 7. 清理 _handles 中对应 turn 的渲染句柄
        for (const [id, handle] of _handles) {
            const removedIds = [turnId, ...(sessionStore.getTurns().length > 0 ? [] : [])];
            // 简单方式：移除该 turn 以及所有后续的 handle
            const turnIds = Array.from(_handles.keys());
            const idx = turnIds.indexOf(turnId);
            if (idx >= 0) {
                for (let i = idx; i < turnIds.length; i++) {
                    _handles.delete(turnIds[i]);
                }
            }
        }

        // 8. 恢复发送按钮状态
        App.ChatV2.chatEntry?.resetProcessingUI?.();

        ChatLog.info?.('RENDERER', `[editUserMessage] 进入编辑模式 turn=${turnId}`);
    }

    // ============ 编辑模式：附件预览渲染 ============

    /**
     * 在编辑重发模式下，将附件渲染到 #attachment-preview 预览区。
     * 不依赖 File 对象，直接使用 session-store 中的附件数据结构。
     * 图片附件显示缩略图，文档附件显示文件名和大小。
     * @param {Array} attachments - [{type, mime, fileName, size, base64?, content?}]
     */
    function renderEditingAttachmentPreview(attachments) {
        if (!attachments || attachments.length === 0) return;

        const preview = document.getElementById('attachment-preview');
        if (!preview) return;

        // 显示预览区
        preview.style.display = 'flex';

        // 隐藏旧版单文件 header
        const legacyHeader = preview.querySelector('.attachment-icon-only');
        if (legacyHeader) legacyHeader.style.display = 'none';

        // 渲染附件列表
        const list = document.getElementById('attachment-list');
        if (!list) return;

        list.innerHTML = '';
        attachments.forEach((att, index) => {
            const item = document.createElement('div');
            item.className = 'attachment-item';

            let icon = '📎';
            let sizeStr = '';

            if (att.type === 'image' && att.base64) {
                // 图片附件：显示缩略图
                icon = '🖼️';
                sizeStr = formatSize(att.size || 0);
                item.innerHTML = `
                    <div class="attachment-icon" style="width:32px;height:32px;display:flex;align-items:center;justify-content:center;">
                        <img src="data:${att.mime || 'image/png'};base64,${att.base64}" alt="" style="max-width:32px;max-height:32px;border-radius:4px;">
                    </div>
                    <div class="attachment-info">
                        <span class="attachment-name" title="${escapeHtml(att.fileName || '')}">${escapeHtml(att.fileName || 'image')}</span>
                        <div class="attachment-meta">
                            <span class="attachment-size">${sizeStr}</span>
                            <span class="attachment-type-badge type-image">图片</span>
                        </div>
                    </div>
                `;
            } else {
                // 文档附件
                icon = fileIconForAttachment(att.fileName || '');
                sizeStr = formatSize(att.size || 0);
                const ext = extOfName(att.fileName || '');
                item.innerHTML = `
                    <span class="attachment-icon">${icon}</span>
                    <div class="attachment-info">
                        <span class="attachment-name" title="${escapeHtml(att.fileName || '')}">${escapeHtml(att.fileName || 'file')}</span>
                        <div class="attachment-meta">
                            <span class="attachment-size">${sizeStr}</span>
                            <span class="attachment-type-badge type-doc">${escapeHtml(ext)}</span>
                        </div>
                    </div>
                `;
            }

            list.appendChild(item);
        });

        list.style.display = 'flex';

        // 单图片时还显示底部缩略图区域（与正常文件选择保持一致）
        const thumb = document.getElementById('attachment-thumbnail');
        if (attachments.length === 1 && attachments[0].type === 'image' && attachments[0].base64) {
            const img = document.getElementById('thumbnail-img');
            if (thumb && img) {
                img.src = 'data:' + (attachments[0].mime || 'image/png') + ';base64,' + attachments[0].base64;
                thumb.style.display = 'block';
            }
        } else if (thumb) {
            thumb.style.display = 'none';
        }
    }

    function fileIconForAttachment(name) {
        const n = (name || '').toLowerCase();
        if (/\.xlsx?$/.test(n)) return '📊';
        if (/\.pdf$/.test(n)) return '📕';
        if (/\.docx?$/.test(n)) return '📝';
        if (/\.txt$/.test(n)) return '📄';
        return '📎';
    }

    function extOfName(name) {
        const i = (name || '').lastIndexOf('.');
        return i >= 0 ? name.slice(i + 1).toUpperCase() : '文件';
    }

    function formatSize(bytes) {
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.formatSize(bytes) : '0 B';
    }

    // ============ PPT 生成配置卡（主题选择 + 版式勾选）============

    const PPT_SLIDE_LAYOUTS = [
        { type: 'title', label: '标题页', desc: '封面' },
        { type: 'table', label: '数据表格', desc: '明细数据' },
        { type: 'text', label: '文本要点', desc: '数据概览' },
        { type: 'chart', label: '图表', desc: '数据可视化' },
        { type: 'kpi', label: 'KPI 指标', desc: '关键数字' },
        { type: 'comparison', label: '对比', desc: '双栏对比' },
        { type: 'timeline', label: '时间线', desc: '里程碑' }
    ];

    function getPptThemeList() {
        const engine = App.ChatV2?.services?.vbaTemplateEngine;
        if (engine && typeof engine.getThemeList === 'function') {
            const list = engine.getThemeList();
            if (Array.isArray(list) && list.length > 0) return list;
        }
        // 兜底：内置四主题
        return [
            { id: 'business_blue', name: '商务蓝', primary: [0, 90, 156] },
            { id: 'tech_green', name: '科技绿', primary: [0, 128, 128] },
            { id: 'vibrant_orange', name: '活力橙', primary: [230, 100, 50] },
            { id: 'elegant_purple', name: '优雅紫', primary: [100, 60, 150] }
        ];
    }

    /**
     * 渲染 PPT 生成配置卡：主题色板 + 版式勾选 + 生成按钮。
     * 用户点击"生成 PPT / 使用默认生成"后，把所选配置交给等待中的 code-pipeline。
     */
    function renderPptConfigCard(handle, config) {
        if (!handle || !config) return;
        // 移除 loading dots（等待用户输入，非流式）
        if (handle.dotsEl && handle.dotsEl.parentNode) {
            handle.dotsEl.remove();
            handle.dotsEl = null;
        }
        // 幂等：若已存在配置卡，先移除
        const existing = handle.main.querySelector('.ppt-config-card');
        if (existing) existing.remove();

        const themeList = getPptThemeList();
        const currentTheme = config.theme || 'business_blue';
        const layout = config.layout || { title: true, table: true, text: true };
        // 图表类型标签（数据层"分析视图"与规划布局"图表页配置"共用）
        const chartTypeLabels = { column: '柱状图', bar: '条形图', line: '折线图', pie: '饼图', doughnut: '环形图', area: '面积图', radar: '雷达图', scatter: '散点图', combo: '组合图', bubble: '气泡图' };
        const chartTypeSelOpts = ['column', 'bar', 'line', 'pie', 'doughnut', 'area', 'radar', 'scatter', 'combo', 'bubble'];

        const card = el('div', 'ppt-config-card');
        card.dataset.theme = currentTheme;
        card.style.cssText = 'border:1px solid #e0e0e0;border-radius:8px;margin:8px 0;padding:12px;background:var(--card-bg,#fafafa);font-size:13px;';

        const head = el('div', 'ppt-config-head');
        head.style.cssText = 'font-weight:600;margin-bottom:10px;color:var(--text-primary,#333);';
        head.textContent = '🎨 PPT 生成配置';
        card.appendChild(head);

        // ---- 三段式层次：样色布局 → 规划布局 → 数据层，每层灰色卡片分组，标题带编号徽标 ----
        const mkAppSec = (no, title, desc) => {
            const wrap = el('div');
            wrap.style.cssText = 'background:#f7f8fa;border:1px solid #ececec;border-radius:8px;padding:2px 10px 10px;margin-bottom:10px;';
            const hRow = el('div');
            hRow.style.cssText = 'display:flex;align-items:center;gap:6px;margin:8px 0 4px;';
            const badge = el('span');
            badge.style.cssText = 'padding:1px 7px;border-radius:4px;background:#0f6cbd;color:#fff;font-size:11px;font-weight:600;';
            badge.textContent = no;
            const name = el('span');
            name.style.cssText = 'font-size:12px;font-weight:600;color:#333;';
            name.textContent = title;
            hRow.appendChild(badge);
            hRow.appendChild(name);
            if (desc) {
                const note = el('span');
                note.style.cssText = 'font-size:11px;color:#999;margin-left:2px;';
                note.textContent = '· ' + desc;
                hRow.appendChild(note);
            }
            wrap.appendChild(hRow);
            return wrap;
        };
        const secStyle = mkAppSec('01', '样色布局', '主题 · 比例 · 模板');
        const secPlan = mkAppSec('02', '规划布局', '页面结构 · 已识别需求');
        const secData = mkAppSec('03', '数据层', '图表形式 · 聚合 · 分析维度 · 洞察');
        card.appendChild(secStyle);
        card.appendChild(secPlan);
        card.appendChild(secData);

        // 推荐设计预览（数据感知设计引擎给出的页面结构，可逐页勾选）
        const recommendedSlides = Array.isArray(config.recommendedSlides) ? config.recommendedSlides : null;
        let recFlags = null;   // 页面级勾选状态（true=保留），供 collectConfig 读取
        if (recommendedSlides && recommendedSlides.length > 0) {
            recFlags = new Array(recommendedSlides.length).fill(true);
            const recLabel = el('div', 'ppt-config-label');
            recLabel.style.cssText = 'margin:2px 0 4px;color:var(--text-secondary,#666);display:flex;align-items:center;justify-content:space-between;';
            const recLabelText = el('span');
            recLabelText.textContent = '推荐设计';
            const recCount = el('span', 'ppt-rec-count');
            recCount.style.cssText = 'font-size:11px;color:var(--text-secondary,#999);font-weight:400;';
            recLabel.appendChild(recLabelText);
            recLabel.appendChild(recCount);
            secPlan.appendChild(recLabel);
            const recRow = el('div', 'ppt-config-recommend');
            recRow.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px;';
            const typeLabels = {
                title: '封面', agenda: '目录', section: '章节', kpi: '关键指标',
                chart: '图表', table: '数据表格', text: '要点', comparison: '对比',
                timeline: '时间线', closing: '结尾'
            };
            const applyChipStyle = (chip, on) => {
                chip.style.cssText = on
                    ? 'padding:2px 7px;border-radius:10px;background:var(--accent-soft,#e8f1fb);color:#0f6cbd;font-size:11px;cursor:pointer;border:1px solid transparent;'
                    : 'padding:2px 7px;border-radius:10px;background:#f0f0f0;color:#aaa;font-size:11px;cursor:pointer;border:1px dashed #d0d0d0;text-decoration:line-through;';
            };
            const updateRecCount = () => {
                const n = recFlags.filter(Boolean).length;
                recCount.textContent = '已选 ' + n + '/' + recommendedSlides.length + ' 页';
            };
            recommendedSlides.forEach((s, i) => {
                const t = String(s?.type || '').toLowerCase();
                const typeName = typeLabels[t] || t;
                const slideTitle = s?.title ? String(s.title) : '';
                const chip = el('span', 'ppt-rec-chip');
                chip.dataset.idx = i;
                applyChipStyle(chip, true);
                chip.textContent = slideTitle ? (typeName + '·' + slideTitle) : typeName;
                chip.title = typeName + '（点击取消/恢复该页）';
                chip.addEventListener('click', () => {
                    recFlags[i] = !recFlags[i];
                    applyChipStyle(chip, recFlags[i]);
                    updateRecCount();
                });
                recRow.appendChild(chip);
            });
            updateRecCount();
            secPlan.appendChild(recRow);

            // 图表·聚合视图规划位：由数据层"分析视图"实时驱动更新；
            // 只要添加了聚合视图，规划布局的图表章节即以此为准（替换推荐图表页），逐项可勾选保留
            const planAggLabel = el('div', 'ppt-config-label');
            planAggLabel.style.cssText = 'margin:2px 0 4px;color:var(--text-secondary,#666);display:none;';
            planAggLabel.textContent = '图表·聚合视图';
            secPlan.appendChild(planAggLabel);
            const planAggBox = el('div', 'ppt-config-aggplan');
            planAggBox.style.cssText = 'display:flex;flex-wrap:wrap;gap:4px;margin-bottom:10px;';
            secPlan.appendChild(planAggBox);
            card.__planAggExcl = new Set();
            card.__syncExcl = () => { card.dataset.aggExcluded = JSON.stringify(Array.from(card.__planAggExcl)); };
            card.__setPlanAggViews = (views) => {
                const excl = card.__planAggExcl || new Set();
                if (!Array.isArray(views) || views.length === 0) {
                    planAggLabel.style.display = 'none';
                    planAggBox.innerHTML = '';
                    return;
                }
                planAggLabel.style.display = '';
                planAggBox.innerHTML = '';
                views.forEach(v => {
                    if (!v) return;
                    const k = [String(v.xField || ''), String(v.measure || ''), String(v.agg || '')].join('|');
                    const on = !excl.has(k);
                    const chip = el('span');
                    chip.style.cssText = on
                        ? 'padding:2px 7px;border-radius:10px;background:#e8f1fb;color:#0f6cbd;font-size:11px;cursor:pointer;border:1px solid transparent;'
                        : 'padding:2px 7px;border-radius:10px;background:#f0f0f0;color:#aaa;font-size:11px;cursor:pointer;border:1px dashed #d0d0d0;text-decoration:line-through;';
                    chip.textContent = '图表·' + v.xField + '·' + v.measure;
                    chip.title = v.xField + ' × ' + v.measure + '（' + v.agg + '）点击保留/取消';
                    chip.addEventListener('click', () => {
                        if (on) { excl.add(k); } else { excl.delete(k); }
                        card.__syncExcl();
                        card.__setPlanAggViews(views);
                    });
                    planAggBox.appendChild(chip);
                });
                const hint = el('span');
                hint.style.cssText = 'font-size:11px;color:#999;';
                hint.textContent = '｜ 已添加聚合视图将替换推荐图表页，逐个勾选保留';
                planAggBox.appendChild(hint);
            };
        }

        // ---- 逐页配置：推荐设计中的图表页可逐页指定 图型 / X维度 / 数值 / 聚合 / 洞察 ----
        const ovSlides = (Array.isArray(config.recommendedSlides) ? config.recommendedSlides : [])
            .map((s, i) => ({ s, i }))
            .filter(o => o.s && String(o.s.type || '').toLowerCase() === 'chart');
        card.__slideOverrides = {};
        if (ovSlides.length > 0) {
            const ovLabel = el('div', 'ppt-config-label');
            ovLabel.style.cssText = 'margin:2px 0 4px;color:var(--text-secondary,#666);';
            ovLabel.textContent = '图表页配置（可选）';
            secPlan.appendChild(ovLabel);
            const ovBox = el('div', 'ppt-config-slideov');
            ovBox.style.cssText = 'margin-bottom:10px;padding:8px;border:1px solid #eee;border-radius:8px;background:#fcfcfc;';
            const ovHeaders = Array.isArray(config.headers) ? config.headers : [];
            const ovRows = Array.isArray(config.rows) ? config.rows : [];
            const ovNum = (v) => { const n = parseFloat(String(v).replace(/[¥￥,%\s]/g, '')); return isNaN(n) ? null : n; };
            const ovIsNum = (h) => ovRows.some(r => ovNum(r[h]) !== null);
            const ovNumCols = ovHeaders.filter(ovIsNum);
            const ovDimCols = ovHeaders.filter(h => !ovIsNum(h));
            const inHeaders = (v) => ovHeaders.some(h => String(h) === String(v));
            // 数据列校验：X 维度/数值列若不在真实数据列中，VB 端会静默回退到默认维度（常是时间轴），
            // 极易造成"配置卡显示地区、实际输出时间"的不透明现象。这里在提交前就标红预警。
            const markMissing = (sel, val) => {
                if (val && !inHeaders(val)) {
                    sel.style.borderColor = '#d9534f';
                    sel.title = '⚠ 「' + val + '」不在数据列中，提交后将按默认维度绘制（可能不是你预期的时间轴）';
                }
            };
            const mkOvSel = (opts, cur, w) => {
                const s = el('select');
                s.style.cssText = `padding:2px 5px;border:1px solid #ddd;border-radius:5px;background:#fff;color:#555;font-size:11px;${w ? `width:${w}px;` : ''}`;
                for (const o of opts) {
                    const opt = el('option');
                    opt.value = o;
                    opt.textContent = o;
                    if (o === cur) opt.selected = true;
                    s.appendChild(opt);
                }
                return s;
            };
            const ovCommit = () => { card.dataset.slideOverrides = JSON.stringify(card.__slideOverrides); };
            ovSlides.forEach(({ s, i }) => {
                const rowWrap = el('div');
                rowWrap.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:5px;margin:4px 0;';
                const ttl = el('span');
                ttl.style.cssText = 'font-size:11px;color:#333;font-weight:600;min-width:86px;max-width:132px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;';
                ttl.textContent = '图表·' + (s.title || '图表');
                ttl.title = String(s.title || '');
                rowWrap.appendChild(ttl);
                const ov = card.__slideOverrides[i] = {
                    __title: String(s.title || ''),
                    chartType: String(s.chartType || 'column'),
                    xField: String(s.xField || (ovDimCols[0] || ovHeaders[0] || '')),
                    series: [String((Array.isArray(s.series) && s.series[0]) || (ovNumCols[0] || ovHeaders[0] || ''))],
                    agg: String(['sum', 'avg', 'count'].includes(s.agg) ? s.agg : 'sum'),
                    insight: String(s.insight || '')
                };
                const tSel = mkOvSel(chartTypeSelOpts, ov.chartType, 82);
                Array.from(tSel.options).forEach(o => { o.textContent = chartTypeLabels[o.value] || o.value; });
                tSel.addEventListener('change', () => { ov.chartType = tSel.value; ovCommit(); });
                rowWrap.appendChild(tSel);
                const xSel = mkOvSel(ovDimCols.length ? ovDimCols : ovHeaders, ov.xField, 90);
                xSel.addEventListener('change', () => { ov.xField = xSel.value; ovCommit(); markMissing(xSel, ov.xField); });
                markMissing(xSel, ov.xField);
                rowWrap.appendChild(xSel);
                const mSel = mkOvSel(ovNumCols.length ? ovNumCols : ovHeaders, ov.series[0], 90);
                mSel.addEventListener('change', () => { ov.series = [mSel.value]; ovCommit(); markMissing(mSel, ov.series[0]); });
                markMissing(mSel, ov.series[0]);
                rowWrap.appendChild(mSel);
                const gSel = mkOvSel(['sum', 'avg', 'count'], ov.agg, 58);
                Array.from(gSel.options).forEach(o => { o.textContent = ({ sum: '求和', avg: '均值', count: '计数' })[o.value] || o.value; });
                gSel.addEventListener('change', () => { ov.agg = gSel.value; ovCommit(); });
                rowWrap.appendChild(gSel);
                const ipt = el('input');
                ipt.type = 'text';
                ipt.placeholder = '洞察（可选）';
                ipt.value = ov.insight;
                ipt.style.cssText = 'flex:1;min-width:110px;padding:2px 6px;border:1px solid #ddd;border-radius:5px;font-size:11px;color:#555;';
                ipt.addEventListener('input', () => { ov.insight = ipt.value; ovCommit(); });
                rowWrap.appendChild(ipt);
                ovBox.appendChild(rowWrap);
            });
            const ovHint = el('div');
            ovHint.style.cssText = 'font-size:10px;color:#999;margin-top:4px;';
            ovHint.textContent = '仅覆盖上方"推荐设计"中对应图表页的字段；不改动则沿用设计引擎的自动推荐。';
            ovBox.appendChild(ovHint);
            secPlan.appendChild(ovBox);
        }

        // 已识别需求（规则解析 / LLM 规划的用户要求）
        const recognized = Array.isArray(config.recognizedRequirements) ? config.recognizedRequirements : null;
        if (recognized && recognized.length > 0) {
            const recLabel = el('div', 'ppt-config-label');
            recLabel.style.cssText = 'margin:2px 0 4px;color:var(--text-secondary,#666);';
            recLabel.textContent = '已识别需求';
            secPlan.appendChild(recLabel);
            const recBox = el('div', 'ppt-config-recognized');
            recBox.style.cssText = 'margin-bottom:10px;padding:8px 10px;border-left:3px solid #2e9e5b;background:var(--card-bg,#fafafa);font-size:12px;color:var(--text-secondary,#555);line-height:1.6;';
            recBox.textContent = recognized.join('；');
            secPlan.appendChild(recBox);
        }

        // 图表形式：图片（稳定不可编辑）vs 带数据的原生图（可编辑）；auto 自动回退
        const chartModeRow = el('div', 'ppt-config-chartmode');
        chartModeRow.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-bottom:10px;';
        const chartModeLabel = el('div', 'ppt-config-label');
        chartModeLabel.style.cssText = 'margin:0 2px 0 0;color:var(--text-secondary,#666);';
        chartModeLabel.textContent = '图表形式';
        chartModeRow.appendChild(chartModeLabel);
        const currentMode = config.chartMode || 'auto';
        const mkChartModeChip = (mode, label, tip) => {
            const lab = el('button', 'ppt-chartmode-chip');
            lab.type = 'button';
            lab.dataset.chartMode = mode;
            const on = mode === currentMode;
            lab.style.cssText = on
                ? 'padding:4px 10px;border:1.5px solid #0f6cbd;border-radius:6px;background:var(--accent-soft,#e8f1fb);color:#0f6cbd;font-size:12px;cursor:pointer;'
                : 'padding:4px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;color:#555;font-size:12px;cursor:pointer;';
            lab.textContent = label;
            lab.title = tip;
            lab.addEventListener('click', () => {
                chartModeRow.querySelectorAll('.ppt-chartmode-chip').forEach(b => {
                    b.style.borderColor = '#ddd'; b.style.borderWidth = '1px'; b.style.background = '#fff'; b.style.color = '#555';
                });
                lab.style.border = '1.5px solid #0f6cbd';
                lab.style.background = 'var(--accent-soft,#e8f1fb)';
                lab.style.color = '#0f6cbd';
                card.dataset.chartMode = mode;
            });
            return lab;
        };
        chartModeRow.appendChild(mkChartModeChip('auto', '自动（推荐）', '原生图表优先，失败自动降级为图片'));
        chartModeRow.appendChild(mkChartModeChip('picture', '图片', '固定用图片，稳定不可编辑，兼容性最好'));
        chartModeRow.appendChild(mkChartModeChip('native', '带数据的图', '固定用原生图表，可点击编辑数据'));
        secData.appendChild(chartModeRow);

        // 聚合方式（用户驱动模型）：数据驱动自动判别 vs 用户显式指定 求和/均值/计数
        const aggRow = el('div', 'ppt-config-agg');
        aggRow.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-bottom:10px;';
        const aggLabel = el('div', 'ppt-config-label');
        aggLabel.style.cssText = 'margin:0 2px 0 0;color:var(--text-secondary,#666);';
        aggLabel.textContent = '聚合方式';
        aggRow.appendChild(aggLabel);
        const currentAgg = config.agg || 'auto';
        const mkAggChip = (agg, label, tip) => {
            const lab = el('button', 'ppt-agg-chip');
            lab.type = 'button';
            lab.dataset.agg = agg;
            const on = agg === currentAgg;
            lab.style.cssText = on
                ? 'padding:4px 10px;border:1.5px solid #0f6cbd;border-radius:6px;background:var(--accent-soft,#e8f1fb);color:#0f6cbd;font-size:12px;cursor:pointer;'
                : 'padding:4px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;color:#555;font-size:12px;cursor:pointer;';
            lab.textContent = label;
            lab.title = tip;
            lab.addEventListener('click', () => {
                aggRow.querySelectorAll('.ppt-agg-chip').forEach(b => {
                    b.style.borderColor = '#ddd'; b.style.borderWidth = '1px'; b.style.background = '#fff'; b.style.color = '#555';
                });
                lab.style.border = '1.5px solid #0f6cbd';
                lab.style.background = 'var(--accent-soft,#e8f1fb)';
                lab.style.color = '#0f6cbd';
                card.dataset.agg = agg;
            });
            return lab;
        };
        aggRow.appendChild(mkAggChip('auto', '自动（推荐）', '按数据维度自动生成聚合视图，选信息量最佳的聚合函数'));
        aggRow.appendChild(mkAggChip('sum', '求和', '各类别数值列求和'));
        aggRow.appendChild(mkAggChip('avg', '均值', '各类别数值列取平均'));
        aggRow.appendChild(mkAggChip('count', '计数', '各类别出现次数统计'));
        secData.appendChild(aggRow);

        // ---- 分析维度多视图选择器（用户驱动，可同时添加多个分类聚合 + 实时预览）----
        // 数据在手，让用户逐一添加"看哪一维 × 看哪个数 × 怎么聚"，多个视图同时进入 PPT 图表页
        const hasPv = Array.isArray(config.headers) && config.headers.length > 0
            && Array.isArray(config.rows) && config.rows.length > 0;
        if (hasPv) {
            const pvHeaders = config.headers;
            const pvRows = config.rows;
            const stripNum = (v) => parseFloat(String(v).replace(/[¥￥,\s]/g, ''));
            const isNumCol = (h) => pvRows.some(r => {
                const v = r[h];
                if (typeof v === 'number') return true;
                if (typeof v === 'string') return !isNaN(stripNum(v));
                return false;
            });
            const numCols = pvHeaders.filter(isNumCol);
            const dimCols = pvHeaders.filter(h => !isNumCol(h));
            const aggOptLabel = { sum: '求和', avg: '均值', count: '计数' };
            // 初始视图列表：沿用上次配置的 aggViews；否则默认"推荐首图维度×数值×聚合"
            const recCharts = (Array.isArray(config.recommendedSlides) ? config.recommendedSlides : [])
                .filter(s => String(s?.type || '').toLowerCase() === 'chart');
            const recChart = recCharts[0];
            const initViews = Array.isArray(config.aggViews) && config.aggViews.length
                ? config.aggViews.map(v => ({
                    xField: String(v.xField ?? ''),
                    measure: String(v.measure ?? (Array.isArray(v.series) ? v.series[0] : v.xField)),
                    agg: String(['sum', 'avg', 'count'].includes(v.agg) ? v.agg : 'sum'),
                    chartType: chartTypeSelOpts.includes(v.chartType) ? v.chartType : 'column',
                    categories: Array.isArray(v.categories) ? v.categories.map(String) : []
                }))
                : recCharts.map(s => ({
                    xField: String(s.xField ?? ''),
                    measure: String(Array.isArray(s.series) && s.series[0] ? s.series[0] : ''),
                    agg: String(['sum', 'avg', 'count'].includes(s.agg) ? s.agg : 'sum'),
                    chartType: chartTypeSelOpts.includes(s.chartType) ? s.chartType : 'column',
                    categories: [],
                    title: String(s.title ?? '')
                })).filter(v => v.xField && v.measure);
            const defDim = (dimCols.includes(config.aggDim) ? config.aggDim : '')
                || (Number.isInteger(dimCols.indexOf(recChart?.xField)) && dimCols.indexOf(recChart?.xField) >= 0 ? recChart.xField : (dimCols[0] || ''));
            const defChartType = chartTypeSelOpts.includes(recChart?.chartType) ? recChart.chartType : 'column';
            const defMeas = (numCols.includes(config.aggMeasure) ? config.aggMeasure : '')
                || (Array.isArray(recChart?.series) && numCols.includes(recChart.series[0]) ? recChart.series[0] : (numCols[0] || ''));
            const defAgg = ['sum', 'avg', 'count'].includes(config.aggFunc) ? config.aggFunc
                : (['sum', 'avg', 'count'].includes(recChart?.agg) ? recChart.agg : 'sum');

            const mkSel = (opts, cur, w) => {
                const s = el('select');
                s.style.cssText = `padding:3px 6px;border:1px solid #ddd;border-radius:6px;background:#fff;color:#555;font-size:12px;${w ? `width:${w}px;` : ''}`;
                for (const o of opts) {
                    const opt = el('option');
                    opt.value = o;
                    opt.textContent = o;
                    if (o === cur) opt.selected = true;
                    s.appendChild(opt);
                }
                return s;
            };
            const dimSel = mkSel(dimCols, defDim);
            const measSel = mkSel(numCols, defMeas);
            const aggSel = mkSel(['sum', 'avg', 'count'], defAgg, 76);
            Array.from(aggSel.options).forEach(o => { o.textContent = aggOptLabel[o.value] || o.value; });
            const chartTypeSel = mkSel(chartTypeSelOpts, defChartType, 84);
            Array.from(chartTypeSel.options).forEach(o => { o.textContent = chartTypeLabels[o.value] || o.value; });

            const dimRow = el('div');
            dimRow.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-bottom:6px;';
            const dimLabel = el('div');
            dimLabel.style.cssText = 'margin:0 2px 0 0;color:var(--text-secondary,#666);width:60px;';
            dimLabel.textContent = '分析视图';
            dimRow.appendChild(dimLabel);
            dimRow.appendChild(dimSel);
            dimRow.appendChild(measSel);
            dimRow.appendChild(aggSel);
            dimRow.appendChild(chartTypeSel);
            const addBtn = el('button');
            addBtn.type = 'button';
            addBtn.textContent = '＋ 添加此聚合';
            addBtn.style.cssText = 'padding:3px 10px;border:1px solid #0f6cbd;border-radius:6px;background:#fff;color:#0f6cbd;font-size:12px;cursor:pointer;';
            dimRow.appendChild(addBtn);
            secData.appendChild(dimRow);

            // 已添加视图列表 + 预览
            const viewsBox = el('div');
            viewsBox.style.cssText = 'margin:0 0 6px 66px;display:flex;flex-wrap:wrap;gap:6px;min-height:6px;';
            secData.appendChild(viewsBox);
            const pvBox = el('div');
            pvBox.style.cssText = 'margin:0 0 10px 66px;padding:8px;border:1px solid #eee;border-radius:8px;background:#fcfcfc;';
            secData.appendChild(pvBox);

            // 对比类别多选面板：列出该视图维度的全部取值，勾选即写回 view.categories（空=全部）
            const catPanel = el('div');
            catPanel.style.cssText = 'margin:0 0 10px 66px;padding:8px;border:1px solid #d6e4f5;border-radius:8px;background:#f5f9ff;display:none;';
            secData.appendChild(catPanel);
            const distinctValues = (xf) => {
                const seen = new Set(); const out = [];
                for (const r of pvRows) {
                    const k = String(r[xf] ?? '（空）');
                    if (!seen.has(k)) { seen.add(k); out.push(k); }
                }
                return out;
            };
            const openCatPanel = (v) => {
                const vals = distinctValues(v.xField);
                catPanel.innerHTML = '';
                const head = el('div');
                head.style.cssText = 'font-size:11px;color:#0f6cbd;margin-bottom:6px;';
                head.textContent = `对比类别（按「${v.xField}」）：勾选要对比的项，默认全部`;
                catPanel.appendChild(head);
                const grid = el('div');
                grid.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;max-height:160px;overflow:auto;';
                const sel = new Set((Array.isArray(v.categories) && v.categories.length) ? v.categories.map(String) : vals);
                vals.forEach(val => {
                    const cb = el('label');
                    cb.style.cssText = 'display:inline-flex;align-items:center;gap:3px;font-size:12px;padding:2px 6px;border:1px solid #c9d8ea;border-radius:6px;background:#fff;cursor:pointer;';
                    const input = el('input');
                    input.type = 'checkbox';
                    input.checked = sel.has(val);
                    input.addEventListener('change', () => {
                        if (input.checked) sel.add(val); else sel.delete(val);
                        v.categories = Array.from(sel);
                        renderViews();
                        previewOf(v);
                    });
                    cb.appendChild(input);
                    cb.appendChild(document.createTextNode(val));
                    grid.appendChild(cb);
                });
                catPanel.appendChild(grid);
                const tools = el('div');
                tools.style.cssText = 'margin-top:6px;display:flex;gap:6px;';
                const allBtn = el('button'); allBtn.type = 'button'; allBtn.textContent = '全选';
                allBtn.style.cssText = 'padding:2px 10px;border:1px solid #0f6cbd;border-radius:6px;background:#fff;color:#0f6cbd;font-size:12px;cursor:pointer;';
                allBtn.addEventListener('click', () => { v.categories = vals.slice(); renderViews(); previewOf(v); openCatPanel(v); });
                const clearBtn = el('button'); clearBtn.type = 'button'; clearBtn.textContent = '清除（全部）';
                clearBtn.style.cssText = 'padding:2px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;color:#555;font-size:12px;cursor:pointer;';
                clearBtn.addEventListener('click', () => { v.categories = []; renderViews(); previewOf(v); openCatPanel(v); });
                tools.appendChild(allBtn); tools.appendChild(clearBtn);
                catPanel.appendChild(tools);
                catPanel.style.display = 'block';
            };

            const views = initViews;
            // 前端聚合（仅预览用；最终 PPT 数值由后端按同规格聚合，口径一致）
            const aggregate = (x, m, fn, cats) => {
                const groups = new Map();
                for (const r of pvRows) {
                    const k = String(r[x] ?? '（空）');
                    const v = parseFloat(String(r[m]).replace(/[¥￥,\s]/g, ''));
                    if (!groups.has(k)) groups.set(k, { sum: 0, cnt: 0 });
                    const g = groups.get(k);
                    if (!isNaN(v)) { g.sum += v; g.cnt++; }
                }
                const all = Array.from(groups.entries()).map(([k, g]) => ({
                    k, v: fn === 'avg' ? (g.cnt ? g.sum / g.cnt : 0) : (fn === 'count' ? g.cnt : g.sum)
                }));
                // 用户显式指定对比类别（白名单）：仅保留勾选项、按勾选顺序，不做 Top8 截断与排序
                if (Array.isArray(cats) && cats.length) {
                    return cats.map(String).map(c => all.find(it => it.k === c)).filter(Boolean);
                }
                return all.sort((a, b) => b.v - a.v).slice(0, 8);
            };
            const evalView = (x, m, fn, cats) => ({ x, m, fn, items: aggregate(x, m, fn, cats) });
            const renderBars = (box, view, chartType) => {
                const items = view.items || [];
                const max = items.length ? items.reduce((mx, it) => (it.v > mx ? it.v : mx), -Infinity) : 1;
                box.innerHTML = '';
                const title = el('div');
                title.style.cssText = 'font-size:11px;color:var(--text-secondary,#666);margin-bottom:6px;';
                const tname = chartTypeLabels[chartType] || chartType || '柱状图';
                title.textContent = `预览·${tname}：按【${view.x}】对【${view.m}】${aggOptLabel[view.fn] || view.fn}`;
                box.appendChild(title);
                if (items.length === 0) {
                    const empty = el('div');
                    empty.style.cssText = 'font-size:11px;color:#aaa;';
                    empty.textContent = '该维度暂无可用数据';
                    box.appendChild(empty);
                    return;
                }
                const fmtVal = (v) => view.fn === 'avg' ? Number(v).toFixed(1) : Math.round(v).toLocaleString();
                const CAT = ['#0f6cbd', '#2e9e5b', '#e08214', '#9b59b6', '#16a085', '#d9534f', '#3498db', '#f39c12'];
                const pvType = String(chartType || 'column');
                const pvLabels = items.map(it => String(it.k));
                const pvValues = items.map(it => Number(it.v) || 0);
                // 优先用 ECharts 真实渲染（图型与最终 PPT 一致）；不可用时回退下方简易渲染
                if (typeof echarts !== 'undefined' && echarts && typeof echarts.init === 'function') {
                    try {
                        if (box.__echarts && typeof box.__echarts.dispose === 'function') { box.__echarts.dispose(); }
                        box.__echarts = null;
                        const PALETTE = ['#0f6cbd', '#2e9e5b', '#e08214', '#9b59b6', '#16a085', '#d9534f', '#3498db', '#f39c12', '#7f8c8d', '#c0392b'];
                        const axisSty = { axisLabel: { fontSize: 10, hideOverlap: true, interval: 'auto' } };
                        const gridSty = { top: 18, left: 8, right: 16, bottom: 22, containLabel: true };
                        const buildOpt = (t) => {
                            const base = { color: PALETTE, animationDuration: 400 };
                            if (t === 'pie' || t === 'doughnut') {
                                base.tooltip = { trigger: 'item', formatter: '{b}: {c} ({d}%)' };
                                base.legend = { bottom: 0, type: 'scroll', textStyle: { fontSize: 10 }, itemWidth: 10, itemHeight: 8 };
                                base.series = [{
                                    type: 'pie',
                                    radius: t === 'doughnut' ? ['45%', '70%'] : '62%',
                                    center: ['50%', '46%'],
                                    data: pvLabels.map((l, i) => ({ name: l, value: pvValues[i] })),
                                    label: { fontSize: 10, formatter: '{b}\n{d}%' }
                                }];
                                return base;
                            }
                            if (t === 'radar') {
                                const mx = Math.max(1, ...pvValues);
                                base.tooltip = {};
                                base.radar = {
                                    indicator: pvLabels.map(l => ({ name: l, max: Math.ceil(mx * 1.15) })),
                                    radius: '62%', center: ['50%', '52%'],
                                    axisName: { fontSize: 10, color: '#666' }
                                };
                                base.series = [{ type: 'radar', data: [{ value: pvValues, name: view.m }], areaStyle: { opacity: 0.25 }, symbolSize: 4 }];
                                return base;
                            }
                            if (t === 'scatter' || t === 'bubble') {
                                const mx = Math.max(1, ...pvValues);
                                const data = pvValues.map((v, i) => t === 'bubble' ? [i, v, Math.max(8, Math.round((v / mx) * 34) + 8)] : [i, v]);
                                const ser = { type: 'scatter', data };
                                if (t === 'bubble') { ser.symbolSize = (d) => d[2]; } else { ser.symbolSize = 9; }
                                base.tooltip = { trigger: 'item', formatter: (p) => (pvLabels[p.data[0]] || '') + ': ' + p.data[1] };
                                base.grid = gridSty;
                                base.xAxis = Object.assign({ type: 'category', data: pvLabels }, axisSty);
                                base.yAxis = { type: 'value', axisLabel: { fontSize: 10 } };
                                base.series = [ser];
                                return base;
                            }
                            if (t === 'combo') {
                                base.tooltip = { trigger: 'axis' };
                                base.grid = gridSty;
                                base.legend = { bottom: 0, textStyle: { fontSize: 10 }, itemWidth: 12, itemHeight: 8 };
                                base.xAxis = Object.assign({ type: 'category', data: pvLabels }, axisSty);
                                base.yAxis = { type: 'value', axisLabel: { fontSize: 10 } };
                                base.series = [
                                    { name: view.m, type: 'bar', data: pvValues, barMaxWidth: 28, itemStyle: { borderRadius: [3, 3, 0, 0] } },
                                    { name: view.m + '（折线）', type: 'line', data: pvValues, smooth: true, symbolSize: 5, lineStyle: { width: 2 } }
                                ];
                                return base;
                            }
                            if (t === 'bar') {
                                base.tooltip = { trigger: 'axis' };
                                base.grid = gridSty;
                                base.xAxis = { type: 'value', axisLabel: { fontSize: 10 } };
                                base.yAxis = Object.assign({ type: 'category', data: pvLabels }, axisSty);
                                base.series = [{ type: 'bar', data: pvValues, barMaxWidth: 18, itemStyle: { borderRadius: [0, 3, 3, 0] } }];
                                return base;
                            }
                            // column / line / area
                            const isLine = t === 'line';
                            const isArea = t === 'area';
                            base.tooltip = { trigger: 'axis' };
                            base.grid = gridSty;
                            base.xAxis = Object.assign({ type: 'category', data: pvLabels, boundaryGap: !(isLine || isArea) }, axisSty);
                            base.yAxis = { type: 'value', axisLabel: { fontSize: 10 } };
                            base.series = [{
                                type: (isLine || isArea) ? 'line' : 'bar',
                                data: pvValues,
                                smooth: isLine || isArea,
                                symbolSize: 5,
                                barMaxWidth: 32,
                                areaStyle: isArea ? { opacity: 0.22 } : undefined,
                                itemStyle: (isLine || isArea) ? {} : { borderRadius: [3, 3, 0, 0] }
                            }];
                            return base;
                        };
                        const host = el('div');
                        host.style.cssText = 'width:100%;height:240px;';
                        box.appendChild(host);
                        const chart = echarts.init(host);
                        chart.setOption(buildOpt(pvType));
                        box.__echarts = chart;
                        // 卡片可能在本帧稍后才完成布局，延迟一次 resize 防止 0 尺寸
                        setTimeout(() => { try { chart.resize(); } catch (_) { /* noop */ } }, 60);
                        if (window.ResizeObserver) {
                            new ResizeObserver(() => { try { chart.resize(); } catch (_) { /* noop */ } }).observe(host);
                        } else {
                            window.addEventListener('resize', () => { try { chart.resize(); } catch (_) { /* noop */ } });
                        }
                        return;
                    } catch (e) {
                        ChatLog.warn?.('RENDER', '[ppt-preview] ECharts 预览失败，回退简易渲染:', e?.message);
                        box.innerHTML = '';
                    }
                }
                // 折线 / 面积：SVG
                if (chartType === 'line' || chartType === 'area') {
                    const W = 250, H = 110, pad = 16, n = items.length;
                    const stepX = n > 1 ? (W - pad * 2) / (n - 1) : 0;
                    const yOf = (v) => H - pad - (max ? (v / max) * (H - pad * 2) : 0);
                    const pts = items.map((it, i) => [pad + i * stepX, yOf(it.v)]);
                    const svgNS = 'http://www.w3.org/2000/svg';
                    const svg = document.createElementNS(svgNS, 'svg');
                    svg.setAttribute('width', '100%');
                    svg.setAttribute('viewBox', `0 0 ${W} ${H}`);
                    svg.style.cssText = 'display:block;max-width:250px;';
                    if (chartType === 'area') {
                        const d = 'M' + pts.map(p => p[0] + ',' + p[1]).join(' L') + ` L${pad + (n - 1) * stepX},${H - pad} L${pad},${H - pad} Z`;
                        const ar = document.createElementNS(svgNS, 'path');
                        ar.setAttribute('d', d); ar.setAttribute('fill', 'rgba(15,108,189,0.15)');
                        svg.appendChild(ar);
                    }
                    const pl = document.createElementNS(svgNS, 'polyline');
                    pl.setAttribute('points', pts.map(p => p[0] + ',' + p[1]).join(' '));
                    pl.setAttribute('fill', 'none'); pl.setAttribute('stroke', '#0f6cbd'); pl.setAttribute('stroke-width', '2');
                    svg.appendChild(pl);
                    pts.forEach(p => {
                        const c = document.createElementNS(svgNS, 'circle');
                        c.setAttribute('cx', p[0]); c.setAttribute('cy', p[1]); c.setAttribute('r', '2.5');
                        c.setAttribute('fill', '#0f6cbd'); svg.appendChild(c);
                    });
                    box.appendChild(svg);
                    const cap = el('div');
                    cap.style.cssText = 'font-size:10px;color:#999;margin-top:2px;';
                    cap.textContent = items.map(it => it.k).join(' · ');
                    box.appendChild(cap);
                    return;
                }
                // 饼 / 环形：conic-gradient
                if (chartType === 'pie' || chartType === 'doughnut') {
                    const total = items.reduce((s, it) => s + Math.max(0, it.v), 0) || 1;
                    let acc = 0; const segs = [];
                    items.forEach((it, i) => {
                        const start = acc / total * 360;
                        acc += Math.max(0, it.v);
                        segs.push(`${start}deg ${acc / total * 360}deg ${CAT[i % CAT.length]}`);
                    });
                    const size = 96;
                    const pie = el('div');
                    pie.style.cssText = `position:relative;width:${size}px;height:${size}px;border-radius:50%;background:conic-gradient(${segs.join(',')});`;
                    if (chartType === 'doughnut') {
                        const inner = el('div');
                        inner.style.cssText = `position:absolute;top:50%;left:50%;width:${size * 0.55}px;height:${size * 0.55}px;background:#fff;border-radius:50%;transform:translate(-50%,-50%);`;
                        pie.appendChild(inner);
                    }
                    box.appendChild(pie);
                    const lg = el('div');
                    lg.style.cssText = 'margin-top:4px;font-size:10px;color:#666;line-height:1.5;';
                    lg.textContent = items.map((it, i) => `${it.k}: ${fmtVal(it.v)}`).join('   ');
                    box.appendChild(lg);
                    return;
                }
                // 柱状（竖向）
                if (chartType === 'column') {
                    const wrap = el('div');
                    wrap.style.cssText = 'display:flex;align-items:flex-end;gap:6px;height:118px;padding:4px 0;';
                    items.forEach((it, i) => {
                        const pct = max ? Math.max(3, Math.round((it.v / max) * 100)) : 3;
                        const col = el('div');
                        col.style.cssText = `width:22px;height:${pct}%;background:${CAT[i % CAT.length]};border-radius:3px 3px 0 0;opacity:.85;display:flex;align-items:flex-end;justify-content:center;`;
                        const vlab = el('div');
                        vlab.style.cssText = 'font-size:9px;color:#fff;margin-bottom:1px;';
                        vlab.textContent = fmtVal(it.v);
                        col.appendChild(vlab);
                        const colWrap = el('div');
                        colWrap.style.cssText = 'display:flex;flex-direction:column;align-items:center;gap:2px;';
                        colWrap.appendChild(col);
                        const klab = el('div');
                        klab.style.cssText = 'font-size:9px;color:#777;width:34px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:center;';
                        klab.textContent = it.k;
                        colWrap.appendChild(klab);
                        wrap.appendChild(colWrap);
                    });
                    box.appendChild(wrap);
                    return;
                }
                // 条形 / 雷达 / 散点 / 组合 / 气泡 等：横向条形（回退，标题已标图型）
                const barWrap = el('div');
                for (const it of items) {
                    const rowLine = el('div');
                    rowLine.style.cssText = 'display:flex;align-items:center;gap:6px;margin:2px 0;';
                    const lab = el('span');
                    lab.style.cssText = 'font-size:11px;color:#555;width:80px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;text-align:right;flex:none;';
                    lab.textContent = it.k;
                    const bar = el('div');
                    const pct = max ? Math.max(2, Math.round((it.v / max) * 100)) : 2;
                    bar.style.cssText = `height:12px;border-radius:3px;background:var(--accent,#0f6cbd);opacity:.82;width:${pct}%;min-width:4px;flex:none;`;
                    const val = el('span');
                    val.style.cssText = 'font-size:11px;color:#333;';
                    val.textContent = fmtVal(it.v);
                    rowLine.appendChild(lab);
                    const barTrack = el('div');
                    barTrack.style.cssText = 'flex:1;min-width:0;display:flex;align-items:center;';
                    barTrack.appendChild(bar);
                    rowLine.appendChild(barTrack);
                    rowLine.appendChild(val);
                    barWrap.appendChild(rowLine);
                }
                box.appendChild(barWrap);
            };
            const syncSel = () => {
                card.dataset.aggDim = dimSel.value || defDim;
                card.dataset.aggMeasure = measSel.value || defMeas;
                card.dataset.aggFunc = aggSel.value || defAgg;
            };
            const previewOf = (v) => renderBars(pvBox, evalView(v.xField, v.measure, v.agg, v.categories), v.chartType);
            const renderViews = () => {
                card.dataset.aggViews = JSON.stringify(views.map(v => ({ xField: v.xField, series: [v.measure], agg: v.agg, chartType: v.chartType || 'column', categories: v.categories || [] })));
                viewsBox.innerHTML = '';
                views.forEach((v, i) => {
                    const pill = el('span');
                    pill.style.cssText = 'display:inline-flex;align-items:center;gap:4px;padding:3px 8px;border:1px solid #c9d8ea;border-radius:12px;background:#e8f1fb;color:#0f6cbd;font-size:11px;cursor:pointer;user-select:none;';
                    const catTag = (Array.isArray(v.categories) && v.categories.length) ? `·${v.categories.length}类` : '';
                    const txt = el('span', null, `${v.xField}·${v.measure}·${aggOptLabel[v.agg] || v.agg}·${chartTypeLabels[v.chartType] || v.chartType}${catTag}`);
                    txt.style.cssText = 'font-weight:600;';
                    const catBtn = el('span');
                    catBtn.style.cssText = 'color:#0f6cbd;font-size:10px;cursor:pointer;padding:0 3px;border:1px solid #c9d8ea;border-radius:8px;background:#fff;';
                    catBtn.textContent = '类别';
                    catBtn.addEventListener('click', e => { e.stopPropagation(); openCatPanel(v); });
                    const del = el('span');
                    del.style.cssText = 'color:#d9534f;font-size:12px;cursor:pointer;padding:0 2px;';
                    del.textContent = '×';
                    del.addEventListener('click', e => { e.stopPropagation(); views.splice(i, 1); renderViews(); });
                    pill.appendChild(txt);
                    pill.appendChild(catBtn);
                    pill.appendChild(del);
                    pill.addEventListener('click', () => previewOf(v));
                    viewsBox.appendChild(pill);
                });
                if (views.length === 0) {
                    const hint = el('span');
                    hint.style.cssText = 'font-size:11px;color:#aaa;';
                    hint.textContent = '尚未添加聚合视图——需多维度对比时逐一添加，每项生成一页图表';
                    viewsBox.appendChild(hint);
                }
                // 聚合驱动规划布局：任一视图增删都实时刷新"规划布局"的图表·聚合视图规划位
                if (typeof card.__setPlanAggViews === 'function') card.__setPlanAggViews(views);
                if (typeof card.__syncExcl === 'function') card.__syncExcl();
            };

            addBtn.addEventListener('click', () => {
                const x = dimSel.value || defDim;
                const m = measSel.value || defMeas;
                const fn = aggSel.value || defAgg;
                const ct = chartTypeSel.value || defChartType;
                const dup = views.find(v => v.xField === x && v.measure === m && v.agg === fn && (v.chartType || 'column') === ct);
                if (dup) { previewOf(dup); return; }
                views.push({ xField: x, measure: m, agg: fn, chartType: ct, categories: [] });
                renderViews();
                previewOf(views[views.length - 1]);
            });
            [dimSel, measSel, aggSel].forEach(s => s.addEventListener('change', syncSel));
            syncSel();
            renderViews();
            if (views.length) previewOf(views[views.length - 1]);
        }

        // 版式比例：16:9（宽屏，默认）vs 4:3（标准屏）vs 竖版（720x960）
        const ratioRow = el('div', 'ppt-config-ratio');
        ratioRow.style.cssText = 'display:flex;flex-wrap:wrap;align-items:center;gap:6px;margin-bottom:10px;';
        const ratioLabel = el('div', 'ppt-config-label');
        ratioLabel.style.cssText = 'margin:0 2px 0 0;color:var(--text-secondary,#666);';
        ratioLabel.textContent = '版式比例';
        ratioRow.appendChild(ratioLabel);
        const currentRatio = config.ratio || '16x9';
        const mkRatioChip = (ratio, orientation, label, tip) => {
            const lab = el('button', 'ppt-ratio-chip');
            lab.type = 'button';
            lab.dataset.ratio = ratio || '';
            lab.dataset.orientation = orientation || '';
            const on = (ratio ? currentRatio === ratio : (config.orientation || '') === orientation);
            lab.style.cssText = on
                ? 'padding:4px 10px;border:1.5px solid #0f6cbd;border-radius:6px;background:var(--accent-soft,#e8f1fb);color:#0f6cbd;font-size:12px;cursor:pointer;'
                : 'padding:4px 10px;border:1px solid #ddd;border-radius:6px;background:#fff;color:#555;font-size:12px;cursor:pointer;';
            lab.textContent = label;
            lab.title = tip;
            lab.addEventListener('click', () => {
                ratioRow.querySelectorAll('.ppt-ratio-chip').forEach(b => {
                    b.style.borderColor = '#ddd'; b.style.borderWidth = '1px'; b.style.background = '#fff'; b.style.color = '#555';
                });
                lab.style.border = '1.5px solid #0f6cbd';
                lab.style.background = 'var(--accent-soft,#e8f1fb)';
                lab.style.color = '#0f6cbd';
                card.dataset.ratio = ratio || '';
                card.dataset.orientation = orientation || '';
            });
            return lab;
        };
        ratioRow.appendChild(mkRatioChip('16x9', '', '宽屏 16:9（推荐）', '宽屏，默认，适配现代投影'));
        ratioRow.appendChild(mkRatioChip('4x3', '', '标准屏 4:3', '标准屏，传统投影'));
        ratioRow.appendChild(mkRatioChip('', 'portrait', '竖版', '纵向，适合长表/手机上阅读'));
        secStyle.appendChild(ratioRow);

        // 主题选择（简化：紧凑色点 + 小加号导入）
        const themeRow = el('div', 'ppt-config-themes');
        themeRow.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin-bottom:6px;';
        for (const t of themeList) {
            const on = t.id === currentTheme;
            const swatch = el('button', 'ppt-theme-swatch');
            swatch.type = 'button';
            swatch.dataset.theme = t.id;
            swatch.title = t.name;
            swatch.style.cssText = 'width:22px;height:22px;border-radius:50%;border:' + (on ? '2px solid #0f6cbd;box-shadow:0 0 0 2px rgba(15,108,189,0.25);' : '2px solid transparent;') + 'background:rgb(' + t.primary.join(',') + ');cursor:pointer;padding:0;flex:none;';
            swatch.addEventListener('click', () => {
                themeRow.querySelectorAll('.ppt-theme-swatch').forEach(b => { b.style.border = '2px solid transparent'; b.style.boxShadow = 'none'; });
                swatch.style.border = '2px solid #0f6cbd';
                swatch.style.boxShadow = '0 0 0 2px rgba(15,108,189,0.25)';
                card.dataset.theme = t.id;
                if (typeof updateLayoutApplyVisibility === 'function') updateLayoutApplyVisibility();
            });
            themeRow.appendChild(swatch);
        }
        // 导入模板小按钮（悬停显示说明）
        const importBtn = el('button', 'ppt-import-btn');
        importBtn.type = 'button';
        importBtn.textContent = '＋';
        importBtn.title = '导入模板：选择 .pptx，提取其配色与字体作为新主题';
        importBtn.style.cssText = 'width:22px;height:22px;padding:0;background:#fff;color:#aaa;border:1px dashed #ccc;border-radius:50%;cursor:pointer;font-size:14px;line-height:1;flex:none;';
        themeRow.appendChild(importBtn);
        secStyle.appendChild(themeRow);
        importBtn.addEventListener('click', async () => {
            statusEl.style.display = 'block';
            statusEl.textContent = '⏳ 请选择公司模板文件…';
            importBtn.disabled = true;
            try {
                // 裸 WebView2 宿主对象优先，【不要】套 BridgeAccessor 的包装层：
                // 包装层调用会弹 UI 的 COM 方法（BrowsePptFile）时，是在【宿主已经弹出文件对话框之后】
                // 才抛 'e is not a function'，于是同一次导入里对话框被弹出两次（先桥、后 MCP 回落）。
                // 裸桥要么一次成功，要么方法不存在直接抛错（压根不弹窗），两种情形都只弹一次。
                const rawBridge = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) || null;
                const bridge = rawBridge
                    || (window.BridgeAccessor && window.BridgeAccessor.getBridge())
                    || App.excelBridge || window.excelBridge;
                if (!bridge) {
                    statusEl.textContent = '❌ 未获取到宿主桥对象';
                    return;
                }
                // 本地 MCP 回落：宿主 DLL 未暴露 BrowsePptFile / ImportPptTheme 时，改调
                // PptMcpTools 的 ppt_browse_template / ppt_import_theme —— 二者走同一个
                // PptGenerationService，但经 CallLocalTool 调用，不依赖新增的 COM 桥方法，
                // 因此旧版宿主 DLL 下也能完成"选文件 + 导入主题"。
                // MCP 返回形态不统一（裸路径 / {success,filePath} / ToolsCallResult{content:[{text}]} /
                // {result:{...}}），只剥一层会出现"选到了文件却取不到路径"，这里递归剥壳。
                const unwrapObj = (o) => {
                    if (!o) return null;
                    if (typeof o === 'string') {
                        const t = o.trim();
                        if (!t) return null;
                        if (/^[a-zA-Z]:[\\/]|^\\\\/.test(t)) return { success: true, filePath: t };
                        try { return unwrapObj(JSON.parse(t)); } catch (_) { return { success: false, error: t }; }
                    }
                    if (Array.isArray(o.content) && o.content[0]) {
                        const c = o.content[0];
                        const txt = typeof c === 'string' ? c : (c && c.text);
                        if (typeof txt === 'string' && txt.trim()) {
                            try {
                                const inner = JSON.parse(txt);
                                if (inner && typeof inner === 'object') return inner;
                            } catch (_) { return { success: false, error: txt }; }
                        }
                    }
                    // result 常是【嵌套的 JSON 字符串】（实测：{"success":true,"result":"{\"success\":true,\"filePath\":\"...\"}"}），
                    // 只对 object 递归会漏掉字符串这层，导致"选到了文件却取不到路径"。
                    if (o.result) return unwrapObj(o.result);
                    return o;
                };
                // 注意：这里【不】用 callLocalToolWithFallback。它对失败会逐个换桥来源重试，
                // 而 ppt_browse_template 会弹系统文件对话框——重试一次就多弹一个框（实测连弹三次）。
                // 会弹 UI 的工具只调一次：调用失败直接上抛，由调用方决定是否报错。
                const callMcp = async (toolName, argsObj) => {
                    // 裸 WebView2 宿主对象优先：桥的包装层调用 CallLocalTool 会在【宿主已经弹出
                    // 文件对话框之后】才抛 'e is not a function'，此时换来源重试 = 让用户重选一遍
                    //（实测连弹三次）。因此裸桥优先，且只调一次，失败直接上抛，不再换来源。
                    const rawBridge = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) || null;
                    const wrapped = (window.BridgeAccessor && window.BridgeAccessor.getBridge())
                        || (typeof App !== 'undefined' && App.excelBridge) || window.excelBridge || null;
                    const cands = [];
                    if (rawBridge) cands.push({ b: rawBridge, label: 'hostObjects(裸桥)' });
                    if (wrapped && wrapped !== rawBridge) cands.push({ b: wrapped, label: 'wrapped' });
                    for (let i = 0; i < cands.length; i++) {
                        const c = cands[i];
                        if (!c || !c.b || typeof c.b.CallLocalTool !== 'function') continue;
                        const raw = await c.b.CallLocalTool(toolName, JSON.stringify(argsObj || {}));
                        const out = unwrapObj(raw);
                        try { console.log('[PPT导入] MCP', toolName, '来源:', c.label || i, '原始:', raw, '解析:', out); } catch (_) {}
                        return out;
                    }
                    return null;
                };
                let path = '';
                let browseRes = null;
                if (typeof bridge.BrowsePptFile === 'function') {
                    try { browseRes = JSON.parse((await bridge.BrowsePptFile()) || '{}'); } catch (pe) { browseRes = null; }
                }
                let mcpBrowse = null;
                if (browseRes && browseRes.success) {
                    path = browseRes.filePath || '';
                } else if (!browseRes || browseRes.error === 'bridge-call-failed') {
                    // 宿主未提供文件对话框 → 回落本地 MCP 工具（同样是原生对话框）
                    mcpBrowse = await callMcp('ppt_browse_template', {}).catch((e) => ({ __error: (e && e.message) || String(e) }));
                    if (mcpBrowse && mcpBrowse.success) path = mcpBrowse.filePath || mcpBrowse.path || '';
                }
                if (!path) {
                    // success=false 且无 error = 用户在原生对话框点了取消，静默关闭
                    if (browseRes && browseRes.success === false && !browseRes.error) {
                        statusEl.style.display = 'none';
                        return;
                    }
                    statusEl.textContent = '❌ 无法选择模板文件（'
                        + ((browseRes && browseRes.detail) || 'BrowsePptFile 调用失败')
                        + '；MCP 返回：' + (mcpBrowse ? JSON.stringify(mcpBrowse).slice(0, 180) : 'CallLocalTool 不可用')
                        + '）。请重新生成 Excel 宿主并重启 Excel 后重试。';
                    return;
                }
                statusEl.textContent = '⏳ 正在导入主题…';
                let res = null;
                try { res = JSON.parse((await bridge.ImportPptTheme(path)) || '{}'); } catch (pe) { res = null; }
                if (!res || !res.success) {
                    // 宿主未提供 ImportPptTheme → 回落本地 MCP 工具
                    const m = await callMcp('ppt_import_theme', { pptxPath: path }).catch(() => null);
                    if (m && m.success) res = m;
                }
                if (!res || !res.success) {
                    statusEl.textContent = '❌ 导入失败：' + (res && res.error ? res.error : '未知错误');
                    return;
                }
                const engine = App.ChatV2?.services?.vbaTemplateEngine;
                if (engine && typeof engine.loadPptTemplates === 'function') {
                    await engine.loadPptTemplates();
                }
                renderPptConfigCard(handle, { ...config, theme: res.themeId });
                const newCard = handle.main.querySelector('.ppt-config-card');
                const newStatus = newCard && newCard.querySelector('.ppt-config-status');
                if (newStatus) {
                    newStatus.style.display = 'block';
                    const layoutInfo = res && res.layout ? ('（已提取版式：' + res.layout + '）') : '';
                    newStatus.textContent = '✅ 已导入主题「' + (res.themeName || res.themeId || '') + '」并自动选中' + layoutInfo;
                }
            } catch (e) {
                statusEl.textContent = '❌ 导入异常：' + (e && e.message ? e.message : String(e));
            } finally {
                importBtn.disabled = false;
            }
        });

        // 模板版式套用（仅当所选主题带版式画像时显示）
        const layoutApplyRow = el('div', 'ppt-config-layout-apply');
        layoutApplyRow.style.cssText = 'display:none;flex-wrap:wrap;gap:6px;margin-bottom:12px;padding:8px 10px;border:1px dashed #c9d8ea;border-radius:6px;background:var(--card-bg,#fafafa);';
        const layoutApplyLabel = el('div', 'ppt-config-label');
        layoutApplyLabel.style.cssText = 'width:100%;margin:0 0 4px;color:var(--text-secondary,#666);';
        layoutApplyLabel.textContent = '模板版式';
        layoutApplyRow.appendChild(layoutApplyLabel);
        const mkApplyToggle = (key, label, def) => {
            const lab = el('label', 'ppt-layout-apply-chip');
            lab.style.cssText = 'display:inline-flex;align-items:center;gap:4px;padding:4px 8px;border:1px solid #ddd;border-radius:6px;background:#fff;cursor:pointer;font-size:12px;';
            const checked = config[key] !== undefined ? !!config[key] : def;
            lab.innerHTML = '<input type="checkbox" data-ppt-apply="' + key + '"' + (checked ? ' checked' : '') + ' style="accent-color:#0f6cbd;">' + escapeHtml(label);
            return lab;
        };
        layoutApplyRow.appendChild(mkApplyToggle('applyLayout', '套用模板版式（位置+形状）', true));
        layoutApplyRow.appendChild(mkApplyToggle('layoutAwarePlanning', 'LLM 规划感知版式', true));
        secStyle.appendChild(layoutApplyRow);
        const updateLayoutApplyVisibility = () => {
            const engine = App.ChatV2?.services?.vbaTemplateEngine;
            const hasLayout = engine && typeof engine.hasThemeLayout === 'function'
                ? engine.hasThemeLayout(card.dataset.theme || currentTheme)
                : false;
            layoutApplyRow.style.display = hasLayout ? 'flex' : 'none';
        };
        updateLayoutApplyVisibility();

        // 洞察预览（数据感知设计引擎给出的关键发现，结构化分区渲染）
        const insights = Array.isArray(config.insights) ? config.insights : null;
        if (insights && insights.length > 0) {
            const insLabel = el('div', 'ppt-config-label');
            insLabel.style.cssText = 'margin:2px 0 4px;color:var(--text-secondary,#666);';
            insLabel.textContent = '洞察预览';
            secData.appendChild(insLabel);
            const insBox = el('div', 'ppt-config-insights');
            insBox.style.cssText = 'margin-bottom:10px;border:1px solid #e0e0e0;border-radius:6px;background:var(--card-bg,#fafafa);overflow:hidden;';
            const isStructured = typeof insights[0] === 'object' && insights[0] !== null && typeof insights[0].kind === 'string';
            if (isStructured) {
                // KPI 指标卡（并排，合计大数字 + 平均/区间小字）
                const kpis = insights.filter(i => i.kind === 'kpi');
                if (kpis.length > 0) {
                    const kpiRow = el('div');
                    kpiRow.style.cssText = 'display:flex;flex-wrap:wrap;gap:6px;padding:8px;border-bottom:1px solid #eee;';
                    for (const k of kpis) {
                        const c = el('div');
                        c.style.cssText = 'flex:1;min-width:110px;padding:6px 8px;border:1px solid #e0e0e0;border-radius:6px;background:#fff;';
                        c.innerHTML =
                            '<div style="font-size:11px;color:#888;">' + escapeHtml(String(k.name ?? '')) + '</div>' +
                            '<div style="font-size:15px;font-weight:600;color:#0f6cbd;margin:1px 0;">' + escapeHtml(String(k.total ?? '')) + '</div>' +
                            '<div style="font-size:10px;color:#aaa;">均 ' + escapeHtml(String(k.avg ?? '')) + ' · ' + escapeHtml(String(k.min ?? '')) + '~' + escapeHtml(String(k.max ?? '')) + '</div>';
                        kpiRow.appendChild(c);
                    }
                    insBox.appendChild(kpiRow);
                }
                // 趋势 / 分布 / 备注（行内徽标 + 迷你进度条）
                const items = insights.filter(i => i.kind !== 'kpi');
                if (items.length > 0) {
                    const listEl = el('div');
                    listEl.style.cssText = 'padding:6px 8px;';
                    for (const it of items) {
                        const row = el('div');
                        row.style.cssText = 'display:flex;align-items:center;gap:6px;padding:4px 2px;font-size:12px;color:var(--text-secondary,#555);line-height:1.5;';
                        if (it.kind === 'trend') {
                            const up = it.direction === 'up';
                            const down = it.direction === 'down';
                            const badge = el('span');
                            badge.style.cssText = 'flex:none;font-size:10px;padding:1px 6px;border-radius:8px;color:#fff;background:' + (up ? '#2e9e5b' : down ? '#d9534f' : '#999') + ';';
                            badge.textContent = up ? '↑ 上升' : down ? '↓ 下降' : '→ 平稳';
                            row.appendChild(badge);
                            row.appendChild(el('span', null, '按「' + String(it.axis ?? '') + '」，「' + String(it.name ?? '') + '」从 ' + String(it.first ?? '') + ' 变化到 ' + String(it.last ?? '') + '（' + String(it.pct ?? '') + '%）'));
                        } else if (it.kind === 'composition') {
                            const pct = Math.min(100, parseFloat(it.pct) || 0);
                            const badge = el('span');
                            badge.style.cssText = 'flex:none;font-size:10px;padding:1px 6px;border-radius:8px;background:#e8f1fb;color:#0f6cbd;font-weight:600;';
                            badge.textContent = String(it.pct ?? '') + '%';
                            row.appendChild(badge);
                            const barWrap = el('div');
                            barWrap.style.cssText = 'flex:none;width:60px;height:6px;border-radius:3px;background:#eee;overflow:hidden;';
                            const bar = el('div');
                            bar.style.cssText = 'height:100%;background:#0f6cbd;border-radius:3px;width:' + pct + '%;';
                            barWrap.appendChild(bar);
                            row.appendChild(barWrap);
                            const name = el('span', null, String(it.name ?? '') + '·' + String(it.top ?? ''));
                            name.style.cssText = 'flex:none;font-weight:600;color:#333;';
                            row.appendChild(name);
                            row.appendChild(el('span', null, '出现最多（' + String(it.count ?? '') + ' 次）'));
                        } else {
                            row.appendChild(el('span', null, escapeHtml(String(it.text ?? ''))));
                        }
                        listEl.appendChild(row);
                    }
                    insBox.appendChild(listEl);
                }
            } else {
                // 旧格式兼容（纯文本数组）
                const box = el('div');
                box.style.cssText = 'padding:8px 10px;font-size:12px;color:var(--text-secondary,#555);line-height:1.6;';
                box.textContent = insights.join('；');
                insBox.appendChild(box);
            }
            secData.appendChild(insBox);
        }

        // 按钮
        const btnRow = el('div', 'ppt-config-actions');
        btnRow.style.cssText = 'display:flex;gap:8px;';
        const genBtn = el('button', 'ppt-gen-btn');
        genBtn.type = 'button';
        genBtn.textContent = '生成 PPT';
        genBtn.style.cssText = 'padding:7px 16px;background:#0f6cbd;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;';
        const defaultBtn = el('button', 'ppt-default-btn');
        defaultBtn.type = 'button';
        defaultBtn.textContent = '使用默认生成';
        defaultBtn.style.cssText = 'padding:7px 12px;background:#f0f0f0;color:#333;border:1px solid #ddd;border-radius:6px;cursor:pointer;font-size:13px;';
        btnRow.appendChild(genBtn);
        btnRow.appendChild(defaultBtn);
        card.appendChild(btnRow);

        // 状态提示
        const statusEl = el('div', 'ppt-config-status');
        statusEl.style.cssText = 'margin-top:8px;color:var(--text-secondary,#888);font-size:12px;display:none;';
        card.appendChild(statusEl);

        const collectConfig = () => {
            const theme = card.dataset.theme || currentTheme;
            // 版式类型过滤器：UI 已并入"推荐设计"逐页勾选，这里沿用配置默认（由推荐设计反推，全保留），
            // 避免把图表/KPI 等类型页误过滤掉；仅在仍有独立版式勾选时取其勾选结果。
            let selected = {};
            card.querySelectorAll('input[data-layout]').forEach(cb => { selected[cb.dataset.layout] = cb.checked; });
            if (Object.keys(selected).length === 0) selected = config.layout || { title: true, table: true, text: true };
            const apply = {};
            card.querySelectorAll('input[data-ppt-apply]').forEach(cb => { apply[cb.dataset.pptApply] = cb.checked; });
            // 页面级勾选：用户取消的推荐页索引（默认全选 → 空数组）
            let excludedSlides = [];
            if (recFlags) {
                excludedSlides = recFlags.map((v, i) => v ? -1 : i).filter(i => i >= 0);
            }
            const _ov = (card.__slideOverrides && Object.keys(card.__slideOverrides).length) ? card.__slideOverrides : undefined;
            return { theme, layout: selected, ...apply, excludedSlides, chartMode: card.dataset.chartMode || 'auto',
                slideOverrides: _ov,
                agg: card.dataset.agg || 'auto',
                aggDim: card.dataset.aggDim || '',
                aggMeasure: card.dataset.aggMeasure || '',
                aggFunc: card.dataset.aggFunc || '',
                aggViews: (() => { try { const v = JSON.parse(card.dataset.aggViews || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; } })(),
                aggExcluded: (() => { try { const v = JSON.parse(card.dataset.aggExcluded || '[]'); return Array.isArray(v) ? v : []; } catch (e) { return []; } })(),
                ratio: card.dataset.ratio || '16x9', orientation: card.dataset.orientation || '' };
        };
        const disableButtons = () => {
            genBtn.disabled = true;
            defaultBtn.disabled = true;
            genBtn.style.opacity = '0.6';
            defaultBtn.style.opacity = '0.6';
            statusEl.style.display = 'block';
            statusEl.textContent = '⏳ 正在生成，请稍候…';
        };
        genBtn.addEventListener('click', () => {
            // 防重入：同一个 turnId 只允许 resolve 一次，避免多 turn/多窗口并行触发多次生成
            if (App.ChatV2._pptResolvedTurns && App.ChatV2._pptResolvedTurns.has(handle._turnId)) return;
            (App.ChatV2._pptResolvedTurns || (App.ChatV2._pptResolvedTurns = new Set())).add(handle._turnId);
            disableButtons();
            App.ChatV2.resolvePptConfig?.(handle._turnId, collectConfig());
        });
        defaultBtn.addEventListener('click', () => {
            if (App.ChatV2._pptResolvedTurns && App.ChatV2._pptResolvedTurns.has(handle._turnId)) return;
            (App.ChatV2._pptResolvedTurns || (App.ChatV2._pptResolvedTurns = new Set())).add(handle._turnId);
            disableButtons();
            // 使用默认生成：版式用推荐默认，但页面勾选仍尊重用户取消操作
            let excludedSlides = [];
            if (recFlags) {
                excludedSlides = recFlags.map((v, i) => v ? -1 : i).filter(i => i >= 0);
            }
            App.ChatV2.resolvePptConfig?.(handle._turnId, { theme: currentTheme, layout: { ...layout }, applyLayout: true, layoutAwarePlanning: true, excludedSlides, chartMode: card.dataset.chartMode || 'auto',
                ratio: card.dataset.ratio || '16x9', orientation: card.dataset.orientation || '' });
        });

        // 插入到 AI 气泡：状态指示器之后、正文区之前（配置在上，结果在下）
        if (handle.statusIndicatorEl && handle.statusIndicatorEl.nextSibling) {
            handle.main.insertBefore(card, handle.statusIndicatorEl.nextSibling);
        } else {
            handle.main.appendChild(card);
        }
        autoScroll();
    }

    // ============ AI 消息气泡（带 loading + 思考区 + 正文区 + 工具区）============

    function createAIBubble(turn) {
        const container = getMessagesContainer();
        if (!container) return null;

        const wrapper = el('div', 'message-wrapper ai');
        wrapper.dataset.turnId = turn.id;
        wrapper.dataset.kind = turn.kind || 'text';

        const avatar = el('div', 'message-avatar', '🤖');
        const main = el('div', 'message-main');

        const header = el('div', 'message-header');
        const modelEl = el('span', 'message-model');
        modelEl.style.color = 'var(--text-secondary, #0f21c5ff)';
        modelEl.style.fontSize = '1em';
        modelEl.style.marginLeft = '2px';
        // 模型名不再显示在消息头部
        modelEl.style.display = 'none';
        header.appendChild(modelEl);
        header.appendChild(el('span', 'message-time', formatTime(turn.startedAt)));
        main.appendChild(header);

        // 思考区（默认折叠）
        const thinkingEl = el('div', 'thinking-process');
        thinkingEl.style.display = 'none';

        const thinkingHeader = el('div', 'thinking-header');
        const thinkingToggle = el('div', 'thinking-toggle', '▸ 思考过程');
        thinkingToggle.style.cursor = 'pointer';
        thinkingToggle.addEventListener('click', () => {
            const body = thinkingEl.querySelector('.thinking-body');
            if (body) {
                const isOpen = body.style.display !== 'none';
                body.style.display = isOpen ? 'none' : 'block';
                thinkingToggle.textContent = (isOpen ? '▸' : '▾') + ' 思考过程';
            }
        });
        thinkingHeader.appendChild(thinkingToggle);
        thinkingEl.appendChild(thinkingHeader);

        const thinkingBody = el('div', 'thinking-body');
        thinkingBody.style.display = 'none';

        const thinkingContent = el('div', 'thinking-content');
        thinkingBody.appendChild(thinkingContent);
        thinkingEl.appendChild(thinkingBody);
        main.appendChild(thinkingEl);

        // 状态指示区（独立于思考区，轻量显示）
        const statusIndicatorEl = el('div', 'ai-status-indicator');
        statusIndicatorEl.style.display = 'none';
        main.appendChild(statusIndicatorEl);

        // 工具调用区（在正文区之前，确保显示顺序与生成顺序一致：先执行过程，后执行结果）
        const toolsEl = el('div', 'execution-flow-tools');
        toolsEl.style.display = 'none';
        main.appendChild(toolsEl);

        // 正文区（Skill 卡片等最终结果渲染在此）
        const contentEl = el('div', 'message-content');
        main.appendChild(contentEl);

        // Loading dots
        const dotsEl = el('div', 'thinking-dots');
        dotsEl.innerHTML = '<span></span><span></span><span></span>';
        main.appendChild(dotsEl);

        wrapper.appendChild(avatar);
        wrapper.appendChild(main);
        container.appendChild(wrapper);

        return {
            wrapper,
            avatar,
            main,
            header,
            modelEl,
            thinkingEl,
            thinkingHeader,
            thinkingToggle,
            thinkingBody,
            statusIndicatorEl,
            thinkingContent,
            toolsEl,
            contentEl,
            dotsEl,
            rounds: new Map(),
            errorEl: null,
            _planInserted: false
        };
    }

    // ============ 模型名显示 ============

    function setModelName(handle, modelName) {
        // 模型名不再显示在消息头部
        if (handle?.modelEl) handle.modelEl.style.display = 'none';
    }

    // ============ 思考区增量 ============

    function appendReasoning(handle, delta) {
        if (!delta || !handle.thinkingContent) return;
        // 思考区可见：有推理内容时显示，但默认折叠
        if (handle.thinkingEl.style.display === 'none') {
            handle.thinkingEl.style.display = 'block';
            handle.thinkingEl.classList.add('has-reasoning');
        }

        // 流式跟随：标记为正在跟随最新行尾
        const body = handle.thinkingBody;
        if (body) {
            body.dataset.followEnd = 'true';
            body.classList.add('streaming');
        }

        handle.thinkingContent.textContent += delta;

        // 流式跟随滚动：使用 requestAnimationFrame 驱动
        if (body && body.dataset.followEnd === 'true') {
            cancelAnimationFrame(body._followRaf);
            body._followRaf = requestAnimationFrame(() => {
                body.scrollTop = body.scrollHeight;
            });
        }

        // 保持折叠状态，用户点击展开
        autoScroll();
    }

    // ============ 正文增量 ============

    function appendContent(handle, delta) {
        if (!delta) return;
        // 第一个 token 时移除 loading dots（已开始有内容）
        if (handle.dotsEl && handle.dotsEl.parentNode) {
            handle.dotsEl.remove();
            handle.dotsEl = null;
        }
        scheduleContentRender(handle._turnId);
    }

    // ============ 工具调用渲染 ============

    function getOrCreateRound(handle, round, roundLabel) {
        if (handle.rounds.has(round)) {
            return handle.rounds.get(round);
        }
        if (handle.toolsEl.style.display === 'none') {
            handle.toolsEl.style.display = 'block';
        }
        // 【规划前置】_planHtml 已全量废弃（thoughtList 迭代替换模型接管所有步骤展示），
        // 不再插入旧的 exec-plan-card 绿面板。toolsEl 容器保留给流式过程（如有其他工具卡需要），
        // 但步骤规划面板统一走 contentEl 里的 thoughtList。
        const roundDiv = el('div', 'exec-round');
        roundDiv.dataset.round = round;
        const roundBadge = roundLabel
            ? el('div', 'exec-tool-badge round-badge', roundLabel)
            : null;
        if (roundBadge) roundDiv.appendChild(roundBadge);
        const cardsContainer = el('div', 'exec-round-cards');
        roundDiv.appendChild(cardsContainer);
        handle.toolsEl.appendChild(roundDiv);
        const entry = { roundDiv, cardsContainer, cards: [], roundBadge };
        handle.rounds.set(round, entry);

        // 如果创建的是第2轮及以后，给第1轮也补上"步骤1"标签（如果还没有）
        if (round > 1 && handle.rounds.has(1)) {
            const firstRound = handle.rounds.get(1);
            if (!firstRound.roundBadge) {
                const badge = el('div', 'exec-tool-badge round-badge', ' 步骤 1');
                firstRound.roundDiv.insertBefore(badge, firstRound.cardsContainer);
                firstRound.roundBadge = badge;
            }
        }

        return entry;
    }

    function renderToolCard(handle, toolCall) {
        // 隐藏的工具调用不渲染（如 python-pipeline 的内部执行记录）
        if (toolCall.hidden) return;

        const round = toolCall.round || 1;
        // 轮次标签：第1轮默认不显示标签（除非有后续步骤，会在创建第2轮时补上）
        let roundLabel = null;
        if (round > 1) {
            roundLabel = ` 步骤 ${round}`;
        }
        const roundEntry = getOrCreateRound(handle, round, roundLabel);

        // 获取卡片渲染意图（里程碑三：Card Union）
        const registry = App.ChatV2?.toolRegistry;
        let cardIntent = null;
        if (registry && typeof registry.renderCardIntent === 'function') {
            try {
                cardIntent = registry.renderCardIntent(toolCall.name, 'call', toolCall.args);
            } catch (e) { /* 静默 */ }
        }
        const cardType = cardIntent?.card || 'generic';

        const card = el('div', 'exec-tool-card');
        card.dataset.toolCallId = toolCall.id;
        card.dataset.toolName = toolCall.name || '';
        card.dataset.status = toolCall.status || 'pending';
        card.dataset.cardType = cardType;
        if (cardType !== 'generic') {
            card.classList.add('exec-tool-card--' + cardType);
        }

        const header = el('div', 'exec-tool-header');
        const displayName = toolCall.displayName || toolCall.name || 'unknown';
        const statusEl = el('span', 'exec-tool-status status-pending', '⏳ 执行中...');

        // ── 根据卡片类型渲染不同头部内容 ──
        if (cardType === 'chart') {
            // 左侧：图表类型色标 + 标题
            const leftGroup = el('div', 'exec-tool-chart-left');
            const typeTag = el('span', 'exec-tool-chart-type-tag', cardIntent.chartType || '图表');
            const titleEl = el('span', 'exec-tool-chart-title', escapeHtml(cardIntent.title || '创建图表'));
            leftGroup.appendChild(typeTag);
            leftGroup.appendChild(titleEl);
            header.appendChild(leftGroup);
            header.appendChild(statusEl);
            card.appendChild(header);

            // 数据范围信息不再单独占一行，改为存到卡片数据属性里
            // 详情展开时一并展示
            card._dataRange = cardIntent.dataRange || '';
        } else if (cardType === 'terminal') {
            const nameEl = el('span', 'exec-tool-name', '💻 ' + escapeHtml(displayName));
            header.appendChild(nameEl);
            header.appendChild(statusEl);
            card.appendChild(header);

            // 代码预览（截断到80字符）
            if (cardIntent && cardIntent.description) {
                const codePreview = el('div', 'exec-tool-code-preview');
                codePreview.textContent = cardIntent.description;
                card.appendChild(codePreview);
            }
        } else {
            // 通用卡片（默认）
            const nameEl = el('span', 'exec-tool-name', '🔧 ' + escapeHtml(displayName));
            header.appendChild(nameEl);
            header.appendChild(statusEl);
            card.appendChild(header);
        }

        const resultEl = el('div', 'exec-tool-result');
        card.appendChild(resultEl);

        // 详情区域（折叠）— 完成后填充参数+结果详情
        const detailEl = el('div', 'exec-tool-detail');
        detailEl.style.display = 'none';
        card.appendChild(detailEl);

        // 存储 args 供完成后填入详情
        card._toolArgs = toolCall.args;

        // 【交互规则】步骤卡片默认折叠：详情只在用户点"详情"时展开。
        // （2026-10-06 确认：步骤确认完成后即折叠，不做任何自动展开。）

        roundEntry.cardsContainer.appendChild(card);
        roundEntry.cards.push({ card, statusEl, resultEl, detailEl });
        autoScroll();
    }

    function updateToolCard(handle, toolCallId, patch, turnId) {
        let card = null;
        for (const [, round] of handle.rounds) {
            const found = round.cards.find(c => c.card.dataset.toolCallId === toolCallId);
            if (found) { card = found; break; }
        }
        if (!card) return;

        const status = patch.status || 'pending';
        card.card.dataset.status = status;
        const cardType = card.card.dataset.cardType || 'generic';

        if (status === 'completed') {
            // 获取卡片完成渲染意图（里程碑三：Card Union）
            const registry = App.ChatV2?.toolRegistry;
            let resultIntent = null;
            if (registry && typeof registry.renderCardIntent === 'function') {
                try {
                    resultIntent = registry.renderCardIntent(card.card.dataset.toolName, 'result', card.card._toolArgs, patch);
                } catch (e) { /* 静默 */ }
            }

            // ── 根据卡片类型设置状态文本 ──
            if (cardType === 'chart') {
                card.statusEl.textContent = '✅ 图表已创建';
                card.statusEl.className = 'exec-tool-status status-completed';
            } else if (cardType === 'terminal') {
                card.statusEl.textContent = '✅ 执行完成';
                card.statusEl.className = 'exec-tool-status status-completed';
            } else {
                card.statusEl.textContent = '✅ 完成';
                card.statusEl.className = 'exec-tool-status status-completed';
            }
            if (patch.durationMs != null) {
                const dur = el('span', 'exec-tool-duration', ` · ${patch.durationMs}ms`);
                card.card.querySelector('.exec-tool-header').appendChild(dur);
            }

            // 构建详情内容（参数 + 结果）
            const args = card.card._toolArgs;
            const hasArgs = args && Object.keys(args).length > 0;
            const hasResult = patch.result != null || patch.resultSummary;
            const showDetailToggle = hasArgs || hasResult;

            let summary = '';
            if (resultIntent && resultIntent.content) {
                // 使用 Card Union 提供的展示内容
                summary = resultIntent.content;
            } else if (patch.resultSummary) {
                summary = patch.resultSummary;
            } else if (patch.result != null) {
                summary = autoSummarizeToolResult(card.card.dataset.toolName, patch.result);
            }

            // 【修复 代码溢出】工具卡片"结果摘要"显示前，先清洗 LLM 误抄进 summary 里的裸 JSON / CSS 字典尾巴
            if (summary) summary = sanitizeLeakedJsonSnippets(summary);

            // 结果摘要
            if (summary) {
                // 默认折叠 summary 和详情，只显示一行 header。
                // 用户点击"详情"后才展开。避免卡片占用过多垂直空间。
                card.resultEl.innerHTML = '';
            }

            // 【步骤规划面板已全量废弃】旧绿面板（exec-plan-card）由 thoughtList 迭代替换模型接管，
            // 不再从 patch.result 里 extractStepsPlan 渲面板。工具卡只展示自身执行结果即可。

            // 反馈明细已移入"返回结果"（以 JSON 键并入返回结果，返回结果保持 JSON 结构），
            // 常显区不再单独渲染 ✓/✕ 清单，避免与返回结果重复。

            // 详情折叠区（合并参数+结果）
            if (showDetailToggle) {
                const parts = [];
                // 渲染提示：① 模型 args._view ② 本轮用户原话 ③ 技能 outputFormat（统一走 resolveToolViewHint）。
                // 命中则按对应格式渲染，并默认展开详情让结果直接可见。
                const _hintView = resolveToolViewHint({ name: card.card.dataset.toolName, args: args }, userMessageForTurn(turnId));
                const _viewArgs = _hintView ? Object.assign({}, args, { _view: _hintView, __userView: !!detectViewFromMessage(userMessageForTurn(turnId)) }) : args;
                // 图表卡片的数据范围也放到详情区
                if (cardType === 'chart' && card.card._dataRange) {
                    parts.push(`<div class="exec-tool-result-summary"><div class="exec-detail-label">数据范围</div>${escapeHtml(card.card._dataRange)}</div>`);
                }
                if (hasArgs) {
                    parts.push(`<div class="exec-tool-args"><div class="exec-detail-label">参数</div>${formatToolArgsHtml(args)}</div>`);
                }
                if (patch.result != null) {
                    // 把反馈明细（清空内容/清除格式等）以 JSON 键并入返回结果，返回结果保持 JSON 结构
                    let resultForJson = patch.result;
                    const fb = (resultIntent && Array.isArray(resultIntent.feedback)) ? resultIntent.feedback : [];
                    if (fb.length) {
                        const fbObj = {};
                        for (const it of fb) {
                            const key = it && it.label ? String(it.label) : '';
                            if (!key) continue;
                            fbObj[key] = (it && it.value !== undefined) ? it.value : !!(it && it.done);
                        }
                        let baseVal = resultForJson;
                        try { baseVal = unwrapToolResult(resultForJson); } catch (_) { baseVal = resultForJson; }
                        if (baseVal && typeof baseVal === 'object' && !Array.isArray(baseVal)) {
                            resultForJson = { ...baseVal, ...fbObj };
                        } else if (resultForJson && typeof resultForJson === 'object' && !Array.isArray(resultForJson)) {
                            // 解包后只剩数组/标量（如 {success, mergedRows, data:[[…]]} 的 data 矩阵）：
                            // 绝不能把外层对象换成解包值 —— 统计字段（来源表/行数/表头）全在外层，
                            // 丢掉后渲染层无法识别合并结果，只能把整份数据矩阵摊成裸 JSON（2026-10-07 用户截图）。
                            resultForJson = { ...resultForJson, ...fbObj };
                        } else {
                            resultForJson = { '返回结果': (baseVal == null ? '' : baseVal), ...fbObj };
                        }
                    }
                    const resultHtml = formatToolResultHtml(resultForJson, _viewArgs);
                    parts.push(`<div class="exec-tool-result-detail"><div class="exec-detail-label">返回结果</div>${resultHtml}</div>`);
                }
                // 如有 summary，也一并放入详情区
                if (summary) {
                    parts.unshift(`<div class="exec-tool-result-summary"><div class="exec-detail-label">摘要</div>${escapeHtml(summary)}</div>`);
                }
                card.detailEl.innerHTML = parts.join('');

                // 添加切换按钮 —— 放在 header 里 inline 显示，不单独占一行
                const toggle = el('div', 'exec-tool-toggle', '▸ 详情');
                toggle.style.cursor = 'pointer';
                toggle.addEventListener('click', () => {
                    const isOpen = card.detailEl.style.display !== 'none';
                    card.detailEl.style.display = isOpen ? 'none' : 'block';
                    toggle.textContent = (isOpen ? '▸' : '▾') + ' 详情';
                });
                // 【完成即折叠】命中 _view/展示诉求也不再自动展开（2026-10-06 用户确认的交互规则）：
                // 步骤卡片在确认完成（✓/✗）后一律收起为一行摘要，保持消息流紧凑。
                // 表格/JSON/结果仍可查看：点"详情"展开，或看最终"任务完成 · 结果视图"与 LLM 正文总结。
                // （旧版"命中 _view 默认展开"导致多张卡全部保持展开撑爆消息流，实测截图多次。）
                // 放到 header 末尾（statusEl 之后），inline 不占独立一行
                const headerEl = card.card.querySelector('.exec-tool-header');
                if (headerEl) {
                    headerEl.appendChild(toggle);
                } else {
                    card.resultEl.appendChild(toggle);
                }
            }

            // 导出/保存成功：常显路径 + 打开文件夹 / 复制路径
            try {
                const pathInfo = extractExportFilePath(
                    card.card.dataset.toolName,
                    patch.result,
                    card.card._toolArgs
                );
                if (pathInfo && pathInfo.path) {
                    if (summary && !card.resultEl.querySelector('.exec-tool-result-summary-line')) {
                        const sumLine = el('div', 'exec-tool-result-summary-line', summary);
                        card.resultEl.appendChild(sumLine);
                    }
                    attachExportPathActions(card.resultEl, pathInfo.path);
                }
            } catch (_) { /* ignore */ }
        } else if (status === 'failed') {
            card.statusEl.textContent = '❌ 失败';
            card.statusEl.className = 'exec-tool-status status-failed';
            const errType = patch.errorType;
            const errReason = patch.errorReason;
            const errSuggestion = patch.errorSuggestion;
            const errText = String(patch.error || '未知错误');
            card.resultEl.innerHTML = formatToolErrorHtml(errText, {
                toolName: card.card.dataset.toolName,
                errorType: errType,
                errorReason: errReason,
                errorSuggestion: errSuggestion
            });
            attachToolRecoveryActions(card, { mode: 'failed', turnId: turnId });
        } else if (status === 'needs-input') {
            // 必填参数无法兜底，向用户反问补齐（提示而非失败）
            card.statusEl.textContent = '❓ 需补充参数';
            card.statusEl.className = 'exec-tool-status status-needs-input';
            const needText = String(patch.errorSuggestion || patch.error || '缺少必要参数，请补充后重试');
            const missing = Array.isArray(patch.needsInput) ? patch.needsInput : [];
            const missingList = missing.length
                ? '<div style="margin-top:6px;color:#8c8c8c;">缺少：' + missing.map(m => escapeHtml(m.label || m.parameter || '')).join('、') + '</div>'
                : '';
            card.resultEl.innerHTML =
                '<div style="color:#8c6420;background:#fff8e6;border:1px solid #f0d9a8;border-radius:6px;padding:8px 10px;">' +
                    '⚠️ ' + escapeHtml(needText) +
                '</div>' + missingList;
            attachNeedsInputForm(card, missing, turnId);
        }
        autoScroll();
    }

    function extractExportFilePath(toolName, rawResult, args) {
        const tc = App.ChatV2?.renderers?.toolCard;
        return tc && typeof tc.extractExportFilePath === 'function'
            ? tc.extractExportFilePath(toolName, rawResult, args)
            : null;
    }

    /** 复制文本（剪贴板不可用时降级 toast） */
    async function copyTextToClipboard(text) {
        const s = String(text || '');
        if (!s) return false;
        try {
            if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
                await navigator.clipboard.writeText(s);
                return true;
            }
        } catch (_) { /* fallthrough */ }
        try {
            const ta = document.createElement('textarea');
            ta.value = s;
            ta.style.cssText = 'position:fixed;left:-9999px;top:-9999px';
            document.body.appendChild(ta);
            ta.select();
            const ok = document.execCommand('copy');
            document.body.removeChild(ta);
            return !!ok;
        } catch (_) {
            return false;
        }
    }

    /**
     * 在资源管理器中打开路径（优先桥 OpenPathInExplorer；否则复制路径并提示）
     */
    async function openPathInExplorer(filePath) {
        const path = String(filePath || '').trim();
        if (!path) {
            App.showToast?.('路径为空', 'warning');
            return;
        }
        try {
            let bridge = null;
            if (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function') {
                bridge = BridgeAccessor.getBridge();
            }
            if (!bridge && App.excelBridge) bridge = App.excelBridge;
            if (!bridge && window.chrome?.webview?.hostObjects?.excelBridge) {
                bridge = window.chrome.webview.hostObjects.excelBridge;
            }
            if (bridge && typeof bridge.OpenPathInExplorer === 'function') {
                let raw;
                if (window.BridgeAccessor && typeof BridgeAccessor.callAsync === 'function') {
                    raw = await BridgeAccessor.callAsync('OpenPathInExplorer', path);
                } else {
                    raw = await bridge.OpenPathInExplorer(path);
                }
                let obj = raw;
                if (typeof raw === 'string') {
                    try { obj = JSON.parse(raw); } catch (_) { obj = { success: true }; }
                }
                if (obj && obj.success === false) {
                    throw new Error(obj.error || '打开失败');
                }
                return;
            }
        } catch (e) {
            ChatLog.warn?.('RENDERER', 'OpenPathInExplorer 失败:', e && e.message);
        }
        const copied = await copyTextToClipboard(path);
        const tr = (k, v) => { try { return App.i18n?.t?.(k, v) || k; } catch (_) { return k; } };
        App.showToast?.(
            copied
                ? tr('tool.copiedPath')
                : tr('tool.openFolderFail', { path: path }),
            'info'
        );
    }

    /** 导出成功：路径展示 + 操作按钮 */
    function attachExportPathActions(container, filePath) {
        if (!container || !filePath) return;
        if (container.querySelector('.exec-tool-path-actions')) return;
        const bar = el('div', 'exec-tool-path-actions');
        const pathEl = el('div', 'exec-tool-path-text', filePath);
        pathEl.title = filePath;
        const btns = el('div', 'exec-tool-path-btns');
        const tr = (k, v) => { try { return App.i18n?.t?.(k, v) || k; } catch (_) { return k; } };
        const copyBtn = el('button', 'exec-tool-action-btn', tr('tool.copyPath'));
        copyBtn.type = 'button';
        copyBtn.addEventListener('click', async () => {
            const ok = await copyTextToClipboard(filePath);
            App.showToast?.(ok ? tr('toast.pathCopied') : tr('toast.copyFail'), ok ? 'success' : 'error');
        });
        const openBtn = el('button', 'exec-tool-action-btn exec-tool-action-primary', tr('tool.openFolder'));
        openBtn.type = 'button';
        openBtn.addEventListener('click', () => { openPathInExplorer(filePath); });
        btns.appendChild(copyBtn);
        btns.appendChild(openBtn);
        bar.appendChild(pathEl);
        bar.appendChild(btns);
        container.appendChild(bar);
    }

    /** 把重试请求填入输入框（无 turn / 用户想改措辞时用） */
    function fillRetryPrompt(cardEl, argsOverride) {
        const toolName = cardEl?.dataset?.toolName || '';
        const args = argsOverride || cardEl?._toolArgs || {};
        const input = document.getElementById('message-input');
        if (!input) return;
        let argText = '';
        try { argText = JSON.stringify(args); } catch (_) { argText = ''; }
        input.value = '请重新执行工具 ' + toolName + (argText ? ('，参数：' + argText) : '');
        input.focus();
        try {
            input.style.height = 'auto';
            input.style.height = Math.min(input.scrollHeight, 200) + 'px';
        } catch (_) { /* ignore */ }
        const ti = (k, v) => { try { return App.i18n?.t?.(k, v) || k; } catch (_) { return k; } };
        App.showToast?.(ti('toast.retryFilled'), 'info');
    }

    /**
     * 一键重试失败/缺参工具（走 invokeTracked，生成新卡片；无 turn 则填入输入框）
     */
    async function retryToolFromCard(cardEl, argsOverride) {
        const ti = (k, v) => { try { return App.i18n?.t?.(k, v) || k; } catch (_) { return k; } };
        const toolName = cardEl?.dataset?.toolName || '';
        if (!toolName) {
            App.showToast?.(ti('toast.retryNoTool'), 'warning');
            return;
        }
        const args = argsOverride || cardEl._toolArgs || {};
        const tr = App.ChatV2?.toolRegistry;
        const ss = App.ChatV2?.sessionStore;
        if (!tr || typeof tr.invokeTracked !== 'function') {
            fillRetryPrompt(cardEl, args);
            return;
        }
        let turn = ss && typeof ss.getCurrentTurn === 'function' ? ss.getCurrentTurn() : null;
        if (!turn && ss && typeof ss.getTurns === 'function') {
            const turns = ss.getTurns() || [];
            turn = turns.length ? turns[turns.length - 1] : null;
        }
        if (!turn) {
            fillRetryPrompt(cardEl, args);
            return;
        }
        App.showToast?.(ti('toast.retrying'), 'info');
        try {
            const result = await tr.invokeTracked(turn, toolName, args, {
                round: 1,
                displayName: toolName,
                ctx: { turn: turn }
            });
            if (result && result.ok === false) {
                App.showToast?.(result.error || ti('toast.retryFail'), 'error');
            } else {
                App.showToast?.(ti('toast.retryDone'), 'success');
            }
        } catch (e) {
            App.showToast?.(ti('toast.retryFail') + ': ' + (e?.message || e), 'error');
        }
    }

    /** 失败卡：重试 / 改参数 */
    function attachToolRecoveryActions(cardHandle, opts) {
        if (!cardHandle || !cardHandle.resultEl) return;
        if (cardHandle.resultEl.querySelector('.exec-tool-recovery')) return;
        const ti = (k, v) => { try { return App.i18n?.t?.(k, v) || k; } catch (_) { return k; } };
        const bar = el('div', 'exec-tool-recovery');
        const retryBtn = el('button', 'exec-tool-action-btn exec-tool-action-primary', ti('tool.retry'));
        retryBtn.type = 'button';
        retryBtn.addEventListener('click', () => retryToolFromCard(cardHandle.card));
        const editBtn = el('button', 'exec-tool-action-btn', ti('tool.editRetry'));
        editBtn.type = 'button';
        editBtn.addEventListener('click', () => fillRetryPrompt(cardHandle.card));
        bar.appendChild(retryBtn);
        bar.appendChild(editBtn);
        cardHandle.resultEl.appendChild(bar);
    }

    /** needs-input：简易补参表单 + 提交重试 */
    function attachNeedsInputForm(cardHandle, missing, turnId) {
        if (!cardHandle || !cardHandle.resultEl) return;
        if (cardHandle.resultEl.querySelector('.exec-tool-needs-form')) return;
        const list = Array.isArray(missing) ? missing : [];
        const form = el('div', 'exec-tool-needs-form');
        if (!list.length) {
            attachToolRecoveryActions(cardHandle, { mode: 'needs-input', turnId: turnId });
            return;
        }
        const fields = [];
        list.forEach(function (m) {
            const key = m.parameter || m.name || m.key || '';
            if (!key) return;
            const row = el('div', 'exec-tool-needs-row');
            const label = el('label', 'exec-tool-needs-label', m.label || key);
            const input = document.createElement('input');
            input.className = 'exec-tool-needs-input';
            input.type = 'text';
            input.placeholder = m.placeholder || ('请输入 ' + (m.label || key));
            input.dataset.param = key;
            if (m.value != null) input.value = String(m.value);
            row.appendChild(label);
            row.appendChild(input);
            form.appendChild(row);
            fields.push(input);
        });
        const submit = el('button', 'exec-tool-action-btn exec-tool-action-primary', '补齐并重试');
        submit.type = 'button';
        submit.addEventListener('click', function () {
            const merged = Object.assign({}, cardHandle.card._toolArgs || {});
            fields.forEach(function (inp) {
                const k = inp.dataset.param;
                const v = (inp.value || '').trim();
                if (k && v) merged[k] = v;
            });
            cardHandle.card._toolArgs = merged;
            retryToolFromCard(cardHandle.card, merged);
        });
        form.appendChild(submit);
        cardHandle.resultEl.appendChild(form);
    }

    /**
     * 清理工具参数用于显示：截断过长的字符串值（如 HTML 内容、大段代码），
     * 避免参数展开时撑爆界面。超过 MAX_ARG_VALUE_LEN 的字符串截断为 "...(N chars)"
     */
    function sanitizeArgsForDisplay(obj, depth) {
        return App.ChatV2.renderers.toolCard.sanitizeArgsForDisplay(obj, depth);
    }

    /**
     * 深度解包工具返回结果，处理各种嵌套包裹格式：
     * - { result: {...}, taskCompleted: true }  (executeLocalExcelTool 包装)
     * - { success: true, result: {...} }       (CallLocalTool 包装)
     * - { ok: true, result: {...} }            (toolRegistry.invoke 包装)
     * - "JSON string"                          (字符串形式的 JSON)
     * - 多层嵌套：{result: {result: {...}, taskCompleted: true}}
     * 返回解包后的实际数据（可能是对象、数组、字符串）
     */
    function unwrapToolResult(raw) {
        return App.ChatV2.renderers.toolCard.unwrapToolResult(raw);
    }

    /**
     * 检测工具结果中是否包含 skill_card（在 outputs.content 中）
     * 如果是，返回 skill 信息对象；否则返回 null
     */
    function extractSkillCardFromResult(rawResult) {
        return App.ChatV2.renderers.toolCard.extractSkillCardFromResult(rawResult);
    }

    /**
     * 根据工具名和返回结果自动生成友好的摘要文本
     * 返回 null 表示无法生成摘要，回退到原始 JSON 展示
     */
    function autoSummarizeToolResult(toolName, rawResult) {
        return App.ChatV2.renderers.toolCard.autoSummarizeToolResult(toolName, rawResult);
    }

    function truncateForDisplay(result) {
        return App.ChatV2.renderers.toolCard.truncateForDisplay(result);
    }

    /**
     * 尝试把工具结果解析并格式化为可读 JSON
     * 支持：外层 JSON + 内层 JSON 字符串（如 { result: "{...}" }）
     */
    /**
     * 格式化工具错误，展示：错误类型徽章 + 原始错误 + 原因 + 修复建议
     * @param {string} errorText - 原始错误文本
     * @param {Object} [meta] - 可选的错误分类信息
     * @param {string} [meta.toolName] - 工具名
     * @param {string} [meta.errorType] - 错误类型
     * @param {string} [meta.errorReason] - 错误原因
     * @param {string} [meta.errorSuggestion] - 修复建议
     */
    function formatToolErrorHtml(errorText, meta) {
        return App.ChatV2.renderers.toolCard.formatToolErrorHtml(errorText, meta);
    }

    /**
     * 为未被 classifyError 覆盖的通用错误生成兜底建议
     */
    function _generateGenericSuggestion(errText, toolName) {
        return App.ChatV2.renderers.toolCard._generateGenericSuggestion(errText, toolName);
    }

    /**
     * 通用对话里模型可能"只写计划不真执行"：把正文中【执行】[工具] / 【执行】工具 标注提取为"规划态"步骤。
     * - 返回 { text, steps }：text 为清洗掉声明占位（保留其后自然语言描述）后的干净正文；
     *   steps 为解析出的规划步骤（status='pending'）。
     * - 仅当本轮【没有任何真实执行工具】时才产出 steps：若工具确实执行过（有已完成调用），
     *   其操作已由工具卡片（✅完成）表达，再补"待执行"绿面板会语义冲突；此时只清洗正文不造面板。
     * 【执行】[tool] 是模型自发的声明式标注（非系统协议），工具名可能并不存在、不会真正执行；
     * 用"待执行"态渲染成绿面板，诚实反映"这是模型规划、尚未真实执行"。
     */
    function prepareConversationContent(turn, text) {
        if (!text || typeof text !== 'string') return { text: text || '', steps: [] };
        // 模型对"计划步骤"的标注写法不统一，常见：
        //   ① 括号式【执行】[工具名] ② 无括号式【执行】工具名（工具名是 ASCII 标识）
        // 1. 清洗正文：按优先级移除「【执行】+工具名」或残留的【执行】分隔符，保留其后的自然语言描述
        const cleanedText = text
            .replace(/\u3010\u6267\u884c\u3011\s*\[[^\]]+\]/g, '')            // 【执行】[工具名]
            .replace(/\u3010\u6267\u884c\u3011\s*[A-Za-z_][A-Za-z0-9_.-]*/g, '') // 【执行】工具名
            .replace(/\u3010\u6267\u884c\u3011/g, '');                          // 兜底 bare/分隔符【执行】
        // 2. 只有本轮没有任何"真实执行"工具时才生成"规划态"步骤，
        //    避免与真实工具卡片（✅完成）冲突而出现"待执行"面板。
        //    注：仅对"带工具名锚定的声明"泛面板化；任意正文（如安全守卫提示、徽章 HTML）绝不据此造步骤。
        let steps = [];
        const hasRealExecution = turn && Array.isArray(turn.toolCalls) &&
            turn.toolCalls.some(tc => tc && tc.status === 'completed');
        const hasRealSteps = (turn && Array.isArray(turn._steps) && turn._steps.length) || hasRealExecution;
        if (!hasRealSteps) {
            const markerRe = /\u3010\u6267\u884c\u3011\s*(?:\[([^\]]+)\]|([A-Za-z_][A-Za-z0-9_.-]*))/g;
            const matches = [];
            let m;
            markerRe.lastIndex = 0;
            while ((m = markerRe.exec(text)) !== null) {
                matches.push({ start: m.index, end: markerRe.lastIndex, tool: (m[1] || m[2] || '') });
            }
            if (matches.length) {
                steps = matches.map((cur, i) => {
                    const nextStart = (i + 1 < matches.length) ? matches[i + 1].start : text.length;
                    const desc = text.slice(cur.end, nextStart).trim();
                    // 步骤标题一律 HTML 转义，防止描述/徽章等原始 HTML 在面板里被当源码渲染
                    return { title: desc ? escapeHtml(desc) : (cur.tool || undefined), status: 'pending' };
                });
            }
        }
        return { text: cleanedText, steps };
    }

    /**
     * 从工具执行结果中提取"步骤规划"数组。
     * 兼容几种存放位置：顶层 steps / output.executionResult.steps / executionResult.steps 等。
     * @param {Object} result 原始执行结果（可为字符串/对象/包裹层）
     * @returns {Array<string>|null} 步骤数组；未找到返回 null
     */
    // 实现已拆分到 renderers/steps-plan.js（纯函数，可单测）
    function extractStepsPlan(result) {
        return App.ChatV2.renderers.stepsPlan.extractStepsPlan(result);
    }

    /**
     * 把步骤数组渲染成醒目的分步清单（带序号+完成标识）。
     */
    // 实现已拆分到 renderers/steps-plan.js（纯函数，可单测）
    function renderStepsPlanHtml(steps) {
        return App.ChatV2.renderers.stepsPlan.renderStepsPlanHtml(steps);
    }

    /** 收集过程列表各行的展开态 + 推理内容滚动位置（供流式全量重渲后恢复） */
    function collectCbxOpenStates(rootEl) {
        const states = {};
        if (!rootEl) return states;
        rootEl.querySelectorAll('details.cbx-item').forEach((d, i) => {
            const body = d.querySelector('.cbx-detail--text');
            if (!body) { states[i] = { open: !!d.open, wasRunning: d.classList.contains('cbx-running') }; return; }
            const distFromBottom = body.scrollHeight - body.scrollTop - body.clientHeight;
            // running 行：默认跟随底部；仅当用户明显上翻（离开底部）时才停止跟随
            const stick = d.classList.contains('cbx-running')
                ? !(body.scrollTop > 1 && distFromBottom > 40)
                : false;
            // wasRunning：收集时该行是否"执行中"——恢复时据此判断"执行中展开→完成后应收起"
            states[i] = { open: !!d.open, wasRunning: d.classList.contains('cbx-running'), scrollTop: body.scrollTop, stick: stick };
        });
        return states;
    }

    /** 恢复过程列表各行的展开态与滚动位置；running 行贴底跟随最新推理内容 */
    function restoreCbxOpenStates(rootEl, states) {
        if (!rootEl) return;
        rootEl.querySelectorAll('details.cbx-item').forEach((d, i) => {
            const body = d.querySelector('.cbx-detail--text');
            const s = states ? states[i] : null;
            // 【完成即折叠】收集时"执行中"、现已完成的行 → 强制收起。
            // 否则流式重渲会把执行期间的展开态原样回写，步骤完成后永远不折叠。
            // 用户手动展开的行（收集时已完成）不受影响，照常恢复。
            if (s && s.wasRunning && !d.classList.contains('cbx-running')) {
                d.open = false;
            } else if (s && Object.prototype.hasOwnProperty.call(states, i)) {
                d.open = !!s.open;
            }
            if (!body) return;
            if (d.classList.contains('cbx-running') && d.open &&
                (!s || s.stick)) {
                // 跟随模式：钉在底部看最新推理
                body.scrollTop = body.scrollHeight;
            } else if (s && s.scrollTop != null) {
                body.scrollTop = s.scrollTop;
            }
        });
    }

    /** 判断是否为"统一过程列表"承载的 text 管线 turn（thoughtList 非空）。
     *  这类 turn 的过程展示统一走 contentEl 顶部的 CodeBuddy 风格扁平列表，
     *  toolsEl 的旧规划卡/工具卡全部抑制，避免同一过程出现三份展示。 */
    function isTextProcessTurn(turnId) {
        try {
            const t = (App.ChatV2?.sessionStore?.getTurns?.() || []).find(x => x.id === turnId);
            if (!t || !Array.isArray(t.thoughtList) || t.thoughtList.length === 0) return false;
            // 文本管线：一旦有 thoughtList，过程只走这一条，不再另开工具卡。
            if (t.kind === 'text') return true;
            // 技能/代码：只有真正的步骤或工具块才独占通道。
            // 仅有路由思考时仍允许工具卡，避免执行记录被收掉。
            return t.thoughtList.some(th => th && (
                th.kind === 'tool' || th.kind === 'step' || th.kind === 'skill_step' || th.kind === 'skill'
            ));
        } catch (e) { return false; }
    }

    /**
     * 渲染通道的可见性只在这里算一次。
     * turn_start、kind 变更、_enableSteps 补丁必须用同一套规则——
     * 以前 _enableSteps 补丁把所有轮次都当成文本，技能轮会被误收起。
     */
    function applyRenderChannel(handle, turn) {
        if (!handle) return;
        const isText = !!(turn && turn.kind === 'text');
        const hasSteps = !!(turn && (turn._enableSteps || (turn.context && turn.context._enableSteps)));
        handle._multiStep = hasSteps;
        handle._suppressSteps = isText && !hasSteps;
        handle._suppressToolCards = !hasSteps;
        handle.finalRenderOnly = !hasSteps;
    }

    // ── 结果展示格式解析：AI 选择渲染器 + 确定性兜底 ────────────────────
    // 优先级：① 用户原话里点名的展示格式（最可靠、且诉求必须被尊重）
    //         ② 模型在 args 里显式传的 _view（本环境本地模型常不输出，仅作兜底）
    //         ③ 技能 outputFormat。
    // 说明：本地模型不一定输出 schema 之外的 _view（实测 2026-10-03），
    // 因此加了 ② 这条确定性兜底，保证"用户要表格就给表格"。
    const VIEW_HINT_KEYWORDS = [
        { view: 'table', re: /表格|table/i },
        { view: 'json', re: /json/i },
        { view: 'code', re: /代码块|代码|code/i },
        { view: 'markdown', re: /markdown|\bmd\b/i },
        { view: 'steps', re: /分步|步骤|清单|steps/i },
        { view: 'text', re: /纯文本/ }
    ];
    // 必须同时出现"展示动词"，避免误伤（如"读取表格A1:C10"只提到表格、并非要求展示格式）
    const VIEW_HINT_DEMAND = /展示|显示|呈现|输出|渲染|格式|整理|做成|转成|生成|查看|看看|列出|列举|列成/;

    /** 从用户原话解析"要求的结果展示格式"；多个格式词并存时取最先提到的
     *  （"把结果用表格/json/代码块展示" → table）。无明确诉求返回 null。 */
    function detectViewFromMessage(msg) {
        if (!msg || typeof msg !== 'string') return null;
        if (!VIEW_HINT_DEMAND.test(msg)) return null;
        let best = null;
        let bestIdx = Infinity;
        for (let i = 0; i < VIEW_HINT_KEYWORDS.length; i++) {
            const k = VIEW_HINT_KEYWORDS[i];
            const m = msg.match(k.re);
            if (m && m.index < bestIdx) { bestIdx = m.index; best = k.view; }
        }
        return best;
    }

    /** 解析一次工具调用最终生效的展示格式（无则 null）。tc 需含 name/args。 */
    function resolveToolViewHint(tc, userMessage) {
        if (!tc) return null;
        const registry = App.ChatV2?.toolRegistry;
        let modelView = null;
        if (registry && typeof registry.getViewTypeFromArgs === 'function') {
            try { modelView = registry.getViewTypeFromArgs(tc.args); } catch (e) { modelView = null; }
        }
        const userView = detectViewFromMessage(userMessage);
        const skillView = resolveSkillView(tc);
        // 用户显式点名优先：本环境模型常不输出 _view，且用户诉求必须被尊重；
        // 模型 _view 仅在用户未指定时生效，技能 outputFormat 作为最后兜底。
        return userView || modelView || skillView || null;
    }

    /** 取某轮的用户原话（展示格式兜底用）；turnId 缺失时回落到当前轮 */
    function userMessageForTurn(turnId) {
        try {
            const ss = App.ChatV2?.sessionStore;
            if (!ss) return '';
            let t = null;
            if (turnId && typeof ss.getTurns === 'function') {
                t = (ss.getTurns() || []).find(x => x && x.id === turnId);
            }
            if (!t && typeof ss.getCurrentTurn === 'function') t = ss.getCurrentTurn();
            if (!t) return '';
            return (t.userMessage || t._originalUserMessage || '');
        } catch (e) { return ''; }
    }

    /** 汇总"本轮用户原话"：多来源兜底。
     *  渲染层拿到的 turn 可能是回放/历史/事件里的副本，未必带 userMessage，
     *  故依次尝试 _originalUserMessage / userMessage / message → sessionStore(按 id、当前轮)
     *  → 会话历史最近一条 user 消息。保证"用户点名展示格式"这一判定不会凭空丢失。 */
    function turnUserText(turn) {
        try {
            const parts = [];
            if (turn) {
                if (turn._originalUserMessage) parts.push(turn._originalUserMessage);
                if (turn.userMessage) parts.push(turn.userMessage);
                if (turn.message) parts.push(turn.message);
                if (turn.id) {
                    const fromStore = userMessageForTurn(turn.id);
                    if (fromStore) parts.push(fromStore);
                }
            }
            if (!parts.length) {
                const ss = App.ChatV2?.sessionStore;
                const hist = (ss && typeof ss.get === 'function') ? ss.get('history') : null;
                if (Array.isArray(hist)) {
                    for (let i = hist.length - 1; i >= 0; i--) {
                        const h = hist[i];
                        if (h && h.role === 'user' && (h.originalContent || h.content)) {
                            parts.push(h.originalContent || h.content);
                            break;
                        }
                    }
                }
            }
            return parts.filter(Boolean).join('\n');
        } catch (e) { return ''; }
    }

    const VIEW_HINT_LABELS = { table: '表格视图', json: 'JSON 视图', code: '代码视图', markdown: 'Markdown 视图', steps: '步骤视图', text: '纯文本视图' };

    /** 本轮是否存在展示诉求（任一工具调用命中即可） */
    function turnViewHint(turn, userText) {
        const calls = (turn && Array.isArray(turn.toolCalls)) ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
        const um = (userText != null) ? userText : turnUserText(turn);
        for (let i = 0; i < calls.length; i++) {
            if (resolveToolViewHint(calls[i], um)) return true;
        }
        return false;
    }

    /** 终态兜底：命中展示诉求时仅强制展开【外层汇总盒】（任务完成 · 结果视图）。
     *  【完成即折叠·2026-10-06】不再强制展开任何工具行（旧版对 data-view-hint 行
     *  d.open=true，会在终态渲染时把我方"确认完成即收起"的折叠全部覆盖回去——
     *  这正是"总是有部分不能折叠"的真凶）。步骤行一律保持收起，点击可展开。 */
    function enforceViewHintOpen(rootEl, turn) {
        try {
            if (!rootEl || !turn) return;
            if (!turnViewHint(turn)) return;
            const box = rootEl.querySelector('details.cbx-proc-box');
            if (box) box.open = true;
        } catch (e) { /* 兜底失败不影响渲染 */ }
    }

    /** 判断 turn 是否真的"做过事"（有工具调用 / 步骤 / Skill 执行）。
     *  纯寒暄、知识问答这类无执行 turn 不展示过程列表与思考过程，避免"你好"也挂一串思考行。 */
    function hasRealExecution(turn) {
        if (!turn) return false;
        const calls = Array.isArray(turn.toolCalls) ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
        if (calls.some(tc => tc.status === 'completed' || tc.status === 'failed' || tc.status === 'running')) return true;
        if (Array.isArray(turn._steps) && turn._steps.length > 0) return true;
        if (turn._planHtml) return true;
        if (turn._matchedSkill || (turn.context && turn.context._matchedSkill)) return true;
        const list = Array.isArray(turn.thoughtList) ? turn.thoughtList : [];
        if (list.some(th => th && (th.kind === 'tool' || th.kind === 'step' || th.kind === 'skill'))) return true;
        // 路由期"意图分析"块也算真实活动：否则从发送到第一个正文 token 之间
        // （阻塞式 LLM 意图分析最长 3s 超时 + agent-loop 首 token 延迟）全程只有三个点。
        // 纯寒暄 turn 不会创建 routing 块（planCheck 未命中时直接跳过 LLM 分析），不受影响。
        if (list.some(th => th && th.kind === 'routing')) return true;
        // 一键编排工具（如 excel_generate_dashboard）本身即代表真实产出
        if (calls.some(tc => tc && String(tc.name || '').indexOf('generate_dashboard') >= 0)) return true;
        return false;
    }

    /**
     * 【终态汇总盒】turn 结束后把过程列表折叠为一个汇总盒：
     * 「任务完成 · 执行 N 个工具 · 耗时 X」——默认收起，点击展开明细行。
     * 流式期（status 为 running 等）保持平铺渐进展示，不包装。
     */
    function wrapProcessSummaryBox(turn, listHtml) {
        if (!turn || !listHtml) return listHtml || '';
        const st = String(turn.status || '');
        if (st !== 'completed' && st !== 'failed' && st !== 'canceled') return listHtml;
        const calls = Array.isArray(turn.toolCalls) ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
        let durMs = (turn.completedAt && turn.startedAt) ? Math.max(0, turn.completedAt - turn.startedAt) : 0;
        if (!durMs && calls.length) {
            durMs = calls.reduce((s, tc) => s + (Number(tc.durationMs) || 0), 0);
        }
        const durText = durMs >= 1000 ? ((durMs / 1000).toFixed(1) + ' 秒') : (durMs + ' 毫秒');
        const label = st === 'completed' ? '任务完成' : (st === 'failed' ? '任务失败' : '已取消');
        // 步骤面板承载的确定性链（批量美化等）不补裸工具卡 → 没有工具调用但有步骤行，
        // 此时摘要按步骤数显示，避免自相矛盾的"未调用工具"（实测 2026-10-22 22:31）。
        const stepCount = Array.isArray(turn.thoughtList)
            ? turn.thoughtList.filter(t => t && t.kind === 'step').length
            : 0;
        const toolText = calls.length > 0
            ? ('执行 ' + calls.length + ' 个工具')
            : (stepCount > 0 ? ('完成 ' + stepCount + ' 个步骤') : '未调用工具');
        // 【Token 用量】真实统计显示 入/出；仅估算时标注 ≈与(估算)
        let tokenText = '';
        const u = turn.usage;
        if (u && (Number(u.in) > 0 || Number(u.out) > 0)) {
            const fmt = n => n >= 10000 ? ((n / 1000).toFixed(1) + 'k') : String(n);
            tokenText = u.estimated
                ? (' · ≈' + fmt(Number(u.out) || 0) + ' token(估算)')
                : (' · token ' + fmt(Number(u.in) || 0) + '入/' + fmt(Number(u.out) || 0) + '出');
        }
        const boxCls = st === 'completed' ? 'cbx-proc-box' : ('cbx-proc-box cbx-proc-' + st);
        // 用户点名了展示格式（或模型/技能指定了 _view）→ 汇总盒默认展开，
        // 否则渲染好的表格/JSON/代码块会被折叠在「任务完成 ›」里，用户根本看不见。
        const _um = turnUserText(turn);
        const _hintOn = turnViewHint(turn, _um);
        const _hintOpen = _hintOn ? ' open' : '';
        // 命中展示诉求时在摘要里标注视图类型：既是提示，也用于确认判定链路是否生效
        const _hintMark = _hintOn ? (' · ' + (VIEW_HINT_LABELS[detectViewFromMessage(_um)] || '结果视图')) : '';
        return '<details class="' + boxCls + '"' + _hintOpen + '>' +
            '<summary>' +
            '<span class="cbx-proc-title">' + escapeHtml(label) + '</span>' +
            '<span class="cbx-proc-meta">' + escapeHtml(toolText + ' · 耗时 ' + durText + tokenText + _hintMark) + '</span>' +
            '<span class="cbx-proc-arrow">›</span>' +
            '</summary>' +
            '<div class="cbx-proc-box-body">' + listHtml + '</div>' +
            '</details>';
    }

    /**
     * 步骤行（图钉）和工具行是同一次调用的两份记录。
     * 工具行已经带参数、耗时和返回结果，同名步骤行就不再进列表。
     */
    function stepCoveredByToolRow(th, list) {
        if (!th || th.kind !== 'step') return false;
        const name = String(th.toolName || cbxCleanLabel(th.title) || '').trim().toLowerCase();
        if (!name) return false;
        return (list || []).some(function (other) {
            if (!other || other === th || other.kind !== 'tool') return false;
            const tn = String(other.toolName || cbxCleanLabel(other.title) || '').trim().toLowerCase();
            return tn === name;
        });
    }

    /** 清洗 thought 块标题为纯动作名（去 emoji、**加粗**、结尾 ✓/状态尾巴） */
    function cbxCleanLabel(title) {
        let s = String(title || '').replace(/\*\*/g, '');
        s = s.replace(/^[\p{Extended_Pictographic}\uFE0F\u200D\s]+/u, '');
        s = s.replace(/\s*[✓✕✔■].*$/, '');
        s = s.replace(/\s*[—–-]\s*执行中$/,'');
        return s.trim();
    }

    /** 渲染 thoughtList 为 CodeBuddy 风格扁平活动列表（迭代替换模型唯一过程展示）：
     *  每行 = 图标 + 动作名 + 状态，点击整行展开推理文字/结果详情；
     *  running 行默认展开，completed/failed 默认收起。 */
    function renderThoughtListHtml(turn) {
        const list = turn && Array.isArray(turn.thoughtList) ? turn.thoughtList : null;
        if (!list || list.length === 0) return '';
        // 本轮用户原话：用于展示格式兜底（用户点名"表格/json/代码块"时按对应渲染器输出并默认展开）
        const turnUserMessage = turnUserText(turn);

        // ── 工具行详情数据源 ──
        // thought 块只记录"标题/状态/推理文字"，没有工具的参数与返回结果；
        // 展开区要能看到"参数 / 返回结果"，必须从 turn.toolCalls 把真实数据回填到对应行。
        const calls = (turn && Array.isArray(turn.toolCalls))
            ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
        const toolThoughtIdx = [];
        for (let i = 0; i < list.length; i++) {
            if (list[i] && list[i].kind === 'tool') toolThoughtIdx.push(i);
        }
        const usedCallIdx = new Set();
        const mapCallFor = (th, orderInTools) => {
            if (calls.length === 0) return null;
            // 工具行数与调用数一致 → 按序一一对应（最可靠）
            // 【必须标记占用】否则末尾"补全未占用工具"循环会把同一调用再渲一遍（两行一模一样）
            if (toolThoughtIdx.length === calls.length) {
                const byOrder = calls[orderInTools] || null;
                if (byOrder) usedCallIdx.add(orderInTools);
                return byOrder;
            }
            // 否则按工具名匹配（取第一个未被占用的同名调用）
            const name = String(th.toolName || cbxCleanLabel(th.title) || '').toLowerCase();
            let hit = -1;
            for (let i = 0; i < calls.length; i++) {
                if (usedCallIdx.has(i)) continue;
                const c = calls[i];
                const cn = String(c.name || c.displayName || '').toLowerCase();
                if (cn && (cn === name || cn.indexOf(name) >= 0 || name.indexOf(cn) >= 0)) { hit = i; break; }
            }
            if (hit < 0) {
                for (let i = 0; i < calls.length; i++) {
                    if (!usedCallIdx.has(i)) { hit = i; break; }
                }
            }
            if (hit < 0) return null;
            usedCallIdx.add(hit);
            return calls[hit];
        };

        // ── 预配对：按渲染顺序把 tool thought 块与 toolCall 一一配对（只标记占用，不渲染）。
        // 渲染循环里不能再调 mapCallFor（重复消费会错位），统一从 thoughtCallMap 取。
        const thoughtCallMap = new Map();
        {
            let seq = 0;
            for (const th of list) {
                if (th && th.kind === 'tool') thoughtCallMap.set(th, mapCallFor(th, seq++));
            }
        }
        // 孤儿调用：未被任何 tool 块占用的 toolCalls（某些快速路径只记 toolCall 不记 thought 块）
        const orphanCalls = [];
        for (let i = 0; i < calls.length; i++) {
            if (!usedCallIdx.has(i) && calls[i]) orphanCalls.push(calls[i]);
        }

        // ── 时间线归并：孤儿调用按 round 插入 thought 块之间，保持与旧版一致的
        // "思考 ↔ 执行"交错时间线，而不是全部堆在末尾；无 round 的孤儿排最后。
        // JS Array.prototype.sort 为稳定排序（ES2019+），同 round 保持插入顺序。
        const timeline = [];
        for (const th of list) {
            if (!th) continue;
            // 图钉步骤行只是同一次工具调用的第二份勾。已经有工具行时不再画。
            if (stepCoveredByToolRow(th, list)) continue;
            timeline.push({ round: Number(th.round) || 0, th: th });
        }
        for (const c of orphanCalls) {
            timeline.push({ round: (c && c.round != null) ? Number(c.round) : Infinity, call: c });
        }
        timeline.sort((a, b) => a.round - b.round);

        let html = '<div class="cbx-list cbx-list--process">';
        let _rowCount = 0;
        for (const item of timeline) {
            // ── 孤儿工具调用行（没有对应 thought 块的执行记录）──
            if (item.call) {
                const c = item.call;
                const cBad = c.status === 'failed';
                const cRunning = (c.status === 'running' || c.status === 'pending');
                const cOpen = cRunning ? ' open' : '';
                const cName = cbxCleanLabel(c.displayName || c.name || '工具调用');
                const cIcon = cbxIconForTool(c.name || c.displayName || '');
                const cSt = cBad
                    ? '<span class="cbx-status bad">✕ 失败</span>'
                    : (c.status === 'completed' ? '<span class="cbx-status ok">✓</span>' : '<span class="cbx-status">⏳ 执行中</span>');
                const cDur = (c.durationMs != null) ? '<span class="cbx-dur">' + escapeHtml(String(c.durationMs)) + 'ms</span>' : '';
                const cTarget = escapeHtml(cbxArgsSummary(c));
                const cDetail = buildToolCallDetailHtml(c, turnUserMessage);
                // 【完成即折叠】孤儿工具行同样只在执行中（cOpen）展开；命中 _view 不再自动展开
                html += '<details class="cbx-item cbx-tool' + (cBad ? ' cbx-failed' : '') + (cRunning ? ' cbx-running' : '') + '"' + cOpen + '>' +
                    '<summary>' +
                        '<span class="cbx-icon">' + cIcon + '</span>' +
                        '<span class="cbx-label">' + escapeHtml(cName) + '</span>' +
                        '<span class="cbx-target">' + cTarget + '</span>' +
                        '<span class="cbx-meta">' + cSt + cDur + '</span>' +
                        '<span class="cbx-chevron">›</span>' +
                    '</summary>' +
                    '<div class="cbx-detail">' + (cDetail || '<div class="cbx-detail--empty">（无返回内容）</div>') + '</div>' +
                '</details>';
                _rowCount++;
                continue;
            }
            const th = item.th;
            const isTool = (th.kind === 'tool');
            // 【步骤统一到 cbx v20261003】多步骤 / Skill / 计算机链的步骤块（kind='step'/'skill_step'）
            // 与工具行一样属于"真实执行记录"，不受下面的空思考块过滤影响 ——
            // 否则"已完成但没有详情文字"的步骤会整行消失，用户看不到执行过什么。
            const isStep = (th.kind === 'step' || th.kind === 'skill_step');
            const running = th.status === 'running';
            // 【空思考块过滤】已完成、既无推理文字、又非工具/步骤行的思考/路由块不再渲染——
            // agent-loop 每轮都会 startThought('🔍 正在思考...')，收尾时标题被改成"🔍 思考过程"；
            // 若本轮模型没走 thinking 通道（如"仅展示"场景只回了"已展示"），该块内容为空，
            // 以前会固定挂出一行"思考过程 →（无推理文字）"的空壳（实测反馈："为什么每次后面还有空的思考过程"）。
            // running 行仍保留（"正在思考…"是有效活动），有内容的行仍保留。
            if (!isTool && !isStep && !running && !(th.content && String(th.content).trim())) continue;
            const bad = th.status === 'failed';
            let openAttr = running ? ' open' : '';
            const toolName = isTool ? (cbxCleanLabel(th.toolName) || cbxCleanLabel(th.title) || '工具调用') : '';
            const label = escapeHtml(isTool ? toolName : (cbxCleanLabel(th.title) || '思考中...'));
            const icon = isTool ? cbxIconForTool(th.toolName || th.title || '') : (isStep ? '📌' : '💭');
            const call = isTool ? (thoughtCallMap.get(th) || null) : null;
            // 【完成即折叠】thought 工具行不再因命中 _view 而自动展开（2026-10-06 规则）：
            // 仅 running 行保持 open；确认完成即收起，点 summary 可手动展开。
            const st = bad
                ? '<span class="cbx-status bad">✕ 失败</span>'
                : (th.status === 'completed'
                    ? '<span class="cbx-status ok">✓</span>'
                    : '<span class="cbx-status">⏳ 执行中</span>');
            const durMs = (th.durationMs != null) ? th.durationMs : (call ? call.durationMs : null);
            const dur = (durMs != null)
                ? '<span class="cbx-dur">' + escapeHtml(String(durMs)) + 'ms</span>' : '';
            // 目标摘要（工具行）：参数里的关键值，如 [数据源 1]
            const target = (isTool && call) ? escapeHtml(cbxArgsSummary(call)) : '';

            // ── 展开区 ──
            // 工具行：参数/返回结果/错误（与工具卡"详情"一致） + 该步推理/结果提示文字
            // 思考行：推理文字；无内容时给占位，避免"点了没反应"
            let bodyHtml = '';
            const callDetail = (isTool && call) ? buildToolCallDetailHtml(call, turnUserMessage) : '';
            const textBody = th.content ? escapeHtml(String(th.content)) : '';
            if (callDetail) bodyHtml += callDetail;
            if (textBody) {
                bodyHtml += '<div class="cbx-detail--text">' + textBody + '</div>';
            } else if (!callDetail) {
                bodyHtml = '<div class="cbx-detail--empty">' + (isTool ? '（无返回内容）' : (isStep ? '（无详细信息）' : '（无推理文字）')) + '</div>';
            } else if (isTool && !call.result) {
                bodyHtml += '<div class="cbx-detail--empty">（该工具无返回内容）</div>';
            }

            html += '<details class="cbx-item cbx-thought' + (isTool ? ' cbx-tool' : '') +
                (bad ? ' cbx-failed' : '') + (running ? ' cbx-running' : '') + '"' + openAttr + '>' +
                '<summary>' +
                    '<span class="cbx-icon">' + icon + '</span>' +
                    '<span class="cbx-label">' + label + '</span>' +
                    '<span class="cbx-target">' + target + '</span>' +
                    '<span class="cbx-meta">' + st + dur + '</span>' +
                    '<span class="cbx-chevron">›</span>' +
                '</summary>' +
                '<div class="cbx-detail">' + bodyHtml + '</div>' +
            '</details>';
            _rowCount++;
        }
        // 所有行都被过滤掉（如只剩空思考块）→ 不渲染空容器，交由调用方兜底
        if (_rowCount === 0) return '';
        html += '</div>';
        return html;
    }

    // 【步骤统一到 cbx v20261003】旧的"步骤规划"绿面板补丁函数已随 .exec-plan-card 一并下线删除：
    //   collectPlanOpenStates / restorePlanOpenStates / applyPlanProgress / applyStepPatch
    // 步骤现在走 thoughtList（kind='step'），由 renderThoughtListHtml 统一渲染，
    // 渲染层收到 thought_update 即全量重渲，不再需要按 data-step-id 打 DOM 补丁。

    /**
     * 检测后端占位消息（Skill 未真正执行，仅返回"已识别/由前端处理"提示）。
     * 命中时返回友好提示文案，避免在工具卡片"返回结果"区泄露原生 JSON。
     * @param {*} data unwrapToolResult 之后的数据
     * @returns {string|null} 占位提示文案；非占位返回 null
     */
    function extractPlaceholderMessage(data) {
        return App.ChatV2.renderers.toolCard.extractPlaceholderMessage(data);
    }

    /**
     * 从操作类技能结果中提取可读摘要（用于工具卡片常显区）。
     * 剥离 "Skill 'xxx' Hybrid 执行成功：" 前缀与 "执行步骤：..." 列表
     * （步骤已由绿面板展示，避免重复），只保留结果摘要，
     * 如 "✓ 合并所有Sheet完成\n已合并8张工作表：...\n合并结果共 239 行（含表头）"。
     * @param {*} rawResult 工具原始结果
     * @returns {string} 摘要文本；无可用摘要返回空串
     */
    function extractSkillResultSummary(rawResult) {
        return App.ChatV2.renderers.toolCard.extractSkillResultSummary(rawResult);
    }

    /**
     * 操作类结果（目前是"合并工作表"）的友好渲染：
     *   ① 统计条（合并 N 张表 → 共 M 行 / 表头已保留 / 目标表）
     *   ② 数据表（复用 tryRowsMatrixTable：{success,data:[[表头],[行]…]} → 表格，超限截断）
     *   ③ 折叠的原始 JSON（需要核对细节时展开，不占视觉主位）
     * @param {*} data 解包后的结果（unwrapToolResult 可能只剩数据矩阵）
     * @param {{text:string, chips:Array<{label:string,value:string}>, source:object}} sum summarizeMergeFromRaw 的返回
     * @param {*} rawResult 未解包的原始结果（可能还被 {result:…} 包着）
     */
    function renderOperationResultHtml(data, sum, rawResult) {
        // 不可见自检标记：合并结果卡片命中此渲染分支时植入，便于排查"仍是裸矩阵"时确认新代码是否已加载。
        // 右键结果区 → 检查元素 → 若见 <!--op-merge-render:...--> 说明新渲染逻辑已生效（问题在数据结构）；
        // 若仍只见裸 JSON，说明仍跑旧代码（需完全退出 Excel 重开 / 硬刷新任务窗格）。
        const _marker = '<!--op-merge-render:' + (window.__APP_VERSION__ || 'na') + '-->';
        const chips = (sum && Array.isArray(sum.chips) ? sum.chips : []).map(function (c) {
            return '<span class="exec-op-chip">'
                + '<span class="exec-op-chip-label">' + escapeHtml(c.label) + '</span>'
                + '<span class="exec-op-chip-value">' + escapeHtml(c.value) + '</span>'
                + '</span>';
        }).join('');
        // 命中对象（含 sourceSheets/mergedRows 的那一层，通常也带 data 矩阵）优先：
        // 解包结果可能只剩矩阵，外层统计字段只在命中对象上 —— 表格与折叠 JSON 都以它为首选。
        const src = (sum && sum.source && typeof sum.source === 'object' && !Array.isArray(sum.source)) ? sum.source : null;
        let table = '';
        try { table = tryRowsMatrixTable(src) || tryRowsMatrixTable(data) || tryRowsMatrixTable(rawResult) || ''; } catch (e) { table = ''; }
        // 折叠 JSON：优先"含统计字段的那一层"，否则用户展开只看到一坨没有来源/行数信息的矩阵
        const jsonSrc = src || ((rawResult && typeof rawResult === 'object' && !Array.isArray(rawResult)) ? rawResult : data);
        const pretty = JSON.stringify(jsonSrc, null, 2);
        const truncated = pretty.length > 2000
            ? pretty.slice(0, 2000) + '\n...(结果已截断，共 ' + pretty.length + ' 字符)'
            : pretty;
        return _marker + '<div class="exec-op-summary">'
            + '<div class="exec-op-title">✅ ' + escapeHtml((sum && sum.text) || '操作完成') + '</div>'
            + (chips ? '<div class="exec-op-chips">' + chips + '</div>' : '')
            + '</div>'
            + (table ? '<div class="py-data-table-wrapper exec-op-table">' + table + '</div>' : '')
            + renderCollapsibleJson(truncated);
    }

    function formatToolResultHtml(result, viewArgs) {
        let data = unwrapToolResult(result);

        // 步骤规划绿面板已由工具卡片"常显区"统一渲染（updateToolCard 中 renderStepsPlanHtml），
        // 此处 JSON 路径不再重复渲染步骤，避免同一份步骤出现两次；
        // 仅占位消息 / 纯字符串路径保留 planHtmlP（这两种场景不经过常显区步骤渲染）。
        const planStepsP = extractStepsPlan(data);
        const planHtmlP = (planStepsP && planStepsP.length) ? renderStepsPlanHtml(planStepsP) : '';

        // 后端占位消息 → 显示友好提示，不泄露原生 JSON
        const placeholderMsg = extractPlaceholderMessage(data);
        if (placeholderMsg) {
            return planHtmlP + '<div style="padding:8px 10px;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;font-size:12px;color:#555;">' +
                escapeHtml(placeholderMsg) + '</div>';
        }

        // Skill 执行结果：outputs.content 是 skill_card 时，**直接把卡片渲染到工具结果区**。
        // 为什么不能只靠主消息区渲染：走 agent-loop 调用 skill.SkillCreator（text-pipeline 通道）时，
        // 主消息区只有 LLM 的总结文本（"已成功创建…"），skill_card 仅存在于工具结果里。
        // 原先这里把 content 换成 "[Skill 卡片：xxx]" 占位 + 折叠原始 JSON：
        // 用户既看不到卡片，也拿不到"保存"按钮（实测 2026-10-02 21:59）。
        const skillCard = extractSkillCardFromResult(result);
        if (skillCard && data && typeof data === 'object') {
            // 卡片本体的渲染交给"最后的结果"区：tool-registry.appendSkillCardToTurn 会把 skill_card
            // 追加到 turn 正文末尾，由 findSkillCardJson → renderSkillCardFromJson 渲染成完整卡片
            // （含步骤预览与保存按钮）。此处只给摘要 + 折叠原始结果，避免同一卡片出现两次。

            // 附带折叠的原始结果（便于审查细节；默认收起）
            const display = JSON.parse(JSON.stringify(data));
            if (display.outputs && typeof display.outputs === 'object') {
                if (display.outputs.skillCardJson) delete display.outputs.skillCardJson;
                if (display.outputs.generatedSkillJson) delete display.outputs.generatedSkillJson;
            }
            const pretty = JSON.stringify(display, null, 2);
            const truncated = pretty.length > 2000
                ? pretty.slice(0, 2000) + '\n...(结果已截断，共 ' + pretty.length + ' 字符)'
                : pretty;
            const _dn = escapeHtml(skillCard.skillDisplayName || skillCard.skillName || '');
            const _ct = escapeHtml(skillCard.skillCategory || '');
            const _summary = '<div style="padding:8px 10px;background:#f6f8fa;border:1px solid #e1e4e8;border-radius:6px;font-size:12px;color:#555;">'
                + '已生成 Skill：' + _dn + (_ct ? '（' + _ct + '）' : '') + '，卡片见下方最终结果</div>';
            // 不再拼接 planHtmlP：skill_card 携带的 steps 是"新技能的步骤定义"，不是本轮真实执行进度，
            // 被 extractStepsPlan 当执行计划渲染成「0/1 步完成（0%）等待」的绿面板，既误导又是冗余
            // （真实预览在下方最终结果的卡片里）。工具结果区只保留摘要 + 折叠原始结果。
            return _summary + renderCollapsibleJson(truncated);
        }

        // ── 操作结果（合并工作表等）：统计摘要 + 数据表 + 折叠原始 JSON ──
        // 【必须排在下面 view 提示分支之前】合并结果的 data 矩阵（100 行 × 9 列）一旦被
        // _vt === 'json' / 'code' 分支接走，就会原样摊成一屏原始 JSON（实测 2026-10-07 用户反馈
        // "渲染有点不好 / 没变化"）。操作语义优先于通用视图提示：
        //   ① 一行统计（合并了几张表 / 多少行 / 表头是否保留）
        //   ② 合并后的数据渲染成可读表格（首行即表头，行数上限由 tryRowsMatrixTable 控制）
        //   ③ 完整原始 JSON 收进"查看原始结果（JSON）"折叠面板
        const _opSum = (App.ChatV2?.renderers?.toolCard?.summarizeMergeFromRaw)
            ? App.ChatV2.renderers.toolCard.summarizeMergeFromRaw(result) : null;
        if (_opSum) return planHtmlP + renderOperationResultHtml(data, _opSum, result);
        // 【形状无关兜底】结构化解包漏一层（双层编码 / payload/output 等未知包装键）时，
        // 从文本的 JSON 块里再找一次合并结果；命中同样渲染"统计条 + 数据表 + 折叠 JSON"，
        // 绝不把整坨 payload 摊进"返回结果"（2026-10-09 用户截图）。
        try {
            const _tu = App.ChatV2?.toolRegistryUtil;
            const _raw = (typeof result === 'string') ? result : JSON.stringify(result || '');
            const _hit = (_tu && typeof _tu.findMergePayloadInText === 'function')
                ? _tu.findMergePayloadInText(_raw) : null;
            if (_hit) {
                const _s = App.ChatV2?.renderers?.toolCard?.summarizeMergeResult?.(_hit.obj);
                if (_s) return planHtmlP + renderOperationResultHtml(
                    _hit.obj, Object.assign({}, _s, { source: _hit.obj }), result);
            }
        } catch (e) { /* 兜底失败不影响原路径 */ }

        // ── AI 显式渲染提示（view.type，封闭枚举，校验不过则忽略）──
        // 让 LLM/技能在工具结果里声明"用哪种渲染器"，覆盖下方基于内容的嗅探，
        // 但只能选 toolRegistry.RENDERER_WHITELIST 里的类型（由 getViewType 校验+归一）。
        const _reg = App.ChatV2?.toolRegistry;
        const _vt = _reg && typeof _reg.getViewType === 'function'
            ? (_reg.getViewTypeFromArgs(viewArgs) || _reg.getViewType(result))
            : null;
        if (_vt) {
            if (_vt === 'table') {
                // 行矩阵优先：excel_read_range 的 {headers,result} 与二维数组都不在 tryConvertToTable
                // 的识别范围内（它只认统计/聚合/对象数组，还会把"数组的数组"误当对象数组渲染）。
                let _p = data;
                if (typeof data === 'string') { try { _p = JSON.parse(data); } catch (e) { _p = null; } }
                let _t = tryRowsMatrixTable(_p);
                if (!_t) _t = tryConvertToTable(data, typeof data === 'string' ? data : JSON.stringify(data));
                if (_t) return planHtmlP + _t;            // 非表格则回落到默认嗅探
            } else if (_vt === 'json') {
                const _p = JSON.stringify(data, null, 2);
                const _tr = _p.length > 2000 ? _p.slice(0, 2000) + '\n...(结果已截断，共 ' + _p.length + ' 字符)' : _p;
                // 用户显式点名"用 json 展示"→ 默认展开 JSON，避免结果藏在折叠面板里看不见
                return planHtmlP + renderCollapsibleJson(_tr, !!(viewArgs && viewArgs.__userView));
            } else if (_vt === 'code') {
                const _c = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
                const _hl = highlightSyntax(_c, 'json');
                return planHtmlP + '<pre class="code-block"><code>' + (_hl || escapeHtml(_c)) + '</code></pre>';
            } else if (_vt === 'markdown') {
                const _m = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
                return planHtmlP + renderTextPreservingCharts(_m);
            } else if (_vt === 'text') {
                const _x = typeof data === 'string' ? data : JSON.stringify(data, null, 2);
                return planHtmlP + '<div style="white-space:pre-wrap;">' + escapeHtml(_x) + '</div>';
            } else if (_vt === 'steps') {
                const _s = extractStepsPlan(data);
                if (_s && _s.length) return renderStepsPlanHtml(_s);
            }
            // 声明了渲染器却无法处理该数据 → 静默回落到默认嗅探，不报错
        }

        if (typeof data === 'string') {
            return planHtmlP + `<pre>${escapeHtml(data)}</pre>`;
        }
        if (data && typeof data === 'object') {
            const pretty = JSON.stringify(data, null, 2);
            const truncated = pretty.length > 2000
                ? pretty.slice(0, 2000) + '\n...(结果已截断，共 ' + pretty.length + ' 字符)'
                : pretty;
            return renderCollapsibleJson(truncated);
        }
        return planHtmlP + '<pre>' + escapeHtml(String(data)) + '</pre>';
    }

    function renderJsonHtml(jsonText) {
        return App.ChatV2.renderers.json.renderJsonHtml(jsonText);
    }

    /**
     * 把 JSON 文本渲染为可折叠面板（默认收起），
     * 消除"步骤规划"绿面板与底部 JSON 的重复展示，同时保留原始数据可展开查看。
     * 使用原生 <details>/<summary>，无需额外 JS 事件绑定。
     */
    function renderCollapsibleJson(jsonText, open) {
        return App.ChatV2.renderers.json.renderCollapsibleJson(jsonText, open);
    }

    // ============ 可折叠面板渲染（PY_COLLAPSIBLE）============

    /**
     * 将文本中的 <!-- PY_COLLAPSIBLE --> 标记内容与普通内容分离，
     * 仅对普通内容执行 Markdown 渲染，可折叠内容保持原样。
     * 处理流式渲染场景：未闭合的标记会显示加载中状态。
     * 返回 { html: 渲染后的HTML, panels: 可折叠面板数据 }
     */
    function processCollapsibleContent(text) {
        if (!text) return { html: '', panels: [] };

        var panels = [];
        var segments = [];  // { type: 'text'|'panel'|'loading', content, panelData }
        var panelIndex = 0;

        // 完整标记对的正则
        var completePattern = /<!--\s*PY_COLLAPSIBLE\s+type="(\w+)"\s+title="([^"]*)"\s*-->([\s\S]*?)<!--\s*PY_COLLAPSIBLE_END\s*-->/g;
        // 仅有开始标记（流式中间状态）
        var openStartPattern = /<!--\s*PY_COLLAPSIBLE\s+type="(\w+)"\s+title="([^"]*)"\s*-->/g;

        var lastIndex = 0;
        var match;

        // 第一遍：匹配完整的标记对
        var tempSegments = [];
        var tempLastIndex = 0;
        while ((match = completePattern.exec(text)) !== null) {
            if (match.index > tempLastIndex) {
                tempSegments.push({ type: 'raw', content: text.substring(tempLastIndex, match.index) });
            }
            var panelType = match[1];
            var panelTitle = match[2];
            var panelContent = match[3].trim();
            var panelId = 'PY_COLLAPSIBLE_PANEL_' + panelIndex;
            panels.push({
                id: panelId,
                type: panelType,
                title: panelTitle,
                content: renderMarkdown(panelContent),
                complete: true
            });
            tempSegments.push({ type: 'panel', panelId: panelId });
            tempLastIndex = match.index + match[0].length;
            panelIndex++;
        }
        if (tempLastIndex < text.length) {
            tempSegments.push({ type: 'raw', content: text.substring(tempLastIndex) });
        }

        // 第二遍：在剩余的 raw 片段中检测所有未闭合的开始标记
        for (var si = 0; si < tempSegments.length; si++) {
            var seg = tempSegments[si];
            if (seg.type !== 'raw') {
                segments.push(seg);
                continue;
            }
            var rawText = seg.content;
            // 遍历查找所有未闭合的开始标记
            var lastScan = 0;
            openStartPattern.lastIndex = 0;
            var openMatch;
            while ((openMatch = openStartPattern.exec(rawText)) !== null) {
                // 标记前的文本（正常渲染）
                if (openMatch.index > lastScan) {
                    segments.push({ type: 'text', content: rawText.substring(lastScan, openMatch.index) });
                }
                // 未闭合的面板，显示加载中状态
                // 标记后的所有内容都归入面板（包括部分到达的内容），不暴露原始标记
                var openType = openMatch[1];
                var openTitle = openMatch[2];
                var openPanelId = 'PY_COLLAPSIBLE_PANEL_LOADING_' + panelIndex;
                // 提取面板内部已到达的部分内容
                var partialContent = rawText.substring(openMatch.index + openMatch[0].length);
                // 清理：移除任何可能的 PY_COLLAPSIBLE_END 标记（如果已到达但还没被 completePattern 捕获）
                partialContent = partialContent.replace(/<!--\s*PY_COLLAPSIBLE_END\s*-->/g, '').trim();
                panels.push({
                    id: openPanelId,
                    type: openType,
                    title: openTitle,
                    content: renderMarkdown(partialContent),
                    complete: false
                });
                segments.push({ type: 'panel', panelId: openPanelId });
                panelIndex++;
                // 不再处理标记后的内容（已被面板吸收）
                lastScan = rawText.length;
                break;  // 只处理第一个未闭合标记，后续都归入这个面板
            }
            // 如果没有找到未闭合标记，添加整个 raw 文本
            if (lastScan === 0) {
                segments.push({ type: 'text', content: rawText });
            }
        }

        // 组装 HTML：普通文本执行 Markdown 渲染，面板用占位符标记，skill_card JSON 结构化渲染
        var htmlParts = [];
        for (var i = 0; i < segments.length; i++) {
            var s = segments[i];
            if (s.type === 'text') {
                // 先尝试从文本中提取 skill_card JSON（可能混合了前后说明文字）
                var cardInfo = findSkillCardJson(s.content);
                if (cardInfo) {
                    var before = s.content.substring(0, cardInfo.start);
                    var after = s.content.substring(cardInfo.end);
                    if (before.trim()) htmlParts.push(renderMarkdown(before));
                    htmlParts.push(renderSkillCardFromJson(cardInfo.obj));
                    if (after.trim()) htmlParts.push(renderMarkdown(after));
                } else if (isHtmlContent(s.content)) {
                    htmlParts.push(s.content);
                } else {
                    // 保护内嵌 ECharts 节点不被 markdown 破坏，其余照常渲染
                    htmlParts.push(renderTextPreservingCharts(s.content));
                }
            } else {
                htmlParts.push('<!--' + s.panelId + '-->');
            }
        }
        return { html: htmlParts.join(''), panels: panels };
    }

    /**
     * 检测文本是否包含 Skill 卡片结构化 JSON 数据
     * 支持纯 JSON 或以 '{' 开头的混合文本，返回 { match, start, end } 便于替换
     */
    function findSkillCardJson(text) {
        if (!text || typeof text !== 'string') return null;
        const trimmed = text.trim();
        // 不再要求整段以 '{' 开头：主消息区常把卡片 JSON 与状态/说明文字拼在一起
        // （"正在执行…"/进度行等），原先的 startsWith 判据直接 return null，
        // 卡片退化为 JSON 原文——用户看到的是一坨 JSON 而不是可审查/可保存的卡片。
        // 先整段解析，失败再走下面的子串扫描兜底。
        try {
            const obj = JSON.parse(trimmed);
            if (obj && obj.type === 'skill_card') {
                return { match: trimmed, start: text.indexOf('{'), end: text.indexOf('{') + trimmed.length, obj: obj };
            }
        } catch (e) {
            // 尝试定位文本中的 skill_card JSON 块
            try {
                const idx = text.indexOf('"type"');
                if (idx < 0) return null;
                // 向前找最近的 '{'
                let start = idx;
                while (start >= 0 && text[start] !== '{') start--;
                if (start < 0) return null;
                // 从 start 开始尝试解析，跳过可能的多余尾部
                let depth = 0;
                let inStr = false;
                let escape = false;
                for (let i = start; i < text.length; i++) {
                    const c = text[i];
                    if (escape) { escape = false; continue; }
                    if (c === '\\') { escape = true; continue; }
                    if (c === '"') { inStr = !inStr; continue; }
                    if (inStr) continue;
                    if (c === '{') depth++;
                    else if (c === '}') {
                        depth--;
                        if (depth === 0) {
                            const candidate = text.substring(start, i + 1);
                            const obj = JSON.parse(candidate);
                            if (obj && obj.type === 'skill_card') {
                                return { match: candidate, start: start, end: i + 1, obj: obj };
                            }
                            return null;
                        }
                    }
                }
            } catch (e2) {
                // ignore
            }
        }
        return null;
    }

    function isSkillCardData(text) {
        return !!findSkillCardJson(text);
    }

    /**
     * 将后端技术处理模式映射为用户友好名称
     */
    function getFriendlyProcessingMode(mode) {
        const map = {
            'VBACode': 'VBA 代码',
            'PythonCode': 'Python 代码',
            'MultimodalLLM': '多模态 AI',
            'Hybrid': '混合模式',
            'LocalOCR': '本地 OCR',
            'AutoCAD': 'AutoCAD',
            'ExcelFormula': 'Excel 公式',
            'ShellCommand': '命令行'
        };
        return map[mode] || mode || '';
    }

    /**
     * 从 turn 的工具调用结果里提取 Skill 卡片并渲染（供"最终结果"区使用）。
     * 场景：agent-loop 直调 skill.SkillCreator 时，skill_card 只存在于工具结果
     * （turn.toolCalls[].result），主消息正文只有 LLM 的总结文本——用户看不到
     * 可审查/可保存的卡片。此处在正文末尾补齐。
     * 取最后一个可解析的 skill_card（同一 turn 可能生成多次）。
     */
    function buildSkillCardFromTurnTools(turn) {
        try {
            if (!turn || !Array.isArray(turn.toolCalls)) return '';
            const renderers = App.ChatV2 && App.ChatV2.renderers;
            const extractor = renderers && renderers.toolCard && renderers.toolCard.extractSkillCardFromResult;
            if (typeof extractor !== 'function') return '';

            for (let i = turn.toolCalls.length - 1; i >= 0; i--) {
                const tc = turn.toolCalls[i];
                if (!tc || tc.result == null) continue;
                const card = extractor(tc.result);
                if (!card || card.type !== 'skill_card') continue;

                const full = Object.assign({}, card);
                // 保存按钮的数据源：generatedSkillJson（SkillManager 路径）优先；
                // 若工具结果被历史持久化转成了字符串，unwrap 由 extractor 内部完成，这里只需兜底取字段。
                if (!full.skillCardJson && !full.skillJson) {
                    const outs = (tc.result && (tc.result.outputs || tc.result.Outputs)) || {};
                    const gen = outs.generatedSkillJson || outs.GeneratedSkillJson;
                    if (typeof gen === 'string' && gen.trim()) full.skillCardJson = gen;
                }
                return renderSkillCardFromJson(full);
            }
            return '';
        } catch (e) {
            return '';
        }
    }

    /**
     * 从 turn 的工具调用结果里提取「可保存脚本」技能结果并渲染成盒子（供"最终结果"区使用）。
     * 与 buildSkillCardFromTurnTools 同源：都直接读 turn.toolCalls（历史重载也可靠），
     * 差别在于：
     *   - buildSkillCardFromTurnTools 面向 outputs.content 为 {"type":"skill_card"} 的 SkillCreator；
     *   - 本函数面向 outputs.content 为**原始 Markdown 脚本**、且技能自身声明了 uiActions 的技能
     *     （如 GenerateOCRParseScript），按 uiActions 组装 复制/保存为MD/重新生成 按钮。
     * 渲染位置与 skill.SkillCreator 的卡片完全一致（正文末尾），取代底部与技能无关的通用 chips。
     */
    function buildSkillScriptCardFromTurnTools(turn) {
        try {
            if (!turn || !Array.isArray(turn.toolCalls)) return '';
            const norm = App.ChatV2 && App.ChatV2._normalizeUiActions;
            const builder = App.ChatV2 && App.ChatV2._buildSaveableScriptCard;
            const extract = App.ChatV2 && App.ChatV2._extractSaveableScript;
            if (typeof norm !== 'function' || typeof builder !== 'function' || typeof extract !== 'function') return '';

            for (let i = turn.toolCalls.length - 1; i >= 0; i--) {
                const tc = turn.toolCalls[i];
                if (!tc || tc.result == null) continue;
                const nm = String(tc.name || '');
                const m = nm.match(/^skill[._](.+)$/);
                if (!m) continue;
                const skillName = m[1];
                const skill = (App.cachedSkills || []).find(function (s) {
                    return s && (s.name === skillName || s.Name === skillName);
                });
                // 未声明 uiActions 的技能不在此渲染（避免与工具卡/正文重复展示）
                if (!norm(skill)) continue;
                const script = extract(tc.result);
                if (!script) continue;
                const prompt = (tc.args && (tc.args.input || tc.args.prompt || tc.args.text)) || (turn.userMessage || '');
                return builder(script, skill, tc.result, prompt);
            }
            return '';
        } catch (e) {
            return '';
        }
    }

    /**
     * 从结构化 JSON 渲染 Skill 卡片 HTML
     * 如果 data 中带 skillCardJson，生成带 data-skill-json 的保存按钮
     */
    function renderSkillCardFromJson(jsonText) {
        try {
            let data;
            if (typeof jsonText === 'string') {
                data = JSON.parse(jsonText);
            } else {
                data = jsonText;
            }
            if (!data || data.type !== 'skill_card') return typeof jsonText === 'string' ? jsonText : JSON.stringify(jsonText);

            const displayName = escapeHtml(data.skillDisplayName || data.skillName || '');
            const name = escapeHtml(data.skillName || '');
            const category = escapeHtml(data.skillCategory || '') || '';
            const description = escapeHtml(data.skillDescription || '') || '';
            // 兼容两条后处理路径的字段名：
            //   ① SkillManager.PostProcessGeneratedSkill → outputs.generatedSkillJson
            //      → 前端 skill-pipeline 合入为 skillCardJson；
            //   ② SkillExecutionRouter.ProcessGeneratedSkill（兜底）→ 直接把定义内嵌进 content，字段名 skillJson。
            // 只认 ① 会出现"卡片渲染出来了，但保存/全编辑按钮不显示"——看得到却存不了。
            const skillJson = data.skillCardJson || data.skillJson || '';

            // 简洁卡片
            let cardHtml = '<div class="skill-card" style="border:1px solid #e0e0e0;border-radius:6px;margin:6px 0;font-size:13px;">';
            
            // 头部
            cardHtml += '<div style="padding:8px 12px;background:#f8f9fa;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid #eee;">';
            cardHtml += '<span style="font-weight:600;">' + displayName + '</span>';
            if (category) cardHtml += '<span style="color:#999;font-size:12px;">' + category + '</span>';
            cardHtml += '</div>';
            
            // 内容
            cardHtml += '<div class="skill-view" id="skill-view-' + name + '" style="padding:10px 12px;color:#666;">';
            if (description) cardHtml += '<div style="margin-bottom:8px;">' + description + '</div>';

            // 步骤预览（可审查性）：把生成的 steps 列出来，工具名未在注册表里时标红。
            // 之前卡片只有名称/分类/描述，用户无法判断产物是否可用，只能盲点"保存"
            // ——实测编造的 excel_send_email 就是这么被保存进去的（现在保存前会拦，但卡片上也应可见）。
            try {
                const parsedSkill = skillJson ? JSON.parse(skillJson) : null;
                const steps = parsedSkill && (parsedSkill.steps || parsedSkill.Steps);
                if (Array.isArray(steps) && steps.length > 0) {
                    const reg = App.ChatV2 && App.ChatV2.toolRegistry;
                    const items = steps.map(function (s, i) {
                        const tool = (s && (s.tool || s.Tool)) || '';
                        const sName = (s && (s.name || s.Name)) || ('步骤 ' + (i + 1));
                        const isCode = String((s && (s.strategy || s.Strategy)) || 'direct') === 'code';
                        let mark = '';
                        if (!isCode && tool && reg && typeof reg.get === 'function' && !reg.get(tool)) {
                            // 只标红不够用：编造名大多是"两个真名拼接"（实测 excel_clean_range =
                            // excel_clean_data + excel_read_range），用户不知道该改成什么，
                            // 要等点保存才由 app-skill-mgmt.js 给出建议。这里提前用同一套
                            // 别名模糊匹配（toolRegistry._resolveToolAlias）把建议显示出来。
                            let hint = '';
                            try {
                                const sugg = typeof reg._resolveToolAlias === 'function' ? reg._resolveToolAlias(tool) : null;
                                if (sugg) hint = '（是否应为 ' + escapeHtml(sugg) + '？）';
                            } catch (e) { /* 别名解析失败则只标红，不影响预览渲染 */ }
                            mark = ' <span style="color:#c62828;" title="该工具未在工具注册表中，保存时会被拒绝">⚠ 工具不存在</span>' +
                                (hint ? ' <span style="color:#c62828;">' + hint + '</span>' : '');
                        }
                        return '<li>' + escapeHtml(sName) + '：' + (isCode ? 'VBA 代码' : escapeHtml(tool)) + mark + '</li>';
                    }).join('');
                    cardHtml += '<div style="margin-bottom:8px;">' +
                        '<div style="color:#999;margin-bottom:4px;">步骤预览（' + steps.length + ' 步）</div>' +
                        '<ol style="margin:0;padding-left:18px;">' + items + '</ol></div>';
                }
            } catch (e) { /* 解析失败则不渲染预览，不影响卡片主体 */ }

            // 按钮行
            cardHtml += '<div style="display:flex;gap:6px;flex-wrap:wrap;">';
            if (skillJson && typeof skillJson === 'string') {
                // 【全编辑/复制不内联 JSON】把完整 JSON 存入全局存储，onclick 只传 name，
                // 避免 VBA 代码中的引号/反斜杠/换行破坏 HTML 属性导致按钮失效
                if (!window.__skillJsonStore) window.__skillJsonStore = {};
                window.__skillJsonStore[name] = skillJson;

                const safeJson = skillJson.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
                cardHtml += '<button onclick="App.saveGeneratedSkill(this)" data-skill-json="' + safeJson + '" style="padding:4px 10px;background:#4CAF50;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">保存</button>';
                // 全编辑按钮：从全局存储读取 JSON
                cardHtml += '<button onclick="App.openJsonEditor(\'' + name + '\')" style="padding:4px 10px;background:#FF5722;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">全编辑</button>';
            }
            cardHtml += '<button onclick="App.editSkillCard(\'' + name + '\')" style="padding:4px 10px;background:#2196F3;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">简单编辑</button>';
            cardHtml += '<button onclick="App.copySkillJson(\'' + name + '\')" style="padding:4px 10px;background:#9e9e9e;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">复制</button>';
            cardHtml += '</div>';
            cardHtml += '</div>';
            
            // 简单编辑视图（默认隐藏）
            cardHtml += '<div class="skill-edit-view" id="skill-edit-' + name + '" style="display:none;padding:10px 12px;background:#fafafa;">';
            cardHtml += '<div style="display:grid;grid-template-columns:70px 1fr;gap:6px;font-size:12px;">';
            cardHtml += '<label>名称：</label><input type="text" id="edit-displayName-' + name + '" value="' + displayName.replace(/"/g, '&quot;') + '" style="padding:4px 6px;border:1px solid #ddd;border-radius:3px;">';
            if (category) cardHtml += '<label>分类：</label><input type="text" id="edit-category-' + name + '" value="' + category.replace(/"/g, '&quot;') + '" style="padding:4px 6px;border:1px solid #ddd;border-radius:3px;">';
            if (description) cardHtml += '<label>描述：</label><textarea id="edit-description-' + name + '" rows="2" style="padding:4px 6px;border:1px solid #ddd;border-radius:3px;resize:vertical;">' + description + '</textarea>';
            cardHtml += '</div>';
            cardHtml += '<div style="margin-top:8px;display:flex;gap:6px;">';
            cardHtml += '<button onclick="App.applySkillEdit(\'' + name + '\')" style="padding:4px 10px;background:#4CAF50;color:white;border:none;border-radius:4px;cursor:pointer;font-size:12px;">确定</button>';
            cardHtml += '<button onclick="App.cancelSkillEdit(\'' + name + '\')" style="padding:4px 10px;background:#ccc;color:#666;border:none;border-radius:4px;cursor:pointer;font-size:12px;">取消</button>';
            cardHtml += '</div>';
            cardHtml += '</div>';
            
            cardHtml += '</div>';
            return cardHtml;
        } catch (e) {
            console.warn('[SkillCard] 渲染失败:', e);
            return typeof jsonText === 'string' ? jsonText : '';
        }
    }

    /**
     * 检测文本是否已经是 HTML（Skill 执行结果等场景）
     * 通过检测是否包含常见 HTML 标签来判断
     */
    function isHtmlContent(text) {
        if (!text || typeof text !== 'string') return false;
        const trimmed = text.trim();
        // 如果以 < 开头，很可能是 HTML
        if (trimmed.startsWith('<')) {
            // 检查是否包含常见 HTML 标签
            const htmlTagPattern = /<\/?(div|span|button|strong|em|p|a|img|table|tr|td|th|ul|li|ol|form|input|hr|br|h[1-6]|pre|code|blockquote|details|summary|style|script)\b/i;
            return htmlTagPattern.test(trimmed);
        }
        return false;
    }

    /**
     * 将面板占位符替换为可展开/折叠的面板 HTML。
     * 支持 type: debug(调试) / result(结果) / error(错误)
     * 未完成的面板：有部分内容时显示内容+加载指示，无内容时显示加载中
     */
    function injectCollapsiblePanels(html, panels) {
        if (!html || !panels || panels.length === 0) return html;
        var result = html;
        for (var i = 0; i < panels.length; i++) {
            var panel = panels[i];
            var placeholder = '<!--' + panel.id + '-->';
            var safeType = escapeHtml(panel.type || 'debug');
            var safeTitle = escapeHtml(panel.title || '调试信息');
            
            var panelHtml;
            if (panel.complete === false) {
                // 未完成的面板
                if (panel.content && panel.content.trim()) {
                    // 有部分内容：显示内容+加载指示
                    panelHtml = '<div class="py-collapsible py-collapsible-' + safeType + ' py-collapsible-loading">' +
                        '<div class="py-collapsible-header" data-collapsible-toggle="' + safeType + '">' +
                        '<span class="py-collapsible-icon">▸</span>' +
                        '<span class="py-collapsible-title">' + safeTitle + ' <span class="py-collapsible-loading-text">生成中...</span></span>' +
                        '</div>' +
                        '<div class="py-collapsible-body">' + panel.content + '</div>' +
                        '</div>';
                } else {
                    // 无内容：仅显示加载中
                    panelHtml = '<div class="py-collapsible py-collapsible-' + safeType + ' py-collapsible-loading">' +
                        '<div class="py-collapsible-header">' +
                        '<span class="py-collapsible-icon py-collapsible-spinner">⏳</span>' +
                        '<span class="py-collapsible-title">' + safeTitle + ' <span class="py-collapsible-loading-text">加载中...</span></span>' +
                        '</div>' +
                        '</div>';
                }
            } else {
                // 已完成的面板：可展开折叠
                panelHtml = '<div class="py-collapsible py-collapsible-' + safeType + '">' +
                    '<div class="py-collapsible-header" data-collapsible-toggle="' + safeType + '">' +
                    '<span class="py-collapsible-icon">▸</span>' +
                    '<span class="py-collapsible-title">' + safeTitle + '</span>' +
                    '</div>' +
                    '<div class="py-collapsible-body" style="display:none;">' + panel.content + '</div>' +
                    '</div>';
            }
            // 使用全局正则替换
            var placeholderRegex = new RegExp(
                placeholder.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'),
                'g'
            );
            result = result.replace(placeholderRegex, panelHtml);
        }
        return result;
    }

    /**
     * 绑定可折叠面板的点击事件（事件委托到容器级别，避免重复绑定）
     */
    function bindCollapsibleEvents(container) {
        if (!container) return;
        // 使用事件委托，每个容器只需绑定一次
        if (container._collapsibleBound) return;
        container._collapsibleBound = true;

        container.addEventListener('click', function(e) {
            const toggle = e.target.closest('[data-collapsible-toggle]');
            if (!toggle) return;
            const panel = toggle.closest('.py-collapsible');
            if (!panel) return;
            const body = panel.querySelector('.py-collapsible-body');
            const icon = panel.querySelector('.py-collapsible-icon');
            if (!body) return;
            const isOpen = body.style.display !== 'none';
            body.style.display = isOpen ? 'none' : 'block';
            if (icon) icon.textContent = isOpen ? '▸' : '▾';
            panel.classList.toggle('expanded', !isOpen);
        });
    }

    // ============ JSON 结果 → 表格增强 ============

    /**
     * 扫描内容中的 JSON 代码块，尝试将结构化数据转换为 HTML 表格
     * 支持的格式：
     *   1. 描述统计: {"列名": {count, mean, median, ...}, ...} → 汇总表
     *   2. 分组聚合: {"组名": {value, count, agg}, ...} → 聚合表
     *   3. 相关性矩阵: {"col1": {"col1": 1.0, "col2": 0.5}, ...} → 矩阵表
     *   4. 异常值检测: {"列名": {total, anomaly_count, anomalies: [...]}, ...} → 异常表
     *   5. 数据分布: {"列名": {min, max, bins, distribution: [...]}, ...} → 分布表
     */
    function enhanceJsonTables(container) {
        if (!container) return;
        // 查找所有可能的 JSON 代码块位置：
        // 1. 原始 marked 输出: pre > code.language-json
        // 2. beautifyCodeBlocks 包装后: .code-block-content > pre > code.language-json
        // 3. py-collapsible 内部的: .py-collapsible-body pre > code.language-json
        const jsonCodes = container.querySelectorAll('code.language-json');
        jsonCodes.forEach(function(codeEl) {
            const text = codeEl.textContent || '';
            try {
                const data = JSON.parse(text);
                const tableHtml = tryConvertToTable(data, text);
                if (tableHtml) {
                    // 向上查找到 pre 元素，再替换为表格
                    const preEl = codeEl.closest('pre');
                    if (preEl) {
                        const wrapper = document.createElement('div');
                        wrapper.className = 'py-json-table-wrapper';
                        wrapper.innerHTML = tableHtml;
                        // 如果 pre 在 code-block-wrapper 内，替换整个 wrapper
                        const blockWrapper = preEl.closest('.code-block-wrapper');
                        if (blockWrapper) {
                            blockWrapper.parentNode.replaceChild(wrapper, blockWrapper);
                        } else {
                            preEl.parentNode.replaceChild(wrapper, preEl);
                        }
                    }
                }
            } catch (e) {
                // 不是有效 JSON，跳过
            }
        });
        // 为所有表格添加独立的插入按钮
        addTableInsertButtons(container);
    }

    function initInlineCharts(container, retryCount) {
        return App.ChatV2.renderers.inlineChart ? App.ChatV2.renderers.inlineChart.initInlineCharts(container, retryCount) : null;
    }

    /**
     * 尝试将 JSON 数据转换为表格 HTML
     * 返回 null 表示无法识别的格式，保持原样
     */
    function tryConvertToTable(data, rawText) {
        return App.ChatV2.renderers.json.tryConvertToTable(data, rawText);
    }

    /** 行矩阵（excel_read_range 等读取类结果）→ 表格；供 _view:"table" 覆盖层使用 */
    function tryRowsMatrixTable(data) {
        return App.ChatV2.renderers.json.tryRowsMatrixTable
            ? App.ChatV2.renderers.json.tryRowsMatrixTable(data) : null;
    }

    function renderDescriptiveStatsTable(data, orientation) {
        return App.ChatV2.renderers.json.renderDescriptiveStatsTable(data, orientation);
    }

    /**
     * 切换描述性统计表格的方向（行↔列）
     * 由"切换方向"按钮触发，用户可手动调整布局
     */
    function toggleStatsTableOrientation(tableEl) {
        if (!tableEl) return;
        const orient = tableEl.getAttribute('data-stats-orient') || 'row';
        const newOrient = orient === 'row' ? 'col' : 'row';

        // 从表格当前 DOM 提取数据 → 重建表格
        const statLabels = {
            '样本数': 'count', '均值': 'mean', '中位数': 'median', '标准差': 'stdev',
            '最小值': 'min', '最大值': 'max', 'Q1(25%)': 'q1', 'Q3(75%)': 'q3',
            '统计量': '__stat__'
        };
        const labelReverse = {
            count: '样本数', mean: '均值', median: '中位数', stdev: '标准差',
            min: '最小值', max: '最大值', q1: 'Q1(25%)', q3: 'Q3(75%)'
        };

        const data = {};
        const headRow = tableEl.querySelector('thead tr');
        const bodyRows = tableEl.querySelectorAll('tbody tr');
        if (!headRow || !bodyRows.length) return;

        const ths = headRow.querySelectorAll('th');
        if (orient === 'row') {
            // 当前：统计量为行，列为列名 → 转置为统计量为列
            const colNames = [];
            for (let i = 1; i < ths.length; i++) colNames.push(ths[i].textContent.trim());
            bodyRows.forEach(function(tr) {
                const tds = tr.querySelectorAll('td');
                if (!tds.length) return;
                const statLabel = tds[0].textContent.trim();
                for (let i = 1; i < tds.length; i++) {
                    const col = colNames[i - 1] || ('col' + (i - 1));
                    if (!data[col]) data[col] = {};
                    const statKey = statLabels[statLabel] || statLabel;
                    data[col][statKey] = tds[i].textContent.trim();
                }
            });
        } else {
            // 当前：统计量为列，行为列名 → 转置为统计量为行
            const statKeys = [];
            for (let i = 0; i < ths.length; i++) statKeys.push(ths[i].textContent.trim());
            bodyRows.forEach(function(tr) {
                const tds = tr.querySelectorAll('td');
                if (!tds.length) return;
                const colName = tds[0].textContent.trim();
                if (!data[colName]) data[colName] = {};
                for (let i = 1; i < tds.length; i++) {
                    const statLabel = statKeys[i - 1] || ('stat' + (i - 1));
                    const statKey = statLabels[statLabel] || statLabel;
                    data[colName][statKey] = tds[i].textContent.trim();
                }
            });
        }

        // 用新方向重新渲染并替换
        const newHtml = renderDescriptiveStatsTable(data, newOrient);
        const wrapper = document.createElement('div');
        wrapper.innerHTML = newHtml;
        const newTable = wrapper.querySelector('table');
        if (newTable && tableEl.parentNode) {
            tableEl.parentNode.replaceChild(newTable, tableEl);
        }
    }

    function renderAggregationTable(data) {
        return App.ChatV2.renderers.json.renderAggregationTable(data);
    }

    function renderCorrelationMatrix(data) {
        return App.ChatV2.renderers.json.renderCorrelationMatrix(data);
    }

    function renderAnomalyTable(data) {
        return App.ChatV2.renderers.json.renderAnomalyTable(data);
    }

    function renderDistributionTable(data) {
        return App.ChatV2.renderers.json.renderDistributionTable(data);
    }

    function renderTrendTable(data) {
        return App.ChatV2.renderers.json.renderTrendTable(data);
    }

    function renderArrayTable(data) {
        return App.ChatV2.renderers.json.renderArrayTable(data);
    }

    // ============ AssistantBlock 内容块分类器 ============

    /**
     * 将 LLM 输出内容按 text/reasoning/tool-call 分类（DSH AssistantBlock 借鉴）
     * 用于最终渲染时对混合内容进行独立处理。
     * @param {string} rawContent - LLM 原始输出文本
     * @returns {Array<{kind: string, text: string}>} 分类后的块列表
     */
    function toAssistantBlocks(rawContent) {
        if (!rawContent) return [{ kind: 'text', text: '' }];
        const blocks = [];
        let remaining = rawContent;

        // 1. 提取 <thinking>...</thinking> 标签内的推理内容
        const thinkingRegex = /<thinking>([\s\S]*?)<\/thinking>/gi;
        let lastIndex = 0;
        let match;

        while ((match = thinkingRegex.exec(rawContent)) !== null) {
            // 标签前的文本
            const before = rawContent.slice(lastIndex, match.index).trim();
            if (before) {
                blocks.push({ kind: 'text', text: before });
            }
            // 推理内容
            const thinkingText = match[1].trim();
            if (thinkingText) {
                blocks.push({ kind: 'reasoning', text: thinkingText });
            }
            lastIndex = match.index + match[0].length;
        }

        // 剩余文本
        const after = rawContent.slice(lastIndex).trim();
        if (after) {
            blocks.push({ kind: 'text', text: after });
        }

        // 如果没有找到任何 thinking 标签，整个作为 text 块
        if (blocks.length === 0 && rawContent.trim()) {
            blocks.push({ kind: 'text', text: rawContent.trim() });
        }

        return blocks;
    }

    // ============ Turn 终态 ============

    // 图片识别 turn 出口兜底：把 content 中重复输出的完全相同表格块折叠为一份（保证最终只渲染一份结果）
    function dedupeRepeatedBlocks(text) {
        const input = String(text || '');
        const blocks = input.split(/\n\s*\n/);
        const seen = new Set();
        const out = [];
        for (const b of blocks) {
            const trimmed = b.trim();
            if (!trimmed) { out.push(b); continue; }
            if (seen.has(trimmed)) continue; // 完全重复的块只保留第一次
            seen.add(trimmed);
            out.push(b);
        }
        return out.join('\n\n');
    }

    function finalizeAIBubble(handle, turn) {
        console.log('[RENDER] finalizeAIBubble called, content length=' + (turn?.content?.length || 0) +
            ', hasDataEcharts=' + (turn?.content?.includes('data-echarts') || false));
        // 移除 loading dots
        if (handle.dotsEl && handle.dotsEl.parentNode) {
            handle.dotsEl.remove();
            handle.dotsEl = null;
        }
        // 最终渲染
        if (_pendingRender.has(handle._turnId)) {
            clearTimeout(_pendingRender.get(handle._turnId));
            _pendingRender.delete(handle._turnId);
        }

        // 推理区收拢：移除流式跟随标记，标记为已稳定
        if (handle.thinkingBody) {
            delete handle.thinkingBody.dataset.followEnd;
            handle.thinkingBody.classList.remove('streaming');
            handle.thinkingBody.classList.add('settled');
            cancelAnimationFrame(handle.thinkingBody._followRaf);
            // 收拢到第一行：滚动回顶部
            handle.thinkingBody.scrollTop = 0;
            // 图片识别 turn：思考区同样去重（多轮输出会把重复思考段累积进 turn.reasoning）
            if (turn.kind === 'image' && handle.thinkingContent) {
                const dd = dedupeRepeatedBlocks(handle.thinkingContent.textContent);
                if (dd !== handle.thinkingContent.textContent) {
                    handle.thinkingContent.textContent = dd;
                }
            }
        }

        // 【纯对话不显示过程】寒暄/知识问答等无执行 turn（无工具调用、无步骤、无 Skill、
        // thoughtList 里也没有工具块）不渲染过程列表，避免"你好"也挂一串"思考过程"行。
        const _realExec = hasRealExecution(turn);
        let processHtml = _realExec ? renderThoughtListHtml(turn) : '';
        // 【三态渲染统一】thoughtList 不持久化：历史重载后为空，但 toolCalls 数据完整。
        // 此时直接复用同一套 cbx 列表作为过程列表渲染在正文顶部（与实时终态一致），
        // 不再走下方 toolsEl 的「执行过程」折叠兜底，避免同一 turn 在
        // 流式 / 终态 / 历史 呈现三种不同结构。
        if (!processHtml && _realExec) {
            const visibleCalls = Array.isArray(turn.toolCalls)
                ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
            if (visibleCalls.some(tc => tc.status === 'completed' || tc.status === 'failed')) {
                processHtml = buildCbExecListHtml(visibleCalls, turnUserText(turn));
            }
        }
        // 【终态折叠】过程列表折叠为汇总盒（历史重载同样生效）
        processHtml = wrapProcessSummaryBox(turn, processHtml);

        if (handle.contentEl) {
            if (isCodeTurn(turn)) {
                renderCodeOnly(handle.contentEl, turn.content || '', turn._codeLang, processHtml);
            } else {
                const rawContent = turn.content || '';
                // 图片识别 turn 出口兜底：模型/多轮输出累积成多份相同表格时，折叠为一份，只渲染最终结果
                const finalContent = turn.kind === 'image' ? dedupeRepeatedBlocks(rawContent) : rawContent;
                // 通用对话：（1）sanitizeFakeToolTags 清洗残留的"工具名: 结果"前缀；
                // （2）prepareConversationContent 解析【执行】标注并清洗正文。
                const sanitized = sanitizeFakeToolTags(sanitizeLeakedJsonSnippets(sanitizeToolJsonEcho(finalContent)));
                // 无真实执行步骤时用规划步骤兜底渲染绿面板（与 streaming 路径视觉一致）
                const prepared = prepareConversationContent(turn, sanitized);
                const processed = processCollapsibleContent(prepared.text);
                let renderedHtml = injectCollapsiblePanels(processed.html, processed.panels);
                // Spec 技能（清洗/美化/图表）的步骤挂到 turn._steps，
                // 气泡最终态用与合并/拆分/关联同一 renderStepsPlanHtml 绿面板置顶渲染；
                // 与流式路径（flushContentRender L228）一致，检查 _suppressSteps。
                const stepSource = (turn._steps && turn._steps.length) ? turn._steps : prepared.steps;
                // 【步骤统一到 cbx v20261003】步骤兜底：thoughtList 不持久化（历史重载后为空），
                // 此时用 turn._steps / 正文【执行】标注里的步骤补一张列表。
                // renderStepsPlanHtml 已改为输出 cbx 扁平列表（不再是旧的 exec-plan-card 绿面板），
                // 与 thoughtList 完全同一套视觉；已有 processHtml 时不重复渲染。
                if (!handle._suppressSteps && stepSource && stepSource.length && !processHtml) {
                    renderedHtml = renderStepsPlanHtml(stepSource) + renderedHtml;
                }
                if (processHtml) renderedHtml = processHtml + renderedHtml;
                // 【Skill 卡片落地】与流式路径同一兜底，但这一步才是决定性的：
                // finalizeAIBubble 会整体重设 contentEl.innerHTML，流式期补的卡片会被这次覆盖冲掉。
                // 历史重载同样复用本函数 → 卡片在历史里也能补渲染。
                try {
                    if (renderedHtml.indexOf('class="skill-card"') < 0) {
                        const _cardHtml = buildSkillCardFromTurnTools(turn);
                        if (_cardHtml) renderedHtml += _cardHtml;
                    }
                } catch (e) { /* 卡片兜底失败不影响正文渲染 */ }
                // 【Skill 脚本卡落地】与上面 skill.SkillCreator 的卡片完全同一处钩子：
                // 声明了 uiActions、产出 Markdown 脚本的技能（如 GenerateOCRParseScript）在此渲染
                // 脚本盒子 + 复制/保存/重新生成 按钮。finalize 会整体重设 innerHTML，
                // 流式期补的会被覆盖，所以必须在这里再补一次；历史重载同样走到这里。
                try {
                    if (renderedHtml.indexOf('ocr-script-card') < 0) {
                        const _scriptHtml = buildSkillScriptCardFromTurnTools(turn);
                        if (_scriptHtml) renderedHtml += _scriptHtml;
                    }
                } catch (e) { /* 脚本卡兜底失败不影响正文渲染 */ }
                handle.contentEl.innerHTML = renderedHtml;
                // 【防御】后置渲染/增强调用各自兜底：压缩构建中个别辅助函数若被改名/裁剪，
                // 单点抛错不应升级为未处理的 Promise 拒绝（表现为重复刷屏的 "e is not a function"）。
                try { beautifyCodeBlocks(handle.contentEl); } catch (_e) { ChatLog.warn?.('RENDER', '[finalize] beautifyCodeBlocks 失败', _e?.message); }
                try { bindCollapsibleEvents(handle.contentEl); } catch (_e) { ChatLog.warn?.('RENDER', '[finalize] bindCollapsibleEvents 失败', _e?.message); }
                try { enhanceJsonTables(handle.contentEl); } catch (_e) { ChatLog.warn?.('RENDER', '[finalize] enhanceJsonTables 失败', _e?.message); }
                // 【关键】最终渲染后初始化面板内嵌图表（ECharts data-echarts 节点）
                console.log('[RENDER] finalizeAIBubble: 调用 initInlineCharts, hasPlanHtml=' + !!turn._planHtml);
                try { initInlineCharts(handle.contentEl); } catch (_e) { ChatLog.warn?.('RENDER', '[finalize] initInlineCharts 失败', _e?.message); }
                // 终态兜底：命中展示诉求时强制展开过程盒与命中行（覆盖 restore 的状态回写）
                try { enforceViewHintOpen(handle.contentEl, turn); } catch (_e) { ChatLog.warn?.('RENDER', '[finalize] enforceViewHintOpen 失败', _e?.message); }
            }
        }
        // 清除状态文本（执行完毕，状态不再显示）
        if (handle.statusIndicatorEl) {
            handle.statusIndicatorEl.textContent = '';
            handle.statusIndicatorEl.style.display = 'none';
        }
        // 思考区：有推理内容时保留但默认折叠（用户可点击展开查看）；完全无推理则隐藏。
        // 纯对话（无执行）turn 连"▸ 思考过程"折叠条一起隐藏——"你好"不需要思考过程。
        const _isTextChat = !turn || turn.kind === 'text' || !turn.kind;
        const hasReasoning = handle.thinkingContent && handle.thinkingContent.textContent.trim() &&
            !(_isTextChat && !_realExec);
        if (hasReasoning) {
            handle.thinkingEl.style.display = 'block';
            handle.thinkingEl.classList.add('has-reasoning');
            // 结束态强制折叠 body 只留 toggle
            if (handle.thinkingBody) handle.thinkingBody.style.display = 'none';
            if (handle.thinkingToggle && handle.thinkingToggle.textContent.startsWith('▾')) {
                handle.thinkingToggle.textContent = '▸ 思考过程';
            }
        } else {
            handle.thinkingEl.style.display = 'none';
        }
        // 如果模型名被 LLM 实际响应覆盖过（model_set 事件），此处 turn.model 已有值
        // 如果 modelEl 仍显示初始配置值，确保它可见
        if (handle.modelEl && handle.modelEl.textContent && handle.modelEl.textContent.trim()) {
            handle.modelEl.style.display = 'inline';
        }
        // 显示 Skill 来源标识
        showSkillBadge(handle, turn);
        // 图片识别来源标识（每次识别都标注走了什么途径）
        showRecognitionSourceBadge(handle, turn);
        // 添加操作按钮（复制）
        addMessageActions(handle, turn);
        // 通用：Skill 声明 quickPrompts 字段时，输出卡片后追加快捷提示词按钮（点击直接执行）
        addSkillQuickPrompts(handle, turn);
        // 多轮继承：基于当前管道给出后续建议 pill（点击触发继承上一轮任务结构）。
        // 注：产出「可保存脚本」的技能（uiActions，如 GenerateOCRParseScript）的脚本盒子
        // 已在 contentEl 末尾由 buildSkillScriptCardFromTurnTools 渲染（与 skill_card 同一钩子）；
        // addFollowUpSuggestions 检测到 .ocr-script-card 会直接跳过，不再追加与技能无关的 chips。
        addFollowUpSuggestions(handle, turn);
        // 【多步骤管道兜底】如果 _planHtml 尚未通过工具卡插入（getOrCreateRound），
        // 在气泡最终态中确保渲染（步骤规划面板在工具区上方）。
        // 已插入时，用最新 _planHtml 更新面板 HTML（反映步骤完成后的状态变化）。
        // 【CodeBuddy 统一过程列表】text 管线 turn：过程已由 contentEl 顶部的 cbx 列表唯一承载，
        // 工具区（旧规划卡/工具卡/执行过程折叠块）整体隐藏，不再重复展示。
        // 【步骤统一到 cbx v20261003】旧的 turn._planHtml 规划卡（.exec-plan-card）已全站下线：
        // 步骤由产出方写进 thoughtList，统一由 contentEl 顶部的 cbx 列表承载。
        // 这里不再往 toolsEl 插入任何规划卡；有过程列表时工具区整体隐藏，避免同一过程出现两份展示。
        if (handle.toolsEl && processHtml) {
            handle.toolsEl.style.display = 'none';
        } else if (handle.toolsEl) {
            // ── 最终态兜底渲染 ──
            // 某些路径（如 dispatcher 快速路径 + _suppressToolCards 屏蔽）会导致
            // tool_call/tool_call_update 事件被跳过 → DOM 里没有 .exec-round，
            // 但 turn.toolCalls 数据仍然完整。从数据直接渲染 DOM，确保展开过程可见。
            const visibleCalls = (turn && Array.isArray(turn.toolCalls))
                ? turn.toolCalls.filter(tc => tc && !tc.hidden) : [];
            const _realCalls = visibleCalls.filter(tc => tc.status === 'completed' || tc.status === 'failed');
            if (_realCalls.length > 0) {
                handle.toolsEl.style.display = 'block';
                // 【结束态统一渲染】CodeBuddy 风格扁平活动列表：
                // 每行 = 图标 + 动作名 + 目标参数摘要 + 状态/耗时，点击整行展开参数/返回结果详情。
                // 整块重建，天然幂等（多次 finalize 不叠加）；同时替换掉流式期间的
                // 旧工具卡/规划卡，避免同一次执行出现两份展示。
                handle.toolsEl.innerHTML = buildCbExecListHtml(visibleCalls);
            }
            // 【结束态统一折叠执行过程】用原生 <details> 包裹 toolsEl，默认收起，
            // 点击 summary 展开能看到完整的工具调用 + thoughtList + 规划面板。
            // 使用 <details> 而非 JS toggle：浏览器原生支持，不会因 finalizeAIBubble 多次调用出问题。
            if (handle.toolsEl.style.display !== 'none') {
                // 避免重复包裹（finalizeAIBubble 可能被多次 flush 调用）
                const parent = handle.toolsEl.parentNode;
                if (parent && parent.tagName === 'DETAILS' && parent.classList.contains('exec-process-details')) {
                    // 已包裹过，只更新 summary 文字
                    const summary = parent.querySelector('summary');
                    if (summary) {
                        const roundCount = handle.toolsEl.querySelectorAll('.exec-round, .cbx-item').length;
                        const thoughtCount = handle.toolsEl.querySelectorAll('.inline-thought').length;
                        const parts = [];
                        if (roundCount > 0) parts.push(roundCount + ' 个工具调用');
                        if (thoughtCount > 0) parts.push(thoughtCount + ' 步思考');
                        const labelText = '执行过程（' + (parts.length ? parts.join(' · ') : '详情') + '）';
                        const labelEl = summary.querySelector('.epd-label');
                        if (labelEl) labelEl.textContent = labelText;
                        else summary.textContent = labelText;
                    }
                } else {
                    // 首次包裹
                    const wrapper = document.createElement('details');
                    wrapper.className = 'exec-process-details';
                    wrapper.open = false; // 默认折叠

                    const summary = document.createElement('summary');
                    summary.style.cssText = 'display:flex;align-items:center;gap:6px;cursor:pointer;font-size:12px;color:#888;padding:4px 0;border-top:1px dashed #e0e0e0;margin-top:8px;user-select:none;list-style:none;';
                    summary.innerHTML = '<span class="epd-label">执行过程</span><span class="epd-arrow" style="margin-left:auto;color:#b6bdc9;">›</span>';
                    // 自定义 marker（隐藏浏览器默认三角）
                    wrapper.style.cssText = 'list-style:none;';
                    summary.style.listStyle = 'none';
                    summary.style.markerDisplay = 'none';

                    const roundCount = handle.toolsEl.querySelectorAll('.exec-round, .cbx-item').length;
                    const thoughtCount = handle.toolsEl.querySelectorAll('.inline-thought').length;
                    const parts = [];
                    if (roundCount > 0) parts.push(roundCount + ' 个工具调用');
                    if (thoughtCount > 0) parts.push(thoughtCount + ' 步思考');
                    summary.innerHTML = '<span class="epd-label">执行过程（' + (parts.length ? parts.join(' · ') : '详情') + '）</span>' +
                        '<span class="epd-arrow" style="margin-left:auto;color:#b6bdc9;">›</span>';

                    // 右侧尖括号随 open 状态变化（⌄ / ›），与 CodeBuddy 一致
                    wrapper.addEventListener('toggle', () => {
                        const arrow = summary.querySelector('.epd-arrow');
                        if (arrow) arrow.textContent = wrapper.open ? '⌄' : '›';
                    });

                    // 移动 DOM：把 toolsEl 挪进 details
                    const next = handle.toolsEl.nextSibling;
                    wrapper.appendChild(summary);
                    wrapper.appendChild(handle.toolsEl);
                    parent.insertBefore(wrapper, next);
                }
            }
        }
        autoScroll();
    }

    // ── 参数人性化渲染 ──
    // 键值对行式布局；嵌套对象/JSON 字符串解析后缩进展示，避免一坨转义 JSON。
    function renderArgEntries(obj, depth) {
        return Object.keys(obj).map(k =>
            '<div class="cbx-arg-row">' +
                '<span class="cbx-arg-key">' + escapeHtml(k) + '</span>' +
                '<span class="cbx-arg-val">' + formatToolArgsValue(obj[k], depth) + '</span>' +
            '</div>'
        ).join('');
    }

    function formatToolArgsValue(v, depth) {
        if (v == null) return '<span class="cbx-arg-scalar">-</span>';
        if (typeof v === 'boolean') return '<span class="cbx-arg-scalar">' + (v ? '是' : '否') + '</span>';
        if (typeof v === 'number') return '<span class="cbx-arg-scalar">' + escapeHtml(String(v)) + '</span>';
        if (typeof v === 'string') {
            const s = v.trim();
            // 值本身是 JSON 字符串（如 alignment:"{\"horizontal\":...}"）→ 解析后结构化展示
            if ((s.startsWith('{') && s.endsWith('}')) || (s.startsWith('[') && s.endsWith(']'))) {
                try { return formatToolArgsValue(JSON.parse(s), depth); } catch (e) { /* 普通字符串 */ }
            }
            return '<span class="cbx-arg-scalar">' + escapeHtml(s) + '</span>';
        }
        if (Array.isArray(v)) {
            if (v.length === 0) return '<span class="cbx-arg-scalar">（空）</span>';
            const items = v.map(x =>
                (x !== null && typeof x === 'object')
                    ? '<div class="cbx-arg-obj">' + renderArgEntries(x, depth + 1) + '</div>'
                    : '<div>' + formatToolArgsValue(x, depth) + '</div>');
            return '<div class="cbx-arg-arr">' + items.join('') + '</div>';
        }
        if (typeof v === 'object') {
            const keys = Object.keys(v);
            if (keys.length === 0) return '<span class="cbx-arg-scalar">{ }</span>';
            return '<div class="cbx-arg-obj">' + renderArgEntries(v, depth + 1) + '</div>';
        }
        return '<span class="cbx-arg-scalar">' + escapeHtml(String(v)) + '</span>';
    }

    /**
     * 展示用参数。底色声明只涂表头时，从 range 析出表头那一行。
     * A1:I27 + bgScope=header → 表头范围 A1:I1。不改真正下发的参数：
     * 列类型、列宽、表身字色仍作用在整段 range 上。
     */
    function presentFormatArgs(args) {
        if (!args || typeof args !== 'object' || Array.isArray(args)) return args;
        let fmt = args.format;
        if (typeof fmt === 'string') {
            const s = fmt.trim();
            if (s.charAt(0) !== '{') return args;
            try { fmt = JSON.parse(s); } catch (e) { return args; }
        }
        if (!fmt || typeof fmt !== 'object' || fmt.bgScope !== 'header') return args;
        const range = args.range == null ? '' : String(args.range).trim();
        if (!range) return args;
        const RR = App.ChatV2 && App.ChatV2.rangeResolver;
        if (!RR || typeof RR.headerRowOf !== 'function' || typeof RR.normRangeAddr !== 'function') return args;
        const header = RR.headerRowOf(range);
        if (!header || header === RR.normRangeAddr(range)) return args;
        const out = {};
        let placed = false;
        Object.keys(args).forEach(function (k) {
            out[k] = args[k];
            if (k === 'range') {
                out['表头范围'] = header;
                placed = true;
            }
        });
        if (!placed) out['表头范围'] = header;
        return out;
    }

    function formatToolArgsHtml(args) {
        if (!args || typeof args !== 'object' || Object.keys(args).length === 0) return '';
        try {
            return '<div class="cbx-args">' + renderArgEntries(sanitizeArgsForDisplay(presentFormatArgs(args)), 0) + '</div>';
        } catch (e) {
            return '<div class="cbx-args"><pre>' + escapeHtml(JSON.stringify(args, null, 2)) + '</pre></div>';
        }
    }

    // 结束态：把一次工具调用的 摘要/耗时/参数/返回结果/错误 构建成详情 HTML
    // （内容结构与 updateToolCard 的 exec-tool-detail 一致），
    // 供步骤规划卡存在时并入对应步骤的展开区，替代底部独立工具卡列表。
    // 带 uiActions 的 Skill 工具调用：把"返回结果"从原始 JSON 改为友好脚本卡
    // （由 skill-pipeline 暴露的 _buildSaveableScriptCard 生成：可读预览 + 复制/保存/重新生成）。
    // 其余工具仍走原始 JSON 折叠展示。
    // 把技能调用的 outputFormat 桥接成渲染器 view（供 formatToolResultHtml 的覆盖层使用）。
    // 仅在技能显式声明 outputFormat 且能映射到有效 view 时返回，否则返回 null（沿用默认嗅探）。
    function resolveSkillView(tc) {
        const nm = tc && tc.name ? String(tc.name) : '';
        const m = nm.match(/^skill\.(.+)$/);
        if (!m) return null;
        const skill = (App.cachedSkills || []).find(function (s) { return s.name === m[1] || s.Name === m[1]; });
        if (!skill) return null;
        const fmt = (skill.outputFormat || skill.OutputFormat || '').toLowerCase();
        if (!fmt) return null;
        const reg = App.ChatV2?.toolRegistry;
        return (reg && typeof reg.mapOutputFormatToView === 'function') ? reg.mapOutputFormatToView(fmt) : null;
    }

    function renderSkillResultCardHtml(tc) {
        const nm = tc && tc.name ? String(tc.name) : '';
        const m = nm.match(/^skill\.(.+)$/);
        if (!m) return '';
        const skillName = m[1];
        const skill = (App.cachedSkills || []).find(function (s) { return s.name === skillName || s.Name === skillName; });
        const norm = (App.ChatV2 && typeof App.ChatV2._normalizeUiActions === 'function')
            ? App.ChatV2._normalizeUiActions(skill)
            : null;
        if (!norm) return '';
        let res = tc.result;
        if (typeof res === 'string') {
            try { res = JSON.parse(res); } catch (e) { return ''; }
        }
        if (!res || typeof res !== 'object') return '';
        const outs = res.outputs || res.Outputs || {};
        const content = outs.content || outs.Content;
        if (typeof content !== 'string' || !content.trim()) return '';
        // 可读脚本卡（含复制/保存/重新生成）统一由正文末尾的 skill 渲染器渲染
        // （buildSkillScriptCardFromTurnTools，与 skill_card 同一钩子）。
        // 工具卡此处不再重复整份脚本，也不再把原文塞进"查看原始结果(JSON)"，仅留一行指引。
        return '<div class="exec-tool-result-note" style="font-size:12px;color:#94a3b8;padding:4px 0;">📄 已生成脚本：见下方脚本卡片，可直接复制、保存或重新生成。</div>';
    }

    function buildToolCallDetailHtml(tc, userMessage) {
        if (!tc) return '';
        const parts = [];
        const registry = App.ChatV2?.toolRegistry;
        let resultIntent = null;
        if (registry && typeof registry.renderCardIntent === 'function') {
            try {
                resultIntent = registry.renderCardIntent(tc.name, 'result', tc.args, { result: tc.result, resultSummary: tc.resultSummary });
            } catch (e) { /* 静默 */ }
        }
        let summary = '';
        if (resultIntent && resultIntent.content) {
            summary = resultIntent.content;
        } else if (tc.resultSummary) {
            summary = tc.resultSummary;
        }
        if (summary) summary = sanitizeLeakedJsonSnippets(String(summary));
        if (summary) {
            parts.push('<div class="exec-tool-result-summary"><div class="exec-detail-label">摘要</div>' + escapeHtml(summary) + '</div>');
        }
        if (tc.durationMs != null) {
            parts.push('<div style="font-size:11px;color:#94a3b8;margin:4px 0;">⏱ 耗时 ' + escapeHtml(String(tc.durationMs)) + 'ms</div>');
        }
        const args = tc.args;
        if (args && Object.keys(args).length > 0) {
            parts.push('<div class="exec-tool-args"><div class="exec-detail-label">参数</div>' + formatToolArgsHtml(args) + '</div>');
        }
        if (tc.result != null) {
            const skillCard = renderSkillResultCardHtml(tc);
            if (skillCard) {
                // 带 uiActions 的 Skill：用友好脚本卡取代原始 JSON，避免"整份内容塞在查看原始结果(JSON)里"
                parts.push('<div class="exec-tool-result-detail"><div class="exec-detail-label">返回结果</div>' + skillCard + '</div>');
            } else {
                // 把反馈明细（清空内容/清除格式等）以 JSON 键并入返回结果，返回结果保持 JSON 结构
                let resultForJson = tc.result;
                const fb = (resultIntent && Array.isArray(resultIntent.feedback)) ? resultIntent.feedback : [];
                if (fb.length) {
                    const fbObj = {};
                    for (const it of fb) {
                        const key = it && it.label ? String(it.label) : '';
                        if (!key) continue;
                        fbObj[key] = (it && it.value !== undefined) ? it.value : !!(it && it.done);
                    }
                    let baseVal = resultForJson;
                    try { baseVal = unwrapToolResult(resultForJson); } catch (_) { baseVal = resultForJson; }
                    if (baseVal && typeof baseVal === 'object' && !Array.isArray(baseVal)) {
                        resultForJson = Object.assign({}, baseVal, fbObj);
                    } else if (resultForJson && typeof resultForJson === 'object' && !Array.isArray(resultForJson)) {
                        // 解包后只剩数组/标量时保留外层对象（统计字段在外层，换掉就只剩裸矩阵 —— 见 updateToolCard 同处注释）
                        resultForJson = Object.assign({}, resultForJson, fbObj);
                    } else {
                        resultForJson = Object.assign({ '返回结果': (baseVal == null ? '' : baseVal) }, fbObj);
                    }
                }
                const _hint = resolveToolViewHint(tc, userMessage);
                const _va = _hint ? Object.assign({}, args, { _view: _hint, __userView: !!detectViewFromMessage(userMessage) }) : args;
                parts.push('<div class="exec-tool-result-detail"><div class="exec-detail-label">返回结果</div>' + formatToolResultHtml(resultForJson, _va) + '</div>');
            }
        }
        if (tc.error) {
            parts.push('<div class="exec-tool-args"><div class="exec-detail-label">错误</div><pre>' + escapeHtml(String(tc.error)) + '</pre></div>');
        }
        return parts.join('');
    }

    // CodeBuddy 风格：按工具名挑一个贴近语义的图标
    function cbxIconForTool(name) {
        const n = String(name || '').toLowerCase();
        if (/read|读取|load|fetch|import/.test(n)) return '📖';
        if (/search|搜索|find|query|查询|match|查找/.test(n)) return '🔍';
        if (/chart|图表|plot|visual/.test(n)) return '📊';
        if (/html|render|format|美化|生成|create|build/.test(n)) return '🎨';
        if (/write|写入|save|export|保存|update|更新/.test(n)) return '💾';
        if (/python|exec|run|code|执行|script/.test(n)) return '⌨️';
        if (/image|截图|screen|ocr|识别|图/.test(n)) return '🖼️';
        if (/sort|排序|filter|筛选|clear|清除/.test(n)) return '🧹';
        return '🔧';
    }

    // CodeBuddy 风格：从 args 里挑 1~2 个有意义的标量值作"目标摘要"（如 [数据源 1]、范围、文件名）
    // 数组 / JSON 数组字符串（excel_merge_sheets 的 sourceSheets 等）单独美化：
    // 直接把 `["大客户部","销售三部",…]` 原样贴出来又长又难读（2026-10-07 用户反馈），
    // 改为 `[大客户部／销售三部／市场部 等 5 项]`。
    function cbxArgsSummary(tc) {
        const a = tc && tc.args;
        if (!a || typeof a !== 'object') return '';
        const parts = [];
        for (const k of Object.keys(a)) {
            const v = a[k];
            if (v == null || typeof v === 'boolean') continue;
            if (/^_/.test(k) || k === 'view' || k === 'viewType' || k === 'returnFormat') continue;
            let arr = null;
            if (Array.isArray(v)) {
                arr = v;
            } else if (typeof v === 'string' && /^\s*\[/.test(v)) {
                try { const p = JSON.parse(v); if (Array.isArray(p)) arr = p; } catch (e) { /* 非 JSON 数组，按普通字符串处理 */ }
            }
            if (arr) {
                if (!arr.length) continue;
                const names = arr.slice(0, 3).map(x => String(x));
                parts.push('[' + names.join('／') + (arr.length > 3 ? ' 等 ' + arr.length + ' 项' : '') + ']');
                if (parts.length >= 2) break;
                continue;
            }
            if (typeof v === 'object') continue;
            const s = String(v).trim();
            if (!s) continue;
            parts.push(s.length > 30 ? s.slice(0, 30) + '…' : s);
            if (parts.length >= 2) break;
        }
        return parts.join(' · ');
    }

    // CodeBuddy 风格扁平活动列表：每行一个原生 <details>，summary 是单行动作条，
    // 展开区复用 buildToolCallDetailHtml（摘要/耗时/参数/返回结果/错误）。
    function buildCbExecListHtml(calls, userMessage) {
        let html = '<div class="cbx-list">';
        for (const tc of calls || []) {
            if (!tc) continue;
            const ok = tc.status === 'completed';
            const bad = tc.status === 'failed';
            const label = escapeHtml(tc.displayName || tc.name || 'unknown');
            const target = escapeHtml(cbxArgsSummary(tc));
            const dur = (tc.durationMs != null)
                ? '<span class="cbx-dur">' + escapeHtml(String(tc.durationMs)) + 'ms</span>' : '';
            const st = bad
                ? '<span class="cbx-status bad">✕ 失败</span>'
                : (ok ? '<span class="cbx-status ok">✓ 完成</span>' : '<span class="cbx-status">⏳ 执行中</span>');
            // 【完成即折叠】命中 _view 也不再自动展开（2026-10-06 规则）：
            // 步骤确认完成（✓/✗）后一律收起，点 summary 行可手动展开详情。
            const detail = buildToolCallDetailHtml(tc, userMessage);
            html += '<details class="cbx-item cbx-tool' + (bad ? ' cbx-failed' : '') + '">' +
                '<summary>' +
                    '<span class="cbx-icon">' + cbxIconForTool(tc.name) + '</span>' +
                    '<span class="cbx-label">' + label + '</span>' +
                    '<span class="cbx-target">' + target + '</span>' +
                    '<span class="cbx-meta">' + st + dur + '</span>' +
                    '<span class="cbx-chevron">›</span>' +
                '</summary>' +
                (detail ? '<div class="cbx-detail">' + detail + '</div>' : '') +
            '</details>';
        }
        html += '</div>';
        return html;
    }

    // 图片识别来源徽章：每次识别都标注走了什么途径（视觉模型 / OCR / ImageFieldExtractor 兜底）
    function showRecognitionSourceBadge(handle, turn) {
        if (!handle || !handle.main) return;
        const source = turn && turn.recognitionSource;
        if (!source) return;

        const existing = handle.main.querySelector('.recognition-source-badge');
        if (existing) existing.remove();

        const badge = el('div', 'skill-badge recognition-source-badge');
        const icon = el('span', 'skill-badge-icon', '⚙️');
        const label = el('span', 'skill-badge-label', '图片识别');
        badge.appendChild(icon);
        badge.appendChild(label);
        badge.appendChild(document.createTextNode(' · '));
        badge.appendChild(el('span', 'skill-pipeline-tag pipeline-mcp', source));

        handle.main.insertBefore(badge, handle.contentEl ? handle.contentEl.nextSibling : null);
    }

    function showSkillBadge(handle, turn) {
        // 渲染通道收敛：不再展示"Skill: 名称 · 执行管线"这类技能专属标识条，
        // 使 skill/code/text 统一为"主要步骤 + 最终结果"的通用多步骤渲染。
        // （图片识别的 recognition-source-badge 走 showRecognitionSourceBadge，不受影响）
        return;
    }

    function extractFirstTableFromElement(element) {
        if (!element) return null;
        // 支持直接传 table 元素或容器元素
        const table = element.tagName === 'TABLE' ? element : element.querySelector('table');
        if (!table) return null;
        const rows = [];
        table.querySelectorAll('tr').forEach(tr => {
            // 跳过"...共 N 条，仅显示前 M 条"这类占位行，避免混入导出数据
            if (tr.querySelector('td[colspan]')) return;
            const row = [];
            tr.querySelectorAll('th, td').forEach(cell => {
                row.push(cell.textContent.trim());
            });
            if (row.length > 0) rows.push(row);
        });
        return rows.length > 0 ? rows : null;
    }

    /** 【全量表格数据】优先从渲染器埋的 JSON 数据岛（script.table-full-data）取完整行列，
     *  取不到（旧渲染/非表格卡）再回落 DOM 提取。返回二维数组（首行为表头）或 null。
     *  背景：DOM 表格只渲染前 N 条（renderArrayTable DISPLAY_MAX / tryRowsMatrixTable MAX），
     *  下载/复制/插入Excel 若直接读 DOM，导出文件就是"省略的"（实测：100 条只导出 20 条）。 */
    function extractFullTableData(container) {
        try {
            const island = container && container.querySelector
                ? container.querySelector('script.table-full-data') : null;
            if (island && island.textContent) {
                const parsed = JSON.parse(island.textContent);
                if (parsed && Array.isArray(parsed.rows) && parsed.rows.length) {
                    let headers = Array.isArray(parsed.headers) && parsed.headers.length
                        ? parsed.headers.map(v => v == null ? '' : String(v))
                        : (parsed.rows[0] || []).map((_, i) => '列' + (i + 1));
                    const body = parsed.rows.map(r => (Array.isArray(r) ? r : [r])
                        .map(v => v == null ? '' : String(v)));
                    return [headers].concat(body);
                }
            }
        } catch (e) { /* 数据岛损坏 → 回落 DOM 提取 */ }
        return null;
    }

    // DOM <td> 的 textContent 一律是字符串；插入 Excel 前把纯数值恢复为 number，
    // 避免"数值变文本"（Excel 单元格左上角绿三角）。非数值（含 %/¥/日期等）保留原样。
    function toTypedValue(str) {
        if (str == null) return str;
        const s = String(str).trim();
        if (s === '') return s;
        const cleaned = s.replace(/,/g, '').trim(); // 千分位
        if (/^-?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(cleaned)) {
            const n = Number(cleaned);
            if (Number.isFinite(n)) return n;
        }
        return str;
    }

    App.ChatV2.toTypedValue = toTypedValue;

    function convertTableRowsToParsedData(rows) {
        if (!rows || rows.length === 0) return null;

        // 键值对表格检测：第一行表头为 [字段/字段名/项目/属性/名称, 内容/值/数值/结果]
        if (rows.length >= 2 && rows[0].length === 2) {
            const firstHeader = rows[0][0].trim();
            const secondHeader = rows[0][1].trim();
            const keyHeaders = ['字段', '字段名', '项目', '属性', '名称', 'key', 'field', 'item', 'property'];
            const valueHeaders = ['内容', '值', '数值', '结果', 'value', 'content', 'result'];
            const lowerFirst = firstHeader.toLowerCase();
            const lowerSecond = secondHeader.toLowerCase();
            const isKeyValue = keyHeaders.some(h => lowerFirst.includes(h.toLowerCase())) &&
                              valueHeaders.some(h => lowerSecond.includes(h.toLowerCase()));

            if (isKeyValue) {
                const record = {};
                for (let i = 1; i < rows.length; i++) {
                    if (rows[i].length >= 2 && rows[i][0]) {
                        record[rows[i][0]] = toTypedValue(rows[i][1] || '');
                    }
                }
                return [record];
            }
        }

        // 标准表格：第一行为表头
        const headers = rows[0];
        const records = [];
        for (let i = 1; i < rows.length; i++) {
            const record = {};
            rows[i].forEach((cell, idx) => {
                if (idx < headers.length && headers[idx]) {
                    record[headers[idx]] = toTypedValue(cell || '');
                }
            });
            if (Object.keys(record).length > 0) records.push(record);
        }
        return records;
    }

    async function insertTableToExcel(tableEl, rowsOverride) {
        try {
            // rowsOverride：调用方从全量数据岛提取的完整行列（DOM 可能只渲染了前 N 条）
            const rows = (Array.isArray(rowsOverride) && rowsOverride.length > 0)
                ? rowsOverride
                : extractFirstTableFromElement(tableEl);
            if (!rows || rows.length === 0) {
                App.showToast?.('⚠️ 没有可插入的表格数据', 'warning');
                return;
            }

            const parsedData = convertTableRowsToParsedData(rows);
            if (!parsedData || parsedData.length === 0) {
                App.showToast?.('⚠️ 无法解析表格数据', 'warning');
                return;
            }

            if (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge) {
                const result = await window.chrome.webview.hostObjects.excelBridge.InsertParsedData(JSON.stringify(parsedData));
                const response = JSON.parse(result);
                if (response.success) {
                    App.showToast?.('✅ ' + (response.message || '表格已智能插入Excel'), 'success');
                } else {
                    throw new Error(response.error || '插入失败');
                }
            } else {
                throw new Error('无法访问后端');
            }
        } catch (error) {
            ChatLog.error('CHAT_RENDERER', '智能插入表格到Excel失败:', error);
            App.showToast?.('❌ 智能插入失败: ' + error.message, 'error');
        }
    }

    function createTableCard(targetEl, tableRef) {
        // targetEl: 要包裹的元素（table 或 wrapper）
        // tableRef: 实际的 table 元素（用于操作）
        if (!targetEl || !tableRef) return;
        
        // 保存原位置信息
        const parentNode = targetEl.parentNode;
        const nextSibling = targetEl.nextSibling;

        // 创建卡片式包装容器
        const container = el('div', 'table-card');
        
        // 创建卡片头部
        const header = el('div', 'table-card-header');
        const title = el('div', 'table-card-title', '📊 表格');
        header.appendChild(title);
        
        // 创建头部操作按钮组
        const headerActions = el('div', 'table-card-actions');
        
        // 描述性统计表：切换方向按钮（行↔列）
        if (tableRef.classList && tableRef.classList.contains('py-stats-table')) {
            const toggleBtn = el('button', 'table-action-btn', '⇄ 转置');
            toggleBtn.title = '切换行列方向（统计量为行/统计量为列）';
            toggleBtn.addEventListener('click', () => {
                // 找到当前 table 元素（可能在 toggle 过程中被替换）
                const curTable = container.querySelector('table.py-stats-table');
                if (curTable) {
                    toggleStatsTableOrientation(curTable);
                    // 转置后重新绑定按钮（table 被替换了，但 tableRef 指向旧元素）
                    const newTable = container.querySelector('table.py-stats-table');
                    if (newTable) {
                        // 更新 copyBtn 和 insertBtn 关联的 tableRef
                        tableRef = newTable;
                    }
                }
            });
            headerActions.appendChild(toggleBtn);
        }

        // 复制表格按钮
        const copyBtn = el('button', 'table-action-btn', '📋');
        copyBtn.title = '复制表格';
        copyBtn.addEventListener('click', () => {
            try {
                const curTable = container.querySelector('table') || tableRef;
                // 优先全量数据岛（DOM 可能只渲染了前 N 条）
                const rows = extractFullTableData(container) || extractFirstTableFromElement(curTable);
                if (rows) {
                    const tsv = rows.map(r => r.join('\t')).join('\n');
                    navigator.clipboard?.writeText(tsv);
                    App.showToast?.('表格已复制到剪贴板（全量 ' + (rows.length - 1) + ' 行）', 'info');
                }
            } catch (e) {
                App.showToast?.('复制失败', 'error');
            }
        });
        headerActions.appendChild(copyBtn);

        // 下载表格按钮（导出 CSV；JSON→表格同样走这里）
        const dlBtn = el('button', 'table-action-btn', '⬇ 下载');
        dlBtn.title = '下载表格为 CSV 文件';
        dlBtn.addEventListener('click', () => {
            try {
                const curTable = container.querySelector('table') || tableRef;
                // 优先全量数据岛（DOM 可能只渲染了前 N 条，实测 100 条只导出 20 条）
                const rows = extractFullTableData(container) || extractFirstTableFromElement(curTable);
                if (!rows || rows.length === 0) {
                    App.showToast?.('没有可下载的表格数据', 'warning');
                    return;
                }
                const ts = (function () { const d = new Date(); const p = function (n) { return String(n).padStart(2, '0'); }; return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '_' + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds()); })();
                // ⚠ CSV 必须带 UTF-8 BOM：无 BOM 时 Excel 按本地 ANSI 解码，中文全乱码（实测截图）
                downloadText('表格_' + ts + '.csv', '\ufeff' + tableRowsToCsv(rows), 'text/csv;charset=utf-8');
            } catch (e) {
                App.showToast?.('下载失败', 'error');
            }
        });
        headerActions.appendChild(dlBtn);

        // 插入到宿主：仅 Excel 显示「插入到 Excel」；Word/PPT 无对应写入能力，避免误导
        const isExcelHost = !(window.SmartPilotHost) || !!window.SmartPilotHost.isExcel;
        if (isExcelHost) {
            const insertBtn = el('button', 'table-action-btn table-action-primary', '插入到 Excel');
            insertBtn.addEventListener('click', () => {
                const curTable = container.querySelector('table') || tableRef;
                // 优先全量数据岛（DOM 可能只渲染了前 N 条）
                insertTableToExcel(curTable, extractFullTableData(container));
            });
            headerActions.appendChild(insertBtn);
        }
        
        header.appendChild(headerActions);
        container.appendChild(header);
        
        // 创建卡片内容区并放入目标元素
        // 【折叠滚动】展示上限提到 200 行后，大表会把消息流撑得极长（实测截图）。
        // 用固定最大高度 + 内部滚动收纳：卡片高度可控、表头吸顶；
        // 下载/复制/插入Excel 仍读全量数据岛，不受滚动收纳影响。
        const body = el('div', 'table-card-body');
        const scrollWrap = el('div', 'table-card-scroll');
        scrollWrap.appendChild(targetEl);
        body.appendChild(scrollWrap);
        container.appendChild(body);
        
        // 将容器插入到原位置
        if (nextSibling) {
            parentNode.insertBefore(container, nextSibling);
        } else {
            parentNode.appendChild(container);
        }
    }

    function addTableInsertButtons(contentEl) {
        if (!contentEl) return;
        
        // 1. 先处理独立的 py-json-table-wrapper（JSON 转换的表格）
        const jsonWrappers = contentEl.querySelectorAll('.py-json-table-wrapper');
        jsonWrappers.forEach(wrapper => {
            if (wrapper.closest('.table-card') || wrapper.dataset.cardAdded) return;
            wrapper.dataset.cardAdded = 'true';
            
            const table = wrapper.querySelector('table');
            if (table) {
                createTableCard(wrapper, table);
            }
        });
        
        // 2. 处理普通 Markdown 表格（不在任何 wrapper 中）
        const tables = contentEl.querySelectorAll('table');
        tables.forEach(table => {
            // 跳过已在 table-card、py-json-table-wrapper 或已标记的表格
            if (table.closest('.table-card') || 
                table.closest('.py-json-table-wrapper') || 
                table.dataset.insertBtnAdded) return;
            table.dataset.insertBtnAdded = 'true';
            createTableCard(table, table);
        });
    }

    function addMessageActions(handle, turn) {
        if (handle.main.querySelector('.message-actions')) return;
        if (!turn.content) return;
        const actions = el('div', 'message-actions');

        const copyBtn = el('button', 'message-action-btn copy-btn', '📋 复制');
        copyBtn.addEventListener('click', () => {
            try {
                navigator.clipboard?.writeText(turn.content || '');
                App.showToast?.('已复制到剪贴板', 'info');
            } catch (e) {
                App.showToast?.('复制失败', 'error');
            }
        });
        actions.appendChild(copyBtn);
        handle.main.appendChild(actions);
    }

    // ============ Skill 快捷提示词（通用）============

    /**
     * 通用快捷提示词渲染：任意 Skill 在 frontmatter 中声明 quickPrompts 列表
     * （[{label, prompt}]）即自动在输出卡片后渲染按钮，点击直接执行。
     * 无需任何 Skill 专用逻辑，新增 Skill 只需声明字段。
     * @param {Object} handle - 气泡句柄（main 容器）
     * @param {Object} turn - 当前 turn（含 context._matchedSkill）
     */
    function addSkillQuickPrompts(handle, turn) {
        return; // 已取消继续操作块
        if (!handle?.main || !turn?.context?._matchedSkill) return;
        // 生成新 Skill 的气泡：skill 卡片本身已是操作入口（保存/编辑/复制），
        // 再叠加泛指"继续操作：xxx"快捷按钮无上下文意义且干扰，直接跳过。
        if (handle.main.querySelector('.skill-card')) return;
        const skill = turn.context._matchedSkill;
        // 兼容后端 PascalCase / 前端 camelCase 序列化
        const prompts = skill.quickPrompts || skill.QuickPrompts;
        if (!Array.isArray(prompts) || prompts.length === 0) return;

        // 移除所有已存在的快捷操作卡片，只保留最新的一个
        const chatEl = document.getElementById('chat-messages');
        if (chatEl) {
            chatEl.querySelectorAll('.skill-quick-prompts').forEach(n => n.remove());
        }

        const box = el('div', 'skill-quick-prompts');
        box.appendChild(el('div', 'skill-quick-prompts-title', '继续操作：'));

        const chips = el('div', 'skill-quick-prompts-chips');
        for (const p of prompts) {
            if (!p || !p.label || !p.prompt) continue;
            const label = p.label;
            const prompt = p.prompt;
            const chip = el('button', 'skill-quick-prompt-chip', escapeHtml(label));
            chip.title = prompt;
            chip.addEventListener('click', () => {
                const input = document.getElementById('message-input');
                if (!input) return;
                input.value = prompt;
                if (typeof window.sendMessage === 'function') {
                    window.sendMessage();
                }
            });
            chips.appendChild(chip);
        }
        if (!chips.children.length) return;
        box.appendChild(chips);
        handle.main.appendChild(box);
    }

    // ============ 后续建议 pill（多轮继承）============

    /**
     * 基于当前 turn 的管道给出可点击的后续建议（融合 Copilot 后续建议 pill）。
     * 点击后填充消息框并触发多轮继承（复用上一轮任务结构）。
     * @param {Object} handle - 气泡句柄（main 容器）
     * @param {Object} turn - 当前 turn（含 _execSpec / context._matchedSkill）
     */
    function addFollowUpSuggestions(handle, turn) {
        return; // 已取消继续操作块
        if (!handle?.main) return;
        // 生成新 Skill 的气泡已被 skill_card 卡片覆盖"操作入口"，不再显示通用后续建议
        if (handle.main.querySelector('.skill-card')) return;
        // 产出可保存脚本的技能（uiActions，如 GenerateOCRParseScript）已渲染脚本盒子 + 操作按钮，
        // 同样不再显示与技能无关的通用后续建议（调整图表/换成类型/做成PPT）
        if (handle.main.querySelector('.ocr-script-card')) return;
        // 已有 Skill 快捷提示词卡片时不重复（skill-quick-prompts 已覆盖）
        if (handle.main.querySelector('.skill-quick-prompts')) return;
        const suggestions = getFollowUpSuggestions(turn);
        if (!suggestions.length) return;

        // 移除所有已存在的后续建议卡片，只保留最新的一个
        const chatEl = document.getElementById('chat-messages');
        if (chatEl) {
            chatEl.querySelectorAll('.follow-up-suggestions').forEach(n => n.remove());
        }

        const box = el('div', 'follow-up-suggestions');
        box.appendChild(el('div', 'follow-up-suggestions-title', '继续操作：'));

        const chips = el('div', 'follow-up-suggestions-chips');
        for (const s of suggestions) {
            if (!s || !s.label || !s.prompt) continue;
            const chip = el('button', 'follow-up-suggestion-chip', escapeHtml(s.label));
            chip.title = s.prompt;
            chip.addEventListener('click', () => {
                const input = document.getElementById('message-input');
                if (!input) return;
                input.value = s.prompt;
                if (typeof window.sendMessage === 'function') {
                    window.sendMessage();
                }
            });
            chips.appendChild(chip);
        }
        if (!chips.children.length) return;
        box.appendChild(chips);
        handle.main.appendChild(box);
    }

    function getTurnPipeline(turn) {
        return App.ChatV2.renderers.suggestions ? App.ChatV2.renderers.suggestions.getTurnPipeline(turn) : '';
    }
    function getFollowUpSuggestions(pipeline) {
        return App.ChatV2.renderers.suggestions ? App.ChatV2.renderers.suggestions.getFollowUpSuggestions(pipeline) : [];
    }


    function showTurnError(handle, error) {
        if (handle.dotsEl && handle.dotsEl.parentNode) {
            handle.dotsEl.remove();
            handle.dotsEl = null;
        }
        // 清除状态文本
        if (handle.statusIndicatorEl) { handle.statusIndicatorEl.textContent = ''; handle.statusIndicatorEl.style.display = 'none'; }
        const errEl = el('div', 'message-error');
        errEl.innerHTML = `<div class="error-box">⚠️ ${escapeHtml(error?.message || '发生错误')}<br><code>${escapeHtml(error?.code || 'UNKNOWN')}</code></div>`;
        handle.main.appendChild(errEl);
        handle.errorEl = errEl;
        autoScroll();
    }

    function showTurnCanceled(handle, reason) {
        if (handle.dotsEl && handle.dotsEl.parentNode) {
            handle.dotsEl.remove();
            handle.dotsEl = null;
        }
        if (handle.statusIndicatorEl) { handle.statusIndicatorEl.textContent = ''; handle.statusIndicatorEl.style.display = 'none'; }
        const cancelEl = el('div', 'message-canceled');
        cancelEl.innerHTML = `🚫 已取消${reason ? `（${escapeHtml(reason)}）` : ''}`;
        handle.main.appendChild(cancelEl);
        autoScroll();
    }

    // ============ 工具函数 ============

    function formatTime(ts) {
        if (!ts) return '';
        try {
            const d = new Date(ts);
            return d.getHours().toString().padStart(2, '0') + ':' +
                   d.getMinutes().toString().padStart(2, '0');
        } catch (e) { return ''; }
    }

    function fileIcon(name) {
        const n = (name || '').toLowerCase();
        if (/\.xlsx?$/.test(n)) return '📊';
        if (/\.pdf$/.test(n)) return '📕';
        if (/\.docx?$/.test(n)) return '📝';
        if (/\.txt$/.test(n)) return '📄';
        return '📎';
    }

    // ============ SessionStore 订阅 ============

    function handleChange(change) {
        try {
            const type = change?.type;
            const sessionStore = App.ChatV2.sessionStore;
            if (!sessionStore) return;

            switch (type) {
                case 'turn_start': {
                    const turn = change.turn;
                    if (!turn) return;
                    // 首次发送消息后移除欢迎面板
                    hideWelcomePanel();
                    // 渲染用户气泡
                    renderUserBubble(turn);
                    // 创建 AI 气泡
                    const handle = createAIBubble(turn);
                    if (handle) {
                        handle._turnId = turn.id;
                        // === 渲染通道收敛（执行保留 · 渲染统一）===
                        // 技能/code 由 skill-pipeline 真实执行；过程列表是唯一的多步骤通道。
                        // 纯文本且尚未进入多步骤时不展示过程面板。可见性规则见 applyRenderChannel。
                        applyRenderChannel(handle, turn);
                        _handles.set(turn.id, handle);
                    }
                    break;
                }

                case 'delta': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) return;
                    if (change.kind === 'reasoning') {
                        // 【迭代替换模型 v20260925】reasoning 不再累积到 turn.reasoning，
                        // 由 sessionStore.appendDelta('reasoning') 内部转发到 thoughtList；
                        // 这里只剩 content 流处理，reasoning 事件走 thought_update 链路
                    } else if (change.kind === 'content') {
                        // 图片识别 turn：过程中的正文不上屏，待 turn 完成时一次性渲染最终结果
                        if (handle.finalRenderOnly) break;
                        appendContent(handle, change.delta);
                    }
                    break;
                }

                // 【迭代替换模型】thoughtList 更新 → 全量重渲（thought 块 + content 混排）
                case 'thought_update': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) break;
                    scheduleContentRender(change.turnId);
                    break;
                }

                case 'model_set': {
                    const handle = _handles.get(change.turnId);
                    ChatLog.info?.('RENDERER', `[renderer] model_set turn=${change.turnId} model=${change.model} handle=${!!handle}`);
                    if (handle) setModelName(handle, change.model);
                    else ChatLog.warn?.('RENDERER', `[renderer] no handle for turn ${change.turnId}`);
                    break;
                }

                case 'status_update': {
                    const handle = _handles.get(change.turnId);
                    if (!handle || !change.text) break;
                    // 在独立的状态指示器中显示状态文本（不侵入思考区）
                    if (handle.statusIndicatorEl) {
                        handle.statusIndicatorEl.textContent = change.text;
                        handle.statusIndicatorEl.style.display = 'block';
                    }
                    break;
                }

                case 'ppt_config': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) break;
                    renderPptConfigCard(handle, change.config);
                    break;
                }

                case 'turn_update': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) break;
                    const patch = change.patch || {};

                    // ════════════════════════════════════════════════════════════
                    // 路由确认卡片：LLM 判 text 但 regex 有 Skill 命中时弹出
                    // 两个选择按钮（keep=走 Skill / drop=走 text-pipeline），
                    // 由 orchestrator.resumeSkillConfirm(turnId, choice) 解析 Promise。
                    // patch._pendingSkillConfirm === null 时清除（用户已选 / 超时 / abort）。
                    // ════════════════════════════════════════════════════════════
                    if ('_pendingSkillConfirm' in patch) {
                        // 先移除已存在的确认卡（无论 set 还是 clear）
                        const chatEl = document.getElementById('chat-messages');
                        if (chatEl) chatEl.querySelectorAll('.skill-confirm-panel').forEach(n => n.remove());

                        const confirm = patch._pendingSkillConfirm;
                        if (confirm && typeof confirm === 'object' && handle.main) {
                            const _turnId = change.turnId;
                            const _hint = escapeHtml(confirm.skill?.hint || confirm.skill?.name || '技能');
                            const _llmReason = confirm.llmReason ? `（模型理由：${escapeHtml(String(confirm.llmReason).slice(0, 30))}）` : '';
                            const box = el('div', 'skill-confirm-panel');
                            box.appendChild(el('div', 'skill-confirm-title',
                                `🤔 系统识别到你可能想【${_hint}】${_llmReason}`));
                            box.appendChild(el('div', 'skill-confirm-desc',
                                '你希望我接下来怎么做？'));
                            const chips = el('div', 'follow-up-suggestions-chips');
                            // 选项 A：走 Skill（keep）——保守优先
                            const keepChip = el('button', 'follow-up-suggestion-chip', `✅ 就生成【${_hint}】`);
                            keepChip.title = '走 Skill 管道生成图表/看板/清洗等';
                            keepChip.addEventListener('click', () => {
                                try { App.ChatV2.orchestrator?.resumeSkillConfirm(_turnId, 'keep'); } catch(e) { ChatLog.warn?.('UI', '[chat-renderer] 确认走 Skill 失败:', e); }
                            });
                            // 选项 B：走 text（drop）——尊重 LLM 判断
                            const dropChip = el('button', 'follow-up-suggestion-chip', '💬 先给我分析/说明');
                            dropChip.title = '走通用对话管道，先给你分析文本再决定';
                            dropChip.addEventListener('click', () => {
                                try { App.ChatV2.orchestrator?.resumeSkillConfirm(_turnId, 'drop'); } catch(e) { ChatLog.warn?.('UI', '[chat-renderer] 确认走对话失败:', e); }
                            });
                            chips.appendChild(keepChip);
                            chips.appendChild(dropChip);
                            box.appendChild(chips);
                            handle.main.appendChild(box);
                        }
                        break;   // 确认卡片自己管自己，不走后续渲染通道
                    }

                    // === 渲染通道收敛（执行保留 · 渲染统一）===
                    // kind 与 _enableSteps 用同一套 applyRenderChannel，避免补丁路径把技能轮当成纯文本收起。
                    if (('kind' in patch && patch.kind && patch.kind !== handle._lastKind) || ('_enableSteps' in patch)) {
                        if ('kind' in patch && patch.kind) handle._lastKind = patch.kind;
                        const curTurn = (App.ChatV2?.sessionStore?.getTurns?.() || []).find(t => t.id === change.turnId);
                        applyRenderChannel(handle, curTurn);
                    }

                    // 【步骤统一到 cbx v20261003】_stepPatch 已废弃：步骤改由 thoughtList 承载，
                    // 渲染层收到 thought_update 就会重渲 contentEl 顶部的 cbx 列表。
                    // 这里不再往 toolsEl 插 .exec-plan-card（旧绿面板全站下线）。
                    if ('_stepPatch' in patch && patch._stepPatch) {
                        handle._multiStep = true;
                        handle._suppressSteps = false;
                        scheduleContentRender(change.turnId);
                    }

                    // 【多步骤管道实时进度】_planHtml 更新时实时重渲染步骤面板，
                    // 让进度条/步骤状态徽章/可展开的步骤结果随执行实时刷新，
                    // 而非等全部完成后一次性输出。
                    // 【步骤统一到 cbx v20261003】_planHtml 实时规划卡（.exec-plan-card）已下线：
                    // 步骤数据由产出方写进 thoughtList，这里只触发一次正文重渲即可。
                    if ('_planHtml' in patch && !patch._stepPatch) {
                        scheduleContentRender(change.turnId);
                    }

                    if (patch.content === '' && patch.reasoning === '') {
                        // 清空正文和思考区
                        if (handle.contentEl) handle.contentEl.innerHTML = '';
                        if (handle.thinkingContent) handle.thinkingContent.innerHTML = '';
                        if (handle.statusIndicatorEl) { handle.statusIndicatorEl.textContent = ''; handle.statusIndicatorEl.style.display = 'none'; }
                        if (handle.thinkingBody) handle.thinkingBody.style.display = 'none';
                        if (handle.thinkingEl) handle.thinkingEl.style.display = 'none';
                        // 仅在 turn 仍处于 streaming 状态时显示 loading dots
                        const turn = sessionStore.getCurrentTurn();
                        if (turn && turn.id === change.turnId && turn.status === 'streaming') {
                            if (handle.dotsEl) handle.dotsEl.style.display = '';
                        }
                    } else if (typeof patch.content === 'string') {
                        // content 被更新（如 code-pipeline 清理内容、工具失败清理文字）
                        const turn = sessionStore.getCurrentTurn();
                        if (turn && turn.id === change.turnId) {
                            if (handle.contentEl) {
                                if (isCodeTurn(turn)) {
                                    renderCodeOnly(handle.contentEl, turn.content || '', turn._codeLang);
                                } else {
                                    const rawContent = turn.content || '';
                                    const processed = processCollapsibleContent(rawContent);
                                    handle.contentEl.innerHTML = injectCollapsiblePanels(processed.html, processed.panels);
                                    beautifyCodeBlocks(handle.contentEl);
                                    bindCollapsibleEvents(handle.contentEl);
                                    enhanceJsonTables(handle.contentEl);
                                    // turn_update 重渲染后同样初始化内嵌图表
                                    initInlineCharts(handle.contentEl);
                                }
                            }
                        }
                    }
                    break;
                }

                case 'tool_call': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) return;
                    if (handle._suppressToolCards) break;   // 通用对话不渲染工具卡片，只显示结果
                    // 【CodeBuddy 统一过程列表】text 管线：工具调用由 cbx 列表承载，不再渲染旧工具卡
                    if (isTextProcessTurn(change.turnId)) break;
                    renderToolCard(handle, change.toolCall);
                    break;
                }

                case 'tool_call_update': {
                    const handle = _handles.get(change.turnId);
                    if (!handle) return;
                    if (handle._suppressToolCards) break;   // 通用对话不渲染工具卡片，只显示结果
                    if (isTextProcessTurn(change.turnId)) break;
                    updateToolCard(handle, change.toolCallId, change.patch, change.turnId);
                    break;
                }

                case 'turn_done': {
                    const handle = _handles.get(change.turnId);
                    if (handle) {
                        finalizeAIBubble(handle, change.turn || sessionStore.getCurrentTurn() || {});
                        const cfgCard = handle.main?.querySelector('.ppt-config-card');
                        if (cfgCard) {
                            const cfgStatus = cfgCard.querySelector('.ppt-config-status');
                            if (cfgStatus) cfgStatus.textContent = '✅ 已生成完成（可再次点击重新生成）';
                            // 恢复按钮：允许非一次性重复生成
                            const gBtn = cfgCard.querySelector('.ppt-gen-btn');
                            const dBtn = cfgCard.querySelector('.ppt-default-btn');
                            if (gBtn) { gBtn.disabled = false; gBtn.style.opacity = '1'; }
                            if (dBtn) { dBtn.disabled = false; dBtn.style.opacity = '1'; }
                        }
                    }
                    // 恢复发送按钮
                    App.ChatV2.chatEntry?.resetProcessingUI?.();
                    // 清理 PPT 防重入标记 —— 允许再次点击重新生成
                    if (App.ChatV2._pptResolvedTurns) App.ChatV2._pptResolvedTurns.delete(change.turnId);
                    // 保留句柄用于后续查看；从 _handles 移除以释放内存
                    setTimeout(() => _handles.delete(change.turnId), 60000);
                    break;
                }

                case 'turn_failed': {
                    const handle = _handles.get(change.turnId);
                    if (handle) {
                        showTurnError(handle, change.error);
                        // 恢复 PPT 配置卡按钮
                        const cfgCard = handle.main?.querySelector('.ppt-config-card');
                        if (cfgCard) {
                            const cfgStatus = cfgCard.querySelector('.ppt-config-status');
                            if (cfgStatus) cfgStatus.textContent = '❌ 生成失败（可点击重新生成）';
                            const gBtn = cfgCard.querySelector('.ppt-gen-btn');
                            const dBtn = cfgCard.querySelector('.ppt-default-btn');
                            if (gBtn) { gBtn.disabled = false; gBtn.style.opacity = '1'; }
                            if (dBtn) { dBtn.disabled = false; dBtn.style.opacity = '1'; }
                        }
                    }
                    App.ChatV2.chatEntry?.resetProcessingUI?.();
                    if (App.ChatV2._pptResolvedTurns) App.ChatV2._pptResolvedTurns.delete(change.turnId);
                    setTimeout(() => _handles.delete(change.turnId), 60000);
                    break;
                }

                case 'turn_canceled': {
                    const handle = _handles.get(change.turnId);
                    if (handle) {
                        showTurnCanceled(handle, change.reason);
                        const cfgCard = handle.main?.querySelector('.ppt-config-card');
                        if (cfgCard) {
                            const cfgStatus = cfgCard.querySelector('.ppt-config-status');
                            if (cfgStatus) cfgStatus.textContent = '🚫 已取消（可点击重新生成）';
                            // 恢复按钮：允许非一次性重复生成
                            const gBtn = cfgCard.querySelector('.ppt-gen-btn');
                            const dBtn = cfgCard.querySelector('.ppt-default-btn');
                            if (gBtn) { gBtn.disabled = false; gBtn.style.opacity = '1'; }
                            if (dBtn) { dBtn.disabled = false; dBtn.style.opacity = '1'; }
                        }
                    }
                    App.ChatV2.chatEntry?.resetProcessingUI?.();
                    if (App.ChatV2._pptResolvedTurns) App.ChatV2._pptResolvedTurns.delete(change.turnId);
                    setTimeout(() => _handles.delete(change.turnId), 60000);
                    break;
                }

                case 'turns_removed': {
                    // 编辑重发时，sessionStore 已删除 turn 记录，DOM 由 editUserMessage 同步删除
                    // 此处清理对应的 _handles + PPT 防重入 Set
                    const removedIds = change.turnIds || [];
                    for (const id of removedIds) {
                        _handles.delete(id);
                        if (App.ChatV2._pptResolvedTurns) App.ChatV2._pptResolvedTurns.delete(id);
                    }
                    break;
                }

                case 'history_cleared': {
                    const c = getMessagesContainer();
                    if (c) c.innerHTML = '';
                    _handles.clear();
                    if (App.ChatV2._pptResolvedTurns) App.ChatV2._pptResolvedTurns.clear();
                    break;
                }

                default:
                    // 忽略其他事件
                    break;
            }
        } catch (e) {
            ChatLog.error?.('RENDERER', '[chat-renderer] handle change failed', e);
        }
    }

    // ============ 初始化 ============

    let _subscribed = false;

    function init() {
        if (_subscribed) return;
        const sessionStore = App.ChatV2.sessionStore;
        if (!sessionStore) {
            ChatLog.warn?.('CHAT-V2', '[chat-renderer] sessionStore not ready, retry in 200ms');
            setTimeout(init, 200);
            return;
        }
        sessionStore.subscribe(handleChange);
        _subscribed = true;
        ChatLog.info?.('CHAT-V2', '✅ chat-renderer subscribed to sessionStore');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }

    App.ChatV2.renderer = {
        init,
        handleChange,
        forceFlush: (turnId) => flushContentRender(turnId),
        renderStepsPlanHtml,
        detectAndFormatCodeBlocks,
        stripCodeFences,
        renderJsonHtml,
        tryConvertToTable,
        unwrapToolResult,
        autoSummarizeToolResult,
        formatToolErrorHtml,
        formatToolResultHtml,
        presentFormatArgs,
        stepCoveredByToolRow,
        cbxArgsSummary
    };
    ChatLog.info?.('CHAT-V2', '✅ chat-renderer loaded');
})();