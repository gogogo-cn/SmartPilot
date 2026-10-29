/**
 * ============================================================
 * app-editor.js — Code Editor
 * ============================================================
 * 功能：代码编辑器模块，提供 VBA 代码编辑、调试、高亮与代码片段管理功能。
 *
 * 核心类/函数：
 *   - initEditor() — 初始化编辑器
 *   - setEditorValue() — 设置编辑器内容
 *   - getEditorValue() — 获取编辑器内容
 *   - formatVBA() — 格式化 VBA 代码
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    // ==================== 独立调试窗口（BroadcastChannel）====================
    const debugChannel = (typeof BroadcastChannel !== 'undefined') ? new BroadcastChannel('vba-debugger') : null;
    let debugWindow = null;

    // 打开独立调试窗口
    function openDebugWindow() {
        if (debugWindow && !debugWindow.closed) {
            debugWindow.focus();
            return;
        }

        const width = 900;
        const height = 600;
        const left = (screen.width - width) / 2;
        const top = (screen.height - height) / 2;

        debugWindow = window.open(
            'debug-window.html',
            'VBADebugger',
            `width=${width},height=${height},left=${left},top=${top},resizable=yes,scrollbars=yes,status=no,menubar=no,toolbar=no`
        );

        if (!debugWindow) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">弹窗被阻止，请允许弹窗后重试</span>', true) : App.showToast('弹窗被阻止，请允许弹窗后重试', 'error');
            return;
        }

        // 同步当前调试状态到新窗口
        setTimeout(() => {
            syncDebugStateToWindow();
        }, 500);
    }

    // 同步当前调试状态到独立窗口
    function syncDebugStateToWindow() {
        if (!debugChannel) return;

        // 发送状态
        const status = App.debugState.isDebugging
            ? (App.debugState.isPaused ? '已暂停' : '运行中')
            : '就绪';
        const lineInfo = App.debugState.isDebugging && App.debugState.lines.length > 0
            ? `行 ${App.debugState.currentLineIndex + 1}/${App.debugState.lines.length}`
            : '行 0/0';

        debugChannel.postMessage({ type: 'status', status, lineInfo });

        // 发送变量（统一使用 parsedVariables 中的当前值）
        if (App.debugState.parsedVariables && App.debugState.parsedVariables.length > 0) {
            const vars = App.debugState.parsedVariables.map(v => ({
                name: v.name,
                value: v.value || '<未初始化>',
                type: v.type || 'Variant'
            }));
            debugChannel.postMessage({ type: 'watch', variables: vars });
        }
    }

    // 接收独立窗口的命令
    if (debugChannel) {
        debugChannel.onmessage = function(event) {
            const msg = event.data;
            if (!msg || !msg.type) return;

            if (msg.type === 'command') {
                switch (msg.command) {
                    case 'stepInto': stepInto(); break;
                    case 'startOrStepInto': startOrStepInto(); break;
                    case 'stepOver': stepOver(); break;
                    case 'continue': continueExecution(); break;
                    case 'stop': stopDebug(); break;
                    case 'refreshWatch': refreshWatchWindow(); break;
                }
            } else if (msg.type === 'ready') {
                // 独立窗口已就绪，同步状态
                syncDebugStateToWindow();
            } else if (msg.type === 'ping') {
                // 回应心跳
                debugChannel.postMessage({ type: 'heartbeat' });
            }
        };
    }

    // ==================== 编辑器 UI 功能 ====================

    // --- 来自 app.js:L20581-L20602 ---
    // 生成/更新行号
    function updateLineNumbers() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;

        const lines = editor.value.split('\n');
        const count = lines.length;

        let html = '';
        for (let i = 1; i <= count; i++) {
            html += `<div class="line-number" data-line="${i}"><span class="line-number-text">${i}</span></div>`;
        }

        lineNumbers.innerHTML = html;

        // ✅ 不设置任何高度属性，让CSS控制
        // CSS已设置 height: 100% 确保与编辑器同高

        // 立即同步位置
        syncScroll();

        // 更新折叠按钮
        updateFoldButtons();
    }

    // --- 来自 app.js:L20605-L20623 ---
    // 同步滚动 - 使用比例确保行号和代码编辑器完全同步
    function syncScroll() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;

        // ✅ 使用滚动比例（最可靠的方式）
        const editorScrollHeight = editor.scrollHeight - editor.clientHeight;
        const lineNumbersScrollHeight = lineNumbers.scrollHeight - lineNumbers.clientHeight;

        if (editorScrollHeight > 0 && lineNumbersScrollHeight > 0) {
            // 计算滚动比例并应用到行号容器
            const scrollRatio = editor.scrollTop / editorScrollHeight;
            lineNumbers.scrollTop = scrollRatio * lineNumbersScrollHeight;
        } else if (editorScrollHeight <= 0) {
            // 如果编辑器不需要滚动，重置行号滚动位置
            lineNumbers.scrollTop = 0;
        }
    }

    // --- 来自 app.js:L20626-L20640 ---
    // 行号容器滚动时也同步到编辑器（双向同步）
    function syncScrollFromLineNumbers() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;

        // ✅ 反向同步：从行号到编辑器
        const lineNumbersScrollHeight = lineNumbers.scrollHeight - lineNumbers.clientHeight;
        const editorScrollHeight = editor.scrollHeight - editor.clientHeight;

        if (lineNumbersScrollHeight > 0 && editorScrollHeight > 0) {
            const scrollRatio = lineNumbers.scrollTop / lineNumbersScrollHeight;
            editor.scrollTop = scrollRatio * editorScrollHeight;
        }
    }

    // --- 来自 app.js:L20643-L20684 ---
    // 高亮当前行（编辑器模式）
    function highlightCurrentLine(lineNum = null) {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;

        // 移除所有高亮
        document.querySelectorAll('.line-number').forEach(el => {
            el.classList.remove('current-line');
        });

        // 计算当前光标所在行
        if (lineNum === null) {
            const pos = editor.selectionStart;
            const textBeforeCursor = editor.value.substring(0, pos);
            lineNum = textBeforeCursor.split('\n').length;
        }

        // 高亮指定行
        const targetLine = lineNumbers.querySelector(`[data-line="${lineNum}"]`);
        if (targetLine) {
            targetLine.classList.add('current-line');

            // 确保该行可见（平滑滚动）
            const lineTop = targetLine.offsetTop;
            const lineHeight = 19.2;
            const containerHeight = editor.clientHeight;
            const scrollTop = editor.scrollTop;

            // 计算目标行应该在的位置
            const targetScrollTop = lineTop - containerHeight / 2 + lineHeight / 2;

            // 只在需要时滚动
            if (lineTop < scrollTop || lineTop + lineHeight > scrollTop + containerHeight) {
                // 使用平滑滚动
                editor.scrollTo({
                    top: Math.max(0, targetScrollTop),
                    behavior: 'smooth'
                });
            }
        }
    }

    // --- 来自 app.js:L20687-L20706 ---
    // ✅ 滚动到代码编辑器顶部
    function scrollToTop() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (editor) {
            editor.scrollTo({
                top: 0,
                left: 0, // ✅ 同时重置水平滚动
                behavior: 'smooth'
            });
        }

        if (lineNumbers) {
            lineNumbers.scrollTo({
                top: 0,
                left: 0, // ✅ 同时重置水平滚动
                behavior: 'smooth'
            });
        }
    }

    // --- 来自 app.js:L20709-L20721 ---
    // 标记已执行/错误/当前行（调试用）
    function markLineStatus(lineNumber, status) {
        const lineEl = document.querySelector(`.line-number[data-line="${lineNumber}"]`);
        if (!lineEl) return;

        lineEl.classList.remove('current-line', 'executed-line', 'error-line');
        if (status === 'executed') {
            lineEl.classList.add('executed-line');
        } else if (status === 'error') {
            lineEl.classList.add('error-line');
        } else if (status === 'current') {
            lineEl.classList.add('current-line');
            // 确保当前行在可视区域
            lineEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
        }
    }

    // --- 来自 app.js:L20724-L20728 ---
    // 清除所有行标记
    function clearAllLineMarks() {
        document.querySelectorAll('.line-number').forEach(el => {
            el.classList.remove('current-line', 'executed-line', 'error-line');
        });
    }

    // --- 来自 app.js:L20730-L20737 --- 已扩展重置语言 ---
    function clearEditor() {
        const editor = document.getElementById('code-editor');
        const highlightContent = document.getElementById('code-highlight-content');
        if (editor) {
            editor.value = '';
            updateLineNumbers();
            clearAllLineMarks();
        }
        if (highlightContent) {
            highlightContent.innerHTML = '';
        }
        // 清空折叠状态
        foldedBlocks.clear();
        // 重置为默认语言
        switchEditorLanguage('vba');
    }

    // 格式化编辑器中的代码
    function formatEditorCode() {
        const editor = document.getElementById('code-editor');
        if (!editor || !editor.value) return;

        try {
            const lang = editorLanguage || 'vba';
            const formatted = autoFormatCode(editor.value, lang);
            if (formatted !== editor.value) {
                const scrollTop = editor.scrollTop;
                editor.value = formatted;
                editor.scrollTop = scrollTop;
                updateLineNumbers();
                updateSyntaxHighlight();
                updateFoldButtons();
            }
        } catch(e) {
            console.error('格式化失败:', e);
        }
    }

    // --- 来自 app.js:L20740-L20793 --- 已扩展显示语言切换器 ---
    // 显示代码编辑器（展开左侧面板）
    async function showCodeEditor(code, language) {
        const leftPanel = document.getElementById('left-panel');
        const codeContainer = document.getElementById('code-with-lines');
        const actions = document.getElementById('editor-actions');
        const title = document.getElementById('editor-title');
        const langSelect = document.getElementById('editor-language');

        if (!leftPanel || !codeContainer) {
            ChatLog.error('EDITOR', 'showCodeEditor: 元素未找到', { leftPanel: !!leftPanel, codeContainer: !!codeContainer });
            return;
        }

        // 设置语言
        if (language) {
            switchEditorLanguage(language);
        }

        const isAlreadyExpanded = leftPanel.classList.contains('expanded');

        if (!isAlreadyExpanded && App.excelBridge) {
            try {
                const currentWidth = await App.excelBridge.GetSidebarWidth();
                App.sidebarWidthBeforeEditor = currentWidth;
                ChatLog.info('EDITOR', '当前宽度:', currentWidth);

                const newWidth = currentWidth + 400;
                ChatLog.info('EDITOR', '设置新宽度:', newWidth);
                await App.excelBridge.SetSidebarWidth(newWidth);

                await new Promise(resolve => setTimeout(resolve, 100));
            } catch (e) {
                ChatLog.info('EDITOR', '调整宽度失败:', e);
            }
        }

        if (!isAlreadyExpanded) {
            leftPanel.classList.add('expanded');
        }
        // 移除 is-hidden，否则 !important 会覆盖内联 display
        codeContainer.classList.remove('is-hidden');
        codeContainer.style.display = 'flex';
        if (actions) {
            actions.classList.remove('is-hidden');
            actions.style.display = 'flex';
        }
        if (title) {
            title.classList.remove('is-hidden');
            title.style.display = 'inline-flex';
        }
        if (langSelect) {
            langSelect.classList.remove('is-hidden');
            langSelect.style.display = '';
        }

        // ✅ 先初始化行号和所有编辑器功能
        initLineNumbers();

        // ✅ 然后设置代码并触发更新
        if (code) {
            const editor = document.getElementById('code-editor');
            if (editor) {
                // 自动格式化代码缩进
                const formattedCode = autoFormatCode(code, editorLanguage || 'vba');
                editor.value = formattedCode;
                // 延迟触发更新，确保 DOM 完全就绪
                setTimeout(() => {
                    updateLineNumbers();
                    updateSyntaxHighlight();
                    updateFoldButtons();
                    highlightCurrentLine();
                }, 80);
            }
        }

        if (App.excelBridge && App.sidebarWidthBeforeEditor !== null) {
            try {
                const expectedWidth = App.sidebarWidthBeforeEditor + 400;
                const actualWidth = await App.excelBridge.GetSidebarWidth();
                ChatLog.info('EDITOR', '实际宽度:', actualWidth, '期望宽度:', expectedWidth);
                if (actualWidth < expectedWidth) {
                    await App.excelBridge.SetSidebarWidth(expectedWidth);
                }
            } catch (e) {
                ChatLog.info('EDITOR', '验证宽度失败:', e);
            }
        }
    }

    // --- 来自 app.js:L20442-L20464 ---
    // 初始化行号显示
    function initLineNumbers() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;
        if (editor._lineNumbersInitialized) return; // 防止重复绑定
        editor._lineNumbersInitialized = true;

        // 如果编辑器已有内容，先自动格式化
        if (editor.value && editor.value.trim().length > 0) {
            try {
                const formatted = autoFormatCode(editor.value, editorLanguage || 'vba');
                if (formatted !== editor.value) {
                    editor.value = formatted;
                }
            } catch(e) {
                // 格式化失败，保持原样
            }
        }

        // 立即生成行号
        updateLineNumbers();

        editor.addEventListener('input', () => {
            // 编辑时自动展开所有折叠块（防止行号偏移）
            if (foldedBlocks.size > 0) {
                unfoldAll();
            }
            updateLineNumbers();
            highlightCurrentLine();
            updateSyntaxHighlight();
            // 重建折叠按钮，确保手动输入/粘贴代码时折叠功能实时可用
            updateFoldButtons();
        });

        // ✅ 双向滚动同步
        editor.addEventListener('scroll', syncScroll);
        lineNumbers.addEventListener('scroll', syncScrollFromLineNumbers);

        editor.addEventListener('keyup', highlightCurrentLine);
        editor.addEventListener('click', highlightCurrentLine);

        // ✅ 初始化语法高亮
        initSyntaxHighlight();
        // ✅ 初始化代码缩进
        initCodeIndent();
        // ✅ 初始化代码折叠
        initCodeFolding();

        // 初始化后立即更新一次折叠按钮和高亮
        updateFoldButtons();
        updateSyntaxHighlight();
    }

    // ==================== 语法高亮引擎 ====================

    let highlightThrottleTimer = null;

    function escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    function highlightSyntax(code, lang) {
        if (!code) return '';

        switch (lang) {
            case 'vba': case 'vb': return highlightVBA(code);
            case 'python': case 'py': return highlightPython(code);
            case 'javascript': case 'js': case 'typescript': case 'ts': return highlightJS(code);
            default: return escapeHtml(code);
        }
    }

    function protectAndHighlight(code, patterns) {
        const protectedRanges = [];

        patterns.protect.forEach(function(p) {
            let m;
            const regex = new RegExp(p.regex.source, p.regex.flags.replace('g', 'g'));
            while ((m = regex.exec(code)) !== null) {
                protectedRanges.push({
                    start: m.index,
                    end: m.index + m[0].length,
                    text: m[0],
                    type: p.type || 'string'
                });
                if (m[0].length === 0) regex.lastIndex++;
            }
        });

        protectedRanges.sort(function(a, b) { return a.start - b.start || (b.end - b.start) - (a.end - a.start); });

        const filteredRanges = [];
        for (let i = 0; i < protectedRanges.length; i++) {
            const r = protectedRanges[i];
            let covered = false;
            for (let j = 0; j < filteredRanges.length; j++) {
                const existing = filteredRanges[j];
                if (r.start >= existing.start && r.end <= existing.end) {
                    covered = true;
                    break;
                }
                if (r.start < existing.end && r.end > existing.start) {
                    if ((r.end - r.start) > (existing.end - existing.start)) {
                        filteredRanges.splice(j, 1);
                        j--;
                    } else {
                        covered = true;
                        break;
                    }
                }
            }
            if (!covered) {
                filteredRanges.push(r);
            }
        }

        let output = '';
        let pos = 0;
        const protectClassMap = {
            'string': 'tok-string',
            'comment': 'tok-comment',
            'tag': 'tok-tag',
            'attr': 'tok-attr',
            'key': 'tok-key'
        };

        filteredRanges.forEach(function(range) {
            if (range.start > pos) {
                const segment = code.slice(pos, range.start);
                output += applyHighlight(escapeHtml(segment), patterns.highlight);
            }
            const escaped = escapeHtml(range.text);
            const cls = protectClassMap[range.type] || 'tok-string';
            output += '<span class="' + cls + '">' + escaped + '</span>';
            pos = range.end;
        });

        if (pos < code.length) {
            output += applyHighlight(escapeHtml(code.slice(pos)), patterns.highlight);
        }

        return output;
    }

    function applyHighlight(html, highlightRules) {
        let result = html;
        highlightRules.forEach(function(h) {
            const placeholders = [];
            result = result.replace(/<span[^>]*>[\s\S]*?<\/span>/g, function(match) {
                const ph = '\u0001' + placeholders.length + '\u0001';
                placeholders.push(match);
                return ph;
            });
            result = result.replace(h.regex, h.replacer);
            placeholders.forEach(function(ph, i) {
                result = result.replace('\u0001' + i + '\u0001', ph);
            });
        });
        return result;
    }

    function highlightVBA(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /'.*$/gm, type: 'comment' },
                { regex: /Rem\s.*$/gmi, type: 'comment' },
                { regex: /"([^"]*)"/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+(\.\d+)?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(End\s+(If|Sub|Function|With|Type|Enum|Select|Property|Get|Let|Set|Do|For|While))\b/gi, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b(Sub|Function|Dim|As|If|Then|Else|ElseIf|For|Next|To|While|Wend|Do|Loop|Until|Each|In|Step|Exit|Public|Private|Const|Type|ReDim|Set|Let|Get|New|With|On|Error|Resume|GoTo|Call|ByVal|ByRef|Optional|ParamArray|Variant|Integer|Long|Double|Single|String|Boolean|Date|Object|Nothing|True|False|Me|Static|Preserve|Erase)\b/gi, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b(MsgBox|InputBox|Debug|Print|Cells|Rows|Columns|Sheets|Worksheets|ActiveSheet|ActiveWorkbook|ThisWorkbook|Selection|Range|Offset|Resize|Value|Formula|Text|Interior|Font|Borders|Merge|Unmerge|Copy|Paste|Delete|Insert|Clear|Sort|Filter|AutoFilter|Find|Replace|Like|Goto|Worksheet|Workbook|Application|ScreenUpdating|EnableEvents|DisplayAlerts|Calculation|Count)\b/g, replacer: '<span class="tok-builtin">$1</span>' }
            ]
        });
    }

    function highlightPython(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /"""[\s\S]*?"""/g, type: 'string' },
                { regex: /'''[\s\S]*?'''/g, type: 'string' },
                { regex: /#[^\n]*/g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /'([^'\\]|\\.)*'/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+(\.\d+)?j?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(def|class|import|from|as|return|if|elif|else|for|while|break|continue|pass|try|except|finally|with|lambda|yield|global|nonlocal|assert|del|in|not|and|or|is|None|True|False|self|super|raise|async|await|yield|lambda)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b([a-zA-Z_][\w]*)\s*\(/g, replacer: '<span class="tok-function">$1</span>(' }
            ]
        });
    }

    function highlightJS(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /\/\/[^\n]*/g, type: 'comment' },
                { regex: /\/\*[\s\S]*?\*\//g, type: 'comment' },
                { regex: /`([^`\\]|\\.)*`/g, type: 'string' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /'([^'\\]|\\.)*'/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+(\.\d+)?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(function|const|let|var|return|if|else|for|while|switch|case|break|continue|new|class|extends|super|import|export|default|try|catch|finally|throw|typeof|instanceof|in|of|await|async|yield|delete|void|this|null|undefined|true|false|static|get|set|extends|implements|interface|package|private|protected|public)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b([a-zA-Z_$][\w$]*)\s*\(/g, replacer: '<span class="tok-function">$1</span>(' }
            ]
        });
    }

    // 初始化语法高亮
    function initSyntaxHighlight() {
        const editor = document.getElementById('code-editor');
        const overlay = document.getElementById('code-highlight-overlay');
        const highlightContent = document.getElementById('code-highlight-content');

        if (!editor || !overlay || !highlightContent) return;

        // 同步滚动
        editor.addEventListener('scroll', syncHighlightScroll);

        // 立即更新一次高亮
        updateSyntaxHighlight();
    }

    // 更新语法高亮显示（行级渲染，支持折叠）
    function updateSyntaxHighlight() {
        const editor = document.getElementById('code-editor');
        const highlightContent = document.getElementById('code-highlight-content');

        if (!editor || !highlightContent) return;

        const code = editor.value;
        const lang = editorLanguage || 'vba';
        const lines = code.split('\n');

        try {
            let html = '';
            for (let i = 0; i < lines.length; i++) {
                const lineHighlighted = highlightSyntax(lines[i], lang);
                html += `<div class="code-line" data-line="${i}" style="display:block;">${lineHighlighted || '&nbsp;'}</div>`;
            }
            highlightContent.innerHTML = html;

            // 应用折叠状态
            applyFoldToHighlight();
        } catch (e) {
            // 出错时显示纯文本
            highlightContent.textContent = code;
        }

        syncHighlightScroll();
    }

    // 将折叠状态应用到高亮层和行号
    function applyFoldToHighlight() {
        const highlightContent = document.getElementById('code-highlight-content');
        const lineNumbers = document.getElementById('line-numbers');

        if (!highlightContent) return;

        const codeLines = highlightContent.querySelectorAll('.code-line');
        const lineNumberEls = lineNumbers ? lineNumbers.querySelectorAll('.line-number') : [];

        // 先全部显示
        codeLines.forEach(el => {
            el.style.display = '';
            // 移除之前插入的折叠占位符（避免折叠→展开后残留提示词、或多次折叠堆叠出多个提示）
            const placeholders = el.querySelectorAll(':scope > .fold-placeholder');
            placeholders.forEach(p => p.remove());
        });
        lineNumberEls.forEach(el => {
            el.style.display = '';
        });

        // 处理折叠块
        if (foldedBlocks.size === 0) return;

        // 按起始行从大到小排序（从后往前处理，避免影响前面的行号）
        const sortedStarts = Array.from(foldedBlocks.keys()).sort((a, b) => b - a);

        sortedStarts.forEach(startLine => {
            const folded = foldedBlocks.get(startLine);
            if (folded.visible) return;

            const endLine = folded.endLine;
            const foldCount = endLine - startLine;

            // 隐藏被折叠的代码行（startLine+1 到 endLine）
            for (let i = startLine + 1; i <= endLine; i++) {
                if (codeLines[i]) {
                    codeLines[i].style.display = 'none';
                }
            }

            // 隐藏被折叠的行号行（startLine+1 到 endLine）
            for (let i = startLine + 1; i <= endLine; i++) {
                if (lineNumberEls[i]) {
                    lineNumberEls[i].style.display = 'none';
                }
            }

            // 在起始行后插入折叠占位符（修改起始行显示）
            if (codeLines[startLine]) {
                // 在行尾添加折叠标记
                const placeholder = document.createElement('span');
                placeholder.className = 'tok-comment fold-placeholder';
                placeholder.textContent = '  ⋯ (' + foldCount + ' 行已折叠)';
                codeLines[startLine].appendChild(placeholder);
            }
        });
    }

    // 同步高亮层滚动位置
    function syncHighlightScroll() {
        const editor = document.getElementById('code-editor');
        const overlay = document.getElementById('code-highlight-overlay');

        if (!editor || !overlay) return;

        overlay.scrollTop = editor.scrollTop;
        overlay.scrollLeft = editor.scrollLeft;
    }

    // ==================== 代码缩进功能 ====================

    const INDENT_SIZE = 4;
    const INDENT_CHAR = ' ';

    function initCodeIndent() {
        const editor = document.getElementById('code-editor');
        if (!editor) return;

        editor.addEventListener('keydown', handleIndentKeydown);

        // 粘贴时自动格式化代码
        editor.addEventListener('paste', function(e) {
            // 延迟处理，等粘贴内容到位后再格式化
            setTimeout(function() {
                const code = editor.value;
                if (code && code.trim().length > 0) {
                    const formatted = autoFormatCode(code, editorLanguage || 'vba');
                    const scrollTop = editor.scrollTop;
                    editor.value = formatted;
                    editor.scrollTop = scrollTop;
                    updateLineNumbers();
                    updateSyntaxHighlight();
                    updateFoldButtons();
                }
            }, 10);
        });
    }

    function handleIndentKeydown(e) {
        const editor = e.target;

        // Tab 键缩进
        if (e.key === 'Tab' && !e.shiftKey) {
            e.preventDefault();
            handleTabIndent(editor);
            return;
        }

        // Shift+Tab 反缩进
        if (e.key === 'Tab' && e.shiftKey) {
            e.preventDefault();
            handleShiftTab(editor);
            return;
        }

        // Enter 键自动缩进
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleAutoIndent(editor);
            return;
        }

        // Backspace 键智能删除缩进
        if (e.key === 'Backspace') {
            const pos = editor.selectionStart;
            const lineStart = getLineStartPos(editor.value, pos);
            const lineText = editor.value.substring(lineStart, pos);

            // 如果光标在行首且前面有缩进，删除一个缩进级别
            if (lineText.length > 0 && lineText.trim() === '' && lineText.length >= INDENT_SIZE) {
                e.preventDefault();
                const indentToRemove = lineText.length % INDENT_SIZE === 0 ? INDENT_SIZE : lineText.length % INDENT_SIZE;
                const newValue = editor.value.substring(0, pos - indentToRemove) + editor.value.substring(pos);
                editor.value = newValue;
                editor.selectionStart = editor.selectionEnd = pos - indentToRemove;
                triggerEditorInput(editor);
                return;
            }
        }
    }

    function handleTabIndent(editor) {
        const start = editor.selectionStart;
        const end = editor.selectionEnd;
        const value = editor.value;

        // 如果有选中文本，对选中的所有行增加缩进
        if (start !== end) {
            const lineStart = getLineStartPos(value, start);
            const lineEnd = getLineEndPos(value, end);
            const selectedText = value.substring(lineStart, lineEnd);
            const lines = selectedText.split('\n');

            const indentedLines = lines.map(line => {
                if (line.trim() === '') return line;
                return INDENT_CHAR.repeat(INDENT_SIZE) + line;
            });

            const newText = indentedLines.join('\n');
            const newValue = value.substring(0, lineStart) + newText + value.substring(lineEnd);

            editor.value = newValue;
            editor.selectionStart = lineStart;
            editor.selectionEnd = lineStart + newText.length;
        } else {
            // 没有选中文本，在光标位置插入缩进
            const indent = INDENT_CHAR.repeat(INDENT_SIZE);
            const newValue = value.substring(0, start) + indent + value.substring(end);
            editor.value = newValue;
            editor.selectionStart = editor.selectionEnd = start + indent.length;
        }

        triggerEditorInput(editor);
    }

    function handleShiftTab(editor) {
        const start = editor.selectionStart;
        const end = editor.selectionEnd;
        const value = editor.value;

        const lineStart = getLineStartPos(value, start);
        const lineEnd = getLineEndPos(value, end);
        const selectedText = value.substring(lineStart, lineEnd);
        const lines = selectedText.split('\n');

        const unindentedLines = lines.map(line => {
            if (line.startsWith(INDENT_CHAR.repeat(INDENT_SIZE))) {
                return line.substring(INDENT_SIZE);
            } else if (line.startsWith(' ') || line.startsWith('\t')) {
                // 删除尽可能多的空格
                let i = 0;
                while (i < INDENT_SIZE && i < line.length && (line[i] === ' ' || line[i] === '\t')) {
                    i++;
                }
                return line.substring(i);
            }
            return line;
        });

        const newText = unindentedLines.join('\n');
        const newValue = value.substring(0, lineStart) + newText + value.substring(lineEnd);

        editor.value = newValue;
        editor.selectionStart = lineStart;
        editor.selectionEnd = lineStart + newText.length;

        triggerEditorInput(editor);
    }

    function handleAutoIndent(editor) {
        const pos = editor.selectionStart;
        const value = editor.value;

        // 获取当前行内容
        const lineStart = getLineStartPos(value, pos);
        const currentLine = value.substring(lineStart, pos);

        // 计算当前行的缩进
        const indentMatch = currentLine.match(/^[\s]*/);
        let baseIndent = indentMatch ? indentMatch[0] : '';

        // 检测是否需要增加缩进（块开始）
        const shouldIncreaseIndent = checkShouldIncreaseIndent(currentLine, editorLanguage);

        // 插入换行和缩进
        let newIndent = baseIndent;
        if (shouldIncreaseIndent) {
            newIndent += INDENT_CHAR.repeat(INDENT_SIZE);
        }

        const insertText = '\n' + newIndent;
        const newValue = value.substring(0, pos) + insertText + value.substring(pos);

        editor.value = newValue;
        editor.selectionStart = editor.selectionEnd = pos + insertText.length;

        triggerEditorInput(editor);
    }

    function checkShouldIncreaseIndent(line, lang) {
        const trimmed = line.trim();
        if (!trimmed) return false;

        switch (lang) {
            case 'vba':
                return /^(If|For|While|Do|With|Select|Sub|Function|Property|Type|Enum)\b/i.test(trimmed) &&
                       !/\bEnd\s+(If|For|While|Do|With|Select|Sub|Function|Property|Type|Enum)\b/i.test(trimmed);
            case 'python':
                return /:\s*$/.test(trimmed) && !/^\s*#/.test(trimmed);
            case 'javascript':
                return /\{\s*$/.test(trimmed) ||
                       /^(function|class|if|for|while|switch|try|catch|finally|else|do)\b.*[^}]\s*$/.test(trimmed);
            default:
                return false;
        }
    }

    function getLineStartPos(text, pos) {
        const idx = text.lastIndexOf('\n', pos - 1);
        return idx === -1 ? 0 : idx + 1;
    }

    function getLineEndPos(text, pos) {
        const idx = text.indexOf('\n', pos);
        return idx === -1 ? text.length : idx;
    }

    function triggerEditorInput(editor) {
        // 触发 input 事件以更新行号和高亮
        const event = new Event('input', { bubbles: true });
        editor.dispatchEvent(event);
    }

    // 自动格式化代码缩进
    function autoFormatCode(code, lang) {
        if (!code) return code;

        const lines = code.split('\n');
        const formattedLines = [];
        let indentLevel = 0;
        const indentStr = INDENT_CHAR.repeat(INDENT_SIZE);

        for (let i = 0; i < lines.length; i++) {
            let line = lines[i];
            const trimmed = line.trim();

            // 空行保持空
            if (trimmed === '') {
                formattedLines.push('');
                continue;
            }

            // 检测是否是块结束（先减少缩进）
            const isBlockEnd = checkIsBlockEnd(trimmed, lang);
            if (isBlockEnd && indentLevel > 0) {
                indentLevel--;
            }

            // 检测单行块（如 VBA 的 If ... Then ... 单行形式）
            const isSingleLineBlock = checkIsSingleLineBlock(trimmed, lang);

            // 应用当前缩进
            const currentIndent = indentStr.repeat(Math.max(0, indentLevel));
            formattedLines.push(currentIndent + trimmed);

            // 检测是否是块开始（增加缩进）
            if (!isSingleLineBlock && checkIsBlockStart(trimmed, lang)) {
                indentLevel++;
            }
        }

        return formattedLines.join('\n');
    }

    function checkIsBlockEnd(line, lang) {
        const trimmed = line.trim();
        switch (lang) {
            case 'vba':
                if (/^End\s+(Sub|Function|Property|Type|Enum|If|With|Select)$/i.test(trimmed)) return true;
                if (/^Next(\s|$)/i.test(trimmed)) return true;
                if (/^Wend$/i.test(trimmed)) return true;
                if (/^Loop(\s|$)/i.test(trimmed)) return true;
                if (/^Else$/i.test(trimmed) || /^ElseIf\b/i.test(trimmed)) return true;
                if (/^Case\s/i.test(trimmed)) return true;
                break;
            case 'python':
                // Python 通过缩进判断，这里简化处理
                break;
            case 'javascript':
                if (/^\}/.test(trimmed) || /\}\s*$/.test(trimmed)) return true;
                break;
        }
        return false;
    }

    function checkIsSingleLineBlock(line, lang) {
        const trimmed = line.trim();
        switch (lang) {
            case 'vba':
                // VBA 单行 If: If ... Then ... (同一行)
                if (/^If\s+.+\s+Then\s+.+/i.test(trimmed)) return true;
                // 单行 For 不太常见，忽略
                break;
            case 'python':
                break;
            case 'javascript':
                if (/\{.*\}/.test(trimmed)) return true;
                break;
        }
        return false;
    }

    function checkIsBlockStart(line, lang) {
        const trimmed = line.trim();
        switch (lang) {
            case 'vba':
                if (/^(Sub|Function|Property\s+(Get|Let|Set)|Type|Enum)\s+\w+/i.test(trimmed)) return true;
                if (/^If\s+.+\s+Then\s*$/i.test(trimmed)) return true;
                if (/^For\s+/i.test(trimmed) && !/Next/i.test(trimmed)) return true;
                if (/^While\s+/i.test(trimmed)) return true;
                if (/^Do\s*(While|Until)?/i.test(trimmed)) return true;
                if (/^With\s+/i.test(trimmed)) return true;
                if (/^Select\s+Case\s+/i.test(trimmed)) return true;
                break;
            case 'python':
                if (/^(def|class)\s+\w+.*:\s*$/.test(trimmed)) return true;
                if (/^(if|elif|for|while|with|try|except|finally)\b.*:\s*$/.test(trimmed)) return true;
                break;
            case 'javascript':
                if (/^(function|class)\s+\w+/.test(trimmed) && /\{[^}]*$/.test(trimmed)) return true;
                if (/\{[^}]*$/.test(trimmed) && /(if|for|while|switch|try|catch|else|do)\b/.test(trimmed)) return true;
                break;
        }
        return false;
    }

    // ==================== 代码折叠功能 ====================

    const foldedBlocks = new Map(); // key: originalStartLine, value: { endLine: number, visible: boolean, foldedLines: string[], placeholderLine: number }

    function initCodeFolding() {
        const lineNumbers = document.getElementById('line-numbers');
        if (!lineNumbers) return;

        // 事件委托：点击折叠按钮
        lineNumbers.addEventListener('click', function(e) {
            const foldBtn = e.target.closest('.code-fold-btn');
            if (foldBtn) {
                e.stopPropagation();
                const lineIndex = parseInt(foldBtn.dataset.lineIndex);
                toggleFold(lineIndex);
            }
        });

        // 初始更新折叠按钮
        updateFoldButtons();
    }

    function updateFoldButtons() {
        const editor = document.getElementById('code-editor');
        const lineNumbers = document.getElementById('line-numbers');

        if (!editor || !lineNumbers) return;

        const code = editor.value;
        const lines = code.split('\n');
        const lang = editorLanguage || 'vba';

        // 清除旧的折叠按钮
        lineNumbers.querySelectorAll('.code-fold-btn').forEach(btn => btn.remove());

        // 为每个可折叠块的起始行添加折叠按钮
        const foldableLines = findFoldableBlocks(lines, lang);

        foldableLines.forEach(block => {
            const lineEl = lineNumbers.querySelector(`[data-line="${block.startLine + 1}"]`);
            if (lineEl) {
                const foldBtn = document.createElement('span');
                foldBtn.className = 'code-fold-btn';
                foldBtn.dataset.lineIndex = block.startLine;
                foldBtn.dataset.endLine = block.endLine;
                foldBtn.textContent = '▼';
                foldBtn.title = '点击折叠';

                // 检查是否已折叠
                const folded = foldedBlocks.get(block.startLine);
                if (folded && !folded.visible) {
                    foldBtn.classList.add('folded');
                    foldBtn.textContent = '▶';
                    foldBtn.title = '点击展开';
                }

                // 直接绑定点击事件（双重保险）
                foldBtn.addEventListener('click', function(e) {
                    e.stopPropagation();
                    e.preventDefault();
                    const idx = parseInt(this.dataset.lineIndex);
                    toggleFold(idx);
                });

                lineEl.insertBefore(foldBtn, lineEl.firstChild);
            }
        });
    }

    function findFoldableBlocks(lines, lang) {
        const blocks = [];
        const stack = [];

        // 块类型匹配映射
        const blockTypeMap = {
            'procedure': 'procedure-end',
            'if': 'if-end',
            'for': 'for-end',
            'while': 'while-end',
            'do': 'do-end',
            'with': 'with-end',
            'select': 'select-end',
            'function': 'function-end',
            'block': 'block-end'
        };

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;

            const blockStartType = getBlockStartType(line, lang);
            const blockEndType = getBlockEndType(line, lang);

            // 如果是块结束，找到匹配的块开始
            if (blockEndType && stack.length > 0) {
                // 从栈顶往下找匹配的类型
                for (let j = stack.length - 1; j >= 0; j--) {
                    const expectedEnd = blockTypeMap[stack[j].type];
                    if (expectedEnd === blockEndType) {
                        const block = stack.splice(j, 1)[0];
                        if (i > block.startLine) {
                            blocks.push({
                                startLine: block.startLine,
                                endLine: i,
                                type: block.type
                            });
                        }
                        break;
                    }
                }
            }

            // 如果是块开始，压入栈（块开始和结束可能在同一行，先处理结束再处理开始）
            if (blockStartType) {
                stack.push({ startLine: i, type: blockStartType });
            }
        }

        // 按起始行排序
        blocks.sort(function(a, b) { return a.startLine - b.startLine; });

        return blocks;
    }

    function getBlockStartType(line, lang) {
        const trimmed = line.trim();
        switch (lang) {
            case 'vba':
                if (/^(Sub|Function|Property\s+(Get|Let|Set)|Type|Enum)\s+\w+/i.test(trimmed)) return 'procedure';
                if (/^If\s+.+\s+Then\s*$/i.test(trimmed)) return 'if';
                if (/^For\s+/i.test(trimmed) && !/\bNext\b/i.test(trimmed)) return 'for';
                if (/^While\s+/i.test(trimmed)) return 'while';
                if (/^Do\s*(While|Until)?/i.test(trimmed)) return 'do';
                if (/^With\s+/i.test(trimmed)) return 'with';
                if (/^Select\s+Case\s+/i.test(trimmed)) return 'select';
                break;
            case 'python':
                if (/^(def|class)\s+\w+.*:\s*$/.test(trimmed)) return 'function';
                if (/^(if|elif|for|while|with|try|except|finally)\b.*:\s*$/.test(trimmed)) return 'block';
                break;
            case 'javascript':
                if (/^(function|class)\s+\w+/.test(trimmed) && /\{[^}]*$/.test(trimmed)) return 'function';
                if (/\{[^}]*$/.test(trimmed) && /(if|for|while|switch|try|catch|else|do)\b/.test(trimmed)) return 'block';
                break;
        }
        return null;
    }

    function getBlockEndType(line, lang) {
        const trimmed = line.trim();
        switch (lang) {
            case 'vba':
                if (/^End\s+(Sub|Function|Property|Type|Enum)$/i.test(trimmed)) return 'procedure-end';
                if (/^End\s+If$/i.test(trimmed)) return 'if-end';
                if (/^Next(\s|$)/i.test(trimmed)) return 'for-end';
                if (/^Wend$/i.test(trimmed)) return 'while-end';
                if (/^Loop(\s|$)/i.test(trimmed)) return 'do-end';
                if (/^End\s+With$/i.test(trimmed)) return 'with-end';
                if (/^End\s+Select$/i.test(trimmed)) return 'select-end';
                break;
            case 'python':
                // Python 通过缩进来判断，这里简化处理
                break;
            case 'javascript':
                if (/^\}/.test(trimmed) || /\}\s*$/.test(trimmed)) return 'block-end';
                break;
        }
        return null;
    }

    function toggleFold(lineIndex) {
        const lineNumbers = document.getElementById('line-numbers');
        if (!lineNumbers) return;

        const foldBtn = lineNumbers.querySelector(`.code-fold-btn[data-line-index="${lineIndex}"]`);
        if (!foldBtn) return;

        const endLine = parseInt(foldBtn.dataset.endLine);
        const currentState = foldedBlocks.get(lineIndex);

        if (currentState && !currentState.visible) {
            // 展开
            unfoldBlock(lineIndex);
        } else {
            // 折叠
            foldBlock(lineIndex, endLine);
        }
    }

    function foldBlock(startLine, endLine) {
        // 保存折叠状态（不修改代码内容）
        foldedBlocks.set(startLine, {
            endLine: endLine,
            visible: false
        });

        // 更新按钮状态
        const lineNumbers = document.getElementById('line-numbers');
        if (lineNumbers) {
            const foldBtn = lineNumbers.querySelector(`.code-fold-btn[data-line-index="${startLine}"]`);
            if (foldBtn) {
                foldBtn.classList.add('folded');
                foldBtn.textContent = '▶';
                foldBtn.title = '点击展开';
            }
        }

        // 应用折叠到显示层
        applyFoldToHighlight();
    }

    function unfoldBlock(startLine) {
        const folded = foldedBlocks.get(startLine);
        if (!folded || folded.visible) return;

        // 移除折叠状态
        foldedBlocks.delete(startLine);

        // 更新按钮状态
        const lineNumbers = document.getElementById('line-numbers');
        if (lineNumbers) {
            const foldBtn = lineNumbers.querySelector(`.code-fold-btn[data-line-index="${startLine}"]`);
            if (foldBtn) {
                foldBtn.classList.remove('folded');
                foldBtn.textContent = '▼';
                foldBtn.title = '点击折叠';
            }
        }

        // 应用展开到显示层
        applyFoldToHighlight();
    }

    // 展开所有折叠块（编辑代码时自动调用）
    function unfoldAll() {
        if (foldedBlocks.size === 0) return;
        foldedBlocks.clear();
        applyFoldToHighlight();
        updateFoldButtons();
    }

    // --- 来自 app.js:L24031-L24086 --- 已扩展语言切换器显示 ---
    // 切换编辑器显示/隐藏
    async function toggleEditor() {
        // 代码编辑器属于对话视图的子面板：切回对话视图并点亮左侧「对话」按钮
        if (typeof window.activateChatView === 'function') window.activateChatView();
        const leftPanel = document.getElementById('left-panel');
        const codeContainer = document.getElementById('code-with-lines');
        const actions = document.getElementById('editor-actions');
        const title = document.getElementById('editor-title');
        const langSelect = document.getElementById('editor-language');

        if (!leftPanel || !codeContainer) {
            ChatLog.error('EDITOR', '编辑器元素未找到', { leftPanel: !!leftPanel, codeContainer: !!codeContainer });
            return;
        }

        if (leftPanel.classList.contains('expanded')) {
            // 判断当前是编辑器还是任务规划
            var tpPanel = document.getElementById('task-planner-panel');
            var isTpActive = tpPanel && tpPanel.classList.contains('active');
            if (isTpActive) {
                // 当前是任务规划面板，切换到编辑器
                tpPanel.classList.remove('active');
                tpPanel.style.display = 'none';
                // 恢复编辑器面板显示
                var editorPanel = document.getElementById('editor-panel');
                if (editorPanel) editorPanel.style.display = '';
                leftPanel.classList.add('expanded');
                codeContainer.classList.remove('is-hidden');
                codeContainer.style.display = 'flex';
                if (actions) {
                    actions.classList.remove('is-hidden');
                    actions.style.display = 'flex';
                }
                if (title) {
                    title.classList.remove('is-hidden');
                    title.style.display = 'inline-flex';
                }
                if (langSelect) {
                    langSelect.classList.remove('is-hidden');
                    langSelect.style.display = 'inline-block';
                }
                return;
            }

            leftPanel.classList.remove('expanded');
            codeContainer.classList.add('is-hidden');
            codeContainer.style.display = 'none';
            if (actions) {
                actions.classList.add('is-hidden');
                actions.style.display = 'none';
            }
            if (title) {
                title.classList.add('is-hidden');
                title.style.display = 'none';
            }
            if (langSelect) {
                langSelect.classList.add('is-hidden');
                langSelect.style.display = 'none';
            }
            if (App.excelBridge && App.sidebarWidthBeforeEditor !== null) {
                try {
                    await App.excelBridge.SetSidebarWidth(App.sidebarWidthBeforeEditor);
                } catch (e) {
                    ChatLog.info('EDITOR', '恢复宽度失败:', e);
                }
            }
        } else {
            // 打开编辑器时，隐藏任务规划面板
            var tpPanel = document.getElementById('task-planner-panel');
            if (tpPanel) {
                tpPanel.classList.remove('active');
                tpPanel.style.display = 'none';
            }
            if (App.excelBridge) {
                try {
                    const currentWidth = await App.excelBridge.GetSidebarWidth();
                    App.sidebarWidthBeforeEditor = currentWidth;
                    ChatLog.info('EDITOR', '当前宽度:', currentWidth);
                    const newWidth = currentWidth + 400;
                    ChatLog.info('EDITOR', '设置新宽度:', newWidth);
                    await App.excelBridge.SetSidebarWidth(newWidth);
                    await new Promise(resolve => setTimeout(resolve, 100));
                } catch (e) {
                    ChatLog.info('EDITOR', '调整宽度失败:', e);
                }
            }
            leftPanel.classList.add('expanded');
            codeContainer.classList.remove('is-hidden');
            codeContainer.style.display = 'flex';
            if (actions) {
                actions.classList.remove('is-hidden');
                actions.style.display = 'flex';
            }
            if (title) {
                title.classList.remove('is-hidden');
                title.style.display = 'inline-flex';
            }
            if (langSelect) {
                langSelect.classList.remove('is-hidden');
                langSelect.style.display = '';
            }
            setTimeout(() => { initLineNumbers(); }, 50);
            if (App.excelBridge && App.sidebarWidthBeforeEditor !== null) {
                try {
                    const expectedWidth = App.sidebarWidthBeforeEditor + 400;
                    const actualWidth = await App.excelBridge.GetSidebarWidth();
                    ChatLog.info('EDITOR', '实际宽度:', actualWidth, '期望宽度:', expectedWidth);
                    if (actualWidth < expectedWidth) {
                        await App.excelBridge.SetSidebarWidth(expectedWidth);
                    }
                } catch (e) {
                    ChatLog.info('EDITOR', '验证宽度失败:', e);
                }
            }
        }
    }

    // --- 来自 app.js:L23990-L23995 ---
    // 关闭图片放大视图
    function closeZoom() {
        const overlay = document.querySelector('.image-zoom-overlay');
        if (overlay) {
            overlay.classList.remove('active');
            const overlayImg = overlay.querySelector('img');
            setTimeout(() => {
                if (overlayImg) overlayImg.src = '';
            }, 300);
        }
    }

    // --- 打开图片放大视图 ---
    function openZoom(imageSrc) {
        const overlay = document.querySelector('.image-zoom-overlay');
        const overlayImg = document.getElementById('zoom-image');
        if (overlay && overlayImg) {
            overlayImg.src = imageSrc;
            overlay.classList.add('active');
        }
    }

    // 将函数暴露到全局作用域
    window.App.closeZoom = closeZoom;
    window.App.openZoom = openZoom;

    // ==================== 代码执行功能 ====================

    // --- 来自 app.js:L20795-L20865 --- 已扩展为多语言分发 ---
    async function executeCode() {
        // 根据当前语言分发给对应的执行器
        if (editorLanguage === 'python') {
            return executePythonCode();
        } else if (editorLanguage === 'javascript') {
            return executeJSCode();
        }

        // 默认: VBA
        const code = document.getElementById('code-editor').value.trim();
        if (!code) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入代码</span>', true) : App.showToast('请输入代码');
            return;
        }

        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
            return;
        }

        const startTime = performance.now();

        try {
            const macroName = extractMacroName(code) || 'GeneratedMacro';

            let result = await App.excelBridge.ExecuteVBA(code, macroName);

            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);

            if (typeof result === 'object') {
                result = JSON.stringify(result);
            }

            try {
                const resultObj = JSON.parse(result);
                const timestamp = new Date().toLocaleString('zh-CN');

                let displayInfo = `📊 执行结果\n`;
                displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
                displayInfo += `⏰ 执行时间: ${executionTime} 秒\n`;
                displayInfo += `📝 宏名称: ${resultObj.macroName || macroName}\n`;
                displayInfo += `🕐 时间戳: ${timestamp}\n`;
                displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;

                if (resultObj.success) {
                    displayInfo += `✅ 状态: 执行成功\n`;
                    if (resultObj.returnValue !== undefined) {
                        displayInfo += `📤 返回值: ${resultObj.returnValue}\n`;
                    }
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + displayInfo + '</span>', true) : App.showToast(displayInfo, 'success', 5000);
                } else {
                    displayInfo += `❌ 状态: 执行失败\n`;
                    displayInfo += `⚠️ 错误信息: ${resultObj.error || '未知错误'}\n`;
                    App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + displayInfo + '</span>', true) : App.showToast(displayInfo, 'error', 5000);
                }
            } catch (e) {
                let displayInfo = `📊 执行结果\n`;
                displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
                displayInfo += `⏰ 执行时间: ${executionTime} 秒\n`;
                displayInfo += `📝 宏名称: ${macroName}\n`;
                displayInfo += `🕐 完成时间: ${new Date().toLocaleString('zh-CN')}\n`;
                displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
                displayInfo += `📄 返回结果:\n${result}`;
                App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + displayInfo + '</span>', true) : App.showToast(displayInfo, 'info', 5000);
            }
        } catch (error) {
            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);

            let errorInfo = `❌ 执行失败\n`;
            errorInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            errorInfo += `⏰ 失败时间: ${executionTime} 秒后\n`;
            errorInfo += `⚠️ 错误类型: ${error.name || 'Error'}\n`;
            errorInfo += `💥 错误消息: ${error.message}\n`;
            errorInfo += `━━━━━━━━━━━━━━━━━━━━`;
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + errorInfo + '</span>', true) : App.showToast(errorInfo, 'error', 5000);
        }
    }

    // --- 来自 app.js:L20885-L20937 ---
    // 解析代码中的变量声明
    function parseVariablesFromCode(code) {
        const variables = [];
        const lines = code.split('\n');

        for (let line of lines) {
            const trimmed = line.trim();

            // 匹配 Dim 语句：Dim var1 As Type, var2 As Type
            const dimMatch = trimmed.match(/^Dim\s+(.+?)\s+As\s+(\w+)/i);
            if (dimMatch) {
                const varNames = dimMatch[1].split(',').map(v => v.trim());
                const varType = dimMatch[2];

                varNames.forEach(varName => {
                    if (varName && !varName.includes('(')) {
                        variables.push({
                            name: varName,
                            type: varType,
                            value: '<未初始化>',
                            line: lines.indexOf(line) + 1
                        });
                    }
                });
            }

            // 匹配 Set 语句：Set obj = something
            const setMatch = trimmed.match(/^Set\s+(\w+)\s*=\s*(.+)/i);
            if (setMatch) {
                const existingVar = variables.find(v => v.name === setMatch[1]);
                if (existingVar) {
                    existingVar.value = setMatch[2].substring(0, 50) + (setMatch[2].length > 50 ? '...' : '');
                    existingVar.isObject = true;
                }
            }

            // 匹配赋值语句：var = value
            const assignMatch = trimmed.match(/^(\w[\w\d]*)\s*=\s*(.+)/);
            if (assignMatch && !trimmed.toLowerCase().startsWith('dim ') &&
                !trimmed.toLowerCase().startsWith('set ')) {
                const existingVar = variables.find(v => v.name === assignMatch[1]);
                if (existingVar) {
                    let value = assignMatch[2];
                    if (value.length > 50) value = value.substring(0, 50) + '...';
                    existingVar.value = value;
                }
            }

            // 匹配 For 循环：For col = 1 To ...
            const forMatch = trimmed.match(/^For\s+(\w[\w\d]*)\s*=\s*(.+)/i);
            if (forMatch) {
                const existingVar = variables.find(v => v.name === forMatch[1]);
                if (existingVar) {
                    let value = forMatch[2];
                    if (value.length > 50) value = value.substring(0, 50) + '...';
                    existingVar.value = value;
                }
            }

            // 匹配 For Each：For Each cell In rng
            const forEachMatch = trimmed.match(/^For\s+Each\s+(\w[\w\d]*)\s+In\s+(.+)/i);
            if (forEachMatch) {
                const existingVar = variables.find(v => v.name === forEachMatch[1]);
                if (existingVar) {
                    let value = forEachMatch[2];
                    if (value.length > 50) value = value.substring(0, 50) + '...';
                    existingVar.value = value;
                }
            }
        }

        return variables;
    }

    // 执行完一行后，根据该行内容更新变量值
    function updateVariableValuesFromLine(lineContent) {
        const trimmed = lineContent.trim();
        if (!trimmed) return;

        // 普通赋值：var = value
        const assignMatch = trimmed.match(/^(\w[\w\d]*)\s*=\s*(.+)/);
        if (assignMatch && !trimmed.toLowerCase().startsWith('dim ') &&
            !trimmed.toLowerCase().startsWith('set ')) {
            const existingVar = App.debugState.parsedVariables.find(v => v.name === assignMatch[1]);
            if (existingVar) {
                let value = assignMatch[2];
                if (value.length > 50) value = value.substring(0, 50) + '...';
                existingVar.value = value;
                ChatLog.info('EDITOR', '[变量更新] 赋值:', existingVar.name, '=', existingVar.value);
            }
        }

        // Set 赋值
        const setMatch = trimmed.match(/^Set\s+(\w[\w\d]*)\s*=\s*(.+)/i);
        if (setMatch) {
            const existingVar = App.debugState.parsedVariables.find(v => v.name === setMatch[1]);
            if (existingVar) {
                existingVar.value = setMatch[2].substring(0, 50) + (setMatch[2].length > 50 ? '...' : '');
                ChatLog.info('EDITOR', '[变量更新] Set:', existingVar.name, '=', existingVar.value);
            }
        }

        // For 循环
        const forMatch = trimmed.match(/^For\s+(\w[\w\d]*)\s*=\s*(.+)/i);
        if (forMatch) {
            const existingVar = App.debugState.parsedVariables.find(v => v.name === forMatch[1]);
            if (existingVar) {
                let value = forMatch[2];
                if (value.length > 50) value = value.substring(0, 50) + '...';
                existingVar.value = value;
                ChatLog.info('EDITOR', '[变量更新] For:', existingVar.name, '=', existingVar.value);
            } else {
                ChatLog.info('EDITOR', '[变量更新] For 未找到变量:', forMatch[1]);
            }
        }

        // For Each
        const forEachMatch = trimmed.match(/^For\s+Each\s+(\w[\w\d]*)\s+In\s+(.+)/i);
        if (forEachMatch) {
            const existingVar = App.debugState.parsedVariables.find(v => v.name === forEachMatch[1]);
            if (existingVar) {
                let value = forEachMatch[2];
                if (value.length > 50) value = value.substring(0, 50) + '...';
                existingVar.value = value;
                ChatLog.info('EDITOR', '[变量更新] For Each:', existingVar.name, '=', existingVar.value);
            } else {
                ChatLog.info('EDITOR', '[变量更新] For Each 未找到变量:', forEachMatch[1]);
            }
        }
    }

    // --- 来自 app.js:L20939-L20965 ---
    // 刷新变量监视窗口
    function refreshWatchWindow() {
        const watchList = document.getElementById('watch-list');

        if (!App.debugState.parsedVariables || App.debugState.parsedVariables.length === 0) {
            if (watchList) watchList.innerHTML = '<div class="watch-empty">无变量（代码中需要Dim声明）</div>';
            return;
        }

        let html = '';
        const syncVars = [];
        App.debugState.parsedVariables.forEach((v, index) => {
            // 若值仍为初始占位，判断该变量声明行是否已执行过
            const isExecuted = Array.from(App.debugState.executedLines).some(i => {
                const stmt = App.debugState.lines[i];
                return stmt && stmt.lineNumber >= v.line;
            });
            const displayValue = (v.value === '<未初始化>' && !isExecuted) ? '<等待执行>' : v.value;

            html += `
                <div class="watch-item ${isExecuted ? '' : 'dimmed'}">
                    <span class="watch-var-name">${v.name}</span>
                    <span class="watch-var-value">${displayValue}</span>
                    <span class="watch-var-type">${v.type}</span>
                </div>
            `;

            syncVars.push({
                name: v.name,
                value: displayValue,
                type: v.type || 'Variant'
            });
        });

        if (watchList) watchList.innerHTML = html || '<div class="watch-empty">无变量</div>';

        // 同步变量到独立调试窗口
        if (debugChannel && syncVars.length > 0) {
            debugChannel.postMessage({ type: 'watch', variables: syncVars });
        }
    }

    // ==================== 调试器核心功能 ====================

    // --- 来自 app.js:L20968-L20982 ---
    // ✅ 统一的逐行执行入口：根据状态自动选择启动或继续
    async function startOrStepInto() {
        if (!App.debugState.isDebugging) {
            // 未在调试模式 → 启动调试会话
            await startStepByStepExecution();
        } else if (App.debugState.currentLineIndex < App.debugState.lines.length) {
            // 正在调试且未完成 → 执行下一步
            await stepInto();
        } else {
            // ✅ 已执行完毕 → 重新开始运行！
            restartDebugSession();
        }

        // 更新按钮显示文本
        updateStepButtonUI();
    }

    // --- 来自 app.js:L20985-L21015 ---
    // ✅ 重新开始调试会话（重置所有状态）
    function restartDebugSession() {
        // 清除之前的执行状态
        clearAllLineMarks();

        // 重置调试状态
        App.debugState.currentLineIndex = 0;
        App.debugState.executedLines.clear();
        App.debugState.variables = {};
        App.debugState.isPaused = true;
        App.debugState.isRunning = false;
        App.debugState.stopRequested = false;
        App.debugState.startTime = performance.now();
        App.debugState.executionLog = [];
        App.debugState.logIndex = -1;

        // 刷新UI
        updateDebugUI();
        refreshWatchWindow();

        // 清空日志并显示重新启动信息
        const logContainer = document.getElementById('debug-log');
        if (logContainer) logContainer.innerHTML = '';

        addDebugLog('info', '🔄 重新启动调试会话');
        addDebugLog('info', `📝 共 ${App.debugState.lines.length} 行可执行语句`);
        if (App.debugState.parsedVariables.length > 0) {
            addDebugLog('info', `👁️ 发现 ${App.debugState.parsedVariables.length} 个变量`);
        }
        addDebugLog('info', '🚀 按 [Enter] 或点击 [下一步] 开始');

        // 高亮第一行
        highlightCurrentLineDebug();
    }

    // --- 来自 app.js:L21018-L21046 ---
    // 更新按钮的显示文本和状态
    function updateStepButtonUI() {
        const mainBtn = document.getElementById('step-exec-btn');

        if (App.debugState.isDebugging) {
            const isCompleted = App.debugState.currentLineIndex >= App.debugState.lines.length;
            const newText = isCompleted ? '🔄 重新运行' : '▶ 下一步';
            const newTitle = isCompleted ?
                '点击重新开始运行 (或按Enter)' :
                '执行下一行 (点击或按Enter)';

            if (mainBtn) {
                mainBtn.textContent = newText;
                mainBtn.title = newTitle;

                // ✅ 完成状态时添加脉冲动画提示用户可点击
                if (isCompleted) {
                    mainBtn.classList.add('pulse-animation');
                    setTimeout(() => {
                        mainBtn.classList.remove('pulse-animation');
                    }, 2000);
                }
            }
        } else {
            if (mainBtn) {
                mainBtn.textContent = '逐行执行';
                mainBtn.title = '开始逐行执行代码';
            }
        }
    }

    // --- 来自 app.js:L21048-L21099 ---
    async function startStepByStepExecution() {
        const code = document.getElementById('code-editor').value.trim();

        if (!code) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入代码</span>', true) : App.showToast('请输入代码', 'warning');
            return;
        }

        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接', 'error');
            return;
        }

        App.debugState.originalCode = code;
        App.debugState.lines = parseVBALines(code);
        App.debugState.currentLineIndex = 0;
        App.debugState.executedLines.clear();
        App.debugState.variables = {};

        // ✅ 解析变量并显示在监视窗口
        App.debugState.parsedVariables = parseVariablesFromCode(code);
        refreshWatchWindow();

        App.debugState.isDebugging = true;
        App.debugState.isPaused = true;
        App.debugState.isRunning = false;
        App.debugState.stopRequested = false;
        App.debugState.startTime = performance.now();

        // ✅ 显示调试面板并立即聚焦
        showDebugPanel();
        updateDebugUI();

        addDebugLog('info', '🚀 调试会话已启动');
        addDebugLog('info', `📝 共 ${App.debugState.lines.length} 行可执行语句`);
        if (App.debugState.parsedVariables.length > 0) {
            addDebugLog('info', `👁️ 发现 ${App.debugState.parsedVariables.length} 个变量`);
        }

        // 发送完整代码到后端执行并获取日志
        addDebugLog('info', '⏳ 正在执行代码并收集运行时状态...');
        try {
            const result = await App.excelBridge.DebugExecuteVBA(code);
            let resultObj = null;
            try {
                resultObj = JSON.parse(result);
            } catch (e) {
                addDebugLog('error', '❌ 无法解析后端返回结果');
                App.debugState.isDebugging = false;
                updateStepButtonUI();
                return;
            }

            if (!resultObj.success) {
                addDebugLog('error', `❌ 执行失败: ${resultObj.error || '未知错误'}`);
                App.debugState.isDebugging = false;
                updateStepButtonUI();
                return;
            }

            App.debugState.executionLog = resultObj.log || [];
            App.debugState.logIndex = -1;

            addDebugLog('success', `✅ 代码执行完成，收集到 ${App.debugState.executionLog.length} 条状态记录`);
            addDebugLog('info', '👆 点击 [下一步] 或按 [Enter] 开始逐行回放');

        } catch (error) {
            addDebugLog('error', `❌ 调试执行异常: ${error.message}`);
            App.debugState.isDebugging = false;
        }

        highlightCurrentLineDebug();

        const codeContainer = document.getElementById('code-with-lines');
        if (codeContainer) {
            codeContainer.classList.add('current-line-highlight');
        }

        scrollToTop();
        updateStepButtonUI();
    }

    // --- 来自 app.js:L21101-L21125 ---
    function parseVBALines(code) {
        const lines = code.split('\n');
        const executableLines = [];

        for (let i = 0; i < lines.length; i++) {
            const line = lines[i];
            const trimmed = line.trim();

            // 跳过空行和注释行
            if (!trimmed || trimmed.startsWith("'")) {
                continue;
            }

            // ✅ 新策略：每一行都作为独立可执行语句
            // 这样用户可以真正地"逐句"执行每行代码
            executableLines.push({
                lineNumber: i + 1,
                content: line,
                type: 'single',  // 全部标记为单行
                isExecutable: true
            });
        }

        return executableLines;
    }

    // --- 来自 app.js:L21127-L21138 ---
    function isBlockStart(line) {
        const lower = line.toLowerCase();
        return lower.startsWith('if ') && lower.includes(' then') ||
               lower.startsWith('for ') ||
               lower.startsWith('while ') ||
               lower.startsWith('do ') ||
               lower.startsWith('select case') ||
               lower.startsWith('with ') ||
               lower.startsWith('sub ') ||
               lower.startsWith('function ') ||
               lower.startsWith('type ');
    }

    // --- 来自 app.js:L21140-L21151 ---
    function getBlockType(line) {
        const lower = line.toLowerCase();
        if (lower.startsWith('if ')) return 'if';
        if (lower.startsWith('for ')) return 'for';
        if (lower.startsWith('while ')) return 'while';
        if (lower.startsWith('do ')) return 'do';
        if (lower.startsWith('select')) return 'select';
        if (lower.startsWith('with ')) return 'with';
        if (lower.startsWith('sub ')) return 'sub';
        if (lower.startsWith('function ')) return 'function';
        return 'unknown';
    }

    // 应用一条日志记录到调试状态（回放）
    function applyLogEntry(entry) {
        if (!entry) return;

        // 找到对应的代码行索引
        const lineIdx = App.debugState.lines.findIndex(stmt => stmt.lineNumber === entry.lineNumber);
        if (lineIdx >= 0) {
            App.debugState.currentLineIndex = lineIdx;
            App.debugState.executedLines.add(lineIdx);
        }

        // 更新变量值
        if (entry.variables) {
            Object.entries(entry.variables).forEach(([name, value]) => {
                const variable = App.debugState.parsedVariables.find(v => v.name === name);
                if (variable) {
                    variable.value = value;
                }
            });
            refreshWatchWindow();
        }

        // 添加日志
        const isError = entry.code && entry.code.startsWith('ERROR');
        const status = isError ? 'error' : (entry.lineNumber === 0 ? 'skipped' : 'success');
        addDebugLog(status, `行 ${entry.lineNumber}: ${(entry.code || '').substring(0, 60)}`);
        if (entry.variables && Object.keys(entry.variables).length > 0) {
            const varStr = Object.entries(entry.variables).map(([k, v]) => `${k}=${v}`).join(', ');
            addDebugLog('info', `   变量: ${varStr.substring(0, 120)}`);
        }

        // 高亮当前行
        highlightCurrentLineDebug();
        updateDebugUI();
    }

    // --- 来自 app.js:L21153-L21170 ---
    async function stepInto() {
        if (!App.debugState.isDebugging) return;

        const log = App.debugState.executionLog;
        if (!log || log.length === 0) {
            addDebugLog('warning', '⚠️ 没有执行日志可回放');
            return;
        }

        App.debugState.logIndex++;
        if (App.debugState.logIndex >= log.length) {
            completeExecution();
            updateStepButtonUI();
            return;
        }

        applyLogEntry(log[App.debugState.logIndex]);
        updateStepButtonUI();
    }

    // --- 来自 app.js:L21172-L21200 ---
    async function stepOver() {
        // 日志回放模式下，stepOver 与 stepInto 行为相同
        await stepInto();
    }

    // --- 来自 app.js:L21202-L21230 ---
    async function continueExecution() {
        if (!App.debugState.isDebugging) return;

        App.debugState.isPaused = false;
        App.debugState.isRunning = true;
        updateDebugUI();
        addDebugLog('info', '▶️ 继续执行...');

        const log = App.debugState.executionLog;
        while (App.debugState.logIndex < log.length - 1 && !App.debugState.isPaused && !App.debugState.stopRequested) {
            App.debugState.logIndex++;
            applyLogEntry(log[App.debugState.logIndex]);
            await new Promise(resolve => setTimeout(resolve, 80));
        }

        if (App.debugState.stopRequested) {
            stopExecutionComplete();
        } else if (App.debugState.logIndex >= log.length - 1) {
            completeExecution();
        }
    }

    // --- 来自 app.js:L21233-L21284 ---
    // 检查某行是否可独立执行
    function isExecutableStatement(lineContent) {
        const trimmed = lineContent.trim();

        // 空行不能执行
        if (!trimmed) return false;

        // 注释行不执行（但保留注释内容用于显示）
        if (trimmed.startsWith("'")) return false;

        // ❌ 跳过过程/函数定义
        const declarationPatterns = [
            /^(sub|function|property)\s+/i,           // 过程/函数定义
            /^end\s+(sub|function|property|type)/i,   // 结束定义
            /^type\s+\w+/i,                           // 类型定义
            /^dim\s+/i,                               // 变量声明
        ];

        for (let pattern of declarationPatterns) {
            if (pattern.test(trimmed)) {
                return false;
            }
        }

        // ❌ 跳过控制流语句（这些必须配对，不能单独执行）
        const controlFlowPatterns = [
            /^(for\s+each|for\s+\w+\s*=\s*)/i,       // For Each / For i =
            /^(while\s|do\s|loop\s)/i,               // While / Do / Loop
            /^(if\s+.+\bthen)\s*$/i,                 // If ... Then (单行)
            /^else(\s+if)?\s*$/i,                     // Else / Else If
            /^(select\s+case|case\s)/i,              // Select Case / Case
            /^with\s+/i,                              // With
            /^(next\b|wend\b|end\s+(if|with|select))/i,  // 结束语句
            /^(on\s+error\s+(goto|resume))/i,         // 错误处理声明
            /^(exit\s+(sub|function|for|do|while|property))/i, // Exit语句
            /^(goto\s|gosub\s|return\b)/i,            // 跳转语句
            /^(resume\s(next)?\b)/i,                  // Resume
        ];

        for (let pattern of controlFlowPatterns) {
            if (pattern.test(trimmed)) {
                return false; // 控制流语句不可单独执行
            }
        }

        // ✅ 只有真正的操作语句可以独立执行：
        // - Set赋值: Set wb = ThisWorkbook
        // - 普通赋值: x = 1, sheetName = ws.Name
        // - 方法调用: .Copy, .SaveAs, MsgBox(), Worksheets.Add
        // - 对象操作: Range("A1").Value = 100

        return true;
    }

    // --- 来自 app.js:L21286-L21414 ---
    async function executeCurrentLine() {
        // 日志回放模式下，executeCurrentLine 等价于 stepInto
        await stepInto();
    }

    // --- 来自 app.js:L21416-L21475 ---
    function buildExecutableCode(stmt, macroName) {
        let code = '';

        // ✅ 收集前面已执行的变量声明和赋值语句（保持变量状态）
        const preambleStatements = [];
        for (let i = 0; i < App.debugState.currentLineIndex; i++) {
            const prevStmt = App.debugState.lines[i];
            // 只收集Dim、Set、赋值等能建立变量的语句
            if (prevStmt && isVariableStatement(prevStmt.content)) {
                preambleStatements.push(prevStmt.content);
            }
            // 提取 For 循环的变量初始化（For col = 1 To ... → col = 1）
            const forInit = extractForLoopInit(prevStmt ? prevStmt.content : '');
            if (forInit) {
                preambleStatements.push(forInit);
            }
        }

        if (stmt.type === 'single') {
            // ✅ 单行语句：先执行前面的变量声明，再执行当前行
            code = `Sub ${macroName}()\n`;
            code += `    On Error GoTo ErrorHandler\n`;

            // 添加前面的变量声明
            if (preambleStatements.length > 0) {
                code += `    ' ===== 前面行的变量声明（保持作用域）=====\n`;
                preambleStatements.forEach(s => {
                    code += `    ${s}\n`;
                });
                code += `    ' =========================================\n`;
            }

            // 执行当前行
            code += `    ${stmt.content}\n`;

            // 添加运行时变量值输出
            const varOutput = buildVariableOutputCode();
            if (varOutput) {
                code += `    ' ===== 调试变量输出 =====\n`;
                code += varOutput;
                code += `    ' ========================\n`;
            }

            code += `    Exit Sub\n`;
            code += `ErrorHandler:\n`;
            code += `    Debug.Print "Error " & Err.Number & ": " & Err.Description\n`;
            code += `End Sub`;
        } else {
            // ✅ 多行代码块
            code = `Sub ${macroName}()\n`;
            code += `    On Error GoTo ErrorHandler\n`;

            // 添加前面的变量声明
            if (preambleStatements.length > 0) {
                preambleStatements.forEach(s => {
                    code += `    ${s}\n`;
                });
            }

            const lines = stmt.content.split('\n');
            lines.forEach(line => {
                if (line.trim()) {
                    code += `    ${line}\n`;
                }
            });

            const varOutput = buildVariableOutputCode();
            if (varOutput) {
                code += `    ' ===== 调试变量输出 =====\n`;
                code += varOutput;
                code += `    ' ========================\n`;
            }

            code += `    Exit Sub\n`;
            code += `ErrorHandler:\n`;
            code += `    Debug.Print "Error " & Err.Number & ": " & Err.Description\n`;
            code += `End Sub`;
        }

        return code;
    }

    // 构建 VBA 变量值输出代码，供后端读取并返回真正的运行时值
    function buildVariableOutputCode() {
        if (!App.debugState.parsedVariables || App.debugState.parsedVariables.length === 0) {
            return '';
        }
        let code = '    Dim __debugOutput As String\n';
        code += '    On Error Resume Next\n';
        App.debugState.parsedVariables.forEach(v => {
            const typeLower = (v.type || '').toLowerCase();
            if (typeLower === 'worksheet' || typeLower === 'workbook') {
                code += `    __debugOutput = __debugOutput & "${v.name}=" & ${v.name}.Name & "|"\n`;
            } else if (typeLower === 'range') {
                code += `    __debugOutput = __debugOutput & "${v.name}=" & ${v.name}.Address & "|"\n`;
            } else {
                code += `    __debugOutput = __debugOutput & "${v.name}=" & ${v.name} & "|"\n`;
            }
        });
        code += '    ActiveSheet.Cells(ActiveSheet.Rows.Count, ActiveSheet.Columns.Count).Value = __debugOutput\n';
        code += '    On Error GoTo ErrorHandler\n';
        return code;
    }

    // 解析后端返回的运行时变量值字符串，更新 parsedVariables
    function parseRuntimeVariableValues(debugOutput) {
        if (!debugOutput) return;
        const pairs = debugOutput.split('|');
        pairs.forEach(pair => {
            const eqIdx = pair.indexOf('=');
            if (eqIdx > 0) {
                const name = pair.substring(0, eqIdx).trim();
                const value = pair.substring(eqIdx + 1).trim();
                const variable = App.debugState.parsedVariables.find(v => v.name === name);
                if (variable) {
                    variable.value = value;
                }
            }
        });
    }

    // --- 来自 app.js:L21478-L21511 ---
    // ✅ 判断是否是变量相关的语句（需要保留以维持作用域）
    function isVariableStatement(lineContent) {
        const trimmed = lineContent.trim().toLowerCase();

        // 这些语句会创建或修改变量，需要保留
        const variablePatterns = [
            /^dim\s+/i,                              // Dim 声明
            /^set\s+\w+\s*=/i,                      // Set 赋值
            /^\w+\s*=\s*.*$/i,                       // 普通赋值（排除关键字）
            /^const\s+/i,                           // Const 常量
            /^redim\s+/i,                            // ReDim 数组
        ];

        // 排除这些不是变量操作的语句
        const excludePatterns = [
            /^(sub|function|end|for|next|if|else|endif|while|wend|do|loop|select|case|with|end\s+with|on\s+error|exit|resume|goto|err\.)/i,
            /^(msgbox|debug\.print)/i,
        ];

        // 检查是否应该排除
        for (let pattern of excludePatterns) {
            if (pattern.test(trimmed)) {
                return false;
            }
        }

        // 检查是否是变量相关语句
        for (let pattern of variablePatterns) {
            if (pattern.test(trimmed)) {
                return true;
            }
        }

        return false;
    }

    // 从 For 循环语句中提取变量初始赋值，用于 preamble 重建变量状态
    function extractForLoopInit(lineContent) {
        if (!lineContent) return null;
        const trimmed = lineContent.trim();
        // For col = 1 To ws.UsedRange.Columns.Count → col = 1
        const forNumMatch = trimmed.match(/^For\s+(\w+)\s*=\s*(.+?)\s+To\b/i);
        if (forNumMatch) {
            return `${forNumMatch[1]} = ${forNumMatch[2]}`;
        }
        return null;
    }

    // --- 来自 app.js:L21513-L21516 ---
    function stopExecution() {
        App.debugState.stopRequested = true;
        addDebugLog('warning', '⏹️ 停止请求已发送...');
    }

    // --- 来自 app.js:L21518-L21533 ---
    function stopExecutionComplete() {
        App.debugState.isDebugging = false;
        App.debugState.isRunning = false;
        App.debugState.isPaused = false;

        const totalTime = ((performance.now() - App.debugState.startTime) / 1000).toFixed(2);
        addDebugLog('warning', `⏹️ 执行已停止 (总时间: ${totalTime}s, 已执行: ${App.debugState.executedLines.size} 行)`);

        updateExecutionStatus('stopped');
        disableDebugControls();
        const codeContainer = document.getElementById('code-with-lines');
        if (codeContainer) {
            codeContainer.classList.remove('current-line-highlight');
        }
        clearAllLineMarks();
    }

    // --- 来自 app.js:L21535-L21554 ---
    function completeExecution() {
        App.debugState.isDebugging = false;
        App.debugState.isRunning = false;
        App.debugState.isPaused = false;

        const totalTime = ((performance.now() - App.debugState.startTime) / 1000).toFixed(2);
        addDebugLog('success', `🎉 执行完成！总时间: ${totalTime}s, 共执行 ${App.debugState.executedLines.size} 行`);

        updateExecutionStatus('completed');
        disableDebugControls();
        const codeContainer = document.getElementById('code-with-lines');
        if (codeContainer) {
            codeContainer.classList.remove('current-line-highlight');
        }

        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">逐句执行完成！共 ' + App.debugState.executedLines.size + ' 行，耗时 ' + totalTime + 's</span>', true) : App.showToast(`逐句执行完成！共 ${App.debugState.executedLines.size} 行，耗时 ${totalTime}s`, 'success');

        // 更新按钮状态为"完成"
        updateStepButtonUI();
    }

    // ==================== 调试器 UI 功能 ====================

    // --- 来自 app.js:L21556-L21581 ---
    function showDebugPanel() {
        // 不再显示内嵌面板，改为打开独立调试窗口
        openDebugWindow();

        const log = document.getElementById('debug-log');
        if (log) log.innerHTML = '';

        // 添加全局键盘快捷键：按 Enter 执行下一步
        if (!window._debugKeyListener) {
            window._debugKeyListener = function(e) {
                if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey && App.debugState && App.debugState.isDebugging) {
                    e.preventDefault();
                    startOrStepInto();  // 使用统一的入口函数
                }
            };
            document.addEventListener('keydown', window._debugKeyListener);
        }
    }

    // --- 来自 app.js:L21583-L21600 ---
    function closeDebugPanel() {
        if (App.debugState.isRunning) {
            if (confirm('调试正在进行中，确定要关闭吗？')) {
                stopExecution();
            }
            return;
        }

        resetDebugState();
        const panel = document.getElementById('step-debug-panel');
        const codeContainer = document.getElementById('code-with-lines');
        if (panel) panel.style.display = 'none';
        if (codeContainer) codeContainer.classList.remove('current-line-highlight');
        clearAllLineMarks();

        // 重置按钮状态
        updateStepButtonUI();
    }

    // --- 来自 app.js:L21602-L21615 ---
    function resetDebugState() {
        App.debugState = {
            isDebugging: false,
            isPaused: false,
            isRunning: false,
            currentLineIndex: -1,
            lines: [],
            originalCode: '',
            executedLines: new Set(),
            variables: {},
            parsedVariables: [], // 存储解析出的变量信息
            startTime: null,
            stopRequested: false,
            executionLog: [],
            logIndex: -1
        };
    }

    // --- 来自 app.js:L21617-L21653 ---
    // 高亮当前行（调试模式版本，使用 debugState）
    function highlightCurrentLineDebug() {
        const editor = document.getElementById('code-editor');
        if (!editor) return;

        const lines = editor.value.split('\n');

        if (App.debugState.currentLineIndex >= 0 && App.debugState.currentLineIndex < App.debugState.lines.length) {
            const currentStmt = App.debugState.lines[App.debugState.currentLineIndex];

            // 安全更新行号显示（如果元素存在）
            const currentLineEl = document.getElementById('current-line');
            if (currentLineEl) {
                currentLineEl.textContent = currentStmt.lineNumber;
            }

            // 使用行号系统高亮
            if (currentStmt.type === 'block') {
                for (let i = currentStmt.startLine; i <= currentStmt.endLine; i++) {
                    markLineStatus(i, 'clear');
                }
                markLineStatus(currentStmt.lineNumber, 'current');
            } else {
                clearAllLineMarks();
                markLineStatus(currentStmt.lineNumber, 'current');
            }

            // 同时选中文本
            const start = currentStmt.type === 'block' ? (currentStmt.startLine - 1) : (currentStmt.lineNumber - 1);
            const end = currentStmt.type === 'block' ? (currentStmt.endLine - 1) : (currentStmt.lineNumber - 1);

            editor.focus();
            editor.setSelectionRange(
                lines.slice(0, start).join('\n').length + (start > 0 ? 1 : 0),
                lines.slice(0, end + 1).join('\n').length
            );
        } else {
            clearAllLineMarks();
        }
    }

    // --- 来自 app.js:L21655-L21671 ---
    function updateDebugUI() {
        // ✅ 安全访问新UI元素
        const totalLinesEl = document.getElementById('total-lines');
        const currentLineEl = document.getElementById('current-line');

        if (totalLinesEl) {
            totalLinesEl.textContent = App.debugState.lines.length;
        }
        if (currentLineEl) {
            currentLineEl.textContent = App.debugState.currentLineIndex >= 0 ?
                App.debugState.lines[App.debugState.currentLineIndex]?.lineNumber || '0' : '0';
        }

        if (App.debugState.isPaused && !App.debugState.isRunning) {
            updateExecutionStatus('paused');
        }

        // 同步状态到独立调试窗口
        if (debugChannel) {
            const status = App.debugState.isDebugging
                ? (App.debugState.isPaused ? '已暂停' : '运行中')
                : '就绪';
            const lineInfo = App.debugState.isDebugging && App.debugState.lines.length > 0
                ? `行 ${App.debugState.currentLineIndex + 1}/${App.debugState.lines.length}`
                : '行 0/0';
            debugChannel.postMessage({ type: 'status', status, lineInfo });
        }
    }

    // --- 来自 app.js:L21673-L21689 ---
    function updateExecutionStatus(status) {
        const statusDot = document.getElementById('debug-status-dot');

        if (!statusDot) return;

        // 更新状态指示点颜色
        const colors = {
            'ready': '#4caf50',      // 绿色
            'running': '#2196f3',     // 蓝色
            'paused': '#ff9800',      // 橙色
            'completed': '#4caf50',   // 绿色
            'stopped': '#9e9e9e',     // 灰色
            'error': '#f44336'        // 红色
        };

        statusDot.style.backgroundColor = colors[status] || colors['ready'];
    }

    // --- 来自 app.js:L21691-L21707 ---
    function enableDebugControls() {
        const nextBtn = document.querySelector('.next-step-btn');
        if (nextBtn) nextBtn.disabled = false;
    }

    // 合并两处 disableDebugControls 定义（来自 app.js:L21696-L21707）
    function disableDebugControls() {
        // 禁用单个下一步按钮
        const nextBtn = document.querySelector('.next-step-btn');
        if (nextBtn) nextBtn.disabled = true;

        // 禁用所有调试按钮（除了停止按钮）
        document.querySelectorAll('.debug-btn').forEach(btn => {
            if (!btn.classList.contains('stop-btn')) {
                btn.disabled = true;
            }
        });
    }

    // --- 来自 app.js:L21709-L21719 ---
    function addDebugLog(type, message) {
        const logContainer = document.getElementById('debug-log');
        const time = new Date().toLocaleTimeString('zh-CN');

        const entry = document.createElement('div');
        entry.className = `log-entry ${type}`;
        entry.innerHTML = `<span class="log-time">[${time}]</span>${message}`;

        if (logContainer) {
            logContainer.appendChild(entry);
            logContainer.scrollTop = logContainer.scrollHeight;
        }

        // 同步到独立调试窗口
        if (debugChannel) {
            debugChannel.postMessage({ type: 'log', text: `[${time}] ${message}`, level: type });
        }
    }

    // --- 来自 app.js:L21894-L21899 ---
    function toggleDebug() {
        const debugInfo = document.getElementById('debug-info');
        if (debugInfo) {
            debugInfo.style.display = debugInfo.style.display === 'none' ? 'block' : 'none';
        }
    }

    // ==================== 撤销/恢复功能 ====================
    // 架构入口：所有撤销/重做统一通过后端 VB.NET 的 UndoBridge 执行
    //  - 撤销栈/重做栈由后端维护（Bridges/UndoBridge.vb）
    //  - 前端不维护独立历史，仅负责调用、显示消息、刷新按钮状态
    //  - 键盘快捷键见 keyboard-shortcuts.js: Ctrl+Z / Ctrl+Y / Ctrl+Shift+Z
    // --- 来自 app.js:L24098-L24110 （取第二处更完整的定义）---
    async function undoLastAction() {
        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
            return;
        }

        try {
            const result = await App.excelBridge.UndoLastAction();            // 解析结果，显示友好消息
            let msg = '已撤销';
            try {
                const data = typeof result === 'string' ? JSON.parse(result) : result;
                if (data && data.mock) {
                    msg = '【模拟】已撤销（Web版为模拟数据）';
                } else if (data && data.message) {
                    msg = data.message;
                }
            } catch (e) { (window.ChatLog || console).debug?.('UI', '[editor] 撤销结果非 JSON，保留默认文案:', e); }
            if (App.showToast) { App.showToast(msg, 'info'); } else if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') { App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + msg + '</span>', true); }
            updateUndoButtonState();
        } catch (error) {
            if (App.showToast) { App.showToast('撤销失败: ' + error.message, 'error'); } else if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') { App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">撤销失败: ' + error.message + '</span>', true); }
        }
    }

    async function redoLastAction() {
        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
            return;
        }

        try {
            const result = await App.excelBridge.RedoLastAction();
            let msg = '已恢复';
            try {
                const data = typeof result === 'string' ? JSON.parse(result) : result;
                if (data && data.mock) {
                    msg = '【模拟】已恢复（Web版为模拟数据）';
                } else if (data && data.message) {
                    msg = data.message;
                }
            } catch (e) { (window.ChatLog || console).debug?.('UI', '[editor] 恢复结果非 JSON，保留默认文案:', e); }
            if (App.showToast) { App.showToast(msg, 'info'); } else if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') { App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">' + msg + '</span>', true); }
            updateUndoButtonState();
        } catch (error) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">恢复失败: ' + error.message + '</span>', true) : App.showToast('恢复失败: ' + error.message);
        }
    }

    // --- 来自 app.js:L24112-L24122 （取第二处更完整的定义）---
    async function updateUndoButtonState() {
        const undoBtn = document.getElementById('undo-btn');
        const redoBtn = document.getElementById('redo-btn');
        // Web 模式 Mock Bridge 没有真实实现，跳过
        if (!App.excelBridge || typeof App.excelBridge.GetUndoStack !== 'function') return;
        // Word/PPT：撤销栈多为 not supported / bridge-call-failed，勿刷屏
        const _soft = (x) => {
            if (x == null) return true;
            const s = typeof x === 'string' ? x : (x.error || x.message || '');
            return /bridge-call-failed|not supported|0x80020006|未知名称|e is not a function/i.test(String(s));
        };

        try {
            const stackInfo = await App.excelBridge.GetUndoStack();
            const data = typeof stackInfo === 'string' ? JSON.parse(stackInfo) : stackInfo;
            if (_soft(data) || (data && data.success === false)) {
                if (undoBtn) undoBtn.disabled = true;
            } else if (undoBtn) {
                undoBtn.disabled = !data || data.count === 0;
            }
        } catch (e) {
            if (!_soft(e)) ChatLog.info('EDITOR', '更新撤销按钮状态失败:', e && e.message);
            if (undoBtn) undoBtn.disabled = true;
        }

        try {
            const redoStackInfo = await App.excelBridge.GetRedoStack();
            const redoData = typeof redoStackInfo === 'string' ? JSON.parse(redoStackInfo) : redoStackInfo;
            if (_soft(redoData) || (redoData && redoData.success === false)) {
                if (redoBtn) redoBtn.disabled = true;
            } else if (redoBtn) {
                redoBtn.disabled = !redoData || redoData.count === 0;
            }
        } catch (e) {
            if (!_soft(e)) ChatLog.info('EDITOR', '更新恢复按钮状态失败:', e && e.message);
            if (redoBtn) redoBtn.disabled = true;
        }
    }

    // ==================== 语言切换与多语言执行 ====================

    // 当前编辑器语言状态
    let editorLanguage = 'vba';

    function switchEditorLanguage(lang) {
        editorLanguage = lang;
        const editor = document.getElementById('code-editor');
        const langSelect = document.getElementById('editor-language');

        if (langSelect) langSelect.value = lang;

        const placeholders = {
            vba: '在此粘贴或编辑 VBA 代码...',
            python: '# 在此粘贴或编辑 Python 代码...\n# 提示: 使用 win32com.client 操作 Excel',
            javascript: '// 在此粘贴或编辑 JavaScript 代码...'
        };
        if (editor) editor.placeholder = placeholders[lang] || placeholders.vba;

        // 根据语言调整按钮可见性
        const stepBtn = document.getElementById('step-exec-btn');

        if (lang === 'vba') {
            if (stepBtn) stepBtn.style.display = '';
        } else {
            // Python/JS 不支持逐行调试（目前）
            if (stepBtn) stepBtn.style.display = 'none';
        }

        // 切换语言后重新应用语法高亮和折叠按钮
        updateSyntaxHighlight();
        updateFoldButtons();

        App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">已切换到 ' + lang.toUpperCase() + ' 模式</span>', true) : App.showToast(`已切换到 ${lang.toUpperCase()} 模式`, 'info', 1500);
    }

    async function executePythonCode() {
        const code = document.getElementById('code-editor').value.trim();
        if (!code) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入 Python 代码</span>', true) : App.showToast('请输入 Python 代码');
            return;
        }

        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
            return;
        }

        const startTime = performance.now();

        try {
            // 先检测 Python 环境
            const envCheck = await App.excelBridge.CheckPythonEnvironment();
            let envInfo;
            try { envInfo = JSON.parse(envCheck); } catch(e) { envInfo = {}; }

            if (!envInfo.available) {
                App.showToast('未检测到 Python 环境，请先安装 Python 3.8+', 'error', 5000);
                return;
            }

            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">正在执行 Python...</span>') : App.showToast('正在执行 Python...', 'info', 2000);
            let result = await App.excelBridge.ExecutePython(code);

            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);

            let resultObj;
            try {
                resultObj = JSON.parse(result);
            } catch (e) {
                App.showToast('执行结果解析失败: ' + result, 'error', 5000);
                return;
            }

            let displayInfo = `🐍 Python 执行结果\n`;
            displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            displayInfo += `⏰ 执行时间: ${executionTime} 秒\n`;
            displayInfo += `🐍 ${envInfo.version || 'Python'}\n`;
            displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;

            if (resultObj.success) {
                displayInfo += `✅ 状态: 执行成功 (退出码 ${resultObj.exitCode || 0})\n`;
                if (resultObj.stdout) {
                    displayInfo += `📤 输出:\n${resultObj.stdout}\n`;
                }
            } else {
                displayInfo += `❌ 状态: 执行失败\n`;
                if (resultObj.stderr) {
                    displayInfo += `⚠️ 错误:\n${resultObj.stderr}\n`;
                }
                if (resultObj.error) {
                    displayInfo += `⚠️ 系统错误: ${resultObj.error}\n`;
                }
            }

            App.showToast(displayInfo, resultObj.success ? 'success' : 'error', 8000);

        } catch (error) {
            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);
            let errorInfo = `❌ Python 执行异常\n`;
            errorInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            errorInfo += `⏰ 时间: ${executionTime} 秒\n`;
            errorInfo += `💥 ${error.message}\n`;
            App.showToast(errorInfo, 'error', 5000);
        }
    }

    // ==================== JavaScript 执行环境 ====================

    // $excel API - 供 JS 沙箱中操作 Excel
    function createExcelSandbox() {
        const logs = [];

        function log(...args) {
            const line = args.map(a => typeof a === 'object' ? JSON.stringify(a, null, 2) : String(a)).join(' ');
            logs.push(line);
        }

        const $excel = {
            log,

            range: function(address) {
                return {
                    async getValue() {
                        if (!App.excelBridge) throw new Error('Excel 未连接');
                        return await App.excelBridge.GetCellValue(address);
                    },
                    async setValue(value) {
                        if (!App.excelBridge) throw new Error('Excel 未连接');
                        const result = await App.excelBridge.SetCellValue(address, String(value));
                        const data = JSON.parse(result);
                        if (!data.success) throw new Error(data.error || '设置失败');
                        return data;
                    },
                    async getValues() {
                        if (!App.excelBridge) throw new Error('Excel 未连接');
                        const sel = await App.excelBridge.GetSelection();
                        return JSON.parse(sel);
                    }
                };
            },

            async getSelection() {
                if (!App.excelBridge) throw new Error('Excel 未连接');
                const result = await App.excelBridge.GetSelection();
                return JSON.parse(result);
            },

            async getActiveSheet() {
                if (!App.excelBridge) throw new Error('Excel 未连接');
                return await App.excelBridge.GetActiveSheetName();
            },

            async getActiveWorkbook() {
                if (!App.excelBridge) throw new Error('Excel 未连接');
                return await App.excelBridge.GetActiveWorkbookName();
            },

            async run(vbaCode) {
                if (!App.excelBridge) throw new Error('Excel 未连接');
                return await App.excelBridge.ExecuteVBA(vbaCode, 'JSExecMacro');
            },

            async fillData(data, type) {
                if (!App.excelBridge) throw new Error('Excel 未连接');
                const json = typeof data === 'string' ? data : JSON.stringify(data);
                return await App.excelBridge.FillDataToExcel(json, type || 'json');
            },

            getLogs() {
                return logs.join('\n');
            }
        };

        return $excel;
    }

    async function executeJSCode() {
        const code = document.getElementById('code-editor').value.trim();
        if (!code) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">请输入 JavaScript 代码</span>', true) : App.showToast('请输入 JavaScript 代码');
            return;
        }

        if (!App.excelBridge) {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">Excel 未连接</span>', true) : App.showToast('Excel 未连接');
            return;
        }

        const startTime = performance.now();
        const $excel = createExcelSandbox();

        // 注入 toolRegistry（聊天链路全局对象），供代码以裸标识符访问，避免 ReferenceError
        const toolRegistry = (App.ChatV2 && App.ChatV2.toolRegistry) || null;

        // 保存原始 console.log
        const originalLog = console.log;
        const originalError = console.error;
        const capturedLogs = [];

        function captureLog(...args) {
            const line = args.map(a => typeof a === 'object' ? JSON.stringify(a, null, 2) : String(a)).join(' ');
            capturedLogs.push(line);
            originalLog.apply(console, args);
        }
        function captureError(...args) {
            const line = args.map(a => typeof a === 'object' ? JSON.stringify(a, null, 2) : String(a)).join(' ');
            capturedLogs.push('[ERROR] ' + line);
            originalError.apply(console, args);
        }

        console.log = captureLog;
        console.error = captureError;

        try {
            App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function' ? App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">正在执行 JavaScript...</span>') : App.showToast('正在执行 JavaScript...', 'info', 1500);

            // 检测代码是否包含 async/await，决定如何包装
            const isAsync = /\basync\b|\bawait\b/.test(code);
            let fn;
            if (isAsync) {
                fn = new Function('$excel', 'toolRegistry', `return (async function() {\n${code}\n})();`);
            } else {
                fn = new Function('$excel', 'toolRegistry', code);
            }

            const result = await Promise.race([
                fn($excel, toolRegistry),
                new Promise((_, reject) => setTimeout(() => reject(new Error('执行超时（30秒）')), 30000))
            ]);

            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);

            let displayInfo = `🟨 JavaScript 执行结果\n`;
            displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            displayInfo += `⏰ 执行时间: ${executionTime} 秒\n`;
            displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            displayInfo += `✅ 状态: 执行成功\n`;

            const allOutput = [...capturedLogs, $excel.getLogs()].filter(Boolean).join('\n');
            if (allOutput) {
                displayInfo += `📤 输出:\n${allOutput}\n`;
            }
            if (result !== undefined) {
                const resultStr = typeof result === 'object' ? JSON.stringify(result, null, 2) : String(result);
                displayInfo += `📦 返回值: ${resultStr}\n`;
            }

            App.showToast(displayInfo, 'success', 8000);

        } catch (error) {
            const endTime = performance.now();
            const executionTime = ((endTime - startTime) / 1000).toFixed(3);

            let displayInfo = `❌ JavaScript 执行异常\n`;
            displayInfo += `━━━━━━━━━━━━━━━━━━━━\n`;
            displayInfo += `⏰ 时间: ${executionTime} 秒\n`;
            displayInfo += `💥 ${error.name}: ${error.message}\n`;
            if (!toolRegistry && /toolRegistry/.test(code)) {
                displayInfo += `💡 提示: 代码使用了 toolRegistry，但聊天引擎未初始化（tool-registry.js 未加载），请重启插件后重试\n`;
            }
            if (error.stack) {
                displayInfo += `📍 ${error.stack.split('\n')[1] || ''}\n`;
            }

            const allOutput = capturedLogs.join('\n');
            if (allOutput) {
                displayInfo += `📤 输出（异常前）:\n${allOutput}\n`;
            }

            App.showToast(displayInfo, 'error', 8000);
        } finally {
            console.log = originalLog;
            console.error = originalError;
        }
    }

    // ==================== 辅助函数 ====================

    // --- 来自 app.js:L20431-L20434 ---
    function extractMacroName(code) {
        // 支持中文、日文、韩文等 Unicode 标识符，不只限于 [A-Za-z0-9_]
        const match = code.match(/Sub\s+([^\s(]+)\s*\(/i);
        return match ? match[1] : null;
    }

    // --- 辅助函数：获取控制流提示信息 ---
    function getControlFlowHint(trimmedLine) {
        if (/^for/i.test(trimmedLine)) return 'For 循环需要配对的 Next 语句';
        if (/^while/i.test(trimmedLine)) return 'While 循环需要配对的 Wend 语句';
        if (/^do/i.test(trimmedLine)) return 'Do 循环需要配对的 Loop 语句';
        if (/^if/i.test(trimmedLine)) return 'If 语句需要配对的 End If (或多行结构)';
        if (/^select/i.test(trimmedLine)) return 'Select Case 需要配对的 End Select';
        if (/^with/i.test(trimmedLine)) return 'With 语句需要配对的 End With';
        return '此语句需要配对使用';
    }

    // ==================== 模块初始化 ====================
    
    // 初始化调试状态对象（如果尚未初始化）
    if (!App.debugState) {
        App.debugState = {
            isDebugging: false,
            isPaused: false,
            isRunning: false,
            currentLineIndex: -1,
            lines: [],
            originalCode: '',
            executedLines: new Set(),
            variables: {},
            parsedVariables: [],
            startTime: null,
            stopRequested: false
        };
    }

    // 保存打开编辑器前的侧边栏宽度
    if (!App.sidebarWidthBeforeEditor) {
        App.sidebarWidthBeforeEditor = null;
    }

    // 定期更新撤销按钮状态
    setInterval(updateUndoButtonState, 2000);

    // ==================== 暴露公共 API ====================
    // 同时挂载到 App 顶层，供其他模块直接调用
    App.updateLineNumbers = updateLineNumbers;
    App.initLineNumbers = initLineNumbers;
    App.showCodeEditor = showCodeEditor;
    App.openDebugWindow = openDebugWindow;
    App.autoFormatCode = autoFormatCode;
    App.updateSyntaxHighlight = updateSyntaxHighlight;
    App.updateFoldButtons = updateFoldButtons;

    // 兼容 index.html 中 onclick 直接调用
    window.toggleEditor = toggleEditor;
    window.clearEditor = clearEditor;
    window.formatEditorCode = formatEditorCode;
    window.executeCode = executeCode;
    window.switchEditorLanguage = switchEditorLanguage;
    window.undoLastAction = undoLastAction;
    window.redoLastAction = redoLastAction;

    window.App.Editor = {
        // 编辑器 UI
        updateLineNumbers,
        syncScroll,
        syncScrollFromLineNumbers,
        highlightCurrentLine,
        scrollToTop,
        markLineStatus,
        clearAllLineMarks,
        clearEditor,
        showCodeEditor,
        toggleEditor,
        closeZoom,
        initLineNumbers,

        // 语法高亮
        updateSyntaxHighlight,
        syncHighlightScroll,
        highlightSyntax,

        // 代码缩进
        initCodeIndent,
        handleTabIndent,
        handleShiftTab,
        handleAutoIndent,

        // 代码折叠
        initCodeFolding,
        updateFoldButtons,
        toggleFold,

        // 语言切换与多语言执行
        switchEditorLanguage,
        executePythonCode,
        executeJSCode,
        createExcelSandbox,

        // 代码执行
        executeCode,
        parseVariablesFromCode,
        refreshWatchWindow,

        // 调试器核心
        startOrStepInto,
        restartDebugSession,
        updateStepButtonUI,
        startStepByStepExecution,
        parseVBALines,
        isBlockStart,
        getBlockType,
        stepInto,
        stepOver,
        continueExecution,
        isExecutableStatement,
        executeCurrentLine,
        buildExecutableCode,
        isVariableStatement,
        stopExecution,
        stopExecutionComplete,
        completeExecution,
        openDebugWindow,

        // 调试器 UI
        showDebugPanel,
        closeDebugPanel,
        resetDebugState,
        highlightCurrentLineDebug,
        updateDebugUI,
        updateExecutionStatus,
        enableDebugControls,
        disableDebugControls,
        addDebugLog,
        toggleDebug,

        // 撤销/恢复功能
        undoLastAction,
        redoLastAction,
        updateUndoButtonState,

        // 辅助函数
        extractMacroName,

        // 兼容别名
        stopDebug: stopExecution
    };

})();