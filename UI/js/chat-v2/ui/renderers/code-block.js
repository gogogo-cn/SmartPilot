/**
 * ============================================================
 * chat-v2/ui/renderers/code-block.js — 代码块检测/清洗（纯函数子模块）
 * ============================================================
 * 从 chat-renderer.js 拆出的独立子模块：
 *   - detectAndFormatCodeBlocks(text) ：自动识别未包裹的 VBA/代码片段并加 ``` 围栏
 *   - stripCodeFences(code)           ：剥掉代码块首尾的 ``` 围栏
 *
 * 两者都是纯字符串，DOM-free，可被 node 直接单元测试。
 * 渲染/美化（beautifyCodeBlocks/renderCodeOnly）仍留在 chat-renderer.js，
 * 因依赖 document 与 executeCodeBlock/copyCode 等宿主能力。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 代码特征检测：识别未格式化的代码片段并补成 Markdown 代码块。
     * 纯函数，无 DOM 依赖。
     * @param {string} text 原始文本
     * @returns {string} 处理后的文本
     */
    function detectAndFormatCodeBlocks(text) {
        if (!text) return text;

        // 首先清理已存在的 ``` 代码块标记（可能是 AI 错误生成的转义版本如 \`\`\`）
        // 也清理 `` ` `` 包裹的行内代码
        text = text.replace(/\\`\\`\\`/g, '```');

        const lines = text.split('\n');
        const codeBlocks = [];
        let i = 0;

        while (i < lines.length) {
            const line = lines[i];
            const trimmed = line.trim();

            // 跳过已在代码块中的内容
            if (trimmed.startsWith('```')) {
                i++;
                while (i < lines.length && !lines[i].trim().startsWith('```')) i++;
                i++;
                continue;
            }

            // 检测 VBA 代码起始（Sub/Function 开头，且是独立一行）
            if (/^(Sub|Function)\s+\w+/i.test(trimmed) &&
                !trimmed.includes('。') && !trimmed.includes('，')) {
                const startIdx = i;
                let codeLines = [line];
                i++;
                // 收集到 End Sub / End Function
                while (i < lines.length && !/^End\s+(Sub|Function|Property)/i.test(lines[i].trim())) {
                    codeLines.push(lines[i]);
                    i++;
                }
                if (i < lines.length) {
                    codeLines.push(lines[i]); // 包含 End 行
                    i++;
                }
                // 至少 3 行才算有效代码块
                if (codeLines.length >= 3) {
                    codeBlocks.push({ startIdx, endIdx: i, code: codeLines.join('\n'), lang: 'vba' });
                }
                continue;
            }

            // 检测孤立的 VBA 代码行（Dim, Set, Range, Cells, Sheets 等）
            if (/^\s*(Dim|Set|Range|Cells|Sheets|Worksheets|ThisWorkbook|ActiveSheet|Workbooks)\b/i.test(line) &&
                /\.(Formula|Value|Interior|Font|Borders|Range|Cells|Column|Row|Copy|Delete|Insert|Clear|Merge|Unmerge)\b|=\s*"[^"]*"|=\s*'[^']*'/.test(line)) {
                // 向前查找是否在代码块中（检查上下文）
                const prevLines = lines.slice(Math.max(0, i - 5), i);
                const hasCodeContext = prevLines.some(l => /(Sub|Function|Dim|Set|End)\s/i.test(l));

                if (hasCodeContext) {
                    const startIdx = Math.max(0, i - 5);
                    const codeLines = [...prevLines.filter(l => l.trim() !== ''), line];
                    i++;
                    // 继续收集后续可能的代码行
                    while (i < lines.length &&
                           /^\s*(Dim|Set|Range|Cells|\.Formula|\.Value|\.Interior|\.Font|\.Borders|\.Copy|\.Delete|End\s+Sub|End\s+Function)\b/i.test(lines[i])) {
                        codeLines.push(lines[i]);
                        i++;
                    }
                    if (codeLines.length >= 2) {
                        codeBlocks.push({ startIdx, endIdx: i, code: codeLines.join('\n'), lang: 'vba' });
                    }
                    continue;
                }
            }

            i++;
        }

        // 如果检测到代码块，重新构建文本
        if (codeBlocks.length > 0) {
            // 去重和合并重叠
            const uniqueBlocks = [];
            for (const block of codeBlocks) {
                const isOverlapped = uniqueBlocks.some(b =>
                    !(block.endIdx <= b.startIdx || block.startIdx >= b.endIdx)
                );
                if (!isOverlapped) {
                    uniqueBlocks.push(block);
                }
            }

            // 从后向前替换（避免位置偏移）
            const resultLines = [...lines];
            uniqueBlocks.sort((a, b) => b.startIdx - a.startIdx);

            for (const block of uniqueBlocks) {
                const replacement = ['```' + block.lang, block.code, '```'];
                resultLines.splice(block.startIdx, block.endIdx - block.startIdx, ...replacement);
            }

            return resultLines.join('\n');
        }

        return text;
    }

    /**
     * 剥掉代码块的首尾 ``` 围栏。
     * 纯函数，无 DOM 依赖。
     * @param {string} code 代码内容
     * @returns {string} 去除围栏后的代码
     */
    function stripCodeFences(code) {
        if (!code) return '';
        let c = code.trim();
        // Remove opening fence: ```lang or ```
        c = c.replace(/^```[\w\-]*\s*\n?/i, '');
        // Remove closing fence: ```
        c = c.replace(/\n?```[\s\S]*$/i, '');
        return c.trim();
    }

    App.ChatV2.renderers.codeBlock = {
        detectAndFormatCodeBlocks,
        stripCodeFences
    };
    ChatLog.info?.('CHAT-V2', '✅ renderers/code-block 已加载');
})();