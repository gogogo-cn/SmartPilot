/*
 * chat-v2/ui/renderers/markdown.js - Markdown / text-cleaning pure-function submodule
 * Extracted from chat-renderer.js.
 * Pure string processing (DOM-free), unit-testable with node.
 * Self-contained escapeHtml; delegates codeBlock detection to code-block.js.
 */

(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;

    function escapeHtml(str) {
        if (str == null) return '';
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(str) : '';
    }

    function detectAndFormatCodeBlocks(text) {
        return App.ChatV2.renderers.codeBlock ? App.ChatV2.renderers.codeBlock.detectAndFormatCodeBlocks(text) : text;
    }
    function renderMarkdown(text) {
        if (!text) return '';

        // 预处理：自动检测并格式化未包裹的代码块
        text = detectAndFormatCodeBlocks(text);
        // 【空表清洗】弱模型总结时经常输出只有表头、"$"/空白占位数据行的假表格，
        // 渲染前统一清掉（详见 sanitizePlaceholderTables 注释）
        text = sanitizePlaceholderTables(text);
        
        try {
            if (window.marked && typeof marked.parse === 'function') {
                const result = marked.parse(text, { breaks: true, gfm: true });
                if (typeof result === 'string') return result;
                // 某些版本 marked.parse 返回 Promise，同步场景无法使用
                if (result && typeof result.then === 'function') {
                    ChatLog.warn?.('RENDERER', '[chat-renderer] marked.parse 返回 Promise，使用降级渲染');
                }
            }
        } catch (e) {
            ChatLog.warn?.('RENDERER', '[chat-renderer] marked.parse 失败，使用降级渲染', e);
        }
        // 极简降级：至少处理 ``` 代码块
        return renderMarkdownFallback(text);
    }

    function renderMarkdownFallback(text) {
        // 转义普通文本中的 HTML
        let html = escapeHtml(text);
        // 处理 ```language\ncode\n``` 代码块
        html = html.replace(/```(\w*)\n([\s\S]*?)\n```/g, (match, lang, code) => {
            return `<pre class="code-block${lang ? ' language-' + lang : ''}"><code>${code}</code></pre>`;
        });
        // 处理行内 `code`
        html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
        // 段落化
        html = html.split(/\n{2,}/).map(p => `<p>${p.replace(/\n/g, '<br>')}</p>`).join('');
        return html;
    }

    /**
     * 渲染一段可能混合了纯文本 Markdown 与内嵌 ECharts 节点（data-echarts）的内容。
     * marked 处理整段 Markdown 时可能丢弃/破坏 <div data-echarts> 节点，因此先摘出
     * 这些节点用占位符顶替，其余照常 renderMarkdown，最后再把节点原样插回。
     * 这样即使图表 div 前面有总结文字、无法走 isHtmlContent 原生 HTML 分支也能保留。
     */
    function renderTextPreservingCharts(text) {
        if (!text || !String(text).includes('data-echarts')) return renderMarkdown(text);
        const blocks = [];
        const protectedText = String(text).replace(/(<div\s+data-echarts='[^']*'[^>]*>\s*<\/div>)/g, function(m) {
            blocks.push(m);
            return '\uE000CHART' + (blocks.length - 1) + '\uE000';
        });
        let html = renderMarkdown(protectedText);
        if (blocks.length) {
            html = html.replace(/\uE000CHART(\d+)\uE000/g, function(_, i) {
                return blocks[Number(i)];
            });
        }
        return html;
    }

    /**
     * 清洗 LLM 总结回复里"不小心抄进来的裸 JSON / CSS 字典片段"，也就是用户说的"代码溢出"。
     * 只在纯文本 / Markdown 渲染路径执行，显式 ```code``` 代码块不受影响。
     *   1. 裸贴 CSS 字典: ,"border":{"top":"1px solid black",...}  / {"fill":{"color":"#F2F2F2"},...}
     *   2. 裸露 style 对象片段、从 tool result 截断后残留的尾部 JSON
     *   3. 行尾孤立的逗号 JSON 尾巴
     */
    // 图片来源可信判定（供 3.6 幻觉图片清洗使用）
    // 仅放行：① 本地截屏专用主机 https://screenassets.local/；② 数据 URI 且 base64 真头正确
    // （PNG/JPEG/GIF/WebP）。普通路径 / URL / 畸形 data URI 一律不可信。
    function _isImageMagicOk(b64) {
        return /^(iVBORw0KGgo|\/9j\/|R0lGOD|UklGR)/i.test(String(b64 || '').trim());
    }
    function _isTrustedImageSrc(src) {
        const u = String(src || '').trim();
        if (/^https:\/\/screenassets\.local\//i.test(u)) return true;
        const m = u.match(/^data:image\/[a-z0-9.+-]+;base64,(.+)$/i);
        if (m) return _isImageMagicOk(m[1]);
        return false;
    }

    // ── 3.7 工具结果 JSON 回显守卫的实现 ─────────────────────────────
    // 结果特征键：出现任一即认为"这个 JSON 块是给机器看的工具结果"，不是给人看的内容。
    const _ECHO_RESULT_KEY_RE = /"(?:sourceSheets|splitSheets|mergedRows|dataRows|writtenTo|targetSheet|outputSheet|resultSheet|cleanedEmptySheets|previewRows|files_created|headerPreserved|usedSelectedSheets|usedSelectedSheets|groups|createdSheets)"\s*:/;
    const _ECHO_SUCCESS_RE = /"success"\s*:\s*(true|false|"true"|"false")/i;

    /**
     * 从回显 JSON 块（容忍无逗号/无缩进的复述变体，不做 JSON.parse）提取确定性摘要。
     * 按已知工具族给出可读文案；认不出细节时给通用一行，绝不把 JSON 原文留在正文。
     */
    function _friendlyEchoText(block) {
        const grab = (re) => { const m = block.match(re); return m ? m[1] : undefined; };
        const num = (re) => { const v = grab(re); const n = Number(v); return (v !== undefined && isFinite(n)) ? n : 0; };
        // 拆分族：splitSheets:[{name,rows}]（幻觉形态）或 groups/created（真实后端形态）
        const names = [];
        const nameRe = /"name"\s*:\s*"([^"]{1,40})"/g;
        let nm;
        while ((nm = nameRe.exec(block)) !== null && names.length < 8) names.push(nm[1]);
        const createdRe = /"created"\s*:\s*\[([^\]]*)\]/;
        if (!names.length && createdRe.test(block)) {
            const cm = block.match(createdRe);
            if (cm) {
                const cn = cm[1].match(/"([^"]{1,60})"/g) || [];
                cn.slice(0, 8).forEach(x => names.push(x.replace(/"/g, '')));
            }
        }
        const groups = num(/"(?:groups|splitCount|sheetCount)"\s*:\s*(\d+)/);
        if (names.length || groups) {
            const list = names.length ? '：' + names.join('、') + (names.length >= 8 ? ' 等' : '') : '';
            return '✅ 已按列拆分为 ' + (groups || names.length) + ' 个工作表' + list;
        }
        // 合并族：sourceSheets/mergedRows/targetSheet
        const mergedRows = num(/"(?:mergedRows|dataRows)"\s*:\s*(\d+)/);
        const target = grab(/"(?:targetSheet|writtenTo|outputSheet)"\s*:\s*"([^"]{1,40})"/);
        const srcCount = (block.match(/"sourceSheets"/) ? (block.match(/"([^"]{1,20})"\s*(?:,|\n|\])/g) || []).length : 0) || num(/"mergedSheets"\s*:\s*(\d+)/);
        if (mergedRows || /"sourceSheets"/.test(block)) {
            return '✅ 已合并工作表' + (mergedRows ? '，共 ' + mergedRows + ' 行数据' : '')
                + (target ? ' → 「' + target + '」' : '');
        }
        return '✅ 操作已完成（结果详情见上方工具卡片）';
    }

    /**
     * 扫描正文中的平衡 {...} 块（容忍换行/无逗号），命中"工具结果特征"整块替换为摘要。
     * 只在渲染层执行，不改 turn.content。代码围栏（```）内容先摘出保护，扫描后原样放回。
     */
    function _replaceToolResultEchoBlocks(text) {
        if (!text || text.indexOf('"success"') < 0) return text;
        let s = text;
        // 保护代码围栏：```...```（含未闭合尾围栏）不参与扫描
        const fences = [];
        s = s.replace(/```[\s\S]*?(?:```|$)/g, function (m) {
            fences.push(m);
            return '\uE001FENCE' + (fences.length - 1) + '\uE001';
        });
        let out = '';
        let i = 0;
        while (i < s.length) {
            const start = s.indexOf('{', i);
            if (start < 0) { out += s.slice(i); break; }
            // 找配对右括号（字符串感知，容忍转义）
            let depth = 0, j = start, inStr = false, closed = false;
            for (; j < s.length; j++) {
                const ch = s.charAt(j);
                if (inStr) {
                    if (ch === '\\') { j++; }
                    else if (ch === '"') { inStr = false; }
                    continue;
                }
                if (ch === '"') inStr = true;
                else if (ch === '{') depth++;
                else if (ch === '}') { depth--; if (depth === 0) { closed = true; break; } }
            }
            if (!closed) {
                // 【截断 JSON 尾巴 2026-10-10】safeDisplayText 等上游会把 JSON 截到 1000 字符，
                // 括号不配对。旧逻辑原样放行 → 裸 JSON 尾巴进正文（实测 23:28 合并）。
                // 这里对"未闭合但带工具结果特征"的尾巴也做摘要替换，从 { 起整段替换。
                const rest = s.slice(start);
                if (rest.length > 40 && _ECHO_SUCCESS_RE.test(rest) && _ECHO_RESULT_KEY_RE.test(rest)) {
                    out += s.slice(i, start) + _friendlyEchoText(rest);
                } else {
                    out += s.slice(i);
                }
                break;
            }
            const block = s.slice(start, j + 1);
            if (block.length > 40 && _ECHO_SUCCESS_RE.test(block) && _ECHO_RESULT_KEY_RE.test(block)) {
                out += s.slice(i, start) + _friendlyEchoText(block);
            } else {
                out += s.slice(i, j + 1);
            }
            i = j + 1;
        }
        s = out.replace(/\uE001FENCE(\d+)\uE001/g, function (_, k) { return fences[Number(k)]; });
        return s;
    }

    function sanitizeLeakedJsonSnippets(text) {
        if (!text || typeof text !== 'string') return text;
        let s = text;

        // 1) 最典型泄漏：CSS 字典尾巴（开头常常带多余逗号）—— border/fill/font/alignment/...
        //    注意用非贪婪 + 计数括号兜底，宁可多匹配一点尾巴，也别把正文吞掉
        s = s.replace(/,?\s*\{?\s*"border"\s*:\s*\{[\s\S]*?\}(?:\s*,\s*"[^"]+"\s*:\s*(?:"[^"]*"|\{[^}]*\}|\[[^\]]*\]|true|false|null|-?\d+(?:\.\d+)?))*\s*\}?\s*$/g, '');

        // 2) 更宽泛：行尾一段完整的 {...}（深度>1 且包含明显"内部字典键"border/fill/font/alignment/color）
        //    只清尾巴，不清整段（避免误伤真正的 JSON 输出）
        s = s.replace(/,\s*\{(?:[^{}]*"border"[^{}]*|[^{}]*"fill"[^{}]*|[^{}]*"font"[^{}]*|[^{}]*"alignment"[^{}]*|[\s\S]){10,}\s*$/g, m => {
            // 简单结构判断：里面若还有花括号嵌套，则当"裸贴尾巴"删掉
            return (m.includes('{') && (m.match(/\{/g) || []).length >= 2) ? '' : m;
        });

        // 3) 结尾孤立的逗号尾巴（前面的 JSON 切了一半留下的）
        s = s.replace(/[,，]\s*$/gm, '');

        // 3.5) 工具调用参数泄漏：模型把 excel_* 工具的 JSON 参数以 ```json 围栏回显在正文里
        //      （如 {"range":"$A$1:$I$101","returnFormat":...} / {"dimColumn":"部门","valueColumn":"销售额"}）。
        //      特征：围栏内容含 ExcelAI 工具参数特征键；真实给用户看的 JSON 示例不会带这些键，整块移除。
        //      流式与收尾共用本函数 → 参数卡不再"流式期可见、收尾消失"。
        const _toolArgKeyRe = /"(?:dimColumns?|valueColumns?|aggFunc|returnFormat|anchorRange|anchorCell|chartType|groupColumn|seriesColumn|targetSheet|numberFormat|sheet|range)"\s*:/;
        s = s.replace(/```json\s*\{[\s\S]*?\}\s*```/gi, m => _toolArgKeyRe.test(m) ? '' : m);
        // 流式中尚未闭合的尾部 ```json 围栏同样按特征键清除，避免参数卡在流式期闪现
        s = s.replace(/```json\s*\{[\s\S]*$/gi, m => _toolArgKeyRe.test(m) ? '' : m);

        // 3.7) 【工具结果 JSON 回显/编造守卫 2026-10-09】模型会把"回喂过/历史里出现过的工具结果 JSON"
        //      原样（几乎必然丢逗号、拆行）复述或凭空编造进气泡正文（实测：合并→sourceSheets 矩阵、
        //      拆分→splitSheets 数组，后者字段甚至后端根本不存在，纯幻觉）。逐工具打补丁追不上，
        //      这里按【结构特征】通用拦截：一个 {...} 块同时含 "success": 和结果特征键
        //      （标准 JSON 与无逗号变体都认，靠键名正则而非 JSON.parse）→ 整块替换为确定性摘要。
        //      ``` 代码围栏内的 JSON 不动（那是用户显式要看的输出）。
        s = _replaceToolResultEchoBlocks(s);

        // 3.6) 幻觉内嵌图片：模型在没有真实图像产物时，常编造 Markdown 图片
        //      （如 ![美化后的表格](假路径) 或 ![x](data:image/png;base64,IVBORw0KGgo…)），
        //      渲染后变成浏览器"破图"占位符 + alt 文本。用【白名单】彻底处理：
        //      仅放行可信来源（数据 URI 且 base64 真头正确，或本地截屏主机 screenassets.local），
        //      其余一律退化为纯文本（保留 alt 文案），杜绝破图。
        //      ── 情形 1：Markdown 图片 ![alt](src) ──
        s = s.replace(/!\[([^\]\n]{0,120})\]\(\s*([^)\s]{0,4000})\s*\)/gi, function (_m, alt, src) {
            if (_isTrustedImageSrc(src)) return _m;
            const a = String(alt || '').trim();
            return a || '（图片已移除）';
        });
        // ── 情形 2：裸 <img> 标签（marked 透传 HTML）──
        s = s.replace(/<img\b[^>]*>/gi, function (_m) {
            const srcM = _m.match(/\bsrc\s*=\s*["']([^"']*)["']/i);
            if (srcM && _isTrustedImageSrc(srcM[1])) return _m;
            const altM = _m.match(/\balt\s*=\s*["']([^"']*)["']/i) || _m.match(/\btitle\s*=\s*["']([^"']*)["']/i);
            return (altM && altM[1] && altM[1].trim()) || '（图片已移除）';
        });
        // ── 情形 3：裸 data:image base64（无 ![] 包裹，主要防 base64 失控）──
        s = s.replace(/data:[a-z]+\/[a-z0-9.+-]+;base64,[A-Za-z0-9+\/=]{64,}/gi, '（内嵌图片数据已移除）');
        // 兜底：无 data: 前缀的超长裸 base64 串（幻觉变体）
        s = s.replace(/[A-Za-z0-9+\/=]{512,}/g, '（内嵌图片数据已移除）');

        // 4) 多余空白
        s = s.replace(/[ \t]{2,}/g, ' ').replace(/\n{3,}/g, '\n\n').trimEnd();
        return s;
    }

    /**
     * 移除 LLM 总结里残留的"仿造工具标签"（纯展示层）。
     * 真实执行的工具调用已由 agent-loop 剥除并单独渲染成工具卡片，
     * 若这里还出现 [excel-operation__xxx] / [excel-local__xxx] / [excel_xxx] / [excel-xxx]
     * 之类括号，说明模型在总结里杜撰了一个并未以该名执行的工具，属虚报，渲染时不应展示。
     * 只在渲染时清理，不改动 turn.content / agent 上下文。
     */
    function sanitizeFakeToolTags(text) {
        if (!text || typeof text !== 'string') return text;
        let s = text;
        // 1) 类工具 bracket：[excel-operation__clear_format] / [excel-local__excel_clear_range]{...} /
        //    [excel-local_excel_read_range] / [excel_clear_range $A$1:$H$6]（含空格分隔的位置参数形式）
        s = s.replace(/\[(?:excel[-_]local?_{1,2}|excel[-_])[a-zA-Z0-9_-]+[^\]\r\n]*\](?:[ \t]*\{[^{}\n]{0,400}\})?/g, '');
        // 2) [EXECUTE:SkillName]
        s = s.replace(/\[EXECUTE:[\w.\s-]+\]/g, '');
        // 3) 若"调用工具："等引导语因上方标签被清而孤立，一并去掉
        s = s.replace(/\s*调用工具\s*[：:]/g, '');
        // 4) 清除 LLM 总结中可能残留的"工具名: 结果"前缀（如"excel_create_sheet: 成功创建工作表"）
        //    prompt 已禁止 LLM 输出工具名，但某些模型仍可能输出，渲染层兜底清除
        //    使用 ^\s* 兼容流式首帧可能带换行的情况
        s = s.replace(/^\s*excel_[a-zA-Z0-9_]+:\s*/g, '');
        // 5) 函数样式假调用前缀（"excel_split_by_column(column="部门"…) → 成功: "，实测 2026-10-09 23:05）：
        //    模型冒充已执行时把调用写成函数样式。agent-loop 已把可解析的转成真实调用，
        //    这里清洗渲染层残留（只清带箭头→成功声明的强特征形态，普通散文提及不动）。
        s = s.replace(/\b(?:excel|ppt|word|computer)_[a-z0-9_]{2,40}\s*\([^()]{0,300}\)\s*(?:→|->|=>)\s*(?:成功|完成|done)?\s*[:：]?\s*/gi, '');
        return s.replace(/\n{3,}/g, '\n\n').trim();
    }

    // ── 空表格 / 占位表格清洗（2026-10-16）─────────────────────────
    // 实测：用户说"按销售额从高到低排序"，qwen2.5 执行 excel_sort_range 后总结时
    // "承诺画表格但填不出数据"——输出只有表头一行 + 数据行全是空白/"$"的占位符
    // （| 地区 | 部门 | 产品 | 销售额 | + |  |  |  | $ |）。渲染层照单全收 → 用户看到
    // 一张空的"表格卡片"，点"插入到 Excel"也只会插入一张空表，比没有表格更误导。
    // 这里在渲染前把这类假表格清掉：
    //   1) 逐行判定：整行所有单元格都是占位符（空白 / $ / - / … / 待填充 / N/A …）→ 删行；
    //   2) 删完一条不剩（表头+分隔行孤存）→ 整张表删除，只留一行中性提示；
    //   3) 代码围栏（VBA/Python 里大量竖线）先摘出保护，扫描后原样放回。
    // 占位符判定只认"明显不是数据"的形态；真实表格（含部分空单元格）不受影响。
    const _PLACEHOLDER_CELL_RE = /^(?:[$¥￥€£\-\u2013\u2014_*.…、/\\]{1,8}|n\/a|null|undefined|tbd|todo|待填充|待补充|待定|待填|略|无数据|\(空\)|\(无\)|\(略\))$/i;
    const _TABLE_SEPARATOR_RE = /^\s*\|?\s*:?-{1,}:?\s*(?:\|\s*:?-{1,}:?\s*)*\|?\s*$/;
    const _EMPTY_TABLE_NOTE = '（未生成有效的表格数据）';

    function _isPlaceholderCell(cell) {
        const s = String(cell == null ? '' : cell).trim();
        if (s === '') return true;
        return _PLACEHOLDER_CELL_RE.test(s);
    }

    function _splitTableRow(line) {
        let s = String(line).trim();
        if (s.charAt(0) === '|') s = s.slice(1);
        if (s.charAt(s.length - 1) === '|') s = s.slice(0, -1);
        // 不拆转义竖线（\| 是单元格内容，不是列分隔）
        return s.split(/(?<!\\)\|/).map(c => c.replace(/\\\|/g, '|').trim());
    }

    /** 清洗一个表块（lines[0]=表头，lines[1]=分隔行，其余=数据行）。
     *  返回保留的行数组；整张表没有真实数据行时返回 null（调用方整表删除）。 */
    function _sanitizeTableBlock(lines) {
        const keptBody = lines.slice(2).filter(line => {
            // 只要还有一个非占位单元格，就认为是真实数据行
            return _splitTableRow(line).some(c => !_isPlaceholderCell(c));
        });
        if (keptBody.length === 0) return null;
        return lines.slice(0, 2).concat(keptBody);
    }

    function sanitizePlaceholderTables(text) {
        if (!text || typeof text !== 'string') return text;
        if (text.indexOf('|') < 0) return text;
        let s = text;
        // 保护代码围栏：```...```（含未闭合尾围栏）不参与表格扫描
        const fences = [];
        s = s.replace(/```[\s\S]*?(?:```|$)/g, function (m) {
            fences.push(m);
            return '\uE002FENCE' + (fences.length - 1) + '\uE002';
        });
        const lines = s.split('\n');
        const out = [];
        let i = 0;
        let dropped = 0;
        let notePending = false; // 已输出过提示且其后没有实体内容 → 不重复刷屏
        const pushLine = (line) => {
            if (line.trim() !== '') notePending = false;
            out.push(line);
        };
        while (i < lines.length) {
            const next = lines[i + 1];
            // 表块起点：当前行含竖线，且下一行是 GFM 分隔行（|---|:--:|）
            if (lines[i].indexOf('|') >= 0 && next != null && _TABLE_SEPARATOR_RE.test(next)) {
                let j = i + 2;
                while (j < lines.length && lines[j].indexOf('|') >= 0 && lines[j].trim() !== '') j++;
                const cleaned = _sanitizeTableBlock(lines.slice(i, j));
                if (cleaned) {
                    for (const l of cleaned) pushLine(l);
                } else {
                    dropped++;
                    // 连续删除多张空表时只留一条提示，避免刷屏
                    if (!notePending) {
                        out.push(_EMPTY_TABLE_NOTE);
                        notePending = true;
                    }
                }
                i = j;
                continue;
            }
            pushLine(lines[i]);
            i++;
        }
        s = out.join('\n');
        s = s.replace(/\uE002FENCE(\d+)\uE002/g, function (_, k) { return fences[Number(k)]; });
        if (dropped > 0) {
            ChatLog.info?.('RENDERER', `[chat-renderer] 已清除 ${dropped} 张无数据占位表格（模型未填充真实数据）`);
        }
        return s;
    }

    function isCodeTurn(turn) {
        return !!(turn && turn._codeLang);
    }

    App.ChatV2.renderers.markdown = {
        renderMarkdown,
        renderMarkdownFallback,
        renderTextPreservingCharts,
        sanitizeLeakedJsonSnippets,
        sanitizeFakeToolTags,
        sanitizePlaceholderTables,
        isCodeTurn
    };
    ChatLog.info?.('CHAT-V2', 'renderers/markdown loaded');
})();