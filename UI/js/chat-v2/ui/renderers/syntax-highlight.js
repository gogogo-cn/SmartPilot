/*
 * chat-v2/ui/renderers/syntax-highlight.js - Syntax-highlighting pure-function submodule
 * Extracted from chat-renderer.js (Token-based, avoids regex nesting).
 * All functions are pure string processing (DOM-free), unit-testable with node.
 * Self-contained escapeHtml delegating to App.ChatV2.renderHtml.escapeHtml.
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
    function highlightSyntax(code, lang) {
        if (!code) return null;
        if (code.length > 8000) return null;

        switch (lang) {
            case 'vba': case 'vb': return highlightVBA(code);
            case 'python': case 'py': return highlightPython(code);
            case 'javascript': case 'js': case 'typescript': case 'ts': return highlightJS(code);
            case 'json': return highlightJSON(code);
            case 'html': return highlightHTML(code);
            case 'css': return highlightCSS(code);
            case 'sql': return highlightSQL(code);
            case 'bash': case 'shell': return highlightBash(code);
            case 'powershell': return highlightPowerShell(code);
            case 'csharp': case 'cs': return highlightCSharp(code);
            case 'java': return highlightJava(code);
            default: return null;
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

    function highlightJSON(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /"([^"\\]|\\.)*"\s*:/g, type: 'key' },
                { regex: /:\s*"([^"\\]|\\.)*"/g, type: 'string' }
            ],
            highlight: [
                { regex: /:\s*(\d+(\.\d+)?)/g, replacer: ': <span class="tok-number">$1</span>' },
                { regex: /:\s*(true|false|null)/g, replacer: ': <span class="tok-bool">$1</span>' }
            ]
        });
    }

    function highlightHTML(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /<!--[\s\S]*?-->/g, type: 'comment' },
                { regex: /<[^>]*>/g, type: 'tag' }
            ],
            highlight: []
        });
    }

    function highlightCSS(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /\/\*[\s\S]*?\*\//g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /'([^'\\]|\\.)*'/g, type: 'string' }
            ],
            highlight: [
                { regex: /([.#][\w-]+)\s*\{/g, replacer: '<span class="tok-tag">$1</span> {' },
                { regex: /([\w-]+)\s*:\s*([^;]+)/g, replacer: '<span class="tok-property">$1</span>: <span class="tok-value">$2</span>' }
            ]
        });
    }

    function highlightSQL(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /--[^\n]*/g, type: 'comment' },
                { regex: /\/\*[\s\S]*?\*\//g, type: 'comment' },
                { regex: /'[^']*'/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+(\.\d+)?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(SELECT|FROM|WHERE|INSERT|INTO|VALUES|UPDATE|SET|DELETE|CREATE|TABLE|DROP|ALTER|INDEX|JOIN|INNER|LEFT|RIGHT|OUTER|CROSS|ON|GROUP|BY|ORDER|HAVING|LIMIT|OFFSET|AND|OR|NOT|IN|EXISTS|BETWEEN|LIKE|ILIKE|AS|UNION|ALL|DISTINCT|PRIMARY|KEY|FOREIGN|REFERENCES|DEFAULT|CHECK|NULL|IS|ASC|DESC|INTEGER|VARCHAR|TEXT|DATE|BOOLEAN|FLOAT|DECIMAL|INT|CHAR|DATETIME|TIMESTAMP|CASE|WHEN|THEN|ELSE|END|BEGIN|TRANSACTION|COMMIT|ROLLBACK|REPLACE|REGEXP|SOUNDEX|CONTAINS)\b/gi, replacer: '<span class="tok-keyword">$1</span>' }
            ]
        });
    }

    function highlightBash(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /#[^\n]*/g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /'([^'\\]|\\.)*'/g, type: 'string' },
                { regex: /\$\([^)]*\)/g, type: 'string' },
                { regex: /`[^`]*`/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(echo|cd|ls|mkdir|rm|cp|mv|cat|grep|sed|awk|if|then|else|elif|fi|for|do|done|while|case|esac|function|return|export|import|source|alias|unset|read|local|declare|eval|exec|exit|break|continue|in|not|and|or)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b([a-zA-Z_][\w-]*)\s*(?=\()/g, replacer: '<span class="tok-function">$1</span>' }
            ]
        });
    }

    function highlightPowerShell(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /#[^\n]*/g, type: 'comment' },
                { regex: /<#[\s\S]*?#>/g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /'([^'\\]|\\.)*'/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(if|else|elseif|for|foreach|while|do|switch|case|default|break|continue|return|exit|function|param|begin|process|end|try|catch|finally|throw|trap|in|not|and|or|xor|true|false|null)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\$-?[a-zA-Z_][\w]*/g, replacer: '<span class="tok-variable">$&</span>' },
                { regex: /\b(Get-[a-zA-Z]+|Set-[a-zA-Z]+|New-[a-zA-Z]+|Remove-[a-zA-Z]+|Invoke-[a-zA-Z]+|Import-[a-zA-Z]+|Export-[a-zA-Z]+)\b/g, replacer: '<span class="tok-function">$1</span>' }
            ]
        });
    }

    function highlightCSharp(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /\/\/[^\n]*/g, type: 'comment' },
                { regex: /\/\*[\s\S]*?\*\//g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' },
                { regex: /@"([^"\\]|\\.)*"/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+(\.\d+)?[fld]?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(using|namespace|class|interface|struct|enum|public|private|protected|internal|static|readonly|void|var|new|return|if|else|for|foreach|while|switch|case|break|continue|try|catch|finally|throw|async|await|public|private|get|set|this|base|null|true|false|in|is|as|typeof|default|out|ref|params)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b([a-zA-Z_][\w]*)\s*\(/g, replacer: '<span class="tok-function">$1</span>(' }
            ]
        });
    }

    function highlightJava(code) {
        return protectAndHighlight(code, {
            protect: [
                { regex: /\/\/[^\n]*/g, type: 'comment' },
                { regex: /\/\*[\s\S]*?\*\//g, type: 'comment' },
                { regex: /"([^"\\]|\\.)*"/g, type: 'string' }
            ],
            highlight: [
                { regex: /\b(\d+[lLfFdD]?)\b/g, replacer: '<span class="tok-number">$1</span>' },
                { regex: /\b(public|private|protected|class|interface|extends|implements|void|static|final|new|return|if|else|for|while|do|switch|case|break|continue|try|catch|finally|throw|throws|import|package|this|super|null|true|false|in|instanceof|void|int|long|short|byte|float|double|char|boolean|String|System|out|println)\b/g, replacer: '<span class="tok-keyword">$1</span>' },
                { regex: /\b([a-zA-Z_][\w]*)\s*\(/g, replacer: '<span class="tok-function">$1</span>(' }
            ]
        });
    }

    App.ChatV2.renderers.syntaxHighlighter = {
        highlightSyntax,
        protectAndHighlight,
        applyHighlight,
        highlightVBA,
        highlightPython,
        highlightJS,
        highlightJSON,
        highlightHTML,
        highlightCSS,
        highlightSQL,
        highlightBash,
        highlightPowerShell,
        highlightCSharp,
        highlightJava
    };
    ChatLog.info?.('CHAT-V2', 'renderers/syntax-highlight loaded');
})();