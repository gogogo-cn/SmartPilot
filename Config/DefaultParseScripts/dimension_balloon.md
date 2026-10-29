---
id: dimension_balloon
name: 尺寸标注（球标）
description: 解析带球标的尺寸。圆球标为非重要，六边形球标为重要，并拆出序号、类型、标称值、下公差、上公差。
icon: 🎯
category: 工程图
outputFormat: table
tags: 尺寸, 球标, 形位公差, 序号, 重要性
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 尺寸序号 | seq | - | 否 | 是 |
| 重要性 | importance | - | 否 | 是 |
| 类型 | dim_type | - | 否 | 是 |
| 标称值 | nominal | - | 否 | 是 |
| 下公差 | lower_tol | - | 否 | 是 |
| 上公差 | upper_tol | - | 否 | 是 |
| 标注条数 | count | - | 否 | 是 |
| 尺寸明细 | items | - | 否 | 是 |

## 解析脚本

```javascript
/**
 * 机械图纸OCR文本解析：提取球标、尺寸、公差、粗糙度、形位公差
 * 序号支持：①~㊿(1-50) ⓐ~ⓩ(51-76) ⓪~⑳(77-100)
 * @param {string} rawText OCR原始文本（可带<br/> html）
 * @param {Array} [fields] 字段配置，同原FIELDS结构，可选
 * @returns {Object} { items: [], tableData: {}, count, 匹配到的自定义字段... }
 */
function parse(rawText, fields) {
    // 默认字段配置（替代全局FIELDS，消除依赖报错）
    const DEFAULT_FIELDS = [
        { key: 'seq', name: '尺寸序号', pattern: '-', visible: true },
        { key: 'importance', name: '重要性', pattern: '-', visible: true },
        { key: 'dim_type', name: '类型', pattern: '-', visible: true },
        { key: 'nominal', name: '标称值', pattern: '-', visible: true },
        { key: 'lower_tol', name: '下公差', pattern: '-', visible: true },
        { key: 'upper_tol', name: '上公差', pattern: '-', visible: true }
    ];
    // 优先用入参；否则用宿主注入的 FIELDS；再否则默认配置（勿再声明名为 FIELDS 的局部变量，会遮蔽宿主注入并触发 TDZ）
    const injectedFields = (typeof FIELDS !== 'undefined' && Array.isArray(FIELDS)) ? FIELDS : null;
    const fieldList = Array.isArray(fields) ? fields : (injectedFields || DEFAULT_FIELDS);

    let text = String(rawText || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/[Φφ⌀]/g, 'Ø')
        .replace(/＋/g, '+')
        .replace(/－/g, '-')
        .replace(/[°º]/g, '°')
        // 全角数字转半角
        .replace(/[０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
        // 全角字母转半角
        .replace(/[Ａ-Ｚａ-ｚ]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
        .replace(/R\s*α/gi, 'Ra')
        .replace(/R\s*([azqy])/gi, (_, c) => 'R' + c)
        .replace(/[ \t]+/g, ' ')
        .trim();

    // ========== 序号集：1~100 ==========
    const circled = '①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳㉑㉒㉓㉔㉕㉖㉗㉘㉙㉚㉛㉜㉝㉞㉟㊱㊲㊳㊴㊵㊶㊷㊸㊹㊺㊻㊼㊽㊾㊿';
    const circledAlphaLower = 'ⓐⓑⓒⓓⓔⓕⓖⓗⓘⓙⓚⓛⓜⓝⓞⓟⓠⓡⓢⓣⓤⓥⓦⓧⓨⓩ';
    const circledDouble = '⓪①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳';
    const filled = '❶❷❸❹❺❻❼❽❾❿';
    const parenNum = '⑴⑵⑶⑷⑸⑹⑺⑻⑼⑽⑾⑿⒀⒁⒂⒃⒄⒅⒆⒇';

    // ========== 预编译正则（提升大批量性能） ==========
    const RE_DEV = /^[+-]\d+(?:\.\d+)?$|^0$|^±\s*\d+(?:\.\d+)?$/;
    const RE_ROUGH_HEADER = /^(?:(?:表面)?粗糙度\s*[:：]?|R\s*[azqyα]|[√∨˅✓✔▽▼])\s*$/i;
    const RE_ROUGH_TOKEN = /(?:(?:表面)?粗糙度|[√∨˅✓✔▽▼]|Ra|Rz|Ry|Rq)\s*[:：=≤≦]?\s*\d+(?:\.\d+)?(?:\s*[\/~～]\s*\d+(?:\.\d+)?)?/gi;
    const RE_GDT_SYM_LINE = /([⊕⌖◎∥⊥▱⌒⌓∠↗])\s*(\d+(?:\.\d+)?)(?:\s+([A-Z](?:\s+[A-Z-])*))?/g;
    const RE_CIRCLE_ROUND = /([○])\s*(\d+\.\d+)(?:\s+([A-Z](?:\s+[A-Z]){0,2}))?/g;
    const RE_N_TIMES = /\d+\s*[×xX]\s*(M\d+(?:\.\d+)?(?:\s*[×xX]\s*\d+(?:\.\d+)?)(?:-\d+[A-Za-z])?|Ø\s*\d+(?:\.\d+)?|R\s*\d+(?:\.\d+)?|C\s*\d+(?:\.\d+)?)/g;
    const RE_METRIC_THREAD = /M\d+(?:\.\d+)?(?:\s*[×xX]\s*\d+(?:\.\d+)?)(?:-\d+[A-Za-z])?/g;
    // 直径：裸 Ø18；±；空格双偏差；斜杠 0/-0.025、+0.012/-0.006
    const RE_DIAM = /Ø\s*\d+(?:\.\d+)?(?:\s*±\s*\d+(?:\.\d+)?|\s*[+-]?\d+(?:\.\d+)?\s*\/\s*[+-]?\d+(?:\.\d+)?|\s*[+-]\d+(?:\.\d+)?(?:\s*(?:[+-]\d+(?:\.\d+)?|0))?)?/g;
    const RE_RADIUS = /(?:^|\s)R\s*\d+(?:\.\d+)?/g;
    const RE_CHAMFER = /(?:^|\s)C\s*\d+(?:\.\d+)?/g;
    const RE_ANGLE = /\d+(?:\.\d+)?\s*°/g;
    // 配合：16H7/f6、H7/f6、以及单侧 32h7 / 28H7 / 18f7
    const RE_FIT = /\d+(?:\.\d+)?\s*[A-Z]\d+\s*\/\s*[a-z]\d+|[A-Z]\d+\s*\/\s*[a-z]\d+|\d+(?:\.\d+)?\s*[A-Za-z]\d+/g;
    const RE_TOL = /\d+(?:\.\d+)?\s*±\s*\d+(?:\.\d+)?|\d+(?:\.\d+)?\s+[+-]?\d+(?:\.\d+)?\s*\/\s*[+-]?\d+(?:\.\d+)?|\d+(?:\.\d+)?\s+[+-]\d+(?:\.\d+)?\s*[+-]\d+(?:\.\d+)?|\d+(?:\.\d+)?\s+[+-]\d+(?:\.\d+)?/g;
    const RE_SKIP = /^(重量|材料|设计|校对|审核|工艺|批准|比例|零件|姓名|日期|单位|阶段|共\s*\d|第\s*\d|未注(?:尺寸|形位)?公差)/;
    const RE_ISO_TOL = /GB\/T\s*1804|GB\/T\s*1184/;
    const RE_DATE = /^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}/;

    const gdtNames = ['直线度', '平面度', '圆度', '圆柱度', '线轮廓度', '面轮廓度', '平行度', '垂直度', '倾斜度', '位置度', '同轴度', '同心度', '对称度', '圆跳动', '全跳动'];
    const gdtSymbols = {
        '⊕': '位置度', '⌖': '位置度', '◎': '同轴度', '∥': '平行度', '⊥': '垂直度',
        '▱': '平面度', '⌒': '线轮廓度', '⌓': '面轮廓度', '∠': '倾斜度', '↗': '圆跳动', '○': '圆度'
    };

    function circledSeq(ch) {
        const sets = [circled, filled, parenNum];
        for (const s of sets) {
            const at = s.indexOf(ch);
            if (at >= 0) return String(at + 1);
        }
        const aIdx = circledAlphaLower.indexOf(ch);
        if (aIdx >= 0) {
            return String(circled.length + aIdx + 1);
        }
        const dIdx = circledDouble.indexOf(ch);
        if (dIdx >= 0) {
            return String(circled.length + circledAlphaLower.length + dIdx + 1);
        }
        return '';
    }

    function readBalloon(line) {
        if (!line) return null;
        let seq = circledSeq(line.charAt(0));
        if (seq) return { seq, importance: '非重要', rest: line.slice(1).trim() };

        let m = line.match(/^[⬡⬢⎔]\s*(\d+)\s*/);
        if (m) return { seq: m[1], importance: '重要', rest: line.slice(m[0].length).trim() };

        m = line.match(/^(?:六边形|六角)\s*[:：]?\s*(\d+)\s*/);
        if (m) return { seq: m[1], importance: '重要', rest: line.slice(m[0].length).trim() };

        m = line.match(/^(?:圆形|圆球标|圆标)\s*[:：]?\s*(\d+)\s*/);
        if (m) return { seq: m[1], importance: '非重要', rest: line.slice(m[0].length).trim() };

        m = line.match(/^[（(]\s*(\d+)\s*[）)]\s*/);
        if (m) return { seq: m[1], importance: '非重要', rest: line.slice(m[0].length).trim() };

        // OCR 常把左括号吃掉：52) Ø50 0/-0.025（限 1~3 位，避免误吃年份等数字）
        m = line.match(/^(\d{1,3})\s*[）)]\s*/);
        if (m) return { seq: m[1], importance: '非重要', rest: line.slice(m[0].length).trim() };

        m = line.match(/^[○◯〇●⊙]\s*(\d+)(?![.\d])\s*/);
        if (m) return { seq: m[1], importance: '非重要', rest: line.slice(m[0].length).trim() };
        return null;
    }

    function isDeviation(line) {
        return RE_DEV.test(line);
    }
    function isGdtLine(line) {
        for (const name of gdtNames) {
            if (line.includes(name)) return true;
        }
        return /^[⊕⌖◎∥⊥▱⌒⌓∠↗]/.test(line);
    }
    function isRoughLine(line) {
        return /粗糙度|(?:^|[^A-Za-z])R\s*[azqyα](?![A-Za-z])|[√∨˅✓✔▽▼]/i.test(line);
    }
    function roughKind(token) {
        const t = String(token || '').replace(/\s+/g, '').replace(/α/gi, 'a');
        if (/^ra$/i.test(t)) return 'Ra';
        if (/^rz$/i.test(t)) return 'Rz';
        if (/^ry$/i.test(t)) return 'Ry';
        if (/^rq$/i.test(t)) return 'Rq';
        return '';
    }
    function isRoughHeader(line) {
        return RE_ROUGH_HEADER.test(String(line || '').trim());
    }
    function isBalloonOnly(line) {
        const b = readBalloon(line);
        return !!(b && !b.rest);
    }

    function parseRoughnessToken(s) {
        const str = String(s).trim();
        const roughRe = /(?:(?:表面)?粗糙度\s*[:：=≤≦]?)?([√∨˅✓✔▽▼])?\s*(Ra|Rz|Ry|Rq)?\s*[:：=≤≦]?\s*(\d+(?:\.\d+)?)\s*(?:[\/~～]\s*(\d+(?:\.\d+)?))?/i;
        const m = str.match(roughRe);
        if (!m) return null;
        const p = roughKind(m[2] || 'Ra');
        const v1 = m[3];
        const v2 = m[4] || '';
        if (!v1) return null;
        return { param: p, valMin: v1, valMax: v2 };
    }

    const rawLines = text.split('\n')
        .map(line => line.trim())
        .filter(Boolean);
    const lines = [];
    for (const line of rawLines) {
        if (lines.length) {
            const last = lines[lines.length - 1];
            if (
                isDeviation(line)
                && /\d/.test(last)
                && !isGdtLine(last)
                && !isRoughLine(last)
                && !isBalloonOnly(last)
            ) {
                lines[lines.length - 1] += ' ' + line;
                continue;
            }
        }
        lines.push(line);
    }

    const items = [];
    const seen = new Set();
    const roughUnique = new Set();
    let pendingQueue = []; // 球标队列，支持连续多个孤立球标
    let pendingRough = null;

    function bindKey(aliases, fallback) {
        for (const field of fieldList) {
            if (!field) continue;
            if (aliases.includes(field.key) || aliases.includes(field.name)) {
                return field.key || fallback;
            }
        }
        return fallback;
    }
    const K = {
        seq: bindKey(['seq', '尺寸序号', '序号'], 'seq'),
        importance: bindKey(['importance', '重要性'], 'importance'),
        dim_type: bindKey(['dim_type', '类型'], 'dim_type'),
        nominal: bindKey(['nominal', '标称值'], 'nominal'),
        lower_tol: bindKey(['lower_tol', '下公差'], 'lower_tol'),
        upper_tol: bindKey(['upper_tol', '上公差'], 'upper_tol'),
        items: bindKey(['items', '尺寸明细'], 'items'),
        count: bindKey(['count', '标注条数'], 'count')
    };

    const MAX_ITEMS = 1200;
    const PARSE_TIMEOUT = 800;
    const startTick = Date.now();

    function push(balloon, type, nominal, lower, upper) {
        if (items.length >= MAX_ITEMS) return;
        if (Date.now() - startTick > PARSE_TIMEOUT) return;

        nominal = String(nominal || '').replace(/\s+/g, ' ').trim();
        lower = String(lower || '').trim();
        upper = String(upper || '').trim();
        if (!nominal && !type) return;
        const seq = balloon ? balloon.seq : '';
        const importance = balloon ? balloon.importance : '';

        if (['Ra','Rz','Ry','Rq'].includes(type)) {
            // 同球标同参数同数值才去重，允许同一序号下不同粗糙度值
            const rKey = `${seq}|${type}|${nominal}|${upper}`;
            if (roughUnique.has(rKey)) return;
            roughUnique.add(rKey);
        }

        const key = [seq, importance, type, nominal, lower, upper].join('|').replace(/\s+/g, '');
        if (seen.has(key)) return;
        seen.add(key);

        const row = {};
        row[K.seq] = seq;
        row[K.importance] = importance;
        row[K.dim_type] = type;
        row[K.nominal] = nominal;
        row[K.lower_tol] = lower;
        row[K.upper_tol] = upper;
        items.push(row);
    }

    function fmtDev(v, side) {
        const t = String(v || '').trim();
        if (t === '0' || t === '+0' || t === '-0') return '0';
        if (/^[+-]/.test(t)) return t;
        return (side === 'upper' ? '+' : '-') + t;
    }

    function splitSize(raw) {
        const s = String(raw || '').replace(/\s+/g, ' ').trim();
        let m = s.match(/^(.*?)\s*±\s*(\d+(?:\.\d+)?)$/);
        if (m) return { nominal: m[1].trim(), lower: '-' + m[2], upper: '+' + m[2] };
        // 斜杠公差：0/-0.025、+0.012/-0.006、+0.03/0
        m = s.match(/^(.*?)\s*([+-]?\d+(?:\.\d+)?)\s*\/\s*([+-]?\d+(?:\.\d+)?)$/);
        if (m) {
            return {
                nominal: m[1].trim(),
                lower: fmtDev(m[3], 'lower'),
                upper: fmtDev(m[2], 'upper')
            };
        }
        m = s.match(/^(.*?)\s*([+-]\d+(?:\.\d+)?)\s*([+-]\d+(?:\.\d+)?)$/);
        if (m) {
            const t1 = m[2], t2 = m[3];
            const low = parseFloat(t1) < parseFloat(t2) ? t1 : t2;
            const up = parseFloat(t1) > parseFloat(t2) ? t1 : t2;
            return { nominal: m[1].trim(), lower: low, upper: up };
        }
        m = s.match(/^(.*?)\s*\+(\d+(?:\.\d+)?)$/);
        if (m) return { nominal: m[1].trim(), lower: '', upper: '+' + m[2] };
        m = s.match(/^(.*?)\s*-(\d+(?:\.\d+)?)$/);
        if (m) return { nominal: m[1].trim(), lower: '-' + m[2], upper: '' };
        return { nominal: s, lower: '', upper: '' };
    }

    function classifyNominal(nominal) {
        const s = String(nominal || '').trim();
        if (/^Ø/.test(s)) return { type: '直径', nominal: s.replace(/^Ø\s*/, '') };
        if (/^R\s*\d/.test(s)) return { type: '半径', nominal: s.replace(/^R\s*/, '') };
        if (/^C\s*\d/.test(s)) return { type: '倒角', nominal: s.replace(/^C\s*/, '') };
        if (/^M\d/.test(s)) return { type: '螺纹', nominal: s.replace(/\s+/g, '') };
        if (/°$/.test(s)) return { type: '角度', nominal: s.replace(/\s*°$/, '').trim() };
        // 配合：H7/f6、16H7/f6、32h7、28H7、18f7
        if (
            /[A-Z]\d+\s*\/\s*[a-z]\d+/.test(s)
            || /^\d+(?:\.\d+)?\s*[A-Za-z]\d+$/.test(s)
        ) {
            return { type: '配合', nominal: s.replace(/\s+/g, '') };
        }
        return { type: '线性尺寸', nominal: s };
    }

    function addSize(raw, balloon) {
        const parts = splitSize(raw);
        const kind = classifyNominal(parts.nominal);
        if (!kind.nominal) return;
        push(balloon, kind.type, kind.nominal, parts.lower, parts.upper);
    }

    // ==========【改动点1：修改pushGdt，nominal、lower_tol为空，公差值放到upper_tol】==========
    function pushGdt(balloon, type, zone) {
        const num = String(zone || '').replace(/Ø/g, '').match(/\d+(?:\.\d+)?/);
        if (!num) return;
        push(balloon, type, '', '', num[0]);
    }

    function frameSymbolType(symbol) {
        if (!symbol) return '形位公差';
        if (symbol === '∩' || symbol === '⌓') return '面轮廓度';
        if (symbol === '⌒') return '线轮廓度';
        return gdtSymbols[symbol] || '形位公差';
    }

    function takeFeatureFrame(line, balloon) {
        const m = String(line || '').trim().match(
            /^(?:(\d+)\s+)?(\d+)\s*[xX×]\s*([⊕⌖◎∥⊥▱⌒⌓∠↗∩])?\s*(Ø?\s*\d+(?:\.\d+)?)(?:\s*[ⓂⓁⓅ])?(?:\s+([A-Z](?:\s+[A-Z]){0,5}))?\s*$/
        );
        if (!m) return false;
        const seq = balloon?.seq ?? m[1] ?? '';
        const importance = balloon?.importance ?? (seq ? '非重要' : '');
        const holder = balloon || { seq, importance };
        pushGdt(holder, frameSymbolType(m[3] || ''), m[4]);
        return true;
    }

    function takeAll(source, re, handler) {
        const src = String(source);
        const matches = [...src.matchAll(re)];
        for (const m of matches) handler(m);
        return matches.length > 0;
    }

    for (const src of lines) {
        if (RE_SKIP.test(src)) continue;
        if (RE_ISO_TOL.test(src)) continue;
        if (/^\d+\s*[:：]\s*\d+$/.test(src)) continue;
        if (RE_DATE.test(src)) continue;
        if (/^[A-Za-z]{1,8}[-.][A-Za-z0-9.-]{2,}$/.test(src)) continue;

        let balloon = readBalloon(src);
        let rest = src;
        if (balloon) {
            rest = balloon.rest;
            if (!rest) {
                pendingQueue.push(balloon);
                continue;
            }
        }
        let balloonBind = balloon;
        if (!balloonBind && pendingQueue.length > 0) {
            balloonBind = pendingQueue.shift();
        }

        if (takeFeatureFrame(rest, balloonBind)) {
            pendingRough = null;
            continue;
        }

        takeAll(rest, RE_GDT_SYM_LINE, m => {
            const sName = gdtSymbols[m[1]] || '形位公差';
            pushGdt(balloonBind, sName, m[2]);
        });
        takeAll(rest, RE_CIRCLE_ROUND, m => {
            pushGdt(balloonBind, '圆度', m[2]);
        });

        if (isRoughHeader(rest)) {
            pendingRough = { balloon: balloonBind, param: roughKind(rest) };
            continue;
        }
        if (pendingRough && /^\d+(?:\.\d+)?$/.test(rest)) {
            push(pendingRough.balloon || balloonBind, pendingRough.param || '表面粗糙度', rest, '', '');
            pendingRough = null;
            continue;
        }
        pendingRough = null;

        // ==========【改动点2：文本形位公差正则，捕获基准，基准暂不进表字段，如果你后续要渲染tooltip，我再加存放extra逻辑】==========
        let gdtHit = false;
        for (const name of gdtNames) {
            const gRe = new RegExp(name + '\\s*[:：]?\\s*(Ø?\\s*\\d+(?:\\.\\d+)?)(?:\\s*[ⓂⓁⓅ])?(?:\\s+([A-Z](?:\\s+[A-Z-]){0,2}))?');
            const gm = rest.match(gRe);
            if (!gm) continue;
            pushGdt(balloonBind, name, gm[1]);
            gdtHit = true;
        }
        if(gdtHit) continue;

        takeAll(rest, RE_ROUGH_TOKEN, m => {
            const r = parseRoughnessToken(m[0]);
            if (!r) return;
            push(balloonBind, r.param, r.valMin, '', r.valMax);
        });

        let hitSpecial = false;
        hitSpecial = takeAll(rest, RE_N_TIMES, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_METRIC_THREAD, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_DIAM, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_RADIUS, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_CHAMFER, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_ANGLE, m => addSize(m[0], balloonBind)) || hitSpecial;
        hitSpecial = takeAll(rest, RE_FIT, m => addSize(m[0], balloonBind)) || hitSpecial;
        // 直径/配合等已命中时不再走线性公差，避免 Ø50 0/-0.025 再拆出线性尺寸
        if (!hitSpecial) {
            takeAll(rest, RE_TOL, m => addSize(m[0], balloonBind));
            const bare = rest.trim();
            if (/^\d+(?:\.\d+)?$/.test(bare) && !/^\d{4}$/.test(bare)) addSize(bare, balloonBind);
        }
    }

    const r = {};
    for (const f of fieldList) {
        if (!f.pattern || f.pattern === '-') continue;
        try {
            const fm = text.match(new RegExp(f.pattern));
            if (fm && fm[1]) r[f.key] = fm[1].trim();
        } catch (e) { }
    }
    r[K.items] = items;
    r[K.count] = String(items.length);

    const columnAlias = new Set([K.seq, K.importance, K.dim_type, K.nominal, K.lower_tol, K.upper_tol]);
    const configured = fieldList.filter(col => col && col.key && columnAlias.has(col.key));
    let cols;
    if (configured.length) {
        cols = configured.filter(col => col.visible !== false);
    } else {
        cols = [
            { key: K.seq, name: '尺寸序号' },
            { key: K.importance, name: '重要性' },
            { key: K.dim_type, name: '类型' },
            { key: K.nominal, name: '标称值' },
            { key: K.lower_tol, name: '下公差' },
            { key: K.upper_tol, name: '上公差' }
        ];
    }
    if (items.length && cols.length) {
        r.tableData = {
            header: cols.map(col => col.name || col.key),
            rows: items.map(it => cols.map(col => {
                const v = it[col.key];
                return v === undefined || v === null ? '' : String(v);
            }))
        };
    }
    return r;
}

```

## 示例文本

```
① Ø30±0.015
② R12.5
③ C2
④ M8×1.25-6H
⑤ 45°
⑥ 20g6
⑦ 平行度 0.03 A
⑧ 垂直度 0.02 B
⑨ 位置度 0.05 A B C
⑩ Ra1.6
⑪ Rz6.3
⑫ Ø25 +0.021/-0.007
⑬ 100
⑭
75.5
⑮ 15 位置度 0.04 A
⑯ 2×Ø6
⑰ 4×M5×0.8
⑱ 圆度 0.01
⑲ 同轴度 0.02 A
⑳ 对称度 0.03 B
㉑ 线轮廓度 0.04
㉒ 面轮廓度 0.05 A
㉓ 倾斜度 0.025
㉔ 圆跳动 0.01 A
㉕ 全跳动 0.02 A
㊿ Ø120
(51) Ø50 +0.03/0
(52) Ø50 0/-0.025
(53) 32h7
(54) 16H7/f6
(55)
(56)
(57) Ø18
(58) Ra0.8
(59) R3
(60) C0.5
(61) M12×1.75
(62) 30°
(63) 平面度 0.02
(64) 直线度 0.015
(65) 平行度 Ø0.04 A
(66) 66 垂直度 0.03 A
(67) 6×R4
(68) 3×C1
(69) 2×Ø8±0.02
(70) 85
(71) 未注公差GB/T1804-m
(72) 材料 45#钢
(73) 重量 0.85kg
(74) 校对：张三
(75) Ø42 +0.018 -0.009
(76) 22.5
(77) Ra12.5
(78) Rz25
(79) 位置度 Ø0.1 A B
(80) 同轴度 Ø0.03 A
(81) Ø35 +0.012/-0.006
(82) 120°
(83) M16×2-6g
(84) 28H7
(85) 18f7
(86) 对称度 0.04 A
(87) 圆跳动 0.008 A
(88) 全跳动 0.03 A
(89) 线轮廓度 0.06 A
(90) 面轮廓度 Ø0.07
(91) 91 位置度 Ø0.06 A B
(92) 4×Ø10 +0.015/0
(93) 5×R2.5
(94) 2×C3
(95) Ø72 ±0.04
(96) 90
(97) 粗糙度 Ra3.2
(98) 粗糙度 Rz16
(99) 倾斜度 0.03
(100) 平面度 0.05

```
