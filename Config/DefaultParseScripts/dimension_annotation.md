---
id: dimension_annotation
name: 尺寸标注
description: 从机械图 OCR 文本中提取尺寸标注和形位公差，包括直径、半径、倒角、螺纹、角度、配合、表面粗糙度、基准和未注公差。
icon: 📏
category: 工程图
outputFormat: table
tags: 尺寸, 形位公差, 机械制图, 标注
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 标注条数 | count | - | 否 | 是 |
| 未注公差 | general_tolerance | - | 否 | 是 |
| 尺寸与形位公差 | annotations | - | 否 | 是 |

## 解析脚本

```javascript
function parse(rawText) {
    var text = String(rawText || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/[Φφ⌀]/g, 'Ø')
        .replace(/＋/g, '+')
        .replace(/－/g, '-')
        .replace(/[°º]/g, '°')
        .replace(/[０-９]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
        .replace(/[Ａ-Ｚａ-ｚ]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 0xFEE0); })
        .replace(/R\s*α/gi, 'Ra')
        .replace(/[ \t]+/g, ' ')
        .trim();

    var gdtNames = ['直线度', '平面度', '圆度', '圆柱度', '线轮廓度', '面轮廓度', '平行度', '垂直度', '倾斜度', '位置度', '同轴度', '同心度', '对称度', '圆跳动', '全跳动'];
    var gdtSymbols = {
        '⊕': '位置度', '⌖': '位置度', '◎': '同轴度', '∥': '平行度', '⊥': '垂直度',
        '▱': '平面度', '⌒': '线轮廓度', '⌓': '面轮廓度', '∠': '倾斜度', '↗': '圆跳动', '○': '圆度'
    };

    function isDeviation(line) {
        return /^[+-]\d+(?:\.\d+)?$/.test(line) || line === '0';
    }
    function isGdtLine(line) {
        var n;
        for (n = 0; n < gdtNames.length; n++) {
            if (line.indexOf(gdtNames[n]) >= 0) return true;
        }
        return /^[⊕⌖◎∥⊥▱⌒⌓∠↗○]/.test(line);
    }
    function isRoughLine(line) {
        return /粗糙度|(?:^|[^A-Za-z])R\s*[azqyα](?![A-Za-z])|[√∨˅✓✔▽▼]/i.test(line);
    }
    function roughKind(token) {
        var t = String(token || '').replace(/\s+/g, '').replace(/α/gi, 'a');
        if (/^ra$/i.test(t)) return 'Ra';
        if (/^rz$/i.test(t)) return 'Rz';
        if (/^ry$/i.test(t)) return 'Ry';
        if (/^rq$/i.test(t)) return 'Rq';
        return '';
    }
    function isRoughHeader(line) {
        return /^(?:(?:表面)?粗糙度\s*[:：]?|R\s*[azqyα]|[√∨˅✓✔▽▼])\s*$/i.test(String(line || '').trim());
    }

    var rawLines = text.split('\n').map(function (line) { return line.trim(); }).filter(function (line) { return line; });
    var lines = [];
    var i;
    for (i = 0; i < rawLines.length; i++) {
        var line = rawLines[i];
        if (lines.length && isDeviation(line) && /\d/.test(lines[lines.length - 1]) && !isGdtLine(lines[lines.length - 1]) && !isRoughLine(lines[lines.length - 1])) {
            lines[lines.length - 1] += ' ' + line;
        } else {
            lines.push(line);
        }
    }

    var items = [];
    var seen = {};
    var general = '';
    var pendingRough = null;

    function add(type, callout, tolerance, datum) {
        callout = String(callout || '').replace(/\s+/g, ' ').trim();
        tolerance = String(tolerance || '').replace(/\s+/g, ' ').trim();
        datum = String(datum || '').replace(/\s+/g, ' ').trim();
        if (!callout) return;
        var key = type + '|' + callout + '|' + datum;
        if (seen[key]) return;
        seen[key] = true;
        items.push({ '类型': type, '标注': callout, '公差': tolerance, '基准': datum });
    }

    function toleranceOf(s) {
        var found = String(s).match(/(±\s*\d+(?:\.\d+)?|\+\d+(?:\.\d+)?(?:\s+(?:-\d+(?:\.\d+)?|0))?)/);
        return found ? found[1].replace(/\s+/g, ' ').trim() : '';
    }

    function take(source, re, handler) {
        var flags = re.flags.indexOf('g') >= 0 ? re.flags : re.flags + 'g';
        var globalRe = new RegExp(re.source, flags);
        var match;
        var hits = [];
        while ((match = globalRe.exec(source)) !== null) {
            hits.push(match);
            if (match.index === globalRe.lastIndex) globalRe.lastIndex++;
        }
        hits.forEach(handler);
        return source.replace(new RegExp(re.source, flags), ' ');
    }

    var skipPrefix = /^(重量|材料|图号|图样|设计|校对|校核|审核|工艺|标准化|批准|比例|零件|姓名|日期|单位|阶段|共\s*\d|第\s*\d)/;

    for (i = 0; i < lines.length; i++) {
        var src = lines[i];
        if (/未注(?:尺寸|形位)?公差|GB\/T\s*1804|GB\/T\s*1184/.test(src)) {
            if (!general) general = src;
            add('未注公差', src, '', '');
            continue;
        }
        if (skipPrefix.test(src)) continue;
        if (/^\d+\s*[:：]\s*\d+$/.test(src)) continue;
        if (/^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}/.test(src)) continue;
        if (/^[A-Za-z]{1,8}[-.][A-Za-z0-9.-]{2,}$/.test(src)) continue;

        var rest = src;
        if (isRoughHeader(rest)) {
            pendingRough = roughKind(rest);
            continue;
        }
        if (pendingRough !== null && /^\d+(?:\.\d+)?$/.test(rest)) {
            var head = pendingRough;
            pendingRough = null;
            add('表面粗糙度', (head || '') + rest, '', '');
            continue;
        }
        pendingRough = null;
        rest = take(rest, /(?:(?:表面)?粗糙度\s*[:：]?\s*)?(?:[√∨˅✓✔▽▼]\s*)?((?:R\s*[azqyα]))\s*[:：=≤≦]?\s*(\d+(?:\.\d+)?)(?:\s*[\/~～]\s*(\d+(?:\.\d+)?))?|(?:表面粗糙度|粗糙度)\s*[:：]?\s*(\d+(?:\.\d+)?)(?:\s*[\/~～]\s*(\d+(?:\.\d+)?))?|[√∨˅✓✔▽▼]\s*(\d+(?:\.\d+)?)/gi, function (m) {
            var param = roughKind(m[1]);
            var value = m[2] || m[4] || m[6] || '';
            var lower = m[3] || m[5] || '';
            if (!value) return;
            var callout = (param || '') + value;
            if (lower) callout += '/' + lower;
            add('表面粗糙度', callout, '', '');
        });
        rest = take(rest, /基准\s*[:：]?\s*([A-Z](?:\s*[、,，]\s*[A-Z])*)/g, function (m) {
            var letters = m[1].replace(/[、,，]/g, ' ').replace(/\s+/g, ' ').trim();
            add('基准', '基准 ' + letters, '', letters);
        });

        var gi;
        for (gi = 0; gi < gdtNames.length; gi++) {
            var name = gdtNames[gi];
            var gRe = new RegExp(name + '\\s*[:：]?\\s*(Ø?\\s*\\d+(?:\\.\\d+)?)(?:\\s*\\(?[MLS]\\)?)?(?:\\s+([A-Z](?:\\s+[A-Z]){0,2}))?');
            var gm = rest.match(gRe);
            if (!gm) continue;
            var gTol = gm[1].replace(/\s+/g, '');
            var gDat = gm[2] ? gm[2].replace(/\s+/g, ' ').trim() : '';
            add('形位公差', name + ' ' + gTol + (gDat ? ' ' + gDat : ''), gTol, gDat);
            rest = rest.replace(gm[0], ' ');
        }

        rest = take(rest, /([⊕⌖◎∥⊥▱⌒⌓∠↗○])\s*(Ø?\s*\d+(?:\.\d+)?)(?:\s*\(?[MLS]\)?)?(?:\s+([A-Z](?:\s+[A-Z]){0,2}))?/g, function (m) {
            var sName = gdtSymbols[m[1]] || '形位公差';
            var sTol = m[2].replace(/\s+/g, '');
            var sDat = m[3] ? m[3].replace(/\s+/g, ' ').trim() : '';
            add('形位公差', sName + ' ' + sTol + (sDat ? ' ' + sDat : ''), sTol, sDat);
        });
        rest = take(rest, /M\d+(?:\.\d+)?(?:\s*[×xX]\s*\d+(?:\.\d+)?)?(?:-\d+[A-Za-z])?/g, function (m) {
            add('螺纹', m[0].replace(/\s+/g, ''), '', '');
        });
        rest = take(rest, /Ø\s*\d+(?:\.\d+)?(?:\s*±\s*\d+(?:\.\d+)?)?(?:\s*\+\d+(?:\.\d+)?(?:\s+(?:-\d+(?:\.\d+)?|0))?)?(?:\s*[A-Z]\d+(?:\/[a-z]\d+)?)?/g, function (m) {
            var d = m[0].replace(/\s+/g, ' ').trim().replace(/^Ø\s+/, 'Ø');
            add('直径', d, toleranceOf(d), '');
        });
        rest = take(rest, /(?:^|\s)R\s*\d+(?:\.\d+)?/g, function (m) {
            add('半径', m[0].replace(/^[^R]+/, '').replace(/\s+/g, ''), '', '');
        });
        rest = take(rest, /(?:^|\s)C\s*\d+(?:\.\d+)?/g, function (m) {
            add('倒角', m[0].replace(/^[^C]+/, '').replace(/\s+/g, ''), '', '');
        });
        rest = take(rest, /\d+(?:\.\d+)?\s*°/g, function (m) {
            add('角度', m[0].replace(/\s+/g, ''), '', '');
        });
        rest = take(rest, /\d+(?:\.\d+)?\s*[A-Z]\d+\s*\/\s*[a-z]\d+|[A-Z]\d+\s*\/\s*[a-z]\d+/g, function (m) {
            add('配合', m[0].replace(/\s+/g, ''), '', '');
        });
        rest = take(rest, /\d+(?:\.\d+)?\s*±\s*\d+(?:\.\d+)?|\d+(?:\.\d+)?\s+\+\d+(?:\.\d+)?(?:\s+(?:-\d+(?:\.\d+)?|0))/g, function (m) {
            var lv = m[0].replace(/\s+/g, ' ').trim();
            add('线性尺寸', lv, toleranceOf(lv), '');
        });

        var bare = rest.trim();
        if (/^\d+(?:\.\d+)?$/.test(bare) && !/^\d{4}$/.test(bare)) {
            add('线性尺寸', bare, '', '');
        }
    }

    var r = {};
    for (i = 0; i < FIELDS.length; i++) {
        var f = FIELDS[i];
        if (!f.pattern || f.pattern === '-') continue;
        try {
            var fm = text.match(new RegExp(f.pattern));
            if (fm && fm[1]) r[f.key] = fm[1].trim();
        } catch (e) {
        }
    }
    if (general) r.general_tolerance = general;
    r.annotations = items;
    r.count = String(items.length);
    if (items.length) {
        r.tableData = {
            header: ['类型', '标注', '公差', '基准'],
            rows: items.map(function (it) { return [it['类型'], it['标注'], it['公差'], it['基准']]; })
        };
    }
    return r;
}
```

## 示例文本

```
Ø25±0.02
Ø12.5
+0.021
0
R3
C1
M8×1.25-6H
45°
120±0.1
50
+0.02
-0.01
位置度 Ø0.05 A B
垂直度 0.03 A
圆度 0.01
同轴度 Ø0.02 A
⊥ 0.03 A
Ra 1.6
基准 A
H7/g6
未注公差按 GB/T 1804-m
比例
1:2
图号
HT-01-002
2024.03.15
```
