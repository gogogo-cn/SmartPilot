---
id: mechanical_drawing
name: 机械制图
description: 解析机械制图 PDF 的标题栏与技术要求。标题和内容写在同一行，或各占一行，都可以提取。
icon: 📐
category: 工程图
outputFormat: table
tags: 机械制图, 工程图, PDF, 标题栏
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 图样名称 | part_name | `(?:零件名称|图样名称)[：:\s]*([^\n：:]{1,40})` | 否 | 是 |
| 图样代号 | drawing_no | `(?:图样代号|图号)[：:\s]*([A-Za-z0-9][A-Za-z0-9./_-]{1,30})` | 否 | 是 |
| 材料 | material | `材料[：:\s]*([^\n：:]{1,20})` | 否 | 是 |
| 比例 | scale | `比例[：:\s]*(\d+\s*[:：]\s*\d+)` | 否 | 是 |
| 重量 | weight | `重量[：:\s]*([\d.]+\s*(?:kg|KG|千克)?)` | 否 | 是 |
| 共张 | sheet_total | `共\s*(\d+)\s*张` | 否 | 是 |
| 第张 | sheet_index | `第\s*(\d+)\s*张` | 否 | 是 |
| 设计 | designer | `设计(?!日期)[：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 设计日期 | design_date | `设计(?:日期)?[：:\s]*[\u4e00-\u9fa5A-Za-z]{2,8}\s+(\d{4}[.\-/年]\d{1,2}[.\-/月]\d{1,2}日?)` | 否 | 是 |
| 校对 | checker | `校[对核][：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 审核 | reviewer | `审核[：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 工艺 | process_engineer | `工艺[：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 标准化 | standardizer | `标准化[：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 批准 | approver | `批准[：:\s]*([\u4e00-\u9fa5A-Za-z]{2,8})` | 否 | 是 |
| 单位名称 | company | - | 否 | 是 |
| 阶段标记 | phase | `阶段标记[：:\s]*([^\n：:]{1,12})` | 否 | 是 |
| 技术要求 | tech_requirements | - | 否 | 是 |

## 解析脚本

```javascript
function parse(rawText) {
    var text = String(rawText || '')
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n{3,}/g, '\n\n')
        .trim();

    var r = {};
    var i;
    for (i = 0; i < FIELDS.length; i++) {
        var f = FIELDS[i];
        if (!f.pattern || f.pattern === '-') continue;
        try {
            var m = text.match(new RegExp(f.pattern));
            if (m && m[1]) r[f.key] = m[1].replace(/[ \t]+/g, ' ').trim();
        } catch (e) {
        }
    }
    if (r.scale) r.scale = r.scale.replace(/\s/g, '');

    var labels = [
        ['零件名称', 'part_name'],
        ['图样名称', 'part_name'],
        ['图样代号', 'drawing_no'],
        ['图号', 'drawing_no'],
        ['材料', 'material'],
        ['比例', 'scale'],
        ['重量', 'weight'],
        ['设计日期', 'design_date'],
        ['设计', 'designer'],
        ['校对', 'checker'],
        ['校核', 'checker'],
        ['审核', 'reviewer'],
        ['工艺', 'process_engineer'],
        ['标准化', 'standardizer'],
        ['批准', 'approver'],
        ['单位名称', 'company'],
        ['阶段标记', 'phase']
    ];

    function findLabel(line) {
        var k;
        for (k = 0; k < labels.length; k++) {
            if (line === labels[k][0] || line.indexOf(labels[k][0]) === 0) return labels[k];
        }
        return null;
    }

    function isPureLabel(line) {
        var k;
        for (k = 0; k < labels.length; k++) {
            if (line === labels[k][0] || line === labels[k][0] + '：' || line === labels[k][0] + ':') return true;
        }
        return line === '技术要求' || line === '明细栏' || line === '标题栏';
    }

    var lines = text.split('\n').map(function (line) { return line.trim(); }).filter(function (line) { return line; });
    var n;
    for (n = 0; n < lines.length; n++) {
        var hit = findLabel(lines[n]);
        if (!hit || r[hit[1]]) continue;
        var same = lines[n].replace(new RegExp('^' + hit[0] + '[：:\\s]*'), '').trim();
        var value = same;
        var usedNext = false;
        if (!value && n + 1 < lines.length && !isPureLabel(lines[n + 1])) {
            value = lines[n + 1];
            usedNext = true;
        }
        if (!value) continue;
        if (hit[1] === 'drawing_no') value = value.split(/\s+/)[0];
        if (hit[1] === 'scale') value = value.replace(/\s/g, '');
        var person = /^(designer|checker|reviewer|process_engineer|standardizer|approver)$/.test(hit[1]);
        if (person) {
            var pm = value.match(/^([\u4e00-\u9fa5A-Za-z·]{2,8})(?:\s+(\d{4}[.\-\/年]\d{1,2}[.\-\/月]\d{1,2}日?))?/);
            if (pm) {
                r[hit[1]] = pm[1];
                if (hit[1] === 'designer' && pm[2] && !r.design_date) r.design_date = pm[2];
                if (hit[1] === 'designer' && !r.design_date && usedNext && n + 2 < lines.length && /^\d{4}[.\-\/年]/.test(lines[n + 2])) {
                    r.design_date = lines[n + 2].trim();
                }
                continue;
            }
        }
        r[hit[1]] = value.replace(/[ \t]+/g, ' ').trim();
    }

    if (!r.sheet_total) {
        var total = text.match(/共\s*(\d+)\s*张/);
        if (total) r.sheet_total = total[1];
    }
    if (!r.sheet_index) {
        var index = text.match(/第\s*(\d+)\s*张/);
        if (index) r.sheet_index = index[1];
    }

    if (!r.tech_requirements) {
        var start = -1;
        for (n = 0; n < lines.length; n++) {
            if (lines[n] === '技术要求' || lines[n].indexOf('技术要求') === 0) {
                start = n;
                break;
            }
        }
        if (start >= 0) {
            var buf = [];
            var head = lines[start].replace(/^技术要求[：:\s]*/, '');
            if (head) buf.push(head);
            for (n = start + 1; n < lines.length; n++) {
                if (isPureLabel(lines[n]) || findLabel(lines[n])) break;
                buf.push(lines[n]);
            }
            if (buf.length) r.tech_requirements = buf.join('\n');
        }
    }

    if (!r.company) {
        var c;
        for (c = lines.length - 1; c >= 0; c--) {
            if (lines[c].length <= 40 && /公司|研究院|厂$/.test(lines[c]) && !findLabel(lines[c])) {
                r.company = lines[c];
                break;
            }
        }
    }

    return r;
}
```

## 示例文本

```
技术要求
1.未注倒角 C0.5
2.去除毛刺
零件名称：法兰盘
图样代号：HT-01-002
材料：45
比例：1:2
重量：0.86kg
共 2 张  第 1 张
设计：张三 2024.03.15
校对：李四
审核：王五
某某机械有限公司
```
