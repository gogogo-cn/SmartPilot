/**
 * 机械制图「引线规则」守栏（2026-10-25）。
 *
 * 背景（用户截图反馈）：识别结果里 "49-0 业 352-0 +0.2 Rz25"、"9 11 10 北" 这类碎片
 * —— 文字框没有按引线（尺寸线 + 箭头）归组，箭头/引线末端的误识字（业、北、®）
 * 和丢小数点的公差（+02）原样留在输出里，公差栈被拆成三行。
 *
 * 本测试钉死四层：
 *   1) GdtSymbolRecognizer 新增 leaderRules 阶段：圆圈数字归一、公差小数修回、箭头碎片清除
 *   2) LeaderInk 互斥规则：一条引线最多收一条完整标注（公差栈除外）
 *   3) DrawingDimensionGrouper：名义值先当种子，偏差才能贴上去（公差栈合并）
 *   4) UI：专项启用兜底默认开 + 识别前等模块清单（否则整条专项被静默跳过）
 * 运行：node Tests/leader-rules-mechanical.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8');

const gdt = read('Services/GdtSymbolRecognizer.vb');
const leader = read('Services/LeaderInk.vb');
const grouper = read('Services/DrawingDimensionGrouper.vb');
const profile = read('Services/DomainPack/CorrectionProfile.vb');
const html = read('UI/ocr-panel.html');

test('leaderRules 阶段已接线：无图路径和主路径都跑文字层修整', () => {
    // 无图（云端 OCR / 临时文件被清）时也要修整：公差小数、箭头碎片都不需要墨迹
    assert.ok(/If targets IsNot Nothing AndAlso profile\.IsStageOn\("leaderRules"\) Then/.test(gdt),
        '无图早退路径应跑 leaderRules 文字层修整');
    // 公差小数和碎片清除要放在引线归组之前
    assert.ok(/If profile\.IsStageOn\("leaderRules"\) Then\s*\r?\n\s*RepairBareDeviations\(doc, targets\)\r?\n\s*DropLeaderJunk\(doc, targets\)/.test(gdt),
        'RepairBareDeviations + DropLeaderJunk 要放在 leader 阶段之前');
});

test('圆圈数字当球标：保留 ① 原文（解析脚本认 ① 不认裸 1）', () => {
    // 识别层扩展到带圈数字，但文本不改写——readBalloon 只认 ①/⬡/(1)
    assert.ok(gdt.includes('Private Shared ReadOnly BalloonDigit As New Regex("^(?:\\d{1,2}|[\\u2460-\\u2473\\u2776-\\u277F\\u24EA-\\u24FF\\u3251-\\u325F\\u32B1-\\u32BF])$"'),
        'BalloonDigit 要认五套带圈数字（①..⑳ ❶..❿ ⓪..⑳ ㉑..㉟ ㊱..㊿）');
    assert.ok(!gdt.includes('NormalizeCircledBalloons') && !gdt.includes('CircledDigits'),
        '不得再把 ① 改写成裸数字（会丢解析脚本的球标序号）');
    // LeaderInk 两侧同步扩展，且不写死 ^\\d{1,2}$
    assert.ok(/Private Shared ReadOnly BalloonPiece As New Regex\("\^\(\?:\\d\{1,2\}\|\[\\u2460-\\u2473/.test(leader),
        'LeaderInk.BalloonPiece 要认带圈数字');
    assert.strictEqual((leader.match(/If Not BalloonPiece\.IsMatch\(text\) Then Return False/g) || []).length, 2,
        'IsBalloonLike 与 IsArrowTag 都要改用 BalloonPiece');
    // 归组时球标排最前：readBalloon 只认行首的 ①
    assert.ok(/OrderBy\(Function\(el\) If\(el\.Type = OcrElementType\.Balloon, 0, 1\)\)/.test(grouper),
        '球标要排在块文本最前（期望「① Ø25±0.02」）');
});

test('公差小数修回：只修「符号+0+一位数字」，-0 和 +20 不碰', () => {
    assert.ok(gdt.includes('Private Shared ReadOnly LostDecimalDeviation As New Regex("(?<=^|\\s)([+±-])0(\\d)(?=\\s|$)"'),
        '正则应是「符号 + 0 + 一位数字」且自成一段');
    assert.ok(gdt.includes('LostDecimalDeviation.Replace(txt, "${1}0.$2")'),
        '替换串必须用 ${1} 占位（$10 会被当成第 10 组，且旧写法多插了一个空格）');

    // 语义镜像：VB 正则与 JS 等价，直接在 node 里验行为
    const re = /(?<=^|\s)([+±-])0(\d)(?=\s|$)/g;
    const fix = (s) => s.replace(re, '$1 0.$2'.replace(' ', '')); // $1 + '0.' + $2
    const fix2 = (s) => s.replace(re, (m, a, b) => a + '0.' + b);
    assert.strictEqual(fix2('+02'), '+0.2');
    assert.strictEqual(fix2('-01'), '-0.1');
    assert.strictEqual(fix2('±02'), '±0.2');
    assert.strictEqual(fix2('22.4-0 +02'), '22.4-0 +0.2');
    assert.strictEqual(fix2('-0'), '-0');            // 真零偏差
    assert.strictEqual(fix2('+20'), '+20');          // 可能是真公差
    assert.strictEqual(fix2('Ø50 0/-0.025'), 'Ø50 0/-0.025');
    assert.strictEqual(fix2('M8×1.25-6H'), 'M8×1.25-6H');
    assert.strictEqual(fix2('A-05'), 'A-05');        // 符号不在段首
    assert.strictEqual(typeof fix, 'function');
});

test('箭头碎片清除：业/北/® 清掉，深/均布/沉孔和标题栏字保留', () => {
    assert.ok(gdt.includes('Private Shared Sub DropLeaderJunk'), '应有碎片清除入口');
    assert.ok(gdt.includes('Private Shared Function IsLeaderJunk'), '应有判定函数');
    // 白名单：图纸标注词 + 标题栏单字
    for (const ch of ['深', '均', '布', '沉', '锪', '孔', '重', '量', '材', '料', '设', '计']) {
        assert.ok(gdt.includes('"' + ch + '"c'), '白名单缺 ' + ch);
    }
    // 符号碎片
    for (const ch of ['®', '©', '™', '§', '※']) {
        assert.ok(gdt.includes('"' + ch + '"c'), '符号碎片表缺 ' + ch);
    }
    // 长度上限 2：更长的文本（注释、标题）一律不动
    assert.ok(/If txt = "" OrElse txt\.Length > 2 Then Return False/.test(gdt),
        '只清 1-2 个字符的碎片');
    // 已定类型的不清（GDT 符号、球标、基准圆）
    assert.ok(/If el\.Type = OcrElementType\.Balloon OrElse el\.Type = OcrElementType\.DatumCircle OrElse el\.Type = OcrElementType\.GdtSymbol Then Return False/.test(gdt),
        'Balloon/DatumCircle/GdtSymbol 不得当碎片清掉');
    // 双份列表 + RawText 行都要清
    assert.ok(/Dim other = If\(Object\.ReferenceEquals\(targets, doc\.Elements\), doc\.RawCellElements, doc\.Elements\)\r?\n\s*If other IsNot Nothing Then other\.RemoveAll/.test(gdt),
        '要同步清掉另一份镜像列表');
    assert.ok(gdt.includes('RemoveRawTextLines(doc, If(el.Text, "").Trim())'),
        'RawText 里对应的行也要删');
    // 判定语义镜像
    const keep = new Set([...'深均布沉锪孔宽厚长高度重量材料比例设计校对审核批准日期阶段共张第名称图号代号密级数质']);
    const junkSym = new Set([...'®©™§※†‡•¶']);
    const isJunk = (t) => {
        if (!t || t.length > 2) return false;
        for (const ch of t) {
            if (ch >= '\u4e00' && ch <= '\u9fff') { if (!keep.has(ch)) return true; }
            else if (junkSym.has(ch)) return true;
        }
        return false;
    };
    assert.strictEqual(isJunk('业'), true);
    assert.strictEqual(isJunk('北'), true);
    assert.strictEqual(isJunk('®'), true);
    assert.strictEqual(isJunk('®B'), true);
    assert.strictEqual(isJunk('深'), false);
    assert.strictEqual(isJunk('重'), false);
    assert.strictEqual(isJunk('Ø'), false);
    assert.strictEqual(isJunk('R'), false);
    assert.strictEqual(isJunk('○'), false);   // 圆度符号
    assert.strictEqual(isJunk('±'), false);
    assert.strictEqual(isJunk('未注公差'), false); // 长度 > 2
});

test('基准圆收编原字母框，不再输出「ⓍA」+游离「A」两份', () => {
    assert.ok(/If TryAddDatumCircle\(ink, box, gcx, gcy, txt, el, doc\) Then consumed\.Add\(el\)/.test(gdt),
        '单字母命中基准圆后要收编原框');
    assert.ok(/Private Shared Function TryAddDatumCircle[\s\S]*?As Boolean/.test(gdt),
        'TryAddDatumCircle 要返回是否命中');
    assert.ok(/targets\.Remove\(el\)\r?\n\s*Dim twin = If\(Object\.ReferenceEquals\(targets, doc\.Elements\), doc\.RawCellElements, doc\.Elements\)/.test(gdt),
        '收编要同时删 targets 和镜像列表');
    assert.ok(/RemoveRawTextLines\(doc, If\(el\.Text, ""\)\.Trim\(\), 1\)/.test(gdt),
        'RawText 只删一行（同名基准字母可能是另一条标注）');
});

test('引线归组互斥：一条引线最多一条完整标注，公差栈除外', () => {
    // CanJoin：候选是完整标注时，种子也是完整标注直接挡；种子是裸偏差只放行叠放的
    assert.ok(/If IsSeparateCallout\(text\) Then\s*\r?\n\s*If IsSeparateCallout\(Compact\(seed\.Text\)\) Then Return False\s*\r?\n\s*If Not StackedOnSeed\(seed, other\) Then Return False\s*\r?\n\s*End If/.test(leader),
        'CanJoin 要按「块里只能有一条完整标注」挡，裸偏差种子只放行公差栈');
    assert.ok(leader.includes('Private Shared Function StackedOnSeed'),
        '应有 StackedOnSeed 判定（上下叠放 vs 并排）');
    // Collect 整块兜底
    assert.ok(/Dim hasCallout = IsSeparateCallout\(Compact\(seed\.Text\)\)/.test(leader),
        'Collect 要按整块统计完整标注');
    assert.ok(/If IsSeparateCallout\(Compact\(other\.Text\)\) Then\s*\r?\n\s*If hasCallout Then Continue For\s*\r?\n\s*hasCallout = True\s*\r?\n\s*End If/.test(leader),
        'Collect 要拒绝第二条完整标注');
    // ± 也是裸偏差（旧正则不含 ±，会把 ±0.2 的公差栈挡掉）
    assert.ok(/Return Regex\.IsMatch\(text, "\^\(\?:\[±\+-\]\\d\+\(\?:\\\.\\d\+\)\?\|0\)\$"\)/.test(leader),
        'LeaderInk.IsBareDeviation 要含 ±');
});

test('公差栈合并：名义值先当种子，偏差才贴得上去', () => {
    assert.ok(/OrderBy\(Function\(el\) If\(IsDeviation\(Compact\(el\.Text\)\), 1, 0\)\)/.test(grouper),
        'GroupLoose 要把裸偏差排到名义值之后（否则上偏差先当种子，公差栈被拆碎）');
    // DeviationToken 是排序依据，必须覆盖 ± 和裸 0
    assert.ok(grouper.includes('DeviationToken As New Regex("^(?:±\\s*\\d+(?:\\.\\d+)?|[+-]\\d+(?:\\.\\d+)?|0)$"'),
        'DeviationToken 要含 ± 与裸 0');
});

test('叠放偏差必须与名义值水平重叠（防合并块把远处的公差吸进来）', () => {
    // 沙盒实测：合并块的框是整条公差栈的高度，dy 阈值按 nh*1.35 放大，
    // 会把隔着一条尺寸的另一个 +0.2 并进块里（"22.4 +0.2 -0 +0.2"）。
    assert.ok(/Dim overlap = Math\.Min\(nominal\.BoundingBox\.Right, candidate\.BoundingBox\.Right\) -\s*\r?\n\s*Math\.Max\(nominal\.BoundingBox\.Left, candidate\.BoundingBox\.Left\)\s*\r?\n\s*Return overlap > 0/.test(grouper),
        'IsStackedDeviation 结尾要有水平重叠判定');
    // 语义镜像：重叠为 0 时必须拒绝
    const stacked = (n, c) => {
        const nh = Math.max(8, n.b - n.t), ch = c.b - c.t;
        if (ch > nh * 0.95 && !/^[±+-]/.test('+0.2')) return false;
        const dy = Math.abs((c.t + c.b) / 2 - (n.t + n.b) / 2);
        if (dy < nh * 0.2 || dy > nh * 1.35) return false;
        const cx = (c.l + c.r) / 2;
        if (cx < n.l - nh * 0.3 || cx > n.r + nh) return false;
        return Math.min(n.r, c.r) - Math.max(n.l, c.l) > 0;
    };
    const nominal = { l: 300, t: 224, r: 352, b: 250 };      // 22.4
    assert.strictEqual(stacked(nominal, { l: 300, t: 196, r: 342, b: 222 }), true,  '正上方公差应接受');
    assert.strictEqual(stacked(nominal, { l: 302, t: 250, r: 322, b: 276 }), true,  '正下方公差应接受');
    assert.strictEqual(stacked(nominal, { l: 390, t: 330, r: 432, b: 354 }), false, '隔着一条尺寸的公差必须拒绝');
});

test('阶段开关与配置：leaderRules 默认开', () => {
    assert.ok(/For Each name In New String\(\) \{[^}]*"leaderRules"/.test(profile),
        'CorrectionProfile.CreateDefault 要登记 leaderRules');
    const json = JSON.parse(read('Config/Correction/mechanical_drawing.json'));
    assert.strictEqual(json.stages.leaderRules, true, '配置模板里 leaderRules 默认开');
    assert.ok(Array.isArray(json.builtin.leaderRules) && json.builtin.leaderRules.length >= 3,
        '配置模板应有 leaderRules 的阶段说明');
});

test('UI：阶段标签、专项启用兜底、识别前等模块清单', () => {
    assert.ok(html.includes("['leaderRules', '引线规则']"), '矫正配置的阶段标签要有引线规则');
    // 兜底默认启用（与后端 disabled:[] 同口径），只有显式点掉才关
    assert.ok(html.includes("localStorage.getItem('ocr-mechanical-drawing') !== '0'"),
        '兜底应从「默认关」改为「默认开」，否则 GetDomainPacks 未回来时机械制图被静默跳过');
    assert.ok(/if \(!window\.__domainPacks && typeof loadDomainPacks === 'function'\) \{[\s\S]{0,200}await loadDomainPacks\(\)/.test(html),
        '识别前要等模块清单（DOMContentLoaded 才异步拉，存在竞态）');
    // 内嵌识别器源码与 .vb 保持一致（domain-packs.test.js 逐字节比对，这里防呆）
    const open = '<script type="text/plain" id="builtin-recognizer-source">';
    const start = html.indexOf(open);
    const end = html.indexOf('</script>', start);
    const embedded = html.slice(start + open.length, end).replace(/^\n/, '').replace(/\s+$/, '').replace(/\r\n/g, '\n');
    const source = read('Services/GdtSymbolRecognizer.vb').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\s+$/, '');
    assert.strictEqual(embedded, source, 'UI/ocr-panel.html 内嵌源码应与 Services/GdtSymbolRecognizer.vb 一致');
    assert.ok(embedded.includes('DropLeaderJunk'), '内嵌源码应包含新阶段');
});
