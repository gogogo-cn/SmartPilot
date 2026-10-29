/**
 * 竖排（旋转 90°）文字识别守栏（2026-10-25）。
 *
 * 背景（用户截图反馈）：机械图侧边的竖排标注「被拆成两个盒子、还重叠」，
 * 气泡丢掉，翻转过（旋转）的数字几乎全读错，识别盒子永远按左上角-右下角画，
 * 看不出文字是横是竖、从哪头读。
 *
 * 根因与修复：
 *   1) 检测器给的是轴对齐矩形，竖排文字整个框又高又窄，识别器按横排喂进去——
 *      实测竖排 F8 读成 "B8"（RapidOCR 0.43）/ "8."（Paddle 0.51）、竖排 Ø78.5 直接读空。
 *      两个引擎都加了 _rerecognize_vertical：又高又窄的框顺/逆时针各转 90° 识别一次
 *      （不走方向分类器，两个方向结果不同才开始比），置信度高的赢，并记下 vdir。
 *   2) 竖排标注常被检测器切成上下两半（Ø44.6 + F8），两个盒子重叠着显示。
 *      专项新增 vertical 阶段：左右对齐、上下挨着就按 vdir 并成一条；
 *      包在文字框里的小碎片（字少、置信度低）是误检，直接丢；带圈数字永不丢（球标）。
 *   3) 竖排 Ø 被读成 0（078.5），意思全变。按 vdir 在文字起点那一侧滑窗找斜线圈补回 Ø。
 *   4) 盒子朝向：ElementStructItem 输出按阅读方向排序的四边形 + vdir，
 *      UI 在文字起点那个角画方向箭头，不再永远左上角到右下角。
 *
 * 运行：node Tests/vertical-text-ocr.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const read = (rel) => fs.readFileSync(path.join(ROOT, rel), 'utf8');
const paddle = read('Services/paddle_inference_engine.py');
const rapid = read('Services/rapidocr_inference_engine.py');
const gdt = read('Services/GdtSymbolRecognizer.vb');
const bridge = read('Bridges/OcrBridge.vb');
const html = read('UI/ocr-panel.html');
const cfg = JSON.parse(read('Config/Correction/mechanical_drawing.json'));
const profile = read('Services/DomainPack/CorrectionProfile.vb');

test('两个引擎都有竖排重识别：转 90° 各认一次，不走方向分类器', () => {
    for (const [name, src, clsArg] of [['paddle', paddle, 'use_angle_cls=False'], ['rapid', rapid, 'use_cls=False']]) {
        assert.ok(src.includes('def _rerecognize_vertical('), `${name} 缺少 _rerecognize_vertical`);
        // 只处理明显竖排的框；矮框不碰
        assert.ok(/if h < w \* 1\.15 or h < 8:/.test(src), `${name} 竖排阈值应是 h >= w*1.15`);
        // 两个方向都试：270=顺时针 90°（原文逆时针放，从下往上读 up），90=逆时针（down）
        assert.ok(src.includes('(270, "up"), (90, "down")'), `${name} 要试顺/逆时针两个方向`);
        // 方向分类器必须关掉：两个方向结果相同就比不出阅读方向（paddle 用 use_angle_cls，rapid 用 use_cls）
        assert.ok(src.includes(clsArg), `${name} 重识别要关掉方向分类器（${clsArg}）`);
        // 文本只在旋转后更通时替换；方向与文本替换解耦（已读对的框也要有方向）
        assert.ok(/rot_best\[1\] > orig_conf/.test(src), `${name} 文本替换要明显更通才换`);
        assert.ok(/if rot_best\[2\] and rot_best\[1\] >= 0\.3 and not item\.get\("vdir"\)/.test(src), `${name} 方向要独立记录`);
    }
    // paddle 侧要给 _recognize_crop_batch 加 use_angle_cls 参数
    assert.ok(/def _recognize_crop_batch\(self, crops, use_angle_cls=None\)/.test(paddle), 'paddle 识别批要支持关分类器');
    // 两条出口都接上重识别（paddle 用 self. 前缀；rapid 匹配赋值行，避开 def 行）
    assert.strictEqual((paddle.match(/self\._rerecognize_vertical\(image_path,/g) || []).length, 2, 'paddle 的 ocr_text/ocr_pos 都要接');
    assert.strictEqual((rapid.match(/items = _rerecognize_vertical\(image_path, items\)/g) || []).length, 2, 'rapid 的 cmd_ocr/cmd_ocr_pos 都要接');
});

test('VB：竖排阶段已接线并默认开', () => {
    assert.ok(/If profile\.IsStageOn\("vertical"\) Then\s*\r?\n\s*MergeSplitVerticalText\(doc, targets\)\r?\n\s*RestoreVerticalDiameter\(bmp, targets\)/.test(gdt),
        'vertical 阶段要跑 合并 + 直径修复');
    // 放在球标检测之后：球标此时已定型，合并/清碎片不会误伤
    const balloonAt = gdt.indexOf('If profile.IsStageOn("balloon") Then DetectBalloons');
    const verticalAt = gdt.indexOf('If profile.IsStageOn("vertical") Then');
    assert.ok(balloonAt > 0 && verticalAt > balloonAt, 'vertical 阶段要排在 balloon 之后');
    assert.ok(profile.includes('"vertical"'), 'CorrectionProfile.CreateDefault 要登记 vertical');
    assert.strictEqual(cfg.stages.vertical, true, '配置模板 vertical 默认开');
});

test('VB：竖排拆分合并（左右对齐 + 上下挨着 + 按 vdir 接文本）', () => {
    assert.ok(gdt.includes('Private Shared Sub MergeSplitVerticalText'), '要有合并入口');
    assert.ok(gdt.includes('Private Shared Function VerticalStackPairs'), '要有配对判定');
    assert.ok(/overlapX < narrow \* 0\.5 Then Return False/.test(gdt), '左右要对齐（重叠过半）');
    assert.ok(/Return gap <= tall \* 0\.6 AndAlso gap >= -tall \* 0\.3/.test(gdt), '上下要挨着（间隙不超过六成高，允许略微重叠）');
    // up=从下往上读 → 下面的盒子是文本开头
    assert.ok(/Dim merged As OcrElement = If\(String\.Equals\(dir, "up", StringComparison\.OrdinalIgnoreCase\), bottom, top\)/.test(gdt),
        'up 方向时下盒为文本开头');
    assert.ok(/If dir = "" Then Continue For/.test(gdt), '不知道阅读方向就不乱并');
    // 带圈数字是球标，既不当碎片丢也不并进尺寸文本
    assert.strictEqual((gdt.match(/If CircledText\.IsMatch/g) || []).length, 2, '两处都要用 CircledText 保护球标');
});

test('VB：竖排 Ø 修复（起点侧滑窗找斜线圈）', () => {
    assert.ok(gdt.includes('Private Shared Sub RestoreVerticalDiameter'), '要有 Ø 修复入口');
    assert.ok(/If Not Regex\.IsMatch\(txt, "\^0\\d"\) Then Continue For/.test(gdt), '只修 0 开头的（0 就是被读错的 Ø）');
    // 滑窗：框内 3 格 + 框外 1 格，覆盖检测框偏移
    assert.ok(/For k = -1 To 3/.test(gdt), '要从起点往里滑窗再往框外探一格');
    assert.ok(/Dim fromBottom = String\.Equals\(dir, "up"/.test(gdt), 'up 方向 Ø 在下方');
    // 特征判定抽成共用函数，横竖排都用
    assert.ok(gdt.includes('Private Shared Function LooksLikeDiameterSymbol'), '直径符号特征要抽成共用函数');
});

test('VB：尾随符号裸偏差修回（0- → -0）', () => {
    assert.ok(/Dim tail = Regex\.Match\(txt, "\^\(\\d\+\(\?:\\\.\\d\+\)\?\)\(\[\+-\]\)\$"\)/.test(gdt), '要识别 0- 这种尾随符号');
    assert.ok(/If tail\.Success Then fixed_ = tail\.Groups\(2\)\.Value & tail\.Groups\(1\)\.Value/.test(gdt), '要把符号挪到前面');
});

test('盒子按阅读方向画：四边形按 vdir 排序 + UI 方向箭头', () => {
    assert.ok(bridge.includes('Private Shared Function OrientedQuad'), '桥要输出按方向排序的四边形');
    assert.ok(/pts = New List\(Of Point\) From \{\s*\r?\n\s*New Point\(b\.Left, b\.Bottom\),/.test(bridge),
        'up 方向起点是左下角（逆时针转一圈）');
    assert.ok(/\.vdir = e\.VerticalDir/.test(bridge), '结构化结果要带 vdir');
    // UI：竖排盒子画阅读方向箭头
    assert.ok(/const vdir = item\.vdir;/.test(html), 'UI 要读 vdir');
    assert.ok(/if \(vdir === 'up' \|\| vdir === 'down'\) \{/.test(html), '竖排才画箭头');
    assert.ok(/const dir = vdir === 'up' \? -1 : 1;/.test(html), '箭头方向跟着阅读方向');
});

test('引擎输出 vdir 一路通到 VB 元素', () => {
    assert.ok(/If item\("vdir"\) IsNot Nothing Then result\.VerticalDir = item\("vdir"\)\.ToString\(\)/.test(read('Services/OcrEngineProcess.vb')),
        'OcrEngineProcess 要解析 vdir');
    assert.ok(/If item\("vdir"\) IsNot Nothing Then result\.VerticalDir = item\("vdir"\)\.ToString\(\)/.test(read('Services/PaddleOcrService.vb')),
        '降级提取也要解析 vdir');
    assert.ok(gdt.includes('Public Property VerticalDir As String = ""') || read('Services/OcrFormatResult.vb').includes('Public Property VerticalDir As String = ""'),
        'OcrElement 要有 VerticalDir');
    assert.ok(read('Services/OcrFormatResult.vb').includes('Public Property Quad As List(Of Point) = Nothing'),
        'OcrElement 要有 Quad（原始四边形，画朝向用）');
    assert.ok(/cellElement\.VerticalDir = If\(r\.VerticalDir, ""\)/.test(read('Services/PaddleOcrService.vb')),
        '构建文档时要带上 vdir');
    assert.ok(/cellElement\.Quad = New List\(Of Point\)\(r\.Box\.Take\(4\)\)/.test(read('Services/PaddleOcrService.vb')),
        '构建文档时要带上四边形');
});
