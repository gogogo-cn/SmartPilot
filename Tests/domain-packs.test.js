'use strict';
/**
 * 专项识别模块（DomainPack）守栏。
 *
 * 背景（2026-10-21）："机械图纸"开关原先硬编码在三处（UI 复选框 / localStorage /
 * OcrServiceManager.IsMechanicalMode 字符串比对），新增建筑图之类的专项每加一个都要
 * 改源码 + 重新编译。本次改为注册表：用户目录里的 .md 清单 + 内置模块登记。
 *
 * 本测试钉死：
 *   1. 三个新 .vb 已登记进 Core 工程（漏登记 = 文件白写，VB 里等于不存在）
 *   2. OcrServiceManager 的调用点已换成注册表分派，且保留向后兼容
 *   3. OcrBridge 暴露了 5 个 COM 方法
 *   4. 面板头部不显示任何专项状态（用户 2026-10-21 要求）：只保留
 *      「管理专项模块」入口，勾选/安装/卸载全部在弹窗内完成
 *   5. 示例模块清单字段齐全、options 是合法 JSON
 */
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

function read(rel) {
    return fs.readFileSync(path.join(ROOT, rel), 'utf8');
}

test('三个新 .vb 已登记进 Core 工程', () => {
    const proj = read('Core/SmartPilot.Core.vbproj');
    ['DomainPackManifest.vb', 'DomainPackLoader.vb', 'DomainPackRegistry.vb'].forEach((f) => {
        assert.ok(proj.indexOf(f) >= 0, 'Core.vbproj 应登记 Services/DomainPack/' + f);
        // 必须带 Link，否则在 Core 工程里的虚拟路径会散掉
        const idx = proj.indexOf(f);
        const chunk = proj.slice(idx - 120, idx + 200);
        assert.ok(/<Link>Ocr\\DomainPack\\/.test(chunk), f + ' 应链接到 Ocr\\DomainPack\\ 下');
    });
});

test('OcrServiceManager 已改用注册表分派', () => {
    const src = read('Services/OcrServiceManager.vb');
    assert.ok(src.indexOf('ApplyDomainPack') >= 0, '应通过 ApplyDomainPack 执行专项后处理');
    assert.ok(src.indexOf('ResolveDomainPack') >= 0, '应通过 ResolveDomainPack 解析模块');
    assert.ok(!/If IsMechanicalMode\(recognitionMode\) Then GdtSymbolRecognizer\.Apply/.test(src),
        '不得再硬编码 IsMechanicalMode → GdtSymbolRecognizer');
    // 向后兼容：旧 mode 字符串仍能映射到内置 GDT（mechanical / gdt / drawing）
    const reg = read('Services/DomainPack/DomainPackRegistry.vb');
    assert.ok(/mode = "mechanical" OrElse mode = "gdt" OrElse mode = "drawing"/.test(reg),
        '注册表必须保留 mechanical/gdt/drawing 到内置 GDT 的映射');
    // IsMechanicalMode 现在只是兼容壳，必须走注册表
    assert.ok(/Private Shared Function IsMechanicalMode[\s\S]*?ResolveDomainPack\(recognitionMode\)/.test(src),
        'IsMechanicalMode 应改为委托注册表（不再自己比字符串）');
});

test('内置机械制图模块登记正确', () => {
    const reg = read('Services/DomainPack/DomainPackRegistry.vb');
    assert.ok(reg.indexOf('BUILTIN_MECHANICAL_ID As String = "mechanical_drawing"') >= 0, '内置 id 应为 mechanical_drawing');
    assert.ok(reg.indexOf('BUILTIN_MECHANICAL_MODULE As String = "GdtSymbolRecognizer"') >= 0, '内置实现代号应为 GdtSymbolRecognizer');
    assert.ok(reg.indexOf('DomainPackPipeline.BuiltIn') >= 0, '机械制图应登记为 BuiltIn 流水线');
    const osm = read('Services/OcrServiceManager.vb');
    assert.ok(/GdtSymbolRecognizer\.Apply\(doc\)/.test(osm), 'GDT 实际调用点应保留在 OcrServiceManager');
});

test('OcrBridge 暴露 5 个专项模块 COM 方法', () => {
    const br = read('Bridges/OcrBridge.vb');
    ['GetDomainPacks', 'InstallDomainPackFromMarkdown', 'SetDomainPackEnabled',
        'UninstallDomainPack', 'SetDomainPackOption'].forEach((m) => {
        assert.ok(br.indexOf('Public Function ' + m) >= 0, 'OcrBridge 应有 ' + m);
    });
    // 每个方法前应有 ComVisible
    const seg = br.slice(br.indexOf('Public Function GetDomainPacks') - 200);
    assert.ok(seg.indexOf('<ComVisible(True)>') >= 0, '新方法必须 ComVisible(True) 才能被 JS 调用');

    // 识图面板拿到的是 ExcelBridge，不是 OcrBridge。漏转发时 JS 调用会得到 0x80020006 未知名称。
    const excel = read('ExcelBridge.vb');
    ['GetDomainPacks', 'InstallDomainPackFromMarkdown', 'SetDomainPackEnabled',
        'UninstallDomainPack', 'SetDomainPackOption'].forEach((m) => {
        assert.ok(excel.indexOf('Public Function ' + m) >= 0, 'ExcelBridge 应转发 ' + m);
        assert.ok(excel.indexOf('OcrBridge.' + m) >= 0, 'ExcelBridge.' + m + ' 应委托给 OcrBridge');
    });
});

test('面板头部不显示专项状态：只留管理入口，勾选/安装/卸载全在弹窗', () => {
    const html = read('UI/ocr-panel.html');
    assert.ok(html.indexOf('chk-mechanical-drawing') < 0, '不得再保留硬编码复选框');
    // 2026-10-21：用户要求头部彻底不显示专项状态（"专项识别：机械制图"也不要），
    // 启用状态在「管理专项模块」弹窗里一目了然，头部只保留入口按钮。
    assert.ok(html.indexOf('domain-packs-box') < 0, '头部不应再有专项状态容器');
    assert.ok(html.indexOf('renderDomainPacks') < 0, '不应再渲染头部状态标签');
    assert.ok(html.indexOf('packs-on') < 0, '不应再引用启用态样式');
    assert.ok(html.indexOf('packT') < 0, 'packT 随状态标签一并删除');
    assert.ok(html.indexOf('id="btn-manage-packs"') >= 0, '应保留「管理专项模块」入口');
    assert.ok(html.indexOf('async function loadDomainPacks') >= 0, '仍应从 GetDomainPacks 拉取模块（供识别模式推导）');
    assert.ok(html.indexOf('id="domain-pack-modal"') >= 0, '应有模块管理弹窗');
    assert.ok(html.indexOf('function activeRecognitionMode') >= 0, '识别模式应改由勾选模块推导');
    // 旧 mode 字符串仍被接受（向后兼容）
    assert.ok(/mode === 'mechanical'/.test(html), 'activeRecognitionMode 应继续接受旧 mechanical');
    // 勾选状态仍要写 localStorage 兜底（服务未就绪时机械制图不能消失）
    assert.ok(html.indexOf("localStorage.setItem('ocr-mechanical-drawing'") >= 0, '应保留 localStorage 兜底');
    // 弹窗里的启用复选框（data-dp）
    const mgrStart = html.indexOf('function renderDomainPackManager');
    assert.ok(mgrStart >= 0 && /data-dp/.test(html.slice(mgrStart, mgrStart + 4000)),
        '管理弹窗应保留启用复选框（data-dp）');

    // 状态文案键已全部删除，i18n 与面板都不得残留
    const i18n = read('UI/js/i18n.js');
    ['ocr.packs.active', 'ocr.packs.inactive', 'ocr.packs.empty'].forEach((k) => {
        assert.ok(i18n.indexOf("'" + k + "'") < 0, 'i18n 不应残留 ' + k);
        assert.ok(html.indexOf(k) < 0, '面板不应残留引用 ' + k);
    });
    // 管理按钮的键要保留且中英齐平
    const manageHits = i18n.match(/'ocr\.packs\.manage'\s*:/g) || [];
    assert.strictEqual(manageHits.length, 2, 'ocr.packs.manage 应中英各一份');
});

test('示例模块清单字段齐全且 options 是合法 JSON', () => {
    const md = read('Config/DomainPacks/title_block.md');
    const fm = md.slice(0, md.indexOf('---', 3));
    assert.ok(/^id:\s*\S+/m.test(fm), '示例应有 id');
    assert.ok(/^name:\s*\S+/m.test(fm), '示例应有 name');
    assert.ok(/^pipeline:\s*text/m.test(fm), '示例应为可用的 text 流水线');
    assert.ok(/^parseScript:\s*dimension_annotation/m.test(fm), '示例应复用现有 dimension_annotation 解析脚本');

    const m = md.match(/```options\s*\n([\s\S]*?)```/);
    assert.ok(m, '示例应有 options 代码块');
    const opts = JSON.parse(m[1]);
    assert.ok(Array.isArray(opts) && opts.length > 0, 'options 应是数组');
    opts.forEach((o) => {
        assert.ok(o.key, '每个选项都要有 key');
        assert.ok(o.type, '每个选项都要有 type');
        if (o.type === 'select') assert.ok(Array.isArray(o.values) && o.values.length > 0, 'select 选项要有 values');
    });
    // 示例引用的解析脚本必须真实存在
    const ps = path.join(ROOT, 'Config', 'DefaultParseScripts', 'dimension_annotation.md');
    assert.ok(fs.existsSync(ps), '示例引用的 dimension_annotation.md 必须存在');
});

test('识别改写入口已接线，并且只接受模块目录里的脚本', () => {
    const proj = read('Core/SmartPilot.Core.vbproj');
    assert.ok(proj.indexOf('DomainPackPythonRunner.vb') >= 0, 'Core.vbproj 应登记 DomainPackPythonRunner.vb');
    const runnerIdx = proj.indexOf('DomainPackPythonRunner.vb');
    const chunk = proj.slice(runnerIdx - 160, runnerIdx + 220);
    assert.ok(/<Link>Ocr\\DomainPack\\/.test(chunk), 'DomainPackPythonRunner.vb 应链接到 Ocr\\DomainPack\\ 下');

    const osm = read('Services/OcrServiceManager.vb');
    const fnStart = osm.indexOf('Private Shared Sub ApplyDomainPack');
    const fnEnd = osm.indexOf('Private Shared Function IsMechanicalMode');
    assert.ok(fnStart >= 0 && fnEnd > fnStart, '应能定位 ApplyDomainPack');
    const fn = osm.slice(fnStart, fnEnd);
    assert.ok(fn.indexOf('DomainPackPythonRunner.ApplyEnabled') >= 0, '识别完成后应调用 python 改写入口');
    assert.ok(fn.indexOf('GdtSymbolRecognizer.Apply') >= 0, '机械制图仍走内置识别器');
    assert.ok(!/If pack Is Nothing Then Return/.test(fn), '没有点名模块时也要跑已勾选的 python 模块');

    const runner = read('Services/DomainPack/DomainPackPythonRunner.vb');
    assert.ok(runner.indexOf('Function SafeScriptFileName') >= 0, '应限制脚本文件名');
    assert.ok(runner.indexOf('Contains("..")') >= 0, '应拒绝路径穿越');
    assert.ok(runner.indexOf('TimeoutMs As Integer = 15000') >= 0, '脚本应有超时，失败时不改识别结果');
    assert.ok(runner.indexOf('success') >= 0, '应按脚本返回的 success 决定是否回写');

    const reg = read('Services/DomainPack/DomainPackRegistry.vb');
    assert.ok(reg.indexOf('ExtractFencedBlock(content, "python")') >= 0, '面板安装应把 python 代码块写成脚本文件');
    assert.ok(reg.indexOf('File.Delete(scriptPath)') >= 0, '卸载应一并删除脚本');
});

test('机械制图矫正由配置开关控制，默认与原来一致', () => {
    const proj = read('Core/SmartPilot.Core.vbproj');
    assert.ok(proj.indexOf('CorrectionProfile.vb') >= 0, 'Core.vbproj 应登记 CorrectionProfile.vb');

    const gdt = read('Services/GdtSymbolRecognizer.vb');
    assert.ok(gdt.indexOf('CorrectionProfile.LoadMechanical') >= 0, '机械制图应读取矫正配置');
    ['gdtFrame', 'balloon', 'datum', 'roughness', 'mark', 'fcf', 'leader', 'groupDimensions'].forEach((stage) => {
        assert.ok(gdt.indexOf('IsStageOn("' + stage + '")') >= 0, '应能单独关闭阶段 ' + stage);
    });
    assert.ok(gdt.indexOf('ResolveRingScore()') >= 0, '球标圆环阈值应来自配置');
    assert.ok(gdt.indexOf('ResolveInk(') >= 0, '墨迹亮度阈值应来自配置');

    const profile = read('Services/DomainPack/CorrectionProfile.vb');
    assert.ok(profile.indexOf('DefaultInkLuminance As Integer = 150') >= 0, '默认墨迹阈值应仍是 150');
    assert.ok(profile.indexOf('DefaultRingScore As Double = 0.34') >= 0, '默认圆环得分应仍是 0.34');
    assert.ok(profile.indexOf('mechanical_drawing.json') >= 0, '用户用 JSON 覆盖，不走界面');

    const json = JSON.parse(read('Config/Correction/mechanical_drawing.json'));
    assert.strictEqual(json.inkLuminance, 150);
    assert.strictEqual(json.ringScore, 0.34);
    assert.deepStrictEqual(json.rules, []);
    ['gdtFrame', 'balloon', 'datum', 'roughness', 'mark', 'fcf', 'leader', 'groupDimensions'].forEach((stage) => {
        assert.strictEqual(json.stages[stage], true, stage + ' 默认应开启');
    });
});

test('安装框展示内置机械制图的识别代码', () => {
    const html = read('UI/ocr-panel.html');
    const open = '<script type="text/plain" id="builtin-recognizer-source">';
    const start = html.indexOf(open);
    const end = html.indexOf('</script>', start);
    assert.ok(start >= 0 && end > start, '页面里应带上识别器源码，不靠另外请求');
    const embedded = html.slice(start + open.length, end).replace(/^\n/, '').replace(/\s+$/, '').replace(/\r\n/g, '\n');
    const source = read('Services/GdtSymbolRecognizer.vb').replace(/^\uFEFF/, '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').replace(/\s+$/, '');
    assert.strictEqual(embedded, source, '页面里的代码应与 GdtSymbolRecognizer.vb 一致');
    assert.ok(embedded.indexOf('Private Shared Function Classify') >= 0, '应包含符号判定');
    assert.ok(embedded.indexOf('Function HasDiameterGlyph') >= 0, '应包含直径补写');
    assert.ok(html.indexOf('function renderBuiltinSource') >= 0, '打开安装框应渲染源码');
    assert.ok(html.indexOf('readEmbeddedRecognizer') >= 0, '打开时直接读页面里的源码');
    assert.ok(html.indexOf('DOMAIN_PACK_EXAMPLE') < 0, '安装框不再预填 JSON 示例');
    assert.ok(html.indexOf('title_block') < 0, '安装框不再预填字段抽取示例');
    assert.ok(html.indexOf('SaveMechanicalCorrection') >= 0, '保存机械制图配置应走专用入口');

    const excel = read('ExcelBridge.vb');
    assert.ok(excel.indexOf('Public Function SaveMechanicalCorrection') >= 0, 'ExcelBridge 应转发保存机械制图配置');
});
