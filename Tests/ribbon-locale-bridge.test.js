/**
 * ribbon-locale-bridge.test.js — 功能区（Ribbon）中英文切换链路回归
 *
 * 背景（2026-10-11 用户反馈「中英文切换对功能区按钮无效」）：
 *   Word/PPT 的桥对象继承 Hosts\Shared\HostBridgeBase.vb，自带 SaveUiLocale / GetUiLocale；
 *   Excel 的 ExcelBridge 是不继承该基类的门面类 → 前端 App.i18n.setLocale() 调
 *   bridge.SaveUiLocale() 在 Excel 里必然失败，且 WebView2 hostObject 对任意方法名
 *   typeof 都是 function，失败被 BridgeAccessor 吞掉 → 功能区语言永远停在旧值、无日志。
 *
 * 本测试锁死四件事：
 *   1) Excel 桥必须实现 SaveUiLocale / GetUiLocale，并写入 ui-locale.txt + Invalidate 功能区；
 *   2) 该实现必须在主工程 vbproj 中参与编译；
 *   3) 共享 Ribbon.xml 里每个控件都挂 getLabel，且 id 都在 GetControlLabel 中有中英分支；
 *   4) 前端 i18n.js 必须真的调用宿主的这两个方法（不能靠 typeof 提前 return 静默跳过）。
 *
 * 运行：node Tests/ribbon-locale-bridge.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const read = p => fs.readFileSync(p, 'utf8');
const exists = p => fs.existsSync(p);

const RIBBON_XML = path.join(ROOT, 'Hosts', 'Shared', 'Ribbon.xml');
const HOST_RIBBON_VB = path.join(ROOT, 'Hosts', 'Shared', 'HostRibbon.vb');
const HOST_BRIDGE_VB = path.join(ROOT, 'Hosts', 'Shared', 'HostBridgeBase.vb');
const EXCEL_LOCALE_VB = path.join(ROOT, 'ExcelBridge.LocaleApi.vb');
const EXCEL_VBPROJ = path.join(ROOT, 'ExcelAIVSTOAddIn.vbproj');
const I18N_JS = path.join(ROOT, 'UI', 'js', 'i18n.js');

test('Excel 桥实现功能区语言同步方法（SaveUiLocale / GetUiLocale）', () => {
    assert.ok(exists(EXCEL_LOCALE_VB), '缺少 ExcelBridge.LocaleApi.vb（Excel 此前无该桥方法）');
    const vb = read(EXCEL_LOCALE_VB);
    assert.match(vb, /Partial\s+Public\s+Class\s+ExcelBridge/, '必须是 ExcelBridge 的 partial 实现');
    assert.match(vb, /Function\s+SaveUiLocale\s*\(\s*locale\s+As\s+String\s*\)\s+As\s+String/, '缺 SaveUiLocale');
    assert.match(vb, /Function\s+GetUiLocale\s*\(\s*\)\s+As\s+String/, '缺 GetUiLocale');
    assert.match(vb, /HostRibbonLocale\.SaveLocale\(/, 'SaveUiLocale 必须落盘语言文件');
    assert.match(vb, /HostRibbonLocale\.Invalidate\(\)/, 'SaveUiLocale 必须刷新功能区（否则按钮不变）');
    // 两个方法都要暴露给 WebView2
    const comVisibleCount = (vb.match(/<ComVisible\(True\)>/g) || []).length;
    assert.ok(comVisibleCount >= 2, 'SaveUiLocale / GetUiLocale 都需 <ComVisible(True)>，实际 ' + comVisibleCount);
});

test('语言同步实现参与主工程编译', () => {
    const proj = read(EXCEL_VBPROJ);
    assert.match(proj, /<Compile Include="ExcelBridge\.LocaleApi\.vb"\s*\/>/, 'vbproj 未包含 ExcelBridge.LocaleApi.vb');
    assert.match(proj, /<Compile Include="Hosts\\Shared\\HostRibbon\.vb">/, 'vbproj 未包含共享 HostRibbon.vb（HostRibbonLocale 在此）');
});

test('Word/PPT 宿主桥仍保留同款语言同步方法（防回退）', () => {
    const vb = read(HOST_BRIDGE_VB);
    assert.match(vb, /Function\s+SaveUiLocale\s*\(\s*locale\s+As\s+String\s*\)\s+As\s+String/);
    assert.match(vb, /Function\s+GetUiLocale\s*\(\s*\)\s+As\s+String/);
    assert.match(vb, /HostRibbonLocale\.Invalidate\(\)/);
});

test('功能区 XML 每个控件都走 getLabel，且 GetControlLabel 都有中英分支', () => {
    const xml = read(RIBBON_XML);
    const ribbon = read(HOST_RIBBON_VB);

    const ids = [...xml.matchAll(/<(\w+)\s+id="([^"]+)"/g)].map(m => ({ tag: m[1], id: m[2] }));
    assert.ok(ids.length >= 4, '功能区控件数量异常: ' + ids.length);
    assert.ok(!xml.includes('btnAiModel') && !xml.includes('btnOcrSettings') && !xml.includes('btnMobileControl'),
        '功能区只保留智能聊天和识图入表；模型设置与手机遥控在侧栏');

    // 1) 每个控件都必须声明 getLabel（否则 label 写死在 XML 里，无法随语言切换）
    const missingGetLabel = [];
    const lineRe = /<(\w+)\s+id="([^"]+)"[^>]*?>/g;
    let m;
    while ((m = lineRe.exec(xml))) {
        if (!/getLabel="GetControlLabel"/.test(m[0])) missingGetLabel.push(m[1] + '#' + m[2]);
    }
    assert.deepStrictEqual(missingGetLabel, [], '这些控件未挂 getLabel="GetControlLabel": ' + missingGetLabel.join(', '));

    // 2) 每个 id 都要在 GetControlLabel 的 Select Case 中有分支
    const labelFn = /Function\s+GetControlLabel[\s\S]*?End Function/.exec(ribbon);
    assert.ok(labelFn, '找不到 GetControlLabel');
    const missing = ids.filter(x => !labelFn[0].includes('"' + x.id + '"')).map(x => x.id);
    assert.deepStrictEqual(missing, [], 'GetControlLabel 缺少分支（会显示原始 id）: ' + missing.join(', '));

    // 3) 每个分支都要有中英两套文案（If(en, "...", "...")）
    const branches = [...labelFn[0].matchAll(/Case\s+"([^"]+)"([\s\S]*?)(?=Case\s+"|Case\s+Else)/g)];
    assert.ok(branches.length >= ids.length - 1, '分支数量少于控件数量');
    const noPair = branches.filter(b => !/If\(en\s*,/.test(b[2])).map(b => b[1]);
    assert.deepStrictEqual(noPair, [], '这些分支没有中英两套文案: ' + noPair.join(', '));
});

test('界面语言目录：前端 LOCALE_CATALOG 与宿主 LocaleCatalog 短码一致', () => {
    const js = read(I18N_JS);
    const vb = read(HOST_RIBBON_VB);
    const jsStart = js.indexOf('var LOCALE_CATALOG');
    const jsEnd = js.indexOf('var FALLBACK_LOCALE', jsStart);
    assert.ok(jsStart > 0 && jsEnd > jsStart, 'i18n.js 缺少 LOCALE_CATALOG');
    const jsIds = [...js.slice(jsStart, jsEnd).matchAll(/id:\s*'([^']+)'/g)].map(m => m[1]);
    const vbStart = vb.indexOf('Function LocaleCatalog');
    const vbEnd = vb.indexOf('End Function', vbStart);
    assert.ok(vbStart > 0 && vbEnd > vbStart, 'HostRibbon.vb 缺少 LocaleCatalog');
    const vbIds = [...vb.slice(vbStart, vbEnd).matchAll(/New String\(\)\s*\{\s*"([^"]+)"/g)].map(m => m[1]);
    assert.deepStrictEqual(jsIds, vbIds, '追加语言时两边短码必须一起改');
    assert.ok(jsIds.length >= 2, '至少保留 zh / en');
    assert.match(vb, /Function\s+Normalize\s*\(/, '宿主需要 Normalize，保存 en-US 这类标签时收成短码');
    assert.match(vb, /MatchLocale\(/, '读写语言文件都要走同一套匹配');
});

test('前端 i18n 真的把语言同步给宿主（不靠 typeof 静默早退）', () => {
    const js = read(I18N_JS);
    assert.match(js, /bridge\.SaveUiLocale\s*\(/, 'i18n.js 未调用宿主 SaveUiLocale');
    assert.ok(!/typeof\s+bridge\.SaveUiLocale\s*!==\s*'function'/.test(js),
        '不能再用 typeof 判断桥方法可用性——WebView2 hostObject 对任意方法名 typeof 都是 function，' +
        '早退会让「切了语言但功能区不变」完全无日志（这正是本次 Excel 事故根因）');
    assert.match(js, /BridgeAccessor\._usingLocalBridge/, '网页版本地桥应被排除在功能区同步之外');
    assert.match(js, /function\s+resolveHostBridge/, '缺少宿主桥解析函数');
    // 首屏对齐：每个 WebView（含 PPT 主题模板这种独立窗格）都向宿主取语言。
    // 不能因为本页 localStorage 已有值就跳过——那份存储和主聊天不是同一个用户目录，
    // 跳过之后若再把本页语言写回宿主，会把用户刚切的语言覆盖掉。
    assert.match(js, /pullLocaleFromHostIfUnset/, '缺少首屏向宿主取语言的逻辑');
    assert.match(js, /bridge\.GetUiLocale\s*\(/, '首屏对齐需要调用宿主 GetUiLocale');
    assert.ok(!/if\s*\(\s*stored\s*===\s*'en'\s*\|\|\s*stored\s*===\s*'zh'\s*\)\s*return false/.test(js),
        '不能因本页 localStorage 已有语言就跳过宿主对齐');
    assert.ok(!/if\s*\(\s*!pulling\s*\)\s*syncLocaleToHost/.test(js),
        '启动时不能把本页 localStorage 反写宿主（PPT 主题模板窗格会因此把语言切回去）');
    assert.match(js, /function\s+applyExternalLocale/, '缺少 applyExternalLocale，已打开的主题模板窗格收不到语言切换');
    assert.match(js, /syncHost:\s*false/, '跟随宿主语言时必须禁止回写');
});

test('语言切换会推进已打开的 PPT 主题模板窗格，且不会清掉正在编辑的配色', () => {
    const excel = read(EXCEL_LOCALE_VB);
    const host = read(HOST_BRIDGE_VB);
    const ribbon = read(HOST_RIBBON_VB);
    const panes = read(path.join(ROOT, 'Hosts', 'Shared', 'HostPaneManager.vb'));
    const tasks = read(path.join(ROOT, 'TaskPaneManager.vb'));
    const editor = read(path.join(ROOT, 'UI', 'ppt-theme-editor.html'));

    assert.match(excel, /BroadcastUiLocale\(\)/, 'Excel SaveUiLocale 必须广播到其它 WebView');
    assert.match(excel, /ExecuteScriptAll\(script\)/, 'Excel 广播要覆盖任务窗格（主题模板在这里）');
    assert.match(excel, /SetPaneTitle\("PptTheme"/, 'Excel 要走到主题窗格标题入口');
    assert.ok(!/pane\.Title\s*=/.test(tasks),
        'Excel CustomTaskPane.Title 是只读的，赋值会 BC30526');
    assert.match(host, /ExecuteScriptAll\(HostRibbonLocale\.ApplyLocaleScript\(\)\)/, 'Word/PPT 宿主也要广播');
    assert.match(host, /SetPaneTitle\("ppttheme"/, 'Word/PPT 要更新主题模板窗格标题');
    assert.match(ribbon, /Function ApplyLocaleScript/, '缺少跨 WebView 的语言脚本');
    assert.match(ribbon, /Function PptThemePaneTitle/, '主题窗格标题要随语言变');
    assert.match(panes, /Sub ExecuteScriptAll/, 'HostPaneManager 缺少 ExecuteScriptAll');
    assert.match(tasks, /Sub ExecuteScriptAll/, 'TaskPaneManager 缺少 ExecuteScriptAll');

    const handlerAt = editor.indexOf("addEventListener('smartpilot:locale-changed'");
    assert.ok(handlerAt > 0, '主题编辑器没有监听语言切换');
    const handler = editor.slice(handlerAt, handlerAt + 1800);
    const snap = handler.indexOf('snapshotColors()');
    const rebuild = handler.indexOf('buildColorGrid()');
    const restore = handler.indexOf('restoreColors(colors)');
    assert.ok(snap >= 0 && rebuild > snap && restore > rebuild,
        '切换语言时必须先记住配色、重建色板后再写回，否则主题会被刷成默认商务蓝');
});
