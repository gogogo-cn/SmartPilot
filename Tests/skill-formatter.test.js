/**
 * ============================================================
 * Tests/skill-formatter.test.js — 格式整理能力合并测试
 * ============================================================
 * SmartTableFormatter 已并入 SmartTableBeautifier。
 * 验证：合并后 Beautifier 渲染器（planAndRender）支持格式整理字段
 * （字体/字号/数字格式/日期格式/换行/清理冗余），
 * 且"美化 + 格式整理"混合 spec 一次渲染完成。
 * 运行：node Tests/skill-formatter.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadEngine() {
    const window = {};
    window.App = { ChatV2: { services: {} }, config: {} };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/skill-spec-engine.js'), 'utf8'));
    return App.ChatV2.services.skillSpecEngine;
}

test('合并后: Beautifier 渲染器支持格式整理字段（美化+整理一次完成）', () => {
    const engine = loadEngine();
    // 混合 spec：美化（主题/条件格式）+ 格式整理（字体/数字/日期/换行/清理冗余）
    const spec = {
        theme: 'tech_green',
        headerColor: '深蓝',
        fontName: '微软雅黑', fontSize: 12,
        alignment: 'right', addBorders: true, autofit: true,
        freezeHeader: true, numberFormat: '#,##0', dateFormat: 'yyyy-mm-dd',
        wrapText: true, clearRedundant: true,
        conditionalFormat: 'dataBar', bandedRows: true
    };
    const r = engine.planAndRender(spec, {});
    assert.ok(r.ok, '渲染成功');
    // 美化能力
    assert.ok(r.code.includes('Sub SmartBeautifyTable()'), 'Beautifier 子程序');
    assert.ok(r.code.includes('AddDatabar'), '条件格式数据条');
    assert.ok(r.code.includes('MOD(ROW(),2)=0'), '斑马纹');
    // 格式整理能力（已并入）
    assert.ok(r.code.includes('.Name = "微软雅黑"'), '字体');
    assert.ok(r.code.includes('.Size = 12'), '字号');
    assert.ok(r.code.includes('.NumberFormat = "#,##0"'), '数字格式');
    assert.ok(r.code.includes('.NumberFormat = "yyyy-mm-dd"'), '日期格式');
    assert.ok(r.code.includes('FreezePanes'), '冻结首行');
    assert.ok(r.code.includes('.WrapText = True'), '自动换行');
    assert.ok(r.code.includes('selRng.ClearFormats'), '清理冗余格式');
    assert.ok(r.code.includes('RGB(30, 70, 140)'), '深蓝表头');
    console.log('✅ 合并后：美化+格式整理字段一次渲染完成');
});

test('合并后: 默认值（无格式字段时用 Beautifier 默认）', () => {
    const engine = loadEngine();
    const r = engine.planAndRender({}, {});
    assert.ok(r.code.includes('.Name = "微软雅黑"'), '默认字体微软雅黑');
    assert.ok(r.code.includes('.Size = 11'), '默认字号11（合并 Formatter 默认）');
    console.log('✅ 默认值');
});

test('合并后: 白名单防护（非法字段/注入）', () => {
    const engine = loadEngine();
    const malicious = {
        fontName: '"; MsgBox("hacked")',   // 注入
        fontSize: 999,                      // 非法字号
        numberFormat: '"; delete',          // 注入
        evilField: 'xxx'
    };
    const r = engine.planAndRender(malicious, {});
    assert.ok(r.ok, '非法 spec 不导致失败');
    assert.strictEqual(r.spec.fontName, '微软雅黑', '非法字体回退默认');
    assert.strictEqual(r.spec.fontSize, 11, '非法字号回退默认');
    assert.strictEqual(r.spec.numberFormat, null, '非法数字格式忽略');
    assert.ok(!('evilField' in r.spec), '未知字段过滤');
    assert.ok(!r.code.includes('MsgBox("hacked"'), '无注入');
    console.log('✅ 白名单防护');
});

test('合并后: SmartTableFormatter.md 已删除，Beautifier 定义完整', () => {
    const fs = require('fs');
    const skillsDir = path.join(__dirname, '..', 'Skills');
    assert.ok(!fs.existsSync(path.join(skillsDir, 'SmartTableFormatter.md')), 'Formatter.md 应已删除');
    const beautifier = fs.readFileSync(path.join(skillsDir, 'SmartTableBeautifier.md'), 'utf8');
    assert.ok(beautifier.includes('格式整理'), 'Beautifier 应含格式整理字段');
    assert.ok(beautifier.includes('fontName'), 'Beautifier specFields 含字体');
    assert.ok(beautifier.includes('dateFormat'), 'Beautifier specFields 含日期');
    assert.ok(beautifier.includes('wrapText'), 'Beautifier specFields 含换行');
    console.log('✅ Formatter 已删除，Beautifier 定义完整');
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
