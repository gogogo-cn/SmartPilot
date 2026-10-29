/**
 * 意图路由：模型自创的技能名不能进入路由。
 * 运行：node --test Tests/intent-route-catalog.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const BASE = path.join(__dirname, '..', 'UI', 'js', 'chat-v2');

function loadAnalyzer() {
    const window = {};
    window.App = { ChatV2: {} };
    global.window = window;
    globalThis.App = window.App;
    (0, eval)(fs.readFileSync(path.join(BASE, 'config/routing-constants.js'), 'utf8'));
    (0, eval)(fs.readFileSync(path.join(BASE, 'services/intent-analyzer.js'), 'utf8'));
    return App.ChatV2.intentAnalyzer;
}

const ia = loadAnalyzer();
const catalog = [
    { id: 'formatter', name: 'formatter' },
    { id: 'skill_creator', name: 'SkillCreator' },
    { id: 'free_chat', name: null }
];

test('目录外的技能名改为 free_chat，并记下被丢掉的名字', () => {
    const route = ia.coerceRouteToCatalog({
        pipeline: 'skill',
        skillId: 'SmartExcelFormatter',
        confidence: 1,
        reason: '目录里有格式化技能'
    }, catalog);
    assert.strictEqual(route.pipeline, 'text');
    assert.strictEqual(route.skillId, 'free_chat');
    assert.strictEqual(route.confidence, 0);
    assert.strictEqual(route.droppedSkillId, 'SmartExcelFormatter');
});

test('目录里的领域 id 保留', () => {
    const route = ia.coerceRouteToCatalog({
        pipeline: 'skill',
        skillId: 'formatter',
        confidence: 0.9,
        reason: ''
    }, catalog);
    assert.strictEqual(route.pipeline, 'skill');
    assert.strictEqual(route.skillId, 'formatter');
    assert.strictEqual(route.confidence, 0.9);
    assert.strictEqual(route.droppedSkillId, undefined);
});

test('技能展示名归一成目录 id', () => {
    const route = ia.coerceRouteToCatalog({
        pipeline: 'skill',
        skillId: 'SkillCreator',
        confidence: 0.8,
        reason: ''
    }, catalog);
    assert.strictEqual(route.skillId, 'skill_creator');
    assert.strictEqual(route.pipeline, 'skill');
});

test('目录未就绪时不改路由', () => {
    const raw = { pipeline: 'skill', skillId: 'SmartExcelFormatter', confidence: 1, reason: '' };
    assert.strictEqual(ia.coerceRouteToCatalog(raw, null), raw);
    assert.strictEqual(ia.coerceRouteToCatalog(raw, []), raw);
});

test('free_chat 保持通用路径', () => {
    const route = ia.coerceRouteToCatalog({
        pipeline: 'skill',
        skillId: 'free_chat',
        confidence: 0.4,
        reason: ''
    }, catalog);
    assert.strictEqual(route.pipeline, 'text');
    assert.strictEqual(route.skillId, 'free_chat');
    assert.strictEqual(route.droppedSkillId, undefined);
});

(0, eval)(fs.readFileSync(path.join(BASE, 'services/turn-goal.js'), 'utf8'));

test('合并并设置成中国红：模型只写成格式化时，意图补回两项', () => {
    const msg = '将选中工作表进行数据合并，并设置成中国红';
    const out = ia.completeIntent({
        needsPlanning: false,
        operations: ['format'],
        summary: '格式化选中工作表为红色',
        reason: '用户明确要求对选中工作表进行格式化操作'
    }, msg);
    assert.strictEqual(out.needsPlanning, true);
    assert.ok(out.operations.indexOf('mergeSheets') >= 0);
    assert.ok(out.operations.indexOf('format') >= 0);
    assert.ok(/合并/.test(out.summary), out.summary);
    assert.ok(/中国红/.test(out.summary), out.summary);
    assert.ok(!/格式化选中工作表为红色/.test(out.summary));
    assert.deepStrictEqual(out.operationLabels, ['合并工作表', '设置成中国红']);
});
