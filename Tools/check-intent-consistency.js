#!/usr/bin/env node
/**
 * ============================================================
 * Tools/check-intent-consistency.js — 意图配置一致性校验
 * ============================================================
 * 作用：检测意图配置是否同步，避免"换词就失效"回归
 *   ① Config/intent_config.json 是唯一事实源（单一数据源）
 *   ② 前端内嵌兜底：UI/js/chat-v2/services/intent-config.js 的 DEFAULT_CONFIG
 *   ③ 后端内嵌兜底：Services/Mcp/ExcelMcpTools.IntentConfigLoader.vb 的 BuildDefaultBackendConfig
 *   ④ 前端消费方不得再内嵌关键词字面量（intent-detector / skill-spec-engine / TaskPlanner / StepExecutor）
 *
 * 用法：node Tools/check-intent-consistency.js
 * 退出码：0=一致  1=存在差异（可用于构建前拦截）
 * ============================================================
 */
'use strict';
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const jsonFile = path.join(ROOT, 'Config', 'intent_config.json');
const jsFile = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'intent-config.js');
const vbFile = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.IntentConfigLoader.vb');
const consumers = [
    'UI/js/chat-v2/services/intent-detector.js',
    'UI/js/chat-v2/services/skill-spec-engine.js',
    'UI/js/chat-v2/services/TaskPlanner.js',
    'UI/js/chat-v2/services/StepExecutor.js'
];

let hasDiff = false;
const out = (msg) => console.log(msg);

// --- 文件存在检查 ---
for (const f of [jsonFile, jsFile, vbFile]) {
    if (!fs.existsSync(f)) {
        out(`[ERROR] File not found: ${f}`);
        process.exit(1);
    }
}

// --- 1. 解析 JSON ---
const jsonObj = JSON.parse(fs.readFileSync(jsonFile, 'utf8'));
out(`[1/4] intent_config.json 已解析 (version=${jsonObj.version})`);

// --- 辅助：从 JS/VB 源码提取字符串数组 ---
function extractStringArray(content, varName) {
    // 匹配 varName = ["a", "b"] / varName: ['a', 'b'] / result.X = {"a", "b"}
    const esc = varName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const re = new RegExp(esc + '\\s*[:=]\\s*[\\[\\{]([^\\]\\}]*)[\\]\\}]');
    const m = content.match(re);
    if (!m) return null;
    const items = [];
    const strRe = /"([^"]*)"|'([^']*)'/g;
    let sm;
    while ((sm = strRe.exec(m[1])) !== null) items.push(sm[1] !== undefined ? sm[1] : sm[2]);
    return items;
}

function arraysEqual(a, b, label) {
    if (!Array.isArray(a) || !Array.isArray(b)) {
        out(`[DIFF] ${label} : 无法提取数组（结构可能已变更）`);
        return false;
    }
    const sa = [...a].sort();
    const sb = [...b].sort();
    if (sa.length !== sb.length) {
        out(`[DIFF] ${label} : 数量不一致 (JS=${sa.length} vs JSON=${sb.length})`);
        return false;
    }
    for (let i = 0; i < sa.length; i++) {
        if (sa[i] !== sb[i]) {
            out(`[DIFF] ${label} : 内容不一致 (JS='${sa[i]}' vs JSON='${sb[i]}')`);
            return false;
        }
    }
    return true;
}

// --- 2. 前端 DEFAULT_CONFIG 关键段 vs JSON ---
const jsContent = fs.readFileSync(jsFile, 'utf8');
const jsChecks = [
    { label: '前端 chartWords', jsVar: 'chartWords', jsonVal: jsonObj.chartWords },
    { label: '前端 browserDisplayKeywords', jsVar: 'browserDisplayKeywords', jsonVal: jsonObj.browserDisplayKeywords },
    { label: '前端 quickPass', jsVar: 'quickPass', jsonVal: jsonObj.quickPass },
    { label: '前端 backend.moneyHints', jsVar: 'moneyHints', jsonVal: jsonObj.backend.moneyHints }
];
for (const c of jsChecks) {
    const jsArr = extractStringArray(jsContent, c.jsVar);
    if (!arraysEqual(jsArr, c.jsonVal, c.label)) hasDiff = true;
}

// 前端 synonyms 键集 vs JSON
const jsSynKeys = [...jsContent.matchAll(/'([^']+)':\s*\[/g)].map(m => m[1]).filter(k => k !== 'keywords');
const jsonSynKeys = Object.keys(jsonObj.backend.synonyms || {});
const synKeyDiff = [...new Set([...jsSynKeys.filter(k => !jsonSynKeys.includes(k)), ...jsonSynKeys.filter(k => !jsSynKeys.includes(k))])];
if (synKeyDiff.length > 0) {
    hasDiff = true;
    out(`[DIFF] 前端 DEFAULT_CONFIG.synonyms 键集与 JSON 不一致: ${synKeyDiff.join(', ')}`);
} else {
    out('[2/4] 前端 DEFAULT_CONFIG 关键段与 JSON 一致');
}

// --- 3. 后端 BuildDefaultBackendConfig vs JSON backend 段 ---
const vbContent = fs.readFileSync(vbFile, 'utf8');
const vbMoney = extractStringArray(vbContent, 'result.MoneyHints');
if (!arraysEqual(vbMoney, jsonObj.backend.moneyHints, '后端 MoneyHints')) hasDiff = true;

const vbDimFallback = extractStringArray(vbContent, 'result.DimensionFallback');
if (!arraysEqual(vbDimFallback, jsonObj.backend.dimensionFallback, '后端 DimensionFallback')) hasDiff = true;

const vbSynKeys = [...vbContent.matchAll(/\{"([^"]+)",\s*New List\(Of String\)/g)].map(m => m[1]);
const vbSynKeyDiff = [...new Set([...vbSynKeys.filter(k => !jsonSynKeys.includes(k)), ...jsonSynKeys.filter(k => !vbSynKeys.includes(k))])];
if (vbSynKeyDiff.length > 0) {
    hasDiff = true;
    out(`[DIFF] 后端 BuildDefaultBackendConfig.synonyms 键集与 JSON 不一致: ${vbSynKeyDiff.join(', ')}`);
} else {
    out('[3/4] 后端 BuildDefaultBackendConfig 关键段与 JSON 一致');
}

// --- 4. 前端消费方不再内嵌关键词字面量（检测正则） ---
// 只检查核心检测引擎 intent-detector.js 的"正则字面量中的关键词"。
// 该模块是唯一从用户消息提取关键词的引擎；若其检测正则内嵌中文关键词，
// 则"换词就失效"。TaskPlanner/StepExecutor 的正则匹配的是工具名/提示词，不涉及用户消息检测。
const bannedLiterals = ['饼图', '柱状图', '折线图', '汇总', '求和', '部门', '地区', '销售额', '利润', '客户数'];
const consumerDiffs = [];
for (const rel of consumers) {
    if (!rel.includes('intent-detector')) continue; // 仅检查核心检测引擎
    const full = path.join(ROOT, rel);
    if (!fs.existsSync(full)) continue;
    const content = fs.readFileSync(full, 'utf8');
    // 剥离 // 与 /* */ 注释，避免注释里的关键词误报
    const noComments = content
        .replace(/\/\*[\s\S]*?\*\//g, '')
        .replace(/(^|[^:])\/\/.*$/gm, '$1');
    // 提取正则字面量 /.../（含 flags），仅保留含中文的（检测关键词的标志）
    const reLiterals = [];
    const reRe = /\/((?:[^\\/\n]|\\.)+)\/[gimsuy]*/g;
    let rm;
    while ((rm = reRe.exec(noComments)) !== null) {
        if (/[\u4e00-\u9fa5]/.test(rm[1])) reLiterals.push(rm[1]);
    }
    for (const lit of bannedLiterals) {
        if (reLiterals.some(r => r.includes(lit))) {
            consumerDiffs.push(`${rel} 检测正则含 '${lit}'`);
        }
    }
}
if (consumerDiffs.length > 0) {
    hasDiff = true;
    out('[DIFF] 前端消费方仍内嵌关键词字面量（应改由 intent_config.json 下发）:');
    for (const d of [...new Set(consumerDiffs)]) out(`  - ${d}`);
} else {
    out('[4/4] 前端消费方无内嵌关键词字面量');
}

out('');
if (hasDiff) {
    out('[RESULT] INCONSISTENT! intent_config.json 是唯一事实源，内嵌兜底必须与其一致。');
    out('Fix: 修改 Config/intent_config.json 后，同步 intent-config.js DEFAULT_CONFIG 与 IntentConfigLoader.vb BuildDefaultBackendConfig。');
    process.exit(1);
} else {
    out(`[RESULT] CONSISTENT (intent_config.json 为单一事实源, version=${jsonObj.version}).`);
    out('  前端兜底:        intent-config.js DEFAULT_CONFIG 一致');
    out('  后端兜底:        IntentConfigLoader.vb 一致');
    out('  前端消费方:      无内嵌关键词字面量');
    process.exit(0);
}
