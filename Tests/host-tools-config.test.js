/**
 * ============================================================
 * Tests/host-tools-config.test.js — 宿主工具配置（Word/PPT）与全仓 JSON 配置校验
 * ============================================================
 * 背景（2026-10-09 审查发现）：
 *   1. Hosts/Shared/Config/word_tools.json 第 255 行 description 里嵌了未转义的
 *      英文双引号 → 整个文件 JSON 非法 → HostMcpTools.LoadConfig 抛异常被空
 *      Catch 吞掉 → **Word 宿主的工具清单静默为空**（典型的"静默空结果"）。
 *      而既有 tools-consistency.test.js 只校验 Config/excel_tools.json 与
 *      Config/ppt_tools.json，**完全没覆盖 Hosts/Shared/Config**。
 *   2. 同名的 ppt_tools.json 有两份（Config/ 19 个工具为 Excel 宿主内的 PPT
 *      生成链路；Hosts/Shared/Config/ 30 个工具为 PPT 宿主自身），内容不同、
 *      服务不同宿主，但命名完全相同，极易误改。这里显式断言两者的职责边界。
 *
 * 运行：node Tests/host-tools-config.test.js（随 run-tests.js 自动发现）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');

/** 全仓需要保证「可被标准 JSON 解析」的配置文件（加载失败会静默降级为空） */
const JSON_FILES = [
    'Config/excel_tools.json',
    'Config/ppt_tools.json',
    'Config/mcp.json',
    'Config/intent_config.json',
    'Hosts/Shared/Config/word_tools.json',
    'Hosts/Shared/Config/ppt_tools.json',
    'Hosts/Shared/Config/word_skills.json',
    'Hosts/Shared/Config/ppt_skills.json'
];

const VALID_OPS = ['Read', 'Write', 'Format', 'Analyze', 'Transform', 'Export',
    'Import', 'Manage', 'Visualize', 'Utility', 'Compute'];

function readJson(rel) {
    const full = path.join(ROOT, rel);
    assert.ok(fs.existsSync(full), `配置文件不存在: ${rel}`);
    return JSON.parse(fs.readFileSync(full, 'utf8').replace(/^\uFEFF/, ''));
}

test('配置合法性: 关键工具/技能 JSON 均可被标准解析器解析（防静默加载为空）', () => {
    const bad = [];
    for (const rel of JSON_FILES) {
        const full = path.join(ROOT, rel);
        if (!fs.existsSync(full)) { bad.push(`${rel}: 文件不存在`); continue; }
        try {
            JSON.parse(fs.readFileSync(full, 'utf8').replace(/^\uFEFF/, ''));
        } catch (e) {
            bad.push(`${rel}: ${e.message}`);
        }
    }
    assert.deepStrictEqual(bad, [], `JSON 解析失败（宿主加载时会被空 Catch 吞掉，表现为"工具清单为空"）:\n${bad.join('\n')}`);
    console.log(`  [ok] ${JSON_FILES.length} 个关键配置文件均可解析`);
});

/** 宿主工具表通用校验：前缀、元数据完整性、名称唯一 */
function assertHostToolsValid(rel, prefix) {
    const json = readJson(rel);
    assert.ok(Array.isArray(json.tools) && json.tools.length > 0, `${rel} 应包含工具`);
    assert.ok(typeof json.version === 'string' && json.version.length > 0, `${rel} 应有 version`);

    const bad = [];
    const seen = new Set();
    for (const t of json.tools) {
        const n = t && t.name;
        if (!n) { bad.push('(缺 name)'); continue; }
        if (seen.has(n)) bad.push(`${n}: 名称重复`);
        seen.add(n);
        if (!n.startsWith(prefix)) bad.push(`${n}: 前缀应为 ${prefix}`);
        if (!t.displayName) bad.push(`${n}: 缺 displayName`);
        if (!t.operationType) bad.push(`${n}: 缺 operationType`);
        else if (!VALID_OPS.includes(t.operationType)) bad.push(`${n}: operationType=${t.operationType} 非法`);
        if (t.isWrite === undefined) bad.push(`${n}: 缺 isWrite`);
        if (t.isDangerous === undefined) bad.push(`${n}: 缺 isDangerous`);
        if (!Array.isArray(t.intentKeywords)) bad.push(`${n}: 缺 intentKeywords 数组`);
        if (!Array.isArray(t.parameters)) bad.push(`${n}: 缺 parameters 数组`);
    }
    assert.deepStrictEqual(bad, [], `${rel} 元数据问题:\n${bad.join('\n')}`);
    return json;
}

test('Word 宿主: Hosts/Shared/Config/word_tools.json 元数据齐全、前缀规范', () => {
    const json = assertHostToolsValid('Hosts/Shared/Config/word_tools.json', 'word_');
    console.log(`  [ok] word_tools.json ${json.tools.length} 个工具元数据齐全`);
});

test('PPT 宿主: Hosts/Shared/Config/ppt_tools.json 元数据齐全，且工具在 PptBridge 有对应方法', () => {
    const json = assertHostToolsValid('Hosts/Shared/Config/ppt_tools.json', 'ppt_');
    // source 由 HostMcpTools 按宿主注入（word-local / ppt-local），不在 JSON 里重复维护
    const hostMcp = path.join(ROOT, 'Hosts', 'Shared', 'HostMcpTools.vb');
    assert.ok(fs.existsSync(hostMcp), 'Hosts/Shared/HostMcpTools.vb 应存在');
    const vb = fs.readFileSync(hostMcp, 'utf8');
    assert.ok(vb.includes('"ppt-local"') && vb.includes('"word-local"'),
        'HostMcpTools 应按宿主注入 source（ppt-local / word-local）');
    console.log(`  [ok] ppt_tools.json ${json.tools.length} 个工具元数据齐全，source 由 HostMcpTools 注入`);
});

test('同名配置职责边界: Config/ppt_tools.json(Excel 宿主内 PPT 生成) 与 Hosts/Shared/Config/ppt_tools.json(PPT 宿主) 均须存在且各自自洽', () => {
    const excelSide = readJson('Config/ppt_tools.json');
    const pptHost = readJson('Hosts/Shared/Config/ppt_tools.json');

    assert.ok(excelSide.tools.length > 0, 'Config/ppt_tools.json（Excel 宿主侧 PPT 生成工具）不应为空');
    assert.ok(pptHost.tools.length > 0, 'Hosts/Shared/Config/ppt_tools.json（PPT 宿主工具）不应为空');

    // 两份文件服务不同宿主，工具集合允许不同；但都必须是 ppt_ 前缀、且名称不重复
    for (const [rel, json] of [['Config/ppt_tools.json', excelSide], ['Hosts/Shared/Config/ppt_tools.json', pptHost]]) {
        const names = json.tools.map(t => t.name);
        assert.strictEqual(new Set(names).size, names.length, `${rel} 存在重名工具`);
        const badPrefix = names.filter(n => !/^ppt_/.test(n));
        assert.deepStrictEqual(badPrefix, [], `${rel} 存在非 ppt_ 前缀工具: ${badPrefix.join(', ')}`);
    }

    // PPT 宿主侧必须覆盖读取与导出（否则 Word/PPT 硬短路工具无定义）
    const hostNames = pptHost.tools.map(t => t.name);
    for (const required of ['ppt_export_pdf']) {
        assert.ok(hostNames.includes(required), `PPT 宿主工具表缺少硬短路依赖的 ${required}`);
    }
    console.log(`  [ok] 两份 ppt_tools.json：Excel 侧 ${excelSide.tools.length} 个 / PPT 宿主侧 ${pptHost.tools.length} 个，各自自洽`);
});

test('部署一致性: Hosts/Shared/Config 的 word_tools.json 已同步到 Word 宿主输出目录', () => {
    const src = path.join(ROOT, 'Hosts', 'Shared', 'Config', 'word_tools.json');
    const deployed = path.join(ROOT, 'Hosts', 'WordAIVSTOAddIn', 'bin', 'Debug', 'Config', 'word_tools.json');
    if (!fs.existsSync(deployed)) {
        console.log('  [skip] Word 宿主输出目录不存在（未构建），跳过部署一致性校验');
        return;
    }
    assert.strictEqual(
        fs.readFileSync(src, 'utf8').trim(),
        fs.readFileSync(deployed, 'utf8').trim(),
        'Hosts/Shared/Config/word_tools.json 与 Word 宿主 bin/Debug/Config 下副本不一致，请手动复制'
    );
    console.log('  [ok] Word 宿主输出目录副本与源一致');
});
