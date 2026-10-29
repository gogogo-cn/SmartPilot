/**
 * ============================================================
 * Tests/tools-consistency.test.js — MCP 工具元数据一致性校验
 * ============================================================
 * 对齐「工具元数据统一」目标：excel_tools.json 是唯一数据源。
 *   ① 后端 VB 注册：Services/Mcp/ExcelMcpTools.vb（兜底，元数据须与 JSON 一致）
 *   ② 配置文件：Config/excel_tools.json（权威：name/layer/description/parameters +
 *      displayName/intentKeywords/operationType/isWrite/isDangerous）
 *   ③ 前端：tool-registry.js 不得再硬编码显示名映射（显示名由桥接下发）
 *
 * 运行：node --test Tests/tools-consistency.test.js（或随 run-tests.js 聚合）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const VB_FILE = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.vb');
const META_FILE = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.ToolDefinitions.Metadata.vb');
const JSON_FILE = path.join(ROOT, 'Config', 'excel_tools.json');
const JS_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'tool-registry.js');
const PPT_VB_FILE = path.join(ROOT, 'Services', 'Mcp', 'PptMcpTools.vb');
const PPT_VB_PROJECT_FILE = path.join(ROOT, 'Services', 'Mcp', 'PptMcpTools.Project.vb');
const PPT_JSON_FILE = path.join(ROOT, 'Config', 'ppt_tools.json');

/** ① 提取后端 VB 注册的工具名（定义拆到 ExcelMcpTools.vb 及 ToolDefinitions.*.vb partial，须一并扫描） */
function extractVbTools() {
    const SRC_DIR = path.join(ROOT, 'Services', 'Mcp');
    const files = fs.readdirSync(SRC_DIR)
        .filter(f => /^ExcelMcpTools(\.ToolDefinitions\.\w+)?\.vb$/.test(f))
        .map(f => path.join(SRC_DIR, f));
    const re = /tools\.Add\(New McpTool\("([^"]+)"/g;
    const names = [];
    for (const f of files) {
        const src = fs.readFileSync(f, 'utf8');
        let m;
        while ((m = re.exec(src)) !== null) names.push(m[1]);
    }
    return [...new Set(names)].sort();
}

/** ② 读取 excel_tools.json */
function loadJson() {
    const raw = fs.readFileSync(JSON_FILE, 'utf8');
    return JSON.parse(raw.replace(/^\uFEFF/, ''));
}

/** 解析 VB GetToolMetadata 的 Select Case 显式覆盖（ApplyExplicitToolMetadata 在 Metadata.vb） */
function parseVbMetaOverrides() {
    const src = fs.readFileSync(META_FILE, 'utf8');
    const start = src.indexOf('Sub ApplyExplicitToolMetadata');
    const end = src.indexOf('End Sub', start);
    const block = src.slice(start, end);
    const meta = {};
    const segments = block.split('Case "').slice(1);
    for (const seg of segments) {
        const name = seg.slice(0, seg.indexOf('"'));
        const body = seg.slice(seg.indexOf('"') + 1);
        const entry = { displayName: null, keywords: null, category: null, isWrite: null, isDangerous: null };
        const dn = body.match(/meta\.DisplayName = "([^"]*)"/);
        if (dn) entry.displayName = dn[1];
        const kw = body.match(/meta\.Keywords = New List\(Of String\)\((\{.*?\})\)/);
        if (kw) entry.keywords = [...kw[1].matchAll(/"([^"]+)"/g)].map(x => x[1]);
        const cat = body.match(/meta\.Category = CapabilityCategory\.(\w+)/);
        if (cat) entry.category = cat[1];
        const w = body.match(/meta\.IsWrite = (True|False)/);
        if (w) entry.isWrite = w[1] === 'True';
        const d = body.match(/meta\.IsDangerous = (True|False)/);
        if (d) entry.isDangerous = d[1] === 'True';
        meta[name] = entry;
    }
    return meta;
}

/** 复制 VB 启发式默认分类逻辑（顺序敏感，ElseIf 链） */
function defaultCategory(name) {
    if (/_read_|_get_|_list_|_search|_browse_/.test(name)) return 'Read';
    if (/_aggregate|_date_group|_remove_duplicates|_clean_data|_join_tables/.test(name)) return 'Compute';
    if (/_write_|_set_|_fill_|_add_|_insert_|_delete_|_remove_|_clear_|_clean_|_merge_cells|_unmerge_cells/.test(name)) return 'Write';
    if (/_format_|_conditional_format|_autofit/.test(name)) return 'Format';
    if (/_pivot_|_sort_|_filter/.test(name)) return 'Analyze';
    if (/_split_|_merge_workbooks|_merge_sheets/.test(name)) return 'Transform';
    if (/_import_/.test(name)) return 'Import';
    if (/_export_/.test(name)) return 'Export';
    if (/_sheet|_workbook/.test(name)) return 'Manage';
    if (/_chart|_display_html/.test(name)) return 'Visualize';
    if (/_clipboard_|_screenshot|_ocr_/.test(name)) return 'Utility';
    return 'Utility';
}
function defaultIsWrite(cat) {
    return ['Write', 'Format', 'Transform', 'Import', 'Export', 'Visualize'].includes(cat);
}
function defaultDanger(name) {
    // 与 ExcelMcpTools.GetHardcodedToolMetadata / excel_tools.json 危险集合对齐
    return /_clear_sheet|_delete_sheet|_delete_row|_delete_column|_clear_range|_split_by_column|_merge_workbooks|_execute_vba|_run_python_script|_clean_data|_remove_duplicates/.test(name);
}

/** 计算 VB 兜底元数据（启发式 + 显式覆盖） */
function computeVbMeta(name, overrides) {
    const md = overrides[name] || {};
    const cat = md.category || defaultCategory(name);
    return {
        displayName: md.displayName || name,
        keywords: md.keywords || [],
        operationType: cat,
        isWrite: md.isWrite !== null && md.isWrite !== undefined ? md.isWrite : defaultIsWrite(cat),
        isDangerous: md.isDangerous !== null && md.isDangerous !== undefined ? md.isDangerous : defaultDanger(name)
    };
}

function diffOnly(a, b) {
    return a.filter(x => !b.includes(x));
}

test('MCP 一致性: ①后端 VB 名称 = ②excel_tools.json 名称', () => {
    const vb = extractVbTools();
    const json = loadJson();
    const jsonNames = (json.tools || []).map(t => t && t.name).filter(Boolean).sort();

    assert.ok(vb.length > 0, 'VB 应提取到工具');
    assert.ok(jsonNames.length > 0, 'excel_tools.json 应提取到工具');
    assert.strictEqual(vb.length, jsonNames.length,
        `工具数量不一致: VB=${vb.length} JSON=${jsonNames.length}\n` +
        `VB 有而 JSON 缺: ${diffOnly(vb, jsonNames).join(', ') || '(无)'}\n` +
        `JSON 有而 VB 缺: ${diffOnly(jsonNames, vb).join(', ') || '(无)'}`);
    assert.deepStrictEqual(vb, jsonNames, '①后端 VB 与 ②excel_tools.json 名称集合不一致');
    console.log(`  [ok] VB 与 JSON 名称一致，共 ${vb.length} 个工具`);
});

test('MCP 一致性: 前端 tool-registry.js 不再硬编码显示名映射', () => {
    const src = fs.readFileSync(JS_FILE, 'utf8');
    assert.ok(!src.includes('EXCEL_TOOL_DISPLAY_NAMES'),
        'tool-registry.js 不应再包含硬编码 EXCEL_TOOL_DISPLAY_NAMES（显示名由 excel_tools.json 经桥接统一下发）');
    console.log('  [ok] 前端无硬编码显示名映射');
});

test('JSON 完整性: excel_tools.json 元数据字段齐全（displayName/operationType/isWrite/isDangerous）', () => {
    const json = loadJson();
    assert.ok(typeof json.version === 'string' && json.version.length > 0, 'excel_tools.json 应有 version');
    assert.ok(typeof json.lastUpdated === 'string' && json.lastUpdated.length > 0, 'excel_tools.json 应有 lastUpdated');

    const validOps = ['Read', 'Write', 'Format', 'Analyze', 'Transform', 'Export', 'Import', 'Manage', 'Visualize', 'Utility', 'Compute'];
    const missing = [];
    const badOps = [];
    for (const t of json.tools) {
        if (!t.name) { missing.push('(无name)'); continue; }
        if (!t.displayName || !t.operationType || t.isWrite === undefined || t.isDangerous === undefined) {
            missing.push(t.name);
        }
        if (t.operationType && !validOps.includes(t.operationType)) badOps.push(`${t.name}:${t.operationType}`);
        if (!Array.isArray(t.intentKeywords)) missing.push(`${t.name}:intentKeywords`);
    }
    assert.deepStrictEqual(missing, [], `JSON 元数据不完整: ${missing.join(', ')}`);
    assert.deepStrictEqual(badOps, [], `JSON operationType 非法: ${badOps.join(', ')}`);
    console.log(`  [ok] ${json.tools.length} 个工具元数据完整，operationType 合法`);
});

test('MCP 一致性: excel_tools.json 与后端 VB 兜底在"可由规则推导的字段"一致（富元数据以 JSON 为唯一权威）', () => {
    const json = loadJson();
    const overrides = parseVbMetaOverrides();
    const vbNames = extractVbTools();

    // 名称一致性（覆盖全工具）：VB 兜底是 JSON 缺失时的备用，命名不得漂移
    const missingInVb = vbNames.filter(n => !json.tools.some(t => t.name === n))
        .concat(json.tools.map(t => t.name).filter(n => !vbNames.includes(n)));
    assert.deepStrictEqual(missingInVb, [],
        `VB 兜底工具名与 excel_tools.json 不一致: ${missingInVb.join(', ') || '(无)'}`);

    // 校验 VB Select Case 覆盖了全部工具
    const noOverride = vbNames.filter(n => !overrides[n]);
    assert.deepStrictEqual(noOverride, [], `VB GetToolMetadata 未覆盖的工具: ${noOverride.join(', ')}`);

    // 富元数据（displayName/intentKeywords）只维护在 excel_tools.json（单一数据源），
    // 兜底不重复维护、不做全字段复刻；此处只校验"可由规则推导"的字段与兜底一致，防兜底漂移。
    const diffs = [];
    for (const t of json.tools) {
        const vbMeta = computeVbMeta(t.name, overrides);
        if (t.operationType !== vbMeta.operationType) diffs.push(`${t.name}.operationType: JSON=${t.operationType} VB兜底=${vbMeta.operationType}`);
        if (t.isWrite !== vbMeta.isWrite) diffs.push(`${t.name}.isWrite: JSON=${t.isWrite} VB兜底=${vbMeta.isWrite}`);
        if (t.isDangerous !== vbMeta.isDangerous) diffs.push(`${t.name}.isDangerous: JSON=${t.isDangerous} VB兜底=${vbMeta.isDangerous}`);
    }
    assert.deepStrictEqual(diffs, [], `JSON 与 VB 兜底（可推导字段）不一致:\n${diffs.join('\n')}`);
    console.log(`  [ok] ${json.tools.length} 个工具：名称与可推导字段与 VB 兜底一致，富元数据以 JSON 为唯一权威`);
});

// ============================================================
// PPT MCP 工具一致性（Config/ppt_tools.json ↔ Services/Mcp/PptMcpTools.vb）
// ============================================================

/** 读取 ppt_tools.json */
function loadPptJson() {
    const raw = fs.readFileSync(PPT_JSON_FILE, 'utf8');
    return JSON.parse(raw.replace(/^\uFEFF/, ''));
}

/** 复制 VB ToolNameToMethodName：ppt_analyze_data → ExecuteAnalyzeData */
function pptToolToMethodName(toolName) {
    let short = toolName;
    if (short.startsWith('ppt_')) short = short.slice(4);
    return 'Execute' + short.split('_').filter(Boolean)
        .map(p => p[0].toUpperCase() + p.slice(1)).join('');
}

/** 提取 PptMcpTools.vb 中定义的 Execute 方法名（partial 拆到 PptMcpTools.Project.vb，须一并扫描） */
function extractPptExecuteMethods() {
    const sources = [PPT_VB_FILE, PPT_VB_PROJECT_FILE]
        .filter(f => fs.existsSync(f))
        .map(f => fs.readFileSync(f, 'utf8'))
        .join('\n');
    const re = /(?:Private|Public|Protected) (?:Shared )?Function (Execute\w+)\s*\(/g;
    return [...sources.matchAll(re)].map(m => m[1]);
}

/** 复制 VB GetHardcodedToolMetadata 启发式（顺序敏感） */
function pptDefaultMeta(name) {
    let cat = 'Utility', isWrite = false, isDangerous = false;
    if (/_get_|_browse_|_list/.test(name)) cat = 'Read';
    else if (/_analyze_/.test(name)) cat = 'Compute';
    else if (/_generate|_import_|_project_|_add_|_update|_remove|_build/.test(name)) {
        cat = 'Write';
        isWrite = true;
        if (/_remove/.test(name)) isDangerous = true;
    }
    return { operationType: cat, isWrite, isDangerous };
}

test('PPT 一致性: ppt_tools.json 工具名全部以 ppt_ 开头且元数据齐全', () => {
    const json = loadPptJson();
    assert.ok(typeof json.version === 'string' && json.version.length > 0, 'ppt_tools.json 应有 version');
    assert.ok(typeof json.lastUpdated === 'string' && json.lastUpdated.length > 0, 'ppt_tools.json 应有 lastUpdated');
    assert.ok(Array.isArray(json.tools) && json.tools.length > 0, 'ppt_tools.json 应包含工具');

    const validOps = ['Read', 'Write', 'Format', 'Analyze', 'Transform', 'Export', 'Import', 'Manage', 'Visualize', 'Utility', 'Compute'];
    const bad = [];
    for (const t of json.tools) {
        if (!/^ppt_/.test(t.name || '')) bad.push(`${t.name}: 不以 ppt_ 开头`);
        if (!t.displayName || !t.operationType || t.isWrite === undefined || t.isDangerous === undefined) {
            bad.push(`${t.name}: 缺 displayName/operationType/isWrite/isDangerous`);
        }
        if (t.operationType && !validOps.includes(t.operationType)) bad.push(`${t.name}:operationType=${t.operationType}`);
        if (!Array.isArray(t.intentKeywords)) bad.push(`${t.name}: 缺 intentKeywords 数组`);
        if (typeof t.layer !== 'number') bad.push(`${t.name}: 缺 layer`);
    }
    assert.deepStrictEqual(bad, [], `ppt_tools.json 元数据问题: ${bad.join(', ') || '(无)'}`);
    console.log(`  [ok] ${json.tools.length} 个 PPT 工具：元数据齐全、前缀规范`);
});

test('PPT 一致性: ppt_tools.json 每个工具在 PptMcpTools.vb 有对应 Execute 方法（反射路由兜底）', () => {
    const json = loadPptJson();
    const methods = extractPptExecuteMethods();
    const names = (json.tools || []).map(t => t.name).filter(Boolean);
    const missing = names.filter(n => !methods.includes(pptToolToMethodName(n)));
    assert.deepStrictEqual(missing, [], `缺少 Execute 方法的 PPT 工具: ${missing.join(', ') || '(无)'}`);
    console.log(`  [ok] ${names.length} 个 PPT 工具均有 Execute 方法（${pptToolToMethodName(names[0])} ...）`);
});

test('PPT 一致性: ppt_tools.json 与 VB GetHardcodedToolMetadata 兜底在可推导字段一致', () => {
    const json = loadPptJson();
    const diffs = [];
    for (const t of json.tools) {
        const vbMeta = pptDefaultMeta(t.name);
        if (t.operationType !== vbMeta.operationType) diffs.push(`${t.name}.operationType: JSON=${t.operationType} VB兜底=${vbMeta.operationType}`);
        if (t.isWrite !== vbMeta.isWrite) diffs.push(`${t.name}.isWrite: JSON=${t.isWrite} VB兜底=${vbMeta.isWrite}`);
        if (t.isDangerous !== vbMeta.isDangerous) diffs.push(`${t.name}.isDangerous: JSON=${t.isDangerous} VB兜底=${vbMeta.isDangerous}`);
    }
    assert.deepStrictEqual(diffs, [], `ppt_tools.json 与 VB 兜底（可推导字段）不一致:\n${diffs.join('\n')}`);
    console.log(`  [ok] ${json.tools.length} 个 PPT 工具：可推导字段与 VB 兜底一致`);
});

test('PPT 一致性: 分步构建工具集完整（project 生命周期 + 页型添加）', () => {
    const json = loadPptJson();
    const names = (json.tools || []).map(t => t.name);

    const LIFECYCLE = ['ppt_project_start', 'ppt_project_list', 'ppt_project_update', 'ppt_project_remove', 'ppt_project_build'];
    const ADDER = ['ppt_add_cover_slide', 'ppt_add_text_slide', 'ppt_add_table_slide', 'ppt_add_chart_slide',
        'ppt_add_kpi_slide', 'ppt_add_structure_slide', 'ppt_add_shapes_slide', 'ppt_add_picture_slide', 'ppt_add_smartart_slide'];
    const missing = [...LIFECYCLE, ...ADDER].filter(n => !names.includes(n));
    assert.deepStrictEqual(missing, [], `分步构建工具缺失: ${missing.join(', ') || '(无)'}`);

    // 需要 projectId 的工具必须有必填参数 projectId
    const projectTools = [...LIFECYCLE.slice(1), ...ADDER]; // 除 start 外
    const bad = [];
    for (const t of json.tools) {
        if (!projectTools.includes(t.name)) continue;
        const params = t.parameters || [];
        const proj = params.find(p => p.name === 'projectId');
        if (!proj) bad.push(`${t.name}: 缺 projectId 参数`);
        else if (!proj.required) bad.push(`${t.name}: projectId 未标记必填`);
    }
    assert.deepStrictEqual(bad, [], `projectId 契约问题: ${bad.join(', ') || '(无)'}`);

    // 页面添加工具的 Execute 方法名映射（与后端 ToolNameToMethodName 同构）
    const methods = extractPptExecuteMethods();
    const noMethod = [...ADDER, 'ppt_project_start', 'ppt_project_build'].filter(n => !methods.includes(pptToolToMethodName(n)));
    assert.deepStrictEqual(noMethod, [], `缺少 Execute 方法: ${noMethod.join(', ') || '(无)'}`);

    console.log(`  [ok] 分步构建工具集 ${LIFECYCLE.length + ADDER.length} 个：projectId 契约与 Execute 映射齐全`);
});

// 测试结束后强制退出（与其它测试文件一致，防止定时器挂起进程）
if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
