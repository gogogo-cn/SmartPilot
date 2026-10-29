/**
 * ============================================================
 * Tests/ppt-templates-consistency.test.js — PPT 模板配置一致性校验
 * ============================================================
 * 对齐「JSON 单一数据源」目标：Config/ppt_templates.json 是 PPT 主题/版式的权威源。
 *   ① 配置文件：Config/ppt_templates.json（主题色 + 版式注册表）
 *   ② 后端：Services/PptGenerationService.vb（SUPPORTED_SLIDE_TYPES 白名单须覆盖全部版式）
 *
 * 运行：node --test Tests/ppt-templates-consistency.test.js（或随 run-tests.js 聚合）
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const JSON_FILE = path.join(ROOT, 'Config', 'ppt_templates.json');
const VB_FILE = path.join(ROOT, 'Services', 'PptGenerationService.vb');
const MCP_ADVANCED_FILE = path.join(ROOT, 'Services', 'Mcp', 'ExcelMcpTools.Advanced.Aggregate.vb');
const DESIGN_ENGINE_FILE = path.join(ROOT, 'Services', 'PptDesignEngine.vb');
const EXCEL_BRIDGE_FILE = path.join(ROOT, 'ExcelBridge.vb');
const CODE_PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'code-pipeline.js');
const PYTHON_PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'python-pipeline.js');
const SKILL_PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill-pipeline.js');
// 【skill-pipeline 拆分适配】specMode 已迁至 skill-spec-execute.js，源码断言需拼接读取
const SKILL_SPEC_EXEC_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill', 'skill-spec-execute.js');
// 【常量收敛适配】正则/注册表已收敛到 skill-constants.js 单一数据源，字符串断言须并入该文件
const SKILL_CONSTANTS_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'skill', 'skill-constants.js');
const DOC_PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'doc-pipeline.js');
const TEXT_PIPELINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'pipelines', 'text-pipeline.js');
const TEMPLATE_ENGINE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'vba-template-engine.js');
const RENDERER_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'ui', 'chat-renderer.js');
const JSON_RENDERER_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'ui', 'renderers', 'json.js');
const ORCHESTRATOR_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'chat-orchestrator.js');
const SESSION_STORE_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'session-store.js');
const INTENT_ANALYZER_FILE = path.join(ROOT, 'UI', 'js', 'chat-v2', 'services', 'intent-analyzer.js');
const INDEX_HTML_FILE = path.join(ROOT, 'UI', 'index.html');
const SKILL_MGMT_JS_FILE = path.join(ROOT, 'UI', 'js', 'app-skill-mgmt.js');
const SKILLS_DIR = path.join(ROOT, 'Skills');

const REQUIRED_COLOR_KEYS = ['primary', 'secondary', 'accent', 'bg', 'light', 'text', 'gray'];

function loadJson() {
    const raw = fs.readFileSync(JSON_FILE, 'utf8');
    return JSON.parse(raw.replace(/^\uFEFF/, ''));
}

/** 【PptGenerationService 分部拆分适配】实现已拆到 PptGenerationService.*.vb partial，须一并拼接扫描 */
function loadPptVbAll() {
    const dir = path.join(ROOT, 'Services');
    const files = fs.readdirSync(dir)
        .filter(f => /^PptGenerationService(\.\w+)?\.vb$/.test(f))
        .map(f => path.join(dir, f))
        .sort();
    return files.map(f => fs.readFileSync(f, 'utf8')).join('\n');
}

/** 提取后端 VB SUPPORTED_SLIDE_TYPES 白名单 */
function extractSupportedSlideTypes() {
    const src = fs.readFileSync(VB_FILE, 'utf8');
    const start = src.indexOf('SUPPORTED_SLIDE_TYPES');
    const end = src.indexOf('}', start);
    const block = src.slice(start, end);
    const types = [...block.matchAll(/"([^"]+)"/g)].map(m => m[1]);
    return [...new Set(types)];
}

/** 提取设计引擎 BuildDesign 中 JObject 字面量声明的版式类型（{ "type", "xxx" }） */
function extractDesignEngineSlideTypes() {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    return [...src.matchAll(/\{\s*"type",\s*"([^"]+)"\s*\}/g)].map(m => m[1]);
}

/** 提取前端 code-pipeline.js filterSlidesByLayout 的结构页 alwaysKeep 集合 */
function extractAlwaysKeepTypes() {
    const src = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const start = src.indexOf('alwaysKeep = new Set(');
    const end = src.indexOf(');', start);
    const block = src.slice(start, end);
    return [...block.matchAll(/'([^']+)'/g)].map(m => m[1]);
}

/** 提取前端 code-pipeline.js detectPptTheme 规则里的主题 id */
function extractFrontendThemeIds() {
    const src = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const start = src.indexOf('detectPptTheme');
    const end = src.indexOf('return \'business_blue\';', start);
    const block = src.slice(start, end);
    return [...block.matchAll(/\{ id: '([^']+)'/g)].map(m => m[1]);
}

/** 提取前端 vba-template-engine.js COLOR_THEMES_BACKUP 的主题 id（键名） */
function extractBackupThemeIds() {
    const src = fs.readFileSync(TEMPLATE_ENGINE_FILE, 'utf8');
    const start = src.indexOf('COLOR_THEMES_BACKUP = {');
    const end = src.indexOf('};', start);
    const block = src.slice(start, end);
    return [...block.matchAll(/'([^']+)': \{/g)].map(m => m[1]);
}

test('PPT 模板: ppt_templates.json 基础结构完整', () => {
    const json = loadJson();
    assert.ok(typeof json.version === 'number' && json.version > 0, '应有 version');
    assert.ok(typeof json.lastUpdated === 'string' && json.lastUpdated.length > 0, '应有 lastUpdated');
    assert.ok(typeof json.defaultTheme === 'string' && json.defaultTheme.length > 0, '应有 defaultTheme');
    assert.ok(Array.isArray(json.themes) && json.themes.length > 0, '应有 themes 数组');
    assert.ok(Array.isArray(json.slideLayouts) && json.slideLayouts.length > 0, '应有 slideLayouts 数组');
    console.log(`  [ok] 结构完整，${json.themes.length} 个主题，${json.slideLayouts.length} 种版式`);
});

test('PPT 模板: defaultTheme 必须存在于 themes 中', () => {
    const json = loadJson();
    const ids = json.themes.map(t => t.id);
    assert.ok(ids.includes(json.defaultTheme),
        `defaultTheme=${json.defaultTheme} 不在 themes 中，可用: ${ids.join(', ')}`);
    console.log(`  [ok] 默认主题 ${json.defaultTheme} 存在`);
});

test('PPT 模板: 每个主题 id 唯一且颜色字段齐全（7 色 RGB 数组）', () => {
    const json = loadJson();
    const ids = json.themes.map(t => t.id);
    assert.strictEqual(new Set(ids).size, ids.length, `主题 id 重复: ${ids.join(', ')}`);

    const problems = [];
    for (const t of json.themes) {
        if (!t.id) { problems.push('(无id)'); continue; }
        if (!t.name) problems.push(`${t.id}:缺name`);
        if (!t.font) problems.push(`${t.id}:缺font`);
        if (!t.colors || typeof t.colors !== 'object') { problems.push(`${t.id}:缺colors`); continue; }
        for (const key of REQUIRED_COLOR_KEYS) {
            const rgb = t.colors[key];
            if (!Array.isArray(rgb) || rgb.length !== 3) {
                problems.push(`${t.id}.${key}:非RGB数组`);
            } else {
                for (const v of rgb) {
                    if (typeof v !== 'number' || v < 0 || v > 255) problems.push(`${t.id}.${key}:值越界(${v})`);
                }
            }
        }
    }
    assert.deepStrictEqual(problems, [], `主题配置问题: ${problems.join(', ')}`);
    console.log(`  [ok] ${json.themes.length} 个主题颜色字段完整`);
});

test('PPT 模板: slideLayouts 类型须被后端 SUPPORTED_SLIDE_TYPES 白名单覆盖', () => {
    const json = loadJson();
    const vbTypes = extractSupportedSlideTypes();
    assert.ok(vbTypes.length > 0, '后端应提取到 SUPPORTED_SLIDE_TYPES');

    const layoutTypes = json.slideLayouts.map(l => l && l.type).filter(Boolean);
    const missing = layoutTypes.filter(t => !vbTypes.includes(t));
    assert.deepStrictEqual(missing, [],
        `版式类型未被后端白名单覆盖: ${missing.join(', ')}（后端支持: ${vbTypes.join(', ')}）`);
    console.log(`  [ok] ${layoutTypes.length} 种版式全部被后端白名单覆盖`);
});

test('PPT 模板: 前端 code-pipeline.js 主题 id 须与 JSON 配置一致', () => {
    const json = loadJson();
    const jsonIds = json.themes.map(t => t.id);
    const feIds = extractFrontendThemeIds();
    assert.ok(feIds.length > 0, '前端应提取到 detectPptTheme 规则主题 id');

    const unknown = feIds.filter(id => !jsonIds.includes(id));
    assert.deepStrictEqual(unknown, [],
        `前端主题 id 不在 JSON 配置中: ${unknown.join(', ')}（JSON: ${jsonIds.join(', ')}）`);
    console.log(`  [ok] 前端 detectPptTheme ${feIds.length} 个主题 id 与 JSON 一致`);
});

test('PPT 模板: 前端 vba-template-engine.js 兜底主题须覆盖 JSON 全部主题', () => {
    const json = loadJson();
    const jsonIds = json.themes.map(t => t.id);
    const backupIds = extractBackupThemeIds();
    assert.ok(backupIds.length > 0, '前端应提取到 COLOR_THEMES_BACKUP 主题 id');

    const missing = jsonIds.filter(id => !backupIds.includes(id));
    assert.deepStrictEqual(missing, [],
        `内置兜底主题缺少 JSON 中的主题: ${missing.join(', ')}（兜底: ${backupIds.join(', ')}）`);
    console.log(`  [ok] 前端兜底主题 ${backupIds.length} 个覆盖 JSON 全部主题`);
});

// ============================================================
// 数据感知设计引擎（PptDesignEngine.vb）一致性
// ============================================================

test('设计引擎: PptDesignEngine.vb 存在且核心入口齐全', () => {
    assert.ok(fs.existsSync(DESIGN_ENGINE_FILE), '缺少 Services/PptDesignEngine.vb');
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    for (const entry of ['AnalyzeData', 'BuildDesign', 'GenerateInsights', 'RecommendChartType', 'DesignFromData', 'AnalyzeReferenceTemplate']) {
        assert.ok(src.includes(`Function ${entry}`) || src.includes(`${entry}(`),
            `设计引擎缺少入口 ${entry}`);
    }
    console.log('  [ok] 设计引擎 6 个核心入口齐全');
});

test('设计引擎: BuildDesign 输出的版式类型 ⊆ 后端 SUPPORTED_SLIDE_TYPES 白名单', () => {
    const vbTypes = extractSupportedSlideTypes();
    const engineTypes = extractDesignEngineSlideTypes();
    assert.ok(engineTypes.length > 0, '设计引擎应提取到版式类型');

    const unknown = engineTypes.filter(t => !vbTypes.includes(t));
    assert.deepStrictEqual(unknown, [],
        `设计引擎输出的版式类型未被后端白名单覆盖: ${unknown.join(', ')}（后端: ${vbTypes.join(', ')}）`);
    console.log(`  [ok] 设计引擎 ${engineTypes.length} 种版式类型全部被后端白名单覆盖`);
});

test('设计引擎: 前端 filterSlidesByLayout 结构页类型 ⊆ 后端 SUPPORTED_SLIDE_TYPES 白名单', () => {
    const vbTypes = extractSupportedSlideTypes();
    const keepTypes = extractAlwaysKeepTypes();
    assert.ok(keepTypes.length > 0, '前端应提取到 alwaysKeep 结构页类型');

    const unknown = keepTypes.filter(t => !vbTypes.includes(t));
    assert.deepStrictEqual(unknown, [],
        `前端结构页类型未被后端白名单覆盖: ${unknown.join(', ')}（后端: ${vbTypes.join(', ')}）`);
    console.log(`  [ok] 前端结构页 ${keepTypes.join('/')} 全部被后端白名单覆盖`);
});

test('设计引擎: 前端 requestPptDesign 契约与后端 RecommendPptDesign 对齐', () => {
    const fe = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const be = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    const vb = fs.readFileSync(VB_FILE, 'utf8');

    assert.ok(fe.includes('RecommendPptDesign'), '前端应调用 RecommendPptDesign');
    assert.ok(be.includes('Public Function RecommendPptDesign'), '后端应暴露 RecommendPptDesign');
    assert.ok(be.includes('result("slides")') && be.includes('result("insights")') && be.includes('result("profile")'),
        '后端应返回 slides/insights/profile 三字段');
    assert.ok(fe.includes('design?.slides') && fe.includes('design?.insights'),
        '前端应消费 slides/insights 字段');

    // 后端 autoDesign 入口：未显式传 slides 时由设计引擎生成
    assert.ok(vb.includes('autoDesign') && vb.includes('PptDesignEngine.DesignFromData'),
        '后端 Generate 应支持 autoDesign 走设计引擎');
    console.log('  [ok] 前端 requestPptDesign ↔ 后端 RecommendPptDesign 契约一致');
});

test('设计引擎: 数据画像逻辑完整（列类型/趋势/时间序列/分类）', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const checks = [
        ['数值列判定', 'cp.Kind = "numeric"'],
        ['日期列判定', 'cp.Kind = "date"'],
        ['分类列判定', 'cp.Kind = "categorical"'],
        ['趋势方向', 'cp.TrendDirection = "up"'],
        ['时间序列标记', 'cp.IsTimeSeries = True'],
        ['主数值列', 'MainNumericColumn'],
        ['主分类列', 'MainCategoryColumn'],
        ['时间序列列', 'TimeSeriesColumn']
    ];
    const missing = checks.filter(([, code]) => !src.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `数据画像缺少逻辑: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 数据画像 ${checks.length} 项逻辑齐全`);
});

test('设计引擎: 图表推荐规则（时间→折线/分类少→饼图/分类多→柱状）', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    assert.ok(src.includes('Return "line"'), '时间序列应推荐折线图');
    assert.ok(src.includes('Return "pie"'), '分类少应推荐饼图');
    assert.ok(src.includes('Return "column"'), '分类多/兜底应推荐柱状图');
    console.log('  [ok] 图表推荐规则齐全');
});

test('设计引擎: ID 列识别与排除（防止 ID 当主角）', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const checks = [
        ['识别函数', 'IsIdLikeColumn'],
        ['IsIdLike 属性', 'IsIdLike As Boolean'],
        ['名称启发式-编号', 'Contains("编号")'],
        ['名称启发式-序号', 'Contains("序号")'],
        ['名称启发式-No', 'StartsWith("no.")'],
        ['数据启发式-高唯一性', 'UniqueCount >= rowCount * 0.9'],
        ['主数值列跳过 ID', 'Not c.IsIdLike'],
        ['主分类列跳过 ID', 'Not cp.IsIdLike'],
        ['KPI 过滤 ID', 'Where(Function(c) Not c.IsIdLike)']
    ];
    const missing = checks.filter(([, code]) => !src.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `ID 列识别/排除缺少逻辑: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] ID 列识别与排除 ${checks.length} 项逻辑齐全`);
});

test('设计引擎: LLM 列理解层（大模型思考"拿什么列"）契约对齐', () => {
    const fe = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const be = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    const engine = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');

    // 前端：LLM 列理解函数 + 下发 columnHints
    assert.ok(fe.includes('async function analyzeColumnsViaLlm'), '前端应实现 analyzeColumnsViaLlm');
    assert.ok(fe.includes('columnHints: columnHints || null'), '前端应下发 columnHints');
    assert.ok(fe.includes('idColumns') && fe.includes('mainMetric'), '前端应识别 ID 列与主角指标');

    // 后端：解析 columnHints + 应用方法
    assert.ok(be.includes('args("columnHints")'), '后端应解析 columnHints');
    assert.ok(engine.includes('Public Shared Sub ApplyColumnHints'), '设计引擎应实现 ApplyColumnHints');
    assert.ok(engine.includes('profile.MainNumericColumn = c'), '应支持主角指标覆盖');
    assert.ok(engine.includes('c.IsIdLike = True'), '应支持 ID 列强制排除');
    console.log('  [ok] LLM 列理解层前端 ↔ 后端契约一致');
});

test('设计引擎: 洞察文案规则（聚合/趋势/占比）', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const checks = [
        ['聚合', '「{0}」合计 {1}'],
        ['趋势', '整体呈上升趋势'],
        ['占比', '出现最多']
    ];
    const missing = checks.filter(([, code]) => !src.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `洞察文案缺少规则: ${missing.map(([n]) => n).join(', ')}`);
    console.log('  [ok] 洞察文案规则齐全');
});

test('设计引擎: 图表数据按 X 轴分类聚合（防止同一分类拆成几十个扇区）', () => {
    const vb = loadPptVbAll();
    // 聚合函数存在
    assert.ok(vb.includes('Private Shared Function AggregateChartRows'), '应实现 AggregateChartRows 聚合函数');
    // 两条写入路径（原生内嵌工作簿 + 宿主 Excel 兜底）都必须使用聚合结果
    // 【签名演进】新增 categories 出参（回传聚合后的分类顺序），调用点实参随之变化。
    const aggCall = 'AggregateChartRows(headers, rows, xIndex, seriesIndexes, maxCategories, agg, categories)';
    const nativeUses = vb.split(aggCall).length - 1 >= 2;
    assert.ok(nativeUses, 'WriteChartData 应使用聚合后的行（原生内嵌工作簿 + 宿主 Excel 兜底两条路径）');
    // 分类图截断到容量上限 + 其他；折线保留全部时间点；散点按明细
    assert.ok(vb.includes('maxCategories = 0') && vb.includes('maxCategories = -1') && vb.includes('maxCategories = If(cap Is Nothing, 8, cap.MaxChartCategories)'),
        '应按图表类型区分聚合策略（折线/散点/分类）');
    assert.ok(vb.includes('d(headers(xIndex)) = "其他"'), '超出 Top N 的分类应合并为「其他」');
    console.log('  [ok] 图表分类聚合逻辑齐全（原生 + Excel 兜底双路径）');
});

test('图表列推断: 缺省 X 轴选非数值/非ID/唯一值适中文本列，缺省系列只取数值列', () => {
    const vb = loadPptVbAll();
    const fns = ['ColumnUniqueCount', 'IsNumericColumn', 'IsIdLikeColumn', 'FindDefaultXIndex'];
    const missing = fns.filter(f => !vb.includes(`Private Shared Function ${f}`));
    assert.deepStrictEqual(missing, [], `缺省列推断函数缺失: ${missing.join(', ')}`);
    assert.ok(vb.includes('numericCols As List(Of Integer)'), '缺省系列应只取数值列');
    assert.ok(vb.includes('u >= 2 AndAlso u <= 24'), 'X 轴候选应要求唯一值适中');
    console.log('  [ok] 缺省列类型推断契约完整（非数值/非ID/唯一值适中 + 数值系列）');
});

test('图表列推断: 缺省数值 X 轴高唯一时 maxCategories=0 不聚合', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('ColumnUniqueCount(headers, rows, xIndex) >= Math.Max(15'),
        '数值 X 高唯一应触发 maxCategories=0');
    console.log('  [ok] 数值高唯一 X 轴不聚合契约完整');
});

test('图表 OLE: 三级降级含 SaveCopyAs+AddOLEObject 文件嵌入路径', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('wb.SaveCopyAs(tempXlsx)'), '应实现 SaveCopyAs 临时 xlsx');
    assert.ok(vb.includes('s.Shapes.AddOLEObject(FileName:=tempXlsx'), '应通过 AddOLEObject 嵌入 PPT');
    assert.ok(vb.includes('chartObj.CopyPicture(2, 1)'), '静态位图兜底路径应保留');
    assert.ok(vb.includes('_olePasteSupported'), '应区分剪贴板 OLE 与文件 OLE 能力缓存');
    console.log('  [ok] 图表三级降级（OLE 复制 → AddOLEObject → 位图）契约完整');
});

test('PPT 生成总结: 返回消息含文件路径/页数页型/主题/比例/图表形式', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('Private Shared Function BuildSlideTypeStats'), '应实现页型统计函数');
    assert.ok(vb.includes('Private Shared Function DescribeChartMode'), '应实现图表形式描述函数');
    assert.ok(vb.includes('summary.AppendLine("文件：" & filePath)'), '总结应含文件路径');
    assert.ok(vb.includes('summary.AppendLine("比例：" & ratioDesc)'), '总结应含比例');
    assert.ok(vb.includes('summary.AppendLine(DescribeChartMode(chartMode, chartPages))'), '总结应含图表形式');
    assert.ok(vb.includes('summary.AppendLine("页数：" & pageCount.ToString()'), '总结应含页数页型统计');
    console.log('  [ok] 生成总结说明契约完整（路径/页数页型/主题/比例/图表形式）');
});

test('PPT 诊断日志: Excel 图表诊断与系列绑定成功日志', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('"Excel 图表诊断: chartType="'), '应输出 Excel 图表诊断行');
    assert.ok(vb.includes('", xField=" & headers(xIndex) & " 系列=["'), '诊断应含 xIndex/xField/系列清单');
    assert.ok(vb.includes('"Excel 图表系列 " & (c + 1).ToString() & " 绑定成功: 行="'),
        '应输出逐系列绑定成功日志');
    assert.ok(vb.includes(', 求和=" & sumVal.ToString("F2")'), '绑定日志应含求和');
    console.log('  [ok] PPT_GEN 诊断日志契约完整');
});

test('图表真实字段: 系列名/分类名显式写入，避免"系列 1/类别 1"占位符', () => {
    const vb = loadPptVbAll();
    // 两条路径都必须按列绘制（xlColumns=-4131），否则行列颠倒
    const plotByCount = (vb.match(/SetSourceData\(dataRange, -4131\)/g) || []).length;
    assert.ok(plotByCount >= 2,
        `原生 + Excel 两条路径都应指定 PlotBy=xlColumns（实际 ${plotByCount} 处）`);
    // 原生路径（WriteChartData）：显式设置系列名与分类标签
    assert.ok(vb.includes('sc.Name = headers(seriesIndexes(c))'),
        'WriteChartData 应显式设置系列名为真实列名');
    assert.ok(vb.includes('sc.XValues = xRange'),
        'WriteChartData 应显式设置分类标签为真实分类值');
    // Excel 兜底路径：同样显式设置
    assert.ok(vb.includes('chartObj.Chart.SeriesCollection(c + 1)'),
        'Excel 兜底路径也应显式设置系列名');
    console.log('  [ok] 图表真实字段名写入契约完整（原生 + Excel 双路径）');
});

test('设计引擎: 生成后自动打开 PPT 供预览', () => {
    const vb = fs.readFileSync(VB_FILE, 'utf8');
    assert.ok(vb.includes('Process.Start(New ProcessStartInfo(filePath)'),
        '生成后应自动打开 PPT 文件');
    assert.ok(vb.includes('UseShellExecute = True'),
        '应使用系统默认程序（PowerPoint）打开');
    assert.ok(vb.includes('PPT 已生成并已打开'),
        '成功消息应提示已打开');
    console.log('  [ok] 生成后自动打开 PPT 逻辑存在');
});

// ============================================================
// 版式多样化 + 数据驱动强调（不再固定套用单一模板）
// ============================================================

test('版式多样化: 设计引擎变体选择函数齐全（按数据特征选变体）', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const fns = ['PickTitleVariant', 'PickKpiVariant', 'PickChartVariant', 'PickAgendaVariant', 'PickTextVariant', 'PickSectionVariant'];
    const missing = fns.filter(f => !src.includes(`Function ${f}`));
    assert.deepStrictEqual(missing, [],
        `设计引擎缺少变体选择函数: ${missing.join(', ')}`);
    console.log(`  [ok] 设计引擎 ${fns.length} 个变体选择函数齐全`);
});

test('版式多样化: BuildDesign 为每类页面注入 variant 字段', () => {
    const src = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const count = (src.match(/\{ "variant", /g) || []).length;
    assert.ok(count >= 8,
        `BuildDesign 应为至少 8 个页面注入 variant（实际 ${count}）`);
    console.log(`  [ok] BuildDesign 注入 ${count} 处 variant 字段`);
});

test('版式多样化: 后端各 builder 变体分支覆盖设计引擎变体 id', () => {
    const vb = loadPptVbAll();
    const checks = [
        ['标题页', 'Case "centered"', 'Case "modern"', 'Case Else \' split'],
        ['KPI 页', 'Case "hero"', 'Case "row"', 'BuildKpiGrid(s, toRelease, kpiList, highlightList, theme)'],
        ['图表页', 'slideVariant.ToLowerInvariant() = "full"', 'slideVariant.ToLowerInvariant() = "insight"'],
        ['文本页', 'Case "cards"', 'Case "numbered"', 'BuildTextBullets(s, toRelease, lines, theme)'],
        ['目录页', 'BuildAgendaCards(s, toRelease, lines, theme)', 'BuildAgendaTimeline(s, toRelease, lines, theme)', 'BuildAgendaNumbered(s, toRelease, lines, theme)'],
        ['章节页', 'Case "centered"', 'Case "bar"'],
        ['结束页', 'Case "centered"', 'Case Else \' thanks']
    ];
    const missing = checks.filter(([name, ...codes]) => codes.some(c => !vb.includes(c)));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `后端变体分支缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 后端 ${checks.length} 类页面变体分支齐全`);
});

test('数据驱动强调: KPI 高亮 + 图表洞察注入与消费', () => {
    const engine = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const vb = loadPptVbAll();
    // 设计引擎注入
    assert.ok(engine.includes('{ "highlight", (i = maxIdx) }'),
        '设计引擎应为最大指标注入 highlight');
    assert.ok(engine.includes('{ "insight", BuildTrendInsight(profile) }'),
        '趋势图应注入洞察');
    assert.ok(engine.includes('{ "insight", BuildCompositionInsight(profile) }'),
        '构成图应注入洞察');
    assert.ok(engine.includes('BuildTrendInsight') && engine.includes('BuildCompositionInsight'),
        '洞察生成函数存在');
    // 后端消费
    assert.ok(vb.includes('highlightList.Add(obj.Value(Of Boolean)("highlight"))'),
        '后端应解析 highlight');
    assert.ok(vb.includes('AddBadge(s, toRelease, "TOP"'),
        '高亮 KPI 应加 TOP 徽标');
    assert.ok(vb.includes('Dim insight As String = spec.Value(Of String)("insight")'),
        '后端应解析 insight');
    assert.ok(vb.includes('AddTextBox(s, toRelease, insight,'),
        '后端应渲染洞察文本');
    console.log('  [ok] 数据驱动强调（KPI 高亮 + 图表洞察）前后端契约一致');
});

test('KPI 防溢出: 前端约束单值 + 后端拆分多值/自动缩字号', () => {
    // 方案 A：前端 LLM 规划 prompt 约束 KPI item 单值化（预测值拆独立项）
    const code = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    assert.ok(code.includes('KPI 项约束'),
        'PPT 规划 prompt 应包含 KPI 项约束');
    assert.ok(code.includes('value 只能是单个数值'),
        'KPI 约束应要求 value 单值化');
    assert.ok(code.includes('把每个预测值拆成独立 item'),
        'KPI 约束应要求预测值拆分为独立 item');
    // 方案 B：后端渲染防御——多值 KPI 拆分为独立项
    const vb = loadPptVbAll();
    assert.ok(vb.includes('单 item value 含多值'),
        '后端应识别多值 KPI 场景');
    assert.ok(vb.includes('StringSplitOptions.RemoveEmptyEntries'),
        '后端拆分多值应剔除空项');
    assert.ok(vb.includes('kpiList = expanded'),
        '后端拆分后应回写 kpiList');
    // 方案 B：AddTextBox 自动缩小字号防溢出
    assert.ok(vb.includes('自动缩小字号以适应文本框'),
        'AddTextBox 应具备自动缩字号防溢出');
    assert.ok(vb.includes('effSize = Math.Max(8, CInt(fontSize * height / neededH))'),
        'AddTextBox 缩字号应设下限 8pt 防过小');
    console.log('  [ok] KPI 防溢出（前端约束 + 后端拆分/缩字号）契约完整');
});

test('版式精致化: 视觉增强原语齐全（圆角/渐变/装饰圆/徽标/垂直居中）', () => {
    // PPT 服务已拆分为多个分部类文件（视觉原语在 PptGenerationService.Shapes.vb），须拼接读取
    const vb = loadPptVbAll();
    const primitives = [
        ['圆角矩形', 'Private Shared Function AddRoundedRect'],
        ['渐变矩形', 'Private Shared Sub AddGradientRect'],
        ['装饰圆', 'Private Shared Sub AddDecorativeCircle'],
        ['徽标', 'Private Shared Sub AddBadge'],
        ['垂直居中文本', 'Private Shared Sub AddCenteredTextBox'],
        ['圆角枚举', 'msoShapeRoundedRectangle = 5'],
        ['渐变 API', 'shape.Fill.TwoColorGradient(angle, 1)'],
        ['透明装饰', 'shape.Fill.Transparency = transparency'],
        ['垂直锚点', 'tb.TextFrame.VerticalAnchor = 3']
    ];
    const missing = primitives.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `视觉增强原语缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 视觉增强原语 ${primitives.length} 项齐全`);
});

// ============================================================
// 公司模板主题导入契约（ImportPptTheme）
// 从已有 .pptx 提取配色 + 字体 → custom_ppt_themes.json → 合并进主题列表
// ============================================================

test('主题导入: 后端 ExcelBridge.vb 暴露 ImportPptTheme 接口', () => {
    const be = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    assert.ok(be.includes('Public Function ImportPptTheme'),
        'ExcelBridge 应暴露 ImportPptTheme COM 接口');
    assert.ok(be.includes('PptGenerationService.ImportThemeFromPptx'),
        'ImportPptTheme 应委托给 PptGenerationService.ImportThemeFromPptx');
    console.log('  [ok] ExcelBridge.ImportPptTheme → PptGenerationService.ImportThemeFromPptx');
});

test('主题导入: 后端导入管线函数齐全（提取/保存/缓存失效/合并）', () => {
    const vb = loadPptVbAll();
    const fns = [
        ['导入入口', 'Public Shared Function ImportThemeFromPptx'],
        ['主题提取', 'Private Shared Function ExtractThemeFromPptx'],
        ['配色提取', 'Private Shared Function ExtractColorFromScheme'],
        ['ID 清洗', 'Private Shared Function SanitizeThemeId'],
        ['保存自定义主题', 'Private Shared Sub SaveCustomTheme'],
        ['缓存失效', 'Public Shared Sub InvalidateThemeCache'],
        ['合并加载', 'Public Shared Function GetMergedTemplatesJson']
    ];
    const missing = fns.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `导入管线缺失函数: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 导入管线 ${fns.length} 个函数齐全`);
});

test('主题导入: 提取逻辑覆盖配色（srgbClr/sysClr 兜底）与字体（Hans 兜底）', () => {
    const vb = loadPptVbAll();
    const checks = [
        ['主题 XML 入口', 'ppt/theme/theme1.xml'],
        ['配色方案', 'clrScheme'],
        ['hex 色值', 'srgbClr'],
        ['系统色兜底', 'sysClr'],
        ['lastClr 兜底', 'lastClr'],
        ['accent1 主色', '"accent1"'],
        ['dk1 文字色', '"dk1"'],
        ['lt1 背景色', '"lt1"'],
        ['字体方案', 'fontScheme'],
        ['东亚字体', 'minorFont'],
        ['Hans 脚本兜底', '"Hans"'],
        ['latin 兜底', 'minor.Element(ns + "latin")']
    ];
    const missing = checks.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `提取逻辑缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 配色/字体提取 ${checks.length} 项逻辑齐全`);
});

test('主题导入: 合并逻辑读 custom_ppt_themes.json 且按 id 去重（大小写不敏感）', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('custom_ppt_themes.json'),
        '应读取/写入 custom_ppt_themes.json');
    assert.ok(vb.includes('StringComparison.OrdinalIgnoreCase'),
        '合并去重应大小写不敏感');
    assert.ok(vb.includes('existing.Replace(t)'),
        '同 id 自定义主题应覆盖内置主题');
    assert.ok(vb.includes('themesArr.Add(t)'),
        '新 id 自定义主题应追加');
    // 保存路径同样按 id 去重
    assert.ok(vb.includes('existing.Replace(newTheme)'),
        'SaveCustomTheme 同 id 应覆盖');
    console.log('  [ok] 合并逻辑（读取 + 去重 + 覆盖/追加）齐全');
});

test('主题导入: 导入校验（文件存在/扩展名/配色缺失）', () => {
    const vb = loadPptVbAll();
    const checks = [
        ['未提供路径', '"未提供 PPT 文件路径"'],
        ['文件不存在', '"文件不存在："'],
        ['扩展名校验', 'Path.GetExtension(pptxPath).ToLowerInvariant()'],
        ['仅支持 pptx', '"仅支持 .pptx 文件'],
        ['配色缺失报错', '"未能从 PPT 提取到配色方案"'],
        ['字体兜底', 'tj.Font = "微软雅黑"']
    ];
    const missing = checks.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `导入校验缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 导入校验 ${checks.length} 项齐全`);
});

test('主题导入: 前端配置卡提供「导入模板」并刷新主题', () => {
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    assert.ok(renderer.includes('导入模板'),
        '配置卡应有「导入模板」按钮');
    assert.ok(renderer.includes('ppt-import-btn'),
        '按钮应带 ppt-import-btn 类');
    assert.ok(renderer.includes('bridge.ImportPptTheme'),
        '应调用后端 ImportPptTheme');
    assert.ok(renderer.includes('engine.loadPptTemplates'),
        '导入后应刷新主题列表');
    assert.ok(renderer.includes('renderPptConfigCard(handle, { ...config, theme: res.themeId })'),
        '导入后应自动选中新主题');
    console.log('  [ok] 前端导入按钮 + 刷新 + 自动选中齐全');
});

test('配置卡推荐设计: 展示数据感知标题而非仅通用类型标签', () => {
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    // 推荐设计 chips 应优先显示数据感知标题（如"销售额与利润概览"），类型名作前缀
    assert.ok(renderer.includes('const typeLabels = {'),
        '应有页面类型标签映射');
    assert.ok(renderer.includes('slideTitle = s?.title ? String(s.title) : \'\''),
        '应读取幻灯片数据感知标题');
    assert.ok(renderer.includes("chip.textContent = slideTitle ? (typeName + '·' + slideTitle) : typeName"),
        'chip 应显示 类型·数据感知标题');
    assert.ok(renderer.includes('chip.title = typeName'),
        'chip 应带类型名工具提示');
    console.log('  [ok] 推荐设计展示数据感知标题契约完整');
});

test('通用多轮消息继承: 存储层/检测路由/上下文复用/LLM 保留原页全链路', () => {
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    const sessionStore = fs.readFileSync(SESSION_STORE_FILE, 'utf8');
    // 存储层：通用 _execSpec API（set/get/clear）
    assert.ok(sessionStore.includes('setExecSpec(turnId, spec)'),
        'session-store 应实现 setExecSpec');
    assert.ok(sessionStore.includes('getLastExecSpec()'),
        'session-store 应实现 getLastExecSpec');
    assert.ok(sessionStore.includes('clearExecSpec(turnId)'),
        'session-store 应实现 clearExecSpec');
    assert.ok(sessionStore.includes('t._execSpec'),
        '应依据 _execSpec 定位上一轮');
    // 检测/路由：orchestrator 通用 FOLLOWUP_PATTERNS + detectFollowUp + _execFollowUp
    assert.ok(orchestrator.includes('FOLLOWUP_PATTERNS'),
        'orchestrator 应有按管道细分的后续消息关键词表');
    assert.ok(orchestrator.includes('ppt: /补充|添加|增加'),
        '关键词应覆盖 PPT 的"补充/添加/增加"');
    assert.ok(orchestrator.includes('章节'),
        '关键词应覆盖"章节"');
    assert.ok(orchestrator.includes('function detectFollowUp'),
        'orchestrator 应实现通用后续消息判定');
    assert.ok(orchestrator.includes('getLastExecSpec'),
        'orchestrator 应调用通用上下文取回');
    assert.ok(orchestrator.includes('context._execFollowUp'),
        '应标记后续消息上下文');
    assert.ok(orchestrator.includes("target === 'ppt' || target === 'code'"),
        'PPT 后续消息应路由回 code-pipeline');
    // 存储：直出路径写入通用 _execSpec（pipeline='ppt'）
    assert.ok(pipeline.includes('sessionStore.setExecSpec(turn.id,'),
        '直出路径应写入 _execSpec');
    assert.ok(pipeline.includes("pipeline: 'ppt'"),
        '_execSpec 应标注来源管道为 ppt');
    // 继承：无选区时从上一轮 _execSpec.data 重建（归一化 skill/text/python 数据快照）
    assert.ok(pipeline.includes('prevExec.spec.data'),
        '应复用上一轮数据上下文');
    assert.ok(pipeline.includes('buildDirectSpecFromExecData(prevExec.spec.data)'),
        '应从 _execSpec.data 归一化重建 spec');
    // 参数继承：主题/推荐设计从 _execSpec.params 取回
    assert.ok(pipeline.includes('prevExec?.spec?.params?.theme'),
        '应继承上一轮主题');
    assert.ok(pipeline.includes('prevExec?.spec?.params?.recommendedSlides'),
        '应继承上一轮推荐设计');
    // LLM 保留原页：提示词含上一版页面结构
    assert.ok(pipeline.includes('previousSlides'),
        'LLM 生成应支持上一版页面结构');
    assert.ok(pipeline.includes('输出完整的新页面列表'),
        '应要求输出完整新页面（保留原页+新增）');
    // 版式感知规划步骤也应保留后续消息的上一版设计（避免冲掉继承结果）
    assert.ok(pipeline.includes('isPptFollowUp ? finalSlides : null'),
        '版式优化时应保留后续消息的当前设计');
    console.log('  [ok] 通用多轮消息继承全链路契约完整');
});

test('VBA 继承: code-pipeline 写入 _execSpec 且后续消息复用上轮代码', () => {
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    // 路由：orchestrator code 关键词表覆盖 VBA 后续诉求（加个按钮/改成弹窗等）
    assert.ok(orchestrator.includes('code: /加个|加一|改成|改为|换成|按钮|弹窗|宏|vba|代码|补充|修改|删除/'),
        'code 关键词表应覆盖 VBA 后续诉求');
    // 存储：代码路径写入通用 _execSpec（pipeline='code'）
    assert.ok(pipeline.includes("pipeline: 'code'"),
        '代码路径应写入 _execSpec 并标注 pipeline=code');
    assert.ok(pipeline.includes('code: extracted || \'\''),
        '_execSpec.result 应存上轮代码');
    assert.ok(pipeline.includes('params: {\n                        lang: actualLang || lang'),
        '_execSpec.params 应记录代码语言');
    // 继承：后续消息从 _execSpec.result.code 取回上轮代码
    assert.ok(pipeline.includes('prevExec.spec.result?.code'),
        '后续消息应从 _execSpec.result.code 取回上轮代码');
    // 注入：buildMessages 把上轮代码注入 LLM 提示
    assert.ok(pipeline.includes('【上一轮生成的代码】'),
        'LLM 提示应注入上一轮代码');
    assert.ok(pipeline.includes('请基于以上代码，按我的新需求修改或追加'),
        '应要求基于上轮代码修改/追加');
    // 数据继承：无新选区时从 _execSpec.data 补全上下文
    assert.ok(pipeline.includes('prevExec.spec.data'),
        '应继承上一轮数据快照');
    // 语言继承：上轮 lang 覆盖当前检测
    assert.ok(pipeline.includes("prevLang === 'vba' || prevLang === 'python'"),
        '应继承上一轮代码语言');
    // 放行：后续消息继承时不被模板拦截（上轮代码已通过验证，LLM 基于它修改）
    assert.ok(pipeline.includes("&& !prevCode)"),
        '后续消息继承时应放行 LLM 原始输出');
    console.log('  [ok] VBA 继承全链路契约完整');
});

test('Python 继承: python-pipeline 写入 _execSpec 且后续消息复用上轮分析结论', () => {
    const pipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    // 路由：orchestrator python 关键词表覆盖分析后续诉求（再分析/细分/对比等）
    // 【实现演进适配】关键词表已精化（去掉过宽的单字"再"，保留核心诉求词）
    assert.ok(orchestrator.includes("python: /继续|再次|再来|换成|改为|换个|维度|细分|对比|预测|趋势|异常/i"),
        'python 关键词表应覆盖分析后续诉求');
    // 存储：分析路径写入通用 _execSpec（pipeline='python'）
    assert.ok(pipeline.includes("pipeline: 'python'"),
        '分析路径应写入 _execSpec 并标注 pipeline=python');
    assert.ok(pipeline.includes('success: overallSuccess'),
        '_execSpec.result 应存分析结论');
    assert.ok(pipeline.includes('summary: combinedSummary || lastAnalysisSummary || \'\''),
        '_execSpec.result 应存自然语言总结');
    assert.ok(pipeline.includes('steps: steps'),
        '_execSpec.result 应存各步骤结论');
    // 继承：后续消息从 _execSpec 取回上轮分析结论
    assert.ok(pipeline.includes('getPythonFollowUp(turn)'),
        '应实现 python 后续消息读取');
    assert.ok(pipeline.includes("prev.spec.pipeline !== 'python'"),
        '应校验上一轮管道为 python');
    // 注入：buildSelectorMessages / buildCustomMessages 注入上轮分析结论
    assert.ok(pipeline.includes('【上一轮分析结论】'),
        'LLM 提示应注入上轮分析结论');
    assert.ok(pipeline.includes('请基于以上分析结果与任务规划，按我的新需求继续分析'),
        '应要求基于上轮结论继续分析');
    // 数据继承：无新选区时继承上一轮数据快照
    assert.ok(pipeline.includes('if (!ctx2.selection && pd.selection)'),
        '应继承上一轮数据快照');
    console.log('  [ok] Python 继承全链路契约完整');
});

test('Python 单步分析: combinedSummary 声明在 if 块外（修复 ReferenceError）', () => {
    const pipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    // 回归：单步分析（allResults.length<=1）时跳过综合总结块，但下方 setExecSpec 仍引用
    // combinedSummary。若声明在 if 块内（let 块级作用域），块外引用抛
    // ReferenceError: combinedSummary is not defined（截图：描述统计单步分析报错）。
    const declIdx = pipeline.indexOf('let combinedSummary = \'\';');
    const ifIdx = pipeline.indexOf('if (allResults.length > 1)');
    assert.ok(declIdx >= 0, 'combinedSummary 应有声明');
    assert.ok(ifIdx >= 0, '应存在多步综合总结块');
    assert.ok(declIdx < ifIdx,
        'combinedSummary 声明必须先于 if 块出现（函数级作用域，单步分析也可引用）');
    console.log('  [ok] Python 单步分析 combinedSummary 作用域契约完整');
});

test('生命周期: clearExecSpec 新任务清理防跨任务误继承', () => {
    const sessionStore = fs.readFileSync(SESSION_STORE_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    const codePipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const pythonPipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    // 存储层：clearExecSpec 支持清指定 turn 或最近一个有 _execSpec 的 turn
    assert.ok(sessionStore.includes('clearExecSpec(turnId)'),
        'session-store 应实现 clearExecSpec');
    assert.ok(sessionStore.includes('不传则清最近一个有 _execSpec 的 turn'),
        'clearExecSpec 应支持无参清最近上下文');
    // 管道层：新任务（非后续消息）时清空上一轮执行上下文
    assert.ok(codePipeline.includes('if (!turn.context?._execFollowUp && sessionStore.clearExecSpec)'),
        'code-pipeline 新任务应清空执行上下文');
    assert.ok(pythonPipeline.includes('if (!turn.context?._execFollowUp && sessionStore.clearExecSpec)'),
        'python-pipeline 新任务应清空执行上下文');
    // 路由层：orchestrator 显式新任务标记 + 清理
    assert.ok(orchestrator.includes('NEW_TASK_PATTERNS'),
        'orchestrator 应有显式新任务标记');
    assert.ok(orchestrator.includes('新任务|换个任务|换一个任务|另起一个|重新开始|从头开始|新做一个|新建一个|重新做一个|换一个话题'),
        '新任务标记应覆盖明确换任务措辞');
    assert.ok(orchestrator.includes('sessionStore.clearExecSpec()'),
        'orchestrator 命中新任务应清空执行上下文');
    console.log('  [ok] 生命周期契约完整');
});

test('Skill 继承: skill-pipeline 写入 _execSpec 且后续消息复用上轮结果', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    // 存储：Skill 路径写入通用 _execSpec（pipeline='skill'）
    assert.ok(pipeline.includes('writeSkillExecSpec(turn, sessionStore,'),
        'skill-pipeline 完成路径应写入 _execSpec');
    assert.ok(pipeline.includes("pipeline: 'skill'"),
        '_execSpec 应标注来源管道为 skill');
    assert.ok(pipeline.includes('skillName:'),
        '_execSpec.params 应记录技能名（供后续消息锁定 Skill）');
    assert.ok(pipeline.includes('result:'),
        '_execSpec.result 应存上轮结果');
    // 继承：后续消息从 _execSpec 取回上轮结果
    assert.ok(pipeline.includes('getSkillFollowUp(turn)'),
        'skill-pipeline 应实现 getSkillFollowUp');
    assert.ok(pipeline.includes('applySkillFollowUp(turn)'),
        'skill-pipeline 应实现 applySkillFollowUp');
    assert.ok(pipeline.includes('_originalUserMessage'),
        '后续消息应保留原始用户消息');
    // 生命周期：新任务清空上一轮上下文
    assert.ok(pipeline.includes('if (!turn.context?._execFollowUp && sessionStore.clearExecSpec)'),
        'skill-pipeline 新任务应清空执行上下文');
    // 路由：orchestrator 从 params.skillName 锁定上一轮 Skill
    assert.ok(orchestrator.includes('followUp.spec.params?.skillName'),
        'orchestrator 应从 _execSpec.params.skillName 取回技能名');
    assert.ok(orchestrator.includes("kind = 'skill'"),
        'orchestrator 后续消息应路由回 skill-pipeline');
    console.log('  [ok] Skill 继承全链路契约完整');
});

test('Skill 图表: 聚合与图表同源（快照选区），选区移开不产生矛盾', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // 聚合必须传快照选区 range（剥离 sheet 名），与 excel_create_chart 同一数据源
    assert.ok(pipeline.includes('range: aggRangeAddr || undefined'),
        'excel_aggregate 应传快照选区 range，避免按实时 Selection 读到空区域');
    assert.ok(pipeline.includes("String(aggRangeAddr).split('!').slice(1).join('!')"),
        'range 应剥离 sheet 名（后端用 srcWs.Range 定位）');
    // 聚合失败且为"选区无效"类错误时中止流程，避免"聚合失败但图表成功"矛盾
    assert.ok(pipeline.includes('聚合分析失败：${aggErrStr'),
        '选区无效类聚合失败应中止并提示重新选择');
    // 【常量收敛适配】选区无效类错误文本正则已收敛到 skill-constants.js
    const constants = fs.readFileSync(SKILL_CONSTANTS_FILE, 'utf8');
    assert.ok(constants.includes('至少需要表头|选区为空'),
        '常量模块应识别选区无效类错误文本');
    assert.ok(pipeline.includes('AGGREGATE_INVALID_RANGE_RE'),
        'executor 应引用常量模块的选区无效正则');
    console.log('  [ok] 聚合与图表同源契约完整（快照 range + 选区无效中止）');
});

test('Skill 图表: 显式快照选区到后端 RangeContext（不依赖实时选区）', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // 图表路径应显式调用 excel_read_range 读取快照地址，触发后端 RangeContext.Update
    assert.ok(pipeline.includes('快照选区到后端 RangeContext'),
        '应实现快照选区到 RangeContext 的步骤');
    assert.ok(pipeline.includes("toolRegistry.get('excel_read_range')"),
        '快照应通过 excel_read_range 工具执行');
    assert.ok(pipeline.includes("String(dataRange).match(/^(.*)!(.+)$/)"),
        '应拆分 dataRange 的 sheet 名与地址');
    assert.ok(pipeline.includes('snapArgs.range = snapMatch[2]'),
        '快照应传剥离 sheet 名后的地址');
    // 快照失败不阻断主流程（仅优化）
    assert.ok(pipeline.includes('快照选区到 RangeContext 失败（不阻断）'),
        '快照失败应仅告警不阻断');
    assert.ok(pipeline.includes("if (e?.message === 'user_aborted') throw e"),
        '用户中止应继续抛出，其余异常不阻断');
    console.log('  [ok] 显式快照选区到 RangeContext 契约完整');
});

test('数据兜底: 快照地址失效时自动回退 UsedRange（前后端一致）', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    const vb = fs.readFileSync(MCP_ADVANCED_FILE, 'utf8');
    // 前端：聚合结果解析抽成共用辅助函数（主路径 + 兜底重试）；现 async（内部可能补写图表数据源）
    assert.ok(/const applyAggregateResult = async \(res\) =>/.test(pipeline),
        '应抽公共聚合结果解析函数');
    assert.ok(pipeline.includes('applyAggregateResult(aggResult)'),
        '主路径应调用公共解析函数');
    // 前端：选区无效失败时无 range 兜底重试（让后端走 RangeContext → Selection → UsedRange）
    assert.ok(pipeline.includes('快照地址失效，尝试无 range 兜底重试（UsedRange）'),
        '选区无效时应先无 range 兜底重试');
    assert.ok(pipeline.includes('const retryArgs = {'),
        '应构造兜底重试参数');
    assert.ok(pipeline.includes('applyAggregateResult(retryResult)'),
        '兜底重试成功应复用公共解析函数');
    assert.ok(pipeline.includes('兜底重试）'),
        '兜底重试应标注（兜底重试）');
    // 后端：excel_aggregate 未传 range 时走 RangeContext 回退链（与 create_chart 一致）
    assert.ok(vb.includes('RangeContext.TryGet(srcWs.Name, ctxAddr)'),
        'ExecuteAggregate 未传 range 时应命中 RangeContext');
    // 【断言适配】行为契约：地址解析失败/行数不足两类分支都应 RangeContext.Clear() 并回退
    assert.ok(vb.includes('srcRange.Rows.Count < 2 Then') && vb.includes('RangeContext.Clear()'),
        'RangeContext 区域行数不足应视为失效并清除');
    assert.ok(/Catch[^]*?RangeContext\.Clear\(\)/.test(vb),
        'RangeContext 地址异常应清除并回退 Selection/UsedRange');
    console.log('  [ok] 数据兜底契约完整（前端无 range 重试 + 后端 RangeContext 回退链）');
});

test('聚合结果表跨轮复用: 上一轮聚合结果写入 _execSpec 并供下一轮直接引用', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // 图表路径：聚合成功后保存 aggregateSnapshot（地址+参数）
    assert.ok(pipeline.includes('let aggregateSnapshot = null'),
        '应声明聚合结果表快照变量');
    assert.ok(pipeline.includes('aggregateSnapshot = {'),
        '聚合成功后应保存聚合结果表快照');
    assert.ok(pipeline.includes('aggregate: aggregateSnapshot'),
        'yield result 应携带聚合结果表');
    // writeSkillExecSpec：保存聚合结果表到 _execSpec.data.aggregateResult
    assert.ok(pipeline.includes('aggregateResult: result?.aggregate || null'),
        'writeSkillExecSpec 应保存聚合结果表到 _execSpec');
    // applySkillFollowUp：继承上一轮聚合结果表到 ctx._aggregateResult
    // 【v20260924 适配】聚合结果迁入 data.ext.skill 命名空间，继承时优先读 ext.skill、兼容旧根字段
    assert.ok(pipeline.includes('extSkill.aggregateResult !== undefined ? extSkill.aggregateResult : pd.aggregateResult'),
        '聚合结果应从 ext.skill 读取并兼容旧根字段');
    assert.ok(pipeline.includes('ctx._aggregateResult = aggResult'),
        '后续消息应继承上一轮聚合结果表');
    assert.ok(pipeline.includes('复用上一轮聚合结果表 '),
        '继承时应记录复用日志');
    // 图表路径：优先复用上一轮聚合结果表，避免重新按原数据聚合
    assert.ok(pipeline.includes('复用上一轮聚合结果表（'),
        '检测到复用条件时应直接引用聚合结果表');
    assert.ok(pipeline.includes('const reuseAgg = prevAgg && prevAgg.dataRange'),
        '应判断复用条件（上一轮聚合结果表存在）');
    // 【常量收敛适配】复用判定/新维度判定正则已收敛到 skill-constants.js
    const constants = fs.readFileSync(SKILL_CONSTANTS_FILE, 'utf8');
    assert.ok(constants.includes('统计结果|上轮|刚才|这个|它|上面|之前|图表'),
        '常量模块复用条件应含引用上一轮的措辞');
    assert.ok(constants.includes('按|依照|根据)(?:部门|地区|产品|日期|人员|渠道|类型|月份|季度|年份'),
        '常量模块应识别明确新聚合维度措辞（不复用）');
    assert.ok(pipeline.includes('REUSE_AGGREGATE_RE') && pipeline.includes('AGGREGATE_REDIMENSION_RE'),
        'executor 应引用常量模块的复用/重聚合正则');
    console.log('  [ok] 聚合结果表跨轮复用契约完整');
});

test('多列分布图表: 数据分布每列生成一张图（三张表 → 三张图）', () => {
    const pipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    // buildTemplateChartCfg：data_distribution 遍历所有列生成配置数组
    assert.ok(pipeline.includes("if (templateId === 'data_distribution')"),
        '应处理 data_distribution 模板');
    assert.ok(pipeline.includes('const cfgs = []'),
        '应收集每列的图表配置');
    assert.ok(pipeline.includes('Object.keys(obj).forEach'),
        '应遍历所有列（而非只取首个列）');
    assert.ok(pipeline.includes('cfgs.push({'),
        '每列应生成一个图表配置');
    assert.ok(pipeline.includes('if (cfgs.length === 1) return cfgs[0]'),
        '单列应兼容旧返回类型（单配置）');
    assert.ok(pipeline.includes('if (cfgs.length > 1) return cfgs'),
        '多列应返回配置数组');
    // buildResultChartDiv：配置数组 → 每列渲染一个图表 div
    assert.ok(pipeline.includes('if (Array.isArray(chartCfg))'),
        'buildResultChartDiv 应处理配置数组');
    assert.ok(pipeline.includes('chartCfg.map(function (cfg) { return buildChartDiv(cfg); }).join'),
        '多列应渲染多个图表 div');
    console.log('  [ok] 多列分布图表契约完整');
});

test('PPT 跨管道继承分析结果: python 保存完整结构化结果并注入 PPT 规划', () => {
    const py = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    const code = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    // python-pipeline：_execSpec 保存完整分析结果（details 结构化数据）
    assert.ok(py.includes('details: allResults.map(function (r)'),
        'python 应保存每个分析步骤的完整结构化结果');
    assert.ok(py.includes('detail: String(detail).slice(0, 4000)'),
        '分析结果详情应限制大小');
    // code-pipeline：继承时把完整分析结果注入 PPT 规划 prompt
    assert.ok(code.includes('const prevDetails = prevRes.details'),
        'code-pipeline 应读取上一轮分析结果详情');
    assert.ok(code.includes('上一轮分析结果（请直接引用这些数据填充 KPI/图表/要点页'),
        '应指示 LLM 直接引用分析结果而非重新计算');
    assert.ok(code.includes('detailLines.slice(0, 6000)'),
        '注入的分析结果详情应限制大小');
    console.log('  [ok] PPT 跨管道继承分析结果契约完整');
});

test('Skill 定义一致性: 描述性统计报告声明 python 且模板为 Python 代码', () => {
    const skillMd = fs.readFileSync(path.join(ROOT, 'Skills', 'DescriptiveStatsReporter.md'), 'utf8');
    // 名字"描述性统计报告"是数据分析 → pipeline 应声明 python（而非 skill/VBA）
    assert.ok(skillMd.includes('pipeline: python'),
        '描述性统计报告应声明 pipeline: python（与名字一致）');
    assert.ok(!skillMd.includes('pipeline: skill'),
        '描述性统计报告不应再声明 pipeline: skill');
    // 模板应为 Python 代码（而非 VBA）
    assert.ok(skillMd.includes('```python'),
        'Template Code 应为 Python 代码块');
    assert.ok(!skillMd.includes('```vba'),
        'Template Code 不应再是 VBA 代码块');
    assert.ok(!skillMd.includes('vbaCode'),
        '输出格式不应再含 vbaCode 字段');
    assert.ok(skillMd.includes("import json, statistics"),
        'Python 模板应含描述统计实现');
    assert.ok(skillMd.includes("'stdev': round(variance**0.5, 4)"),
        'Python 模板应计算标准差');
    console.log('  [ok] Skill 定义一致性契约完整');
});

test('Skill 定义一致性: SmartDataAnalyzer 模板为 Python 且无 VBA', () => {
    const skillMd = fs.readFileSync(path.join(ROOT, 'Skills', 'SmartDataAnalyzer.md'), 'utf8');
    // pipeline 已声明 python（数据分析 → python 管道）
    assert.ok(skillMd.includes('pipeline: python'),
        'SmartDataAnalyzer 应声明 pipeline: python');
    assert.ok(skillMd.includes('processingMode: PythonCode'),
        'SmartDataAnalyzer 应声明 processingMode: PythonCode');
    // 模板应为 Python 代码（而非 VBA）
    assert.ok(skillMd.includes('```python'),
        'Template Code 应为 Python 代码块');
    assert.ok(!skillMd.includes('```vba'),
        'Template Code 不应再是 VBA 代码块');
    assert.ok(!skillMd.includes('Sub SmartDataAnalyzer'),
        '不应再含 VBA Sub 过程');
    assert.ok(skillMd.includes("import json, math"),
        'Python 模板应含异常检测实现');
    assert.ok(skillMd.includes("Z-Score") || skillMd.includes('z = (v - mean) / std'),
        'Python 模板应含 Z-Score 异常检测');
    assert.ok(!skillMd.includes('VBA 要执行的操作'),
        '字段说明不应再引用 VBA');
    console.log('  [ok] SmartDataAnalyzer 定义一致性契约完整');
});

test('Skill 模板一致性: 声明与实际执行模式不符的模板已清理', () => {
    // 1. SmartTableBeautifier：SkillSpec 模式，不应再有完整手写 VBA 宏
    const beautifier = fs.readFileSync(path.join(SKILLS_DIR, 'SmartTableBeautifier.md'), 'utf8');
    assert.ok(beautifier.includes('processingMode: SkillSpec'),
        'SmartTableBeautifier 应声明 SkillSpec 模式');
    assert.ok(!beautifier.includes('Sub SmartTableBeautify'),
        'SkillSpec 模式不应再含完整手写 VBA 宏');
    assert.ok(beautifier.includes('Spec 引擎'),
        '应标注实际由 Spec 引擎动态生成 VBA');
    assert.ok(beautifier.includes('不参与执行'),
        '应明确旧模板不参与执行');

    // 2. DataMerger/DataSplitter/TableJoiner：Hybrid 但实际走 MCP 工具，不应再含 MsgBox 空壳
    const mergeFamily = ['DataMerger', 'DataSplitter', 'TableJoiner'];
    mergeFamily.forEach(name => {
        const md = fs.readFileSync(path.join(SKILLS_DIR, name + '.md'), 'utf8');
        assert.ok(md.includes('不再执行任何 VBA 模板代码'),
            name + ' 应声明不执行 VBA 模板');
        assert.ok(md.includes('此模板仅作参考'),
            name + ' 应标注模板仅作参考');
        assert.ok(!/Sub \w+Operation\(\)[\s\S]*?MsgBox/.test(md),
            name + ' 不应再含 MsgBox 空壳代码');
    });
    console.log('  [ok] Skill 模板一致性契约完整（Beautifier/合并家族）');
});

test('Skill 管理 UI: 反馈按钮(有用/无用)已移除', () => {
    const html = fs.readFileSync(INDEX_HTML_FILE, 'utf8');
    const mgmt = fs.readFileSync(SKILL_MGMT_JS_FILE, 'utf8');
    // 1. HTML 不再含 有用/无用 checkbox
    assert.ok(!html.includes('skill-ui-show-like'),
        'HTML 不应再含"有用"checkbox');
    assert.ok(!html.includes('skill-ui-show-dislike'),
        'HTML 不应再含"无用"checkbox');
    assert.ok(!html.includes('>👍 有用</span>') && !html.includes('>👎 无用</span>'),
        'HTML 不应再渲染"有用/无用"按钮');
    // 2. JS 默认 UIActions 不应再含 ShowLike/ShowDislike 死字段
    assert.ok(!mgmt.includes('ShowLike') && !mgmt.includes('ShowDislike'),
        '默认 UIActions 不应再含 ShowLike/ShowDislike');
    // 3. 保留的按钮仍齐全
    assert.ok(html.includes('skill-ui-show-copy') && html.includes('skill-ui-show-save-md') && html.includes('skill-ui-show-regenerate'),
        '复制/保存MD/重新生成按钮应保留');
    console.log('  [ok] 反馈按钮移除契约完整');
});

test('Skill 管理 UI: 代码模板标签通用化（不再写死 VBA）', () => {
    const html = fs.readFileSync(INDEX_HTML_FILE, 'utf8');
    const mgmt = fs.readFileSync(SKILL_MGMT_JS_FILE, 'utf8');
    // 标签应为通用"代码模板"（模板可能是 VBA 或 Python）
    // 注：标签上允许挂 data-i18n 等属性（中英文化），故用正则而非整串相等
    assert.ok(/<label[^>]*>代码模板<\/label>/.test(html),
        '模板标签应为通用"代码模板"');
    assert.ok(!html.includes('VBA 代码模板'),
        '不应再写死"VBA 代码模板"标签');
    // id 一致性：index.html 与 app-skill-mgmt.js 引用同一个 id
    assert.ok(html.includes('id="skill-code-group"'),
        'index.html 应使用通用 id skill-code-group');
    assert.ok(mgmt.includes("getElementById('skill-code-group')"),
        'app-skill-mgmt.js 应引用 skill-code-group');
    assert.ok(!mgmt.includes('skill-vba-code-group'),
        'app-skill-mgmt.js 不应再引用 skill-vba-code-group');
    // placeholder 通用化（支持 VBA / Python）
    assert.ok(html.includes('在这里编写代码'),
        'placeholder 应为通用代码提示');
    console.log('  [ok] Skill 管理 UI 标签契约完整');
});

test('描述性统计表: AI 自动选方向 + 用户可转置切换（行↔列）', () => {
    // 表格渲染纯函数已拆至 renderers/json.js，转置交互仍在 chat-renderer.js
    const jsonRenderer = fs.readFileSync(JSON_RENDERER_FILE, 'utf8');
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    // 1. renderDescriptiveStatsTable 支持 orientation 参数
    assert.ok(jsonRenderer.includes('renderDescriptiveStatsTable(data, orientation)'),
        'renderDescriptiveStatsTable 应支持 orientation 参数');
    // 2. AI 智能选默认方向：列数少(≤3)→统计量为列，列数多→统计量为行
    assert.ok(jsonRenderer.includes("cols.length <= 3 ? 'col' : 'row'"),
        'AI 应按列数智能选择默认方向（≤3列→宽表，>3列→长表）');
    // 3. 表格带 data-stats-orient 属性（转置时用）
    assert.ok(jsonRenderer.includes('data-stats-orient'),
        '表格应带 data-stats-orient 属性标识方向');
    // 4. 支持统计量为列（横向/宽表）渲染
    assert.ok(jsonRenderer.includes("if (orient === 'col')"),
        '应支持统计量为列（宽表）渲染模式');
    // 5. 用户可手动切换（转置按钮）：交互在 chat-renderer
    assert.ok(renderer.includes('toggleStatsTableOrientation'),
        '应提供 toggleStatsTableOrientation 转置函数');
    assert.ok(renderer.includes('⇄ 转置'),
        '描述性统计卡片应有"转置"按钮');
    assert.ok(jsonRenderer.includes('py-stats-table'),
        '应通过 py-stats-table class 识别描述性统计表');
    // 6. 转置后 copy/insert 仍用最新 table
    assert.ok(renderer.includes('container.querySelector(\'table\') || tableRef'),
        '复制/插入按钮应使用转置后的最新 table');
    console.log('  [ok] 描述统计表格方向契约完整（AI 选 + 用户切）');
});

test('Doc 继承: doc-pipeline 写入 _execSpec 且后续消息复用上轮文档内容', () => {
    const pipeline = fs.readFileSync(DOC_PIPELINE_FILE, 'utf8');
    // 存储：文档路径写入通用 _execSpec（pipeline='doc'）
    assert.ok(pipeline.includes('writeDocExecSpec(turn, sessionStore, content)'),
        'doc-pipeline 完成路径应写入 _execSpec');
    assert.ok(pipeline.includes("pipeline: 'doc'"),
        '_execSpec 应标注来源管道为 doc');
    assert.ok(pipeline.includes('result: { content: content || \'\' }'),
        '_execSpec.result 应存上轮文档内容');
    // 【v20260924 适配】data 统一由 execSpecFactory.buildCommonData 构建，
    // doc 专有 attachments 挂到 data.ext.doc 命名空间（不污染公共 data 字段）
    assert.ok(pipeline.includes('fact.buildCommonData(turn, { doc: { attachments: atts } })'),
        '_execSpec.data 应存附件快照（ext.doc.attachments）');
    // 继承：后续消息从 _execSpec.result.content 取回上轮文档内容
    assert.ok(pipeline.includes('getDocFollowUp(turn)'),
        'doc-pipeline 应实现 getDocFollowUp');
    assert.ok(pipeline.includes('applyDocFollowUp(turn)'),
        'doc-pipeline 应实现 applyDocFollowUp');
    assert.ok(pipeline.includes('prevSpec.result?.content'),
        '后续消息应复用上轮文档内容');
    assert.ok(pipeline.includes('_originalUserMessage'),
        '后续消息应保留原始用户消息');
    // 生命周期：新任务清空上一轮上下文
    assert.ok(pipeline.includes('if (!turn.context?._execFollowUp && sessionStore.clearExecSpec)'),
        'doc-pipeline 新任务应清空执行上下文');
    console.log('  [ok] Doc 继承全链路契约完整');
});

test('聚合快照作用域: aggregateSnapshot 声明在函数顶层（美化/清洗路径不报 ReferenceError）', () => {
    const pipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // 1. 声明必须在 isChart 分支之前（函数顶层），而不是图表路径的 if 块内
    const declIdx = pipeline.indexOf('let aggregateSnapshot = null');
    const chartBranchIdx = pipeline.indexOf('if (isChart) {', pipeline.indexOf('function* executeSkillSpecMode'));
    assert.ok(declIdx !== -1, '应存在 aggregateSnapshot 声明');
    assert.ok(declIdx !== -1 && chartBranchIdx !== -1 && declIdx < chartBranchIdx,
        'aggregateSnapshot 声明应在 isChart 分支之前（函数顶层）');
    // 2. 图表路径内不应再有重复的 let 声明（应只有赋值）
    const chartBlockStart = pipeline.indexOf('if (typeof intentDetector.detectAggregateIntent');
    const chartBlockEnd = pipeline.indexOf('aggregateSnapshot = {', chartBlockStart);
    const between = pipeline.slice(chartBlockStart, chartBlockEnd);
    assert.ok(!/\blet aggregateSnapshot\b/.test(between),
        '图表路径块内不应重复声明 aggregateSnapshot（避免遮蔽顶层声明）');
    // 3. result 应引用 aggregateSnapshot
    assert.ok(pipeline.includes('aggregate: aggregateSnapshot'),
        'result 应携带 aggregate 快照');
    console.log('  [ok] 聚合快照作用域契约完整（顶层声明，各路径可访问）');
});

test('Text 继承: text-pipeline 写入 _execSpec 且后续消息复用上轮图表结果', () => {
    const pipeline = fs.readFileSync(TEXT_PIPELINE_FILE, 'utf8');
    // 存储：图表路径写入通用 _execSpec（pipeline='skill'，供 FOLLOWUP_PATTERNS.skill 匹配）
    assert.ok(pipeline.includes('writeChartExecSpec(turn, sessionStore, result)'),
        'text-pipeline 图表路径应写入 _execSpec');
    assert.ok(pipeline.includes("pipeline: 'skill'"),
        '_execSpec 应标注 pipeline=skill（对齐 FOLLOWUP_PATTERNS.skill）');
    assert.ok(pipeline.includes('mode: \'text\''),
        '_execSpec.params 应记录 mode=text');
    assert.ok(pipeline.includes('result: result || {}'),
        '_execSpec.result 应存上轮图表结果');
    // 继承：后续消息从 _execSpec 取回上轮图表结果
    assert.ok(pipeline.includes('getChartFollowUp(turn)'),
        'text-pipeline 应实现 getChartFollowUp');
    assert.ok(pipeline.includes('applyChartFollowUp(turn)'),
        'text-pipeline 应实现 applyChartFollowUp');
    assert.ok(pipeline.includes('_originalUserMessage'),
        '后续消息应保留原始用户消息');
    assert.ok(pipeline.includes('turn._originalUserMessage || turn.userMessage'),
        '内置任务判定应使用原始消息（避免注入上下文干扰图表词统计）');
    // 生命周期：新任务清空上一轮上下文
    // 【实现演进适配】判定抽为 isNewTask 变量（额外排除子管道轮，契约语义不变）
    assert.ok(pipeline.includes("const isNewTask = !turn.context?._execFollowUp")
        && pipeline.includes('if (isNewTask && sessionStore.clearExecSpec)'),
        'text-pipeline 新任务应清空执行上下文');
    console.log('  [ok] Text 继承全链路契约完整');
});

test('PPT 跨管道继承: 分析结果(skill/text/python) → PPT 规划不断层', () => {
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    // 路由：orchestrator 识别"将上面的分析转成PPT"类跨管道诉求
    assert.ok(orchestrator.includes('PPT_CROSS_PIPELINE_PATTERNS'),
        'orchestrator 应定义 PPT 跨管道继承模式');
    assert.ok(/上面的分析|刚才的报告|把.*分析.*转.*PPT/i.test('将上面的分析转成PPT'),
        'PPT_CROSS_PIPELINE_PATTERNS 应命中"将上面的分析转成PPT"');
    assert.ok(orchestrator.includes('PPT 跨管道继承'),
        'orchestrator 应实现 PPT 跨管道继承检测');
    assert.ok(orchestrator.includes('prevPipeline === \'skill\' || prevPipeline === \'text\' || prevPipeline === \'python\''),
        '跨管道继承应覆盖 skill/text/python 分析管道');
    // 继承：code-pipeline 把分析类管道视为 PPT 后续消息
    assert.ok(pipeline.includes('prevExec.spec?.pipeline === \'skill\'') &&
        pipeline.includes('prevExec.spec?.pipeline === \'python\''),
        'isPptFollowUp 应接受 skill/text/python 分析管道');
    // 数据归一化：skill/text/python 的 {selection, referenceRegions} → directSpec
    assert.ok(pipeline.includes('function buildDirectSpecFromExecData'),
        '应实现 buildDirectSpecFromExecData 归一化数据快照');
    // 结论注入：把上一轮分析结论附加到 LLM 规划提示
    assert.ok(pipeline.includes('pptAnalysisContext'),
        '应提取上一轮分析结论');
    assert.ok(pipeline.includes('function buildPptPlanningMessage'),
        '应实现 buildPptPlanningMessage 注入分析结论');
    assert.ok(pipeline.includes('buildPptPlanningMessage(turn.userMessage, pptAnalysisContext, prevExec?.spec'),
        'LLM 规划调用应注入分析结论（含上一轮分析上下文）');
    // 报告标题：取上一轮技能名（如"描述性统计报告"）
    assert.ok(pipeline.includes('analysisName && directSpec && !directSpec.reportTitle'),
        '报告标题应取上一轮分析技能名');
    console.log('  [ok] PPT 跨管道继承全链路契约完整');
});

test('图表跨管道继承: 分析结果(skill/text/python) → 图表向导不断层', () => {
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    const skillPipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    const pythonPipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');

    // 路由：orchestrator 识别"生成图表"类跨管道诉求（无新选区时继承上一轮数据）
    assert.ok(orchestrator.includes('图表跨管道继承'),
        'orchestrator 应实现图表跨管道继承检测');
    assert.ok(orchestrator.includes("context._matchedSkill = chartSkill"),
        '图表跨管道继承应锁定 SmartChartWizard');
    assert.ok(orchestrator.includes("prevPipeline === 'skill' || prevPipeline === 'text' || prevPipeline === 'python'"),
        '图表跨管道继承应覆盖 skill/text/python 分析管道');
    assert.ok(orchestrator.includes('!ctxHasData'),
        '图表跨管道继承应在无新选区时触发');

    // 继承：skill-pipeline 接受 skill/text/python 上一轮（跨管道）
    assert.ok(skillPipeline.includes("p !== 'skill' && p !== 'text' && p !== 'python'"),
        'getSkillFollowUp 应接受 skill/text/python 跨管道继承');
    assert.ok(skillPipeline.includes('【上一轮分析结论】'),
        'buildSkillFollowUpPrompt 应注入上一轮分析结论(summary)');

    // 数据快照：python/skill 管道存储完整选区（含 rows/sheetName），供下游直接使用
    // 【v20260924 适配】data 收敛到 execSpecFactory.buildCommonData 单一数据源
    // （消除 5 个 pipeline 各自组装导致的字段漂移），rows/sheetName/headers 由工厂统一产出。
    const sessionStore = fs.readFileSync(SESSION_STORE_FILE, 'utf8');
    assert.ok(pythonPipeline.includes('fact.buildCommonData(turn)'),
        'python-pipeline 应通过工厂存储选区数据快照');
    assert.ok(skillPipeline.includes('fact.buildCommonData(turn'),
        'skill-pipeline 应通过工厂存储选区数据快照');
    assert.ok(sessionStore.includes('rows: rows'),
        '选区数据快照应含 rows');
    assert.ok(sessionStore.includes('sheetName: ctxSel.sheetName || ctxSel.sheet || ctxSel.worksheet'),
        '选区数据快照应含 sheetName');
    console.log('  [ok] 图表跨管道继承全链路契约完整');
});

test('PPT 反死模板: 结构由需求驱动而非"一列一章"套路', () => {
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    // 分析继承提示词：按分析类型给出结构建议
    assert.ok(pipeline.includes('【分析类型与推荐叙事结构】'),
        '分析继承提示词应给出按分析类型的结构建议');
    assert.ok(pipeline.includes('不要为每个数据列单独建"章节+表格"页'),
        '仍应禁止"一列一章"死模板');
    assert.ok(pipeline.includes('KPI：均值、中位数、标准差'),
        '应根据分析类型给出推荐结构（描述统计含 KPI+表格+结论）');
    // 需求驱动提示词：明确禁止"一列一章"死模板 + 按需求选择侧重点
    assert.ok(pipeline.includes('禁止为每个数据列单独建"章节+表格"页'),
        '通用规划提示词应禁止"一列一章"死模板');
    assert.ok(pipeline.includes('页面结构必须服务于用户的具体需求，不同需求用不同结构'),
        '提示词应要求结构随用户需求变化，而非固定套路');
    assert.ok(pipeline.includes('根据用户需求选择侧重点'),
        '提示词应按需求类型给出结构侧重点（趋势/对比/占比/汇报/明细）');
    assert.ok(pipeline.includes('用户要"趋势/走势/变化/增长"'),
        '趋势需求应导向折线图页');
    assert.ok(pipeline.includes('用户要"对比/比较"'),
        '对比需求应导向对比页');
    assert.ok(pipeline.includes('用户要"占比/构成/份额"'),
        '占比需求应导向饼图页');
    // 需求驱动基线：规则引擎推荐作为参考起点（可调整，非照搬）
    assert.ok(pipeline.includes('baselineSlides'),
        'generateSlideSpecViaLlm 应支持规则引擎基线参考');
    assert.ok(pipeline.includes('仅作起点参考，必须根据用户需求调整，不要照搬'),
        '基线应声明为可调整参考而非固定结构');
    // LLM 首选规划：普通请求也走 LLM（不再只对复杂需求），规则引擎兜底
    assert.ok(pipeline.includes('需求驱动: LLM 按用户需求规划页面（规则推荐为基线）'),
        '普通请求应首选 LLM 按需求规划');
    assert.ok(pipeline.includes('LLM 失败回退：规则引擎 + 需求应用'),
        'LLM 失败应回退规则引擎+需求应用');
    // 反死模板安全网：检测并折叠"一列一章+表格"套路
    assert.ok(pipeline.includes('function collapseRigidSections'),
        '应实现反死模板折叠函数');
    assert.ok(pipeline.includes('collapseRigidSections(valid)'),
        'LLM 输出应经过反死模板折叠');
    assert.ok(pipeline.includes('sectionTablePairs < 3'),
        '折叠触发条件应为≥3个章节+表格对');
    console.log('  [ok] PPT 需求驱动反死模板契约完整');
});

test('PPT 需求偏置: 用户需求透传到后端规则引擎，兜底结构也随需求变化', () => {
    const engine = fs.readFileSync(DESIGN_ENGINE_FILE, 'utf8');
    const bridge = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    const vb = fs.readFileSync(VB_FILE, 'utf8');
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');

    // 后端设计引擎：需求偏置模型 + 解析 + 应用
    assert.ok(engine.includes('Private Class RequirementBias'),
        '设计引擎应实现需求偏置模型');
    assert.ok(engine.includes('Private Shared Function ParseRequirement'),
        '设计引擎应实现需求解析（趋势/对比/构成/明细/图表类型）');
    assert.ok(engine.includes('Optional requirement As String = "") As JArray'),
        'BuildDesign 应接收用户需求参数');
    assert.ok(engine.includes('Dim bias As RequirementBias = ParseRequirement(requirement)'),
        'BuildDesign 应解析并应用需求偏置');
    assert.ok(engine.includes('Not bias.SkipTrend') && engine.includes('Not bias.SkipComposition'),
        '需求偏置应能裁剪趋势/构成章节');
    assert.ok(engine.includes('bias.WantComparison'),
        '需求偏置应能新增对比页');
    assert.ok(engine.includes('bias.WantDetail'),
        '需求偏置应能裁剪数据明细');
    assert.ok(engine.includes('If Not String.IsNullOrWhiteSpace(bias.ChartType) Then chartType = bias.ChartType'),
        '需求偏置应能覆盖图表类型');
    assert.ok(engine.includes('Optional requirement As String = "") As JArray'),
        'DesignFromData 应透传需求参数');

    // ExcelBridge：读取 requirement 并传给 BuildDesign
    assert.ok(bridge.includes('args.Value(Of String)("requirement")'),
        'RecommendPptDesign 应读取用户需求');
    assert.ok(bridge.includes('PptDesignEngine.BuildDesign(profile, reportTitle, src, requirement)'),
        'RecommendPptDesign 应把需求传给 BuildDesign');

    // PptGenerationService：autoDesign 兜底路径透传需求
    // 【实参演进】数据来源固定为 Excel 选区，原 src 变量已内联为字面量 "Excel 选区"；
    // requirement 仍透传 —— 这才是本用例真正要守护的契约。
    assert.ok(vb.includes('PptDesignEngine.DesignFromData(headers, rows, reportTitle, "Excel 选区", requirement)'),
        'Generate autoDesign 兜底应透传需求');

    // 前端：requestPptDesign 透传用户需求
    assert.ok(pipeline.includes('requirement: requirement || \'\''),
        '前端 requestPptDesign 应透传用户需求');
    assert.ok(pipeline.includes('requestPptDesign(directSpec, columnHints, turn.userMessage)'),
        '前端应把用户消息作为需求传给设计推荐');
    console.log('  [ok] PPT 需求偏置前后端契约完整');
});

test('主题导入: 前端模板引擎支持动态加载与解析', () => {
    const engine = fs.readFileSync(TEMPLATE_ENGINE_FILE, 'utf8');
    assert.ok(engine.includes('async function loadPptTemplates'),
        '模板引擎应实现 loadPptTemplates');
    assert.ok(engine.includes('function resolveTheme'),
        '模板引擎应实现 resolveTheme');
    assert.ok(engine.includes('COLOR_THEMES_BACKUP'),
        '应有内置兜底主题');
    assert.ok(engine.includes('loadPptTemplates,') && engine.includes('resolveTheme,'),
        '应导出 loadPptTemplates/resolveTheme');
    console.log('  [ok] 前端模板引擎动态加载/解析/兜底齐全');
});

test('主题导入: 原生文件浏览（BrowsePptFile）契约', () => {
    const vb = loadPptVbAll();
    const be = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');

    // 后端：PptGenerationService 实现 + ExcelBridge 暴露
    assert.ok(vb.includes('Public Shared Function BrowsePptFile'),
        'PptGenerationService 应实现 BrowsePptFile');
    assert.ok(vb.includes('System.Windows.Forms.OpenFileDialog'),
        '应使用原生 OpenFileDialog');
    assert.ok(vb.includes('"PowerPoint 演示文稿 (*.pptx)|*.pptx|'),
        '过滤器应限定 .pptx');
    assert.ok(vb.includes('Key .filePath = dlg.FileName'),
        '应返回所选文件路径');
    assert.ok(vb.includes('"用户取消选择"'),
        '用户取消应返回明确提示');
    assert.ok(be.includes('Public Function BrowsePptFile'),
        'ExcelBridge 应暴露 BrowsePptFile COM 接口');

    // 前端：优先原生浏览，Web 模式兜底手动输入
    assert.ok(renderer.includes('bridge.BrowsePptFile'),
        '前端应调用 BrowsePptFile');
    assert.ok(renderer.includes('browseRes.filePath'),
        '前端应消费 filePath');
    // 【兜底演进】原「Web 模式手动输入路径」已升级为本地 MCP 工具回落（ppt_browse_template）：
    // 同样是原生文件对话框，且不依赖新增 COM 桥方法，旧版宿主 DLL 下也能选文件。
    assert.ok(renderer.includes("callMcp('ppt_browse_template'"),
        '无原生对话框时应回落本地 MCP 选文件');
    assert.ok(renderer.includes('mcpBrowse.filePath'),
        'MCP 回落返回的路径应被消费');
    console.log('  [ok] BrowsePptFile 前后端契约一致（原生浏览 + Web 兜底）');
});

test('版式画像: 提取模型与函数齐全（标题/正文/表格/形状）', () => {
    const vb = loadPptVbAll();
    const models = [
        ['PptLayoutProfile', 'Private Class PptLayoutProfile'],
        ['PptRect', 'Private Class PptRect'],
        ['PptTableProfile', 'Private Class PptTableProfile'],
        ['PptShapeProfile', 'Private Class PptShapeProfile'],
        ['提取入口', 'Private Shared Function ExtractLayoutProfile'],
        ['布局查找', 'Private Shared Function FindFirstLayoutEntry'],
        ['EMU 解析', 'Private Shared Function ParseEmu'],
        ['摘要生成', 'Private Shared Function LayoutSummary'],
        ['列宽套用', 'Private Shared Sub ApplyTableColumnWidths']
    ];
    const missing = models.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `版式画像缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 版式画像模型/函数 ${models.length} 项齐全`);
});

test('版式画像: 提取逻辑覆盖占位符/表格/形状与 EMU 换算', () => {
    const vb = loadPptVbAll();
    const checks = [
        ['幻灯片尺寸', 'ppt/presentation.xml'],
        ['sldSz 尺寸', 'sldSz'],
        ['EMU→设计空间', 'scaleX'],
        ['标题占位符', 'phType = "title"'],
        ['正文占位符', 'phType = "body"'],
        ['slideLayout 提取', 'ppt/slideLayouts/slideLayout'],
        ['表格提取', 'graphicFrame'],
        ['列宽提取', 'gridCol'],
        ['表头填充色', 'HeaderFill'],
        ['形状提取', 'prstGeom'],
        ['形状填充色', 'solidFill'],
        ['跳过占位符形状', 'nvPr.Element(spns + "ph")'],
        ['跳过文本框形状', 'String.IsNullOrWhiteSpace(txBody.Value)'],
        ['至少一项才算成功', 'profile.Title Is Nothing AndAlso profile.Body Is Nothing']
    ];
    const missing = checks.filter(([, code]) => !vb.includes(code));
    assert.deepStrictEqual(missing.map(([n]) => n), [],
        `提取逻辑缺失: ${missing.map(([n]) => n).join(', ')}`);
    console.log(`  [ok] 提取逻辑 ${checks.length} 项齐全`);
});

test('版式画像: 导入流程保存并返回 layout，生成时套用', () => {
    const vb = loadPptVbAll();
    // 导入：提取 → 挂到主题 → 保存 → 响应返回
    // 【实现演进适配】提取调用增加 mediaDir 参数（版式内图片落盘）
    assert.ok(vb.includes('tj.Layout = ExtractLayoutProfile(pptxPath'),
        '导入应调用版式提取');
    assert.ok(vb.includes('Key .layout = LayoutSummary(tj.Layout)'),
        '导入响应应返回版式摘要');
    assert.ok(vb.includes('newTheme("layout") = JToken.FromObject(tj.Layout,'),
        'SaveCustomTheme 应序列化 layout（camelCase）');
    // 加载：BuildTheme 复制 LayoutProfile
    assert.ok(vb.includes('theme.LayoutProfile = tj.Layout'),
        'BuildTheme 应复制 LayoutProfile');
    // 生成：激活布局画像（applyLayout=false 时跳过）
    assert.ok(vb.includes('_activeLayout = If(applyLayout, theme.LayoutProfile, Nothing)'),
        'Generate 应激活布局画像');
    // 套用：标题/内容区/表格
    assert.ok(vb.includes('TitleLeft()') && vb.includes('TitleTop()'),
        '标题位置应套用画像');
    assert.ok(vb.includes('ContentTop()'),
        '内容区顶部应套用画像');
    assert.ok(vb.includes('SX(TableLeft())') && vb.includes('ApplyTableColumnWidths'),
        '表格位置与列宽应套用画像');
    // 套用：装饰形状（SetSlideBackground 中绘制，避免遮挡正文）
    assert.ok(vb.includes('_activeLayout.Shapes'),
        'SetSlideBackground 应套用装饰形状');
    assert.ok(vb.includes('Case "roundrect"') && vb.includes('Case "oval"'),
        '形状类型映射应覆盖圆角矩形与椭圆');
    assert.ok(vb.includes('SetSlideBackground(s, toRelease, theme.Bg)'),
        '各 builder 应把释放列表传入 SetSlideBackground');
    console.log('  [ok] 导入→保存→加载→套用链路完整');
});

test('版式画像: 前端导入成功展示版式摘要', () => {
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    assert.ok(renderer.includes('已提取版式'),
        '导入成功提示应展示版式摘要');
    assert.ok(renderer.includes('res.layout'),
        '应消费后端返回的 layout 摘要');
    console.log('  [ok] 前端版式摘要展示齐全');
});

test('版式画像: UI 开关控制套用与 LLM 规划感知', () => {
    const vb = fs.readFileSync(VB_FILE, 'utf8');
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const engine = fs.readFileSync(TEMPLATE_ENGINE_FILE, 'utf8');
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    // 后端：applyLayout=false 时跳过版式画像套用
    assert.ok(vb.includes('args.Value(Of Boolean)("applyLayout")'),
        '后端应读取 applyLayout 参数');
    assert.ok(vb.includes('_activeLayout = If(applyLayout, theme.LayoutProfile, Nothing)'),
        'applyLayout=false 时应跳过版式画像');
    // 前端引擎：暴露版式画像查询
    assert.ok(engine.includes('hasThemeLayout') && engine.includes('getThemeLayoutSummary'),
        '主题引擎应暴露版式画像查询函数');
    assert.ok(engine.includes('normalizeLayout(t && t.layout)'),
        '主题归一化应保留并归一化版式画像');
    // 配置卡：两个开关
    assert.ok(renderer.includes("mkApplyToggle('applyLayout'") && renderer.includes("mkApplyToggle('layoutAwarePlanning'"),
        '配置卡应提供套用/规划感知两个开关');
    // 管线：透传 applyLayout + LLM 版式感知重规划
    assert.ok(pipeline.includes('applyLayout,') && pipeline.includes('theme,'),
        'argsJson 应透传 applyLayout');
    assert.ok(pipeline.includes('getThemeLayoutSummary(theme)'),
        '管线应按所选主题读取版式摘要');
    assert.ok(pipeline.includes('generateSlideSpecViaLlm(buildPptPlanningMessage(turn.userMessage, pptAnalysisContext') &&
            pipeline.includes('directSpec, signal, layoutSummary'),
        '开启规划感知时应带版式上下文重新规划');
    console.log('  [ok] UI 开关→管线→后端 契约完整');
});

test('Zone 容量限制: 模型/推导/摘要/序列化/构建器强制全链路', () => {
    const vb = loadPptVbAll();
    const engine = fs.readFileSync(TEMPLATE_ENGINE_FILE, 'utf8');
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    // 后端模型：PptCapacityProfile + PptLayoutProfile.Capacity
    assert.ok(vb.includes('Private Class PptCapacityProfile'),
        '应定义容量画像类');
    assert.ok(vb.includes('Public Property MaxBullets As Integer'),
        '容量应含要点上限');
    assert.ok(vb.includes('Public Property MaxTableRows As Integer'),
        '容量应含表格行上限');
    assert.ok(vb.includes('Public Property MaxTableCols As Integer'),
        '容量应含表格列上限');
    assert.ok(vb.includes('Public Property MaxChartSeries As Integer'),
        '容量应含图表系列上限');
    assert.ok(vb.includes('Public Property MaxChartCategories As Integer'),
        '容量应含图表分类上限');
    assert.ok(vb.includes('Public Property Capacity As PptCapacityProfile'),
        '版式画像应挂载容量');
    // 推导：正文区高度→要点数，表格区高度→行数
    assert.ok(vb.includes('profile.Body.Height / 52'),
        '要点数应由正文区高度推导');
    assert.ok(vb.includes('profile.Table.Height / 14'),
        '表格行数应由表格区高度推导');
    assert.ok(vb.includes('profile.Table.Columns.Count'),
        '表格列数应由模板列数推导');
    // 解析兜底：CapacityOf 旧主题按区域尺寸推导
    assert.ok(vb.includes('Private Shared Function CapacityOf'),
        '应提供容量解析兜底函数');
    // 摘要：LayoutSummary 包含容量
    assert.ok(vb.includes('容量(要点≤'),
        '版式摘要应包含容量信息');
    // 序列化：camelCase（前端读取兼容）
    assert.ok(vb.includes('CamelCasePropertyNamesContractResolver'),
        'layout 应以 camelCase 序列化');
    // 构建器强制：表格/要点/图表
    assert.ok(vb.includes('cap.MaxTableCols') && vb.includes('cap.MaxTableRows'),
        '表格构建器应强制列/行容量');
    assert.ok(vb.includes('cap.MaxBullets'),
        '要点构建器应强制要点容量');
    assert.ok(vb.includes('cap.MaxChartSeries') && vb.includes('cap.MaxChartCategories'),
        '图表构建器应强制系列/分类容量');
    // 前端：normalizeLayout 兼容大小写 + 容量字段
    assert.ok(engine.includes('function normalizeLayout'),
        '前端应归一化版式画像');
    assert.ok(engine.includes("pick(l, ['title', 'Title'])"),
        '归一化应兼容 camelCase/PascalCase');
    assert.ok(engine.includes('capacity: capRaw'),
        '归一化应保留容量字段');
    assert.ok(engine.includes('容量上限(每页要点≤'),
        '版式摘要应包含容量上限');
    // LLM 规划：提示词含严格容量约束
    assert.ok(pipeline.includes('严格容量约束'),
        'LLM 提示词应含容量约束');
    assert.ok(pipeline.includes('bullets 条数 ≤ 容量上限'),
        '应约束要点条数');
    assert.ok(pipeline.includes('series 数组元素数 ≤ 容量上限'),
        '应约束图表系列数');
    console.log('  [ok] Zone 容量限制全链路契约完整');
});

test('生成历史日志: 每次生成成功/失败追加记录到固定 TXT', () => {
    const vb = loadPptVbAll();
    // 辅助方法存在
    assert.ok(vb.includes('Private Shared Sub AppendGenerationLog'),
        '应存在生成历史日志辅助方法');
    assert.ok(vb.includes('PPT生成历史.log'),
        '日志应写入固定文件 PPT生成历史.log');
    assert.ok(vb.includes('Environment.SpecialFolder.MyDocuments'),
        '日志应写入我的文档目录');
    assert.ok(vb.includes('File.AppendAllText'),
        '应追加写入而非覆盖');
    // 成功/失败路径都调用
    assert.ok(vb.includes("AppendGenerationLog(True, reportTitle, theme.Name"),
        '成功路径应记录标题/主题/页数/文件路径');
    assert.ok(vb.includes("AppendGenerationLog(False, reportTitle, \"\", 0, \"\", ex.Message)"),
        '失败路径应记录错误原因');
    console.log('  [ok] 生成历史日志契约完整');
});

test('PPT 直出路径: 状态消息以 TXT 代码块展示而非误标 JSON', () => {
    const pipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    // 直出路径在配置确认后把 _codeLang 标成 txt，使状态消息以「TXT」代码块展示
    const clearIdx = pipeline.indexOf("sessionStore.updateTurn(turn.id, { _codeLang: 'txt' });");
    assert.ok(clearIdx > 0, '直出路径应把 _codeLang 标成 txt');
    // 标记发生在 waitForPptConfig 之后（配置卡已渲染，不受影响）
    const waitIdx = pipeline.indexOf('await waitForPptConfig(turn.id, signal)');
    assert.ok(waitIdx > 0 && clearIdx > waitIdx,
        '_codeLang 标记应在配置确认之后');
    // 渲染层：LANG_DISPLAY 应把 txt 映射为 TXT 标签
    const renderer = fs.readFileSync(RENDERER_FILE, 'utf8');
    assert.ok(renderer.includes("txt: 'TXT'"),
        '渲染层应把 txt 映射为 TXT 标签');
    assert.ok(renderer.includes('if (turn._codeLang) {'),
        '渲染层应仅在 _codeLang 存在时按代码块渲染');
    console.log('  [ok] 状态消息以 TXT 代码块展示契约完整');
});

test('任务结构继承: 各管道持久化 taskStructure 供后续消息规划下一个任务', () => {
    const pythonPipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    const textPipeline = fs.readFileSync(TEXT_PIPELINE_FILE, 'utf8');
    const codePipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const skillPipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // python：分析计划（任务清单）持久化到 _execSpec.taskStructure
    assert.ok(pythonPipeline.includes('taskStructure: {'),
        'python-pipeline 应持久化 taskStructure');
    assert.ok(pythonPipeline.includes("displayName: (s.tpl && s.tpl.name) || s.templateName || ''"),
        'python taskStructure 应含分析步骤 displayName');
    // text：图表任务清单持久化到 _execSpec.taskStructure
    assert.ok(textPipeline.includes('taskStructure: (result && result.taskStructure) || null,'),
        'text-pipeline 应持久化 taskStructure');
    // 【归一化收敛】图表任务清单的归一化已从 text-pipeline 内联映射收敛到 intent-analyzer
    // 的通用 normalizer（兼容 string/object 两种任务形态：chartLabel→name、chartEnum→type）。
    const intentAnalyzer = fs.readFileSync(INTENT_ANALYZER_FILE, 'utf8');
    assert.ok(intentAnalyzer.includes("name: t.displayName || t.chartLabel || t.title || t.templateId || t.name || '任务'"),
        'taskStructure 归一化应取图表任务名（chartLabel）');
    assert.ok(intentAnalyzer.includes("type: t.type || t.chartEnum || t.templateId || ''"),
        'taskStructure 归一化应取图表类型（chartEnum）');
    // code：PPT 页面规划持久化到 _execSpec.taskStructure
    assert.ok(codePipeline.includes('taskStructure: {'),
        'code-pipeline 应持久化 taskStructure');
    assert.ok(codePipeline.includes('slides: recommendedSlides || finalSlides || null'),
        'code taskStructure 应含 PPT 页面规划');
    // skill：方案规格/执行步骤持久化到 _execSpec.taskStructure
    assert.ok(skillPipeline.includes('taskStructure: {'),
        'skill-pipeline 应持久化 taskStructure');
    assert.ok(skillPipeline.includes('steps: Array.isArray(result?.steps) ? result.steps : null'),
        'skill taskStructure 应含执行步骤');
    console.log('  [ok] 各管道持久化 taskStructure 契约完整');
});

test('任务结构注入: 后续消息 prompt 注入【上一轮任务规划】而非只靠关键词', () => {
    const pythonPipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    const textPipeline = fs.readFileSync(TEXT_PIPELINE_FILE, 'utf8');
    const codePipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const skillPipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // python：buildFollowUpPrompt 注入任务规划
    assert.ok(pythonPipeline.includes('【上一轮任务规划】'),
        'python buildFollowUpPrompt 应注入上一轮任务规划');
    assert.ok(pythonPipeline.includes('ts.tasks.map'),
        'python 任务规划应来自 taskStructure.tasks');
    // text：applyChartFollowUp 注入任务规划
    assert.ok(textPipeline.includes('【上一轮任务规划】'),
        'text applyChartFollowUp 应注入上一轮任务规划');
    assert.ok(textPipeline.includes('planTasks'),
        'text 任务规划应兼容 ts.tasks/ts.steps');
    // code：PPT 跨管道注入分析任务规划
    assert.ok(codePipeline.includes('【上一轮分析任务规划（请按此结构组织 PPT 章节）】'),
        'code PPT 跨管道应注入分析任务规划');
    assert.ok(codePipeline.includes('prevTs.tasks.map'),
        'code PPT 任务规划应来自 taskStructure.tasks');
    // skill：buildSkillFollowUpPrompt 注入任务规划
    assert.ok(skillPipeline.includes('【上一轮任务规划】'),
        'skill buildSkillFollowUpPrompt 应注入上一轮任务规划');
    assert.ok(skillPipeline.includes('ts.steps'),
        'skill 任务规划应兼容 ts.steps/ts.tasks');
    console.log('  [ok] 后续消息注入任务规划契约完整');
});

test('意图分析器: intent-analyzer 服务存在且提供路由分析/规划 prompt 构造', () => {
    const analyzer = fs.readFileSync(INTENT_ANALYZER_FILE, 'utf8');
    // 服务注册
    assert.ok(analyzer.includes('App.ChatV2.intentAnalyzer ='),
        'intent-analyzer 应注册到 App.ChatV2.intentAnalyzer');
    // 路由层意图分析：结合用户输入 + 上一轮任务结构
    assert.ok(analyzer.includes('function analyzeFollowUpIntent'),
        '应提供 analyzeFollowUpIntent 路由层意图分析');
    assert.ok(analyzer.includes('isFollowUp'),
        '意图分析应返回 isFollowUp 判定');
    assert.ok(analyzer.includes('targetPipeline'),
        '意图分析应返回 targetPipeline 路由目标');
    // 引用检测：结合任务结构内容（实体词）而非纯关键词
    assert.ok(analyzer.includes('function referencesPreviousTask'),
        '应提供 referencesPreviousTask 引用检测');
    assert.ok(analyzer.includes('normalizeTaskStructure'),
        '应提供 normalizeTaskStructure 统一任务结构');
    // 规划层 prompt 构造：注入用户新需求 + 意图 + 规划指令
    assert.ok(analyzer.includes('function buildFollowUpPlanningPrompt'),
        '应提供 buildFollowUpPlanningPrompt 规划 prompt 构造');
    assert.ok(analyzer.includes('【用户新需求】'),
        '规划 prompt 应注入用户新需求');
    assert.ok(analyzer.includes('【意图分析】'),
        '规划 prompt 应注入意图分析');
    assert.ok(analyzer.includes('规划下一个任务'),
        '规划 prompt 应明确要求规划下一个任务');
    console.log('  [ok] 意图分析器契约完整');
});

test('意图分析路由: orchestrator 用意图分析器替换纯关键词 detectFollowUp', () => {
    const orchestrator = fs.readFileSync(ORCHESTRATOR_FILE, 'utf8');
    // 主路径：意图分析器 analyzeFollowUpIntent（结合用户输入 + 任务结构）
    assert.ok(orchestrator.includes('intentAnalyzer.analyzeFollowUpIntent'),
        'orchestrator 应调用意图分析器 analyzeFollowUpIntent');
    assert.ok(orchestrator.includes('_execFollowUp = { turnId: followUp.turnId, intent:'),
        'orchestrator 应把意图分析结果传给管道');
    // 兜底：旧关键词 detectFollowUp 保留（意图分析器不可用时降级）
    assert.ok(orchestrator.includes('function detectFollowUp'),
        'orchestrator 应保留 detectFollowUp 兜底');
    console.log('  [ok] orchestrator 意图分析路由契约完整');
});

test('意图规划注入: 四管道后续消息 prompt 注入用户新需求 + 意图分析 + 规划指令', () => {
    const pythonPipeline = fs.readFileSync(PYTHON_PIPELINE_FILE, 'utf8');
    const textPipeline = fs.readFileSync(TEXT_PIPELINE_FILE, 'utf8');
    const codePipeline = fs.readFileSync(CODE_PIPELINE_FILE, 'utf8');
    const skillPipeline = fs.readFileSync(SKILL_PIPELINE_FILE, 'utf8') + fs.readFileSync(SKILL_SPEC_EXEC_FILE, 'utf8');
    // 各管道后续消息 prompt 应调用意图分析器 buildFollowUpPlanningPrompt
    assert.ok(pythonPipeline.includes('buildFollowUpPlanningPrompt'),
        'python 后续消息应使用 buildFollowUpPlanningPrompt');
    assert.ok(textPipeline.includes('buildFollowUpPlanningPrompt'),
        'text 后续消息应使用 buildFollowUpPlanningPrompt');
    assert.ok(codePipeline.includes('buildFollowUpPlanningPrompt'),
        'code PPT 跨管道应使用 buildFollowUpPlanningPrompt');
    assert.ok(skillPipeline.includes('buildFollowUpPlanningPrompt'),
        'skill 后续消息应使用 buildFollowUpPlanningPrompt');
    // 意图分析器本身注入用户新需求 + 意图分析（规划输入完整）
    const analyzer = fs.readFileSync(INTENT_ANALYZER_FILE, 'utf8');
    assert.ok(analyzer.includes('【用户新需求】') && analyzer.includes('【意图分析】'),
        '规划 prompt 应同时注入用户新需求与意图分析');
    console.log('  [ok] 四管道规划注入契约完整');
});

test('实体识别: intent-analyzer 提取上一轮数据列并构造列名映射 prompt', () => {
    const analyzer = fs.readFileSync(INTENT_ANALYZER_FILE, 'utf8');
    // 列提取：从上一轮 spec.data 提取列名 + 数值/日期类型推断
    assert.ok(analyzer.includes('function extractColumns'),
        '应提供 extractColumns 列提取');
    assert.ok(analyzer.includes('data.selection.headers'),
        '列提取应读取选区表头');
    assert.ok(analyzer.includes('data.referenceRegions'),
        '列提取应读取引用区域表头');
    assert.ok(analyzer.includes('numeric') && analyzer.includes('date'),
        '列提取应做数值列/日期列类型推断');
    // 列名映射 prompt：用户术语 → 实际列名（融合 Julius 实体识别）
    assert.ok(analyzer.includes('function buildColumnMappingPrompt'),
        '应提供 buildColumnMappingPrompt 列名映射 prompt 构造');
    assert.ok(analyzer.includes('【上一轮数据列】'),
        '列名映射 prompt 应注入上一轮数据列');
    assert.ok(analyzer.includes('可用列：'),
        '列名映射 prompt 应列出可用列');
    assert.ok(analyzer.includes('数值列：'),
        '列名映射 prompt 应列出数值列');
    assert.ok(analyzer.includes('映射到实际列名'),
        '列名映射 prompt 应指示模型把用户术语映射到实际列名');
    // 规划 prompt 应注入列名映射段
    assert.ok(analyzer.includes('const colPrompt = buildColumnMappingPrompt(prevSpec)'),
        '规划 prompt 应调用列名映射构造');
    console.log('  [ok] 实体识别/列名映射契约完整');
});

test('上下文应用: 规划 prompt 处理代词与省略表达并基于上一轮结论推断指代', () => {
    const analyzer = fs.readFileSync(INTENT_ANALYZER_FILE, 'utf8');
    // 引用检测应覆盖代词/指示词（上面的/刚才的/它/这个/这些）
    // 【实现演进适配】词表已刻意收紧（去掉"这个/那个/它"等 Excel 高频误判词，见 intent-analyzer 注释），
    // 契约不变：仍需覆盖指向上文的代词/指示措辞。
    assert.ok(analyzer.includes('上面的|刚才的|上一步|之前的|前面|它们'),
        '引用检测应覆盖代词/指示词');
    // 规划 prompt 应包含上下文应用指令：推断代词/省略表达指代对象
    assert.ok(analyzer.includes('若用户消息含代词或省略表达'),
        '规划 prompt 应包含代词/省略表达处理指令');
    assert.ok(analyzer.includes('推断其指代对象'),
        '规划 prompt 应指示基于上一轮结论推断指代对象');
    // 规划 prompt 应注入上一轮分析结论与任务规划（上下文应用的基础）
    assert.ok(analyzer.includes('【上一轮分析结论】'),
        '规划 prompt 应注入上一轮分析结论');
    assert.ok(analyzer.includes('【上一轮任务规划】'),
        '规划 prompt 应注入上一轮任务规划');
    console.log('  [ok] 上下文应用（代词/省略表达）契约完整');
});

// ============================================================
// PPT 主题模板编辑器 — 实际功能对接契约
// ============================================================
const THEME_EDITOR_FILE = path.join(ROOT, 'UI', 'ppt-theme-editor.html');
const SIDE_TOOLBAR_FILE = path.join(ROOT, 'UI', 'js', 'side-toolbar.js');
const TASK_PANE_MGR_FILE = path.join(ROOT, 'TaskPaneManager.vb');
const THIS_ADDIN_PANELS_FILE = path.join(ROOT, 'ThisAddIn.Panels.vb');
const HOST_BRIDGE_FILE = path.join(ROOT, 'Hosts', 'Shared', 'HostBridgeBase.vb');

test('主题编辑器: 入口走 OpenPptThemeEditor（任务窗格注入桥），非裸 window.open', () => {
    const index = fs.readFileSync(INDEX_HTML_FILE, 'utf8');
    const toolbar = fs.readFileSync(SIDE_TOOLBAR_FILE, 'utf8');
    assert.ok(index.includes('openPptThemeEditor()'),
        'index.html 🎨 应调用 openPptThemeEditor()');
    assert.ok(!/onclick="window\.open\('ppt-theme-editor\.html'/.test(index),
        'index.html 不应再裸 window.open 主题编辑器（popup 无 hostObjects）');
    assert.ok(toolbar.includes('window.openPptThemeEditor'),
        'side-toolbar 应定义 openPptThemeEditor');
    assert.ok(toolbar.includes('OpenPptThemeEditor'),
        '应优先调桥 OpenPptThemeEditor');
    console.log('  [ok] 入口对接任务窗格 / OpenPptThemeEditor');
});

test('主题编辑器: ExcelBridge + TaskPane + ThisAddIn 面板链路齐全', () => {
    const be = fs.readFileSync(EXCEL_BRIDGE_FILE, 'utf8');
    const panes = fs.readFileSync(THIS_ADDIN_PANELS_FILE, 'utf8');
    const tpm = fs.readFileSync(TASK_PANE_MGR_FILE, 'utf8');
    assert.ok(be.includes('Public Sub OpenPptThemeEditor'),
        'ExcelBridge 应暴露 OpenPptThemeEditor');
    assert.ok(be.includes('ShowPptThemeEditorPanel'),
        'OpenPptThemeEditor 应委托 ThisAddIn.ShowPptThemeEditorPanel');
    assert.ok(panes.includes('ShowPptThemeEditorPanel'),
        'ThisAddIn 应有 ShowPptThemeEditorPanel');
    assert.ok(tpm.includes('CreatePptThemeEditorTaskPaneAsync'),
        'TaskPaneManager 应有 CreatePptThemeEditorTaskPaneAsync');
    assert.ok(tpm.includes('ppt-theme-editor.html'),
        '任务窗格应加载 ppt-theme-editor.html');
    assert.ok(tpm.includes('PPT_THEME_PANE_WIDTH'),
        '主题编辑器应使用更宽的任务窗格');
    console.log('  [ok] Excel 任务窗格链路齐全');
});

test('主题编辑器: 前端 callHost 支持 opener 桥 + 列表按 id 载入（无双 stringify）', () => {
    const html = fs.readFileSync(THEME_EDITOR_FILE, 'utf8');
    assert.ok(html.includes('function resolveBridge'),
        '应有 resolveBridge（本窗/opener/App）');
    assert.ok(html.includes('window.opener'),
        'popup 降级时应尝试 opener.excelBridge');
    assert.ok(html.includes('savedThemesCache'),
        '应缓存已保存主题列表');
    assert.ok(html.includes('loadThemeById'),
        '列表点击应走 loadThemeById');
    assert.ok(!html.includes('JSON.stringify(JSON.stringify(t))'),
        '禁止双 stringify 塞进 onclick（会把字符串当对象传给 loadThemeObj）');
    assert.ok(html.includes("typeof t === 'string'"),
        'loadThemeObj 应兼容 JSON 字符串入参');
    console.log('  [ok] 前端桥接与列表载入契约完整');
});

test('主题编辑器: PPT/Word HostBridgeBase 暴露主题 CRUD + OpenPptThemeEditor', () => {
    const hb = fs.readFileSync(HOST_BRIDGE_FILE, 'utf8');
    assert.ok(hb.includes('Function SaveCustomPptTheme'),
        'HostBridgeBase 应有 SaveCustomPptTheme');
    assert.ok(hb.includes('Function ListCustomPptThemes'),
        'HostBridgeBase 应有 ListCustomPptThemes');
    assert.ok(hb.includes('Function DeleteCustomPptTheme'),
        'HostBridgeBase 应有 DeleteCustomPptTheme');
    assert.ok(hb.includes('Sub OpenPptThemeEditor'),
        'HostBridgeBase 应有 OpenPptThemeEditor');
    assert.ok(hb.includes('custom_ppt_themes.json'),
        '宿主 CRUD 应写 custom_ppt_themes.json');
    console.log('  [ok] 宿主桥主题 CRUD 齐全');
});

test('主题解析: ResolveTheme 支持按显示名匹配（编辑器「用主题名生成」）', () => {
    const vb = loadPptVbAll();
    assert.ok(vb.includes('String.Equals(t.Name, themeId, StringComparison.OrdinalIgnoreCase)'),
        'ResolveTheme 应在 id 未命中时按 Name 匹配');
    console.log('  [ok] ResolveTheme 支持主题名');
});

// 测试结束后强制退出（与其它测试文件一致，防止定时器挂起进程）
if (require.main === module) { setTimeout(() => process.exit(0), 3000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
