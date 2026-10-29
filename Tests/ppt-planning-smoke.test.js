/**
 * ============================================================
 * Tests/ppt-planning-smoke.test.js — PPT 规划提示词（第1/2步）回归冒烟
 * ============================================================
 * 覆盖：
 *   1. 第1步取数规则：有分析上下文时注入"直接引用分析结果、不重算"强制规则
 *      无分析上下文时原样返回（普通生成零污染）
 *   2. 第2步故事线：按上一轮分析 skill displayName 选择对应叙事模板
 *   3. 集成：真实销售数据 _execSpec → 生成完整规划提示词（desc_stats + 取数规则 + 完整叙事）
 * 说明：buildPptPlanningMessage/STORYBOARDS 是 code-pipeline.js 闭包内部函数，
 *       读取磁盘源码按结构边界实时提取（与实现同步、防漂移），不改生产代码。
 * 运行：node Tests/ppt-planning-smoke.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');

const SRC = path.join(__dirname, '..', 'UI', 'js', 'chat-v2', 'pipelines', 'code-pipeline.js');
const code = fs.readFileSync(SRC, 'utf8');

function closingBraceAt(src, openIdx) {
  let depth = 0;
  for (let i = openIdx; i < src.length; i++) {
    const ch = src[i];
    if (ch === '{') depth++;
    else if (ch === '}') { depth--; if (depth === 0) return i + 1; }
  }
  return -1;
}

/** 从磁盘提取 { deriveAnalysisKind, deriveAnalysisKinds, buildPptPlanningMessage, detectDataFeatures, buildStructureConstraints } */
function loadPlanner() {
  const start = code.indexOf('const SUPPORTED_SLIDE_TYPES = new Set(');
  const sig = code.indexOf('function buildPptPlanningMessage', start);
  const openBrace = code.indexOf('{', sig);
  const end = closingBraceAt(code, openBrace);
  assert.ok(start >= 0 && sig >= 0 && end >= 0, '源码中无法定位版式白名单/故事线/规划函数（结构可能变更）');
  const extracted = code.slice(start, end) +
    '; return { deriveAnalysisKind, deriveAnalysisKinds, buildPptPlanningMessage, detectDataFeatures, buildStructureConstraints, composeStoryline, deriveAudience, deriveDimensionFeatures, buildSeedOutline, deriveDataMapKeys, validateSlideDataMapping, AUDIENCE_STORYLINES };';
  return new Function(extracted)();
}

const planner = loadPlanner();
const { buildPptPlanningMessage, deriveAnalysisKinds, detectDataFeatures, buildStructureConstraints, deriveAudience, deriveDimensionFeatures, buildSeedOutline, validateSlideDataMapping } = planner;

const ANALYSIS_CTX = [
  '【上一轮分析结果】',
  '{"销售额":{"count":12,"sum":1955,"mean":162.9,"median":165,"stdev":25.7,"min":120,"max":200,"trend":"up"}}'
].join('\n');

const TRIGGER = /\u3010上一轮分析/;
const FIRST_STEP_GUARDS = ['严禁用样本原始数据重新计算', '直接引用分析结果已给出'];

test('规划提示词: 无分析上下文时原样返回（普通生成不被污染）', () => {
  const msg = '把销售做成PPT';
  assert.strictEqual(buildPptPlanningMessage(msg, ''), msg);
  console.log('  [ok] 无分析上下文原样返回');
});

test('规划提示词: 第1步取数规则就位且随分析上下文触发', () => {
  for (const g of FIRST_STEP_GUARDS) {
    assert.ok(code.includes(g), `源码应含强制取数规则: ${g}`);
  }
  const prompt = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, null);
  assert.ok(TRIGGER.test(prompt), '有分析上下文时应命中【上一轮分析标签（触发第1步注入）');
  console.log(`  [ok] 第1步取数规则 ${FIRST_STEP_GUARDS.length}/${FIRST_STEP_GUARDS.length} 就位且触发`);
});

const CASES = [
  ['描述性统计报告', 'desc_stats', '均值、中位数、标准差'],
  ['数据差异对比分析器', 'diff', '差异记录数'],
  ['智能数据分析(AI+VBA)', 'anomaly_quality', '异常记录数'],
  ['智能数据清洗大师', 'anomaly_quality', '缺失值数'],
  ['Python数据分析（执行）', 'generic', '核心指标'],
  [null, 'generic', '核心指标'],
];

test('规划提示词: 第2步按分析类型选择故事线模板', () => {
  for (const [displayName, wantKind, marker] of CASES) {
    const spec = displayName === null
      ? null
      : { taskStructure: { tasks: [{ displayName }] }, params: {} };
    const prompt = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, spec);
    assert.ok(prompt.includes(wantKind === 'desc_stats'
      ? '描述性统计' : wantKind === 'diff' ? '数据差异/新旧对比分析'
      : wantKind === 'anomaly_quality' ? '异常/数据质量问题分析' : '通用数据分析'),
      `[${displayName}] 应输出 ${wantKind} 故事线`);
    assert.ok(prompt.includes(marker), `[${displayName}] 故事线应含内容标记: ${marker}`);
    assert.ok(TRIGGER.test(prompt), `[${displayName}] 应触发第1步取数规则`);
    console.log(`  [ok] ${String(displayName).padEnd(22)} → ${wantKind}`);
  }
});

test('规划提示词: 端到端销售数据整链一致', () => {
  const prevSpec = { taskStructure: { tasks: [{ displayName: '描述性统计报告' }] }, params: {} };
  const prompt = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, prevSpec);
  assert.ok(prompt.includes('描述性统计'), '应为描述性统计叙事');
  assert.ok(prompt.includes('指标 → 依据 → 结论'), '应强调完整叙事');
  assert.ok(prompt.includes('table') || prompt.includes('表格'), '叙事应含明细表格建议');
  assert.ok(/均值|中位数|标准差/.test(prompt), '叙事应引导填充统计指标');
  assert.ok(TRIGGER.test(prompt), '应触发第1步取数规则');
  console.log('  [ok] 端到端：desc_stats 叙事 + 取数规则 + 完整叙事链全部就位');
});

test('规划提示词: 复合任务 → 组合叙事（多类分析并存时覆盖全部）', () => {
  const prevSpec = { taskStructure: { tasks: [{ displayName: '描述性统计报告' }, { displayName: '数据差异对比分析器' }] }, params: {} };
  const kinds = deriveAnalysisKinds(prevSpec);
  assert.deepStrictEqual(kinds, ['desc_stats', 'diff'], '复合任务应有序去重返回两类分析');
  const prompt = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, prevSpec);
  assert.ok(prompt.includes('复合分析（描述性统计 + 数据差异/新旧对比分析）'), '复合任务应输出组合叙事综述');
  assert.ok(prompt.includes('均值、中位数、标准差') && prompt.includes('差异记录数'), '组合叙事应同时覆盖两类内容模块');
  console.log(`  [ok] 复合任务组合叙事: ${kinds.join(' + ')}`);
});

test('规划提示词: 数据特征动态注入（趋势/增幅/分组提示）', () => {
  const ctx = '【上一轮分析结果】整体呈上升趋势，较期初增长 66.7%，「华东」占比最高；销售额均值 162.9。';
  const f = detectDataFeatures(ctx);
  assert.strictEqual(f.trend, 'up', '应识别上升趋势');
  assert.strictEqual(f.growthPct, 66.7, '应解析增幅百分比');
  assert.ok(f.group, '应识别分组/占比迹象');
  const prompt = buildPptPlanningMessage('把这次分析做成PPT', ctx, null);
  assert.ok(prompt.includes('【本次数据特征提示】'), '应动态注入本次特征提示块');
  assert.ok(prompt.includes('上升趋势（增幅 66.7%）'), '特征值应进入叙事取值提示');
  console.log(`  [ok] 数据特征注入: trend=${f.trend} growth=${f.growthPct}% group=${f.group}`);
});

test('规划提示词: 结构约束块含 type 白名单与 KPI 单值约束', () => {
  const c = buildStructureConstraints();
  for (const t of ['title', 'agenda', 'kpi', 'chart', 'table', 'text', 'comparison', 'closing']) {
    assert.ok(c.includes(t), `结构约束应含合法 type: ${t}`);
  }
  assert.ok(/KPI 每项 value 只能是单个数值/.test(c), '应含 KPI 单值约束');
  assert.ok(/禁止凭空编造数值/.test(c) || /"—"/.test(c), '应含缺失用"—"' + '的约束');
  console.log('  [ok] 结构约束: 白名单 + KPI单值 + 缺失— 齐全');
});

test('规划提示词: 推荐结构基线并入（以其为骨架）且缺省不注入', () => {
  const prevSpec = { taskStructure: { tasks: [{ displayName: '描述性统计报告' }] }, params: {} };
  const withBaseline = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, prevSpec, [{ type: 'kpi', title: '关键指标', items: [] }]);
  assert.ok(withBaseline.includes('【推荐结构基线'), '传入 recommendedSlides 应注入基线块');
  assert.ok(withBaseline.includes('以其为骨架'), '应声明以基线为骨架');
  const without = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, prevSpec);
  assert.ok(!without.includes('【推荐结构基线'), '未传 recommendedSlides 应不注入基线块');
  console.log('  [ok] 推荐基线融合: 注入/缺省行为正确');
});

test('深化: 受众/场景分层叙事（汇报为结论先行，缺省中性）', () => {
  const prep = '把这次分析做成PPT';
  const rep = '给领导做个汇报PPT';
  const business = '这个要给业务运营看，怎么做改善';
  assert.strictEqual(deriveAudience(rep), 'leadership', '含"领导/汇报"应判为领导视角');
  assert.strictEqual(deriveAudience(business), 'business', '含"业务/运营/怎么"应判为业务视角');
  assert.strictEqual(deriveAudience(prep), null, '中性诉求不应误判受众');
  const prompt = buildPptPlanningMessage(rep, ANALYSIS_CTX, null);
  assert.ok(prompt.includes('【受众叙事视角】') && prompt.includes('汇报/领导视角'), '应注入受众叙事视角块');
  assert.ok(prompt.includes('结论先行'), '领导视角应强调结论先行（倒金字塔）');
  console.log('  [ok] 受众分层: leadership=' + (prompt.includes('汇报/领导视角') ? 'Y' : 'N'));
});

test('深化: 维度驱动补拼（时间→趋势线、分组+数值→构成、≥2数值→对比）', () => {
  const directSpec = {
    headers: ['日期', '地区', '销售额', '利润'],
    rows: [{ '日期': '2024-01', '地区': '华东', '销售额': 100, '利润': 20 }]
  };
  const dims = deriveDimensionFeatures(directSpec, null);
  assert.ok(dims.hasTime, '含"日期"列应识别时间维度');
  assert.ok(dims.hasGroup, '含"地区"列应识别分组维度');
  assert.ok(dims.numericCount >= 2, '含销售额/利润应识别 ≥2 数值列');
  const prompt = buildPptPlanningMessage('把这次分析做成PPT', ANALYSIS_CTX, null, null, { directSpec });
  assert.ok(prompt.includes('【维度驱动补拼】'), '应注入维度驱动补拼块');
  assert.ok(prompt.includes('时间维度'), '时间维度应进入叙事');
  assert.ok(prompt.includes('页级大纲种子'), '应注入结构化页级大纲种子');
  assert.ok(/type=chart.*趋势|type=comparison/.test(prompt), '大纲应含趋势/对比页种子');
  console.log(`  [ok] 维度驱动: time=${dims.hasTime} group=${dims.hasGroup} num=${dims.numericCount}`);
});

test('深化: 结构化页级大纲种子（含每页数据提示，供 LLM 填值）', () => {
  const pages = buildSeedOutline(['desc_stats', 'trend'], { hasTime: true, hasGroup: true, numericCount: 2 }, 'leadership');
  const typ = pages.map(p => p.type);
  assert.ok(typ.includes('kpi') && typ.includes('chart') && typ.includes('table') && typ.includes('text'), '大纲应含 kpi/chart/table/text');
  assert.ok(pages[0].title === '执行摘要', '领导视角大纲首屏应为执行摘要');
  const chart = pages.find(p => p.type === 'chart');
  assert.ok(chart && chart.chartType && typeof chart.hint === 'string' && chart.hint.length > 0, '图表种子应含 chartType 与数据提示');
  console.log(`  [ok] 页级大纲: ${pages.length} 页, 首屏=${pages[0].title}, 图表=${chart ? chart.chartType : '-'}`);
});

test('深化: 页内数据映射校验（KPI单值、占位回退"—"、字段对齐）', () => {
  const slides = [
    { type: 'kpi', title: '关键指标', items: [
      { label: '总销售额', value: '100、200、300' },
      { label: '缺失指标', value: 'XX' },
      { label: '正常指标', value: '162.9' }
    ]},
    { type: 'comparison', title: '对比', items: [
      { name: '销售额', value: 100 },
      { name: '假编造字段', value: 50 }
    ]}
  ];
  // 分析结果不含"假编造字段"，表头也不含 → 应回退为 —
  const r = validateSlideDataMapping(slides, '均值为162.9，占比30%', ['地区', '销售额']);
  assert.strictEqual(r.slides[0].items[0].value, '100', '多值应保留首值');
  assert.strictEqual(r.slides[0].items[1].value, '—', '占位 XX 应回退为 —');
  assert.strictEqual(r.slides[0].items[2].value, '162.9', '正常单值应原样保留');
  assert.strictEqual(r.slides[1].items[0].name, '销售额', '表头内字段应保留');
  assert.strictEqual(r.slides[1].items[1].name, '—', '编造字段应回退为 —');
  assert.ok(r.warnings.length >= 3, '应产生修正提示');
  console.log(`  [ok] 数据映射校验: warnings=${r.warnings.length}, KPI单值/占位回退/字段对齐均生效`);
});

test('深化: PPT 图表次坐标轴（secondaryAxis 字段约定就位且声明默认不启用）', () => {
  // 规划提示词字段约定应支持可选 secondaryAxis，且只在量纲差异大的组合图启用
  const guidance = /chart: \{"type":"chart"[^]*?secondaryAxis[\s\S]*?次轴/.exec(code);
  assert.ok(guidance, '源码应含 chart 页 secondaryAxis 字段说明');
  assert.ok(code.includes('secondaryAxis'), '规划提示词应允许 LLM 输出 secondaryAxis 字段');
  assert.ok(/仅当.*量纲差异大|\u53ea\u6709.*\u7ec4\u5408|\u4e0d\u8981\u7528/.test(code) || code.includes('同量纲或单系列'), '应声明同量纲/单系列勿用次轴以防误导');
  console.log('  [ok] 次坐标轴字段约定: secondaryAxis 就位且默认不启用');
});