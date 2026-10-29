/**
 * python-template-validator.test.js
 * 验证 Python 模板库扩展机制 + 参数校验层：
 *  1) linear_regression 模板已注册进模板库（扩展=新增模板）
 *  2) validateTemplateParams：schema 合规 + 列存在性/数值性校验
 *  3) buildColumnSuggestions：坏列给出近义候选 + 可用列清单（防"利润率"式误导）
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2 } = require('./helpers/harness');

const HEADERS = ['日期', '部门', '地区', '销售额', '成本', '利润'];
const NUMERIC = ['销售额', '成本', '利润'];

function tmp() {
    return loadChatV2().App.PythonTemplateValidator;
}

test('模板库：linear_regression 与 6 个内置模板均已注册', () => {
    const ids = tmp().getTemplateIds();
    for (const id of ['descriptive_stats', 'trend_analysis', 'correlation', 'anomaly_detection', 'data_distribution', 'group_aggregation', 'linear_regression']) {
        assert.ok(ids.includes(id), `模板应注册 "${id}"，实际: ${ids.join(', ')}`);
    }
});

test('校验合法参数：descriptive_stats 通过', () => {
    const r = tmp().validateTemplateParams('descriptive_stats', { columns: ['销售额', '成本'] }, HEADERS, NUMERIC);
    assert.strictEqual(r.ok, true);
    assert.strictEqual(r.invalidColumns.length, 0, '不应有无效列');
    assert.strictEqual(r.issues.length, 0);
});

test('校验缺必填参数：trend_analysis 缺失 value_column 判不通过', () => {
    const r = tmp().validateTemplateParams('trend_analysis', {}, HEADERS, NUMERIC);
    assert.strictEqual(r.ok, false);
    assert.ok(r.missingRequired.includes('value_column'), `应缺 value_column，实际: ${r.missingRequired.join(', ')}`);
});

test('坏列命中：利润率(不存在)被判无效列并给出近义候选 + 可用列清单', () => {
    const V = tmp();
    const r = V.validateTemplateParams('descriptive_stats', { columns: ['利润率'] }, HEADERS, NUMERIC);
    assert.strictEqual(r.ok, true, '列失效不应判必填缺失失败');
    assert.ok(r.invalidColumns.length > 0, '应检测到无效列');
    assert.ok(r.issues.some(i => i.includes('利润率')), 'issues 应提到利润率');
    assert.ok(r.columnSuggestions.length > 0, '应生成列建议');
    assert.ok(r.columnSuggestions.includes('利润'), '建议应含近义候选/可用列 利润');
});

test('坏列并列示：可用列清单完整列出真实列', () => {
    const sug = tmp().buildColumnSuggestions([{ name: '利润率', numeric: true }], HEADERS, NUMERIC);
    assert.ok(sug.includes('利润率'), '建议应点名坏列');
    assert.ok(sug.includes('销售额'), '建议应列出可用数值列 销售额');
    assert.ok(sug.includes('利润'), '建议应列出可用数值列 利润');
});

test('linear_regression：合法参数通过，坏自变量被判无效', () => {
    const V = tmp();
    assert.strictEqual(V.validateTemplateParams('linear_regression', { xColumn: '成本', yColumn: '利润' }, HEADERS, NUMERIC).ok, true);
    const bad = V.validateTemplateParams('linear_regression', { xColumn: '利润率', yColumn: '利润' }, HEADERS, NUMERIC);
    assert.ok(bad.invalidColumns.length > 0, '无效自变量应被检出');
});

test('linear_regression.generate：产出引用 input_data 且含两列名的可执行脚本', () => {
    const script = tmp().generate('linear_regression', { xColumn: '成本', yColumn: '利润', forecastPoints: 3 });
    assert.ok(script && typeof script === 'string', '应产出脚本字符串');
    assert.ok(script.includes('input_data'), '应通过系统注入的 input_data');
    assert.ok(script.includes('成本') && script.includes('利润'), '应引用两列名');
    assert.ok(script.includes('fps = '), '应注入预测点数');
    assert.ok(script.includes('r2'), '应计算 R²');
    assert.ok(script.includes('parse_dt'), '应含日期解析（含序列号）支持');
    assert.ok(script.includes('1899, 12, 30'), '应含 Excel 序列号基准日期');
    assert.ok(script.includes('to_num'), '应把自变量/因变量统一转数值');
});

test('未知模板：校验报告未知模板', () => {
    const r = tmp().validateTemplateParams('no_such_template', {}, HEADERS, NUMERIC);
    assert.strictEqual(r.ok, false);
    assert.ok(r.issues.some(i => i.includes('未知模板')));
    assert.strictEqual(tmp().generate('no_such_template', {}), null);
});

test('trend_analysis：提供 time_column 时产出时间轴处理脚本', () => {
    const script = tmp().generate('trend_analysis', { value_column: '销售额', time_column: '日期', forecast_steps: 3 });
    assert.ok(script.includes('parse_dt'), '应含日期解析函数');
    assert.ok(script.includes("from datetime import datetime"), '应引入 datetime');
    assert.ok(script.includes('has_time_axis'), '应计算时间轴标记');
    assert.ok(script.includes("'按时间轴'"), '应区分时间轴模式');
    assert.ok(script.includes('tcol = "日期"'), '应注入时间列名');
    assert.ok(script.includes('forecast_x'), '应输出预测时间标签');
    assert.ok(script.includes('short') === false, '不应出现未知字段');
});

test('trend_analysis：未提供 time_column 时保持行序回退', () => {
    const script = tmp().generate('trend_analysis', { value_column: '销售额' });
    assert.ok(script.includes('tcol = ""'), '时间列应为空');
    assert.ok(script.includes('按行序'), '应回退行序模式');
});

test('detectDateColumn：识别日期列，忽略数值列', () => {
    const V = tmp();
    const headers = ['日期', '部门', '地区', '销售额', '成本', '利润'];
    const rows = [['2023-01-01', '销售部', '北京', 1000, 600, 400],
        ['2023-01-02', '销售部', '上海', 1200, 700, 500],
        ['2023-01-03', '销售部', '广州', 1100, 650, 450]];
    const r = V.detectDateColumn(headers, rows);
    assert.strictEqual(r, '日期', `应识别日期列"日期"，实际: ${r}`);
    const rowsNoDate = [['销售部', '北京', 1000], ['销售部', '上海', 1200], ['销售部', '广州', 1100]];
    const r2 = V.detectDateColumn(['部门', '地区', '销售额'], rowsNoDate);
    assert.strictEqual(r2, '', '无日期列时不应误判');
});

test('trend_analysis：validate 接受 time_column 为合法列', () => {
    const V = tmp();
    const ok = V.validateTemplateParams('trend_analysis', { value_column: '销售额', time_column: '日期' }, HEADERS, NUMERIC);
    assert.strictEqual(ok.ok, true);
    assert.strictEqual(ok.invalidColumns.length, 0);
    const bad = V.validateTemplateParams('trend_analysis', { value_column: '销售额', time_column: '利润率' }, HEADERS, NUMERIC);
    assert.ok(bad.invalidColumns.length > 0, '无效时间列应被检出');
});

test('trend_analysis：generate 脚本含 Excel 序列号转日期逻辑', () => {
    const script = tmp().generate('trend_analysis', { value_column: '销售额', time_column: '日期' });
    assert.ok(script.includes('1899, 12, 30'), '应含 Excel 序列号基准日期');
    assert.ok(script.includes('timedelta(days=n)'), '应按天数转日期');
    assert.ok(script.includes('1 <= n <= 100000'), '应限制序列号合理范围');
});

test('detectDateColumn：识别 Excel 序列号日期列，且不误伤普通数值列', () => {
    const V = tmp();
    // 列名带日期语义 + 序列号数值（45123 ≈ 2023-07-21）
    const serRows = [['日期', '销售额'], ['43466', 1000], ['43467', 2000], ['43468', 3000]];
    const r = V.detectDateColumn(serRows[0], serRows.slice(1));
    assert.strictEqual(r, '日期', `应识别序列号日期列，实际: ${r}`);
    // 普通数值列（列名不含日期语义）且值是纯数字 → 不应误判为日期列
    const numRows = [['部门', '地区', '销售额'], ['销售部', '北京', 1000], ['销售部', '上海', 2000], ['销售部', '广州', 3000]];
    const r2 = V.detectDateColumn(numRows[0], numRows.slice(1));
    assert.strictEqual(r2, '', '无日期语义的纯数值列不应判为日期');
});

test('columnGuardMessage：列护栏文案回传可疑列与可用列清单', () => {
    const V = tmp();
    const msg = V.columnGuardMessage(['利润率'], ['日期', '部门', '地区', '销售额', '成本', '利润']);
    assert.ok(msg.includes('利润率'), '应点名可疑列');
    assert.ok(msg.includes('可用列'), '应标注可用列');
    assert.ok(msg.includes('销售额'), '可用列清单应含真实列');
    const msg2 = V.columnGuardMessage([], []);
    assert.ok(msg2.includes('未知占位符'), '空可疑列时以占位符提示');
});

test('shouldForceCustomInput 路由判定：模板优先，仅强"自定义"信号走custom', () => {
    const f = tmp().shouldForceCustomInput;
    // 带"用python" + 明确活 → false → 交给模板选择器（分组/异常/趋势/回归/统计）
    assert.strictEqual(f('用Python对当前数据做异常检测'), false, '异常检测应走模板');
    assert.strictEqual(f('用python分析销售额趋势'), false, '趋势应走模板');
    assert.strictEqual(f('写个python脚本按地区汇总销售额'), false, '分组汇总应走模板');
    assert.strictEqual(f('用python做线性回归'), false, '回归应走模板');
    assert.strictEqual(f('python脚本做描述统计'), false, '描述统计应走模板');
    assert.strictEqual(f('用Python对当前数据做分组聚合统计'), false, '分组聚合应走模板');
    // 无明确模板但也没有强自定义信号 → 也交给模板选择器（选不到再兜底 custom）
    assert.strictEqual(f('用python算一下我觉得奇怪的东西'), false, '无强自定义信号应走模板选择器');
    // 强烈表达"自由/自定义写脚本" → true → custom
    assert.strictEqual(f('用python做点自定义的自由处理'), true, '自定义信号应走 custom');
    assert.strictEqual(f('用python自己写个脚本处理'), true, '自己写信号应走 custom');
    assert.strictEqual(f('python随便处理完全自己写'), true, '完全自己/随便写信号应走 custom');
    // 不带 python → 不强制 custom
    assert.strictEqual(f('异常检测'), false, '无python不强制 custom');
    assert.strictEqual(f('分组聚合统计'), false, '无python不强制 custom');
});

test('templateResultToMarkdown：对象数组→表格，矩阵→表格，纯对象→键值，标量→原样', () => {
    const md = tmp().templateResultToMarkdown;
    // 同构对象数组 → markdown 表格
    const t1 = md([{ '地区': '北京', '销售额': 3000 }, { '地区': '上海', '销售额': 5000 }]);
    assert.ok(t1.includes('| 地区 | 销售额 |'), '对象数组应生成表头');
    assert.ok(t1.includes('北京'), '应含行数据');
    assert.strictEqual(t1.includes('```'), false, '表格不应是 json 块');
    // 二维矩阵（相关性）→ 表格
    const t2 = md([[1, 0.8], [0.8, 1]]);
    assert.ok(t2.includes('| # |'), '矩阵应生成索引表头');
    assert.ok(t2.includes('0.8'), '应含矩阵数值');
    // 纯对象 → 键值列表
    const t3 = md({ n: 10, r2: 0.94, slope: 1.5 });
    assert.ok(t3.includes('- **n**：10'), '纯对象应生成键值列表');
    // 标量 → 原样
    assert.strictEqual(md('hello'), 'hello');
});

test('shouldForceCustomInput：带python但明确活→false(走模板)，强自定义→true', () => {
    const sc = tmp().shouldForceCustomInput;
    assert.strictEqual(sc('用Python对当前数据做分组聚合统计'), false, '带python+明确活应走模板');
    assert.strictEqual(sc('用python做异常检测'), false, '异常检测应走模板');
    assert.strictEqual(sc('用python分析销售额趋势'), false, '趋势应走模板');
    assert.strictEqual(sc('写python脚本算均值'), false, '描述统计应走模板');
    assert.strictEqual(sc('用python自己写个脚本处理数据'), true, '明确自定义应走 custom');
    assert.strictEqual(sc('用python随便分析点什么'), true, '随意/自由应走 custom');
    assert.strictEqual(sc('异常检测'), false, '不带python永不强制 custom');
});

test('matchTemplateByIntent：明确意图确定性绑定模板', () => {
    const mt = tmp().matchTemplateByIntent;
    assert.deepStrictEqual(mt('用Python对当前数据做分组聚合统计'), { id: 'group_aggregation' }, '分组聚合→group_aggregation');
    assert.deepStrictEqual(mt('按地区汇总销售额'), { id: 'group_aggregation' }, '按地区汇总→group_aggregation');
    assert.deepStrictEqual(mt('异常检测'), { id: 'anomaly' }, '异常→anomaly');
    assert.deepStrictEqual(mt('线性回归分析成本与利润'), { id: 'linear_regression' }, '回归→linear_regression');
    assert.deepStrictEqual(mt('销售额相关性分析'), { id: 'correlation' }, '相关性→correlation');
    assert.deepStrictEqual(mt('按日期分析销售额趋势'), { id: 'trend' }, '趋势→trend');
    assert.deepStrictEqual(mt('描述统计'), { id: 'descriptive_stats' }, '描述统计→descriptive_stats');
    assert.deepStrictEqual(mt('数据分布'), { id: 'distribution' }, '分布→distribution');
    assert.strictEqual(mt('算一下奇怪的东西'), null, '无明确意图→null（交给模板选择器/兜底）');
});

test('validateTemplateParams：日期列不可作 value_column（numericHeaders 剔除日期后）', () => {
    const { validateTemplateParams } = tmp();
    const headers = ['日期', '部门', '地区', '销售额', '成本', '利润', '客户数'];
    // 仿真 execute：numericHeaders 已剔除日期列
    const r = validateTemplateParams(
        'group_aggregation',
        { value_column: '日期', group_column: '部门', aggregation: 'sum' },
        headers, ['销售额', '成本', '利润', '客户数']
    );
    const dateHit = (r.invalidColumns || []).find(c => c.name === '日期');
    assert.ok(dateHit, '日期当 value_column 应判无效列');
    assert.strictEqual(dateHit.numeric, true, '应标记为非数值列待修');
    // 正确用法应通过
    const ok = validateTemplateParams(
        'group_aggregation',
        { value_column: '销售额', group_column: '部门', aggregation: 'sum' },
        headers, ['销售额', '成本', '利润']
    );
    assert.strictEqual(ok.invalidColumns.length, 0, '销售额数值列应通过');
});

test('buildTemplateChartCfg：分组/分布→bar，趋势→line，无法构造→null', () => {
    const b = tmp().buildTemplateChartCfg;
    // 分组聚合结果：对象数组（地区|销售额）
    const g = b('group_aggregation', [{ 地区: '华北', 销售额: 3000 }, { 地区: '华东', 销售额: 5000 }]);
    assert.strictEqual(g.type, 'bar', '分组聚合应是柱状图');
    assert.deepStrictEqual(g.labels, ['华北', '华东']);
    assert.deepStrictEqual(g.values, [3000, 5000]);
    assert.ok(String(g.title).includes('销售额'), '标题应变销售额');
    // 趋势结果：对象带 labels/values
    const t = b('trend_analysis', { labels: ['2024-06', '2024-07'], values: [100, 130] });
    assert.strictEqual(t.type, 'line', '趋势应是折线图');
    assert.strictEqual(b('data_distribution', { 销售额: { min: 0, max: 10, bins: 2, distribution: [{ range: '0-5', count: 3 }, { range: '5-10', count: 7 }] } }).type, 'bar', '分布应是柱状图');
    // 无法构造：非数组/无数值 → null（安全回退，不画图）
    assert.strictEqual(b('group_aggregation', { message: '数据不足' }), null, '无结构化数据应返回 null');
    assert.strictEqual(b('descriptive_stats', [{ 列: '销售额', 均值: 123 }]), null, '描述统计不映射图类型应 null');
});

test('buildTemplateChartCfg：平面对象映射 {类别:数值} 也能生成柱状图', () => {
    const b = tmp().buildTemplateChartCfg;
    const m = b('group_aggregation', { 华北: 3000, 华东: 5000, 华南: 2000 });
    assert.strictEqual(m.type, 'bar', '平面对象映射应是柱状图');
    assert.deepStrictEqual(m.labels, ['华北', '华东', '华南']);
    assert.deepStrictEqual(m.values, [3000, 5000, 2000]);
    assert.strictEqual(b('group_aggregation', { message: '数据不足' }), null, '无不含数值的平面对象应返回 null');
});

test('buildTemplateChartCfg：分组聚合 dict-of-dicts 标题按 agg 显示真实指标名(avg→平均值, sum→求和)', () => {
    const b = tmp().buildTemplateChartCfg;
    const avg = b('group_aggregation', { 食品: { value: 31558.58, count: 10, agg: 'avg' }, 服装: { value: 31634.92, count: 9, agg: 'avg' } });
    assert.strictEqual(avg.title, '平均值', 'avg 标题应为平均值, 实际=' + avg.title);
    const sum = b('group_aggregation', { 华北: { value: 15000, count: 3, agg: 'sum' }, 华东: { value: 20000, count: 4, agg: 'sum' } });
    assert.strictEqual(sum.title, '求和', 'sum 标题应为求和, 实际=' + sum.title);
    // 无 agg 字段时降级为字段名，不再强制英文 value 表现无意义
    const noagg = b('group_aggregation', { 食品: { value: 10, count: 1 }, 服装: { value: 20, count: 1 } });
    assert.ok(noagg.title === 'value' || noagg.title === '平均值', '无 agg 时应降级, 实际=' + noagg.title);
});

test('每个模板的图表映射：trend_analysis→双序列折线(历史+预测), data_distribution→频数直方图, linear_regression→拟合线, correlation→热力图', () => {
    const b = tmp().buildTemplateChartCfg;
    // 趋势（带历史序列）→ 双系列折线：历史数据（实线）+ 预测数据（虚线），x轴合并，预测线在最后历史点衔接
    const trend = b('trend_analysis', {
        column: '销售额', data_points: 4, direction: '上升', has_time_axis: true,
        historical_x: ['2024-01-01', '2024-01-02', '2024-01-03', '2024-01-04'],
        historical_values: [100, 110, 120, 130],
        forecast: [140, 150, 160],
        forecast_x: ['2024-01-05', '2024-01-06', '2024-01-07']
    });
    assert.strictEqual(trend.type, 'line', '趋势应为折线图');
    assert.deepStrictEqual(trend.labels, ['2024-01-01','2024-01-02','2024-01-03','2024-01-04','2024-01-05','2024-01-06','2024-01-07'], 'x轴应为历史+预测合并');
    assert.ok(trend.title.includes('趋势与预测') && trend.title.includes('销售额'), '趋势标题应为 列名·趋势与预测, 实际=' + trend.title);
    assert.ok(Array.isArray(trend.series) && trend.series.length === 2, '应为两个系列, 实际=' + (trend.series && trend.series.length));
    assert.strictEqual(trend.series[0].name, '历史数据');
    assert.strictEqual(trend.series[1].name, '预测数据');
    assert.deepStrictEqual(trend.series[0].data, [100, 110, 120, 130, null, null, null], '历史系列：前4个真实值+后3个空');
    assert.deepStrictEqual(trend.series[1].data, [null, null, null, 130, 140, 150, 160], '预测系列：在最后历史点130处衔接，然后预测值');
    assert.strictEqual(trend.series[1].lineStyle && trend.series[1].lineStyle.type, 'dashed', '预测线应为虚线');
    // 趋势（无历史，只有预测）→ 降级单条折线，保持兼容
    const trendOnly = b('trend_analysis', { column: '销售额', forecast: [101, 102, 103], forecast_x: ['2024-01-01', '2024-01-02', '2024-01-03'] });
    assert.strictEqual(trendOnly.type, 'line');
    assert.strictEqual(trendOnly.series, undefined, '无历史时不应有多系列');
    assert.deepStrictEqual(trendOnly.labels, ['2024-01-01', '2024-01-02', '2024-01-03']);
    assert.ok(trendOnly.title.includes('未来预测'), '降级时标题为未来预测, 实际=' + trendOnly.title);
    // 分布：首个含 distribution 的列 → bar 直方图
    const dist = b('data_distribution', { 销售额: { min: 0, max: 100, bins: 3, distribution: [{ range: '0-33', count: 5, percentage: 10 }, { range: '33-66', count: 8, percentage: 16 }] } });
    assert.strictEqual(dist.type, 'bar', '数据分布应为柱状图');
    assert.deepStrictEqual(dist.labels, ['0-33', '33-66']);
    assert.deepStrictEqual(dist.values, [5, 8]);
    assert.ok(dist.title.includes('频数分布'), '分布标题应含 频数分布, 实际=' + dist.title);
    // 回归：forecast[{x,y}] → line 拟合线
    const reg = b('linear_regression', { n: 5, slope: 1, intercept: 2, r2: 0.99, equation: 'y=1*x+2', forecast: [{ x: 1, y: 3 }, { x: 2, y: 4 }, { x: 3, y: 5 }] });
    assert.strictEqual(reg.type, 'line', '回归应为折线图');
    assert.deepStrictEqual(reg.labels, ['1', '2', '3']);
    assert.deepStrictEqual(reg.values, [3, 4, 5]);
    assert.ok(reg.title.includes('y=1*x+2'), '回归标题应含方程, 实际=' + reg.title);
    // 相关性：矩阵 → heatmap，data 长度 = n*n
    const heat = b('correlation', { 销售额: { 销售额: 1, 利润: 0.8 }, 利润: { 销售额: 0.8, 利润: 1 } });
    assert.strictEqual(heat.type, 'heatmap', '相关性应为热力图');
    assert.deepStrictEqual(heat.xLabels, ['销售额', '利润']);
    assert.strictEqual(heat.data.length, 4, '热力图数据应为 n*n 个点');
    const diagOk = heat.data.every(function (p) { return p[0] === p[1] ? p[2] === 1 : Math.abs(p[2] - 0.8) < 1e-9; });
    assert.ok(diagOk, '热力图对角线=1, 两两相关=0.8');
    // 无对应图能力的模板（descriptive_stats/anomaly_detection）应返回 null（安全回退仅表格）
    assert.strictEqual(b('descriptive_stats', { 销售额: { count: 3, mean: 1 } }), null);
    assert.strictEqual(b('anomaly_detection', { 销售额: { total: 5 } }), null);
});

test('buildResultChartDiv：模板执行结果 → 可见 data-echarts 节点，无法构造→空串', () => {
    const d = tmp().buildResultChartDiv;
    const div = d('group_aggregation', {
        ok: true,
        result: JSON.stringify({ success: true, result: [{ 地区: '华北', 销售额: 3000 }, { 地区: '华东', 销售额: 5000 }] })
    });
    assert.ok(div.includes('data-echarts'), '应生成 data-echarts 节点');
    assert.ok(div.includes('<div'), '应包含 div 节点');
    assert.strictEqual(d('group_aggregation', { ok: false, error: 'x' }), '', '失败结果不生成图表');
    assert.strictEqual(d('', { ok: true }), '', '无 templateId 不生成图表');
    assert.strictEqual(d('group_aggregation', { ok: true, result: '{this is not json', output: '无数据' }), '',
        '无法构造时应返回空串安全回退');
});

test('buildTemplateChartCfg：分组聚合实况 {组名:{value,count,agg}} 字典套字典也能出柱状图', () => {
    const b = tmp().buildTemplateChartCfg;
    const real = {
        家居: { value: 27410.49, count: 5, agg: 'avg' },
        服装: { value: 32591.26, count: 6, agg: 'avg' },
        办公: { value: 33915.682, count: 4, agg: 'avg' }
    };
    const r = b('group_aggregation', real);
    assert.strictEqual(r.type, 'bar', '字典套字典应是柱状图');
    assert.deepStrictEqual(r.labels, ['家居', '服装', '办公']);
    assert.deepStrictEqual(r.values, [27410.49, 32591.26, 33915.682]);
    assert.strictEqual(r.title, '平均值', '标题应按 agg=avg 显示真实指标名(平均值), 实际=' + r.title);
});

test('buildResultChartDiv：相关性矩阵（带"Python脚本执行成功"前置文本）也能生成热力图节点', () => {
    const d = tmp().buildResultChartDiv;
    const raw = 'Python 脚本执行成功\r\n' + JSON.stringify({
        销售额: { 销售额: 1, 成本: 0.9546, 利润: 0.8096 },
        成本: { 销售额: 0.9546, 成本: 1, 利润: 0.5981 },
        利润: { 销售额: 0.8096, 成本: 0.5981, 利润: 1 }
    });
    const div = d('correlation', { ok: true, result: raw });
    assert.ok(div && div.includes('data-echarts'), '相关性(带前置文本)应生成 data-echarts, 实际 len=' + String(div).length);
    const m = div.match(/data-echarts='([\s\S]*?)'/);
    assert.ok(m, '应能提取相关性配置');
    const cfg = JSON.parse(m[1]);
    assert.strictEqual(cfg.type, 'heatmap', '应为热力图');
    assert.strictEqual(cfg.data.length, 9, '3×3 矩阵应为 9 个数据点');
});

test('buildResultChartDiv：分组聚合实况 JSON（result 为字典套字典）能生成 data-echarts', () => {
    const d = tmp().buildResultChartDiv;
    const div = d('group_aggregation', {
        ok: true,
        result: JSON.stringify({
            success: true,
            summary: '按部门分组',
            result: { 家居: { value: 27410.49, count: 5, agg: 'avg' }, 办公: { value: 33915.682, count: 4, agg: 'avg' } }
        })
    });
    assert.ok(div.includes('data-echarts'), '字典套字典结果应生成 data-echarts');
    assert.ok(div.includes('家居') && div.includes('办公'), '标签应含分组名');
});

test('buildChartDiv：图表外包横向滚动容器，分类少自适应/分类多给最小宽度防右侧被裁', () => {
    const d = tmp().buildChartDiv;
    // 分类少(2个)：应保持 100% 宽自适应填满气泡，不引入滚动
    const few = d({ type: 'bar', labels: ['食品', '服装'], values: [1, 2] });
    assert.ok(few.startsWith('\n<div style="width:100%;max-width:100%;overflow-x:auto;'), '少分类也应包横向滚动容器');
    assert.ok(few.includes('width:100%;"'), '少分类内层应为 100% 宽(自适应)');
    // 分类很多(30个)：内层应给定最小宽度，外层可横向滚动看全，不再被 overflow:hidden 右侧裁掉
    const manyLabels = [];
    for (let i = 0; i < 30; i++) manyLabels.push('类别' + i);
    const many = d({ type: 'bar', labels: manyLabels, values: manyLabels.map(function (_, i) { return i; }) });
    assert.ok(many.includes('overflow-x:auto'), '多分类外层应有横向滚动容器');
    const m = many.match(/style="min-width:100%;height:300px;width:(\d+)px;/);
    assert.ok(m, '多分类内层应给数值最小宽度');
    assert.ok(Number(m[1]) >= 420, '多分类最小宽度应 ≥420px, 实际=' + m[1]);
    const e = many.match(/data-echarts='([\s\S]*?)'/);
    assert.ok(e, '多分类仍应保留 data-echarts 节点');
    const cfg = JSON.parse(e[1]);
    assert.strictEqual(cfg.labels.length, 30, '标签应完整保留 30 个');
});

test('toTypedValue：纯数值→number，含%/中文/日期→保留字符串', () => {
    const tv = loadChatV2().App.ChatV2.toTypedValue;
    assert.ok(tv && typeof tv === 'function', '应暴露 toTypedValue');
    assert.strictEqual(tv('73'), 73, '整数应转 number');
    assert.strictEqual(tv('0'), 0, '零应转 number');
    assert.strictEqual(tv('0.9572'), 0.9572, '小数应转 number');
    assert.strictEqual(tv('-1,234.5'), -1234.5, '千分位负小数应转 number');
    assert.strictEqual(tv('1.5e3'), 1500, '科学计数应转 number');
    // 非纯数值保持原样（异常检测的比率列）
    assert.strictEqual(tv('0.0%'), '0.0%', '含%不应转 number');
    assert.strictEqual(tv('北京'), '北京', '中文不应转 number');
    assert.strictEqual(tv('2023-01-01'), '2023-01-01', '日期不应转 number');
    assert.strictEqual(tv('销售额'), '销售额', '文本表头不应转 number');
});

test('buildResultChartDiv：Python dict repr 结果（单引号+True）也能解析出图表', () => {
    // 回归：部分 Python 执行工具把 output 以 str() 回传（单引号、大写 True），
    // 此前 JS JSON.parse 解析失败 → divLen=0 → 图表全都不显示。加 parsePythonRepr 兜底修复。
    const d = tmp().buildResultChartDiv;
    const pyrepr = "{'success': True, 'summary': '分组聚合按部门', 'result': {'家居': {'value': 27410.49, 'count': 18, 'agg': 'avg'}, '服装': {'value': 32591.26, 'count': 15, 'agg': 'avg'}}}";
    const div = d('group_aggregation', { ok: true, result: pyrepr });
    assert.ok(div.includes('data-echarts'), 'Python repr 结果应能生成 data-echarts');
    assert.ok(div.includes('家居') && div.includes('服装'), '应解析出分组标签');
    // 解析出的配置应是合法柱状图
    const m = div.match(/data-echarts='([\s\S]*?)'/);
    assert.ok(m, '应能提取图表配置');
    const cfg = JSON.parse(m[1]);
    assert.strictEqual(cfg.type, 'bar', '分组聚合应为柱状图');
    assert.deepStrictEqual(cfg.values, [27410.49, 32591.26], '应提取出数值序列');
});

test('buildResultChartDiv：后端"JSON 变体"（前置提示 + 成员缺逗号）也能解析出图表', () => {
    // 回归：真实后端把结果回传为 "Python 脚本执行成功\n{...}"，且成员间缺逗号，
    // 标准 JSON.parse / extractJson 都失败 → 图表不显示。加 parseLenientJson 容错解析修复。
    const d = tmp().buildResultChartDiv;
    const raw = 'Python 脚本执行成功\r\n{\r\n "食品": {\r\n "value": 31558.58\r\n "count": 10\r\n "agg": "avg"\r\n }\n "服装": {\r\n "value": 31634.9189\r\n "count": 9\r\n "agg": "avg"\r\n }\n}';
    const div = d('group_aggregation', { ok: true, result: raw });
    assert.ok(div.includes('data-echarts'), '后端变体结果应能生成 data-echarts');
    const m = div.match(/data-echarts='([\s\S]*?)'/);
    assert.ok(m, '应能提取图表配置');
    const cfg = JSON.parse(m[1]);
    assert.strictEqual(cfg.type, 'bar', '应为柱状图');
    assert.deepStrictEqual(cfg.labels, ['食品', '服装'], '应解析出分组标签');
    assert.deepStrictEqual(cfg.values, [31558.58, 31634.9189], '应提取出数值序列');
});