/**
 * ============================================================
 * Tests/table-style-presets.test.js — 表格「样式方案」回归测试
 * ============================================================
 * 背景（用户反馈 2026-10-07）：
 *   以前的"主题"只有颜色不同（商务蓝/科技绿/…），版式一模一样 → 换主题等于换配色。
 * 修复：新增样式方案预设（config/table-style-presets.js），三个宿主共享：
 *   ① 版式 layout（三线表/无边框留白/卡片条纹/深表头/首列强调/合计行/紧凑/杂志…）
 *   ② 配色 palette（8 套）
 *   ③ 密度 density（紧凑 15-16 / 标准 20-22 / 宽松 28-30 磅）
 *   ④ 开关：隔行底纹、首列加粗、末行合计、数字右对齐、垂直居中
 * 本文件验证：预设识别 → Excel spec 字段 → VBA 渲染；Word/PPT 工具参数。
 * 运行：node Tests/table-style-presets.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');
const { loadChatV2, BASE } = require('./helpers/harness');

/** 加载预设模块（独立环境） */
function loadPresets() {
    const { App } = loadChatV2(['config/table-style-presets.js']);
    return App.ChatV2.tableStylePresets;
}

/** 加载预设 + Excel Spec 引擎（共享同一个 window/App） */
function loadEngineWithPresets() {
    const { window, App } = loadChatV2(['config/table-style-presets.js', 'services/skill-spec-engine.js']);
    return { App, engine: App.ChatV2.services.skillSpecEngine };
}

/** 加载预设 + dispatcher（Word/PPT 硬短路参数） */
function loadDispatcherWithPresets() {
    // routing-constants 提供硬短路关键词常量（FMT_VERBS_HOST 等），必须一并加载
    const { window, App } = loadChatV2(['config/routing-constants.js', 'config/table-style-presets.js', 'services/mcp-dispatcher.js']);
    return { window, App, dispatcher: App.ChatV2.mcpDispatcher };
}

test('样式预设：版式识别 — 三线表 / 无边框 / 卡片条纹 / 深表头', () => {
    const P = loadPresets();
    const d1 = P.resolve('把当前表格做成三线表');
    assert.strictEqual(d1.layoutId, 'three_line');
    assert.strictEqual(d1.borderStyle, 'threeLine');
    assert.strictEqual(d1.headerMode, 'underline');
    assert.strictEqual(d1.zebra, false);

    const d2 = P.resolve('无边框留白版式');
    assert.strictEqual(d2.layoutId, 'borderless');
    assert.strictEqual(d2.borderStyle, 'none');
    assert.strictEqual(d2.density, 'loose');

    const d3 = P.resolve('卡片条纹');
    assert.strictEqual(d3.layoutId, 'card_stripe');
    assert.strictEqual(d3.zebra, true);

    const d4 = P.resolve('深表头');
    assert.strictEqual(d4.layoutId, 'deep_header');
    assert.strictEqual(d4.headerMode, 'deep');
});

test('样式预设：配色与版式可组合（科技绿 + 三线表）', () => {
    const P = loadPresets();
    const d = P.resolve('用科技绿三线表美化');
    assert.strictEqual(d.layoutId, 'three_line');
    assert.strictEqual(d.paletteId, 'tech_green');
    const args = P.toHostArgs(d);
    assert.strictEqual(args.borderStyle, 'threeLine');
    assert.strictEqual(args.borderColor, P.PALETTES.tech_green.border);
});

test('样式预设：密度决定行高（紧凑 / 标准 / 宽松）', () => {
    const P = loadPresets();
    assert.strictEqual(P.toHostArgs(P.resolve('紧凑报表')).rowHeight, P.ROW_HEIGHT.host.compact);
    assert.strictEqual(P.toHostArgs(P.resolve('美化表格')).rowHeight, P.ROW_HEIGHT.host.normal);
    assert.ok(P.toHostArgs(P.resolve('宽松一点')).rowHeight >= P.ROW_HEIGHT.host.loose);
    assert.strictEqual(P.toExcelSpec(P.resolve('紧凑报表')).rowHeight, P.ROW_HEIGHT.excel.compact);
});

test('样式预设：强调开关（首列 / 合计行 / 数字右对齐）', () => {
    const P = loadPresets();
    const s1 = P.toExcelSpec(P.resolve('首列加粗'));
    assert.strictEqual(s1.boldFirstColumn, true);
    const s2 = P.toExcelSpec(P.resolve('加个合计行'));
    assert.strictEqual(s2.emphasizeLastRow, true);
    assert.strictEqual(s2.bodyNumberRight, true);
    const s3 = P.toExcelSpec(P.resolve('数字右对齐'));
    assert.strictEqual(s3.bodyNumberRight, true);
    const a1 = P.toHostArgs(P.resolve('首列强调'));
    assert.strictEqual(a1.firstColBold, true);
    const a2 = P.toHostArgs(P.resolve('合计行'));
    assert.strictEqual(a2.lastRowBold, true);
});

test('Excel：specFromMessage 识别版式 → 渲染出对应 VBA（三线表无竖线无底色）', () => {
    const { engine } = loadEngineWithPresets();
    const spec = engine.specFromMessage('把当前表格做成三线表');
    assert.strictEqual(spec.borderStyle, 'threeLine', '应识别为三线表版式');
    assert.strictEqual(spec.headerLayout, 'underline');

    const code = engine.renderBeautifyVBA(spec, {});
    assert.ok(/xlEdgeTop/.test(code), '三线表应有顶线/底线');
    assert.ok(/xlEdgeBottom/.test(code), '三线表应有底线');
    assert.ok(!/\.Borders\.LineStyle = xlContinuous\s*$/m.test(code) || /xlEdgeTop/.test(code));
    assert.ok(!/MOD\(ROW\(\),2\)=0/.test(code), '三线表默认不加斑马纹');
    // 数据区不留整片底色（留白）
    assert.ok(/\.Interior\.Pattern = xlNone/.test(code), '三线表数据区应无底色');
});

test('Excel：stylePreset 一键展开（只给预设 id 也能得到完整排版字段）', () => {
    const { engine } = loadEngineWithPresets();
    const spec = engine.normalizeSpec({ stylePreset: 'compact_report' });
    assert.strictEqual(spec.borderStyle, 'all');
    assert.strictEqual(spec.bodyNumberRight, true);
    assert.ok(spec.rowHeight && spec.rowHeight <= 16, '紧凑版式行高应压紧');
    const code = engine.renderBeautifyVBA(spec, {});
    assert.ok(/RowHeight = /.test(code), '紧凑版式应写行高');
});

test('Excel：不同版式渲染出不同 VBA（不再只有颜色不同）', () => {
    const { engine } = loadEngineWithPresets();
    const a = engine.renderBeautifyVBA(engine.normalizeSpec({ stylePreset: 'three_line' }), {});
    const b = engine.renderBeautifyVBA(engine.normalizeSpec({ stylePreset: 'card_stripe' }), {});
    const c = engine.renderBeautifyVBA(engine.normalizeSpec({ stylePreset: 'compact_report' }), {});
    assert.notStrictEqual(a, b);
    assert.notStrictEqual(b, c);
    assert.ok(/xlEdgeTop/.test(a) && !/MOD\(ROW\(\),2\)=0/.test(a), '三线表：三线无斑马纹');
    assert.ok(/MOD\(ROW\(\),2\)=0/.test(b), '卡片条纹：应有整行浅底');
    assert.ok(/RowHeight = /.test(c), '紧凑报表：应压行高');
});

test('Excel：白名单防护 — 非法版式/行高被忽略', () => {
    const { engine } = loadEngineWithPresets();
    const s = engine.normalizeSpec({ borderStyle: 'DROP TABLE', headerLayout: '<script>', rowHeight: 9999, stylePreset: 'evil' });
    assert.strictEqual(s.borderStyle, null);
    assert.strictEqual(s.headerLayout, null);
    assert.strictEqual(s.rowHeight, null);
    assert.strictEqual(s.stylePreset, null);
});

test('Word/PPT：硬短路参数带上版式（三线表 → borderStyle=threeLine、表头无底色）', async () => {
    const { window, dispatcher } = loadDispatcherWithPresets();
    window.SmartPilotHost = { isExcel: false, isPowerPoint: false, kind: 'word' };
    const res = await dispatcher.tryDispatch({
        message: '把当前选中的表格做成三线表',
        toolRegistry: null,
        sessionStore: null,
        turn: { context: { selectedObject: { selectionType: 'table', rowCount: 4, colCount: 5, summary: '表格 4行5列' } } }
    });
    assert.strictEqual(res.hardMatched, true, '应命中选中表格格式化硬短路');
    assert.strictEqual(res.toolName, 'word_format_selected_table_cell');
    assert.strictEqual(res.args.borderStyle, 'threeLine', '应带上边框版式');
    assert.strictEqual(res.args.headerFillColor, '', '三线表表头不要底色');
    assert.strictEqual(res.args.bodyAlign, 'left');
});

test('Word/PPT：主题色仍然生效（商务蓝 → 蓝边框 + 浅蓝表头）', async () => {
    const { window, dispatcher } = loadDispatcherWithPresets();
    window.SmartPilotHost = { isExcel: false, isPowerPoint: true, kind: 'powerpoint' };
    const res = await dispatcher.tryDispatch({
        message: '用商务蓝主题美化当前选中的表格',
        toolRegistry: null,
        sessionStore: null,
        turn: { context: { selectedObject: { selectionType: 'table', rowCount: 3, colCount: 4, summary: '表格 3行4列' } } }
    });
    assert.strictEqual(res.hardMatched, true);
    assert.strictEqual(res.toolName, 'ppt_format_selected_table_cell');
    assert.strictEqual(res.args.borderColor, '#4472C4', '商务蓝边框色');
    assert.strictEqual(res.args.headerFillColor, '#D9E2F3', '商务蓝浅表头');
    assert.strictEqual(res.args.zebra, true, '经典网格默认隔行底纹');
    assert.strictEqual(res.args.rowHeight > 0, true, '应带行高（排版维度）');
});

test('技能文档：三个宿主的技能都声明了新版式字段', () => {
    const root = path.join(__dirname, '..');
    const beautifier = fs.readFileSync(path.join(root, 'Skills', 'SmartTableBeautifier.md'), 'utf8');
    for (const f of ['stylePreset', 'borderStyle', 'headerLayout', 'rowHeight', 'boldFirstColumn', 'emphasizeLastRow', 'bodyNumberRight']) {
        assert.ok(beautifier.includes(f), `SmartTableBeautifier.md 应声明字段 ${f}`);
    }
    for (const md of ['Skills/WordTableThemer.md', 'Skills/PptTableStyle.md']) {
        const txt = fs.readFileSync(path.join(root, md), 'utf8');
        for (const p of ['borderStyle', 'headerFontColor', 'rowHeight', 'valign', 'bodyAlign', 'firstColBold', 'lastRowBold']) {
            assert.ok(txt.includes(p), `${md} 应说明参数 ${p}`);
        }
    }
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); }
