/**
 * ============================================================
 * Tests/tool-required-args.test.js — 通用对话必填参数"推断+反问结合"恢复策略
 * ============================================================
 * 覆盖：
 *   1) 可安全推断的参数自动补全后执行（如 excel_create_sheet.name → "新建工作表N"）
 *   2) 无法兜底的必填参数拦截为 needsInput（不执行，交给 agent-loop 反问用户）
 *   3) 仅通用对话/文本路径生效（kind=text）；skill 路径不受影响
 * 运行：node Tests/tool-required-args.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadChatV2, stubDangerConfirm } = require('./helpers/harness');

test('必填参数自动补全：excel_create_sheet.name 缺失时自动生成唯一默认表名并执行', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '帮我创建一个新的工作表',
        kind: 'text',
        context: { sheets: ['Sheet1', '新建工作表1'] }
    });
    assert.strictEqual(turn.kind, 'text');

    let executedArgs = null;
    tr.register({
        name: 'excel_create_sheet', displayName: '新建工作表', description: '创建一个新工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '新工作表名称' } },
            required: ['name']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { created: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_create_sheet', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true, '缺 name 时应自动补全并执行成功');
    assert.ok(executedArgs, '工具应被真正执行');
    assert.strictEqual(executedArgs.name, '新建工作表2', '应跳过已存在的 Sheet1 与 新建工作表1，生成 新建工作表2');
    const tc = turn.toolCalls.find(c => c.name === 'excel_create_sheet');
    assert.ok(tc, '应创建 toolCall 记录');
    assert.strictEqual(tc.status, 'completed');
});

test('必填参数自动补全：用户消息里显式指定表名时应提取并采用，而非兜底"新建工作表N"', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '帮我创建一个新的工作表名为"新工作表"',
        kind: 'text',
        context: { sheets: ['Sheet1'] }
    });

    let executedArgs = null;
    tr.register({
        name: 'excel_create_sheet', displayName: '新建工作表', description: '创建一个新工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '新工作表名称' } },
            required: ['name']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { created: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_create_sheet', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true);
    assert.ok(executedArgs, '工具应被真正执行');
    assert.strictEqual(executedArgs.name, '新工作表', '应提取用户显式指定的表名，而非改用默认名');
});

test('无法兜底的必填参数 → 拦截为 needsInput（反问用户，不执行）', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: '导出数据到文件', kind: 'text', context: {} });

    let executed = false;
    tr.register({
        name: 'excel_export_data', displayName: '导出数据', description: '导出数据到外部文件', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { filePath: { type: 'string', description: '目标导出文件路径' } },
            required: ['filePath']
        },
        async execute() { executed = true; return { ok: true, result: {} }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_export_data', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, false);
    assert.ok(Array.isArray(res.needsInput) && res.needsInput.length > 0, '应返回 needsInput 缺失清单');
    assert.strictEqual(executed, false, '无法兜底的参数不应执行工具');
    const tc = turn.toolCalls.find(c => c.name === 'excel_export_data');
    assert.ok(tc, '应创建 needs-input toolCall 记录');
    assert.strictEqual(tc.status, 'needs-input');
    assert.match(String(res.error), /文件路径/, '反问措辞应包含参数描述');
});

test('skill 路径不受影响：不触发自动补全/反问，保持原有执行流程', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({ message: '创建一个新的工作表', kind: 'skill', context: {} });

    let executed = false;
    tr.register({
        name: 'excel_create_sheet', displayName: '新建工作表', description: '创建一个新工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '新工作表名称' } },
            required: ['name']
        },
        async execute(args) { executed = true; return { ok: true, result: { created: true, gotName: args && args.name } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_create_sheet', {}, { round: 1, ctx: { turn } });

    // skill 路径不拦截、不补全：走原始 invoke（无必填校验），工具正常执行
    assert.strictEqual(res.ok, true);
    assert.ok(executed, 'skill 路径应按原逻辑执行');
    assert.ok(!res.needsInput, 'skill 路径不应触发 needsInput');
});

// ── 真实失败场景（00:08 复现）："导出当前工作表为 CSV"缺 filePath 被 needsInput 拦截。
//    修复：inferExportFilePath 自动落盘到桌面/下载（文件名=工作表名_时间戳.csv），
//    无桥接时才回退"需用户补充"，绝不编造路径。
test('excel_export_csv: 缺 filePath 时自动推断桌面路径并执行（不再 needsInput 拦截）', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    // 模拟真实 WebView 桥接（app-model.js 会把 window.excelBridge 同步到 App.excelBridge）
    App.excelBridge = {
        GetDesktopPath: async () => 'C:\\Users\\test\\Desktop',
        GetDownloadsPath: async () => 'C:\\Users\\test\\Downloads'
    };

    const turn = ss.startTurn({
        message: '导出当前工作表为 CSV',
        kind: 'text',
        context: { activeSheet: { sheet: '销售报表' } }
    });

    let executedArgs = null;
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '将工作表导出为 CSV', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { filePath: { type: 'string', description: 'CSV 保存路径' } },
            required: ['filePath']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { success: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_export_csv', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true, '缺 filePath 应自动推断并执行成功，不应 needsInput');
    assert.ok(executedArgs, '工具应被真正执行');
    assert.ok(executedArgs.filePath, '应补全 filePath');
    assert.ok(executedArgs.filePath.startsWith('C:\\Users\\test\\Desktop\\销售报表_'),
        `应落在桌面且文件名以工作表名开头，实际: ${executedArgs.filePath}`);
    assert.ok(/\.csv$/.test(executedArgs.filePath), `应以 .csv 结尾，实际: ${executedArgs.filePath}`);
    const tc = turn.toolCalls.find(c => c.name === 'excel_export_csv');
    assert.strictEqual(tc.status, 'completed');
    console.log(`\n✅ CSV 导出自动落盘：${executedArgs.filePath}`);
});

test('excel_export_csv: LLM 输出 savePath 别名 → 归一化为 filePath（命中必填校验）', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    App.excelBridge = { GetDesktopPath: async () => 'C:\\Users\\test\\Desktop' };

    const turn = ss.startTurn({ message: '导出数据', kind: 'text', context: { activeSheet: { sheet: 'Sheet1' } } });

    let executedArgs = null;
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '将工作表导出为 CSV', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { filePath: { type: 'string', description: 'CSV 保存路径' } },
            required: ['filePath']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { success: true } }; }
    });

    // LLM 输出了 save_path 别名（后端只认 filePath）
    const res = await tr.invokeTracked(turn, 'excel_export_csv', { save_path: 'C:\\out\\data.csv' }, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true);
    assert.ok(executedArgs, '工具应被真正执行');
    assert.ok(executedArgs.filePath, 'save_path 应归一化为 filePath');
    assert.strictEqual(executedArgs.filePath, 'C:\\out\\data.csv', '应保留 LLM 显式给出的路径（不覆盖）');
    console.log(`\n✅ save_path 别名归一化：${executedArgs.filePath}`);
});

// ── 回归：`删除工作表"异常场景测试表"` 曾缺 name 被 needsInput 拦截，
//    又因 excel_delete_sheet 仅被 dashboard 域 guardedTools 收录（关键词=看板/仪表盘…）
//    被 agent-loop 的"无关调用"护栏静默降级为失败 —— 既不执行也不反问。
test('extractSheetNameFromMessage: 兼容 删除/重命名/复制 工作表"X" 与 名为"X"', () => {
    const { App } = loadChatV2();
    const util = App.ChatV2.toolRegistryUtil;
    const ex = (m) => util.extractSheetNameFromMessage({ userMessage: m });

    assert.strictEqual(ex('删除工作表"异常场景测试表"'), '异常场景测试表', '紧邻工作表后的引号名');
    assert.strictEqual(ex('删除“异常场景测试表”工作表'), '异常场景测试表', '引号名在工作表之前');
    assert.strictEqual(ex('重命名工作表“趋势”'), '趋势', '中文引号');
    assert.strictEqual(ex('帮我创建一个新的工作表名为"新工作表"'), '新工作表', '保持原有"名为"语义');
    assert.strictEqual(ex('筛选"订单金额"列'), null, '非工作表场景的引号列名不得被当成表名');
    assert.strictEqual(ex('帮我创建一个新的工作表'), null, '无显式表名时返回 null（交默认名生成器）');
});

test('excel_delete_sheet: 消息里显式表名应自动补全 name 并执行（不再 needsInput 拦截）', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '删除工作表"异常场景测试表"',
        kind: 'text',
        context: { activeSheet: { sheet: '异常场景测试表' } }
    });
    // excel_delete_sheet 在 excel_tools.json 标为 isDangerous：守卫在**所有**思考模式
    // （含 expert）都会弹确认框。本用例只验证"表名能否自动补全"，故模拟用户点
    // 「仍要继续执行」（确认流程另有用例覆盖）。
    turn.mode = 'expert';
    stubDangerConfirm(true);

    let executedArgs = null;
    tr.register({
        name: 'excel_delete_sheet', displayName: '删除工作表', description: '删除指定工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '要删除的工作表名称' } },
            required: ['name']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { deleted: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_delete_sheet', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true, '应自动补全 name 并执行成功');
    assert.ok(executedArgs, '工具应被真正执行');
    assert.strictEqual(executedArgs.name, '异常场景测试表', '应提取引号内的目标表名，而非兜底活动表名');
    assert.ok(!res.needsInput, '不应返回 needsInput');
});

test('excel_delete_sheet: 消息未显式给表名时回落到当前活动工作表名', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '把这个工作表删掉',
        kind: 'text',
        context: { activeSheet: { sheet: '临时表' } }
    });
    turn.mode = 'expert';
    stubDangerConfirm(true);

    let executedArgs = null;
    tr.register({
        name: 'excel_delete_sheet', displayName: '删除工作表', description: '删除指定工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '要删除的工作表名称' } },
            required: ['name']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: { deleted: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_delete_sheet', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true);
    assert.ok(executedArgs, '工具应被真正执行');
    assert.strictEqual(executedArgs.name, '临时表', '无显式表名时应回落到活动工作表');
});

test('excel_delete_sheet: standard 模式下应被"安全守卫"拦截并要求用户确认', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '删除工作表"异常场景测试表"',
        kind: 'text',
        context: { activeSheet: { sheet: '异常场景测试表' } }
    });
    turn.mode = 'standard';

    let executed = false;
    tr.register({
        name: 'excel_delete_sheet', displayName: '删除工作表', description: '删除指定工作表', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { name: { type: 'string', description: '要删除的工作表名称' } },
            required: ['name']
        },
        async execute() { executed = true; return { ok: true, result: {} }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_delete_sheet', {}, { round: 1, ctx: { turn } });

    assert.strictEqual(executed, false, '未经确认不得真正执行删除');
    assert.strictEqual(res.ok, false, '应被守卫拦截（ok=false）');
    assert.match(JSON.stringify(res), /安全守卫/, '拦截文案应带【安全守卫】标识，供 UI 弹确认框');
});

// ── 回归：模型把用户名当占位符，传 C:\Users\Yourusername\Desktop\output.csv。
//    该值非空 → 跳过 inferExportFilePath → 后端 CreateDirectory 真造出垃圾目录。
test('isPlaceholderPath: 识别编造的占位符路径，且不误伤正常路径', () => {
    const { App } = loadChatV2();
    const isP = App.ChatV2.toolRegistryUtil.isPlaceholderPath;

    // 正例：模型常见的占位符写法
    assert.strictEqual(isP('C:\\Users\\Yourusername\\Desktop\\output.csv'), true, 'Yourusername');
    assert.strictEqual(isP('C:\\Users\\用户名\\Desktop\\out.csv'), true, '中文占位符');
    assert.strictEqual(isP('C:\\Users\\<username>\\Desktop\\out.csv'), true, '尖括号模板');
    assert.strictEqual(isP('C:\\path\\to\\your\\file.csv'), true, 'path/to 模板');
    assert.strictEqual(isP('%USERNAME%\\Desktop\\out.csv'), true, '未展开环境变量');
    assert.strictEqual(isP('D:\\dummy\\x.csv'), true, 'dummy 段');
    assert.strictEqual(isP('D:\\someuser\\out.csv'), true, 'someuser 段');

    // 反例：必须放行，否则会丢失用户指定的合法路径/文件名
    assert.strictEqual(isP('C:\\Users\\Administrator\\Desktop\\销售报表.csv'), false, '真实用户路径');
    assert.strictEqual(isP('C:\\Users\\Administrator\\Desktop\\username_report.csv'), false, '文件名含 username 不得误判');
    assert.strictEqual(isP('D:\\工作\\导出\\明细.csv'), false, '普通中文路径');
    assert.strictEqual(isP(''), false, '空串');
    assert.strictEqual(isP(null), false, 'null');
});

test('excel_export_csv: 编造的 C:\\Users\\Yourusername 路径不得采信，应改用真实桌面', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;
    // 桥接 stub：提供"真实桌面"，模拟 VSTO 环境
    App.excelBridge = { async GetDesktopPath() { return 'C:\\Users\\Administrator\\Desktop'; } };

    const turn = ss.startTurn({
        message: '将当前工作表导出为 CSV 并保存到桌面',
        kind: 'text',
        context: { activeSheet: { sheet: '销售数据' } }
    });

    let executedArgs = null;
    tr.register({
        name: 'excel_export_csv', displayName: '导出CSV', description: '导出当前工作表为 CSV', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { filePath: { type: 'string', description: 'CSV 输出文件路径' } },
            required: ['filePath']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: {} }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_export_csv',
        { filePath: 'C:\\Users\\Yourusername\\Desktop\\output.csv' },
        { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true, '应正常执行（不再把假路径当有效参数而短路推断）');
    assert.ok(executedArgs, '工具应被执行');
    assert.ok(!/Yourusername/i.test(String(executedArgs.filePath)), `不得使用编造的用户目录：${executedArgs.filePath}`);
    assert.ok(String(executedArgs.filePath).startsWith('C:\\Users\\Administrator\\Desktop\\'),
        `应落到真实桌面，实际：${executedArgs.filePath}`);
});

test('ppt_export_pdf: 编造的 C:\\Users\\YourUsername 路径应交给后端兜底，不得拦截为"缺少必要参数"', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '将演示文稿导出为 PDF',
        kind: 'text',
        context: {}
    });

    let executedArgs = null;
    tr.register({
        name: 'ppt_export_pdf', displayName: '导出PDF', description: '将演示文稿导出为 PDF 文件', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: { filePath: { type: 'string', description: '导出 PDF 路径' } },
            required: ['filePath']
        },
        async execute(args) { executedArgs = args; return { ok: true, result: {} }; }
    });

    const res = await tr.invokeTracked(turn, 'ppt_export_pdf',
        { filePath: 'C:\\Users\\YourUsername\\Documents\\report_20261005.pdf' },
        { round: 1, ctx: { turn } });

    assert.strictEqual(res.ok, true, `占位符路径应交后端兜底而非拦截，实际错误：${res.error}`);
    assert.ok(executedArgs, '工具应被执行');
});

test('美化时底色只标到表头，点名表身或整表才铺满', () => {
    const { App } = loadChatV2();
    const confine = App.ChatV2.toolRegistryUtil.confineFormatFillToHeader;
    const beautify = { range: 'A1:I101', format: { bgColor: 'FFD700', bold: true } };
    const header = confine(beautify, { userMessage: '美化当前表格' });
    assert.strictEqual(beautify.format.bgColor, undefined, '整段参数里不应再带表身底色');
    assert.strictEqual(beautify.format.bold, true);
    assert.strictEqual(header.range, 'A1:I1');
    assert.strictEqual(header.format.bgColor, 'FFD700');

    const whole = { range: 'A1:I101', format: { bgColor: 'FFFF00' } };
    assert.strictEqual(confine(whole, { userMessage: '把整表底色涂成黄色' }), null);
    assert.strictEqual(whole.format.bgColor, 'FFFF00');

    const named = { range: 'A1:D10', format: { bgColor: 'FFFF00' } };
    assert.strictEqual(confine(named, { userMessage: 'A1:D10 底色改成黄色' }), null);

    const dash = { range: 'A1:I20', format: { bgColor: '1F4E79' } };
    assert.strictEqual(confine(dash, { userMessage: '做一个销售看板', context: { _forceDashboard: true } }), null);

    const setup = { range: 'A1:I101', format: { bgColor: '006600', fontName: 'Arial' } };
    const setupHeader = confine(setup, { userMessage: '把选中的表格设置一下格式' });
    assert.strictEqual(setup.format.bgColor, undefined);
    assert.strictEqual(setup.format.fontName, 'Arial');
    assert.strictEqual(setupHeader.range, 'A1:I1');
    assert.strictEqual(setupHeader.format.bgColor, '006600');

    const asText = { range: 'A1:I101', format: '{"bgColor":"006600","bold":true}' };
    const fromText = confine(asText, { userMessage: '美化当前表格' });
    assert.strictEqual(asText.format.bgColor, undefined);
    assert.strictEqual(asText.format.bold, true);
    assert.strictEqual(fromText.range, 'A1:I1');
    assert.strictEqual(fromText.format.bgColor, '006600');

    const oneRow = { range: 'A1:I1', format: { bgColor: 'FF0000' } };
    assert.strictEqual(confine(oneRow, { userMessage: '美化当前表格' }), null);
    assert.strictEqual(oneRow.format.bgColor, 'FF0000');

    const marked = { range: 'A2:I20', format: { bgColor: 'FFFF00', bgScope: 'all' } };
    assert.strictEqual(confine(marked, { userMessage: '美化当前表格' }), null);
    assert.strictEqual(marked.format.bgColor, 'FFFF00');

    const china = { range: 'A1:I21', sheet: '市场部', format: { bgColor: 'DE2910', bgScope: 'header', headerFontColor: 'FFFFFF' } };
    assert.strictEqual(confine(china, { userMessage: '美化工作簿中的所有工作表，统一为中国红' }), null);
    assert.strictEqual(china.format.bgColor, 'DE2910', '已声明只涂表头时，底色必须留在原调用里');
    assert.strictEqual(china.format.bgScope, 'header');
});

test('excel_format_range: 美化所有表时，光标单格不能填进 range，也不补 autoRange', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '美化工作簿中的所有工作表，统一为科技绿',
        kind: 'text',
        context: {
            activeSheet: { sheet: '市场部' },
            selection: { address: '$G$30', sheetName: '市场部' }
        }
    });

    let executed = false;
    tr.register({
        name: 'excel_format_range', displayName: '格式化单元格', description: '格式化指定范围', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: {
                sheet: { type: 'string' },
                range: { type: 'string' },
                format: { type: 'string' },
                autoRange: { type: 'boolean' }
            },
            required: ['range', 'format']
        },
        async execute() { executed = true; return { ok: true, result: { formatted: true } }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_format_range', {
        sheet: '市场部',
        format: { headerFontColor: 'FFFFFF' }
    }, { round: 1, ctx: { turn } });

    assert.strictEqual(executed, false, '数据区还没解析出来时不能拿光标格执行');
    assert.ok(res.needsInput, '缺的是这张表自己的数据区，不是活动表 UsedRange');
    const tc = turn.toolCalls.find(c => c.name === 'excel_format_range');
    assert.notStrictEqual(tc && tc.args && tc.args.range, 'G30');
    assert.notStrictEqual(tc && tc.args && tc.args.range, '$G$30');
    assert.notStrictEqual(tc && tc.args && tc.args.autoRange, true);
});

test('excel_format_range: 圈定的一行、嘴里说的地址，按原样执行', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    function register() {
        tr.register({
            name: 'excel_format_range', displayName: '格式化单元格', description: '格式化指定范围', source: 'mcp',
            inputSchema: {
                type: 'object',
                properties: {
                    sheet: { type: 'string' },
                    range: { type: 'string' },
                    format: { type: 'string' }
                },
                required: ['range', 'format']
            },
            async execute(args) { return { ok: true, result: args }; }
        });
    }

    const rowTurn = ss.startTurn({
        message: '美化选中的这一行',
        kind: 'text',
        context: { selection: { address: 'A5:I5', sheetName: '市场部' } }
    });
    register();
    const rowRes = await tr.invokeTracked(rowTurn, 'excel_format_range', {
        format: { bold: true }
    }, { round: 1, ctx: { turn: rowTurn } });
    assert.strictEqual(rowRes.ok, true);
    const rowCall = rowTurn.toolCalls.find(c => c.name === 'excel_format_range');
    assert.strictEqual(rowCall.args.range, 'A5:I5');
    assert.strictEqual(rowCall.args.sheet, '市场部');

    const saidTurn = ss.startTurn({
        message: '把所有工作表的 A1:D10 统一成中国红',
        kind: 'text',
        context: { activeSheet: { sheet: '市场部' }, selection: { address: 'G30', sheetName: '市场部' } }
    });
    register();
    const saidRes = await tr.invokeTracked(saidTurn, 'excel_format_range', {
        sheet: '销售一部',
        format: { bgColor: 'DE2910', bgScope: 'header' }
    }, { round: 1, ctx: { turn: saidTurn } });
    assert.strictEqual(saidRes.ok, true);
    const saidCall = saidTurn.toolCalls.find(c => c.name === 'excel_format_range');
    assert.strictEqual(saidCall.args.range, 'A1:D10');
    assert.strictEqual(saidCall.args.sheet, '销售一部');
});

test('excel_format_range: 表名和范围都没有时仍然反问，不静默格式化整本工作簿', async () => {
    const { App } = loadChatV2();
    const tr = App.ChatV2.toolRegistry;
    const ss = App.ChatV2.sessionStore;

    const turn = ss.startTurn({
        message: '美化工作簿中的所有工作表，统一为科技绿',
        kind: 'text'
    });

    let executed = false;
    tr.register({
        name: 'excel_format_range', displayName: '格式化单元格', description: '格式化指定范围', source: 'mcp',
        inputSchema: {
            type: 'object',
            properties: {
                sheet: { type: 'string' },
                range: { type: 'string' },
                format: { type: 'string' }
            },
            required: ['range', 'format']
        },
        async execute() { executed = true; return { ok: true, result: {} }; }
    });

    const res = await tr.invokeTracked(turn, 'excel_format_range', {
        format: { headerFontColor: 'FFFFFF' }
    }, { round: 1, ctx: { turn } });

    assert.strictEqual(executed, false, '没有表名也没有范围时不应执行');
    assert.ok(Array.isArray(res.needsInput) && res.needsInput.some(m => m.parameter === 'range'));
});

if (require.main === module) { setTimeout(() => process.exit(0), 4000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出