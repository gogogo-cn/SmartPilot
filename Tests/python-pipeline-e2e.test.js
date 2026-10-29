/**
 * ============================================================
 * Tests/python-pipeline-e2e.test.js — python-pipeline 端到端集成测试
 * ============================================================
 * 目标：在 Node 中完整执行 python-pipeline 的自定义脚本路径，
 *       用 stub LLM 模拟"LLM 生成含伪造 input_data 的脚本"，验证：
 *   1. ensureSelectionLoaded 预读真实 rows 到 context
 *   2. buildExecArgs 移除伪造 input_data 赋值
 *   3. invokeTracked 收到修复后脚本 + referenceRegions 真实数据
 * 运行：node Tests/python-pipeline-e2e.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const path = require('path');
const fs = require('fs');
const { loadChatV2, DEFAULT_ORDER, BASE, stubDangerConfirm } = require('./helpers/harness');

// 与真实 index.html 一致的加载顺序，但截断到 chat-orchestrator（不加载 chat-entry/index，避免入口副作用）
const ORDER = DEFAULT_ORDER.slice(0, DEFAULT_ORDER.indexOf('ui/chat-renderer.js'));

test('python-pipeline: 端到端 - 用户选真实数据 + LLM 生成伪造脚本 → 真实数据注入执行', async () => {
    const App = loadChatV2(ORDER).App;
    // excel_run_python_script 是高危工具：所有思考模式（含下面的 expert）都会被【安全守卫】
    // 拦下并弹确认框。本用例验证的是"脚本注入真实数据"，故模拟用户点「仍要继续执行」。
    stubDangerConfirm(true);

    // 1. 注册 excel_run_python_script 工具（stub 执行：捕获收到的 script + referenceRegions）
    const tr = App.ChatV2.toolRegistry;
    let capturedExecArgs = null;
    tr.register({
        name: 'excel_run_python_script', displayName: '执行 Python', description: '执行 Python 脚本',
        inputSchema: { type: 'object', properties: { script: { type: 'string' } }, required: ['script'] },
        source: 'mcp',
        async execute(args) { capturedExecArgs = args; return { ok: true, result: { output: '分析完成' } }; }
    });

    // 2. stub excel_read_range（预读真实数据）
    const REAL_DATA = [
        { '日期': '2023-01-01', '部门': '销售部', '地区': '北京', '销售额': 1000, '成本': 600, '利润': 400 },
        { '日期': '2023-01-02', '部门': '销售部', '地区': '上海', '销售额': 1500, '成本': 900, '利润': 600 },
        { '日期': '2023-01-03', '部门': '市场部', '地区': '北京', '销售额': 800, '成本': 500, '利润': 300 }
    ];
    tr.register({
        name: 'excel_read_range', displayName: '读取数据', description: '读取 Excel 数据',
        inputSchema: { type: 'object', properties: {} }, source: 'mcp',
        async execute() {
            return { ok: true, result: { success: true, headers: Object.keys(REAL_DATA[0]), rows: REAL_DATA } };
        }
    });

    // 3. stub llmClient：自定义脚本模式，生成含伪造 input_data 的脚本
    const FAKE_SCRIPT = [
        'import pandas as pd',
        '',
        '# 假设 input_data 已经定义',
        'input_data = [',
        '    {"日期": "2023-01-01", "部门": "销售部", "地区": "北京", "销售额": 1000, "成本": 600, "利润": 400, "客户数": 50},',
        '    # ... 其他数据',
        ']',
        '',
        'df = pd.DataFrame(input_data)',
        'print(df.groupby("地区").agg({"销售额": "sum"}))',
        'output = {"rows": len(df)}',
        ''
    ].join('\n');

    let llmCallCount = 0;
    App.ChatV2.llmClient = {
        stream({ messages }) {
            return (async function* () {
                llmCallCount++;
                const hasCustomPrompt = messages.some(m => m.role === 'system' && String(m.content).includes('input_data 变量注入'));
                if (hasCustomPrompt) {
                    yield { type: 'token', delta: '```python\n' + FAKE_SCRIPT + '\n```' };
                } else {
                    yield { type: 'token', delta: '{}' };
                }
                yield { type: 'done', finishReason: 'stop' };
            })();
        }
    };

    // 4. 模拟 turn：context.selection 只有地址无 rows（模拟 context-engine 只给 headers）
    const sessionStore = App.ChatV2.sessionStore;
    const turn = sessionStore.startTurn({
        message: '不要用模板，自己写 python 脚本，按地区汇总销售额',
        kind: 'text',
        context: { selection: { address: 'A1:F10', sheetName: 'Sheet1' } }
    });
    // 守卫读 turn.mode：expert 放行 excel_run_python_script（高危工具）
    turn.mode = 'expert';

    // 5. 运行 python-pipeline
    const pipeline = App.ChatV2.pipelines.python;
    const gen = pipeline.execute(turn, new AbortController().signal);
    const events = [];
    for await (const evt of gen) {
        events.push(evt.type);
        if (evt.type === 'status') console.log('  [status]', evt.text);
    }

    console.log('  [llm-calls]', llmCallCount);

    // 6. 断言：capturedExecArgs 收到修复后脚本
    assert.ok(capturedExecArgs, 'execute 应被调用（捕获 execArgs）');
    const finalScript = capturedExecArgs.script;
    console.log('\n=== 最终执行的脚本 ===');
    console.log(finalScript);

    // 6a. 伪造 input_data 赋值被移除
    assert.ok(!/input_data\s*=\s*\[/.test(finalScript), '伪造 input_data = [...] 应被移除');
    // 6b. 使用逻辑保留
    assert.ok(finalScript.includes('pd.DataFrame(input_data)'), 'pd.DataFrame(input_data) 使用保留');
    // 6c. referenceRegions 携带真实数据
    assert.ok(capturedExecArgs.referenceRegions, '应携带 referenceRegions');
    const realRows = capturedExecArgs.referenceRegions[0]?.rows || [];
    assert.strictEqual(realRows.length, 3, 'referenceRegions 应含 3 行真实数据');
    assert.strictEqual(realRows[0]['地区'], '北京', '第一行数据应为真实地区');
    console.log('\n✅ referenceRegions 携带真实数据:', JSON.stringify(realRows[0]));

    // 7. 事件流完整
    assert.ok(events.includes('round_start'), '应有 round_start');
    console.log('✅ pipeline 事件流:', events.join(', '));
    console.log('\nALL E2E PASS');
});

test('python-pipeline: 模板模式 prompt 包含真实列名注入与禁造约束', async () => {
    const src = fs.readFileSync(path.join(BASE, 'pipelines/python-pipeline.js'), 'utf8');
    assert.ok(src.includes('【可用数据列】'), '模板选择 prompt 含数据列注入');
    assert.ok(src.includes('严禁编造不存在的列名'), '模板选择 prompt 禁止编造列名');
    console.log('✅ 模板模式 prompt 检查通过');
});

if (require.main === module) { setTimeout(() => process.exit(0), 5000).unref(); } // 仅独立运行时强制退出；聚合运行（run-tests.js require）不注册，避免进程提前退出
