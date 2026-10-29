/**
 * ============================================================
 * Tests/routing-constants.test.js — 路由常量收敛回归测试
 * ============================================================
 * 验证：
 *   1. config/routing-constants.js 是路由常量单一事实源（可被 Node 直接 require）。
 *   2. detectPythonAnalysisIntent 正确识别数据分析意图（标准集/保守集）。
 *   3. chat-orchestrator.js 不再内联重复的 Python 路由正则（源码 + 部署）。
 *
 * 运行：node Tests/routing-constants.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const routingConstants = require(path.join(ROOT, 'UI', 'js', 'chat-v2', 'config', 'routing-constants.js'));

function readFile(p) {
    const buf = fs.readFileSync(p);
    if (buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) return buf.toString('utf8', 3);
    return buf.toString('utf8');
}

test('routing-constants 导出完整', () => {
    assert.ok(routingConstants.COMMON_2CHAR_WORDS instanceof Set, 'COMMON_2CHAR_WORDS 应为 Set');
    assert.strictEqual(typeof routingConstants.detectLanguageDirective, 'function');
    assert.strictEqual(typeof routingConstants.detectCodeGenerationIntent, 'function');
    assert.strictEqual(typeof routingConstants.detectPythonAnalysisIntent, 'function');
    assert.strictEqual(typeof routingConstants.stripQuotedSegments, 'function');
    assert.strictEqual(typeof routingConstants.isSheetManagementIntent, 'function');
    assert.ok(routingConstants.PYTHON_ANALYSIS_INTENT, 'PYTHON_ANALYSIS_INTENT 缺失');
    assert.ok(routingConstants.PYTHON_ANALYSIS_INTENT_CONSERVATIVE, 'PYTHON_ANALYSIS_INTENT_CONSERVATIVE 缺失');
});

test('detectPythonAnalysisIntent 标准集识别分析意图', () => {
    const cases = [
        ['分析销售趋势', 'trend'],
        ['找出异常数据', 'anomaly'],
        ['看下相关性', 'correlation'],
        ['做描述性统计', 'descriptive'],
        ['看数据分布', 'distribution'],
        ['趋势和异常一起分析', 'multi']
    ];
    for (const [msg, key] of cases) {
        const r = routingConstants.detectPythonAnalysisIntent(msg);
        assert.strictEqual(r[key], true, `${msg} 应命中 ${key}`);
        assert.strictEqual(r.any, true, `${msg} 应 any=true`);
    }
    // 纯分组请求：groupAgg 单独命中不算分析意图（与 orchestrator 决策一致，仅纯分组不走 python）
    const groupOnly = routingConstants.detectPythonAnalysisIntent('按部门分组');
    assert.strictEqual(groupOnly.groupAgg, true, '按部门分组 应命中 groupAgg');
    assert.strictEqual(groupOnly.any, false, '纯分组请求不应 any=true');
});

test('detectPythonAnalysisIntent 保守集排除宽泛词（最终兜底防误路由）', () => {
    // 保守集不含"变化"，避免"颜色变化/格式变化"等无 Skill/无选区时误路由
    const r = routingConstants.detectPythonAnalysisIntent('把颜色变化一下', { conservative: true });
    assert.strictEqual(r.trend, false, '保守集不应因"变化"命中 trend');
    assert.strictEqual(r.any, false, '保守集下纯"变化"不应路由 python');

    // 标准集（matchedSkill 分支）仍保留"变化"（与旧行为一致）
    const r2 = routingConstants.detectPythonAnalysisIntent('分析销售变化', {});
    assert.strictEqual(r2.trend, true, '标准集应识别"变化"为趋势信号');
});

test('工作表管理指令优先：引号内表名不得劫持为 Python 分析', () => {
    // 回归：`删除工作表"异常场景测试表"` 曾因表名含"异常"被强制路由到 python（异常值检测）
    const msg = '删除工作表"异常场景测试表"';
    assert.strictEqual(routingConstants.isSheetManagementIntent(msg), true, '应识别为工作表管理指令');
    assert.strictEqual(routingConstants.detectPythonAnalysisIntent(msg).any, false, '不得判定为分析意图');

    // 引号（中英文）内的分析词均应被剥离
    assert.ok(!/异常/.test(routingConstants.stripQuotedSegments(msg)), 'stripQuotedSegments 应剥离引号内"异常"');
    const msgCn = '删除工作表“异常场景测试表”';
    assert.strictEqual(routingConstants.detectPythonAnalysisIntent(msgCn).any, false, '中文引号表名同样不得判定为分析意图');

    // 其余工作表管理动词
    for (const m of ['新建工作表"汇总"', '创建 sheet 明细', '把工作表"异常表"重命名', '复制工作表“趋势”', '删除"异常"工作表']) {
        assert.strictEqual(routingConstants.isSheetManagementIntent(m), true, `${m} 应识别为工作表管理指令`);
    }
    // 引号内出现的"删除"等词不得让非工作表管理消息误判（表名污染反向护栏）
    for (const m of ['筛选"删除记录"列', '统计明细表', '建议看一下工作表结构']) {
        assert.strictEqual(routingConstants.isSheetManagementIntent(m), false, `${m} 不应识别为工作表管理指令`);
    }

    // 引号外的分析词仍应正常命中（不能因护栏误伤真实分析请求）
    const ok = routingConstants.detectPythonAnalysisIntent('对"销售额"列做异常检测');
    assert.strictEqual(ok.anomaly, true, '引号外"异常检测"仍应命中 anomaly');
    assert.strictEqual(ok.any, true, '真实分析请求仍应 any=true');

    // 表名在引号内的分析词不构成分析意图（无引号外表名时同样安全）
    assert.strictEqual(routingConstants.detectPythonAnalysisIntent('筛选"异常值"列').any, false, '引号内"异常值"不应触发分析');
    assert.strictEqual(routingConstants.isSheetManagementIntent('分析销售趋势'), false, '纯分析请求不是工作表管理');
});

test('detectLanguageDirective / detectCodeGenerationIntent', () => {
    assert.strictEqual(routingConstants.detectLanguageDirective('用python写脚本'), 'python');
    assert.strictEqual(routingConstants.detectLanguageDirective('写个VBA宏'), 'vba');
    assert.strictEqual(routingConstants.detectLanguageDirective('你好'), null);
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('帮我写一段代码'), true);
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('清洗数据'), false);
    // 【防误路由】"解析脚本/解析器"是 OCR 解析器产物域，不是在要代码
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('帮我做一个发票解析脚本'), false,
        '"解析脚本"不应判为代码生成意图（否则被 GenerateVBA 抢走）');
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('为火车票创建一个解析器'), false);
    // 明确点名代码载体/语言时，仍然算代码意图
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('写一个解析脚本的VBA代码'), true);
    assert.strictEqual(routingConstants.detectCodeGenerationIntent('帮我写个脚本处理数据'), true);
});

test('chat-orchestrator.js 不再内联重复 Python 路由正则（源码+部署）', () => {
    const files = [
        path.join(ROOT, 'UI', 'js', 'chat-v2', 'chat-orchestrator.js'),
        path.join(ROOT, 'bin', 'Debug', 'UI', 'js', 'chat-v2', 'chat-orchestrator.js')
    ];
    for (const f of files) {
        if (!fs.existsSync(f)) continue;
        const src = readFile(f);
        // 旧的重复内联正则变量已移除，统一走 routing().detectPythonAnalysisIntent
        assert.ok(!src.includes('forceHasTrend'), `${f} 仍含旧内联正则 forceHasTrend`);
        assert.ok(!src.includes('finalHasTrend'), `${f} 仍含旧内联正则 finalHasTrend`);
        assert.ok(!src.includes('const COMMON_2CHAR_WORDS = new Set'), `${f} 仍内联 COMMON_2CHAR_WORDS`);
        assert.ok(src.includes('detectPythonAnalysisIntent'), `${f} 应调用统一检测函数`);
    }
});
