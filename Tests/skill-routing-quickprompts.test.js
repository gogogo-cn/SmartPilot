/**
 * ============================================================
 * Tests/skill-routing-quickprompts.test.js — 继续操作快捷提示词路由回归测试
 * ============================================================
 * 验证：
 *   1. chat-orchestrator.js 的 strongMatch 逻辑包含"多候选明显领先（分差>=3）也路由"条件，
 *      修复"去重 · 仅高亮"快捷提示词因 SmartDataAnalyzer 的"标记"成为第二候选而掉进 text-pipeline 的问题。
 *   2. 复刻 shouldRouteToSkill 路由逻辑，验证 SmartDataCleaner 全部 5 条快捷提示词均正确路由。
 *
 * 背景：单独发"去重/高亮"没问题，但通过"继续操作"按钮连续发送时，
 *   "对当前数据继续清洗，去重并只把重复行黄色高亮标记（不删除、不新增列）"无法路由到 SmartDataCleaner，
 *   导致掉进 text-pipeline（LLM 光说不做，报"实际没有调用任何工具"）。
 *
 * 运行：node Tests/skill-routing-quickprompts.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const ROOT = path.join(__dirname, '..');
const ORCH_SRC = path.join(ROOT, 'UI', 'js', 'chat-v2', 'chat-orchestrator.js');
const ORCH_DEP = path.join(ROOT, 'bin', 'Debug', 'UI', 'js', 'chat-v2', 'chat-orchestrator.js');
const SKILL_FILE = path.join(ROOT, 'Skills', 'SmartDataCleaner.md');
// 路由常量单一事实源（与 chat-orchestrator 共用，避免测试内联副本漂移）
const routingConstants = require(path.join(ROOT, 'UI', 'js', 'chat-v2', 'config', 'routing-constants.js'));
const COMMON_2CHAR_WORDS = routingConstants.COMMON_2CHAR_WORDS;

function readFile(p) {
    const buf = fs.readFileSync(p);
    if (buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) return buf.toString('utf8', 3);
    if (buf[0] === 0xFF && buf[1] === 0xFE) return buf.toString('utf16le', 2);
    if (buf[0] === 0xFE && buf[1] === 0xFF) return buf.toString('utf16be', 2);
    return buf.toString('utf8');
}

test('chat-orchestrator.js 包含多候选明显领先路由条件（源码+部署）', () => {
    for (const f of [ORCH_SRC, ORCH_DEP]) {
        const src = readFile(f);
        // 【常量收敛适配】领先分差已提取为命名常量 ROUTE_T.ROUTE_LEAD_SCORE（值 3），
        // 语义不变：单候选直接路由，或多候选中最佳明显领先（分差 >= 领先阈值）也路由。
        assert.ok(src.includes('candidates.length === 1 || best.score - candidates[1].score >= ROUTE_T.ROUTE_LEAD_SCORE'),
            `${f} 缺少多候选明显领先路由条件`);
    }
});

test('SmartDataCleaner 快捷提示词全部正确路由', () => {
    // 从 skill 文件提取 keywords
    const skillMd = readFile(SKILL_FILE);
    const kwMatch = skillMd.match(/^keywords:[\s\S]*?^[a-zA-Z]+:/m);
    assert.ok(kwMatch, 'SmartDataCleaner.md 缺少 keywords 段');
    const cleanerKeywords = kwMatch[0]
        .split('\n')
        .map(l => l.trim().replace(/^-\s*/, ''))
        .filter(l => l && !/^keywords:$/.test(l) && !/^[a-zA-Z]+:/.test(l));

    // 与 SmartDataCleaner 竞争意图的候选 skill 关键词（与真实文件一致）
    const skills = [
        { name: 'SmartDataCleaner', keywords: cleanerKeywords },
        { name: 'SmartDataAnalyzer', keywords: ['异常值','异常检测','异常行','智能分析','问题发现','数据洞察','发现','标记','问题行'] },
        { name: 'DataMerger', keywords: ['合并','数据合并','多文件合并','合并Sheet','合并工作表','合并工作簿','合并表格','合并数据','去重合并','合并去重','汇总合并','合并多表','合并到一个表','合并到一张表','合并所有Sheet','把多个文件合并到一起','把所有Sheet合成一个','横向汇总','两张表放一起'] },
        { name: 'DataSplitter', keywords: ['拆分','拆分Sheet','拆分工作表','按条件拆分','拆分成','拆分数据','分割','按列拆分','按行拆分'] },
        { name: 'DataDiffAnalyzer', keywords: ['对比','差异','比较','对账','核对','版本对比'] },
        { name: 'DescriptiveStatsReporter', keywords: ['统计','描述统计','均值','中位数','标准差','频率分布','统计报告'] },
        { name: 'SmartTableBeautifier', keywords: ['美化','美化表格','美化数据','表格美化','格式化表格','美化样式','排版','样式美化'] },
        { name: 'TableJoiner', keywords: ['关联','连接','join','两表关联','表关联','数据关联','按列关联','匹配关联'] }
    ];

    const COMMON_2CHAR_WORDS = routingConstants.COMMON_2CHAR_WORDS;

    function shouldRouteToSkill(message) {
        const lower = message.toLowerCase();
        const candidates = [];
        for (const skill of skills) {
            let score = 0, hasLongKwMatch = false, hasMidKwMatch = false, hasShortKwMatch = false, hasStrong2Char = false, lastStrongPos = -1;
            for (const kw of skill.keywords) {
                if (typeof kw !== 'string') continue;
                const kwLower = kw.toLowerCase();
                const kwLen = kwLower.length;
                if (kwLen < 2) continue;
                if (lower.includes(kwLower)) {
                    if (kwLen >= 5) { score += 8; hasLongKwMatch = true; }
                    else if (kwLen >= 4) { score += 4; hasMidKwMatch = true; }
                    else if (kwLen >= 3) { score += 3; hasShortKwMatch = true; }
                    else { score += 1; }
                }
            }
            if (score === 0) continue;
            if (!hasLongKwMatch && !hasMidKwMatch && !hasShortKwMatch) {
                let kwSumLen = 0, hitCount = 0;
                for (const kw of skill.keywords) {
                    if (typeof kw !== 'string') continue;
                    const kwLower = kw.toLowerCase();
                    if (kwLower.length >= 2 && lower.includes(kwLower)) {
                        kwSumLen += kwLower.length;
                        hitCount++;
                        if (kwLower.length === 2 && !COMMON_2CHAR_WORDS.has(kwLower)) {
                            hasStrong2Char = true;
                            const idx = lower.lastIndexOf(kwLower);
                            if (idx > lastStrongPos) lastStrongPos = idx;
                        }
                    }
                }
                if (!hasStrong2Char && (kwSumLen < 4 || hitCount === 0)) continue;
            }
            candidates.push({ skill, score, _hasLongMatch: hasLongKwMatch, _hasMidKwMatch: hasMidKwMatch, _hasShortKwMatch: hasShortKwMatch, _hasStrong2Char: hasStrong2Char, _lastStrongPos: lastStrongPos });
        }
        if (candidates.length === 0) return null;
        let maxStrongPos = -1;
        for (const c of candidates) if (c._lastStrongPos > maxStrongPos) maxStrongPos = c._lastStrongPos;
        if (maxStrongPos >= 0) for (const c of candidates) if (c._lastStrongPos === maxStrongPos) c.score += 2;
        candidates.sort((a, b) => b.score - a.score);
        const best = candidates[0];
        const strongMatch = (best.score >= 8)
            || (best.score >= 5 && candidates[0]._hasLongMatch)
            || (best.score >= 4 && candidates[0]._hasMidKwMatch)
            || (best.score >= 3 && candidates[0]._hasShortKwMatch && (candidates.length === 1 || best.score - candidates[1].score >= 3))
            || (candidates.length === 1 && candidates[0]._hasStrong2Char)
            || (candidates.length > 0 && candidates[0]._hasStrong2Char && candidates[0].score >= 3);
        return strongMatch ? best.skill.name : null;
    }

    const quickPrompts = [
        '对当前数据继续清洗，去重并只把重复行黄色高亮标记（不删除、不新增列）',
        '对当前数据继续清洗，删除重复行（真正删除，仅保留首条唯一数据）',
        '对当前数据继续清洗，填充空值',
        '对当前数据继续清洗，去除文本前后空格',
        '对当前数据继续清洗，统一日期格式'
    ];

    for (const p of quickPrompts) {
        const matched = shouldRouteToSkill(p);
        assert.strictEqual(matched, 'SmartDataCleaner', `快捷提示词未路由到 SmartDataCleaner: ${p} → ${matched}`);
    }
});

test('SmartDataCleaner.md 快捷提示词声明完整', () => {
    const skillMd = readFile(SKILL_FILE);
    const qpMatch = skillMd.match(/^quickPrompts:[\s\S]*?^[a-zA-Z]+:/m);
    assert.ok(qpMatch, 'SmartDataCleaner.md 缺少 quickPrompts 段');
    const prompts = qpMatch[0].match(/prompt:\s*(.+)$/gm) || [];
    assert.ok(prompts.length >= 5, `SmartDataCleaner 快捷提示词数量不足: ${prompts.length}`);
});
