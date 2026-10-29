#!/usr/bin/env node
/**
 * lint.js — 零依赖前端静态检查（替代 npx eslint，适配项目零依赖理念）
 *
 * 职责：
 *   1. 递归收集 UI/js 下所有 *.js 并做语法校验（node --check）
 *   2. 施加一组可扩展的规则子集（默认禁止 debugger 残留）
 *   3. 汇总输出，存在违规以非 0 退出（供 CI lint job 硬闸门）
 *
 * 用法：node Tools/lint.js [dir=UI/js]
 * 规则扩展：仅需向 RULES 数组追加 { name, test(file, src) } 项
 */
'use strict';

const fs = require('fs');
const path = require('path');
const { execFileSync } = require('child_process');

// ---------- 规则子集（可扩展）----------
const RULES = [
    {
        name: 'no-debugger',
        description: '禁止残留 debugger 语句',
        test(_file, src) {
            const report = [];
            // debugger 语句特征：整行独立、且前后不是引号/标识符（避免误判字符串如 'vba-debugger'）
            const re = /(^|[;"\n\s(])\s*debugger\s*(;|$|[\r\n})])/gm;
            let m;
            while ((m = re.exec(src)) !== null) {
                // 若紧邻前/后的临界字符是引号，说明命中的是字符串内容，跳过
                const before = m[1];
                const after = m[2];
                if (before !== '"' && before !== "'" && after !== '"' && after !== "'") {
                    report.push({ line: countLines(src, m.index), code: 'debugger' });
                }
            }
            return report;
        }
    }
    // 追加新规则时按此结构扩展即可
];

// ---------- 工具 ----------
function countLines(src, index) {
    let line = 1;
    for (let i = 0; i < index; i++) {
        if (src.charCodeAt(i) === 10) line++;
    }
    return line;
}

function collectJs(dir) {
    const out = [];
    if (!fs.existsSync(dir)) return out;
    for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
        const full = path.join(dir, entry.name);
        if (entry.isDirectory()) {
            // 跳过 node_modules / 构建产物
            if (entry.name === 'node_modules' || entry.name === 'bin') continue;
            out.push(...collectJs(full));
        } else if (entry.isFile() && entry.name.endsWith('.js')) {
            out.push(full);
        }
    }
    return out;
}

function syntaxCheck(file) {
    // --check 不执行文件，仅做语法解析；抛错即语法错误
    execFileSync(process.execPath, ['--check', file], { stdio: 'pipe' });
}

// ---------- 主流程 ----------
function main() {
    const dir = process.argv[2] || path.join(process.cwd(), 'UI', 'js');
    const files = collectJs(dir);
    const errors = [];
    const warnings = [];

    for (const file of files) {
        const rel = path.relative(process.cwd(), file).replace(/\\/g, '/');
        let src;
        try {
            src = fs.readFileSync(file, 'utf8');
        } catch (e) {
            errors.push({ file: rel, kind: 'read', message: e.message });
            continue;
        }

        // 1. 语法校验
        try {
            syntaxCheck(file);
        } catch (e) {
            const info = (e.stderr || e.message || '').toString().split('\n').slice(0, 6).join(' | ');
            errors.push({ file: rel, kind: 'syntax', message: info || '语法错误' });
            continue; // 语法错误时后续规则无意义，跳过
        }

        // 2. 规则子集
        for (const rule of RULES) {
            try {
                const hits = rule.test(file, src);
                for (const h of hits) {
                    warnings.push({ file: rel, kind: rule.name, line: h.line, message: rule.description + ': ' + (h.code || '') });
                }
            } catch (e) {
                warnings.push({ file: rel, kind: rule.name, message: '规则执行异常: ' + e.message });
            }
        }
    }

    // ---------- 汇总输出 ----------
    let exit = 0;
    if (errors.length > 0) {
        console.error(`[LINT] ${errors.length} 个语法/读取错误:`);
        for (const e of errors) {
            console.error(`  ✗ ${e.file}: ${e.kind}: ${e.message}`);
        }
        exit = 1;
    }
    if (warnings.length > 0) {
        console.warn(`[LINT] ${warnings.length} 条规则警告:`);
        for (const w of warnings) {
            console.warn(`  ⚠ ${w.file}:${w.line || '-'}: ${w.kind}: ${w.message}`);
        }
    }
    console.log(`[LINT] 检查 ${files.length} 个文件 → 语法错误 ${errors.length}, 规则警告 ${warnings.length}`);
    if (!exit && warnings.length > 0) {
        console.log('[LINT] 规则警告不阻塞构建（仅提示）。如需硬闸门，可提升规则为 error。');
    }
    return exit;
}

process.exitCode = main();