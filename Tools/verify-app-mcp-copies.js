/**
 * ============================================================
 * verify-app-mcp-copies.js — app-mcp.js 副本一致性校验
 * ============================================================
 * 背景：app-mcp.js 的单一事实源是 UI/js/app-mcp.js（见该文件头注释）。
 *   运行时副本由 Tools/sync-ui-to-debug.js 在构建时复制到各宿主的
 *   <Host>/bin/Debug/UI/js/app-mcp.js（及 Excel 的 bin/Debug/UI/js/app-mcp.js）。
 *   若开发者改了源却忘记跑同步，或误改了 bin 副本，线上就会用旧/分叉代码，
 *   且极难排查（界面行为"偶尔不对"但源里看不出问题）。
 *
 * 本脚本扫描仓库内所有 bin 目录下的 UI/js/app-mcp.js 副本，与单一事实源逐字节比对，
 * 发现漂移即报错并以退出码 1 结束，建议在 CI / 发布前运行。
 *
 * 用法：
 *   node Tools/verify-app-mcp-copies.js
 *   node Tools/verify-app-mcp-copies.js --json     # 机器可读输出
 * ============================================================
 */

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const SRC = path.join(ROOT, 'UI', 'js', 'app-mcp.js');
const args = process.argv.slice(2);
const asJson = args.includes('--json');

function fail(msg) {
    if (asJson) console.log(JSON.stringify({ ok: false, error: msg }, null, 2));
    else console.error('[ERR] ' + msg);
    process.exit(1);
}

if (!fs.existsSync(SRC)) fail('单一事实源缺失: ' + SRC);

const srcBuf = fs.readFileSync(SRC);

// 递归扫描：进入 bin 目录后收集其下所有 app-mcp.js；其余目录继续下钻
// （跳过 node_modules / .git 等无关目录，避免误扫与性能浪费）。
function findCopies(dir, out) {
    let entries;
    try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch (_) {
        return;
    }
    for (const e of entries) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) {
            if (e.name === 'bin') collectUnder(full, out);
            else if (e.name === 'node_modules' || e.name === '.git') continue;
            else findCopies(full, out);
        } else if (e.name === 'app-mcp.js') {
            out.push(full);
        }
    }
}

// 收集某个 bin 目录子树下的所有 app-mcp.js
function collectUnder(dir, out) {
    let entries;
    try {
        entries = fs.readdirSync(dir, { withFileTypes: true });
    } catch (_) {
        return;
    }
    for (const e of entries) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) collectUnder(full, out);
        else if (e.name === 'app-mcp.js') out.push(full);
    }
}

const copies = [];
findCopies(ROOT, copies);

const results = [];
let mismatched = 0;
for (const c of copies) {
    let ok = false;
    let sizeMatch = false;
    let err = null;
    try {
        const buf = fs.readFileSync(c);
        sizeMatch = buf.length === srcBuf.length;
        ok = sizeMatch && buf.equals(srcBuf);
    } catch (e) {
        err = e.message;
    }
    if (!ok) mismatched++;
    results.push({ copy: path.relative(ROOT, c), match: ok, sizeMatch, err });
}

if (asJson) {
    console.log(JSON.stringify({
        ok: mismatched === 0,
        source: path.relative(ROOT, SRC),
        copies: results,
        mismatched
    }, null, 2));
} else {
    console.log('单一事实源: ' + path.relative(ROOT, SRC) + ' (' + srcBuf.length + ' bytes)');
    console.log('发现副本: ' + copies.length + ' 个');
    for (const r of results) {
        if (r.match) console.log('  [OK]   ' + r.copy);
        else console.error('  [DIFF] ' + r.copy +
            (r.err ? ' (读错误: ' + r.err + ')' : (r.sizeMatch ? ' (内容不一致)' : ' (大小不一致)')));
    }
    if (mismatched === 0) {
        if (copies.length === 0) console.log('  [i] 未发现任何 bin 副本（尚未构建？）');
        console.log('✅ 所有 app-mcp.js 副本与单一事实源一致');
    } else {
        console.error('❌ ' + mismatched + ' 个副本与单一事实源不一致，请运行 Tools/sync-ui-to-debug.js 重新同步');
    }
}

process.exit(mismatched === 0 ? 0 : 1);
