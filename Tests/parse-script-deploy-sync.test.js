/**
 * 内置解析脚本部署一致性闸门。
 *
 * 背景：运行时 ParseScriptManager 读的是「输出目录」的 Config\DefaultParseScripts\*.md
 * （PathHelper.ConfigDirectory → bin\Debug\Config），再缓存到
 * %APPDATA%\ExcelAIVSTOAddIn\ParseScripts。只改仓库源码 Config\ 而不部署，
 * 就会出现「测试全绿、用户那边还是旧逻辑（尺寸/配合解析不出来）」。
 *
 * 运行：node Tests/parse-script-deploy-sync.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const srcDir = path.join(root, 'Config', 'DefaultParseScripts');
const outDir = path.join(root, 'bin', 'Debug', 'Config', 'DefaultParseScripts');

function scriptBody(file) {
    const md = fs.readFileSync(file, 'utf8').replace(/\r\n/g, '\n');
    const m = md.match(/```javascript\n([\s\S]*?)\n```/);
    return m ? m[1] : '';
}

test('仓库源码的内置解析脚本必须已部署到输出目录', () => {
    if (!fs.existsSync(outDir)) {
        // 从未构建过的机器上跳过，而不是误报失败
        return;
    }
    const srcFiles = fs.readdirSync(srcDir).filter(f => f.endsWith('.md'));
    assert.ok(srcFiles.length > 0, '源码目录应有内置脚本');

    const stale = [];
    const missing = [];
    for (const name of srcFiles) {
        const outFile = path.join(outDir, name);
        if (!fs.existsSync(outFile)) {
            missing.push(name);
            continue;
        }
        if (scriptBody(path.join(srcDir, name)) !== scriptBody(outFile)) {
            stale.push(name);
        }
    }
    assert.deepStrictEqual(missing, [], '这些内置脚本还没部署到 ' + outDir);
    assert.deepStrictEqual(stale, [], '这些内置脚本的部署副本已过期，请复制源码到 ' + outDir);
});
