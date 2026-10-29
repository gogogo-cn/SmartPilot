'use strict';
/**
 * compat.test.js — 运营台与 gen-key 命令行的 CSV 互认守栏。
 *
 * 这是本目录最重要的测试：两条写入链路（命令行 CLI / 运营台 API）必须
 * 写同一套格式，且运营台的解析器必须能读回 CLI 写的每一行（含引号转义）。
 * 哪天任一边改了 CSV 格式而忘了同步，这里先红。
 *
 * ⚠️ CLI 带 --out 时会追加真实台账（LicenseServer/data/issued-keys.csv），
 * 因此每个用例都先快照、finally 还原，绝不在用户账本里留测试行。
 */
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execFileSync } = require('child_process');
const ledger = require('../lib/ledger');
const lk = require('../../LicenseServer/lib/licenseKey.js');

const CLI = path.join(__dirname, '..', '..', 'LicenseServer', 'tools', 'gen-license-key.js');
const LEDGER = path.join(__dirname, '..', '..', 'LicenseServer', 'data', 'issued-keys.csv');

/** 快照真实台账；返回 restore()。文件原本不存在则还原为“不存在”。 */
function snapshotLedger() {
    const existed = fs.existsSync(LEDGER);
    const original = existed ? fs.readFileSync(LEDGER) : null;
    return function restore() {
        if (existed) fs.writeFileSync(LEDGER, original);
        else if (fs.existsSync(LEDGER)) fs.unlinkSync(LEDGER);
    };
}

test('CLI 写的台账行，运营台解析器能完整读回（含逗号备注）', () => {
    const restore = snapshotLedger();
    try {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-compat-'));
        const note = '兼容测试, 张三 微信订单 20261021';
        const r = execFileSync(process.execPath, [
            CLI, '--login', 'compat@example.com', '--plan', 'pro_yearly',
            '--days', '365', '--count', '3', '--out', path.join(dir, 'keys.csv'), '--note', note
        ], { encoding: 'utf8' });
        assert.ok(r.includes('台账已登记'), 'CLI 应登记台账');
        assert.ok(fs.existsSync(LEDGER), 'CLI --out 应追加主台账');

        const rows = ledger.readLedger(LEDGER).slice(-3);
        assert.strictEqual(rows.length, 3);
        for (const row of rows) {
            assert.strictEqual(lk.verify(row.key).valid, true, 'CLI 写的 key 必须能验签');
            assert.ok(row.login.startsWith('compat'), '批量 login 应带 +N 后缀');
            assert.strictEqual(row.note, note, '含逗号的备注必须完整还原');
        }
    } finally {
        restore();
    }
});

test('运营台写的行与 CLI 逐字节同构（除时间戳外完全一致）', () => {
    const restore = snapshotLedger();
    try {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-compat2-'));
        execFileSync(process.execPath, [
            CLI, '--login', 'fmt@example.com', '--plan', 'team', '--seats', '2',
            '--days', '31', '--out', path.join(dir, 'tmp.csv')
        ], { encoding: 'utf8', stdio: 'pipe' });

        const cliRow = ledger.readLedger(LEDGER).slice(-1)[0];
        const cliText = fs.readFileSync(LEDGER, 'utf8').split('\r\n').filter(Boolean).slice(-1)[0];

        const opsFile = path.join(dir, 'ops.csv');
        ledger.appendRows(opsFile, [{
            key: cliRow.key, login: cliRow.login, plan: cliRow.plan,
            seats: cliRow.seats, expiresAt: cliRow.expiresAt
        }], '');
        const opsText = fs.readFileSync(opsFile, 'utf8').split('\r\n')[1];

        const strip = (s) => s.replace(/^[^,]+,/, ''); // 去掉 issuedAt 时间戳列
        assert.strictEqual(strip(opsText), strip(cliText), '除时间戳外必须逐字节一致');
        assert.ok(/\r\n$/.test(fs.readFileSync(opsFile, 'utf8')), '运营台也必须用 CRLF');
        assert.strictEqual(ledger.HEADER, 'issuedAt,key,login,plan,seats,expiresAt,note');
    } finally {
        restore();
    }
});

test('运营台生成的 key 与 CLI 生成的 key 算法同源（互相验签）', () => {
    // 不带 --out，CLI 只打印 key 不动台账
    const cliOut = execFileSync(process.execPath, [CLI, '--login', 'x@y.z', '--plan', 'pro_yearly', '--days', '30'],
        { encoding: 'utf8' });
    const cliKey = cliOut.split(/\r?\n/)[0].trim();
    assert.strictEqual(lk.verify(cliKey).valid, true, 'lib 应验通过 CLI 生成的 key');

    const libKey = lk.createKey({ login: 'x@y.z', plan: 'pro_yearly', days: 30 });
    const r = execFileSync(process.execPath, [CLI, '--verify', libKey], { encoding: 'utf8' });
    assert.ok(r.includes('"valid": true'), 'CLI --verify 应认可 lib 生成的 key');
});

test('快照还原有效性：连续跑两遍用例后台账行数不变', () => {
    const before = fs.existsSync(LEDGER) ? ledger.readLedger(LEDGER).length : 0;
    const restore = snapshotLedger();
    try {
        const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'ops-compat3-'));
        execFileSync(process.execPath, [
            CLI, '--login', 'snap@example.com', '--plan', 'pro_monthly', '--days', '31',
            '--out', path.join(dir, 'k.csv')
        ], { encoding: 'utf8', stdio: 'pipe' });
        assert.strictEqual(ledger.readLedger(LEDGER).length, before + 1);
    } finally {
        restore();
    }
    const after = fs.existsSync(LEDGER) ? ledger.readLedger(LEDGER).length : 0;
    assert.strictEqual(after, before, '还原后台账行数必须与快照一致');
});
