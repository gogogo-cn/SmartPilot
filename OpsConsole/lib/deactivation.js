'use strict';
/**
 * deactivation.js — 「消 key 退款」闭环（运营台侧）。
 *
 * 业务闭环：
 *   1. 用户在本机点「注销并申请退款」→ 本机激活被清除，该 key 在本机永久失效
 *   2. 客户端生成「消 key 码」（SPD1-...）→ 用户发给客服
 *   3. 运营台用本模块验签 → 定位台账 → 确认未重复退款 → 标记退款 + 加吊销名单
 *
 * 为什么要加吊销名单：本机注销只让**那台机器**失效。若用户事先备份了激活码，
 * 换一台机器仍能激活。退款后写入吊销名单，联网客户端拉到名单后即可拦住。
 */
const path = require('path');
const lk = require(path.join(__dirname, '..', '..', 'LicenseServer', 'lib', 'licenseKey.js'));
const { Revocation, hashKey } = require(path.join(__dirname, '..', '..', 'LicenseServer', 'lib', 'revocation.js'));

/**
 * 验证注销码并定位台账记录。**不产生任何副作用**（只读）。
 * @param {string} code 用户发来的消 key 码
 * @param {Array<{key:string,login:string,plan:string,expiresAt:string,status:string}>} rows 台账行
 * @returns {{ok:boolean, error?:string, keyHash?:string, plan?:string, destroyedAt?:string,
 *            row?:object, alreadyRefunded?:boolean}}
 */
function verifyProof(code, rows) {
    const v = lk.verifyDeactivationProof(code);
    if (!v.valid) {
        return {
            ok: false,
            error: v.reason === 'bad_signature'
                ? '注销码无效或已被篡改 —— 不要退款'
                : '注销码格式不对（应以 SPD1- 开头，请让用户重新复制完整）'
        };
    }
    const row = (rows || []).find(r => hashKey(r.key) === v.keyHash);
    if (!row) {
        return { ok: false, error: '台账里查不到这张激活码（可能不是本系统签发的）' };
    }
    return {
        ok: true,
        keyHash: v.keyHash,
        plan: v.plan,
        destroyedAt: v.destroyedAt,
        row: row,
        alreadyRefunded: row.status === 'refunded'
    };
}

/**
 * 验签通过后的退款动作：标记 refunded + 写入吊销名单。
 * @param {{code:string, store:object, dataDir:string}} opts
 *        dataDir 为 LicenseServer 的 data 目录（吊销名单存放处）
 * @returns {{ok:boolean, error?:string, row?:object, state?:object, revoked?:object}}
 */
function refundByProof(opts) {
    opts = opts || {};
    const store = opts.store;
    const v = verifyProof(opts.code, store.rows());
    if (!v.ok) return v;
    if (v.alreadyRefunded) {
        return { ok: false, error: '这张激活码已经退过款了（禁止重复退款）', row: v.row };
    }

    const state = store.setStatus(v.row.key, {
        status: 'refunded',
        opsNote: '消key退款 ' + (v.destroyedAt || '')
    });

    // 联动吊销：换机备份的那份也一并作废
    let revoked = null;
    try {
        if (opts.dataDir) {
            const rev = new Revocation(opts.dataDir);
            const r = rev.add(v.row.key);
            revoked = { hash: r.hash, version: r.version, count: rev.count() };
        }
    } catch (e) {
        revoked = { error: e && e.message ? e.message : String(e) };
    }

    return { ok: true, row: v.row, state: state, revoked: revoked, destroyedAt: v.destroyedAt };
}

module.exports = { verifyProof, refundByProof, hashKey };
