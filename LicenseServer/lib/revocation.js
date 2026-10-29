'use strict';
/**
 * 激活码吊销名单（Revocation List）。
 *
 * 设计要点：
 * 1. 名单里**绝不保存激活码原文**，只保存 SHA256(key) 的前 16 位十六进制。
 *    这样名单可以公开分发（GET /api/revoked 无需鉴权），即使整份泄露也无法还原出可用激活码。
 * 2. version 单调递增：客户端可用它判断"我缓存的名单是不是旧的"，避免每次全量比对。
 * 3. 原子写（tmp + rename），避免运营台与 admin API 并发写坏文件。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// 吊销标识长度（hex 字符数）。16 hex = 64 bit，碰撞概率可忽略，且名单体积极小。
const HASH_LEN = 16;
const HASH_RE = /^[0-9a-f]{16}$/;

/** 激活码 → 吊销标识（不可逆，不含原文信息）。 */
function hashKey(key) {
    return crypto.createHash('sha256')
        .update(String(key === undefined || key === null ? '' : key), 'utf8')
        .digest('hex')
        .slice(0, HASH_LEN);
}

/** 判断一个字符串是否已经是吊销标识（16 位小写 hex）。 */
function isHash(s) {
    return HASH_RE.test(String(s === undefined || s === null ? '' : s));
}

class Revocation {
    constructor(dataDir) {
        this.dir = dataDir;
        this.file = path.join(dataDir, 'revoked.json');
    }

    read() {
        try {
            const raw = fs.readFileSync(this.file, 'utf8');
            const o = JSON.parse(raw);
            const revoked = Array.isArray(o.revoked)
                ? o.revoked.filter(isHash)
                : [];
            return {
                version: Number(o.version) || 0,
                updatedAt: typeof o.updatedAt === 'string' ? o.updatedAt : '',
                revoked: revoked
            };
        } catch (e) {
            // 文件不存在或损坏 → 视为空名单（不是错误，服务端仍可启动）
            return { version: 0, updatedAt: '', revoked: [] };
        }
    }

    write(state) {
        if (!fs.existsSync(this.dir)) fs.mkdirSync(this.dir, { recursive: true });
        const tmp = this.file + '.tmp';
        fs.writeFileSync(tmp, JSON.stringify(state, null, 2), 'utf8');
        fs.renameSync(tmp, this.file);
    }

    list() {
        return this.read();
    }

    count() {
        return this.read().revoked.length;
    }

    /**
     * 吊销一个激活码。入参可以是激活码原文，也可以是已算好的吊销标识。
     * 返回 { added, hash, version }；重复吊销返回 added:false 且不改 version。
     */
    add(keyOrHash) {
        const raw = String(keyOrHash === undefined || keyOrHash === null ? '' : keyOrHash).trim();
        if (!raw) return { added: false, hash: '', version: this.read().version };
        const h = isHash(raw) ? raw : hashKey(raw);
        const s = this.read();
        if (s.revoked.indexOf(h) >= 0) {
            return { added: false, hash: h, version: s.version };
        }
        s.revoked.push(h);
        s.version = (s.version || 0) + 1;
        s.updatedAt = new Date().toISOString();
        this.write(s);
        return { added: true, hash: h, version: s.version };
    }

    /** 解除吊销（误操作回滚）。 */
    remove(hash) {
        const h = String(hash || '').trim();
        const s = this.read();
        const i = s.revoked.indexOf(h);
        if (i < 0) return { removed: false, version: s.version };
        s.revoked.splice(i, 1);
        s.version = (s.version || 0) + 1;
        s.updatedAt = new Date().toISOString();
        this.write(s);
        return { removed: true, version: s.version };
    }

    hasKey(key) {
        return this.read().revoked.indexOf(hashKey(key)) >= 0;
    }

    hasHash(h) {
        return this.read().revoked.indexOf(String(h || '')) >= 0;
    }
}

module.exports = { Revocation, hashKey, isHash, HASH_LEN };
