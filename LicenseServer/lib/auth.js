'use strict';
const crypto = require('crypto');

const SCRYPT_N = 16384;
const SCRYPT_R = 8;
const SCRYPT_P = 1;
const KEYLEN = 32;

function hashPassword(password) {
    const salt = crypto.randomBytes(16);
    const hash = crypto.scryptSync(String(password), salt, KEYLEN, {
        N: SCRYPT_N, r: SCRYPT_R, p: SCRYPT_P
    });
    return 'scrypt$' + salt.toString('hex') + '$' + hash.toString('hex');
}

function verifyPassword(password, stored) {
    try {
        const parts = String(stored || '').split('$');
        if (parts.length !== 3 || parts[0] !== 'scrypt') return false;
        const salt = Buffer.from(parts[1], 'hex');
        const expect = Buffer.from(parts[2], 'hex');
        const hash = crypto.scryptSync(String(password), salt, expect.length, {
            N: SCRYPT_N, r: SCRYPT_R, p: SCRYPT_P
        });
        return crypto.timingSafeEqual(hash, expect);
    } catch (_) {
        return false;
    }
}

function publicUser(u) {
    if (!u) return null;
    return { id: u.id, email: u.email, phone: u.phone, createdAt: u.createdAt };
}

function entitlementView(ent) {
    if (!ent) {
        return {
            plan: 'none',
            status: 'inactive',
            seats: 0,
            periodEnd: null,
            active: false,
            byok: true
        };
    }
    const end = ent.periodEnd ? new Date(ent.periodEnd).getTime() : 0;
    const active = ent.status === 'active' && end > Date.now();
    return {
        plan: ent.plan,
        status: active ? 'active' : (ent.status === 'active' ? 'expired' : ent.status),
        seats: ent.seats || 1,
        periodEnd: ent.periodEnd,
        active,
        provider: ent.provider || null,
        byok: true,
        note: 'Model API costs are paid by you to your model provider. This subscription is SmartPilot software only.'
    };
}

module.exports = { hashPassword, verifyPassword, publicUser, entitlementView };
