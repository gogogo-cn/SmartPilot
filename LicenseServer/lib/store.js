'use strict';
/**
 * JSON file store — no native DB dependency for the MVP license API.
 * Tables: users, sessions, entitlements, subscriptions, orders
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

function nowIso() {
    return new Date().toISOString();
}

function id(prefix) {
    return prefix + '_' + crypto.randomBytes(12).toString('hex');
}

function loadEnvFile(root) {
    const envPath = path.join(root, '.env');
    if (!fs.existsSync(envPath)) return;
    const text = fs.readFileSync(envPath, 'utf8');
    for (const line of text.split(/\r?\n/)) {
        const t = line.trim();
        if (!t || t.startsWith('#')) continue;
        const i = t.indexOf('=');
        if (i < 1) continue;
        const k = t.slice(0, i).trim();
        let v = t.slice(i + 1).trim();
        if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) {
            v = v.slice(1, -1);
        }
        if (process.env[k] == null || process.env[k] === '') process.env[k] = v;
    }
}

class Store {
    constructor(dataDir) {
        this.dataDir = dataDir;
        this.file = path.join(dataDir, 'license-db.json');
        this._db = { users: [], sessions: [], entitlements: [], subscriptions: [], orders: [] };
        this._ensure();
    }

    _ensure() {
        if (!fs.existsSync(this.dataDir)) fs.mkdirSync(this.dataDir, { recursive: true });
        if (fs.existsSync(this.file)) {
            try {
                const raw = JSON.parse(fs.readFileSync(this.file, 'utf8'));
                this._db = Object.assign(this._db, raw);
            } catch (_) { /* keep empty */ }
        } else {
            this.save();
        }
    }

    save() {
        const tmp = this.file + '.tmp';
        fs.writeFileSync(tmp, JSON.stringify(this._db, null, 2), 'utf8');
        fs.renameSync(tmp, this.file);
    }

    createUser({ email, phone, passwordHash }) {
        const user = {
            id: id('usr'),
            email: (email || '').trim().toLowerCase() || null,
            phone: (phone || '').trim() || null,
            passwordHash,
            createdAt: nowIso()
        };
        this._db.users.push(user);
        this.save();
        return user;
    }

    findUserByLogin(login) {
        const key = String(login || '').trim().toLowerCase();
        if (!key) return null;
        return this._db.users.find(u =>
            (u.email && u.email === key) ||
            (u.phone && u.phone === key) ||
            (u.phone && u.phone === String(login || '').trim())
        ) || null;
    }

    findUserById(userId) {
        return this._db.users.find(u => u.id === userId) || null;
    }

    createSession(userId, days) {
        const expiresAt = new Date(Date.now() + (days || 30) * 86400000).toISOString();
        const session = {
            id: id('ses'),
            token: crypto.randomBytes(32).toString('hex'),
            userId,
            expiresAt,
            createdAt: nowIso()
        };
        this._db.sessions.push(session);
        this.save();
        return session;
    }

    findSession(token) {
        if (!token) return null;
        const s = this._db.sessions.find(x => x.token === token);
        if (!s) return null;
        if (new Date(s.expiresAt).getTime() < Date.now()) return null;
        return s;
    }

    revokeSessionsForUser(userId) {
        this._db.sessions = this._db.sessions.filter(s => s.userId !== userId);
        this.save();
    }

    getEntitlement(userId) {
        return this._db.entitlements.find(e => e.userId === userId) || null;
    }

    upsertEntitlement(userId, patch) {
        let e = this.getEntitlement(userId);
        if (!e) {
            e = {
                userId,
                plan: 'trial',
                seats: 1,
                status: 'active',
                periodEnd: null,
                updatedAt: nowIso()
            };
            this._db.entitlements.push(e);
        }
        Object.assign(e, patch, { updatedAt: nowIso() });
        this.save();
        return e;
    }

    createSubscription(row) {
        const sub = Object.assign({
            id: id('sub'),
            createdAt: nowIso(),
            updatedAt: nowIso()
        }, row);
        this._db.subscriptions.push(sub);
        this.save();
        return sub;
    }

    findSubscriptionByProviderId(provider, providerSubId) {
        return this._db.subscriptions.find(s =>
            s.provider === provider && s.providerSubId === providerSubId
        ) || null;
    }

    updateSubscription(subId, patch) {
        const s = this._db.subscriptions.find(x => x.id === subId);
        if (!s) return null;
        Object.assign(s, patch, { updatedAt: nowIso() });
        this.save();
        return s;
    }

    createOrder(row) {
        const order = Object.assign({
            id: id('ord'),
            status: 'pending',
            createdAt: nowIso(),
            updatedAt: nowIso()
        }, row);
        this._db.orders.push(order);
        this.save();
        return order;
    }

    findOrder(orderId) {
        return this._db.orders.find(o => o.id === orderId) || null;
    }

    updateOrder(orderId, patch) {
        const o = this.findOrder(orderId);
        if (!o) return null;
        Object.assign(o, patch, { updatedAt: nowIso() });
        this.save();
        return o;
    }
}

function hashPassword(password, salt) {
    const s = salt || crypto.randomBytes(16).toString('hex');
    const hash = crypto.pbkdf2Sync(String(password), s, 120000, 32, 'sha256').toString('hex');
    return s + ':' + hash;
}

function verifyPassword(password, stored) {
    if (!stored || !stored.includes(':')) return false;
    const [salt, hash] = stored.split(':');
    const check = crypto.pbkdf2Sync(String(password), salt, 120000, 32, 'sha256').toString('hex');
    try {
        return crypto.timingSafeEqual(Buffer.from(hash, 'hex'), Buffer.from(check, 'hex'));
    } catch (_) {
        return false;
    }
}

function entitlementView(ent) {
    if (!ent) {
        return {
            plan: 'none',
            seats: 0,
            status: 'inactive',
            periodEnd: null,
            active: false,
            reason: 'no_entitlement'
        };
    }
    const end = ent.periodEnd ? new Date(ent.periodEnd).getTime() : null;
    const active = ent.status === 'active' && (end == null || end > Date.now());
    return {
        plan: ent.plan,
        seats: ent.seats || 1,
        status: active ? 'active' : (ent.status === 'active' ? 'expired' : ent.status),
        periodEnd: ent.periodEnd,
        active,
        reason: active ? 'ok' : (end != null && end <= Date.now() ? 'expired' : 'inactive')
    };
}

module.exports = {
    Store,
    loadEnvFile,
    hashPassword,
    verifyPassword,
    entitlementView,
    id,
    nowIso
};
