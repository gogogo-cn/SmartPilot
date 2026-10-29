/**
 * license.js — BYOK 席位订阅客户端
 * 订阅覆盖 SmartPilot 软件许可（含 AI 对话 + OCR 等模块）；模型 API Key 仍在本机。
 * 主入口：侧栏「账户」按钮 → 账户抽屉（非埋在 AI 设置里）。
 */
(function () {
    'use strict';

    var ChatLog = window.ChatLog || console;

    function t(key, vars) {
        try {
            if (window.App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(key, vars);
        } catch (_) { /* ignore */ }
        var fallback = {
            'license.card': '账户 / 订阅',
            'license.byok': '订阅覆盖 SmartPilot 软件许可（含 AI 对话与 OCR 等模块）。模型与云 OCR API Key 由你自行配置，费用与对应服务商结算。',
            'license.status.loading': '正在读取订阅状态…',
            'license.status.none': '未登录',
            'license.status.activeOffline': '已激活（激活码）：{plan}，到期 {end}',
            'license.btnLogin': '登录',
            'license.btnRegister': '注册试用',
            'license.btnRefresh': '刷新状态',
            'license.btnStripeMonth': '升级（Stripe 月付）',
            'license.btnStripeYear': '升级（Stripe 年付）',
            'license.btnCnYear': '升级（国内年付）',
            'license.btnPortal': '管理账单',
            'license.btnLogout': '退出登录',
            'license.login': '邮箱或手机',
            'license.password': '密码',
            'nav.account': '账户 / 订阅',
            'license.badge.trial': '试用 / 未订阅',
            'license.badge.active': '已订阅',
            'license.badge.none': '账户',
            'license.err.credentials': '请填写账号和密码',
            'license.err.auth': '登录或注册失败',
            'license.err.weakPassword': '密码至少 8 位',
            'license.err.serverDown': '暂时无法连接授权服务，请稍后重试。',
            'license.err.noBridge': '无法连接插件桥接，请从 Office 任务窗格内操作。',
            'license.passwordHint': '密码至少 8 位',
            'license.server.ok': '授权服务：已连接',
            'license.server.down': '授权服务：未连接',
            'license.server.offline': '授权服务：未连接（激活码离线可用）',
            'license.server.bridgeOnly': '授权服务：将通过插件连接（页面无法直连 localhost）',
            'license.gate.blocked': '试用已结束或订阅未生效。请打开「账户 / 订阅」登录。付费通道开通前，过期账户不能继续使用需授权的功能。',
            'license.checkout.pending': '付费通道尚未开通。试用期内可以继续使用。不会进入模拟支付。',
            'license.trialScope': '注册试用为 14 天，期内可使用对话和云端识别。到期后这两项停止，直到订阅生效。付费尚未开通时不会出现升级按钮。未登录也可以先用本机识图。',
            'license.gate.needAccount': '对话和云端识别需要登录。打开侧栏「账户」注册试用。付费尚未开通时，本机识图可以先用。',
            'license.openAccount': '打开账户与订阅',
            'license.scope': '软件许可覆盖 AI 与 OCR',
            'license.key': '激活码',
            'license.key.hint': '购买后收到的激活码（SPK1- 开头），粘贴后点「激活」。激活码在本机校验，不需要联网，也不需要登录账户。',
            'license.key.hintRenew': '已激活中。续费时点「输入新激活码」，粘贴新码后点「激活」即可延长有效期。',
            'license.key.placeholder': 'SPK1-…',
            'license.btnActivate': '激活',
            'license.btnRenewKey': '输入新激活码',
            'license.btnClearKey': '清除激活',
            'license.local.active': '离线激活：{plan}，到期 {end}',
            'license.local.none': '未使用激活码',
            'license.key.ok': '激活成功',
            'license.key.cleared': '已清除本地激活',
            'ocr.quota.exhausted': '今日免费额度已用完（{limit} 次/天），明天 0 点自动重置；激活后不限次数。',
            'ocr.quota.badge': '今日剩余 {remaining}/{limit}',
            'ocr.quota.unlimited': '已激活 · 不限次数',
            'ocr.quota.cloudFree': '云端识别不计额度',
            'license.btnRefund': '注销并申请退款',
            'license.refund.confirm': '确定注销本机激活吗？注销后这台机器立刻失去授权，且这张激活码不能再在本机激活。你会拿到一个「消 key 码」，发给客服验证后才会退款。',
            'license.refund.codeTitle': '消 key 码（发给客服完成退款）',
            'license.refund.codeHint': '把它复制发给客服。客服验证通过后会为你退款。此码不能恢复授权——如需再用请重新购买。',
            'license.refund.copy': '复制消 key 码',
            'license.refund.copied': '已复制，发给客服即可',
            'license.refund.done': '已注销，本机授权已失效',
            'license.refund.none': '当前没有可注销的离线激活',
            'license.refund.failed': '注销失败',
            'license.key.empty': '请先粘贴激活码',
            'license.buy.title': '购买 / 续费',
            'license.buy.titleRenew': '续费',
            'license.buy.price': '激活码方案（付款时微信里说明选哪个）',
            'license.buy.planMonthly': '专业版 · 月付 —— ¥19 / 月',
            'license.buy.planYearly': '专业版 · 年付 —— ¥199 / 年（含 1 个席位）',
            'license.buy.planTeam': '团队版（5人以上） —— ¥129 / 席位 / 年（几人买几席，每人一个独立激活码，总价 = 129 × 人数）',
            'license.buy.planHint': '有效期由激活码决定：月付 30 天、年付 365 天、团队版按年。激活后本机离线可用，不需要联网、不需要登录账户。',
            'license.buy.step1': '扫码添加微信（也可直接复制下方联系方式）',
            'license.buy.step2': '微信转账付款，备注里留下你的邮箱或微信号',
            'license.buy.step3': '把付款成功截图发给我，几分钟内回复激活码，回到上方粘贴激活即可',
            'license.buy.qr': '微信二维码（扫码添加）',
            'license.buy.qrMissing': '微信二维码未配置（放到 UI/images/wechat-qr.jpg），请用下方联系方式添加',
            'license.buy.contactLabel': '联系方式',
            'license.buy.copyContact': '复制联系方式',
            'license.buy.copied': '联系方式已复制',
            'license.buy.copyFailed': '复制失败，请手动选择复制',
            'license.buy.contactEmpty': '联系方式未配置，请通过文档中的渠道联系作者',
            'license.ok.auth': '登录成功',
            'license.ok.logout': '已退出登录'
        };
        return fallback[key] || key;
    }

    function notify(msg, level) {
        try {
            if (typeof App !== 'undefined' && typeof App.showToast === 'function') App.showToast(msg, level || 'info');
            else if (typeof window.showToast === 'function') window.showToast(msg, level || 'info');
        } catch (_) { /* ignore */ }
        var statusEl = document.getElementById('license-status');
        if (statusEl && (level === 'error' || level === 'warning')) statusEl.textContent = msg;
    }

    function friendlyAuthError(res) {
        if (!res) return t('license.err.auth');
        if (res.code === 'no_bridge') return t('license.err.noBridge');
        var err = String(res.error || res.code || '');
        if (/refused|无法连接|timed out|timeout|ECONNREFUSED|actively refused|名称.*解析|failed to fetch|网络/i.test(err)) {
            return t('license.err.serverDown');
        }
        if (res.code === 'weak_password') return t('license.err.weakPassword');
        if (res.code === 'checkout_not_ready') return t('license.checkout.pending');
        return err || t('license.err.auth');
    }

    var checkoutReady = { stripe: false, cn: false };
    // 离线激活码是否生效（三端共享会话里查到的）。路线①用户本就不依赖授权服务，
    // 服务器连不上时不该弹红色警告打扰他们。
    var offlineActivationActive = false;

    // ---- 购买 / 续费面板（路线①：人工收款，无在线支付）----
    // 开卖前只改这两项：
    //   qrImage  ：微信二维码图片（客户扫码加你好友后再转账），路径相对 UI 根目录；
    //              当前用 UI/images/wechat-qr.jpg。文件缺失时自动隐藏图片，
    //              提示改用联系方式。将来若拿到微信收款码，换成收款码图即可。
    //   contact  ：文字版联系方式（微信 / QQ / 邮箱），供客户复制。
    // 价格文案在 i18n.js 的 license.buy.price（中英各一处）。
    var PURCHASE = {
        qrImage: 'images/wechat-qr.jpg',
        contact: '微信：hellokitty360730'
    };

    async function pingLicenseServer() {
        var el = document.getElementById('license-server-status');
        var base = 'http://localhost:9530';
        var info = await call('LicenseGetApiBaseUrl');
        if (info && info.baseUrl) base = String(info.baseUrl).replace(/\/$/, '');

        // 优先走 .NET 桥探测：WebView2 对 localhost 的 fetch 经常失败，不能当真实状态
        var bridgePing = await call('LicensePingHealth');
        if (bridgePing && bridgePing.code !== 'no_bridge') {
            var okBridge = bridgePing.ok === true;
            checkoutReady.stripe = bridgePing.stripeReady === true;
            checkoutReady.cn = bridgePing.cnReady === true;
            if (el) {
                if (okBridge) {
                    el.textContent = t('license.server.ok');
                    el.className = 'license-server-status is-ok';
                } else if (offlineActivationActive) {
                    el.textContent = t('license.server.offline');
                    el.className = 'license-server-status is-neutral';
                } else {
                    el.textContent = t('license.server.down');
                    el.className = 'license-server-status is-down';
                }
            }
            applyCheckoutVisibility();
            return okBridge;
        }

        // 旧插件未暴露 LicensePingHealth：有桥则允许登录（由桥发 HTTP），仅用 fetch 做辅助显示
        var hasBridge = !!(info && info.code !== 'no_bridge' && info.success !== false);
        var fetchOk = false;
        try {
            var ctrl = new AbortController();
            var tm = setTimeout(function () { ctrl.abort(); }, 1200);
            var r = await fetch(base + '/health', { signal: ctrl.signal, mode: 'cors' });
            clearTimeout(tm);
            if (r && r.ok) {
                fetchOk = true;
                try {
                    var body = await r.json();
                    checkoutReady.stripe = !!(body && body.stripeReady === true);
                    checkoutReady.cn = !!(body && body.cnReady === true);
                } catch (_) { /* ignore */ }
            }
        } catch (_) { fetchOk = false; }
        if (el) {
            if (fetchOk) {
                el.textContent = t('license.server.ok');
                el.className = 'license-server-status is-ok';
            } else             if (hasBridge) {
                el.textContent = t('license.server.bridgeOnly');
                el.className = 'license-server-status is-ok';
            } else if (offlineActivationActive) {
                el.textContent = t('license.server.offline');
                el.className = 'license-server-status is-neutral';
            } else {
                el.textContent = t('license.server.down');
                el.className = 'license-server-status is-down';
            }
        }
        applyCheckoutVisibility();
        return fetchOk || hasBridge;
    }

    function applyCheckoutVisibility() {
        var root = document.getElementById('account-license-drawer') || document;
        ['stripe-m', 'stripe-y', 'portal'].forEach(function (id) {
            var btn = root.querySelector('[data-lic="' + id + '"]');
            if (btn) btn.style.display = checkoutReady.stripe ? '' : 'none';
        });
        var cn = root.querySelector('[data-lic="cn-y"]');
        if (cn) cn.style.display = checkoutReady.cn ? '' : 'none';
        var note = document.getElementById('license-checkout-note');
        // 离线激活码已生效的用户既不在试用、也不等付费通道，这条提示对他们纯属噪音
        if (note) note.style.display = (offlineActivationActive || checkoutReady.stripe || checkoutReady.cn) ? 'none' : 'block';
    }

    function bridge() {
        try {
            if (window.BridgeAccessor && typeof BridgeAccessor.getBridge === 'function') {
                var b = BridgeAccessor.getBridge();
                if (b) return b;
            }
        } catch (_) { /* ignore */ }
        try {
            return window.chrome && chrome.webview && chrome.webview.hostObjects && chrome.webview.hostObjects.excelBridge;
        } catch (_) {
            return null;
        }
    }

    async function call(method, arg) {
        var b = bridge();
        if (!b || typeof b[method] !== 'function') {
            return { success: false, error: 'License bridge unavailable', code: 'no_bridge' };
        }
        var raw = arg === undefined ? await b[method]() : await b[method](arg);
        if (typeof raw === 'string') {
            try { raw = JSON.parse(raw); } catch (_) { return { success: false, error: raw, code: 'bridge_failed' }; }
        }
        if (raw && raw.error === 'bridge-call-failed') {
            return { success: false, code: 'bridge_failed', error: raw.detail || raw.error };
        }
        return raw || { success: false, error: 'empty', code: 'bridge_failed' };
    }

    function formatEntitlement(ent) {
        if (!ent) return t('license.status.none');
        var reason = String(ent.reason || ent.status || '');
        if (!ent.active && /not_logged_in|logged_out/i.test(reason)) return t('license.status.none');
        if (ent.active) {
            var end = ent.periodEnd ? String(ent.periodEnd).slice(0, 10) : '';
            // 路线①离线激活码：标明来源，避免已激活用户看到"未登录"式文案以为还得注册账户
            if (ent.source === 'offline') {
                return t('license.status.activeOffline', { plan: ent.plan || '', end: end });
            }
            return t('license.status.active', { plan: ent.plan || '', end: end });
        }
        return t('license.status.inactive', { reason: ent.reason || ent.status || '' });
    }

    function esc(s) {
        return String(s == null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }

    function licenseCardInnerHtml(opts) {
        opts = opts || {};
        var title = opts.includeTitle === false
            ? ''
            : '<div class="license-card-title" data-i18n="license.card">' + t('license.card') + '</div>';
        return '' + title +
            '<p id="license-byok-hint" class="license-byok-hint" data-i18n="license.byok">' + t('license.byok') + '</p>' +
            '<p id="license-trial-scope" class="license-byok-hint" data-i18n="license.trialScope">' + t('license.trialScope') + '</p>' +
            '<p id="license-checkout-note" class="license-byok-hint" data-i18n="license.checkout.pending">' + t('license.checkout.pending') + '</p>' +
            '<div id="license-scope-chip" class="license-scope-chip" data-i18n="license.scope">' + t('license.scope') + '</div>' +
            '<div id="license-server-status" class="license-server-status">…</div>' +
            '<div id="license-status" class="license-status" data-i18n="license.status.loading">' + t('license.status.loading') + '</div>' +
            '<div id="license-key" class="license-key">' +
            '  <div class="license-key-title" data-i18n="license.key">' + t('license.key') + '</div>' +
            '  <div class="license-row" id="license-key-input-row"><input type="text" id="license-key-input" class="license-input" placeholder="' + t('license.key.placeholder') + '" autocomplete="off" spellcheck="false"></div>' +
            '  <div class="license-btn-row">' +
            '    <button type="button" class="license-btn primary" data-lic="activate" id="license-activate-btn" data-i18n="license.btnActivate">' + t('license.btnActivate') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="clear-key" id="license-clear-key-btn" style="display:none" data-i18n="license.btnClearKey">' + t('license.btnClearKey') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="refund" id="license-refund-btn" style="display:none" data-i18n="license.btnRefund">' + t('license.btnRefund') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="renew-key" id="license-renew-btn" style="display:none" data-i18n="license.btnRenewKey">' + t('license.btnRenewKey') + '</button>' +
            '  </div>' +
            '  <div id="license-local-status" class="license-local-status is-none" data-i18n="license.local.none">' + t('license.local.none') + '</div>' +
            '  <div id="license-refund-box" class="license-refund-box" style="display:none">' +
            '    <div class="license-key-title" data-i18n="license.refund.codeTitle">' + t('license.refund.codeTitle') + '</div>' +
            '    <div class="license-row"><input type="text" id="license-refund-code" class="license-input" readonly spellcheck="false"></div>' +
            '    <div class="license-btn-row">' +
            '      <button type="button" class="license-btn primary" data-lic="copy-refund" data-i18n="license.refund.copy">' + t('license.refund.copy') + '</button>' +
            '    </div>' +
            '    <div class="license-byok-hint" data-i18n="license.refund.codeHint">' + t('license.refund.codeHint') + '</div>' +
            '  </div>' +
            '  <div id="license-key-hint" class="license-byok-hint" data-i18n="license.key.hint">' + t('license.key.hint') + '</div>' +
            '  <div id="license-key-hint-renew" class="license-byok-hint" data-i18n="license.key.hintRenew" style="display:none">' + t('license.key.hintRenew') + '</div>' +
            '</div>' +
            '<div id="license-buy" class="license-buy">' +
            '  <div class="license-key-title" id="license-buy-title" data-i18n="license.buy.title">' + t('license.buy.title') + '</div>' +
            '  <div class="license-buy-price" data-i18n="license.buy.price">' + t('license.buy.price') + '</div>' +
            '  <ul class="license-buy-plans">' +
            '    <li data-i18n="license.buy.planMonthly">' + t('license.buy.planMonthly') + '</li>' +
            '    <li data-i18n="license.buy.planYearly">' + t('license.buy.planYearly') + '</li>' +
            '    <li data-i18n="license.buy.planTeam">' + t('license.buy.planTeam') + '</li>' +
            '  </ul>' +
            '  <div class="license-byok-hint" data-i18n="license.buy.planHint">' + t('license.buy.planHint') + '</div>' +
            '  <ol class="license-buy-steps">' +
            '    <li data-i18n="license.buy.step1">' + t('license.buy.step1') + '</li>' +
            '    <li data-i18n="license.buy.step2">' + t('license.buy.step2') + '</li>' +
            '    <li data-i18n="license.buy.step3">' + t('license.buy.step3') + '</li>' +
            '  </ol>' +
            '  <div class="license-buy-qr-row">' +
            '    <img id="license-buy-qr" class="license-buy-qr" src="' + esc(PURCHASE.qrImage) + '" alt="' + t('license.buy.qr') + '">' +
            '    <div id="license-buy-qr-missing" class="license-byok-hint" style="display:none" data-i18n="license.buy.qrMissing">' + t('license.buy.qrMissing') + '</div>' +
            '  </div>' +
            '  <div class="license-buy-contact-row">' +
            '    <span class="license-buy-contact-label" data-i18n="license.buy.contactLabel">' + t('license.buy.contactLabel') + '</span>' +
            '    <span id="license-buy-contact" class="license-buy-contact">' + esc(PURCHASE.contact) + '</span>' +
            '    <button type="button" class="license-btn" data-lic="copy-contact" data-i18n="license.buy.copyContact">' + t('license.buy.copyContact') + '</button>' +
            '  </div>' +
            '</div>' +
            '<div id="license-auth" class="license-auth">' +
            '  <div class="license-row"><label for="license-login" data-i18n="license.login">' + t('license.login') + '</label>' +
            '  <input type="text" id="license-login" class="license-input" autocomplete="username" placeholder="you@example.com"></div>' +
            '  <div class="license-row"><label for="license-password" data-i18n="license.password">' + t('license.password') + '</label>' +
            '  <input type="password" id="license-password" class="license-input" autocomplete="current-password" minlength="8" placeholder="' + t('license.passwordHint') + '"></div>' +
            '  <div class="license-password-hint" data-i18n="license.passwordHint">' + t('license.passwordHint') + '</div>' +
            '  <div class="license-btn-row">' +
            '    <button type="button" class="license-btn primary" data-lic="login" data-i18n="license.btnLogin">' + t('license.btnLogin') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="register" data-i18n="license.btnRegister">' + t('license.btnRegister') + '</button>' +
            '  </div>' +
            '</div>' +
            '<div id="license-actions" class="license-actions" style="display:none;">' +
            '  <div class="license-btn-row">' +
            '    <button type="button" class="license-btn primary" data-lic="refresh" data-i18n="license.btnRefresh">' + t('license.btnRefresh') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="stripe-m" data-i18n="license.btnStripeMonth" style="display:none">' + t('license.btnStripeMonth') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="stripe-y" data-i18n="license.btnStripeYear" style="display:none">' + t('license.btnStripeYear') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="cn-y" data-i18n="license.btnCnYear" style="display:none">' + t('license.btnCnYear') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="portal" data-i18n="license.btnPortal" style="display:none">' + t('license.btnPortal') + '</button>' +
            '    <button type="button" class="license-btn" data-lic="logout" data-i18n="license.btnLogout">' + t('license.btnLogout') + '</button>' +
            '  </div>' +
            '</div>';
    }

    function bindLicenseCardClicks(root) {
        if (!root || root._licBound) return;
        root._licBound = true;
        root.addEventListener('click', function (e) {
            var btn = e.target && e.target.closest ? e.target.closest('[data-lic]') : null;
            if (!btn) return;
            var act = btn.getAttribute('data-lic');
            if (act === 'login') registerOrLogin('login');
            else if (act === 'register') registerOrLogin('register');
            else if (act === 'refresh') refreshAccountCard();
            else if (act === 'stripe-m') checkout('stripe', 'pro_monthly');
            else if (act === 'stripe-y') checkout('stripe', 'pro_yearly');
            else if (act === 'cn-y') checkout('cn', 'pro_yearly');
            else if (act === 'portal') openPortal();
            else if (act === 'logout') logout();
            else if (act === 'activate') activateKey();
            else if (act === 'renew-key') revealKeyInput();
            else if (act === 'clear-key') clearLocalActivation();
            else if (act === 'refund') deactivateForRefund();
            else if (act === 'copy-refund') copyRefundCode();
            else if (act === 'copy-contact') copyPurchaseContact();
        });
        // 收款码图片缺失/加载失败时隐藏图片，提示改用联系方式付款
        var qr = root.querySelector('#license-buy-qr');
        if (qr && !qr._licQrBound) {
            qr._licQrBound = true;
            qr.addEventListener('error', function () {
                qr.style.display = 'none';
                var miss = document.getElementById('license-buy-qr-missing');
                if (miss) miss.style.display = 'block';
            });
        }
    }

    /** 复制购买联系方式（路线①人工收款）。navigator.clipboard 在非安全上下文不可用，退回 execCommand。 */
    function copyPurchaseContact() {
        var text = String(PURCHASE.contact || '').trim();
        if (!text) {
            notify(t('license.buy.contactEmpty'), 'warning');
            return;
        }
        var done = function () { notify(t('license.buy.copied'), 'success'); };
        var fail = function () { notify(t('license.buy.copyFailed'), 'warning'); };
        try {
            if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(text).then(done, function () { legacyCopy(text, done, fail); });
                return;
            }
        } catch (_) { /* fall through */ }
        legacyCopy(text, done, fail);
    }

    function legacyCopy(text, done, fail) {
        try {
            var ta = document.createElement('textarea');
            ta.value = text;
            ta.style.position = 'fixed';
            ta.style.opacity = '0';
            document.body.appendChild(ta);
            ta.select();
            var ok = document.execCommand('copy');
            document.body.removeChild(ta);
            if (ok) done(); else fail();
        } catch (_) {
            fail();
        }
    }

    /** 单一入口：侧栏 👤 按钮上反映订阅状态（不再另做徽章，避免两个图标） */
    function updateAccountEntry(ent, loggedIn) {
        var btn = document.getElementById('side-account-btn');
        if (!btn) return;
        btn.classList.remove('lic-active', 'lic-warn', 'lic-muted');
        var label = t('nav.account');
        if (ent && ent.active) {
            btn.classList.add('lic-active');
            btn.title = label + ' — ' + formatEntitlement(ent);
        } else if (loggedIn) {
            btn.classList.add('lic-warn');
            btn.title = label + ' — ' + formatEntitlement(ent);
        } else {
            btn.classList.add('lic-muted');
            btn.title = label + ' — ' + t('license.status.none');
        }
    }

    async function refreshAccountCard() {
        var statusEl = document.getElementById('license-status');
        var authBox = document.getElementById('license-auth');
        var actions = document.getElementById('license-actions');
        var trialHint = document.getElementById('license-trial-scope');
        var result = await call('LicenseGetEntitlement');
        rememberGateResult(entitlementToGate(result));
        var ent = result && result.entitlement;
        // 告诉 OCR 免费额度模块当前是否已激活（已激活 → 不计数，见 js/ocr-quota.js）
        if (window.SmartPilotOcrQuota) {
            window.SmartPilotOcrQuota.setLicensed(!!(ent && ent.active === true) || softBypassEnabled());
        }
        if (statusEl) statusEl.textContent = formatEntitlement(ent);
        var loggedIn = !!(result && result.loggedIn);
        if (authBox) authBox.style.display = loggedIn ? 'none' : 'block';
        if (actions) actions.style.display = loggedIn ? 'flex' : 'none';
        // 已订阅/已激活的人不需要再读"注册试用 14 天"的说明
        if (trialHint) trialHint.style.display = (loggedIn || (ent && ent.active === true)) ? 'none' : '';
        var hint = document.getElementById('license-byok-hint');
        if (hint) hint.style.display = 'block';
        updateAccountEntry(ent, loggedIn);
        var local = await call('LicenseGetLocalActivation');
        var localActive = renderLocalActivation(local);
        offlineActivationActive = localActive;
        if (localActive) {
            // 离线激活码已生效（三端共用同一份共享会话）：主状态、登录框、账户入口
            // 全部按"已激活"呈现。否则会同时出现"未登录"和"离线激活：…"自相矛盾两行。
            if (statusEl) statusEl.textContent = formatEntitlement(local.entitlement);
            if (authBox) authBox.style.display = 'none';
            if (trialHint) trialHint.style.display = 'none';
            if (actions) {
                actions.style.display = 'flex';
                // 没登录过账户的人不该看到"退出登录"；刷新状态仍保留
                var logoutBtn = actions.querySelector('[data-lic="logout"]');
                if (logoutBtn) logoutBtn.style.display = 'none';
            }
            updateAccountEntry(local.entitlement, true);
            // 离线激活码用户也要解锁 OCR 无限额度（renderLocalActivation 已排除被注销的 key）
            if (window.SmartPilotOcrQuota) window.SmartPilotOcrQuota.setLicensed(true);
        }
        applyCheckoutVisibility();
        return result;
    }

    async function registerOrLogin(mode) {
        var login = String((document.getElementById('license-login') || {}).value || '').trim();
        var password = String((document.getElementById('license-password') || {}).value || '');
        if (!login || !password) {
            notify(t('license.err.credentials'), 'warning');
            return;
        }
        if (password.length < 8) {
            notify(t('license.err.weakPassword'), 'warning');
            return;
        }
        var serverOk = await pingLicenseServer();
        if (!serverOk) {
            notify(t('license.err.serverDown'), 'error');
            return;
        }
        var body = { password: password };
        if (login.indexOf('@') >= 0) body.email = login;
        else body.phone = login;
        body.login = login;
        var method = mode === 'register' ? 'LicenseRegister' : 'LicenseLogin';
        var res = await call(method, JSON.stringify(body));
        if (!res || !res.success) {
            notify(friendlyAuthError(res), 'error');
            return;
        }
        notify(t('license.ok.auth'), 'success');
        await refreshAccountCard();
        await pingLicenseServer();
    }

    async function logout() {
        gateCache = null;
        await call('LicenseLogout');
        await refreshAccountCard();
        notify(t('license.ok.logout'), 'info');
    }

    /** 离线激活码激活（路线①：.NET 侧本地验签，无需授权服务器）。 */
    async function activateKey() {
        var input = document.getElementById('license-key-input');
        var key = String((input || {}).value || '').trim();
        if (!key) {
            notify(t('license.key.empty'), 'warning');
            return;
        }
        var res = await call('LicenseActivateKey', JSON.stringify({ key: key }));
        if (!res || !res.success) {
            notify(friendlyAuthError(res) || t('license.key.err'), 'error');
            return;
        }
        if (input) input.value = '';
        gateCache = null;
        notify(t('license.key.ok'), 'success');
        await refreshAccountCard();
    }

    async function clearLocalActivation() {
        var res = await call('LicenseClearLocalActivation');
        if (!res || !res.success) {
            notify(friendlyAuthError(res) || t('license.key.err'), 'error');
            return;
        }
        gateCache = null;
        notify(t('license.key.cleared'), 'info');
        await refreshAccountCard();
    }

    /**
     * 注销本机激活以申请退款（消 key 闭环的客户端一环）。
     * 注销后本机立即失效，且该激活码在本机不能再被激活；返回「消 key 码」，
     * 用户把它发给客服，运营台验签通过后才退款。
     */
    async function deactivateForRefund() {
        // confirm 在个别宿主里可能不可用，缺失时不拦截（用户已明确点了注销按钮）
        var confirmed = true;
        try {
            if (typeof window.confirm === 'function') {
                confirmed = window.confirm(t('license.refund.confirm'));
            }
        } catch (_) { confirmed = true; }
        if (!confirmed) return;

        var res = await call('LicenseDeactivateForRefund');
        if (!res || !res.success) {
            notify(res && res.code === 'not_activated'
                ? t('license.refund.none')
                : (friendlyAuthError(res) || t('license.refund.failed')), 'error');
            return;
        }
        gateCache = null;
        var input = document.getElementById('license-refund-code');
        var box = document.getElementById('license-refund-box');
        if (input) input.value = String(res.code || '');
        if (box) box.style.display = '';
        notify(t('license.refund.done'), 'success');
        await refreshAccountCard();
    }

    /**
     * 把"是否已激活"告诉 OCR 免费额度模块（见 js/ocr-quota.js）。
     * 不依赖任何账户面板 DOM，所以在 OCR 工作台这种独立页面同样生效。
     */
    async function syncOcrQuota() {
        if (!window.SmartPilotOcrQuota) return;
        if (softBypassEnabled()) {
            window.SmartPilotOcrQuota.setLicensed(true);
            return;
        }
        // 先查本机离线激活（秒回，不联网），命中就结束——避免没人服务器时白等在线探测
        try {
            var local = await call('LicenseGetLocalActivation');
            if (local && local.active === true && local.code !== 'bridge_failed') {
                window.SmartPilotOcrQuota.setLicensed(true);
                return;
            }
        } catch (_) { /* ignore */ }
        try {
            var sub = await call('LicenseIsSubscriptionActive');
            if (sub && sub.success !== false && sub.active === true) {
                window.SmartPilotOcrQuota.setLicensed(true);
            }
        } catch (_) { /* ignore */ }
    }

    /** 复制消 key 码（clipboard 在非安全上下文不可用，退回 execCommand）。 */
    function copyRefundCode() {
        var input = document.getElementById('license-refund-code');
        var text = String((input || {}).value || '').trim();
        if (!text) return;
        var done = function () { notify(t('license.refund.copied'), 'success'); };
        var fallback = function () {
            try {
                if (input) { input.select(); document.execCommand('copy'); }
                done();
            } catch (_) { notify(t('license.buy.copyFailed'), 'warning'); }
        };
        if (navigator.clipboard && navigator.clipboard.writeText) {
            try { navigator.clipboard.writeText(text).then(done, fallback); }
            catch (_) { fallback(); }
        } else {
            fallback();
        }
    }

    /** 已激活用户点「输入新激活码」：展开激活入口（续费用，激活后会自动收起）。 */
    function revealKeyInput() {
        var row = document.getElementById('license-key-input-row');
        var activateBtn = document.getElementById('license-activate-btn');
        var renewBtn = document.getElementById('license-renew-btn');
        if (row) row.style.display = '';
        if (activateBtn) activateBtn.style.display = '';
        if (renewBtn) renewBtn.style.display = 'none';
        var input = document.getElementById('license-key-input');
        if (input) { try { input.focus(); } catch (_) { /* ignore */ } }
    }

    /** 渲染本地激活状态行；返回是否处于离线激活生效状态。 */
    function renderLocalActivation(local) {
        var el = document.getElementById('license-local-status');
        var clearBtn = document.getElementById('license-clear-key-btn');
        var active = !!(local && local.active === true && local.code !== 'bridge_failed');
        if (el) {
            if (active) {
                // 主状态行已显示同一份权益（"已激活（激活码）：…"），这里再重复一遍只会显得乱
                el.style.display = 'none';
            } else {
                el.textContent = t('license.local.none');
                el.className = 'license-local-status is-none';
                el.style.display = '';
            }
        }
        if (clearBtn) clearBtn.style.display = active ? '' : 'none';
        // 已激活时收起激活入口（输入框 + 激活按钮），只留「输入新激活码」揭示按钮——
        // 否则会摆一个绿色大按钮邀请已激活用户再做一次激活，纯属误导。续费路径走揭示按钮。
        var inputRow = document.getElementById('license-key-input-row');
        var activateBtn = document.getElementById('license-activate-btn');
        var renewBtn = document.getElementById('license-renew-btn');
        if (inputRow) inputRow.style.display = active ? 'none' : '';
        if (activateBtn) activateBtn.style.display = active ? 'none' : '';
        if (renewBtn) renewBtn.style.display = active ? '' : 'none';
        // 只有已激活时才允许「注销并申请退款」；未激活时收起上一次的消 key 码
        var refundBtn = document.getElementById('license-refund-btn');
        if (refundBtn) refundBtn.style.display = active ? '' : 'none';
        var refundBox = document.getElementById('license-refund-box');
        if (refundBox && !active) refundBox.style.display = 'none';
        // 已激活用户看到的是「续费」而不是「购买」
        var buyTitle = document.getElementById('license-buy-title');
        if (buyTitle) buyTitle.textContent = t(active ? 'license.buy.titleRenew' : 'license.buy.title');
        // 激活码区提示按状态切换：未激活讲"怎么激活"，已激活讲"怎么续费"
        var hint = document.getElementById('license-key-hint');
        var hintRenew = document.getElementById('license-key-hint-renew');
        if (hint) hint.style.display = active ? 'none' : '';
        if (hintRenew) hintRenew.style.display = active ? '' : 'none';
        return active;
    }

    async function checkout(channel, plan) {
        if (channel === 'cn' && !checkoutReady.cn) {
            notify(t('license.checkout.pending'), 'info');
            return;
        }
        if (channel !== 'cn' && !checkoutReady.stripe) {
            notify(t('license.checkout.pending'), 'info');
            return;
        }
        var res = await call('LicenseStartCheckout', JSON.stringify({
            channel: channel || 'stripe',
            plan: plan || (channel === 'cn' ? 'pro_yearly' : 'pro_monthly')
        }));
        if (!res.success) {
            notify(friendlyAuthError(res) || t('license.err.checkout'), 'error');
            return;
        }
        notify(t('license.ok.checkout'), 'info');
    }

    async function openPortal() {
        var res = await call('LicenseOpenBillingPortal');
        if (!res.success) notify(friendlyAuthError(res) || t('license.err.portal'), 'error');
    }

    function ensureDrawerStyles() {
        if (document.getElementById('account-license-drawer-style')) return;
        var style = document.createElement('style');
        style.id = 'account-license-drawer-style';
        style.textContent = [
            '.account-license-drawer{position:fixed;inset:0;z-index:100060;display:none;pointer-events:none;font-family:"Segoe UI",system-ui,sans-serif;}',
            '.account-license-drawer.open{display:block;pointer-events:auto;}',
            '.account-license-drawer-backdrop{position:absolute;inset:0;background:rgba(15,23,42,.45);}',
            '.account-license-drawer-panel{position:absolute;top:0;right:0;bottom:0;width:min(420px,100%);background:#f8fafc;box-shadow:-8px 0 28px rgba(0,0,0,.18);display:flex;flex-direction:column;}',
            '.account-license-drawer-header{flex:0 0 auto;display:flex;align-items:center;justify-content:space-between;padding:12px 16px;background:#0f766e;color:#fff;font-size:15px;font-weight:600;}',
            '.account-license-drawer-close{border:none;background:rgba(255,255,255,.2);color:#fff;width:28px;height:28px;border-radius:50%;cursor:pointer;font-size:16px;line-height:1;}',
            '.account-license-drawer-body{flex:1;overflow:auto;padding:16px;color:#334155;line-height:1.5;}',
            '.account-license-drawer .license-byok-hint{margin:0 0 10px;font-size:12px;color:#64748b;}',
            '.account-license-drawer .license-scope-chip{display:inline-block;margin:0 0 12px;padding:4px 10px;border-radius:999px;background:#ccfbf1;color:#0f766e;font-size:11px;font-weight:600;}',
            '.account-license-drawer .license-server-status{margin:0 0 10px;font-size:12px;font-weight:600;}',
            '.account-license-drawer .license-server-status.is-ok{color:#047857;}',
            '.account-license-drawer .license-server-status.is-down{color:#b91c1c;}',
            '.account-license-drawer .license-server-status.is-neutral{color:#64748b;}',
            '.account-license-drawer .license-status{margin:0 0 14px;font-weight:700;font-size:14px;color:#0f172a;}',
            '.account-license-drawer .license-row{margin-bottom:10px;}',
            '.account-license-drawer .license-row label{display:block;margin-bottom:4px;font-size:12px;color:#475569;}',
            '.account-license-drawer .license-input{display:block;width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:13px;background:#fff;}',
            '.account-license-drawer .license-password-hint{margin:-2px 0 8px;font-size:11px;color:#94a3b8;}',
            '.account-license-drawer .license-btn-row{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px;}',
            '.account-license-drawer .license-btn{border:1px solid #cbd5e1;background:#fff;color:#334155;border-radius:6px;padding:7px 12px;font-size:12px;cursor:pointer;}',
            '.account-license-drawer .license-btn.primary{background:#0f766e;border-color:#0f766e;color:#fff;}',
            '.account-license-drawer .license-actions{display:flex;flex-direction:column;gap:8px;}',
            '.account-license-drawer .license-key{margin:0 0 14px;padding:10px;border:1px dashed #cbd5e1;border-radius:8px;background:#fff;}',
            '.account-license-drawer .license-key-title{font-size:12px;font-weight:600;color:#475569;margin-bottom:8px;}',
            '.account-license-drawer .license-local-status{margin-top:8px;font-size:12px;font-weight:600;color:#0f766e;}',
            '.account-license-drawer .license-local-status.is-none{color:#94a3b8;font-weight:400;}',
            '.account-license-drawer .license-refund-box{margin-top:8px;padding:10px;border:1px solid #e2e8f0;border-radius:8px;background:#f8fafc;}',
            '.account-license-drawer .license-buy{margin:0 0 14px;padding:10px;border:1px solid #bbf7d0;border-radius:8px;background:#f0fdf4;}',
            '.account-license-drawer .license-buy-price{font-size:13px;font-weight:700;color:#166534;margin-bottom:8px;}',
            '.account-license-drawer .license-buy-plans{margin:0 0 8px;padding-left:18px;font-size:12px;color:#334155;}',
            '.account-license-drawer .license-buy-plans li{margin-bottom:4px;}',
            '.account-license-drawer .license-buy-steps{margin:0 0 10px;padding-left:18px;font-size:12px;color:#334155;}',
            '.account-license-drawer .license-buy-steps li{margin-bottom:4px;}',
            '.account-license-drawer .license-buy-qr-row{text-align:center;margin-bottom:8px;}',
            '.account-license-drawer .license-buy-qr{max-width:180px;width:100%;border:1px solid #cbd5e1;border-radius:8px;background:#fff;}',
            '.account-license-drawer .license-buy-contact-row{display:flex;align-items:center;justify-content:space-between;gap:8px;font-size:12px;color:#334155;}',
            '.account-license-drawer .license-buy-contact-label{color:#475569;flex:0 0 auto;}',
            '.account-license-drawer .license-buy-contact{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}'
        ].join('');
        (document.head || document.documentElement).appendChild(style);
    }

    function ensureAccountDrawer() {
        ensureDrawerStyles();
        var existing = document.getElementById('account-license-drawer');
        if (existing) return existing;
        var overlay = document.createElement('div');
        overlay.id = 'account-license-drawer';
        overlay.className = 'account-license-drawer';
        overlay.setAttribute('aria-hidden', 'true');
        overlay.innerHTML =
            '<div class="account-license-drawer-backdrop" data-lic-close="1"></div>' +
            '<div class="account-license-drawer-panel" role="dialog" aria-modal="true" aria-label="' + t('nav.account') + '">' +
            '  <div class="account-license-drawer-header">' +
            '    <span data-i18n="nav.account">' + t('nav.account') + '</span>' +
            '    <button type="button" class="account-license-drawer-close" data-lic-close="1" title="Esc">×</button>' +
            '  </div>' +
            '  <div class="account-license-drawer-body" id="license-card">' + licenseCardInnerHtml({ includeTitle: false }) + '</div>' +
            '</div>';
        document.body.appendChild(overlay);
        overlay.addEventListener('click', function (e) {
            var el = e.target;
            if (el && (el.getAttribute('data-lic-close') === '1' || el.classList.contains('account-license-drawer-backdrop'))) {
                closeAccountPanel();
            }
        });
        bindLicenseCardClicks(overlay.querySelector('#license-card'));
        if (!overlay._escBound) {
            overlay._escBound = true;
            document.addEventListener('keydown', function (e) {
                if (e.key === 'Escape' && overlay.classList.contains('open')) closeAccountPanel();
            });
        }
        return overlay;
    }

    function softBypassEnabled() {
        try {
            if (window.__LICENSE_SOFT_BYPASS__ === true) return true;
            if (localStorage.getItem('LICENSE_SOFT_BYPASS') === '1') return true;
        } catch (_) { /* ignore */ }
        return false;
    }

    function openAccountPanel() {
        // 主界面 iframe 内嵌设置：交给父页打开一级账户抽屉
        try {
            if (window.parent && window.parent !== window) {
                try {
                    if (typeof window.parent.openAccountPanel === 'function' &&
                        window.parent.document && window.parent.document.getElementById('side-toolbar')) {
                        window.parent.openAccountPanel();
                        return;
                    }
                } catch (_) { /* cross-origin or unavailable */ }
                try {
                    window.parent.postMessage({ type: 'smartpilot:open-account' }, '*');
                    if (/[?&]embed=1(?:&|$)/.test(String(location.search || ''))) return;
                } catch (_) { /* fall through to local drawer */ }
            }
        } catch (_) { /* fall through */ }
        var drawer = ensureAccountDrawer();
        drawer.classList.add('open');
        drawer.setAttribute('aria-hidden', 'false');
        refreshAccountCard()
            .then(function () { return pingLicenseServer(); })
            .catch(function (e) {
                ChatLog.warn?.('LICENSE', e && e.message);
            });
    }

    function closeAccountPanel() {
        var drawer = document.getElementById('account-license-drawer');
        if (!drawer) return;
        // 先把焦点移出抽屉再隐藏：aria-hidden="true" 的子树里若留有焦点，浏览器会
        // 拦截并告警（Blocked aria-hidden ... descendant retained focus），键盘用户
        // 的焦点位置也会丢进隐藏节点。侧栏按钮是 div（无 tabindex，鼠标导向设计），
        // 不擅自给工具栏加键盘模型，就地失焦让焦点回到 body。
        try {
            var active = document.activeElement;
            if (active && drawer.contains(active) && typeof active.blur === 'function') {
                active.blur();
            }
        } catch (_) { /* ignore */ }
        drawer.classList.remove('open');
        drawer.setAttribute('aria-hidden', 'true');
    }

    /**
     * 席位门禁。对话和云端 OCR 需要有效试用或订阅。
     * 本机识图在未登录、且付费通道还没开通时放行。
     * 已登录但试用过期或订阅失效时，本机识别同样拦截。
     * 本地开发跳过：localStorage.LICENSE_SOFT_BYPASS='1' 或 window.__LICENSE_SOFT_BYPASS__=true
     * @param {{ surface?: string }} [opts] surface='local-ocr' 表示本机引擎
     * @returns {Promise<boolean>}
     */
    // 发送按钮每次都等 LicenseIsSubscriptionActive。桥实现会打一次授权 HTTP，
    // 超时 20 秒，授权服务没开时点击就像没反应。打开面板时已经拉过权益，这里复用一会儿。
    var GATE_TTL_MS = 60000;
    var gateCache = null;

    function rememberGateResult(res) {
        if (!res || res.code === 'bridge_failed') {
            gateCache = null;
            return;
        }
        gateCache = { at: Date.now(), res: res };
    }

    function entitlementToGate(result) {
        if (!result || result.code === 'bridge_failed') return null;
        if (result.code === 'no_bridge') return result;
        var ent = result.entitlement || null;
        if (result.active !== undefined && result.loggedIn !== undefined && !ent) return result;
        return {
            success: result.success !== false,
            active: !!(ent && (ent.active === true || result.active === true)),
            loggedIn: result.loggedIn === true,
            entitlement: ent,
            error: result.error || (ent && ent.reason) || null,
            code: result.code
        };
    }

    /** @returns {boolean|null} null = 交给调用方走本机识别的额外探测 */
    function decideGate(res, opts) {
        opts = opts || {};
        if (res && res.code === 'no_bridge') return true;
        if (res && res.active) return true;
        if (opts.surface === 'local-ocr' && res && res.loggedIn === false && res.code !== 'bridge_failed') {
            return null;
        }
        var loggedIn = isLoggedInEntitlement(res);
        var msg = loggedIn ? t('license.gate.blocked') : t('license.gate.needAccount');
        try {
            if (typeof App.showToast === 'function') App.showToast(msg, 'warning');
            else if (typeof window.showToast === 'function') window.showToast(msg, 'warning');
        } catch (_) { /* ignore */ }
        if (document.getElementById('side-account-btn')) {
            try { setTimeout(function () { openAccountPanel(); }, 80); } catch (_) { /* ignore */ }
        }
        return false;
    }

    async function ensureSubscriptionActive(opts) {
        opts = opts || {};
        if (softBypassEnabled()) return true;

        // 离线激活优先：激活码在 .NET 侧本地验签，授权服务不在线也能用（路线①）
        var localGate = await call('LicenseGetLocalActivation');
        if (localGate && localGate.code !== 'bridge_failed' && localGate.active === true) return true;

        if (!opts.force && gateCache && (Date.now() - gateCache.at) < GATE_TTL_MS) {
            var cachedDecision = decideGate(gateCache.res, opts);
            if (cachedDecision !== null) return cachedDecision;
        }

        var res = await call('LicenseIsSubscriptionActive');
        rememberGateResult(res);
        var decision = decideGate(res, opts);
        if (decision !== null) return decision;

        // 只有明确未登录才放行本机识别。调用失败或结果里没有 loggedIn 时不放行。
        if (opts.surface === 'local-ocr' && res && res.loggedIn === false && res.code !== 'bridge_failed') {
            await pingLicenseServer();
            if (!checkoutReady.stripe && !checkoutReady.cn) return true;
        }

        var loggedIn = isLoggedInEntitlement(res);
        var msg = loggedIn ? t('license.gate.blocked') : t('license.gate.needAccount');
        try {
            if (typeof App.showToast === 'function') App.showToast(msg, 'warning');
            else if (typeof window.showToast === 'function') window.showToast(msg, 'warning');
        } catch (_) { /* ignore */ }
        // 登录只挂在主界面侧栏。OCR 等独立窗格只提示，不再另开一套表单。
        if (document.getElementById('side-account-btn')) {
            try { setTimeout(function () { openAccountPanel(); }, 80); } catch (_) { /* ignore */ }
        }
        return false;
    }

    function isLoggedInEntitlement(res) {
        return !!(res && res.loggedIn === true);
    }

    window.SmartPilotLicense = {
        refreshAccountCard: refreshAccountCard,
        register: function () { return registerOrLogin('register'); },
        login: function () { return registerOrLogin('login'); },
        logout: logout,
        checkoutStripeMonthly: function () { return checkout('stripe', 'pro_monthly'); },
        checkoutStripeYearly: function () { return checkout('stripe', 'pro_yearly'); },
        checkoutCnYearly: function () { return checkout('cn', 'pro_yearly'); },
        openPortal: openPortal,
        activateKey: activateKey,
        clearLocalActivation: clearLocalActivation,
        getLocalActivation: function () { return call('LicenseGetLocalActivation'); },
        openAccountPanel: openAccountPanel,
        closeAccountPanel: closeAccountPanel,
        ensureSubscriptionActive: ensureSubscriptionActive,
        licenseCardInnerHtml: licenseCardInnerHtml,
        bindLicenseCardClicks: bindLicenseCardClicks
    };

    window.openAccountPanel = openAccountPanel;
    window.closeAccountPanel = closeAccountPanel;
    window.ensureSubscriptionActive = ensureSubscriptionActive;

    document.addEventListener('DOMContentLoaded', function () {
        var staticCard = document.getElementById('license-card');
        if (staticCard) {
            if (!staticCard.querySelector('#license-status')) {
                staticCard.innerHTML = licenseCardInnerHtml();
            }
            bindLicenseCardClicks(staticCard);
        }
        // 主界面 / OCR 页：刷新侧栏状态或账户卡（未登录时按钮呈灰色）
        if (staticCard || document.getElementById('side-account-btn')) {
            refreshAccountCard()
                .then(function () { return pingLicenseServer(); })
                .catch(function (e) {
                    ChatLog.warn?.('LICENSE', e && e.message);
                });
        }
        // OCR 工作台等页面没有账户面板 DOM，但也要让 OCR 额度知道是否已激活，
        // 否则已付费用户在那里会被当成未激活、白白限额。
        syncOcrQuota().catch(function () { /* ignore */ });
    });
})();
