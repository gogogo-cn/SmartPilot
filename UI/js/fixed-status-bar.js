/**
 * ============================================================
 * fixed-status-bar.js — Fixed Status Bar
 * ============================================================
 * 功能：固定状态栏模块，在页面底部显示全局状态、进度与快捷操作按钮。
 *
 * 核心类/函数：
 *   - showStatus() — 显示状态
 *   - hideStatus() — 隐藏状态
 *   - setProgress() — 设置进度
 *   - showActionButtons() — 显示操作按钮
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    ChatLog.info('STATUS', '[FixedStatusBar] 📦 脚本开始加载...');

    class FixedStatusBar {
        constructor() {
            this.element = null;
            this.textEl = null;
            this.progressEl = null;
            this.progressFill = null;
            this.percentageEl = null;

            this.thinProgressContainer = null;
            this.thinProgressBar = null;

            this.currentStatus = 'idle';
            this.isInitialized = false;
            this.extraStatusText = '';  // 额外状态文本（如思考模式）
            this.mainMessage = (window.App && App.i18n && App.i18n.t) ? App.i18n.t('status.ready') : '系统就绪';  // 主消息（单独存储，不从DOM提取）

            ChatLog.info('STATUS', '[FixedStatusBar] ✅ 构造函数完成');
        }

        init() {
            ChatLog.info('STATUS', '[FixedStatusBar] 🚀 开始初始化...');

            if (this.isInitialized) {
                ChatLog.warn('STATUS', '[FixedStatusBar] ⚠️ 已初始化');
                return;
            }

            this.element = document.getElementById('fixed-status-bar');
            if (!this.element) {
                ChatLog.error('STATUS', '[FixedStatusBar] ❌ 找不到元素');
                alert('错误：找不到 #fixed-status-bar');
                return;
            }

            ChatLog.info('STATUS', '[FixedStatusBar] ✅ 找到元素');

            this.textEl = document.getElementById('fsb-text');
            this.progressEl = document.getElementById('fsb-progress');
            this.progressFill = document.getElementById('fsb-progress-fill');
            this.percentageEl = document.getElementById('fsb-percentage');

            this.createThinProgressBar();
            this.injectThinProgressStyles();

            this.isInitialized = true;
            ChatLog.info('STATUS', '[FixedStatusBar] 🎉 初始化完成！');

            this.setIdle();
        }

        createThinProgressBar() {
            if (this.thinProgressContainer) return;

            this.thinProgressContainer = document.getElementById('fsb-thin-progress');

            if (this.thinProgressContainer) {
                this.thinProgressBar = document.getElementById('fsb-thin-progress-bar');
                ChatLog.info('STATUS', '[FixedStatusBar] ✅ 找到HTML中的细进度条');

                // ★ 强制设置样式确保可见
                this.thinProgressContainer.style.cssText = 'position:absolute;bottom:0;left:0;width:100%;height:3px;background:rgba(255,0,0,0.2);overflow:hidden;z-index:99999;display:block;';
                
                if (this.thinProgressBar) {
                    this.thinProgressBar.style.cssText = 'height:100%;width:50%;background:linear-gradient(90deg,#6366f1,#8b5cf6,#a78bfa);box-shadow:0 0 15px rgba(99,102,241,1);transition:width 0.3s;';
                }
                return;
            }

            ChatLog.info('STATUS', '[FixedStatusBar] 📝 动态创建细进度条...');
            
            this.thinProgressContainer = document.createElement('div');
            this.thinProgressContainer.id = 'fsb-thin-progress';
            this.thinProgressContainer.style.cssText = 'position:absolute;bottom:0;left:0;width:100%;height:3px;background:rgba(255,0,0,0.2);z-index:99999;display:block;';

            this.thinProgressBar = document.createElement('div');
            this.thinProgressBar.id = 'fsb-thin-progress-bar';
            this.thinProgressBar.style.cssText = 'height:100%;width:50%;background:linear-gradient(90deg,#6366f1,#8b5cf6,#a78bfa);box-shadow:0 0 15px rgba(99,102,241,1);';

            this.thinProgressContainer.appendChild(this.thinProgressBar);
            this.element.appendChild(this.thinProgressContainer);
        }

        injectThinProgressStyles() {
            if (document.getElementById('fsb-styles')) return;

            const style = document.createElement('style');
            style.id = 'fsb-styles';
            style.textContent = `
                .fsb-thin-progress{position:absolute!important;bottom:0!important;left:0!important;width:100%!important;height:3px!important;z-index:99999!important;display:block!important}
                .fsb-thin-progress-bar{height:100%!important;width:50%!important;background:linear-gradient(90deg,#6366f1,#8b5cf6,#a78bfa)!important;box-shadow:0 0 15px rgba(99,102,241,1)!important;transition:width 0.3s ease-out!important}
                @keyframes fsb-indeterminate{0%{transform:translateX(-100%)}100%{transform:translateX(400%)}}
                .fsb-thin-progress.indeterminate .fsb-thin-progress-bar{width:30%!important;animation:fsb-indeterminate 1.5s ease-in-out infinite!important}
                #fixed-status-bar{position:relative!important}
            `;
            document.head.appendChild(style);
        }

        updateThinProgress(pct = 0, indeterminate = false) {
            if (!this.thinProgressBar || !this.thinProgressContainer) return;
            
            if (indeterminate) {
                this.thinProgressContainer.classList.add('indeterminate');
            } else {
                this.thinProgressContainer.classList.remove('indeterminate');
                this.thinProgressBar.style.width = Math.max(0, Math.min(100, pct)) + '%';
            }
        }

        showThinProgress(indeterminate = false) {
            if (this.thinProgressContainer) {
                this.thinProgressContainer.style.display = 'block';
                this.updateThinProgress(0, indeterminate);
            }
        }

        hideThinProgress() {
            if (this.thinProgressContainer) {
                this.thinProgressContainer.style.display = 'none';
                this.updateThinProgress(0);
            }
        }

        setIdle(message) {
            if (message == null || message === '') {
                message = (window.App && App.i18n && App.i18n.t) ? App.i18n.t('status.ready') : '系统就绪';
            }
            // 追加版本戳到就绪消息（自动跟随 __APP_VERSION__，无需手动更新）
            const versionStamp = ' ' + (window.__APP_VERSION__ || (window.App && App.SCRIPT_VERSION) || 'v0');
            const msgWithVersion = message + '<span style="color:#999;font-size:11px;margin-left:6px;">' + versionStamp + '</span>';
            this.setStatus('idle', msgWithVersion);
            this.hideProgress();

            // ★ 就绪状态：显示色弱友好-蓝色满条（100%）作为视觉反馈
            if (this.thinProgressBar && this.thinProgressContainer) {
                this.thinProgressContainer.classList.remove('indeterminate');

                // 使用 !important 确保覆盖CSS默认值
                this.thinProgressBar.style.setProperty('width', '100%', 'important');
                this.thinProgressBar.style.setProperty('transition', 'width 0.5s ease-out', 'important');

                // 色弱友好：深蓝色渐变（高对比度 + 蓝色对大多数色盲友好）
                this.thinProgressBar.style.setProperty('background', 'linear-gradient(90deg, #1e40af 0%, #3b82f6 50%, #60a5fa 100%)', 'important');

                // 蓝色发光效果
                this.thinProgressBar.style.setProperty('boxShadow', '0 0 12px rgba(59,130,246,0.8)', 'important');

                // 停止微光动画
                this.thinProgressBar.style.setProperty('animation', 'none', 'important');

                this.thinProgressContainer.style.setProperty('display', 'block', 'important');

                ChatLog.info('STATUS', '[FixedStatusBar] ✅ 系统就绪 - 显示色弱友好-蓝色满条');
            }
        }
        
        setLoading(msg, pct = 0) { this.setStatus('loading', msg); this.showProgress(pct); this.showThinProgress(false); this.updateThinProgress(pct); }
        
        setProcessing(msg) { this.setStatus('processing', msg); this.hideProgress(); this.showThinProgress(true); }
        
        setSuccess(msg, dur = 2000) {
            this.setStatus('success', msg);
            this.hideProgress();
            if (this.thinProgressBar && this.thinProgressContainer) {
                this.thinProgressContainer.classList.remove('indeterminate');
                this.thinProgressBar.style.transition = 'width 0.5s';
                this.thinProgressBar.style.width = '100%';
                setTimeout(() => { this.hideThinProgress(); if(this.thinProgressBar) this.thinProgressBar.style.transition = 'width 0.3s'; }, 500);
            }
            if (dur > 0) setTimeout(() => { if(this.currentStatus === 'success') this.setIdle(); }, dur);
        }
        
        setError(msg, dur = 3000) { this.setStatus('error', msg); this.hideProgress(); this.hideThinProgress(); if(dur > 0) setTimeout(() => { if(this.currentStatus === 'error') this.setIdle(); }, dur); }

        setInfo(msg, dur = 2500) {
            this.setStatus('idle', msg);
            this.hideProgress();
            this.hideThinProgress();
            if (dur > 0) {
                if (this._infoTimer) clearTimeout(this._infoTimer);
                this._infoTimer = setTimeout(() => {
                    if (this.currentStatus === 'idle') this.setIdle();
                }, dur);
            }
        }

        setWarning(msg, dur = 2500) { this.setError(msg, dur); }

        updateProgress(pct, msg = null) {
            if (this.progressFill) this.progressFill.style.width = Math.max(0, Math.min(100, pct)) + '%';
            if (this.percentageEl) this.percentageEl.textContent = Math.round(pct) + '%';
            if (msg && this.textEl) this.textEl.textContent = msg;
            this.updateThinProgress(pct);
        }

        setStatus(status, msg) {
            this.currentStatus = status;
            this.mainMessage = msg || '系统就绪';  // 存储主消息
            if (!this.element || !this.textEl) return;
            this.element.className = '';

            var iconMap = {
                idle: '<span style="color:#3b82f6;font-size:16px;display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;margin-right:4px;vertical-align:text-bottom;">●</span>',
                loading: '',
                processing: '',
                success: '',
                error: ''
            };

            var icon = iconMap[status] || '';
            // 保留额外的状态信息（如思考模式）
            var extraInfo = this.extraStatusText || '';
            // 只更新文字内容，保留 spinner DOM 结构
            var contentEl = this.textEl.querySelector('.fsb-text-content');
            if (contentEl) {
                contentEl.innerHTML = icon + this.mainMessage + extraInfo;
            } else {
                this.textEl.innerHTML = icon + this.mainMessage + extraInfo;
            }

            if (status === 'loading' || status === 'processing') this.element.classList.add('active');
            else if (status === 'success') this.element.classList.add('success');
            else if (status === 'error') this.element.classList.add('error');
        }
        
        // 设置额外状态文本 —— 已禁用：所有状态必须只显示在固定状态条主消息中，
        // 不允许动态追加额外 DOM 内容或第二行信息。
        setExtraStatus(html, autoClear = false) {
            // 忽略所有动态扩展状态，保持状态条单行、固定
            return;
        }

        _refreshStatusText() {
            if (!this.textEl || !this.currentStatus) return;
            const iconMap = {
                idle: '<span style="color:#3b82f6;font-size:16px;display:inline-flex;align-items:center;justify-content:center;width:16px;height:16px;margin-right:4px;vertical-align:text-bottom;">●</span>',
                loading: '',
                processing: '',
                success: '',
                error: ''
            };
            const icon = iconMap[this.currentStatus] || '';
            // 只显示固定主消息，不附加任何动态扩展文本
            this.textEl.innerHTML = icon + this.mainMessage;
        }

        showProgress(pct = 0) { if (this.progressEl) this.progressEl.style.display = 'flex'; this.updateProgress(pct); }
        
        hideProgress() {
            if (this.progressEl) this.progressEl.style.display = 'none';
            if (this.progressFill) this.progressFill.style.width = '0%';
            if (this.percentageEl) this.percentageEl.textContent = '0%';
        }
    }

    window.App.fixedStatusBar = new FixedStatusBar();
    ChatLog.info('STATUS', '[FixedStatusBar] ✅ 实例化完成');
})();