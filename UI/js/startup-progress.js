/**
 * ============================================================
 * startup-progress.js — Startup Progress
 * ============================================================
 * 功能：启动进度模块，在应用初始化时显示加载进度、步骤状态与错误信息。
 *
 * 核心类/函数：
 *   - showStartupProgress() — 显示启动进度
 *   - updateStartupStep() — 更新启动步骤
 *   - hideStartupProgress() — 隐藏启动进度
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

// ChatLog 兜底（本文件在 app-bootstrap.js 之前加载，ChatLog 可能尚未定义）
window.ChatLog = window.ChatLog || console;
(function() {
    'use strict';

    class StartupProgressManager {
        constructor() {
            this.steps = [];
            this.currentStepIndex = -1;
            this.maxProgress = 0;  // 记录历史最大进度，防止回退
            this.progressElement = null;
            this.progressBar = null;
            this.statusText = null;
            this.percentageText = null;
            this.detailText = null;
            this.isInitialized = false;
        }

        /**
         * 初始化进度管理器
         * @param {Array} steps - 步骤配置数组
         */
        init(steps) {
            if (this.isInitialized) {
                ChatLog.warn('STARTUP', '[StartupProgress] 已经初始化，跳过重复初始化');
                return;
            }

            this.steps = steps;

            // 在聊天区域创建进度条容器
            const chatMessages = document.getElementById('chat-messages');
            if (!chatMessages) {
                ChatLog.error('STARTUP', '[StartupProgress] 找不到聊天消息容器');
                return;
            }

            // 创建居中容器
            const centerWrapper = document.createElement('div');
            centerWrapper.className = 'startup-progress-center';
            centerWrapper.id = 'startup-progress-center';

            // 创建美化的进度条DOM
            this.progressElement = document.createElement('div');
            this.progressElement.className = 'startup-progress-card';
            this.progressElement.id = 'startup-progress-message';
            this.progressElement.innerHTML = `
                <div class="sp-card-inner">
                    <div class="sp-header">
                        <div class="sp-icon-wrapper">
                            <span class="sp-icon">🚀</span>
                        </div>
                        <h3 class="sp-title">系统准备中</h3>
                    </div>
                    
                    <div class="sp-progress-container">
                        <div class="sp-progress-track">
                            <div class="sp-progress-fill" id="progress-fill"></div>
                            <div class="sp-progress-glow" id="progress-glow"></div>
                        </div>
                    </div>
                    
                    <div class="sp-status-row">
                        <div class="sp-status" id="progress-status">
                            <div class="sp-spinner-mini"></div>
                            <span>正在加载...</span>
                        </div>
                        <div class="sp-percentage" id="progress-percentage">0%</div>
                    </div>
                    
                    <div class="sp-detail" id="progress-detail">
                        <span class="sp-detail-text" id="progress-detail-text"></span>
                    </div>
                </div>
            `;

            centerWrapper.appendChild(this.progressElement);
            
            // 插入到聊天区域
            chatMessages.insertBefore(centerWrapper, chatMessages.firstChild);

            // 注入样式（如果尚未注入）
            if (!document.getElementById('startup-progress-styles')) {
                this.injectStyles();
            }

            // 缓存DOM引用
            this.progressBar = document.getElementById('progress-fill');
            this.progressGlow = document.getElementById('progress-glow');
            this.statusText = document.getElementById('progress-status');
            this.percentageText = document.getElementById('progress-percentage');
            this.detailText = document.getElementById('progress-detail-text');

            // 重置进度状态
            this.maxProgress = 0;
            this.currentStepIndex = -1;

            this.isInitialized = true;
            ChatLog.info('STARTUP', `[StartupProgress] ✅ 初始化完成（居中美化版），共 ${steps.length} 个步骤`);
        }

        /**
         * 注入CSS样式
         */
        injectStyles() {
            const style = document.createElement('style');
            style.id = 'startup-progress-styles';
            style.textContent = `
                /* 居中容器 */
                .startup-progress-center {
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    min-height: 100%;
                    padding: 40px 20px;
                }

                /* 进度卡片 */
                .startup-progress-card {
                    width: 100%;
                    max-width: 420px;
                    background: linear-gradient(135deg, #ffffff 0%, #f8f9ff 100%);
                    border-radius: 16px;
                    padding: 28px 32px;
                    box-shadow: 
                        0 4px 24px rgba(0, 0, 0, 0.08),
                        0 1px 2px rgba(0, 0, 0, 0.04),
                        inset 0 1px 0 rgba(255, 255, 255, 0.8);
                    border: 1px solid rgba(102, 126, 234, 0.12);
                    transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
                    animation: spFadeInUp 0.6s ease-out;
                }

                @keyframes spFadeInUp {
                    from {
                        opacity: 0;
                        transform: translateY(20px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }

                .sp-card-inner {
                    display: flex;
                    flex-direction: column;
                    gap: 18px;
                }

                /* 头部 */
                .sp-header {
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    gap: 10px;
                }

                .sp-icon-wrapper {
                    width: 52px;
                    height: 52px;
                    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                    border-radius: 14px;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-size: 26px;
                    box-shadow: 0 4px 14px rgba(102, 126, 234, 0.35);
                    transition: all 0.3s ease;
                }

                .sp-title {
                    margin: 0;
                    font-size: 17px;
                    font-weight: 600;
                    color: #1a1a2e;
                    letter-spacing: -0.02em;
                }

                /* 进度条 */
                .sp-progress-container {
                    position: relative;
                }

                .sp-progress-track {
                    height: 8px;
                    background: #e8ecf4;
                    border-radius: 100px;
                    overflow: visible;
                    position: relative;
                }

                .sp-progress-fill {
                    height: 100%;
                    background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
                    border-radius: 100px;
                    width: 0%;
                    transition: width 0.5s cubic-bezier(0.4, 0, 0.2, 1);
                    position: relative;
                    z-index: 2;
                }

                .sp-progress-glow {
                    position: absolute;
                    top: 50%;
                    left: 0;
                    transform: translateY(-50%);
                    width: 60px;
                    height: 24px;
                    background: radial-gradient(ellipse, rgba(102, 126, 234, 0.4), transparent);
                    filter: blur(8px);
                    transition: left 0.5s cubic-bezier(0.4, 0, 0.2, 1);
                    z-index: 1;
                    pointer-events: none;
                }

                /* 状态行 */
                .sp-status-row {
                    display: flex;
                    justify-content: space-between;
                    align-items: center;
                }

                .sp-status {
                    display: flex;
                    align-items: center;
                    gap: 8px;
                    font-size: 13px;
                    color: #555;
                    font-weight: 500;
                }

                .sp-spinner-mini {
                    width: 16px;
                    height: 16px;
                    border: 2px solid #e0e4ec;
                    border-top-color: #667eea;
                    border-radius: 50%;
                    animation: spSpin 0.8s linear infinite;
                }

                @keyframes spSpin {
                    to { transform: rotate(360deg); }
                }

                .sp-percentage {
                    font-size: 13px;
                    font-weight: 600;
                    color: #667eea;
                    min-width: 36px;
                    text-align: right;
                    font-variant-numeric: tabular-nums;
                }

                /* 详情 */
                .sp-detail {
                    background: #f0f4ff;
                    border-radius: 8px;
                    padding: 10px 14px;
                    display: none;
                    animation: spSlideDown 0.25s ease-out;
                }

                @keyframes spSlideDown {
                    from {
                        opacity: 0;
                        transform: translateY(-4px);
                    }
                    to {
                        opacity: 1;
                        transform: translateY(0);
                    }
                }

                .sp-detail-text {
                    font-size: 12px;
                    color: #5a67d8;
                    line-height: 1.5;
                }

                /* 完成状态 */
                .startup-progress-card.completed {
                    background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%);
                    border-color: rgba(34, 197, 94, 0.2);
                }

                .startup-progress-card.completed .sp-icon-wrapper {
                    background: linear-gradient(135deg, #22c55e 0%, #16a34a 100%);
                    box-shadow: 0 4px 14px rgba(34, 197, 94, 0.35);
                }

                .startup-progress-card.completed .sp-progress-fill {
                    background: linear-gradient(90deg, #22c55e 0%, #16a34a 100%);
                }

                .startup-progress-card.completed .sp-spinner-mini {
                    display: none;
                }

                .startup-progress-card.completed .sp-percentage {
                    color: #22c55e;
                }
            `;
            document.head.appendChild(style);
        }

        /**
         * 开始指定步骤
         * @param {string} stepId - 步骤ID
         */
        startStep(stepId) {
            const stepIndex = this.steps.findIndex(s => s.id === stepId);
            if (stepIndex === -1) {
                ChatLog.warn('STARTUP', `[StartupProgress] 未找到步骤: ${stepId}`);
                return;
            }

            this.currentStepIndex = stepIndex;
            this.updateUI();

            ChatLog.info('STARTUP', `[StartupProgress] ▶️ [${stepIndex + 1}/${this.steps.length}] ${this.steps[stepIndex].name}`);
        }

        /**
         * 完成当前步骤
         * @param {string} stepId - 步骤ID（可选）
         * @param {string} detail - 详细信息（可选）
         */
        completeStep(stepId, detail) {
            let stepIndex;

            if (stepId) {
                stepIndex = this.steps.findIndex(s => s.id === stepId);
            } else {
                stepIndex = this.currentStepIndex;
            }

            if (stepIndex === -1 || stepIndex >= this.steps.length) {
                return;
            }

            // 显示详细信息
            if (detail && this.detailText) {
                const detailContainer = document.getElementById('progress-detail');
                if (detailContainer) {
                    detailContainer.style.display = 'block';
                    this.detailText.textContent = detail;
                }
                ChatLog.info('STARTUP', `[StartupProgress] ✅ [${stepIndex + 1}/${this.steps.length}] ${this.steps[stepIndex].name} - ${detail}`);
            } else {
                ChatLog.info('STARTUP', `[StartupProgress] ✅ [${stepIndex + 1}/${this.steps.length}] ${this.steps[stepIndex].name}`);
            }

            // 步骤推进由调用方显式控制，避免与 js-app-init.js 的 startStep 重复触发
        }

        /**
         * 设置详细信息
         * @param {string} detail - 详细信息文本
         */
        setDetail(detail) {
            if (detail && this.detailText) {
                const detailContainer = document.getElementById('progress-detail');
                if (detailContainer) {
                    detailContainer.style.display = 'block';
                    this.detailText.textContent = detail;
                }
            }
        }

        /**
         * 更新UI显示（进度只增不减，防止回退跳动）
         */
        updateUI() {
            if (!this.progressBar || !this.statusText || !this.percentageText) return;

            // 计算当前步骤对应的进度
            const completedSteps = Math.max(0, this.currentStepIndex);
            const totalSteps = this.steps.length;
            let progress = ((completedSteps + 0.5) / totalSteps) * 100;

            // 🚀 【关键修复】取历史最大值，确保进度条单调递增
            this.maxProgress = Math.max(this.maxProgress, progress);
            progress = this.maxProgress;

            // 更新进度条和光晕位置
            const widthPercent = `${Math.min(progress, 100)}%`;
            this.progressBar.style.width = widthPercent;
            if (this.progressGlow) {
                this.progressGlow.style.left = `calc(${widthPercent} - 30px)`;
            }

            // 更新百分比
            this.percentageText.textContent = `${Math.round(progress)}%`;

            // 更新状态文字（始终显示当前正在执行的步骤）
            if (this.currentStepIndex >= 0 && this.currentStepIndex < this.steps.length) {
                const currentStep = this.steps[this.currentStepIndex];
                this.statusText.innerHTML = `
                    <div class="sp-spinner-mini"></div>
                    <span>${currentStep.name}</span>
                `;

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setLoading(currentStep.name, progress);
                }
            }
        }

        /**
         * 标记完成并隐藏进度条（带动画）
         */
        async hide() {
            return new Promise((resolve) => {
                if (!this.progressElement) {
                    resolve();
                    return;
                }

                // 显示100%和完成状态
                if (this.progressBar) {
                    this.progressBar.style.width = '100%';
                }
                if (this.progressGlow) {
                    this.progressGlow.style.left = 'calc(100% - 30px)';
                }
                if (this.percentageText) {
                    this.percentageText.textContent = '100%';
                }
                if (this.statusText) {
                    this.statusText.innerHTML = '<span>✅ 准备就绪</span>';
                }

                // 同步更新固定状态栏
                if (window.App && App.fixedStatusBar && App.fixedStatusBar.isInitialized) {
                    App.fixedStatusBar.setSuccess('系统准备完成', 2000);
                }

                // 添加完成样式类
                this.progressElement.classList.add('completed');

                // 更新标题和图标
                const title = this.progressElement.querySelector('.sp-title');
                if (title) {
                    title.textContent = '系统已就绪';
                }

                const icon = this.progressElement.querySelector('.sp-icon');
                if (icon) {
                    icon.textContent = '✨';
                }

                // 延迟后淡出移除
                setTimeout(() => {
                    if (this.progressElement && this.progressElement.parentNode) {
                        const wrapper = this.progressElement.closest('#startup-progress-center');
                        if (wrapper) {
                            wrapper.style.opacity = '0';
                            wrapper.style.transform = 'scale(0.95)';
                            wrapper.style.transition = 'all 0.5s cubic-bezier(0.4, 0, 0.2, 1)';
                        }

                        setTimeout(() => {
                            if (wrapper && wrapper.parentNode) {
                                wrapper.remove();
                            } else if (this.progressElement && this.progressElement.parentNode) {
                                this.progressElement.remove();
                            }
                            ChatLog.info('STARTUP', '[StartupProgress] 进度条已移除');
                            resolve();
                        }, 500);
                    } else {
                        resolve();
                    }
                }, 1200);
            });
        }

        /**
         * 显示进度条（如果被隐藏）
         */
        show() {
            const wrapper = document.querySelector('#startup-progress-center');
            if (wrapper) {
                wrapper.style.display = 'flex';
                wrapper.style.opacity = '1';
                wrapper.style.transform = 'scale(1)';
            }
        }
    }

    // 创建全局单例实例
    window.startupProgress = new StartupProgressManager();

    ChatLog.info('STARTUP', '✅ StartupProgressManager 已创建（居中美化版）');

})();