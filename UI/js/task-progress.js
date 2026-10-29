/**
 * ============================================================
 * task-progress.js — Task Progress
 * ============================================================
 * 功能：任务进度模块，显示长时间运行任务的进度条、步骤详情与取消控制。
 *
 * 核心类/函数：
 *   - showTaskProgress() — 显示任务进度
 *   - updateProgress() — 更新进度
 *   - completeTaskProgress() — 完成任务进度
 *   - cancelTaskProgress() — 取消任务进度
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

window.ChatLog = window.ChatLog || console;

(function () {
    'use strict';

    const ChatLog = window.ChatLog;

    // 步骤类型 → 图标映射
    const TYPE_ICONS = {
        analysis: '💡',
        data: '📊',
        formula: '📐',
        format: '🎨',
        chart: '📈',
        verify: '✅',
        default: '📌'
    };

    // 步骤状态枚举
    const STATUS = { PENDING: 'pending', RUNNING: 'running', COMPLETED: 'completed', FAILED: 'failed' };

    // ============================================================
    // 工具函数
    // ============================================================
    function resolveApp() {
        return (typeof App !== 'undefined') ? App : window.App;
    }

    /** 安全的 DOM 滚动 */
    function safeScrollToBottom() {
        try {
            const container = document.getElementById('chat-messages');
            if (container) {
                container.scrollTop = container.scrollHeight;
            }
        } catch (_) { /* 忽略滚动异常 */ }
    }

    /** 对 step.type 做兼容处理 */
    function normalizeType(type) {
        if (!type || typeof type !== 'string') return 'default';
        return TYPE_ICONS[type.toLowerCase()] ? type.toLowerCase() : 'default';
    }

    // ============================================================
    // 核心类
    // ============================================================
    class TaskProgressPanel {
        constructor() {
            /** @type {boolean} */
            this.isInitialized = false;
            /** @type {Array<Object>} 原始步骤数据 */
            this.steps = [];
            /** @type {number} 当前正在执行的步骤索引，-1 表示无 */
            this.currentStepIndex = -1;
            /** @type {string} 面板标题 */
            this.title = '任务进度';
            /** @type {string} 所属消息 wrapper 的 DOM id */
            this.ownerMessageId = null;
            /** @type {string[]} 步骤状态快照: 'pending'|'running'|'completed'|'failed' */
            this._stepStatus = [];
            /** @type {string[]} 每步详情 HTML（累积） */
            this._stepDetailHtml = [];

            // DOM 缓存（惰性创建）
            /** @type {HTMLElement|null} */
            this._panelEl = null;
            /** @type {HTMLElement|null} */
            this._counterEl = null;
            /** @type {HTMLElement|null} */
            this._progressFill = null;
            /** @type {HTMLElement[]} */
            this._stepEls = [];
            /** @type {Map<number, HTMLElement>} */
            this._detailEls = null;

            /** @type {boolean} 整体折叠 */
            this._collapsed = false;
        }

        // ============================================================
        // 公共 API
        // ============================================================

        /**
         * 初始化面板，在指定的 AI 消息气泡内创建 DOM。
         * @param {Array<Object>} steps - [{id, name, description, type, estimatedTime}, ...]
         * @param {Object} [options]
         * @param {string} [options.title] - 面板标题
         * @param {string} [options.ownerMessageId] - 所属消息 wrapper 的 DOM id
         */
        init(steps, options) {
            if (!steps || !Array.isArray(steps) || steps.length === 0) {
                ChatLog.warn('TASK', '[TaskProgressPanel] init() 收到空步骤，跳过');
                this.isInitialized = false;
                return;
            }

            // 重置状态
            this.steps = steps.slice();
            this.currentStepIndex = -1;
            this.title = (options && options.title) || '任务进度';
            this.ownerMessageId = (options && options.ownerMessageId) || null;
            this._stepStatus = this.steps.map(() => STATUS.PENDING);
            this._stepDetailHtml = this.steps.map(() => '');
            this._stepEls = [];
            this._detailEls = new Map();
            this._collapsed = false;

            // 挂载到目标消息
            const parentEl = this._resolveParentElement();
            if (!parentEl) {
                ChatLog.warn('TASK', '[TaskProgressPanel] 找不到父元素，面板无法挂载');
                this.isInitialized = false;
                return;
            }

            // 移除旧面板（如果存在）
            const oldPanel = parentEl.querySelector('.task-progress-panel');
            if (oldPanel) { oldPanel.remove(); }

            // 构建面板 DOM
            this._panelEl = this._buildPanelDom();
            parentEl.appendChild(this._panelEl);

            this.isInitialized = true;
            ChatLog.info('TASK', `[TaskProgressPanel] 初始化完成: "${this.title}", ${this.steps.length} 步`);
            safeScrollToBottom();
        }

        /**
         * 标记步骤开始执行。
         * @param {number} index - 步骤索引
         */
        startStep(index) {
            if (!this.isInitialized || !this._isValidIndex(index)) return;
            this.currentStepIndex = index;
            this._setStepStatus(index, STATUS.RUNNING);
            this._autoExpandStep(index);
            this._updateProgressBar();
            safeScrollToBottom();
        }

        /**
         * 标记步骤执行完成。
         * @param {number} index - 步骤索引
         */
        completeStep(index) {
            if (!this.isInitialized || !this._isValidIndex(index)) return;
            this._setStepStatus(index, STATUS.COMPLETED);
            this._updateProgressBar();
        }

        /**
         * 标记步骤执行失败。
         * @param {number} index - 步骤索引
         * @param {string} [errorMsg] - 错误信息（追加到详情末尾）
         */
        failStep(index, errorMsg) {
            if (!this.isInitialized || !this._isValidIndex(index)) return;
            this._setStepStatus(index, STATUS.FAILED);
            this._autoExpandStep(index);
            if (errorMsg) {
                this.appendStepDetail(index,
                    `<div style="margin-top:6px;color:#ef4444;">❌ ${String(errorMsg)}</div>`
                );
            }
            this._updateProgressBar();
        }

        /**
         * 设置步骤的详情内容（替换）。
         * @param {number} index - 步骤索引
         * @param {string} html - HTML 片段
         */
        setStepDetail(index, html) {
            if (!this.isInitialized || !this._isValidIndex(index)) return;
            this._stepDetailHtml[index] = html || '';
            this._renderDetail(index);
            safeScrollToBottom();
        }

        /**
         * 向步骤详情追加内容。
         * @param {number} index - 步骤索引
         * @param {string} html - HTML 片段
         */
        appendStepDetail(index, html) {
            if (!this.isInitialized || !this._isValidIndex(index) || !html) return;
            this._stepDetailHtml[index] = (this._stepDetailHtml[index] || '') + html;
            this._renderDetail(index);
            safeScrollToBottom();
        }

        /** 强制刷新面板 DOM */
        render() {
            if (!this.isInitialized) return;
            const parentEl = this._resolveParentElement();
            if (!parentEl) return;

            const oldPanel = parentEl.querySelector('.task-progress-panel');
            if (oldPanel) oldPanel.remove();

            this._panelEl = this._buildPanelDom();
            parentEl.appendChild(this._panelEl);

            // 恢复状态
            this._stepStatus.forEach((status, i) => {
                this._updateStepDomClass(i, status);
                if (this._stepEls[i] && this._stepDetailHtml[i]) {
                    this._renderDetail(i);
                }
            });
            this._updateProgressBar();
            safeScrollToBottom();
        }

        /**
         * 将所有未完成的步骤标记为完成。
         */
        finishAll() {
            if (!this.isInitialized) return;
            this.steps.forEach((_, i) => {
                if (this._stepStatus[i] === STATUS.PENDING || this._stepStatus[i] === STATUS.RUNNING) {
                    this._setStepStatus(i, STATUS.COMPLETED);
                }
            });
            this._updateProgressBar();
        }

        /**
         * 切换某个步骤详情的展开/折叠。
         * @param {number} index - 步骤索引
         */
        toggleStepDetail(index) {
            if (!this.isInitialized || !this._isValidIndex(index)) return;
            const stepEl = this._stepEls[index];
            if (!stepEl) return;
            stepEl.classList.toggle('expanded');
            this._updateToggleIcon(index);
        }

        /**
         * 切换整个面板的折叠/展开。
         */
        toggleCollapse() {
            if (!this.isInitialized || !this._panelEl) return;
            this._collapsed = !this._collapsed;
            const stepsEl = this._panelEl.querySelector('.tp-steps');
            const barEl = this._panelEl.querySelector('.step-progress-bar');
            if (stepsEl) stepsEl.style.display = this._collapsed ? 'none' : '';
            if (barEl) barEl.style.display = this._collapsed ? 'none' : '';
        }

        /**
         * 重置面板状态并移除 DOM。
         */
        reset() {
            if (this._panelEl && this._panelEl.parentNode) {
                this._panelEl.parentNode.removeChild(this._panelEl);
            }
            this.isInitialized = false;
            this.steps = [];
            this.currentStepIndex = -1;
            this.title = '任务进度';
            this.ownerMessageId = null;
            this._stepStatus = [];
            this._stepDetailHtml = [];
            this._stepEls = [];
            this._detailEls = new Map();
            this._panelEl = null;
            this._counterEl = null;
            this._progressFill = null;
            this._collapsed = false;
            ChatLog.info('TASK', '[TaskProgressPanel] 已重置');
        }

        /** 滚动到聊天容器底部。 */
        scrollToBottom() {
            safeScrollToBottom();
        }

        /**
         * 静态工厂：从 JSON 快照创建面板实例（用于历史恢复）。
         * @param {Object} json
         * @returns {TaskProgressPanel|null}
         */
        static fromJSON(json) {
            if (!json || !json.steps || !Array.isArray(json.steps)) return null;
            const panel = new TaskProgressPanel();
            panel.steps = json.steps.slice();
            panel.title = json.title || '任务进度';
            panel.ownerMessageId = json.ownerMessageId || null;
            panel._stepStatus = (json._stepStatus && Array.isArray(json._stepStatus))
                ? json._stepStatus.slice()
                : panel.steps.map(() => STATUS.PENDING);
            panel._stepDetailHtml = (json._stepDetailHtml && Array.isArray(json._stepDetailHtml))
                ? json._stepDetailHtml.slice()
                : panel.steps.map(() => '');
            return panel;
        }

        /** 导出快照（用于序列化到历史）。 */
        toJSON() {
            return {
                steps: this.steps.slice(),
                title: this.title,
                ownerMessageId: this.ownerMessageId,
                _stepStatus: this._stepStatus.slice(),
                _stepDetailHtml: this._stepDetailHtml.slice()
            };
        }

        // ============================================================
        // 私有方法
        // ============================================================

        /** 解析父容器（AI 消息的 .message 元素） */
        _resolveParentElement() {
            if (this.ownerMessageId) {
                const wrapper = document.getElementById(this.ownerMessageId);
                if (wrapper) {
                    return wrapper.querySelector('.message.ai') || wrapper;
                }
            }
            // 回退：寻找最后一个 AI 消息
            const allAiMsgs = document.querySelectorAll('#chat-messages .message.ai');
            return allAiMsgs.length > 0 ? allAiMsgs[allAiMsgs.length - 1] : null;
        }

        /** 索引合法性检查 */
        _isValidIndex(index) {
            return index >= 0 && index < this.steps.length;
        }

        /** 更新步骤状态并刷新 DOM */
        _setStepStatus(index, newStatus) {
            if (this._stepStatus[index] === newStatus) return;
            this._stepStatus[index] = newStatus;
            this._updateStepDomClass(index, newStatus);
        }

        /** 根据状态更新步骤 DOM 的图标类与 CSS 类 */
        _updateStepDomClass(index, newStatus) {
            const stepEl = this._stepEls[index];
            if (!stepEl) return;

            // 清除旧状态类
            stepEl.classList.remove('pending', 'running', 'completed', 'failed');
            stepEl.classList.add(newStatus);

            // 更新图标
            const iconEl = stepEl.querySelector('.tp-icon');
            if (!iconEl) return;

            switch (newStatus) {
                case STATUS.PENDING:
                    iconEl.innerHTML = '<span class="tp-icon-pending"></span>';
                    break;
                case STATUS.RUNNING:
                    iconEl.innerHTML = '<span class="tp-icon-running"></span>';
                    break;
                case STATUS.COMPLETED:
                    iconEl.innerHTML = '<span class="tp-icon-done">✓</span>';
                    break;
                case STATUS.FAILED:
                    iconEl.innerHTML = '<span class="tp-icon-failed">✗</span>';
                    break;
            }
        }

        /** 构建整个面板 DOM（仅在 init/render 时调用） */
        _buildPanelDom() {
            const panel = document.createElement('div');
            panel.className = 'task-progress-panel';

            // ---- 头部 ----
            const header = document.createElement('div');
            header.className = 'tp-header';

            const titleSpan = document.createElement('span');
            titleSpan.className = 'tp-title';
            titleSpan.textContent = this.title;
            header.appendChild(titleSpan);

            this._counterEl = document.createElement('span');
            this._counterEl.className = 'tp-counter';
            header.appendChild(this._counterEl);

            const closeBtn = document.createElement('button');
            closeBtn.className = 'tp-close-btn';
            closeBtn.textContent = '×';
            closeBtn.title = '折叠面板';
            closeBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                this.toggleCollapse();
            });
            header.appendChild(closeBtn);

            panel.appendChild(header);

            // ---- 进度条 ----
            const progressBar = document.createElement('div');
            progressBar.className = 'step-progress-bar';
            this._progressFill = document.createElement('div');
            this._progressFill.className = 'step-progress-fill';
            this._progressFill.style.width = '0%';
            progressBar.appendChild(this._progressFill);
            panel.appendChild(progressBar);

            // ---- 步骤列表 ----
            const stepsEl = document.createElement('div');
            stepsEl.className = 'tp-steps';

            this._stepEls = [];
            this._detailEls = new Map();

            this.steps.forEach((step, i) => {
                const stepEl = document.createElement('div');
                stepEl.className = 'tp-step pending';
                stepEl.dataset.stepIndex = i;

                // 主行
                const mainRow = document.createElement('div');
                mainRow.className = 'tp-step-main';
                mainRow.addEventListener('click', () => this.toggleStepDetail(i));

                // 图标
                const icon = document.createElement('span');
                icon.className = 'tp-icon';
                icon.innerHTML = '<span class="tp-icon-pending"></span>';
                mainRow.appendChild(icon);

                // 名称
                const name = document.createElement('span');
                name.className = 'tp-step-name';
                name.textContent = `${i + 1}. ${step.name || '步骤' + (i + 1)}`;
                mainRow.appendChild(name);

                // 类型小图标（通过 CSS 隐藏，供可选展示）
                const typeIcon = document.createElement('span');
                typeIcon.className = 'tp-type-icon';
                typeIcon.textContent = TYPE_ICONS[normalizeType(step.type)] || TYPE_ICONS.default;
                mainRow.appendChild(typeIcon);

                // 预估时间
                const est = document.createElement('span');
                est.className = 'tp-step-est';
                if (step.estimatedTime) est.textContent = step.estimatedTime;
                mainRow.appendChild(est);

                // 展开箭头
                const toggle = document.createElement('span');
                toggle.className = 'tp-step-toggle';
                mainRow.appendChild(toggle);

                stepEl.appendChild(mainRow);

                // 详情区
                const detail = document.createElement('div');
                detail.className = 'tp-step-detail';
                detail.style.display = 'none';

                const detailContent = document.createElement('div');
                detailContent.className = 'tp-step-detail-content';
                detail.appendChild(detailContent);

                stepEl.appendChild(detail);

                this._stepEls[i] = stepEl;
                this._detailEls.set(i, detailContent);
                stepsEl.appendChild(stepEl);
            });

            panel.appendChild(stepsEl);

            // 初始化计数器
            this._updateProgressBar();

            return panel;
        }

        /** 渲染步骤详情 */
        _renderDetail(index) {
            const detailContent = this._detailEls.get(index);
            if (!detailContent) return;
            const html = this._stepDetailHtml[index];
            if (html) {
                detailContent.innerHTML = html;
                // 确保详情区域可见
                const detailEl = detailContent.parentElement;
                if (detailEl) detailEl.style.display = '';
            }
        }

        /** 自动展开步骤（running/failed 时自动展开） */
        _autoExpandStep(index) {
            const stepEl = this._stepEls[index];
            if (!stepEl) return;
            stepEl.classList.add('expanded');
            const detailEl = stepEl.querySelector('.tp-step-detail');
            if (detailEl) detailEl.style.display = '';
            this._updateToggleIcon(index);
        }

        /** 同步展开箭头的旋转 */
        _updateToggleIcon(index) {
            const stepEl = this._stepEls[index];
            if (!stepEl) return;
            const detailEl = stepEl.querySelector('.tp-step-detail');
            if (detailEl) {
                detailEl.style.display = stepEl.classList.contains('expanded') ? '' : 'none';
            }
        }

        /** 更新进度条和完成数 */
        _updateProgressBar() {
            const total = this.steps.length;
            let doneCount = 0;

            this._stepStatus.forEach(s => {
                if (s === STATUS.COMPLETED) doneCount++;
            });

            if (this._counterEl) {
                this._counterEl.textContent = `${doneCount}/${total}`;
            }

            if (this._progressFill) {
                const pct = total > 0 ? Math.round((doneCount / total) * 100) : 0;
                this._progressFill.style.width = `${pct}%`;
            }
        }
    }

    // ============================================================
    // 全局暴露
    // ============================================================
    window.TaskProgressPanel = TaskProgressPanel;

    if (typeof window.App === 'undefined') {
        window.App = {};
    }
    window.App.TaskProgressPanel = TaskProgressPanel;

    // 复用已有实例或创建新实例（兼容多次加载场景）
    if (!window.App.taskProgress) {
        window.App.taskProgress = new TaskProgressPanel();
    }

    ChatLog.info('TASK', '[TaskProgressPanel] 模块已加载');
})();
