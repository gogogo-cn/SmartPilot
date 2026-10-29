/**
 * ============================================================
 * chat-v2/ui/renderers/steps-plan.js — 步骤规划清单渲染器（纯函数子模块）
 * ============================================================
 * 从 chat-renderer.js 拆出的独立子模块：
 *   - extractStepsPlan(result)   ：递归抽取"步骤规划"数组
 *   - renderStepsPlanHtml(steps) ：渲染成绿色"步骤规划"HTML 字符串
 *
 * 两者都是纯字符串/DOM-free，可被 node 直接单元测试。
 * 自包含 escapeHtml（与 chat-renderer 同款委托到 App.ChatV2.renderHtml.escapeHtml），
 * 避免跨模块依赖。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.renderers) App.ChatV2.renderers = {};

    const ChatLog = window.ChatLog || console;

    function escapeHtml(str) {
        if (str == null) return '';
        return App.ChatV2.renderHtml ? App.ChatV2.renderHtml.escapeHtml(str) : '';
    }

    /**
     * 从工具执行结果中提取"步骤规划"数组。
     * 兼容几种存放位置：顶层 steps / output.executionResult.steps / executionResult.steps 等。
     * @param {Object} result 原始执行结果（可为字符串/对象/包裹层）
     * @returns {Array<string>|null} 步骤数组；未找到返回 null
     */
    function extractStepsPlan(result) {
        if (result == null) return null;
        // 可能是 JSON 字符串 → 解析
        let scan = result;
        if (typeof scan === 'string') {
            try { scan = JSON.parse(scan); } catch (e) { return null; }
        }
        // 递归向下查找 steps 数组（限层数，避免误探）
        let found = null;
        const MAX_DEPTH = 8;
        (function walk(obj, depth) {
            if (found || obj == null || depth > MAX_DEPTH) return;
            // 字符串节点：可能内嵌"包着 JSON 的字符串"（如 executionResult.Message 为 "✓ … ：{…steps…}"）
            if (typeof obj === 'string') {
                const jsonStart = obj.indexOf('{');
                if (jsonStart >= 0) {
                    let parsed = null;
                    try { parsed = JSON.parse(obj.slice(jsonStart)); } catch (e) { parsed = null; }
                    if (parsed) { walk(parsed, depth + 1); return; }
                }
                return;
            }
            if (typeof obj !== 'object') return;
            if (obj.steps && Array.isArray(obj.steps) && obj.steps.length > 0) {
                const arr = obj.steps.filter(s => typeof s === 'string' && s.trim());
                if (arr.length) { found = arr; return; }
            }
            if (Array.isArray(obj)) {
                for (const item of obj) { walk(item, depth + 1); }
            } else {
                for (const key of Object.keys(obj)) {
                    walk(obj[key], depth + 1);
                }
            }
        })(scan, 0);
        return found;
    }

    /**
     * 【步骤统一到 cbx v20261003】步骤不再渲染旧的 .exec-plan-card 绿面板，
     * 改为与 thoughtList 完全同一套 CodeBuddy 风格 cbx 扁平列表（全站唯一过程视觉）。
     * 这样 多步骤 / Skill / 计算机链 / 工具结果内嵌步骤 四条路径共用一种样式，
     * 杜绝任何路径退回旧的"步骤规划 (N 步)"绿面板。
     * 状态兼容：done/success/ok → completed，error/err → failed。
     */
    function renderStepsPlanHtml(steps) {
        if (!steps || !steps.length) return '';

        const normStatus = function (raw) {
            const s = String(raw || 'done').toLowerCase();
            if (s === 'done' || s === 'success' || s === 'ok' || s === 'completed') return 'completed';
            if (s === 'error' || s === 'err' || s === 'failed') return 'failed';
            if (s === 'running') return 'running';
            return 'pending';
        };

        const rows = steps.map(function (s) {
            const obj = (s && typeof s === 'object') ? s : { title: String(s) };
            const st = normStatus(obj.status);
            const title = String(obj.title || obj.name || obj.description || obj.detail || '步骤');
            const detail = String(obj.detail || obj.description || '').trim();
            const running = (st === 'running');
            const bad = (st === 'failed');
            const done = (st === 'completed');
            const statusHtml = bad ? '<span class="cbx-status bad">✕ 失败</span>'
                : (done ? '<span class="cbx-status ok">✓</span>'
                    : (running ? '<span class="cbx-status">⏳ 执行中</span>' : '<span class="cbx-status">○ 等待</span>'));
            const body = detail
                ? '<div class="cbx-detail--text">' + escapeHtml(detail) + '</div>'
                : '<div class="cbx-detail--empty">（无详细信息）</div>';
            return '<details class="cbx-item cbx-thought' + (bad ? ' cbx-failed' : '') + (running ? ' cbx-running' : '') + '"' + (running ? ' open' : '') + '>' +
                '<summary>' +
                    '<span class="cbx-icon">📌</span>' +
                    '<span class="cbx-label">' + escapeHtml(title) + '</span>' +
                    '<span class="cbx-target"></span>' +
                    '<span class="cbx-meta">' + statusHtml + '</span>' +
                    '<span class="cbx-chevron">›</span>' +
                '</summary>' +
                '<div class="cbx-detail">' + body + '</div>' +
            '</details>';
        }).join('');

        return '<div class="cbx-list cbx-list--process">' + rows + '</div>';
    }

    App.ChatV2.renderers.stepsPlan = {
        extractStepsPlan,
        renderStepsPlanHtml
    };
    ChatLog.info?.('CHAT-V2', '✅ renderers/steps-plan 已加载');
})();