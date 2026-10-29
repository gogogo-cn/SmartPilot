/**
 * ============================================================
 * chat-v2/services/ProgressTracker.js — 多步骤进度追踪与UI渲染
 * ============================================================
 * 职责：
 *   - 维护步骤状态机（pending → running → completed/failed）
 *   - 构建规划面板 HTML
 *   - 更新状态变化后的 UI
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // HTML 转义（局部定义，不依赖外部 h()）
    function h(text) {
        if (text == null) return '';
        return String(text)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/'/g, '&#039;');
    }

    // 状态对应的徽章样式（text 运行时经 i18n 覆盖）
    const STATUS_BADGES = {
        pending: { text: '等待', color: '#94a3b8', bg: 'rgba(148, 163, 184, 0.1)' },
        running: { text: '执行中', color: '#3b82f6', bg: 'rgba(59, 130, 246, 0.1)' },
        completed: { text: '完成', color: '#10b981', bg: 'rgba(16, 185, 129, 0.1)' },
        failed: { text: '失败', color: '#ef4444', bg: 'rgba(239, 68, 68, 0.1)' }
    };

    function statusBadgeText(status) {
        const base = STATUS_BADGES[status] || STATUS_BADGES.pending;
        try {
            if (App && App.i18n && typeof App.i18n.t === 'function') {
                const key = 'progress.' + (status || 'pending');
                const s = App.i18n.t(key);
                if (s && s !== key) return s;
            }
        } catch (_e) { /* ignore */ }
        return base.text;
    }

    /**
     * 从步骤结果中提取简短的数据预览文本
     */
    /**
     * 解析工具返回的字符串结果
     * - 有效 JSON → 直接解析
     * - 无效 JSON（缺逗号）→ 修复后解析
     * - 非 JSON 文本 → 返回 { text: 原文 }
     */
    function parseToolResult(str) {
        if (typeof str !== 'string') return str;
        const trimmed = str.trim();
        if (!trimmed) return str;

        // 尝试直接解析
        try {
            const parsed = JSON.parse(trimmed);
            if (parsed && typeof parsed === 'object') return parsed;
        } catch (_) { /* 无效 JSON，尝试修复 */ }

        // 修复缺少逗号的 JSON（MCP 返回的常见格式问题）
        try {
            const repaired = trimmed
                .replace(/}\s*{/g, '},{')                       // 对象间补逗号
                .replace(/"\s*\n\s*"/g, '",\n"')               // key-value 间补逗号
                .replace(/(\d+)\s*\n\s*"/g, '$1,\n"');         // 值后补逗号
            const parsed = JSON.parse(repaired);
            if (parsed && typeof parsed === 'object') return parsed;
        } catch (_) { /* 修复后仍失败 */ }

        // 非 JSON 文本，包装为 { text }
        return { text: trimmed };
    }

    function extractResultPreview(result) {
        if (!result) return '';

        const data = deepFindData(result);
        if (data && data.length > 0) {
            const firstRow = data[0];
            const keys = (typeof firstRow === 'object' && !Array.isArray(firstRow))
                ? Object.keys(firstRow) : [];
            const colHint = keys.length > 0 ? `列: ${keys.join(', ')}` : '';
            const rowHint = `${data.length} 行`;
            return `📊 ${rowHint}${colHint ? ` · ${colHint}` : ''}`;
        }

        if (result.summary) return String(result.summary).slice(0, 60);
        if (result.message) return String(result.message).slice(0, 60);
        if (result.text) return String(result.text).slice(0, 60);
        if (typeof result.result === 'string') return result.result.slice(0, 60);

        return '';
    }

    /**
     * 渲染完整的数据详情（展开后显示）
     */
    /**
     * 渲染步骤展开后的详细内容
     * - 聚合步骤 → 展示数据表格（聚合结果）
     * - 读取步骤 → 展示元数据（工作表/范围/列名/行数），不展示数据
     * - 其他 → 回退到预览行或深度搜索
     */
    function renderFullResult(result) {
        if (!result) return '';

        // 读操作步骤：展示元数据（工作表、范围、列名、行数），range 显示实际选区地址
        if (result.worksheet && !result.aggregated) {
            const parts = [];
            parts.push(`<div class="plan-step-meta">`);
            parts.push(`<div class="meta-row"><span class="meta-label">工作表</span><span class="meta-value">${h(result.worksheet)}</span></div>`);
            if (result.range) {
                parts.push(`<div class="meta-row"><span class="meta-label">范围</span><span class="meta-value">${h(result.range)}</span></div>`);
            }
            if (Array.isArray(result.headers) && result.headers.length > 0) {
                parts.push(`<div class="meta-row"><span class="meta-label">列名</span><span class="meta-value">${result.headers.map(h).join(' · ')}</span></div>`);
            }
            if (result.rows) {
                parts.push(`<div class="meta-row"><span class="meta-label">行数</span><span class="meta-value">${result.rows} 行</span></div>`);
            }
            parts.push(`</div>`);
            // 操作结果消息（如 "成功创建工作表: 汇总"）
            if (result.message) {
                parts.push(`<div class="plan-step-text-result">${h(String(result.message))}</div>`);
            }
            return parts.join('\n');
        }

        // 聚合工具展示 result.data（聚合结果，小数据量）
        // 支持 excel_aggregate（aggregated=true）和 excel_date_group（dateGrouped=true）
        let data = null;
        if ((result.aggregated === true || result.dateGrouped === true) && Array.isArray(result.data) && result.data.length > 0) {
            data = result.data;
        } else if (Array.isArray(result.previewRows) && result.previewRows.length > 0) {
            data = result.previewRows;
        } else {
            data = deepFindData(result);
        }
        if (data && data.length > 0) {
            const isObjects = typeof data[0] === 'object' && !Array.isArray(data[0]);
            const keys = isObjects ? Object.keys(data[0]) : [];
            const maxRows = Math.min(data.length, 20);

            if (isObjects && keys.length > 0) {
                let table = '<table class="multi-step-data-table"><thead><tr>';
                for (const key of keys) {
                    table += `<th>${h(key)}</th>`;
                }
                table += '</tr></thead><tbody>';
                for (let r = 0; r < maxRows; r++) {
                    table += '<tr>';
                    for (const key of keys) {
                        const val = data[r][key];
                        const display = val != null ? String(val) : '';
                        table += `<td>${h(display)}</td>`;
                    }
                    table += '</tr>';
                }
                table += '</tbody></table>';
                if (data.length > maxRows) {
                    table += `<div class="plan-step-more">... 还有 ${data.length - maxRows} 行</div>`;
                }
                return table;
            } else if (Array.isArray(data[0])) {
                const maxRows = Math.min(data.length, 21);
                let table = '<table class="multi-step-data-table"><tbody>';
                for (let r = 0; r < maxRows; r++) {
                    const isHeader = r === 0;
                    table += `<tr>`;
                    for (const cell of data[r]) {
                        const tag = isHeader ? 'th' : 'td';
                        const display = cell != null ? String(cell) : '';
                        table += `<${tag}>${h(display)}</${tag}>`;
                    }
                    table += '</tr>';
                }
                table += '</tbody></table>';
                if (data.length > maxRows) {
                    table += `<div class="plan-step-more">... 还有 ${data.length - maxRows} 行</div>`;
                }
                return table;
            }
        }

        if (result.summary) return `<div class="plan-step-text-result">${h(String(result.summary))}</div>`;
        if (result.message) return `<div class="plan-step-text-result">${h(String(result.message))}</div>`;
        if (result.text) return `<div class="plan-step-text-result">${h(String(result.text))}</div>`;
        if (typeof result.result === 'string') return `<div class="plan-step-text-result">${h(result.result)}</div>`;

        // 调试用：显示原始 JSON
        try {
            const json = JSON.stringify(result, null, 2);
            if (json && json.length > 4 && json.length < 5000) {
                return `<pre class="plan-step-json">${h(json)}</pre>`;
            }
        } catch (e) { /* ignore */ }

        return '';
    }

    /**
     * 深度搜索结果对象，找到第一个有意义的数据数组
     * 处理 3~4 层嵌套：{ success: true, result: { data: [...] } }
     */
    function deepFindData(obj, depth = 0) {
        if (!obj || typeof obj !== 'object' || depth > 5) return null;

        // 直接是数组
        if (Array.isArray(obj)) return obj;

        // 常见的数组字段名
        const arrayKeys = ['data', 'rows', 'records', 'items', 'values', 'list'];
        for (const key of arrayKeys) {
            if (Array.isArray(obj[key]) && obj[key].length > 0) {
                return obj[key];
            }
        }

        // result 字段可能是对象，递归深入
        if (obj.result && typeof obj.result === 'object') {
            const found = deepFindData(obj.result, depth + 1);
            if (found) return found;
        }

        // 遍历所有对象属性，递归找数组
        for (const key of Object.keys(obj)) {
            if (key === 'result' || key === 'success' || key === 'message' || key === 'summary') continue;
            const val = obj[key];
            if (val && typeof val === 'object') {
                const found = deepFindData(val, depth + 1);
                if (found) return found;
            }
        }

        return null;
    }

    // 步骤结构
    /**
     * @typedef {Object} Step
     * @property {string} id - 步骤ID
     * @property {string} name - 步骤名称
     * @property {string} description - 步骤描述
     * @property {string} status - pending|running|completed|failed
     * @property {string} strategy - direct|agent|pipeline
     * @property {any} result - 执行结果
     */

    // [multi-turn-isolated] steps state, keyed by turn.id (prevents concurrent turns overwriting each other)
    const _stepsByTurn = new Map();
    let _activeTurnId = null;

    /** Internal helper: resolve steps for a given turn (or activeTurnId as fallback) */
    function _resolveSteps(turn) {
        const turnId = (turn && turn.id) || _activeTurnId;
        if (!turnId) return [];
        const arr = _stepsByTurn.get(turnId);
        return Array.isArray(arr) ? arr : [];
    }
    function _setSteps(turn, steps) {
        const turnId = (turn && turn.id) || _activeTurnId;
        if (turnId) { _stepsByTurn.set(turnId, steps); _activeTurnId = turnId; }
    }

    /**
     * 初始化追踪器
     * @param {Array<Step>} steps - 步骤数组
     * @param {Object} turn - turn 对象
     * @param {Object} sessionStore - sessionStore
     */
    function init(steps, turn, sessionStore) {
        const copied = (steps || []).map(s => ({ ...s }));
        // ✨ 去重：同一 turn 若已 init 且内容完全一致才跳过（防三重 init 闪烁）；
        //    只要 status/detail/toolName/params 任一变化就用最新状态覆盖，
        //    否则"骨架先出→逐步填色"时展开区永远显示旧的 pending 骨架。
        const turnId = (turn && turn.id) || _activeTurnId;
        const existing = turnId ? _stepsByTurn.get(turnId) : null;
        if (existing && Array.isArray(existing) && existing.length === copied.length) {
            const identical = copied.every((s, i) => {
                const e = existing[i] || {};
                const sKey = (s.params == null) ? '' : JSON.stringify(s.params);
                const eKey = (e.params == null) ? '' : JSON.stringify(e.params);
                const sSub = (s.subSteps == null) ? '' : JSON.stringify(s.subSteps);
                const eSub = (e.subSteps == null) ? '' : JSON.stringify(e.subSteps);
                return s.status === e.status
                    && (s.detail === e.detail)
                    && (s.toolName === e.toolName)
                    && (sKey === eKey)
                    && (sSub === eSub);
            });
            if (identical) {
                ChatLog.info?.('PROGRESS-TRACKER', `[init] 跳过（turn=${turnId} 已有 ${copied.length} 步，数据一致）`);
                return;
            }
        }
        _setSteps(turn, copied);
        _activeTurnId = turnId;
        ChatLog.info?.('PROGRESS-TRACKER', '[init] ' + copied.length + ' 个步骤初始化');

        // 【任务套任务】监听子步骤状态更新事件，实时刷新嵌套UI
        if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.on === 'function') {
            App.ChatV2.agentEvents.on('multi_step:substep_update', (payload) => {
                if (!payload || !payload.stepId) return;
                const cbSteps = _resolveSteps({ id: payload.turnId });
                const targetStep = cbSteps.find(s => s.id === payload.stepId);
                if (targetStep && Array.isArray(payload.subSteps)) {
                    targetStep.subSteps = payload.subSteps;
                    // 触发规划面板重渲染
                    if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.emit === 'function') {
                        App.ChatV2.agentEvents.emit('multi_step:update', {
                            stepIndex: cbSteps.indexOf(targetStep),
                            status: targetStep.status,
                            turnId: payload.turnId
                        });
                    }
                }
            });
        }
    }

    /**
     * 更新步骤状态
     * @param {number} index - 步骤索引
     * @param {Step} step - 更新后的步骤
     * @param {Object} turn - turn 对象
     * @param {Object} sessionStore - sessionStore
     */
    function updateStep(index, step, turn, sessionStore) {
        if (index >= 0 && index < _resolveSteps().length) {
            _resolveSteps()[index] = { ..._resolveSteps()[index], ...step };
        }
        // 状态变化广播事件（供渲染层监听）
        if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.emit === 'function') {
            App.ChatV2.agentEvents.emit('multi_step:update', {
                stepIndex: index,
                status: step.status,
                turnId: turn.id
            });
        }
    }

    /**
     * 构建规划面板 HTML（增强版：进度条 + 动画 + 匹配演示文档效果）
     * @returns {string} HTML
     */
    function buildPlanHtml() {
        if (!Array.isArray(_resolveSteps()) || _resolveSteps().length === 0) return '';

        const stats = getStats();
        const rows = _resolveSteps().map((step, i) => renderStepHtml(step, i)).join('');
        const title = _resolveSteps().length === 1 ? '📋 步骤规划' : `📋 步骤规划 (${_resolveSteps().length} 步)`;

        return `<div class="exec-plan-card">
  <div class="plan-title">${title}</div>
  ${buildProgressBar(stats.completed, stats.total)}
  ${rows}
</div>`;
    }

    /**
     * 进度条 HTML（进度条 + 文案），整块渲染与增量渲染共用，保证两处视觉一致。
     * @param {number} completed - 已完成步数
     * @param {number} total - 总步数
     * @returns {string} HTML
     */
    function buildProgressBar(completed, total) {
        const pct = total > 0 ? Math.round((completed / total) * 100) : 0;
        return `<div class="plan-progress-bar">
  <div class="plan-progress-fill" style="width:${pct}%"></div>
</div>
<div class="plan-progress-label">${completed}/${total} 步完成 (${pct}%)</div>`;
    }

    /**
     * 渲染单个步骤行 HTML（自 buildPlanHtml 抽取，视觉零改动）。
     * 额外输出 data-step-id / data-step-index，供渲染层按 id 做增量原位替换。
     * @param {Step} step - 步骤对象
     * @param {number} index - 步骤在面板中的 0-based 索引
     * @returns {string} HTML
     */
    function renderStepHtml(step, index) {
        const stepId = (step.stepId != null) ? String(step.stepId)
            : (step.id != null ? String(step.id) : 'step-' + index);
        const status = step.status || 'pending';
        const badge = STATUS_BADGES[status] || STATUS_BADGES.pending;
        const badgeLabel = statusBadgeText(status);
        const desc = step.description || step.title || step.name || step.toolName || `步骤 ${index + 1}`;

        const strategyLabel = step.strategy === 'direct' ? '工具' :
                              step.strategy === 'agent' ? '推理' :
                              step.strategy === 'pipeline' ? '管道' : '';
        const strategyClass = step.strategy || 'direct';
        const strategyBadge = strategyLabel ?
            `<span class="plan-strategy-badge ${strategyClass}">${strategyLabel}</span>` : '';

        // 步骤主行（始终可见）
        const stepRow = `<div class="plan-step status-${status}" data-step-id="${h(stepId)}" data-step-index="${index}">
  <div class="plan-step-index">${index + 1}</div>
  <div class="plan-step-desc">${desc}${strategyBadge}</div>
  <span class="plan-step-badge status-${status}">${badgeLabel}</span>
</div>`;

        // 可折叠的数据详情（仅已完成步骤有数据）
        let detailHtml = '';
        if (status === 'completed') {
            // ✨ 展示该步真实操作：工具名 + 人类可读参数（让'展开'看到实际执行的调用）
            const _toolName = step.toolName || '';
            const _showArgs = step.params && typeof step.params === 'object' ? step.params : null;
            let _argText = '';
            if (_showArgs) {
                try { _argText = Object.keys(_showArgs).map(k => k + ':' + (typeof _showArgs[k] === 'object' ? JSON.stringify(_showArgs[k]) : String(_showArgs[k]))).join('，'); } catch (err) { _argText = String(_showArgs); }
            }
            const _argSuffix = _argText ? (' · ' + h(_argText)) : '';
            const _toolInfo = (_toolName || _argText)
                ? '<div class="plan-step-tool">🔨 <b>' + h(_toolName) + '</b>' + _argSuffix + '</div>'
                : '';
            // ✨ 优先 step.detail（骨架预写的 + progress update 传的），
            // 再 fallback step.result（invokeTracked 原始返回，需层层解析）
            const stepDetailText = typeof step.detail === 'string' ? step.detail.trim() : '';
            if (stepDetailText) {
                detailHtml = _toolInfo + `<div class="plan-step-detail">${stepDetailText}</div>`;
            } else if (step.result) {
                // step.result = { ok: true, result: { ok: true, result: { data: [...] }, ... }, ... }
                // 逐层展开：executeStep → invokeTracked → tool data
                const invokeResult = step.result.result || step.result;   // invokeTracked 返回
                let toolResult = invokeResult.result || invokeResult;     // 实际工具数据
                // 保险：工具返回的可能是 JSON 字符串（含无效 JSON，需修复后解析）
                if (typeof toolResult === 'string') {
                    toolResult = parseToolResult(toolResult);
                }
                // 处理 MCP 标准响应格式 { content: [{ type: "text", text: "..." }] }
                // 提取 content[0].text 中的 JSON 并解析，替代原始 JSON 展示
                if (toolResult && typeof toolResult === 'object' && Array.isArray(toolResult.content) && toolResult.content.length > 0) {
                    for (const item of toolResult.content) {
                        if (item && typeof item.text === 'string' && item.text.trim()) {
                            const trimmed = item.text.trim();
                            if (trimmed.startsWith('{') || trimmed.startsWith('[')) {
                                const parsed = parseToolResult(trimmed);
                                if (parsed && typeof parsed === 'object') {
                                    toolResult = parsed;
                                    break;
                                }
                            }
                            // 非 JSON 文本（如 "成功复制数据：..."），直接作为 result.text
                            if (!trimmed.startsWith('{') && !trimmed.startsWith('[')) {
                                toolResult = { text: trimmed };
                                break;
                            }
                        }
                    }
                }
                const preview = extractResultPreview(toolResult);
                const fullData = renderFullResult(toolResult);
                detailHtml = _toolInfo + `<div class="plan-step-detail">${fullData || preview || ''}</div>`;
            } else { detailHtml = _toolInfo; } // end else if (step.result)
        } // end if (status === 'completed')

            // ✨ 嵌套子步骤（Skill 大步骤嵌 op 小步骤）——所有状态都渲染，pending 时也能看到骨架
        const subSteps = step.subSteps || step.substeps || step.children || [];
        let subHtml = '';
        if (subSteps.length > 0) {
            const subItems = subSteps.map((sub, si) => {
                const subStatus = sub.status || sub.state || 'pending';
                const subName = typeof sub === 'string' ? sub : (sub.description || sub.title || sub.name || sub);
                let subIcon = '▹';
                if (subStatus === 'completed') subIcon = '✓';
                else if (subStatus === 'running') subIcon = '⟳';
                else if (subStatus === 'failed') subIcon = '✕';
                const subBadgeText = subStatus === 'completed' ? '完成' :
                                     subStatus === 'running' ? '执行中' :
                                     subStatus === 'failed' ? '失败' : '等待';
                // ✨ 子步骤也展示真实工具名 + 人类可读参数（与主步骤的 plan-step-tool 一致）
                let _subTool = '';
                const _subToolName = sub && sub.toolName ? String(sub.toolName) : '';
                let _subArgText = '';
                if (sub && sub.params && typeof sub.params === 'object') {
                    try { _subArgText = Object.keys(sub.params).map(k => k + ':' + (typeof sub.params[k] === 'object' ? JSON.stringify(sub.params[k]) : String(sub.params[k]))).join('，'); } catch (err) { _subArgText = String(sub.params); }
                }
                if (_subToolName || _subArgText) {
                    _subTool = '<div class="plan-substep-tool">🔨 <b>' + h(_subToolName) + '</b>' + (_subArgText ? (' · ' + h(_subArgText)) : '') + '</div>';
                }
                const subDetail = typeof sub.detail === 'string' && sub.detail.trim() && subStatus === 'completed'
                    ? `<div class="plan-substep-detail">${sub.detail.trim()}</div>` : '';
                return `<div class="plan-substep status-${subStatus}" data-substep="${si}">
  <span class="plan-substep-icon">${subIcon}</span>
  <span class="plan-substep-text">${subName}</span>
  <span class="plan-substep-badge status-${subStatus}">${subBadgeText}</span>
  ${_subTool}
  ${subDetail}
</div>`;
            }).join('');
            subHtml = `<div class="plan-substeps">${subItems}</div>`;
        }

        // 用 <details> 实现可折叠（有 subSteps 或 completed 时可折叠）
        const isExpandable = subSteps.length > 0 || status === 'completed';
        if (isExpandable) {
            return `<details class="plan-step-wrapper status-${status}" data-step-id="${h(stepId)}" data-step-index="${index}" ${subSteps.length > 0 ? 'open' : ''}>
  <summary>${stepRow}</summary>
  ${detailHtml}
  ${subHtml}
</details>`;
        }
        return stepRow;
    }

    /**
     * 增量 upsert 单个步骤（不整块重建面板，避免闪烁与 <details> 展开态丢失）。
     * - 命中已有 stepId → 合并更新该步
     * - 未命中且给了 step.index → 按位插入（占位骨架被真实步骤替换）
     * - 未命中且无 index → 追加到末尾
     * @param {string|number} id - 步骤标识
     * @param {Step} step - 步骤数据（可含 index 指定插入位）
     * @returns {number} 更新后的索引；未生效返回 -1
     */
    function upsertStep(id, step) {
        if (id == null) return -1;
        const key = String(id);
        const steps = _resolveSteps();
        let idx = steps.findIndex(s => String((s && s.stepId != null) ? s.stepId : (s && s.id)) === key);
        const incoming = Object.assign({}, step || {}, { stepId: key });
        if (idx === -1) {
            const rawAt = Number(incoming.index);
            const at = (incoming.index != null && Number.isInteger(rawAt) && rawAt >= 0)
                ? Math.min(rawAt, steps.length) : steps.length;
            steps.splice(at, 0, incoming);
            idx = at;
        } else {
            steps[idx] = Object.assign({}, steps[idx], incoming);
        }
        _setSteps(null, steps);
        return idx;
    }

    /**
     * 无状态渲染"现代"步骤规划面板：从任意步骤数组生成与 buildPlanHtml 完全一致的可折叠富面板。
     * 供通用兜底 / computer 操作链等"无 ProgressTracker 实例"的路径使用，保证全站只有这一种步骤面板视觉。
     * 兼容旧简单形状：{ title, detail, status:'done'|'pending'|'running'|'failed' }，done→completed 归一。
     */
    function renderStepsPanel(steps) {
        if (!Array.isArray(steps) || steps.length === 0) return '';
        const items = steps.map((s, idx) => {
            const obj = (s && typeof s === 'object') ? s : { title: String(s) };
            const status = obj.status === 'done' ? 'completed' : (obj.status || 'pending');
            return Object.assign({}, obj, {
                status,
                description: obj.description || obj.title || obj.name || obj.toolName || ('步骤 ' + (idx + 1))
            });
        });
        const total = items.length;
        const completed = items.filter(s => s.status === 'completed').length;
        const pct = total > 0 ? Math.round((completed / total) * 100) : 0;
        const progressBar = `<div class="plan-progress-bar">
  <div class="plan-progress-fill" style="width:${pct}%"></div>
</div>
<div class="plan-progress-label">${completed}/${total} 步完成 (${pct}%)</div>`;
        const title = total === 1 ? '📋 步骤规划' : `📋 步骤规划 (${total} 步)`;
        const rows = items.map((st, i) => {
            const badgeLabel = statusBadgeText(st.status);
            const strategyLabel = st.strategy === 'direct' ? '工具' :
                                  st.strategy === 'agent' ? '推理' :
                                  st.strategy === 'pipeline' ? '管道' : '';
            const strategyBadge = strategyLabel ?
                `<span class="plan-strategy-badge ${st.strategy || 'direct'}">${strategyLabel}</span>` : '';
            const stepId = (st.stepId != null) ? String(st.stepId)
                : (st.id != null ? String(st.id) : 'step-' + i);
            const row = `<div class="plan-step status-${st.status}" data-step-id="${h(stepId)}" data-step-index="${i}">
  <div class="plan-step-index">${i + 1}</div>
  <div class="plan-step-desc">${h(st.description)}${strategyBadge}</div>
  <span class="plan-step-badge status-${st.status}">${badgeLabel}</span>
</div>`;
            let detailHtml = '';
            if (st.status === 'completed') {
                const hasTool = st.toolName || (st.params && typeof st.params === 'object' && Object.keys(st.params).length > 0);
                let toolInfo = '';
                if (hasTool) {
                    let argText = '';
                    if (st.params && typeof st.params === 'object') {
                        try { argText = Object.keys(st.params).map(k => k + ':' + (typeof st.params[k] === 'object' ? JSON.stringify(st.params[k]) : String(st.params[k]))).join('，'); } catch (err) { argText = String(st.params); }
                    }
                    toolInfo = '<div class="plan-step-tool">🔨 <b>' + h(st.toolName || '') + '</b>' + (argText ? (' · ' + h(argText)) : '') + '</div>';
                }
                const d = typeof st.detail === 'string' && st.detail.trim();
                detailHtml = toolInfo + (d ? `<div class="plan-step-detail">${h(st.detail.trim())}</div>` : '');
            }
            if (st.status === 'completed' || detailHtml) {
                return `<details class="plan-step-wrapper status-${st.status}" data-step-id="${h(stepId)}" data-step-index="${i}">
  <summary>${row}</summary>
  ${detailHtml}
</details>`;
            }
            return row;
        }).join('');
        return `<div class="exec-plan-card">
  <div class="plan-title">${title}</div>
  ${progressBar}
  ${rows}
</div>`;
    }

    /**
     * 获取当前所有步骤
     * @returns {Array<Step>}
     */
    function getSteps() {
        return [..._resolveSteps()];
    }

    /**
     * 获取已完成的结果列表
     * @returns {Array}
     */
    function getResults() {
        return _resolveSteps().filter(s => s.result !== undefined).map(s => s.result);
    }

    /**
     * 统计完成/失败
     * @returns {Object} { total, completed, failed }
     */
    function getStats() {
        const completed = _resolveSteps().filter(s => s.status === 'completed').length;
        const failed = _resolveSteps().filter(s => s.status === 'failed').length;
        return { total: _resolveSteps().length, completed, failed };
    }

    App.ChatV2.ProgressTracker = {
        init,
        updateStep,
        upsertStep,
        buildPlanHtml,
        buildProgressBar,
        renderStepHtml,
        renderStepsPanel,
        getSteps,
        getResults,
        getStats
    };

    ChatLog.info?.('CHAT-V2', '✅ ProgressTracker loaded');
})();