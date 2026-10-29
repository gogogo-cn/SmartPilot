/**
 * ============================================================
 * chat-v2/services/tool-registry.js — 工具注册中心
 * ============================================================
 * 职责：统一注册 Skill / MCP / 内置工具，提供一致调用接口
 * 设计：新增工具零代码修改接入；execute 异步返回结果
 * 兼容：从 app-mcp.js 和 app-skill-mgmt.js 拉取工具列表
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    const ERROR_HINT = /(错误|失败|异常|出错|error|fail|exception|invalid|不可用|无法)/i;

    function hasErrorField(v, depth) {
        if (v == null) return false;
        if (typeof v === 'string') return ERROR_HINT.test(v);
        if (depth > 3) return false;
        if (Array.isArray(v)) {
            return v.some(item => {
                if (typeof item === 'string' && item.trim() && ERROR_HINT.test(item)) return true;
                if (item && typeof item === 'object') return hasErrorField(item, depth + 1);
                return false;
            });
        }
        if (typeof v === 'object') {
            for (const k of ['error', 'Error', 'errMsg', 'ErrMsg', 'errorMessage', 'ErrorMessage', 'errorDetail', 'exception']) {
                const val = v[k];
                if (typeof val === 'string' && val.trim()) return true;
                if (val && typeof val === 'object' && hasErrorField(val, depth + 1)) return true;
            }
            const warnList = v.warnings || v.Warnings;
            if (Array.isArray(warnList) && warnList.length > 0) {
                // 切片器等工具多层回退会写「FALLBACK / 参数错误 / 异常来自 HRESULT」，最终仍可能 SUCCESS；
                // 诊断行不算失败；仅「缓存创建失败 / ❌ / 全部失败」等终态才算。
                if (warnList.some(w => {
                    if (typeof w !== 'string' || !w.trim()) return false;
                    if (/\[SLICER_DIAG\]|FALLBACK|回退/.test(w)) {
                        return /缓存创建失败|全部失败|未成功创建|❌\s*PivotChart\s*FAIL/i.test(w);
                    }
                    return ERROR_HINT.test(w);
                })) return true;
            }
            for (const key of Object.keys(v)) {
                if (key === 'result' || key === 'Result') {
                    const r = v[key];
                    if (typeof r === 'string') {
                        try {
                            const parsed = JSON.parse(r);
                            if (parsed && typeof parsed === 'object' && hasErrorField(parsed, depth + 1)) return true;
                        } catch (_) { if (typeof ChatLog !== 'undefined') ChatLog.warn('TOOL', '工具注册操作失败:', _ && _.message); }
                    } else if (r && typeof r === 'object' && hasErrorField(r, depth + 1)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    // 提取最深层的可读错误文本；找不到返回 null。
    function extractErrorField(v, depth) {
        if (v == null) return null;
        if (typeof v === 'string') return v;
        if (depth > 3) return null;
        if (typeof v === 'object') {
            for (const k of ['error', 'Error', 'errMsg', 'ErrMsg', 'errorMessage', 'ErrorMessage', 'errorDetail', 'exception']) {
                const val = v[k];
                if (typeof val === 'string' && val.trim()) return val.trim();
                if (val && typeof val === 'object') {
                    const inner = extractErrorField(val, depth + 1);
                    if (inner) return inner;
                }
            }
        }
        return null;
    }

    const _tools = new Map();      // name -> tool definition
    const _nameMap = new Map();   // 规范化名 -> 原始名（用于查找）
    const _initialized = { mcp: false, skill: false, mcpLoadAttempted: false };

    // ── Post-Execute 钩子链 ────────────────────────────────────────
    /** @type {Array<{name: string, fn: PostExecuteHook, priority: number}>} */
    const _postExecuteHooks = [];

    /**
     * Post-execute 钩子签名：
     * @callback PostExecuteHook
     * @param {Object} ctx - 钩子上下文
     * @param {string} ctx.toolName - 已执行的工具名
     * @param {Object} ctx.args - 工具调用参数
     * @param {Object} ctx.result - 工具执行结果（{ok, result, error, durationMs}）
     * @param {Object} [ctx.turn] - 当前 turn 对象（由 invoke 的 ctx 透传）
     * @returns {Promise<{kind: 'accept'|'replace', payload?: Object}|void>}
     *   - {kind:'accept'}: 不做任何变更，继续传递
     *   - {kind:'replace', payload: {ok, result, error}}: 替换结果
     *   - 返回 undefined/void: 等同于 accept
     */

    // ── 单调守卫链 ──────────────────────────────────────────────────
    /** @type {Array<{name: string, fn: ToolGuard, priority: number}>} */
    const _guards = [];
    // Pure-function relay to tool-registry.util.js (implementations moved out).
    function resolveChartType(...a){ return App.ChatV2.toolRegistryUtil?.['resolveChartType']?.(...a); }
    function _presentSucceeded(...a){ return App.ChatV2.toolRegistryUtil?.['_presentSucceeded']?.(...a); }
    function _backendReported(...a){ return App.ChatV2.toolRegistryUtil?.['_backendReported']?.(...a); }
    function normalizeName(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeName']?.(...a); }
    function normalizeSheetArg(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeSheetArg']?.(...a); }
    function inferDimColumnFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['inferDimColumnFromMessage']?.(...a); }
    function inferTimeGroupFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['inferTimeGroupFromMessage']?.(...a); }
    function _existingSheetNames(...a){ return App.ChatV2.toolRegistryUtil?.['_existingSheetNames']?.(...a); }
    function extractSheetNameFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['extractSheetNameFromMessage']?.(...a); }
    function defaultNewSheetName(...a){ return App.ChatV2.toolRegistryUtil?.['defaultNewSheetName']?.(...a); }
    function _hasValue(...a){ return App.ChatV2.toolRegistryUtil?.['_hasValue']?.(...a); }
    function _requiredParamKeys(...a){ return App.ChatV2.toolRegistryUtil?.['_requiredParamKeys']?.(...a); }
    function _paramLabel(...a){ return App.ChatV2.toolRegistryUtil?.['_paramLabel']?.(...a); }
    function resolveClearIntentFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['resolveClearIntentFromMessage']?.(...a); }
    function resolveIntentFlagsFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['resolveIntentFlagsFromMessage']?.(...a); }
    function applyExplicitSortColumn(...a){ return App.ChatV2.toolRegistryUtil?.['applyExplicitSortColumn']?.(...a); }
    function pickArgsKeys(...a){ return App.ChatV2.toolRegistryUtil?.['pickArgsKeys']?.(...a); }
    function toBool(...a){ return App.ChatV2.toolRegistryUtil?.['toBool']?.(...a); }
    function parseRangeAddr(...a){ return App.ChatV2.toolRegistryUtil?.['parseRangeAddr']?.(...a); }
    function combineSourceRanges(...a){ return App.ChatV2.toolRegistryUtil?.['combineSourceRanges']?.(...a); }
    function normalizeChartTypeValueForTool(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeChartTypeValueForTool']?.(...a); }
    function normalizeChartArgs(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeChartArgs']?.(...a); }
    function normalizeShapeArgs(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeShapeArgs']?.(...a); }

    function normalizeExportCsvArgs(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeExportCsvArgs']?.(...a); }

    function inferExportFilePath(...a){ return App.ChatV2.toolRegistryUtil?.['inferExportFilePath']?.(...a); }
    function inferHostExportFilePath(...a){ return App.ChatV2.toolRegistryUtil?.['inferHostExportFilePath']?.(...a); }
    function isPlaceholderPath(...a){ return App.ChatV2.toolRegistryUtil?.['isPlaceholderPath']?.(...a); }
    function isPlaceholderPathForParam(...a){ return App.ChatV2.toolRegistryUtil?.['isPlaceholderPathForParam']?.(...a); }
    function inferChartTitleFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['inferChartTitleFromMessage']?.(...a); }
    function inferValueColumnFromMessage(...a){ return App.ChatV2.toolRegistryUtil?.['inferValueColumnFromMessage']?.(...a); }
    function safeDisplayText(...a){ return App.ChatV2.toolRegistryUtil?.['safeDisplayText']?.(...a); }
    function inferChartDefaults(...a){ return App.ChatV2.toolRegistryUtil?.['inferChartDefaults']?.(...a); }
    function classifyError(...a){ return App.ChatV2.toolRegistryUtil?.['classifyError']?.(...a); }
    function normalizeMcpResult(...a){ return App.ChatV2.toolRegistryUtil?.['normalizeMcpResult']?.(...a); }
    function computeNextRound(...a){ return App.ChatV2.toolRegistryUtil?.['computeNextRound']?.(...a); }
    function validateOutput(...a){ return App.ChatV2.toolRegistryUtil?.['validateOutput']?.(...a); }
    function finalizeToolContent(...a){ return App.ChatV2.toolRegistryUtil?.['finalizeToolContent']?.(...a); }
    function presentToolCall(...a){ return App.ChatV2.toolRegistryUtil?.['presentToolCall']?.(...a); }
    function presentToolResult(...a){ return App.ChatV2.toolRegistryUtil?.['presentToolResult']?.(...a); }
    function buildSkillAction(...a){ return App.ChatV2.toolRegistryUtil?.['buildSkillAction']?.(...a); }
    function buildSkillInputs(...a){ return App.ChatV2.toolRegistryUtil?.['buildSkillInputs']?.(...a); }
    function resolveGuardMode(...a){ return App.ChatV2.toolRegistryUtil?.['resolveGuardMode']?.(...a); }

    /**
     * 把 excel_sort_range 的 sortBy 数组渲染成人类可读的排序依据串。
     * 还原原生"排序"对话框语义：列 + 排序依据（值/单元格颜色/字体颜色/条件格式图标）+ 次序。
     * @param {string|Array} sortBy JSON 字符串或已解析数组
     * @returns {string} 如 "销售额 降序" / "状态（单元格颜色 #FF0000）升序" / "次要关键字 部门 升序"
     */
    function describeSortKeys(sortBy) {
        let sb = sortBy;
        if (typeof sb === 'string') {
            try { sb = JSON.parse(sb); } catch (e) { sb = null; }
        }
        if (!Array.isArray(sb) || !sb.length) return '第一列 · 升序';
        return sb.map((k, i) => {
            const col = (k && k.column) || 'A';
            const dir = (k && k.order === 'descending') ? '降序' : '升序';
            let on = '';
            if (k && k.sortOn === 'cellColor') on = `（单元格颜色${k.color ? ' ' + k.color : ''}）`;
            else if (k && k.sortOn === 'fontColor') on = `（字体颜色${k.color ? ' ' + k.color : ''}）`;
            else if (k && k.sortOn === 'icon') on = '（条件格式图标）';
            return (i === 0 ? '' : '次要关键字 ') + `${col}${on} ${dir}`;
        }).join('，');
    }

    // ── 危险操作审批记忆（同一 Turn 内只问一次）────────────────────────
    // 背景：agent-loop 串行执行同一轮里的多个工具调用。若模型为删除多行而连续发出
    // 多次 excel_delete_row，每次都独立弹确认框 → 用户确认第 1 次后行已被删除，
    // 第 2 次弹窗时看到"数据已没才提示"，产生"先斩后奏"的错觉。
    // 做法：用户对某个危险工具一旦做出「确认 / 取消」决定，本 Turn 内后续同名工具
    // 沿用该决定，不再重复弹窗。作用域仅限当前 Turn（提问/任务），不影响后续对话。
    const GUARD_DECISION = new WeakMap();

    /**
     * 读取本 Turn 内对"删除/清空类危险操作"这一族已做出的决定。
     * 按 Turn 整族记忆而非按工具名：防止技能的 LLM 在被取消后换一个工具
     * （如 excel_delete_row 被取消后改用 excel_execute_vba）绕过用户决定。
     * @param {Object|undefined} turn 当前 turn 对象（无则不做记忆）
     * @returns {'allow'|'deny'|null}
     */
    function getTurnDangerousDecision(turn) {
        if (!turn || typeof turn !== 'object') return null;
        return GUARD_DECISION.get(turn) || null;
    }

    /**
     * 记录本 Turn 内对危险操作族的决定（第一次询问时由用户做出）
     * @param {Object|undefined} turn
     * @param {'allow'|'deny'} decision
     */
    function setTurnDangerousDecision(turn, decision) {
        if (!turn || typeof turn !== 'object') return;
        GUARD_DECISION.set(turn, decision);
    }

    /** 设置项：危险操作始终询问（不沿用本轮 allow） */
    function isDangerAlwaysAsk() {
        try {
            return !!(App && App.config && App.config.dangerAlwaysAsk);
        } catch (_e) {
            return false;
        }
    }

    // ── 通用"能力去重"：同一目标多路径可达，本轮已做过就不再重复做 ──────────
    // 背景：删空行可用 delete_row / clean_data+写回 / VBA；算指标可用 add_formula /
    // aggregate / date_group。按"工具对"写规则是组合爆炸且新工具必漏，因此改为按
    // 能力组 + 目标区域做通用去重：本轮同组且区域重叠 → 判定为重复劳动，自动跳过。
    const CAPABILITY_GROUPS = {
        // 删除/清洗组（会毁数据）
        'excel_delete_row': 'delete-clean',
        'excel_delete_column': 'delete-clean',
        'excel_clean_data': 'delete-clean',
        'excel_remove_duplicates': 'delete-clean',
        'excel_clear_range': 'delete-clean',
        'excel_clear_sheet': 'delete-clean',
        // 统计/计算组
        'excel_aggregate': 'stats',
        'excel_add_formula': 'stats',
        'excel_date_group': 'stats'
    };
    // 命中重复后自动跳过的组（写入/美化组不动，避免误伤正常分步流程）
    const AUTO_SKIP_GROUPS = { 'delete-clean': true, 'stats': true };

    const CAPABILITY_LOG = new WeakMap();  // turn -> Array<{toolName, group, sheet, rect}>

    /** 工具的能力组；VBA 按代码内容判定（含删除/清空语句 → delete-clean，否则不参与去重） */
    function capabilityGroupOf(toolName, args) {
        if (toolName === 'excel_execute_vba') {
            return isDestructiveVba(args && args.code) ? 'delete-clean' : null;
        }
        return CAPABILITY_GROUPS[toolName] || null;
    }

    /** "A1" / "$A$1:$G$201" → {r1,c1,r2,c2}；无法解析返回 null（视为整表） */
    function parseRangeRect(text) {
        if (!text) return null;
        const s = String(text).replace(/\$/g, '');
        const parts = s.split(':');
        const cell = (t) => {
            const m = /^([A-Za-z]+)?(\d+)?$/.exec(t.trim());
            if (!m) return null;
            let c = m[1] ? 0 : 1;
            if (m[1]) for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
            const r = m[2] ? Number(m[2]) : 1;
            return { r, c };
        };
        const a = cell(parts[0]);
        if (!a) return null;
        const b = parts[1] ? cell(parts[1]) : a;
        if (!b) return null;
        return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) };
    }

    /** 取本次调用的目标矩形；无范围信息返回 null（视为覆盖整表） */
    function targetRectOf(toolName, args) {
        if (!args) return null;
        if (args.range) return parseRangeRect(args.range);
        if (args.cell) return parseRangeRect(args.cell);
        if (args.address) return parseRangeRect(args.address);
        if (toolName === 'excel_delete_row' && args.row) {
            const r1 = Number(args.row), cnt = Number(args.count) || 1;
            return { r1, c1: 1, r2: r1 + cnt - 1, c2: 16384 };
        }
        if (toolName === 'excel_delete_column' && (args.column || args.col)) {
            const c1 = Number(args.column || args.col), cnt = Number(args.count) || 1;
            return { r1: 1, c1, r2: 1048576, c2: c1 + cnt - 1 };
        }
        return null;
    }

    function rectsOverlap(a, b) {
        if (!a || !b) return true;   // 任一方范围未知 → 保守判定为重叠
        return !(a.r2 < b.r1 || b.r2 < a.r1 || a.c2 < b.c1 || b.c2 < a.c1);
    }

    /** 记录一次已执行的同类操作 */
    function recordCapability(turn, toolName, group, sheet, rect) {
        if (!turn || typeof turn !== 'object' || !group) return;
        let list = CAPABILITY_LOG.get(turn);
        if (!list) { list = []; CAPABILITY_LOG.set(turn, list); }
        list.push({ toolName, group, sheet: String(sheet || '').toLowerCase(), rect });
    }

    /** 查找本轮是否已执行过同组且目标重叠的操作 */
    function findDuplicateCapability(turn, group, sheet, rect) {
        if (!turn || typeof turn !== 'object' || !group) return null;
        const list = CAPABILITY_LOG.get(turn);
        if (!list || !list.length) return null;
        const sh = String(sheet || '').toLowerCase();
        for (let i = list.length - 1; i >= 0; i--) {
            const it = list[i];
            if (it.group !== group) continue;
            if (it.sheet !== sh) continue;
            if (rectsOverlap(it.rect, rect)) return it;
        }
        return null;
    }


    /**
     * 守卫签名：
     * @callback ToolGuard
     * @param {Object} ctx - 守卫上下文
     * @param {string} ctx.toolName - 即将执行的工具名
     * @param {Object} ctx.args - 工具调用参数
     * @param {Object} [ctx.turn] - 当前 turn 对象
     * @returns {Promise<string|void|null>}
     *   - 返回 string: 拒绝执行，该字符串为拒绝原因（单调性：一旦被拒，后续守卫不可翻盘）
     *   - 返回 null/undefined/void: 允许执行
     */

    /**
     * 注册一个单调守卫。
     * 守卫按 priority 升序执行（小数字优先），同优先级按注册顺序。
     * 单调性保证：一旦某个守卫返回拒绝原因，后续守卫即使返回 null 也不可翻盘。
     * @param {string} name - 守卫名称（用于日志）
     * @param {ToolGuard} fn - 守卫函数
     * @param {number} [priority=100] - 优先级（默认 100）
     * @returns {Function} 取消注册的 disposer 函数
     */
    function guard(name, fn, priority = 100) {
        const entry = { name, fn, priority };
        let insertAt = _guards.length;
        for (let i = 0; i < _guards.length; i++) {
            if (_guards[i].priority > priority) {
                insertAt = i;
                break;
            }
        }
        _guards.splice(insertAt, 0, entry);
        ChatLog.info?.('TOOLS', `[guard] 注册守卫: ${name} (priority=${priority})`);
        return () => {
            const idx = _guards.indexOf(entry);
            if (idx >= 0) {
                _guards.splice(idx, 1);
                ChatLog.info?.('TOOLS', `[guard] 移除守卫: ${name}`);
            }
        };
    }

    /**
     * 运行守卫链（执行前调用）。
     * 单调性：一旦某个守卫返回拒绝原因，后续守卫即使返回 null 也不可翻盘。
     * @param {string} toolName - 即将执行的工具名
     * @param {Object} args - 工具调用参数
     * @param {Object} [ctx] - 执行上下文
     * @returns {Promise<string|null>} null=允许执行，string=拒绝原因
     */
    async function runGuards(toolName, args, ctx = {}) {
        let denied = null;
        for (const g of _guards) {
            try {
                const reason = await g.fn({
                    toolName,
                    args,
                    turn: ctx.turn || null
                });
                if (typeof reason === 'string') {
                    // 第一个拒绝即生效（单调性）
                    if (denied === null) {
                        denied = reason;
                        ChatLog.warn?.('TOOLS', `[guard] 守卫 ${g.name} 拒绝了 ${toolName}: ${reason}`);
                    }
                    // 即使已被拒绝，继续遍历剩余守卫（记录日志），但不翻盘
                }
            } catch (e) {
                ChatLog.warn?.('TOOLS', `[guard] 守卫 ${g.name} 异常:`, e);
            }
        }
        return denied;
    }

    /**
     * 从工具参数生成简短「将影响什么」预览（危险确认用）
     */
    function buildDangerArgsPreview(args) {
        if (!args || typeof args !== 'object') return '';
        const prefer = ['sheet', 'worksheet', 'range', 'address', 'dataRange',
            'filePath', 'path', 'outputPath', 'targetSheet', 'rows', 'startRow', 'endRow'];
        const parts = [];
        const push = (k, v) => {
            if (v == null || v === '') return;
            let s = typeof v === 'object' ? JSON.stringify(v) : String(v);
            if (s.length > 100) s = s.slice(0, 100) + '…';
            parts.push(k + '：' + s);
        };
        for (const k of prefer) {
            if (Object.prototype.hasOwnProperty.call(args, k)) push(k, args[k]);
            if (parts.length >= 6) break;
        }
        if (!parts.length) {
            for (const k of Object.keys(args)) {
                if (k.startsWith('_')) continue;
                push(k, args[k]);
                if (parts.length >= 5) break;
            }
        }
        return parts.join('\n');
    }

    /**
     * 危险操作确认对话框（Promise-based，WebView2 安全，z-index 高于一切模态）。
     * 当安全守卫在 quick/standard 模式拦截危险操作时弹出确认框：
     *   用户点"仍要继续执行" → 放行（返回 true，invoke 继续执行工具）；
     *   用户点"取消" / Esc / 点遮罩 → 拒绝（返回 false）。
     * @param {string} toolName 被拦截/即将执行的工具名
     * @param {string} reason 守卫返回的完整拒绝原因（含风险提示与严重程度）
     * @param {Object} [options]
     * @param {Object} [options.args] 工具参数（用于预览影响范围）
     * @returns {Promise<boolean>}
     */
    function confirmDangerousAction(toolName, reason, options) {
        // DOM 不可用（如无 headless 测试环境）或无法正常弹窗时 fail-safe 到 false（=取消），
        // 避免依赖异常兜底或导致调用方挂起。
        if (typeof document === 'undefined' || !document
            || typeof document.createElement !== 'function'
            || !document.body || typeof document.body.appendChild !== 'function') {
            return Promise.resolve(false);
        }
        const opts = options && typeof options === 'object' ? options : {};
        const preview = buildDangerArgsPreview(opts.args);
        return new Promise((resolve) => {
            const overlay = document.createElement('div');
            overlay.setAttribute('role', 'presentation');
            overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(20,20,30,0.55);z-index:100000;display:flex;align-items:center;justify-content:center;';
            const box = document.createElement('div');
            box.setAttribute('role', 'alertdialog');
            box.setAttribute('aria-modal', 'true');
            box.setAttribute('aria-labelledby', 'danger-confirm-title');
            box.style.cssText = 'background:#fff;border-radius:10px;padding:20px 22px;max-width:460px;width:92%;box-shadow:0 8px 30px rgba(0,0,0,0.25);font-family:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;';
            const escHtml = (s) => String(s)
                .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
            const tr = (k) => {
                try {
                    if (App && App.i18n && typeof App.i18n.t === 'function') return App.i18n.t(k);
                } catch (_e) { /* ignore */ }
                return k;
            };
            const alwaysAsk = isDangerAlwaysAsk();
            const hint = alwaysAsk ? tr('danger.hintAlwaysAsk') : tr('danger.hintDefault');
            const previewHtml = preview
                ? '<div style="font-size:12px;color:#64748b;margin-bottom:6px;font-weight:600;">' + escHtml(tr('danger.preview')) + '</div>' +
                  '<div style="font-size:12px;color:#334155;line-height:1.55;background:#f8fafc;border:1px solid #e2e8f0;padding:8px 10px;border-radius:6px;margin-bottom:12px;white-space:pre-wrap;font-family:Consolas,monospace;">' +
                  escHtml(preview) + '</div>'
                : '';
            box.innerHTML =
                '<div id="danger-confirm-title" style="font-size:15px;font-weight:600;color:#b91c1c;margin-bottom:10px;">' + escHtml(tr('danger.title')) + '</div>' +
                '<div style="margin-bottom:8px;font-size:13px;font-weight:500;color:#334155;">' + escHtml(toolName) + '</div>' +
                '<div style="font-size:13px;color:#475569;line-height:1.6;background:#fef2f2;border-left:3px solid #ef4444;padding:9px 11px;border-radius:4px;margin-bottom:12px;">' + escHtml(reason) + '</div>' +
                previewHtml +
                '<div style="font-size:11px;color:#94a3b8;margin-bottom:14px;line-height:1.45;">' + escHtml(hint) + '</div>' +
                '<div style="display:flex;gap:10px;justify-content:flex-end;">' +
                    '<button type="button" data-ac="cancel" style="padding:7px 16px;border:1px solid #d1d5db;background:#fff;color:#374151;border-radius:6px;cursor:pointer;font-size:13px;">' + escHtml(tr('danger.cancel')) + '</button>' +
                    '<button type="button" data-ac="ok" style="padding:7px 16px;border:none;background:#dc2626;color:#fff;border-radius:6px;cursor:pointer;font-size:13px;">' + escHtml(tr('danger.continue')) + '</button>' +
                '</div>';
            overlay.appendChild(box);
            document.body.appendChild(overlay);
            let settled = false;
            let onKey = null;
            const done = (val) => {
                if (settled) return;
                settled = true;
                try {
                    if (onKey && typeof document.removeEventListener === 'function') {
                        document.removeEventListener('keydown', onKey, true);
                    }
                } catch (_e) { /* ignore */ }
                try { overlay.remove(); } catch (_e) { /* 已移除则忽略 */ }
                resolve(val);
            };
            onKey = (e) => {
                if (!e) return;
                if (e.key === 'Escape') {
                    try { e.preventDefault(); } catch (_e) { /* ignore */ }
                    done(false);
                }
            };
            const okBtn = overlay.querySelector('[data-ac="ok"]');
            const cancelBtn = overlay.querySelector('[data-ac="cancel"]');
            if (okBtn) okBtn.onclick = () => done(true);
            if (cancelBtn) cancelBtn.onclick = () => done(false);
            overlay.onclick = (e) => {
                if (e && e.target === overlay) done(false);
            };
            try { if (box && typeof box.addEventListener === 'function') box.addEventListener('click', (e) => e.stopPropagation()); } catch (_e) { /* ignore */ }
            try {
                if (typeof document.addEventListener === 'function') {
                    document.addEventListener('keydown', onKey, true);
                }
            } catch (_e) { /* ignore */ }
            try { cancelBtn && typeof cancelBtn.focus === 'function' && cancelBtn.focus(); } catch (_e) { /* ignore */ }
        });
    }

    /**
     * 判断 VBA 代码是否包含删除/清空类破坏性操作。
     * 供技能管线在直连 COM（excelBridge.ExecuteVBA，绕过工具守卫）执行前把关。
     * @param {string} code VBA 代码
     * @returns {boolean}
     */
    function isDestructiveVba(code) {
        if (!code) return false;
        // FormatConditions.Delete 只清"条件格式规则"，不清数据 → 不算破坏性。
        // 否则每次美化（带数据条/色阶）都会弹危险确认框、把确定性 VBA 拦下来。
        const stripped = String(code).replace(/\bFormatConditions\s*\.\s*Delete\b/gi, '');
        return /\.(Delete|Clear|ClearContents|ClearFormats|ClearComments|ClearHyperlinks|ClearNotes|ClearArrows)\b/i.test(stripped);
    }

    /**
     * Excel MCP 工具中文显示名：由 excel_tools.json 统一维护，
     * 经桥接服务器下发（mt.displayName），此处不再硬编码。
     */

    // ── 图片类工具（text-pipeline 不可见）─────────────────────────────
    // 只记录规范化名（单一权威），前缀别名（excel-local__ 与否）由 isImageOnlyTool
    // 在比对时消解。管道只声明"过滤策略"，不接触这些名字。
    const IMAGE_ONLY_TOOLS = ['excel_ocr_recognize'];

    // ── 工具卡片渲染意图注册表 ──────────────────────────────────────
    /**
     * 工具名 → { presentCall, presentResult } 映射。
     * presentCall(args): 返回工具调用时的卡片渲染意图
     * presentResult(args, result): 返回工具完成后的卡片渲染意图
     * 返回格式：{ card: 'generic'|'chart'|'terminal'|'diff'|'search', ... }
     * 不在此表中的工具使用默认通用渲染（里程碑三：Card Union）
     */
    // 图表类型枚举 → 中文名称映射（用于 UI 卡片显示）
    // 兼容 LLM 输出的 snake_case 和英文无前缀形态

    /** 从 args 中获取 chartType 枚举值（兼容 chart_type / chartType / 英文无前缀） */

    /**
     * 判断工具执行是否成功，用于生成反馈性输出（如"清空内容：是"）。
     * presentResult 的 result 形参存在两种调用形态：
     *   - invokeTracked 传入 invoke() 结果：{ ok, result, error, ... }（有 ok）
     *   - updateToolCard 传入 sessionStore patch：{ status, error, ... }（无 ok，靠 status==='completed'）
     * 因此做宽容判定：ok===true 或 status==='completed' 均视为成功。
     * @param {Object|undefined} result - presentResult 收到的 result 形参
     * @returns {boolean}
     */

    /**
     * 读取后端返回内容里实际报告完成的关键动作。
     * 用途：当某子操作因降级而未真正执行时，前端不展示该"未执行"项（而非显示为"否"虚报）。
     * 当前仅用于"取消合并并填充"：后端在确实展开填充时返回含关键词的文案，降级为仅取消合并时不含。
     * @param {Object|undefined} result
     * @param {RegExp} keyword - 判定"已实际执行"的关键词
     * @returns {boolean}
     */

    // 渲染器白名单（封闭枚举）：AI/技能可在工具结果里携带 view.type 选择渲染器，
    // 但只能从本集合中选。不在集合内（含拼写/近义）会被 _normalizeViewType 归一或忽略，
    // 避免重蹈 excel_clean_range 那样"模型编造不存在渲染器"的覆辙。
    const RENDERER_WHITELIST = {
        table: 'table',      // 表格
        json: 'json',        // 可折叠 JSON
        code: 'code',        // 代码块
        markdown: 'markdown',// Markdown
        text: 'text',        // 纯文本
        steps: 'steps'       // 分步清单
    };
    // view.type 的近义词归一（简写/英文/常见误拼）→ 白名单键
    const VIEW_TYPE_ALIASES = {
        tbl: 'table', grid: 'table', table: 'table',
        raw: 'json', jsonc: 'json', json: 'json',
        codeblock: 'code', snippet: 'code', code: 'code',
        md: 'markdown', markdown: 'markdown',
        plain: 'text', text: 'text',
        plan: 'steps', steps: 'steps', checklist: 'steps'
    };
    function _normalizeViewType(t) {
        if (!t || typeof t !== 'string') return null;
        const k = t.trim().toLowerCase();
        if (RENDERER_WHITELIST[k]) return k;
        const alias = VIEW_TYPE_ALIASES[k];
        if (alias) return alias;
        return null;
    }
    // 从工具结果里挖出 AI/技能显式声明的 view.type（浅挖两层，兼容常见包裹）
    function _digViewType(result) {
        const candidates = [result, result && result.result, result && result.outputs];
        for (const c of candidates) {
            if (!c || typeof c !== 'object') continue;
            let raw = null;
            if (c.view && typeof c.view === 'object' && c.view.type) raw = c.view.type;
            else if (typeof c.viewType === 'string') raw = c.viewType;
            if (raw) {
                const norm = _normalizeViewType(raw);
                if (!norm) {
                    ChatLog.warn?.('TOOLS', `[view] 未知渲染器 "${raw}"，已忽略（可用：${Object.keys(RENDERER_WHITELIST).join('/')}）`);
                }
                return norm;
            }
        }
        return null;
    }

    const TOOL_CARD_RENDERERS = {
        'excel_create_chart': {
            presentCall(args) {
                const chartType = resolveChartType(args);
                const typeLabel = App.ChatV2.toolRegistryUtil?.CHART_TYPE_CN[chartType] || '柱状图';
                return { card: 'chart', title: '创建图表', chartType: typeLabel, dataRange: args?.dataRange || '' };
            },
            presentResult(args, result) {
                const chartType = resolveChartType(args);
                const typeLabel = App.ChatV2.toolRegistryUtil?.CHART_TYPE_CN[chartType] || '柱状图';
                const title = args?.title || '图表';
                return { card: 'chart', title, chartType: typeLabel, dataRange: args?.dataRange || '', completed: true };
            }
        },
        'excel_write_range': {
            presentCall(args) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                return { card: 'generic', title: '写入数据', kind: 'edit', content: `目标: ${target}` };
            },
            presentResult(args, result) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                return {
                    card: 'generic', title: '写入完成', kind: 'edit', content: `数据已写入 ${target}`,
                    feedback: _presentSucceeded(result) ? [{ label: '写入数据', done: true }] : []
                };
            }
        },
        'excel_clear_range': {
            presentCall(args) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                const cc = args?.clearContent !== false;
                const cf = args?.clearFormat !== false;
                const what = (cc && cf) ? '内容与格式' : (cf ? '格式' : (cc ? '内容' : '—'));
                return { card: 'generic', title: '清空单元格', kind: 'edit', content: `目标: ${target}（清除${what}）` };
            },
            presentResult(args, result) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                // 反馈随实际意图（clearContent/clearFormat）动态生成，而非写死
                const cc = args?.clearContent !== false;
                const cf = args?.clearFormat !== false;
                const what = (cc && cf) ? '内容与格式' : (cf ? '格式' : (cc ? '内容' : '—'));
                return {
                    card: 'generic', title: '清空单元格', kind: 'edit', content: `已清除 ${target} 的${what}`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '清空内容', done: cc },
                        { label: '清除格式', done: cf }
                    ] : []
                };
            }
        },
        'excel_clear_sheet': {
            presentCall(args) {
                const sheet = args?.sheet || '当前工作表';
                return { card: 'generic', title: '清空工作表', kind: 'edit', content: `目标: ${sheet}` };
            },
            presentResult(args, result) {
                const sheet = args?.sheet || '当前工作表';
                return {
                    card: 'generic', title: '清空工作表', kind: 'edit', content: `已清空 ${sheet} 的内容和格式`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '清空内容', done: true },
                        { label: '清除格式', done: true }
                    ] : []
                };
            }
        },
        'excel_clear_format': {
            presentCall(args) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                return { card: 'generic', title: '清除格式', kind: 'edit', content: `目标: ${target}（仅清除格式，内容不变）` };
            },
            presentResult(args, result) {
                const target = args?.range ? `${args.sheet || ''}!${args.range}` : '选中区域';
                return {
                    card: 'generic', title: '清除格式', kind: 'edit',
                    content: `已清除 ${target} 的格式（内容保持不变）`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '清除格式', done: true },
                        { label: '清空内容', done: false }
                    ] : []
                };
            }
        },
        'excel_close_workbook': {
            presentCall(args) {
                const sc = args?.saveChanges !== false;
                return { card: 'generic', title: '关闭工作簿', kind: 'edit', content: `动作: ${sc ? '保存并关闭' : '关闭（不保存）'}` };
            },
            presentResult(args, result) {
                const sc = args?.saveChanges !== false;
                return {
                    card: 'generic', title: '关闭工作簿', kind: 'edit', content: `工作簿已关闭（${sc ? '保存了更改' : '未保存更改'}）`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '保存更改', done: sc },
                        { label: '不保存关闭', done: !sc }
                    ] : []
                };
            }
        },
        'excel_apply_conditional_format': {
            presentCall(args) {
                const lit = args?.markLiteral;
                const colorName = args?.markColorName || '黄色';
                if (lit) return { card: 'generic', title: '标记单元格', kind: 'edit', content: `把等于「${lit}」的单元格标成${colorName}` };
                return { card: 'generic', title: '条件格式', kind: 'edit', content: args?.formatType || '条件格式' };
            },
            presentResult(args, result) {
                const lit = args?.markLiteral;
                const colorName = args?.markColorName || '黄色';
                if (lit) {
                    return {
                        card: 'generic', title: '已标成' + colorName, kind: 'edit',
                        content: `等于「${lit}」的单元格已涂成${colorName}`,
                        feedback: _presentSucceeded(result) ? [{ label: '已标成' + colorName, done: true }] : []
                    };
                }
                return { card: 'generic', title: '条件格式已应用', kind: 'edit', content: '已应用到目标范围' };
            }
        },
        'excel_clean_data': {
            presentCall(args) {
                const wb = args?.writeBack === true || args?.writeBack === 'true';
                return { card: 'generic', title: wb ? '清洗并写回' : '清洗预览', kind: wb ? 'edit' : 'other', content: wb ? '将直接修改工作表' : '仅预览，不改工作表' };
            },
            presentResult(args, result) {
                const wb = args?.writeBack === true || args?.writeBack === 'true';
                return {
                    card: 'generic', title: wb ? '已写回工作表' : '仅预览', kind: wb ? 'edit' : 'other',
                    content: wb ? '清洗结果已写回当前工作表' : '仅计算，未写回 Excel',
                    feedback: _presentSucceeded(result) ? [{ label: wb ? '已写回工作表' : '仅预览', done: true }] : []
                };
            }
        },
        'excel_remove_duplicates': {
            presentCall(args) {
                const cols = Array.isArray(args?.columns) && args.columns.length ? args.columns.join('、') : '全部列';
                const wb = args?.writeBack === true || args?.writeBack === 'true';
                const mark = args?.markOnly === true;
                const content = mark ? `按 ${cols} 把重复行标黄，不删除` : (wb ? `按 ${cols} 删除重复行并写回` : `按 ${cols} 返回去重后的数据`);
                return { card: 'generic', title: mark ? '标记重复行' : '按列去重', kind: (wb || mark) ? 'edit' : 'other', content };
            },
            presentResult(args, result) {
                const wb = args?.writeBack === true || args?.writeBack === 'true';
                const mark = args?.markOnly === true;
                return {
                    card: 'generic',
                    title: mark ? '已标黄重复行' : (wb ? '已写回工作表' : '去重完成'),
                    kind: (wb || mark) ? 'edit' : 'other',
                    content: mark ? '重复行已标黄，数据未删除' : (wb ? '重复行已从工作表删除' : '已返回去重后的唯一数据（仅计算，未写回 Excel）'),
                    feedback: _presentSucceeded(result) ? [{ label: mark ? '已标黄' : (wb ? '已写回工作表' : '返回去重数据'), done: true }] : []
                };
            }
        },
        'excel_merge_cells': {
            presentCall(args) {
                const range = args?.range || '选中区域';
                const c = args?.mergeType === 'mergeAndCenter';
                return { card: 'generic', title: '合并单元格', kind: 'edit', content: `${range}（${c ? '合并后居中' : '仅合并'}）` };
            },
            presentResult(args, result) {
                const range = args?.range || '选中区域';
                const c = args?.mergeType === 'mergeAndCenter';
                return {
                    card: 'generic', title: '合并单元格', kind: 'edit', content: `已合并 ${range}（${c ? '文字居中' : '未改对齐'}）`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '合并单元格', done: true },
                        { label: '合并后居中', done: c }
                    ] : []
                };
            }
        },
        'excel_unmerge_cells': {
            presentCall(args) {
                const range = args?.range || '选中区域';
                const fill = args?.fillAfterUnmerge === true;
                return { card: 'generic', title: '取消合并单元格', kind: 'edit', content: `${range}（取消合并${fill ? '并展开填充值' : ''}）` };
            },
            presentResult(args, result) {
                const range = args?.range || '选中区域';
                const fillWanted = args?.fillAfterUnmerge === true;
                // 仅当后端确实执行了展开填充才展示该项；降级为仅取消合并时不展示（避免虚报未执行的动作）
                const fillExecuted = fillWanted && _backendReported(result, /展开填充到整个区域/);
                const items = [{ label: '取消合并', done: true }];
                if (fillExecuted) items.push({ label: '展开填充值', done: true });
                return {
                    card: 'generic', title: '取消合并单元格', kind: 'edit',
                    content: `已取消合并 ${range}${fillExecuted ? '，并将值展开填充到整个区域' : ''}`,
                    feedback: _presentSucceeded(result) ? items : []
                };
            }
        },
        'excel_split_by_rows': {
            presentCall(args) {
                const mode = args?.outputMode === 'files' ? '拆分为文件' : '拆分到工作表';
                const per = args?.rowsPerFile || 500;
                return { card: 'generic', title: '按行拆分', kind: 'edit', content: `${mode}，每份 ${per} 行` };
            },
            presentResult(args, result) {
                const mode = args?.outputMode === 'files' ? '已拆分为独立文件' : '已拆分到新工作表';
                return { card: 'generic', title: '按行拆分', kind: 'edit', content: `${mode} 完成` };
            }
        },
        'excel_sort_range': {
            presentCall(args) {
                const keys = describeSortKeys(args?.sortBy);
                const hdr = args?.hasHeader === false ? '不含表头' : '含表头';
                return { card: 'generic', title: '排序', kind: 'edit', content: `${keys}（${hdr}）` };
            },
            presentResult(args, result) {
                const keys = describeSortKeys(args?.sortBy);
                let rowNote = '';
                try {
                    const r = typeof result === 'string' ? JSON.parse(result) : result;
                    const u = (r && typeof r === 'object') ? (r.data && typeof r.data === 'object' ? r.data : r) : null;
                    if (u && (u.rowCount != null || u.headers)) {
                        rowNote = u.rowCount != null ? `，共 ${u.rowCount} 行` : '';
                    }
                } catch (e) { /* 非 JSON 结果不附加行数 */ }
                return {
                    card: 'generic', title: '排序', kind: 'edit', content: `已按 ${keys} 排序${rowNote}`,
                    feedback: _presentSucceeded(result) ? [
                        { label: '排序依据', done: true, note: keys },
                        { label: '忽略表头', done: args?.hasHeader === false }
                    ] : []
                };
            }
        },
        'excel_read_range': {
            presentCall(args) {
                const range = args?.range || '选中区域';
                return { card: 'generic', title: '读取数据', kind: 'read', content: `范围: ${range}` };
            },
            presentResult(args) {
                const range = args?.range || '选中区域';
                return { card: 'generic', title: '读取完成', kind: 'read', content: `已读取 ${range} 的数据` };
            }
        },
        'excel_execute_vba': {
            presentCall(args) {
                const codePreview = (args?.code || '').substring(0, 80);
                return { card: 'terminal', title: '执行 VBA', description: codePreview };
            },
            presentResult(args) {
                return { card: 'terminal', title: 'VBA 执行完成', output: '' };
            }
        },
        'excel_aggregate': {
            presentCall(args) {
                const dim = args?.dimColumn || args?.groupBy || '';
                const agg = args?.aggFunc || 'sum';
                return { card: 'generic', title: '聚合分析', kind: 'other', content: `按 ${dim} 汇总 (${agg})` };
            },
            presentResult(args) {
                const dim = args?.dimColumn || args?.groupBy || '';
                return { card: 'generic', title: '聚合完成', kind: 'other', content: `按 ${dim} 的聚合结果已返回（仅计算，未生成图表/工作表）` };
            }
        },
        'excel_date_group': {
            presentCall(args) {
                const col = args?.dateColumn || '';
                const g = args?.groupBy || 'month';
                const gMap = { year: '年', quarter: '季度', month: '月', week: '周', day: '日' };
                return { card: 'generic', title: '日期分组', kind: 'other', content: `按${gMap[g] || g}分组「${col}」` };
            },
            presentResult(args) {
                const g = args?.groupBy || 'month';
                const gMap = { year: '年', quarter: '季度', month: '月', week: '周', day: '日' };
                return { card: 'generic', title: '分组完成', kind: 'other', content: `已按${gMap[g] || g}完成日期分组` };
            }
        }
    };

    /**
     * 规范化工具名为符合 LLM function calling 规范的格式
     * 规则：必须以字母开头，只能包含字母、数字、下划线和连字符
     * 例如：'skill.SmartDataCleaner' → 'skill_SmartDataCleaner'
     *       'mcp:excel.read' → 'mcp_excel_read'
     *       '123foo' → 'fn_123foo'
     */

    /**
     * 透明化补全参数：对声明了 range/sheet 的选区类工具（如 excel_clear_range），
     * 用当前 Turn 的上下文（当前选区 / 活动工作表）补全缺失参数，
     * 使操作明确指向具体工作表与单元格区域，而非依赖后端隐性回退，提升操作透明度。
     * 仅当调用方/模型未提供该参数时才补全，绝不覆盖显式传入值。
     * @param {string} toolName - 工具名
     * @param {Object} args - 当前生效参数（可能是 effectiveArgs / 原始 args）
     * @param {Object} turn - turn 对象（含 context.selection / context.activeSheet）
     * @returns {Object} 需要补全的字段（如 { range, sheet }）
     */
    function fillSelectionArgs(toolName, args, turn) {
        const tool = _tools.get(normalizeName(toolName)) || _tools.get(toolName);
        if (!tool || !turn || !turn.context) return {};
        let paramNames = [];
        if (tool.inputSchema && tool.inputSchema.properties) {
            paramNames = Object.keys(tool.inputSchema.properties);
        } else if (Array.isArray(tool.inputSchema)) {
            paramNames = tool.inputSchema.map(p => p && p.name).filter(Boolean);
        } else if (Array.isArray(tool.parameters)) {
            paramNames = tool.parameters.map(p => p && p.name).filter(Boolean);
        }
        const sel = turn.context.selection;
        const ctxSheet = turn.context.activeSheet;
        const hasStr = (o, key) => typeof o === 'object' && o != null && typeof o[key] === 'string' && o[key].trim() !== '';
        const fill = {};
        const _isFormat = normalizeName(toolName) === 'excel_format_range';
        const _plan = (_isFormat && App.ChatV2.rangeResolver && typeof App.ChatV2.rangeResolver.decideFormatTarget === 'function')
            ? App.ChatV2.rangeResolver.decideFormatTarget(turn.userMessage || turn.message || '', turn.context)
            : null;
        if (paramNames.includes('range') && !hasStr(args, 'range')) {
            if (_plan) {
                // 光标单格、自动捕获、别的表上的选区，都不能填进「每张表自己的数据区」。
                const _useSel = _plan.rangeMode === 'selection' || _plan.rangeMode === 'cell' || _plan.rangeMode === 'copy-selection';
                if (_plan.rangeMode === 'spoken' && _plan.spokenRange) fill.range = _plan.spokenRange;
                else if (_useSel && _plan.selectionRange) fill.range = _plan.selectionRange;
            } else {
                const addr = (sel && typeof sel.address === 'string' && sel.address) ? sel.address
                    : (sel && typeof sel.range === 'string' && sel.range) ? sel.range : null;
                if (addr) fill.range = addr;
            }
        }
        if (paramNames.includes('sheet') && !hasStr(args, 'sheet')) {
            if (_plan && (_plan.sheetScope === 'all' || _plan.sheetScope === 'others')) {
                // 口头要每张表：不要把活动表名填进去，否则一次调用看起来只做了当前表。
            } else if (_plan && _plan.sheetScope === 'named' && _plan.namedSheets.length === 1) {
                fill.sheet = _plan.namedSheets[0];
            } else if (_plan && (_plan.rangeMode === 'selection' || _plan.rangeMode === 'cell') && _plan.selectionSheet) {
                fill.sheet = _plan.selectionSheet;
            } else if (ctxSheet) {
                fill.sheet = ctxSheet;
            }
        }
        // 扩展：复制类工具(excel_copy_range)的 sourceSheet/sourceRange 补全
        // 使用 turn.context 中捕获的原始选区信息（创建新表后活动表已变，但 turn.context 保存了请求时的状态）
        if (paramNames.includes('sourceSheet') && !hasStr(args, 'sourceSheet')) {
            const srcSheet = ctxSheet && typeof ctxSheet === 'object' ? (ctxSheet.sheet || ctxSheet.Name || '') : ctxSheet || '';
            if (srcSheet) fill.sourceSheet = srcSheet;
        }
        if (paramNames.includes('sourceRange') && !hasStr(args, 'sourceRange')) {
            const addr = (sel && typeof sel.address === 'string' && sel.address) ? sel.address
                : (sel && typeof sel.range === 'string' && sel.range) ? sel.range : null;
            if (addr) {
                // 去除工作表限定符（如 "Sheet1!A1:B2" → "A1:B2"）
                const cleanAddr = addr.includes('!') ? addr.split('!').pop().trim() : addr;
                if (/^[A-Za-z]+\d+/.test(cleanAddr)) fill.sourceRange = cleanAddr;
            }
        }
        // 图表数据源：excel_create_chart 声明的是 dataRange（非 range），
        // 未显式传 dataRange 且未用 dataValues（内存数组模式，与 dataRange 二选一）时，
        // 用当前选中区域补全，避免"对选中区域生成图表"被必填校验拦截。
        if (paramNames.includes('dataRange') && !hasStr(args, 'dataRange') && !hasStr(args, 'dataValues')) {
            const addr = (sel && typeof sel.address === 'string' && sel.address) ? sel.address
                : (sel && typeof sel.range === 'string' && sel.range) ? sel.range : null;
            if (addr) fill.dataRange = addr;
        }
        return fill;
    }

    /**
     * 把异常的 sheet 参数归一化为字符串。
     * 不同模型可能把 sheet 传成上下文对象（如 {sheet:"员工提成计算表", workbook:"..."}），
     * 若不归一化，模板拼接会出现 "[object Object]" 且后端 GetArg(String) 也取不到正确值。
     * @param {Object|null} args - 将就地改写其 sheet 字段
     */

    /**
     * 必填参数安全默认生成器（推断优先、兜底重生成本）。
     * 仅收录"缺了也明显该自动补一个合理值"的参数，避免对无安全默认的参数盲目造值扩面。
     * 键名格式：<工具规范化名>.<必填参数键>
     */
    const SAFE_DEFAULT_GENERATORS = {
        'excel_create_sheet.name': defaultNewSheetName,
        'excel_create_sheet.sheetName': defaultNewSheetName,
        'excel_create_sheet.sheet_name': defaultNewSheetName,
        'excel_aggregate.dimColumn': inferDimColumnFromMessage,
        'excel_aggregate.dim_column': inferDimColumnFromMessage,
        'excel_format_range.format': () => '{}',
        // 图表类型：未指定时兜底为 auto，交由后端启发式自动选型（原子工具的默认决策路径）
        'excel_create_chart.chartType': () => 'auto',
        // 导出 CSV 的保存路径：用户说"保存到桌面/下载"或未指定时自动落盘（支持异步生成器）
        'excel_export_csv.filePath': inferExportFilePath,
        'excel_export_csv.file_path': inferExportFilePath,
        'excel_export_csv.savePath': inferExportFilePath,
        'excel_export_csv.save_path': inferExportFilePath,
        // PPT/Word 导出（PDF/图片）：未指定路径时**前端就把真实桌面路径算出来**。
        // 【为什么不能只靠 BACKEND_INFERRED_PARAMS 留空给后端】宿主 DLL 不一定随源码重编译，
        // 旧版 Bridge.ExportPdf 仍有 `filePath 不能为空` 硬校验 → 工具直接失败，
        // 原始校验语还会被当成结果甩给用户（实测 2026-10-05 23:00，PPT 导出 PDF 连报两次）。
        // 前端补齐后新旧后端都能落盘；BACKEND_INFERRED_PARAMS 保留作第二层兜底。
        'ppt_export_pdf.filePath': (t) => inferHostExportFilePath(t, { ext: 'pdf', base: '演示文稿' }),
        'ppt_export_pdf.file_path': (t) => inferHostExportFilePath(t, { ext: 'pdf', base: '演示文稿' }),
        'word_export_pdf.filePath': (t) => inferHostExportFilePath(t, { ext: 'pdf', base: '文档' }),
        'word_export_pdf.file_path': (t) => inferHostExportFilePath(t, { ext: 'pdf', base: '文档' }),
        'ppt_export_image.filePath': (t) => inferHostExportFilePath(t, { ext: 'png', base: '幻灯片' }),
        'ppt_export_image.file_path': (t) => inferHostExportFilePath(t, { ext: 'png', base: '幻灯片' }),
    };

    /**
     * 后端可自动推断的参数：生成器/上下文推断均失败时，留空交给后端兜底，不反问用户。
     * 例：excel_aggregate.dimColumn 由后端 dimensionFallback 从数据自动选分类列，
     *     "找出卖得好的产品"这类未指定分组列的请求不应被拦截，而应让后端推断。
     */
    const BACKEND_INFERRED_PARAMS = {
        'excel_aggregate.dimColumn': true,
        'excel_aggregate.dim_column': true,
        'excel_aggregate.dimField': true,
        'excel_aggregate.groupBy': true,
        // PPT / Word 导出 PDF：filePath 留空或为占位符路径时，后端自动落到真实桌面
        'ppt_export_pdf.filePath': true,
        'ppt_export_pdf.file_path': true,
        'word_export_pdf.filePath': true,
        'word_export_pdf.file_path': true,
        // PPT 导出图片：filePath 留空或为占位符路径时，后端自动落到真实桌面
        'ppt_export_image.filePath': true,
        'ppt_export_image.file_path': true,
    };

    /**
     * 可选参数的前端推断填充白名单（键 = 工具规范化名，值 = 参数名数组）。
     * 背景（实测 2026-10-06 00:26）：ppt_export_pdf 的 filePath 已改为可选（留空由后端
     * 自动落盘桌面），但 SAFE_DEFAULT_GENERATORS 只对"必填参数"生效 → 前端不填 →
     * 工具卡片的"参数"区显示 `filePath:`（空），用户看不出文件会存到哪。
     * 对白名单内的可选参数也跑 SAFE_DEFAULT_GENERATORS：前端把真实桌面路径填进参数，
     * 卡片"参数"区直接显示完整路径；后端的留空兜底逻辑保留不变（双保险）。
     */
    const OPTIONAL_INFERRED_PARAMS = {
        ppt_export_pdf: ['filePath', 'file_path'],
        word_export_pdf: ['filePath', 'file_path'],
        ppt_export_image: ['filePath', 'file_path'],
    };

    /** 从用户消息提取分组维度列名（如"按部门汇总" → "部门"） */

    /** 从用户消息提取时间分组粒度（如"按年汇总" → "year"，"按月统计" → "month"） */

    /** 从 turn.context 提取已存在的工作表名集合，用于生成不与现有表重名的默认表名 */

    /** 从用户消息提取显式声明的目标工作表名（名为/命名为/叫/叫做 等）。提取失败返回 null */


    /** 判断一个值是否"已提供"（非 undefined/null/空串） */

    /** 解出工具的必填参数键数组（兼容 inputSchema{properties,required} 与 inputSchema数组/parameters数组） */

    /** 参数的友好描述（用于反问措辞） */

    /**
     * 从会话上下文推断缺失的参数值。
     * 策略：
     *   1. "source/fromSheet" 类参数 → 使用活动工作表名
     *   2. "sheet/sheetName/worksheet/target" 类参数 → 从会话历史中最近创建的工作表推断
     */
    function inferFromContext(name, pk, ctxTurn) {
        const ctx = ctxTurn?.context || {};

        // 1) 参数名含"源"语义 → 使用活动工作表
        if (pk === 'source' || pk === 'fromSheet' || pk === 'from_sheet' || pk === 'sourceSheet') {
            const active = ctx.activeSheet?.sheet || ctx.activeSheet?.Name || '';
            if (active) return active;
        }

        // 1.5) 工作表管理工具的目标表名：excel_delete_sheet.name / excel_rename_sheet.oldName。
        //      优先取用户消息里显式写的表名（如 `删除工作表"异常场景测试表"`），其次回落到活动工作表。
        //      实测缺此分支：`删除工作表"异常场景测试表"` 的 name 无法补全 → 判为"缺参数"，
        //      又因 excel_delete_sheet 仅被 dashboard 域 guardedTools 收录、消息不含看板关键词，
        //      被"无关调用"护栏静默降级为失败 —— 用户既没被执行、也没被反问。
        if ((name === 'excel_delete_sheet' && pk === 'name') ||
            (name === 'excel_rename_sheet' && pk === 'oldName')) {
            const fromMsg = extractSheetNameFromMessage(ctxTurn);
            if (fromMsg) return fromMsg;
            const active = ctx.activeSheet?.sheet || ctx.activeSheet?.Name || '';
            if (active) return active;
        }

        // 2) 参数名含"目标/表名"语义 → 尝试从会话历史中最近创建的工作表推断
        if (pk === 'sheet' || pk === 'sheetName' || pk === 'sheet_name' || pk === 'target' || pk === 'destinationSheet') {
            const fromMsg = extractSheetNameFromMessage(ctxTurn);
            if (fromMsg) return fromMsg;
            try {
                const ss = App.ChatV2?.sessionStore;
                if (ss && typeof ss.getHistory === 'function') {
                    const history = ss.getHistory();
                    if (Array.isArray(history)) {
                        for (let i = history.length - 1; i >= 0; i--) {
                            const h = history[i];
                            if (h.role !== 'assistant' || !Array.isArray(h.toolCalls)) continue;
                            for (const tc of h.toolCalls) {
                                if (tc.name !== 'excel_create_sheet' || tc.status !== 'completed') continue;
                                // 从 MCP 结果文本中提取表名
                                const result = tc.result;
                                if (result && result.content && Array.isArray(result.content)) {
                                    for (const item of result.content) {
                                        const text = item.text || item.Text || '';
                                        const m = text.match(/成功创建工作表[：:]\s*(.+)/);
                                        if (m && m[1]) return m[1].trim();
                                    }
                                }
                                // 从工具参数中提取表名
                                if (tc.args?.name) return tc.args.name;
                                if (tc.args?.sheetName) return tc.args.sheetName;
                            }
                        }
                    }
                }
            } catch (e) { /* 忽略 */ }
        }

        return null;
    }

    /**
     * 必填参数兜底：能安全推断/默认的补全；否则收集为"需用户补充"。
     * 与 fillSelectionArgs 互补：fillSelectionArgs 已补全选区/工作表，此函数处理其余必填参数。
     * @param {string} name - 工具名（原始或规范化）
     * @param {Object} args - 当前生效参数（就地可能不回写，仅返回 added）
     * @param {Object} ctxTurn - turn 对象（含 context）
     * @returns {{added:Object, missing:Array<{parameter:string,label:string,question:string}>}}
     */
    async function completeRequiredArgs(name, args, ctxTurn) {
        const key = normalizeName(name);
        const tool = _tools.get(key) || _tools.get(name);
        if (!tool || !ctxTurn) return { added: {}, missing: [] };
        const required = _requiredParamKeys(tool);

        const added = {};
        const missing = [];
        for (const pk of required) {
            // 已有值即尊重——**除非**它是模型编造的占位符路径（如 C:\Users\Yourusername\Desktop\output.csv）。
            // 这类值非空却不可信：若直接放过，下面的 SAFE_DEFAULT_GENERATORS（如 inferExportFilePath）
            // 会被整体跳过，后端 CreateDirectory 就会把文件写进编造出来的目录（已实测发生）。
            // 【收窄适用范围】占位符规则是面向【路径】设计的（花括号/省略号等），
            // 不能对所有必填参数一律套用：像 excel_run_python_script.script 这类代码/文本参数
            // 天然含 {...}、... ，会被误判成编造值而被清空反问，脚本根本执行不了
            // （实测 2026-10-09：script 已传入却被判"缺少必要参数：请提供script"）。
            // 故只对路径型参数名 / 值本身像绝对路径的参数启用检测。
            if (_hasValue(args[pk]) && !isPlaceholderPathForParam(pk, args[pk])) continue;
            // 1) 用户消息推断（可扩展；目前默认生成器已覆盖创建类命名）
            // 2) 安全默认生成器
            const gen = SAFE_DEFAULT_GENERATORS[key + '.' + pk];
            let val = null;
            if (typeof gen === 'function') {
                try {
                    val = gen(ctxTurn);
                    // 支持异步生成器（如 excel_export_csv.filePath 需 await 桥接取桌面路径）
                    if (val && typeof val.then === 'function') val = await val;
                } catch (e) { val = null; }
            }
            if (_hasValue(val)) {
                added[pk] = val;
                continue;
            }
            // 2.5) 上下文推断：从 Excel 当前状态（活动工作表）和会话历史（最近创建的工作表）推断参数
            if (!_hasValue(val)) {
                try { val = inferFromContext(key, pk, ctxTurn); } catch (e) { val = null; }
            }
            if (_hasValue(val)) {
                added[pk] = val;
                continue;
            }
            // 2.6) 后端可自动推断的参数：留空交给后端兜底，不反问用户
            if (BACKEND_INFERRED_PARAMS[key + '.' + pk]) continue;
            // 格式化缺 range 时不要补 autoRange。autoRange 会落到活动表的 UsedRange，
            // 未激活的表经常是光标单格。范围由 decideFormatTarget 决定：口头地址、真选区，
            // 或调用前按该表数据区解析。这里解析不了就仍然算缺 range。
            // 2.7) 图表内存数组模式：dataValues 与 dataRange 二选一（与后端 schema 语义一致），
            // 已提供 dataValues（内存二维数组）时 dataRange 不再视为缺失。
            if (key === 'excel_create_chart' && pk === 'dataRange' && _hasValue(args.dataValues)) continue;
            // 3) 无法兜底 → 收集为需用户补充
            const label = _paramLabel(tool, pk);
            missing.push({ parameter: pk, label, question: '请提供' + label + '；提供后再重试。' });
        }

        // 【可选参数的前端推断填充】白名单内的可选参数（导出类 filePath）留空时也跑
        // SAFE_DEFAULT_GENERATORS，把真实桌面路径填进参数 → 卡片"参数"区显示完整路径。
        // 生成失败/目录拿不到时留空交后端兜底，不反问、不拦截。
        const optParams = OPTIONAL_INFERRED_PARAMS[key];
        if (optParams && optParams.length) {
            for (const pk of optParams) {
                if (_hasValue(args[pk]) && !isPlaceholderPath(args[pk])) continue;
                const gen = SAFE_DEFAULT_GENERATORS[key + '.' + pk];
                if (typeof gen !== 'function') continue;
                try {
                    let val = gen(ctxTurn);
                    if (val && typeof val.then === 'function') val = await val;
                    if (_hasValue(val)) added[pk] = val;
                } catch (e) { /* 留空交后端兜底 */ }
            }
        }
        return { added, missing };
    }

    /**
     * 解析"清空"类工具的意图粒度（清空内容 vs 清除格式）。
     * 【硬路由兜底】模型常只口述意图却不传 clearContent/clearFormat（甚至自造 excel_clear_format 等名），
     * 直接按当前用户消息的关键词判定并覆盖模型漏传/误传，避免"清空格式"误删单元格内容。
     * 仅当用户消息能明确判定为"仅格式"或"仅内容"时才改写；否则返回 null（交由模型参数与后端默认"全部清除"）。
     * @param {Object} turn - 当前 turn（含 userMessage）
     * @returns {{clearContent:boolean, clearFormat:boolean}|null}
     */

    /**
     * "多项操作"工具的意图归属参数表：工具名 → [应随用户意图装配的参数]。规划者（模型）若漏传这些参数，
     * 则按用户消息补全；已显式传入的参数一律尊重、不覆盖（"每个动作是一个零件，只装配需要的"）。
     */

    /**
     * 按当前用户消息为"多项操作"工具装配缺省的粒度参数（零件装配）。
     * @param {Object} turn - 当前 turn（含 userMessage）
     * @param {string} name - 工具名
     * @returns {Object|null} 需并入 args 的参数对；无明确意图时返回 null
     */
    // 把 LLM 输出的"通用/蛇形"参数形态，翻译成后端 MCP 工具期望的"驼峰"参数名。
    // 否则如 chart_type/data_source/include_legend 等参数不会被后端识别，
    // 工具会命中默认值（如柱状图）导致规划与执行不一致。
    const TOOL_ARG_ADAPTERS = {
        'excel_create_chart': normalizeChartArgs,
        'excel_add_shape': normalizeShapeArgs,
        // LLM 常输出 file_path/savePath/save_path/path，后端只认 filePath →
        // 归一化后必填参数校验才能命中，避免"缺少必要参数"拦截
        'excel_export_csv': (...a) => App.ChatV2.toolRegistryUtil?.['normalizeExportCsvArgs']?.(...a)
    };



    /** 解析 "$B$2:$B$74" / "Sheet1!B2:B74" 地址 → {sheet,col1,row1,col2,row2} */

    /** 合并 category_source(标签列) 与 data_source(数值列) 为后端可用的单一 dataRange */

    /** 图表类型值归一化：pie/柱状图/xlPie 等 → 后端接受的枚举/别名 */

    /** excel_create_chart 参数适配器：chart_type→chartType, data_source/category_source→dataRange, include_legend→hasLegend */

    /** 从用户消息推断图表标题（"按XX汇总"→"按XX汇总"；否则用图表类型中文名） */

    /** 从用户消息推断纵轴（数值列）标题 */

    /**
     * 图表工具智能默认值：未显式提供时补合理默认，避免图表空白/简陋或被参数卡死。
     * 仅填充缺失参数；用户/规划者显式传入的一律尊重、不覆盖。
     */

    /**
     * 注册一个 post-execute 钩子。
     * 钩子按 priority 升序执行（小数字优先），同优先级按注册顺序。
     * @param {string} name - 钩子名称（用于日志）
     * @param {PostExecuteHook} fn - 钩子函数
     * @param {number} [priority=100] - 优先级（默认 100）
     * @returns {Function} 取消注册的 disposer 函数
     */
    function onPostExecute(name, fn, priority = 100) {
        const entry = { name, fn, priority };
        // 按 priority 插入排序
        let insertAt = _postExecuteHooks.length;
        for (let i = 0; i < _postExecuteHooks.length; i++) {
            if (_postExecuteHooks[i].priority > priority) {
                insertAt = i;
                break;
            }
        }
        _postExecuteHooks.splice(insertAt, 0, entry);
        ChatLog.info?.('TOOLS', `[post-execute] 注册钩子: ${name} (priority=${priority})`);
        return () => {
            const idx = _postExecuteHooks.indexOf(entry);
            if (idx >= 0) {
                _postExecuteHooks.splice(idx, 1);
                ChatLog.info?.('TOOLS', `[post-execute] 移除钩子: ${name}`);
            }
        };
    }

    /**
     * 运行 post-execute 钩子链。
     * 每个钩子可以 accept（保持结果）或 replace（替换结果）。
     * 钩子不应抛出异常——异常会被捕获并记录，不影响主流程。
     * @param {string} toolName - 已执行的工具名
     * @param {Object} args - 工具调用参数
     * @param {Object} result - 当前执行结果（{ok, result, error, durationMs}）
     * @param {Object} [ctx] - 执行上下文（turn, sessionStore 等）
     * @returns {Promise<Object>} 最终结果（可能被钩子替换）
     */
    async function runPostExecuteHooks(toolName, args, result, ctx = {}) {
        let current = { ...result };
        for (const hook of _postExecuteHooks) {
            try {
                const decision = await hook.fn({
                    toolName,
                    args,
                    result: current,
                    turn: ctx.turn || null
                });
                if (decision && decision.kind === 'replace' && decision.payload) {
                    ChatLog.info?.('TOOLS', `[post-execute] 钩子 ${hook.name} 替换了 ${toolName} 的结果`);
                    current = { ...current, ...decision.payload };
                }
                // kind === 'accept' 或 undefined/void：不做变更
            } catch (e) {
                ChatLog.warn?.('TOOLS', `[post-execute] 钩子 ${hook.name} 异常:`, e);
            }
        }
        return current;
    }

    /**
     * Tool 接口契约：
     * {
     *   name: string,           // 唯一名称（规范化后），如 'excel_read_range' / 'skill_SmartDataCleaner'
     *   originalName: string,   // 原始名称（用于内部调用），如 'skill.SmartDataCleaner'
     *   description: string,    // 给 LLM 看的描述
     *   inputSchema: object,    // JSON Schema
     *   source: 'mcp'|'skill'|'builtin',
     *   execute(args, ctx): Promise<{ ok: boolean, result: any, error?: string }>
     * }
     */

    /**
     * 错误分类器：将原始错误消息分类为结构化错误对象，包含类型、原因和修复建议
     * @param {string|object} rawError - 原始错误
     * @param {string} toolName - 工具名
     * @returns {{type: string, reason: string, suggestion: string}|null}
     */

    /**
     * 规范化 MCP 工具返回结果，识别成功/失败
     * MCP 后端返回格式多样：{result: ...}、{success:false, message:"..."}、{error:"..."}、纯字符串错误等
     */

    /**
     * 同一 turn 内自动分配下一个步骤号：
     * 取现有 toolCalls 中最大 round + 1；无任何调用时为 1。
     * invokeTracked 在 round 未显式指定时调用，各管线无需再手动维护 nextRound 计数器。
     */

    // ── 工具契约（Tool Contract）──────────────────────────────────
    // 参考 DSH ToolDefinition 设计：每个工具可声明
    //   outputSchema     — 规范输出声明（JSON Schema 子集），invoke 后校验
    //   finalizeContent  — 最后一英里内容归一化（同步回调，总函数，不得抛异常）
    //   presentCall      — 调用中状态的 UI 投影（纯函数，可回放）
    //   presentResult    — 完成状态的 UI 投影（纯函数，可回放）
    // 所有字段均为可选；未声明时行为与旧版完全一致，保证向后兼容。

    /**
     * 轻量 JSON Schema 校验器（支持 DSH assertObjectJsonSchema 的实用子集）：
     *   type / required / properties / items / enum / const / oneOf / additionalProperties
     * @param {Object} schema - JSON Schema 节点
     * @param {*} value - 待校验值
     * @returns {{valid: boolean, errors: string[]}}
     */

    /**
     * 工具结果最后一英里归一化：调用工具声明的 finalizeContent（若有）。
     * 与 DSH 约定一致：回调必须是总函数（total），不得抛异常；
     * 返回 undefined 表示保留原内容，返回其他值则替换 content。
     * @param {Object} tool - 工具定义
     * @param {Object} args - 调用参数
     * @param {Object} result - 规范化结果 {ok, result, error, durationMs}
     * @returns {Object} 可能被替换 content 的结果
     */

    /**
     * 计算工具调用"进行中"状态的 UI 投影（供渲染层使用）。
     * @param {Object} tool - 工具定义
     * @param {Object} args - 调用参数
     * @returns {Object|undefined} 投影视图或 undefined（回退通用展示）
     */

    /**
     * 计算工具调用"完成"状态的 UI 投影（供渲染层使用）。
     * @param {Object} tool - 工具定义
     * @param {Object} args - 调用参数
     * @param {Object} result - 规范化结果
     * @returns {Object|undefined} 投影视图或 undefined（回退通用展示）
     */


    /**
     * 执行 Skill 的统一入口
     * @param {Object} skill - Skill 定义对象
     * @param {Object} args - LLM 传入的参数
     * @param {Object} ctx - 执行上下文
     * @returns {Promise<{ok: boolean, result: any, error?: string}>}
     */
    async function executeSkill(skill, args, ctx) {
        const skillName = skill.name || skill.Name;
        const processingMode = skill.processingMode || skill.ProcessingMode || 0;

        // 数字模式映射
        const modeMap = { 0: 'VBACode', 1: 'MultimodalLLM', 2: 'Hybrid', 3: 'LocalOCR', 4: 'PythonCode' };
        const mode = typeof processingMode === 'number' ? (modeMap[processingMode] || 'VBACode') : processingMode;

        ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} mode=${mode}`);

        // 【SkillSpec 空转治本】SkillSpec 模式的 Skill（SmartTableBeautifier / SmartTableFormatter 等）
        // 设计上由前端 skill-pipeline 的 Spec 引擎执行（LLM 规划 spec → 渲染 VBA）；agent-loop 把它
        // 当普通工具直调时，后端 lane（SkillExecutionRouter.ExecuteSkillSpecSkill）会返回
        // "由前端 Spec 引擎处理" 的假成功（Success=True，实测 ~18ms 返回，什么都没做）——
        // 美化/列宽全部未执行，随后覆盖度自检误报"excel_autofit 未执行"、用户看到"没有执行"告警。
        // 此处直接拦截返回引导性失败，让模型改用具体工具（excel_format_range / excel_autofit）真正完成操作。
        // （SkillSpec 在后端枚举中的数值序位为 5，兼容数字形态的 processingMode）
        if (String(mode) === 'SkillSpec' || processingMode === 5) {
            ChatLog.warn?.('TOOLS', `[executeSkill] ${skillName} 为 SkillSpec 模式，工具直调通道不执行（后端仅假成功透传）→ 返回引导性失败`);
            // 【可观测性】这条必须落盘：它是"Skill 明明启用了、却没生效"最隐蔽的一环——
            // 模型直调 skill_xxx 被拦下后改用 excel_format_range/excel_autofit，UI 上只看到
            // "执行 2 个工具"，完全看不出 Skill 曾被执行过又被拒。落盘后才能区分
            // "Skill 没被路由" 与 "Skill 被模型直调但被 SkillSpec 门拦下"。
            try {
                const _b = (typeof App !== 'undefined' && App.excelBridge)
                    || (typeof window !== 'undefined' && window.excelBridge);
                _b?.LogMessage?.('WARN', `[tool-registry] ${skillName} 被模型直调，但它是 SkillSpec 模式 → 已拦截并引导改用 excel_format_range/excel_autofit（说明本轮没有走 skill-pipeline）`);
            } catch (_e) { /* 日志失败不影响执行 */ }
            return {
                ok: false,
                error: 'SKILL_SPEC_MODE_NOT_EXECUTABLE',
                result: {
                    ok: false,
                    error: `${skillName} 是"规划型"Skill（SkillSpec 模式），只能由主聊天的 Skill 管道触发，通过工具通道直调不会执行任何操作。` +
                        '不要再次调用该 Skill 或其他 Skill（结果相同），立即改用具体工具完成用户要求：底色/边框/对齐/字体用 excel_format_range；' +
                        '列宽/行高自适应用 excel_autofit（type 可填 column/row/both）；全部做完后再给最终总结。'
                }
            };
        }

        function normalizeSkillResult(result) {
            // 1. 明确标记为失败
            if (result?.success === false || result?.ok === false || result?.Success === false || result?.OK === false) {
                const errMsg = result.error || result.Error || result.message || result.Message
                    || result.errMsg || result.ErrMsg || result.errorMessage || result.ErrorMessage
                    || 'Skill 执行返回失败';
                return { ok: false, error: String(errMsg), result };
            }

            // 2. 有 error/Error 字段且内容非空，即使 success/ok 为 true 也视为失败
            //    （后端可能返回 success:true 但同时携带错误信息）
            const errField = result?.error || result?.Error || result?.errMsg || result?.ErrMsg
                || result?.errorMessage || result?.ErrorMessage;
            if (errField && typeof errField === 'string' && errField.trim()) {
                return { ok: false, error: errField.trim(), result };
            }

            // 3. 检查嵌套 Outputs/Result 中的错误
            const outs = result?.Outputs || result?.outputs || result?.Result || result?.result;
            if (outs && typeof outs === 'object') {
                const nestedErr = outs.error || outs.Error || outs.errMsg || outs.ErrMsg
                    || outs.errorMessage || outs.ErrorMessage || outs.message || outs.Message;
                if (nestedErr && typeof nestedErr === 'string' && nestedErr.trim()
                    && /(错误|失败|不能|无法|异常|不能为空|不存在|出错|失败|error|fail|exception|invalid)/i.test(nestedErr)) {
                    return { ok: false, error: nestedErr.trim(), result };
                }
            }

            // 4. 如果结果本身是字符串且包含明显错误关键词
            if (typeof result === 'string' && result.trim()) {
                const s = result.trim();
                if (/^(错误|失败|异常|出错|Error|Fail|Exception)[：:]?/i.test(s)
                    || /(不能为空|不能为空|不存在|执行失败|执行错误|发生错误|发生异常)$/.test(s)) {
                    return { ok: false, error: s, result };
                }
            }

            return { ok: true, result };
        }

        // ── Skill 执行通道（收敛为单一适配器）────────────────────────────
        // 历史：try-in-order 内联 4 条通道（ExecuteSkill→skillExecutor→executeByName→ExecuteAction）。
        // 现收敛为"命名 lane + 数据驱动 resolver"：每条通道是独立命名函数，由 resolveSkillLanes
        // 按 skill.executor（可声明）排序；未声明时保持原降级链，行为零回退。
        // lane 返回 { executed, result?:{ok,error,result} }：executed=false=API 不可用→尝试下一条；
        // executed=true=已真正执行（成功或失败均立即返回，不回退，与原逻辑一致）。
        const laneSet = {
            backend: async () => {
                const bridge = window.chrome?.webview?.hostObjects?.excelBridge;
                if (!bridge || typeof bridge.ExecuteSkill !== 'function') return { executed: false };
                const inputs = buildSkillInputs(args, ctx);
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} via lane=backend (excelBridge.ExecuteSkill)`);
                const resultJson = await bridge.ExecuteSkill(skillName, JSON.stringify(inputs));
                const result = JSON.parse(resultJson);
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} result:`, result);
                return { executed: true, result: normalizeSkillResult(result) };
            },
            skillExecutor: async () => {
                if (!window.skillExecutor || typeof window.skillExecutor.executeAction !== 'function') return { executed: false };
                const action = buildSkillAction(skill, args, mode);
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} via lane=skillExecutor`);
                const result = await window.skillExecutor.executeAction(action);
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} result:`, result);
                return { executed: true, result: normalizeSkillResult(result) };
            },
            // 【GeneratePPTReport 专用 lane】后端将该 skill 当 MultimodalLLM 执行只返回 JSON 结构描述
            // 且 success:true，UI 误判成功 [object Object]，实际磁盘无文件。
            // 此处改为直接调 excelBridge.GeneratePPTFromData（真实 PowerPoint COM 生成）。
            pptDirect: async () => {
                const isPpt = skillName === 'GeneratePPTReport' || skillName === 'skill.GeneratePPTReport';
                if (!isPpt) return { executed: false };
                const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
                if (!bridge || typeof bridge.GeneratePPTFromData !== 'function') return { executed: false };
                // 数据源：优先 LLM 调用参数里显式的 rangeAddress/headers/rows，其次 ctx 选区/参考区
                const ctxSel = ctx?.selection;
                const ctxRef = Array.isArray(ctx?.referenceRegions) && ctx.referenceRegions.length > 0 ? ctx.referenceRegions[0] : null;
                const src = (ctxSel && Array.isArray(ctxSel.headers) && ctxSel.headers.length) ? ctxSel
                    : (ctxRef && Array.isArray(ctxRef.headers) && ctxRef.headers.length) ? ctxRef : null;
                const headers = Array.isArray(args?.headers) ? args.headers
                    : (Array.isArray(args?.columns) ? args.columns
                    : (src && Array.isArray(src.headers) ? src.headers : []));
                const rows = Array.isArray(args?.rows) ? args.rows
                    : (src && Array.isArray(src.rows) ? src.rows : []);
                const rangeAddress = String(args?.rangeAddress || args?.range || args?.address || (src?.address) || '');
                const payload = {
                    reportTitle: String(args?.reportTitle || args?.title || '数据分析报告'),
                    rangeAddress,
                    headers,
                    rows,
                    theme: String(args?.theme || ''),
                    chartMode: String(args?.chartMode || 'auto') || 'auto',
                    ratio: String(args?.ratio || '16x9') || '16x9',
                    orientation: String(args?.orientation || '')
                };
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} via lane=pptDirect (GeneratePPTFromData, range=${rangeAddress || '无地址'} headers=${headers.length} rows=${rows.length})`);
                const genJson = await bridge.GeneratePPTFromData(JSON.stringify(payload));
                let gen = {};
                try { gen = JSON.parse(genJson || '{}') || {}; } catch { gen = { raw: genJson }; }
                const ok = gen?.success === true || gen?.Success === true;
                const err = gen?.error || gen?.Error || '';
                if (!ok) {
                    return { executed: true, result: { ok: false, error: String(err || 'GeneratePPTFromData 返回失败'), result: gen } };
                }
                return { executed: true, result: { ok: true, ...gen } };
            },
            executeByName: async () => {
                if (!App.executeSkillByName || typeof App.executeSkillByName !== 'function') return { executed: false };
                const base64Data = args?.imageBase64 || args?.base64 || args?.image || '';
                const fileName = args?.fileName || args?.filename || `${skillName}_output`;
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} via lane=executeByName`);
                const result = await App.executeSkillByName(skillName, base64Data, fileName);
                return { executed: true, result: normalizeSkillResult(result) };
            },
            executeAction: async () => {
                const bridge = window.chrome?.webview?.hostObjects?.excelBridge;
                if (!bridge || typeof bridge.ExecuteAction !== 'function') return { executed: false };
                const action = buildSkillAction(skill, args, mode);
                ChatLog.info?.('TOOLS', `[executeSkill] ${skillName} via lane=executeAction (excelBridge.ExecuteAction)`);
                const resultJson = await bridge.ExecuteAction(JSON.stringify(action));
                const result = JSON.parse(resultJson);
                return { executed: true, result: normalizeSkillResult(result) };
            }
        };

        // 默认降级链（保持原 try-in-order 顺序，行为零回退）
        const DEFAULT_LANE_ORDER = ['backend', 'skillExecutor', 'executeByName', 'executeAction'];

        // 数据驱动 lane 选择：优先 skill.executor.via（单值或有序数组），否则默认降级链。
        function resolveSkillLanes() {
            // 【GeneratePPTReport 强优先】agent-loop 调用该 skill 时，必须走真实生成通道（pptDirect），
            // 否则后端 MultimodalLLM 只返回 JSON 描述且 success:true，被误判成功却无文件。仅它独享。
            if (skillName === 'GeneratePPTReport' || skillName === 'skill.GeneratePPTReport') {
                return ['pptDirect', ...DEFAULT_LANE_ORDER.filter(k => k !== 'pptDirect')];
            }
            const declared = skill?.executor?.via;
            if (typeof declared === 'string' && Object.prototype.hasOwnProperty.call(laneSet, declared)) {
                return [declared];
            }
            if (Array.isArray(declared)) {
                const keys = declared.filter(k => Object.prototype.hasOwnProperty.call(laneSet, k));
                if (keys.length) return keys;
            }
            return DEFAULT_LANE_ORDER.slice();
        }

        try {
            const laneOrder = resolveSkillLanes();
            for (const key of laneOrder) {
                const lane = laneSet[key];
                try {
                    const out = await lane();
                    if (out.executed) return out.result;
                } catch (err) {
                    // lane 已真正执行但抛异常 → 视为执行失败，立即返回（不回退，与原行为一致）
                    ChatLog.error?.('TOOLS', `[executeSkill] lane=${key} ${skillName} failed:`, err);
                    return { ok: false, error: err?.message || String(err) };
                }
            }
            return { ok: false, error: 'Skill executor not available (ExecuteSkill / skillExecutor / executeSkillByName / excelBridge 均不可用)' };
        } catch (err) {
            ChatLog.error?.('TOOLS', `[executeSkill] ${skillName} failed:`, err);
            return { ok: false, error: err?.message || String(err) };
        }
    }

    /**
     * 构建 Skill Action 对象（供 SkillExecutor / excelBridge.ExecuteAction 使用）
     */

    /**
     * 在工具出口附上 Value / Context / Error 信封。模块未加载时原样返回。
     */
    function stampEnvelope(name, payload) {
        if (!payload || typeof payload !== 'object' || payload.envelope) return payload;
        try {
            const api = App.ChatV2 && App.ChatV2.toolEnvelope;
            if (api && typeof api.fromInvoke === 'function') payload.envelope = api.fromInvoke(name, payload);
        } catch (_) { /* 信封失败不影响工具结果 */ }
        return payload;
    }

    /**
     * 构建后端 SkillExecutionRouter 所需的 inputs 字典
     */

    const registry = {
        /**
         * 注册工具
         * @param {Object} tool - 工具定义，name 字段会被规范化
         */
        register(tool) {
            if (!tool?.name) {
                ChatLog.warn?.('TOOLS', '[tool-registry] register 失败：缺少 name');
                return false;
            }
            // 保留原始名用于内部调用
            const originalName = tool.originalName || tool.name;
            // 规范化工具名（符合 LLM function calling 命名规则）
            const normalizedName = normalizeName(tool.name);
            tool.originalName = originalName;
            tool.name = normalizedName;

            if (_tools.has(normalizedName)) {
                ChatLog.debug?.('TOOLS', `[tool-registry] 覆盖已存在工具: ${normalizedName} (原始: ${originalName})`);
            }
            _tools.set(normalizedName, tool);
            _nameMap.set(normalizedName, originalName);
            return true;
        },

        /**
         * 批量注册
         */
        registerAll(tools) {
            if (!Array.isArray(tools)) return 0;
            let count = 0;
            for (const t of tools) {
                if (this.register(t)) count++;
            }
            return count;
        },

        /**
         * 取消注册
         */
        unregister(name) {
            const normalizedName = normalizeName(name);
            _nameMap.delete(normalizedName);
            return _tools.delete(normalizedName);
        },

        /**
         * 列出所有工具（给 LLM function calling 用）
         */
        list() {
            return Array.from(_tools.values()).map(t => ({
                name: t.name,
                description: t.description,
                inputSchema: t.inputSchema || { type: 'object', properties: {} }
            }));
        },

        /**
         * 列出工具名
         */
        names() {
            return Array.from(_tools.keys());
        },

        /**
         * 按来源筛选
         */
        listBySource(source) {
            return Array.from(_tools.values()).filter(t => t.source === source);
        },

        /**
         * 获取单个工具定义
         * @param {string} name - 规范化名或原始名均可
         */
        get(name) {
            // 先尝试直接查找（规范化名）
            let tool = _tools.get(name);
            if (tool) return tool;
            // 再尝试规范化后查找
            const normalizedName = normalizeName(name);
            return _tools.get(normalizedName);
        },

        /**
         * 该工具的某参数是否有安全默认生成器（SAFE_DEFAULT_GENERATORS）。
         * 供 agent-loop 强制兜底的"缺必填参数"检查使用：有生成器的参数缺了也能被
         * invokeTracked 自动补全（如 excel_export_csv.filePath 自动落桌面），
         * 不应因缺参而放弃兜底执行。
         * @param {string} toolName - 工具名（裸名，如 excel_export_csv）
         * @param {string} paramKey - 参数名（如 filePath）
         * @returns {boolean}
         */
        hasSafeDefault(toolName, paramKey) {
            if (!toolName || !paramKey) return false;
            const tool = _tools.get(toolName) || _tools.get(normalizeName(toolName));
            const bare = tool ? String(tool.name || tool.originalName || '').toLowerCase() : String(toolName).toLowerCase();
            return typeof SAFE_DEFAULT_GENERATORS[bare + '.' + paramKey] === 'function'
                || typeof SAFE_DEFAULT_GENERATORS[String(toolName).toLowerCase() + '.' + paramKey] === 'function';
        },

        /**
         * LLM 编造工具名 → 真实工具名翻译兜底
         * 链式多步时 LLM 经常输出 excel_paste / excel_get_range_value 这种看起来合理但实际不存在的名字。
         * 用 ngram+keywords 匹配找最接近的真实工具。
         * @param {string} fakeName LLM 编造的工具名
         * @returns {string|null} 最接近的真实工具名（score>=15），或 null
         */
        _resolveToolAlias(fakeName) {
            if (!fakeName || typeof fakeName !== 'string') return null;
            const msg = fakeName.toLowerCase();
            const all = Array.from(_tools.values());
            // 🔥 家族守卫（与后端 ActionEngine 同族翻译策略保持一致，杜绝跨族误译）：
            //   - 原名带家族前缀（excel_/ppt_/computer_/web_/time_）→ 只允许译到同族；
            //   - 裸名 → 本宿主为 Excel，默认排除 ppt_ 族；仅当原名含 ppt/幻灯片/演示 提示才允许 ppt_。
            // 实测：裸名 "chart" / "skill.chart" 曾跨族译成 ppt_add_chart_slide（后端报 "PPT 项目不存在"）。
            const _familyOf = (n) => (String(n || '').toLowerCase().match(/^([a-z]+)_/) || [])[1] || '';
            const _fakePrefix = _familyOf(fakeName);
            const _pptHint = /ppt|slide|幻灯片|演示|汇报稿/.test(msg);
            const _desiredFamily = _fakePrefix || (_pptHint ? 'ppt' : '');
            const _allowFamily = (tn) => {
                const fam = _familyOf(tn);
                if (_desiredFamily) return fam === _desiredFamily;
                return fam !== 'ppt';
            };
            let best = null, bestScore = 0;
            for (const t of all) {
                let s = 0;
                const tn = (t.name || t.originalName || '').toLowerCase();
                if (!tn) continue;
                if (!_allowFamily(tn)) continue;   // 家族守卫：禁止跨族翻译
                const fakePrefix = _fakePrefix;
                const realPrefix = _familyOf(tn);
                if (fakePrefix && fakePrefix === realPrefix) s += 30;
                const dn = (t.displayName || '').toLowerCase();
                if (dn) { for (let n = 2; n <= 4; n++) { for (let i = 0; i <= fakeName.length - n; i++) { const ng = fakeName.slice(i, i + n); if (dn.includes(ng)) s += n * 5; } } }
                const desc = (t.description || '').toLowerCase();
                if (desc) { for (let n = 2; n <= 4; n++) { for (let i = 0; i <= fakeName.length - n; i++) { const ng = fakeName.slice(i, i + n); if (desc.includes(ng)) s += n * 3; } } }
                const kwList = [...(Array.isArray(t.keywords) ? t.keywords : [])].filter(k => typeof k === 'string').map(k => k.toLowerCase());
                for (const kw of kwList) { if (kw.length >= 2 && msg.includes(kw)) s += kw.length * 3; }
                if (s > bestScore) { bestScore = s; best = t.name || t.originalName; }
            }
            if (bestScore >= 15) return best;
            return null;
        },

        /**
         * 判断是否为「图片类」工具（仅适用于图片输入）。
         * text-pipeline 只处理纯文本，不应向 LLM 暴露图片类工具（图片走 image-pipeline，
         * kind='image'）。规范名由 IMAGE_ONLY_TOOLS 单一维护，此处兼容规范化名与
         * excel-local__ 前缀原始名。
         * @param {string|Object} nameOrTool - 工具名，或含 name/originalName 的工具对象
         * @returns {boolean}
         */
        isImageOnlyTool(nameOrTool) {
            const candidates = [];
            if (typeof nameOrTool === 'string') {
                candidates.push(nameOrTool);
            } else if (nameOrTool) {
                if (nameOrTool.name) candidates.push(nameOrTool.name);
                if (nameOrTool.originalName) candidates.push(nameOrTool.originalName);
            }
            for (const raw of candidates) {
                if (!raw) continue;
                const stripped = raw.replace(/^[a-z]+-local__/, '');
                if (IMAGE_ONLY_TOOLS.includes(raw) || IMAGE_ONLY_TOOLS.includes(stripped)) return true;
            }
            return false;
        },

        /**
         * 调用工具（带 UI 追踪）—— 通用方法，任何 pipeline / skill / 代码路径都可以使用。
         * 自动完成：addToolCall（创建UI卡片）→ invoke（执行工具）→ updateToolCall（更新结果）。
         * UI 渲染完全由 sessionStore 通知驱动，调用方无需关心渲染细节。
         *
         * @param {Object} turn - 当前 turn 对象（必须有 id）
         * @param {string} name - 工具名
         * @param {Object} args - 调用参数（实际传给工具的参数）
         * @param {Object} [options] - 可选参数
         * @param {number} [options.round=1] - 步骤/轮次编号，同一 turn 内递增以分组显示
         * @param {string} [options.displayName] - 覆盖显示名（默认取工具注册表的 displayName）
         * @param {Object} [options.displayArgs] - 覆盖 UI 展示的参数（默认使用 args；用于隐藏大段内容如脚本源码）
         * @param {Object} [options.ctx] - 额外执行上下文（signal 等）
         * @returns {Promise<{ok, result, error, durationMs, toolCallId}>}
         */
        async invokeTracked(turn, name, args, options = {}) {
            const sessionStore = App.ChatV2?.sessionStore;
            if (!sessionStore) {
                ChatLog.warn?.('TOOLS', '[invokeTracked] sessionStore 不可用，降级为普通 invoke');
                const res = await this.invoke(name, args, options.ctx || {});
                return { ...res, toolCallId: null };
            }

            // round 未显式指定时自动递增（同一 turn 内按调用顺序分配），
            // 各管线无需再手动维护 nextRound 计数器，步骤编号由统一 invokeTracked 层负责。
            const round = options.round !== undefined ? options.round : computeNextRound(turn);
            let tool = this.get(name);
            // 【域空壳治本 · 统一执行层】invokeTracked 是所有管线（skill/text/agent/dashboard）的唯一执行咽喉。
            // 'skill.<域id>'（如 skill.chart）或裸域 id 并非真实工具——域只有 name+keywords，
            // 其真实语义是该域的确定性工具链（DESCRIPTORS.guardedTools）。
            // 若放任其进入下方 ngram 兜底翻译，极易跨族把 skill.chart 译成 ppt_add_chart_slide
            // （后端报 "PPT 项目不存在"，实测）。此处按域的 guardedTools 首领归一到真实工具，
            // 并把 skill 形态参数（prompt/text/message/excelData + selection 对象）清洗成目标工具的合法参数
            // （selection 对象 → dataRange 选区地址）。
            if (!tool) {
                const _domainId = String(name || '').replace(/^skill\./i, '').toLowerCase();
                const _dd = App.ChatV2?.skillDomain?.getById?.(_domainId);
                const _guards = (_dd && Array.isArray(_dd.guardedTools)) ? _dd.guardedTools : [];
                const _real = _guards.map(n => this.get(n)).find(Boolean);
                if (_real) {
                    const _sel = args && args.selection;
                    const _addr = _sel && (typeof _sel === 'object' ? (_sel.address || _sel.range) : String(_sel));
                    const _junk = ['prompt', 'text', 'message', 'excelData', 'selection', 'referenceRegions', 'attachments'];
                    const _clean = {};
                    for (const k of Object.keys(args || {})) {
                        if (_junk.includes(k)) continue;
                        _clean[k] = args[k];
                    }
                    if (_addr && !_clean.dataRange) _clean.dataRange = String(_addr);
                    ChatLog.warn?.('TOOLS', `[invokeTracked] 域空壳 "${name}" → 归一为真实工具 "${_real.name}"（域 ${_domainId} 首个受护工具；selection→dataRange=${_addr || '∅'}）`);
                    tool = _real;
                    name = _real.name;
                    args = _clean;
                    options = { ...options, displayName: _real.displayName || _real.name, displayArgs: undefined };
                }
            }
            // 🔥 工具名翻译兜底：LLM 链式调用经常编不存在的工具名（如 excel_paste），
            // 用前端 ngram+keywords 匹配找最接近的真实工具（跟 McpDispatcher 同源）
            if (!tool) {
                const resolved = this._resolveToolAlias(name);
                if (resolved && resolved !== name) {
                    ChatLog.warn?.('TOOLS', `[invokeTracked] LLM 编造工具名 "${name}" → 翻译为 "${resolved}"`);
                    tool = this.get(resolved);
                    name = resolved;
                }
            }
            const displayName = options.displayName || tool?.displayName || name;
            const displayArgs = options.displayArgs !== undefined ? options.displayArgs : args;
            const isHidden = !!options.hidden;

            // LLM 输出通用/蛇形参数 → 后端期望的驼峰参数，
            // 归一化后同时用于 UI 投影与执行，保证"规划一致、执行一致"。
            const adapter = TOOL_ARG_ADAPTERS[name] || TOOL_ARG_ADAPTERS[normalizeName(name)];
            const effectiveArgs = adapter ? adapter(args || {}) : { ...(args || {}) };
            // 透明化补全：声明了 range/sheet 的选区类工具（如 excel_clear_range），
            // 用当前选区 / 活动工作表补全缺失参数，使工具卡片与实际执行都明确显示具体工作表与区域。
            const ctxTurn = (options.ctx && options.ctx.turn) || turn;
            const fillSel = fillSelectionArgs(name, effectiveArgs, ctxTurn);
            Object.assign(effectiveArgs, fillSel || {});
            // 图表工具智能默认值：未显式提供时补合理默认（标题/图例/轴标题/数据标签/样式），
            // 避免图表空白或被参数卡死；用户/规划者显式传入的一律尊重、不覆盖。
            if (normalizeName(name) === 'excel_create_chart') {
                const chartDefaults = inferChartDefaults(ctxTurn, effectiveArgs);
                if (chartDefaults) {
                    Object.assign(effectiveArgs, chartDefaults);
                }
            }
            // effectiveDisplayArgs 用于卡片渲染，必须先于下方对其的改写声明
            //（与 effectiveArgs 一致地合并 adapter 归一化、透明补全的 range/sheet 等关键字段）
            const effectiveDisplayArgs = Object.assign(
                { ...(adapter ? adapter(displayArgs || {}) : (displayArgs || {})), ...(fillSel || {}) }
            );
            // sheet 参数可能是对象形态，归一化为字符串，避免展示/后端取值出错
            normalizeSheetArg(effectiveArgs);
            normalizeSheetArg(effectiveDisplayArgs);
            // 【诊断】excel_create_chart 定位参数侦探：捕捉 position/anchorRange 是否被错误注入为 HTML/非 JSON。
            // 现象："position JSON 错误: Unexpected character ... value: <"，需确认谁把 position 填成 HTML。
            if (normalizeName(name) === 'excel_create_chart' && effectiveArgs.position !== undefined) {
                ChatLog.info?.('TOOLREG', `[position侦探] excel_create_chart.position 类型=${typeof effectiveArgs.position}; 预览="${String(effectiveArgs.position).slice(0, 160)}"`);
            }
            // 多项操作工具：每个子操作是一个独立"零件"，由规划者按需装配。
            // 仅当规划者（模型）漏传了该工具的意图归属参数时，才按当前用户消息补齐所需粒度，
            // 避免多余零件被一并装配；规划者已显式指定的参数一律尊重、不覆盖。
            const INTENT_PARAM_KEYS = App.ChatV2.toolRegistryUtil?.INTENT_PARAM_KEYS || {};
            const intentKeys = INTENT_PARAM_KEYS[name] || INTENT_PARAM_KEYS[normalizeName(name)];
            if (intentKeys && intentKeys.some((k) => effectiveArgs[k] === undefined)) {
                const intent = resolveIntentFlagsFromMessage(ctxTurn, name);
                if (intent) {
                    for (const k of Object.keys(intent)) {
                        if (effectiveArgs[k] === undefined) {
                            effectiveArgs[k] = intent[k];
                            effectiveDisplayArgs[k] = intent[k];
                        }
                    }
                }
            }
            // 「按销售额从高到低」：模型常照抄 schema 示例把 sortBy.column 写成 A（次序倒是对的）。
            // 漏传 sortBy 时上面的意图补全也会落到默认 A。这里用用户点名的表头列覆盖，
            // 后端按表头解析成真实列，避免整表被按第一列重排。
            if (normalizeName(name) === 'excel_sort_range' && applyExplicitSortColumn(effectiveArgs, ctxTurn)) {
                effectiveDisplayArgs.sortBy = effectiveArgs.sortBy;
            }
            // 以下合并 effectiveArgs 中的关键展示字段（如 chartType），防止卡片类型标签回退到默认值。
            if (effectiveArgs.chartType && !effectiveDisplayArgs.chartType) {
                effectiveDisplayArgs.chartType = effectiveArgs.chartType;
            }
            if (effectiveArgs.dataRange && !effectiveDisplayArgs.dataRange) {
                effectiveDisplayArgs.dataRange = effectiveArgs.dataRange;
            }
            if (effectiveArgs.title && !effectiveDisplayArgs.title) {
                effectiveDisplayArgs.title = effectiveArgs.title;
            }
            if (effectiveArgs.xAxisTitle && !effectiveDisplayArgs.xAxisTitle) {
                effectiveDisplayArgs.xAxisTitle = effectiveArgs.xAxisTitle;
            }
            if (effectiveArgs.yAxisTitle && !effectiveDisplayArgs.yAxisTitle) {
                effectiveDisplayArgs.yAxisTitle = effectiveArgs.yAxisTitle;
            }
            if (effectiveArgs.hasLegend !== undefined && effectiveDisplayArgs.hasLegend === undefined) {
                effectiveDisplayArgs.hasLegend = effectiveArgs.hasLegend;
            }
            if (effectiveArgs.showDataLabels !== undefined && effectiveDisplayArgs.showDataLabels === undefined) {
                effectiveDisplayArgs.showDataLabels = effectiveArgs.showDataLabels;
            }
            if (effectiveArgs.chartStyle !== undefined && effectiveDisplayArgs.chartStyle === undefined) {
                effectiveDisplayArgs.chartStyle = effectiveArgs.chartStyle;
            }
            if (effectiveArgs.dimColumn && !effectiveDisplayArgs.dimColumn) {
                effectiveDisplayArgs.dimColumn = effectiveArgs.dimColumn;
            }
            if (effectiveArgs.valueColumn && !effectiveDisplayArgs.valueColumn) {
                effectiveDisplayArgs.valueColumn = effectiveArgs.valueColumn;
            }
            if (effectiveArgs.aggFunc && !effectiveDisplayArgs.aggFunc) {
                effectiveDisplayArgs.aggFunc = effectiveArgs.aggFunc;
            }
            // excel_date_group 展示字段同步
            if (effectiveArgs.dateColumn && !effectiveDisplayArgs.dateColumn) {
                effectiveDisplayArgs.dateColumn = effectiveArgs.dateColumn;
            }
            if (effectiveArgs.groupBy && !effectiveDisplayArgs.groupBy) {
                effectiveDisplayArgs.groupBy = effectiveArgs.groupBy;
            }
            if (effectiveArgs.weekStart && !effectiveDisplayArgs.weekStart) {
                effectiveDisplayArgs.weekStart = effectiveArgs.weekStart;
            }
            if (effectiveArgs.outputColumn && !effectiveDisplayArgs.outputColumn) {
                effectiveDisplayArgs.outputColumn = effectiveArgs.outputColumn;
            }
            if (effectiveArgs.fiscalYearStart && !effectiveDisplayArgs.fiscalYearStart) {
                effectiveDisplayArgs.fiscalYearStart = effectiveArgs.fiscalYearStart;
            }
            if (effectiveArgs.fiscalQuarter !== undefined && effectiveDisplayArgs.fiscalQuarter === undefined) {
                effectiveDisplayArgs.fiscalQuarter = effectiveArgs.fiscalQuarter;
            }
            if (effectiveArgs.aggColumn && !effectiveDisplayArgs.aggColumn) {
                effectiveDisplayArgs.aggColumn = effectiveArgs.aggColumn;
            }
            if (effectiveArgs.outputSheet && !effectiveDisplayArgs.outputSheet) {
                effectiveDisplayArgs.outputSheet = effectiveArgs.outputSheet;
            }
            if (normalizeName(name) === 'excel_format_range') {
                const confine = App.ChatV2.toolRegistryUtil && App.ChatV2.toolRegistryUtil.confineFormatFillToHeader;
                if (typeof confine === 'function') {
                    const headerFill = confine(effectiveArgs, ctxTurn);
                    if (headerFill) options = { ...options, _headerFill: headerFill };
                    const executedStillFills = !!(effectiveArgs.format && typeof effectiveArgs.format === 'object' && effectiveArgs.format.bgColor);
                    if (!executedStillFills && effectiveDisplayArgs.format) {
                        if (typeof effectiveDisplayArgs.format === 'object') {
                            delete effectiveDisplayArgs.format.bgColor;
                            delete effectiveDisplayArgs.format.bgScope;
                        } else if (typeof effectiveDisplayArgs.format === 'string') {
                            try {
                                const shown = JSON.parse(effectiveDisplayArgs.format);
                                if (shown && typeof shown === 'object') {
                                    delete shown.bgColor;
                                    delete shown.bgScope;
                                    effectiveDisplayArgs.format = shown;
                                }
                            } catch (_) { /* 展示串保持原样 */ }
                        }
                    }
                }
            }

            // 必填参数兜底（"推断+反问结合"恢复策略）——仅通用对话/文本路径启用。
            // 能安全推断/默认的参数自动补全（如 excel_create_sheet.name → "新建工作表N"），
            // 无法兜底的收集为 needsInput，交由 agent-loop 向用户反问补齐，而不是硬失败。
            let needsInput = null;
            if (ctxTurn && ctxTurn.kind === 'text') {
                const filled = await completeRequiredArgs(name, effectiveArgs, ctxTurn);
                if (filled && filled.added && Object.keys(filled.added).length) {
                    Object.assign(effectiveArgs, filled.added);
                    Object.assign(effectiveDisplayArgs, filled.added);
                }
                if (filled && filled.missing && filled.missing.length) {
                    needsInput = filled.missing;
                }
            }

            // 1a. 拦截：必填参数无法兜底 → 不执行工具，创建 needs-input 卡片并反问用户
            if (needsInput && needsInput.length) {
                const missingTc = sessionStore.addToolCall(turn.id, {
                    name,
                    displayName,
                    args: effectiveDisplayArgs || {},
                    round,
                    hidden: isHidden,
                    status: 'needs-input',
                    presentCall: this.renderCardIntent(name, 'call', effectiveArgs) || undefined
                }, { silent: isHidden });
                const question = needsInput.map(m => m.question).join(' ');
                sessionStore.updateToolCall(turn.id, missingTc.id, {
                    status: 'needs-input',
                    result: null,
                    error: '缺少必要参数',
                    errorSuggestion: question,
                    needsInput,
                    durationMs: 0,
                    completedAt: Date.now()
                }, { silent: isHidden });
                this.needsInputPending = { turnId: turn.id, name, missing: needsInput, question };
                return stampEnvelope(name, { ok: false, needsInput, error: '缺少必要参数：' + question, durationMs: 0, toolCallId: missingTc.id });
            }

            // 0. 能力去重：本轮已执行过同组且目标重叠的操作 → 跳过（不执行、不再弹危险确认），
            //    仍生成一张卡片显式说明"已由上一步完成，已跳过"，让用户看得见省掉了哪一步
            const capGroup = capabilityGroupOf(name, effectiveArgs);
            if (capGroup && AUTO_SKIP_GROUPS[capGroup]) {
                const dup = findDuplicateCapability(turn, capGroup, (effectiveArgs && effectiveArgs.sheet) || '', targetRectOf(name, effectiveArgs));
                if (dup) {
                    const skipMsg = `已跳过：本轮已执行过同类操作（${dup.toolName}），目标区域重叠，无需重复执行。`;
                    ChatLog.info?.('TOOLS', `[capability-dedup] ${name} 命中重复（同组=${capGroup}，前序=${dup.toolName}），已跳过`);
                    const skipTc = sessionStore.addToolCall(turn.id, {
                        name,
                        displayName,
                        args: effectiveDisplayArgs || {},
                        round,
                        hidden: isHidden,
                        presentCall: this.renderCardIntent(name, 'call', effectiveArgs) || undefined
                    }, { silent: isHidden });
                    sessionStore.updateToolCall(turn.id, skipTc.id, {
                        status: 'completed',
                        result: skipMsg,
                        error: null,
                        durationMs: 0,
                        completedAt: Date.now(),
                        presentResult: { card: 'generic', kind: 'info', content: skipMsg }
                    }, { silent: isHidden });
                    return stampEnvelope(name, { ok: true, result: skipMsg, durationMs: 0, skipped: true, toolCallId: skipTc.id });
                }
            }

            // 1. 创建 toolCall 记录（hidden=true 时静默创建，不触发 UI 渲染）
            const tc = sessionStore.addToolCall(turn.id, {
                name,
                displayName,
                args: effectiveDisplayArgs || {},
                round,
                hidden: isHidden,
                // 调用中状态投影（供渲染层使用，可选）
                presentCall: this.renderCardIntent(name, 'call', effectiveArgs) || undefined
            }, { silent: isHidden });

            // 2. 执行工具
            const result = await this.invoke(name, effectiveArgs || {}, {
                turn,
                sessionStore,
                ...(options.ctx || {})
            });
            if (result && result.ok && options._headerFill) {
                try {
                    await this.invoke(name, options._headerFill, { turn, sessionStore, ...(options.ctx || {}) });
                } catch (fillErr) {
                    ChatLog.warn?.('TOOLS', '[invokeTracked] 表头底色补涂失败: ' + (fillErr && fillErr.message));
                }
            }

            // 3. 工具失败时，附加错误分类信息（类型 + 原因 + 建议）
            let classifiedError = null;
            if (!result.ok && result.error) {
                classifiedError = classifyError(result.error, name);
            }

            // 4. 更新 toolCall 结果（hidden 时静默更新）
            const patch = {
                status: result.ok ? 'completed' : 'failed',
                cancelled: result.cancelled === true || undefined,
                result: result.result,
                error: result.error,
                durationMs: result.durationMs,
                completedAt: Date.now(),
                // 完成状态投影（供渲染层使用，可选）
                presentResult: this.renderCardIntent(name, 'result', effectiveArgs, result) || undefined,
                // output schema 校验结果（可选）
                contractErrors: result._contractErrors || null,
                contractValid: result._contractValid === true ? true : null,
                envelope: result.envelope || null
            };
            if (classifiedError) {
                patch.errorType = classifiedError.type;
                patch.errorReason = classifiedError.reason;
                patch.errorSuggestion = classifiedError.suggestion;
            }
            sessionStore.updateToolCall(turn.id, tc.id, patch, { silent: isHidden });

            return { ...result, toolCallId: tc.id, classifiedError };
        },

        /**
         * 调用工具
         * @param {string} name - 工具名（原始名或规范化名均可）
         * @param {Object} args - 调用参数
         * @param {Object} ctx - 执行上下文（turn, sessionStore, signal 等）
         * @returns {Promise<{ok, result, error, durationMs, envelope}>}
         */
        async invoke(name, args, ctx = {}) {
            // 先直接查找，再规范化后查找
            let tool = _tools.get(name);
            if (!tool) {
                const normalizedName = normalizeName(name);
                tool = _tools.get(normalizedName);
            }
            if (!tool) {
                return stampEnvelope(name, {
                    ok: false,
                    error: `Tool not found: ${name}`,
                    durationMs: 0
                });
            }

            // 运行执行前守卫，拒绝则处理：
            //   ——【安全守卫】危险操作：弹确认框，用户确认后放行（任何模式都可执行）；
            //   ——其他硬拒绝：直接返回（不执行）。
            const denyReason = await runGuards(name, args, ctx);
            if (denyReason) {
                const isGuardDeny = typeof denyReason === 'string' && denyReason.indexOf('【安全守卫】') === 0;
                if (isGuardDeny) {
                    let confirmed = false;
                    try {
                        confirmed = await confirmDangerousAction(name, denyReason, { args: args || {} });
                    } catch (_e) { confirmed = false; }
                    if (confirmed) {
                        ChatLog.warn?.('GUARD', `[tool-registry] ${name} 危险操作已由用户确认，继续执行`);
                        // 「始终询问」开启时不记忆 allow，下次同类操作仍弹窗
                        if (!isDangerAlwaysAsk()) {
                            setTurnDangerousDecision(ctx && ctx.turn, 'allow');
                        }
                    } else {
                        // 取消也记住：本轮所有危险操作自动拒绝（防止技能换工具绕过用户决定）
                        setTurnDangerousDecision(ctx && ctx.turn, 'deny');
                        return stampEnvelope(name, {
                            ok: false,
                            error: `已取消，未执行任何操作。原因：${denyReason}\n（若相关删除/清空已由前面的步骤完成，此取消不影响已有结果；请勿再重复发起同类调用。）`,
                            durationMs: 0,
                            guarded: true,
                            cancelled: true
                        });
                    }
                } else {
                    ChatLog.warn?.('TOOLS', `[tool-registry] ${name} 被守卫拒绝: ${denyReason}`);
                    return stampEnvelope(name, {
                        ok: false,
                        error: denyReason,
                        durationMs: 0,
                        guarded: true
                    });
                }
            }

            const start = Date.now();
            try {
                ChatLog.info?.('TOOLS', `[tool-registry] invoke ${name}`, args, { source: tool?.source, orig: tool?.originalName, skill: tool?.skillInfo?.name, nkeys: _tools.size });
                // _view 是纯前端渲染提示（见 llm-client.withViewHint）：只用于结果展示，绝不能下发给后端。
                // 在副本上剥离，保证后端收不到未知字段，同时调用方持有的原始 args（含 _view）不受影响，
                // 渲染层 updateToolCard 仍能从原 args 读到 _view。
                let _execArgs = args || {};
                if (_execArgs && typeof _execArgs === 'object' && Object.prototype.hasOwnProperty.call(_execArgs, '_view')) {
                    _execArgs = Object.assign({}, _execArgs);
                    delete _execArgs._view;
                }
                const raw = await tool.execute(_execArgs, ctx);
                const durationMs = Date.now() - start;

                // 标准化结果
                // 统一判定逻辑：
                // 1) 后端明确报告的 success/isError 字段优先（可信状态），不再被 warnings 中的诊断文本覆盖；
                //    —— 切片器等工具走多层回退，warnings 里会带"FALLBACK"及异常消息，但只要 success=true 即视为成功。
                // 2) 无显式状态字段时，用 hasErrorField 启发式检测（兼容旧工具/裸数据返回）。
                const explicitSuccess = raw?.success === true || raw?.Success === true
                    || raw?.isError === false || raw?.IsError === false
                    || raw?.ok === true || raw?.OK === true;
                const explicitFailure = raw?.ok === false || raw?.OK === false
                    || raw?.success === false || raw?.Success === false
                    || raw?.isError === true || raw?.IsError === true;
                let ok;
                if (explicitSuccess) {
                    ok = true;
                } else if (explicitFailure) {
                    ok = false;
                } else {
                    ok = !hasErrorField(raw);
                }
                let result = {
                    ok,
                    result: raw?.result !== undefined ? raw.result : raw,
                    error: extractErrorField(raw),
                    durationMs
                };

                // output schema 校验（声明了才校验，不影响旧工具）
                if (tool.outputSchema) {
                    const outVal = result.ok ? result.result : undefined;
                    if (result.ok && outVal !== undefined) {
                        const check = validateOutput(tool.outputSchema, outVal);
                        if (!check.valid) {
                            ChatLog.warn?.('TOOLS', `[tool-registry] ${name} output 契约校验失败: ${check.errors.join('; ')}`);
                            result._contractErrors = check.errors;
                        } else {
                            result._contractValid = true;
                        }
                    }
                }

                // 最后一英里内容归一化（finalizeContent）
                result = finalizeToolContent(tool, args, result);

                // 运行 post-execute 钩子，允许替换/增强结果
                result = await runPostExecuteHooks(name, args, result, ctx);

                return stampEnvelope(name, result);
            } catch (err) {
                const durationMs = Date.now() - start;
                ChatLog.error?.('TOOLS', `[tool-registry] invoke ${name} failed`, err);
                return stampEnvelope(name, {
                    ok: false,
                    error: err?.message || String(err),
                    durationMs
                });
            }
        },

        /**
         * 读取 MCP 工具源（按优先级：App.mcpTools → App.mcp.listTools() → window.mcpTools → window._mcpTools）
         * 唯一真实来源是 window._mcpTools（由 app-mcp.js 的 loadMcpToolsForAI 异步填充）。
         */
        _readMcpToolSource() {
            return (
                App.mcpTools ||
                App.mcp?.listTools?.() ||
                window.mcpTools ||
                window._mcpTools ||
                []
            );
        },

        /**
         * 从 app-mcp.js 拉取并注册所有 MCP 工具
         */
        async syncFromMcp() {
            try {
                let mcpTools = this._readMcpToolSource();
                // 【修复】与 loadMcpToolsForAI 的竞态：
                // window._mcpTools 由 app-mcp.js 异步填充（Web API / COM 桥 GetAvailableMcpTools）。
                // 若本次 init 早于该加载完成，源为空会被永久锁死（_initialized.mcp = true，不重试），
                // 导致工具注册表整个会话为空：卡片显示裸工具名而非 displayName、必填参数闸门静默失效。
                // 因此：源为空且本会话尚未尝试加载时，主动触发 loadMcpToolsForAI 并 await，再读一次。
                if ((!Array.isArray(mcpTools) || mcpTools.length === 0) && !_initialized.mcpLoadAttempted) {
                    _initialized.mcpLoadAttempted = true;
                    try {
                        const loader = (typeof App.loadMcpToolsForAI === 'function' && App.loadMcpToolsForAI)
                            || (typeof window.loadMcpToolsForAI === 'function' && window.loadMcpToolsForAI);
                        if (loader && !App._mcpToolsLoaded) {
                            await loader();
                        }
                    } catch (e) {
                        ChatLog.warn?.('TOOLS', '[tool-registry] 触发 MCP 工具加载失败', e);
                    }
                    mcpTools = this._readMcpToolSource();
                }
                if (!Array.isArray(mcpTools) || mcpTools.length === 0) {
                    ChatLog.debug?.('TOOLS', '[tool-registry] 未发现 MCP 工具');
                    _initialized.mcp = true;
                    return 0;
                }
                let count = 0;
                for (const mt of mcpTools) {
                    const fullName = mt.name || mt.Name;
                    if (!fullName) continue;
                    // 对 excel-local__/ppt-local__ 等本地工具，注册名去掉来源前缀（LLM 调用和文本解析更稳定）
                    // 保留完整 prefixed name 作为 originalName 传给 executeMcpToolCall
                    const prefixMatch = fullName.match(/^([a-z]+-local)__/i);
                    const name = prefixMatch ? fullName.slice(prefixMatch[0].length) : fullName;
                    const originalName = fullName;
                    const displayName = mt.displayName || mt.DisplayName || name;
                    this.register({
                        name,
                        originalName,
                        displayName,
                        description: mt.description || mt.Description || '',
                        inputSchema: mt.inputSchema || mt.parameters || mt.schema || { type: 'object', properties: {} },
                        // 透传危险标记 + 守卫文案（权威源：excel_tools.json → MCP 桥）
                        isDangerous: !!(mt.isDangerous || (mt.meta && mt.meta.isDangerous)),
                        dangerSeverity: mt.dangerSeverity || (mt.meta && mt.meta.dangerSeverity) || undefined,
                        dangerMessage: mt.dangerMessage || (mt.meta && mt.meta.dangerMessage) || undefined,
                        // 【修复·静默空转】intentKeywords 必须一起透传。
                        // buildMcpToolMeta（app-mcp.js）已经把 excel_tools.json 的 intentKeywords 放进
                        // meta，但这里只搬了 name/displayName/description/inputSchema/isDangerous →
                        // 注册后的工具对象没有 intentKeywords → agent-loop 的"要求覆盖度自检"
                        // （detectUncoveredRequirements）对每个工具都命中 kws.length===0 而 continue
                        // → uncovered 恒为空 → 用户一次列 5 条要求只做 1 条也永远不会被纠正。
                        // 这是"代码改了却完全没效果"的根因：检测依赖的数据源在注册时被丢弃。
                        intentKeywords: Array.isArray(mt.intentKeywords) ? mt.intentKeywords.slice() : undefined,
                        source: 'mcp',
                        async execute(args, ctx) {
                            // 委托给 app-mcp.js 的 executeMcpToolCall（已暴露为 App.executeMcpToolCall）
                            if (App.executeMcpToolCall && typeof App.executeMcpToolCall === 'function') {
                                // 【修复】把 ctx 透传给 executeMcpToolCall（之前完全没传，导致 executeMcpToolCall 拿不到 turn.context.activeSheet 做兜底）
                                // 同时透传 tool 信息（含 inputSchema），executeMcpToolCall 里需要遍历 schema.required
                                // 自动补全"路径/文件"类必填参数，避免每个工具都要单独 patch
                                // 外层 const 只有 name/originalName/displayName（line 536-538）。
                                // description 和 inputSchema 都是 register() 参数对象里的字段名，不是外层 const —— 之前写 shorthand 连续报错 3 次。
                                // 全部用 mt.xxx 显式引用，杜绝 shorthand 误用。
                                const toolMeta = {
                                    name: mt.name || mt.Name,
                                    originalName: mt.name || mt.Name,
                                    displayName: mt.displayName || mt.DisplayName || name,
                                    description: mt.description || mt.Description || '',
                                    inputSchema: mt.inputSchema || mt.parameters || mt.schema || { type: 'object', properties: {} },
                                    isDangerous: !!(mt.isDangerous || (mt.meta && mt.meta.isDangerous)),
                                    dangerSeverity: mt.dangerSeverity || (mt.meta && mt.meta.dangerSeverity) || undefined,
                                    dangerMessage: mt.dangerMessage || (mt.meta && mt.meta.dangerMessage) || undefined
                                };
                                const raw = await App.executeMcpToolCall({
                                    function: { name: originalName, arguments: args || {} }
                                }, { ...(ctx || {}), tool: toolMeta });
                                let result = raw;
                                if (typeof raw === 'string') {
                                    try { result = JSON.parse(raw); } catch { /* keep string */ }
                                }
                                return normalizeMcpResult(result);
                            }
                            if (window.invokeMcpTool) {
                                return await window.invokeMcpTool(name, args);
                            }
                            return { ok: false, error: 'MCP executor not available' };
                        }
                    });
                    count++;
                }
                _initialized.mcp = true;
                ChatLog.info?.('TOOLS', `[tool-registry] 已同步 ${count} 个 MCP 工具`);
                return count;
            } catch (e) {
                ChatLog.warn?.('TOOLS', '[tool-registry] syncFromMcp failed', e);
                _initialized.mcp = true;
                return 0;
            }
        },

        /**
         * 从 app-skill-mgmt.js 拉取并注册所有 Skill
         */
        async syncFromSkills() {
            try {
                const skills = (
                    App.cachedSkills ||
                    window.cachedSkills ||
                    (App.skillMgmt?.list?.()) ||
                    []
                );
                if (!Array.isArray(skills) || skills.length === 0) {
                    ChatLog.debug?.('TOOLS', '[tool-registry] 未发现 Skill');
                    _initialized.skill = true;
                    return 0;
                }
                let count = 0;
                for (const skill of skills) {
                    const name = skill.name || skill.Name || '';
                    if (!name) continue;
                    const displayName = skill.displayName || skill.DisplayName || name;
                    const desc = skill.description || skill.Description || '';
                    const processingMode = skill.processingMode || skill.ProcessingMode || 'VBACode';

                    this.register({
                        name: 'skill.' + name,
                        description: `[Skill] ${displayName}: ${desc}`,
                        inputSchema: { type: 'object', properties: {} },
                        source: 'skill',
                        skillInfo: skill,
                        async execute(args, ctx) {
                            // 【不要在执行期往 content 里追加卡片 JSON】agent-loop 收尾会用 LLM 总结
                            // 覆盖 content，追加的内容会被冲掉（实测：追加后最终消息里卡片仍不出现）。
                            // 卡片改由渲染层补：chat-renderer.buildSkillCardFromTurnTools 渲染正文时
                            // 从工具结果提取 skill_card，渲染到"最终结果"区末尾。
                            return await executeSkill(skill, args, ctx);
                        }
                    });
                    count++;
                }
                _initialized.skill = true;
                ChatLog.info?.('TOOLS', `[tool-registry] 已同步 ${count} 个 Skill`);
                return count;
            } catch (e) {
                ChatLog.warn?.('TOOLS', '[tool-registry] syncFromSkills failed', e);
                _initialized.skill = true;
                return 0;
            }
        },

        /**
         * 初始化：同步所有工具源
         */
        async init() {
            await Promise.all([
                this.syncFromMcp(),
                this.syncFromSkills()
            ]);
            const keys = Array.from(_tools.keys());
            ChatLog.info?.('TOOLS', `[tool-registry] 初始化完成，共 ${_tools.size} 个工具`, keys);
            return _tools.size;
        },

        /**
         * 重置（用于热重载）
         */
        reset() {
            _tools.clear();
            _nameMap.clear();
            _initialized.mcp = false;
            _initialized.skill = false;
        },

        /**
         * 调试信息
         */
        getDebugInfo() {
            const tools = Array.from(_tools.values());
            return {
                total: _tools.size,
                mcp: this.listBySource('mcp').length,
                skill: this.listBySource('skill').length,
                builtin: this.listBySource('builtin').length,
                postExecuteHooks: _postExecuteHooks.length,
                guards: _guards.length,
                // 契约化工具统计
                contractTools: tools.filter(t => t.outputSchema || t.finalizeContent || t.presentCall || t.presentResult).length,
                outputSchemaTools: tools.filter(t => t.outputSchema).length,
                presentTools: tools.filter(t => t.presentCall || t.presentResult).length
            };
        },

        /**
         * 注册 post-execute 钩子。
         * @param {string} name - 钩子名称
         * @param {PostExecuteHook} fn - 钩子函数
         * @param {number} [priority=100] - 优先级
         * @returns {Function} 取消注册的 disposer
         */
        onPostExecute(name, fn, priority = 100) {
            return onPostExecute(name, fn, priority);
        },

        /**
         * 注册单调守卫。
         * @param {string} name - 守卫名称
         * @param {ToolGuard} fn - 守卫函数
         * @param {number} [priority=100] - 优先级
         * @returns {Function} 取消注册的 disposer
         */
        guard(name, fn, priority = 100) {
            return guard(name, fn, priority);
        },

        /**
         * 获取工具的自定义卡片渲染器（如果有）
         * @param {string} toolName - 工具名
         * @returns {Object|null} { presentCall, presentResult } 或 null
         */
        getCardRenderer(toolName) {
            return TOOL_CARD_RENDERERS[toolName] || null;
        },

        /**
         * 计算工具定义级"调用中"UI 投影（DSH presentCall 对齐）。
         * 仅使用工具定义上的 presentCall，不做静态表回退（需回退请用 renderCardIntent）。
         * @param {string} toolName - 工具名
         * @param {Object} args - 调用参数
         * @returns {Object|undefined}
         */
        presentToolCall(toolName, args) {
            const tool = this.get(toolName);
            return presentToolCall(tool, args);
        },

        /**
         * 计算工具定义级"完成"UI 投影（DSH presentResult 对齐）。
         * 仅使用工具定义上的 presentResult，不做静态表回退（需回退请用 renderCardIntent）。
         * @param {string} toolName - 工具名
         * @param {Object} args - 调用参数
         * @param {Object} result - 规范化结果
         * @returns {Object|undefined}
         */
        presentToolResult(toolName, args, result) {
            const tool = this.get(toolName);
            return presentToolResult(tool, args, result);
        },

        /**
         * 校验工具输出是否符合其 outputSchema 声明（DSH output contract 对齐）。
         * @param {string} toolName - 工具名
         * @param {*} value - 待校验的输出值
         * @returns {{valid: boolean, errors: string[]}}
         */
        validateToolOutput(toolName, value) {
            const tool = this.get(toolName);
            if (!tool || !tool.outputSchema) return { valid: true, errors: [] };
            return validateOutput(tool.outputSchema, value);
        },

        /**
         * 获取工具在指定阶段的卡片渲染意图（里程碑三：Card Union）
         * @param {string} toolName - 工具名
         * @param {'call'|'result'} phase - 阶段：'call' 调用时，'result' 完成后
         * @param {Object} args - 工具调用参数
         * @param {Object} [result] - 工具执行结果（仅 phase='result' 时需要）
         * @returns {Object|null} 渲染意图对象或 null（使用默认通用渲染）
         */
        renderCardIntent(toolName, phase, args, result) {
            // 优先使用工具定义上的动态契约（presentCall/presentResult），
            // 回退到静态 TOOL_CARD_RENDERERS 表。动态契约与 DSH ToolDefinition 对齐。
            const tool = this.get(toolName);
            if (tool) {
                if (phase === 'call' && typeof tool.presentCall === 'function') {
                    try { return tool.presentCall(args); } catch (e) {
                        ChatLog.warn?.('TOOLS', `[renderCardIntent] ${toolName} presentCall 异常:`, e);
                    }
                }
                if (phase === 'result' && typeof tool.presentResult === 'function') {
                    try { return tool.presentResult(args, result); } catch (e) {
                        ChatLog.warn?.('TOOLS', `[renderCardIntent] ${toolName} presentResult 异常:`, e);
                    }
                }
            }
            // 静态表按"无 excel-local_/ppt-local_ 等来源前缀"的规范名匹配（LLM 可能输出带前缀的名称）
            const tableName = String(toolName || '').replace(/^[a-z]+-local_{1,2}/i, '');
            const renderer = TOOL_CARD_RENDERERS[tableName] || TOOL_CARD_RENDERERS[toolName];
            if (!renderer) return null;
            try {
                if (phase === 'call' && typeof renderer.presentCall === 'function') {
                    return renderer.presentCall(args);
                }
                if (phase === 'result' && typeof renderer.presentResult === 'function') {
                    return renderer.presentResult(args, result);
                }
            } catch (e) {
                ChatLog.warn?.('TOOLS', `[renderCardIntent] ${toolName} ${phase} 异常:`, e);
            }
            return null;
        },

        /**
         * 读取工具结果里 AI/技能显式声明的渲染器（view.type），校验并归一到白名单。
         * 返回 canonical 类型（'table'|'json'|'code'|'markdown'|'text'|'steps'）或 null。
         * 非法值会被忽略并打日志，绝不把未审过的渲染器交给前端执行。
         * @param {*} result 工具原始结果
         * @returns {string|null}
         */
        getViewType(result) {
            return _digViewType(result);
        },

        /** 渲染器白名单（封闭枚举）：AI 可声明的 view.type 全集，供提示词/代码引用 */
        RENDERER_WHITELIST: RENDERER_WHITELIST,

        /**
         * 从工具调用参数里读取 AI 显式指定的渲染器（前端只读，不发给后端）。
         * 支持 args._view（字符串）或 args.view.type（对象）。校验不过则忽略。
         * 这是"让 AI 在自由对话里选择渲染器"的落地通道：LLM 在调用 excel_* 工具时
         * 附带 _view 提示，前端据此覆盖默认嗅探，但只能选白名单里的类型。
         * @param {Object} [args] 工具调用参数
         * @returns {string|null}
         */
        getViewTypeFromArgs(args) {
            if (!args || typeof args !== 'object') return null;
            const raw = (args.view && typeof args.view === 'object' && args.view.type) || args._view || args.viewType;
            if (!raw) return null;
            const norm = _normalizeViewType(raw);
            if (!norm) {
                ChatLog.warn?.('TOOLS', `[view] args 中未知渲染器 "${raw}"，已忽略（可用：${Object.keys(RENDERER_WHITELIST).join('/')}）`);
            }
            return norm;
        },

        /**
         * 把技能的 outputFormat（text/json/python/vba/markdown）桥接成渲染器 view 类型。
         * 让 SkillCreator 生成的技能只需声明 outputFormat，其结果即按对应方式渲染，
         * 无需再手动在结果里塞 view。仅返回有意义覆盖的值，text/其它返回 null（沿用默认嗅探）。
         * @param {string} format 技能 outputFormat
         * @returns {string|null}
         */
        mapOutputFormatToView(format) {
            if (!format || typeof format !== 'string') return null;
            const f = format.trim().toLowerCase();
            if (f === 'json') return 'json';
            if (f === 'python' || f === 'vba' || f === 'python_script' || f === 'vba_code') return 'code';
            if (f === 'markdown') return 'markdown';
            return null; // text 及未知：不覆盖默认嗅探
        },

        /**
         * 【推荐操作】基于已执行的工具，规则推导一组可继续执行的后续操作。
         * 仅面向"零件装配"型的 excel_* MCP 工具（通用对话路径），返回：
         *   [{ label: String, send: String }, ...]  ——点击时把 send 作为新消息发回通用对话继续执行。
         * 规则表保持静态，可在需要时依据 args/result 分支细化。
         * @param {string} toolName - 已执行工具名（可带 excel-local_/ppt-local_ 等来源前缀）
         * @param {Object} [args] - 本次工具参数
         * @param {*} [result] - 本次工具返回结果
         * @returns {Array<{label:string, send:string}>}
         */
        getRecommendedOps(toolName, args, result) {
            const norm = String(toolName || '').replace(/^[a-z]+-local_{1,2}/i, '');
            const a = (args && typeof args === 'object') ? args : {};

            // ── 返回结果提取辅助（按结果微调推荐）──
            const unwrap = (v, depth) => {
                depth = depth || 0;
                if (!v || typeof v !== 'object' || depth > 4) return v;
                for (const k of ['result', 'data', 'payload']) {
                    if (v[k] !== undefined && typeof v[k] === 'object') return unwrap(v[k], depth + 1);
                }
                if ((v.ok === true || v.success === true) && v.result !== undefined) return unwrap(v.result, depth + 1);
                return v;
            };
            const txt = () => {
                let u = unwrap(result);
                if (typeof u === 'string') { try { const j = JSON.parse(u); if (j && typeof j === 'object') u = j; } catch (_) { ChatLog.debug?.('TOOLS', '[tool-registry] 结果字符串非 JSON，保留原文:', _); } }
                if (typeof u === 'object') { try { return JSON.stringify(u); } catch (_) { return String(u); } }
                return typeof u === 'string' ? u : '';
            };
            const num = (keys, word) => {
                let src = unwrap(result);
                if (typeof src === 'string') { try { src = JSON.parse(src); } catch (_) { src = null; } }
                if (src && typeof src === 'object') {
                    for (const k of keys) {
                        const v = src[k];
                        if (typeof v === 'number' && isFinite(v) && v >= 0) return v;
                        if (typeof v === 'string' && /^\d+$/.test(v)) return parseInt(v, 10);
                    }
                }
                if (word) { const m = txt().match(new RegExp(word + '\\s*[（(]?\\s*(\\d+)')); if (m) return parseInt(m[1], 10); }
                return null;
            };
            const hasTxt = (w) => txt().indexOf(w) > -1;

            try {
                switch (norm) {
                    case 'excel_clear_range': {
                        const clearedContent = !!a.clearContent;
                        const clearedFormat = !!a.clearFormat;
                        const ops = [];
                        if (clearedFormat && !clearedContent) ops.push({ label: '清空内容', send: '清空这一选中区域的内容' });
                        if (clearedContent && !clearedFormat) ops.push({ label: '清除格式', send: '清除这一选中区域的格式' });
                        if (clearedFormat && !clearedContent) ops.push({ label: '重设样式', send: '给这个区域加上边框和填充颜色，美化一下' });
                        if (clearedContent && !clearedFormat) ops.push({ label: '重填数据', send: '帮这个区域重新填入数据' });
                        // 由返回结果微调：确实清空了内容时，追加撤销入口
                        if (hasTxt('内容') || hasTxt('值')) ops.push({ label: '撤销', send: '撤销刚才的清空操作' });
                        return ops;
                    }
                    case 'excel_clear_format':
                        return [
                            { label: '重新设置样式', send: '给这个区域加上边框和填充颜色，美化一下' },
                            { label: '撤销清格式', send: '撤销刚才清除格式的操作' }
                        ];
                    case 'excel_clear_sheet':
                        return [
                            { label: '撤销清空', send: '撤销刚才的清空工作表操作' }
                        ];
                    case 'excel_merge_cells':
                        return [
                            { label: '取消合并', send: '取消这些单元格的合并' },
                            { label: '合并后居中', send: '把选中的单元格合并并居中' }
                        ];
                    case 'excel_unmerge_cells':
                        return [
                            { label: '重新合并并居中', send: '把选中的单元格合并并居中' },
                            { label: '展开填充值', send: '取消合并并把原值铺满整个区域' }
                        ];
                    case 'excel_remove_duplicates':
                    case 'excel_clean_data': {
                        // 纯计算：已返回结果数据，未写回 Excel。推荐把结果落盘以便后续使用
                        return [
                            { label: '写入结果到新表', send: '把刚才返回的结果数据写入一个新的工作表' },
                            { label: '导出为CSV', send: '把刚才返回的结果数据导出为 CSV 文件' }
                        ];
                    }
                    case 'excel_sort_range': {
                        // 排序成功后最高频的后续诉求是"看看排完之后长什么样"（实测截图：
                        // 总结里只有一张模型填不出数据的空表格）。这里直接给一条查看前 10 行的
                        // 推荐。文案刻意避开"排序/按/排名"等词——① 避免再次触发排序类意图
                        // 重复执行；② 命中 detectDisplayIntent 的"仅展示"分支，让模型用
                        // excel_read_range 只读表头+前10行并以表格渲染，而不是又跑一遍分析。
                        let sortCol = '';
                        let orderDesc = true;
                        try {
                            const sb = typeof a.sortBy === 'string' ? JSON.parse(a.sortBy) : a.sortBy;
                            if (Array.isArray(sb) && sb.length) {
                                sortCol = String(sb[0].column || '').trim();
                                orderDesc = sb[0].order !== 'ascending';
                            }
                        } catch (e) { /* 解析失败走通用文案 */ }
                        const viewSend = /^[A-Za-z]{1,3}$/.test(sortCol)
                            ? '读取当前数据的前10行并用表格展示'
                            : `读取当前数据的前10行并用表格展示，第1行是${sortCol}${orderDesc ? '最高' : '最低'}的记录`;
                        return [
                            { label: '恢复原顺序', send: '撤销刚才的排序，恢复原来的顺序' },
                            { label: '查看前10行', send: viewSend }
                        ];
                    }
                    case 'excel_autofit':
                        return [{ label: '同时调行高列宽', send: '同时自动调整选区的行高和列宽' }];
                    case 'excel_apply_conditional_format':
                        return [{ label: '清除条件格式', send: '清除选中区域的条件格式' }];
                    case 'excel_format_range':
                        return [{ label: '清空格式', send: '清空选中区域的格式' }];
                    case 'excel_write_range': {
                        const wrote = num(['affectedRows', 'writtenRows', 'rowCount', 'rows', 'cells'], '写入');
                        const ops = [{ label: '撤销写入', send: '撤销刚才写入选中区域的数据' }];
                        if (wrote != null && wrote > 0) ops.push({ label: `已写 ${wrote} 行表格`, send: '展示一下刚才写入的数据作为表格' });
                        return ops.slice(0, 3);
                    }
                    case 'excel_set_cell_value':
                        return [
                            { label: '撤销修改', send: '撤销刚才对单元格的修改' },
                            { label: '清除该单元格', send: '清除这个单元格的内容' }
                        ];
                    case 'excel_add_formula':
                        return [
                            { label: '转成固定数值', send: '把刚写入的公式区域转成固定数值' },
                            { label: '应用范围扩大', send: '把该公式应用到相邻的更多单元格' }
                        ];
                    case 'excel_fill_data':
                        return [
                            { label: '改为按列填充', send: '改成像刚才那样按列方向填充数据' },
                            { label: '清除填充区域', send: '清除该区域刚填充的数据' }
                        ];
                    case 'excel_insert_row':
                    case 'excel_insert_rows':
                        return [
                            { label: '删除刚插入的行', send: '删除刚才插入的行' },
                            { label: '下方再插入一行', send: '在刚才插入的行下方再插入一行' }
                        ];
                    case 'excel_delete_row': {
                        const del = num(['deletedRows', 'deletedCount', 'affectedRows', 'count'], '删除');
                        return (del != null && del > 0)
                            ? [
                                { label: '撤销删行', send: '撤销刚才删除行的操作' },
                                { label: `共删 ${del} 行确认`, send: `确认一下刚才删除了 ${del} 行数据` }
                              ]
                            : [{ label: '重新选行删除', send: '重新选择要删除的行再删一次' }];
                    }
                    case 'excel_insert_column':
                        return [
                            { label: '删除刚插入的列', send: '删除刚才插入的列' },
                            { label: '右侧再插入一列', send: '在刚才插入列的右侧再插入一列' }
                        ];
                    case 'excel_delete_column': {
                        const del = num(['deletedColumns', 'affectedColumns', 'count'], '删除');
                        return (del != null && del > 0)
                            ? [
                                { label: '撤销删列', send: '撤销刚才删除列的操作' },
                                { label: `共删 ${del} 列确认`, send: `确认一下刚才删除了 ${del} 列数据` }
                              ]
                            : [{ label: '重新选列删除', send: '重新选择要删除的列再删一次' }];
                    }
                    case 'excel_apply_filter':
                        return [
                            { label: '清除筛选', send: '清除选中区域的筛选' },
                            { label: '按其他列筛选', send: '换一列重新设置筛选' }
                        ];
                    case 'excel_clear_filter':
                        return [{ label: '再次筛选', send: '给选中区域重新加一层筛选' }];
                    case 'excel_split_by_column':
                        return [{ label: '用其他分隔符拆分', send: '把选中的列用其他分隔符拆分开' }];
                    case 'excel_split_by_rows': {
                        const base = a.outputMode === 'files'
                            ? [{ label: '改拆到工作表', send: '把选中的内容按行拆到多个工作表' }]
                            : [{ label: '改拆到文件', send: '把选中的内容按行拆分成多个文件' }];
                        const per = num(['perCount', 'rowsPerFile', 'rowsPerSheet', 'perRow'], '份');
                        if (per != null && per > 0) base.push({ label: `每份 ${per} 行`, send: `告诉我拆分结果一共分成了几份` });
                        return base;
                    }
                    case 'excel_save_workbook':
                        return [
                            { label: '导出为 CSV', send: '把当前工作表导出为 CSV 文件' },
                            { label: '关闭工作簿', send: '保存并关闭当前工作簿' }
                        ];
                    case 'excel_close_workbook':
                        return [{ label: '重新打开', send: '重新打开刚才关闭的工作簿' }];
                    case 'excel_open_workbook':
                        return [{ label: '关闭工作簿', send: '关闭刚才打开的工作簿' }];
                    case 'excel_import_csv':
                        return [{ label: '导出为 CSV', send: '把导入的数据再导出成 CSV 文件' }];
                    case 'excel_export_csv':
                        return [{ label: '查看导出的数据', send: '读取刚才导出区域的数据给我看看' }];
                    case 'excel_protect_sheet':
                        return [{ label: '取消保护', send: '取消工作表保护' }];
                    case 'excel_unprotect_sheet':
                        return [{ label: '重新保护', send: '重新保护这个工作表' }];
                    case 'excel_create_sheet':
                        return [
                            { label: '重命名新表', send: '把刚创建的工作表改成别的名称' },
                            { label: '删除新表', send: '删除刚创建的工作表' }
                        ];
                    case 'excel_rename_sheet':
                        return [{ label: '恢复原名', send: '把工作表名称改回去' }];
                    case 'excel_copy_sheet':
                        return [{ label: '删除副本', send: '删除刚才复制出来的工作表副本' }];
                    case 'excel_delete_sheet':
                        return [{ label: '恢复该表', send: '尝试恢复刚删除的工作表' }];
                    case 'excel_create_pivot_table':
                        return [
                            { label: '刷新透视表', send: '刷新刚才的透视表' },
                            { label: '删除透视表', send: '删除刚才创建的透视表' }
                        ];
                    case 'excel_refresh_pivot_table':
                        return [{ label: '新建另一张透视表', send: '再创建一张新的透视表' }];
                    case 'excel_create_chart':
                        return [
                            { label: '改图表标题', send: '修改刚创建图表的标题' },
                            { label: '导出图表为图片', send: '把刚创建的图表导出为图片' },
                            { label: '删除图表', send: '删除刚创建的图表' }
                        ];
                    case 'excel_update_chart_data':
                        return [{ label: '查看图表数据范围', send: '读取该图表的当前数据范围' }];
                    case 'excel_set_chart_title':
                        return [{ label: '恢复默认标题', send: '把图表标题恢复成默认' }];
                    case 'excel_delete_chart':
                        return [{ label: '重新创建图表', send: '为选中的数据重新创建一张图表' }];
                    case 'excel_export_chart_image':
                        return [{ label: '查看导出的图片', send: '把导出的图表图片展示给我看' }];
                    case 'excel_merge_sheets':
                        return [{ label: '查看原始工作表', send: '读取合并前的原始工作表数据' }];
                    case 'excel_join_tables':
                        return [{ label: '换种联接方式', send: '换一种联接方式（内连接或左连接）重新合并这两个表' }];
                    case 'excel_aggregate':
                        return [
                            { label: '换聚合方式', send: '换一种聚合方式（求和或求平均）重新汇总' },
                            { label: '按其他列分组', send: '按另一列分组重新做汇总' }
                        ];
                    default:
                        return [];
                }
            } catch (e) {
                ChatLog.warn?.('TOOLS', `[getRecommendedOps] ${toolName} 异常:`, e);
                return [];
            }
        },
    /**
         * 解析守卫生效模式（turnMode 优先 → thinkingMode 映射 → 安全默认 standard）。
         * 暴露给测试与调试入口，与危险操作守卫共用同一实现。
         * @param {string} [turnMode] - turn.mode（Skill 路径临时置位，如 'expert'）
         * @param {string} [thinkingMode] - 界面思考模式选择器值（AUTO/Quick/Standard/Deep/Expert）
         * @returns {'quick'|'standard'|'deep'|'expert'}
         */
        resolveGuardMode,

        /** 危险操作确认对话框（供技能管线直连 COM 绕过守卫时复用） */
        confirmDangerous: confirmDangerousAction,
        /** VBA 破坏性操作检测（删除/清空类语句） */
        isDestructiveVba,
        /** 读取本轮对危险操作族的用户决定（'allow'|'deny'|null） */
        getTurnDangerousDecision,
        /** 记录本轮对危险操作族的用户决定（供技能管线取消时同步） */
        setTurnDangerousDecision,
    };

    App.ChatV2.toolRegistry = registry;
    ChatLog.info?.('CHAT-V2', '✅ toolRegistry loaded');

    // ── 内置 Post-Execute 钩子 ──────────────────────────────────────

    /**
     * 内置钩子 0：记录本轮已执行的能力（工具组 + 目标区域），供"能力去重"判定重复劳动。
     * 仅在工具执行成功后记录，失败的操作不算数（后续仍需真正执行）。
     */
    onPostExecute('capability-dedup-record', async ({ toolName, args, result, turn }) => {
        try {
            if (!result || result.ok !== true) return;
            const group = capabilityGroupOf(toolName, args);
            if (!group) return;
            if (!turn) return;
            recordCapability(turn, toolName, group, (args && args.sheet) || '', targetRectOf(toolName, args));
        } catch (_e) { /* 去重记录失败不影响主流程 */ }
    });

    /**
     * 内置钩子 1：数据写入后自动刷新图表。
     * 当 excel_write_range / excel_clear_range / excel_set_cell_value / excel_insert_row
     * 执行成功后，异步触发对应工作表的图表刷新（不阻塞主流程）。
     */
    onPostExecute('auto-refresh-charts', async ({ toolName, args, result }) => {
        if (!result.ok) return;
        const WRITE_TOOLS = ['excel_write_range', 'excel_clear_range', 'excel_set_cell_value',
            'excel_insert_row', 'excel_delete_row', 'excel_insert_column', 'excel_delete_column',
            'excel_clear_sheet', 'excel_sort_range', 'excel_apply_filter'];
        if (!WRITE_TOOLS.includes(toolName)) return;

        const sheet = (args && (args.sheet || args.dataSheet || '')) || '';
        // 异步触发，不阻塞主流程
        setTimeout(() => {
            try {
                // 通过 excelBridge 触发图表刷新（如果后端支持）
                if (App.excelBridge && typeof App.excelBridge.Invoke === 'function') {
                    App.excelBridge.Invoke('RefreshCharts', sheet)
                        .then(() => ChatLog.info?.('TOOLS', `[auto-refresh-charts] ${sheet || '当前工作表'} 图表已刷新`))
                        .catch(() => { /* 后端可能不支持 RefreshCharts，静默忽略 */ });
                }
            } catch (_e) { /* 静默 */ }
        }, 200);
    }, 50);  // 优先执行（priority=50）

    ChatLog.info?.('CHAT-V2', '✅ 内置 post-execute 钩子已注册');

    // ── 内置守卫 ────────────────────────────────────────────────────

    /**
     * 内置守卫 1：危险操作拦截。
     * 在 quick/standard 模式下，直接拒绝危险操作，要求切换到 deep/expert 模式。
     * 在 deep/expert 模式下，记录警告但不阻止（由 UI 确认层处理）。
     * 单调性保证：一旦被拒，后续守卫不可翻盘。
     */
    /**
     * 解析守卫生效模式（纯函数，零副作用，可单测）。
     * turnMode 优先（Skill 路径会临时置位 'expert'）；
     * 缺省时映射界面思考模式选择器（AUTO/Quick/Standard/Deep/Expert）；
     * 兜底安全默认 'standard'：保持默认拦截自定义高危脚本，仅显式 深度/专家 解锁。
     */

    /** @type {Object.<string, {severity:string, message:string}>|null} */
    let _dangerCatalog = null;

    /** 从 excel_tools.json 加载危险工具目录（权威源；供未透传 meta 的注册路径兜底） */
    function _ingestDangerTools(cat, tools) {
        if (!Array.isArray(tools)) return;
        for (const t of tools) {
            if (!t || !t.name) continue;
            const n = String(t.name).replace(/^[a-z]+-local__/i, '').toLowerCase();
            const dangerous = !!(t.isDangerous || (t.meta && t.meta.isDangerous));
            if (!n || !dangerous || cat[n]) continue;
            cat[n] = {
                severity: t.dangerSeverity || (t.meta && t.meta.dangerSeverity) || 'medium',
                message: t.dangerMessage || (t.meta && t.meta.dangerMessage)
                    || ((t.displayName || n) + ' 是危险操作，需要确认后才会执行。')
            };
        }
    }

    function ensureDangerCatalog() {
        if (_dangerCatalog) return _dangerCatalog;
        const cat = Object.create(null);
        try {
            const bundled = (typeof globalThis !== 'undefined' && globalThis.__SMARTPILOT_EXCEL_TOOLS__)
                || (typeof window !== 'undefined' && window.__SMARTPILOT_EXCEL_TOOLS__);
            if (bundled && Array.isArray(bundled.tools)) _ingestDangerTools(cat, bundled.tools);
        } catch (_) { /* ignore */ }
        try {
            const mcp = (typeof window !== 'undefined' && Array.isArray(window._mcpTools)) ? window._mcpTools : [];
            _ingestDangerTools(cat, mcp);
        } catch (_) { /* ignore */ }
        _dangerCatalog = cat;
        return cat;
    }

    /**
     * 危险工具信息：权威源 excel_tools.json（isDangerous/dangerSeverity/dangerMessage）。
     * 优先读已注册工具上的透传字段；缺省时查 JSON/MCP 目录（不再维护硬编码名单）。
     */
    function resolveDangerInfo(toolName) {
        const name = String(toolName || '');
        if (!name) return null;
        let tool = null;
        try {
            if (registry && typeof registry.get === 'function') tool = registry.get(name);
        } catch (_) { /* ignore */ }
        if (!tool && App.ChatV2 && App.ChatV2.toolRegistry && typeof App.ChatV2.toolRegistry.get === 'function') {
            tool = App.ChatV2.toolRegistry.get(name);
        }
        if (tool) {
            const dangerous = !!(tool.isDangerous || tool.IsDangerous || (tool.meta && tool.meta.isDangerous));
            if (dangerous || tool.dangerSeverity || (tool.meta && tool.meta.dangerSeverity)) {
                return {
                    severity: tool.dangerSeverity || (tool.meta && tool.meta.dangerSeverity) || 'medium',
                    message: tool.dangerMessage || (tool.meta && tool.meta.dangerMessage)
                        || ((tool.displayName || name) + ' 是危险操作，需要确认后才会执行。')
                };
            }
        }
        const key = name.replace(/^[a-z]+-local__/i, '').toLowerCase();
        return ensureDangerCatalog()[key] || null;
    }

    // ── 用户本轮已拒绝危险操作时的"整表改动"拦截范围 ──────────────────
    // 背景：数据清洗技能的实际改动路径是 excel_clean_data（只算不写）+ excel_write_range 原位写回，
    // 与 excel_delete_row 不是同一个工具；若只按工具名拦截，用户取消 delete_row 后技能仍会写回完成删除。
    // 因此用户一旦拒绝，本轮所有会改动工作表的操作一律拦截，让"取消"真正生效。
    const MODIFY_TOOLS_ON_DENY = [
        'excel_write_range', 'excel_clean_data', 'excel_remove_duplicates', 'excel_clear_range',
        'excel_clear_sheet', 'excel_delete_sheet', 'excel_delete_row', 'excel_delete_column',
        'excel_execute_vba', 'excel_run_python_script', 'excel_split_by_column', 'excel_merge_workbooks',
        'excel_insert_row', 'excel_insert_column', 'excel_sort_range', 'excel_apply_filter',
        'excel_set_cell_value'
    ];

    // 优先级 9（早于 dangerous-tools 的 10）：用户已拒绝 → 直接拦截，不再弹窗
    guard('denied-modify-family', async ({ toolName, turn }) => {
        if (getTurnDangerousDecision(turn) !== 'deny') return;
        if (MODIFY_TOOLS_ON_DENY.indexOf(toolName) < 0) return;
        ChatLog.warn?.('GUARD', `[denied-modify-family] ${toolName} 本轮用户已取消危险操作，自动拦截`);
        return `本次对话中你已取消过危险操作，为尊重你的决定，本轮后续所有会改动工作表的操作（含数据清洗写回）均不再执行。如确需执行，请重新发起对话。`;
    }, 9);

    guard('dangerous-tools', async ({ toolName, args, turn }) => {
        const danger = resolveDangerInfo(toolName);
        if (!danger) return;  // 非危险操作，放行

        // 模板执行（固定安全代码）放行：
        //   excel_run_python_script 的模板模式脚本由系统内置模板生成（描述统计/分组聚合等），
        //   代码固定且安全，不涉及文件/网络/破坏性操作，标准模式下无需拦截。
        //   自定义脚本（LLM 生成）仍按原规则拦截，要求切换到 deep/expert 模式。
        if (toolName === 'excel_run_python_script' && args && args._templateMode === true) {
            ChatLog.info?.('GUARD', `[dangerous-tools] ${toolName} 模板模式执行已放行（固定安全代码）`);
            return;
        }

        // 任务规划：单次调度执行（_tpBypassGuard）或用户显式开启「自动化免确认」（_tpAutoBypass）
        if (window._tpBypassGuard || window._tpAutoBypass) {
            ChatLog.info?.('GUARD', `[dangerous-tools] ${toolName} 任务规划/免确认已放行 (bypass=${!!window._tpBypassGuard}, auto=${!!window._tpAutoBypass})`);
            return;
        }

        // 商业默认：所有思考模式均要求确认（含 deep/expert）。
        // 仅上述 bypass 标志可跳过；「始终询问」控制本轮是否复用 allow。
        const mode = resolveGuardMode(turn?.mode, App?.config?.thinkingMode);
        const alwaysAsk = isDangerAlwaysAsk();
        const prev = getTurnDangerousDecision(turn);
        if (prev === 'allow' && !alwaysAsk) {
            ChatLog.info?.('GUARD', `[dangerous-tools] ${toolName} 本轮已确认危险操作，直接放行（不重复弹窗） mode=${mode}`);
            return;
        }
        if (prev === 'deny') {
            return `本次对话中你已取消过危险操作，为尊重你的决定，本轮后续删除/清空类危险操作均不再执行。如确需执行，请重新发起对话。`;
        }
        let followHint = alwaysAsk
            ? '（已开启「始终询问」：本轮内每次危险操作都会再次确认）'
            : '（本轮对话中后续相同操作将沿用你的决定，不再重复询问）';
        try {
            if (App && App.i18n && typeof App.i18n.t === 'function') {
                followHint = App.i18n.t(alwaysAsk ? 'danger.followAlwaysAsk' : 'danger.followDefault');
            }
        } catch (_e) { /* ignore */ }
        ChatLog.warn?.('GUARD', `[dangerous-tools] ${toolName} 待确认 mode=${mode} thinkingMode=${App?.config?.thinkingMode || 'auto'}`);
        return `【安全守卫】${danger.message}\n\n此操作存在风险，需要您确认后才会执行。\n\n被拦截工具: ${toolName}\n严重程度: ${danger.severity === 'high' ? '⚠️ 高危险' : '⚠️ 中风险'}\n\n${followHint}`;
    }, 10);  // 高优先级，在其他守卫之前执行

    ChatLog.info?.('CHAT-V2', '✅ 内置守卫已注册');
})();