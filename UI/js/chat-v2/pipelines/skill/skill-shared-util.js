/**
 * ============================================================
 * skill-shared-util.js — Skill 管道共享工具函数（从 skill-pipeline.js 拆出）
 * ------------------------------------------------------------
 * 职责：地址/颜色/步骤摘要等纯函数工具，供 skill 家族各模块复用。
 * 挂载：App.ChatV2.skillPipelineModules.util
 * 行为零变更重构：函数体自 skill-pipeline.js 逐字迁移，未改动任何逻辑/正则/文案。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.skillPipelineModules) App.ChatV2.skillPipelineModules = {};
    const MODS = App.ChatV2.skillPipelineModules;
    const SP = MODS;
    const ChatLog = window.ChatLog || console;

    // 根据用户原话强制判定去重模式，纠正小模型在 spec 中丢失 markDuplicate 的问题。
    // 返回：false=真正删除重复行；true=仅高亮标记；null=措辞含糊，交给 LLM/默认（默认仅高亮，安全）。
    function detectCleanDedupeMode(text) {
        const t = (text || '').toLowerCase();
        // 否定措辞（"不删除/不要删除/不删"等）→ 绝不进入删除模式。
        // 否则"仅高亮（不删除、不新增列）"里的"不删除"会被下方"重复.*删"误判为删除意图。
        const hasNegation = /(不删除|不要删除|不删行|不删|别删|不进行删除)/.test(t);
        // ✨ 新增：单独"去重"或"去重数据"没说"高亮/标记" → 默认就是删除重复行（用户心智模型：去重=删重复）。
        // 只有明确说"高亮/标记/仅高亮"才是 markDuplicate=true。
        // 放在否定措辞之后——"去重但不删除"走后面的高亮分支。
        const simpleDedupe = /^(去重|去重数据|去重行|去重一下|去重复|去下重复|去除重复)$/.test(t.trim()) || /(去重|去重复)[^，。,.;;\s]*/.test(t);
        // 明确"仅高亮/只标记"优先于"简单去重"——用户说"去重但不删除/只标记"应走高亮
        if (/(仅高亮|只高亮|高亮标记|只标记|仅标记|只要标记|标记重复|做标记|高亮.*不删|不删除重复|不删重复)/.test(t)) {
            return true; // markDuplicate=true → 仅高亮
        }
        // 简单"去重"没说高亮 → 默认删除（最常见用户意图）
        if (simpleDedupe && !hasNegation) {
            return false; // markDuplicate=false → 删除重复行
        }
        // 明确"删除"意图 → 真正删除（markDuplicate=false）。优先于"高亮/不删除"判定，
        // 因为"删除重复行（真正删除，仅保留首条唯一数据）"这类删除措辞绝不能退回高亮。
        if (/(删除重复行|删除重复|去除重复|去掉重复|清除重复|删掉重复|真正删除|真正去重|真正地去重|去重删除|去重并删除|去重,.*删除|删.*重复)/.test(t) && !hasNegation) {
            return false;
        }
        // 兜底：带"重复...删"的删除措辞（如"把重复的删除"），同样排除否定。
        if (/(重复.*删)/.test(t) && !hasNegation) {
            return false;
        }
        // 明确"高亮"不带"标记"也走高亮（上面已处理过"仅高亮"，这里兜底"高亮"单独出现）
        if (/高亮/.test(t) && /重复|去重/.test(t)) {
            return true;
        }
        return null;
    }
    /**
     * 从后端 MCP 工具返回结果中提取"真实执行步骤"（优先于 LLM 规划步骤）。
     * @param {*} o
     * @returns {Array<string>}
     */
    function findStepsDeep(o) {
        if (o == null) return [];
        if (typeof o === 'string') {
            const s = o.trim();
            if (!s) return [];
            // JSON 字符串 → 解析后递归（只抽 steps/message 等摘要字段）
            if ((s.startsWith('{') || s.startsWith('[')) && s.length < 4000) {
                try { o = JSON.parse(s); } catch { return []; } // 解析失败也丢弃（长原始 JSON 塞进 detail 太丑）
            } else {
                // 纯文本摘要（非 JSON、非 workbook/worksheet/range 这类字段值）
                return [s];
            }
        }
        if (Array.isArray(o)) {
            const strs = o.filter(x => typeof x === 'string' && x.trim() && !/^(workbook|worksheet|sheet|range|headers|data|cleaned|success|error)$/i.test(x.trim()));
            if (strs.length) return strs;
            for (const it of o) {
                const sub = findStepsDeep(it);
                if (sub.length) return sub;
            }
            return [];
        }
        if (typeof o === 'object') {
            // 优先找 steps 数组（最规范的后端返回）
            if (Array.isArray(o.steps)) {
                const arr = o.steps.filter(x => typeof x === 'string' && x.trim());
                if (arr.length) return arr;
            }
            // ✨ 只抽明确的摘要字段（白名单），跳过 workbook/worksheet/range/headers 等元数据
            const SUMMARY_KEYS = ['message', 'summary', 'result', 'output', 'description', 'detail'];
            for (const k of SUMMARY_KEYS) {
                if (typeof o[k] === 'string' && o[k].trim() && !o[k].trim().startsWith('{') && !o[k].trim().startsWith('[')) {
                    return [o[k].trim()];
                }
            }
        }
        return [];
    }

    /**
     * 解析 Skill 的声明式执行模式。
     * 优先级：executionMode 显式声明 > steps 字段（SkillCreator 默认产物）> processingMode 推导（SkillSpec → spec）> fixed。
     * @param {Object} skill
     * @param {*} processingMode - 原始 processingMode 值
     * @param {string} pmStr - 小写字符串形态的 processingMode
     * @returns {'steps'|'spec'|'agent'|'fixed'}
     */
    function resolveExecutionMode(skill, processingMode, pmStr) {
        // 🛡️ 合成空壳（__fromDescriptor=true）一律强制 spec 模式
        // 这些虚拟 Skill 来自 DESCRIPTORS（cleaner/crud/beautifier/formatter 等），
        // 没有 VB 后端注册，必须走 spec 模式调 MCP 原子工具。
        // 走 agent 模式会让 LLM 看到全部 50+ 工具 → 脑补 aggregate/chart → 复杂化。
        if (skill?.__fromDescriptor) {
            const n = String(skill?.name || '');
            const _exemptAgentDomain = /chart|dashboard/i.test(n); // chart/dashboard 空壳保留 agent（需要多步骤 aggregate→create_chart 链）
            if (!_exemptAgentDomain) {
                ChatLog.info?.('PIPELINE', `[skill-pipeline] 合成空壳强制 spec 模式: ${n}`);
                return 'spec';
            }
        }
        const declared = skill?.executionMode ? String(skill.executionMode).toLowerCase() : '';
        if (declared === 'steps') return 'steps';
        if (declared === 'spec' || declared === 'agent' || declared === 'fixed') return declared;
        // 声明式步骤：Skill 自带非空 steps（SkillCreator 默认产物）→ 走 steps 执行器
        const raw = skill && (skill.steps ?? skill.Steps);
        if (Array.isArray(raw) && raw.length > 0) return 'steps';
        // 向后兼容：未声明 executionMode 时，按 processingMode 推导（对齐 v20260818 SkillSpec 兼容逻辑）
        if (pmStr === 'skillspec' || Number(processingMode) === 5) return 'spec';
        // ✨ 关键：按 skill 名称强制 spec 模式——Beautifier/Formatter/Chart/Cleaner/CRUD
        // 这些 skill 内部都是 specEngine.planAndRender + VBA 渲染，不管后端 processingMode
        // 声明的是什么（很多是 VBACode=0），都应该走 spec 路径拿到 buildSpecSteps 多步骨架。
        const skillName = String(skill?.name || skill?.displayName || '');
        if (/SmartTableBeautifier|Beautifier|SmartTableFormatter|Formatter|SmartChartWizard|ChartWizard|SmartDataCleaner|Cleaner|数据清洗|ExcelCRUD|\bCRUD\b|增删改查|DashboardGenerator|Dashboard|看板|驾驶舱/i.test(skillName)) {
            ChatLog.info?.('PIPELINE', `[skill-pipeline] 按 skill 名称强制 spec 模式: ${skillName} (原 processingMode=${processingMode})`);
            return 'spec';
        }
        return 'fixed';
    }

    /**
     * 把 LLM 规划的 spec plan 转成有序步骤字符串数组。
     * 内容全部来自 plan 真实字段（不编造）；由 renderStepsPlanHtml 统一渲染成绿面板。
     * kind: 'beautify' | 'clean' | 'chart'
     */
    /**
     * 统一"步骤=真实工具"框架的工具函数。
     * 思路：美化等 Skill 的每个规划步骤都对应一个真实 MCP 工具调用（excel_format_range / excel_autofit 等），
     * 步骤卡展示真实工具 + 参数 + 后端结果；不使用的手写骨架文案。若本框架无法表达的选项，则整体回退原 VBA 渲染。
     */
    function rgbToHex(arr) {
        if (!Array.isArray(arr)) return null;
        return '#' + arr.slice(0, 3).map(x => ('0' + Math.max(0, Math.min(255, Number(x) | 0)).toString(16)).slice(-2)).join('');
    }
    /**
     * 从区域 addr 中截取第 r1..r2 行（1-based 相对位置，r1<=r2）。
     * 例：rowRangeOf('A3:H20', 1, 1) = 'A3:H3'（选中区域首行作表头）。
     * 注意 r1/r2 是基于 addr 起始行的相对行号，不能当绝对行号直接拼串，
     * 否则会把表头误设到工作簿第 1 行（实测"按选中区域美化只见第一行"）。
     */
    function rowRangeOf(addr, r1, r2) {
        const m = /^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$/.exec(String(addr || '').replace(/\$/g, ''));
        if (!m || r2 < r1) return addr;
        const start = Number(m[2]);
        return m[1] + (start + r1 - 1) + ':' + m[3] + (start + r2 - 1);
    }
    function summarizeToolResult(r) {
        const msg = (r && (r.message || r.result?.message)) || '';
        if (msg) return String(msg).slice(0, 60);
        if (r && (r.ok === true || r.success === true)) return '已应用';
        return r?.error || '执行完成';
    }
    function _colToNum(col) { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }

    /**
     * 判断一行是否为表头行（内容识别，与 VBA IsHeaderRow 语义一致）。
     * header 判定：① 合并标题行（仅首个单元格非空，其余全空）→ 表头；
     *              ② 列头行（整行无纯数字单元格，即全为文本）→ 表头；
     *              数据行（含至少一个数值列）或空行 → 非表头。
     */
    function rowClassify(row) {
        const arr = Array.isArray(row) ? row : [];
        let nc = 0, t = 0;
        const vals = [];
        for (let i = 0; i < arr.length; i++) {
            const s = String(arr[i] === null || arr[i] === undefined ? '' : arr[i]).trim();
            vals.push(s);
            if (s === '') continue;
            nc++;
            if (!/^-?\d+(\.\d+)?$/.test(s)) t++;
        }
        const mergedTitle = vals.length > 0 && vals[0] !== '' && vals.slice(1).every(s => s === '');
        return { nc, textShare: nc ? t / nc : 0, mergedTitle };
    }
    /**
     * 依前若干行内容累计"表头高度"（含合并标题行与列头行），从数据首行开始连续计数，遇数据/空行停。
     * @param {Array<Array>} rows 自上而下的原始行（含可能的标题行/列头行/数据行）
     * @returns {number} 表头行数，至少 1，至多 5
     */
    function detectHeaderRows(rows) {
        if (!Array.isArray(rows) || rows.length === 0) return 1;
        let hr = 0;
        let idx = 0;
        for (const row of rows) {
            idx++;
            // 空行（nc=0，整行无内容）：跳过不终止。常见"标题行 + 空白分隔行 + 列头行"布局，
            // 若在空白处 break 会永远到不了下方的真列头行（实测只美化到标题、列头白底黑字）。
            if (!Array.isArray(row) || row.length === 0) continue;
            const c = rowClassify(row);
            if (c.nc === 0) continue;
            // 表头判定须整行无纯数字单元格（textShare≈1）：列头几乎都是文本，
            // 而数据行通常含至少一个数值列（如业绩/提成），否则文本型数据行
            //（如 员工编号+姓名+部门）会被误判为表头（实测 textShare 0.75 的 EMP001 行）。
            if (c.mergedTitle || c.textShare >= 0.999) {
                // hr 记"最后一行表头层的位置(1基)"而非计数，使表头范围能覆盖中间的空白分隔行。
                hr = idx;
                if (hr >= 5) break;
                continue;
            }
            break; // 数据行（含数值列）→ 表头到此结束
        }
        return hr >= 1 ? hr : 1;
    }

    MODS.util = {
        detectCleanDedupeMode,
        findStepsDeep,
        resolveExecutionMode,
        rgbToHex,
        rowRangeOf,
        summarizeToolResult,
        rowClassify,
        detectHeaderRows,
        _colToNum
    };
})();
