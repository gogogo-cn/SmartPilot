/**
 * ============================================================
 * agent-loop/file-delivery-gate.js — 文件产物交付闸门
 * ============================================================
 * 解决的问题（实测 2026-10-28 19:04 / 19:19，两次截图）：
 *   用户："将当前工作表导出为 CSV 并保存到桌面"
 *   实际：只执行了 excel_save_workbook（存了 .xlsx），
 *         正文却写"工作表已导出为 CSV 文件并保存到桌面，文件路径为：C:\...\xxx.csv"
 *         —— 那个 CSV 从来没有被生成过，路径是模型编的。
 *
 * 根因：模型把"导出 CSV"理解成"先保存一下工作簿"，框架又没有用
 *       "产物类工具是否真的执行过"这一硬事实去校验结论，于是假结论放行。
 *
 * 本模块只做判定/话术（纯函数，便于单测），执行与流程控制放在 agent-loop.js：
 *   - detect(rawMsg, host)      → 用户是否要"某个产物文件"，以及该由哪个工具产出
 *   - isDelivered(turn, spec)   → 该工具本轮是否真的执行过（唯一可信证据）
 *   - buildNudge(spec, turn)    → 再劝一轮时的硬指令
 *   - scrubClaims(text, spec)   → 删除"已导出/已保存 xx.csv"这类没有事实支撑的句子
 *   - describeSuccess/Failure   → 用真实工具结果替换/补充正文
 * ============================================================
 */
(function () {
    'use strict';
    if (typeof window === 'undefined') return;
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const EXPORT_VERB = /(导出|输出|另存|转成|转换成|生成|保存为|保存成|存为|存成)/;
    const CSV_WORD = /(csv|逗号分隔|文本文件)/i;
    // 需要先做别的活（筛选/汇总/排版…）或多表批量 → 交给规划，本闸门不插手
    const NEEDS_PLAN = /(筛选|汇总|统计|分析|图表|画图|透视|美化|格式化|格式设置|条件格式|数字格式|日期格式|清洗|合并|拆分|排序|去重|删除|每个|每张|所有|全部|各个|分别|批量)/;
    // 询问/教学语气（"怎么导出csv""如何导出成csv？"）不是命令，不能强行导出
    const ASKING = /(怎么|如何|怎样|咋样|咋办|为什么|是什么|能不能|可不可以|可以吗|吗[？?]?\s*$|[?？]\s*$)/;

    /** 宿主类型：excel / word / ppt / unknown */
    function hostKind(host) {
        const h = host || (window.SmartPilotHost || null) || (App.SmartPilotHost || null);
        if (!h) return 'excel';   // 与 mcp-dispatcher 一致：宿主信息缺失时按 Excel 处理
        if (h.kind) return String(h.kind).toLowerCase();
        if (h.isExcel) return 'excel';
        if (h.isPowerPoint) return 'ppt';
        if (h.isWord) return 'word';
        return 'unknown';
    }

    /**
     * 识别"产物型请求"。
     * @returns {{id:string, tool:string, label:string, displayName:string}|null}
     */
    function detect(rawMsg, host) {
        const msg = String(rawMsg || '');
        if (!msg || msg.length > 400) return null;
        if (ASKING.test(msg)) return null;
        if (hostKind(host) !== 'excel') return null;
        if (EXPORT_VERB.test(msg) && CSV_WORD.test(msg) && !NEEDS_PLAN.test(msg)) {
            return { id: 'csv', tool: 'excel_export_csv', label: 'CSV 文件', displayName: '导出 CSV 文件' };
        }
        return null;
    }

    /** 产物是否真的产出过（工具执行记录是唯一可信证据，不信模型措辞） */
    function isDelivered(turn, spec) {
        if (!turn || !spec || !spec.tool) return false;
        const target = String(spec.tool).toLowerCase();
        const list = Array.isArray(turn._executedTools) ? turn._executedTools : [];
        if (list.some(n => String(n || '').toLowerCase() === target)) return true;
        const calls = Array.isArray(turn._executedToolCalls) ? turn._executedToolCalls : [];
        if (calls.some(c => c && String(c.name || '').toLowerCase() === target)) return true;
        return false;
    }

    /** 补执行时的参数：只带 sheet（filePath 留空由后端落盘桌面） */
    function buildArgs(spec, turn) {
        const args = {};
        try {
            const ctx = (turn && turn.context) || {};
            const sel = ctx.selection || null;
            const sheet = (sel && (sel.sheetName || sel.sheet))
                || (ctx.activeSheet && (ctx.activeSheet.sheet || ctx.activeSheet.name))
                || '';
            if (sheet) args.sheet = String(sheet);
        } catch (_e) { /* 忽略 */ }
        return args;
    }

    /** 再劝一轮的硬指令（点名工具 + 禁止用保存工作簿冒充 + 禁止编造结论） */
    function buildNudge(spec, turn) {
        const done = (turn && Array.isArray(turn._executedTools) ? turn._executedTools : []).filter(Boolean);
        return {
            text: `⏳ 交付校验未通过：用户要的是「${spec.label}」，不是别的动作。\n\n`
                + `本轮已执行的工具：${done.length ? done.join('、') : '（无）'}\n`
                + `但【${spec.tool}】一次都没有执行，产物文件并不存在。\n\n`
                + `硬约束：\n`
                + `1. excel_save_workbook 只是保存工作簿，不等于导出 ${spec.label}，不能代替 ${spec.tool}；\n`
                + `2. 不许写"已导出/已保存到某路径/文件路径为…"这类结论，除非 ${spec.tool} 真的返回了路径；\n`
                + `3. 现在就调用 ${spec.tool}（不要传 filePath 也可以，后端会保存到桌面），成功后再写最终总结。`
        };
    }

    /**
     * 删除没有事实支撑的"已完成产物"句子（保留其它正常内容）。
     * 只删同时提到该产物关键词或该产物扩展名的句子，避免误删无关内容。
     */
    function scrubClaims(text, spec) {
        const src = String(text || '');
        if (!src) return '';
        const ext = spec && spec.id === 'csv' ? /\.csv\b|csv/i : /(pdf|xlsx?|pptx|docx)/i;
        const CLAIM = /[^。！？!?\n]*(?:已(?:经)?(?:成功)?(?:导出|保存|另存|生成)|导出成功|保存成功|生成成功)[^。！？!?\n]*[。！？!?]?/g;
        return src
            .replace(CLAIM, (m) => (ext.test(m) ? '' : m))
            .replace(/[ \t]+\n/g, '\n')
            .replace(/\n{3,}/g, '\n\n')
            .trim();
    }

    function _pickResultData(res) {
        let data = res && (res.result !== undefined ? res.result : (res.data !== undefined ? res.data : res));
        if (typeof data === 'string') {
            try { data = JSON.parse(data); } catch (_e) { /* 保留原串 */ }
        }
        return data;
    }

    /** 用真实工具返回拼一句交付说明 */
    function describeSuccess(spec, res) {
        const data = _pickResultData(res);
        let path = '';
        if (data && typeof data === 'object') {
            path = data.path || data.filePath || data.file || data.outputPath
                || (data.data && (data.data.path || data.data.filePath)) || '';
        }
        if (path) return `✅ 已导出 ${spec.label}：${path}`;
        const msg = (data && typeof data === 'object' && (data.message || data.msg)) || '';
        return `✅ 已导出 ${spec.label}${msg ? '：' + String(msg).slice(0, 200) : '（工具已成功返回）'}`;
    }

    /** 补执行失败时的诚实说明 */
    function describeFailure(spec, err) {
        const reason = String(err || '未知原因').slice(0, 160);
        return `⚠️ 这次没能真正导出 ${spec.label}（${reason}）。\n`
            + `上文中出现的相关文件路径并不存在，请不要当作已完成。`;
    }

    App.ChatV2.agentLoopFileDeliveryGate = {
        hostKind,
        detect,
        isDelivered,
        buildArgs,
        buildNudge,
        scrubClaims,
        describeSuccess,
        describeFailure
    };
})();
