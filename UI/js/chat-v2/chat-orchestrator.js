/**
 * ============================================================
 * chat-v2/chat-orchestrator.js — L2 Turn 生命周期编排器
 * ============================================================
 * 职责：
 *   - Turn lifecycle：CREATED → WORKING → COMPLETED/FAILED/CANCELED
 *   - 按 message.kind 选择 Pipeline（text / image / doc）
 *   - 管理 AbortController，提供取消能力（取消作为一等公民）
 *   - 消费 Pipeline 的 typed event 流，转发到 SessionStore
 *
 * 不做的事：
 *   - 不构造 messages（由 prompt-builder 做）
 *   - 不直接调用 LLM（由 agent-loop 做）
 *   - 不修改 DOM（由 chat-renderer 做）
 *
 * 依赖：App.ChatV2.{sessionStore, pipelines}
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 【P2-8】Skill 路由/评分阈值常量表——消除散落魔法数字。
     * 原先 score>=8 / kwLen>=5 / confidence>=0.5 等阈值散落在 shouldRouteToSkill 与 LLM route 多处，
     * 调参时极易漏改、造成行为漂移。统一在此声明，语义分组命名，便于一次性调参。
     */
    const ROUTE_T = {
        // —— 关键词长度分档 ——
        KW_MIN_LEN: 2,            // 关键词最小长度（<2 视为单字跳过）
        KW_LONG_LEN: 5,           // 长关键词阈值（强意图）
        KW_MID_LEN: 4,            // 中关键词阈值
        KW_SHORT_LEN: 3,          // 短关键词阈值（具体图表名/操作名）
        KW_SUM_LEN_MIN: 4,        // 2字词合计长度护栏下限

        // —— 关键词命中得分 ——
        SCORE_WEAK_2CHAR: 1,      // 2字词弱权重
        SCORE_SHORT_KW: 3,        // 3字词得分
        SCORE_MID_KW: 4,          // 4字词得分
        SCORE_LONG_KW: 8,         // 5+字词得分

        // —— 语言/代码生成意图加减分 ——
        LANG_MATCH_BONUS: 20,     // 语言指令命中加成
        LANG_MISMATCH_PENALTY: 5, // 语言不匹配惩罚
        CODEGEN_MATCH_BONUS: 10,  // 代码生成类 Skill 命中加成
        CODEGEN_EXEC_PENALTY: 15, // 执行类 Skill 在写代码意图下惩罚

        // —— 强意图 2 字词与路由阈值 ——
        STRONG2_POS_BONUS: 2,     // 强意图2字词位置加成
        STRONG2_SORT_SCORE: 3,    // 排序时强意图2字词最低分
        ROUTE_SCORE_LONG: 8,      // 强匹配：长词命中
        ROUTE_SCORE_LONG2: 5,     // 强匹配：长词(5分)且 _hasLongMatch
        ROUTE_SCORE_MID: 4,       // 强匹配：中词(4分)且 _hasMidKwMatch
        ROUTE_SCORE_SHORT: 3,     // 强匹配：短词(3分)且 _hasShortKwMatch
        ROUTE_SCORE_MIN: 3,       // 强匹配最低分
        ROUTE_LEAD_SCORE: 3,      // 多候选时最佳领先分差

        // —— 快通道与 LLM route ——
        SHORT_CMD_MAX_LEN: 12,                 // 短祈使句快通道最大长度
        LLM_ROUTE_MIN_CONFIDENCE: 0.5,         // LLM route 覆盖 regex 的最低置信度
    };

    /**
     * 【可观测性】Skill 路由的关键判定必须落盘到宿主日志文件
     * （%LOCALAPPDATA%\ExcelAIVSTOAddIn\Logs\chat_log_*.txt）。
     * 背景：前端 ChatLog 只写 WebView2 console（app-bootstrap.js 里 `App.ChatLog = { ... console.* }`），
     * 生产环境完全看不到。实测"Skill 已启用、工具清单里也有 skill_SmartTableBeautifier，
     * 但最终走的是通用 excel_format_range/excel_autofit"——到底卡在 cachedSkills 为空 /
     * 候选评分不够 / 意图守卫 / ctxOk 过滤哪一步，此前无任何落盘证据，只能靠猜。
     * ExcelBridge.LogMessage(level, message) 经 C# 侧 ChatLogger 落盘（tag [FRONTEND]）。
     */
    function _orchHostLog(level, message) {
        try {
            const b = (typeof App !== 'undefined' && App.excelBridge)
                || (typeof window !== 'undefined' && window.excelBridge);
            if (b && typeof b.LogMessage === 'function') {
                const r = b.LogMessage(level || 'INFO', `[orchestrator] ${message}`);
                if (r && typeof r.catch === 'function') r.catch(function () { /* 宿主日志失败可忽略 */ });
            }
        } catch (_e) { /* 宿主日志失败绝不影响主流程 */ }
    }

    // 【#6 收敛】错误文本化已统一到 App.ChatV2.toolRegistryUtil.safeErrorText（见 tool-registry.util.js），
    // 不再在此重复实现。运行时经 App.ChatV2.toolRegistryUtil.safeErrorText 引用。

    // 【#6 收敛】展示文本化统一出口是 App.ChatV2.toolRegistryUtil.safeDisplayText（见 tool-registry.util.js）。
    // 本函数仅作为 tool-registry.util 未加载时的极端兜底别名，不再维护独立序列化逻辑，
    // 避免"改一处漏三处"。调用点 :1713 以 `safeDisplayText || safeDisplayTextWrapper` 形式保留防御。
    function safeDisplayTextWrapper(value, maxLen) {
        const f = App.ChatV2?.toolRegistryUtil?.safeDisplayText;
        if (typeof f === 'function') return f(value, maxLen);
        // 极端兜底：tool-registry.util 未加载时的最小干净序列化
        if (value == null) return '';
        if (typeof value === 'string') return value;
        try { return JSON.stringify(value) || String(value); } catch (e) { return String(value); }
    }

    // computer_* 单步操作：不倾倒 pid/target 等原始字段（如 "pid=21220 target=C:\...\calc.exe"），
    // 按工具动词给一行友好结果；非 computer_* 或无法友好化时返回 null，由调用方走原兜底。
    function computerFriendlyText(d, resultObj) {
        const tool = (d && (d.toolName || '')) || '';
        if (tool.indexOf('computer_') !== 0) return null;
        const o = (resultObj && typeof resultObj === 'object') ? resultObj : {};
        let noun = '';
        if (typeof o.app === 'string' && o.app.trim()) noun = o.app.trim();
        else if (typeof o.name === 'string' && o.name.trim()) noun = o.name.trim();
        else if (typeof o.target === 'string' && o.target.trim()) {
            noun = (o.target.replace(/\\/g, '/').split('/').pop() || '').replace(/\.(lnk|exe|url|bat|cmd)$/i, '');
        }
        const verb = tool === 'computer_open_application' ? '已打开' :
                      tool === 'computer_close_application' ? '已退出' :
                      tool === 'computer_play_media' ? '已触发播放' :
                      tool === 'computer_key_press' ? '已发送按键' : '';
        if (!verb) return null;
        // ⚠ 关键防伪：仅当后端确实成功时才拼"已成功"文案；后端返回任何失败/未命中标记时
        // 一律不覆盖（返回 null），让真实的后端信息（not_found/未检测到/失败）原样透传给用户，
        // 杜绝"关闭失败却显示已退出"这类假确认。
        const textBits = [o && typeof o.result === 'string' ? o.result : '',
                          o && typeof o.message === 'string' ? o.message : '',
                          o && typeof o.resultText === 'string' ? o.resultText : '',
                          o && typeof o.error === 'string' ? o.error : ''].join(' ');
        if (/not_found|未找到|未检测|不存在|找不到|失败|error|exception|无法|assert_failed/i.test(textBits)) return null;
        if (!noun) return verb;
        return verb + ' ' + noun;
    }

    // 把产物类工具返回的 steps 归一化成 string[]。
    // 元素可能是字符串（如 excel_split_by_column 的"新建工作表 '西南'，写入表头 + 39 行数据"），
    // 也可能是 {title, detail, status} 对象（如 spec/buildSpecSteps 产出）→ 统一取可读文案。
    function normalizeStepList(raw) {
        if (!Array.isArray(raw)) return [];
        return raw.map(s => {
            if (typeof s === 'string') return s.trim();
            if (s && typeof s === 'object') {
                const title = (typeof s.title === 'string' && s.title.trim()) ? s.title.trim()
                    : ((typeof s.text === 'string' && s.text.trim()) ? s.text.trim()
                        : ((typeof s.message === 'string' && s.message.trim()) ? s.message.trim() : ''));
                const detail = (typeof s.detail === 'string' && s.detail.trim()) ? s.detail.trim() : '';
                if (title && detail) return title + '：' + detail;
                return title || detail;
            }
            return '';
        }).filter(Boolean);
    }

    // 取"内层业务对象"：
    //   ① 字符串型（宿主 CallLocalTool 原样回传的 JSON 文本）→ 原样返回，由调用方 JSON.parse；
    //   ② 对象型（app-mcp / py 桥等通道会先把内层 JSON 解析成对象再交上来）→ 只认"像业务层"的对象，
    //      绝不把 data 矩阵（数组）当包裹层剥掉，否则统计字段（来源表/行数/表头）全丢。
    // 实测（2026-10-09）：合并工作表的结果以"对象"形态回传时，旧逻辑只在 result 是**字符串**时才下钻，
    // 于是停在信封层 → 找不到 summary/steps → 返回 null → 整坨 JSON 被 safeDisplayText 倒进气泡。
    function _innerEnvelope(o) {
        if (!o || typeof o !== 'object' || Array.isArray(o)) return null;
        if (typeof o.result === 'string' && o.result.trim()) return o.result;
        if (typeof o.data === 'string' && o.data.trim()) return o.data;
        if (o.result && typeof o.result === 'object' && !Array.isArray(o.result)) return o.result;
        if (o.data && typeof o.data === 'object' && !Array.isArray(o.data) &&
            (typeof o.data.summary === 'string' || Array.isArray(o.data.steps) || Array.isArray(o.data.files_created))) return o.data;
        return null;
    }

    // excel_* 产物类工具（拆分/合并/导出/落盘等）的返回 JSON 常带 summary / files_created / steps，
    // 直接走 safeDisplayText 会把原始 JSON（含转义反斜杠路径）整段倾倒进气泡。
    // 这里优先取 summary 拼一行可读结果并附产物路径；识别不到"产物类"结构时返回 null，交回原兜底。
    function excelFriendlyText(toolName, resultObj) {
        const tool = String(toolName || '');
        if (tool.indexOf('excel_') !== 0) return null;
        if (resultObj && resultObj.removedValue) {
            const n = Number(resultObj.removedCount) || (Array.isArray(resultObj.removedRows) ? resultObj.removedRows.length : 0);
            const loc = Array.isArray(resultObj.removedRows) && resultObj.removedRows.length
                ? '（第 ' + resultObj.removedRows.slice(0, 8).join('、') + (resultObj.removedRows.length > 8 ? '…' : '') + ' 行）'
                : '';
            if (!n) return '没有找到「' + resultObj.removedValue + '」，工作表没有改。';
            return '✅ 已移除「' + resultObj.removedValue + '」' + n + ' 行' + loc;
        }
        if (tool === 'excel_apply_conditional_format' && resultObj && resultObj.markLiteral) {
            const where = (resultObj.sheet ? '「' + resultObj.sheet + '」' : '') + (resultObj.range ? ' ' + resultObj.range : '');
            const colorName = resultObj.markColorName || '黄色';
            return '✅ 已把「' + resultObj.markLiteral + '」标成' + colorName + (where ? '：' + where : '') + '<br>等于该值的单元格已涂成' + colorName;
        }
        // ✨ excel_get_sheets：表清单转人话（实测 2026-10-22 22:04：通用兜底把"美化工作簿中的
        // 所有工作表"匹配到本工具后，原始 JSON 数组被整段倾倒进正文冒充回答）
        if (tool === 'excel_get_sheets') {
            let o = resultObj;
            let arr = null;
            for (let depth = 0; depth < 3 && arr == null; depth++) {
                if (Array.isArray(o)) { arr = o; break; }
                if (!o || typeof o !== 'object') break;
                if (Array.isArray(o.result)) { arr = o.result; break; }
                if (typeof o.result === 'string') {
                    try { const p = JSON.parse(o.result); if (Array.isArray(p)) { arr = p; break; } o = p; continue; } catch (_) { break; }
                }
                if (Array.isArray(o.sheets)) { arr = o.sheets; break; }
                if (Array.isArray(o.content)) {
                    const t = o.content.find(function (c) { return c && typeof c.text === 'string'; });
                    if (!t) break;
                    try { o = JSON.parse(t.text); continue; } catch (_) { break; }
                }
                break;
            }
            if (Array.isArray(arr) && arr.length > 0) {
                const names = arr.filter(function (x) { return x && typeof x.name === 'string'; })
                    .map(function (x) { return x.name.trim(); });
                if (names.length) {
                    return '📋 当前工作簿共有 ' + names.length + ' 张可见工作表：' + names.join('、');
                }
            }
            if (Array.isArray(arr) && arr.length === 0) return '当前工作簿没有可见的工作表';
            return null;   // 解析不出结构 → 走原通用路径
        }
        let obj = resultObj;
        // 兼容嵌套包装：桥返回常见 {"result":"<内层JSON字符串>"}，内层才是真正的产物对象
        for (let depth = 0; depth < 3; depth++) {
            if (typeof obj === 'string') {
                const s = obj.trim();
                if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return null;
                try { obj = JSON.parse(s); } catch (e) { return null; }
            }
            if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
            // 已经拿到产物结构 → 不再下钻
            if (typeof obj.summary === 'string' || Array.isArray(obj.files_created) || Array.isArray(obj.steps)) break;
            // 还有一层包装（result/data/payload 里是字符串**或已解析对象**）→ 继续解一层
            const inner = _innerEnvelope(obj);
            if (!inner) break;
            obj = inner;
        }
        if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;

        // ✨ excel_clean_data / excel_remove_duplicates：清洗结果转统计摘要
        //（清洗结果带 headers+data 结构，会被读数据元数据卡误判成"已读取数据 0 行"）
        if (tool === 'excel_clean_data' || tool === 'excel_remove_duplicates') {
            if (obj.success === false) {
                const err = (typeof obj.error === 'string' && obj.error.trim()) ? obj.error.trim() : '未知原因';
                return '⚠️ 清洗失败：' + err;
            }
            const scope = (obj.worksheet ? '「' + obj.worksheet + '」' : '') + (obj.range ? ' ' + obj.range : '');
            const parts = [];
            const _num = (v) => (typeof v === 'number' && Number.isFinite(v) && v > 0) ? v : 0;
            const cr = _num(obj.cleanedRows);
            if (cr > 0) parts.push('处理 ' + cr + ' 行数据');
            const tc = _num(obj.trimmedCells);
            if (tc > 0) parts.push('去首尾空格 ' + tc + ' 处');
            const de = _num(obj.removedEmptyRows);
            if (de > 0) parts.push('删除空行 ' + de + ' 行');
            const dupRows = Array.isArray(obj.highlightedRows) && obj.highlightedRows.length
                ? obj.highlightedRows.filter(n => Number.isFinite(Number(n)))
                : (Array.isArray(obj.duplicateRows) ? obj.duplicateRows.filter(n => Number.isFinite(Number(n))) : []);
            const dup = _num(obj.removedDuplicateRows) || dupRows.length;
            const markOnly = obj.markOnly === true;
            if (dup > 0) {
                const loc = dupRows.length ? '（第 ' + dupRows.slice(0, 8).map(n => String(n)).join('、') + (dupRows.length > 8 ? '…' : '') + ' 行）' : '';
                parts.push((markOnly ? '标黄重复行 ' : '发现重复行 ') + dup + ' 行' + loc + (markOnly ? '，数据未删除' : ''));
            }
            const nd = _num(obj.normalizedDates);
            if (nd > 0) parts.push('规范日期 ' + nd + ' 个');
            const nn = _num(obj.normalizedNumbers);
            if (nn > 0) parts.push('规范数字 ' + nn + ' 个');
            const rr = _num(obj.removedRows);
            const accounted = de + dup;
            if (rr > accounted) parts.push('移除 ' + (rr - accounted) + ' 行');
            else if (rr > 0 && dup === 0 && de === 0) parts.push('移除 ' + rr + ' 行');
            if (parts.length === 0) parts.push('数据本身已符合要求，无需改动');
            const note = (typeof obj.note === 'string' && obj.note.trim()) ? '<br><span style="color:#94a3b8;font-size:12px;">' + obj.note + '</span>' : '';
            const written = typeof obj.writtenTo === 'string' ? obj.writtenTo : '';
            const previewOnly = obj.mode === 'dryRun_preview' || /未修改|dryRun|预览/.test(written);
            const wrote = /已直接修改|已修改并保存/.test(written);
            let title = tool === 'excel_remove_duplicates' ? '去重完成' : '清洗完成';
            if (markOnly) title = dup > 0 ? '已标黄重复行' : '未发现重复行';
            else if (previewOnly) title = '仅预览，工作表未修改';
            else if (wrote) title = '已写回工作表';
            return '✅ ' + title + (scope ? '：' + scope : '') + '<br>' + parts.join(' · ') + note;
        }

        // ✨ excel_check_errors：错误检查结果转人类可读摘要（否则裸 JSON 倾倒进气泡）
        if (tool === 'excel_check_errors') {
            if (obj.success === false) {
                const err = (typeof obj.error === 'string' && obj.error.trim()) ? obj.error.trim() : '未知原因';
                return '⚠️ 检查失败：' + err;
            }
            const errs = Array.isArray(obj.errors) ? obj.errors : [];
            const scope = (obj.sheet ? '「' + obj.sheet + '」' : '当前工作表') +
                ' ' + (obj.scannedRange || '已用区域') + '（共 ' + (obj.scannedCells || 0) + ' 个单元格）';
            if (errs.length === 0) {
                return '✅ 检查完成：' + scope + '，未发现公式错误。';
            }
            const items = errs.map(e => '• <b>' + (e.cell || '?') + '</b>：' + (e.error || '错误值') +
                (e.formula ? '（公式 ' + e.formula + '）' : ''));
            return '⚠️ 检查完成：' + scope + '，发现 <b>' + errs.length + '</b> 个公式错误：<br>' + items.join('<br>');
        }

        // ✨ excel_merge_sheets：合并结果转统计摘要（否则 success+data 矩阵整坨 JSON 倾倒进气泡）
        // 后端回传形态（Services/Mcp/ExcelMcpTools.SheetOps.SplitMerge.vb ExecuteMergeSheets）：
        //   { success:true, sourceSheets:[…], mergedRows:N, dataRows:N, headerPreserved:true,
        //     data:[[表头],[行]…], steps:[…], writtenTo:"合并结果", targetSheet:"合并结果", cleanedEmptySheets:0 }
        // 统计字段可能在外层、也可能被包在 result 里，且 result 既可能是 JSON 字符串也可能是
        // 已解析对象 → 这里按"合并专属字段"逐层下钻，两种形态都能命中（只在信封层就识别 = 失败根因）。
        if (tool === 'excel_merge_sheets') {
            let m = obj;
            for (let i = 0; i < 4 && m && typeof m === 'object' && !Array.isArray(m); i++) {
                if (m.mergedRows !== undefined || m.dataRows !== undefined || Array.isArray(m.sourceSheets)) break;
                const nxt = _innerEnvelope(m);
                if (!nxt) break;
                if (typeof nxt === 'string') {
                    try { m = JSON.parse(nxt); } catch (e) { break; }
                } else { m = nxt; }
            }
            if (!m || typeof m !== 'object' || Array.isArray(m)) return null;
            if (m.success === false || m.isError === true) {
                const err = (typeof m.error === 'string' && m.error.trim()) ? m.error.trim() : '未知原因';
                return '⚠️ 合并失败：' + err;
            }
            const _n = (v) => (typeof v === 'number' && Number.isFinite(v) && v > 0) ? v : 0;
            const srcs = (Array.isArray(m.sourceSheets) ? m.sourceSheets : [])
                .filter(s => typeof s === 'string' && s.trim()).map(s => s.trim());
            const rows = _n(m.mergedRows) || _n(m.dataRows);
            const target = (typeof m.targetSheet === 'string' && m.targetSheet.trim()) ? m.targetSheet.trim()
                : ((typeof m.writtenTo === 'string' && m.writtenTo.trim()) ? m.writtenTo.trim() : '');
            // 源表过多时只点名前 6 张，避免标题行被表名撑爆
            const srcText = srcs.length === 0 ? ''
                : (srcs.length <= 6 ? srcs.join('、') : (srcs.slice(0, 6).join('、') + ' 等 ' + srcs.length + ' 张'));
            const lines = [];
            lines.push('✅ 已合并 ' + (srcs.length > 0 ? srcs.length + ' 张工作表' : '工作表')
                + (srcText ? '（' + srcText + '）' : '')
                + (rows > 0 ? '，共 ' + rows + ' 行数据' : '')
                + (target ? ' → 「' + target + '」' : ''));
            const detail = [];
            if (m.headerPreserved === true) detail.push('表头已保留');
            else if (m.headerPreserved === false) detail.push('表头未保留（首行当数据处理）');
            const cl = _n(m.cleanedEmptySheets);
            if (cl > 0) detail.push('清理空白表 ' + cl + ' 个');
            if (detail.length) lines.push(detail.join(' · '));
            const st = normalizeStepList(m.steps);
            if (st.length) lines.push('<span style="color:#94a3b8;font-size:12px;">' + st.join(' · ') + '</span>');
            return lines.join('<br>');
        }

        const summary = typeof obj.summary === 'string' ? obj.summary.trim() : '';
        const files = Array.isArray(obj.files_created) ? obj.files_created.filter(f => typeof f === 'string' && f.trim()) : [];
        const failed = Array.isArray(obj.failed) ? obj.failed.filter(Boolean) : [];
        // 【新增】产物类工具常把执行过程写在 steps 里（拆分/合并/导出每张表的写入明细）。
        // 它本质就是一串子步骤，之前不在识别范围内 → 返回 null → 整段原始 JSON 倾倒进气泡。
        const steps = normalizeStepList(obj.steps);
        // 非"产物类"结果（如读数据的 headers/rows）交回原兜底，避免抢走元数据卡片的渲染
        if (!summary && files.length === 0 && steps.length === 0) return null;

        if (obj.success === false) {
            const err = (typeof obj.error === 'string' && obj.error.trim()) ? obj.error.trim() : '未知原因';
            return '⚠️ 执行失败：' + err;
        }
        const lines = [];
        if (summary) lines.push('✅ ' + summary);
        if (files.length) lines.push(files.map(f => '📄 ' + f.replace(/\\/g, '/')).join('<br>'));
        // steps 渲染成编号子步骤；无 summary 时用 created/skipped 给产物概览
        if (steps.length) {
            const created = Array.isArray(obj.created) ? obj.created.filter(c => typeof c === 'string' && c.trim()) : [];
            const skipped = Array.isArray(obj.skipped) ? obj.skipped.filter(c => typeof c === 'string' && c.trim()) : [];
            if (!summary && created.length) {
                lines.push('✅ 新建 ' + created.length + ' 个工作表：' + created.join('、'));
            }
            if (!summary && skipped.length && created.length === 0) {
                lines.push('✅ 拆分表已存在（' + skipped.join('、') + '），未重复创建');
            } else if (!summary && skipped.length) {
                lines.push('ℹ️ 已跳过同名表：' + skipped.join('、'));
            }
            lines.push(steps.map((s, i) => (i + 1) + '. ' + s).join('<br>'));
        }
        if (failed.length) lines.push('⚠️ 未处理：' + failed.map(f => typeof f === 'string' ? f : JSON.stringify(f)).join('、'));
        return lines.join('<br>');
    }

    function pptFriendlyText(toolName, resultObj) {
        const tool = String(toolName || '');
        if (tool.indexOf('ppt_') !== 0) return null;
        // 表格格式化类 ppt_ 工具返回 {summary,…}，若直接落到函数尾部会被误报成"✅ PPT 已生成并打开"。
        // 这类工具窄路优先走内层 summary；其余 ppt_ 工具行为保持不变（不扩大改动面）。
        if (/^ppt_format(?:_selected)?_table_cell$/i.test(tool)) {
            const p = unwrapEnvelopeToPayload(resultObj);
            if (!p || typeof p !== 'object' || Array.isArray(p)) return null;
            if (p.success === false || p.isError === true) {
                const err = (typeof p.error === 'string' && p.error.trim()) ? p.error.trim() : '未知原因';
                return '⚠️ 执行失败：' + err;
            }
            const s = (typeof p.summary === 'string' && p.summary.trim()) ? p.summary.trim()
                : ((typeof p.message === 'string' && p.message.trim()) ? p.message.trim() : '');
            return s ? '✅ ' + s : null;
        }
        let obj = resultObj;
        // 兼容嵌套包装
        for (let depth = 0; depth < 3; depth++) {
            if (typeof obj === 'string') {
                const s = obj.trim();
                if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return null;
                try { obj = JSON.parse(s); } catch (e) { return null; }
            }
            if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
            const inner = (typeof obj.result === 'string' && obj.result.trim()) ? obj.result
                : ((typeof obj.data === 'string' && obj.data.trim()) ? obj.data : null);
            if (!inner) break;
            obj = inner;
        }
        if (!obj || typeof obj !== 'object') return null;

        if (obj.success === false) {
            const err = (typeof obj.error === 'string' && obj.error.trim()) ? obj.error.trim() : '未知原因';
            return '⚠️ PPT 生成失败：' + err;
        }
        const path = typeof obj.path === 'string' ? obj.path : '';
        const msg = typeof obj.message === 'string' ? obj.message.trim() : '';
        const fileMatch = msg.match(/文件[：:]\s*([^\n]+)/);
        const themeMatch = msg.match(/主题[：:]\s*([^\n]+)/);
        const pageMatch = msg.match(/页数[：:]\s*([^\n]+)/);
        const lines = ['✅ PPT 已生成并打开'];
        if (fileMatch) lines.push('📄 ' + fileMatch[1].trim());
        else if (path) lines.push('📄 ' + path.replace(/\\/g, '/'));
        if (themeMatch) lines.push('🎨 主题：' + themeMatch[1].trim());
        if (pageMatch) lines.push('📑 ' + pageMatch[1].trim());
        return lines.join('<br>');
    }

    // 拆多层信封，取出"真正的业务对象"：
    //   {"success":true,"result":{"success":true,"summary":"…"},"taskCompleted":true} → 内层。
    // result/data 既可能是对象，也可能是 JSON 字符串（executeLocalExcelTool/CallLocalTool 的包装）。
    // 只带壳字段（success/taskCompleted/method…）就继续往里钻；一旦出现业务字段即返回。
    function unwrapEnvelopeToPayload(value) {
        let obj = value;
        for (let depth = 0; depth < 4; depth++) {
            if (typeof obj === 'string') {
                const s = obj.trim();
                if (!s || (s.charAt(0) !== '{' && s.charAt(0) !== '[')) return null;
                try { obj = JSON.parse(s); } catch (e) { return null; }
                continue;
            }
            if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
            if (typeof obj.summary === 'string' || typeof obj.scope === 'string' ||
                Array.isArray(obj.changes) || (obj.applied && typeof obj.applied === 'object') ||
                typeof obj.path === 'string' || Array.isArray(obj.steps)) return obj;
            const inner = (obj.result != null) ? obj.result : ((obj.data != null) ? obj.data : null);
            if (inner == null) return obj;
            obj = inner;
        }
        return (obj && typeof obj === 'object') ? obj : null;
    }

    // Word 宿主的工具结果友好化。
    // 背景：本文件的收尾只认 computer_/excel_/ppt_ 三个前缀，**没有 word_**，
    // 于是 word_* 的结果会掉到 safeDisplayTextWrapper 被整坨 JSON.stringify 成正文 ——
    // 实测 word_format_selected_table_cell 执行成功后气泡里出现
    // {"success":true,"result":{…"summary":"表格#1（19行×9列）整表已修改…"},"taskCompleted":true}（用户反馈"没有渲染"）。
    // 这里拆开外层信封，优先把内层 summary/message 作为一行结论。
    function wordFriendlyText(toolName, resultObj) {
        const tool = String(toolName || '');
        if (tool.indexOf('word_') !== 0) return null;
        const obj = unwrapEnvelopeToPayload(resultObj);
        if (!obj || typeof obj !== 'object' || Array.isArray(obj)) return null;
        if (obj.success === false || obj.isError === true) {
            const err = (typeof obj.error === 'string' && obj.error.trim()) ? obj.error.trim() : '未知原因';
            return '⚠️ 执行失败：' + err;
        }
        const summary = (typeof obj.summary === 'string' && obj.summary.trim()) ? obj.summary.trim()
            : ((typeof obj.message === 'string' && obj.message.trim()) ? obj.message.trim() : '');
        if (summary) return '✅ ' + summary;
        return null;   // 无法友好化 → 交回统一兜底
    }

    // ============================================================
    // 【写类工具裁决闸门】——命中项是"候选"，不是"命令"
    //
    // 背景：McpDispatcher 排在 LLM 之前，只做工具名/关键词匹配，看不懂意图。
    // "帮我写一段 VBA 代码：拆分工作簿" 里的"拆分工作簿"会命中 excel_split_workbook
    // 并**真实执行**（落盘产出新文件），而用户只想要代码 —— 顺序错了：先执行、后理解。
    //
    // 设计（按副作用分级，不搞全量 AI 裁决）：
    //   - 只读/瞬时类工具（读数据、截图、打开程序…）：保持确定性短路，零延迟。
    //   - 写入/产物类工具（拆分/合并/导出/写单元格/落盘…）：命中只作为**候选**，必须先过裁决。
    //
    // 裁决优先级：规则（零延迟、确定性覆盖常见歧义）→ LLM 二次确认（可选）→ 保守沿用现有能力。
    // 之所以不把每条命令都交给 LLM：本地小模型会超时/不稳，而"打开计算器"这类命令多等几秒
    // 是纯粹的体验损失；歧义只集中在有副作用的工具上，只在该处付费。
    // ============================================================

    // 副作用动词：工具名按 _ 切分后，任一段命中即视为写类工具
    // （excel_read_range/get_cell_value/list_sheets/screenshot/open_app 均不含，保持短路）
    const SIDE_EFFECT_VERBS = new Set([
        'split', 'merge', 'export', 'save', 'write', 'create', 'rename', 'delete', 'remove',
        'replace', 'format', 'beautify', 'clean', 'dedup', 'dedupe', 'pivot', 'insert',
        'append', 'set', 'fill', 'update', 'convert', 'generate', 'build', 'move', 'copy', 'add'
    ]);

    function isSideEffectTool(toolName) {
        const t = String(toolName || '').toLowerCase();
        if (!t) return false;
        return t.split('_').some(seg => SIDE_EFFECT_VERBS.has(seg));
    }

    // ── 【副作用分级】只产出文件 vs 改动宿主内容 ────────────────────────
    // 初衷：isSideEffectTool 把 export/save 和 write/delete/format 塞进同一个布尔闸门，
    // 而这两类的冒犯性完全不同：
    //   · 改动类（write/delete/insert/format/update/replace/set/fill…）→ 破坏用户文档，
    //     必须严，漏判代价高；
    //   · 产出类（export/save + 无改动动词）→ **不改原文档**，只是多一份文件，最坏情况
    //     是桌面多一个 PDF（删掉即可），而且空白文档也能导出（见 2026-10-06 问答）。
    // 于是"导出"被同一套闸门对待时，只要规则表没覆盖该措辞，就掉进 LLM 二次裁决，
    // 被小模型判成 other → 放行 agent-loop → 自由发挥（实测：导出前先美化表格）。
    // 这里显式分级：产出文件类不需要 LLM 再裁决一次是否该执行。
    const DESTRUCTIVE_VERBS = new Set([
        'write', 'delete', 'remove', 'format', 'beautify', 'clean', 'dedup', 'dedupe',
        'update', 'replace', 'set', 'fill', 'insert', 'append', 'move', 'rename',
        'merge', 'split', 'pivot', 'clear'
    ]);
    const ARTIFACT_VERBS = new Set(['export', 'save', 'saveas']);

    /**
     * 是否为"只产出文件、不改动宿主内容"的工具（excel_export_csv / word_export_pdf / ppt_save…）。
     * 判定：工具名含产出动词 **且** 不含任何改动类动词。
     */
    function isArtifactOnlyTool(toolName) {
        const segs = String(toolName || '').toLowerCase().split('_');
        if (!segs.some(s => ARTIFACT_VERBS.has(s))) return false;
        return !segs.some(s => DESTRUCTIVE_VERBS.has(s));
    }

    // 规则闸门 1：消息明确在"要代码/脚本/宏" → 只出代码，绝不执行
    // 补充 写/生成 + vba|vb 变体：'帮我写个VBA把A1加粗' 此前不命中任何一条，
    // 会掉到下方"格式设置命令"规则被误判成执行命令（本该只出代码）。
    const CODE_REQUEST_RE = /写.*代码|生成.*代码|给我.*代码|要.*代码|来.*代码|写.*脚本|生成.*脚本|写.*宏|生成.*宏|写个.*宏|写.*vba|生成.*vba|写.*vb\b|生成.*vb\b|教我|怎么写|如何写|示例代码|代码示例|vba\s*代码|vb\s*代码|python\s*代码|\bcode\b/i;
    // 规则闸门 2：疑问/示例/假设语气 → 不视为执行命令
    // 补充 怎么做/怎么办/怎样/咋… ：像"导出PDF怎么做？"这类在问方法的句子必须判 other，
    // 否则新增的"文档级导出"执行规则会把提问当成命令，直接落盘一份 PDF。
    const NON_COMMAND_TONE_RE = /怎么写|如何写|怎么做|怎么办|咋做|咋办|怎样|教我|示例|例子|假设|假如|如果|能不能|可以吗|是否|什么|为什么|解释|介绍|说明/i;
    // 规则闸门 3：明确执行命令标记 —— 含命令式结构 + 副作用动词 + 产物/路径关键词
    // 这类确定性命令（"将X导出为Y"、"把数据保存到Z"）不走 LLM 裁决，直接判定 command
    const EXEC_COMMAND_PATTERNS = [
        /(?:将|把|给|帮)[^。？?\n]{2,30}(?:导出|保存|写入|输出|生成|创建|另存|存为)[^。？?\n]{0,50}(?:文件|csv|xlsx|txt|图片|图表|桌面|文件夹|目录|路径|工作表|数据)/i,
        /(?:导出|保存|写入|输出|生成|创建|另存|存为)[^。？?\n]{2,50}(?:文件|csv|xlsx|txt|图片|图表|桌面|文件夹|目录|路径|工作表|数据)/i,
        // 宿主文档级导出（"将当前文档导出为 PDF" / "把演示文稿另存为 PDF"）：
        // PDF/文档/幻灯片本身就是产物，不会像 Excel 那样跟"文件/桌面/工作表"等词，
        // 因此上面的产物关键词白名单命中不了 → 掉进 LLM 二次裁决，被小模型判成
        // "既非命令也非代码" → 放行通用路由 → agent-loop 自由发挥（实测 2026-10-06 00:59：
        // 导出前先跑了一次 word_format_selected_table_cell 美化表格，纯属多余）。
        // 这里是"命令式 + 导出动词 + 文档/PDF 产物"的确定性标记，直接判定 command 走硬短路。
        /(?:将|把|给|帮)?[^。？?\n]{0,20}(?:导出|输出|另存|转换|转成|保存|存为|存成)\s*(?:为|成|到)?\s*(?:pdf|文档|幻灯片|演示文稿|图片|png|jpe?g|word|excel|pptx?|docx?)/i,
        /(?:保存到|存到|放到|写入到|输出到)[^。？?\n]{1,30}(?:桌面|文件夹|目录|路径|文件)/i,
        /(?:按|根据)[^。？?\n]{2,30}(?:汇总|统计|求和|计算|分组|排序|筛选|分析|制作|创建|生成)[^。？?\n]{0,30}(?:图表|饼图|柱图|折线图|透视|报表)/i,
        // PPT/幻灯片/演示文稿 生成命令（"把刚才的结果做成PPT"、"将这份数据生成幻灯片"）
        /(?:把|将|给|帮)[^。？?\n]{2,30}(?:做|做成|制作|生成|创建|出|转换|转成)[^。？?\n]{0,20}(?:PPT|ppt|幻灯片|演示文稿)/i,
        /(?:做|做成|制作|生成|创建|出)[^。？?\n]{0,10}(?:PPT|ppt|幻灯片|演示文稿)/i,
        // 数据整形命令（"把这份表按'销售区域'列拆分成多个工作表"、"按部门把当前数据拆分到多个工作表"）：
        // 句中既没有"导出/保存/生成"等落盘动词，又常超过短祈使句阈值（>12 字），
        // 此前只能落到 LLM 裁决，实测被本地小模型误判成 code → __forceCodePipeline
        // → 只吐 VBA 不执行（测试用例"按销售区域列拆分成多个工作表"就是这么失败的）。
        // 两条都用句首锚定：目标是挡掉"写VBA把表拆分成多个工作表"这类**要代码**的句子
        // （带"写…宏/代码/脚本"的句子更已被上方 CODE_REQUEST_RE 规则闸门 1 先一步拦下）。
        /^(?:把|将|给|帮)[^。？?\n]{2,30}(?:拆分|拆成|拆为|分成|分为|分割|合并|汇合|去重|分列|转置)[^。？?\n]{0,30}(?:工作表|工作簿|多份|多个|一张表|一个表|sheet|文件)/i,
        /^(?:按|根据)[^。？?\n]{2,30}(?:拆分|拆成|拆为|分成|分为|分割|合并|汇合|去重|分列|转置)[^。？?\n]{0,30}(?:工作表|工作簿|多份|多个|一张表|一个表|sheet|文件)/i,
        // 工作表级合并/拆分的"自然语序"（"将选中工作表进行数据合并"、"把这几张表合并成一张"）：
        // 上面两条要求动作词后面紧跟产物词（工作表/多份/一张表），而中文更常见的写法是把动作词放句尾
        // （"…工作表…合并"），于是三条规则全不命中 → 落到 LLM 二次裁决，小模型易判 other →
        // 放行 agent-loop 自由发挥（实测 2026-10-07 19:26：合并请求被演绎成"读数→聚合→画饼图"）。
        /(?:合并|汇合|合并成|合并到|汇总成)[^。？?\n]{0,20}(?:工作表|sheet|多张表|多表|数据|这几张表|选中表|所选表)/i,
        /(?:工作表|sheet|多张表|多表|这几张表|选中表|所选表)[^。？?\n]{0,20}(?:合并|汇合|合并到|汇总成)/i,
        // 格式设置命令（"给 A1:D1 加粗并水平居中"、"把表头加粗"、"选中区域设成百分比格式"）：
        // 与上面的拆分/合并同理——句中既无"导出/保存/生成"等落盘动词，又常超过短祈使句阈值(>12 字)，
        // 此前只能落到 LLM 裁决，实测被本地小模型误判成 code → __forceCodePipeline
        // → 只吐 excel_format_range(...) 命令不执行（实测"给 A1:D1 加粗并水平居中"就是这么失败的）。
        // 句首锚定 + 代码意图已由上方 CODE_REQUEST_RE 规则闸门 1 先一步拦下。
        /^(?:把|将|给|帮|对|为)[^。？?\n]{0,30}(?:加粗|加斜|倾斜|下划线|水平居中|垂直居中|居中|左对齐|右对齐|顶端对齐|底端对齐|格式化|设置格式|调整格式|统一格式|美化|字号|字体|底色|背景色|填充色|字体颜色|边框|描边|保留小数|百分比|货币|数字格式|日期格式|自动换行|行高|列宽|条件格式|数据条|色阶|图标集|合并居中)/i,
        // 句首就是美化/格式化（"美化工作簿中的所有工作表，统一为商务蓝"）。
        // 上一条要求句首是把/将/给，这条对不上，就会落到 LLM，小模型判成 code。
        /^(?:美化|格式化|排版)[^。？?\n]{0,80}(?:工作表|工作簿|表格|sheet|商务蓝|深蓝)/i,
        // 「统一为/转为商务蓝」不从美化起句时同样是要执行的预设，不是要 VBA。
        /(?:统一(?:为|成)|转为|转成|改成|换成|设为)[^。？?\n]{0,12}(?:商务蓝|深蓝)/
    ];

    /**
     * 对一个"写类工具命中"做裁决。
     * @returns {Promise<{verdict:'command'|'code'|'other', source:string, reason?:string}>}
     *   command = 确实要执行；code = 只想要代码文本（交给 code-pipeline）；other = 都不是（交回通用路由）
     */
    async function adjudicateWriteHit(opts) {
        const message = String(opts?.message || '').trim();
        if (!message) return { verdict: 'command', source: 'rule', reason: '空消息' };

        // 1) 规则闸门：明确要代码 → code
        if (CODE_REQUEST_RE.test(message)) {
            return { verdict: 'code', source: 'rule', reason: '消息含代码生成措辞' };
        }
        // 2) 规则闸门：疑问/示例语气 → other
        if (NON_COMMAND_TONE_RE.test(message)) {
            return { verdict: 'other', source: 'rule', reason: '消息为疑问/示例语气' };
        }
        // 3) 短祈使句快通道：短且无问号/逗号（如"拆分工作簿"）→ 直接执行，不为确定性命令付 LLM 延迟
        if (message.length <= ROUTE_T.SHORT_CMD_MAX_LEN && !/[？?，,。；;]/.test(message)) {
            return { verdict: 'command', source: 'rule', reason: '短祈使句' };
        }
        // 3.5) 明确执行命令：含命令式结构 + 副作用动词 + 产物/路径关键词 → 直接执行
        // 这类确定性命令（"将当前工作表导出为 CSV 并保存到桌面"、"按销售额汇总出饼图"）
        // 不走 LLM 二次确认，小模型容易把执行命令误判成代码询问
        for (const pat of EXEC_COMMAND_PATTERNS) {
            if (pat.test(message)) {
                return { verdict: 'command', source: 'rule', reason: '明确执行命令标记' };
            }
        }
        // 3.8) 【产出文件免裁决】命中项是"只导出/保存文件、不改宿主内容"的工具时，
        // 已过"要代码"与"疑问语气"两道规则闸门，接下来就是执行——不必再问 LLM。
        // 代价是非对称的：直接执行的后果只是多一个可删的文件；而问 LLM 有可能被小模型
        // 误判成 other → 放行 agent-loop 自由发挥（实测：导出前先多跑一次表格美化）。
        if (isArtifactOnlyTool(opts?.toolName)) {
            return { verdict: 'command', source: 'rule', reason: '产出文件类工具免 LLM 裁决' };
        }
        // 4) LLM 二次确认（候选交给模型看一眼）；不可用/失败 → 沿用现有能力直接执行
        const llmIo = App.ChatV2?.llmIo;
        if (!llmIo || typeof llmIo.askJson !== 'function') {
            return { verdict: 'command', source: 'no-llm', reason: 'LLM IO 未就绪' };
        }
        try {
            const ask = llmIo.askJson({
                messages: [
                    {
                        role: 'system',
                        content: '你是意图裁决器。用户消息命中了一个**有副作用的工具**候选，但命中不等于用户想执行。'
                            + '只返回 JSON：{"intent":"command|code|other","confidence":0.0~1.0,"reason":"简短理由"}。'
                            + 'command=用户确实要执行这个操作；code=用户只想要代码/脚本/宏的文本、不要执行；'
                            + 'other=既不是要执行也不是要代码（问答/讨论/举例/无关）。'
                    },
                    {
                        role: 'user',
                        content: `用户消息：${message}\n候选工具：${opts?.toolName || ''}（${opts?.displayName || ''}）\n`
                            + '注意：只有用户**明确索要代码/脚本/宏的文本本身**（如"帮我写…""生成一段…代码"）时才判 code；'
                            + '若用户是在描述一件要他去做的事（如"按销售区域拆分成多个工作表""把数据导出到桌面"），一律判 command。',
                    }
                ],
                signal: opts?.signal,
                config: App.config
            });
            const raced = await Promise.race([
                ask,
                new Promise(resolve => setTimeout(() => resolve(null), 4000))
            ]);
            const json = raced && raced.json;
            const intent = json ? String(json.intent || '').toLowerCase() : '';
            if (intent === 'command' || intent === 'code' || intent === 'other') {
                return { verdict: intent, source: 'llm', confidence: json.confidence, reason: json.reason };
            }
            ChatLog.warn?.('ORCH', '[写类工具裁决] LLM 未返回可用 intent（沿用现有能力直接执行）');
        } catch (e) {
            ChatLog.warn?.('ORCH', `[写类工具裁决] LLM 裁决异常（沿用现有能力直接执行）: ${e?.message || e}`);
        }
        return { verdict: 'command', source: 'llm-fallback', reason: 'LLM 不可用/超时' };
    }

    // 截屏工具返回 "base64=<PNG base64>" 时，构造气泡内可展示的 <img>；否则返回 null。
    // 识别 data:image/ 与纯 base64（iVBOR=PNG 特征头）两种形式。
    function buildScreenCaptureImgHtml(toolName, resultObj) {
        // 全屏 + 按窗口截图都会返回 "path=..."/"base64=..."；focus_window 仅当它作为
        // "切到X并截图"链的首步、结果被第二步的截图路径覆盖时才会含 path → 故一并放行，无 path 自然返回 null。
        if (toolName !== 'computer_screen_capture' && toolName !== 'computer_capture_window' && toolName !== 'computer_focus_window') return null;
        let raw = null;
        if (typeof resultObj === 'string') raw = resultObj;
        else if (resultObj && typeof resultObj.result === 'string') raw = resultObj.result;
        else if (resultObj && typeof resultObj.message === 'string') raw = resultObj.message;
        if (!raw) return null;

        // ── 情形 A：后端 returnFormat="path"，本地临时 PNG，用 file:// 引用（绕开超长 base64）──
        let pathVal = null;
        if (raw.startsWith('path=')) pathVal = raw.slice(5);
        else {
            const pi = raw.indexOf('path=');
            if (pi >= 0) pathVal = raw.slice(pi + 5);
        }
        if (pathVal) {
            pathVal = pathVal.trim();
            // 去掉可能残留的消息后缀（路径后若带空格+文案则截到第一个空白）
            const rel = pathVal.split(/\s+/)[0];
            // 用 WebView2 虚拟主机同源引用 Temp 下的截图文件（绕开 file:// 跨源与超大 base64）
            const fileName = rel.split(/[\\/]/).pop();
            const src = 'https://screenassets.local/' + encodeURIComponent(fileName);
            ChatLog.info?.('ORCH', `[screen-capture] path-mode local=${rel} src=${src}`);
            const safeSrc = String(src).replace(/[<>"']/g, '');
            return '<div class="screen-capture-wrap" style="position:relative;">' +
                '<img class="screen-capture-img" src="' + safeSrc + '" alt="屏幕截图" ' +
                'onerror="this.outerHTML=\'&lt;div style=&quot;padding:10px;color:#666;font-size:13px;&quot;&gt;⚠️ 截屏图片加载失败，原始文件：&lt;/div&gt;\'" ' +
                'style="max-width:100%;max-height:360px;border:1px solid var(--border-color,#e0e0e0);border-radius:8px;cursor:zoom-in;" title="点击放大" />' +
                '</div>';
        }

        // ── 情形 B：base64 模式（小区域截图或显式 returnFormat="base64"）──
        let b64 = raw.startsWith('data:image/') ? raw : (raw.startsWith('base64=') ? raw.slice(7) : (raw.startsWith('base64:') ? raw.slice(7) : (raw.indexOf('base64=') >= 0 ? raw.slice(raw.indexOf('base64=') + 7) : raw)));
        b64 = b64.trim();
        ChatLog.info?.('ORCH', `[screen-capture] base64-mode tool=${toolName} rawLen=${(raw||'').length} b64Len=${(b64||'').length} b64Head=${(b64||'').slice(0,20)} matches=${/^(data:image\/|iVBORw0KGgo|iVBOR)/.test(b64||'')}`);
        if (!b64 || !/^(data:image\/|iVBORw0KGgo|iVBOR)/.test(b64)) return null;
        const src = /^data:image\//.test(b64) ? b64 : 'data:image/png;base64,' + b64;
        const safeSrc = String(src).replace(/[<>"']/g, '');
        return '<div class="screen-capture-wrap">' +
            '<img class="screen-capture-img" src="' + safeSrc + '" alt="屏幕截图" ' +
            'onerror="this.outerHTML=\'&lt;div style=&quot;padding:10px;color:#666;font-size:13px;&quot;&gt;⚠️ 截屏图片加载失败（base64 异常/过长）&lt;/div&gt;\'" ' +
            'style="max-width:100%;max-height:360px;border:1px solid var(--border-color,#e0e0e0);border-radius:8px;cursor:zoom-in;" title="点击放大" />' +
            '</div>';
    }

    /**
     * 构建 excel_read_range 的可视化结果：元数据卡片 + 数据预览表。
     * 取代裸 JSON 输出，让"读取数据"类操作返回清晰的结构化信息（范围/行列/表头/预览行）。
     * @param {object} resultObj - 后端 excel_read_range 的返回对象
     * @returns {string|null} HTML 片段；无法解析返回 null 交回默认渲染
     */
    /**
     * 深解包 excel_read_range 的返回负载。
     * 【根因修复】后端常把读数据结果序列化成 JSON 字符串，包在 {result} / {Text} /
     * MCP content 里（McpTextContent.Text → McpManager 统一包成 { result: "<json字符串>" }）。
     * 旧代码只解"result 是 object"的分支，遇到字符串直接跳过 → hasMeta 判不中 → 落回裸 JSON，
     * 表现就是"读取数据有时候只显示 JSON、不显示元数据卡片"。
     * 现统一：字符串可解析成读数据 meta，或从常见承载键 result/data/text/Outputs 中取出对象/字符串再解。
     * @returns {object|null} 含读数据 meta 的对象；递归上限 depth 层，仍无则返回当前层（交由 hasMeta 兜底）
     */
    function unwrapReadPayload(input, depth) {
        if (input == null) return input;
        if (depth <= 0) return input;
        let cur = input;
        if (typeof cur === 'string') {
            const t = cur.trim();
            if (!t) return null;
            try { cur = JSON.parse(t); } catch (_) { return null; }
        }
        if (!cur || typeof cur !== 'object' || Array.isArray(cur)) return cur;
        // 已是读数据 meta（含表头/预览行/工作表）→ 直接用
        if ('headers' in cur || 'previewRows' in cur || 'worksheet' in cur) return cur;
        // 常见承载键：result/数据/文本/output/Outputs
        for (const k of ['result', 'data', 'text', 'Text', 'output', 'Outputs', 'Output']) {
            if (!(k in cur)) continue;
            const v = cur[k];
            if (v == null || typeof v === 'number' || typeof v === 'boolean') continue;
            const next = unwrapReadPayload(v, depth - 1);
            if (next && typeof next === 'object' && !Array.isArray(next)) {
                if ('headers' in next || 'previewRows' in next || 'worksheet' in next ||
                    'rows' in next || 'range' in next || 'rowCount' in next || 'columns' in next) {
                    return next;
                }
            }
        }
        // MCP content 数组：content:[{type:'text', text:"<json>"}]
        if (Array.isArray(cur.content)) {
            for (const item of cur.content) {
                if (!item || typeof item !== 'object') continue;
                const next = unwrapReadPayload(item.text || item.Text, depth - 1);
                if (next && typeof next === 'object' && !Array.isArray(next) &&
                    ('headers' in next || 'previewRows' in next || 'worksheet' in next ||
                     'rows' in next || 'range' in next || 'rowCount' in next || 'columns' in next)) {
                    return next;
                }
            }
        }
        return cur;
    }

    function buildReadRangeMetaCardHtml(resultObj) {
        if (resultObj == null) return null;
        const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g,
            c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
        // 兼容 direct object / { result: object|json字符串 } / { wrapper:{...} } / MCP content
        const data = unwrapReadPayload(resultObj, 3);

        const hasMeta = data && typeof data === 'object' &&
            ('worksheet' in data || 'range' in data || 'headers' in data || 'rows' in data || 'columns' in data || 'rowCount' in data || 'previewRows' in data);
        if (!hasMeta) return null;

        const worksheet = data.worksheet ? esc(data.worksheet) : '';
        const range = data.range ? esc(data.range) : '';
        const headers = Array.isArray(data.headers) ? data.headers : [];
        const rows = Number(data.rows) || (data.rowCount != null ? Number(data.rowCount) : 0);
        const cols = Number(data.columns) || (data.colCount != null ? Number(data.colCount) : headers.length) || 0;
        const preview = Array.isArray(data.previewRows) ? data.previewRows
            : (Array.isArray(data.rows) ? data.rows : null);

        let html = '<div class="read-range-meta-card" style="border:1px solid var(--border-color,#e0e0e0);border-radius:8px;overflow:hidden;margin:6px 0;font-family:\'Microsoft YaHei\',\'PingFang SC\',sans-serif;">';
        html += '<div style="background:var(--bg-secondary,#f7f8fa);padding:8px 12px;font-size:13px;color:#333;border-bottom:1px solid var(--border-color,#e0e0e0);">📊 已读取数据';
        if (worksheet || range || rows || cols) {
            const metaBits = [];
            if (worksheet) metaBits.push(`工作表 ${worksheet}`);
            if (range) metaBits.push(`范围 ${range}`);
            if (rows || cols) metaBits.push(`${rows} 行 × ${cols} 列`);
            if (metaBits.length) html += ` <span style="color:#888;font-size:12px;">（${metaBits.join(' · ')}）</span>`;
        }
        html += '</div>';
        if (headers.length) {
            html += '<div style="padding:5px 12px;font-size:12px;color:#555;background:#fbfcfd;">';
            html += '<span style="color:#999;">列名：</span>' + headers.map(esc).join(' · ');
            html += '</div>';
        }
        if (preview && preview.length) {
            const maxRows = Math.min(preview.length, 8);
            const keys = (typeof preview[0] === 'object' && !Array.isArray(preview[0])) ? Object.keys(preview[0]) : [];
            html += '<div style="overflow:auto;max-height:260px;">';
            html += '<table style="width:100%;border-collapse:collapse;font-size:12px;">';
            if (keys.length) {
                html += '<thead><tr>';
                for (const k of keys) html += `<th style="border:1px solid var(--border-color,#e0e0e0);padding:4px 8px;background:var(--bg-secondary,#f7f8fa);text-align:left;white-space:nowrap;font-weight:600;">${esc(k)}</th>`;
                html += '</tr></thead><tbody>';
                for (let r = 0; r < maxRows; r++) {
                    html += '<tr>';
                    for (const k of keys) {
                        const v = preview[r][k];
                        html += `<td style="border:1px solid var(--border-color,#e0e0e0);padding:4px 8px;white-space:nowrap;">${esc(v)}</td>`;
                    }
                    html += '</tr>';
                }
                html += '</tbody></table>';
            } else if (Array.isArray(preview[0])) {
                for (let r = 0; r < maxRows; r++) {
                    const isHdr = r === 0;
                    html += '<tr>';
                    for (const cell of preview[r]) {
                        const tag = isHdr ? 'th' : 'td';
                        html += `<${tag} style="border:1px solid var(--border-color,#e0e0e0);padding:4px 8px;${isHdr ? 'background:var(--bg-secondary,#f7f8fa);font-weight:600;' : ''}white-space:nowrap;">${esc(cell)}</${tag}>`;
                    }
                    html += '</tr>';
                }
                html += '</table>';
            }
            if (preview.length > maxRows) html += `<div style="padding:4px 12px;font-size:12px;color:#999;">… 还有 ${preview.length - maxRows} 行</div>`;
            html += '</div>';
        }
        html += '</div>';
        return html;
    }

    // 每个 active turn 的 AbortController：turnId -> AbortController
    const _activeControllers = new Map();

    // ══════════════════════════════════════════════════════════════════
    // 路由确认等待：LLM 判 text 但 regex 有 Skill 命中时，挂起 dispatch
    // 等待用户从两个候选（keep=走 Skill / drop=走 text-pipeline）中选一个，
    // 而非直接覆盖为 text-pipeline 截断 Skill 意图。
    // 超时 2min → 保守默认 keep（宁可多一次确认，不可错误截断）。
    // ══════════════════════════════════════════════════════════════════
    const _confirmResolvers = new Map();  // turnId → resolve({ choice: 'keep'|'drop'|'aborted' })
    const ROUTE_CONFIRM_TIMEOUT_MS = 120000;

    // LLM 意图分析超时（ms）：超时后回退关键词路由，防止慢速 LLM 一直堵住后面的工具。
    // 6000ms 仍会在本地模型吐出第一个字之前截断（18:52 整 6 秒，卡片里只有超时句）。
    // 12 秒内返回的，卡片写出意图/理由/路由；到点仍没有，保留已经写下的步骤并继续。
    const INTENT_ANALYSIS_TIMEOUT_MS = 12000;

    // 路由常量单一事实源：config/routing-constants.js
    // （COMMON_2CHAR_WORDS / detectLanguageDirective / detectCodeGenerationIntent /
    //   PYTHON_ANALYSIS_INTENT / detectPythonAnalysisIntent）
    // 运行时取值，避免加载顺序问题；缺失时降级为空对象。
    function routing() {
        return App.ChatV2.routingConstants || {};
    }

    // ============================================================
    // 通用后续消息继承（多轮消息继承）
    // 上一轮有执行上下文（_execSpec）时，本条消息若为"补充/修改/加一页"等修改诉求，
    // 路由回原管道，继承上一轮数据上下文真正重新执行，避免走通用 LLM 路径"光说没做"。
    // 设计参考：Dialogflow Contexts（输入上下文偏置意图分类，带 lifespan）
    //          + Follow-up Intents（后续意图显式建模）。
    // 关键词按管道细分，避免误触发。
    // ============================================================
    const FOLLOWUP_PATTERNS = {
        ppt: /补充|添加|增加|加一页|加个|加一|加页|章节|页面|插入|新增|修改|更新|重新|再来|继续|改成|改为|换成|换|调整|删除|去掉|目录|封面|结尾|要点|表格|图表|主题|配色|颜色/i,
        python: /继续|再次|再来|换成|改为|换个|维度|细分|对比|预测|趋势|异常/i,
        code: /加个|加一|改成|改为|换成|按钮|弹窗|宏|vba|代码|补充|修改|删除/i,
        skill: /图表|柱状|饼图|折线|改成|换成|按.*分组|按.*汇总|主题|美化|格式|样式|配色|颜色|字体|对齐|边框|列宽|冻结|筛选|清洗|去重|增删改查|数据清洗|格式化|图表|画图/i,
        doc: /补充|修改|标题|段落|章节|文档|报告/i
    };

    // 显式新任务标记（P6 生命周期）：命中即清空上一轮执行上下文，防止跨任务误继承。
    // 仅收录"明确换任务"的措辞，避免误伤"重新生成"这类可能仍是后续消息的诉求。
    const NEW_TASK_PATTERNS = /新任务|换个任务|换一个任务|另起一个|重新开始|从头开始|新做一个|新建一个|重新做一个|换一个话题/i;

    // PPT 跨管道继承：明确引用上一轮分析/报告/结果并要转成 PPT 的措辞。
    // 命中且上一轮是分析类管道（skill/text/python）时，继承分析结果规划 PPT，
    // 避免管道切换断层生成通用"数据分析报告"。
    // 跨管道 PPT 意图：引用上一轮分析/报告/结果 + 明确要 PPT（幻灯片/演示文稿）
    // 所有备选方案都必须含 PPT/幻灯片/演示文稿 关键词，避免 "刚才的结果" 在非 PPT 上下文被误命中
    const PPT_CROSS_PIPELINE_PATTERNS = /(?:上面|刚才|之前)(?:的)?(?:分析|报告|结果).*(?:PPT|ppt|幻灯片|演示文稿|做成ppt|转成ppt)|(?:把|将)(?:上面|刚才|之前)?(?:的)?(?:分析|报告|结果).*(?:转|做|做成|制作|生成|出).*(?:PPT|ppt|幻灯片|演示文稿)/i;

    /** 判定：必须有上一轮执行上下文 + 本条是修改诉求（双重门槛防误触发） */
    function detectFollowUp(message, prevExec) {
        if (!prevExec || !prevExec.spec) return null;
        const pattern = FOLLOWUP_PATTERNS[prevExec.spec.pipeline];
        return pattern && pattern.test(String(message || '')) ? prevExec : null;
    }

    /**
     * 根据 kind 选择 pipeline
     * 优先走 pipelineRegistry（支持降级链），
     * 注册表未加载时回退旧 switch 逻辑。
     */
    function pickPipeline(kind) {
        const registry = App.ChatV2.pipelineRegistry;
        if (registry && typeof registry.pick === 'function') {
            return registry.pick(kind);
        }
        // 旧逻辑兜底（注册表未加载时）
        const pipelines = App.ChatV2.pipelines || {};
        switch (kind) {
            case 'image':  return pipelines.image;
            case 'doc':    return pipelines.doc;
            case 'python': return pipelines.python;
            case 'html':   return pipelines.html;
            case 'skill':  return pipelines.skill;
            case 'code':   return pipelines.code;
            case 'text':
            default:       return pipelines.text;
        }
    }

    /**
     * 根据附件类型推断消息 kind
     * 【P1 修复】代码类附件确定性兜底：原逻辑把所有非图片附件硬归 doc，
     * 附件消息完全绕过 python/code/skill 路由（这些路由均要求 kind==='text'），
     * 意图层失败时"运行这个脚本"类执行诉求会掉进 doc 云分析而丢失确定性执行能力。
     * 扩展名是确定性信号，不依赖意图层/LLM，超时或失败时仍然生效。
     */
    function inferKind(message, attachments) {
        if (attachments && attachments.length > 0) {
            const first = attachments[0];
            // 图片附件
            if (first.type === 'image' || (first.mime || '').startsWith('image/')) return 'image';
            // 代码类附件：按扩展名确定性路由
            //   .py/.ipynb → python-pipeline（可读取并执行/分析脚本）
            //   .vba/.bas/.vbs/.cls/.frm → code-pipeline（代码生成管道）
            const fname = String(first.name || first.fileName || first.FileName || first.Name || '');
            const ext = ((fname.match(/\.([A-Za-z0-9]+)$/) || [])[1] || '').toLowerCase();
            if (ext === 'py' || ext === 'ipynb') return 'python';
            if (ext === 'vba' || ext === 'bas' || ext === 'vbs' || ext === 'cls' || ext === 'frm') return 'code';
            // 其他附件统一走 doc
            return 'doc';
        }
        return 'text';
    }

    /**
     * 根据 Skill 的 tags / outputType 判断其绑定的编程语言
     * @returns {string|null} 'python' | 'vba' | null
     */
    function getSkillLanguage(skill) {
        // 兼容后端 PascalCase 和前端 camelCase
        const tags = Array.isArray(skill.tags) ? skill.tags
            : Array.isArray(skill.Tags) ? skill.Tags
            : [];
        const outputType = skill.outputType || skill.OutputType || '';
        if (tags.includes('python') || outputType === 'python_script') return 'python';
        if (tags.includes('vba') || outputType === 'vba_code') return 'vba';
        return null;
    }

    /**
     * 【Phase2】通用 Skill 意图路由 — 命中时返回匹配的 Skill
     *
     * 设计原则：
     *   - 不硬编码任何具体 Skill 名称或业务逻辑。
     *   - 只通过 App.cachedSkills 中声明的 keywords / inputContext 做通用匹配。
     *   - 新增 Skill 只需在 Skills/*.md 中注册并声明 keywords/pipeline，无需改 orchestrator。
     *
     * Skill 可声明 pipeline 字段来指定路由到哪个 pipeline：
     *   - 不声明或 pipeline: "skill" → 走 skill-pipeline（VB后端执行）
     *   - pipeline: "html" → 走 html-display-pipeline（前端确定性执行）
     *   - pipeline: "python" → 走 python-pipeline
     *   - pipeline: "text" → 走 text-pipeline（agent-loop）
     */

    /**
     * 【降级守卫 A】判断一个 skill 对象是否为"当前真正可用"的 Skill。
     *
     * 唯一权威来源：App.cachedSkills（后端 GetAvailableSkills 已按 isEnabled 过滤）。
     * 判定规则（从强到弱）：
     *   1. 一个启用的 Skill 都没有（全部停用）→ 任何 Skill 都不可执行，含 DESCRIPTORS 空壳；
     *      这是"无 Skill 场景"的核心判据，也是唯一需要拦掉的场景。
     *   2. 合成空壳（__fromDescriptor）→ **不可用**。空壳由 shouldRouteToSkill / LLM 路由在
     *      cachedSkills 找不到真实 Skill 时注入，只有 name+keywords，**没有任何 steps /
     *      processingMode / executionMode**，进 skill-pipeline 后拿不到可执行内容，只能退化成
     *      "单次错误工具调用"——实测"读取当前选中数据，按部门汇总销售额，画柱状图，把汇总结果写回新表"：
     *      空壳 chart 只调 excel_create_chart(dataRange=$A$1:$I$101) 对原始明细建图，
     *      图例是一堆"企业客户/个人客户"原始行、完全没按部门聚合（用户反馈"生成了一个没用图"）；
     *      更早还曾被 ngram 兜底跨族译成 ppt_add_chart_slide，报"PPT 项目不存在"。
     *      置空后由【补救守卫 D/E】接管（图表 → _forceChart 注入"先 aggregate 再 create_chart"引导；
     *      看板 → _forceDashboard/_forceChart/_forceMultiStep），再落到 text-pipeline 多步骤管道，
     *      用真实工具链执行——这正是守卫 C 注释所写的"仅 DESCRIPTORS 空壳 → 丢弃"的语义。
     *   3. 真实 Skill → 必须能在 cachedSkills（启用清单）中找到同名项；找不到说明它已被停用 → 不可用。
     */
    function isRealAvailableSkill(s) {
        if (!s) return false;
        const cached = (App.cachedSkills || []).filter(x => x.isEnabled !== false && x.IsEnabled !== false);
        // 1) 全部停用 → 一律不可用
        if (cached.length === 0) return false;
        // 2) 合成空壳（无 steps）→ 不可用，交回守卫 D/E + text-pipeline 多步骤管道
        if (s.__fromDescriptor) return false;
        // 3) 真实 Skill → 必须在启用清单中
        const name = s.name || s.Name || '';
        if (!name) return false;
        return cached.some(x => (x.name || x.Name || '') === name);
    }

    /**
     * 【降级守卫 B】判断消息命中的确定性 Skill 域里，是否存在"真正可用"的 Skill。
     *
     * 用于 McpDispatcher 短路块：域命中但无可用 Skill 时不再拦截，放行通用管道。
     * 判定顺序：
     *   1. 完全没有可用 Skill（全部停用）→ false（明确回落）
     *   2. 命中确定性域 → 用域声明的 names 正则反查可用 Skill；有则 true，无则 false
     *   3. 未命中确定性域（如 SkillCreator 等路由策略意图）→ 退化为关键词匹配，
     *      仍无法判定时保守返回 true，维持"命中就拦截"的原有行为，避免误伤已启用 Skill。
     */
    function hasExecutableSkillForMessage(message) {
        try {
            const avail = (App.cachedSkills || []).filter(x => x.isEnabled !== false && x.IsEnabled !== false);
            if (avail.length === 0) return false;

            const _sd = App.ChatV2?.skillDomain;
            if (_sd && typeof _sd.matchDomains === 'function' && typeof _sd.getAll === 'function') {
                const domains = _sd.matchDomains(message || '') || [];
                const descs = _sd.getAll() || [];
                for (const did of domains) {
                    const d = descs.find(x => x.id === did);
                    if (d && d.names && avail.some(s => d.names.test(s.name || s.Name || ''))) return true;
                }
                if (domains.length > 0) return false; // 有明确域但该域无可用 Skill → 放行通用管道
            }

            for (const s of avail) {
                const kws = Array.isArray(s.keywords) ? s.keywords : (Array.isArray(s.Keywords) ? s.Keywords : []);
                for (const kw of kws) {
                    if (kw && (message || '').includes(kw)) return true;
                }
            }
            return true; // 无法判定 → 保守维持原行为
        } catch (e) {
            return true; // 守卫自身异常时保守放行，不影响已启用 Skill 的路由
        }
    }

    // 技能名归一化：忽略大小写与 _/- 分隔符。
    // LLM 常把技能名写成 snake_case（SkillCreator → skill_creator），
    // 用严格相等比较会让这类路由被判"目录里没有"而静默丢弃。
    function _normSkillName(s) {
        return String(s || '').toLowerCase().replace(/[\s_\-]/g, '');
    }

    /**
     * 校验 LLM 意图分析给的 route.skillId 能否真正落地成技能，并解析出会执行的 SkillSpec。
     * 与下方「LLM route 覆盖 regex」保持同一口径，供意图块如实显示"这条建议是否被采用"。
     *
     * 背景（实测 2026-10-03 08:59）：用户说"为火车票创建一个解析器"，LLM 返回
     *   route: { pipeline:'skill', skillId:'skill_creator', confidence:0.8 }
     * 但 skill_creator 只是 getSkillCatalog 里给 LLM 看的**展示别名**，既不在 DESCRIPTORS，
     * 也不是真实技能名 → 覆盖逻辑解析不到、保留 regex 结果（实际执行 GenerateOCRParseScript）。
     * 结果：执行是对的，界面上却写着 skill_creator，用户以为选错了技能。
     *
     * 判定顺序：
     *   ① DESCRIPTORS 领域 id（chart/dashboard/...）→ 用该域 names 反查真实 SkillSpec
     *   ② 展示别名 skill_creator / html_report → 必须过该技能自己的意图守卫（与 regex 路径同一把尺）
     *   ③ cachedSkills 真实技能名（大小写/下划线不敏感）
     * ②③ 的顺序不能反：_normSkillName('skill_creator') === _normSkillName('SkillCreator')，
     * 先做③的话别名会被"真实技能名"这一条吞掉，意图守卫形同虚设——
     * 实测后果就是"为火车票创建一个解析器"被 SkillCreator 抢走（这正是要防的那个 bug）。
     * @returns {{ok:boolean, skill:Object|null, reason:string}}
     */
    function resolveLlmRouteTarget(route, message) {
        const rid = String((route && route.skillId) || '').trim();
        if (!rid || rid === 'free_chat') return { ok: false, skill: null, reason: '无 skillId' };
        const sd = App.ChatV2 && App.ChatV2.skillDomain;
        const cached = Array.isArray(App.cachedSkills) ? App.cachedSkills : [];

        // ① 领域 id
        const descs = (sd && typeof sd.getAll === 'function') ? sd.getAll() : [];
        const found = descs.find(d => d.id === rid);
        if (found) {
            const real = cached.find(s => found.names.test(s.name || s.Name || ''));
            return { ok: true, skill: real || { name: found.id, __fromDescriptor: true }, reason: 'DESCRIPTORS 领域' };
        }

        // ② 目录展示别名：不是真实的"技能名/领域名"，必须过该技能自己的意图守卫才认
        if (_normSkillName(rid) === 'skillcreator') {
            const pol = sd && sd.ROUTING_POLICIES && sd.ROUTING_POLICIES.SkillCreator;
            const pass = !!(pol && typeof sd._testPolicyIntent === 'function' && sd._testPolicyIntent(pol, message || ''));
            if (pass) {
                return { ok: true, skill: cached.find(s => _normSkillName(s.name || s.Name) === 'skillcreator') || null, reason: 'SkillCreator 意图守卫通过' };
            }
            return {
                ok: false, skill: null,
                reason: 'skill_creator 要求消息里真的出现 skill/技能/插件/智能体 字样（本消息只有"创建"，是"创建解析器"）'
            };
        }
        if (_normSkillName(rid) === 'htmlreport') {
            return { ok: false, skill: null, reason: 'html_report 需命中 HTML 报表意图（生成动词 + html/网页关键词）' };
        }

        // ③ 真实技能名（归一化：大小写与 _/- 不敏感）
        const byName = cached.find(s => _normSkillName(s.name || s.Name) === _normSkillName(rid));
        if (byName) return { ok: true, skill: byName, reason: 'cachedSkills 技能名' };

        return { ok: false, skill: null, reason: '目录里没有该 id' };
    }

    function shouldRouteToSkill(message, context) {
        if (!message) return null;
        const lower = message.toLowerCase();
        let skills = App.cachedSkills || [];

        // 兜底过滤：跳过显式停用的 skill（isEnabled === false）
        // 正常情况下 GetAvailableSkills 已过滤，但 synthetic skill / 兜底路径可能漏掉
        skills = skills.filter(s => s.isEnabled !== false && s.IsEnabled !== false);

        // ════════════════════════════════════════════════════════════════
        // DESCRIPTORS 注入：把前端声明了但后端 GetAllSkills 没返回的 skill 也加进来。
        // 之前 candidates 只遍历 cachedSkills，若后端漏发 DashboardGenerator 则 DESCRIPTORS 白写。
        // 现在 DESCRIPTORS 是路由的完整 skill 清单——后端有就用后端的，后端没有就用前端声明的构造 synthetic skill。
        // ════════════════════════════════════════════════════════════════
        {
            const _sd = App.ChatV2?.skillDomain;
            if (_sd && typeof _sd.getAll === 'function') {
                const _descs = _sd.getAll();
                for (const _d of _descs) {
                    const _exists = skills.some(s => {
                        const _n = s.name || s.Name || '';
                        return _d.names.test(_n);
                    });
                    if (!_exists) {
                        // 【停用感知】"后端没返回"有两种截然不同的原因，此前被混为一谈：
                        //   ① 后端真的漏发（DESCRIPTORS 才有存在意义）；
                        //   ② 用户刚在面板里把它停用（GetAvailableSkills 过滤掉了）。
                        // ② 若也走注入，停用就被抵消——实测：停用 SmartTableBeautifier 后
                        // cachedSkills 里它消失，这个分支立刻补回一个空壳 beautifier(kw=9)，
                        // 继续参与关键词打分与 LLM 目录，用户看到"我把 skill 关了，怎么还在用"。
                        if (typeof _sd.isDomainDisabledByUser === 'function' && _sd.isDomainDisabledByUser(_d)) {
                            ChatLog.info?.('ORCH', `[shouldRouteToSkill] DESCRIPTORS 跳过已被用户停用的领域: ${_d.id}`);
                            _orchHostLog('INFO', `[shouldRouteToSkill] DESCRIPTORS 跳过已停用领域: ${_d.id}（用户已关闭，不再注入空壳复活）`);
                            continue;
                        }
                        // 【背书包袱】第三种"后端没返回"：后端压根没有与该领域对应的 Skill。
                        // 实测（2026-10-01 10:40）：本环境从未安装 formatter skill，DESCRIPTORS 里的
                        // formatter 被反复补成空壳（kw=32 → score=5），而用户已启用的真 Skill
                        // SmartTableBeautifier 只拿到 score=1 → best=formatter。下方 strongMatch 里
                        // "唯一候选+强意图2字词"这条本可放行 beautifier，却因空壳占了第一、候选变两个而失效，
                        // 最终 beautifier 被 skill-pipeline 拒之门外（用户看到"我明明打开了 beautifier，
                        // 却命中 formatter"，然后 generic 工具兜着做）。无真实技能背书的领域不注入空壳，
                        // 把路由让给真正存在的 Skill。判据未就绪时 isDomainBacked 返回 true（保守保留）。
                        if (typeof _sd.isDomainBacked === 'function' && !_sd.isDomainBacked(_d)) {
                            ChatLog.info?.('ORCH', `[shouldRouteToSkill] DESCRIPTORS 跳过无真实技能背书的领域: ${_d.id}`);
                            _orchHostLog('INFO', `[shouldRouteToSkill] DESCRIPTORS 跳过无背书的幽灵领域: ${_d.id}（后端无对应 Skill，不再注入空壳抢路由）`);
                            continue;
                        }
                        const _kw = _d.keywords || [];
                        skills = skills.concat({
                            name: _d.id,
                            keywords: _kw,
                            displayName: _d.id,
                            __fromDescriptor: true
                        });
                        ChatLog.info?.('ORCH', `[shouldRouteToSkill] DESCRIPTORS 注入 synthetic skill: ${_d.id} (keywords=${_kw.length})`);
                    }
                }
            }
        }

        if (skills.length === 0) {
            ChatLog.info?.('ORCH', `[shouldRouteToSkill] cachedSkills 为空！`);
            _orchHostLog('WARN', `[shouldRouteToSkill] 可用 Skill 清单为空 → 任何 Skill 都无法路由（全部已停用，或 Skill 未加载）`);
            return null;
        }
        // ✨ 关键 debug：打印所有 cachedSkills 的名称+keywords 数量，确认 Beautifier 是否在里面
        {
            const skillList = skills.map(s => {
                const name = s.name || s.Name || '?';
                const kwLen = Array.isArray(s.keywords) ? s.keywords.length : (Array.isArray(s.Keywords) ? s.Keywords.length : 0);
                const hasBeautifierKw = name.includes('Beautifier') || name.includes('beautifier') || name.includes('美化');
                return `${name}(kw=${kwLen}${hasBeautifierKw ? ' ⚠️' : ''})`;
            }).join(', ');
            ChatLog.info?.('ORCH', `[shouldRouteToSkill] cachedSkills(${skills.length}): ${skillList}`);
            _orchHostLog('INFO', `[shouldRouteToSkill] cachedSkills(${skills.length}): ${skillList}`);
        }

        // 记录 SkillCreator 是否在列表中
        const skillCreatorSkill = skills.find(s => (s.name || s.Name) === 'SkillCreator');
        if (skillCreatorSkill) {
            const scKeywords = Array.isArray(skillCreatorSkill.keywords) ? skillCreatorSkill.keywords : [];
            const matchedKw = scKeywords.filter(kw => lower.includes(kw.toLowerCase()));
            ChatLog.debug?.('ORCH', `[shouldRouteToSkill] SkillCreator 关键词匹配: ${matchedKw.join(', ')} (共${scKeywords.length}个关键词)`);
        } else {
            ChatLog.debug?.('ORCH', `[shouldRouteToSkill] SkillCreator 不在 cachedSkills 列表中 (total skills: ${skills.length})`);
        }

        // === 语言指令检测：用户明确指定了编程语言 ===
        const langDirective = routing().detectLanguageDirective(message);
        const codeGenIntent = routing().detectCodeGenerationIntent(message);

        const candidates = [];
        for (const skill of skills) {
            let keywords = Array.isArray(skill.keywords) ? skill.keywords
                : Array.isArray(skill.Keywords) ? skill.Keywords
                : [];
            if (keywords.length === 0) {
                const _sd = App.ChatV2?.skillDomain;
                if (_sd && typeof _sd.findDescriptorBySkill === 'function') {
                    const _desc = _sd.findDescriptorBySkill(skill);
                    if (_desc && Array.isArray(_desc.keywords)) {
                        keywords = _desc.keywords;
                    }
                }
            }
            if (keywords.length === 0) continue;

            let score = 0;
            let hasLongKwMatch = false;
            let hasMidKwMatch = false;  // 4字词命中 = 强意图信号
            let hasShortKwMatch = false; // 3字词命中（具体图表名/操作名）= 强意图信号
            let hasStrong2Char = false; // 单2字强意图词命中（如"去重"）
            let lastStrongPos = -1; // 最后一个强意图2字词在消息中的位置（用于判定主动作）
            for (const kw of keywords) {
                if (typeof kw !== 'string') continue;
                const kwLower = kw.toLowerCase();
                const kwLen = kwLower.length;
                if (kwLen < ROUTE_T.KW_MIN_LEN) continue; // 跳过单字关键词，避免"技能"里的"技"等误匹配
                if (lower.includes(kwLower)) {
                    if (kwLen >= ROUTE_T.KW_LONG_LEN) {
                        score += ROUTE_T.SCORE_LONG_KW;
                        hasLongKwMatch = true;
                    } else if (kwLen >= ROUTE_T.KW_MID_LEN) {
                        score += ROUTE_T.SCORE_MID_KW;
                        hasMidKwMatch = true;
                    } else if (kwLen >= ROUTE_T.KW_SHORT_LEN) {
                        score += ROUTE_T.SCORE_SHORT_KW;
                        hasShortKwMatch = true;
                    } else {
                        score += ROUTE_T.SCORE_WEAK_2CHAR; // 2字词弱权重，避免通用词（"技能"/"创建"）误匹配
                    }
                }
            }
            if (score === 0) continue;

            // [方案B] 2字词通用词（skill/技能/创建）必须有一个3+字短语共同命中，
            // 否则得分视为无效，交给 text-pipeline 的 LLM 决策
            // 例外：单 2 字"强意图词"（如"去重/美化/排序"）也应放行——
            // 否则"去重"这类请求会掉进 text-pipeline 让小模型幻觉执行。
            // 护栏只在"没有任何 3+ 字词命中"时收紧 kwSumLen；
            //   3+ 字词命中（hasShortKwMatch/hasMidKwMatch/hasLongKwMatch）本身就是强意图，
            //   不应再被 kwSumLen<4 误杀（如"柱状图"3字词：画个柱状图 → 应路由 SmartChartWizard）。
            if (!hasLongKwMatch && !hasMidKwMatch && !hasShortKwMatch) {
                let kwSumLen = 0;
                let hitCount = 0;
                for (const kw of keywords) {
                    if (typeof kw !== 'string') continue;
                    const kwLower = kw.toLowerCase();
                    if (kwLower.length >= ROUTE_T.KW_MIN_LEN && lower.includes(kwLower)) {
                        kwSumLen += kwLower.length;
                        hitCount++;
                        // 强意图 2 字词：不是通用虚词（技能/创建/生成/处理/分析/数据/文件/操作 等）
                        if (kwLower.length === 2 && !routing().COMMON_2CHAR_WORDS.has(kwLower)) {
                            hasStrong2Char = true;
                            const idx = lower.lastIndexOf(kwLower);
                            if (idx > lastStrongPos) lastStrongPos = idx;
                        }
                    }
                }
                // 原有护栏：2字词合计长度不足且无强意图词 → 剔除
                if (!hasStrong2Char && (kwSumLen < ROUTE_T.KW_SUM_LEN_MIN || hitCount === 0)) continue;
            }

            // === 语言指令加分：用户明确说了 Python/VBA，匹配 Skill 语言 ===
            if (langDirective) {
                const skillLang = getSkillLanguage(skill);
                if (skillLang === langDirective) {
                    // 强匹配：用户明确指定了语言，且 Skill 用该语言
                    score += ROUTE_T.LANG_MATCH_BONUS;
                } else if (skillLang && skillLang !== langDirective) {
                    // 弱惩罚：Skill 用的是不同语言
                    score -= ROUTE_T.LANG_MISMATCH_PENALTY;
                }
            }

            // === 代码生成意图：区分代码生成类 Skill 和执行类 Skill ===
            if (codeGenIntent) {
                const skillPipeline = skill.pipeline || skill.Pipeline || '';
                const skillName = skill.name || skill.Name || '';
                const isCodeGenSkill = (skillPipeline === 'code' || skillPipeline === 'python')
                    || skillName.startsWith('general-');
                if (isCodeGenSkill) {
                    score += ROUTE_T.CODEGEN_MATCH_BONUS;
                } else {
                    // 执行类 Skill（如 excel-split-sheets）在"写代码"意图下重惩罚
                    score -= ROUTE_T.CODEGEN_EXEC_PENALTY;
                }
            }

            // 检查所需上下文是否满足
            const requiredCtx = Array.isArray(skill.inputContext) ? skill.inputContext
                : Array.isArray(skill.InputContext) ? skill.InputContext
                : [];
            let ctxOk = true;
            // SkillSpec / executionMode=spec 的蓝图引擎能自动从活动表已用区域兜底读取，
            // selection 仅是"更精准"的配置项而非硬依赖（DashboardGenerator/configFields.selection
            // 被后端自动映射成 inputContext=['selection'] 后，无选区时会被这里误过滤——
            // 直接导致 shouldRouteToSkill 返回 null → HtmlBrowserDisplay 兜底乱入的 bug）。
            const _isSpecMode = (skill.executionMode || skill.ExecutionMode || skill.processingMode || '') === 'spec';
            for (const req of requiredCtx) {
                if (req === 'selection') {
                    if (_isSpecMode) continue;
                    if (!(context?.selection?.range || context?.selection?.address)) { ctxOk = false; break; }
                } else if (req === 'referenceRegions') {
                    if (!Array.isArray(context?.referenceRegions) || context.referenceRegions.length === 0) { ctxOk = false; break; }
                } else if (req === 'attachments') {
                    if (!Array.isArray(context?.attachments) || context.attachments.length === 0) { ctxOk = false; break; }
                }
            }
            if (!ctxOk) {
                const sn = skill.name || skill.Name || '?';
                const reqCtx = requiredCtx.join(',');
                ChatLog.info?.('ORCH', `[shouldRouteToSkill] Skill ${sn} 被 ctxOk 过滤（inputContext=[${reqCtx}] 但 selection=${!!context?.selection?.range} refRegions=${(context?.referenceRegions?.length || 0)}）`);
                _orchHostLog('INFO', `[shouldRouteToSkill] Skill ${sn} 被 ctxOk 过滤（inputContext=[${reqCtx}] selection=${!!(context?.selection?.range || context?.selection?.address)} refRegions=${(context?.referenceRegions?.length || 0)}）`);
                continue;
            }

            // _kws：本命中的关键词表，供意图守卫判定"句首查询词是否其实是本技能的动作词"
            candidates.push({ skill, score, _hasLongMatch: hasLongKwMatch, _hasMidKwMatch: hasMidKwMatch, _hasShortKwMatch: hasShortKwMatch, _hasStrong2Char: hasStrong2Char, _lastStrongPos: lastStrongPos, _kws: keywords });
        }

        if (candidates.length === 0) return null;

        // 消息中最后一个强意图 2 字词通常是主动作（如"对合并后的数据做清洗"→清洗）。
        // 给命中该词的候选 +2 分，使主动作 Skill 得分>=3，触发下方强匹配路由，
        // 避免"清洗/合并"等多候选平局时掉进 text-pipeline 让小模型幻觉执行。
        let maxStrongPos = -1;
        for (const c of candidates) {
            if (c._lastStrongPos > maxStrongPos) maxStrongPos = c._lastStrongPos;
        }
        if (maxStrongPos >= 0) {
            for (const c of candidates) {
                if (c._lastStrongPos === maxStrongPos) c.score += ROUTE_T.STRONG2_POS_BONUS;
            }
        }

        // 按匹配度排序，取最高分
        // 同分时，命中强意图 2 字词（饼图/去重/清洗/合并等明确动作词）的候选优先：
        // 实测"使用饼图按部门展示销售额"——SmartChartWizard 命中"饼图"(2字强词→+2加成后 score=3)，
        // GeneratePython 命中"销售额"(3字数据描述词 score=3) 同分。若按 cachedSkills 插入顺序
        // GeneratePython 排前，strongMatch 判定只认 best 的 strong2 → false → 图表请求掉进
        // text-pipeline 让小模型幻觉（整表画图、缺参数）。"饼图"是明确的图表类型意图，
        // "销售额"只是数据列名，强意图词应优先于泛数据词。
        candidates.sort((a, b) => {
            const d = b.score - a.score;
            if (d !== 0) return d;
            const aStrong = (a._hasStrong2Char && a.score >= ROUTE_T.STRONG2_SORT_SCORE) ? 1 : 0;
            const bStrong = (b._hasStrong2Char && b.score >= ROUTE_T.STRONG2_SORT_SCORE) ? 1 : 0;
            return bStrong - aStrong;
        });
        let best = candidates[0];

        // [方案B] 大幅提升阈值：避免通用2字词（技能/创建/skill）凑分误匹配。
        // 短关键词命中只作为弱信号，真正的决策交给 text-pipeline 里的 LLM。
        // 只有强匹配（有长关键词命中 或 总分 >= 8）才走 Skill pipeline。
        // 放宽点：
        //   1. 4字词命中（hasMidKwMatch）也是强意图信号，如"美化表格"得4分即可路由；
        //   2. 唯一候选且命中强意图 2 字词（如"去重"）也路由；
        //   3. 3字词命中（hasShortKwMatch，如"柱状图/折线图/去重"）且总分>=3 也路由——
        //      具体图表名/操作名是明确意图，不该掉进 text-pipeline 让小模型幻觉。
        // 路由策略加成（数据驱动，取代原 SkillCreator 硬编码特判）：声明了 boostMinScore 的 Skill
        // 命中且 score>=阈值即进 strongMatch（原 hasSkillCreationIntent && score>=3）。
        const _routeDomain = App.ChatV2?.skillDomain || null;
        const hasIntentBoost = !!(_routeDomain && candidates.some(c => {
            const p = _routeDomain.getRoutingPolicy(c.skill.name || c.skill.Name);
            return !!(p && p.boostMinScore != null && c.score >= p.boostMinScore);
        }));

        const strongMatch = (best.score >= ROUTE_T.ROUTE_SCORE_LONG)
            || (best.score >= ROUTE_T.ROUTE_SCORE_LONG2 && candidates[0]._hasLongMatch)
            || (best.score >= ROUTE_T.ROUTE_SCORE_MID && candidates[0]._hasMidKwMatch)
            // 3字词命中（如"重复行"）且总分>=3：唯一候选直接路由；
            //   多候选时若最佳候选明显领先（分差>=3）也路由——否则"对当前数据继续清洗，去重并只把
            //   重复行黄色高亮标记"因 SmartDataAnalyzer 的"标记"成为第二候选，破坏单候选条件而掉进
            //   text-pipeline（LLM 光说不做，报"实际没有调用任何工具"）。best 已是最高分且明显领先=明确意图。
            || (best.score >= ROUTE_T.ROUTE_SCORE_SHORT && candidates[0]._hasShortKwMatch && (candidates.length === 1 || best.score - candidates[1].score >= ROUTE_T.ROUTE_LEAD_SCORE))
            || (candidates.length === 1 && candidates[0]._hasStrong2Char)
            // 最佳候选命中强意图 2 字词（清洗/去重/合并/拆分等）且总分>=3（命中多个关键词）时，
            //   即使存在多个候选也路由——避免"对合并后的数据做清洗"因 DataMerger 也命中"合并"而掉进
            //   text-pipeline 走后端占位消息（"共1步"+原生JSON泄露）。best 已是最高分，强词+多命中=明确意图。
            || (candidates.length > 0 && candidates[0]._hasStrong2Char && candidates[0].score >= ROUTE_T.ROUTE_SCORE_MIN)
            // 路由策略加成放行（数据驱动，取代原 SkillCreator 特判）
            || hasIntentBoost;

        // 记录所有候选 Skill 的得分（总是打印，不限关键词——定位"为什么某个请求没命中 skill"）
        {
            const candidateSummary = candidates.slice(0, 8).map(c => {
                const name = c.skill.name || c.skill.Name || 'unknown';
                return `${name}(score=${c.score},long=${c._hasLongMatch},mid=${c._hasMidKwMatch},short=${c._hasShortKwMatch},strong2=${c._hasStrong2Char})`;
            }).join('; ');
            ChatLog.info?.('ORCH', `[shouldRouteToSkill] Skill候选(${candidates.length}): ${candidateSummary || 'none'}, strongMatch=${strongMatch}, best=${best.skill.name || best.skill.Name}, bestScore=${best.score}`);
            _orchHostLog('INFO', `[shouldRouteToSkill] 候选(${candidates.length}): ${candidateSummary || 'none'} | strongMatch=${strongMatch} best=${best.skill.name || best.skill.Name} bestScore=${best.score}`);
        }

        if (!strongMatch) {
            ChatLog.info?.('ORCH', `[shouldRouteToSkill] 最佳匹配 ${best.skill.name || best.skill.Name} 得分 ${best.score}，不够强(strongMatch=${strongMatch})，交给 LLM`);
            _orchHostLog('INFO', `[shouldRouteToSkill] ✗ 未达强匹配阈值：best=${best.skill.name || best.skill.Name} score=${best.score}（阈值>=8 或 >=5+5字词 或 >=4+4字词 或 >=3+3字词 或 唯一候选+2字强词）→ 交给 LLM，本 Skill 不会执行`);
            return null;
        }

        // ── per-skill 路由策略 + 通用意图门 ──
        // 声明见 config/skill-domain.js：
        //   - 所有 skill 都过通用意图门（查询拦截 + 生成动词检测）
        //   - 显式 intent（如 SkillCreator）用双命中规则（动词+skill词）
        //   - block : 命中阻却正则即放行给 text-pipeline（如 HtmlBrowserDisplay/SmartFormulaAssistant）
        if (_routeDomain && _routeDomain.getRoutingPolicy) {
            // 意图守卫：best 不过守卫时**回退次优候选**（否则"做个excel看板"被 AutoGenerateWebReport(纯展示)
            // 抢到最高分被拦下后直接 return null → text-pipeline 幻觉"已生成"，DashboardGenerator 永远接不住）。
            let picked = null;
            for (const cand of candidates) {
                const _cname = cand.skill.name || cand.skill.Name;
                // 强关键词豁免（2026-10-09 放宽到中词与强 2 字词）：
                // 候选命中具体关键词时不再要求通用门的生成动词或 DESCRIPTORS 域关键词——
                //   长词(≥5字)"每1000行"/"按部门拆分" / 中词(4字)"跨表匹配" / 强2字"总结""对比"
                // 关键词本身就是比通用动词更硬的意图证据。
                // 三重安全边界（均在 _testPolicyIntent 内部生效，此处不会被绕过）：
                //   ① 查询拦截仍生效——且"句首查询词 = 本技能动作词"（如 DataDiffAnalyzer 的
                //      "对比"）才放行，"对比和合并有什么区别"因句中疑问语气仍被拦；
                //   ② blockedWords 仍生效；
                //   ③ 豁免仅对**非疑问句**生效（吗/什么/区别/？…）。
                const _strongKw = cand._hasLongMatch === true || cand._hasMidKwMatch === true || cand._hasStrong2Char === true;
                if (_strongKw) {
                    ChatLog.info?.('ORCH', `[shouldRouteToSkill] 候选 ${_cname} 命中强关键词(long=${cand._hasLongMatch}/mid=${cand._hasMidKwMatch}/strong2=${cand._hasStrong2Char})，豁免通用门动词/域要求（查询拦截与疑问句兜底仍生效）`);
                }
                if (!_routeDomain.matchesIntent(_cname, message, { strongKeyword: _strongKw, skillKeywords: cand._kws })) {
                    ChatLog.info?.('ORCH', `[shouldRouteToSkill] 候选 ${_cname} 未通过意图守卫，跳过`);
                    continue;
                }
                const _policy = _routeDomain.getRoutingPolicy(_cname) || {};
                if (_policy.block && new RegExp(_policy.block).test(message.toLowerCase())) {
                    ChatLog.info?.('ORCH', `[shouldRouteToSkill] 候选 ${_cname} 命中阻却规则，跳过`);
                    continue;
                }
                picked = cand;
                break;
            }
            if (!picked) {
                ChatLog.info?.('ORCH', `[shouldRouteToSkill] 全部候选均未通过意图守卫，交给 LLM`);
                _orchHostLog('INFO', `[shouldRouteToSkill] ✗ 全部候选未通过意图守卫/阻却规则 → 交给 LLM`);
                return null;
            }
            if (picked !== best) {
                ChatLog.info?.('ORCH', `[shouldRouteToSkill] 意图守卫回退: ${best.skill.name || best.skill.Name} → ${picked.skill.name || picked.skill.Name}`);
            }
            best = picked;
        }

        ChatLog.info?.('ORCH', `[shouldRouteToSkill] 命中 Skill 意图: ${best.skill.name} (score=${best.score}, lang=${langDirective || 'auto'}, codeGen=${codeGenIntent})`);
        _orchHostLog('INFO', `[shouldRouteToSkill] ✓ 命中 Skill：${best.skill.name} score=${best.score}（将由 skill-pipeline 执行）`);
        return best.skill;
    }

    const orchestrator = {
        /**
         * 派发一次 turn（主入口）
         * @param {Object} turnInput - { message, kind?, attachments?, context? }
         * @returns {Promise<string>} turnId
         */
        async dispatch(turnInput) {
            const sessionStore = App.ChatV2.sessionStore;
            if (!sessionStore) {
                ChatLog.error?.('ORCH', '[orchestrator] sessionStore not loaded');
                return null;
            }

            const { message, attachments } = turnInput || {};
            let context = (turnInput || {}).context;
            let kind = turnInput?.kind || inferKind(message, attachments);
            const dispatchStack = new Error().stack.split('\n').slice(2, 5).map(l => l.trim());
            ChatLog.info?.('ORCH', `[dispatch] 进入: msg="${(message||'').slice(0,40)}" kind=${kind} context=${JSON.stringify(context||{}).slice(0,120)} caller=${dispatchStack.join(' ← ')}`);

            // PPT 意图检测（单一处定义，供 Skill 强制 python 兜底与最终兜底共用）。
            // 含"趋势/图表"等分析词但明确是 PPT 诉求（如"在PPT汇报中增加趋势分析页"）
            // 时，PPT 路由优先于 python 分析兜底，避免误触发 Excel/Python 分析。
            const pptIntent = /ppt|powerpoint|幻灯片|演示文稿|做ppt|做成ppt|生成ppt|汇报/i.test(message || '');

            // PPT 结构诉求检测：仅当用户明确描述"页面结构"（多页/具体页型/分节等）时才判定为
            // 自由创作诉求——放行 agent-loop（AI 自由调用 ppt_ MCP 工具逐页自由创作），不锁定
            // GeneratePPTReport 技能。注意："风格/主题/配色/版式/模板"是主题修饰词而非页面结构，
            // 不在此列：用户说"用商务风格生成PPT"应走 GeneratePPTReport（skill 支持 theme 参数自动
            // 应用商务蓝主题），若误判为自由创作会被 agent-loop 的弱模型带偏（只聚合不生成→假成功）。
            // 技能仅作为"简单生成（无明确结构 + 有选区数据）"的扩展兜底，符合 SKILL=SOP、MCP=工具基石。
            const pptFreeBuild = pptIntent && /(?:第?\d+\s*[页张]|多页|分页|封面|目录|章节|结尾|封底|要点|页面|图表页|表格页|KPI|对比|时间线|里程碑|流程图|布局|分节|组织结构)/i.test(message || '');

            // 确保 Skills 已加载（preloadAllSkills 是非阻塞的后台任务，首条用户消息时可能尚未完成）
            // 若 cachedSkills 为空数组，主动加载一次，避免 shouldRouteToSkill 因 skills.length===0 返回 null
            if (kind === 'text' && typeof App.loadSkillsForLLM === 'function' &&
                (!Array.isArray(App.cachedSkills) || App.cachedSkills.length === 0)) {
                try {
                    await App.loadSkillsForLLM();
                } catch (e) {
                    ChatLog.warn?.('ORCH', '[orchestrator] Skills 按需加载失败，继续无 Skill 路由:', e);
                }
            }

            // === 立即启动 turn（用户气泡 + 思考态马上出现）===
            // 之前 startTurn 在路由末尾才调用，而路由中 LLM 意图分析（analyzeNewIntent）
            // 是阻塞的完整 LLM 往返（本地模型可达数秒），导致用户按下发送后长时间看不到
            // 任何反馈（"消息发出去太慢了"）。此处提前启动 turn，让渲染器立即显示用户气泡
            // 与思考态；路由（含慢速意图分析）在气泡可见的情况下进行。
            // 注意：kind 此时为初始推断值，路由完成后会同步最终 kind 到 turn（见下方）。
            if (_activeControllers.size > 0) {
                ChatLog.info?.('ORCH', '[orchestrator] 已有进行中的 turn，先取消');
                this.cancelCurrent('replaced_by_new_turn');
            }
            const turn = sessionStore.startTurn({
                message: message || '',
                kind,
                attachments: attachments || [],
                context: context || {}
            });
            const controller = new AbortController();
            _activeControllers.set(turn.id, controller);
            try {
                if (App.ChatV2.turnGoal && typeof App.ChatV2.turnGoal.build === 'function') {
                    turn.goal = App.ChatV2.turnGoal.build(message || '');
                    if (sessionStore.updateTurn) sessionStore.updateTurn(turn.id, { goal: turn.goal });
                }
            } catch (_goalErr) { /* 目标记录失败不影响派发 */ }

            // 路由 + 管道派发整体保护：turn 已提前启动，若路由阶段抛异常，
            // 必须 failTurn + 清理 controller，避免 turn 卡在 working 且 controller 泄漏。
            try {
            // === 生命周期（Phase 4）：明确新任务时清空执行上下文 ===
            // 用户显式表达"新任务/换个任务/重新开始"等意图时，清空上一轮 _execSpec，
            // 防止后续消息继承误路由到旧管道（P6 生命周期）。清空后下方 detectFollowUp
            // 因 getLastExecSpec() 返回 null 而自然不命中，本条按新任务路由。
            if (NEW_TASK_PATTERNS.test(message || '')) {
                if (sessionStore.clearExecSpec) {
                    sessionStore.clearExecSpec();
                    ChatLog.info?.('ORCH', '[orchestrator] 明确新任务: 已清空上一轮执行上下文');
                }
            }

            // 读取上一轮执行上下文（供后续消息继承判断）
            const prevExec = sessionStore.getLastExecSpec ? sessionStore.getLastExecSpec() : null;

            // ============================================================
            // 后续消息继承（优先于 Skill 匹配）
            // 有上一轮任务结构（_execSpec）时，先判断本条是否引用上一轮结果：
            //   1. PPT 跨管道：明确引用上一轮分析并要转 PPT → code-pipeline
            //   2. 图表跨管道：图表意图 + 无新选区 + 上一轮分析 → skill-pipeline
            //   3. 通用后续消息：修改/继续诉求 → 路由回原管道
            // 命中则继承任务结构规划下一个任务，避免"停留在关键词匹配"、
            // 无法结合上下文规划任务。否则才做 Skill 关键词匹配。
            // ============================================================
            let inheritedFollowUp = false;

            // 1. PPT 跨管道继承（Phase 5 扩展）
            // Bug1/Bug2 修复后 ppt_generate/ppt_project_start 等原子工具通过 CallLocalTool 真实可用，
            // PPT 跨管道不再强制 GeneratePPTReport skill，而是放行 dispatcher（直接命中 ppt_generate → 执行）
            // 或 agent-loop（LLM 自由调 ppt_ 工具）。
            // 不设置 inheritedFollowUp/context._execFollowUp —— 否则 _skillFollowUpPending 会阻塞 dispatcher。
            // 上一轮分析结果已自然存在于 agent-loop 的 messages 历史中，无需显式继承标志。
            if (pptIntent && PPT_CROSS_PIPELINE_PATTERNS.test(message || '')) {
                const prevPipeline = prevExec?.spec?.pipeline;
                // 放宽触发条件：不再强制要求 prevExec 存在，用户直接说"做成PPT"也能标记
                // （无 prevExec 时默认 kind='text' 已由 inferKind 处理，此处主要是设 _pptCrossPipeline 标记）
                if (!prevExec || prevPipeline === 'skill' || prevPipeline === 'text' || prevPipeline === 'python') {
                    kind = 'text';
                    context = context || {};
                    context._pptCrossPipeline = true;
                    ChatLog.info?.('ORCH', `[orchestrator] PPT 跨管道: ${prevPipeline ? '上一轮 ' + prevPipeline : '无前置对话'} → kind=text，放行 dispatcher/agent-loop 自由调 ppt_ 工具`);
                }
            }

            // 2. 图表跨管道继承（Phase 5 扩展）
            if (!inheritedFollowUp) {
                const chartIntent = /图表|画图|做图|可视化|柱状图|折线图|饼图|组合图|漏斗图|瀑布图|箱线图|树状图|雷达图|生成图表|做图表|画图表|创建图表|同比|导出图片|图表推荐/i.test(message || '');
                if (chartIntent) {
                    const prevPipeline = prevExec?.spec?.pipeline;
                    const ctxHasData =
                        (Array.isArray(context?.selection?.headers) && context.selection.headers.length > 0) ||
                        (Array.isArray(context?.referenceRegions) &&
                            context.referenceRegions.some(r => Array.isArray(r.headers) && r.headers.length > 0));
                    // 明确"仅切图表类型"（换成柱状图/改成折线图）必须继承上一轮图表数据上下文，
                    // 即使当前选区仍带着数据（ctxHasData）——用户诉求是"沿用当前这张图的数据只换类型"，
                    // 若被 ctxHasData 拦下会回退到实时选区，画出错误内容（实测第二条消息内容不对）。
                    const rc = App.ChatV2.routingConstants;
                    const chartTypeSwitch = !!(rc && rc.CHART_TYPE_SWITCH_RE && rc.CHART_TYPE_SWITCH_RE.test(message || ''));
                    const shouldInherit = chartTypeSwitch || !ctxHasData;
                    if (prevExec && shouldInherit &&
                        (prevPipeline === 'skill' || prevPipeline === 'text' || prevPipeline === 'python')) {
                        const chartSkill = (App.cachedSkills || []).find(
                            s => s.name === 'SmartChartWizard' || s.Name === 'SmartChartWizard');
                        if (chartSkill) {
                            context = context || {};
                            context._matchedSkill = chartSkill;
                            kind = 'skill';
                            context._execFollowUp = { turnId: prevExec.turnId, crossPipeline: true };
                            inheritedFollowUp = true;
                            ChatLog.info?.('ORCH', `[orchestrator] 图表跨管道继承: 上一轮 ${prevPipeline} 分析数据 → skill-pipeline`);
                        }
                    }
                }
            }

            // 3. 通用后续消息继承（意图分析：结合用户输入 + 上一轮任务结构，而非纯关键词）
            // 意图分析器 analyzeFollowUpIntent 综合判断：
            //   - 引用上一轮任务（代词/继续/修改/细化 + 任务结构内容）→ 路由回原管道
            //   - 跨管道 PPT/图表 → 路由到 code/skill
            // 未加载时降级到旧关键词 detectFollowUp（兜底）。
            // PPT 跨管道已在上面处理完毕（context._pptCrossPipeline=true），此处不再重复走 followUp 继承，
            // 避免设置 _execFollowUp → _skillFollowUpPending 阻塞 dispatcher。
            if (!inheritedFollowUp && !context?._pptCrossPipeline) {
                const intentAnalysis = (App.ChatV2.intentAnalyzer && App.ChatV2.intentAnalyzer.analyzeFollowUpIntent)
                    ? App.ChatV2.intentAnalyzer.analyzeFollowUpIntent({ message, prevSpec: prevExec, context })
                    : null;
                const detectResult = detectFollowUp(message, prevExec);
                ChatLog.info?.('ORCH', `[orchestrator] follow-up 检测: msg="${(message||'').slice(0,40)}" prevPipeline=${prevExec?.spec?.pipeline || 'null'} prevSkill=${prevExec?.spec?.params?.skillName || 'null'} → intentAnalyzer=${JSON.stringify(intentAnalysis).slice(0,200)} detectFollowUp=${detectResult ? 'YES' : 'NO'}`);
                const followUp = (intentAnalysis && intentAnalysis.isFollowUp)
                    ? { turnId: prevExec.turnId, spec: prevExec.spec, _intent: intentAnalysis }
                    : detectResult;
                if (followUp) {
                    context = context || {};
                    context._execFollowUp = { turnId: followUp.turnId, intent: followUp._intent || null };
                    const target = (followUp._intent && followUp._intent.targetPipeline) || followUp.spec.pipeline;
                    // 【关键修复】PPT 跨管道意图（context._pptCrossPipeline=true，已在 L1030 设置）
                    // 说明用户要"直接把当前数据做成PPT"——放行 dispatcher 调 ppt_generate，
                    // 不要锁 GeneratePPTReport skill，不要改 kind='code'。
                    if (target === 'ppt' || target === 'code') {
                        if (context._pptCrossPipeline) {
                            // 跨管道：跳过 skill/code 锁定，保持 kind='text' 让 dispatcher 命中 ppt_generate
                            ChatLog.info?.('ORCH', '[orchestrator] 后续消息继承被 PPT 跨管道豁免 → 保持 kind=text，放行 dispatcher');
                        } else if (pptFreeBuild) {
                            // 结构诉求：不锁 GeneratePPTReport，放行 agent-loop 自由调 ppt_ 工具
                            kind = 'text';
                            inheritedFollowUp = true;
                            ChatLog.info?.('ORCH', `[orchestrator] 后续消息继承(PPT结构诉求): ${followUp._intent ? '意图分析(' + followUp._intent.reason + ')' : '命中上一轮 PPT 上下文'} → agent-loop`);
                        } else {
                            // PPT 需锁定 GeneratePPTReport 技能（code-pipeline 直出路径）
                            const pptSkill = (App.cachedSkills || []).find(
                                s => s.name === 'GeneratePPTReport' || s.Name === 'GeneratePPTReport');
                            if (pptSkill) {
                                context._matchedSkill = pptSkill;
                                kind = 'code';
                                inheritedFollowUp = true;
                                ChatLog.info?.('ORCH', `[orchestrator] 后续消息继承: ${followUp._intent ? '意图分析(' + followUp._intent.reason + ')' : '命中上一轮 PPT 上下文'} → code-pipeline`);
                            }
                        }
                    } else if (target === 'skill') {
                        // 图表/技能后续消息：锁定上一轮 Skill（从 _execSpec.params 取回技能名），
                        // 否则 skill-pipeline 因 _matchedSkill 为空直接失败（Phase 5）。
                        const prevSkillName = followUp.spec.params?.skillName;
                        const skill = prevSkillName && (App.cachedSkills || []).find(
                            s => s.name === prevSkillName || s.Name === prevSkillName);
                        if (skill) {
                            context._matchedSkill = skill;
                            kind = 'skill';
                            inheritedFollowUp = true;
                            ChatLog.info?.('ORCH', `[orchestrator] 后续消息继承: ${followUp._intent ? '意图分析(' + followUp._intent.reason + ')' : '命中上一轮 Skill ' + prevSkillName} → skill-pipeline`);
                        }
                    } else {
                        kind = target;
                        inheritedFollowUp = true;
                        ChatLog.info?.('ORCH', `[orchestrator] 后续消息继承: ${followUp._intent ? '意图分析(' + followUp._intent.reason + ')' : '路由回'} ${target}-pipeline`);
                    }
                }
            }

            // ============================================================
            // 【McpDispatcher 确定性调度】短路 Ollama 小模型
            // 所有本地 MCP 工具（computer_*/ppt_*/excel_*/browser_*/wps_*）统一在这里匹配执行。
            // 加新 MCP 类型 → 只改 mcp-dispatcher.js 的 PREFIXES 一行，这里永远不动。
            // 标注 dispatcherBlock：写类工具裁决不通过时用它 break 出来，放行下游路由，
            // 而不是 return 结束整轮（return=短路；break=让路，语义完全不同）。
            // ============================================================
            // 【PPT/Word 意图卡补位 · 方案 A】非 Excel 宿主走 dispatcher 硬短路（命中即提前 return），
            // 整段 LLM 意图分析被跳过 → UI 永远看不到"意图/路由"卡片（Excel 的"✅ 意图分析完成"卡）。
            // 此处为 dispatcher 命中路径提前建一个 round=0 的 routing 块；命中确定性工具时把它 complete 成
            // "确定性命中 xxx"，未命中（放行通用链路）时 complete 成"转通用处理"，与 Excel 卡片对齐观感。
            // 不引入任何 LLM 时延（title/工具名直接复用 dispatcher 已匹配结果）。Excel 仍走后段意图分析块建卡，不重复。
            const _orchHost = App.SmartPilotHost || window.SmartPilotHost;
            const _isNonExcel = !!(_orchHost && !_orchHost.isExcel);
            let routingThoughtStarted = false;
            dispatcherBlock: {
                const toolRegistry = App.ChatV2?.toolRegistry;
                const dispatcher = App.mcpDispatcher || App.ChatV2?.mcpDispatcher;
                ChatLog.info?.('ORCH', `[orchestrator] McpDispatcher 检查: toolRegistry=${!!toolRegistry} dispatcher=${!!dispatcher} tryDispatch=${typeof dispatcher?.tryDispatch} execFollowUp=${!!context?._execFollowUp} inherited=${!!inheritedFollowUp}`);

                // ── 注册表就绪兜底（关键）──
                // dispatcher 早于 text-pipeline 运行，而注册表初始化只在 text-pipeline（阶段1）里做。
                // 若本会话注册表仍为空（桥/工具加载时序竞态），会连锁导致：
                //   ① 工具卡显示裸名（无 displayName）；
                //   ② dispatcher 的"必填参数缺失闸门"查不到 schema → 静默失效 → 空参裸调后端
                //      （实测"筛选 订单金额>3000 且 区域=华东"→ excel_apply_filter 无 range → "范围不能为空"）。
                // 故此处仅在注册表为空时兜底初始化一次，恢复 displayName + 必填闸门 + 参数归一化。
                try {
                    const _regEmpty = !(toolRegistry && typeof toolRegistry.list === 'function' && toolRegistry.list().length > 0);
                    if (_regEmpty && toolRegistry && typeof toolRegistry.init === 'function') {
                        await toolRegistry.init();
                        const _n = (typeof toolRegistry.list === 'function') ? toolRegistry.list().length : 0;
                        ChatLog.info?.('ORCH', `[orchestrator] 注册表为空 → dispatcher 前兜底初始化（${_n} 个工具）`);
                    }
                } catch (_regErr) {
                    ChatLog.warn?.('ORCH', '[orchestrator] 注册表兜底初始化失败: ' + (_regErr?.message || _regErr));
                }
                // ── Skill follow-up 排他 ──
                // 后续消息（"换成柱状图"/"继续美化"/"改成红色"等）已继承上一轮 skill 上下文（_execFollowUp）
                // 或跨管道图表继承（inheritedFollowUp）时，必须交给 skill-pipeline 走完整流程，
                // 不能让 dispatcher 裸调 excel_read_range/create_chart 等工具短路 —— 否则"换图表类型"
                // 会退化成一次读数据、输出原始 JSON。后续意图交给 skill-followup 而非确定性短工具。
                // 【降级守卫】上一轮 Skill 已被停用/不可用时，续轮不再强行续 skill，
                // 交由下面的 __skillDomainMsg 守卫与确定性路由守卫统一回落通用管道。
                // 【PPT 跨管道例外】已标记 context._pptCrossPipeline 时跳过 follow-up 守卫，
                // 让 dispatcher 直接命中 ppt_generate 执行，不被 GeneratePPTReport skill 阻塞。
                const _pptCrossPipelineHit = !!context?._pptCrossPipeline;
                const _skillFollowUpPending = !_pptCrossPipelineHit
                    && (!!inheritedFollowUp || !!context?._execFollowUp)
                    && hasExecutableSkillForMessage(message);

                // ── 确定性 Skill 域 / 路由意图排他（关键）──
                // 图表/美化/格式化/清洗/CRUD 是"多步骤 Skill 域"：内部需先选区识别 + LLM 参数规划，
                // 再组合裸工具执行。若在此让 dispatcher 短路，后端 MatchCapability 常把
                // "换成柱状图"匹配成 excel_read_range（读数据工具，不在受护名单）而**先短路返回**，
                // 导致该命令退化成一次读数据 + 元数据卡片，图表根本没转换。
                // 同理，SkillCreator 的"创建/生成一个 xx 的 skill"意图若让 dispatcher 短路，
                // 后端会把它误匹配成 computer_/读数据工具先返回 not_found，技能创建走不到
                // shouldRouteToSkill。因此只要消息命中 Skill 域关键词或路由策略意图守卫，
                // 就一律跳过 dispatcher、放行给 skill-pipeline。
                // computer_ 类工具指令不含这些词，不受影响。
                // 关键词单一来源：config/skill-domain.js 的 isSkillRoutingMessage()（含 isDomainMessage）。
                const _skillDomainHit = (App.ChatV2?.skillDomain?.isSkillRoutingMessage)
                    ? App.ChatV2.skillDomain.isSkillRoutingMessage(message || '')
                    : (App.ChatV2?.skillDomain?.isDomainMessage
                        ? App.ChatV2.skillDomain.isDomainMessage(message || '')
                        : false);
                // 【降级守卫】命中 Skill 域，但该域已无"真正可用"的 Skill（被停用 / 只有 DESCRIPTORS 空壳）
                // → 不再拦截 dispatcher，放行裸 MCP 工具或通用 text-pipeline。这是"无 Skill 也能完成任务"的关键。
                const _skillDomainMsg = _skillDomainHit && hasExecutableSkillForMessage(message);

                // ── 代码生成意图排他（关键）──
                // "帮我写一段 VBA 代码：拆分工作簿" 这类请求要的是**代码文本**，不是真去拆分工作簿。
                // 但 dispatcher 的关键词匹配会因"拆分工作簿"命中 excel_split_workbook 并**直接执行**，
                // 造成"写代码"变成"实际干活"（实测现象：气泡里返回 files_created 的原始 JSON）。
                // 命中代码生成意图时一律跳过 dispatcher，放行给下游 code-pipeline（只展示、不执行）。
                // 判定与 L1531 `routing().detectCodeGenerationIntent` 同源，保证两处一致。
                const _codeGenMsg = (typeof routing().detectCodeGenerationIntent === 'function')
                    && !!routing().detectCodeGenerationIntent(message || '');
                if (_codeGenMsg && !_skillFollowUpPending && !_skillDomainMsg) {
                    ChatLog.info?.('ORCH', '[orchestrator] 命中代码生成意图 → 跳过 McpDispatcher（只生成代码，不实际执行）');
                }

                // 【PPT 生成意图 · 总豁免】只要用户明确要做 PPT，就跳过所有守卫让 dispatcher 放行，
                // 直接命中 ppt_generate 快速执行（避免 agent-loop 先读数据+设计结构+输出 JSON 确认的冗长链路）。
                // PPT 是"确定产物"诉求（要一份 .pptx 文件），不是"写代码""格式化""图表美化"等需要 Skill 域处理的复合意图。
                const _pptGenIntent = pptIntent && /(?:做|做成|制作|生成|创建|出|转换|转成|新建|弄|搞|整个)[^。？?\n]{0,20}/i.test(message || '');
                if (_pptGenIntent && (_skillFollowUpPending || _skillDomainMsg || _codeGenMsg)) {
                    ChatLog.info?.('ORCH', '[orchestrator] PPT 生成意图 · 豁免 followUp/skillDomain/codeGen 守卫，放行 McpDispatcher');
                }
                const _skipAllGuards = _pptGenIntent; // PPT 生成意图无视所有前置守卫
                // ── 看板/大屏意图排他（关键）──
                // "帮我生成excel大屏看板"是"多步骤复合创作"（读数→聚合→建表→图表→格式化），
                // 但 dispatcher 的后端 Match 常把"看板"消息误匹配成 excel_read_range（读数据工具，
                // 不在受护名单、无必填参数）→ 裸调一次 + 渲染表格卡后 completeTurn 短路收尾，
                // LLM 从未被调用，补救守卫 E（_forceDashboard 看板五步工具链）永远到不了。
                // 无可用看板 Skill（_skillDomainMsg=false）时跳过 dispatcher，放行下游：
                // 守卫 E 注入 _forceDashboard/_forceChart/_forceMultiStep → 多步骤 agent-loop
                // 自由串联 create_sheet/aggregate/create_chart/format_range 原子工具。
                // 看板 Skill 可用时 _skillDomainMsg=true，仍走 skill-pipeline 蓝图引擎（不受影响）。
                const _dashboardIntentMsg = /看板|仪表盘|驾驶舱|dashboard|数据大屏|大屏|数据面板/i.test(message || '')
                    && !_skillDomainMsg;
                if (_dashboardIntentMsg) {
                    ChatLog.info?.('ORCH', '[orchestrator] 看板/大屏意图且无可用 Skill → 跳过 McpDispatcher（避免 read_range 短路），放行 agent-loop 看板工具链');
                }
                if (dispatcher && typeof dispatcher.attachGroupView === 'function' && turn && turn.context) {
                    try { dispatcher.attachGroupView(message, turn); }
                    catch (eGrp) { ChatLog.warn?.('ORCH', '[orchestrator] 分组计划跳过: ' + (eGrp && eGrp.message || eGrp)); }
                }
                const _bareBeautify = !!(dispatcher && typeof dispatcher.isBareTableBeautify === 'function' && dispatcher.isBareTableBeautify(message || ''));
                if (_bareBeautify && _skillDomainMsg) {
                    ChatLog.info?.('ORCH', '[orchestrator] 「美化表格」不走技能规划，改由预设直接格式化');
                }
                // 带文档/图片附件的轮次问的是附件（例如「读取图纸中的尺寸」+ PDF）。
                // 「读取」会被 MatchCapability 配成 excel_read_range，短路后只读选区、正文为空，图纸根本没进模型。
                const _attachmentKind = kind === 'doc' || kind === 'image';
                if (_attachmentKind) {
                    ChatLog.info?.('ORCH', `[orchestrator] 本轮 kind=${kind} 带附件，跳过 McpDispatcher，交给文档/图片管道`);
                    _orchHostLog('INFO', `[orchestrator] 附件轮次 kind=${kind} 跳过 McpDispatcher`);
                }
                const _dispatcherOk = toolRegistry && dispatcher && typeof dispatcher.tryDispatch === 'function'
                    && !_attachmentKind
                    && (_skipAllGuards || _bareBeautify || (!_skillFollowUpPending && !_skillDomainMsg && !_codeGenMsg && !_dashboardIntentMsg));
                // 被守卫跳过时必须留痕：此前这里没有任何日志，"dispatcher 到底跑没跑"完全不可知，
                // 导致"硬短路没生效"只能靠猜（实测踩坑：非 Excel 宿主的选中表格格式化就是这么被绕过的）。
                if (!_dispatcherOk) {
                    _orchHostLog('WARN', `[orchestrator] McpDispatcher 被跳过: registry=${!!toolRegistry} dispatcher=${!!dispatcher} followUp=${_skillFollowUpPending} skillDomain=${_skillDomainMsg} codeGen=${_codeGenMsg} dashboard=${_dashboardIntentMsg} attachment=${_attachmentKind}`);
                }

                if (_dispatcherOk) {
                    // 【PPT/Word 意图卡补位】dispatcher 执行前先建 routing 块（仅非 Excel、非 follow-up）。
                    // Excel 的 routing 块由后段 LLM 意图分析块统一创建，这里跳过避免双卡。
                    if (_isNonExcel && kind === 'text' && !inheritedFollowUp && !context?._execFollowUp && sessionStore.startThought && !routingThoughtStarted) {
                        try {
                            sessionStore.startThought(turn.id, 0, '🔍 正在理解你的意图...', { kind: 'routing' });
                            if (sessionStore.appendThoughtDelta) {
                                // 勿写"Excel 模式"：Word/PPT 用户会误判跑错宿主（实测截图）
                                const _hostLabel = (_orchHost && _orchHost.isPowerPoint) ? 'PowerPoint'
                                    : (_orchHost && (_orchHost.isWord || _orchHost.kind === 'word')) ? 'Word' : '当前宿主';
                                sessionStore.appendThoughtDelta(turn.id, `（${_hostLabel}：确定性工具调度，跳过 LLM 语义理解）`);
                            }
                            routingThoughtStarted = true;
                            ChatLog.info?.('ORCH', '[orchestrator] 非 Excel routing 块已建（dispatcher 前）');
                        } catch (_reErr) { _orchHostLog('WARN', `[orchestrator] 非Excel routing 块渲染异常（忽略）: ${_reErr?.message || _reErr}`); }
                    }
                    try {
                        const d = await dispatcher.tryDispatch({ message, toolRegistry, sessionStore, turn, signal: controller.signal });
                        ChatLog.info?.('ORCH', `[orchestrator] McpDispatcher 结果: hit=${d.hit} tool=${d.toolName || 'none'} score=${d.score || 0}`);
                        _orchHostLog('INFO', `[orchestrator] McpDispatcher 结果: hit=${d.hit} tool=${d.toolName || 'none'} reason=${d.reason || ''} score=${d.score || 0}`);
                        if (d.hit) {
                            // ── 写类工具裁决闸门 ──
                            // 命中 ≠ 命令。有副作用的工具（拆分/合并/导出/写入…）必须先过裁决，
                            // 通过才执行；不通过就 break 出 dispatcherBlock，把本轮让给下游路由
                            // （code-pipeline 出代码 / agent-loop 通用处理），**不结束整轮对话**。
                            // 【PPT 生成意图豁免】用户明确"做PPT"就是确定产物命令，无需写类裁决再确认
                            // （与 L1229 _skipAllGuards 精神一致）。否则 ppt_generate 会被裁决闸门拦下
                            // break 到 agent-loop 再次执行 → 反而制造重复文件、复读"已完成"。
                            // 调度器返回时表已经改完。再问模型「要不要代码」会把这次成功丢掉，
                            // 界面变成「未调用工具」并另生成一段 VBA（19:04 商务蓝：6 张表已格式化）。
                            const _alreadyRan = !!(d.execResult && d.execResult.ok === true);
                            if (_alreadyRan && isSideEffectTool(d.toolName)) {
                                _orchHostLog('INFO', `[orchestrator] ${d.toolName} 已执行成功，跳过写类裁决`);
                            }
                            if (isSideEffectTool(d.toolName) && !_pptGenIntent && !_alreadyRan) {
                                const verdict = await adjudicateWriteHit({
                                    message,
                                    toolName: d.toolName,
                                    displayName: toolRegistry.get(d.toolName)?.displayName || d.toolName,
                                    signal: controller.signal
                                });
                                ChatLog.info?.('ORCH', `[orchestrator] 写类工具裁决: tool=${d.toolName} verdict=${verdict.verdict} source=${verdict.source} reason=${verdict.reason || '-'}`);
                                if (verdict.verdict !== 'command') {
                                    if (verdict.verdict === 'code') {
                                        context = context || {};
                                        context.__forceCodePipeline = true;   // 下游据此置 kind='code'（只出代码）
                                    }
                                    ChatLog.info?.('ORCH', `[orchestrator] 命中项被裁决为 ${verdict.verdict} → 不执行 ${d.toolName}，放行下游路由`);
                                    break dispatcherBlock;
                                }
                            }
                            const resultObj = d.execResult?.result || {};
                            let resultText = '';
                            const execOk = d.execResult && (d.execResult.ok === true || (resultObj.success !== false && !d.execResult.error));
                            if (execOk) {
                                // 截屏软件返回 "base64=<PNG>"：在气泡内直接渲染成图片，不再贴原始 base64 长串
                                const imgHtml = buildScreenCaptureImgHtml(d.toolName, resultObj);
                                if (imgHtml != null) {
                                    resultText = imgHtml;
                                } else {
                                    // Excel 读取数据 → 元数据卡片 + 数据预览表，而非裸 JSON。
                                    // 触发放宽到所有 excel_ 工具：只需结果带读数据元数据（headers/rows/worksheet 等）
                                    // 即可渲染卡片，兼容 read_range/read_selection/get_cell_value/read_file 等变体；
                                    // 非读类结果 hasMeta 判不中 → buildReadRangeMetaCardHtml 返回 null → 走正常文本。
                                    // 清洗/去重结果带 headers+data 结构，会被下方元数据卡误判成
                                    // "已读取数据 N 行"（行数还解析错），必须排除走专用清洗摘要。
                                    const _isCleanTool = d.toolName === 'excel_clean_data' || d.toolName === 'excel_remove_duplicates'
                                        || (d.toolName === 'excel_delete_row' && d.args && d.args.deleteMatchingRows);
                                    const readCard = (!_isCleanTool && (d.toolName || '').indexOf('excel_') === 0)
                                        ? buildReadRangeMetaCardHtml(resultObj) : null;
                                    if (readCard != null) {
                                        resultText = readCard;
                                    } else {
                                        // computer_* 单步：友好结果优先；excel_* 产物类结果同样友好化；
                                        // 其余才走统一安全出口（绝不 String 成 "[object Object]"）
                                        const computerText = computerFriendlyText(d, resultObj);
                                        let excelText = computerText != null ? null : excelFriendlyText(d.toolName, resultObj);
                                        // 【合并兜底 2026-10-10】合并结果以"字符串字段"形态过桥时，若 JSON.parse
                                        // 因截断/转义/编码任何原因失败，excelFriendlyText 返回 null → 整段 pretty
                                        // JSON 会经 safeDisplayText 倒进正文（实测 23:28，截断的 JSON 连渲染层守卫
                                        // 都救不了——括号不配对）。这里用 replaceJsonEchoWithFriendly 兜底：
                                        // 它对合法/无逗号/被截断的合并 JSON 都能免 parse 抽出字段生成摘要。
                                        if (excelText == null && (d.toolName || '') === 'excel_merge_sheets') {
                                            try {
                                                const _txt = (resultObj && typeof resultObj.result === 'string') ? resultObj.result
                                                    : (typeof resultObj === 'string' ? resultObj : JSON.stringify(resultObj || {}));
                                                const _efOut = App.ChatV2 && App.ChatV2.toolRegistryUtil
                                                    && typeof App.ChatV2.toolRegistryUtil.replaceJsonEchoWithFriendly === 'function'
                                                    ? App.ChatV2.toolRegistryUtil.replaceJsonEchoWithFriendly(_txt) : null;
                                                if (_efOut) {
                                                    ChatLog.info?.('ORCH', '[orchestrator] merge_sheets 结果经 replaceJsonEchoWithFriendly 兜底摘要化');
                                                    excelText = _efOut;
                                                }
                                            } catch (_me) { /* 兜底失败不影响主流程 */ }
                                        }
                                        const pptText = (computerText != null || excelText != null) ? null : pptFriendlyText(d.toolName, resultObj);
                                        // Word 宿主原先缺这一支 → word_* 结果整坨 JSON 进正文（用户反馈"没有渲染"）
                                        const wordText = (computerText != null || excelText != null || pptText != null)
                                            ? null : wordFriendlyText(d.toolName, resultObj);
                                        resultText = computerText != null
                                            ? computerText
                                            : (excelText != null
                                                ? excelText
                                                : (pptText != null
                                                    ? pptText
                                                    : (wordText != null
                                                        ? wordText
                                                        : (App.ChatV2?.toolRegistryUtil?.safeDisplayText || safeDisplayTextWrapper)(resultObj, 1000))));
                                    }
                                }
                            } else {
                                // 【#3 修复】dispatcher 命中但执行失败：不静默收尾，放行下游 agent-loop 二次尝试，
                                // 让 LLM 换工具/参数或如实说明失败原因（agent-loop 有 failStreak 熔断，不会无限重放同一失败调用）。
                                const _failMsg = App.ChatV2.toolRegistryUtil.safeErrorText(resultObj.error || d.execResult?.error || d.error || '未知原因');
                                _orchHostLog('ERROR', `[orchestrator] dispatcher 命中 ${d.toolName} 但执行失败，放行下游 agent-loop: ${_failMsg}`);
                                ChatLog.warn?.('ORCH', `[orchestrator] dispatcher 执行失败 → 放行下游: ${d.toolName} :: ${_failMsg}`);
                                try { if (sessionStore.appendThoughtDelta) sessionStore.appendThoughtDelta(turn.id, `\n⚠️ 确定性工具 ${d.toolName} 执行失败（${_failMsg.slice(0, 120)}），转交通用处理尝试其他方式。`); } catch (_dErr) { /* 忽略 */ }
                                if (!context) context = {};
                                context._dispatcherFailed = { tool: d.toolName, error: _failMsg };
                                break dispatcherBlock;
                            }
                            const displayName = toolRegistry.get(d.toolName)?.displayName || d.toolName;
                            // ── 解除 renderer 的工具卡屏蔽 ──
                            // turn_start 时 isText 且无 _enableSteps → handle._suppressToolCards = true
                            // 导致后续 addToolCall / updateToolCall 事件被 renderer 跳过（L3786 break）
                            // 必须在发工具卡事件前先设 _enableSteps = true 并通知 renderer 解除 suppress
                            turn._enableSteps = true;
                            if (context) context._enableSteps = true;
                            if (sessionStore.updateTurn) sessionStore.updateTurn(turn.id, { _enableSteps: true });

                            // ── dispatcher 快速路径：补渲染工具卡（agent-loop 才会自动发 tool_call 事件）──
                            // 必须走 sessionStore.addToolCall / updateToolCall，
                            // 这俩会 notify renderer → renderToolCard 往 toolsEl 里真插 .exec-round DOM，
                            // finalizeAIBubble 才能 querySelectorAll('.exec-round') 到东西。
                            // ⚠ 例外：确定性步骤链（批量美化/拆分后逐表美化）已经把主调用作为步骤第一行展示，
                            // 再补一张裸工具卡就是 `excel_format_range {} · 销售二部 · 1ms` 这种重复行
                            //（实测 2026-10-22 22:31 用户反馈"没有任何改善"）。dispatcher 置
                            // execResult.suppressToolCard=true → 跳过补卡，过程只由步骤列表呈现。
                            const _suppressCard = !!(d.execResult && d.execResult.suppressToolCard === true
                                && Array.isArray(d.execResult.steps) && d.execResult.steps.length > 0);
                            try {
                                if (_suppressCard) {
                                    ChatLog.info?.('ORCH', `[orchestrator] 跳过重复工具卡: ${d.toolName}（步骤面板已覆盖主调用，${d.execResult.steps.length} 步）`);
                                } else {
                                    const dispTc = sessionStore.addToolCall(turn.id, {
                                        name: d.toolName,
                                        displayName: displayName,
                                        args: d.args || {}
                                    });
                                    sessionStore.updateToolCall(turn.id, dispTc.id, {
                                        status: execOk ? 'completed' : 'failed',
                                        result: resultObj,
                                        error: execOk ? null : (resultObj.error || d.execResult?.error || d.error || '未知原因'),
                                        durationMs: d.execResult?.elapsedMs || 1
                                    });
                                }
                            } catch (_tcErr) { _orchHostLog('WARN', `[orchestrator] 工具卡渲染异常（不影响主流程）: ${_tcErr?.message || _tcErr}`); }

                            // 链式操作（computer_ 多工具，如"打开→播放"）：把执行步骤挂到 turn，走通用步骤渲染
                            const stepArr = d.execResult?.steps;
                            if (Array.isArray(stepArr) && stepArr.length) {
                                turn._steps = stepArr;
                                turn._enableSteps = true;
                                if (context) context._enableSteps = true;
                                // 发 turn_update patch：渲染层据 _enableSteps 重算 _suppressSteps=false，
                                // 否则 turn_start 时已锁 _suppressSteps=true，步骤面板不会显示
                                if (sessionStore.updateTurn) sessionStore.updateTurn(turn.id, { _enableSteps: true });
                                // 【步骤统一到 cbx v20261003】链式步骤写进 thoughtList，
                                // 由 contentEl 顶部的 cbx 列表展示（不再渲染旧 exec-plan-card 绿面板）
                                if (typeof sessionStore.upsertStepThought === 'function') {
                                    stepArr.forEach(function (s, i) {
                                        sessionStore.upsertStepThought(turn.id, {
                                            id: (s && (s.id || s.title)) || ('chain_step_' + i),
                                            title: (s && (s.title || s.name)) || ('步骤 ' + (i + 1)),
                                            detail: (s && (s.detail || s.description)) || '',
                                            status: (s && s.status) || 'done',
                                            index: i
                                        });
                                    });
                                }
                                // 链式操作：不把原始执行对象(pid/target 等)当正文倾倒——可读信息已由步骤面板承载，
                                // 只给一行忠于实际状态的中性小结，避免标题下方再叠一段不可读裸文本。
                                const hasFailed = stepArr.some(s => s.status === 'failed');
                                // 若正文已是渲染产物（如窗口截图 <img>），保留它而不被小结覆盖——截图即交付物
                                const alreadyRendered = typeof resultText === 'string' &&
                                    (resultText.includes('screen-capture') || resultText.includes('read-range'));
                                if (!alreadyRendered) {
                                    resultText = (hasFailed ? '⚠️ ' : '✅ ') + '操作完成：' +
                                        stepArr.map(s => {
                                            const mark = s.status === 'failed' ? '✕' : (s.status === 'done' ? '✓' : '…');
                                            return (s.title || '步骤') + ' ' + mark;
                                        }).join(' · ');
                                }
                            }
                            // 【PPT/Word 意图卡补位】dispatcher 命中确定性工具 → 收尾 routing 块
                            // （标题对齐 Excel 的"✅ 意图分析完成"卡片；不引入 LLM 时延）。
                            if (routingThoughtStarted && sessionStore.completeThought) {
                                const _dn = displayName || d.toolName || '工具';
                                if (sessionStore.appendThoughtDelta) {
                                    sessionStore.appendThoughtDelta(turn.id, `\n命中确定性工具：${_dn}（score=${d.score || 0}，跳过 LLM 语义理解，直接执行）`);
                                }
                                sessionStore.completeThought(turn.id, `✅ 已识别：确定性命中 ${_dn}（跳过 LLM 意图分析）`, { round: 0, kind: 'routing' });
                            }
                            try {
                                const envApi = App.ChatV2 && App.ChatV2.toolEnvelope;
                                if (envApi && typeof envApi.fromRaw === 'function') {
                                    const env = envApi.fromRaw(d.toolName, resultObj, execOk);
                                    turn._envelopes = [env];
                                    if (typeof resultText === 'string' && envApi.looksLikeRawJson(resultText)) {
                                        const friendly = envApi.formatForUser(env, displayName || d.toolName);
                                        if (friendly) resultText = friendly;
                                    }
                                }
                            } catch (_envErr) { /* 信封失败不影响原文 */ }
                            sessionStore.appendDelta(turn.id, 'content', resultText || `✅ ${displayName} 执行成功`);
                            sessionStore.completeTurn(turn.id);
                            _activeControllers.delete(turn.id);
                            return turn.id;
                        } else if (d && d.deferToPlanner) {
                            // 关键词盖不住的外观（如「统一为中国红」）不在这里执行。
                            // 强制多步骤，让模型按原话规划：每张表 + 点名的颜色都要留下。
                            const bag = (turn && turn.context) ? turn.context : (context || {});
                            bag._forceMultiStep = true;
                            bag._enableSteps = true;
                            bag._semanticStyleTarget = String(d.styleTarget || '');
                            if (turn) {
                                turn.context = bag;
                                turn._enableSteps = true;
                            }
                            context = bag;
                            ChatLog.info?.('ORCH', `[orchestrator] 不走硬短路，交给模型规划${bag._semanticStyleTarget ? '（外观「' + bag._semanticStyleTarget + '」）' : ''}`);
                            _orchHostLog('INFO', `[orchestrator] 不走硬短路，交给模型规划${bag._semanticStyleTarget ? ' style=' + bag._semanticStyleTarget : ''}`);
                        }
                    } catch (mcpErr) {
                        ChatLog.warn?.('ORCH', `[orchestrator] McpDispatcher 异常: ${mcpErr?.message || mcpErr}`);
                        _orchHostLog('ERROR', `[orchestrator] McpDispatcher 异常: ${mcpErr?.message || mcpErr}`);
                    }
                }
                // 【PPT/Word 意图卡补位】dispatcher 未提前 return（未命中工具 / 被写类裁决放行到下游）→
                // 把前面建的 routing 块收尾为"转通用处理"，避免它永远停在 running 态。
                // （命中且已执行的路径在 return 前已收尾，不会重复。）
                if (routingThoughtStarted && sessionStore.completeThought) {
                    sessionStore.completeThought(turn.id, '✅ 意图分析完成（无匹配确定性工具，转通用处理）', { round: 0, kind: 'routing' });
                }
            }

            // ============================================================
            // LLM 意图分析前置（Phase 6）：非琐碎请求先经 LLM 语义理解，
            // 产出结构化意图（要做什么/对什么数据/产出什么）驱动路由与执行。
            // 关键词只做快速通道（问候/琐碎）与 LLM 失败时的兜底。
            // 触发范围：needPlanning 返回 true（关键词命中多类别）或 needs_llm_check（关键词不确定）的请求。
            //   关键：needPlanning()===true 只是"关键词计数命中"，不等于"理解意图"——
            //   这类请求正是"命中关键词但干不了活"的高发区（如"分析这个表格"、"看看哪些产品卖得好"）。
            //   因此只要非琐碎（非 false），都先经 LLM 语义理解，LLM 的 needsPlanning 是权威决策。
            // 结果写入 context：
            //   - _forceMultiStep=true（LLM 判定需多步骤）→ 下游 Skill/CRUD/python 路由据此跳过，
            //     text-pipeline 据此直接进入多步骤管道。
            //   - _intentSingleOp=true（LLM 判定单操作/问答）→ 阻止关键词计数把单操作误判为多步骤。
            //   - _intent=结构化意图（operations/targetData/output），供下游执行参考。
            // ============================================================
            let llmIntent = null;
            // ============================================================
            // 确定性 Skill 预判 — 但**只跳过 needsPlanning 副作用，不跳过 operations 提取**
            // ------------------------------------------------------------
            // 历史原因跳过 LLM：LLM 会把"格式化表格"这类单操作判成 needsPlanning=true →
            // forceMultiStep → Beautifier/Formatter/Chart 等确定性 Skill 被拦进 text-pipeline。
            // 但 LLM analyzeNewIntent 返回的 operations 字段没有副作用——它就是一张"用户要做什么"
            // 的能力需求清单，Route Validator 需要它做「LLM 动态意图 → capability 匹配」。
            //
            // 所以精细拆分：
            //   preMatchDeterministic=true → 只屏蔽 needsPlanning 的 forceMultiStep 设置
            //   LLM analyzeNewIntent 照常跑 → operations 正常提取 → 喂给 Route Validator
            //   LLM 超时/失败 → operations=null → Route Validator 自动降级到正则
            //
            // 这相当于给路由层注入了杰瑞VBE那种"AI 动态理解"的能力，但又有
            // capability 确定性校验兜底——不会因为 LLM 弱模型误判就路由错。
            // ============================================================
            const deterministicKwTest = () => (
                (App.ChatV2?.skillDomain?.isDomainMessage)
                    ? App.ChatV2.skillDomain.isDomainMessage(message || '')
                    : false
            );
            const isDeterministicHtmlReportIntent = () => {
                if (!(App.cachedSkills || []).some(s => String(s.pipeline || s.Pipeline).toLowerCase() === 'html')) return false;
                const m = (message || '').toLowerCase();
                const hasAction = /生成|创建|做出|做一个|做一份|出一份|展示|输出/.test(m);
                const hasHtmlHint = /html|网页报表|网页 |\.html|web 报表|web报表/.test(m);
                return hasAction && hasHtmlHint;
            };
            const preMatchDeterministic = deterministicKwTest() || isDeterministicHtmlReportIntent();
            // PPT 结构诉求（pptFreeBuild）需放行 agent-loop，仍交 LLM 意图分析
            const pptNeedsIntentAnalysis = pptFreeBuild || !pptIntent;

            // === LLM 意图分析：始终尝试跑，只是消费方式因 preMatchDeterministic 不同 ===
            if (kind === 'text' && !inheritedFollowUp && !context?._execFollowUp && pptNeedsIntentAnalysis) {
                try {
                    const taskPlanner = App.ChatV2.TaskPlanner;
                    // preMatchDeterministic 时跳过 TaskPlanner.needPlanning（确定性 Skill 不需要 LLM 多步骤判定），
                    // 但仍跑 LLM 提取 operations → Route Validator 需要这个动态意图信号
                    const planCheck = preMatchDeterministic ? false
                        : ((taskPlanner && typeof taskPlanner.needPlanning === 'function')
                            ? taskPlanner.needPlanning(message || '') : false);
                    const shouldCallLlm = preMatchDeterministic || planCheck === true || planCheck === 'needs_llm_check';
                    if (shouldCallLlm) {
                        // 【统一 thoughtList】路由阶段的"理解意图"变成 thoughtList 的第 0 个块（kind='routing'），
                        // 与后续 agent-loop 的第 1 轮思考/工具调用同一条时间轴，不再走 statusIndicatorEl 独立通道。
                        // routingThoughtStarted 为真说明前面（非 Excel dispatcher 路径）已建过 round=0 routing 块，
                        // 这里跳过，避免双卡；非 Excel 一般不走到此（needPlanning 闸门为 Excel 中心）。
                        if (!routingThoughtStarted && sessionStore.startThought) {
                            sessionStore.startThought(turn.id, 0, '🔍 正在理解你的意图...', { kind: 'routing' });
                        } else if (!routingThoughtStarted && sessionStore.setTurnStatus) {
                            sessionStore.setTurnStatus(turn.id, '🔍 正在理解你的意图...');
                        }
                        // 模型返回前，卡片里先写下已经能确定的步骤。否则 超时后只剩一句降级说明，
                        // 看起来像分析没有过程。
                        if (sessionStore.appendThoughtDelta) {
                            const _styleNow = String((context && context._semanticStyleTarget) || '').trim();
                            const _enabledNow = (App.cachedSkills || []).filter(function (s) {
                                return s && s.isEnabled !== false && s.IsEnabled !== false;
                            });
                            const _proc = ['请求：' + String(message || '').trim().slice(0, 80)];
                            if (_styleNow) {
                                let _styleLine = '外观「' + _styleNow + '」不走预设，和原话里的其他要求一起交给模型';
                                try {
                                    const _g = App.ChatV2.turnGoal && App.ChatV2.turnGoal.build(message || '');
                                    const _kinds = [];
                                    const _seenK = {};
                                    ((_g && _g.doneWhen) || []).forEach(function (d) {
                                        if (!d || d.kind === 'reply' || _seenK[d.kind]) return;
                                        _seenK[d.kind] = true;
                                        _kinds.push(d.kind);
                                    });
                                    if (_kinds.indexOf('merge') >= 0) {
                                        _styleLine = '原话有两项：合并选中工作表，以及外观「' + _styleNow + '」。两项都交给模型，不单做成格式化';
                                    }
                                } catch (_) { /* 目标拆不开时保留外观一句 */ }
                                _proc.push(_styleLine);
                            }
                            _proc.push(_enabledNow.length
                                ? ('可用技能 ' + _enabledNow.length + ' 个，正在让模型从目录里选')
                                : '可用技能为 0，已关闭的技能不参与，这次不走技能');
                            _proc.push('正在等模型给出结构化意图…');
                            sessionStore.appendThoughtDelta(turn.id, _proc.join('\n'));
                        }
                        const ia = App.ChatV2.intentAnalyzer;
                        if (ia && typeof ia.analyzeNewIntent === 'function') {
                            // 流式 reasoning 直接追加进 routing 块（无 running 块时 appendThoughtDelta 会兜底建块，
                            // 但此处 routing 块必然存在且 running）
                            const intentAbort = new AbortController();
                            const _onTurnAbort = function () { intentAbort.abort(); };
                            if (controller.signal) {
                                if (controller.signal.aborted) intentAbort.abort();
                                else controller.signal.addEventListener('abort', _onTurnAbort);
                            }
                            let intentTimedOut = false;
                            const intentPromise = ia.analyzeNewIntent({
                                message, context, signal: intentAbort.signal,
                                onReasoning: (delta) => {
                                    if (delta && !intentTimedOut && sessionStore.appendThoughtDelta) {
                                        sessionStore.appendThoughtDelta(turn.id, delta);
                                    }
                                }
                            }).catch(function () { return null; });
                            const timeoutPromise = new Promise(resolve => setTimeout(() => {
                                intentTimedOut = true;
                                intentAbort.abort();
                                resolve(null);
                            }, INTENT_ANALYSIS_TIMEOUT_MS));
                            llmIntent = await Promise.race([intentPromise, timeoutPromise]);
                            if (controller.signal) controller.signal.removeEventListener('abort', _onTurnAbort);
                            const _waitLine = '正在等模型给出结构化意图…';
                            if (llmIntent) {
                                context = context || {};
                                context._intent = llmIntent;
                                // ── 路由建议的落地判定（算一次，显示与 needsPlanning 决策共用） ──
                                // route.skillId 只是"建议"：能不能被执行，取决于下方「LLM route 覆盖 regex」
                                // 能否解析成真实技能，以及会不会被"多步骤分流"这道闸门整段跳过（见下）。
                                const _routeConf = llmIntent.route ? Number(llmIntent.route.confidence) : 0;
                                const _routeSkillRoute = !!(llmIntent.route && String(llmIntent.route.pipeline) === 'skill'
                                    && llmIntent.route.skillId && String(llmIntent.route.skillId) !== 'free_chat');
                                const _routeRes = (_routeSkillRoute && _routeConf >= ROUTE_T.LLM_ROUTE_MIN_CONFIDENCE)
                                    ? resolveLlmRouteTarget(llmIntent.route, message)
                                    : { ok: false, skill: null, reason: _routeSkillRoute ? '置信度不足 0.5' : '非技能路由' };
                                const _routeMetaName = (_routeRes.ok && _routeRes.skill)
                                    ? String(_routeRes.skill.name || _routeRes.skill.Name || '') : '';
                                // 元技能判据与下方 RouteValidator 前（isMetaSkillPre）完全一致：
                                // ROUTING_POLICIES 里声明了独占 intent 的生成器（SkillCreator / GenerateOCRParseScript）。
                                const _sdRt = App.ChatV2?.skillDomain;
                                const _routeMetaPolicy = (_routeMetaName && _sdRt && _sdRt.ROUTING_POLICIES)
                                    ? (_sdRt.ROUTING_POLICIES[_routeMetaName] || _sdRt.ROUTING_POLICIES[_routeMetaName.toLowerCase()])
                                    : null;
                                const _routeIsMetaSkill = !!(_routeMetaPolicy && _routeMetaPolicy.intent);
                                const _routePinned = !!context?._pinnedSkill;
                                // 非确定性预判 + needsPlanning：只有当路由不是元技能时，下面才会真的设 _forceMultiStep
                                const _routeWillBeDiverted = !preMatchDeterministic && !!llmIntent.needsPlanning && !_routeIsMetaSkill;

                                // 纯操作意图（格式化/样式/结构）不应被 forceMultiStep：
                                // 否则 text-pipeline 的 PURE_FMT_RE 短路会被跳过，格式化请求会被过度编排成
                                // "读数据 → 分析 → 总结"，并触发 ResultAggregator 的数据分析摘要。
                                const pb = App.ChatV2.promptBuilder;
                                const _isOperationOnly = pb && typeof pb.detectOperationOnlyIntent === 'function'
                                    ? pb.detectOperationOnlyIntent(message || '')
                                    : false;

                                // 把分析结论写进 routing 块 content（此前永远为空 → 展开显示"（无推理文字）"）
                                if (sessionStore.replaceThoughtText) {
                                    sessionStore.replaceThoughtText(turn.id, _waitLine, '');
                                }
                                if (sessionStore.appendThoughtDelta) {
                                    const lines = [];
                                    if (llmIntent.summary) lines.push('意图：' + llmIntent.summary);
                                    if (llmIntent.reason) lines.push('理由：' + llmIntent.reason);
                                    if (llmIntent.route) {
                                        // ✨ 如实显示 LLM 的路由建议：route.skillId 只是"建议"，
                                        // 能不能真的被执行取决于下方「LLM route 覆盖 regex」的解析结果。
                                        // 旧实现用一张硬编码白名单（含 skill_creator/html_report）判断"id 存在"，
                                        // 而这两个别名恰恰是解析器**认不出**的（既不在 DESCRIPTORS 也不是技能名）
                                        // → 白名单压制了"未执行"标记，界面写着 skill_creator，实际跑的是
                                        // GenerateOCRParseScript（实测"为火车票创建一个解析器"，置信度 0.8）。
                                        // 现在统一走 resolveLlmRouteTarget，与执行层同一口径，并补上"实际执行谁"。
                                        const _rp = String(llmIntent.route.pipeline || 'text');
                                        const _ridRaw = llmIntent.route.skillId;
                                        let _routeTail = '';
                                        if (llmIntent.route.droppedSkillId) {
                                            _routeTail = '（目录中没有「' + llmIntent.route.droppedSkillId + '」，已改为通用路径）';
                                            _orchHostLog('WARN', '意图路由丢弃目录外技能「' + llmIntent.route.droppedSkillId + '」→ free_chat');
                                        } else if (!_routeSkillRoute) {
                                            // 非技能路由（text / free_chat）：无需校验，照原样显示
                                        } else if (!_routeRes.ok) {
                                            _routeTail = '〔LLM 建议，未采用：' + _routeRes.reason + '〕';
                                            // 只有确认后面真的会走一次 regex 路由时，才敢断言"实际按谁执行"
                                            // （多步骤分流会让整段 Skill 路由被跳过，此时谈不上 regex 命中）
                                            const _willRunRouting = !inheritedFollowUp && !context?._forceMultiStep
                                                && !_routePinned && !_routeWillBeDiverted;
                                            if (_willRunRouting) {
                                                const _fb = shouldRouteToSkill(message, context);
                                                const _fbName = _fb && (_fb.displayName || _fb.name || _fb.Name);
                                                if (_fbName) _routeTail += '，实际按关键词命中「' + _fbName + '」执行';
                                            }
                                        } else if (_routeWillBeDiverted) {
                                            // 元技能以外的技能路由 + needsPlanning → 下面仍会设 _forceMultiStep，
                                            // 整段 Skill 路由（含这个技能）被跳过，实际由 agent 自由挑工具。
                                            // 显示必须说清楚，否则又是一条"说要执行、其实没执行"的假路由。
                                            _routeTail = '〔该技能通道将被多步骤规划跳过，实际由 agent 自行挑工具执行〕';
                                        }
                                        lines.push('路由：' + _rp +
                                            (_ridRaw ? ' → ' + _ridRaw : '') +
                                            (llmIntent.route.confidence ? '（置信度 ' + llmIntent.route.confidence + '）' : '') +
                                            _routeTail);
                                    }
                                    if (Array.isArray(llmIntent.operationLabels) && llmIntent.operationLabels.length) {
                                        lines.push('操作：' + llmIntent.operationLabels.join('、'));
                                    } else if (Array.isArray(llmIntent.operations) && llmIntent.operations.length) {
                                        lines.push('操作：' + llmIntent.operations.join('、'));
                                    }
                                    if (lines.length) sessionStore.appendThoughtDelta(turn.id, '\n' + lines.join('\n'));
                                }
                                // 关键：preMatchDeterministic 时不消费 needsPlanning —— 确定性 Skill
                                // 本身就是内部多步骤管道（planAndRender + VBA），forceMultiStep 会把它们错拦。
                                // operations 字段则无论如可供 Route Validator 使用。
                                //
                                // 【元技能同样不消费 needsPlanning】（实测 2026-10-03 09:06）
                                // 用户："帮我创建一个自动整理表格的skill"，LLM 给出
                                //   route={pipeline:'skill', skillId:'skill_creator', confidence:0.9}
                                //   operations=[read,analyze,write] → needsPlanning=true
                                // 旧逻辑无条件消费 needsPlanning → _forceMultiStep=true → 下面 1831 行
                                // 的 Skill 路由闸门被关掉，matchedSkill 永远 null → 掉进通用多步骤 agent-loop，
                                // 由弱模型自己挑工具 → 它调了 excel_read_range 读 Sheet1!A1，回
                                // "表里没数据，请先输入数据"，SkillCreator 一次都没跑（界面却写着路由 0.9）。
                                // 根因：operations 里的 read/analyze/write 是**将要被创建的那个技能**要做的动作，
                                // 不是本回合要执行的动作，LLM 因此误判 needsPlanning。
                                // 元技能是"生成器"，产物是新技能/解析脚本，本就不需要 Excel 选区或多步骤编排，
                                // 且下方 2157 行本来就写着"元技能必须独占 skill-pipeline，不交给 LLM 自由发挥"——
                                // 那道保护正因为这里的 _forceMultiStep 而永远进不去。
                                // preMatchDeterministic 可能因为初始化时序（skillDomain 还没注入 App.ChatV2）
                                // 在这个时间点返回 false——再加一道直接调 isDomainMessage 的二次检查兜底。
                                // 场景（实测 2026-10-07 19:11）：cleaner 合成空壳在 L2127 才被 shouldRouteToSkill 赋值，
                                // 但那之前 L1973 已经因为 preMatchDeterministic=false（isDomainMessage 返回空）
                                // + needsPlanning=true → _forceMultiStep=true → L2126 守卫被卡死 → matchedSkill=null
                                // → 整条 skill-pipeline 路由链掐死 → 掉进 agent-loop 自由发挥调 aggregate/create_chart。
                                const _kwTestNow = (() => {
                                    const sd = App.ChatV2?.skillDomain;
                                    if (sd && typeof sd.isDomainMessage === 'function') return sd.isDomainMessage(message || '');
                                    return false;
                                })();
                                const _skipNeedsPlanning = preMatchDeterministic || _kwTestNow;
                                if (!_skipNeedsPlanning) {
                                    if (_isOperationOnly) {
                                        // 纯格式化/样式请求：即使 LLM 误判 needsPlanning（如"设置表头底色"
                                        // 因含"设置"被算成 write 类别、凑够多类别），也按单操作路由，
                                        // 不进多步骤管道 —— 避免被编排成"读数据→分析→总结"。
                                        context._intentSingleOp = true;
                                        ChatLog.info?.('ORCH', `[orchestrator] 纯操作意图（格式化/样式）覆盖 LLM needsPlanning=${!!llmIntent.needsPlanning} → _intentSingleOp，不走多步骤管道`);
                                    } else if (llmIntent.needsPlanning && !_routeIsMetaSkill) {
                                        context._forceMultiStep = true;
                                        context._enableSteps = true;
                                        ChatLog.info?.('ORCH', `[orchestrator] LLM 意图分析: 需多步骤规划 → text-pipeline (${llmIntent.reason || 'no reason'})`);
                                        if (_routeRes.ok) {
                                            ChatLog.warn?.('ORCH', `[orchestrator] 多步骤规划跳过 Skill 路由(${llmIntent.route.skillId}) → agent 自行挑工具`);
                                            _orchHostLog('WARN', `LLM 判 needsPlanning → _forceMultiStep，Skill 路由(${llmIntent.route.skillId})被跳过，改由 agent 自由挑工具（如需它执行请设 route needsPlanning=false）`);
                                        }
                                    } else if (llmIntent.needsPlanning && _routeIsMetaSkill) {
                                        // 元技能优先：保留 Skill 路由，让 skill-pipeline 独占执行
                                        context._intentSingleOp = true;
                                        ChatLog.info?.('ORCH', `[orchestrator] 元技能路由(${_routeMetaName}，置信度 ${_routeConf}) → 不消费 needsPlanning，交由 skill-pipeline 独占执行`);
                                        _orchHostLog('INFO', `LLM 路由命中元技能 ${_routeMetaName}（置信度 ${_routeConf}）→ 忽略 needsPlanning，强制 skill-pipeline 独占执行`);
                                    } else {
                                        context._intentSingleOp = true;
                                        ChatLog.info?.('ORCH', `[orchestrator] LLM 意图分析: 单操作/问答 → 保持关键词路由 (${llmIntent.reason || 'no reason'})`);
                                    }
                                } else {
                                    ChatLog.info?.('ORCH', `[orchestrator] 确定性路由预判命中 → LLM 仅提取 operations（跳过 needsPlanning 消费），operations=${JSON.stringify(llmIntent.operations || []).slice(0,200)}`);
                                }
                            } else {
                                ChatLog.warn?.('ORCH', `[orchestrator] LLM 意图分析超时(${INTENT_ANALYSIS_TIMEOUT_MS}ms)，Route Validator 将降级到正则兜底`);
                                const _timeoutLine = intentTimedOut
                                    ? ('模型在 ' + INTENT_ANALYSIS_TIMEOUT_MS + 'ms 内没有返回结构化意图，后面按关键词和工具继续')
                                    : '模型没有返回结构化意图，后面按关键词和工具继续';
                                const _replaced = sessionStore.replaceThoughtText
                                    && sessionStore.replaceThoughtText(turn.id, _waitLine, _timeoutLine);
                                if (!_replaced && sessionStore.appendThoughtDelta) {
                                    sessionStore.appendThoughtDelta(turn.id, '\n' + _timeoutLine);
                                }
                            }
                        }
                        // 意图分析结束 → complete round=0 routing 块，统一 thoughtList 时间轴
                        if (sessionStore.completeThought) {
                            sessionStore.completeThought(turn.id, llmIntent ? '✅ 意图分析完成' : '意图分析未完成，已继续', { round: 0, kind: 'routing' });
                        }
                    }
                } catch (iaErr) {
                    ChatLog.warn?.('ORCH', '[orchestrator] LLM 意图分析失败（不影响主流程，Route Validator 将降级到正则兜底）:', iaErr?.message);
                    // 失败也要 complete 掉 routing 块，防止它永远 running
                    if (sessionStore.completeThought) {
                        sessionStore.completeThought(turn.id, '⚠️ 意图分析失败', { round: 0, kind: 'routing', status: 'failed' });
                    }
                }
            }

            // ════════════════════════════════════════════════════════════════
            // 路由决策：LLM 优先，regex 兜底
            // ────────────────────────────────────────────────────────────────
            // 核心变化：之前 regex (shouldRouteToSkill) 是路由决策者，LLM operations 只是 Route Validator 的输入。
            // 现在 LLM route 字段直接决定走哪条通道——regex 只在 LLM 超时/失败时才用。
            //
            // 为什么 LLM 更准：
            //   "做个看板" → LLM 语义理解 → route.pipeline="skill", skillId="dashboard"
            //   "什么是看板" → LLM 语义理解 → route.pipeline="text", skillId="free_chat"
            //   regex 永远分不清楚这两个——因为它只看关键词命中，不理解意图是"创建"还是"查询"
            //
            // 执行顺序：
            //   ① regex shouldRouteToSkill 先跑（快路径，给 UX 进度条即时反馈）
            //   ② LLM route 结果**覆盖** regex——LLM 说啥就是啥（只要 confidence > 0.5）
            //   ③ LLM 没 route / confidence 太低 → 保留 regex 结果
            // ════════════════════════════════════════════════════════════════
            let matchedSkill = null;
            // ✨【错误检查短路】"检查/检测是否包含错误"类意图必须走 text-pipeline，
            // 由 agent 直接调用 excel_check_errors（COM SpecialCells 精确定位 #DIV/0!、#N/A 等错误单元格）。
            // 不能被 LLM 路由到 SmartDataAnalyzer——它是通用"异常值分析"，不认识公式错误，
            // 实测"检查当前工作表是否包含错误"被路由过去后只会报"0 值异常"，漏掉 #DIV/0!。
            const _msgStr = String(message || '');
            // --------------------------------------------------------------
            // 【展示结果意图守卫】"把结果用表格展示/显示/呈现/列出"是"把数据渲染成表格展示"，
            // 不应被 LLM 误路由到 SmartTableBeautifier / SmartTableFormatter（美化/格式化工件表）。
            // 美化只在用户显式说"美化/格式化/调整样式/配色/边框/加粗/主题"时才触发。
            // 实测："把结果用表格展示"被判 beautifier → 跑了一堆改工作表的 VBA，结果却藏在折叠盒里。
            // 处理：纯展示诉求时强制 text-pipeline(free_chat)，由 agent 读取数据并以表格/图表渲染展示。
            // --------------------------------------------------------------
            const _hasDisplayVerb = /(展示|显示|呈现|列出|输出|渲染|查看|看看|做成)/.test(_msgStr);
            const _hasDisplayTarget = /(表格|表|图表|列表|清单|明细|数据|结果)/.test(_msgStr);
            const _hasBeautify = /(美化|美一下|漂亮|美观|好看|样式|风格|配色|颜色|边框|冻结|首行|表头|字体|对齐|居中|加粗|底色|背景|主题|格式|格式化|调整样式|自动调整|统一字体|autofit|beautif)/i.test(_msgStr);
            const _displayResultIntent = _hasDisplayVerb && _hasDisplayTarget && !_hasBeautify;
            if (_displayResultIntent) {
                // 展示诉求 ≠ 分析诉求：剥离 analyze，避免触发"必须给分析结论"
                // （MultiStepEngine 见 analyze → ReAct 动态规划；prompt-builder 见 analyze 语义 → 注入分析提示）
                if (Array.isArray(llmIntent?.operations) && llmIntent.operations.some(op => /^analyze$/i.test(String(op)))) {
                    llmIntent.operations = llmIntent.operations.filter(op => !/^analyze$/i.test(String(op)));
                    ChatLog.info?.('ORCH', '[orchestrator] 展示结果意图 → 剥离 operations.analyze（避免强制输出分析结论）');
                }
                // 误判成 beautifier/formatter/html_report → 拉回 text-pipeline（在聊天里渲染结果）
                // html_report 只在用户确实要 html/网页/浏览器展示时才放行；"用表格/json展示结果"应留在聊天。
                if (llmIntent && llmIntent.route && llmIntent.route.pipeline === 'skill') {
                    const _sid = String(llmIntent.route.skillId || '').toLowerCase();
                    const _hasHtmlAsk = /(html|网页|浏览器|大屏|web|报表)/i.test(_msgStr);
                    const _isBeautifyLike = (_sid === 'beautifier' || _sid === 'formatter');
                    const _isHtmlReport = (_sid === 'html_report' || _sid === 'htmlbrowserdisplay') && !_hasHtmlAsk;
                    if (_isBeautifyLike || _isHtmlReport) {
                        ChatLog.info?.('ORCH', '[orchestrator] 展示结果意图 → 强制 text-pipeline（结果在聊天中渲染），忽略 LLM 误路由: ' + _sid);
                        llmIntent.route = {
                            pipeline: 'text',
                            skillId: 'free_chat',
                            confidence: llmIntent.route.confidence || 0.8,
                            reason: '展示结果诉求=在聊天中渲染结果，非美化/报表生成'
                        };
                    }
                }
            }
            const _errorCheckIntent = /(检查|检测|排查|查找|看看|查看)/.test(_msgStr) && /(错误|出错|error|#DIV|#N\/A|#VALUE|#REF)/i.test(_msgStr);
            if (_errorCheckIntent && llmIntent && llmIntent.route && llmIntent.route.pipeline === 'skill') {
                ChatLog.info?.('ORCH', `[orchestrator] 错误检查意图 → 强制 text-pipeline（excel_check_errors），忽略 LLM skill 路由: ${llmIntent.route.skillId}`);
                llmIntent.route = {
                    pipeline: 'text',
                    skillId: 'free_chat',
                    confidence: llmIntent.route.confidence || 0.8,
                    reason: '错误检查走 excel_check_errors 工具精确定位错误单元格'
                };
                if (Array.isArray(llmIntent.operations) && !llmIntent.operations.includes('analyze')) {
                    llmIntent.operations.push('analyze');
                }
            }
            // ════════════════════════════════════════════════════════════════
            // 【快捷入口直连钉死】用户点 ⚡ 快捷入口（或技能卡片"直接执行"）发起时，
            // 技能是用户**显式选择**的，不允许再被任何意图路由改判。
            // 背景（实测 2026-10-02 23:37）：快捷入口原先只把提示词填进输入框、丢掉 skillName，
            // 于是这条消息走通用路由 —— "为火车票创建一个解析器"（GenerateOCRParseScript 快捷
            // 入口的提示词）被 skill_SkillCreator 抢走执行，用户："我都直接用的快捷键调用的
            // skill 也能匹配到别的技能？"
            // 钉死后：matchedSkill = 该技能 + __forceSkillPipeline=true，跳过 regex 路由与
            // LLM 覆盖（后面的 code/python/PPT 强制分支也已用 _pinnedSkill 放行保护）。
            // ════════════════════════════════════════════════════════════════
            const _forcedSkillName = context?._forcedSkillName ? String(context._forcedSkillName) : '';
            if (_forcedSkillName) {
                const _allSkills = Array.isArray(App.cachedSkills) ? App.cachedSkills : [];
                const _wantLower = _forcedSkillName.toLowerCase();
                const _pinned = _allSkills.find(s => String(s.name || s.Name || '').toLowerCase() === _wantLower)
                    || _allSkills.find(s => String(s.displayName || s.DisplayName || '').toLowerCase() === _wantLower);
                if (_pinned) {
                    matchedSkill = _pinned;
                    context = context || {};
                    context.__forceSkillPipeline = true;
                    context._pinnedSkill = String(_pinned.name || _pinned.Name || '');
                    // 快捷入口 = 一次全新的显式调用，不该被上一轮的 follow-up 继承劫持
                    // （继承了会让 dispatcher 走"继续上一轮"的执行分支，而不是这次选的技能）。
                    if (inheritedFollowUp || context._execFollowUp) {
                        ChatLog.info?.('ORCH', '[直连] 清除上一轮 follow-up 继承（快捷入口视为全新调用）');
                    }
                    inheritedFollowUp = false;
                    context._execFollowUp = null;
                    ChatLog.info?.('ORCH', `[直连] 快捷入口锁定技能 ${context._pinnedSkill}（source=${context._forcedSkillSource || 'unknown'}）→ 跳过 regex 路由与 LLM 覆盖`);
                } else {
                    ChatLog.warn?.('ORCH', `[直连] 快捷入口指定技能 "${_forcedSkillName}" 不在 cachedSkills（已停用/已删除？）→ 回退通用路由`);
                }
            }

            // 强制关键词命中放行：即使 _forceMultiStep=true（上游 needsPlanning 误判），
            // 只要消息里命中了 cleaner/crud/beautifier/formatter 这类确定性 Skill 域的关键词，
            // 也照样跑 shouldRouteToSkill —— 不能让一条误判的 needsPlanning 把整条 Skill 路由链掐死。
            // chart/dashboard 域故意不加：它们必须走 text-pipeline 多步骤引擎（aggregate→create_chart 链），
            // spec 模式拿不到 planAndRender 只能对原始明细建废图。
            const _sdNow = App.ChatV2?.skillDomain;
            const _kwHitDomain = (() => {
                if (!_sdNow || typeof _sdNow.matchDomains !== 'function') return [];
                const hits = _sdNow.matchDomains(message || '');
                return hits.filter(id => /cleaner|crud|beautifier|formatter/i.test(id));
            })();
            const _forceRegexRoute = _kwHitDomain.length > 0;
            if (!inheritedFollowUp && (_forceRegexRoute || !context?._forceMultiStep) && !context?._pinnedSkill) {
                matchedSkill = (kind === 'text') ? shouldRouteToSkill(message, context) : null;
                ChatLog.info?.('ORCH', `[orchestrator] regex 路由: msg="${(message||'').slice(0,40)}" → ${matchedSkill ? (matchedSkill.name || matchedSkill.Name) : 'null'}`);

                // ── LLM route 覆盖 regex ──
                const _sd = App.ChatV2?.skillDomain;
                if (llmIntent && llmIntent.route && llmIntent.route.confidence >= ROUTE_T.LLM_ROUTE_MIN_CONFIDENCE) {
                    const _route = llmIntent.route;
                    if (_route.pipeline === 'skill' && _route.skillId && _route.skillId !== 'free_chat') {
                        // LLM 明确说走 skill-pipeline + 指定 skillId → 按 skillId 从 DESCRIPTORS 找对应 skill。
                        //
                        // ⚠️ 领域 id 而非真实 SkillSpec 的空壳不能直接作为 matchedSkill：
                        // 域名（如 'chart'）只有 name+keywords，缺 processingMode/executionMode/steps，
                        // skill-pipeline resolveExecutionMode 匹配不到 SmartChartWizard 等正则 → 判 fixed →
                        // LLM 自主挑工具（agent）会编造 "skill.chart" → tool-registry 兜底错译成 ppt_add_chart_slide
                        // → 后端报 "MCP Client 'ppt-local' 不存在"，图表请求直接失败（实测"使用饼图按部门展示销售额"）。
                        // 正确做法：命中确定性域（chart/beautifier/cleaner/crud/dashboard/formatter）时，
                        // 用该域 DESCRIPTOR 的 names 正则去 cachedSkills 反查真实 SkillSpec（SmartChartWizard 等），
                        // 走 spec 模式（excel_create_chart 图表链）；查不到真实 spec 才回退域空壳。
                        let _llmSkill = null;
                        if (_sd) {
                            const _descs = typeof _sd.getAll === 'function' ? _sd.getAll() : [];
                            const found = _descs.find(d => d.id === _route.skillId);
                            if (found) {
                                // 反查真实 SkillSpec：域名 names 正则（如 /SmartChartWizard|Chart/i）匹配真实 skill 名
                                const _realSkills = Array.isArray(App.cachedSkills) ? App.cachedSkills : [];
                                const _real = _realSkills.find(s => {
                                    const _n = s.name || s.Name || '';
                                    return found.names.test(_n);
                                });
                                if (_real) {
                                    _llmSkill = _real;
                                    if ((_route.skillId === 'chart' || _route.skillId === 'beautifier' ||
                                        _route.skillId === 'cleaner' || _route.skillId === 'crud' ||
                                        _route.skillId === 'dashboard' || _route.skillId === 'formatter') &&
                                        !String(_real.processingMode ?? _real.ProcessingMode ?? '').match(/skillspec/i) &&
                                        Number(_real.processingMode ?? _real.ProcessingMode) !== 5) {
                                        ChatLog.info?.('ORCH',
                                            `[orchestrator] LLM 领域 "${_route.skillId}" → 真实 SkillSpec: ${(typeof _real === 'object' ? (_real.name || _real.Name) : _real)}（走 spec 模式确定性图表/格式/清洗链）`);
                                    }
                                } else {
                                    _llmSkill = { name: found.id, keywords: found.keywords || [], __fromDescriptor: true };
                                }
                            }
                        }
                        // 【别名补解析】skill_creator/html_report 这类"目录展示别名"既不在 DESCRIPTORS
                        // 也不是真实技能名，旧逻辑直接落到下面的"没找到 → 保留 regex"，LLM 的选择被静默丢弃
                        // （而意图块却还写着它 → 用户以为选错了技能）。这里复用显示层同一个
                        // resolveLlmRouteTarget：别名必须过该技能自己的意图守卫才认，
                        // 所以"为火车票创建一个解析器"不会被 SkillCreator 抢走，只会显式记一条日志。
                        if (!_llmSkill) {
                            const _alias = resolveLlmRouteTarget(_route, message);
                            if (_alias.ok && _alias.skill) {
                                _llmSkill = _alias.skill;
                                ChatLog.info?.('ORCH', `[orchestrator] LLM skillId="${_route.skillId}" 按别名解析 → ${_alias.skill.name || _alias.skill.Name}（${_alias.reason}）`);
                            }
                        }
                        if (_llmSkill) {
                            if (!matchedSkill || (_llmSkill.name || _llmSkill.Name) !== (matchedSkill.name || matchedSkill.Name)) {
                                ChatLog.info?.('ORCH', `[orchestrator] LLM route 覆盖 regex: ${matchedSkill ? (matchedSkill.name||matchedSkill.Name)+'→' : 'null→'}${_llmSkill.name} (confidence=${_route.confidence}, reason=${_route.reason})`);
                            }
                            matchedSkill = _llmSkill;
                        } else {
                            ChatLog.warn?.('ORCH', `[orchestrator] LLM 指定 skillId="${_route.skillId}" 但 DESCRIPTORS/cachedSkills 里没找到 → 保留 regex 结果`);
                        }
                    } else if (_route.pipeline === 'text') {
                        // 创建意图守卫：消息含明确"做/生成/创建…看板"等创建动词且 regex 已命中看板类 Skill 时，
                        // LLM 的 text 路由不得否决——否则执行器会在无工具调用下直接编造"已生成"的幻觉完成文本。
                        const _strongCreate = matchedSkill && /(做|生成|创建|制作|搞|弄|画|出|排|搭)\s*(个|一个|一张|份|下)?\s*(excel\s*|表格\s*)?(看板|驾驶舱|仪表盘|大屏|数据面板|dashboard)/i.test(message || '');
                        if (_strongCreate) {
                            ChatLog.warn?.('ORCH', `[orchestrator] LLM 判 text 但消息含明确创建动词("${(message||'').slice(0,30)}") → 保留 regex 命中 ${matchedSkill.name || matchedSkill.Name}，防幻觉完成`);
                        } else if (matchedSkill) {
                            // ───────────────────────────────────────────────────────────────────
                            // 轻量确认：宁可多一次确认，不可错误截断
                            // LLM 判 text-pipeline，但 regex（shouldRouteToSkill）有明确 Skill 命中 ——
                            // 不再直接 matchedSkill = null 截断 Skill，而是挂起 dispatch、把选择弹给用户。
                            // 超时 2min 或 turn 被 abort → 保守默认 keep（保留 Skill 意图）。
                            // ───────────────────────────────────────────────────────────────────
                            const _skillName = matchedSkill.displayName || matchedSkill.name || matchedSkill.Name || '技能';
                            const _skillDomain = (App.ChatV2?.skillDomain?.detectTypeByName?.(matchedSkill.name || matchedSkill.Name)) || 'skill';
                            const _skillHint = ({
                                chart: '生成图表', beautifier: '美化表格', formatter: '格式化',
                                cleaner: '清洗数据', crud: '增删改查', dashboard: '生成看板'
                            })[_skillDomain] || _skillName;

                            ChatLog.info?.('ORCH', `[confirm] LLM text vs regex Skill(${_skillHint}) 冲突 → 挂起等待用户选择 (confidence=${_route.confidence}, reason=${_route.reason})`);
                            sessionStore.setTurnStatus(turn.id, '🤔 需要确认一下你的意图');
                            sessionStore.updateTurn(turn.id, {
                                _pendingSkillConfirm: {
                                    skill: { name: matchedSkill.name || matchedSkill.Name, domain: _skillDomain, hint: _skillHint },
                                    llmReason: _route.reason || '',
                                    llmConfidence: _route.confidence,
                                    question: `系统识别到你可能想【${_skillHint}】，但也可能只是想看分析。你希望？`
                                }
                            });

                            // 挂起等待用户选择：Promise + Abort + Timeout 三重保护
                            let confirmResult = { choice: 'keep', timeout: false };
                            try {
                                confirmResult = await new Promise((resolve) => {
                                    _confirmResolvers.set(turn.id, resolve);
                                    // Timeout → 默认 keep（保守优先）
                                    const _to = setTimeout(() => {
                                        if (_confirmResolvers.get(turn.id) === resolve) {
                                            _confirmResolvers.delete(turn.id);
                                            ChatLog.info?.('ORCH', `[confirm] turn=${turn.id} 超时 ${ROUTE_CONFIRM_TIMEOUT_MS}ms，保守默认 keep`);
                                            resolve({ choice: 'keep', timeout: true });
                                        }
                                    }, ROUTE_CONFIRM_TIMEOUT_MS);
                                    // Abort → 清理 resolver，choice='aborted' 让上层 return
                                    const onAbort = () => {
                                        clearTimeout(_to);
                                        if (_confirmResolvers.get(turn.id) === resolve) {
                                            _confirmResolvers.delete(turn.id);
                                            resolve({ choice: 'aborted' });
                                        }
                                    };
                                    controller.signal.addEventListener('abort', onAbort, { once: true });
                                });
                            } catch (_) { /* 异常兜底：keep */ }

                            if (confirmResult.choice === 'aborted') {
                                // 用户点取消 / 发新消息 → abort 已触发，直接 return 让 cancel 流程接管
                                ChatLog.info?.('ORCH', `[confirm] turn=${turn.id} 等待中被 abort → return`);
                                return turn.id;
                            }

                            // 清除 pending 标记（让渲染器移除确认卡片）
                            sessionStore.updateTurn(turn.id, { _pendingSkillConfirm: null });

                            if (confirmResult.choice === 'drop') {
                                // 用户明确要 text-pipeline → 尊重选择
                                ChatLog.info?.('ORCH', `[confirm] 用户选 text-pipeline: 放弃 ${matchedSkill.name || matchedSkill.Name} (timeout=${confirmResult.timeout})`);
                                matchedSkill = null;
                            } else {
                                // keep 或 超时兜底 → 保留 regex 命中
                                ChatLog.info?.('ORCH', `[confirm] 用户选 keep: 维持 ${matchedSkill.name || matchedSkill.Name} (choice=${confirmResult.choice}, timeout=${confirmResult.timeout})`);
                            }
                        } else {
                            // matchedSkill 为 null（regex 本身没命中），说明 LLM text 路由没有 Skill 被截——正常 drop
                            ChatLog.info?.('ORCH', `[orchestrator] LLM route: regex 本身 null → text-pipeline (confidence=${_route.confidence}, reason=${_route.reason})`);
                        }
                    }
                } else if (llmIntent && llmIntent.route && llmIntent.route.confidence < 0.5) {
                    ChatLog.info?.('ORCH', `[orchestrator] LLM route confidence=${llmIntent.route.confidence} < 0.5 → 保留 regex 结果`);
                }
            } else {
                ChatLog.info?.('ORCH', `[orchestrator] 路由被跳过: inheritedFollowUp=${inheritedFollowUp} _forceMultiStep=${context?._forceMultiStep}`);
            }

            // 确定性 html 意图兜底：仅当 isDeterministicHtmlReportIntent 预判**真·HTML 报表意图**
            // （生成动作 + html/网页关键词同时命中）且 shouldRouteToSkill 返回 null 时才塞 HtmlBrowserDisplay。
            // 兜底进来的 skill 也必须能通过下面的 Route Validator——如果用户说"做个看板"
            // （wantsDashboard=true），HtmlBrowserDisplay 的 pureDisplayOnly 标记会被 validator 拦下。
            const _htmlIntentOnly = isDeterministicHtmlReportIntent();
            if (_htmlIntentOnly && !matchedSkill) {
                const htmlSk = (App.cachedSkills || []).find(s => String(s.pipeline || s.Pipeline).toLowerCase() === 'html');
                if (htmlSk) {
                    matchedSkill = htmlSk;
                    ChatLog.info?.('ORCH', `[orchestrator] 确定性 html 意图兜底：matchedSkill=null → 采用缓存 pipeline=html 技能 ${htmlSk.name}`);
                } else {
                    ChatLog.warn?.('ORCH', '[orchestrator] 确定性 html 意图但缓存中无 pipeline=html 技能');
                }
            }

            // ============================================================
            // Route Validator：Capability 校验（核心防线）
            // ------------------------------------------------------------
            // 不管 matchedSkill 来自 shouldRouteToSkill 还是 html 兜底，
            // 都必须过这层：skill 自声明的 capabilities 能不能覆盖用户消息里的
            // 意图需求。意图需求优先从 LLM operations 提取（LLM 动态理解，类似
            // 杰瑞VBE 的"AI 动态生成+修正"思路，让路由层也具备语义理解能力），
            // LLM 超时/失败时自动降级到正则兜底。
            //
            // 这比堆 block 补丁强得多——不需要枚举"看板|仪表盘|驾驶舱|..."，
            // 能力边界一眼看清，新增 skill 自动生效。
            // ============================================================
            // 元技能（ROUTING_POLICIES 里声明了独占 intent 的生成器，如 SkillCreator）必须在能力校验前识别：
            // 它不直接操作 Excel，产物是"一个新技能"，拿消息里的 read/write 操作词去校验它属于范畴错误。
            // 实测（2026-10-02 21:59）："生成一个批量发送邮件的Skill"被 LLM 抽成 operations=[read,analyze,write]
            // → needsEditExcel=true → SkillCreator 自声明无 canEditExcel → 被驳回 → matchedSkill=null
            // → 落回 text-pipeline 的 agent-loop → skill_card 丢失，用户只看到"已成功创建…"的总结文本，
            //   看不到卡片、无法审查、无法保存。
            const _metaNamePre = matchedSkill ? (matchedSkill.name || matchedSkill.Name || '') : '';
            const _sdPre = App.ChatV2?.skillDomain;
            const _metaPolicyPre = (_sdPre && _sdPre.ROUTING_POLICIES && _metaNamePre)
                ? (_sdPre.ROUTING_POLICIES[_metaNamePre] || _sdPre.ROUTING_POLICIES[_metaNamePre.toLowerCase()])
                : null;
            const isMetaSkillPre = !!(_metaPolicyPre && _metaPolicyPre.intent);

            if (matchedSkill && App.ChatV2?.skillDomain?.validateSkillRoute) {
                if (context?._pinnedSkill) {
                    // 快捷入口直连：用户显式选的技能不拿 capability 去驳（它可能本就不直接操作 Excel，
                    // 如 OCR 解析脚本生成器）；驳回等于把用户的选择抹掉，回到"选错技能"的老问题。
                    ChatLog.info?.('ORCH', `[RouteValidator] 快捷入口直连技能(${_metaNamePre}) 跳过 capability 校验（用户显式选择优先于自动校验）`);
                } else if (isMetaSkillPre) {
                    ChatLog.info?.('ORCH', `[RouteValidator] 元技能(${_metaNamePre}) 跳过 capability 校验（生成器型：产物是新技能，不直接操作 Excel）`);
                } else {
                    const _ops = (llmIntent && Array.isArray(llmIntent.operations)) ? llmIntent.operations : null;
                    const _v = App.ChatV2.skillDomain.validateSkillRoute(matchedSkill, message, _ops);
                    if (!_v.ok) {
                        ChatLog.info?.('ORCH', `[RouteValidator] ${matchedSkill.name || matchedSkill.Name} 被 capability 校验驳回: ${_v.reason} → 丢弃，交给 text-pipeline`);
                        matchedSkill = null;
                    } else {
                        ChatLog.info?.('ORCH', `[RouteValidator] ${matchedSkill.name || matchedSkill.Name} 通过能力校验 ✓ (intentSource=${_ops ? 'llm-operations' : 'regex-fallback'})`);
                    }
                }
            }

            // 【降级守卫 C】matchedSkill 若指向一个"当前不可用"的 Skill（已被停用 / 仅 DESCRIPTORS 空壳），
            // 一律置空 → 后续 isSkDomainName=false 不再置 __forceSkillPipeline、kind 保持 'text'，
            // 自然回落通用 text-pipeline，避免把请求塞进不存在的 Skill 管道后失败或静默无输出。
            // 此处是 matchedSkill 的唯一收敛点：regex 匹配、LLM 覆盖的域空壳、DESCRIPTORS 合成空壳三条来源均已就位。
            if (matchedSkill && !isRealAvailableSkill(matchedSkill)) {
                // 若该不可用 Skill 属于 chart 域（真实 SmartChartWizard 已停用 / 仅 chart 域空壳），
                // 则"无真实 Skill"的图表意图必须显式强制走 text-pipeline 多步骤通道——
                // 不依赖守卫 D 的强图表引导词命中，也不依赖 LLM 单/多步判定：
                // _forceMultiStep 让 decideMultiStep 第一优先级短路返回 true，稳定进入
                // runMultiStepPipeline，由真实工具链（aggregate→write→create_chart）出图。
                // 否则 chart 域空壳被丢弃后若消息缺强图表引导词（如"做成图表"），会退回单步 agent-loop 不出图。
                const _dropName = matchedSkill.name || matchedSkill.Name || '';
                const _dropIsChart = /SmartChartWizard|Chart|图表/i.test(_dropName) ||
                    !!(App.ChatV2?.skillDomain?.isDomainName && App.ChatV2.skillDomain.isDomainName(_dropName));
                if (_dropIsChart) {
                    context = context || {};
                    context._forceChart = true;
                    context._forceMultiStep = true;
                    ChatLog.info?.('ORCH', `[orchestrator] 守卫 C: 丢弃的不可用 Skill(${_dropName}) 属 chart 域 → 显式 _forceChart + _forceMultiStep，强制 text-pipeline 多步骤通道`);
                }
                ChatLog.info?.('ORCH', `[orchestrator] Skill(${_dropName}) 当前不可用（已停用/仅域空壳）→ 丢弃，回落 text-pipeline`);
                matchedSkill = null;
            }

            // 【补救守卫 D】matchedSkill 被置空后，检查消息有没有强图表意图词。
            // 场景：LLM route 把"按销售额汇总出饼图"误判成 crud（"汇总"触发 CRUD），
            // crud skill 不可用 → 被守卫 C 丢弃 → 回落 text-pipeline → LLM 只调 aggregate 不调 create_chart。
            // 如果有明确图表类型词（饼图/柱图/折线图/可视化/仪表盘...），说明用户意图是出图，
            // 重找 chart skill；找不到也在 context 上打 _forceChart 标记让 text-pipeline 的 LLM 知道要出图。
            const STRONG_CHART_INTENT_RE = routing().STRONG_CHART_INTENT_RE;
            if (!matchedSkill && STRONG_CHART_INTENT_RE.test(message || '')) {
                // 尝试重找 chart skill（从 cachedSkills 和 DESCRIPTORS 反查）
                let _chartSkill = null;
                const _realSkills = Array.isArray(App.cachedSkills) ? App.cachedSkills : [];
                _chartSkill = _realSkills.find(s => {
                    const _n = s.name || s.Name || '';
                    return /SmartChartWizard|Chart|图表/i.test(_n);
                });
                if (!_chartSkill && App.ChatV2?.skillDomain) {
                    const _descs = typeof App.ChatV2.skillDomain.getAll === 'function' ? App.ChatV2.skillDomain.getAll() : [];
                    const chartDesc = _descs.find(d => d.id === 'chart');
                    if (chartDesc) {
                        _chartSkill = _realSkills.find(s => {
                            const _n = s.name || s.Name || '';
                            return chartDesc.names?.test(_n) || /Chart/i.test(_n);
                        });
                    }
                }
                if (_chartSkill && isRealAvailableSkill(_chartSkill)) {
                    ChatLog.info?.('ORCH', `[orchestrator] 补救守卫 D: 检测到强图表意图 → 重路由到 chart skill (${_chartSkill.name || _chartSkill.Name})`);
                    matchedSkill = _chartSkill;
                } else {
                    // chart skill 不可用（已停用 / 全部 Skill 关闭 → 仅 DESCRIPTORS 空壳）→
                    // 显式强制走 text-pipeline 多步骤通道，不再只靠 _forceChart 的 prompt 引导词命中。
                    //
                    // 为什么还要 _forceMultiStep：text-pipeline.decideMultiStep 的优先级是
                    //   _forceMultiStep → _intentSingleOp → TaskPlanner.needPlanning
                    // _forceChart 只往 prompt 注入"记得调 excel_create_chart"的引导，弱模型可能读了也不调，
                    // 仍退回单步 agent-loop（光说不做）；而 _forceMultiStep 让 decideMultiStep 第一优先级
                    // 短路返回 true，100% 进入 runMultiStepPipeline（agent-loop 循环真实工具链：
                    // excel_aggregate / excel_read_range → excel_write_range → excel_create_chart），
                    // 且优先级高于 _intentSingleOp，图类请求被 LLM 误判"单操作"时也不会被降级退化。
                    //
                    // 这正是"无 Skill 通用通道"对图表意图的确定性兜底：无论引导词命中与否、无论 LLM 单/多步判定，
                    // 强图表意图都稳定落到多步骤通道，由真实工具链执行（杜绝"生成了一个没用图"的空壳单次错图）。
                    ChatLog.info?.('ORCH', '[orchestrator] 补救守卫 D: 强图表意图但 chart skill 不可用 → 显式 _forceChart + _forceMultiStep，强制 text-pipeline 多步骤通道');
                    context = context || {};
                    context._forceChart = true;       // 保留：prompt-builder 注入"先 aggregate 再 create_chart"引导
                    context._forceMultiStep = true;   // 新增：决定层短路进入多步骤管道，不依赖引导词/LLM 判定
                }
            }

            // 【补救守卫 E】看板/仪表盘意图。
            // DashboardGenerator 在后端可能不存在或不可用 → matchedSkill 被置空 → 回落 text-pipeline →
            // LLM 只输出文字分析（列出"应该用哪些图表"的建议）但不调任何原子工具。
            // 此处检测看板意图，设:
            //   _forceDashboard=true  → prompt-builder 注入看板强引导（5 步工具链）
            //   _forceChart=true      → 同时注入图表生成引导
            //   _forceMultiStep=true  → 强制进入 runMultiStepPipeline，agentLoop.run 带 persistTools=true
            //                            + maxIterations=MAX_TOOL_ROUNDS（默认 5），确保工具失败不会锁死、LLM 有足够轮次串联多工具
            // 为什么三个标记都要：看板 = prompt 引导 + 多步骤管道 + 图表生成，缺一个都会断链。
            const STRONG_DASHBOARD_INTENT_RE = /看板|仪表盘|驾驶舱|dashboard|数据大屏|大屏|数据面板|经营看板|销售看板|监控大屏|仪表板/i;
            if (!matchedSkill && !context?._forceDashboard && STRONG_DASHBOARD_INTENT_RE.test(message || '')) {
                ChatLog.info?.('ORCH', '[orchestrator] 补救守卫 E: 检测到看板/仪表盘意图 → 三标记（_forceDashboard + _forceChart + _forceMultiStep）');
                context = context || {};
                context._forceDashboard = true;
                context._forceChart = true;
                context._forceMultiStep = true;  // 关键：让 decideMultiStep 直接 true → 进入多步骤管道（persistTools）
            }

                // 【多步骤请求拦截】如果消息包含"然后/再/接着"等顺序词且涉及多个操作类别，
                // 即使有 Skill 匹配也不应路由到 skill-pipeline，而应交由 text-pipeline 的多步骤管道处理。
                // 例如"读取 Sheet1 的数据，然后按部门汇总"包含"读取"+"汇总"两个操作，
                // 如果路由到"Excel 数据增删改查"Skill，只会执行"读取"而丢失"汇总"。
                // 例外：LLM 意图分析已判定为单操作/问答（_intentSingleOp=true）时，
                // 关键词计数（needPlanning===true）不再拦截——LLM 语义理解是权威，关键词只是兜底。
                // 额外例外：Beautifier/Formatter/Chart/Cleaner/CRUD 这些"确定性 Skill"本身就是
                // 多步骤管道（内部 planAndRender + VBA/多工具链），不应被 needPlanning 拦截丢进
                // 通用 text-pipeline MultiStepEngine（后者只会调 excel_format_range 一个 MCP）。
                // pipeline='html' 的技能是确定性"网页报表生成"（前端 html-display-pipeline 真生成本地 html 文件），
                // 与 Beautifier/Chart 等同属确定性 Skill：不被需多步骤的 needPlanning 拦截、强制走确定性链路。
                const skillHTMLPipe = String(matchedSkill?.pipeline || matchedSkill?.Pipeline || '').toLowerCase() === 'html';
                // 确定性 Skill 判定统一走声明式注册表 config/skill-domain.js（isDomainName 派生自 names 正则），
                // 不再本地私写 beautifier|formatter|chartwizard|... 的正则副本（避免修一处漏一处漂移）。
                const skillDomain = App.ChatV2?.skillDomain;
                const isSkDomainName = !!(skillDomain?.isDomainName) && skillDomain.isDomainName(matchedSkill?.name || matchedSkill?.Name || '');
                // ⚠️ 合成空壳不是"确定性 Skill"——它是 shouldRouteToSkill / LLM 路由在 cachedSkills
                // 找不到真实 Skill 时注入的 DESCRIPTORS 壳（__fromDescriptor，name 即域 id，只有 keywords，没有 steps）。
                // 真实 SmartChartWizard 等 Skill 内部才有确定性多工具链（planAndRender + 聚合 + 建图）；
                // 空壳进 skill-pipeline 后拿不到任何 steps，只能执行**单次**工具调用：
                // 实测"读取当前选中数据，按部门汇总销售额，画柱状图，把汇总结果写回新表"——
                //   空壳 chart 享受确定性豁免 → `__forceSkillPipeline` 把它强制钉在 skill-pipeline
                //   → 只调 excel_create_chart(dataRange=$A$1:$I$101) 对原始明细建图（图例是一堆
                //     "企业客户/个人客户"原始行，完全没按部门聚合，用户反馈"生成了一个没用图"）。
                // 因此合成空壳不享受该豁免：交回下面的 needPlanning 判定，由多步骤引擎用该域的真实
                // 工具链（excel_aggregate → excel_write_range → excel_create_chart）执行。
                const _matchedName = matchedSkill?.name || matchedSkill?.Name || '';
                const isSyntheticShell = !!(matchedSkill && matchedSkill.__fromDescriptor);
                // 合成空壳豁免细分：
                //   - chart/dashboard 空壳 → 不豁免（必须进多步骤引擎调 aggregate→create_chart 工具链，
                //     spec 模式拿不到 planAndRender 只能对原始明细建废图）
                //   - cleaner/crud/beautifier/formatter 空壳 → 豁免（spec 模式已有独立原子工具链：
                //     executeStructureTransform 直调 split/merge，不需要多步骤引擎脑补）
                const _shellNeedsMultiStep = isSyntheticShell && /chart|dashboard/i.test(_matchedName);
                // 元技能豁免：SkillCreator 这类"生成器"Skill 不在 DESCRIPTORS 域表内，
                // 必须独占 skill-pipeline（由后端固定通道执行，不交给 LLM 自由发挥）。
                // isMetaSkillPre 已在 RouteValidator 之前完成元技能判定，此处复用避免两处判据漂移。
                const isMetaRoutingSkill = isMetaSkillPre && !!matchedSkill;
                const isDeterministicSkill = !_shellNeedsMultiStep && (skillHTMLPipe || isSkDomainName || isMetaRoutingSkill);

                // ── [多意图观测哨] (Step 1) 感知复合意图，不改路由行为，只打日志观察频度 ──
                // matchDomains 返回消息命中的全部确定性 Skill 域 id；>=2 即"美化后再画图/清洗并生成看板"类复合句。
                // 刻意不劫走路由：确定性 Skill 的豁免（isDeterministicSkill）保持不变，避免重蹈
                // "被多步劫持后只调一个 MCP / 美化表格掉进 text-pipeline 幻觉"的历史血案。
                // 目的仅是采集真实复合意图占比，为 Step 2（多 skill 顺序执行）的排期提供数据。
                let multiIntentDomains = [];
                try {
                    if (skillDomain && typeof skillDomain.matchDomains === 'function') {
                        multiIntentDomains = skillDomain.matchDomains(message || '') || [];
                    }
                } catch (me) { /* 观测哨失败不影响主流程 */ }
                if (multiIntentDomains.length >= 2) {
                    const _msn = matchedSkill?.name || matchedSkill?.Name || '无';
                    ChatLog.info?.('ORCH',
                        `[多意图] 命中${multiIntentDomains.length}个域: ${multiIntentDomains.join('+')} | ` +
                        `matchedSkill=${_msn} deterministic=${isDeterministicSkill} intentSingleOp=${!!context?._intentSingleOp} | msg="${(message || '').slice(0, 40)}"`);
                }
                if (isDeterministicSkill) {
                    context = context || {};
                    context.__forceSkillPipeline = true;
                    ChatLog.info?.('ORCH', `[orchestrator] ${isMetaRoutingSkill ? '元技能' : '确定性 Skill'}(${matchedSkill.name || matchedSkill.Name}) → 强制走 skill-pipeline（标记 __forceSkillPipeline=true）`);
                }
                // 合成空壳额外豁免 _intentSingleOp：空壳没有任何确定性链可用，即使 LLM 把图类请求判成
                // "单操作"（图类请求极易被判单操作），保留在 skill-pipeline 的后果也只有一张错图。
                // 而 needPlanning===true 仅在"≥2 个工具类别"或"图表词+分组维度词"这类硬结构信号下成立，足够可靠。
                if (matchedSkill && kind === 'text' && (!context?._intentSingleOp || isSyntheticShell) && !isDeterministicSkill && !context?._pinnedSkill) {
                    try {
                        const taskPlanner = App.ChatV2.TaskPlanner;
                        if (taskPlanner && typeof taskPlanner.needPlanning === 'function') {
                            const planCheck = taskPlanner.needPlanning(message || '');
                            if (planCheck === true) {
                                ChatLog.info?.('ORCH', `[orchestrator] 检测到多步骤请求(类别>=2)，跳过 Skill 路由: ${matchedSkill.name || matchedSkill.Name} → text-pipeline 多步骤管道`);
                                matchedSkill = null;
                                context = context || {};
                                context._skipMultiStep = true; // 标记：此请求已判定为多步骤，后续 CRUD 兜底也不应拦截
                                context._enableSteps = true;   // 告诉渲染层：需要显示步骤面板和工具卡片
                                // 合成空壳覆盖了 LLM 的"单操作"判定（图类请求极易被判单操作）时，必须同时清掉
                                // _intentSingleOp：text-pipeline.decideMultiStep 第 2 步会据此直接 return false
                                // → 退化成 agent-loop 兜底（可能光说不做）。needPlanning===true 是硬结构信号，以此为准。
                                if (isSyntheticShell) context._intentSingleOp = false;
                            }
                        }
                    } catch (planErr) {
                        ChatLog.warn?.('ORCH', '[orchestrator] 多步骤检测失败（不影响主流程）:', planErr?.message);
                        _orchHostLog('WARN', `[orchestrator] 多步骤检测失败（不影响主流程）: ${planErr?.message}`);
                    }
                } else if (isDeterministicSkill) {
                    ChatLog.info?.('ORCH', `[orchestrator] 确定性 Skill(${matchedSkill.name || matchedSkill.Name}) 不拦截多步骤判定 → 保持 skill-pipeline 路由`);
                }
            // 【CRUD 兜底路由】未命中任何 Skill，但用户意图是明确的增删改查（读/写/插入/删除/清空）
            // → 走 ExcelCRUD Skill 的确定性工具链，而不是让通用对话直接调 excel_* CRUD 裸 MCP 工具。
            // 仅当没有任何 Skill 命中时才兜底，避免覆盖去重/清洗/美化/图表/分析等专属 Skill；
            // crudSpecFromMessage 已排除去重/清洗/美化/图表/分析/排序等其它 Skill 专属意图，双保险。
            // 多步骤请求（_skipMultiStep=true / LLM 意图分析 _forceMultiStep=true）也不应走 CRUD 兜底。
            if (!inheritedFollowUp && !matchedSkill && kind === 'text' && !context?._skipMultiStep && !context?._forceMultiStep) {
                const specEngine = App.ChatV2.services && App.ChatV2.services.skillSpecEngine;
                if (specEngine && typeof specEngine.crudSpecFromMessage === 'function' &&
                    specEngine.crudSpecFromMessage(message || '')) {
                    const crudSkill = (App.cachedSkills || App.ln || []).find(
                        s => (s.name === 'ExcelCRUD' || s.Name === 'ExcelCRUD' ||
                              s.displayName === 'Excel 数据增删改查' || s.DisplayName === 'Excel 数据增删改查'));
                    if (crudSkill) {
                        matchedSkill = crudSkill;
                        ChatLog.info?.('ORCH', `[orchestrator] 检测到增删改查意图，路由到 ExcelCRUD Skill（确定性工具链）`);
                    }
                }
            }
            // 【内置确定性任务判定（与 Skill 无关）】该判定必须放在 `if (matchedSkill)` 之外：
            // 全部 Skill 停用时 matchedSkill 必为 null，仍需识别"多图表/聚合"这类请求，
            // 用于下面的"多图表强制走 text-pipeline"决策。
            // 注意：原先把命中的请求强制交给 text-pipeline 的内置编排（runMultiChartPipeline），
            // 该内置编排已按用户要求删除，这里不再设 context.__forceTextPipeline。
            let builtinIntent = null;
            if (kind === 'text') {
                try {
                    builtinIntent = App.ChatV2.taskPlanner.detectBuiltinIntent(message || '');
                } catch (biErr) {
                    ChatLog.warn?.('ORCH', '[orchestrator] 内置确定性任务判定失败（不影响主流程）:', biErr?.message);
                    _orchHostLog('WARN', `[orchestrator] 内置确定性任务判定失败（不影响主流程）: ${biErr?.message}`);
                }
            }
            const forceTextBuiltin = !!(builtinIntent && builtinIntent.builtin);

            if (matchedSkill) {
                // SmartChartWizard 是单 Skill 单任务模式，
                // 用户同时说多个图表任务时（如"用饼图按部门...用柱状图按地区..."），
                // 必须走 text-pipeline 的多任务循环，否则 Skill 只处理第一个图表，后面的丢失。
                // 用局部变量跟踪"多图表强制 text"，替代原来的 `if (kind === 'text')`
                // （该判断永远为 true：matchedSkill 只在 kind==='text' 时匹配，导致所有 Skill 被错误路由到
                //  text-pipeline，LLM 光说不做不执行 Skill）。
                let forceTextPipeline = false;
                const isSmartChartWizard = !!(matchedSkill && (matchedSkill.name === 'SmartChartWizard' || matchedSkill.Name === 'SmartChartWizard'));
                // R2 收敛：内置确定性任务(聚合/多图表/词级多图表)的判定统一走
                // taskPlanner.detectBuiltinIntent（单一实现），已在 `if (matchedSkill)` 之前完成，
                // 这里直接复用上面的 builtinIntent / forceTextBuiltin，不再重复判定。
                // （原 __forceTextPipeline 强制通道已随 text-pipeline 的内置编排一并移除，
                //  此处仅保留"多图表强制走 text-pipeline"的判定。）
                // 多图表：若命中 SmartChartWizard 单 Skill，则必须改走 text-pipeline 多任务循环
                // （否则 Skill 只处理第一个图表，后面的丢失）。
                if (forceTextBuiltin && isSmartChartWizard &&
                    (builtinIntent.multiChartByCount || builtinIntent.multiChartBySplit)) {
                    forceTextPipeline = true;
                    // 多图表必然是多步骤（每张图各需聚合+建图），且这里 _matchedSkill 仍在，
                    // 会挡住 text-pipeline 的多步骤判定 → 显式打标记放行到 MultiStepEngine。
                    if (context) context._forceMultiStep = true;
                    ChatLog.info?.('ORCH', `[orchestrator] 多图表/聚合检测(图表词=${builtinIntent.chartCount}, 拆分任务=${builtinIntent.chartTasks.length}, 聚合=${builtinIntent.hasAggregate})，强制走 text-pipeline + 多步骤引擎，跳过 SmartChartWizard 单 Skill 模式`);
                }

                // 路由到 Skill 声明的 pipeline（默认 skill-pipeline）：
                //   1. 多图表强制 text（forceTextPipeline=true）→ 保留 text-pipeline
                //   2. 代码生成意图 → 统一走 code pipeline（只展示，不执行）
                //   3. 确定性 Skill(__forceSkillPipeline) → 强制 'skill'（覆盖后端 pipeline 声明）
                //   4. 其他 → 按 Skill pipeline 声明路由，默认 'skill'
                if (forceTextPipeline) {
                    kind = 'text';
                } else if (!context?._pinnedSkill && (context?.__forceCodePipeline || routing().detectCodeGenerationIntent(message))) {
                    // 快捷入口直连的技能优先于"代码生成意图"：用户点的是技能卡片，不是"写段代码"。
                    kind = 'code';
                } else if (context?.__forceSkillPipeline) {
                    kind = 'skill';
                    ChatLog.info?.('ORCH', `[orchestrator] __forceSkillPipeline → 强制 kind='skill' (忽略 Skill pipeline 声明 ${matchedSkill.pipeline || matchedSkill.Pipeline || ''})`);
                } else {
                    kind = matchedSkill.pipeline || matchedSkill.Pipeline || 'skill';
                }

                // === PPT 意图优先 ===
                // 用户明确要 PPT（含"趋势分析页/图表页"等 PPT 页面诉求）时，即使分析类 Skill
                // （如 GeneratePython 的"趋势分析"）在评分上反超 GeneratePPTReport，也必须改回
                // GeneratePPTReport → code-pipeline。否则"在PPT汇报中增加趋势分析页"会被路由到
                // python-pipeline 触发 Excel/Python 分析（GeneratePython 声明 pipeline: python）。
                // 例外：明确页面结构/主题要求（pptFreeBuild）时放行 agent-loop——
                // AI 自由调用 ppt_ MCP 工具自由创作，GeneratePPTReport 技能仅是简单生成的扩展兜底。
                if (pptFreeBuild && !context?._pinnedSkill) {
                    if (matchedSkill) {
                        ChatLog.info?.('ORCH', `[orchestrator] PPT 结构诉求: 不锁定 ${matchedSkill.name || matchedSkill.Name} → 放行 agent-loop 自由生成`);
                    }
                    matchedSkill = null;
                    kind = 'text';
                } else if (!context?._pinnedSkill && pptIntent && (matchedSkill.name || matchedSkill.Name) !== 'GeneratePPTReport') {
                    const pptSkill = (App.cachedSkills || []).find(
                        s => s.name === 'GeneratePPTReport' || s.Name === 'GeneratePPTReport');
                    if (pptSkill) {
                        ChatLog.info?.('ORCH', `[orchestrator] PPT 意图优先: ${matchedSkill.name || matchedSkill.Name} → GeneratePPTReport`);
                        matchedSkill = pptSkill;
                        kind = 'code';
                    }
                }

                // === 强制兜底：数据分析请求一律路由到 python-pipeline ===
                // 无论 Skill 声明什么 pipeline，只要用户请求包含数据分析关键词，
                // 都强制走 python-pipeline，避免 VB skill-pipeline 参数问题。
                // 例外：明确 PPT 意图（含"趋势分析页"等 PPT 页面诉求）不强制 python，
                //       否则"在PPT汇报中增加趋势分析页"会被误路由到 python 分析。
                // 正则单一事实源：config/routing-constants.js
                const pyIntent = routing().detectPythonAnalysisIntent(message);

                if (pyIntent.any && !pptIntent && !context?.__forceSkillPipeline) {
                    // 只要有分析意图，一律走 python（即使同时有分组意图）
                    // 仅纯分组请求才不走 python
                    kind = 'python';
                    ChatLog.info?.('ORCH', `[orchestrator] 强制路由: 检测到数据分析意图 → python-pipeline (trend=${pyIntent.trend}, anomaly=${pyIntent.anomaly}, multi=${pyIntent.multi}, group=${pyIntent.groupAgg})`);
                }

                ChatLog.info?.('ORCH', `[orchestrator] skillName=${matchedSkill.name || matchedSkill.Name}, pipeline=${matchedSkill.Pipeline || matchedSkill.pipeline}`);

                context = context || {};
                context._matchedSkill = matchedSkill;
                ChatLog.info?.('ORCH', `[orchestrator] 命中 Skill: ${matchedSkill.name || matchedSkill.Name} → pipeline=${kind}`);
            } else {
                // 未命中任何 Skill（或已被多步骤判定清空 matchedSkill=null），
                // 但用户明确是"写代码/写宏"意图 → 强制走 code-pipeline（只生成展示，不执行）。
                // 根因：此前"帮我写一个按工作表拆分工作簿的 VBA 宏"这类请求会掉进 text-pipeline
                // agent-loop，LLM 误当执行任务直接调用 excel_* 工具拆分干活——用户要的是"写代码"，
                // 不是"用代码干活"。代码生成意图必须无条件走只读的 code-pipeline。
                if (kind === 'text' && (context?.__forceCodePipeline || routing().detectCodeGenerationIntent(message)) && !inheritedFollowUp && !context?._execFollowUp && !pptFreeBuild) {
                    const langDir = (routing().detectLanguageDirective || function () { return null; })(message);
                    const codeSkillName = langDir === 'python' ? 'GeneratePython' : 'GenerateVBA';
                    const codeSkill = (App.cachedSkills || []).find(s => (s.name || s.Name) === codeSkillName);
                    if (codeSkill) {
                        matchedSkill = codeSkill;
                        kind = 'code';
                        context = context || {};
                        context._matchedSkill = codeSkill;
                        ChatLog.info?.('ORCH', `[orchestrator] 无 Skill 但检测到代码生成意图 → 强制 ${codeSkillName} → code-pipeline（只生成展示，不执行）`);
                    } else {
                        // 全部 Skill 停用时 GenerateVBA/GeneratePython 取不到，但"要代码"这个产出目标
                        // 与 Skill 是否启用无关：code-pipeline 是前端只读管道（只展示、不执行），
                        // matchedSkill 为空也能按消息语言生成代码。此处若不放行，写代码请求会掉进
                        // agent-loop 被小模型当成执行任务，直接用 excel_* 工具把活干了。
                        kind = 'code';
                        ChatLog.info?.('ORCH', '[orchestrator] 无可用代码 Skill（已停用）但命中代码生成意图 → code-pipeline（无 Skill 兜底，只生成展示，不执行）');
                    }
                }

                // 未命中任何 Skill，但检测到数据分析意图 → 直接走 python-pipeline
                // 多步骤请求（_skipMultiStep / LLM 意图分析 _forceMultiStep）不强制 python，
                // 应交由 text-pipeline 多步骤管道统一规划（读取→分析→写入），否则"分析+写入"会被 python 劫走。
                // PPT 结构诉求（pptFreeBuild）与一切 PPT 意图（pptIntent）都放行 agent-loop，
                // 不受 python 分析关键词影响——无 Skill 时由大模型自主调用内置 ppt_* MCP 工具生成。
                if (kind === 'text' && !pptFreeBuild && !pptIntent && context?.selection && !context?._skipMultiStep && !context?._forceMultiStep) {
                    const pyIntent = routing().detectPythonAnalysisIntent(message);
                    if (pyIntent.any) {
                        kind = 'python';
                        ChatLog.info?.('ORCH', `[orchestrator] 无 Skill 但检测到数据分析意图 → python-pipeline`);
                    }
                }
            }

            // === PPT 直出安全网 ===
            // shouldRouteToSkill 对短关键词("PPT")命中不足会被 kwSumLen<4 护栏跳过 → matchedSkill=null → kind='text'
            // → 走 agent-loop，可能被后端引导调用 skill_GeneratePPTReport 工具（LLM 反问/编造工具"光说不做"）。
            // 此处：用户明确要 PPT 且已有选区/参考区时，在入口强制锁定 GeneratePPTReport → code-pipeline，
            // 触发前端"模板直出+自动执行"真正生成 PPT，不依赖小模型输出。
            // 例外：明确页面结构/主题要求（pptFreeBuild）时跳过此兜底，放行 agent-loop 自由调 ppt_ 工具。
            // 有"主题/风格"词时同样走直出（GeneratePPTReport 支持 theme 参数，detectPptTheme 会应用商务蓝等）。
            if (!context?._matchedSkill && !pptFreeBuild) {
                context = context || {};
                const ctxHasData =
                    (Array.isArray(context.selection?.headers) && context.selection.headers.length > 0) ||
                    (Array.isArray(context.referenceRegions) &&
                        context.referenceRegions.some(r => Array.isArray(r.headers) && r.headers.length > 0));
                if (pptIntent && ctxHasData) {
                    const pptSkill = (App.cachedSkills || []).find(
                        s => s.name === 'GeneratePPTReport' || s.Name === 'GeneratePPTReport');
                    if (pptSkill) {
                        context._matchedSkill = pptSkill;
                        kind = 'code';
                        ChatLog.info?.('ORCH', '[orchestrator] PPT 直出: 命中 GeneratePPTReport + 选区 → code-pipeline');
                    }
                }
            }

            // === Python 语言指令强制路由 ===
            // 用户明确说"用Python/用pandas"是硬性语言约束：无论 LLM 意图分析是否判定多步骤
            // （_forceMultiStep）、是否命中 Skill，都应走 python-pipeline 生成 Python 分析
            // （脚本+图表+结论），而非被 text-pipeline 的 agent-loop 用 excel_* MCP 工具
            // "宏式执行"（读表→聚合→写表）。纯分组请求（groupAgg）同样适用——
            // 否则"用Python对当前数据做分组聚合统计"会被 LLM 判为多步骤后绕开 Python。
            // 例外：PPT 意图、确定性 Skill 强制（__forceSkillPipeline）不覆盖。
            const langDirective = routing().detectLanguageDirective(message);
            if (langDirective === 'python' && !pptIntent && !context?.__forceSkillPipeline) {
                const pyLang = routing().detectPythonAnalysisIntent(message);
                if (pyLang.any || pyLang.groupAgg) {
                    kind = 'python';
                    ChatLog.info?.('ORCH', `[orchestrator] Python 语言指令强制路由 → python-pipeline (lang=${langDirective}, groupAgg=${pyLang.groupAgg})`);
                }
            }

            // === 最终兜底：在 pickPipeline 之前强制路由 ===
            // 仅当未命中任何 Skill / PPT 直出（kind 仍为 text）时，才按数据分析关键词强制走 python。
            // 若已路由到具体 pipeline（skill/code/html/python 等）或明确 PPT 意图，不再覆盖——
            // 否则"在PPT汇报中增加趋势分析页"会被 PPT 直出锁定为 code 后又被"趋势"覆盖回 python。
            // 使用保守集（conservative），避免"变化"等宽泛词在无 Skill/无选区时误路由。
            // LLM 意图分析已判定需多步骤（_forceMultiStep=true）时不强制 python：
            // 多步骤请求应交由 text-pipeline 多步骤管道统一规划（读取→分析→写入），
            // 否则"分析+写入"会被 python 关键词劫走，意图分析结果被丢弃。
            const pyFinal = routing().detectPythonAnalysisIntent(message, { conservative: true });
            if (pyFinal.any && kind === 'text' && !pptIntent && !context?._forceMultiStep) {
                ChatLog.info?.('ORCH', `[FINAL] 强制路由: ${kind} → python (trend=${pyFinal.trend}, anomaly=${pyFinal.anomaly}, desc=${pyFinal.descriptive}, multi=${pyFinal.multi}, group=${pyFinal.groupAgg})`);
                kind = 'python';
            }

            // === 路由完成：同步最终 kind / context 到 turn ===
            // turn 在路由前已启动（气泡立即显示），此处把路由决定的最终 kind 与
            // 注入的上下文（_forceMultiStep/_intent/_matchedSkill 等）同步回 turn，
            // 渲染器据此修正步骤/工具卡显示，pipeline 据此选择。
            ChatLog.info?.('ORCH', `[FINAL-ROUTE] kind=${kind} matchedSkill=${matchedSkill ? (matchedSkill.name || matchedSkill.Name) : 'null'} inheritedFollowUp=${inheritedFollowUp} __forceSkillPipeline=${context?.__forceSkillPipeline || 'null'} _execFollowUp=${!!context?._execFollowUp}`);
            turn.context = context || {};
            if (turn.kind !== kind) {
                turn.kind = kind;
                // 触发 turn_update，让渲染器按最终 kind 重评估 _suppressSteps/_suppressToolCards/finalRenderOnly
                sessionStore.updateTurn(turn.id, { kind });
            }

            // 路由期间若已被新 turn 取消（controller 被 abort），不再执行本 turn 的 pipeline
            if (controller.signal.aborted) {
                ChatLog.info?.('ORCH', `[orchestrator] turn ${turn.id} 在路由期间被取消，跳过执行`);
                return turn.id;
            }

            // 选择 pipeline
            const pipeline = pickPipeline(kind);
            if (!pipeline || typeof pipeline.execute !== 'function') {
                ChatLog.error?.('ORCH', `[orchestrator] no pipeline for kind=${kind}`);
                sessionStore.failTurn(turn.id, {
                    code: 'NO_PIPELINE',
                    message: `未找到处理流程: ${kind}`
                });
                _activeControllers.delete(turn.id);
                return turn.id;
            }

            ChatLog.info?.('ORCH', `[orchestrator] dispatch turn=${turn.id} kind=${kind} attachments=${attachments?.length || 0}`);

            // 异步执行 pipeline，消费 typed events
            //    （不阻塞调用方；事件由 sessionStore 广播给 renderer）
            this._runPipeline(turn, pipeline, controller);

            return turn.id;
            } catch (routeErr) {
                // 路由/派发阶段异常：清理 turn 与 controller，避免卡死
                ChatLog.error?.('ORCH', `[orchestrator] 路由阶段异常 turn=${turn.id}:`, routeErr?.message || routeErr);
                _orchHostLog('ERROR', `[orchestrator] 路由阶段异常 turn=${turn.id}: ${routeErr?.message || routeErr}`);
                try {
                    if (sessionStore.getCurrentTurn()?.id === turn.id) {
                        sessionStore.failTurn(turn.id, {
                            code: 'ROUTE_ERROR',
                            message: routeErr?.message || String(routeErr)
                        });
                    }
                } catch (_) { /* 忽略二次异常 */ }
                _activeControllers.delete(turn.id);
                return turn.id;
            }
        },

        /**
         * 内部：运行 pipeline 并消费事件
         */
        async _runPipeline(turn, pipeline, controller) {
            const sessionStore = App.ChatV2.sessionStore;
            const signal = controller.signal;

            // 订阅 agentEvents，桥接 agent/* 生命周期事件
            // 供 UI/调试面板订阅；不改变现有 sessionStore 通知机制
            const agentEvents = App.ChatV2?.agentEvents;
            let offTurnEnd = null;
            if (agentEvents) {
                offTurnEnd = agentEvents.on('turn/end', (payload) => {
                    if (payload?.turnId !== turn.id) return;
                    ChatLog.debug?.('ORCH', `[orchestrator] agent turn/end: ${turn.id} reason=${payload.reason}`);
                });
            }

            try {
                for await (const evt of pipeline.execute(turn, signal)) {
                    // sessionStore 在 agentLoop 内已处理 delta / toolCall 等更新
                    // 此处仅处理 pipeline / agent-loop 层面的终态事件
                    switch (evt.type) {
                        case 'canceled':
                            // 已由 cancel() 触发（abort + cancelTurn），此处确保状态一致
                            if (sessionStore.getCurrentTurn()?.id === turn.id) {
                                sessionStore.cancelTurn(turn.id, evt.reason || 'aborted');
                            }
                            ChatLog.info?.('ORCH', `[orchestrator] turn ${turn.id} canceled`);
                            break;

                        case 'error':
                            // pipeline 已经在 catch 中 failTurn；这里仅记录
                            ChatLog.error?.('ORCH', `[orchestrator] turn ${turn.id} error`, evt.error);
                            break;

                        case 'status':
                            // 状态更新：传递给 sessionStore，渲染器显示在思考区
                            if (sessionStore.getCurrentTurn()?.id === turn.id) {
                                sessionStore.setTurnStatus(turn.id, evt.text);
                            }
                            ChatLog.info?.('ORCH', `[orchestrator] status: ${evt.text}`);
                            break;

                        case 'pipeline_done':
                            ChatLog.info?.('ORCH', `[orchestrator] turn ${turn.id} pipeline_done (${evt.pipeline})`);
                            break;

                        case 'loop_done':
                            ChatLog.debug?.('ORCH', `[orchestrator] turn ${turn.id} loop_done: ${evt.reason}`);
                            break;

                        default:
                            // token / reasoning / tool_call_start / tool_call_result / round_start / round_end / usage / done
                            // 这些事件由 sessionStore 通知系统广播
                            break;
                    }
                }
            } catch (err) {
                ChatLog.error?.('ORCH', `[orchestrator] turn ${turn.id} uncaught error`, err);
                if (sessionStore.getCurrentTurn()?.id === turn.id) {
                    sessionStore.failTurn(turn.id, {
                        code: 'UNCAUGHT',
                        message: err?.message || String(err)
                    });
                }
            } finally {
                if (offTurnEnd) offTurnEnd();
                _activeControllers.delete(turn.id);
            }
        },

        /**
         * 取消指定 turn
         */
        cancel(turnId, reason = 'user_aborted') {
            const controller = _activeControllers.get(turnId);
            if (controller) {
                // reason 必须传下去：不传时 signal.reason 是默认 DOMException
                // （message 即 "signal is aborted without reason"），下游 fetch/LLM
                // 调用的 catch 只能看到这句默认文案，分不清取消来源。
                try { controller.abort(new DOMException(String(reason), 'AbortError')); }
                catch (_) { try { controller.abort(); } catch (_2) { /* ignore */ } }
                _activeControllers.delete(turnId);
            }
            const sessionStore = App.ChatV2.sessionStore;
            if (sessionStore) {
                // cancelTurn 内部检查 currentTurn，幂等
                sessionStore.cancelTurn(turnId, reason);
            }
            ChatLog.info?.('ORCH', `[orchestrator] cancel turn ${turnId}: ${reason}`);
        },

        /**
         * 取消当前正在进行的 turn
         */
        cancelCurrent(reason = 'user_aborted') {
            const current = App.ChatV2.sessionStore?.getCurrentTurn();
            if (current && current.status === 'working') {
                this.cancel(current.id, reason);
                return true;
            }
            return false;
        },

        /**
         * 用户点击"确认候选"按钮后的回调入口。
         * 由 chat-renderer 渲染的 skill-confirm-chip 通过 window.App.ChatV2.orchestrator.resumeSkillConfirm(turnId, choice) 触发。
         * choice ∈ { 'keep' → 保留 regex Skill, 'drop' → 接受 LLM text }
         * 找不到 resolver（已超时或 abort 清掉）则静默忽略。
         */
        resumeSkillConfirm(turnId, choice) {
            if (!turnId) return;
            const resolver = _confirmResolvers.get(turnId);
            if (resolver) {
                _confirmResolvers.delete(turnId);
                ChatLog.info?.('ORCH', `[confirm] resumeSkillConfirm turn=${turnId} choice=${choice}`);
                resolver({ choice: choice === 'drop' ? 'drop' : 'keep' });
            } else {
                ChatLog.info?.('ORCH', `[confirm] resumeSkillConfirm turn=${turnId} 无挂起 resolver（可能已超时/abort），忽略`);
            }
        },

        /**
         * 是否有 turn 正在进行
         */
        isRunning() {
            return _activeControllers.size > 0;
        },

        /**
         * 当前 turn（代理到 sessionStore）
         */
        getCurrentTurn() {
            return App.ChatV2.sessionStore?.getCurrentTurn();
        },

        /**
         * 调试信息
         */
        getDebugInfo() {
            return {
                activeTurns: Array.from(_activeControllers.keys()),
                running: _activeControllers.size > 0
            };
        },

        /**
         * 写类工具裁决（暴露给测试/调试）：判断命中项是"执行命令"还是"只想看代码"。
         * 典型回归：'将当前文档导出为 PDF' 必须判为 command（否则被放行到 agent-loop，
         * 小模型会先做一次多余的表格美化再导出）。
         */
        adjudicateWriteHit,

        /** 工具名是否含副作用动词（导出/写入/格式化…），决定是否走裁决闸门 */
        isSideEffectTool,

        /** 工具名是否"只产出文件、不改动宿主内容"（导出/保存类），决定是否免 LLM 裁决 */
        isArtifactOnlyTool,

        /**
         * excel_* 工具结果友好化（暴露给测试/调试）。
         * 回归：excel_merge_sheets 成功返回（含 success/sourceSheets/mergedRows/data/steps）
         * 必须落成一行可读摘要，绝不能整坨 JSON 倒进气泡（2026-10-09 用户反馈）。
         */
        excelFriendlyText,
    };

    App.ChatV2.orchestrator = orchestrator;
    ChatLog.info?.('CHAT-V2', '✅ orchestrator loaded');
})();
