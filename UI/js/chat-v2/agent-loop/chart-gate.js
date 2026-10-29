/**
 * chat-v2/agent-loop/chart-gate.js — 图表交付闸门 + 补链
 * ============================================================
 * 从 agent-loop.js 抽离。技术债：forceRetryCount 跨守卫共用、
 * continue 不计 round —— 见主循环注释，本轮不改预算语义。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    function _hostLog(level, message) {
        try {
            if (typeof App.ChatV2.agentLoopHostLog === 'function') {
                App.ChatV2.agentLoopHostLog(level, message);
                return;
            }
        } catch (_e) { /* ignore */ }
        const fn = level === 'WARN' ? ChatLog.warn : ChatLog.info;
        fn?.('AGENT', '[chart-gate] ' + message);
    }

    function emitAgent(event, payload) {
        try { App.ChatV2?.agentEvents?.emit(event, payload); } catch (e) {
            ChatLog.warn?.('AGENT', '[chart-gate] emit ' + event + ' 失败:', e);
        }
    }

    /**
     * 根据用户消息推断图表类型枚举值（供 _forceChart 兜底自动建图使用）。
     * 优先使用调用方传入的覆盖值，其次读取 intent-config 的 chartTypeDefs，
     * 最后按关键词做简单兜底。
     */
    function inferChartType(userMsg, override) {
        if (override) return override;
        if (!userMsg) return 'xlColumnClustered';
        try {
            const defs = (typeof App !== 'undefined' && App.ChatV2 && App.ChatV2.intentConfig && typeof App.ChatV2.intentConfig.defs === 'function')
                ? App.ChatV2.intentConfig.defs() : null;
            const chartTypeDefs = defs && Array.isArray(defs.chartTypeDefs) ? defs.chartTypeDefs : [];
            for (const d of chartTypeDefs) {
                if (d && d.re && d.re.test(userMsg)) return d.enum || 'xlColumnClustered';
            }
        } catch (_e) { /* 意图配置不可用 → 走兜底 */ }
        if (/饼|pie|占比|扇形/i.test(userMsg)) return 'xlPie';
        if (/折线|line|趋势/i.test(userMsg)) return 'xlLineMarkers';
        if (/横向柱|条形|bar/i.test(userMsg)) return 'xlBarClustered';
        if (/柱|column/i.test(userMsg)) return 'xlColumnClustered';
        if (/圆环|环形|doughnut/i.test(userMsg)) return 'xlDoughnut';
        if (/面积|area/i.test(userMsg)) return 'xlArea';
        if (/散点|scatter/i.test(userMsg)) return 'xlXYScatter';
        if (/雷达|radar/i.test(userMsg)) return 'xlRadar';
        return 'xlColumnClustered';
    }

    /**
     * 【图表交付闸门】判定本 turn 是不是"必须产出图表"的请求。
     *
     * 背景（实测 2026-10-10 07:46）："使用饼图按部门展示销售额"在全部 Skill 被停用后
     * 走通用多步骤管道，弱模型第 1 轮零工具调用只写了 712 字分析，守卫链补了一次
     * excel_read_range，覆盖度自检判"未完成=0"直接收尾 —— 用户拿到一大段文字、零张图。
     * 图表请求的"完成"必须由**图的存在**来定义，不能由模型自称完成来定义。
     *
     * 疑问句（"饼图怎么做" / "图表有什么区别"）不是出图指令，不强制。
     *
     * @param {Object} trn turn 对象
     * @param {string} rawMsg 用户原始消息（未经 prompt-builder 元信息污染）
     * @returns {boolean}
     */
    function isChartIntent(trn, rawMsg) {
        if (trn && trn.context && (trn.context._forceChart === true || trn.context._forceDashboard === true)) return true;
        const msg = String(rawMsg || (trn && (trn.userMessage || trn.message)) || '');
        if (!msg) return false;
        // 疑问句排除：同一套语气词在 skill-domain 里已用于"是动作还是提问"的分流，此处复用语义。
        if (/[?？]|怎么|怎样|咋|什么|为何|为什么|区别|吗|呢/.test(msg)) return false;
        let re = null;
        try { re = (App.ChatV2?.routingConstants?.STRONG_CHART_INTENT_RE) || null; } catch (_e) { re = null; }
        // 本地兜底副本：routing-constants 未加载时仍要能判定（与本文件共存的独立性优先）
        if (!re) re = /饼图|柱图|柱状图|条形图|折线图|散点图|面积图|雷达图|直方图|趋势图|仪表盘|dashboard|数据大屏|可视化|图形化|画图|制图|做图|出图|生成.*图|创建.*图|画.*图|图表类型|chart\s*(pie|bar|line|scatter|area|radar|histogram|gauge|funnel)?/i;
        return re.test(msg);
    }

    /**
     * 【图表补链】为"要出图、但模型连聚合都没做"的场景推断聚合参数。
     *
     * 数据源优先级：最近一次成功 excel_read_range 的实参（模型自己认定过的数据区）
     *              → 当前选区。
     * 维度/数值列：只用**选区表头**推断（可靠）——第一个非数值列=维度、第一个数值列=数值；
     *              推断不出维度列时返回 null（猜错会做出废图，宁可交回模型）。
     *
     * @param {Object} trn turn 对象
     * @returns {{range:string,sheet:string,dimColumn:string,valueColumn:string,title:string}|null}
     */
    function resolveAggregateSource(trn) {
        if (!trn) return null;
        const ctx = trn.context || {};
        const sel = ctx.selection || null;
        let range = '';
        let sheet = '';
        const calls = Array.isArray(trn._executedToolCalls) ? trn._executedToolCalls : [];
        for (let i = calls.length - 1; i >= 0; i--) {
            const c = calls[i];
            if (c && c.name === 'excel_read_range' && c.args) {
                range = String(c.args.range || '');
                sheet = String(c.args.sheet || '');
                if (range) break;
            }
        }
        if (!range && sel) range = String(sel.address || sel.range || '');
        if (!sheet && sel) sheet = String(sel.sheetName || sel.sheet || sel.worksheet || '');
        if (!range) return null;
        const headers = Array.isArray(sel && sel.headers)
            ? sel.headers.filter(h => h !== null && h !== undefined && String(h).trim() !== '')
            : [];
        const rows = Array.isArray(sel && sel.rows) ? sel.rows.slice(0, 20) : [];
        const toNum = (v) => {
            if (typeof v === 'number') return v;
            const s = String(v === null || v === undefined ? '' : v).replace(/[,%￥$¥\s]/g, '');
            if (s === '') return NaN;
            const n = Number(s);
            return Number.isNaN(n) ? NaN : n;
        };
        const isNumCol = (ci) => {
            let seen = 0;
            let num = 0;
            for (const r of rows) {
                const v = Array.isArray(r) ? r[ci] : (r && typeof r === 'object' ? r[headers[ci]] : undefined);
                if (v === null || v === undefined || String(v).trim() === '') continue;
                seen++;
                if (!Number.isNaN(toNum(v))) num++;
            }
            return seen > 0 && num >= Math.ceil(seen * 0.8);
        };
        let valueColumn = '';
        for (let ci = 0; ci < headers.length; ci++) {
            if (!valueColumn && isNumCol(ci)) valueColumn = String(headers[ci]);
        }
        // 【维度择优】旧逻辑取"第一个非数值列"当维度 —— 实测 2026-10-28 会选中「日期」列，
        // 而在「地区」列只有 1 个取值的数据上也会选中「地区」→ 图表退化成一根柱。
        // 改为按"取值基数 + 表头典型性"打分择优：基数权重高于关键词，distinct<=1 直接淘汰。
        const distinctOf = (ci) => {
            const set = new Set();
            for (const r of rows) {
                const v = Array.isArray(r) ? r[ci] : (r && typeof r === 'object' ? r[headers[ci]] : undefined);
                if (v === null || v === undefined || String(v).trim() === '') continue;
                set.add(String(v).trim());
            }
            return set.size;
        };
        const isDateCol = (ci) => {
            if (/日期|时间|年份|年度|月份|编号|序号|编码|id/i.test(String(headers[ci] || ''))) return true;
            let seen = 0, hit = 0;
            for (const r of rows) {
                const v = Array.isArray(r) ? r[ci] : (r && typeof r === 'object' ? r[headers[ci]] : undefined);
                if (v === null || v === undefined || String(v).trim() === '') continue;
                seen++;
                if (v instanceof Date) { hit++; continue; }
                if (/^\d{4}[-/年]\d{1,2}[-/月]\d{1,2}/.test(String(v).trim())) hit++;
            }
            return seen > 0 && hit >= Math.ceil(seen * 0.8);
        };
        const keywordScore = (h) => {
            if (/部门/.test(h)) return 100;
            if (/地区/.test(h)) return 90;
            if (/区域/.test(h)) return 88;
            if (/大区/.test(h)) return 87;
            if (/产品|商品|品类/.test(h)) return 85;
            if (/渠道|来源/.test(h)) return 80;
            if (/分类|类别|类型/.test(h)) return 75;
            if (/城市|省份|国家/.test(h)) return 70;
            if (/客户/.test(h)) return 60;
            return 0;
        };
        const scoreOf = (ci) => {
            const d = distinctOf(ci);
            if (d <= 1) return -1;
            let base = (d >= 2 && d <= 20) ? 90 : (d <= 50 ? 60 : 30);
            if (rows.length > 0 && d >= rows.length) base -= 20; // 每行唯一 = 明细列
            if (base < 0) base = 0;
            return base * 2 + keywordScore(String(headers[ci] || ''));
        };
        let dimColumn = '';
        let bestScore = 0;
        let firstNonNumeric = '';
        for (let ci = 0; ci < headers.length; ci++) {
            if (String(headers[ci] || '').trim() === '') continue;
            if (isNumCol(ci) || isDateCol(ci)) continue;
            if (!firstNonNumeric) firstNonNumeric = String(headers[ci]);
            const sc = scoreOf(ci);
            if (sc > bestScore) { bestScore = sc; dimColumn = String(headers[ci]); }
        }
        // 所有分类列都只有 1 个取值（数据本身无区分度）→ 退回旧行为，交给模型判断
        if (!dimColumn) dimColumn = firstNonNumeric;
        // 【2026-10-28 实测】没有选区快照 headers（用户没手动选中、上下文也没带 headers）时，
        // 旧逻辑直接 return null → 整个补链跳过 → 模型空手收尾（"图表没有执行"）。
        // 现在改为把 dimColumn 留空交给后端自动择优（后端 fallback 会按取值基数+关键词挑列，
        // 只有 1 组时还会自动换列），补链永不因"前端猜不出列"而放弃。
        if (!dimColumn) {
            return {
                range: range,
                sheet: sheet,
                dimColumn: '',
                valueColumn: '',
                autoDim: true,
                title: '数据汇总'
            };
        }
        return {
            range: range,
            sheet: sheet,
            dimColumn: dimColumn,
            valueColumn: valueColumn,
            autoDim: false,
            title: dimColumn + (valueColumn ? '按' + valueColumn : '') + '统计'
        };
    }

    /**
     * 构造"补图引导"提示词（【图表交付闸门】用；抽成纯函数以便回归测试与判定保持同步）。
     * @returns {{text:string, src:Object|null, chartType:string}}
     */
    function buildChartNudgeHint(trn, rawMsg, chartTypeOverride) {
        const _srcHint = resolveAggregateSource(trn);
        const _ctHint = inferChartType(rawMsg, chartTypeOverride);
        const _aggTail = (_srcHint ? (_srcHint.valueColumn ? `，valueColumn="${_srcHint.valueColumn}"` : '')
                + (_srcHint.range ? `，range="${_srcHint.range}"` : '')
                + (_srcHint.sheet ? `，sheet="${_srcHint.sheet}"` : '') : '')
            + '，aggFunc="sum"）';
        const _step1 = _srcHint
            ? (_srcHint.dimColumn
                ? `excel_aggregate（dimColumn="${_srcHint.dimColumn}"` + _aggTail
                : 'excel_aggregate（dimColumn 留空=由系统自动择优维度' + _aggTail)
            : 'excel_aggregate（dimColumn=分组维度列名，valueColumn=数值列名，aggFunc="sum"）';
        const text = '⏳ 任务还没完成：用户这次要的是【图表】，但到目前为止没有生成任何图表。\n\n'
            + '只有文字分析不算完成。请立刻按下面两步执行（缺一不可）：\n'
            + `  1) 调用 ${_step1}，得到"分组→汇总"两列结果；\n`
            + `  2) 调用 excel_create_chart：dataRange=第 1 步聚合结果所在区域（含表头行），chartType="${_ctHint}"。\n\n`
            + '严禁直接用原始明细数据建图（那样的图例是一堆明细行，不是按维度汇总的结果）。做完这两步再给最终总结。';
        return { text: text, src: _srcHint, chartType: _ctHint };
    }

    /**
     * 图表补链：合成 excel_aggregate → excel_create_chart。
     * @param {object} p { turn, toolCalls, localResults, round, stepNo, signal, rawMsg, chartTypeOverride }
     */
    async function* runChartBackfill(p) {
        const toolRegistry = App.ChatV2.toolRegistry;
        const sessionStore = App.ChatV2.sessionStore;
        const out = { ran: false, anyToolFailed: false, chartDone: false };
        const trn = p.turn;
        const _srcAuto = resolveAggregateSource(trn);
        if (!_srcAuto) {
            _hostLog('WARN', '图表补链跳过：推断不出聚合维度列（选区无表头）或没有可用数据区 → 交回模型');
            return out;
        }
        trn._aggFallbackTried = true;
        out.ran = true;
        const _aggSheet = String((_srcAuto.dimColumn || 'AI') + '汇总').replace(/[\\/:*?\[\]]/g, '').slice(0, 28) || 'AI汇总';
        const _aggArgs = { aggFunc: 'sum', targetSheet: _aggSheet };
        // dimColumn 留空 = 交给后端自动择优（后端 fallback 按取值基数+关键词挑列，1 组时自动换列）
        if (_srcAuto.dimColumn) _aggArgs.dimColumn = _srcAuto.dimColumn;
        if (_srcAuto.valueColumn) _aggArgs.valueColumn = _srcAuto.valueColumn;
        // 行简写地址（"A1:21" 这类）不直接传给后端：Excel 会把它解析成 16384 列整行，
        // 交给后端的 RangeContext（上次 read_range 的真实区域）更可靠。
        if (_srcAuto.range && !/^[A-Za-z]{1,3}\d+:\d+$/.test(String(_srcAuto.range))) _aggArgs.range = _srcAuto.range;
        if (_srcAuto.sheet) _aggArgs.sheet = _srcAuto.sheet;
        const _aggTc = { id: 'syn_agg_' + Math.random().toString(36).slice(2, 8), name: 'excel_aggregate', args: _aggArgs };
        ChatLog.warn?.('AGENT', '[agent-loop] 图表补链：LLM 未聚合也未建图 → 自动合成 excel_aggregate（'
            + _srcAuto.dimColumn + (_srcAuto.valueColumn ? '×' + _srcAuto.valueColumn : '') + '@' + _srcAuto.range + '）');
        _hostLog('WARN', '图表补链：自动合成 excel_aggregate（dimColumn=' + _srcAuto.dimColumn
            + (_srcAuto.valueColumn ? '，valueColumn=' + _srcAuto.valueColumn : '') + '，range=' + _srcAuto.range
            + (_srcAuto.sheet ? '，sheet=' + _srcAuto.sheet : '') + '，targetSheet=' + _aggSheet + '）');
        let _aggRes = null;
        try {
            _aggRes = await toolRegistry.invokeTracked(trn, 'excel_aggregate', _aggArgs, {
                round: p.round + 1, ctx: { turn: trn, signal: p.signal }
            });
        } catch (_aggErr2) {
            _aggRes = { ok: false, error: _aggErr2?.message || String(_aggErr2), durationMs: 0, toolCallId: null };
        }
        if (!_aggRes) return out;
        p.toolCalls.push(_aggTc);
        p.localResults.push(_aggRes);
        emitAgent('tool/result', {
            turnId: trn.id, step: p.stepNo, toolName: _aggTc.name, callId: _aggTc.id,
            ok: !!_aggRes.ok, error: _aggRes.error || null,
            durationMs: _aggRes.durationMs || 0, toolCallId: _aggRes.toolCallId || null
        });
        try {
            sessionStore?.appendStepLog?.(trn.id, 'tool_return',
                '工具 ' + _aggTc.name + ' ' + (_aggRes.ok ? '成功' : '失败') + '（图表补链）', p.stepNo);
        } catch (e) { /* 忽略 */ }
        yield {
            type: 'tool_call_result', toolCallId: _aggRes.toolCallId, toolName: _aggTc.name,
            ok: _aggRes.ok, result: _aggRes.result, error: _aggRes.error, durationMs: _aggRes.durationMs
        };
        if (!_aggRes.ok) { out.anyToolFailed = true; return out; }
        const _m = (_aggRes.result && typeof _aggRes.result === 'object') ? _aggRes.result : {};
        const _writtenTo = _m.writtenTo || _m.resultSheet || _m.targetSheet || _aggSheet;
        const _gc = Number(_m.groupCount)
            || (Array.isArray(_m.data) ? _m.data.length : 0)
            || (Array.isArray(_m.rows) ? _m.rows.length : 0);
        if (!_gc) {
            _hostLog('WARN', '图表补链：聚合成功但取不到分组行数 → 跳过自动建图，交回模型');
            return out;
        }
        if (_gc <= 1) {
            // 后端已做过"维度择优"仍只有 1 组 = 数据里确实没有可分组维度（所有分类列都只有 1 个取值）。
            // 这种情况建图必然是一根柱/一个扇区，属于废图 —— 不建，交回模型如实说明数据问题。
            _hostLog('WARN', '图表补链：聚合只有 1 组（维度无区分度）→ 不建废图，交回模型说明数据问题');
            return out;
        }
        const _cArgs = {
            chartType: inferChartType(p.rawMsg, p.chartTypeOverride),
            dataRange: _writtenTo + '!A1:B' + (_gc + 1),
            sheet: _writtenTo,
            title: _srcAuto.title || (String(_m.dimColumn || '') ? String(_m.dimColumn) + '汇总' : '图表')
        };
        const _cTc = { id: 'syn_chart_' + Math.random().toString(36).slice(2, 8), name: 'excel_create_chart', args: _cArgs };
        ChatLog.warn?.('AGENT', '[agent-loop] 图表补链：自动合成 excel_create_chart，数据源=' + _cArgs.dataRange);
        _hostLog('WARN', '图表补链：自动合成 excel_create_chart（dataRange=' + _cArgs.dataRange + '，chartType=' + _cArgs.chartType + '）');
        let _cRes = null;
        try {
            _cRes = await toolRegistry.invokeTracked(trn, 'excel_create_chart', _cArgs, {
                round: p.round + 1, ctx: { turn: trn, signal: p.signal }
            });
        } catch (_chartErr2) {
            _cRes = { ok: false, error: _chartErr2?.message || String(_chartErr2), durationMs: 0, toolCallId: null };
        }
        if (!_cRes) return out;
        p.toolCalls.push(_cTc);
        p.localResults.push(_cRes);
        emitAgent('tool/result', {
            turnId: trn.id, step: p.stepNo, toolName: _cTc.name, callId: _cTc.id,
            ok: !!_cRes.ok, error: _cRes.error || null,
            durationMs: _cRes.durationMs || 0, toolCallId: _cRes.toolCallId || null
        });
        try {
            sessionStore?.appendStepLog?.(trn.id, 'tool_return',
                '工具 ' + _cTc.name + ' ' + (_cRes.ok ? '成功' : '失败') + '（图表补链）', p.stepNo);
        } catch (e) { /* 忽略 */ }
        yield {
            type: 'tool_call_result', toolCallId: _cRes.toolCallId, toolName: _cTc.name,
            ok: _cRes.ok, result: _cRes.result, error: _cRes.error, durationMs: _cRes.durationMs
        };
        if (_cRes.ok) { trn._chartDone = true; out.chartDone = true; }
        else { out.anyToolFailed = true; }
        return out;
    }

    App.ChatV2.agentLoopChartGate = {
        inferChartType,
        isChartIntent,
        resolveAggregateSource,
        buildChartNudgeHint,
        runChartBackfill
    };
    ChatLog.info?.('CHAT-V2', '✅ agentLoopChartGate loaded');
})();
