/**
 * ============================================================
 * chat-v2/services/task-planner.js — 任务拆分与规划
 * ============================================================
 * 职责：从 text-pipeline 抽离的任务拆分逻辑。
 *   - splitChartTasks：把用户消息拆成图表/浏览器展示/写表子任务
 *   - splitAggregateTasks：拆聚合子任务（含多子句拆分）
 *   - buildSingleAggregateTask：构建单个聚合任务对象
 *   - buildAggregateSystemPrompt / buildAggregateUserText：聚合任务的 prompt 构造
 * 依赖：intent-detector（所有 detect* 函数与正则常量）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 从 intentDetector 引用全部 detect 函数（避免在 planner 内重复定义）
    function D() { return App.ChatV2.intentDetector; }

    const CHART_TASK_SYSTEM_PROMPT =
        '你是专注于 Excel 图表的助手。用户会给你一个明确的图表任务，请只调用 excel_create_chart 创建这一个图表，' +
        '创建完成后立即结束。严格禁止一次创建多个图表，也严格禁止调用 OCR/图片识别工具（Excel 数据直接用 excel_read_range 读取）。';

    /**
     * 浏览器展示意图判定（收紧版）："用饼图展示销售额"是图表任务，不是浏览器展示。
     * 仅当满足以下任一条件才判为浏览器展示：
     *   1. 文本含强浏览器关键词（浏览器/网页/html/打开等）——用户明确要面板/浏览器展示
     *   2. 文本既不含图表词也不含聚合意图——"展示数据/用表格展示"等纯展示诉求
     * 否则（图表词/聚合意图与"展示"同现）一律按图表/聚合任务处理。
     */
    function isBrowserDisplayTask(text, det) {
        if (!text || !det || typeof det.isBrowserDisplayIntent !== 'function') return false;
        if (!det.isBrowserDisplayIntent(text)) return false;
        if (/浏览器|网页|html|在浏览器中|用浏览器|打开.*查看/i.test(text)) return true;
        if (det.CHART_WORD_RE && det.CHART_WORD_RE.test(text)) return false;
        if (typeof det.detectAggregateIntent === 'function' && det.detectAggregateIntent(text)) return false;
        return true;
    }

    /**
     * 把一条用户消息拆分为多个子任务，覆盖三类操作：
     *   - 图表子任务：chartLabel/chartEnum
     *   - 浏览器展示子任务：isBrowserDisplay=true
     *   - 写入工作表子任务：targetSheet=表名
     */
    function splitChartTasks(text) {
        if (!text) return null;
        const det = D();
        if (!det) return null;

        let protectedText = String(text);
        // 拆分词表复用 intent-detector 的单一源 CHART_SPLIT_WORDS（别再手写字面量）
        const splitWords = det.CHART_SPLIT_WORDS || ['表格', '图表', '饼图', '柱状图', '条形图', '折线图', '散点图', '面积图', '雷达图', '图形', '视图', '数据', '列表', '明细'];
        const splitA = splitWords.map(w => w).join('|');
        protectedText = protectedText.replace(
            new RegExp(`(?:${splitA})\\s*(?:和|与|及)\\s*(?:${splitA})`, 'g'),
            function(match) { return match.replace(/\s*(?:和|与|及)\s*/, '\x00'); }
        );

        const parts = protectedText
            .split(/[，,；;。、]|\s*(?:并|且|以及|同时)\s*/)
            .map(s => s.replace(/\x00/g, '和').trim())
            .filter(Boolean);
        if (parts.length === 0) return null;

        if (parts.length <= 1) {
            if (isBrowserDisplayTask(text, det)) {
                return [{
                    text: text.trim(), isBrowserDisplay: true, chartLabel: 'HTML展示',
                    displayMode: det.detectDisplayMode(text)
                }];
            }
            const target = det.extractTargetSheet(text);
            if (target && !/^=.+|[=\(\)]/.test(target)) {
                return [{ text: text.trim(), targetSheet: target, chartLabel: '写入' }];
            }
            if (!det.CHART_WORD_RE.test(text)) return null;
            const ct = det.detectChartType(text);
            return [{
                text: text.trim(),
                chartLabel: ct ? ct.label : '图表',
                chartEnum: ct ? ct.enum : null
            }];
        }

        let hasAnyTask = false;
        const tasks = [];
        for (const p of parts) {
            if (isBrowserDisplayTask(p, det)) {
                tasks.push({ text: p, isBrowserDisplay: true, chartLabel: 'HTML展示', displayMode: det.detectDisplayMode(p) });
                hasAnyTask = true;
            } else if (det.extractTargetSheet(p)) {
                const ts = det.extractTargetSheet(p);
                if (ts && !/^=.+|[=\(\)]/.test(ts)) {
                    tasks.push({ text: p, targetSheet: ts, chartLabel: '写入' });
                    hasAnyTask = true;
                }
            } else if (det.CHART_WORD_RE.test(p)) {
                const ct = det.detectChartType(p);
                tasks.push({
                    text: p,
                    chartLabel: ct ? ct.label : '图表',
                    chartEnum: ct ? ct.enum : null
                });
                hasAnyTask = true;
            } else if (tasks.length > 0) {
                const last = tasks[tasks.length - 1];
                last.text += '，' + p;
            }
        }
        return hasAnyTask && tasks.length >= 1 ? tasks : null;
    }

    /**
     * 构建单个聚合任务（内部函数）
     */
    function buildSingleAggregateTask(text, forceChartType) {
        const det = D();
        if (!det) return null;
        const dims = det.detectDimensionColumns(text);
        const vals = det.detectValueColumns(text);
        const aggDef = det.detectAggFunc(text);
        const advanced = det.detectAdvancedScenario(text);
        const filterCond = det.detectFilterConditions(text);
        const chartTypeEnum = forceChartType || (det.detectChartType(text)?.enum) || null;

        // 场景 8：筛选条件汇总 → Python
        if (filterCond && dims.length > 0 && vals.length > 0) {
            return {
                tool: 'excel_run_python_script',
                params: {
                    scriptType: 'aggregate_filtered',
                    advancedType: 'filtered',
                    filterColumn: filterCond.filterColumn,
                    filterValue: filterCond.filterValue,
                    filterText: filterCond.filterText,
                    dimColumns: dims,
                    valueColumns: vals,
                    aggFunc: aggDef.func,
                    aggLabel: aggDef.label,
                    text: text
                },
                label: `筛选${filterCond.filterText}后按${dims.join('+')}${aggDef.label}${vals.join('、')}`,
                type: 'filtered_aggregate',
                text: text
            };
        }

        // 场景 9-11：高级场景 → Python
        if (advanced) {
            return {
                tool: 'excel_run_python_script',
                params: {
                    scriptType: 'aggregate_advanced',
                    advancedType: advanced.type,
                    dimColumns: dims,
                    valueColumns: vals,
                    aggFunc: aggDef.func,
                    label: advanced.label,
                    text: text
                },
                label: `${advanced.label}：按${dims.join('+')}${aggDef.label}${vals.join('、')}`,
                type: 'advanced_aggregate',
                text: text
            };
        }

        // 场景 2-3：多维度 OR 多指标 → excel_create_pivot_table
        if (dims.length >= 2 || vals.length >= 2) {
            const valueFields = vals.map(v => v);
            const valueFunctions = vals.map(() =>
                aggDef.func === 'sum' ? 'Sum' :
                aggDef.func === 'avg' ? 'Average' :
                aggDef.func === 'count' ? 'Count' :
                aggDef.func === 'max' ? 'Max' : 'Min');
            return {
                tool: 'excel_create_pivot_table',
                params: {
                    rowFields: dims,
                    valueFields: valueFields,
                    valueFunctions: valueFunctions
                },
                label: `按${dims.join('+')}${aggDef.label}${vals.join('、')}（透视表）`,
                type: 'pivot_aggregate',
                text: text
            };
        }

        // 场景 1/4/5/6/7：单维度 + 单指标 → excel_aggregate
        // 单维度 + 无明确指标（如"按产品分类汇总"）也返回 excel_aggregate：
        // valueColumn 留空，由下游"参数不足→模型补全"机制推断指标列，避免聚合被静默跳过。
        if (dims.length >= 1) {
            const v = vals[0] || '';
            return {
                tool: 'excel_aggregate',
                params: {
                    dimColumn: dims[0],
                    valueColumn: v,
                    aggFunc: aggDef.func,
                    chartType: chartTypeEnum || (aggDef.func === 'count' ? 'xlBarClustered' : 'xlColumnClustered'),
                    // 后端纯聚合默认不落盘；显式落盘"聚合结果"表，供后续图表步骤引用
                    targetSheet: '聚合结果'
                },
                label: `按${dims[0]}${aggDef.label}${v}`,
                type: 'simple_aggregate',
                text: text
            };
        }

        return null;
    }

    /**
     * 解析聚合意图，生成结构化任务
     */
    function splitAggregateTasks(text) {
        if (!text) return null;
        const det = D();
        if (!det || !det.detectAggregateIntent(text)) return null;

        const AND_PH = '\u0001';
        const protectedText = String(text)
            .replace(/(求)和/g, '$1' + AND_PH)
            .replace(/(按|依照|根据)([^，,；;。、\s]{1,4})和(?=[^，,；;。、\s]{1,4})/g, '$1$2' + AND_PH);

        const parts = protectedText
            .split(/[，,；;。、]|\s*(?:并|且|和|以及|同时)\s*/)
            .map(s => s.trim().split(AND_PH).join('和'))
            .filter(Boolean);

        if (parts.length <= 1) {
            const task = buildSingleAggregateTask(text);
            return task ? [task] : null;
        }

        const tasks = [];
        for (const p of parts) {
            if (!det.detectAggregateIntent(p)) continue;
            const ct = det.detectChartType(p);
            const task = buildSingleAggregateTask(p, ct ? ct.enum : null);
            if (task) tasks.push(task);
        }
        return tasks.length > 0 ? tasks : null;
    }

    /**
     * 根据聚合任务类型构建强约束 System Prompt
     */
    function buildAggregateSystemPrompt(task) {
        const toolName = task.tool;
        const params = task.params;
        const paramDesc = Object.entries(params)
            .map(([k, v]) => {
                if (Array.isArray(v)) return `${k}: ${v.join(', ')}`;
                return `${k}: ${v}`;
            })
            .join(', ');

        if (task.type === 'pivot_aggregate') {
            return `你是专注于 Excel 数据透视表的助手。请使用 excel_create_pivot_table 工具创建数据透视表。` +
                `参数要求：${paramDesc}。` +
                `源数据使用当前 Excel 选区。创建完成后立即结束，不要创建其他图表或调用其他工具。`;
        }
        if (task.type === 'filtered_aggregate') {
            return `你是专注于 Excel 分类汇总的助手。请使用 ${task.tool} 工具。` +
                `参数要求：filterColumn=${params.filterColumn}, filterValue=${params.filterValue}, dimColumns=${(params.dimColumns || []).join('、')}, valueColumns=${(params.valueColumns || []).join('、')}, aggFunc=${params.aggFunc}。` +
                `源数据使用当前 Excel 选区（已应用筛选）。创建完成后立即结束。`;
        }
        return `你是专注于 Excel 分类汇总的助手。请使用 excel_aggregate 工具。` +
            `参数要求：dimColumn="${params.dimColumn}", valueColumn="${params.valueColumn}", aggFunc=${params.aggFunc}, chartType=${params.chartType}, targetSheet="${params.targetSheet || '聚合结果'}"。` +
            `源数据使用当前 Excel 选区。直接调用工具，参数必须与上述完全一致。创建完成后立即结束，不要调用其他工具。`;
    }

    function buildAggregateUserText(task, originalText) {
        const params = task.params;
        if (task.type === 'pivot_aggregate') {
            return `请按以下要求创建透视表：行字段=${(params.rowFields || []).join('、')}，值字段=${(params.valueFields || []).join('、')}，聚合函数=${(params.valueFunctions || []).join('、')}。`;
        }
        if (task.type === 'filtered_aggregate') {
            return `请在筛选${params.filterText || params.filterValue || ''}后的数据上，按 ${(params.dimColumns || []).join('、')} ${params.aggFunc} ${(params.valueColumns || []).join('、')}。`;
        }
        return `请按 ${params.dimColumn} ${params.aggFunc} ${params.valueColumn}，生成${params.chartType === 'xlPie' ? '饼图' : params.chartType === 'xlBarClustered' ? '条形图' : '柱状图'}。`;
    }

    /**
     * R2 收敛：图表/聚合"内置确定性任务"判定的单一实现（唯一一处定义，供 orchestrator 与 text-pipeline 共同消费）。
     * 同时封装三条互补策略，覆盖不同路由时刻，仅在"一处"统一定义、不改任意一条策略的既有语义：
     *   - splitChartTasks >= 2      → 真正拆分为多个图表子任务
     *   - splitAggregateTasks >= 1  → 命中聚合任务
     *   - countChartWords >= 2     → 词级多图表（兼容"…和…"被拆分合并成单任务、但字面有 ≥2 图表的场景）
     * 返回的 builtin(builtin === hasAggregate || multiChartBySplit || multiChartByCount) 供路由判定。
     */
    function detectBuiltinIntent(text) {
        const userMsg = String(text || '');
        const plannerDet = D();
        const chartTasks = splitChartTasks(userMsg);
        const aggregateTasks = splitAggregateTasks(userMsg);
        const chartCount = (plannerDet && typeof plannerDet.countChartWords === 'function') ? plannerDet.countChartWords(userMsg) : 0;
        const multiChartBySplit = Array.isArray(chartTasks) && chartTasks.length >= 2;
        const hasAggregate = Array.isArray(aggregateTasks) && aggregateTasks.length >= 1;
        const multiChartByCount = chartCount >= 2;
        return {
            chartTasks,
            aggregateTasks,
            chartCount,
            multiChartBySplit,
            multiChartByCount,
            hasAggregate,
            builtin: hasAggregate || multiChartBySplit || multiChartByCount
        };
    }

    App.ChatV2.taskPlanner = {
        CHART_TASK_SYSTEM_PROMPT,
        splitChartTasks,
        splitAggregateTasks,
        buildSingleAggregateTask,
        buildAggregateSystemPrompt,
        buildAggregateUserText,
        detectBuiltinIntent
    };

    ChatLog.info?.('CHAT-V2', '✅ taskPlanner loaded');
})();
