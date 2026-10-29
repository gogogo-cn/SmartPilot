﻿/**
 * ============================================================
 * chat-v2/services/StepExecutor.js — 步骤执行器
 * ============================================================
 * 职责：根据步骤类型（strategy）选择不同的执行方式：
 *   - direct: 直接调用 toolRegistry.invokeTracked() 执行MCP工具
 *   - agent: 在子agent-loop中执行LLM推理（用于需要思考的步骤）
 *   - pipeline: 通过 pipelineRegistry 路由到子管道执行
 *
 * 设计：
 *   - 单一职责：只负责执行步骤，不负责状态追踪
 *   - 错误隔离：步骤执行失败不影响后续步骤（支持重试）
 *   - 可扩展：新增策略只需新增 case
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;
    const MAX_RETRIES = 2;

    /**
     * 执行单个步骤
     * @param {Object} step - 步骤定义 {tool, name, description, strategy, params, ...}
     * @param {number} stepIndex - 步骤索引（用于日志）
     * @param {Object} ctx - 执行上下文
     *   - toolRegistry: 工具注册表
     *   - agentLoop: Agent循环（用于agent策略）
     *   - pipelineRegistry: 管道注册表（用于pipeline策略）
     *   - sessionStore: Session存储
     *   - turn: 当前turn
     *   - signal: AbortSignal
     * @returns {Promise<{ok: boolean, result: any, error: any, iterations: number}>}
     */
    async function executeStep(step, stepIndex, ctx) {
        const { toolRegistry, agentLoop, pipelineRegistry, sessionStore, turn, signal, results, steps: planSteps } = ctx;
        const strategy = step.strategy || 'direct';
        // 优先用 toolName（TaskPlanner 保证已映射为 excel_* 工具名），再 tool，最后 name。
        // 之前用 step.tool || step.name 会忽略 toolName，导致 name 不含聚合/图表关键词时
        // 前端聚合/格式化/建图等判定失效（如 name="分析哪些产品卖得好" 匹配不到"汇总/聚合"）。
        const toolName = step.toolName || step.tool || step.name;

        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 执行步骤 ${stepIndex + 1}, 策略=${strategy}, 工具=${toolName}`);

        // 【步骤上下文填充】从之前步骤执行结果中提取的上下文（如新表名）填充当前步骤缺失参数
        // 独立于 MCP 响应解析，更可靠
        const stepContext = ctx.stepContext;
        if (stepContext && Object.keys(stepContext).length > 0) {
            if (!step.params) step.params = {};
            try {
                fillStepContext(step, stepContext);
            } catch (fillErr) {
                ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 步骤上下文填充异常:', fillErr?.message);
            }
        }

        let retries = 0;
        let lastError = null;

        while (retries <= MAX_RETRIES) {
            try {
                if (signal?.aborted) {
                    return { ok: false, error: new Error('执行已取消'), iterations: retries + 1 };
                }

                // 【预执行】从之前步骤的结果数据中自动填充列名（解决 LLM 猜测列名问题）
                if (strategy === 'direct' && results && results.length > 0) {
                    if (!step.params) step.params = {};
                    try {
                        prepareStepParams(step, results, turn);
                    } catch (prepErr) {
                        ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 参数预填充异常:', prepErr?.message);
                    }
                }

                let result;

                // 【直接修正】excel_format_range 的 sheet 和 range 参数从结果中提取
                // 放在这里（而非 executeDirect）是因为 executeStep 有 results 和 ctx 上下文
                const isFormatRange = toolName === 'excel_format_range' || /格式化|format/i.test(toolName || '');
                if (isFormatRange && results && results.length > 0) {
                    if (!step.params) step.params = {};
                    // 遍历所有前序结果，提取目标工作表名和数据范围
                    let targetSheet = null;
                    let dataRange = null;
                    for (const prevResult of results) {
                        if (!prevResult || !prevResult.ok || !prevResult.result) continue;
                        // 提取实际工具数据（兼容字符串与对象）
                        const toolData = extractToolData(prevResult.result);
                        if (toolData && toolData.worksheet && !targetSheet) {
                            targetSheet = toolData.worksheet;
                        }
                        // 提取返回文本（兼容字符串 / { content: [...] } / { result: string }）
                        let text = '';
                        const ir = prevResult.result;
                        if (typeof ir === 'string') {
                            text = ir;
                        } else if (ir && typeof ir === 'object') {
                            if (typeof ir.result === 'string') {
                                text = ir.result;
                            } else if (ir.result && Array.isArray(ir.result.content) && ir.result.content.length > 0) {
                                text = ir.result.content[0]?.text || '';
                            } else if (Array.isArray(ir.content) && ir.content.length > 0) {
                                text = ir.content[0]?.text || '';
                            }
                        }
                        if (text) {
                            // 提取目标工作表：→ 汇总!A1 或 → 汇总!A1:D15
                            const arrowMatch = text.match(/→\s*([^!]+)!/);
                            if (arrowMatch && !targetSheet) {
                                targetSheet = arrowMatch[1].trim();
                                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从复制结果提取目标表: '${targetSheet}'`);
                            }
                            // 提取数据范围：Sheet1!A1:D15 或 Sheet1!A1（源范围，复制到目标表 A1 起始）
                            const rangeMatch = text.match(/[^!→]+!([A-Z]+\d+(?::[A-Z]+\d+)?)/);
                            if (rangeMatch && !dataRange) {
                                dataRange = rangeMatch[1];
                                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从复制结果提取数据范围: '${dataRange}'`);
                            }
                        }
                    }
                    // 填充 sheet：目标表优先（汇总），否则用 stepContext 兜底
                    if (targetSheet) {
                        step.params.sheet = targetSheet;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 sheet='${targetSheet}'（格式化目标表）`);
                    } else if (stepContext && (stepContext.lastDestinationSheet || stepContext.lastCreatedSheet)) {
                        step.params.sheet = stepContext.lastDestinationSheet || stepContext.lastCreatedSheet;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 sheet='${step.params.sheet}'（stepContext 兜底）`);
                    }
                    // 填充 range：从复制结果提取的数据范围
                    if (dataRange) {
                        step.params.range = dataRange;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 range='${dataRange}'（格式化目标表数据）`);
                    }
                }

                // 【直接修正】excel_create_chart 的 sheet 和 dataRange 参数从聚合结果中提取
                // 聚合步骤已把结果写入"聚合_XX"工作表，建图步骤应直接引用该表的数据范围
                const isCreateChart = toolName === 'excel_create_chart' || /图表|柱状图|折线图|饼图|面积图|散点图|雷达图|环形图|画图|chart/i.test(toolName || '');
                if (isCreateChart && results && results.length > 0) {
                    if (!step.params) step.params = {};
                    let chartSheet = null;
                    let chartRange = null;
                    let prevChartType = null;
                    // 【多图表：一对一取源】把当前图表步骤与"它自己前面的那个聚合步骤"配对，
                    // 避免第二张图复用第一张图的数据源（原实现固定取 results 里第一份聚合结果）。
                    // 配对算法：按同样的规则先算出前面的图表步骤都认领了哪个聚合步骤，
                    // 当前图表步骤取"最近的、尚未被认领的"聚合步骤。
                    let pairedIdx = -1;
                    if (Array.isArray(planSteps)) {
                        const stepToolName = (s) => String((s && (s.toolName || s.tool || s.name)) || '');
                        const isAgg = (s) => /aggregate|date_group|汇总|聚合|统计|分组|求和|平均|计数/.test(stepToolName(s));
                        const isChart = (s) => /create_chart|图表|柱状图|折线图|饼图|环形图|面积图|散点图|雷达图|画图/.test(stepToolName(s));
                        const claimed = new Set();
                        for (let j = 0; j < stepIndex; j++) {
                            if (!isChart(planSteps[j])) continue;
                            for (let k = j - 1; k >= 0; k--) {
                                if (isAgg(planSteps[k])) { claimed.add(k); break; }
                            }
                        }
                        for (let j = stepIndex - 1; j >= 0; j--) {
                            if (isAgg(planSteps[j]) && !claimed.has(j)) { pairedIdx = j; break; }
                        }
                    }
                    const sourceResults = (pairedIdx >= 0 && results[pairedIdx]) ? [results[pairedIdx]] : results;
                    if (pairedIdx >= 0) {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 图表步骤与聚合步骤 ${pairedIdx + 1} 配对取源`);
                    }
                    for (const prevResult of sourceResults) {
                        if (!prevResult || !prevResult.ok || !prevResult.result) continue;
                        const toolData = extractToolData(prevResult.result);
                        if (toolData && toolData.resultSheet && !chartSheet) {
                            chartSheet = toolData.resultSheet;
                        }
                        if (toolData && toolData.resultRange && !chartRange) {
                            chartRange = toolData.resultRange;
                        }
                        // 记录前置聚合步骤生成的图表类型（用于冗余检测）
                        if (toolData && toolData.chartType && toolData.resultSheet && !prevChartType) {
                            prevChartType = toolData.chartType;
                        }
                    }
                    // 【兜底落盘】前置聚合步骤只做了内存计算（无 resultSheet —— 图表类型规划在
                    // 独立的图表步骤上而非聚合步上时，聚合步不会写结果表）时，把配对的聚合
                    // 结果落盘为新表供图表引用。否则图表步骤缺 dataRange/sheet 会触发
                    // needsInput 反问硬失败（真实 bug："使用饼图按部门展示销售额" 3 步计划，
                    // 第 3 步报"缺少必要参数：数据范围"）。
                    if (!chartSheet) {
                        for (let ri = sourceResults.length - 1; ri >= 0; ri--) {
                            const prevAgg = sourceResults[ri];
                            if (!prevAgg || !prevAgg.ok || !prevAgg.result) continue;
                            const tdAgg = extractToolData(prevAgg.result);
                            if (!tdAgg || !Array.isArray(tdAgg.data) || tdAgg.data.length === 0) continue;
                            if (tdAgg.aggregated !== true && tdAgg.dateGrouped !== true) continue;
                            try {
                                // 表头以数据行实际键为准（VB 聚合返回的 headers 是源表列，不是聚合结果列）
                                const firstRow = tdAgg.data[0] || {};
                                const keys = Object.keys(firstRow);
                                const dimKey = (tdAgg.dimColumn && keys.includes(tdAgg.dimColumn)) ? tdAgg.dimColumn : (keys[0] || '类别');
                                const valKey = (tdAgg.valueColumn && keys.includes(tdAgg.valueColumn) && tdAgg.valueColumn !== dimKey)
                                    ? tdAgg.valueColumn : (keys.find(k => k !== dimKey) || '数量');
                                const values2d = [[dimKey, valKey]].concat(tdAgg.data.map(r => [
                                    (r && r[dimKey] !== undefined && r[dimKey] !== null) ? r[dimKey] : '',
                                    (r && r[valKey] !== undefined && r[valKey] !== null) ? r[valKey] : ''
                                ]));
                                // 表名带步骤索引，保证同一轮里多张图各自的落盘表互不覆盖
                                const newSheet = '聚合_' + dimKey + '_' + stepIndex + '_' + (Date.now() % 10000);
                                await toolRegistry.invokeTracked(turn, 'excel_write_range', {
                                    sheet: newSheet,
                                    range: 'A1',
                                    values: JSON.stringify(values2d)
                                }, { hidden: true });
                                chartSheet = newSheet;
                                chartRange = 'A1:' + getExcelColumnName(2) + (tdAgg.data.length + 1);
                                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 前置聚合仅有内存结果，已落盘 '${newSheet}' (${chartRange}) 供图表引用`);
                            } catch (landErr) {
                                ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 聚合结果兜底落盘失败:', landErr?.message);
                            }
                            break;  // 只取最近一份聚合结果
                        }
                    }
                    // 冗余检测：前置聚合步骤已生成同类型图表时，跳过本次建图，避免重复图表
                    if (prevChartType) {
                        const stepChartType = (step.params && step.params.chartType) || 'xlColumnClustered';
                        const normPrev = String(prevChartType).toLowerCase().replace(/[^a-z]/g, '');
                        const normStep = String(stepChartType).toLowerCase().replace(/[^a-z]/g, '');
                        if (normPrev === normStep) {
                            // 将技术枚举名翻译为用户友好的中文名称
                            const CHART_TYPE_CN = {
                                'xlPie': '饼图', 'xlDoughnut': '圆环图',
                                'xlColumnClustered': '柱状图', 'xlColumnStacked': '堆叠柱状图',
                                'xlColumnStacked100': '百分比柱状图',
                                'xlBarClustered': '条形图', 'xlBarStacked': '堆叠条形图',
                                'xlBarStacked100': '百分比条形图',
                                'xlLine': '折线图', 'xlLineMarkers': '折线带点图',
                                'xlArea': '面积图', 'xlXYScatter': '散点图',
                                'xlRadar': '雷达图', 'xlRadarFilled': '填充雷达图',
                                'xlBubble': '气泡图', 'xlSunburst': '旭日图',
                                'xlTreemap': '树状图', 'xlHistogram': '直方图',
                                'xlBoxwhisker': '箱线图', 'xlFunnel': '漏斗图',
                                'xlWaterfall': '瀑布图'
                            };
                            const typeLabel = CHART_TYPE_CN[prevChartType] || '图表';
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 前置聚合步骤已生成 ${typeLabel}，跳过建图步骤`);
                            return {
                                ok: true,
                                result: { skipped: true, message: `上一步已生成${typeLabel}，无需重复创建` },
                                error: null,
                                iterations: 1
                            };
                        }
                    }
                    if (chartSheet) {
                        step.params.sheet = chartSheet;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 sheet='${chartSheet}'（图表数据源表）`);
                    }
                    if (chartRange) {
                        step.params.dataRange = chartRange;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 dataRange='${chartRange}'（图表数据源）`);
                    }
                }

                // 【多列混合表：图表改走内存数组模式（dataValues）】
                // range 模式下 Excel 会把源区域每一列都绑成一个系列，文本列（地区/部门/产品/销售员…）
                // 被当成数值系列 → 图上出现"文本当数值"的乱码数据标签（实测：对整表 A1:I101 建柱状图）。
                // 这里只保留"首列分类 + 数值列"组装 dataValues，交给后端已有的内存数组模式（后端只绑数值列）。
                // 数据行来源：① 图表步骤已带 dataRange → 读该区域；② 未带 dataRange（依赖"选区自动注入"）
                // → 复用前置步骤已读到的数据行（否则选区一注入就是整表，文本列又被全部绑成系列）。
                // 读失败 / 无数值列 / 无可剔除的文本列时，保持原参数不动。
                if (isCreateChart && step.params && !step.params.dataValues) {
                    const rngAddr = String(step.params.dataRange || '');
                    let chartRows = null;
                    let chartSheetForRead = step.params.sheet || '';
                    try {
                        if (rngAddr && rangeColumnCount(rngAddr) > 2) {
                            const colonIdx = rngAddr.lastIndexOf('!');
                            chartSheetForRead = colonIdx >= 0
                                ? rngAddr.slice(0, colonIdx).replace(/^'|'$/g, '')
                                : (step.params.sheet || '');
                            const addrForRead = colonIdx >= 0 ? rngAddr.slice(colonIdx + 1) : rngAddr;
                            const readOut = await toolRegistry.invokeTracked(turn, 'excel_read_range', {
                                sheet: chartSheetForRead, range: addrForRead, hasHeaders: true, autoProbe: false
                            }, { hidden: true, round: 1 });
                            chartRows = findSourceRows([{ ok: true, result: readOut }]);
                        } else if (!rngAddr) {
                            chartRows = findSourceRows(results);
                        }
                    } catch (dvErr) {
                        ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 图表数据源转内存模式失败（保持原参数不变）:', dvErr?.message);
                    }
                    if (chartRows && chartRows.length > 0) {
                        const keys = Object.keys(chartRows[0]);
                        const catKey = keys[0];
                        const numKeys = keys.slice(1).filter(k => chartRows.some(r => isNumericCellValue(r && r[k])));
                        const droppedTextCols = keys.length - 1 - numKeys.length;
                        if (catKey && numKeys.length > 0 && droppedTextCols > 0) {
                            const arr = [[catKey].concat(numKeys)];
                            for (const r of chartRows) {
                                arr.push([(r && r[catKey] != null) ? r[catKey] : '']
                                    .concat(numKeys.map(k => (r && r[k] != null) ? r[k] : '')));
                            }
                            step.params.dataValues = JSON.stringify(arr);
                            delete step.params.dataRange;
                            if (!step.params.sheet && chartSheetForRead) step.params.sheet = chartSheetForRead;
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 图表改走内存模式：分类列=${catKey}，数值列=${numKeys.join('/')}，剔除文本列 ${droppedTextCols} 个`);
                        }
                    }
                }

                // 【前端聚合】聚合是计算操作，绝不在 Excel 工作表中实现。
                // 始终优先使用前端内存聚合（clientAggregate），不碰 Excel。
                // 当需要图表数据源时，在内存聚合完成后把结果写入新表（仅此一次写入），
                // 图表步骤直接引用新表。聚合本身不写入、不覆盖原数据。
                const isAggregateTool = /aggregate|汇总|聚合|统计|分组|求和|平均|计数/.test(String(toolName));
                const chartRequested = !!(step.params && step.params.chartType);
                if (isAggregateTool && strategy === 'direct' && results && results.length > 0) {
                    const sourceRows = findSourceRows(results);
                    if (sourceRows && sourceRows.length > 0) {
                        const aggResult = clientAggregate(sourceRows, step.params || {});
                        if (aggResult) {
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 前端聚合完成: 按「${aggResult.dimColumn}」分组，${aggResult.groupCount} 组（内存计算，不操作 Excel）`);

                            // 如果需要图表，把聚合结果写入新表供图表引用（这是写入操作，不是聚合操作）
                            if (chartRequested) {
                                try {
                                    const headers = aggResult.headers || [];
                                    const data = aggResult.data || [];
                                    const values2d = [headers, ...data.map(r => headers.map(h => r[h] != null ? r[h] : ''))];
                                    const sheetName = '聚合_' + (aggResult.dimColumn || '结果') + '_' + (Date.now() % 100000);
                                    const colCount = headers.length;
                                    const rowCount = data.length + 1; // +1 for header
                                    const endCol = getExcelColumnName(colCount);
                                    await toolRegistry.invokeTracked(turn, 'excel_write_range', {
                                        sheet: sheetName,
                                        range: 'A1',
                                        values: JSON.stringify(values2d)
                                    }, { hidden: true });
                                    aggResult.resultSheet = sheetName;
                                    aggResult.resultRange = `A1:${endCol}${rowCount}`;
                                    aggResult.chartDataReady = true;
                                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 聚合结果已写入新表 '${sheetName}' (${aggResult.resultRange}) 供图表引用`);
                                } catch (writeErr) {
                                    ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 写入聚合结果到图表数据源失败:', writeErr?.message);
                                }
                            }

                            return {
                                ok: true,
                                result: aggResult,
                                error: null,
                                iterations: retries + 1
                            };
                        }
                    }
                }

                // 【前端日期分组】日期分组步骤且数据在前端可用时，在面板中计算并显示。
                const isDateGroupTool = /date_group|日期分组|时间分组/.test(String(toolName));
                if (isDateGroupTool && strategy === 'direct' && ctx.enableClientAggregate &&
                    results && results.length > 0) {
                    const sourceRows = findSourceRows(results);
                    if (sourceRows && sourceRows.length > 0) {
                        const dgResult = clientDateGroup(sourceRows, step.params || {});
                        if (dgResult) {
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 前端日期分组完成: 按「${dgResult.dateColumn}」${dgResult.groupBy}分组，${dgResult.groupCount} 组（未生成 Excel 表）`);
                            return {
                                ok: true,
                                result: dgResult,
                                error: null,
                                iterations: retries + 1
                            };
                        }
                    }
                }

                // 【写入前清理】写入步骤写入管道创建的结果表时，先清空整表，
                // 防止上一次运行残留的公式/内容导致 #N/A 或脏数据（用户反馈结果表出现大量 #N/A）。
                // 仅清理管道创建的表（lastCreatedSheet/lastDestinationSheet），避免误清源数据。
                const isWriteTool = toolName === 'excel_write_range' || /write|写入|填充|填入/.test(String(toolName));
                if (isWriteTool && strategy === 'direct' && step.params && step.params.sheet && stepContext) {
                    const isPipelineSheet = (step.params.sheet === stepContext.lastCreatedSheet) ||
                                            (step.params.sheet === stepContext.lastDestinationSheet);
                    if (isPipelineSheet) {
                        try {
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 写入前清空结果表: '${step.params.sheet}'`);
                            await toolRegistry.invokeTracked(turn, 'excel_clear_sheet', { sheet: step.params.sheet }, { hidden: true });
                        } catch (clearErr) {
                            ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 清空结果表失败（继续写入）:', clearErr?.message);
                        }
                    }
                }

                // =================================================================
                // 【原数据保护区 · 极其重要】禁止任何写入/格式化/删除步骤覆盖用户原始选区
                // 聚合（aggregate）是纯内存计算，不在此列——绝不在 Excel 中实现聚合。
                // 原则：目标表=原数据表 且 目标范围与原始选区重叠/未指定 → 自动偏移到安全位置
                //       格式化步骤若无明确目标表（指向原数据表） → 跳过格式化避免美化原表格
                // =================================================================
                if (strategy === 'direct' && step.params && turn && turn.context && turn.context.selection) {
                    const sel = turn.context.selection;
                    const origSheet = sel.sheetName || sel.worksheet || '';
                    const origRange = sel.range || sel.address || '';

                    if (origSheet) {
                        const isDestructive =
                            toolName === 'excel_write_range' || /write|写入|填充|填入/.test(String(toolName)) ||
                            toolName === 'excel_date_group' || /date_group|日期分组|时间分组/.test(String(toolName)) ||
                            toolName === 'excel_format_range' || /format|格式化|美化/.test(String(toolName)) ||
                            toolName === 'excel_copy_range' || /copy|复制|拷贝/.test(String(toolName)) ||
                            toolName === 'excel_delete_range' || /delete|删除|清空/.test(String(toolName));

                        if (isDestructive) {
                            const targetSheet = step.params.sheet || step.params.sheetName || '';
                            // 如果目标表为空（默认当前活动表），或明确指向原始数据表 → 需要检查
                            const hitsOrigSheet = !targetSheet || targetSheet === origSheet;
                            const hasSafeTarget =
                                (stepContext && stepContext.lastCreatedSheet && targetSheet === stepContext.lastCreatedSheet) ||
                                (stepContext && stepContext.lastDestinationSheet && targetSheet === stepContext.lastDestinationSheet);

                            if (hitsOrigSheet && !hasSafeTarget) {
                                // 解析原始选区范围，计算安全偏移位置（右侧 2 列或下方 3 行）
                                function parseRange(rangeStr) {
                                    if (!rangeStr) return null;
                                    const m = rangeStr.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
                                    if (!m) return null;
                                    const colToIdx = (c) => {
                                        let idx = 0;
                                        for (let i = 0; i < c.length; i++) idx = idx * 26 + (c.charCodeAt(i) - 64);
                                        return idx;
                                    };
                                    return {
                                        col1: m[1], col1Idx: colToIdx(m[1]),
                                        row1: parseInt(m[2], 10),
                                        col2: m[3] || m[1], col2Idx: colToIdx(m[3] || m[1]),
                                        row2: m[4] ? parseInt(m[4], 10) : parseInt(m[2], 10)
                                    };
                                }
                                const idxToCol = (idx) => {
                                    let s = '';
                                    while (idx > 0) {
                                        const mod = (idx - 1) % 26;
                                        s = String.fromCharCode(65 + mod) + s;
                                        idx = Math.floor((idx - 1) / 26);
                                    }
                                    return s || 'A';
                                };

                                // 格式化步骤额外严格：如果没有之前步骤创建的结果表继承，直接跳过避免误美化原表
                                const isFormatTool = toolName === 'excel_format_range' || /format|格式化|美化/.test(String(toolName));
                                if (isFormatTool && !stepContext?.lastCreatedSheet && !stepContext?.lastDestinationSheet) {
                                    ChatLog.warn?.('MULTI-STEP', `[StepExecutor] 【原数据保护】格式化步骤未指定结果表，跳过以避免美化原数据（orig=${origSheet}）`);
                                    return {
                                        ok: true,
                                        result: { ok: true, skipped: true, reason: '避免格式化原数据' },
                                        error: null,
                                        iterations: retries + 1
                                    };
                                }

                                // 生成安全写入位置：原始选区右侧偏移 2 列，从第 1 行开始
                                const orig = parseRange(origRange);
                                let safeStartCell = 'A1';
                                if (orig) {
                                    const safeColIdx = orig.col2Idx + 2; // 右侧偏移 2 列
                                    safeStartCell = idxToCol(safeColIdx) + '1';
                                }

                                // 写入/日期分组 类工具：把目标写到安全偏移位置或新建结果表
                                // 注：聚合（excel_aggregate）不在此列——聚合是纯计算，走前端 clientAggregate，不操作 Excel
                                const needsSafeRedirect =
                                    toolName === 'excel_write_range' || /write|写入|填充|填入/.test(String(toolName)) ||
                                    toolName === 'excel_date_group' || /date_group|日期分组|时间分组/.test(String(toolName));

                                if (needsSafeRedirect) {
                                    // 优先：结果表存在则继承结果表（由 VB 聚合返回 resultSheet 字段）
                                    // 否则：用安全起始单元格写入原数据表
                                    if (!step.params.range && !step.params.startCell) {
                                        step.params.range = safeStartCell;
                                        ChatLog.warn?.('MULTI-STEP',
                                            `[StepExecutor] 【原数据保护】目标指向原数据表，自动偏移写入位置: A1 → ${safeStartCell}（origRange=${origRange}）`);
                                    }
                                }

                                // 删除类工具：如果目标在原数据区内，直接拒绝
                                const isDeleteTool = toolName === 'excel_delete_range' || /delete|删除|清空/.test(String(toolName));
                                if (isDeleteTool) {
                                    const dr = step.params.range || origRange;
                                    if (dr === origRange) {
                                        ChatLog.warn?.('MULTI-STEP', `[StepExecutor] 【原数据保护】删除操作与原始选区重叠，拒绝执行（origRange=${origRange}）`);
                                        return { ok: false, error: new Error('原数据保护：禁止删除用户原始数据选区'), iterations: retries + 1 };
                                    }
                                }
                            }
                        }
                    }
                }

                switch (strategy) {
                    case 'direct':
                        result = await executeDirect(step, toolRegistry, sessionStore, turn, stepContext);
                        break;
                    case 'agent':
                        result = await executeAgent(step, agentLoop, ctx);
                        break;
                    case 'pipeline':
                        result = await executePipeline(step, pipelineRegistry, ctx);
                        break;
                    default:
                        throw new Error(`未知执行策略: ${strategy}`);
                }

                // 执行成功
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤 ${stepIndex + 1} 完成`);
                return {
                    ok: true,
                    result,
                    error: null,
                    iterations: retries + 1
                };
            } catch (e) {
                lastError = e;
                retries++;
                ChatLog.warn?.('MULTI-STEP', `[StepExecutor] 步骤 ${stepIndex + 1} 失败 (重试 ${retries}/${MAX_RETRIES}):`, e?.message);

                if (retries > MAX_RETRIES) {
                    ChatLog.error?.('MULTI-STEP', `[StepExecutor] 步骤 ${stepIndex + 1} 最终失败`);
                    break;
                }

                // 首次失败时，尝试从之前步骤的结果数据中自动修正参数（如列名）
                if (strategy === 'direct' && retries === 1 && results && results.length > 0) {
                    try {
                        const fixed = tryFixParamsFromPreviousResult(step, e, results);
                        if (fixed) {
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤 ${stepIndex + 1} 参数已自动修正，准备重试`);
                        }
                    } catch (fixErr) {
                        ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 参数修正异常:', fixErr?.message);
                    }
                }

                // 等待一小段再重试（避免立即重试导致的问题）
                await new Promise(resolve => setTimeout(resolve, 500));
            }
        }

        // 重试用完仍失败
        return {
            ok: false,
            result: null,
            error: lastError,
            iterations: retries
        };
    }

    /**
     * direct 策略：直接调用MCP工具
     */
    async function executeDirect(step, toolRegistry, sessionStore, turn, stepContext) {
        const toolName = step.toolName || step.tool;
        let params = step.params || {};

        if (!toolName) {
            throw new Error('direct 策略缺少 tool 字段');
        }

        // 检查工具是否存在（toolRegistry.get 返回 undefined 表示不存在）
        // 先尝试 alias 翻译（LLM 常编不存在的工具名如 excel_new_sheet → excel_create_sheet）
        let _realTool = toolRegistry.get(toolName);
        if (!_realTool && typeof toolRegistry._resolveToolAlias === 'function') {
            const _alias = toolRegistry._resolveToolAlias(toolName);
            if (_alias) {
                ChatLog.warn?.('StepExec', `[StepExecutor] LLM 编造工具名 "${toolName}" → 翻译为 "${_alias}"`);
                toolName = _alias;
                _realTool = toolRegistry.get(_alias);
            }
        }
        if (!_realTool) {
            throw new Error(`找不到工具: ${toolName}`);
        }

        // 【直接修正】excel_format_range 的 sheet 参数从步骤描述中提取
        // 避免 fillSelectionArgs 用 turn.context.activeSheet（原始表）覆盖
        const isFormatRange = toolName === 'excel_format_range' || /格式化|format/i.test(toolName || '');
        if (isFormatRange && (!params.sheet || params.sheet === 'Sheet1')) {
            // 从 step.description 提取（如 "格式化 '汇总' 表格"）
            const desc = step.description || step.name || '';
            ChatLog.info?.('MULTI-STEP', `[StepExecutor] excel_format_range 尝试从描述提取 sheet: desc='${desc}'`);
            // 支持单引号、双引号、中文引号
            const wsMatch = desc.match(/['"''""]([^'""]+)['"''""]/);
            if (wsMatch) {
                params.sheet = wsMatch[1];
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 excel_format_range sheet='${params.sheet}'（从步骤描述提取）`);
            }
        }

        // 【直接修正】excel_write_range 的 sheet 参数与上下文校验
        // 当 LLM 规划的 sheet 名称与之前创建的不一致时，用上下文的正确名称覆盖
        const isWriteRange = toolName === 'excel_write_range' || /write|写入|填充|填入/i.test(toolName || '');
        if (isWriteRange && stepContext && stepContext.lastCreatedSheet) {
            const plannedSheet = params.sheet;
            const contextSheet = stepContext.lastCreatedSheet;
            // 如果规划的 sheet 为空、为默认表、或与上下文不一致，则修正
            if (!plannedSheet || plannedSheet === 'Sheet1' || plannedSheet !== contextSheet) {
                params.sheet = contextSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 强制修正 excel_write_range sheet='${contextSheet}'（来自步骤上下文，原值='${plannedSheet}'）`);
            }
        }

        // 【关键】excel_write_range 的 range 参数必须是起始单元格（如 A1），不能是完整范围（如 A1:B6）
        // VB 后端的 GetEndCell 函数只解析起始单元格，传入完整范围会导致解析错误（如 A1:B6 → AB + 16 → AC21）
        if (isWriteRange && params.range) {
            const rangeStr = String(params.range);
            const colonIdx = rangeStr.indexOf(':');
            if (colonIdx > 0) {
                const startCell = rangeStr.substring(0, colonIdx).trim();
                params.range = startCell;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 修正 excel_write_range range='${startCell}'（原值含结束位置 '${rangeStr}'，VB 后端只需起始单元格）`);
            }
        }

        // 多步骤管道中，工具卡片由规划面板（plan-step）替代，故使用 hidden: true
        // 避免每个中间步骤都渲染独立工具卡片（导致汇总后残留多余卡片）。
        // invokeTracked 签名: (turn, name, args, options)
        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 调用 invokeTracked: tool=${toolName}, params=${JSON.stringify(params)}`);
        const result = await toolRegistry.invokeTracked(turn, toolName, params, { hidden: true });

        // invokeTracked 返回 { ok, result, error, needsInput? }。
        // 当参数缺失时（needsInput），invokeTracked 仍返回 { ok: false, ... } 而非抛出异常，
        // 此处需显式检查 ok 标志，避免 executeStep 错误地将参数缺失视为"执行成功"。
        if (result && result.ok === false) {
            const errMsg = result.error || (result.needsInput ? '缺少必要参数' : '工具执行失败');
            throw new Error(errMsg);
        }

        return result;
    }

    /**
     * agent 策略：在子agent-loop中执行
     * agentLoop.run 返回 async generator，必须通过 for await...of 迭代消费，
     * 否则 Promise 永不 resolve 导致步骤永久卡在"执行中"。
     * 同时加入超时控制，防止 LLM 响应缓慢或无响应导致永久阻塞。
     */
    async function executeAgent(step, agentLoop, ctx) {
        const { sessionStore, signal } = ctx;
        const prompt = step.prompt || step.description;

        if (!prompt) {
            throw new Error('agent 策略缺少 prompt/description 字段');
        }

        // 构建子消息
        const messages = ctx.messages ? [...ctx.messages] : [];
        messages.push({
            role: 'user',
            content: prompt
        });

        // 本地 AbortController：统一管理超时与外部取消，使 agent-loop 真正被中止
        const localAbort = new AbortController();
        const abortWithReason = (ac, msg) => {
            try { ac.abort(new DOMException(msg, 'AbortError')); }
            catch (_) { try { ac.abort(); } catch (_2) { /* ignore */ } }
        };
        const onOuterAbort = () => abortWithReason(localAbort, '上层已取消本步骤');
        signal?.addEventListener?.('abort', onOuterAbort);

        // agent 步骤超时控制：防止 LLM 响应缓慢或无响应导致永久阻塞
        const AGENT_TIMEOUT_MS = 120000; // 120 秒
        const timeoutHandle = setTimeout(() => abortWithReason(localAbort, 'agent 步骤超过 120 秒，已中止'), AGENT_TIMEOUT_MS);

        try {
            let accumulatedResult = '';
            let toolCalls = [];

            // 【任务套任务】初始化 agent 子步骤占位
            const agentSubSteps = step.subSteps || step.substeps || [
                '初始化 Agent 推理',
                '执行 LLM 推理',
                '整理分析结果'
            ];
            step.subSteps = agentSubSteps.map((name, idx) => ({
                name: typeof name === 'string' ? name : (name.name || name.description || `子步骤 ${idx + 1}`),
                status: idx === 0 ? 'running' : 'pending'
            }));

            // 【子步骤UI实时刷新】初始化后立即触发 ProgressTracker 更新
            notifySubStepUpdate(step, ctx);

            // 迭代 agentLoop.run 返回的 async generator（必须消费，否则 Promise 永不 resolve）
            let subStepIdx = 0;
            for await (const evt of agentLoop.run({
                turn: ctx.turn,
                messages,
                tools: ctx.tools || [],
                signal: localAbort.signal,
                sessionStore,
                silentContent: true,
                parentTurn: ctx.turn
            })) {
                if (evt.type === 'token') {
                    accumulatedResult += evt.delta;
                } else if (evt.type === 'tool_call_result' && evt.ok) {
                    toolCalls.push({ name: evt.toolName, result: evt.result });
                    // 工具调用时更新子步骤状态
                    if (step.subSteps && subStepIdx < step.subSteps.length - 1) {
                        step.subSteps[subStepIdx].status = 'completed';
                        subStepIdx++;
                        step.subSteps[subStepIdx].status = 'running';
                        // 【子步骤UI实时刷新】状态变化后触发 ProgressTracker 更新
                        notifySubStepUpdate(step, ctx);
                    }
                } else if (evt.type === 'error') {
                    throw evt.error || new Error('agent 执行失败');
                } else if (evt.type === 'canceled') {
                    throw new Error('agent 执行已取消');
                }
            }

            // agent 完成后标记所有子步骤为完成。
            // 子步骤推进原依赖"工具调用次数"（每次 tool_call_result 推进一格），但 agent 可能
            // 只发生一次工具调用（如单一 skill_GeneratePPTReport），导致后续子步骤永远停在
            // pending"等待"，UI 上看起来像失败。只要 agent 正常落盘（未抛错/未超时到此行），
            // 即整体成功，所有子步骤统一 completed。
            if (step.subSteps) {
                step.subSteps.forEach((sub) => {
                    sub.status = 'completed';
                });
                // 【子步骤UI实时刷新】完成后触发 ProgressTracker 更新
                notifySubStepUpdate(step, ctx);
            }

            // 超时中止时 agent-loop 会 yield canceled，此处兜底检测
            if (localAbort.signal.aborted) {
                throw new Error('agent 步骤执行超时或已取消');
            }

            // 从 turn 读取最终内容（agent-loop 会把最终结果写入 sessionStore）
            let finalText = accumulatedResult;
            try {
                const currentTurn = sessionStore.getCurrentTurn?.();
                if (currentTurn && currentTurn.id === ctx.turn.id && currentTurn.content) {
                    finalText = currentTurn.content;
                }
            } catch (e) { /* 忽略 */ }

            // 【数据上下文传递】从 agent 内部工具调用中提取结构化数据，供 ResultAggregator 汇总使用
            let extractedData = null;
            for (const tc of toolCalls) {
                const r = tc.result;
                if (r && (r.aggregated === true || r.dateGrouped === true || Array.isArray(r.data))) {
                    extractedData = r;
                    break;
                }
            }

            return {
                type: 'agent',
                text: finalText,
                toolCalls,
                data: extractedData  // 传递结构化数据给 ResultAggregator
            };
        } finally {
            clearTimeout(timeoutHandle);
            signal?.removeEventListener?.('abort', onOuterAbort);
        }
    }

    /**
     * pipeline 策略：路由到子管道执行
     */
    async function executePipeline(step, pipelineRegistry, ctx) {
        const { sessionStore, signal } = ctx;
        const pipelineName = step.pipeline || step.name;

        if (!pipelineName) {
            throw new Error('pipeline 策略缺少 pipeline/name 字段');
        }

        // 获取管道
        const pipeline = pipelineRegistry.getPipeline(pipelineName);
        if (!pipeline) {
            throw new Error(`找不到管道: ${pipelineName}`);
        }

        // 构建子上下文
        // 使用 _isSubPipeline 标记子管道调用（区别于用户后续消息的 _execFollowUp）
        // 使子管道能进入多步骤规划，实现任务套任务的嵌套步骤
        const subTurn = {
            ...ctx.turn,
            userMessage: step.prompt || step.description || ctx.turn.userMessage,
            _parentStep: step,
            _isSubPipeline: true
        };

        // 【任务套任务】初始化子步骤占位，供 ProgressTracker 渲染嵌套面板
        // 子管道执行前预设子步骤列表，UI 可提前展示嵌套结构
        const subStepNames = step.subSteps || step.substeps || [
            `路由到 ${pipelineName}`,
            '执行子管道逻辑',
            '返回结果'
        ];
        step.subSteps = subStepNames.map((name, idx) => ({
            name: typeof name === 'string' ? name : (name.name || name.description || `子步骤 ${idx + 1}`),
            status: idx === 0 ? 'running' : 'pending'
        }));

        // 【子步骤UI实时刷新】初始化后立即触发 ProgressTracker 更新
        notifySubStepUpdate(step, ctx);

        // 执行子管道
        const result = await pipeline.run({
            turn: subTurn,
            signal,
            sessionStore,
            parentContext: ctx
        });

        // 子管道完成后标记所有子步骤为完成
        step.subSteps.forEach((sub, idx) => {
            sub.status = result && !result.error ? 'completed' : 'failed';
        });

        // 【子步骤UI实时刷新】完成后触发 ProgressTracker 更新
        notifySubStepUpdate(step, ctx);

        return result;
    }

    /**
     * 预执行阶段：从之前步骤的结果数据中自动填充列名
     * LLM 的 params 中 valueColumn/dimColumn 可能为空字符串 ""，
     * 此处从实际读取的数据中提取列名并填入
     */
    function prepareStepParams(step, results, turn) {
        const params = step.params;
        if (!params) {
            ChatLog.info?.('MULTI-STEP', '[StepExecutor] prepareStepParams: step.params 为空，跳过');
            return;
        }
        ChatLog.info?.('MULTI-STEP', `[StepExecutor] prepareStepParams: 开始填充, results.length=${results.length}, params=${JSON.stringify(params)}`);

        // 仅当 valueColumn 为空时才删除（让后端自动推断）；用户/LLM 显式指定的数值列必须保留，
        // 否则"按部门分类汇总客户数"会被后端自动推断成"销售额"（金额优先），导致字段名被忽略。
        // 若保留的列名与实际数据不匹配，后端会报"未找到数值列"，由 tryFixParamsFromPreviousResult 兜底修正。
        if ('valueColumn' in params && (!params.valueColumn || !String(params.valueColumn).trim())) {
            delete params.valueColumn;
        }
        if ('valueField' in params && (!params.valueField || !String(params.valueField).trim())) {
            delete params.valueField;
        }

        // 从之前步骤的结果中提取实际数据列名（用于填充 dimColumn）
        // 以及 sheetName/range 等上下文信息
        let headers = [];

        // 【关键】第一步（无前置结果）时，从 turn.context.selection 获取列名
        // 避免静态模式下第一步就是汇总/分析时，因无 headers 而无法校验列名
        if (results.length === 0 && turn && turn.context && turn.context.selection) {
            const sel = turn.context.selection;
            if (Array.isArray(sel.headers) && sel.headers.length > 0) {
                headers = sel.headers.map(String);
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 第一步从 turn.context.selection 获取列名: ${headers.join(', ')}`);
            } else if (sel.range && sel.sheetName) {
                // 如果只有 range 没有 headers，尝试从 range 推断列数（如 A1:C10 → 3列）
                const colMatch = sel.range.match(/[A-Z]+/g);
                if (colMatch && colMatch.length >= 2) {
                    // 无法知道具体列名，但至少知道有数据
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 第一步无 headers，但有 range='${sel.range}'，列名校验将依赖后端`);
                }
            }
        }
        // 聚合步骤标记：影响表头提取（优先前置聚合结果列）与是否提前 break
        const isAggStep = /aggregate|汇总|聚合|统计/.test(String(step.toolName || step.tool || ''));
        // 仅复制/格式化步骤需要 sourceSheet/destinationSheet/sourceRange 等源表元数据；
        // 写入/聚合步骤只需 sheet/range/values，填充这些冗余参数会污染工具调用参数。
        const isFormatStep = /format|格式化/.test(String(step.toolName || step.tool || ''));
        const isCopyStep = /copy|复制/.test(String(step.toolName || step.tool || ''));
        const needsSourceMeta = isFormatStep || isCopyStep;
        for (const prevResult of results) {
            if (!prevResult || !prevResult.ok || !prevResult.result) {
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] prepareStepParams: 跳过无效结果, ok=${prevResult?.ok}, hasResult=${!!prevResult?.result}`);
                continue;
            }

            // 【直接解析】MCP 响应格式: { ok, result: { content: [{ type: "text", text: JSON }] } }
            // 不走递归，直接硬编码穿透三层结构
            const invokeResult = prevResult.result;  // { ok, result: { content: [...] } }
            const mcpResult = invokeResult && invokeResult.result;  // { content: [...] }
            ChatLog.info?.('MULTI-STEP', `[StepExecutor] prepareStepParams: invokeResult keys=${Object.keys(invokeResult).join(',')}, has mcpResult=${!!mcpResult}, mcpResult type=${typeof mcpResult}`);
            if (mcpResult) {
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] prepareStepParams: mcpResult keys=${Object.keys(mcpResult).join(',')}, has content=${Array.isArray(mcpResult.content)}, content length=${mcpResult.content?.length}`);
                if (mcpResult.content && mcpResult.content.length > 0) {
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] prepareStepParams: content[0] type=${mcpResult.content[0]?.type}, text=${mcpResult.content[0]?.text?.slice(0, 100)}`);
                }
            }
            let toolData = null;
            // mcpResult 可能是字符串（MCP 直接返回文本），也可能是 { content: [...] } 对象
            if (mcpResult && typeof mcpResult === 'string') {
                // 情况 1: 字符串 → 直接尝试 JSON 解析，失败则作为纯文本
                const trimmed = mcpResult.trim();
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 原始字符串前 120 字符: ${trimmed.slice(0, 120)}`);
                try {
                    const parsed = JSON.parse(trimmed);
                    if (parsed && typeof parsed === 'object') {
                        // Case A: JSON 解析出 { content: [...] } 格式（MCP 标准响应嵌套）
                        if (Array.isArray(parsed.content) && parsed.content.length > 0) {
                            for (const item of parsed.content) {
                                if (item && typeof item.text === 'string' && item.text.trim()) {
                                    try {
                                        const innerParsed = JSON.parse(item.text);
                                        if (innerParsed && typeof innerParsed === 'object') {
                                            toolData = innerParsed;
                                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 字符串 JSON→content 解析成功: keys=${Object.keys(innerParsed).join(',')}`);
                                            break;
                                        }
                                    } catch (_) { /* inner JSON 解析失败，继续 */ }
                                }
                            }
                            // content 数组解析失败时，回退到 parsed 本身
                            if (!toolData && parsed.worksheet) {
                                toolData = parsed;
                            }
                        }
                        // Case B: JSON 直接解析出业务对象
                        if (!toolData) {
                            toolData = parsed;
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 字符串 JSON 解析成功: keys=${Object.keys(parsed).join(',')}`);
                        }
                    }
                } catch (_) {
                    // JSON 解析失败 → 尝试修复缺少逗号的 JSON
                    try {
                        const repaired = trimmed
                            .replace(/}\s*{/g, '},{')                       // 对象间补逗号
                            .replace(/"\s*\n\s*"/g, '",\n"')               // key-value 间补逗号
                            .replace(/(\d+)\s*\n\s*"/g, '$1,\n"');         // 值后补逗号
                        const parsed = JSON.parse(repaired);
                        if (parsed && typeof parsed === 'object') {
                            toolData = parsed;
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] JSON 修复后解析成功: keys=${Object.keys(parsed).join(',')}`);
                        }
                    } catch (_) { /* 修复后仍失败，走文本回退 */ }
                }
                // 文本回退提取：从原始文本中用正则提取所有字段（不依赖 JSON 格式）
                if (!toolData) {
                    // 提取 worksheet/name
                    let wsName = null;
                    const m1 = trimmed.match(/工作表[：:]\s*([^"\s，、。；;]+)/);
                    if (m1) wsName = m1[1];
                    if (!wsName) {
                        const m2 = trimmed.match(/"(?:worksheet|name)":\s*"([^"]+)"/);
                        if (m2) wsName = m2[1];
                    }
                    // 提取 previousSheet/source
                    let prevSheet = null;
                    const ps1 = trimmed.match(/"previousSheet":\s*"([^"]+)"/);
                    if (ps1) prevSheet = ps1[1];
                    // 提取 message
                    let msg = null;
                    const msg1 = trimmed.match(/"message":\s*"([^"]+)"/);
                    if (msg1) msg = msg1[1];
                    if (wsName || prevSheet) {
                        toolData = {};
                        if (wsName) toolData.worksheet = wsName;
                        if (prevSheet) toolData.previousSheet = prevSheet;
                        if (msg) toolData.message = msg;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 正则提取: worksheet='${wsName}', previousSheet='${prevSheet}'`);
                    }
                }
            } else if (mcpResult && Array.isArray(mcpResult.content) && mcpResult.content.length > 0) {
                // 情况 2: { content: [...] } 对象
                for (const item of mcpResult.content) {
                    if (item && typeof item.text === 'string' && item.text.trim()) {
                        try {
                            const parsed = JSON.parse(item.text);
                            if (parsed && typeof parsed === 'object') {
                                toolData = parsed;
                                break;
                            }
                        } catch (_) { /* continue */ }
                    }
                }
            }
            // 文本回退（针对 { content: [...] } 结构中的纯文本）
            if (!toolData && mcpResult && Array.isArray(mcpResult.content) && mcpResult.content.length > 0) {
                for (const item of mcpResult.content) {
                    if (item && typeof item.text === 'string') {
                        const text = item.text.trim();
                        const wsMatch = text.match(/工作表[：:]\s*(\S+)/);
                        if (wsMatch) {
                            toolData = { worksheet: wsMatch[1] };
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] content 文本回退提取 worksheet: '${wsMatch[1]}'`);
                            break;
                        }
                    }
                }
            }
            // 兜底：普通 invoke 格式 { ok, result: { worksheet, range, ... } }
            if (!toolData && invokeResult && invokeResult.result && typeof invokeResult.result === 'object' && !Array.isArray(invokeResult.result)) {
                toolData = invokeResult.result;
            }

            if (toolData && typeof toolData === 'object') {
                // 聚合步骤的后续聚合：数据源应指向前置聚合的结果表（resultSheet），而非原始源表。
                // 否则后续聚合会在源表上查找"数量/销售额_求和"等结果列 → 报"未找到数值列"。
                // destinationSheet 修正仅对格式化步骤生效（excel_format_range）：
                // 聚合/写入/复制等步骤的 sheet 由各自逻辑决定，不能被 destinationSheet（源表）覆盖。
                // 聚合步骤优先继承 resultSheet：即使 sheet 已被前置读取步骤的 worksheet 填充，
                // 只要当前 sheet 仍是源表（等于 worksheet），就用结果表覆盖，避免在源表上找结果列。
                if (isAggStep && toolData.resultSheet && (!params.sheet || params.sheet === toolData.worksheet) && !params.sheetName) {
                    params.sheet = toolData.resultSheet;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 聚合步骤继承结果表 sheet: '${toolData.resultSheet}'`);
                } else if (toolData.worksheet && !params.sheet && !params.sheetName) {
                    params.sheet = toolData.worksheet;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 继承 sheet: '${toolData.worksheet}'`);
                }
                if (needsSourceMeta && toolData.worksheet && !params.sourceSheet) {
                    // sourceSheet 仅在 previousSheet 可用时填充（VB 编译返回 JSON 时有此字段）
                    // 文本回退时 previousSheet 不可用，留空让 turn 上下文填充
                    if (toolData.previousSheet) {
                        params.sourceSheet = toolData.previousSheet;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 继承 sourceSheet: '${toolData.previousSheet}'`);
                    } else {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] sourceSheet 留空（无 previousSheet），等待 turn 上下文填充`);
                    }
                }
                if (needsSourceMeta && toolData.worksheet && !params.destinationSheet) {
                    params.destinationSheet = toolData.worksheet;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 继承 destinationSheet: '${toolData.worksheet}'`);
                }
                // 当 destinationSheet 已设置（即当前是格式化新表步骤）而 range 为空时，
                // 显式标记 autoRange=true 让 VB 端自动推断有数据的区域，避免格式化原始表格。
                // 不再用 range='A1' 魔法值（会被误解为"只格式化 A1"或"格式化整表"）。
                if (params.destinationSheet && !params.range && !params.startCell) {
                    params.autoRange = true;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 自动填充 autoRange=true (VB 自动推断有数据区域)`);
                }
                // 格式化步骤的 sheet 应使用新创建/复制的目标工作表，而非原始活动表
                // （仅格式化步骤：聚合/写入/复制等步骤的 sheet 由各自逻辑决定，不能被 destinationSheet 覆盖）
                if (isFormatStep && params.destinationSheet && params.sheet && params.sheet !== params.destinationSheet) {
                    const oldSheet = params.sheet;
                    params.sheet = params.destinationSheet;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 修正 sheet: '${oldSheet}' → '${params.destinationSheet}'`);
                }
                // 格式化步骤的 sheet 应优先使用 destinationSheet
                if (isFormatStep && params.destinationSheet && !params.sheet && !params.sheetName) {
                    params.sheet = params.destinationSheet;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从 destinationSheet 继承 sheet: '${params.destinationSheet}'`);
                }
                if (toolData.range && !params.range) {
                    params.range = toolData.range;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 继承 range: '${toolData.range}'`);
                }
                if (needsSourceMeta && toolData.range && !params.sourceRange) {
                    params.sourceRange = toolData.range;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 继承 sourceRange: '${toolData.range}'`);
                }
                // 聚合步骤：优先使用前置聚合的结果表列（data 数组的键），而非源表列名。
                // 后续聚合的数据源是前置聚合的结果表（resultSheet），其列是聚合结果列（如 部门/数量），
                // 用源表列名校验 valueColumn 会把"数量"误改成"销售额"。
                if (isAggStep && Array.isArray(toolData.data) && toolData.data.length > 0) {
                    const firstRow = toolData.data[0];
                    if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
                        const resultCols = Object.keys(firstRow);
                        if (resultCols.length > 0) {
                            headers = resultCols;
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 聚合步骤使用前置聚合结果列: ${resultCols.join(', ')}`);
                        }
                    }
                }
                if (headers.length === 0 && Array.isArray(toolData.headers)) {
                    headers = toolData.headers;
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从元数据取列名: ${headers.join(', ')}`);
                }
            }
            if (headers.length === 0) {
                headers = extractHeaders(prevResult.result);
            }
            // 聚合步骤不因源表列名提前 break：继续遍历以找到前置聚合结果（其结果列/resultSheet 优先）
            if (headers.length > 0 && !isAggStep) break;
        }

        // 写入步骤：从前置聚合结果提取数据，填充 values（JSON 二维数组）。
        // 独立于上方 headers 循环——读取结果含 headers 会提前 break，导致聚合 data 未被处理。
        // 逆序遍历取最近一个含 data 数组的聚合结果。
        const isWriteStep = /write|写入|填充|填入/.test(String(step.toolName || step.tool || ''));
        if (isWriteStep && !params.values) {
            for (let i = results.length - 1; i >= 0; i--) {
                const prevResult = results[i];
                if (!prevResult || !prevResult.ok || !prevResult.result) continue;
                const td = extractToolData(prevResult.result);
                if (td && Array.isArray(td.data) && td.data.length > 0) {
                    const values2d = buildValuesFromAggregate(td);
                    if (values2d) {
                        params.values = values2d;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从聚合结果填充 values（${td.data.length} 行数据）`);
                        // 【关键】根据 values 的实际尺寸计算精确的 range，避免 VB 后端错误扩展为超大范围
                        try {
                            const parsed = JSON.parse(values2d);
                            if (Array.isArray(parsed) && parsed.length > 0) {
                                const rowCount = parsed.length;
                                const colCount = Array.isArray(parsed[0]) ? parsed[0].length : 1;
                                const endCol = getExcelColumnName(colCount);
                                const endRow = rowCount;
                                const preciseRange = `A1:${endCol}${endRow}`;
                                params.range = preciseRange;
                                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 根据 values 尺寸计算精确 range='${preciseRange}'（${rowCount} 行 × ${colCount} 列）`);
                            }
                        } catch (e) {
                            ChatLog.warn?.('MULTI-STEP', `[StepExecutor] 计算精确 range 失败（继续用默认值）:`, e?.message);
                        }
                        break;
                    }
                }
            }
        }

        // sourceSheet 仍未填充时，从 turn.context.selection 获取请求时的原始活动表名
        // （仅复制/格式化步骤需要；写入/聚合步骤不需要源表元数据）
        if (needsSourceMeta && !params.sourceSheet && turn && turn.context && turn.context.selection) {
            const selSheet = turn.context.selection.sheetName;
            if (selSheet) {
                params.sourceSheet = selSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 从 turn 上下文继承 sourceSheet: '${selSheet}'`);
            }
        }

        if (headers.length === 0) return;

        // 分类列名：文本列 vs 数值列（关键词唯一事实源：Config/intent_config.json → intentConfig.defs().columnHints）
        // 注意：日期/时间列不作为分组维度候选（避免"按产品分类汇总"被错误分组到日期列）
        const ic = App.ChatV2.intentConfig;
        const ch = (ic && typeof ic.defs === 'function' && ic.defs().columnHints) || {};
        const textHints = ch.textHints || null;
        const categoryHints = ch.categoryHints || null;
        const nonCategoryHints = ch.nonCategoryHints || null;
        const numericHints = ch.numericHints || null;

        const textColumns = headers.filter(h => !textHints || textHints.test(h));
        const categoryColumns = headers.filter(h => (!categoryHints || categoryHints.test(h)) && (!nonCategoryHints || !nonCategoryHints.test(h)));
        const numericColumns = headers.filter(h => !numericHints || numericHints.test(h));
        const otherColumns = headers.filter(h => (!textHints || !textHints.test(h)) && (!numericHints || !numericHints.test(h)));

        // 列名映射规则：聚合步骤即使 LLM 未产出 dimColumn，也从实际数据挑选分组列
        // （避免"找出卖得好的产品"这类未指定分组列的请求因 dimColumn 缺失而被 completeRequiredArgs 拦截）
        const isAggregateStep = /aggregate|汇总|聚合|统计/.test(String(step.toolName || step.tool || ''));
        const isDateGroupStep = /date_group|日期分组|时间分组/.test(String(step.toolName || step.tool || ''));
        const hasDimField = 'dimColumn' in params || 'dimField' in params || 'groupBy' in params;
        if (isAggregateStep || isDateGroupStep || hasDimField) {
            const field = hasDimField ? ('dimColumn' in params ? 'dimColumn' : 'dimField' in params ? 'dimField' : 'groupBy') : 'dimColumn';
            const current = params[field];
            // 只在当前值为空或不在实际列名中时才填充
            const needsFill = !current || !headers.some(h => h.toLowerCase() === String(current).toLowerCase());
            if (needsFill) {
                // 日期分组步骤优先找日期列
                if (isDateGroupStep) {
                    const dateHints = /日期|时间|date|time|年|月|日|created|updated/i;
                    const dateCol = headers.find(h => dateHints.test(h));
                    if (dateCol) {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 日期分组预填充 dateColumn: '${current || ''}' → '${dateCol}'`);
                        params.dateColumn = dateCol;
                    } else {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 未找到日期列，dateColumn 留空交由后端兜底（可用列: ${headers.join(', ')}）`);
                    }
                } else {
                    // 分类列优先，其次普通文本列；无合适文本列时不填充，交由后端兜底（避免误选日期列）
                    // 【意图感知】从步骤描述中提取用户意图的分组列（如"分析哪些产品卖得好"→优先选"产品"）
                    let fill = '';
                    if (categoryColumns.length > 0) {
                        const stepText = `${step.name || ''} ${step.description || ''}`;
                        // 优先匹配步骤描述中明确提到的分类列
                        const matchedCol = categoryColumns.find(h => stepText.includes(h));
                        fill = matchedCol || categoryColumns[0];
                        if (matchedCol) {
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 意图感知: 步骤描述匹配到分组列 '${matchedCol}'`);
                        }
                    } else if (textColumns.length > 0) {
                        fill = textColumns[0];
                    }
                    if (fill) {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 预填充 ${field}: '${current || ''}' → '${fill}'`);
                        params[field] = fill;
                        // 【透明度】当 LLM 指定的列名被自动修正为完全不同的列时，更新步骤描述避免误导用户
                        if (current && current.toLowerCase() !== fill.toLowerCase()) {
                            const oldDesc = step.description || step.name || '';
                            // 无论描述中是否显式包含原列名，都强制更新描述以反映实际分组字段
                            // 避免"按产品汇总"被改成按"地区"汇总后，描述仍显示"按产品"
                            const newDesc = oldDesc
                                ? oldDesc.replace(new RegExp(current, 'g'), fill)
                                : `按${fill}汇总`;
                            step.description = newDesc;
                            if (step.name) {
                                step.name = step.name.includes(current)
                                    ? step.name.replace(new RegExp(current, 'g'), fill)
                                    : `按${fill}汇总`;
                            }
                            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤描述同步修正: '${oldDesc}' → '${step.description}'`);
                        }
                    } else {
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 无合适分类列，${field} 留空交由后端兜底（可用列: ${headers.join(', ')}）`);
                    }
                }
            } else {
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 保留 ${field}: '${current}'（与实际列名匹配）`);
            }
        }

        // 数值列校验：LLM 猜测的 valueColumn 若不在实际列名中，自动替换为真实数值列。
        // （避免"分析哪些产品卖得好"被 LLM 猜测为"数量"而实际表无此列 → 聚合全 0 或报"未找到数值列"）
        const hasValueField = 'valueColumn' in params || 'valueField' in params || 'metric' in params;
        const hasAggColumn = 'aggColumn' in params;
        if ((isAggregateStep && hasValueField) || (isDateGroupStep && hasAggColumn)) {
            const vField = isDateGroupStep ? 'aggColumn' :
                          'valueColumn' in params ? 'valueColumn' : 'valueField' in params ? 'valueField' : 'metric';
            const vCurrent = params[vField];
            // 仅在 LLM 猜测了列名但该列不存在时才修正；留空（count 行数）不强制填充，交由后端兜底
            const vNeedsFix = vCurrent && !headers.some(h => h.toLowerCase() === String(vCurrent).toLowerCase());
            if (vNeedsFix) {
                const fill = (numericColumns.length > 0 ? numericColumns[0] :
                             otherColumns.length > 0 ? otherColumns[0] : '');
                if (fill) {
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 修正 ${vField}: '${vCurrent}' → '${fill}'（LLM 猜测列不存在，用实际数值列）`);
                    params[vField] = fill;
                } else {
                    ChatLog.info?.('MULTI-STEP', `[StepExecutor] 无合适数值列，${vField} 保留 '${vCurrent}'（可用列: ${headers.join(', ')}）`);
                }
            } else {
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 保留 ${vField}: '${vCurrent}'（与实际列名匹配）`);
            }
        }
    }

    /**
     * 从之前步骤的结果数据中提取列名，自动修正当前步骤的参数
     * 解决 LLM 规划时猜测的列名与实际数据列名不匹配的问题
     */
    function tryFixParamsFromPreviousResult(step, error, results) {
        const errMsg = error && (typeof error === 'string' ? error : (error.message || ''));
        if (!errMsg) return false;

        // 只处理"列未找到"类错误
        if (!/未找到数值列|未找到.*列|找不到.*列|column.*not found/i.test(errMsg)) {
            return false;
        }

        // 从错误信息中提取被猜测的列名
        // 错误格式为 JSON 字符串: {'success':'false','result':'未找到数值列 '金额'，请检查列名',...}
        // 优先从 JSON 的 result 字段中提取，其次直接匹配 "未找到数值列 'xxx'"
        let guessedColumn = null;
        try {
            // 尝试将错误信息中的单引号替换为双引号后解析 JSON
            const jsonStr = errMsg.replace(/'/g, '"');
            const parsed = JSON.parse(jsonStr);
            const resultText = parsed.result || parsed.message || '';
            const colMatch = resultText.match(/未找到数值列\s+['"]([^'"]+)['"]/);
            if (colMatch) guessedColumn = colMatch[1];
        } catch (_) {
            // JSON 解析失败，直接从错误文本中提取
            const colMatch = errMsg.match(/未找到数值列\s+'([^']+)'/);
            if (colMatch) guessedColumn = colMatch[1];
        }
        if (!guessedColumn) return false;

        // 遍历之前步骤的结果，提取数据列名
        for (const prevResult of results) {
            if (!prevResult || !prevResult.ok || !prevResult.result) continue;
            const headers = extractHeaders(prevResult.result);
            if (headers.length === 0) continue;

            // 如果猜测的列名已在数据中，说明不是列名问题
            if (headers.includes(guessedColumn)) return false;

            // 找最佳匹配列：优先数值含义列，否则取最后一列
            const ic = App.ChatV2.intentConfig;
            const ch = (ic && typeof ic.defs === 'function' && ic.defs().columnHints) || {};
            const numericHints = ch.numericHints || null;
            let bestMatch = null;

            for (const h of headers) {
                if (!numericHints || numericHints.test(h)) {
                    bestMatch = h;
                    break;
                }
            }
            // 没找到数值含义列时，排除"部门/姓名/日期/月份/年份"等非数值列，取第一个非文本列
            if (!bestMatch && headers.length > 1) {
                const nonTextHints = ch.nonTextHints || null;
                for (const h of headers) {
                    if ((!nonTextHints || !nonTextHints.test(h)) && h !== guessedColumn) {
                        bestMatch = h;
                        break;
                    }
                }
            }
            // 仍然没找到，取最后一列
            if (!bestMatch && headers.length > 1) {
                bestMatch = headers[headers.length - 1];
            }

            if (bestMatch && bestMatch !== guessedColumn && step.params) {
                const isDateGroupStep = /date_group|日期分组|时间分组/.test(String(step.toolName || step.tool || ''));
                const oldVal = isDateGroupStep ? step.params.aggColumn : step.params.valueColumn;
                if (isDateGroupStep) {
                    step.params.aggColumn = bestMatch;
                } else {
                    step.params.valueColumn = bestMatch;
                }
                // 同时修正数据源：前置聚合的结果表（resultSheet）作为后续聚合的数据源，
                // 否则修正后的 valueColumn（如"数量"）在源表上依然不存在 → 重试仍失败
                const toolData = extractToolData(prevResult.result);
                if (toolData && toolData.resultSheet) {
                    const isAggStep = /aggregate|汇总|聚合|统计/.test(String(step.toolName || step.tool || ''));
                    if (isAggStep) {
                        step.params.sheet = toolData.resultSheet;
                        ChatLog.info?.('MULTI-STEP', `[StepExecutor] 修正聚合数据源 sheet → '${toolData.resultSheet}'`);
                    }
                }
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 自动修正列名: '${oldVal || guessedColumn}' → '${bestMatch}'（可用列: ${headers.join(', ')}）`);
                return true;
            }
        }

        return false;
    }

    /**
     * 解析工具返回的字符串结果（含修复无效 JSON）
     * - 有效 JSON → 直接解析
     * - 无效 JSON（缺逗号）→ 修复后解析
     * - 非 JSON 文本 → 返回 { text: 原文 }
     */
    function parseToolResult(str) {
        if (typeof str !== 'string') return str;
        const trimmed = str.trim();
        if (!trimmed) return str;

        try {
            const parsed = JSON.parse(trimmed);
            if (parsed && typeof parsed === 'object') return parsed;
        } catch (_) { /* 无效 JSON，尝试修复 */ }

        try {
            const repaired = trimmed
                .replace(/}\s*{/g, '},{')                       // 对象间补逗号
                .replace(/"\s*\n\s*"/g, '",\n"')               // key-value 间补逗号
                .replace(/(\d+)\s*\n\s*"/g, '$1,\n"');         // 值后补逗号
            const parsed = JSON.parse(repaired);
            if (parsed && typeof parsed === 'object') return parsed;
        } catch (_) { /* 修复后仍失败 */ }

        return { text: trimmed };
    }

    /**
     * 从 invokeTracked 结果中提取实际工具返回数据。
     * 实际结构是嵌套的:
     *   executeStep → { ok, result: invokeTrackedResult }
     *   invokeTrackedResult → { ok, result: invokeResult }
     *   invokeResult → { content: [{ type: "text", text: '{...}' }], isError: false }
     *               或 { worksheet, range, headers, ... }
     * 此函数递归穿透直到拿到实际数据对象。
     */
    function extractToolData(invokeResult, depth = 0) {
        if (!invokeResult || depth > 5) return null;

        // 字符串结果：尝试解析（含修复无效 JSON），如 MCP 返回的 "{\r\n \"worksheet\": \"汇总\"...}"
        if (typeof invokeResult === 'string') {
            const parsed = parseToolResult(invokeResult);
            if (parsed && typeof parsed === 'object') return parsed;
            return null;
        }
        if (typeof invokeResult !== 'object') return null;

        // 如果已经是实际数据对象（有 worksheet 或 range 等业务字段），直接返回
        if (invokeResult.worksheet || invokeResult.range || Array.isArray(invokeResult.headers)) {
            return invokeResult;
        }

        // 情况1: MCP 响应格式 — content 数组包含 JSON 文本
        if (Array.isArray(invokeResult.content) && invokeResult.content.length > 0) {
            for (const item of invokeResult.content) {
                if (item && typeof item.text === 'string' && item.text.trim()) {
                    const parsed = parseToolResult(item.text);
                    if (parsed && typeof parsed === 'object') return parsed;
                }
            }
        }

        // 情况2: 有 result 字段 → 递归深入（可能是对象或字符串）
        if (invokeResult.result && (typeof invokeResult.result === 'object' || typeof invokeResult.result === 'string')) {
            const deeper = extractToolData(invokeResult.result, depth + 1);
            if (deeper) return deeper;
        }

        // 情况3: 遍历所有属性，找第一个有业务字段的对象
        for (const key of Object.keys(invokeResult)) {
            if (key === 'ok' || key === 'error' || key === 'durationMs' || key === 'toolCallId' || key === 'isError') continue;
            const val = invokeResult[key];
            if (val && (typeof val === 'object' || typeof val === 'string')) {
                const deeper = extractToolData(val, depth + 1);
                if (deeper) return deeper;
            }
        }

        return null;
    }

    /**
     * 从工具调用结果中提取数据列名（表头）
     * 支持任意深层嵌套，使用 deepFindData 递归搜索
     */
    function extractHeaders(result) {
        if (!result) return [];
        // 工具结果可能被包裹在 result.result 中（invokeTracked 返回 { ok, result, ... }）
        let data = result.result || result;
        if (!data) return [];

        // 深度搜索数据数组
        const arr = deepFindData(data);
        if (arr && arr.length > 0) {
            // 数组元素是对象 → 取 key 作为列名
            if (typeof arr[0] === 'object' && !Array.isArray(arr[0])) {
                return Object.keys(arr[0]);
            }
            // 二维数组 [[标题行], [数据行]] → 取第一行作为列名
            if (Array.isArray(arr[0]) && arr.length >= 1) {
                return arr[0].filter(h => h != null && h !== '').map(String);
            }
        }
        return [];
    }

    /**
     * 深度搜索结果对象，找到第一个有意义的数据数组
     * 处理 3~4 层嵌套：{ success: true, result: { data: [...] } }
     */
    function deepFindData(obj, depth = 0) {
        if (!obj || typeof obj !== 'object' || depth > 5) return null;
        if (Array.isArray(obj)) return obj;

        const arrayKeys = ['data', 'rows', 'records', 'items', 'values', 'list'];
        for (const key of arrayKeys) {
            if (Array.isArray(obj[key]) && obj[key].length > 0) return obj[key];
        }

        if (obj.result && typeof obj.result === 'object') {
            const found = deepFindData(obj.result, depth + 1);
            if (found) return found;
        }

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

    /**
     * 将列索引转换为 Excel 列名（1→A, 2→B, 27→AA）
     * @param {number} index - 列索引（从1开始）
     * @returns {string} Excel 列名
     */
    function getExcelColumnName(index) {
        let result = '';
        let n = index;
        while (n > 0) {
            n--;
            result = String.fromCharCode(65 + (n % 26)) + result;
            n = Math.floor(n / 26);
        }
        return result;
    }

    /**
     * 将聚合结果的 data 数组（对象数组）转换为 JSON 二维数组，供 excel_write_range 写入。
     * [{部门:'家居',数量:5},...] → [["部门","数量"],["家居",5],...]
     * @param {Object} toolData - 聚合工具返回的数据对象（含 data 数组）
     * @returns {string|null} JSON 二维数组字符串；数据不可用时返回 null
     */
    function buildValuesFromAggregate(toolData) {
        if (!toolData || !Array.isArray(toolData.data) || toolData.data.length === 0) return null;
        const rows = toolData.data;
        const first = rows[0];
        if (!first || typeof first !== 'object' || Array.isArray(first)) return null;
        const headers = Object.keys(first);
        if (headers.length === 0) return null;
        const values = [headers];
        for (const row of rows) {
            if (!row || typeof row !== 'object') continue;
            values.push(headers.map(h => (row[h] !== undefined && row[h] !== null ? row[h] : '')));
        }
        return JSON.stringify(values);
    }

    /**
     * 解析区域地址包含的列数：'A1:I101' / '$A$1:$I$101' → 9；解析失败返回 0。
     * 用途：图表步骤只有源区域跨多列时才需要"剔除文本列"的收敛处理。
     */
    function rangeColumnCount(addr) {
        const m = String(addr || '').replace(/\$/g, '').match(/([A-Za-z]+)\d*\s*:\s*([A-Za-z]+)\d*/);
        if (!m) return 0;
        const idx = (s) => s.toUpperCase().split('').reduce((acc, ch) => acc * 26 + (ch.charCodeAt(0) - 64), 0);
        return Math.abs(idx(m[2]) - idx(m[1])) + 1;
    }

    /**
     * 单元格值是否为数值：数字类型，或可解析为数字的字符串（去掉千分位/货币/百分号）。
     * 注意 '2024-01-09' 这类日期串 Number() 为 NaN → 判为非数值，避免把日期列当数值系列。
     */
    function isNumericCellValue(v) {
        if (typeof v === 'number') return isFinite(v);
        if (typeof v === 'string') {
            const s = v.trim();
            return s !== '' && !isNaN(Number(s.replace(/[,¥$%\s]/g, '')));
        }
        return false;
    }

    /**
     * 从之前步骤的结果中查找可用的源数据行（对象数组）。
     * 从最近的结果往前找，优先取聚合结果（data 数组）或读取结果（result 数组）。
     * 供前端聚合使用：数据在前端可用时，聚合在面板中计算显示，不再生成 Excel 聚合表。
     */
    function findSourceRows(results) {
        if (!Array.isArray(results)) return null;
        for (let i = results.length - 1; i >= 0; i--) {
            const prevResult = results[i];
            if (!prevResult || !prevResult.ok || !prevResult.result) continue;
            const data = deepFindData(prevResult.result);
            if (Array.isArray(data) && data.length > 0 &&
                typeof data[0] === 'object' && !Array.isArray(data[0])) {
                return data;
            }
        }
        return null;
    }

    /**
     * 前端聚合：在面板中计算聚合结果，不生成 Excel 聚合表。
     * 按 dimColumn 分组，对 valueColumn 计算 aggFunc（sum/count/avg/max/min）。
     * 返回格式与后端 excel_aggregate 兼容（aggregated: true + data 数组），
     * 使 ProgressTracker 面板能直接渲染为表格、后续写入步骤能提取数据。
     */
    function clientAggregate(rows, params) {
        if (!Array.isArray(rows) || rows.length === 0) return null;
        const dimColumn = params.dimColumn || '';
        const valueColumn = params.valueColumn || '';
        const aggFunc = params.aggFunc || 'sum';
        if (!dimColumn) return null;

        const groups = new Map();
        for (const row of rows) {
            if (!row || typeof row !== 'object') continue;
            const key = row[dimColumn] != null ? String(row[dimColumn]) : '(空)';
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(row);
        }

        const data = [];
        for (const [key, groupRows] of groups) {
            const out = { [dimColumn]: key };
            if (valueColumn) {
                let val;
                switch (aggFunc) {
                    case 'count': val = groupRows.length; break;
                    case 'avg':
                    case 'average':
                        val = groupRows.length
                            ? groupRows.reduce((s, r) => s + (Number(r[valueColumn]) || 0), 0) / groupRows.length
                            : 0;
                        break;
                    case 'max': val = groupRows.reduce((m, r) => Math.max(m, Number(r[valueColumn]) || 0), -Infinity); break;
                    case 'min': val = groupRows.reduce((m, r) => Math.min(m, Number(r[valueColumn]) || 0), Infinity); break;
                    default: val = groupRows.reduce((s, r) => s + (Number(r[valueColumn]) || 0), 0);
                }
                out[valueColumn] = Math.round(val * 100) / 100;
            } else {
                out['数量'] = groupRows.length;
            }
            data.push(out);
        }

        return {
            success: true,
            aggregated: true,
            clientComputed: true,
            dimColumn,
            valueColumn,
            aggFunc,
            groupCount: data.length,
            headers: valueColumn ? [dimColumn, valueColumn] : [dimColumn, '数量'],
            data,
            message: `前端聚合完成：按「${dimColumn}」分组${valueColumn ? `，${aggFunc}(${valueColumn})` : '（计数）'}，共 ${data.length} 组`
        };
    }

    /**
     * 前端日期分组：在面板中按日期粒度分组，支持 year/quarter/month/week/day。
     * 返回格式与后端 excel_date_group 兼容，使 ProgressTracker 能渲染结果。
     */
    function clientDateGroup(rows, params) {
        if (!Array.isArray(rows) || rows.length === 0) return null;
        const dateColumn = params.dateColumn || '';
        const groupBy = params.groupBy || 'month';
        const aggColumn = params.aggColumn || '';
        if (!dateColumn) return null;

        const parseDate = (v) => {
            if (v == null) return null;
            if (v instanceof Date) return v;
            if (typeof v === 'number') {
                // Excel 序列号（1900 日期系统近似）
                return new Date((v - 25569) * 86400000);
            }
            const d = new Date(v);
            return isNaN(d.getTime()) ? null : d;
        };

        // 周起始日配置
        const weekStart = params.weekStart || 'monday';
        // 财年配置
        const fiscalYearStart = params.fiscalYearStart || 1;
        const useFiscalQuarter = params.fiscalQuarter === true;

        const getGroupKey = (d) => {
            if (!d) return '(无效日期)';
            let y = d.getFullYear();
            const m = d.getMonth() + 1;
            const day = d.getDate();

            switch (groupBy) {
                case 'year': return `${y}年`;
                case 'quarter': {
                    if (useFiscalQuarter && fiscalYearStart > 1) {
                        // 财年季度计算
                        const adjustedMonth = ((m - fiscalYearStart + 12) % 12) + 1;
                        const fq = Math.ceil(adjustedMonth / 3);
                        const fy = m >= fiscalYearStart ? y : y - 1;
                        return `${fy}财年Q${fq}`;
                    }
                    return `${y}年Q${Math.ceil(m / 3)}`;
                }
                case 'month': return `${y}年${m}月`;
                case 'week': {
                    // 支持周一/周日起始的 ISO-like 周计算
                    const isSundayStart = weekStart === 'sunday';
                    const jsDay = d.getDay(); // 0=周日, 1=周一, ..., 6=周六
                    // 计算该日期所在周的周一（或周日）
                    const offsetToWeekStart = isSundayStart
                        ? jsDay // 到本周日的偏移（往前推 jsDay 天）
                        : (jsDay === 0 ? 6 : jsDay - 1); // 到本周一的偏移
                    const weekStartDate = new Date(d);
                    weekStartDate.setDate(day - offsetToWeekStart);
                    const wsYear = weekStartDate.getFullYear();
                    const wsMonth = weekStartDate.getMonth() + 1;
                    const wsDay = weekStartDate.getDate();
                    // 计算该周起始日是一年中的第几周
                    const yearFirstDay = new Date(wsYear, 0, 1);
                    const yearFirstDayOfWeek = yearFirstDay.getDay();
                    const firstWeekStartOffset = isSundayStart
                        ? yearFirstDayOfWeek
                        : (yearFirstDayOfWeek === 0 ? 6 : yearFirstDayOfWeek - 1);
                    const firstWeekStart = new Date(wsYear, 0, 1 - firstWeekStartOffset);
                    const dayDiff = Math.floor((weekStartDate - firstWeekStart) / 86400000);
                    const weekNum = Math.floor(dayDiff / 7) + 1;
                    return `${wsYear}年第${weekNum}周`;
                }
                case 'day': return `${y}-${String(m).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
                default: return `${y}年${m}月`;
            }
        };

        const groups = new Map();
        for (const row of rows) {
            if (!row || typeof row !== 'object') continue;
            const rawDate = row[dateColumn];
            const d = parseDate(rawDate);
            const key = getGroupKey(d);
            if (!groups.has(key)) groups.set(key, []);
            groups.get(key).push(row);
        }

        const data = [];
        for (const [key, groupRows] of groups) {
            const out = { [groupBy === 'day' ? '日期' : '时间']: key };
            if (aggColumn) {
                const sum = groupRows.reduce((s, r) => s + (Number(r[aggColumn]) || 0), 0);
                out[aggColumn] = Math.round(sum * 100) / 100;
            } else {
                out['数量'] = groupRows.length;
            }
            data.push(out);
        }

        // 按时间顺序排序
        data.sort((a, b) => {
            const ta = a[groupBy === 'day' ? '日期' : '时间'];
            const tb = b[groupBy === 'day' ? '日期' : '时间'];
            return String(ta).localeCompare(String(tb));
        });

        const gMap = { year: '年', quarter: '季度', month: '月', week: '周', day: '日' };
        return {
            success: true,
            dateGrouped: true,
            clientComputed: true,
            dateColumn,
            groupBy,
            aggColumn,
            groupCount: data.length,
            headers: aggColumn ? ['时间', aggColumn] : ['时间', '数量'],
            data,
            message: `前端日期分组完成：按「${dateColumn}」${gMap[groupBy] || groupBy}分组${aggColumn ? `，汇总(${aggColumn})` : '（计数）'}，共 ${data.length} 组`
        };
    }

    /**
     * 检查步骤是否有可执行配置
     */
    function validateStep(step) {
        const strategy = step.strategy || 'direct';

        switch (strategy) {
            case 'direct':
                return !!step.tool;
            case 'agent':
                return !!step.prompt || !!step.description;
            case 'pipeline':
                return !!step.pipeline || !!step.name;
            default:
                return false;
        }
    }

    /**
     * 从步骤执行结果中提取上下文信息（如新创建的工作表名），供后续步骤使用
     * @param {Object} step - 步骤定义
     * @param {Object} stepResult - executeStep 的返回值 { ok, result, error, ... }
     * @param {Object} stepContext - 上下文存储对象（在 MultiStepEngine 中维护）
     */
    function extractStepContext(step, stepResult, stepContext) {
        const toolName = step.tool || step.toolName || step.name;
        if (!toolName) return;

        // 尝试从结果中提取工具数据
        const toolData = extractToolData(stepResult.result);
        if (!toolData) return;

        // excel_create_sheet: 新建了工作表
        if (toolName === 'excel_create_sheet') {
            let sheetName = toolData.worksheet || toolData.sheet || toolData.name;
            // 从文本消息中提取工作表名（如 "成功创建工作表: 分析结果"）
            if (!sheetName && toolData.text) {
                const textMatch = toolData.text.match(/工作表[：:]\s*([^\s,，]+)/);
                if (textMatch) sheetName = textMatch[1].trim();
            }
            if (sheetName) {
                stepContext.lastCreatedSheet = sheetName;
                if (toolData.previousSheet) {
                    stepContext.lastSourceSheet = toolData.previousSheet;
                }
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文: lastCreatedSheet='${sheetName}'`);
            }
        }

        // excel_copy_range: 复制数据到目标工作表
        if (toolName === 'excel_copy_range') {
            // 提取目标工作表名（从 params 或结果文本）
            if (step.params && step.params.destinationSheet) {
                stepContext.lastDestinationSheet = step.params.destinationSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文: lastDestinationSheet='${step.params.destinationSheet}'`);
            }
            if (step.params && step.params.sourceSheet) {
                stepContext.lastSourceSheet = step.params.sourceSheet;
            }
        }

        // excel_date_group: 日期分组结果表
        if (toolName === 'excel_date_group' && toolData.outputSheet) {
            stepContext.lastCreatedSheet = toolData.outputSheet;
            if (toolData.range) stepContext.lastSourceRange = toolData.range;
            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文: dateGroup outputSheet='${toolData.outputSheet}'`);
        }

        // excel_aggregate（前端聚合 + 图表数据源写入）: 记录结果表名供后续图表步骤引用
        if ((toolName === 'excel_aggregate' || /aggregate|汇总|聚合/.test(String(toolName))) && toolData.resultSheet) {
            stepContext.lastCreatedSheet = toolData.resultSheet;
            ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文: 聚合结果表='${toolData.resultSheet}'（供图表引用）`);
        }
    }

    /**
     * 从步骤上下文中填充当前步骤的缺失参数
     * @param {Object} step - 当前步骤定义
     * @param {Object} stepContext - 上下文存储对象
     */
    function fillStepContext(step, stepContext) {
        const toolName = step.tool || step.toolName || step.name;
        if (!toolName || !stepContext) return;
        if (!step.params) step.params = {};

        let changed = false;

        // excel_format_range: 格式化新创建的表
        if (toolName === 'excel_format_range') {
            // sheet 未设置时，优先使用最近创建的目标表
            if (!step.params.sheet && stepContext.lastDestinationSheet) {
                step.params.sheet = stepContext.lastDestinationSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 sheet='${stepContext.lastDestinationSheet}'（来自 lastDestinationSheet）`);
                changed = true;
            } else if (!step.params.sheet && stepContext.lastCreatedSheet) {
                step.params.sheet = stepContext.lastCreatedSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 sheet='${stepContext.lastCreatedSheet}'（来自 lastCreatedSheet）`);
                changed = true;
            }
            // 未指定 range 时，显式标记 autoRange=true 让 VB 端自动推断有数据区域。
            // 不再用 range='A1' 魔法值，避免被误解为"只格式化 A1"或"格式化整表"。
            if (!step.params.range) {
                step.params.autoRange = true;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 autoRange=true`);
                changed = true;
            }
        }

        // excel_copy_range: 复制数据到目标表
        if (toolName === 'excel_copy_range') {
            if (!step.params.sourceSheet && stepContext.lastSourceSheet) {
                step.params.sourceSheet = stepContext.lastSourceSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 sourceSheet='${stepContext.lastSourceSheet}'`);
                changed = true;
            }
            if (!step.params.destinationSheet && stepContext.lastCreatedSheet) {
                step.params.destinationSheet = stepContext.lastCreatedSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 destinationSheet='${stepContext.lastCreatedSheet}'`);
                changed = true;
            }
        }

        // excel_write_range: 写入新创建的表
        if (toolName === 'excel_write_range') {
            if (!step.params.sheet && stepContext.lastCreatedSheet) {
                step.params.sheet = stepContext.lastCreatedSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 sheet='${stepContext.lastCreatedSheet}'（写入目标表）`);
                changed = true;
            }
            if (!step.params.range) {
                step.params.range = 'A1';
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 步骤上下文填充 range='A1'`);
                changed = true;
            }
        }

        // excel_date_group: 日期分组步骤继承上下文
        if (toolName === 'excel_date_group') {
            if (!step.params.sheet && stepContext.lastDestinationSheet) {
                step.params.sheet = stepContext.lastDestinationSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 日期分组步骤继承 sheet='${stepContext.lastDestinationSheet}'`);
                changed = true;
            } else if (!step.params.sheet && stepContext.lastCreatedSheet) {
                step.params.sheet = stepContext.lastCreatedSheet;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 日期分组步骤继承 sheet='${stepContext.lastCreatedSheet}'`);
                changed = true;
            }
            if (!step.params.range && stepContext.lastSourceRange) {
                step.params.range = stepContext.lastSourceRange;
                ChatLog.info?.('MULTI-STEP', `[StepExecutor] 日期分组步骤继承 range='${stepContext.lastSourceRange}'`);
                changed = true;
            }
        }

        return changed;
    }

    /**
     * 【子步骤UI实时刷新】通知 ProgressTracker 更新子步骤状态
     * 当 agent/pipeline 策略的子步骤状态变化时，触发规划面板重渲染
     */
    function notifySubStepUpdate(step, ctx) {
        if (!step || !ctx) return;
        const { turn, sessionStore } = ctx;
        if (!turn || !sessionStore) return;

        // 通过 agentEvents 广播子步骤更新事件，让 MultiStepEngine/ProgressTracker 刷新 UI
        if (window.App.ChatV2.agentEvents && typeof App.ChatV2.agentEvents.emit === 'function') {
            App.ChatV2.agentEvents.emit('multi_step:substep_update', {
                stepId: step.id,
                subSteps: step.subSteps,
                turnId: turn.id
            });
        }

        // 直接触发 ProgressTracker 重渲染（如果可用）
        const progressTracker = App.ChatV2.ProgressTracker;
        if (progressTracker && typeof progressTracker.getSteps === 'function') {
            try {
                // 【步骤统一到 cbx v20261003】不再写 _planHtml，改把最新步骤状态 upsert 进 thoughtList
                const steps = progressTracker.getSteps() || [];
                if (sessionStore && typeof sessionStore.upsertStepThought === 'function') {
                    steps.forEach(function (s, i) {
                        sessionStore.upsertStepThought(turn.id, {
                            id: s.id || s.stepId || ('se_step_' + i),
                            title: s.description || s.name || ('步骤 ' + (i + 1)),
                            detail: s.detail || '',
                            status: s.status || 'pending',
                            toolName: s.toolName || null,
                            index: i
                        });
                    });
                }
            } catch (e) {
                ChatLog.warn?.('MULTI-STEP', '[StepExecutor] 子步骤UI刷新失败:', e?.message);
            }
        }
    }

    // 导出
    App.ChatV2.StepExecutor = {
        executeStep,
        validateStep,
        extractStepContext,
        fillStepContext,
        clientAggregate,
        clientDateGroup,
        findSourceRows
    };

})();