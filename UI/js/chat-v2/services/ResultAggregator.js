﻿﻿﻿﻿﻿﻿﻿/**
 * ============================================================
 * chat-v2/services/ResultAggregator.js — 结果汇总器
 * ============================================================
 * 职责：汇总多步骤管道中各步骤的执行结果，生成简洁的最终回答。
 *   - 收集所有步骤的执行结果和工具调用输出
 *   - 调用 LLM 生成自然语言摘要（可选）
 *   - 直接构建结构化摘要（快速路径）
 *
 * 设计：
 *   - 快速路径：轻量级结构化摘要，无需 LLM 调用
 *   - LLM 路径：调用 LLM 汇总生成自然语言回答（需上下文丰富）
 *   - 错误处理：部分步骤失败时也能生成有意义的摘要
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    /**
     * 汇总所有步骤的执行结果，生成最终摘要
     * @param {Object} opts
     * @param {Object[]} opts.steps - 所有步骤定义
     * @param {Object[]} opts.results - 对应的执行结果
     * @param {Object} [opts.llmIo] - 可选的 llmIo 实例（用于 LLM 汇总）
     * @param {Object[]} [opts.messages] - 原始对话消息（用于 LLM 汇总）
     * @param {Object} [opts.config] - 配置
     * @param {AbortSignal} [opts.signal]
     * @returns {Promise<{summary: string, type: string, stepResults: Object[]}>}
     */
    async function aggregate(opts) {
        const { steps, results, llmIo, messages, config, signal } = opts || {};

        if (!Array.isArray(steps) || !Array.isArray(results)) {
            return { summary: '无法汇总：缺少步骤或结果数据', type: 'error', stepResults: [] };
        }

        const stepResults = buildStepResults(steps, results);
        const allOk = stepResults.every(sr => sr.ok);
        const anyOk = stepResults.some(sr => sr.ok);
        const okCount = stepResults.filter(sr => sr.ok).length;
        const totalCount = stepResults.length;

        let summary;

        // 全部成功 -> 快速路径
        if (allOk && totalCount > 0) {
            summary = buildQuickSummary(stepResults, 'all_ok');
        }
        // 部分成功 -> 快速路径
        else if (anyOk && totalCount > 0) {
            summary = buildQuickSummary(stepResults, 'partial_ok');
        }
        // 全部失败 -> 快速路径
        else {
            summary = buildQuickSummary(stepResults, 'all_failed');
        }

        // 如果提供了 llmIo，尝试生成更自然的 LLM 摘要
        // 只在步骤数较少（<=5）且全部成功时使用 LLM 汇总
        if (llmIo && allOk && totalCount <= 5 && totalCount > 0) {
            try {
                const llmSummary = await buildLlmSummary(stepResults, llmIo, messages, config, signal);
                if (llmSummary) {
                    summary = llmSummary;
                }
            } catch (e) {
                ChatLog.warn?.('MULTI-STEP', '[ResultAggregator] LLM 汇总失败，使用快速路径:', e?.message);
                // 快速路径的 summary 已就绪，无需 fallback
            }
        }

        return { summary, type: allOk ? 'success' : (anyOk ? 'partial' : 'failure'), stepResults };
    }

    /**
     * 构建结构化步骤结果
     */
    function buildStepResults(steps, results) {
        return steps.map((step, i) => {
            const result = results[i] || {};
            const ok = result.ok === true;
            let toolResult = null;

            // 从结果中提取工具调用输出
            if (ok && result.result) {
                toolResult = result.result;
                // 如果结果是对象且有 text 或 content 字段，提取以便显示
                // 【保留结构化数据】agent 策略的结果包含 data 字段（内部工具调用的结构化数据），
                // 必须保留原始对象以便 extractDataContext 提取真实数据，避免 LLM 汇总时编造数字
                if (typeof toolResult === 'object' && !Array.isArray(toolResult)) {
                    const hasStructuredData = toolResult.data && (
                        toolResult.data.aggregated === true ||
                        toolResult.data.dateGrouped === true ||
                        Array.isArray(toolResult.data.data)
                    );
                    if (!hasStructuredData) {
                        const displayText = toolResult.text || toolResult.content || toolResult.summary || null;
                        if (displayText) {
                            toolResult = displayText;
                        }
                    }
                }
            }

            return {
                index: i,
                name: step.description || step.name || step.tool || `步骤 ${i + 1}`,
                tool: step.tool || '',
                strategy: step.strategy || 'direct',
                ok,
                error: result.error ? formatError(result.error) : null,
                toolResult: toolResult,
                iterations: result.iterations || 1
            };
        });
    }

    /**
     * 快速路径：构建结构化摘要（无需 LLM）
     */
    function buildQuickSummary(stepResults, mode) {
        const total = stepResults.length;
        const okList = stepResults.filter(sr => sr.ok);
        const failList = stepResults.filter(sr => !sr.ok);

        let lines = [];

        switch (mode) {
            case 'all_ok': {
                lines.push(`✅ 已完成所有 ${total} 个步骤的操作。`);
                if (total <= 5) {
                    for (const sr of stepResults) {
                        if (sr.toolResult) {
                            const resultStr = typeof sr.toolResult === 'string'
                                ? sr.toolResult.slice(0, 200)
                                : JSON.stringify(sr.toolResult).slice(0, 200);
                            lines.push(`  • ${sr.name}: ${resultStr}`);
                        } else {
                            lines.push(`  • ${sr.name}: 完成`);
                        }
                    }
                } else {
                    const names = stepResults.map(sr => sr.name).join('、');
                    lines.push(`  具体操作：${names}`);
                }
                break;
            }
            case 'partial_ok': {
                lines.push(`⚠️ 部分完成：${okList.length} 个步骤成功，${failList.length} 个失败。`);
                if (okList.length > 0) {
                    const okNames = okList.map(sr => sr.name).join('、');
                    lines.push(`  ✅ 成功：${okNames}`);
                }
                if (failList.length > 0) {
                    const failNames = failList.map(sr => `${sr.name}（${sr.error || '未知错误'}）`).join('、');
                    lines.push(`  ❌ 失败：${failNames}`);
                }
                break;
            }
            case 'all_failed': {
                lines.push(`❌ 所有 ${total} 个步骤均执行失败。`);
                for (const sr of stepResults) {
                    lines.push(`  • ${sr.name}: ${sr.error || '未知错误'}`);
                }
                break;
            }
        }

        return lines.join('\n');
    }

    /**
     * LLM 路径：调用 LLM 生成自然语言摘要
     * 增强：从步骤结果中提取用户数据上下文（如销售数据），让 LLM 生成更精准、更有洞察的总结
     */
    async function buildLlmSummary(stepResults, llmIo, messages, config, signal) {
        if (!llmIo || typeof llmIo.askText !== 'function') {
            return null;
        }

        // 【数据上下文提取】从步骤结果中提取实际业务数据，供 LLM 生成数据感知的总结
        const dataContext = extractDataContext(stepResults);

        // 构建步骤执行摘要作为 LLM 的上下文
        const stepsText = stepResults.map(sr => {
            const resultText = sr.toolResult
                ? (typeof sr.toolResult === 'string' ? sr.toolResult : JSON.stringify(sr.toolResult).slice(0, 500))
                : '执行成功';
            return `步骤 ${sr.index + 1}: ${sr.name}\n状态: 成功\n结果: ${resultText}`;
        }).join('\n\n');

        const systemPrompt = `你是一个 Excel 数据分析助手。你的任务是总结刚刚完成的${stepResults.length}个操作步骤的结果。

【绝对规则——必须严格遵守】
1. 你只能基于下方提供的【数据上下文】和【步骤结果】进行总结。严禁编造、推测或引入任何未在提供的数据中出现的信息（如具体产品名称、利润、占比、增长率等）。
2. 如果你无法从提供的数据中确认某个事实，必须忽略它，绝不能臆测填充。
3. 总结中的每一个数字都必须能在【数据上下文】中找到明确来源。如果数据上下文为空或不足，只描述已完成的操作，不要进行数据分析。
4. 不要为显得"有洞察"而强行解读。宁肯简短如实，也不要胡说八道。

【内容要求】
- 不要逐条复述步骤，给用户一个整体的、连贯的回答。
- 说明已经完成了什么操作。
- 如果数据上下文确实包含聚合/分析数据，可以提炼关键数字（如最大值、最小值、合计），但必须明确标注这些数据来自提供的数据。
- 如果合适，可以给出后续操作建议，但保持简短（1-2句话）。`;

        // 构建用户消息：包含数据上下文 + 步骤结果
        let userMsg = '';
        if (dataContext) {
            userMsg += `【数据上下文】\n${dataContext}\n\n`;
        } else {
            userMsg += `【数据上下文】\n（无可用数据上下文）\n\n`;
        }
        userMsg += `【步骤结果】\n${stepsText}\n\n请基于以上数据上下文和步骤结果进行总结。如果数据上下文为空或信息不足，请仅描述已完成的操作，不要进行数据分析或编造数字。`;

        const result = await llmIo.askText({
            messages: [
                { role: 'system', content: systemPrompt },
                ...(messages?.filter(m => m.role === 'user').slice(-1) || []),
                { role: 'user', content: userMsg }
            ],
            config,
            signal
        });

        if (result?.error) {
            throw new Error(result.error.message || 'LLM 汇总失败');
        }

        return result?.text || null;
    }

    /**
     * 从步骤结果中提取数据上下文（如聚合结果、读取的数据样本）
     * 用于增强 LLM 汇总的数据感知能力
     */
    function extractDataContext(stepResults) {
        const contexts = [];
        for (const sr of stepResults) {
            if (!sr.ok || !sr.toolResult) continue;
            let data = sr.toolResult;

            // 【Agent 策略数据提取】agent 步骤返回的 toolResult 可能包含 data 字段（内部工具调用的结构化数据）
            if (data.type === 'agent' && data.data) {
                data = data.data;
            }

            // 提取聚合结果数据（excel_aggregate）
            if (data.aggregated === true && Array.isArray(data.data) && data.data.length > 0) {
                const dimCol = data.dimColumn || '维度';
                const valCol = data.valueColumn || '数值';
                const aggFn = data.aggFunc || '汇总';
                const rows = data.data;
                // 计算合计、最大、最小
                let total = 0, maxVal = -Infinity, minVal = Infinity;
                let maxItem = null, minItem = null;
                for (const row of rows) {
                    const val = Number(row[valCol]) || 0;
                    total += val;
                    if (val > maxVal) { maxVal = val; maxItem = row; }
                    if (val < minVal) { minVal = val; minItem = row; }
                }
                const lines = [
                    `聚合结果（按「${dimCol}」${aggFn}「${valCol}」）：共 ${rows.length} 组`,
                    `合计: ${total.toLocaleString()}`,
                    `最高: ${maxItem?.[dimCol] || '-'} = ${maxVal.toLocaleString()}`,
                    `最低: ${minItem?.[dimCol] || '-'} = ${minVal.toLocaleString()}`
                ];
                // 添加前5行数据样本
                const sample = rows.slice(0, 5).map(r => `  ${r[dimCol]}: ${r[valCol]}`).join('\n');
                contexts.push(lines.join('，') + '\n' + sample);
            }
            // 提取日期分组结果数据（excel_date_group）
            else if (data.dateGrouped === true && Array.isArray(data.data) && data.data.length > 0) {
                const dateCol = data.dateColumn || '日期';
                const groupBy = data.groupBy || 'month';
                const gMap = { year: '年', quarter: '季度', month: '月', week: '周', day: '日' };
                const groupLabel = gMap[groupBy] || groupBy;
                const aggCol = data.aggColumn;
                const rows = data.data;
                const valCol = aggCol || '数量';
                // 计算合计、最大、最小
                let total = 0, maxVal = -Infinity, minVal = Infinity;
                let maxItem = null, minItem = null;
                for (const row of rows) {
                    const val = Number(row[valCol]) || 0;
                    total += val;
                    if (val > maxVal) { maxVal = val; maxItem = row; }
                    if (val < minVal) { minVal = val; minItem = row; }
                }
                const timeKey = groupBy === 'day' ? '日期' : '时间';
                const lines = [
                    `日期分组结果（按「${dateCol}」${groupLabel}分组）：共 ${rows.length} 组`,
                    `合计: ${total.toLocaleString()}`,
                    `最高: ${maxItem?.[timeKey] || '-'} = ${maxVal.toLocaleString()}`,
                    `最低: ${minItem?.[timeKey] || '-'} = ${minVal.toLocaleString()}`
                ];
                const sample = rows.slice(0, 5).map(r => `  ${r[timeKey]}: ${r[valCol]}`).join('\n');
                contexts.push(lines.join('，') + '\n' + sample);
            }
            // 提取读取的数据元数据
            else if (data.worksheet && data.rows) {
                contexts.push(`读取工作表「${data.worksheet}」：${data.rows} 行 × ${Array.isArray(data.headers) ? data.headers.length : '?'} 列`);
            }
            // 提取写入/复制结果
            else if (data.message && typeof data.message === 'string') {
                if (data.message.includes('成功') || data.message.includes('已创建') || data.message.includes('已复制')) {
                    contexts.push(data.message);
                }
            }
        }
        return contexts.length > 0 ? contexts.join('\n') : null;
    }

    /**
     * 格式化错误信息
     */
    function formatError(err) {
        if (!err) return '未知错误';
        if (typeof err === 'string') return err;
        if (err instanceof Error) return err.message;
        if (err.message) return err.message;
        try { return JSON.stringify(err); } catch (_) { return '未知错误'; }
    }

    // 导出
    App.ChatV2.ResultAggregator = {
        aggregate
    };

})();