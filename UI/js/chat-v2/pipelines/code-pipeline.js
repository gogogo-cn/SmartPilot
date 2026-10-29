/**
 * ============================================================
 * chat-v2/pipelines/code-pipeline.js — 纯代码生成展示流程
 * ============================================================
 * 职责：
 *   - 根据 Skill 的 systemPrompt 生成代码并展示给用户
 *   - 不执行任何代码，仅提供复制/保存功能
 *
 * 触发条件：
 *   - 用户明确要求"写代码/生成代码/写脚本"等纯代码生成意图
 *   - 匹配到 general-python / general-vba 等代码生成类 Skill
 *
 * 与 python-pipeline 的区别：
 *   - python-pipeline：生成 + 自动执行（分析类任务）
 *   - code-pipeline：仅生成展示（"帮我写代码"类请求）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    const ChatLog = window.ChatLog || console;

    const DEFAULT_PROMPT = `你是 SmartPilot 的代码生成助手。

【任务】
根据用户的描述，生成可直接使用的代码。

【严格规则】
1. 仅输出代码，不要解释、不要说明、不要 Markdown 代码块标记、不要任何额外文字。
2. 代码必须可直接使用。
3. 如果是 Python 代码，已预置 pandas(pd)、numpy(np)、matplotlib.pyplot(plt)。
4. 如果是 VBA 代码，必须是 Sub/End Sub 结构，可直接运行。
5. 引用工作表时必须使用名称（如 Sheets("Sheet1")），禁止使用索引（如 Sheets(1)）。
6. 禁止提及任何工具、函数调用、宏执行等内容。只输出纯代码。`;

    function buildContextBlock(ctx) {
        const parts = [];

        const selection = ctx?.selection;
        if (selection && Array.isArray(selection.headers) && selection.headers.length > 0) {
            parts.push(
                `【当前选区】${selection.address || selection.range || ''}\n` +
                `工作簿：${selection.workbookName || selection.workbook || ''}\n` +
                `工作表：${selection.sheetName || selection.worksheet || ''}\n` +
                `表头：${selection.headers.join(', ')}\n` +
                `样本行数：${Array.isArray(selection.rows) ? selection.rows.length : 0}`
            );
        }

        const refs = ctx?.referenceRegions;
        if (Array.isArray(refs) && refs.length > 0) {
            parts.push('【引用区域】');
            refs.forEach((r, i) => {
                const headers = Array.isArray(r.headers) ? r.headers.join(', ') : '';
                const rowCount = Array.isArray(r.rows) ? r.rows.length : 0;
                parts.push(
                    `区域${i + 1}: [${r.workbook || ''}]${r.worksheet || ''}!${r.address || ''}\n` +
                    `表头：${headers}\n样本行数：${rowCount}`
                );
            });
        }

        return parts.join('\n\n');
    }

    function buildMessages(turn, detectedLang, prevCode, ctxOverride) {
        const userMsg = turn.userMessage || '';
        // 后续消息继承：无新选区时用上一轮 _execSpec.data 快照补全上下文
        const contextBlock = buildContextBlock(ctxOverride || turn.context || {});
        let userContent = contextBlock
            ? `${userMsg}\n\n${contextBlock}`
            : userMsg;

        // 代码后续消息继承：把上一轮代码注入提示，让 LLM 基于它修改/追加而非从零生成
        if (prevCode) {
            userContent += `\n\n【上一轮生成的代码】\n${prevCode}\n\n请基于以上代码，按我的新需求修改或追加。只输出修改后的完整代码，不要解释、不要 Markdown 代码块标记。`;
        }

        const matchedSkill = turn.context?._matchedSkill;
        // 后端序列化字段为 llmSystemPrompt（VSTO GetAllSkills / Web 共享 Skill 均一致），
        // 兼容旧的自定义 Skill 可能使用的 systemPrompt 命名。
        let systemPrompt = (matchedSkill && (matchedSkill.llmSystemPrompt || matchedSkill.systemPrompt))
            ? (matchedSkill.llmSystemPrompt || matchedSkill.systemPrompt)
            : DEFAULT_PROMPT;

        // 用户意图（VBA）与 skill 标签（Python）冲突时，以用户意图为准
        // 覆盖 system prompt 中的语言引导，防止 LLM 输出 Python
        // json 模式不覆盖，保留 Skill 自定义的 systemPrompt
        // 例外：GenerateVBA 自带"输出 IntentSpec JSON → 模板引擎渲染"架构提示词，
        //       覆盖会让 LLM 直接输出未经模板验证的原始 VBA，必须保留。
        const isVbaSkill = !!(matchedSkill && (matchedSkill.name === 'GenerateVBA' || matchedSkill.Name === 'GenerateVBA'));
        if (detectedLang === 'vba' && !isVbaSkill) {
            systemPrompt = DEFAULT_PROMPT;
        }

        return [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userContent }
        ];
    }

    function stripToolText(code) {
        if (!code) return '';
        let result = code;
        const cutPatterns = [
            /\n?为了.{0,20}写入[\s\S]*$/i,
            /\n?请调用[\s\S]*$/i,
            /\n?调用以下工具[\s\S]*$/i,
            /\n?工具[：:][\s\S]*$/i
        ];
        for (const p of cutPatterns) {
            result = result.replace(p, '');
        }
        result = result.replace(/\[excel_[a-z_]+\][\s\S]*$/gi, '');
        return result.trim();
    }

    function extractCode(content) {
        if (!content) return '';
        const trimmed = content.trim();

        // 保留 markdown 代码块标记（含语言标签），让渲染器正确识别语言
        const codeBlockMatch = trimmed.match(/```(\w*)\s*\n?([\s\S]*?)```/i);
        if (codeBlockMatch) {
            const lang = (codeBlockMatch[1] || '').trim();
            const code = stripToolText(codeBlockMatch[2].trim());
            if (lang) {
                return '```' + lang + '\n' + code + '\n```';
            }
            return code;
        }

        return stripToolText(trimmed);
    }

    /**
     * 从用户消息中识别 PPT 主题（与 Config/ppt_templates.json themes 的 id 对齐）。
     * 未命中任何主题关键词时回退默认主题 business_blue。
     */
    function detectPptTheme(userMessage) {
        const msg = String(userMessage || '');
        const rules = [
            { id: 'business_blue', keywords: ['商务', '正式', '专业', '蓝色', '商务蓝', '沉稳', '汇报'] },
            { id: 'tech_green', keywords: ['科技', '绿色', '科技绿', '清新', '环保', '创新', '互联网'] },
            { id: 'vibrant_orange', keywords: ['活力', '橙色', '活力橙', '热情', '醒目', '活泼', '年轻'] },
            { id: 'elegant_purple', keywords: ['优雅', '紫色', '优雅紫', '高端', '大气', '奢华', '质感'] },
            { id: 'academic_navy', keywords: ['学术', '深蓝', '学术深蓝', '论文', '严谨', '学院'] },
            { id: 'academic_crimson', keywords: ['学术红', '朱红', '复古红', '典雅红'] },
            { id: 'light_fresh', keywords: ['浅色清新', '小清新', '清爽', '明亮', '马卡龙'] },
            { id: 'light_minimal', keywords: ['浅色极简', '极简', '简约', '留白', '性冷淡'] }
        ];
        for (const rule of rules) {
            if (rule.keywords.some(k => msg.includes(k))) return rule.id;
        }
        return 'business_blue';
    }

    /**
     * 【模板直出】GeneratePPTReport + 用户已给选区时，不依赖小模型，直接用选区数据构造 IntentSpec。
     * 满足条件才直出，否则返回 null 走原 LLM 路径。
     * @returns {object|null}
     */
    // ============================================================
    // 多轮消息继承：后续消息检测与上下文取回已泛化为通用机制
    // （orchestrator 的 FOLLOWUP_PATTERNS + detectFollowUp + session-store._execSpec）。
    // 本管道只需在生成时写入 _execSpec（pipeline='ppt'），
    // 后续消息由 orchestrator 检测并路由回本管道，此处从 getLastExecSpec() 继承。
    // ============================================================
    function buildDirectIntentSpec(ctx, userMessage) {
        const matchedSkill = ctx?._matchedSkill;
        const isPpt = matchedSkill && (matchedSkill.name === 'GeneratePPTReport' || matchedSkill.Name === 'GeneratePPTReport');
        if (!isPpt) return null;

        // 数据源优先取主选区；其次取第一个参考区域（附件切片多半会落到 referenceRegions）。
        const selection = ctx?.selection;
        const ref = Array.isArray(ctx?.referenceRegions) && ctx.referenceRegions.length > 0 ? ctx.referenceRegions[0] : null;
        const hasSel = selection && Array.isArray(selection.headers) && selection.headers.length > 0;
        const hasRef = ref && Array.isArray(ref.headers) && ref.headers.length > 0;
        if (!hasSel && !hasRef) return null;

        const src = hasSel ? selection : ref;
        const headers = src.headers;
        const sheet = src.sheetName || src.worksheet || src.worksheetName || 'Sheet1';
        const dataRange = (src.address && String(src.address).trim())
            || (src.range && String(src.range).trim()) || '';
        const rowCount = Math.max(1, Math.min(
            Number(src.rowCount) || (Array.isArray(src.rows) ? src.rows.length : 0) || 15, 20));

        // 标题优先取用户消息里带书名号/引号的部分，否则用"数据分析报告"
        let reportTitle = '数据分析报告';
        const quoted = (userMessage || '').match(/["“《「]([^"”》」]+)["”》」]/);
        if (quoted && quoted[1]) reportTitle = quoted[1];
        if (String(userMessage || '').includes('PPT') || /ppt/i.test(userMessage || '')) {
            reportTitle = `${sheet} 数据分析报告`;
        }

        const theme = detectPptTheme(userMessage);

        const slides = [
            { type: 'title', title: reportTitle, subtitle: `数据来源：${sheet}${dataRange ? ' · ' + dataRange : ''}` },
            {
                type: 'table',
                title: '数据明细',
                columns: headers,
                rowCount,
                startRow: 2
            },
            {
                type: 'text',
                title: '数据概览',
                bullets: [
                    `数据区域：${sheet}${dataRange ? ' ' + dataRange : ''}`,
                    `字段（${headers.length} 个）：${headers.slice(0, 8).join('、')}`,
                    '双击 Excel 生成后，可在 PowerPoint 中自由调整布局与样式'
                ]
            }
        ];

        return { scenario: 'ppt_from_data', reportTitle, dataSheet: sheet, dataRange, headers, rows: src.rows || [], slides, theme };
    }

    /**
     * 把上一轮执行上下文的数据快照归一化为 directSpec。
     * - ppt 管道：data 本身就是 directSpec（扁平 headers/rows/dataSheet）
     * - skill/text/python 管道：data 是 { selection, referenceRegions }，取首个有表头的区域
     */
    function buildDirectSpecFromExecData(data) {
        if (!data) return null;
        if (Array.isArray(data.headers) && data.headers.length > 0) {
            return data;
        }
        const sel = data.selection || null;
        const ref = Array.isArray(data.referenceRegions) && data.referenceRegions.length > 0
            ? data.referenceRegions[0] : null;
        const src = (sel && Array.isArray(sel.headers) && sel.headers.length > 0) ? sel : ref;
        if (!src || !Array.isArray(src.headers) || src.headers.length === 0) return null;
        const sheet = src.sheetName || src.worksheet || src.sheet || 'Sheet1';
        const dataRange = (src.address && String(src.address).trim())
            || (src.range && String(src.range).trim()) || '';
        return {
            scenario: 'ppt_from_data',
            reportTitle: '数据分析报告',
            dataSheet: sheet,
            dataRange,
            headers: src.headers,
            rows: Array.isArray(src.rows) ? src.rows : [],
            slides: null,
            theme: 'business_blue'
        };
    }

    // ============================================================
    // PPT 配置卡：主题 + 版式（生成前由用户在配置卡上自定义）
    // ============================================================

    /**
     * 根据用户选择的主题 + 版式勾选，生成后端 slides 规格。
     * 数据相关版式（chart/kpi/comparison/timeline）从选区数据推导默认内容，
     * 后端布局引擎对缺省字段有兜底。
     */
    function buildSlidesFromConfig(config, directSpec) {
        const layout = config.layout || {};
        const reportTitle = directSpec.reportTitle || '数据分析报告';
        const sheet = directSpec.dataSheet || '';
        const dataRange = directSpec.dataRange || '';
        const headers = Array.isArray(directSpec.headers) ? directSpec.headers : [];
        const rows = Array.isArray(directSpec.rows) ? directSpec.rows : [];
        const rowCount = rows.length;
        const sourceDesc = `数据来源：${sheet}${dataRange ? ' · ' + dataRange : ''}`;
        const slides = [];

        if (layout.title) {
            slides.push({ type: 'title', title: reportTitle, subtitle: sourceDesc });
        }
        if (layout.table) {
            slides.push({ type: 'table', title: '数据明细', columns: headers, rowCount, startRow: 2 });
        }
        if (layout.chart) {
            slides.push({ type: 'chart', title: '数据图表' });
        }
        if (layout.kpi) {
            slides.push({ type: 'kpi', title: '关键指标' });
        }
        if (layout.comparison) {
            const leftTitle = headers[0] || '方案 A';
            const rightTitle = headers[1] || '方案 B';
            const leftItems = rows.slice(0, 5).map(r => String(r[headers[0]] ?? ''));
            const rightItems = rows.slice(0, 5).map(r => String(r[headers[1]] ?? ''));
            slides.push({ type: 'comparison', title: '对比分析', leftTitle, rightTitle, leftItems, rightItems });
        }
        if (layout.timeline) {
            const dateCol = headers.find(h => /日期|时间|月份|年份|date|time/i.test(h)) || headers[0];
            const valCol = headers.find(h => h !== dateCol) || headers[1] || '';
            const items = rows.slice(0, 6).map(r => ({
                time: String(r[dateCol] ?? ''),
                text: valCol ? String(r[valCol] ?? '') : ''
            }));
            slides.push({ type: 'timeline', title: '时间线', items });
        }
        if (layout.text) {
            slides.push({
                type: 'text',
                title: '数据概览',
                bullets: [
                    `数据区域：${sheet}${dataRange ? ' ' + dataRange : ''}`,
                    `字段（${headers.length} 个）：${headers.slice(0, 8).join('、')}`,
                    '双击 Excel 生成后，可在 PowerPoint 中自由调整布局与样式'
                ]
            });
        }
        return slides;
    }

    /**
     * 用户驱动聚合视图落位（多视图：可同时展示多个分类聚合）：
     * 把配置卡"分析视图"里逐一添加的若干聚合视图（维度×数值×聚合）作为独立图表页应用，
     * 并"紧贴相关正文"——就地填充到推荐设计原有的图表锚点位置（趋势图在前、聚合图随后），
     * 锚点不足时剩余聚合图在"结尾"页前就近补位；未勾选（aggExcluded 命中 key）的视图页剔除。
     * 未添加任何视图（aggViews 空）时原样返回，交由默认模型决定。
     */
    function applySelectedAggViews(slides, aggViews, aggExcluded) {
        if (!Array.isArray(slides) || !Array.isArray(aggViews) || aggViews.length === 0) {
            return slides;
        }
        const exclSet = new Set(Array.isArray(aggExcluded) ? aggExcluded : []);
        const valid = aggViews.filter(v => v && v.xField && (Array.isArray(v.series) && v.series[0] || v.measure))
            .filter(v => {
                const measure = Array.isArray(v.series) ? String(v.series[0]) : String(v.measure);
                const k = [String(v.xField), measure, String(['sum', 'avg', 'count'].includes(v.agg) ? v.agg : 'sum')].join('|');
                return !exclSet.has(k);
            });
        if (valid.length === 0) return slides;
        const aggLabel = { sum: '求和', avg: '均值', count: '计数' };
        const aggOf = (v) => ['sum', 'avg', 'count'].includes(v.agg) ? v.agg : 'sum';
        const newCharts = valid.map(v => {
            const measure = Array.isArray(v.series) ? String(v.series[0]) : String(v.measure);
            return {
                type: 'chart',
                chartType: v.chartType || 'column',
                title: (v.title && String(v.title).trim()) ? String(v.title) : `${v.xField}·${measure}（${aggLabel[aggOf(v)] || aggOf(v)}）`,
                xField: String(v.xField),
                series: [measure],
                agg: aggOf(v),
                categories: (Array.isArray(v.categories) && v.categories.length) ? v.categories.map(String) : null,
                variant: 'clustered'
            };
        });
        const isTrend = (s) => {
            const t = String(s?.type || '').toLowerCase();
            if (t !== 'chart') return false;
            return /[年月日时分周]/.test(String(s?.xField || '')) && /line/i.test(String(s?.chartType || ''));
        };
        // 图表页序列：趋势图在前，各聚合视图随后，紧贴正文数据区
        const seq = slides.filter(isTrend).concat(newCharts);
        if (seq.length === 0) return slides;
        // 就地填充：把聚合图填进原"图表"锚点位置（图表紧贴相关正文的原有节奏）；
        // 锚点不足时，剩余聚合图在"结尾"页之前就近补位，避免堆到结尾页之后
        const out = [];
        let seqIdx = 0;
        for (let i = 0; i < slides.length; i++) {
            const t = String(slides[i]?.type || '').toLowerCase();
            if (t === 'chart') {
                if (seqIdx < seq.length) { out.push(seq[seqIdx++]); }
                // 锚点用尽：跳过原图表页，不再保留
            } else if (t === 'closing') {
                while (seqIdx < seq.length) { out.push(seq[seqIdx++]); }
                out.push(slides[i]);
            } else {
                out.push(slides[i]);
            }
        }
        while (seqIdx < seq.length) { out.push(seq[seqIdx++]); }
        return out;
    }

    /**
     * 请求后端数据感知设计推荐：分析选区数据 → 推荐版式/图表/洞察。
     * columnHints 为 LLM 列理解结果（可选），覆盖规则引擎的启发式选列。
     * requirement 为用户需求原文（可选），后端据此做结构偏置（趋势/对比/构成/明细）。
     * 失败时返回 null（前端回退 buildSlidesFromConfig 简单版式）。
     */
    async function requestPptDesign(directSpec, columnHints, requirement) {
        try {
            const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
            if (!bridge || typeof bridge.RecommendPptDesign !== 'function') return null;
            const argsJson = JSON.stringify({
                reportTitle: directSpec.reportTitle || '数据分析报告',
                sheetName: directSpec.dataSheet || '',
                address: directSpec.dataRange || '',
                headers: Array.isArray(directSpec.headers) ? directSpec.headers : [],
                rows: Array.isArray(directSpec.rows) ? directSpec.rows : [],
                columnHints: columnHints || null,
                requirement: requirement || ''
            });
            const json = await bridge.RecommendPptDesign(argsJson);
            const data = JSON.parse(json || '{}');
            if (data?.success !== true || !Array.isArray(data.slides)) return null;
            return data;
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[code-pipeline] 设计推荐失败，回退简单版式:', e);
            return null;
        }
    }

    /**
     * 【LLM 列理解】让大模型思考"该拿什么列"：识别主角指标/分类/时间/ID 列。
     * 这是对规则引擎启发式选列的语义增强——规则引擎只知"拿列"，不知"拿什么列"。
     * 失败返回 null（调用方回退规则引擎）。
     */
    async function analyzeColumnsViaLlm(userMessage, directSpec, signal) {
        try {
            const llmClient = App.ChatV2?.llmClient;
            if (!llmClient || typeof llmClient.complete !== 'function') return null;

            const headers = Array.isArray(directSpec?.headers) ? directSpec.headers : [];
            const rows = Array.isArray(directSpec?.rows) ? directSpec.rows : [];
            if (headers.length === 0) return null;
            const sample = rows.slice(0, 5).map(r => {
                const o = {};
                for (const h of headers.slice(0, 10)) o[h] = (r && r[h] !== undefined) ? String(r[h]) : '';
                return o;
            });

            const prompt = `你是数据分析师。根据数据字段和样本，判断哪些列适合作为分析主角。

数据字段：${headers.join('、')}
样本数据：${JSON.stringify(sample)}
用户需求：${userMessage}

请输出 JSON 对象（只输出 JSON，不要任何其他文字、解释或代码块标记）：
{
  "mainMetric": "主数值指标列名（趋势/构成/KPI 的主角，如 销售额）",
  "secondaryMetrics": ["次要数值指标列名"],
  "mainCategory": "主分类维度列名（构成分析用，如 地区）",
  "timeColumn": "时间/日期列名（没有则空字符串）",
  "idColumns": ["标识列名（订单ID/客户ID/编号/序号等，不参与指标分析）"],
  "ignoreColumns": ["与主题无关、不应展示的列名"]
}
规则：
- 标识列（ID/编号/序号/单号/流水号）必须放入 idColumns，绝不能当指标
- mainMetric 必须是真正有意义的数值指标（销售额/利润/数量等），不能是 ID
- 没有合适列时对应字段用空字符串或空数组`;

            const events = await llmClient.complete({
                messages: [
                    { role: 'system', content: '你是数据分析师，只输出 JSON 对象，不输出任何其他内容。' },
                    { role: 'user', content: prompt }
                ],
                config: App.config,
                signal
            });

            let text = '';
            for (const evt of events) {
                if (evt.type === 'error') return null;
                if (evt.type === 'token') text += evt.delta || '';
            }

            const obj = parseColumnHintsJson(text);
            if (!obj || typeof obj !== 'object') return null;

            // 校验：字段必须真实存在于 headers 中，防止 LLM 编造列名
            const hset = new Set(headers.map(h => String(h).toLowerCase()));
            const pick = (v) => (typeof v === 'string' && hset.has(v.toLowerCase())) ? v : '';
            const pickArr = (v) => Array.isArray(v)
                ? v.filter(x => typeof x === 'string' && hset.has(x.toLowerCase()))
                : [];
            const valid = {
                mainMetric: pick(obj.mainMetric),
                mainCategory: pick(obj.mainCategory),
                timeColumn: pick(obj.timeColumn),
                secondaryMetrics: pickArr(obj.secondaryMetrics),
                idColumns: pickArr(obj.idColumns),
                ignoreColumns: pickArr(obj.ignoreColumns)
            };
            if (!valid.mainMetric && !valid.mainCategory && !valid.timeColumn && valid.idColumns.length === 0) {
                return null;
            }
            return valid;
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[code-pipeline] LLM 列理解失败，回退规则引擎:', e);
            return null;
        }
    }

    /** 从 LLM 输出中提取并解析 JSON 对象（容忍 ```json 围栏与前后杂文本） */
    function parseColumnHintsJson(text) {
        if (!text) return null;
        const fenced = text.match(/```(?:json)?\s*(\{[\s\S]*?\})\s*```/);
        const candidate = fenced ? fenced[1] : text;
        const start = candidate.indexOf('{');
        const end = candidate.lastIndexOf('}');
        if (start < 0 || end <= start) return null;
        try {
            return JSON.parse(candidate.slice(start, end + 1));
        } catch (e) {
            return null;
        }
    }

    // 后端布局引擎支持的幻灯片类型白名单（与 PptGenerationService.SUPPORTED_SLIDE_TYPES 对齐）
    const SUPPORTED_SLIDE_TYPES = new Set([
        'title', 'cover', 'table', 'text', 'bullets', 'chart',
        'kpi', 'comparison', 'timeline', 'agenda', 'section', 'closing', 'end'
    ]);

    /** 按分析类型给出完整「叙事结构」故事线模板，让 PPT 讲清"指标→依据→结论" */
    const STORYBOARDS = {
        desc_stats: {
            label: '描述性统计',
            timeline: '- 封面(title)：报告标题取分析技能名或用户描述\n' +
                '- 目录(agenda)：指标概览 / 数值分布 / 分组对比 / 结论\n' +
                '- KPI：均值、中位数、标准差、最大值/最小值（直接取分析结果数值，每项单值）\n' +
                '- 表格(table)：各字段 count / mean / median / stdev / min / max / q1 / q3 明细\n' +
                '- 图表(chart,column)或comparison：分组/维度均值对比（若分析含分组）\n' +
                '- 要点(text)：结论性解读（直接采用分析给出的论断）\n' +
                '- 结尾(closing)'
        },
        trend: {
            label: '趋势/时序分析',
            timeline: '- 封面(title)\n' +
                '- 目录(agenda)\n' +
                '- KPI：最新值、累计值、平均增速、预测值（多个预测值拆成独立 item）\n' +
                '- 图表(chart,line)：时序趋势\n' +
                '- 图表(chart,column/pie)：分项或构成（若分析给出）\n' +
                '- 图表/要点：预测外推与展望\n' +
                '- 要点(text)：结论与建议\n' +
                '- 结尾(closing)'
        },
        comparison_share: {
            label: '对比/占比/构成分析',
            timeline: '- 封面(title)\n' +
                '- 目录(agenda)\n' +
                '- KPI：总量、占比最高项及份额\n' +
                '- 图表(chart,pie/doughnut)：结构占比\n' +
                '- 图表(chart,column/bar)：分项对比\n' +
                '- 表格(table)：分组明细\n' +
                '- 要点(text)：结论\n' +
                '- 结尾(closing)'
        },
        anomaly_quality: {
            label: '异常/数据质量问题分析',
            timeline: '- 封面(title)\n' +
                '- 目录(agenda)：异常概览 / 问题清单 / 归因建议\n' +
                '- KPI：异常记录数、缺失值数、重复记录数\n' +
                '- 表格(table)：异常/缺失/重复明细（字段、位置、原始值、标记）\n' +
                '- 要点(text)：成因与影响\n' +
                '- 要点/表格(text)：处理建议\n' +
                '- 结尾(closing)'
        },
        diff: {
            label: '数据差异/新旧对比分析',
            timeline: '- 封面(title)\n' +
                '- 目录(agenda)：差异概览 / 对比 / 明细 / 结论\n' +
                '- KPI：差异记录数、涉及字段数\n' +
                '- comparison：新旧/两侧对比\n' +
                '- 表格(table)：差异明细\n' +
                '- 要点(text)：影响与结论\n' +
                '- 结尾(closing)'
        },
        generic: {
            label: '通用数据分析',
            timeline: '- 封面(title)→目录(agenda)\n' +
                '- KPI：核心指标\n' +
                '- 图表(chart)：核心图表\n' +
                '- 表格(table)：数据明细\n' +
                '- 要点(text)：结论/建议\n' +
                '- 结尾(closing)'
        }
    };

    /**
     * 复合分析任务 → 有序去重的分析类型集合。
     * 与 deriveAnalysisKind 的单值不同，这里覆盖全部任务：多类任务并存时返回完整集合，
     * generic 仅在无任何可识别类型时垫底，不参与组合。
     */
    function deriveAnalysisKinds(prevSpec) {
        if (!prevSpec) return ['generic'];
        const names = [];
        const ts = prevSpec.taskStructure;
        if (ts && Array.isArray(ts.tasks)) {
            for (const t of ts.tasks) {
                if (t && (t.displayName || t.templateId)) names.push(t.displayName || t.templateId);
            }
        }
        const p = prevSpec.params || {};
        if (p.displayName) names.push(p.displayName);
        if (p.skillName) names.push(p.skillName);
        const kindOf = (n) => {
            if (/差异|diff/i.test(n)) return 'diff';
            if (/清洗|异常|缺失|重复|离群|质量|智能数据分析/i.test(n)) return 'anomaly_quality';
            if (/描述性统计|统计|均值|中位数|标准差|分位数/i.test(n)) return 'desc_stats';
            if (/占比|构成|份额|结构/i.test(n)) return 'comparison_share';
            if (/趋势|预测|时序|走势/i.test(n)) return 'trend';
            return null;
        };
        const kinds = [];
        for (const n of names) {
            const k = kindOf(n);
            if (k && !kinds.includes(k)) kinds.push(k);
        }
        if (kinds.length === 0) kinds.push('generic');
        return kinds;
    }

    /** 从上一轮分析上下文中启发式抽取数据特征（趋势方向/增幅/分组迹象），用于动态补强故事线 */
    function detectDataFeatures(analysisContext) {
        const text = String(analysisContext || '');
        const f = { trend: null, growthPct: null, group: false, hasKpis: false };
        if (/(下降|下滑|decrease|down)/i.test(text)) f.trend = 'down';
        else if (/(上升|增长|上涨|increase|up)/i.test(text)) f.trend = 'up';
        const g = text.match(/(?:较期初|增长|上涨|上升)\s*(?:约)?\s*([+-]?\d+(?:\.\d+)?)%/i);
        if (g) f.growthPct = parseFloat(g[1]);
        f.group = /(占比|构成|份额|分组|最高|最多)/.test(text);
        f.hasKpis = /(均值|中位数|标准差|合计|平均值|mean|median|stdev|count)/i.test(text);
        return f;
    }

    /**
     * 把分析类型集合 + 数据特征组装成完整故事线：
     * 多类任务按序串联；特征动态注入取值提示；受众决定叙事重心；数据维度补拼趋势/构成/对比页。
     */
    function composeStoryline(kinds, features, audience, dims) {
        const boards = [];
        for (const k of kinds) {
            if (STORYBOARDS[k]) boards.push(STORYBOARDS[k]);
        }
        if (boards.length === 0) boards.push(STORYBOARDS.generic);
        const label = kinds.length > 1
            ? '复合分析（' + kinds.map(k => (STORYBOARDS[k] || {}).label || k).join(' + ') + '）'
            : (STORYBOARDS[kinds[0]] || STORYBOARDS.generic).label;
        const lines = [];
        for (const b of boards) lines.push(b.timeline);
        const hints = [];
        if (features.trend) {
            const verb = features.trend === 'up' ? '上升' : (features.trend === 'down' ? '下降' : '平稳');
            const pct = features.growthPct != null ? '（增幅 ' + features.growthPct + '%）' : '';
            hints.push('数据整体呈' + verb + '趋势' + pct + '，趋势/预测页请用具体数值体现这一走势');
        }
        if (features.group) hints.push('数据含分组/占比特征，构成或对比页应呈现分组数与占比');
        if (hints.length) lines.push('【本次数据特征提示】\n- ' + hints.join('\n- '));

        // 维度驱动：在分析类型模板之外，按数据维度补拼叙事主线
        const dimLines = [];
        const d = dims || {};
        if (d.hasTime) dimLines.push('数据含时间维度：应有一条"随时间变化"的走势主线（折线图），讲清变化与关键节点');
        if (d.hasGroup && d.numericCount >= 2) dimLines.push('数据含分组维度且 ≥2 个数值列：应有一页构成/占比（饼/柱）与一页分组对比，讲清结构与差异');
        else if (d.numericCount >= 2) dimLines.push('数据含 ≥2 个数值列：首屏 KPI 提炼 2-4 个核心指标，并配一页指标对比(comparison)');
        if (dimLines.length) lines.push('【维度驱动补拼】\n- ' + dimLines.join('\n- '));

        const aud = (audience && AUDIENCE_STORYLINES[audience]) || null;
        // 结构化页级大纲种子（供 LLM 按骨架填充数值，而非自由发挥）
        const pages = buildSeedOutline(kinds, d, audience);
        return {
            label: label,
            body: lines.join('\n'),
            featureHints: hints,
            dimLines: dimLines,
            pages: pages,
            audienceLabel: aud ? aud.label : null,
            audienceLead: aud ? aud.lead : null,
            audienceStructure: aud ? aud.structure : null
        };
    }

    /** 生成对 LLM 的结构约束块（type 白名单来自渲染端支持集合，动态保持一致） */
    function buildStructureConstraints() {
        const types = Array.from(SUPPORTED_SLIDE_TYPES).join(' / ');
        return '【结构约束（必须严格遵守）】\n' +
            '- 每页 type 必须是以下白名单之一：' + types + '\n' +
            '- 页面总数 5-8 页；section 章节页最多 2-3 个；不要为每个数据列单独建"章节+表格"页（死模板）\n' +
            '- KPI 每项 value 只能是单个数值（禁止塞入多个值导致溢出）；含预测/未来值须拆成独立 item，label 指明；数值直接取上一轮分析结果\n' +
            '- 分析结果未提供的指标用"—"或省略，禁止凭空编造数值\n' +
            '- 保持"指标 → 依据 → 结论"的完整叙事';
    }

    /**
     * 受众/场景分层叙事：同一分析在不同受众面前叙事重心不同。
     * 汇报领导 → 结论先行（倒金字塔）；业务运营 → 问题驱动；数据团队 → 方法驱动。
     */
    const AUDIENCE_STORYLINES = {
        leadership: {
            label: '汇报/领导视角',
            lead: '结论先行（倒金字塔）：第一屏就用执行摘要给出核心结论与关键数字，再展开支撑依据，最后给建议。',
            structure: '- 首屏 KPI/要点给执行摘要与核心结论\n- 每章节"结论在前、依据在后"\n- 少堆明细，多讲影响与建议'
        },
        business: {
            label: '业务运营视角',
            lead: '问题驱动：发现问题 → 归因分析 → 量化影响 → 给出可执行建议。',
            structure: '- KPI+图表先定位问题与影响\n- comparison/表格做归因与对比\n- text 落在"怎么办"的可执行建议'
        },
        technical: {
            label: '数据/技术视角',
            lead: '方法驱动：先说清口径与分析方法 → 明细与验证 → 结论与局限。',
            structure: '- section 说明口径/方法与样本\n- table 呈现明细与验证\n- text 给结论并说明局限'
        }
    };

    /** 从用户诉求启发式判定受众/叙事视角（无法判断返回 null → 用通用叙事） */
    function deriveAudience(userMessage) {
        const t = String(userMessage || '');
        if (/(汇报|领导|老板|管理层|董事会|总裁|executive|leadership|汇报ppt|summary|结论先行)/i.test(t)) return 'leadership';
        if (/(业务|运营|销售|营销|客户|决策|怎么做|怎么办|如何|怎么提升|建议|方案|action|business)/i.test(t)) return 'business';
        if (/(技术|开发|算法|数据团队|口径|方法|逻辑|验证|analy|detail|model)/i.test(t)) return 'technical';
        return null;
    }

    /**
     * 维度驱动的数据特征：从列结构（含 LLM 列理解）判定有无时间/分组维度、数值列数量。
     * 用于在该分析类型模板之外，按数据本身的维度动态补拼故事线页（趋势/构成/对比）。
     */
    function deriveDimensionFeatures(directSpec, columnHints) {
        const headers = Array.isArray(directSpec?.headers) ? directSpec.headers : [];
        const rows = Array.isArray(directSpec?.rows) ? directSpec.rows : [];
        const dims = { hasTime: false, hasGroup: false, numericCount: 0, rowCount: rows.length };
        const ch = columnHints || {};
        if (typeof ch.timeColumn === 'string' && ch.timeColumn) dims.hasTime = true;
        if (typeof ch.mainCategory === 'string' && ch.mainCategory) dims.hasGroup = true;
        if (typeof ch.mainMetric === 'string' && ch.mainMetric) dims.numericCount += 1;
        if (Array.isArray(ch.secondaryMetrics)) dims.numericCount += ch.secondaryMetrics.length;
        if (headers.length) {
            const timeRe = /(日期|时间|年月|月份|年度|日$|年$|month|date|time|year|quarter|week)/i;
            const groupRe = /(地区|城市|省份|类别|分类|分组|部门|产品|类型|客户|区域|region|city|province|category|type|dept|group|product)/i;
            const numRe = /^(金额|数值|值|销售额|利润|数量|均价|total|amount|sales|profit|price|value|qty|count|score|score|数值)/i;
            for (const h of headers) {
                const s = String(h);
                if (!dims.hasTime && timeRe.test(s)) dims.hasTime = true;
                if (!dims.hasGroup && groupRe.test(s)) dims.hasGroup = true;
                if (numRe.test(s)) dims.numericCount += 1;
            }
        }
        return dims;
    }

    /**
     * 故事线 → 结构化的页级大纲种子（Page Outline）。
     * 把分析类型 + 数据维度 + 受众翻译成"type + 该页应放什么数据"的骨架，
     * LLM 只需按骨架填充具体数值（直接引分析结果），避免重算/编造。
     */
    function buildSeedOutline(kinds, dims, audience) {
        const pages = [];
        const dims2 = dims || {};
        const kinds2 = Array.isArray(kinds) ? kinds : ['generic'];
        if (audience === 'leadership') {
            pages.push({ type: 'kpi', title: '执行摘要', hint: '结论先行：提炼 2-4 个关键指标（单值）直接给结论' });
            pages.push({ type: 'text', title: '核心结论与建议', hint: '结论在前，依据在后' });
        } else {
            pages.push({ type: 'kpi', title: '关键指标总览', hint: '2-4 个核心指标，直接取分析结果单值' });
        }
        if (kinds2.includes('trend') || dims2.hasTime) {
            pages.push({ type: 'chart', chartType: 'line', title: '趋势/走势', hint: '时序变化，取分析结果趋势/预测数值' });
        }
        if (kinds2.includes('comparison_share') || (dims2.hasGroup && dims2.numericCount >= 2)) {
            pages.push({ type: 'chart', chartType: 'pie', title: '构成/占比', hint: '分项占比，取分析结果的份额/结构数值' });
        }
        if (dims2.numericCount >= 2 && !kinds2.includes('comparison_share')) {
            pages.push({ type: 'comparison', title: '指标对比', hint: '各组/各维度数值成对对比' });
        }
        if (kinds2.includes('anomaly_quality')) {
            pages.push({ type: 'table', title: '异常/问题清单', hint: '异常/缺失/重复明细（字段、位置、原始值、标记）' });
        }
        if (kinds2.includes('desc_stats')) {
            pages.push({ type: 'table', title: '统计明细', hint: 'count / mean / median / stdev / min / max 明细' });
        }
        if (kinds2.includes('diff')) {
            pages.push({ type: 'comparison', title: '差异对比', hint: '新旧/两侧成对差异数字' });
        }
        pages.push({ type: 'text', title: '结论与建议', hint: '直接采用分析结果的结论性语句' });
        return pages;
    }

    /** 从分析结果文本抽取可用作数值依据的关键词集合，供页内数据映射校验对齐 */
    function deriveDataMapKeys(analysisContext) {
        const text = String(analysisContext || '');
        const keys = new Set();
        const kw = ['执行摘要', '均值', '中位数', '标准差', '最大值', '最小值', '合计', '平均值',
            '占比', '份额', '增幅', '增长率', '趋势', '异常', '缺失', '重复', '差异', '总量', '预测'];
        for (const k of kw) if (text.includes(k)) keys.add(k);
        return keys;
    }

    /**
     * 页内数据映射校验（反编造/反溢出）：
     * - KPI 每项 value 必须是单值（多值 → 保留首个并提示）
     * - 置空/占位（"XX"/"[x]"/[]等）→ 回退 "—" 并提示，禁止透传给渲染端
     * - 数据域外字段名（明显是编造的占位列名）→ 回退空并提示
     * 返回 { slides, warnings }。
     */
    function validateSlideDataMapping(slides, analysisContext, headers) {
        const warnings = [];
        if (!Array.isArray(slides)) return { slides, warnings };
        const keys = deriveDataMapKeys(analysisContext);
        const headerSet = new Set((Array.isArray(headers) ? headers : []).map(h => String(h).toLowerCase()));
        const isPlaceholder = (v) => {
            const s = String(v == null ? '' : v).trim();
            if (!s) return true;
            if (/^\[.*\]$/.test(s)) return true;
            if (/^(XX+|xxx+|\[x+\]|TBD|待定|示例|样例)$/i.test(s)) return true;
            if (/^[\s\u3000]+$/.test(s)) return true;
            return false;
        };
        const cleaned = slides.map((s) => {
            if (!s) return s;
            const type = String(s.type || '').toLowerCase();
            if (type === 'kpi' && Array.isArray(s.items)) {
                s.items = s.items.map((it) => {
                    let value = String(it && it.value !== undefined ? it.value : '');
                    if (value.includes('、') || (value.includes(',') && /[0-9]/.test(value))) {
                        const first = String(value.split(/[、,]/)[0]).trim();
                        warnings.push(`KPI「${it && it.label}」value 含多值，已保留首值「${first}」`);
                        return Object.assign({}, it, { value: first });
                    }
                    if (isPlaceholder(value)) {
                        warnings.push(`KPI「${it && it.label}」value 为空/占位，已回退为 —`);
                        return Object.assign({}, it, { value: '—' });
                    }
                    return it;
                });
            } else if ((type === 'chart' || type === 'comparison') && s && typeof s === 'object') {
                // chart/comparison 的数据域引用须与上一轮分析结果或字段集对齐：
                // series/items 里出现的"分类/维度/字段名"若既不在分析关键词集也不在表头，判为编造 → 清空该序列并提示
                const collect = (it) => {
                    if (it && typeof it === 'object') {
                        for (const f of ['name', 'category', 'label', 'field']) {
                            if (typeof it[f] === 'string' && it[f]) {
                                const low = it[f].toLowerCase();
                                if (!headerSet.has(low) && !keys.has(it[f])) {
                                    if (keys.size > 0 || headerSet.size > 0) {
                                        warnings.push(`${type}页「${it[f]}」字段不在分析结果/表头，已回退`);
                                        return Object.assign({}, it, { [f]: '—' });
                                    }
                                }
                            }
                        }
                    }
                    return it;
                };
                const arrKey = type === 'chart' ? 'series' : 'items';
                const arr = Array.isArray(s[arrKey]) ? s[arrKey] : null;
                if (arr) s[arrKey] = arr.map(collect);
            }
            return s;
        });
        return { slides: cleaned, warnings };
    }

    /** 依据上一轮分析的技能名/任务名推断分析类型，用于选择故事线模板（无法判断则用通用） */
    function deriveAnalysisKind(prevSpec) {
        if (!prevSpec) return 'generic';
        const names = [];
        const ts = prevSpec.taskStructure;
        if (ts && Array.isArray(ts.tasks)) {
            for (const t of ts.tasks) {
                if (t && (t.displayName || t.templateId)) names.push(t.displayName || t.templateId);
            }
        }
        const p = prevSpec.params || {};
        if (p.displayName) names.push(p.displayName);
        if (p.skillName) names.push(p.skillName);
        const text = names.join(' ');
        if (/差异|diff/i.test(text)) return 'diff';
        if (/清洗|异常|缺失|重复|离群|质量|智能数据分析/i.test(text)) return 'anomaly_quality';
        if (/描述性统计|统计|均值|中位数|标准差|分位数/i.test(text)) return 'desc_stats';
        if (/占比|构成|份额|结构/i.test(text)) return 'comparison_share';
        if (/趋势|预测|时序|走势/i.test(text)) return 'trend';
        return 'generic';
    }

    /**
     * PPT 规划提示词（跨管道继承分支）：
     * 融合 ①复合任务组合叙事 ②数据特征动态注入 ③结构约束 ④推荐结构基线
     * ⑤受众分层叙事 ⑥维度驱动补拼 ⑦结构化页级大纲种子，
     * 与 generateSlideSpecViaLlm 的 baseline/analysisBlock 协同，prompt 多端给 LLM 一致信号。
     * opts：{ directSpec, columnHints }，提供列结构以做维度驱动；缺省则仅按文本特征。
     */
    function buildPptPlanningMessage(baseMessage, analysisContext, prevSpec, recommendedSlides, opts) {
        if (!analysisContext) return baseMessage;
        const o = opts || {};
        const audience = deriveAudience(baseMessage);
        const dims = deriveDimensionFeatures(o.directSpec, o.columnHints);
        const story = composeStoryline(
            deriveAnalysisKinds(prevSpec),
            detectDataFeatures(analysisContext),
            audience,
            dims
        );
        let out = (baseMessage || '') +
            '\n\n【上一轮分析结论（请据此规划 PPT 页面，把关键指标与统计结论呈现出来）】\n' +
            analysisContext +
            '\n\n请基于以上分析结论规划 PPT 页面结构，优先呈现关键指标、统计结论与洞察，\n' +
            '让 PPT 讲清"指标 → 依据 → 结论"的完整叙事，不要为每个数据列单独建章节+表格。\n' +
            '\n【分析类型与推荐叙事结构】检测到本次分析类型：' + story.label + '\n' +
            '推荐按以下页面结构组织（可依据用户具体诉求增删页面，但请保持叙事完整性）：\n' +
            story.body;

        // 受众/场景分层叙事
        if (story.audienceLabel) {
            out += '\n\n【受众叙事视角】本次面向：' + story.audienceLabel + '\n' +
                '叙事重心：' + story.audienceLead + '\n' + story.audienceStructure;
        }

        // 结构化的页级大纲种子：给 LLM 一个"每页该放什么数据"的骨架，直接用它填充数值
        if (Array.isArray(story.pages) && story.pages.length) {
            const outlineLines = story.pages.map(function (p, i) {
                const chartT = p.chartType ? '(' + p.chartType + ')' : '';
                return '- 第' + (i + 1) + '页 type=' + p.type + chartT + '「' + p.title + '」→ ' + p.hint;
            }).join('\n');
            out += '\n\n【页级大纲种子（请以其为骨架填充具体数据，直接取上一轮分析结果，可增删调整）】\n' +
                outlineLines;
        }

        if (Array.isArray(recommendedSlides) && recommendedSlides.length > 0) {
            out += '\n\n【推荐结构基线（已按数据特征自动生成，请以其为骨架）】\n' +
                JSON.stringify(recommendedSlides) + '\n' +
                '- 以该页面序列为骨架，按用户叙事与分析结论增删调整；KPI/对比/图表数值直接取分析结果\n' +
                '- 它是起点参考而非照搬：最终结构 = 该基线经用户需求优化后的版本';
        }
        out += '\n\n【选择性输出原则（必须遵守）】\n' +
            '- 按需输出而非全量堆页：默认只规划核心页（封面 + 关键分析 1~2 页 + 总结 + 结尾），总页数控制在 6~9 页；\n' +
            '- 目录、对比分析、数据明细等辅助页仅当用户明确要求时才包含；\n' +
            '- 同一数据维度只保留最有价值的一种呈现，不要"章节+图表"重复堆砌；\n' +
            '- 用户只关心部分内容时（如只看趋势/只看某部门），只规划相关页面，其余一律省略。';
        return out + '\n\n' + buildStructureConstraints();
    }

    /**
     * 反死模板：检测 LLM 输出的"一列一章+表格"套路并折叠为合理结构。
     * 特征：≥3 个 section 页，且每个 section 后紧跟同主题 table 页（如
     * 章节:销售额统计 | 数据表格:销售额统计 × 每列）。
     * 折叠为：封面/目录/结尾保留 + KPI 概览 + 数据明细表 + 结论要点。
     * 仅当特征明显时折叠，避免误伤正常的多章节结构。
     */
    function collapseRigidSections(slides) {
        if (!Array.isArray(slides) || slides.length < 6) return slides;
        const types = slides.map(s => String(s?.type || '').toLowerCase());
        let sectionCount = 0;
        let sectionTablePairs = 0;
        for (let i = 0; i < types.length; i++) {
            if (types[i] === 'section') {
                sectionCount++;
                if (types[i + 1] === 'table') sectionTablePairs++;
            }
        }
        if (sectionCount < 3 || sectionTablePairs < 3) return slides;

        const kept = [];
        const seen = new Set();
        for (const s of slides) {
            const t = String(s?.type || '').toLowerCase();
            if (t === 'title' || t === 'agenda' || t === 'closing') {
                kept.push(s);
            } else if (t === 'kpi' && !seen.has('kpi')) {
                kept.push(s);
                seen.add('kpi');
            }
        }
        const firstTable = slides.find(s => String(s?.type || '').toLowerCase() === 'table');
        if (firstTable) {
            kept.push({
                type: 'table',
                title: '数据明细',
                columns: firstTable.columns,
                rowCount: firstTable.rowCount,
                startRow: 2
            });
        }
        const lastText = [...slides].reverse().find(s => String(s?.type || '').toLowerCase() === 'text');
        if (lastText) kept.push(lastText);
        ChatLog.info?.('PIPELINE', '[code-pipeline] 反死模板: 检测到"一列一章+表格"套路，已折叠为 KPI+明细+结论');
        return kept.length >= 3 ? kept : slides;
    }

    /**
     * 【LLM 混合】需求驱动规划：让 LLM 产出幻灯片规格 JSON 数组（匹配后端布局引擎契约）。
     * - previousSlides：后续消息继承，保留原页按诉求增删改
     * - baselineSlides：规则引擎推荐结构作为参考起点（可调整，非照搬）
     * 失败返回 null（调用方回退规则解析结果）。
     */
    async function generateSlideSpecViaLlm(userMessage, directSpec, signal, layoutContext, previousSlides, baselineSlides) {
        try {
            const llmClient = App.ChatV2?.llmClient;
            if (!llmClient || typeof llmClient.complete !== 'function') return null;

            const headers = Array.isArray(directSpec?.headers) ? directSpec.headers : [];
            const rows = Array.isArray(directSpec?.rows) ? directSpec.rows : [];
            const sample = rows.slice(0, 5).map(r => {
                const o = {};
                for (const h of headers.slice(0, 8)) o[h] = (r && r[h] !== undefined) ? String(r[h]) : '';
                return o;
            });

            const layoutBlock = layoutContext
                ? `\n公司模板版式画像（规划时必须对齐该结构）：\n${layoutContext}\n
严格容量约束（每页内容不得超过上限，超出部分拆分到多页或只保留重点）：
- 要点页 bullets 条数 ≤ 容量上限中的"每页要点"
- 表格页 rowCount ≤ 容量上限中的"表格行"（超出时只保留前 N 行）
- 图表页 series 数组元素数 ≤ 容量上限中的"图表系列"
- 图表分类数 ≤ 容量上限中的"分类"\n`
                : '';

            // 后续消息继承：把上一版页面结构交给 LLM，在保留原页的基础上应用补充/修改诉求
            const prevBlock = (Array.isArray(previousSlides) && previousSlides.length > 0)
                ? `\n这是上一版 PPT 的页面结构（JSON）：\n${JSON.stringify(previousSlides)}\n
用户正在对上一版 PPT 提出补充/修改诉求。请保留原有页面，按诉求增删改页面，
输出完整的新页面列表（包含所有保留页 + 新增/修改页），不要只输出改动部分。\n`
                : '';

            // 需求驱动基线：当前参考结构作为起点（LLM 可调整，非照搬）
            const baselineBlock = (Array.isArray(baselineSlides) && baselineSlides.length > 0)
                ? `\n以下是当前参考结构（JSON，仅作起点参考，必须根据用户需求调整，不要照搬）：
${JSON.stringify(baselineSlides)}\n`
                : '';

            // 跨管道继承分析结果时：KPI/图表/要点/对比页必须直接引用上一轮分析给出的具体指标与结论，
            // 而不是拿样本原始数据重算或套模板，保证 PPT 内容与分析结论一致（内容完整度）。
            const hasAnalysis = /\u3010上一轮分析/.test(String(userMessage || ''));
            const analysisBlock = hasAnalysis
                ? `\n分析数据使用（当"用户需求"中包含【上一轮分析】时，以下为强制规则）：
- 【上一轮分析结果】是数字与结论的权威来源：KPI、图表、对比、要点页里的具体数值，优先直接引用分析结果已给出的统计指标（均值/中位数/标准差/极值/占比/预测值/Top值/分组数字等），严禁用样本原始数据重新计算。
- 每个 KPI 项的 value 只能是单个数值，直接从分析结果取值；label 写明指标名（如"总销售额""平均增长率"），多个预测/未来值要拆成独立 item。
- text 要点页(bullets)直接采用分析结果里的结论性语句，提炼为短句，不要复述原始明细。
- comparison 页的 leftItems/rightItems 直接放分析结果里成对的对比数字/占比。
- chart 页依据分析结果已给出的趋势/构成/占比确定 chartType 与 series，并用具体结论作为标题表述。
- 只有当用户明确要"原始明细/完整/全部数据"时才输出明细表(table)，rowCount 控制在容量内。
- 分析结果未提供的指标用"—"或省略，禁止凭空编造数值。\n`
                : '';

            const prompt = `你是 PPT 结构设计师。根据用户需求和数据字段，输出幻灯片规格 JSON 数组。

用户需求：${userMessage}
数据字段：${headers.join('、')}
样本数据：${JSON.stringify(sample)}
${layoutBlock}${prevBlock}${baselineBlock}${analysisBlock}
结构原则（必须严格遵守）：
- 页面结构必须服务于用户的具体需求，不同需求用不同结构，禁止千篇一律
- 禁止为每个数据列单独建"章节+表格"页（那是死模板）
- 页面总数 5-8 页，结构服务于分析目的，不是罗列字段
- 根据用户需求选择侧重点：
  · 用户要"趋势/走势/变化/增长"→ 用折线图页突出时间变化，弱化其他章节
  · 用户要"对比/比较"→ 用对比页或柱状图页呈现差异
  · 用户要"占比/构成/份额"→ 用饼图页呈现结构占比
  · 用户要"汇报/总结/建议"→ 多用要点页呈现结论与建议
  · 用户要"明细/完整数据/原始数据"→ 保留数据明细表页
  · 用户没特别要求 → 用 KPI + 核心图表 + 明细 + 结论的通用结构
- 章节页(section)只用于真正需要分组的场景，最多 2-3 个
- 优先用 KPI 页呈现核心指标、图表页呈现趋势/构成、表格页呈现明细
每页 type 必须是以下之一：title/agenda/section/kpi/chart/table/text/comparison/timeline/closing。
字段约定：
- title: {"type":"title","title":"...","subtitle":"..."}
- agenda: {"type":"agenda","title":"目录","items":["章节1","章节2"]}
- section: {"type":"section","number":"01","title":"章节名","subtitle":"..."}
- kpi: {"type":"kpi","title":"关键指标","items":[{"label":"指标名","value":"数值"}]}
- KPI 项约束（必须严格遵守）：
  · 每个 item 的 value 只能是单个数值，禁止把多个数值塞进同一个 value（如"29958.44、29960.63、29962.81"），否则渲染时文字溢出覆盖 label
  · 上一轮分析含预测/未来值（如趋势预测的多个未来销售额）时，把每个预测值拆成独立 item，label 用其日期或"预测第N天"，value 用该值
- chart: {"type":"chart","title":"图表标题","chartType":"line|column|pie","xField":"X轴字段","series":["数值字段"]}
- chart 页可选"次坐标轴"（secondaryAxis）：仅当是"折线+柱状"等量纲差异大的组合图时，把量级悬殊的系列放入 secondaryAxis 数组（值为系列字段名，与 series 一致），让该系列绑到右侧次轴；同量纲或单系列一律不要用（会误导）。例：chartType 用 column，series:["销售额","利润"]，secondaryAxis:["利润"]。
- table: {"type":"table","title":"数据明细","columns":["字段1"],"rowCount":行数,"startRow":2}
- text: {"type":"text","title":"要点","bullets":["要点1"]}
- comparison: {"type":"comparison","title":"对比","leftTitle":"A","rightTitle":"B","leftItems":[],"rightItems":[]}
- timeline: {"type":"timeline","title":"时间线","items":[{"time":"时间","text":"内容"}]}
- closing: {"type":"closing","title":"谢谢","subtitle":"..."}

只输出 JSON 数组，不要任何其他文字、解释或代码块标记。`;

            const events = await llmClient.complete({
                messages: [
                    { role: 'system', content: '你是 PPT 结构设计师，只输出 JSON 数组，不输出任何其他内容。' },
                    { role: 'user', content: prompt }
                ],
                config: App.config,
                signal
            });

            let text = '';
            for (const evt of events) {
                if (evt.type === 'error') return null;
                if (evt.type === 'token') text += evt.delta || '';
            }

            const slides = parseSlidesJson(text);
            if (!slides || slides.length === 0) return null;

            // 白名单校验：只保留后端支持的版式类型
            const valid = slides.filter(s => s && SUPPORTED_SLIDE_TYPES.has(String(s.type || '').toLowerCase()));
            if (valid.length === 0) return null;
            // 反死模板：折叠"一列一章+表格"套路（小模型常见输出）
            return collapseRigidSections(valid);
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[code-pipeline] LLM 幻灯片规格生成失败，回退规则:', e);
            return null;
        }
    }

    /** 从 LLM 输出中提取并解析 JSON 数组（容忍 ```json 围栏与前后杂文本） */
    function parseSlidesJson(text) {
        if (!text) return null;
        const fenced = text.match(/```(?:json)?\s*(\[[\s\S]*?\])\s*```/);
        const candidate = fenced ? fenced[1] : text;
        const start = candidate.indexOf('[');
        const end = candidate.lastIndexOf(']');
        if (start < 0 || end <= start) return null;
        try {
            const arr = JSON.parse(candidate.slice(start, end + 1));
            return Array.isArray(arr) ? arr : null;
        } catch (e) {
            return null;
        }
    }

    /**
     * 从推荐 slides 提取内容页勾选状态（结构页不参与勾选，封面默认勾选）。
     */
    function buildLayoutFromSlides(slides) {
        if (!Array.isArray(slides) || slides.length === 0) return null;
        const layout = { title: true };
        const contentTypes = ['table', 'text', 'chart', 'kpi', 'comparison', 'timeline'];
        for (const s of slides) {
            const t = String(s?.type || '').toLowerCase();
            if (contentTypes.includes(t)) layout[t] = true;
        }
        return Object.keys(layout).length > 1 ? layout : null;
    }

    /**
     * 按用户勾选的版式过滤推荐 slides。
     * 封面（title）按勾选保留；目录/章节/结尾（agenda/section/closing）作为叙事骨架始终保留，
     * 但孤立章节页（其后无保留内容）剔除；内容页仅保留勾选类型。
     */
    function filterSlidesByLayout(slides, checkedTypes) {
        if (!Array.isArray(slides)) return [];
        const alwaysKeep = new Set(['agenda', 'section', 'closing', 'cover', 'end']);
        const kept = [];
        for (const s of slides) {
            const t = String(s?.type || '').toLowerCase();
            if (t === 'title') {
                // 封面：默认保留，用户取消勾选"标题页"则移除
                if (!checkedTypes || checkedTypes.title !== false) kept.push(s);
            } else if (alwaysKeep.has(t)) {
                kept.push(s);
            } else if (checkedTypes && checkedTypes[t]) {
                kept.push(s);
            }
        }
        // 剔除孤立章节页：章节页后若直接遇到另一结构页（非 section）或结尾，则删除
        return kept.filter((s, i) => {
            if (String(s?.type || '').toLowerCase() !== 'section') return true;
            for (let j = i + 1; j < kept.length; j++) {
                const nt = String(kept[j]?.type || '').toLowerCase();
                if (alwaysKeep.has(nt) && nt !== 'section') return false;
                if (!alwaysKeep.has(nt)) return true;
            }
            return false;
        });
    }

    /**
     * 等待用户在配置卡上确认主题/版式。
     * 渲染器点击"生成 PPT / 使用默认生成"时 resolve(config)；
     * 用户取消（abort）时 resolve(null)。
     */
    function waitForPptConfig(turnId, signal) {
        return new Promise((resolve) => {
            const waiters = App.ChatV2._pptConfigWaiters || (App.ChatV2._pptConfigWaiters = new Map());
            waiters.set(turnId, { resolve });
            if (signal) {
                signal.addEventListener('abort', () => {
                    if (waiters.has(turnId)) {
                        waiters.delete(turnId);
                        resolve(null);
                    }
                }, { once: true });
            }
        });
    }

    /**
     * 供渲染器调用：用户点击配置卡按钮后，把所选配置交给等待中的 pipeline。
     */
    function resolvePptConfig(turnId, config) {
        const waiters = App.ChatV2._pptConfigWaiters;
        const w = waiters?.get(turnId);
        if (w) {
            waiters.delete(turnId);
            w.resolve(config);
        }
    }

    function detectLanguage(content, skill, userMessage) {
        // GeneratePPTReport 输出 JSON 描述 PPT 结构，VBA 由模板引擎渲染
        if (skill && (skill.name === 'GeneratePPTReport' || skill.Name === 'GeneratePPTReport')) {
            return 'json';
        }

        // 优先级 1：用户消息意图 > skill 标签
        const userMsg = (userMessage || '').toLowerCase();
        if (/vba|宏|excel.*代码|excel.*脚本|excel.*程序|ppt|powerpoint|幻灯片|演示文稿|生成.*报告/.test(userMsg)) {
            return 'vba';
        }
        if (/python|py脚本|py代码/.test(userMsg)) {
            return 'python';
        }

        if (!skill) return 'code';
        const outputType = (skill.outputType || skill.OutputType || '').toLowerCase();
        const outputFormat = (skill.outputFormat || skill.OutputFormat || '').toLowerCase();
        if (outputType === 'python_script' || outputFormat === 'python') return 'python';
        if (outputType === 'vba_code' || outputFormat === 'vba') return 'vba';
        if (outputFormat === 'json') return 'json';
        const tags = Array.isArray(skill.tags) ? skill.tags.join(',').toLowerCase()
            : Array.isArray(skill.Tags) ? skill.Tags.join(',').toLowerCase()
            : '';
        if (tags.includes('python')) return 'python';
        if (tags.includes('vba')) return 'vba';
        if (tags.includes('general')) return 'code';
        return 'code';
    }

    /**
     * 根据最终落定的代码内容重新判定语言，避免渲染前的语言猜测与真实内容不符。
     * 典型场景：GeneratePPTReport 在渲染前被 detectLanguage 判成 'json'，
     * 但模板引擎实际渲染出的是 VBA 代码，若仍带 'json' 标签渲染会让代码块显示为 JSON。
     * @param {string} code - 最终代码内容
     * @param {string} fallback - 无法判定时回退的原有语言
     * @returns {string} 重新判定后的语言，'' 表示不是代码（走普通文本渲染）
     */
    function redetectCodeLang(code, fallback) {
        const t = (code || '').trim();
        if (!t) return fallback;

        // 模板兜底提示是 Markdown 文本而非代码：清空语言，走普通文本渲染
        if (/^⚠️/.test(t) || /无法自动生成/.test(t)) return '';

        // VBA（模板渲染出的代码通常以 Sub/Function 或 End Sub 结尾）
        if (/^\s*(Sub|Function|Dim|Public|Private|Const|Type)\s/i.test(t) ||
            /End\s+(Sub|Function|Property)\s*$/im.test(t)) {
            return 'vba';
        }

        // Python 优先判定，避免被当成 VBA
        if (/\b(def |import |from |class |print\(|@\w+|pandas|numpy)\b/.test(t)) {
            return 'python';
        }

        // 真正的 JSON（IntentSpec / PPT 报告结构）
        if (/^\s*\{[\s\S]*\}\s*$/.test(t)) {
            try {
                const obj = JSON.parse(t);
                if (obj && typeof obj === 'object' &&
                    (obj.scenario || obj.slides || obj.reportTitle || obj.dataSheet)) {
                    return 'json';
                }
            } catch (e) { /* 非 JSON，继续 */ }
        }

        return fallback;
    }

    function isToolRelatedText(text) {
        if (!text) return false;
        const toolPatterns = [
            /为了.{0,20}写入/i,
            /请调用|请使用/i,
            /调用以下工具/i,
            /\[excel_[a-z_]+\]/i,
            /工具[：:]/,
            /function_call|tool_call/i,
            /excel_local_/i,
            /excel_write|excel_execute/i
        ];
        return toolPatterns.some(p => p.test(text));
    }

    function detectCodeEnd(content, lang) {
        if (!content) return false;
        const checkWindow = content.slice(-200);
        if (isToolRelatedText(checkWindow)) return true;
        if (lang === 'vba' && /End\s+(Sub|Function|Property)\s*$/im.test(checkWindow)) return true;
        if (lang === 'python' && /(if __name__\s*==\s*['"]__main__['"]|^#[^\n]*TODO)/m.test(checkWindow)) return true;
        if (lang === 'python' && /\b(def |class )[\s\S]*$/m.test(checkWindow.slice(-50))) {
            // Incomplete - wait for more
            return false;
        }
        return false;
    }

    /**
     * 检测 LLM 输出是否给出了 IntentSpec JSON 或 PPT 报告 JSON。
     * 架构说明：架构重构后，LLM 不再直接生成 VBA 完整代码，
     *           而是输出 IntentSpec JSON（如 {"scenario":"ppt_from_data",...}），
     *           真正的代码由 vba-template-engine 渲染。
     *           这种"LLM 提取参数 + 模板拼装"的分层比让 LLM 一次性写 50 行 VBA
     *           安全得多（模板由人工审过，LLM 失误只影响参数不影响 API 调用）。
     */
    function looksLikeVbaSource(text) {
        if (!text) return false;
        return /\b(?:Sub|Function)\s+\S+/i.test(text) && /\bEnd\s+(?:Sub|Function)\b/i.test(text);
    }

    const RUNNABLE_VBA_SYSTEM = `你是 Excel VBA 专家。根据用户需求写一段可以原样粘贴进 VBA 编辑器（Alt+F11，插入模块）并直接运行的代码。

硬性要求：
1. 只输出一个 vba 代码块，不要解释。
2. 模块第一行是 Option Explicit。恰好一个入口 Sub，英文过程名，带括号。
3. 每个变量只 Dim 一次，Dim 放在该 Sub 的最前面。禁止 Attribute VB_Name，禁止在过程内部写 Option Explicit。
4. 工作表和区域用上下文里的真实名称；上下文没有就用 ActiveSheet 和 UsedRange，禁止编造不存在的表名。
5. 字符串只用英文双引号。不要用中文引号、中文括号、中文逗号、中文分号。
6. 用 On Error GoTo ErrHandler。过程末尾先 Exit Sub，再写 ErrHandler:，用 MsgBox 显示 Err.Number 和 Err.Description，并恢复 ScreenUpdating。
7. 修改工作表名之前去掉 \\ / ? * [ ] : ，并截断到 31 个字符。
8. 用户没说“选中区域”就不要用 Selection。`;

    async function requestRunnableVba(turn, signal) {
        const llmClient = App.ChatV2?.llmClient;
        if (!llmClient || typeof llmClient.complete !== 'function') return '';
        const ctxBlock = buildContextBlock(turn.context || {});
        const user = `${turn.userMessage || ''}${ctxBlock ? `\n\n${ctxBlock}` : ''}`;
        try {
            const events = await llmClient.complete({
                messages: [
                    { role: 'system', content: RUNNABLE_VBA_SYSTEM },
                    { role: 'user', content: user }
                ],
                config: App.config,
                signal
            });
            let text = '';
            for (const evt of events) {
                if (evt?.type === 'error') return '';
                if (evt?.type === 'token') text += evt.delta || '';
            }
            return text;
        } catch (e) {
            ChatLog.warn?.('PIPELINE', '[code-pipeline] 可运行 VBA 补生成失败:', e);
            return '';
        }
    }

    function isIntentSpecJson(text) {
        if (!text) return false;
        const m = text.match(/```(?:json)?\s*([\s\S]*?)```/);
        const json = m ? m[1].trim() : text.trim();
        if (!json.startsWith('{') || !json.endsWith('}')) return false;
        try {
            const obj = JSON.parse(json);
            // IntentSpec 必须至少包含 scenario 字段或与已知模板关键词匹配
            // PPT 报告 JSON 包含 reportTitle 或 slides 数组
            return obj && (
                typeof obj.scenario === 'string' ||
                typeof obj.dataSheet === 'string' ||
                typeof obj.fields === 'object' ||
                typeof obj.reportTitle === 'string' ||
                Array.isArray(obj.slides)
            );
        } catch (e) {
            return false;
        }
    }

    async function* execute(turn, signal) {
        const { sessionStore } = App.ChatV2;
        const llmClient = App.ChatV2.llmClient;

        if (!sessionStore || !llmClient) {
            throw new Error('code-pipeline 依赖未就绪');
        }

        try {
            ChatLog.info?.('PIPELINE', `[code-pipeline] turn=${turn.id} 开始生成代码`);

            // 生命周期（Phase 4）：新任务（非后续消息）时清空上一轮执行上下文，
            // 防止跨任务误继承（P6）。后续消息（_execFollowUp）保留上下文供继承。
            if (!turn.context?._execFollowUp && sessionStore.clearExecSpec) {
                sessionStore.clearExecSpec();
            }

            const matchedSkill = turn.context?._matchedSkill;
            let lang = detectLanguage('', matchedSkill, turn.userMessage);

            // ============================================================
            // 代码后续消息继承（Phase 2）：上一轮有 code 执行上下文（_execSpec）
            // 时，把上轮代码注入 LLM 提示，让它基于原代码修改/追加而非从零生成；
            // 同时继承上一轮语言与数据快照（无新选区时）。
            // ============================================================
            let prevCode = null;
            let ctxForMsg = turn.context || {};
            if (turn.context?._execFollowUp) {
                const prevExec = sessionStore.getLastExecSpec ? sessionStore.getLastExecSpec() : null;
                if (prevExec && prevExec.spec?.pipeline === 'code' && prevExec.spec.result?.code) {
                    prevCode = prevExec.spec.result.code;
                    const prevLang = prevExec.spec.params?.lang;
                    if (prevLang === 'vba' || prevLang === 'python') lang = prevLang;
                    const pd = prevExec.spec.data || {};
                    if (!ctxForMsg.selection && pd.selection) ctxForMsg = { ...ctxForMsg, selection: pd.selection };
                    if (!ctxForMsg.referenceRegions && pd.referenceRegions) ctxForMsg = { ...ctxForMsg, referenceRegions: pd.referenceRegions };
                    ChatLog.info?.('PIPELINE', `[code-pipeline] 代码后续消息继承: 复用上一轮代码 (lang=${lang})`);
                }
            }

            const messages = buildMessages(turn, lang, prevCode, ctxForMsg);
            let generatedCode = '';
            let reasoningText = '';
            let tokenReceived = false;
            let captureStopped = false;

            yield { type: 'round_start', round: 1, totalRounds: 1 };

            const langLabels = { python: 'Python', vba: 'VBA', json: 'JSON', code: '代码' };
            const langLabel = langLabels[lang] || '代码';

            // 标记此 turn 为代码响应（供渲染层识别）
            sessionStore.updateTurn(turn.id, { _codeLang: lang });

            // ============================================================
            // 【PPT 架构直出】GeneratePPTReport + 已提供选区：不再生成/编译 VBA，
            // 直接把 报告标题+数据范围 交给后端原生 COM 生成 PPT（固定主题、少报错）。
            // 规避小模型"反问/编造工具/输出非 JSON"以及模板 VBA"编译报错"两类问题。
            // ============================================================
            // 构建直出 spec：优先当前选区；后续消息无选区时从上一轮执行上下文继承（通用机制）
            let directSpec = buildDirectIntentSpec(turn.context, turn.userMessage);
            const prevExec = sessionStore.getLastExecSpec ? sessionStore.getLastExecSpec() : null;
            // PPT 后续消息继承：同管道（ppt）或跨管道（skill/text/python 分析结果 → PPT）均视为后续消息
            const isPptFollowUp = !!(turn.context?._execFollowUp) && !!prevExec &&
                (prevExec.spec?.pipeline === 'ppt' ||
                 prevExec.spec?.pipeline === 'skill' ||
                 prevExec.spec?.pipeline === 'text' ||
                 prevExec.spec?.pipeline === 'python');
            if (!directSpec && isPptFollowUp && prevExec.spec.data) {
                directSpec = buildDirectSpecFromExecData(prevExec.spec.data);
                if (directSpec) {
                    ChatLog.info?.('PIPELINE', '[code-pipeline] PPT 后续消息继承: 复用上一轮数据上下文');
                }
            }
            // PPT 跨管道继承（Phase 5 扩展）：上一轮是分析类管道（skill/text/python）时，
            // 把上一轮分析结论 + 任务结构注入规划提示，让 PPT 按分析结构组织章节、
            // 呈现分析结果而非通用数据表；报告标题取上一轮技能名（如"描述性统计报告"）。
            let pptAnalysisContext = '';
            if (isPptFollowUp && prevExec && prevExec.spec.pipeline !== 'ppt') {
                const prevRes = prevExec.spec.result || {};
                const analysisText = prevRes.outputText || prevRes.summary || prevRes.content || '';
                if (analysisText) {
                    pptAnalysisContext = String(analysisText).slice(0, 3000);
                }
                // 注入上一轮分析任务结构（任务清单/分析步骤），让 PPT 章节对齐分析规划
                const prevTs = prevExec.spec.taskStructure;
                if (prevTs && Array.isArray(prevTs.tasks) && prevTs.tasks.length) {
                    const planLines = prevTs.tasks.map(function (t, i) {
                        return (i + 1) + '. ' + (t.displayName || t.templateId || '分析');
                    }).join('\n');
                    pptAnalysisContext = (pptAnalysisContext ? pptAnalysisContext + '\n\n' : '') +
                        '【上一轮分析任务规划（请按此结构组织 PPT 章节）】\n' + planLines;
                }
                // 注入上一轮分析的完整结构化结果（统计指标/分布数据/结论文本），
                // 让 LLM 规划 PPT 时直接引用这些数据填充 KPI/图表/要点页，而非用原始数据重新组装
                const prevDetails = prevRes.details;
                if (Array.isArray(prevDetails) && prevDetails.length) {
                    const detailLines = prevDetails.map(function (d) {
                        return '【' + (d.displayName || d.templateId || '分析') + '】\n' + (d.detail || '');
                    }).join('\n\n');
                    pptAnalysisContext = (pptAnalysisContext ? pptAnalysisContext + '\n\n' : '') +
                        '【上一轮分析结果（请直接引用这些数据填充 KPI/图表/要点页，不要用原始数据重新计算）】\n' +
                        detailLines.slice(0, 6000);
                }
                // 意图分析：结合用户新需求 + 上一轮任务结构，规划 PPT 章节（而非只靠关键词）
                const intentAnalyzer = App.ChatV2.intentAnalyzer;
                if (intentAnalyzer && typeof intentAnalyzer.buildFollowUpPlanningPrompt === 'function') {
                    const intent = (turn.context && turn.context._execFollowUp && turn.context._execFollowUp.intent) || null;
                    const planning = intentAnalyzer.buildFollowUpPlanningPrompt({
                        prevSpec: prevExec.spec,
                        userMessage: turn.userMessage || '',
                        intent: intent
                    });
                    if (planning) {
                        pptAnalysisContext = (pptAnalysisContext ? pptAnalysisContext + '\n\n' : '') + planning;
                    }
                }
                const prevParams = prevExec.spec.params || {};
                const analysisName = prevParams.displayName || prevParams.skillName || '';
                if (analysisName && directSpec && !directSpec.reportTitle) {
                    directSpec = { ...directSpec, reportTitle: analysisName };
                }
                ChatLog.info?.('PIPELINE', `[code-pipeline] PPT 跨管道继承: 注入上一轮 ${prevExec.spec.pipeline} 分析结论与任务规划`);
            }
            if (directSpec) {
                ChatLog.info?.('PIPELINE', '[code-pipeline] 命中 PPT 直出路径（GeneratePPTReport+选区），请求数据感知设计推荐');

                // 0. 解析用户需求（规则层）：图表类型 / 强调重点 / 页面类型 / 章节
                const reqParser = App.ChatV2?.services?.pptRequirementParser;
                const requirements = reqParser ? reqParser.parsePptRequirements(turn.userMessage) : null;
                const recognized = Array.isArray(requirements?.recognized) ? requirements.recognized : [];

                // 0.5. LLM 列理解：让大模型思考"该拿什么列"（主角指标/分类/时间/ID 列），
                //      覆盖规则引擎的启发式选列；失败静默回退规则引擎
                let columnHints = null;
                if (directSpec) {
                    columnHints = await analyzeColumnsViaLlm(turn.userMessage, directSpec, signal);
                    if (columnHints) {
                        ChatLog.info?.('PIPELINE', '[code-pipeline] LLM 列理解成功:', columnHints);
                    }
                }

                // 1. 请求后端数据感知设计推荐（分析数据 → 版式/图表/洞察，应用 LLM 列理解与用户需求偏置）
                const design = await requestPptDesign(directSpec, columnHints, turn.userMessage);
                let recommendedSlides = design?.slides || null;
                const insights = Array.isArray(design?.insights) ? design.insights : [];
                // 数据感知页面方向：竖版长表 → portrait，宽表/均衡 → landscape（后端按此设置页面尺寸）
                const recommendedOrientation = design?.orientation === 'portrait' ? 'portrait' : 'landscape';

                // 1b. 需求驱动混合策略（反死模板）：
                //   - 后续消息继承：LLM 在上一轮设计基础上应用补充/修改诉求（保留原页增删改）
                //   - 其他所有情况：LLM 作为首选规划器（需求驱动），规则引擎推荐作为基线参考，
                //     让结构随用户要求变化而非固定套路；LLM 失败时回退规则引擎+需求应用
                if (isPptFollowUp) {
                    ChatLog.info?.('PIPELINE', '[code-pipeline] PPT 后续消息继承: LLM 在上一轮设计上应用诉求');
                    const prevSlides = prevExec?.spec?.params?.recommendedSlides || null;
                    const llmSlides = await generateSlideSpecViaLlm(buildPptPlanningMessage(turn.userMessage, pptAnalysisContext, prevExec?.spec, prevSlides, { directSpec, columnHints }), directSpec, signal, null, prevSlides);
                    if (llmSlides && llmSlides.length > 0) {
                        recommendedSlides = llmSlides;
                        recognized.push('（LLM 按后续诉求规划）');
                    }
                } else {
                    // 需求驱动：LLM 首选规划，规则引擎推荐作为基线参考（可调整，非照搬）
                    ChatLog.info?.('PIPELINE', '[code-pipeline] 需求驱动: LLM 按用户需求规划页面（规则推荐为基线）');
                    const llmSlides = await generateSlideSpecViaLlm(
                        buildPptPlanningMessage(turn.userMessage, pptAnalysisContext, prevExec?.spec, recommendedSlides, { directSpec, columnHints }),
                        directSpec, signal, null, null, recommendedSlides
                    );
                    if (llmSlides && llmSlides.length > 0) {
                        recommendedSlides = llmSlides;
                        recognized.push('（LLM 按需求规划）');
                    } else if (requirements && recommendedSlides) {
                        // LLM 失败回退：规则引擎 + 需求应用
                        recommendedSlides = reqParser.applyRequirementsToDesign(recommendedSlides, requirements);
                    }
                }

                // 页内数据映射校验：KPI 单值、占位回退 "—"、字段与上一轮分析结果/表头对齐
                {
                    const mapped = validateSlideDataMapping(recommendedSlides, pptAnalysisContext, directSpec?.headers);
                    if (mapped.warnings && mapped.warnings.length) {
                        ChatLog.info?.('PIPELINE', '[code-pipeline] 页内数据映射校验修正: ' + mapped.warnings.join('; '));
                    }
                    recommendedSlides = mapped.slides;
                }

                // 2. 渲染配置卡：预填推荐设计的主题 + 版式勾选（后续消息预填上一轮主题）
                const defaultTheme = (isPptFollowUp && prevExec?.spec?.params?.theme) || directSpec.theme || 'business_blue';
                const defaultLayout = buildLayoutFromSlides(recommendedSlides) || { title: true, table: true, text: true };
                sessionStore.setPptConfig(turn.id, {
                    reportTitle: directSpec.reportTitle || '数据分析报告',
                    theme: defaultTheme,
                    layout: defaultLayout,
                    recommendedSlides,
                    insights,
                    headers: Array.isArray(directSpec.headers) ? directSpec.headers : [],
                    rows: Array.isArray(directSpec.rows) ? directSpec.rows : [],
                    recognizedRequirements: recognized,
                    chartMode: requirements?.chartMode || (isPptFollowUp ? prevExec?.spec?.params?.chartMode : null) || 'auto',
                    agg: requirements?.agg || (isPptFollowUp ? prevExec?.spec?.params?.agg : null) || 'auto',
                    aggDim: isPptFollowUp ? (prevExec?.spec?.params?.aggDim || '') : '',
                    aggMeasure: isPptFollowUp ? (prevExec?.spec?.params?.aggMeasure || '') : '',
                    aggFunc: isPptFollowUp ? (prevExec?.spec?.params?.aggFunc || '') : '',
                    aggViews: isPptFollowUp && Array.isArray(prevExec?.spec?.params?.aggViews) ? prevExec.spec.params.aggViews : [],
                    aggExcluded: isPptFollowUp && Array.isArray(prevExec?.spec?.params?.aggExcluded) ? prevExec.spec.params.aggExcluded : [],
                    ratio: requirements?.ratio || (isPptFollowUp ? prevExec?.spec?.params?.ratio : null) || '16x9',
                    orientation: requirements?.orientation || (isPptFollowUp ? prevExec?.spec?.params?.orientation : null) || ''
                });

                const config = await waitForPptConfig(turn.id, signal);
                if (!config) {
                    ChatLog.info?.('PIPELINE', '[code-pipeline] 用户取消 PPT 配置，终止');
                    sessionStore.cancelTurn(turn.id, 'ppt_config_canceled');
                    yield { type: 'canceled', reason: 'ppt_config_canceled' };
                    return;
                }

                // PPT 直出路径的输出是状态消息（纯文本），不是 JSON 代码：
                // 把 _codeLang 标成 txt，让状态消息以「TXT」代码块展示而非误标 JSON
                sessionStore.updateTurn(turn.id, { _codeLang: 'txt' });

                const reportTitle = directSpec.reportTitle || '数据分析报告';
                const sheet = directSpec.dataSheet || '';
                const simpleRange = directSpec.dataRange || '';
                // rangeAddress 支持 "Sheet!$A$1:$G$121" 或 "$A$1:$G$121"；已含表名时不再重复拼接
                const rangeAddress = simpleRange.includes('!') ? simpleRange : (sheet ? `${sheet}!${simpleRange}` : simpleRange);
                const theme = config.theme || defaultTheme;
                const applyLayout = config.applyLayout !== false;
                const layoutAwarePlanning = config.layoutAwarePlanning !== false;

                // 2a. 用户逐页配置：对推荐设计中的图表页覆盖 图型/维度/数值/聚合/洞察
                //     按配置卡展示的原推荐索引应用，必须在下面"按索引剔除取消页"之前，否则索引错位
                const _slideOv = config && config.slideOverrides;
                if (_slideOv && typeof _slideOv === 'object' && Array.isArray(recommendedSlides)) {
                    const _hdrSet = new Set((Array.isArray(directSpec?.headers) ? directSpec.headers : []).map(h => String(h)));
                    const _ovWarn = [];
                    Object.keys(_slideOv).forEach((k) => {
                        const idx = parseInt(k, 10);
                        const ov = _slideOv[k];
                        const sl = recommendedSlides[idx];
                        if (!ov || !sl) return;
                        if (String(sl.type || '').toLowerCase() !== 'chart') return;
                        if (ov.chartType) sl.chartType = ov.chartType;
                        if (ov.xField) {
                            if (_hdrSet.size && !_hdrSet.has(String(ov.xField))) {
                                _ovWarn.push(`图表「${sl.title || '图表 ' + (idx + 1)}」的 X 维度「${ov.xField}」不在数据列中，将按默认维度（可能非你预期的时间轴）绘制`);
                            }
                            sl.xField = ov.xField;
                        }
                        if (Array.isArray(ov.series) && ov.series[0]) {
                            if (_hdrSet.size && !_hdrSet.has(String(ov.series[0]))) {
                                _ovWarn.push(`图表「${sl.title || '图表 ' + (idx + 1)}」的数值列「${ov.series[0]}」不在数据列中`);
                            }
                            sl.series = ov.series;
                        }
                        if (ov.agg) sl.agg = ov.agg;
                        if (typeof ov.insight === 'string' && ov.insight) sl.insight = ov.insight;
                    });
                    ChatLog.info?.('PIPELINE', `[code-pipeline] 应用图表页逐页配置: ${Object.keys(_slideOv).length} 页`);
                    if (_ovWarn.length) {
                        _ovWarn.forEach(w => ChatLog.warn?.('PIPELINE', '[code-pipeline] ' + w));
                        yield { type: 'status', text: '⚠️ 部分图表维度未匹配数据列：' + _ovWarn.join('；') };
                    }
                }

                // 2b. 用户页面级勾选：剔除推荐设计中用户取消的页（按索引），后续过滤/版式感知均以剔除后为基线
                if (Array.isArray(config.excludedSlides) && config.excludedSlides.length > 0 && Array.isArray(recommendedSlides)) {
                    const exSet = new Set(config.excludedSlides);
                    recommendedSlides = recommendedSlides.filter((_, i) => !exSet.has(i));
                }

                // 3. 最终 slides：优先按勾选过滤推荐设计；无推荐时回退简单版式
                let finalSlides = recommendedSlides
                    ? filterSlidesByLayout(recommendedSlides, config.layout)
                    : buildSlidesFromConfig(config, directSpec);

                // LLM 版式感知规划：用户开启且所选主题带版式画像时，用版式上下文重新规划页面结构
                if (layoutAwarePlanning) {
                    const engine = App.ChatV2?.services?.vbaTemplateEngine;
                    const layoutSummary = engine && typeof engine.getThemeLayoutSummary === 'function'
                        ? engine.getThemeLayoutSummary(theme)
                        : null;
                    if (layoutSummary) {
                        yield { type: 'status', text: '正在按模板版式优化页面结构...' };
                        // 后续消息继承：把当前设计作为上一版传入，版式优化时保留原页增删改
                        // 其他情况：把当前结构作为基线参考，按版式约束重规划
                        const awareSlides = await generateSlideSpecViaLlm(buildPptPlanningMessage(turn.userMessage, pptAnalysisContext, null, null, { directSpec, columnHints }), directSpec, signal, layoutSummary, isPptFollowUp ? finalSlides : null, finalSlides);
                        if (Array.isArray(awareSlides) && awareSlides.length > 0) {
                            finalSlides = filterSlidesByLayout(awareSlides, config.layout);
                            const remapped = validateSlideDataMapping(finalSlides, pptAnalysisContext, directSpec?.headers);
                            if (remapped.warnings && remapped.warnings.length) {
                                ChatLog.info?.('PIPELINE', '[code-pipeline] 版式感知页映射校验修正: ' + remapped.warnings.join('; '));
                            }
                            finalSlides = remapped.slides;
                        }
                    }
                }
                const slides = applySelectedAggViews(finalSlides, config.aggViews, config.aggExcluded);

                // 强制锁回用户逐页配置（图表类型/维度/度量/聚合/洞察）。
                // layoutAwarePlanning 的 LLM 二次规划会把「趋势」类图表强行改成时间轴，
                // 覆盖掉用户在配置卡里选的 X 维度，导致「配置卡显示地区、实际画时间轴」的不一致。
                // 这里按标题把用户覆盖值重新锁回最终 slides，保证配置卡看到的与最终生成一致。
                {
                    const _ov = config && typeof config.slideOverrides === 'object' ? config.slideOverrides : null;
                    if (_ov && Array.isArray(slides)) {
                        Object.keys(_ov).forEach((k) => {
                            const ov = _ov[k];
                            if (!ov) return;
                            const ovTitle = String(ov.__title || '');
                            if (!ovTitle) return;
                            slides.forEach((sl) => {
                                if (!sl || String(sl.type || '').toLowerCase() !== 'chart') return;
                                if (String(sl.title || '') !== ovTitle) return;
                                if (ov.chartType) sl.chartType = ov.chartType;
                                if (ov.xField) sl.xField = ov.xField;
                                if (Array.isArray(ov.series) && ov.series[0]) sl.series = ov.series;
                                if (ov.agg) sl.agg = ov.agg;
                                if (typeof ov.insight === 'string' && ov.insight) sl.insight = ov.insight;
                            });
                        });
                    }
                }

                yield { type: 'status', text: '正在根据选区生成 PPT...' };

                // 记录生成结果，供写入 _execSpec（多轮继承的 result）
                let genResult = null;

                if (!rangeAddress) {
                    sessionStore.appendDelta(turn.id, 'content', '\n⚠️ **无法确定数据区域**：请在 Excel 中选中要分析的区域后重试。');
                } else {
                    try {
                        // ── 非 Excel 宿主（PowerPoint/Word）分支 ──
                        // 本宿主桥不提供 excelBridge.GeneratePPTFromData（该方法只存在于 ExcelBridge，
                        // Word/PPT 工程的桥是 PptBridge/WordBridge + HostBridgeBase，没有它）。
                        // 实测 PPT 宿主调用必抛 "e is not a function"（宿主代理无此方法）——见
                        // bridge-accessor.js 注释：桥代理对任意属性 typeof 恒为 function，守卫挡不住。
                        // 因此改走本宿主 ppt-local 工具 ppt_generate（与 PptGenerationService 同一渲染引擎），
                        // 并复用多来源桥回退调用器（BridgeAccessor 包装桥在 PPT 宿主不可用，裸桥可用）。
                        const _hostIsExcel = !(window.SmartPilotHost && window.SmartPilotHost.isExcel === false);
                        if (!_hostIsExcel) {
                            const _headers = Array.isArray(directSpec.headers) ? directSpec.headers : [];
                            const _rows = Array.isArray(directSpec.rows) ? directSpec.rows : [];
                            const _hasSlides = Array.isArray(slides) && slides.length > 0;
                            const _caller = (window.App && App.ChatV2 && App.ChatV2.callLocalToolWithFallback) || null;
                            if (typeof _caller !== 'function') {
                                sessionStore.appendDelta(turn.id, 'content', '\n\n❌ 本地工具桥不可用（callLocalToolWithFallback 缺失）：请完全退出 PowerPoint 后重新打开本插件。');
                            } else if (!_hasSlides && (!_headers.length || !_rows.length)) {
                                sessionStore.appendDelta(turn.id, 'content', '\n\n⚠️ **没有可用的数据区域**：请先在 Excel 中选中要分析的数据，再回到 PowerPoint 生成 PPT。');
                            } else {
                                // 只传已在前端拿到的 headers/rows + slides（内联聚合数据），不传 rangeAddress：
                                // PPT 宿主无法读 Excel 选区地址，传数据可完全绕开后端按地址解析。
                                const _payload = {
                                    reportTitle,
                                    requirement: turn.userMessage || '',
                                    autoDesign: true,
                                    theme,
                                    applyLayout,
                                    orientation: config.orientation || recommendedOrientation,
                                    chartMode: config.chartMode || 'auto',
                                    ratio: config.ratio || (requirements?.ratio) || '16x9',
                                    headers: _headers,
                                    rows: _rows
                                };
                                if (_hasSlides) _payload.slides = slides;
                                const _callRes = await _caller('ppt_generate', JSON.stringify(_payload));
                                let gen = {};
                                try { gen = JSON.parse((_callRes && _callRes.raw) || '{}') || {}; } catch { gen = {}; }
                                genResult = gen;
                                if (gen && (gen.success === true || gen.Success === true)) {
                                    const msg = (App.ChatV2?.toolRegistryUtil?.safeDisplayText || ((v) => (typeof v === 'string' ? v : JSON.stringify(v))))(gen.message || gen.Message || 'PPT 已生成');
                                    ChatLog.info?.('PIPELINE', `[code-pipeline] PPT 生成成功（ppt_generate@${_callRes && _callRes.src}）: ${msg}`);
                                    sessionStore.appendDelta(turn.id, 'content', `\n\n✅ ${msg}`);
                                } else {
                                    const errMsg = gen.error || gen.Error || (_callRes && typeof _callRes.raw === 'string' ? _callRes.raw.slice(0, 300) : '') || '生成失败';
                                    ChatLog.warn?.('PIPELINE', `[code-pipeline] PPT 生成失败（ppt_generate）: ${errMsg}`);
                                    sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 生成失败：${errMsg}`);
                                }
                            }
                        } else {
                        // Excel 宿主：统一走 BridgeAccessor.callAsync（已包装 WebView2 代理）。
                        // WebView2 的 hostObjects 代理对"任意属性名" typeof 恒为 'function'、且方法缺失时
                        // 会抛未捕获的 "e is not a function"（栈只有匿名单帧、外层 try/catch 抓不到）。
                        // 包装层把这类失败转成结构化 JSON（success:false），调用方按 success 判定即可。
                        if (window.BridgeAccessor && typeof window.BridgeAccessor.callAsync === 'function') {
                            const argsJson = JSON.stringify({
                                reportTitle,
                                rangeAddress,
                                sheetName: sheet,
                                address: simpleRange,
                                headers: Array.isArray(directSpec.headers) ? directSpec.headers : [],
                                rows: Array.isArray(directSpec.rows) ? directSpec.rows : [],
                                theme,
                                applyLayout,
                                orientation: config.orientation || recommendedOrientation,
                                slides,
                                chartMode: config.chartMode || 'auto',
                                agg: config.agg || 'auto',
                                ratio: config.ratio || (requirements?.ratio) || '16x9'
                            });
                            const genJson = await window.BridgeAccessor.callAsync('GeneratePPTFromData', argsJson);
                            let gen = {};
                            try { gen = JSON.parse(genJson || '{}') || {}; } catch { gen = {}; }
                            genResult = gen;
                            if (gen?.success === true || gen?.Success === true) {
                                const msg = (App.ChatV2?.toolRegistryUtil?.safeDisplayText || ((v) => (typeof v === 'string' ? v : JSON.stringify(v))))(gen.message || gen.Message || 'PPT 已生成');
                                ChatLog.info?.('PIPELINE', `[code-pipeline] PPT 生成成功: ${msg}`);
                                sessionStore.appendDelta(turn.id, 'content', `\n\n✅ ${msg}`);
                            } else {
                                const errMsg = gen.error || gen.Error || (gen.detail && (gen.method + ' — ' + gen.detail)) || '生成失败';
                                // 旧版宿主 DLL 未实现 GeneratePPTFromData（实测报 'e is not a function'）时，
                                // 回落本地 MCP 工具 ppt_generate —— 同一 PptGenerationService 渲染引擎，
                                // 经 CallLocalTool 调用，不依赖该 COM 方法。ppt_generate 不弹 UI，重试无副作用。
                                const _caller = (window.App && App.ChatV2 && App.ChatV2.callLocalToolWithFallback) || null;
                                let _fbDone = false;
                                if (_caller) {
                                    try {
                                        const _payload = {
                                            reportTitle,
                                            requirement: turn.userMessage || '',
                                            autoDesign: true,
                                            theme,
                                            applyLayout,
                                            orientation: config.orientation || recommendedOrientation,
                                            chartMode: config.chartMode || 'auto',
                                            ratio: config.ratio || (requirements?.ratio) || '16x9',
                                            rangeAddress,
                                            headers: Array.isArray(directSpec.headers) ? directSpec.headers : [],
                                            rows: Array.isArray(directSpec.rows) ? directSpec.rows : []
                                        };
                                        if (Array.isArray(slides) && slides.length) _payload.slides = slides;
                                        const _callRes = await _caller('ppt_generate', JSON.stringify(_payload));
                                        let g2 = {};
                                        try { g2 = JSON.parse((_callRes && _callRes.raw) || '{}') || {}; } catch { g2 = {}; }
                                        if (g2 && (g2.success === true || g2.Success === true)) {
                                            genResult = g2;
                                            _fbDone = true;
                                            const _fmt = (App.ChatV2?.toolRegistryUtil?.safeDisplayText || ((v) => (typeof v === 'string' ? v : JSON.stringify(v))));
                                            const okMsg = _fmt(g2.message || g2.Message || 'PPT 已生成');
                                            ChatLog.info?.('PIPELINE', `[code-pipeline] PPT 生成成功（ppt_generate 回落）: ${okMsg}`);
                                            sessionStore.appendDelta(turn.id, 'content', `\n\n✅ ${okMsg}`);
                                        } else {
                                            const e2 = (g2 && (g2.error || g2.Error)) || 'ppt_generate 不可用';
                                            ChatLog.warn?.('PIPELINE', `[code-pipeline] ppt_generate 回落失败: ${e2}`);
                                            sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 生成失败：${errMsg}（ppt_generate 回落也失败：${e2}）`);
                                        }
                                    } catch (fe) {
                                        ChatLog.warn?.('PIPELINE', '[code-pipeline] ppt_generate 回落异常:', fe);
                                        sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 生成失败：${errMsg}（ppt_generate 回落异常：${fe?.message || fe}）`);
                                    }
                                }
                                if (!_fbDone && !_caller) {
                                    ChatLog.warn?.('PIPELINE', `[code-pipeline] PPT 生成失败: ${errMsg}`);
                                    sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 生成失败：${errMsg}`);
                                }
                            }
                        } else {
                            sessionStore.appendDelta(turn.id, 'content', '\n\n❌ BridgeAccessor 不可用，请确认已重新生成后端并重启 Excel。');
                        }
                        } // end else (_hostIsExcel)：Excel 宿主原路径
                    } catch (genErr) {
                        ChatLog.error?.('PIPELINE', '[code-pipeline] PPT 生成异常:', genErr);
                        sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 生成异常：${genErr?.message || String(genErr)}`);
                    }
                }

                // 写入通用执行上下文（多轮继承）：数据快照 + 参数 + 结果，供后续消息继承
                if (sessionStore.setExecSpec) {
                    // v20260924：用工厂统一构建 data，PPT 专用数据挂 data.ext.ppt 命名空间
                    const fact = sessionStore.execSpecFactory;
                    const commonData = fact && typeof fact.buildCommonData === 'function'
                        ? fact.buildCommonData(turn, { ppt: directSpec })
                        : { ext: { ppt: directSpec } }; // 兜底
                    sessionStore.setExecSpec(turn.id, {
                        pipeline: 'ppt',
                        data: commonData,
                        params: {
                            theme,
                            applyLayout,
                            chartMode: config.chartMode || 'auto',
                            agg: config.agg || 'auto',
                            aggDim: config.aggDim || '',
                            aggMeasure: config.aggMeasure || '',
                            aggFunc: config.aggFunc || '',
                            aggViews: Array.isArray(config.aggViews) ? config.aggViews : [],
                            aggExcluded: Array.isArray(config.aggExcluded) ? config.aggExcluded : [],
                            ratio: config.ratio || (requirements?.ratio) || '16x9',
                            recommendedSlides: recommendedSlides || null
                        },
                        result: {
                            success: !!(genResult && (genResult.success === true || genResult.Success === true)),
                            message: (genResult && (genResult.message || genResult.Message)) || '',
                            finalSlides: finalSlides || null
                        },
                        // 任务结构（PPT 页面规划）：供后续消息继承时注入规划输入
                        taskStructure: {
                            slides: recommendedSlides || finalSlides || null
                        },
                        ts: Date.now()
                    });
                }

                sessionStore.completeTurn(turn.id, { reasoning: '' });
                yield { type: 'pipeline_done', pipeline: 'code' };
                return;
            }

            for await (const evt of llmClient.stream({ messages, config: App.config, signal })) {
                if (signal?.aborted) {
                    yield { type: 'canceled', reason: 'user_aborted' };
                    return;
                }

                switch (evt.type) {
                    case 'token': {
                        const delta = evt.delta || '';
                        generatedCode += delta;
                        tokenReceived = true;

                        if (!captureStopped) {
                            sessionStore.appendDelta(turn.id, 'content', delta);
                            if (detectCodeEnd(generatedCode, lang)) {
                                captureStopped = true;
                                ChatLog.info?.('PIPELINE', '[code-pipeline] 检测到代码结束，停止后续展示');
                            }
                        }
                        yield evt;
                        break;
                    }
                    case 'reasoning':
                        reasoningText += evt.delta || '';
                        sessionStore.appendDelta(turn.id, 'reasoning', evt.delta || '');
                        yield evt;
                        break;
                    case 'model':
                        sessionStore.setTurnModel(turn.id, evt.model);
                        yield evt;
                        break;
                    case 'error':
                        yield evt;
                        return;
                    default:
                        yield evt;
                }
            }

            // 提取原始内容
            const rawContent = sessionStore.getCurrentTurn()?.content || '';

            // ============================================================
            // 架构重构：用模板引擎渲染 VBA 代码（取代"让 LLM 一次性写完 50 行 VBA"）
            //
            // 旧问题：让 LLM 负责 VBA 全代码生成
            //   - 同一个 Prompt 越长，模型越容易在长上下文中忘记早期规则（lost-in-the-middle）
            //   - Slides.Add 索引、Layout 常量、Dim 位置等任一细节写错就崩溃
            //   - 加规则只会让 prompt 越来越长，问题越来越严重（补丁困境）
            //
            // 新架构：LLM 只输出 IntentSpec（结构化 JSON），模板引擎渲染代码
            //   - LLM 只负责：理解意图 + 抽取参数
            //   - 模板引擎负责：保证 Sub 结构、变量声明位置、API 调用正确性
            //   - 模板是人工审核过的 VBA 代码片段，覆盖高频场景
            //   - 模板不覆盖的场景才走旧的 LLM 自由生成路径
            // ============================================================
            let finalCode = rawContent;
            let usedTemplate = null;
            let templateFailed = false;
            let templateErrorMessage = '';

            // 健壮性兜底：services 命名空间可能未初始化（脚本加载顺序问题）
            const vbaTplEngine = App.ChatV2?.services?.vbaTemplateEngine;
            // json 模式（GeneratePPTReport）和 vba 模式都走模板引擎
            if ((lang === 'vba' || lang === 'json') && vbaTplEngine && typeof vbaTplEngine.render === 'function') {
                const ctxForTpl = {
                    activeSheet: turn.context?.activeSheet?.sheet,
                    selection: turn.context?.selection
                };

                // 1. 尝试用模板引擎渲染
                if (isIntentSpecJson(rawContent)) {
                    const result = vbaTplEngine.render(
                        rawContent, turn.userMessage, ctxForTpl
                    );
                    if (result.matched && result.code) {
                        finalCode = '```vba\n' + result.code + '\n```';
                        usedTemplate = result.template;
                        ChatLog.info?.('PIPELINE', `[code-pipeline] 模板 ${result.template} 渲染成功，长度=${result.code.length}`);
                    } else {
                        templateFailed = true;
                        templateErrorMessage = result.error || '模板引擎匹配失败（IntentSpec 参数不完整或与模板不兼容）';
                        ChatLog.warn?.('PIPELINE', `[code-pipeline] 模板渲染失败: ${templateErrorMessage}`);
                    }
                } else {
                    // 2. rawContent 不是 IntentSpec JSON（LLM 直接输出了代码），
                    //    尝试用关键词匹配模板
                    const result = vbaTplEngine.render(
                        null, turn.userMessage, ctxForTpl
                    );
                    if (result.matched && result.code) {
                        finalCode = '```vba\n' + result.code + '\n```';
                        usedTemplate = result.template;
                        ChatLog.info?.('PIPELINE', `[code-pipeline] 通过关键词匹配到模板 ${result.template}，已用模板代码覆盖 LLM 输出`);
                    }
                }
            } else if (lang === 'vba' || lang === 'json') {
                // 模板引擎未就绪——记录诊断日志但不影响主流程
                ChatLog.warn?.('PIPELINE', '[code-pipeline] vba-template-engine 未就绪（App.ChatV2.services.vbaTemplateEngine 不可用），回退到 LLM 自由生成');
            }

            // 3. 模板没命中时，不再把模型写出的 VBA 换成“无法生成”提示。
            //    先整理成可粘贴模块；模型只回了 IntentSpec JSON 时，再补一次“只写可运行 VBA”的请求。
            if (lang === 'vba' && !usedTemplate) {
                let source = '';
                if (looksLikeVbaSource(rawContent)) source = rawContent;
                else if (!prevCode) {
                    yield { type: 'status', text: '正在生成可直接粘贴运行的 VBA...' };
                    source = await requestRunnableVba(turn, signal);
                } else {
                    source = rawContent;
                }
                const cleaned = (vbaTplEngine && typeof vbaTplEngine.toRunnableVba === 'function')
                    ? vbaTplEngine.toRunnableVba(source)
                    : '';
                if (cleaned && looksLikeVbaSource(cleaned)) {
                    finalCode = '```vba\n' + cleaned + '\n```';
                    ChatLog.info?.('PIPELINE', `[code-pipeline] 已整理为可直接运行的 VBA，长度=${cleaned.length}`);
                } else if (!prevCode) {
                    finalCode = `⚠️ **这次没能生成可直接运行的 VBA**

请把需求说具体一点，例如要改哪张表、哪一列、结果写到哪里。`;
                    sessionStore.updateTurn(turn.id, { content: finalCode });
                    ChatLog.info?.('PIPELINE', '[code-pipeline] 模板未匹配且补生成失败');
                }
            } else if (usedTemplate && vbaTplEngine && typeof vbaTplEngine.toRunnableVba === 'function') {
                const cleaned = vbaTplEngine.toRunnableVba(finalCode);
                if (cleaned && looksLikeVbaSource(cleaned)) {
                    finalCode = '```vba\n' + cleaned + '\n```';
                }
            }

            const extracted = extractCode(finalCode) || finalCode.trim();

            // 语言标签必须与实际落定的代码一致：渲染前的 detectLanguage 只是猜测
            // （如 GeneratePPTReport 先判成 'json'），模板渲染后真正内容是 VBA 代码。
            // 若仍带 'json' 标签渲染，正文代码块会被错误标成 JSON。
            const actualLang = redetectCodeLang(extracted, lang);
            // 模板兜底提示（⚠️ markdown文本）不是代码，置空 _codeLang 走普通文本渲染
            if (actualLang !== lang) {
                sessionStore.updateTurn(turn.id, { _codeLang: actualLang });
            }

            if (extracted !== rawContent) {
                sessionStore.updateTurn(turn.id, { content: extracted });
            }

            ChatLog.info?.('PIPELINE', `[code-pipeline] 代码生成完成，长度=${extracted.length}` + (usedTemplate ? `, 模板=${usedTemplate}` : ', 自由生成'));

            // 自动执行（Skill 声明 autoExecute: true 时生效）
            if (matchedSkill && (matchedSkill.autoExecute === true || matchedSkill.AutoExecute === true)) {
                ChatLog.info?.('PIPELINE', `[code-pipeline] Skill ${matchedSkill.name} 声明 autoExecute=true，开始自动执行`);
                yield { type: 'status', text: '正在执行生成的代码...' };
                try {
                    const bridge = window.BridgeAccessor && window.BridgeAccessor.getBridge() || (window.chrome?.webview?.hostObjects?.excelBridge || App.excelBridge);
                    if (bridge && typeof bridge.ExecuteVBA === 'function') {
                        const resultJson = await bridge.ExecuteVBA(extracted);
                        const result = JSON.parse(resultJson || '{}');
                        if (result?.success === true || result?.Success === true) {
                            const msg = result.message || result.Message || result.result || result.Result || '执行成功';
                            ChatLog.info?.('PIPELINE', `[code-pipeline] 自动执行成功: ${msg}`);
                            sessionStore.appendDelta(turn.id, 'content', `\n\n✅ 执行成功: ${msg}`);
                        } else {
                            const errMsg = result.error || result.Error || '执行失败';
                            ChatLog.warn?.('PIPELINE', `[code-pipeline] 自动执行失败: ${errMsg}`);
                            sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 执行失败: ${errMsg}`);
                        }
                    } else {
                        ChatLog.warn?.('PIPELINE', '[code-pipeline] excelBridge.ExecuteVBA 不可用，跳过自动执行');
                    }
                } catch (execErr) {
                    ChatLog.error?.('PIPELINE', '[code-pipeline] 自动执行异常:', execErr);
                    sessionStore.appendDelta(turn.id, 'content', `\n\n❌ 执行异常: ${execErr?.message || String(execErr)}`);
                }
            }

            // 写入通用执行上下文（多轮继承）：代码 + 数据快照 + 参数，供后续消息继承
            if (sessionStore.setExecSpec && extracted && !extracted.startsWith('⚠️')) {
                // v20260924：用工厂统一构建 data，消除 code pipeline 精简 selection schema 的漂移
                const fact = sessionStore.execSpecFactory;
                const commonData = fact && typeof fact.buildCommonData === 'function'
                    ? fact.buildCommonData(turn)
                    : {}; // 兜底
                sessionStore.setExecSpec(turn.id, {
                    pipeline: 'code',
                    data: commonData,
                    params: {
                        lang: actualLang || lang,
                        template: usedTemplate || null
                    },
                    result: {
                        success: true,
                        code: extracted || '',
                        message: ''
                    },
                    ts: Date.now()
                });
            }

            sessionStore.completeTurn(turn.id, { reasoning: reasoningText });
            yield { type: 'pipeline_done', pipeline: 'code' };
        } catch (err) {
            ChatLog.error?.('PIPELINE', '[code-pipeline] error', err);
            sessionStore.failTurn(turn.id, { message: err?.message || String(err) });
            yield {
                type: 'error',
                error: { code: 'PIPELINE_ERROR', message: err?.message || String(err), recoverable: false }
            };
        }
    }

    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('code', {
            execute,
            name: '代码生成',
            description: '纯代码生成展示（VBA/Python），不执行，仅复制/保存',
            priority: 100,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.code = { execute };
    }

    // 供渲染器在用户点击 PPT 配置卡按钮时调用
    App.ChatV2.resolvePptConfig = resolvePptConfig;
    ChatLog.info?.('CHAT-V2', '✅ code-pipeline loaded');
})();
