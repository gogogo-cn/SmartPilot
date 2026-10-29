/**
 * ============================================================
 * chat-v2/services/intent-analyzer.js — 意图分析器
 * ============================================================
 * 职责：结合用户输入内容 + 上一轮任务结构，分析后续消息意图并规划下一个任务。
 *   - analyzeFollowUpIntent：路由层意图分析（后续消息 vs 新任务，路由到哪个管道）
 *   - buildFollowUpPlanningPrompt：规划层 prompt 构造（注入任务结构 + 用户新需求 + 规划指令）
 *   - normalizeTaskStructure：统一各管道任务结构格式（tasks/steps/slides）
 * 设计原则：
 *   - 后续消息不是"匹配"出来的，而是"分析"出来的——
 *     结合用户输入内容与上一轮任务结构判断意图，而非纯关键词。
 *   - 纯函数，无副作用，可独立测试。
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 显式新任务标记（与 orchestrator NEW_TASK_PATTERNS 对齐）
    const NEW_TASK_PATTERNS = /新任务|换个任务|换一个任务|另起一个|重新开始|从头开始|新做一个|新建一个|重新做一个|换一个话题/i;

    // 引用上一轮任务的措辞（比纯关键词更宽：含代词/指示词/继续/修改/细化）
    // 【修复-上下文组合】收紧引用词表：去掉单字"这/那/再/换/改/加"、Excel 高频词"区域"，
    // 以及"这个/那个/增加/插入"等常指当前选中区域或新任务操作的词，
    // 避免与上一轮无关的新消息被误判为后续消息、错误继承前文上下文。
    const REFERENCE_PATTERNS = {
        // 指示代词：仅保留明确指向上文的措辞（"这个/那个/这些/那些"在 Excel 中常指当前选中区域）
        pronoun: /上面的|刚才的|上一步|之前的|前面|它们/i,
        // 继续/延伸（去掉裸"再"及口语连接词"然后/接下来"，防止新话题被误判为延续）
        continue: /继续|接着|进一步|往下|延伸|补充|追加|再次|再来/i,
        // 修改/调整（去掉"插入/新增/添加/增加/加个"等更像新增内容且 Excel 高频的词）
        modify: /改成|改为|换成|换个|修改|调整|更新|重新|删除|去掉/i,
        // 分析细化（"区域"在 Excel 语境过于常用，易误伤，移除）
        refine: /换个维度|换个角度|细分|对比|预测|趋势|异常|维度|按.*分组|按.*汇总|筛选|过滤/i,
        // 图表词（针对图表类管道）
        // 【注意】图表词只是"产物类型"，不等于"引用上文"——本项【不】参与下方"通用引用措辞"循环，
        // 由 referencesPreviousTask 单独判定（自带数据维度时不算引用），详见 SELF_CONTAINED_DATA_RE。
        chart: /图表|柱状|饼图|折线|柱形|条形|散点|面积|雷达|圆环|漏斗|瀑布|可视化|画图|做图/i,
        // PPT（针对 PPT 管道）
        ppt: /章节|页面|目录|封面|结尾|要点|主题|配色|颜色|幻灯片|演示/i
    };

    // 跨管道：明确引用上一轮分析并要转 PPT
    // 跨管道：引用上一轮分析/报告/结果 + 明确要 PPT
    const PPT_CROSS_PIPELINE_PATTERNS = /(?:上面|刚才|之前)(?:的)?(?:分析|报告|结果).*(?:PPT|ppt|幻灯片|演示文稿|做成ppt|转成ppt)|(?:把|将)(?:上面|刚才|之前)?(?:的)?(?:分析|报告|结果).*(?:转|做|做成|制作|生成|出).*(?:PPT|ppt|幻灯片|演示文稿)/i;
    // 跨管道：图表意图
    const CHART_INTENT_PATTERNS = /图表|画图|做图|可视化|柱状图|折线图|饼图|组合图|漏斗图|瀑布图|箱线图|树状图|雷达图|生成图表|做图表|画图表|创建图表|同比|导出图片|图表推荐/i;

    // 「自带数据维度/范围」的短语：按<维度>展示/统计/汇总/计算/分组/排名/对比/分析…
    // 命中说明用户在本条消息里已经指明了数据来源与口径 —— 这是一条**独立新命令**，
    // 而不是"在折腾上一轮的产物"。此时图表词（饼图/柱状图…）不能作为"引用上文"的信号。
    // 实测根因（2026-10-10 00:32 日志）：上一轮 python 分析后重发「使用饼图按部门展示销售额」，
    //   被判为后续消息 → targetPipeline 继承上一轮 python → 整段分析重跑、饼图一次没画。
    // 反例（必须仍然继承）：纯换图型「换成柱状图」、无数据线索的「用饼图展示一下」。
    const SELF_CONTAINED_DATA_RE = /按[^，。；、！？\s]{1,14}?(?:展示|统计|汇总|计算|分组|排序|排名|对比|分析|求和|平均|计数)/;

    /**
     * 统一各管道任务结构为通用任务清单。
     * 兼容：python/text 存 ts.tasks；skill 存 ts.steps；code 存 ts.slides。
     * @param {Object} prevSpec 上一轮 _execSpec（或其 spec）
     * @returns {Array<Object>} 通用任务清单 [{ name, detail, type }]
     */
    function normalizeTaskStructure(prevSpec) {
        if (!prevSpec) return [];
        const spec = prevSpec.spec || prevSpec;
        const ts = spec.taskStructure;
        if (!ts) return [];
        let raw = [];
        if (Array.isArray(ts.tasks)) raw = ts.tasks;
        else if (Array.isArray(ts.steps)) raw = ts.steps;
        else if (Array.isArray(ts.slides)) raw = ts.slides;
        return raw.map(function(t) {
            if (typeof t === 'string') return { name: t, detail: '', type: 'text' };
            return {
                name: t.displayName || t.chartLabel || t.title || t.templateId || t.name || '任务',
                detail: t.text || t.detail || t.summary || '',
                type: t.type || t.chartEnum || t.templateId || ''
            };
        }).filter(function(t) { return t.name; });
    }

    /**
     * 提取上一轮数据列（列名 + 数值/日期类型推断），供实体识别/列名映射（融合 Julius 实体识别）。
     * @param {Object} prevSpec 上一轮 _execSpec（或其 spec）
     * @returns {{ headers: string[], numeric: string[], date: string[] }}
     */
    function extractColumns(prevSpec) {
        if (!prevSpec) return { headers: [], numeric: [], date: [] };
        const spec = prevSpec.spec || prevSpec;
        const data = spec.data || {};
        const headers = [];
        const seen = {};
        const addHeaders = function(arr) {
            if (!Array.isArray(arr)) return;
            for (let i = 0; i < arr.length; i++) {
                const key = String(arr[i]);
                if (key && !seen[key]) { seen[key] = true; headers.push(key); }
            }
        };
        if (data.selection && Array.isArray(data.selection.headers)) addHeaders(data.selection.headers);
        if (Array.isArray(data.referenceRegions)) {
            for (let i = 0; i < data.referenceRegions.length; i++) {
                const r = data.referenceRegions[i];
                if (r && Array.isArray(r.headers)) addHeaders(r.headers);
            }
        }
        const numeric = [];
        const date = [];
        const sampleRow = (data.selection && Array.isArray(data.selection.rows) && data.selection.rows[0]) || null;
        for (let i = 0; i < headers.length; i++) {
            const h = headers[i];
            const v = sampleRow ? sampleRow[h] : undefined;
            if (v == null) continue;
            if (typeof v === 'number') numeric.push(h);
            else if (/^\d+(\.\d+)?$/.test(String(v))) numeric.push(h);
            else if (!isNaN(Date.parse(String(v))) && /[-/]/.test(String(v))) date.push(h);
        }
        return { headers, numeric, date };
    }

    /**
     * 构造列名映射 prompt 段（实体识别：用户术语 → 实际列名）。
     * 完整版：列出上一轮实际列名 + 数值/日期类型，并指示模型把用户用词映射到实际列。
     * @param {Object} prevSpec
     * @returns {string} prompt 片段（无列时返回 ''）
     */
    function buildColumnMappingPrompt(prevSpec) {
        const cols = extractColumns(prevSpec);
        if (!cols.headers.length) return '';
        const lines = ['【上一轮数据列】'];
        lines.push('可用列：' + cols.headers.join('、'));
        if (cols.numeric.length) lines.push('数值列：' + cols.numeric.join('、'));
        if (cols.date.length) lines.push('日期列：' + cols.date.join('、'));
        lines.push('当用户消息中的指标/维度用词与列名不完全一致时（如用户说"销售额"而列名为 total_revenue），请先映射到实际列名，再基于该列规划下一个任务。');
        return lines.join('\n');
    }

    /**
     * 判断消息是否引用上一轮任务（结合用户输入 + 任务结构内容）。
     * @param {string} message
     * @param {Object} prevSpec
     * @returns {boolean}
     */
    function referencesPreviousTask(message, prevSpec) {
        const msg = String(message || '');
        if (!msg) return false;
        // 1. 引用措辞（代词/继续/修改/PPT）：明确的引用信号
        for (const key of Object.keys(REFERENCE_PATTERNS)) {
            if (key === 'refine') continue; // 细化词单独走管道门槛（见下方）
            if (key === 'chart') continue;  // 图表词单独判定（见下方 1.5）
            if (REFERENCE_PATTERNS[key].test(msg)) return true;
        }
        // 1.5 图表词：只有"没自带数据维度/范围"时才算引用上一轮。
        //     自带维度的图表命令（"使用饼图按部门展示销售额"）是独立新任务，不继承上一轮管道。
        if (REFERENCE_PATTERNS.chart.test(msg) && !SELF_CONTAINED_DATA_RE.test(msg)) return true;
        // 2. 分析细化词：仅当上一轮是纯分析管道（python/text）时才算引用，
        //    skill 管道多为表格操作类（美化/清洗/CRUD），"对比/趋势"等词更可能是无关新任务
        const prevPipeline = prevSpec && prevSpec.spec ? prevSpec.spec.pipeline : null;
        const isAnalysisPipeline = prevPipeline === 'python' || prevPipeline === 'text';
        if (isAnalysisPipeline && REFERENCE_PATTERNS.refine.test(msg)) return true;
        // 3. 任务结构内容引用：消息包含任务名中的实体词（如"华东"）
        const tasks = normalizeTaskStructure(prevSpec);
        for (const t of tasks) {
            const name = String(t.name || '');
            const entities = name.match(/[\u4e00-\u9fa5]{2,}/g) || [];
            for (const e of entities) {
                if (msg.includes(e)) return true;
            }
        }
        return false;
    }

    /**
     * 路由层意图分析：结合用户输入 + 上一轮任务结构，判断是否为后续消息及路由目标。
     * @param {Object} params
     * @param {string} params.message 当前用户消息
     * @param {Object|null} params.prevSpec 上一轮 _execSpec（含 spec 字段）
     * @param {Object} [params.context] 当前上下文（含选区）
     * @returns {{
     *   isFollowUp: boolean,
     *   followUpType: 'continue'|'modify'|'refine'|'cross_ppt'|'cross_chart'|'new_task'|null,
     *   targetPipeline: string|null,
     *   reason: string,
     *   confidence: 'high'|'medium'|'low'
     * }}
     */
    function analyzeFollowUpIntent({ message, prevSpec, context }) {
        const msg = String(message || '');
        if (!prevSpec || !prevSpec.spec) {
            return { isFollowUp: false, followUpType: null, targetPipeline: null, reason: '无上一轮任务结构', confidence: 'low' };
        }
        const prev = prevSpec.spec;
        const prevPipeline = prev.pipeline;

        // 1. 明确新任务 → 不继承
        if (NEW_TASK_PATTERNS.test(msg)) {
            return { isFollowUp: false, followUpType: 'new_task', targetPipeline: null, reason: '明确新任务', confidence: 'high' };
        }

        // 2. 跨管道 PPT：明确引用上一轮分析并要转 PPT → code-pipeline
        const pptIntent = /ppt|powerpoint|幻灯片|演示文稿|做ppt|做成ppt|生成ppt|汇报/i.test(msg);
        if (pptIntent && PPT_CROSS_PIPELINE_PATTERNS.test(msg)) {
            if (prevPipeline === 'skill' || prevPipeline === 'text' || prevPipeline === 'python') {
                return { isFollowUp: true, followUpType: 'cross_ppt', targetPipeline: 'code', reason: '引用上一轮分析转 PPT', confidence: 'high' };
            }
        }

        // 3. 跨管道图表：图表意图 + （无新选区 或 明确仅切图表类型） + 上一轮分析 → skill-pipeline
        if (CHART_INTENT_PATTERNS.test(msg)) {
            const ctxHasData =
                (Array.isArray(context && context.selection && context.selection.headers) && context.selection.headers.length > 0) ||
                (Array.isArray(context && context.referenceRegions) &&
                    context.referenceRegions.some(function(r) { return Array.isArray(r.headers) && r.headers.length > 0; }));
            // 明确"仅切图表类型"（换成柱状图/改成折线图）即使带选区也必须继承上一轮数据，
            // 否则回退到实时选区画出错误内容（实测第二条消息内容不对）。
            const rc = App.ChatV2.routingConstants;
            const chartTypeSwitch = !!(rc && rc.CHART_TYPE_SWITCH_RE && rc.CHART_TYPE_SWITCH_RE.test(msg));
            // 自带数据维度的图表命令（"使用饼图按部门展示销售额"）本就不该继承上一轮数据，
            // 无论当前有无选区 —— 否则上一轮 python 分析会被复活重跑（实测 2026-10-10 00:32）。
            const selfContained = SELF_CONTAINED_DATA_RE.test(msg);
            if (!selfContained && (!ctxHasData || chartTypeSwitch) && (prevPipeline === 'skill' || prevPipeline === 'text' || prevPipeline === 'python')) {
                return { isFollowUp: true, followUpType: 'cross_chart', targetPipeline: 'skill', reason: '引用上一轮分析生成图表', confidence: 'high' };
            }
        }

        // ✨ 4. 确定性 Skill 继承：上一轮是 Beautifier/Formatter/Chart/Cleaner/CRUD，
        //    本轮消息包含同类领域关键词（主题/美化/格式/图表/清洗/增删改查）→ 强制路由回 skill-pipeline。
        //    解决：用户说"换个紫色主题"这类没有"改成/换成"修改词但意图明显是继续美化的请求，
        //    referencesPreviousTask 因没命中代词/修改词而返回 false，导致后续消息掉进 text-pipeline。
        if (prevPipeline === 'skill') {
            const prevSkillName = prev.params?.skillName || '';
            // 技能名是否属于确定性 Skill 域：统一走声明式注册表（config/skill-domain.js）
            const sd = App.ChatV2?.skillDomain;
            const isDeterministicSkill = !!(sd?.isDomainName) && sd.isDomainName(prevSkillName);
            if (isDeterministicSkill) {
                // 确定性 Skill 域消息命中判定统一走 skill-domain.isDomainMessage（由各 descriptor 关键词派生），
                // 不再本地私写 domainKw 正则副本（避免"主题|美化|图表|清洗…"关键词两处漂移）。
                const domainKwHit = !!(sd?.isDomainMessage) && sd.isDomainMessage(msg);
                if (domainKwHit) {
                    return { isFollowUp: true, followUpType: 'modify', targetPipeline: 'skill', reason: '确定性 Skill 领域关键词匹配（' + prevSkillName + ' → ' + prevSkillName + '）', confidence: 'medium' };
                }
            }
        }

        // 5. 引用上一轮任务 → 路由回原管道
        if (referencesPreviousTask(msg, prevSpec)) {
            let followUpType = 'continue';
            if (REFERENCE_PATTERNS.modify.test(msg)) followUpType = 'modify';
            else if (REFERENCE_PATTERNS.refine.test(msg)) followUpType = 'refine';
            return {
                isFollowUp: true,
                followUpType: followUpType,
                targetPipeline: prevPipeline,
                reason: '引用上一轮任务结构（' + followUpType + '）',
                confidence: 'medium'
            };
        }

        return { isFollowUp: false, followUpType: null, targetPipeline: null, reason: '未引用上一轮任务', confidence: 'low' };
    }

    /**
     * 规划层 prompt 构造：注入上一轮任务结构 + 用户新需求 + 意图分析 + 规划指令，
     * 明确要求 LLM 基于任务结构规划下一个任务。
     * @param {Object} params
     * @param {Object} params.prevSpec 上一轮 _execSpec（或其 spec）
     * @param {string} params.userMessage 当前用户消息
     * @param {Object} [params.intent] 路由层分析结果
     * @returns {string|null} prompt 片段
     */
    function buildFollowUpPlanningPrompt({ prevSpec, userMessage, intent }) {
        if (!prevSpec) return null;
        const spec = prevSpec.spec || prevSpec;
        const parts = [];
        const res = (spec.result) || {};

        // 1. 上一轮分析结论
        if (res.summary) {
            parts.push('【上一轮分析结论】\n' + res.summary);
        }
        if (Array.isArray(res.steps) && res.steps.length) {
            const lines = res.steps.map(function(s) {
                return '- ' + (s.displayName || s.templateId || '分析') + (s.summary ? '：' + s.summary : '');
            });
            parts.push('【上一轮分析内容】\n' + lines.join('\n'));
        }

        // 2. 上一轮任务结构（任务清单/规划步骤）
        const tasks = normalizeTaskStructure(spec);
        if (tasks.length) {
            const planLines = tasks.map(function(t, i) {
                return (i + 1) + '. ' + t.name + (t.detail ? '：' + t.detail : '');
            });
            parts.push('【上一轮任务规划】\n' + planLines.join('\n'));
        }

        // 3. 上一轮数据列（实体识别/列名映射，融合 Julius 实体识别）
        const colPrompt = buildColumnMappingPrompt(prevSpec);
        if (colPrompt) parts.push(colPrompt);

        // 4. 用户新需求 + 意图分析 + 规划指令
        const intentType = (intent && intent.followUpType) || 'continue';
        const intentDesc = {
            continue: '继续/延伸上一轮任务',
            modify: '修改/调整上一轮任务',
            refine: '对上一轮任务进一步细化/细分/对比',
            cross_ppt: '把上一轮分析结果转为 PPT',
            cross_chart: '把上一轮分析结果转为图表'
        }[intentType] || '继续上一轮任务';
        parts.push('【用户新需求】\n' + String(userMessage || ''));
        parts.push('【意图分析】\n用户意图：' + intentDesc + '。请结合上一轮任务规划与用户新需求，规划下一个任务：' +
            '明确要做什么（任务名）、基于哪个数据/维度、产出什么结果。' +
            '若新需求是对上一轮结果的进一步细分、筛选、对比或修改，请在规划中体现与上一轮任务结构的对应关系。' +
            '若用户消息含代词或省略表达（如"再对比一下"、"它"、"这个"、"上面"、"刚才的"），' +
            '请基于上一轮分析结论与任务规划推断其指代对象，再规划下一个任务。');

        return parts.join('\n\n');
    }

    /**
     * 提取当前上下文的列信息（供 LLM 意图分析参考）。
     * @param {Object} context
     * @returns {{headers: string[], numeric: string[], date: string[]}}
     */
    function extractContextColumns(context) {
        const headers = [];
        const seen = {};
        const add = function (arr) {
            if (!Array.isArray(arr)) return;
            for (let i = 0; i < arr.length; i++) {
                const k = String(arr[i]);
                if (k && !seen[k]) { seen[k] = true; headers.push(k); }
            }
        };
        if (context && Array.isArray(context.selection && context.selection.headers)) add(context.selection.headers);
        if (context && Array.isArray(context.referenceRegions)) {
            for (let i = 0; i < context.referenceRegions.length; i++) {
                const r = context.referenceRegions[i];
                if (r && Array.isArray(r.headers)) add(r.headers);
            }
        }
        const numeric = [];
        const date = [];
        const sampleRow = (context && context.selection && Array.isArray(context.selection.rows) && context.selection.rows[0]) || null;
        for (let i = 0; i < headers.length; i++) {
            const h = headers[i];
            const v = sampleRow ? sampleRow[h] : undefined;
            if (v == null) continue;
            if (typeof v === 'number') numeric.push(h);
            else if (/^\d+(\.\d+)?$/.test(String(v))) numeric.push(h);
            else if (!isNaN(Date.parse(String(v))) && /[-/]/.test(String(v))) date.push(h);
        }
        return { headers, numeric, date };
    }

    function normSkillKey(s) {
        return String(s || '').toLowerCase().replace(/[\s_\-]/g, '');
    }

    /**
     * 模型给出的 skillId 必须能在本次发给它的目录里对上（id 或 name，忽略大小写与 _/-）。
     * 对不上就改成 free_chat，避免把自创名字（如 SmartExcelFormatter）当成技能路由。
     * catalog 为空或未就绪时原样返回，避免目录还没加载就把合法路由收掉。
     */
    function coerceRouteToCatalog(route, catalog) {
        if (!route || typeof route !== 'object') return route;
        if (!Array.isArray(catalog) || catalog.length === 0) return route;
        const sid = String(route.skillId || '').trim();
        const key = normSkillKey(sid);
        const asText = function (extra) {
            const out = {
                pipeline: 'text',
                skillId: 'free_chat',
                confidence: extra && extra.droppedSkillId ? 0 : (Number(route.confidence) || 0),
                reason: String(route.reason || '')
            };
            if (extra && extra.droppedSkillId) out.droppedSkillId = extra.droppedSkillId;
            return out;
        };
        if (!key || key === 'freechat') return asText();
        const hit = catalog.find(function (s) {
            if (!s) return false;
            return normSkillKey(s.id) === key || normSkillKey(s.name) === key;
        });
        if (!hit || normSkillKey(hit.id) === 'freechat') return asText(hit ? null : { droppedSkillId: sid });
        return {
            pipeline: String(route.pipeline || 'text'),
            skillId: String(hit.id || sid),
            confidence: Number(route.confidence) || 0,
            reason: String(route.reason || '')
        };
    }

    /**
     * LLM 意图分析（新请求前置门）：语义理解"要做什么/对什么数据/产出什么"。
     * 产出结构化意图驱动路由与执行；关键词只做快速通道与 LLM 失败时的兜底。
     * @param {Object} params
     * @param {string} params.message 用户请求
     * @param {Object} [params.context] 当前上下文（含选区/参考区）
     * @param {AbortSignal} [params.signal]
     * @param {Function} [params.onReasoning] 推理流回调 (delta)=>void（透传给 streamText）
     * @returns {Promise<Object|null>} {
     *   needsPlanning: boolean,
     *   operations: string[],
     *   targetData: { sheet: string, columns: string[] },
     *   output: { type: string, name: string },
     *   summary: string,
     *   reason: string
     * } 失败返回 null（调用方回退关键词路由）
     */
    async function analyzeNewIntent({ message, context, signal, onReasoning }) {
        const msg = String(message || '').trim();
        if (!msg) return null;
        const llmIo = App.ChatV2.llmIo;
        if (!llmIo || typeof llmIo.askJson !== 'function') return null;

        const host = (window.SmartPilotHost || (App && App.SmartPilotHost));
        const isPpt = !!(host && host.isPowerPoint);
        const isWord = !!(host && host.isWord);

        const cols = extractContextColumns(context);
        const ctxLines = [];
        if (cols.headers.length) ctxLines.push('当前数据列：' + cols.headers.join('、'));
        if (cols.numeric.length) ctxLines.push('数值列：' + cols.numeric.join('、'));
        if (cols.date.length) ctxLines.push('日期列：' + cols.date.join('、'));
        // 选中对象（PPT 表格/形状、Word 表格）：这是“无选区”误判的根因——必须显式告知 LLM，
        // 否则它只会看到“（无选区数据）”而追问幻灯片/行列号。
        if (context && context.selectedObject && context.selectedObject.summary) {
            ctxLines.push('已选中对象：' + context.selectedObject.summary);
        }
        if (!ctxLines.length) ctxLines.push('（无选区数据）');

        // LLM 路由用的 skill 目录（优先动态全量：内置领域 + 自定义/扩展 skill；未就绪时回退静态目录）
        let skillCatalogText = '';
        // 目录里真实可用的 id 集合（null = 目录取不到，此时不擅自收窄规则，保持旧行为）
        let availableSkillIds = null;
        // 完整目录条目（id + name）。路由落盘前用来丢掉模型自创的技能名。
        let skillCatalog = null;
        try {
            const sd = App.ChatV2?.skillDomain;
            if (sd) {
                let cat = null;
                if (typeof sd.getDynamicSkillCatalog === 'function') {
                    const cached = (Array.isArray(App.cachedSkills) && App.cachedSkills.length > 0)
                        ? App.cachedSkills
                        : null;
                    cat = sd.getDynamicSkillCatalog(cached);
                } else if (typeof sd.getSkillCatalog === 'function') {
                    cat = sd.getSkillCatalog();
                }
                if (Array.isArray(cat) && cat.length > 0) {
                    skillCatalog = cat;
                    skillCatalogText = '\n【可用 Skill 目录（从下列选最匹配的 route.skillId；skillId 可以是领域 id 或目录中的具体技能名）】\n' +
                        cat.map(s => `- ${s.id}：${s.description}${s.keywords ? '（关键词: ' + s.keywords + '）' : ''}`).join('\n');
                    availableSkillIds = new Set(cat.map(s => String((s && s.id) || '').toLowerCase()).filter(Boolean));
                }
            }
        } catch (_e) { /* skillDomain 未就绪时跳过 */ }

        // 领域 id → 给 LLM 举例的典型说法（只用于下方路由规则行）。
        const SKILL_RULE_EXAMPLES = {
            dashboard: '"做看板/做仪表盘/做驾驶舱/大屏"（即使用户说"什么是看板"这种纯问答，也选 free_chat + text）',
            chart: '"生成图表/做个柱状图/画饼图"',
            beautifier: '"美化/主题/样式/边框"（仅当用户想改工作表外观；"把结果用表格展示/显示/输出结果"是渲染展示，不是美化，不要选它）',
            cleaner: '"去重/清洗/删空行/合并工作表/合并数据/合到一张表/拆分列/按部门拆分数据到不同Sheet/拆分工作簿"',
            formatter: '"格式化/日期格式/数字格式"',
            crud: '"增删改查/读取/写入/插入/删除"',
            html_report: '"生成 HTML/网页报表/web 大屏"（用户要的是浏览器/网页报表；"把结果用表格/json展示"是聊天内展示，不要选它）',
            // 【防抢路由】"创建一个解析器/解析脚本"不是"创建 skill"：
            // 原先这条只写"创建 skill/智能体/插件"，LLM 看到"创建"就照抄 skill_creator，
            // 实测"为火车票创建一个解析器"被路由到 SkillCreator（用户："这都能选错技能？"）。
            // 规则文本必须把边界写死：只认明确出现 skill/技能/智能体/插件的请求。
            skill_creator: '"创建 skill/技能/智能体/插件"（消息里必须真的出现 skill/技能/智能体/插件 字样；' +
                '"创建一个解析器/解析脚本/提取脚本"不算 → 那是 GenerateOCRParseScript）',
            GenerateOCRParseScript: '"创建/生成 解析器、解析脚本、OCR 字段提取脚本"（如"为火车票创建一个解析器"' +
                '"按这个 Excel 表头做个解析脚本"）'
        };
        // 【停用感知·必做】规则行只对"目录里真实存在"的领域举例。
        // 此前这里是硬编码全量列表：用户停用 beautifier 后，上面注入的【可用 Skill 目录】
        // 里确实已经没有它了（DESCRIPTORS 空壳也不再复活），但这 8 行规则照旧写着
        // "美化/主题/样式/边框 → skillId=\"beautifier\""，LLM 严格照规则作答 →
        // 意图分析块显示"路由：skill → beautifier（置信度 1）"。
        // 实测反馈："还知道 skill 的存在，匹配了，不过应该是没使用"——执行层没受影响
        // （守卫 C 丢弃空壳，实际走了通用工具），但提示信息在骗人，用户以为停用无效。
        // 目录取不到时不加过滤（availableSkillIds=null），避免把本来能用的路由也收窄掉。
        // 【大小写不敏感】目录里的动态技能 id 就是技能名（如 GenerateOCRParseScript），
        // 而 availableSkillIds 统一小写化；用大小写敏感比较会让这类驼峰 key 的规则整条被丢掉
        // （规则写了却不生效，最难查）。
        const skillRuleText = Object.keys(SKILL_RULE_EXAMPLES)
            .filter(id => !availableSkillIds || availableSkillIds.has(String(id).toLowerCase()))
            .map(id => `- ${SKILL_RULE_EXAMPLES[id]} → skillId="${id}"`)
            .join('\n');

        const prompt = `你是一个 ${(isPpt ? 'PowerPoint 演示文稿' : (isWord ? 'Word 文档' : 'Excel 工作簿'))} 操作意图分析器 + 路由器。请理解用户请求的语义，判断需要执行哪些操作、针对什么数据、产出什么结果，以及最适合交给哪个 Skill 处理。

【用户请求】
${msg}

【当前上下文】
${ctxLines.join('\n')}
${skillCatalogText}

请输出 JSON（不要任何其他内容）：
{
  "needsPlanning": true 或 false,
  "operations": ["read","analyze","write",...],
  "targetData": {"sheet": "${(isPpt ? '幻灯片/形状定位描述' : (isWord ? '文档位置描述' : '工作表名'))}或空串", "columns": ["列名"]},
  "inferredParams": {"dimColumn": "", "valueColumn": "", "dateColumn": "", "aggColumn": "", "groupBy": "", "aggFunc": "sum", "chartType": ""},
  "output": {"type": "new_sheet|chart|text|same_sheet|html", "name": "产出名称或空串"},
  "route": {
    "pipeline": "skill|text",
    "skillId": "从上面 Skill 目录中选最匹配的 id（领域 id 或具体技能名；纯问答/复合任务用 free_chat）",
    "confidence": 0.0~1.0,
    "reason": "为什么选这个 skill"
  },
  "summary": "一句话总结用户意图",
  "reason": "判断理由"
}

路由规则（最重要）：
- route.pipeline="skill" 时必须从上面的 Skill 目录选一个 skillId（不能是 free_chat）
- 🔒 route.skillId 必须从【可用 Skill 目录】逐字复制现成的 id（如 "SmartDataCleaner"）。
  严禁缩写、翻译、改写或自创 id——"data_cleaning"、"data_analyze"、"cleaner_skill" 等都不存在，
  写了也会被系统拒绝执行。目录里确实没有匹配的技能 → 直接 skillId="free_chat", pipeline="text"
- route.pipeline="text" 时 skillId 填 "free_chat"
${skillRuleText}
- 纯问答（"什么是 X / X 原理 / X 怎么用"） → skillId="free_chat", pipeline="text"
- 复合多步骤（"先读数据然后分析再画图"涉及多个 skill 领域） → skillId="free_chat", pipeline="text"
- 检查/检测工作表或数据是否包含错误（错误值、#DIV/0!、#N/A、#VALUE!、公式错误） → skillId="free_chat", pipeline="text"（由 agent 调用 excel_check_errors 工具精确定位错误单元格；不要路由到数据分析类 skill 做泛化异常值分析）
- 🔒 合并/分拆/拆分工作表类请求（"合并工作表"、"合到一张表"、"合并所有sheet"、"按字段拆分数据到不同Sheet"、"拆分工作簿"、"分拆数据"） → 目录里有 cleaner 时 route.pipeline="skill", route.skillId="cleaner"；目录里没有就 pipeline="text", skillId="free_chat"，不要自造 formatting / formatter。**operations 直接写原子操作名（驼峰，注意大小写）**：合并多工作表→["mergeSheets"]，按列拆分→["splitByColumn"]，拆分工作簿→["splitWorkbook"]，按行数拆分→["splitByRows"]，合并多工作簿→["mergeWorkbooks"]。**严禁脑补**成 read+write+copy+create_sheet 这类通用操作串，也严禁只选 analyze/read 然后让 agent 自己读数据做文本分析声称"已合并"——必须调 excel_merge_sheets / excel_split_by_column 等原子工具产出实体表。
- 同一句里还有第二项（设置成中国红、美化、筛选、图表）时，summary 和 operations 必须两项都写上，needsPlanning=true。禁止只写成「格式化为红色」而丢掉合并，禁止把「中国红」改成笼统的「红色」。

operations 取值（驼峰命名，优先使用原子名而非通用串）：
read(读取)/write(写入)/analyze(分析)/aggregate(汇总)/chart(图表)/format(格式化)/copy(复制)/delete(删除)/create_sheet(建表)/send(发送)/other
mergeSheets(合并多工作表)/mergeWorkbooks(合并多工作簿)/splitByColumn(按字段拆表)/splitWorkbook(拆分工作簿为独立文件)/splitByRows(按行数拆表)
涉及多个独立操作（如"先清洗数据再画柱状图"）→ needsPlanning=true
**单个结构变换原子操作（合并/拆分/清洗/图表）→ needsPlanning=false**（即使内部涉及读写复制，这是一个原子工具一步到位，不是多步骤编排）
简单问答/问候/闲聊 → needsPlanning=false 且 route.pipeline="text"
用户说"分析/看看/研究/评估/对比/找出"某数据 → operations 含 analyze，且通常 needsPlanning=true（需读取+计算+给结论）
用户说"写到/保存到/放到/复制到新表" → operations 含 write 或 copy，output.type="new_sheet"，output.name 填表名
targetData.columns：必须列出【当前上下文】中所有可用的数据列名（供下游做列名约束）
inferredParams：根据用户意图和当前数据列，推断最合适的工具参数：
  * 如果用户请求涉及汇总/分组（如"分析哪些产品卖得好"、"按部门汇总"）：推断 dimColumn（分组列）和 valueColumn（数值列），必须从 targetData.columns 中选择
  * 如果涉及日期分组（如"按月汇总"）：推断 dateColumn（日期列）和 groupBy（year/quarter/month/week/day）
  * 如果涉及图表：推断 chartType（如 xlColumnClustered/xlPie/xlLine 等）
  * aggFunc 默认为 sum，如果用户明确说"平均/计数"则填 average/count
  * 如果无法推断或用户未指定，保持空字符串 ""
- 用户明确指定${(isPpt ? '幻灯片或形状（如"第3张幻灯片里的表格"）' : (isWord ? '文档位置（如"第二段/这个表格"）' : '工作表（如"读取 Sheet1"）'))} → targetData.sheet 填定位描述；未指定则空串`;

        try {
            const res = await llmIo.askJson({
                messages: [{ role: 'user', content: prompt }],
                signal,
                config: App.config,
                onReasoning: (typeof onReasoning === 'function') ? onReasoning : undefined
            });
            const intent = res && res.json;
            if (!intent || typeof intent !== 'object') {
                ChatLog.warn?.('INTENT', '[analyzeNewIntent] LLM 未返回有效 JSON（回退关键词路由）: raw=' + String((res && res.raw) || '').slice(0, 120));
                return null;
            }
            // 【关键】targetData.columns 必须包含所有实际可用列名（供 TaskPlanner 做列名约束），
            // 而不是仅用户提到的列。如果 LLM 返回的列不全，用 extractContextColumns 的结果兜底。
            const llmColumns = (intent.targetData && Array.isArray(intent.targetData.columns))
                ? intent.targetData.columns.map(String)
                : [];
            const allColumns = cols.headers.length > 0 ? cols.headers : llmColumns;
            // 【场景推理】提取 LLM 推断的参数，供 TaskPlanner 直接使用
            const inferred = (intent.inferredParams && typeof intent.inferredParams === 'object')
                ? intent.inferredParams
                : {};
            const inferredParams = {
                dimColumn: String(inferred.dimColumn || ''),
                valueColumn: String(inferred.valueColumn || ''),
                dateColumn: String(inferred.dateColumn || ''),
                aggColumn: String(inferred.aggColumn || ''),
                groupBy: String(inferred.groupBy || ''),
                aggFunc: String(inferred.aggFunc || 'sum'),
                chartType: String(inferred.chartType || '')
            };
            // 校验推断的列名是否在实际列中存在，不存在则清空
            if (allColumns.length > 0) {
                ['dimColumn', 'valueColumn', 'dateColumn', 'aggColumn'].forEach(field => {
                    const val = inferredParams[field];
                    if (val && !allColumns.some(c => c.toLowerCase() === val.toLowerCase())) {
                        ChatLog.warn?.('INTENT', `[analyzeNewIntent] 推断列名 ${field}="${val}" 不在实际列中，清空`);
                        inferredParams[field] = '';
                    }
                });
            }

            // route：LLM 直接给路由决策——是走 skill-pipeline 还是 text-pipeline，
            // 以及选哪个 skillId。这比前端堆 regex 精准得多，LLM 的语义理解天然能
            // 区分"做个看板"→ dashboard 和"什么是看板"→ free_chat。
            let llmRoute = (intent.route && typeof intent.route === 'object')
                ? {
                    pipeline: String(intent.route.pipeline || 'text'),
                    skillId: String(intent.route.skillId || 'free_chat'),
                    confidence: Math.max(0, Math.min(1, Number(intent.route.confidence) || 0)),
                    reason: String(intent.route.reason || '')
                }
                : null;
            // 目录已就绪时，模型自创的技能名不能进入路由。实测会把没装过的
            // SmartExcelFormatter 写成置信度 1，界面仍把它展示成一条技能路由。
            if (llmRoute && Array.isArray(skillCatalog) && skillCatalog.length) {
                const coerced = coerceRouteToCatalog(llmRoute, skillCatalog);
                if (coerced && coerced.droppedSkillId) {
                    ChatLog.warn?.('INTENT', '[analyzeNewIntent] 丢弃目录外技能「' + coerced.droppedSkillId + '」→ free_chat');
                }
                llmRoute = coerced;
            }

            return completeIntent({
                needsPlanning: !!intent.needsPlanning,
                operations: Array.isArray(intent.operations) ? intent.operations.map(String) : [],
                targetData: (intent.targetData && typeof intent.targetData === 'object')
                    ? {
                        sheet: String(intent.targetData.sheet || ''),
                        columns: allColumns
                    }
                    : { sheet: '', columns: allColumns },
                inferredParams,
                output: (intent.output && typeof intent.output === 'object')
                    ? { type: String(intent.output.type || 'text'), name: String(intent.output.name || '') }
                    : { type: 'text', name: '' },
                route: llmRoute,
                summary: String(intent.summary || ''),
                reason: String(intent.reason || '')
            }, msg);
        } catch (e) {
            ChatLog.warn?.('INTENT', '[analyzeNewIntent] LLM 意图分析失败（回退关键词路由）:', e?.message);
            return null;
        }
    }

    /**
     * 模型常把「合并，并设置成中国红」收成「格式化为红色」。
     * 原话里能对上的每一项都要留在 summary / operations 里，卡片不能只结一项。
     */
    function completeIntent(intent, message) {
        if (!intent || typeof intent !== 'object') return intent;
        const tg = (typeof App !== 'undefined' && App.ChatV2 && App.ChatV2.turnGoal) || null;
        if (!tg || typeof tg.build !== 'function') return intent;
        const goal = tg.build(message || '');
        const kinds = [];
        const seen = {};
        (goal.doneWhen || []).forEach(function (d) {
            if (!d || d.kind === 'reply' || seen[d.kind]) return;
            seen[d.kind] = true;
            kinds.push(d.kind);
        });
        const style = String(goal.styleTarget || '');
        const summary = String(intent.summary || '');
        if (kinds.length < 2) {
            if (style && summary.indexOf(style) < 0 && /红色/.test(summary) && /中国红/.test(style)) {
                intent.summary = summary.replace(/红色/g, style);
                intent.summaryCorrected = true;
            }
            return intent;
        }
        const OP = {
            merge: 'mergeSheets', split: 'splitByColumn', format: 'format', chart: 'chart',
            filter: 'filter', sort: 'sort', aggregate: 'aggregate', clean: 'clean'
        };
        const LABEL = {
            merge: '合并工作表', split: '拆分', format: '设置格式', chart: '图表',
            filter: '筛选', sort: '排序', aggregate: '汇总', clean: '清洗'
        };
        const ops = Array.isArray(intent.operations) ? intent.operations.map(String) : [];
        const keyOf = function (s) { return String(s || '').toLowerCase().replace(/[\s_-]/g, ''); };
        const hasOp = function (kind) {
            const want = keyOf(OP[kind] || kind);
            return ops.some(function (o) {
                const k = keyOf(o);
                if (k === want) return true;
                if (kind === 'merge' && k.indexOf('merge') >= 0) return true;
                if (kind === 'split' && k.indexOf('split') >= 0) return true;
                if (kind === 'format' && k.indexOf('format') >= 0) return true;
                return false;
            });
        };
        const bits = kinds.map(function (k) {
            if (k === 'format' && style) return '设置成' + style;
            return LABEL[k] || k;
        });
        const dropsMerge = kinds.indexOf('merge') >= 0 && !/合并/.test(summary);
        const dropsStyle = !!(style && summary.indexOf(style) < 0);
        const missing = kinds.filter(function (k) { return OP[k] && !hasOp(k); });
        if (!dropsMerge && !dropsStyle && !missing.length) return intent;
        kinds.forEach(function (k) {
            if (OP[k] && !hasOp(k)) ops.push(OP[k]);
        });
        intent.operations = ops;
        intent.operationLabels = bits;
        intent.needsPlanning = true;
        intent.summary = bits.join('，并');
        intent.reason = '原话有多项：' + bits.join('、') + '。不能只写成其中一项。';
        intent.summaryCorrected = true;
        return intent;
    }

    App.ChatV2.intentAnalyzer = {
        NEW_TASK_PATTERNS,
        REFERENCE_PATTERNS,
        PPT_CROSS_PIPELINE_PATTERNS,
        CHART_INTENT_PATTERNS,
        // 自带数据维度/范围的短语（"按部门展示…"）→ 图表词不算引用上文，供测试与路由器复用
        SELF_CONTAINED_DATA_RE,
        normalizeTaskStructure,
        extractColumns,
        buildColumnMappingPrompt,
        referencesPreviousTask,
        analyzeFollowUpIntent,
        buildFollowUpPlanningPrompt,
        analyzeNewIntent,
        coerceRouteToCatalog,
        completeIntent
    };

    ChatLog.info?.('CHAT-V2', '✅ intentAnalyzer loaded');
})();