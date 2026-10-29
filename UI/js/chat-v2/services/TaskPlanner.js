/**
 * ============================================================
 * chat-v2/services/TaskPlanner.js — 多步骤任务规划器
 * ============================================================
 * 职责：将用户请求分解为结构化执行步骤。
 *   与 services/task-planner.js 不同，此文件专为多步骤管道设计，
 *   处理的是"通用对话"场景下的多步骤分解，而非图表/聚合任务拆分。
 *
 * 设计：
 *   - needPlanning() 三层漏斗：快速通道→工具计数→LLM 二分类
 *   - plan() 通过 LLM 将用户请求分解为结构化步骤数组
 *   - 每个步骤包含：strategy(direct/agent/pipeline) + toolName + params + description
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    // 关键词唯一事实源：App.ChatV2.intentConfig（Config/intent_config.json）
    function defs() {
        const ic = App.ChatV2.intentConfig;
        return (ic && typeof ic.defs === 'function') ? ic.defs() : {};
    }

    /**
     * 轻量级判断：此用户请求是否需要多步骤规划
     * 三层漏斗，避免不必要的 LLM 调用
     * @param {string} userMessage
     * @returns {boolean} true=需要多步骤规划
     */
    function needPlanning(userMessage) {
        if (!userMessage || typeof userMessage !== 'string') return false;

        const text = userMessage.trim();
        const d = defs();

        // 第一层：快速通道 — 问候语/简单问答/超短输入直接 false
        if (d.quickPassRe && d.quickPassRe.test(text)) return false;
        if (text.length <= 3) return false;  // 超短输入（≤3字符）跳过规划

        // 第二层：工具计数 — 统计涉及的不同工具类别
        let categoryCount = 0;
        for (const re of Object.values(d.toolKeywords || {})) {
            if (re && re.test(text)) categoryCount++;
        }
        // 2 个以上不同工具类别 → 需要多步骤
        if (categoryCount >= 2) return true;

        // 【图表 + 分类维度 = 必须先聚合再建图，天然两步】
        // 如"使用饼图按部门展示销售额""各部门销量占比图"：
        // 明细表直接建图会画出错误的图（整表范围/多系列饼图），必须先 excel_aggregate 汇总。
        // 只有"把 A1:B6 做成饼图"这类已给出明确两列范围的请求才算单步。
        const hasChartWord = d.toolKeywords?.chart && d.toolKeywords.chart.test(text);
        const hasGroupWord = /(按(?!照|时|钮)|各(部门|地区|区域|类|项|产品|渠道|城市|省份|月|季度|年)|分组|分类|汇总|统计|占比|分布|排名|每个)/.test(text);
        if (hasChartWord && hasGroupWord) return true;

        // 单步工具关键词 → 不需要多步骤
        for (const re of d.singleStepKeywords || []) {
            if (re.test(text)) return false;
        }

        // 第三层：LLM 判断（返回 "need_planning" 标记，由调用方决定是否执行 LLM 判断）
        // 此处返回 true 表示"需要进一步 LLM 判断"
        // 调用方（text-pipeline）应使用 needPlanningLLM() 做最终决定
        return 'needs_llm_check';
    }

    /**
     * LLM 二分类判断：此请求是否需要多步骤规划
     * @param {string} userMessage
     * @param {Function} askLLM - async (prompt) => string
     * @returns {Promise<boolean>}
     */
    async function needPlanningLLM(userMessage, askLLM) {
        if (!userMessage || !askLLM) return false;
        try {
            const prompt = `判断以下用户请求是否需要拆分为"多个步骤"来执行。\n\n` +
                `规则：\n` +
                `- 如果用户请求涉及多个操作（如"先读取数据，再分析，然后生成图表"或"美化表格后再写入新表"）→ YES\n` +
                `- 如果用户请求是单个操作（如"读取A1的数据"、"把B2改成100"、"把A1:B6做成饼图"）→ NO\n` +
                `- 如果用户要求"按某维度展示/统计某个指标"（如"按部门展示销售额"、"各部门销量占比"），\n` +
                `  需要先分类汇总再出图，属于多个操作 → YES\n` +
                `- 如果用户请求是简单的问答、问候、闲聊 → NO\n\n` +
                `用户请求：${userMessage.slice(0, 200)}\n\n` +
                `只回答 YES 或 NO，不要任何其他内容。`;

            const result = await askLLM(prompt);
            const answer = (result || '').trim().toUpperCase();
            return answer === 'YES';
        } catch (e) {
            ChatLog.warn?.('TASK-PLANNER', '[needPlanningLLM] LLM 判断失败，保守返回 false:', e?.message);
            return false;
        }
    }

    /**
     * 关键词兜底：将描述性步骤文本映射到具体 excel_* 工具名
     * 用于 LLM 未产出合法工具名时（如"按地区分类汇总客户数"→ excel_aggregate），
     * 避免此类步骤被规划为 agent 策略而进入子 agent-loop 卡住
     * @param {string} text - 步骤描述文本
     * @returns {string|null} 匹配到的工具名，无匹配返回 null
     */
    function resolveToolByKeywords(text) {
        if (!text) return null;
        // 优先级从高到低由配置 toolMapping 顺序决定：图表 > 聚合 > 格式化 > 读取 > 复制 > 建表 > 写入 > 删除
        const rules = defs().toolMapping || [];
        for (const rule of rules) {
            if (rule.re && rule.re.test(text)) return rule.tool;
        }
        return null;
    }

    /**
     * 将用户请求分解为结构化步骤
     * @param {Object} opts
     * @param {string} opts.userMessage - 用户原始消息
     * @param {Array} opts.tools - 可用工具列表
     * @param {Array} opts.messages - 对话历史
     * @param {AbortSignal} opts.signal
     * @param {number} opts.maxSteps - 最大步骤数
     * @returns {Promise<Array<Object>>} 步骤数组
     */
    async function plan(opts) {
        const { userMessage, tools, messages, signal, maxSteps, intent, context } = opts || {};
        if (!userMessage) return [];

        // 提取实际数据列名（优先级：context.selection.headers > intent.targetData.columns）
        let actualColumns = [];
        if (context && context.selection && Array.isArray(context.selection.headers) && context.selection.headers.length > 0) {
            actualColumns = context.selection.headers.map(String);
        } else if (intent && intent.targetData && Array.isArray(intent.targetData.columns)) {
            actualColumns = intent.targetData.columns;
        }
        const columnConstraint = actualColumns.length > 0
            ? `\n- **重要：当前数据可用列名为：${actualColumns.join('、')}。你在 excel_aggregate / excel_date_group 等工具的 params 中填写的 dimColumn、valueColumn、dateColumn、aggColumn 等列名参数，必须是上述列名之一。如果用户请求中提到的维度/数值列名不在上述列表中，你必须传空字符串 ""，由系统在运行时自动推断。严禁编造不存在的列名。**`
            : '';

        // 构造工具列表摘要（供 LLM 参考）
        const toolSummaries = (tools || []).map(t => {
            const name = t.name || t.function?.name || '';
            const desc = t.description || t.function?.description || '';
            return `- ${name}: ${desc.slice(0, 100)}`;
        }).join('\n');

        const systemPrompt = `你是一个任务规划助手。请将用户请求拆分为多个执行步骤。

## 核心原则
- 包含"然后"、"再"、"接着"、"之后"等顺序词的请求，**必须**拆分为多个步骤
- 每个步骤必须是**单一原子操作**（一个步骤只调用一个工具）
- 两个不同的操作（如"读取"+"汇总"）必须拆成两个独立步骤
- 如果有"先...再..."、"先...然后..."模式，每个子句都是一个独立步骤

## 输出格式
返回 JSON 数组，每个元素：
[
  {
    "id": "step_1",
    "name": "简短步骤名",
    "description": "步骤描述（20字以内）",
    "strategy": "direct|agent|pipeline",
    "toolName": "工具名（direct策略必填）",
    "params": { "参数名": "参数值" },
    "displayArgs": { "显示名": "显示值" },
    "displayName": "工具卡片显示名",
    "prompt": "agent策略时的提示词",
    "pipelineKind": "pipeline策略时的管道名",
    "allowedTools": ["agent策略允许的工具列表"],
    "dependsOn": []
  }
]

## 策略选择规则
	- direct: 工具名匹配 excel_* 且参数可完全确定
	- agent: 需要 LLM 推理、参数不确定、或涉及分析/生成
	- pipeline: 需要路由到其他管道（skill/python/code/html）
	
	## 可用工具列表
	${toolSummaries || '（无可用工具列表）'}
	
	## 约束
	- 最多 ${maxSteps || 10} 个步骤
	- 步骤必须按执行顺序排列
	- 只返回 JSON 数组，不要任何多余解释或代码块
	- 如果请求是简单的问候/问答/不需要拆分为多步骤，返回空数组 []
		- **重要：所有 Excel 数据操作（读取/汇总/建图/格式化/复制/写入/删除/日期分组）必须使用 direct 策略 + 对应 excel_* 工具名，严禁使用 agent 策略。agent 策略仅用于需要 LLM 文本推理的步骤（如"总结数据特点"、"给出改进建议"）。**
		- **重要：用户请求"分析哪些产品卖得好/排名/对比"等数据分析需求时，属于数据汇总操作，必须使用 direct 策略 + excel_aggregate 工具（按产品分组汇总销售额），严禁使用 agent 策略。只有在前一步已经得到汇总数据后，如果需要 LLM 对汇总结果进行文字解读（如"给出销售建议"），才使用 agent 策略。**
		- **重要："汇总/聚合/统计/求和/分组/分类/平均/计数/分析（卖得好/排名/对比等数据分析）"等操作必须使用 direct 策略，工具名 excel_aggregate。参数 dimColumn/valueColumn：当用户在请求中明确指定了分组字段和数值字段时（如"按部门汇总客户数"→ dimColumn:"部门", valueColumn:"客户数"），必须填写用户指定的列名；仅当用户未指定时才传空字符串 ""（系统会自动推断列名）。**
		- **重要："按年/季度/月/周/日分组"等日期时间分组操作必须使用 direct 策略，工具名 excel_date_group。参数 dateColumn：用户指定的日期列名；groupBy：year/quarter/month/week/day；aggColumn（可选）：要聚合的数值列名。例如"按月汇总销售额"→ dateColumn:"日期", groupBy:"month", aggColumn:"销售额"。**
		- **重要：当用户请求涉及"季度/月份/年份/周/日"等时间维度分组时（如"按季度汇总销售额""按月统计订单数"），必须使用 excel_date_group 工具，而不是 excel_aggregate。因为 excel_aggregate 无法处理日期时间分组，只有 excel_date_group 能将日期列按时间粒度分组。例如"按季度分类汇总销售额"→ toolName:"excel_date_group", params:{"dateColumn":"日期","groupBy":"quarter","aggColumn":"销售额"}。**
		- **重要：params 中必须包含用户明确指定的数据列名（如"按部门汇总销售额"→ dimColumn:"部门", valueColumn:"销售额"；"按地区汇总客户数"→ dimColumn:"地区", valueColumn:"客户数"）。仅当用户未指定列名时才使用空字符串 ""，由系统自动从已读取的数据中推断。禁止把用户明确指定的列名改成空字符串。**
		- **重要：sheet 参数仅在用户明确指定工作表名时才填写（如"读取 Sheet1 的数据"→ sheet:"Sheet1"）；否则使用空字符串 ""（系统自动使用当前活动工作表）。禁止从用户消息中猜测工作表名（如"读取销售数据"不能推断出 sheet:"销售数据"）**
		- **重要：禁止将"复制数据到目标工作表"拆分为"复制"+"粘贴"两个独立步骤！"复制选中数据到新工作表"是一个完整的原子操作，使用 excel_copy_range 工具一步完成。**
		- **重要：当用户说"格式化表格"但未指定具体格式时，必须使用 direct 策略，工具名 excel_format_range，参数 format 设为空对象 "{}"（系统会自动应用商务蓝专业美化预设：表头蓝底白字加粗、斑马纹、细边框、冻结首行、智能列类型识别——日期/货币/整数格式自动应用）。仅当用户明确指定了具体格式（如"加粗"、"红色字体"、"居中"）时才在 format 中传对应字段。**
		- **重要：用户点名了颜色、主题或「统一为/改成/换成某外观」（例如中国红）时，format 必须写出该颜色（表头底色 + 表头白字），禁止用空对象 {}。空对象会被执行成商务蓝，等于丢掉用户指定的外观。要对多张工作表做时，先取工作表清单，再逐表格式化，不能只做一张。**
		- **重要：使用 excel_write_range 写入数据时，如果目标工作表不存在，系统会自动创建该工作表，不需要额外规划"创建新工作表"步骤。只需在 params.sheet 中指定目标工作表名即可。**
		- **重要：禁止规划重复的写入步骤！如果前面步骤已经使用 excel_write_range 将结果写入工作表，后续严禁再规划另一个写入步骤。一个完整的业务流程只需要一次写入操作。**
		- **重要：一个指标只允许一种计算方式！若已规划 excel_add_formula 公式步骤完成计算，严禁再规划 excel_aggregate 等统计步骤重复计算同一指标；同理，前一步已完成的动作不要再规划重复步骤。**
		- **重要：用户要求"按某维度出图"（如"使用饼图按部门展示销售额""各部门销量占比图"）时，必须规划两步：第1步 excel_aggregate（dimColumn=维度列、valueColumn=数值列、aggFunc=sum）汇总，第2步 excel_create_chart 建图；严禁只规划"建图"一步。图表步骤的 dataRange/sheet 由系统按上一步聚合结果自动填入，你只需给 chartType 和 title。**
		- **重要：饼图/圆环图只能表达单系列。系统会自动用"聚合结果"的类别列+数值列两列作为数据源；如果数据本身不是分类汇总表（而是明细表），更要走上面的"先聚合再建图"两步，禁止直接对明细整表建图。**
		${columnConstraint}
	
	## 示例
	用户请求："读取 Sheet1 的数据，然后按部门汇总"
	正确输出：
	[
	  {"id":"step_1","name":"读取数据","description":"读取 Sheet1 数据","strategy":"direct","toolName":"excel_read_range","params":{"sheet":"Sheet1","range":"A1:Z1000"},"displayArgs":{},"dependsOn":[]},
	  {"id":"step_2","name":"汇总","description":"按部门汇总数据","strategy":"direct","toolName":"excel_aggregate","params":{"dimColumn":"部门","valueColumn":"","aggFunc":"sum"},"displayArgs":{},"dependsOn":["step_1"]}
	]
	
	用户请求："读取销售数据，按地区分类汇总客户数，然后创建柱状图"
	正确输出：
	[
	  {"id":"step_1","name":"读取数据","description":"读取销售数据","strategy":"direct","toolName":"excel_read_range","params":{"sheet":"","range":""},"displayArgs":{},"dependsOn":[]},
	  {"id":"step_2","name":"汇总","description":"按地区分类汇总客户数","strategy":"direct","toolName":"excel_aggregate","params":{"dimColumn":"地区","valueColumn":"客户数","aggFunc":"sum"},"displayArgs":{},"dependsOn":["step_1"]},
	  {"id":"step_3","name":"创建柱状图","description":"创建柱状图","strategy":"direct","toolName":"excel_create_chart","params":{"chartType":"xlColumnClustered"},"displayArgs":{},"dependsOn":["step_2"]}
	]`;

        const userPrompt = `请将以下请求拆分为执行步骤：\n\n${userMessage}`;

        // 注入 LLM 意图分析结果（orchestrator Phase 6）：作为规划参考，
        // 让步骤分解基于语义理解（要做什么/对什么数据/产出什么），而非纯关键词。
        let intentPrompt = '';
        if (intent && typeof intent === 'object') {
            const parts = [];
            if (Array.isArray(intent.operations) && intent.operations.length) {
                parts.push('意图分析-操作：' + intent.operations.join('、'));
            }
            if (intent.targetData && intent.targetData.sheet) {
                parts.push('意图分析-目标工作表：' + intent.targetData.sheet);
            }
            if (intent.targetData && Array.isArray(intent.targetData.columns) && intent.targetData.columns.length) {
                parts.push('意图分析-目标列：' + intent.targetData.columns.join('、'));
                parts.push('【强制约束】上述列名是当前数据的实际列名。你在规划步骤时，params 中的列名参数（dimColumn/valueColumn/dateColumn/aggColumn 等）必须从上述列名中选择。如果用户提到的列名不在列表中，传空字符串 "" 由系统自动推断。');
            }
            if (intent.output && intent.output.type) {
                parts.push('意图分析-产出形式：' + intent.output.type + (intent.output.name ? '（' + intent.output.name + '）' : ''));
            }
            if (intent.summary) {
                parts.push('意图分析-总结：' + intent.summary);
            }
            // 【场景推理】注入意图分析层推断的参数，减少 TaskPlanner LLM 的猜测
            if (intent.inferredParams && typeof intent.inferredParams === 'object') {
                const ip = intent.inferredParams;
                const inferredParts = [];
                if (ip.dimColumn) inferredParts.push(`dimColumn:"${ip.dimColumn}"`);
                if (ip.valueColumn) inferredParts.push(`valueColumn:"${ip.valueColumn}"`);
                if (ip.dateColumn) inferredParts.push(`dateColumn:"${ip.dateColumn}"`);
                if (ip.aggColumn) inferredParts.push(`aggColumn:"${ip.aggColumn}"`);
                if (ip.groupBy) inferredParts.push(`groupBy:"${ip.groupBy}"`);
                if (ip.aggFunc && ip.aggFunc !== 'sum') inferredParts.push(`aggFunc:"${ip.aggFunc}"`);
                if (ip.chartType) inferredParts.push(`chartType:"${ip.chartType}"`);
                if (inferredParts.length) {
                    parts.push('意图分析-推断参数：' + inferredParts.join('，'));
                }
            }
            if (parts.length) {
                intentPrompt = '\n\n【LLM 意图分析参考（请据此规划步骤，与用户请求冲突时以用户请求为准）】\n' + parts.join('\n');
            }
        }

        try {
            const llmIo = App.ChatV2.llmIo;
            if (!llmIo || typeof llmIo.askJson !== 'function') {
                ChatLog.warn?.('TASK-PLANNER', '[plan] llmIo.askJson 不可用，返回空规划');
                return [];
            }

            const allMessages = [
                ...(Array.isArray(messages) ? messages.slice(-6) : []),
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt + intentPrompt }
            ];

            const res = await llmIo.askJson({ messages: allMessages, signal, config: App.config });
            const steps = res && res.json;

            if (!Array.isArray(steps) || steps.length === 0) {
                ChatLog.info?.('TASK-PLANNER', '[plan] LLM 返回空步骤，无需多步骤');
                return [];
            }

            // 校验步骤结构
            const validSteps = steps.filter(s => s && s.id && (s.name || s.description));
            if (validSteps.length === 0) {
                ChatLog.warn?.('TASK-PLANNER', '[plan] 所有步骤均无效，返回空规划');
                return [];
            }

            // 【兜底校验】如果 LLM 只返回了 1 步，但消息包含"然后/再/接着"等顺序词，
            // 说明 LLM 省略了步骤，此时强制要求 LLM 重新规划
            if (validSteps.length === 1 && /然后|再|接着|之后|先.*再|先.*然后/.test(userMessage)) {
                ChatLog.warn?.('TASK-PLANNER', `[plan] LLM 只返回了 1 步但消息包含顺序词，强制要求重新规划`);
                try {
                    const retryRes = await llmIo.askJson({
                        messages: [
                            ...allMessages,
                            { role: 'user', content: `你只返回了 1 个步骤，但用户请求包含"然后"等顺序词，需要拆分。请重新拆分，确保每个独立操作都是一个独立步骤。` }
                        ],
                        signal,
                        config: App.config
                    });
                    const retrySteps = retryRes && retryRes.json;
                    if (Array.isArray(retrySteps) && retrySteps.length >= 2) {
                        const retryValid = retrySteps.filter(s => s && s.id && (s.name || s.description));
                        if (retryValid.length >= 2) {
                            ChatLog.info?.('TASK-PLANNER', `[plan] 重新规划成功: ${retryValid.length} 个步骤`);
                            validSteps.length = 0;
                            validSteps.push(...retryValid);
                        }
                    }
                } catch (retryErr) {
                    ChatLog.warn?.('TASK-PLANNER', '[plan] 重新规划失败，使用原始规划:', retryErr?.message);
                }
            }

            // 补充默认字段
            for (const s of validSteps) {
                if (!s.strategy) s.strategy = 'direct';
                if (!s.params) s.params = {};
                if (!s.dependsOn) s.dependsOn = [];
                s.status = 'pending';

                // 强制后处理：所有 excel_ 前缀的 MCP 工具必须使用 direct 策略
                // agent 策略会进入子 agent-loop 导致步骤卡住或 LLM 绕路
                // 中文显示名（如"新建工作表"）经 toolMapping 配置识别，不内嵌关键词
                const isExcelTool = (s.toolName && s.toolName.startsWith('excel_')) ||
                    resolveToolByKeywords(s.toolName || s.tool || '') !== null;
                if (isExcelTool && s.strategy === 'agent') {
                    ChatLog.info?.('TASK-PLANNER', `[plan] 强制 ${s.toolName || s.tool} 策略从 agent → direct`);
                    s.strategy = 'direct';
                }

                // 关键词兜底：LLM 未产出合法 excel_* 工具名时，按步骤描述关键词映射到对应工具（direct 策略）
                // 覆盖两种情况：① toolName 为空或非 excel_* 前缀；② toolName 是中文显示名（如"分析"）经 toolMapping 识别
                // 避免"按地区分类汇总客户数"这类描述性步骤被规划为 agent 而卡住
                const hasExcelToolName = !!(s.toolName && s.toolName.startsWith('excel_'));
                if (!hasExcelToolName) {
                    const stepText = `${s.toolName || ''} ${s.tool || ''} ${s.name || ''} ${s.description || ''}`;
                    const mappedTool = resolveToolByKeywords(stepText);
                    if (mappedTool) {
                        ChatLog.info?.('TASK-PLANNER', `[plan] 关键词兜底映射: "${stepText.trim().slice(0, 30)}" → ${mappedTool} (direct)`);
                        s.toolName = mappedTool;
                        s.strategy = 'direct';
                    }
                }

                // 【列名校验】如果 LLM 填写的列名不在实际数据列中，自动清空让 prepareStepParams 推断
                if (actualColumns.length > 0 && s.params) {
                    const colFields = ['dimColumn', 'valueColumn', 'dateColumn', 'aggColumn', 'groupBy'];
                    for (const field of colFields) {
                        const val = s.params[field];
                        if (val && typeof val === 'string' && val.trim() !== '') {
                            const exists = actualColumns.some(c => c.toLowerCase() === val.toLowerCase());
                            if (!exists) {
                                ChatLog.warn?.('TASK-PLANNER', `[plan] 列名校验: ${field}="${val}" 不在实际列名 [${actualColumns.join(',')}] 中，自动清空`);
                                s.params[field] = '';
                            }
                        }
                    }
                }

                // 【场景推理】用 intent-analyzer 推断的参数填充 LLM 未填或填空的参数
                if (intent && intent.inferredParams && typeof intent.inferredParams === 'object') {
                    const ip = intent.inferredParams;
                    const paramMap = {
                        'excel_aggregate': ['dimColumn', 'valueColumn', 'aggFunc'],
                        'excel_date_group': ['dateColumn', 'groupBy', 'aggColumn', 'aggFunc'],
                        'excel_create_chart': ['chartType']
                    };
                    const relevantFields = paramMap[s.toolName] || [];
                    for (const field of relevantFields) {
                        const inferredVal = ip[field];
                        const currentVal = s.params[field];
                        // 只有当 LLM 未填（空字符串）或推断值更具体时，才用推断值覆盖
                        if (inferredVal && (!currentVal || currentVal === '')) {
                            ChatLog.info?.('TASK-PLANNER', `[plan] 场景推理填充: ${s.toolName}.${field}="${inferredVal}"（LLM 未填）`);
                            s.params[field] = inferredVal;
                        }
                    }
                }
            }

            ChatLog.info?.('TASK-PLANNER', `[plan] 规划完成: ${validSteps.length} 个步骤`);
            return validSteps;
        } catch (e) {
            ChatLog.warn?.('TASK-PLANNER', '[plan] 规划异常:', e?.message);
            return [];
        }
    }

    /**
     * 动态重规划：基于已执行步骤的观察结果，规划下一步（或决定任务完成）
     * 这是 ReAct 架构的核心：观察 → 推理 → 行动
     * @param {Object} opts
     * @param {string} opts.userMessage - 原始用户请求
     * @param {Array} opts.executedSteps - 已执行的步骤 [{step, result}]
     * @param {Array} opts.tools - 可用工具列表
     * @param {Array} opts.messages - 消息上下文
     * @param {AbortSignal} opts.signal
     * @param {Object} opts.intent - 意图分析结果
     * @returns {Promise<Array>} 下一步步骤数组（空数组表示任务完成）
     */
    async function replan(opts) {
        const { userMessage, executedSteps, tools, messages, signal, intent } = opts || {};
        if (!userMessage) return [];

        // 构建观察上下文：已执行步骤的结果摘要
        const observations = buildObservations(executedSteps);
        const lastResult = executedSteps.length > 0 ? executedSteps[executedSteps.length - 1].result : null;

        // 提取实际数据列名（从最后一步的结果中）
        let actualColumns = [];
        if (intent && intent.targetData && Array.isArray(intent.targetData.columns)) {
            actualColumns = intent.targetData.columns;
        }
        // 如果最后一步是读取操作，从结果中提取列名
        if (lastResult && lastResult.ok) {
            const cols = extractColumnsFromResult(lastResult);
            if (cols.length > 0) actualColumns = cols;
        }

        const columnConstraint = actualColumns.length > 0
            ? `\n- **重要：当前数据可用列名为：${actualColumns.join('、')}。你在 excel_aggregate / excel_date_group 等工具的 params 中填写的 dimColumn、valueColumn、dateColumn、aggColumn 等列名参数，必须是上述列名之一。如果用户请求中提到的维度/数值列名不在上述列表中，你必须传空字符串 ""，由系统在运行时自动推断。严禁编造不存在的列名。**`
            : '';

        const toolSummaries = (tools || []).map(t => {
            const name = t.name || t.function?.name || '';
            const desc = t.description || t.function?.description || '';
            return `- ${name}: ${desc.slice(0, 100)}`;
        }).join('\n');

        const systemPrompt = `你是一个任务规划助手。用户请求正在执行中，部分步骤已完成。请根据【观察结果】判断下一步该做什么，或是否已完成任务。

## 核心原则
- 如果任务已经完成（如已读取数据、已汇总、已写入），返回空数组 []
- 如果还需要继续执行，只返回**下一步**的单个步骤（不要一次性规划多步）
- 每个步骤必须是**单一原子操作**
- 基于观察结果中的实际数据列名来填写参数，严禁编造

## 输出格式
返回 JSON 数组（0 个或 1 个元素）：
[
  {
    "id": "步骤唯一标识",
    "name": "步骤名称",
    "description": "步骤描述",
    "strategy": "direct",
    "toolName": "excel_read_range / excel_aggregate / excel_date_group / excel_create_chart / excel_format_range / excel_copy_range / excel_write_range / excel_delete_range",
    "params": { "参数名": "参数值" },
    "displayArgs": { "显示名": "显示值" },
    "dependsOn": []
  }
]

## 约束
- 只返回 JSON 数组，不要任何多余解释
- 如果任务已完成，返回空数组 []
- 所有 Excel 数据操作必须使用 direct 策略 + 对应 excel_* 工具名
- params 中必须包含用户明确指定的数据列名；仅当用户未指定时才使用空字符串 ""${columnConstraint}

## 可用工具
${toolSummaries || '(无工具列表)'}`;

        const userPrompt = `原始请求：${userMessage}\n\n【已完成的步骤】\n${observations}\n\n请判断：任务是否已完成？如果未完成，请规划下一步该做什么（只返回下一步）。\n\n注意：\n- 如果已经读取了数据但还没分析/汇总，下一步应该是分析/汇总（使用观察结果中的实际列名填写参数）\n- 如果已经分析/汇总完成，下一步应该是写入结果或创建图表\n- 所有参数必须从观察结果中的实际列名中选择，严禁编造不存在的列名`;

        try {
            const llmIo = App.ChatV2.llmIo;
            if (!llmIo || typeof llmIo.askJson !== 'function') {
                ChatLog.warn?.('TASK-PLANNER', '[replan] llmIo.askJson 不可用');
                return [];
            }

            const allMessages = [
                ...(Array.isArray(messages) ? messages.slice(-6) : []),
                { role: 'system', content: systemPrompt },
                { role: 'user', content: userPrompt }
            ];

            const res = await llmIo.askJson({ messages: allMessages, signal, config: App.config });
            const steps = res && res.json;

            if (!Array.isArray(steps)) {
                ChatLog.info?.('TASK-PLANNER', '[replan] LLM 返回非数组，视为任务完成');
                return [];
            }

            if (steps.length === 0) {
                ChatLog.info?.('TASK-PLANNER', '[replan] LLM 返回空数组，任务完成');
                return [];
            }

            // 只取第一个步骤（replan 每次只规划下一步）
            const step = steps[0];
            if (!step || !step.id || !(step.name || step.description)) {
                ChatLog.warn?.('TASK-PLANNER', '[replan] 步骤结构无效');
                return [];
            }

            // 补充默认字段
            if (!step.strategy) step.strategy = 'direct';
            if (!step.params) step.params = {};
            if (!step.dependsOn) step.dependsOn = [];

            // 强制 excel_ 工具使用 direct 策略
            const isExcelTool = (step.toolName && step.toolName.startsWith('excel_')) ||
                resolveToolByKeywords(step.toolName || step.tool || '') !== null;
            if (isExcelTool && step.strategy === 'agent') {
                step.strategy = 'direct';
            }

            // 关键词兜底
            const hasExcelToolName = !!(step.toolName && step.toolName.startsWith('excel_'));
            if (!hasExcelToolName) {
                const stepText = `${step.toolName || ''} ${step.tool || ''} ${step.name || ''} ${step.description || ''}`;
                const mappedTool = resolveToolByKeywords(stepText);
                if (mappedTool) {
                    step.toolName = mappedTool;
                    step.strategy = 'direct';
                }
            }

            // 列名校验
            if (actualColumns.length > 0 && step.params) {
                const colFields = ['dimColumn', 'valueColumn', 'dateColumn', 'aggColumn', 'groupBy'];
                for (const field of colFields) {
                    const val = step.params[field];
                    if (val && typeof val === 'string' && val.trim() !== '') {
                        const exists = actualColumns.some(c => c.toLowerCase() === val.toLowerCase());
                        if (!exists) {
                            ChatLog.warn?.('TASK-PLANNER', `[replan] 列名校验: ${field}="${val}" 不在实际列名中，自动清空`);
                            step.params[field] = '';
                        }
                    }
                }
            }

            // 用 intent 推断参数填充
            if (intent && intent.inferredParams && typeof intent.inferredParams === 'object') {
                const ip = intent.inferredParams;
                const paramMap = {
                    'excel_aggregate': ['dimColumn', 'valueColumn', 'aggFunc'],
                    'excel_date_group': ['dateColumn', 'groupBy', 'aggColumn', 'aggFunc'],
                    'excel_create_chart': ['chartType']
                };
                const relevantFields = paramMap[step.toolName] || [];
                for (const field of relevantFields) {
                    const inferredVal = ip[field];
                    const currentVal = step.params[field];
                    if (inferredVal && (!currentVal || currentVal === '')) {
                        step.params[field] = inferredVal;
                    }
                }
            }

            ChatLog.info?.('TASK-PLANNER', `[replan] 动态规划下一步: ${step.name || step.description} (${step.toolName})`);
            return [step];
        } catch (e) {
            ChatLog.warn?.('TASK-PLANNER', '[replan] 重规划异常:', e?.message);
            return [];
        }
    }

    /**
     * 从执行结果中提取观察摘要
     */
    function buildObservations(executedSteps) {
        return executedSteps.map(({ step, result }, idx) => {
            const status = result && result.ok ? '✅' : '❌';
            const tool = step.toolName || step.tool || '未知工具';
            const desc = step.description || step.name || `步骤${idx + 1}`;
            let detail = '';
            if (result && result.ok) {
                // 提取关键信息：列名、行数、工作表名等
                const data = result.result || result.data || result;
                if (data && typeof data === 'object') {
                    if (Array.isArray(data.headers)) detail += ` 列名:[${data.headers.join(',')}]`;
                    if (Array.isArray(data.data)) detail += ` 行数:${data.data.length}`;
                    if (data.worksheet) detail += ` 工作表:${data.worksheet}`;
                    if (data.range) detail += ` 范围:${data.range}`;
                }
            } else if (result && result.error) {
                detail += ` 错误:${typeof result.error === 'string' ? result.error : result.error.message || '未知错误'}`;
            }
            return `${status} ${idx + 1}. [${tool}] ${desc}${detail}`;
        }).join('\n');
    }

    /**
     * 从执行结果中提取列名
     */
    function extractColumnsFromResult(result) {
        if (!result || !result.ok) return [];
        const data = result.result || result.data || result;
        if (!data || typeof data !== 'object') return [];
        if (Array.isArray(data.headers)) return data.headers.map(String);
        if (Array.isArray(data.data) && data.data.length > 0 && typeof data.data[0] === 'object') {
            return Object.keys(data.data[0]).map(String);
        }
        if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object') {
            return Object.keys(data[0]).map(String);
        }
        return [];
    }

    App.ChatV2.TaskPlanner = { needPlanning, needPlanningLLM, plan, replan, resolveToolByKeywords };
    ChatLog.info?.('CHAT-V2', '✅ TaskPlanner loaded (with dynamic replan)');
})();