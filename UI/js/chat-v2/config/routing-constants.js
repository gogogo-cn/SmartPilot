/**
 * ============================================================
 * chat-v2/config/routing-constants.js — 前端路由常量单一事实源
 * ============================================================
 * 职责：收编散落在 chat-orchestrator 等处的路由常量与正则，统一在此维护。
 *   - COMMON_2CHAR_WORDS          通用 2 字虚词（防误路由护栏）
 *   - detectLanguageDirective     编程语言指令检测（python/vba）
 *   - detectCodeGenerationIntent  代码生成意图检测
 *   - PYTHON_ANALYSIS_INTENT      Python 数据分析意图正则（收编 orchestrator 三处重复内联正则）
 *   - detectPythonAnalysisIntent  统一检测函数（返回结构化结果）
 *
 * 兼容：
 *   - 浏览器：挂载到 App.ChatV2.routingConstants
 *   - Node 测试：module.exports 导出（供 Tests/*.test.js 直接 require）
 *
 * 注意：PYTHON_ANALYSIS_INTENT 为"标准集"（原 matchedSkill 分支使用的完整正则），
 *       PYTHON_ANALYSIS_INTENT_CONSERVATIVE 为"保守集"（原最终兜底分支使用的子集，
 *       不含"变化"等宽泛词，避免无 Skill/无选区时误路由）。两套均在此维护，
 *       改动路由规则只需改本文件。
 * ============================================================
 */
(function() {
    'use strict';
    if (typeof window !== 'undefined' && window) {
        if (!window.App) window.App = {};
        if (!App.ChatV2) App.ChatV2 = {};
    }

    // 通用 2 字虚词集合：
    // 仅命中这些词（无 3+ 字短语）时不视为强意图，防止误路由（如"处理数据"不应命中所有 skill）。
    // 注意：只放"多个 skill 共有、意图含糊"的通用词；"合并/拆分/去重/美化/清洗"等具体操作词
    // 因有"唯一候选"保护（candidates.length === 1）不会误路由，故不放进来。
    const COMMON_2CHAR_WORDS = new Set([
        '技能', '创建', '生成', '处理', '分析', '数据', '文件', '操作', '设置',
        '管理', '查看', '显示', '打开', '关闭', '保存', '导出', '导入', '转换',
        '复制', '粘贴', '查找', '替换', '修改', '更新', '添加', '删除', '插入',
        '选择', '计算', '统计', '汇总', '提取', '识别', '上传', '下载', '打印',
        '预览', '刷新', '重置', '提交', '启用', '禁用', '允许', '阻止', '检查',
        '测试', '验证', '确认', '取消', '帮助', '介绍', '解释', '说明', '什么',
        '如何', '怎样', '怎么', '为什么', '内容', '信息', '结果', '报告', '问题'
    ]);

    /**
     * 检测用户消息中的编程语言指令
     * @returns {string|null} 'python' | 'vba' | null
     */
    function detectLanguageDirective(message) {
        if (!message) return null;
        const lower = message.toLowerCase();
        if (/\bpython\b|py脚本|py代码|python代码/.test(lower)) return 'python';
        if (/\bvba\b|宏|vb代码|vba代码/.test(lower)) return 'vba';
        return null;
    }

    /**
     * OCR/表格"解析脚本、解析器"意图（GenerateOCRParseScript 的产物域）。
     * 与"让 AI 写代码/宏"是两码事，判代码生成意图时必须先排掉。
     */
    const OCR_PARSE_SCRIPT_INTENT = /解析器|解析脚本|解析规则|提取脚本|提取规则|字段提取脚本|ocr\s*解析/i;

    /**
     * 明确点名了代码载体/语言（代码、宏、VBA、Python）时的提示词。
     * 注意：不要用 \b 包中文——中文字符在 JS 正则里是非单词字符，\b代码\b 根本匹配不上。
     */
    const CODE_ARTIFACT_HINT = /代码|宏|vba|python|py脚本/i;

    /**
     * 检测用户是否有代码生成意图
     */
    function detectCodeGenerationIntent(message) {
        if (!message) return false;
        // 【防误路由】"生成/写一个解析脚本"是在要 OCR 解析器，不是在要代码。
        // 原正则里裸的 `脚本` 会把任何含"脚本"二字的消息判成代码生成意图，实测
        // "帮我做一个发票解析脚本" → 强制进 code-pipeline → GenerateVBA
        // （用户："这都能选错技能？"）。只有同时点名代码/宏/VBA/Python 时才保留代码意图。
        if (OCR_PARSE_SCRIPT_INTENT.test(message) && !detectLanguageDirective(message) && !CODE_ARTIFACT_HINT.test(message)) {
            return false;
        }
        const lower = message.toLowerCase();
        return /写.*代码|生成.*代码|写.*脚本|生成.*脚本|写.*宏|生成.*宏|code|脚本/.test(lower);
    }

    // ================================================================
    // Python 数据分析意图正则（单一事实源）
    // 收编自 chat-orchestrator 三处重复内联正则（matchedSkill / 无Skill / 最终兜底）。
    // 标准集 = 原 matchedSkill 分支的完整正则；保守集 = 原最终兜底分支的子集。
    // ================================================================

    const PYTHON_ANALYSIS_INTENT = {
        trend: /趋势|走向|变化|forecast|预测/,
        correlation: /相关|关联|correlation/,
        anomaly: /异常|离群|outlier|anomaly/,
        descriptive: /描述(?:性)?统计|统计报告|均值|中位数|标准差/,
        distribution: /分布|频率|直方图|histogram/,
        normalize: /归一化|标准化|正规化|min-?max|z-?score|z分数|极差法|小数定标/,
        multi: /趋势.*异常|异常.*趋势|同时.*趋势|同时.*异常|找出.*趋势.*异常|找出.*异常.*趋势/,
        groupAgg: /分组|聚合|group|aggregate|按.*分组|各.*部门|分.*部门|按.*类别|分.*类别/
    };

    // 保守集：最终兜底分支使用，排除"变化"等宽泛词，避免无 Skill/无选区时误路由
    const PYTHON_ANALYSIS_INTENT_CONSERVATIVE = {
        trend: /趋势|走向|forecast|预测/,
        correlation: /相关|关联|correlation/,
        anomaly: /异常|离群|outlier|anomaly/,
        descriptive: /描述(?:性)?统计|统计报告|均值|中位数/,
        distribution: /分布|频率|直方图/,
        normalize: /归一化|标准化|正规化|min-?max|z-?score|z分数/,
        multi: /趋势.*异常|异常.*趋势|同时.*趋势|同时.*异常/,
        groupAgg: /分组|聚合|按.*分组|各.*部门|分.*部门|按.*类别|分.*类别/
    };

    /**
     * 剥离消息里"被引用的标识符"（中英文引号/书名号内的内容）。
     * 动机：用户常把工作表名/列名加引号，如 `删除工作表"异常场景测试表"`；
     * 引号内的"异常/趋势/分布/相关…"会污染数据分析意图关键词匹配，导致
     * "删除工作表"被误强制路由到 python-pipeline（异常值检测）。
     * 注意：仅剥离"引号内"文本，引号外的分析词（如 `对"销售额"做异常检测`）仍会被命中。
     * @param {string} message
     * @returns {string}
     */
    function stripQuotedSegments(message) {
        if (!message || typeof message !== 'string') return message || '';
        return message
            .replace(/"[^"]*"/g, ' ')
            .replace(/'[^']*'/g, ' ')
            .replace(/“[^”]*”/g, ' ')
            .replace(/‘[^’]*’/g, ' ')
            .replace(/「[^」]*」/g, ' ')
            .replace(/『[^』]*』/g, ' ')
            .replace(/【[^】]*】/g, ' ');
    }

    /**
     * 是否"工作表管理"指令（删除/新建/创建/重命名/改名/复制工作表）。
     * 这类指令属确定性工具操作（excel_delete_sheet / excel_create_sheet /
     * excel_rename_sheet / excel_copy_sheet），即使消息里含"异常/趋势"等分析词
     * （多来自工作表名）也必须优先，不得被强制路由成 Python 数据分析。
     */
    const SHEET_MANAGE_RE =
        /(?:删除|删掉|移除|去掉|清掉|新建|创建|增加|添加|重命名|改名|复制|拷贝|克隆)[^，。；;!！？?\n]{0,12}?(?:工作表|sheet|页签|标签页)|(?:工作表|sheet|页签|标签页)[^，。；;!！？?\n]{0,12}?(?:删除|删掉|移除|去掉|新建|创建|重命名|改名|复制|拷贝)/i;

    /** 判断消息是否为工作表管理指令 */
    function isSheetManagementIntent(message) {
        if (!message || typeof message !== 'string') return false;
        const stripped = stripQuotedSegments(message);
        return SHEET_MANAGE_RE.test(stripped);
    }

    /**
     * 检测消息是否含 Python 数据分析意图。
     * @param {string} message
     * @param {{conservative?: boolean}} [opts] - conservative=true 时使用保守集（最终兜底）
     * @returns {{trend:boolean, correlation:boolean, anomaly:boolean, descriptive:boolean,
     *            distribution:boolean, multi:boolean, groupAgg:boolean, any:boolean}}
     */
    function detectPythonAnalysisIntent(message, opts) {
        const defs = (opts && opts.conservative) ? PYTHON_ANALYSIS_INTENT_CONSERVATIVE : PYTHON_ANALYSIS_INTENT;
        // ① 工作表管理指令（删/建/重命名/复制工作表）优先：绝不被分析词强制劫持到 python。
        // ② 意图匹配前剥离引号内的标识符，避免工作表名/列名里的"异常/趋势"等词污染判定。
        const _empty = {
            trend: false, correlation: false, anomaly: false, descriptive: false,
            distribution: false, normalize: false, multi: false, groupAgg: false, any: false
        };
        if (isSheetManagementIntent(message)) return _empty;
        const lower = stripQuotedSegments(message).toLowerCase();
        const r = {
            trend: defs.trend.test(lower),
            correlation: defs.correlation.test(lower),
            anomaly: defs.anomaly.test(lower),
            descriptive: defs.descriptive.test(lower),
            distribution: defs.distribution.test(lower),
            normalize: defs.normalize.test(lower),
            multi: defs.multi.test(lower),
            groupAgg: defs.groupAgg.test(lower)
        };
        r.any = r.trend || r.correlation || r.anomaly || r.descriptive || r.distribution || r.multi || r.normalize;
        return r;
    }

    /**
     * 增删改查（CRUD）MCP 工具名集合（单一事实源）。
     * 用途：① text-pipeline 把这类工具从"通用对话"LLM 可见列表剔除——基础读写/行列增删清空
     *       统一改走 ExcelCRUD Skill 的确定性工具链，不再由通用对话直接接触裸 MCP 工具；
     *       ② orchestrator 判定 CRUD 意图。
     * 注意：excel_clear_sheet / excel_delete_sheet 属"工作表管理"，不归 CRUD；
     *       excel_add_formula 属公式写入（路由到 excel_add_formula），不归 CRUD，避免冲突。
     */
    const CRUD_TOOL_NAMES = new Set([
        // 读·查
        'excel_read_range', 'excel_get_cell_value', 'excel_search',
        // 写·改
        'excel_write_range', 'excel_set_cell_value', 'excel_fill_data',
        // 增·插入
        'excel_insert_row', 'excel_insert_rows', 'excel_insert_column',
        // 删·清空
        'excel_delete_row', 'excel_delete_column',
        'excel_clear_range', 'excel_clear_format'
    ]);

    /** 判断工具名是否属增删改查（CRUD）工具 */
    function isCrudTool(name) {
        return typeof name === 'string' && CRUD_TOOL_NAMES.has(name);
    }

    // ================================================================
    // 图表类型切换（仅换类型）路由正则（单一事实源）
    // 命中"换成柱状图/改成折线图"等明确仅切图表类型的后续消息。
    // 与 skill-pipelineModules.constants.CHART_TYPE_SWITCH_RE 语义一致，
    // 但本处为路由层判定（继承上一轮图表数据上下文），split-location 到路由模块。
    // ================================================================
    const CHART_TYPE_SWITCH_RE = /(?:换成|改成|改为|变成|转为|切换(?:为|到)?|换(?:一下|个)?)\s*(?:柱状|柱形|折线|饼|环形|圆环|面积|散点|条形|组合|雷达|漏斗|瀑布|旭日|树状|堆积|堆叠|仪表)[^。，！？]{0,8}\s*图?/;

    // ================================================================
    // Excel 处理意图判定（单一事实源）
    // ----------------------------------------------------------------
    // 用途：agent-loop 在"模型只回文字、没调工具"时的重试/强制兜底机制，只应为
    //       "确有 Excel 操作/选区引用"的请求启用；对无关问答（天气/新闻/闲聊等）
    //       不得启动工具引导，更不得把破坏性工具（clear_range 等）强推到用户数据上。
    // 命中任一"操作动词/选中引用"即视为有 Excel 处理意图；否则视为纯咨询放行文字回复。
    // 注意：勿放入"数据"等过宽泛词，避免把"XX数据怎么样"等查询误判成处理意图。
    // ================================================================
    const EXCEL_OPERATION_INTENT_RE =
        /清除|清空|清洗|删除|去掉|去除|删掉|清掉|格式化|美化|样式|配色|边框|冻结|合并|拆分|分列|插入|新增|追加|写入|写出|读取|读出|查看|更新|替换|修改|编辑|求和|合计|总计|汇总|统计|平均|均值|中位数|图表|柱状|折线|饼图|散点|排序|筛选|透视|去重|填充|调整|转置|隐藏|显示|生成|创建|算出|计算|选中|所选|选区|选好的|这块|这段|这个区域|当前区域|这个区域|这张表|这个表|当前表|工作表|工作簿|单元格|几行|几列|行列/;

    // 纯文本变换（翻译/润色/改写…）：虽常含「选中的内容」，但不是表格美化/Excel 写操作。
    // 若仍判 userOpIntent=true，零工具兜底会强推 word_format_selected_table_cell 空参（实测翻译请求翻车）。
    const CONTENT_TRANSFORM_INTENT_RE = /(翻译|译成|译为|译作|英译|中译|润色|改写|重写|缩写|扩写|总结选中|摘要选中|概括选中|把.*写成|改成英文|改成中文)/;
    // 同句若还有明确格式化动词，仍算表格/Excel 操作意图
    const CONTENT_TRANSFORM_FMT_OVERRIDE_RE = /(美化|格式化|样式|边框|底色|底纹|加粗|居中|对齐|清除格式|配色|主题|隔行|表头底)/;

    /**
     * 用户消息是否带有 Excel/表格写操作意图（操作动词 或 对当前选区/表的引用）。
     * 无关问答（如"东莞今天天气怎么样"）不命中，返回 false。
     * 纯翻译/润色等文本变换也不命中（避免强推格式化工具）。
     */
    function hasExcelOperationIntent(message) {
        if (!message || typeof message !== 'string') return false;
        if (CONTENT_TRANSFORM_INTENT_RE.test(message) && !CONTENT_TRANSFORM_FMT_OVERRIDE_RE.test(message)) {
            return false;
        }
        return EXCEL_OPERATION_INTENT_RE.test(message);
    }

    // ================================================================
    // 确定性硬短路触发关键词（单一事实源）
    // ----------------------------------------------------------------
    // 收编自 mcp-dispatcher / chat-orchestrator 的内联正则与关键词数组。
    // 这些关键词是"McpDispatcher 是否对消息做确定性直调"的触发条件，
    // 此前散落在多处、各写一份，极易"改一处漏一处"导致硬短路行为漂移。
    // 改动硬短路匹配规则只需改本文件。
    // ================================================================

    // 导出/另存为动词（非 Excel 宿主文档级导出到 PDF/图片 的触发动词）
    const EXPORT_VERB = /(导出|输出|另存|转成|转换成|生成|保存为|保存成|存为|存成)/;

    // PDF 产物关键词（与 EXPORT_VERB 配合识别"导出 PDF"）
    const PDF_KW = /(pdf|便携文档)/i;

    // 筛选/过滤触发关键词（命中 → excel_apply_filter 硬短路）
    const FILTER_KW = /(筛选|过滤|筛出)/;

    // 错误检查触发关键词（命中 + 错误关键词 → excel_check_errors 硬短路）
    const CHECK_KW = /(检查|检测|排查|查找|看看|查看|扫描)/;

    // 合并工作表触发关键词（命中 → excel_merge_sheets 硬短路）
    // 与"合并单元格"(excel_merge_cells)、"合并多文件"(excel_merge_workbooks) 共用"合并"这个动词，
    // 因此必须与下面的 MERGE_SHEET_TARGET_KW / MERGE_SHEET_NON_KW 组合判定，不能单靠本正则直调。
    const MERGE_SHEET_KW = /(合并|汇合|合表|合到一张|合并到一张|多表合并|汇总多表)/;

    // 合并对象确认词：消息必须指向"工作表/多表/数据"级别，而不是单元格。
    const MERGE_SHEET_TARGET_KW = /(工作表|sheet|多表|多张表|多张工作表|这几张表|选中表|所选表|明细表|数据|合到一张|一张表|汇总表)/i;

    // 合并意图的"复合/无关"排除词：命中即放行给多步骤管道或多文件工具，不在 dispatcher 里独占。
    // - 图表/看板/分析/报告：合并只是前序步骤，后面还有别的活（如"合并后画图"）。
    // - 清洗/去重/美化/排序/筛选/导出/ppt/vba：同属其它工具族，不能拿合并把它们截走。
    const MERGE_SHEET_NON_KW = /(图表|画图|画个|画一[个张]|绘制|制图|可视化|看板|大屏|饼图|柱状|柱形|折线|条形|散点|雷达|漏斗|甘特|环形|chart|plot|透视|分析|报告|清洗|去重|美化|格式化|排序|筛选|导出|短信|邮件|ppt|幻灯片|vba|代码|宏)/i;

    // 单元格合并措辞（"合并 A1:B2 单元格"）——属 excel_merge_cells，不在本硬短路范围内
    const MERGE_CELL_ONLY_RE = /(?:合并|取消合并)[^。，,]{0,8}单元格|单元格[^。，,]{0,8}(?:合并|取消合并)/;

    // 多文件合并措辞（"合并多个 xlsx/工作簿"）——excel_merge_workbooks 需要文件路径，
    // 无路径无法确定性直调，放行给 agent-loop 让其问询/规划。
    const MERGE_WORKBOOK_ONLY_RE = /(?:合并|汇总)[^。，,]{0,12}(?:文件|xlsx|多个工作簿|所有工作簿|工作簿文件)/i;

    // Excel 单元格格式化动词（命中 → excel_format_range 硬短路）
    const FMT_VERBS_EXCEL = ['加粗', '水平居中', '垂直居中', '居中', '倾斜', '斜体', '边框', '字体颜色', '字体', '字号', '填充色', '背景色', '底色', '行高', '列宽', '自动换行', '高亮', '清除格式', '设置格式', '格式化', '单元格格式', '对齐'];

    // 非 Excel 宿主（PPT/Word）选中表格格式化动词（命中 → ppt/word_format_selected_table_cell 硬短路）。
    // 与 FMT_VERBS_EXCEL 用途不同（不同宿主、不同目标工具），且关键词集合有差异
    // （本列表含「美化/样式/表头/颜色/底纹」、不含「水平居中/垂直居中/行高/列宽/自动换行/高亮/清除格式」），
    // 故保留为独立列表，不可与 FMT_VERBS_EXCEL 合并，否则会改变任一宿主的命中范围。
    // 2026-10-07 增补「样式方案/版式」词：选中表格时说"三线表/无边框留白/卡片条纹/深表头/紧凑…"
    // 也是 100% 确定性的排版请求，不该落到 LLM（此前只有"美化/样式"等词才命中，说"三线表"会漏）。
    const FMT_VERBS_HOST = ['加粗', '居中', '斜体', '倾斜', '边框', '字体', '字号', '填充色', '背景色', '底色', '底纹', '对齐', '设置格式', '格式化', '单元格格式', '美化', '样式', '表头', '颜色',
        '版式', '三线表', '无边框', '留白', '卡片', '条纹', '斑马纹', '隔行', '紧凑', '宽松', '行高', '首列', '合计行', '汇总行', '末行', '主题'];

    // 强图表意图（图表类型词 / 可视化 / 出图）。
    // 用于 orchestrator 补救守卫 D：matchedSkill 被置空后，消息若含明确图表类型词，
    // 重找图表 Skill；找不到则在 context 上打 _forceChart 标记让 text-pipeline 知道要出图。
    const STRONG_CHART_INTENT_RE = /饼图|柱图|柱状图|条形图|折线图|散点图|面积图|雷达图|直方图|趋势图|仪表盘|dashboard|数据大屏|可视化|图形化|画图|制图|做图|出图|生成.*图|创建.*图|画.*图|图表类型|chart\s*(pie|bar|line|scatter|area|radar|histogram|gauge|funnel)?/i;

    const api = {
        COMMON_2CHAR_WORDS,
        detectLanguageDirective,
        detectCodeGenerationIntent,
        PYTHON_ANALYSIS_INTENT,
        PYTHON_ANALYSIS_INTENT_CONSERVATIVE,
        stripQuotedSegments,
        isSheetManagementIntent,
        detectPythonAnalysisIntent,
        CRUD_TOOL_NAMES,
        isCrudTool,
        CHART_TYPE_SWITCH_RE,
        hasExcelOperationIntent,
        EXPORT_VERB,
        PDF_KW,
        FILTER_KW,
        CHECK_KW,
        MERGE_SHEET_KW,
        MERGE_SHEET_TARGET_KW,
        MERGE_SHEET_NON_KW,
        MERGE_CELL_ONLY_RE,
        MERGE_WORKBOOK_ONLY_RE,
        FMT_VERBS_EXCEL,
        FMT_VERBS_HOST,
        STRONG_CHART_INTENT_RE
    };

    if (typeof window !== 'undefined' && window) {
        App.ChatV2.routingConstants = api;
    }
    if (typeof module !== 'undefined' && module.exports) {
        module.exports = api;
    }

    (typeof window !== 'undefined' && window && (window.ChatLog || console) || console).info?.('CHAT-V2', '✅ routingConstants loaded');
})();
