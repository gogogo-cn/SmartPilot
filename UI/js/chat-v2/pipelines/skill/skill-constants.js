/**
 * ============================================================
 * skill-constants.js — Skill 管道单一数据源（常量/注册表/关键词集合）
 * ------------------------------------------------------------
 * 职责：清洗操作注册表（顺序/标签/meta 字段）、派生工作表名、Skill 类型判定表、
 *       意图关键词集合、协议契约（CRUD 映射/后端错误文本模式）。
 * 挂载：App.ChatV2.skillPipelineModules.constants
 * 原则：所有"必须与 X 一致"的注释约束在此变成代码约束——
 *       buildSpecSteps 骨架顺序、executor 排序/标签/meta 累加、clean-crud 工具卡标签
 *       全部引用本模块，改一处即全局生效。
 * 注意：本模块为纯数据/纯函数，不依赖其它 skill 模块，必须最先加载。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.skillPipelineModules) App.ChatV2.skillPipelineModules = {};
    const MODS = App.ChatV2.skillPipelineModules;

    // ── 清洗操作注册表 ──
    // name    ：协议操作名（与后端 excel_clean_data / VB 契约一致，改动需同步后端）
    // label   ：通用中文标签（executor 规划文案 / clean-crud 工具卡名）
    // stepTitle：buildSpecSteps 骨架步骤标题
    // 数组顺序即骨架步骤顺序（原 buildSpecSteps 的 if 出现顺序，曾以注释人工对齐）
    const CLEAN_OPS = [
        { name: 'removeDuplicates', label: '去重',       stepTitle: '去除重复行' },
        { name: 'fillEmpty',        label: '填充空值',   stepTitle: '填充空值' },
        { name: 'trimSpaces',       label: '去除空格',   stepTitle: '去除首尾空格' },
        { name: 'normalizeDates',   label: '统一日期',   stepTitle: '统一日期格式' },
        { name: 'normalizeNumbers', label: '统一数字',   stepTitle: '统一数字格式' },
        { name: 'standardizeText',  label: '文本规范化', stepTitle: '文本规范化' },
        { name: 'removeEmptyRows',  label: '删除空行',   stepTitle: '删除空行' },
        { name: 'splitColumn',      label: '拆分列',     stepTitle: '拆分合并列' },
        { name: 'mergeColumns',     label: '合并列',     stepTitle: '合并指定列' },
        // —— 表结构变换（独立原子工具，非 excel_clean_data）——
        { name: 'splitByColumn',    label: '按字段拆表', stepTitle: '按字段值拆分数据到不同 Sheet' },
        { name: 'splitWorkbook',    label: '拆分工作簿', stepTitle: '拆分工作簿为独立 xlsx 文件' },
        { name: 'splitByRows',      label: '按行数拆表', stepTitle: '按固定行数拆分成多张 Sheet' },
        { name: 'mergeSheets',      label: '合并工作表', stepTitle: '合并多个工作表到一张' },
        { name: 'mergeWorkbooks',   label: '合并工作簿', stepTitle: '合并多个 xlsx 文件' }
    ];
    /** 操作名有序数组（executor 排序 plan.operations 用） */
    const CLEAN_OP_NAMES = CLEAN_OPS.map(o => o.name);
    const CLEAN_OP_BY_NAME = Object.fromEntries(CLEAN_OPS.map(o => [o.name, o]));
    /** 模糊清洗意图 → 自动扩展的全量清洗操作（顺序与 CLEAN_OPS 相对顺序一致） */
    const FULL_CLEAN_OPS = ['removeDuplicates', 'trimSpaces', 'normalizeDates', 'normalizeNumbers', 'removeEmptyRows'];
    function cleanOpLabel(name) {
        const def = CLEAN_OP_BY_NAME[name];
        return def ? def.label : undefined;
    }
    /** 后端返回的汇总 meta 字段契约（executor 逐操作累加） */
    const CLEAN_SUM_FIELDS = ['removedRows', 'deletedRows', 'markedRows', 'filledCells', 'trimmedCells', 'normalizedDates', 'normalizedNumbers', 'removedEmptyRows', 'splitColumns'];
    /** 后端返回的"取首次"字段契约 */
    const CLEAN_KEEP_FIRST_FIELDS = ['cleanedRows'];

    // ── CRUD 操作契约表（spec.operation → 中文标签 / MCP 工具名）──
    const CRUD_OPS = { read: '读取数据', write: '写入/修改', insert: '插入', delete: '删除', clear: '清空' };
    const CRUD_OP_TOOL = {
        read: 'excel_read_range',
        write: 'excel_set_cell_value',
        insertRow: 'excel_insert_row',
        insertAppend: 'excel_insert_rows',
        insertCol: 'excel_insert_column',
        deleteRow: 'excel_delete_row',
        deleteCol: 'excel_delete_column',
        clearRange: 'excel_clear_range',
        clearFormat: 'excel_clear_format'
    };
    function crudOpLabel(op) {
        return CRUD_OPS[op] || String(op || '');
    }

    // ── Skill 类型 → 规划/执行契约（原 executor 内的 if/else 分发链）──
    // promptBuilder / normalizer 是 skillSpecEngine 上的方法名，执行器延迟解析，
    // 保持本模块零依赖（不引用其它 skill 模块）。skeletonKind 供骨架标题生成。
    const SKILL_SPEC_CONTRACT = {
        beautifier: { planLabel: '样式方案',     promptBuilder: 'buildPlanningPrompt',     normalizer: 'normalizeSpec',       skeletonKind: 'beautify' },
        formatter:  { planLabel: '格式方案',     promptBuilder: 'buildFormatPlanningPrompt', normalizer: 'normalizeFormatSpec', skeletonKind: 'format' },
        chart:      { planLabel: '图表方案',     promptBuilder: 'buildChartPlanningPrompt',  normalizer: 'normalizeChartSpec',  skeletonKind: 'chart' },
        cleaner:    { planLabel: '清洗方案',     promptBuilder: 'buildCleanPlanningPrompt',  normalizer: 'normalizeCleanSpec',  skeletonKind: 'clean' },
        crud:       { planLabel: '增删改查方案', promptBuilder: 'buildCrudPlanningPrompt',   normalizer: 'normalizeCrudSpec',   skeletonKind: 'crud' },
        dashboard:  { planLabel: '看板方案',     promptBuilder: 'buildDashboardPlanningPrompt', normalizer: 'normalizeDashboardSpec', skeletonKind: 'dashboard' }
    };

    // ── 派生工作表名（与后端 ExcelMcpTools 落盘契约一致）──
    const SHEET_CHART_DATA = '图表数据源';
    const SHEET_AGGREGATE = '聚合结果';
    // 看板默认工作表名（蓝图未显式声明 sheet 时使用）
    const SHEET_DASHBOARD = '看板';

    // ── Skill 类型判定表 ──
    // 唯一事实源：声明式注册表 config/skill-domain.js 的 detectTypeByName()（names 正则 → 类型 id）。
    // 这里不再维护本地 SKILL_TYPE_RULES 副本，避免与注册表两处漂移。
    function detectSkillType(skillName) {
        const name = String(skillName || '');
        if (App.ChatV2 && App.ChatV2.skillDomain && typeof App.ChatV2.skillDomain.detectTypeByName === 'function') {
            return App.ChatV2.skillDomain.detectTypeByName(name);
        }
        return null;
    }

    // ── 清洗意图关键词 ──
    // 各变体差异是历史行为，刻意保留（合并会改变模糊请求/覆盖判定的结果）：
    //   CLEAN_INTENT_RE          整体清洗意图（"清洗/清理/优化/整理/数据质量/脏数据"）
    //   CLEAN_SPECIFIC_OP_RE     具体操作关键词（模糊请求判定；不含 拆分列/合并列）
    //   CLEAN_SPECIFIC_OP_EXT_RE 同上，另含 拆分列/合并列（纯意图覆盖判定用）
    //   CLEAN_EMPTY_ROW_CONFLICT_RE 空行锁定的冲突集（用"缺值"替"空白"，不含"空行"）
    const CLEAN_INTENT_RE = /(清洗|清理|优化|整理|数据质量|脏数据)/;
    const CLEAN_SPECIFIC_OP_RE = /(去重|重复|空值|填充|补空|去空格|空格|空白|trim|首尾|前后|统一日期|日期格式|统一数字|数字格式|金额|文本规范|文本标准化|大小写|删空行|空行)/;
    const CLEAN_SPECIFIC_OP_EXT_RE = /(去重|重复|空值|填充|补空|去空格|空格|空白|trim|首尾|前后|统一日期|日期格式|统一数字|数字格式|金额|文本规范|文本标准化|大小写|删空行|空行|拆分列|合并列)/;
    const CLEAN_EMPTY_ROW_RE = /(空行)/;
    const CLEAN_EMPTY_ROW_CONFLICT_RE = /(去重|重复|空值|填充|补空|缺值|去空格|空格|trim|首尾|前后|统一日期|日期格式|统一数字|数字格式|金额|文本规范|文本标准化|大小写|拆分列|合并列)/;

    // ── 图表/选区/聚合意图关键词 ──
    const USER_SELECTION_RE = /(选中|选区|选定|框选|所选)/;
    const REUSE_AGGREGATE_RE = /统计结果|上轮|刚才|这个|它|上面|之前|图表/;
    const AGGREGATE_REDIMENSION_RE = /(?:按|依照|根据)(?:部门|地区|产品|日期|人员|渠道|类型|月份|季度|年份)/;
    // 图表类型切换后续消息：仅"换/改成/改为/变成/切换 + 图表类型"，只变类型、沿用上一轮数据口径
    const CHART_TYPE_SWITCH_RE = /(?:换成|改成|改为|变成|转为|切换(?:为|到)?|换(?:一下|个)?)\s*(?:柱状|柱形|折线|饼|环形|圆环|面积|散点|条形|组合|雷达|漏斗|瀑布|旭日|树状|堆积|堆叠|仪表)[^。，！？]{0,8}图?/;

    // ── 后端错误/取值契约（前后端需同步维护）──
    // excel_aggregate 选区无效类错误文案
    const AGGREGATE_INVALID_RANGE_RE = /未选中|有效数据区域|至少需要表头|选区为空|no valid|selection.*empty|empty.*selection/i;
    // 非法工作表名取值（undefined/null/[object Object]/none/nan）
    const INVALID_SHEET_NAME_RE = /^(undefined|null|\[object Object\]|none|nan)$/i;
    function isValidSheetName(name) {
        return typeof name === 'string' && name.length > 0 && !INVALID_SHEET_NAME_RE.test(name);
    }
    // 无意义摘要过滤（executor 总结去重用）
    const GENERIC_SUMMARY_RE = /^(清洗完成|执行完成|操作完成|成功|完成|cleaned|done|success)$/i;

    // ── MCP 直连接口契约（绕过注册表的兜底调用）──
    const MCP_SERVER = 'excel-local';
    const MCP_CHART_TOOL_FQN = 'excel-local__excel_create_chart';

    // ── UI 兜底文案/图表元信息（集中定义便于统一口径）──
    const UI_DEFAULTS = {
        chartMeta: { name: '图表', icon: '📊', category: '通用图表' },
        chartTitle: '默认生成',
        dimHeader: '类别',
        valueHeader: '数量'
    };

    MODS.constants = {
        CLEAN_OPS,
        CLEAN_OP_NAMES,
        FULL_CLEAN_OPS,
        cleanOpLabel,
        CLEAN_SUM_FIELDS,
        CLEAN_KEEP_FIRST_FIELDS,
        CRUD_OPS,
        CRUD_OP_TOOL,
        crudOpLabel,
        SHEET_CHART_DATA,
        SHEET_AGGREGATE,
        SHEET_DASHBOARD,
        detectSkillType,
        SKILL_SPEC_CONTRACT,
        CLEAN_INTENT_RE,
        CLEAN_SPECIFIC_OP_RE,
        CLEAN_SPECIFIC_OP_EXT_RE,
        CLEAN_EMPTY_ROW_RE,
        CLEAN_EMPTY_ROW_CONFLICT_RE,
        USER_SELECTION_RE,
        REUSE_AGGREGATE_RE,
        AGGREGATE_REDIMENSION_RE,
        CHART_TYPE_SWITCH_RE,
        AGGREGATE_INVALID_RANGE_RE,
        INVALID_SHEET_NAME_RE,
        isValidSheetName,
        GENERIC_SUMMARY_RE,
        MCP_SERVER,
        MCP_CHART_TOOL_FQN,
        UI_DEFAULTS
    };
})();
