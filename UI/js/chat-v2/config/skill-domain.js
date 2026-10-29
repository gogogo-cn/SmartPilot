/**
 * ============================================================
 * chat-v2/config/skill-domain.js — 确定性 Skill 域唯一事实源（声明式描述符）
 * ============================================================
 * 职责：用一份声明式数据定义"内置确定性 Skill 域"（美化/格式化/图表/清洗/CRUD）
 *       的类型、触发关键词、受护 MCP 工具、技能名识别正则。所有外部路由守卫
 *       一律从这里派生，禁止各文件再私写一套正则（历史做法导致每次修 bug 都是
 *       再补一个正则/守卫的补丁式堆叠）。
 *
 * 挂载：window.App.ChatV2.skillDomain
 *
 * 消费方（全部改为惰性调用本注册表，消除重复正则）：
 *   - chat-orchestrator.js   : deterministicKwTest + SKILL_DOMAIN_RE → isDomainMessage()
 *   - chat-v2/services/mcp-dispatcher.js : _isSkillGuardedTool 名单 → isGuardedTool()
 *   - chat-v2/services/intent-analyzer.js: isDeterministicSkill 正则 → isDomainName()
 *   - pipelines/skill/skill-constants.js : SKILL_TYPE_RULES/detectSkillType → detectTypeByName()
 *
 * 新增一个确定性 Skill 域时，只需在此加一个 descriptor；四个守卫自动跟随。
 * ============================================================
 */
(function () {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    // ============================================================
    // Skill Capability 声明（正向"我能做什么"，取代反向"别让我处理XX"补丁）
    // ------------------------------------------------------------
    // 每个 Skill 域/独立 Skill 通过 capabilities 自描述能力边界。
    // Route Validator 据此自动判定：skill 的能力能不能覆盖用户消息里的意图需求。
    // 不覆盖 → 驳回，丢给 text-pipeline。**新增 skill 只需在此加声明**，
    // 校验器自动工作，永远不需要写"别让我处理 XX"的补丁正则。
    //
    // capabilities 字段语义：
    //   canAggregate        : 能做聚合/分组/统计（sum/count/avg/按X分组）
    //   canVisualize        : 能生成图表/可视化
    //   canCreateNewSheet   : 能新建工作表并写入
    //   canFormat           : 能格式化/美化单元格样式
    //   canClean            : 能清洗数据（去重/填空/拆分列）
    //   canEditExcel        : 能修改 Excel 内容（写入/CRUD）
    //   canOutputHtml       : 能产出 HTML / 浏览器展示
    //   pureDisplayOnly     : 只展示已有数据（读→展示），不做任何计算/修改。
    //                        这是 HtmlBrowserDisplay 类 skill 的核心边界标记——
    //                        用户要"生成看板/图表/聚合"时一律被 validator 拦下。
    //   inputContextOptional: 选区不是硬依赖（蓝图引擎能从活动表兜底读取）
    // ============================================================

    // ============================================================
    // 通用意图门（一次声明，所有 DESCRIPTORS 内的 Skill 自动生效）
    // ------------------------------------------------------------
    // 写一次正则，新增 skill 不用再手动配 intent 守卫：
    //   GENERIC_QUERY_INTENT   纯查询问句拦截："什么是 X / X 原理 / X 教程"
    //   GENERIC_CREATE_VERBS   生成/执行型动词白名单（skillWords 由每个 DESCRIPTOR 自身 keywords 提供）
    // allowQueryIntent: true   显式声明"我能回答查询"的 skill 才放行（如 SkillCreator）
    // ============================================================
    const GENERIC_QUERY_INTENT = '什么是|什么叫|怎么|为什么|如何|区别|对比|介绍|讲解|解释|原理|教程|是什么|有什么用|干嘛的|做什么的|好不好|有多少';
    const GENERIC_CREATE_VERBS = '做一个|做个|做|创建|生成|制作|设计|搭|搭建|弄个|整个|搞个|帮我做|帮我建|建|出|画|新建|加一个|添加|弄|帮我';
    // 疑问句语气词（**不锚定句首**）——给下面放宽后的意图豁免兜底。
    // GENERIC_QUERY_INTENT 只锚定句首，拦得住"什么是X"，拦不住疑问在句中的问法
    // （"对比和合并有什么区别""清洗怎么用"）。放宽豁免时必须用它二次把关，
    // 否则命中强关键词的问句会被误当成执行指令。
    const QUESTION_TONE_RE = /(吗|呢|什么|怎么|怎样|如何|为什么|为啥|区别|能否|能不能|\?|？)/;

    const DESCRIPTORS = [
        {
            id: 'formatter', names: /SmartTableFormatter|Formatter/i,
            keywords: [
                '格式化', '格式整理', '统一日期', '统一数字', '金额格式', '日期格式', '数字格式', '百分比格式',
                '加粗', '粗体', '倾斜', '斜体', '水平居中', '垂直居中', '居中', '字体颜色', '字体大小', '字号',
                '填充色', '背景色', '底色', '行高', '列宽', '自动换行', '换行', '边框', '高亮', '条件格式',
                '清除格式', '合并单元格', '取消合并', '单元格格式'
            ],
            guardedTools: ['excel_format_range'], planLabel: '格式方案',
            capabilities: { canFormat: true }
        },
        {
            id: 'beautifier', names: /SmartTableBeautifier|Beautifier|美化/i,
            keywords: ['美化', '主题', '样式', '表格边框', '统一字体', '自动调整', '冻结首行', '配色', '表头样式'],
            guardedTools: ['excel_format_range'], planLabel: '样式方案',
            capabilities: { canFormat: true }
        },
        {
            id: 'chart', names: /SmartChartWizard|Chart/i,
            keywords: ['图表', '画图', '做图', '可视化', '柱状图', '柱状', '条形', '条形图', '折线图', '折线', '饼图', '面积图', '散点图', '雷达图', '组合图', '漏斗图', '瀑布图', '箱线图', '树状图', '生成图表', '画图表', '创建图表', '图表推荐', '同比', '导出图片'],
            guardedTools: ['excel_create_chart', 'excel_aggregate'], planLabel: '图表方案',
            capabilities: { canAggregate: true, canVisualize: true }
        },
        {
            id: 'cleaner', names: /SmartDataCleaner|Cleaner|数据清洗/i,
            keywords: ['清洗', '数据清洗', '去重', '移除重复', '重复行', '拆分列', '合并列', '合并单元格', '空值', '删空行', '数据质量', '脏数据', '填充空值', '去除空格', '文本规范化', '数据透视',
                '拆分', '拆分数据', '拆分sheet', '拆分表', '按.*拆分', '拆分到不同', '拆分到多个', '数据拆分', '分割数据', '数据分割', '分拆', '数据分拆',
                '合并', '合并工作表', '合并数据', '数据合并', '合到一张表', '合表', '合并多表', '合并所有sheet', '汇总数据', '汇总表', '合到一张', '合并sheet', '合并工作簿', '合并文件', '汇总文件'],
            guardedTools: ['excel_clean_data', 'excel_merge_cells', 'excel_merge_sheets', 'excel_merge_workbooks',
                'excel_split_by_column', 'excel_split_workbook', 'excel_split_by_rows'], planLabel: '清洗方案',
            capabilities: { canClean: true, canEditExcel: true, canCreateNewSheet: true }
        },
        {
            id: 'crud', names: /ExcelCRUD|Crud|CRUD|增删改查/i,
            keywords: ['增删改查', 'CRUD'],
            guardedTools: ['excel_crud_create', 'excel_crud_read', 'excel_crud_update', 'excel_crud_delete'], planLabel: '增删改查方案',
            capabilities: { canEditExcel: true }
        },
        {
            id: 'dashboard', names: /DashboardGenerator|Dashboard|看板|驾驶舱/i,
            keywords: ['看板', '仪表盘', '大屏', '数据大屏', '驾驶舱', '数据驾驶舱', 'dashboard', '数据面板', '经营看板', '销售看板', '监控大屏'],
            guardedTools: ['excel_generate_dashboard', 'excel_create_chart', 'excel_aggregate', 'excel_format_range', 'excel_write_range', 'excel_create_sheet', 'excel_delete_sheet', 'excel_merge_cells', 'excel_create_slicer', 'excel_display_html_in_browser'],
            planLabel: '看板方案',
            capabilities: { canAggregate: true, canVisualize: true, canCreateNewSheet: true, canFormat: true, canOutputHtml: true, inputContextOptional: true }
        }
    ];

    // CAPABILITY_OVERRIDES：不在 DESCRIPTORS 里的独立 Skill（如 HtmlBrowserDisplay、SkillCreator）。
    // key 按 skill.name 匹配。任何 Skill 都必须能过能力校验才能上路。
    const CAPABILITY_OVERRIDES = {
        HtmlBrowserDisplay: {
            pureDisplayOnly: true, canAggregate: false, canVisualize: false,
            canCreateNewSheet: false, canOutputHtml: true
        },
        SkillCreator: { canAggregate: false, canVisualize: false, canCreateNewSheet: false },
        // 公式助手：唯一声明 canFormula 的技能（needsFormula 校验据此放行）
        SmartFormulaAssistant: { canFormula: true },
        // ── 数据处理类 Skill 能力声明（2026-10-09 补齐）──
        // 背景：getSkillCapabilities 对"既不属 DESCRIPTORS、也不在 CAPABILITY_OVERRIDES"的 Skill
        // 兜底返回 {}，于是 validateSkillRoute 的 `needsEditExcel && !canEditExcel && !canClean
        // && !canFormat` 分支把它们一律驳回。实测日志：
        //   "合并多个表"/"按部门拆分成不同Sheet" 已通过意图守卫命中 DataMerger/DataSplitter，
        //   却输出「被 capability 校验驳回: 用户要修改 Excel，但 Skill 没有任何编辑能力 → 丢弃」。
        // 这些 Skill 的 md 自述明确会改写工作表/生成结果表，故按自述**正向声明**：会写才声明，
        // 不夸大（不声明 canAggregate/canVisualize），也不缩小。
        // 反例：SmartSummary 自述"输出文本描述"、不改动 Excel，不能因它路由失败就虚报编辑能力
        // （它的卡点在意图门，见 GENERIC_QUERY_INTENT / 生成动词判定）。
        DataMerger: { canEditExcel: true, canCreateNewSheet: true },
        DataSplitter: { canEditExcel: true, canCreateNewSheet: true },
        TableJoiner: { canEditExcel: true, canCreateNewSheet: true },
        // 自述"标记新增、删除、修改的单元格，生成差异摘要报告"
        DataDiffAnalyzer: { canEditExcel: true, canCreateNewSheet: true },
        // 自述"输出与原始数据行对齐的结果并支持一键插入 Excel"（pipeline: python）
        DataNormalizer: { canEditExcel: true }
    };

    // 惰性构建索引（模块级只算一次）
    let _byId = null;
    let _searchable = null;   // [{pos, token}] 便于统一做前缀/包含匹配，简单起见直接用全词 includes
    function ensureIndex() {
        if (_byId) return;
        _byId = {};
        for (const d of DESCRIPTORS) _byId[d.id] = d;
    }

    const api = {
        /** 全部描述符 */
        getAll() { ensureIndex(); return DESCRIPTORS.slice(); },

        /**
         * 【停用感知】某个"真实 Skill 名"是否被用户显式停用。
         *
         * 背景（2026-10-01 实测）：DESCRIPTORS 是前端静态领域目录，完全不感知 isEnabled。
         * 用户停用 SmartTableBeautifier 后，后端 GetAvailableSkills 不再返回它，反而让路由层的
         * "后端没返回 → 注入 synthetic 空壳"分支把它复活（日志对照：
         *   SmartTableBeautifier(kw=27) 停用后变成 beautifier(kw=9)，9 个正是 DESCRIPTORS 的关键词），
         * 同时 LLM 意图分析目录里也照旧有 beautifier，于是出现"我把 Skill 关了，怎么还在用/还显示
         * 路由到 beautifier（置信度 1）"。
         *
         * 判据来源（任一可用即可；都不可用则返回 false = 保持旧行为，交由执行层守卫 C 兜底）：
         *  1) App.allSkillsFull      —— 管理面板拉取的完整清单（含停用项 + isEnabled）
         *  2) App.disabledSkillNames —— 面板切换开关时维护的停用名单（会话级）
         */
        isSkillNameDisabledByUser(name) {
            try {
                const n = String(name || '').trim();
                if (!n) return false;
                const all = Array.isArray(App.allSkillsFull) ? App.allSkillsFull : null;
                if (all && all.length > 0) {
                    const matched = all.filter(s => {
                        const sn = String((s && (s.name || s.Name)) || '');
                        return sn && sn.toLowerCase() === n.toLowerCase();
                    });
                    if (matched.length > 0) {
                        return matched.every(s => s.isEnabled === false || s.IsEnabled === false);
                    }
                }
                const off = Array.isArray(App.disabledSkillNames) ? App.disabledSkillNames : [];
                return off.some(x => String(x || '').toLowerCase() === n.toLowerCase());
            } catch (_e) { return false; }
        },

        /** 某个内置领域（DESCRIPTOR）当前是否被用户整域停用（该域下的真实 Skill 全部停用）。 */
        isDomainDisabledByUser(desc) {
            try {
                if (!desc || !desc.names) return false;
                const all = Array.isArray(App.allSkillsFull) ? App.allSkillsFull : null;
                if (all && all.length > 0) {
                    const matched = all.filter(s => desc.names.test(String((s && (s.name || s.Name)) || '')));
                    if (matched.length > 0) {
                        return matched.every(s => s.isEnabled === false || s.IsEnabled === false);
                    }
                }
                const off = Array.isArray(App.disabledSkillNames) ? App.disabledSkillNames : [];
                if (off.length > 0) {
                    return off.some(x => desc.names.test(String(x || '')) || String(x || '') === desc.id);
                }
                return false;
            } catch (_e) { return false; }
        },

        /**
         * 【背书包袱】一个领域是否有"真实 Skill 背书"。
         *
         * 用途：getSkillCatalog 里把"既未被用户停用、也没有任何真实 Skill 与之对应"的内置领域剔除，
         * 否则 LLM 会把这种"幽灵领域"当可用 Skill 列出/路由（如本环境没装 formatter skill，
         * 目录里仍挂着 formatter），用户看到"匹配了 formatter"却与真实能力不符。
         *
         * 判据来源（保守优先）：
         *   - 后端完整清单 App.allSkillsFull 未就绪（空）→ 返回 true，不做剔除，避免误伤本可路由的能力；
         *   - 已就绪 → 只要存在一个 name 命中 desc.names 的 Skill（不论启用与否）即视为"有背书"。
         */
        isDomainBacked(desc) {
            try {
                if (!desc || !desc.names) return true;
                const all = Array.isArray(App.allSkillsFull) ? App.allSkillsFull : null;
                if (!all || all.length === 0) return true;   // 未就绪：保守保留
                return all.some(s => desc.names.test(String((s && (s.name || s.Name)) || '')));
            } catch (_e) { return true; }
        },

        /**
         * LLM 路由用的精简 skill 目录。
         * 给 analyzeNewIntent 的 prompt 注入，让 LLM 从这份目录里选最匹配的 skill。
         * 字段是 LLM 能理解的自然语言描述，不是内部正则。
         *
         * 【停用感知】被用户停用的领域/独立 Skill 不再注入——否则 LLM 会照旧回
         * route.skillId="beautifier"，UI 上显示"路由：skill → beautifier（置信度 1）"，
         * 用户以为停用无效。执行层虽有守卫 C 兜底丢弃空壳，但提示信息已经误导人。
         */
        getSkillCatalog() {
            const domainNames = DESCRIPTORS.map(d => d.names).concat(Object.keys(CAPABILITY_OVERRIDES).map(n => new RegExp(n, 'i')));
            const _liveDescs = DESCRIPTORS.filter(d => !api.isDomainDisabledByUser(d) && api.isDomainBacked(d));
            const skills = _liveDescs.map(d => ({
                id: d.id,
                name: d.id,
                description: {
                    formatter: '统一格式：日期/数字/金额/百分比格式规范、千分位、小数位',
                    beautifier: '美化表格：主题配色、边框、字体、冻结首行、自动调整列宽',
                    chart: '生成图表：柱状图/折线图/饼图等 20+ 类型，自动聚合数据，支持对比/同比',
                    cleaner: '清洗/合并/分拆数据：去重、删空行、合并多表到一张、按字段拆分数据到不同Sheet、拆分工作簿、填充空值、去除空格',
                    crud: '增删改查：读取/写入/插入/删除/清空单元格区域',
                    dashboard: '做看板/仪表盘：一键生成销售看板/经营驾驶舱，自动选 KPI、多图表、明细表，输出 Excel + HTML 大屏'
                }[d.id] || d.id,
                keywords: (d.keywords || []).slice(0, 8).join('、'),
                capabilities: Object.keys(d.capabilities || {}).join(', ')
            }));
            if (!api.isSkillNameDisabledByUser('HtmlBrowserDisplay')) {
                skills.push({ id: 'html_report', name: 'HtmlBrowserDisplay', description: '生成独立 HTML 网页报表/可视化大屏，在浏览器中展示', keywords: 'html、网页、大屏、web报表', capabilities: 'canOutputHtml, pureDisplayOnly' });
            }
            if (!api.isSkillNameDisabledByUser('SkillCreator')) {
                // 【防抢路由】描述必须把边界写死，否则 LLM 看到用户说"创建一个XX"就照抄本条目。
                // 实测（2026-10-03）："为火车票创建一个解析器"被路由到 skill_creator（置信度 0.8），
                // 而该别名不在 DESCRIPTORS、也不是真实技能名、且过不了本技能的意图守卫
                // （守卫要求消息里出现 skill/技能/插件/智能体）→ 执行层丢弃，界面却还写着它。
                skills.push({
                    id: 'skill_creator', name: 'SkillCreator',
                    description: '创建/开发一个新的智能体 Skill（仅当用户明确提到 skill/技能/插件/智能体；' +
                        '"创建一个解析器/解析脚本/提取脚本"不属于本技能 → 选 GenerateOCRParseScript）',
                    keywords: 'skill、技能、插件、智能体', capabilities: ''
                });
            }
            skills.push({ id: 'free_chat', name: null, description: '不属于以上任何 Skill 的请求：通用问答、闲聊、多步骤复合操作、Python 代码生成、PPT 制作等', keywords: '', capabilities: '' });
            return skills;
        },

        /**
         * 【LLM 工具选择 · 动态全量目录】
         * 在 getSkillCatalog（静态领域锚点）基础上，追加后端实际加载的全部 skill
         * （含用户自定义 CustomSkills 与扩展 skill），让 LLM 的路由选择覆盖整个
         * 能力面，而不只是内置 9 个领域。
         *
         * 规则：
         *  - 内置领域条目优先（id 与动态 skill 名冲突时保留内置）；
         *  - 动态条目 id = skill.name（后端 GetAllSkills 的 Name 字段）；
         *  - 仅注入精简字段（id/description/keywords），并设数量上限控制 token；
         *  - 空/未就绪时退化为纯内置目录（保持旧行为）。
         */
        getDynamicSkillCatalog(cachedSkills) {
            const base = this.getSkillCatalog();
            const known = new Set(base.map(s => String(s.id || s.name || '').toLowerCase()));
            const src = Array.isArray(cachedSkills) ? cachedSkills : [];
            const dynamic = [];
            for (const s of src) {
                const name = String(s.name || s.Name || '').trim();
                if (!name) continue;
                if (known.has(name.toLowerCase())) continue;   // 内置领域已有，不重复
                const desc = String(s.description || s.Description || '');
                const kw = (Array.isArray(s.keywords) ? s.keywords
                    : Array.isArray(s.Keywords) ? s.Keywords
                    : []);
                dynamic.push({
                    id: name,
                    name: name,
                    description: desc || name,
                    keywords: kw.slice(0, 8).join('、'),
                    capabilities: ''
                });
                if (dynamic.length >= 30) break;   // token 上限
            }
            if (dynamic.length === 0) return base;
            ChatLog.info?.('SKILL', `[skill-domain] 动态 skill 目录: 内置 ${base.length} + 动态 ${dynamic.length} = ${base.length + dynamic.length}`);
            return base.concat(dynamic);
        },

        /**
         * 按 cachedSkill 对象匹配 DESCRIPTOR —— 用在 candidates 循环里兜底取 keywords。
         * 因为后端 GetAllSkills 返回的 skill 可能不带 keywords 字段，
         * 而 DESCRIPTORS 是 keywords 的单一事实源。
         */
        findDescriptorBySkill(skill) {
            if (!skill) return null;
            const n = String(skill.name || skill.Name || '');
            for (const d of DESCRIPTORS) {
                if (d.names.test(n)) return d;
            }
            return null;
        },

        /**
         * 每 Skill 的路由策略（声明式，取代 orchestrator 里按 skill.name 堆叠的 inline if 安全网）。
         *   - intent  : 意图守卫。仅当 createVerbs+skillWords 同时命中、且非 query/blocked 才放行。
         *   - block   : 命中阻却正则即放行给 text-pipeline（返回 null）。
         *   - boostMinScore: 命中该 Skill（score>=此值）即进 strongMatch（原 SkillCreator 加成）。
         */
        ROUTING_POLICIES: {
            SkillCreator: {
                boostMinScore: 3,
                intent: {
                    createVerbs: '创建|生成|新建|做一个|做个|自定义|开发|搞个|整个|写一个|写个|帮我做|帮我写|新功能',
                    skillWords: '\\bskill\\b|技能|插件|智能体',
                    queryPrefixes: '^(什么是|如何|怎么|为什么|区别|对比|介绍|讲解|解释|列出|查看|显示|有多少|多少|数量)',
                    blockedWords: '删除|修改|更新|编辑|重命名|停用|启用|卸载'
                }
            },
            // OCR 解析脚本生成器（GenerateOCRParseScript）—— 与 SkillCreator 同属"生成器型元技能"：
            // 产物是一段解析脚本（写进 OCR 智能解析面板），不直接操作 Excel，所以必须显式声明策略：
            //   ① shouldRouteToSkill 的守卫循环要求 getRoutingPolicy(name) 给出策略才能放行，
            //      未声明时返回 null → matchesIntent=false → 候选被静默丢弃（该 Skill 永远轮不到）；
            //   ② 声明了 intent 后，orchestrator 会把它识别为元技能并跳过 capability 校验，
            //      否则消息里的 read/write 操作词会被判成"要改 Excel，而它没有编辑能力"而驳回。
            // 实测（2026-10-02 23:24）："为火车票创建一个解析器"因①丢失路由 → 落 text-pipeline，
            // 被 skill_SkillCreator（关键词含泛化的"创建一个"）抢走执行，用户看到"这都能选错技能？"
            GenerateOCRParseScript: {
                boostMinScore: 3,
                intent: {
                    createVerbs: '创建|生成|新建|做一个|做个|写一个|写个|设计|来个|要一个|帮我做|帮我写',
                    // 必须出现"解析"侧的字样，避免把泛化的"创建一个XX"也算进来
                    skillWords: '解析器|解析脚本|解析规则|提取脚本|提取规则|字段提取|ocr解析|ocr脚本'
                }
            },
            DashboardGenerator: { boostMinScore: 1 },
            // 美化技能最易被"表头/汇总行"字样误路由（实测：用户写 SUMIFS 公式的消息里
            // 有"先读表头"，LLM 路由到美化置信度 1）。命中公式诉求一律阻却，交回通用管道。
            SmartTableBeautifier: {
                block: '写入公式|添加公式|填入公式|插入公式|公式写|写.{0,8}公式|SUMIFS|SUMIF\\(|VLOOKUP|COUNTIF|SUMPRODUCT'
            },
            SmartFormulaAssistant: {
                block: '写入公式|在.*单元格.*写(入|上)|在.*单元格.*添加|在.*单元格.*填|公式写.*到|填入公式|插入公式'
            }
        },

        /** 取某 Skill 的路由策略（按技能名匹配 name/Name + DESCRIPTORS 正则兜底） */
        getRoutingPolicy(skillOrName) {
            const name = (typeof skillOrName === 'string') ? skillOrName
                : (skillOrName && (skillOrName.name || skillOrName.Name) || '');
            if (api.ROUTING_POLICIES[name]) return api.ROUTING_POLICIES[name];
            for (const d of DESCRIPTORS) {
                if (d.names.test(name)) {
                    for (const pk of Object.keys(api.ROUTING_POLICIES)) {
                        if (new RegExp(pk, 'i').test(name) || d.names.test(pk)) {
                            return api.ROUTING_POLICIES[pk];
                        }
                    }
                    // 确定性 Skill 域（如 SmartChartWizard/Beautifier/Cleaner）没有显式 ROUTING_POLICIES 条目时，
                    // 返回一个基于其 capabilities 的合成空策略——意图判定交给 _testPolicyIntent 的通用门
                    // （hasGenericVerb + hasStrongDomainIntent 强领域关键词豁免）。
                    // 修复：原先返回 null → matchesIntent 拿到 null 直接 false，导致所有无显式策略的
                    // DOMAIN skill 意图守卫永远拦截（"使用饼图按部门展示销售额"掉进 text-pipeline）。
                    return {};
                }
            }
            // 【无域背书 Skill 不得静默丢弃】（2026-10-09 修复）
            // 背景：DESCRIPTORS 只声明了 6 个确定性域（formatter/beautifier/chart/cleaner/crud/dashboard），
            // 而大量真实 Skill（DataDiffAnalyzer/DataMerger/DataSplitter/TableJoiner/SmartSummary…）
            // 不匹配任何 DESCRIPTOR。原先此处返回 null → matchesIntent 直接 false →
            // shouldRouteToSkill 的守卫循环把**每一个**候选判为"未通过意图守卫"并跳过，
            // 最终 return null 交给 LLM。实测（探针日志）：
            //   "按部门拆分成不同Sheet" → DataSplitter(score=9,long=true) strongMatch=true
            //   却输出「候选 DataSplitter 未通过意图守卫，跳过」→「全部候选均未通过意图守卫」。
            // 即这些 Skill 关键词打分再高也永远轮不到执行，只能掉进 text-pipeline 让小模型幻觉。
            // 修复：返回合成空策略 {}，让它们进入 _testPolicyIntent 的通用门
            // （查询拦截 + blockedWords + hasGenericVerb/hasStrongDomainIntent），
            // 由通用门按消息本身裁决，而非"没配策略 = 一律拒绝"。
            // 安全性：isSkillRoutingMessage 只遍历显式声明 intent 的 ROUTING_POLICIES（且要求 p.intent），
            // 不读 getRoutingPolicy 的返回值，故本改动**不会**放大 _skillDomainMsg、
            // 不会重新触发历史上"McpDispatcher 被整体跳过 → 确定性硬短路全部失效"的回归。
            return {};
        },

        /** 按类型 id 取描述符 */
        getById(id) { ensureIndex(); return _byId[id] || null; },

        /** 按 Skill 名正则判断类型（从 cachedSkills 按 name 判定） */
        detectTypeByName(name) {
            const n = String(name || '');
            for (const d of DESCRIPTORS) {
                if (d.names.test(n)) return d.id;
            }
            return null;
        },

        /** 技能名是否属于确定性 Skill 域（follow-up 继承 / intent-analyzer 用） */
        isDomainName(name) {
            return api.detectTypeByName(name) !== null;
        },

        /**
         * 返回消息命中的全部确定性 Skill 域 id（去重、按 DESCRIPTORS 声明顺序）。
         * 复合句（如"把表美化后再画个图""清洗完做张看板"）会命中多个域，
         * 调用方可据此识别"多意图"而不被首命中的单一域切走。只读，不产生副作用。
         */
        matchDomains(message) {
            if (!message || typeof message !== 'string') return [];
            const lower = message.toLowerCase();
            const hit = [];
            for (const d of DESCRIPTORS) {
                const hasKw = (d.keywords || []).some(kw => kw && lower.includes(kw.toLowerCase()));
                if (hasKw) hit.push(d.id);
            }
            return hit;
        },

        /**
         * 消息是否命中确定性 Skill 域关键词 —— 判定"该不该走 skill-pipeline、
         * 不该进 dispatcher 裸调短路 / 不该被 LLM 意图分析拦下"。
         * 与旧 deterministicKwTest / SKILL_DOMAIN_RE 的并集保持兼容，避免回退。
         */
        isDomainMessage(message) {
            return api.matchDomains(message).length > 0;
        },

        /**
         * 意图守卫：
         *   - 有 policy.intent（如 SkillCreator）→ 用显式规则（createVerbs + skillWords 双命中）
         *   - 没有 policy.intent（如 DashboardGenerator/Beautifier/Cleaner）→ 用通用门：
         *     ① 命中 GENERIC_CREATE_VERBS 中的任一生成动词
         *     ② 不是纯查询问句（GENERIC_QUERY_INTENT）
         *     ③ 没有 blockedWords 里的删除/修改词
         *   - policy.allowQueryIntent: true → 跳过查询拦截（默认 false）
         */
        _testPolicyIntent(policy, message, opts) {
            if (!policy || !message || typeof message !== 'string') return false;
            const lower = message.toLowerCase();

            if (policy.allowQueryIntent) {
                // 显式声明可回答查询 → 不走通用门，交给显式 intent（若有）或全放行
            } else {
                // 【动作词优先于查询词】（2026-10-09）
                // 句首命中的"查询词"若正好是**本技能自己的关键词**（典型：DataDiffAnalyzer 的"对比"），
                // 那它是动作而非追问。同一个"对比"有两种截然不同的意图：
                //   "对比这两个表"          → 该执行（该 Skill 的 md examples 首条就是它）
                //   "对比和合并有什么区别"  → 该回答
                // 二者由**句中疑问语气**区分：带疑问语气仍判查询并拦截，不带则视为动作放行。
                // opts.skillKeywords 默认不传 → 既有调用方行为完全不变（一律按查询拦截）。
                const _trimmed = message.trim();
                const _qm = _trimmed.match(new RegExp('^(?:' + GENERIC_QUERY_INTENT + ')'));
                if (_qm) {
                    const _kws = (opts && Array.isArray(opts.skillKeywords)) ? opts.skillKeywords : [];
                    const _qw = String(_qm[0] || '').toLowerCase();
                    const _isActionWord = !!_qw && _kws.some(k => typeof k === 'string' && k.toLowerCase() === _qw);
                    if (!_isActionWord || QUESTION_TONE_RE.test(_trimmed)) return false;
                }
            }

            const pl = policy.intent || {};
            const hasBlocked = !!pl.blockedWords && new RegExp(pl.blockedWords).test(lower);
            if (hasBlocked) return false;

            if (policy.intent) {
                const hasVerb = !!pl.createVerbs && new RegExp(pl.createVerbs).test(lower);
                const hasSkill = !!pl.skillWords && new RegExp(pl.skillWords).test(lower);
                return hasVerb && hasSkill;
            }

            const hasGenericVerb = new RegExp(GENERIC_CREATE_VERBS).test(lower);
            // 强领域意图豁免：消息命中"确定性 Skill 域"关键词时也视为生成意图放行守卫生成侧，
            // 避免"使用饼图按部门展示销售额"这类**明确的图表操作词**因没有通用生成动词
            // （创建/生成/做/画/建/出）而被意图守卫拦下 → 掉进 text-pipeline 让小模型幻觉（光说不做）。
            // 仅当命中 DESCRIPTORS 里带 canVisualize/canClean/canFormat 等执行能力的域关键词才放行，
            // 纯查询问句（"什么是饼图"）仍被上面的 isQuery 拦截，不会误放。
            const hasStrongDomainIntent = (() => {
                // 疑问句不带执行意图：命中域关键词的问句（实测"清洗数据是什么意思"）应留给
                // LLM 作答，而不是当成清洗指令执行。句首锚定的 isQuery 拦不住疑问在句中的问法，
                // 故在此用疑问语气兜底（与上面强关键词豁免用的是同一把尺子）。
                if (QUESTION_TONE_RE.test(lower)) return false;
                if (!api.getById) return false;
                for (const d of DESCRIPTORS) {
                    if (!d.keywords || d.keywords.length === 0) continue;
                    const caps = d.capabilities || {};
                    const executable = caps.canVisualize || caps.canClean || caps.canFormat || caps.canAggregate || caps.canEditExcel || caps.canCreateNewSheet;
                    if (!executable) continue;
                    if (d.keywords.some(kw => kw && lower.includes(kw.toLowerCase()))) {
                        return true;
                    }
                }
                return false;
            })();
            // 【强关键词豁免】（2026-10-09 新增，仅对无显式 intent 的技能生效）
            // 当候选 Skill 命中了**具体关键词**——5 字以上长词（"每1000行""按部门拆分"）、
            // 4 字中词（"跨表匹配"）或强意图 2 字词（"总结""对比"）——该关键词本身就是
            // 比"生成动词"更硬的意图证据：这类词几乎不会出现在闲聊里。
            // 而 hasGenericVerb/hasStrongDomainIntent 要求消息里额外出现"做/创建/生成"
            // 或 DESCRIPTORS 域关键词，导致明明命中专属关键词却被通用门挡掉：
            //   ① "每1000行拆一个文件" → DataSplitter 命中长词"每1000行"，但消息里既无
            //      生成动词、也不含"拆分"（hasStrongDomainIntent 落空）→ 未路由；
            //   ② "跨表匹配一下这两张表" → TableJoiner 命中中词"跨表匹配"，同理落空；
            //   ③ "总结这段内容" → SmartSummary 命中强 2 字词"总结"，且该技能只输出文本
            //      不改动 Excel（按自述不虚报编辑能力），hasStrongDomainIntent 必然落空。
            // 安全边界（三重，缺一不可）：
            //   a) 上面的**查询拦截**与 **blockedWords** 在此处之前生效，不会被绕过；
            //   b) 豁免**仅对非疑问句**生效：带疑问语气（吗/什么/区别/？…）不豁免，
            //      避免"清洗怎么用"这类问句被当成执行指令；
            //   c) opts 默认不传 → isSkillRoutingMessage 等既有调用方行为完全不变。
            if (opts && opts.strongKeyword && !QUESTION_TONE_RE.test(lower)) return true;

            return hasGenericVerb || hasStrongDomainIntent;
        },

        /**
         * 消息是否属于"必须放行给 skill-pipeline、不该进 dispatcher 裸调短路"的路由意图。
         * 覆盖两类：
         *   1. 确定性 Skill 域关键词命中（isDomainMessage，如"换成柱状图/继续美化"）；
         *   2. 声明了 intent 的路由策略命中（如 SkillCreator 的"创建xx skill"）。
         * 否则后端 MatchCapability 常把"帮我生成一个xx的 skill"误匹配成 computer_/读数据工具
         * 而先短路返回 not_found，技能创建根本走不到 shouldRouteToSkill。
         */
        isSkillRoutingMessage(message) {
            if (!message || typeof message !== 'string') return false;
            if (api.isDomainMessage(message)) return true;
            // 【只认显式声明了 intent 的策略】（SkillCreator / GenerateOCRParseScript）。
            // 为什么不能遍历全部策略：没有 intent 的策略（SmartTableBeautifier / SmartFormulaAssistant /
            // DashboardGenerator 只声明了 block / boostMinScore）会走 _testPolicyIntent 的通用门
            // （hasGenericVerb || hasStrongDomainIntent），而 GENERIC_CREATE_VERBS 里含"导出/生成/保存"
            // 这类普通动词 —— 于是任意一句"将当前文档导出为 PDF"都会命中，
            // 连锁导致 orchestrator 的 _skillDomainMsg=true → **McpDispatcher 被整体跳过**，
            // 所有确定性硬短路（导出/格式化/筛选）在 Word/PPT 下全部失效
            // （实测 2026-10-06 01:10：host_bridge.log 明确记录 "McpDispatcher 被跳过 … skillDomain=true"，
            //   导出请求因此落到 agent-loop，模型零工具调用却声称已完成）。
            for (const name of Object.keys(api.ROUTING_POLICIES)) {
                const p = api.ROUTING_POLICIES[name];
                if (!p || !p.intent) continue;
                if (api._testPolicyIntent(p, message)) return true;
            }
            return false;
        },

        /** 供 orchestrator.shouldRouteToSkill 复用统一意图守卫（单一来源，带 DESCRIPTORS 兜底） */
        matchesIntent(name, message, opts) {
            const policy = api.getRoutingPolicy(name);
            if (!policy) return false;
            return api._testPolicyIntent(policy, message, opts);
        },

        // ============================================================
        // Capability-Based Route Validator（核心防线，取代所有 block 补丁）
        // ------------------------------------------------------------
        // 解决"关键词打分后路由到了 skill，但这个 skill 根本干不了用户要的事"的问题。
        // 不是补丁式的"别让我处理 XX"，而是正向的"我能做什么"声明 + 自动校验。
        // 意图需求优先从 LLM operations 提取（LLM 动态理解，类似杰瑞VBE AI 看板
        // 的"动态生成+修正"思路），LLM 不可用时降级到正则。
        // ============================================================

        /**
         * 从 skill 对象解析 capabilities（能做什么）。
         * 优先级：CAPABILITY_OVERRIDES(按 name) → DESCRIPTORS(按 names 正则) → 兜底推断。
         */
        getSkillCapabilities(skill) {
            if (!skill) return {};
            const name = skill.name || skill.Name || '';
            if (CAPABILITY_OVERRIDES[name]) return Object.assign({}, CAPABILITY_OVERRIDES[name]);
            for (const d of DESCRIPTORS) {
                if (d.names.test(name)) return Object.assign({}, d.capabilities || {});
            }
            const pipeline = String(skill.pipeline || skill.Pipeline || '').toLowerCase();
            if (pipeline === 'html') return { pureDisplayOnly: true, canOutputHtml: true };
            return {};
        },

        /**
         * 从用户消息 + LLM operations 提取意图需求（用户要什么）。
         * LLM operations 优先（语义理解准），operations=null 时降级到正则。
         */
        extractIntentRequirements(message, llmOperations) {
            const m = String(message || '').toLowerCase();
            const ops = Array.isArray(llmOperations) ? llmOperations : null;
            // 正则信号（用于 LLM 兜底 + 补充 output 形式维度）
            const re = {
                wantsDashboard: /看板|仪表盘|驾驶舱|dashboard|数据大屏/.test(m),
                wantsHtmlBrowser: /浏览器|html|网页|投屏|大屏展示|在浏览器里|用浏览器/.test(m)
            };
            if (ops) {
                return {
                    // LLM 语义：aggregate + chart = 看板意图（核心标志是看板必须同时聚合数据+画图）
                    wantsDashboard: ops.includes('aggregate') && ops.includes('chart'),
                    // ★ LLM 常把"饼图/图表/展示销售额/折线图"泛化成 read/write 而漏抽 chart/aggregate
                    //   （本会话已现：operations=["read","analyze","write"]，但消息明显要画饼图）。
                    //   用正则对可视化/聚合补强——即使 LLM 漏抽，消息含明确图表词仍判可视化意图，
                    //   否则 _hasVisualIntent=false 会被 capability 校验误判为"修改 Excel 无编辑能力"而驳回。
                    needsVisualization: ops.includes('chart') ||
                        /图表|画图|做图|可视化|柱状|折线|饼|条形|散点|漏斗|瀑布|面积|雷达|柱形图|折线图|饼图|chart|组合图|主柱|次折|双轴|趋势图/.test(m),
                    needsAggregate: ops.includes('aggregate') ||
                        /分组|聚合|汇总|求和|求总|求平均|平均|统计|合计|总计|按.{0,6}结合|按.{0,6}分组|sum|count|avg|average|max|min|rank|排名|top\d+/.test(m),
                    // ★ 写公式是独立意图维度：美化/清洗/图表等技能均自述"不修改公式"，
                    //   消息含明确公式诉求时必须拒绝路由，交回通用管道用 excel_add_formula 完成
                    needsFormula: /写入公式|添加公式|填入公式|插入公式|公式写|写.{0,8}公式|SUMIFS|SUMIF\(|VLOOKUP|COUNTIF|SUMPRODUCT/.test(m),
                    needsEditExcel: ops.includes('format') || ops.includes('write') || ops.includes('delete') || ops.includes('copy'),
                    needsCreateNewSheet: ops.includes('create_sheet') || /新建.*表|新工作表|创建.*表/.test(m),
                    // output 形式正则兜底（LLM 不管"要在哪输出"，这是输出维度不是 capability 维度）
                    wantsHtmlBrowser: re.wantsHtmlBrowser
                };
            }
            // 正则兜底（LLM 超时/失败）
            return {
                wantsDashboard: re.wantsDashboard,
                needsVisualization: /图表|画图|做图|可视化|柱状|折线|饼|条形|散点|漏斗|瀑布|面积|雷达|柱形图|折线图|饼图|chart/.test(m),
                needsAggregate: /分组|聚合|汇总|求和|求总|求平均|平均|统计|合计|总计|按.*分组|sum|count|avg|average|max|min|rank|排名|top\d+/.test(m),
                needsFormula: /写入公式|添加公式|填入公式|插入公式|公式写|写.{0,8}公式|SUMIFS|SUMIF\(|VLOOKUP|COUNTIF|SUMPRODUCT/.test(m),
                needsEditExcel: /清洗|去重|合并|拆分|删除|移除|填充|写入|填入|插入|格式化|美化/.test(m),
                needsCreateNewSheet: /新建.*表|创建.*表|放到.*工作表|新工作表/.test(m),
                wantsHtmlBrowser: re.wantsHtmlBrowser
            };
        },

        /**
         * 核心校验：skill 的 capabilities 能不能覆盖用户的意图需求。
         * 返回 { ok: boolean, reason: string }。reason 用中文，便于打日志。
         *
         * 校验规则（正向覆盖，不是反向补丁）：
         *   1. pureDisplayOnly → 任何要计算/修改/聚合/可视化的意图都挡
         *   2. wantsDashboard → 必须 canAggregate + canVisualize
         *   3. needsAggregate → 必须 canAggregate
         *   4. needsVisualization → 必须 canVisualize
         *   5. needsCreateNewSheet → 必须 canCreateNewSheet
         *   6. needsEditExcel → 必须至少有一项编辑能力
         */
        validateSkillRoute(skill, message, llmOperations) {
            const caps = api.getSkillCapabilities(skill);
            const intent = api.extractIntentRequirements(message, llmOperations);

            if (caps.pureDisplayOnly) {
                if (intent.wantsDashboard)  return { ok: false, reason: '纯展示型 Skill 不能处理看板意图（需要聚合+图表）' };
                if (intent.needsAggregate)  return { ok: false, reason: '纯展示型 Skill 不能处理聚合/统计意图' };
                if (intent.needsVisualization) return { ok: false, reason: '纯展示型 Skill 不能处理图表/可视化意图' };
                if (intent.needsEditExcel)  return { ok: false, reason: '纯展示型 Skill 不能处理 Excel 修改意图' };
                if (intent.needsCreateNewSheet) return { ok: false, reason: '纯展示型 Skill 不能新建工作表' };
            }

            if (intent.wantsDashboard) {
                if (!caps.canAggregate) return { ok: false, reason: '看板需要聚合能力，但 Skill 自声明 canAggregate=false' };
                if (!caps.canVisualize) return { ok: false, reason: '看板需要可视化能力，但 Skill 自声明 canVisualize=false' };
            }

            if (intent.needsAggregate && !caps.canAggregate)
                return { ok: false, reason: '用户要聚合/分组/合计，但 Skill 不能聚合' };
            if (intent.needsFormula && !caps.canFormula)
                return { ok: false, reason: '用户要写公式，但该 Skill 不具备公式能力（应交回通用管道用 excel_add_formula 完成）' };
            if (intent.needsVisualization && !caps.canVisualize)
                return { ok: false, reason: '用户要图表/可视化，但 Skill 不能画图' };
            if (intent.needsCreateNewSheet && !caps.canCreateNewSheet)
                return { ok: false, reason: '用户要新建工作表，但 Skill 不能新建工作表' };
            if (intent.needsEditExcel && !caps.canEditExcel && !caps.canClean && !caps.canFormat) {
                // 关键：LLM operations 常把"画图/展示销售额"过度抽成含 write，导致 needsEditExcel=true。
                // 但图表/聚合技能用内存绑定（dataValues）不写工作表，本就不需要编辑能力——
                // 且此时 needsVisualization/needsAggregate 是用户更真实的核心意图。
                // 因此仅当消息**没有**明确的图表/聚合/看板意图（真·要改 Excel 内容）时，
                // 才以编辑能力不足驳回；否则放行给具备可视化/聚合能力的技能。
                const _hasVisualIntent = intent.needsVisualization || intent.needsAggregate || intent.wantsDashboard;
                if (!_hasVisualIntent) return { ok: false, reason: '用户要修改 Excel，但 Skill 没有任何编辑能力' };
            }

            return { ok: true };
        },

        /** 该 MCP 工具是否为"确定性 Skill 域受护工具"——dispatcher 不得裸调，
         * 命中必须放行给 skill-pipeline（对应旧 dispatcher 的 _isSkillGuardedTool 名单）。
         */
        isGuardedTool(toolName) {
            if (!toolName || typeof toolName !== 'string') return false;
            for (const d of DESCRIPTORS) {
                if (d.guardedTools.indexOf(toolName) >= 0) return true;
            }
            return false;
        },

        /** 某类型的所有触发关键词 */
        keywordsOf(id) {
            ensureIndex();
            const d = _byId[id];
            return d ? d.keywords.slice() : [];
        },

        /** 某类型的所有受护工具 */
        guardedToolsOf(id) {
            ensureIndex();
            const d = _byId[id];
            return d ? d.guardedTools.slice() : [];
        }
    };

    // ============================================================
    // STEPS_CONTRACT — 自定义 Skill 声明式工具序列唯一事实源
    // ------------------------------------------------------------
    // 产物契约（SkillCreator 生成）：
    //   steps: [
    //     { name, description, tool, strategy, params, code?, waitOn? }
    //   ]
    //   - tool     (必填) MCP 工具名（direct 步），前端经 toolRegistry.invokeTracked 执行；
    //              'skill.<name>' 用于委托给另一个 Skill；既有 'excel_*'/'computer_*' 工具名直接可用。
    //   - strategy (选填) 'direct'（默认，调 MCP 工具）| 'code'（执 VBA，走后端固定模式）
    //   - params   (可选) 传给工具的参数对象
    //   - code     (仅 strategy='code' 必填) 原始 VBA Sub；无 code 时报错（不产生不可审计黑盒）
    //   - waitOn   (保留) 预留：是否依赖上一步结果自动回填参数
    //   - name/description (可选) 步骤名/说明，缺省用 tool 名，供绿面板渲染
    // 设计原则：
    //   - 默认产物是【声明式工具序列】，每个步骤=真实 MCP 工具，前端可审计/可预览/可分步填色；
    //   - VBA 降为显式 strategy='code'，仅当某步确实需要任意逻辑时使用，绝不默认生成；
    //   - 渲染/执行/保存共用同一份校验，杜绝"产物与执行器各写一套"。
    // ============================================================
    const STEPS_STRATEGIES = ['direct', 'code'];

    /**
     * 取字段值：兼容 camelCase（steps 契约/前端产物）与 PascalCase（后端 SerializeSkillCompact、
     * md frontmatter 序列化出的 Steps）。
     * 为什么必须兼容：SkillCreator 卡片保存的数据来自后端序列化结果，steps 元素是
     * { Name, Tool, Strategy, Params, Code }（PascalCase）。只读 s.tool 会让每个步骤都被判成
     * "缺少 tool" → 保存被误拒（实测 2026-10-02 22:0x："步骤[0] 缺少 tool" 但卡片预览里 tool 明明在）。
     */
    function _pickStepField(obj, camel, pascal) {
        if (obj === null || typeof obj !== 'object') return undefined;
        if (obj[camel] !== undefined) return obj[camel];
        if (obj[pascal] !== undefined) return obj[pascal];
        return undefined;
    }

    function normalizeStep(step, index) {
        const s = step && typeof step === 'object' ? step : {};
        const toolRaw = _pickStepField(s, 'tool', 'Tool');
        const tool = typeof toolRaw === 'string' ? toolRaw.trim() : '';
        const strategyRaw = _pickStepField(s, 'strategy', 'Strategy');
        const strategy = (typeof strategyRaw === 'string' ? strategyRaw.trim() : '') || 'direct';
        const paramsRaw = _pickStepField(s, 'params', 'Params');
        const params = (paramsRaw && typeof paramsRaw === 'object' && !Array.isArray(paramsRaw)) ? paramsRaw : {};
        const nameRaw = _pickStepField(s, 'name', 'Name');
        const descRaw = _pickStepField(s, 'description', 'Description');
        const codeRaw = _pickStepField(s, 'code', 'Code');
        const waitRaw = _pickStepField(s, 'waitOn', 'WaitOn');
        return {
            name: (typeof nameRaw === 'string' && nameRaw.trim()) ? nameRaw.trim() : (tool || ('步骤 ' + (index + 1))),
            description: typeof descRaw === 'string' ? descRaw.trim() : '',
            tool,
            strategy: STEPS_STRATEGIES.indexOf(strategy) >= 0 ? strategy : 'direct',
            params,
            code: typeof codeRaw === 'string' ? codeRaw : '',
            waitOn: !!waitRaw
        };
    }

    api.STEPS_CONTRACT = {
        strategies: STEPS_STRATEGIES.slice(),
        displayStrategy: { direct: '工具', code: 'VBA' },

        /** 归一化 steps：数组、丢弃非法项，返回合法步骤列表（可安全执行） */
        normalizeAll(steps) {
            if (!Array.isArray(steps)) return [];
            return steps
                .map((s, i) => normalizeStep(s, i))
                .filter(s => s.tool || (s.strategy === 'code' && s.code));
        },

        /**
         * 校验 steps：返回 { ok, steps, errors[] }。
         * 非法项返回具体原因，便于入库前校验（PreStore）与执行前防线复用。
         */
        validateSteps(steps) {
            if (!Array.isArray(steps) || steps.length === 0) {
                return { ok: false, steps: [], errors: ['steps 必须是非空数组'] };
            }
            const errors = [];
            const out = [];
            steps.forEach((s, i) => {
                const n = normalizeStep(s, i);
                if (typeof s !== 'object' || !s) { errors.push(`步骤[${i}]: 必须是对象`); return; }
                if (!n.tool && !(n.strategy === 'code' && n.code)) {
                    errors.push(`步骤[${i}]: 缺少 tool，且非 code 步，无法识别执行目标`);
                }
                if (STEPS_STRATEGIES.indexOf(n.strategy) < 0) {
                    errors.push(`步骤[${i}]: 非法 strategy="${n.strategy}"（允许 direct|code）`);
                }
                if (n.strategy === 'code' && !n.code.trim()) {
                    errors.push(`步骤[${i}]: strategy='code' 但缺少 code（VBA），违反"不产生不可审计黑盒"约束`);
                }
                out.push(n);
            });
            return { ok: errors.length === 0, steps: out, errors };
        },

        /**
         * 校验 steps 里引用的工具名是否真实注册（入库前 PreStore 防线）。
         * "禁止臆造工具名"目前只是提示词约束，编造的 tool 直到执行期才报错；
         * 这里在保存前用 toolRegistry 逐条核对，并给出最接近的真实工具名建议。
         * 注：skill.<Name> 委托同样以工具形式注册在 registry 中，一并覆盖。
         * @param {Array} steps
         * @param {Object} [toolRegistry] 默认 App.ChatV2.toolRegistry（调用方需先 await init()）
         * @returns {{ok:boolean, skipped:boolean, unknown:Array<{index:number,tool:string,suggestion:string}>}}
         */
        validateTools(steps, toolRegistry) {
            const unknown = [];
            const reg = toolRegistry || (App.ChatV2 && App.ChatV2.toolRegistry);
            // registry 未就绪（如 Web/Mock 环境）→ 放行，不阻断保存
            if (!reg || typeof reg.get !== 'function') return { ok: true, skipped: true, unknown };

            const arr = Array.isArray(steps) ? steps : [];
            arr.forEach((s, i) => {
                const n = normalizeStep(s, i);
                if (n.strategy !== 'direct' || !n.tool) return;
                if (reg.get(n.tool)) return;
                let suggestion = '';
                try {
                    // 前缀剥离优先：模型常把"委托技能"的 skill. 当成工具前缀，写出
                    // skill.excel_read_range / skill.smtp_send_email 这类名字（skill. 只用于
                    // 委托已存在的技能，不是包住工具名）。剥离后能命中真实工具时优先给这条建议，
                    // 比 ngram 模糊匹配精准得多。
                    if (/^skill[._]/i.test(n.tool)) {
                        const stripped = n.tool.replace(/^skill[._]/i, '');
                        if (reg.get(stripped)) suggestion = stripped;
                    }
                    if (!suggestion && typeof reg._resolveToolAlias === 'function') {
                        suggestion = reg._resolveToolAlias(n.tool) || '';
                    }
                } catch (e) { /* 别名解析失败不影响校验结论 */ }
                unknown.push({ index: i, tool: n.tool, suggestion });
            });
            return { ok: unknown.length === 0, skipped: false, unknown };
        }
    };

    App.ChatV2.skillDomain = api;
})();