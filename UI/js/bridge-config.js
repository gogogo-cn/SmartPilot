/**
 * ============================================================
 * bridge-config.js — Bridge Configuration
 * ============================================================
 * 功能：桥接配置模块，集中管理 Local Bridge 与 UI 服务器的连接地址、端口与协议配置。
 *
 * 核心类/函数：
 *   - getBridgeUrl() — 获取桥接地址
 *   - getUiServerUrl() — 获取 UI 服务器地址
 *   - loadBridgeConfig() — 加载桥接配置
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    const DEFAULT_BRIDGE_HOST = 'localhost';
    const DEFAULT_BRIDGE_PORT = 9527;
    const DEFAULT_BRIDGE_BASE = `http://${DEFAULT_BRIDGE_HOST}:${DEFAULT_BRIDGE_PORT}`;

    const DEFAULT_UI_HOST = 'localhost';
    const DEFAULT_UI_PORT = 8088;
    const DEFAULT_UI_BASE = `http://${DEFAULT_UI_HOST}:${DEFAULT_UI_PORT}`;

    /**
     * Excel Bridge 配置对象
     * 提供统一的服务地址和超时参数
     */
    window.ExcelBridgeConfig = window.ExcelBridgeConfig || {
        /** 服务基础地址 */
        baseUrl: DEFAULT_BRIDGE_BASE,
        /** UI服务基础地址 */
        uiBaseUrl: DEFAULT_UI_BASE,

        /** Excel Bridge API 主端点 */
        bridgeUrl: `${DEFAULT_BRIDGE_BASE}/api/excel`,
        /** 
         * 备选 Bridge API 端点列表（用于连接重试）
         * 按优先级排序，第一个为主地址
         */
        bridgeUrls: [
            `${DEFAULT_BRIDGE_BASE}/api/excel`,
            `http://127.0.0.1:${DEFAULT_BRIDGE_PORT}/api/excel`
        ],
        /** 默认 Bridge API 端点 */
        defaultBridgeUrl: `${DEFAULT_BRIDGE_BASE}/api/excel`,

        /** 共享 Skills 列表 API 端点 */
        skillsUrl: `${DEFAULT_BRIDGE_BASE}/api/skills/shared`,

        /**
         * 获取指定 Skill 的 API 地址：
         * @param {string} name - Skill 名称
         * @returns {string} Skill API 地址
         */
        skillUrl: function(name) {
            return `${DEFAULT_BRIDGE_BASE}/api/skills/shared/${encodeURIComponent(name)}`;
        },

        /** 状态检测超时时间（毫秒） */
        statusTimeout: 2000,
        /** 连接超时时间（毫秒） */
        connectTimeout: 5000,
        /** 调用超时时间（毫秒） */
        invokeTimeout: 30000,

        /**
         * 添加缓存穿透参数到 URL
         * 防止浏览器缓存旧版本资源
         * @param {string} url - 原始 URL
         * @returns {string} 添加时间戳后的 URL
         */
        withCacheBuster: function(url) {
            const separator = url.indexOf('?') >= 0 ? '&' : '?';
            return `${url}${separator}t=${Date.now()}`;
        }
    };

    /**
     * 路由配置对象
     * 管理 Skill 路由与浏览器展示相关配置
     */
    window.RoutingConfig = window.RoutingConfig || {
        /** 
         * 浏览器展示关键词正则
         * 命中时直接跳转到浏览器展示模式
         */
        browserKeywords: /浏览器|html.*展示|html.*列出|html.*显示|html.*查看|网页.*报表|交互图表|dashboard|大屏|echarts|用浏览器/,
        
        /** 简单 HTML 展示意图检测配置 */
        simpleHtmlDisplay: {
            requiredKeywords: ['html', '浏览器', '网页', 'browser', 'web'],
            actionKeywords: ['列出', '展示', '显示', '查看', '读取', '表格', '数据'],
            excludeKeywords: ['如何', '怎么', '为什么', '吗', '？', '?']
        },

        /** 数据分析类 Skill 列表（Hybrid 模式专用） */
        dataAnalysisSkills: ['DescriptiveStatsReporter', 'SmartDataAnalyzer', 'DataDiffAnalyzer'],

        /** SkillCreator 触发关键词 */
        skillCreator: {
            autoGenKeywords: ['创建', '生成', '新建', '做一个', '帮我写', '自定义'],
            queryKeywords: ['有多少', '多少', '列表', '数量', '几个', '几个技能', '查看', '显示']
        },

        /** 文档附件扩展名列表（命中时跳过 Skill 匹配） */
        docExtensions: ['xls', 'xlsx', 'doc', 'docx', 'pdf', 'ppt', 'pptx', 'txt', 'csv'],

        /**
         * 判断是否为简单 HTML 展示意图
         * @param {string} message - 用户消息
         * @returns {boolean} 是否为简单 HTML 展示意图
         */
        isSimpleHtmlDisplay: function(message) {
            const lowerMsg = message.toLowerCase();
            const hasRequired = this.simpleHtmlDisplay.requiredKeywords.some(kw => lowerMsg.includes(kw));
            const hasAction = this.simpleHtmlDisplay.actionKeywords.some(kw => lowerMsg.includes(kw));
            const hasExclude = this.simpleHtmlDisplay.excludeKeywords.some(kw => lowerMsg.includes(kw));
            return hasRequired && hasAction && !hasExclude;
        },

        /**
         * 判断是否命中浏览器展示关键词
         * @param {string} message - 用户消息
         * @returns {boolean} 是否命中
         */
        matchBrowserKeywords: function(message) {
            return this.browserKeywords.test(message.toLowerCase());
        },

        /**
         * 判断是否为数据分析类 Skill
         * @param {string} skillName - Skill 名称
         * @returns {boolean} 是否为数据分析类 Skill
         */
        isDataAnalysisSkill: function(skillName) {
            return this.dataAnalysisSkills.includes(skillName);
        },

        /**
         * 判断是否为文档附件
         * @param {File} file - 文件对象
         * @returns {boolean} 是否为文档附件
         */
        isDocAttachment: function(file) {
            if (!file || !file.name) return false;
            const ext = file.name.toLowerCase().split('.').pop();
            return this.docExtensions.includes(ext);
        }
    };

    ChatLog.info('CONFIG', '[BridgeConfig] ✅ 已加载，Bridge地址:', window.ExcelBridgeConfig.bridgeUrl);
    ChatLog.info('CONFIG', '[RoutingConfig] ✅ 已加载，数据分析Skill数量:', window.RoutingConfig.dataAnalysisSkills.length);
})();



