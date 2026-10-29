/**
 * ============================================================
 * app-endpoints.js — 统一端点配置（唯一真源）
 * ============================================================
 * 职责：集中管理所有本地服务端口、API 地址、桥接 URL
 * 原则：
 *   - 所有硬编码 localhost:port 必须引用此文件
 *   - 支持通过 window.ExcelBridgeConfig 覆盖
 *   - 环境相关值从配置读取，禁止散落代码中
 *
 * 端口分配表：
 *   9527 - Excel Bridge HTTP API（Python 本地桥）
 *   9528 - Excel VSTO 内置 HTTP 服务（Pomodoro/Workbook 等）
 *   9529 - 独立任务规划服务（Standalone Agent）
 *   11434 - Ollama 默认端口
 *   6333 - Qdrant 向量数据库
 *   8080 - llama.cpp 默认端口
 *   5173 - 开发模式 Vite 端口（仅开发环境）
 *
 * 依赖：
 *   - 无（纯配置，零依赖）
 * ============================================================
 */

(function() {
    'use strict';

    // 如果已定义则跳过（防止重复加载）
    if (window.App && App.Endpoints) return;

    // ==================== 端口常量 ====================
    var PORTS = {
        EXCEL_BRIDGE: 9527,
        EXCEL_VSTO_HTTP: 9528,
        STANDALONE_AGENT: 9529,
        LICENSE_API: 9530,
        OLLAMA: 11434,
        QDRANT: 6333,
        LLAMACPP: 8080,
        DEV_VITE: 5173
    };

    // ==================== 基础 URL 构造器 ====================
    function buildUrl(port, path) {
        return 'http://localhost:' + port + (path || '');
    }

    // ==================== 端点配置对象 ====================
    var Endpoints = {
        // Excel Bridge（Python 本地桥）
        excelBridge: {
            baseUrl: buildUrl(PORTS.EXCEL_BRIDGE, '/api/excel'),
            invoke: buildUrl(PORTS.EXCEL_BRIDGE, '/api/excel/invoke'),
            skills: function(skillName) {
                return buildUrl(PORTS.EXCEL_BRIDGE, '/api/skills/shared/' + encodeURIComponent(skillName));
            },
            skillsList: buildUrl(PORTS.EXCEL_BRIDGE, '/api/skills/shared')
        },

        // Excel VSTO 内置 HTTP 服务
        excelVstoHttp: {
            baseUrl: buildUrl(PORTS.EXCEL_VSTO_HTTP),
            status: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/status'),
            pomodoro: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/pomodoro'),
            pomodoroGet: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/pomodoro/get'),
            standaloneStart: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/standalone/start'),
            llmExportConfig: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/llm/export-config'),
            hostWorkbook: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/host/workbook'),
            browseExcelFile: buildUrl(PORTS.EXCEL_VSTO_HTTP, '/api/browse/excel-file')
        },

        // 独立任务规划服务
        standaloneAgent: {
            baseUrl: buildUrl(PORTS.STANDALONE_AGENT),
            tasks: buildUrl(PORTS.STANDALONE_AGENT, '/tasks'),
            health: buildUrl(PORTS.STANDALONE_AGENT, '/health'),
            taskPlanner: buildUrl(PORTS.STANDALONE_AGENT, '/task-planner'),
            reloadConfig: buildUrl(PORTS.STANDALONE_AGENT, '/reload-config')
        },

        // 席位订阅授权服务（BYOK：不计量 Token）
        license: {
            baseUrl: buildUrl(PORTS.LICENSE_API),
            health: buildUrl(PORTS.LICENSE_API, '/health'),
            entitlement: buildUrl(PORTS.LICENSE_API, '/me/entitlement')
        },

        // AI 模型服务
        aiModels: {
            ollama: buildUrl(PORTS.OLLAMA, '/api/chat'),
            ollamaTags: buildUrl(PORTS.OLLAMA, '/api/tags'),
            llamacpp: buildUrl(PORTS.LLAMACPP, '/v1/models'),
            qdrant: buildUrl(PORTS.QDRANT),
            localBge: buildUrl(PORTS.OLLAMA)  // 默认复用 Ollama 端口
        },

        // 开发环境（仅开发模式使用）
        dev: {
            vite: buildUrl(PORTS.DEV_VITE, '/api/excel')
        }
    };

    // ==================== 智能获取 Bridge URL ====================
    /**
     * 获取当前环境有效的 Bridge 基础 URL
     * 优先级：window.ExcelBridgeConfig.baseUrl > 生产环境 9527 > 开发环境 5173
     * @returns {string}
     */
    function getBridgeBaseUrl() {
        if (window.ExcelBridgeConfig && window.ExcelBridgeConfig.baseUrl) {
            return window.ExcelBridgeConfig.baseUrl;
        }
        // 生产环境默认 9527
        return Endpoints.excelBridge.baseUrl;
    }

    /**
     * 获取当前环境有效的 Excel Bridge Invoke URL
     * @returns {string}
     */
    function getBridgeInvokeUrl() {
        var base = getBridgeBaseUrl();
        return base.replace(/\/api\/excel$/, '') + '/api/excel/invoke';
    }

    // ==================== 暴露到全局 ====================
    if (!window.App) window.App = {};
    App.Endpoints = Endpoints;
    App.Ports = PORTS;
    App.getBridgeBaseUrl = getBridgeBaseUrl;
    App.getBridgeInvokeUrl = getBridgeInvokeUrl;

    // 兼容旧版 window.ExcelBridgeConfig
    if (!window.ExcelBridgeConfig) {
        window.ExcelBridgeConfig = {
            get baseUrl() { return getBridgeBaseUrl(); },
            get bridgeUrl() { return getBridgeBaseUrl(); }
        };
    }

    if (typeof ChatLog !== 'undefined') {
        ChatLog.info('ENDPOINTS', '✅ 统一端点配置已加载 (Ports: ' + Object.keys(PORTS).length + ')');
    }

})();