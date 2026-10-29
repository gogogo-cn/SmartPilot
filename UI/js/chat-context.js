/**
 * ============================================================
 * chat-context.js — Chat Context
 * ============================================================
 * 功能：聊天上下文模块，维护当前会话的消息上下文、角色信息与历史摘要。
 *
 * 核心类/函数：
 *   - getChatContext() — 获取聊天上下文
 *   - appendUserMessage() — 追加用户消息
 *   - appendAssistantMessage() — 追加助手消息
 *   - clearContext() — 清空上下文
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

class ChatContextManager {
    /** 单例实例 */
    static _instance = null;

    /**
     * 获取单例实例
     * @returns {ChatContextManager} 聊天上下文管理器实例
     */
    static getInstance() {
        if (!ChatContextManager._instance) {
            ChatContextManager._instance = window.chatContextManager || new ChatContextManager();
        }
        return ChatContextManager._instance;
    }

    /**
     * 构造函数
     * 初始化对话历史和会话信息
     */
    constructor() {
        /** 对话历史数组 */
        this.history = [];
        /** 最大保留消息数 */
        this.maxHistoryLength = 20;
        /** 当前会话信息 */
        this.currentSession = {
            id: this.generateSessionId(),
            startTime: new Date(),
            messageCount: 0
        };
    }

    /**
     * 构建请求上下文
     * @param {string} userMessage - 用户消息
     * @param {Object} [fileContext=null] - 文件上下文
     * @returns {Object} 请求上下文对象
     */
    buildContext(userMessage, fileContext = null) {
        return {
            currentMessage: userMessage,
            history: this.getRecentHistory(10),
            fileContext: fileContext,
            excelState: this.getExcelState(),
            sessionInfo: this.currentSession,
            timestamp: Date.now()
        };
    }

    /**
     * 获取最近的对话历史
     * @param {number} maxRounds - 最大轮数（每轮包含用户和助手消息）
     * @returns {Array} 最近的消息数组
     */
    getRecentHistory(maxRounds) {
        const maxMessages = maxRounds * 2;
        return this.history.slice(-maxMessages);
    }

    /**
     * 获取 Excel 状态
     * @returns {Object|null} Excel 状态对象或 null
     */
    getExcelState() {
        if (!this.isExcelBridgeAvailable()) {
            return null;
        }

        try {
            const bridge = window.chrome.webview.hostObjects.excelBridge;
            return {
                activeSheet: bridge.GetActiveSheet(),
                selectedRange: bridge.GetSelectedRange(),
                workbookName: bridge.GetWorkbookName()
            };
        } catch (error) {
            ChatLog.warn('CONTEXT', '[ChatContext] 获取Excel状态失败:', error);
            return null;
        }
    }

    /**
     * 保存消息到历史
     * @param {string} userMessage - 用户消息
     * @param {Object} thinkingResponse - AI思考响应对象
     */
    saveToHistory(userMessage, thinkingResponse) {
        // 保存用户消息
        this.history.push({
            role: 'user',
            content: userMessage,
            timestamp: new Date()
        });

        // 保存AI响应
        this.history.push({
            role: 'assistant',
            content: thinkingResponse.finalAnswer,
            thinkingSteps: thinkingResponse.thinkingSteps,
            actions: thinkingResponse.actions,
            mode: thinkingResponse.mode,
            timestamp: new Date()
        });

        // 更新会话计数
        this.currentSession.messageCount++;

        // 限制历史长度
        if (this.history.length > this.maxHistoryLength) {
            this.history = this.history.slice(-this.maxHistoryLength);
        }
    }

    /**
     * 添加单条消息到上下文（新架构兼容接口）
     * @param {Object} msg - 消息对象
     */
    addMessage(msg) {
        if (!msg) {
            ChatLog.warn('CONTEXT', '[ChatContext] addMessage: 消息为空');
            return;
        }
        const role = msg.role || 'unknown';
        const content = (msg.content !== undefined && msg.content !== null) ? String(msg.content) : '';
        this.history.push({
            role: role,
            content: content,
            timestamp: msg.timestamp || new Date(),
            images: msg.images || null,
            thinkingMode: msg.thinkingMode || null
        });
        this.currentSession.messageCount++;
        if (this.history.length > this.maxHistoryLength) {
            this.history = this.history.slice(-this.maxHistoryLength);
        }
    }

    /**
     * 获取用于后端的消息格式（新架构兼容接口）
     * @returns {Array} 标准化的消息数组
     */
    getMessagesForBackend() {
        return this.history.map(msg => ({
            role: msg.role,
            content: msg.content,
            timestamp: msg.timestamp
        }));
    }

    /**
     * 获取当前 Excel 状态（新架构兼容接口）
     * @returns {Object|null} Excel 状态对象或 null
     */
    getCurrentExcelState() {
        return this.getExcelState();
    }

    /**
     * 更新 Excel 状态（新架构兼容接口）
     * @param {Object} newState - 新的 Excel 状态
     */
    updateExcelState(newState) {
        if (!newState) return;
        if (!this._excelState) {
            this._excelState = {};
        }
        Object.assign(this._excelState, newState);
    }

    /**
     * 获取完整历史文本（用于调试或导出）
     * @returns {string} 格式化的历史文本
     */
    getFullHistoryText() {
        return this.history.map(msg => {
            const role = msg.role === 'user' ? '用户' : 'AI';
            return `${role}: ${msg.content}`;
        }).join('\n\n');
    }

    /**
     * 清空历史
     * 重置会话信息为新会话
     */
    clearHistory() {
        this.history = [];
        this.currentSession = {
            id: this.generateSessionId(),
            startTime: new Date(),
            messageCount: 0
        };
    }

    /**
     * 导出历史（用于保存）
     * @returns {Object} 导出数据对象
     */
    exportHistory() {
        return {
            session: this.currentSession,
            history: this.history,
            exportTime: new Date()
        };
    }

    /**
     * 导入历史（用于恢复）
     * @param {Object} data - 导入数据对象
     */
    importHistory(data) {
        if (data.session) {
            this.currentSession = data.session;
        }
        if (data.history) {
            this.history = data.history;
        }
    }

    /**
     * 获取统计信息
     * @returns {Object} 统计信息对象
     */
    getStats() {
        return {
            totalMessages: this.history.length,
            userMessages: this.history.filter(m => m.role === 'user').length,
            aiMessages: this.history.filter(m => m.role === 'assistant').length,
            sessionDuration: this.getSessionDuration(),
            sessionId: this.currentSession.id
        };
    }

    /**
     * 获取会话持续时间
     * @returns {string} 格式化的持续时间字符串
     */
    getSessionDuration() {
        const now = new Date();
        const diff = now - this.currentSession.startTime;
        const minutes = Math.floor(diff / 60000);
        const seconds = Math.floor((diff % 60000) / 1000);
        return `${minutes}分${seconds}秒`;
    }

    /**
     * 生成会话ID
     * @returns {string} 唯一会话ID
     */
    generateSessionId() {
        return 'session_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
    }

    /**
     * 检查 Excel Bridge 是否可用
     * @returns {boolean} 是否可用
     */
    isExcelBridgeAvailable() {
        return window.chrome && 
               window.chrome.webview && 
               window.chrome.webview.hostObjects && 
               window.chrome.webview.hostObjects.excelBridge;
    }
}

/** 创建全局实例 */
window.chatContextManager = new ChatContextManager();