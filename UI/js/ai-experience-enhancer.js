/**
 * ============================================================
 * ai-experience-enhancer.js — AI Experience Enhancer
 * ============================================================
 * 功能：AI 交互体验增强器，负责流式输出时的打字机效果、视觉反馈与用户体验优化。
 *
 * 核心类/函数：
 *   - enhanceStreamingExperience() — 增强流式输出体验
 *   - applyVisualEffects() — 应用视觉反馈效果
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    // ==================== AI状态管理器 ====================
    
    class AIStateManager {
        constructor() {
            this.currentState = 'idle'; // idle | thinking | executing | error | completed
            this.startTime = 0;
            this.timeout = 60000; // 60秒总超时
            this.watchdogTimer = null;
            this.retryCount = 0;
            this.maxRetries = 2;
        }

        /**
         * 开始AI处理流程
         */
        startProcessing() {
            this.currentState = 'thinking';
            this.startTime = Date.now();
            this.retryCount = 0;
            
            ChatLog.info('AI-State', '🚀 开始AI处理流程');
            
            // 启动看门狗定时器
            this._startWatchdog();
            
            // 更新UI状态
            this._updateUIStatus('processing', '正在思考中');
        }

        /**
         * 完成当前步骤
         * @param {string} stepName - 步骤名称
         */
        completeStep(stepName) {
            ChatLog.info('AI-State', `✅ 完成: ${stepName}`);
            
            // 重置看门狗
            this._resetWatchdog();
        }

        /**
         * 标记错误并尝试恢复
         * @param {Error} error - 错误对象
         * @param {string} context - 错误上下文
         */
        handleError(error, context) {
            const errorMsg = error?.message || error || '未知错误';
            
            ChatLog.error('AI-State', `❌ 错误 [${context}]: ${errorMsg}`);
            
            // 判断是否可恢复
            if (this._isRecoverableError(errorMsg) && this.retryCount < this.maxRetries) {
                this.retryCount++;
                ChatLog.warn('AI-State', `🔄 尝试恢复 (${this.retryCount}/${this.maxRetries}中`);
                
                // 显示恢复提示
                if (typeof App.showToast === 'function') {
                    App.showToast(`⚠️ 遇到小问题，正在自动恢复... (${this.retryCount}/${this.maxRetries})`, 'warning', 3000);
                }
                
                return true; // 可恢复
            } else {
                // 不可恢复或重试次数耗尽
                this.currentState = 'error';
                this._stopWatchdog();
                
                // 显示友好错误信息
                this._showUserFriendlyError(errorMsg, context);
                
                return false; // 不可恢复
            }
        }

        /**
         * 完成整个流程
         */
        complete() {
            this.currentState = 'completed';
            this._stopWatchdog();
            
            const duration = Date.now() - this.startTime;
            ChatLog.info('AI-State', `🎉 流程完成，耗时: ${(duration/1000).toFixed(1)}秒`);
            
            this._updateUIStatus('completed', '✅ 处理完成');
        }

        /**
         * 判断错误是否可恢复
         */
        _isRecoverableError(errorMsg) {
            const recoverablePatterns = [
                /timeout/i,
                /超时/i,
                /网络/i,
                /network/i,
                /临时/i,
                /temporary/i
            ];
            
            return recoverablePatterns.some(pattern => pattern.test(errorMsg));
        }

        /**
         * 启动看门狗定时器
         */
        _startWatchdog() {
            this._stopWatchdog(); // 先清除已有的
            
            this.watchdogTimer = setTimeout(() => {
                const elapsed = Date.now() - this.startTime;
                ChatLog.warn('AI-State', `⏱️ 处理时间过长 (${(elapsed/1000).toFixed(1)}秒)，可能卡死`);
                
                if (typeof App.showToast === 'function') {
                    App.showToast(
                        `⏱️ AI响应时间较长(${(elapsed/1000).toFixed(0)}秒)，请耐心等待或刷新页面`, 
                        'warning', 
                        5000
                    );
                }
                
                // 如果超过2倍超时时间，强制结束
                if (elapsed > this.timeout * 2) {
                    this.handleError(new Error('处理超时'), 'WATCHDOG');
                }
            }, this.timeout);
        }

        /**
         * 重置看门狗
         */
        _resetWatchdog() {
            if (this.watchdogTimer) {
                clearTimeout(this.watchdogTimer);
                this._startWatchdog();
            }
        }

        /**
         * 停止看门狗
         */
        _stopWatchdog() {
            if (this.watchdogTimer) {
                clearTimeout(this.watchdogTimer);
                this.watchdogTimer = null;
            }
        }

        /**
         * 更新UI状态
         */
        _updateUIStatus(status, message) {
            try {
                // 尝试更新固定状态栏
                if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                    const iconMap = {
                        processing: '🤔',
                        completed: '✅',
                        error: '❌',
                        warning: '⚠️'
                    };
                    App.fixedStatusBar.setExtraStatus(
                        ` <span style="color:#666;font-size:11px;margin-left:8px;">${iconMap[status] || ''} ${message}</span>`,
                        true
                    );
                }
            } catch (e) {
                // UI更新失败不影响主流程
            }
        }

        /**
         * 显示用户友好的错误信息
         */
        _showUserFriendlyError(errorMsg, context) {
            let userMessage = '';
            let suggestion = '';
            
            // 根据错误类型生成友好消息
            if (/timeout|超时/i.test(errorMsg)) {
                userMessage = '⏱️ 响应超时';
                suggestion = '可能原因：Excel正在忙碌、网络延迟、或数据量过大<br>建议：稍后重试，或减少选区范围';
            } else if (/bridge|excel|com/i.test(errorMsg)) {
                userMessage = '❌ Excel连接异常';
                suggestion = '可能原因：Excel未正确加载、插件版本不兼容<br>建议：重启Excel或重新加载插件';
            } else if (/network|网络|fetch/i.test(errorMsg)) {
                userMessage = '🌐 网络问题';
                suggestion = '可能原因：无法连接到AI服务器<br>建议：检查网络设置和API配置';
            } else if (/api|token|key/i.test(errorMsg)) {
                userMessage = '🔑 API配置错误';
                suggestion = '可能原因：API密钥无效或额度不足<br>建议：检查设置面板中的API配置';
            } else {
                userMessage = '⚠️ 处理失败';
                suggestion = '发生了未知错误，请查看控制台日志或联系支持';
            }
            
            // 构建完整的错误提示
            const fullMessage = `${userMessage}<br><br><div style="font-size:11px;color:#888;margin-top:4px;">${suggestion}</div>`;
            
            // 显示错误
            if (typeof App.showToast === 'function') {
                App.showToast(fullMessage, 'error', 8000);
            } else if (typeof window.alert !== 'undefined') {
                alert(`${userMessage}\n\n${suggestion.replace(/<[^>]+>/g, '')}`);
            }
            
            // 记录到控制台
            ChatLog.error('AI-Error', `[${context}] ${userMessage}`, { original: errorMsg, suggestion });
        }
    }

    // ==================== 智能降级策略 ====================
    
    class SmartFallbackManager {
        constructor() {
            this.fallbackLevel = 0; // 0=正常, 1=部分降级, 2=完全降级
        }

        /**
         * 根据错误情况选择降级策略
         * @param {string} errorType - 错误类型
         * @returns {Object} 降级配置
         */
        getFallbackStrategy(errorType) {
            const strategies = {
                'CONTEXT_TIMEOUT': {
                    level: 1,
                    action: 'skip_context',
                    message: '跳过上下文获取，使用基础模式',
                    retryable: true
                },
                'BRIDGE_ERROR': {
                    level: 2,
                    action: 'mock_mode',
                    message: '切换到模拟模式（无实时Excel数据）',
                    retryable: false
                },
                'MODEL_ERROR': {
                    level: 1,
                    action: 'retry_with_fallback_model',
                    message: '尝试备用模型',
                    retryable: true
                },
                'NETWORK_ERROR': {
                    level: 1,
                    action: 'retry_with_backoff',
                    message: '网络抖动，等待后重试',
                    retryable: true
                }
            };

            return strategies[errorType] || {
                level: 2,
                action: 'abort',
                message: '无法恢复，终止流程',
                retryable: false
            };
        }

        /**
         * 执行降级操作
         */
        async executeFallback(strategy) {
            ChatLog.warn('AI-Fallback', `📉 执行降级策略: ${strategy.action} - ${strategy.message}`);
            
            switch (strategy.action) {
                case 'skip_context':
                    // 跳过上下文获取，使用空上下文
                    if (App.contextAwareness) {
                        App.contextAwareness.clearCache();
                    }
                    return { success: true, data: null };
                    
                case 'mock_mode':
                    // 切换到模拟模式
                    return { success: true, mockMode: true };
                    
                case 'retry_with_fallback_model':
                    // TODO: 实现模型切换逻辑
                    return { success: false, error: '暂不支持自动切换模型' };
                    
                case 'retry_with_backoff':
                    // 等待后重试
                    await new Promise(resolve => setTimeout(resolve, 2000));
                    return { success: true, shouldRetry: true };
                    
                default:
                    return { success: false, error: strategy.message };
            }
        }
    }

    // ==================== 性能监控 ====================
    
    class PerformanceMonitor {
        constructor() {
            this.metrics = [];
            this.maxMetrics = 100;
        }

        /**
         * 记录性能指标
         */
        recordMetric(name, duration, success = true) {
            this.metrics.push({
                name,
                duration,
                success,
                timestamp: Date.now()
            });
            
            // 保持队列大小
            if (this.metrics.length > this.maxMetrics) {
                this.metrics.shift();
            }
            
            // 输出性能警告
            if (duration > 3000) {
                ChatLog.warn('Perf', `⚠️ ${name} 耗时过长: ${(duration/1000).toFixed(1)}秒`);
            }
        }

        /**
         * 获取性能统计
         */
        getStats(metricName) {
            const filtered = this.metrics.filter(m => m.name === metricName);
            if (filtered.length === 0) return null;
            
            const durations = filtered.map(m => m.duration);
            const successRate = filtered.filter(m => m.success).length / filtered.length;
            
            return {
                count: filtered.length,
                avgDuration: durations.reduce((a, b) => a + b, 0) / durations.length,
                maxDuration: Math.max(...durations),
                minDuration: Math.min(...durations),
                successRate: (successRate * 100).toFixed(1) + '%'
            };
        }
    }

    // ==================== 初始化和注册 ====================
    
    // 创建单例实例
    const aiStateManager = new AIStateManager();
    const fallbackManager = new SmartFallbackManager();
    const performanceMonitor = new PerformanceMonitor();

    // 注册到全局App对象
    if (typeof App !== 'undefined') {
        App.AIStateManager = aiStateManager;
        App.SmartFallbackManager = fallbackManager;
        App.PerformanceMonitor = performanceMonitor;
        
        ChatLog.info('AI-Enhancer', '✅ AI体验增强器已加载');
        ChatLog.info('AI-Enhancer', '   - 智能状态管理');
        ChatLog.info('AI-Enhancer', '   - 降级恢复策略');
        ChatLog.info('AI-Enhancer', '   - 性能监控');
    }

})();