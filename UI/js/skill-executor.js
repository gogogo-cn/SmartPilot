/**
 * ============================================================
 * skill-executor.js — Skill Executor
 * ============================================================
 * 功能：Skill 执行器，负责在沙箱环境中运行 Skill 代码、捕获输出与异常处理。
 *
 * 核心类/函数：
 *   - executeSkillCode() — 执行 Skill 代码
 *   - createSandbox() — 创建沙箱
 *   - getSkillOutput() — 获取 Skill 输出
 *
 * 依赖：
 *   - app-ai-skill.js
 *   - app-utils.js
 * ============================================================
 */

class SkillExecutor {
    constructor() {
        this.executionHistory = [];
        this._currentAbortController = null;  // 当前正在执行的操作的取消控制器
    }

    /**
     * 取消当前正在执行的操作
     */
    cancelCurrentExecution() {
        if (this._currentAbortController) {
            this._currentAbortController.abort();
            this._currentAbortController = null;
            ChatLog.info('SKILL', '已取消当前操作');
            return true;
        }
        return false;
    }

    /**
     * 检查是否已被取消
     */
    _checkCancellation(abortController) {
        if (abortController && abortController.signal.aborted) {
            throw new Error('操作已被用户取消');
        }
    }

    /**
     * 执行动作
     */
    async executeAction(action) {
        // 创建取消控制器
        this._currentAbortController = new AbortController();
        const abortController = this._currentAbortController;

        try {
            ChatLog.info('SKILL', '执行动作', action);

            switch(action.actionType) {
                case 'excel_operation':
                    return await this.executeExcelOperation(action, abortController);
                
                case 'code_generation':
                    return await this.generateAndExecuteCode(action, abortController);
                
                case 'data_extraction':
                    return await this.extractAndFillData(action, abortController);
                
                case 'image_recognition':
                    return await this.recognizeImage(action, abortController);
                
                default:
                    throw new Error(`未知的动作类型: ${action.actionType}`);
            }
        } catch (error) {
            if (error.message === '操作已被用户取消') {
                ChatLog.info('SKILL', '操作已被用户取消');
                return {
                    success: false,
                    message: '操作已被用户取消',
                    cancelled: true
                };
            }
            ChatLog.error('SKILL', '执行失败', error);
            return {
                success: false,
                message: error.message,
                error: error
            };
        } finally {
            // 清理取消控制器
            if (this._currentAbortController === abortController) {
                this._currentAbortController = null;
            }
        }
    }

    /**
     * 执行Excel操作
     * 【修复】统一调用后端 ExecuteAction(actionJson)，不再调用不存在的 ExecuteExcelOperation
     */
    async executeExcelOperation(action, abortController) {
        const { parameters } = action;
        const operation = parameters.operation;

        ChatLog.info('SKILL', 'Excel操作', { operation, parameters });
        this._checkCancellation(abortController);

        if (!this.isExcelBridgeAvailable()) {
            throw new Error('Excel未连接');
        }

        try {
            const bridge = window.chrome.webview.hostObjects.excelBridge;
            const resultJson = await bridge.ExecuteAction(JSON.stringify(action));
            const result = JSON.parse(resultJson);

            this.executionHistory.push({
                action: action,
                result: result,
                timestamp: new Date()
            });

            return result;

        } catch (error) {
            throw new Error(`Excel操作失败: ${error.message}`);
        }
    }

    /**
     * 生成并执行代码
     * 【修复】统一调用后端 ExecuteAction(actionJson)，由后端决定是走 VSTO 直接操作还是生成 VBA。
     * 如果后端返回了生成的代码，则展示代码执行对话框；否则直接返回执行结果。
     */
    async generateAndExecuteCode(action, abortController) {
        const { parameters } = action;
        const description = parameters.description;

        ChatLog.info('SKILL', '代码生成', { description });
        this._checkCancellation(abortController);

        if (!this.isExcelBridgeAvailable()) {
            throw new Error('Excel未连接');
        }

        try {
            const bridge = window.chrome.webview.hostObjects.excelBridge;
            const resultJson = await bridge.ExecuteAction(JSON.stringify(action));
            const result = JSON.parse(resultJson);

            // 如果后端返回了 VBA 代码，展示代码执行对话框
            if (result.success && result.data && result.data.code) {
                return await this.showCodeExecutionDialog(result.data.code, action);
            }

            this.executionHistory.push({
                action: action,
                result: result,
                timestamp: new Date()
            });

            return result;

        } catch (error) {
            throw new Error(`代码生成失败: ${error.message}`);
        }
    }

    /**
     * 请求代码生成
     */
    async requestCodeGeneration(description) {
        const prompt = `请根据以下描述生成VBA代码:

描述: ${description}

【输出要求】
1. 生成完整的 VBA 代码（Sub/End Sub）
2. 包含完整的错误处理
3. 使用清晰的变量名和注释
4. **直接返回代码，不要任何解释**

【绝对禁止】
- 禁止输出"使用说明"、"功能说明"、"执行步骤"等标题
- 禁止输出"按Alt+F11"、"打开VBA编辑器"等操作指南
- 禁止输出任何代码之外的文字

请使用\`\`\`vba代码块格式返回代码。`;

        // 调用AI API
        const response = await this.callAI(prompt);
        
        // 提取代码块
        const codeMatch = response.match(/```vba\s*([\s\S]*?)```/);
        return codeMatch ? codeMatch[1].trim() : response.trim();
    }

    /**
     * 显示代码执行对话框
     */
    async showCodeExecutionDialog(code, action) {
        return new Promise((resolve) => {
            // 创建对话框
            const dialog = document.createElement('div');
            dialog.className = 'code-execution-dialog';
            dialog.innerHTML = `
                <div class="dialog-overlay"></div>
                <div class="dialog-content">
                    <div class="dialog-header">
                        <h3>💻 VBA代码生成完成</h3>
                        <button class="dialog-close">&times;</button>
                    </div>
                    <div class="dialog-body">
                        <p>已为您生成以下VBA代码:</p>
                        <pre><code>${this.escapeHtml(code)}</code></pre>
                        <div class="dialog-actions">
                            <button class="btn btn-primary" id="execute-code-btn">✅ 执行代码</button>
                            <button class="btn btn-secondary" id="copy-code-btn">📋 复制代码</button>
                            <button class="btn btn-secondary" id="cancel-code-btn">❌ 取消</button>
                        </div>
                    </div>
                </div>
            `;

            document.body.appendChild(dialog);

            // 执行代码按钮
            dialog.querySelector('#execute-code-btn').onclick = async () => {
                try {
                    const bridge = window.chrome.webview.hostObjects.excelBridge;
                    // 重新构造 code_generation 动作，让后端 VBACodeSandbox 执行
                    const executeAction = {
                        actionType: 'code_generation',
                        parameters: {
                            ...action.parameters,
                            description: action.parameters.description || '执行生成的VBA代码',
                            generatedCode: code
                        }
                    };
                    const resultJson = await bridge.ExecuteAction(JSON.stringify(executeAction));

                    const result = JSON.parse(resultJson);
                    dialog.remove();
                    resolve(result);
                } catch (error) {
                    if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                        App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">❌ 执行失败: ' + error.message + '</span>', true);
                    } else if (typeof window.showToast === 'function') {
                        window.showToast('❌ 执行失败: ' + error.message, 'error');
                    }
                    dialog.remove();
                    resolve({ success: false, message: error.message });
                }
            };

            // 复制代码按钮
            dialog.querySelector('#copy-code-btn').onclick = () => {
                navigator.clipboard.writeText(code);
                if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                    App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">✅ 代码已复制到剪贴板</span>', true);
                } else if (typeof window.showToast === 'function') {
                    window.showToast('✅ 代码已复制到剪贴板', 'success');
                }
            };

            // 取消按钮
            dialog.querySelector('#cancel-code-btn').onclick = () => {
                dialog.remove();
                resolve({ success: false, message: '用户取消执行' });
            };

            // 关闭按钮
            dialog.querySelector('.dialog-close').onclick = () => {
                dialog.remove();
                resolve({ success: false, message: '用户取消执行' });
            };
        });
    }

    /**
     * 数据提取并填充
     */
    async extractAndFillData(action, abortController) {
        const { parameters } = action;
        const dataType = parameters.dataType;
        const imageBase64 = parameters.imageBase64;
        const targetSheet = parameters.targetSheet;

        ChatLog.info('SKILL', '数据提取', { dataType, targetSheet });
        this._checkCancellation(abortController);

        if (!this.isExcelBridgeAvailable()) {
            throw new Error('Excel未连接');
        }

        try {
            // 统一调用后端 ExecuteAction(actionJson)
            const bridge = window.chrome.webview.hostObjects.excelBridge;
            const resultJson = await bridge.ExecuteAction(JSON.stringify(action));
            const result = JSON.parse(resultJson);

            this.executionHistory.push({
                action: action,
                result: result,
                timestamp: new Date()
            });

            return result;

        } catch (error) {
            throw new Error(`数据提取失败: ${error.message}`);
        }
    }

    /**
     * 识别图片数据
     */
    async recognizeImageData(imageBase64, dataType) {
        // 构建识别提示词
        const prompts = {
            'id_card': '请识别身份证信息,返回JSON格式: { "name": "", "gender": "", "nation": "", "birth": "", "address": "", "id_number": "" }',
            'business_card': '请识别名片信息,返回JSON格式: { "name": "", "title": "", "company": "", "phone": "", "email": "", "address": "" }',
            'invoice': '请识别发票信息,返回JSON格式: { "invoice_code": "", "invoice_number": "", "date": "", "amount": "", "tax": "", "total": "" }',
            'table': '请识别表格数据,返回JSON数组格式: [{ "列标题1": "值1", "列标题2": "值2" }]，使用图片中真实的列名，禁止使用"列1"、"列2"等占位符'
        };

        const prompt = prompts[dataType] || '请识别图片内容,返回JSON格式的结构化数据';

        // 调用AI识别
        const response = await this.callAIWithImage(prompt, imageBase64);

        // 【P1修复】使用安全JSON解析（容错处理）
        return this.safeJsonParse(response);
    }

    /**
     * 【P1修复】安全JSON解析 - 容错处理
     */
    safeJsonParse(content) {
        if (!content || typeof content !== 'string') {
            throw new Error('内容为空或不是字符串');
        }

        try {
            return JSON.parse(content);
        } catch (e) {
            const patterns = [
                /```(?:json)?\s*([\s\S]*?)```/,
                /\{[\s\S]*\}/
            ];

            for (const pattern of patterns) {
                const match = content.match(pattern);
                if (match) {
                    try {
                        const jsonStr = match[1] || match[0];
                        return JSON.parse(jsonStr.trim());
                    } catch (innerError) {
                        continue;
                    }
                }
            }

            throw new Error('无法从响应内容中提取有效的JSON格式');
        }
    }

    /**
     * 图片识别
     */
    async recognizeImage(action, abortController) {
        const { parameters } = action;
        const imageBase64 = parameters.imageBase64;
        const recognitionType = parameters.recognitionType;

        ChatLog.info('SKILL', '图片识别', { recognitionType });
        this._checkCancellation(abortController);

        try {
            // 调用AI识别
            const result = await this.recognizeImageData(imageBase64, recognitionType);

            return {
                success: true,
                message: '图片识别成功',
                data: result
            };

        } catch (error) {
            throw new Error(`图片识别失败: ${error.message}`);
        }
    }

    /**
     * 调用AI API
     */
    async callAI(prompt) {
        // 使用现有的AI调用逻辑
        if (window.callAIAndGetResponse) {
            const response = await window.callAIAndGetResponse(prompt);
            return response.choices[0].message.content;
        }
        
        throw new Error('AI API不可用');
    }

    /**
     * 调用AI API(带图片)
     */
    async callAIWithImage(prompt, imageBase64) {
        // 使用现有的多模态AI调用逻辑
        if (window.callAIWithImage) {
            return await window.callAIWithImage(prompt, imageBase64);
        }
        
        throw new Error('多模态AI API不可用');
    }

    /**
     * 检查ExcelBridge是否可用
     */
    isExcelBridgeAvailable() {
        return window.chrome && 
               window.chrome.webview && 
               window.chrome.webview.hostObjects && 
               window.chrome.webview.hostObjects.excelBridge;
    }

    /**
     * HTML转义
     */
    escapeHtml(text) {
        if (!text) return '';
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    /**
     * 获取执行历史
     */
    getExecutionHistory() {
        return this.executionHistory;
    }

    /**
     * 清空执行历史
     */
    clearHistory() {
        this.executionHistory = [];
    }
}

// 创建全局实例
window.skillExecutor = new SkillExecutor();
