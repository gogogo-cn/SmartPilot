/**
 * ============================================================
 * keyboard-shortcuts.js — Keyboard Shortcuts
 * ============================================================
 * 功能：键盘快捷键模块，注册和管理全局快捷键、命令面板与快捷操作。
 *
 * 核心类/函数：
 *   - registerShortcut() — 注册快捷键
 *   - unregisterShortcut() — 注销快捷键
 *   - handleKeydown() — 处理按键事件
 *   - showCommandPalette() — 显示命令面板
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

class KeyboardShortcuts {
    constructor() {
        // 快捷键映射
        this.shortcuts = new Map();
        // 命令注册表
        this.commands = new Map();
        // 命令面板是否打开
        this.isCommandPanelOpen = false;
        // 当前输入的命令
        this.currentCommand = '';
        // 快捷命令前缀
        this.commandPrefix = '/';
        // 初始化默认快捷键
        this.initDefaultShortcuts();
        // 初始化默认命令
        this.initDefaultCommands();
        // 绑定键盘事件
        this.bindEvents();
    }

    /**
     * 初始化默认快捷键
     */
    initDefaultShortcuts() {
        // 文件操作
        this.registerShortcut('Ctrl+O', 'open_file', '打开文件');
        this.registerShortcut('Ctrl+S', 'save_file', '保存文件');
        this.registerShortcut('Ctrl+Shift+S', 'save_as', '另存为');
        
        // 编辑操作
        this.registerShortcut('Ctrl+Z', 'undo', '撤销');
        this.registerShortcut('Ctrl+Y', 'redo', '重做');
        this.registerShortcut('Ctrl+Shift+Z', 'redo', '重做');
        
        // 对话操作
        this.registerShortcut('Ctrl+N', 'new_chat', '新建对话');
        this.registerShortcut('Ctrl+L', 'clear_chat', '清空对话');
        this.registerShortcut('Ctrl+F', 'search_history', '搜索历史');
        
        // 功能操作
        this.registerShortcut('Ctrl+E', 'toggle_editor', '切换代码编辑器');
        this.registerShortcut('Ctrl+R', 'run_code', '执行代码');
        this.registerShortcut('Ctrl+Shift+R', 'recognize_image', '识别图片');
        
        // 导航操作
        this.registerShortcut('Ctrl+Tab', 'next_tab', '下一个标签页');
        this.registerShortcut('Ctrl+Shift+Tab', 'prev_tab', '上一个标签页');
        this.registerShortcut('Ctrl+1', 'switch_to_chat', '切换到对话');
        this.registerShortcut('Ctrl+2', 'switch_to_history', '切换到历史');
        this.registerShortcut('Ctrl+3', 'switch_to_skills', '切换到技能');
        this.registerShortcut('Ctrl+4', 'switch_to_settings', '切换到设置');
        
        // 帮助
        this.registerShortcut('F1', 'show_help', '显示帮助');
        this.registerShortcut('Ctrl+/', 'show_shortcuts', '显示快捷键列表');
        this.registerShortcut('Ctrl+Shift+P', 'command_palette', '命令面板');
        
        // 其他
        this.registerShortcut('Escape', 'close_dialog', '关闭对话框');
        this.registerShortcut('Ctrl+Enter', 'send_message', '发送消息');
        this.registerShortcut('Shift+Enter', 'new_line', '换行');
    }

    /**
     * 初始化默认命令
     */
    initDefaultCommands() {
        // 对话相关
        this.registerCommand('newchat', '新建对话', '💬', () => this.executeAction('new_chat'));
        this.registerCommand('clear', '清空对话', '🧹', () => this.executeAction('clear_chat'));
        this.registerCommand('history', '查看历史', '📜', () => this.executeAction('show_history'));
        
        // 识别相关
        this.registerCommand('ocr', '识图入表', '👁️', () => this.executeAction('recognize_image'));
        this.registerCommand('invoice', '识别发票', '🧾', () => this.executeAction('recognize_invoice'));
        this.registerCommand('idcard', '识别身份证', '🆔', () => this.executeAction('recognize_idcard'));
        this.registerCommand('card', '识别名片', '📇', () => this.executeAction('recognize_card'));
        
        // 代码相关
        this.registerCommand('code', '生成代码', '💻', () => this.executeAction('generate_code'));
        this.registerCommand('run', '执行代码', '▶️', () => this.executeAction('run_code'));
        this.registerCommand('editor', '代码编辑器', '📝', () => this.executeAction('toggle_editor'));
        
        // Excel 操作
        this.registerCommand('fill', '填入Excel', '📊', () => this.executeAction('fill_excel'));
        this.registerCommand('export', '导出数据', '📤', () => this.executeAction('export_data'));
        this.registerCommand('import', '导入数据', '📥', () => this.executeAction('import_data'));
        
        // 设置
        this.registerCommand('settings', '设置', '⚙️', () => this.executeAction('open_settings'));
        this.registerCommand('ai', 'AI 设置', '🤖', () => this.executeAction('open_ai_settings'));
        this.registerCommand('ocrsettings', 'OCR 设置', '🧠', () => this.executeAction('open_ocr_settings'));
        
        // 帮助
        this.registerCommand('help', '帮助', '❓', () => this.executeAction('show_help'));
        this.registerCommand('shortcuts', '快捷键', '⌨️', () => this.executeAction('show_shortcuts'));
        
        // 批量处理
        this.registerCommand('batch', '批量处理', '📁', () => this.executeAction('batch_process'));
        this.registerCommand('validate', '验证数据', '✅', () => this.executeAction('validate_data'));
        
        // 撤销/重做
        this.registerCommand('undo', '撤销', '↩️', () => this.executeAction('undo'));
        this.registerCommand('redo', '重做', '↪️', () => this.executeAction('redo'));
    }

    /**
     * 注册快捷键
     */
    registerShortcut(keyCombo, action, description) {
        const normalized = this.normalizeKeyCombo(keyCombo);
        this.shortcuts.set(normalized, {
            action: action,
            description: description,
            keyCombo: keyCombo
        });
    }

    /**
     * 注册命令
     */
    registerCommand(name, description, icon, handler) {
        this.commands.set(name.toLowerCase(), {
            name: name,
            description: description,
            icon: icon,
            handler: handler
        });
    }

    /**
     * 绑定键盘事件
     */
    bindEvents() {
        document.addEventListener('keydown', (e) => {
            this.handleKeyDown(e);
        });

        // 输入框特殊处理
        const messageInput = document.getElementById('message-input');
        if (messageInput) {
            messageInput.addEventListener('keydown', (e) => {
                this.handleInputKeyDown(e);
            });

            messageInput.addEventListener('input', (e) => {
                this.handleInput(e);
            });
        }
    }

    /**
     * 处理键盘按下事件
     */
    handleKeyDown(e) {
        // 忽略在输入框中的快捷键（除了特定的）
        if (this.isInputElement(e.target) && !e.ctrlKey && !e.metaKey) {
            return;
        }

        const keyCombo = this.getKeyCombo(e);
        const shortcut = this.shortcuts.get(keyCombo);

        if (shortcut) {
            e.preventDefault();
            this.executeAction(shortcut.action);
        }

        // 命令面板
        if (keyCombo === 'Ctrl+Shift+P') {
            e.preventDefault();
            this.toggleCommandPanel();
        }

        // ESC 关闭对话框
        if (e.key === 'Escape') {
            this.closeAllDialogs();
        }
    }

    /**
     * 处理输入框键盘事件
     */
    handleInputKeyDown(e) {
        const keyCombo = this.getKeyCombo(e);

        // Ctrl+Enter 发送消息
        if (keyCombo === 'Ctrl+Enter') {
            e.preventDefault();
            this.executeAction('send_message');
            return;
        }

        // Shift+Enter 换行（默认行为）
        if (e.key === 'Enter' && !e.shiftKey && !e.ctrlKey) {
            // 如果没有显示命令面板，则发送消息
            if (!this.isCommandPanelOpen) {
                e.preventDefault();
                this.executeAction('send_message');
            }
        }

        // ESC 关闭命令面板
        if (e.key === 'Escape' && this.isCommandPanelOpen) {
            this.hideCommandPanel();
        }

        // 上下箭头选择命令
        if (this.isCommandPanelOpen) {
            if (e.key === 'ArrowDown') {
                e.preventDefault();
                this.selectNextCommand();
            } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                this.selectPrevCommand();
            } else if (e.key === 'Enter') {
                e.preventDefault();
                this.executeSelectedCommand();
            }
        }
    }

    /**
     * 处理输入事件(快捷命令)
     * 注:已禁用"/"快捷命令功能,由app.js中的快捷指令系统统一处理
     */
    handleInput(e) {
        // 禁用"/"开头的命令面板触发
        // const value = e.target.value;
        // if (value.startsWith(this.commandPrefix)) {
        //     this.currentCommand = value.slice(1);
        //     this.showCommandPanel(this.currentCommand);
        // } else if (this.isCommandPanelOpen) {
        //     this.hideCommandPanel();
        // }
    }

    /**
     * 显示命令面板
     */
    showCommandPanel(filter = '') {
        this.isCommandPanelOpen = true;
        let panel = document.getElementById('command-panel');
        
        if (!panel) {
            panel = this.createCommandPanel();
        }

        const commands = this.filterCommands(filter);
        this.renderCommandPanel(panel, commands);
        panel.style.display = 'block';
    }

    /**
     * 隐藏命令面板
     */
    hideCommandPanel() {
        this.isCommandPanelOpen = false;
        const panel = document.getElementById('command-panel');
        if (panel) {
            panel.style.display = 'none';
        }
    }

    /**
     * 切换命令面板
     */
    toggleCommandPanel() {
        if (this.isCommandPanelOpen) {
            this.hideCommandPanel();
        } else {
            this.showCommandPanel();
        }
    }

    /**
     * 创建命令面板
     */
    createCommandPanel() {
        const panel = document.createElement('div');
        panel.id = 'command-panel';
        panel.className = 'command-panel';
        panel.innerHTML = `
            <div class="command-panel-header">
                <span>快捷命令</span>
                <span class="command-hint">↑↓ 选择 · Enter 执行 · Esc 关闭</span>
            </div>
            <div class="command-list"></div>
        `;
        
        // 添加到输入区域
        const inputArea = document.getElementById('input-area');
        if (inputArea) {
            inputArea.appendChild(panel);
        }

        return panel;
    }

    /**
     * 渲染命令面板
     */
    renderCommandPanel(panel, commands) {
        const list = panel.querySelector('.command-list');
        list.innerHTML = '';

        if (commands.length === 0) {
            list.innerHTML = '<div class="command-empty">无匹配命令</div>';
            return;
        }

        commands.forEach((cmd, index) => {
            const item = document.createElement('div');
            item.className = 'command-item';
            if (index === 0) item.classList.add('selected');
            item.dataset.command = cmd.name;
            item.innerHTML = `
                <span class="command-icon">${cmd.icon}</span>
                <span class="command-name">${cmd.name}</span>
                <span class="command-desc">${cmd.description}</span>
            `;
            item.addEventListener('click', () => {
                this.executeCommand(cmd.name);
            });
            list.appendChild(item);
        });
    }

    /**
     * 过滤命令
     */
    filterCommands(filter) {
        const filterLower = filter.toLowerCase();
        return Array.from(this.commands.values())
            .filter(cmd => 
                cmd.name.toLowerCase().includes(filterLower) ||
                cmd.description.toLowerCase().includes(filterLower)
            )
            .slice(0, 10); // 最多显示10个
    }

    /**
     * 选择下一个命令
     */
    selectNextCommand() {
        const items = document.querySelectorAll('.command-item');
        const selected = document.querySelector('.command-item.selected');
        
        if (selected && items.length > 0) {
            selected.classList.remove('selected');
            const next = selected.nextElementSibling || items[0];
            next.classList.add('selected');
            next.scrollIntoView({ block: 'nearest' });
        }
    }

    /**
     * 选择上一个命令
     */
    selectPrevCommand() {
        const items = document.querySelectorAll('.command-item');
        const selected = document.querySelector('.command-item.selected');
        
        if (selected && items.length > 0) {
            selected.classList.remove('selected');
            const prev = selected.previousElementSibling || items[items.length - 1];
            prev.classList.add('selected');
            prev.scrollIntoView({ block: 'nearest' });
        }
    }

    /**
     * 执行选中的命令
     */
    executeSelectedCommand() {
        const selected = document.querySelector('.command-item.selected');
        if (selected) {
            const commandName = selected.dataset.command;
            this.executeCommand(commandName);
        }
    }

    /**
     * 执行命令
     */
    executeCommand(name) {
        const command = this.commands.get(name.toLowerCase());
        if (command && command.handler) {
            command.handler();
            this.hideCommandPanel();
            
            // 清空输入框中的命令
            const input = document.getElementById('message-input');
            if (input && input.value.startsWith(this.commandPrefix)) {
                input.value = '';
            }
        }
    }

    /**
     * 执行动作
     */
    executeAction(action) {
        ChatLog.info('KEYBOARD', `[KeyboardShortcuts] 执行动作: ${action}`);
        
        // 触发自定义事件
        const event = new CustomEvent('app:action', {
            detail: { action: action }
        });
        document.dispatchEvent(event);

        // 执行对应的函数
        switch (action) {
            case 'new_chat':
                if (window.startNewChat) window.startNewChat();
                break;
            case 'clear_chat':
                if (window.clearChat) window.clearChat();
                break;
            case 'undo':
                if (window.undoLastAction) window.undoLastAction();
                break;
            case 'redo':
                if (window.redoLastAction) window.redoLastAction();
                break;
            case 'toggle_editor':
                if (window.toggleEditor) window.toggleEditor();
                break;
            case 'run_code':
                if (window.executeCode) window.executeCode();
                break;
            case 'send_message':
                if (window.sendMessage) window.sendMessage();
                break;
            case 'show_help':
                if (typeof window.showProductHelp === 'function') window.showProductHelp({ tab: 'start' });
                else this.showShortcutsHelp();
                break;
            case 'show_shortcuts':
                this.showShortcutsHelp();
                break;
            case 'command_palette':
                this.toggleCommandPanel();
                break;
            case 'open_ai_settings':
            case 'open_settings':
                if (typeof window.openAiSettings === 'function') window.openAiSettings();
                break;
            case 'recognize_image':
            case 'recognize_invoice':
            case 'recognize_idcard':
            case 'recognize_card':
                if (typeof window.openOcrPanel === 'function') window.openOcrPanel();
                break;
            // 其他动作...
        }
    }

    /**
     * 显示快捷键帮助（优先走应用内帮助面板）
     */
    showShortcutsHelp() {
        if (typeof window.showProductHelp === 'function') {
            window.showProductHelp({ tab: 'shortcuts' });
            return;
        }
        const shortcuts = Array.from(this.shortcuts.entries())
            .map(([key, value]) => ({
                key: value.keyCombo,
                description: value.description
            }));

        const html = `
            <div class="shortcuts-help">
                <h3>⌨️ 键盘快捷键</h3>
                <div class="shortcuts-list">
                    ${shortcuts.map(s => `
                        <div class="shortcut-item">
                            <kbd>${s.key}</kbd>
                            <span>${s.description}</span>
                        </div>
                    `).join('')}
                </div>
            </div>
        `;

        // 显示对话框
        if (window.showDialog) {
            window.showDialog('快捷键帮助', html);
        } else if (App.fixedStatusBar && typeof App.fixedStatusBar.setExtraStatus === 'function') {
                App.fixedStatusBar.setExtraStatus(' <span style="color:#999;font-size:11px;margin-left:8px;">⌨️ 快捷键: ' + shortcuts.map(s => s.key).join(', ') + '</span>', true);
            } else if (typeof window.showToast === 'function') {
                window.showToast('⌨️ 快捷键: ' + shortcuts.map(s => s.key).join(', '), 'info');
        } else {
            ChatLog.info('KEYBOARD', '快捷键:', shortcuts.map(s => `${s.key}: ${s.description}`).join('\n'));
        }
    }

    /**
     * 关闭所有对话框
     */
    closeAllDialogs() {
        const dialogs = document.querySelectorAll('.dialog-overlay');
        dialogs.forEach(dialog => {
            dialog.style.display = 'none';
        });
        this.hideCommandPanel();
    }

    /**
     * 获取按键组合
     */
    getKeyCombo(e) {
        const parts = [];
        
        if (e.ctrlKey || e.metaKey) parts.push('Ctrl');
        if (e.altKey) parts.push('Alt');
        if (e.shiftKey) parts.push('Shift');
        
        const key = e.key.length === 1 ? e.key.toUpperCase() : e.key;
        parts.push(key);
        
        return parts.join('+');
    }

    /**
     * 标准化按键组合
     */
    normalizeKeyCombo(keyCombo) {
        return keyCombo
            .split('+')
            .map(k => k.trim())
            .sort((a, b) => {
                // 修饰键在前
                const modifiers = ['Ctrl', 'Alt', 'Shift'];
                const aIsMod = modifiers.includes(a);
                const bIsMod = modifiers.includes(b);
                if (aIsMod && !bIsMod) return -1;
                if (!aIsMod && bIsMod) return 1;
                return a.localeCompare(b);
            })
            .join('+');
    }

    /**
     * 判断是否是输入元素
     */
    isInputElement(element) {
        const tagName = element.tagName.toLowerCase();
        return tagName === 'input' || 
               tagName === 'textarea' || 
               element.contentEditable === 'true';
    }

    /**
     * 获取所有快捷键
     */
    getAllShortcuts() {
        return Array.from(this.shortcuts.entries()).map(([key, value]) => ({
            keyCombo: value.keyCombo,
            action: value.action,
            description: value.description
        }));
    }

    /**
     * 获取所有命令
     */
    getAllCommands() {
        return Array.from(this.commands.values());
    }
}

// 创建全局实例
window.keyboardShortcuts = new KeyboardShortcuts();

ChatLog.info('KEYBOARD', '✅ 快捷键系统已加载');
