/**
 * ============================================================
 * app-ocr-parser.js — OCR Result Parser
 * ============================================================
 * 功能：OCR 结果解析模块，将 OCR 原始输出转换为结构化数据，支持身份证、营业执照、发票等票据解析。
 *
 * 核心类/函数：
 *   - parseOcrResult() — 解析 OCR 结果
 *   - extractIdCardInfo() — 提取身份证信息
 *   - extractBusinessLicenseInfo() — 提取营业执照信息
 *   - normalizeTableData() — 规范化表格数据
 *
 * 依赖：
 *   - app-ocr.js
 *   - app-utils.js
 * ============================================================
 */

ChatLog.info('OCR', '[Parser] 脚本加载');
(function() {
    'use strict';

    function pt(key, vars) {
        try {
            if (window.App && App.i18n && App.i18n.t) {
                var v = App.i18n.t(key, vars);
                if (v && v !== key) return v;
            }
        } catch (_) {}
        return key;
    }

    var CATEGORY_KEYS = {
        '通用': 'ocr.parser.cat.general',
        '票据': 'ocr.parser.cat.receipt',
        '证件': 'ocr.parser.cat.id',
        '商务': 'ocr.parser.cat.business',
        '自定义': 'ocr.parser.cat.custom'
    };

    function categoryLabel(cat) {
        var key = CATEGORY_KEYS[cat];
        return key ? pt(key) : cat;
    }

    function setParserTitle(name) {
        var el = document.getElementById('parser-current-name');
        if (!el) return;
        if (!name) {
            el.setAttribute('data-i18n', 'ocr.parser.pickOrNew');
            el.textContent = pt('ocr.parser.pickOrNew');
        } else {
            el.removeAttribute('data-i18n');
            el.textContent = name;
        }
    }

    function applyLocale() {
        var select = document.getElementById('parser-category-select');
        if (select) {
            for (var i = 0; i < select.options.length; i++) {
                var opt = select.options[i];
                opt.textContent = opt.value ? categoryLabel(opt.value) : pt('ocr.parser.allCategories');
            }
        }
        var title = document.getElementById('parser-current-name');
        if (title && title.getAttribute('data-i18n') === 'ocr.parser.pickOrNew') {
            title.textContent = pt('ocr.parser.pickOrNew');
        }
        var search = document.getElementById('parser-search-input');
        var cat = document.getElementById('parser-category-select');
        if (typeof renderScriptList === 'function') {
            renderScriptList(search ? search.value : '', cat ? cat.value : '');
        }
        if (typeof populateMainScriptSelector === 'function') populateMainScriptSelector();
        if (typeof populateBatchScriptSelector === 'function') populateBatchScriptSelector();
    }

    // ==================== 状态管理 ====================

    let _scripts = [];           // 所有解析脚本列表
    let _currentScript = null;   // 当前选中的脚本
    let _currentView = 'edit';   // 当前视图: edit | preview | fields
    let _lastResult = null;      // 最近一次的解析结果

    // ==================== 工具函数 ====================

    /**
     * 格式化JavaScript代码（紧凑美化显示）
     * @param {string} code - 压缩或未格式化的JS代码
     * @returns {string} - 格式化后的代码（紧凑风格）
     */
    function formatJsCode(code) {
        if (!code || typeof code !== 'string') return '';
        
        try {
            // 清理输入
            let cleaned = code.replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();
            
            // 逐字符处理，维护正确的缩进
            let result = '';
            let indentLevel = 0;
            const indent = '    ';
            let inString = false;
            let stringChar = '';
            let needNewline = false;
            
            for (let i = 0; i < cleaned.length; i++) {
                const char = cleaned[i];
                
                // 字符串内容原样保留
                if (!inString && (char === '"' || char === "'" || char === '`')) {
                    inString = true;
                    stringChar = char;
                    result += char;
                    continue;
                }
                
                if (inString) {
                    result += char;
                    if (char === stringChar && cleaned[i - 1] !== '\\') {
                        inString = false;
                    }
                    continue;
                }
                
                // 跳过原有空白（由我们控制格式）
                if (/\s/.test(char)) {
                    // 保留关键字后的一个空格
                    if (result.length > 0) {
                        const lastChar = result[result.length - 1];
                        if (/\w/.test(lastChar)) {
                            // 检查是否是关键字
                            const words = ['function','if','else','for','while','do','switch','case','catch','finally','return','var','let','const','new','typeof','throw','try','class','extends'];
                            const text = result;
                            const isKeyword = words.some(w => text.endsWith(w));
                            if (isKeyword) {
                                result += ' ';
                            }
                        }
                    }
                    continue;
                }
                
                // 需要换行时先换行并添加缩进
                if (needNewline) {
                    result += '\n' + indent.repeat(indentLevel);
                    needNewline = false;
                }
                
                // 开括号 {
                if (char === '{') {
                    // 前面加空格（如 function() {）
                    if (result.length > 0 && result[result.length - 1] !== ' ' && result[result.length - 1] !== '\n') {
                        result += ' ';
                    }
                    result += '{';
                    indentLevel++;
                    needNewline = true;
                    continue;
                }
                
                // 闭括号 }
                if (char === '}') {
                    indentLevel = Math.max(0, indentLevel - 1);
                    result += '\n' + indent.repeat(indentLevel) + '}';
                    needNewline = true;
                    continue;
                }
                
                // 分号 ;
                if (char === ';') {
                    result += ';';
                    needNewline = true;
                    continue;
                }
                
                // 逗号 ,
                if (char === ',') {
                    result += ',';
                    // 检查后面是不是 }]) ，如果是就不换行
                    let j = i + 1;
                    while (j < cleaned.length && /\s/.test(cleaned[j])) j++;
                    if (cleaned[j] && !'}])'.includes(cleaned[j])) {
                        needNewline = true;
                    }
                    continue;
                }
                
                // 其他字符原样输出
                result += char;
            }
            
            // 最终清理
            result = result.replace(/\n{3,}/g, '\n\n'); // 最多连续两个空行
            result = result.replace(/[ \t]+$/gm, ''); // 清理行尾空白
            result = result.trim();
            
            return result;
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 格式化失败:', ex);
            return code; // 失败时返回原始代码
        }
    }

    /**
     * HTML转义（防止XSS）
     */

    // ==================== 初始化 ====================
    
    /**
     * 初始化解析模块（在DOM加载完成后调用）
     * 包含自动重试机制，等待WebView2 COM桥接就绪
     */
    function excelBridgeReady() {
        return !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge);
    }

    async function callOcrApi(method, args) {
        args = args || [];
        if (excelBridgeReady()) {
            const bridge = window.chrome.webview.hostObjects.excelBridge;
            if (method === 'GetParseScripts') return await bridge.GetParseScripts();
            if (method === 'GetParseScriptCategories') return await bridge.GetParseScriptCategories();
            if (method === 'GetOcrElementsJson') return await bridge.GetOcrElementsJson();
            if (method === 'GetParseScript') return await bridge.GetParseScript(args[0]);
            if (method === 'SaveParseScript') return await bridge.SaveParseScript(args[0]);
            if (method === 'DeleteParseScript') return await bridge.DeleteParseScript(args[0]);
            if (method === 'ClearParseScriptCache') return await bridge.ClearParseScriptCache();
            if (method === 'ImportParseScriptFromMarkdown') return await bridge.ImportParseScriptFromMarkdown(args[0]);
            if (method === 'DuplicateParseScript') return await bridge.DuplicateParseScript(args[0], args[1]);
            if (method === 'ResetBuiltinScript') return await bridge.ResetBuiltinScript(args[0]);
            return await bridge[method](...args);
        }
        if (typeof window.ocrHostCall === 'function') {
            return await window.ocrHostCall('OcrBridge.' + method, args || []);
        }
        return null;
    }

    async function init(retryCount = 0) {
        if (window.__ocrOpeningHost) return;
        if (retryCount > 0) {
            ChatLog.info('OCR', `[Parser] 初始化重试 (${retryCount}/10)`);
        }
        try {
            // 只有页面在 WebView2 里、桥还没注入时才等待。普通浏览器没有 WebView2，直接继续。
            const hasWebView = !!(window.chrome && window.chrome.webview);
            if (hasWebView && !excelBridgeReady() && retryCount < 10) {
                ChatLog.warn('OCR', `[Parser] WebView2 桥尚未注入，500ms后重试... (${retryCount + 1}/10)`);
                await new Promise(resolve => setTimeout(resolve, 500));
                return init(retryCount + 1);
            }

            await loadScripts();
            loadCategories();
            ChatLog.info('OCR', `[Parser] 脚本加载完成 (${_scripts.length}个)`);

            var ocrTextarea = document.getElementById('ocr-result-text');
            if (ocrTextarea) {
                ocrTextarea.addEventListener('input', function() {
                    updateMainParseSectionVisibility();
                });
                updateMainParseSectionVisibility();
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 初始化失败:', ex);
            if (excelBridgeReady()) {
                showToast('解析模块初始化失败: ' + ex.message);
            }
        }
    }

    // ==================== 脚本加载与管理 ====================

    /**
     * 从后端加载所有解析脚本
     */
    async function loadScripts() {
        try {
            const resultJson = await callOcrApi('GetParseScripts');
            if (!resultJson) return;
            const result = JSON.parse(resultJson);
            
            if (result.success) {
                _scripts = result.scripts || [];
                ChatLog.info('OCR', `[Parser] 脚本加载完成 (${_scripts.length}个)`);
                renderScriptList();

                // 同时填充主页脚本选择器
                populateMainScriptSelector();
                
                // 如果有当前选中的脚本，保持选中状态
                if (_currentScript) {
                    const found = _scripts.find(s => s.id === _currentScript.id);
                    if (!found) _currentScript = null;
                }
            } else {
                ChatLog.error('OCR', '[Parser] 加载失败:', result.error);
                showToast('加载脚本失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 加载脚本失败:', ex);
            ChatLog.error('OCR', '[Parser] 异常堆栈:', ex.stack);
            showToast('加载脚本失败: ' + ex.message);
        }
    }

    /**
     * 加载分类下拉框选项
     */
    async function loadCategories() {
        // 安全检查
        try {
            const resultJson = await callOcrApi('GetParseScriptCategories');
            if (!resultJson) return;
            const result = JSON.parse(resultJson);
            
            if (result.success && result.categories) {
                const select = document.getElementById('parser-category-select');
                if (select) {
                    // 保留"全部分类"选项
                    const firstOption = select.options[0];
                    select.innerHTML = '';
                    select.appendChild(firstOption);
                    
                    result.categories.forEach(cat => {
                        const option = document.createElement('option');
                        option.value = cat;
                        option.textContent = categoryLabel(cat);
                        select.appendChild(option);
                    });
                }
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 加载分类失败:', ex);
        }
    }

    /**
     * 渲染脚本列表到左侧边栏
     */
    function renderScriptList(filterText, filterCategory) {
        const container = document.getElementById('parser-script-list');
        if (!container) return;

        let filtered = _scripts;
        
        // 按文本过滤
        if (filterText) {
            const keyword = filterText.toLowerCase();
            filtered = filtered.filter(s => 
                s.name.toLowerCase().includes(keyword) ||
                (s.description && s.description.toLowerCase().includes(keyword))
            );
        }

        // 按分类过滤
        if (filterCategory) {
            filtered = filtered.filter(s => s.category === filterCategory);
        }

        if (filtered.length === 0) {
            container.innerHTML = '<div class="parser-loading">' +
                (filterText || filterCategory ? pt('ocr.parser.noMatch') : pt('ocr.parser.none')) + '</div>';
            return;
        }

        container.innerHTML = filtered.map(script => `
            <div class="parser-script-item ${_currentScript && _currentScript.id === script.id ? 'active' : ''}"
                 onclick="App.Parser.selectScript('${script.id}')"
                 data-id="${script.id}">
                <span class="parser-script-icon">${script.icon || '📝'}</span>
                <div class="parser-script-info">
                    <div class="parser-script-name">${escapeHtml(script.name)}</div>
                    ${script.description ? `<div class="parser-script-desc-hint">${escapeHtml(script.description)}</div>` : ''}
                </div>
                ${script.isBuiltIn ? '<span class="parser-badge builtin">内置</span>' : ''}
                ${!script.isBuiltIn ? `<button class="parser-script-delete" onclick="event.stopPropagation(); App.Parser.deleteCurrentById('${script.id}')" title="删除脚本">✕</button>` : ''}
            </div>
        `).join('');
    }

    /**
     * 选择并加载一个脚本进行编辑/预览
     */
    async function selectScript(scriptId) {
        const script = _scripts.find(s => s.id === scriptId);
        if (!script) {
            // 从后端重新获取
            try {
                const resultJson = await callOcrApi('GetParseScript', [scriptId]);
                const result = JSON.parse(resultJson);
                if (result.success) {
                    loadScriptToEditor(result.script);
                }
            } catch (ex) {
                showToast('加载脚本失败: ' + ex.message);
            }
            return;
        }

        loadScriptToEditor(script);
    }

    /**
     * 将脚本数据加载到编辑器表单
     */
    function loadScriptToEditor(script) {
        _currentScript = script;

        // ✅ 保存原始未格式化的代码（用于执行）
        if (script.script && !script._rawScript) {
            script._rawScript = script.script;
        }

        // 更新表单字段
        document.getElementById('parser-script-name').value = script.name || '';
        document.getElementById('parser-script-desc').value = script.description || '';
        document.getElementById('parser-script-icon').value = script.icon || '📝';
        document.getElementById('parser-script-category').value = script.category || '通用';
        document.getElementById('parser-script-output-format').value = script.outputFormat || 'table';

        // ✅ 显示时使用格式化后的代码（正常多行显示）
        const codeEditor = document.getElementById('parser-script-code');
        if (codeEditor) {
            const rawCode = script._rawScript || script.script || '';
            codeEditor.value = rawCode;

            // ✅ 鼠标滚轮横向滚动（首次加载时绑定事件）
            if (!codeEditor._wheelBound) {
                codeEditor._wheelBound = true;
                codeEditor.addEventListener('wheel', function(e) {
                    // 如果按住 Shift 或只有横向滚动条，转为横向滚动
                    if (e.deltaY !== 0 && this.scrollWidth > this.clientWidth) {
                        e.preventDefault();
                        this.scrollLeft += e.deltaY;
                    }
                }, { passive: false });
            }
        }

        // 更新工具栏显示
        setParserTitle(script.name || pt('ocr.parser.unnamed'));
        
        const badge = document.getElementById('parser-builtin-badge');
        if (badge) {
            badge.style.display = script.isBuiltIn ? 'inline' : 'none';
        }

        // 显示/隐藏删除按钮（内置不可删）
        const deleteBtn = document.getElementById('parser-delete-btn');
        if (deleteBtn) {
            deleteBtn.style.display = script.isBuiltIn ? 'none' : 'inline-flex';
        }

        // 显示/隐藏还原按钮（仅内置脚本可还原）
        const resetBtn = document.getElementById('parser-reset-btn');
        if (resetBtn) {
            resetBtn.style.display = script.isBuiltIn ? 'inline-flex' : 'none';
        }

        // 渲染字段列表
        renderFieldsList();

        // 更新列表选中状态
        renderScriptList();

        // 加载示例文本到测试区
        if (script.sampleText) {
            document.getElementById('parser-test-input').value = script.sampleText;
        }

        ChatLog.info('OCR', `[Parser] 已加载脚本: ${script.name}`);
    }

    // ==================== 脚本 CRUD 操作 ====================

    /**
     * 清空缓存，重新从源文件加载内置脚本
     */
    async function clearCache() {
        if (!confirm('确定要清空缓存吗？\n\n这将删除所有自定义脚本，并从源文件重新加载内置脚本。\n此操作不可撤销！')) {
            return;
        }

        try {
            ChatLog.info('OCR', '[Parser] 开始清空缓存...');
            const resultJson = await callOcrApi('ClearParseScriptCache');
            const result = JSON.parse(resultJson);
            
            if (result.success) {
                ChatLog.info('OCR', '[Parser] 缓存已清空，重新加载脚本...');
                showToast('缓存已清空，正在重新加载脚本...');
                // 重新加载脚本
                await loadScripts();
                // 清空当前选中状态和编辑器
                _currentScript = null;
                document.getElementById('parser-script-name').value = '';
                document.getElementById('parser-script-desc').value = '';
                document.getElementById('parser-script-code').value = '';
                setParserTitle('');
                document.getElementById('parser-builtin-badge').style.display = 'none';
                renderFieldsList();
            } else {
                ChatLog.error('OCR', '[Parser] 清空缓存失败:', result.error);
                showToast('清空缓存失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 清空缓存异常:', ex);
            showToast('清空缓存失败: ' + ex.message);
        }
    }

    /**
     * 创建新的空白脚本
     */
    function createNewScript() {
        const newScript = {
            id: generateId(),
            name: '新建解析脚本',
            description: '',
            icon: '📝',
            category: '自定义',
            isBuiltIn: false,
            version: 1,
            createdAt: new Date().toISOString(),
            lastModified: new Date().toISOString(),
            fields: [],
            script: `// 新建解析脚本
function parse(text) {
    const result = {};
    
    // 在此编写你的解析逻辑
    // 可用变量:
    //   - text: OCR识别的原始文本
    //   - FIELDS: 字段定义数组
    
    // 示例：使用正则提取
    // for (const field of FIELDS) {
    //     if (field.pattern) {
    //         const regex = new RegExp(field.pattern);
    //         const match = text.match(regex);
    //         if (match && match[1]) {
    //             result[field.key] = match[1].trim();
    //         }
    //     }
    // }
    
    return result;
}`,
            outputFormat: 'table',
            sampleText: '',
            tags: [],
            usageCount: 0
        };

        loadScriptToEditor(newScript);
        switchView('edit');

        // 自动聚焦到名称输入框
        setTimeout(() => {
            document.getElementById('parser-script-name').focus();
            document.getElementById('parser-script-name').select();
        }, 100);

        showToast('已创建新脚本，请编辑后保存');
    }

    /**
     * 上传 MD 文件导入解析脚本
     */
    function importScriptFromFile() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.md,.markdown';
        input.onchange = async function(e) {
            const file = e.target.files[0];
            if (!file) return;

            try {
                const reader = new FileReader();
                reader.onload = async function(event) {
                    const markdownContent = event.target.result;
                    
                    try {
                        const resultJson = await callOcrApi('ImportParseScriptFromMarkdown', [markdownContent]);
                        const result = JSON.parse(resultJson);
                        
                        if (result.success && result.script) {
                            // 添加到本地缓存
                            _scripts.push(result.script);

                            // 调试日志：检查导入的脚本数据
                            ChatLog.info('OCR', '[Parser] 导入的完整脚本对象:', JSON.stringify(result.script, null, 2));
                            ChatLog.info('OCR', '[Parser] 脚本代码长度:', (result.script.script || '').length);
                            ChatLog.info('OCR', '[Parser] 示例文本长度:', (result.script.sampleText || '').length);

                            // 加载到编辑器
                            loadScriptToEditor(result.script);
                            switchView('edit');
                            
                            // 刷新列表
                            renderScriptList();
                            
                            showToast('✅ 脚本导入成功');
                            ChatLog.info('OCR', '[Parser] 脚本已导入:', result.script.name);
                        } else {
                            showToast('导入失败: ' + (result.error || '文件格式不正确'));
                        }
                    } catch (ex) {
                        ChatLog.error('OCR', '[Parser] 导入失败:', ex);
                        showToast('导入失败: ' + ex.message);
                    }
                };
                reader.readAsText(file, 'UTF-8');
            } catch (ex) {
                ChatLog.error('OCR', '[Parser] 读取文件失败:', ex);
                showToast('读取文件失败: ' + ex.message);
            }
        };
        input.click();
    }

    /**
     * 保存当前脚本
     */
    async function saveCurrent() {
        if (!_currentScript) {
            showToast('请先选择或创建脚本');
            return;
        }

        // 收集表单数据
        const name = document.getElementById('parser-script-name').value.trim();
        if (!name) {
            showToast('请输入脚本名称');
            return;
        }

        const scriptCode = document.getElementById('parser-script-code').value.trim();
        if (!scriptCode) {
            showToast('请输入JavaScript解析脚本');
            return;
        }

        // 验证脚本语法
        if (!validateScriptSyntax(scriptCode)) {
            showToast('脚本语法错误，请检查代码');
            return;
        }

        // 构建保存对象
        _currentScript.name = name;
        _currentScript.description = document.getElementById('parser-script-desc').value.trim();
        _currentScript.icon = document.getElementById('parser-script-icon').value.trim() || '📝';
        _currentScript.category = document.getElementById('parser-script-category').value;
        _currentScript.outputFormat = document.getElementById('parser-script-output-format').value;

        // ✅ 保存编辑器中的代码（可能是格式化后的）
        // 同时保存原始未格式化版本用于执行
        _currentScript.script = scriptCode;
        _currentScript._rawScript = scriptCode;  // 保存为原始代码
        _currentScript.lastModified = new Date().toISOString();

        try {
            const scriptJson = JSON.stringify(_currentScript);
            const resultJson = await callOcrApi('SaveParseScript', [scriptJson]);
            const result = JSON.parse(resultJson);

            if (result.success) {
                // 更新本地缓存
                const existingIdx = _scripts.findIndex(s => s.id === _currentScript.id);
                if (existingIdx >= 0) {
                    _scripts[existingIdx] = _currentScript;
                } else {
                    _scripts.push(_currentScript);
                }

                // 更新UI
                setParserTitle(name);
                renderScriptList();

                showToast('✅ 脚本保存成功');
                ChatLog.info('OCR', '[Parser] 脚本已保存:', name);
            } else {
                showToast('保存失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 保存失败:', ex);
            showToast('保存失败: ' + ex.message);
        }
    }

    /**
     * 复制当前脚本
     */
    async function duplicateCurrent() {
        if (!_currentScript) {
            showToast('请先选择要复制的脚本');
            return;
        }

        const newName = prompt('输入新脚本的名称:', _currentScript.name + ' (副本)');
        if (!newName || !newName.trim()) return;

        try {
            const resultJson = await callOcrApi('DuplicateParseScript', [_currentScript.id, newName.trim()]);
            const result = JSON.parse(resultJson);

            if (result.success && result.script) {
                _scripts.push(result.script);
                loadScriptToEditor(result.script);
                showToast('✅ 脚本复制成功');
            } else {
                showToast('复制失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            showToast('复制失败: ' + ex.message);
        }
    }

    /**
     * 删除当前脚本
     */
    async function deleteCurrent() {
        if (!_currentScript) {
            showToast('请先选择要删除的脚本');
            return;
        }

        if (_currentScript.isBuiltIn) {
            showToast('内置模板不可删除');
            return;
        }

        if (!confirm(`确定要删除脚本 "${_currentScript.name}" 吗？\n此操作不可撤销。`)) {
            return;
        }

        try {
            const resultJson = await callOcrApi('DeleteParseScript', [_currentScript.id]);
            const result = JSON.parse(resultJson);

            if (result.success) {
                _scripts = _scripts.filter(s => s.id !== _currentScript.id);
                _currentScript = null;

                // 清空编辑器
                document.getElementById('parser-script-name').value = '';
                document.getElementById('parser-script-desc').value = '';
                document.getElementById('parser-script-code').value = '';
                setParserTitle('');
                document.getElementById('parser-builtin-badge').style.display = 'none';
                document.getElementById('parser-delete-btn').style.display = 'none';

                renderScriptList();
                renderFieldsList();
                showToast('✅ 脚本已删除');
            } else {
                showToast('删除失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            showToast('删除失败: ' + ex.message);
        }
    }

    /**
     * 根据ID删除脚本（从列表中的删除按钮调用）
     */
    async function deleteCurrentById(scriptId) {
        const script = _scripts.find(s => s.id === scriptId);
        if (!script) {
            showToast('脚本不存在');
            return;
        }

        if (script.isBuiltIn) {
            showToast('内置模板不可删除');
            return;
        }

        if (!confirm(`确定要删除脚本 "${script.name}" 吗？\n此操作不可撤销。`)) {
            return;
        }

        try {
            const resultJson = await callOcrApi('DeleteParseScript', [scriptId]);
            const result = JSON.parse(resultJson);

            if (result.success) {
                _scripts = _scripts.filter(s => s.id !== scriptId);
                // 如果删除的是当前选中的脚本，清空编辑器
                if (_currentScript && _currentScript.id === scriptId) {
                    _currentScript = null;
                    document.getElementById('parser-script-name').value = '';
                    document.getElementById('parser-script-desc').value = '';
                    document.getElementById('parser-script-code').value = '';
                    setParserTitle('');
                    document.getElementById('parser-builtin-badge').style.display = 'none';
                    document.getElementById('parser-delete-btn').style.display = 'none';
                    renderFieldsList();
                }
                renderScriptList();
                showToast('✅ 脚本已删除');
            } else {
                showToast('删除失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            showToast('删除失败: ' + ex.message);
        }
    }

    /**
     * 还原当前内置脚本到默认配置
     */
    async function resetCurrent() {
        if (!_currentScript) {
            showToast('请先选择要还原的脚本');
            return;
        }

        if (!_currentScript.isBuiltIn) {
            showToast('仅内置脚本可还原为默认配置');
            return;
        }

        if (!confirm(`确定要将 "${_currentScript.name}" 还原为默认配置吗？\n所有自定义修改将丢失。`)) {
            return;
        }

        try {
            const resultJson = await callOcrApi('ResetBuiltinScript', [_currentScript.id]);
            const result = JSON.parse(resultJson);

            if (result.success && result.script) {
                // 用还原后的脚本替换当前脚本
                _currentScript = result.script;

                // 更新本地缓存
                const existingIdx = _scripts.findIndex(s => s.id === _currentScript.id);
                if (existingIdx >= 0) {
                    _scripts[existingIdx] = _currentScript;
                }

                // 重新加载到编辑器
                loadScriptToEditor(_currentScript);

                showToast('✅ 已还原为默认配置');
                ChatLog.info('OCR', '[Parser] 脚本已还原:', _currentScript.name);
            } else {
                showToast('还原失败: ' + (result.error || '未知错误'));
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 还原失败:', ex);
            showToast('还原失败: ' + ex.message);
        }
    }

    // ==================== 视图切换 ====================

    /**
     * 切换编辑/预览/字段视图
     */
    function switchView(viewName) {
        _currentView = viewName;

        // 更新标签按钮状态
        document.querySelectorAll('#tab-parser .parser-view-tabs .view-tab-btn').forEach(btn => {
            btn.classList.toggle('active', btn.dataset.parserView === viewName);
        });

        // 切换内容区域
        document.querySelectorAll('#tab-parser .parser-view-content').forEach(content => {
            content.classList.remove('active');
        });

        const targetView = document.getElementById(`parser-${viewName}-view`);
        if (targetView) {
            targetView.classList.add('active');
        }
    }

    // ==================== 字段配置 ====================

    /**
     * 添加新字段
     */
    function addField() {
        if (!_currentScript) {
            showToast('请先选择或创建脚本');
            return;
        }

        if (!_currentScript.fields) {
            _currentScript.fields = [];
        }

        const newField = {
            name: '',
            key: '',
            pattern: '',
            required: false,
            defaultValue: '',
            visible: true
        };

        _currentScript.fields.push(newField);
        renderFieldsList();
        switchView('fields');
        onFieldConfigEdited();

        // 滚动到最后一个字段
        setTimeout(() => {
            const list = document.getElementById('parser-fields-list');
            if (list) list.scrollTop = list.scrollHeight;
        }, 50);
    }

    /**
     * 更新字段属性
     */
    function updateField(index, prop, value) {
        if (!_currentScript || !_currentScript.fields) return;
        if (index >= 0 && index < _currentScript.fields.length) {
            _currentScript.fields[index][prop] = value;
            onFieldConfigEdited();
        }
    }

    /**
     * 删除指定索引的字段
     */
    function removeField(index) {
        if (!_currentScript || !_currentScript.fields) return;
        _currentScript.fields.splice(index, 1);
        renderFieldsList();
        onFieldConfigEdited();
    }

    /**
     * 渲染字段配置列表
     */
    function renderFieldsList() {
        const container = document.getElementById('parser-fields-list');
        if (!container) return;

        if (!_currentScript || !_currentScript.fields || _currentScript.fields.length === 0) {
            container.innerHTML = '<div class="parser-empty-fields">' + pt('ocr.parser.noFields') + '</div>';
            return;
        }

        container.innerHTML = _currentScript.fields.map((field, idx) => `
            <div class="parser-field-item">
                <div class="parser-field-row">
                    <input type="text" placeholder="字段名称" value="${escapeHtml(field.name || '')}"
                           style="min-width:80px;flex:2 1 auto" onchange="App.Parser.updateField(${idx}, 'name', this.value)">
                    <input type="text" placeholder="键名(英文)" value="${escapeHtml(field.key || '')}"
                           style="min-width:70px;flex:1.5 1 auto" onchange="App.Parser.updateField(${idx}, 'key', this.value)">
                    <label style="display:flex;align-items:center;gap:3px;font-size:11px;color:#666;white-space:nowrap;">
                        <input type="checkbox" ${field.required ? 'checked' : ''}
                               onchange="App.Parser.updateField(${idx}, 'required', this.checked)"> 必填
                    </label>
                    <label style="display:flex;align-items:center;gap:3px;font-size:11px;color:#2e7d32;white-space:nowrap;" title="勾选后该字段在解析结果中显示">
                        <input type="checkbox" ${field.visible !== false ? 'checked' : ''}
                               onchange="App.Parser.updateField(${idx}, 'visible', this.checked)"> 显示
                    </label>
                    <button class="btn-danger btn-sm" onclick="App.Parser.removeField(${idx})">✕</button>
                </div>
                <div class="parser-field-row">
                    <input type="text" placeholder="正则表达式 (可选)" value="${escapeHtml(field.pattern || '')}"
                           style="min-width:120px;flex:3 1 auto" onchange="App.Parser.updateField(${idx}, 'pattern', this.value)">
                    <input type="text" placeholder="默认值" value="${escapeHtml(field.defaultValue || '')}"
                           style="min-width:60px;flex:1 1 auto" onchange="App.Parser.updateField(${idx}, 'defaultValue', this.value)">
                </div>
            </div>
        `).join('');
    }

    // ==================== 解析执行 ====================

    /**
     * 执行解析（从测试预览界面）
     */
    async function runParse() {
        if (!_currentScript) {
            showToast('请先选择解析脚本');
            return;
        }

        const testInput = document.getElementById('parser-test-input');
        const ocrText = testInput ? testInput.value.trim() : '';

        if (!ocrText) {
            showToast('请输入测试文本（OCR识别结果）');
            return;
        }

        await executeParse(_currentScript, ocrText);
    }

    /**
     * 执行解析（从识别Tab调用）
     */
    async function parseOcrResult(ocrText) {
        if (!_currentScript) {
            showToast('请先在"智能解析"Tab中选择一个解析脚本');
            return null;
        }

        return await executeParse(_currentScript, ocrText);
    }

    /**
     * 核心解析执行函数
     */
    async function executeParse(script, ocrText, target) {
        const resultContainer = document.getElementById('parser-result-container');
        if (resultContainer) {
            resultContainer.innerHTML = '<div style="text-align:center;padding:30px;color:#666;">' + pt('ocr.parser.parsing') + '</div>';
        }

        try {
            // 准备字段数据
            const fields = script.fields || [];

            // ✅ 使用原始未格式化的代码执行（避免格式化导致的语法问题）
            const rawScriptCode = script._rawScript || script.script || '';

            if (!rawScriptCode.trim()) {
                throw new Error('脚本代码为空');
            }

            // 构建并执行JavaScript
            const parseFunc = createParseFunction(rawScriptCode, fields);

            if (!parseFunc) {
                throw new Error('无法创建解析函数，请检查脚本语法');
            }

            // 执行解析
            // 预处理：从OCR原始格式中提取纯文本（如 [1] text="姓名韦小宝" ... → 姓名韦小宝）
            let cleanText = ocrText;
            if (/\[\d+\]\s*text=/.test(ocrText)) {
                const textMatches = ocrText.match(/\[\d+\]\s*text="([^"]*)"/g);
                if (textMatches) {
                    cleanText = textMatches.map(function(m) {
                        var match = m.match(/text="([^"]*)"/);
                        return match ? match[1] : '';
                    }).filter(function(t) { return t; }).join('\n');
                }
            }

            // 如果脚本可能需要坐标，尝试获取 OCR 元素列表
            let elements = null;
            try {
                if (excelBridgeReady() || typeof window.ocrHostCall === 'function') {
                    const elementsJson = await callOcrApi('GetOcrElementsJson');
                    const parsed = JSON.parse(elementsJson);
                    if (parsed && parsed.success && Array.isArray(parsed.items)) {
                        elements = parsed.items;
                    }
                }
            } catch (e) {
                ChatLog.warn('OCR', '[Parser] 获取 OCR 坐标失败:', e.message);
            }

            const startTime = performance.now();
            const result = await parseFunc(cleanText, elements);
            const elapsed = Math.round(performance.now() - startTime);

            _lastResult = result;
            _lastParseElapsed = elapsed;

            // 显示结果（根据调用来源选择输出位置）
            if (target === 'main') {
                _mainParseElapsed = elapsed;
                displayMainResult(result, script.outputFormat, elapsed);
            } else {
                displayResult(result, script.outputFormat, elapsed);
            }

            ChatLog.info('OCR', `[Parser] 解析完成 (${elapsed}ms):`, result);
            return result;

        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 解析执行错误:', ex);
            
            if (resultContainer) {
                // 支持多行错误信息显示
                const formattedError = ex.message.replace(/\n/g, '<br>');
                resultContainer.innerHTML = `
                    <div style="color:#d32f2f;padding:20px;text-align:left;max-width:600px;margin:0 auto;">
                        <div style="font-size:24px;margin-bottom:8px;text-align:center;">⚠️</div>
                        <div style="font-weight:600;margin-bottom:12px;text-align:center;">解析执行出错</div>
                        <div style="font-size:12px;background:#fff3f3;padding:12px;border-radius:6px;border-left:3px solid #d32f2f;line-height:1.6;">${formattedError}</div>
                    </div>`;
            }

            // 只在 toast 中显示简短错误
            const shortError = ex.message.split('\n')[0];
            showToast('解析出错: ' + shortError);
            return null;
        }
    }

    /**
     * 创建可执行的解析函数
     */
    function createParseFunction(scriptCode, fields) {
        try {
            // 验证脚本代码不为空
            if (!scriptCode || !scriptCode.trim()) {
                throw new Error('脚本代码为空，请编写解析逻辑');
            }

            ChatLog.info('OCR', '[Parser] 脚本代码长度:', scriptCode.length);
            
            // 使用 Function 构造器创建函数（比 eval 更安全可控）
            let parseFunc;
            
            // 策略1: 尝试直接作为函数表达式执行
            try {
                const wrappedAsExpression = `
                    "use strict";
                    const FIELDS = ${JSON.stringify(fields)};
                    (${scriptCode})
                `;
                
                const func1 = new Function(wrappedAsExpression);
                const result1 = func1();
                
                if (typeof result1 === 'function') {
                    ChatLog.info('OCR', '[Parser] ✓ 作为函数表达式成功');
                    return result1;
                }
                
                // 如果不是函数但也不是undefined，可能是对象等
                if (result1 !== undefined) {
                    ChatLog.info('OCR', '[Parser] ✓ 返回非函数值，自动包装');
                    return (text) => result1;
                }
            } catch (ex1) {
                ChatLog.info('OCR', '[Parser] 函数表达式方式失败:', ex1.message);
            }
            
            // 策略2: 处理函数声明格式 (function parse(text) {...})
            try {
                const wrappedAsDeclaration = `
                    "use strict";
                    const FIELDS = ${JSON.stringify(fields)};
                    ${scriptCode};
                    if (typeof parse === 'function') return parse;
                    // 尝试其他常见函数名
                    var funcs = [parse, extract, run, execute, main, handler];
                    for (var i = 0; i < funcs.length; i++) {
                        if (typeof funcs[i] === 'function') return funcs[i];
                    }
                `;
                
                const func2 = new Function(wrappedAsDeclaration);
                const result2 = func2();
                
                if (typeof result2 === 'function') {
                    ChatLog.info('OCR', '[Parser] ✓ 作为函数声明成功');
                    return result2;
                }
            } catch (ex2) {
                ChatLog.info('OCR', '[Parser] 函数声明方式失败:', ex2.message);
            }
            
            // 策略3: 将函数声明转换为函数表达式
            try {
                // 移除 function name(text) 中的 name，转换为匿名函数
                let convertedCode = scriptCode.trim();
                
                // 匹配 function xxx( 或 function xxx (
                convertedCode = convertedCode.replace(
                    /^function\s+\w*\s*\(/, 
                    'function('
                );
                
                const wrappedConverted = `
                    "use strict";
                    const FIELDS = ${JSON.stringify(fields)};
                    (${convertedCode})
                `;
                
                const func3 = new Function(wrappedConverted);
                const result3 = func3();
                
                if (typeof result3 === 'function') {
                    ChatLog.info('OCR', '[Parser] ✓ 转换为匿名函数成功');
                    return result3;
                }
            } catch (ex3) {
                ChatLog.info('OCR', '[Parser] 转换方式失败:', ex3.message);
            }
            
            // 所有策略都失败
            throw new Error(`无法创建解析函数

已尝试的格式:
1. 函数表达式: (text) => {...}
2. 函数声明: function parse(text) {...}
3. 匿名函数: function(text) {...}

请检查脚本语法，确保返回一个函数。

推荐格式:
function parse(text) {
    return { 字段: 值 };
}`);
            
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 创建解析函数失败:', ex);
            throw ex;
        }
    }

    /**
     * 获取当前脚本中标记为可见的字段key集合（用于过滤结果显示）
     */
    function getVisibleFieldKeys() {
        const visibleKeys = new Set();
        if (_currentScript && _currentScript.fields) {
            _currentScript.fields.forEach(f => {
                if (f.visible !== false && f.key) visibleKeys.add(f.key);
            });
        }
        return visibleKeys;
    }

    /**
     * 按字段配置整理解析结果：顺序、中文名、显示开关都跟着配置走。
     * 脚本另写的 birth_date_from_id 这类值，归到已配置的 birth_date，不再单独占一行。
     * 配置里没有的键仍保留在最后，避免脚本算出的额外内容被丢掉。
     */
    function presentByFieldConfig(result, fields) {
        const data = {};
        const names = {};

        function projectRowArray(rows, fieldList) {
            if (!Array.isArray(rows) || !rows.length || !fieldList || !fieldList.length) return rows;
            const isRows = rows.every(function (row) {
                return row && typeof row === 'object' && !Array.isArray(row);
            });
            if (!isRows) return rows;
            const present = {};
            rows.forEach(function (row) {
                Object.keys(row).forEach(function (key) { present[key] = true; });
            });
            const cols = fieldList.filter(function (field) { return field && field.key && present[field.key]; });
            if (!cols.length) return rows;
            const visible = cols.filter(function (field) { return field.visible !== false; });
            if (!visible.length) return [];
            return rows.map(function (row) {
                const out = {};
                visible.forEach(function (field) {
                    const value = row[field.key];
                    out[field.name || field.key] = value === undefined || value === null ? '' : value;
                });
                return out;
            });
        }

        if (!result || typeof result !== 'object' || Array.isArray(result)) {
            return { data: data, names: names };
        }
        if (!fields || !fields.length) {
            Object.keys(result).forEach(function (key) {
                if (key === 'tableData') return;
                const value = result[key];
                if (value === undefined || value === null || value === '') return;
                data[key] = value;
            });
            return { data: data, names: names };
        }
        const used = {};
        fields.forEach(function (field) {
            if (!field || !field.key) return;
            const alias = field.key + '_from_id';
            used[field.key] = true;
            used[alias] = true;
            if (field.name) names[field.key] = field.name;
            if (field.visible === false) return;
            let value = result[field.key];
            if (value === undefined || value === null || value === '') value = result[alias];
            if (value === undefined || value === null || value === '') return;
            value = projectRowArray(value, fields);
            if (Array.isArray(value) && value.length === 0) return;
            data[field.key] = value;
        });
        Object.keys(result).forEach(function (key) {
            if (used[key] || key === 'tableData') return;
            let value = result[key];
            if (value === undefined || value === null || value === '') return;
            value = projectRowArray(value, fields);
            if (Array.isArray(value) && value.length === 0) return;
            data[key] = value;
        });
        return { data: data, names: names };
    }

    function projectParseResult(result, script) {
        return presentByFieldConfig(result, script && script.fields);
    }

    /** 明细列名已经换成字段配置里的名称时，按这个表写入，不再用解析时的旧表头。 */
    function tableRowsFromProjected(displayData, fields) {
        const names = {};
        (fields || []).forEach(function (field) {
            if (field && field.name) names[field.name] = true;
        });
        if (!Object.keys(names).length) return null;
        const keys = Object.keys(displayData || {});
        for (let i = 0; i < keys.length; i++) {
            const val = displayData[keys[i]];
            if (!Array.isArray(val) || !val.length || !val[0] || typeof val[0] !== 'object' || Array.isArray(val[0])) continue;
            const cols = Object.keys(val[0]);
            if (!cols.length || !cols.every(function (col) { return names[col]; })) continue;
            const rows = [cols];
            val.forEach(function (item) {
                rows.push(cols.map(function (col) {
                    const cell = item[col];
                    return cell === undefined || cell === null ? '' : String(cell);
                }));
            });
            return rows;
        }
        return null;
    }

    /** 字段名、显示、顺序或正则改了之后，立刻按新配置重画已有结果。 */
    function onFieldConfigEdited() {
        const mainScript = _scripts.find(function (s) { return s.id === _mainScriptId; });
        const editingMain = mainScript && _currentScript && mainScript.id === _currentScript.id;
        let ocrText = '';
        if (typeof OcrState !== 'undefined' && OcrState.recognitionResult) {
            ocrText = String(OcrState.recognitionResult).trim();
        }
        if (!ocrText) {
            const ta = document.getElementById('ocr-result-text');
            ocrText = ta ? ta.value.trim() : '';
        }
        if (editingMain && ocrText) {
            runParseFromMain();
            return;
        }
        const testInput = document.getElementById('parser-test-input');
        if (_currentScript && testInput && testInput.value.trim()) {
            runParse();
            return;
        }
        if (_lastResult && _currentScript) {
            displayResult(_lastResult, _currentScript.outputFormat || 'table', _lastParseElapsed || 0);
        }
        if (_mainLastResult && mainScript) {
            displayMainResult(_mainLastResult, mainScript.outputFormat || 'table', _mainParseElapsed || 0);
        }
    }

    /**
     * 显示解析结果
     */
    function displayResult(result, format, elapsedMs) {
        const container = document.getElementById('parser-result-container');
        if (!container) return;

        if (!result || typeof result !== 'object' || Object.keys(result).length === 0) {
            container.innerHTML = `
                <div style="text-align:center;padding:30px;color:#999;">
                    <div style="font-size:24px;margin-bottom:8px;">📭</div>
                    <div>${pt('ocr.parser.noData')}</div>
                    <div style="font-size:11px;color:#bbb;margin-top:4px;">请检查脚本逻辑和字段定义</div>
                </div>`;
            return;
        }

        // 按字段配置的顺序、名称和显示开关整理。派生键并入对应字段。
        const projected = projectParseResult(result, _currentScript);
        const fieldNameMap = projected.names;
        const displayResult = projected.data;

        if (!displayResult || Object.keys(displayResult).length === 0) {
            container.innerHTML = `
                <div style="text-align:center;padding:30px;color:#999;">
                    <div style="font-size:24px;margin-bottom:8px;">📭</div>
                    <div>${pt('ocr.parser.noData')}</div>
                    <div style="font-size:11px;color:#bbb;margin-top:4px;">请检查脚本逻辑和字段定义</div>
                </div>`;
            return;
        }

        let html = '';

        // 根据输出格式渲染
        switch (format) {
            case 'table':
                html = renderTableResult(displayResult, fieldNameMap);
                break;
            case 'key-value':
                html = renderKeyValueResult(displayResult, fieldNameMap);
                break;
            case 'json':
                html = renderJsonResult(displayResult);
                break;
            default:
                html = renderTableResult(displayResult, fieldNameMap);
        }

        container.innerHTML = html +
            `<div style="margin-top:10px;text-align:right;font-size:11px;color:#999;">
                解析耗时: ${elapsedMs}ms | 字段数: ${Object.keys(displayResult).length}
            </div>`;
    }

    /**
     * 表格格式结果
     */
    function renderTableResult(result, fieldNameMap) {
        const entries = Object.entries(result);
        let html = '<table class="parse-result-table"><thead><tr><th>字段名</th><th>值</th></tr></thead><tbody>';

        entries.forEach(([key, value]) => {
            const displayName = fieldNameMap[key] || key;
            const formattedValue = formatValueForDisplay(value, key);
            html += `<tr><td><strong>${escapeHtml(displayName)}</strong></td><td>${formattedValue}</td></tr>`;
        });
        
        html += '</tbody></table>';
        return html;
    }

    /**
     * 键值对格式结果
     */
    function renderKeyValueResult(result, fieldNameMap) {
        let html = '<div class="parse-kv-list">';

        Object.entries(result).forEach(([key, value]) => {
            const displayName = fieldNameMap[key] || key;
            const formattedValue = formatValueForDisplay(value, key);
            html += `
                <div class="parse-kv-item">
                    <span class="parse-kv-key">${escapeHtml(displayName)}</span>
                    <span class="parse-kv-value">${formattedValue}</span>
                </div>`;
        });
        
        html += '</div>';
        return html;
    }

    /**
     * 格式化值为纯文本（用于复制到剪贴板）
     */
    function formatValueForText(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'object') return JSON.stringify(value);
        return String(value);
    }

    /**
     * 智能格式化值用于显示
     * 处理各种数据类型：字符串、数字、数组、对象等
     */
    function formatValueForDisplay(value, fieldName) {
        // 处理 null 或 undefined
        if (value === null || value === undefined) {
            return '<span style="color:#999;font-style:italic;">(空)</span>';
        }
        
        // 处理字符串和数字（简单值）
        if (typeof value === 'string') {
            // 长文本截断显示
            if (value.length > 100) {
                return `<span title="${escapeHtml(value)}">${escapeHtml(value.substring(0, 100))}...</span>`;
            }
            return escapeHtml(value);
        }
        
        if (typeof value === 'number' || typeof value === 'boolean') {
            return escapeHtml(String(value));
        }
        
        // 处理结构化表格数据 { header, rows }
        if (value && typeof value === 'object' && Array.isArray(value.header) && Array.isArray(value.rows)) {
            return renderTableData(value, fieldName);
        }

        // 处理数组
        if (Array.isArray(value)) {
            if (value.length === 0) {
                return '<span style="color:#999;">[] 空数组</span>';
            }
            
            // 检查是否是对象数组（如 amounts, dates）
            const isObjectArray = typeof value[0] === 'object' && value[0] !== null;
            
            if (isObjectArray) {
                // 对象数组 - 渲染为子表格
                return renderObjectArrayAsTable(value, fieldName);
            } else {
                // 简单数组 - 渲染为逗号分隔的列表
                const items = value.map(item => escapeHtml(String(item))).join(', ');
                return `<span style="color:#0066cc;">[${items}]</span>`;
            }
        }
        
        // 处理普通对象
        if (typeof value === 'object') {
            const entries = Object.entries(value);
            
            // 简单对象 - 显示为键值对
            if (entries.length <= 4) {
                const kvPairs = entries.map(([k, v]) => 
                    `<span style="color:#666;">${escapeHtml(k)}:</span>${escapeHtml(String(v))}`
                ).join(' | ');
                return `<div style="font-size:12px;line-height:1.6;">${kvPairs}</div>`;
            }
            
            // 复杂对象 - 显示为 JSON
            return `<pre style="margin:0;padding:8px;background:#f5f5f5;border-radius:4px;font-size:11px;max-height:200px;overflow:auto;white-space:pre-wrap;">${escapeHtml(JSON.stringify(value, null, 2))}</pre>`;
        }
        
        // 其他类型
        return escapeHtml(String(value));
    }

    /**
     * 将 { header, rows } 结构化表格数据渲染为 HTML 表格
     */
    function renderTableData(tableData, fieldName) {
        if (!tableData || !Array.isArray(tableData.header) || !Array.isArray(tableData.rows)) {
            return '<span style="color:#999;">表格数据格式错误</span>';
        }

        const header = tableData.header;
        const rows = tableData.rows;

        if (header.length === 0 && rows.length === 0) {
            return '<span style="color:#999;">空表格</span>';
        }

        let html = `<div style="margin-top:4px;">
            <div style="font-size:11px;color:#666;margin-bottom:4px;">共 ${rows.length} 行 × ${header.length} 列</div>
            <table style="width:100%;border-collapse:collapse;font-size:11px;background:white;border:1px solid #e0e0e0;border-radius:4px;">
                <thead>
                    <tr style="background:#f8f9fa;">`;

        header.forEach(function(cell) {
            html += `<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #e0e0e0;color:#555;font-weight:600;">${escapeHtml(String(cell))}</th>`;
        });

        html += `</tr>
                </thead>
                <tbody>`;

        rows.forEach(function(row, idx) {
            html += `<tr style="${idx % 2 === 0 ? '' : 'background:#fafbfc'}">`;
            row.forEach(function(cell) {
                const val = cell === null || cell === undefined ? '' : String(cell);
                html += `<td style="padding:6px 8px;border-bottom:1px solid #f0f0f0;">${escapeHtml(val)}</td>`;
            });
            html += `</tr>`;
        });

        html += `</tbody></table></div>`;
        return html;
    }

    /**
     * 将对象数组渲染为表格（用于 amounts, dates 等复杂数据）
     */
    function renderObjectArrayAsTable(arr, fieldName) {
        if (!arr || arr.length === 0) return '';
        
        // 获取所有可能的键
        const allKeys = new Set();
        arr.forEach(item => {
            if (typeof item === 'object' && item !== null) {
                Object.keys(item).forEach(key => allKeys.add(key));
            }
        });
        
        const keys = Array.from(allKeys);
        
        // 根据字段名选择要显示的列
        let displayKeys = keys;
        if (fieldName === 'amounts') {
            displayKeys = keys.filter(k => ['value', 'formatted', 'raw', 'currency', 'type'].includes(k));
        } else if (fieldName === 'dates') {
            displayKeys = keys.filter(k => ['date', 'raw', 'type'].includes(k));
        }
        
        if (displayKeys.length === 0) displayKeys = keys.slice(0, 6); // 最多显示6列
        
        let html = `<div style="margin-top:4px;">
            <div style="font-size:11px;color:#666;margin-bottom:4px;">共 ${arr.length} 条记录</div>
            <table style="width:100%;border-collapse:collapse;font-size:11px;background:white;border:1px solid #e0e0e0;border-radius:4px;">
                <thead>
                    <tr style="background:#f8f9fa;">
                        ${displayKeys.map(k => `<th style="padding:6px 8px;text-align:left;border-bottom:1px solid #e0e0e0;color:#555;font-weight:600;">${escapeHtml(k)}</th>`).join('')}
                    </tr>
                </thead>
                <tbody>`;
        
        arr.forEach((item, idx) => {
            html += `<tr style="${idx % 2 === 0 ? '' : 'background:#fafbfc'}">`;
            displayKeys.forEach(key => {
                const val = item[key];
                let displayVal;
                
                if (val === null || val === undefined) {
                    displayVal = '<span style="color:#ccc;">-</span>';
                } else if (typeof val === 'number') {
                    // 数字右对齐，保留合适的小数位
                    displayVal = `<span style="float:right;">${Number(val).toLocaleString()}</span>`;
                } else {
                    displayVal = escapeHtml(String(val));
                }
                
                html += `<td style="padding:6px 8px;border-bottom:1px solid #f0f0f0;">${displayVal}</td>`;
            });
            html += `</tr>`;
        });
        
        html += `</tbody></table></div>`;
        
        return html;
    }

    /**
     * JSON格式结果
     */
    function renderJsonResult(result) {
        const jsonStr = JSON.stringify(result, null, 2);
        return `<pre style="margin:0;padding:12px;background:white;border-radius:4px;font-size:12px;line-height:1.5;overflow:auto;">${escapeHtml(jsonStr)}</pre>`;
    }

    // ==================== 测试辅助功能 ====================

    /**
     * 加载当前脚本的示例文本
     */
    function loadSampleText() {
        if (_currentScript && _currentScript.sampleText) {
            document.getElementById('parser-test-input').value = _currentScript.sampleText;
            showToast('已加载示例文本');
        } else {
            showToast('当前脚本没有设置示例文本');
        }
    }

    /**
     * 测试运行（保存并执行）
     */
    async function testScript() {
        if (!_currentScript) {
            showToast('请先选择解析脚本');
            return;
        }

        // 如果编辑器有修改，尝试保存；保存失败仍可继续用已有代码测试
        const editor = document.getElementById('parser-script-code');
        if (editor && editor.value.trim() && editor.value.trim() !== (_currentScript._rawScript || _currentScript.script || '')) {
            await saveCurrent();
        }

        // 切换到预览视图
        switchView('preview');

        // 执行解析
        await runParse();
    }

    // ==================== 结果操作 ====================

    /**
     * 复制解析结果
     */
    function copyResult() {
        if (!_lastResult) {
            showToast('没有可复制的解析结果');
            return;
        }

        const text = JSON.stringify(_lastResult, null, 2);
        copyToClipboard(text);
        showToast('✅ 结果已复制到剪贴板');
    }

    /**
     * 将结果插入Excel
     */
    async function insertToExcel() {
        if (!_lastResult) {
            showToast('没有可插入的数据');
            return;
        }

        try {
            if (window.chrome?.webview?.hostObjects?.excelBridge) {
                const editorProjected = projectParseResult(_lastResult, _currentScript);
                const editorTable = tableRowsFromProjected(editorProjected.data, _currentScript && _currentScript.fields);
                if (editorTable) {
                    const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(editorTable), true);
                    const parsed = JSON.parse(resultObj);
                    if (parsed.success) {
                        showToast('已插入到 Excel');
                    } else {
                        showToast('插入失败: ' + (parsed.error || '未知错误'));
                    }
                    return;
                }
                // 优先处理结构化表格数据 { header, rows }
                if (_lastResult.tableData && Array.isArray(_lastResult.tableData.header) && Array.isArray(_lastResult.tableData.rows)) {
                    const rows = [_lastResult.tableData.header].concat(_lastResult.tableData.rows);
                    if (rows.length === 0 || rows[0].length === 0) {
                        showToast('没有可插入的数据');
                        return;
                    }
                    const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows), true);
                    const parsed = JSON.parse(resultObj);
                    if (parsed.success) {
                        showToast('已插入到 Excel');
                    } else {
                        showToast('插入失败: ' + (parsed.error || '未知错误'));
                    }
                    return;
                }

                const projected = projectParseResult(_lastResult, _currentScript);
                const displayData = projected.data;
                const fieldNameMap = projected.names;

                var rows = [];
                Object.entries(displayData).forEach(function(entry) {
                    var key = entry[0], val = entry[1];
                    var name = fieldNameMap[key] || key;
                    var textVal = formatValueForText(val);
                    rows.push([name, textVal]);
                });

                if (rows.length === 0) { showToast('没有可插入的数据'); return; }

                const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows), true);
                const parsed = JSON.parse(resultObj);
                if (parsed.success) {
                    showToast('已插入到 Excel');
                } else {
                    showToast('插入失败: ' + (parsed.error || '未知错误'));
                }
            } else {
                showToast('WebView2 未就绪');
            }
        } catch (ex) {
            ChatLog.error('OCR', '[Parser] 插入Excel失败:', ex);
            showToast('插入失败: ' + ex.message);
        }
    }

    // ==================== 搜索与筛选 ====================

    /**
     * 搜索脚本
     */
    function onSearch(keyword) {
        const category = document.getElementById('parser-category-select')?.value || '';
        renderScriptList(keyword, category);
    }

    /**
     * 按分类筛选
     */
    function filterByCategory(category) {
        const keyword = document.getElementById('parser-search-input')?.value || '';
        renderScriptList(keyword, category);
    }

    // ==================== 工具函数 ====================

    /**
     * 生成唯一ID
     */
    function generateId() {
        return 'ps_' + Date.now().toString(36) + Math.random().toString(36).substring(2, 8);
    }

    /**
     * HTML转义
     */
    function escapeHtml(str) {
        if (!str) return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    /**
     * 验证JavaScript语法
     */
    function validateScriptSyntax(code) {
        try {
            // 尝试包装成函数并编译
            new Function(code);
            return true;
        } catch (e) {
            ChatLog.warn('OCR', '[Parser] 语法验证失败:', e);
            return false;
        }
    }

    /**
     * 复制到剪贴板
     */
    function copyToClipboard(text) {
        if (navigator.clipboard) {
            navigator.clipboard.writeText(text);
        } else {
            // 兼容旧浏览器
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.style.position = 'fixed';
            textarea.style.opacity = '0';
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            document.body.removeChild(textarea);
        }
    }

    // ==================== 主页（识别Tab）集成 ====================

    /** 主页选中的解析脚本ID */
    let _mainScriptId = null;
    /** 主页最后一次解析结果 */
    let _mainLastResult = null;
    let _lastParseElapsed = 0;
    let _mainParseElapsed = 0;

    /**
     * 填充主页脚本选择下拉框
     */
    function populateMainScriptSelector() {
        const select = document.getElementById('main-parse-script-select');
        if (!select) return;

        select.innerHTML = '<option value="">' + pt('ocr.pickScript') + '</option>';
        
        // 按分类分组
        const categories = {};
        (_scripts || []).forEach(function(s) {
            const cat = s.category || '通用';
            if (!categories[cat]) categories[cat] = [];
            categories[cat].push(s);
        });

        Object.keys(categories).sort().forEach(function(cat) {
            const group = document.createElement('optgroup');
            group.label = cat;
            categories[cat].forEach(function(s) {
                const opt = document.createElement('option');
                opt.value = s.id;
                opt.textContent = (s.icon || '') + ' ' + s.name;
                if (s.id === _mainScriptId) opt.selected = true;
                group.appendChild(opt);
            });
            select.appendChild(group);
        });
    }

    /**
     * 主页脚本选择变化
     */
    function onMainScriptChange(scriptId) {
        _mainScriptId = scriptId || null;

        const btnParse = document.getElementById('btn-main-parse');
        if (btnParse) btnParse.disabled = !scriptId;

        if (scriptId) {
            const script = _scripts.find(function(s) { return s.id === scriptId; });
            if (script) ChatLog.info('OCR', '[Parser-主页] 选择脚本:', script.name);
        }
    }

    /**
     * 从主页执行解析（使用识别结果文本）
     */
    async function runParseFromMain() {
        if (!_mainScriptId) {
            showToast('请先选择解析脚本');
            return;
        }

        // 获取OCR识别结果文本
        // 优先使用原始 OCR 文本（避免被格式下拉框的 HTML/Markdown 转换污染）
        const rawOcrText = (typeof OcrState !== 'undefined' && OcrState.recognitionResult)
            ? OcrState.recognitionResult
            : (document.getElementById('ocr-result-text')?.value.trim() || '');

        if (!rawOcrText) {
            showToast('请先执行OCR识别，获取识别结果后再解析');
            return;
        }

        const script = _scripts.find(function(s) { return s.id === _mainScriptId; });
        if (!script) {
            showToast('未找到选择的脚本');
            return;
        }

        const resultContainer = document.getElementById('main-parse-result');
        if (resultContainer) {
            resultContainer.innerHTML = '<div style="text-align:center;padding:20px;color:#666;">' + pt('ocr.parser.parsing') + '</div>';
        }

        try {
            _mainLastResult = null;
            const result = await executeParse(script, rawOcrText, 'main');
            _mainLastResult = result;

            // 显示操作按钮
            var copyBtn = document.getElementById('btn-main-copy-result');
            var insertBtn = document.getElementById('btn-main-insert-excel');
            var bottomInsertBtn = document.getElementById('btn-insert-excel');
            if (copyBtn) copyBtn.style.display = '';
            if (insertBtn) insertBtn.style.display = '';
            if (bottomInsertBtn) bottomInsertBtn.style.display = 'none';

        } catch (e) {
            _mainLastResult = null;
            if (resultContainer) {
                resultContainer.innerHTML = '<div style="text-align:center;padding:20px;color:#d32f2f;">&#9888; 解析失败: ' + escapeHtml(e.message) + '</div>';
            }
        }
    }

    /**
     * 显示主页解析结果（复用displayResult但输出到主页容器）
     */
    function displayMainResult(result, format, elapsedMs) {
        const container = document.getElementById('main-parse-result');
        if (!container) return;

        if (!result || typeof result !== 'object' || Object.keys(result).length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:30px;color:#999;"><div style="font-size:24px;margin-bottom:8px;">&#128230;</div><div>' + pt('ocr.parser.noData') + '</div></div>';
            return;
        }

        const script = _scripts.find(function(s) { return s.id === _mainScriptId; });
        const projected = projectParseResult(result, script);
        const fieldNameMap = projected.names;
        const displayData = projected.data;

        if (!displayData || Object.keys(displayData).length === 0) {
            container.innerHTML = '<div style="text-align:center;padding:30px;color:#999;"><div style="font-size:24px;margin-bottom:8px;">&#128230;</div><div>' + pt('ocr.parser.noData') + '</div></div>';
            return;
        }

        let html = '';
        switch (format) {
            case 'table': html = renderTableResult(displayData, fieldNameMap); break;
            case 'key-value': html = renderKeyValueResult(displayData, fieldNameMap); break;
            case 'json': html = renderJsonResult(displayData); break;
            default: html = renderTableResult(displayData, fieldNameMap);
        }

        container.innerHTML = html +
            '<div style="margin-top:10px;text-align:right;font-size:11px;color:#999;">解析耗时: ' + elapsedMs + 'ms | 字段数: ' + Object.keys(displayData).length + '</div>';
    }

    /**
     * 复制主页解析结果
     */
    function copyMainResult() {
        if (!_mainLastResult) { showToast('没有可复制的结果'); return; }

        const script = _scripts.find(function(s) { return s.id === _mainScriptId; });
        const format = script ? script.outputFormat : 'table';
        const projected = projectParseResult(_mainLastResult, script);
        const displayData = projected.data;
        const fieldNameMap = projected.names;
        let text = '';

        if (format === 'json') {
            text = JSON.stringify(displayData, null, 2);
        } else {
            Object.entries(displayData).forEach(function(entry) {
                var key = entry[0], val = entry[1];
                var name = fieldNameMap[key] || key;
                text += name + '\t' + formatValueForText(val) + '\n';
            });
        }

        navigator.clipboard.writeText(text).then(function() {
            showToast('已复制到剪贴板');
        }).catch(function() {
            showToast('复制失败，请手动选择复制');
        });
    }

    /**
     * 将主页解析结果插入Excel
     */
    async function insertMainResultToExcel() {
        if (!_mainLastResult) { showToast('没有可插入的结果'); return; }

        try {
            if (!window.chrome?.webview?.hostObjects?.excelBridge) throw new Error('WebView2 未就绪');

            const script = _scripts.find(function(s) { return s.id === _mainScriptId; });
            const projected = projectParseResult(_mainLastResult, script);
            const configuredTable = tableRowsFromProjected(projected.data, script && script.fields);
            if (configuredTable) {
                const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(configuredTable), true);
                const parsed = JSON.parse(resultObj);
                if (!parsed.success) showToast('插入失败: ' + (parsed.error || '未知错误'));
                return;
            }

            // 优先处理结构化表格数据 { header, rows }
            if (_mainLastResult.tableData && Array.isArray(_mainLastResult.tableData.header) && Array.isArray(_mainLastResult.tableData.rows)) {
                const rows = [_mainLastResult.tableData.header].concat(_mainLastResult.tableData.rows);
                if (rows.length === 0 || rows[0].length === 0) { showToast('没有可插入的数据'); return; }

                const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertTableData(JSON.stringify(rows), true);
                const parsed = JSON.parse(resultObj);
                if (parsed.success) {
                    // 弹窗提示已移除，仅保留状态栏反馈
                } else {
                    showToast('插入失败: ' + (parsed.error || '未知错误'));
                }
                return;
            }

            const displayData = projected.data;
            const fieldNameMap = projected.names;

            // 与批量导出保持一致：横向表格，字段名为列，解析结果为数据行
            const row = {};
            Object.entries(displayData).forEach(function(entry) {
                var key = entry[0], val = entry[1];
                var name = fieldNameMap[key] || key;
                row[name] = formatValueForText(val);
            });

            if (Object.keys(row).length === 0) { showToast('没有可插入的数据'); return; }

            const resultObj = await window.chrome.webview.hostObjects.excelBridge.InsertParsedData(JSON.stringify([row]));
            const parsed = JSON.parse(resultObj);

            if (parsed.success) {
                // 弹窗提示已移除，仅保留状态栏反馈
            } else {
                showToast('插入失败: ' + (parsed.error || '未知错误'));
            }
        } catch (e) {
            showToast('插入Excel失败: ' + e.message);
        }
    }

    /**
     * 更新主页解析区域显示状态（有OCR结果时显示）
     */
    function updateMainParseSectionVisibility() {
        const section = document.getElementById('main-parse-section');
        const ocrText = document.getElementById('ocr-result-text');
        if (!section || !ocrText) return;

        var stored = '';
        if (typeof OcrState !== 'undefined' && OcrState) {
            stored = String(OcrState.batchOriginalText || OcrState.recognitionResult || '').trim();
        }
        var hasText = ocrText.value.trim().length > 0 || stored.length > 0;
        
        if (!hasText) {
            section.style.display = 'none';
            
            const resultContainer = document.getElementById('main-parse-result');
            if (resultContainer) {
                resultContainer.innerHTML = '';
            }
            
            _mainLastResult = null;
            
            var copyBtn = document.getElementById('btn-main-copy-result');
            var insertBtn = document.getElementById('btn-main-insert-excel');
            if (copyBtn) copyBtn.style.display = 'none';
            if (insertBtn) insertBtn.style.display = 'none';
        } else {
            section.style.display = 'flex';
            
            if (_scripts && _scripts.length > 0 && !document.getElementById('main-parse-script-select').options.length > 1) {
                populateMainScriptSelector();
            }

            // 有 OCR 文本时，根据是否已执行过解析切换插入按钮
            var mainCopyBtn = document.getElementById('btn-main-copy-result');
            var mainInsertBtn = document.getElementById('btn-main-insert-excel');
            var bottomInsertBtn = document.getElementById('btn-insert-excel');
            if (_mainLastResult && Object.keys(_mainLastResult).length > 0) {
                if (mainCopyBtn) mainCopyBtn.style.display = '';
                if (mainInsertBtn) mainInsertBtn.style.display = '';
                if (bottomInsertBtn) bottomInsertBtn.style.display = 'none';
            } else {
                if (mainCopyBtn) mainCopyBtn.style.display = 'none';
                if (mainInsertBtn) mainInsertBtn.style.display = 'none';
                if (bottomInsertBtn) bottomInsertBtn.style.display = '';
            }
        }
    }

    /**
     * 显示Toast提示
     */
    function showToast(message) {
        if (window.App && typeof App.showToast === 'function') {
            App.showToast(message);
        } else {
            alert(message);
        }
    }

    // ==================== 公开API ====================
    
    window.App = window.App || {};
    window.App.Parser = {
        // 初始化
        init: init,
        
        // 脚本管理
        loadScripts: loadScripts,
        selectScript: selectScript,
        createNewScript: createNewScript,
        importScriptFromFile: importScriptFromFile,
        saveCurrent: saveCurrent,
        duplicateCurrent: duplicateCurrent,
        deleteCurrent: deleteCurrent,
        deleteCurrentById: deleteCurrentById,
        resetCurrent: resetCurrent,
        clearCache: clearCache,
        
        // 视图
        switchView: switchView,
        
        // 字段配置
        addField: addField,
        updateField: updateField,
        removeField: removeField,
        
        // 解析执行
        runParse: runParse,
        parseOcrResult: parseOcrResult,
        testScript: testScript,
        
        // 测试辅助
        loadSampleText: loadSampleText,
        
        // 结果操作
        copyResult: copyResult,
        insertToExcel: insertToExcel,

        // 搜索筛选
        onSearch: onSearch,
        filterByCategory: filterByCategory,

        // 主页集成
        runParseFromMain: runParseFromMain,
        onMainScriptChange: onMainScriptChange,
        populateMainScriptSelector: populateMainScriptSelector,
        copyMainResult: copyMainResult,
        insertMainResultToExcel: insertMainResultToExcel,
        updateMainParseSectionVisibility: updateMainParseSectionVisibility,

        // 批量处理集成
        onBatchScriptChange: onBatchScriptChange,
        populateBatchScriptSelector: populateBatchScriptSelector,

        // 内部状态访问（供调试）
        get currentScript() { return _currentScript; },
        get lastResult() { return _lastResult; },
        get mainLastResult() { return _mainLastResult; },
        get mainScriptId() { return _mainScriptId; },
        get scripts() { return _scripts; },
        applyLocale: applyLocale
    };

    /**
     * 批量处理脚本选择变更
     */
    function onBatchScriptChange(scriptId) {
        ChatLog.info('OCR', '[Parser-批量] 选择脚本:', scriptId);
        // 可以在此处添加批量处理特定的逻辑
    }

    /**
     * 填充批量处理的脚本选择器
     */
    function populateBatchScriptSelector() {
        const select = document.getElementById('batch-parser-script');
        if (!select || _scripts.length === 0) return;

        // 清空现有选项（保留第一个）
        select.innerHTML = '<option value="">' + pt('ocr.pickScript') + '</option>';

        // 添加脚本选项
        _scripts.forEach(function(script) {
            const option = document.createElement('option');
            option.value = script.id;
            option.textContent = script.name;
            select.appendChild(option);
        });
    }

    // DOM加载完成后自动初始化
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        // 如果DOM已经加载完毕，延迟初始化以确保WebView2就绪
        setTimeout(init, 500);
    }

})();
