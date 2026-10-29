/**
 * ============================================================
 * skill-shared-loader.js — Skill Shared Loader
 * ============================================================
 * 功能：Skill 共享加载器，负责加载共享 Skill 库、依赖解析与版本管理。
 *
 * 核心类/函数：
 *   - loadSharedSkill() — 加载共享 Skill
 *   - resolveDependencies() — 解析依赖
 *   - getSharedSkillList() — 获取共享列表
 *
 * 依赖：
 *   - app-ai-skill.js
 *   - app-utils.js
 * ============================================================
 */

(function() {
    'use strict';
    
    ChatLog.info('SKILL', '[SharedSkillLoader] ========== 模块开始加载 ==========');
    ChatLog.info('SKILL', '[SharedSkillLoader] window.SharedSkillLoader:', window.SharedSkillLoader);
    
    // 检测是否在 VSTO WebView2 环境（COM 桥接可用）
    function isVstoWebView2() {
        return !!(window.chrome && window.chrome.webview && window.chrome.webview.hostObjects && window.chrome.webview.hostObjects.excelBridge && !window.chrome.webview._mockMode);
    }

    // 从后端API加载共享Skill
    async function loadSharedSkills() {
        ChatLog.info('SKILL', '[SharedSkillLoader.load] 开始执行 loadSharedSkills...');

        // VSTO WebView2 环境下，共享 Skill 已通过 COM 桥接加载到 SkillManager，无需 HTTP
        if (isVstoWebView2()) {
            ChatLog.info('SKILL', '[SharedSkillLoader.load] VSTO WebView2 环境 detected，共享 Skill 已由 COM 桥接加载，跳过 HTTP 请求');
            window.__sharedSkills = [];
            window.__sharedSkillsData = [];
            return [];
        }

        try {
            ChatLog.info('SKILL', '[SharedSkillLoader.load] 开始从后端加载共享Skill...');
            const skillsUrl = App.Endpoints ? App.Endpoints.excelBridge.skillsList : (window.ExcelBridgeConfig ? window.ExcelBridgeConfig.skillsUrl : 'http://localhost:9527/api/skills/shared');
            ChatLog.info('SKILL', '[SharedSkillLoader.load] API URL:', skillsUrl);
            
            const response = await fetch(skillsUrl);
            const data = await response.json();
            
            if (data.success && data.skills && data.skills.length > 0) {
                ChatLog.info('SKILL', '[SharedSkillLoader] ✅ 已加载', data.count, '个共享Skill');
                
                // 存储到全局变量
                window.__sharedSkills = data.skills;
                window.__sharedSkillsData = data.skills;  // 供loadCustomSkills使用
                
                // 【修复】构建分类数据，用于分类导航渲染
                window._skillCategoryData = {};
                data.skills.forEach(skill => {
                    const cat = skill.category || '其他';
                    if (!window._skillCategoryData[cat]) window._skillCategoryData[cat] = [];
                    window._skillCategoryData[cat].push(skill);
                });
                ChatLog.info('SKILL', '[SharedSkillLoader] 分类数据:', Object.keys(window._skillCategoryData));
                
                // 合并到App.skills（如果App已加载）
                if (window.App && window.App.loadSkillsForLLM) {
                    ChatLog.info('SKILL', '[SharedSkillLoader] 将共享Skill合并到App.skills...');
                    mergeSkillsToApp(data.skills);
                }
                
                // 【事件驱动】触发全局事件，通知UI模块数据已就绪
                // UI模块（app-chat-ui.js）会监听此事件进行渲染
                ChatLog.info('SKILL', '[SharedSkillLoader] 触发 skills-data-ready 事件...');
                window.dispatchEvent(new CustomEvent('skills-data-ready', {
                    detail: {
                        skills: data.skills,
                        categories: window._skillCategoryData
                    }
                }));
                
                return data.skills;
            } else {
                ChatLog.info('SKILL', '[SharedSkillLoader] 未找到共享Skill');
                window.__sharedSkills = [];
                return [];
            }
        } catch (error) {
            ChatLog.error('SKILL', '[SharedSkillLoader] 加载失败:', error);
            window.__sharedSkills = [];
            return [];
        }
    }
    
    // 合并Skill到App
    function mergeSkillsToApp(sharedSkills) {
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] ========== 开始合并Skill ==========');
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] sharedSkills数量:', sharedSkills.length);
        
        if (sharedSkills.length > 0) {
            ChatLog.info('SKILL', '[SharedSkillLoader.merge] 第一个Skill结构:');
            ChatLog.info('SKILL', '[SharedSkillLoader.merge] - name:', sharedSkills[0].name);
            ChatLog.info('SKILL', '[SharedSkillLoader.merge] - type:', sharedSkills[0].type);
            ChatLog.info('SKILL', '[SharedSkillLoader.merge] - content长度:', sharedSkills[0].content ? sharedSkills[0].content.length : 0);
            ChatLog.info('SKILL', '[SharedSkillLoader.merge] - content前100字符:', sharedSkills[0].content ? sharedSkills[0].content.substring(0, 100) : 'null');
        }
        
        if (!window.App || !window.App.cachedSkills) {
            ChatLog.warn('SKILL', '[SharedSkillLoader.merge] ❌ App.cachedSkills 未就绪，跳过合并');
            return;
        }
        
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] App.cachedSkills当前数量:', window.App.cachedSkills.length);
        
        let mergedCount = 0;
        let parseFailCount = 0;
        
        sharedSkills.forEach((skill, index) => {
            ChatLog.info('SKILL', `[SharedSkillLoader.merge] [${index+1}/${sharedSkills.length}] 处理:`, skill.name);
            
            // 检查是否已存在
            const existsIdx = window.App.cachedSkills.findIndex(s => s.name === skill.name);
            if (existsIdx >= 0) {
                // 已存在（通常是后端 GetAllSkills 提供的定义，缺前端路由字段）。
                // 不能整体跳过 —— 否则 pipeline/executionMode/steps 丢失会把技能误路由到后端执行。
                // 改为回填解析出的关键路由字段到已存在条目（已存在者优先，空值才补）。
                const parsedSkill = parseMarkdownSkill(skill.content, skill.name);
                if (parsedSkill) {
                    const prev = window.App.cachedSkills[existsIdx];
                    ['pipeline', 'executionMode', 'processingMode', 'steps', 'subSkills', 'quickPrompts', 'uiActions'].forEach(f => {
                        const v = parsedSkill[f];
                        if (v !== undefined && v !== null && v !== '' && (prev[f] === undefined || prev[f] === null || prev[f] === '')) {
                            prev[f] = v;
                        }
                    });
                    const _kwMd = parsedSkill.keywords;
                    if (Array.isArray(_kwMd) && _kwMd.length > 0) {
                        const _kwPrev = Array.isArray(prev.keywords) ? prev.keywords.length : 0;
                        const _kwPrevCap = Array.isArray(prev.Keywords) ? prev.Keywords.length : 0;
                        if (_kwPrev === 0 && _kwPrevCap === 0) {
                            prev.keywords = _kwMd;
                            ChatLog.info('SKILL', '[SharedSkillLoader.merge] ⇢ 回填 keywords:', skill.name, _kwMd.length, '个');
                        }
                    }
                    prev.source = prev.source || 'shared';
                    prev.isCustom = prev.isCustom !== false;
                    ChatLog.info('SKILL', '[SharedSkillLoader.merge] ⇢ 已存在，已回填路由字段:', skill.name, 'pipeline=', parsedSkill.pipeline || prev.pipeline);
                } else {
                    ChatLog.info('SKILL', '[SharedSkillLoader.merge] 已存在且解析失败，跳过:', skill.name);
                }
                mergedCount++;
                return;
            }
            
            // 解析Markdown格式
            if (!skill.content) {
                ChatLog.warn('SKILL', '[SharedSkillLoader.merge] ❌ content为空:', skill.name);
                parseFailCount++;
                return;
            }
            
            const parsedSkill = parseMarkdownSkill(skill.content, skill.name);
            if (parsedSkill) {
                parsedSkill.source = 'shared';
                parsedSkill.isCustom = true;
                window.App.cachedSkills.push(parsedSkill);
                mergedCount++;
                ChatLog.info('SKILL', '[SharedSkillLoader.merge] ✅ 已合并:', skill.name);
            } else {
                parseFailCount++;
                ChatLog.warn('SKILL', '[SharedSkillLoader.merge] ❌ 解析失败:', skill.name);
            }
        });
        
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] ========== 合并完成 ==========');
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] 成功:', mergedCount, '个');
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] 失败:', parseFailCount, '个');
        ChatLog.info('SKILL', '[SharedSkillLoader.merge] App.cachedSkills现在数量:', window.App.cachedSkills.length);
    }
    
    // 解析Markdown格式的Skill文件
    function parseMarkdownSkill(content, fileName) {
        try {
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] ========== 开始解析:', fileName, '==========');
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] 内容长度:', content.length, '字符');
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] 内容前100字符:', content.substring(0, 100));
            
            // 提取YAML frontmatter（兼容Windows和Unix换行符）
            // 【修复】使用\r?\n匹配所有换行符类型
            const yamlMatch = content.match(/^---\r?\n([\s\S]*?)\r?\n---/);
            if (!yamlMatch) {
                ChatLog.warn('SKILL', '[SharedSkillLoader.parse] ❌ 未找到YAML分隔符:', fileName);
                ChatLog.info('SKILL', '[SharedSkillLoader.parse] 内容开头:', content.substring(0, 50));
                return null;
            }
            
            const yamlContent = yamlMatch[1];
            const mdContent = content.replace(/^---\r?\n[\s\S]*?\r?\n---/, '').trim();
            
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] YAML内容:', yamlContent.substring(0, 300));
            
            // 提取 llmSystemPrompt（从 ## System Prompt 部分）
            // 使用负向前瞻排除 ###、#### 等三级及以上标题
            const systemPromptMatch = mdContent.match(/^## System Prompt\s*\r?\n([\s\S]*?)(?=^##(?!#) |$)/m);
            const llmSystemPrompt = systemPromptMatch ? systemPromptMatch[1].trim() : '';
            
            // 提取 textInputRules（嵌套对象）
            const textInputRules = extractYamlObject(yamlContent, 'textInputRules');
            
            // 改进的YAML解析
            const skill = {
                name: extractYamlValue(yamlContent, 'name') || fileName,
                displayName: extractYamlValue(yamlContent, 'displayName') || fileName,
                description: extractYamlValue(yamlContent, 'description') || '',
                category: extractYamlValue(yamlContent, 'category') || 'Custom',
                content: mdContent,
                llmSystemPrompt: llmSystemPrompt,
                keywords: extractYamlArray(yamlContent, 'keywords'),
                examples: extractYamlArray(yamlContent, 'examples'),
                allowedTools: extractYamlArray(yamlContent, 'allowedTools'),
                isGeneric: extractYamlBool(yamlContent, 'isGeneric'),
                processingMode: extractYamlValue(yamlContent, 'processingMode') || 'LLM',
                // 【路由关键】技能声明的 pipeline：用于 orchestrator 把技能路由到对应前端确定性 pipeline。
                //   pipeline: "html" → html-display-pipeline（读选区→生成真报表→浏览器展示，自动完成）
                //   缺省按 'skill' 处理（后端 VB 执行 / agent）。漏读该字段会导致 html 技能被误路由到
                //    agent 推理而编造文件路径、流程不收尾。
                pipeline: extractYamlValue(yamlContent, 'pipeline') || null,
                // 统一声明式执行模式字段：
                //   executionMode: spec  | agent | fixed
                //     - spec  → 确定性 Spec 引擎（等价 processingMode: SkillSpec）
                //     - agent → 自主 Agent 循环（ReAct 多步 + token 水位 + 步骤日志）
                //     - fixed → 传统模式（按 processingMode 走后端 SkillExecutionRouter）
                //   agentConfig: 仅 agent 模式生效，字段与 multi-step-agent-mcp 对齐
                //     - maxIterations / maxContextTokens / tokenWarnRatio / tokenCriticalRatio / keepLatestRound
                executionMode: extractYamlValue(yamlContent, 'executionMode') || null,
                agentConfig: extractYamlObject(yamlContent, 'agentConfig'),
                uiActions: extractYamlObject(yamlContent, 'uiActions'),
                needsExcelContext: extractYamlBool(yamlContent, 'needsExcelContext'),
                needsTaskPlanning: extractYamlBool(yamlContent, 'needsTaskPlanning'),
                allowMcpTools: extractYamlBool(yamlContent, 'allowMcpTools'),
                needsTextInput: extractYamlBool(yamlContent, 'needsTextInput'),
                textInputRules: textInputRules,
                autoImportParseScript: extractYamlBool(yamlContent, 'autoImportParseScript'),
                subSkills: extractYamlArray(yamlContent, 'subSkills'),
                // 解析 quickPrompts 对象数组（输出卡片后渲染快捷按钮）
                quickPrompts: extractYamlObjectArray(yamlContent, 'quickPrompts'),
                // 解析声明式步骤（steps，SkillCreator 默认产物）。
                // 后端 WriteYaml 以单行 JSON 标量输出 params/code，故这里仅需平坦解析 + JSON.parse。
                steps: extractSteps(yamlContent)
            };
            
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] ✅ 解析成功:', skill.name);
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] displayName:', skill.displayName);
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] category:', skill.category);
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] isGeneric:', skill.isGeneric);
            ChatLog.info('SKILL', '[SharedSkillLoader.parse] keywords数量:', skill.keywords.length);
            
            return skill;
        } catch (error) {
            ChatLog.error('SKILL', '[SharedSkillLoader.parse] ❌ 解析异常:', fileName, error);
            ChatLog.error('SKILL', '[SharedSkillLoader.parse] 错误堆栈:', error.stack);
            return null;
        }
    }
    
    // 提取YAML值
    function extractYamlValue(yaml, key) {
        const regex = new RegExp(`^${key}:\\s*(.+)$`, 'm');
        const match = yaml.match(regex);
        if (!match) return null;
        
        let value = match[1].trim();
        // 如果值后面跟着多行内容（直到下一个key或数组结束）
        const nextKeyRegex = new RegExp(`^\\w+:`, 'm');
        const arrayRegex = new RegExp(`^${key}:\\s*\\n`, 'm');
        
        // 如果是数组格式，返回null让extractYamlArray处理
        if (arrayRegex.test(yaml)) {
            return null;
        }
        
        return value;
    }
    
    // 提取YAML数组
    function extractYamlArray(yaml, key) {
        const regex = new RegExp(`^${key}:\\s*\\n([\\s\\S]*?)(?=^\\w+:|$)`, 'm');
        const match = yaml.match(regex);
        if (!match) return [];
        
        const lines = match[1].split('\n');
        return lines
            .filter(line => line.trim().startsWith('- '))
            .map(line => line.trim().substring(2));
    }
    
    // 提取YAML布尔值
    function extractYamlBool(yaml, key) {
        const value = extractYamlValue(yaml, key);
        if (value === null) return false;
        return value.toLowerCase() === 'true';
    }
    
    // 提取YAML对象（如 uiActions）
    function extractYamlObject(yaml, key) {
        const regex = new RegExp(`^${key}:\\s*\\n((?:[ \\t]+.*\\n?)+)`, 'm');
        const match = yaml.match(regex);
        if (!match) return null;
        
        const content = match[1];
        const obj = {};
        const lines = content.split('\n');
        
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed || !trimmed.includes(':')) continue;
            
            const parts = trimmed.split(':', 2);
            if (parts.length === 2) {
                const objKey = parts[0].trim();
                const objValue = parts[1].trim().toLowerCase();
                
                if (objValue === 'true') {
                    obj[objKey] = true;
                } else if (objValue === 'false') {
                    obj[objKey] = false;
                } else if (!isNaN(objValue) && objValue !== '') {
                    obj[objKey] = parseFloat(objValue);
                } else {
                    obj[objKey] = objValue.replace(/^["']|["']$/g, '');
                }
            }
        }
        
        return Object.keys(obj).length > 0 ? obj : null;
    }
    
    // 提取YAML对象数组（如 quickPrompts: [{label, prompt}, ...]）
    function extractYamlObjectArray(yaml, key) {
        // 原 (?=^\w+:|$) 在 m 模式下 $ 匹配行尾，导致只捕获数组第一行；
        // 改用 (?![\s\S]) 表示"仅字符串末尾"，正确捕获到下一个顶层键或 frontmatter 结束。
        const regex = new RegExp(`^${key}:\\s*\\n([\\s\\S]*?)(?=^\\w+:|(?!\\[\\s\\S\\]))`, 'm');
        const match = yaml.match(regex);
        if (!match) return [];
        
        const lines = match[1].split('\n');
        const items = [];
        let current = null;
        
        for (const line of lines) {
            const trimmed = line.trim();
            if (!trimmed) continue;
            
            // 新对象项：以 "- " 开头（支持 "- label: xxx" 内联形式）
            const itemMatch = trimmed.match(/^-\s+(.+)$/);
            if (itemMatch) {
                if (current && current.label && current.prompt) items.push(current);
                const kv = itemMatch[1].split(':', 2);
                if (kv.length === 2) {
                    current = {};
                    current[kv[0].trim()] = kv[1].trim().replace(/^["']|["']$/g, '');
                } else {
                    current = null;
                }
                continue;
            }
            
            // 属性行："  prompt: xxx"
            const kv = trimmed.split(':', 2);
            if (kv.length === 2 && current) {
                current[kv[0].trim()] = kv[1].trim().replace(/^["']|["']$/g, '');
            }
        }
        if (current && current.label && current.prompt) items.push(current);
        return items;
    }
    
    // 提取声明式步骤（steps 对象数组）。
    // 后端 WriteYaml 格式：params/code 均为"单行 JSON 标量"（EscapeYamlString 双引号包裹），
    // name/tool/strategy/description/waitOn 为简单标量 → 平坦提取后对 params/code 做 JSON.parse。
    function extractSteps(yaml) {
        const regex = new RegExp(`^steps:\\s*\n([\\s\\S]*?)(?=^\w+:|(?!\[\s\S\]))`, 'm');
        const match = yaml.match(regex);

        // 也兼容 frontmatter 末尾无后续顶层键的写法
        if (!match) {
            const m2 = yaml.match(/^steps:\s*\n([\s\S]*)$/m);
            if (m2) { return parseStepsBlock(m2[1]); }
            return [];
        }
        return parseStepsBlock(match[1]);
    }

    function parseStepsBlock(block) {
        if (!block) return [];
        const items = [];
        let current = null;

        for (const line of block.split('\n')) {
            const trimmed = line.trim();
            if (!trimmed) continue;

            // 新步骤项：以 "- " 开头
            if (/^-\s+/.test(trimmed)) {
                if (current) items.push(current);
                current = {};
                const rest = trimmed.replace(/^-\s+/, '');
                const ci = rest.indexOf(':');
                if (ci > 0 && rest.slice(ci + 1).trim() !== '') {
                    current[rest.slice(0, ci).trim()] = unquoteScalar(rest.slice(ci + 1).trim());
                }
                continue;
            }

            // 属性行："  key: value"（params/code 为 JSON 标量，须取第一个冒号后的全部内容）
            const ci = trimmed.indexOf(':');
            if (ci > 0 && current) {
                const k = trimmed.slice(0, ci).trim();
                const v = unquoteScalar(trimmed.slice(ci + 1).trim());
                if (k === 'params' || k === 'code') {
                    current[k] = parseJsonField(v);
                } else if (k === 'waitOn') {
                    current[k] = v === 'true';
                } else {
                    current[k] = v;
                }
            }
        }
        if (current) items.push(current);
        return items;
    }

    // 去首尾引号（与 extractYamlValue 保持一致）
    function unquoteScalar(v) {
        return String(v).replace(/^["']|["']$/g, '');
    }

    // 单行 JSON 标量 → 对象/字符串；解析失败原样返回字符串
    function parseJsonField(raw) {
        if (!raw) return raw;
        try {
            const parsed = JSON.parse(raw);
            return parsed;
        } catch (e) {
            return raw;
        }
    }
    
    // 暴露到全局
    window.SharedSkillLoader = {
        load: loadSharedSkills,
        getSharedSkills: () => window.__sharedSkills || []
    };
    
    ChatLog.info('SKILL', '[SharedSkillLoader] 模块已加载');
})();