﻿﻿/**
 * ============================================================
 * skill-dom-fixer.js — Skill DOM Fixer
 * ============================================================
 * 功能：Skill DOM 修复器，处理 Skill 执行后产生的 DOM 异常、样式冲突与事件绑定问题。
 *
 * 核心类/函数：
 *   - fixSkillDom() — 修复 Skill DOM
 *   - cleanupOrphanElements() — 清理孤立元素
 *   - rebindEvents() — 重新绑定事件
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function(){
    'use strict';

    /**
     * 规范化技能卡片的 DOM 结构
     * 将不规则的技能卡片转换为统一的结构：header + body
     * @param {HTMLElement} card - 技能卡片元素
     */
    function normalizeSkillCard(card) {
        try {
            // 跳过已修复的卡片或非技能卡片元素
            if (!card || card.__skillFixed) return;
            if (!card.classList || !card.classList.contains('skill-item')) return;

            // 跳过已经有 header 结构的卡片
            if (card.querySelector('.skill-item-header') || card.querySelector('.snippet-item-header')) {
                card.__skillFixed = true;
                return;
            }

            // 查找标题元素（排除全局 header 中的元素）
            var titleEl = card.querySelector('.snippet-item-name, .skill-item-name, .snippet-title, .card-title, h3, strong, b, label, .item-name');
            if (titleEl && titleEl.closest('.skill-list-header')) titleEl = null;

            // 如果找不到标题元素，尝试从文本节点提取
            if (!titleEl) {
                var textNodes = Array.from(card.childNodes).filter(function(n){
                    return n.nodeType === Node.TEXT_NODE && n.textContent.trim().length>0 && !(n.parentElement && n.parentElement.classList && n.parentElement.classList.contains('skill-list-header'));
                });
                if (textNodes.length>0) {
                    titleEl = document.createElement('div');
                    titleEl.textContent = textNodes[0].textContent.trim();
                    textNodes[0].textContent = '';
                }
            }

            // 收集操作按钮（排除全局 header 中的按钮）
            var actionCandidates = Array.from(card.querySelectorAll('button, a[role="button"], .btn, .action-btn'));
            actionCandidates = actionCandidates.filter(function(el){
                if (!card.contains(el)) return false;
                if (el.closest('.skill-list-header')) return false;
                return true;
            });

            // 如果既没有标题也没有操作按钮，跳过处理
            if (!titleEl && actionCandidates.length===0) {
                card.__skillFixed = true;
                return;
            }

            // 创建 header 包装器
            var header = document.createElement('div');
            header.className = 'skill-item-header';
            header.style.display = 'flex';
            header.style.alignItems = 'center';
            header.style.justifyContent = 'space-between';
            header.style.gap = '8px';

            // 创建标题包装器
            var nameWrap = document.createElement('div');
            nameWrap.className = 'skill-item-name';
            nameWrap.style.flex = '1 1 auto';
            nameWrap.style.minWidth = '0';
            nameWrap.style.overflow = 'hidden';
            nameWrap.style.textOverflow = 'ellipsis';
            nameWrap.style.whiteSpace = 'nowrap';

            if (titleEl) {
                // 将原标题元素从卡片中移除并放入 header
                try {
                    if (titleEl.parentNode) titleEl.parentNode.removeChild(titleEl);
                } catch (e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('SKILL', 'DOM修复失败:', e && e.message); }
                try {
                    nameWrap.appendChild(titleEl);
                } catch (e) {
                    // 降级方案：使用文本内容
                    var txt = document.createElement('span');
                    txt.textContent = titleEl.textContent.trim();
                    nameWrap.appendChild(txt);
                }

                // 将分类/标签元素移动到标题后面
                try {
                    var catSelectors = ['.skill-badge', '.badge', '.tag', '.label', '.skill-tag', '.chip', '.category', '.meta-badge'];
                    var categoryCandidates = [];
                    catSelectors.forEach(function(sel){
                        Array.prototype.push.apply(categoryCandidates, Array.from(card.querySelectorAll(sel)));
                    });
                    categoryCandidates = categoryCandidates.filter(function(el){ return card.contains(el) && !el.closest('.skill-list-header'); });
                    categoryCandidates.forEach(function(cat){
                        try { nameWrap.appendChild(cat); } catch(e) { try { nameWrap.appendChild(cat.cloneNode(true)); } catch(e2){ if (typeof ChatLog !== 'undefined') ChatLog.warn('SKILL', 'DOM修复降级也失败:', e2 && e2.message); } }
                    });
                } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('SKILL', '[skill-dom-fixer] 分类标签重排失败:', e && e.message); }

            } else {
                // 降级方案：使用第一个子元素作为标题
                var firstChild = Array.from(card.children).find(function(ch){
                    return !ch.classList.contains('skill-item-header') && !ch.classList.contains('skill-actions') && !ch.classList.contains('card-actions');
                });
                if (firstChild) {
                    try { nameWrap.appendChild(firstChild); } catch(e) { nameWrap.appendChild(firstChild.cloneNode(true)); }
                }
            }

            // 创建操作按钮包装器
            var actionsWrap = document.createElement('div');
            actionsWrap.className = 'skill-actions';
            actionsWrap.style.display = 'flex';
            actionsWrap.style.gap = '8px';
            actionsWrap.style.flex = '0 0 auto';

            // 将按钮移动到操作区域
            actionCandidates.forEach(function(btn){
                try {
                    if (btn.closest('.skill-actions')) return;
                    actionsWrap.appendChild(btn);
                } catch (e) {
                    // 忽略错误
                }
            });

            // 组装 header
            header.appendChild(nameWrap);
            header.appendChild(actionsWrap);

            // 保存现有的子元素，准备移动到 body 中
            var existingChildren = Array.from(card.children).filter(function(ch){ return ch !== header; });

            // 在卡片顶部插入 header
            card.insertBefore(header, card.firstChild);

            // 创建 body 容器并移动剩余内容
            var body = document.createElement('div');
            body.className = 'skill-item-body';
            body.style.marginTop = '8px';

            existingChildren.forEach(function(ch){
                try { body.appendChild(ch); } catch(e) { try { body.appendChild(ch.cloneNode(true)); } catch(e2){ if (typeof ChatLog !== 'undefined') ChatLog.warn('SKILL', 'DOM修复降级也失败:', e2 && e2.message); } }
            });

            // 在 header 后面追加 body
            card.appendChild(body);

            // 标记卡片已修复
            card.__skillFixed = true;
            card.classList.add('skill-card-fixed');
        } catch (e) {
            ChatLog.warn('DOMFIX', 'normalizeSkillCard error', e);
        }
    }

    /**
     * 处理所有技能卡片
     * 遍历 skill-list-container 中的所有 skill-item 并修复
     */
    function processAll() {
        var container = document.getElementById('skill-list-container');
        if (!container) return;
        var cards = Array.from(container.querySelectorAll('.skill-item'));
        cards.forEach(normalizeSkillCard);
    }

    /**
     * 监听 DOM 变化，实时修复新添加的技能卡片
     * 使用 MutationObserver 监听容器变化
     */
    function observe() {
        var container = document.getElementById('skill-list-container');
        if (!container) return;
        var mo = new MutationObserver(function(mutations){
            mutations.forEach(function(m){
                if (m.addedNodes && m.addedNodes.length>0) {
                    Array.from(m.addedNodes).forEach(function(n){
                        if (n.nodeType===1) {
                            // 如果添加的是 skill-item，直接处理
                            if (n.classList && n.classList.contains('skill-item')) {
                                normalizeSkillCard(n);
                            }
                            // 如果添加的是分组容器，处理里面的 skill-item
                            var items = n.querySelectorAll ? n.querySelectorAll('.skill-item') : [];
                            Array.from(items).forEach(normalizeSkillCard);
                        }
                    });
                }
            });
        });
        mo.observe(container, { childList: true, subtree: true });
    }

    /**
     * 在 DOM 就绪后初始化
     * 延迟 150ms 确保其他脚本已完成渲染
     */
    function initWhenReady(){
        if (document.readyState === 'complete' || document.readyState === 'interactive'){
            setTimeout(function(){ processAll(); observe(); }, 150);
        } else {
            document.addEventListener('DOMContentLoaded', function(){ setTimeout(function(){ processAll(); observe(); }, 150); });
        }
    }

    /** 自动初始化 */
    initWhenReady();
})();