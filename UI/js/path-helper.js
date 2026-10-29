/**
 * ============================================================
 * path-helper.js — Path Helper
 * ============================================================
 * 功能：路径辅助工具，提供跨平台路径拼接、扩展名提取、相对/绝对路径转换。
 *
 * 核心类/函数：
 *   - join() — 拼接路径
 *   - getExtension() — 获取扩展名
 *   - toAbsolute() — 转为绝对路径
 *   - normalize() — 规范化路径
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';
    
    // 获取用户APPDATA路径
    function getAppDataPath() {
        // 在浏览器中无法直接获取系统路径
        // 需要通过Node.js后端或VSTO插件传递
        // 这里提供默认配置
        return null;
    }
    
    // 共享目录配置
    const SharedPaths = {
        // 基础目录（需要在初始化时设置）
        baseDir: null,
        
        // Skill文件目录
        get skillsDir() {
            if (!this.baseDir) return null;
            return this.baseDir + '\\Skills';
        },
        
        // 配置文件目录
        get configDir() {
            if (!this.baseDir) return null;
            return this.baseDir + '\\Config';
        },
        
        // 知识库目录
        get knowledgeDir() {
            if (!this.baseDir) return null;
            return this.baseDir + '\\KnowledgeBase';
        },
        
        // 初始化共享路径
        init: function(appDataPath) {
            this.baseDir = appDataPath + '\\SmartPilot';
            ChatLog.info('PATH', '[PathHelper] 共享路径已初始化:', this.baseDir);
        },
        
        // 获取完整文件路径
        getSkillFilePath: function(fileName) {
            if (!this.skillsDir) return null;
            return this.skillsDir + '\\' + fileName;
        },
        
        getConfigFilePath: function(fileName) {
            if (!this.configDir) return null;
            return this.configDir + '\\' + fileName;
        }
    };
    
    // 暴露到全局
    window.SharedPaths = SharedPaths;
    
    ChatLog.info('PATH', '[PathHelper] 模块已加载');
})();
