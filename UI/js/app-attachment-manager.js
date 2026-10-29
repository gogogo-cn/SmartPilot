/**
 * ============================================================
 * app-attachment-manager.js — Attachment Manager
 * ============================================================
 * 功能：附件管理模块，负责聊天附件的上传、预览、删除与持久化。
 *
 * 核心类/函数：
 *   - addAttachment() — 添加附件
 *   - removeAttachment() — 移除附件
 *   - getAttachments() — 获取附件列表
 *   - clearAttachments() — 清空附件
 *
 * 依赖：
 *   - app.js（App 全局命名空间）
 * ============================================================
 */

(function() {
    'use strict';

    const ChatLog = window.ChatLog || console;

    /**
     * 附件管理器类
     * 提供文件附件的统一管理能力
     */
    class AttachmentManager {
        /**
         * 构造函数
         */
        constructor() {
            this.messageAttachmentsMap = new Map();
            this.instantImageUrls = [];
            this._pendingBase64Conversions = new Map();
        }

        /**
         * 判断文件是否为图片文件
         * @param {File} file - 文件对象
         * @returns {boolean} 是否为图片
         */
        isImageFile(file) {
            if (!file) return false;
            const imageExts = ['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp', 'svg', 'ico', 'tiff', 'tif'];
            const ext = (file.name || '').toLowerCase().split('.').pop();
            return imageExts.includes(ext) || (file.type || '').startsWith('image/');
        }

        /**
         * 渲染图片缩略图HTML
         * @param {File} file - 文件对象
         * @param {string} src - 图片URL（可以是DataURL或ObjectURL）
         * @returns {string} 图片缩略图HTML
         */
        renderImageThumbnail(file, src) {
            const escapeHtml = typeof App !== 'undefined' && App.escapeHtml 
                ? App.escapeHtml 
                : (str => str ? String(str).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])) : '');
            
            return `
                <div class="image-container vertical-top" style="margin-top: 8px;">
                    <img src="${src}" alt="${escapeHtml(file.name)}" style="max-width:260px;max-height:200px;border-radius:8px;object-fit:contain;">
                    <div class="image-caption">
                        <div>🖼️ ${escapeHtml(file.name)}</div>
                        <div>${(file.size / 1024).toFixed(1)} KB</div>
                    </div>
                </div>
            `;
        }

        /**
         * 异步渲染图片并转换为Base64
         * 先用ObjectURL快速显示，后台异步转换为Base64
         * @param {File} file - 文件对象
         * @returns {Promise<string>} 图片缩略图HTML
         */
        async renderImageWithBase64(file) {
            const instantUrl = URL.createObjectURL(file);
            
            const reader = new FileReader();
            const base64Promise = new Promise(resolve => {
                reader.onload = (e) => {
                    resolve(e.target.result);
                };
                reader.readAsDataURL(file);
            });

            base64Promise.then(base64 => {
                const imgEl = document.querySelector('img[src="' + instantUrl + '"]');
                if (imgEl) {
                    imgEl.src = base64;
                    imgEl.setAttribute('data-base64', base64);
                }
            }).catch(e => {
                ChatLog.error('AttachmentManager', 'base64转换失败:', e.message);
            });

            this.instantImageUrls.push(instantUrl);
            setTimeout(() => {
                try { URL.revokeObjectURL(instantUrl); } catch(e) { ChatLog.warn?.('UI', '[AttachmentManager] 撤销 objectURL 失败:', e); }
                const idx = this.instantImageUrls.indexOf(instantUrl);
                if (idx >= 0) this.instantImageUrls.splice(idx, 1);
            }, 60000);

            return this.renderImageThumbnail(file, instantUrl);
        }

        /**
         * 保存附件到消息关联映射
         * @param {string} messageId - 消息ID
         * @param {File[]} files - 文件数组
         */
        saveAttachments(messageId, files) {
            if (!messageId || !files || files.length === 0) return;
            this.messageAttachmentsMap.set(messageId, files);
            
            const lastUserWrapper = document.querySelector('#chat-messages .message-wrapper.user:last-child');
            if (lastUserWrapper) {
                const attachmentsData = files.map(f => ({
                    name: f.name,
                    size: f.size,
                    type: f.type
                }));
                lastUserWrapper.setAttribute('data-attachments', JSON.stringify(attachmentsData));
            }
        }

        /**
         * 获取指定消息的附件
         * @param {string} messageId - 消息ID
         * @returns {File[]|undefined} 文件数组
         */
        getAttachments(messageId) {
            return this.messageAttachmentsMap.get(messageId);
        }

        /**
         * 从文件输入框恢复文件列表
         * @returns {File[]|null} 文件数组
         */
        recoverFiles() {
            const fileInputEl = document.getElementById('file-input');
            if (fileInputEl && fileInputEl.files && fileInputEl.files.length > 0) {
                const files = [];
                for (let i = 0; i < fileInputEl.files.length; i++) {
                    files.push(fileInputEl.files[i]);
                }
                return files;
            }
            return null;
        }

        /**
         * 清理所有附件数据和资源
         */
        clear() {
            this.messageAttachmentsMap.clear();
            this.instantImageUrls.forEach(url => {
                try { URL.revokeObjectURL(url); } catch(e) { ChatLog.warn?.('UI', '[AttachmentManager] 清理 objectURL 失败:', e); }
            });
            this.instantImageUrls = [];
            this._pendingBase64Conversions.clear();
        }
    }

    if (typeof App !== 'undefined') {
        App.AttachmentManager = AttachmentManager;
        App.attachmentManager = new AttachmentManager();
    } else if (typeof window !== 'undefined') {
        if (!window.App) window.App = {};
        App.AttachmentManager = AttachmentManager;
        App.attachmentManager = new AttachmentManager();
    }

    window.AttachmentManager = AttachmentManager;
})();