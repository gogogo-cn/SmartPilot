/**
 * ============================================================
 * llm-stream-receiver.js — 后端 LLM 流式响应接收器
 * ============================================================
 * 职责：
 *   - 注册 window.__onLLMStream 回调，接收 VB 后端推送的流式 chunk
 *   - 将 chunk 转发到 App.ChatV2 事件系统（供 chat-renderer 渲染）
 *   - 管理 stream 生命周期（start / delta / end / error）
 *
 * 使用方式：
 *   后端通过 ExecuteScriptAsync 调用 window.__onLLMStream(payload)
 *   前端通过 App.ChatV2.streamReceiver.onChunk / onEnd / onError 监听
 *
 * 依赖：
 *   - app-bootstrap.js（App 命名空间）
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    var ChatLog = window.ChatLog || console;

    var _listeners = {};
    var _activeStreams = {};

    function ensureStream(streamId) {
        if (!_activeStreams[streamId]) {
            _activeStreams[streamId] = {
                id: streamId,
                fullContent: '',
                fullReasoning: '',
                chunkCount: 0,
                model: '',
                startedAt: Date.now()
            };
        }
        return _activeStreams[streamId];
    }

    function emit(eventType, data) {
        var handlers = _listeners[eventType] || [];
        for (var i = 0; i < handlers.length; i++) {
            try {
                handlers[i](data);
            } catch (e) {
                ChatLog.error?.('STREAM-RECV', 'Listener error for ' + eventType, e);
            }
        }
    }

    function handleStreamEvent(payload) {
        if (!payload || payload.type !== 'llm_stream') return;

        var evt = payload.event;
        var data = payload.data || {};
        var streamId = data.streamId;

        if (!streamId) {
            ChatLog.warn?.('STREAM-RECV', 'Received stream event without streamId');
            return;
        }

        switch (evt) {
            case 'stream_start':
                var stream = ensureStream(streamId);
                ChatLog.info?.('STREAM-RECV', '🔗 Stream started: ' + streamId);
                emit('start', { streamId: streamId, stream: stream });
                break;

            case 'stream_delta':
                var stream = ensureStream(streamId);
                stream.chunkCount++;
                if (data.content) stream.fullContent += data.content;
                if (data.reasoning) stream.fullReasoning += data.reasoning;
                if (data.model) stream.model = data.model;
                emit('delta', {
                    streamId: streamId,
                    content: data.content || '',
                    reasoning: data.reasoning || '',
                    model: data.model || '',
                    chunkIndex: stream.chunkCount,
                    stream: stream
                });
                break;

            case 'stream_end':
                var stream = _activeStreams[streamId] || ensureStream(streamId);
                stream.isCompleted = true;
                if (data.fullContent) stream.fullContent = data.fullContent;
                var elapsed = ((Date.now() - stream.startedAt) / 1000).toFixed(2);
                ChatLog.info?.('STREAM-RECV', '✅ Stream ended: ' + streamId +
                    ', chunks=' + stream.chunkCount +
                    ', chars=' + stream.fullContent.length +
                    ', time=' + elapsed + 's');
                emit('end', {
                    streamId: streamId,
                    finishReason: data.finishReason || 'stop',
                    fullContent: data.fullContent || stream.fullContent,
                    usage: data.usage || null,
                    stream: stream
                });
                delete _activeStreams[streamId];
                break;

            case 'stream_error':
                var stream = _activeStreams[streamId];
                ChatLog.error?.('STREAM-RECV', '❌ Stream error: ' + streamId + ' - ' + data.error);
                emit('error', {
                    streamId: streamId,
                    error: data.error || 'Unknown error',
                    stream: stream
                });
                delete _activeStreams[streamId];
                break;

            default:
                ChatLog.warn?.('STREAM-RECV', 'Unknown stream event: ' + evt);
        }
    }

    window.__onLLMStream = function(jsonStr) {
        try {
            var payload = (typeof jsonStr === 'string') ? JSON.parse(jsonStr) : jsonStr;
            handleStreamEvent(payload);
        } catch (e) {
            ChatLog.error?.('STREAM-RECV', 'Failed to parse stream event', e);
        }
    };

    var streamReceiver = {
        on: function(eventType, handler) {
            if (!_listeners[eventType]) _listeners[eventType] = [];
            _listeners[eventType].push(handler);
            return this;
        },

        off: function(eventType, handler) {
            if (!_listeners[eventType]) return this;
            if (!handler) {
                delete _listeners[eventType];
            } else {
                _listeners[eventType] = _listeners[eventType].filter(function(h) { return h !== handler; });
            }
            return this;
        },

        getActiveStream: function(streamId) {
            return _activeStreams[streamId] || null;
        },

        getActiveStreamIds: function() {
            return Object.keys(_activeStreams);
        }
    };

    App.ChatV2.streamReceiver = streamReceiver;
    ChatLog.info?.('CHAT-V2', '✅ streamReceiver loaded');
})();