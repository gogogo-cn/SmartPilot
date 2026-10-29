/**
 * 列表底部「全选」：跟行勾选框同一侧，点一次只改勾选。
 * 运行：node Tests/ocr-select-all.test.js
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const html = fs.readFileSync(path.join(__dirname, '..', 'UI', 'ocr-panel.html'), 'utf8').replace(/\r\n/g, '\n');

function sliceFn(name) {
    const start = html.indexOf('function ' + name + '(');
    assert.ok(start >= 0, '缺少 ' + name);
    const next = html.indexOf('\n        function ', start + 10);
    return html.slice(start, next > start ? next : start + 1200);
}

test('全选在操作按钮左侧，且不会点一次触发两次', () => {
    const footer = html.slice(html.indexOf('id="result-footer"'), html.indexOf('id="main-parse-section"'));
    const selectAt = footer.indexOf('id="lbl-select-all-annotations"');
    const recognizeAt = footer.indexOf('id="btn-manual-recognize"');
    assert.ok(selectAt >= 0 && selectAt < recognizeAt, '全选应排在「识别选中」左边');
    assert.ok(!/lbl-select-all-annotations[\s\S]{0,240}onclick=/.test(footer), '标签上不应再绑 onclick');
    assert.match(footer, /onchange="toggleSelectAllAnnotations\(this\.checked\)"/);
});

test('连续截图先收集再一次性识别', () => {
    assert.match(html, /window\.beginScreenshotBatch = function/);
    assert.match(html, /window\.appendScreenshot = function/);
    const start = html.indexOf('window.recognizeScreenshotBatch = async function');
    assert.ok(start >= 0, '缺少 recognizeScreenshotBatch');
    const fn = html.slice(start, start + 4500);
    assert.match(fn, /recognizeImage\(dataUrl/);
    assert.match(fn, /outputFormat: 'plaintext'/);
    assert.match(fn, /storeBatchOriginal\(parts, rawParts, excelParts, markdownParts\)/);
});

test('多张截图识别后仍根据原文显示脚本解析', () => {
    const parser = fs.readFileSync(path.join(__dirname, '..', 'UI', 'js', 'app-ocr-parser.js'), 'utf8');
    const vis = parser.slice(parser.indexOf('function updateMainParseSectionVisibility'), parser.indexOf('function showToast'));
    assert.match(vis, /batchOriginalText/);
    assert.match(vis, /recognitionResult/);
    const store = html.slice(html.indexOf('function renderBatchOriginal'), html.indexOf('window.clearScreenshotStrip = function'));
    assert.match(store, /updateMainParseSectionVisibility/);
});

test('多张截图时显示全部识别，并一次识别队列里的图', () => {
    assert.match(html, /id="btn-recognize-all"[^>]*onclick="recognizeScreenshotBatch\(\)"/);
    const syncStart = html.indexOf('function syncRecognizeAllButton()');
    assert.ok(syncStart >= 0, '缺少 syncRecognizeAllButton');
    const sync = html.slice(syncStart, syncStart + 500);
    assert.match(sync, /pendingScreenshots/);
    assert.match(sync, /count > 1/);
});

test('截图载入后显示「识别这张图」，而不是只解开隐藏的开始识别', () => {
    const fnStart = html.indexOf('window.loadScreenshotFromBase64 = function');
    assert.ok(fnStart >= 0, '缺少 loadScreenshotFromBase64');
    const fn = html.slice(fnStart, fnStart + 3500);
    assert.match(fn, /setRecognizeImageReady\(true\)/);
    assert.ok(!fn.includes("getElementById('btn-recognize')"), '截图路径不应只操作已隐藏的底部按钮');
});

test('全选只改勾选，不重绘整份结果', () => {
    const fn = sliceFn('toggleSelectAllAnnotations');
    assert.ok(!fn.includes('updateAnnotationViews'), '全选不应重绘列表和原文');
    assert.match(fn, /syncSelectAllCheckbox\(\)/);
    const sync = sliceFn('syncSelectAllCheckbox');
    assert.match(sync, /indeterminate/);
    assert.match(sync, /ocr-select-all-count/);
});
