/**
 * ============================================================
 * Tests/host-export-path.test.js — PPT/Word 导出（PDF/图片）保存路径兜底
 * ============================================================
 * 覆盖：
 *   1) 模型未传 filePath 时，前端生成器算出真实桌面绝对路径（.pdf）
 *   2) 用户说"下载" → 走 Downloads 目录
 *   3) 用户消息里给了保存目录 → 尊重用户指定的目录
 *   4) 文件名优先取当前文档名（context.activeSheet.workbook），并去扩展名
 *   5) 桥不可用（拿不到目录）→ 返回 null，交回后端兜底/反问，不瞎猜
 *   6) tool-registry.js 里确实注册了 6 个导出工具的生成器键（防误删注册）
 *
 * 背景（2026-10-05）：宿主 DLL 未重编译时旧版 Bridge.ExportPdf 仍有
 * `filePath 不能为空` 硬校验，只靠 BACKEND_INFERRED_PARAMS 留空给后端会直接失败。
 * 运行：node Tests/host-export-path.test.js
 * ============================================================
 */
'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { loadChatV2, BASE } = require('./helpers/harness');

/** Excel 主桥风格：GetDesktopPath 返回裸路径 */
function stubBridge(desktop, downloads) {
    return {
        GetDesktopPath: () => desktop,
        GetDownloadsPath: () => downloads,
    };
}

/**
 * PPT/Word 宿主真实形态（实测 2026-10-05）：
 *   - 没有 override GetDesktopPath → 调基类，返回 {"success":false,"error":"operation not supported on this host"}
 *   - 但能给出当前文档的完整路径（GetPresentationInfo / GetDocumentInfo）
 * 之前用"理想桥"做桩，掩盖了"取不到桌面 → 生成器返回 null → 功能其实没修好"的问题。
 */
function stubHostBridge(fullName, docName) {
    return {
        GetDesktopPath: () => '{"success":false,"error":"operation not supported on this host"}',
        GetPresentationInfo: () => JSON.stringify({ success: true, name: docName || '演示文稿1.pptx', fullName: fullName, slides: 12 }),
    };
}

function gen(App, opts) {
    return App.ChatV2.toolRegistryUtil.inferHostExportFilePath;
}

test('导出 PDF 未指定路径时自动生成桌面绝对路径', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubBridge('C:\\Users\\tester\\Desktop', 'C:\\Users\\tester\\Downloads');
    const out = await gen(App)({ userMessage: '将演示文稿导出为 PDF。', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out, '应生成路径');
    assert.ok(path.win32.isAbsolute(out), '应为绝对路径: ' + out);
    assert.ok(out.toLowerCase().endsWith('.pdf'), '应为 .pdf: ' + out);
    assert.ok(out.startsWith('C:\\Users\\tester\\Desktop'), '应落在桌面: ' + out);
    assert.ok(/_\d{8}_\d{6}\.pdf$/.test(out), '应带时间戳避免覆盖: ' + out);
});

test('用户说"下载"时落到 Downloads 目录', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubBridge('C:\\Users\\tester\\Desktop', 'C:\\Users\\tester\\Downloads');
    const out = await gen(App)({ userMessage: '导出 PDF 到下载', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out.startsWith('C:\\Users\\tester\\Downloads'), '应落在下载目录: ' + out);
});

test('用户消息里给了保存目录时尊重用户指定', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubBridge('C:\\Users\\tester\\Desktop', 'C:\\Users\\tester\\Downloads');
    const out = await gen(App)({ userMessage: '把演示文稿保存到 D:\\季度报告\\最终稿.pdf', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out.startsWith('D:\\季度报告'), '应落在用户指定目录: ' + out);
    assert.ok(!out.includes('最终稿'), '不应直接用用户给的文件名（避免覆盖）: ' + out);
});

test('文件名优先取当前文档名并去掉扩展名', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubBridge('C:\\Users\\tester\\Desktop', 'C:\\Users\\tester\\Downloads');
    const out = await gen(App)(
        { userMessage: '导出为 PDF', context: { activeSheet: { workbook: '2026Q3季报.pptx' } } },
        { ext: 'pdf', base: '演示文稿' }
    );
    assert.ok(path.win32.basename(out).startsWith('2026Q3季报_'), '应使用文档名: ' + out);
    assert.ok(!out.includes('.pptx'), '应去掉原扩展名: ' + out);
});

test('桥不可用时返回 null（交回后端兜底，不瞎猜路径）', async () => {
    const { App } = loadChatV2();
    App.excelBridge = null;
    const out = await gen(App)({ userMessage: '导出 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.strictEqual(out, null);
});

test('PPT 真实桥形态：GetDesktopPath 返回基类 error JSON 时，改由文档路径反推当前用户桌面', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubHostBridge('C:\\Users\\Administrator\\Desktop\\2026Q3季报.pptx', '2026Q3季报.pptx');
    const out = await gen(App)({ userMessage: '将演示文稿导出为 PDF。', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out, '必须生成路径（否则等于没修）');
    assert.ok(out.startsWith('C:\\Users\\Administrator\\Desktop\\'), '应落在当前用户桌面: ' + out);
    assert.ok(out.toLowerCase().endsWith('.pdf'), '应为 .pdf: ' + out);
});

test('文档在 Documents 下时，桌面仍按当前用户 profile 解析（不跟文档目录走）', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubHostBridge('C:\\Users\\Administrator\\Documents\\周会材料.pptx', '周会材料.pptx');
    const out = await gen(App)({ userMessage: '导出为 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out.startsWith('C:\\Users\\Administrator\\Desktop\\'), '应落桌面而非 Documents: ' + out);
    assert.ok(path.win32.basename(out).startsWith('周会材料_'), '文件名取文档名: ' + out);
});

test('error JSON 绝不能被当成路径（否则会建出 {\\"success\\":false... 目录）', async () => {
    const { App } = loadChatV2();
    // 桥既给不出桌面也给不出文档路径（未打开任何文件）
    App.excelBridge = {
        GetDesktopPath: () => '{"success":false,"error":"operation not supported on this host"}',
        GetPresentationInfo: () => '{"success":false,"error":"没有打开的演示文稿"}',
    };
    const out = await gen(App)({ userMessage: '导出 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.strictEqual(out, null, '必须返回 null 而不是把错误 JSON 当路径');
});

test('网络盘文档无法反推 profile 时返回 null', async () => {
    const { App } = loadChatV2();
    App.excelBridge = stubHostBridge('\\\\server\\share\\团队\\季报.pptx', '季报.pptx');
    const out = await gen(App)({ userMessage: '导出 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.strictEqual(out, null);
});

// ============================================================
// 【根因回归】COM 斜杠丢失（2026-10-05 23:45 定位的真根因）
// ------------------------------------------------------------
// 现象：后端 `C:\Users\Administrator\Desktop` 经 WebView2 COM hostObjects 代理往返后，
//       到前端变成 `C:/Users/Administrator/Desktop`（反斜杠被吃；JSON 里的 `\\` 变 `//`）。
// 后果：① `dir.includes('\\')` 判定为假 → 拼出 `C://Users//...` 非法路径；
//       ② `desktopFromUserProfile` 的 `^([A-Za-z]:\\Users\\...)` 正则失配 → 目录解析返回空
//       → 回落桌面 → 导出拿到空/非法路径 → "filePath 不能为空" / "缺少必要参数"。
// 为什么之前没测出来：桩里直接给了带反斜杠的字符串，**没有复现 COM 的斜杠丢失**。
// 下面前置桩一律先过 com() 处理，锁死这个形态。
// ============================================================

/** 模拟 WebView2 COM 代理对字符串的往返：反斜杠被替换成正斜杠 */
function com(s) { return String(s).replace(/\\/g, '/'); }

test('COM 斜杠丢失：裸路径桥仍应产出合法 Windows 反斜杠绝对路径', async () => {
    const { App } = loadChatV2();
    App.excelBridge = { GetDesktopPath: () => com('C:\\Users\\Administrator\\Desktop') };
    const out = await gen(App)({ userMessage: '将演示文稿导出为 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out, '不应返回 null');
    assert.ok(/^C:\\Users\\Administrator\\Desktop\\/.test(out), '应是单反斜杠 Windows 绝对路径，实际: ' + out);
    assert.ok(!out.includes('//'), '不应出现正斜杠，实际: ' + out);
    assert.ok(out.endsWith('.pdf'), '扩展名应为 .pdf，实际: ' + out);
});

test('COM 斜杠丢失：旧 DLL（无 GetDesktopPath override）靠文档路径反推桌面', async () => {
    const { App } = loadChatV2();
    App.excelBridge = {
        GetDesktopPath: () => com('{"success":false,"error":"operation not supported on this host"}'),
        GetPresentationInfo: () => com(JSON.stringify({
            success: true, name: '季度汇报.pptx',
            fullName: 'C:\\Users\\Administrator\\Documents\\季度汇报.pptx', slides: 12
        })),
    };
    const out = await gen(App)({ userMessage: '将演示文稿导出为 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out, '旧 DLL 也必须能反推出桌面，不应返回 null');
    assert.ok(out.startsWith('C:\\Users\\Administrator\\Desktop\\'), '应反推当前用户桌面，实际: ' + out);
    assert.ok(out.includes('季度汇报'), '文件名应取当前文档名，实际: ' + out);
});

test('COM 斜杠丢失：Excel 导出 CSV 不得产出 C://Users// 这类非法路径', async () => {
    const { App } = loadChatV2();
    App.excelBridge = {
        GetDesktopPath: () => com('C:\\Users\\Administrator\\Desktop'),
        GetDownloadsPath: () => com('C:\\Users\\Administrator\\Downloads'),
    };
    const out = await App.ChatV2.toolRegistryUtil.inferExportFilePath(
        { userMessage: '将当前工作表导出为 CSV 并保存到桌面', context: {} });
    assert.ok(out, '不应返回 null');
    assert.ok(out.startsWith('C:\\Users\\Administrator\\Desktop\\'), '应为合法反斜杠路径，实际: ' + out);
    assert.ok(!out.includes('//'), '不得出现正斜杠，实际: ' + out);
    assert.ok(out.endsWith('.csv'), '扩展名应为 .csv，实际: ' + out);
});

test('COM 斜杠丢失：用户消息里用正斜杠写的目录也必须被尊重（不得静默落桌面）', async () => {
    const { App } = loadChatV2();
    App.excelBridge = { GetDesktopPath: () => com('C:\\Users\\Administrator\\Desktop') };
    const out = await gen(App)({ userMessage: '保存到 D:/季度报告/最终稿.pdf', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.ok(out, '不应返回 null');
    assert.ok(out.startsWith('D:\\季度报告\\'), '应尊重用户指定的 D 盘目录，实际: ' + out);
    assert.ok(!out.includes('Administrator'), '不应回落到当前用户桌面，实际: ' + out);
    assert.ok(!out.includes('最终稿'), '文件名不直接用用户的（避免覆盖），实际: ' + out);
});

test('normalizePathSeparators：压缩重复分隔符并保留 UNC 前缀', () => {
    const { App } = loadChatV2();
    const n = App.ChatV2.toolRegistryUtil.normalizePathSeparators;
    assert.strictEqual(n('C:/Users/Administrator/Desktop'), 'C:\\Users\\Administrator\\Desktop');
    assert.strictEqual(n('C:////Users////Administrator'), 'C:\\Users\\Administrator');
    assert.strictEqual(n('//server/share/a.xlsx'), '\\\\server\\share\\a.xlsx', 'UNC 前导双反斜杠必须保留');
    assert.strictEqual(n('D:/季度报告/'), 'D:\\季度报告', '尾部多余分隔符压掉');
    assert.strictEqual(n(''), '');
    assert.strictEqual(n(null), '');
    assert.strictEqual(n(undefined), '');
});

test('joinWinPath：永远用反斜杠拼接，不猜分隔符', () => {
    const { App } = loadChatV2();
    const j = App.ChatV2.toolRegistryUtil.joinWinPath;
    assert.strictEqual(j('C:/Users/Administrator/Desktop', 'a.pdf'), 'C:\\Users\\Administrator\\Desktop\\a.pdf');
    assert.strictEqual(j('C:\\Users\\Administrator\\Desktop\\', 'a.pdf'), 'C:\\Users\\Administrator\\Desktop\\a.pdf');
    assert.strictEqual(j('C:/Users/Admin/Desktop', '/a.pdf'), 'C:\\Users\\Admin\\Desktop\\a.pdf', '文件名侧多余分隔符应被吃掉');
    assert.strictEqual(j('', 'a.pdf'), 'a.pdf');
});

test('COM 化后的 error JSON 绝不能被当成目录使用', async () => {
    const { App } = loadChatV2();
    App.excelBridge = {
        GetDesktopPath: () => com('{"success":false,"error":"operation not supported on this host"}'),
        GetPresentationInfo: () => com('{"success":false,"error":"没有打开的演示文稿"}'),
    };
    const out = await gen(App)({ userMessage: '导出 PDF', context: {} }, { ext: 'pdf', base: '演示文稿' });
    assert.strictEqual(out, null, '拿不到目录时必须返回 null，不能建出 {"success":false... 目录');
});

test('tool-registry 已注册 6 个导出工具的路径生成器键', () => {
    const src = fs.readFileSync(path.join(BASE, 'services', 'tool-registry.js'), 'utf8');
    const keys = ['ppt_export_pdf.filePath', 'ppt_export_pdf.file_path', 'word_export_pdf.filePath',
        'word_export_pdf.file_path', 'ppt_export_image.filePath', 'ppt_export_image.file_path'];
    for (const k of keys) {
        assert.ok(src.includes("'" + k + "':"), '缺少生成器注册: ' + k);
    }
    // util 必须真的导出该函数（否则生成器静默返回 undefined，行为退回旧问题）
    const util = fs.readFileSync(path.join(BASE, 'services', 'tool-registry.util.js'), 'utf8');
    assert.ok(util.includes('inferHostExportFilePath,'), 'toolRegistryUtil 未导出 inferHostExportFilePath');
});
