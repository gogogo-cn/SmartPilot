/**
 * i18n.js — 轻量多语言 UI 文案（主路径扩展）
 *
 * App.i18n.t(key) / setLocale(id) / apply()
 * DOM：data-i18n / data-i18n-title / data-i18n-placeholder / data-i18n-aria / data-i18n-value
 *
 * 追加一种界面语言：在 LOCALE_CATALOG 加一项（并在 HostRibbon.vb 的 LocaleCatalog 加同一短码），
 * 再在 STRINGS 里补同名语言包。归一化、回退、html lang、徽标、下拉框都读这份目录。
 *
 * 2026-10-11 语言切换优化：
 *   - 双向回退（en 缺词回 zh，zh 缺词回 en），缺失只回退到 key 本身，不再静默留空白；
 *   - setLocale 支持 opts.broadcast —— 通过 postMessage 同步父窗口 / 子 iframe / opener，
 *     解决「AI 设置抽屉是 iframe、主窗口改了语言抽屉不跟着变」的老问题；
 *   - 新增 toggleLocale / addLocaleChangeListener / missingKeys / has 等 API，便于 UI 层接哨兵；
 *   - apply(root) 支持按作用域刷新 + <option>/ input.value 场景。
 */
(function () {
    'use strict';

    var STORAGE_KEY = 'smartpilot_ui_locale';

    /**
     * 界面语言目录。id 是落盘短码（localStorage / ui-locale.txt）。
     * aliases 用来把 zh-CN、en-US 这类标签收成短码；更长的地区标签靠「短码或别名 + '-'」前缀匹配。
     * 数组顺序就是工具栏一键切换的轮换顺序，第一项是缺省回退语言。
     */
    var LOCALE_CATALOG = [
        { id: 'zh', htmlLang: 'zh-CN', dateLocale: 'zh-CN', badge: '中', nativeName: '中文', aliases: ['zh', 'zh-cn', 'zh-hans', 'chinese'] },
        { id: 'en', htmlLang: 'en', dateLocale: 'en-US', badge: 'EN', nativeName: 'English', aliases: ['en', 'en-us', 'en-gb', 'english'] }
    ];
    var FALLBACK_LOCALE = LOCALE_CATALOG[0].id;
    var LOCALES = LOCALE_CATALOG.map(function (d) { return d.id; });

    var STRINGS = {
        zh: {
            'app.title': 'SmartPilot 智航',
            'tab.chat': '💬对话',
            'tab.history': '📜历史',
            'tab.skills': '🛠️ 扩展',
            'tab.knowledge': '📚 知识库',
            'nav.chat': '对话',
            'nav.history': '历史',
            'nav.ocr': '识图',
            'nav.more': '更多',
            'nav.settingsShort': '设置',
            'nav.helpShort': '帮助',
            'nav.accountShort': '账户',
            'nav.ocrModel': 'OCR 模型',
            'nav.mobile': '手机遥控',
            'nav.skills': '扩展',
            'nav.knowledge': '知识库',
            'nav.newChat': '新建对话',
            'nav.taskPlanner': '任务规划',
            'nav.editor': '代码编辑器',
            'nav.quickPrompts': '快捷提示词',
            'nav.pomodoro': '独立番茄钟（数据同步）',
            'nav.pomodoroShort': '番茄钟',
            'nav.pptTheme': 'PPT 主题模板',
            'nav.export': '导出当前会话',
            'nav.help': '帮助 (F1)',
            'nav.clearCache': '清空缓存',
            'nav.settings': 'AI 模型设置',
            'nav.account': '账户 / 订阅',
            'nav.collapseToolbar': '收起 / 展开工具栏',
            'nav.scrollUp': '向上滚动',
            'nav.scrollDown': '向下滚动',
            'nav.undo': '撤销插件上一步（不是工作表里的 Ctrl+Z）',
            'nav.redo': '重做插件上一步（不是工作表里的 Ctrl+Y）',
            'status.detecting': '检测中',
            'status.ready': '系统就绪',
            'conn.title': 'Excel 连接状态',
            'conn.checking': '正在检测...',
            'conn.connected': '已连接',
            'conn.connecting': '连接中',
            'conn.disconnected': '已断开',
            'conn.excelClosed': 'Excel已关闭',
            'conn.excelNotOpen': 'Excel未打开',
            'conn.serviceDown': '服务未启动',
            'conn.cannotConnect': '无法连接Excel',
            'conn.linkLost': '连接断开',
            'conn.mode': '模式:',
            'conn.time': '时间:',
            'conn.justNow': '刚刚',
            'conn.minsAgo': '{n}分钟前',
            'conn.hoursAgo': '{n}小时前',
            'conn.modeVsto': 'VSTO 内嵌（当前 Excel）',
            'conn.modeBridge': '本地桥 (pywin32)',
            'conn.pickWorkbook': '选择Excel工作簿',
            'conn.pickSheet': '选择工作表',
            'conn.workbook': '工作簿:',
            'conn.activeSheet': '活动工作表:',
            'conn.sheetCount': '工作表数量:',
            'conn.sheetUnit': '{n}个',
            'conn.activeWb': '(活动工作簿)',
            'conn.loading': '加载中',
            'conn.noWorkbooks': '未找到打开的工作簿，请先在Excel中打开文件',
            'conn.connect': '连接到 Excel',
            'conn.refreshList': '刷新工作簿列表',
            'conn.disconnect': '断开连接',
            'conn.refresh': '刷新状态',
            'conn.confirmDisconnect': '确定断开？',
            'conn.hostHint': '已内嵌于 Excel，工作簿切换会直接作用于当前窗口',
            'conn.refreshed': '已刷新',
            'conn.switchedSheet': '已切换工作表: {name}',
            'conn.switchedWb': '已切换到: {name}',
            'conn.connectOk': '已连接 Excel',
            'conn.needExcel': '请先打开 Excel',
            'conn.needConnect': '请先连接 Excel',
            'conn.loadWbFail': '加载工作簿失败: {msg}',
            'conn.switchSheetFail': '切换工作表失败: {msg}',
            'conn.switchWbFail': '切换失败: {msg}',
            'conn.pickWbWarn': '请选择工作簿',
            'conn.excelClosedToast': 'Excel 已关闭，请重新连接',
            'clear.title': '选择要清空的数据',
            'clear.history': '对话历史记录',
            'clear.historyDesc': '所有聊天记录和对话历史',
            'clear.config': '应用配置',
            'clear.configDesc': 'API设置、模型选择等配置',
            'clear.types': '自定义识别类型',
            'clear.typesDesc': '用户添加的自定义识别类型',
            'clear.snippets': '代码片段',
            'clear.snippetsDesc': '保存的代码片段和代码库',
            'clear.cancel': '取消',
            'clear.confirm': '确认清空',
            'tools.localExcel': '本地 Excel',
            'tools.localPpt': '本地 PPT',
            'tools.localWord': '本地 Word',
            'tools.localPrefix': '本地 {name}',
            'tools.real': '真实',
            'tools.mock': '模拟',
            'tools.expand': '点击展开',
            'tools.call': '调用',
            'tools.loading': '工具加载中，请稍候...',
            'ocr.title': '识图入表',
            'ocr.more': '批量 / 解析 / 历史',
            'ocr.localHint': '本机识别可以先用。云端 OCR 和对话需要账户；付费开通前，未登录不影响本机识别。',
            'ocr.tab.recognize': '识别',
            'ocr.tab.batch': '批量',
            'ocr.tab.parser': '智能解析',
            'ocr.tab.history': '历史',
            'ocr.mobile': '手机遥控',
            'ocr.engine': '引擎',
            'ocr.lang': '语言',
            'ocr.lang.auto': '自动检测',
            'ocr.lang.zh': '简体中文',
            'ocr.lang.zht': '繁体中文',
            'ocr.lang.en': '英文',
            'ocr.lang.ja': '日文',
            'ocr.lang.ko': '韩文',
            'ocr.lang.zhOnly': '当前引擎按中文识别',
            'ocr.recognizeThis': '识别这张图',
            'ocr.drop': '拖拽图片或 PDF 到此处',
            'ocr.dropHint': '支持单文件或多文件同时上传',
            'ocr.pickFile': '选择文件',
            'ocr.pickFolder': '选择文件夹',
            'ocr.screenshot': '截图识别',
            'ocr.preview': '图片预览',
            'ocr.result': '识别结果',
            'ocr.view.text': '原文',
            'ocr.view.list': '列表',
            'ocr.view.preview': '预览',
            'ocr.format': '格式',
            'ocr.format.plain': '纯文本',
            'ocr.format.raw': 'OCR原文',
            'ocr.format.excel': 'Excel 表格',
            'ocr.format.preview': '预览',
            'ocr.start': '开始识别',
            'ocr.insert': '写入当前表',
            'ocr.saveExcel': '另存 Excel',
            'ocr.savedExcel': '已另存 Excel（{n} 行）',
            'ocr.noExportData': '没有可导出的数据',
            'ocr.insertFields': '按字段写入',
            'ocr.export': '导出',
            'ocr.ready': '就绪',
            'ocr.settingsTitle': 'OCR 识别设置',
            'ocr.settingsEngine': 'OCR 识别引擎',
            'ocr.engineStatus': '引擎状态',
            'ocr.detecting': '正在检测...',
            'ocr.paddleLang': '识别语言',
            'ocr.layout': '版面分析',
            'ocr.table': '表格识别',
            'ocr.formula': '公式识别',
            'ocr.verifyPaddle': '验证 PaddleOCR 安装',
            'ocr.saveSettings': '保存 OCR 设置',
            'ocr.opt.rapid': 'RapidOCR - 嵌入式，轻量快速（约 30MB）',
            'ocr.opt.paddle': 'PaddleOCR - 百度开源，中文识别效果最佳',
            'ocr.opt.cloud': '云端 OCR API - 识别准确率最高',
            'ocr.lang.ch': '中文（默认）',
            'ocr.modelSettings': '模型设置',
            'ocr.opt.rapidShort': 'RapidOCR（嵌入式）',
            'ocr.opt.paddleShort': 'PaddleOCR（中文推荐）',
            'ocr.opt.cloudShort': '云端 OCR',
            'ocr.recognizeAll': '全部识别',
            'ocr.mechanical': '机械图纸',
            'ocr.mechanicalTitle': '只在机械图纸上补形位公差符号。不勾选时，RapidOCR 和 PaddleOCR 都保持各自的原文',
            'ocr.packs.manage': '专项模块',
            'ocr.pickScript': '选择解析脚本...',
            'ocr.batchEmpty': '请选择图片或 PDF 文件',
            'ocr.batchEmptyHint': '支持 PNG / JPG / PDF 等格式',
            'ocr.batchPickWait': 'OCR 引擎初始化中，请稍候',
            'ocr.batchStart': '开始批量处理',
            'ocr.batchPause': '暂停',
            'ocr.batchCancel': '取消',
            'ocr.batchInsert': '插入到 Excel',
            'ocr.batchParse': '批量解析',
            'ocr.removeImage': '删除图片',
            'ocr.parser.title': '解析脚本',
            'ocr.parser.clear': '清空缓存',
            'ocr.parser.upload': '上传',
            'ocr.parser.uploadTitle': '上传 MD 文件',
            'ocr.parser.new': '新建',
            'ocr.parser.searchPh': '搜索脚本...',
            'ocr.parser.allCategories': '全部分类',
            'ocr.parser.loading': '加载中',
            'ocr.parser.edit': '脚本编辑',
            'ocr.parser.preview': '测试预览',
            'ocr.parser.fieldsTab': '字段配置',
            'ocr.parser.name': '脚本名称 *',
            'ocr.parser.namePh': '如：增值税发票解析',
            'ocr.parser.desc': '描述',
            'ocr.parser.descPh': '简要描述这个脚本的用途',
            'ocr.parser.icon': '图标',
            'ocr.parser.category': '分类',
            'ocr.parser.cat.general': '通用',
            'ocr.parser.cat.receipt': '票据',
            'ocr.parser.cat.id': '证件',
            'ocr.parser.cat.business': '商务',
            'ocr.parser.cat.custom': '自定义',
            'ocr.parser.output': '输出格式',
            'ocr.parser.out.table': '表格（两列）',
            'ocr.parser.out.kv': '键值对',
            'ocr.parser.code': '脚本代码',
            'ocr.parser.save': '保存脚本',
            'ocr.parser.reset': '还原默认',
            'ocr.parser.test': '测试运行',
            'ocr.parser.testInput': '输入测试文本（OCR 识别结果）',
            'ocr.parser.testPh': '在此粘贴 OCR 识别结果文本，或使用示例文本...',
            'ocr.parser.loadSample': '加载示例文本',
            'ocr.parser.run': '执行解析',
            'ocr.parser.resultEmpty': '执行解析后结果将显示在这里',
            'ocr.parser.copy': '复制',
            'ocr.parser.insert': '插入 Excel',
            'ocr.parser.fields': '字段定义',
            'ocr.parser.addField': '添加字段',
            'ocr.parser.noFields': '暂无字段定义，点击上方按钮添加',
            'ocr.history.title': '识别历史',
            'ocr.history.clear': '清空全部',
            'ocr.history.empty': '暂无识别记录',
            'ocr.history.emptyHint': '完成图片识别后会自动保存到这里',
            'ocr.desc.rapid': 'ONNX 轻量级 OCR，模型小（约 30MB），启动快，随插件分发',
            'ocr.desc.paddle': '百度开源 OCR，中文识别效果最佳，需要安装 Python',
            'ocr.desc.cloud': '云端 OCR 服务，识别准确率高，需要 API Key 和联网',
            'ocr.rapid.intro': 'RapidOCR 是基于 ONNX 的轻量级 OCR 引擎，模型小（约 30MB），启动快，随插件预装',
            'ocr.rapid.checking': '正在检测 RapidOCR...',
            'ocr.rapid.features': 'RapidOCR 特性',
            'ocr.rapid.f1': '基于 ONNX Runtime，无需 Python 环境即可运行',
            'ocr.rapid.f2': '模型体积小（约 30MB），下载安装快速',
            'ocr.rapid.f3': '支持中文、英文等多种语言',
            'ocr.rapid.f4': '启动速度快，适合批量处理',
            'ocr.rapid.check': '检查 RapidOCR 状态',
            'ocr.paddle.intro': 'PaddleOCR 是百度开源的 OCR 引擎，对中文识别效果最佳',
            'ocr.paddle.checking': '正在检测 PaddleOCR...',
            'ocr.paddle.advanced': 'PaddleOCR 高级功能（需 PP-OCRv6 + PP-StructureV3）',
            'ocr.paddle.hint': '新版 PP-OCRv6 支持多种语言统一识别。版面分析识别文档结构，表格识别还原行列，公式识别输出 LaTeX。',
            'ocr.paddle.noPython': '未检测到 Python，PaddleOCR 需要 Python 3.9 或更高版本',
            'ocr.paddle.downloadPython': '下载 Python',
            'ocr.paddle.install': '自动安装 PaddleOCR',
            'ocr.paddle.diagnoseLabel': '诊断结果：',
            'ocr.paddle.log': '安装日志：',
            'ocr.paddle.manualTitle': '手动安装指南（高级用户）',
            'ocr.paddle.manualHint': '如果一键安装也失败，可以手动执行以下命令：',
            'ocr.win.intro': '使用 Windows 10/11 内置的 OCR 功能，无需安装任何额外软件',
            'ocr.win.checking': '正在检测 Windows OCR...',
            'ocr.win.langs': '支持的语言',
            'ocr.cloud.intro': '使用云端 OCR API，识别准确率最高，需要联网和配置 API Key',
            'ocr.cloud.provider': '云服务提供商',
            'ocr.cloud.baidu': '百度智能云 OCR',
            'ocr.cloud.tencent': '腾讯云 OCR',
            'ocr.cloud.baiduKeyPh': '输入百度智能云 API Key',
            'ocr.cloud.baiduSecretPh': '输入百度智能云 Secret Key',
            'ocr.cloud.baiduLink': '前往百度智能云获取 API Key',
            'ocr.cloud.tencentIdPh': '输入腾讯云 SecretId',
            'ocr.cloud.tencentKeyPh': '输入腾讯云 SecretKey',
            'ocr.cloud.tencentLink': '前往腾讯云获取 API Key',
            'ocr.cloud.test': '连接测试',
            'ocr.cloud.testBtn': '测试连接',
            'ocr.status.checking': '正在检测...',
            'ocr.status.waitInit': '等待初始化...',
            'ocr.status.available': '可用',
            'ocr.status.unavailable': '不可用',
            'ocr.status.ready': '就绪',
            'ocr.status.needConfig': '需要配置',
            'ocr.status.rapidOk': 'RapidOCR 已安装并可用',
            'ocr.status.rapidBad': 'RapidOCR 不可用',
            'ocr.status.checkPlugin': '请检查插件安装',
            'ocr.status.fail': '检测失败',
            'ocr.status.checkBackend': '请检查后端服务',
            'ocr.status.notSelected': '未选择',
            'ocr.status.pickEngine': '请从上方选择 OCR 引擎',
            'ocr.langId.ch': '中文',
            'ocr.langId.en': '英文',
            'ocr.langId.japan': '日语',
            'ocr.langId.korean': '韩语',
            'ocr.langId.fr': '法语',
            'ocr.langId.ger': '德语',
            'ocr.langId.es': '西班牙语',
            'ocr.langId.ru': '俄语',
            'ocr.langId.arabic': '阿拉伯语',
            'ocr.langId.structure': '自动检测',
            'ocr.fileHint': '本地文件无法直接连接识图服务。打开后会进入本机页面。',
            'ocr.openOcr': '打开识图',
            'ocr.desc.default': '可选 OCR 引擎，支持多种语言，本地运行无需联网',
            'ocr.formulaInstall': '安装公式增强包',
            'ocr.formulaReady': '公式识别已就绪',
            'ocr.formulaInstalling': '正在安装（约 3-5 分钟）...',
            'ocr.pdfPrev': '上一页',
            'ocr.pdfNext': '下一页',
            'ocr.pdfPause': '暂停',
            'ocr.pdfResume': '继续',
            'ocr.pdfCancel': '取消',
            'ocr.resultPh': '识别结果将显示在这里...',
            'ocr.formatHint': '选择导出到 Excel 或复制时的输出格式',
            'ocr.copyTitle': '复制',
            'ocr.clearTitle': '清空',
            'ocr.remove': '删除',
            'ocr.export.excel': '导出为 Excel 表格',
            'ocr.export.csv': '导出为 CSV',
            'ocr.export.text': '导出为文本文件',
            'ocr.batch.track': '跟踪',
            'ocr.batch.trackTitle': '处理过程中自动滚动显示结果',
            'ocr.batch.total': '总计',
            'ocr.batch.done': '完成',
            'ocr.batch.failed': '失败',
            'ocr.batch.clear': '清空',
            'ocr.batch.summary': '识别结果汇总',
            'ocr.batch.copyAll': '全部复制',
            'ocr.batch.emptyFiles': '请选择图片文件或文件夹',
            'ocr.batch.addFile': '选择文件',
            'ocr.batch.addFolder': '选择文件夹',
            'ocr.batch.remove': '移除',
            'ocr.batch.pending': '等待处理',
            'ocr.batch.processing': '处理中',
            'ocr.batch.completed': '完成',
            'ocr.batch.failedMsg': '失败: {msg}',
            'ocr.batch.unknownError': '未知错误',
            'ocr.batch.engineWait': 'OCR 引擎初始化中，请稍候再试',
            'ocr.parser.pickOrNew': '选择或新建脚本',
            'ocr.parser.builtin': '内置',
            'ocr.parser.unnamed': '未命名脚本',
            'ocr.parser.parsing': '正在执行解析...',
            'ocr.parser.noData': '未提取到任何数据',
            'ocr.parser.parseFail': '解析失败: {msg}',
            'ocr.parser.noMatch': '没有匹配的脚本',
            'ocr.parser.none': '暂无脚本，点击上方按钮新建',
            'ocr.recognizeAllCount': '全部识别（{n}）',
            'ocr.status.engineReady': '引擎就绪',
            'ocr.status.engineNotReady': 'OCR 引擎未就绪：{msg}',
            'ocr.status.hostDown': 'OCR 服务未启动',
            'ocr.status.unknownEngine': '未知引擎类型',
            'ocr.status.noPython': '未检测到 Python',
            'ocr.status.rapidNeedPy': 'RapidOCR 不可用，请检查 Python 环境',
            'ocr.status.paddleNoModel': 'PaddleOCR 已安装，但模型未下载（请重新安装以触发模型下载）',
            'ocr.status.cloudName': '云端 OCR',
            'ocr.status.cloudReady': '云端 OCR 已配置',
            'ocr.status.cloudNeedKey': '云端 OCR 未配置，请在设置中填写 API Key',
            'ocr.paddle.installed': 'PaddleOCR 已安装 ({name})',
            'ocr.paddle.missingCount': 'PaddleOCR 未安装（已找到 {n} 个 Python，均不可用）',
            'ocr.paddle.ok': 'PaddleOCR 已安装并可用',
            'ocr.paddle.notInstalled': 'PaddleOCR 未安装',
            'ocr.paddle.cannotCheck': '无法检测状态',
            'ocr.status.installing': '正在安装...',
            'ocr.status.startingInstall': '正在启动安装...',
            'ocr.noResult': '暂无识别结果',
            'ocr.noText': '（未识别到文字）',
            'ocr.fit': '适应',
            'ocr.mode.label': '识别模式',
            'ocr.mode.auto': '自动',
            'ocr.mode.text': '纯文本',
            'ocr.mode.formula': '公式',
            'ocr.mode.number': '数字',
            'ocr.mode.date': '日期',
            'ocr.mode.mechanical': '机械图纸',
            'ocr.adjusted': '已调整',
            'ocr.pageLabel': '第 {n} 页',
            'ocr.shotTitle': '第 {n} 张',
            'ocr.tool.expand': '展开工具栏',
            'ocr.tool.rotateLeft': '逆时针旋转 90 度',
            'ocr.tool.rotateRight': '顺时针旋转 90 度',
            'ocr.tool.flip': '水平翻转',
            'ocr.tool.zoomOut': '缩小',
            'ocr.tool.zoomIn': '放大',
            'ocr.tool.resetZoom': '重置缩放',
            'ocr.tool.viewOriginal': '放大查看原图',
            'ocr.tool.collapse': '折叠工具栏',
            'ocr.manual.step1': '# 1. 打开命令提示符或 PowerShell',
            'ocr.manual.step2': '# 2. 执行以下命令（替换为实际 Python 版本号）：',
            'ocr.manual.step3': '# 3. 下载 OCR 识别模型（重要！）：',
            'ocr.manual.done': '模型下载完成',
            'ocr.manual.tip': '提示：如果清华镜像失败，可尝试其他镜像源<br>阿里云 https://mirrors.aliyun.com/pypi/simple/<br>中科大 https://pypi.mirrors.ustc.edu.cn/simple/',
            'ocr.detailTitle': '识别详情',
            'ocr.reRecognize': '重新识别',
            'ocr.parseShort': '解析',
            'ocr.selectAll': '全选',
            'ocr.addRegion': '添加区域',
            'ocr.realtime': '实时识别',
            'ocr.recognizeSelected': '识别选中区域',
            'ocr.recognizeSelectedTitle': '识别左侧复选框勾选的区域',
            'ocr.duplicateScript': '复制脚本',
            'ocr.deleteScript': '删除脚本',
            'ocr.ctx.clearLatex': '清除 LaTeX 格式',
            'ocr.ctx.fixNumbers': '修正数字格式',
            'ocr.ctx.fixDates': '修正日期格式',
            'ocr.ctx.removeSpaces': '移除多余空格',
            'ocr.ctx.edit': '手动编辑',
            'ocr.ctx.deleteRegion': '删除区域',
            'ocr.recognizing': '识别中...',
            'ocr.recognizingNow': '正在识别...',
            'ocr.loadingEngine': '正在加载 OCR 引擎...',
            'ocr.analyzingImage': '正在分析图片...',
            'ocr.readingText': '正在识别文字（首次运行需加载模型，请耐心等待）...',
            'ocr.readingTextWait': '正在识别文字（模型加载中，已完成 {n} 秒）...',
            'ocr.shotAlt': '截图 {n}',
            'ocr.shotProgress': '第 {n}/{total} 张',
            'ocr.shotSwitch': '第 {n}/{total} 张，点缩略图可切换，Enter 识别全部',
            'ocr.parser.signature': '函数签名',
            'ocr.parser.fieldsHint': '字段用系统注入的',
            'ocr.batch.scanning': '正在扫描文件夹...',
            'ocr.batch.scanDone': '扫描完成，共 {n} 个文件',
            'ocr.batch.scanFail': '扫描失败: {msg}',
            'ocr.batch.noBackend': '无法连接后端服务',
            'ocr.batch.running': 'OCR 识别中 {n}/{total}...',
            'ocr.batch.fileProgress': '正在识别 ({n}/{total}): {name}',
            'ocr.batch.doneSummary': '批量完成: {ok} 成功, {bad} 失败',
            'ocr.batch.parsing': '正在解析 ({n}/{total}): {name}',
            'ocr.batch.parseDone': '解析完成: {ok} 成功, {bad} 失败',
            'ocr.batch.paused': '已暂停',
            'ocr.batch.resuming': '继续处理...',
            'ocr.batch.cancelled': '已取消',
            'ocr.batch.cancelledToast': '已取消批量处理',
            'ocr.batch.noFiles': '没有可解析的文件',
            'mobile.title': '手机遥控 OCR',
            'mobile.status': '服务状态',
            'mobile.stopped': '未启动',
            'mobile.running': '服务运行中',
            'mobile.openBrowser': '在浏览器打开',
            'mobile.copyLink': '复制链接',
            'mobile.start': '启动服务',
            'mobile.stop': '停止服务',
            'mobile.refreshQr': '刷新 QR 码',
            'mobile.helpTitle': '使用说明',
            'mobile.help': '1. 确保手机和电脑在同一 WiFi 网络<br>2. 点击「启动服务」按钮<br>3. 手机扫描二维码，打开拍照页面<br>4. 拍照后自动上传并 OCR 识别<br>5. 识别结果实时显示在手机上',
            'input.placeholder': '描述您的需求，或粘贴图片/表格...（可批量选择多个文件）',
            'input.send': '发送',
            'input.stop': '⏹ 停止',
            'input.upload': '上传文件（支持多选）',
            'input.reference': '添加引用区域（点击后在 Excel 中选择区域）',
            'history.search': '搜索历史记录...',
            'history.empty': '暂无历史记录',
            'history.pin': '置顶',
            'history.unpin': '取消置顶',
            'history.export': '导出该会话',
            'history.edit': '编辑标题',
            'history.delete': '删除',
            'history.pinnedBadge': '置顶',
            'history.groupPinned': '置顶',
            'history.groupToday': '今天',
            'history.groupYesterday': '昨天',
            'history.multiSelect': '多选删除',
            'history.deleteSelected': '删除所选',
            'history.clearAll': '清空所有记录',
            'history.selectAll': '全选',
            'history.selectedCount': '已选 {n} 条',
            'history.done': '完成',
            'history.tempStorage': '💡 当前使用临时存储，刷新后历史将重置',
            'history.pinnedToast': '已置顶',
            'history.unpinnedToast': '已取消置顶',
            'welcome.setup': '识图入表不用先配模型。截图、表格、表单和票据都能先写入当前表。清洗、拆分和出图时，再配置你自己的模型。',
            'gate.needLlm': '请先配置 AI 模型：打开设置填写 API 密钥，或启动 Ollama 并选择模型。',
            'welcome.openOcr': '先识别一张图片',
            'welcome.openSettings': '打开 AI 设置',
            'welcome.tryThese': '试试这些（点击填入输入框）',
            'welcome.helpLink': '查看快速开始与常见问题',
            'welcome.excel.subtitle': '把截图、表格、表单和票据收成一张表',
            'welcome.excel.desc': '先识别图片，写入当前工作表。要清洗、拆分或出图时，再接你自己的模型。订阅只收软件席位。',
            'welcome.word.subtitle': '智能 Word 文档助手',
            'welcome.word.desc': '用自然语言润色、总结、排版与导出文档，让写作与审阅更高效。',
            'welcome.ppt.subtitle': '智能 PowerPoint 演示助手',
            'welcome.ppt.desc': '用自然语言生成大纲、美化幻灯片、补充备注并导出，加速演示文稿制作。',
            'phase.turn': '正在准备…',
            'phase.step': '正在执行步骤…',
            'phase.thinking': '正在分析你的需求…',
            'phase.connecting': '正在连接 AI…',
            'phase.tool': '正在执行操作…',
            'phase.toolWrite': '正在写入 / 修改内容…',
            'phase.toolExport': '正在导出 / 保存文件…',
            'phase.toolRead': '正在读取数据…',
            'phase.summarizing': '正在整理结果…',
            'phase.done': '已完成',
            'phase.canceled': '已停止',
            'phase.processing': '处理中…',
            'phase.stepOf': '第 {n}/{total} 轮',
            'phase.runningTool': '正在执行：{name}',
            'progress.pending': '等待',
            'progress.running': '执行中',
            'progress.completed': '完成',
            'progress.failed': '失败',
            'danger.title': '⚠️ 危险操作确认',
            'danger.cancel': '取消',
            'danger.continue': '仍要继续执行',
            'danger.preview': '将影响的范围',
            'danger.hintDefault': '确认后，本轮对话内同类危险操作将不再询问。按 Esc 或点遮罩可取消。',
            'danger.hintAlwaysAsk': '已开启「始终询问」：本轮内每次危险操作都会再次确认。按 Esc 或点遮罩可取消。',
            'danger.followDefault': '（本轮对话中后续相同操作将沿用你的决定，不再重复询问）',
            'danger.followAlwaysAsk': '（已开启「始终询问」：本轮内每次危险操作都会再次确认）',
            'settings.title': '🤖 AI 模型设置',
            'settings.conn': '🔌 连接配置',
            'settings.provider': 'AI 提供商',
            'settings.apiUrl': 'API 地址',
            'settings.apiKey': 'API 密钥',
            'settings.apiKeyPh': '输入您的 API 密钥',
            'settings.modelCard': '🧠 模型选择',
            'settings.modelName': '模型名称',
            'settings.modelPh': '选择或输入模型名',
            'settings.detectLocal': '🔍 检测本地模型',
            'settings.params': '⚙️ 参数配置',
            'settings.temperature': '温度',
            'settings.context': '上下文长度',
            'settings.ctx.auto': '自动（推荐）',
            'settings.ctx.short': '简短（最近 1 轮）',
            'settings.ctx.medium': '中等（最近 3 轮）',
            'settings.ctx.long': '长对话（最近 10 轮）',
            'settings.ctx.unlimited': '完整历史',
            'settings.uiLocale': '界面语言',
            'settings.uiLocale.zh': '中文',
            'settings.uiLocale.en': 'English',
            'settings.dangerAsk': '危险操作',
            'settings.dangerAskLabel': '始终询问确认',
            'settings.dangerAskHint': '开启后，本轮内每次删除/清空等危险操作都会再次弹窗，不再沿用「允许」。',
            'settings.testCard': '连接测试',
            'settings.testBtn': '测试连接',
            'settings.testHint': '点击上方按钮测试 API 连接是否正常',
            'settings.save': '💾 保存 AI 设置',
            'settings.dataCard': '数据与风险说明',
            'settings.data.cloud': '使用云端模型时，你输入的文字、当前选中的表格或文档片段、以及近期对话会发送到你填写的 API 地址。该地址可能在本机以外，也可能在境外。',
            'settings.data.local': '提供商选 Ollama 或 llama.cpp，且地址指向本机时，提示词默认留在本机。把地址改成远程服务器后，数据会发往该服务器。',
            'settings.data.tools': '部分工具会修改当前工作簿、文档或幻灯片，导出文件，或在本机运行 VBA / Python。删除、清空一类操作会先请你确认。',
            'settings.data.keys': 'API 密钥只保存在本机当前用户下，不会打进安装包。请勿把 llm_config.json 发给他人。企业可把自定义 API 指到私有化模型。',
            'settings.undoNote': '聊天窗格上的撤销是插件自己的上一步，撤不回你在工作表里用 Ctrl+Z 能撤的编辑。',
            'settings.deleteLocal': '删除本机密钥和识别记录',
            'settings.deleteLocalConfirm': '再点一次，确认删除',
            'settings.deleteLocalDone': '已删除本机密钥和识别记录。工作簿里已写入的内容还在。',
            'settings.deleteLocalFail': '删除没有完成，请稍后再试。',
            'settings.closeEsc': '关闭 (Esc)',
            'settings.provider.ollama': 'Ollama (本地)',
            'settings.provider.llamacpp': 'llama.cpp (本地)',
            'settings.localGuide.title': '📥 本地模型安装引导',
            'settings.localGuide.statusNone': '尚未检测到本地模型。安装完成后点「检测本地模型」。',
            'settings.localGuide.statusOk': '已检测到 {count} 个本地模型，可在上方模型列表选择。',
            'settings.localGuide.ollamaHint': 'Ollama 免费开源，安装后模型的下载与推理都在本机完成，断网也能用。',
            'settings.localGuide.ollamaStep1': '下载并安装 Ollama（Windows 安装包约 700MB），按提示装完即可。',
            'settings.localGuide.downloadOllama': '⬇ 下载 Ollama',
            'settings.localGuide.ollamaStep2': '打开 PowerShell，拉取一个模型（首次下载约 1～4GB）：',
            'settings.localGuide.ollamaPullCmd': 'ollama pull qwen2.5:7b',
            'settings.localGuide.copy': '复制',
            'settings.localGuide.copied': '已复制',
            'settings.localGuide.copyFail': '复制失败，请手动选择命令',
            'settings.localGuide.openFail': '无法自动打开浏览器，请手动复制链接到浏览器打开：',
            'settings.localGuide.ollamaStep3': 'Ollama 装好后会常驻后台自启。若仍提示无法连接，先手动启动服务：',
            'settings.localGuide.ollamaServeCmd': 'ollama serve',
            'settings.localGuide.llamacppHint': 'llama.cpp 免费开源，通过 llama-server 提供 OpenAI 兼容接口（默认 http://localhost:8080）。',
            'settings.localGuide.llamacppStep1': '下载 llama.cpp Release（Windows 压缩包，解压即用）：',
            'settings.localGuide.downloadLlamacpp': '⬇ 下载 llama.cpp',
            'settings.localGuide.llamacppStep2': '从 Hugging Face 下载 GGUF 模型文件，放入解压目录的 models 文件夹。',
            'settings.localGuide.llamacppStep3': '在解压目录打开 PowerShell 启动服务，保持窗口开启：',
            'settings.localGuide.llamacppServeCmd': 'llama-server -m models\\qwen2.5-7b-instruct-q4_k_m.gguf --port 8080',
            'settings.localGuide.detectNow': '🔍 检测本地模型',
            'settings.provider.wenxin': '百度文心一言',
            'settings.provider.qianwen': '阿里通义千问',
            'settings.provider.zhipu': '智谱 AI (ChatGLM)',
            'settings.provider.xunfei': '讯飞星火',
            'settings.provider.moonshot': 'Moonshot (月之暗面)',
            'settings.provider.deepseek': '深度求索 (DeepSeek)',
            'settings.provider.volcengine': '火山引擎 (豆包)',
            'settings.provider.custom': '自定义 API',
            'settings.saved': '设置已保存',
            'license.card': '账户 / 订阅',
            'license.byok': '订阅覆盖 SmartPilot 软件许可（含 AI 对话与 OCR 等模块）。模型与云 OCR API Key 由你自行配置，费用与对应服务商结算。',
            'license.scope': '软件许可覆盖 AI 与 OCR',
            'license.openAccount': '打开账户与订阅',
            'license.badge.trial': '试用 / 未订阅',
            'license.badge.active': '已订阅',
            'license.badge.none': '账户',
            'license.status.loading': '正在读取订阅状态…',
            'license.status.none': '未登录',
            'license.status.active': '有效：{plan}，到期 {end}',
            'license.status.activeOffline': '已激活（激活码）：{plan}，到期 {end}',
            'license.status.inactive': '不可用（{reason}）',
            'license.login': '邮箱或手机',
            'license.password': '密码',
            'license.btnLogin': '登录',
            'license.btnRegister': '注册试用',
            'license.btnRefresh': '刷新状态',
            'license.btnStripeMonth': '升级（Stripe 月付）',
            'license.btnStripeYear': '升级（Stripe 年付）',
            'license.btnCnYear': '升级（国内年付）',
            'license.btnPortal': '管理账单',
            'license.btnLogout': '退出登录',
            'license.key': '激活码',
            'license.key.hint': '购买后收到的激活码（SPK1- 开头），粘贴后点「激活」。激活码在本机校验，不需要联网，也不需要登录账户。',
            'license.key.hintRenew': '已激活中。续费时点「输入新激活码」，粘贴新码后点「激活」即可延长有效期。',
            'license.key.placeholder': 'SPK1-…',
            'license.btnActivate': '激活',
            'license.btnRenewKey': '输入新激活码',
            'license.btnClearKey': '清除激活',
            'license.local.active': '离线激活：{plan}，到期 {end}',
            'license.local.none': '未使用激活码',
            'license.key.ok': '激活成功',
            'license.key.cleared': '已清除本地激活',
            'ocr.quota.exhausted': '今日免费额度已用完（{limit} 次/天），明天 0 点自动重置；激活后不限次数。',
            'ocr.quota.badge': '今日剩余 {remaining}/{limit}',
            'ocr.quota.unlimited': '已激活 · 不限次数',
            'ocr.quota.cloudFree': '云端识别不计额度',
            'license.btnRefund': '注销并申请退款',
            'license.refund.confirm': '确定注销本机激活吗？注销后这台机器立刻失去授权，且这张激活码不能再在本机激活。你会拿到一个「消 key 码」，发给客服验证后才会退款。',
            'license.refund.codeTitle': '消 key 码（发给客服完成退款）',
            'license.refund.codeHint': '把它复制发给客服。客服验证通过后会为你退款。此码不能恢复授权——如需再用请重新购买。',
            'license.refund.copy': '复制消 key 码',
            'license.refund.copied': '已复制，发给客服即可',
            'license.refund.done': '已注销，本机授权已失效',
            'license.refund.none': '当前没有可注销的离线激活',
            'license.refund.failed': '注销失败',
            'license.key.empty': '请先粘贴激活码',
            'license.buy.title': '购买 / 续费',
            'license.buy.titleRenew': '续费',
            'license.buy.price': '激活码方案（付款时微信里说明选哪个）',
            'license.buy.planMonthly': '专业版 · 月付 —— ¥19 / 月',
            'license.buy.planYearly': '专业版 · 年付 —— ¥199 / 年（含 1 个席位）',
            'license.buy.planTeam': '团队版（5人以上） —— ¥129 / 席位 / 年（几人买几席，每人一个独立激活码，总价 = 129 × 人数）',
            'license.buy.planHint': '有效期由激活码决定：月付 30 天、年付 365 天、团队版按年。激活后本机离线可用，不需要联网、不需要登录账户。',
            'license.buy.step1': '扫码添加微信（也可直接复制下方联系方式）',
            'license.buy.step2': '微信转账付款，备注里留下你的邮箱或微信号',
            'license.buy.step3': '把付款成功截图发给我，几分钟内回复激活码，回到上方粘贴激活即可',
            'license.buy.qr': '微信二维码（扫码添加）',
            'license.buy.qrMissing': '微信二维码未配置（放到 UI/images/wechat-qr.jpg），请用下方联系方式添加',
            'license.buy.contactLabel': '联系方式',
            'license.buy.copyContact': '复制联系方式',
            'license.buy.copied': '联系方式已复制',
            'license.buy.copyFailed': '复制失败，请手动选择复制',
            'license.buy.contactEmpty': '联系方式未配置，请通过文档中的渠道联系作者',
            'license.err.credentials': '请填写账号和密码',
            'license.err.auth': '登录或注册失败',
            'license.err.weakPassword': '密码至少 8 位',
            'license.err.serverDown': '暂时无法连接授权服务，请稍后重试。',
            'license.err.noBridge': '无法连接插件桥接，请从 Office 任务窗格内操作。',
            'license.passwordHint': '密码至少 8 位',
            'license.server.ok': '授权服务：已连接',
            'license.server.down': '授权服务：未连接',
            'license.server.offline': '授权服务：未连接（激活码离线可用）',
            'license.server.bridgeOnly': '授权服务：将通过插件连接（页面无法直连 localhost）',
            'license.err.checkout': '无法打开结账页',
            'license.err.portal': '无法打开账单门户',
            'license.ok.auth': '已登录',
            'license.ok.logout': '已退出',
            'license.ok.checkout': '已在浏览器打开结账页',
            'license.checkout.pending': '付费通道尚未开通。试用期内可以继续使用。不会进入模拟支付。',
            'license.trialScope': '注册试用为 14 天，期内可使用对话和云端识别。到期后这两项停止，直到订阅生效。付费尚未开通时不会出现升级按钮。未登录也可以先用本机识图。',
            'license.gate.needAccount': '对话和云端识别需要登录。打开侧栏「账户」注册试用。付费尚未开通时，本机识图可以先用。',
            'license.gate.blocked': '试用已结束或订阅未生效。请打开侧栏「账户 / 订阅」登录。付费通道开通前，过期账户不能继续使用需授权的功能。',
            'help.title': 'SmartPilot 帮助',
            'help.host': '当前宿主：{label}',
            'help.close': '关闭 (Esc)',
            'help.tabStart': '快速开始',
            'help.tabFaq': '常见问题',
            'help.tabShortcuts': '快捷键',
            'help.tabAbout': '关于',
            'help.about.version': '界面版本 {version}',
            'help.about.diag': '导出诊断信息',
            'help.about.diagHint': '只含版本、宿主、模型服务商和 API 主机名，不含密钥，也不含文档内容。',
            'help.disclosure.title': '发送到模型之前',
            'help.disclosure.accept': '我已了解，继续',
            'help.disclosure.cancel': '取消',
            'help.exportTitle': '导出当前会话',
            'help.exportHistoryTitle': '导出历史会话',
            'help.exportHint': '将对话下载到本地，便于备份或分享。',
            'help.exportMd': '导出 Markdown (.md)',
            'help.exportTxt': '导出纯文本 (.txt)',
            'help.start.1.t': '先识别一张图片',
            'help.start.1.d': '点功能区「识图入表」，把截图、表格、表单或票据写入当前表。清洗、出图、写文档时再配置模型。',
            'help.start.1.a': '打开识图入表',
            'help.start.2.t': '描述你的需求',
            'help.start.2.d': '在输入框用自然语言说明要做什么，或点击欢迎页 / 侧栏的快捷提示词（已按 {host} 适配）。',
            'help.start.2.a': '去对话',
            'help.start.3.t': '查看结果并继续',
            'help.start.3.d': '关注输入区上方的进度提示；导出成功可打开文件夹；失败可重试。需要时可导出本会话备份。',
            'help.start.3.a': '导出本会话',
            'help.sc.send': '发送消息',
            'help.sc.help': '显示帮助',
            'help.sc.list': '快捷键列表',
            'help.sc.esc': '关闭对话框 / 取消危险确认',
            'faq.common.0.q': '提示「请先配置 AI 模型」怎么办？',
            'faq.common.0.a': '点击侧栏 ⚙️ 打开设置，填写云端 API 密钥并选择模型；或启动本地 Ollama 后选择已检测到的模型，再点「保存」。',
            'faq.common.1.q': '导出的文件在哪里？',
            'faq.common.1.a': '导出成功后，工具卡片会显示完整路径，可点「打开文件夹」或「复制路径」。未指定路径时通常保存到桌面。',
            'faq.common.2.q': '工具失败了怎么办？',
            'faq.common.2.a': '失败卡片上可点「重试」或「改参数后重试」；缺参时会给出简易表单，补齐后再执行。',
            'faq.common.3.q': '界面看起来没更新？',
            'faq.common.3.a': 'WebView2 会缓存页面。请关闭并重新打开 SmartPilot 任务面板（不必重启 Office），以加载最新 UI 版本。',
            'faq.common.4.q': '危险操作确认是什么？',
            'faq.common.4.a': '清空、删除等可能改动文档的操作会先弹出确认。可按 Esc 取消；确认后本轮同类操作不再重复询问（除非开启「始终询问」）。',
            'faq.common.5.q': '我的表格会不会被上传？',
            'faq.common.5.a': '使用云端模型时，本次需要的文字和选区会发到你填写的 API 地址（可能在境外）。Ollama / llama.cpp 指向本机时，提示词留在本机。密钥不进安装包。完整说明见帮助「关于」或安装目录 Legal 文件夹。',
            'faq.excel.0.q': 'Excel 里从哪开始？',
            'faq.excel.0.a': '从功能区「识图入表」识别截图、表格、表单或票据。要改表时再选中区域，说清洗、拆分或出图。',
            'faq.excel.1.q': '识别结果如何填入表格？',
            'faq.excel.1.a': '图片、表格或票据识别完成后，点「填入 Excel」写入当前工作簿（仅 Excel 宿主）。身份证等证件只在确有业务需要时识别，结果留在本机工作簿。',
            'faq.word.0.q': 'Word 里能做什么？',
            'faq.word.0.a': '总结润色、排版、查找替换、导出 PDF 等。用自然语言描述，或点欢迎页的快捷提示词。',
            'faq.word.1.q': '为什么没有「插入到 Excel」？',
            'faq.word.1.a': '当前是 Word 宿主，表格结果可复制或下载；填入工作表仅在 Excel 中提供。',
            'faq.ppt.0.q': 'PPT 里能做什么？',
            'faq.ppt.0.a': '生成大纲、美化幻灯片、图表页、演讲备注、导出 PDF 等。可从欢迎页快捷提示词开始。',
            'faq.ppt.1.q': '导出 PDF 失败？',
            'faq.ppt.1.a': '确认演示文稿已保存过至少一次；失败卡片可重试。成功后用「打开文件夹」定位文件。',
            'toast.loading': '功能加载中，请稍候…',
            'toast.moduleLoading': '模块加载中，请稍候…',
            'undo.ok': '已撤销',
            'redo.ok': '已重做',
            'undo.fail': '撤销失败',
            'redo.fail': '重做失败',
            'undo.noHost': 'Excel 未连接',
            'err.network_error': '网络连接失败，请检查您的网络设置',
            'err.timeout': '请求超时，请稍后重试',
            'err.fetch_failed': '无法连接到服务器',
            'err.cors_error': '跨域请求被拒绝，请联系管理员',
            'err.excel_not_ready': '{host} 未就绪，请确保应用正常运行',
            'err.excel_access_denied': '无法访问 {host}，请检查权限',
            'err.excel_operation_failed': '{host} 操作失败',
            'err.cell_write_failed': '单元格写入失败',
            'err.ocr_failed': '文字识别失败，请尝试重新上传',
            'err.ocr_timeout': '文字识别超时',
            'err.ocr_engine_not_available': 'OCR 引擎不可用，请检查配置',
            'err.image_invalid': '图片格式无效，请上传 PNG/JPG 格式',
            'err.image_too_large': '图片过大，请压缩后重试',
            'err.llm_failed': 'AI 响应失败，请稍后重试',
            'err.llm_timeout': 'AI 响应超时',
            'err.llm_api_error': 'API 调用失败',
            'err.token_limit_exceeded': '上下文过长，请开启新对话',
            'err.config_load_failed': '配置加载失败',
            'err.config_save_failed': '配置保存失败',
            'err.config_invalid': '配置格式错误',
            'err.unknown_error': '发生未知错误',
            'err.permission_denied': '权限不足',
            'err.file_not_found': '文件不存在',
            'err.operation_cancelled': '操作已取消',
            'ext.title': '🛠️ 扩展中心',
            'ext.intro': '通过 Skills 创建自定义技能（VBA / AI 视觉 / 混合模式），通过 MCP 连接外部工具服务，扩展 AI 能力边界。',
            'ext.tip': '💡 提示：在对话中说「帮我创建一个 xxx 的 Skill」即可快速生成专属技能。',
            'ext.tabSkills': '🛠️ Skills',
            'ext.tabMcp': '🔌 MCP',
            'mcp.server': '🖥️ MCP Server',
            'mcp.notRunning': '未运行',
            'mcp.running': '运行中',
            'mcp.serverDesc': '将 Excel 操作为 MCP 工具供外部 AI 调用',
            'mcp.portTools': '端口',
            'mcp.toolsLabel': '工具',
            'mcp.start': '▶ 启动 Server',
            'mcp.stop': '⏹ 停止 Server',
            'mcp.clients': '🔗 MCP Clients',
            'mcp.add': '+ 添加',
            'mcp.clientsDesc': '连接外部 MCP Server 扩展 AI 能力',
            'mcp.noClients': '暂无已连接的 MCP Client',
            'mcp.availableTools': '🛠️ 可用工具',
            'mcp.noTools': '暂无可用工具',
            'mcp.reconnect': '🔗 重新连接',
            'mcp.refresh': '🔄 刷新',
            'kb.title': '📚 知识库 (RAG)',
            'kb.intro': '将企业文档（PDF / Word / Excel）导入知识库，AI 对话时自动检索相关内容作为参考。',
            'kb.flow': '💡 使用流程：新建知识库 → 导入文档 → 检索测试 → 开启对话增强',
            'kb.tabList': '📚 知识库列表',
            'kb.tabDocs': '📄 文档管理',
            'kb.tabSearch': '🔍 检索测试',
            'kb.tabSettings': '⚙️ RAG 设置',
            'kb.listTitle': '知识库列表',
            'kb.create': '+ 新建知识库',
            'kb.name': '名称',
            'kb.namePh': '例如：公司产品手册',
            'kb.desc': '描述（可选）',
            'kb.descPh': '简要描述这个知识库的用途...',
            'kb.cancel': '取消',
            'kb.createBtn': '创建',
            'kb.searchPh': '🔍 搜索知识库名称或描述...',
            'kb.filterAll': '全部状态',
            'kb.filterActive': '✓ 活跃',
            'kb.filterPending': '⏳ 待同步',
            'kb.filterArchived': '📦 已归档',
            'kb.empty': '暂无知识库',
            'kb.emptyHint': '点击右上角「+ 新建知识库」开始使用',
            'kb.docsTitle': '文档管理',
            'kb.docsHint': '粘贴本地文件路径或 OSS URL 导入文档：',
            'kb.pathPh': '本地路径如 C:\\docs\\file.pdf，或 OSS URL',
            'kb.browse': '浏览…',
            'kb.selectKb': '-- 请选择知识库 --',
            'kb.import': '导入文档',
            'kb.refreshDocs': '🔄 刷新',
            'kb.docsHint2': '支持格式：PDF / Word / Excel / CSV / TXT / MD。导入后自动解析分块并向量化，处理期间列表会自动刷新。',
            'kb.envChecking': '正在检测本地 RAG 环境…',
            'kb.envInstallTitle': '未检测到 Ollama',
            'kb.envInstallDesc': 'RAG 向量化依赖本地 Ollama 服务（免费开源，约 700MB）。请先下载安装，完成后点「重新检测」。Ollama 与模型不会随插件分发，需单独安装一次。',
            'kb.envDownload': '⬇ 下载 Ollama',
            'kb.envRecheck': '重新检测',
            'kb.envStartTitle': 'Ollama 服务未运行',
            'kb.envStartDesc': '检测到 Ollama 已安装但服务未启动。请打开 Ollama 应用（或在终端运行 ollama serve），然后点「重新检测」。',
            'kb.envPullTitle': '缺少向量模型 {model}',
            'kb.envPullDesc': 'RAG 需要嵌入模型 {model}（约 1.2GB，仅首次）。点「一键拉起」自动下载，或复制命令到终端执行。',
            'kb.envPullBtn': '⬇ 一键拉起 {model}',
            'kb.envPulling': '正在拉取…',
            'kb.envPullDone': '模型拉取完成',
            'kb.envPullFail': '模型拉取失败',
            'kb.envPullStartFail': '启动拉取失败',
            'kb.envPullCancelled': '已取消拉取',
            'kb.envCancel': '取消',
            'kb.envCopyCmd': '复制命令',
            'kb.envCopied': '已复制: {cmd}',
            'kb.envErrorTitle': '环境检测失败',
            'kb.envErrorDesc': '未知错误',
            'kb.envNoBridge': '当前宿主不支持打开外部链接',
            'kb.envOpenFail': '打开链接失败',
            'kb.retry': '重试',
            'kb.searchTitle': '检索测试台',
            'kb.queryPh': '输入测试问题...',
            'kb.runSearch': '执行检索',
            'kb.ragConfig': 'RAG 配置',
            'kb.embedding': '向量化服务 (Embedding)',
            'kb.provider': '默认 Provider:',
            'kb.ollamaEndpoint': '本地 Ollama Endpoint:',
            'kb.localModel': '本地模型名:',
            'kb.cloudEndpoint': '云端 Endpoint（Provider 选云端时使用）:',
            'kb.cloudEndpointPh': '例如：https://dashscope.aliyuncs.com/api/v1',
            'kb.cloudApiKey': '云端 API Key:',
            'kb.cloudModel': '云端模型名:',
            'kb.apiKeyKeepPh': '留空表示不修改',
            'kb.qdrant': '向量数据库 (Qdrant)',
            'kb.apiKeyOpt': 'API Key (可选):',
            'kb.apiKeyPh': '留空则不使用认证',
            'kb.chatEnhance': '对话增强',
            'kb.enableRag': '在对话中启用 RAG 增强',
            'kb.defaultKb': '默认知识库（对话时自动检索）:',
            'kb.defaultKbNone': '-- 不使用（关闭对话增强）--',
            'kb.topk': '默认检索数量 (TopK):',
            'kb.threshold': '相似度阈值 (0-1):',
            'kb.maxChunks': '注入上下文的最大块数:',
            'kb.saveCfg': '保存配置',
            'kb.testConn': '测试连接',
            'tp.title': '📋 任务规划',
            'tp.autoBypass': '自动化免确认（定时任务执行时不再弹确认框）',
            'tp.newTask': '+ 新建任务',
            'tp.refresh': '🔄 刷新',
            'tp.standalone': '🗂 打开独立任务列表',
            'tp.standaloneTitle': '打开独立任务列表（启动独立服务 + 打开独立页，Excel 关闭后仍可执行）',
            'tp.all': '全部',
            'tp.empty': '暂无定时任务',
            'tp.running': '运行中',
            'tp.paused': '已暂停',
            'tp.formTitle': '新建定时任务',
            'tp.taskName': '任务名称',
            'tp.taskNamePh': '如：每日数据清洗',
            'tp.freq': '执行频率',
            'tp.freqDaily': '每天',
            'tp.freqHourly': '每小时',
            'tp.freqWeekdays': '工作日每天',
            'tp.freqWeekly': '每周一',
            'tp.freqOnce': '仅一次',
            'tp.time': '执行时间',
            'tp.command': '执行指令',
            'tp.commandPh': '如：清洗当前工作表数据',
            'editor.title': '代码编辑器',
            'editor.lang': '切换编程语言',
            'editor.clear': '清空编辑器',
            'editor.format': '格式化代码',
            'editor.saveSnippet': '保存代码片段',
            'editor.library': '代码库',
            'editor.run': '执行代码',
            'editor.debug': '打开调试器窗口',
            'editor.placeholder': '在此粘贴或编辑 VBA 代码...',
            'tool.copyPath': '复制路径',
            'tool.openFolder': '打开文件夹',
            'tool.retry': '重试',
            'tool.editRetry': '改参数后重试',
            'tool.copiedPath': '已复制路径。请在资源管理器地址栏粘贴打开（重新编译插件后可一键打开文件夹）',
            'tool.openFolderFail': '无法打开文件夹，请手动定位：{path}',
            'toast.exported': '已导出 {n} 条消息',
            'toast.pathCopied': '路径已复制',
            'toast.copyFail': '复制失败',
            'toast.retryFilled': '已填入重试请求，确认后发送',
            'toast.retrying': '正在重试…',
            'toast.retryDone': '重试已完成',
            'toast.retryFail': '重试失败',
            'toast.retryNoTool': '无法重试：缺少工具名',
            'feat.excel.0.t': '识图入表', 'feat.excel.0.d': '截图、表格、表单和票据写入当前工作表',
            'feat.excel.1.t': '清洗拆分', 'feat.excel.1.d': '去重、合并、按列拆开',
            'feat.excel.2.t': '选区出图', 'feat.excel.2.d': '用当前表生成图表',
            'feat.excel.3.t': '生成汇报', 'feat.excel.3.d': '从这张表做出 PPT 页',
            'feat.word.0.t': '总结润色', 'feat.word.0.d': '全文总结、段落润色与翻译',
            'feat.word.1.t': '智能排版', 'feat.word.1.d': '目录、标题样式与表格',
            'feat.word.2.t': '查找替换', 'feat.word.2.d': '全文检索与批量替换',
            'feat.word.3.t': '校对检查', 'feat.word.3.d': '错别字与不通顺表述',
            'feat.word.4.t': '自定义 Skill', 'feat.word.4.d': '一句话创建专属文档技能',
            'feat.word.5.t': '导出 PDF', 'feat.word.5.d': '一键导出当前文档',
            'feat.ppt.0.t': '大纲生成', 'feat.ppt.0.d': '按主题生成多页大纲',
            'feat.ppt.1.t': '页面美化', 'feat.ppt.1.d': '排版配色与版式应用',
            'feat.ppt.2.t': '图表页', 'feat.ppt.2.d': '新增图表页展示数据',
            'feat.ppt.3.t': '演讲备注', 'feat.ppt.3.d': '为当前页生成演讲者备注',
            'feat.ppt.4.t': '自定义 Skill', 'feat.ppt.4.d': '一句话创建专属演示技能',
            'feat.ppt.5.t': '导出 PDF', 'feat.ppt.5.d': '一键导出演示文稿',
            'ppt.title': 'PPT 主题模板',
            'ppt.badge': '所见即所得',
            'ppt.dirty': '未保存更改',
            'ppt.hint': '改配色、拖版式，右侧大预览即时跟随。保存后对话里说「用 [主题名] 生成 PPT」即可套用。',
            'ppt.info': '主题信息',
            'ppt.name': '主题名称',
            'ppt.namePh': '如：公司红金',
            'ppt.id': '主题 ID',
            'ppt.idPh': '自动生成',
            'ppt.font': '正文字体',
            'ppt.presets': '一键配色',
            'ppt.colors': '精细调色',
            'ppt.contrastOk': '正文与背景对比度良好',
            'ppt.mine': '我的主题',
            'ppt.loading': '加载中…',
            'ppt.reset': '重置',
            'ppt.saveAs': '另存',
            'ppt.save': '保存',
            'ppt.live': '实时预览',
            'ppt.guides': '参考线',
            'ppt.editLayout': '编辑排版',
            'ppt.currentPage': '当前页：',
            'ppt.hintBrowse': '点缩略图切换页型 · 点标题可直接改示例文案',
            'ppt.hintLayout': '拖拽色块调整标题/正文/表格区 · 生成 PPT 时按此排版',
            'ppt.addDeco': '＋ 装饰色块',
            'ppt.resetLayout': '恢复默认布局',
            'ppt.canvasTip': '在大预览上拖拽色块 · 右下角缩放 · 设计坐标 960×720',
            'ppt.capBullets': '每页要点≤',
            'ppt.capRows': '表格行≤',
            'ppt.capCols': '表格列≤',
            'ppt.capSeries': '图表系列≤',
            'ppt.capCats': '图表分类≤',
            'ppt.del': '删除',
            'ppt.fill': '填充',
            'ppt.slide.title': '封面页',
            'ppt.slide.table': '数据表格',
            'ppt.slide.kpi': 'KPI 指标',
            'ppt.slide.agenda': '目录',
            'ppt.slide.chart': '图表',
            'ppt.slide.text': '要点',
            'ppt.slide.comparison': '对比',
            'ppt.slide.timeline': '时间线',
            'ppt.slide.section': '章节',
            'ppt.slide.closing': '结束页',
            'ppt.preset.business_blue': '商务蓝',
            'ppt.preset.tech_green': '科技绿',
            'ppt.preset.vibrant_orange': '活力橙',
            'ppt.preset.elegant_purple': '优雅紫',
            'ppt.preset.academic_navy': '学术深蓝',
            'ppt.preset.light_fresh': '浅色清新',
            'ppt.color.primary': '主色',
            'ppt.color.primary.tip': '标题、表头、章节底',
            'ppt.color.secondary': '辅色',
            'ppt.color.secondary.tip': '次要强调、柱状图',
            'ppt.color.accent': '强调色',
            'ppt.color.accent.tip': '高亮、装饰条',
            'ppt.color.bg': '背景色',
            'ppt.color.bg.tip': '整页底色',
            'ppt.color.light': '浅色块',
            'ppt.color.light.tip': '斑马纹、卡片底',
            'ppt.color.text': '正文色',
            'ppt.color.text.tip': '标题与正文',
            'ppt.color.gray': '辅助灰',
            'ppt.color.gray.tip': '副标题、说明',
            'ppt.contrast.good': '正文对比度 {r}:1 · 阅读清晰',
            'ppt.contrast.warn': '正文对比度 {r}:1 · 建议加深正文或提亮背景',
            'ppt.contrast.bad': '正文对比度偏低 {r}:1 · 投影时可能看不清',
            'ppt.toast.preset': '已套用「{name}」配色',
            'ppt.toast.resetLayout': '已恢复默认布局（960×720）',
            'ppt.toast.needName': '请先填写主题名称',
            'ppt.toast.loadBad': '载入失败：主题数据无效',
            'ppt.toast.loadEmpty': '载入失败：主题为空',
            'ppt.toast.loaded': '已载入「{name}」{layout}',
            'ppt.toast.loadedLayout': '（含版式）',
            'ppt.toast.notFound': '未找到主题：{id}',
            'ppt.toast.reset': '已重置为默认配色与版式',
            'ppt.toast.saved': '✓ 已保存{layout}：{name}',
            'ppt.toast.savedLayout': '（配色+版式）',
            'ppt.toast.saveFail': '保存失败：{err}',
            'ppt.toast.savedAs': '✓ 已另存为新主题',
            'ppt.toast.deleted': '已删除：{id}',
            'ppt.toast.deleteFail': '删除失败：{err}',
            'ppt.toast.unknown': '未知错误',
            'ppt.empty': '还没有自定义主题。调好配色/版式后点「保存」。',
            'ppt.loadFail': '加载失败：{err}',
            'ppt.bridgeFail': '无法连接宿主桥（请从任务面板 🎨 打开）',
            'ppt.layoutTag': ' · 版式',
            'ppt.loadTip': '点击载入',
            'ppt.confirmDelete': '确定删除主题「{id}」？',
            'pomo.title': '番茄钟',
            'pomo.titleStandalone': '番茄钟 · 独立模式',
            'pomo.connecting': '连接中…',
            'pomo.focus': '专注中',
            'pomo.break': '休息中',
            'pomo.session': '第 {n} / {max} 个番茄',
            'pomo.start': '开始',
            'pomo.pause': '暂停',
            'pomo.resume': '继续',
            'pomo.reset': '重置',
            'pomo.statTotal': '累计番茄',
            'pomo.statMinutes': '专注分钟',
            'pomo.statToday': '今日完成',
            'pomo.footer': '独立模式 · 连接 Excel 时双向同步，离线自动保存本地',
            'pomo.settings': '番茄钟设置',
            'pomo.workDur': '专注时长（分钟）',
            'pomo.shortDur': '短休息时长（分钟）',
            'pomo.longDur': '长休息时长（分钟）',
            'pomo.longEvery': '长休息间隔（每 N 个番茄后）',
            'pomo.sound': '完成时播放提示音',
            'pomo.notif': '桌面通知（首次需授权）',
            'pomo.cancel': '取消',
            'pomo.save': '保存',
            'pomo.min': '{n} 分',
            'pomo.count': '{n} 个',
            'pomo.unitMin': '分',
            'pomo.unitCount': '个',
            'pomo.syncTip': '同步状态 / 立即同步一次',
            'pomo.gearTip': '设置',
            'pomo.synced': '已同步',
            'pomo.offline': '离线模式',
            'pomo.syncing': '同步中…',
            'pomo.offlineBanner': '未连接 Excel 服务，当前为本地离线模式',
            'pomo.toast.running': '番茄钟已在运行',
            'pomo.toast.resetStart': '已重置，请再次点击开始',
            'pomo.toast.focusStart': '▶ 专注开始，加油！',
            'pomo.toast.notRunning': '当前并未在运行',
            'pomo.toast.paused': '⏸ 已暂停',
            'pomo.toast.resetTo': '↺ 已重置为 {time}',
            'pomo.toast.saved': '设置已保存',
            'pomo.toast.synced': '已与 Excel 同步',
            'pomo.toast.offlineMode': '未连接 Excel，已切换离线模式',
            'pomo.notify.focusTitle': '番茄钟 · 专注完成',
            'pomo.notify.focusBody': '休息一下，做做伸展！',
            'pomo.notify.breakTitle': '番茄钟 · 休息结束',
            'pomo.notify.breakBody': '开始新一轮专注！',

            // ── 语言切换 / 快捷入口 ──
            'nav.language': '切换界面语言',
            'nav.languageSwitched': '已切换为 {lang}',

            // ── 任务规划 ──
            'tp.runNow': '立即执行',
            'tp.model': '执行模型（选填；不选=使用上方「默认执行模型」；和聊天栏选法完全一致）',
            'tp.modelAuto': 'AUTO（跟随默认执行模型）',
            'tp.modelRefreshTitle': '重新检测 Ollama 本地模型 / 刷新云端模型列表',
            'tp.refreshBtn': '🔄 刷新',
            'tp.modelHint': '选填：给这个任务单独指定执行模型，比如本地 Ollama 跑 qwen2.5、推理跑 deepseek-r1',
            'tp.targetFile': '目标 Excel 文件（独立执行时使用，可选）',
            'tp.targetFilePh': '如：C:\\Users\\...\\销售报表.xlsx   Excel 关闭后由独立服务直接操作此文件',
            'tp.browse': '📁 浏览',
            'tp.cancel': '取消',
            'tp.save': '保存',

            // ── 代码片段 ──
            'snippet.libraryTitle': '代码库',
            'snippet.close': '关闭',
            'snippet.saveTitle': '保存代码片段',
            'snippet.name': '名称',
            'snippet.namePh': '例如：批量处理数据',
            'snippet.desc': '描述（可选）',
            'snippet.descPh': '简要描述这段代码的功能',
            'snippet.code': '代码',
            'snippet.codePh': 'VBA 代码...',
            'snippet.saveAsSkill': '同时保存为 Skill（推荐，可配置关键词匹配）<br><small style="color: #666;">保存为 Skill 后，可以通过自然语言描述直接调用这段代码</small>',
            'snippet.cancel': '取消',
            'snippet.save': '保存',

            // ── 知识库 ──
            'kb.searchKbPh': '🔍 搜索知识库名称或描述...',
            'kb.embLocalBge': '本地 BGE (Ollama)',
            'kb.embQianwen': '通义千问',
            'kb.embZhipu': '智谱 AI',

            // ── Skill 编辑 ──
            'skill.addTitle': '添加 Skill',
            'skill.viewSource': '📄 查看MD源文件',
            'skill.tabForm': '表单编辑',
            'skill.tabJson': 'JSON源码',
            'skill.basicInfo': '基础信息',
            'skill.name': 'Skill 名称（英文标识）',
            'skill.namePh': '例如：MyCustomSkill',
            'skill.displayName': '显示名称',
            'skill.displayPh': '例如：我的自定义操作',
            'skill.desc': '描述',
            'skill.descPh': '描述这个 Skill 的功能',
            'skill.keywords': '关键词（用逗号分隔）',
            'skill.keywordsPh': '例如：自定义,操作,处理',
            'skill.processing': '处理配置',
            'skill.procMode': '处理模式',
            'skill.procVba': 'VBA代码 - 传统Excel 宏',
            'skill.procLlm': '多模态LLM - AI视觉模型',
            'skill.procHybrid': '混合模式 - 自动选择',
            'skill.procOcr': '本地OCR - PaddleOCR引擎',
            'skill.codeTpl': '代码模板',
            'skill.codePh': '在这里编写代码\nVBA 示例：\nSub MyMacro()\n  \' 你的逻辑\nEnd Sub',
            'skill.examples': '示例输入（用逗号分隔）',
            'skill.examplesPh': '例如：执行我的操作，运行自定义宏',
            'skill.specParams': 'Spec 参数',
            'skill.advanced': '高级选项',
            'skill.llmSysPrompt': 'LLM 系统提示',
            'skill.llmSysPh': '你是一个专业的AI助手。请根据用户输入执行相应任务...\n\n【规则】\n1. ...\n2. ...',
            'skill.llmSysHint': '定义 AI 的角色和行为准则（多模态 LLM 模式使用）',
            'skill.llmUserTpl': 'LLM 用户提示词模板',
            'skill.llmUserPh': '请处理以下内容：{input}',
            'skill.llmUserHint': '使用 {input} 作为用户输入的占位符',
            'skill.outputFormat': '输出格式',
            'skill.outText': '纯文本 (text)',
            'skill.outJson': 'JSON 格式',
            'skill.outVba': 'VBA 代码',
            'skill.outHtmlBrowser': 'HTML 浏览器展示',
            'skill.outHtml': 'HTML 输出',
            'skill.outTable': '表格',
            'skill.outChart': '图表',
            'skill.outPivot': '数据透视表',
            'skill.version': '版本',
            'skill.tags': '标签（用逗号分隔）',
            'skill.tagsPh': '例如：OCR,图像处理,AI',
            'skill.supportsImage': '支持图片输入',
            'skill.supportsImageTitle': '启用后，用户可以在对话中上传图片，AI将识别并处理图片内容（适用于OCR、图像分析等场景）',
            'skill.supportsFile': '支持文件输入',
            'skill.supportsFileTitle': '启用后，用户可以上传文件（如Excel、PDF等），AI将读取并分析文件内容',
            'skill.quickShortcut': '⭐ 加入快捷入口',
            'skill.quickShortcutTitle': '标记后，该Skill将显示在欢迎面板的快捷入口区域，方便用户快速调用常用功能',
            'skill.quickConfig': '快捷入口设置',
            'skill.quickIcon': '图标（emoji）',
            'skill.quickPrompt': '点击提示文本',
            'skill.quickPromptPh': '留空则使用第一个示例',
            'skill.quickHint': '设置后此 Skill 将出现在欢迎面板的「快捷提示词」区域，点击即可快速触发',
            'skill.uiButtons': 'UI操作按钮配置',
            'skill.uiCopy': '📋 复制',
            'skill.uiSaveMd': '💾 保存MD',
            'skill.uiRegenerate': '🔄 重新生成',
            'skill.uiHint': '控制AI消息底部显示的按钮（仅Skill配置类建议开启"保存MD"）',
            'skill.jsonLabel': 'Skill JSON 配置',
            'skill.formatJson': '格式化 JSON',
            'skill.validateJson': '验证 JSON',
            'skill.cancel': '取消',
            'skill.restoreDefault': '恢复默认',
            'skill.save': '保存',

            // ── MCP Server 编辑 ──
            'mcpsrv.addTitle': '添加 MCP Server',
            'mcpsrv.name': 'Server 名称',
            'mcpsrv.namePh': '例如：filesystem, sqlite',
            'mcpsrv.desc': '描述',
            'mcpsrv.descPh': '描述此 MCP Server 的功能',
            'mcpsrv.transport': '传输类型',
            'mcpsrv.stdio': 'stdio（本地进程通信）',
            'mcpsrv.http': 'HTTP SSE（远程服务器）',
            'mcpsrv.command': '命令',
            'mcpsrv.commandPh': '例如：npx, uvx, python',
            'mcpsrv.args': '参数（每行一个）',
            'mcpsrv.env': '环境变量（每行一个，格式：KEY=VALUE）',
            'mcpsrv.headers': 'Headers（每行一个，格式：Key=Value）',
            'mcpsrv.envHttp': '环境变量（每行一个，格式：KEY=VALUE）',
            'mcpsrv.enabled': '启用',
            'mcpsrv.cancel': '取消',
            'mcpsrv.test': '🧪 测试连接',
            'mcpsrv.save': '保存并连接',

            // ── 快捷提示词 ──
            'quick.barHint': '在 Skill 管理中标记「⭐快捷入口」即可在此显示',
            'quick.title': '⚡ 快捷提示词',
            'quick.manageTitle': '⚙️ 管理快捷提示词',
            'quick.manage': '⚙️ 管理',
            'quick.searchPh': '搜索提示词',
            'quick.addTitle': '新增快捷提示词',
            'quick.add': '＋ 新增',
            'quick.resetTitle': '恢复默认配置',
            'quick.done': '完成',

            // ── 输入区 / 模型选择 ──
            'model.refreshTitle': '刷新模型列表',
            'think.selectTitle': '选择思考模式',
            'think.auto': 'AUTO（智能自动选择）',
            'think.quick': '⚡快速模式',
            'think.standard': '📝标准模式',
            'think.deep': '🔍深度模式',
            'think.expert': '👑专家模式',
            'input.analyzing': 'AI正在分析...',
            'input.pickRange': '请在 Excel 中选择区域，选择完成后自动添加...',
            'input.quickCmd': '快捷指令 - 输入 / 选择功能'
        },
        en: {
            'app.title': 'SmartPilot',
            'tab.chat': '💬 Chat',
            'tab.history': '📜 History',
            'tab.skills': '🛠️ Extensions',
            'tab.knowledge': '📚 Knowledge',
            'nav.chat': 'Chat',
            'nav.history': 'History',
            'nav.ocr': 'Capture',
            'nav.more': 'More',
            'nav.settingsShort': 'Settings',
            'nav.helpShort': 'Help',
            'nav.accountShort': 'Account',
            'nav.ocrModel': 'OCR model',
            'nav.mobile': 'Mobile',
            'nav.skills': 'Extensions',
            'nav.knowledge': 'Knowledge',
            'nav.newChat': 'New chat',
            'nav.taskPlanner': 'Task planner',
            'nav.editor': 'Code editor',
            'nav.quickPrompts': 'Quick prompts',
            'nav.pomodoro': 'Pomodoro (synced)',
            'nav.pomodoroShort': 'Pomodoro',
            'nav.pptTheme': 'PPT theme templates',
            'nav.export': 'Export this chat',
            'nav.help': 'Help (F1)',
            'nav.clearCache': 'Clear cache',
            'nav.settings': 'AI model settings',
            'nav.account': 'Account / Subscription',
            'nav.collapseToolbar': 'Collapse / expand toolbar',
            'nav.scrollUp': 'Scroll up',
            'nav.scrollDown': 'Scroll down',
            'nav.undo': 'Undo the add-in step (not the worksheet Ctrl+Z)',
            'nav.redo': 'Redo the add-in step (not the worksheet Ctrl+Y)',
            'status.detecting': 'Checking…',
            'status.ready': 'Ready',
            'conn.title': 'Excel connection',
            'conn.checking': 'Checking…',
            'conn.connected': 'Connected',
            'conn.connecting': 'Connecting…',
            'conn.disconnected': 'Disconnected',
            'conn.excelClosed': 'Excel closed',
            'conn.excelNotOpen': 'Excel not open',
            'conn.serviceDown': 'Service not running',
            'conn.cannotConnect': 'Cannot connect to Excel',
            'conn.linkLost': 'Connection lost',
            'conn.mode': 'Mode:',
            'conn.time': 'Time:',
            'conn.justNow': 'Just now',
            'conn.minsAgo': '{n} min ago',
            'conn.hoursAgo': '{n} h ago',
            'conn.modeVsto': 'VSTO embed (this Excel)',
            'conn.modeBridge': 'Local bridge (pywin32)',
            'conn.pickWorkbook': 'Select workbook',
            'conn.pickSheet': 'Select worksheet',
            'conn.workbook': 'Workbook:',
            'conn.activeSheet': 'Active sheet:',
            'conn.sheetCount': 'Sheets:',
            'conn.sheetUnit': '{n}',
            'conn.activeWb': '(active workbook)',
            'conn.loading': 'Loading…',
            'conn.noWorkbooks': 'No open workbooks — open a file in Excel first',
            'conn.connect': 'Connect to Excel',
            'conn.refreshList': 'Refresh workbook list',
            'conn.disconnect': 'Disconnect',
            'conn.refresh': 'Refresh status',
            'conn.confirmDisconnect': 'Disconnect?',
            'conn.hostHint': 'Embedded in Excel — switching workbooks affects this window',
            'conn.refreshed': 'Refreshed',
            'conn.switchedSheet': 'Switched sheet: {name}',
            'conn.switchedWb': 'Switched to: {name}',
            'conn.connectOk': 'Connected to Excel',
            'conn.needExcel': 'Please open Excel first',
            'conn.needConnect': 'Connect to Excel first',
            'conn.loadWbFail': 'Failed to load workbooks: {msg}',
            'conn.switchSheetFail': 'Failed to switch sheet: {msg}',
            'conn.switchWbFail': 'Switch failed: {msg}',
            'conn.pickWbWarn': 'Select a workbook',
            'conn.excelClosedToast': 'Excel closed — reconnect',
            'clear.title': 'Choose data to clear',
            'clear.history': 'Chat history',
            'clear.historyDesc': 'All chat messages and session history',
            'clear.config': 'App settings',
            'clear.configDesc': 'API keys, model selection, and related config',
            'clear.types': 'Custom recognition types',
            'clear.typesDesc': 'User-defined recognition types',
            'clear.snippets': 'Code snippets',
            'clear.snippetsDesc': 'Saved snippets and code library',
            'clear.cancel': 'Cancel',
            'clear.confirm': 'Clear selected',
            'tools.localExcel': 'Local Excel',
            'tools.localPpt': 'Local PPT',
            'tools.localWord': 'Local Word',
            'tools.localPrefix': 'Local {name}',
            'tools.real': 'Live',
            'tools.mock': 'Mock',
            'tools.expand': 'Expand',
            'tools.call': 'Run',
            'tools.loading': 'Loading tools…',
            'ocr.title': 'Capture to sheet',
            'ocr.more': 'Batch / Parse / History',
            'ocr.localHint': 'Local recognition works first. Cloud OCR and chat need an account. Until checkout is open, local capture does not require sign-in.',
            'ocr.tab.recognize': 'Recognize',
            'ocr.tab.batch': 'Batch',
            'ocr.tab.parser': 'Smart parse',
            'ocr.tab.history': 'History',
            'ocr.mobile': 'Mobile',
            'ocr.engine': 'Engine',
            'ocr.lang': 'Language',
            'ocr.lang.auto': 'Auto-detect',
            'ocr.lang.zh': 'Simplified Chinese',
            'ocr.lang.zht': 'Traditional Chinese',
            'ocr.lang.en': 'English',
            'ocr.lang.ja': 'Japanese',
            'ocr.lang.ko': 'Korean',
            'ocr.lang.zhOnly': 'This engine recognizes Chinese',
            'ocr.recognizeThis': 'Recognize this image',
            'ocr.drop': 'Drop images or PDFs here',
            'ocr.dropHint': 'Single or multiple files',
            'ocr.pickFile': 'Choose files',
            'ocr.pickFolder': 'Choose folder',
            'ocr.screenshot': 'Screenshot OCR',
            'ocr.preview': 'Image preview',
            'ocr.result': 'Results',
            'ocr.view.text': 'Text',
            'ocr.view.list': 'List',
            'ocr.view.preview': 'Preview',
            'ocr.format': 'Format',
            'ocr.format.plain': 'Plain text',
            'ocr.format.raw': 'OCR original',
            'ocr.format.excel': 'Excel table',
            'ocr.format.preview': 'Preview',
            'ocr.start': 'Start OCR',
            'ocr.insert': 'Write to this sheet',
            'ocr.saveExcel': 'Save as Excel',
            'ocr.savedExcel': 'Saved Excel ({n} rows)',
            'ocr.noExportData': 'Nothing to export',
            'ocr.insertFields': 'Write by field',
            'ocr.export': 'Export',
            'ocr.ready': 'Ready',
            'ocr.settingsTitle': 'OCR settings',
            'ocr.settingsEngine': 'OCR engine',
            'ocr.engineStatus': 'Engine status',
            'ocr.detecting': 'Checking…',
            'ocr.paddleLang': 'Recognition language',
            'ocr.layout': 'Layout analysis',
            'ocr.table': 'Table recognition',
            'ocr.formula': 'Formula recognition',
            'ocr.verifyPaddle': 'Verify PaddleOCR install',
            'ocr.saveSettings': 'Save OCR settings',
            'ocr.opt.rapid': 'RapidOCR — embedded, lightweight (~30MB)',
            'ocr.opt.paddle': 'PaddleOCR — best Chinese accuracy (needs Python)',
            'ocr.opt.cloud': 'Cloud OCR API — highest accuracy',
            'ocr.lang.ch': 'Chinese (default)',
            'ocr.modelSettings': 'Model settings',
            'ocr.opt.rapidShort': 'RapidOCR (embedded)',
            'ocr.opt.paddleShort': 'PaddleOCR (best for Chinese)',
            'ocr.opt.cloudShort': 'Cloud OCR',
            'ocr.recognizeAll': 'Recognize all',
            'ocr.mechanical': 'Mechanical drawing',
            'ocr.mechanicalTitle': 'Add geometric tolerance symbols on mechanical drawings only. When off, RapidOCR and PaddleOCR keep their own text.',
            'ocr.packs.manage': 'Manage domain packs',
            'ocr.pickScript': 'Choose a parse script...',
            'ocr.batchEmpty': 'Choose images or PDF files',
            'ocr.batchEmptyHint': 'PNG, JPG, PDF, and similar formats',
            'ocr.batchPickWait': 'OCR engine is still starting',
            'ocr.batchStart': 'Start batch',
            'ocr.batchPause': 'Pause',
            'ocr.batchCancel': 'Cancel',
            'ocr.batchInsert': 'Insert into Excel',
            'ocr.batchParse': 'Batch parse',
            'ocr.removeImage': 'Remove image',
            'ocr.parser.title': 'Parse scripts',
            'ocr.parser.clear': 'Clear cache',
            'ocr.parser.upload': 'Upload',
            'ocr.parser.uploadTitle': 'Upload a Markdown file',
            'ocr.parser.new': 'New',
            'ocr.parser.searchPh': 'Search scripts...',
            'ocr.parser.allCategories': 'All categories',
            'ocr.parser.loading': 'Loading',
            'ocr.parser.edit': 'Edit script',
            'ocr.parser.preview': 'Test preview',
            'ocr.parser.fieldsTab': 'Fields',
            'ocr.parser.name': 'Script name *',
            'ocr.parser.namePh': 'Example: VAT invoice',
            'ocr.parser.desc': 'Description',
            'ocr.parser.descPh': 'What this script extracts',
            'ocr.parser.icon': 'Icon',
            'ocr.parser.category': 'Category',
            'ocr.parser.cat.general': 'General',
            'ocr.parser.cat.receipt': 'Receipts',
            'ocr.parser.cat.id': 'IDs',
            'ocr.parser.cat.business': 'Business',
            'ocr.parser.cat.custom': 'Custom',
            'ocr.parser.output': 'Output format',
            'ocr.parser.out.table': 'Table (two columns)',
            'ocr.parser.out.kv': 'Key-value',
            'ocr.parser.code': 'Script code',
            'ocr.parser.save': 'Save script',
            'ocr.parser.reset': 'Reset default',
            'ocr.parser.test': 'Test run',
            'ocr.parser.testInput': 'Test text (OCR result)',
            'ocr.parser.testPh': 'Paste OCR text here, or load the sample...',
            'ocr.parser.loadSample': 'Load sample text',
            'ocr.parser.run': 'Run parse',
            'ocr.parser.resultEmpty': 'The parse result will appear here',
            'ocr.parser.copy': 'Copy',
            'ocr.parser.insert': 'Insert into Excel',
            'ocr.parser.fields': 'Fields',
            'ocr.parser.addField': 'Add field',
            'ocr.parser.noFields': 'No fields yet. Use the button above to add one.',
            'ocr.history.title': 'History',
            'ocr.history.clear': 'Clear all',
            'ocr.history.empty': 'No history yet',
            'ocr.history.emptyHint': 'Recognized images are saved here',
            'ocr.desc.rapid': 'Lightweight ONNX OCR. Small model (~30MB), fast startup, shipped with the add-in.',
            'ocr.desc.paddle': 'Open-source OCR with the best Chinese accuracy. Requires Python.',
            'ocr.desc.cloud': 'Cloud OCR. Highest accuracy. Needs an API key and a network connection.',
            'ocr.rapid.intro': 'RapidOCR is a lightweight ONNX engine. The model is about 30MB, it starts quickly, and it ships with the add-in.',
            'ocr.rapid.checking': 'Checking RapidOCR...',
            'ocr.rapid.features': 'RapidOCR',
            'ocr.rapid.f1': 'Runs on ONNX Runtime, no Python required',
            'ocr.rapid.f2': 'Small model (~30MB), quick to install',
            'ocr.rapid.f3': 'Chinese, English, and other languages',
            'ocr.rapid.f4': 'Fast startup, suited to batches',
            'ocr.rapid.check': 'Check RapidOCR',
            'ocr.paddle.intro': 'PaddleOCR is an open-source engine with the best results on Chinese text.',
            'ocr.paddle.checking': 'Checking PaddleOCR...',
            'ocr.paddle.advanced': 'PaddleOCR extras (PP-OCRv6 + PP-StructureV3)',
            'ocr.paddle.hint': 'PP-OCRv6 recognizes many languages in one model. Layout analysis finds document structure, table recognition keeps rows and columns, and formula recognition returns LaTeX.',
            'ocr.paddle.noPython': 'Python was not found. PaddleOCR needs Python 3.9 or newer.',
            'ocr.paddle.downloadPython': 'Download Python',
            'ocr.paddle.install': 'Install PaddleOCR',
            'ocr.paddle.diagnoseLabel': 'Diagnosis:',
            'ocr.paddle.log': 'Install log:',
            'ocr.paddle.manualTitle': 'Manual install (advanced)',
            'ocr.paddle.manualHint': 'If one-click install fails, run these commands yourself:',
            'ocr.win.intro': 'Uses the OCR built into Windows 10/11. No extra software.',
            'ocr.win.checking': 'Checking Windows OCR...',
            'ocr.win.langs': 'Supported languages',
            'ocr.cloud.intro': 'Cloud OCR is the most accurate. It needs a network connection and an API key.',
            'ocr.cloud.provider': 'Cloud provider',
            'ocr.cloud.baidu': 'Baidu Cloud OCR',
            'ocr.cloud.tencent': 'Tencent Cloud OCR',
            'ocr.cloud.baiduKeyPh': 'Baidu API Key',
            'ocr.cloud.baiduSecretPh': 'Baidu Secret Key',
            'ocr.cloud.baiduLink': 'Get a Baidu API key',
            'ocr.cloud.tencentIdPh': 'Tencent SecretId',
            'ocr.cloud.tencentKeyPh': 'Tencent SecretKey',
            'ocr.cloud.tencentLink': 'Get a Tencent API key',
            'ocr.cloud.test': 'Connection test',
            'ocr.cloud.testBtn': 'Test connection',
            'ocr.status.checking': 'Checking...',
            'ocr.status.waitInit': 'Waiting to start...',
            'ocr.status.available': 'Available',
            'ocr.status.unavailable': 'Unavailable',
            'ocr.status.ready': 'Ready',
            'ocr.status.needConfig': 'Needs setup',
            'ocr.status.rapidOk': 'RapidOCR is installed',
            'ocr.status.rapidBad': 'RapidOCR is unavailable',
            'ocr.status.checkPlugin': 'Check the add-in install',
            'ocr.status.fail': 'Check failed',
            'ocr.status.checkBackend': 'Check the local service',
            'ocr.status.notSelected': 'Not selected',
            'ocr.status.pickEngine': 'Choose an OCR engine above',
            'ocr.langId.ch': 'Chinese',
            'ocr.langId.en': 'English',
            'ocr.langId.japan': 'Japanese',
            'ocr.langId.korean': 'Korean',
            'ocr.langId.fr': 'French',
            'ocr.langId.ger': 'German',
            'ocr.langId.es': 'Spanish',
            'ocr.langId.ru': 'Russian',
            'ocr.langId.arabic': 'Arabic',
            'ocr.langId.structure': 'Auto-detect',
            'ocr.fileHint': 'A local file cannot reach the OCR service. Opening it switches to the local page.',
            'ocr.openOcr': 'Open OCR',
            'ocr.desc.default': 'Choose an OCR engine. It supports many languages and runs locally.',
            'ocr.formulaInstall': 'Install formula pack',
            'ocr.formulaReady': 'Formula recognition is ready',
            'ocr.formulaInstalling': 'Installing (about 3–5 minutes)...',
            'ocr.pdfPrev': 'Previous page',
            'ocr.pdfNext': 'Next page',
            'ocr.pdfPause': 'Pause',
            'ocr.pdfResume': 'Resume',
            'ocr.pdfCancel': 'Cancel',
            'ocr.resultPh': 'Recognition results appear here...',
            'ocr.formatHint': 'Output format for Excel and copy',
            'ocr.copyTitle': 'Copy',
            'ocr.clearTitle': 'Clear',
            'ocr.remove': 'Remove',
            'ocr.export.excel': 'Export as Excel table',
            'ocr.export.csv': 'Export as CSV',
            'ocr.export.text': 'Export as text file',
            'ocr.batch.track': 'Follow',
            'ocr.batch.trackTitle': 'Scroll to each result while processing',
            'ocr.batch.total': 'Total',
            'ocr.batch.done': 'Done',
            'ocr.batch.failed': 'Failed',
            'ocr.batch.clear': 'Clear',
            'ocr.batch.summary': 'Results',
            'ocr.batch.copyAll': 'Copy all',
            'ocr.batch.emptyFiles': 'Choose images or a folder',
            'ocr.batch.addFile': 'Choose files',
            'ocr.batch.addFolder': 'Choose folder',
            'ocr.batch.remove': 'Remove',
            'ocr.batch.pending': 'Waiting',
            'ocr.batch.processing': 'Processing',
            'ocr.batch.completed': 'Done',
            'ocr.batch.failedMsg': 'Failed: {msg}',
            'ocr.batch.unknownError': 'Unknown error',
            'ocr.batch.engineWait': 'The OCR engine is still starting',
            'ocr.parser.pickOrNew': 'Choose or create a script',
            'ocr.parser.builtin': 'Built-in',
            'ocr.parser.unnamed': 'Untitled script',
            'ocr.parser.parsing': 'Parsing...',
            'ocr.parser.noData': 'No fields were extracted',
            'ocr.parser.parseFail': 'Parse failed: {msg}',
            'ocr.parser.noMatch': 'No matching scripts',
            'ocr.parser.none': 'No scripts yet. Use New above.',
            'ocr.recognizeAllCount': 'Recognize all ({n})',
            'ocr.status.engineReady': 'Engine ready',
            'ocr.status.engineNotReady': 'OCR engine is not ready: {msg}',
            'ocr.status.hostDown': 'OCR service is not running',
            'ocr.status.unknownEngine': 'Unknown engine',
            'ocr.status.noPython': 'Python was not found',
            'ocr.status.rapidNeedPy': 'RapidOCR is unavailable. Check the Python environment.',
            'ocr.status.paddleNoModel': 'PaddleOCR is installed, but the models are not downloaded. Install again to download them.',
            'ocr.status.cloudName': 'Cloud OCR',
            'ocr.status.cloudReady': 'Cloud OCR is configured',
            'ocr.status.cloudNeedKey': 'Cloud OCR is not configured. Add an API key in settings.',
            'ocr.paddle.installed': 'PaddleOCR is installed ({name})',
            'ocr.paddle.missingCount': 'PaddleOCR is not installed ({n} Python installs found, none usable)',
            'ocr.paddle.ok': 'PaddleOCR is installed',
            'ocr.paddle.notInstalled': 'PaddleOCR is not installed',
            'ocr.paddle.cannotCheck': 'Status could not be checked',
            'ocr.status.installing': 'Installing...',
            'ocr.status.startingInstall': 'Starting install...',
            'ocr.noResult': 'No recognition result',
            'ocr.noText': '(No text recognized)',
            'ocr.fit': 'Fit',
            'ocr.mode.label': 'Mode',
            'ocr.mode.auto': 'Auto',
            'ocr.mode.text': 'Plain text',
            'ocr.mode.formula': 'Formula',
            'ocr.mode.number': 'Number',
            'ocr.mode.date': 'Date',
            'ocr.mode.mechanical': 'Mechanical drawing',
            'ocr.adjusted': 'Adjusted',
            'ocr.pageLabel': 'Page {n}',
            'ocr.shotTitle': 'Shot {n}',
            'ocr.tool.expand': 'Expand toolbar',
            'ocr.tool.rotateLeft': 'Rotate 90° left',
            'ocr.tool.rotateRight': 'Rotate 90° right',
            'ocr.tool.flip': 'Flip horizontal',
            'ocr.tool.zoomOut': 'Zoom out',
            'ocr.tool.zoomIn': 'Zoom in',
            'ocr.tool.resetZoom': 'Reset zoom',
            'ocr.tool.viewOriginal': 'View original',
            'ocr.tool.collapse': 'Collapse toolbar',
            'ocr.manual.step1': '# 1. Open Command Prompt or PowerShell',
            'ocr.manual.step2': '# 2. Run these commands (use your Python version):',
            'ocr.manual.step3': '# 3. Download the OCR models (required):',
            'ocr.manual.done': 'Models downloaded',
            'ocr.manual.tip': 'If the Tsinghua mirror fails, try another index<br>Alibaba https://mirrors.aliyun.com/pypi/simple/<br>USTC https://pypi.mirrors.ustc.edu.cn/simple/',
            'ocr.detailTitle': 'Recognition details',
            'ocr.reRecognize': 'Recognize again',
            'ocr.parseShort': 'Parse',
            'ocr.selectAll': 'Select all',
            'ocr.addRegion': 'Add region',
            'ocr.realtime': 'Live recognition',
            'ocr.recognizeSelected': 'Recognize selected',
            'ocr.recognizeSelectedTitle': 'Recognize the checked regions',
            'ocr.duplicateScript': 'Duplicate script',
            'ocr.deleteScript': 'Delete script',
            'ocr.ctx.clearLatex': 'Clear LaTeX',
            'ocr.ctx.fixNumbers': 'Fix number format',
            'ocr.ctx.fixDates': 'Fix date format',
            'ocr.ctx.removeSpaces': 'Remove extra spaces',
            'ocr.ctx.edit': 'Edit',
            'ocr.ctx.deleteRegion': 'Delete region',
            'ocr.recognizing': 'Recognizing...',
            'ocr.recognizingNow': 'Recognizing...',
            'ocr.loadingEngine': 'Loading the OCR engine...',
            'ocr.analyzingImage': 'Analyzing the image...',
            'ocr.readingText': 'Reading text (the first run loads the model)...',
            'ocr.readingTextWait': 'Reading text (model still loading, {n}s)...',
            'ocr.shotAlt': 'Screenshot {n}',
            'ocr.shotProgress': 'Shot {n} of {total}',
            'ocr.shotSwitch': 'Shot {n} of {total}. Click a thumbnail to switch. Press Enter to recognize all.',
            'ocr.parser.signature': 'Function',
            'ocr.parser.fieldsHint': 'Fields come from',
            'ocr.batch.scanning': 'Scanning folder...',
            'ocr.batch.scanDone': 'Scan finished, {n} files',
            'ocr.batch.scanFail': 'Scan failed: {msg}',
            'ocr.batch.noBackend': 'Cannot reach the local service',
            'ocr.batch.running': 'Recognizing {n}/{total}...',
            'ocr.batch.fileProgress': 'Recognizing ({n}/{total}): {name}',
            'ocr.batch.doneSummary': 'Batch finished: {ok} succeeded, {bad} failed',
            'ocr.batch.parsing': 'Parsing ({n}/{total}): {name}',
            'ocr.batch.parseDone': 'Parse finished: {ok} succeeded, {bad} failed',
            'ocr.batch.paused': 'Paused',
            'ocr.batch.resuming': 'Resuming...',
            'ocr.batch.cancelled': 'Cancelled',
            'ocr.batch.cancelledToast': 'Batch cancelled',
            'ocr.batch.noFiles': 'No files to parse',
            'mobile.title': 'Mobile remote OCR',
            'mobile.status': 'Service status',
            'mobile.stopped': 'Stopped',
            'mobile.running': 'Service running',
            'mobile.openBrowser': 'Open in browser',
            'mobile.copyLink': 'Copy link',
            'mobile.start': 'Start service',
            'mobile.stop': 'Stop service',
            'mobile.refreshQr': 'Refresh QR code',
            'mobile.helpTitle': 'How to use',
            'mobile.help': '1. Phone and PC on the same Wi‑Fi<br>2. Click Start service<br>3. Scan the QR code on your phone<br>4. Capture/upload — OCR runs automatically<br>5. Results appear on the phone in real time',
            'input.placeholder': 'Describe what you need, or paste images/tables… (multi-file OK)',
            'input.send': 'Send',
            'input.stop': '⏹ Stop',
            'input.upload': 'Upload files (multi-select)',
            'input.reference': 'Add a reference range (then select cells in Excel)',
            'history.search': 'Search history…',
            'history.empty': 'No history yet',
            'history.pin': 'Pin',
            'history.unpin': 'Unpin',
            'history.export': 'Export this chat',
            'history.edit': 'Rename',
            'history.delete': 'Delete',
            'history.pinnedBadge': 'Pinned',
            'history.groupPinned': 'Pinned',
            'history.groupToday': 'Today',
            'history.groupYesterday': 'Yesterday',
            'history.multiSelect': 'Select to delete',
            'history.deleteSelected': 'Delete selected',
            'history.clearAll': 'Clear all history',
            'history.selectAll': 'Select all',
            'history.selectedCount': '{n} selected',
            'history.done': 'Done',
            'history.tempStorage': '💡 Using temporary storage — history resets on refresh',
            'history.pinnedToast': 'Pinned',
            'history.unpinnedToast': 'Unpinned',
            'welcome.setup': 'Capture screenshots, tables, forms, and documents without a model. Set up your own model when you need to clean, split, or chart.',
            'gate.needLlm': 'Set up an AI model first: open Settings for an API key, or start Ollama and pick a model.',
            'welcome.openOcr': 'Capture a document',
            'welcome.openSettings': 'Open AI settings',
            'welcome.tryThese': 'Try these (click to fill the box)',
            'welcome.helpLink': 'Quick start & FAQ',
            'welcome.excel.subtitle': 'Turn screenshots, tables, forms, and documents into a sheet',
            'welcome.excel.desc': 'Recognize a document into the current sheet first. Connect your own model when you need to clean, split, or chart. The subscription is the software seat only.',
            'welcome.word.subtitle': 'Your AI partner for Word',
            'welcome.word.desc': 'Polish, summarize, format, and export documents with natural language.',
            'welcome.ppt.subtitle': 'Your AI partner for PowerPoint',
            'welcome.ppt.desc': 'Generate outlines, refine slides, add speaker notes, and export — faster decks.',
            'phase.turn': 'Getting ready…',
            'phase.step': 'Working on the next step…',
            'phase.thinking': 'Understanding your request…',
            'phase.connecting': 'Connecting to AI…',
            'phase.tool': 'Running an action…',
            'phase.toolWrite': 'Writing / updating content…',
            'phase.toolExport': 'Exporting / saving…',
            'phase.toolRead': 'Reading data…',
            'phase.summarizing': 'Putting results together…',
            'phase.done': 'Done',
            'phase.canceled': 'Stopped',
            'phase.processing': 'Working…',
            'phase.stepOf': 'Step {n}/{total}',
            'phase.runningTool': 'Running: {name}',
            'progress.pending': 'Pending',
            'progress.running': 'Running',
            'progress.completed': 'Done',
            'progress.failed': 'Failed',
            'danger.title': '⚠️ Confirm risky action',
            'danger.cancel': 'Cancel',
            'danger.continue': 'Continue anyway',
            'danger.preview': 'What will be affected',
            'danger.hintDefault': 'After you confirm, similar risky actions in this turn won’t ask again. Esc or click outside to cancel.',
            'danger.hintAlwaysAsk': '“Always ask” is on: every risky action in this turn will ask again. Esc or click outside to cancel.',
            'danger.followDefault': '(Later similar actions in this turn will reuse your choice.)',
            'danger.followAlwaysAsk': '(“Always ask” is on: each risky action will ask again.)',
            'settings.title': '🤖 AI model settings',
            'settings.conn': '🔌 Connection',
            'settings.provider': 'AI provider',
            'settings.apiUrl': 'API URL',
            'settings.apiKey': 'API key',
            'settings.apiKeyPh': 'Enter your API key',
            'settings.modelCard': '🧠 Model',
            'settings.modelName': 'Model name',
            'settings.modelPh': 'Pick or type a model name',
            'settings.detectLocal': '🔍 Detect local models',
            'settings.params': '⚙️ Parameters',
            'settings.temperature': 'Temperature',
            'settings.context': 'Context length',
            'settings.ctx.auto': 'Auto (recommended)',
            'settings.ctx.short': 'Short (last 1 turn)',
            'settings.ctx.medium': 'Medium (last 3 turns)',
            'settings.ctx.long': 'Long (last 10 turns)',
            'settings.ctx.unlimited': 'Full history',
            'settings.uiLocale': 'UI language',
            'settings.uiLocale.zh': '中文',
            'settings.uiLocale.en': 'English',
            'settings.dangerAsk': 'Risky actions',
            'settings.dangerAskLabel': 'Always ask for confirmation',
            'settings.dangerAskHint': 'When on, each delete/clear-style action asks again; “Allow” is not reused for the turn.',
            'settings.testCard': 'Connection test',
            'settings.testBtn': 'Test connection',
            'settings.testHint': 'Click the button above to verify the API connection',
            'settings.save': '💾 Save AI settings',
            'settings.dataCard': 'Data and risk',
            'settings.data.cloud': 'With a cloud model, your text, the current selection, and recent turns are sent to the API address you entered. That server may be outside this PC and outside your country.',
            'settings.data.local': 'Ollama or llama.cpp aimed at this PC keeps prompts local. Point the address at a remote server and the data goes there.',
            'settings.data.tools': 'Some tools edit the open workbook, document, or slides, export files, or run VBA / Python on this PC. Delete and clear actions ask first.',
            'settings.data.keys': 'API keys stay on this PC under your Windows user and are not shipped in the installer. Do not share llm_config.json. Enterprises can point Custom API at a private model.',
            'settings.undoNote': 'Undo on the chat pane reverts the add-in step. It does not undo edits you can revert with Ctrl+Z in the worksheet.',
            'settings.deleteLocal': 'Delete local keys and capture history',
            'settings.deleteLocalConfirm': 'Click again to confirm',
            'settings.deleteLocalDone': 'Local keys and capture history were deleted. Content already written into the workbook is unchanged.',
            'settings.deleteLocalFail': 'Could not finish deleting. Try again later.',
            'settings.closeEsc': 'Close (Esc)',
            'settings.provider.ollama': 'Ollama (local)',
            'settings.provider.llamacpp': 'llama.cpp (local)',
            'settings.localGuide.title': '📥 Local model setup',
            'settings.localGuide.statusNone': 'No local model detected yet. Finish the steps below, then click Detect local models.',
            'settings.localGuide.statusOk': 'Found {count} local model(s). Pick one in the model list above.',
            'settings.localGuide.ollamaHint': 'Ollama is free and open source. Model download and inference stay on this PC, so it keeps working offline.',
            'settings.localGuide.ollamaStep1': 'Download and install Ollama (Windows installer, about 700 MB).',
            'settings.localGuide.downloadOllama': '⬇ Download Ollama',
            'settings.localGuide.ollamaStep2': 'Open PowerShell and pull a model (first download is about 1-4 GB):',
            'settings.localGuide.ollamaPullCmd': 'ollama pull qwen2.5:7b',
            'settings.localGuide.copy': 'Copy',
            'settings.localGuide.copied': 'Copied',
            'settings.localGuide.copyFail': 'Copy failed. Please select the command manually.',
            'settings.localGuide.openFail': 'Could not open the browser automatically. Copy the link into a browser manually:',
            'settings.localGuide.ollamaStep3': 'Ollama runs in the background after install. If the connection still fails, start it manually:',
            'settings.localGuide.ollamaServeCmd': 'ollama serve',
            'settings.localGuide.llamacppHint': 'llama.cpp is free and open source. llama-server exposes an OpenAI-compatible API on http://localhost:8080 by default.',
            'settings.localGuide.llamacppStep1': 'Download a llama.cpp release (Windows zip, no install needed):',
            'settings.localGuide.downloadLlamacpp': '⬇ Download llama.cpp',
            'settings.localGuide.llamacppStep2': 'Download a GGUF model from Hugging Face into the models folder of the unzipped directory.',
            'settings.localGuide.llamacppStep3': 'Open PowerShell in that folder, start the server, and keep the window open:',
            'settings.localGuide.llamacppServeCmd': 'llama-server -m models\\qwen2.5-7b-instruct-q4_k_m.gguf --port 8080',
            'settings.localGuide.detectNow': '🔍 Detect local models',
            'settings.provider.wenxin': 'Baidu ERNIE',
            'settings.provider.qianwen': 'Alibaba Qwen',
            'settings.provider.zhipu': 'Zhipu AI (ChatGLM)',
            'settings.provider.xunfei': 'iFLYTEK Spark',
            'settings.provider.moonshot': 'Moonshot (Kimi)',
            'settings.provider.deepseek': 'DeepSeek',
            'settings.provider.volcengine': 'Volcengine (Doubao)',
            'settings.provider.custom': 'Custom API',
            'settings.saved': 'Settings saved',
            'license.card': 'Account / Subscription',
            'license.byok': 'Your subscription covers the SmartPilot software license (AI chat and OCR modules). You bring your own model / cloud OCR API keys; those providers bill usage separately.',
            'license.scope': 'License covers AI + OCR',
            'license.openAccount': 'Open account & subscription',
            'license.badge.trial': 'Trial / inactive',
            'license.badge.active': 'Subscribed',
            'license.badge.none': 'Account',
            'license.status.loading': 'Loading subscription…',
            'license.status.none': 'Not signed in',
            'license.status.active': 'Active: {plan}, ends {end}',
            'license.status.activeOffline': 'Active (key): {plan}, ends {end}',
            'license.status.inactive': 'Inactive ({reason})',
            'license.login': 'Email or phone',
            'license.password': 'Password',
            'license.btnLogin': 'Sign in',
            'license.btnRegister': 'Start trial',
            'license.btnRefresh': 'Refresh',
            'license.btnStripeMonth': 'Upgrade (Stripe monthly)',
            'license.btnStripeYear': 'Upgrade (Stripe yearly)',
            'license.btnCnYear': 'Upgrade (CN yearly)',
            'license.btnPortal': 'Manage billing',
            'license.btnLogout': 'Sign out',
            'license.key': 'License key',
            'license.key.hint': 'Paste the key you received after purchase (starts with SPK1-). It is verified on this PC — no internet or account needed.',
            'license.key.hintRenew': 'Already active. To renew, click "Enter new key", paste the new key, then click Activate to extend the period.',
            'license.key.placeholder': 'SPK1-…',
            'license.btnActivate': 'Activate',
            'license.btnRenewKey': 'Enter new key',
            'license.btnClearKey': 'Remove activation',
            'license.local.active': 'Offline activation: {plan}, ends {end}',
            'license.local.none': 'No activation key in use',
            'license.key.ok': 'Activated',
            'license.key.cleared': 'Local activation cleared',
            'ocr.quota.exhausted': 'Daily free quota used up ({limit}/day). It resets at midnight; activate a license for unlimited OCR.',
            'ocr.quota.badge': '{remaining}/{limit} left today',
            'ocr.quota.unlimited': 'Activated · unlimited',
            'ocr.quota.cloudFree': 'Cloud OCR not counted',
            'license.btnRefund': 'Deactivate & request refund',
            'license.refund.confirm': 'Deactivate this PC? You will lose access immediately, and this key can never be activated here again. You will get a deactivation code — send it to support, who refunds only after verifying it.',
            'license.refund.codeTitle': 'Deactivation code (send to support for refund)',
            'license.refund.codeHint': 'Copy it and send to support. They refund after verification. This code cannot restore access — buy again if you want to continue.',
            'license.refund.copy': 'Copy deactivation code',
            'license.refund.copied': 'Copied — send it to support',
            'license.refund.done': 'Deactivated — this PC is no longer licensed',
            'license.refund.none': 'No local activation to deactivate',
            'license.refund.failed': 'Deactivation failed',
            'license.key.empty': 'Paste an activation key first',
            'license.buy.title': 'Purchase / Renew',
            'license.buy.titleRenew': 'Renew',
            'license.buy.price': 'Activation plans (mention your choice when paying)',
            'license.buy.planMonthly': 'Pro · Monthly —— ¥19 / month',
            'license.buy.planYearly': 'Pro · Yearly —— ¥199 / year (1 seat included)',
            'license.buy.planTeam': 'Team (5+ people) —— ¥129 / seat / year (one independent activation code per person; total = 129 × people)',
            'license.buy.planHint': 'Validity is set by the activation code: monthly 30 days, yearly 365 days, team per year. Once activated it works offline — no internet or account login required.',
            'license.buy.step1': 'Scan to add me on WeChat (or copy the contact below)',
            'license.buy.step2': 'Transfer the payment via WeChat; leave your email or WeChat ID in the note',
            'license.buy.step3': 'Send me the payment screenshot; the activation code arrives within minutes — paste it above',
            'license.buy.qr': 'WeChat QR (scan to add)',
            'license.buy.qrMissing': 'WeChat QR not configured (drop it at UI/images/wechat-qr.jpg); please use the contact below',
            'license.buy.contactLabel': 'Contact',
            'license.buy.copyContact': 'Copy contact',
            'license.buy.copied': 'Contact copied',
            'license.buy.copyFailed': 'Copy failed — please copy manually',
            'license.buy.contactEmpty': 'Contact not configured; reach the author via the docs',
            'license.err.credentials': 'Enter account and password',
            'license.err.auth': 'Sign-in failed',
            'license.err.weakPassword': 'Password must be at least 8 characters',
            'license.err.serverDown': 'Cannot reach the license service. Try again later.',
            'license.err.noBridge': 'Host bridge unavailable — open this from the Office task pane.',
            'license.passwordHint': 'At least 8 characters',
            'license.server.ok': 'License server: connected',
            'license.server.down': 'License server: offline',
            'license.server.offline': 'License server: offline (activation key works standalone)',
            'license.server.bridgeOnly': 'License server: via host bridge (page cannot reach localhost)',
            'license.err.checkout': 'Could not open checkout',
            'license.err.portal': 'Could not open billing portal',
            'license.ok.auth': 'Signed in',
            'license.ok.logout': 'Signed out',
            'license.ok.checkout': 'Checkout opened in your browser',
            'license.checkout.pending': 'Paid upgrade is not available yet. You can keep using the trial. This will not open a simulated payment.',
            'license.trialScope': 'Registration starts a 14-day trial for chat and cloud OCR. Those stop when the trial ends, until a subscription is active. Upgrade buttons stay hidden until checkout is open. Local capture works without signing in.',
            'license.gate.needAccount': 'Chat and cloud OCR need a sign-in. Open Account in the sidebar to start a trial. Until checkout is open, local capture works without an account.',
            'license.gate.blocked': 'Trial ended or subscription inactive. Open Account in the side toolbar to sign in. Paid upgrade is not available yet, so an expired account cannot use licensed features.',
            'help.title': 'SmartPilot Help',
            'help.host': 'Host: {label}',
            'help.close': 'Close (Esc)',
            'help.tabStart': 'Quick start',
            'help.tabFaq': 'FAQ',
            'help.tabShortcuts': 'Shortcuts',
            'help.tabAbout': 'About',
            'help.about.version': 'UI version {version}',
            'help.about.diag': 'Export diagnostics',
            'help.about.diagHint': 'Version, host, provider, and API hostname only. No keys and no document text.',
            'help.disclosure.title': 'Before sending to a model',
            'help.disclosure.accept': 'I understand, continue',
            'help.disclosure.cancel': 'Cancel',
            'help.exportTitle': 'Export current chat',
            'help.exportHistoryTitle': 'Export history chat',
            'help.exportHint': 'Download the conversation for backup or sharing.',
            'help.exportMd': 'Export Markdown (.md)',
            'help.exportTxt': 'Export plain text (.txt)',
            'help.start.1.t': 'Capture a document first',
            'help.start.1.d': 'Use Capture to sheet on the ribbon for a screenshot, table, form, or document. Local recognition writes into the current sheet. Set up a model when you need to clean, chart, or edit text.',
            'help.start.1.a': 'Open capture',
            'help.start.2.t': 'Describe your task',
            'help.start.2.d': 'Type what you need in plain language, or use a quick prompt from the welcome page / sidebar (tuned for {host}).',
            'help.start.2.a': 'Go to chat',
            'help.start.3.t': 'Review and continue',
            'help.start.3.d': 'Watch the progress line above the input; open the folder after export; retry on failure. Export the chat anytime for backup.',
            'help.start.3.a': 'Export this chat',
            'help.sc.send': 'Send message',
            'help.sc.help': 'Show help',
            'help.sc.list': 'Shortcut list',
            'help.sc.esc': 'Close dialog / cancel risk confirm',
            'faq.common.0.q': 'I see “Please configure an AI model” — what now?',
            'faq.common.0.a': 'Open ⚙️ Settings, enter a cloud API key and pick a model; or start local Ollama, select a detected model, then Save.',
            'faq.common.1.q': 'Where are exported files?',
            'faq.common.1.a': 'After a successful export, the tool card shows the full path — use Open folder or Copy path. Default is usually the Desktop.',
            'faq.common.2.q': 'A tool failed — what can I do?',
            'faq.common.2.a': 'On the failure card, use Retry or Edit params & retry. Missing args open a short form before re-running.',
            'faq.common.3.q': 'The UI looks outdated?',
            'faq.common.3.a': 'WebView2 caches pages. Close and reopen the SmartPilot task pane (no need to restart Office) to load the latest UI.',
            'faq.common.4.q': 'What is the risky-action confirm?',
            'faq.common.4.a': 'Clear/delete-style changes ask first. Esc cancels. After you allow, similar actions in this turn won’t ask again (unless Always ask is on).',
            'faq.common.5.q': 'Is my spreadsheet uploaded?',
            'faq.common.5.a': 'Cloud models send the text and selection needed for this turn to the API address you set (which may be abroad). Ollama / llama.cpp on localhost keeps prompts on this PC. Keys are not in the installer. See Help → About, or the Legal folder.',
            'faq.excel.0.q': 'Where do I start in Excel?',
            'faq.excel.0.a': 'Use Capture to sheet for a screenshot, table, form, or document. To change the grid, select a range and ask to clean, split, or chart.',
            'faq.excel.1.q': 'How do I fill recognition results into the sheet?',
            'faq.excel.1.a': 'After an image, table, or document is recognized, choose Fill Excel (Excel only). Recognize an ID document only when you have a business need; the result stays in the workbook on this PC.',
            'faq.word.0.q': 'What can I do in Word?',
            'faq.word.0.a': 'Summarize, polish, format, find/replace, export PDF, and more — describe it in natural language or use a quick prompt.',
            'faq.word.1.q': 'Why is there no “Insert into Excel”?',
            'faq.word.1.a': 'You’re in the Word host. Copy or download table results; sheet fill is Excel-only.',
            'faq.ppt.0.q': 'What can I do in PowerPoint?',
            'faq.ppt.0.a': 'Outlines, slide polish, chart slides, speaker notes, export PDF — start from welcome quick prompts.',
            'faq.ppt.1.q': 'PDF export failed?',
            'faq.ppt.1.a': 'Save the presentation at least once; retry from the failure card. Use Open folder after success.',
            'toast.loading': 'Still loading, please wait…',
            'toast.moduleLoading': 'Module still loading, please wait…',
            'undo.ok': 'Undone',
            'redo.ok': 'Redone',
            'undo.fail': 'Undo failed',
            'redo.fail': 'Redo failed',
            'undo.noHost': 'Host not connected',
            'err.network_error': 'Network error — check your connection',
            'err.timeout': 'Request timed out — try again later',
            'err.fetch_failed': 'Could not reach the server',
            'err.cors_error': 'Cross-origin request blocked — contact admin',
            'err.excel_not_ready': '{host} is not ready — make sure the app is running',
            'err.excel_access_denied': 'Cannot access {host} — check permissions',
            'err.excel_operation_failed': '{host} operation failed',
            'err.cell_write_failed': 'Failed to write cells',
            'err.ocr_failed': 'OCR failed — try uploading again',
            'err.ocr_timeout': 'OCR timed out',
            'err.ocr_engine_not_available': 'OCR engine unavailable — check configuration',
            'err.image_invalid': 'Invalid image — use PNG/JPG',
            'err.image_too_large': 'Image too large — compress and retry',
            'err.llm_failed': 'AI response failed — try again later',
            'err.llm_timeout': 'AI response timed out',
            'err.llm_api_error': 'API call failed',
            'err.token_limit_exceeded': 'Context too long — start a new chat',
            'err.config_load_failed': 'Failed to load settings',
            'err.config_save_failed': 'Failed to save settings',
            'err.config_invalid': 'Invalid settings format',
            'err.unknown_error': 'An unknown error occurred',
            'err.permission_denied': 'Permission denied',
            'err.file_not_found': 'File not found',
            'err.operation_cancelled': 'Operation cancelled',
            'ext.title': '🛠️ Extensions',
            'ext.intro': 'Build custom Skills (VBA / vision / hybrid) and connect MCP servers to extend AI capabilities.',
            'ext.tip': '💡 Tip: In chat, say “create a Skill for …” to generate one quickly.',
            'ext.tabSkills': '🛠️ Skills',
            'ext.tabMcp': '🔌 MCP',
            'mcp.server': '🖥️ MCP Server',
            'mcp.notRunning': 'Stopped',
            'mcp.running': 'Running',
            'mcp.serverDesc': 'Expose Excel actions as MCP tools for external AI clients',
            'mcp.portTools': 'Port',
            'mcp.toolsLabel': 'Tools',
            'mcp.start': '▶ Start server',
            'mcp.stop': '⏹ Stop server',
            'mcp.clients': '🔗 MCP Clients',
            'mcp.add': '+ Add',
            'mcp.clientsDesc': 'Connect external MCP servers to extend AI',
            'mcp.noClients': 'No MCP clients connected',
            'mcp.availableTools': '🛠️ Available tools',
            'mcp.noTools': 'No tools available',
            'mcp.reconnect': '🔗 Reconnect',
            'mcp.refresh': '🔄 Refresh',
            'kb.title': '📚 Knowledge (RAG)',
            'kb.intro': 'Import company docs (PDF / Word / Excel). Chat will retrieve relevant snippets as context.',
            'kb.flow': '💡 Flow: create a knowledge base → import docs → test search → enable chat RAG',
            'kb.tabList': '📚 Knowledge bases',
            'kb.tabDocs': '📄 Documents',
            'kb.tabSearch': '🔍 Search test',
            'kb.tabSettings': '⚙️ RAG settings',
            'kb.listTitle': 'Knowledge bases',
            'kb.create': '+ New knowledge base',
            'kb.name': 'Name',
            'kb.namePh': 'e.g. Product handbook',
            'kb.desc': 'Description (optional)',
            'kb.descPh': 'Briefly describe what this knowledge base is for…',
            'kb.cancel': 'Cancel',
            'kb.createBtn': 'Create',
            'kb.searchPh': '🔍 Search by name or description…',
            'kb.filterAll': 'All statuses',
            'kb.filterActive': '✓ Active',
            'kb.filterPending': '⏳ Pending sync',
            'kb.filterArchived': '📦 Archived',
            'kb.empty': 'No knowledge bases yet',
            'kb.emptyHint': 'Click “+ New knowledge base” to get started',
            'kb.docsTitle': 'Documents',
            'kb.docsHint': 'Paste a local file path or OSS URL to import:',
            'kb.pathPh': 'Local path like C:\\docs\\file.pdf, or an OSS URL',
            'kb.browse': 'Browse…',
            'kb.selectKb': '-- Select a knowledge base --',
            'kb.import': 'Import document',
            'kb.refreshDocs': '🔄 Refresh',
            'kb.docsHint2': 'Supported: PDF / Word / Excel / CSV / TXT / MD. Documents are parsed, chunked and embedded automatically; the list refreshes while processing.',
            'kb.envChecking': 'Checking local RAG environment...',
            'kb.envInstallTitle': 'Ollama not found',
            'kb.envInstallDesc': 'RAG embedding requires the local Ollama service (free & open source, ~700MB). Download and install it, then click "Re-check". Ollama and models are NOT bundled with the add-in; install them once.',
            'kb.envDownload': '⬇ Download Ollama',
            'kb.envRecheck': 'Re-check',
            'kb.envStartTitle': 'Ollama service not running',
            'kb.envStartDesc': 'Ollama is installed but the service is not running. Open the Ollama app (or run "ollama serve" in a terminal), then click "Re-check".',
            'kb.envPullTitle': 'Missing embedding model {model}',
            'kb.envPullDesc': 'RAG needs the embedding model {model} (~1.2GB, one-time). Click "Pull" to download automatically, or copy the command to a terminal.',
            'kb.envPullBtn': '⬇ Pull {model}',
            'kb.envPulling': 'Pulling...',
            'kb.envPullDone': 'Model pull complete',
            'kb.envPullFail': 'Model pull failed',
            'kb.envPullStartFail': 'Failed to start pull',
            'kb.envPullCancelled': 'Pull cancelled',
            'kb.envCancel': 'Cancel',
            'kb.envCopyCmd': 'Copy command',
            'kb.envCopied': 'Copied: {cmd}',
            'kb.envErrorTitle': 'Environment check failed',
            'kb.envErrorDesc': 'Unknown error',
            'kb.envNoBridge': 'This host cannot open external links',
            'kb.envOpenFail': 'Failed to open link',
            'kb.retry': 'Retry',
            'kb.searchTitle': 'Search playground',
            'kb.queryPh': 'Enter a test question…',
            'kb.runSearch': 'Search',
            'kb.ragConfig': 'RAG configuration',
            'kb.embedding': 'Embedding service',
            'kb.provider': 'Default provider:',
            'kb.ollamaEndpoint': 'Local Ollama endpoint:',
            'kb.localModel': 'Local model name:',
            'kb.cloudEndpoint': 'Cloud endpoint (used when provider is a cloud API):',
            'kb.cloudEndpointPh': 'e.g. https://dashscope.aliyuncs.com/api/v1',
            'kb.cloudApiKey': 'Cloud API key:',
            'kb.cloudModel': 'Cloud model name:',
            'kb.apiKeyKeepPh': 'Leave empty to keep unchanged',
            'kb.qdrant': 'Vector DB (Qdrant)',
            'kb.apiKeyOpt': 'API key (optional):',
            'kb.apiKeyPh': 'Leave empty for no auth',
            'kb.chatEnhance': 'Chat enhancement',
            'kb.enableRag': 'Enable RAG in chat',
            'kb.defaultKb': 'Default knowledge base (auto-retrieved in chat):',
            'kb.defaultKbNone': '-- None (disable chat enhancement) --',
            'kb.topk': 'Default TopK:',
            'kb.threshold': 'Similarity threshold (0–1):',
            'kb.maxChunks': 'Max context chunks:',
            'kb.saveCfg': 'Save settings',
            'kb.testConn': 'Test connection',
            'tp.title': '📋 Task planner',
            'tp.autoBypass': 'Skip confirms for scheduled runs',
            'tp.newTask': '+ New task',
            'tp.refresh': '🔄 Refresh',
            'tp.standalone': '🗂 Open standalone list',
            'tp.standaloneTitle': 'Open standalone task list (service + page; keeps running after Excel closes)',
            'tp.all': 'All',
            'tp.empty': 'No scheduled tasks',
            'tp.running': 'Running',
            'tp.paused': 'Paused',
            'tp.formTitle': 'New scheduled task',
            'tp.taskName': 'Task name',
            'tp.taskNamePh': 'e.g. Daily data clean',
            'tp.freq': 'Frequency',
            'tp.freqDaily': 'Daily',
            'tp.freqHourly': 'Hourly',
            'tp.freqWeekdays': 'Weekdays',
            'tp.freqWeekly': 'Every Monday',
            'tp.freqOnce': 'Once',
            'tp.time': 'Run time',
            'tp.command': 'Command',
            'tp.commandPh': 'e.g. Clean the current sheet',
            'editor.title': 'Code editor',
            'editor.lang': 'Switch language',
            'editor.clear': 'Clear editor',
            'editor.format': 'Format code',
            'editor.saveSnippet': 'Save snippet',
            'editor.library': 'Code library',
            'editor.run': 'Run code',
            'editor.debug': 'Open debugger',
            'editor.placeholder': 'Paste or edit VBA code here…',
            'tool.copyPath': 'Copy path',
            'tool.openFolder': 'Open folder',
            'tool.retry': 'Retry',
            'tool.editRetry': 'Edit params & retry',
            'tool.copiedPath': 'Path copied. Paste it in Explorer’s address bar (rebuild the add-in for one-click Open folder).',
            'tool.openFolderFail': 'Could not open folder. Locate manually: {path}',
            'toast.exported': 'Exported {n} messages',
            'toast.pathCopied': 'Path copied',
            'toast.copyFail': 'Copy failed',
            'toast.retryFilled': 'Retry prompt filled — send when ready',
            'toast.retrying': 'Retrying…',
            'toast.retryDone': 'Retry finished',
            'toast.retryFail': 'Retry failed',
            'toast.retryNoTool': 'Cannot retry: missing tool name',
            'feat.excel.0.t': 'Capture to sheet', 'feat.excel.0.d': 'Screenshots, tables, forms, and documents into the current sheet',
            'feat.excel.1.t': 'Clean and split', 'feat.excel.1.d': 'Deduplicate, merge, split by column',
            'feat.excel.2.t': 'Chart the selection', 'feat.excel.2.d': 'Build a chart from the current sheet',
            'feat.excel.3.t': 'Make a deck', 'feat.excel.3.d': 'Turn this sheet into PPT slides',
            'feat.word.0.t': 'Polish & summarize', 'feat.word.0.d': 'Summaries, rewrites, translation',
            'feat.word.1.t': 'Smart layout', 'feat.word.1.d': 'TOC, styles, tables',
            'feat.word.2.t': 'Find & replace', 'feat.word.2.d': 'Search and bulk replace',
            'feat.word.3.t': 'Proofreading', 'feat.word.3.d': 'Typos and awkward phrasing',
            'feat.word.4.t': 'Custom skills', 'feat.word.4.d': 'Document skills in one sentence',
            'feat.word.5.t': 'Export PDF', 'feat.word.5.d': 'Export the current document',
            'feat.ppt.0.t': 'Outline', 'feat.ppt.0.d': 'Multi-slide outlines by topic',
            'feat.ppt.1.t': 'Slide polish', 'feat.ppt.1.d': 'Layout and color themes',
            'feat.ppt.2.t': 'Chart slides', 'feat.ppt.2.d': 'Add chart slides for data',
            'feat.ppt.3.t': 'Speaker notes', 'feat.ppt.3.d': 'Notes for the current slide',
            'feat.ppt.4.t': 'Custom skills', 'feat.ppt.4.d': 'Deck skills in one sentence',
            'feat.ppt.5.t': 'Export PDF', 'feat.ppt.5.d': 'Export the presentation',
            'ppt.title': 'PPT Theme Templates',
            'ppt.badge': 'WYSIWYG',
            'ppt.dirty': 'Unsaved changes',
            'ppt.hint': 'Edit colors and layout — the live preview updates instantly. After saving, say “use [theme name] to generate a PPT” in chat.',
            'ppt.info': 'Theme info',
            'ppt.name': 'Theme name',
            'ppt.namePh': 'e.g. Company Red Gold',
            'ppt.id': 'Theme ID',
            'ppt.idPh': 'Auto-generated',
            'ppt.font': 'Body font',
            'ppt.presets': 'Color presets',
            'ppt.colors': 'Fine-tune colors',
            'ppt.contrastOk': 'Body/background contrast looks good',
            'ppt.mine': 'My themes',
            'ppt.loading': 'Loading…',
            'ppt.reset': 'Reset',
            'ppt.saveAs': 'Save as',
            'ppt.save': 'Save',
            'ppt.live': 'Live preview',
            'ppt.guides': 'Guides',
            'ppt.editLayout': 'Edit layout',
            'ppt.currentPage': 'Current:',
            'ppt.hintBrowse': 'Click a thumbnail to switch · click titles to edit sample text',
            'ppt.hintLayout': 'Drag blocks to adjust title/body/table · used when generating PPT',
            'ppt.addDeco': '+ Accent shape',
            'ppt.resetLayout': 'Reset default layout',
            'ppt.canvasTip': 'Drag shapes on the preview · resize from corner · design size 960×720',
            'ppt.capBullets': 'Bullets/page ≤',
            'ppt.capRows': 'Table rows ≤',
            'ppt.capCols': 'Table cols ≤',
            'ppt.capSeries': 'Chart series ≤',
            'ppt.capCats': 'Chart cats ≤',
            'ppt.del': 'Delete',
            'ppt.fill': 'Fill',
            'ppt.slide.title': 'Cover',
            'ppt.slide.table': 'Data table',
            'ppt.slide.kpi': 'KPI',
            'ppt.slide.agenda': 'Agenda',
            'ppt.slide.chart': 'Chart',
            'ppt.slide.text': 'Bullets',
            'ppt.slide.comparison': 'Compare',
            'ppt.slide.timeline': 'Timeline',
            'ppt.slide.section': 'Section',
            'ppt.slide.closing': 'Closing',
            'ppt.preset.business_blue': 'Business Blue',
            'ppt.preset.tech_green': 'Tech Green',
            'ppt.preset.vibrant_orange': 'Vibrant Orange',
            'ppt.preset.elegant_purple': 'Elegant Purple',
            'ppt.preset.academic_navy': 'Academic Navy',
            'ppt.preset.light_fresh': 'Light Fresh',
            'ppt.color.primary': 'Primary',
            'ppt.color.primary.tip': 'Titles, headers, section fill',
            'ppt.color.secondary': 'Secondary',
            'ppt.color.secondary.tip': 'Secondary accents, bars',
            'ppt.color.accent': 'Accent',
            'ppt.color.accent.tip': 'Highlights, accent bars',
            'ppt.color.bg': 'Background',
            'ppt.color.bg.tip': 'Page background',
            'ppt.color.light': 'Light block',
            'ppt.color.light.tip': 'Zebra rows, card fill',
            'ppt.color.text': 'Body text',
            'ppt.color.text.tip': 'Titles and body',
            'ppt.color.gray': 'Muted gray',
            'ppt.color.gray.tip': 'Subtitles, captions',
            'ppt.contrast.good': 'Contrast {r}:1 · easy to read',
            'ppt.contrast.warn': 'Contrast {r}:1 · darken text or lighten background',
            'ppt.contrast.bad': 'Low contrast {r}:1 · may be hard to see on a projector',
            'ppt.toast.preset': 'Applied “{name}” colors',
            'ppt.toast.resetLayout': 'Restored default layout (960×720)',
            'ppt.toast.needName': 'Enter a theme name first',
            'ppt.toast.loadBad': 'Load failed: invalid theme data',
            'ppt.toast.loadEmpty': 'Load failed: empty theme',
            'ppt.toast.loaded': 'Loaded “{name}”{layout}',
            'ppt.toast.loadedLayout': ' (with layout)',
            'ppt.toast.notFound': 'Theme not found: {id}',
            'ppt.toast.reset': 'Reset to default colors and layout',
            'ppt.toast.saved': '✓ Saved{layout}: {name}',
            'ppt.toast.savedLayout': ' (colors + layout)',
            'ppt.toast.saveFail': 'Save failed: {err}',
            'ppt.toast.savedAs': '✓ Saved as a new theme',
            'ppt.toast.deleted': 'Deleted: {id}',
            'ppt.toast.deleteFail': 'Delete failed: {err}',
            'ppt.toast.unknown': 'Unknown error',
            'ppt.empty': 'No custom themes yet. Tune colors/layout, then Save.',
            'ppt.loadFail': 'Load failed: {err}',
            'ppt.bridgeFail': 'Host bridge unavailable (open from the task pane 🎨)',
            'ppt.layoutTag': ' · layout',
            'ppt.loadTip': 'Click to load',
            'ppt.confirmDelete': 'Delete theme “{id}”?',
            'pomo.title': 'Pomodoro',
            'pomo.titleStandalone': 'Pomodoro · Standalone',
            'pomo.connecting': 'Connecting…',
            'pomo.focus': 'Focus',
            'pomo.break': 'Break',
            'pomo.session': 'Pomodoro {n} / {max}',
            'pomo.start': 'Start',
            'pomo.pause': 'Pause',
            'pomo.resume': 'Resume',
            'pomo.reset': 'Reset',
            'pomo.statTotal': 'Total pomodoros',
            'pomo.statMinutes': 'Focus minutes',
            'pomo.statToday': 'Done today',
            'pomo.footer': 'Standalone · syncs both ways with Excel; saves locally when offline',
            'pomo.settings': 'Pomodoro settings',
            'pomo.workDur': 'Focus length (minutes)',
            'pomo.shortDur': 'Short break (minutes)',
            'pomo.longDur': 'Long break (minutes)',
            'pomo.longEvery': 'Long break every N pomodoros',
            'pomo.sound': 'Play sound when done',
            'pomo.notif': 'Desktop notifications (permission on first use)',
            'pomo.cancel': 'Cancel',
            'pomo.save': 'Save',
            'pomo.min': '{n} min',
            'pomo.count': '{n}',
            'pomo.unitMin': 'min',
            'pomo.unitCount': '',
            'pomo.syncTip': 'Sync status / sync now',
            'pomo.gearTip': 'Settings',
            'pomo.synced': 'Synced',
            'pomo.offline': 'Offline',
            'pomo.syncing': 'Syncing…',
            'pomo.offlineBanner': 'Excel service offline — using local mode',
            'pomo.toast.running': 'Pomodoro is already running',
            'pomo.toast.resetStart': 'Reset — click Start again',
            'pomo.toast.focusStart': '▶ Focus started — you’ve got this!',
            'pomo.toast.notRunning': 'Not running',
            'pomo.toast.paused': '⏸ Paused',
            'pomo.toast.resetTo': '↺ Reset to {time}',
            'pomo.toast.saved': 'Settings saved',
            'pomo.toast.synced': 'Synced with Excel',
            'pomo.toast.offlineMode': 'Excel offline — switched to local mode',
            'pomo.notify.focusTitle': 'Pomodoro · Focus done',
            'pomo.notify.focusBody': 'Take a break — stretch a bit!',
            'pomo.notify.breakTitle': 'Pomodoro · Break over',
            'pomo.notify.breakBody': 'Time for another focus block!',

            // ── Language switch / quick launch ──
            'nav.language': 'Switch UI language',
            'nav.languageSwitched': 'Switched to {lang}',

            // ── Task planner ──
            'tp.runNow': 'Run now',
            'tp.model': 'Execution model (optional; leave empty to use the default model above)',
            'tp.modelAuto': 'AUTO (follow default model)',
            'tp.modelRefreshTitle': 'Re-detect local Ollama models / refresh cloud model list',
            'tp.refreshBtn': '🔄 Refresh',
            'tp.modelHint': 'Optional: pin a model for this task, e.g. local Ollama qwen2.5, deepseek-r1 for reasoning',
            'tp.targetFile': 'Target Excel file (standalone runner, optional)',
            'tp.targetFilePh': 'e.g. C:\\Users\\...\\sales.xlsx — the standalone service operates this file after Excel closes',
            'tp.browse': '📁 Browse',
            'tp.cancel': 'Cancel',
            'tp.save': 'Save',

            // ── Snippets ──
            'snippet.libraryTitle': 'Code library',
            'snippet.close': 'Close',
            'snippet.saveTitle': 'Save snippet',
            'snippet.name': 'Name',
            'snippet.namePh': 'e.g. batch data cleanup',
            'snippet.desc': 'Description (optional)',
            'snippet.descPh': 'Briefly describe what this code does',
            'snippet.code': 'Code',
            'snippet.codePh': 'VBA code...',
            'snippet.saveAsSkill': 'Also save as Skill (recommended, supports keyword matching)<br><small style="color: #666;">Once saved, you can invoke this snippet with a plain-language request</small>',
            'snippet.cancel': 'Cancel',
            'snippet.save': 'Save',

            // ── Knowledge base ──
            'kb.searchKbPh': '🔍 Search knowledge base name or description...',
            'kb.embLocalBge': 'Local BGE (Ollama)',
            'kb.embQianwen': 'Qwen (Tongyi Qianwen)',
            'kb.embZhipu': 'Zhipu AI',

            // ── Skill editor ──
            'skill.addTitle': 'Add Skill',
            'skill.viewSource': '📄 View Markdown source',
            'skill.tabForm': 'Form editor',
            'skill.tabJson': 'JSON source',
            'skill.basicInfo': 'Basic info',
            'skill.name': 'Skill name (English ID)',
            'skill.namePh': 'e.g. MyCustomSkill',
            'skill.displayName': 'Display name',
            'skill.displayPh': 'e.g. My custom action',
            'skill.desc': 'Description',
            'skill.descPh': 'Describe what this Skill does',
            'skill.keywords': 'Keywords (comma separated)',
            'skill.keywordsPh': 'e.g. custom,action,processing',
            'skill.processing': 'Processing',
            'skill.procMode': 'Processing mode',
            'skill.procVba': 'VBA macro - classic Excel automation',
            'skill.procLlm': 'Multimodal LLM - AI vision model',
            'skill.procHybrid': 'Hybrid - auto select',
            'skill.procOcr': 'Local OCR - PaddleOCR engine',
            'skill.codeTpl': 'Code template',
            'skill.codePh': 'Write the code here\nVBA example:\nSub MyMacro()\n  \' your logic\nEnd Sub',
            'skill.examples': 'Sample inputs (comma separated)',
            'skill.examplesPh': 'e.g. run my action, execute custom macro',
            'skill.specParams': 'Spec parameters',
            'skill.advanced': 'Advanced options',
            'skill.llmSysPrompt': 'LLM system prompt',
            'skill.llmSysPh': 'You are a professional AI assistant. Complete the task based on the user input...\n\n[Rules]\n1. ...\n2. ...',
            'skill.llmSysHint': 'Defines the AI role and behavior rules (multimodal LLM mode)',
            'skill.llmUserTpl': 'LLM user prompt template',
            'skill.llmUserPh': 'Process the following: {input}',
            'skill.llmUserHint': 'Use {input} as the placeholder for user input',
            'skill.outputFormat': 'Output format',
            'skill.outText': 'Plain text',
            'skill.outJson': 'JSON',
            'skill.outVba': 'VBA code',
            'skill.outHtmlBrowser': 'HTML (browser view)',
            'skill.outHtml': 'HTML output',
            'skill.outTable': 'Table',
            'skill.outChart': 'Chart',
            'skill.outPivot': 'Pivot table',
            'skill.version': 'Version',
            'skill.tags': 'Tags (comma separated)',
            'skill.tagsPh': 'e.g. OCR,image,AI',
            'skill.supportsImage': 'Supports image input',
            'skill.supportsImageTitle': 'When enabled, users can upload images in chat and the AI will process them (OCR, image analysis, etc.)',
            'skill.supportsFile': 'Supports file input',
            'skill.supportsFileTitle': 'When enabled, users can upload files (Excel, PDF, etc.) and the AI will read them',
            'skill.quickShortcut': '⭐ Add to quick launch',
            'skill.quickShortcutTitle': 'Marked Skills appear in the quick-launch area of the welcome panel',
            'skill.quickConfig': 'Quick launch settings',
            'skill.quickIcon': 'Icon (emoji)',
            'skill.quickPrompt': 'Click hint text',
            'skill.quickPromptPh': 'Leave empty to use the first sample',
            'skill.quickHint': 'Once set, this Skill shows up in the "Quick prompts" area of the welcome panel',
            'skill.uiButtons': 'UI action buttons',
            'skill.uiCopy': '📋 Copy',
            'skill.uiSaveMd': '💾 Save MD',
            'skill.uiRegenerate': '🔄 Regenerate',
            'skill.uiHint': 'Controls buttons under AI messages (enable "Save MD" only for Skill-configuration flows)',
            'skill.jsonLabel': 'Skill JSON config',
            'skill.formatJson': 'Format JSON',
            'skill.validateJson': 'Validate JSON',
            'skill.cancel': 'Cancel',
            'skill.restoreDefault': 'Restore defaults',
            'skill.save': 'Save',

            // ── MCP Server editor ──
            'mcpsrv.addTitle': 'Add MCP Server',
            'mcpsrv.name': 'Server name',
            'mcpsrv.namePh': 'e.g. filesystem, sqlite',
            'mcpsrv.desc': 'Description',
            'mcpsrv.descPh': 'Describe what this MCP Server does',
            'mcpsrv.transport': 'Transport',
            'mcpsrv.stdio': 'stdio (local process)',
            'mcpsrv.http': 'HTTP SSE (remote server)',
            'mcpsrv.command': 'Command',
            'mcpsrv.commandPh': 'e.g. npx, uvx, python',
            'mcpsrv.args': 'Arguments (one per line)',
            'mcpsrv.env': 'Environment variables (one per line, KEY=VALUE)',
            'mcpsrv.headers': 'Headers (one per line, Key=Value)',
            'mcpsrv.envHttp': 'Environment variables (one per line, KEY=VALUE)',
            'mcpsrv.enabled': 'Enabled',
            'mcpsrv.cancel': 'Cancel',
            'mcpsrv.test': '🧪 Test connection',
            'mcpsrv.save': 'Save and connect',

            // ── Quick prompts ──
            'quick.barHint': 'Mark a Skill with "⭐ Quick launch" in Skill management to show it here',
            'quick.title': '⚡ Quick prompts',
            'quick.manageTitle': '⚙️ Manage quick prompts',
            'quick.manage': '⚙️ Manage',
            'quick.searchPh': 'Search prompts',
            'quick.addTitle': 'Add quick prompt',
            'quick.add': '＋ Add',
            'quick.resetTitle': 'Restore defaults',
            'quick.done': 'Done',

            // ── Input area / model picker ──
            'model.refreshTitle': 'Refresh model list',
            'think.selectTitle': 'Select thinking mode',
            'think.auto': 'AUTO (smart auto-select)',
            'think.quick': '⚡ Quick',
            'think.standard': '📝 Standard',
            'think.deep': '🔍 Deep',
            'think.expert': '👑 Expert',
            'input.analyzing': 'AI is analyzing...',
            'input.pickRange': 'Select a range in Excel; it will be added automatically...',
            'input.quickCmd': 'Quick commands - type / to pick'
        }
    };

    function canonicalizeTag(loc) {
        return String(loc).trim().toLowerCase().replace(/_/g, '-');
    }

    /**
     * 把 zh-CN / en_US / EN 收成目录短码。无法识别返回 null（不悄悄当成中文）。
     */
    function resolveLocale(loc) {
        if (loc == null || loc === '') return null;
        var s = canonicalizeTag(loc);
        if (!s) return null;
        var i, j, def, aliases, alias, best, bestLen;
        for (i = 0; i < LOCALE_CATALOG.length; i++) {
            def = LOCALE_CATALOG[i];
            if (s === def.id) return def.id;
            aliases = def.aliases || [];
            for (j = 0; j < aliases.length; j++) {
                if (s === aliases[j]) return def.id;
            }
        }
        best = null;
        bestLen = -1;
        for (i = 0; i < LOCALE_CATALOG.length; i++) {
            def = LOCALE_CATALOG[i];
            aliases = [def.id].concat(def.aliases || []);
            for (j = 0; j < aliases.length; j++) {
                alias = aliases[j];
                if (!alias) continue;
                if (s.length > alias.length && s.indexOf(alias) === 0 && s.charAt(alias.length) === '-') {
                    if (alias.length > bestLen) {
                        best = def.id;
                        bestLen = alias.length;
                    }
                }
            }
        }
        return best;
    }

    /** 已知标签收成短码；无法识别时回退到目录第一项。 */
    function normalizeLocale(loc) {
        return resolveLocale(loc) || FALLBACK_LOCALE;
    }

    function localeMeta(loc) {
        var id = (loc == null || loc === '') ? _locale : (resolveLocale(loc) || FALLBACK_LOCALE);
        var def = LOCALE_CATALOG[0];
        for (var i = 0; i < LOCALE_CATALOG.length; i++) {
            if (LOCALE_CATALOG[i].id === id) { def = LOCALE_CATALOG[i]; break; }
        }
        return {
            id: def.id,
            htmlLang: def.htmlLang,
            dateLocale: def.dateLocale,
            badge: def.badge,
            nativeName: def.nativeName
        };
    }

    function localeCatalog() {
        var out = [];
        for (var i = 0; i < LOCALE_CATALOG.length; i++) out.push(localeMeta(LOCALE_CATALOG[i].id));
        return out;
    }

    /** 按目录顺序轮换到下一种语言（末项回到第一项）。 */
    function nextLocale(loc) {
        var id = resolveLocale(loc) || _locale;
        var i = LOCALES.indexOf(id);
        if (i < 0) i = 0;
        return LOCALES[(i + 1) % LOCALES.length];
    }

    function detectDefaultLocale() {
        try {
            var saved = resolveLocale(localStorage.getItem(STORAGE_KEY));
            if (saved) return saved;
        } catch (_) { /* ignore */ }
        try {
            var navRaw = navigator.language || navigator.userLanguage || FALLBACK_LOCALE;
            var nav = resolveLocale(navRaw);
            if (nav) return nav;
            // 未登记的系统语言沿用原规则：非中文环境默认英文
            return resolveLocale('en') || FALLBACK_LOCALE;
        } catch (_) {
            return FALLBACK_LOCALE;
        }
    }

    var _locale = detectDefaultLocale();

    function interpolate(s, vars) {
        if (!vars || typeof vars !== 'object') return String(s);
        return String(s).replace(/\{(\w+)\}/g, function (_, k) {
            return vars[k] != null ? String(vars[k]) : '';
        });
    }

    /**
     * 按顺序查找词条，返回 { value, locale }；都找不到返回 null。
     * 顺序 = 当前语言 → 回退语言（双向回退，避免任一侧漏词时出现空白）。
     */
    function lookup(key) {
        var order = [_locale];
        if (FALLBACK_LOCALE !== _locale) order.push(FALLBACK_LOCALE);
        for (var n = 0; n < LOCALES.length; n++) {
            if (order.indexOf(LOCALES[n]) < 0) order.push(LOCALES[n]);
        }
        for (var i = 0; i < order.length; i++) {
            var pack = STRINGS[order[i]];
            if (pack && pack[key] != null) return { value: pack[key], locale: order[i] };
        }
        return null;
    }

    function has(key) {
        return lookup(key) !== null;
    }

    function t(key, vars) {
        var found = lookup(key);
        return interpolate(found ? found.value : key, vars);
    }

    /**
     * 自检：返回目标语言相对基准语言缺失的 key 列表（默认 en 相对 zh）。
     * 供 Tests 与开发者面板使用，UI 层无需关心。
     */
    function missingKeys(loc) {
        var target = normalizeLocale(loc);
        var base = target === FALLBACK_LOCALE
            ? (LOCALES.length > 1 ? LOCALES[1] : FALLBACK_LOCALE)
            : FALLBACK_LOCALE;
        var tp = STRINGS[target] || {};
        var bp = STRINGS[base] || {};
        var out = [];
        Object.keys(bp).forEach(function (k) {
            if (tp[k] == null) out.push(k);
        });
        return out;
    }

    /** 解析当前环境的宿主桥（WebView2）。网页版本地 HTTP 桥不参与功能区语言同步。 */
    function resolveHostBridge() {
        try {
            if (window.BridgeAccessor) {
                if (BridgeAccessor._usingLocalBridge) return null;   // 网页版：没有功能区
                if (typeof BridgeAccessor.getBridge === 'function') {
                    var b = BridgeAccessor.getBridge();
                    // getBridge() 可能刚刚 init 出「本地 HTTP 桥」，此时同样没有功能区可刷
                    if (b && !BridgeAccessor._usingLocalBridge) return b;
                }
            }
        } catch (_) { /* ignore */ }
        try {
            var hosts = window.chrome && chrome.webview && chrome.webview.hostObjects;
            if (hosts) {
                // 同步代理优先：能在页面刷新/卸载前把语言落盘
                try { if (hosts.sync && hosts.sync.excelBridge) return hosts.sync.excelBridge; } catch (_s) { /* ignore */ }
                if (hosts.excelBridge) return hosts.excelBridge;
            }
        } catch (_) { /* ignore */ }
        try { if (window.App && App.excelBridge) return App.excelBridge; } catch (_) { /* ignore */ }
        return null;
    }

    function localeSyncWarn(msg) {
        try {
            if (window.ChatLog && typeof ChatLog.warn === 'function') ChatLog.warn('I18N', msg);
            else console.warn('[i18n] ' + msg);
        } catch (_) { /* ignore */ }
    }

    function checkLocaleSyncResult(raw) {
        try {
            var s = raw;
            if (typeof s === 'string' && s.charAt(0) === '{') s = JSON.parse(s);
            if (s && s.success === false) {
                localeSyncWarn('宿主功能区语言同步失败: ' + (s.detail || s.error || 'unknown'));
            }
        } catch (_) { /* ignore */ }
    }

    /**
     * 把界面语言同步给宿主（写 %LocalAppData%\SmartPilot\ui-locale.txt 并 Invalidate 功能区）。
     * 注意：不能靠 `typeof bridge.SaveUiLocale === 'function'` 判断可用性 ——
     * WebView2 hostObject 对任意方法名 typeof 都是 function，真实失败要到调用后才暴露；
     * 早退会让「切了语言但功能区不变」完全没有日志可查（2026-10-11 Excel 版根因）。
     */
    function localeSyncErrText(e) {
        if (!e) return 'unknown';
        if (typeof e === 'string') return e;
        if (e.message) return e.message;
        try {
            if (e.parameters && e.parameters.error) return String(e.parameters.error);
        } catch (_) { /* ignore */ }
        return String(e);
    }

    function syncLocaleToHost(loc) {
        // 优先 BridgeAccessor.callAsync：把 WebView2 的「t/e is not a function」吞成结构化 JSON，
        // 避免设置面板等未加载包装层时直接打 raw hostObjects 冒泡未处理拒绝。
        try {
            if (window.BridgeAccessor && typeof BridgeAccessor.callAsync === 'function' && !BridgeAccessor._usingLocalBridge) {
                BridgeAccessor.callAsync('SaveUiLocale', loc).then(checkLocaleSyncResult).catch(function (e) {
                    localeSyncWarn('宿主功能区语言同步异常: ' + localeSyncErrText(e));
                });
                return true;
            }
        } catch (_) { /* fall through */ }

        var bridge = resolveHostBridge();
        if (!bridge) return false;
        try {
            var ret = bridge.SaveUiLocale(loc);
            if (ret && typeof ret.then === 'function') {
                ret.then(checkLocaleSyncResult).catch(function (e) {
                    localeSyncWarn('宿主功能区语言同步异常: ' + localeSyncErrText(e));
                });
            } else {
                checkLocaleSyncResult(ret);
            }
            return true;
        } catch (e) {
            localeSyncWarn('宿主功能区语言同步异常: ' + localeSyncErrText(e));
            return false;
        }
    }

    /**
     * 首屏对齐宿主功能区语言（%LocalAppData%\SmartPilot\ui-locale.txt）。
     *
     * 必须无视本页 localStorage 是否已有值：PPT 主题模板、OCR 等是独立 WebView，
     * 用户数据目录和主聊天不是同一份，localStorage 会停在旧语言。若此处因「已经存过」
     * 就跳过，并在启动时把本页语言 sync 回宿主，会把用户刚切的语言覆盖掉
     * （表现为「语言切换对 PPT 主题模板失效」）。
     * 对齐时 syncHost:false，禁止这个 WebView 把旧语言写回宿主。
     * 返回 true 表示已向宿主发起读取。
     */
    function pullLocaleFromHostIfUnset() {
        var bridge = resolveHostBridge();
        if (!bridge || typeof bridge.GetUiLocale !== 'function') return false;
        try {
            var ret = bridge.GetUiLocale();
            if (ret && typeof ret.then === 'function') {
                ret.then(adoptHostLocale).catch(function () { /* 读失败就保持本页语言，不反写宿主 */ });
                return true;
            }
            adoptHostLocale(ret);
            return true;
        } catch (_) { /* ignore */ }
        return false;
    }

    /**
     * 宿主 / 其它 WebView 推过来的语言。已是当前语言则什么都不做，
     * 避免把主题编辑器里的配色和预览再刷一遍。
     */
    function applyExternalLocale(loc) {
        loc = resolveLocale(loc);
        if (!loc || loc === _locale) return _locale;
        return setLocale(loc, {
            apply: true,
            reloadWelcome: true,
            reloadHistory: true,
            broadcast: false,
            syncHost: false
        });
    }

    function adoptHostLocale(raw) {
        try {
            var s = raw;
            if (typeof s === 'string' && s.charAt(0) === '{') s = JSON.parse(s);
            var loc = s && (s.locale || (s.data && s.data.locale));
            if (!resolveLocale(loc)) return;
            applyExternalLocale(loc);
        } catch (_) { /* ignore */ }
    }

    /** 同一用户数据目录下的其它文档，以及宿主 PostWebMessage，都走这条。 */
    function bindExternalLocaleBus() {
        try {
            window.addEventListener('storage', function (e) {
                if (!e || e.key !== STORAGE_KEY) return;
                if (!resolveLocale(e.newValue)) return;
                applyExternalLocale(e.newValue);
            });
        } catch (_) { /* ignore */ }
        try {
            var wv = window.chrome && chrome.webview;
            if (wv && typeof wv.addEventListener === 'function') {
                wv.addEventListener('message', function (ev) {
                    var d = ev && (ev.data != null ? ev.data : ev);
                    if (typeof d === 'string') {
                        try { d = JSON.parse(d); } catch (_) { return; }
                    }
                    if (!d || (d.type !== LOCALE_MSG && d.type !== 'smartpilot:locale-changed')) return;
                    applyExternalLocale(d.locale);
                });
            }
        } catch (_) { /* ignore */ }
    }

    /**
     * 跨窗口广播。AI 设置面板是嵌入的 iframe（独立 document），番茄钟/调试窗是 window.open
     * 出来的新窗口 —— 任一窗口切换语言后，其它窗口必须同步，否则会出现「一半英文一半中文」。
     */
    var LOCALE_MSG = 'smartpilot:locale-change';

    function broadcastLocale(loc) {
        var payload = { type: LOCALE_MSG, locale: loc };
        try {
            if (window.parent && window.parent !== window && typeof window.parent.postMessage === 'function') {
                window.parent.postMessage(payload, '*');
            }
        } catch (_) { /* ignore */ }
        try {
            if (window.opener && typeof window.opener.postMessage === 'function') {
                window.opener.postMessage(payload, '*');
            }
        } catch (_) { /* ignore */ }
        try {
            for (var i = 0; i < window.frames.length; i++) {
                try { window.frames[i].postMessage(payload, '*'); } catch (_) { /* ignore */ }
            }
        } catch (_) { /* ignore */ }
    }

    var _listeners = [];

    function addLocaleChangeListener(fn) {
        if (typeof fn !== 'function') return function () { /* noop */ };
        _listeners.push(fn);
        return function remove() { removeLocaleChangeListener(fn); };
    }

    function removeLocaleChangeListener(fn) {
        for (var i = _listeners.length - 1; i >= 0; i--) {
            if (_listeners[i] === fn) _listeners.splice(i, 1);
        }
    }

    function fireLocaleChange(loc, prev) {
        var detail = { locale: loc, previousLocale: prev };
        for (var i = 0; i < _listeners.length; i++) {
            try { _listeners[i](detail); } catch (_) { /* 单个监听失败不影响其余 */ }
        }
        try {
            window.dispatchEvent(new CustomEvent('smartpilot:locale-changed', { detail: detail }));
        } catch (_) { /* ignore */ }
    }

    function setLocale(loc, opts) {
        opts = opts || {};
        loc = normalizeLocale(loc);
        var prev = _locale;
        _locale = loc;
        if (opts.persist !== false) {
            try { localStorage.setItem(STORAGE_KEY, loc); } catch (_) { /* ignore */ }
        }
        try {
            if (window.App && App.config) App.config.uiLocale = loc;
        } catch (_) { /* ignore */ }
        try {
            if (document.documentElement) document.documentElement.lang = localeMeta(loc).htmlLang;
        } catch (_) { /* ignore */ }
        if (opts.syncHost !== false) syncLocaleToHost(loc);
        if (opts.apply !== false) apply();
        if (opts.reloadWelcome !== false) {
            try {
                var panel = document.querySelector('.welcome-panel');
                if (panel && panel.parentNode) {
                    panel.parentNode.removeChild(panel);
                    if (typeof window.showWelcomePanel === 'function') window.showWelcomePanel();
                }
            } catch (_) { /* ignore */ }
        }
        if (opts.reloadHistory !== false) {
            try {
                if (typeof window.renderHistory === 'function') window.renderHistory();
                else if (window.App && typeof App.renderHistory === 'function') App.renderHistory();
            } catch (_) { /* ignore */ }
        }
        fireLocaleChange(loc, prev);
        if (opts.broadcast !== false) broadcastLocale(loc);
        return loc;
    }

    /** 按目录顺序切到下一种语言（工具栏按钮 / Alt+L 快捷键入口） */
    function toggleLocale(opts) {
        return setLocale(nextLocale(_locale), opts);
    }

    function getLocale() { return _locale; }

    function setTextContent(el, text) {
        if (el.hasAttribute('data-i18n-html') || /<[a-z][\s\S]*>/i.test(text)) {
            el.innerHTML = text;
        } else {
            el.textContent = text;
        }
    }

    function eachNode(scope, selector, fn) {
        var list;
        try { list = scope.querySelectorAll(selector); } catch (_) { return; }
        if (!list) return;
        for (var i = 0; i < list.length; i++) {
            try { fn(list[i]); } catch (_) { /* 单个节点失败不中断整体刷新 */ }
        }
    }

    /**
     * 把当前语言的文案刷进 DOM。
     * 支持的属性：data-i18n（文本）/ data-i18n-html / data-i18n-title / data-i18n-placeholder
     *             / data-i18n-aria-label / data-i18n-value
     */
    function apply(root) {
        var scope = root || document;
        if (!scope || !scope.querySelectorAll) return;

        eachNode(scope, '[data-i18n]', function (el) {
            var key = el.getAttribute('data-i18n');
            if (!key) return;
            var text = t(key);
            var tag = (el.tagName || '').toUpperCase();
            if (tag === 'INPUT') {
                // <input type=button/submit/reset> 的可见文案就是 value；其余输入框保持用户输入
                var type = (el.getAttribute('type') || 'text').toLowerCase();
                if (type === 'button' || type === 'submit' || type === 'reset') el.value = text;
                return;
            }
            setTextContent(el, text);
        });

        eachNode(scope, '[data-i18n-value]', function (el) {
            var key = el.getAttribute('data-i18n-value');
            if (key) el.value = t(key);
        });

        eachNode(scope, '[data-i18n-title]', function (el) {
            var key = el.getAttribute('data-i18n-title');
            if (!key) return;
            el.setAttribute('title', t(key));
            if (el.hasAttribute('aria-label') || el.hasAttribute('data-i18n-aria')) {
                el.setAttribute('aria-label', t(key));
            }
        });

        eachNode(scope, '[data-i18n-placeholder]', function (el) {
            var key = el.getAttribute('data-i18n-placeholder');
            if (key) el.setAttribute('placeholder', t(key));
        });

        eachNode(scope, '[data-i18n-aria-label]', function (el) {
            var key = el.getAttribute('data-i18n-aria-label');
            if (key) el.setAttribute('aria-label', t(key));
        });

        try {
            var hostDoc = scope.getElementById ? scope : document;
            var sendBtn = hostDoc.getElementById('send-btn');
            if (sendBtn && !sendBtn.classList.contains('processing')) {
                sendBtn.textContent = t('input.send');
            }
        } catch (_) { /* ignore */ }
        // 只改主聊天页标题。主题模板等独立窗格有自己的 document.title，
        // 这里无条件写成 app.title 会把「PPT 主题模板」盖掉。
        try {
            if ((!root || root === document) && document.getElementById('send-btn')) {
                document.title = t('app.title');
            }
        } catch (_) { /* ignore */ }
    }

    var OCR_MESSAGE_KEYS = {
        '未知引擎类型': 'ocr.status.unknownEngine',
        '未检测到 Python': 'ocr.status.noPython',
        'RapidOCR 已安装并可用': 'ocr.status.rapidOk',
        'RapidOCR 不可用，请检查 Python 环境': 'ocr.status.rapidNeedPy',
        'RapidOCR 不可用': 'ocr.status.rapidBad',
        '云端 OCR': 'ocr.status.cloudName',
        '云端 OCR 引擎': 'ocr.status.cloudName',
        '云端 OCR 已配置': 'ocr.status.cloudReady',
        '云端 OCR 未配置，请在设置中填写 API Key': 'ocr.status.cloudNeedKey',
        '引擎就绪': 'ocr.status.engineReady',
        '检测失败': 'ocr.status.fail',
        'OCR 服务未启动': 'ocr.status.hostDown',
        '就绪': 'ocr.status.ready',
        '不可用': 'ocr.status.unavailable',
        '需要配置': 'ocr.status.needConfig',
        '公式识别已就绪': 'ocr.formulaReady',
        'PaddleOCR 已安装并可用': 'ocr.paddle.ok',
        'PaddleOCR 未安装': 'ocr.paddle.notInstalled',
        'PaddleOCR 已安装，但模型未下载（请重新安装以触发模型下载）': 'ocr.status.paddleNoModel',
        '无法检测状态': 'ocr.paddle.cannotCheck',
        '无法检测状态，等待初始化...': 'ocr.status.waitInit',
        '等待初始化...': 'ocr.status.waitInit',
        '正在检测...': 'ocr.status.checking'
    };

    function localizeOcrMessage(msg, vars) {
        if (msg == null || msg === '') return '';
        var text = String(msg);
        if (OCR_MESSAGE_KEYS[text]) return t(OCR_MESSAGE_KEYS[text], vars);
        var installed = /^PaddleOCR 已安装 \((.+)\)$/.exec(text);
        if (installed) return t('ocr.paddle.installed', { name: installed[1] });
        var missing = /^PaddleOCR 未安装（已找到 (\d+) 个 Python，均不可用）$/.exec(text);
        if (missing) return t('ocr.paddle.missingCount', { n: missing[1] });
        return text;
    }

    var api = {
        t: t,
        localizeOcrMessage: localizeOcrMessage,
        setLocale: setLocale,
        getLocale: getLocale,
        toggleLocale: toggleLocale,
        apply: apply,
        applyExternalLocale: applyExternalLocale,
        has: has,
        missingKeys: missingKeys,
        addLocaleChangeListener: addLocaleChangeListener,
        removeLocaleChangeListener: removeLocaleChangeListener,
        normalizeLocale: normalizeLocale,
        resolveLocale: resolveLocale,
        localeMeta: localeMeta,
        localeCatalog: localeCatalog,
        nextLocale: nextLocale,
        locales: function () { return LOCALES.slice(); },
        strings: STRINGS
    };

    window.App = window.App || {};
    window.App.i18n = api;
    window.t = t;

    /**
     * 语言同步消息总线：setLocale 的 broadcast 会把新语言 postMessage 出去，
     * 这里负责接收（父窗口 ↔ iframe ↔ window.open 的独立窗口）。
     * 收到后本地切换但不再回广播（broadcast:false），避免父子之间来回抖动。
     */
    function bindLocaleMessageBus() {
        try {
            window.addEventListener('message', function (e) {
                var d = e && e.data;
                if (!d || d.type !== LOCALE_MSG) return;
                var incoming = resolveLocale(d.locale);
                if (!incoming || incoming === _locale) return;
                setLocale(incoming, {
                    apply: true,
                    reloadWelcome: true,
                    reloadHistory: true,
                    broadcast: false
                });
            });
        } catch (_) { /* ignore */ }
    }

    function boot() {
        bindLocaleMessageBus();
        try {
            if (window.App && App.config) App.config.uiLocale = _locale;
        } catch (_) { /* ignore */ }
        try {
            if (document.documentElement) document.documentElement.lang = localeMeta(_locale).htmlLang;
        } catch (_) { /* ignore */ }
        // 跟随宿主语言文件，且启动时绝不把本页 localStorage 反写回去
        // （独立 WebView 的旧语言会覆盖用户刚刚的切换）。
        bindExternalLocaleBus();
        pullLocaleFromHostIfUnset();
        apply();
    }
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', boot);
    } else {
        boot();
    }
})();
