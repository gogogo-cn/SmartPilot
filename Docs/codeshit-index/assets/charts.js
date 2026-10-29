// codeshit-index 图表脚本
(function () {
  var style = getComputedStyle(document.documentElement);
  var accent = style.getPropertyValue('--accent').trim();
  var accent2 = style.getPropertyValue('--accent2').trim();
  var ink = style.getPropertyValue('--ink').trim();
  var muted = style.getPropertyValue('--muted').trim();
  var rule = style.getPropertyValue('--rule').trim();
  var bg2 = style.getPropertyValue('--bg2').trim();

  // --- 雷达图：六维评分 ---
  var radar = echarts.init(document.getElementById('chart-radar'), null, { renderer: 'svg' });
  radar.setOption({
    title: { text: '六维技术债评分', left: 'center', textStyle: { color: ink, fontSize: 13, fontWeight: 600 } },
    animation: false,
    tooltip: { trigger: 'item', appendToBody: true, formatter: function (p) {
      return p.name + '：' + p.value + ' / 5';
    } },
    radar: {
      indicator: [
        { name: 'God文件/巨型类', max: 5 },
        { name: '日志体系冗余', max: 5 },
        { name: '错误静默失败', max: 5 },
        { name: '硬编码/配置散落', max: 5 },
        { name: '新旧架构并存', max: 5 },
        { name: '配置/状态冲突', max: 5 }
      ],
      shape: 'polygon', radius: '62%',
      splitNumber: 5,
      axisName: { color: ink, fontSize: 11 },
      splitLine: { lineStyle: { color: rule } },
      splitArea: { areaStyle: { color: [bg2, 'transparent'] } },
      axisLine: { lineStyle: { color: rule } }
    },
    series: [{
      type: 'radar',
      data: [{ value: [4.5, 4.0, 4.5, 4.0, 4.5, 4.0], name: '六维评分' }],
      symbolSize: 4,
      lineStyle: { color: accent, width: 2 },
      itemStyle: { color: accent },
      areaStyle: { color: accent, opacity: 0.18 }
    }]
  });
  window.addEventListener('resize', function () { radar.resize(); });

  // --- 横向条形：超大文件 TOP10 ---
  var files = [
    'ThinkingEngine.vb', 'WorkflowEngine.vb', 'app-editor.js', 'app-file-batch.js',
    'tool-registry.js', 'local-bridge-server.js', 'PptGenerationService.vb',
    'skill-pipeline.js', 'python-pipeline.js', 'chat-renderer.js'
  ];
  var lines = [1836, 1892, 2562, 2602, 2646, 2732, 2695, 2894, 2904, 3397];
  var bar = echarts.init(document.getElementById('chart-god'), null, { renderer: 'svg' });
  bar.setOption({
    title: { text: '超大文件 TOP10（行数）', left: 'center', textStyle: { color: ink, fontSize: 13, fontWeight: 600 } },
    animation: false,
    tooltip: { trigger: 'axis', appendToBody: true, axisPointer: { type: 'shadow' } },
    grid: { left: 8, right: 24, top: 40, bottom: 8, containLabel: true },
    xAxis: { type: 'value', axisLabel: { color: muted, fontSize: 11 }, splitLine: { lineStyle: { color: rule } } },
    yAxis: { type: 'category', data: files, axisLabel: { color: ink, fontSize: 11 }, axisLine: { lineStyle: { color: rule } } },
    series: [{
      type: 'bar',
      data: lines.map(function (v, i) { return { value: v, itemStyle: { color: i < 4 ? accent2 : accent } }; }),
      barWidth: '55%',
      label: { show: true, position: 'right', color: muted, fontSize: 11 }
    }]
  });
  window.addEventListener('resize', function () { bar.resize(); });
})();