/**
 * ============================================================
 * chat-v2/pipelines/python-pipeline.js — Python 数据分析流程
 * ============================================================
 * 职责：
 *   - 处理 Python 数据分析任务
 *   - 使用预置分析模板，LLM 只负责选择模板+填充参数
 *   - 支持多模板并行执行（一次请求同时运行多个分析）
 *   - 用户明确要求"写 Python 脚本"时才走 AI 自由生成
 *
 * 设计原则：
 *   - 预置模板经过测试，不会语法错误
 *   - LLM 输出结构化 JSON（模板选择+参数），不直接写代码
 *   - 模板处理边界情况（空数据、非数值列、缺失值）
 *   - 执行失败有清晰错误提示，不阻塞核心流程
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};
    if (!App.ChatV2.pipelines) App.ChatV2.pipelines = {};

    let ChatLog = window.ChatLog || console;

    // ============================================================
    // 预置分析模板
    // ============================================================
    // 每个模板是 generate(params) 函数，返回可执行的 Python 脚本字符串
    // 模板代码已处理边界情况，无需 AI 关心防御性编程
    // 注意：JS 中生成 Python 代码时，变量替换用 JS 字符串拼接，
    //       Python 代码内的 f-string 转义为普通 format 或字符串拼接
    // ============================================================
    // 共享 Python 片段：日期解析（文本日期 + 8位YYYYMMDD + Excel序列号）。
    // trend_analysis 与 linear_regression 复用，避免注入脚本里 parse_dt 重复漂移。
    // 注意：仅用于注入产物的脚本字符串，非源码层共享逻辑。
    const PY_PARSE_DT_DEF = [
        "    def parse_dt(s):",
        "        s = str(s).strip()",
        "        for fmt in ('%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d', '%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S'):",
        "            try:",
        "                return datetime.strptime(s, fmt)",
        "            except ValueError:",
        "                pass",
        "        if s.isdigit() and len(s) == 8:",
        "            try:",
        "                return datetime.strptime(s, '%Y%m%d')",
        "            except ValueError:",
        "                pass",
        "        try:",
        "            n = float(s)",
        "            if 1 <= n <= 100000:",
        "                return datetime(1899, 12, 30) + timedelta(days=n)",
        "        except ValueError:",
        "            pass",
        "        return None"
    ];

    let TEMPLATES = {
        descriptive_stats: {
            name: '描述统计',
            description: '计算数值列的均值、中位数、标准差、最小值、最大值、四分位数',
            params: {
                columns: { type: 'string_array', description: '要分析的列名列表', required: true }
            },
            generate: function(params) {
                let cols = JSON.stringify(params.columns || []);
                return [
                    "import json, statistics",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list) or len(data) == 0:",
                    "    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域（如 A1:G7）后重试'}",
                    "else:",
                    "    target_cols = " + cols,
                    "    headers = list(data[0].keys()) if isinstance(data[0], dict) else []",
                    "    missing_cols = [c for c in target_cols if c not in headers]",
                    "    results = {}",
                    "    skipped_cols = []",
                    "    for col in target_cols:",
                    "        if col not in headers:",
                    "            continue",
                    "        vals = []",
                    "        for row in data:",
                    "            try:",
                    "                v = row.get(col) if isinstance(row, dict) else None",
                    "                if v is not None and str(v).strip():",
                    "                    vals.append(float(v))",
                    "            except (ValueError, TypeError):",
                    "                pass",
                    "        if vals:",
                    "            n = len(vals)",
                    "            mean = sum(vals) / n",
                    "            s = sorted(vals)",
                    "            median = s[n//2] if n % 2 else (s[n//2-1] + s[n//2]) / 2",
                    "            variance = sum((x - mean)**2 for x in vals) / (n - 1) if n > 1 else 0",
                    "            results[col] = {",
                    "                'count': n,",
                    "                'mean': round(mean, 4),",
                    "                'median': round(median, 4),",
                    "                'stdev': round(variance**0.5, 4),",
                    "                'min': round(min(vals), 4),",
                    "                'max': round(max(vals), 4),",
                    "                'q1': round(s[n//4], 4),",
                    "                'q3': round(s[3*n//4], 4)",
                    "            }",
                    "        else:",
                    "            skipped_cols.append(col)",
                    "    if not results:",
                    "        detail = []",
                    "        if missing_cols:",
                    "            detail.append('列不存在: ' + ', '.join(missing_cols))",
                    "        if skipped_cols:",
                    "            detail.append('无有效数值: ' + ', '.join(skipped_cols))",
                    "        if not detail:",
                    "            detail.append('数据为空或非数值')",
                    "        output = {'success': False, 'message': '所选列均无法计算描述统计（' + '；'.join(detail) + '）。请检查列名是否与表头一致、数据是否为数值。'}",
                    "    else:",
                    "        print(json.dumps(results, ensure_ascii=False, indent=2))",
                    "        output = {'success': True, 'summary': '完成' + str(len(results)) + '列的描述统计', 'result': results}"
                ].join('\n');
            }
        },

        normalize: {
            name: '数据归一化',
            description: '对数值列做 Min-Max 归一化或 Z-Score 标准化，输出与原始数据行对齐的归一化结果（新增 *_minmax / *_zscore 列），可一键插入到 Excel。',
            params: {
                columns: { type: 'string_array', description: '要归一化的数值列名列表', required: true },
                method: { type: 'string', description: "归一化方法: 'minmax'(默认, 缩放到[0,1]) 或 'zscore'(标准化为均值0标准差1)", required: false, default: 'minmax' }
            },
            generate: function(params) {
                let cols = JSON.stringify(params.columns || []);
                let method = JSON.stringify(params.method || 'minmax');
                return [
                    "import json, math",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list) or len(data) == 0:",
                    "    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域（如 A1:G7）后重试'}",
                    "else:",
                    "    target_cols = " + cols,
                    "    method = " + method,
                    "    headers = list(data[0].keys()) if isinstance(data[0], dict) else []",
                    "    missing_cols = [c for c in target_cols if c not in headers]",
                    "    series = {}",
                    "    skipped_cols = []",
                    "    for col in target_cols:",
                    "        if col not in headers:",
                    "            continue",
                    "        vals = []",
                    "        for row in data:",
                    "            try:",
                    "                v = float(row.get(col)) if isinstance(row, dict) else None",
                    "                if v is not None:",
                    "                    vals.append(v)",
                    "            except (ValueError, TypeError):",
                    "                pass",
                    "        if vals:",
                    "            series[col] = vals",
                    "        else:",
                    "            skipped_cols.append(col)",
                    "    if not series:",
                    "        detail = []",
                    "        if missing_cols:",
                    "            detail.append('列不存在: ' + ', '.join(missing_cols))",
                    "        if skipped_cols:",
                    "            detail.append('无有效数值: ' + ', '.join(skipped_cols))",
                    "        if not detail:",
                    "            detail.append('数据为空或非数值')",
                    "        output = {'success': False, 'message': '所选列均无法归一化（' + '；'.join(detail) + '）。请检查列名是否与表头一致、数据是否为数值。'}",
                    "    else:",
                    "        suffix = '_minmax' if method == 'minmax' else '_zscore'",
                    "        out = []",
                    "        for row in data:",
                    "            newrow = dict(row) if isinstance(row, dict) else {}",
                    "            for col in series:",
                    "                try:",
                    "                    v = float(row.get(col)) if isinstance(row, dict) else None",
                    "                except (ValueError, TypeError):",
                    "                    v = None",
                    "                if v is None:",
                    "                    newrow[col + suffix] = None",
                    "                    continue",
                    "                arr = series[col]",
                    "                if method == 'minmax':",
                    "                    mn, mx = min(arr), max(arr)",
                    "                    nv = (v - mn) / (mx - mn) if mx > mn else 0.0",
                    "                else:",
                    "                    n = len(arr)",
                    "                    mean = sum(arr) / n",
                    "                    var = sum((x - mean) ** 2 for x in arr) / (n - 1) if n > 1 else 0",
                    "                    std = var ** 0.5",
                    "                    nv = (v - mean) / std if std > 0 else 0.0",
                    "                newrow[col + suffix] = round(nv, 4)",
                    "            out.append(newrow)",
                    "        print(json.dumps(out, ensure_ascii=False, indent=2))",
                    "        output = {'success': True, 'summary': '完成' + str(len(series)) + '列归一化（方法=' + method + '），共' + str(len(out)) + '行（已新增' + suffix + '列）', 'result': out}"
                ].join('\n');
            }
        },

        trend_analysis: {
            name: '趋势分析',
            description: '【时间趋势】对单个数值列进行趋势分析（上升/下降/平稳），计算斜率并预测未来值。适用于"分析XX的趋势"、"XX随时间变化"等问题。参数需一个数值列名；若数据含日期/时间列，建议同时指定 time_column，趋势将按真实时间轴分析并给出预测日期',
            params: {
                value_column: { type: 'string', description: '要分析的数值列名', required: true },
                time_column: { type: 'string', description: '时间列名（日期/时间，可选；提供后按时间轴分析）', required: false, default: '' },
                forecast_steps: { type: 'number', description: '预测期数（默认5）', required: false, default: 5 }
            },
            generate: function(params) {
                let col = JSON.stringify(params.value_column || '');
                let tcol = JSON.stringify(params.time_column || '');
                let steps = parseInt(params.forecast_steps, 10) || 5;
                return [
                    "import json, math",
                    "from datetime import datetime, timedelta",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list):",
                    "    output = {'success': False, 'message': '无可用数据'}",
                    "else:",
                    "    col = " + col,
                    "    tcol = " + tcol,
                    "    steps = " + steps,
                ].concat(PY_PARSE_DT_DEF).concat([
                    "    vals = []; tvals = []; raw_x = []",
                    "    for row in data:",
                    "        if not isinstance(row, dict):",
                    "            continue",
                    "        try:",
                    "            v = row.get(col)",
                    "            if v is None or str(v).strip() == '':",
                    "                continue",
                    "            fv = float(str(v).replace(',', '').replace('¥', ''))",
                    "        except (ValueError, TypeError):",
                    "            continue",
                    "        vals.append(fv)",
                    "        tv = row.get(tcol) if tcol else None",
                    "        if tv is not None and str(tv).strip() != '':",
                    "            tvals.append(parse_dt(tv)); raw_x.append(str(tv))",
                    "        else:",
                    "            tvals.append(None); raw_x.append('')",
                    "    n = len(vals)",
                    "    if n < 3:",
                    "        output = {'success': False, 'message': '列有效数据不足3个，无法分析趋势'}",
                    "    else:",
                    "        has_time = bool(tcol) and all(t is not None for t in tvals)",
                    "        if has_time:",
                    "            epoch = min(tvals)",
                    "            X = [(t - epoch).days for t in tvals]  # 仅保留天数整数部分(忽略时间)，按日粒度计算",
                    "        else:",
                    "            X = [float(i) for i in range(n)]",
                    "        x_mean = sum(X) / n",
                    "        y_mean = sum(vals) / n",
                    "        num = sum((X[i] - x_mean) * (vals[i] - y_mean) for i in range(n))",
                    "        den = sum((X[i] - x_mean) ** 2 for i in range(n))",
                    "        slope = num / den if den != 0 else 0",
                    "        intercept = y_mean - slope * x_mean",
                    "        direction = '上升' if slope > 0.001 else ('下降' if slope < -0.001 else '平稳')",
                    "        step_x = (X[-1] - X[0]) / (n - 1) if (n > 1 and X[-1] != X[0]) else 1",
                    "        fx = [X[-1] + step_x * (i + 1) for i in range(steps)]",
                    "        forecast = [round(slope * f + intercept, 2) for f in fx]",
                    "        forecast_x = []",
                    "        if has_time:",
                    "            for f in fx:",
                    "                d = epoch + timedelta(days=int(round(f)))  # 整日偏移，统一输出日期格式，不带时间",
                    "                forecast_x.append(d.strftime('%Y-%m-%d'))",
                    "        else:",
                    "            forecast_x = [int(round(f)) for f in fx]",
                    "        # 历史序列：用于双序列折线图（历史真实值 + 未来预测值同图）",
                    "        historical_values = [round(v, 4) for v in vals]",
                    "        if has_time:",
                    "            historical_x = [tvals[i].strftime('%Y-%m-%d') for i in range(n)]",
                    "        else:",
                    "            historical_x = [i for i in range(n)]",
                    "        result = {",
                    "            'column': col,",
                    "            'data_points': n,",
                    "            'slope': round(slope, 4),",
                    "            'intercept': round(intercept, 4),",
                    "            'direction': direction,",
                    "            'forecast': forecast,",
                    "            'has_time_axis': has_time,",
                    "            'time_column': tcol if has_time else '',",
                    "            'forecast_x': forecast_x,",
                    "            'historical_x': historical_x,",
                    "            'historical_values': historical_values",
                    "        }",
                    "        print(json.dumps(result, ensure_ascii=False, indent=2))",
                    "        output = {'success': True, 'summary': '趋势分析完成(' + ('按时间轴' if has_time else '按行序') + ')，方向:' + direction, 'result': result}"
                ]).join('\n');
            }
        },

        correlation: {
            name: '相关性分析',
            description: '计算数值列之间的 Pearson 相关系数矩阵',
            params: {
                columns: { type: 'string_array', description: '要分析的列名列表（至少2列）', required: true }
            },
            generate: function(params) {
                let cols = JSON.stringify(params.columns || []);
                return [
                    "import json, math",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list):",
                    "    output = {'success': False, 'message': '无可用数据'}",
                    "else:",
                    "    target_cols = " + cols,
                    "    if len(target_cols) < 2:",
                    "        output = {'success': False, 'message': '需要至少2列做相关性分析'}",
                    "    else:",
                    "        col_vals = {}",
                    "        for col in target_cols:",
                    "            vals = []",
                    "            for row in data:",
                    "                try:",
                    "                    v = row.get(col) if isinstance(row, dict) else None",
                    "                    if v is not None and str(v).strip():",
                    "                        vals.append(float(v))",
                    "                except: pass",
                    "            if len(vals) >= 3:",
                    "                col_vals[col] = vals",
                    "        if len(col_vals) < 2:",
                    "            output = {'success': False, 'message': '有效数值列不足2列'}",
                    "        else:",
                    "            corr_matrix = {}",
                    "            cols_list = list(col_vals.keys())",
                    "            for i, c1 in enumerate(cols_list):",
                    "                corr_matrix[c1] = {}",
                    "                for j, c2 in enumerate(cols_list):",
                    "                    if i == j:",
                    "                        corr_matrix[c1][c2] = 1.0",
                    "                    else:",
                    "                        v1, v2 = col_vals[c1], col_vals[c2]",
                    "                        min_len = min(len(v1), len(v2))",
                    "                        v1, v2 = v1[:min_len], v2[:min_len]",
                    "                        n = len(v1)",
                    "                        m1, m2 = sum(v1)/n, sum(v2)/n",
                    "                        num = sum((v1[k]-m1)*(v2[k]-m2) for k in range(n))",
                    "                        d1 = math.sqrt(sum((v1[k]-m1)**2 for k in range(n)))",
                    "                        d2 = math.sqrt(sum((v2[k]-m2)**2 for k in range(n)))",
                    "                        r = num/(d1*d2) if d1*d2 != 0 else 0",
                    "                        corr_matrix[c1][c2] = round(r, 4)",
                    "            print(json.dumps(corr_matrix, ensure_ascii=False, indent=2))",
                    "            output = {'success': True, 'summary': '完成' + str(len(cols_list)) + '列的相关性分析', 'result': corr_matrix}"
                ].join('\n');
            }
        },

        anomaly_detection: {
            name: '异常值检测',
            description: '使用 Z-Score 方法检测数值列中的异常值（默认阈值=3）',
            params: {
                columns: { type: 'string_array', description: '要检测的列名列表', required: true },
                threshold: { type: 'number', description: 'Z-Score 阈值（默认3）', required: false, default: 3 }
            },
            generate: function(params) {
                let cols = JSON.stringify(params.columns || []);
                let thresh = params.threshold || 3;
                return [
                    "import json, math",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list):",
                    "    output = {'success': False, 'message': '无可用数据'}",
                    "else:",
                    "    target_cols = " + cols,
                    "    threshold = " + thresh,
                    "    results = {}",
                    "    for col in target_cols:",
                    "        vals = []",
                    "        for row in data:",
                    "            try:",
                    "                v = row.get(col) if isinstance(row, dict) else None",
                    "                if v is not None and str(v).strip():",
                    "                    vals.append((float(v), row))",
                    "            except: pass",
                    "        if len(vals) < 3:",
                    "            continue",
                    "        nums = [v[0] for v in vals]",
                    "        n = len(nums)",
                    "        mean = sum(nums)/n",
                    "        var = sum((x-mean)**2 for x in nums)/(n-1) if n>1 else 0",
                    "        std = math.sqrt(var) if var > 0 else 1",
                    "        anomalies = []",
                    "        for v, row in vals:",
                    "            z = (v - mean)/std",
                    "            if abs(z) > threshold:",
                    "                row_label = str(row.get(list(row.keys())[0], '')) if isinstance(row, dict) else ''",
                    "                anomalies.append({'value': round(v,4), 'z_score': round(z,4), 'label': row_label})",
                    "        results[col] = {",
                    "            'total': n,",
                    "            'anomaly_count': len(anomalies),",
                    "            'anomaly_rate': round(len(anomalies)/n*100, 2),",
                    "            'anomalies': anomalies[:20]",
                    "        }",
                    "    total_anomalies = sum(r['anomaly_count'] for r in results.values())",
                    "    print(json.dumps(results, ensure_ascii=False, indent=2))",
                    "    output = {'success': True, 'summary': '检测完成，共发现' + str(total_anomalies) + '个异常值', 'result': results}"
                ].join('\n');
            }
        },

        data_distribution: {
            name: '数据分布',
            description: '计算数值列的频率分布、分箱统计',
            params: {
                columns: { type: 'string_array', description: '要分析的列名列表', required: true },
                bins: { type: 'number', description: '分箱数（默认10）', required: false, default: 10 }
            },
            generate: function(params) {
                let cols = JSON.stringify(params.columns || []);
                let bins = params.bins || 10;
                return [
                    "import json, math",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list):",
                    "    output = {'success': False, 'message': '无可用数据'}",
                    "else:",
                    "    target_cols = " + cols,
                    "    num_bins = " + bins,
                    "    results = {}",
                    "    for col in target_cols:",
                    "        vals = []",
                    "        for row in data:",
                    "            try:",
                    "                v = row.get(col) if isinstance(row, dict) else None",
                    "                if v is not None and str(v).strip():",
                    "                    vals.append(float(v))",
                    "            except: pass",
                    "        if not vals:",
                    "            continue",
                    "        min_v, max_v = min(vals), max(vals)",
                    "        bin_width = (max_v - min_v) / num_bins if max_v > min_v else 1",
                    "        bins_list = [round(min_v + i*bin_width, 2) for i in range(num_bins+1)]",
                    "        counts = [0]*num_bins",
                    "        for v in vals:",
                    "            idx = min(int((v - min_v) / bin_width), num_bins-1) if bin_width > 0 else 0",
                    "            counts[idx] += 1",
                    "        distribution = []",
                    "        for i in range(num_bins):",
                    "            distribution.append({",
                    "                'range': str(bins_list[i]) + '-' + str(bins_list[i+1]),",
                    "                'count': counts[i],",
                    "                'percentage': round(counts[i]/len(vals)*100, 2)",
                    "            })",
                    "        results[col] = {",
                    "            'min': round(min_v, 4),",
                    "            'max': round(max_v, 4),",
                    "            'bins': num_bins,",
                    "            'distribution': distribution",
                    "        }",
                    "    print(json.dumps(results, ensure_ascii=False, indent=2))",
                    "    output = {'success': True, 'summary': '完成' + str(len(results)) + '列的数据分布分析', 'result': results}"
                ].join('\n');
            }
        },

        group_aggregation: {
            name: '分组聚合',
            description: '【按分类分组统计】按某列分组，对数值列做求和/平均/计数。需要同时指定分组列(如部门、地区)和数值列(如销售额、利润)。注意：如果只是分析单个数值列的趋势，用 trend_analysis 而非 group_aggregation',
            params: {
                group_column: { type: 'string', description: '分组依据的列名', required: true },
                value_column: { type: 'string', description: '要聚合的数值列名', required: true },
                aggregation: { type: 'string', description: '聚合方式: sum/avg/count/min/max（默认avg）', required: false, default: 'avg' }
            },
            generate: function(params) {
                let groupCol = JSON.stringify(params.group_column || '');
                let valCol = JSON.stringify(params.value_column || '');
                let agg = params.aggregation || 'avg';
                return [
                    "import json",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list):",
                    "    output = {'success': False, 'message': '无可用数据'}",
                    "else:",
                    "    group_col = " + groupCol,
                    "    val_col = " + valCol,
                    "    agg = '" + agg + "'",
                    "    groups = {}",
                    "    for row in data:",
                    "        if not isinstance(row, dict): continue",
                    "        g = row.get(group_col, '')",
                    "        try:",
                    "            v = float(row.get(val_col, 0) or 0)",
                    "        except: continue",
                    "        if g not in groups:",
                    "            groups[g] = {'vals': [], 'count': 0}",
                    "        groups[g]['vals'].append(v)",
                    "        groups[g]['count'] += 1",
                    "    result = {}",
                    "    for g, info in groups.items():",
                    "        vals = info['vals']",
                    "        n = len(vals)",
                    "        if n == 0: continue",
                    "        if agg == 'sum':",
                    "            result[g] = {'value': round(sum(vals), 4), 'count': n, 'agg': 'sum'}",
                    "        elif agg == 'min':",
                    "            result[g] = {'value': round(min(vals), 4), 'count': n, 'agg': 'min'}",
                    "        elif agg == 'max':",
                    "            result[g] = {'value': round(max(vals), 4), 'count': n, 'agg': 'max'}",
                    "        elif agg == 'count':",
                    "            result[g] = {'value': n, 'count': n, 'agg': 'count'}",
                    "        else:",
                    "            result[g] = {'value': round(sum(vals)/n, 4), 'count': n, 'agg': 'avg'}",
                    "    print(json.dumps(result, ensure_ascii=False, indent=2))",
                    "    output = {'success': True, 'summary': '按' + group_col + '分组，对' + val_col + '做' + agg + '聚合，共' + str(len(result)) + '组', 'result': result}"
                ].join('\n');
            }
        },
        // ============ 扩展模板示例：线性回归分析（验证"新增一个活=新增一个模板"的扩展机制） ============
        linear_regression: {
            name: '线性回归分析',
            description: '对两个数值列做一元线性回归，输出斜率、截距、决定系数 R² 及预测点',
            params: {
                xColumn: { type: 'string', description: '自变量列名（数值列，如成本）', required: true },
                yColumn: { type: 'string', description: '因变量列名（数值列，如利润）', required: true },
                forecastPoints: { type: 'int', description: '预测点数（默认5）', default: 5 }
            },
            generate: function(params) {
                let xcol = JSON.stringify(params.xColumn || '');
                let ycol = JSON.stringify(params.yColumn || '');
                let fpt = Math.max(2, parseInt(params.forecastPoints, 10) || 5);
                return [
                    "import json",
                    "from datetime import datetime, timedelta",
                    "data = json.loads(json.dumps(input_data)) if isinstance(input_data, (list, dict)) else []",
                    "if not data or not isinstance(data, list) or len(data) == 0:",
                    "    output = {'success': False, 'message': '未读取到数据行，请先选中包含表头和数据的区域（如 A1:G7）后重试'}",
                    "else:",
                    "    xcol = " + xcol,
                    "    ycol = " + ycol,
                ].concat(PY_PARSE_DT_DEF).concat([
                    "    def to_num(v):",
                    "        if v is None: return None",
                    "        try:",
                    "            return float(str(v).replace(',', '').replace('%', '').replace('¥', ''))",
                    "        except ValueError:",
                    "            pass",
                    "        dt = parse_dt(v)",
                    "        if dt is not None:",
                    "            return (dt - datetime(1899, 12, 30)).days + (dt - datetime(1899, 12, 30)).seconds / 86400.0",
                    "        return None",
                    "    xs = []; ys = []",
                    "    for row in data:",
                    "        xv = row.get(xcol); yv = row.get(ycol)",
                    "        if (xv is None or yv is None or str(xv).strip() == '' or str(yv).strip() == ''):",
                    "            continue",
                    "        xf = to_num(xv); yf = to_num(yv)",
                    "        if xf is None or yf is None:",
                    "            continue",
                    "        xs.append(xf); ys.append(yf)",
                    "    n = len(xs)",
                    "    if n < 2:",
                    "        output = {'success': False, 'message': '有效数值点不足(<2)，无法进行线性回归，请检查所选的列是否含足够数值'}",
                    "    else:",
                    "        mx = sum(xs) / n; my = sum(ys) / n",
                    "        sxx = sum((x - mx) ** 2 for x in xs)",
                    "        sxy = sum((x - mx) * (y - my) for x, y in zip(xs, ys))",
                    "        slope = (sxy / sxx) if sxx != 0 else 0.0",
                    "        intercept = my - slope * mx",
                    "        ss_res = sum((y - (slope * x + intercept)) ** 2 for x, y in zip(xs, ys))",
                    "        ss_tot = sum((y - my) ** 2 for y in ys)",
                    "        r2 = (1 - ss_res / ss_tot) if ss_tot != 0 else 0.0",
                    "        fps = " + fpt,
                    "        xmin = min(xs); xmax = max(xs)",
                    "        step = ((xmax - xmin) / (fps - 1)) if (fps > 1 and xmax != xmin) else 1.0",
                    "        forecast = [{'x': round(xmin + step * i, 4), 'y': round(slope * (xmin + step * i) + intercept, 4)} for i in range(fps)]",
                    "        eq = 'y = {0:.4f}*x {1:+.4f}'.format(slope, intercept)",
                    "        output = {'success': True, 'summary': '对 ' + str(xcol) + ' 与 ' + str(ycol) + ' 做一元线性回归: ' + eq + ', R²=' + ('{0:.4f}'.format(r2)) + ', 样本n=' + str(n), 'result': {'n': n, 'slope': round(slope, 6), 'intercept': round(intercept, 6), 'r2': round(r2, 4), 'equation': eq, 'forecast': forecast}}",
                    "    print(json.dumps(output, ensure_ascii=False, indent=2))"
                ]).join('\n');
            }
        }
    };

    // ============================================================
    // 参数校验层（纯函数，供执行路径调用与 node:test 单测）
    // 目标：把"静默替换坏列"升级为"显式校验 + 可用列提示"，杜绝运行在不存在的列上造成误导
    // ============================================================

    // 轻量字符重合度（用于给坏列推荐近义候选列），中文列名场景足够
    function _colSimilarity(a, b) {
        if (!a || !b) return 0;
        if (a === b) return 1;
        let common = 0;
        const sbSet = {};
        for (let i = 0; i < b.length; i++) sbSet[b[i]] = true;
        for (let j = 0; j < a.length; j++) if (sbSet[a[j]]) common++;
        return common / Math.max(a.length, b.length, 1);
    }

    // 给每个无效列生成"所见即可用"的提示：近义候选 + 可用列清单
    function buildColumnSuggestions(invalidColumns, availableHeaders, numericHeaders) {
        availableHeaders = availableHeaders || [];
        numericHeaders = numericHeaders || availableHeaders.slice();
        if (!invalidColumns || invalidColumns.length === 0) return '';
        return invalidColumns.map(function (ic) {
            let target = (ic && ic.name != null) ? String(ic.name) : '';
            let best = null, bs = -1;
            for (let hi = 0; hi < availableHeaders.length; hi++) {
                let s = _colSimilarity(String(availableHeaders[hi]).toLowerCase(), target.toLowerCase());
                if (s > bs) { bs = s; best = availableHeaders[hi]; }
            }
            let near = (best && bs >= 0.5 && best !== target) ? '，近义候选: ' + best : '';
            let pool = (ic && ic.numeric) ? numericHeaders : availableHeaders;
            return '「' + (target || '(空)') + '」不在可用列中' + near + '（可用' + ((ic && ic.numeric) ? '数值 ' : '') + '列: ' + ((pool && pool.length) ? pool.join('、') : '无') + '）';
        }).join('；');
    }

    // 参数 schema 合规 + 列存在性/数值性校验（纯函数，不修改 params，只报告）
    function validateTemplateParams(templateId, params, availableHeaders, numericHeaders) {
        availableHeaders = availableHeaders || [];
        numericHeaders = numericHeaders || availableHeaders.slice();
        const lowerH = availableHeaders.map(function (h) { return String(h).toLowerCase(); });
        const lowerN = numericHeaders.map(function (h) { return String(h).toLowerCase(); });
        const tpl = TEMPLATES[templateId];
        const issues = [];
        const missingRequired = [];
        const invalidColumns = [];
        const columnKeysRegex = /(value_column|group_column|xColumn|yColumn|time_column)/i;
        const columnArrayKeysRegex = /columns/i;

        if (tpl) {
            for (const pname in tpl.params) {
                const spec = tpl.params[pname];
                if (spec.required && (params[pname] === undefined || params[pname] === null || params[pname] === '')) {
                    missingRequired.push(pname);
                    issues.push('缺必填参数: ' + pname);
                }
                const needNum = /(value_column|xColumn|yColumn)/i.test(pname);
                if (spec.type === 'string' && columnKeysRegex.test(pname)) {
                    const col = params[pname];
                    if (col) {
                        const idx = lowerH.indexOf(String(col).toLowerCase());
                        if (idx < 0) {
                            invalidColumns.push({ name: col, numeric: needNum });
                        } else if (needNum && lowerN.indexOf(lowerH[idx]) < 0) {
                            invalidColumns.push({ name: col, numeric: needNum, reason: '非数值列' });
                        }
                    }
                } else if (spec.type === 'string_array' && columnArrayKeysRegex.test(pname) && Array.isArray(params[pname])) {
                    for (const c of params[pname]) {
                        if (!c) continue;
                        const idx = lowerH.indexOf(String(c).toLowerCase());
                        if (idx < 0) {
                            invalidColumns.push({ name: c, numeric: true });
                        } else if (lowerN.indexOf(lowerH[idx]) < 0) {
                            invalidColumns.push({ name: c, numeric: true, reason: '非数值列' });
                        }
                    }
                }
            }
        } else {
            issues.push('未知模板: ' + templateId);
        }

        for (let ii = 0; ii < invalidColumns.length; ii++) {
            issues.push('无效列引用: ' + String(invalidColumns[ii].name) + (invalidColumns[ii].reason ? '（' + invalidColumns[ii].reason + '）' : ''));
        }

        const columnSuggestions = invalidColumns.length > 0 ? buildColumnSuggestions(invalidColumns, availableHeaders, numericHeaders) : '';
        return {
            ok: !!tpl && missingRequired.length === 0,
            templateId: templateId,
            missingRequired: missingRequired,
            invalidColumns: invalidColumns,
            columnSuggestions: columnSuggestions,
            issues: issues
        };
    }

    // 列名护栏文案：custom 分支命中坏列时，中止并把可用列清单回传给用户
    function columnGuardMessage(suspectCols, availableHeaders) {
        const s = (Array.isArray(suspectCols) ? suspectCols : []).slice(0, 5)
            .map(function (x) { return '`' + x + '`'; }).join('、');
        const a = (Array.isArray(availableHeaders) ? availableHeaders : []).slice(0, 8).join('、');
        return '自定义脚本引用了不存在的列：' + (s || '(未知占位符)') +
            (a ? '。可用列：' + a + '。' : '。') +
            '请改用可用列重写脚本，或用内置模板（如"异常检测"）更稳定。';
    }

    // 是否强制走 custom（LLM 自由生成）。采用「模板优先」策略：
    // 只要没有强烈表达"自由/自定义写脚本"，带 python 也应交给模板选择器（内置真实列 + 模板列表），
    // 由选择器按意图挑模板；选择器兜底逻辑会保证选不到模板时降级 custom（见 selector 解析失败分支）。
    // 这样避免"用python+分组聚合/异常检测"这类明确活因触发词漏网而误入 custom，也不依赖意图词穷举。
    function shouldForceCustomInput(userMessage) {
        const m = String(userMessage || '').toLowerCase();
        const py = /(写\s*python|生成\s*python|python\s*脚本|python\s*代码|自定义.*分析|用\s*python|\bpython\b)/.test(m);
        if (!py) return false;
        // 仅当强烈表达"自定义/自由/自己写"时才强制 custom，其余一律交给模板选择器
        const strongCustomize = /(自定义|自由发挥|自由脚本|自由处理|自己写|自己写代码|完全自己|随便|随意|任意发挥|不想用模板|不要模板)/.test(m);
        return strongCustomize;
    }

    // 意图→模板确定性绑定：把"分组聚合/异常/趋势/回归/相关/分布/描述统计"等明确活
    // 直接映射到内置模板，跳过 qwen 从模板列表"二次选择"这一步——避免它选模板失败
    // 时在 L1649 附近降级回 custom（从而再次自由生成触发列幻觉）。
    // 顺序偏具体优先，避免"预测趋势"等同时命中多个模板时误分配。
    const TEMPLATE_INTENTS = [
        { id: 'normalize', re: /归一化|标准化|正规化|min-?max|z-?score|z分数|极差法|小数定标/ },
        { id: 'anomaly', re: /异常|离群|越界|极端|不符/ },
        { id: 'linear_regression', re: /线性回归|回归分析|回归|拟合|趋势线/ },
        { id: 'correlation', re: /相关系数|相关性|相关|关联/ },
        { id: 'trend', re: /趋势|走势|时间序列|随时间|预测未来|未来预测/ },
        { id: 'group_aggregation', re: /分组|按.{0,8}(汇总|统计|合计)|汇总|合计/ },
        { id: 'distribution', re: /分布|频数|频率分布|直方/ },
        { id: 'descriptive_stats', re: /描述统计|均值|中位数|标准差|统计量/ },
        // 通用"区域总结/概况"兜底：无更具体意图时确定性走描述统计，避免 LLM 选择器
        // 返回无效 JSON 后降级 custom 自由生成（列名幻觉高发），此规则必须放在最末（最不具体）
        { id: 'descriptive_stats', re: /(?:区域|选区|表格).*(?:总结|概况|总览)|分析.*(?:总结|概况|总览|一下)|给出总结|做.*总结|总结.*(?:区域|选区|表格|数据)|概况|总览/ }
    ];
    function matchTemplateByIntent(userMessage) {
        const m = String(userMessage || '').toLowerCase();
        for (const it of TEMPLATE_INTENTS) {
            if (it.re.test(m)) {
                const res = { id: it.id };
                // 归一化：z-score / z分数 / 明确"标准化为均值0标准差1" → 走 zscore 方法，
                // 其余（归一化 / 标准化 / min-max / 正规化）默认 minmax，避免方法误判。
                if (it.id === 'normalize' && /z-?score|z分数|标准化为.*标准差|均值为\s*0|正态化/.test(m)) {
                    res.params = { method: 'zscore' };
                }
                return res;
            }
        }
        return null;
    }

    // 日期/时间列名识别：即使以 Excel 序列号(数值形态)存储，本质仍是维度列，
    // 不得当作"数值列"做聚合（value_column），也不宜作分组列(group_column)导致每组仅1条。
    const DATE_COL_RE = /(日\s*期|日期|时间|年月|季度|月度|\bdate\b|datetime|-dt\b|\btime\b)/i;

    // 从表头/数据中识别日期时间列（供 trend 模板自动填补 time_column）
    // 依据：列名启发（日期/时间/date/time）+ 前若干行的值形态占比。
    // 支持两种日期形态：文本日期（2023-01-01 等）与 Excel 序列号（纯数字 1~100000）；
    // 序列号列仅当列名带日期语义时才加权，避免误伤销售额等普通数值列。
    function detectDateColumn(headers, rows) {
        headers = headers || [];
        if (!headers.length) return '';
        const nameRe = /日\s*期|日期|时间|date|time|年|季度|month|day|月日/i;
        const textDateish = function (v) {
            if (v === null || v === undefined) return false;
            const s = String(v).trim();
            if (!s) return false;
            if (/^\d{8}$/.test(s)) return true;
            return /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(s) || /^\d{1,2}[-/.]\d{1,2}[-/.]\d{4}/.test(s);
        };
        // Excel 序列号日期的数字形态：纯数字、非 8 位（8 位是 YYYYMMDD，已归 textDateish）、值 1~100000
        const serialNumish = function (v) {
            if (v === null || v === undefined) return false;
            const s = String(v).trim();
            if (!/^\d+(\.\d+)?$/.test(s)) return false;
            const n = parseFloat(s);
            return n >= 1 && n <= 100000 && s.replace('.', '').length < 8;
        };
        let best = null, bestScore = 0;
        for (let hi = 0; hi < headers.length; hi++) {
            const h = headers[hi];
            if (h === null || h === undefined) continue;
            const hasDateName = nameRe.test(String(h));
            let score = hasDateName ? 1.5 : 0;
            if (rows && rows.length) {
                let textCnt = 0, serialCnt = 0, tot = 0;
                const limit = Math.min(rows.length, 15);
                for (let ri = 0; ri < limit; ri++) {
                    const row = rows[ri];
                    const v = Array.isArray(row) ? row[hi] : (row ? row[h] : null);
                    tot++;
                    if (textDateish(v)) textCnt++;
                    else if (serialNumish(v)) serialCnt++;
                }
                if (tot) {
                    score += (textCnt / tot) * 3;
                    // 序列号只在列名带日期语义时计入，普通数值列（如销售额）不会因此被误判
                    if (hasDateName) score += (serialCnt / tot) * 2;
                }
            }
            if (score > bestScore) { bestScore = score; best = h; }
        }
        return (best && bestScore >= 1.5) ? best : '';
    }

    function getTemplateIds() {
        return Object.keys(TEMPLATES);
    }

    function generateTemplate(templateId, params) {
        const tpl = TEMPLATES[templateId];
        if (!tpl) return null;
        return tpl.generate(params || {});
    }

    if (window.App) {
        window.App.PythonTemplateValidator = {
            validateTemplateParams: validateTemplateParams,
            buildColumnSuggestions: buildColumnSuggestions,
            detectDateColumn: detectDateColumn,
            columnGuardMessage: columnGuardMessage,
            shouldForceCustomInput: shouldForceCustomInput,
            matchTemplateByIntent: matchTemplateByIntent,
            buildTemplateChartCfg: buildTemplateChartCfg,
            buildResultChartDiv: buildResultChartDiv,
            buildChartDiv: buildChartDiv,
            templateResultToMarkdown: templateResultToMarkdown,
            getTemplateIds: getTemplateIds,
            generate: generateTemplate
        };
    }

    // ============================================================
    // 模板选择 Prompt
    // ============================================================

    let TEMPLATE_SELECTOR_PROMPT = '你是 SmartPilot 的数据分析助手。根据用户需求，从预置模板中选择最合适的分析方案并填充参数。\n\n' +
        '【可用模板】\n{%TEMPLATES%}\n\n' +
        '【可用数据列】\n{%HEADERS%}\n\n' +
        '【模板选择规则】\n' +
        '- 关键词"趋势"、"走向"、"变化" → trend_analysis（value_column 必填；若数据表含日期/时间列，补 time_column 则按时间轴分析并输出预测日期）\n' +
        '- 关键词"异常"、"离群"、"outlier" → anomaly_detection\n' +
        '- 关键词"统计"、"汇总"、"均值/中位数" → descriptive_stats\n' +
        '- 关键词"相关性"、"关联"、"相关系数" → correlation\n' +
        '- 关键词"回归"、"线性回归"、"拟合"、"预测趋势"、"因果关系" → linear_regression（需指定自变量 xColumn 和因变量 yColumn 两个数值列）\n' +
        '- 关键词"分布"、"频率"、"直方图" → data_distribution\n' +
        '- 关键词"分组"、"按XX统计"、"聚合" → group_aggregation（需要同时指定分组列和数值列）\n' +
        '- 关键词"趋势和异常"、"同时" → 输出多个模板\n' +
        '- 如果用户明确要求"写Python"、"生成代码" → 输出 custom\n\n' +
        '【重要】优先使用模板！只有在所有模板都无法满足需求时才使用 custom。\n' +
        '对于"相关性 + 分组"这类组合需求，应该分别输出 correlation 和 group_aggregation 两个模板，而不是用 custom。\n\n' +
        '【输出格式】仅输出 JSON，不要任何额外文字：\n' +
        '单模板: {"templates": [{"id": "模板ID", "params": {...}}]}\n' +
        '多模板: {"templates": [{"id": "模板ID1", "params": {...}}, {"id": "模板ID2", "params": {...}}]}\n\n' +
        '【参数规则】\n' +
        '- 参数名必须严格匹配模板定义（如 trend_analysis 用 value_column，不是 dimColumn 或 column）\n' +
        '- 列名必须使用"可用数据列"中列出的实际列名，必须逐字完全匹配（大小写、括号、单位全部一致）\n' +
        '- 如果"可用数据列"为空（你不知道真实列名），所有 params 置为空字符串，之后由系统自动选择默认列\n' +
        '- 严禁编造不存在的列名（例如"销售额"、"收入"、"日期"等必须在可用数据列里真的出现）\n' +
        '- 禁止使用占位词如"列名"、"xxx"、"字段名"';


    // 旧 prompt 保留为参考
    // var _OLD_PROMPT = '...';

    // ============================================================
    // 模板列表说明（注入 TEMPLATE_SELECTOR_PROMPT）
    // ============================================================

    function buildTemplateList() {
        let lines = [];
        let entries = Object.entries(TEMPLATES);
        for (let i = 0; i < entries.length; i++) {
            let id = entries[i][0];
            let tpl = entries[i][1];
            lines.push('【' + tpl.name + '】(ID: ' + id + ')');
            lines.push('  说明：' + tpl.description);
            let paramEntries = Object.entries(tpl.params);
            if (paramEntries.length > 0) {
                lines.push('  参数：');
                for (let j = 0; j < paramEntries.length; j++) {
                    let pname = paramEntries[j][0];
                    let pinfo = paramEntries[j][1];
                    let req = pinfo.required ? '必填' : '可选(默认=' + pinfo.default + ')';
                    lines.push('    - ' + pname + ' (' + pinfo.type + '): ' + pinfo.description + ' [' + req + ']');
                }
            }
            lines.push('');
        }
        return lines.join('\n');
    }

    function buildSelectorMessages(turn) {
        let userMsg = turn.userMessage || '';
        let contextBlock = buildContextBlock(turn.context || {});
        let userContent = contextBlock
            ? userMsg + '\n\n' + contextBlock
            : userMsg;

        let templateList = buildTemplateList();
        let systemPrompt = TEMPLATE_SELECTOR_PROMPT.replace('{%TEMPLATES%}', templateList);

        // 注入可用数据列名，帮助 LLM 选择正确的列
        let headersText = '（当前选区无表头信息）';
        let ctx = turn.context || {};
        let sel = ctx.selection;
        let availableHeaders = [];
        if (sel && Array.isArray(sel.headers) && sel.headers.length > 0) {
            availableHeaders = sel.headers;
            headersText = sel.headers.map(function(h) { return '  - ' + h; }).join('\n');
            // 同时注入到 user message 里，确保 LLM 不会忽略
            userContent += '\n\n【可用数据列名（必须用实际列名，不要用"列名"等占位词）】\n' + headersText;
        } else {
            let refRegions = ctx.referenceRegions;
            if (Array.isArray(refRegions) && refRegions.length > 0) {
                for (let ri = 0; ri < refRegions.length; ri++) {
                    let rh = refRegions[ri];
                    if (Array.isArray(rh.headers)) {
                        for (let rhi = 0; rhi < rh.headers.length; rhi++) {
                            if (availableHeaders.indexOf(rh.headers[rhi]) === -1) availableHeaders.push(rh.headers[rhi]);
                        }
                    }
                }
                if (availableHeaders.length > 0) {
                    headersText = availableHeaders.map(function(h) { return '  - ' + h; }).join('\n');
                    userContent += '\n\n【可用数据列名（必须用实际列名，不要用"列名"等占位词）】\n' + headersText;
                }
            }
        }
        systemPrompt = systemPrompt.replace('{%HEADERS%}', headersText);

        // 后续消息继承：注入上轮分析结论 + 任务规划 + 意图（Phase 3）
        const followUpSpec = getPythonFollowUp(turn);
        if (followUpSpec) {
            const prevCtx = buildFollowUpPrompt(followUpSpec.spec, turn.userMessage || '', followUpSpec.intent);
            if (prevCtx) userContent += '\n\n' + prevCtx;
        }

        ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] buildSelectorMessages headers=' + headersText);

        return [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userContent }
        ];
    }

    // ============================================================
    // 工具函数
    // ============================================================

    function buildContextBlock(ctx) {
        let parts = [];
        let selection = ctx && ctx.selection;
        if (selection && Array.isArray(selection.headers) && selection.headers.length > 0) {
            parts.push(
                '【当前选区】' + (selection.address || selection.range || '') + '\n' +
                '工作簿：' + (selection.workbookName || selection.workbook || '') + '\n' +
                '工作表：' + (selection.sheetName || selection.worksheet || '') + '\n' +
                '表头：' + selection.headers.join(', ') + '\n' +
                '样本行数：' + (Array.isArray(selection.rows) ? selection.rows.length : 0)
            );
        }
        let refs = ctx && ctx.referenceRegions;
        if (Array.isArray(refs) && refs.length > 0) {
            parts.push('【引用区域】');
            for (let i = 0; i < refs.length; i++) {
                let r = refs[i];
                let headers = Array.isArray(r.headers) ? r.headers.join(', ') : '';
                let rowCount = Array.isArray(r.rows) ? r.rows.length : 0;
                parts.push(
                    '区域' + (i + 1) + ': [' + (r.workbook || '') + ']' + (r.worksheet || '') + '!' + (r.address || '') + '\n' +
                    '表头：' + headers + '\n样本行数：' + rowCount
                );
            }
        }
        return parts.join('\n\n');
    }

    // ============================================================
    // 后续消息继承（Phase 3）：读取上一轮 python _execSpec，注入分析结论
    // ============================================================
    function getPythonFollowUp(turn) {
        if (!turn || !turn.context || !turn.context._execFollowUp) return null;
        const ss = App.ChatV2.sessionStore;
        if (!ss || typeof ss.getLastExecSpec !== 'function') return null;
        const prev = ss.getLastExecSpec();
        if (!prev || !prev.spec || prev.spec.pipeline !== 'python') return null;
        return {
            spec: prev.spec,
            intent: turn.context._execFollowUp.intent || null
        };
    }

    function buildFollowUpPrompt(spec, userMessage, intent) {
        // 失败结果不注入：上一轮执行失败（含列名护栏拦截）时，其报错/护栏文案
        // 会污染本轮生成（模型把"数据为空，无法进行分析。"等报错串当作列名引用，形成幻觉循环）
        const prevRes = (spec && spec.result) || {};
        if (prevRes.success === false) return '';

        // 优先用意图分析器构造规划 prompt：注入任务结构 + 用户新需求 + 意图 + 规划指令，
        // 让 LLM 基于上一轮任务结构规划下一个任务，而非只靠关键词。
        const analyzer = App.ChatV2.intentAnalyzer;
        if (analyzer && typeof analyzer.buildFollowUpPlanningPrompt === 'function') {
            const prompt = analyzer.buildFollowUpPlanningPrompt({
                prevSpec: { spec: spec },
                userMessage: userMessage || '',
                intent: intent || null
            });
            if (prompt) return prompt;
        }
        // 兜底：旧逻辑（保持兼容）
        const parts = [];
        const res = (spec && spec.result) || {};
        if (res.summary) {
            parts.push('【上一轮分析结论】\n' + res.summary);
        }
        if (Array.isArray(res.steps) && res.steps.length) {
            const lines = res.steps.map(function (s) {
                return '- ' + (s.displayName || s.templateId || '分析') + (s.summary ? '：' + s.summary : '');
            });
            parts.push('【上一轮分析内容】\n' + lines.join('\n'));
        }
        // 注入上一轮任务结构（分析步骤/任务清单），让后续分析基于上轮规划继续而非只靠关键词
        const ts = spec && spec.taskStructure;
        if (ts && Array.isArray(ts.tasks) && ts.tasks.length) {
            const planLines = ts.tasks.map(function (t, i) {
                return (i + 1) + '. ' + (t.displayName || t.templateId || '分析');
            });
            parts.push('【上一轮任务规划】\n' + planLines.join('\n'));
        }
        if (parts.length) {
            parts.push('请基于以上分析结果与任务规划，按我的新需求继续分析。若新需求是对上一轮结果的进一步细分、筛选或对比，请在方案中体现。');
            return parts.join('\n\n');
        }
        return '';
    }

    function extractResultSummary(execResult) {
        if (!execResult) return '';
        let obj = execResult.result;
        if (typeof obj === 'string') {
            try { obj = JSON.parse(obj); } catch (e) { return String(obj).slice(0, 200); }
        }
        if (obj && obj.summary) return String(obj.summary);
        return '';
    }

    // JSON 解析已收敛到 services/json-utils.js（App.ChatV2.jsonUtils）
    // 以下为兼容转发，保持本文件的内部调用点不变、行为不变。
    function extractJSON(text) {
        const ju = App.ChatV2.jsonUtils;
        return ju ? ju.extractJsonObject(text) : null;
    }

    /**
     * 从执行结果文本中提取 JSON（支持对象或数组，兼容前缀文本）
     */
    function extractJson(text) {
        const ju = App.ChatV2.jsonUtils;
        return ju ? ju.extractJson(text) : text;
    }

    /**
     * 将 LLM 输出的 JSON 规范化为统一的 {templates: [...], reasoning: "..."} 格式
     * 支持新格式 (templates 数组) 和旧格式 (template 单值)
     */
    function normalizeSelection(raw) {
        if (!raw) return null;

        if (Array.isArray(raw.templates)) {
            return raw;
        }

        if (raw.template) {
            let entry = { id: raw.template };
            if (raw.params) entry.params = raw.params;
            let reasoning = raw.reasoning || '';
            return { templates: [entry], reasoning: reasoning };
        }

        return null;
    }

    function fuzzyMatch(a, b) {
        if (a === b) return 1.0;
        if (a.indexOf(b) !== -1 || b.indexOf(a) !== -1) return 0.8;
        let la = a.length, lb = b.length;
        let maxLen = Math.max(la, lb);
        let dist = levenshtein(a, b);
        return 1 - dist / maxLen;
    }

    function levenshtein(a, b) {
        let la = a.length, lb = b.length;
        if (la === 0) return lb;
        if (lb === 0) return la;
        let prev = [];
        for (let j = 0; j <= lb; j++) prev[j] = j;
        for (let i = 1; i <= la; i++) {
            let curr = [i];
            for (let j2 = 1; j2 <= lb; j2++) {
                let cost = a[i - 1] === b[j2 - 1] ? 0 : 1;
                curr[j2] = Math.min(
                    prev[j2] + 1,
                    curr[j2 - 1] + 1,
                    prev[j2 - 1] + cost
                );
            }
            prev = curr;
        }
        return prev[lb];
    }

    function extractPythonCode(content) {
        if (!content) return '';
        let trimmed = content.trim();
        let codeBlockMatch = trimmed.match(/```(?:python)?\s*([\s\S]*?)```/);
        if (codeBlockMatch) return codeBlockMatch[1].trim();
        return trimmed;
    }

    function resolvePythonToolName(toolRegistry) {
        let candidates = [
            'excel-local__excel_run_python_script',
            'excel_run_python_script',
            'run_python_script'
        ];
        for (let i = 0; i < candidates.length; i++) {
            if (toolRegistry.get(candidates[i])) return candidates[i];
        }
        return 'excel_run_python_script';
    }

    /**
     * 预读选区数据到 turn.context（如果还没有 rows）。
     * 调用 excel_read_range 读取当前选区/已用范围的真实 headers+rows，
     * 这样 buildSelectorMessages 能拿到真实列名，buildExecArgs 能拿到真实数据，
     * 从根源杜绝 LLM 编造列名和假数据。
     */
    async function ensureSelectionLoaded(turn, toolRegistry, signal) {
        if (!turn.context) turn.context = {};
        let ctx = turn.context;
        let sel = ctx.selection;

        // 已有有效 rows：无需再读
        let hasRows = false;
        if (sel && Array.isArray(sel.rows) && sel.rows.length > 0) hasRows = true;
        if (!hasRows && Array.isArray(ctx.referenceRegions)) {
            for (let i = 0; i < ctx.referenceRegions.length; i++) {
                if (ctx.referenceRegions[i] && Array.isArray(ctx.referenceRegions[i].rows) && ctx.referenceRegions[i].rows.length > 0) {
                    hasRows = true; break;
                }
            }
        }
        if (hasRows) return;

        let readToolName = null;
        if (toolRegistry.get('excel_read_range')) readToolName = 'excel_read_range';
        if (!readToolName) return;

        // 组装读取参数：优先用 selection 中的 sheet/range，缺省用活动工作表
        let args = {};
        let activeSheet = ctx.activeSheet || {};
        if (sel) {
            if (sel.sheetName) args.sheet = String(sel.sheetName);
            else if (sel.worksheet) args.sheet = String(sel.worksheet);
            if (sel.address) args.range = String(sel.address);
            else if (sel.range) args.range = String(sel.range);
        }
        if (!args.sheet && activeSheet.sheet) args.sheet = String(activeSheet.sheet);

        // object 格式返回（[{列名:值}, ...]），脚本和列名解析都更方便
        args.returnFormat = 'object';
        args.hasHeaders = true;
        args.maxRows = 10000;

        ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 预读选区: ' + JSON.stringify(args));

        let readResult;
        try {
            readResult = await toolRegistry.invoke(readToolName, args, { ctx: { turn: turn, signal: signal } });
        } catch (invErr) {
            ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] excel_read_range 调用失败:', invErr.message);
            return;
        }

        let parsed = null;
        if (typeof readResult === 'string') {
            parsed = safeParse(readResult);
        } else if (readResult && typeof readResult === 'object') {
            parsed = readResult.result ? (typeof readResult.result === 'string' ? safeParse(readResult.result) : readResult.result) : readResult;
        }
        if (!parsed) return;

        // excel_read_range 返回格式可能是 {success:true, headers:[...], rows:[...]} 或直接数组
        let newHeaders = null;
        let newRows = null;
        if (parsed.success === false) {
            ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 预读失败: ' + (parsed.message || parsed.error || 'unknown'));
            return;
        }
        if (Array.isArray(parsed)) {
            newRows = parsed;
            newHeaders = parsed.length > 0 && typeof parsed[0] === 'object' ? Object.keys(parsed[0]) : [];
        } else {
            newHeaders = Array.isArray(parsed.headers) ? parsed.headers : (Array.isArray(parsed.columns) ? parsed.columns : null);
            newRows = Array.isArray(parsed.rows) ? parsed.rows : (Array.isArray(parsed.data) ? parsed.data : null);
        }

        if (newRows && newRows.length > 0 && !newHeaders && typeof newRows[0] === 'object') {
            newHeaders = Object.keys(newRows[0]);
        }

        // 回填 selection
        if (!ctx.selection) ctx.selection = {};
        if (newHeaders && Array.isArray(newHeaders) && newHeaders.length > 0) {
            ctx.selection.headers = newHeaders;
        }
        if (newRows && Array.isArray(newRows) && newRows.length > 0) {
            ctx.selection.rows = newRows;
        }
        if (!ctx.selection.address && args.range) ctx.selection.address = args.range;
        if (!ctx.selection.sheetName && args.sheet) ctx.selection.sheetName = args.sheet;
        if (!ctx.selection.workbookName && activeSheet.workbook) ctx.selection.workbookName = activeSheet.workbook;

        // 如果没有 referenceRegions，同时注入一个作为 buildExecArgs 的兜底
        if ((!Array.isArray(ctx.referenceRegions) || ctx.referenceRegions.length === 0) && newRows && newRows.length > 0) {
            ctx.referenceRegions = [{
                workbook: activeSheet.workbook,
                worksheet: ctx.selection.sheetName || args.sheet,
                address: ctx.selection.address || args.range,
                headers: ctx.selection.headers || [],
                rows: newRows
            }];
        }

        ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 预读完成: headers=' + (ctx.selection.headers || []).length + ', rows=' + (ctx.selection.rows || []).length);
    }

    function safeParse(str) {
        const ju = App.ChatV2.jsonUtils;
        return ju ? ju.safeParse(str) : (() => { try { return JSON.parse(str); } catch (e) { return null; } })();
    }

    function buildExecArgs(script, turn, isTemplateMode) {
        let args = { script: script, timeout: 120 };
        // 模板执行（固定安全代码）标记：标准模式下安全守卫放行模板，仅拦截自定义脚本
        if (isTemplateMode) args._templateMode = true;
        let refs = turn.context && turn.context.referenceRegions;
        let built = [];
        if (Array.isArray(refs) && refs.length > 0) {
            for (let ri = 0; ri < refs.length; ri++) {
                let r = refs[ri];
                if (r && (Array.isArray(r.rows) && r.rows.length > 0 || Array.isArray(r.headers) && r.headers.length > 0)) {
                    built.push({
                        workbook: r.workbook || r.workbookName,
                        worksheet: r.worksheet || r.sheetName,
                        address: r.address,
                        headers: r.headers,
                        rows: r.rows
                    });
                }
            }
        }
        // 兜底：如果 referenceRegions 没数据，把当前 selection 注入
        if (built.length === 0 && turn.context && turn.context.selection) {
            let sel = turn.context.selection;
            if (Array.isArray(sel.rows) && sel.rows.length > 0) {
                built.push({
                    workbook: sel.workbook || sel.workbookName,
                    worksheet: sel.worksheet || sel.sheetName,
                    address: sel.address || sel.range,
                    headers: Array.isArray(sel.headers) ? sel.headers : [],
                    rows: sel.rows
                });
            }
        }
        if (built.length > 0) args.referenceRegions = built;

        // 关键修复：LLM 生成的脚本可能硬编码了占位符
        //   `input_data = [{"日期": "2023-01-01", ...}, # ... 其他数据]`，
        //   这会覆盖后端 PythonExecutionService 注入的真实 input_data（来自 Excel 选中区域）。
        // 修复：移除脚本中所有对【后端注入变量】的赋值语句（保留读取/使用），
        //   让后端注入的真实数据生效；同时移除空的初始化，避免遮蔽。
        //   只在确认存在真实数据源（referenceRegions/selection.rows）时重写。
        //   覆盖全部后端注入变量：
        //     input_data（真实数据）、selection_data（选区字典）、reference_data（引用区域）、
        //     input_headers（表头）、_sel（安全访问助手）——一律禁止 LLM 赋值。
        let hasRealData = built.length > 0;
        if (hasRealData && script && typeof script === 'string') {
            let origScript = script;
            // 1. 移除各注入变量的赋值（多行数组/字典/json.loads/None/字符串）
            //    正则按变量名分组，避免误伤使用语句（如 pd.DataFrame(input_data)）
            let INJECTED_VARS = ['input_data', 'selection_data', 'reference_data', 'input_headers', '_sel'];
            for (let vi = 0; vi < INJECTED_VARS.length; vi++) {
                let varname = INJECTED_VARS[vi];
                // 多行赋值：`var = [ ... ]` / `var = { ... }` / `var = json.loads(...)` / `var = None` / `var = "..."` 直到行尾
                let multiRe = new RegExp(
                    '^\\s*' + varname + '\\s*=\\s*(?:\\[[\\s\\S]*?\\]|\\{[\\s\\S]*?\\}|json\\.loads?\\([^)]*\\)|None|""|\'\')\\s*(?:#[^\\n]*)?\\n?', 'gm');
                script = script.replace(multiRe, '# [自动] ' + varname + ' 由系统注入真实数据，移除 LLM 占位符赋值\n');
                // 单行数组赋值兜底
                let singleRe = new RegExp('^\\s*' + varname + '\\s*=\\s*\\[[^\\n]*\\]\\s*$', 'gm');
                script = script.replace(singleRe, '# [自动] ' + varname + ' 由系统注入真实数据，移除 LLM 占位符赋值');
            }
            if (script !== origScript) {
                ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 已移除脚本中 LLM 编造的注入变量赋值（input_data/selection_data/reference_data/input_headers/_sel），改用后端注入的真实数据');
            }
        }

        args.script = script;
        return args;
    }

    // 参数名 → 中文标签映射
    let PARAM_LABELS = {
        columns: '分析列',
        value_column: '数值列',
        group_column: '分组列',
        aggregation: '聚合方式',
        threshold: '异常阈值',
        bins: '分箱数',
        forecast_steps: '预测期数',
        column: '列名',
        cols: '列名',
        group_by: '分组依据',
        agg: '聚合方式'
    };

    let AGG_LABELS = {
        avg: '平均',
        sum: '求和',
        count: '计数',
        min: '最小',
        max: '最大'
    };

    /**
     * 构建用户友好的参数展示信息
     */
    function buildDisplayArgs(templateId, templateName, filledParams) {
        let info = { '分析方法': templateName || templateId };
        if (!filledParams || typeof filledParams !== 'object') return info;

        let keys = Object.keys(filledParams);
        for (let i = 0; i < keys.length; i++) {
            let key = keys[i];
            let label = PARAM_LABELS[key] || key;
            let val = filledParams[key];

            // 聚合方式特殊处理
            if (key === 'aggregation' || key === 'agg') {
                let aggVal = AGG_LABELS[String(val).toLowerCase()] || val;
                info[label] = aggVal;
                continue;
            }

            // 数组类型（如 columns）：显示列名列表
            if (Array.isArray(val)) {
                if (val.length === 0) {
                    info[label] = '(未指定)';
                } else if (val.length <= 5) {
                    info[label] = val.join('、');
                } else {
                    info[label] = val.slice(0, 5).join('、') + ' 等' + val.length + '列';
                }
                continue;
            }

            // 字符串：直接显示
            if (typeof val === 'string') {
                info[label] = val || '(未指定)';
                continue;
            }

            // 其他类型：转为字符串
            info[label] = String(val);
        }

        return info;
    }

    /**
     * 生成参数摘要行（Markdown 文本）
     * 格式：**参数**：分析列: 销售额, 利润 | 聚合方式: 平均
     */
    function buildParamsText(filledParams) {
        if (!filledParams || typeof filledParams !== 'object') return '';

        let parts = [];
        let keys = Object.keys(filledParams);
        for (let i = 0; i < keys.length; i++) {
            let key = keys[i];
            let label = PARAM_LABELS[key] || key;
            let val = filledParams[key];

            // 聚合方式特殊处理
            if (key === 'aggregation' || key === 'agg') {
                let aggVal = AGG_LABELS[String(val).toLowerCase()] || val;
                parts.push(label + ': ' + aggVal);
                continue;
            }

            // 数组类型
            if (Array.isArray(val)) {
                if (val.length === 0) {
                    parts.push(label + ': (未指定)');
                } else {
                    parts.push(label + ': ' + val.join('、'));
                }
                continue;
            }

            // 其他类型
            let strVal = (typeof val === 'string') ? val : String(val);
            parts.push(label + ': ' + (strVal || '(未指定)'));
        }

        if (parts.length === 0) return '';
        return '**参数**：' + parts.join(' | ');
    }

    function extractImagePaths(text) {
        if (!text || typeof text !== 'string') return [];
        let pattern = /([a-zA-Z]:[\\/][^\s"<>|:*?]+\.(?:png|jpg|jpeg|gif|bmp|webp))/gi;
        let matches = text.match(pattern) || [];
        return Array.from(new Set(matches));
    }

    function isPythonExecSuccess(execResult) {
        if (!execResult || !execResult.ok) return false;
        let r = execResult.result;
        if (r && (r.success === false || r.isError === true)) return false;
        return true;
    }

    /**
     * 执行前扫描脚本，检测可能的列名幻觉问题和变量名错误
     * 返回检测结果：{ hallucinatedColumns: [...], badVariableNames: [...] }
     */
    function detectScriptIssues(script, headers) {
        let result = { hallucinatedColumns: [], badVariableNames: [] };
        if (!script) return result;

        // 1. 检测可疑列名：只识别"列访问上下文"中的字符串字面量
        //    （df['X'] / df[['a','b']] / groupby('X') / sort_values('X') / row.get('X') 等）。
        //    不再扫描全部字面量——print 标签、图表文件名、防御性错误消息
        //    （如 print('数据为空，无法进行分析。')、_save_chart('地区分布.png')）会被误判为列名幻觉，
        //    导致"自定义脚本引用了不存在的列"拦截掉本可正常执行的脚本。
        if (headers && headers.length > 0) {
            const seen = {};
            const excluded = ['success', 'message', 'summary', 'result', 'output', 'error',
                'True', 'False', 'None', 'count', 'value', 'agg', 'mean', 'sum',
                'csv', 'excel', 'sheet', 'workbook', 'header', 'rows', 'index',
                'plot', 'chart', 'graph', 'figure', 'bar', 'line', 'pie',
                '标题', '标签', '图例', '颜色', '字体', '保存'];
            function checkLiteral(lit) {
                if (!lit || seen[lit]) return;
                seen[lit] = true;
                if (!/[\u4e00-\u9fff_]/.test(lit) || lit.length < 2) return;
                if (excluded.indexOf(lit.toLowerCase()) !== -1 || headers.indexOf(lit) !== -1) return;
                result.hallucinatedColumns.push(lit);
            }
            function checkInnerLiterals(str) {
                if (!str) return;
                const innerRe = /['"]([^'"]{1,30})['"]/g;
                let im;
                while ((im = innerRe.exec(str)) !== null) checkLiteral(im[1]);
            }
            const accessPatterns = [
                /\[\[([^\]]*)\]\]/g,                                  // df[['a','b']]
                /\[(['"])([^'"\]\\]{1,30})\1\]/g,                     // df['a'] / row['a']
                /\.(?:groupby|sort_values|drop|get|astype|fillna|melt|pivot|stack|unstack|set_index|reset_index)\(['"]([^'")\\]{1,30})['"]\)/g,
                /\.(?:groupby|sort_values|drop|melt|pivot|set_index)\(\[([^\]]*)\]\)/g,
                /\.rename\(columns=\{[^}]*\}/g,
                /on=['"]([^'"]{1,30})['"]/g
            ];
            for (let ai = 0; ai < accessPatterns.length; ai++) {
                const re = accessPatterns[ai];
                re.lastIndex = 0;
                let m;
                while ((m = re.exec(script)) !== null) {
                    if (re === accessPatterns[0] || re === accessPatterns[3] || re === accessPatterns[4]) {
                        checkInnerLiterals(m[1]);
                    } else {
                        // accessPatterns[1]（df['a']）第1组是引号、第2组才是列名
                        const cap = (re === accessPatterns[1]) ? m[2] : m[1];
                        if (cap !== undefined) checkLiteral(cap);
                    }
                }
            }
        }

        // 2. 检测变量名错误（使用 df 但未定义）
        let hasDataFrameDef = /df\s*=\s*pd\.DataFrame/.test(script) || /df\s*=\s*DataFrame/.test(script);
        let usesDf = /\bdf\./.test(script);
        if (usesDf && !hasDataFrameDef) {
            result.badVariableNames.push('df（未定义，应为 df = pd.DataFrame(input_data)）');
        }

        // 3. 检测使用 input_data 之外的数据入口
        let usesOtherDataEntry = /pd\.read_excel|pd\.read_csv|open\(.+\.xlsx|open\(.+\.csv/.test(script);
        if (usesOtherDataEntry) {
            result.badVariableNames.push('使用了文件读取而非 input_data 变量');
        }

        return result;
    }

    /**
     * 增强错误信息：区分框架错误、模板错误、Skill/LLM 生成错误
     */
    function classifyExecutionError(execResult, isCustomMode, displayName) {
        let errorMsg = getPythonError(execResult);
        if (!errorMsg) return null;

        let info = { errorType: '未知错误', suggestion: '' };

        if (errorMsg.includes('未读取到数据') || errorMsg.includes('无可用数据') || errorMsg.includes('当前未选中任何数据')) {
            info.errorType = '数据读取问题（未获取到数据）';
            info.suggestion = '未读取到有效数据。请先选中包含表头和数据的单元格区域（如 A1:G7），再重新执行。';
        } else if (errorMsg.includes('列不存在') || errorMsg.includes('无有效数值') || errorMsg.includes('无法计算描述统计')) {
            info.errorType = '模板参数与数据不匹配';
            info.suggestion = '指定的分析列与数据表头不匹配，或所选列没有有效数值。请检查列名是否与表头完全一致（如「销售额」≠「销售金额」），并确认数值列不为空。';
        } else if (errorMsg.includes('IndexError') || errorMsg.includes('KeyError') || errorMsg.includes('not found')) {
            info.errorType = 'Skill/LLM 生成错误（列名不匹配）';
            info.suggestion = 'LLM 生成的脚本中使用了数据中不存在的列名。请检查实际列名是否匹配，或在 Prompt 中明确指定列名。';
        } else if (errorMsg.includes('SyntaxError') || errorMsg.includes('syntax error')) {
            info.errorType = 'Skill/LLM 生成错误（语法错误）';
            info.suggestion = 'LLM 生成的 Python 脚本存在语法错误。建议使用模板模式而非自定义模式。';
        } else if (errorMsg.includes('timed out') || errorMsg.includes('timeout') || errorMsg.includes('超时')) {
            info.errorType = '框架/执行超时';
            info.suggestion = 'Python 脚本执行超时。可能是数据量过大或脚本陷入循环。建议分批处理或简化分析请求。';
        } else if (errorMsg.includes('ImportError') || errorMsg.includes('ModuleNotFound')) {
            info.errorType = '框架错误（环境缺失）';
            info.suggestion = 'Python 环境缺少必要的库。请检查 pandas、numpy 等是否已安装。';
        } else if (errorMsg.includes('permission') || errorMsg.includes('权限') || errorMsg.includes('Access')) {
            info.errorType = '框架/权限错误';
            info.suggestion = '文件访问被拒绝。请检查文件是否已关闭或路径是否正确。';
        } else if (isCustomMode) {
            info.errorType = 'Skill/LLM 生成错误';
            info.suggestion = '自定义脚本执行失败。建议改用预置模板（如描述统计、分组聚合等），或简化分析需求。';
        } else {
            info.errorType = '模板执行错误';
            // 展示真实错误信息（截断），避免笼统文案掩盖实际原因
            let detail = String(errorMsg || '').replace(/\s+/g, ' ').trim();
            if (detail.length > 200) detail = detail.slice(0, 200) + '…';
            info.suggestion = displayName + ' 模板执行失败。' +
                (detail && detail !== 'Python 执行失败' ? '原始错误：' + detail : '请检查数据是否符合该模板的要求（如数值列是否存在、数据是否为空）。');
        }

        return info;
    }

    function getPythonError(execResult) {
        if (!execResult) return '执行结果为空';
        if (execResult.error) return execResult.error;
        let r = execResult.result;
        if (r) {
            if (typeof r.result === 'string') return r.result;
            if (typeof r.error === 'string') return r.error;
            if (typeof r.message === 'string') return r.message;
        }
        return 'Python 执行失败';
    }

    function formatExecutionResult(execResult) {
        if (!execResult) return '无执行结果';
        if (!execResult.ok) return '❌ 执行失败：' + (execResult.error || '未知错误');

        let raw = execResult.result;
        let output = '';
        let parsed = null;

        if (typeof raw === 'string') {
            output = raw;
            parsed = extractJson(raw);
        } else if (raw && typeof raw === 'object') {
            if (raw.output !== undefined) {
                output = typeof raw.output === 'string' ? raw.output : JSON.stringify(raw.output);
                parsed = extractJson(output);
                if (!parsed) parsed = raw;
            } else {
                parsed = raw;
                output = JSON.stringify(raw);
            }
        }

        let lines = [];

        if (typeof parsed === 'object' && parsed !== null) {
            if (parsed.success === false) {
                return '❌ ' + (parsed.message || '执行失败');
            }
            if (parsed.summary) {
                lines.push('**分析结果**：' + parsed.summary);
            }
            if (parsed.result !== undefined) {
                if (typeof parsed.result === 'object') {
                    lines.push('```json');
                    lines.push(JSON.stringify(parsed.result, null, 2));
                    lines.push('```');
                } else {
                    lines.push(String(parsed.result));
                }
            } else if (parsed.output) {
                lines.push(String(parsed.output));
            } else if (parsed.data) {
                if (typeof parsed.data === 'object') {
                    lines.push('```json');
                    lines.push(JSON.stringify(parsed.data, null, 2));
                    lines.push('```');
                } else {
                    lines.push(String(parsed.data));
                }
            }
        } else if (output) {
            lines.push(output);
        }

        if (lines.length === 0 && raw) {
            lines.push(String(raw));
        }

        let imagePaths = extractImagePaths(output || '');
        if (imagePaths.length > 0) {
            lines.push('');
            lines.push('**生成的图表：**');
            for (let pi = 0; pi < imagePaths.length; pi++) {
                lines.push('![分析图表](' + chartImageUrl(imagePaths[pi]) + ')');
            }
        }

        return lines.join('\n');
    }

    // 图表图片 URL：WebView2 禁 file:// 本地图，转交给本地桥接服务的图片端点
    // 返回桥接 http 地址，<img> 跨协议加载无需 CORS；未拿到 baseUrl 时回退 localhost:9527
    function chartImageUrl(imgPath) {
        const base = (window.ExcelBridgeConfig && window.ExcelBridgeConfig.baseUrl)
            ? String(window.ExcelBridgeConfig.baseUrl).replace(/\/+$/, '')
            : (App.Endpoints ? App.Endpoints.excelBridge.baseUrl : 'http://localhost:9527');
        return base + '/api/chart-image?path=' + encodeURIComponent(String(imgPath).replace(/^file:\/\//i, ''));
    }

    // 把模板的结构化 result 转成 ECharts 内嵌配置（供 initInlineCharts 渲染 data-echarts 节点）。
    // 类型映射只覆盖适合"单分类轴"展示的模板（group/distribution→bar，trend/regression→line）；
    // 无法构造配置时返回 null（安全回退：不画图，仍显示表格/文本）。
    // 聚合方式 → 中文指标名（用于图表标题/表格列名，避免出现 "value"/"指标" 这类无意义标签）
    var AGG_METRIC_LABELS = { avg: '平均值', sum: '求和', count: '计数', min: '最小值', max: '最大值', median: '中位数' };

    // 相关性矩阵 → ECharts 热力图配置
    function buildCorrelationHeatmapCfg(matrix) {
        if (!matrix || typeof matrix !== 'object') return null;
        const cols = Object.keys(matrix);
        if (cols.length < 2) return null;
        const data = [];
        for (let i = 0; i < cols.length; i++) {
            for (let j = 0; j < cols.length; j++) {
                const v = matrix[cols[i]] && matrix[cols[i]][cols[j]];
                data.push([i, j, (typeof v === 'number') ? v : 0.0]);
            }
        }
        return { type: 'heatmap', xLabels: cols, yLabels: cols, data: data, labels: cols, values: cols.map(String), title: '相关性矩阵' };
    }

    function buildTemplateChartCfg(templateId, obj) {
        if (obj === null || obj === undefined) return null;
        if (typeof obj === 'string') { try { obj = JSON.parse(obj); } catch (e) { return null; } }
        if (obj === null || typeof obj !== 'object') return null;
        const typeMap = { group_aggregation: 'bar', data_distribution: 'bar', trend_analysis: 'line', linear_regression: 'line', correlation: 'heatmap', descriptive_stats: 'bar' };
        const type = typeMap[templateId];
        if (!type) return null;

        // ---- 各模板专用结果结构映射（优先于通用解析，避免把纯标量或嵌套数组误判成无意义图）----
        // 趋势分析：forecast_x(日期/日序号) + forecast(预测值) → 预测折线
        // 趋势分析：historical（历史真实值）+ forecast（未来预测值）→ 双序列折线图
        if (templateId === 'trend_analysis' && Array.isArray(obj.forecast_x) && obj.forecast_x.length && Array.isArray(obj.forecast)) {
            const hasHist = Array.isArray(obj.historical_x) && obj.historical_x.length && Array.isArray(obj.historical_values);
            const fx = obj.forecast_x.map(String);
            const fv = obj.forecast.map(Number);
            if (!hasHist) {
                // 降级：只有预测数据，单条折线
                return { type: 'line', labels: fx, values: fv, title: (obj.column || '数值') + ' · 未来预测' };
            }
            const hx = obj.historical_x.map(String);
            const hv = obj.historical_values.map(Number);
            // 合并 x 轴：历史 + 预测（预测线从最后一个历史点开始，保持视觉连续）
            const allLabels = hx.concat(fx);
            const histData = hv.concat(new Array(fv.length).fill(null));
            // 预测序列：在最后一个历史点处衔接，然后是预测值
            const forecastData = new Array(hv.length - 1).fill(null).concat([hv[hv.length - 1]]).concat(fv);
            return {
                type: 'line',
                labels: allLabels,
                values: histData, // 兼容旧版单 series 取值
                title: (obj.column || '数值') + ' · 趋势与预测',
                series: [
                    { name: '历史数据', data: histData, smooth: false },
                    { name: '预测数据', data: forecastData, smooth: false, lineStyle: { type: 'dashed' } }
                ]
            };
        }
        // 数据分布：每列生成一个频数直方图(bar)。
        // 多列（如"销售额/成本/利润分布"）时返回配置数组 → 每列一张图，
        // 避免只取首个列导致"三张表只有一张图"。
        if (templateId === 'data_distribution') {
            const cfgs = [];
            Object.keys(obj).forEach(function (k) {
                const col = obj[k];
                if (col && Array.isArray(col.distribution) && col.distribution.length) {
                    cfgs.push({
                        type: 'bar',
                        labels: col.distribution.map(function (d) { return String(d.range); }),
                        values: col.distribution.map(function (d) { return Number(d.count || 0); }),
                        title: k + ' · 频数分布'
                    });
                }
            });
            if (cfgs.length === 1) return cfgs[0];  // 单列 → 兼容旧返回类型（单配置）
            if (cfgs.length > 1) return cfgs;       // 多列 → 返回配置数组（每列一图）
            return null;  // 形状不符（无 distribution 数组）→ 安全回退表格，不落通用解析
        }
        // 描述统计：每列一张"统计指标柱状图"（平均值/中位数/标准差/最小/最大），
        // 多列时返回配置数组（每列一图——不同列量纲不同，混排会互相压扁）。
        if (templateId === 'descriptive_stats') {
            const cfgs = [];
            Object.keys(obj).forEach(function (k) {
                const st = obj[k];
                if (st && typeof st === 'object') {
                    const labels = ['平均值', '中位数', '标准差', '最小值', '最大值'];
                    const values = [st.mean, st.median, st.stdev, st.min, st.max];
                    // 至少 3 个统计量为数值才视为有效统计结果（真实模板输出 5 个全有；
                    // {count,mean} 这类残缺对象不画图，避免 1 根柱 + 4 空柱的伪图）
                    const numCount = values.filter(function (v) { return typeof v === 'number'; }).length;
                    if (numCount >= 3) {
                        cfgs.push({
                            type: 'bar',
                            labels: labels,
                            values: values.map(function (v) { return (typeof v === 'number') ? v : null; }),
                            title: k + ' · 描述统计'
                        });
                    }
                }
            });
            if (cfgs.length === 1) return cfgs[0];  // 单列 → 兼容旧返回类型（单配置）
            if (cfgs.length > 1) return cfgs;       // 多列 → 返回配置数组（每列一图）
            return null;  // 形状不符（如 [{列,均值}] 行式数组/残缺统计对象）→ 安全回退表格，不落通用解析
        }
        // 线性回归：forecast[{x,y}] → 拟合线；形状不符（如标量/公式段）→ 安全回退 null
        if (templateId === 'linear_regression' && Array.isArray(obj.forecast) && obj.forecast.length) {
            return { type: 'line', labels: obj.forecast.map(function (p) { return String(p.x); }), values: obj.forecast.map(function (p) { return Number(p.y); }), title: obj.equation || '线性拟合' };
        }
        if (templateId === 'linear_regression') return null;
        // 相关性：相关系数矩阵 → 热力图；非矩阵 → 安全回退 null
        if (templateId === 'correlation') {
            return buildCorrelationHeatmapCfg(obj) || null;
        }

        let labels = null, values = null, title = '';
        if (Array.isArray(obj)) {
            if (obj.length && typeof obj[0] === 'object' && obj[0] !== null) {
                const o = obj[0];
                const keys = Object.keys(o);
                const labelKey = keys.find(function (k) { return typeof o[k] === 'string'; });
                const valueKey = keys.find(function (k) { return typeof o[k] === 'number'; });
                if (labelKey && valueKey !== undefined) {
                    labels = obj.map(function (r) { return String(r ? r[labelKey] : ''); });
                    values = obj.map(function (r) { return Number(r ? (r[valueKey] || 0) : 0); });
                    title = '按 ' + labelKey + ' 的 ' + valueKey;
                } else if (valueKey !== undefined) {
                    labels = obj.map(function (_, i) { return String(i + 1); });
                    values = obj.map(function (r) { return Number(r ? (r[valueKey] || 0) : 0); });
                }
            } else if (obj.length && obj.every(function (x) { return typeof x === 'number'; })) {
                labels = obj.map(function (_, i) { return String(i + 1); });
                values = obj.slice();
            }
        } else if (Array.isArray(obj.labels) && obj.labels.length && Array.isArray(obj.values)) {
            labels = obj.labels;
            values = obj.values;
            title = obj.title || '';
        } else {
            // 平面对象映射：{类别: 数值}（如 {华北:3000}）或 {类别: {数值字段...}}（分组聚合正值 {组名:{value,count}}）
            const keys = Object.keys(obj);
            const numKeys = keys.filter(function (k) { return typeof obj[k] === 'number'; });
            if (numKeys.length > 0) {
                labels = keys.map(String);
                values = keys.map(function (k) { return Number(obj[k] || 0); });
                title = numKeys[0];
            } else {
                // 值本身是对象：在所有条目中找一个"数值且跨组一致"的字段作为序列值（优先 value/平均值/求和值/频数 等）
                const preferred = ['value', '求和值', '平均值', '均值', '频数', 'count', 'anomaly_count', '值'];
                const sampleKey = keys.find(function (k) { return obj[k] && typeof obj[k] === 'object'; });
                if (sampleKey !== undefined) {
                    const subKeys = Object.keys(obj[sampleKey] || {});
                    const isAllNumeric = function (field) {
                        return keys.length > 0 && keys.every(function (g) {
                            const v = obj[g] && obj[g][field];
                            return typeof v === 'number';
                        });
                    };
                    let valSubKey = null;
                    for (let pi = 0; pi < preferred.length; pi++) {
                        if (subKeys.indexOf(preferred[pi]) >= 0 && isAllNumeric(preferred[pi])) { valSubKey = preferred[pi]; break; }
                    }
                    if (valSubKey === null) {
                        valSubKey = subKeys.find(isAllNumeric) || null;
                    }
                    if (valSubKey !== null) {
                        labels = keys.map(String);
                        values = keys.map(function (k) { return Number((obj[k] || {})[valSubKey] || 0); });
                        // 标题优先用真实聚合方式（avg→平均值、sum→求和、count→计数…），
                        // 避免出现裸英文 "value" 这种对用户无意义的标签。
                        const aggLabel = AGG_METRIC_LABELS[obj[sampleKey] && obj[sampleKey].agg];
                        title = aggLabel || valSubKey;
                    }
                }
            }
        }

        if (labels && labels.length && Array.isArray(values) && values.length) {
            return { type: type, labels: labels, values: values, title: title };
        }
        return null;
    }

    // 从模板执行结果中构建可见的 ECharts 内嵌节点（data-echarts，被前端 initInlineCharts 扫描渲染）。
    // 与 formatAnalysisResult 复用同一套 parsed.result 解析逻辑；无法构造时返回 ''（安全回退）。
    function parseTemplateResultJson(text) {
        if (typeof text !== 'string') return null;
        const t = text.trim();
        if (t.charAt(0) === '{' || t.charAt(0) === '[') {
            try { return JSON.parse(t); } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
        }
        try {
            const ju = App.ChatV2 && App.ChatV2.jsonUtils;
            if (ju && ju.safeParse) { const r = ju.safeParse(text); if (r) return r; }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
        // 兜底：复用 formatAnalysisResult 的 extractJson（兼容 print 前缀/混合输出的结果文本）
        try {
            const ju = App.ChatV2 && App.ChatV2.jsonUtils;
            if (ju && ju.extractJson) {
                const r = ju.extractJson(text);
                if (r && typeof r === 'object') return r;
            }
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
        // 兜底：兼容 Python dict repr（单引号 + True/False/None + 尾逗号，JS JSON.parse 解析不了）
        try {
            const pr = parsePythonRepr(text);
            if (pr && typeof pr === 'object') return pr;
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
        // 兜底：容错解析后端"JSON 变体"（前置成功提示 + 成员缺逗号，如 "Python 脚本执行成功\n{...}"）
        try {
            const lj = parseLenientJson(text);
            if (lj && typeof lj === 'object') return lj;
        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
        return null;
    }

    /**
     * 把最外层为 Python 字面量（dict/list repr，单引号、True/False/None、尾逗号）的文本转成 JS 对象。
     * 背景：部分 Python 执行工具把 output 以 str() 的形式回传（例：{'success': True, 'result': {...}}），
     * 这类文本既不是合法 JSON，extractJson 的冒号/花括号正则也会被单引号破坏，导致图表取不到数据。
     */
    function parsePythonRepr(text) {
        if (typeof text !== 'string') return null;
        const start = text.indexOf('{');
        const arrStart = text.indexOf('[');
        if (start === -1 && arrStart === -1) return null;
        let s = start, brace = true;
        if (arrStart !== -1 && (start === -1 || arrStart < start)) { s = arrStart; brace = false; }
        let depth = 0, end = -1;
        for (let i = s; i < text.length; i++) {
            const c = text[i];
            const open = brace ? '{' : '[';
            const close = brace ? '}' : ']';
            if (c === open) depth++;
            else if (c === close) { depth--; if (depth === 0) { end = i + 1; break; } }
        }
        if (end === -1) return null;
        const body = text.substring(s, end);
        try {
            return JSON.parse(pythonReprToJsonString(body));
        } catch(e) { return null; }
    }

    // 把单引号 Python repr 文本改写成合法 JSON 文本（正确处理 Python 字符串转义）
    function pythonReprToJsonString(body) {
        const escMap = { n: 'n', t: 't', r: 'r', b: 'b', f: 'f', '\\': '\\', '/': '/' };
        let out = '';
        let inStr = false;
        let i = 0;
        while (i < body.length) {
            const c = body[i];
            if (inStr) {
                if (c === '\\') {
                    const e = body[i + 1];
                    // \uXXXX / \xHH 等 Unicode 转义：原样透传（JSON 同样支持 \u）
                    if (e === 'u') { out += body.substr(i, 6); i += 6; continue; }
                    if (e === 'x') {
                        const hex = body.substr(i, 4);
                        const code = parseInt(hex.slice(2), 16);
                        out += (isNaN(code) ? hex : String.fromCharCode(code));
                        i += 4; continue;
                    }
                    if (e === "'") { out += "'"; i += 2; continue; }      // Python 对单引号的转义 → 还原为字面单引号
                    if (e === '"') { out += '\\"'; i += 2; continue; }     // 双引号需 JSON 转义
                    if (escMap[e]) { out += '\\' + escMap[e]; i += 2; continue; }
                    out += '\\' + e; i += 2; continue;
                }
                if (c === "'") { out += '"'; inStr = false; i++; continue; }
                if (c === '"') { out += '\\"'; i++; continue; }
                out += c; i++; continue;
            }
            if (c === "'") { out += '"'; inStr = true; i++; continue; }
            if (/[a-zA-Z]/.test(c)) {
                let j = i;
                while (j < body.length && /[a-zA-Z0-9_]/.test(body[j])) j++;
                const id = body.substring(i, j);
                const rep = { 'True': 'true', 'False': 'false', 'None': 'null' }[id];
                out += (rep !== undefined ? rep : id);
                i = j; continue;
            }
            out += c; i++;
        }
        // 去掉尾逗号（, 直接出现在 ] 或 } 之前）
        return out.replace(/,\s*([\]}])/g, '$1');
    }

    /**
     * 容错解析后端返回的"JSON 变体"文本：
     * - 允许前置文本（如 "Python 脚本执行成功\n"）
     * - 允许成员/元素之间缺逗号（如 {"a": 1\n "b": 2}）
     * - 兼容单/双引号、True/False/None 等 Python 风格字面量
     * 采用递归下降解析最外层 {...} / [...]，最健壮；返回 JS 对象/数组，失败返回 null。
     */
    function parseLenientJson(text) {
        if (typeof text !== 'string') return null;
        const start = text.indexOf('{');
        const arrStart = text.indexOf('[');
        let s = start, isObj = true;
        if (start === -1 && arrStart === -1) return null;
        if (arrStart !== -1 && (start === -1 || arrStart < start)) { s = arrStart; isObj = false; }
        const open = isObj ? '{' : '[';
        const close = isObj ? '}' : ']';
        let depth = 0, end = -1;
        for (let i = s; i < text.length; i++) {
            const c = text[i];
            if (c === open) depth++;
            else if (c === close) { depth--; if (depth === 0) { end = i + 1; break; } }
        }
        if (end === -1) return null;
        try {
            return lenientParse(text.substring(s, end), isObj);
        } catch(e) { return null; }
    }

    function lenientParse(src, isObj) {
        let i = 0;
        function skipWs() { while (i < src.length && /\s/.test(src[i])) i++; }
        function parseString() {
            const q = src[i]; i++;
            let out = '';
            while (i < src.length && src[i] !== q) {
                if (src[i] === '\\' && i + 1 < src.length) {
                    const e = src[i + 1];
                    if (e === 'n') out += '\n';
                    else if (e === 't') out += '\t';
                    else if (e === 'r') out += '\r';
                    else if (e === 'u') { const code = parseInt(src.substring(i + 2, i + 6), 16); out += isNaN(code) ? '' : String.fromCharCode(code); i += 5; }
                    else out += e;
                    i += 2; continue;
                }
                out += src[i]; i++;
            }
            i++;
            return out;
        }
        function parseValue() {
            skipWs();
            if (i >= src.length) return undefined;
            const c = src[i];
            if (c === '{' || c === '[') return parseContainer(c === '{');
            if (c === '"' || c === "'") return parseString();
            // 数字
            if (/[-+0-9]/.test(c)) {
                let j = i;
                while (j < src.length && /[-+0-9.eE]/.test(src[j])) j++;
                const num = src.substring(i, j);
                i = j;
                if (/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(num)) return parseFloat(num);
                return num;
            }
            // 裸词（True/False/None/其他）
            let k = i;
            while (k < src.length && /[a-zA-Z0-9_.\-]/.test(src[k])) k++;
            const word = src.substring(i, k); i = k;
            if (word === 'True' || word === 'true') return true;
            if (word === 'False' || word === 'false') return false;
            if (word === 'None' || word === 'null') return null;
            return word;
        }
        function parseContainer(isContainerObj) {
            i++;
            const result = isContainerObj ? {} : [];
            skipWs();
            while (i < src.length) {
                skipWs();
                if (i >= src.length) break;
                const c = src[i];
                if ((isContainerObj && c === '}') || (!isContainerObj && c === ']')) { i++; break; }
                if (isContainerObj) {
                    let key;
                    if (c === '"' || c === "'") key = parseString();
                    else {
                        let k = i;
                        while (k < src.length && !/[\s,:{}[\]]/.test(src[k])) k++;
                        key = src.substring(i, k); i = k;
                    }
                    skipWs();
                    if (src[i] === ':') i++;
                    result[key] = parseValue();
                } else {
                    result.push(parseValue());
                }
                skipWs();
                if (src[i] === ',') i++;
                // 缺逗号也能继续（下一轮循环直接解析下一个成员）
            }
            return result;
        }
        return parseContainer(isObj);
    }

    function buildResultChartDiv(templateId, execResult) {
        if (!templateId || !execResult || !execResult.ok) {
            ChatLog.info && ChatLog.info('PYCHART', '[buildResultChartDiv] skip: templateId=' + templateId + ' hasExec=' + !!execResult + ' ok=' + (execResult && execResult.ok));
            return '';
        }
        let raw = execResult.result;
        let parsed = null;
        // 【修复-趋势图表】提取"结构化 JSON 候选文本"：
        // 优先取 [OUTPUT_JSON] 标记后的完整 output（VB 端附加的结构化数据），
        // 否则用整段文本（兼容 "Python 脚本执行成功\n{...}" 前置文本）。
        const OUTPUT_JSON_MARK = '[OUTPUT_JSON]';
        function extractJsonText(t) {
            if (typeof t !== 'string') return null;
            const markIdx = t.indexOf(OUTPUT_JSON_MARK);
            const part = markIdx >= 0 ? t.slice(markIdx + OUTPUT_JSON_MARK.length) : t;
            const trimmed = part.trim();
            return trimmed || null;
        }
        if (typeof raw === 'string') {
            const jt = extractJsonText(raw);
            if (jt) parsed = parseTemplateResultJson(jt);
        } else if (raw && typeof raw === 'object') {
            if (Array.isArray(raw.content)) {
                // MCP content 包装（外部 MCP 工具结果）：拼接 text 项后解析，优先 [OUTPUT_JSON]
                const texts = raw.content
                    .filter(function (c) { return c && typeof c === 'object' && typeof c.text === 'string'; })
                    .map(function (c) { return c.text; });
                const joined = texts.join('\n');
                const jt = extractJsonText(joined);
                if (jt) parsed = parseTemplateResultJson(jt);
                if (!parsed) parsed = raw;
            } else if (raw.output !== undefined) {
                parsed = parseTemplateResultJson(typeof raw.output === 'string' ? raw.output : JSON.stringify(raw.output));
                if (!parsed) parsed = raw;
            } else {
                parsed = raw;
            }
        }
        const chartObj = (parsed && typeof parsed === 'object' && parsed !== null && parsed.result != null)
            ? parsed.result : parsed;
        const chartCfg = buildTemplateChartCfg(templateId, chartObj);
        if (!chartCfg) {
            ChatLog.info && ChatLog.info('PYCHART', '[buildResultChartDiv] no-cfg: templateId=' + templateId + ' raw=' + (typeof raw === 'string' ? 'string(' + raw.length + ')' : (raw ? typeof raw + '+' + Object.keys(raw).join(',') : String(raw))) + ' parsedKeys=' + (parsed && typeof parsed === 'object' ? Object.keys(parsed).join(',') : 'null'));
            return '';
        }
        // 多列分布：buildTemplateChartCfg 返回配置数组 → 每列渲染一个图表 div（三张表 → 三张图）
        if (Array.isArray(chartCfg)) {
            const divs = chartCfg.map(function (cfg) { return buildChartDiv(cfg); }).join('');
            ChatLog.info && ChatLog.info('PYCHART', '[buildResultChartDiv] built multi-chart div len=' + divs.length + ' count=' + chartCfg.length);
            return divs;
        }
        const div = buildChartDiv(chartCfg);
        ChatLog.info && ChatLog.info('PYCHART', '[buildResultChartDiv] built div len=' + div.length + ' type=' + chartCfg.type + ' labels=' + chartCfg.labels.length);
        return div;
    }

    // 生成前端内嵌 ECharts 节点：data-echarts 属性会被 initInlineCharts 扫描并渲染。
    // 自适应宽度：外包装一层横向滚动容器(width:100%)，内层图表按分类数估算一个最小宽度。
    // 分类少时 min-width:100% 填满气泡(自适应)；分类多/标签长时宽度超过气泡，可横向拖动看全，避免右侧被裁。
    function buildChartDiv(cfg) {
        if (!cfg) return '';
        const safe = JSON.stringify(cfg).replace(/'/g, '&#39;');
        let innerW = null;
        if (cfg.type === 'bar' && Array.isArray(cfg.labels) && cfg.labels.length > 0) {
            const n = cfg.labels.length;
            const maxChars = Math.max.apply(null, cfg.labels.map(function (l) { return String(l == null ? '' : l).length; })) || 1;
            const est = n * (Math.min(maxChars, 8) * 16 + 40); // 每分类约「标签宽 + 柱位」
            if (est > 420) innerW = Math.min(Math.max(est, 420), 1400);
        }
        const innerStyle = "min-width:100%;height:300px;" + (innerW ? "width:" + innerW + "px;" : "width:100%;");
        return "\n<div style=\"width:100%;max-width:100%;overflow-x:auto;margin:10px 0;\"><div data-echarts='" + safe + "' style=\"" + innerStyle + "\"></div></div>";
    }

    // 把模板的结构化 result 转成友好 markdown：
    // 同构对象数组 → 表格；二维矩阵（相关性）→ 表格；纯对象 → 键值列表；标量 → 原样。
    function templateResultToMarkdown(r) {
        const isObj = function (x) { return typeof x === 'object' && x !== null && !Array.isArray(x); };
        if (Array.isArray(r) && r.length && r.every(isObj)) {
            const cols = []; const seen = {};
            r.forEach(function (o) { Object.keys(o).forEach(function (k) { if (!seen[k]) { seen[k] = 1; cols.push(k); } }); });
            let s = '| ' + cols.join(' | ') + ' |\n|' + cols.map(function () { return '---'; }).join('|') + '|';
            r.forEach(function (o) {
                s += '\n| ' + cols.map(function (c) { return o[c] === undefined || o[c] === null ? '' : String(o[c]); }).join(' | ') + ' |';
            });
            return s;
        }
        const isMat = Array.isArray(r) && r.length && r.every(function (x) { return Array.isArray(x); });
        if (isMat) {
            const w = Math.max.apply(null, r.map(function (x) { return x.length; }));
            const head = r[0].map(function (_, j) { return (j + 1); });
            let s = '| # | ' + head.join(' | ') + ' |\n|--|' + r[0].map(function () { return '---'; }).join('|') + '|';
            r.forEach(function (row, i) {
                s += '\n| ' + (i + 1) + ' | ' + row.map(function (v) { return String(v); }).join(' | ') + ' |';
            });
            return s;
        }
        if (isObj(r)) {
            return Object.keys(r).map(function (k) { return '- **' + k + '**：' + String(r[k]); }).join('\n');
        }
        return String(r);
    }

    function formatAnalysisResult(execResult, templateId, templateName) {
        if (!execResult) return '无结果';
        if (!execResult.ok) return '❌ 执行失败：' + (execResult.error || '未知错误');

        let raw = execResult.result;
        let parsed = null;
        let rawOutput = '';
        let rawJsonText = '';

        if (typeof raw === 'string') {
            rawOutput = raw;
            parsed = extractJson(raw);
            if (!parsed) {
                parsed = safeParse(raw);
            }
        } else if (raw && typeof raw === 'object') {
            // 优先取 output 字段（Python 执行工具的标准输出）
            if (raw.output !== undefined) {
                rawOutput = typeof raw.output === 'string' ? raw.output : JSON.stringify(raw.output);
                parsed = extractJson(rawOutput);
                if (!parsed) {
                    parsed = safeParse(rawOutput) || raw;
                }
            } else {
                parsed = raw;
                rawOutput = JSON.stringify(raw);
            }
        }

        if (!parsed) parsed = rawOutput || String(raw || '');
        rawJsonText = rawOutput || JSON.stringify(raw || '');

        let name = templateName || templateId || '分析';
        let lines = [];

        // 如果解析出的是完整的模板输出（包含 success/result），从中取 result
        if (typeof parsed === 'object' && parsed !== null) {
            if (parsed.success === false) {
                return '❌ ' + name + '：' + (parsed.message || '执行失败');
            }

            // 有 parsed.result，取它作为模板特定格式化的数据源
            let r = parsed.result;
            if (r !== undefined && r !== null) {
                if (parsed.summary) {
                    lines.push('**' + name + '**：' + parsed.summary);
                    lines.push('');
                }
                lines.push(templateResultToMarkdown(r));
                lines.push('');
                lines.push('<details><summary>查看原始 JSON</summary>\n\n```json\n' + JSON.stringify(r, null, 2) + '\n```\n\n</details>');
            } else if (parsed.data !== undefined) {
                // 有些脚本把结果放在 data 字段
                if (parsed.summary) {
                    lines.push('**' + name + '**：' + parsed.summary);
                }
                if (r !== undefined && r !== null && typeof parsed.data === 'object') {
                    lines.push(templateResultToMarkdown(parsed.data));
                } else {
                    lines.push(String(parsed.data));
                }
            } else if (parsed.summary) {
                lines.push('**' + name + '**：' + parsed.summary);
            } else {
                // 直接输出：对象 → 友好键值，否则 json 折叠
                if (typeof parsed === 'object') {
                    lines.push(templateResultToMarkdown(parsed));
                } else {
                    lines.push('```json');
                    lines.push(JSON.stringify(parsed, null, 2));
                    lines.push('```');
                }
            }
        } else {
            lines.push(String(parsed));
        }

        let imagePaths = extractImagePaths(rawJsonText);
        if (imagePaths.length > 0) {
            lines.push('');
            lines.push('**生成的图表：**');
            for (let pi = 0; pi < imagePaths.length; pi++) {
                lines.push('![分析图表](' + chartImageUrl(imagePaths[pi]) + ')');
            }
        }

        return lines.join('\n');
    }

    // ============================================================
    // 自定义 Python 生成（旧模式，仅作 fallback）
    // ============================================================

    let CUSTOM_PYTHON_PROMPT = '你是 SmartPilot 的 Python 数据分析专家。\n\n' +
        '【任务】\n' +
        '根据用户的 Excel 数据和问题，生成可直接执行的 Python 脚本来完成数据分析任务。\n\n' +
        '【数据访问规则 - 极其重要】\n' +
        '- 数据通过 input_data 变量注入（list[dict]），这是唯一的数据入口\n' +
        '- 必须使用 input_data，不允许使用 df、data、frame 等其他变量名\n' +
        '- 示例：df = pd.DataFrame(input_data) 然后用 df 做分析\n' +
        '- 当前选区信息通过 selection_data 字典注入（含 workbookName, sheetName, address, headers, rows）\n' +
        '- 引用区域数据通过 reference_data 列表注入\n' +
        '- 禁止通过文件路径读取 Excel 文件\n' +
        '- 绝对禁止在脚本中定义或赋值以下变量（它们由系统注入真实数据）：\n' +
        '  input_data（真实数据）、selection_data（当前选区字典）、reference_data（引用区域列表）、\n' +
        '  input_headers（表头）、_sel（选区安全访问助手）。若你伪造赋值，将覆盖真实数据导致分析错误。\n' +
        '  直接使用这些变量即可，它们已包含用户选中的真实数据。\n\n' +
        '【列名强制规则 - 极其重要】\n' +
        '- 脚本中所有列名必须来自 selection_data["headers"] 或下方列出的实际列名\n' +
        '- 禁止使用任何不在实际数据中的列名（如 region/category/value 等占位词）\n' +
        '- 用 pd.DataFrame(input_data).columns.tolist() 获取实际列名\n' +
        '- 如果用户说"按地区分组"，必须使用数据中实际存在的列名\n\n' +
        '【输出要求】\n' +
        '1. 仅输出 Python 脚本代码，不要解释、不要 Markdown 代码块标记。\n' +
        '2. 脚本必须可直接执行。\n' +
        '3. 分析结论、关键统计量等请用 print() 输出。\n' +
        '4. 如需返回结构化结果，赋值给 output 变量。\n' +
        '5. 已预置库：pandas(pd)、numpy(np)、scipy.stats、matplotlib.pyplot(plt)，无需手动 import。\n' +
        '6. 脚本应进行防御性处理：处理空数据、非数值列、缺失值等边界情况。\n\n' +
        '【图表生成规则】\n' +
        '- 使用 _save_chart(filename) 保存图表，它会自动保存并输出路径。\n' +
        '- 已有内置中文字体支持，图表标题/标签可直接写中文。';

    // ============================================================
    // 主流程
    // ============================================================

    async function* execute(turn, signal) {
        let sessionStore = App.ChatV2.sessionStore;
        let toolRegistry = App.ChatV2.toolRegistry;
        let llmClient = App.ChatV2.llmClient;

        if (!sessionStore || !llmClient || !toolRegistry) {
            throw new Error('python-pipeline 依赖未就绪');
        }

        try {
            await toolRegistry.init();
            ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] turn=' + turn.id + ' 开始');

            // 生命周期（Phase 4）：新任务（非后续消息）时清空上一轮执行上下文，
            // 防止跨任务误继承（P6）。后续消息（_execFollowUp）保留上下文供继承。
            if (!turn.context?._execFollowUp && sessionStore.clearExecSpec) {
                sessionStore.clearExecSpec();
            }

            // ========== 预读数据：确保 turn.context.selection/referenceRegions 有真实 rows ==========
            // context-engine 可能只返回了 headers/address 而没实际 rows，导致 Python 脚本 input_data=[]
            // 这里主动调 excel_read_range 读取真实数据，注入上下文供后续所有步骤使用
            yield { type: 'status', text: '正在读取 Excel 数据...' };
            try {
                await ensureSelectionLoaded(turn, toolRegistry, signal);
            } catch (loadErr) {
                ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 预读数据失败，继续用现有上下文:', loadErr.message);
            }

            // 后续消息继承：无新选区时继承上一轮数据快照（元数据，供 LLM 理解上下文）
            const followUpSpec = getPythonFollowUp(turn);
            if (followUpSpec && followUpSpec.data) {
                const ctx2 = turn.context || {};
                const pd = followUpSpec.data;
                if (!ctx2.selection && pd.selection) ctx2.selection = pd.selection;
                if (!ctx2.referenceRegions && pd.referenceRegions) ctx2.referenceRegions = pd.referenceRegions;
            }

            let userMessage = (turn.userMessage || '').toLowerCase();
            // 去掉 \b 词边界：\b 依赖 \w（字母数字），对中文（如"脚本"后接中文）不可靠，
            // 导致"写个 python 脚本分析"这类消息无法命中自定义模式。改为纯文本包含匹配。
            let isCustomScript = shouldForceCustomInput(userMessage);

            yield { type: 'round_start', round: 1, totalRounds: 1 };

            let pythonToolName = resolvePythonToolName(toolRegistry);
            if (!toolRegistry.get(pythonToolName)) {
                throw new Error('未找到 Python 执行工具 (' + pythonToolName + ')');
            }

            let allScripts = [];
            let allResults = [];  // 收集执行结果，用于最终总结
            let isCustomMode = false;
            let lastAnalysisSummary = '';  // 最近一步的自然语言总结（写入 _execSpec 供后续继承）

            if (isCustomScript) {
                isCustomMode = true;
                yield { type: 'status', text: '正在生成 Python 脚本...' };

                let customMessages = buildCustomMessages(turn);
                let generatedScript = '';
                let tokenReceived = false;

                for await (let evt of llmClient.stream({ messages: customMessages, config: App.config, signal: signal })) {
                    if (signal && signal.aborted) { yield { type: 'canceled', reason: 'user_aborted' }; return; }
                    if (evt.type === 'token') {
                        generatedScript += evt.delta || '';
                        tokenReceived = true;
                        yield evt;
                    } else if (evt.type === 'reasoning') {
                        sessionStore.appendDelta(turn.id, 'reasoning', evt.delta || '');
                        yield evt;
                    } else if (evt.type === 'model') {
                        sessionStore.setTurnModel(turn.id, evt.model);
                        yield evt;
                    } else if (evt.type === 'error') {
                        yield evt; return;
                    } else {
                        yield evt;
                    }
                }

                let script = extractPythonCode(generatedScript);
                if (!script) {
                    if (!tokenReceived) throw new Error('模型未返回任何内容');
                    throw new Error('未能从 AI 响应中提取到可执行的 Python 脚本');
                }
                allScripts.push({ script: script, templateName: 'custom', tpl: null });
            } else {
                yield { type: 'status', text: '正在选择分析方案...' };

                let selection;
                const directMatch = matchTemplateByIntent(userMessage);
                if (directMatch) {
                    // 确定性模板绑定：命中明确意图直接定模板，跳过 LLM 二次选择，
                    // 避免 qwen 从模板列表选模板失败时在下方降级回 custom。
                    selection = { templates: [directMatch] };
                } else {
                    let selectorMessages = buildSelectorMessages(turn);
                    let selectorOutput = '';
                    let selTokenReceived = false;
                    for await (let evt2 of llmClient.stream({ messages: selectorMessages, config: App.config, signal: signal })) {
                        if (signal && signal.aborted) { yield { type: 'canceled', reason: 'user_aborted' }; return; }
                        if (evt2.type === 'token') {
                            selectorOutput += evt2.delta || '';
                            selTokenReceived = true;
                            yield evt2;
                        } else if (evt2.type === 'reasoning') {
                            sessionStore.appendDelta(turn.id, 'reasoning', evt2.delta || '');
                            yield evt2;
                        } else if (evt2.type === 'model') {
                            sessionStore.setTurnModel(turn.id, evt2.model);
                            yield evt2;
                        } else if (evt2.type === 'error') {
                            yield evt2; return;
                        } else {
                            yield evt2;
                        }
                    }
                    selection = extractJSON(selectorOutput);
                }
                let normalized = normalizeSelection(selection);

                if (!normalized || !normalized.templates || normalized.templates.length === 0) {
                    ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] JSON 解析失败，降级为 custom');
                    yield { type: 'status', text: '方案识别失败，尝试自定义生成...' };

                    let fallbackMessages = buildCustomMessages(turn);
                    let fallbackScript = '';

                    // 【修复-上下文组合】生成阶段不渲染折叠面板（脚本由执行阶段统一渲染），
                    // 避免与执行阶段的 "N. 名称 + Python 脚本" 面板重复组合
                    for await (let evtFb of llmClient.stream({ messages: fallbackMessages, config: App.config, signal: signal })) {
                        if (signal && signal.aborted) { yield { type: 'canceled', reason: 'user_aborted' }; return; }
                        if (evtFb.type === 'token') {
                            fallbackScript += evtFb.delta || '';
                            yield evtFb;
                        } else if (evtFb.type === 'reasoning') {
                            sessionStore.appendDelta(turn.id, 'reasoning', evtFb.delta || '');
                            yield evtFb;
                        } else if (evtFb.type === 'model') {
                            sessionStore.setTurnModel(turn.id, evtFb.model);
                            yield evtFb;
                        } else if (evtFb.type === 'error') {
                            yield evtFb; return;
                        } else {
                            yield evtFb;
                        }
                    }

                    let fbScript = extractPythonCode(fallbackScript);
                    if (!fbScript) throw new Error('未能生成 Python 脚本');
                    allScripts.push({ script: fbScript, templateName: 'custom', tpl: null });
                    isCustomMode = true;
                } else if (normalized.templates[0].id === 'custom') {
                    isCustomMode = true;
                    yield { type: 'status', text: '正在生成自定义 Python 脚本...' };

                    // 【修复-上下文组合】同降级分支：生成阶段不渲染折叠面板，脚本由执行阶段统一渲染
                    let customMessages2 = buildCustomMessages(turn);
                    let generatedScript2 = '';

                    for await (let evt3 of llmClient.stream({ messages: customMessages2, config: App.config, signal: signal })) {
                        if (signal && signal.aborted) { yield { type: 'canceled', reason: 'user_aborted' }; return; }
                        if (evt3.type === 'token') {
                            generatedScript2 += evt3.delta || '';
                            yield evt3;
                        } else if (evt3.type === 'reasoning') {
                            sessionStore.appendDelta(turn.id, 'reasoning', evt3.delta || '');
                            yield evt3;
                        } else if (evt3.type === 'model') {
                            sessionStore.setTurnModel(turn.id, evt3.model);
                            yield evt3;
                        } else if (evt3.type === 'error') {
                            yield evt3; return;
                        } else {
                            yield evt3;
                        }
                    }

                    let script2 = extractPythonCode(generatedScript2);
                    if (!script2) throw new Error('未能生成 Python 脚本');
                    allScripts.push({ script: script2, templateName: 'custom', tpl: null });
                } else {
                    let templatesList = normalized.templates;

                    for (let ti = 0; ti < templatesList.length; ti++) {
                        let tEntry = templatesList[ti];
                        let tId = tEntry.id;
                        let tParams = tEntry.params || {};
                        let tpl = TEMPLATES[tId];

                        if (!tpl) {
                            // 尝试模糊匹配模板 ID
                            let keys = Object.keys(TEMPLATES);
                            let bestMatch = null;
                            let bestScore = 0;
                            for (let ki = 0; ki < keys.length; ki++) {
                                let score = fuzzyMatch(tId.toLowerCase(), keys[ki].toLowerCase());
                                if (score > bestScore) {
                                    bestScore = score;
                                    bestMatch = keys[ki];
                                }
                            }
                            if (bestMatch && bestScore >= 0.6) {
                                ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 模糊匹配模板: ' + tId + ' → ' + bestMatch + ' (score=' + bestScore.toFixed(2) + ')');
                                tId = bestMatch;
                                tpl = TEMPLATES[tId];
                            } else {
                                ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 未知模板: ' + tId + '，跳过');
                                continue;
                            }
                        }

                        let filledParams = {};
                        let pKeys = Object.keys(tParams);
                        for (let pk = 0; pk < pKeys.length; pk++) {
                            filledParams[pKeys[pk]] = tParams[pKeys[pk]];
                        }

                        // 容错：自动修正参数名不匹配的常见错误
                        let tplParamKeys = Object.keys(tpl.params);
                        let tplParamKeysLower = {};
                        for (let tkl = 0; tkl < tplParamKeys.length; tkl++) {
                            tplParamKeysLower[tplParamKeys[tkl].toLowerCase()] = tplParamKeys[tkl];
                        }

                        // 构建常见错误名→正确名的映射
                        let PARAM_ALIASES = {
                            'dimcolumn': 'group_column',
                            'dim_columns': 'group_column',
                            'dimension': 'group_column',
                            'groupby': 'group_column',
                            'group_by': 'group_column',
                            'column': 'value_column',
                            'col': 'value_column',
                            'field': 'value_column',
                            'measure': 'value_column',
                            'metric': 'value_column',
                            'value': 'value_column',
                            'target_column': 'value_column',
                            'targetcol': 'value_column',
                            'valuecol': 'value_column',
                            'columns': 'columns',
                            'cols': 'columns',
                            'fields': 'columns',
                            'field_list': 'columns',
                            'column_list': 'columns',
                            'aggregation_method': 'aggregation',
                            'agg': 'aggregation',
                            'func': 'aggregation'
                        };

                        let currentKeys = Object.keys(filledParams);
                        for (let ck = 0; ck < currentKeys.length; ck++) {
                            let ckLower = currentKeys[ck].toLowerCase();
                            let correctName = null;
                            // 先检查精确的大小写不敏感匹配
                            if (tplParamKeysLower[ckLower]) {
                                correctName = tplParamKeysLower[ckLower];
                            }
                            // 再检查别名映射
                            if (!correctName && PARAM_ALIASES[ckLower] && tplParamKeysLower[PARAM_ALIASES[ckLower]]) {
                                correctName = PARAM_ALIASES[ckLower];
                            }
                            if (correctName && currentKeys[ck] !== correctName) {
                                filledParams[correctName] = filledParams[currentKeys[ck]];
                                delete filledParams[currentKeys[ck]];
                                ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 自动修正参数: ' + currentKeys[ck] + ' → ' + correctName);
                            }
                        }

                        // 额外智能补全：如果只有一个数值列参数缺失，从已提供的参数中推断
                        if (tplParamKeysLower['value_column'] && filledParams.value_column === undefined) {
                            let hasColParam = false;
                            for (let k in filledParams) {
                                if (k !== 'group_column' && typeof filledParams[k] === 'string' && filledParams[k].length > 0) {
                                    filledParams.value_column = filledParams[k];
                                    hasColParam = true;
                                    ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 智能补全 value_column = ' + filledParams[k]);
                                    break;
                                }
                            }
                        }

                        // 容错：columns → value_column（取数组第一个元素）
                        if (filledParams.columns !== undefined && filledParams.value_column === undefined && tplParamKeysLower['value_column']) {
                            let cv2 = filledParams.columns;
                            filledParams.value_column = Array.isArray(cv2) ? cv2[0] : cv2;
                        }
                        // 容错：value_column → columns（包装成数组）
                        if (filledParams.value_column !== undefined && filledParams.columns === undefined && tplParamKeysLower['columns']) {
                            let vc2 = filledParams.value_column;
                            filledParams.columns = Array.isArray(vc2) ? vc2 : [vc2];
                        }
                        // 容错：column → columns（单数→复数）
                        if (filledParams.column !== undefined && filledParams.columns === undefined && tplParamKeysLower['columns']) {
                            let cl2 = filledParams.column;
                            filledParams.columns = Array.isArray(cl2) ? cl2 : [cl2];
                        }
                        // ============ 新增：列名校验 + 自动兜底（防止 LLM 编造不存在的列名） ============
                        // 1) 收集真实可用列名
                        let realHeaders = [];
                        let hCtx = turn.context || {};
                        if (hCtx.selection && Array.isArray(hCtx.selection.headers)) {
                            realHeaders = hCtx.selection.headers.slice();
                        } else if (Array.isArray(hCtx.referenceRegions)) {
                            for (let rhi = 0; rhi < hCtx.referenceRegions.length; rhi++) {
                                let rh = hCtx.referenceRegions[rhi];
                                if (rh && Array.isArray(rh.headers)) realHeaders = realHeaders.concat(rh.headers);
                            }
                        }
                        // 去重
                        let _hset = {}; realHeaders = realHeaders.filter(function (x) { if (x == null) return false; let s = String(x); if (_hset[s]) return false; _hset[s] = true; return true; });
                        let realHeadersLower = realHeaders.map(function (h) { return String(h).toLowerCase(); });

                        // 2) 真实数值列（从 rows 推断），用于兜底时选默认值列
                        let numericHeaders = [];
                        try {
                            let sampleRows = null;
                            if (hCtx.selection && Array.isArray(hCtx.selection.rows) && hCtx.selection.rows.length > 0) {
                                sampleRows = hCtx.selection.rows;
                            } else if (Array.isArray(hCtx.referenceRegions) && hCtx.referenceRegions.length > 0) {
                                for (let ri = 0; ri < hCtx.referenceRegions.length; ri++) {
                                    if (hCtx.referenceRegions[ri] && Array.isArray(hCtx.referenceRegions[ri].rows) && hCtx.referenceRegions[ri].rows.length > 0) {
                                        sampleRows = hCtx.referenceRegions[ri].rows; break;
                                    }
                                }
                            }
                            if (sampleRows && sampleRows.length > 0 && realHeaders.length > 0) {
                                let numericVote = {};
                                let voteCount = Math.min(sampleRows.length, 20);
                                for (let vi = 0; vi < voteCount; vi++) {
                                    let vrow = sampleRows[vi];
                                    if (Array.isArray(vrow)) {
                                        for (let ci = 0; ci < Math.min(vrow.length, realHeaders.length); ci++) {
                                            let cv = vrow[ci];
                                            if (cv === null || cv === undefined || cv === '') continue;
                                            if (typeof cv === 'number' || !isNaN(parseFloat(String(cv).replace(/[,\s¥%￥]/g, '')))) {
                                                numericVote[realHeaders[ci]] = (numericVote[realHeaders[ci]] || 0) + 1;
                                            }
                                        }
                                    } else if (vrow && typeof vrow === 'object') {
                                        for (let hni = 0; hni < realHeaders.length; hni++) {
                                            let hn = realHeaders[hni];
                                            let cv2 = vrow[hn];
                                            if (cv2 === null || cv2 === undefined || cv2 === '') continue;
                                            if (typeof cv2 === 'number' || !isNaN(parseFloat(String(cv2).replace(/[,\s¥%￥]/g, '')))) {
                                                numericVote[hn] = (numericVote[hn] || 0) + 1;
                                            }
                                        }
                                    }
                                }
                                numericHeaders = Object.keys(numericVote).filter(function (h) { return numericVote[h] >= Math.ceil(voteCount / 2); });
                            }
                        } catch (nhErr) {
                            numericHeaders = [];
                        }
                        // 兜底：没推断出就全列当作候选
                        if (numericHeaders.length === 0) numericHeaders = realHeaders.slice();
                        // 日期/时间列即使以序列号(数值形态)投票入围，也剔除——不能当 value_column 做数值聚合
                        numericHeaders = numericHeaders.filter(function (h) { return !DATE_COL_RE.test(String(h)); });
                        if (numericHeaders.length === 0 && realHeaders.length) {
                            numericHeaders = realHeaders.filter(function (h) { return !DATE_COL_RE.test(String(h)); });
                        }

                        // 自动填充 columns：模板要求数值列数组但 LLM 未提供时，用推断出的数值列兜底。
                        // 支持"分析/总结"等免参数请求走确定性模板，避免缺参被跳过或降级 custom。
                        if (tplParamKeysLower['columns'] && filledParams.columns === undefined) {
                            filledParams.columns = numericHeaders.slice(0, Math.min(3, numericHeaders.length));
                            ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 自动填充 columns = ' + (filledParams.columns || []).join(', '));
                        }

                        // 3) 修正 columns / value_column / group_column：不存在于真实列名时 → 自动替换
                        function pickFirstExisting(candidates, mustBeNumeric) {
                            let pool = mustBeNumeric ? numericHeaders : realHeaders;
                            for (let i = 0; i < candidates.length; i++) {
                                let idx = realHeadersLower.indexOf(String(candidates[i]).toLowerCase());
                                if (idx >= 0) return realHeaders[idx];
                            }
                            return pool.length > 0 ? pool[0] : '';
                        }

                        function isValidHeader(col, mustBeNumeric) {
                            if (!col) return false;
                            let idx = realHeadersLower.indexOf(String(col).toLowerCase());
                            if (idx < 0) return false;
                            if (mustBeNumeric) {
                                return numericHeaders.some(function (nh) { return String(nh).toLowerCase() === String(col).toLowerCase(); });
                            }
                            return true;
                        }

                        // 预校验：在校正前记录 LLM 填写的原始列引用是否无效（用于显式告警，不只静默替换）
                        let preCheck = validateTemplateParams(tId, filledParams, realHeaders, numericHeaders);
                        let colWarnings = preCheck.invalidColumns || [];

                        let columnarParams = ['value_column', 'columns', 'group_column', 'time_column'];
                        for (let cpi = 0; cpi < columnarParams.length; cpi++) {
                            let cpName = columnarParams[cpi];
                            if (!tplParamKeysLower[cpName.toLowerCase()]) continue;  // 模板不含此参数跳过
                            let isNumericCol = (cpName === 'value_column' || cpName === 'columns');
                            let val = filledParams[cpName];

                            if (cpName === 'columns') {
                                let arr = Array.isArray(val) ? val.slice() : (val ? [val] : []);
                                let correctedArr = [];
                                for (let ai = 0; ai < arr.length; ai++) {
                                    if (isValidHeader(arr[ai], isNumericCol)) {
                                        correctedArr.push(pickFirstExisting([arr[ai]], isNumericCol));
                                    }
                                }
                                if (correctedArr.length === 0) {
                                    correctedArr = numericHeaders.slice(0, Math.min(3, numericHeaders.length));
                                    ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] columns 参数列无效，回退默认列: ' + correctedArr.join(', '));
                                }
                                filledParams.columns = correctedArr;
                            } else {
                                if (!isValidHeader(val, isNumericCol)) {
                                    let fixed = pickFirstExisting(val ? [val] : [], isNumericCol);
                                    ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] ' + cpName + ' 列不存在或非数值(' + (val || '(空)') + ')，修正为: ' + fixed);
                                    filledParams[cpName] = fixed;
                                } else {
                                    filledParams[cpName] = pickFirstExisting([val], isNumericCol); // 保留大小写精确匹配
                                }
                            }
                        }
                        // ============ 列名校验结束 ============

                        // 分组列兜底：若 LLM 把 group_column 填成日期/时间列（每日唯一→每组1条→"数据不足"），
                        // 修正为第一个非日期列（部门/地区等分类列），避免“分组聚合”产出空结论。
                        if (tplParamKeysLower['group_column'] && filledParams.group_column != null && String(filledParams.group_column) !== '') {
                            const gCol = String(filledParams.group_column);
                            if (DATE_COL_RE.test(gCol)) {
                                const categoryFallback = realHeaders.filter(function (g) { return !DATE_COL_RE.test(String(g)); });
                                if (categoryFallback.length) {
                                    const oldG = filledParams.group_column;
                                    filledParams.group_column = categoryFallback[0];
                                    ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] group_column 为日期列(' + oldG + ')，修正为分类列: ' + filledParams.group_column);
                                }
                            }
                        }

                        // 分组列智能兜底（自动补列 + 单值列修正）：
                        // 1) 免 LLM 选择器直接绑定模板时 group_column 为空 → 选"取值最丰富"的分类列；
                        // 2) 分组列在实际数据中仅有 ≤1 个取值（如整列全为"华南"）→ 按它聚合退化为单组、
                        //    结论无意义，自动改选取值更丰富的分类列（部门/产品/销售员等）。
                        //    与上方日期列修正同理：选了退化分组列就换掉。用户显式指定的分组列除外。
                        // 3) value_column 为空 → 用推断出的首个数值列（销售额/利润等）兜底，避免必填校验跳过模板。
                        if (tplParamKeysLower['group_column'] || tplParamKeysLower['value_column']) {
                            const gCol2 = (filledParams.group_column == null || String(filledParams.group_column) === '') ? '' : String(filledParams.group_column);
                            const gRows2 = [];
                            if (hCtx.selection && Array.isArray(hCtx.selection.rows)) {
                                for (let gr2 = 0; gr2 < hCtx.selection.rows.length; gr2++) gRows2.push(hCtx.selection.rows[gr2]);
                            } else if (Array.isArray(hCtx.referenceRegions)) {
                                for (let gr3 = 0; gr3 < hCtx.referenceRegions.length; gr3++) {
                                    const rr2 = hCtx.referenceRegions[gr3];
                                    if (rr2 && Array.isArray(rr2.rows)) {
                                        for (let gr4 = 0; gr4 < rr2.rows.length; gr4++) gRows2.push(rr2.rows[gr4]);
                                    }
                                }
                            }
                            const distinctCountOf = function (colName, maxCap) {
                                const seen2 = {};
                                let cnt2 = 0;
                                const cap = (maxCap && maxCap > 0) ? maxCap : 1000000;
                                const lowerIdx = realHeadersLower.indexOf(String(colName).toLowerCase());
                                for (let ri2 = 0; ri2 < gRows2.length; ri2++) {
                                    const row2 = gRows2[ri2];
                                    let v2 = Array.isArray(row2) ? (lowerIdx >= 0 ? row2[lowerIdx] : null) : (row2 ? row2[colName] : null);
                                    if (v2 === null || v2 === undefined) v2 = '';
                                    const key2 = String(v2).trim();
                                    if (!key2 || seen2[key2]) continue;
                                    seen2[key2] = true;
                                    cnt2++;
                                    if (cnt2 >= cap) return cnt2;
                                }
                                return cnt2;
                            };
                            const isNumericCol2 = function (colName) {
                                const low = String(colName).toLowerCase();
                                for (let ni = 0; ni < numericHeaders.length; ni++) {
                                    if (String(numericHeaders[ni]).toLowerCase() === low) return true;
                                }
                                return false;
                            };
                            const bestCategorical2 = function (excludeCol) {
                                let b2 = '', bc2 = 0;
                                for (let ci2 = 0; ci2 < realHeaders.length; ci2++) {
                                    const cand2 = String(realHeaders[ci2]);
                                    // 分组列必须是分类列：排除日期列与数值列（销售额/成本/利润等）
                                    if (DATE_COL_RE.test(cand2) || isNumericCol2(cand2)) continue;
                                    if (excludeCol && cand2.toLowerCase() === String(excludeCol).toLowerCase()) continue;
                                    // 选"最丰富"列用完整去重计数（上限50），低基数判断用 cap=2 提前退出
                                    const cc2 = distinctCountOf(cand2, 50);
                                    if (cc2 > bc2) { bc2 = cc2; b2 = cand2; }
                                }
                                return { col: b2, cnt: bc2 };
                            };

                            if (tplParamKeysLower['value_column'] && (filledParams.value_column === undefined || filledParams.value_column === '')) {
                                filledParams.value_column = numericHeaders.length ? numericHeaders[0] : (realHeaders.length ? realHeaders[0] : '');
                                ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 自动填充 value_column = ' + filledParams.value_column);
                            }

                            if (tplParamKeysLower['group_column'] && (gCol2 === '' || (gRows2.length > 1 && distinctCountOf(gCol2, 2) < 2))) {
                                if (gCol2 === '') {
                                    const fb2 = bestCategorical2('');
                                    filledParams.group_column = fb2.col || (realHeaders.filter(function (g) { return !DATE_COL_RE.test(String(g)); })[0] || realHeaders[0] || '');
                                    ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 自动填充 group_column = ' + filledParams.group_column);
                                } else if (String(userMessage || '').indexOf(gCol2) === -1) {
                                    const fb3 = bestCategorical2(gCol2);
                                    if (fb3.col && fb3.cnt >= 2) {
                                        const oldG2 = filledParams.group_column;
                                        filledParams.group_column = fb3.col;
                                        ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] group_column "' + oldG2 + '" 仅有 1 个取值（聚合退化为单组），修正为分类列: ' + fb3.col + '（' + fb3.cnt + ' 个取值）');
                                    }
                                }
                            }
                        }

                        // 显式告警：把"静默替换坏列"升级为带可用列清单的提示，避免分析跑在错误列上误导用户
                        if (colWarnings.length > 0) {
                            let sug = buildColumnSuggestions(colWarnings, realHeaders, numericHeaders);
                            ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 参数校验: 模板 "' + tId + '" 引用了 ' + colWarnings.length + ' 个不存在的列，已自动修正。\n' + sug);
                        }

                        // 自动填补时间列：趋势类模板在 LLM 未指定 time_column 时，若数据含日期/时间列则自动识别并绑定
                        if (tplParamKeysLower['time_column'] && (filledParams.time_column === undefined || filledParams.time_column === '')) {
                            let dateRows = [];
                            if (hCtx.selection && Array.isArray(hCtx.selection.rows)) {
                                dateRows = hCtx.selection.rows;
                            } else if (Array.isArray(hCtx.referenceRegions)) {
                                for (let dri = 0; dri < hCtx.referenceRegions.length; dri++) {
                                    if (hCtx.referenceRegions[dri] && Array.isArray(hCtx.referenceRegions[dri].rows)) {
                                        dateRows = dateRows.concat(hCtx.referenceRegions[dri].rows);
                                    }
                                }
                            }
                            let dt = detectDateColumn(realHeaders, dateRows);
                            if (dt) {
                                filledParams.time_column = dt;
                                ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 自动识别时间列: ' + dt + ' → time_column');
                            }
                        }

                        for (let pi = 0; pi < tplParamKeys.length; pi++) {
                            let pname = tplParamKeys[pi];
                            if (filledParams[pname] === undefined && tpl.params[pname].default !== undefined) {
                                filledParams[pname] = tpl.params[pname].default;
                            }
                        }

                        let allValid = true;
                        for (let pj = 0; pj < tplParamKeys.length; pj++) {
                            let pname2 = tplParamKeys[pj];
                            if (tpl.params[pname2].required &&
                                (filledParams[pname2] === undefined || filledParams[pname2] === null || filledParams[pname2] === '')) {
                                allValid = false;
                                ChatLog.warn && ChatLog.warn('PIPELINE', '[python-pipeline] 模板 "' + tpl.name + '" 缺少必填参数: ' + pname2 + '，跳过');
                                break;
                            }
                        }

                        if (!allValid) continue;

                        let genScript = tpl.generate(filledParams);
                        allScripts.push({
                            script: genScript,
                            templateName: tId,
                            tpl: tpl,
                            filledParams: filledParams
                        });

                        ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 模板: ' + tId + ', params:', filledParams);
                    }

                    if (allScripts.length === 0) {
                        let errDetails = '所有模板参数验证均失败。\n';
                        errDetails += 'LLM 输出: ' + JSON.stringify(normalized.templates, null, 2) + '\n';
                        errDetails += '可用模板: ' + Object.keys(TEMPLATES).join(', ') + '\n';
                        errDetails += '请尝试：用更明确的关键词提问（如"分析销售额的趋势"而非"分析销售额"）';
                        ChatLog.error && ChatLog.error('PIPELINE', '[python-pipeline] ' + errDetails);
                        throw new Error(errDetails);
                    }
                }
            }

            if (allScripts.length === 0) throw new Error('未能生成可执行的 Python 脚本');

            ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] 共生成 ' + allScripts.length + ' 个脚本');

            let overallSuccess = true;

            for (let idx = 0; idx < allScripts.length; idx++) {
                let entry = allScripts[idx];
                let displayName = entry.tpl ? entry.tpl.name : entry.templateName;
                let templateId = entry.templateName;

                // custom 模式：提前检测列名幻觉和变量错误（用于状态提示）
                let hallucinatedCols = [];
                let badVarNames = [];
                // preHeaders 需在块外声明：下方警告分支（hallucinatedCols 非空时）也会引用
                let preHeaders = [];
                if (isCustomMode) {
                    let preCtx = turn.context || {};
                    if (preCtx.selection && Array.isArray(preCtx.selection.headers)) {
                        preHeaders = preCtx.selection.headers;
                    } else if (Array.isArray(preCtx.referenceRegions) && preCtx.referenceRegions.length > 0) {
                        for (let pi = 0; pi < preCtx.referenceRegions.length; pi++) {
                            let prh = preCtx.referenceRegions[pi];
                            if (Array.isArray(prh.headers)) {
                                for (let pj = 0; pj < prh.headers.length; pj++) {
                                    if (preHeaders.indexOf(prh.headers[pj]) === -1) preHeaders.push(prh.headers[pj]);
                                }
                            }
                        }
                    }
                }

                let hasWarnings = (hallucinatedCols && hallucinatedCols.length > 0) || (badVarNames && badVarNames.length > 0);

                if (allScripts.length > 1) {
                    if (hasWarnings) {
                        yield { type: 'status', text: '⚠️ 检测到脚本问题，仍在执行 ' + displayName + ' (' + (idx + 1) + '/' + allScripts.length + ')...' };
                    } else {
                        yield { type: 'status', text: '正在执行 ' + displayName + ' (' + (idx + 1) + '/' + allScripts.length + ')...' };
                    }
                } else {
                    yield { type: 'status', text: '正在执行 Python 分析脚本...' };
                }

                // === 写入步骤标题和结果 ===
                if (allScripts.length > 1 && idx > 0) {
                    sessionStore.appendDelta(turn.id, 'content', '\n---\n\n');
                }

                // 步骤标题（始终显示）
                sessionStore.appendDelta(turn.id, 'content',
                    '### ' + (idx + 1) + '. ' + displayName + '\n\n'
                );

                // 参数摘要（简洁显示）
                if (!isCustomMode) {
                    let paramText = buildParamsText(entry.filledParams || {});
                    if (paramText) {
                        sessionStore.appendDelta(turn.id, 'content', paramText + '\n\n');
                    }
                }

                // === 调试区域（可折叠，标记包裹） ===
                let debugParts = [];
                if (!isCustomMode) {
                    debugParts.push('**参数**：' + buildParamsText(entry.filledParams || {}));
                }
                debugParts.push('```python\n' + entry.script + '\n```');
                let debugContent = debugParts.join('\n\n');
                sessionStore.appendDelta(turn.id, 'content',
                    '<!-- PY_COLLAPSIBLE type="debug" title="Python 脚本" -->\n' +
                    debugContent + '\n' +
                    '<!-- PY_COLLAPSIBLE_END -->\n\n'
                );

                let execArgs = buildExecArgs(entry.script, turn, !isCustomMode);

                // 列名幻觉/变量错误改为基于"剥离伪造注入变量后"的最终执行脚本判断：
                // LLM 伪造的 input_data=[{...}] 字面量里的假列（如"客户数"）会被 buildExecArgs 移除，
                // 若在剥离前检测会误判为"引用了不存在的列"而拦截本可正常执行的脚本。
                if (isCustomMode) {
                    let issues = detectScriptIssues(execArgs.script || entry.script, preHeaders);
                    hallucinatedCols = issues.hallucinatedColumns;
                    badVarNames = issues.badVariableNames;
                }

                // 检测到问题时警告用户
                if (hallucinatedCols.length > 0 || badVarNames.length > 0) {
                    let warnParts = [];
                    if (hallucinatedCols.length > 0) {
                        warnParts.push('> ⚠️ **可疑列名**：`' + hallucinatedCols.slice(0, 3).join('`、`') + '`（可用列：`' + (preHeaders.length > 0 ? preHeaders.slice(0, 5).join('`、`') : '未知') + '`）');
                    }
                    if (badVarNames.length > 0) {
                        warnParts.push('> ⚠️ **变量问题**：' + badVarNames.join('、'));
                    }
                    warnParts.push('> 脚本可能会执行失败，建议改用模板模式或检查代码\n');
                    sessionStore.appendDelta(turn.id, 'content', warnParts.join('\n') + '\n');
                }

                // 构建工具卡片展示信息（供 hidden 模式下的历史记录）
                let displayInfo;
                if (isCustomMode) {
                    displayInfo = { '执行模式': '自定义脚本', '脚本长度': entry.script.length + ' 字符' };
                    if (hallucinatedCols.length > 0) {
                        displayInfo['⚠️ 检测到可疑列名'] = hallucinatedCols.slice(0, 5).join('、') + '...';
                    }
                } else {
                    displayInfo = buildDisplayArgs(templateId, displayName, entry.filledParams || {});
                }

                // === 列名护栏：custom 分支命中坏列则中止，不回传执行、回传可用列清单 ===
                // detectScriptIssues 已识别 hallucinatedColumns（只警告）；此处升级为硬拦截，
                // 避免 qwen2.5:7b 等小模型自由生成时引用不存在的列直接打崩 Python 执行。
                let execResult = null;
                let guardedByColumnFilter = false;
                let placeholderColHit = /[\{<]\s*col[}>]|\b(xxx|占位|待填)\b|\[\s*'[^']*列[^']*'\s*\]/.test(execArgs.script || entry.script);
                if (isCustomMode && (hallucinatedCols.length > 0 || placeholderColHit)) {
                    guardedByColumnFilter = true;
                    const avail = preHeaders.length ? preHeaders : ['(未检测到选区列)'];
                    const guardMsg = columnGuardMessage(hallucinatedCols.length > 0 ? hallucinatedCols : ['占位符'], avail);
                    execResult = { ok: false, success: false, result: guardMsg, output: guardMsg, error: guardMsg };
                    yield { type: 'status', text: '检测到脚本引用了不存在的列，已阻止执行并回传可用列清单' };
                } else {
                    execResult = await toolRegistry.invokeTracked(turn, pythonToolName, execArgs, {
                        round: 1,
                        displayArgs: displayInfo || {},
                        hidden: true,
                        ctx: { turn: turn, signal: signal }
                    });
                }

                let execSuccess = isPythonExecSuccess(execResult);
                // 错误分类：区分框架错误、模板错误、Skill/LLM 生成错误、列名护栏
                let errorClassification = null;
                if (!execSuccess) {
                    if (guardedByColumnFilter) {
                        errorClassification = {
                            errorType: '列名引用错误',
                            suggestion: columnGuardMessage(hallucinatedCols.length > 0 ? hallucinatedCols : ['占位符'], preHeaders.length ? preHeaders : ['(未检测到选区列)'])
                        };
                    } else {
                        errorClassification = classifyExecutionError(execResult, isCustomMode, displayName);
                    }
                }
                // hidden 卡片静默更新
                if (execResult.toolCallId) {
                    let patch = {
                        status: execSuccess ? 'completed' : 'failed',
                        resultSummary: execSuccess ? '分析完成' : getPythonError(execResult),
                        error: execSuccess ? null : getPythonError(execResult)
                    };
                    if (errorClassification) {
                        patch.errorType = errorClassification.errorType;
                        patch.errorReason = errorClassification.suggestion;
                        patch.errorSuggestion = errorClassification.suggestion;
                    }
                    sessionStore.updateToolCall(turn.id, execResult.toolCallId, patch, { silent: true });
                }

                if (!execSuccess) overallSuccess = false;

                let resultText;
                if (isCustomMode || !templateId) {
                    resultText = formatExecutionResult(execResult);
                } else {
                    resultText = formatAnalysisResult(execResult, templateId, displayName);
                }

                // 写入分析结果（简洁模式 + 可折叠 JSON）
                if (!execSuccess && errorClassification) {
                    sessionStore.appendDelta(turn.id, 'content',
                        '**❌ 执行失败**\n\n' +
                        '**错误类型**：' + errorClassification.errorType + '\n\n' +
                        '**建议**：' + errorClassification.suggestion + '\n\n' +
                        '<!-- PY_COLLAPSIBLE type="error" title="原始错误详情" -->\n' +
                        '```\n' + (getPythonError(execResult) || resultText) + '\n```\n' +
                        '<!-- PY_COLLAPSIBLE_END -->\n'
                    );
                } else {
                    // 生成自然语言总结（与 WPS AI 同类体验）
                    yield { type: 'status', text: '正在生成分析总结...' };

                    let analysisSummary = '';
                    try {
                        analysisSummary = await generateAnalysisSummary(
                            execResult, templateId, displayName, llmClient, App.config, signal
                        );
                    } catch(e) {
                        ChatLog.warn && ChatLog.warn('PIPELINE', '[summary] 生成异常:', e.message);
                    }
                    if (analysisSummary) lastAnalysisSummary = analysisSummary;

                    // 如果没有生成总结，尝试从原始结果中提取 summary 字段
                    if (!analysisSummary) {
                        try {
                            let rawObj2 = typeof execResult?.result === 'string' ? JSON.parse(execResult.result) : execResult?.result;
                            if (rawObj2 && rawObj2.summary) {
                                analysisSummary = '**结果**：' + rawObj2.summary;
                            }
                        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
                    }

                    // 主输出：自然语言总结
                    if (analysisSummary) {
                        sessionStore.appendDelta(turn.id, 'content', analysisSummary + '\n\n');
                    } else {
                        // 降级：直接显示原始摘要
                        let summaryText2 = '';
                        try {
                            let rawObj = typeof execResult?.result === 'string' ? JSON.parse(execResult.result) : execResult?.result;
                            if (rawObj && rawObj.summary) {
                                summaryText2 = rawObj.summary;
                            }
                        } catch(e) { if (typeof ChatLog !== 'undefined') ChatLog.warn('PYTHON', '解析失败:', e && e.message); }
                        if (summaryText2) {
                            sessionStore.appendDelta(turn.id, 'content', '**结果**：' + summaryText2 + '\n\n');
                        }
                    }

                    // 技术细节（可折叠）：原始 JSON 结果
                    // ECharts 图表：以"可见"节点追加到总结之后（不藏在折叠面板/表格卡内，
                    // 否则 display:none 或 <details> 折叠下渲染为 0 宽高不可见）。
                    let resultChartDiv = '';
                    if (templateId && execResult && execResult.ok) {
                        try { resultChartDiv = buildResultChartDiv(templateId, execResult); }
                        catch(e) { ChatLog.warn && ChatLog.warn('PIPELINE', '[chart] 构建图表节点失败:', e.message); }
                    }
                    ChatLog.info && ChatLog.info('PYCHART', '[python-pipeline] 追加图表: templateId=' + templateId + ' execOk=' + (execResult && execResult.ok) + ' divLen=' + resultChartDiv.length + ' hasDiv=' + resultChartDiv.includes('data-echarts') + ' rawType=' + (execResult && typeof execResult.result) + ' rawPreview=' + JSON.stringify(String(execResult && execResult.result)).slice(0, 240));
                    if (resultChartDiv) {
                        sessionStore.appendDelta(turn.id, 'content', resultChartDiv);
                    }

                    sessionStore.appendDelta(turn.id, 'content',
                        '<!-- PY_COLLAPSIBLE type="result" title="分析结果详情" -->\n' +
                        resultText + '\n' +
                        '<!-- PY_COLLAPSIBLE_END -->\n'
                    );
                }
                // 收集执行结果用于最终总结
                if (execSuccess) {
                    allResults.push({
                        templateId: templateId,
                        displayName: displayName,
                        result: execResult
                    });
                }

                yield { type: 'token', delta: resultText };
            }

            // 生成最终综合总结（多个步骤时合并展示）
            // combinedSummary 必须在 if 块外声明：单步分析（allResults.length<=1）时
            // 下方 setExecSpec 仍会引用它，块内 let 声明会导致 ReferenceError（"combinedSummary is not defined"）。
            let combinedSummary = '';
            if (allResults.length > 1) {
                yield { type: 'status', text: '正在生成综合分析总结...' };

                try {
                    // 合并所有结果
                    let combinedText = '';
                    for (let ri = 0; ri < allResults.length; ri++) {
                        let r = allResults[ri];
                        let rText = '';
                        if (r.result.result) {
                            rText = typeof r.result.result === 'string' ? r.result.result : JSON.stringify(r.result.result, null, 2);
                        } else if (r.result.output) {
                            rText = typeof r.result.output === 'string' ? r.result.output : JSON.stringify(r.result.output);
                        }
                        combinedText += '【' + r.displayName + '】\n' + rText + '\n\n';
                    }

                    if (combinedText.length > 8000) {
                        combinedText = combinedText.substring(0, 8000) + '\n...';
                    }

                    let finalMsg = '你需要将以下多个分析步骤的结果合并为一份综合总结。\n\n' +
                        '要求：\n' +
                        '1. 使用中文\n' +
                        '2. 先给出各步骤的简要结论\n' +
                        '3. 最后给出整体结论和建议\n' +
                        '4. 使用 Markdown 格式，多行同构明细必须用表格呈现，适当使用表格\n' +
                        '5. 不要提及"Python"或"代码"\n' +
                        '6. 不要使用 h1/h2 大标题，改用 h3（### 小节名）或加粗文本\n\n' +
                        '分析结果：\n```\n' + combinedText + '\n```';

                    let combinedMessages = [
                        { role: 'system', content: '你是数据分析总结专家，擅长将多个分析步骤整合为一份清晰的综合报告。' },
                        { role: 'user', content: finalMsg }
                    ];

                    const _llmIo = App.ChatV2.llmIo;
                    if (_llmIo) {
                        const _sumR = await _llmIo.streamText({ messages: combinedMessages, config: App.config, signal });
                        combinedSummary = _sumR.text || '';
                    }
                } catch(e) {
                    ChatLog.warn && ChatLog.warn('PIPELINE', '[summary] 综合总结失败:', e.message);
                }

                if (combinedSummary) {
                    sessionStore.appendDelta(turn.id, 'content', '\n---\n\n### 📋 综合总结\n\n' + combinedSummary + '\n\n');
                }
            }

            // 完成提示
            if (overallSuccess) {
                let successCount = allResults.length;
                sessionStore.appendDelta(turn.id, 'content',
                    '✅ **分析完成！** 共完成 ' + successCount + ' 项分析\n'
                );
            } else {
                sessionStore.appendDelta(turn.id, 'content',
                    '⚠️ 部分分析未能完成（' + allResults.length + '/' + allScripts.length + ' 成功）\n'
                );
            }

            ChatLog.info && ChatLog.info('PIPELINE', '[python-pipeline] turn=' + turn.id + ' 完成 total=' + allScripts.length + ' ok=' + overallSuccess);

            // 写入通用执行上下文（多轮继承）：分析结论 + 数据快照 + 参数，供后续消息继承
            if (sessionStore.setExecSpec) {
                const fact = sessionStore.execSpecFactory;
                // v20260924：用工厂统一构建 data，selection/referenceRegions schema 不再手写
                const commonData = fact && typeof fact.buildCommonData === 'function'
                    ? fact.buildCommonData(turn)
                    : {}; // 兜底（旧 session-store 版本）
                const steps = allResults.map(function (r) {
                    return {
                        templateId: r.templateId,
                        displayName: r.displayName,
                        summary: extractResultSummary(r.result)
                    };
                });
                sessionStore.setExecSpec(turn.id, {
                    pipeline: 'python',
                    data: commonData,
                    params: {
                        mode: isCustomMode ? 'custom' : 'template',
                        templates: allScripts.map(function (s) { return s.templateName || 'custom'; })
                    },
                    result: {
                        success: overallSuccess,
                        summary: combinedSummary || lastAnalysisSummary || '',
                        steps: steps,
                        // 完整分析结果（结构化）：供 PPT 跨管道继承时直接引用统计指标/分布数据/结论文本，
                        // 避免 PPT 用原始数据重新组装
                        details: allResults.map(function (r) {
                            let detail = '';
                            try {
                                const obj = r.result && r.result.result;
                                if (obj) {
                                    detail = typeof obj === 'string' ? obj : JSON.stringify(obj, null, 2);
                                } else if (r.result && r.result.output) {
                                    detail = typeof r.result.output === 'string' ? r.result.output : JSON.stringify(r.result.output);
                                }
                            } catch (e) { detail = ''; }
                            return {
                                templateId: r.templateId,
                                displayName: r.displayName,
                                detail: String(detail).slice(0, 4000)
                            };
                        })
                    },
                    // 任务结构（分析步骤/任务清单）：供后续消息继承时注入规划输入
                    taskStructure: {
                        tasks: allScripts.map(function (s) {
                            return {
                                templateId: s.templateName || '',
                                displayName: (s.tpl && s.tpl.name) || s.templateName || '',
                                params: s.filledParams || null
                            };
                        })
                    },
                    ts: Date.now()
                });
            }

            sessionStore.completeTurn(turn.id, {});
            yield { type: 'pipeline_done', pipeline: 'python' };
        } catch (err) {
            ChatLog.error && ChatLog.error('PIPELINE', '[python-pipeline] error', err);
            sessionStore.failTurn(turn.id, { message: err && err.message ? err.message : String(err) });
            yield {
                type: 'error',
                error: { code: 'PIPELINE_ERROR', message: err && err.message ? err.message : String(err), recoverable: false }
            };
        }
    }

    // ============================================================
    // 分析总结生成：用 LLM 将 Python 结果转为自然语言总结
    // ============================================================

    let SUMMARY_PROMPT = '你是数据分析总结专家。\n\n' +
        '请根据以下 Python 数据分析结果，生成一份简洁、专业的分析总结。\n\n' +
        '【输出格式】\n' +
        '1. 使用中文\n' +
        '2. 使用 Markdown 格式\n' +
        '3. 不要使用 h1/h2 大标题（如 # 标题 或 ## 标题），改用 h3（### 小节名）或加粗文本作为标题\n' +
        '4. 【强制】凡涉及多条同构明细记录（同组字段名重复出现多条，如异常检测里"总数/异常值数量/占比"对每个指标各一条），必须统一用 Markdown 表格呈现（表头=字段，每行=一条记录），严禁拆成逐字段的文本列表排版\n' +
        '5. 【强制】数值列的列名必须用真实聚合名称或原列名，例如"平均值/求和/计数/最小值/最大值"，或直接用原列名（如"销售额"）；严禁用"指标"、"数值"、"value"这类笼统占位词。若结果含 agg 字段（avg/sum/count/min/max），列名就对应写"平均值/求和/计数/最小值/最大值"\n' +
        '6. 关键发现用加粗标出\n' +
        '7. 使用中文序号分段（一、二、三、）\n' +
        '8. 不要提及"Python"或"代码"，直接呈现分析结论\n\n' +
        '【示例格式】\n' +
        '```\n' +
        '**分析完成**\n\n' +
        '### 分组统计结果\n' +
        '| 分组 | 平均值 |\n' +
        '|------|--------|\n' +
        '| A    | 100    |\n' +
        '| B    | 50     |\n\n' +
        '### 关键发现\n' +
        '- **最高平均值**：A 组的 100\n' +
        '- **最低平均值**：B 组的 50\n\n' +
        '### 结论\n' +
        '数据表明...\n' +
        '```\n';

    /**
     * 调用 LLM 将 Python 执行结果转化为自然语言总结
     * @param {object} execResult - Python 执行结果
     * @param {string} templateId - 模板 ID
     * @param {string} displayName - 显示名称
     * @param {object} llmClient - LLM 客户端
     * @param {object} config - 配置
     * @param {object} signal - 中断信号
     * @returns {Promise<string>} 总结文本
     */
    async function generateAnalysisSummary(execResult, templateId, displayName, llmClient, config, signal) {
        try {
            // 先判断执行是否失败/无数据——这些情况不调 LLM，直接返回确定性文本，防止 LLM 编造
            let parsedExec = null;
            if (execResult && typeof execResult === 'object') {
                parsedExec = execResult;
            } else if (typeof execResult === 'string') {
                parsedExec = safeParse(execResult);
            }

            // 1. 明确失败：返回错误，不调 LLM
            if (parsedExec && parsedExec.success === false) {
                let msg = parsedExec.message || parsedExec.error || '脚本执行失败';
                return '**分析失败**\n\n> ' + String(msg).replace(/\n/g, '\n> ');
            }

            // 提取原始结果文本
            let rawText = '';
            if (execResult && execResult.result) {
                if (typeof execResult.result === 'string') {
                    rawText = execResult.result;
                } else {
                    rawText = JSON.stringify(execResult.result, null, 2);
                }
            } else if (execResult && execResult.output) {
                rawText = typeof execResult.output === 'string' ? execResult.output : JSON.stringify(execResult.output);
            }

            if (!rawText || rawText.trim().length < 5) {
                return '';
            }

            // 2. 无可用数据：返回确定性提示，不调 LLM
            let noDataPattern = /(无可用数据|no data|没有数据|input_data\s*为空|数据为空|empty data)/i;
            if (noDataPattern.test(rawText)) {
                return '**分析失败**\n\n> 未读取到有效数据。请先框选要分析的数据区域（需包含表头），或确认工作表中存在实际数据后重试。';
            }

            // 截断过长的结果（避免 token 超限）
            let truncatedText = rawText.length > 8000 ? rawText.substring(0, 8000) + '\n...（结果过长，已截断）' : rawText;

            // 3. 在 SUMMARY_PROMPT 中追加"严禁编造"指令，有证据才能输出
            let hardenedPrompt = SUMMARY_PROMPT + '\n\n【严禁编造】\n- 所有结论只能引用"执行结果"中明确出现的数字、字段和事实，执行结果里没有的内容绝对不得虚构（包括但不限于时间、金额、样本数、列名、百分比、日期区间、趋势描述）。\n- 如果执行结果信息不足、结构无法解读，请直接写"**数据不足，无法生成分析结论**"，不要自行假设或补充任何数值。\n- 严禁提及不存在的列（例如"销售额"、"收入"等除非结果里明确出现），严禁伪造样本数量。';

            let userMsg = '分析任务：' + (displayName || templateId || '数据分析') + '\n\n' +
                '执行结果：\n```\n' + truncatedText + '\n```\n\n' +
                '请生成分析总结。';

            let messages = [
                { role: 'system', content: hardenedPrompt },
                { role: 'user', content: userMsg }
            ];

            const _llmIo = App.ChatV2.llmIo;
            const _res = _llmIo ? await _llmIo.streamText({ messages: messages, config: config, signal: signal }) : { text: '' };
            return String(_res.text || '').trim();
        } catch (e) {
            // 总结生成失败不影响主流程
            ChatLog.warn && ChatLog.warn('PIPELINE', '[summary] 生成失败:', e.message);
            return '';
        }
    }

    function buildCustomMessages(turn) {
        let userMsg = turn.userMessage || '';
        let contextBlock = buildContextBlock(turn.context || {});
        let userContent = contextBlock
            ? userMsg + '\n\n' + contextBlock
            : userMsg + '\n\n（注：当前未提供具体选区数据，请在脚本中通过 Excel 接口读取活动工作表或当前选区，并基于实际数据完成分析。）';

        // 后续消息继承：注入上轮分析结论 + 任务规划 + 意图（Phase 3）
        const followUpSpec = getPythonFollowUp(turn);
        if (followUpSpec) {
            const prevCtx = buildFollowUpPrompt(followUpSpec.spec, turn.userMessage || '', followUpSpec.intent);
            if (prevCtx) userContent += '\n\n' + prevCtx;
        }

        let matchedSkill = turn.context && turn.context._matchedSkill;
        let systemPrompt = (matchedSkill && matchedSkill.systemPrompt) ? matchedSkill.systemPrompt : CUSTOM_PYTHON_PROMPT;

        return [
            { role: 'system', content: systemPrompt },
            { role: 'user', content: userContent }
        ];
    }

    // 优先走 pipelineRegistry 声明式注册
    if (App.ChatV2.pipelineRegistry && typeof App.ChatV2.pipelineRegistry.register === 'function') {
        App.ChatV2.pipelineRegistry.register('python', {
            execute,
            name: 'Python 数据分析',
            description: '数据分析（模板+参数确定性执行 / 自由生成 Python 脚本）',
            priority: 90,
            fallbackKind: 'text'
        });
    } else {
        App.ChatV2.pipelines.python = { execute: execute };
    }
    ChatLog.info && ChatLog.info('CHAT-V2', '✅ python-pipeline loaded (v3 - multi-template)');
})();