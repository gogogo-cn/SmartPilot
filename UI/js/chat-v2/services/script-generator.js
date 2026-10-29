/**
 * ============================================================
 * chat-v2/services/script-generator.js — Python 脚本生成（结构化参数）
 * ============================================================
 * 职责：从 text-pipeline 抽离的 generateAdvancedScript，改为
 *   "参数结构化传参"模式：所有用户输入（列名/筛选值）通过 JSON
 *   前置块注入，Python 代码从 _P 字典取值，杜绝字符串拼接注入。
 *
 *   对比旧实现（text-pipeline generateAdvancedScript）：
 *     旧：script += `df['${pyStr(col)}'] = ...`  ← 手写转义，易注入
 *     新：_P = json.loads('{...}')              ← JSON 序列化，安全
 * ============================================================
 */
(function() {
    'use strict';
    if (!window.App) window.App = {};
    if (!App.ChatV2) App.ChatV2 = {};

    const ChatLog = window.ChatLog || console;

    const AGG_MAP = { sum: 'sum', avg: 'mean', count: 'count', max: 'max', min: 'min' };

    /**
     * 生成高级/筛选场景的 Python 脚本
     * @param {Object} task - 聚合任务对象（含 params）
     * @returns {string} 可执行 Python 脚本
     */
    function generateAdvancedScript(task) {
        const p = task.params;
        const type = p.advancedType;

        // 构建 JSON 参数块：所有用户输入值序列化为 JSON，Python 侧 json.loads
        const paramMap = {
            dim_cols: p.dimColumns || [],
            val_cols: p.valueColumns || [],
            agg_func: p.aggFunc || 'sum',
            pandas_agg: AGG_MAP[p.aggFunc] || 'sum',
            filter_col: p.filterColumn || '',
            filter_val: p.filterValue || '',
            filter_text: p.filterText || '',
            n: 3
        };
        if (type === 'topn') {
            const m = (p.label || '').match(/\d+/);
            paramMap.n = m ? parseInt(m[0]) : 3;
        }
        const jsonParams = JSON.stringify(paramMap);

        // 通用前置
        const header = [
            'import json, pandas as pd, numpy as np',
            'import matplotlib; matplotlib.use("Agg")',
            'import matplotlib.pyplot as plt',
            '',
            '_P = ' + jsonParams + '',
            '_P = json.loads(json.dumps(_P))',
            'DIM_COLS = _P.get("dim_cols", [])',
            'VAL_COLS = _P.get("val_cols", [])',
            'AGG_FUNC = _P.get("pandas_agg", "sum")',
            'FILTER_COL = _P.get("filter_col", "")',
            'FILTER_VAL = _P.get("filter_val", "")',
            'TOP_N = _P.get("n", 3)',
            '',
            '# input_data 是从 Excel 选区读取的二维数组（含表头）',
            'try:',
            '    if "input_data" in dir() and input_data and len(input_data) > 1:',
            '        df = pd.DataFrame(input_data[1:], columns=input_data[0])',
            '    else:',
            '        df = pd.DataFrame()',
            'except Exception as e:',
            '    df = pd.DataFrame()',
            '    print(f"数据读取失败: {e}")',
            ''
        ].join('\n');

        // 数值列转换（遍历 VAL_COLS，不拼接列名）
        const numericConvert = [
            '# 数值列转换',
            'for _col in VAL_COLS:',
            '    if _col in df.columns:',
            '        df[_col] = pd.to_numeric(df[_col], errors="coerce")',
            ''
        ].join('\n');

        let body = '';

        if (type === 'filtered') {
            body = [
                '# 筛选 + 聚合',
                'if len(df) > 0:',
                '    if FILTER_COL and FILTER_COL in df.columns:',
                '        df_filtered = df[df[FILTER_COL].astype(str).str.contains(str(FILTER_VAL), na=False)]',
                '        print(f"筛选前: {len(df)} 行, 筛选后: {len(df_filtered)} 行")',
                '    else:',
                '        df_filtered = df',
                '        print(f"未找到列 {FILTER_COL}，使用全部数据")',
                '    _dim = DIM_COLS[0] if DIM_COLS else ""',
                '    _val = VAL_COLS[0] if VAL_COLS else "值"',
                '    if _dim and _dim in df_filtered.columns and _val in df_filtered.columns:',
                '        grouped = df_filtered.groupby(_dim)[_val].agg(AGG_FUNC).reset_index()',
                '        print(f"\\n按 {_dim} 聚合 {_val} 结果：")',
                '        print(grouped.to_string(index=False))',
                '        plt.figure(figsize=(10, 6))',
                '        plt.bar(grouped[_dim].astype(str), grouped[_val])',
                '        plt.title(f"按{_dim} 聚合 {_val}（筛选: {FILTER_VAL}）")',
                '        plt.xticks(rotation=45)',
                '        plt.tight_layout()',
                '        plt.show()',
                '        output = grouped.to_dict("records")',
                '    else:',
                '        print(f"未找到列: {_dim} 或 {_val}")',
                '        output = []',
                'else:',
                '    output = []'
            ].join('\n');

        } else if (type === 'topn') {
            body = [
                '# Top N 分析',
                '_dim = DIM_COLS[0] if DIM_COLS else ""',
                '_val = VAL_COLS[0] if VAL_COLS else "值"',
                'if _dim and _dim in df.columns and _val in df.columns and len(df) > 0:',
                '    grouped = df.groupby(_dim)[_val].agg(AGG_FUNC).reset_index()',
                '    top_n = grouped.nlargest(TOP_N, _val)',
                '    print(f"Top {TOP_N} 结果：")',
                '    print(top_n.to_string(index=False))',
                '    plt.figure(figsize=(10, 6))',
                '    plt.bar(top_n[_dim].astype(str), top_n[_val])',
                '    plt.title(f"Top {TOP_N} {_dim} - {_val}")',
                '    plt.xticks(rotation=45)',
                '    plt.tight_layout()',
                '    plt.show()',
                '    output = top_n.to_dict("records")',
                'elif len(df) > 0:',
                '    output = df.nlargest(TOP_N, df.columns[-1]).to_dict("records")',
                'else:',
                '    output = []'
            ].join('\n');

        } else if (type === 'compare') {
            body = [
                '# 同比/环比分析',
                'if len(df) > 0:',
                '    print("数据概览：")',
                '    print(df.head(10).to_string())',
                '    date_cols = [c for c in df.columns if any(k in str(c) for k in ["日期","时间","月","年"])]',
                '    val_cols = [c for c in df.columns if c in VAL_COLS or any(k in str(c) for k in ["销售","利润","金额","数量"])]',
                '    if date_cols and val_cols:',
                '        _dc = date_cols[0]',
                '        _vc = val_cols[0]',
                '        df_tmp = df.copy()',
                '        df_tmp[_dc] = pd.to_datetime(df_tmp[_dc], errors="coerce")',
                '        df_tmp = df_tmp.dropna(subset=[_dc, _vc])',
                '        df_tmp["月份"] = df_tmp[_dc].dt.to_period("M")',
                '        monthly = df_tmp.groupby("月份")[_vc].sum().sort_index().reset_index()',
                '        monthly["环比%"] = monthly[_vc].pct_change() * 100',
                '        print("\\n按月汇总及环比：")',
                '        print(monthly.to_string(index=False))',
                '        plt.figure(figsize=(12, 6))',
                '        plt.plot(monthly["月份"].astype(str), monthly[_vc], marker="o")',
                '        plt.title(f"{_vc} 趋势分析")',
                '        plt.xticks(rotation=45)',
                '        plt.tight_layout()',
                '        plt.show()',
                '        output = monthly.to_dict("records")',
                '    else:',
                '        output = {"error": "未找到日期列或数值列"}',
                'else:',
                '    output = []'
            ].join('\n');

        } else if (type === 'date_group') {
            body = [
                '# 日期分组汇总',
                'if len(df) > 0:',
                '    date_cols = [c for c in df.columns if any(k in str(c) for k in ["日期","时间","月","年","周"])]',
                '    val_cols = [c for c in df.columns if c in VAL_COLS or any(k in str(c) for k in ["销售","利润","金额","数量"])]',
                '    if date_cols and val_cols:',
                '        _dc = date_cols[0]',
                '        _vc = val_cols[0]',
                '        df_tmp = df.copy()',
                '        df_tmp[_dc] = pd.to_datetime(df_tmp[_dc], errors="coerce")',
                '        df_tmp = df_tmp.dropna(subset=[_dc, _vc])',
                '        df_tmp["月份"] = df_tmp[_dc].dt.to_period("M")',
                '        monthly = df_tmp.groupby("月份")[_vc].agg(AGG_FUNC).reset_index()',
                '        print("按月份汇总结果：")',
                '        print(monthly.to_string(index=False))',
                '        plt.figure(figsize=(12, 6))',
                '        plt.bar(monthly["月份"].astype(str), monthly[_vc])',
                '        plt.title(f"按月份汇总 {_vc}")',
                '        plt.xticks(rotation=45)',
                '        plt.tight_layout()',
                '        plt.show()',
                '        output = monthly.to_dict("records")',
                '    else:',
                '        output = {"error": "未找到日期列或数值列"}',
                'else:',
                '    output = []'
            ].join('\n');

        } else {
            body = [
                '# 通用聚合分析',
                'if len(df) > 0:',
                '    output = df.to_dict("records")',
                'else:',
                '    output = []'
            ].join('\n');
        }

        const footer = [
            '',
            '_cnt = len(output) if isinstance(output, list) else 1',
            'print(f"\\n分析完成，共 {_cnt} 条结果")',
            ''
        ].join('\n');

        return header + numericConvert + body + footer;
    }

    App.ChatV2.scriptGenerator = { generateAdvancedScript };

    ChatLog.info?.('CHAT-V2', '✅ scriptGenerator loaded');
})();
