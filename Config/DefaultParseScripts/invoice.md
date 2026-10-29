---
id: invoice
name: 增值税发票
description: 解析增值税专用/普通/电子发票，提取完整票面信息
icon: 
category: 票据
outputFormat: table
tags: 发票, 税务, 财务
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 发票类型 | invoice_type | `(增值税(?:专用|普通|电子)发票)` | 否 | 是 |
| 发票代码 | invoice_code | `发票代码[：:\s]*(\d+)` | 是 | 是 |
| 发票号码 | invoice_number | `发票号码[：:\s]*(\d+)` | 是 | 是 |
| 开票日期 | issue_date | `开票日期[：:\s]*([\d\-年月日]+)` | 是 | 是 |
| 购买方名称 | buyer_name | `购[\s\S]*?名称[：:\s]*([^\n：:]+)` | 否 | 是 |
| 购买方税号 | buyer_tax_id | `购[\s\S]*?纳税人识别号[：:\s]*([A-Z0-9]+)` | 否 | 是 |
| 销售方名称 | seller_name | `售[\s\S]*?名称[：:\s]*([^\n：:]+)` | 否 | 是 |
| 销售方税号 | seller_tax_id | `售[\s\S]*?纳税人识别号[：:\s]*([A-Z0-9]+)` | 否 | 是 |
| 合计金额 | total_amount | `合[\s\S]*?计[\s\S]*?[¥￥]?([\d,]+\.?\d*)` | 否 | 是 |
| 价税合计 | total_with_tax | `价税合计[\s\S]*?[¥￥]?([\d,]+\.?\d*)` | 否 | 是 |
| 税率 | tax_rate | `税率[：:\s]*(\d+(?:\.\d+)?%)` | 否 | 是 |
| 税额 | tax_amount | `税额[：:\s]*[¥￥]?([\d,]+\.?\d*)` | 否 | 是 |
| 开票人 | issuer | `(?:开票人|收款人)[：:\s]*([^\n：:]+)` | 否 | 是 |

## 解析脚本

```javascript
function parse(text) {
    var r = {};
    for (var i = 0; i < FIELDS.length; i++) {
        var f = FIELDS[i];
        if (f.pattern) {
            try {
                var rx = new RegExp(f.pattern);
                var m = text.match(rx);
                if (m && m[1]) {
                    r[f.key] = m[1].trim();
                }
            } catch (e) {
            }
        }
    }
    return r;
}
```

## 示例文本

```
增值税普通发票
发票代码：1100193130
发票号码：290008765432
开票日期：2023年05月15日
购买方名称：ABC科技有限公司
购买方纳税人识别号：91110108MA01ABCD12
销售方名称：XYZ贸易有限公司
销售方纳税人识别号：91310115MA02EFGH34
合计金额：¥1000.00
税率：13%
税额：¥130.00
价税合计：¥1130.00
开票人：张三
```
