---
id: invoice_electronic
name: 电子发票(增强)
description: 解析增值税电子专用/普通发票，支持完整的货物明细、税率、税额提取
icon: 
category: 票据
outputFormat: table
tags: 发票, 电子发票, 税务, 财务
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 发票代码 | invoice_code | `发票代码[：:\s]*(\d+)` | 是 | 是 |
| 发票号码 | invoice_number | `发票号码[：:\s]*(\d+)` | 是 | 是 |
| 开票日期 | issue_date | `开票日期[：:\s]*([\d\-年月日]+)` | 是 | 是 |
| 购买方名称 | buyer_name | `购[\s\S]*?名称[：:\s]*([^\n：:]+)` | 是 | 是 |
| 购买方税号 | buyer_tax_id | `购[\s\S]*?纳税人识别号[：:\s]*([A-Z0-9]+)` | 否 |
| 销售方名称 | seller_name | `售[\s\S]*?名称[：:\s]*([^\n：:]+)` | 是 | 是 |
| 销售方税号 | seller_tax_id | `售[\s\S]*?纳税人识别号[：:\s]*([A-Z0-9]+)` | 否 |
| 合计金额 | total_amount | `合[\s\S]*?计[\s\S]*?[¥￥]?([\d,]+\.?\d*)` | 是 |
| 税率 | tax_rate | `税率[：:\s]*(\d+(?:\.\d+)?%)` | 否 | 是 |
| 税额 | tax_amount | `税额[：:\s]*[¥￥]?([\d,]+\.?\d*)` | 否 |
| 价税合计(小写) | total_with_tax | `价税合计[\s\S]*?小写[）)：:\s]*[¥￥]?([\d,]+\.?\d*)` | 否 | 是 |
| 价税合计(大写) | total_in_words | `价税合计[\s\S]*?大写[）)：:\s]*([\u4e00-\u9fa5零壹贰叁肆伍陆柒捌玖拾佰万亿圆角分整]+)` | 否 | 否 |
| 开票人 | issuer | `(?:开票人|收款人)[：:\s]*([^\n：:]+)` | 否 |
| 备注 | remark | `(?:备注|说明)[：:\s]*([^\n：:]+)` | 否 |

## 解析脚本

```javascript
function parse(text) {
  var r = {}
  for (var i = 0; i < FIELDS.length; i++) {
    var f = FIELDS[i]
    if (f.pattern) {
      try {
        var rx = new RegExp(f.pattern)
        var m = text.match(rx)
        if (m && m[1]) r[f.key] = m[1].trim()
      } catch (e) {}
    }
  }
  var items = []
  var dataLines = text.split('\n').map(function(l) {
    return l.trim()
  }).filter(function(l) {
    return l
  })
  var inSection = false
  for (var li = 0; li < dataLines.length; li++) {
    var line = dataLines[li]
    if (line.indexOf('货物') > -1 || line.indexOf('劳务') > -1 || line.indexOf('服务') > -1 || line.indexOf('应税') > -1) {
      inSection = true
      continue
    }
    if (inSection && (line.indexOf('合计') > -1 || line.indexOf('价税合计') > -1)) {
      break
    }
    if (inSection && line.indexOf('*') === 0) {
      var parts = line.split('*').filter(function(p) {
        return p.trim()
      })
      if (parts.length >= 7) {
        items.push({
          category: parts[0] || '',
          name: parts[1] || '',
          spec: parts[2] || '',
          unit: parts[3] || '',
          qty: parseFloat(parts[4]) || 0,
          price: parseFloat(parts[5]) || 0,
          amt: parseFloat(parts[6]) || 0
        })
      }
    }
  }
  r.items = items
  r.item_count = items.length
  if (items.length > 0) r.items_total = items.reduce(function(s, it) {
    return s + (it.amt || 0)
  }, 0)
  return r
}
```
