---
id: business_license
name: 营业执照
description: 解析企业/个体工商户营业执照，提取完整工商登记信息
icon: 
category: 证件
outputFormat: table
tags: 营业执照, 企业信息, 工商
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 统一社会信用代码 | credit_code | `(?:统一社会信用代码|信用代码)[：:\s]*([A-Z0-9]{18})` | 是 |
| 主体类型 | entity_type | `(?:类型|主体类型|企业类型)[：:\s]*([^\n：:]+)` | 否 |
| 名称 | name | `(?:名称|企业名称|字号名称)[：:\s]*([^\n：:]+)` | 是 |
| 法定代表人 | legal_person | `(?:法定代表人|负责人|经营者)[：:\s]*([^\n：:]+)` | 是 |
| 注册资本 | capital | `(?:注册资本|注册资本金|资金数额)[：:\s]*([\d.]+(?:万)?(?:元)?(?:人民币|RMB)?)` | 否 |
| 成立日期 | establish_date | `(?:成立日期|核准日期|注册日期)[：:\s]*([\d\-年月日]+)` | 否 |
| 住所/经营场所 | address | `(?:住所|经营场所|营业场所|地址)[：:\s]*([^\n：:]+)` | 否 |
| 经营范围 | business_scope | `(?:经营范围)[：:\s]*([\s\S]+?)(?=($|登记机关| 登记机关|年度报告|核准日期))` | 否 |
| 登记机关 | registry_office | `(?:登记机关| 登记机关)[：:\s]*([^\n：:]+)` | 否 |
| 营业期限 | business_period | `(?:营业期限|经营期限)[：:\s]*([^\n：:]+)` | 否 |

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
        if (m && m[1]) r[f.key] = m[1].trim().replace(/\s+/g, ' ')
      } catch (e) {}
    }
  }
  if (!r.credit_code) {
    var cm = text.match(/[0-9A-Z]/g)
    if (cm && cm.length >= 18) r.credit_code = cm.join('').substring(0, 18)
  }
  if (r.business_scope) r.business_scope = r.business_scope.replace(/\s+/g, ' ').trim()
  return r
}
```
