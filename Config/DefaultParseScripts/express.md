---
id: express
name: 快递单
description: 解析快递/物流面单信息，完整提取收发件人和运单详情
icon: 
category: 票据
outputFormat: table
tags: 快递, 物流, 运单
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 运单号 | tracking_no | `(?:运单号|快递单号|单号|追踪号)[：:\s]*([A-Z0-9]{10,20})` | 是 | 是 |
| 快递公司 | express_company | `(?:快递公司|承运商|物流)[：:\s]*([^\n：:]{2,20})` | 否 | 是 |
| 寄件人姓名 | sender_name | `(?:寄件人|发件人|寄件)[^\n]*姓名[：:\s]*([^\n：:]+)` | 否 | 是 |
| 寄件人电话 | sender_phone | `(?:寄件人|发件人|寄件)[^\n]*(?:电话|手机|联系方式)[：:\s]*(1[3-9]\d{9})` | 否 | 是 |
| 寄件地址 | sender_address | `(?:寄件人|发件人|寄件)[^\n]*(?:地址|详细地址)[：:\s]*([^\n]+)` | 否 | 否 |
| 收件人姓名 | receiver_name | `(?:收件人|收货人|收件)[^\n]*姓名[：:\s]*([^\n：:]+)` | 是 | 是 |
| 收件人电话 | receiver_phone | `(?:收件人|收货人|收件)[^\n]*(?:电话|手机|联系方式)[：:\s]*(1[3-9]\d{9})` | 是 | 是 |
| 收件地址 | receiver_address | `(?:收件人|收货人|收件)[^\n]*(?:地址|详细地址)[：:\s]*([^\n]+)` | 是 | 是 |
| 物品名称 | goods_name | `(?:内件品名|物品名称|货物名称)[：:\s]*([^\n]+)` | 否 | 是 |
| 重量 | weight | `(?:重量|计费重量)[：:\s]*([\d.]+)(?:kg|公斤|千克)?` | 否 | 否 |
| 运费 | shipping_fee | `(?:运费|费用)[：:\s]*[¥￥]?([\d.]+)元?` | 否 | 否 |

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
    
    if (!r.sender_phone || !r.receiver_phone) {
        var pr = /1[3-9]\d{9}/g;
        var phs = text.match(pr) || [];
        if (phs.length >= 1 && !r.sender_phone) {
            r.sender_phone = phs[0];
        }
        if (phs.length >= 2 && !r.receiver_phone) {
            r.receiver_phone = phs[1];
        }
    }
    
    return r;
}
```

## 示例文本

```
运单号：SF1234567890
快递公司：顺丰速运
寄件人姓名：张三
寄件人电话：13800138000
收件人姓名：李四
收件人电话：13900139000
收件地址：北京市朝阳区建国路88号SOHO现代城A座1201室
内件品名：电子产品
重量：2.5kg
```
