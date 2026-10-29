---
id: business_card
name: 名片
description: 解析名片信息，智能提取联系人详情和公司信息
icon: 
category: 商务
outputFormat: table
tags: 名片, 联系人, 商务
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 姓名 | name | `^([\w\u4e00-\u9fa5]{1,4})` | 是 | 是 |
| 职位 | title | `(?:职位|职务|头衔)[：:\s]*([\u4e00-\u9fa5]+)` | 否 | 是 |
| 公司 | company | `(?<![\u4e00-\u9fa5])(?:公司|单位|组织|企业)[：:\s]*([^\n：:]+)` | 否 | 是 |
| 部门 | department | `(?:部门)[：:\s]*([^\n：:]+)` | 否 | 是 |
| 手机 | mobile | `(?:手机|Mobile|移动电话)[：:\s]*(1[3-9]\d{9})` | 否 | 是 |
| 电话 | phone | `(?:电话|Tel|座机)[：:\s]*(0\d{2,3}-?\d{7,8}|400[-]?\d{3,4}-?\d{4})` | 否 | 是 |
| 邮箱 | email | `[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}` | 否 | 是 |
| 网站 | website | `(?:网站|Web|网址)[：:\s]*(https?://[^\s]+|www\.[^\s]+)` | 否 | 是 |
| 地址 | address | `(?:地址|Addr|通讯地址)[：:\s]*([^\n]+)` | 否 | 是 |

## 解析脚本

```javascript
function parse(text) {
    var r = {};
    var lines = text.split('\n').map(function(l) {
        return l.trim();
    }).filter(function(l) {
        return l;
    });
    
    if (lines.length > 0) {
        r.name = lines[0];
    }
    
    if (lines.length > 1 && /(经理|总监|主管|总裁|专员|工程师|教授|博士|顾问|代表|主任|部长|处长|科长)/.test(lines[1])) {
        r.title = lines[1];
    }
    
    for (var i = 0; i < FIELDS.length; i++) {
        var f = FIELDS[i];
        if (f.pattern && f.key !== 'name' && f.key !== 'title') {
            try {
                var rx = new RegExp(f.pattern, 'i');
                var m = text.match(rx);
                if (m) {
                    r[f.key] = (m[1] || m[0]).trim();
                }
            } catch (e) {
            }
        }
    }
    
    if (!r.company) {
        for (var j = 0; j < lines.length; j++) {
            if (/公司|企业|集团|有限|股份|合伙|事务所/.test(lines[j])) {
                r.company = lines[j];
                break;
            }
        }
    }
    
    var phoneAll = /1[3-9]\d{9}/g;
    var phones = text.match(phoneAll) || [];
    if (phones.length && !r.mobile) {
        r.mobile = phones[0];
    }
    
    return r;
}
```

## 示例文本

```
李四
技术总监
研发部
ABC科技有限公司
手机: 13800138000
电话: 010-88888888
邮箱: lisi@abc.com
网站: www.abc.com
地址: 北京市海淀区中关村大街1号
微信: lisi_abc
```
