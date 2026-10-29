---
id: driver_license
name: 驾驶证
description: 解析中华人民共和国机动车驾驶证正页/副页完整信息
icon: 
category: 证件
outputFormat: table
tags: 驾驶证, 驾照, 证件, 交通
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 证号 | license_no | `(?:证号|驾驶证号|档案编号)[：:\s]*([A-Z0-9]{12,18})` | 是 |
| 姓名 | name | `姓名[：:\s]*([^\n：:]+)` | 是 | 是 |
| 性别 | gender | `性别[：:\s]*([男女])` | 否 | 是 |
| 国籍 | nationality | `国籍[：:\s]*([^\n：:]+)` | 否 | 否 |
| 出生日期 | birth_date | `(?:出生日期|出生)[：:\s]*([\d\-年月日]+)` | 否 |
| 准驾车型 | vehicle_class | `(?:准驾车型|准驾车型代号)[：:\s]*([A-Z][A-Z0-9]*)` | 是 |
| 有效期限 | valid_period | `(?:有效期限|有效期)[：:\s]*([^\n：:]+)` | 否 |
| 住址 | address | `住址[：:\s]*([\s\S]+?)(?=(?:准驾车型|有效期限|证号|$))` | 否 |

## 解析脚本

```javascript
function parse(text) {
  var r = {}
  var classMap = {
    'A1': '大型客车',
    'A2': '牵引车',
    'A3': '城市公交车',
    'B1': '中型客车',
    'B2': '大型货车',
    'C1': '小型汽车',
    'C2': '小型自动挡汽车',
    'C3': '低速载货汽车',
    'C4': '三轮汽车',
    'C5': '残疾人专用小型自动挡汽车',
    'D': '普通三轮摩托车',
    'E': '普通二轮摩托车',
    'F': '轻便摩托车',
    'M': '轮式自行机械车',
    'N': '无轨电车',
    'P': '有轨电车'
  }
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
  if (r.vehicle_class) {
    r.vehicle_class_desc = classMap[r.vehicle_class] || '未知'
  }
  if (r.license_no && !r.birth_date && r.license_no.length >= 14) {
    r.birth_date_from_id = r.license_no.substring(6, 10) + '-' + r.license_no.substring(10, 12) + '-' + r.license_no.substring(12, 14)
  }
  return r
}
```
