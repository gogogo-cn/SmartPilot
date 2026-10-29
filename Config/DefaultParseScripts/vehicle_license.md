---
id: vehicle_license
name: 行驶证
description: 解析机动车行驶证信息，提取车辆详细参数与注册信息
icon: 
category: 证件
outputFormat: table
tags: 行驶证, 车辆, 证件, 交通
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 车牌号码 | plate_no | `(?:车牌号码|号牌号码)[：:\s]*([京津沪渝冀豫云辽黑湘皖鲁新苏浙赣鄂桂甘晋蒙陕吉闽贵粤川青藏琼宁][A-Z][·.・]?[A-Z0-9]{5,6})` | 是 |
| 车辆类型 | vehicle_type | `(?:车辆类型)[：:\s]*([^\n：:]+)` | 是 | 是 |
| 所有人 | owner | `(?:所有人|车主)[：:\s]*([^\n：:]+)` | 是 |
| 品牌型号 | brand_model | `(?:品牌型号)[：:\s]*([^\n：:]+)` | 否 | 是 |
| VIN码 | vin | `(?:车辆识别代号|VIN码|VIN)[：:\s]*([A-HJ-NPR-Z0-9]{17})` | 否 |
| 发动机号 | engine_no | `(?:发动机号码|发动机号)[：:\s]*([A-Za-z0-9]{6,20})` | 否 |
| 注册日期 | register_date | `(?:注册日期|初次登记日期)[：:\s]*([\d\-年月日]+)` | 否 |
| 发证日期 | issue_date | `(?:发证日期|发证日期|核发日期)[：:\s]*([\d\-年月日]+)` | 否 |
| 使用性质 | use_type | `(?:使用性质)[：:\s]*([^\n：:]+)` | 否 | 是 |
| 核定载人 | seats | `(?:核定载人数|核载人数|座位数)[：:\s]*(\d+)人?` | 否 |
| 整备质量 | curb_weight | `(?:整备质量|总质量)[：:\s]*([\d.]+)kg` | 否 |
| 住址 | address | `(?:住址|住所)[：:\s]*([^\n：:]+)` | 否 |

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
  if (r.plate_no) r.plate_clean = r.plate_no.replace(/[·.]/g, '').toUpperCase()
  return r
}
```
