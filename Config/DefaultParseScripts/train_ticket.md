---
id: train_ticket
name: 火车票
description: 解析中国铁路火车票/高铁票信息，提取车次、站点、座位等全部信息
icon: 
category: 票据
outputFormat: table
tags: 火车票, 高铁, 出行
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 姓名 | passenger_name | `(?:(?:姓名|乘客)[：:\s]*([^\n：:]{2,10})|\d{17}[\dXx]([^\n：:]{2,10}))` | 是 |
| 出发站 | from_station | `([\u4e00-\u9fa5]{2,10}站)(?=.*[GDCKTZEHLYSPWJ]\d{1,4})` | 是 |
| 到达站 | to_station | `[GDCKTZEHLYSPWJ]\d{1,4}(?:次)?[^\n]*[\n\s]*([\u4e00-\u9fa5]{2,10}站)` | 是 |
| 车次 | train_no | `([GDCKTZEHLYSPWJ]\d{1,4})\s*次?` | 是 |
| 出发日期 | date | `(\d{4}[\-/年]\d{1,2}[\-/月]\d{1,2}日?)` | 是 |
| 出发时间 | time | `(\d{1,2}:\d{2})\s*开` | 是 |
| 车厢 | carriage | `(?<![:\d])(\d{1,2})\s*车` | 否 |
| 座位号 | seat_no | `(\d{1,3})\s*号` | 否 |
| 票价 | price | `[¥￥]\s*(\d+\.?\d*)\s*元` | 否 |
| 席别 | seat_class | `(新空调硬卧|新空调软卧|硬卧|软卧|硬座|软座|二等座|一等座|商务座|特等座|无座)` | 否 |
| 候车地点 | waiting_hall | `(?:候车|检票口|进站口)[：:\s]*([^\n：:]+)` | 否 |

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
        if (m) {
          var val = m[1] || m[2] || ''
          if (val) r[f.key] = val.trim()
        }
      } catch (e) {}
    }
  }
  if (r.train_no) {
    r.train_info = r.train_no + '次'
    if (r.date) r.train_info += ' ' + r.date
    if (r.time) r.train_info += ' ' + r.time + '开'
  }
  if (r.carriage || r.seat_no) {
    r.seat_info = (r.carriage || '') + '车' + (r.seat_no || '')
    if (r.seat_class) r.seat_info += ' ' + r.seat_class
  }
  if (r.from_station && r.to_station) {
    r.route = r.from_station + ' → ' + r.to_station
  }
  return r
}
```

## 示例文本

```
[1] text="R093443"  confidence=0.9974  box=[[1019,566],[1697,577],[1695,729],[1016,718]]
[2] text="银川站"  confidence=0.9979  box=[[1151,718],[1925,718],[1925,904],[1151,904]]
[3] text="北京站"  confidence=0.9994  box=[[2787,719],[3584,708],[3587,898],[2790,910]]
[4] text="K1178"  confidence=0.9940  box=[[2108,753],[2634,753],[2634,909],[2108,909]]
[5] text="Yinchuan"  confidence=0.9892  box=[[1277,898],[1767,910],[1764,1032],[1274,1020]]
[6] text="Beijing"  confidence=0.9888  box=[[2954,908],[3354,920],[3351,1047],[2950,1035]]
[7] text="2019年10月06日16:05开"  confidence=0.9749  box=[[1042,1046],[2475,1051],[2475,1183],[1042,1178]]
[8] text="14车015号上铺"  confidence=0.9958  box=[[2749,1045],[3651,1056],[3650,1193],[2748,1182]]
[9] text="网"  confidence=0.9995  box=[[2138,1207],[2272,1207],[2272,1339],[2138,1339]]
[10] text="￥280.5元"  confidence=0.9677  box=[[1057,1217],[1607,1217],[1607,1339],[1057,1339]]
[11] text="新空调硬卧"  confidence=0.9929  box=[[2798,1217],[3418,1217],[3418,1344],[2798,1344]]
[12] text="限乘当日当次车"  confidence=0.9928  box=[[1042,1363],[1890,1363],[1890,1495],[1042,1495]]
[13] text="3424231998****1540裴丽丽"  confidence=0.9866  box=[[1032,1666],[2728,1671],[2728,1818],[1032,1813]]
[14] text="买票请到12306发货请到95306"  confidence=0.9965  box=[[1335,1857],[2768,1867],[2767,1979],[1334,1969]]
[15] text="中国铁路祝您旅途愉快"  confidence=0.9957  box=[[1524,1988],[2575,1999],[2574,2116],[1522,2106]]
[16] text="14846301011007R093443 银川售"  confidence=0.9800  box=[[1038,2154],[2635,2175],[2633,2307],[1036,2286]]
```
