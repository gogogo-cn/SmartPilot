---
id: idcard
name: 身份证
description: 解析中国居民身份证信息，支持完整地址提取与智能拆分（含派出所）
icon: 🆔
category: 证件
outputFormat: table
tags: 身份证, 证件, 个人信息
---

## 字段定义

| 字段名 | 键名 | 正则表达式 | 必填 | 可见 |
|--------|------|-----------|------|------|
| 姓名 | name | `姓名[：:\s]*([^\n]+)` | 是 | 是 |
| 性别 | gender | `性别[：:\s]*([男女])` | 是 | 是 |
| 民族 | ethnicity | `民族[：:\s]*([^\n]+)` | 否 | 是 |
| 出生日期 | birth_date | `出生[：:\s]*([\d\-年月日]+)` | 是 | 是 |
| 住址 | address | `住址[：:\s]*([\s\S]+?)(?=(?:公民.{0,2}份证号码|身份号码|身份证号码|\d{18}|$))` | 是 | 是 |
| 省份 | province | - | 否 | 是 |
| 城市 | city | - | 否 | 是 |
| 区县 | district | - | 否 | 是 |
| 街道/门牌 | street | - | 否 | 是 |
| 派出所 | police_station | - | 否 | 是 |
| 公民身份号码 | id_number | `(?:公民.{0,2}份证号码|身份号码|身份证号)[：:\s]*(\d{17}[\dXx])` | 是 | 是 |
| 签发机关 | issuing_authority | `(?:签发机关)[：:\s]*([^\n]+)` | 否 | 否 |
| 有效期限 | valid_period | `(?:有效期限)[：:\s]*([^\n]+)` | 否 | 否 |

## 解析脚本

```javascript
function parse(rawText) {
    // 预处理：去除 HTML 标签、转义字符，统一空白
    var text = rawText
        .replace(/<br\s*\/?>/gi, '\n')
        .replace(/<\/p>|<p[^>]*>/gi, '\n')
        .replace(/<[^>]+>/g, '')
        .replace(/&nbsp;/g, ' ')
        .replace(/&[a-z]+;/gi, '')
        .replace(/[ \t]+\n/g, '\n')
        .replace(/\n[ \t]+/g, '\n')
        .replace(/\n+/g, '\n')
        .trim();

    var r = {};
    for (var i = 0; i < FIELDS.length; i++) {
        var f = FIELDS[i];
        if (f.pattern && f.pattern !== '-') {
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

    // 兜底：直接搜索 18 位身份证号
    if (!r.id_number) {
        var idMatch = text.match(/(\d{17}[\dXx])/);
        if (idMatch) r.id_number = idMatch[1];
    }

    if (r.id_number) {
        var id = r.id_number;
        if (!r.birth_date && id.length >= 14) {
            r.birth_date_from_id = id.substring(6, 10) + '-' + id.substring(10, 12) + '-' + id.substring(12, 14);
        }
        if (!r.gender && id.length >= 17) {
            r.gender_from_id = parseInt(id.substring(16, 17)) % 2 === 1 ? '男' : '女';
        }
    }

    if (r.address) {
        var addr = r.address.replace(/\s+/g, '');
        var psMatch = addr.match(/([^0-9]*派出所)/);
        if (psMatch) {
            r.police_station = psMatch[1];
        }

        var addrClean = addr.replace(/[^\u4e00-\u9fa5\d\-a-zA-Z]/g, '');
        var provList = ['北京市', '天津市', '上海市', '重庆市', '河北省', '山西省', '辽宁省', '吉林省', '黑龙江省', '江苏省', '浙江省', '安徽省', '福建省', '江西省', '山东省', '河南省', '湖北省', '湖南省', '广东省', '海南省', '四川省', '贵州省', '云南省', '陕西省', '甘肃省', '青海省', '台湾省', '内蒙古自治区', '广西壮族自治区', '西藏自治区', '宁夏回族自治区', '新疆维吾尔自治区', '香港特别行政区', '澳门特别行政区'];

        var provMatch = null;
        for (var p = 0; p < provList.length; p++) {
            if (addrClean.indexOf(provList[p]) === 0) {
                provMatch = provList[p];
                break;
            }
        }

        if (provMatch) {
            r.province = provMatch;
            var rest = addrClean.substring(provMatch.length);
            var mr = parseMunicipality(rest, provMatch);
            r.city = mr.city;
            r.district = mr.district;
            r.street = mr.street;
        } else {
            var ap = addrClean.match(/^(.*?[\u5e02\u5dde\u76df])(.*?[\u53bf\u533a\u65d7])(.*)/);
            if (ap) {
                r.city = ap[1];
                r.district = ap[2];
                r.street = ap[3];
            } else {
                r.street = addrClean;
            }
        }
    }

    return r;
}

function parseMunicipality(rest, province) {
    var city = '', district = '', street = rest;
    var dm = ['北京市', '天津市', '上海市', '重庆市'];

    if (dm.indexOf(province) > -1) {
        // 直辖市没有独立的“市”，直辖市本身同时是省和市
        city = province;
        var dMatch = rest.match(/^(.*?[\u53bf\u533a])/);
        if (dMatch) {
            district = dMatch[1];
            street = rest.substring(dMatch[0].length);
        } else {
            street = rest;
        }
    } else {
        var cMatch = rest.match(/^(.*?[\u5e02\u5dde\u76df])/);
        if (cMatch) {
            city = cMatch[1];
            var ac = rest.substring(cMatch[0].length);
            var dm2 = ac.match(/^(.*?[\u53bf\u533a\u65d7])/);
            if (dm2) {
                district = dm2[1];
                street = ac.substring(dm2[0].length);
            } else {
                street = ac;
            }
        } else {
            street = rest;
        }
    }

    return {
        city: city,
        district: district,
        street: street
    };
}
```

## 示例文本

```
姓名：张三
性别：男
民族：汉
出生：1990年1月1日
住址：北京市朝阳区建国路88号SOHO现代城A座1201室
公民身份号码：110105199001011234
```
