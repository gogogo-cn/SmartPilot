/**
 * 把行对象写成真正的 xlsx（Office Open XML）。不用 HTML 冒充 xls。
 */
(function (root, factory) {
    var api = factory();
    if (typeof module === 'object' && module.exports) module.exports = api;
    if (root) root.XlsxTable = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';

    var CRC_TABLE = (function () {
        var table = new Uint32Array(256);
        for (var n = 0; n < 256; n++) {
            var c = n;
            for (var k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
            table[n] = c >>> 0;
        }
        return table;
    })();

    function crc32(bytes) {
        var c = 0xFFFFFFFF;
        for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
        return (c ^ 0xFFFFFFFF) >>> 0;
    }

    function u16(n) {
        return [n & 255, (n >>> 8) & 255];
    }

    function u32(n) {
        return [n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255];
    }

    function concat(parts) {
        var size = 0;
        for (var i = 0; i < parts.length; i++) size += parts[i].length;
        var out = new Uint8Array(size);
        var offset = 0;
        for (var j = 0; j < parts.length; j++) {
            out.set(parts[j], offset);
            offset += parts[j].length;
        }
        return out;
    }

    function bytesOf(list) {
        return new Uint8Array(list);
    }

    function utf8(text) {
        if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(text);
        var encoded = unescape(encodeURIComponent(text));
        var out = new Uint8Array(encoded.length);
        for (var i = 0; i < encoded.length; i++) out[i] = encoded.charCodeAt(i);
        return out;
    }

    function storeZip(files) {
        var locals = [];
        var centrals = [];
        var offset = 0;
        files.forEach(function (file) {
            var name = utf8(file.name);
            var data = file.data;
            var crc = crc32(data);
            var local = concat([
                bytesOf([0x50, 0x4B, 0x03, 0x04]),
                bytesOf(u16(20)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u32(crc)),
                bytesOf(u32(data.length)),
                bytesOf(u32(data.length)),
                bytesOf(u16(name.length)),
                bytesOf(u16(0)),
                name,
                data
            ]);
            var central = concat([
                bytesOf([0x50, 0x4B, 0x01, 0x02]),
                bytesOf(u16(20)),
                bytesOf(u16(20)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u32(crc)),
                bytesOf(u32(data.length)),
                bytesOf(u32(data.length)),
                bytesOf(u16(name.length)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u16(0)),
                bytesOf(u32(0)),
                bytesOf(u32(offset)),
                name
            ]);
            locals.push(local);
            centrals.push(central);
            offset += local.length;
        });
        var centralDir = concat(centrals);
        var end = concat([
            bytesOf([0x50, 0x4B, 0x05, 0x06]),
            bytesOf(u16(0)),
            bytesOf(u16(0)),
            bytesOf(u16(files.length)),
            bytesOf(u16(files.length)),
            bytesOf(u32(centralDir.length)),
            bytesOf(u32(offset)),
            bytesOf(u16(0))
        ]);
        return concat(locals.concat([centralDir, end]));
    }

    function xml(text) {
        return String(text == null ? '' : text)
            .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;');
    }

    function columnName(index) {
        var name = '';
        var n = index + 1;
        while (n > 0) {
            var rem = (n - 1) % 26;
            name = String.fromCharCode(65 + rem) + name;
            n = Math.floor((n - 1) / 26);
        }
        return name;
    }

    function headersOf(rows) {
        var headers = [];
        rows.forEach(function (row) {
            if (!row) return;
            Object.keys(row).forEach(function (key) {
                if (headers.indexOf(key) < 0) headers.push(key);
            });
        });
        return headers;
    }

    function sheetXml(rows) {
        var headers = headersOf(rows);
        var lines = [
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        ];
        function rowXml(rowIndex, values) {
            var cells = values.map(function (value, col) {
                var text = xml(value);
                var space = /^\s|\s$/.test(String(value == null ? '' : value)) ? ' xml:space="preserve"' : '';
                return '<c r="' + columnName(col) + rowIndex + '" t="inlineStr"><is><t' + space + '>' + text + '</t></is></c>';
            }).join('');
            return '<row r="' + rowIndex + '">' + cells + '</row>';
        }
        lines.push(rowXml(1, headers));
        rows.forEach(function (row, index) {
            lines.push(rowXml(index + 2, headers.map(function (header) {
                return row && row[header] != null ? row[header] : '';
            })));
        });
        lines.push('</sheetData></worksheet>');
        return lines.join('');
    }

    function build(rows) {
        var list = Array.isArray(rows) ? rows : [];
        var files = [
            {
                name: '[Content_Types].xml',
                data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
                    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
                    '<Default Extension="xml" ContentType="application/xml"/>' +
                    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
                    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
                    '</Types>')
            },
            {
                name: '_rels/.rels',
                data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
                    '</Relationships>')
            },
            {
                name: 'xl/workbook.xml',
                data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
                    '<sheets><sheet name="识别结果" sheetId="1" r:id="rId1"/></sheets></workbook>')
            },
            {
                name: 'xl/_rels/workbook.xml.rels',
                data: utf8('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
                    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
                    '</Relationships>')
            },
            {
                name: 'xl/worksheets/sheet1.xml',
                data: utf8(sheetXml(list))
            }
        ];
        return storeZip(files);
    }

    return { build: build };
});
