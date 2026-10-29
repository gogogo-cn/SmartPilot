/**
 * ocr-box-dedupe.js — 去掉同一行相邻检测框交界处被识别两次的字
 *
 * 检测框扩边后会互相压住一个字，识别时这个字会同时出现在左右两个结果里，
 * 例如「出生」和「生1654年12月20日」。只在框确实重叠、且交界文字一致时，
 * 把这个字留给更靠近它的那个框。
 */
(function (root, factory) {
    var api = factory();
    if (typeof module === 'object' && module.exports) {
        module.exports = api;
    }
    if (root) {
        root.dedupeOverlappingOcrItems = api.dedupeOverlappingOcrItems;
    }
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
    'use strict';

    function charUnits(ch) {
        var c = ch.charCodeAt(0);
        if (c <= 0x007E && (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r')) return 0.33;
        if (c <= 0x007E) return 0.55;
        return 1;
    }

    function textUnits(text) {
        var u = 0;
        for (var i = 0; i < text.length; i++) u += charUnits(text.charAt(i));
        return u || 1;
    }

    function pointXY(p) {
        if (!p) return null;
        if (Array.isArray(p) && p.length >= 2 && isFinite(+p[0]) && isFinite(+p[1])) {
            return { x: +p[0], y: +p[1] };
        }
        if (typeof p.x === 'number' && typeof p.y === 'number' && isFinite(p.x) && isFinite(p.y)) {
            return { x: p.x, y: p.y };
        }
        return null;
    }

    function boundsOf(item) {
        var box = item && item.box;
        if (!box || box.length < 4) return null;
        var pts = [];
        for (var i = 0; i < box.length; i++) {
            var xy = pointXY(box[i]);
            if (!xy) return null;
            pts.push(xy);
        }
        var left = pts[0].x, right = pts[0].x, top = pts[0].y, bottom = pts[0].y;
        for (var j = 1; j < pts.length; j++) {
            if (pts[j].x < left) left = pts[j].x;
            if (pts[j].x > right) right = pts[j].x;
            if (pts[j].y < top) top = pts[j].y;
            if (pts[j].y > bottom) bottom = pts[j].y;
        }
        if (right - left < 1 || bottom - top < 1) return null;
        return { left: left, right: right, top: top, bottom: bottom };
    }

    function coreText(text, side) {
        if (side === 'left') {
            var coreL = text.replace(/\s+$/, '');
            return { core: coreL, pad: text.length - coreL.length };
        }
        var coreR = text.replace(/^\s+/, '');
        return { core: coreR, pad: text.length - coreR.length };
    }

    function vOverlapRatio(a, b) {
        var inter = Math.min(a.bottom, b.bottom) - Math.max(a.top, b.top);
        if (inter <= 0) return 0;
        var shorter = Math.min(a.bottom - a.top, b.bottom - b.top);
        if (shorter <= 0) return 0;
        return inter / shorter;
    }

    function sameLine(a, b) {
        var ha = a.bottom - a.top;
        var hb = b.bottom - b.top;
        if (Math.min(ha, hb) <= 0) return false;
        if (Math.max(ha, hb) / Math.min(ha, hb) > 2.2) return false;
        return vOverlapRatio(a, b) >= 0.5;
    }

    function clusterLines(records) {
        var ordered = records.slice().sort(function (a, b) {
            var cy = (a.top + a.bottom) / 2 - (b.top + b.bottom) / 2;
            return cy || (a.left - b.left);
        });
        var lines = [];
        ordered.forEach(function (rec) {
            for (var i = 0; i < lines.length; i++) {
                if (sameLine(lines[i][0], rec)) {
                    lines[i].push(rec);
                    return;
                }
            }
            lines.push([rec]);
        });
        return lines;
    }

    function matchingPrefixLen(aText, bText, sharedUnits) {
        if (sharedUnits < 0.45) return 0;
        var maxN = Math.min(aText.length, bText.length);
        var best = 0;
        for (var n = 1; n <= maxN; n++) {
            var piece = aText.slice(-n);
            if (piece !== bText.slice(0, n)) continue;
            if (/\s/.test(piece)) continue;
            var units = textUnits(piece);
            if (units <= sharedUnits + 0.55) best = n;
        }
        return best;
    }

    function keepOnLeft(a, b, glyphX) {
        var aMid = (a.left + a.right) / 2;
        var bMid = (b.left + b.right) / 2;
        var da = Math.abs(glyphX - aMid);
        var db = Math.abs(glyphX - bMid);
        var slack = Math.max(2, 0.15 * Math.min(a.right - a.left, b.right - b.left));
        if (Math.abs(da - db) > slack) return da < db;
        var ca = +a.item.confidence || 0;
        var cb = +b.item.confidence || 0;
        if (Math.abs(ca - cb) > 0.02) return ca >= cb;
        return true;
    }

    function syncOriginal(item, previous, next) {
        if (item.originalText == null || item.originalText === previous) {
            item.originalText = next;
        }
    }

    function tryPair(a, b) {
        if (a.isFormula || b.isFormula) return;
        if (!sameLine(a, b)) return;
        var overlap = Math.min(a.right, b.right) - Math.max(a.left, b.left);
        var aText = a.item.text || '';
        var bText = b.item.text || '';
        var aCore = coreText(aText, 'left');
        var bCore = coreText(bText, 'right');
        if (!aCore.core || !bCore.core) return;
        var aUnit = (a.right - a.left) / textUnits(aCore.core);
        var bUnit = (b.right - b.left) / textUnits(bCore.core);
        if (aUnit <= 0 || bUnit <= 0) return;
        var shared;
        if (overlap > 1) {
            shared = Math.min(overlap / aUnit, overlap / bUnit);
        } else {
            var gap = Math.max(a.left, b.left) - Math.min(a.right, b.right);
            if (gap > 0.45 * Math.min(aUnit, bUnit)) return;
            shared = 1;
        }
        var n = matchingPrefixLen(aCore.core, bCore.core, shared);
        if (n <= 0) return;
        var pieceUnits = textUnits(aCore.core.slice(-n));
        var spanA = pieceUnits * aUnit;
        var spanB = pieceUnits * bUnit;
        var glyphX = ((a.right - spanA / 2) + (b.left + spanB / 2)) / 2;
        if (keepOnLeft(a, b, glyphX)) {
            var nextR = bText.slice(bCore.pad + n);
            b.item.text = nextR;
            syncOriginal(b.item, bText, nextR);
        } else {
            var nextL = aText.slice(0, aText.length - aCore.pad - n);
            a.item.text = nextL;
            syncOriginal(a.item, aText, nextL);
        }
    }

    function copyItem(item) {
        var copied = {};
        for (var k in item) if (Object.prototype.hasOwnProperty.call(item, k)) copied[k] = item[k];
        if (Array.isArray(item.box)) {
            copied.box = item.box.map(function (p) {
                if (Array.isArray(p)) return p.slice();
                if (p && typeof p === 'object') {
                    var o = {};
                    for (var key in p) if (Object.prototype.hasOwnProperty.call(p, key)) o[key] = p[key];
                    return o;
                }
                return p;
            });
        }
        return copied;
    }

    function dedupeOverlappingOcrItems(items) {
        if (!Array.isArray(items) || items.length < 2) return items;
        var indexed = [];
        var active = [];
        items.forEach(function (item) {
            if (!item || typeof item !== 'object') {
                indexed.push({ skip: true, item: item });
                return;
            }
            var copied = copyItem(item);
            var b = boundsOf(copied);
            var rec = { skip: !b || !!copied.is_formula, item: copied, isFormula: !!copied.is_formula };
            if (b) {
                rec.left = b.left;
                rec.right = b.right;
                rec.top = b.top;
                rec.bottom = b.bottom;
            }
            indexed.push(rec);
            if (!rec.skip) active.push(rec);
        });
        clusterLines(active).forEach(function (line) {
            line.sort(function (a, b) { return a.left - b.left; });
            for (var i = 0; i < line.length; i++) {
                var seam = Math.max(1, 0.6 * (line[i].bottom - line[i].top));
                for (var j = i + 1; j < line.length; j++) {
                    if (line[j].left >= line[i].right + seam) break;
                    tryPair(line[i], line[j]);
                }
            }
        });
        var out = [];
        indexed.forEach(function (rec) {
            var text = rec.item && rec.item.text;
            if (rec.skip || text) out.push(rec.item);
        });
        return out;
    }

    return { dedupeOverlappingOcrItems: dedupeOverlappingOcrItems };
});
