/* =============================================================================
 * wps_macro.js —— WPS 演示「模板配色自适应」宏（在 WPS 内部运行，无需外部工具）
 * -----------------------------------------------------------------------------
 * 作用：选中刚插入的模板内容（图标 / 关系图 / 流程图 / 图表 / 表格 / SmartArt…），
 *       运行本宏，它会自动读取「当前演示的整体配色 + 当前页的实际配色」，
 *       把选中对象的颜色整体改成这套配色，并保留原有的明暗层次与渐变方向。
 *
 * 安装：WPS 演示 → 开发工具 → 宏编辑器（WPS 宏编辑器）→ 插入 → 模块 →
 *       把本文件整段粘进去 → 保存。
 * 运行：宏列表里选 recolorSelection 运行；或把 recolorSelection 加到快速访问工具栏。
 *
 * 入口函数：
 *   recolorSelection()      改当前选中的对象（没有选中就改本页最后插入的那个）
 *   recolorLastInserted()   改本页最后插入的对象
 *   recolorWholeSlide()     改整页所有形状
 *   showSlidePalette()      只显示取到的配色，不改任何东西（用来先确认取色对不对）
 *
 * 说明：本宏大量使用 try/catch 做能力降级——WPS 各版本 JSAPI 支持度不同，
 *       某个属性取不到就跳过该处，不会整段报错中断。
 * ========================================================================== */

/* ------------------------------ 常量 ------------------------------ */
var MSO_GROUP = 6, MSO_CHART = 3, MSO_TABLE = 19, MSO_SMARTART = 24, MSO_DIAGRAM = 21;
var MSO_FILL_SOLID = 1, MSO_FILL_GRADIENT = 3, MSO_FILL_BACKGROUND = 5;
var MSO_TRUE = -1, MSO_FALSE = 0;

/* --------------------------- 颜色基础工具 --------------------------- */
// VBA/WPS 的 RGB 属性是 R + G*256 + B*65536 的整数
function vbaR(v) { return (v >>> 0) & 0xFF; }
function vbaG(v) { return (v >>> 8) & 0xFF; }
function vbaB(v) { return (v >>> 16) & 0xFF; }
function toVba(r, g, b) { return (r & 255) + ((g & 255) * 256) + ((b & 255) * 65536); }

function hex2rgb(h) {
    h = String(h).replace('#', '');
    if (h.length === 8) h = h.slice(2);
    return [parseInt(h.substr(0, 2), 16), parseInt(h.substr(2, 2), 16), parseInt(h.substr(4, 2), 16)];
}
function rgb2hex(r, g, b) {
    function c(x) { x = Math.max(0, Math.min(255, Math.round(x))); var s = x.toString(16).toUpperCase(); return s.length < 2 ? '0' + s : s; }
    return c(r) + c(g) + c(b);
}
function rgb2hsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    var mx = Math.max(r, g, b), mn = Math.min(r, g, b), h = 0, s = 0, l = (mx + mn) / 2, d = mx - mn;
    if (d !== 0) {
        s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
        if (mx === r) h = ((g - b) / d + (g < b ? 6 : 0));
        else if (mx === g) h = (b - r) / d + 2;
        else h = (r - g) / d + 4;
        h *= 60;
    }
    return [h, s, l];
}
function hsl2rgb(h, s, l) {
    h = ((h % 360) + 360) % 360 / 360;
    s = Math.max(0, Math.min(1, s)); l = Math.max(0, Math.min(1, l));
    function hue2rgb(p, q, t) {
        if (t < 0) t += 1; if (t > 1) t -= 1;
        if (t < 1 / 6) return p + (q - p) * 6 * t;
        if (t < 1 / 2) return q;
        if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
        return p;
    }
    if (s === 0) { var v = l * 255; return [v, v, v]; }
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    return [hue2rgb(p, q, h + 1 / 3) * 255, hue2rgb(p, q, h) * 255, hue2rgb(p, q, h - 1 / 3) * 255];
}
function hexHsl(h) { var c = hex2rgb(h); return rgb2hsl(c[0], c[1], c[2]); }
function hslHex(h, s, l) { var c = hsl2rgb(h, s, l); return rgb2hex(c[0], c[1], c[2]); }
function hueDist(a, b) { var d = Math.abs(a - b) % 360; return Math.min(d, 360 - d); }
function relLum(hex) { var c = hex2rgb(hex); return (0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]) / 255; }

// 中性色判定必须用 RGB 绝对彩度：HSL 的 S 在接近白/黑时会被放大（#EAF1F9 的 S 高达 0.56）
function isNeutral(hex) {
    var c = hex2rgb(hex), mx = Math.max(c[0], c[1], c[2]), mn = Math.min(c[0], c[1], c[2]);
    var chroma = (mx - mn) / 255, l = (mx + mn) / 510;
    return chroma < 0.12 || l > 0.945 || l < 0.075;
}

/* --------------------------- WPS 对象访问 --------------------------- */
function activePresentation() { return Application.ActivePresentation; }

function currentSlide() {
    var s = null;
    try { s = Application.ActiveWindow.View.Slide; } catch (e) { }
    if (!s) {
        try { s = Application.ActiveWindow.Selection.SlideRange.Item(1); } catch (e) { }
    }
    if (!s) {
        try {
            var idx = Application.ActiveWindow.Selection.SlideRange.Item(1).SlideIndex;
            s = activePresentation().Slides.Item(idx);
        } catch (e) { }
    }
    if (!s) { try { s = activePresentation().Slides.Item(1); } catch (e) { } }
    return s;
}

function selectedShapes(slide) {
    var out = [];
    try {
        var sel = Application.ActiveWindow.Selection;
        if (sel && sel.Type === 2 && sel.ShapeRange && sel.ShapeRange.Count > 0) {
            for (var i = 1; i <= sel.ShapeRange.Count; i++) out.push(sel.ShapeRange.Item(i));
        }
    } catch (e) { }
    if (out.length === 0 && slide) {
        // 没有选中 → 本页最后插入的非占位形状（模板粘贴后通常就在最后）
        try {
            var n = slide.Shapes.Count;
            for (var j = n; j >= 1; j--) {
                var sh = slide.Shapes.Item(j);
                var isPh = false;
                try { isPh = sh.PlaceholderFormat !== null && sh.PlaceholderFormat !== undefined; } catch (e) { }
                var hasPh = false;
                try { hasPh = (sh.Type === 14); } catch (e) { }
                if (!isPh && !hasPh) { out.push(sh); break; }
            }
            if (out.length === 0 && n > 0) out.push(slide.Shapes.Item(n));
        } catch (e) { }
    }
    return out;
}

function shapeArea(sh) {
    try { return sh.Width * sh.Height; } catch (e) { }
    return 1;
}

// 递归展开组合
function eachShape(sh, fn, depth) {
    depth = depth || 0;
    fn(sh, depth);
    var t = 0;
    try { t = sh.Type; } catch (e) { }
    if (t === MSO_GROUP) {
        try {
            for (var i = 1; i <= sh.GroupItems.Count; i++) {
                try { eachShape(sh.GroupItems.Item(i), fn, depth + 1); } catch (e) { }
            }
        } catch (e) { }
    }
}

/* ------------------------------ 取色 ------------------------------ */
// 从形状上摘一个颜色（返回 {hex, kind, area} 或 null）
function grab(sh, getColor, kind) {
    try {
        var v = getColor();
        if (v === null || v === undefined) return null;
        var r = vbaR(v), g = vbaG(v), b = vbaB(v);
        return { hex: rgb2hex(r, g, b), kind: kind, area: shapeArea(sh) };
    } catch (e) { return null; }
}

var KIND_W = { fill: 1.0, stop: 0.75, line: 0.35, text: 0.25 };

function harvestShape(sh, out, targetSet) {
    eachShape(sh, function (s) {
        // 填充
        var f = grab(s, function () {
            if (!s.Fill || s.Fill.Type === MSO_FILL_BACKGROUND) return null;
            return s.Fill.ForeColor.RGB;
        }, 'fill');
        if (f) out.push(f);
        // 渐变：把同一渐变的所有 stop 绑定成一组，保证只做深浅变化
        var gradId = null;
        try {
            if (s.Fill && s.Fill.Type === MSO_FILL_GRADIENT && s.Fill.GradientColorStops) {
                gradId = 'g' + out.length + '_' + Math.random();
                var stops = s.Fill.GradientColorStops;
                for (var i = 1; i <= stops.Count; i++) {
                    try {
                        out.push({
                            hex: rgb2hex(vbaR(stops.Item(i).Color.RGB), vbaG(stops.Item(i).Color.RGB), vbaB(stops.Item(i).Color.RGB)),
                            kind: 'stop', area: shapeArea(s), grad: gradId
                        });
                    } catch (e) { }
                }
            }
        } catch (e) { }
        // 线条
        var l = grab(s, function () {
            if (!s.Line || !s.Line.Visible) return null;
            return s.Line.ForeColor.RGB;
        }, 'line');
        if (l) out.push(l);
        // 文字
        var t = grab(s, function () {
            if (!s.TextFrame || !s.TextFrame.HasText) return null;
            return s.TextFrame.TextRange.Font.Color.RGB;
        }, 'text');
        if (t) out.push(t);
    });
}

// 聚类：按色相分族（与本地引擎同一套规则）
function cluster(items) {
    var chroma = [], neutral = [];
    for (var i = 0; i < items.length; i++) {
        if (isNeutral(items[i].hex)) neutral.push(items[i]); else chroma.push(items[i]);
    }
    var bins = {};
    for (var i = 0; i < chroma.length; i++) {
        var h = hexHsl(chroma[i].hex)[0], k = Math.floor(h / 12);
        (bins[k] = bins[k] || []).push(chroma[i]);
    }
    var keys = Object.keys(bins).map(Number).sort(function (a, b) { return a - b; });
    var groups = [];
    for (var i = 0; i < keys.length; i++) {
        var mem = bins[keys[i]], mh = 0;
        for (var j = 0; j < mem.length; j++) mh += hexHsl(mem[j].hex)[0];
        mh /= mem.length;
        if (groups.length && hueDist(groups[groups.length - 1].h, mh) < 26) {
            groups[groups.length - 1].mem = groups[groups.length - 1].mem.concat(mem);
        } else groups.push({ h: mh, mem: mem });
    }
    var clusters = [];
    for (var i = 0; i < groups.length; i++) {
        var mem = groups[i].mem, tot = 0, sx = 0, sy = 0, sl = 0, ss = 0;
        for (var j = 0; j < mem.length; j++) {
            var w = KIND_W[mem[j].kind] || 0.2;
            w *= Math.pow(Math.max(0.05, mem[j].area / 5e6), 0.45);
            tot += w;
            var hs = hexHsl(mem[j].hex);
            sx += Math.cos(hs[0] * Math.PI / 180) * w;
            sy += Math.sin(hs[0] * Math.PI / 180) * w;
            sl += hs[2] * w; ss += hs[1] * w;
        }
        tot = tot || 1;
        var mh = Math.atan2(sy, sx) * 180 / Math.PI; mh = ((mh % 360) + 360) % 360;
        clusters.push({ hue: mh, l: sl / tot, s: ss / tot, weight: tot, hex: hslHex(mh, ss / tot, sl / tot), mem: mem });
    }
    clusters.sort(function (a, b) { return b.weight - a.weight; });
    return { clusters: clusters, neutral: neutral };
}

// 同一渐变的所有 stop 必须落在同一色族
function mergeGradient(clusters) {
    var map = {}, changed = false;
    for (var i = 0; i < clusters.length; i++) {
        for (var j = 0; j < clusters[i].mem.length; j++) {
            var g = clusters[i].mem[j].grad;
            if (g) { (map[g] = map[g] || []).push(i); }
        }
    }
    for (var g in map) {
        var idxs = map[g].filter(function (v, i, a) { return a.indexOf(v) === i; });
        if (idxs.length < 2) continue;
        changed = true;
        idxs.sort(function (a, b) { return clusters[b].weight - clusters[a].weight; });
        var main = clusters[idxs[0]];
        for (var k = 1; k < idxs.length; k++) {
            main.mem = main.mem.concat(clusters[idxs[k]].mem);
            main.weight += clusters[idxs[k]].weight;
            clusters[idxs[k]]._drop = true;
        }
    }
    if (!changed) return clusters;
    var out = [];
    for (var i = 0; i < clusters.length; i++) if (!clusters[i]._drop) out.push(clusters[i]);
    out.sort(function (a, b) { return b.weight - a.weight; });
    return out;
}

// 读取主题色（不同版本 API 不同，逐级降级）
function themeAccents(slide) {
    var out = [];
    var names = [['accent1', 5], ['accent2', 6], ['accent3', 7], ['accent4', 8], ['accent5', 9], ['accent6', 10]];
    try {
        var tcs = null;
        try { tcs = slide.ThemeColorScheme; } catch (e) { }
        if (!tcs) { try { tcs = slide.CustomLayout.ThemeColorScheme; } catch (e) { } }
        if (tcs) {
            for (var i = 0; i < names.length; i++) {
                try {
                    var v = tcs.Colors(names[i][1]).RGB;
                    out.push(rgb2hex(vbaR(v), vbaG(v), vbaB(v)));
                } catch (e) { }
            }
        }
    } catch (e) { }
    if (out.length === 0) {
        try {
            var cs = slide.ColorScheme;
            if (cs) {
                for (var i = 0; i < names.length; i++) {
                    try { var v = cs.Colors(6 + i).RGB; out.push(rgb2hex(vbaR(v), vbaG(v), vbaB(v))); } catch (e) { }
                }
            }
        } catch (e) { }
    }
    if (out.length === 0) {
        // 兜底：扫母版 / 版式上形状的实际颜色
        try {
            var lay = slide.CustomLayout;
            if (lay) {
                var items = [];
                for (var i = 1; i <= lay.Shapes.Count; i++) harvestShape(lay.Shapes.Item(i), items);
                var c = cluster(items);
                for (var j = 0; j < c.clusters.length && out.length < 6; j++) out.push(c.clusters[j].hex);
            }
        } catch (e) { }
    }
    return out;
}

/* ------------------------------ 建调色板 ------------------------------ */
function buildPalette(slide, targets, useThemeFirst) {
    var items = [], i, j;
    var isTarget = function (sh) {
        for (var k = 0; k < targets.length; k++) if (targets[k] === sh) return true;
        return false;
    };
    try {
        for (i = 1; i <= slide.Shapes.Count; i++) {
            var sh = slide.Shapes.Item(i);
            if (isTarget(sh)) continue;
            harvestShape(sh, items);
        }
    } catch (e) { }

    var res = cluster(items);
    var clusters = mergeGradient(res.clusters);
    var neutral = res.neutral;
    var accents = themeAccents(slide);

    // 当页颜色不够 → 用主题色补位。只补「与当页已有色族还算协调」的（色相距离升序，
    // 且不超过 HARMONY_MAX）；太跳的宁可不补，让源色族走主色的深浅派生。
    // 往一份只用蓝青的稿子里硬塞主题里的橙色是帮倒忙。
    var HARMONY_MAX = 90, cands = [];
    for (j = 0; j < accents.length; j++) {
        var a = accents[j];
        if (isNeutral(a)) continue;
        var ah = hexHsl(a)[0], dmin = 1e9;
        for (i = 0; i < clusters.length; i++) dmin = Math.min(dmin, hueDist(ah, clusters[i].hue));
        if (clusters.length === 0) dmin = 0;
        if (dmin <= 25) continue;
        var hs = hexHsl(a);
        cands.push({ d: dmin, hue: hs[0], l: hs[2], s: hs[1], hex: a });
    }
    cands.sort(function (x, y) { return x.d - y.d; });
    for (j = 0; j < cands.length && j < 3; j++) {
        if (cands[j].d > HARMONY_MAX) break;
        clusters.push({ hue: cands[j].hue, l: cands[j].l, s: cands[j].s,
                        weight: 0.001, hex: cands[j].hex, mem: [] });
    }
    if (clusters.length === 0) {
        for (j = 0; j < accents.length; j++) {
            if (!isNeutral(accents[j])) {
                var hs2 = hexHsl(accents[j]);
                clusters.push({ hue: hs2[0], l: hs2[2], s: hs2[1], weight: 1, hex: accents[j], mem: [] });
                break;
            }
        }
    }
    if (clusters.length === 0) clusters.push({ hue: 210, l: 0.45, s: 0.6, weight: 1, hex: '2F6FB5', mem: [] });

    if (useThemeFirst && accents.length && !isNeutral(accents[0])) {
        var hs3 = hexHsl(accents[0]);
        clusters.unshift({ hue: hs3[0], l: hs3[2], s: hs3[1], weight: 1e9, hex: accents[0], mem: [] });
    }
    clusters.sort(function (x, y) { return y.weight - x.weight; });

    // 中性色两端
    var darks = [], lights = [];
    for (i = 0; i < neutral.length; i++) (relLum(neutral[i].hex) < 0.42 ? darks : lights).push(neutral[i]);
    var nd = '3A3A3A', nl = 'F2F2F2';
    if (darks.length) { nd = darks.sort(function (a, b) { return b.area - a.area; })[0].hex; }
    if (lights.length) { nl = lights.sort(function (a, b) { return b.area - a.area; })[0].hex; }

    return {
        chroma: clusters.map(function (c) { return { hex: c.hex, hue: c.hue, l: c.l, s: c.s }; }),
        neutralDark: nd, neutralLight: nl
    };
}

/* ------------------------------ 映射 ------------------------------ */
// 给定源色族与目标色，算出每个成员的新颜色（保留层次）
function mapCluster(mem, tgt, depth) {
    var out = [], i;
    var ls = [], ss = [], tot = 0;
    for (i = 0; i < mem.length; i++) {
        var hs = hexHsl(mem[i].hex);
        ls.push(hs[2]); ss.push(hs[1]);
        tot += (KIND_W[mem[i].kind] || 0.2);
    }
    var meanL = 0;
    for (i = 0; i < mem.length; i++) meanL += ls[i] * (KIND_W[mem[i].kind] || 0.2);
    meanL /= (tot || 1);
    var meanS = 0;
    for (i = 0; i < mem.length; i++) meanS += ss[i];
    meanS = (meanS / mem.length) || 0.001;

    for (i = 0; i < mem.length; i++) {
        var nl = Math.max(0.08, Math.min(0.94, tgt.l + (ls[i] - meanL) * depth));
        var ratio = ss[i] / meanS;
        var ns = Math.max(0.10, Math.min(0.98, tgt.s * (0.55 + 0.45 * Math.min(2, ratio))));
        out.push({ item: mem[i], hex: hslHex(tgt.hue, ns, nl) });
    }
    // 渐变保险：两端明暗差太小就撑开
    var grads = {};
    for (i = 0; i < mem.length; i++) if (mem[i].grad) (grads[mem[i].grad] = grads[mem[i].grad] || []).push(out[i]);
    for (var g in grads) {
        var arr = grads[g];
        if (arr.length < 2) continue;
        arr.sort(function (a, b) { return hexHsl(b.item.hex)[2] - hexHsl(a.item.hex)[2]; });
        var l0 = hexHsl(arr[0].hex)[2], l1 = hexHsl(arr[arr.length - 1].hex)[2];
        if (Math.abs(l0 - l1) >= 0.16) continue;
        var baseH = hexHsl(arr[0].hex)[0], baseS = hexHsl(arr[0].hex)[1];
        var center = (l0 + l1) / 2, span = 0.16;
        for (i = 0; i < arr.length; i++) {
            var t = arr.length === 1 ? 0 : i / (arr.length - 1);
            arr[i].hex = hslHex(baseH, baseS, Math.max(0.10, Math.min(0.92, center + span / 2 - t * span)));
        }
    }
    return out;
}

function mapNeutral(mem, nd, nl, primaryHue) {
    var out = [], ls = [], i;
    for (i = 0; i < mem.length; i++) ls.push(hexHsl(mem[i].hex)[2]);
    var lo = Math.min.apply(null, ls), hi = Math.max.apply(null, ls);
    var ndL = hexHsl(nd)[2], nlL = hexHsl(nl)[2];
    for (i = 0; i < mem.length; i++) {
        var li = ls[i], hex;
        if (li > 0.965) hex = 'FFFFFF';
        else if (li < 0.045) hex = '000000';
        else {
            var t = hi > lo ? (li - lo) / (hi - lo) : 0.5;
            var s = Math.max(0.02, Math.min(0.12, hexHsl(mem[i].hex)[1] * 0.6 + 0.04));
            hex = hslHex(primaryHue, s, Math.max(0.03, Math.min(0.97, ndL + t * (nlL - ndL))));
        }
        out.push({ item: mem[i], hex: hex });
    }
    return out;
}

/* ------------------------------ 应用 ------------------------------ */
function applyHex(item, hex) {
    var c = hex2rgb(hex), v = toVba(c[0], c[1], c[2]);
    try {
        if (item.kind === 'stop' && item.el && item.el.Color) { item.el.Color.RGB = v; return true; }
        if (item.kind === 'fill' && item.el) { item.el.ForeColor.RGB = v; return true; }
        if (item.kind === 'line' && item.el) { item.el.ForeColor.RGB = v; return true; }
        if (item.kind === 'text' && item.el) { item.el.RGB = v; return true; }
    } catch (e) { }
    return false;
}

// 收集目标对象里「可直接改」的颜色对象引用（同时记录用途，便于映射后写回）
function collectWritable(shapes) {
    var out = [];
    for (var i = 0; i < shapes.length; i++) {
        eachShape(shapes[i], function (s) {
            try {
                if (s.Fill && s.Fill.Type !== MSO_FILL_BACKGROUND) {
                    if (s.Fill.Type === MSO_FILL_GRADIENT) {
                        var gid = 'g' + Math.random();
                        try {
                            var stops = s.Fill.GradientColorStops;
                            for (var k = 1; k <= stops.Count; k++) {
                                var v = stops.Item(k).Color.RGB;
                                out.push({
                                    hex: rgb2hex(vbaR(v), vbaG(v), vbaB(v)), kind: 'stop',
                                    area: shapeArea(s), el: stops.Item(k).Color, grad: gid, owner: s
                                });
                            }
                        } catch (e) {
                            var v2 = s.Fill.ForeColor.RGB;
                            out.push({ hex: rgb2hex(vbaR(v2), vbaG(v2), vbaB(v2)), kind: 'fill', area: shapeArea(s), el: s.Fill, owner: s });
                        }
                    } else {
                        var v3 = s.Fill.ForeColor.RGB;
                        out.push({ hex: rgb2hex(vbaR(v3), vbaG(v3), vbaB(v3)), kind: 'fill', area: shapeArea(s), el: s.Fill, owner: s });
                    }
                }
            } catch (e) { }
            try {
                if (s.Line && s.Line.Visible) {
                    var v4 = s.Line.ForeColor.RGB;
                    out.push({ hex: rgb2hex(vbaR(v4), vbaG(v4), vbaB(v4)), kind: 'line', area: shapeArea(s), el: s.Line, owner: s });
                }
            } catch (e) { }
            try {
                if (s.TextFrame && s.TextFrame.HasText) {
                    var v5 = s.TextFrame.TextRange.Font.Color.RGB;
                    out.push({ hex: rgb2hex(vbaR(v5), vbaG(v5), vbaB(v5)), kind: 'text', area: shapeArea(s), el: s.TextFrame.TextRange.Font.Color, owner: s });
                }
            } catch (e) { }
            // 表格单元格
            try {
                if (s.Type === MSO_TABLE && s.Table) {
                    for (var r = 1; r <= s.Table.Rows.Count; r++) {
                        for (var cc = 1; cc <= s.Table.Columns.Count; cc++) {
                            try {
                                var f = s.Table.Cell(r, cc).Shape.Fill;
                                if (f && f.Type !== MSO_FILL_BACKGROUND) {
                                    var v6 = f.ForeColor.RGB;
                                    out.push({ hex: rgb2hex(vbaR(v6), vbaG(v6), vbaB(v6)), kind: 'fill', area: 1, el: f, owner: s });
                                }
                            } catch (e) { }
                        }
                    }
                }
            } catch (e) { }
            // 图表系列
            try {
                if (s.HasChart && s.Chart) {
                    var ser = s.Chart.SeriesCollection();
                    for (var q = 1; q <= ser.Count; q++) {
                        try {
                            var ff = ser.Item(q).Format.Fill;
                            var v7 = ff.ForeColor.RGB;
                            out.push({ hex: rgb2hex(vbaR(v7), vbaG(v7), vbaB(v7)), kind: 'fill', area: shapeArea(s), el: ff, owner: s });
                        } catch (e) { }
                    }
                }
            } catch (e) { }
        });
    }
    return out;
}

/* ------------------------------ 主流程 ------------------------------ */
var CFG = { depth: 0.85, useThemeFirst: false, keepWhiteBlack: true };

function runRecolor(shapes, slide) {
    if (!slide) { notify('拿不到当前页，先切到普通视图再运行。'); return; }
    if (!shapes || shapes.length === 0) { notify('没有选中任何对象，本页也没有形状。'); return; }

    var pal = buildPalette(slide, shapes, CFG.useThemeFirst);
    var items = collectWritable(shapes);
    var res = cluster(items);
    var clusters = mergeGradient(res.clusters);

    var changes = [], k, i;
    for (k = 0; k < clusters.length; k++) {
        var tgt = pal.chroma[k % pal.chroma.length];
        var cycle = Math.floor(k / pal.chroma.length);
        var t = { hue: tgt.hue, s: tgt.s, l: tgt.l };
        if (cycle) t.l = Math.max(0.22, Math.min(0.86, t.l + (cycle % 2 ? 0.16 : -0.16) * cycle));
        changes = changes.concat(mapCluster(clusters[k].mem, t, CFG.depth));
    }
    if (res.neutral.length) {
        changes = changes.concat(mapNeutral(res.neutral, pal.neutralDark, pal.neutralLight, pal.chroma[0].hue));
    }

    var done = 0, fail = 0, log = [];
    for (i = 0; i < changes.length; i++) {
        var ch = changes[i];
        if (ch.item.hex.toUpperCase() === ch.hex.toUpperCase()) continue;
        if (applyHex(ch.item, ch.hex)) { done++; log.push(ch.item.hex + ' → ' + ch.hex); }
        else fail++;
    }

    var msg = '取到的配色（当页为主）\n  主色 #' + pal.chroma[0].hex +
        '\n  色序 ' + pal.chroma.map(function (c) { return '#' + c.hex; }).join(' ') +
        '\n  灰阶 #' + pal.neutralDark + ' ~ #' + pal.neutralLight +
        '\n\n已改写 ' + done + ' 处颜色' + (fail ? '（' + fail + ' 处该版本 WPS 不支持改写，已跳过）' : '') +
        '\n\n' + log.slice(0, 24).join('\n') + (log.length > 24 ? '\n…' : '');
    notify(msg);
}

function notify(msg) {
    try { alert(msg); } catch (e) {
        try { Application.ActivePresentation.Application.StatusBar = msg; } catch (e2) { }
    }
}

/* ------------------------------ 入口 ------------------------------ */
function recolorSelection() {
    var slide = currentSlide();
    runRecolor(selectedShapes(slide), slide);
}

function recolorLastInserted() {
    var slide = currentSlide();
    if (!slide) return;
    try { runRecolor([slide.Shapes.Item(slide.Shapes.Count)], slide); } catch (e) { notify('取不到最后插入的对象。'); }
}

function recolorWholeSlide() {
    var slide = currentSlide();
    if (!slide) return;
    var all = [];
    for (var i = 1; i <= slide.Shapes.Count; i++) all.push(slide.Shapes.Item(i));
    runRecolor(all, slide);
}

function showSlidePalette() {
    var slide = currentSlide();
    if (!slide) { notify('拿不到当前页。'); return; }
    var pal = buildPalette(slide, [], CFG.useThemeFirst);
    notify('当前页取到的配色\n\n主色 #' + pal.chroma[0].hex +
        '\n色序 ' + pal.chroma.map(function (c) { return '#' + c.hex; }).join(' ') +
        '\n灰阶 #' + pal.neutralDark + ' ~ #' + pal.neutralLight +
        '\n主题色 ' + themeAccents(slide).map(function (h) { return '#' + h; }).join(' '));
}

// 一键：选中→改色。绑到工具栏按钮时就用这个。
function PPT配色自适应() { recolorSelection(); }
