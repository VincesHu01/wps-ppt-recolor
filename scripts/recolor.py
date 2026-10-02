#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
recolor.py —— PPT/WPS 演示「模板配色自适应」引擎（本地通道，零第三方依赖）

用途：把某一页中插入的模板内容（图标 / 图表 / 关系图 / 流程图 / SmartArt / 表格…）
      自动改成当前演示文稿的整体配色，尤其是**当页**的实际配色。

原理：
  1. 取色 —— 从 ppt/theme/theme*.xml 读主题色（整体配色），
             从当页（可选：版式页 / 母版页 / 全篇）已有形状读实际用色（当页配色），
             按「面积 × 用途权重」聚类出主色 / 辅色 / 点缀色 / 中性色。
  2. 建映射 —— 把目标对象内部的颜色按色相聚类，
             源主色族 → 当页主色，源第 k 个色族 → 当页第 k 个色族；
             族内按明度排序做「深浅派生」，保留原有的明暗层次与渐变方向；
             中性色映射到当页的深灰/浅灰，并向主色轻微靠色，避免死灰。
  3. 写回 —— 直接改写 OOXML 颜色节点（solidFill / gradFill 的每个 stop /
             线条 / 文字 / 表格底纹 / 图表 / SmartArt），不动其它任何格式。

用法：
  python3 recolor.py deck.pptx --slide 5                       # 改第 5 页最后一个插入的对象
  python3 recolor.py deck.pptx --slide 5 --target all          # 改整页
  python3 recolor.py deck.pptx --slide 5 --target name:流程     # 改名字含"流程"的形状
  python3 recolor.py deck.pptx --slide 5 --dry-run             # 只看配色报告，不改文件
  python3 recolor.py deck.pptx --slide 5 --out new.pptx        # 另存
"""

import argparse
import colorsys
import json
import math
import os
import re
import shutil
import sys
import zipfile
from xml.etree import ElementTree as ET

# ---------------------------------------------------------------- 命名空间
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
for _k, _v in NS.items():
    ET.register_namespace(_k, _v)
ET.register_namespace("", NS["p"])


def qn(ns, tag):
    return "{%s}%s" % (NS[ns], tag)


SHAPE_TAGS = (
    qn("p", "sp"), qn("p", "pic"), qn("p", "graphicFrame"),
    qn("p", "grpSp"), qn("p", "cxnSp"), qn("p", "contentPart"),
)

# 常见 prstClr 名称 → hex
PRST = {
    "aliceBlue": "F0F8FF", "antiqueWhite": "FAEBD7", "aqua": "00FFFF",
    "aquamarine": "7FFFD4", "azure": "F0FFFF", "beige": "F5F5DC",
    "bisque": "FFE4C4", "black": "000000", "blanchedAlmond": "FFEBCD",
    "blue": "0000FF", "blueViolet": "8A2BE2", "brown": "A52A2A",
    "burlyWood": "DEB887", "cadetBlue": "5F9EA0", "chartreuse": "7FFF00",
    "chocolate": "D2691E", "coral": "FF7F50", "cornflowerBlue": "6495ED",
    "cornsilk": "FFF8DC", "crimson": "DC143C", "cyan": "00FFFF",
    "darkBlue": "00008B", "darkCyan": "008B8B", "darkGoldenrod": "B8860B",
    "darkGray": "A9A9A9", "darkGreen": "006400", "darkKhaki": "BDB76B",
    "darkMagenta": "8B008B", "darkOliveGreen": "556B2F", "darkOrange": "FF8C00",
    "darkOrchid": "9932CC", "darkRed": "8B0000", "darkSalmon": "E9967A",
    "darkSeaGreen": "8FBC8F", "darkSlateBlue": "483D8B", "darkSlateGray": "2F4F4F",
    "darkTurquoise": "00CED1", "darkViolet": "9400D3", "deepPink": "FF1493",
    "deepSkyBlue": "00BFFF", "dimGray": "696969", "dodgerBlue": "1E90FF",
    "firebrick": "B22222", "floralWhite": "FFFAF0", "forestGreen": "228B22",
    "fuchsia": "FF00FF", "gainsboro": "DCDCDC", "ghostWhite": "F8F8FF",
    "gold": "FFD700", "goldenrod": "DAA520", "gray": "808080",
    "green": "008000", "greenYellow": "ADFF2F", "honeydew": "F0FFF0",
    "hotPink": "FF69B4", "indianRed": "CD5C5C", "indigo": "4B0082",
    "ivory": "FFFFF0", "khaki": "F0E68C", "lavender": "E6E6FA",
    "lavenderBlush": "FFF0F5", "lawnGreen": "7CFC00", "lemonChiffon": "FFFACD",
    "lightBlue": "ADD8E6", "lightCoral": "F08080", "lightCyan": "E0FFFF",
    "lightGoldenrodYellow": "FAFAD2", "lightGray": "D3D3D3", "lightGreen": "90EE90",
    "lightPink": "FFB6C1", "lightSalmon": "FFA07A", "lightSeaGreen": "20B2AA",
    "lightSkyBlue": "87CEFA", "lightSlateGray": "778899", "lightSteelBlue": "B0C4DE",
    "lightYellow": "FFFFE0", "lime": "00FF00", "limeGreen": "32CD32",
    "linen": "FAF0E6", "magenta": "FF00FF", "maroon": "800000",
    "mediumAquamarine": "66CDAA", "mediumBlue": "0000CD", "mediumOrchid": "BA55D3",
    "mediumPurple": "9370DB", "mediumSeaGreen": "3CB371", "mediumSlateBlue": "7B68EE",
    "mediumSpringGreen": "00FA9A", "mediumTurquoise": "48D1CC", "mediumVioletRed": "C71585",
    "midnightBlue": "191970", "mintCream": "F5FFFA", "mistyRose": "FFE4E1",
    "moccasin": "FFE4B5", "navajoWhite": "FFDEAD", "navy": "000080",
    "oldLace": "FDF5E6", "olive": "808000", "oliveDrab": "6B8E23",
    "orange": "FFA500", "orangeRed": "FF4500", "orchid": "DA70D6",
    "paleGoldenrod": "EEE8AA", "paleGreen": "98FB98", "paleTurquoise": "AFEEEE",
    "paleVioletRed": "DB7093", "papayaWhip": "FFEFD5", "peachPuff": "FFDAB9",
    "peru": "CD853F", "pink": "FFC0CB", "plum": "DDA0DD",
    "powderBlue": "B0E0E6", "purple": "800080", "red": "FF0000",
    "rosyBrown": "BC8F8F", "royalBlue": "4169E1", "saddleBrown": "8B4513",
    "salmon": "FA8072", "sandyBrown": "F4A460", "seaGreen": "2E8B57",
    "seaShell": "FFF5EE", "sienna": "A0522D", "silver": "C0C0C0",
    "skyBlue": "87CEEB", "slateBlue": "6A5ACD", "slateGray": "708090",
    "snow": "FFFAFA", "springGreen": "00FF7F", "steelBlue": "4682B4",
    "tan": "D2B48C", "teal": "008080", "thistle": "D8BFD8",
    "tomato": "FF6347", "turquoise": "40E0D0", "violet": "EE82EE",
    "wheat": "F5DEB3", "white": "FFFFFF", "whiteSmoke": "F5F5F5",
    "yellow": "FFFF00", "yellowGreen": "9ACD32",
    # DrawingML 自带的 dk1/lt1 等占位值（少数工具会写）
    "dk1": "000000", "lt1": "FFFFFF",
}


# ---------------------------------------------------------------- 颜色工具
def hex2rgb(h):
    h = (h or "000000").lstrip("#").upper()
    if len(h) == 8:      # ARGB
        h = h[2:]
    if len(h) != 6:
        h = (h + "000000")[:6]
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb2hex(r, g, b):
    return "%02X%02X%02X" % (
        max(0, min(255, int(round(r)))),
        max(0, min(255, int(round(g)))),
        max(0, min(255, int(round(b)))),
    )


def to_hls(h):
    r, g, b = hex2rgb(h)
    hh, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
    return hh * 360.0, l, s


def from_hls(h, l, s):
    h = (h % 360.0) / 360.0
    l = max(0.0, min(1.0, l))
    s = max(0.0, min(1.0, s))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return rgb2hex(r * 255, g * 255, b * 255)


def hue_dist(a, b):
    d = abs(a - b) % 360.0
    return min(d, 360.0 - d)


def rel_lum(h):
    """感知亮度（Rec.709），用于判断深浅对比"""
    r, g, b = hex2rgb(h)
    return (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255.0


def is_neutral(h, chroma_max=0.12):
    """中性色判定必须用 RGB 绝对彩度：HSL 的 S 在接近白/黑时会被放大（如 #EAF1F9 的 S 高达 0.56）"""
    r, g, b = hex2rgb(h)
    mx, mn = max(r, g, b), min(r, g, b)
    chroma = (mx - mn) / 255.0
    l = (mx + mn) / 510.0
    return chroma < chroma_max or l > 0.945 or l < 0.075


def darken(h, amount):
    hh, l, s = to_hls(h)
    return from_hls(hh, max(0.02, l - amount), s)


def lighten(h, amount):
    hh, l, s = to_hls(h)
    return from_hls(hh, min(0.98, l + amount), s)


def parse_sys_clr(el):
    v = el.get("lastClr")
    if v:
        return v.upper()
    name = (el.get("val") or "").lower()
    return {"windowtext": "000000", "window": "FFFFFF",
            "highlighttext": "FFFFFF", "highlight": "3399FF",
            "menutext": "000000", "menu": "F0F0F0",
            "captiontext": "000000", "btntext": "000000",
            "graytext": "808080", "infotext": "000000",
            "3dlight": "C0C0C0", "btnhighlight": "FFFFFF"}.get(name, "000000")


# ---------------------------------------------------------------- XML 读写
def color_elements(root, include_scheme=False):
    """找出子树里所有颜色定义节点"""
    out = []
    for parent in root.iter():
        for child in list(parent):
            t = child.tag
            if t == qn("a", "srgbClr"):
                out.append(child)
            elif t == qn("a", "prstClr"):
                out.append(child)
            elif t == qn("a", "sysClr"):
                out.append(child)
            elif t == qn("a", "scrgbClr"):
                out.append(child)
            elif t == qn("a", "hslClr"):
                out.append(child)
            elif include_scheme and t == qn("a", "schemeClr"):
                out.append(child)
    return out


def read_color(el):
    t = el.tag
    if t == qn("a", "srgbClr"):
        return (el.get("val") or "000000").upper()
    if t == qn("a", "prstClr"):
        return PRST.get(el.get("val") or "", "808080").upper()
    if t == qn("a", "sysClr"):
        return parse_sys_clr(el).upper()
    if t == qn("a", "scrgbClr"):
        f = lambda k: max(0, min(100, int(el.get(k) or 0)))
        return rgb2hex(f("r") * 2.55, f("g") * 2.55, f("b") * 2.55)
    if t == qn("a", "hslClr"):
        hue = int(el.get("hue") or 0) / 60000.0
        sat = int(el.get("sat") or 0) / 1000.0
        lum = int(el.get("lum") or 0) / 1000.0
        return from_hls(hue, lum, sat)
    if t == qn("a", "schemeClr"):
        return (el.get("val") or "").upper()
    return "000000"


def write_color(el, hexval):
    """把任意颜色节点改写成 a:srgbClr，保留 alpha/lumMod/lumOff 等修饰子节点"""
    hexval = hexval.upper().lstrip("#")
    if len(hexval) == 8:
        hexval = hexval[2:]
    keep = []
    for ch in list(el):
        if ch.tag in (qn("a", "alpha"), qn("a", "lumMod"), qn("a", "lumOff"),
                      qn("a", "tint"), qn("a", "shade"), qn("a", "satMod"),
                      qn("a", "satOff"), qn("a", "hueMod"), qn("a", "hueOff"),
                      qn("a", "red"), qn("a", "green"), qn("a", "blue"),
                      qn("a", "gamma"), qn("a", "invGamma"), qn("a", "gray"),
                      qn("a", "comp"), qn("a", "inv"), qn("a", "alphaMod"),
                      qn("a", "alphaOff"), qn("a", "blend")):
            keep.append(ch)
    for ch in list(el):
        el.remove(ch)
    for ch in keep:
        el.append(ch)
    el.tag = qn("a", "srgbClr")
    el.attrib.clear()
    el.set("val", hexval)


def has_modifier(el):
    for ch in el:
        if ch.tag in (qn("a", "lumMod"), qn("a", "lumOff"),
                      qn("a", "tint"), qn("a", "shade")):
            return True
    return False


def shape_extent(el):
    """返回形状面积（EMU²），取不到返回 None"""
    for xfrm in el.iter(qn("a", "xfrm")):
        ext = xfrm.find(qn("a", "ext"))
        if ext is not None:
            try:
                w = int(ext.get("cx") or 0)
                h = int(ext.get("cy") or 0)
                if w > 0 and h > 0:
                    return w * h
            except ValueError:
                pass
        break
    for ext in el.iter(qn("a", "ext")):
        try:
            w = int(ext.get("cx") or 0)
            h = int(ext.get("cy") or 0)
            if w > 0 and h > 0:
                return w * h
        except ValueError:
            pass
    return None


def shape_name(el):
    nv = None
    for tag in ("nvSpPr", "nvPicPr", "nvGraphicFramePr", "nvGrpSpPr", "nvCxnSpPr"):
        nv = el.find(qn("p", tag))
        if nv is not None:
            break
    if nv is None:
        for c in el.iter(qn("p", "cNvPr")):
            return c.get("name") or "", c.get("id") or ""
        return "", ""
    c = nv.find(qn("p", "cNvPr"))
    return (c.get("name") or "", c.get("id") or "") if c is not None else ("", "")


def is_placeholder(el):
    for ph in el.iter(qn("p", "ph")):
        return True
    return False


def top_shapes(sp_tree):
    return [c for c in list(sp_tree) if c.tag in SHAPE_TAGS]


# ---------------------------------------------------------------- 容器
class Deck(object):
    def __init__(self, path):
        self.path = path
        self.z = zipfile.ZipFile(path, "r")
        self.parts = {}          # zip 内路径 → bytes
        self.trees = {}          # zip 内路径 → Element
        for n in self.z.namelist():
            self.parts[n] = self.z.read(n)

    def xml(self, name):
        if name not in self.trees:
            if name not in self.parts:
                return None
            self.trees[name] = ET.fromstring(self.parts[name].decode("utf-8"))
        return self.trees[name]

    def slide_names(self):
        return sorted(
            [n for n in self.parts if re.match(r"ppt/slides/slide\d+\.xml$", n)],
            key=lambda s: int(re.search(r"(\d+)", os.path.basename(s)).group(1)),
        )

    def rels_targets(self, part_name):
        """返回 {rId: 包内绝对路径}"""
        d = os.path.dirname(part_name)
        rp = "%s/_rels/%s.rels" % (d, os.path.basename(part_name))
        if rp not in self.parts:
            return {}
        root = ET.fromstring(self.parts[rp].decode("utf-8"))
        out = {}
        for rel in root.findall(qn("rel", "Relationship")):
            tgt = rel.get("Target") or ""
            if tgt.startswith("/"):
                out[rel.get("Id")] = tgt.lstrip("/")
            else:
                out[rel.get("Id")] = os.path.normpath(os.path.join(d, tgt)).replace("\\", "/")
        return out

    def theme_palette(self):
        """从主题读整体配色：dk1 lt1 dk2 lt2 accent1..6"""
        out = []
        for n in sorted([x for x in self.parts if re.match(r"ppt/theme/theme\d+\.xml$", x)]):
            root = self.xml(n)
            if root is None:
                continue
            cs = root.find(".//" + qn("a", "clrScheme"))
            if cs is None:
                continue
            for ch in cs:
                if not ch.tag.startswith(qn("a", "")):
                    continue
                key = ch.tag.split("}")[1]
                if key not in ("dk1", "lt1", "dk2", "lt2",
                               "accent1", "accent2", "accent3",
                               "accent4", "accent5", "accent6"):
                    continue
                for cc in ch:
                    if cc.tag in (qn("a", "srgbClr"), qn("a", "sysClr"),
                                  qn("a", "prstClr"), qn("a", "scrgbClr"),
                                  qn("a", "hslClr")):
                        hexv = read_color(cc)
                        if hexv:
                            out.append((key, hexv))
                        break
        return out

    def save(self, out_path):
        tmp = out_path + ".tmp"
        zin = self.z
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
            for item in zin.infolist():
                data = self.parts[item.filename]
                if item.filename in self.trees and self.trees[item.filename] is not None:
                    data = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + \
                        ET.tostring(self.trees[item.filename], encoding="utf-8")
                zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
                zi.compress_type = item.compress_type
                zi.external_attr = item.external_attr
                zo.writestr(zi, data)
        zin.close()
        os.replace(tmp, out_path)


# ---------------------------------------------------------------- 取色
KIND_WEIGHT = {"fill": 1.0, "stop": 0.75, "line": 0.35, "text": 0.25, "other": 0.15}


def build_parent_map(root):
    pm = {}
    for p in root.iter():
        for c in list(p):
            pm[c] = p
    return pm


def harvest(shape_el, parent_map=None, default_area=None):
    """收集一个形状（含组内子形状，递归）里的所有颜色元素及其视觉权重"""
    items = []
    pm = parent_map if parent_map is not None else build_parent_map(shape_el)
    base_area = default_area or shape_extent(shape_el) or 10 ** 10

    for col in color_elements(shape_el, include_scheme=False):
        hexv = read_color(col)
        if not hexv:
            continue
        kinds, area, grad_id = [], base_area, None
        node = pm.get(col)
        while node is not None:
            t = node.tag
            if grad_id is None and t == qn("a", "gradFill"):
                grad_id = id(node)
            if t == qn("a", "ln"):
                kinds.append("line")
            elif t in (qn("a", "rPr"), qn("a", "defRPr"), qn("a", "endParaRPr")):
                kinds.append("text")
            elif t == qn("a", "gradFill"):
                kinds.append("stop")
            elif t in (qn("a", "solidFill"), qn("a", "pattFill"),
                       qn("a", "blipFill"), qn("a", "bgFillStyle"),
                       qn("a", "fillStyleLst"), qn("a", "highlight")):
                kinds.append("fill")
            elif t in (qn("a", "effectLst"), qn("a", "effectDag")):
                kinds.append("other")
            if t in SHAPE_TAGS:
                area = shape_extent(node) or area
                break
            node = pm.get(node)
        kind = kinds[-1] if kinds else "other"
        w = KIND_WEIGHT.get(kind, 0.2) * (max(0.05, area / 1e12)) ** 0.45
        items.append({"el": col, "hex": hexv, "kind": kind, "grad": grad_id,
                      "weight": w, "mod": has_modifier(col)})
    return items


def cluster(items):
    """把颜色按色相聚类，返回 (chroma_clusters, neutral_items)"""
    chroma, neutral = [], []
    for it in items:
        if is_neutral(it["hex"]):
            neutral.append(it)
        else:
            chroma.append(it)
    bins = {}
    for it in chroma:
        h, l, s = to_hls(it["hex"])
        key = int(h // 12)
        bins.setdefault(key, []).append((it, h, l, s))
    # 合并相邻色相桶（< 26°）
    keys = sorted(bins.keys())
    groups = []
    for k in keys:
        members = bins[k]
        mh = sum(x[1] for x in members) / len(members)
        if groups:
            pmh = groups[-1]["h"]
            if hue_dist(pmh, mh) < 26:
                groups[-1]["members"] += members
                groups[-1]["h"] = (pmh + mh) / 2
                continue
        groups.append({"h": mh, "members": members})
    if len(groups) > 1 and hue_dist(groups[0]["h"], groups[-1]["h"]) < 26:
        groups[-1]["members"] += groups[0]["members"]
        groups.pop(0)

    clusters = []
    for g in groups:
        mem = g["members"]
        tot = sum(m[0]["weight"] for m in mem) or 1.0
        # 加权圆均值色相
        sx = sy = sl = ss = 0.0
        for it, h, l, s in mem:
            w = it["weight"] / tot
            sx += math.cos(math.radians(h)) * w
            sy += math.sin(math.radians(h)) * w
            sl += l * w
            ss += s * w
        mh = math.degrees(math.atan2(sy, sx)) % 360
        clusters.append({
            "hue": mh, "l": sl, "s": ss, "weight": tot,
            "hex": from_hls(mh, sl, ss),
            "members": [m[0] for m in mem],
        })
    clusters.sort(key=lambda c: -c["weight"])
    return clusters, neutral


# ---------------------------------------------------------------- 映射
class Palette(object):
    def __init__(self):
        self.chroma = []      # [{'hex','hue','l','s'}]
        self.neutral_dark = "3A3A3A"
        self.neutral_light = "F2F2F2"
        self.primary = "2F6FB5"


def build_palette(deck, slide_part, exclude_els, scope, args_priority="page",
                  harmony_max=90):
    """取色：当页为主，母版/版式/全篇补充"""
    slide_root = deck.xml(slide_part)
    sp_tree = slide_root.find(".//" + qn("p", "spTree"))

    page_items = []
    for sh in top_shapes(sp_tree):
        if any(sh is e for e in exclude_els):
            continue
        page_items += harvest(sh)
    page_items += _background_colors(slide_root)

    theme = deck.theme_palette()
    theme_map = {k: v for k, v in theme}

    layout_items = []
    if scope in ("layout", "deck"):
        for rid, tgt in deck.rels_targets(slide_part).items():
            if re.search(r"slideLayout\d+\.xml$", tgt):
                lr = deck.xml(tgt)
                if lr is None:
                    continue
                lst = lr.find(".//" + qn("p", "spTree"))
                if lst is None:
                    continue
                for sh in top_shapes(lst):
                    layout_items += harvest(sh)
    deck_items = []
    if scope == "deck":
        for sn in deck.slide_names():
            if sn == slide_part:
                continue
            sr = deck.xml(sn)
            st = sr.find(".//" + qn("p", "spTree")) if sr is not None else None
            if st is None:
                continue
            for sh in top_shapes(st):
                deck_items += harvest(sh)

    pal = Palette()
    clusters, neutral = cluster(page_items)
    if len(clusters) < 2 and layout_items:
        c2, n2 = cluster(layout_items)
        clusters += c2
        neutral += n2
    if len(clusters) < 2 and deck_items:
        c3, n3 = cluster(deck_items)
        clusters += c3
        neutral += n3

    # 排序：权重降序；权重打平时（如几个等大色块）优先选更接近主题主色的色相
    ah = None
    for key in ("accent1", "accent2", "accent3", "dk2"):
        v = theme_map.get(key)
        if v and not is_neutral(v):
            ah = to_hls(v)[0]
            break
    clusters.sort(key=lambda c: (-c["weight"],
                                 hue_dist(c["hue"], ah) if ah is not None else 0.0))
    if args_priority == "theme":
        for key in ("accent1", "accent2", "accent3"):
            v = theme_map.get(key)
            if v and not is_neutral(v):
                h, l, s = to_hls(v)
                clusters.insert(0, {"hue": h, "l": l, "s": s,
                                    "weight": max([c["weight"] for c in clusters] or [1]) + 1,
                                    "hex": v, "members": []})
                break

    # 主题 accent 补位：只补「与当页已有色族还算协调」的色。
    # 排序按与已有色族的最小色相距离升序；太跳的（> HARMONY_MAX）宁可不补，
    # 让源色族走主色的深浅派生 —— 往一份只用蓝青的稿子里硬塞主题里的橙色是帮倒忙。
    HARMONY_MAX = harmony_max
    cands = []
    for key in ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]:
        v = theme_map.get(key)
        if not v:
            continue
        h, l, s = to_hls(v)
        if s < 0.12 or l > 0.95 or l < 0.08:
            continue
        if clusters:
            d = min(hue_dist(h, c["hue"]) for c in clusters)
        else:
            d = 0.0
        if d <= 25:
            continue
        cands.append((d, h, l, s, v))
    for d, h, l, s, v in sorted(cands)[:3]:
        if d > HARMONY_MAX:
            break
        clusters.append({"hue": h, "l": l, "s": s, "weight": 0.001,
                         "hex": v, "members": []})
    if not clusters:
        for key in ["accent1", "accent2", "accent3", "dk2", "lt2"]:
            v = theme_map.get(key)
            if v and not is_neutral(v):
                h, l, s = to_hls(v)
                clusters.append({"hue": h, "l": l, "s": s, "weight": 1,
                                 "hex": v, "members": []})
                break
    if not clusters:
        clusters = [{"hue": 210, "l": 0.45, "s": 0.6, "weight": 1,
                     "hex": "2F6FB5", "members": []}]

    pal.chroma = [{"hex": c["hex"], "hue": c["hue"], "l": c["l"], "s": c["s"]}
                  for c in clusters]
    pal.primary = pal.chroma[0]["hex"]

    # 中性色
    darks = [n for n in neutral if rel_lum(n["hex"]) < 0.42]
    lights = [n for n in neutral if rel_lum(n["hex"]) >= 0.42]
    if darks:
        pal.neutral_dark = max(darks, key=lambda n: n["weight"])["hex"]
    else:
        pal.neutral_dark = theme_map.get("dk1") or theme_map.get("dk2") or "3A3A3A"
    if lights:
        pal.neutral_light = max(lights, key=lambda n: n["weight"])["hex"]
    else:
        pal.neutral_light = theme_map.get("lt1") or theme_map.get("lt2") or "F2F2F2"
    if is_neutral(pal.neutral_light) is False:
        pass
    return pal, page_items, theme_map


def _background_colors(slide_root):
    out = []
    bg = slide_root.find(qn("p", "cSld") + "/" + qn("p", "bg"))
    if bg is not None:
        for col in color_elements(bg):
            out.append({"el": col, "hex": read_color(col), "kind": "fill",
                        "weight": 3.0, "mod": has_modifier(col)})
    return out


def _cluster_mean(members):
    tot = sum(m["weight"] for m in members) or 1.0
    sx = sy = sl = ss = 0.0
    for m in members:
        h, l, s = to_hls(m["hex"])
        w = m["weight"] / tot
        sx += math.cos(math.radians(h)) * w
        sy += math.sin(math.radians(h)) * w
        sl += l * w
        ss += s * w
    mh = math.degrees(math.atan2(sy, sx)) % 360
    return mh, sl, ss, from_hls(mh, sl, ss), tot


def merge_gradient_clusters(clusters):
    """同一个渐变（gradFill）里的所有 stop 必须落在同一色族 —— 否则渐变会被改成两个色相"""
    grad_map = {}
    for i, c in enumerate(clusters):
        for m in c["members"]:
            g = m.get("grad")
            if g:
                grad_map.setdefault(g, set()).add(i)
    groups = [sorted(s) for s in grad_map.values() if len(s) > 1]
    if not groups:
        return clusters
    touched = set()
    for g in groups:
        touched.update(g)
    keep = [c for i, c in enumerate(clusters) if i not in touched]
    for g in groups:
        g = sorted(g, key=lambda i: -clusters[i]["weight"])
        members = []
        for i in g:
            members += clusters[i]["members"]
        mh, sl, ss, hexv, tot = _cluster_mean(members)
        keep.append({"hue": mh, "l": sl, "s": ss, "weight": tot,
                     "hex": hexv, "members": members})
    keep.sort(key=lambda c: -c["weight"])
    return keep


def _enforce_gradient_spread(mapping, src_clusters, min_spread=0.16):
    """渐变保险：同一 gradFill 的 stop 若映射后明暗差太小，撑开到至少 min_spread"""
    grads = {}
    for cl in src_clusters:
        for m in cl["members"]:
            g = m.get("grad")
            if g and m.get("_new"):
                grads.setdefault(g, []).append(m)
    for g, mem in grads.items():
        if len(mem) < 2:
            continue
        mem = sorted(mem, key=lambda m: -to_hls(m["hex"])[1])
        ls = [to_hls(m["_new"])[1] for m in mem]
        if max(ls) - min(ls) >= min_spread:
            continue
        base = mem[0]["_new"]
        bh, bl, bs = to_hls(sum_hex(mem))
        span = min(0.5, min_spread)
        for i, m in enumerate(mem):
            t = i / (len(mem) - 1.0)
            nl = max(0.10, min(0.92, bl + span / 2 - t * span))
            nh, _, ns = to_hls(m["_new"])
            new_hex = from_hls(nh, nl, ns)
            mapping[id(m["el"])] = (m["el"], new_hex)
            m["_new"] = new_hex


def sum_hex(mem):
    """一组颜色的加权平均色（用于取渐变的基准色相）"""
    tot = sum(m["weight"] for m in mem) or 1.0
    sx = sy = sl = ss = 0.0
    for m in mem:
        h, l, s = to_hls(m["hex"])
        w = m["weight"] / tot
        sx += math.cos(math.radians(h)) * w
        sy += math.sin(math.radians(h)) * w
        sl += l * w
        ss += s * w
    return from_hls(math.degrees(math.atan2(sy, sx)) % 360, sl, ss)


def build_map(src_clusters, src_neutral, pal, depth, keep_white_black,
              args_spread=0.16):
    """生成 {颜色元素: 新hex} 映射 + 可读报告"""
    mapping = {}
    report = []
    targets = pal.chroma
    nT = len(targets)

    for k, cl in enumerate(src_clusters):
        tgt = targets[k % nT]
        cycle = k // nT
        base_h = tgt["hue"]
        base_l = tgt["l"]
        base_s = tgt["s"]
        if cycle:
            # 溢出：沿明度轴派生，保持色相，避免撞色
            base_l = min(0.86, max(0.22, base_l + (0.16 if cycle % 2 else -0.16) * cycle))

        mem = sorted(cl["members"], key=lambda m: -to_hls(m["hex"])[1])
        ls = [to_hls(m["hex"])[1] for m in mem]
        ss = [to_hls(m["hex"])[2] for m in mem]
        # 以源簇的加权平均明度为锚，对齐到目标主色明度；
        # 每个成员按自己相对锚点的偏离量 * depth 派生 —— 保留原有的深浅层次与渐变方向
        tot_w = sum(m["weight"] for m in mem) or 1.0
        mean_l = sum(l * m["weight"] for l, m in zip(ls, mem)) / tot_w
        mean_s = sum(ss) / len(ss) or 0.001

        for i, m in enumerate(mem):
            li, si = ls[i], ss[i]
            new_l = base_l + (li - mean_l) * depth
            ratio = si / mean_s if mean_s > 0.001 else 1.0
            new_s = max(0.10, min(0.98, base_s * (0.55 + 0.45 * min(2.0, ratio))))
            new_hex = from_hls(base_h, max(0.08, min(0.94, new_l)), new_s)
            if m["mod"]:
                # 带 lumMod/tint 修饰的节点：只换色相与饱和度，明度交给原修饰
                new_hex = from_hls(base_h, max(0.05, min(0.95, base_l)), new_s)
            mapping[id(m["el"])] = (m["el"], new_hex)
            m["_new"] = new_hex
            report.append({"from": m["hex"], "to": new_hex, "kind": m["kind"],
                           "group": "chroma#%d" % k})

    _enforce_gradient_spread(mapping, src_clusters, min_spread=args_spread)

    # 中性色
    nd_h, nd_l, nd_s = to_hls(pal.neutral_dark)
    nl_h, nl_l, nl_s = to_hls(pal.neutral_light)
    ph = pal.chroma[0]["hue"]
    if src_neutral:
        ls = [to_hls(n["hex"])[1] for n in src_neutral]
        lo, hi = min(ls), max(ls)
        for n in src_neutral:
            li, si, hi_h = to_hls(n["hex"])
            if keep_white_black and li > 0.965:
                new_hex = "FFFFFF"
            elif keep_white_black and li < 0.045:
                new_hex = "000000"
            else:
                t = (li - lo) / (hi - lo) if hi > lo else 0.5
                new_l = nd_l + t * (nl_l - nd_l)
                new_s = max(0.02, min(0.12, si * 0.6 + 0.04))
                new_hex = from_hls(ph, max(0.03, min(0.97, new_l)), new_s)
            mapping[id(n["el"])] = (n["el"], new_hex)
            report.append({"from": n["hex"], "to": new_hex, "kind": n["kind"],
                           "group": "neutral"})
    return mapping, report


# ---------------------------------------------------------------- 主流程
def pick_targets(sp_tree, mode):
    shapes = top_shapes(sp_tree)
    if not shapes:
        return []
    if mode == "all":
        return shapes
    if mode.startswith("name:"):
        kw = mode[5:]
        return [s for s in shapes if kw in shape_name(s)[0]]
    if mode.startswith("id:"):
        sid = mode[3:]
        return [s for s in shapes if shape_name(s)[1] == sid]
    if mode.startswith("idx:"):
        i = int(mode[4:])
        return [shapes[i]] if -len(shapes) <= i < len(shapes) else []
    # last：Z 序最后插入的非占位形状；找不到就取最后一个
    for s in reversed(shapes):
        if not is_placeholder(s):
            return [s]
    return [shapes[-1]]


def extra_parts_for(deck, slide_part, targets):
    """目标含图表 / SmartArt 时，返回需要连带处理的部件名"""
    out = set()
    rels = deck.rels_targets(slide_part)
    for sh in targets:
        for gd in sh.iter(qn("a", "graphicData")):
            uri = gd.get("uri") or ""
            if uri.endswith("/chart") or uri.endswith("/chartEx"):
                for ch in gd.iter():
                    cid = ch.get(qn("r", "id"))
                    if cid and cid in rels:
                        out.add(rels[cid])
            elif uri.endswith("/diagram"):
                for relids in sh.iter(qn("dgm", "relIds")):
                    for k, v in relids.attrib.items():
                        if k.endswith("cs") or k.endswith("qs") or k.endswith("dm"):
                            rid = v
                            if rid in rels and re.search(r"diagrams/(colors|quickStyle)", rels[rid]):
                                out.add(rels[rid])
    return sorted(out)


def recolor(args):
    deck = Deck(args.file)
    slides = deck.slide_names()
    if not slides:
        sys.exit("错误：文件里没有找到幻灯片。")

    idx = args.slide - 1
    if not (0 <= idx < len(slides)):
        sys.exit("错误：--slide %d 越界，共 %d 页。" % (args.slide, len(slides)))
    slide_part = slides[idx]
    slide_root = deck.xml(slide_part)
    sp_tree = slide_root.find(".//" + qn("p", "spTree"))
    if sp_tree is None:
        sys.exit("错误：该页没有形状树。")

    targets = pick_targets(sp_tree, args.target)
    if not targets:
        sys.exit("错误：--target %s 没有匹配到任何形状。" % args.target)

    excluded = list(targets)
    for sel in (args.exclude or []):
        excluded += pick_targets(sp_tree, sel)

    pal, page_items, theme_map = build_palette(
        deck, slide_part, excluded, args.scope, args.priority, args.harmony_max)

    # 收集目标对象自己的颜色（含连带部件）
    src_items = []
    for sh in targets:
        src_items += harvest(sh)
    extras = extra_parts_for(deck, slide_part, targets)
    extra_items = []
    for part in extras:
        root = deck.xml(part)
        if root is None:
            continue
        for col in color_elements(root, include_scheme=args.include_scheme):
            hexv = read_color(col)
            if not hexv:
                continue
            extra_items.append({"el": col, "hex": hexv, "kind": "fill",
                                "weight": 1.0, "mod": has_modifier(col)})
    src_items += extra_items

    src_clusters, src_neutral = cluster(src_items)
    src_clusters = merge_gradient_clusters(src_clusters)
    # 目标对象内部：按视觉权重降序（主色族先映射）
    src_clusters.sort(key=lambda c: -c["weight"])

    mapping, report = build_map(src_clusters, src_neutral, pal,
                                args.depth, not args.no_keep_extremes,
                                args.gradient_spread)

    changed = 0
    if not args.dry_run:
        for key, (el, hexv) in mapping.items():
            if read_color(el) != hexv:
                write_color(el, hexv)
                changed += 1
        deck.trees[slide_part] = slide_root
        for part in extras:
            deck.trees[part] = deck.xml(part)

    summary = {
        "file": os.path.abspath(args.file),
        "slide": args.slide,
        "slide_part": slide_part,
        "targets": [{"name": shape_name(s)[0], "id": shape_name(s)[1]}
                    for s in targets],
        "extra_parts": extras,
        "palette": {
            "primary": pal.primary,
            "chroma": [c["hex"] for c in pal.chroma],
            "neutral_dark": pal.neutral_dark,
            "neutral_light": pal.neutral_light,
            "theme": theme_map,
        },
        "source_clusters": [{"hex": c["hex"], "weight": round(c["weight"], 3),
                             "count": len(c["members"])} for c in src_clusters],
        "changed_color_nodes": changed,
        "mapping": report,
    }

    if not args.dry_run:
        out = args.out or args.file
        if out == args.file and not args.no_backup:
            shutil.copy2(args.file, args.file + ".recolor.bak")
        deck.save(out)
        summary["output"] = os.path.abspath(out)
    else:
        summary["output"] = None

    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        _print_text(summary)
    return summary


def _print_text(s):
    p = s["palette"]
    print("\n=== 配色自适应报告 ===")
    print("文件：%s" % s["file"])
    print("页码：第 %d 页（%s）" % (s["slide"], s["slide_part"]))
    print("目标：%s" % ", ".join("%s(#%s)" % (t["name"] or "( unnamed )", t["id"])
                                 for t in s["targets"]))
    if s["extra_parts"]:
        print("连带部件：%s" % ", ".join(s["extra_parts"]))
    print("\n-- 取到的配色（当页优先） --")
    print("  主色  : #%s" % p["primary"])
    print("  色序  : %s" % " ".join("#" + c for c in p["chroma"]))
    print("  深灰  : #%s    浅灰: #%s" % (p["neutral_dark"], p["neutral_light"]))
    if p["theme"]:
        print("  主题  : %s" % " ".join("%s=#%s" % (k, v)
                                        for k, v in sorted(p["theme"].items())))
    print("\n-- 模板内识别到的色族 --")
    for c in s["source_clusters"]:
        print("  #%s  权重 %.2f  元素 %d 个" % (c["hex"], c["weight"], c["count"]))
    print("\n-- 颜色改写（%d 处）--" % s["changed_color_nodes"])
    seen = {}
    for m in s["mapping"]:
        seen.setdefault((m["from"], m["to"], m["group"]), 0)
        seen[(m["from"], m["to"], m["group"])] += 1
    for (f, t, g), n in sorted(seen.items(), key=lambda x: -x[1]):
        print("  %s  #%s -> #%s   ×%d" % (g.ljust(9), f, t, n))
    print("\n输出：%s" % (s["output"] or "(dry-run，未写入)"))


def main():
    ap = argparse.ArgumentParser(
        description="把插入的模板内容自动改成当前 PPT/WPS 演示的配色")
    ap.add_argument("file", help="pptx 文件路径")
    ap.add_argument("--slide", type=int, default=1, help="页码，1 起（默认 1）")
    ap.add_argument("--target", default="last",
                    help="重着色对象：last | all | name:<关键字> | id:<shapeId> | idx:<序号>，默认 last")
    ap.add_argument("--scope", default="page", choices=["page", "layout", "deck"],
                    help="取色范围：page=当页；layout=当页+版式；deck=当页+版式+全篇（默认 page）")
    ap.add_argument("--priority", default="page", choices=["page", "theme"],
                    help="主色来源偏好：page=当页用得最多的颜色；theme=主题 accent1（默认 page）")
    ap.add_argument("--exclude", action="append",
                    help="取色时要排除的形状（同批插入的其它模板部分），同 --target 语法，可重复")
    ap.add_argument("--depth", type=float, default=0.85,
                    help="深浅派生强度 0~1.5，越大层次越明显（默认 0.85）")
    ap.add_argument("--out", help="另存为；省略则原地覆盖（自动生成 .recolor.bak）")
    ap.add_argument("--dry-run", action="store_true", help="只输出报告，不写文件")
    ap.add_argument("--json", action="store_true", help="以 JSON 输出报告")
    ap.add_argument("--no-backup", action="store_true", help="原地覆盖时不留备份")
    ap.add_argument("--harmony-max", type=float, default=90, metavar="DEG",
                    help="补位色与当页已有色族的最大色相距离（默认 90）。调小=只用当页色系派生；"
                         "调大=更敢引入主题里的对比色")
    ap.add_argument("--gradient-spread", type=float, default=0.16,
                    help="渐变两端明度差的最小跨度 0~0.5（默认 0.16，太小会看不出渐变）")
    ap.add_argument("--no-keep-extremes", action="store_true",
                    help="连纯白/纯黑也一起改（默认保留）")
    ap.add_argument("--include-scheme", action="store_true",
                    help="连 a:schemeClr 主题色引用也强行改成硬编码（默认不改，跟随主题）")
    args = ap.parse_args()
    if not os.path.isfile(args.file):
        sys.exit("错误：找不到文件 %s" % args.file)
    recolor(args)


if __name__ == "__main__":
    main()
