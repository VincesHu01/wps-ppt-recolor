#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_sample.py — builds the sanitized demo deck used in the README.

Everything here is synthetic: a generic quarterly review deck with a
deep blue / teal palette, plus two "vendor template" objects pasted in
with clashing colors (orange, purple, green, warm gradient).

    python3 make_sample.py

Writes sample_before.pptx next to this script. Then run the engine:

    python3 ../scripts/recolor.py sample_before.pptx --slide 2 --target last --out sample_after.pptx
    python3 ../scripts/recolor.py sample_after.pptx  --slide 3 --target last --out sample_after.pptx
"""

import os

from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))

NAVY = "1B3A5C"
TEAL = "2E8B8B"
MIST = "EDF1F5"
INK = "12283F"

# Colors a pasted template would bring along — deliberately off palette
CLASH = ["F26B21", "8E44AD", "27AE60", "F7B733", "E74C3C"]
CLASH_LINE = "7F3F00"


def slide_title(prs, text, sub):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
    bar.name = "Title band"
    bar.fill.solid()
    bar.fill.fore_color.rgb = RGBColor.from_string(NAVY)
    bar.line.fill.background()

    box = s.shapes.add_textbox(Inches(0.8), Inches(2.4), Inches(11.7), Inches(1.2))
    box.name = "Deck title"
    r = box.text_frame.paragraphs[0].add_run()
    r.text = text
    r.font.size = Pt(40)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(NAVY)

    sub_box = s.shapes.add_textbox(Inches(0.8), Inches(3.6), Inches(11.7), Inches(0.8))
    sub_box.name = "Deck subtitle"
    r2 = sub_box.text_frame.paragraphs[0].add_run()
    r2.text = sub
    r2.font.size = Pt(18)
    r2.font.color.rgb = RGBColor.from_string(TEAL)
    return s


def kpi_tile(s, x, y, w, h, fill, label, value):
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    sh.name = "KPI %s" % label
    sh.fill.solid()
    sh.fill.fore_color.rgb = RGBColor.from_string(fill)
    sh.line.fill.background()
    tf = sh.text_frame
    tf.text = "%s\n%s" % (value, label)
    for i, p in enumerate(tf.paragraphs):
        p.runs[0].font.size = Pt(26 if i == 0 else 12)
        p.runs[0].font.bold = i == 0
        p.runs[0].font.color.rgb = RGBColor.from_string("FFFFFF")
    return sh


def pasted_diagram(s):
    """A grouped relationship diagram in colors that clash with the deck."""
    shapes = []
    for i, c in enumerate(CLASH):
        sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9 + i * 2.3), Inches(4.7), Inches(1.7), Inches(1.4))
        sh.name = "Vendor node %d" % i
        sh.fill.solid()
        sh.fill.fore_color.rgb = RGBColor.from_string(c)
        sh.line.color.rgb = RGBColor.from_string(CLASH_LINE)
        sh.line.width = Pt(1.5)
        shapes.append(sh)

    arrow = s.shapes.add_shape(MSO_SHAPE.CHEVRON, Inches(0.9 + len(CLASH) * 2.3), Inches(4.7), Inches(1.7), Inches(1.4))
    arrow.name = "Vendor arrow"
    gf = arrow.fill
    gf.gradient()
    gf.gradient_stops[0].color.rgb = RGBColor.from_string("FC4A1A")
    gf.gradient_stops[1].color.rgb = RGBColor.from_string("F7B733")
    arrow.line.fill.background()
    shapes.append(arrow)
    return shapes


def clashing_chart(s):
    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("Plan", (18, 22, 21, 26))
    data.add_series("Actual", (15, 24, 19, 29))
    gf = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(0.9), Inches(1.7),
                            Inches(11.5), Inches(4.6), data)
    gf.name = "Vendor chart"
    chart = gf.chart
    for i, c in enumerate(("E74C3C", "8E44AD")):
        ser = chart.plots[0].series[i]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = RGBColor.from_string(c)
        ser.format.line.fill.background()
    return gf


def group_shapes(path, slide_index, prefix, group_name):
    """Wrap every top level shape whose name starts with `prefix` into one group.

    python-pptx cannot create groups, so this post-processes the saved file.
    It matters for the demo: a template pasted into a slide normally arrives as
    a single group, which is exactly what `--target last` selects by default.
    """
    import re
    import shutil
    import zipfile
    from xml.etree import ElementTree as ET

    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    for k, v in (("a", A), ("p", P)):
        ET.register_namespace(k, v)

    def q(ns, tag):
        return "{%s}%s" % (ns, tag)

    zin = zipfile.ZipFile(path)
    part = "ppt/slides/slide%d.xml" % slide_index
    root = ET.fromstring(zin.read(part))
    tree = root.find(q(P, "cSld")).find(q(P, "spTree"))

    picked = []
    for sp in list(tree):
        if sp.tag != q(P, "sp"):
            continue
        c = sp.find(q(P, "nvSpPr")).find(q(P, "cNvPr"))
        if (c.get("name") or "").startswith(prefix):
            picked.append(sp)

    xs, ys, xe, ye = [], [], [], []
    for sp in picked:
        xf = sp.find(q(P, "spPr")).find(q(A, "xfrm"))
        x = int(xf.find(q(A, "off")).get("x"))
        y = int(xf.find(q(A, "off")).get("y"))
        w = int(xf.find(q(A, "ext")).get("cx"))
        h = int(xf.find(q(A, "ext")).get("cy"))
        xs.append(x); ys.append(y); xe.append(x + w); ye.append(y + h)
    X, Y, X2, Y2 = min(xs), min(ys), max(xe), max(ye)

    max_id = 0
    for el in tree.iter(q(P, "cNvPr")):
        try:
            max_id = max(max_id, int(el.get("id") or 0))
        except ValueError:
            pass

    grp = ET.Element(q(P, "grpSp"))
    nv = ET.SubElement(grp, q(P, "nvGrpSpPr"))
    c = ET.SubElement(nv, q(P, "cNvPr"))
    c.set("id", str(max_id + 100))
    c.set("name", group_name)
    ET.SubElement(nv, q(P, "cNvGrpSpPr"))
    ET.SubElement(nv, q(P, "nvPr"))

    gsp = ET.SubElement(grp, q(P, "grpSpPr"))
    xf = ET.SubElement(gsp, q(A, "xfrm"))
    for tag, vals in (("off", (X, Y)), ("ext", (X2 - X, Y2 - Y)),
                      ("chOff", (X, Y)), ("chExt", (X2 - X, Y2 - Y))):
        el = ET.SubElement(xf, q(A, tag))
        el.set("x", str(int(vals[0])))
        el.set("y", str(int(vals[1])))
        el.set("cx", str(int(vals[0])))
        el.set("cy", str(int(vals[1])))
    xf.find(q(A, "ext")).set("cx", str(int(X2 - X)))
    xf.find(q(A, "ext")).set("cy", str(int(Y2 - Y)))
    xf.find(q(A, "chExt")).set("cx", str(int(X2 - X)))
    xf.find(q(A, "chExt")).set("cy", str(int(Y2 - Y)))

    for sp in picked:
        tree.remove(sp)
        grp.append(sp)
    tree.append(grp)

    new_xml = ET.tostring(root, encoding="utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zin.infolist():
            data = new_xml if item.filename == part else zin.read(item.filename)
            zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
            zi.compress_type = item.compress_type
            zi.external_attr = item.external_attr
            zo.writestr(zi, data)
    zin.close()
    shutil.move(tmp, path)


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    slide_title(prs, "Quarterly Business Review", "Sample deck — synthetic data, no real figures")

    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    head = s2.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.9))
    head.name = "Section title"
    r = head.text_frame.paragraphs[0].add_run()
    r.text = "Performance overview"
    r.font.size = Pt(30)
    r.font.bold = True
    r.font.color.rgb = RGBColor.from_string(NAVY)

    kpi_tile(s2, 0.8, 1.7, 3.7, 1.6, NAVY, "Revenue", "12.4M")
    kpi_tile(s2, 4.8, 1.7, 3.7, 1.6, TEAL, "Margin", "31%")
    kpi_tile(s2, 8.8, 1.7, 3.7, 1.6, MIST, "Churn", "2.1%")

    note = s2.shapes.add_textbox(Inches(0.8), Inches(3.6), Inches(11.7), Inches(0.7))
    note.name = "Diagram caption"
    r3 = note.text_frame.paragraphs[0].add_run()
    r3.text = "Value chain — pasted from a vendor template"
    r3.font.size = Pt(14)
    r3.font.color.rgb = RGBColor.from_string(INK)

    pasted_diagram(s2)

    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    head3 = s3.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(11.7), Inches(0.9))
    head3.name = "Section title"
    r4 = head3.text_frame.paragraphs[0].add_run()
    r4.text = "Plan versus actual"
    r4.font.size = Pt(30)
    r4.font.bold = True
    r4.font.color.rgb = RGBColor.from_string(NAVY)
    clashing_chart(s3)

    out = os.path.join(HERE, "sample_before.pptx")
    prs.save(out)
    group_shapes(out, 2, "Vendor ", "Vendor diagram (pasted)")
    print("wrote", out)


if __name__ == "__main__":
    main()
