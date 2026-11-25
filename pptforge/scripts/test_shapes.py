"""形状层验收：组合/解组、移动、缩放、属性、对齐（含真机几何回读）。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.util import Emu, Inches

from mcp_ppt import comtools
from mcp_ppt.design import Theme, Painter
from mcp_ppt.inspect import validate_deck
from mcp_ppt.shapes import (add_element, align_shapes, box_of, find_shape, group_shapes,
                            move, place, resize, set_props, ungroup_shape)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
os.makedirs(OUT, exist_ok=True)
TARGET = os.path.join(OUT, "shapes-demo.pptx")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = int(prs.slide_width), int(prs.slide_height)
painter = Painter(prs, Theme.preset("dark"))
build = painter.blank({"title": "shapes"})
slide = build.slide
# 铺一层背景
bg = add_element(slide, {"name": "bg", "kind": "rect", "x": 0, "y": 0, "w": "100%", "h": "100%",
                         "fill": "0B1020"}, slide_w=SW, slide_h=SH)
print("[bg]", bg["box"])

# ---- 元素定位：绝对 + 百分比 + 锚点
a = add_element(slide, {"name": "a", "kind": "rect", "x": 1.0, "y": 1.2, "w": 2.2, "h": 1.1,
                        "text": "A", "fill": "4C8DFF"}, slide_w=SW, slide_h=SH)
b = add_element(slide, {"name": "b", "kind": "round-rect", "x": "30%", "y": "20%", "w": 2.0, "h": 1.2,
                        "text": "B", "fill": "6FE3C4", "alpha": 0.7}, slide_w=SW, slide_h=SH)
c = add_element(slide, {"name": "c", "kind": "ellipse", "x": "50%", "y": "50%", "w": 1.6, "h": 1.6,
                        "anchor": "center", "text": "C", "fill": "FFC53D", "rot": 15,
                        "flip_h": True}, slide_w=SW, slide_h=SH)
for rec in (a, b, c):
    print("[elem]", rec["name"], rec["box"], rec["applied"])

# ---- 移动 / 缩放 / 属性
print("[move a]", move(find_shape(slide, "a"), dx=0.5, dy="60pt"))
print("[resize a]", resize(find_shape(slide, "a"), scale=0.8, anchor="center"))
print("[props b]", set_props(find_shape(slide, "b"), size=18, bold=True, line_color="FFFFFF",
                             line_width=1.5, z_order="front"))

# ---- 组合
shapes = [find_shape(slide, n) for n in ("a", "b", "c")]
info = group_shapes(slide, shapes, name="trio")
print("[group]", json.dumps(info, ensure_ascii=False))
grp = find_shape(slide, "group:trio")
print("[group box]", box_of(grp))
xml = grp._element.xml
print("[group xml ok]", "chOff" in xml, "chExt" in xml)

# ---- 组移动：只改 off
before = box_of(grp)
move(grp, dx=0.8, dy=0.3)
after = box_of(grp)
xoff = grp._element.find(".//" + "{http://schemas.openxmlformats.org/drawingml/2006/main}chOff").get("x")
print("[group move] before", before, "after", after, "chOff.x", xoff)

# ---- 组缩放：只改 ext
before = box_of(grp)
resize(grp, scale=0.75, anchor="center")
after = box_of(grp)
chext = grp._element.find(".//" + "{http://schemas.openxmlformats.org/drawingml/2006/main}chExt")
print("[group scale] before", before, "after", after,
      "chExt", [chext.get("cx"), chext.get("cy")])

# ---- 解组
res = ungroup_shape(slide, find_shape(slide, "group:trio"))
print("[ungroup]", json.dumps(res, ensure_ascii=False))
for shp in slide.shapes:
    print("   after ungroup:", shp.name, box_of(shp) if shp.name in ("a", "b", "c") else "")

# ---- 对齐 / 分布
lefts = [1.0, 4.0, 8.0]
for i, shp in enumerate([s for s in slide.shapes if s.name in ("a", "b", "c")]):
    place(shp, x=lefts[i], y=5.0, anchor="left-top", slide_w=SW, slide_h=SH)
print("[align]", align_shapes(slide, [find_shape(slide, n) for n in ("a", "b", "c")],
                              "distribute-h", to="slide", slide_w=SW, slide_h=SH))
print("[align top]", align_shapes(slide, [find_shape(slide, n) for n in ("a", "b", "c")],
                                  "top", to="slide", slide_w=SW, slide_h=SH))

prs.save(TARGET)
print("[save]", TARGET)
print("[validate]", validate_deck(TARGET)["stats"], validate_deck(TARGET)["problems"])

com = comtools.verify(TARGET)
if com.get("available"):
    page = com["slides"][0]
    print("[com] shapes:", page["slide_shapes"])
else:
    print("[com] unavailable")

# 用 COM 读几何做交叉验证
if comtools.available():
    from mcp_ppt.comtools import _app, _open
    app = _app()
    pres = _open(app, TARGET)
    s = pres.Slides(1)
    for i in range(1, s.Shapes.Count + 1):
        shp = s.Shapes(i)
        print(f"   COM {shp.Name}: type={shp.Type} "
              f"({round(float(shp.Left) / 72, 2)}in, {round(float(shp.Top) / 72, 2)}in, "
              f"{round(float(shp.Width) / 72, 2)}in x {round(float(shp.Height) / 72, 2)}in) "
              f"rot={shp.Rotation}")
    pres.Close()
    app.Quit()