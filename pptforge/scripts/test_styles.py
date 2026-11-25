"""属性层验收：渐变/图案填充、线条虚线箭头、阴影/发光/柔化、富文本、图片裁剪、表格。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.util import Inches

from mcp_ppt import comtools
from mcp_ppt.comtools import _app, _open
from mcp_ppt.design import Painter, Theme
from mcp_ppt.inspect import validate_deck
from mcp_ppt.layout import layout_report
from mcp_ppt.shapes import add_element, find_shape, set_props

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
os.makedirs(OUT, exist_ok=True)
TARGET = os.path.join(OUT, "styles-demo.pptx")
IMG = os.path.join(ROOT, "rev", "out", "test-image.png")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = int(prs.slide_width), int(prs.slide_height)
painter = Painter(prs, Theme.preset("dark"))
build = painter.blank({"title": "styles"})
slide = build.slide
add_element(slide, {"name": "bg", "kind": "rect", "x": 0, "y": 0, "w": "100%", "h": "100%",
                    "fill": "#0B1020"}, slide_w=SW, slide_h=SH)

# ---- 渐变（线性 + 径向）与图案
g1 = add_element(slide, {"name": "grad-linear", "kind": "round-rect", "x": 0.8, "y": 0.9,
                         "w": 3.4, "h": 1.6, "text": "linear", "fill": "#FFFFFF"},
                 slide_w=SW, slide_h=SH)
print("[fill grad-linear]", set_props(find_shape(slide, "grad-linear"), fill={
    "type": "gradient", "angle": 45,
    "stops": [{"pos": 0, "color": "#4C8DFF"}, {"pos": 100, "color": "#6FE3C4"}]}))
g2 = add_element(slide, {"name": "grad-radial", "kind": "ellipse", "x": 4.7, "y": 0.9,
                         "w": 3.2, "h": 1.6, "text": "radial"},
                 slide_w=SW, slide_h=SH)
print("[fill grad-radial]", set_props(find_shape(slide, "grad-radial"), fill={
    "type": "gradient", "style": "radial",
    "stops": [{"pos": 0, "color": "#FFD166", "alpha": 0.95},
              {"pos": 100, "color": "#1A1110", "alpha": 0.6}]}))

# ---- 线条：虚线 / 箭头
ln = add_element(slide, {"name": "wire", "kind": "line", "x": 0.9, "y": 3.1, "w": 4.0, "h": 0},
                 slide_w=SW, slide_h=SH)
print("[line dash]", set_props(find_shape(slide, "wire"),
                               line_color="#6FE3C4", line_width=3, line_dash="dash",
                               line_alpha=0.8))
arr = add_element(slide, {"name": "arrow", "kind": "line", "x": 5.6, "y": 3.1, "w": 3.6, "h": 0.9},
                  slide_w=SW, slide_h=SH)
print("[line arrow]", set_props(find_shape(slide, "arrow"),
                                line_color="#FFD166", line_width=2.5,
                                line_head="arrow", line_tail="triangle"))

# ---- 效果：阴影 / 发光 / 柔化
sh = add_element(slide, {"name": "card", "kind": "round-rect", "x": 0.9, "y": 4.1,
                         "w": 3.0, "h": 1.7, "text": "shadow", "fill": "#18233F"},
                 slide_w=SW, slide_h=SH)
print("[shadow]", set_props(find_shape(slide, "card"), effects={
    "shadow": {"blur": 14, "dist": 10, "dir": 45, "color": "#000000", "alpha": 0.55}}))
gl = add_element(slide, {"name": "glow", "kind": "round-rect", "x": 4.3, "y": 4.1,
                         "w": 3.0, "h": 1.7, "text": "glow", "fill": "#0D2038",
                         "line_color": "#36C2FF", "line_width": 1.5},
                 slide_w=SW, slide_h=SH)
print("[glow]", set_props(find_shape(slide, "glow"), effects={
    "glow": {"radius": 14, "color": "#36C2FF", "alpha": 0.7}}))
se = add_element(slide, {"name": "soft", "kind": "round-rect", "x": 7.7, "y": 4.1,
                         "w": 3.0, "h": 1.7, "text": "soft-edge", "fill": "#225533"},
                 slide_w=SW, slide_h=SH)
print("[soft]", set_props(find_shape(slide, "soft"), effects={"soft_edge": 12}))

# ---- 富文本：多段 + 段内混排
tb = add_element(slide, {"name": "rich", "kind": "text", "x": 0.9, "y": 6.1, "w": 9.6, "h": 1.1},
                 slide_w=SW, slide_h=SH)
print("[rich]", set_props(find_shape(slide, "rich"), paragraphs=[
    {"align": "left", "space_after": 6, "line_spacing": 1.2,
     "runs": [{"text": "渐变 ", "bold": True, "size": 26, "color": "#6FE3C4"},
              {"text": "与 ", "size": 26, "color": "#EAF3FF"},
              {"text": "阴影", "bold": True, "size": 26, "color": "#FFD166",
               "underline": True}]},
    {"align": "left", "level": 1, "bullet": {"char": "▪", "color": "#36C2FF"},
     "runs": [{"text": "混排 run：粗体 / 斜体 / 小字号", "italic": True, "size": 16,
               "color": "#9BABCE"}]},
]))

# ---- 图片裁剪与亮度
if os.path.exists(IMG):
    pic = add_element(slide, {"name": "photo", "kind": "image", "image": IMG,
                              "x": 10.6, "y": 0.9, "w": 2.2, "h": 3.0},
                      slide_w=SW, slide_h=SH)
    print("[crop]", set_props(find_shape(slide, "photo"), crop={"l": 0.1, "t": 0.08, "r": 0.05},
                               brightness=0.1, contrast=0.15))

prs.save(TARGET)
print("[save]", TARGET)
print("[validate]", validate_deck(TARGET)["stats"], validate_deck(TARGET)["problems"])
lay = layout_report(TARGET)
print("[layout] issue_count:", lay.get("issue_count"),
      [p["issues"] for p in lay.get("pages", []) if p["issues"]])

# ---- XML 抽查
import zipfile
import re
xml = zipfile.ZipFile(TARGET).read("ppt/slides/slide1.xml").decode("utf-8")
for pat in (r"<a:gradFill.*?</a:gradFill>", r"<a:effectLst>.*?</a:effectLst>",
            r"<a:ln[ >].*?</a:ln>", r"<a:srcRect[^>]*/>", r"<a:lum[^>]*/>"):
    for m in re.finditer(pat, xml, re.S):
        seg = re.sub(r"\s+", " ", m.group(0))
        print("   XML:", seg[:260])

if comtools.available():
    app = _app()
    pres = _open(app, TARGET)
    s = pres.Slides(1)
    print("[com]")
    for i in range(1, s.Shapes.Count + 1):
        shp = s.Shapes(i)
        try:
            fill_type = shp.Fill.Type
        except Exception:  # noqa: BLE001
            fill_type = "?"
        try:
            shadow_vis = shp.Shadow.Visible
        except Exception:  # noqa: BLE001
            shadow_vis = "?"
        crop = ""
        if shp.Type == 13:
            pf = shp.PictureFormat
            crop = f" crop=({pf.CropLeft},{pf.CropTop},{pf.CropRight})"
        print(f"   {shp.Name}: fill.Type={fill_type} shadow={shadow_vis}{crop}")
    pres.Close()
    app.Quit()