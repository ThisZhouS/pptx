"""定位哪个属性写入导致 PowerPoint 打不开：逐项生成并试探打开。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32
from pptx import Presentation
from pptx.util import Inches

from mcp_ppt.design import Painter, Theme
from mcp_ppt.shapes import add_element, find_shape, set_props

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "diag")
os.makedirs(OUT, exist_ok=True)
IMG = os.path.join(ROOT, "rev", "out", "test-image.png")

CASES = {}


def build(name, painter_fn):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    SW, SH = int(prs.slide_width), int(prs.slide_height)
    painter = Painter(prs, Theme.preset("dark"))
    slide = painter.blank({}).slide
    add_element(slide, {"name": "bg", "kind": "rect", "x": 0, "y": 0, "w": "100%", "h": "100%",
                        "fill": "#0B1020"}, slide_w=SW, slide_h=SH)
    painter_fn(slide, SW, SH)
    path = os.path.join(OUT, f"{name}.pptx")
    prs.save(path)
    return path


def case_gradient(slide, SW, SH):
    s = add_element(slide, {"name": "g", "kind": "round-rect", "x": 1, "y": 1, "w": 4, "h": 2,
                            "text": "grad"}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "g"), fill={"type": "gradient", "angle": 45,
                                            "stops": [{"pos": 0, "color": "#4C8DFF"},
                                                      {"pos": 100, "color": "#6FE3C4"}]})


def case_gradient_radial(slide, SW, SH):
    s = add_element(slide, {"name": "g", "kind": "ellipse", "x": 1, "y": 1, "w": 4, "h": 2},
                    slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "g"), fill={"type": "gradient", "style": "radial",
                                            "stops": [{"pos": 0, "color": "#FFD166", "alpha": 0.95},
                                                      {"pos": 100, "color": "#1A1110", "alpha": 0.6}]})


def case_line(slide, SW, SH):
    add_element(slide, {"name": "l", "kind": "line", "x": 1, "y": 2, "w": 5, "h": 0},
                slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "l"), line_color="#6FE3C4", line_width=3,
              line_dash="dash", line_alpha=0.8)


def case_line_arrow(slide, SW, SH):
    add_element(slide, {"name": "l", "kind": "line", "x": 1, "y": 2, "w": 5, "h": 1},
                slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "l"), line_color="#FFD166", line_width=2.5,
              line_head="arrow", line_tail="triangle")


def case_shadow(slide, SW, SH):
    add_element(slide, {"name": "c", "kind": "round-rect", "x": 1, "y": 1, "w": 3, "h": 2,
                        "fill": "#18233F"}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "c"), effects={"shadow": {"blur": 14, "dist": 10, "dir": 45,
                                                         "color": "#000000", "alpha": 0.55}})


def case_glow(slide, SW, SH):
    add_element(slide, {"name": "c", "kind": "round-rect", "x": 1, "y": 1, "w": 3, "h": 2,
                        "fill": "#0D2038"}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "c"), effects={"glow": {"radius": 14, "color": "#36C2FF",
                                                       "alpha": 0.7}})


def case_softedge(slide, SW, SH):
    add_element(slide, {"name": "c", "kind": "round-rect", "x": 1, "y": 1, "w": 3, "h": 2,
                        "fill": "#225533"}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "c"), effects={"soft_edge": 12})


def case_text(slide, SW, SH):
    add_element(slide, {"name": "t", "kind": "text", "x": 1, "y": 1, "w": 8, "h": 2},
                slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "t"), paragraphs=[
        {"align": "left", "runs": [{"text": "渐变 ", "bold": True, "size": 26,
                                    "color": "#6FE3C4"},
                                   {"text": "普通", "size": 26, "color": "#EAF3FF"}]}])


def case_text_bullet(slide, SW, SH):
    add_element(slide, {"name": "t", "kind": "text", "x": 1, "y": 1, "w": 8, "h": 2},
                slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "t"), paragraphs=[
        {"align": "left", "level": 1, "bullet": {"char": "▪", "color": "#36C2FF"},
         "space_before": 12, "space_after": 8, "line_spacing": 1.4,
         "runs": [{"text": "要点", "size": 18}]}])


def case_picture(slide, SW, SH):
    add_element(slide, {"name": "p", "kind": "image", "image": IMG, "x": 1, "y": 1,
                        "w": 4, "h": 3}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "p"), crop={"l": 0.1, "t": 0.08, "r": 0.05})


def case_picture_lum(slide, SW, SH):
    add_element(slide, {"name": "p", "kind": "image", "image": IMG, "x": 1, "y": 1,
                        "w": 4, "h": 3}, slide_w=SW, slide_h=SH)
    set_props(find_shape(slide, "p"), brightness=0.1, contrast=0.15)


CASES = {
    "gradient": case_gradient,
    "gradient_radial": case_gradient_radial,
    "line": case_line,
    "line_arrow": case_line_arrow,
    "shadow": case_shadow,
    "glow": case_glow,
    "softedge": case_softedge,
    "text": case_text,
    "text_bullet": case_text_bullet,
    "picture": case_picture,
    "picture_lum": case_picture_lum,
}

app = win32.Dispatch("PowerPoint.Application")
app.DisplayAlerts = 1
for name, fn in CASES.items():
    path = build(name, fn)
    try:
        pres = app.Presentations.Open(os.path.abspath(path), ReadOnly=-1, WithWindow=0)
        count = pres.Slides.Count
        pres.Close()
        print(f"  OK    {name} (slides={count})")
    except Exception as exc:  # noqa: BLE001
        print(f"  FAIL  {name}: {exc}")
try:
    app.Quit()
except Exception:  # noqa: BLE001
    pass