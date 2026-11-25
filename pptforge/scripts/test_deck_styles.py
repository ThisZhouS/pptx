"""通过 deck 模型（elements + ops）写样式：渐变/效果/富文本/裁剪全部走 DSL。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import validate_deck
from mcp_ppt.layout import layout_report

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
TARGET = os.path.join(OUT, "styles-deck.pptx")
IMG = os.path.join(ROOT, "rev", "out", "test-image.png").replace("\\", "/")

MODEL = {
    "theme": "warm",
    "footer": "styles via DSL",
    "slides": [
        {
            "layout": "blank",
            "background": "#1A1110",
            "elements": [
                {"name": "hero", "kind": "round-rect", "x": 0.9, "y": 1.1, "w": 5.4, "h": 2.4,
                 "fill": {"type": "gradient", "angle": 30,
                          "stops": [{"pos": 0, "color": "#FF7A45"},
                                    {"pos": 100, "color": "#FFC53D"}]},
                 "effects": {"shadow": {"blur": 18, "dist": 12, "dir": 45, "alpha": 0.5}},
                 "paragraphs": [
                     {"align": "left", "runs": [
                         {"text": "渐变 + 外阴影", "bold": True, "size": 32, "color": "#1A1110"}]},
                     {"align": "left", "space_before": 8, "runs": [
                         {"text": "全部由 deck 模型的 elements 写出", "size": 14,
                          "color": "#3B1E12"}]},
                 ]},
                {"name": "ring", "kind": "ellipse", "x": 6.9, "y": 1.1, "w": 2.4, "h": 2.4,
                 "fill": {"type": "gradient", "style": "radial",
                          "stops": [{"pos": 0, "color": "#FFC53D", "alpha": 0.95},
                                    {"pos": 100, "color": "#1A1110", "alpha": 0.5}]},
                 "effects": {"glow": {"radius": 16, "color": "#FF7A45", "alpha": 0.6}}},
                {"name": "tag", "kind": "round-rect", "x": 9.7, "y": 1.1, "w": 2.4, "h": 0.7,
                 "text": "SOFT EDGE", "size": 12, "bold": True, "color": "#1A1110",
                 "align": "center", "fill": "#FFC53D",
                 "effects": {"soft_edge": 6}},
                {"name": "rule", "kind": "line", "x": 0.9, "y": 4.0, "w": 11.2, "h": 0,
                 "line_color": "#FF7A45", "line_width": 2, "line_dash": "dash",
                 "line_alpha": 0.7},
                {"name": "arrow", "kind": "line", "x": 0.9, "y": 4.6, "w": 6.0, "h": 0.8,
                 "line_color": "#FFC53D", "line_width": 2.5,
                 "line_head": "triangle", "line_tail": "arrow"},
                {"name": "notes", "kind": "text", "x": 0.9, "y": 5.4, "w": 7.6, "h": 1.4,
                 "paragraphs": [
                     {"align": "left", "bullet": {"char": "▸", "color": "#FF7A45"},
                      "space_after": 6,
                      "runs": [{"text": "fill 支持 solid / gradient / none",
                                "size": 16, "color": "#FFF3EC"}]},
                     {"align": "left", "level": 1, "bullet": {"char": "·", "color": "#FFC53D"},
                      "runs": [{"text": "effects 支持 shadow / glow / soft_edge", "size": 14,
                                "color": "#D8B4A6"}]},
                 ]},
                {"name": "photo", "kind": "image", "image": IMG, "x": 9.0, "y": 3.6,
                 "w": 3.1, "h": 2.4, "crop": {"l": 0.08, "t": 0.05, "b": 0.05},
                 "brightness": 0.08, "contrast": 0.12},
            ],
            "ops": [
                {"op": "props", "target": "tag", "rot": -4},
                {"op": "group", "targets": ["ring", "tag"], "name": "badge"},
                {"op": "resize", "target": "group:badge", "scale": 0.95, "anchor": "center"},
            ],
            "transition": {"effect": "glitter-diamond-left", "duration": 0.9},
            "animations": [
                {"target": "hero", "effect": "fade", "trigger": "click", "duration": 0.6},
                {"target": "group:badge", "effect": "zoom", "trigger": "after_previous"},
                {"target": "notes", "effect": "wipe", "trigger": "after_previous",
                 "build": "paragraph"},
            ],
        }
    ],
}

report = build_deck(MODEL, TARGET)
print("[build]", report["slide_count"], "warnings:", report["warnings"])
page = report["slides"][0]
for rec in page.get("elements", []):
    print("  element", rec["name"], json.dumps(rec["applied"], ensure_ascii=False)[:150])
print("[groups]", json.dumps(page.get("groups"), ensure_ascii=False))
for rec in page.get("ops", []):
    print("  op", rec["op"], "ok" if rec["ok"] else rec["error"])

val = validate_deck(TARGET)
print("[validate]", val["stats"], val["problems"])
lay = layout_report(TARGET)
print("[layout] issue_count", lay.get("issue_count"),
      [p["issues"] for p in lay.get("pages", []) if p["issues"]])
if comtools.available():
    com = comtools.verify(TARGET)
    print("[com] shapes:", com["slides"][0]["slide_shapes"],
          "anims:", [(a["type"], a["trigger"], a["shape_name"])
                     for a in com["slides"][0]["animations"]])