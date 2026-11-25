"""布局层验收：自定义元素 + 组合 + ops（定位/移动/缩放/属性/对齐/复制/删除）+ 组动画。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.comtools import _app, _open
from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import inspect_deck, validate_deck

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
os.makedirs(OUT, exist_ok=True)
TARGET = os.path.join(OUT, "layout-demo.pptx")

MODEL = {
    "theme": "midnight",
    "footer": "layout demo",
    "slides": [
        {
            "layout": "title",
            "title": "元素定位与组合",
            "subtitle": "坐标 / 属性 / 组合 / 移动 / 缩放",
            "meta": ["ops"],
            "transition": {"effect": "pan-left", "duration": 0.9},
            "animations": [
                {"target": "title", "effect": "fade", "trigger": "click"},
                {"target": "subtitle", "effect": "fade", "trigger": "after_previous"},
            ],
        },
        {
            "layout": "blank",
            "background": "071426",
            "kicker": "diagram",
            "elements": [
                {"name": "tag", "kind": "round-rect", "x": "72%", "y": "10%", "w": 1.5, "h": 0.46,
                 "text": "NEW", "fill": "FFD166", "color": "101828", "size": 12, "bold": True,
                 "align": "center"},
                {"name": "headline", "kind": "text", "x": 1.0, "y": 0.9, "w": 7.0, "h": 0.9,
                 "text": "从形状到组合", "size": 34, "bold": True, "color": "EAF3FF"},
                {"name": "nodeA", "kind": "ellipse", "x": 1.4, "y": 3.4, "w": 1.3, "h": 1.3,
                 "text": "A", "fill": "36C2FF", "color": "04121F", "size": 20, "bold": True,
                 "align": "center"},
                {"name": "nodeB", "kind": "ellipse", "x": 3.9, "y": 3.4, "w": 1.3, "h": 1.3,
                 "text": "B", "fill": "36C2FF", "color": "04121F", "size": 20, "bold": True,
                 "align": "center"},
                {"name": "nodeC", "kind": "ellipse", "x": 6.4, "y": 3.4, "w": 1.3, "h": 1.3,
                 "text": "C", "fill": "36C2FF", "color": "04121F", "size": 20, "bold": True,
                 "align": "center"},
                {"name": "link", "kind": "line", "x": 2.7, "y": 4.05, "w": 1.2, "h": 0,
                 "line_color": "36C2FF", "line_width": 2},
                {"name": "caption", "kind": "text", "x": 1.4, "y": 5.2, "w": 7.0, "h": 0.8,
                 "text": "组合后整体移动与缩放，子元素坐标不变", "size": 14, "color": "8FB0D6"},
            ],
            "ops": [
                {"op": "move", "target": "headline", "dy": "40pt"},
                {"op": "props", "target": "tag", "rot": 6, "alpha": 0.85},
                {"op": "resize", "target": "tag", "scale": 1.15, "anchor": "right-top"},
                {"op": "group", "targets": ["nodeA", "nodeB", "link"], "name": "row"},
                {"op": "move", "target": "group:row", "dx": 0.4, "dy": 0.2},
                {"op": "resize", "target": "group:row", "scale": 0.9, "anchor": "center"},
                {"op": "duplicate", "target": "nodeC", "name": "nodeD", "dx": 2.5},
                {"op": "align", "targets": ["nodeD"], "mode": "top", "to": "nodeC"},
                {"op": "delete", "target": "caption"},
            ],
            "transition": {"effect": "gallery-left", "duration": 0.8},
            "animations": [
                {"target": "headline", "effect": "fade", "trigger": "click"},
                {"target": "group:row", "effect": "zoom", "trigger": "after_previous", "duration": 0.6},
                {"target": "nodeC", "effect": "spin", "trigger": "after_previous", "duration": 0.8},
            ],
        },
    ],
}

report = build_deck(MODEL, TARGET)
print("[build] slides:", report["slide_count"], "| warnings:", report["warnings"])
page2 = report["slides"][1]
print("[elements]", json.dumps(page2.get("elements"), ensure_ascii=False, indent=1)[:1200])
print("[groups]", json.dumps(page2.get("groups"), ensure_ascii=False))
print("[ops]")
for rec in page2.get("ops") or []:
    print("   ", rec["op"], "ok" if rec["ok"] else f"FAIL {rec.get('error')}",
          json.dumps(rec.get("result"), ensure_ascii=False)[:150])

val = validate_deck(TARGET)
print("[validate]", val["stats"], val["problems"])
deck = inspect_deck(TARGET)
print("[inspect#2 shapes]")
for s in deck["slides"][1]["shapes"]:
    print("   ", s["id"], s["name"], s["tag"])
print("[inspect#2 animations]", deck["slides"][1]["animations"])

com = comtools.verify(TARGET)
print("[com]", "slides:", com.get("slide_count"),
      [(a["type"], a["trigger"], a["shape_name"]) for a in com["slides"][1]["animations"]])

if comtools.available():
    app = _app()
    pres = _open(app, TARGET)
    s = pres.Slides(2)
    print("[com geometry]")
    for i in range(1, s.Shapes.Count + 1):
        shp = s.Shapes(i)
        extra = ""
        try:
            if shp.Type == 6:  # msoGroup
                extra = f" groupItems={shp.GroupItems.Count}"
        except Exception:  # noqa: BLE001
            pass
        print(f"   {shp.Name}: ({round(float(shp.Left)/72,2)}, {round(float(shp.Top)/72,2)}) "
              f"{round(float(shp.Width)/72,2)}x{round(float(shp.Height)/72,2)}in "
              f"rot={round(float(shp.Rotation),1)}{extra}")
    pres.Close()
    app.Quit()