"""形状层 MCP 工具联调：add_elements → edit_elements → element_info → 动画挂到组。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import server as S
from mcp_ppt import comtools

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
TARGET = os.path.join(OUT, "shape-tools.pptx")


def show(title, obj, limit=1400):
    print(f"\n=== {title} ===")
    text = json.dumps(obj, ensure_ascii=False, indent=1)
    print(text if len(text) <= limit else text[:limit] + " ...")


# 1) 建一份最小 deck
build = S.build_deck({
    "theme": "dark", "footer": "shape tools",
    "slides": [{"layout": "blank", "background": "0B1020", "transition": {"effect": "fade"}}],
}, TARGET)
show("build", {"slide_count": build["slide_count"], "warnings": build["warnings"]})

# 2) 加元素
added = S.add_elements(TARGET, 1, [
    {"name": "n1", "kind": "ellipse", "x": 1.2, "y": 2.4, "w": 1.3, "h": 1.3, "text": "A",
     "fill": "4C8DFF", "align": "center", "size": 20, "bold": True},
    {"name": "n2", "kind": "ellipse", "x": 3.6, "y": 2.4, "w": 1.3, "h": 1.3, "text": "B",
     "fill": "4C8DFF", "align": "center", "size": 20, "bold": True},
    {"name": "n3", "kind": "ellipse", "x": 6.0, "y": 2.4, "w": 1.3, "h": 1.3, "text": "C",
     "fill": "4C8DFF", "align": "center", "size": 20, "bold": True},
    {"name": "rule", "kind": "line", "x": 2.5, "y": 3.05, "w": 1.1, "h": 0,
     "line_color": "6FE3C4", "line_width": 2},
    {"name": "note", "kind": "text", "x": 1.2, "y": 5.2, "w": 8.0, "h": 0.8,
     "text": "先建元素，再用 ops 组合与调整", "size": 16, "color": "9BABCE"},
], groups={"nodes": ["n1", "n2"]})
show("add_elements", {"patched": added["patched"][0]["applied"].get("elements"),
                      "groups": added["patched"][0]["applied"].get("groups"),
                      "warnings": added["warnings"]})

# 3) ops：定位/移动/缩放/属性/对齐/复制/组合
edited = S.edit_elements(TARGET, 1, [
    {"op": "place", "target": "note", "x": "10%", "y": 5.0},
    {"op": "props", "target": "n3", "fill": "FFD166", "rot": 10, "alpha": 0.9},
    {"op": "resize", "target": "group:nodes", "scale": 1.15, "anchor": "center"},
    {"op": "duplicate", "target": "n3", "name": "n4", "dx": 2.4},
    {"op": "align", "targets": ["n4"], "mode": "top", "to": "n3"},
    {"op": "align", "targets": ["n1", "n2", "n3", "n4"], "mode": "distribute-h", "to": "slide"},
    {"op": "group", "targets": ["n3", "rule"], "name": "pair"},
])
for rec in edited["patched"][0]["applied"].get("ops", []):
    print("  op", rec["op"], "ok" if rec["ok"] else f"FAIL {rec.get('error')}",
          json.dumps(rec.get("result"), ensure_ascii=False)[:120])

# 4) 动画挂到组
anim = S.add_animations(TARGET, 1, [
    {"target": "group:nodes", "effect": "fade", "trigger": "click", "duration": 0.6},
    {"target": "group:pair", "effect": "zoom", "trigger": "after_previous", "duration": 0.5},
    {"target": "note", "effect": "wipe", "trigger": "after_previous", "build": "paragraph"},
])
show("add_animations", {"warnings": anim["warnings"],
                        "animations": anim["patched"][0]["applied"].get("animations")})

# 5) 读回
info = S.element_info(TARGET, 1)
show("element_info", {"count": info["count"],
                      "names": [(s["name"], s["type"], [s["left_in"], s["top_in"]],
                                 s.get("rot"), bool(s.get("group_items")))
                                for s in info["shapes"]]})
one = S.element_info(TARGET, 1, "group:nodes")
show("element_info(group)", one["shapes"][0])

show("validate", S.validate(TARGET)["stats"])
print("[com]", [(a["type"], a["trigger"], a["shape_name"]) for a in
                (comtools.verify(TARGET)["slides"][0]["animations"] if comtools.available() else [])])