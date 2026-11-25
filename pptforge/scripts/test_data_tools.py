"""数据层 MCP 工具联调：add_table → edit_table → add_chart → element_info 摘要。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt import server as S

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TARGET = os.path.join(ROOT, "samples", "data-tools.pptx")


def show(title, obj, limit=900):
    print(f"\n=== {title} ===")
    text = json.dumps(obj, ensure_ascii=False, indent=1)
    print(text if len(text) <= limit else text[:limit] + " ...")


build = S.build_deck({"theme": "dark", "slides": [
    {"layout": "blank", "background": "0B1020", "transition": {"effect": "fade"}}]}, TARGET)
show("build", {"slides": build["slide_count"], "warnings": build["warnings"]})

tbl = S.add_table(TARGET, 1, {
    "name": "grid", "x": 0.9, "y": 1.2, "w": 11.4, "h": 2.6,
    "rows": [["渠道", "Q3", "Q4", "变化"],
             ["直营", "18.0", "22.0", "+22%"],
             ["分销", "9.5", "12.5", "+32%"]],
    "header": True, "col_widths": [3.2, 2.4, 2.4, 2.4], "font_size": 15,
    "cell": {"valign": "middle"},
    "header_cell": {"align": "center"},
})
show("add_table", {"elements": tbl["patched"][0]["applied"].get("elements"),
                   "warnings": tbl["warnings"]})

cols = S.edit_table(TARGET, 1, "grid", [
    {"op": "cells", "cells": [
        {"row": 1, "col": 3, "text": "+22%", "color": "#6FE3C4"},
        {"row": 2, "col": 3, "text": "+32%", "color": "#6FE3C4"},
        {"row": 1, "col": 0, "bold": True},
    ]},
    {"op": "merge", "row1": 3, "col1": 0, "row2": 3, "col2": 3},
    {"op": "cells", "cells": [{"row": 3, "col": 0, "text": "（季度合计）", "align": "right"}]},
    {"op": "style", "style": {"first_row": True, "band_row": True}},
])
show("edit_table", cols["results"], 1200)

chart = S.add_chart(TARGET, 1, {
    "name": "bars", "type": "bar", "x": 0.9, "y": 4.2, "w": 11.4, "h": 2.9,
    "categories": ["直营", "分销", "订阅"],
    "series": [{"name": "Q4", "values": [22.0, 12.5, 8.1], "color": "#4C8DFF"}],
    "legend": "none", "data_labels": True, "number_format": "0.0",
})
show("add_chart", {"elements": chart["patched"][0]["applied"].get("elements"),
                   "warnings": chart["warnings"]})

info = S.element_info(TARGET, 1, "grid")
show("element_info(table)", info["shapes"][0], 1500)
info2 = S.element_info(TARGET, 1, "bars")
show("element_info(chart)", {k: v for k, v in info2["shapes"][0].items() if k != "chart"},
     400)
show("chart data", info2["shapes"][0].get("chart"), 600)

anim = S.add_animations(TARGET, 1, [
    {"target": "grid", "effect": "fade", "trigger": "click"},
    {"target": "bars", "effect": "wipe", "trigger": "after_previous", "duration": 0.7},
])
show("add_animations", anim["patched"][0]["applied"].get("animations"))
show("validate", S.validate(TARGET)["stats"])
show("layout", {"issue_count": S.layout_check(TARGET).get("issue_count")})
if comtools.available():
    print("[com]", [(a["type"], a["trigger"], a["shape_name"])
                    for a in comtools.verify(TARGET)["slides"][0]["animations"]])