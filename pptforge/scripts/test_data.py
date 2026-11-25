"""数据层验收：表格（单元格/边框/合并/表样式）+ 图表（柱状/折线/饼图）+ 图表整体动画。"""
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
TARGET = os.path.join(OUT, "data-demo.pptx")

MODEL = {
    "theme": "dark",
    "footer": "data layer",
    "slides": [
        {
            "layout": "chart",
            "kicker": "数据",
            "title": "季度收支",
            "chart": {
                "type": "column",
                "categories": ["Q1", "Q2", "Q3", "Q4"],
                "series": [
                    {"name": "收入", "values": [12.0, 18.0, 15.0, 22.0], "color": "#4C8DFF"},
                    {"name": "成本", "values": [8.0, 9.5, 11.0, 12.5], "color": "#6FE3C4"},
                ],
                "legend": "bottom",
                "data_labels": True,
                "number_format": "0.0",
                "gridlines": True,
                "gap_width": 60,
            },
            "transition": {"effect": "cover-left", "duration": 0.6},
            "animations": [
                {"target": "title", "effect": "fade", "trigger": "click"},
                {"target": "chart", "effect": "wipe", "trigger": "after_previous", "duration": 0.8},
            ],
        },
        {
            "layout": "chart",
            "kicker": "结构",
            "title": "收入构成",
            "chart": {
                "type": "doughnut",
                "categories": ["产品", "服务", "订阅"],
                "series": [{"name": "占比", "values": [45.0, 30.0, 25.0],
                            "colors": ["#4C8DFF", "#6FE3C4", "#FFD166"]}],
                "legend": "right",
                "data_labels": True,
                "number_format": "0\"%\"",
            },
            "animations": [{"target": "chart", "effect": "zoom", "trigger": "click"}],
        },
        {
            "layout": "table",
            "kicker": "明细",
            "title": "分渠道表现",
            "rows": [
                ["渠道", "Q3", "Q4", "变化"],
                ["直营", "18.0", "22.0", "+22%"],
                ["分销", "9.5", "12.5", "+32%"],
                ["订阅", "6.2", "8.1", "+31%"],
            ],
            "col_widths": [2.6, 2.2, 2.2, 2.2],
            "font_size": 15,
            "chart_title": False,
            "transition": {"effect": "fade", "duration": 0.6},
            "ops": [
                {"op": "table", "target": "table", "cells": [
                    {"row": 0, "col": 0, "text": "渠道", "align": "left"},
                    {"row": 0, "col": 3, "text": "同比", "align": "right"},
                    {"row": 1, "col": 3, "text": "+22%", "color": "#6FE3C4"},
                    {"row": 2, "col": 3, "text": "+32%", "color": "#6FE3C4"},
                    {"row": 3, "col": 3, "text": "+31%", "color": "#6FE3C4"},
                ]},
                {"op": "table", "target": "table", "style": {"first_row": True, "band_row": True}},
            ],
            "animations": [
                {"target": "title", "effect": "fade", "trigger": "click"},
                {"target": "table", "effect": "fade", "trigger": "after_previous"},
            ],
        },
    ],
}

report = build_deck(MODEL, TARGET)
print("[build]", report["slide_count"], "warnings:", report["warnings"])
for s in report["slides"]:
    names = [k for k in s["shapes"]]
    print(f"  slide {s['index']} ({s['layout']}): shapes={names}")
    for rec in s.get("ops", []):
        print("    op", rec["op"], "ok" if rec["ok"] else rec["error"],
              json.dumps(rec.get("result"), ensure_ascii=False)[:120])

val = validate_deck(TARGET)
print("[validate]", val["stats"], val["problems"])
lay = layout_report(TARGET)
print("[layout] issue_count", lay.get("issue_count"),
      [p["issues"] for p in lay.get("pages", []) if p["issues"]])
if comtools.available():
    com = comtools.verify(TARGET)
    for page in com["slides"]:
        print(f"  COM slide {page['index']}: shapes={page['slide_shapes']} "
              f"anims={[(a['type'], a['trigger'], a['shape_name']) for a in page['animations']]}")

# XML 抽查：表格 tcPr 边框 / 图表部件
import re
import zipfile
z = zipfile.ZipFile(TARGET)
parts = [n for n in z.namelist() if "chart" in n or "embed" in n]
print("[parts]", parts[:8])
xml = z.read("ppt/slides/slide3.xml").decode("utf-8")
for pat in (r"<a:tblPr[^>]*>", r"<a:tcPr>.{0,160}", r'<a:gridCol w="\d+"'):
    m = re.search(pat, xml, re.S)
    if m:
        print("   ", re.sub(r"\s+", " ", m.group(0))[:200])
chart_xml = z.read(parts[0]).decode("utf-8") if parts else ""
print("[chart xml head]", re.sub(r"\s+", " ", chart_xml)[:300])