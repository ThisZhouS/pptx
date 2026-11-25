"""二分定位：表格页 / 图表页 各自单独生成，看 PowerPoint 能否打开。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32

from mcp_ppt.deck import build_deck

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "diag")
os.makedirs(OUT, exist_ok=True)

CHART = {
    "layout": "chart", "title": "图表页",
    "chart": {"type": "column", "categories": ["Q1", "Q2"],
              "series": [{"name": "A", "values": [1, 2]}], "legend": "bottom"},
}
TABLE = {
    "layout": "table", "title": "表格页",
    "rows": [["A", "B"], ["1", "2"]],
}
TABLE_OPS = {
    "layout": "table", "title": "表格 + ops",
    "rows": [["A", "B"], ["1", "2"]],
    "ops": [{"op": "table", "target": "table",
             "cells": [{"row": 1, "col": 1, "text": "x"}]}],
}

CASES = {"chart_only": [CHART], "table_only": [TABLE], "table_ops": [TABLE_OPS]}
app = win32.Dispatch("PowerPoint.Application")
app.DisplayAlerts = 1
for name, slides in CASES.items():
    path = os.path.join(OUT, f"{name}.pptx")
    try:
        build_deck({"theme": "dark", "slides": slides}, path)
    except Exception as exc:  # noqa: BLE001
        print(f"  BUILD-FAIL {name}: {type(exc).__name__}: {exc}")
        continue
    try:
        pres = app.Presentations.Open(os.path.abspath(path), ReadOnly=-1, WithWindow=0)
        n = pres.Slides.Count
        shapes = pres.Slides(1).Shapes.Count
        pres.Close()
        print(f"  OK    {name} (slides={n}, shapes={shapes})")
    except Exception as exc:  # noqa: BLE001
        print(f"  FAIL  {name}: {exc}")
app.Quit()