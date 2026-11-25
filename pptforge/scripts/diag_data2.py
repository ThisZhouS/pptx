"""更细二分：区分"构建器"与"版式"哪一层写出非法 XML。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import win32com.client as win32
from pptx import Presentation
from pptx.util import Inches

from mcp_ppt.chart import build_chart
from mcp_ppt.design import Painter, Theme
from mcp_ppt.table import build_table

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "diag")
os.makedirs(OUT, exist_ok=True)


def new_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs


def save_case(name, fn):
    prs = new_prs()
    painter = Painter(prs, Theme.preset("dark"))
    kicker = painter.blank({})
    fn(prs, kicker.slide, int(prs.slide_width), int(prs.slide_height))
    path = os.path.join(OUT, f"{name}.pptx")
    prs.save(path)
    return path


def case_table_builder(prs, slide, sw, sh):
    build_table(slide.shapes, {"name": "t1", "x": 1.0, "y": 1.0, "w": 8.0, "h": 3.0,
                               "rows": [["A", "B"], ["1", "2"]], "header": True},
                None, sw, sh)


def case_table_builder_plain(prs, slide, sw, sh):
    # 最简：只建表，不设任何样式
    slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(8), Inches(3))


def case_chart_builder(prs, slide, sw, sh):
    build_chart(slide.shapes, {"name": "c1", "type": "column", "categories": ["A", "B"],
                               "series": [{"name": "S", "values": [1, 2]}],
                               "x": 1.0, "y": 1.0, "w": 8.0, "h": 4.0}, None, sw, sh)


def case_chart_plain(prs, slide, sw, sh):
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    data = CategoryChartData()
    data.categories = ["A", "B"]
    data.add_series("S", (1, 2))
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(1),
                           Inches(8), Inches(4), data)


def case_layout_table(prs, slide, sw, sh):
    painter = Painter(prs, Theme.preset("dark"))
    # 用同一个 slide 不行，直接建新页
    painter.table({"title": "T", "rows": [["A", "B"], ["1", "2"]]})


def case_layout_chart(prs, slide, sw, sh):
    painter = Painter(prs, Theme.preset("dark"))
    painter.chart({"title": "C", "chart": {"type": "column", "categories": ["A", "B"],
                                           "series": [{"name": "S", "values": [1, 2]}]}})


CASES = [
    ("tbl_plain_pptx", case_table_builder_plain),
    ("tbl_builder", case_table_builder),
    ("chart_plain_pptx", case_chart_plain),
    ("chart_builder", case_chart_builder),
    ("layout_table", case_layout_table),
    ("layout_chart", case_layout_chart),
]

app = win32.Dispatch("PowerPoint.Application")
app.DisplayAlerts = 1
for name, fn in CASES:
    path = save_case(name, fn)
    try:
        pres = app.Presentations.Open(os.path.abspath(path), ReadOnly=-1, WithWindow=0)
        n = pres.Slides.Count
        pres.Close()
        print(f"  OK    {name} (slides={n})")
    except Exception as exc:  # noqa: BLE001
        print(f"  FAIL  {name}: {exc}")
app.Quit()