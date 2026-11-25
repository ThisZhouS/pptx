"""逆向第 6 层：表格（单元格级）与图表（含图表动画）。

表格：COM 建表 → 合并/列宽行高/单元格填充/边框/字体 → 解包看 a:tbl 结构。
图表：python-pptx 生成 chart → PowerPoint 打开加"按系列/按类别"动画 → 解包看 p:timing 里的
      bldGraphic / graphicEl / 图表分级写法。

产物: out/gt_table_chart.pptx, out/table_chart_raw/, out/table_chart_map.json
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

import win32com.client as win32
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
PPTX = os.path.join(OUT, "gt_table_chart.pptx")
RAW = os.path.join(OUT, "table_chart_raw")
SEED = os.path.join(OUT, "chart_seed.pptx")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoTrue = -1


def seed_charts() -> None:
    """用 python-pptx 先铺两张图表页（PowerPoint 后续加动画与重写样式）。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    data = CategoryChartData()
    data.categories = ["Q1", "Q2", "Q3", "Q4"]
    data.add_series("收入", (12.0, 18.0, 15.0, 22.0))
    data.add_series("成本", (8.0, 9.5, 11.0, 12.5))
    s1 = prs.slides.add_slide(blank)
    gf = s1.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED,
                             Inches(0.8), Inches(1.0), Inches(11.8), Inches(5.6), data)
    chart = gf.chart
    chart.has_title = True
    chart.chart_title.text_frame.text = "季度收支"
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.legend.include_in_layout = False
    chart.plots[0].has_data_labels = True
    chart.plots[0].data_labels.show_value = True
    chart.plots[0].data_labels.number_format = "0.0"
    chart.plots[0].data_labels.number_format_is_linked = False
    for s, color in zip(chart.series, ("4C8DFF", "6FE3C4")):
        s.format.fill.solid()
        s.format.fill.fore_color.rgb = __import__("pptx.dml.color", fromlist=["RGBColor"]).RGBColor.from_string(color)

    data2 = CategoryChartData()
    data2.categories = ["产品", "服务", "订阅"]
    data2.add_series("占比", (45.0, 30.0, 25.0))
    s2 = prs.slides.add_slide(blank)
    gf2 = s2.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT,
                              Inches(1.6), Inches(1.2), Inches(9.5), Inches(5.4), data2)
    gf2.chart.has_title = True
    gf2.chart.chart_title.text_frame.text = "收入结构"
    gf2.chart.has_legend = True
    gf2.chart.legend.position = XL_LEGEND_POSITION.RIGHT
    prs.save(SEED)
    print(f"[seed] {SEED}")


def build_table(app) -> dict:
    """表格页：合并、列宽行高、单元格填充、边框、字体。"""
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    slide = pres.Slides.Add(1, ppLayoutBlank)
    shape = slide.Shapes.AddTable(3, 4, 80, 90, 780, 220)
    tbl = shape.Table
    info = {"case": "table", "shape_id": int(shape.Id), "borders_probe": []}

    # 表头合并前两格
    try:
        tbl.Cell(1, 1).Merge(tbl.Cell(1, 2))
    except Exception as exc:  # noqa: BLE001
        info["merge_err"] = repr(exc)
    data = [["指标", "", "Q3", "Q4"], ["收入", "万元", "18.0", "22.0"], ["成本", "万元", "11.0", "12.5"]]
    for r in range(1, 4):
        for c in range(1, 5):
            cell = tbl.Cell(r, c)
            cell.Shape.TextFrame.TextRange.Text = data[r - 1][c - 1]
            cell.Shape.TextFrame.TextRange.Font.Size = 14 if r > 1 else 15
            cell.Shape.TextFrame.TextRange.Font.Bold = msoTrue if r == 1 else 0
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = 0x4A2A1B if r == 1 else (0x38201A if c % 2 else 0x26140F)
    # 列宽 / 行高
    tbl.Columns(1).Width = 140
    tbl.Columns(2).Width = 90
    tbl.Rows(1).Height = 56
    try:
        tbl.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", True)
    except Exception as exc:  # noqa: BLE001
        info["style_err"] = repr(exc)

    # 边框：探查可用属性名
    cell = tbl.Cell(2, 1)
    for n in ("Borders", "Border"):
        if hasattr(cell, n):
            info["borders_probe"].append(n)
    try:
        b = cell.Borders
        info["borders_dir"] = [x for x in dir(b) if not x.startswith("_")][:30]
        # ppBorderTop=1, ppBorderLeft=2, ppBorderBottom=3, ppBorderRight=4
        bd = b(3)   # bottom
        bd.ForeColor.RGB = 0xFF7A45
        bd.Weight = 1.5
        bd.DashStyle = 1
    except Exception as exc:  # noqa: BLE001
        info["border_err"] = repr(exc)

    info["cols"] = [float(tbl.Columns(i).Width) for i in range(1, 5)]
    info["rows"] = [float(tbl.Rows(i).Height) for i in range(1, 4)]
    return {"pres": pres, "info": info}


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    if not os.path.exists(SEED):
        seed_charts()

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone

    # ---- 表格页（第 2 页）
    tbl_built = build_table(app)
    pres = tbl_built["pres"]
    info = tbl_built["info"]

    # ---- 图表页：把 seed 的两页插进来
    seed = app.Presentations.Open(os.path.abspath(SEED), ReadOnly=msoTrue, WithWindow=0)
    seed.Slides(1).Copy()
    pres.Slides.Paste()
    seed.Slides(2).Copy()
    pres.Slides.Paste()
    seed.Close()

    # 给两张图表加动画：柱状按系列、环形按类别
    chart_targets = []
    for idx in (2, 3):
        slide = pres.Slides(idx)
        chart_shape = None
        for j in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(j)
            if getattr(shp, "HasChart", False):
                chart_shape = shp
                break
        if chart_shape is None:
            continue
        level = 10 if idx == 2 else 8    # msoAnimateChartBySeries=10 / msoAnimateChartByCategory=8
        try:
            eff = slide.TimeLine.MainSequence.AddEffect(chart_shape, 10, level, 1)
            eff.Timing.Duration = 1.0
            chart_targets.append({"slide": idx, "shape_id": int(chart_shape.Id),
                                  "level": level, "ok": True})
        except Exception as exc:  # noqa: BLE001
            chart_targets.append({"slide": idx, "level": level, "ok": False, "err": repr(exc)})
    info["chart_anims"] = chart_targets

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    out = {"table": info, "chart_anims": chart_targets, "slides": len(slides)}
    U.save_json(out, os.path.join(OUT, "table_chart_map.json"))

    import re
    for s in slides:
        if s["index"] == 1:
            print("=" * 70, "\ntable XML")
            for pat in (r"<a:tblPr[^>]*>", r"<a:tcPr.*?</a:tcPr>", r"<a:ln[BLTR] [^>]*>",
                        r'gridSpan="\d+"', r'rowSpan="\d+"', r"<a:gridCol w=\"\d+\""):
                for m in re.finditer(pat, s["xml"], re.S):
                    seg = re.sub(r"\s+", " ", m.group(0))
                    print("   ", seg[:300])
                    break
        else:
            tm = U.walk_timing(s["xml"])
            if not tm:
                continue
            print("=" * 70, f"\nslide {s['index']} timing")
            for pat in (r"<p:bldGraphic.*?</p:bldGraphic>", r"<p:bldSub>.*?</p:bldSub>",
                        r"<p:graphicEl>.*?</p:graphicEl>", r"<p:chart[^>]*>.*?</p:chart>"):
                for m in re.finditer(pat, tm["xml"], re.S):
                    seg = re.sub(r"\s+", " ", m.group(0))
                    print("   ", seg[:400])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())