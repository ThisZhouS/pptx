"""探查图表动画与表格边框的正确 COM 调用方式。"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
SEED = os.path.join(OUT, "chart_seed.pptx")

app = win32.Dispatch("PowerPoint.Application")
app.Visible = True
app.DisplayAlerts = 1

# ---- 图表动画组合探查
pres = app.Presentations.Open(os.path.abspath(SEED), ReadOnly=0, WithWindow=0)
slide = pres.Slides(1)
chart_shape = None
for j in range(1, slide.Shapes.Count + 1):
    shp = slide.Shapes(j)
    print("shape", shp.Name, "type", shp.Type, "hasChart", getattr(shp, "HasChart", "?"))
    if getattr(shp, "HasChart", False):
        chart_shape = shp
if chart_shape is None:
    print("no chart found")
else:
    combos = [(10, 0), (21, 0), (10, 10), (21, 10), (10, 8), (2, 10), (1, 10), (10, 7)]
    for effect_id, level in combos:
        try:
            eff = slide.TimeLine.MainSequence.AddEffect(chart_shape, effect_id, level, 1)
            print(f"  OK   effect={effect_id} level={level} -> {eff.EffectType}")
            slide.TimeLine.MainSequence.Item(1).Delete()
        except Exception as exc:  # noqa: BLE001
            print(f"  FAIL effect={effect_id} level={level}: {exc}")
pres.Close()

# ---- 表格边框
pres2 = app.Presentations.Add()
pres2.PageSetup.SlideWidth = 960
pres2.PageSetup.SlideHeight = 540
s = pres2.Slides.Add(1, 12)
tbl = s.Shapes.AddTable(3, 3, 80, 80, 600, 200).Table
cell = tbl.Cell(2, 2)
print("Borders count:", cell.Borders.Count)
for idx in range(1, 7):
    try:
        bd = cell.Borders.Item(idx)
        bd.ForeColor.RGB = 0xFF7A45
        bd.Weight = 1.5
        bd.DashStyle = 1
        print(f"  border {idx} OK visible={bd.Visible} weight={bd.Weight}")
    except Exception as exc:  # noqa: BLE001
        print(f"  border {idx} FAIL: {exc}")
pres2.SaveAs(os.path.join(OUT, "gt_table_border.pptx"), 24)
pres2.Close()
app.Quit()
print("[saved] gt_table_border.pptx")