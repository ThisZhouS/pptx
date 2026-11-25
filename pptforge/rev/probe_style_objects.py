"""探查 PowerPoint 样式相关 COM 对象（填充/阴影/发光/线条/文本/段落/图片）。"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import win32com.client as win32

app = win32.Dispatch("PowerPoint.Application")
app.Visible = True
app.DisplayAlerts = 1
pres = app.Presentations.Add()
pres.PageSetup.SlideWidth = 960
pres.PageSetup.SlideHeight = 540
slide = pres.Slides.Add(1, 12)
shp = slide.Shapes.AddShape(1, 100, 100, 300, 160)
shp.TextFrame.TextRange.Text = "样例文字 Sample"


def show(obj, title, limit=90):
    names = sorted(n for n in dir(obj) if not n.startswith("_"))
    print(f"--- {title} ({len(names)}) ---")
    print(", ".join(names[:limit]))
    print()


show(shp.Fill, "Shape.Fill")
show(shp.Line, "Shape.Line")
show(shp.Shadow, "Shape.Shadow")
try:
    show(shp.Glow, "Shape.Glow")
except Exception as e:
    print("Glow err:", e)
try:
    show(shp.SoftEdge, "Shape.SoftEdge")
except Exception as e:
    print("SoftEdge err:", e)
try:
    show(shp.ThreeD, "Shape.ThreeD")
except Exception as e:
    print("ThreeD err:", e)
tr = shp.TextFrame.TextRange
show(tr.Font, "TextRange.Font")
show(tr.ParagraphFormat, "TextRange.ParagraphFormat")
show(shp.TextFrame, "TextFrame")

# 渐变 stops
try:
    shp.Fill.OneColorGradient(3, 1, 0.4)
    show(shp.Fill.GradientStops, "Fill.GradientStops")
    gs = shp.Fill.GradientStops
    print("gstops count:", gs.Count)
    print("stop props:", [n for n in dir(gs.Item(1)) if not n.startswith("_")])
    print("GradientStyle:", shp.Fill.GradientStyle, "Variant:", shp.Fill.GradientVariant,
          "Angle:", shp.Fill.GradientAngle if hasattr(shp.Fill, "GradientAngle") else "n/a")
except Exception as e:
    print("gradient err:", e)

# 图片
pic_path = None
for cand in (r"C:\Windows\Web\Wallpaper\Windows\img0.jpg",):
    if os.path.exists(cand):
        pic_path = cand
        break
if pic_path:
    pic = slide.Shapes.AddPicture(pic_path, 0, -1, 500, 100, 300, 200)
    show(pic.PictureFormat, "Picture.PictureFormat")
    print("default crop:", pic.PictureFormat.CropLeft, pic.PictureFormat.CropTop,
          pic.PictureFormat.CropRight, pic.PictureFormat.CropBottom)

# 表格
tbl = slide.Shapes.AddTable(2, 2, 100, 320, 300, 100).Table
show(tbl, "Table")
try:
    print("table style id:", tbl.Style if hasattr(tbl, "Style") else "n/a")
except Exception as e:
    print("table style err:", e)

pres.Close()
app.Quit()