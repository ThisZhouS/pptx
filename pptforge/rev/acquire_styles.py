"""逆向第 5 层：绘图属性（填充/线条/效果/文本/图片/表格）。

真机采样，把每个属性的 XML 落点钉死：
  填充 solid+alpha / 双色渐变 / 单色渐变 / 图案 / 无填充
  线条 虚线 / 箭头 / 透明度 / 粗细
  效果 外阴影 / 发光 / 柔化边缘 / 3D 相机
  文本 粗斜体下划线 / 字体 / 字号 / 颜色 / 段前段后 / 行距 / 项目符号 / 缩进
  图片 裁剪 / 亮度对比度
  表格 表样式 / 单元格填充

产物: out/gt_styles.pptx, out/styles_raw/, out/styles_map.json
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
PPTX = os.path.join(OUT, "gt_styles.pptx")
RAW = os.path.join(OUT, "styles_raw")
IMG = os.path.join(OUT, "test-image.png")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoShapeRectangle = 1
msoTrue = -1
msoFalse = 0


def make_image() -> str:
    if os.path.exists(IMG):
        return IMG
    from PIL import Image, ImageDraw
    img = Image.new("RGB", (640, 400), "#3366CC")
    d = ImageDraw.Draw(img)
    for i in range(0, 640, 40):
        d.line([(i, 0), (i, 400)], fill="#66CCFF", width=3)
    for j in range(0, 400, 40):
        d.line([(0, j), (640, j)], fill="#FFCC66", width=3)
    d.ellipse([180, 100, 460, 300], outline="#FFFFFF", width=6)
    img.save(IMG)
    return IMG


def box(slide, left, top, w, h, text=""):
    shp = slide.Shapes.AddShape(msoShapeRectangle, left, top, w, h)
    if text:
        shp.TextFrame.TextRange.Text = text
    return shp


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    img_path = make_image()
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    meta: dict[int, dict] = {}

    # ---- 1. 填充
    s = pres.Slides.Add(1, ppLayoutBlank)
    a = box(s, 60, 60, 260, 150, "solid+alpha")
    a.Fill.Solid()
    a.Fill.ForeColor.RGB = 0xCC8844
    a.Fill.Transparency = 0.35
    b = box(s, 360, 60, 260, 150, "grad2")
    b.Fill.TwoColorGradient(1, 1)
    b.Fill.ForeColor.RGB = 0x3366CC
    b.Fill.BackColor.RGB = 0x66CCFF
    b.Fill.GradientAngle = 45
    c = box(s, 660, 60, 240, 150, "grad1-center")
    c.Fill.OneColorGradient(7, 2, 0.4)
    c.Fill.ForeColor.RGB = 0xFF6633
    d = box(s, 60, 260, 260, 150, "pattern")
    d.Fill.Patterned(5)
    d.Fill.ForeColor.RGB = 0xFFFFFF
    d.Fill.BackColor.RGB = 0x112233
    e = box(s, 360, 260, 260, 150, "no-fill")
    e.Fill.Visible = msoFalse
    e.Line.ForeColor.RGB = 0xFFFFFF
    f = box(s, 660, 260, 240, 150, "solid50")
    f.Fill.Solid()
    f.Fill.ForeColor.RGB = 0x66FFCC
    f.Fill.Transparency = 0.5
    meta[1] = {"case": "fill", "ids": [int(a.Id), int(b.Id), int(c.Id), int(d.Id), int(e.Id), int(f.Id)]}

    # ---- 2. 线条
    s = pres.Slides.Add(2, ppLayoutBlank)
    l1 = s.Shapes.AddLine(60, 80, 420, 80)
    l1.Line.DashStyle = 4          # msoLineDash
    l1.Line.Weight = 3
    l1.Line.ForeColor.RGB = 0xFFFFFF
    l2 = s.Shapes.AddLine(60, 140, 420, 140)
    l2.Line.BeginArrowheadStyle = 2   # triangle
    l2.Line.EndArrowheadStyle = 3     # open
    l2.Line.Weight = 2
    l2.Line.ForeColor.RGB = 0x66CCFF
    l3 = s.Shapes.AddLine(60, 200, 420, 200)
    l3.Line.Transparency = 0.5
    l3.Line.Weight = 6
    l3.Line.ForeColor.RGB = 0xFFCC66
    l4 = s.Shapes.AddLine(60, 260, 420, 260)
    l4.Line.DashStyle = 3          # round dot
    l4.Line.Weight = 2
    meta[2] = {"case": "line", "ids": [int(l1.Id), int(l2.Id), int(l3.Id), int(l4.Id)]}

    # ---- 3. 效果
    s = pres.Slides.Add(3, ppLayoutBlank)
    e1 = box(s, 80, 80, 240, 150, "shadow")
    e1.Fill.Solid()
    e1.Fill.ForeColor.RGB = 0x333333
    e1.Shadow.Visible = msoTrue
    e1.Shadow.Blur = 12
    e1.Shadow.OffsetX = 6
    e1.Shadow.OffsetY = 6
    e1.Shadow.Transparency = 0.35
    e1.Shadow.ForeColor.RGB = 0x000000
    e2 = box(s, 380, 80, 240, 150, "glow")
    e2.Fill.Solid()
    e2.Fill.ForeColor.RGB = 0x112244
    e2.Glow.Radius = 12
    e2.Glow.Color.RGB = 0x66CCFF
    e2.Glow.Transparency = 0.3
    e3 = box(s, 680, 80, 200, 150, "softedge")
    e3.Fill.Solid()
    e3.Fill.ForeColor.RGB = 0x225533
    e3.SoftEdge.Radius = 10
    e4 = box(s, 80, 300, 240, 140, "3d")
    e4.Fill.Solid()
    e4.Fill.ForeColor.RGB = 0x446688
    try:
        e4.ThreeD.RotationX = 25
        e4.ThreeD.RotationY = 15
        e4.ThreeD.Depth = 20
    except Exception as exc:  # noqa: BLE001
        print("3d err", exc)
    meta[3] = {"case": "effects", "ids": [int(e1.Id), int(e2.Id), int(e3.Id), int(e4.Id)]}

    # ---- 4. 文本与段落
    s = pres.Slides.Add(4, ppLayoutBlank)
    t = s.Shapes.AddTextbox(1, 60, 40, 820, 120)
    tr = t.TextFrame.TextRange
    tr.Text = "加粗 斜体 下划线 删除线 上标"
    tr.Font.Size = 26
    tr.Font.Name = "Arial"
    tr.Font.NameFarEast = "Microsoft YaHei"
    tr.Font.Bold = msoTrue
    tr.Font.Italic = msoTrue
    tr.Font.Underline = msoTrue
    tr.Font.Color.RGB = 0xFFCC66
    tr.Font.Emboss = msoTrue
    # 部分字符单独格式（模拟 run 级混排）
    tr.Characters(1, 3).Font.Bold = msoTrue
    tr.Characters(1, 3).Font.Size = 34
    tr.Characters(1, 3).Font.Color.RGB = 0x66CCFF
    t2 = s.Shapes.AddTextbox(1, 60, 200, 820, 200)
    tr2 = t2.TextFrame.TextRange
    tr2.Text = "第一条要点\r第二条要点\r第三条要点"
    tr2.Font.Size = 20
    pf = tr2.ParagraphFormat
    pf.Alignment = 2               # 居中
    pf.SpaceBefore = 12
    pf.SpaceAfter = 8
    pf.LineRuleWithin = msoTrue
    pf.SpaceWithin = 1.4
    pf.Bullet.Visible = msoTrue
    pf.Bullet.Character = 8226
    pf.Bullet.RelativeSize = 1.1
    tr2.IndentLevel = 2
    meta[4] = {"case": "text", "ids": [int(t.Id), int(t2.Id)],
               "bullet_props": [n for n in dir(pf.Bullet) if not n.startswith("_")]}

    # ---- 5. 图片裁剪与亮度
    s = pres.Slides.Add(5, ppLayoutBlank)
    p1 = s.Shapes.AddPicture(img_path, 0, msoTrue, 60, 60, 360, 240)
    p1.PictureFormat.CropLeft = 0.1
    p1.PictureFormat.CropRight = 0.05
    p1.PictureFormat.CropTop = 0.08
    p2 = s.Shapes.AddPicture(img_path, 0, msoTrue, 480, 60, 360, 240)
    p2.PictureFormat.Brightness = 0.2
    p2.PictureFormat.Contrast = 0.25
    meta[5] = {"case": "picture", "ids": [int(p1.Id), int(p2.Id)]}

    # ---- 6. 表格
    s = pres.Slides.Add(6, ppLayoutBlank)
    shape = s.Shapes.AddTable(3, 3, 80, 80, 600, 200)
    tbl = shape.Table
    try:
        tbl.ApplyStyle("{5940675A-B579-460E-94D1-54222C63F5DA}", True)  # No Style, Table Grid
    except Exception as exc:  # noqa: BLE001
        print("table style err", exc)
    for r in range(1, 4):
        for c in range(1, 4):
            cell = tbl.Cell(r, c)
            cell.Shape.TextFrame.TextRange.Text = f"{r}-{c}"
            cell.Shape.Fill.Solid()
            cell.Shape.Fill.ForeColor.RGB = 0x1B2A4A if r == 1 else 0x0D2038
    meta[6] = {"case": "table", "table_id": int(shape.Id), "style": str(tbl.Style)}

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    out = {}
    for s in slides:
        info = meta[s["index"]]
        out[info["case"]] = {"meta": info, "xml": s["xml"]}
    U.save_json({k: v["meta"] for k, v in out.items()}, os.path.join(OUT, "styles_map.json"))

    # 打印每个 case 里关键节点
    import re
    for case, payload in out.items():
        print("=" * 70)
        print(case)
        xml = payload["xml"]
        for pat in (r"<a:gradFill.*?</a:gradFill>", r"<a:pattFill.*?</a:pattFill>",
                    r"<a:effectLst>.*?</a:effectLst>", r"<a:ln[ >].*?</a:ln>",
                    r"<a:srcRect[^>]*/>", r"<a:lum[^>]*/>", r"<a:tblPr[^>]*>",
                    r"<a:buChar[^>]*/>", r"<a:buClr>.*?</a:buClr>"):
            for m in re.finditer(pat, xml, re.S):
                seg = re.sub(r"\s+", " ", m.group(0))
                print("   ", seg[:400])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
