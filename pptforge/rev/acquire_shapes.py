"""逆向第 4 层：形状层的组合、变换、层级与组动画。

覆盖真机采样：
  · 组合后 <p:grpSp> 的结构（a:off/a:ext 与 a:chOff/a:chExt 的关系）
  · 组的旋转/翻转、组整体移动/缩放后子形状坐标是否变化
  · 组合动画的 target 写法（p:grpSpTgt vs p:spTgt）
  · 组内单形状动画的 target 写法
  · 形状属性：旋转、翻转、透明度、线宽、渐变填充、层级顺序

产物: out/gt_shapes.pptx, out/shapes_raw/, out/shapes_map.json
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
PPTX = os.path.join(OUT, "gt_shapes.pptx")
RAW = os.path.join(OUT, "shapes_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoShapeRectangle = 1
msoShapeOval = 9
msoTrue = -1
msoFalse = 0
msoFlipHorizontal = 0
msoFlipVertical = 1
msoAnimEffectFade = 10
msoAnimEffectSpin = 8


def add_box(slide, left, top, w, h, text, oval=False):
    shp = slide.Shapes.AddShape(msoShapeOval if oval else msoShapeRectangle, left, top, w, h)
    shp.TextFrame.TextRange.Text = text
    shp.Fill.ForeColor.RGB = 0x4488CC
    return shp


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    meta: dict[int, dict] = {}

    # ---- 1. 纯组合：看 grpSp / chOff / chExt
    s1 = pres.Slides.Add(1, ppLayoutBlank)
    a = add_box(s1, 100, 100, 200, 100, "A")
    b = add_box(s1, 350, 160, 160, 120, "B")
    c = add_box(s1, 200, 300, 240, 90, "C", oval=True)
    g = s1.Shapes.Range().Group()
    meta[1] = {"case": "group_only", "group_name": g.Name, "group_shape_id": int(g.Id),
               "children": [int(a.Id), int(b.Id), int(c.Id)],
               "child_positions": [[float(a.Left), float(a.Top), float(a.Width), float(a.Height)],
                                   [float(b.Left), float(b.Top), float(b.Width), float(b.Height)],
                                   [float(c.Left), float(c.Top), float(c.Width), float(c.Height)]]}

    # ---- 2. 组 + 组动画
    s2 = pres.Slides.Add(2, ppLayoutBlank)
    a2 = add_box(s2, 120, 120, 200, 100, "A")
    b2 = add_box(s2, 380, 200, 180, 120, "B")
    g2 = s2.Shapes.Range().Group()
    eff = s2.TimeLine.MainSequence.AddEffect(g2, msoAnimEffectFade, 0, 1)
    meta[2] = {"case": "group_anim", "group_shape_id": int(g2.Id),
               "children": [int(a2.Id), int(b2.Id)]}

    # ---- 3. 组内单形状动画
    s3 = pres.Slides.Add(3, ppLayoutBlank)
    a3 = add_box(s3, 120, 120, 200, 100, "A")
    b3 = add_box(s3, 380, 200, 180, 120, "B")
    g3 = s3.Shapes.Range().Group()
    eff3 = s3.TimeLine.MainSequence.AddEffect(g3.GroupItems(1), msoAnimEffectSpin, 0, 1)
    meta[3] = {"case": "child_anim", "group_shape_id": int(g3.Id),
               "children": [int(a3.Id), int(b3.Id)], "animated_child": int(a3.Id)}

    # ---- 4. 旋转 / 翻转
    s4 = pres.Slides.Add(4, ppLayoutBlank)
    r = add_box(s4, 120, 120, 240, 120, "rot30")
    r.Rotation = 30
    f = add_box(s4, 420, 120, 240, 120, "flipH")
    f.Flip(msoFlipHorizontal)
    v = add_box(s4, 120, 320, 200, 100, "flipV")
    v.Flip(msoFlipVertical)
    meta[4] = {"case": "rotate_flip", "ids": [int(r.Id), int(f.Id), int(v.Id)]}

    # ---- 5. 透明度 / 线宽 / 渐变
    s5 = pres.Slides.Add(5, ppLayoutBlank)
    t1 = add_box(s5, 120, 120, 240, 120, "alpha50")
    t1.Fill.Transparency = 0.5
    t2 = add_box(s5, 420, 120, 240, 120, "line2pt")
    t2.Line.Weight = 2
    t2.Line.ForeColor.RGB = 0xFF6600
    t3 = add_box(s5, 120, 320, 240, 120, "gradient")
    t3.Fill.OneColorGradient(3, 1, 0.4)   # msoGradientDiagonalUp, 变体 1, 亮度 0.4
    meta[5] = {"case": "style", "ids": [int(t1.Id), int(t2.Id), int(t3.Id)]}

    # ---- 6. 组合后整体移动
    s6 = pres.Slides.Add(6, ppLayoutBlank)
    a6 = add_box(s6, 100, 100, 200, 100, "A")
    b6 = add_box(s6, 360, 180, 160, 120, "B")
    g6 = s6.Shapes.Range().Group()
    before = [float(a6.Left), float(a6.Top)]
    g6.Left = g6.Left + 120
    g6.Top = g6.Top + 40
    meta[6] = {"case": "group_move", "group_shape_id": int(g6.Id),
               "children": [int(a6.Id), int(b6.Id)],
               "child_before": before,
               "child_after": [float(a6.Left), float(a6.Top)],
               "group_box": [float(g6.Left), float(g6.Top), float(g6.Width), float(g6.Height)]}

    # ---- 7. 组合后整体缩放
    s7 = pres.Slides.Add(7, ppLayoutBlank)
    a7 = add_box(s7, 100, 100, 200, 100, "A")
    b7 = add_box(s7, 360, 180, 160, 120, "B")
    g7 = s7.Shapes.Range().Group()
    box_before = [float(g7.Left), float(g7.Top), float(g7.Width), float(g7.Height)]
    child_before = [float(a7.Left), float(a7.Top), float(a7.Width), float(a7.Height)]
    g7.Width = g7.Width * 0.6
    g7.Height = g7.Height * 0.6
    meta[7] = {"case": "group_scale", "group_shape_id": int(g7.Id),
               "children": [int(a7.Id), int(b7.Id)],
               "group_before": box_before,
               "group_after": [float(g7.Left), float(g7.Top), float(g7.Width), float(g7.Height)],
               "child_before": child_before,
               "child_after": [float(a7.Left), float(a7.Top), float(a7.Width), float(a7.Height)]}

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    out = {}
    for s in slides:
        info = meta[s["index"]]
        info["shape_tree"] = U.shape_tree(s["xml"])
        tm = U.walk_timing(s["xml"])
        info["timing_xml"] = tm["xml"] if tm else None
        info["targets"] = tm["paths"] if tm else None
        out[info["case"]] = info
    U.save_json(out, os.path.join(OUT, "shapes_map.json"))

    for case, info in out.items():
        print("=" * 70)
        print(case, json.dumps({k: v for k, v in info.items()
                                if k in ("group_shape_id", "children", "child_positions",
                                         "child_before", "child_after", "group_before",
                                         "group_after", "group_box", "animated_child")},
                               ensure_ascii=False))
        for node in info["shape_tree"]:
            print("   ", json.dumps(node, ensure_ascii=False))
        if info.get("timing_xml"):
            import re
            print("    targets:", re.findall(r"<p:(?:spTgt|grpSpTgt)[^>]*>", info["timing_xml"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())