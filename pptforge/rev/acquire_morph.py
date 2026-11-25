"""逆向第 8 层：morph（平滑）切换 + 元素命名匹配 + 自动换片时间轴。

要验证的事：
  1. 连续多页写 <p159:morph> 是否被 PowerPoint 接受（EntryEffect = 3954）
  2. 页间"同一元素"是否靠 cNvPr@name 匹配；PowerPoint 官方的 "!!" 前缀强制匹配是否只是命名约定
  3. 自动换片 advTm + p14:dur 能否组成"无需点击的时间轴"
  4. 元素跨页改变 位置/尺寸/旋转/透明度 后，XML 里各是什么

产物: out/gt_morph.pptx, out/morph_raw/, out/morph_map.json
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

import win32com.client as win32
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
SEED = os.path.join(OUT, "morph_seed.pptx")
PPTX = os.path.join(OUT, "gt_morph.pptx")
RAW = os.path.join(OUT, "morph_raw")

ppSaveAsOpenXMLPresentation = 24
ppAlertsNone = 1


def build_seed() -> None:
    """3 页：同一元素 !!hero 逐帧位移/放大/变色；一个伴生元素 side 同步漂移。"""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    frames = [
        {"hero": (1.0, 1.0, 2.4, 1.2, "4C8DFF"), "side": (9.5, 1.0, 2.0, 1.0, "FFD166")},
        {"hero": (4.2, 3.0, 5.0, 2.0, "6FE3C4"), "side": (10.5, 4.5, 1.4, 1.4, "FF7A45")},
        {"hero": (1.4, 5.2, 3.0, 1.4, "B388FF"), "side": (2.0, 0.6, 2.6, 0.8, "36C2FF")},
    ]
    for i, frame in enumerate(frames, start=1):
        slide = prs.slides.add_slide(blank)
        for name, (x, y, w, h, color) in frame.items():
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                           Inches(x), Inches(y), Inches(w), Inches(h))
            shape.name = f"!!{name}"          # PowerPoint morph 的强制匹配约定
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor.from_string(color)
            shape.line.fill.background()
            tf = shape.text_frame
            tf.text = f"frame {i}"
            tf.paragraphs[0].runs[0].font.size = Pt(14)
            tf.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("101828")
    prs.save(SEED)
    print(f"[seed] {SEED}（3 页，元素名 !!hero / !!side）")


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    if not os.path.exists(SEED):
        build_seed()

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Open(os.path.abspath(SEED), ReadOnly=0, WithWindow=0)

    info = {"slides": pres.Slides.Count, "frames": []}
    for i in range(1, pres.Slides.Count + 1):
        slide = pres.Slides(i)
        t = slide.SlideShowTransition
        # morph 用 COM 写（对应 ppEffectMorphByObject=3954）
        try:
            t.EntryEffect = 3954
            t.Duration = 1.2
            t.AdvanceOnClick = 0
            t.AdvanceOnTime = -1
            t.AdvanceTime = 1.5
            info["frames"].append({"slide": i, "entry": int(t.EntryEffect),
                                   "duration": round(float(t.Duration), 2),
                                   "advance_time": round(float(t.AdvanceTime), 2),
                                   "advance_on_click": int(t.AdvanceOnClick)})
        except Exception as exc:  # noqa: BLE001
            info["frames"].append({"slide": i, "error": repr(exc)})
    # 放映时用排练计时自动推进
    ss = pres.SlideShowSettings
    ss.AdvanceMode = 2      # ppSlideShowUseSlideTimings
    ss.ShowWithAnimation = -1
    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    out = {"com": info, "pages": []}
    import re
    for s in slides:
        tr = U.find_transition(s["xml"])
        names = re.findall(r'<p:cNvPr id="(\d+)" name="([^"]+)"', s["xml"])
        adv = re.search(r'advTm="(\d+)"', s["xml"])
        out["pages"].append({
            "index": s["index"],
            "transition": (tr or {}).get("xml"),
            "advTm": adv.group(1) if adv else None,
            "shapes": [(i2, n) for i2, n in names if n.startswith("!!")],
        })
    U.save_json(out, os.path.join(OUT, "morph_map.json"))

    for p in out["pages"]:
        print(f"--- slide {p['index']} advTm={p['advTm']} shapes={p['shapes']}")
        print("   ", (p["transition"] or "").replace("\n", " ")[:260])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())