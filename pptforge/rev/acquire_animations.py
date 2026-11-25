"""逆向第 3 层：ground-truth 动画采样。

对 MsoAnimEffect 的每个取值，让 PowerPoint 加一次效果（单击触发、1 秒），
再把 slideN.xml 的 <p:timing> 抠出来 —— 得到 presetID/presetClass/presetSubtype
与 XML 节点结构的权威对照。

产物
----
out/gt_animations.pptx
out/animations_raw/slide_*.xml
out/animations_map.json
out/animations_labels.json
"""
from __future__ import annotations

import json
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import ooxml_util as U

import win32com.client as win32

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
PPTX = os.path.join(OUT, "gt_animations.pptx")
RAW = os.path.join(OUT, "animations_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoShapeRectangle = 1
msoAnimTriggerOnPageClick = 1
msoAnimateLevelNone = 0


def enum_members() -> dict[str, int]:
    with open(os.path.join(OUT, "typelib_constants.json"), encoding="utf-8") as fh:
        return json.load(fh)["enums"]["MsoAnimEffect"]["members"]


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    members = enum_members()
    plan = [{"label": n, "value": v} for n, v in members.items() if isinstance(v, int) and v > 0]
    print(f"[plan] {len(plan)} animation effects")

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540

    for i, item in enumerate(plan, start=1):
        slide = pres.Slides.Add(i, ppLayoutBlank)
        box = slide.Shapes.AddTextbox(1, 60, 120, 840, 80)
        box.TextFrame.TextRange.Text = f"#{i:03d} {item['label']}"
        box.TextFrame.TextRange.Font.Size = 20
        shape = slide.Shapes.AddShape(msoShapeRectangle, 260, 240, 440, 180)
        shape.Fill.ForeColor.RGB = 0x3366CC
        shape.TextFrame.TextRange.Text = "target"
        try:
            eff = slide.TimeLine.MainSequence.AddEffect(
                shape, item["value"], msoAnimateLevelNone, msoAnimTriggerOnPageClick)
            eff.Timing.Duration = 1.0
            item["applied"] = True
            item["effect_type"] = int(eff.EffectType)
        except Exception as exc:  # noqa: BLE001
            item["applied"] = False
            item["apply_error"] = f"{type(exc).__name__}: {exc}"
        if i % 25 == 0:
            print(f"  ... {i}/{len(plan)}")

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    mapping = {}
    for s in slides:
        item = plan[s["index"] - 1]
        tm = U.walk_timing(s["xml"])
        mapping[item["label"]] = {
            "value": item["value"],
            "applied": item.get("applied"),
            "apply_error": item.get("apply_error"),
            "effect_type": item.get("effect_type"),
            "presets": tm["presets"] if tm else None,
            "bhvr_count": tm["bhvr_count"] if tm else None,
            "timing_xml": tm["xml"] if tm else None,
        }
    U.save_json(mapping, os.path.join(OUT, "animations_map.json"))
    U.save_json({str(s["index"]): plan[s["index"] - 1] for s in slides},
                os.path.join(OUT, "animations_labels.json"))

    ok = [k for k, v in mapping.items() if v["timing_xml"]]
    print(f"[done] with_timing={len(ok)}/{len(mapping)}")
    cls: dict[str, int] = {}
    for v in mapping.values():
        for p in (v["presets"] or []):
            cls[p["presetClass"]] = cls.get(p["presetClass"], 0) + 1
    print("[done] presetClass distribution:", cls)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())