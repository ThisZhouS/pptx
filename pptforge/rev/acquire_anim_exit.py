"""逆向第 3c 层：exit（退出）类动画采样。

基础采样只覆盖 entr/emph/path；exit 在 COM 里是 effect.Exit=True 变体。
这里对同一批效果再采样一遍，得到 exit 的行为节点模板（transition="out" 等）。

产物: out/gt_anim_exit.pptx, out/anim_exit_raw/, out/anim_exit_map.json
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
PPTX = os.path.join(OUT, "gt_anim_exit.pptx")
RAW = os.path.join(OUT, "anim_exit_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoShapeRectangle = 1


def main() -> int:
    with open(os.path.join(OUT, "animations_map.json"), encoding="utf-8") as fh:
        base = json.load(fh)
    plan = [{"label": k, "value": v["value"]} for k, v in base.items() if v.get("applied")]
    print(f"[plan] {len(plan)} effects for exit variant")

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    for i, item in enumerate(plan, start=1):
        slide = pres.Slides.Add(i, ppLayoutBlank)
        shape = slide.Shapes.AddShape(msoShapeRectangle, 260, 240, 440, 180)
        shape.Fill.ForeColor.RGB = 0x3366CC
        shape.TextFrame.TextRange.Text = "target"
        try:
            eff = slide.TimeLine.MainSequence.AddEffect(shape, item["value"], 0, 1)
            eff.Exit = 1
            eff.Timing.Duration = 1.0
            item["applied"] = True
        except Exception as exc:  # noqa: BLE001
            item["applied"] = False
            item["err"] = f"{type(exc).__name__}: {exc}"
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
            "err": item.get("err"),
            "presets": tm["presets"] if tm else None,
            "timing_xml": tm["xml"] if tm else None,
        }
    U.save_json(mapping, os.path.join(OUT, "anim_exit_map.json"))
    cls: dict[str, int] = {}
    for v in mapping.values():
        for p in (v["presets"] or []):
            cls[p["presetClass"]] = cls.get(p["presetClass"], 0) + 1
    print("[done] presetClass:", cls)
    print("[done] with_timing:", sum(1 for v in mapping.values() if v["timing_xml"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())