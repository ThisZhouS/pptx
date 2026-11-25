"""逆向第 2b 层：切换的"属性方言"采样（时长 / 换片方式 / 声音 / 黑场）。

第一轮采样证明 spd 由 Speed 决定、Duration 没落地。这里逐项单变量试，
把 PowerPoint 实际写的属性名与取值形式钉死。

产物: out/gt_transition_props.pptx, out/transition_props.json
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
PPTX = os.path.join(OUT, "gt_transition_props.pptx")
RAW = os.path.join(OUT, "transition_props_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
ppEffectFade = 1793
ppEffectPushLeft = 3853
ppEffectMorphByObject = 3954

CASES = [
    # (label, entryEffect, kwargs)
    ("fade_base", ppEffectFade, {"Speed": 2}),
    ("fade_dur_only", ppEffectFade, {"Duration": 2.5}),
    ("fade_dur03_fast", ppEffectFade, {"Duration": 0.3, "Speed": 3}),
    ("fade_slow", ppEffectFade, {"Speed": 1}),
    ("fade_fast", ppEffectFade, {"Speed": 3}),
    ("fade_adv_click_off", ppEffectFade, {"AdvanceOnClick": 0}),
    ("fade_adv_time3", ppEffectFade, {"AdvanceOnTime": -1, "AdvanceTime": 3.0}),
    ("fade_adv_click_off_time1", ppEffectFade, {"AdvanceOnClick": 0, "AdvanceOnTime": -1, "AdvanceTime": 1.0}),
    ("morph_dur2", ppEffectMorphByObject, {"Duration": 2.0}),
    ("push_dur15", ppEffectPushLeft, {"Duration": 1.5, "Speed": 2}),
    ("none_plain", 0, {}),
    ("none_adv_time2", 0, {"AdvanceOnTime": -1, "AdvanceTime": 2.0}),
]


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    meta = {}
    for i, (label, effect, kwargs) in enumerate(CASES, start=1):
        slide = pres.Slides.Add(i, ppLayoutBlank)
        box = slide.Shapes.AddTextbox(1, 60, 200, 840, 120)
        box.TextFrame.TextRange.Text = f"{i:02d} {label}"
        t = slide.SlideShowTransition
        info = {"label": label, "entry_effect": effect, "kwargs": kwargs, "results": {}}
        try:
            t.EntryEffect = effect
        except Exception as exc:  # noqa: BLE001
            info["results"]["EntryEffect"] = f"ERR {exc}"
        for k, v in kwargs.items():
            try:
                setattr(t, k, v)
                info["results"][k] = ("ok", getattr(t, k, None))
            except Exception as exc:  # noqa: BLE001
                info["results"][k] = f"ERR {exc}"
        meta[i] = info
    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()

    slides = U.dump_pptx(PPTX, RAW)
    out = {}
    for s in slides:
        info = U.find_transition(s["xml"])
        meta[s["index"]]["transition"] = info
        out[meta[s["index"]]["label"]] = meta[s["index"]]
    U.save_json(out, os.path.join(OUT, "transition_props.json"))

    for label, info in out.items():
        tr = info.get("transition")
        print("=" * 60)
        print(label, "| kwargs:", info["kwargs"], "| applied:", info["results"])
        if tr:
            print("  attrs:", tr["attrs"], "| children:", tr["children"], "| holder:", tr["holder"])
            print("  " + tr["xml"].replace("\n", "\n  ").strip()[:600])
        else:
            print("  <no transition>")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())