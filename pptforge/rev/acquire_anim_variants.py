"""逆向第 3b 层：动画"参数方言"采样。

基础采样只给出 presetID/presetClass/subtype。真实生成需要知道：
退出类怎么标、三种触发怎么写、时长/延迟/重复/加减速落到哪个属性、
文字按段落分级如何生成 bldP、同页多效果如何分 click 组、路径动画的 path 字符串。

产物: out/gt_anim_variants.pptx, out/anim_variants_raw/, out/anim_variants.json
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
PPTX = os.path.join(OUT, "gt_anim_variants.pptx")
RAW = os.path.join(OUT, "anim_variants_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1
msoShapeRectangle = 1
msoShapeOval = 9

MIXED = {  # 触发方式
    "click": 1, "with_prev": 2, "after_prev": 3, "shape_click": 4,
}
LEVEL = {"none": 0, "all_levels": 1, "by_first": 2, "by_second": 3}


def load_values() -> dict[str, int]:
    with open(os.path.join(OUT, "animations_map.json"), encoding="utf-8") as fh:
        data = json.load(fh)
    return {k: v["value"] for k, v in data.items()}


CASES = [
    # ---- 触发方式 ----
    {"label": "entr_click", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "trigger": "click"}]},
    {"label": "entr_with_prev", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "trigger": "with_prev"}]},
    {"label": "entr_after_prev", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "trigger": "after_prev"}]},
    {"label": "entr_shape_click", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "trigger": "shape_click"}]},
    # ---- 时间参数 ----
    {"label": "entr_dur05_delay03", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "timing": {"Duration": 0.5, "TriggerDelayTime": 0.3}}]},
    {"label": "entr_repeat2", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "timing": {"RepeatCount": 2}}]},
    {"label": "entr_accel", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "timing": {"Acceleration": 0.5, "Deceleration": 0.5}}]},
    {"label": "entr_dur2_accel_full", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFly", "timing": {"Duration": 2.0, "Acceleration": 1.0, "Deceleration": 1.0}}]},
    # ---- 退出 ----
    {"label": "exit_fade", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "effect": {"Exit": 1}}]},
    {"label": "exit_zoom_after", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectZoom", "trigger": "after_prev", "effect": {"Exit": 1}}]},
    # ---- 强调 ----
    {"label": "emph_spin_after", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectSpin", "trigger": "after_prev", "timing": {"Duration": 1.5}}]},
    {"label": "emph_pulse", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectGrowShrink"}]},
    # ---- 文字分级 ----
    {"label": "text_by_first_level", "shapes": [{"kind": "text", "paras": ["一级要点 A", "一级要点 B", "一级要点 C"]}],
     "effects": [{"value": "msoAnimEffectFade", "level": "by_first"}]},
    {"label": "text_all_levels", "shapes": [{"kind": "text", "paras": ["一级", "  二级", "    三级"]}],
     "effects": [{"value": "msoAnimEffectFade", "level": "all_levels"}]},
    # ---- 多效果 / click 组 ----
    {"label": "two_click_effects", "shapes": [{"kind": "rect", "text": "A"}, {"kind": "rect", "text": "B", "x": 560}],
     "effects": [{"value": "msoAnimEffectFade", "shape": 0, "trigger": "click"},
                 {"value": "msoAnimEffectZoom", "shape": 1, "trigger": "click"}]},
    {"label": "two_with_effects", "shapes": [{"kind": "rect", "text": "A"}, {"kind": "rect", "text": "B", "x": 560}],
     "effects": [{"value": "msoAnimEffectFade", "shape": 0, "trigger": "click"},
                 {"value": "msoAnimEffectZoom", "shape": 1, "trigger": "with_prev"}]},
    {"label": "three_mixed", "shapes": [{"kind": "rect", "text": "A"}, {"kind": "rect", "text": "B", "x": 560},
                                        {"kind": "oval", "text": "C", "x": 300, "y": 380}],
     "effects": [{"value": "msoAnimEffectFade", "shape": 0, "trigger": "click"},
                 {"value": "msoAnimEffectFly", "shape": 1, "trigger": "with_prev",
                  "timing": {"Duration": 0.75}},
                 {"value": "msoAnimEffectSpin", "shape": 2, "trigger": "after_prev"}]},
    # ---- 路径 ----
    {"label": "path_circle", "shapes": [{"kind": "oval", "text": "A"}],
     "effects": [{"value": "msoAnimEffectPathCircle", "timing": {"Duration": 2.0}}]},
    {"label": "path_zigzag", "shapes": [{"kind": "oval", "text": "A"}],
     "effects": [{"value": "msoAnimEffectPathZigzag", "timing": {"Duration": 2.0}}]},
    {"label": "path_bounce_left", "shapes": [{"kind": "oval", "text": "A"}],
     "effects": [{"value": "msoAnimEffectPathBounceLeft", "timing": {"Duration": 1.0}}]},
    # ---- 播放后行为 ----
    {"label": "entr_dim_after", "shapes": [{"kind": "rect", "text": "A"}],
     "effects": [{"value": "msoAnimEffectFade", "effect": {"AfterEffect": 1}}]},
]


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    values = load_values()
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    meta: dict[int, dict] = {}

    for i, case in enumerate(CASES, start=1):
        slide = pres.Slides.Add(i, ppLayoutBlank)
        shapes = []
        for spec in case["shapes"]:
            x = spec.get("x", 120)
            y = spec.get("y", 180)
            if spec["kind"] == "text":
                tb = slide.Shapes.AddTextbox(1, 80, 90, 800, 340)
                tr = tb.TextFrame.TextRange
                tr.Text = spec["paras"][0]
                for extra in spec["paras"][1:]:
                    tr.InsertAfter("\r" + extra)
                tr.Font.Size = 24
                shapes.append(tb)
            else:
                kind = msoShapeOval if spec["kind"] == "oval" else msoShapeRectangle
                sh = slide.Shapes.AddShape(kind, x, y, 260, 140)
                sh.TextFrame.TextRange.Text = spec.get("text", "")
                sh.Fill.ForeColor.RGB = 0x4477AA
                shapes.append(sh)
        info = {"label": case["label"], "effects": []}
        seq = slide.TimeLine.MainSequence
        for es in case["effects"]:
            rec = dict(es)
            try:
                shp = shapes[es.get("shape", 0)]
                eff = seq.AddEffect(shp, values[es["value"]], LEVEL[es.get("level", "none")],
                                    MIXED[es.get("trigger", "click")])
                for k, v in (es.get("effect") or {}).items():
                    setattr(eff, k, v)
                for k, v in (es.get("timing") or {}).items():
                    setattr(eff.Timing, k, v)
                rec["ok"] = True
                rec["shape_id"] = int(shp.Id)
            except Exception as exc:  # noqa: BLE001
                rec["ok"] = False
                rec["err"] = f"{type(exc).__name__}: {exc}"
            info["effects"].append(rec)
        meta[i] = info

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()
    print(f"[save] {PPTX}")

    slides = U.dump_pptx(PPTX, RAW)
    out = {}
    for s in slides:
        info = meta[s["index"]]
        tm = U.walk_timing(s["xml"])
        info["timing_xml"] = tm["xml"] if tm else None
        info["presets"] = tm["presets"] if tm else None
        info["paths"] = tm["paths"] if tm else None
        out[info["label"]] = info
        # 关键属性摘要
        if tm:
            import re
            attrs = re.findall(r'<p:cTn[^>]*presetID[^>]*>', tm["xml"])
            info["ctn_nodes"] = attrs
    U.save_json(out, os.path.join(OUT, "anim_variants.json"))

    for label, info in out.items():
        print("=" * 66)
        print(label, "| effects ok:", [e.get("ok") for e in info["effects"]])
        for n in (info.get("presets") or []):
            print("   ", n)
        for e in info["effects"]:
            if not e.get("ok"):
                print("    ERR", e.get("err"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())