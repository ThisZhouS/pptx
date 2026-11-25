"""逆向第 2 层：ground-truth 切换效果采样。

对 PpEntryEffect 里每一个取值，让真实 PowerPoint 设置一次幻灯片切换并保存，
再把 slideN.xml 里的 <p:transition> 抠出来 —— 得到"枚举值 → OOXML 方言"的权威映射表。

产物
----
out/gt_transitions.pptx          真实 PowerPoint 写出的样本
out/transitions_raw/slide_*.xml  逐页 XML
out/transitions_map.json         值 → {name, prst/p14 元素, attrs, xml}
out/transitions_labels.json      页序 → 取值 的对照表
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
PPTX = os.path.join(OUT, "gt_transitions.pptx")
RAW = os.path.join(OUT, "transitions_raw")

ppSaveAsOpenXMLPresentation = 24
ppLayoutBlank = 12
ppAlertsNone = 1


def enum_members() -> dict[str, int]:
    with open(os.path.join(OUT, "typelib_constants.json"), encoding="utf-8") as fh:
        data = json.load(fh)
    return {k: v for k, v in data["enums"]["PpEntryEffect"]["members"].items() if isinstance(v, int)}


def build(app, plan: list[dict]) -> None:
    pres = app.Presentations.Add()
    pres.PageSetup.SlideWidth = 960
    pres.PageSetup.SlideHeight = 540
    total = len(plan)
    for i, item in enumerate(plan, start=1):
        slide = pres.Slides.Add(i, ppLayoutBlank)
        box = slide.Shapes.AddTextbox(1, 60, 200, 840, 140)
        tf = box.TextFrame
        tf.TextRange.Text = f"#{i:03d} {item['label']}"
        tf.TextRange.Font.Size = 28

        t = slide.SlideShowTransition
        try:
            t.EntryEffect = item["value"]
            item["applied"] = True
        except Exception as exc:  # noqa: BLE001
            item["applied"] = False
            item["apply_error"] = f"{type(exc).__name__}: {exc}"
        for attr, val in item.get("props", {}).items():
            try:
                setattr(t, attr, val)
            except Exception as exc:  # noqa: BLE001
                item.setdefault("prop_errors", {})[attr] = f"{type(exc).__name__}: {exc}"
        if i % 40 == 0:
            print(f"  ... {i}/{total} slides built")

    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    print(f"[save] {PPTX}")


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    members = enum_members()
    plan: list[dict] = []
    for name, value in members.items():
        if value < 0:
            continue
        plan.append({"label": name, "value": value, "kind": "enum",
                     "props": {"Duration": 1.0, "Speed": 2}})
    # 变体采样：同一效果的不同时长 / 速度
    variants = [
        ("ppEffectFade", 1793, {"Duration": 0.5, "Speed": 2}),
        ("ppEffectFade", 1793, {"Duration": 2.5, "Speed": 2}),
        ("ppEffectFade", 1793, {"Duration": 1.0, "Speed": 1}),
        ("ppEffectFade", 1793, {"Duration": 1.0, "Speed": 3}),
        ("ppEffectPushLeft", 1281, {"Duration": 0.25, "Speed": 3}),
        ("ppEffectWipeRight", 1026, {"Duration": 3.0, "Speed": 1}),
    ]
    for name, value, props in variants:
        plan.append({"label": f"{name}@{props}", "value": value, "kind": "variant", "props": props})

    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    try:
        build(app, plan)
    except Exception:
        print(traceback.format_exc())
        try:
            app.Quit()
        except Exception:
            pass
        return 1

    slides = U.dump_pptx(PPTX, RAW)
    labels = {s["index"]: plan[s["index"] - 1] for s in slides if s["index"] - 1 < len(plan)}
    mapping: dict[str, dict] = {}
    for s in slides:
        info = U.find_transition(s["xml"])
        item = labels.get(s["index"], {})
        key = f"{item.get('value')}|{item.get('label')}"
        mapping[key] = {
            "label": item.get("label"),
            "value": item.get("value"),
            "kind": item.get("kind"),
            "requested_props": item.get("props"),
            "applied": item.get("applied"),
            "apply_error": item.get("apply_error"),
            "transition": info,
        }
    U.save_json(mapping, os.path.join(OUT, "transitions_map.json"))
    U.save_json(labels, os.path.join(OUT, "transitions_labels.json"))
    app.Quit()

    ok = sum(1 for v in mapping.values() if v["transition"])
    print(f"[done] pages={len(slides)} with_transition={ok} mapped={len(mapping)}")
    # 只看样式指纹
    seen = {}
    for v in mapping.values():
        tr = v["transition"]
        if not tr:
            continue
        sig = ",".join(tr["children"]) + "|" + ",".join(sorted(tr["attrs"]))
        seen.setdefault(sig, []).append(v["label"])
    print(f"[done] distinct transition signatures: {len(seen)}")
    for sig, names in list(seen.items())[:40]:
        print(f"   {sig}  <- {len(names)} effects, e.g. {names[:3]}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())