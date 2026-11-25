"""分析 animations_map.json：提取 preset 三要素与行为特征，输出摘要和代表样本。"""
import json, os, collections

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
d = json.load(open(os.path.join(OUT, "animations_map.json"), encoding="utf-8"))

rows = []
for label, v in d.items():
    ps = v["presets"] or []
    head = None
    for p in ps:
        if p["presetClass"] in ("entr", "emph", "exit", "path"):
            head = p
            break
    rows.append({
        "label": label, "value": v["value"], "applied": v["applied"],
        "err": v.get("apply_error"),
        "class": head["presetClass"] if head else None,
        "presetID": head["presetID"] if head else None,
        "subtype": head["presetSubtype"] if head else None,
        "node": head["node"] if head else None,
        "n_presets": len(ps),
        "has_timing": v["timing_xml"] is not None,
    })

lines = [f"total {len(rows)}  with_timing {sum(r['has_timing'] for r in rows)}"]
lines.append("\n== by class ==")
for c, n in collections.Counter(r["class"] for r in rows).most_common():
    lines.append(f"  {c!s:8} {n}")
lines.append("\n== by presetID (top) ==")
for c, n in collections.Counter(r["presetID"] for r in rows).most_common(30):
    lines.append(f"  presetID={c!s:6} {n}")
lines.append("\n== rows ==")
for r in sorted(rows, key=lambda x: (str(x["class"]), int(x["presetID"] or 0), int(x["subtype"] or 0))):
    lines.append(f"{r['value']:5d} {str(r['class']):5s} pid={str(r['presetID']):5s} sub={str(r['subtype']):5s} {r['label']:34s} {r['err'] or ''}")
open(os.path.join(OUT, "animations_summary.txt"), "w", encoding="utf-8").write("\n".join(lines))
print("\n".join(lines[:50]))

samples = ["msoAnimEffectFade", "msoAnimEffectAppear", "msoAnimEffectFly", "msoAnimEffectSpin",
           "msoAnimEffectGrowShrink", "msoAnimEffectPathCircle", "msoAnimEffectZoom",
           "msoAnimEffectWipe", "msoAnimEffectBounce", "msoAnimEffectTransparency"]
with open(os.path.join(OUT, "animations_samples.xml"), "w", encoding="utf-8") as fh:
    for label in samples:
        v = d.get(label)
        if not v or not v["timing_xml"]:
            continue
        fh.write(f"\n<!-- ===== {label} (value {v['value']}) ===== -->\n")
        fh.write(v["timing_xml"])
print("[written] animations_samples.xml / animations_summary.txt")