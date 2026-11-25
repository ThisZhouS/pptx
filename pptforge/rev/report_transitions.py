"""分析 transitions_map.json：抽取每个效果的 OOXML 指纹，打印摘要与样本。"""
import json, os, collections

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
d = json.load(open(os.path.join(OUT, "transitions_map.json"), encoding="utf-8"))

rows = []
for k, v in d.items():
    tr = v["transition"]
    if not tr:
        rows.append((v["label"], v["value"], None, None, None, None, v.get("applied")))
        continue
    child = tr["children"][0] if tr["children"] else None
    ns = tr["child_ns"][0] if tr["child_ns"] else None
    attrs = {kk: vv for kk, vv in tr["attrs"].items() if kk not in ("spd",)}
    el_attrs = {}
    if child and tr["xml"]:
        import re
        m = re.search(r"<([\w:]+)([^>]*)/?>", tr["xml"].split("\n", 1)[-1] if False else tr["xml"])
    rows.append((v["label"], v["value"], child, ns, attrs, tr["holder"], v.get("applied")))

lines = []
lines.append(f"total mapped entries: {len(rows)}")
applied = [r for r in rows if r[6]]
lines.append(f"applied=True: {len(applied)}   applied=False: {len(rows)-len(applied)}")
by_child = collections.Counter(r[2] for r in rows)
lines.append("\n== child element counts ==")
for c, n in by_child.most_common():
    lines.append(f"  {c!s:28s} {n}")

lines.append("\n== rows ==")
for lab, val, child, ns, attrs, holder, ok in sorted(rows, key=lambda r: (r[2] is None, str(r[2]), r[1])):
    lines.append(f"{str(val):5s} {str(child):14s} {str(ns or '')[-24:]:24s} {json.dumps(attrs, ensure_ascii=False):40s} {holder if holder!='p:sld' else '':24s} {lab}")

open(os.path.join(OUT, "transitions_summary.txt"), "w", encoding="utf-8").write("\n".join(lines))
print("\n".join(lines[:80]))
print("...")
print(f"[written] {os.path.join(OUT,'transitions_summary.txt')}")

# 样本 XML
samples = ["ppEffectNone", "ppEffectFade", "ppEffectPushLeft", "ppEffectCubeLeft",
           "ppEffectMorphByObject", "ppEffectMorphByWord", "ppEffectGalleryLeft",
           "ppEffectSwitchLeft", "ppEffectShredStripsIn", "ppEffectDoorsVertical",
           "ppEffectRevealBlackLeft", "ppEffectFadeSmoothly", "ppEffectFlashbulb"]
with open(os.path.join(OUT, "transitions_samples.xml"), "w", encoding="utf-8") as fh:
    for k, v in d.items():
        if v["label"] in samples and v["kind"] == "enum":
            fh.write(f"\n<!-- {v['label']} = {v['value']} -->\n")
            if v["transition"]:
                fh.write(v["transition"]["xml"])
print("[written] transitions_samples.xml")