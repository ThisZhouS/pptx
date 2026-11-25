"""把动画变体按"结构差异"打印：cond / nodeType / build / 行为节点。"""
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
d = json.load(open(os.path.join(OUT, "anim_variants.json"), encoding="utf-8"))

IGNORE = ("timing", "paths", "presets", "ctn_nodes", "timing_xml")
PATTERNS = [
    ("cTn", r"<p:cTn\b[^>]*>"),
    ("cond", r"<p:cond\b[^>]*/?>"),
    ("nodeType", r'nodeType="[^"]+"'),
    ("bld", r"<p:bldP\b[^>]*/?>"),
    ("behaviour", r"<p:(?:animEffect|animMotion|anim|set|animRot|animScale)\b[^>]*>"),
    ("tgt", r"<p:(?:spTgt|sldTgt|grpSpTgt)\b[^>]*/?>"),
]

out = []
for label, info in d.items():
    out.append("=" * 78)
    out.append(f"{label}   effects_ok={[e.get('ok') for e in info['effects']]}")
    xml = info.get("timing_xml") or ""
    if not xml:
        out.append("   <no timing>")
        continue
    for name, pat in PATTERNS:
        hits = re.findall(pat, xml)
        if not hits:
            continue
        uniq = []
        for h in hits:
            if h not in uniq:
                uniq.append(h)
        out.append(f"  [{name}] x{len(hits)}")
        for h in uniq[:8]:
            out.append(f"     {h}")

text = "\n".join(out)
open(os.path.join(OUT, "anim_variants_summary.txt"), "w", encoding="utf-8").write(text)
# 只打印前 6 个 case 到控制台
head = "\n".join(out[:140])
print(head)
print(f"\n[written] {os.path.join(OUT, 'anim_variants_summary.txt')}")