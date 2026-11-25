"""提取每个变体案例的 (nodeType, stCondLst) 序列，用于确定分组/触发写法。"""
import json, os, re

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
d = json.load(open(os.path.join(OUT, "anim_variants.json"), encoding="utf-8"))

for label, info in d.items():
    xml = info.get("timing_xml")
    if not xml:
        continue
    # 找 <p:cTn ... presetID ...> 及其后的 stCondLst
    rows = []
    for m in re.finditer(r'<p:cTn\b[^>]*presetID[^>]*>', xml):
        seg = xml[m.end():m.end() + 400]
        cond = re.search(r'<p:cond\b[^>]*/?>', seg)
        tag = re.search(r'nodeType="([^"]+)"', m.group(0))
        pid = re.search(r'presetID="([^"]+)"', m.group(0))
        cls = re.search(r'presetClass="([^"]+)"', m.group(0))
        rep = re.search(r'repeatCount="([^"]+)"', m.group(0))
        acc = re.search(r'accel="([^"]+)"', m.group(0))
        rows.append({
            "pid": pid.group(1) if pid else None,
            "class": cls.group(1) if cls else None,
            "node": tag.group(1) if tag else None,
            "cond": cond.group(0) if cond else None,
            "repeat": rep.group(1) if rep else None,
            "accel": acc.group(1) if acc else None,
        })
    print("=" * 74)
    print(label, "| effects:", [e.get("value") for e in info["effects"]])
    for r in rows:
        print("   ", r)
    # 组数 = delay="indefinite" 的 cond 数
    print("    click-groups(indefinite):", xml.count('delay="indefinite"'),
          "| txEl:", xml.count("<p:txEl>"), "| bldP:", xml.count("<p:bldP"))