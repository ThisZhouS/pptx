"""挑出 3 个体积小、风格不同的模板放进 skill 作为参考，并生成索引。"""
from __future__ import annotations

import json
import os
import shutil

ROOT = r"E:\pptx"
RES = os.path.join(ROOT, "Resource")
SKILL = os.path.join(ROOT, "pptx-skill")
TPL = os.path.join(SKILL, "templates")
os.makedirs(TPL, exist_ok=True)

deep = json.load(open(os.path.join(ROOT, "pptforge", "samples", "resource_deep.json"),
                      encoding="utf-8"))
ok = [r for r in deep if r and not r.get("error")]

# 目标：三种气质 —— 零动画极简 / 中等动画商务 / 高动画动态
wanted = [
    ("极简参考", lambda r: (r.get("anim_per_slide_avg") or 0) == 0 and (r.get("slides") or 0) >= 12),
    ("商务参考", lambda r: 1 <= (r.get("anim_per_slide_avg") or 0) <= 4 and (r.get("media") or 0) >= 5),
    ("动态参考", lambda r: (r.get("anim_per_slide_avg") or 0) >= 6),
]
index = []
used = set()
for label, pred in wanted:
    cand = [r for r in ok if pred(r) and r["file"] not in used]
    cand.sort(key=lambda r: os.path.getsize(os.path.join(RES, r["file"])))
    for r in cand:
        src = os.path.join(RES, r["file"])
        if not os.path.exists(src) or os.path.getsize(src) > 6 * 1024 * 1024:
            continue
        name = os.path.basename(r["file"]).replace(" ", "_")
        dst = os.path.join(TPL, f"{label}-{name}")
        shutil.copyfile(src, dst)
        used.add(r["file"])
        index.append({
            "file": os.path.basename(dst), "label": label, "origin": r["file"],
            "group": r.get("group"), "pages": r.get("slides"),
            "anim_per_page": r.get("anim_per_slide_avg"),
            "shapes_per_page": r.get("shapes_avg"),
            "layouts": r.get("layouts"), "media": r.get("media"),
            "transition": list((r.get("transitions") or {}).keys())[:1],
            "accent1": (r.get("colors") or {}).get("accent1"),
            "fonts": (r.get("fonts") or [])[:2],
            "size_kb": round(os.path.getsize(dst) / 1024),
        })
        print(f"[pick] {label}: {r['file'][:60]} -> {os.path.basename(dst)} "
              f"({round(os.path.getsize(dst)/1024)}KB)")
        break

with open(os.path.join(TPL, "index.json"), "w", encoding="utf-8") as fh:
    json.dump({"note": "三份参考模板：极简（0 动画）/ 商务（低动画）/ 动态（高动画）",
               "templates": index}, fh, ensure_ascii=False, indent=1)
print(f"[done] {len(index)} templates")
