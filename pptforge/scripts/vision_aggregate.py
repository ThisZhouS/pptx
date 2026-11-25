"""识图聚合：300 个首页的风格标签分布、视觉指纹均值、与结构分档的交叉。"""
from __future__ import annotations

import json
import os
import statistics
from collections import Counter, defaultdict

S = r"E:\pptx\pptforge\samples"
SKILL_ASSETS = r"E:\pptx\pptx-skill\assets"


def main() -> int:
    p = os.path.join(S, "resource_vision_all.json")
    if not os.path.exists(p):
        print("缺少", p)
        return 1
    data = json.load(open(p, encoding="utf-8"))
    ok = [d for d in data if (d.get("features") or [{}])[0].get("style")]
    print(f"样本 {len(ok)}/{len(data)}")
    styles = Counter()
    inks, bris, sats, edges = [], [], [], []
    for d in ok:
        f = d["features"][0]
        styles[f["style"]] += 1
        inks.append(f["ink_ratio"])
        bris.append(f["brightness"])
        sats.append(f["saturation"])
        edges.append(f["edge_density"])

    summary = {
        "scope": f"{len(ok)} 个模板的首页（从 2120 个均匀抽样）",
        "note": "全量 2120 个渲染需 40–60 分钟 COM 调用；识图仅用于风格分布，故均匀抽 300",
        "style_distribution": dict(styles.most_common()),
        "fingerprints": {
            "ink_ratio_median": round(statistics.median(inks), 3),
            "brightness_median": round(statistics.median(bris), 3),
            "saturation_median": round(statistics.median(sats), 3),
            "edge_density_median": round(statistics.median(edges), 3),
        },
        "samples": [{"file": d["file"], "style": d["features"][0]["style"],
                     "ink": d["features"][0]["ink_ratio"],
                     "brightness": d["features"][0]["brightness"]}
                    for d in ok[:12]],
    }
    with open(os.path.join(S, "vision_all_summary.json"), "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=1)

    # 同步进 skill 资产
    with open(os.path.join(SKILL_ASSETS, "vision.json"), "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=1)

    print("=== 风格标签分布 ===")
    for k, v in styles.most_common():
        print(f"  {k:22s} {v:4d}  {v/len(ok)*100:5.1f}%")
    print("=== 视觉指纹中位 ===")
    for k, v in summary["fingerprints"].items():
        print(f"  {k:24s} {v}")
    print(f"[written] {os.path.join(SKILL_ASSETS, 'vision.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())