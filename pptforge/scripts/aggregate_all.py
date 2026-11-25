"""全量聚合：2125 个 pptx 的分布、分位数、板块画像与风格分档。

输出 samples/resource_all_summary.json，供 skill 资产重算使用。
"""
from __future__ import annotations

import json
import os
import statistics
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
S = os.path.join(ROOT, "samples")
IN = os.path.join(S, "resource_deep_all.json")
OUT = os.path.join(S, "resource_all_summary.json")


def short_tr(k: str) -> str:
    return k.split("|")[0].replace("p:", "").replace("p14:", "").replace("p15:", "").replace("p159:", "")


def pct(vals, p):
    vals = sorted(v for v in vals if v is not None)
    if not vals:
        return None
    k = max(0, min(len(vals) - 1, int(round(p / 100 * (len(vals) - 1)))))
    return round(vals[k], 2)


def main() -> int:
    if not os.path.exists(IN):
        print("缺少", IN)
        return 1
    data = json.load(open(IN, encoding="utf-8"))
    ok = [r for r in data if r and not r.get("error")]
    print(f"样本 {len(ok)}/{len(data)}")

    tr, pre, dur = Counter(), Counter(), Counter()
    pages, shapes, anims, layouts, media = [], [], [], [], []
    for r in ok:
        for k, v in (r.get("transitions") or {}).items():
            tr[short_tr(k)] += v
        for item in r.get("presets_top") or []:
            k, v = item.split("×")
            pre[k] += int(v)
        for d, n in r.get("durations_top") or []:
            dur[int(d)] += n
        if r.get("slides"):
            pages.append(r["slides"])
        if r.get("shapes_avg"):
            shapes.append(r["shapes_avg"])
        if r.get("anim_per_slide_avg") is not None:
            anims.append(r["anim_per_slide_avg"])
        if r.get("layouts"):
            layouts.append(r["layouts"])
        if r.get("media"):
            media.append(r["media"])

    # 风格分档：动画密度
    tiers = Counter()
    for a in anims:
        if a == 0:
            tiers["zero"] += 1
        elif a < 3:
            tiers["low"] += 1
        elif a < 8:
            tiers["mid"] += 1
        elif a < 15:
            tiers["high"] += 1
        else:
            tiers["extreme"] += 1

    # 板块画像
    by_group = defaultdict(list)
    for r in ok:
        by_group[r.get("group") or "?"].append(r)
    groups = []
    for g, items in sorted(by_group.items(), key=lambda kv: -len(kv[1])):
        n = len(items)
        gtr, gpre = Counter(), Counter()
        for r in items:
            for k, v in (r.get("transitions") or {}).items():
                gtr[short_tr(k)] += v
            for item in r.get("presets_top") or []:
                k, v = item.split("×")
                gpre[k] += int(v)
        groups.append({
            "group": g, "files": n,
            "pages_median": pct([r.get("slides") for r in items], 50),
            "anim_median": pct([r.get("anim_per_slide_avg") for r in items], 50),
            "shapes_median": pct([r.get("shapes_avg") for r in items], 50),
            "layouts_median": pct([r.get("layouts") for r in items], 50),
            "media_median": pct([r.get("media") for r in items], 50),
            "top_transition": gtr.most_common(1)[0][0] if gtr else "-",
            "top_preset": gpre.most_common(1)[0][0] if gpre else "-",
        })

    summary = {
        "scope": "全部 2125 个 pptx",
        "sampled": len(ok),
        "global": {
            "transitions_pages": dict(tr.most_common(10)),
            "presets": dict(pre.most_common(12)),
            "durations_ms": dict(dur.most_common(10)),
        },
        "distributions": {
            "pages": {"p10": pct(pages, 10), "p50": pct(pages, 50), "p90": pct(pages, 90)},
            "shapes_per_page": {"p10": pct(shapes, 10), "p50": pct(shapes, 50), "p90": pct(shapes, 90)},
            "anim_per_page": {"p10": pct(anims, 10), "p50": pct(anims, 50), "p90": pct(anims, 90)},
            "layouts": {"p10": pct(layouts, 10), "p50": pct(layouts, 50), "p90": pct(layouts, 90)},
            "media": {"p10": pct(media, 10), "p50": pct(media, 50), "p90": pct(media, 90)},
        },
        "motion_tiers": dict(tiers),
        "groups": groups,
    }
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=1)

    print("\n=== 全局切换（页） ===")
    print("  " + " | ".join(f"{k} {v}" for k, v in list(summary["global"]["transitions_pages"].items())[:8]))
    print("=== 全局预设 ===")
    print("  " + " | ".join(f"{k} {v}" for k, v in list(summary["global"]["presets"].items())[:8]))
    print("=== 时长 ===")
    print("  " + " | ".join(f"{k} {v}" for k, v in list(summary["global"]["durations_ms"].items())[:6]))
    print("=== 分布（P10/P50/P90） ===")
    for k, v in summary["distributions"].items():
        print(f"  {k:16s} {v}")
    print("=== 动画密度分档 ===", summary["motion_tiers"])
    print(f"[written] {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())