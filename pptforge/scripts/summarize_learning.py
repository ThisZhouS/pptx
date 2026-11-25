"""汇总 Resource 拆解结果：全局分布 + 每板块画像 + 风格标签统计。"""
from __future__ import annotations

import json
import os
import sys
from collections import Counter, defaultdict

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
S = os.path.join(ROOT, "samples")


def load(name):
    p = os.path.join(S, name)
    if not os.path.exists(p):
        return []
    return json.load(open(p, encoding="utf-8"))


def short_tr(k: str) -> str:
    return k.split("|")[0].replace("p:", "").replace("p14:", "").replace("p15:", "").replace("p159:", "")


def main() -> int:
    deep = load("resource_deep.json")
    vision = load("resource_vision.json")
    ok = [r for r in deep if r and not r.get("error")]
    print(f"结构样本 {len(ok)}/{len(deep)}，识图样本 {len(vision)}")

    # 全局分布
    tr, pre, dur = Counter(), Counter(), Counter()
    for r in ok:
        for k, v in (r.get("transitions") or {}).items():
            tr[short_tr(k)] += v
        for item in r.get("presets_top") or []:
            k, v = item.split("×")
            pre[k] += int(v)
        for d, n in r.get("durations_top") or []:
            dur[int(d)] += n
    print("\n=== 全局切换（按页累计）===")
    print("  " + " | ".join(f"{k} {v}" for k, v in tr.most_common(8)))
    print("=== 全局动画预设 ===")
    print("  " + " | ".join(f"{k} {v}" for k, v in pre.most_common(8)))
    print("=== 全局时长(ms) ===")
    print("  " + " | ".join(f"{k} {v}" for k, v in dur.most_common(6)))

    # 按板块聚合
    by_group = defaultdict(list)
    for r in ok:
        by_group[r.get("group") or "?"].append(r)

    styles = defaultdict(Counter)
    for v in vision:
        feat = (v.get("features") or [{}])[0]
        if feat.get("style"):
            styles[v.get("group") or "?"][feat["style"]] += 1

    rows = []
    for g, items in sorted(by_group.items(), key=lambda kv: -len(kv[1])):
        n = len(items)
        tr_g, pre_g, dur_g = Counter(), Counter(), Counter()
        for r in items:
            for k, v in (r.get("transitions") or {}).items():
                tr_g[short_tr(k)] += v
            for item in r.get("presets_top") or []:
                k, v = item.split("×")
                pre_g[k] += int(v)
            for d, c in r.get("durations_top") or []:
                dur_g[int(d)] += c
        palette = Counter()
        for r in items:
            for k in ("accent1", "accent2", "dk2", "lt2"):
                c = (r.get("colors") or {}).get(k)
                if c:
                    palette[f"{k}:{c}"] += 1
        rows.append({
            "group": g,
            "samples": n,
            "pages_avg": round(sum(r.get("slides") or 0 for r in items) / n, 1),
            "anim_per_page": round(sum(r.get("anim_per_slide_avg") or 0 for r in items) / n, 1),
            "shapes_per_page": round(sum(r.get("shapes_avg") or 0 for r in items) / n, 1),
            "layouts_avg": round(sum(r.get("layouts") or 0 for r in items) / n, 1),
            "media_avg": round(sum(r.get("media") or 0 for r in items) / n, 1),
            "top_transition": tr_g.most_common(1)[0][0] if tr_g else "-",
            "top_preset": pre_g.most_common(1)[0][0] if pre_g else "-",
            "top_duration": dur_g.most_common(1)[0][0] if dur_g else "-",
            "palette": [c for c, _ in palette.most_common(2)],
            "styles": dict(styles.get(g, {})),
        })

    print(f"\n=== 板块画像（{len(rows)} 个）===")
    print(f"{'板块':26s} {'样本':>3s} {'页':>4s} {'动画/页':>6s} {'形状/页':>6s} {'版式':>4s} {'主切换':>10s} {'主预设':>12s} {'时长':>5s}  风格")
    for row in rows:
        st = "/".join(f"{k}×{v}" for k, v in (row["styles"] or {}).items()) or "-"
        print(f"{row['group'][:26]:26s} {row['samples']:3d} {row['pages_avg']:4.0f} "
              f"{row['anim_per_page']:6.1f} {row['shapes_per_page']:6.1f} {row['layouts_avg']:4.0f} "
              f"{row['top_transition']:>10s} {row['top_preset']:>12s} {str(row['top_duration']):>5s}  {st}")

    out = os.path.join(S, "resource_summary.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump({"global": {"transitions": dict(tr), "presets": dict(pre), "durations": dict(dur)},
                   "groups": rows}, fh, ensure_ascii=False, indent=1)
    print(f"\n[written] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
