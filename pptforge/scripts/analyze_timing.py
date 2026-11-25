"""拆解高动画模板的时序配方：每页几组 click、组内几条、delay 错峰间隔、预设组合。

只看动画密度高的样本（anim_per_slide_avg 排序），因为目标是"要动画"的版本。
"""
from __future__ import annotations

import json
import os
import re
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RES = os.path.join(os.path.dirname(ROOT), "Resource")
S = os.path.join(ROOT, "samples")

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def timing_rows(xml: str) -> list[dict]:
    """按文档顺序取出每个效果的 (preset, class, nodeType, delay, dur)。"""
    rows = []
    # 效果 par：<p:cTn ... presetID=... > 之后紧跟 stCondLst/cond@delay，行为节点里有 dur
    for m in re.finditer(
            r'<p:cTn id="\d+" presetID="(\d+)" presetClass="(\w+)" presetSubtype="(\d+)"([^>]*)>'
            r'.*?<p:cond delay="(\d+)"', xml, re.S):
        pid, cls, sub, attrs, delay = m.groups()
        node = re.search(r'nodeType="(\w+)"', attrs)
        tail = xml[m.end():m.end() + 1200]
        dur = re.search(r'<p:cTn id="\d+" dur="(\d+)"', tail)
        spid = re.search(r'<p:spTgt spid="(\d+)"', tail)
        rows.append({"preset": pid, "class": cls, "node": node.group(1) if node else "?",
                     "delay": int(delay), "dur": int(dur.group(1)) if dur else None,
                     "spid": spid.group(1) if spid else None})
    return rows


def page_recipe(rows: list[dict]) -> dict:
    """把一页的动画序列归纳成配方。"""
    if not rows:
        return {"count": 0}
    groups, cur = [], []
    for r in rows:
        if r["node"] == "clickEffect":
            if cur:
                groups.append(cur)
            cur = [r]
        else:
            cur.append(r)
    if cur:
        groups.append(cur)
    delays = [r["delay"] for r in rows if r["delay"] not in (0,)]
    durs = [r["dur"] for r in rows if r["dur"] and r["dur"] != 1]
    return {
        "count": len(rows),
        "click_groups": sum(1 for r in rows if r["node"] == "clickEffect"),
        "with_prev": sum(1 for r in rows if r["node"] == "withEffect"),
        "after_prev": sum(1 for r in rows if r["node"] == "afterEffect"),
        "group_sizes": [len(g) for g in groups],
        "presets": [r["preset"] + r["class"][:1] for r in rows],
        "max_delay": max(delays) if delays else 0,
        "median_dur": sorted(durs)[len(durs) // 2] if durs else None,
        "total_dur_ms": sum(durs) if durs else 0,
    }


def main() -> int:
    deep = json.load(open(os.path.join(S, "resource_deep.json"), encoding="utf-8"))
    ok = [r for r in deep if r and not r.get("error") and (r.get("anim_per_slide_avg") or 0) >= 4]
    ok.sort(key=lambda r: -(r.get("anim_per_slide_avg") or 0))
    print(f"动画密度 ≥4 的模板：{len(ok)} 个（按密度排序）")
    picks = ok[:14]
    out = []
    for r in picks:
        path = os.path.join(RES, r["file"])
        if not os.path.exists(path):
            continue
        recipes = []
        try:
            with zipfile.ZipFile(path) as zf:
                slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                                key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
                for part in slides[:12]:
                    xml = zf.read(part).decode("utf-8", errors="ignore")
                    rec = page_recipe(timing_rows(xml))
                    if rec.get("count"):
                        recipes.append(rec)
        except Exception as exc:  # noqa: BLE001
            print(f"  ERR {r['file']}: {exc}")
            continue
        if not recipes:
            continue
        avg_groups = sum(x["click_groups"] for x in recipes) / len(recipes)
        avg_with = sum(x["with_prev"] for x in recipes) / len(recipes)
        avg_after = sum(x["after_prev"] for x in recipes) / len(recipes)
        all_presets = Counter(p for x in recipes for p in x["presets"])
        max_delays = [x["max_delay"] for x in recipes if x["max_delay"]]
        row = {
            "file": r["file"], "group": r.get("group"), "pages": r.get("slides"),
            "anim_per_page": r.get("anim_per_slide_avg"),
            "avg_click_groups": round(avg_groups, 1),
            "avg_with": round(avg_with, 1),
            "avg_after": round(avg_after, 1),
            "avg_group_size": round(sum(sum(x["group_sizes"]) / len(x["group_sizes"])
                                        for x in recipes) / len(recipes), 1),
            "max_delay_median": sorted(max_delays)[len(max_delays) // 2] if max_delays else 0,
            "presets_top": [f"{k}×{v}" for k, v in all_presets.most_common(5)],
            "median_dur": sorted(x["median_dur"] for x in recipes if x["median_dur"])[:5],
        }
        out.append(row)
        print(f"\n{row['file'][:52]:52s} {row['anim_per_page']:>5} 条/页")
        print(f"   click 组 {row['avg_click_groups']} | with {row['avg_with']} | after {row['avg_after']}"
              f" | 组内均 {row['avg_group_size']} 条 | 最大 delay {row['max_delay_median']}ms")
        print(f"   预设 {row['presets_top']}")

    with open(os.path.join(S, "timing_recipes.json"), "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print(f"\n[written] {os.path.join(S, 'timing_recipes.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())