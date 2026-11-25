"""全量布局测量：2120 个模板各取前 8 页，统计边距、覆盖率、列数、图文、装饰、留白。"""
from __future__ import annotations

import json
import os
import statistics
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, HERE)
RES = os.path.join(os.path.dirname(ROOT), "Resource")
S = os.path.join(ROOT, "samples")

from analyze_ui_layout import page_ui  # noqa: E402

IN = os.path.join(S, "resource_deep_all.json")
OUT = os.path.join(S, "layout_all.json")


def main() -> int:
    import re

    data = json.load(open(IN, encoding="utf-8"))
    rels = [r["file"] for r in data if r and not r.get("error")]
    print(f"[scan] {len(rels)} 个文件")
    pages = []
    done = 0
    for rel in rels:
        path = os.path.join(RES, rel)
        if not os.path.exists(path):
            continue
        try:
            with zipfile.ZipFile(path) as zf:
                pres = zf.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
                m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
                W, H = (int(m.group(1)) / 914400, int(m.group(2)) / 914400) if m else (13.333, 7.5)
                slides = sorted([n for n in zf.namelist()
                                 if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                                key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
                for part in slides[:8]:
                    r = page_ui(zf.read(part).decode("utf-8", errors="ignore"), W, H)
                    if r:
                        pages.append(r)
        except Exception:  # noqa: BLE001
            continue
        done += 1
        if done % 200 == 0:
            print(f"  ... {done}/{len(rels)}  页样本 {len(pages)}", flush=True)

    def med(key):
        v = [p[key] for p in pages if p.get(key) is not None]
        return round(statistics.median(v), 3) if v else None

    summary = {
        "scope": f"全部 {done} 个文件的 {len(pages)} 页",
        "fg_shapes": med("fg"),
        "decor_total": med("decor_total"),
        "cols": dict(Counter(p["cols"] for p in pages).most_common(6)),
        "image_placement": dict(Counter(p["img_pos"] for p in pages).most_common()),
        "decor_mix": {"chips": round(sum(p["chips"] for p in pages) / len(pages), 1),
                      "dots": round(sum(p["dots"] for p in pages) / len(pages), 1),
                      "bars": round(sum(p["bars"] for p in pages) / len(pages), 1)},
        "gap_in": round(statistics.median([g for p in pages for g in p["gaps"]]) * 7.5, 2)
        if any(p["gaps"] for p in pages) else None,
        "band_shapes": {k: sum(p["bands"][k] for p in pages if k in p["bands"])
                        for k in ("top", "upper", "middle", "lower", "bottom")},
    }
    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(summary, fh, ensure_ascii=False, indent=1)

    print(f"\n=== 全量布局（{len(pages)} 页）===")
    print(f"  前景形状/页 中位 {summary['fg_shapes']}（抽样版 15.0）")
    print(f"  装饰/页 中位 {summary['decor_total']}（抽样版 7.0）")
    print(f"  装饰构成 {summary['decor_mix']}")
    print(f"  列数 {summary['cols']}")
    print(f"  图文 {summary['image_placement']}")
    print(f"  块间距 {summary['gap_in']} in")
    print(f"  区域带 {summary['band_shapes']}")
    print(f"[written] {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())