"""UI/布局深一层：区域带、列栅格、图文关系、装饰元素、垂直留白节奏。

对每页（归一化坐标）统计：
  bands      五个横向区域里各有多少形状与覆盖（判断页面如何分带）
  columns    x 方向的聚类峰值 → 1/2/3/4/更多列
  image     图片的面积占比与位置（全屏 / 左侧 / 右侧 / 居中 / 无图）
  decor      细长条、小色块、圆点等装饰元素的数量与形态
  gaps       垂直方向相邻元素之间的空白分布（留白节奏）
"""
from __future__ import annotations

import json
import os
import re
import statistics
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RES = os.path.join(os.path.dirname(ROOT), "Resource")
S = os.path.join(ROOT, "samples")
EMU = 914400

BANDS = [("top", 0.0, 0.2), ("upper", 0.2, 0.45), ("middle", 0.45, 0.7),
         ("lower", 0.7, 0.9), ("bottom", 0.9, 1.0)]


def shapes_of(xml: str, W: float, H: float) -> list[dict]:
    out = []
    # 按出现顺序扫 sp / pic，各自取最近的 xfrm
    for m in re.finditer(r"<p:(sp|pic|graphicFrame)>.*?(?=<p:(?:sp|pic|graphicFrame)>|</p:spTree>)",
                         xml, re.S):
        kind = m.group(1)
        block = m.group(0)
        xf = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"/>', block)
        if not xf:
            continue
        x, y, cx, cy = (int(v) / EMU for v in xf.groups())
        if cx <= 0 or cy <= 0:
            continue
        out.append({"kind": kind, "x": x / W, "y": y / H, "w": cx / W, "h": cy / H})
    return out


def page_ui(xml: str, W: float, H: float) -> dict:
    shp = shapes_of(xml, W, H)
    if not shp:
        return {}
    full_bg = [s for s in shp if s["w"] > 0.9 and s["h"] > 0.9]
    fg = [s for s in shp if not (s["w"] > 0.9 and s["h"] > 0.9)]

    bands = {}
    for name, lo, hi in BANDS:
        inb = [s for s in fg if lo <= s["y"] < hi]
        area = sum(s["w"] * s["h"] for s in inb)
        bands[name] = {"n": len(inb), "area": round(area, 3)}

    # 列：把左边界量化到 5%，取出现 >=2 次的 x
    xs = Counter(round(s["x"] * 20) / 20 for s in fg)
    cols = [k for k, n in xs.items() if n >= 2]
    if cols:
        cols.sort()
        # 合并相邻（间距 < 0.08 视为同一列）
        merged = [cols[0]]
        for c in cols[1:]:
            if c - merged[-1] > 0.08:
                merged.append(c)
        n_cols = len(merged)
    else:
        n_cols = 1

    # 图片
    pics = [s for s in fg if s["kind"] == "pic"]
    pic_area = round(sum(s["w"] * s["h"] for s in pics), 3)
    if any(s["w"] > 0.85 and s["h"] > 0.7 for s in pics):
        img_pos = "全屏"
    elif pics:
        first = pics[0]
        centre = first["x"] + first["w"] / 2
        if centre < 0.42:
            img_pos = "左侧"
        elif centre > 0.58:
            img_pos = "右侧"
        elif first["y"] + first["h"] / 2 < 0.4:
            img_pos = "上方"
        else:
            img_pos = "居中"
    else:
        img_pos = "无图"

    # 装饰：细长条 / 小色块 / 圆点
    bars = [s for s in fg if s["w"] / max(s["h"], 0.001) > 8 and s["h"] < 0.04]
    chips = [s for s in fg if 0.04 <= s["w"] <= 0.3 and 0.02 <= s["h"] <= 0.12
             and s["w"] * s["h"] < 0.03]
    dots = [s for s in fg if abs(s["w"] - s["h"]) < 0.02 and s["w"] < 0.06]

    # 垂直留白：把元素按 y 排序，算相邻间隙
    spans = sorted((s["y"], s["y"] + s["h"]) for s in fg)
    gaps = []
    cur = 0.0
    for top, bot in spans:
        if top > cur:
            gaps.append(round(top - cur, 3))
        cur = max(cur, bot)
    if cur < 1.0:
        gaps.append(round(1 - cur, 3))

    return {
        "fg": len(fg), "bg": len(full_bg),
        "bands": {k: v["n"] for k, v in bands.items()},
        "band_area": {k: v["area"] for k, v in bands.items()},
        "cols": n_cols,
        "pics": len(pics), "pic_area": pic_area, "img_pos": img_pos,
        "bars": len(bars), "chips": len(chips), "dots": len(dots),
        "decor_total": len(bars) + len(chips) + len(dots),
        "gaps": gaps,
    }


def main() -> int:
    plan = json.load(open(os.path.join(S, "resource_plan.json"), encoding="utf-8"))
    picks = []
    for row in plan["plan"]:
        for rel in row["picked"][:2]:
            picks.append(rel)
    print(f"[scan] {len(picks)} 个文件")
    page_rows = []
    file_rows = []
    for i, rel in enumerate(picks, start=1):
        path = os.path.join(RES, rel)
        if not os.path.exists(path):
            continue
        try:
            with zipfile.ZipFile(path) as zf:
                pres = zf.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
                m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
                W, H = (int(m.group(1)) / EMU, int(m.group(2)) / EMU) if m else (13.333, 7.5)
                slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                                key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
                rows = []
                for part in slides[1:6]:   # 跳过封面，看内容页
                    xml = zf.read(part).decode("utf-8", errors="ignore")
                    r = page_ui(xml, W, H)
                    if r:
                        rows.append(r)
        except Exception as exc:  # noqa: BLE001
            print(f"  ERR {rel}: {exc}")
            continue
        if not rows:
            continue
        page_rows.extend(rows)
        file_rows.append({"file": rel, "pages": len(rows)})
        if i % 10 == 0:
            print(f"  ... {i}/{len(picks)}", flush=True)

    with open(os.path.join(S, "ui_layout.json"), "w", encoding="utf-8") as fh:
        json.dump({"files": file_rows, "pages": page_rows}, fh, ensure_ascii=False, indent=1)

    n = len(page_rows)
    print(f"\n=== 内容页样本 {n} 页 ===")

    def agg(fn, label, fmt="{:.2f}"):
        vals = [fn(r) for r in page_rows if fn(r) is not None]
        if vals:
            print(f"  {label:22s} 中位 {fmt.format(statistics.median(vals))}")

    agg(lambda r: r["fg"], "前景形状数", "{:.1f}")
    agg(lambda r: r["decor_total"], "装饰元素数", "{:.1f}")
    print("  区域带形状数:", dict(Counter(k for r in page_rows for k in r["bands"]
                                    for _ in range(r["bands"][k])).most_common()))
    print("  列数分布:", dict(Counter(r["cols"] for r in page_rows).most_common()))
    print("  图片位置:", dict(Counter(r["img_pos"] for r in page_rows).most_common()))
    agg(lambda r: r["pic_area"], "图片面积占比")
    print("  装饰构成:", {"bars": sum(r["bars"] for r in page_rows),
                        "chips": sum(r["chips"] for r in page_rows),
                        "dots": sum(r["dots"] for r in page_rows)})
    all_gaps = [g for r in page_rows for g in r["gaps"]]
    if all_gaps:
        print(f"  垂直留白间距 中位 {statistics.median(all_gaps):.3f} "
              f"（≈{statistics.median(all_gaps) * 7.5:.2f}in）")
    print(f"[written] {os.path.join(S, 'ui_layout.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())