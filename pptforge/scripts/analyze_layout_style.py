"""布局与风格分析：从模板里量化边距、对齐基线、字号层级、配色规律。

对每个样本取前 8 页，把形状坐标归一化到页面比例，统计：
  边距（左/上/右/下最小空白）、内容覆盖率、左边界聚类（对齐线数量）、
  标题带高度、字号档位与最大/最小字号比、主题 12 色的色相分布与明度。
"""
from __future__ import annotations

import colorsys
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

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU = 914400


def page_metrics(xml: str, W: float, H: float) -> dict:
    """单页的布局指标。"""
    boxes = []
    for m in re.finditer(r"<a:off x=\"(-?\d+)\" y=\"(-?\d+)\"/>\s*<a:ext cx=\"(\d+)\" cy=\"(\d+)\"/>", xml):
        x, y, cx, cy = (int(v) / EMU for v in m.groups())
        if cx <= 0 or cy <= 0:
            continue
        boxes.append((x / W, y / H, cx / W, cy / H))
    if not boxes:
        return {"shapes": 0}
    # 满版形状（背景图/底纹）会把边距压到 0，先剥离再算布局
    bg = [b for b in boxes if b[2] > 0.9 and b[3] > 0.9]
    fg = [b for b in boxes if not (b[2] > 0.9 and b[3] > 0.9)]
    out_extra = {"bg_layers": len(bg), "fg_shapes": len(fg)}
    if len(fg) >= 2:
        boxes = fg
    lefts = [b[0] for b in boxes]
    tops = [b[1] for b in boxes]
    rights = [b[0] + b[2] for b in boxes]
    bottoms = [b[1] + b[3] for b in boxes]
    # 对齐线：左边界按 1% 量化后的众数聚类
    q = Counter(round(v * 100) for v in lefts)
    lines = [k for k, n in q.items() if n >= 2]
    coverage = min(sum(b[2] * b[3] for b in boxes), 1.0)
    return {
        **out_extra,
        "shapes": len(boxes),
        "margin_left": round(max(min(lefts), 0), 3),
        "margin_top": round(max(min(tops), 0), 3),
        "margin_right": round(max(1 - max(rights), 0), 3),
        "margin_bottom": round(max(1 - max(bottoms), 0), 3),
        "coverage": round(coverage, 3),
        "align_lines": len(lines),
        "title_band": round(min(max(bottoms[i] for i in range(len(boxes)) if tops[i] < 0.2), 1.0), 3)
        if any(t < 0.2 for t in tops) else None,
    }


def color_stats(hexv: str) -> dict:
    try:
        r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:  # noqa: BLE001
        return {}
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    return {"hue": round(h * 360), "sat": round(s, 2), "val": round(v, 2)}


def analyze(path: str, max_pages: int = 8) -> dict:
    out = {"file": os.path.relpath(path, RES)}
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        # 页面尺寸
        pres = zf.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
        m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
        W, H = (int(m.group(1)) / EMU, int(m.group(2)) / EMU) if m else (13.333, 7.5)
        out["aspect"] = round(W / H, 2)
        # 主题配色
        themes = [n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n)]
        colors = {}
        if themes:
            t = zf.read(sorted(themes)[0]).decode("utf-8", errors="ignore")
            cs = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", t, re.S)
            if cs:
                for mm in re.finditer(
                        r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*'
                        r'<a:(?:srgbClr val|sysClr[^>]*lastClr)="([0-9A-Fa-f]{6})"', cs.group(1)):
                    colors[mm.group(1)] = mm.group(2).upper()
        out["colors"] = colors
        acc = color_stats(colors.get("accent1", "")) if colors.get("accent1") else {}
        out["accent1_hsv"] = acc
        dk = color_stats(colors.get("dk1", "")) if colors.get("dk1") else {}
        out["dk1_hsv"] = dk
        # 每页指标
        slides = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        pages = []
        sizes = Counter()
        for part in slides[:max_pages]:
            xml = zf.read(part).decode("utf-8", errors="ignore")
            pm = page_metrics(xml, W, H)
            if pm.get("shapes"):
                pages.append(pm)
            for sz in re.findall(r'<a:rPr[^>]*sz="(\d+)"', xml):
                sizes[int(sz) // 100] += 1
            pics = len(re.findall(r"<p:pic>", xml))
            out["_pics"] = out.get("_pics", 0) + pics
            out["_pages_for_pic"] = out.get("_pages_for_pic", 0) + 1
        if not pages:
            return out
        out["pages"] = len(pages)
        for key in ("margin_left", "margin_top", "margin_right", "margin_bottom",
                    "coverage", "align_lines", "bg_layers", "fg_shapes"):
            vals = [p[key] for p in pages if p.get(key) is not None]
            if vals:
                out[key] = round(statistics.median(vals), 3)
        # 字号层级
        if out.get("_pages_for_pic"):
            out["pics_per_page"] = round(out.pop("_pics") / out["_pages_for_pic"], 1)
            out.pop("_pages_for_pic", None)
        if sizes:
            distinct = sorted(sizes)
            out["size_levels"] = len(distinct)
            out["size_top"] = [s for s, _ in sizes.most_common(4)]
            big = max(distinct)
            body = statistics.median([s for s, n in sizes.items() for _ in range(min(n, 20))])
            out["size_ratio"] = round(big / body, 1) if body else None
        return out


def main() -> int:
    deep = json.load(open(os.path.join(S, "resource_deep.json"), encoding="utf-8"))
    ok = [r for r in deep if r and not r.get("error")]
    # 每板块取 2 个，控制总量
    by_group: dict[str, list] = {}
    for r in ok:
        by_group.setdefault(r.get("group") or "?", []).append(r)
    picks = []
    for g, items in by_group.items():
        picks.extend(items[:2])
    print(f"[scan] {len(picks)} 个样本做布局/风格分析")
    results = []
    for i, r in enumerate(picks, start=1):
        path = os.path.join(RES, r["file"])
        if not os.path.exists(path):
            continue
        try:
            info = analyze(path)
        except Exception as exc:  # noqa: BLE001
            info = {"file": r["file"], "error": f"{type(exc).__name__}: {exc}"}
        info["group"] = r.get("group")
        results.append(info)
        if i % 10 == 0:
            print(f"  ... {i}/{len(picks)}", flush=True)
    with open(os.path.join(S, "layout_style.json"), "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)

    good = [r for r in results if r.get("margin_left") is not None]
    print(f"\n有效 {len(good)}/{len(results)}")
    for key in ("margin_left", "margin_top", "margin_right", "margin_bottom",
                "coverage", "align_lines", "bg_layers", "fg_shapes",
                "pics_per_page", "size_levels", "size_ratio"):
        vals = [r[key] for r in good if r.get(key) is not None]
        if vals:
            print(f"  {key:14s} 中位 {statistics.median(vals):>6.2f}  "
                  f"范围 {min(vals):.2f}–{max(vals):.2f}")
    hues = Counter()
    for r in good:
        h = (r.get("accent1_hsv") or {}).get("hue")
        if h is None:
            continue
        band = "红0-30" if h < 30 or h >= 330 else "橙黄30-70" if h < 70 else \
               "绿70-160" if h < 160 else "青160-200" if h < 200 else \
               "蓝200-250" if h < 250 else "紫250-290" if h < 290 else "品红290-330"
        hues[band] += 1
    print("\n=== accent1 色相分布 ===")
    for k, v in hues.most_common():
        print(f"  {k:12s} {v}")
    sats = [(r.get("accent1_hsv") or {}).get("sat") for r in good if r.get("accent1_hsv")]
    vals2 = [(r.get("accent1_hsv") or {}).get("val") for r in good if r.get("accent1_hsv")]
    if sats:
        print(f"=== accent1 饱和度中位 {statistics.median(sats):.2f} / 明度中位 {statistics.median(vals2):.2f}")
    print(f"[written] {os.path.join(S, 'layout_style.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
