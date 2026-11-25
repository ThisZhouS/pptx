"""栅格吻合度 + 配色公式：模板的形状是否落在栅格上，配色是不是从主色派生。

栅格：形状 x/y/w 换算成英寸，看是否落在 0.25in / 0.5in 网格，以及 12 列栅格。
配色：theme 12 色的 HSV、背景与前景的相对亮度对比（WCAG 公式），
      accent1 与背景的关系，判断模板是"深色版"还是"浅色版"、accent 如何派生。
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
EMU = 914400


def rel_lum(hexv: str) -> float:
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    def f(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a: str, b: str) -> float:
    l1, l2 = sorted((rel_lum(a), rel_lum(b)), reverse=True)
    return round((l1 + 0.05) / (l2 + 0.05), 2)


def hsv(hexv: str) -> tuple:
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return colorsys.rgb_to_hsv(r, g, b)


def main() -> int:
    plan = json.load(open(os.path.join(S, "resource_plan.json"), encoding="utf-8"))
    picks = [rel for row in plan["plan"] for rel in row["picked"][:2]]
    grid_hits = Counter()
    total_boxes = 0
    palettes = []
    for i, rel in enumerate(picks, start=1):
        path = os.path.join(RES, rel)
        if not os.path.exists(path):
            continue
        try:
            with zipfile.ZipFile(path) as zf:
                pres = zf.read("ppt/presentation.xml").decode("utf-8", errors="ignore")
                m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
                W, H = (int(m.group(1)) / EMU, int(m.group(2)) / EMU) if m else (13.333, 7.5)

                # 栅格吻合
                slides = [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)][:6]
                for part in slides:
                    xml = zf.read(part).decode("utf-8", errors="ignore")
                    for mm in re.finditer(
                            r'<a:off x="(-?\d+)" y="(-?\d+)"/>\s*<a:ext cx="(\d+)" cy="(\d+)"/>', xml):
                        x, y, cx, cy = (int(v) / EMU for v in mm.groups())
                        if cx <= 0 or cy <= 0:
                            continue
                        total_boxes += 1
                        for step in (0.25, 0.5, 1.0):
                            fx = (x / step) % 1
                            fy = (y / step) % 1
                            if min(fx, 1 - fx) < 0.06 or min(fy, 1 - fy) < 0.06:
                                grid_hits[f"{step}in"] += 1
                        fx = x / W * 12
                        if abs(fx - round(fx)) < 0.03:
                            grid_hits["12col"] += 1

                # 配色
                themes = [n for n in zf.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
                if not themes:
                    continue
                t = zf.read(sorted(themes)[0]).decode("utf-8", errors="ignore")
                cs = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", t, re.S)
                colors = {}
                if cs:
                    for mm in re.finditer(
                            r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*'
                            r'<a:(?:srgbClr val|sysClr[^>]*lastClr)="([0-9A-Fa-f]{6})"', cs.group(1)):
                        colors[mm.group(1)] = mm.group(2).upper()
                if "accent1" not in colors:
                    continue
                bg_key = "lt1" if rel_lum(colors.get("lt1", "FFFFFF")) > rel_lum(colors.get("dk1", "000000")) else "dk1"
                fg_key = "dk1" if bg_key == "lt1" else "lt1"
                bg = colors.get(bg_key, "FFFFFF")
                fg = colors.get(fg_key, "000000")
                palettes.append({
                    "mode": "浅底" if bg_key == "lt1" else "深底",
                    "bg": bg, "fg": fg,
                    "bg_v": round(hsv(bg)[2], 2), "fg_v": round(hsv(fg)[2], 2),
                    "contrast_fg_bg": contrast(fg, bg),
                    "accent1": colors["accent1"],
                    "a_h": round(hsv(colors["accent1"])[0] * 360),
                    "a_s": round(hsv(colors["accent1"])[1], 2),
                    "a_v": round(hsv(colors["accent1"])[2], 2),
                    "contrast_accent_bg": contrast(colors["accent1"], bg),
                    "accent_hue_vs_fg": round(abs(hsv(colors["accent1"])[0] - hsv(fg)[0]) * 360),
                })
        except Exception:  # noqa: BLE001
            continue
        if i % 15 == 0:
            print(f"  ... {i}/{len(picks)}", flush=True)

    print(f"\n=== 栅格吻合（{total_boxes} 个形状）===")
    for k, v in grid_hits.most_common():
        print(f"  {k:8s} {v:6d}  {v / max(total_boxes, 1) * 100:5.1f}%")

    print(f"\n=== 配色（{len(palettes)} 个主题）===")
    print("  模式:", dict(Counter(p["mode"] for p in palettes)))
    for key in ("bg_v", "fg_v", "contrast_fg_bg", "a_s", "a_v", "contrast_accent_bg"):
        vals = [p[key] for p in palettes]
        print(f"  {key:20s} 中位 {statistics.median(vals):>5.2f}   "
              f"范围 {min(vals):.2f}–{max(vals):.2f}")
    print("  accent1 色相:", dict(Counter(
        "蓝" if 200 <= p["a_h"] < 250 else "红橙" if p["a_h"] < 60 or p["a_h"] >= 330 else
        "青" if 160 <= p["a_h"] < 200 else "绿" if 70 <= p["a_h"] < 160 else "紫"
        for p in palettes).most_common()))

    with open(os.path.join(S, "grid_palette.json"), "w", encoding="utf-8") as fh:
        json.dump({"grid": dict(grid_hits), "boxes": total_boxes, "palettes": palettes},
                  fh, ensure_ascii=False, indent=1)
    print(f"[written] {os.path.join(S, 'grid_palette.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())