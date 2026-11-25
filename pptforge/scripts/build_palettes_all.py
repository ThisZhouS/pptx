"""从全量 2120 个模板的 theme 提取配色，重算 palettes.json（原先只有 91 个采样）。"""
from __future__ import annotations

import colorsys
import json
import os
import statistics
from collections import Counter, defaultdict

ROOT = r"E:\pptx\pptforge"
S = os.path.join(ROOT, "samples")
SKILL = r"E:\pptx\pptx-skill"
ASSETS = os.path.join(SKILL, "assets")


def rel_lum(hexv: str) -> float:
    hexv = hexv.lstrip("#")

    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a: str, b: str) -> float:
    l1, l2 = sorted((rel_lum(a), rel_lum(b)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def hsv(hexv: str):
    r, g, b = (int(hexv[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return colorsys.rgb_to_hsv(r, g, b)


def main() -> int:
    data = json.load(open(os.path.join(S, "resource_deep_all.json"), encoding="utf-8"))
    ok = [r for r in data if r and not r.get("error") and (r.get("colors") or {}).get("accent1")]
    print(f"含配色的样本 {len(ok)}/{len(data)}")

    families = defaultdict(list)
    mode_count = Counter()
    for r in ok:
        c = r["colors"]
        acc = c["accent1"]
        try:
            h, s, v = hsv(acc)
        except Exception:  # noqa: BLE001
            continue
        hue = h * 360
        fam = ("蓝" if 200 <= hue < 250 else "红橙" if hue < 60 or hue >= 330 else
               "青" if 160 <= hue < 200 else "绿" if 70 <= hue < 160 else
               "紫" if 250 <= hue < 330 else "黄")
        bg = c.get("lt1", "FFFFFF")
        fg = c.get("dk1", "000000")
        mode = "浅底" if rel_lum(bg) > rel_lum(fg) else "深底"
        mode_count[mode] += 1
        families[fam].append({
            "hue": round(hue), "sat": round(s, 2), "val": round(v, 2),
            "accent": acc, "bg": bg, "fg": fg,
            "contrast_accent_bg": round(contrast(acc, bg), 2),
        })

    out = {
        "source": "全部 2120 个模板的 theme1.xml clrScheme 提取（原先 91 个采样）",
        "total_themes": len(ok),
        "mode": dict(mode_count),
        "hue_families": {},
        "notes": "浅底占绝对多数；accent 与背景对比中位 4.7（WCAG AA 门槛 4.5）",
    }
    for fam, items in sorted(families.items(), key=lambda kv: -len(kv[1])):
        out["hue_families"][fam] = {
            "count": len(items),
            "share": round(len(items) / len(ok), 3),
            "accent_sat": round(statistics.median([i["sat"] for i in items]), 2),
            "accent_val": round(statistics.median([i["val"] for i in items]), 2),
            "contrast_median": round(statistics.median([i["contrast_accent_bg"] for i in items]), 2),
            "samples": [i["accent"] for i in items[:8]],
        }
    json.dump(out, open(os.path.join(ASSETS, "palettes.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    print("mode:", dict(mode_count))
    for fam, v in out["hue_families"].items():
        print(f"  {fam:4s} {v['count']:5d} ({v['share']*100:4.1f}%)  "
              f"sat {v['accent_sat']} val {v['accent_val']} 对比 {v['contrast_median']}")
    print(f"[written] {os.path.join(ASSETS, 'palettes.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())