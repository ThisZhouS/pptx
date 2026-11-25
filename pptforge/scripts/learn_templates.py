"""拆解 Resource/ 下的优质模板：切换、动画、版式系统、配色、字体、密度。

只读 zip 里的 XML 元数据（不解压媒体），因此扫描 2000+ 文件也可控。
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

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def analyze(path: str, max_slides: int = 40) -> dict:
    out = {"file": os.path.relpath(path, RES), "mb": round(os.path.getsize(path) / 1048576, 2)}
    try:
        zf = zipfile.ZipFile(path)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"zip: {exc}"
        return out
    with zf:
        names = zf.namelist()
        slides = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        out["slides"] = len(slides)
        out["layouts"] = len([n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n)])
        out["masters"] = len([n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n)])
        out["media"] = len([n for n in names if n.startswith("ppt/media/")])
        out["charts"] = len([n for n in names if n.startswith("ppt/charts/")])
        out["notes"] = len([n for n in names if n.startswith("ppt/notesSlides/")])

        # 配色与字体
        colors, fonts = {}, Counter()
        themes = [n for n in names if re.match(r"ppt/theme/theme\d+\.xml$", n)]
        if themes:
            t = zf.read(sorted(themes)[0]).decode("utf-8", errors="ignore")
            cs = re.search(r"<a:clrScheme[^>]*>(.*?)</a:clrScheme>", t, re.S)
            if cs:
                for m in re.finditer(r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(?:srgbClr val|sysClr[^>]*lastClr)="([0-9A-Fa-f]{6})"', cs.group(1)):
                    colors[m.group(1)] = m.group(2).upper()
            for m in re.finditer(r'<a:(majorFont|minorFont)><a:latin typeface="([^"]*)"', t):
                fonts[f"{m.group(1)}:{m.group(2)}"] += 1
        out["colors"] = colors
        out["fonts"] = [k for k, _ in fonts.most_common(4)]

        transitions = Counter()
        presets = Counter()
        durations = Counter()
        anim_per_slide = []
        shapes_per_slide = []
        sizes = Counter()
        for part in slides[:max_slides]:
            xml = zf.read(part).decode("utf-8", errors="ignore")
            m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>", xml, re.S)
            if m:
                eff = re.findall(r"<(p\d*:\w+)(?:\s[^>]*)?/>", m.group(1))
                dur = re.search(r'p14:dur="(\d+)"', m.group(0))
                transitions[(eff[0] if eff else "none", int(dur.group(1)) if dur else None)] += 1
            else:
                transitions[("none", None)] += 1
            pid = re.findall(r'presetID="(\d+)" presetClass="(\w+)"', xml)
            anim_per_slide.append(len(pid))
            for p1, c1 in pid:
                presets[f"{c1}#{p1}"] += 1
            for d in re.findall(r'<p:cTn [^>]*dur="(\d+)"', xml):
                if d not in ("1", "indefinite"):
                    durations[int(d)] += 1
            for s1 in re.findall(r'<a:rPr[^>]*sz="(\d+)"', xml):
                sizes[int(s1) // 100] += 1
            shapes_per_slide.append(len(re.findall(r'<p:cNvPr id="\d+"', xml)) - 1)
    out["transitions"] = {f"{k[0]}|{k[1]}": v for k, v in transitions.most_common(6)}
    out["anim_total"] = sum(anim_per_slide)
    out["anim_per_slide_avg"] = round(sum(anim_per_slide) / max(len(anim_per_slide), 1), 1)
    out["presets_top"] = [f"{k}×{v}" for k, v in presets.most_common(6)]
    out["durations_top"] = durations.most_common(5)
    out["sizes_top"] = [f"{s}pt×{n}" for s, n in sizes.most_common(6)]
    out["shapes_avg"] = round(sum(shapes_per_slide) / max(len(shapes_per_slide), 1), 1)
    return out


def pick_samples(limit_groups: int = 18) -> list[str]:
    """按顶层目录分组，每组取体积适中的若干个，优先"动态"目录。"""
    files = []
    for dirpath, _dirnames, filenames in os.walk(RES):
        for fn in filenames:
            if fn.lower().endswith(".pptx"):
                files.append(os.path.join(dirpath, fn))
    groups: dict[str, list[str]] = {}
    for f in files:
        rel = os.path.relpath(f, RES)
        top = rel.split(os.sep)[0]
        groups.setdefault(top, []).append(f)
    picked: list[str] = []
    # 动态 / 平滑 优先，每个这类目录取 3 个
    for name in sorted(groups):
        if re.search(r"动态|平滑|轮播|切换", name):
            wanted = 3
        elif re.search(r"创赛|商务|国潮|中国风|极简|画册|INS|逻辑", name):
            wanted = 2
        else:
            wanted = 1
        cand = sorted(groups[name], key=lambda p: os.path.getsize(p))
        mid = cand[len(cand) // 2:] + cand[:len(cand) // 2]
        picked.extend(mid[:wanted])
    return picked[:60]


def main() -> int:
    if not os.path.isdir(RES):
        print("找不到 Resource/：", RES)
        return 1
    samples = pick_samples()
    print(f"[scan] 抽样 {len(samples)} 个文件")
    results = []
    for i, path in enumerate(samples, start=1):
        r = analyze(path)
        results.append(r)
        ta = r.get("transitions") or {}
        print(f"  {i:2d}. {r['file'][:58]:58s} {r['slides']:3d}页 "
              f"动画{r['anim_total']:4d} 切换{list(ta)[:1]}")
    out = os.path.join(ROOT, "samples", "resource_analysis.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(results, fh, ensure_ascii=False, indent=1)
    print(f"[written] {out}")

    # 汇总
    all_tr = Counter()
    all_pre = Counter()
    all_dur = Counter()
    for r in results:
        for k, v in (r.get("transitions") or {}).items():
            all_tr[k.split("|")[0]] += v
        for item in r.get("presets_top") or []:
            k, v = item.split("×")
            all_pre[k] += int(v)
        for d, n in r.get("durations_top") or []:
            all_dur[int(d)] += n
    print("\n=== 切换类型（页数）===")
    for k, v in all_tr.most_common(8):
        print(f"  {k:24s} {v}")
    print("\n=== 动画预设 top ===")
    for k, v in all_pre.most_common(10):
        print(f"  {k:16s} {v}")
    print("\n=== 动画时长 top(ms) ===")
    for k, v in all_dur.most_common(8):
        print(f"  {k:6d} {v}")
    print("\n=== 每页动画数 & 形状数 ===")
    for r in results[:12]:
        print(f"  {r['file'][:44]:44s} 动画/页≈{r.get('anim_per_slide_avg')} "
              f"形状/页≈{r.get('shapes_avg')} 版式{r.get('layouts')} 媒体{r.get('media')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
