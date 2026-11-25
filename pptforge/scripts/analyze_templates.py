"""分析"人类设计师"的带动画模板：切换类型/时长、对象动画数量与参数、每页元素密度。

用来给引擎的默认参数找参照（而不是凭感觉定 0.4/0.5/1.2 秒）。
"""
import json
import os
import re
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(os.path.dirname(HERE), "samples", "templates")

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def analyze(path: str) -> dict:
    info = {"file": os.path.basename(path),
            "size_mb": round(os.path.getsize(path) / 1048576, 2)}
    with zipfile.ZipFile(path) as zf:
        slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        info["slides"] = len(slides)
        transitions = Counter()
        anim_per_slide = []
        durations = Counter()
        presets = Counter()
        triggers = Counter()
        shapes_per_slide = []
        for part in slides:
            xml = zf.read(part).decode("utf-8", errors="ignore")
            # 切换
            m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>", xml, re.S)
            if m:
                effect = re.findall(r"<(p\d*:\w+)(?:\s[^>]*)?/>", m.group(1))
                dur = re.search(r'p14:dur="(\d+)"', m.group(0))
                transitions[(effect[0] if effect else "none", dur.group(1) if dur else None)] += 1
            # 对象动画
            pid = re.findall(r'presetID="(\d+)" presetClass="(\w+)" presetSubtype="(\d+)"', xml)
            anim_per_slide.append(len(pid))
            for p1, c1, s1 in pid:
                presets[(c1, p1)] += 1
            for node in re.findall(r'nodeType="(\w+)"', xml):
                triggers[node] += 1
            for d in re.findall(r'<p:cTn [^>]*dur="(\d+)"', xml):
                if d != "1":
                    durations[int(d)] += 1
            shapes_per_slide.append(len(re.findall(r'<p:cNvPr id="\d+"', xml)) - 1)
        info["transitions"] = {f"{k[0]}|dur={k[1]}": v for k, v in transitions.items()}
        info["animations_per_slide"] = anim_per_slide
        info["animations_total"] = sum(anim_per_slide)
        info["preset_top"] = [f"{c}#{p}" for (c, p), _ in presets.most_common(10)]
        info["preset_counts"] = [v for _, v in presets.most_common(10)]
        info["node_types"] = dict(triggers)
        info["durations_ms_top"] = durations.most_common(8)
        info["shapes_per_slide"] = shapes_per_slide
        info["avg_shapes"] = round(sum(shapes_per_slide) / max(len(shapes_per_slide), 1), 1)
    return info


def main() -> int:
    if not os.path.isdir(TPL):
        print("没有模板目录：", TPL)
        return 1
    out = []
    for name in sorted(os.listdir(TPL)):
        if not name.endswith(".pptx"):
            continue
        r = analyze(os.path.join(TPL, name))
        out.append(r)
        print("=" * 72)
        print(f"{r['file']}  {r['size_mb']}MB  {r['slides']} 页  平均每页 {r['avg_shapes']} 个形状")
        print("  切换:", r["transitions"])
        print("  动画/页:", r["animations_per_slide"], "共", r["animations_total"])
        print("  预设 top:", list(zip(r["preset_top"], r["preset_counts"])))
        print("  触发:", r["node_types"])
        print("  时长(ms) top:", r["durations_ms_top"])
    path = os.path.join(os.path.dirname(HERE), "samples", "template_analysis.json")
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print(f"[written] {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
