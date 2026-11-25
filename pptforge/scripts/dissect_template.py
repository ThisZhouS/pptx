"""把 Resource 模板里出现频次最高的 presetID 翻译成效果名，并拆解一个高动画样本。"""
import json
import os
import re
import sys
import zipfile
from collections import Counter

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
RES = os.path.join(os.path.dirname(ROOT), "Resource")

# 1) presetID → 效果名
an = json.load(open(os.path.join(ROOT, "schema", "animations.json"), encoding="utf-8"))["effects"]
by_pid = {}
for k, v in an.items():
    if v["preset"]["presetClass"] == "entr":
        by_pid.setdefault(v["preset"]["presetID"], (k, v["zh"]))
print("=== 高频 presetID 对照 ===")
for pid in ("22", "10", "2", "42", "53", "14", "16", "47", "23", "12"):
    hit = by_pid.get(pid)
    print(f"  {pid:>3} -> {hit if hit else '?'}")

# 2) 拆解一个高动画样本：每页的动画序列
TARGET = os.path.join(RES, "创赛点击进去—先下载国赛", "200套绝美PPT模板！含多份乡村振兴PPT", "98潮流.pptx")


def detail(path: str, max_pages: int = 4) -> None:
    print(f"\n=== 拆解 {os.path.basename(path)} ===")
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        slides = sorted([n for n in names if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        print(f"  页数 {len(slides)}  版式 {len([n for n in names if 'slideLayout' in n])} "
              f" 媒体 {len([n for n in names if n.startswith('ppt/media/')])}")
        for part in slides[:max_pages]:
            xml = zf.read(part).decode("utf-8", errors="ignore")
            shapes = []
            for m in re.finditer(r'<p:cNvPr id="(\d+)" name="([^"]*)"', xml):
                shapes.append(m.group(2))
            seq = re.findall(
                r'<p:cTn id="\d+" presetID="(\d+)" presetClass="(\w+)" presetSubtype="(\d+)"[^>]*'
                r'(?:nodeType="(\w+)")?[^>]*>.*?<p:cond delay="(\d+)"', xml, re.S)
            durs = re.findall(r'<p:cTn [^>]*dur="(\d+)"', xml)
            print(f"\n  -- {part} 形状 {len(shapes)} 个: {shapes[:8]}")
            print(f"     动画 {len(seq)} 条, 时长分布 {Counter([d for d in durs if d not in ('1','indefinite')]).most_common(3)}")
            for p1, c1, s1, node, delay in seq[:8]:
                print(f"       preset={p1:>2} {c1:5s} sub={s1:>2} node={node:12s} delay={delay}")


# 3) 统计所有抽样的"每页动画数 vs 形状数"关系
data = json.load(open(os.path.join(ROOT, "samples", "resource_analysis.json"), encoding="utf-8"))
with_anim = [d for d in data if d.get("anim_total")]
print(f"\n=== 有动画的模板 {len(with_anim)}/{len(data)} ===")
print(f"  每页动画均值 {sum(d['anim_per_slide_avg'] for d in with_anim) / max(len(with_anim), 1):.1f}")
print(f"  每页形状均值 {sum(d['shapes_avg'] for d in with_anim) / max(len(with_anim), 1):.1f}")
morph_files = [d for d in data if any("morph" in k for k in (d.get("transitions") or {}))]
print(f"\n=== 用 morph 的模板 {len(morph_files)} 个 ===")
for d in morph_files:
    print(f"  {d['file'][:60]:60s} {d['slides']}页 动画{d['anim_total']} {list(d['transitions'])[:1]}")

if os.path.exists(TARGET):
    detail(TARGET)
else:
    print("\n[skip] 目标文件不存在")