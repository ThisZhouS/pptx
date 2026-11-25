"""Morph 关键帧引擎验收：生成序列 → 真机回读 morph/advTm/元素位置 → 渲染。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation

from mcp_ppt import comtools
from mcp_ppt.inspect import inspect_deck, validate_deck
from mcp_ppt.morph import build_morph_deck, decompose, plan

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
TARGET = os.path.join(OUT, "morph-sequence.pptx")

MODEL = {
    "theme": "dark",
    "background": "#0B1020",
    "frame_duration": 1.6,
    "transition": {"effect": "morph", "duration": 1.2},
    "base": {
        "elements": [
            {"name": "title", "kind": "text", "x": 0.9, "y": 0.6, "w": 8.0, "h": 0.7,
             "text": "关键帧叙事：从左上到右下", "size": 26, "bold": True, "color": "#EAF0FF"},
            {"name": "note", "kind": "text", "x": 0.9, "y": 6.7, "w": 11.0, "h": 0.5,
             "text": "morph 1.2s · 每帧停留 1.6s · 自动换片", "size": 12, "color": "#9BABCE"},
        ]
    },
    "elements": [
        {
            "name": "hero",
            "kind": "round-rect",
            "fill": "#4C8DFF",
            "text": "HERO",
            "size": 28,
            "bold": True,
            "color": "#04121F",
            "keyframes": [
                {"x": 1.0, "y": 1.6, "w": 2.4, "h": 1.2},
                {"x": 4.6, "y": 3.0, "w": 4.6, "h": 2.2, "fill": "#6FE3C4", "rot": 8},
                {"x": 8.6, "y": 1.4, "w": 3.4, "h": 1.6, "fill": "#FFD166", "rot": -6},
                {"x": 5.2, "y": 4.6, "w": 2.2, "h": 2.2, "fill": "#B388FF", "rot": 0,
                 "alpha": 0.75},
            ],
        },
        {
            "name": "dot",
            "kind": "ellipse",
            "fill": "#36C2FF",
            "keyframes": [
                {"x": 10.6, "y": 1.6, "w": 1.1, "h": 1.1},
                {"x": 9.2, "y": 4.2, "w": 1.8, "h": 1.8, "fill": "#FF7A45"},
                {"x": 2.2, "y": 4.2, "w": 0.9, "h": 0.9, "alpha": 0.5},
                {"x": 11.2, "y": 5.2, "w": 1.4, "h": 1.4, "fill": "#6FE3C4"},
            ],
        },
    ],
}

p = plan(MODEL)
print("[plan] frames:", p["frames"], "total:", p["total_duration"], "s")
for row in p["timeline"]:
    print(f"   frame {row['frame']} @{row['start']}s +{row['duration']}s "
          f"hero={({k: row['elements']['hero'].get(k) for k in ('x', 'y', 'w', 'h')})}")

report = build_morph_deck(MODEL, TARGET)
print("[build] slides:", len(report["slides"]), "file:", report["file"])
for s in report["slides"]:
    print(f"   frame {s['frame']}: {s['shapes']} -> {s['transition'].get('effect')}")

val = validate_deck(TARGET)
print("[validate]", val["stats"], val["problems"])

deck = inspect_deck(TARGET)
for s in deck["slides"]:
    tr = s["transition"] or {}
    names = [x["name"] for x in s["shapes"]]
    print(f"   slide {s['index']}: transition={tr.get('effects')} shapes={names}")

if comtools.available():
    com = comtools.verify(TARGET)
    for page in com["slides"]:
        print(f"[com] slide {page['index']}: entry={page['entry_effect']} shapes={page['slide_shapes']}")

# 元素位置是否逐帧变化（用 python-pptx 复核）
prs = Presentation(TARGET)
print("[positions]")
for i, slide in enumerate(prs.slides, start=1):
    for shp in slide.shapes:
        if shp.name in ("!!hero", "!!dot"):
            print(f"   slide {i} {shp.name}: "
                  f"({round(shp.left / 914400, 2)}, {round(shp.top / 914400, 2)}) "
                  f"{round(shp.width / 914400, 2)}x{round(shp.height / 914400, 2)}in "
                  f"rot={shp.rotation}")

print("[decompose]", json.dumps(
    {k: v for k, v in decompose(TARGET).items() if k in ("pages", "coded_across_pages")},
    ensure_ascii=False))