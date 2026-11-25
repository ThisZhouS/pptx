"""导出三版对比 PNG 并用 PowerPoint 回读切换/动画。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.layout import png_stats

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = r"C:\Users\zs\Desktop\CodeCraft AI  基于大语言模型的自主软件架构优化与重构引擎.pptx"
CASES = [
    ("before", SRC),
    ("after", os.path.join(ROOT, "samples", "codecraft-optimized.pptx")),
    ("motion", os.path.join(ROOT, "samples", "codecraft-motion-only.pptx")),
]

out = {}
for tag, path in CASES:
    render_dir = os.path.join(ROOT, "samples", "render-codecraft", tag)
    stats = png_stats(path, render_dir, 1200)
    pixels = stats.get("pixel_check") or []
    entry = {"file": path, "pages": len(pixels),
             "ink": [p["ink_ratio"] for p in pixels],
             "files": stats.get("files", [])[:2]}
    if comtools.available():
        v = comtools.verify(path)
        if v.get("available"):
            entry["entry_effects"] = sorted({s["entry_effect"] for s in v["slides"]})
            entry["animations"] = sum(len(s["animations"]) for s in v["slides"])
            entry["anim_types"] = sorted({a["type"] for s in v["slides"] for a in s["animations"]})
    out[tag] = entry
    print(f"[{tag}] pages={entry['pages']} ink={entry['ink'][:6]} "
          f"entry={entry.get('entry_effects')} anims={entry.get('animations')}")

with open(os.path.join(ROOT, "samples", "codecraft-compare.json"), "w", encoding="utf-8") as fh:
    json.dump(out, fh, ensure_ascii=False, indent=1)
print("[written] samples/codecraft-compare.json")