"""按四步流程给 CodeCraft 做一段 morph 叙事示例（新增页，不改原件）。

流程：
  ① 读取与拆解 → ② 元素编码 → ③ 每帧 plan → ④ 起点/中间/终点编写
演示内容：章节转场（大号编号 01 → 02，缩放移位 + 标题交叉变形）。
"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.inspect import validate_deck
from mcp_ppt.morph import build_morph_deck, decompose, plan

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = r"C:\Users\zs\Desktop\CodeCraft AI  基于大语言模型的自主软件架构优化与重构引擎.pptx"
OUT = os.path.join(ROOT, "samples", "codecraft-morph-chapter.pptx")

# ① 拆解源文件，拿配色与章节文案
dec = decompose(SRC)
chapters = []
for page in dec["pages_detail"]:
    texts = [s["text"] for s in page["shapes"] if s["text"].strip()]
    if len(texts) == 2 and texts[1].isdigit() is False and texts[0].isdigit() is False:
        continue
    if any(len(t) == 2 and t.isdigit() for t in texts):
        num = next(t for t in texts if len(t) == 2 and t.isdigit())
        title = next((t for t in texts if not (len(t) == 2 and t.isdigit())), "")
        chapters.append({"slide": page["slide"], "num": num, "title": title})
print("[① 拆解] 章节页:", json.dumps(chapters, ensure_ascii=False))

# ② 元素编码：同一个逻辑元素在每帧都叫 !!num / !!chapter-title / !!rule
CODED = ["num", "chapter-title", "rule"]

# ③ 每帧 plan（4 帧：开场 → 落位 → 换章 → 收束）
MODEL = {
    "theme": {"preset": "dark", "bg": "#0B1020", "fg": "#EAF0FF",
              "accent": "#4C8DFF", "accent_2": "#6FE3C4"},
    "background": "#0B1020",
    "frame_duration": [2.0, 1.2, 1.2, 1.6],
    "transition": {"effect": "morph", "duration": 1.2},
    "base": {
        "elements": [
            {"name": "kicker", "kind": "text", "x": 0.95, "y": 6.6, "w": 6.0, "h": 0.4,
             "text": "CodeCraft AI · 章节转场", "size": 12, "color": "#9BABCE"},
        ]
    },
    "elements": [
        {
            # ④ 起点：满屏居中；中间点：缩到左上；终点：换成 02 再回到中上
            "name": "num",
            "kind": "text",
            "text": "01",
            "size": 220,
            "bold": True,
            "color": "#4C8DFF",
            "keyframes": [
                {"x": 4.4, "y": 1.6, "w": 4.5, "h": 3.6},
                {"x": 0.9, "y": 0.7, "w": 1.5, "h": 1.2, "size": 72, "color": "#6FE3C4"},
                {"x": 0.9, "y": 0.7, "w": 1.5, "h": 1.2, "text": "02", "size": 72,
                 "color": "#FFD166"},
                {"x": 1.0, "y": 1.4, "w": 1.8, "h": 1.4, "text": "02", "size": 84,
                 "color": "#FFD166"},
            ],
        },
        {
            # 标题：从下方淡入 → 移到右侧 → 文字交叉变形
            "name": "chapter-title",
            "kind": "text",
            "text": "核心理念与痛点",
            "size": 46,
            "bold": True,
            "color": "#EAF0FF",
            "keyframes": [
                {"x": 3.2, "y": 5.4, "w": 7.0, "h": 1.2, "alpha": 0.35},
                {"x": 2.9, "y": 2.6, "w": 8.6, "h": 1.2, "alpha": 1},
                {"x": 2.9, "y": 2.6, "w": 8.6, "h": 1.2, "text": "核心技术栈与产品"},
                {"x": 3.2, "y": 3.0, "w": 8.6, "h": 1.4, "text": "核心技术栈与产品"},
            ],
        },
        {
            # 装饰条：随编号一起缩放
            "name": "rule",
            "kind": "rect",
            "fill": "#4C8DFF",
            "keyframes": [
                {"x": 4.4, "y": 5.3, "w": 4.5, "h": 0.06},
                {"x": 0.9, "y": 2.0, "w": 1.5, "h": 0.06, "fill": "#6FE3C4"},
                {"x": 0.9, "y": 2.0, "w": 1.5, "h": 0.06, "fill": "#FFD166"},
                {"x": 1.0, "y": 2.9, "w": 1.9, "h": 0.06, "fill": "#FFD166"},
            ],
        },
    ],
}

p = plan(MODEL)
print("[③ 时间轴]", f"{p['frames']} 帧 / {p['total_duration']}s")
for row in p["timeline"]:
    num = row["elements"]["num"]
    print(f"   frame {row['frame']} @{row['start']}s +{row['duration']}s "
          f"num={num.get('text')} ({num.get('x')},{num.get('y')}) {num.get('size')}pt")

report = build_morph_deck(MODEL, OUT)
print("[④ 编写]", report["file"], "slides:", len(report["slides"]))
for s in report["slides"]:
    print(f"   frame {s['frame']}: {s['shapes']} -> {s['transition'].get('effect')}")

print("[validate]", validate_deck(OUT)["stats"], validate_deck(OUT)["problems"])
if comtools.available():
    com = comtools.verify(OUT)
    print("[com]", [(s["index"], s["entry_effect"], s["slide_shapes"]) for s in com["slides"]])
    from mcp_ppt.layout import png_stats
    print("[render]", json.dumps(
        png_stats(OUT, os.path.join(ROOT, "samples", "render-morph"), 1200).get("pixel_check"),
        ensure_ascii=False))