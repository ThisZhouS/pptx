"""冒烟测试：生成一份带回切换 + 对象动画的 demo deck，并读回结构。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt.deck import build_deck, dump_slide_xml
from pptx import Presentation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
os.makedirs(OUT, exist_ok=True)
TARGET = os.path.join(OUT, "smoke.pptx")

MODEL = {
    "theme": "dark",
    "footer": "pptforge demo",
    "defaults": {"transition": {"effect": "fade", "duration": 0.7}},
    "slides": [
        {
            "layout": "title",
            "title": "AI 生成 PPT 与切换动画",
            "subtitle": "从 PowerPoint OOXML 逆向到可调用的 MCP 引擎",
            "meta": ["2026", "pptforge"],
            "transition": {"effect": "morph", "duration": 1.2},
        },
        {
            "layout": "bullets",
            "kicker": "逆向",
            "title": "三层结构：包 / 页 / 时间轴",
            "bullets": [
                "包层：presentation.xml 定义页序，rels 绑定 slide 部件",
                "页层：p:sld → cSld / clrMapOvr / transition / timing",
                "时间轴：tmRoot → mainSeq → click 组 → 效果 par → 行为节点",
                {"text": "切换与动画是两套独立机制，必须分别写入", "level": 1},
            ],
            "transition": {"effect": "cover-left", "duration": 0.6},
            "animations": [
                {"target": "title", "effect": "fade", "trigger": "click", "duration": 0.5},
                {"target": "bullets", "effect": "fly", "trigger": "after_previous",
                 "duration": 0.6, "build": "paragraph"},
            ],
        },
        {
            "layout": "metrics",
            "kicker": "验证",
            "title": "采样覆盖度",
            "cards": [
                {"value": "150", "label": "切换效果（含 p14/p15/p159 扩展）"},
                {"value": "150", "label": "动画预设（entr/emph/exit/path）"},
                {"value": "20", "label": "参数变体（触发/时长/分级/路径）"},
            ],
            "transition": {"effect": "vortex-left", "duration": 0.8},
            "animations": [
                {"target": ["card1", "card2", "card3"], "effect": "zoom", "trigger": "click",
                 "duration": 0.5},
            ],
        },
    ],
}

report = build_deck(MODEL, TARGET)
print(json.dumps(report, ensure_ascii=False, indent=1))

prs = Presentation(TARGET)
print("\n[readback] slides:", len(prs.slides), "size:", prs.slide_width, prs.slide_height)
for i, slide in enumerate(prs.slides, start=1):
    print(f"  slide {i}: shapes={len(slide.shapes)} "
          f"transition={'yes' if b'transition' in slide.element.xml.encode() else 'no'}")

print("\n[slide 2 xml tail]")
xml = dump_slide_xml(TARGET, 2)
print(xml[-2500:])