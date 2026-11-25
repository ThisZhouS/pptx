"""函数级联调：逐个调用 MCP 工具函数，产出一份真实 demo。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import server as S

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
DEMO = os.path.join(OUT, "pptforge-demo.pptx")


def show(title, obj):
    print(f"\n=== {title} ===")
    text = json.dumps(obj, ensure_ascii=False, indent=1)
    print(text if len(text) < 1600 else text[:1600] + " ...")


show("capability", S.capability_report())
show("catalog(涡流)", S.catalog_transitions("涡流"))
show("catalog(路径)", S.catalog_animations("path-"))

sections = [
    {"kicker": "背景", "title": "AI 生成 PPT，动效总是丢",
     "points": ["python-pptx 只能写内容，切换与动画不在它的对象模型里",
                {"text": "结果：文件能开，但放映时没有过渡、没有分条出现", "level": 1},
                "手工补动效的成本，抵消了自动化的收益"]},
    {"kicker": "逆向", "title": "把 PowerPoint 拆成三层",
     "points": ["包层：presentation.xml 排序，rels 绑定 slide 部件",
                "页层：cSld → clrMapOvr → transition → timing 的固定顺序",
                {"text": "时间轴：tmRoot → mainSeq → click 组 → 效果 par → 行为节点", "level": 1},
                "切换与动画是两套机制，必须分别写"]},
    {"kicker": "方法", "title": "用真机采样代替猜 schema",
     "points": ["枚举 COM 类型库：150 个切换、150 个动画预设",
                "让 PowerPoint 自己写出 XML，再逐一比对",
                "p14 / p15 / p159 扩展必须包在 mc:AlternateContent 里"]},
    {"kicker": "结果", "title": "一次采样，换来可复用的目录",
     "cards": [{"value": "150", "label": "切换效果可生成"},
               {"value": "150", "label": "动画预设可生成"},
               {"value": "0", "label": "COM 回读失配"}]},
    {"kicker": "验证", "title": "生成 → 真机回读 → 出图",
     "points": ["结构校验：动画形状引用、XML 顺序",
                "PowerPoint 回读：EntryEffect 与预设逐项对齐",
                "版式体检：用 BoundHeight 判断文本溢出"]},
]
show("build_deck_from_outline", S.build_deck_from_outline(
    title="从 PowerPoint 逆向到 MCP",
    subtitle="让 AI 生成的演示自带切换与动画",
    sections=sections,
    out_path=DEMO,
    theme="midnight",
    motion_style="cinematic",
    footer="pptforge · 2026",
))

show("validate", S.validate(DEMO))
show("inspect#3", S.inspect(DEMO)["slides"][2])
show("layout_check", {k: v for k, v in S.layout_check(DEMO).items() if k != "pages"})
show("render_png", S.render_png(DEMO, os.path.join(OUT, "render"), 1400))
show("export_pdf", S.export_pdf(DEMO))