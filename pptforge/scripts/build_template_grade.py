"""按模板规则从零生成一份 deck 并对标：
  · 浅底 + 黑字（91/91 模板如此），accent 从色相派生（S=0.65, V=0.77, 对比背景≈4.7）
  · 边距 0.5in（落在 0.25in 网格上），形状按英寸整数/半英寸对齐
  · 每页 15 个前景形状（13 内容 + 装饰），动画 16-28 条
"""
from __future__ import annotations

import colorsys
import json
import os
import statistics
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import validate_deck
from mcp_ppt.layout import layout_report

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples", "template-grade.pptx")


def _rel(hexv: str) -> float:
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    hexv = hexv.lstrip("#")
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def accent_from_hue(hue: float = 205, sat: float = 0.72, val: float = 0.80,
                    min_contrast: float = 4.5) -> str:
    """模板实测 accent 与白底对比≈4.7；这里按对比度下限往下压明度，保证可读。"""
    best = None
    for v in (val, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5):
        r, g, b = colorsys.hsv_to_rgb(hue / 360, sat, v)
        hexv = "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
        c = (1.05) / (_rel(hexv) + 0.05)
        best = hexv
        if c >= min_contrast:
            return hexv
    return best


ACCENT = accent_from_hue()
ACCENT2 = accent_from_hue(168, 0.6, 0.80)   # 辅色：青绿，同样按对比度下限

MODEL = {
    "theme": {"preset": "light",
              "colors": {"accent1": ACCENT, "accent2": ACCENT2},
              "fonts": {"major": "Microsoft YaHei", "minor": "Microsoft YaHei"}},
    "footer": "template-grade deck",
    "defaults": {"transition": {"effect": "fade", "duration": 0.8}},
    "slides": [
        {"layout": "title", "title": "Template Grade", "subtitle": "按 91 个模板的中位参数生成",
         "meta": ["2026", "对标验证"], "title_size": 60},
        {"layout": "metrics", "kicker": "CONTENTS", "title": "四个部分",
         "cards": [{"value": "01", "label": "栅格与边距"},
                   {"value": "02", "label": "字号层级"},
                   {"value": "03", "label": "配色公式"},
                   {"value": "04", "label": "动效节奏"}]},
        {"layout": "section", "index": 1, "title": "栅格与边距", "subtitle": "0.5in 基线，0.25in 网格"},
        {"layout": "bullets", "kicker": "01 · 栅格", "title": "模板不用 12 列",
         "bullets": [
             {"text": "33% 的形状落在 0.25in 网格", "level": 0},
             {"text": "0.5in 与 1.0in 网格的吻合度同样是 33%，说明是按英寸手工对齐，不是按列宽切分。", "level": 1},
             {"text": "12 列栅格只有 17%", "level": 0},
             {"text": "Web 那套列栅格在 PPT 里不是主流做法，直接用英寸更稳。", "level": 1},
         ]},
        {"layout": "bullets", "kicker": "01 · 边距", "title": "上紧下松",
         "bullets": [
             {"text": "上 0.16in、左右 0.48in、下 0.60in", "level": 0},
             {"text": "上边距只有下边距的三分之一，底部留白用来放页脚与呼吸。", "level": 1},
             {"text": "内容覆盖 71%", "level": 0},
             {"text": "留白约 29%，页面是填满的，只是填得有秩序。", "level": 1},
         ]},
        {"layout": "section", "index": 2, "title": "字号层级", "subtitle": "最大字是正文的 5 倍"},
        {"layout": "bullets", "kicker": "02 · 字号", "title": "11 档，但只有三档是主角",
         "bullets": [
             {"text": "封面 54–72 / 页标题 28–40", "level": 0},
             {"text": "小节 20–24，正文 14–18，说明 12，页脚与图注 10。", "level": 1},
             {"text": "层级靠比例拉开", "level": 0},
             {"text": "最大字是正文的 5 倍（我原来只有 1.7 倍，看着就没有主次）。", "level": 1},
         ], "columns": 1},
        {"layout": "metrics", "kicker": "02 · 对比", "title": "三个关键比例",
         "cards": [{"value": "5×", "label": "最大字 ÷ 正文"},
                   {"value": "71%", "label": "内容覆盖率"},
                   {"value": "3 条", "label": "对齐基线"}]},
        {"layout": "section", "index": 3, "title": "配色公式", "subtitle": "浅底 + 中饱和 accent"},
        {"layout": "bullets", "kicker": "03 · 配色", "title": "91 个主题全是浅底",
         "bullets": [
             {"text": "白底黑字，对比度 21:1", "level": 0},
             {"text": "模板默认不是深色版：背景明度 1.00、文字明度 0.00，纯黑白对比。", "level": 1},
             {"text": "accent 参数固定", "level": 0},
             {"text": "饱和度 0.65、明度 0.77、与背景对比 4.7——刚好过 WCAG AA 的 4.5。", "level": 1},
         ]},
        {"layout": "quote", "quote": "设计感来自约束：对齐、间距、层级、留白。",
         "source": "2dogppt · 排版设计"},
        {"layout": "section", "index": 4, "title": "动效节奏", "subtitle": "1 click + delay 瀑布"},
        {"layout": "bullets", "kicker": "04 · 动效", "title": "高动画模板怎么排",
         "bullets": [
             {"text": "每页 15–34 条动画，click 组 0–1 个", "level": 0},
             {"text": "其余全是 withEffect，靠 delay 瀑布式连播，观众不需要点 20 次。", "level": 1},
             {"text": "一页混用 4–6 种预设", "level": 0},
             {"text": "标题 wipe、正文 ascend、图片 faded-zoom、装饰 fade，而不是全篇一种效果。", "level": 1},
         ], "columns": 2},
        {"layout": "table", "kicker": "04 · 参数", "title": "对标参数表",
         "rows": [["指标", "模板中位", "本页"],
                  ["前景形状/页", "15.0", "13–15"],
                  ["装饰/页", "7.0", "10"],
                  ["动画/页", "15–34", "16–28"],
                  ["上边距", "0.16in", "0.16in"],
                  ["最大/正文", "5.0×", "5.0×"]],
         "col_widths": [3.0, 2.6, 2.6]},
        {"layout": "section", "index": 5, "title": "谢谢", "subtitle": "THE END"},
    ],
}

ENTER, DUR, STAG = "wipe", 0.5, 0.1
DECOR = [
    {"target": ["decor-dot1", "decor-dot2", "decor-dot3"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2", "decor-chip3"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": "footer-rule", "effect": "wipe", "trigger": "with_previous",
     "duration": 0.4, "stagger": 0.1},
    {"target": "page-num", "effect": "fade", "trigger": "with_previous",
     "duration": 0.3, "stagger": 0.1},
]


def anims_for(spec: dict) -> list[dict]:
    layout = spec.get("layout")
    if layout == "title":
        return [
            {"target": "accent-bar", "effect": "wipe", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": 0.6, "stagger": 0.1},
            {"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.15},
            {"target": "meta", "effect": "fade", "trigger": "with_previous",
             "duration": 0.4, "stagger": 0.2},
        ]
    if layout == "section":
        return [
            {"target": "index", "effect": "wipe", "trigger": "click", "duration": 0.5},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.12},
            {"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.15},
        ] + DECOR
    if layout == "metrics":
        n = len(spec.get("cards") or [])
        out = [{"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
               {"target": "title", "effect": "fly", "trigger": "with_previous",
                "duration": 0.5, "stagger": 0.1},
               {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
                "duration": 0.4, "stagger": 0.12}]
        out.append({"target": [f"card{i}" for i in range(1, n + 1)], "effect": "faded-zoom",
                    "trigger": "with_previous", "duration": 0.5, "stagger": 0.1})
        out.append({"target": [f"cardValue{i}" for i in range(1, n + 1)], "effect": "fade",
                    "trigger": "with_previous", "duration": 0.4, "stagger": 0.08})
        out.append({"target": [f"cardLabel{i}" for i in range(1, n + 1)], "effect": "fade",
                    "trigger": "with_previous", "duration": 0.4, "stagger": 0.08})
        return out + DECOR
    if layout == "table":
        return [{"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
                {"target": "title", "effect": "fly", "trigger": "with_previous",
                 "duration": 0.5, "stagger": 0.1},
                {"target": "table", "effect": "faded-zoom", "trigger": "with_previous",
                 "duration": 0.6, "stagger": 0.1}] + DECOR
    if layout == "quote":
        return [{"target": "mark", "effect": "fade", "trigger": "click", "duration": 0.4},
                {"target": "quote", "effect": "ascend", "trigger": "with_previous",
                 "duration": 0.6, "stagger": 0.15},
                {"target": "source", "effect": "fade", "trigger": "with_previous",
                 "duration": 0.4, "stagger": 0.2}]
    two_col = int(spec.get("columns") or 1) >= 2
    targets = ["bullets-1", "bullets-2"] if two_col else ["bullets"]
    return [
        {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": "fly", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.1},
        {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.4, "stagger": 0.12},
        {"target": targets, "effect": "ascend", "trigger": "with_previous",
         "duration": 0.5, "stagger": 0.2, "build": "paragraph"},
    ] + DECOR


MODEL["slides"] = [dict(s, animations=anims_for(s)) for s in MODEL["slides"]]

report = build_deck(MODEL, OUT)
print("[build]", report["slide_count"], "warnings:", report["warnings"][:3])
print("[validate]", validate_deck(OUT)["stats"], validate_deck(OUT)["problems"])
print("[layout] overflow:", layout_report(OUT).get("issue_count"))

# 对标指标
import re
import zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from analyze_ui_layout import page_ui  # noqa: E402
from analyze_templates import analyze  # noqa: E402
from analyze_grid_palette import contrast, rel_lum  # noqa: E402

rows = []
with zipfile.ZipFile(OUT) as zf:
    for n in sorted([x for x in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", x)]):
        r = page_ui(zf.read(n).decode("utf-8"), 13.333, 7.5)
        if r:
            rows.append(r)
a = analyze(OUT)
from analyze_layout_style import analyze as layout_analysis  # noqa: E402
lm = layout_analysis(OUT)
theme = json.load(open(os.path.join(ROOT, "samples", "layout_style.json"), encoding="utf-8"))
good = [r for r in theme if r.get("margin_left") is not None]


def med(k):
    v = [r[k] for r in good if r.get(k) is not None]
    return round(statistics.median(v), 3) if v else None


print("\n=== 对标 ===")
print(f"{'指标':18s} {'模板':>8s} {'本deck':>8s}")
for label, tpl, mine in [
    ("前景形状/页", med("fg_shapes"), statistics.median([r["fg"] for r in rows])),
    ("装饰/页", 7.0, statistics.median([r["decor_total"] for r in rows])),
    ("上边距(归一)", med("margin_top"), lm.get("margin_top")),
    ("左右边距", med("margin_left"), lm.get("margin_left")),
    ("内容覆盖", med("coverage"), lm.get("coverage")),
    ("字号档数", med("size_levels"), lm.get("size_levels")),
    ("最大/正文", med("size_ratio"), lm.get("size_ratio")),
    ("动画/页", 15, round(a["animations_total"] / max(a["slides"], 1), 1)),
]:
    print(f"{label:18s} {str(tpl):>8s} {str(mine):>8s}")
print("配色: bg=%s fg=%s accent=%s 对比 %.1f" % (
    lm["colors"]["lt1"], lm["colors"]["dk1"], lm["colors"]["accent1"],
    contrast(lm["colors"]["accent1"], lm["colors"]["lt1"])))
print("[done]", OUT)
