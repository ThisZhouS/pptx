"""pptx-skill 生成器：按任务类型与风格选择素材，从大纲生成对标模板的 pptx。

用法：
  python build.py outline.json --out out.pptx --task 路演 --style 科技
  python build.py --demo --out demo.pptx --style 商务

大纲格式：
  {"title": "...", "subtitle": "...", "author": "...", "date": "...",
   "sections": [{"heading": "章节名", "kicker": "可选",
                 "type": "bullets|metrics|quote|table|image",
                 "points": [{"title": "要点", "detail": "说明"}],
                 "cards": [{"value": "12%", "label": "说明"}],
                 "rows": [["表头","表头"], ["值","值"]],
                 "quote": "句子", "source": "出处", "image": "图片路径"}]}
"""
from __future__ import annotations

import argparse
import colorsys
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL, "engine"))

from mcp_ppt.deck import build_deck  # noqa: E402

ASSETS = os.path.join(SKILL, "assets")


def _load(name: str) -> dict:
    with open(os.path.join(ASSETS, name), encoding="utf-8") as fh:
        return json.load(fh)


# ------------------------------------------------------------------ 配色

def _rel_lum(hexv: str) -> float:
    hexv = hexv.lstrip("#")

    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def _contrast(a: str, b: str) -> float:
    l1, l2 = sorted((_rel_lum(a), _rel_lum(b)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def accent_from_hsv(hue: float, sat: float, val: float, bg: str = "#FFFFFF",
                    min_contrast: float = 4.5) -> str:
    """模板实测：accent 与背景对比中位 4.72（WCAG AA）。按这个下限压明度。"""
    best = "#000000"
    for v in (val, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5, 0.45):
        r, g, b = colorsys.hsv_to_rgb(hue / 360, sat, v)
        hexv = "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
        best = hexv
        if _contrast(hexv, bg) >= min_contrast:
            return hexv
    return best


# ------------------------------------------------------------------ 风格

def pick_style(style: str | None = None, task_type: str | None = None) -> dict:
    """按风格名或任务类型选配置；两者都空时用商务。"""
    doc = _load("styles.json")
    styles = doc["styles"]
    name = style
    if not name and task_type:
        name = doc["task_type_default"].get(task_type)
    if name not in styles:
        name = "商务" if "商务" in styles else next(iter(styles))
    cfg = dict(styles[name])
    cfg["name"] = name
    return cfg


def theme_for(cfg: dict) -> dict:
    acc = cfg["accent"]
    preset = acc.get("preset", "light")
    bg = "#FFFFFF" if preset != "midnight" else "#0B1020"
    accent = accent_from_hsv(acc["hue"], acc["sat"], acc["val"], bg)
    accent2 = accent_from_hsv((acc["hue"] + 40) % 360, max(acc["sat"] - 0.12, 0.3),
                              min(acc["val"] + 0.1, 0.95), bg)
    return {"preset": preset,
            "colors": {"accent1": accent, "accent2": accent2},
            "fonts": {"major": "Microsoft YaHei", "minor": "Microsoft YaHei"}}


# ------------------------------------------------------------------ 动效

DECOR_ANIMS = [
    {"target": ["decor-dot1", "decor-dot2", "decor-dot3"], "effect": "fade",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": ["decor-chip1", "decor-chip2", "decor-chip3"], "effect": "wipe",
     "trigger": "with_previous", "duration": 0.3, "stagger": 0.05},
    {"target": "footer-rule", "effect": "wipe", "trigger": "with_previous",
     "duration": 0.4, "stagger": 0.1},
    {"target": "page-num", "effect": "fade", "trigger": "with_previous",
     "duration": 0.3, "stagger": 0.1},
]


def anims_for(spec: dict, recipe: str) -> list[dict]:
    """recipe: none / subtle / rich / cinematic —— 来自 motion.json 的分档。"""
    if recipe == "none":
        return []
    layout = spec.get("layout")
    step = 0.1 if recipe == "subtle" else 0.12
    dur = 0.5 if recipe != "cinematic" else 0.8

    if layout == "title":
        out = [
            {"target": "accent-bar", "effect": "wipe", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": dur, "stagger": step},
            {"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
             "duration": dur, "stagger": step + 0.05},
            {"target": "meta", "effect": "fade", "trigger": "with_previous",
             "duration": 0.4, "stagger": step},
        ]
        return out   # 封面不加右下装饰（设计上通常留白）
    if layout == "section":
        out = [
            {"target": "index", "effect": "wipe", "trigger": "click", "duration": dur},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": dur, "stagger": step},
        ]
        if spec.get("subtitle"):
            out.append({"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
                        "duration": dur, "stagger": step})
        out.append({"target": "page-num", "effect": "fade", "trigger": "with_previous",
                    "duration": 0.3, "stagger": step})
        return out
    if layout == "metrics":
        n = len(spec.get("cards") or [])
        out = [
            {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": dur, "stagger": step},
            {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
             "duration": 0.4, "stagger": step},
            {"target": [f"card{i}" for i in range(1, n + 1)], "effect": "faded-zoom",
             "trigger": "with_previous", "duration": dur, "stagger": step},
            {"target": [f"cardValue{i}" for i in range(1, n + 1)], "effect": "fade",
             "trigger": "with_previous", "duration": 0.4, "stagger": step},
            {"target": [f"cardLabel{i}" for i in range(1, n + 1)], "effect": "fade",
             "trigger": "with_previous", "duration": 0.4, "stagger": step},
        ]
        return out if recipe != "rich" else out + DECOR_ANIMS
    if layout == "table":
        return [
            {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": dur, "stagger": step},
            {"target": "table", "effect": "faded-zoom", "trigger": "with_previous",
             "duration": dur + 0.1, "stagger": step},
        ] + (DECOR_ANIMS if recipe == "rich" else [])
    if layout == "quote":
        return [
            {"target": "mark", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "quote", "effect": "ascend", "trigger": "with_previous",
             "duration": dur + 0.1, "stagger": step},
            {"target": "source", "effect": "fade", "trigger": "with_previous",
             "duration": 0.4, "stagger": step},
        ]
    two_col = int(spec.get("columns") or 1) >= 2
    targets = ["bullets-1", "bullets-2"] if two_col else ["bullets"]
    out = [
        {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
        {"target": "title", "effect": "fly", "trigger": "with_previous",
         "duration": dur, "stagger": step},
        {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
         "duration": 0.4, "stagger": step},
    ]
    if spec.get("image"):
        out.append({"target": "image", "effect": "faded-zoom", "trigger": "with_previous",
                    "duration": dur + 0.1, "stagger": step})
    out.append({"target": targets, "effect": "ascend", "trigger": "with_previous",
                "duration": dur, "stagger": step + 0.1, "build": "paragraph"})
    return out + (DECOR_ANIMS if recipe == "rich" else [])


# ------------------------------------------------------------------ 组装

def _bullets(points: list) -> list[dict]:
    out = []
    for p in points or []:
        if isinstance(p, str):
            out.append({"text": p, "level": 0})
            continue
        out.append({"text": p.get("title") or p.get("text", ""), "level": 0})
        if p.get("detail"):
            out.append({"text": p["detail"], "level": 1})
    return out


def make_model(outline: dict, cfg: dict) -> dict:
    theme = theme_for(cfg)
    prefs = cfg.get("layout_prefs") or {}
    recipe = (cfg.get("motion") or {}).get("recipe", "subtle")
    sections = outline.get("sections") or []
    slides: list[dict] = [{
        "layout": "title",
        "title": outline.get("title", "演示文稿"),
        "subtitle": outline.get("subtitle", ""),
        "meta": [m for m in (outline.get("date"), outline.get("author")) if m],
        "title_size": 60,
    }]
    # 目录
    if len(sections) >= 3:
        slides.append({"layout": "metrics", "kicker": "CONTENTS", "title": "目录",
                       "cards": [{"value": f"{i:02d}", "label": s.get("heading", "")}
                                 for i, s in enumerate(sections, start=1)][:4]})
    for i, sec in enumerate(sections, start=1):
        slides.append({"layout": "section", "index": i,
                       "title": sec.get("heading", f"第 {i} 部分"),
                       "subtitle": sec.get("subtitle", "")})
        kind = sec.get("type")
        if not kind:                      # 按字段自动推断版式
            if sec.get("cards"):
                kind = "metrics"
            elif sec.get("rows"):
                kind = "table"
            elif sec.get("quote"):
                kind = "quote"
            else:
                kind = "bullets"
        body = {"layout": kind, "kicker": sec.get("kicker") or sec.get("heading", ""),
                "title": sec.get("title") or sec.get("heading", ""),
                "columns": prefs.get("columns", 1)}
        if kind == "metrics":
            body["cards"] = sec.get("cards") or []
        elif kind == "table":
            body["rows"] = sec.get("rows") or []
            body["col_widths"] = sec.get("col_widths")
        elif kind == "quote":
            body["quote"] = sec.get("quote", "")
            body["source"] = sec.get("source", "")
        else:
            body["bullets"] = _bullets(sec.get("points") or [])
            if sec.get("image") or (prefs.get("image_slot") and sec.get("image")):
                body["image"] = sec.get("image")
        slides.append(body)
    slides.append({"layout": "section", "index": len(sections) + 1,
                   "title": outline.get("end_title", "谢谢"),
                   "subtitle": outline.get("end_subtitle", "THE END")})

    for s in slides:
        s["animations"] = anims_for(s, recipe)
    return {"theme": theme, "footer": outline.get("footer") or outline.get("title", ""),
            "defaults": {"transition": {"effect": "fade", "duration": 0.8}},
            "slides": slides}


def build(outline: dict, out_path: str, style: str | None = None,
          task_type: str | None = None) -> dict:
    cfg = pick_style(style, task_type)
    model = make_model(outline, cfg)
    report = build_deck(model, out_path)
    report["style"] = cfg["name"]
    report["recipe"] = (cfg.get("motion") or {}).get("recipe")
    report["theme"] = theme_for(cfg)
    return report


DEMO = {
    "title": "生成式演示",
    "subtitle": "按任务类型与风格选择模板素材",
    "author": "pptx-skill",
    "date": "2026",
    "footer": "pptx-skill demo",
    "sections": [
        {"heading": "素材怎么来", "kicker": "01 · 拆解", "title": "从 2125 个模板里提炼",
         "points": [{"title": "结构拆解", "detail": "读包内 XML：切换、动画预设、时长、版式数、配色、字体。"},
                    {"title": "识图指纹", "detail": "渲染首页算主色占比、信息密度、饱和度，归到六类风格标签。"},
                    {"title": "布局测量", "detail": "444 个内容页归一化坐标，得到边距、列数、图文关系与装饰构成。"},
                    {"title": "时序拆解", "detail": "14 个高动画模板：click 组数、with 占比、delay 瀑布、预设混用。"}]},
        {"heading": "风格怎么选", "kicker": "02 · 选择", "title": "任务类型决定默认风格",
         "points": [{"title": "汇报 → 商务", "detail": "浅底、中饱和蓝、每页 8 条动画。"},
                    {"title": "路演 → 科技", "detail": "亮度更高的蓝青、每页 18 条动画、左侧图位。"},
                    {"title": "答辩 → 极简", "detail": "零动画，靠版式与留白。"}]},
        {"heading": "效果如何", "kicker": "03 · 对标", "title": "三个硬指标",
         "cards": [{"value": "13", "label": "前景形状/页（模板 15.5）"},
                   {"value": "5.0", "label": "accent 对比（模板 4.7）"},
                   {"value": "12.7", "label": "动画/页（模板 15）"}]},
        {"heading": "怎么用", "kicker": "04 · 用法", "title": "一条命令",
         "type": "quote", "quote": "给大纲和风格，剩下的交给引擎。",
         "source": "pptx-skill"},
    ],
}


def main() -> int:
    ap = argparse.ArgumentParser(description="按风格生成对标模板的 pptx")
    ap.add_argument("outline", nargs="?", help="大纲 JSON 路径")
    ap.add_argument("--out", default="out.pptx")
    ap.add_argument("--style", default=None, help="风格名，见 assets/styles.json")
    ap.add_argument("--task", default=None, help="任务类型：汇报/答辩/路演/培训/作品集/文化/发布会")
    ap.add_argument("--demo", action="store_true", help="用内置示例大纲")
    args = ap.parse_args()

    if args.demo or not args.outline:
        outline = DEMO
    else:
        with open(args.outline, encoding="utf-8") as fh:
            outline = json.load(fh)
    report = build(outline, args.out, args.style, args.task)
    print(json.dumps({"file": report["file"], "slides": report["slide_count"],
                      "style": report["style"], "recipe": report["recipe"],
                      "accent": report["theme"]["colors"]["accent1"],
                      "warnings": report["warnings"][:3]}, ensure_ascii=False, indent=1))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
