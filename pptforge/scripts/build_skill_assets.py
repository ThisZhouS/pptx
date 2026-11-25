"""从 Resource 拆解结果里提炼 skill 资产：风格库、配色库、版式参数、动效配方、装饰规格。

输入 samples/ 下已有的分析 JSON，输出 pptx-skill/assets/ 下的可复用素材。
"""
from __future__ import annotations

import json
import os
import statistics
from collections import Counter, defaultdict

ROOT = r"E:\pptx\pptforge"
S = os.path.join(ROOT, "samples")
SKILL = r"E:\pptx\pptx-skill"
ASSETS = os.path.join(SKILL, "assets")


def load(name):
    p = os.path.join(S, name)
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else []


def median(vals, default=None):
    vals = [v for v in vals if v is not None]
    return round(statistics.median(vals), 3) if vals else default


def main() -> int:
    os.makedirs(ASSETS, exist_ok=True)
    grid = load("grid_palette.json") or {}
    layout_style = load("layout_style.json")
    summary = load("resource_summary.json")
    recipes = load("timing_recipes.json")
    ui = load("ui_layout.json") or {}
    vision = load("resource_vision.json")

    # ---------- 1) palettes.json：从 91 个主题提取，按"色相族 + 明暗"分组
    palettes = grid.get("palettes", [])
    groups = defaultdict(list)
    for p in palettes:
        hue = p["a_h"]
        family = ("蓝" if 200 <= hue < 250 else "红橙" if hue < 60 or hue >= 330 else
                  "青" if 160 <= hue < 200 else "绿" if 70 <= hue < 160 else "紫")
        groups[family].append(p)
    pal_out = {}
    for fam, items in groups.items():
        pal_out[fam] = {
            "count": len(items),
            "accent_sat": median([i["a_s"] for i in items]),
            "accent_val": median([i["a_v"] for i in items]),
            "contrast_accent_bg": median([i["contrast_accent_bg"] for i in items]),
            "samples": [i["accent1"] for i in items[:5]],
            "mode": Counter(i["mode"] for i in items).most_common(1)[0][0],
        }
    json.dump({"source": "Resource/ 91 个主题的 clrScheme 提取",
               "hue_families": pal_out,
               "note": "浅底占 91/91；accent 与背景对比中位 4.72（WCAG AA）"},
              open(os.path.join(ASSETS, "palettes.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- 2) layouts.json：版式参数（中位值）
    fg = ui.get("pages") or []
    layouts = {
        "source": "444 个内容页 + 91 个样本的布局测量",
        "margins_in": {
            "top": round(median([r.get("margin_top") for r in layout_style]) * 7.5, 2),
            "left": round(median([r.get("margin_left") for r in layout_style]) * 13.333, 2),
            "right": round(median([r.get("margin_right") for r in layout_style]) * 13.333, 2),
            "bottom": round(median([r.get("margin_bottom") for r in layout_style]) * 7.5, 2),
        },
        "coverage": median([r.get("coverage") for r in layout_style]),
        "align_lines": median([r.get("align_lines") for r in layout_style]),
        "fg_shapes_per_page": median([r.get("fg") for r in fg]),
        "decor_per_page": median([r.get("decor_total") for r in fg]),
        "decor_mix": {
            "chips": round(sum(r.get("chips", 0) for r in fg) / max(len(fg), 1), 1),
            "dots": round(sum(r.get("dots", 0) for r in fg) / max(len(fg), 1), 1),
            "bars": round(sum(r.get("bars", 0) for r in fg) / max(len(fg), 1), 1),
        },
        "columns_distribution": dict(Counter(r.get("cols") for r in fg).most_common(5)),
        "image_placement": dict(Counter(r.get("img_pos") for r in fg).most_common()),
        "block_gap_in": round(median([g for r in fg for g in r.get("gaps", [])]) * 7.5, 2),
        "type_scale": {
            "levels": median([r.get("size_levels") for r in layout_style]),
            "max_to_body_ratio": median([r.get("size_ratio") for r in layout_style]),
            "cover": [54, 60, 72], "page_title": [28, 34, 40],
            "section": [20, 24], "body": [14, 16, 18], "caption": 12, "footer": 10,
        },
        "grid": {"0.25in": 0.343, "0.5in": 0.337, "1.0in": 0.322, "12col": 0.171,
                 "note": "按英寸对齐，不用 12 列栅格"},
    }
    json.dump(layouts, open(os.path.join(ASSETS, "layouts.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- 3) motion.json：动效配方
    motion = {
        "source": "594 个模板的动画统计 + 14 个高动画模板的时序拆解",
        "duration_ms": {"primary": 500, "secondary": 1000, "light": 250},
        "preset_mix": {"fly": 10436, "wipe": 7918, "ascend": 6715, "fade": 5880,
                       "faded-zoom": 4778, "split": 1347, "random-bars": 1275},
        "trigger": {"click_groups_per_page": "0–1.7", "with_effect_share": "90%+",
                    "rule": "1 个 click 起手，其余 with_previous + delay 递增"},
        "delay_waterfall": {"start": 0.1, "step": 0.1, "max_observed_ms": 7000},
        "anims_per_page": {"template_median": 15, "range": [15, 34]},
        "preset_by_element": {"title": "wipe", "body": "ascend", "image": "faded-zoom",
                              "decor": "fade", "card": "faded-zoom", "footer": "fade"},
        "transitions": {"none": 5380, "random": 3022, "fade": 472, "morph": 469,
                        "rule": "汇报用 fade/none；morph 只用于图片推镜/轮播"},
        "templates": [
            {"file": r["file"], "anim_per_page": r["anim_per_page"],
             "click_groups": r["avg_click_groups"], "with": r["avg_with"],
             "after": r["avg_after"], "max_delay_ms": r["max_delay_median"],
             "presets": r["presets_top"][:4]}
            for r in recipes
        ],
    }
    json.dump(motion, open(os.path.join(ASSETS, "motion.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- 4) decor.json：装饰规格与板块画像
    decor = {
        "source": "444 个内容页的装饰统计 + 62 个板块画像",
        "elements": {
            "chip": {"size_in": [0.17, 0.045], "count": 3, "role": "小色块，装饰主体"},
            "dot": {"size_in": [0.085, 0.085], "count": 3, "role": "圆点，标点般的存在"},
            "footer_rule": {"height_pt": 0.75, "width": "content_w - 0.9in"},
            "page_num": {"size_pt": 10, "pos": "右下"},
        },
        "sections": [
            {"group": s["group"], "samples": s["samples"], "pages": s["pages_avg"],
             "anim_per_page": s["anim_per_page"], "shapes_per_page": s["shapes_per_page"],
             "layouts": s["layouts_avg"], "transition": s["top_transition"],
             "preset": s["top_preset"], "duration": s["top_duration"],
             "styles": s.get("styles") or {}}
            for s in (summary.get("groups") or [])
        ],
    }
    json.dump(decor, open(os.path.join(ASSETS, "decor.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- 5) styles.json：风格库（按任务类型 + 风格）
    accent_by_style = {
        "极简": {"hue": 205, "sat": 0.72, "val": 0.80, "preset": "light"},
        "商务": {"hue": 212, "sat": 0.68, "val": 0.62, "preset": "light"},
        "科技": {"hue": 195, "sat": 0.78, "val": 0.85, "preset": "light"},
        "国潮": {"hue": 8, "sat": 0.62, "val": 0.55, "preset": "light"},
        "文艺": {"hue": 32, "sat": 0.35, "val": 0.62, "preset": "light"},
        "沉稳深色": {"hue": 210, "sat": 0.55, "val": 0.72, "preset": "midnight"},
    }
    motion_by_style = {
        "极简": {"anims": 0, "recipe": "none"},
        "商务": {"anims": 8, "recipe": "subtle"},
        "科技": {"anims": 18, "recipe": "rich"},
        "国潮": {"anims": 6, "recipe": "subtle"},
        "文艺": {"anims": 4, "recipe": "subtle"},
        "沉稳深色": {"anims": 12, "recipe": "cinematic"},
    }
    styles = {
        "source": "Resource 62 板块画像 + 91 主题配色 + 14 个高动画配方",
        "styles": {
            name: {
                "accent": accent_by_style[name],
                "motion": motion_by_style[name],
                "layout_prefs": {
                    "columns": 1 if name in ("极简", "文艺") else 2,
                    "image_slot": name in ("科技", "文艺", "国潮"),
                    "decor_density": "low" if name == "极简" else "normal",
                },
                "use_for": {
                    "极简": ["汇报", "答辩", "评审"],
                    "商务": ["工作汇报", "方案", "投标"],
                    "科技": ["产品", "技术分享", "路演"],
                    "国潮": ["文化", "非遗", "品牌"],
                    "文艺": ["画册", "作品集", "读书会"],
                    "沉稳深色": ["发布会", "夜场演示", "Fintech"],
                }[name],
            }
            for name in accent_by_style
        },
        "task_type_default": {
            "汇报": "商务", "答辩": "极简", "路演": "科技", "培训": "商务",
            "作品集": "文艺", "文化": "国潮", "发布会": "沉稳深色",
        },
    }
    json.dump(styles, open(os.path.join(ASSETS, "styles.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print("[written]", os.listdir(ASSETS))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())