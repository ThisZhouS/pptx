"""按全量 2120 个模板重算 skill 资产（替换原先基于 595 抽样的版本）。"""
from __future__ import annotations

import json
import os
import statistics
from collections import Counter

ROOT = r"E:\pptx\pptforge"
S = os.path.join(ROOT, "samples")
SKILL = r"E:\pptx\pptx-skill"
ASSETS = os.path.join(SKILL, "assets")


def load(name):
    p = os.path.join(S, name)
    return json.load(open(p, encoding="utf-8")) if os.path.exists(p) else None


allsum = load("resource_all_summary.json")
grid = load("grid_palette.json") or {}
ui = load("ui_layout.json") or {}
layout_style = load("layout_style.json") or []
recipes = load("timing_recipes.json") or []


def main() -> int:
    os.makedirs(ASSETS, exist_ok=True)
    g = allsum["global"]
    dist = allsum["distributions"]
    tiers = allsum["motion_tiers"]
    total = sum(tiers.values())

    # ---------- motion.json：全量分布 + 分档
    motion = {
        "source": "全部 2120 个 pptx 的动画统计（不再抽样）",
        "duration_ms": {k: v for k, v in list(g["durations_ms"].items())},
        "duration_rule": "500ms 占绝对多数（165962 次），1000ms 次之（114903）",
        "preset_mix": {k: v for k, v in list(g["presets"].items())[:10]},
        "preset_rule": "fly 29522 > wipe 26761 > ascend 25371 > fade 20220 > faded-zoom 14508",
        "transitions_pages": {k: v for k, v in list(g["transitions_pages"].items())[:8]},
        "transition_rule": "none 16644 页最多（35%）；morph 1584 页，集中在图片轮播类",
        "anim_per_page": dist["anim_per_page"],
        "motion_tiers": {
            "zero": {"files": tiers.get("zero", 0), "share": round(tiers.get("zero", 0) / total, 3),
                     "recipe": "none", "note": "37% 的模板完全不加动画"},
            "low": {"files": tiers.get("low", 0), "share": round(tiers.get("low", 0) / total, 3),
                    "recipe": "subtle"},
            "mid": {"files": tiers.get("mid", 0), "share": round(tiers.get("mid", 0) / total, 3),
                    "recipe": "rich"},
            "high": {"files": tiers.get("high", 0), "share": round(tiers.get("high", 0) / total, 3),
                     "recipe": "rich"},
            "extreme": {"files": tiers.get("extreme", 0), "share": round(tiers.get("extreme", 0) / total, 3),
                        "recipe": "rich"},
        },
        "trigger": {"click_groups_per_page": "0–1.7", "with_effect_share": "90%+",
                    "rule": "1 个 click 起手，其余 with_previous + delay 递增（瀑布）"},
        "preset_by_element": {"title": "wipe", "body": "ascend", "image": "faded-zoom",
                              "decor": "fade", "card": "faded-zoom", "footer": "fade"},
        "templates_sampled": [
            {"file": r["file"], "anim_per_page": r["anim_per_page"],
             "click_groups": r["avg_click_groups"], "with": r["avg_with"],
             "after": r["avg_after"], "max_delay_ms": r["max_delay_median"],
             "presets": r["presets_top"][:4]}
            for r in recipes
        ],
    }
    json.dump(motion, open(os.path.join(ASSETS, "motion.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- layouts.json：全量分布
    fg = ui.get("pages") or []
    layouts = {
        "source": "全部 2120 个模板（结构）+ 444 个内容页（布局）+ 15407 个形状（栅格）",
        "margins_in": {
            "top": round(statistics.median([r.get("margin_top") for r in layout_style if r.get("margin_top") is not None]) * 7.5, 2),
            "left": round(statistics.median([r.get("margin_left") for r in layout_style if r.get("margin_left") is not None]) * 13.333, 2),
            "right": round(statistics.median([r.get("margin_right") for r in layout_style if r.get("margin_right") is not None]) * 13.333, 2),
            "bottom": round(statistics.median([r.get("margin_bottom") for r in layout_style if r.get("margin_bottom") is not None]) * 7.5, 2),
        },
        "coverage": round(statistics.median([r["coverage"] for r in layout_style if r.get("coverage") is not None]), 3),
        "align_lines": statistics.median([r.get("align_lines") for r in layout_style if r.get("align_lines") is not None]),
        "shapes_per_page": dist["shapes_per_page"],
        "pages": dist["pages"],
        "media": dist["media"],
        "layouts": dist["layouts"],
        "decor_per_page": statistics.median([r["decor_total"] for r in fg]) if fg else None,
        "decor_mix": {
            "chips": round(sum(r.get("chips", 0) for r in fg) / max(len(fg), 1), 1),
            "dots": round(sum(r.get("dots", 0) for r in fg) / max(len(fg), 1), 1),
            "bars": round(sum(r.get("bars", 0) for r in fg) / max(len(fg), 1), 1),
        },
        "columns_distribution": dict(Counter(r.get("cols") for r in fg).most_common(5)),
        "image_placement": dict(Counter(r.get("img_pos") for r in fg).most_common()),
        "block_gap_in": round(statistics.median([g for r in fg for g in r.get("gaps", [])]) * 7.5, 2) if fg else None,
        "type_scale": {
            "levels": statistics.median([r.get("size_levels") for r in layout_style if r.get("size_levels")]),
            "max_to_body_ratio": statistics.median([r.get("size_ratio") for r in layout_style if r.get("size_ratio")]),
            "cover": [54, 60, 72], "page_title": [28, 34, 40],
            "section": [20, 24], "body": [14, 16, 18], "caption": 12, "footer": 10,
        },
        "grid": {"0.25in": 0.343, "0.5in": 0.337, "1.0in": 0.322, "12col": 0.171,
                 "note": "按英寸对齐，不用 12 列栅格"},
    }
    json.dump(layouts, open(os.path.join(ASSETS, "layouts.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- decor.json：全量板块画像
    decor = {
        "source": "全部 2120 个模板的板块画像 + 444 个内容页的装饰统计",
        "elements": {
            "chip": {"size_in": [0.17, 0.045], "count": 3, "role": "小色块，装饰主体"},
            "dot": {"size_in": [0.085, 0.085], "count": 3, "role": "圆点"},
            "footer_rule": {"height_pt": 0.75, "width": "content_w - 0.9in"},
            "page_num": {"size_pt": 10, "pos": "右下"},
        },
        "sections": allsum["groups"],
    }
    json.dump(decor, open(os.path.join(ASSETS, "decor.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    # ---------- styles.json：按全量分档重排风格
    # 全量 2120 个主题实测中位：蓝族 sat 0.58 / val 0.77 / 对比 4.03；
    # 红橙 sat 0.50 / val 0.80；青 sat 0.64 / val 0.62。
    # 生成端仍按对比度 >=4.5 压明度（比模板中位略深，换取正文可读性）。
    accents = {
        "极简": {"hue": 208, "sat": 0.55, "val": 0.80, "preset": "light", "tier": "zero"},
        "商务": {"hue": 212, "sat": 0.58, "val": 0.70, "preset": "light", "tier": "low"},
        "科技": {"hue": 196, "sat": 0.64, "val": 0.82, "preset": "light", "tier": "high"},
        "国潮": {"hue": 8, "sat": 0.50, "val": 0.62, "preset": "light", "tier": "low"},
        "文艺": {"hue": 32, "sat": 0.30, "val": 0.66, "preset": "light", "tier": "zero"},
        "沉稳深色": {"hue": 210, "sat": 0.55, "val": 0.72, "preset": "midnight", "tier": "mid"},
    }
    styles = {
        "source": "全部 2120 个模板：动画分档 {}，形状/页 P10/50/90 = {}/{}/{}".format(
            tiers, dist["shapes_per_page"]["p10"], dist["shapes_per_page"]["p50"],
            dist["shapes_per_page"]["p90"]),
        "styles": {
            name: {
                "accent": {k: v for k, v in acc.items() if k != "tier"},
                "motion": {"tier": acc["tier"],
                           "recipe": {"zero": "none", "low": "subtle",
                                      "mid": "rich", "high": "rich"}[acc["tier"]]},
                "layout_prefs": {
                    "columns": 1 if name in ("极简", "文艺") else 2,
                    "image_slot": name in ("科技", "文艺", "国潮"),
                    "decor_density": "low" if name in ("极简", "文艺") else "normal",
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
            for name, acc in accents.items()
        },
        "task_type_default": {
            "汇报": "商务", "答辩": "极简", "路演": "科技", "培训": "商务",
            "作品集": "文艺", "文化": "国潮", "发布会": "沉稳深色",
        },
    }
    json.dump(styles, open(os.path.join(ASSETS, "styles.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    print("[rebuilt]", sorted(os.listdir(ASSETS)))
    print("tiers:", tiers)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())