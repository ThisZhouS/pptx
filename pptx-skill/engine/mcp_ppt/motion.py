"""动效套餐：把"风格"翻译成具体的切换 + 动画组合。

参数来自 Resource/ 下 2125 个模板的抽样拆解（docs/TEMPLATE_LEARNING.md）：

  预设频次   wipe 1184 > fade 1146 > fly 907 > ascend 575 > faded-zoom 567
  时长      500ms 7752 次（绝对主流）> 1000ms 2916 > 250ms 310
  密度      有动画的模板每页 6.8 条动画 / 19.8 个形状
  切换      none 416 页、fade 112 页；random/blinds 多见于模板展示，汇报场景回避
  morph     仅 5/60 个模板使用，且全部是"图片放大/轮播"，时长 1250–2000ms

所以默认分工：标题 wipe、正文 fade、卡片与图片 faded-zoom，全部 500ms，
同组用 with_previous + stagger 0.1 错峰；切换用 fade（不用花哨切换）。
"""
from __future__ import annotations

from dataclasses import dataclass, field

DUR = 0.5          # 主流时长
DUR_BIG = 1.0      # 重大元素 / 叙事页
STAGGER = 0.1      # 同组错峰


@dataclass
class MotionRecipe:
    name: str
    cover_transition: dict
    section_transition: dict
    content_transition: dict
    cover_animations: list[dict] = field(default_factory=list)
    content_animations: list[dict] = field(default_factory=list)
    card_animations: list[dict] = field(default_factory=list)


RECIPES: dict[str, MotionRecipe] = {
    # 克制：汇报主力。标题擦出、正文淡入、卡片淡出缩放
    "subtle": MotionRecipe(
        name="subtle",
        cover_transition={"effect": "fade", "duration": 0.8},
        section_transition={"effect": "fade", "duration": 0.8},
        content_transition={"effect": "fade", "duration": 0.8},
        cover_animations=[
            {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": 0.6, "stagger": 0.1},
            {"target": "subtitle", "effect": "ascend", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.15},
            {"target": "meta", "effect": "fade", "trigger": "with_previous",
             "duration": 0.4, "stagger": 0.2},
            {"target": "rule", "effect": "wipe", "trigger": "with_previous",
             "duration": 0.4, "stagger": 0.25},
        ],
        content_animations=[
            {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.1},
            {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
             "duration": 0.4, "stagger": 0.15},
            {"target": "bullets", "effect": "ascend", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.2, "build": "paragraph"},
            {"target": "footer", "effect": "fade", "trigger": "with_previous",
             "duration": 0.3, "stagger": 0.3},
        ],
        card_animations=[
            {"target": "kicker", "effect": "fade", "trigger": "click", "duration": 0.4},
            {"target": "title", "effect": "fly", "trigger": "with_previous",
             "duration": 0.5, "stagger": 0.1},
            {"target": "title-rule", "effect": "wipe", "trigger": "with_previous",
             "duration": 0.4, "stagger": 0.15},
            {"target": ["card1", "card2", "card3", "card4"], "effect": "faded-zoom",
             "trigger": "with_previous", "duration": 0.5, "stagger": 0.12},
            {"target": ["cardValue1", "cardValue2", "cardValue3", "cardValue4"],
             "effect": "fade", "trigger": "with_previous", "duration": 0.4, "stagger": 0.08},
            {"target": ["cardLabel1", "cardLabel2", "cardLabel3", "cardLabel4"],
             "effect": "fade", "trigger": "with_previous", "duration": 0.4, "stagger": 0.08},
        ],
    ),
    # 有力：发布会。飞入 + 更强切换
    "bold": MotionRecipe(
        name="bold",
        cover_transition={"effect": "push-left", "duration": 0.6},
        section_transition={"effect": "cover-left", "duration": 0.6},
        content_transition={"effect": "push-up", "duration": 0.6},
        cover_animations=[
            {"target": "title", "effect": "fly", "trigger": "click", "duration": 0.6},
            {"target": "subtitle", "effect": "fly", "trigger": "with_previous",
             "duration": 0.6, "stagger": STAGGER},
        ],
        content_animations=[
            {"target": "title", "effect": "wipe", "trigger": "click", "duration": 0.5},
            {"target": "bullets", "effect": "fly", "trigger": "with_previous",
             "duration": 0.5, "stagger": STAGGER},
        ],
        card_animations=[
            {"target": ["card1", "card2", "card3", "card4"], "effect": "zoom",
             "trigger": "click", "duration": 0.5, "stagger": STAGGER},
        ],
    ),
    # 叙事：morph 切页 + 长时长淡入（适合配图、宣传片式讲述）
    "cinematic": MotionRecipe(
        name="cinematic",
        cover_transition={"effect": "morph", "duration": 2.0},
        section_transition={"effect": "morph", "duration": 2.0},
        content_transition={"effect": "fade", "duration": 1.0},
        cover_animations=[
            {"target": "title", "effect": "wipe", "trigger": "click", "duration": 0.8},
            {"target": "subtitle", "effect": "fade", "trigger": "with_previous",
             "duration": DUR_BIG, "stagger": 0.2, "smooth_start": 0.3, "smooth_end": 0.3},
        ],
        content_animations=[
            {"target": "title", "effect": "wipe", "trigger": "click", "duration": 0.6},
            {"target": "bullets", "effect": "fade", "trigger": "with_previous",
             "duration": DUR_BIG, "stagger": 0.2},
        ],
        card_animations=[
            {"target": ["card1", "card2", "card3", "card4"], "effect": "faded-zoom",
             "trigger": "click", "duration": DUR_BIG, "stagger": 0.2},
        ],
    ),
    # 技术：擦除为主，节奏偏快
    "tech": MotionRecipe(
        name="tech",
        cover_transition={"effect": "fade", "duration": 0.8},
        section_transition={"effect": "fade", "duration": 0.8},
        content_transition={"effect": "fade", "duration": 0.8},
        cover_animations=[
            {"target": "title", "effect": "wipe", "trigger": "click", "duration": DUR},
            {"target": "subtitle", "effect": "wipe", "trigger": "with_previous",
             "duration": DUR, "stagger": 0.12},
        ],
        content_animations=[
            {"target": "title", "effect": "wipe", "trigger": "click", "duration": 0.5},
            {"target": "bullets", "effect": "wipe", "trigger": "with_previous",
             "duration": DUR, "stagger": 0.12},
        ],
        card_animations=[
            {"target": ["card1", "card2", "card3", "card4"], "effect": "faded-zoom",
             "trigger": "click", "duration": DUR, "stagger": 0.12},
        ],
    ),
}


def get_recipe(name: str | None) -> MotionRecipe:
    key = str(name or "subtle").strip().lower()
    aliases = {"克制": "subtle", "稳重": "subtle", "有力": "bold", "发布会": "bold",
               "电影": "cinematic", "宣传": "cinematic", "科技": "tech", "技术": "tech"}
    return RECIPES.get(key) or RECIPES.get(aliases.get(key, "subtle"), RECIPES["subtle"])


def classify_slide(shape_names: list[str], layout: str | None = None) -> str:
    names = set(shape_names)
    if layout in ("title",):
        return "cover"
    if layout in ("section",):
        return "section"
    if any(str(n).startswith("card") for n in names):
        return "cards"
    if layout in ("metrics",):
        return "cards"
    return "content"


def recipe_for(recipe: MotionRecipe, kind: str) -> dict:
    trans = {"cover": recipe.cover_transition, "section": recipe.section_transition,
             "content": recipe.content_transition, "cards": recipe.content_transition}[kind]
    anims = {"cover": recipe.cover_animations, "section": recipe.cover_animations,
             "content": recipe.content_animations, "cards": recipe.card_animations}[kind]
    return {"transition": dict(trans), "animations": [dict(a) for a in anims]}


def apply_recipe_to_model(model: dict, style: str) -> dict:
    """给整份 model 套用动效套餐（未显式指定动效的页）。"""
    recipe = get_recipe(style)
    out = dict(model)
    slides = []
    for spec in model.get("slides") or []:
        spec = dict(spec)
        kind = classify_slide(list(spec.keys()), str(spec.get("layout") or spec.get("type") or ""))
        suggestion = recipe_for(recipe, kind)
        if "transition" not in spec:
            spec["transition"] = suggestion["transition"]
        if "animations" not in spec and suggestion["animations"]:
            spec["animations"] = suggestion["animations"]
        slides.append(spec)
    out["slides"] = slides
    out.setdefault("motion_style", recipe.name)
    return out
