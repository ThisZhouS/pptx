"""效果目录：把逆向得到的 schema/transitions.json 与 schema/animations.json 读进内存。

AI 只需给"名字"（id / 中文别名 / COM 枚举名 / PowerPoint 常见叫法），
catalog 负责解析成 OOXML 事实（元素、属性、preset 三要素、行为节点模板）。
"""
from __future__ import annotations

import json
import os
import re
from functools import lru_cache

HERE = os.path.dirname(os.path.abspath(__file__))
SCHEMA_DIR = os.path.join(os.path.dirname(HERE), "schema")

# AI / 用户最可能用的写法 → 目录 id
TRANSITION_ALIASES = {
    "淡出": "fade", "淡入": "fade", "fade-smooth": "fade-smoothly", "平滑淡出": "fade-smoothly",
    "淡出到黑场": "fade", "黑场淡出": "fade", "cut": "cut", "直接切换": "cut",
    "推开": "push", "推进": "push", "push-left": "push", "擦除": "wipe-left",
    "wipe": "wipe-left", "覆盖": "cover-left", "cover": "cover-left",
    "揭开": "uncover-left", "uncover": "uncover-left", "分割": "split-horizontal-out",
    "split": "split-horizontal-out", "溶解": "dissolve", "dissolve": "dissolve",
    "百叶窗": "blinds-horizontal", "棋盘": "checkerboard-across", "随机": "random",
    "形状": "circle-out", "圆形": "circle-out", "菱形": "diamond-out",
    "轮子": "wheel1-spoke", "wheel": "wheel1-spoke", "缩放": "box-out", "zoom": "box-out",
    "涡流": "vortex-left", "vortex": "vortex-left", "涟漪": "ripple-center",
    "ripple": "ripple-center", "闪耀": "glitter-diamond-left",
    "glitter": "glitter-diamond-left", "glitter-left": "glitter-diamond-left",
    "图库": "gallery-left", "gallery": "gallery-left", "传送带": "conveyor-left",
    "门": "doors-vertical", "窗户": "window-vertical", "曲变": "warp-in",
    "飞越": "fly-through-in", "蜂巢": "honeycomb", "摩天轮": "ferris-wheel-left",
    "切换": "switch-left", "翻转": "flip-left", "碎片": "shred-strips-in",
    "立方体": "cube-left", "轨道": "orbit-left", "平移": "pan-left",
    "变形": "morph-by-object", "morph": "morph-by-object",
    "morph-by-object": "morph-by-object", "morph-by-word": "morph-by-word",
    "翻覆": "fall-over-left", "垂坠": "drape-left", "窗帘": "curtains",
    "风": "wind-left", "剥落": "peel-off-left", "卷页": "page-curl-single-left",
    "飞机": "airplane-left", "折纸": "origami-left", "华丽": "prestige",
    "碎裂": "fracture", "压碎": "crush", "闪光灯": "flashbulb",
    "无": "none", "none": "none", "无切换": "none",
}

ANIM_ALIASES = {
    "淡入": "fade", "淡出（进入）": "fade", "fade": "fade", "出现": "appear",
    "appear": "appear", "飞入": "fly", "fly": "fly", "擦除": "wipe", "wipe": "wipe",
    "缩放": "zoom", "zoom": "zoom", "百叶窗": "blinds", "棋盘": "checkerboard",
    "圆形": "circle", "菱形": "diamond", "溶解": "dissolve", "轮子": "wheel",
    "分割": "split", "伸展": "stretch", "剥离": "strips", "螺旋": "spiral",
    "旋转强调": "spin", "spin": "spin", "放大缩小": "grow-shrink",
    "grow-shrink": "grow-shrink", "弹跳": "bounce", "bounce": "bounce",
    "浮动": "float", "升起": "rise-up", "上升": "ascend", "下降": "descend",
    "翻转": "flip", "闪烁一次": "flash-once", "透明": "transparency",
    "跷跷板": "teeter", "波浪": "wave", "脉冲": "grow-shrink",
    "路径": "path-circle", "圆形路径": "path-circle", "path-circle": "path-circle",
    "锯齿": "path-zigzag", "path-zigzag": "path-zigzag",
}

TRIGGER_ALIASES = {
    "click": "click", "单击时": "click", "on_click": "click", "点击": "click",
    "with_previous": "with_previous", "与上一动画同时": "with_previous",
    "with": "with_previous", "同时": "with_previous",
    "after_previous": "after_previous", "上一动画之后": "after_previous",
    "after": "after_previous", "之后": "after_previous",
}


@lru_cache(maxsize=1)
def _load() -> tuple[dict, dict]:
    with open(os.path.join(SCHEMA_DIR, "transitions.json"), encoding="utf-8") as fh:
        transitions = json.load(fh)
    with open(os.path.join(SCHEMA_DIR, "animations.json"), encoding="utf-8") as fh:
        animations = json.load(fh)
    return transitions, animations


class Catalog:
    def __init__(self) -> None:
        self.transitions_doc, self.animations_doc = _load()
        self.transitions: dict[str, dict] = self.transitions_doc["effects"]
        self.animations: dict[str, dict] = self.animations_doc["effects"]
        self._t_index = self._build_index(self.transitions)
        self._a_index = self._build_index(self.animations)

    @staticmethod
    def _build_index(effects: dict[str, dict]) -> dict[str, str]:
        idx: dict[str, str] = {}
        for eid, e in effects.items():
            idx[eid.lower()] = eid
            idx[e["enum_name"].lower()] = eid
            zh = e.get("zh")
            if zh:
                idx.setdefault(zh, eid)
        return idx

    # ---------- 解析 ----------
    def resolve_transition(self, name: str | None) -> str:
        if not name:
            return "none"
        key = str(name).strip()
        low = key.lower()
        if low in self._t_index:
            return self._t_index[low]
        if key in TRANSITION_ALIASES:
            return TRANSITION_ALIASES[key]
        if low in TRANSITION_ALIASES:
            return TRANSITION_ALIASES[low]
        # 宽松匹配：cover-left ↔ cover_left ↔ cover left
        norm = re.sub(r"[\s_]+", "-", low)
        if norm in self._t_index:
            return self._t_index[norm]
        if norm in TRANSITION_ALIASES:
            return TRANSITION_ALIASES[norm]
        raise KeyError(f"未知切换效果: {name!r}。可用效果见 transitions_catalog 工具。")

    def resolve_animation(self, name: str | None) -> str:
        if not name:
            return "fade"
        key = str(name).strip()
        low = key.lower()
        if low in self._a_index:
            return self._a_index[low]
        if key in ANIM_ALIASES:
            return ANIM_ALIASES[key]
        if low in ANIM_ALIASES:
            return ANIM_ALIASES[low]
        norm = re.sub(r"[\s_]+", "-", low)
        if norm in self._a_index:
            return self._a_index[norm]
        if norm in ANIM_ALIASES:
            return ANIM_ALIASES[norm]
        raise KeyError(f"未知动画效果: {name!r}。可用效果见 animations_catalog 工具。")

    @staticmethod
    def resolve_trigger(name: str | None) -> str:
        if not name:
            return "click"
        key = str(name).strip()
        if key in TRIGGER_ALIASES:
            return TRIGGER_ALIASES[key]
        low = key.lower()
        if low in TRIGGER_ALIASES:
            return TRIGGER_ALIASES[low]
        raise KeyError(f"未知触发方式: {name!r}（click / with_previous / after_previous）")

    # ---------- 查询 ----------
    def transition_spec(self, name: str | None) -> dict:
        return self.transitions[self.resolve_transition(name)]

    def animation_spec(self, name: str | None, exit_: bool = False) -> dict:
        eff = self.animations[self.resolve_animation(name)]
        if exit_ and "exit" in eff["variants"]:
            return {
                "id": eff["id"],
                "zh": eff["zh"],
                "preset": eff["variants"]["exit"]["preset"],
                "behaviors": eff["variants"]["exit"]["behaviors"],
            }
        return {
            "id": eff["id"],
            "zh": eff["zh"],
            "preset": eff["preset"],
            "behaviors": eff["behaviors"],
        }

    def list_transitions(self, query: str | None = None) -> list[dict]:
        out = []
        for eid, e in self.transitions.items():
            if query and query.lower() not in (eid + e.get("zh", "") + e["enum_name"]).lower():
                continue
            out.append({"id": eid, "zh": e.get("zh"), "element": e.get("element"),
                        "requires": e.get("requires")})
        return out

    def list_animations(self, query: str | None = None, only_exit: bool = False) -> list[dict]:
        out = []
        for eid, e in self.animations.items():
            if query and query.lower() not in (eid + e.get("zh", "") + e["enum_name"]).lower():
                continue
            has_exit = "exit" in e["variants"]
            if only_exit and not has_exit:
                continue
            out.append({"id": eid, "zh": e.get("zh"), "presetID": e["preset"]["presetID"],
                        "class": e["preset"]["presetClass"], "exit_variant": has_exit})
        return out


_catalog: Catalog | None = None


def get_catalog() -> Catalog:
    global _catalog
    if _catalog is None:
        _catalog = Catalog()
    return _catalog