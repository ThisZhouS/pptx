"""把 ground-truth 采样编译成引擎目录（schema/transitions.json, schema/animations.json）。

编译规则（来自逆向结论）
------------------------
* 切换：id = COM 枚举名去 ppEffect 前缀的 kebab 形式；附 element/attrs/requires；
  p14/p15/p159 元素必须走 mc:AlternateContent。
* 动画：id = MsoAnimEffect 名去 msoAnimEffect 前缀的 kebab 形式；
  preset 三要素 + 行为节点模板（spid 占位），entr/emph/path 与 exit 变体各一份。
* 中文别名由族名 + 方向规则生成，覆盖 PowerPoint 中文 UI 的叫法。
"""
from __future__ import annotations

import json
import os
import re

from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
ROOT = os.path.dirname(HERE)
SCHEMA = os.path.join(ROOT, "schema")
os.makedirs(SCHEMA, exist_ok=True)

NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
URI2PREFIX = {v: k for k, v in NSMAP.items()}

# ---------------------------------------------------------------- 中文别名

FAMILY_ZH = {
    "cut": "切割", "cut-through-black": "通过黑场切割", "random": "随机",
    "blinds": "百叶窗", "checkerboard": "棋盘", "cover": "覆盖", "dissolve": "溶解",
    "fade": "淡出", "fade-smoothly": "平滑淡出", "uncover": "揭开",
    "random-bars": "随机线条", "strips": "剥离", "wipe": "擦除", "box": "缩放",
    "split": "分割", "flash-once": "闪烁一次", "appear": "出现", "circle": "圆形",
    "diamond": "菱形", "comb": "梳状", "newsflash": "新闻快报", "plus": "加号",
    "push": "推进", "wedge": "楔形", "wheel": "轮子", "wheel-reverse": "反向轮子",
    "vortex": "涡流", "ripple": "涟漪", "glitter": "闪耀", "gallery": "图库",
    "conveyor": "传送带", "doors": "门", "window": "窗户", "warp": "曲变",
    "fly-through": "飞越", "reveal": "揭开", "honeycomb": "蜂巢",
    "ferris-wheel": "摩天轮", "switch": "切换", "flip": "翻转",
    "flashbulb": "闪光灯", "shred-strips": "碎片（条状）", "shred-rectangle": "碎片（矩形）",
    "cube": "立方体", "orbit": "轨道", "pan": "平移", "fall-over": "翻覆",
    "drape": "垂坠", "curtains": "窗帘", "wind": "风", "prestige": "华丽",
    "fracture": "碎裂", "crush": "压碎", "peel-off": "剥落",
    "page-curl-single": "卷页（单页）", "page-curl-double": "卷页（双页）",
    "airplane": "飞机", "origami": "折纸", "morph": "变形",
}

DIR_ZH = {
    "left": "左", "right": "右", "up": "上", "down": "下",
    "left-up": "左上", "right-up": "右上", "left-down": "左下", "right-down": "右下",
    "up-left": "左上", "up-right": "右上", "down-left": "左下", "down-right": "右下",
    "horizontal": "水平", "vertical": "垂直", "across": "横向",
    "in": "向内", "out": "向外", "in-slightly": "略向内", "out-slightly": "略向外",
    "horizontal-out": "水平向外", "horizontal-in": "水平向内",
    "vertical-out": "垂直向外", "vertical-in": "垂直向内",
}

EXTRA_ZH = {
    "box-in": "缩放（向内）", "box-out": "缩放（向外）",
    "zoom-in": "缩放（向内）", "zoom-in-slightly": "缩放（略向内）",
    "zoom-out": "缩放（向外）", "zoom-out-slightly": "缩放（略向外）",
    "zoom-center": "缩放（中心）", "zoom-bottom": "缩放（底部）",
    "stretch-across": "伸展（横跨）", "spiral": "螺旋", "swivel": "旋转",
    "stretch-left": "伸展（左）", "stretch-up": "伸展（上）",
    "stretch-right": "伸展（右）", "stretch-down": "伸展（下）",
    "circle-out": "圆形（向外）", "diamond-out": "菱形（向外）", "plus-out": "加号（向外）",
    "flash-once-fast": "闪烁一次（快）", "flash-once-medium": "闪烁一次（中）",
    "flash-once-slow": "闪烁一次（慢）",
    "wheel-1-spoke": "轮子（1 根辐条）", "wheel-2-spokes": "轮子（2 根辐条）",
    "wheel-3-spokes": "轮子（3 根辐条）", "wheel-4-spokes": "轮子（4 根辐条）",
    "wheel-8-spokes": "轮子（8 根辐条）",
    "wheel-reverse-1-spoke": "反向轮子（1 根辐条）",
    "fly-through-in-bounce": "飞越（进入带弹跳）", "fly-through-out-bounce": "飞越（退出带弹跳）",
    "shred-strips-in": "碎片（条状进入）", "shred-strips-out": "碎片（条状退出）",
    "shred-rectangle-in": "碎片（矩形进入）", "shred-rectangle-out": "碎片（矩形退出）",
    "reveal-smooth-left": "平滑揭开（左）", "reveal-smooth-right": "平滑揭开（右）",
    "reveal-black-left": "揭开（黑场左）", "reveal-black-right": "揭开（黑场右）",
    "morph-by-object": "变形（按对象）", "morph-by-word": "变形（按单词）",
    "morph-by-char": "变形（按字符）",
    "page-curl-single-left": "卷页（单页，左）", "page-curl-single-right": "卷页（单页，右）",
    "page-curl-double-left": "卷页（双页，左）", "page-curl-double-right": "卷页（双页，右）",
}


def zh_for(eid: str) -> str:
    if eid in EXTRA_ZH:
        return EXTRA_ZH[eid]
    parts = eid.split("-")
    # 尝试最长族名匹配
    for i in range(len(parts), 0, -1):
        fam = "-".join(parts[:i])
        if fam in FAMILY_ZH:
            rest = parts[i:]
            if not rest:
                return FAMILY_ZH[fam]
            dir_key = "-".join(rest)
            if dir_key in DIR_ZH:
                return f"{FAMILY_ZH[fam]}（{DIR_ZH[dir_key]}）"
            return f"{FAMILY_ZH[fam]}（{dir_key}）"
    return eid


ANIM_ZH = {
    "appear": "出现", "fly": "飞入", "blinds": "百叶窗", "box": "盒状",
    "checkerboard": "棋盘", "circle": "圆形", "crawl": "爬行", "diamond": "菱形",
    "dissolve": "溶解", "fade": "淡入", "flash-once": "闪烁一次", "peek": "切入",
    "plus": "十字形扩展", "random-bars": "随机线条", "spiral": "螺旋", "split": "分割",
    "stretch": "伸展", "strips": "剥离", "swivel": "旋转", "wedge": "楔入",
    "wheel": "轮子", "wipe": "擦除", "zoom": "缩放", "random-effects": "随机效果",
    "boomerang": "回旋", "bounce": "弹跳", "color-reveal": "颜色显现", "credits": "字幕",
    "ease-in": "缓动进入", "float": "浮动", "grow-and-turn": "增长并旋转",
    "light-speed": "光速", "pinwheel": "纸风车", "rise-up": "上升", "swish": "嗖嗖声效",
    "thin-line": "细线", "unfold": "展开", "whip": "鞭打", "ascend": "上升",
    "center-revolve": "中心旋转", "faded-swivel": "淡出旋转", "descend": "下降",
    "sling": "投掷", "spinner": "回旋镖", "stretchy": "弹性伸展", "zip": "快速移动",
    "arc-up": "向上弧线", "faded-zoom": "淡出缩放", "glide": "滑行", "expand": "展开",
    "flip": "翻转", "shimmer": "微光闪烁", "fold": "折叠",
    "change-fill-color": "填充颜色", "change-font": "更改字体",
    "change-font-color": "字体颜色", "change-font-size": "字号",
    "change-font-style": "字形", "grow-shrink": "放大/缩小",
    "change-line-color": "线条颜色", "spin": "旋转", "transparency": "透明",
    "bold-flash": "加粗闪烁", "blast": "爆炸", "bold-reveal": "加粗显现",
    "brush-on-color": "画笔着色", "brush-on-underline": "画笔下划线",
    "color-blend": "颜色混合", "color-wave": "颜色波浪",
    "complementary-color": "互补色", "complementary-color2": "互补色 2",
    "contrasting-color": "对比色", "darken": "变暗", "desaturate": "去饱和",
    "flash-bulb": "闪光灯", "flicker": "闪烁", "grow-with-color": "带颜色放大",
    "lighten": "变亮", "style-emphasis": "样式强调", "teeter": "跷跷板",
    "vertical-grow": "垂直放大", "wave": "波浪", "media-play": "播放媒体",
    "media-pause": "暂停媒体", "media-stop": "停止媒体",
}


def anim_zh(eid: str) -> str:
    if eid in ANIM_ZH:
        return ANIM_ZH[eid]
    if eid.startswith("path-"):
        shape = eid[5:]
        known = {
            "circle": "圆形", "square": "正方形", "diamond": "菱形", "hexagon": "六边形",
            "right-triangle": "直角三角形", "5-point-star": "五角星", "crescent-moon": "新月",
            "trapezoid": "梯形", "heart": "心形", "octagon": "八边形", "6-point-star": "六角星",
            "football": "足球", "equal-triangle": "等边三角形", "parallelogram": "平行四边形",
            "pentagon": "五边形", "4-point-star": "四角星", "8-point-star": "八角星",
            "teardrop": "泪滴", "pointy-star": "尖角星", "curved-square": "曲线方形",
            "curved-x": "曲线 X", "vertical-figure8": "垂直 8 字",
            "curvy-star": "曲线星", "loopde-loop": "环绕", "buzzsaw": "圆锯",
            "horizontal-figure8": "水平 8 字", "peanut": "花生", "figure8-four": "四重 8 字",
            "neutron": "中子", "swoosh": "嗖", "bean": "豆形", "plus": "十字",
            "inverted-triangle": "倒三角", "inverted-square": "倒方形", "left": "向左",
            "turn-right": "右转", "arc-down": "向下弧线", "zigzag": "锯齿",
            "s-curve2": "S 形曲线 2", "sine-wave": "正弦波", "bounce-left": "向左弹跳",
            "down": "向下", "turn-up": "上转", "arc-up": "向上弧线",
            "heartbeat": "心跳", "spiral-right": "右螺旋", "wave": "波浪",
            "curvy-left": "曲线向左", "diagonal-down-right": "右下斜线",
            "turn-down": "下转", "arc-left": "向左弧线", "funnel": "漏斗",
            "spring": "弹簧", "bounce-right": "向右弹跳", "spiral-left": "左螺旋",
            "diagonal-up-right": "右上斜线", "turn-up-right": "右上转",
            "arc-right": "向右弧线", "s-curve1": "S 形曲线 1",
            "decaying-wave": "衰减波浪", "curvy-right": "曲线向右",
            "stairs-down": "向下台阶", "up": "向上", "right": "向右",
        }
        shape_zh = known.get(shape, shape)
        return f"路径（{shape_zh}）"
    return eid


# ---------------------------------------------------------------- 切换

def strip_ns_text(xml: str) -> str:
    return re.sub(r'\sxmlns:[a-zA-Z0-9]+="[^"]*"', "", xml)


def effect_element(tr: dict) -> dict | None:
    if not tr or not tr.get("children"):
        return None
    root = etree.fromstring(tr["xml"].encode("utf-8"))
    kids = [c for c in root if isinstance(c.tag, str)]
    if not kids:
        return None
    child = kids[0]
    q = etree.QName(child)
    prefix = URI2PREFIX.get(q.namespace, "p")
    attrs = {k.split("}")[-1]: v for k, v in child.attrib.items()}
    return {"prefix": prefix, "local": q.localname, "attrs": attrs,
            "transition_attrs": {k.split("}")[-1]: v for k, v in root.attrib.items()}}


def kebab(enum_name: str, prefix: str) -> str:
    s = enum_name[len(prefix):]
    s = re.sub(r"(?<!^)(?=[A-Z])", "-", s)
    s = re.sub(r"-(\d)", r"\1", s)          # wheel-1-spoke 保留数字连写
    s = re.sub(r"-+", "-", s).strip("-").lower()
    return s


def build_transitions() -> dict:
    src = json.load(open(os.path.join(OUT, "transitions_map.json"), encoding="utf-8"))
    effects: dict[str, dict] = {}
    seen_sig: dict[tuple, str] = {}
    for key, v in src.items():
        if v.get("kind") != "enum":
            continue
        name = v["label"]
        if name in ("ppEffectMixed", "ppEffectNone"):
            continue
        el = effect_element(v["transition"])
        if not el:
            continue
        eid = kebab(name, "ppEffect")
        sig = (el["prefix"], el["local"], tuple(sorted(el["attrs"].items())))
        if sig in seen_sig:            # 同一 XML 的重复枚举，保留首个
            effects[seen_sig[sig]].setdefault("also", []).append(name)
            continue
        seen_sig[sig] = eid
        needs_alt = el["prefix"] in ("p14", "p15", "p159")
        effects[eid] = {
            "id": eid,
            "zh": zh_for(eid),
            "enum_name": name,
            "pp_value": v["value"],
            "element": f'{el["prefix"]}:{el["local"]}',
            "prefix": el["prefix"],
            "local": el["local"],
            "element_attrs": el["attrs"],
            "needs_alternate_content": needs_alt,
            "requires": el["prefix"] if needs_alt else None,
            "fallback": "p:fade" if needs_alt else None,
            "ground_truth_attrs": el["transition_attrs"],
        }
    # 兜底：ppEffectNone 记为 none（不写 transition）
    effects["none"] = {
        "id": "none", "zh": "无切换", "enum_name": "ppEffectNone", "pp_value": 0,
        "element": None, "prefix": None, "local": None, "element_attrs": {},
        "needs_alternate_content": False, "requires": None, "fallback": None,
        "ground_truth_attrs": {},
    }
    return {
        "meta": {
            "source": "PowerPoint 16.0 (build 20625) ground-truth sampling via COM SlideShowTransition",
            "generated_by": "rev/acquire_transitions.py + rev/acquire_transition_props.py",
            "rules": {
                "placement": "<p:transition> 位于 <p:sld> 内、<p:clrMapOvr> 之后、<p:timing> 之前",
                "p_elements": "p 命名空间元素可直接写裸 <p:transition>",
                "p14_p15_p159": "非 p 命名空间元素必须包在 mc:AlternateContent/mc:Choice Requires=\"p14|p15|p159\"；mc:Fallback 写 <p:transition><p:fade/></p:transition>",
                "p14_dur": "精确时长用属性 p14:dur（毫秒）；一旦使用就必须走 mc:AlternateContent Requires=\"p14\"，Fallback 为去掉该属性的同名 transition",
                "spd": "spd=slow|med|fast 对应 COM Speed=1/2/3；fast 为默认可省略",
                "fast_default": "Duration 与 Speed 互相覆盖，先设者被后者覆盖（COM 行为，写 PPT 时二选一）",
                "advTm": "自动换片毫秒；advClick=\"0\" 关闭点击换片",
                "empty": "无效果但需要换片条件时写空 <p:transition spd=\"...\"/>",
            },
        },
        "effects": effects,
    }


# ---------------------------------------------------------------- 动画

def behavior_templates(timing_xml: str) -> list[str]:
    if not timing_xml:
        return []
    root = etree.fromstring(timing_xml.encode("utf-8"))
    out: list[str] = []
    for ctn in root.iter("{%s}cTn" % NSMAP["p"]):
        if ctn.get("presetID") is None:
            continue
        child_lst = ctn.find("{%s}childTnLst" % NSMAP["p"])
        if child_lst is None:
            continue
        for beh in child_lst:
            if not isinstance(beh.tag, str):
                continue
            chunk = etree.tostring(beh, encoding="unicode", with_tail=False)
            chunk = strip_ns_text(chunk)
            chunk = re.sub(r'spid="\d+"', 'spid="{{SPID}}"', chunk)
            out.append(re.sub(r"\s+", " ", chunk).strip())
    return out


def preset_of(rec: dict) -> dict | None:
    for p in (rec.get("presets") or []):
        if p["presetClass"] in ("entr", "emph", "exit", "path"):
            return p
    return None


def build_animations() -> dict:
    base = json.load(open(os.path.join(OUT, "animations_map.json"), encoding="utf-8"))
    exit_path = os.path.join(OUT, "anim_exit_map.json")
    exits = json.load(open(exit_path, encoding="utf-8")) if os.path.exists(exit_path) else {}
    items: dict[str, dict] = {}
    for label, rec in base.items():
        p = preset_of(rec)
        if not p or not rec.get("timing_xml"):
            continue
        eid = kebab(label, "msoAnimEffect")
        entry = {
            "id": eid,
            "zh": anim_zh(eid),
            "enum_name": label,
            "pp_value": rec["value"],
            "preset": {"presetID": p["presetID"], "presetClass": p["presetClass"],
                       "presetSubtype": p["presetSubtype"]},
            "behaviors": behavior_templates(rec["timing_xml"]),
            "variants": {},
        }
        ex = exits.get(label)
        if ex and ex.get("timing_xml"):
            ep = preset_of(ex)
            if ep and ep["presetClass"] == "exit":
                entry["variants"]["exit"] = {
                    "preset": {"presetID": ep["presetID"], "presetClass": ep["presetClass"],
                               "presetSubtype": ep["presetSubtype"]},
                    "behaviors": behavior_templates(ex["timing_xml"]),
                }
        items[eid] = entry
    return {
        "meta": {
            "source": "PowerPoint 16.0 ground-truth sampling via COM MainSequence.AddEffect",
            "generated_by": "rev/acquire_animations.py + rev/acquire_anim_exit.py + rev/acquire_anim_variants.py",
            "preset_class": {"entr": "进入", "emph": "强调", "exit": "退出", "path": "动作路径"},
            "trigger_to_node": {
                "click": {"nodeType": "clickEffect", "cond": 'delay="0"'},
                "with_previous": {"nodeType": "withEffect", "cond": 'evt="onBegin" delay="0"'},
                "after_previous": {"nodeType": "afterEffect", "cond": 'evt="onBegin" delay="0"'},
            },
            "placeholders": {"{{SPID}}": "形状 id", "{{DUR}}": "行为时长（毫秒）"},
            "units": {"dur": "毫秒", "delay": "毫秒", "repeatCount": "次数×1000",
                      "accel/decel": "百分比×1000", "smooth": "SmoothStart/SmoothEnd 秒"},
            "structure": "p:timing/p:tnLst/p:par/p:cTn(tmRoot)/p:childTnLst/p:seq/p:cTn(mainSeq)/p:childTnLst/p:par(click 组)/.../p:par(效果)/p:cTn(presetID...)",
        },
        "effects": items,
    }


def main() -> int:
    tr = build_transitions()
    an = build_animations()
    for name, obj in (("transitions.json", tr), ("animations.json", an)):
        with open(os.path.join(SCHEMA, name), "w", encoding="utf-8") as fh:
            json.dump(obj, fh, ensure_ascii=False, indent=1)
    print(f"[transitions] {len(tr['effects'])}")
    for eid, e in tr["effects"].items():
        print(f"   {eid:24s} {e['zh']:12s} {str(e['element']):16s} {json.dumps(e['element_attrs'], ensure_ascii=False)}")
    print(f"\n[animations] {len(an['effects'])}")
    for eid, e in list(an["effects"].items()):
        ex = "exit" if e["variants"].get("exit") else "   "
        print(f"   {eid:24s} {e['zh']:14s} pid={e['preset']['presetID']:>4s} {e['preset']['presetClass']:5s} {ex} beh={len(e['behaviors'])}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())