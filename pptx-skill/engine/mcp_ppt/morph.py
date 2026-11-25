"""Morph 关键帧引擎：用"多页元素变换 + 平滑切换 + 自动换片"做复杂动画。

思路
----
对象动画（p:timing）表达能力有限：一条时间轴、单向、必须点击或等前一条结束。
要做"元素从 A 漂移到 B，途中放大变色，再缩小归位"这类叙事动画，更合适的做法是
**把每一帧画成一页，页间用 morph（p159:morph）平滑插值，并用 advTm 自动换片**：

    页1（起始） --morph 1.2s--> 页2（中间点1） --morph--> 页3（中间点2） --> 页4（结束）

PowerPoint 的 morph 依据"名称相同"匹配对象，名称以 ``!!`` 开头时强制按名称匹配，
因此同一逻辑元素在每页用同一个 ``!!name``，插值就会发生在它自己身上。

真机验证（rev/acquire_morph.py）：
  · 连续 3 页 <p159:morph> + p14:dur=1200 + advTm=1500 被 PowerPoint 完全接受
  · EntryEffect 回读 3954（ppEffectMorphByObject）
  · 元素名 ``!!hero`` / ``!!side`` 原样保留
"""
from __future__ import annotations

import os
from dataclasses import dataclass, field

from pptx import Presentation
from pptx.util import Inches

from .catalog import get_catalog
from .design import SLIDE_H, SLIDE_W, Theme
from .ooxml import set_slide_transition
from .shapes import add_element


@dataclass
class ElementPlan:
    name: str
    kind: str = "round-rect"
    base: dict = field(default_factory=dict)
    keyframes: list[dict] = field(default_factory=list)

    def state_at(self, index: int) -> dict:
        """第 index 帧的状态：基础值 + 截至该帧最近一次的关键帧覆盖。"""
        state = dict(self.base)
        for kf in self.keyframes[: index + 1]:
            state.update({k: v for k, v in kf.items() if k != "at"})
        return state


def _frame_count(elements: list[ElementPlan]) -> int:
    return max([len(e.keyframes) for e in elements] + [1])


def plan(model: dict) -> dict:
    """展开成逐帧计划（不落盘），便于先审后建。"""
    elements = [ElementPlan(name=e["name"], kind=e.get("kind", "round-rect"),
                            base={k: v for k, v in e.items()
                                  if k not in ("name", "kind", "keyframes")},
                            keyframes=e.get("keyframes") or [])
                for e in model.get("elements") or []]
    base_elements = model.get("base", {}).get("elements") or []
    frames = _frame_count(elements)
    durations = model.get("frame_duration") or model.get("frame_durations") or 2.5
    if isinstance(durations, (int, float)):
        durations = [durations] * frames
    while len(durations) < frames:
        durations.append(durations[-1] if durations else 2.5)

    timeline = []
    t = 0.0
    for i in range(frames):
        state = {e.name: e.state_at(i) for e in elements}
        timeline.append({"frame": i + 1, "start": round(t, 2),
                         "duration": float(durations[i]),
                         "elements": state})
        t += float(durations[i])
    return {"frames": frames, "timeline": timeline,
            "base": base_elements,
            "transition": model.get("transition") or {"effect": "morph", "duration": 2.0},
            "total_duration": round(t, 2)}


def build_morph_deck(model: dict, out_path: str) -> dict:
    """按关键帧生成 morph 序列（N 页 = 关键帧数）。"""
    cat = get_catalog()
    theme_value = model.get("theme")
    theme = Theme.preset(theme_value if isinstance(theme_value, str) else
                         (theme_value or {}).get("preset", "dark")
                         if isinstance(theme_value, dict) else "dark")
    if isinstance(theme_value, dict):
        overrides = {k: v for k, v in theme_value.items()
                     if k in ("bg", "bg_alt", "card", "fg", "muted", "accent", "accent_2",
                              "font_title", "font_body")}
        if overrides:
            from dataclasses import replace as _replace
            theme = _replace(theme, **overrides)

    p = plan(model)
    trans = dict(p["transition"])
    if str(trans.get("effect", "morph")).lower() in ("morph", "变形"):
        trans["effect"] = "morph"
    option = model.get("morph_option") or model.get("option")
    if option and trans["effect"] == "morph-by-object":
        pass

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]
    SW, SH = int(SLIDE_W), int(SLIDE_H)

    background = model.get("background") or theme.bg
    report_slides = []
    for frame in p["timeline"]:
        slide = prs.slides.add_slide(blank)
        fill = slide.background.fill
        fill.solid()
        from pptx.dml.color import RGBColor
        fill.fore_color.rgb = RGBColor.from_string(str(background).lstrip("#").upper())

        drawn = []
        # 背景底板（可选）
        if model.get("backdrop", True):
            add_element(slide, {"name": "!!backdrop", "kind": "rect", "x": 0, "y": 0,
                                "w": "100%", "h": "100%", "fill": background},
                        theme, SW, SH)
        # 每帧都存在的静态元素（同名 → morph 会保持不动）
        for spec in p["base"]:
            base_spec = dict(spec)
            base_spec["name"] = f"!!{spec['name']}" if not str(spec["name"]).startswith("!!") else spec["name"]
            add_element(slide, base_spec, theme, SW, SH)
            drawn.append(base_spec["name"])
        # 关键帧元素
        for el in model.get("elements") or []:
            state = frame["elements"][el["name"]]
            spec = {"kind": el.get("kind", "round-rect"), **state}
            name = str(el["name"])
            spec["name"] = name if name.startswith("!!") else f"!!{name}"
            if spec.get("visible") is False:
                # 用全透明占位，保证 morph 仍能匹配到这个对象
                spec.setdefault("alpha", 0.0)
                if "fill" not in spec:
                    spec["fill"] = "#FFFFFF"
            add_element(slide, spec, theme, SW, SH)
            drawn.append(spec["name"])

        adv = float(frame["duration"])
        applied = set_slide_transition(slide.element, {**trans, "advance_after": adv}, cat)
        report_slides.append({"frame": frame["frame"], "start": frame["start"],
                              "duration": adv, "shapes": drawn,
                              "transition": applied})

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    prs.save(out_path)
    return {"file": os.path.abspath(out_path), "frames": p["frames"],
            "total_duration": p["total_duration"], "transition": trans,
            "slides": report_slides,
            "timeline": p["timeline"]}


def decompose(path: str) -> dict:
    """读取一页或多页的元素清单 + 指纹，用于给"同一元素"编同一个码。"""
    import re
    import zipfile

    from lxml import etree

    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    pages = []
    with zipfile.ZipFile(path) as zf:
        parts = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                       key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for idx, part in enumerate(parts, start=1):
            root = etree.fromstring(zf.read(part))
            tree = root.find(".//{%s}spTree" % P)
            items = []
            for sp in tree if tree is not None else []:
                if not isinstance(sp.tag, str):
                    continue
                local = etree.QName(sp).localname
                if local not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                    continue
                nv = sp.find(".//{%s}cNvPr" % P)
                if nv is None:
                    continue
                text = "".join(sp.itertext()).strip()
                xfrm = sp.find(".//{%s}xfrm" % A)
                box = None
                if xfrm is not None:
                    off, ext = xfrm.find("{%s}off" % A), xfrm.find("{%s}ext" % A)
                    if off is not None and ext is not None:
                        box = [round(int(off.get("x")) / 914400, 2), round(int(off.get("y")) / 914400, 2),
                               round(int(ext.get("cx")) / 914400, 2), round(int(ext.get("cy")) / 914400, 2)]
                items.append({"id": int(nv.get("id")), "name": nv.get("name"), "tag": local,
                              "text": text[:40], "box_in": box,
                              "coded": str(nv.get("name")).startswith("!!")})
            pages.append({"slide": idx, "shapes": items})

    # 跨页同名候选（morph 可插值的对象）
    seen: dict[str, list[int]] = {}
    for page in pages:
        for s in page["shapes"]:
            key = str(s["name"])
            if s["coded"]:
                seen.setdefault(key, []).append(page["slide"])
    return {"file": os.path.abspath(path), "pages": len(pages), "pages_detail": pages,
            "coded_across_pages": {k: v for k, v in seen.items() if len(v) > 1}}


def assign_codes(path: str, mapping: dict[str, str]) -> dict:
    """把若干元素的名称改成 !!code，便于 morph 跨页匹配。

    mapping: {"第1页的旧名或 shape:12": "hero", ...}；同名元素在各自页上都会改。
    """
    from pptx import Presentation

    prs = Presentation(path)
    changed = []
    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shp in slide.shapes:
            old = shp.name
            for key, code in mapping.items():
                if key in (old, f"shape:{shp.shape_id}", str(shp.shape_id)):
                    new = code if str(code).startswith("!!") else f"!!{code}"
                    shp.name = new
                    changed.append({"slide": slide_idx, "id": shp.shape_id,
                                    "from": old, "to": new})
    if changed:
        prs.save(path)
    return {"file": os.path.abspath(path), "changed": changed}
