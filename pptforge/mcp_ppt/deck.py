"""Deck 构建器：一份 JSON 描述 → 可播放的 .pptx（含切换与对象动画）。

模型（ai 生成时按这个结构写）
---------------------------
{
  "theme": "dark" | "light" | "midnight" | "warm" | {"bg": "#0B1020", ...},
  "defaults": {"transition": {"effect": "fade", "duration": 0.7}},
  "slides": [
    {"layout": "title", "title": "...", "subtitle": "...", "meta": ["2026"]},
    {"layout": "bullets", "title": "...", "bullets": ["a", {"text": "b", "level": 1}],
     "transition": {"effect": "cover-left", "duration": 0.6},
     "animations": [
        {"target": "title", "effect": "fade", "trigger": "click", "duration": 0.5},
        {"target": "bullets", "effect": "fly", "trigger": "after_previous", "build": "paragraph"}
     ]}
  ]
}
"""
from __future__ import annotations

import os
from dataclasses import replace

from pptx import Presentation
from pptx.util import Inches

from .catalog import get_catalog
from .design import (Painter, SLIDE_H, SLIDE_W, ShapeMeta, Theme, resolve_layout)
from .ns import to_xml
from .ooxml import AnimItem, read_timing, read_transition, set_slide_timing, set_slide_transition
from .shapes import add_element, apply_ops, find_shape, group_shapes, iter_shapes
from .theme import apply_theme, extract_theme, set_show


def _theme(value) -> Theme:
    if isinstance(value, dict):
        base = Theme.preset(value.get("preset") or value.get("name") or "dark")
        overrides = {k: v for k, v in value.items()
                     if k in ("bg", "bg_alt", "card", "fg", "muted", "accent", "accent_2",
                              "font_title", "font_body")}
        colors = value.get("colors") or {}
        if colors.get("accent1"):
            overrides.setdefault("accent", str(colors["accent1"]))
        if colors.get("accent2"):
            overrides.setdefault("accent_2", str(colors["accent2"]))
        fonts = value.get("fonts") or {}
        if fonts.get("minor"):
            overrides.setdefault("font_body", str(fonts["minor"]))
        if fonts.get("major"):
            overrides.setdefault("font_title", str(fonts["major"]))
        return replace(base, **overrides) if overrides else base
    return Theme.preset(value)


def _resolve_targets(target, shapes: dict[str, ShapeMeta]) -> tuple[list[ShapeMeta], list[str]]:
    """解析动画目标；未知目标返回在 unknown 里（不抛错，由调用方记 warning）。"""
    if target is None:
        return [], []
    names = target if isinstance(target, (list, tuple)) else [target]
    out: list[ShapeMeta] = []
    unknown: list[str] = []
    for name in names:
        key = str(name).strip()
        if key in ("*", "all", "全部"):
            out.extend([m for m in shapes.values() if m.shape_id > 0 and m.kind != "error"])
            continue
        # group:<name> / name:<name> 前缀
        lookup = key
        if ":" in key:
            head, tail = key.split(":", 1)
            if head.lower() in ("group", "name"):
                lookup = tail.strip()
        if lookup in shapes:
            out.append(shapes[lookup])
            continue
        low = lookup.lower()
        hit = [m for n, m in shapes.items() if n.lower() == low and m.shape_id > 0]
        if hit:
            out.extend(hit)
            continue
        if key.lower().startswith("shape:"):
            try:
                sid = int(key.split(":", 1)[1])
            except ValueError:
                unknown.append(key)
                continue
            found = [m for m in shapes.values() if m.shape_id == sid]
            if found:
                out.extend(found)
            else:
                unknown.append(key)
            continue
        unknown.append(key)
    # 去重保序
    seen, uniq = set(), []
    for m in out:
        if m.shape_id not in seen:
            seen.add(m.shape_id)
            uniq.append(m)
    return uniq, unknown


def _anim_items(anim_specs, shapes: dict[str, ShapeMeta], cat) -> tuple[list[AnimItem], list[str]]:
    items: list[AnimItem] = []
    warnings: list[str] = []
    stagger_acc = 0.0
    for spec in anim_specs or []:
        if not isinstance(spec, dict):
            continue
        trigger = cat.resolve_trigger(spec.get("trigger") or spec.get("when"))
        effect = spec.get("effect") or spec.get("name") or "fade"
        is_exit = bool(spec.get("exit")) or str(spec.get("kind", "")).lower() in ("exit", "退出")
        duration = spec.get("duration", 0.8)
        base_delay = float(spec.get("delay") or 0.0)
        stagger = float(spec.get("stagger") or 0.0)
        repeat = spec.get("repeat")
        smooth_start = spec.get("smooth_start")
        smooth_end = spec.get("smooth_end")
        build = str(spec.get("build") or spec.get("by") or "all").lower()
        targets, unknown = _resolve_targets(spec.get("target") or spec.get("targets"), shapes)
        if unknown:
            available = [k for k, v in shapes.items() if v.shape_id > 0]
            warnings.append(f"动画 {effect!r} 的目标 {unknown} 不存在，已跳过；本页可用: {available}")
        if not targets:
            if not unknown:
                warnings.append(f"动画 {effect!r} 没有匹配到目标形状")
            continue
        if trigger == "click":
            stagger_acc = 0.0
        for meta in targets:
            delay = base_delay
            if trigger == "with_previous" and stagger:
                stagger_acc += stagger
                delay += stagger_acc
            elif trigger != "with_previous":
                stagger_acc = 0.0
            if build in ("paragraph", "段落", "para") and meta.paragraphs > 1:
                for i in range(meta.paragraphs):
                    items.append(AnimItem(
                        shape_id=meta.shape_id, effect=effect, trigger=trigger,
                        duration=duration, delay=delay + i * 0.1, repeat=repeat,
                        smooth_start=smooth_start, smooth_end=smooth_end,
                        exit=is_exit, paragraph=i, paragraph_end=i,
                        raw=spec))
            else:
                items.append(AnimItem(
                    shape_id=meta.shape_id, effect=effect, trigger=trigger,
                    duration=duration, delay=delay, repeat=repeat,
                    smooth_start=smooth_start, smooth_end=smooth_end,
                    exit=is_exit, raw=spec))
    # 规范化：组内第一个效果若是 with/after，仍保留原触发（与 PowerPoint 行为一致）
    for i, item in enumerate(items):
        if i == 0 and item.trigger != "click":
            warnings.append("首个动画设为 with/after，已按 PowerPoint 行为保留（组同时可被点击触发）")
            break
    return items, warnings


def _apply_elements(build, specs: list[dict], slide_w: int, slide_h: int) -> tuple[list[dict], list[dict]]:
    """创建自定义元素，并按 element.group 字段自动组合；返回 (创建记录, 组合记录)。"""
    added: list[dict] = []
    buckets: dict[str, list] = {}
    for spec in specs:
        rec = add_element(build.slide, spec, slide_w=slide_w, slide_h=slide_h)
        shp = rec.pop("_shape", None)
        if shp is not None:
            paragraphs = len(shp.text_frame.paragraphs) if getattr(shp, "has_text_frame", False) else 1
            build.add(rec["name"], shp, kind="element", paragraphs=paragraphs)
        added.append(rec)
        group_name = spec.get("group")
        if group_name and shp is not None:
            buckets.setdefault(str(group_name), []).append(shp)
    grouped = []
    for gname, members in buckets.items():
        if len(members) < 2:
            continue
        try:
            info = group_shapes(build.slide, members, name=gname)
            grp = find_shape(build.slide, f"group:{gname}")
            build.shapes[f"group:{gname}"] = ShapeMeta(f"group:{gname}", grp.shape_id, "group", 1)
            build.shapes.setdefault(gname, build.shapes[f"group:{gname}"])
            grouped.append(info)
        except Exception as exc:  # noqa: BLE001
            grouped.append({"group": gname, "error": f"{type(exc).__name__}: {exc}"})
    return added, grouped


def _sync_shapes(build) -> None:
    """把 ops 之后新增/改名的形状同步进 build.shapes，供动画和后续 op 引用。"""
    for shp in iter_shapes(build.slide.shapes):
        key = shp.name
        if not key:
            continue
        if key not in build.shapes:
            paragraphs = len(shp.text_frame.paragraphs) if getattr(shp, "has_text_frame", False) else 1
            build.shapes[key] = ShapeMeta(key, shp.shape_id, "shape", paragraphs)


def build_deck(model: dict, out_path: str) -> dict:
    cat = get_catalog()
    theme = _theme(model.get("theme"))
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    painter = Painter(prs, theme)

    defaults = model.get("defaults") or {}
    default_transition = defaults.get("transition")
    report = {"file": os.path.abspath(out_path), "theme": theme.name, "slides": [], "warnings": []}

    # ---- 主题部件（theme1.xml）与放映设置（presProps.xml）
    theme_spec = model.get("theme")
    theme_from = model.get("theme_from")
    if isinstance(theme_spec, dict) or theme_from:
        colors = dict(theme_spec.get("colors") or {}) if isinstance(theme_spec, dict) else {}
        fonts = dict(theme_spec.get("fonts") or {}) if isinstance(theme_spec, dict) else {}
        if theme_from:
            ext = extract_theme(theme_from)
            colors = {**ext.get("colors", {}), **colors}
            fonts = {**ext.get("fonts", {}), **fonts}
            report["theme_from"] = {"source": theme_from, "colors": len(ext.get("colors", {}))}
        report["theme_applied"] = apply_theme(
            prs, colors, fonts,
            name=(theme_spec.get("name") if isinstance(theme_spec, dict) else None))
    if model.get("show"):
        report["show"] = set_show(prs, model["show"]).get("show")

    slides = model.get("slides") or []
    if not slides:
        raise ValueError("slides 为空：至少需要一页")

    for idx, sspec in enumerate(slides, start=1):
        sspec = dict(sspec)
        layout = resolve_layout(sspec.get("layout") or sspec.get("type"))
        sspec.setdefault("index", idx)
        if layout != "title" and not sspec.get("footer") and sspec.get("show_footer", True) and model.get("footer"):
            sspec["footer"] = model["footer"]
        build = getattr(painter, layout)(sspec)
        slide_w, slide_h = int(prs.slide_width), int(prs.slide_height)

        # ---- 自定义元素 / 组合 / 形状操作（在动效之前，动效按最终几何与名字绑定）
        elements_meta, groups_meta = [], []
        if sspec.get("elements"):
            elements_meta, groups_meta = _apply_elements(build, sspec["elements"], slide_w, slide_h)
        ops_meta = []
        if sspec.get("ops"):
            ops_meta = apply_ops(build.slide, sspec["ops"], slide_w, slide_h)
            for rec in ops_meta:
                if not rec.get("ok"):
                    report["warnings"].append(f"第 {idx} 页 ops[{rec['index']}] {rec['op']}: {rec.get('error')}")
        if sspec.get("elements") or sspec.get("ops"):
            _sync_shapes(build)

        # ---- 切换
        trans_spec = sspec.get("transition", default_transition)
        trans_applied = None
        if trans_spec is not None:
            if isinstance(trans_spec, str):
                trans_spec = {"effect": trans_spec}
            trans_applied = set_slide_transition(build.slide.element, trans_spec, cat)
            for w in trans_applied.get("warnings", []):
                report["warnings"].append(f"第 {idx} 页切换: {w}")

        # ---- 动画
        items, warns = _anim_items(sspec.get("animations"), build.shapes, cat)
        anim_meta = set_slide_timing(build.slide.element, items, cat) if items else {"items": 0}
        for w in warns:
            report["warnings"].append(f"第 {idx} 页动画: {w}")

        entry = {
            "index": idx,
            "layout": layout,
            "title": sspec.get("title"),
            "shapes": {n: {"id": m.shape_id, "kind": m.kind, "paragraphs": m.paragraphs}
                       for n, m in build.shapes.items() if m.shape_id > 0},
            "transition": trans_applied,
            "animations": anim_meta,
        }
        if elements_meta:
            entry["elements"] = elements_meta
        if groups_meta:
            entry["groups"] = groups_meta
        if ops_meta:
            entry["ops"] = ops_meta
        report["slides"].append(entry)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    prs.save(out_path)
    report["slide_count"] = len(prs.slides)
    return report


def patch_deck(path: str, patches: list[dict]) -> dict:
    """对已存在的 pptx 逐页编辑：元素/形状操作/切换/动画（不重建内容）。"""
    from .design import SlideBuild

    cat = get_catalog()
    prs = Presentation(path)
    slide_w, slide_h = int(prs.slide_width), int(prs.slide_height)
    report = {"file": os.path.abspath(path), "patched": [], "warnings": []}
    for patch in patches:
        index = int(patch.get("slide", patch.get("index", 1)))
        if not 1 <= index <= len(prs.slides):
            report["warnings"].append(f"页号越界: {index}")
            continue
        slide = prs.slides[index - 1]
        applied = {}

        build = SlideBuild(slide)
        for shape in iter_shapes(slide.shapes):
            key = shape.name or f"shape{shape.shape_id}"
            build.shapes[key] = ShapeMeta(
                key, shape.shape_id, "shape",
                paragraphs=len(shape.text_frame.paragraphs) if shape.has_text_frame else 1)

        if patch.get("elements"):
            elements_meta, groups_meta = _apply_elements(build, patch["elements"], slide_w, slide_h)
            applied["elements"] = elements_meta
            if groups_meta:
                applied["groups"] = groups_meta
        if patch.get("ops"):
            ops_meta = apply_ops(slide, patch["ops"], slide_w, slide_h)
            applied["ops"] = ops_meta
            for rec in ops_meta:
                if not rec.get("ok"):
                    report["warnings"].append(
                        f"第 {index} 页 ops[{rec['index']}] {rec['op']}: {rec.get('error')}")
        if patch.get("elements") or patch.get("ops"):
            _sync_shapes(build)

        if patch.get("transition") is not None:
            spec = patch["transition"]
            if isinstance(spec, str):
                spec = {"effect": spec}
            applied["transition"] = set_slide_transition(slide.element, spec, cat)
        if "animations" in patch:
            items, warns = _anim_items(patch["animations"], build.shapes, cat)
            # 空列表也要走一遍：set_slide_timing 会先移除旧 timing，实现"清空动画"
            applied["animations"] = set_slide_timing(slide.element, items, cat)
            report["warnings"].extend(f"第 {index} 页: {w}" for w in warns)
        report["patched"].append({"index": index, "applied": applied})
    prs.save(path)
    return report


def dump_slide_xml(path: str, index: int) -> str:
    prs = Presentation(path)
    slide = prs.slides[index - 1]
    return to_xml(slide.element)
