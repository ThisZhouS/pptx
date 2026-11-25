"""读取 pptx 的切换 / 动画 / 形状结构，用于校验与展示。"""
from __future__ import annotations

import os
import zipfile

from lxml import etree

from .ns import NS, q, to_xml


def _slide_elements(path: str):
    """按页序产出 (index, slide_xml_root)。"""
    with zipfile.ZipFile(path) as zf:
        pres = etree.fromstring(zf.read("ppt/presentation.xml"))
        rels = etree.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
        rel_map = {r.get("Id"): r.get("Target") for r in rels}
        for i, sld in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), start=1):
            target = rel_map[sld.get(q("r:id"))]
            part = target if target.startswith("ppt/") else os.path.normpath(os.path.join("ppt", target)).replace("\\", "/")
            yield i, etree.fromstring(zf.read(part))


def inspect_deck(path: str, with_xml: bool = False) -> dict:
    from .ooxml import read_timing, read_transition

    slides = []
    for idx, slde in _slide_elements(path):
        c_sld = slde.find(q("p:cSld"))
        shapes = []
        if c_sld is not None:
            tree = c_sld.find(q("p:spTree"))
            if tree is not None:
                for sp in tree:
                    if not isinstance(sp.tag, str):
                        continue
                    local = etree.QName(sp).localname
                    if local not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                        continue
                    nv = sp.find(".//" + q("p:cNvPr"))
                    if nv is None:
                        continue
                    text = "".join(sp.itertext()).strip()[:40]
                    shapes.append({"id": int(nv.get("id")), "name": nv.get("name"),
                                   "tag": local, "text": text})
        entry = {
            "index": idx,
            "shapes": shapes,
            "transition": read_transition(slde),
            "animations": read_timing(slde),
        }
        if with_xml:
            entry["xml"] = to_xml(slde)
        slides.append(entry)
    return {"file": os.path.abspath(path), "slide_count": len(slides), "slides": slides}


def validate_deck(path: str) -> dict:
    """结构级校验：zip 完整性、XML 可解析、动画引用形状是否存在、transition 位置。"""
    problems: list[str] = []
    info = {"file": os.path.abspath(path), "problems": problems, "stats": {}}
    if not os.path.exists(path):
        problems.append("文件不存在")
        return info
    try:
        with zipfile.ZipFile(path) as zf:
            bad = zf.testzip()
            if bad:
                problems.append(f"zip 损坏: {bad}")
            names = set(zf.namelist())
            for required in ("ppt/presentation.xml", "[Content_Types].xml"):
                if required not in names:
                    problems.append(f"缺少必需部件: {required}")
    except zipfile.BadZipFile:
        problems.append("不是合法的 zip/pptx")
        return info

    total_anim = 0
    trans_count = 0
    page_count = 0
    for idx, slde in _slide_elements(path):
        page_count = idx
        try:
            etree.tostring(slde)
        except Exception as exc:  # noqa: BLE001
            problems.append(f"第 {idx} 页 XML 解析失败: {exc}")
            continue
        shape_ids = {
            int(nv.get("id"))
            for nv in slde.iter(q("p:cNvPr"))
        }
        timing = slde.find(q("p:timing"))
        if timing is not None:
            for tgt in timing.iter(q("p:spTgt")):
                try:
                    sid = int(tgt.get("spid"))
                except (TypeError, ValueError):
                    problems.append(f"第 {idx} 页动画目标 spid 非法")
                    continue
                if sid not in shape_ids:
                    problems.append(f"第 {idx} 页动画引用了不存在的形状 id={sid}")
            total_anim += sum(1 for ctn in timing.iter(q("p:cTn")) if ctn.get("presetID"))
        if slde.find(q("p:transition")) is not None or any(
                isinstance(c.tag, str) and c.tag == q("mc:AlternateContent") for c in slde):
            trans_count += 1
        # 顺序检查：transition 必须在 clrMapOvr 之后、timing 之前
        order = [etree.QName(c).localname for c in slde if isinstance(c.tag, str)]
        if "transition" in order and "timing" in order:
            if order.index("transition") > order.index("timing"):
                problems.append(f"第 {idx} 页 transition 位于 timing 之后（违反 schema 顺序）")
        if "AlternateContent" in order and "timing" in order:
            if order.index("AlternateContent") > order.index("timing"):
                problems.append(f"第 {idx} 页 AlternateContent 位于 timing 之后")
    info["stats"] = {"slides": page_count,
                     "pages_with_transition": trans_count,
                     "animation_effects": total_anim,
                     "problems": len(problems)}
    return info