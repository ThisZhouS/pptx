"""OOXML 解包与命名空间工具（逆向层公用）。"""
from __future__ import annotations

import json
import os
import zipfile
from typing import Iterable

from lxml import etree

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
    "p16": "http://schemas.microsoft.com/office/powerpoint/2016/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def qn(tag: str) -> str:
    """'p:transition' -> '{ns}transition'"""
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


def slide_parts(zf: zipfile.ZipFile) -> list[tuple[int, str]]:
    """按 presentation.xml 的 sldIdLst 顺序返回 [(1-based index, part name)]。"""
    pres = etree.fromstring(zf.read("ppt/presentation.xml"))
    rels = etree.fromstring(zf.read("ppt/_rels/presentation.xml.rels"))
    rel_map = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for i, sld in enumerate(pres.findall("p:sldIdLst/p:sldId", NS), start=1):
        rid = sld.get(qn("r:id"))
        target = rel_map[rid]
        part = target if target.startswith("ppt/") else os.path.normpath(os.path.join("ppt", target)).replace("\\", "/")
        out.append((i, part))
    return out


def pretty(elem) -> str:
    return etree.tostring(elem, pretty_print=True, encoding="unicode")


def dump_pptx(path: str, outdir: str, what: str = "slides") -> list[dict]:
    """解包把所有 slide XML 落盘，返回 [{index, part, xml}]。"""
    os.makedirs(outdir, exist_ok=True)
    result = []
    with zipfile.ZipFile(path) as zf:
        for idx, part in slide_parts(zf):
            raw = zf.read(part)
            tree = etree.fromstring(raw)
            text = pretty(tree)
            fn = os.path.join(outdir, f"slide_{idx:03d}.xml")
            with open(fn, "w", encoding="utf-8") as fh:
                fh.write(text)
            result.append({"index": idx, "part": part, "xml": text})
    return result


def find_transition(xml_text: str) -> dict | None:
    """从一个 slide XML 里定位 <p:transition>（含 mc:AlternateContent 包裹）并拆出属性。"""
    root = etree.fromstring(xml_text.encode("utf-8"))
    tr = root.find("p:transition", NS)
    holder = "p:sld"
    if tr is None:
        alt = root.find("mc:AlternateContent", NS)
        if alt is not None:
            choice = alt.find("mc:Choice", NS)
            if choice is not None:
                tr = choice.find("p:transition", NS)
                holder = "mc:AlternateContent/mc:Choice"
    if tr is None:
        return None
    children = [etree.QName(c).localname for c in tr if isinstance(c.tag, str)]
    ns_of_children = [etree.QName(c).namespace for c in tr if isinstance(c.tag, str)]
    info = {
        "attrs": {k.split("}")[-1]: v for k, v in tr.attrib.items()},
        "children": children,
        "child_ns": ns_of_children,
        "holder": holder,
        "xml": etree.tostring(tr, pretty_print=True, encoding="unicode"),
    }
    return info


def walk_timing(xml_text: str) -> dict | None:
    """定位 <p:timing>，返回原始 XML 与结构指纹（节点路径计数）。"""
    root = etree.fromstring(xml_text.encode("utf-8"))
    tm = root.find("p:timing", NS)
    if tm is None:
        return None
    paths: dict[str, int] = {}
    for el in tm.iter():
        if not isinstance(el.tag, str):
            continue
        chain = []
        node = el
        while node is not None and node is not tm:
            chain.append(etree.QName(node).localname)
            node = node.getparent()
        chain.append("p:timing")
        key = "/".join(reversed(chain))
        paths[key] = paths.get(key, 0) + 1
    behaviours = [dict(el.attrib) for el in tm.iter(qn("p:cBhvr"))]
    effects = [dict(el.attrib) for el in tm.iter(qn("p:animEffect"))]
    presets = [
        {
            "node": etree.QName(el.getparent()).localname,
            "presetID": el.get("presetID"),
            "presetClass": el.get("presetClass"),
            "presetSubtype": el.get("presetSubtype"),
        }
        for el in tm.iter()
        if isinstance(el.tag, str) and el.get("presetClass") is not None
    ]
    return {
        "xml": etree.tostring(tm, pretty_print=True, encoding="unicode"),
        "paths": paths,
        "bhvr_count": len(behaviours),
        "anim_effects": effects,
        "presets": presets,
    }


def save_json(obj, path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, ensure_ascii=False, indent=1)


def _xfrm_info(node) -> dict:
    """从 a:xfrm（或 p:grpSpPr/a:xfrm）取偏移、尺寸、旋转、翻转、子坐标系。"""
    xfrm = node.find(".//" + qn("a:xfrm"))
    info = {}
    if xfrm is None:
        return info
    off = xfrm.find(qn("a:off"))
    ext = xfrm.find(qn("a:ext"))
    choff = xfrm.find(qn("a:chOff"))
    chext = xfrm.find(qn("a:chExt"))
    if off is not None:
        info["off"] = [int(off.get("x")), int(off.get("y"))]
    if ext is not None:
        info["ext"] = [int(ext.get("cx")), int(ext.get("cy"))]
    if choff is not None:
        info["chOff"] = [int(choff.get("x")), int(choff.get("y"))]
    if chext is not None:
        info["chExt"] = [int(chext.get("cx")), int(chext.get("cy"))]
    if xfrm.get("rot"):
        info["rot"] = int(xfrm.get("rot"))
    flips = [k for k in ("flipH", "flipV") if xfrm.get(k) in ("1", "true")]
    if flips:
        info["flip"] = flips
    return info


def shape_tree(xml_text: str) -> list[dict]:
    """递归解析 spTree：形状/图片/组合的 id、名称、几何与子坐标系。"""
    root = etree.fromstring(xml_text.encode("utf-8"))
    tree = root.find(".//" + qn("p:spTree"))
    if tree is None:
        return []

    def walk(container, depth=0):
        out = []
        for node in container:
            if not isinstance(node.tag, str):
                continue
            local = etree.QName(node).localname
            if local not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                continue
            nv = node.find(".//" + qn("p:cNvPr"))
            rec = {"tag": local, "depth": depth}
            if nv is not None:
                rec["id"] = int(nv.get("id"))
                rec["name"] = nv.get("name")
            rec.update(_xfrm_info(node))
            text = "".join(node.itertext()).strip()
            if text and local != "grpSp":
                rec["text"] = text[:24]
            if local == "grpSp":
                # 组内子形状直接挂在 grpSp 下（没有额外的 spTree 包裹）
                rec["children"] = walk(node, depth + 1)
            out.append(rec)
        return out

    return walk(tree)