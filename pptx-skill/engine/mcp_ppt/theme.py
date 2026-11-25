"""主题（theme1.xml）与放映设置（presProps.xml/showPr）。

真机结构（rev/out/theme_map.json、gt_theme.pptx）

theme1.xml
  <a:theme><a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="1F497D"/></a:dk2> … <a:accent1>…<a:accent6>
      <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink>…
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>
        <a:font script="Hans" typeface="宋体"/> …</a:majorFont>
      <a:minorFont>…</a:minorFont>
    </a:fontScheme>
    <a:fmtScheme>…</a:fmtScheme>
  </a:themeElements>

presProps.xml（放映设置不在 presentation.xml 里）
  <p:presentationPr><p:showPr loop="1"><p:present/><p:sldAll/><p:penClr>…</p:penClr></p:showPr></p:presentationPr>
  改放映设置后 PowerPoint 只在 showPr 上写 loop="1"，其余默认值省略；展示方式用子元素 present/browse/kiosk + sldAll。
"""
from __future__ import annotations

import re
import zipfile

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

PRES_PROPS_CT = "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"

DEFAULT_COLORS = {
    "dk1": "000000", "lt1": "FFFFFF", "dk2": "44546A", "lt2": "E7E6E6",
    "accent1": "4472C4", "accent2": "ED7D31", "accent3": "A5A5A5",
    "accent4": "FFC000", "accent5": "5B9BD5", "accent6": "70AD47",
    "hlink": "0563C1", "folHlink": "954F72",
}
COLOR_KEYS = list(DEFAULT_COLORS)


def _q(ns: str, tag: str) -> str:
    return "{%s}%s" % (ns, tag)


def _theme_parts(prs):
    return [p for p in prs.part.package.iter_parts()
            if str(p.partname).startswith("/ppt/theme/")]


def apply_theme(prs, colors: dict | None = None, fonts: dict | None = None,
                name: str | None = None) -> dict:
    """改 theme1.xml 的配色与字体（保留 fmtScheme/其余结构）。

    theme 部件在 python-pptx 里通常是普通 Part（blob），不是 XmlPart，
    所以统一按"解析 blob → 修改 → 写回 _blob"处理；若是 XmlPart 则直接改 element。
    """
    applied = {"colors": {}, "fonts": {}}
    for part in _theme_parts(prs):
        element = getattr(part, "_element", None)
        is_xml_part = element is not None and etree.QName(element).localname == "theme"
        root = element if is_xml_part else etree.fromstring(part.blob)
        if etree.QName(root).localname != "theme":
            continue
        clr = root.find(".//" + _q(A, "clrScheme"))
        if clr is not None:
            if name:
                clr.set("name", str(name))
            for key, value in (colors or {}).items():
                key = str(key).lower()
                if key not in COLOR_KEYS or not value:
                    continue
                node = clr.find(_q(A, key))
                if node is None or not len(node):
                    continue
                child = node[0]
                hexv = str(value).lstrip("#").upper()
                if key in ("dk1", "lt1"):
                    child.set("lastClr", hexv)
                else:
                    child.set("val", hexv)
                applied["colors"][key] = hexv
        fs = root.find(".//" + _q(A, "fontScheme"))
        if fs is not None and fonts:
            for scope, tag in (("major", "majorFont"), ("minor", "minorFont")):
                font_el = fs.find(_q(A, tag))
                if font_el is None:
                    continue
                latin = fonts.get(f"{scope}_latin") or fonts.get(scope)
                ea = fonts.get(f"{scope}_ea") or fonts.get(scope)
                hans = fonts.get("hans") or ea
                if latin:
                    node = font_el.find(_q(A, "latin"))
                    if node is None:
                        node = etree.SubElement(font_el, _q(A, "latin"))
                    node.set("typeface", str(latin))
                    applied["fonts"][f"{scope}_latin"] = str(latin)
                if ea:
                    node = font_el.find(_q(A, "ea"))
                    if node is None:
                        node = etree.SubElement(font_el, _q(A, "ea"))
                    node.set("typeface", str(ea))
                    applied["fonts"][f"{scope}_ea"] = str(ea)
                if hans:
                    for f in font_el.findall(_q(A, "font")):
                        if f.get("script") == "Hans":
                            f.set("typeface", str(hans))
                            applied["fonts"]["hans"] = str(hans)
        if not is_xml_part:
            part._blob = etree.tostring(root, xml_declaration=True,
                                        encoding="UTF-8", standalone=True)
    return applied


def extract_theme(path: str) -> dict:
    """从已有 pptx 提取配色与字体（用于贴合企业模板）。"""
    out = {"colors": {}, "fonts": {}, "source": path}
    with zipfile.ZipFile(path) as zf:
        themes = [n for n in zf.namelist() if re.match(r"ppt/theme/theme\d+\.xml$", n)]
        if not themes:
            return out
        xml = zf.read(sorted(themes)[0])
    root = etree.fromstring(xml)
    clr = root.find(".//" + _q(A, "clrScheme"))
    if clr is not None:
        out["theme_name"] = clr.get("name")
        for key in COLOR_KEYS:
            node = clr.find(_q(A, key))
            if node is None or not len(node):
                continue
            child = node[0]
            val = child.get("val")
            if child.tag.endswith("sysClr"):
                val = child.get("lastClr") or val
            if val:
                out["colors"][key] = val.upper()
    fs = root.find(".//" + _q(A, "fontScheme"))
    if fs is not None:
        for scope, tag in (("major", "majorFont"), ("minor", "minorFont")):
            font_el = fs.find(_q(A, tag))
            if font_el is None:
                continue
            latin = font_el.find(_q(A, "latin"))
            ea = font_el.find(_q(A, "ea"))
            if latin is not None and latin.get("typeface"):
                out["fonts"][f"{scope}_latin"] = latin.get("typeface")
            if ea is not None and ea.get("typeface"):
                out["fonts"][f"{scope}_ea"] = ea.get("typeface")
            for f in font_el.findall(_q(A, "font")):
                if f.get("script") == "Hans":
                    out["fonts"]["hans"] = f.get("typeface")
    return out


# ------------------------------------------------------------------ 放映设置

def _pres_props_part(prs):
    for rel in prs.part.rels.values():
        if rel.reltype == RT.PRES_PROPS:
            return rel.target_part
    return None


def _build_pres_props(loop: bool, show_type: str, narration: bool,
                      animation: bool, timings: bool, scrollbar: bool) -> bytes:
    ns = f'xmlns:a="{A}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="{P}"'
    attrs = []
    if loop:
        attrs.append('loop="1"')
    if not narration:
        attrs.append('showNarration="0"')
    if not animation:
        attrs.append('showAnimation="0"')
    if not timings:
        attrs.append('useTimings="0"')
    if scrollbar:
        attrs.append('showScrollbar="1"')
    kind = {"kiosk": "<p:kiosk/>", "browse": "<p:browse/>"}.get(str(show_type).lower(), "<p:present/>")
    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f"<p:presentationPr {ns}><p:showPr{' ' if attrs else ''}{' '.join(attrs)}>"
        f"{kind}<p:sldAll/><p:penClr><a:prstClr val=\"red\"/></p:penClr></p:showPr></p:presentationPr>"
    )
    return xml.encode("utf-8")


def set_show(prs, spec: dict | None) -> dict:
    """写放映设置。spec: {loop, show_type: speaker|kiosk|browse, narration, animation, timings, scrollbar}"""
    if not spec:
        return {"show": None}
    loop = bool(spec.get("loop"))
    show_type = str(spec.get("show_type") or spec.get("type") or "speaker")
    narration = bool(spec.get("narration"))
    animation = bool(spec.get("animation", True))
    timings = bool(spec.get("timings", True))
    scrollbar = bool(spec.get("scrollbar"))
    blob = _build_pres_props(loop, show_type, narration, animation, timings, scrollbar)
    part = _pres_props_part(prs)
    if part is None:
        part = Part(PackURI("/ppt/presProps.xml"), PRES_PROPS_CT, blob, package=prs.part.package)
        prs.part.relate_to(part, RT.PRES_PROPS)
    else:
        # XmlPart 没有 setter，整体替换底层 blob
        try:
            part._blob = blob  # type: ignore[attr-defined]
        except Exception:  # noqa: BLE001
            part._element = etree.fromstring(blob)  # type: ignore[attr-defined]
    return {"show": {"loop": loop, "show_type": show_type, "narration": narration,
                     "animation": animation, "timings": timings,
                     "created_part": part.partname}}


def read_show(path: str) -> dict:
    try:
        with zipfile.ZipFile(path) as zf:
            if "ppt/presProps.xml" not in zf.namelist():
                return {"show": None}
            root = etree.fromstring(zf.read("ppt/presProps.xml"))
    except Exception as exc:  # noqa: BLE001
        return {"error": repr(exc)}
    show = root.find(_q(P, "showPr"))
    if show is None:
        return {"show": None}
    kids = [etree.QName(c).localname for c in show if isinstance(c.tag, str)]
    return {"show": {**{k: v for k, v in show.attrib.items()}, "modes": kids}}