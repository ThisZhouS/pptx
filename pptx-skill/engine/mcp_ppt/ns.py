"""OOXML 命名空间与元素工具。"""
from __future__ import annotations

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
}

NSDECL = " ".join(f'xmlns:{k}="{v}"' for k, v in NS.items())


def q(tag: str) -> str:
    """'p:transition' -> '{http://...}transition'"""
    prefix, local = tag.split(":")
    return "{%s}%s" % (NS[prefix], local)


def elem(tag: str, nsmap: dict | None = None, **attrs) -> etree._Element:
    e = etree.Element(q(tag), nsmap=nsmap if nsmap else None)
    for k, v in attrs.items():
        set_attr(e, k, v)
    return e


def sub(parent: etree._Element, tag: str, nsmap: dict | None = None, **attrs) -> etree._Element:
    e = elem(tag, nsmap=nsmap, **attrs)
    parent.append(e)
    return e


def set_attr(e: etree._Element, name: str, value) -> None:
    if value is None:
        return
    key = q(name) if ":" in name else name
    e.set(key, str(value))


def parse_fragment(xml: str) -> etree._Element:
    """把不带命名空间声明的片段解析成元素（前缀统一由 NSDECL 提供）。"""
    wrapper = f"<frag {NSDECL}>{xml}</frag>"
    root = etree.fromstring(wrapper.encode("utf-8"))
    node = root[0]
    root.remove(node)
    return node


def to_xml(e: etree._Element, pretty: bool = True) -> str:
    return etree.tostring(e, pretty_print=pretty, encoding="unicode")