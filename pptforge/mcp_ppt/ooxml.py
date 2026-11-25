"""幻灯片层 OOXML 写入器：切换 (<p:transition>) 与动画 (<p:timing>)。

所有写法都来自 PowerPoint 16.0 的 ground-truth 采样（见 pptforge/rev 与 docs/REVERSE_REPORT.md）：

切换
----
* 位置：<p:sld> / <p:clrMapOvr> 之后，<p:timing> 之前
* p 命名空间效果 → 裸 <p:transition>；p14/p15/p159 效果 → mc:AlternateContent + mc:Fallback
* 精确时长 p14:dur（毫秒）必须走 mc:AlternateContent Requires="p14"

动画
----
* 骨架：timing/tnLst/par/cTn(tmRoot)/childTnLst/seq/cTn(mainSeq)/childTnLst/
         par(cTn fill=hold, cond delay=indefinite)/childTnLst/
         par(cTn fill=hold, cond delay=0)/childTnLst/
         par(cTn presetID/presetClass/presetSubtype/grpId/nodeType)/childTnLst/行为节点
* nodeType：clickEffect / withEffect / afterEffect
* 组级第二个 cond：evt="onBegin" + <p:tn val="2"/>（首位效果不是 click 时）
* 文本分级：p:spTgt/p:txEl/p:pRg st/end，bldP build="p"
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field

from lxml import etree

from .catalog import Catalog, get_catalog
from .ns import NS, NSDECL, elem, parse_fragment, q, set_attr, sub, to_xml

X_BLANK = "blank"


# ----------------------------------------------------------------- 工具

def _int_ms(seconds: float | int | None) -> int | None:
    if seconds is None:
        return None
    return int(round(float(seconds) * 1000))


def _speed_attr(speed: str | None) -> str | None:
    if not speed:
        return None
    table = {"slow": "slow", "medium": "med", "med": "med", "fast": "fast", 1: "slow", 2: "med", 3: "fast"}
    return table.get(speed, "med")


def _find_existing_transition(sld: etree._Element) -> list[etree._Element]:
    """返回需要被替换的节点（裸 transition 或其外层 AlternateContent）。"""
    out = []
    for child in list(sld):
        if not isinstance(child.tag, str):
            continue
        if child.tag == q("p:transition"):
            out.append(child)
        elif child.tag == q("mc:AlternateContent"):
            for node in child.iter(q("p:transition")):
                out.append(child)
                break
    return out


def _insert_in_sld_order(sld: etree._Element, node: etree._Element) -> None:
    """把 node 插到 clrMapOvr 之后、timing/extLst 之前。"""
    timing = sld.find(q("p:timing"))
    if timing is not None:
        timing.addprevious(node)
        return
    ext = sld.find(q("p:extLst"))
    if ext is not None:
        ext.addprevious(node)
        return
    sld.append(node)


# ----------------------------------------------------------------- 切换

def build_transition(spec: dict, cat: Catalog | None = None) -> dict:
    """构造 <p:transition> 节点（可能带 mc:AlternateContent）。

    返回 {"node": Element|None, "warnings": [...], "applied": {...}}
    """
    cat = cat or get_catalog()
    warnings: list[str] = []
    eff_id = cat.resolve_transition(spec.get("effect") or spec.get("name"))
    eff = cat.transitions[eff_id]

    duration = spec.get("duration")
    dur_ms = _int_ms(duration)
    spd = _speed_attr(spec.get("speed")) or ("med" if duration is None else None)
    adv_after = spec.get("advance_after") or spec.get("advTm")
    adv_ms = _int_ms(adv_after)
    adv_click = spec.get("advance_on_click")

    base_attrs: dict[str, str] = {}
    if spd:
        base_attrs["spd"] = spd
    if adv_click is False:
        base_attrs["advClick"] = "0"
    if adv_ms:
        base_attrs["advTm"] = str(adv_ms)

    def make_core(with_dur: bool, forced_child: str | None = None) -> etree._Element:
        nsmap = {"p14": NS["p14"]} if (with_dur and dur_ms) else None
        t = elem("p:transition", nsmap=nsmap)
        for k, v in base_attrs.items():
            set_attr(t, k, v)
        if with_dur and dur_ms:
            set_attr(t, "p14:dur", dur_ms)
        child_tag = forced_child or (eff["element"] if eff.get("element") else None)
        if child_tag:
            child = elem(child_tag)
            if forced_child is None:
                for k, v in (eff.get("element_attrs") or {}).items():
                    set_attr(child, k, v)
            t.append(child)
        return t

    if eff_id == "none" and not (adv_ms or adv_click is False):
        return {"node": None, "warnings": warnings, "applied": {"effect": "none"}}

    requires = eff.get("requires")
    needs_alt = bool(requires) or bool(dur_ms and eff["prefix"] == "p")

    if not needs_alt:
        node = make_core(with_dur=False)
        return {"node": node, "warnings": warnings,
                "applied": {"effect": eff_id, "zh": eff.get("zh"), "spd": spd,
                            "duration_ms": dur_ms, "advance_ms": adv_ms,
                            "alternate_content": False}}

    req = requires or "p14"
    ac_nsmap = {"mc": NS["mc"]}
    if req in NS:
        ac_nsmap[req] = NS[req]
    ac = elem("mc:AlternateContent", nsmap=ac_nsmap)
    choice = elem("mc:Choice")
    choice.set("Requires", req)
    choice.append(make_core(with_dur=True))
    fallback = elem("mc:Fallback")
    if eff["prefix"] == "p":
        fallback.append(make_core(with_dur=False))
    else:
        fb = elem("p:transition")
        for k, v in base_attrs.items():
            set_attr(fb, k, v)
        fb.append(elem("p:fade"))
        fallback.append(fb)
    ac.append(choice)
    ac.append(fallback)
    return {"node": ac, "warnings": warnings,
            "applied": {"effect": eff_id, "zh": eff.get("zh"), "requires": req,
                        "spd": spd, "duration_ms": dur_ms, "advance_ms": adv_ms,
                        "advance_on_click": adv_click,
                        "alternate_content": True}}


def set_slide_transition(slde: etree._Element, spec: dict | None,
                         cat: Catalog | None = None) -> dict:
    """写入/清除某页切换。spec=None 表示清除。"""
    for old in _find_existing_transition(slde):
        old.getparent().remove(old)
    if spec is None:
        return {"effect": "none", "cleared": True}
    built = build_transition(spec, cat)
    node = built["node"]
    if node is not None:
        _insert_in_sld_order(slde, node)
    return built["applied"] | {"warnings": built["warnings"]}


def read_transition(slde: etree._Element) -> dict | None:
    """读回一页的切换信息（兼容裸 transition 与 AlternateContent）。"""
    holder = None
    tr = slde.find(q("p:transition"))
    if tr is not None:
        holder = "p:sld"
    else:
        for child in slde:
            if isinstance(child.tag, str) and child.tag == q("mc:AlternateContent"):
                ch = child.find(q("mc:Choice"))
                if ch is not None:
                    tr = ch.find(q("p:transition"))
                    if tr is not None:
                        holder = "mc:AlternateContent/mc:Choice"
                        break
    if tr is None:
        return None
    kids = [etree.QName(c) for c in tr if isinstance(c.tag, str)]
    return {
        "holder": holder,
        "attrs": {etree.QName(k).localname if "}" in k else k: v for k, v in tr.attrib.items()},
        "effects": [f"{k.namespace and NS_ABBR(k.namespace)}:{k.localname}" for k in kids],
        "xml": to_xml(tr),
    }


def NS_ABBR(uri: str) -> str:
    for k, v in NS.items():
        if v == uri:
            return k
    return "?"


# ----------------------------------------------------------------- 动画

@dataclass
class AnimItem:
    shape_id: int
    effect: str = "fade"
    trigger: str = "click"
    duration: float | None = 0.5
    delay: float | None = 0.0
    repeat: float | None = None
    smooth_start: float | None = None   # 秒
    smooth_end: float | None = None
    exit: bool = False
    paragraph: int | None = None        # 文本分级：段落下标
    paragraph_end: int | None = None
    raw: dict = field(default_factory=dict)


class _IdGen:
    def __init__(self) -> None:
        self.n = 0

    def next(self) -> int:
        self.n += 1
        return self.n


def _node_type(trigger: str) -> str:
    return {"click": "clickEffect", "with_previous": "withEffect",
            "after_previous": "afterEffect"}[trigger]


def _behavior_node(template: str, spid: int, dur_ms: int, ids: _IdGen,
                   paragraph: tuple[int, int] | None) -> etree._Element:
    xml = template.replace("{{SPID}}", str(spid))
    node = parse_fragment(xml)
    for ctn in node.iter(q("p:cTn")):
        ctn.set("id", str(ids.next()))
        if ctn.get("dur") not in (None, "1"):
            ctn.set("dur", str(dur_ms))
    if paragraph is not None:
        st, end = paragraph
        for tgt in node.iter(q("p:spTgt")):
            if tgt.get("spid") == str(spid) and tgt.find(q("p:txEl")) is None:
                txel = elem("p:txEl")
                prg = elem("p:pRg")
                prg.set("st", str(st))
                prg.set("end", str(end))
                txel.append(prg)
                tgt.append(txel)
    return node


def build_timing(items: list[AnimItem], cat: Catalog | None = None) -> tuple[etree._Element, dict]:
    """按 ground-truth 结构生成整页 <p:timing>。"""
    cat = cat or get_catalog()
    ids = _IdGen()
    timing = elem("p:timing")
    tn_lst = sub(timing, "p:tnLst")
    par0 = sub(tn_lst, "p:par")
    root = sub(par0, "p:cTn", id=ids.next(), dur="indefinite", restart="never", nodeType="tmRoot")
    root_child = sub(root, "p:childTnLst")
    seq = sub(root_child, "p:seq", concurrent="1", nextAc="seek")
    main = sub(seq, "p:cTn", id=ids.next(), dur="indefinite", nodeType="mainSeq")
    main_child = sub(main, "p:childTnLst")

    # ---- 分组：click 开新组，with/after 顺延到当前组
    groups: list[list[AnimItem]] = []
    for item in items:
        if item.trigger == "click" or not groups:
            groups.append([item])
        else:
            groups[-1].append(item)

    bld: dict[int, str | None] = {}
    for group in groups:
        g_par = sub(main_child, "p:par")
        g_ctn = sub(g_par, "p:cTn", id=ids.next(), fill="hold")
        stc = sub(g_ctn, "p:stCondLst")
        sub(stc, "p:cond", delay="indefinite")
        if group[0].trigger != "click":
            c2 = sub(stc, "p:cond", evt="onBegin", delay="0")
            sub(c2, "p:tn", val=2)
        g_child = sub(g_ctn, "p:childTnLst")
        e_par = sub(g_child, "p:par")
        e_ctn = sub(e_par, "p:cTn", id=ids.next(), fill="hold")
        esc = sub(e_ctn, "p:stCondLst")
        sub(esc, "p:cond", delay="0")
        e_child = sub(e_ctn, "p:childTnLst")

        for item in group:
            spec = cat.animation_spec(item.effect, exit_=item.exit)
            preset = spec["preset"]
            eff_par = sub(e_child, "p:par")
            attrs = {
                "id": ids.next(),
                "presetID": preset["presetID"],
                "presetClass": preset["presetClass"],
                "presetSubtype": preset["presetSubtype"],
                "fill": "hold",
                "grpId": "0",
                "nodeType": _node_type(item.trigger),
            }
            if item.repeat and item.repeat != 1:
                attrs["repeatCount"] = int(round(item.repeat * 1000))
            if item.smooth_start:
                attrs["accel"] = int(round(item.smooth_start * 100000))
            if item.smooth_end:
                attrs["decel"] = int(round(item.smooth_end * 100000))
            eff_ctn = sub(eff_par, "p:cTn", **attrs)
            esc2 = sub(eff_ctn, "p:stCondLst")
            sub(esc2, "p:cond", delay=_int_ms(item.delay) or 0)
            ech = sub(eff_ctn, "p:childTnLst")
            dur_ms = _int_ms(item.duration) or 500
            par_range = None
            if item.paragraph is not None:
                st = item.paragraph
                end = item.paragraph_end if item.paragraph_end is not None else item.paragraph
                par_range = (st, end)
            for tmpl in spec["behaviors"]:
                ech.append(_behavior_node(tmpl, item.shape_id, dur_ms, ids, par_range))
            if item.shape_id not in bld:
                bld[item.shape_id] = "p" if item.paragraph is not None else None

    # ---- seq 的播放条件
    prev = sub(seq, "p:prevCondLst")
    c = sub(prev, "p:cond", evt="onPrev", delay="0")
    tgt = sub(c, "p:tgtEl")
    sub(tgt, "p:sldTgt")
    nxt = sub(seq, "p:nextCondLst")
    c = sub(nxt, "p:cond", evt="onNext", delay="0")
    tgt = sub(c, "p:tgtEl")
    sub(tgt, "p:sldTgt")

    # ---- build 列表
    bld_lst = sub(timing, "p:bldLst")
    for spid, mode in bld.items():
        node = sub(bld_lst, "p:bldP", spid=spid, grpId="0")
        if mode:
            node.set("build", mode)
        else:
            node.set("animBg", "1")

    meta = {
        "items": len(items),
        "groups": len(groups),
        "built_shapes": list(bld.keys()),
        "ids_used": ids.n,
    }
    return timing, meta


def set_slide_timing(slde: etree._Element, items: list[AnimItem],
                     cat: Catalog | None = None) -> dict:
    for old in slde.findall(q("p:timing")):
        slde.remove(old)
    if not items:
        return {"items": 0, "groups": 0, "built_shapes": []}
    timing, meta = build_timing(items, cat)
    _insert_in_sld_order(slde, timing)
    return meta


def read_timing(slde: etree._Element) -> list[dict]:
    """读回一页动画序列（形状、presetClass、nodeType、触发）。"""
    timing = slde.find(q("p:timing"))
    if timing is None:
        return []
    out = []
    for ctn in timing.iter(q("p:cTn")):
        pid = ctn.get("presetID")
        if pid is None:
            continue
        spids = []
        for tgt in ctn.iter(q("p:spTgt")):
            spids.append(int(tgt.get("spid")))
        cond = ctn.find(q("p:stCondLst"))
        delay = None
        if cond is not None:
            c = cond.find(q("p:cond"))
            if c is not None:
                delay = c.get("delay")
        out.append({
            "presetID": pid,
            "presetClass": ctn.get("presetClass"),
            "presetSubtype": ctn.get("presetSubtype"),
            "nodeType": ctn.get("nodeType"),
            "repeatCount": ctn.get("repeatCount"),
            "accel": ctn.get("accel"),
            "decel": ctn.get("decel"),
            "delay": delay,
            "shape_ids": sorted(set(spids)),
        })
    return out