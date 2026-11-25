"""DrawingML 属性写入器：填充 / 线条 / 效果 / 富文本 / 图片裁剪。

写法全部来自 PowerPoint 16.0 真机采样（rev/acquire_styles.py，证据见 rev/out/styles_raw）：

填充
  solid      <a:solidFill><a:srgbClr val="RRGGBB"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
  gradient   <a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0">…</a:gs><a:gs pos="100000">…</a:gs></a:gsLst>
             <a:lin ang="2700000" scaled="0"/> 或 <a:path path="rect"><a:fillToRect …/></a:path></a:gradFill>
  none       <a:noFill/>
线条
  <a:ln w="38100" cap="rnd"><a:solidFill>…</a:solidFill><a:prstDash val="dash"/>
    <a:headEnd type="triangle"/><a:tailEnd type="arrow"/></a:ln>
效果（effectLst 子元素顺序：blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge）
  <a:outerShdw blurRad="152400" dist="107763" dir="2699994" rotWithShape="0"><a:srgbClr val="000000">
    <a:alpha val="65000"/></a:srgbClr></a:outerShdw>
  <a:glow rad="152400">…</a:glow>   <a:softEdge rad="127000"/>
文本
  a:pPr@algn,@lvl,@marL,@indent + lnSpc/spcBef/spcAft + buSzPct/buFont/buChar|buNone
  a:rPr@sz(1/100pt),@b,@i,@u,@strike,@baseline + solidFill + latin/ea
图片
  <a:blipFill><a:srcRect l="10000" t="0" r="0" b="0"/><a:blip><a:lum bright="20000" contrast="25000"/>
单位：pct 值 = 百分比 ×1000；spcPts = 1/100 pt；EMU = 914400/英寸；裁剪 srcRect = 比例 ×100000
"""
from __future__ import annotations

from lxml import etree

from .ns import NS, elem, q, set_attr, sub

A = NS["a"]

SPPR_ORDER = ["a:xfrm", "a:custGeom", "a:prstGeom", "a:noFill", "a:solidFill", "a:gradFill",
              "a:blipFill", "a:pattFill", "a:grpFill", "a:ln", "a:effectLst", "a:effectDag",
              "a:scene3d", "a:sp3d", "a:extLst"]
FILL_TAGS = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill", "a:grpFill")
EFFECT_ORDER = ["a:blur", "a:fillOverlay", "a:glow", "a:innerShdw", "a:outerShdw",
                "a:prstShdw", "a:reflection", "a:softEdge"]
LN_ORDER = ["a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:prstDash",
            "a:round", "a:bevel", "a:miter", "a:headEnd", "a:tailEnd", "a:extLst"]
PPR_ORDER = ["a:lnSpc", "a:spcBef", "a:spcAft", "a:buClrTx", "a:buClr", "a:buSzTx",
             "a:buSzPct", "a:buSzPts", "a:buFontTx", "a:buFont", "a:buNone", "a:buAutoNum",
             "a:buChar", "a:tabLst", "a:defRPr", "a:extLst"]
RPR_ORDER = ["a:ln", "a:noFill", "a:solidFill", "a:gradFill", "a:pattFill", "a:effectLst",
             "a:highlight", "a:uLnTx", "a:uLn", "a:uFillTx", "a:uFill", "a:latin", "a:ea",
             "a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"]


# ------------------------------------------------------------------ 通用工具

def _local(el) -> str:
    return etree.QName(el).localname if isinstance(el.tag, str) else ""


def _tag(el) -> str:
    return f"a:{_local(el)}"


def insert_ordered(parent, child, order: list[str]) -> None:
    """按 schema 顺序插入子元素。"""
    name = _tag(child)
    if name not in order:
        parent.append(child)
        return
    idx = order.index(name)
    for existing in list(parent):
        if not isinstance(existing.tag, str):
            continue
        other = _tag(existing)
        if other in order and order.index(other) > idx:
            existing.addprevious(child)
            return
    parent.append(child)


def _drop(parent, tags: list[str]) -> None:
    for t in tags:
        for el in parent.findall(q(t)):
            parent.remove(el)


def color_hex(value: str) -> str:
    return str(value).lstrip("#").upper()


def _clr(color: str, alpha: float | None = None):
    """<a:srgbClr val="RRGGBB"><a:alpha val=".."/></a:srgbClr>"""
    clr = elem("a:srgbClr")
    clr.set("val", color_hex(color))
    if alpha is not None:
        a = elem("a:alpha")
        a.set("val", str(int(round(max(0.0, min(1.0, float(alpha))) * 100000))))
        clr.append(a)
    return clr


def _color_el(tag: str, color: str, alpha: float | None = None):
    """带颜色的容器元素：solidFill / buClr / ln 的填充等。"""
    el = elem(tag)
    el.append(_clr(color, alpha))
    return el


# ------------------------------------------------------------------ 填充

def set_fill(sp_pr, spec) -> dict:
    """spec: "#RRGGBB" | "none" | {type, color, alpha} | {type:gradient, stops, angle, radial}."""
    _drop(sp_pr, list(FILL_TAGS))
    if spec is None:
        return {"fill": None}
    if isinstance(spec, str):
        if spec.strip().lower() in ("none", "transparent", "无"):
            insert_ordered(sp_pr, elem("a:noFill"), SPPR_ORDER)
            return {"fill": "none"}
        node = _color_el("a:solidFill", spec)
        insert_ordered(sp_pr, node, SPPR_ORDER)
        return {"fill": color_hex(spec)}

    kind = str(spec.get("type") or ("gradient" if spec.get("stops") or spec.get("from") else "solid")).lower()
    if kind in ("none", "no", "无"):
        insert_ordered(sp_pr, elem("a:noFill"), SPPR_ORDER)
        return {"fill": "none"}
    if kind in ("gradient", "grad", "渐变"):
        stops = spec.get("stops")
        if not stops:
            stops = [{"pos": 0, "color": spec.get("from", "#FFFFFF")},
                     {"pos": 100, "color": spec.get("to", "#000000")}]
        grad = elem("a:gradFill")
        grad.set("flip", "none")
        grad.set("rotWithShape", "1")
        gs_lst = sub(grad, "a:gsLst")
        for stop in sorted(stops, key=lambda s: float(s.get("pos", 0))):
            gs = sub(gs_lst, "a:gs")
            gs.set("pos", str(int(round(float(stop.get("pos", 0)) * 1000))))
            gs.append(_clr(stop.get("color", "#FFFFFF"), stop.get("alpha")))
        if str(spec.get("style", "linear")).lower() in ("radial", "path", "径向"):
            path = sub(grad, "a:path")
            path.set("path", str(spec.get("path", "rect")))
            rect = sub(path, "a:fillToRect")
            for k in ("l", "t", "r", "b"):
                rect.set(k, str(int(round(float(spec.get(k, 50)) * 1000))))
        else:
            lin = sub(grad, "a:lin")
            lin.set("ang", str(int(round(float(spec.get("angle", 0)) * 60000)) % 21600000))
            lin.set("scaled", "0")
        insert_ordered(sp_pr, grad, SPPR_ORDER)
        return {"fill": "gradient", "stops": len(stops), "angle": spec.get("angle")}

    node = _color_el("a:solidFill", spec.get("color", "#FFFFFF"), spec.get("alpha"))
    insert_ordered(sp_pr, node, SPPR_ORDER)
    return {"fill": color_hex(spec.get("color", "#FFFFFF")), "alpha": spec.get("alpha")}


# ------------------------------------------------------------------ 线条

DASHES = {"solid", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash",
          "sysDashDot", "sysDashDotDot", "sysDot", "dot", "round", "square", "none"}


def set_line(sp_pr, spec: dict) -> dict:
    """spec: {color, width(pt), dash, alpha, head, tail, cap, join}. color=none 表示无线条。"""
    if spec is None:
        return {"line": None}
    ln = sp_pr.find(q("a:ln"))
    if ln is None:
        ln = elem("a:ln")
        insert_ordered(sp_pr, ln, SPPR_ORDER)
    _drop(ln, LN_ORDER)
    color = spec.get("color", spec.get("line_color"))
    if str(spec.get("visible", True)).lower() in ("false", "0", "none") or str(color).lower() in ("none", "无"):
        insert_ordered(ln, elem("a:noFill"), LN_ORDER)
        return {"line": "none"}
    if spec.get("width") is not None:
        ln.set("w", str(int(round(float(spec["width"]) * 12700))))
    if spec.get("cap"):
        ln.set("cap", {"round": "rnd", "rnd": "rnd", "square": "sq", "flat": "flat"}.get(
            str(spec["cap"]).lower(), "flat"))
    if color:
        insert_ordered(ln, _color_el("a:solidFill", color, spec.get("alpha")), LN_ORDER)
    if spec.get("dash"):
        dash = str(spec["dash"])
        prst = dash if dash in DASHES else "dash"
        node = elem("a:prstDash")
        node.set("val", prst)
        insert_ordered(ln, node, LN_ORDER)
    join = str(spec.get("join") or "").lower()
    if join in ("round", "bevel", "miter"):
        insert_ordered(ln, elem(f"a:{join}"), LN_ORDER)
    for side, key in (("a:headEnd", "head"), ("a:tailEnd", "tail")):
        if spec.get(key):
            node = elem(side)
            node.set("type", str(spec[key]))
            if spec.get("arrow_size"):
                node.set("w", str(spec["arrow_size"]))
                node.set("len", str(spec["arrow_size"]))
            insert_ordered(ln, node, LN_ORDER)
    return {"line": {"color": color_hex(color) if color else None,
                     "width": spec.get("width"), "dash": spec.get("dash"),
                     "alpha": spec.get("alpha"), "head": spec.get("head"),
                     "tail": spec.get("tail")}}


# ------------------------------------------------------------------ 效果

def set_effects(sp_pr, spec: dict) -> dict:
    """spec: {shadow:{blur,dist,dir,color,alpha,inner}, glow:{radius,color,alpha}, soft_edge:pt}."""
    if spec is None:
        return {"effects": None}
    if str(spec.get("clear", False)).lower() in ("true", "1"):
        _drop(sp_pr, ["a:effectLst", "a:effectDag"])
        return {"effects": "cleared"}
    effect_lst = sp_pr.find(q("a:effectLst"))
    if effect_lst is None:
        effect_lst = elem("a:effectLst")
        insert_ordered(sp_pr, effect_lst, SPPR_ORDER)
    applied = {}
    shadow = spec.get("shadow")
    if shadow is not None:
        if str(shadow).lower() in ("none", "false", "0"):
            _drop(effect_lst, ["a:outerShdw", "a:innerShdw"])
            applied["shadow"] = "none"
        else:
            sh = shadow if isinstance(shadow, dict) else {}
            inner = bool(sh.get("inner"))
            _drop(effect_lst, ["a:innerShdw"] if inner else ["a:outerShdw"])
            node = elem("a:innerShdw" if inner else "a:outerShdw")
            if sh.get("blur") is not None:
                node.set("blurRad", str(int(round(float(sh["blur"]) * 12700))))
            if sh.get("dist") is not None:
                node.set("dist", str(int(round(float(sh["dist"]) * 12700))))
            node.set("dir", str(int(round(float(sh.get("dir", 45)) * 60000)) % 21600000))
            if not inner:
                node.set("rotWithShape", "0")
            node.append(_clr(sh.get("color", "#000000"), sh.get("alpha", 0.4)))
            insert_ordered(effect_lst, node, EFFECT_ORDER)
            applied["shadow"] = {"blur": sh.get("blur"), "dist": sh.get("dist"),
                                 "dir": sh.get("dir", 45), "inner": inner}
    glow = spec.get("glow")
    if glow is not None:
        if str(glow).lower() in ("none", "false", "0"):
            _drop(effect_lst, ["a:glow"])
            applied["glow"] = "none"
        else:
            g = glow if isinstance(glow, dict) else {}
            _drop(effect_lst, ["a:glow"])
            node = elem("a:glow")
            node.set("rad", str(int(round(float(g.get("radius", 8)) * 12700))))
            node.append(_clr(g.get("color", "#FFFFFF"), g.get("alpha", 0.6)))
            insert_ordered(effect_lst, node, EFFECT_ORDER)
            applied["glow"] = {"radius": g.get("radius", 8)}
    soft = spec.get("soft_edge", spec.get("softEdge"))
    if soft is not None:
        _drop(effect_lst, ["a:softEdge"])
        if str(soft).lower() not in ("none", "false", "0"):
            node = elem("a:softEdge")
            node.set("rad", str(int(round(float(soft) * 12700))))
            insert_ordered(effect_lst, node, EFFECT_ORDER)
            applied["soft_edge"] = float(soft)
    return {"effects": applied or "noop"}


# ------------------------------------------------------------------ 富文本

def _rpr_for(run: dict, defaults: dict) -> etree._Element:
    rpr = elem("a:rPr")
    rpr.set("lang", str(run.get("lang") or defaults.get("lang") or "zh-CN"))
    rpr.set("altLang", "en-US")
    size = run.get("size", defaults.get("size"))
    if size is not None:
        rpr.set("sz", str(int(round(float(size) * 100))))
    bold = run.get("bold", defaults.get("bold"))
    if bold is not None:
        rpr.set("b", "1" if bold else "0")
    italic = run.get("italic", defaults.get("italic"))
    if italic is not None:
        rpr.set("i", "1" if italic else "0")
    underline = run.get("underline", defaults.get("underline"))
    if underline is not None:
        rpr.set("u", "sng" if underline is True else str(underline))
    strike = run.get("strike", defaults.get("strike"))
    if strike is not None:
        rpr.set("strike", "sngStrike" if strike is True else str(strike))
    if run.get("baseline") is not None:
        rpr.set("baseline", str(int(round(float(run["baseline"]) * 1000))))
    if run.get("spc") is not None:
        rpr.set("spc", str(int(round(float(run["spc"]) * 100))))
    color = run.get("color", defaults.get("color"))
    if color:
        insert_ordered(rpr, _color_el("a:solidFill", color, run.get("alpha", defaults.get("alpha"))), RPR_ORDER)
    if run.get("emboss") or defaults.get("emboss"):
        fx = elem("a:effectLst")
        sh = sub(fx, "a:innerShdw")
        sh.set("dist", "107763")
        sh.set("dir", "13500000")
        sub(sh, "a:scrgbClr", r="0", g="0", b="0")
        insert_ordered(rpr, fx, RPR_ORDER)
    font = run.get("font", defaults.get("font"))
    if font:
        latin = elem("a:latin")
        latin.set("typeface", str(font))
        insert_ordered(rpr, latin, RPR_ORDER)
        ea = elem("a:ea")
        ea.set("typeface", str(run.get("font_ea") or defaults.get("font_ea") or font))
        insert_ordered(rpr, ea, RPR_ORDER)
    return rpr


def _ppr_for(para: dict, defaults: dict) -> etree._Element:
    ppr = elem("a:pPr")
    align = para.get("align", defaults.get("align"))
    if align:
        ppr.set("algn", {"left": "l", "center": "ctr", "right": "r", "justify": "just",
                         "l": "l", "ctr": "ctr", "r": "r", "just": "just"}.get(str(align).lower(), "l"))
    level = para.get("level", defaults.get("level"))
    if level:
        ppr.set("lvl", str(int(level)))
    indent = para.get("indent_emu")
    mar_l = para.get("marL_emu")
    if level and mar_l is None:
        mar_l = 914400 * int(level)
        indent = indent if indent is not None else -457200
    if mar_l is not None:
        ppr.set("marL", str(int(mar_l)))
    if indent is not None:
        ppr.set("indent", str(int(indent)))
    line_spacing = para.get("line_spacing", defaults.get("line_spacing"))
    if line_spacing is not None:
        ln_spc = sub(ppr, "a:lnSpc")
        pct = sub(ln_spc, "a:spcPct")
        pct.set("val", str(int(round(float(line_spacing) * 100000))))
    before = para.get("space_before", defaults.get("space_before"))
    if before is not None:
        node = sub(ppr, "a:spcBef")
        sub(node, "a:spcPts").set("val", str(int(round(float(before) * 100))))
    after = para.get("space_after", defaults.get("space_after"))
    if after is not None:
        node = sub(ppr, "a:spcAft")
        sub(node, "a:spcPts").set("val", str(int(round(float(after) * 100))))
    bullet = para.get("bullet", defaults.get("bullet"))
    if bullet is not None:
        if str(bullet).lower() in ("none", "false", "0", "无"):
            insert_ordered(ppr, elem("a:buNone"), PPR_ORDER)
        else:
            b = bullet if isinstance(bullet, dict) else {"char": bullet if isinstance(bullet, str) else "•"}
            if b.get("color"):
                insert_ordered(ppr, _color_el("a:buClr", b["color"]), PPR_ORDER)
            if b.get("size_pct") is not None:
                sz = elem("a:buSzPct")
                sz.set("val", str(int(round(float(b["size_pct"]) * 100000))))
                insert_ordered(ppr, sz, PPR_ORDER)
            font = elem("a:buFont")
            font.set("typeface", str(b.get("font") or "Arial"))
            insert_ordered(ppr, font, PPR_ORDER)
            if b.get("auto_num"):
                num = elem("a:buAutoNum")
                num.set("type", str(b.get("auto_num") if b.get("auto_num") is not True else "arabicPeriod"))
                insert_ordered(ppr, num, PPR_ORDER)
            else:
                ch = elem("a:buChar")
                ch.set("char", str(b.get("char") or "•"))
                insert_ordered(ppr, ch, PPR_ORDER)
    return ppr


def set_text_body(tx_body, paragraphs: list, defaults: dict | None = None) -> dict:
    """把 a:txBody 的段落整体重建。

    paragraphs: [{runs:[{text, bold, size, color, ...}], align, level, line_spacing,
                  space_before, space_after, bullet}] 或 纯字符串列表。
    """
    defaults = defaults or {}
    if not paragraphs:
        return {"paragraphs": 0}
    body_pr = tx_body.find(q("a:bodyPr"))
    lst_style = tx_body.find(q("a:lstStyle"))
    for p in tx_body.findall(q("a:p")):
        tx_body.remove(p)
    count_runs = 0
    for para in paragraphs:
        if isinstance(para, str):
            para = {"runs": [{"text": para}]}
        p_el = elem("a:p")
        p_el.append(_ppr_for(para, defaults))
        runs = para.get("runs") or []
        if not runs and para.get("text"):
            runs = [{"text": para["text"]}]
        for run in runs:
            r = sub(p_el, "a:r")
            r.append(_rpr_for(run, defaults))
            t = sub(r, "a:t")
            t.text = str(run.get("text", ""))
            count_runs += 1
        p_el.append(elem("a:endParaRPr"))
        tx_body.append(p_el)
    # 保持 bodyPr/lstStyle 在最前
    for node in (lst_style, body_pr):
        if node is not None and node.getparent() is tx_body:
            tx_body.remove(node)
            tx_body.insert(0, node)
    return {"paragraphs": len(paragraphs), "runs": count_runs}


# ------------------------------------------------------------------ 图片

def set_crop(pic_el, crop: dict) -> dict:
    """crop 用比例（0..1），写 a:srcRect（比例 ×100000）。"""
    blip_fill = pic_el.find(q("p:blipFill"))
    if blip_fill is None:
        blip_fill = pic_el.find(".//" + q("p:blipFill"))
    if blip_fill is None:
        return {"crop": "no blipFill"}
    src = blip_fill.find(q("a:srcRect"))
    if src is None:
        src = elem("a:srcRect")
        blip = blip_fill.find(q("a:blip"))
        if blip is not None:
            blip.addnext(src)
        else:
            blip_fill.insert(0, src)
    applied = {}
    for key, attr in (("l", "l"), ("t", "t"), ("r", "r"), ("b", "b"),
                      ("left", "l"), ("top", "t"), ("right", "r"), ("bottom", "b")):
        if key in crop:
            val = int(round(float(crop[key]) * 100000))
            src.set(attr, str(val))
            applied[attr] = round(float(crop[key]), 4)
    return {"crop": applied}


def set_lum(pic_el, brightness: float | None = None, contrast: float | None = None) -> dict:
    """brightness/contrast 用 -1..1，写 a:blip/a:lum（×100000）。"""
    blip = pic_el.find(".//" + q("a:blip"))
    if blip is None:
        return {"lum": "no blip"}
    lum = blip.find(q("a:lum"))
    if lum is None:
        lum = elem("a:lum")
        blip.append(lum)
    applied = {}
    if brightness is not None:
        lum.set("bright", str(int(round(float(brightness) * 100000))))
        applied["bright"] = brightness
    if contrast is not None:
        lum.set("contrast", str(int(round(float(contrast) * 100000))))
        applied["contrast"] = contrast
    return {"lum": applied}
