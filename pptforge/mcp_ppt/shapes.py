"""形状层：定位、移动、缩放、组合/解组、属性编辑、层级与对齐。

逆向依据（rev/acquire_shapes.py 真机采样，见 rev/out/shapes_map.json）
-------------------------------------------------------------------
* 几何：`a:xfrm` 里 `a:off(x,y)` + `a:ext(cx,cy)`，单位 EMU（914400 = 1 英寸）
  旋转 `rot`（度×60000），翻转 `flipH/flipV="1"`。
* 组合：`p:grpSp` 内 `a:xfrm` 同时有 `a:off/a:ext`（组在父坐标系的位置尺寸）与
  `a:chOff/a:chExt`（子坐标系）。PowerPoint 组合时 chOff = 子包围盒左上、chExt = 包围盒尺寸、
  off = chOff、ext = chExt（1:1）；子形状的 off/ext 保持不变。
* 平移 = 只改 `off`（`chOff` 不动）；缩放 = 只改 `ext`（`chExt` 不动）——
  子形状视觉位置 = off + (child_off - chOff) × (ext/chExt)。
* 组动画的 target 仍是 `<p:spTgt spid="组id"/>`，组内单形状动画用子形状 id。
"""
from __future__ import annotations

import copy
import math
from dataclasses import dataclass

from lxml import etree
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Cm, Emu, Inches, Pt

from .ns import NS, q
from . import dml

EMU_PER_INCH = 914400
EMU_PER_POINT = 12700


# ------------------------------------------------------------------ 单位

def to_emu(value, basis: int | None = None) -> int:
    """数字按英寸；字符串支持 50% / 1.5in / 120pt / 3cm / 914400emu。"""
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return int(round(value * EMU_PER_INCH))
    text = str(value).strip().lower()
    if text.endswith("%"):
        if basis is None:
            raise ValueError(f"百分比需要基准长度: {value!r}")
        return int(round(basis * float(text[:-1]) / 100.0))
    for suffix, factor in (("emu", 1), ("in", EMU_PER_INCH), ("inch", EMU_PER_INCH),
                           ("pt", EMU_PER_POINT), ("cm", 360000), ("mm", 36000)):
        if text.endswith(suffix):
            return int(round(float(text[: -len(suffix)]) * factor))
    return int(round(float(text) * EMU_PER_INCH))


def emu_to_in(emu: int) -> float:
    return round(emu / EMU_PER_INCH, 4)


def to_emu_len(value, basis: int | None = None) -> int:
    """长度解析：python-pptx Length / 大整数（EMU）/ 英寸数字 / 带单位字符串。

    python-pptx 的 Emu 做算术后会退化成 int（如 Emu - int*Emu），所以用"数值量级"兜底。
    """
    from pptx.util import Length

    if isinstance(value, Length):
        return int(value)
    if isinstance(value, int) and not isinstance(value, bool) and abs(value) >= 10000:
        return int(value)
    return to_emu(value, basis)


ANCHORS = {
    "left-top": (0.0, 0.0), "top-left": (0.0, 0.0), "nw": (0.0, 0.0),
    "center": (0.5, 0.5), "middle": (0.5, 0.5), "c": (0.5, 0.5),
    "left": (0.0, 0.5), "right": (1.0, 0.5), "top": (0.5, 0.0), "bottom": (0.5, 1.0),
    "right-top": (1.0, 0.0), "top-right": (1.0, 0.0), "ne": (1.0, 0.0),
    "left-bottom": (0.0, 1.0), "bottom-left": (0.0, 1.0), "sw": (0.0, 1.0),
    "right-bottom": (1.0, 1.0), "bottom-right": (1.0, 1.0), "se": (1.0, 1.0),
}


def anchor_fraction(anchor: str | None) -> tuple[float, float]:
    return ANCHORS.get(str(anchor or "left-top").strip().lower(), (0.0, 0.0))


# ------------------------------------------------------------------ 元素访问

def _xfrm(shape) -> etree._Element:
    el = shape._element
    xfrm = el.find(".//" + q("a:xfrm"))
    if xfrm is None:
        # 没有 xfrm 的形状（如 graphicFrame 没有；sp/pic 有 spPr/xfrm）
        sp_pr = el.find(q("p:spPr")) if el.tag == q("p:sp") else None
        if sp_pr is None and el.find(q("p:grpSpPr")) is not None:
            sp_pr = el.find(q("p:grpSpPr"))
        if sp_pr is None:
            raise ValueError(f"形状没有可编辑的几何: {getattr(shape, 'name', '?')}")
        xfrm = etree.SubElement(sp_pr, q("a:xfrm"))
    return xfrm


def _set_box(shape, x=None, y=None, w=None, h=None) -> None:
    xfrm = _xfrm(shape)
    off = xfrm.find(q("a:off"))
    if off is None:
        off = etree.SubElement(xfrm, q("a:off"))
        xfrm.remove(off)
        xfrm.insert(0, off)
    ext = xfrm.find(q("a:ext"))
    if ext is None:
        ext = etree.SubElement(xfrm, q("a:ext"))
    if x is not None:
        off.set("x", str(int(x)))
    if y is not None:
        off.set("y", str(int(y)))
    if w is not None:
        ext.set("cx", str(int(w)))
    if h is not None:
        ext.set("cy", str(int(h)))


def iter_shapes(slide_shapes, recursive: bool = True):
    """遍历形状（含组内）。"""
    for shp in slide_shapes:
        yield shp
        if recursive and getattr(shp, "shape_type", None) is not None and shp.__class__.__name__ == "GroupShape":
            yield from iter_shapes(shp.shapes, recursive)


def find_shape(slide, target):
    """按名称或 id 定位形状；支持 group:<name>、shape:<id>。"""
    if target is None:
        raise ValueError("target 不能为空")
    key = str(target).strip()
    prefix = ""
    if ":" in key:
        head, tail = key.split(":", 1)
        if head.lower() in ("group", "shape", "id", "name"):
            prefix, key = head.lower(), tail.strip()
    candidates = []
    for shp in iter_shapes(slide.shapes):
        try:
            if prefix == "group" and shp.__class__.__name__ != "GroupShape":
                continue
            sid = str(shp.shape_id)
        except Exception:  # noqa: BLE001
            continue
        if prefix in ("shape", "id"):
            if sid == key:
                candidates.append(shp)
        elif prefix == "name":
            if shp.name == key:
                candidates.append(shp)
        else:
            if shp.name == key or sid == key:
                candidates.append(shp)
    if not candidates:
        names = [f"{s.name}({s.shape_id})" for s in iter_shapes(slide.shapes)]
        raise KeyError(f"找不到形状 {target!r}；可用: {names}")
    return candidates[0]


def find_shapes(slide, targets) -> list:
    if targets in (None, "*", "all", "全部"):
        return list(iter_shapes(slide.shapes))
    names = targets if isinstance(targets, (list, tuple)) else [targets]
    out, seen = [], set()
    for t in names:
        try:
            shp = find_shape(slide, t)
        except KeyError:
            continue
        if shp.shape_id not in seen:
            seen.add(shp.shape_id)
            out.append(shp)
    return out


def box_of(shape) -> tuple[int, int, int, int]:
    return (int(shape.left or 0), int(shape.top or 0), int(shape.width or 0), int(shape.height or 0))


# ------------------------------------------------------------------ 定位 / 移动 / 缩放

def place(shape, x=None, y=None, w=None, h=None, anchor: str = "left-top",
          slide_w: int = 0, slide_h: int = 0, of=None) -> dict:
    """按锚点放置：x/y 表示"把形状的哪个点放到该坐标"。of 给出基准形状。"""
    ax, ay = anchor_fraction(anchor)
    basis_w = int(getattr(of, "width", 0) or slide_w) or slide_w
    basis_h = int(getattr(of, "height", 0) or slide_h) or slide_h
    cur = box_of(shape)
    width = to_emu(w, slide_w) if w is not None else cur[2]
    height = to_emu(h, slide_h) if h is not None else cur[3]
    ox = int(getattr(of, "left", 0) or 0)
    oy = int(getattr(of, "top", 0) or 0)
    left, top = cur[0], cur[1]
    if x is not None:
        left = ox + to_emu(x, basis_w) - ax * width
    if y is not None:
        top = oy + to_emu(y, basis_h) - ay * height
    _set_box(shape, x=left, y=top, w=width if w is not None else None,
             h=height if h is not None else None)
    return {"shape": shape.name, "left": left, "top": top, "width": width, "height": height}


def move(shape, dx=None, dy=None, to_x=None, to_y=None, anchor: str = "left-top",
         slide_w: int = 0, slide_h: int = 0) -> dict:
    ax, ay = anchor_fraction(anchor)
    left, top, width, height = box_of(shape)
    if to_x is not None:
        left = to_emu(to_x, slide_w) - ax * width
    if to_y is not None:
        top = to_emu(to_y, slide_h) - ay * height
    if dx is not None:
        left += to_emu(dx, slide_w)
    if dy is not None:
        top += to_emu(dy, slide_h)
    _set_box(shape, x=left, y=top)
    return {"shape": shape.name, "left": left, "top": top}


def resize(shape, scale=None, w=None, h=None, keep_aspect: bool = True,
           anchor: str = "left-top", slide_w: int = 0, slide_h: int = 0) -> dict:
    """缩放：scale 可为单数或 [sx, sy]；anchor 为固定不动的点。"""
    left, top, width, height = box_of(shape)
    ax, ay = anchor_fraction(anchor)
    new_w, new_h = width, height
    if scale is not None:
        if isinstance(scale, (list, tuple)):
            sx, sy = float(scale[0]), float(scale[1])
        else:
            sx = sy = float(scale)
        new_w, new_h = int(width * sx), int(height * sy)
    if w is not None:
        new_w = to_emu(w, slide_w)
        if keep_aspect and h is None and width:
            new_h = int(height * new_w / width)
    if h is not None:
        new_h = to_emu(h, slide_h)
        if keep_aspect and w is None and height:
            new_w = int(width * new_h / height)
    left = int(left + (width - new_w) * ax)
    top = int(top + (height - new_h) * ay)
    _set_box(shape, x=left, y=top, w=new_w, h=new_h)
    return {"shape": shape.name, "left": left, "top": top, "width": new_w, "height": new_h}


# ------------------------------------------------------------------ 组合 / 解组

def _max_shape_id(slide) -> int:
    ids = [0]
    for nv in slide.element.iter(q("p:cNvPr")):
        try:
            ids.append(int(nv.get("id")))
        except (TypeError, ValueError):
            pass
    return max(ids)


def group_shapes(slide, shapes: list, name: str | None = None) -> dict:
    """把同层形状组合成 p:grpSp（坐标 1:1，子形状 off/ext 不变）。"""
    sp_tree = slide.shapes._spTree
    elements = []
    for shp in shapes:
        el = shp._element
        if el.getparent() is not sp_tree:
            raise ValueError(f"只能组合同一层的形状；{shp.name} 在其它容器里")
        if el.tag == q("p:grpSp"):
            raise ValueError(f"{shp.name} 已经是组合，请先解组或改选子形状")
        elements.append(el)
    if len(elements) < 2:
        raise ValueError("组合至少需要 2 个形状")
    order = list(sp_tree)
    elements.sort(key=order.index)
    boxes = []
    for el in elements:
        xfrm = el.find(".//" + q("a:xfrm"))
        off = xfrm.find(q("a:off")) if xfrm is not None else None
        ext = xfrm.find(q("a:ext")) if xfrm is not None else None
        if off is None or ext is None:
            continue
        boxes.append((int(off.get("x")), int(off.get("y")),
                      int(ext.get("cx")), int(ext.get("cy"))))
    if not boxes:
        raise ValueError("选中的形状都没有几何信息")
    min_x = min(b[0] for b in boxes)
    min_y = min(b[1] for b in boxes)
    max_x = max(b[0] + b[2] for b in boxes)
    max_y = max(b[1] + b[3] for b in boxes)
    cx, cy = max_x - min_x, max_y - min_y

    grp_id = _max_shape_id(slide) + 1
    grp = etree.Element(q("p:grpSp"))
    nv = etree.SubElement(grp, q("p:nvGrpSpPr"))
    c_nv_pr = etree.SubElement(nv, q("p:cNvPr"))
    c_nv_pr.set("id", str(grp_id))
    c_nv_pr.set("name", name or f"Group {grp_id}")
    etree.SubElement(nv, q("p:cNvGrpSpPr"))
    etree.SubElement(nv, q("p:nvPr"))
    grp_spr = etree.SubElement(grp, q("p:grpSpPr"))
    xfrm = etree.SubElement(grp_spr, q("a:xfrm"))
    for tag, attrs in (
        ("a:off", {"x": str(min_x), "y": str(min_y)}),
        ("a:ext", {"cx": str(cx), "cy": str(cy)}),
        ("a:chOff", {"x": str(min_x), "y": str(min_y)}),
        ("a:chExt", {"cx": str(cx), "cy": str(cy)}),
    ):
        node = etree.SubElement(xfrm, q(tag))
        for k, v in attrs.items():
            node.set(k, v)

    first = elements[0]
    first.addprevious(grp)
    for el in elements:
        sp_tree.remove(el)
        grp.append(el)
    return {"group": name or c_nv_pr.get("name"), "id": grp_id,
            "children": [etree.QName(e).localname for e in elements],
            "off": [min_x, min_y], "ext": [cx, cy],
            "children_ids": [int(e.find(".//" + q("p:cNvPr")).get("id")) for e in elements]}


def ungroup_shape(slide, group_slide_shape) -> dict:
    """解组：把子形状搬回父容器，按 ext/chExt 换算几何，并处理组的旋转/翻转。"""
    el = group_slide_shape._element
    if el.tag != q("p:grpSp"):
        raise ValueError(f"{group_slide_shape.name} 不是组合")
    parent = el.getparent()
    grp_spr = el.find(q("p:grpSpPr"))
    xfrm = grp_spr.find(q("a:xfrm")) if grp_spr is not None else None
    off = xfrm.find(q("a:off")) if xfrm is not None else None
    ext = xfrm.find(q("a:ext")) if xfrm is not None else None
    ch_off = xfrm.find(q("a:chOff")) if xfrm is not None else None
    ch_ext = xfrm.find(q("a:chExt")) if xfrm is not None else None
    gx, gy = (int(off.get("x")), int(off.get("y"))) if off is not None else (0, 0)
    gw, gh = (int(ext.get("cx")), int(ext.get("cy"))) if ext is not None else (0, 0)
    cx = int(ch_off.get("x")) if ch_off is not None else gx
    cy = int(ch_off.get("y")) if ch_off is not None else gy
    cw = int(ch_ext.get("cx")) if ch_ext is not None else gw
    ch = int(ch_ext.get("cy")) if ch_ext is not None else gh
    sx = (gw / cw) if cw else 1.0
    sy = (gh / ch) if ch else 1.0
    rot = int(xfrm.get("rot", 0)) / 60000.0 if xfrm is not None else 0.0
    flip_h = xfrm is not None and xfrm.get("flipH") in ("1", "true")
    flip_v = xfrm is not None and xfrm.get("flipV") in ("1", "true")

    kids = [k for k in el if isinstance(k.tag, str) and k.tag in
            (q("p:sp"), q("p:pic"), q("p:graphicFrame"), q("p:grpSp"), q("p:cxnSp"))]
    out = []
    gc_x, gc_y = gx + gw / 2, gy + gh / 2
    anchor = el
    anchor.addprevious(etree.Comment(" ungrouped "))
    for kid in kids:
        kx = kid.find(".//" + q("a:xfrm"))
        if kx is not None:
            koff = kx.find(q("a:off"))
            kext = kx.find(q("a:ext"))
            if koff is not None and kext is not None:
                nx = gx + (int(koff.get("x")) - cx) * sx
                ny = gy + (int(koff.get("y")) - cy) * sy
                nw = int(int(kext.get("cx")) * sx)
                nh = int(int(kext.get("cy")) * sy)
                if rot or flip_h or flip_v:
                    # 绕组中心旋转/翻转，再把自身 rot 累加
                    kcx, kcy = nx + nw / 2, ny + nh / 2
                    if flip_h:
                        kcx = 2 * gc_x - kcx
                    if flip_v:
                        kcy = 2 * gc_y - kcy
                    if rot:
                        theta = math.radians(rot)
                        dx, dy = kcx - gc_x, kcy - gc_y
                        kcx = gc_x + dx * math.cos(theta) - dy * math.sin(theta)
                        kcy = gc_y + dx * math.sin(theta) + dy * math.cos(theta)
                    nx, ny = kcx - nw / 2, kcy - nh / 2
                    kid_rot = int(kx.get("rot", 0)) / 60000.0 + rot
                    kx.set("rot", str(int(round(kid_rot * 60000)) % 21600000))
                    if flip_h:
                        kx.set("flipH", "0" if kx.get("flipH") in ("1", "true") else "1")
                    if flip_v:
                        kx.set("flipV", "0" if kx.get("flipV") in ("1", "true") else "1")
                koff.set("x", str(int(round(nx))))
                koff.set("y", str(int(round(ny))))
                kext.set("cx", str(nw))
                kext.set("cy", str(nh))
        el.remove(kid)
        anchor.addprevious(kid)
        nv = kid.find(".//" + q("p:cNvPr"))
        out.append({"id": int(nv.get("id")) if nv is not None else None,
                    "name": nv.get("name") if nv is not None else None})
    parent.remove(el)
    for c in [x for x in parent if isinstance(x.tag, str) is False and x.text == " ungrouped "]:
        parent.remove(c)
    return {"ungrouped": group_slide_shape.name, "children": out,
            "scale": [round(sx, 4), round(sy, 4)], "rot": rot}


# ------------------------------------------------------------------ 属性编辑

def _current_solid_color(sp_pr) -> str | None:
    """读现有纯色填充的 RGB hex（用于只改透明度）。"""
    solid = sp_pr.find(q("a:solidFill"))
    if solid is None or not len(solid):
        return None
    clr = solid[0]
    return clr.get("val") if clr.tag == q("a:srgbClr") else None


def _solid_fill_alpha(shape, alpha: float | None) -> None:
    """给纯色填充加/改 <a:alpha>（0..1）。"""
    if alpha is None:
        return
    sp_pr = shape._element.find(".//" + q("p:spPr"))
    if sp_pr is None:
        return
    solid = sp_pr.find(q("a:solidFill"))
    if solid is None:
        return
    clr = solid[0] if len(solid) else None
    if clr is None:
        return
    for old in clr.findall(q("a:alpha")):
        clr.remove(old)
    a = etree.SubElement(clr, q("a:alpha"))
    a.set("val", str(int(round(max(0.0, min(1.0, alpha)) * 100000))))


def set_props(shape, slide=None, **props) -> dict:
    """统一属性编辑入口。未知键会被忽略并写进 ignored。"""
    from pptx.dml.color import RGBColor

    applied, ignored = {}, []
    name = props.pop("name", None)
    if name:
        shape.name = str(name)
        applied["name"] = shape.name
    text = props.pop("text", None)
    if text is not None and getattr(shape, "has_text_frame", False):
        tf = shape.text_frame
        first = tf.paragraphs[0]
        if first.runs:
            first.runs[0].text = str(text)
            for extra in first.runs[1:]:
                extra._r.getparent().remove(extra._r)
        else:
            first.add_run().text = str(text)
        applied["text"] = str(text)
    lines = props.pop("lines", None)
    if lines is not None and getattr(shape, "has_text_frame", False):
        tf = shape.text_frame
        tf.clear()
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.add_run().text = str(line)
        applied["lines"] = len(lines)
    font = props.pop("font", None)
    size = props.pop("size", None)
    bold = props.pop("bold", None)
    italic = props.pop("italic", None)
    color = props.pop("color", None)
    align = props.pop("align", None)
    if any(v is not None for v in (font, size, bold, italic, color, align)) and getattr(shape, "has_text_frame", False):
        for para in shape.text_frame.paragraphs:
            if align:
                para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                                  "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}.get(
                    str(align).lower(), PP_ALIGN.LEFT)
            for run in para.runs:
                if font:
                    run.font.name = str(font)
                if size is not None:
                    run.font.size = Pt(float(size))
                if bold is not None:
                    run.font.bold = bool(bold)
                if italic is not None:
                    run.font.italic = bool(italic)
                if color:
                    run.font.color.rgb = RGBColor.from_string(str(color).lstrip("#").upper())
        applied.update({k: v for k, v in
                        {"font": font, "size": size, "bold": bold, "italic": italic,
                         "color": color, "align": align}.items() if v is not None})
    fill = props.pop("fill", None)
    alpha = props.pop("alpha", props.pop("fill_alpha", None))
    sp_pr = shape._element.find(q("p:spPr"))
    if fill is not None or alpha is not None:
        if sp_pr is None:
            ignored.append("fill(该形状没有 p:spPr)")
        else:
            if isinstance(fill, dict):
                spec = dict(fill)
                if alpha is not None:
                    spec.setdefault("alpha", alpha)
                applied["fill"] = dml.set_fill(sp_pr, spec)
            elif fill is not None:
                if alpha is not None:
                    applied["fill"] = dml.set_fill(sp_pr, {"color": str(fill), "alpha": alpha})
                else:
                    applied["fill"] = dml.set_fill(sp_pr, str(fill))
            else:
                cur = _current_solid_color(sp_pr)
                if cur:
                    applied["fill"] = dml.set_fill(sp_pr, {"color": cur, "alpha": alpha})
                    applied["fill"]["alpha"] = alpha
                else:
                    ignored.append("alpha(当前填充不是纯色)")
    line_keys = ("line_color", "line_width", "line_dash", "line_alpha", "line_head",
                 "line_tail", "line_join", "line_none", "line_cap", "arrow_size")
    if any(k in props for k in line_keys):
        spec = {
            "color": props.pop("line_color", None),
            "width": props.pop("line_width", None),
            "dash": props.pop("line_dash", None),
            "alpha": props.pop("line_alpha", None),
            "head": props.pop("line_head", None),
            "tail": props.pop("line_tail", None),
            "join": props.pop("line_join", None),
            "cap": props.pop("line_cap", None),
            "arrow_size": props.pop("arrow_size", None),
            "visible": not bool(props.pop("line_none", None)),
        }
        if sp_pr is None:
            ignored.append("line(该形状没有 p:spPr)")
        else:
            applied["line"] = dml.set_line(sp_pr, {k: v for k, v in spec.items() if v is not None})
    effects = props.pop("effects", None)
    if effects is not None:
        if sp_pr is None:
            ignored.append("effects(该形状没有 p:spPr)")
        else:
            applied["effects"] = dml.set_effects(sp_pr, effects)
    crop = props.pop("crop", None)
    if crop is not None:
        applied["crop"] = dml.set_crop(shape._element, crop)
    brightness = props.pop("brightness", None)
    contrast = props.pop("contrast", None)
    if brightness is not None or contrast is not None:
        applied["lum"] = dml.set_lum(shape._element, brightness, contrast)
    paragraphs = props.pop("paragraphs", None)
    runs = props.pop("runs", None)
    if paragraphs is not None or runs is not None:
        if not getattr(shape, "has_text_frame", False):
            ignored.append("paragraphs(该形状没有文本)")

        else:
            body = shape.text_frame._txBody
            payload = paragraphs if paragraphs is not None else [{"runs": runs}]
            text_defaults = {k: props.get(k) for k in ("font", "size", "bold", "italic", "color")}
            applied["paragraphs"] = dml.set_text_body(body, payload, text_defaults)
    rot = props.pop("rot", props.pop("rotation", None))
    if rot is not None:
        shape.rotation = float(rot) % 360
        applied["rotation"] = shape.rotation
    flip_h = props.pop("flip_h", None)
    flip_v = props.pop("flip_v", None)
    if flip_h is not None or flip_v is not None:
        xfrm = _xfrm(shape)
        if flip_h is not None:
            xfrm.set("flipH", "1" if flip_h else "0")
        if flip_v is not None:
            xfrm.set("flipV", "1" if flip_v else "0")
        applied["flip"] = [bool(flip_h), bool(flip_v)]
    z = props.pop("z_order", props.pop("z", None))
    if z is not None:
        el = shape._element
        parent = el.getparent()
        if str(z).lower() in ("front", "top", "最前", "置顶"):
            parent.remove(el)
            parent.append(el)
        elif str(z).lower() in ("back", "bottom", "最后", "置底"):
            parent.remove(el)
            parent.insert(2 if parent.tag == q("p:spTree") else 0, el)
        elif str(z).lower() in ("forward", "up", "上移"):
            nxt = el.getnext()
            if nxt is not None:
                parent.remove(el)
                nxt.addnext(el)
        elif str(z).lower() in ("backward", "down", "下移"):
            prev = el.getprevious()
            if prev is not None:
                parent.remove(el)
                prev.addprevious(el)
        applied["z_order"] = z
    for key in list(props):
        ignored.append(key)
    return {"shape": shape.name, "id": shape.shape_id, "applied": applied, "ignored": ignored}


# ------------------------------------------------------------------ 对齐 / 分布

ALIGN_MODES = {
    "left": ("x", 0.0), "right": ("x", 1.0), "center-h": ("x", 0.5), "centerx": ("x", 0.5),
    "top": ("y", 0.0), "bottom": ("y", 1.0), "center-v": ("y", 0.5), "centery": ("y", 0.5),
    "center": ("c", 0.5),
}


def align_shapes(slide, shapes: list, mode: str, to: str | None = None,
                 slide_w: int = 0, slide_h: int = 0, spacing=None) -> dict:
    """对齐/分布。to 可为形状名或 "slide"；spacing 用于分布间距。"""
    if not shapes:
        return {"aligned": 0}
    mode_key = str(mode).strip().lower()
    if to in (None, "", "slide", "页面"):
        base = (0, 0, slide_w, slide_h)
    else:
        ref = find_shape(slide, to)
        base = box_of(ref)
    bx, by, bw, bh = base

    def shift_to(axis: str, frac: float) -> int:
        moved = 0
        for shp in shapes:
            left, top, width, height = box_of(shp)
            if axis == "x":
                target = bx + frac * bw - frac * width
                _set_box(shp, x=int(target))
            else:
                target = by + frac * bh - frac * height
                _set_box(shp, y=int(target))
            moved += 1
        return moved

    if mode_key in ("distribute-h", "distribute-v", "distribute"):
        horizontal = mode_key != "distribute-v"
        ordered = sorted(shapes, key=lambda s: box_of(s)[0] if horizontal else box_of(s)[1])
        if len(ordered) < 2:
            return {"aligned": 0}
        first, last = box_of(ordered[0]), box_of(ordered[-1])
        if spacing is not None:
            gap = to_emu(spacing, slide_w if horizontal else slide_h)
            cursor = (first[0] + first[2] + gap) if horizontal else (first[1] + first[3] + gap)
            for shp in ordered[1:-1]:
                left, top, width, height = box_of(shp)
                if horizontal:
                    _set_box(shp, x=int(cursor))
                    cursor += width + gap
                else:
                    _set_box(shp, y=int(cursor))
                    cursor += height + gap
        else:
            total = (last[0] + last[2] - first[0]) if horizontal else (last[1] + last[3] - first[1])
            sizes = sum(box_of(s)[2 if horizontal else 3] for s in ordered[1:-1])
            gap = (total - (first[2] if horizontal else first[3])
                   - (last[2] if horizontal else last[3]) - sizes) / (len(ordered) - 1)
            cursor = (first[0] + first[2] + gap) if horizontal else (first[1] + first[3] + gap)
            for shp in ordered[1:-1]:
                left, top, width, height = box_of(shp)
                if horizontal:
                    _set_box(shp, x=int(cursor))
                    cursor += width + gap
                else:
                    _set_box(shp, y=int(cursor))
                    cursor += height + gap
        return {"distributed": len(ordered), "axis": "h" if horizontal else "v",
                "spacing": round(gap)}

    if mode_key not in ALIGN_MODES:
        raise KeyError(f"未知对齐方式 {mode!r}；可用: {sorted(ALIGN_MODES)}")
    axis, frac = ALIGN_MODES[mode_key]
    if axis == "c":
        shift_to("x", 0.5)
        shift_to("y", 0.5)
        return {"aligned": len(shapes), "mode": "center"}
    return {"aligned": shift_to(axis, frac), "mode": mode_key}


# ------------------------------------------------------------------ 新建元素

SHAPE_KINDS = {
    "rect": MSO_SHAPE.RECTANGLE, "rectangle": MSO_SHAPE.RECTANGLE, "矩形": MSO_SHAPE.RECTANGLE,
    "round-rect": MSO_SHAPE.ROUNDED_RECTANGLE, "圆角矩形": MSO_SHAPE.ROUNDED_RECTANGLE,
    "ellipse": MSO_SHAPE.OVAL, "oval": MSO_SHAPE.OVAL, "circle": MSO_SHAPE.OVAL, "圆": MSO_SHAPE.OVAL,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE, "三角形": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "line": MSO_SHAPE.RECTANGLE, "线条": MSO_SHAPE.RECTANGLE,
    "arrow": MSO_SHAPE.RIGHT_ARROW, "箭头": MSO_SHAPE.RIGHT_ARROW,
    "chevron": MSO_SHAPE.CHEVRON, "人字": MSO_SHAPE.CHEVRON,
    "pentagon": MSO_SHAPE.PENTAGON, "五边形": MSO_SHAPE.PENTAGON,
    "star": MSO_SHAPE.STAR_5_POINT, "星形": MSO_SHAPE.STAR_5_POINT,
    "diamond": MSO_SHAPE.DIAMOND, "菱形": MSO_SHAPE.DIAMOND,
}


@dataclass
class ElementSpec:
    name: str
    kind: str = "rect"
    x: object = 0
    y: object = 0
    w: object = 4.0
    h: object = 1.0
    text: str | None = None
    fill: str | None = None
    line_color: str | None = None
    line_width: float | None = None
    alpha: float | None = None
    rot: float | None = None
    flip_h: bool | None = None
    flip_v: bool | None = None
    font: str | None = None
    size: float | None = None
    bold: bool | None = None
    color: str | None = None
    align: str | None = None
    anchor: str = "left-top"
    group: str | None = None
    z_order: object = None
    image: str | None = None


def add_element(slide, spec: dict, theme=None, slide_w: int = 0, slide_h: int = 0) -> dict:
    """按 spec 新建形状/文本框/图片/表格/图表。x,y,w,h 支持 12% / "50%" / "1.5in" / 120pt。"""
    from pptx.dml.color import RGBColor

    kind = str(spec.get("kind") or spec.get("type") or "rect").strip().lower()

    # 表格 / 图表走专门构建器（返回同样的记录结构）
    if kind in ("table", "表格"):
        from .table import build_table
        info = build_table(slide.shapes, spec, theme, slide_w, slide_h)
        return {**info, "kind": "table", "applied": {"style": info.get("style")},
                "group": spec.get("group"),
                "_shape": find_shape(slide, f"shape:{info['id']}")}
    if kind in ("chart", "图表"):
        from .chart import build_chart
        info = build_chart(slide.shapes, spec, theme, slide_w, slide_h)
        return {**info, "kind": "chart", "applied": {"series": info.get("series")},
                "group": spec.get("group"),
                "_shape": find_shape(slide, f"shape:{info['id']}")}

    name = str(spec.get("name") or f"{kind}-{len(slide.shapes) + 1}")
    w = to_emu(spec.get("w", spec.get("width", 4.0)), slide_w)
    h = to_emu(spec.get("h", spec.get("height", 1.0)), slide_h)
    ax, ay = anchor_fraction(spec.get("anchor"))
    x = to_emu(spec.get("x", 0), slide_w) - ax * w
    y = to_emu(spec.get("y", 0), slide_h) - ay * h

    if kind in ("text", "textbox", "文本", "文本框"):
        shp = slide.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(w), Emu(h))
        shp.text_frame.word_wrap = True
        if spec.get("text"):
            shp.text_frame.text = str(spec["text"])
    elif kind in ("image", "pic", "图片"):
        path = spec.get("image") or spec.get("path")
        shp = slide.shapes.add_picture(path, Emu(int(x)), Emu(int(y)), Emu(w), Emu(h))
    elif kind in ("line", "线条", "connector", "连接线"):
        x1, y1, x2, y2 = x, y, x + w, y + h
        if x2 < x1:
            x1, x2 = x2, x1
        if y2 < y1:
            y1, y2 = y2, y1
        shp = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Emu(int(x1)), Emu(int(y1)),
                                         Emu(int(x2)), Emu(int(y2)))
        if spec.get("line_color") is None:
            spec.setdefault("line_color", spec.get("color") or "FFFFFF")
    else:
        shape_type = SHAPE_KINDS.get(kind, MSO_SHAPE.RECTANGLE)
        shp = slide.shapes.add_shape(shape_type, Emu(int(x)), Emu(int(y)), Emu(w), Emu(h))
        shp.shadow.inherit = False
        if spec.get("text"):
            shp.text_frame.text = str(spec["text"])
    shp.name = name

    # 除结构键外的属性全部透传给 set_props（fill / effects / line_* / paragraphs / crop …）
    reserved = {"name", "kind", "type", "x", "y", "w", "h", "width", "height",
                "anchor", "group", "image", "path", "z", "text"}
    props = {k: v for k, v in spec.items() if k not in reserved}
    applied = set_props(shp, **props) if props else {"applied": {}}
    return {"name": name, "id": shp.shape_id, "kind": kind,
            "box": [int(shp.left), int(shp.top), int(shp.width), int(shp.height)],
            "applied": applied.get("applied", {}), "group": spec.get("group"),
            "_shape": shp}


# ------------------------------------------------------------------ 复制 / 删除

def _renumber(element, start: int) -> int:
    """给元素内所有 cNvPr 重新分配 id，返回用到的最大 id。"""
    nid = start
    for nv in element.iter(q("p:cNvPr")):
        nid += 1
        nv.set("id", str(nid))
    return nid


def duplicate_shape(slide, shape, name: str | None = None, dx=None, dy=None,
                    slide_w: int = 0, slide_h: int = 0) -> dict:
    """复制形状（组会连同子形状一起复制并重编号）。"""
    src = shape._element
    new_el = copy.deepcopy(src)
    max_id = _renumber(new_el, _max_shape_id(slide))
    nv = new_el.find(".//" + q("p:cNvPr"))
    if nv is not None:
        nv.set("name", str(name or f"{shape.name} copy"))
    src.addnext(new_el)
    target_id = int(new_el.find(".//" + q("p:cNvPr")).get("id"))
    new_shape = find_shape(slide, f"shape:{target_id}")
    if dx is not None or dy is not None:
        move(new_shape, dx=dx, dy=dy, slide_w=slide_w, slide_h=slide_h)
    return {"name": new_shape.name, "id": target_id, "box": box_of(new_shape),
            "max_id": max_id}


def delete_shape(shape) -> dict:
    el = shape._element
    parent = el.getparent()
    if parent is None:
        raise ValueError(f"{shape.name} 没有父节点")
    parent.remove(el)
    return {"deleted": shape.name, "id": shape.shape_id}


# ------------------------------------------------------------------ ops 调度

def _shape_or_raise(slide, target):
    return find_shape(slide, target)


def apply_ops(slide, ops: list[dict], slide_w: int = 0, slide_h: int = 0) -> list[dict]:
    """执行一串形状操作：place/move/resize/props/group/ungroup/align/duplicate/delete。

    每个 op 独立 try，失败只记录不中断后续（返回里 ok=False + error）。
    """
    results: list[dict] = []
    for i, op in enumerate(ops or []):
        kind = str(op.get("op") or op.get("action") or "").strip().lower()
        rec = {"index": i, "op": kind, "ok": True}
        try:
            if kind in ("place", "定位", "at"):
                shp = _shape_or_raise(slide, op.get("target"))
                rec["result"] = place(shp, x=op.get("x"), y=op.get("y"),
                                      w=op.get("w"), h=op.get("h"),
                                      anchor=op.get("anchor", "left-top"),
                                      slide_w=slide_w, slide_h=slide_h,
                                      of=find_shape(slide, op["of"]) if op.get("of") else None)
            elif kind in ("move", "移动", "offset"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = [move(s, dx=op.get("dx"), dy=op.get("dy"),
                                      to_x=op.get("to_x"), to_y=op.get("to_y"),
                                      anchor=op.get("anchor", "left-top"),
                                      slide_w=slide_w, slide_h=slide_h) for s in targets]
            elif kind in ("resize", "scale", "缩放", "尺寸"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = [resize(s, scale=op.get("scale"),
                                        w=op.get("w", op.get("width")),
                                        h=op.get("h", op.get("height")),
                                        keep_aspect=bool(op.get("keep_aspect", True)),
                                        anchor=op.get("anchor", "left-top"),
                                        slide_w=slide_w, slide_h=slide_h) for s in targets]
            elif kind in ("props", "set", "属性", "edit"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                payload = {k: v for k, v in op.items()
                           if k not in ("op", "action", "target", "targets")}
                rec["result"] = [set_props(s, **payload) for s in targets]
            elif kind in ("group", "组合"):
                members = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = group_shapes(slide, members, name=op.get("name"))
            elif kind in ("ungroup", "解组"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = [ungroup_shape(slide, s) for s in targets
                                 if s.__class__.__name__ == "GroupShape"]
            elif kind in ("align", "对齐"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = align_shapes(slide, targets, op.get("mode", "left"),
                                             to=op.get("to"), slide_w=slide_w,
                                             slide_h=slide_h, spacing=op.get("spacing"))
            elif kind in ("duplicate", "copy", "复制"):
                src = _shape_or_raise(slide, op.get("target"))
                rec["result"] = duplicate_shape(slide, src, name=op.get("name"),
                                                dx=op.get("dx"), dy=op.get("dy"),
                                                slide_w=slide_w, slide_h=slide_h)
            elif kind in ("table", "表格", "cell", "cells", "merge", "table_style", "widths", "heights"):
                from .table import edit_table as _edit_table
                rec["result"] = _edit_table(slide, [op])
            elif kind in ("delete", "remove", "删除"):
                targets = find_shapes(slide, op.get("targets") or op.get("target"))
                rec["result"] = [delete_shape(s) for s in targets]
            else:
                rec["ok"] = False
                rec["error"] = (f"未知操作 {kind!r}；可用: place/move/resize/props/"
                                f"group/ungroup/align/duplicate/delete")
        except Exception as exc:  # noqa: BLE001
            rec["ok"] = False
            rec["error"] = f"{type(exc).__name__}: {exc}"
        results.append(rec)
    return results


# ------------------------------------------------------------------ 读取

def shape_info(shape, deep: bool = False) -> dict:
    """读出一个形状的几何与外观，供 AI 校验自己的操作结果。"""
    left, top, width, height = box_of(shape)
    rec = {
        "name": shape.name,
        "id": shape.shape_id,
        "type": shape.__class__.__name__,
        "left_in": emu_to_in(left), "top_in": emu_to_in(top),
        "width_in": emu_to_in(width), "height_in": emu_to_in(height),
        "left_emu": left, "top_emu": top, "width_emu": width, "height_emu": height,
        "rot": round(float(getattr(shape, "rotation", 0) or 0), 2),
    }
    xfrm = shape._element.find(".//" + q("a:xfrm"))
    if xfrm is not None:
        rec["flip"] = [xfrm.get("flipH") in ("1", "true"), xfrm.get("flipV") in ("1", "true")]
    if getattr(shape, "has_text_frame", False):
        tf = shape.text_frame
        rec["text"] = tf.text[:300]
        rec["paragraphs"] = len(tf.paragraphs)
    if shape.__class__.__name__ == "GroupShape":
        rec["group_items"] = [shape_info(s) for s in shape.shapes]
        grp_pr = shape._element.find(q("p:grpSpPr"))
        gx = grp_pr.find(q("a:xfrm")) if grp_pr is not None else None
        if gx is not None:
            ch_off = gx.find(q("a:chOff"))
            ch_ext = gx.find(q("a:chExt"))
            if ch_off is not None:
                rec["ch_off"] = [int(ch_off.get("x")), int(ch_off.get("y"))]
            if ch_ext is not None:
                rec["ch_ext"] = [int(ch_ext.get("cx")), int(ch_ext.get("cy"))]
    if deep:
        rec["xml"] = shape._element.xml[:4000]
    return rec


def slide_shapes_info(slide, target=None, deep: bool = False) -> list[dict]:
    """列出某页形状（可按 target 过滤）。"""
    if target:
        return [shape_info(find_shape(slide, target), deep=deep)]
    return [shape_info(s) for s in iter_shapes(slide.shapes)]