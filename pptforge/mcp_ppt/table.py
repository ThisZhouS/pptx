"""表格：单元格级别的内容、填充、边框、合并、行高列宽、表样式。

真机结构（rev/out/gt_table_border.pptx、gt_table_chart.pptx 的 slide_001.xml）：

  p:graphicFrame/a:graphic/a:graphicData(uri=…/table)/a:tbl
    a:tblPr firstRow / bandRow / firstCol / lastRow + a:tableStyleId{GUID}
    a:tblGrid/a:gridCol@w          列宽（EMU）
    a:tr@h                         行高（EMU）
      a:tc
        a:txBody / a:p / a:r / a:rPr / a:t
        a:tcPr
          a:lnL|a:lnR|a:lnT|a:lnB   @w(EMU) @cmpd 内嵌 solidFill + prstDash
          a:lnTlToBr / a:lnBlToTr   对角线
          a:solidFill …             单元格填充（在边框之后）
    合并：主单元格 gridSpan="2"；被吞掉的单元格保留 a:tc 并写 hMerge="1"（纵向 rowSpan/vMerge）
"""
from __future__ import annotations

from lxml import etree
from pptx.util import Emu, Inches, Length

from .ns import q
from . import dml

TC_PR_ORDER = ["a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:lnTlToBr", "a:lnBlToTr", "a:cell3D",
               "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
               "a:grpFill", "a:headers", "a:extLst"]

SIDES = {"left": "a:lnL", "right": "a:lnR", "top": "a:lnT", "bottom": "a:lnB",
         "diag_down": "a:lnTlToBr", "diag_up": "a:lnBlToTr"}


def _tc_pr(cell):
    tc = cell._tc
    tcpr = tc.tcPr
    if tcpr is None:
        tcpr = tc.get_or_add_tcPr()
    return tcpr


def set_borders(cell, borders: dict) -> dict:
    """borders: {"top": {"color":"#RRGGBB","width":1.0,"dash":"solid","none":False}, ...}"""
    tcpr = _tc_pr(cell)
    applied = {}
    for side, tag in SIDES.items():
        spec = borders.get(side)
        if spec is None:
            continue
        old = tcpr.find(q(tag))
        if old is not None:
            tcpr.remove(old)
        if str(spec).lower() in ("none", "false", "0") or (isinstance(spec, dict) and spec.get("none")):
            node = etree.Element(q(tag))
            node.set("w", "0")
            dml.insert_ordered(tcpr, node, TC_PR_ORDER)
            node.append(etree.Element(q("a:noFill")))
            applied[side] = "none"
            continue
        conf = spec if isinstance(spec, dict) else {}
        node = etree.Element(q(tag))
        node.set("w", str(int(round(float(conf.get("width", 1.0)) * 12700))))
        node.set("cmpd", "sng")
        dml.insert_ordered(tcpr, node, TC_PR_ORDER)
        node.append(dml._clr(conf.get("color", "#FFFFFF")))
        dash = etree.Element(q("a:prstDash"))
        dash.set("val", str(conf.get("dash", "solid")))
        node.append(dash)
        applied[side] = {"color": conf.get("color", "#FFFFFF"), "width": conf.get("width", 1.0)}
    return applied


def set_cell(cell, spec: dict, theme=None) -> dict:
    """单元格：text / fill / alpha / font / size / bold / color / align / valign / margins / borders。"""
    applied = {}
    text = spec.get("text")
    if text is not None:
        cell.text_frame.text = str(text)
        applied["text"] = str(text)
    lines = spec.get("lines")
    if lines is not None:
        tf = cell.text_frame
        tf.clear()
        for i, line in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.add_run().text = str(line)
        applied["lines"] = len(lines)
    fill = spec.get("fill")
    if fill is not None:
        if isinstance(fill, dict) and str(fill.get("type", "solid")).lower() == "gradient":
            applied["fill"] = dml.set_fill(_tc_pr(cell), fill)
        elif str(fill).lower() in ("none", "transparent", "无"):
            cell.fill.background()
            applied["fill"] = "none"
        else:
            cell.fill.solid()
            from pptx.dml.color import RGBColor
            cell.fill.fore_color.rgb = RGBColor.from_string(str(fill).lstrip("#").upper())
            applied["fill"] = str(fill)
    alpha = spec.get("alpha")
    if alpha is not None:
        _tc_alpha(cell, float(alpha))
        applied["alpha"] = float(alpha)
    align = spec.get("align")
    valign = spec.get("valign", spec.get("vertical"))
    for para in cell.text_frame.paragraphs:
        if align:
            from pptx.enum.text import PP_ALIGN
            para.alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                              "right": PP_ALIGN.RIGHT}.get(str(align).lower(), PP_ALIGN.LEFT)
            applied["align"] = align
        for run in para.runs:
            if spec.get("size") is not None:
                from pptx.util import Pt
                run.font.size = Pt(float(spec["size"]))
            if spec.get("bold") is not None:
                run.font.bold = bool(spec["bold"])
            if spec.get("italic") is not None:
                run.font.italic = bool(spec["italic"])
            if spec.get("color"):
                from pptx.dml.color import RGBColor
                run.font.color.rgb = RGBColor.from_string(str(spec["color"]).lstrip("#").upper())
            if spec.get("font"):
                run.font.name = str(spec["font"])
        if spec.get("size") is not None or spec.get("bold") is not None:
            applied["font"] = {"size": spec.get("size"), "bold": spec.get("bold"),
                               "color": spec.get("color")}
    if valign:
        from pptx.enum.text import MSO_ANCHOR
        cell.vertical_anchor = {"top": MSO_ANCHOR.TOP, "middle": MSO_ANCHOR.MIDDLE,
                                "bottom": MSO_ANCHOR.BOTTOM}.get(str(valign).lower(), MSO_ANCHOR.MIDDLE)
        applied["valign"] = valign
    margins = spec.get("margins")
    if margins:
        from pptx.util import Pt
        cell.margin_left = Pt(float(margins.get("left", 0)))
        cell.margin_right = Pt(float(margins.get("right", 0)))
        cell.margin_top = Pt(float(margins.get("top", 0)))
        cell.margin_bottom = Pt(float(margins.get("bottom", 0)))
    if spec.get("borders"):
        applied["borders"] = set_borders(cell, spec["borders"])
    return applied


def _tc_alpha(cell, alpha: float) -> None:
    tcpr = _tc_pr(cell)
    solid = tcpr.find(q("a:solidFill"))
    if solid is None or not len(solid):
        return
    clr = solid[0]
    for old in clr.findall(q("a:alpha")):
        clr.remove(old)
    a = etree.Element(q("a:alpha"))
    a.set("val", str(int(round(max(0.0, min(1.0, alpha)) * 100000))))
    clr.append(a)


def style_table(table, spec: dict) -> dict:
    """表样式：first_row / band_row / first_col / last_row + tableStyleId。"""
    applied = {}
    mapping = {"first_row": "firstRow", "band_row": "bandRow", "first_col": "firstCol",
               "last_row": "lastRow", "band_col": "bandCol"}
    for key, attr in mapping.items():
        if key in spec:
            setattr(table, key, bool(spec[key]))
            applied[key] = bool(spec[key])
    style_id = spec.get("style_id") or spec.get("tableStyleId")
    if style_id:
        tbl = table._tbl
        tbl_pr = tbl.tblPr
        old = tbl_pr.find(q("a:tableStyleId"))
        if old is not None:
            tbl_pr.remove(old)
        node = etree.SubElement(tbl_pr, q("a:tableStyleId"))
        node.text = str(style_id)
        applied["style_id"] = str(style_id)
    return applied


def column_widths(table, widths: list) -> dict:
    applied = {}
    for i, w in enumerate(widths):
        if i >= len(table.columns):
            break
        table.columns[i].width = Emu(int(float(w))) if isinstance(w, (int, float)) and w > 100 else Inches(float(w))
        applied[i] = int(table.columns[i].width)
    return applied


def row_heights(table, heights: list) -> dict:
    applied = {}
    for i, h in enumerate(heights):
        if i >= len(table.rows):
            break
        table.rows[i].height = Emu(int(float(h))) if isinstance(h, (int, float)) and h > 100 else Inches(float(h))
        applied[i] = int(table.rows[i].height)
    return applied


def build_table(shapes, spec: dict, theme=None, slide_w: int = 0, slide_h: int = 0) -> dict:
    """按 rows + 样式创建表格。spec: {name, x,y,w,h, rows, col_widths, row_heights, header, style, cell}"""
    rows = spec.get("rows") or [["A", "B"], ["1", "2"]]
    n_rows = len(rows)
    n_cols = max(len(r) for r in rows)
    from .shapes import anchor_fraction, to_emu_len

    w = to_emu_len(spec.get("w", 10.0), slide_w)
    h = to_emu_len(spec.get("h", 3.0), slide_h)
    ax, ay = anchor_fraction(spec.get("anchor"))
    x = to_emu_len(spec.get("x", 0), slide_w) - ax * w
    y = to_emu_len(spec.get("y", 0), slide_h) - ay * h
    frame = shapes.add_table(n_rows, n_cols, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    frame.name = spec.get("name") or "table"
    table = frame.table

    base = dict(spec.get("cell") or {})
    header = spec.get("header")
    header_style = dict(spec.get("header_cell") or {})
    body_style = dict(spec.get("body_cell") or {})
    for r, row in enumerate(rows):
        for c in range(n_cols):
            cell = table.cell(r, c)
            conf = dict(base)
            if header and r == 0:
                conf.update(header_style)
            elif not (header and r == 0):
                conf.update(body_style)
            if r == 0 and header and "bold" not in conf:
                conf["bold"] = True
            conf["text"] = str(row[c]) if c < len(row) else ""
            merged = dict(conf)
            merged.pop("borders", None)
            set_cell(cell, merged, theme)
            if conf.get("borders"):
                set_borders(cell, conf["borders"])

    if spec.get("col_widths"):
        column_widths(table, spec["col_widths"])
    if spec.get("row_heights"):
        row_heights(table, spec["row_heights"])
    style_conf = dict(spec.get("style") or {})
    if header:
        style_conf.setdefault("first_row", True)
    if spec.get("banding", True) and header:
        style_conf.setdefault("band_row", True)
    style_info = style_table(table, style_conf) if style_conf else {}

    return {"name": frame.name, "id": frame.shape_id,
            "rows": n_rows, "cols": n_cols,
            "box": [int(frame.left), int(frame.top), int(frame.width), int(frame.height)],
            "style": style_info}


def _find_table(slide, target):
    from .shapes import find_shape
    shp = find_shape(slide, target)
    if not getattr(shp, "has_table", False):
        raise ValueError(f"{target!r} 不是表格")
    return shp, shp.table


def edit_table(slide, ops: list[dict]) -> list[dict]:
    """表格编辑 ops：cell / merge / style / widths / heights；也接受合并写法 op=table。"""
    results = []
    for i, op in enumerate(ops or []):
        kind = str(op.get("op") or op.get("action") or "cell").lower()
        if kind in ("table", "表格"):
            target = op.get("target") or op.get("table")
            sub: list[dict] = []
            if op.get("cells"):
                sub.append({"op": "cell", "target": target, "cells": op["cells"]})
            if op.get("merge"):
                m = op["merge"]
                if isinstance(m, dict):
                    sub.append({"op": "merge", "target": target, **m})
                else:
                    sub.append({"op": "merge", "target": target, "ranges": m})
            if op.get("style"):
                sub.append({"op": "style", "target": target, "style": op["style"]})
            if op.get("widths") or op.get("col_widths"):
                sub.append({"op": "widths", "target": target,
                            "widths": op.get("widths") or op.get("col_widths")})
            if op.get("heights") or op.get("row_heights"):
                sub.append({"op": "heights", "target": target,
                            "heights": op.get("heights") or op.get("row_heights")})
            results.append({"index": i, "op": kind, "ok": True,
                            "result": edit_table(slide, sub) if sub else {"noop": True}})
            continue
        rec = {"index": i, "op": kind, "ok": True}
        try:
            frame, table = _find_table(slide, op.get("target") or op.get("table"))
            if kind in ("cell", "cells", "单元格"):
                cells = op.get("cells") or [op]
                applied = []
                for item in cells:
                    r = int(item.get("row", item.get("r", 0)))
                    c = int(item.get("col", item.get("c", 0)))
                    applied.append(set_cell(table.cell(r, c), item))
                rec["result"] = applied
            elif kind in ("merge", "合并"):
                ranges = op.get("ranges")
                merged = []
                if ranges:
                    for rg in ranges:
                        table.cell(int(rg[0]), int(rg[1])).merge(table.cell(int(rg[2]), int(rg[3])))
                        merged.append(rg)
                else:
                    r1 = int(op.get("row1", op.get("r1", 0)))
                    c1 = int(op.get("col1", op.get("c1", 0)))
                    r2 = int(op.get("row2", op.get("r2", r1)))
                    c2 = int(op.get("col2", op.get("c2", c1)))
                    table.cell(r1, c1).merge(table.cell(r2, c2))
                    merged.append([r1, c1, r2, c2])
                rec["result"] = {"merged": merged}
            elif kind in ("style", "样式"):
                rec["result"] = style_table(table, op.get("style") or op)
            elif kind in ("widths", "列宽"):
                rec["result"] = column_widths(table, op.get("widths") or op.get("col_widths") or [])
            elif kind in ("heights", "行高"):
                rec["result"] = row_heights(table, op.get("heights") or op.get("row_heights") or [])
            else:
                rec["ok"] = False
                rec["error"] = f"未知表格操作 {kind!r}；可用 cell/merge/style/widths/heights"
        except Exception as exc:  # noqa: BLE001
            rec["ok"] = False
            rec["error"] = f"{type(exc).__name__}: {exc}"
        results.append(rec)
    return results


def table_info(slide, target) -> dict:
    frame, table = _find_table(slide, target)
    grid = [int(col.width) for col in table.columns]
    rows = [int(row.height) for row in table.rows]
    cells = []
    for r, row in enumerate(table.rows):
        for c, cell in enumerate(row.cells):
            cells.append({"r": r, "c": c, "text": cell.text[:60]})
    return {"name": frame.name, "id": frame.shape_id,
            "rows": len(table.rows), "cols": len(table.columns),
            "col_widths": grid, "row_heights": rows, "cells": cells}
