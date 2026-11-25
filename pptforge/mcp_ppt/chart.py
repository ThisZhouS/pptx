"""图表：原生 chart（数据内嵌工作簿），含标题/图例/数据标签/配色/轴格式。

图表部件结构（python-pptx 生成，PowerPoint 完全接受）：
  ppt/charts/chartN.xml            图表定义（c:chartSpace/c:chart/c:plotArea/c:barChart…）
  ppt/embeddings/Microsoft_Excel_SheetN.xlsx   数据源（图表可"编辑数据"）
  p:graphicFrame/a:graphic/a:graphicData(uri=…/chart)/c:chart r:id="rIdN"

动画：图表作为 graphicFrame，动画目标就是它自己的 shape id；
      PowerPoint 的"按系列/按类别"分级动画在 COM 里 AddEffect 一律返回 Invalid request
     （已实测 7/8/10 三档全失败），因此引擎只支持整体动画（淡入/擦除/缩放…）。
"""
from __future__ import annotations

from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.util import Emu, Length, Pt

from .shapes import anchor_fraction, to_emu_len

CHART_TYPES = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column-clustered": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column-stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "column-stacked-100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar-stacked": XL_CHART_TYPE.BAR_STACKED,
    "line": XL_CHART_TYPE.LINE,
    "line-markers": XL_CHART_TYPE.LINE_MARKERS,
    "area": XL_CHART_TYPE.AREA,
    "area-stacked": XL_CHART_TYPE.AREA_STACKED,
    "pie": XL_CHART_TYPE.PIE,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "radar": XL_CHART_TYPE.RADAR,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
}

LEGEND_POS = {
    "bottom": XL_LEGEND_POSITION.BOTTOM, "top": XL_LEGEND_POSITION.TOP,
    "left": XL_LEGEND_POSITION.LEFT, "right": XL_LEGEND_POSITION.RIGHT,
    "none": None, "false": None,
}

PALETTE = ["#4C8DFF", "#6FE3C4", "#FFD166", "#FF7A45", "#B388FF",
           "#36C2FF", "#F06292", "#9CCC65", "#FFB74D", "#4DD0E1"]


def _rgb(value: str) -> RGBColor:
    return RGBColor.from_string(str(value).lstrip("#").upper())


def build_chart(shapes, spec: dict, theme=None, slide_w: int = 0, slide_h: int = 0) -> dict:
    """spec: {name, type, categories, series:[{name, values, color, colors}], title, legend,
             data_labels, x, y, w, h, font_size, number_format, gap_width, overlap}"""
    chart_type = CHART_TYPES.get(str(spec.get("type") or "column").lower())
    if chart_type is None:
        raise KeyError(f"未知图表类型 {spec.get('type')!r}；可用: {sorted(CHART_TYPES)}")
    categories = spec.get("categories") or [f"C{i + 1}" for i in range(4)]
    series_specs = spec.get("series") or [{"name": "系列 1", "values": [1, 2, 3, 4]}]

    data = CategoryChartData()
    data.categories = categories
    for s in series_specs:
        values = s.get("values") or [0] * len(categories)
        data.add_series(str(s.get("name") or "系列"), tuple(values))

    w = to_emu_len(spec.get("w", 11.0), slide_w)
    h = to_emu_len(spec.get("h", 5.0), slide_h)
    ax, ay = anchor_fraction(spec.get("anchor"))
    x = to_emu_len(spec.get("x", 0), slide_w) - ax * w
    y = to_emu_len(spec.get("y", 0), slide_h) - ay * h
    frame = shapes.add_chart(chart_type, Emu(int(x)), Emu(int(y)), Emu(w), Emu(h), data)
    frame.name = spec.get("name") or "chart"
    chart = frame.chart

    # 标题
    title_text = spec.get("title")
    if title_text:
        chart.has_title = True
        chart.chart_title.text_frame.text = str(title_text)
        for run in chart.chart_title.text_frame.paragraphs[0].runs:
            run.font.size = Pt(float(spec.get("title_size", 16)))
            run.font.bold = True
            if theme:
                run.font.color.rgb = _rgb(theme.fg)
    else:
        chart.has_title = False

    # 图例
    legend = spec.get("legend", "bottom")
    pos = LEGEND_POS.get(str(legend).lower(), XL_LEGEND_POSITION.BOTTOM)
    chart.has_legend = pos is not None
    if pos is not None:
        chart.legend.position = pos
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(float(spec.get("legend_size", 11)))
        if theme:
            chart.legend.font.color.rgb = _rgb(theme.muted)

    # 数据标签
    plots = chart.plots
    if spec.get("data_labels"):
        for plot in plots:
            plot.has_data_labels = True
            dl = plot.data_labels
            dl.show_value = True
            if spec.get("number_format"):
                dl.number_format = str(spec["number_format"])
                dl.number_format_is_linked = False
            dl.font.size = Pt(float(spec.get("label_size", 10)))
            if theme:
                dl.font.color.rgb = _rgb(theme.fg)

    # 系列配色
    is_pie = chart_type in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT)
    palette = list(PALETTE)
    for si, s in enumerate(chart.series):
        colors = None
        if si < len(series_specs):
            colors = series_specs[si].get("colors")
            single = series_specs[si].get("color")
        else:
            single = None
        if is_pie:
            pts = colors or (palette[: len(categories)])
            for pi in range(len(categories)):
                point = s.points[pi]
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = _rgb(pts[pi % len(pts)])
        else:
            try:
                s.format.fill.solid()
                s.format.fill.fore_color.rgb = _rgb(colors[0] if colors else (single or palette[si % len(palette)]))
            except Exception:  # noqa: BLE001  折线图没有填充
                pass
            try:
                s.format.line.color.rgb = _rgb(colors[0] if colors else (single or palette[si % len(palette)]))
            except Exception:  # noqa: BLE001
                pass

    # 坐标轴格式
    try:
        if chart_type not in (XL_CHART_TYPE.PIE, XL_CHART_TYPE.DOUGHNUT):
            va = chart.value_axis
            va.has_major_gridlines = bool(spec.get("gridlines", True))
            va.tick_labels.font.size = Pt(float(spec.get("axis_size", 10)))
            if theme:
                va.tick_labels.font.color.rgb = _rgb(theme.muted)
            ca = chart.category_axis
            ca.tick_labels.font.size = Pt(float(spec.get("axis_size", 10)))
            if theme:
                ca.tick_labels.font.color.rgb = _rgb(theme.muted)
    except Exception:  # noqa: BLE001
        pass

    # 柱间距（仅柱/条）
    try:
        if spec.get("gap_width") is not None:
            plots[0].gap_width = int(spec["gap_width"])
        if spec.get("overlap") is not None:
            plots[0].overlap = int(spec["overlap"])
    except Exception:  # noqa: BLE001
        pass

    return {"name": frame.name, "id": frame.shape_id,
            "type": str(spec.get("type") or "column"),
            "categories": len(categories), "series": len(series_specs),
            "box": [int(frame.left), int(frame.top), int(frame.width), int(frame.height)]}


def chart_info(slide, target) -> dict:
    from .shapes import find_shape
    shp = find_shape(slide, target)
    if not getattr(shp, "has_chart", False):
        raise ValueError(f"{target!r} 不是图表")
    chart = shp.chart
    out = {"name": shp.name, "id": shp.shape_id, "type": str(chart.chart_type),
           "has_title": bool(chart.has_title), "has_legend": bool(chart.has_legend),
           "series": []}
    try:
        cats = list(chart.plots[0].categories)
        out["categories"] = [str(c) for c in cats]
    except Exception:  # noqa: BLE001
        pass
    for s in chart.series:
        try:
            values = list(s.values)
        except Exception:  # noqa: BLE001
            values = []
        out["series"].append({"name": s.name, "values": values})
    return out
