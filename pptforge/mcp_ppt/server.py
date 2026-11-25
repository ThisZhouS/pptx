"""pptforge MCP server：把 PowerPoint OOXML 能力暴露给任意 LLM 客户端。

启动：python -m mcp_ppt.server        （stdio 传输）
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp.server.fastmcp import FastMCP
from pptx import Presentation

from mcp_ppt import comtools
from mcp_ppt.catalog import get_catalog
from mcp_ppt.deck import build_deck as _build_deck
from mcp_ppt.deck import dump_slide_xml as _dump_slide_xml
from mcp_ppt.deck import patch_deck
from mcp_ppt.inspect import inspect_deck, validate_deck
from mcp_ppt.layout import layout_report, png_stats
from mcp_ppt.motion import apply_recipe_to_model, get_recipe
from mcp_ppt.shapes import slide_shapes_info
from mcp_ppt.theme import apply_theme as theme_apply
from mcp_ppt.theme import extract_theme, read_show, set_show as theme_show

mcp = FastMCP(
    "pptforge",
    instructions=(
        "用 pptforge 生成带切换与对象动画的 PowerPoint。典型流程：\n"
        "1) deck_model_template 了解 JSON 模型结构（或直接按说明写 model）；\n"
        "2) catalog_transitions / catalog_animations 挑选效果名；\n"
        "3) build_deck 生成 .pptx（可传 motion_style 套用整套动效风格）；\n"
        "4) validate_deck / layout_check 校验结构与版式；\n"
        "5) render_png 或 export_pdf 出图。\n"
        "只改已有文件时用 set_transition / add_animations / add_elements / edit_elements，不会重建内容。\n"
        "形状层：add_elements 新建元素，edit_elements 做定位/移动/缩放/属性/组合/对齐/复制/删除，\n"
        "set_text 重写富文本，add_table / add_chart / edit_table 处理表格与图表，element_info 读回核对。\\n"
        "动画目标可以指向组（group:<组名>）；图表动画只支持整体（按系列分级在 COM 上不可用）。\\n"
        "属性：fill 支持纯色/渐变/无填充，effects 支持外阴影/发光/柔化边缘，line_* 支持虚线/箭头/透明度。"
    ),
)


def _parse_slides(text, count: int) -> list[int]:
    if text is None or str(text).strip().lower() in ("", "all", "全部", "*"):
        return list(range(1, count + 1))
    if isinstance(text, int):
        return [text]
    if isinstance(text, (list, tuple)):
        return [int(x) for x in text]
    out: list[int] = []
    for part in str(text).replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            out.extend(range(int(a), int(b) + 1))
        else:
            out.append(int(part))
    return [i for i in out if 1 <= i <= count]


# ------------------------------------------------------------------ 目录

@mcp.tool()
def catalog_transitions(query: str = "") -> dict:
    """列出可用的幻灯片切换效果（id / 中文名 / OOXML 元素 / 需要的扩展命名空间）。

    query 可匹配 id、中文名或 COM 枚举名；留空返回全部 150 个。
    """
    cat = get_catalog()
    items = cat.list_transitions(query)
    return {"count": len(items), "effects": items}


@mcp.tool()
def catalog_animations(query: str = "", only_exit: bool = False) -> dict:
    """列出可用的对象动画效果（id / 中文名 / presetID / 类别 entr|emph|exit|path）。"""
    cat = get_catalog()
    items = cat.list_animations(query, only_exit)
    return {"count": len(items), "effects": items}


@mcp.tool()
def motion_styles() -> dict:
    """列出内置动效风格（subtle / bold / cinematic / tech）及其切换与动画配置。"""
    from mcp_ppt.motion import RECIPES

    return {k: {
        "cover": v.cover_transition, "section": v.section_transition,
        "content": v.content_transition,
        "cover_animations": v.cover_animations,
        "content_animations": v.content_animations,
        "card_animations": v.card_animations,
    } for k, v in RECIPES.items()}


# ------------------------------------------------------------------ 模型

@mcp.tool()
def deck_model_template() -> dict:
    """返回 deck JSON 模型的完整说明与可直接改写的样板。"""
    return {
        "说明": {
            "theme": "dark | light | midnight | warm，或 {preset, bg, fg, muted, accent, card, font_title, font_body}",
            "defaults.transition": "整份 deck 的默认切换，页内 transition 覆盖它",
            "slides[].layout": "title | section | bullets | metrics | quote | image | table | blank",
            "slides[].transition": "{effect, duration(秒), speed(slow|med|fast), advance_after(秒), advance_on_click(false 关闭点击换片)}",
            "slides[].animations[]": "{target, effect, trigger(click|with_previous|after_previous), duration(秒), delay(秒), repeat, smooth_start, smooth_end, exit, build(all|paragraph)}",
            "target": "形状语义名：title / subtitle / bullets / kicker / cards 的 card1..cardN / cardValue1.. / accent-bar / image / table / footer，也支持 * 或 shape:<id>；组用 group:<组名>",
            "build=paragraph": "文本框按段落逐条出现（生成 p:txEl/p:pRg 分级动画）",
            "slides[].elements": "自定义元素：{name, kind, x, y, w, h, anchor, text, fill, alpha, rot, flip_h, flip_v, line_color, line_width, line_dash, line_head, line_tail, effects, crop, brightness, contrast, paragraphs, font, size, bold, color, align, z_order, group}；kind=rect|round-rect|ellipse|triangle|diamond|star|pentagon|chevron|arrow|line|text|image",
            "fill": "\"#RRGGBB\" | \"none\" | {type:solid,color,alpha} | {type:gradient,angle,style:linear|radial,stops:[{pos,color,alpha}]}",
            "effects": "{shadow:{blur,dist,dir,color,alpha,inner}, glow:{radius,color,alpha}, soft_edge:pt}",
            "paragraphs": "[{align, level, line_spacing, space_before, space_after, bullet:{char,color,size_pct}|none, runs:[{text,bold,italic,underline,strike,size,color,font,baseline}]}]",
            "crop": "{l,t,r,b} 比例（0..1），配合 brightness/contrast（-1..1）",
            "单位": "x/y/w/h 支持数字（英寸）、\"50%\"、\"1.5in\"、\"120pt\"、\"3cm\"；anchor=left-top（默认）/center/right-top/left-bottom/…",
            "slides[].ops": "对已有形状批量操作：place/move/resize/props/group/ungroup/align/duplicate/delete，按顺序执行",
            "element.group": "同一个元素上写 \"group\": \"组名\" 会自动组合（≥2 个成员）",
        },
        "样板": {
            "theme": "dark",
            "footer": "公司 · 2026",
            "defaults": {"transition": {"effect": "fade", "duration": 0.7}},
            "slides": [
                {"layout": "title", "title": "季度复盘", "subtitle": "Q3 关键结果与下一步",
                 "meta": ["2026-10", "内部"],
                 "transition": {"effect": "morph", "duration": 1.2},
                 "animations": [
                     {"target": "title", "effect": "fade", "trigger": "click", "duration": 0.8},
                     {"target": "subtitle", "effect": "fade", "trigger": "after_previous"}]},
                {"layout": "bullets", "kicker": "进展", "title": "三件事做成了",
                 "bullets": ["留存率提升 12%", {"text": "来自新引导流程", "level": 1}, "成本下降 8%"],
                 "transition": {"effect": "cover-left", "duration": 0.5},
                 "animations": [
                     {"target": "title", "effect": "fade", "trigger": "click"},
                     {"target": "bullets", "effect": "fly", "trigger": "after_previous",
                      "build": "paragraph", "duration": 0.5}]},
                {"layout": "metrics", "title": "关键指标",
                 "cards": [{"value": "12%", "label": "留存提升"},
                           {"value": "8%", "label": "成本下降"},
                           {"value": "2.1x", "label": "转化"}],
                 "animations": [{"target": ["card1", "card2", "card3"], "effect": "zoom",
                                 "trigger": "click", "duration": 0.5}]},
                {"layout": "blank", "background": "#071426",
                 "kicker": "diagram",
                 "elements": [
                     {"name": "n1", "kind": "ellipse", "x": "18%", "y": "45%", "w": 1.3, "h": 1.3,
                      "text": "A", "fill": "#36C2FF", "align": "center", "group": "nodes"},
                     {"name": "n2", "kind": "ellipse", "x": "40%", "y": "45%", "w": 1.3, "h": 1.3,
                      "text": "B", "fill": "#36C2FF", "align": "center", "group": "nodes"},
                     {"name": "wire", "kind": "line", "x": 3.4, "y": 4.3, "w": 1.2, "h": 0,
                      "line_color": "#36C2FF", "line_width": 2}],
                 "ops": [
                     {"op": "move", "target": "kick", "dy": "-8pt"},
                     {"op": "resize", "target": "group:nodes", "scale": 1.1, "anchor": "center"},
                     {"op": "duplicate", "target": "n2", "name": "n3", "dx": 2.6},
                     {"op": "align", "targets": ["n3"], "mode": "top", "to": "n2"}],
                 "animations": [{"target": "group:nodes", "effect": "fade", "trigger": "click"},
                                {"target": "n3", "effect": "spin", "trigger": "after_previous"}]},
            ],
        },
    }


@mcp.tool()
def build_deck(model: dict, out_path: str, motion_style: str = "") -> dict:
    """按 JSON 模型生成 pptx（含切换与对象动画）。

    model: 见 deck_model_template 的结构。
    motion_style: 传入 subtle|bold|cinematic|tech 时，未写动效的页自动套用该风格。
    """
    payload = model
    if motion_style:
        payload = apply_recipe_to_model(model, motion_style)
    report = _build_deck(payload, out_path)
    report["validate"] = validate_deck(out_path)["stats"]
    return report


@mcp.tool()
def build_deck_from_outline(title: str, sections: list[dict], out_path: str,
                            subtitle: str = "", theme: str = "dark",
                            motion_style: str = "subtle", footer: str = "") -> dict:
    """用最简大纲直接出片：sections=[{title, points:[...], kicker?, image?}]。"""
    slides: list[dict] = [{
        "layout": "title", "title": title, "subtitle": subtitle or (sections[0].get("title", "") if sections else ""),
        "meta": [footer] if footer else [],
    }]
    for i, sec in enumerate(sections, start=1):
        points = sec.get("points") or sec.get("bullets") or []
        if sec.get("image"):
            slides.append({"layout": "image", "title": sec["title"], "image": sec["image"],
                           "bullets": points, "kicker": sec.get("kicker")})
        elif sec.get("cards"):
            slides.append({"layout": "metrics", "title": sec["title"], "cards": sec["cards"],
                           "kicker": sec.get("kicker")})
        elif sec.get("quote"):
            slides.append({"layout": "quote", "quote": sec["quote"], "source": sec.get("source"),
                           "title": sec["title"]})
        else:
            slides.append({"layout": "bullets", "title": sec["title"], "bullets": points,
                           "kicker": sec.get("kicker")})
    slides.append({"layout": "section", "title": "谢谢", "subtitle": footer or title})
    model = {"theme": theme, "footer": footer, "slides": slides}
    payload = apply_recipe_to_model(model, motion_style)
    report = _build_deck(payload, out_path)
    report["validate"] = validate_deck(out_path)["stats"]
    return report


# ------------------------------------------------------------------ 改造已有文件

@mcp.tool()
def set_transition(path: str, slides: str = "all", effect: str = "fade",
                   duration: float = 0.0, speed: str = "",
                   advance_after: float = 0.0, advance_on_click: bool = True) -> dict:
    """给已有 pptx 的若干页设置切换。slides 支持 "all" / "1,3" / "2-5"。"""
    deck = inspect_deck(path)
    targets = _parse_slides(slides, deck["slide_count"])
    spec: dict = {"effect": effect}
    if duration:
        spec["duration"] = duration
    if speed:
        spec["speed"] = speed
    if advance_after:
        spec["advance_after"] = advance_after
    if not advance_on_click:
        spec["advance_on_click"] = False
    patches = [{"slide": i, "transition": spec} for i in targets]
    return patch_deck(path, patches)


@mcp.tool()
def add_animations(path: str, slide: int, animations: list[dict]) -> dict:
    """重写某一页的对象动画（覆盖该页原动画，不动内容与其他页）。

    animations 结构同 deck 模型：{target, effect, trigger, duration, delay, build, exit}
    target 可以是组：写 "group:<组名>"。
    """
    return patch_deck(path, [{"slide": slide, "animations": animations}])


# ------------------------------------------------------------------ morph 时间轴

@mcp.tool()
def decompose_deck(path: str) -> dict:
    """拆解每页元素（id / 名称 / 标签 / 位置 / 文本），并列出跨页同编码元素。

    跨页同编码（!!name）的对象就是 morph 能插值的对象——这是做"多页元素变换"动画的前提。
    """
    from mcp_ppt.morph import decompose
    return decompose(path)


@mcp.tool()
def assign_codes(path: str, mapping: dict) -> dict:
    """把元素改名为 !!code，让 PowerPoint 的 morph 强制按名称匹配。

    mapping 形如 {"标题 5": "hero", "shape:12": "dot"}；同一页内的同名元素都会被改。
    """
    from mcp_ppt.morph import assign_codes as _assign
    return _assign(path, mapping)


@mcp.tool()
def plan_morph(model: dict) -> dict:
    """只算关键帧时间轴（不落盘）：每帧起止时间、各元素在那一帧的状态。

    model: {frame_duration, transition, base:{elements:[…]}, elements:[{name, kind, …base, keyframes:[…]}]}
    """
    from mcp_ppt.morph import plan
    return plan(model)


@mcp.tool()
def build_morph(model: dict, out_path: str) -> dict:
    """把关键帧展开成 morph 序列：N 页 = 关键帧数，页间 morph 插值 + advTm 自动换片。

    这就是"用多页元素变换实现复杂动画"的写法：
    同一元素在每页叫 !!name，改变位置/尺寸/旋转/透明度/填充，页间自动插值成动画。
    关键帧里省略的字段沿用上一帧，因此"中间点"可写 0 个、1 个或多个。
    """
    from mcp_ppt.morph import build_morph_deck
    return build_morph_deck(model, out_path)


# ------------------------------------------------------------------ 形状层

@mcp.tool()
def add_elements(path: str, slide: int, elements: list[dict],
                 groups: dict | None = None) -> dict:
    """在指定页新增形状/文本框/图片，并可选地组合。

    elements[]: {name, kind, x, y, w, h, anchor, text, fill, alpha, rot, flip_h, flip_v,
                 line_color, line_width, font, size, bold, color, align, z_order, group, image}
    · kind: rect | round-rect | ellipse | triangle | diamond | star | pentagon | chevron |
            arrow | line | text | image
    · x/y/w/h 支持数字（英寸）、"50%"、"1.5in"、"120pt"、"3cm"；anchor 默认 left-top，
      常用 center / right-top / left-bottom
    · groups: {"组名": ["元素名", ...]} —— 低于 2 个成员会被跳过
    """
    report = patch_deck(path, [{"slide": slide, "elements": elements}])
    if groups:
        ops = [{"op": "group", "targets": members, "name": name}
               for name, members in groups.items() if len(members) >= 2]
        if ops:
            more = patch_deck(path, [{"slide": slide, "ops": ops}])
            applied = report["patched"][0]["applied"]
            applied.setdefault("ops", []).extend(more["patched"][0]["applied"].get("ops", []))
            report["warnings"].extend(more.get("warnings", []))
    return report


@mcp.tool()
def edit_elements(path: str, slide: int, ops: list[dict]) -> dict:
    """对某页已有形状执行批量操作（按顺序执行，单条失败不影响后续）。

    ops[] 支持：
      {"op":"place","target":"card1","x":"10%","y":4,"anchor":"left-top"}
      {"op":"move","targets":["card1","card2"],"dx":0.4,"dy":"-10pt"}
      {"op":"resize","target":"card1","scale":1.1,"anchor":"center"}   或 w/h
      {"op":"props","target":"title","size":40,"color":"FFC53D","alpha":0.8,
       "rot":8,"flip_h":true,"z_order":"front","fill":"18233F"}
      {"op":"group","targets":["a","b"],"name":"row"}
      {"op":"ungroup","target":"row"}
      {"op":"align","targets":["card1","card2","card3"],"mode":"distribute-h","to":"slide"}
         mode: left|right|center-h|top|bottom|center-v|center|distribute-h|distribute-v
      {"op":"duplicate","target":"card1","name":"card4","dx":2.4}
      {"op":"delete","target":"old"}
    """
    return patch_deck(path, [{"slide": slide, "ops": ops}])


@mcp.tool()
def set_theme(path: str, colors: dict | None = None, fonts: dict | None = None,
              name: str = "") -> dict:
    """改文档主题部件（theme1.xml）的配色方案与字体方案。

    colors: {"accent1":"#FF7A45", "accent2":"#6FE3C4", "dk1","lt1","dk2","lt2",
             "accent3".."accent6", "hlink","folHlink"}（只写给出的键）
    fonts:  {"major":"Microsoft YaHei", "minor":"Microsoft YaHei", "hans":"Microsoft YaHei"}
            也接受 major_latin/major_ea/minor_latin/minor_ea 分项
    name:   配色方案名称（clrScheme@name）
    """
    prs = Presentation(path)
    applied = theme_apply(prs, colors, fonts, name or None)
    prs.save(path)
    return {"file": os.path.abspath(path), "applied": applied}


@mcp.tool()
def set_show(path: str, show: dict) -> dict:
    """写放映设置（presProps.xml 的 showPr）。

    show: {loop:true, show_type:"kiosk"|"speaker"|"browse", narration:false,
           animation:true, timings:true, scrollbar:false}
    """
    prs = Presentation(path)
    info = theme_show(prs, show)
    prs.save(path)
    return {"file": os.path.abspath(path), **info, "readback": read_show(path)}


@mcp.tool()
def read_theme(path: str) -> dict:
    """读取文档的主题配色/字体与放映设置（可用于"照着这份模板做"）。"""
    info = extract_theme(path)
    info["show"] = read_show(path)
    return info


@mcp.tool()
def set_text(path: str, slide: int, target: str, paragraphs: list[dict]) -> dict:
    """重写某个文本框的内容与格式（支持段内多 run 混排、项目符号、行距）。

    paragraphs[]: {align, level, line_spacing, space_before, space_after,
                   bullet: {char, color, size_pct} | "none",
                   runs: [{text, bold, italic, underline, strike, size, color, alpha,
                           font, font_ea, baseline, spc}]}
    """
    return patch_deck(path, [{"slide": slide,
                              "ops": [{"op": "props", "target": target,
                                       "paragraphs": paragraphs}]}])


@mcp.tool()
def add_table(path: str, slide: int, table: dict) -> dict:
    """在指定页新增表格。

    table: {name, x, y, w, h, rows:[[...]], header:bool, col_widths:[inch...],
            row_heights:[inch...], font_size, cell:{fill,color,size,align,valign,borders},
            header_cell:{...}, body_cell:{...}, style:{first_row,band_row,style_id}}
    borders 示例：{"bottom": {"color":"#FF7A45","width":1.5,"dash":"solid"}}
    """
    return patch_deck(path, [{"slide": slide, "elements": [{"kind": "table", **table}]}])


@mcp.tool()
def add_chart(path: str, slide: int, chart: dict) -> dict:
    """在指定页新增原生图表（数据内嵌工作簿，可在 PowerPoint 里"编辑数据"）。

    chart: {name, type, categories:[], series:[{name, values, color, colors}],
            title, legend:"bottom"|"right"|"none", data_labels, number_format,
            gridlines, x, y, w, h, gap_width, overlap}
    type: column|column-stacked|bar|line|line-markers|area|pie|doughnut|radar|scatter
    """
    return patch_deck(path, [{"slide": slide, "elements": [{"kind": "chart", **chart}]}])


@mcp.tool()
def edit_table(path: str, slide: int, target: str, ops: list[dict]) -> dict:
    """编辑表格：单元格 / 合并 / 表样式 / 列宽 / 行高。

    ops[] 支持：
      {"op":"table","cells":[{row,col,text,fill,color,size,bold,align,valign,borders}]}
      {"op":"table","merge":{"row1":0,"col1":0,"row2":0,"col2":1}}
      {"op":"table","style":{"first_row":true,"band_row":true,"style_id":"{GUID}"}}
      {"op":"table","widths":[2.6,2.2,2.2],"heights":[0.6,0.5,0.5]}
    也接受直接写 {"op":"cell"|"merge"|"style"|"widths"|"heights"}。
    """
    from mcp_ppt.table import edit_table as _edit_table
    prs = Presentation(path)
    if not 1 <= slide <= len(prs.slides):
        return {"error": f"页号越界: {slide}（共 {len(prs.slides)} 页）"}
    s = prs.slides[slide - 1]
    payload = [{**op, "target": target} for op in ops]
    result = _edit_table(s, payload)
    prs.save(path)
    return {"file": os.path.abspath(path), "results": result}


@mcp.tool()
def element_info(path: str, slide: int = 1, target: str = "", deep: bool = False) -> dict:
    """读取某页形状的几何、文本、旋转/翻转、组内结构（英寸 + EMU）。

    target 留空列出全部；支持 "组名"、"group:组名"、"shape:12"。
    指定表格/图表时会附带内容摘要（表格的单元格文本、图表的数据系列）。
    """
    prs = Presentation(path)
    if not 1 <= slide <= len(prs.slides):
        return {"error": f"页号越界: {slide}（共 {len(prs.slides)} 页）"}
    s = prs.slides[slide - 1]
    items = slide_shapes_info(s, target or None, deep)
    if target:
        from mcp_ppt.shapes import find_shape
        from mcp_ppt.table import table_info
        from mcp_ppt.chart import chart_info

        shp = find_shape(s, target)
        if getattr(shp, "has_table", False):
            items[0]["table"] = table_info(s, target)
        if getattr(shp, "has_chart", False):
            items[0]["chart"] = chart_info(s, target)
    return {"slide": slide, "count": len(items), "shapes": items,
            "slide_size_in": [round(int(prs.slide_width) / 914400, 3),
                              round(int(prs.slide_height) / 914400, 3)]}


# ------------------------------------------------------------------ 检查 / 出图

@mcp.tool()
def inspect(path: str, with_xml: bool = False) -> dict:
    """读取 pptx 结构：每页形状、切换（元素与前缀）、动画（preset 与触发）。"""
    return inspect_deck(path, with_xml)


@mcp.tool()
def validate(path: str) -> dict:
    """结构校验：zip 完整性、动画形状引用、transition 与 timing 的顺序。"""
    return validate_deck(path)


@mcp.tool()
def layout_check(path: str) -> dict:
    """用 PowerPoint 计算文本实际占位，报告溢出与越界（需本机 PowerPoint）。"""
    return layout_report(path)


@mcp.tool()
def slide_xml(path: str, slide: int) -> dict:
    """导出某一页的完整 XML（逆向/排查用）。"""
    return {"slide": slide, "xml": _dump_slide_xml(path, slide)}


@mcp.tool()
def render_png(path: str, out_dir: str = "", width: int = 1600) -> dict:
    """逐页导出 PNG（需本机 PowerPoint），附带像素统计判断页面是否有内容。"""
    out_dir = out_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "render")
    return png_stats(path, out_dir, width)


@mcp.tool()
def export_pdf(path: str, out_path: str = "") -> dict:
    """导出 PDF（需本机 PowerPoint）。"""
    out_path = out_path or os.path.splitext(os.path.abspath(path))[0] + ".pdf"
    return comtools.export_pdf(path, out_path)


@mcp.tool()
def capability_report() -> dict:
    """当前环境能力：效果目录规模、COM/PowerPoint 是否可用。"""
    cat = get_catalog()
    return {
        "transitions": len(cat.transitions),
        "animations": len(cat.animations),
        "powerpoint_com": comtools.available(),
        "python": sys.version.split()[0],
        "cwd": os.getcwd(),
    }


if __name__ == "__main__":
    mcp.run()

