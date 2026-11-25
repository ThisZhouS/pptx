"""版式与设计系统：把一份 slide spec 画成可编辑的 PowerPoint 形状。

设计目标
--------
* 只用原生 shape（textbox / rect / picture / table），保证 AI 之后仍可编辑
* 每个形状带语义名（title / bullets / card1 ...），动画层用名字定位形状
* 深色/浅色两套主题，字体、字号、栅格固定，避免溢出
"""
from __future__ import annotations

from dataclasses import dataclass, field

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Emu, Inches, Pt
from lxml import etree

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
MARGIN = Inches(0.62)
CONTENT_W = SLIDE_W - 2 * MARGIN


def rgb(value: str) -> RGBColor:
    v = value.lstrip("#")
    if len(v) == 3:
        v = "".join(c * 2 for c in v)
    return RGBColor.from_string(v.upper())


def _set_eastasian(run, typeface: str) -> None:
    """给 run 补 <a:ea>/<a:cs> 字体，保证中文走同一字体。"""
    rpr = run.font._rPr
    if rpr is None:
        return
    from pptx.oxml.ns import qn

    for tag, before in (("a:ea", ("a:cs", "a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst")),
                        ("a:cs", ("a:sym", "a:hlinkClick", "a:hlinkMouseOver", "a:rtl", "a:extLst"))):
        el = rpr.find(qn(tag))
        if el is None:
            el = rpr.makeelement(qn(tag), {})
            rpr.insert_element_before(el, *before)
        el.set("typeface", typeface)


@dataclass
class Theme:
    name: str = "dark"
    bg: str = "0B1020"
    bg_alt: str = "121A31"
    card: str = "18233F"
    fg: str = "F2F6FF"
    muted: str = "9BABCE"
    accent: str = "4C8DFF"
    accent_2: str = "6FE3C4"
    font_title: str = "Microsoft YaHei"
    font_body: str = "Microsoft YaHei"

    @classmethod
    def preset(cls, name: str | None) -> "Theme":
        name = (name or "dark").lower()
        if name in ("light", "亮", "浅色"):
            return cls(name="light", bg="FFFFFF", bg_alt="F4F6FB", card="EEF2FA",
                       fg="101828", muted="5A6683", accent="2F6FED", accent_2="12B886")
        if name in ("midnight", "蓝", "深海"):
            return cls(name="midnight", bg="071426", bg_alt="0D2038", card="122C4A",
                       fg="EAF3FF", muted="8FB0D6", accent="36C2FF", accent_2="FFD166")
        if name in ("warm", "暖", "橙"):
            return cls(name="warm", bg="1A1110", bg_alt="261917", card="33211E",
                       fg="FFF3EC", muted="D8B4A6", accent="FF7A45", accent_2="FFC53D")
        return cls(name="dark")


@dataclass
class ShapeMeta:
    name: str
    shape_id: int
    kind: str = "shape"
    paragraphs: int = 1


@dataclass
class SlideBuild:
    slide: object
    shapes: dict[str, ShapeMeta] = field(default_factory=dict)

    def add(self, name: str, shape, kind: str = "shape", paragraphs: int = 1) -> object:
        try:
            shape.name = name
        except Exception:
            pass
        self.shapes[name] = ShapeMeta(name, shape.shape_id, kind, paragraphs)
        return shape


class Painter:
    """把语义 spec 画到 slide 上。"""

    def __init__(self, prs, theme: Theme):
        self.prs = prs
        self.theme = theme

    # -------------------------------------------------- 基础件
    def _slide(self):
        return self.prs.slides.add_slide(self.prs.slide_layouts[6])  # blank

    def _paint_bg(self, slide, color: str | None = None) -> None:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = rgb(color or self.theme.bg)

    def _rect(self, build: SlideBuild, name: str, x, y, w, h, color: str,
              line: str | None = None, shape=MSO_SHAPE.RECTANGLE):
        shp = build.slide.shapes.add_shape(shape, x, y, w, h)
        shp.fill.solid()
        shp.fill.fore_color.rgb = rgb(color)
        if line:
            shp.line.color.rgb = rgb(line)
            shp.line.width = Pt(1)
        else:
            shp.line.fill.background()
        shp.shadow.inherit = False
        if shp.has_text_frame:
            shp.text_frame.text = ""
        return build.add(name, shp)

    def _text(self, build: SlideBuild, name: str, x, y, w, h,
              paragraphs: list[dict], *, size=24, color=None, bold=False,
              font=None, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
              line_spacing=1.25, space_after=10, italic=False) -> object:
        box = build.slide.shapes.add_textbox(x, y, w, h)
        tf = box.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = 0
        tf.margin_right = 0
        tf.margin_top = 0
        tf.margin_bottom = 0
        font = font or self.theme.font_body
        first = True
        for para in paragraphs:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = para.get("align", align)
            p.line_spacing = para.get("line_spacing", line_spacing)
            p.space_after = Pt(para.get("space_after", space_after))
            level = para.get("level", 0)
            p.level = level
            text = para.get("text", "")
            run = p.add_run()
            run.text = text
            f = run.font
            f.size = Pt(para.get("size", size if level == 0 else max(size - 6, 12)))
            f.bold = para.get("bold", bold)
            f.italic = para.get("italic", italic)
            f.name = para.get("font", font)
            f.color.rgb = rgb(para.get("color", color or self.theme.fg))
            _set_eastasian(run, para.get("font", font))
            if para.get("bulleted"):
                self._bullet(p, para.get("bullet_color", self.theme.accent))
        return build.add(name, box, kind="text", paragraphs=len(paragraphs))

    @staticmethod
    def _bullet(paragraph, color: str) -> None:
        ppr = paragraph._pPr if paragraph._pPr is not None else paragraph._p.get_or_add_pPr()
        ns = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        for tag in ("buNone", "buChar", "buAutoNum", "buClr", "buFont", "buSzPct"):
            for el in ppr.findall(ns + tag):
                ppr.remove(el)
        bu_clr = etree.SubElement(ppr, ns + "buClr")
        srgb = etree.SubElement(bu_clr, ns + "srgbClr")
        srgb.set("val", color.lstrip("#").upper())
        bu_font = etree.SubElement(ppr, ns + "buFont")
        bu_font.set("typeface", "Arial")
        bu_char = etree.SubElement(ppr, ns + "buChar")
        bu_char.set("char", "•")
        # 缩进
        ppr.set("marL", str(int(Inches(0.34))))
        ppr.set("indent", str(int(-Inches(0.34))))

    # -------------------------------------------------- 版式
    def title(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        t = self.theme
        self._rect(b, "accent-bar", MARGIN, Inches(2.05), Inches(0.12), Inches(1.5), t.accent)
        title_size = spec.get("title_size", 54)
        self._text(b, "title", MARGIN + Inches(0.32), Inches(2.0), CONTENT_W - Inches(0.4), Inches(1.6),
                   [{"text": spec.get("title", ""), "size": title_size, "bold": True,
                     "color": spec.get("title_color", t.fg), "line_spacing": 1.1}],
                   font=t.font_title, anchor=MSO_ANCHOR.TOP)
        if spec.get("subtitle"):
            self._text(b, "subtitle", MARGIN + Inches(0.32), Inches(3.85), CONTENT_W - Inches(0.4), Inches(1.1),
                       [{"text": spec["subtitle"], "size": 22, "color": t.muted}], font=t.font_body)
        meta = spec.get("meta") or []
        if meta:
            self._text(b, "meta", MARGIN + Inches(0.32), Inches(6.25), CONTENT_W - Inches(0.4), Inches(0.5),
                       [{"text": "   ·   ".join(str(m) for m in meta), "size": 14, "color": t.muted}])
        self._rule(b, Inches(6.15))
        return b

    def _rule(self, b: SlideBuild, y) -> None:
        self._rect(b, "rule", MARGIN, y, CONTENT_W, Emu(int(Pt(1))), self.theme.muted)

    def section(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background") or self.theme.bg_alt)
        b = SlideBuild(slide)
        t = self.theme
        num = spec.get("index")
        if num is not None:
            self._text(b, "index", MARGIN, Inches(1.32), Inches(3), Inches(1.92),
                       [{"text": f"{int(num):02d}", "size": 108, "bold": True, "color": t.accent}],
                       font=t.font_title)
        self._text(b, "title", MARGIN, Inches(3.32), CONTENT_W, Inches(1.4),
                   [{"text": spec.get("title", ""), "size": 52, "bold": True, "color": t.fg}],
                   font=t.font_title)
        if spec.get("subtitle"):
            self._text(b, "subtitle", MARGIN, Inches(4.72), CONTENT_W, Inches(1.0),
                       [{"text": spec["subtitle"], "size": 20, "color": t.muted}])
        if spec.get("index"):
            self._text(b, "page-num", SLIDE_W - MARGIN - Inches(0.7), Inches(6.6),
                       Inches(0.7), Inches(0.3),
                       [{"text": f"{int(spec['index']):02d}", "size": 10, "color": t.muted}],
                       align=PP_ALIGN.RIGHT)
        self._decor(b, spec, None)   # 装饰层（页码已在上面单独画）
        return b

    def bullets(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        t = self.theme
        self._header(b, spec)
        items = spec.get("bullets") or []
        paras = []
        for item in items:
            if isinstance(item, str):
                paras.append({"text": item, "size": spec.get("body_size", 18),
                              "bulleted": True, "space_after": 16, "color": t.fg})
            else:
                paras.append({
                    "text": item.get("text", ""),
                    "size": item.get("size", spec.get("body_size", 18) - 3 * int(item.get("level", 0))),
                    "level": int(item.get("level", 0)),
                    "bulleted": item.get("level", 0) == 0 or item.get("bulleted", True),
                    "space_after": item.get("space_after", 14),
                    "color": item.get("color", t.fg if item.get("level", 0) == 0 else t.muted),
                })
        top = Inches(2.15)
        body_h = Inches(4.35)
        image = spec.get("image")
        columns = int(spec.get("columns") or 1)
        if image:
            # 左 40% 图 + 右 60% 文（模板里"左侧图"占 27%，是最常见的图文关系）
            img_w = Emu(int(CONTENT_W * 0.38))
            img_h = body_h
            try:
                pic = b.slide.shapes.add_picture(image, MARGIN, top, img_w, img_h)
                b.add("image", pic, kind="image")
            except Exception as exc:  # noqa: BLE001
                b.shapes["__image_error__"] = ShapeMeta(f"图片加载失败: {exc}", -1, "error")
            body_x = MARGIN + img_w + Inches(0.55)
            body_w = CONTENT_W - img_w - Inches(0.55)
            self._text(b, "bullets", body_x, top, body_w, body_h, paras,
                       size=spec.get("body_size", 18), space_after=16)
        elif columns >= 2:
            # 两列：按"要点 + 说明"成对切分，避免把一对拆到两列
            pairs, cur = [], []
            for p in paras:
                if p.get("level", 0) == 0 and cur:
                    pairs.append(cur)
                    cur = [p]
                else:
                    cur.append(p)
            if cur:
                pairs.append(cur)
            half = (len(pairs) + 1) // 2
            gap = Inches(0.5)
            col_w = Emu(int((CONTENT_W - gap) / 2))
            for i, chunk in enumerate((pairs[:half], pairs[half:])):
                flat = [p for group in chunk for p in group]
                if not flat:
                    continue
                x = MARGIN if i == 0 else MARGIN + col_w + gap
                self._text(b, f"bullets-{i + 1}", x, top, col_w, body_h, flat,
                           size=spec.get("body_size", 18), space_after=16)
        else:
            self._text(b, "bullets", MARGIN, top, CONTENT_W, body_h, paras,
                       size=spec.get("body_size", 18), space_after=16)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b

    def metrics(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        t = self.theme
        self._header(b, spec)
        cards = (spec.get("cards") or [])[:4]
        n = max(len(cards), 1)
        gap = Inches(0.32)
        width = Emu(int((CONTENT_W - gap * (n - 1)) / n))
        top = Inches(2.5)
        height = Inches(2.9)
        for i, card in enumerate(cards):
            x = MARGIN + Emu(int((width + gap) * i))
            self._rect(b, f"card{i + 1}", x, top, width, height, t.card)
            self._rect(b, f"cardBar{i + 1}", x, top, width, Emu(int(Pt(4))), t.accent)
            self._text(b, f"cardValue{i + 1}", x + Inches(0.36), top + Inches(0.5),
                       width - Inches(0.7), Inches(1.0),
                       [{"text": str(card.get("value", "")), "size": 40, "bold": True, "color": t.accent}],
                       font=t.font_title)
            self._text(b, f"cardLabel{i + 1}", x + Inches(0.36), top + Inches(1.55),
                       width - Inches(0.7), Inches(1.1),
                       [{"text": str(card.get("label", "")), "size": 16, "color": t.muted}])
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b

    def quote(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background") or self.theme.bg_alt)
        b = SlideBuild(slide)
        t = self.theme
        self._text(b, "mark", MARGIN, Inches(1.2), Inches(2), Inches(1.6),
                   [{"text": "\u201c", "size": 72, "bold": True, "color": t.accent}], font="Georgia")
        self._text(b, "quote", MARGIN + Inches(0.9), Inches(2.35), CONTENT_W - Inches(1.8), Inches(2.6),
                   [{"text": spec.get("quote") or spec.get("title", ""), "size": 30,
                     "italic": True, "color": t.fg, "line_spacing": 1.35}],
                   font=t.font_title)
        if spec.get("subtitle") or spec.get("source"):
            self._text(b, "source", MARGIN + Inches(0.9), Inches(5.4), CONTENT_W - Inches(1.8), Inches(0.6),
                       [{"text": f"— {spec.get('source') or spec.get('subtitle')}",
                         "size": 18, "color": t.muted}])
        return b

    def image(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        t = self.theme
        self._header(b, spec)
        path = spec.get("image")
        img_w = Inches(6.1)
        img_h = Inches(4.3)
        top = Inches(2.1)
        if path:
            try:
                pic = slide.shapes.add_picture(path, MARGIN, top, img_w, img_h)
                b.add("image", pic, kind="image")
            except Exception as exc:  # noqa: BLE001
                b.shapes["__image_error__"] = ShapeMeta(f"图片加载失败: {exc}", -1, "error")
        else:
            self._rect(b, "image", MARGIN, top, img_w, img_h, t.card)
        items = spec.get("bullets") or []
        paras = [{"text": it if isinstance(it, str) else it.get("text", ""),
                  "size": 18, "bulleted": True, "space_after": 12} for it in items]
        if paras:
            self._text(b, "bullets", MARGIN + img_w + Inches(0.5), top, CONTENT_W - img_w - Inches(0.5),
                       img_h, paras, size=18, space_after=12)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b

    def table(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        t = self.theme
        self._header(b, spec)
        from .table import build_table
        from .shapes import find_shape

        tspec = dict(spec)
        tspec.setdefault("name", "table")
        tspec.setdefault("x", MARGIN)
        tspec.setdefault("y", Inches(2.15))
        tspec.setdefault("w", CONTENT_W)
        tspec.setdefault("h", Inches(0.6) * max(len(spec.get("rows") or []), 1))
        tspec.setdefault("header", True)
        tspec.setdefault("cell", {})
        tspec["cell"] = {
            "size": spec.get("font_size", 14),
            "color": spec.get("fg", t.fg),
            "fill": spec.get("cell_fill", t.card),
            "valign": "middle",
            **tspec["cell"],
        }
        tspec.setdefault("header_cell", {})
        tspec["header_cell"] = {"bold": True, "fill": spec.get("header_fill", t.accent),
                                "color": spec.get("header_color", "FFFFFF"),
                                **tspec["header_cell"]}
        tspec.setdefault("body_cell", {})
        tspec["body_cell"] = {"fill": spec.get("cell_fill", t.card), **tspec["body_cell"]}
        info = build_table(slide.shapes, tspec, t, int(self.prs.slide_width), int(self.prs.slide_height))
        shp = find_shape(slide, f"shape:{info['id']}")
        b.add("table", shp, kind="table")
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b

    def chart(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        b = SlideBuild(slide)
        self._header(b, spec)
        from .chart import build_chart
        from .shapes import find_shape

        cspec = dict(spec.get("chart") or {})
        cspec.setdefault("name", "chart")
        cspec.setdefault("x", MARGIN)
        cspec.setdefault("y", Inches(2.05))
        cspec.setdefault("w", CONTENT_W)
        cspec.setdefault("h", Inches(4.4))
        cspec.setdefault("title", spec.get("chart_title", False))
        info = build_chart(slide.shapes, cspec, self.theme,
                           int(self.prs.slide_width), int(self.prs.slide_height))
        shp = find_shape(slide, f"shape:{info['id']}")
        b.add("chart", shp, kind="chart")
        self._footer(b, spec)
        return b

    def blank(self, spec: dict) -> SlideBuild:
        slide = self._slide()
        self._paint_bg(slide, spec.get("background"))
        return SlideBuild(slide)

    # -------------------------------------------------- 公共装饰
    def _header(self, b: SlideBuild, spec: dict) -> None:
        t = self.theme
        if spec.get("kicker"):
            self._text(b, "kicker", MARGIN, Inches(0.34), CONTENT_W, Inches(0.4),
                       [{"text": str(spec["kicker"]).upper(), "size": 13, "bold": True,
                         "color": t.accent}])
            title_top = Inches(0.78)
        else:
            title_top = Inches(0.46)
        self._text(b, "title", MARGIN, title_top, CONTENT_W, Inches(0.95),
                   [{"text": spec.get("title", ""), "size": spec.get("title_size", 40),
                     "bold": True, "color": t.fg}], font=t.font_title)
        self._rect(b, "title-rule", MARGIN, title_top + Inches(0.95), Inches(1.5),
                   Emu(int(Pt(3))), t.accent)

    def _footer(self, b: SlideBuild, spec: dict) -> None:
        if spec.get("footer") is False:
            return
        t = self.theme
        text = spec.get("footer") or ""
        if text:
            self._text(b, "footer", MARGIN, Inches(6.72), CONTENT_W - Inches(1.2), Inches(0.4),
                       [{"text": str(text), "size": 10, "color": t.muted}])

    def _decor(self, b: SlideBuild, spec: dict, index=None) -> None:
        """装饰层：模板实测每页约 7 个装饰（6–7 个小色块 + 5–6 个圆点 + 细线）。

        这里给 5 个（3 色点 + 页码 + 细分隔线），既补上形状密度，又不过度装饰。
        """
        if spec.get("decor") is False:
            return
        t = self.theme
        # 右上角三个点（第一个用主色，其余弱化）
        x0 = SLIDE_W - MARGIN - Inches(0.76)
        for i in range(3):
            color = t.accent if i == 0 else (t.accent_2 if i == 1 else t.muted)
            self._rect(b, f"decor-dot{i + 1}", x0 + Inches(0.26) * i, Inches(0.40),
                       Inches(0.085), Inches(0.085), color, shape=MSO_SHAPE.OVAL)
        # 三个小色块（模板里 chips 占装饰大头）
        for i in range(3):
            self._rect(b, f"decor-chip{i + 1}", x0 + Inches(0.26) * i, Inches(0.60),
                       Inches(0.17), Inches(0.045), t.card)
        # 底部分隔线（比正文窄，制造第三条对齐线）
        self._rect(b, "footer-rule", MARGIN, Inches(6.62), CONTENT_W - Inches(0.9),
                   Emu(int(Pt(0.75))), t.card)
        # 页码
        if index:
            self._text(b, "page-num", SLIDE_W - MARGIN - Inches(0.7), Inches(6.66),
                       Inches(0.7), Inches(0.3),
                       [{"text": f"{int(index):02d}", "size": 10, "color": t.muted}],
                       align=PP_ALIGN.RIGHT)


LAYOUTS = {
    "title": "title", "封面": "title",
    "section": "section", "分节": "section",
    "bullets": "bullets", "列表": "bullets", "content": "bullets", "内容": "bullets",
    "metrics": "metrics", "指标": "metrics", "cards": "metrics",
    "quote": "quote", "引用": "quote",
    "image": "image", "图文": "image", "picture": "image",
    "table": "table", "表格": "table",
    "chart": "chart", "图表": "chart",
    "blank": "blank", "空白": "blank",
}


def resolve_layout(name: str | None) -> str:
    if not name:
        return "bullets"
    return LAYOUTS.get(str(name).strip().lower(), LAYOUTS.get(str(name).strip(), "bullets"))


