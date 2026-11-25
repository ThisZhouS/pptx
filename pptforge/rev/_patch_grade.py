"""两处修正：章节页补装饰层；accent 按对比度下限迭代（模板 4.7）。"""
import io

# 1) design.py：章节页也加装饰（不重复页码）
P1 = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P1, encoding="utf-8").read()
old = '''        if spec.get("index"):
            self._text(b, "page-num", SLIDE_W - MARGIN - Inches(0.7), Inches(6.6),
                       Inches(0.7), Inches(0.3),
                       [{"text": f"{int(spec['index']):02d}", "size": 10, "color": t.muted}],
                       align=PP_ALIGN.RIGHT)
        return b'''
new = '''        if spec.get("index"):
            self._text(b, "page-num", SLIDE_W - MARGIN - Inches(0.7), Inches(6.6),
                       Inches(0.7), Inches(0.3),
                       [{"text": f"{int(spec['index']):02d}", "size": 10, "color": t.muted}],
                       align=PP_ALIGN.RIGHT)
        self._decor(b, spec, None)   # 装饰层（页码已在上面单独画）
        return b'''
assert old in t
t = t.replace(old, new, 1)
io.open(P1, "w", encoding="utf-8").write(t)
print("[ok] section 加装饰层")

# 2) build_template_grade.py：accent 按对比度迭代到 >= 4.5
P2 = r"E:\pptx\pptforge\scripts\build_template_grade.py"
t2 = io.open(P2, encoding="utf-8").read()
old2 = '''def accent_from_hue(hue: float = 205, sat: float = 0.65, val: float = 0.77) -> str:
    """模板实测：accent 饱和度中位 0.65、明度 0.77，与浅底对比≈4.7。"""
    r, g, b = colorsys.hsv_to_rgb(hue / 360, sat, val)
    return "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))'''
new2 = '''def _rel(hexv: str) -> float:
    def f(c):
        c /= 255
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (int(hexv[i:i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def accent_from_hue(hue: float = 205, sat: float = 0.72, val: float = 0.80,
                    min_contrast: float = 4.5) -> str:
    """模板实测 accent 与白底对比≈4.7；这里按对比度下限往下压明度，保证可读。"""
    best = None
    for v in (val, 0.75, 0.7, 0.65, 0.6, 0.55, 0.5):
        r, g, b = colorsys.hsv_to_rgb(hue / 360, sat, v)
        hexv = "#%02X%02X%02X" % (round(r * 255), round(g * 255), round(b * 255))
        c = (1.05) / (_rel(hexv) + 0.05)
        best = hexv
        if c >= min_contrast:
            return hexv
    return best'''
assert old2 in t2
t2 = t2.replace(old2, new2, 1)
t2 = t2.replace('ACCENT2 = accent_from_hue(168, 0.55, 0.82)   # 辅色：青绿，同明度区间',
                'ACCENT2 = accent_from_hue(168, 0.6, 0.80)   # 辅色：青绿，同样按对比度下限')
io.open(P2, "w", encoding="utf-8").write(t2)
print("[ok] accent 对比度迭代")