"""布局体检：用 PowerPoint 计算文本实际占位，判断溢出与越界。

PowerPoint 的 TextRange.BoundHeight/BoundWidth 是排版后的真实尺寸，
比任何静态估算都准，因此这一层用 COM 做"自动视觉检查"。
"""
from __future__ import annotations

import os

from .comtools import _HAVE_COM, _app, _open, win32

PP_ALERTS_NONE = 1


def layout_report(path: str, tolerance: float = 0.06) -> dict:
    if not _HAVE_COM:
        return {"available": False, "reason": "win32com / PowerPoint 不可用"}
    app = _app()
    pres = None
    try:
        pres = _open(app, path)
        sw = float(pres.PageSetup.SlideWidth)
        sh = float(pres.PageSetup.SlideHeight)
        pages = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            issues = []
            shapes = []
            for j in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(j)
                try:
                    left, top = float(shp.Left), float(shp.Top)
                    width, height = float(shp.Width), float(shp.Height)
                except Exception:  # noqa: BLE001
                    continue
                rec = {"name": shp.Name, "left": round(left), "top": round(top),
                       "width": round(width), "height": round(height)}
                if left < -1 or top < -1 or left + width > sw + 1 or top + height > sh + 1:
                    issues.append(f"{shp.Name}: 超出页面边界")
                if getattr(shp, "HasTextFrame", 0):
                    tf = shp.TextFrame
                    try:
                        tr = tf.TextRange
                        if tr.Length > 0:
                            bh = float(tr.BoundHeight)
                            bw = float(tr.BoundWidth)
                            rec["text_height"] = round(bh)
                            rec["text_width"] = round(bw)
                            if bh > height * (1.0 + tolerance) + 2:
                                issues.append(
                                    f"{shp.Name}: 文本可能溢出（文本高 {round(bh)}pt > 框高 {round(height)}pt）")
                            if bw > width * (1.0 + tolerance) + 2 and tr.Length < 200:
                                issues.append(
                                    f"{shp.Name}: 文本宽度超出（{round(bw)}pt > {round(width)}pt）")
                    except Exception:  # noqa: BLE001
                        pass
                shapes.append(rec)
            pages.append({"index": i, "shapes": shapes, "issues": issues})
        return {"available": True, "slide_size": [sw, sh], "pages": pages,
                "issue_count": sum(len(p["issues"]) for p in pages)}
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


def png_stats(path: str, outdir: str, width: int = 1280) -> dict:
    """导出 PNG 并做像素统计（无法看图时用于确认页面非空白/有内容）。"""
    from .comtools import export_png

    res = export_png(path, outdir, width)
    if not res.get("available"):
        return res
    try:
        from PIL import Image
    except Exception:  # noqa: BLE001
        return {"available": True, "files": res["files"], "pixel_check": "Pillow 不可用"}
    rows = []
    for f in res["files"]:
        img = Image.open(f).convert("RGB")
        small = img.resize((160, 90))
        px = list(small.getdata())
        uniq = len(set(px))
        # 与最常见的背景色不同的像素比例
        from collections import Counter
        common = Counter(px).most_common(1)[0][1]
        ink_ratio = 1 - common / len(px)
        rows.append({"file": os.path.basename(f), "colors": uniq,
                     "ink_ratio": round(ink_ratio, 4)})
    return {"available": True, "files": res["files"], "pixel_check": rows}