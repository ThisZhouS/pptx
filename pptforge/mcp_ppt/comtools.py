"""本机 PowerPoint COM：用真实 PowerPoint 验证与渲染（可选依赖 win32com）。

用途
----
* verify()  : 打开 pptx，读回切换 EntryEffect 与动画序列，确认 PowerPoint 真正接受我们写的 XML
* export_png(): 逐页导出 PNG
* export_pdf(): 导出 PDF

无 COM 时所有函数返回 {"available": False}，不影响生成功能。
"""
from __future__ import annotations

import os
import sys

try:  # pragma: no cover
    import win32com.client as win32
    _HAVE_COM = True
except Exception:  # noqa: BLE001
    win32 = None
    _HAVE_COM = False

PP_SAVE_AS_PPTX = 24
PP_ALERTS_NONE = 1
MSO_FALSE = 0
MSO_TRUE = -1


def available() -> bool:
    return _HAVE_COM


def _app():
    app = win32.Dispatch("PowerPoint.Application")
    app.DisplayAlerts = PP_ALERTS_NONE
    try:
        app.Visible = True
    except Exception:  # noqa: BLE001
        pass
    return app


def _open(app, path: str):
    return app.Presentations.Open(os.path.abspath(path), ReadOnly=MSO_TRUE, WithWindow=MSO_FALSE)


def verify(path: str) -> dict:
    """用 PowerPoint 读回结构。返回每页的切换 EntryEffect 与动画条目。"""
    if not _HAVE_COM:
        return {"available": False, "reason": "win32com / PowerPoint 不可用"}
    app = _app()
    pres = None
    try:
        pres = _open(app, path)
        slides = []
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            t = slide.SlideShowTransition
            entry = None
            try:
                entry = int(t.EntryEffect)
            except Exception:  # noqa: BLE001
                pass
            effects = []
            try:
                seq = slide.TimeLine.MainSequence
                for j in range(1, seq.Count + 1):
                    eff = seq.Item(j)
                    effects.append({
                        "type": int(eff.EffectType),
                        "exit": bool(eff.Exit),
                        "trigger": int(eff.Timing.TriggerType),
                        "duration": round(float(eff.Timing.Duration), 3),
                        "shape_id": int(eff.Shape.Id) if eff.Shape is not None else None,
                        "shape_name": getattr(eff.Shape, "Name", None),
                    })
            except Exception as exc:  # noqa: BLE001
                effects = [{"error": repr(exc)}]
            slides.append({
                "index": i,
                "entry_effect": entry,
                "slide_shapes": pres.Slides(i).Shapes.Count,
                "animations": effects,
            })
        return {"available": True, "slides": slides, "slide_count": pres.Slides.Count}
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


def export_png(path: str, outdir: str, width: int = 1600, max_pages: int | None = None) -> dict:
    if not _HAVE_COM:
        return {"available": False}
    os.makedirs(outdir, exist_ok=True)
    app = _app()
    pres = None
    try:
        pres = _open(app, path)
        height = int(round(width * pres.PageSetup.SlideHeight / pres.PageSetup.SlideWidth))
        files = []
        count = pres.Slides.Count if max_pages is None else min(max_pages, pres.Slides.Count)
        for i in range(1, count + 1):
            target = os.path.abspath(os.path.join(outdir, f"slide_{i:03d}.png"))
            pres.Slides(i).Export(target, "PNG", width, height)
            files.append(target)
        return {"available": True, "files": files, "size": [width, height]}
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


def export_pdf(path: str, out_path: str) -> dict:
    if not _HAVE_COM:
        return {"available": False}
    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    app = _app()
    pres = None
    try:
        pres = _open(app, path)
        pres.SaveAs(os.path.abspath(out_path), 32)  # ppSaveAsPDF
        return {"available": True, "file": os.path.abspath(out_path)}
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass


if __name__ == "__main__":
    print(verify(sys.argv[1]))
