"""逆向第 7 层：主题（theme1.xml）与放映设置（showPr）。

采样：
  · python-pptx 默认模板的 theme1.xml（配色方案 / 字体方案 / 格式方案）
  · 用 PowerPoint 改放映设置（循环/旁白/计时/展台）后 presentation.xml 的 <p:showPr>
  · 用 PowerPoint 改主题配色后 theme1.xml 的变化范围
  · slideMaster / slideLayout 的部件结构

产物: out/gt_theme.pptx, out/theme_raw/, out/theme_map.json
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import win32com.client as win32
from pptx import Presentation
from pptx.util import Inches

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
SEED = os.path.join(OUT, "theme_seed.pptx")
PPTX = os.path.join(OUT, "gt_theme.pptx")
RAW = os.path.join(OUT, "theme_raw")

ppSaveAsOpenXMLPresentation = 24
ppAlertsNone = 1
msoTrue = -1
ppSlideShowUseSlideTimings = 2


def seed() -> None:
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]
    s1 = prs.slides.add_slide(blank)
    box = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    box.text_frame.text = "主题采样页"
    s2 = prs.slides.add_slide(blank)
    s2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1)).text_frame.text = "第二页"
    prs.save(SEED)
    print(f"[seed] {SEED}")


def main() -> int:
    os.makedirs(OUT, exist_ok=True)
    if not os.path.exists(SEED):
        seed()
    import zipfile
    import re

    src = zipfile.ZipFile(SEED)
    parts = [n for n in src.namelist() if "theme" in n or "slideMaster" in n or "slideLayout" in n]
    theme_name = next(n for n in src.namelist() if n.startswith("ppt/theme/theme") and n.endswith(".xml"))
    theme_xml = src.read(theme_name).decode("utf-8")
    print("[parts]", parts[:8])
    print("[theme len]", len(theme_xml))

    # 默认主题的配色与字体
    colors = {}
    for m in re.finditer(r'<a:(dk1|lt1|dk2|lt2|accent[1-6]|hlink|folHlink)>\s*<a:(srgbClr|sysClr)[^>]*?/?>', theme_xml):
        tag = m.group(1)
        seg = theme_xml[m.start():m.start() + 200]
        val = re.search(r'val="([0-9A-Fa-f]{6})"', seg)
        last = re.search(r'lastClr="([0-9A-Fa-f]{6})"', seg)
        colors[tag] = (val.group(1) if val else None) or (last.group(1) if last else None)
    fonts = {k: re.search(rf'<a:{k}>.*?<a:latin typeface="([^"]*)"', theme_xml, re.S).group(1)
             for k in ("majorFont", "minorFont")
             if re.search(rf'<a:{k}>.*?<a:latin typeface="([^"]*)"', theme_xml, re.S)}
    ea_fonts = {k: re.search(rf'<a:{k}>.*?<a:ea typeface="([^"]*)"', theme_xml, re.S).group(1)
                for k in ("majorFont", "minorFont")
                if re.search(rf'<a:{k}>.*?<a:ea typeface="([^"]*)"', theme_xml, re.S)}
    print("[colors]", colors)
    print("[fonts]", fonts, ea_fonts)

    # ---- PowerPoint：改放映设置 + 改主题色
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True
    app.DisplayAlerts = ppAlertsNone
    pres = app.Presentations.Open(os.path.abspath(SEED), ReadOnly=0, WithWindow=0)
    info = {}
    try:
        ss = pres.SlideShowSettings
        info["settings_dir"] = [n for n in dir(ss) if not n.startswith("_")][:24]
        ss.LoopUntilStopped = msoTrue
        ss.ShowWithNarration = 0
        ss.ShowWithAnimation = msoTrue
        ss.AdvanceMode = ppSlideShowUseSlideTimings
        info["loop"] = int(ss.LoopUntilStopped)
        info["narration"] = int(ss.ShowWithNarration)
        info["animation"] = int(ss.ShowWithAnimation)
        info["advance_mode"] = int(ss.AdvanceMode)
        info["show_type"] = int(ss.ShowType)
    except Exception as exc:  # noqa: BLE001
        info["settings_err"] = repr(exc)
    # 主题配色
    try:
        theme = pres.SlideMaster.Theme
        scheme = theme.ThemeColorScheme
        info["colorscheme_dir"] = [n for n in dir(scheme) if not n.startswith("_")]
        info["colors_before"] = {}
        for idx in range(1, 13):
            try:
                info["colors_before"][idx] = int(scheme.Colors(idx).RGB)
            except Exception:  # noqa: BLE001
                pass
        scheme.Colors(6).RGB = 0xFF7A45      # ppAccent1
        scheme.Colors(7).RGB = 0x6FE3C4      # ppAccent2
        info["color_patched"] = True
    except Exception as exc:  # noqa: BLE001
        info["theme_err"] = repr(exc)
    pres.SaveAs(PPTX, ppSaveAsOpenXMLPresentation)
    pres.Close()
    app.Quit()

    os.makedirs(RAW, exist_ok=True)
    z = zipfile.ZipFile(PPTX)
    pres_xml = z.read("ppt/presentation.xml").decode("utf-8")
    show_m = re.search(r"<p:showPr[^>]*>", pres_xml)
    idx = pres_xml.find("showPr")
    print("\n[showPr]", (pres_xml[max(0, idx-120):idx+300] if idx >= 0 else "<none>"))
    print("[pres len]", len(pres_xml))
    theme2 = z.read(theme_name).decode("utf-8")
    changed = []
    for m in re.finditer(r'<a:(accent[1-6]|dk1|lt1|dk2|lt2)>\s*<a:srgbClr val="([0-9A-Fa-f]{6})"', theme2):
        changed.append((m.group(1), m.group(2)))
    print("[theme after]", changed)
    print("[master]", z.read("ppt/slideMasters/slideMaster1.xml").decode("utf-8")[:400])

    out = {
        "theme_part": theme_name,
        "default_colors": colors,
        "default_fonts": {**fonts, **{f"ea_{k}": v for k, v in ea_fonts.items()}},
        "showPr": show_m.group(0) if show_m else None,
        "com": info,
        "theme_after_colors": changed,
        "parts": [n for n in z.namelist() if "theme" in n or "slideMaster" in n or "slideLayout" in n][:12],
    }
    with open(os.path.join(OUT, "theme_map.json"), "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    with open(os.path.join(RAW, "theme1.xml"), "w", encoding="utf-8") as fh:
        os.makedirs(RAW, exist_ok=True)
        fh.write(theme2)
    print(f"[written] {os.path.join(OUT, 'theme_map.json')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
