"""主题与放映设置验收：theme1.xml 配色/字体 + presProps.xml showPr + 真机回读。"""
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools
from mcp_ppt.comtools import _app, _open
from mcp_ppt.deck import build_deck
from mcp_ppt.inspect import validate_deck
from mcp_ppt.layout import layout_report
from mcp_ppt.theme import extract_theme, read_show

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, "samples")
TARGET = os.path.join(OUT, "theme-demo.pptx")
TEMPLATE = os.path.join(ROOT, "rev", "out", "gt_theme.pptx")

MODEL = {
    "theme": {
        "preset": "dark",
        "name": "Corporate",
        "colors": {"accent1": "#FF7A45", "accent2": "#6FE3C4",
                   "dk2": "#1B2A4A", "lt2": "#EAF3FF", "hlink": "#36C2FF"},
        "fonts": {"major": "Microsoft YaHei", "minor": "Microsoft YaHei", "hans": "Microsoft YaHei"},
    },
    "show": {"loop": True, "show_type": "kiosk", "narration": False, "timings": True},
    "footer": "theme demo",
    "slides": [
        {"layout": "title", "title": "主题与放映", "subtitle": "写入 theme1.xml 与 presProps.xml",
         "transition": {"effect": "morph", "duration": 1.0}},
        {"layout": "bullets", "kicker": "结构", "title": "两个部件",
         "bullets": ["theme1.xml 管配色与字体方案",
                     "presProps.xml 管放映设置（showPr）",
                     {"text": "presentation.xml 里没有 showPr", "level": 1}],
         "animations": [{"target": "bullets", "effect": "wipe", "trigger": "click",
                         "build": "paragraph"}]},
    ],
}

report = build_deck(MODEL, TARGET)
print("[build]", report["slide_count"], "warnings:", report["warnings"])
print("[theme_applied]", json.dumps(report.get("theme_applied"), ensure_ascii=False))
print("[show]", json.dumps(report.get("show"), ensure_ascii=False))
print("[validate]", validate_deck(TARGET)["stats"], validate_deck(TARGET)["problems"])
print("[layout]", layout_report(TARGET).get("issue_count"))
print("[read_show]", json.dumps(read_show(TARGET), ensure_ascii=False))
print("[extract_theme(template)]",
      json.dumps({k: v for k, v in extract_theme(TEMPLATE).items() if k != "source"},
                 ensure_ascii=False)[:400])

# ---- 真机回读
if comtools.available():
    app = _app()
    pres = _open(app, TARGET)
    ss = pres.SlideShowSettings
    print("[com show] loop=%s narration=%s animation=%s advance=%s type=%s" % (
        ss.LoopUntilStopped, ss.ShowWithNarration, ss.ShowWithAnimation,
        ss.AdvanceMode, ss.ShowType))
    scheme = pres.SlideMaster.Theme.ThemeColorScheme
    got = {}
    for idx in (5, 6, 3, 4, 11):
        try:
            got[idx] = "%06X" % int(scheme.Colors(idx).RGB)
        except Exception:  # noqa: BLE001
            pass
    print("[com colors] (5=accent1 6=accent2 3=dk2 4=lt2 11=hlink)", got)
    pres.Close()
    app.Quit()

# ---- 贴合模板：从 gt_theme.pptx 提取主题应用到新 deck
MODEL2 = dict(MODEL)
MODEL2.pop("theme")
MODEL2["theme_from"] = TEMPLATE
TARGET2 = os.path.join(OUT, "theme-from-template.pptx")
rep2 = build_deck(MODEL2, TARGET2)
print("[theme_from]", json.dumps({k: rep2.get(k) for k in ("theme_from", "theme")},
                                 ensure_ascii=False))

# ---- MCP 工具层
from mcp_ppt import server as S

TARGET3 = os.path.join(OUT, "theme-tools.pptx")
S.build_deck({"theme": "dark", "slides": [{"layout": "title", "title": "工具测试"}]}, TARGET3)
print("[tool set_theme]", json.dumps(
    S.set_theme(TARGET3, colors={"accent1": "#36C2FF", "dk2": "#0D2038"},
                fonts={"major": "Microsoft YaHei", "hans": "Microsoft YaHei"}),
    ensure_ascii=False)[:320])
print("[tool set_show]", json.dumps(
    S.set_show(TARGET3, {"loop": True, "show_type": "browse", "timings": False}),
    ensure_ascii=False)[:320])
print("[tool read_theme]", json.dumps(
    {k: v for k, v in S.read_theme(TARGET3).items() if k in ("colors", "fonts", "show")},
    ensure_ascii=False)[:400])