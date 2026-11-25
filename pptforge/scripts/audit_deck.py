"""审计任意 pptx：版式尺寸、每页形状与文本、字体/字号/颜色分布、切换与动画、溢出。"""
import json
import os
import re
import sys
import zipfile
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lxml import etree

from mcp_ppt.inspect import inspect_deck
from mcp_ppt.layout import layout_report

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def audit(path: str, out_json: str | None = None) -> dict:
    info = {"file": os.path.abspath(path), "exists": os.path.exists(path)}
    if not info["exists"]:
        return info
    info["size_kb"] = round(os.path.getsize(path) / 1024, 1)

    deck = inspect_deck(path)
    info["slide_count"] = deck["slide_count"]

    fonts, sizes, colors = Counter(), Counter(), Counter()
    pages = []
    with zipfile.ZipFile(path) as zf:
        slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for idx, part in enumerate(slides, start=1):
            xml = zf.read(part)
            root = etree.fromstring(xml)
            for rpr in root.iter("{%s}rPr" % A):
                if rpr.get("sz"):
                    sizes[int(rpr.get("sz")) // 100] += 1
                for tag in ("latin", "ea"):
                    for node in root.iter():
                        pass
            for node in root.iter("{%s}latin" % A):
                if node.get("typeface"):
                    fonts[node.get("typeface")] += 1
            for node in root.iter("{%s}ea" % A):
                if node.get("typeface"):
                    fonts[node.get("typeface")] += 1
            for node in root.iter("{%s}srgbClr" % A):
                if node.get("val"):
                    colors[node.get("val").upper()] += 1
            tr = root.find("{%s}transition" % P)
            tm = root.find("{%s}timing" % P)
            n_anim = 0
            if tm is not None:
                n_anim = sum(1 for c in tm.iter("{%s}cTn" % P) if c.get("presetID"))
            pages.append({"index": idx, "transition": tr is not None, "animations": n_anim,
                          "shapes": len(deck["slides"][idx - 1]["shapes"]) if idx <= len(deck["slides"]) else 0,
                          "texts": [s["text"] for s in (deck["slides"][idx - 1]["shapes"] if idx <= len(deck["slides"]) else []) if s.get("text")][:4]})
    info["fonts"] = fonts.most_common(8)
    info["sizes_pt"] = sorted(sizes.items(), key=lambda kv: -kv[1])[:10]
    info["colors"] = colors.most_common(10)
    info["pages"] = pages
    info["pages_with_transition"] = sum(1 for p in pages if p["transition"])
    info["animations_total"] = sum(p["animations"] for p in pages)

    lay = layout_report(path)
    info["layout"] = {"available": lay.get("available"),
                      "issue_count": lay.get("issue_count"),
                      "issues": [f"#{p['index']} " + "; ".join(p["issues"])
                                 for p in lay.get("pages", []) if p["issues"]][:12]}

    # 页尺寸
    try:
        with zipfile.ZipFile(path) as zf:
            pres = etree.fromstring(zf.read("ppt/presentation.xml"))
            sz = pres.find("{%s}sldSz" % P)
            if sz is not None:
                info["slide_size_emu"] = [int(sz.get("cx")), int(sz.get("cy"))]
                info["slide_size_in"] = [round(int(sz.get("cx")) / 914400, 2),
                                         round(int(sz.get("cy")) / 914400, 2)]
    except Exception:  # noqa: BLE001
        pass

    if out_json:
        with open(out_json, "w", encoding="utf-8") as fh:
            json.dump(info, fh, ensure_ascii=False, indent=1)
    return info


if __name__ == "__main__":
    target = sys.argv[1]
    result = audit(target, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                        "..", "samples", "audit.json"))
    print(json.dumps(result, ensure_ascii=False, indent=1)[:4000])