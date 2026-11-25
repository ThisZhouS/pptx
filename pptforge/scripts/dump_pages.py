"""列出每页形状的语义信息（名字/位置/尺寸/文本/字号），用于设计动画分组。"""
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
EMU_IN = 914400


def shapes_of(slide_root):
    out = []
    tree = slide_root.find(".//{%s}spTree" % P)
    if tree is None:
        return out
    for sp in tree:
        if not isinstance(sp.tag, str):
            continue
        local = etree.QName(sp).localname
        if local not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
            continue
        nv = sp.find(".//{%s}cNvPr" % P)
        rec = {"tag": local,
               "id": int(nv.get("id")) if nv is not None else None,
               "name": nv.get("name") if nv is not None else None}
        xfrm = sp.find(".//{%s}xfrm" % A)
        if xfrm is not None:
            off = xfrm.find("{%s}off" % A)
            ext = xfrm.find("{%s}ext" % A)
            if off is not None and ext is not None:
                rec["x"] = round(int(off.get("x")) / EMU_IN, 2)
                rec["y"] = round(int(off.get("y")) / EMU_IN, 2)
                rec["w"] = round(int(ext.get("cx")) / EMU_IN, 2)
                rec["h"] = round(int(ext.get("cy")) / EMU_IN, 2)
        texts, sizes = [], []
        for p in sp.iter("{%s}p" % A):
            t = "".join(p.itertext()).strip()
            if t:
                texts.append(t)
        for rpr in sp.iter("{%s}rPr" % A):
            if rpr.get("sz"):
                sizes.append(int(rpr.get("sz")) / 100)
        for d in sp.iter("{%s}defRPr" % A):
            if d.get("sz"):
                sizes.append(int(d.get("sz")) / 100)
        rec["text"] = " | ".join(texts)[:90]
        rec["paras"] = len(texts)
        rec["sizes"] = sorted(set(sizes), reverse=True)[:3]
        out.append(rec)
    return out


def main(path: str) -> int:
    with zipfile.ZipFile(path) as zf:
        slides = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                        key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        report = []
        for idx, part in enumerate(slides, start=1):
            root = etree.fromstring(zf.read(part))
            report.append({"index": idx, "shapes": shapes_of(root)})
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "samples",
                       "pages.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(report, fh, ensure_ascii=False, indent=1)
    for page in report:
        print(f"=== slide {page['index']} ({len(page['shapes'])} shapes) ===")
        for s in page["shapes"]:
            print(f"  id={s['id']:2d} {s['tag']:12s} {str(s['name'])[:22]:22s} "
                  f"({s.get('x')},{s.get('y')}) {s.get('w')}x{s.get('h')} "
                  f"sz={s['sizes']} p={s['paras']} :: {s['text'][:60]}")
    print(f"[written] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1]))