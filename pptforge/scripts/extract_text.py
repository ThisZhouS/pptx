"""提取 pptx 全文：按页、按形状、按段落，供重排设计使用。"""
import json
import os
import re
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lxml import etree

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def extract(path: str) -> list[dict]:
    pages = []
    with zipfile.ZipFile(path) as zf:
        parts = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                       key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for idx, part in enumerate(parts, start=1):
            root = etree.fromstring(zf.read(part))
            tree = root.find(".//{%s}spTree" % P)
            items = []
            for sp in tree if tree is not None else []:
                if not isinstance(sp.tag, str):
                    continue
                nv = sp.find(".//{%s}cNvPr" % P)
                name = nv.get("name") if nv is not None else ""
                paras = []
                for p in sp.iter("{%s}p" % A):
                    text = "".join(p.itertext()).strip()
                    if text:
                        sizes = [int(r.get("sz")) / 100 for r in p.iter("{%s}rPr" % A) if r.get("sz")]
                        paras.append({"text": text, "size": max(sizes) if sizes else None})
                if paras:
                    items.append({"name": name, "paras": paras})
            pages.append({"slide": idx, "items": items})
    return pages


def main() -> int:
    path = sys.argv[1]
    pages = extract(path)
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "samples", "codecraft_text.json")
    with open(out, "w", encoding="utf-8") as fh:
        json.dump(pages, fh, ensure_ascii=False, indent=1)
    for page in pages:
        print(f"===== slide {page['slide']} =====")
        for it in page["items"]:
            head = it["paras"][0]
            print(f"  [{it['name']}] {head['text'][:70]}  (size={head['size']}, {len(it['paras'])} 段)")
    print(f"[written] {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())