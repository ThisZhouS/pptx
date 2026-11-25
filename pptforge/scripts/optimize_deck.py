"""用 pptforge 优化一份真实演示：字体可用性 → 切换 → 按语义分组的动画。

策略（保守）：
  1. 原件只读，优化结果另存；所有改动可回滚
  2. 字体替换为系统可用字体（缺字体会让演示机回退成宋体，版式崩）
  3. 全篇统一切换；按每页形状语义（标题/装饰/图片/正文/目录条目）自动分组加动画
  4. 不移动、不改写任何文字

用法：
  python scripts/optimize_deck.py "<源文件.pptx>" "<输出.pptx>" [--keep-fonts]
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sys
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from lxml import etree

from mcp_ppt.deck import patch_deck
from mcp_ppt.inspect import inspect_deck, validate_deck
from mcp_ppt.layout import layout_report
from mcp_ppt.theme import extract_theme

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# 常见"看起来会崩"的字体 → 系统可用替代
FONT_MAP = {
    "MiSans Normal": "Microsoft YaHei",
    "MiSans": "Microsoft YaHei",
    "汉仪刚艺体-85W": "Microsoft YaHei",
    "汉仪刚艺体": "Microsoft YaHei",
    "HarmonyOS Sans SC": "Microsoft YaHei",
    "思源黑体": "Microsoft YaHei",
    "Source Han Sans": "Microsoft YaHei",
}


def replace_fonts(src: str, dst: str, mapping: dict[str, str]) -> dict:
    """在包内所有 XML 里替换 typeface（含 theme/layout/master/slide）。"""
    counts: dict[str, int] = {}
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                text = data.decode("utf-8")
                for old, new in mapping.items():
                    n = text.count(f'typeface="{old}"')
                    if n:
                        text = text.replace(f'typeface="{old}"', f'typeface="{new}"')
                        counts[old] = counts.get(old, 0) + n
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return counts


def shape_kind(rec: dict) -> str:
    name = str(rec.get("name") or "")
    tag = rec.get("tag")
    if tag == "pic":
        return "image"
    if "装饰" in name:
        return "deco"
    if "标题" in name or "副标题" in name:
        return "title"
    if "图片" in name or "内容占位符" in name:
        return "image"
    if "文本占位符" in name:
        return "text"
    return "other"


def slide_shapes(path: str) -> list[list[dict]]:
    """每页的形状（保持 spTree 顺序 = z 序）。"""
    pages = []
    with zipfile.ZipFile(path) as zf:
        parts = sorted([n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                       key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))
        for part in parts:
            root = etree.fromstring(zf.read(part))
            tree = root.find(".//{%s}spTree" % P)
            recs = []
            for sp in tree if tree is not None else []:
                if not isinstance(sp.tag, str):
                    continue
                local = etree.QName(sp).localname
                if local not in ("sp", "pic", "graphicFrame", "cxnSp", "grpSp"):
                    continue
                nv = sp.find(".//{%s}cNvPr" % P)
                if nv is None:
                    continue
                text = "".join(sp.itertext()).strip()
                xfrm = sp.find(".//{%s}xfrm" % A)
                top = None
                if xfrm is not None:
                    off = xfrm.find("{%s}off" % A)
                    if off is not None:
                        top = int(off.get("y", 0))
                recs.append({"id": int(nv.get("id")), "name": nv.get("name"), "tag": local,
                             "text": text, "top": top, "paras": len([p for p in sp.iter("{%s}p" % A)
                                                                     if "".join(p.itertext()).strip()])})
            pages.append(recs)
    return pages


def plan_page(recs: list[dict], index: int) -> dict:
    """按语义给一页排动画。

    参数取自真实模板（analyze_templates.py）：进入用「上升」、时长 800ms、
    触发用 with_previous + stagger 错峰（而不是 after_previous 串行）。
    """
    kinds = [(r, shape_kind(r)) for r in recs]
    titles = [r for r, k in kinds if k == "title"]
    decos = [r for r, k in kinds if k == "deco"]
    images = [r for r, k in kinds if k == "image"]
    texts = [r for r, k in kinds if k == "text"]
    codes = [r for r in texts if re.fullmatch(r"0\d", str(r.get("text", "")).strip())]
    ENTER, DUR, STAG = "ascend", 0.8, 0.12

    # 目录页：标题 → 每个条目（编号 + 名称）成组错峰出现
    if len(codes) >= 3:
        anims = []
        for r in titles[:1]:
            anims.append({"target": r["name"], "effect": ENTER, "trigger": "click",
                          "duration": DUR})
        for r in titles[1:]:
            anims.append({"target": r["name"], "effect": ENTER, "trigger": "with_previous",
                          "duration": DUR, "stagger": 0.15})
        ordered = [r for r in recs if r in codes or (shape_kind(r) == "text" and r not in codes)]
        first = True
        for r in ordered:
            anims.append({"target": r["name"], "effect": ENTER,
                          "trigger": "click" if first else "with_previous",
                          "duration": DUR, "stagger": 0.0 if first else STAG})
            first = False
        return {"animations": anims, "transition": {"effect": "fade", "duration": 0.8}}

    # 章节页：编号 + 标题，轻微上浮
    if len(recs) <= 3 and titles:
        anims = []
        first = True
        for r in recs:
            if not str(r.get("text", "")).strip():
                continue
            anims.append({"target": r["name"], "effect": ENTER,
                          "trigger": "click" if first else "with_previous",
                          "duration": DUR, "stagger": 0.0 if first else 0.15})
            first = False
        return {"animations": anims, "transition": {"effect": "fade", "duration": 0.8}}

    # 普通页：标题 click → 其余元素 with_previous 错峰（总时长控制在 ~1.3s）
    anims = []
    if titles:
        anims.append({"target": titles[0]["name"], "effect": ENTER, "trigger": "click",
                      "duration": DUR})
        seq = decos + images + texts + decos[len(decos):]
        for r in seq:
            anims.append({"target": r["name"], "effect": ENTER, "trigger": "with_previous",
                          "duration": DUR, "stagger": STAG})
    else:
        for i, r in enumerate(recs):
            anims.append({"target": r["name"], "effect": ENTER,
                          "trigger": "click" if i == 0 else "with_previous",
                          "duration": DUR, "stagger": 0.0 if i == 0 else STAG})
    return {"animations": anims, "transition": {"effect": "fade", "duration": 0.8}}


def fix_width_overflow(path: str, ratio: float = 1.05, min_width: float = 20.0) -> list[str]:
    """只修"文字宽度超出文本框"：给 bodyPr 加 <a:normAutofit/>（PowerPoint 自动缩字）。

    做法是 COM 只读扫描拿坐标，XML 层改 bodyPr —— 不用 PowerPoint 保存，
    因为 PowerPoint 的 Save 会重新序列化 XML，实测会丢掉手写的 timing 动画（101 → 98）。
    不动高度溢出（目录页大号编号本就是故意超出框的），也不碰窄装饰条（其"文本"是空行高，属误报）。
    """
    from mcp_ppt import comtools

    if not comtools.available():
        return ["COM 不可用，跳过溢出修复"]
    targets: list[tuple[int, int, str, int]] = []
    app = comtools._app()
    pres = None
    try:
        pres = comtools._open(app, path)
        for i in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(i)
            for j in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(j)
                try:
                    if not getattr(shp, "HasTextFrame", 0):
                        continue
                    width = float(shp.Width)
                    if width < min_width:
                        continue
                    tr = shp.TextFrame.TextRange
                    if tr.Length == 0:
                        continue
                    # 短文本（目录编号 01/02…、"+" 等）是设计意图，不缩
                    if len(str(tr.Text).strip()) <= 3:
                        continue
                    bound_w = float(tr.BoundWidth)
                    if bound_w > width * ratio:
                        scale = max(60, min(99, int(width / bound_w * 100) - 1))
                        targets.append((i, int(shp.Id),
                                        f"#{i} {shp.Name} ({round(bound_w)}→{round(width)}pt, {scale}%)",
                                        scale))
                except Exception:  # noqa: BLE001
                    continue
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass

    if not targets:
        return []
    # XML 层写入 normAutofit（python-pptx 保存会保留 timing）
    from lxml import etree
    from pptx import Presentation

    A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
    prs = Presentation(path)
    fixed: list[str] = []
    for slide_idx, shape_id, label, scale in targets:
        slide = prs.slides[slide_idx - 1]
        for shp in slide.shapes:
            if shp.shape_id != shape_id or not shp.has_text_frame:
                continue
            body_pr = shp.text_frame._txBody.find("{%s}bodyPr" % A_NS)
            if body_pr is None:
                continue
            for tag in ("noAutofit", "normAutofit", "spAutoFit"):
                for old in body_pr.findall("{%s}%s" % (A_NS, tag)):
                    body_pr.remove(old)
            node = etree.Element("{%s}normAutofit" % A_NS)
            node.set("fontScale", str(scale * 1000))
            warp = body_pr.find("{%s}prstTxWarp" % A_NS)
            if warp is not None:
                warp.addnext(node)
            else:
                body_pr.insert(0, node)
            fixed.append(label)
            break
    if fixed:
        prs.save(path)
    return fixed


def main() -> int:
    if len(sys.argv) < 3:
        print(__doc__)
        return 2
    src, dst = sys.argv[1], sys.argv[2]
    keep_fonts = "--keep-fonts" in sys.argv
    work = os.path.join(os.path.dirname(os.path.abspath(dst)), "_opt_work.pptx")

    before = inspect_deck(src)
    audit_before = {"slides": before["slide_count"], "theme": extract_theme(src)}

    # 1) 字体替换（写到临时文件）
    font_stats = {}
    if keep_fonts:
        shutil.copyfile(src, work)
    else:
        font_stats = replace_fonts(src, work, FONT_MAP)
        print("[fonts]", font_stats or "无需替换")

    # 2) 计划每页动画与切换
    pages = slide_shapes(work)
    patches = []
    for i, recs in enumerate(pages, start=1):
        plan = plan_page(recs, i)
        patches.append({"slide": i, "transition": plan["transition"],
                        "animations": plan["animations"]})
        print(f"  slide {i:2d}: {len(plan['animations'])} 个动画 "
              f"({plan['transition']['effect']})")

    result = patch_deck(work, patches)
    print("[patch] warnings:", result["warnings"][:6])

    # 3) 收尾
    if os.path.exists(dst):
        os.remove(dst)
    shutil.move(work, dst)

    # 4) 宽度溢出修复（com 侧，写回文件）
    fixed = fix_width_overflow(dst) if "--no-overflow-fix" not in sys.argv else []
    print("[overflow fixed]", fixed or "无")

    val = validate_deck(dst)
    lay = layout_report(dst)
    print("[validate]", val["stats"], val["problems"][:4])
    print("[layout] issue_count", lay.get("issue_count"),
          [p["issues"] for p in lay.get("pages", []) if p["issues"]][:6])
    out = {"source": os.path.abspath(src), "output": os.path.abspath(dst),
           "fonts_replaced": font_stats, "validate": val["stats"],
           "layout_issues": lay.get("issue_count"),
           "overflow_fixed": fixed,
           "layout_detail": [f"#{p['index']} " + "; ".join(p["issues"])
                             for p in lay.get("pages", []) if p["issues"]],
           "pages": [{"index": i, "shapes": len(r)} for i, r in enumerate(pages, start=1)]}
    with open(os.path.join(os.path.dirname(os.path.abspath(dst)), "optimize_report.json"),
              "w", encoding="utf-8") as fh:
        json.dump(out, fh, ensure_ascii=False, indent=1)
    print(f"[done] {dst}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

