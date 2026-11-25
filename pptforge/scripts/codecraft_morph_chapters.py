"""按 morph 流程改造 CodeCraft：5 个章节页 → 各 3 帧 morph 序列，内容页保持 fade。

做法（不动原章节页内容）：
  ① 拆解：找出章节页（含 01–04 编号）与其编号/标题的实际几何（COM 读）
  ② 编码：章节页的编号/标题改名 !!num / !!title（morph 强制匹配）
  ③ plan：每章 3 帧 —— 起点（编号居中放大）→ 中间点（缩到左上）→ 终点（原章节页）
  ④ 编写：用同一 layout 新建前两帧（继承背景与占位符样式），插到原章节页之前，
          三页都设 morph + advTm；原章节页清空对象动画（morph 与页内时间轴不叠加）
"""
from __future__ import annotations

import json
import os
import shutil
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.util import Emu, Inches

from mcp_ppt import comtools
from mcp_ppt.catalog import get_catalog
from mcp_ppt.deck import patch_deck
from mcp_ppt.inspect import validate_deck
from mcp_ppt.morph import decompose
from mcp_ppt.ooxml import set_slide_transition

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "samples", "codecraft-optimized-v2.pptx")
DST = os.path.join(ROOT, "samples", "codecraft-morph-chapters.pptx")
EMU = 914400

MORPH = {"effect": "morph", "duration": 2.0}
FRAME_ADV = 2.2      # 前两帧停留（模板实测 morph 2s + 停留 2s 左右最稳）
END_ADV = 2.8        # 章节页停留（给观众认字）


def chapters(path: str) -> list[dict]:
    dec = decompose(path)
    out = []
    for page in dec["pages_detail"]:
        texts = [s["text"].strip() for s in page["shapes"] if s["text"].strip()]
        nums = [t for t in texts if len(t) == 2 and t.isdigit()]
        if nums and len(page["shapes"]) <= 3:
            out.append({"slide": page["slide"], "num": nums[0],
                        "title": next((t for t in texts if t not in nums), ""),
                        "shapes": page["shapes"]})
    return out


def read_geometry(path: str, slide_idx: int) -> dict:
    """用 COM 读出该页编号/标题的实际几何（英寸）。"""
    app = comtools._app()
    pres = None
    info = {}
    try:
        pres = comtools._open(app, path)
        slide = pres.Slides(slide_idx)
        for j in range(1, slide.Shapes.Count + 1):
            shp = slide.Shapes(j)
            try:
                text = str(shp.TextFrame.TextRange.Text).strip()
            except Exception:  # noqa: BLE001
                continue
            if not text:
                continue
            key = "num" if (len(text) == 2 and text.isdigit()) else "title"
            info[key] = {
                "name": shp.Name,
                "box": [round(float(shp.Left) / 72, 3), round(float(shp.Top) / 72, 3),
                        round(float(shp.Width) / 72, 3), round(float(shp.Height) / 72, 3)],
                "size": round(float(shp.TextFrame.TextRange.Font.Size), 1),
            }
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    return info


def set_placeholder_text(slide, text: str, kind: str = "body") -> object | None:
    """在新建页上写好占位符文本（保留 layout 继承的字体/颜色）。"""
    from pptx.enum.shapes import PP_PLACEHOLDER

    chosen = None
    for shp in slide.placeholders:
        ptype = shp.placeholder_format.type
        if kind == "title" and ptype in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            chosen = shp
            break
        if kind == "body" and ptype not in (PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE):
            chosen = shp
            break
    if chosen is None:
        return None
    chosen.text_frame.text = text
    return chosen


def main() -> int:
    if not os.path.exists(SRC):
        print("缺少源文件：", SRC)
        return 1
    shutil.copyfile(SRC, DST)
    chs = chapters(DST)
    print(f"[① 拆解] {len(chs)} 个章节页:",
          json.dumps([{k: c[k] for k in ("slide", "num", "title")} for c in chs],
                     ensure_ascii=False))

    geo = {c["slide"]: read_geometry(DST, c["slide"]) for c in chs}
    print("[① 几何]", json.dumps(geo, ensure_ascii=False)[:400])

    prs = Presentation(DST)
    sld_id_lst = prs.slides._sldIdLst
    plans = []
    frame_pages: list[dict] = []
    # 最终页码 = 原页码 + 2×(排在前面的章节数) + 2（自己在前面插的两帧）
    for i, c in enumerate(sorted(chs, key=lambda x: x["slide"])):
        end = c["slide"] + 2 * i + 2
        c["frame_pages"] = [end - 2, end - 1, end]
    # 从后往前插，避免索引错位
    for ch in sorted(chs, key=lambda c: -c["slide"]):
        idx = ch["slide"] - 1
        src_slide = prs.slides[idx]
        layout = src_slide.slide_layout
        g = geo[ch["slide"]]
        num_box = g.get("num", {}).get("box") or [4.4, 2.2, 4.5, 3.6]
        title_box = g.get("title", {}).get("box") or [2.9, 2.6, 8.6, 1.2]

        # 帧 1（起点）：编号居中放大，标题在下方
        a = prs.slides.add_slide(layout)
        num_a = set_placeholder_text(a, ch["num"], "body")
        ttl_a = set_placeholder_text(a, ch["title"], "title")
        if num_a is not None:
            num_a.left, num_a.top = Emu(int(Inches(0.40))), Emu(int(Inches(4.05)))
            num_a.width, num_a.height = Emu(int(Inches(2.68))), Emu(int(Inches(2.32)))
            num_a.name = "!!num"
        if ttl_a is not None:
            ttl_a.left, ttl_a.top = Emu(int(Inches(4.08))), Emu(int(Inches(4.30)))
            ttl_a.width, ttl_a.height = Emu(int(Inches(7.80))), Emu(int(Inches(1.60)))
            ttl_a.name = "!!title"

        # 帧 2（中间点）：编号缩到左上，标题上移
        b = prs.slides.add_slide(layout)
        num_b = set_placeholder_text(b, ch["num"], "body")
        ttl_b = set_placeholder_text(b, ch["title"], "title")
        if num_b is not None:
            num_b.left, num_b.top = Emu(int(Inches(0.40))), Emu(int(Inches(3.18)))
            num_b.width, num_b.height = Emu(int(Inches(2.68))), Emu(int(Inches(2.32)))
            num_b.name = "!!num"
        if ttl_b is not None:
            ttl_b.left, ttl_b.top = Emu(int(Inches(4.08))), Emu(int(Inches(3.40)))
            ttl_b.width, ttl_b.height = Emu(int(Inches(7.80))), Emu(int(Inches(1.60)))
            ttl_b.name = "!!title"

        # 终点帧（原章节页）：改名参与 morph，清空页内对象动画
        for shp in src_slide.shapes:
            try:
                text = str(shp.text_frame.text).strip()
            except Exception:  # noqa: BLE001
                continue
            if not text:
                continue
            shp.name = "!!num" if (len(text) == 2 and text.isdigit()) else "!!title"

        # 把两帧移到原章节页之前（保持 帧1 → 帧2 → 章节页）
        ids = list(sld_id_lst)
        a_id, b_id = ids[-2], ids[-1]
        target = ids[idx]
        sld_id_lst.remove(a_id)
        sld_id_lst.remove(b_id)
        target.addprevious(a_id)
        target.addprevious(b_id)

        a_page, b_page, end_page = ch["frame_pages"]
        frame_pages.append({"num": ch["num"], "title": ch["title"],
                            "frames": ch["frame_pages"],
                            "num_box_final": num_box, "title_box_final": title_box})

        plans.append({"original_slide": ch["slide"], "num": ch["num"],
                      "title": ch["title"], "frames": 3,
                      "num_box_final": num_box, "title_box_final": title_box})
        print(f"  [③ plan] 章节 {ch['num']} {ch['title']}: 3 帧 → 页码 {a_page}/{b_page}/{end_page}")

    prs.save(DST)
    print(f"[④ 编写] {DST}（共 {len(prs.slides.__iter__.__self__._sldIdLst)} 页）")

    # 给新帧与章节页设定切换；章节页清空页内动画
    total = len(Presentation(DST).slides)
    print("[③ 帧位置]", [fp["frames"] for fp in frame_pages])

    patches = []
    for fp in frame_pages:
        for page_no, adv in zip(fp["frames"], (FRAME_ADV, FRAME_ADV, END_ADV)):
            if 1 <= page_no <= total:
                patches.append({"slide": page_no,
                                "transition": {**MORPH, "advance_after": adv},
                                "animations": []})
    result = patch_deck(DST, patches)
    print("[patch]", len(patches), "页；warnings:", result["warnings"][:4])

    val = validate_deck(DST)
    print("[validate]", val["stats"], val["problems"][:4])
    if comtools.available():
        com = comtools.verify(DST)
        entries = [(s["index"], s["entry_effect"]) for s in com["slides"]]
        morph_pages = [e for e in entries if e[1] == 3954]
        print("[com] 总页数", len(entries), "morph 页", len(morph_pages),
              "其它切换", sorted({e[1] for e in entries if e[1] != 3954}))
    with open(os.path.join(ROOT, "samples", "codecraft-morph-plan.json"), "w",
              encoding="utf-8") as fh:
        json.dump({"chapters": plans, "frame_pages": frame_pages,
                   "validate": val["stats"]}, fh, ensure_ascii=False, indent=1)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



