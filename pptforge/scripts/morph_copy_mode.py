"""正确的 morph 帧做法：复制原页（Duplicate）再改位置，保证元素属性完全一致。

为什么不能新建同版式页：
  新建页的占位符 bodyPr（autofit / insets / anchor）继承路径与原页不同，
  框位置一样、文字落点不一样 → 观感是"元素错位"。
  官方教程给的做法就是"复制幻灯片保名称，再校准位置"。

用法：
  python scripts/morph_copy_mode.py <base.pptx> <slide_index> <out.pptx>
"""
from __future__ import annotations

import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from mcp_ppt import comtools

MORPH = 3954          # ppEffectMorphByObject
DUR = 2.0             # 切换时长（秒）——对齐 designer 模板的 2000ms
ADV = 2.2             # 每页停留（秒）


def build(base: str, slide_idx: int, out: str, frames: list[dict],
          morph_duration: float = DUR, advance: float = ADV) -> dict:
    """frames[0] 是原页（不改）；frames[1:] 每项是 {形状名: {x,y,w,h,rot}} 的偏移。

    坐标单位：英寸，相对**原始位置**的绝对位置（与原页几何同坐标系）。
    """
    import shutil

    shutil.copyfile(base, out)
    app = comtools._app()
    pres = None
    log = []
    try:
        pres = app.Presentations.Open(os.path.abspath(out), ReadOnly=0, WithWindow=0)
        src = pres.Slides(slide_idx)

        # 第②步：先给参与 morph 的元素编码（两位数字 → !!num，其余文本 → !!title）
        for j in range(1, src.Shapes.Count + 1):
            shp = src.Shapes(j)
            try:
                text = str(shp.TextFrame.TextRange.Text).strip()
            except Exception:  # noqa: BLE001
                continue
            if not text:
                continue
            shp.Name = "!!num" if (len(text) == 2 and text.isdigit()) else "!!title"

        # 记录原页每个形状的几何，便于报告
        def geo(slide):
            out_map = {}
            for j in range(1, slide.Shapes.Count + 1):
                shp = slide.Shapes(j)
                try:
                    out_map[shp.Name] = [round(float(shp.Left) / 72, 3), round(float(shp.Top) / 72, 3),
                                         round(float(shp.Width) / 72, 3), round(float(shp.Height) / 72, 3)]
                except Exception:  # noqa: BLE001
                    continue
            return out_map

        base_geo = geo(src)
        log.append({"frame": 1, "mode": "base", "geo": base_geo})

        copies = []
        for i, frame in enumerate(frames[1:], start=2):
            dup = src.Duplicate()
            new_slide = dup.Item(dup.Count)
            copies.append(new_slide)
            for j in range(1, new_slide.Shapes.Count + 1):
                shp = new_slide.Shapes(j)
                target = frame.get(shp.Name)
                if not target:
                    continue
                if "x" in target:
                    shp.Left = float(target["x"]) * 72
                if "y" in target:
                    shp.Top = float(target["y"]) * 72
                if "w" in target:
                    shp.Width = float(target["w"]) * 72
                if "h" in target:
                    shp.Height = float(target["h"]) * 72
                if "rot" in target:
                    shp.Rotation = float(target["rot"])
                if "alpha" in target:
                    shp.Fill.Transparency = 1 - float(target["alpha"])
            log.append({"frame": i, "mode": "duplicate", "geo": geo(new_slide)})

        # 复制出来的帧排在原页之后，要移到原页**之前**（起点 → 中间 → 终点）
        for k, slide in enumerate(copies):
            slide.MoveTo(slide_idx + k)

        # 切换：第一帧从上一页用普通淡入；之后两页用 morph 做连续插值
        t_first = copies[0].SlideShowTransition
        t_first.EntryEffect = 1793       # ppEffectFade
        t_first.Duration = 0.8
        t_first.AdvanceOnClick = 0
        t_first.AdvanceOnTime = -1
        t_first.AdvanceTime = advance

        for slide in copies[1:] + [src]:
            t = slide.SlideShowTransition
            t.EntryEffect = MORPH
            t.Duration = morph_duration
            t.AdvanceOnClick = 0
            t.AdvanceOnTime = -1
            t.AdvanceTime = advance

        # 放映用计时自动推进
        ss = pres.SlideShowSettings
        ss.AdvanceMode = 2
        ss.ShowWithAnimation = -1
        pres.Save()
    finally:
        try:
            if pres is not None:
                pres.Close()
            app.Quit()
        except Exception:  # noqa: BLE001
            pass
    # 清空这几页的页内对象动画：morph 是页级过渡，两者叠加会让时间轴打架
    from mcp_ppt.deck import patch_deck
    patch_deck(out, [{"slide": slide_idx + i, "animations": []} for i in range(len(frames))])
    return {"file": os.path.abspath(out), "frames": len(frames), "log": log}


def main() -> int:
    if len(sys.argv) < 4:
        print(__doc__)
        return 2
    base, slide_idx, out = sys.argv[1], int(sys.argv[2]), sys.argv[3]
    # 3 帧：原位 → 下方 1.4in（同尺寸）→ 再回到中间态（只演示位移，不错位）
    frames = [
        {},
        {"!!num": {"y": 3.98}, "!!title": {"y": 3.96}},
        {"!!num": {"y": 3.28}, "!!title": {"y": 3.24}},
    ]
    r = build(base, slide_idx, out, frames)
    print(json.dumps({"file": r["file"], "frames": r["frames"]}, ensure_ascii=False))
    for row in r["log"]:
        print("  frame", row["frame"], row["mode"],
              {k: v for k, v in list(row["geo"].items())[:3]})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
