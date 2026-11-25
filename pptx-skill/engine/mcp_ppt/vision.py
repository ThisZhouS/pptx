"""泛用识图：把 pptx 渲染成图，算出可量化的视觉指纹。

为什么需要：结构拆解只能看到 XML，看不到"这一页长什么样"。
渲染成图后用像素统计，才能区分「深色极简」「图片主导」「大色块装饰」「文字密集」这些风格。

特征（全部 0–1 或计数）：
  dominant      前 5 主色及其占比（判断基调）
  ink_ratio     非背景像素占比（信息密度）
  edge_density  边缘像素占比（元素多不多、细节多不多）
  saturation    平均饱和度（莫兰迪低、马卡龙高、国潮中等）
  brightness    平均明度（深色主题 < 0.35）
  contrast      亮度标准差（对比强不强）
  regions       上/下/左/右 四块的主色（判断有没有大色块/页眉页脚）
  colorfulness  颜色多样性（唯一色数量）
"""
from __future__ import annotations

import os
import sys
from collections import Counter

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image, ImageFilter, ImageStat


def _hex(rgb) -> str:
    return "#%02X%02X%02X" % (int(rgb[0]), int(rgb[1]), int(rgb[2]))


def image_features(png_path: str, small: int = 240) -> dict:
    img = Image.open(png_path).convert("RGB")
    w, h = img.size
    thumb = img.resize((small, int(small * h / w)))
    px = list(thumb.getdata())
    total = len(px)

    # 主色（量化到 16 级，避免噪点）
    quant = [(r // 16 * 16, g // 16 * 16, b // 16 * 16) for r, g, b in px]
    top = Counter(quant).most_common(5)
    dominant = [{"color": _hex(c), "share": round(n / total, 3)} for c, n in top]

    # 信息密度：与最常见色不同的像素
    ink = 1 - top[0][1] / total if top else 0

    # 边缘密度
    edges = thumb.convert("L").filter(ImageFilter.FIND_EDGES)
    epx = list(edges.getdata())
    edge_density = sum(1 for v in epx if v > 40) / len(epx)

    # 明度 / 饱和度 / 对比
    hsv = thumb.convert("HSV")
    stat_v = ImageStat.Stat(hsv.getchannel("V"))
    stat_s = ImageStat.Stat(hsv.getchannel("S"))
    stat_l = ImageStat.Stat(thumb.convert("L"))

    # 四区域主色
    regions = {}
    boxes = {"top": (0, 0, w, h // 4), "bottom": (0, 3 * h // 4, w, h),
             "left": (0, 0, w // 4, h), "right": (3 * w // 4, 0, w, h)}
    for key, box in boxes.items():
        crop = img.crop(box).resize((40, 40))
        cpx = list(crop.getdata())
        q = [(r // 16 * 16, g // 16 * 16, b // 16 * 16) for r, g, b in cpx]
        c, n = Counter(q).most_common(1)[0]
        regions[key] = {"color": _hex(c), "share": round(n / len(cpx), 3)}

    return {
        "dominant": dominant,
        "ink_ratio": round(ink, 3),
        "edge_density": round(edge_density, 3),
        "saturation": round(stat_s.mean[0] / 255, 3),
        "brightness": round(stat_v.mean[0] / 255, 3),
        "contrast": round(stat_l.stddev[0] / 128, 3),
        "colorfulness": len({c for c in quant}),
        "regions": regions,
    }


def classify_style(f: dict) -> str:
    """把指纹归到一个可读风格标签。"""
    b, s, ink, edge = f["brightness"], f["saturation"], f["ink_ratio"], f["edge_density"]
    if b < 0.32 and ink < 0.45:
        return "深色极简"
    if b > 0.82 and ink < 0.18:
        return "浅色留白"
    if s < 0.18:
        return "低饱和（莫兰迪/灰调）"
    if s > 0.45 and b > 0.6:
        return "高饱和（马卡龙/潮流）"
    if ink > 0.6 and edge > 0.12:
        return "信息密集"
    if ink > 0.45:
        return "图片/色块主导"
    return "均衡型"


def deck_vision(path: str, pages: int = 1, width: int = 900) -> dict:
    """渲染前 N 页并逐页算特征。"""
    from mcp_ppt.comtools import export_png
    from mcp_ppt.layout import png_stats  # noqa: F401  （保持依赖一致）

    import tempfile

    tmp = tempfile.mkdtemp(prefix="vision_")
    res = export_png(path, tmp, width, max_pages=pages)
    if not res.get("available"):
        return {"file": path, "available": False}
    files = res.get("files") or []
    out = {"file": os.path.basename(path), "pages_rendered": len(files), "features": []}
    for f in files[:pages]:
        try:
            feat = image_features(f)
            feat["style"] = classify_style(feat)
            out["features"].append(feat)
        except Exception as exc:  # noqa: BLE001
            out["features"].append({"error": f"{type(exc).__name__}: {exc}"})
    return out

