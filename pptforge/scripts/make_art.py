"""生成几张技术风抽象图，供 bullets 版的左侧图片位使用。"""
import math
import os

from PIL import Image, ImageDraw


def gradient_grid(path: str, c1: str, c2: str, w: int = 720, h: int = 780) -> None:
    img = Image.new("RGB", (w, h))
    px = img.load()
    r1, g1, b1 = (int(c1[i:i + 2], 16) for i in (0, 2, 4))
    r2, g2, b2 = (int(c2[i:i + 2], 16) for i in (0, 2, 4))
    for y in range(h):
        for x in range(w):
            t = (x / w * 0.55 + y / h * 0.45)
            px[x, y] = (int(r1 + (r2 - r1) * t), int(g1 + (g2 - g1) * t), int(b1 + (b2 - b1) * t))
    d = ImageDraw.Draw(img, "RGBA")
    # 网格
    for gx in range(0, w, 45):
        d.line([(gx, 0), (gx, h)], fill=(255, 255, 255, 18), width=1)
    for gy in range(0, h, 45):
        d.line([(0, gy), (w, gy)], fill=(255, 255, 255, 18), width=1)
    img.save(path)


def rings(path: str, c: str, w: int = 720, h: int = 780, rings: int = 7) -> None:
    img = Image.new("RGB", (w, h), "#0B1020")
    d = ImageDraw.Draw(img, "RGBA")
    r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    cx, cy = w // 2, h // 2
    for i in range(rings):
        rad = 60 + i * 42
        alpha = 200 - i * 22
        d.ellipse([cx - rad, cy - rad, cx + rad, cy + rad], outline=(r, g, b, alpha), width=3)
    for i in range(24):
        a = i / 24 * math.tau
        x2 = cx + math.cos(a) * 40
        y2 = cy + math.sin(a) * 40
        d.line([(cx, cy), (x2, y2)], fill=(r, g, b, 90), width=1)
    img.save(path)


def nodes(path: str, c: str, w: int = 720, h: int = 780) -> None:
    img = Image.new("RGB", (w, h), "#0B1020")
    d = ImageDraw.Draw(img, "RGBA")
    r, g, b = (int(c[i:i + 2], 16) for i in (0, 2, 4))
    pts = [(140, 150), (360, 110), (560, 230), (180, 380), (420, 400),
           (600, 520), (240, 610), (460, 690)]
    for i, p in enumerate(pts):
        for q in pts[i + 1:]:
            if abs(p[0] - q[0]) + abs(p[1] - q[1]) < 300:
                d.line([p, q], fill=(r, g, b, 55), width=2)
    for p in pts:
        d.ellipse([p[0] - 9, p[1] - 9, p[0] + 9, p[1] + 9], fill=(r, g, b, 230))
    img.save(path)


def main() -> int:
    out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                       "samples", "art")
    os.makedirs(out, exist_ok=True)
    gradient_grid(os.path.join(out, "art-ast.png"), "1B2A4A", "36C2FF")
    rings(os.path.join(out, "art-gnn.png"), "6FE3C4")
    nodes(os.path.join(out, "art-llm.png"), "FFD166")
    print("generated:", os.listdir(out))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())