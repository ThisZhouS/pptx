"""design.py 改进：bullets 支持图片位与两列、装饰层补 3 个色块。"""
import io
import re

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()

# ---------- 1) 装饰层补 3 个色块（模板实测 chips 6.8/页）
old_decor_tail = '''        # 底部分隔线（比正文窄，制造第三条对齐线）'''
new_decor_tail = '''        # 三个小色块（模板里 chips 占装饰大头）
        for i in range(3):
            self._rect(b, f"decor-chip{i + 1}", x0 + Inches(0.26) * i, Inches(0.60),
                       Inches(0.17), Inches(0.045), t.card)
        # 底部分隔线（比正文窄，制造第三条对齐线）'''
assert old_decor_tail in t
t = t.replace(old_decor_tail, new_decor_tail, 1)

# ---------- 2) bullets 支持 image（左图右文）与 columns=2
old_bullets = '''        top = Inches(2.15)
        self._text(b, "bullets", MARGIN, top, CONTENT_W, Inches(4.6), paras,
                   size=spec.get("body_size", 18), space_after=16)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b'''

new_bullets = '''        top = Inches(2.15)
        body_h = Inches(4.35)
        image = spec.get("image")
        columns = int(spec.get("columns") or 1)
        if image:
            # 左 40% 图 + 右 60% 文（模板里"左侧图"占 27%，是最常见的图文关系）
            img_w = Emu(int(CONTENT_W * 0.38))
            img_h = body_h
            try:
                pic = b.slide.shapes.add_picture(image, MARGIN, top, img_w, img_h)
                b.add("image", pic, kind="image")
            except Exception as exc:  # noqa: BLE001
                b.shapes["__image_error__"] = ShapeMeta(f"图片加载失败: {exc}", -1, "error")
            body_x = MARGIN + img_w + Inches(0.55)
            body_w = CONTENT_W - img_w - Inches(0.55)
            self._text(b, "bullets", body_x, top, body_w, body_h, paras,
                       size=spec.get("body_size", 18), space_after=16)
        elif columns >= 2:
            # 两列：按"要点 + 说明"成对切分，避免把一对拆到两列
            pairs, cur = [], []
            for p in paras:
                if p.get("level", 0) == 0 and cur:
                    pairs.append(cur)
                    cur = [p]
                else:
                    cur.append(p)
            if cur:
                pairs.append(cur)
            half = (len(pairs) + 1) // 2
            gap = Inches(0.5)
            col_w = Emu(int((CONTENT_W - gap) / 2))
            for i, chunk in enumerate((pairs[:half], pairs[half:])):
                flat = [p for group in chunk for p in group]
                if not flat:
                    continue
                x = MARGIN if i == 0 else MARGIN + col_w + gap
                self._text(b, f"bullets-{i + 1}", x, top, col_w, body_h, flat,
                           size=spec.get("body_size", 18), space_after=16)
        else:
            self._text(b, "bullets", MARGIN, top, CONTENT_W, body_h, paras,
                       size=spec.get("body_size", 18), space_after=16)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))
        return b'''

assert old_bullets in t, "bullets 段落没匹配上"
t = t.replace(old_bullets, new_bullets, 1)

# ---------- 3) image 版式也补装饰层
old_img = '''        if paras:
            self._text(b, "bullets", MARGIN + img_w + Inches(0.5), top, CONTENT_W - img_w - Inches(0.5),
                       img_h, paras, size=18, space_after=12)
        self._footer(b, spec)'''
new_img = '''        if paras:
            self._text(b, "bullets", MARGIN + img_w + Inches(0.5), top, CONTENT_W - img_w - Inches(0.5),
                       img_h, paras, size=18, space_after=12)
        self._footer(b, spec)
        self._decor(b, spec, spec.get("index"))'''
if old_img in t:
    t = t.replace(old_img, new_img, 1)
    print("[ok] image 版式加装饰")

io.open(P, "w", encoding="utf-8").write(t)
print("[ok] design.py patched")