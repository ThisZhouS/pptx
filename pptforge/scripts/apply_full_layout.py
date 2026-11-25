"""按全量布局数据修正 skill 资产与装饰层密度。"""
import io
import json
import os

ROOT = r"E:\pptx\pptforge"
S = os.path.join(ROOT, "samples")
ASSETS = r"E:\pptx\pptx-skill\assets"

# ---------- 1) layouts.json 用全量布局覆盖形状/装饰数据
lay = json.load(open(os.path.join(ASSETS, "layouts.json"), encoding="utf-8"))
full = json.load(open(os.path.join(S, "layout_all.json"), encoding="utf-8"))
lay["shapes_per_page_full"] = {"p50": full["fg_shapes"],
                               "note": "全量 16157 页的中位（抽样版曾给 15.0）"}
lay["decor_per_page"] = full["decor_total"]
lay["decor_mix"] = full["decor_mix"]
lay["columns_distribution_full"] = full["cols"]
lay["image_placement_full"] = full["image_placement"]
lay["band_shapes"] = full["band_shapes"]
lay["coverage_note"] = "内容页覆盖率 0.71；全量含章节页时中位会低一些"
lay["source"] = ("全部 2120 个模板（结构）+ 16157 页（布局全量）+ "
                 "15407 个形状（栅格）+ 2120 主题（配色）")
json.dump(lay, open(os.path.join(ASSETS, "layouts.json"), "w", encoding="utf-8"),
          ensure_ascii=False, indent=1)
print("[ok] layouts.json 全量覆盖：形状中位", full["fg_shapes"],
      "装饰中位", full["decor_total"])

# ---------- 2) design.py：装饰层减量（点 3→2、块 3→2），并支持 density
P = os.path.join(ROOT, "mcp_ppt", "design.py")
t = io.open(P, encoding="utf-8").read()
old = '''        t = self.theme
        # 右上角三个点（第一个用主色，其余弱化）
        x0 = SLIDE_W - MARGIN - Inches(0.76)
        for i in range(3):
            color = t.accent if i == 0 else (t.accent_2 if i == 1 else t.muted)
            self._rect(b, f"decor-dot{i + 1}", x0 + Inches(0.26) * i, Inches(0.40),
                       Inches(0.085), Inches(0.085), color, shape=MSO_SHAPE.OVAL)
        # 三个小色块（模板里 chips 占装饰大头）
        for i in range(3):
            self._rect(b, f"decor-chip{i + 1}", x0 + Inches(0.26) * i, Inches(0.60),
                       Inches(0.17), Inches(0.045), t.card)'''
new = '''        t = self.theme
        # 装饰数量按 density 控制：全量实测每页装饰中位 3 个，不铺满
        density = str(spec.get("decor_density") or "normal")
        n_dot = 2 if density == "normal" else 1
        n_chip = 2 if density == "normal" else 1
        x0 = SLIDE_W - MARGIN - Inches(0.76)
        for i in range(n_dot):
            color = t.accent if i == 0 else t.muted
            self._rect(b, f"decor-dot{i + 1}", x0 + Inches(0.26) * i, Inches(0.40),
                       Inches(0.085), Inches(0.085), color, shape=MSO_SHAPE.OVAL)
        for i in range(n_chip):
            self._rect(b, f"decor-chip{i + 1}", x0 + Inches(0.26) * i, Inches(0.60),
                       Inches(0.17), Inches(0.045), t.card)'''
assert old in t
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] design._decor 减量")