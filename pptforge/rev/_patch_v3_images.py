"""给 v3 的三页技术页加左侧图片位，其余页保持纯文字版式。"""
import io
import os

P = r"E:\pptx\pptforge\scripts\build_codecraft_v3.py"
t = io.open(P, encoding="utf-8").read()

ART = r"samples\art".replace("\\", "/")

# AST / GNN / LLM 三页加图（模板里"左侧图"占 27%，是最常见的图文关系）
pairs = [
    ('''        {"layout": "bullets", "kicker": "02 · 技术 1/4", "title": "代码抽象语法树解析",''',
     f'''        {{"layout": "bullets", "kicker": "02 · 技术 1/4", "title": "代码抽象语法树解析",
         "image": "{ART}/art-ast.png",'''),
    ('''        {"layout": "bullets", "kicker": "02 · 技术 2/4", "title": "图神经网络建模",''',
     f'''        {{"layout": "bullets", "kicker": "02 · 技术 2/4", "title": "图神经网络建模",
         "image": "{ART}/art-gnn.png",'''),
    ('''        {"layout": "bullets", "kicker": "02 · 技术 3/4", "title": "领域专用大模型",''',
     f'''        {{"layout": "bullets", "kicker": "02 · 技术 3/4", "title": "领域专用大模型",
         "image": "{ART}/art-llm.png",'''),
    # 亮点页用两列
    ('''        {"layout": "bullets", "kicker": "03 · 深度", "title": "三种能力叠在一起",''',
     '''        {"layout": "bullets", "kicker": "03 · 深度", "title": "三种能力叠在一起", "columns": 2,'''),
]
for old, new in pairs:
    if old not in t:
        print("[miss]", old.strip()[:50])
        continue
    t = t.replace(old, new, 1)
    print("[ok]", old.strip()[:46])

io.open(P, "w", encoding="utf-8").write(t)
print("patched")