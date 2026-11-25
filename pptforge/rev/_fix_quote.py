"""修 quote 装饰的溢出：字号 110→72，框高 1.2→1.6in。"""
import io

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()

old = '''        self._text(b, "mark", MARGIN, Inches(1.35), Inches(2), Inches(1.2),
                   [{"text": "\u201c", "size": 110, "bold": True, "color": t.accent}], font="Georgia")'''
new = '''        self._text(b, "mark", MARGIN, Inches(1.2), Inches(2), Inches(1.6),
                   [{"text": "\u201c", "size": 72, "bold": True, "color": t.accent}], font="Georgia")'''
if old in t:
    t = t.replace(old, new, 1)
    io.open(P, "w", encoding="utf-8").write(t)
    print("[ok] quote mark 调整")
else:
    print("[miss] 检查当前尺寸：")
    i = t.find('"mark"')
    print(repr(t[i - 20:i + 200]))