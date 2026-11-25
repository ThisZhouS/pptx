"""修 quote 装饰溢出（源码里是转义序列 \\u201c，需用 raw string 匹配）。"""
import io

P = r"E:\pptx\pptforge\mcp_ppt\design.py"
t = io.open(P, encoding="utf-8").read()

old = r'''        self._text(b, "mark", MARGIN, Inches(1.35), Inches(2), Inches(1.2),
                   [{"text": "\u201c", "size": 110, "bold": True, "color": t.accent}], font="Georgia")'''
new = r'''        self._text(b, "mark", MARGIN, Inches(1.2), Inches(2), Inches(1.6),
                   [{"text": "\u201c", "size": 72, "bold": True, "color": t.accent}], font="Georgia")'''
assert old in t, "未匹配"
t = t.replace(old, new, 1)
io.open(P, "w", encoding="utf-8").write(t)
print("[ok] quote mark: size 110→72, 框高 1.2→1.6in")