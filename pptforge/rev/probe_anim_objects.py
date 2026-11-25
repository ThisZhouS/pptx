"""探测动画 COM 对象的可用属性名（Acceleration / AfterEffect / 触发等）。"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import win32com.client as win32

app = win32.Dispatch("PowerPoint.Application")
app.Visible = True
app.DisplayAlerts = 1
pres = app.Presentations.Add()
pres.PageSetup.SlideWidth = 960
pres.PageSetup.SlideHeight = 540
slide = pres.Slides.Add(1, 12)
sh = slide.Shapes.AddShape(1, 100, 100, 200, 120)
seq = slide.TimeLine.MainSequence
eff = seq.AddEffect(sh, 10, 0, 1)   # fade

def dump(obj, title, limit=120):
    names = []
    for n in dir(obj):
        if n.startswith("_"):
            continue
        names.append(n)
    print(f"--- {title} ({len(names)}) ---")
    print(", ".join(sorted(names)[:limit]))

dump(eff, "Effect")
dump(eff.Timing, "Effect.Timing")
dump(slide.TimeLine, "TimeLine")
dump(seq, "Sequence")
try:
    dump(seq[1].Behaviors, "Behavior[]")
except Exception as e:
    print("Behaviors err:", e)

# 逐项试设置
for cand in ["SmoothStart", "SmoothEnd", "Acceleration", "Deceleration"]:
    try:
        setattr(eff.Timing, cand, 0.5)
        print(f"Timing.{cand} = OK ->", getattr(eff.Timing, cand))
    except Exception as e:
        print(f"Timing.{cand} -> {type(e).__name__}")
for cand in ["AfterEffect", "Exit", "StopAfterSlides", "Dim"]:
    try:
        setattr(eff, cand, 1)
        print(f"Effect.{cand} = OK ->", getattr(eff, cand))
    except Exception as e:
        print(f"Effect.{cand} -> {type(e).__name__}")

# 触发方式：先加一个效果再改 TriggerType
for t in (1, 2, 3, 4):
    try:
        eff.Timing.TriggerType = t
        print("TriggerType", t, "OK ->", eff.Timing.TriggerType)
    except Exception as e:
        print("TriggerType", t, "->", type(e).__name__, e)

pres.Close()
app.Quit()