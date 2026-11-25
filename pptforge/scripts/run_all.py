"""一键回归：冒烟 → 验收矩阵 → MCP stdio → 工具联调。

用法：python scripts/run_all.py
需要本机 PowerPoint 的步骤在缺 COM 时会自动降级，不算失败。
"""
import os
import subprocess
import time
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STEPS = [
    ("冒烟：模型 → pptx", ["scripts/smoke_test.py"]),
    ("验收矩阵：15 切换 + 9 动画（含真机回读）", ["scripts/matrix_test.py"]),
    ("形状层：组合/移动/缩放/属性/对齐", ["scripts/test_shapes.py"]),
    ("绘图属性：渐变/效果/富文本/裁剪", ["scripts/test_styles.py"]),
    ("绘图属性单变量二分开箱检查", ["scripts/diag_styles.py"]),
    ("布局层：elements + ops + 组动画", ["scripts/test_layout.py"]),
    ("样式走 deck DSL", ["scripts/test_deck_styles.py"]),
    ("数据层：表格 / 图表 / 图表动画", ["scripts/test_data.py"]),
    ("数据层版式开箱检查", ["scripts/diag_data2.py"]),
    ("数据层 MCP 工具联调", ["scripts/test_data_tools.py"]),
    ("主题与放映设置", ["scripts/test_theme.py"]),
    ("Morph 关键帧序列", ["scripts/test_morph.py"]),
    ("形状层 MCP 工具联调", ["scripts/test_shape_tools.py"]),
    ("MCP stdio 端到端", ["scripts/test_stdio.py"]),
    ("工具联调：catalog / build / layout / render / pdf", ["scripts/test_tools.py"]),
    ("结构校验 + 真机回读 demo", ["scripts/verify_demo.py", os.path.join(ROOT, "samples", "pptforge-demo.pptx")]),
]

results = []
for title, args in STEPS:
    print("\n" + "=" * 78)
    print("[step]", title)
    print("=" * 78)
    # 每步前清掉残留的 PowerPoint 实例，避免 COM 抢占导致偶发失败
    try:
        subprocess.run(["taskkill", "/F", "/IM", "POWERPNT.EXE"],
                       capture_output=True, text=True)
        time.sleep(1.0)
    except Exception:  # noqa: BLE001
        pass
    proc = subprocess.run([sys.executable] + args, cwd=ROOT, text=True,
                          encoding="utf-8", errors="replace")
    results.append((title, proc.returncode))

print("\n" + "=" * 78)
print("[summary]")
for title, code in results:
    print(f"  {'OK ' if code == 0 else 'FAIL'}  {title}  (exit {code})")
failed = [t for t, c in results if c != 0]
print(f"\n{len(results) - len(failed)}/{len(results)} steps passed")
sys.exit(1 if failed else 0)

