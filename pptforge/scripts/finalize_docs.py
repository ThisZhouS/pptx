"""把全量基准写进 SKILL.md 对标表，并重新生成 examples。"""
import io
import os
import subprocess

SK = r"E:\pptx\pptx-skill\SKILL.md"
t = io.open(SK, encoding="utf-8").read()

start = t.find("## 对标指标")
end = t.find("## 结构")
assert start > 0 and end > start

NEW = '''## 对标指标（`scripts/verify_benchmark.py`，基准为全量 2120 个模板）

同一份大纲（CodeCraft 内容，6 节）跑两次：`full` 生成章节页，`compact` 不生成。

| 指标 | 全量基准 | full（15 页） | compact（9 页） |
| --- | --- | --- | --- |
| 前景形状/页 | **10**（16157 页中位） | 8 | **11** |
| 装饰/页 | **3**（中位）／9（均值） | 6 | 7 |
| 内容覆盖率 | 0.714（内容页口径） | 0.546 | **0.715** |
| 字号档数 | 11 | 11 | 9 |
| 最大/正文 | 5.0 | 6.8 | 3.8 |
| 上边距 | 0.021 | 0.049 | 0.045 |
| 动画/页 | 1.8（全量中位）｜科技档 8–15 | 9.3 | 13.4 |
| `layout_check` 溢出 | — | 0 | 0 |

**覆盖率只受章节页占比影响**：章节页只有编号与标题，天然空。
`--compact` 跳过章节页后覆盖率与形状密度立刻落到模板口径；需要章节过渡时用 `full`。

动画数按风格档位而非全量中位：全量中位 1.8 条/页（37% 的模板零动画），
`科技` 档刻意给到 8–15 条/页 —— 那是 `high` 档模板的水平。

```bash
python scripts/build.py outline.json --out deck.pptx --task 路演 --compact
```

'''
t = t[:start] + NEW + t[end:]
io.open(SK, "w", encoding="utf-8").write(t)
print("[ok] SKILL.md 对标表 → 全量基准")

# 重新生成 examples
os.chdir(r"E:\pptx\pptx-skill")
for style in ("科技", "商务", "极简"):
    out = os.path.join("examples", f"demo-{style}.pptx")
    r = subprocess.run(["python", "scripts/build.py", "--demo", "--out", out, "--style", style],
                       capture_output=True, text=True, encoding="utf-8")
    line = [x for x in r.stdout.splitlines() if '"accent"' in x or '"slides"' in x]
    print(f"  {style}: {' '.join(line[:2])}")