# PowerPoint 结构逆向报告

对象：PowerPoint 16.0（`C:\Program Files\Microsoft Office\Root\Office16`，build 20625）
目标：把"幻灯片切换"与"对象动画"的 OOXML 写法完整、可复现地拆出来，供程序生成。
结论产物：[`../schema/ooxml_structure.json`](../schema/ooxml_structure.json)、
[`../schema/transitions.json`](../schema/transitions.json)、
[`../schema/animations.json`](../schema/animations.json)。

---

## 1. 方法与顺序

逆向三件套不是读文档，而是**让 PowerPoint 自己说话**：

| 步骤 | 手段 | 脚本 | 产出 |
| --- | --- | --- | --- |
| ① 枚举名字空间 | 解析 COM 类型库 `MSPPT.OLB`，取 `PpEntryEffect`(192) / `MsoAnimEffect`(155) / 速度、触发、方向等枚举 | `rev/discover_typelib.py` | `rev/out/typelib_constants.json` |
| ② 真机采样切换 | 对每个枚举值设置 `SlideShowTransition.EntryEffect`，保存，解包 `slideN.xml` 抽 `<p:transition>` | `rev/acquire_transitions.py` | `rev/out/gt_transitions.pptx`(197 页)、`transitions_map.json` |
| ③ 真机采样切换属性 | 单变量试 `Duration` / `Speed` / `AdvanceOnTime` / `AdvanceOnClick` | `rev/acquire_transition_props.py` | `gt_transition_props.pptx`(12 页)、`transition_props.json` |
| ④ 真机采样动画 | 对每个动画值 `MainSequence.AddEffect(...)`，抽 `<p:timing>`；再对同一批做 `Exit=True` 采退出类 | `rev/acquire_animations.py`、`rev/acquire_anim_exit.py` | `gt_animations.pptx`(154 页)、`gt_anim_exit.pptx`(150 页) |
| ⑤ 参数方言采样 | 触发/时长/延迟/重复/平滑/文字分级/多效果/路径 逐项单变量 | `rev/acquire_anim_variants.py` | `gt_anim_variants.pptx`(21 例)、`anim_variants.json` |
| ⑥ 形状层采样 | 组合/旋转/翻转/透明度/线宽/渐变；组整体移动与缩放；组动画与组内子动画的 target | `rev/acquire_shapes.py` | `gt_shapes.pptx`(7 页)、`shapes_map.json` |
| ⑦ 绘图属性采样 | 填充/线条/效果/文本/图片/表格 逐项单变量 | `rev/acquire_styles.py` | `gt_styles.pptx`(6 页)、`styles_map.json` |
| ⑧ 数据层采样 | 表格合并/边框/列宽行高；图表部件与图表动画组合探查 | `rev/acquire_table_chart.py`、`rev/probe_chart_anim.py` | `gt_table_chart.pptx`(3 页)、`gt_table_border.pptx`、`table_chart_map.json` |
| ⑨ 主题与放映采样 | theme1.xml 配色/字体；presProps.xml 的 showPr 与展示方式 | `rev/acquire_theme.py` | `gt_theme.pptx`、`theme_map.json` |
| ⑩ 编译目录 | 把"枚举 → 元素/属性/preset/行为节点模板"编译成 JSON | `rev/build_catalog.py` | `schema/*.json` |
| ⑪ 闭环验证 | 用逆向结果生成新文件，让 PowerPoint 读回 `EntryEffect`/`EffectType`/`TriggerType` 对账 | `scripts/matrix_test.py` | 24/24 一致 |

采样总量：**554 页**（197 + 12 + 154 + 150 + 21 + 7 + 6 + 4 + 3），全部由真实 PowerPoint 写出，未手写 XML。

---

## 2. 切换层（`<p:transition>`）

### 2.1 位置与元素形态

`p:sld` 的子元素顺序固定：`p:cSld → p:clrMapOvr → p:transition → p:timing → p:extLst`。
顺序错了 PowerPoint 会提示修复，因此写入时统一"插在 `clrMapOvr` 之后、`timing` 之前"。

效果不是属性，而是 `p:transition` 的**子元素**：

```xml
<p:transition spd="med"><p:fade thruBlk="1"/></p:transition>
```

- `ppEffectFade`(1793) → `<p:fade thruBlk="1"/>`（经黑场）
- `ppEffectFadeSmoothly`(3849) → `<p:fade/>`（平滑）
- `ppEffectCoverLeft`(1281) → `<p:cover/>`（无 `dir` 即默认左），`ppEffectCoverUp` → `<p:cover dir="u"/>`
- `ppEffectUncoverLeft` → `<p:pull/>`（PPT 的"揭开"落成 `pull`）
- `ppEffectBoxOut/In` → `<p:zoom/>`、`<p:zoom dir="in"/>`；`ppEffectSplitHorizontalOut` → `<p:split/>`（`out`/`horz` 是默认，省略）
- `ppEffectCubeLeft` → `<p14:prism dir="l"/>`；`ppEffectOrbit*` → `<p14:prism isContent="1" .../>`；
  `ppEffectRotateLeft` → `<p14:prism isInverted="1" .../>`（同一元素不同属性区分三种视觉）
- `ppEffectGlitterHexagonLeft` → `<p14:glitter dir="l" pattern="hexagon"/>`
- `ppEffectFallOverLeft` → `<p15:prstTrans prst="fallOver" invX="1"/>`（`invX=1` 是左向）
- `ppEffectMorphByObject` → `<p159:morph option="byObject"/>`

### 2.2 扩展命名空间必须走 `mc:AlternateContent`

`p14` / `p15` / `p159` 元素不能裸写，必须给出降级路径：

```xml
<mc:AlternateContent xmlns:mc="...markup-compatibility/2006">
  <mc:Choice xmlns:p14="...office/powerpoint/2010/main" Requires="p14">
    <p:transition spd="med"><p14:conveyor dir="l"/></p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="med"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

证据：`rev/out/transitions_raw/slide_117.xml`（conveyor）。

### 2.3 精确时长 `p14:dur` 是"扩展属性"，同样受这条规则约束

单变量实验（`rev/out/transition_props.json`）：

| 设置 | PowerPoint 写出 |
| --- | --- |
| `Speed=2` | `<p:transition spd="med">` |
| `Duration=2.5` | 整页包 `mc:AlternateContent/Choice(Requires="p14")`；`<p:transition spd="slow" p14:dur="2500">`，Fallback 去掉 `p14:dur` |
| `Duration=0.3, Speed=3` | `<p:transition>`（无任何属性 —— `fast` 与默认时长都被省略） |
| `Speed=3` | `<p:transition>`（`fast` 是默认值） |
| `AdvanceOnTime=True, AdvanceTime=3.0` | `<p:transition ... advTm="3000">` |
| `AdvanceOnClick=0` | `<p:transition ... advClick="0">` |
| 无效果 + 自动换片 | `<p:transition spd="slow" p14:dur="2000" advTm="2000"/>`（空 transition 合法） |

### 2.4 一个 COM 层的坑

`Duration` 与 `Speed` 互相覆盖，**后写者赢**：先 `Duration=1.5` 再 `Speed=2`，结果是
`<p:transition spd="med">`，`p14:dur` 消失（`push_dur15`）。所以生成端二选一：要精确就 `duration`，
要三档就 `speed`。

---

## 3. 动画层（`<p:timing>`）

### 3.1 骨架

```
p:timing
└ p:tnLst / p:par / p:cTn(id=1, nodeType=tmRoot, dur=indefinite, restart=never)
  └ p:childTnLst / p:seq(concurrent=1, nextAc=seek)
    ├ p:cTn(id=2, nodeType=mainSeq, dur=indefinite)
    │ └ p:childTnLst
    │   └ p:par                      ← click 组
    │     └ p:cTn(id=3, fill=hold) + stCondLst(cond delay=indefinite)
    │       └ p:childTnLst / p:par   ← 效果容器
    │         └ p:cTn(id=4, fill=hold) + stCondLst(cond delay=0)
    │           └ p:childTnLst / p:par  ← 一个效果
    │             └ p:cTn(id=5, presetID, presetClass, presetSubtype, grpId, nodeType, ...)
    │               └ p:childTnLst
    │                 ├ p:set / p:animEffect / p:animMotion / p:anim  ← 行为节点
    ├ p:prevCondLst(cond evt=onPrev + sldTgt)
    └ p:nextCondLst(cond evt=onNext + sldTgt)
p:bldLst / p:bldP(spid, grpId, animBg|build)
```

### 3.2 触发方式 = `nodeType` + 组级 `cond`

| 用户语义 | 效果 `cTn@nodeType` | 组级 `stCondLst` |
| --- | --- | --- |
| 单击时 | `clickEffect` | `<p:cond delay="indefinite"/>` |
| 与上一动画同时 | `withEffect` | `<p:cond delay="indefinite"/>` + `<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond>` |
| 上一动画之后 | `afterEffect` | 同上（两条 cond 是"或"关系：既能点击触发，也能随主序列开始） |

效果自身永远只有一条 `<p:cond delay="{毫秒}"/>`，`evt` 不复现在效果层。
证据：`anim_variants.json` 的 `entr_click` / `entr_with_prev` / `entr_after_prev`。

### 3.3 分组

`three_mixed`（click + with + after）只生成 **1 个 click 组**，三个效果并列在同一个
"效果容器" `par` 下，`nodeType` 依次 `clickEffect / withEffect / afterEffect`。
所以分组规则是：`click` 开新组，`with`/`after` 追加当前组。

### 3.4 参数落点

| 参数 | 落点 | 示例 |
| --- | --- | --- |
| 时长 | 行为节点 `cTn@dur`（毫秒） | `<p:cTn id="7" dur="500"/>` |
| 延迟 | 效果 `stCondLst/cond@delay` | `<p:cond delay="300"/>` |
| 重复 | 效果 `cTn@repeatCount`（次×1000） | `repeatCount="2000"`（2 次） |
| 平滑 | 效果 `cTn@accel` / `@decel`（%×1000） | 路径动画默认 `accel="50000" decel="50000"` |
| 退出 | 行为节点 `transition="out"`，并在末尾追加 `p:set`（`delay="999"` 把对象在动画后藏起来） | `exit_fade` |
| 播放后 | `Effect.AfterEffect` 在 COM 上不存在（4 个候选属性全部 `AttributeError`），当前不支持 | — |

注意 `<p:set>` 内部 `cTn@dur="1"` 是瞬时标记，**不能**被时长替换；生成器只改 `dur != "1"` 的节点。

### 3.5 文字分级（按段落逐条出现）

`AddEffect(shape, effect, Level=msoAnimateTextByFirstLevel)` 的结果：
3 段文字 → **3 个 click 组**，每组一个 `clickEffect`，目标里带段落范围，`bldLst` 标记 `build="p"`：

```xml
<p:spTgt spid="2"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>
...
<p:bldP spid="2" grpId="0" build="p"/>
```

`Level=msoAnimateTextByAllLevels` 则生成 1 个 click 组 + 2 个 `withEffect`，标记 `build="allAtOnce"`。

### 3.6 动作路径

```xml
<p:animMotion origin="layout" pathEditMode="relative" ptsTypes=""
  path="M 0 0 C 0.069 0 0.125 0.056 0.125 0.125 C ... Z"/>
```

`path` 用相对坐标（0..1 相对形状尺寸），`origin="layout"`，`presetClass="path"`，
并且 PowerPoint 默认给 `accel="50000" decel="50000"`。

### 3.7 id 唯一性

`p:cTn@id` 在**同一页**内必须唯一，`p:rtn` 同理。采样里一次 1 效果的页用掉 id 1..7，
多效果页线性增长。生成端用一个自增分配器，整棵树（含行为节点内部）统一重编号。

---

## 4. 形状层（`p:grpSp`）

### 4.1 几何

位置与尺寸都在 `a:xfrm`：`a:off@x, @y` + `a:ext@cx, @cy`，单位 EMU（914400 = 1 英寸）。
旋转 `@rot` 以 1/60000 度计（`Rotation=30` → `rot="1800000"`），翻转是 `@flipH` / `@flipV="1"`。
透明度落在纯色填充上：`<a:srgbClr val="CC8844"><a:alpha val="50000"/></a:srgbClr>`（0.5 → 50000）。
线宽 `a:ln@w` 同样是 EMU（2pt → `w="19050"`），渐变填充是 `a:gradFill`。

### 4.2 组合结构

```xml
<p:grpSp>
  <p:nvGrpSpPr>
    <p:cNvPr id="5" name="组合 4"/>
    <p:cNvGrpSpPr/><p:nvPr/>
  </p:nvGrpSpPr>
  <p:grpSpPr>
    <a:xfrm>
      <a:off x="1270000" y="1270000"/><a:ext cx="5207000" cy="3683000"/>
      <a:chOff x="1270000" y="1270000"/><a:chExt cx="5207000" cy="3683000"/>
    </a:xfrm>
  </p:grpSpPr>
  <p:sp>…</p:sp>          <!-- 子形状直接挂在 grpSp 下，没有 p:spTree 包裹 -->
  <p:sp>…</p:sp>
</p:grpSp>
```

真机三形状组合的实测：`off == chOff`、`ext == chExt`，**子形状的 off/ext 完全不变**。
也就是 PowerPoint 组合时把子坐标系原点设在包围盒左上，比例 1:1。

### 4.3 移动与缩放的"看不见的手"

| 动作 | XML 变化 | 子形状 |
| --- | --- | --- |
| 组 `Left += 120, Top += 40` | `off` 从 (1270000,1270000) → (2794000,1778000)，`chOff` 保持 (1270000,1270000) | 不变 |
| 组 `Width ×= 0.6, Height ×= 0.6` | `ext` 5334000×2540000 → 3200400×1524000，`chExt` 保持 5334000×2540000 | 不变 |

所以视觉位置 = `off + (child_off − chOff) × (ext / chExt)`。
实现"移动组"只改 `off`，"缩放组"只改 `ext`，比逐个改子形状安全得多；
`ungroup` 时按同一公式反算，并把组的 `rot`/`flip` 折算进每个子形状。

### 4.4 组动画的 target

组加动画后 timing 里仍是 `<p:spTgt spid="4"/>`（组 id），**没有** `p:grpSpTgt`；
组内单个子形状加动画则写子形状自己的 `spid`。生成端只要把 target 解析成组 id 即可。

## 5. 绘图属性（填充 / 线条 / 效果 / 文本 / 图片）

`rev/acquire_styles.py` 生成 6 页真机样本，逐属性比对。所有单位都是"整千/整万"的定点数。

### 5.1 填充

```xml
<!-- 纯色 + 透明度 -->
<a:solidFill><a:srgbClr val="CC6633"><a:alpha val="50000"/></a:srgbClr></a:solidFill>
<!-- 线性渐变（TwoColorGradient + GradientAngle=45） -->
<a:gradFill flip="none" rotWithShape="1">
  <a:gsLst><a:gs pos="0"><a:srgbClr val="CC6633"/></a:gs><a:gs pos="100000"><a:srgbClr val="FFCC66"/></a:gs></a:gsLst>
  <a:lin ang="2700000" scaled="0"/>
</a:gradFill>
<!-- 径向渐变 -->
… <a:path path="rect"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>
<!-- 图案 / 无填充 -->
<a:pattFill prst="pct30"><a:fgClr>…</a:fgClr><a:bgClr>…</a:bgClr></a:pattFill>
<a:noFill/>
```

角度 `ang` 单位是 1/60000 度（45° → 2700000，因为 DrawingML 的角度以 x 轴正向为 0 顺时针）。

### 5.2 线条

```xml
<a:ln w="38100" cap="rnd">
  <a:solidFill><a:srgbClr val="6FE3C4"><a:alpha val="80000"/></a:srgbClr></a:solidFill>
  <a:prstDash val="dash"/>
  <a:headEnd type="triangle"/><a:tailEnd type="arrow"/>
</a:ln>
```

`w` 是 EMU（3pt → 38100）；`cap` 取 `flat|rnd|sq`；虚线取 `prstDash@val`（`dash`/`dot`/`lgDash`…）；
箭头取 `headEnd/tailEnd@type`（`triangle`/`stealth`/`diamond`/`oval`/`arrow`）。

### 5.3 效果

`a:effectLst` 的子元素顺序是固定的：`blur → fillOverlay → glow → innerShdw → outerShdw → prstShdw → reflection → softEdge`。顺序错了 PowerPoint 会提示修复。

| COM 设置 | XML |
| --- | --- |
| `Shadow.Visible=True, Blur=12, OffsetX/Y=6, Transparency=0.35` | `<a:outerShdw blurRad="152400" dist="107763" dir="2699994" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="65000"/></a:srgbClr></a:outerShdw>` |
| `Glow.Radius=12, Transparency=0.3` | `<a:glow rad="152400"><a:srgbClr val="FFCC66"><a:alpha val="70000"/></a:srgbClr></a:glow>` |
| `SoftEdge.Radius=10` | `<a:softEdge rad="127000"/>` |

注意 `blurRad/rad = 磅×12700`，`dist` 是偏移合成距离（6,6 → 107763），`dir` 是 1/60000 度，
`Transparency` 要换算成 `1-x` 的 alpha。

### 5.4 文本与段落

```xml
<a:pPr marL="914400" lvl="1" indent="-457200" algn="ctr">
  <a:lnSpc><a:spcPct val="140000"/></a:lnSpc>
  <a:spcBef><a:spcPts val="1200"/></a:spcBef>
  <a:spcAft><a:spcPts val="800"/></a:spcAft>
  <a:buSzPct val="110000"/><a:buFont typeface="Arial"/><a:buChar char="•"/>
</a:pPr>
<a:r><a:rPr lang="zh-CN" altLang="en-US" sz="3400" b="1" i="1" u="sng">
  <a:solidFill>…</a:solidFill><a:latin typeface="Arial"/><a:ea typeface="Microsoft YaHei"/>
</a:rPr><a:t>加粗 </a:t></a:r>
```

- `sz` = 1/100 pt（34pt → 3400）
- 行距用 `spcPct`（140% → 140000），段前段后用 `spcPts`（12pt → 1200）
- 缩进 `marL` / `indent` 是 EMU，`lvl` 从 0 开始（PowerPoint UI 的"1 级"= `lvl="0"`）
- 下划线 `u="sng"`、删除线 `strike="sngStrike"`、上下标 `baseline="±30000"`
- 中英字体分开：`a:latin` + `a:ea`（漏掉 `a:ea` 中文会回退到主题字体）
- `Emboss` 属性会生成 `rPr/a:effectLst/a:innerShdw`（不是单独标签）

### 5.5 图片

```xml
<p:blipFill><a:blip r:embed="rId2"><a:lum bright="10000" contrast="15000"/></a:blip>
  <a:srcRect l="10000" t="8000" r="5000"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
```

两个坑：

1. `p:blipFill` 是 **p 命名空间**（不是 `a:blipFill`），`srcRect` 才是 `a:`；
2. `srcRect` 的值是比例 ×100000，而 COM 的 `PictureFormat.CropLeft` 单位是**磅**（相对图片原始尺寸）——
   `CropLeft=0.1` 作用在 480pt 宽的图上 → `srcRect l="10000"`（21/100000 是四舍五入后的显示差异）。

### 5.6 表格

`<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5940675A-…}</a:tableStyleId></a:tblPr>`；
单元格在 `p:graphicFrame/a:graphic/a:graphicData/a:tbl/a:tr/a:tc`，填充写 `tcPr`。

### 5.7 COM 的 BGR 陷阱

`ForeColor.RGB = 0x3366CC` 在 XML 里是 `srgbClr val="CC6633"` —— COM 用 `0xBBGGRR`，
DrawingML 用 `#RRGGBB`。只做生成时（python-pptx 的 `RGBColor`）不受影响，
但拿 COM 采样值去对 XML 时必须翻转，否则会误判"颜色写错了"。

## 6. 表格与图表

### 6.1 表格

```xml
<a:tbl>
  <a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5940675A-…}</a:tableStyleId></a:tblPr>
  <a:tblGrid><a:gridCol w="1778000"/>…</a:tblGrid>
  <a:tr h="846667">
    <a:tc><a:txBody>…</a:txBody><a:tcPr/></a:tc>
    <a:tc hMerge="1"><a:txBody>…</a:txBody><a:tcPr/></a:tc>
    <a:tc><a:txBody>…</a:txBody>
      <a:tcPr>
        <a:lnL w="19050" cmpd="sng"><a:solidFill><a:srgbClr val="457AFF"/></a:solidFill><a:prstDash val="solid"/></a:lnL>
        <a:lnR …/><a:lnT …/><a:lnB …/><a:lnTlToBr …/><a:lnBlToTr …/>
      </a:tcPr>
    </a:tc>
  </a:tr>
</a:tbl>
```

- 合并：主单元格写 `gridSpan="2"`，被吞掉的单元格保留 `a:tc` 并写 `hMerge="1"`（纵向是 `rowSpan` / `vMerge`）。
- `a:tcPr` 的子元素顺序：`lnL → lnR → lnT → lnB → lnTlToBr → lnBlToTr → 填充 → extLst`。
- 列宽 `a:gridCol@w`、行高 `a:tr@h` 都是 EMU；表样式是 `a:tblPr/a:tableStyleId{GUID}`。
- 边框实测：`Cell(2,2).Borders.Item(ppBorderBottom).Weight=1.5` → `<a:lnB w="19050" cmpd="sng">`。

### 6.2 图表

图表不是形状，而是 `p:graphicFrame` 指向独立部件：

```
ppt/charts/chart1.xml                       c:chartSpace / c:chart / c:plotArea / c:barChart…
ppt/charts/_rels/chart1.xml.rels
ppt/embeddings/Microsoft_Excel_Sheet1.xlsx  数据源（"编辑数据"打开的就是它）
```

所以图表必须带内嵌工作簿，否则 PowerPoint 打开会报错或降级成图片。
生成端用 python-pptx 的 `add_chart`（自动写 xlsx），只需补配色、图例、数据标签、轴格式。

### 6.3 图表动画：一个 COM 死路

| 尝试 | 结果 |
| --- | --- |
| `AddEffect(chart, msoAnimEffectFade, msoAnimateLevelNone, click)` | OK（整体动画） |
| `AddEffect(chart, msoAnimEffectFade, msoAnimateChartBySeries=10, click)` | `Invalid request` |
| `AddEffect(chart, 21/2/1, 10)` 各效果组合 | 全部 `Invalid request` |
| `AddEffect(chart, 10, msoAnimateChartByCategory=8 / AllAtOnce=7)` | 全部 `Invalid request` |

结论：这一版 PowerPoint 的 COM 不接受图表分级参数，`Sequence.AddEffect` 只给整体效果。
"按系列出现"只能在 UI 里做，或手写 `p:bldGraphic/bldSub/a:bldChart` —— 后者无真机样本可比对，
不放进目录，避免产出 PowerPoint 会拒绝的文件。

## 7. 主题与放映设置

### 7.1 theme1.xml

```xml
<a:theme><a:themeElements>
  <a:clrScheme name="Office">
    <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
    <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
    <a:dk2><a:srgbClr val="1F497D"/></a:dk2> … <a:accent1>…<a:accent6>
    <a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
    <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>
      <a:font script="Hans" typeface="宋体"/>…</a:majorFont>
    <a:minorFont>…</a:minorFont>
  </a:fontScheme>
  <a:fmtScheme>…</a:fmtScheme>
</a:themeElements></a:theme>
```

- 12 个配色键里，只有 `dk1`/`lt1` 是 `sysClr`（`val="windowText"/"window"` + `lastClr`），其余 10 个是 `srgbClr@val`。
- 改主题只需改这些值，`fmtScheme` 与 `fontScheme` 的 script 列表不用动（结构合法）。
- 中文走两层：`<a:ea typeface>` 与 `<a:font script="Hans">` —— 只改 `latin` 不动中文会"看起来没生效"。
- COM 侧 `Theme.ThemeColorScheme.Colors(n)` 的索引是：1=dk1 2=lt1 3=dk2 4=lt2 **5=accent1** 6=accent2 …，
  且 `RGB` 仍是 BGR（设 `0xFF7A45` → XML `val="457AFF"`）。

### 7.2 presProps.xml（放映设置不在 presentation.xml）

```xml
<p:presentationPr xmlns:a="…" xmlns:r="…" xmlns:p="…">
  <p:showPr loop="1" showNarration="0">
    <p:kiosk/><p:sldAll/>
    <p:penClr><a:prstClr val="red"/></p:penClr>
  </p:showPr>
</p:presentationPr>
```

| COM 设置 | XML |
| --- | --- |
| `SlideShowSettings.LoopUntilStopped = msoTrue` | `showPr@loop="1"` |
| `ShowWithNarration = 0` | `showPr@showNarration="0"`（默认就是关，省不省看版本） |
| `AdvanceMode = ppSlideShowUseSlideTimings` | 默认，`useTimings` 省略 |
| `ShowType = ppShowTypeKiosk(3)` | `<p:kiosk/>` + `<p:sldAll/>`（演讲者是 `<p:present/>`，浏览是 `<p:browse/>`） |

实测：写入 `loop + kiosk` 后 PowerPoint 回读 `LoopUntilStopped=-1`、`ShowType=3`，说明手写 `presProps.xml` 被完全接受。
注意这个部件在 python-pptx 默认模板里**不存在**，需要新建 Part 并挂 `presentation → presProps` 关系。

## 8. 目录编译

`rev/build_catalog.py` 把采样结果编译成两张表：

- `schema/transitions.json`：150 条（`p` 58 / `p14` 68 / `p15` 20 / `p159` 3 / `none` 1），
  每条含 `element`、`element_attrs`、`requires`、`needs_alternate_content`、中文名，
  以及同一 XML 的重复枚举别名（7 组进了 `also`）。
- `schema/animations.json`：150 条动画预设，每条含 `preset{presetID,presetClass,presetSubtype}`、
  `behaviors[]`（XML 片段，`spid` 用 `{{SPID}}` 占位）、退出变体（150 条里 53 条能转退出）。
  中文名按"族 + 方向"规则生成，覆盖 PowerPoint 中文 UI 的叫法。

编译时踩到的两个命名陷阱，已固化进规则：

1. `ppEffectStripsUpLeft` 与 `ppEffectStripsLeftUp` 是同一 XML，去重后进 `also`，只保留一条。
2. `ppEffectGlitterDiamond*` 与 `ppEffectGlitterHexagon*` 共享 `<p14:glitter>`，靠 `pattern` 区分，
   所以 id 必须来自枚举名而不是元素名（否则两者会撞车）。

---

## 9. 闭环验证

`scripts/matrix_test.py` 生成 24 页（15 种切换 + 9 种动画场景），用 PowerPoint COM 打开逐项对账：

- 切换比对 `SlideShowTransition.EntryEffect` 与目录里的 `pp_value`：15/15 一致
  （含 `p14:vortex`=3863、`p15:prstTrans fallOver`=3934、`p159:morph`=3954）。
- 动画比对 `MainSequence.Count`、`EffectType`、`Timing.TriggerType`、`Effect.Exit`：一致。
- `validate` 对生成件的结构检查：0 problem。
- `layout_check`（PowerPoint `BoundHeight`）：0 溢出。

`scripts/test_stdio.py` 另跑一遍完整 MCP 协议：`initialize → tools/list(18) → tools/call`，
生成的 pptx 结构校验 0 problem。

形状层另有两条闭环（`scripts/test_shapes.py` / `scripts/test_layout.py`）：

- 3 形状组合后 PowerPoint 报 `GroupItems = 3`，`Shape.Type = 6`（msoGroup）；
- 组移动后子形状 XML 坐标不变、`chOff` 不变、`off` 改变；组缩放 0.6 倍后 `ext` 变而 `chExt` 不变；
- 组动画读回为 `(EffectType=23, TriggerType=3, shape_name='pair')`，组内子形状动画读回子形状名；
- 解组后几何按 `off + (child_off − chOff) × scale` 落位，PowerPoint 显示的英寸数与换算值一致。

---

## 10. 未覆盖

| 项 | 原因 | 当前行为 |
| --- | --- | --- |
| 媒体类动画（`MediaPlay/Pause/Stop/PlayFromBookmark`） | `AddEffect` 对普通形状返回 `Invalid request`（需要媒体形状） | 目录无条目，调用报"未知动画效果" |
| 图表/图示分级（`msoAnimateChart*` / `Diagram*`） | 需要真实图表/图示对象 | 同上 |
| 单击对象触发（`InteractiveSequences`，`TriggerType=4`） | COM 要求形状先有交互动作，`TriggerType=4` 直接 `Invalid request` | 仅支持 click / with / after |
| 切换音效（`SoundEffect`） | 需要音频文件与包内关系 | 未纳入 DSL |
| "播放后变暗/隐藏"（`MsoAnimAfterEffect`） | `Effect` 对象无 `AfterEffect` 属性 | 未纳入 DSL |
| 嵌套组的旋转折算 | `ungroup` 只展开一层；多层嵌套需连续调用，且每层的 `rot`/`flip` 各自折算 | 每次调用展开一层 |
| 图片裁剪 / 蒙版 | 需要 `a:srcRect` 与效果链 | 仅支持整图缩放定位 |
| 自定义几何（`a:custGeom`）与自由曲线 | 未采样 | 只提供预设几何（`SHAPE_KINDS`） |
| 组内时序（`p:grpSpTgt` 子目标） | PowerPoint 2024 对组动画直接写组的 `spTgt`，不用 `grpSpTgt` | 与 PowerPoint 行为一致 |
| 图表分级动画（`p:bldGraphic/bldSub`） | COM 拒绝该参数，没有真机样本可对照 | 只做整体动画 |

这些都不影响主体能力：150 种切换、150 种动画预设、三种触发、按段落分级、动作路径、退出动画已闭环。

---

## 11. 复现

```bash
cd E:\pptx\pptforge

python rev/discover_typelib.py          # ① 枚举名字空间
python rev/acquire_transitions.py       # ② 197 页切换采样
python rev/acquire_transition_props.py  # ③ 12 页属性单变量
python rev/acquire_animations.py        # ④ 154 页动画采样
python rev/acquire_anim_exit.py         # ④b 150 页退出变体
python rev/acquire_anim_variants.py     # ⑤ 21 例参数方言
python rev/acquire_shapes.py            # ⑥ 7 页形状/组合/组动画
python rev/acquire_styles.py            # ⑥b 6 页绘图属性
python rev/build_catalog.py             # ⑦ 编译目录
python rev/acquire_table_chart.py       # ⑧ 3 页表格/图表 + 1 页边框
python rev/acquire_theme.py             # ⑨ 主题部件 + 放映设置
python rev/build_structure_map.py        # ⑨b 结构地图
python scripts/matrix_test.py           # ⑩ 闭环验证
python scripts/test_data.py             # 表格/图表闭环
python scripts/test_theme.py            # 主题/放映闭环
python scripts/run_all.py               # 全量回归（8 步）
```

采样证据都在 `rev/out/`（`*_raw/slide_NNN.xml` 是逐页原始 XML，`*_map.json` 是映射表）。


