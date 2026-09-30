# 症状 · 往返后效果被吃掉

> PowerPoint 打开再保存一遍（round-trip），你的效果就变了或没了。注意：往返**既会救活**你写不对的东西，**也会改掉**你写对的东西。
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 往返后效果被吃掉

注入时好好的，**PowerPoint 打开再保存之后就没了** —— 因为 PowerPoint 会按自己的理解重写 XML。判据：`motion.ps1 -Strict` 往返普查。

### §14 同一形状上"入场 + 强调"不兼容 PowerPoint 往返（实测限制，未解决）

这是本机实测的硬限制，**不是本 skill 的 bug**，但必须知道：

| spec | 引擎写出的 preset | PowerPoint `Save()` 之后 |
| --- | --- | --- |
| 单效果 `spin` | 1 个 | **保留** |
| `fly` + `growShrink`（同形状） | 2 个 | 只剩 `fly` |
| `fly` + `spin`（同形状，after / with 都试过） | 2 个 | 只剩 `fly` |
| `fly` + `fade` + `spin`（spin 与 fly 同形状） | 3 个 | 只剩 `fly`、`fade` |

即：**同一形状上叠加"入场 + 强调"时，强调必定丢失**。不同形状之间不受影响。

因此 `motion.ps1` 加了往返普查（effect census）：打开前数文件里的
`<p:cTn presetID=…>`，`SaveCopyAs` 之后再数一次，少了就报 `LOSS`。
`-Strict` 时以非零退出。**注意两个数不能跨口径比较**：PowerPoint 的
`MainSequence.Count` 会把一个入场拆成"可见性 set + 动画"两项，通常大于 preset 数
（自测 deck 是 11 vs 6），只有"文件 vs 文件"的往返比较才有意义。

<p align="right"><sub>来源 com-pitfalls §14</sub></p>

### §18 Round-trip 会以第二种方式咬"一个形状多个效果"

§14 讲的是"入场+强调"被丢掉。还有第二种成因不同、症状相似的情况：

**逐段揭示**（by-paragraph build）在 PowerPoint 里就是同一形状挂多个效果。你写进
spec，`motion.py apply` 老实写成多个 `<p:par>`，`--assert-geometry`、
`verify_motion`、`motion.py check` 全过，然后 `motion.ps1 -Strict` 报 `LOSS`：

```
round-trip census: 153 effect(s) written back, 155 were in the input
```

**实测**：两个逐段块各多一个效果，正好丢 2 个。PowerPoint 重新读时间轴时把重复形状
**折叠成一行**，多出来的效果被丢掉。

**判断方法**：`LOSS` 的数字正好等于 `逐段块数 × (段数-1)`，不是任意数字。
看到这个规律就别去查"入场+强调"了。

**修法是改版面，不是改 spec**：一行一个文本框，各自独立 shape id、各自一个效果，
什么都折叠不了，版面看起来完全一样。见 `reference/authoring-rules.md` §J。

<p align="right"><sub>来源 com-pitfalls §18</sub></p>

### §21 3D 相机（`scene3d`）可以被注入，且角度原样保留

3D 相机**可以注入**，角度也能往返保留。

> 注意：§20 与本节分别在**两台不同机器**上验证过（§20 及
> `morph-and-3d-recipes.md` 是本机的 Office LTSC 2024，`16.0.17928.20148`；
> 本节的原始观察来自上游机器的 build 20228）。**能力判定请以本机实测为准。**

```xml
<a:scene3d>
  <a:camera prst="perspectiveRelaxedModerately">
    <a:rot lat="17400000" lon="0" rev="0"/>   <!-- 1/60000 度；17400000 = 290° -->
  </a:camera>
  <a:lightRig rig="threePt" dir="t"/>
</a:scene3d>
```

放进 `<p:spPr>` 内、**追加在 `</p:spPr>` 之前**。`SaveAs` 往返后角度**一位不差**。

⚠️ **轴不能放错**：`lat` 是俯仰（做"平面翻倒"），`lon` 是水平自转。
上例放 `lat` 才是 template1 那种"躺下"效果；把 290° 写进 `lon` 会得到水平自转，
**结构完全合法、PowerPoint 正常打开，但效果完全不同**。界面角度与 `lat/lon`
的换算见 §23。

**要点**：

- 「三维旋转」调的是**相机**，不是物体转。物体在平面内转是 `<a:xfrm/@rot>`，
  两者完全不同，别混。
- `perspectiveRelaxedModerately` 是 OOXML **62 个标准预设相机之一**（中文界面叫
  「适度宽松」）；同族的还有 `perspectiveRelaxed` / `perspectiveFront` /
  `perspectiveAbove` 等。
- 预设名**不要自己编**。完整的 62 个预设名和它们的实测角度值，LibreOffice 有一份
  整理好的表：`oox/source/drawingml/scene3dhelper.cxx`
  （注释注明是 experimental 实测所得）。
  ⚠️ 但那份表里**角度数值的单位换算我核对不上**（`perspectiveRelaxedModerately`
  的值与其单位注释差约 725 倍，原因未查明），**所以角度以实测为准，不要照抄那份数**。

<p align="right"><sub>来源 com-pitfalls §21</sub></p>

---

---

### §41 core 切换的 Fallback 被 PowerPoint 改掉

16 个 core 切换往返后全被判「被改写」，但子元素一个都没变。原因在 Fallback：
**core 元素本来就向后兼容，PowerPoint 要求 Fallback 重复它自己**，
只有 p14/p15/p159 的子元素才降级成 `<p:fade/>`。

```xml
<mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>   <!-- 写的 -->
<mc:Fallback><p:transition><p:push dir="u"/></p:transition></mc:Fallback>  <!-- 保存后 -->
```

同源两条：**每个切换都会被包进 `mc:AlternateContent`**（连 core 的也是，因为
`p14:dur` 是 2010 年属性）；**值等于默认值的属性会被删掉**（`<p:push dir="l"/>` →
`<p:push/>`）。

完整 48 项对照表见 [`transitions.md`](transitions.md)。

<p align="right"><sub>来源 com-pitfalls §41</sub></p>
