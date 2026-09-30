# 症状 · 文件打不开 / 报损坏

> PowerPoint 报 `0x80070570`「文件或目录已损坏」，且**不告诉你是哪个元素**。这一栏全部是**结构非法**，不是内容错。
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 文件打不开 / 报损坏

PowerPoint 报 `0x80070570`「文件或目录已损坏」，而它**不会告诉你是哪个元素**，所以每次都要自己定位。**这一栏全部是结构非法**，不是内容错。有一条通用规则能防掉其中大部分：**一个序列型元素只能出现一次 —— 存在就替换，不存在才插入**（见 §31）。

### §1 多余 preset 包装层 → PowerPoint 拒开（最致命）

PowerPoint 自己写的动画节点长这样：外层 `<p:cTn presetID=… nodeType="afterEffect">` 包着
`<p:set>` / `<p:animEffect>`：

```xml
<p:cTn id="5" presetID="10" presetClass="entr" presetSubtype="0" fill="hold"
       grpId="0" nodeType="afterEffect">
  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
  <p:childTnLst>
    <p:set>…style.visibility…</p:set>
    <p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="7" dur="500"/>…</p:cBhvr></p:animEffect>
  </p:childTnLst>
</p:cTn>
```

**把这个结构原样嵌到骨架 `<p:cTn>` 下，PowerPoint 打开直接 `E_FAIL`**——连从 PowerPoint 自己
文件里抠出来的原封节点也不行（已二分确认，不是模板被改坏）。

必须改成**内联**：效果子节点直接作为骨架 cTn 的子节点，preset 属性合并到骨架 cTn 上。

```xml
<p:par><p:cTn id="4" fill="hold" presetID="10" presetClass="entr" presetSubtype="0"
              grpId="0" nodeType="afterEffect">
  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
  <p:childTnLst>
    <p:set>…</p:set>
    <p:animEffect …>…</p:animEffect>
  </p:childTnLst>
</p:cTn></p:par>
```

`motion.py` 的 `split_effect_template()` 就是做这件事；`build_timing()` 负责合并属性。
实测结论：内联（带/不带 preset 属性、带/不带 `bldLst`）**全部可开**；包装层**全部拒开**。

<p align="right"><sub>来源 com-pitfalls §1</sub></p>

### §2 重复属性 → XML 非法

模板的外层 cTn 自带 `id="5"`、`dur`、`fill`。合并属性到骨架 cTn 时如果不过滤，
就会产生 `id="4" … id="5"` 这种**重复属性**，lxml 直接报
`Attribute id redefined`。必须过滤 `id|dur|fill`。

<p align="right"><sub>来源 com-pitfalls §2</sub></p>

### §26 `a:rot` 的 `lat/lon` 必须是 0~21600000，负数会让**整个文件**损坏

```xml
<a:rot lat="-1800000" .../>     <!-- -30° -->
```

PowerPoint 打开报 **`0x80070570`（文件或目录损坏且无法读取）**——
不是忽略这个元素，是**直接认为整个文件坏了**。

**必须归一化到 0~360°**：`-30°` 要写成 `lat="18900000"`（330°）。

**这是个很容易踩的坑**，因为同一个角度在别处（窗格、COM）**是可以写负数的**，
只有进了 OOXML 的 `lat/lon` 才必须绕回正区间。

```python
def norm(deg):
    v = int(round(deg * 60000)) % 21600000
    return v + 21600000 if v < 0 else v
```

顺带记一下这个错误码：**`0x80070570` 在 OOXML 注入里几乎总是指
"属性值越界或结构非法"**，不是文件真的坏了。遇到它先查数值范围。

<p align="right"><sub>来源 com-pitfalls §26</sub></p>

### §31 「往序列里插一个已经存在的元素」—— 本项目**所有**损坏文件都是这一个原因

`0x80070570`（"文件或目录已损坏"）在这个项目里出现过 **六次**，每次都只修表面症状。
回头数，**六次是同一个错误**：

| # | 重复了什么 | 触发场景 |
| --- | --- | --- |
| 1 | `<p:transition>` ×2 | python-pptx 已经写过切换，我又插一个 |
| 2 | `<a:solidFill>` + `<a:blipFill>` | 填充是 `xsd:choice`，**取第一个**，旧填充没删 |
| 3 | `useBgFill="1"` 加到**文本框**上 | 全局替换命中不该改的形状 |
| 4 | `<a:effectLst>` ×2 | 形状已有空的自闭合 `<a:effectLst/>`，我又插一个带阴影的 |
| 5 | `<p:transition>` ×2（第二次） | 同 1 |
| 6 | `<a:effectLst>` ×2（第二次） | 同 4 |

**OOXML 的 schema 大多是 `xsd:sequence` + 可选成员，所以第二个副本不是"多余的"，
而是让文档非法。** 而 PowerPoint **只报"已损坏"，不告诉你哪个元素有错**，
所以每次都得从零开始猜。

### 规则

**永远不要写 `s.replace("</p:spPr>", X + "</p:spPr>")`。** 用
`motion.set_singleton()` —— 存在就**替换**，不存在才插入。

```python
xml, status = motion.set_singleton(xml, "a:effectLst", new_block,
                                   inside="spPr",          # 限定父元素
                                   before=("a:scene3d", "a:extLst"))  # 插入时的顺序
```

### 加了一道闸

```bash
python scripts/verify_singletons.py --pptx out.pptx
```

它检查形状属性、幻灯片子元素、切换的重复，**并检查填充冲突**
（两个不同名的填充并存也非法，只查同名重复是查不出来的）。

### 造这道闸时又犯了三个错，值得单独记

**判据不能自己骗自己 —— 一个"永远报 OK"的检查比没有检查更糟。**

| 错 | 症状 |
| --- | --- |
| `motion.element_spans` **自己拼 `p:` 前缀**，要传裸名 `"spPr"`。我传 `"p:spPr"` | 拼成 `<p:p:spPr`，**匹配不到任何东西**，于是在损坏文件上报 OK |
| 直接子元素判定写成 `depth == 1` | 但 spPr 的直接子元素在 **depth 0**，同样永远匹配不到 |
| 用全局正则数 `<p:transition` | **把正确的 morph 当成重复** —— morph 本来就有两个 `transition`（`mc:Choice` 一个、`mc:Fallback` 一个） |

前两个都属于同一类：**检查器悄悄匹配不到，然后报"没问题"。**
所以这个工具的回归测试里，**专门断言"带前缀的名字匹配不到任何东西"** ——
把陷阱本身写成测试。

**另外 `element_spans` 只认成对标签** `<x>…</x>`，对自闭合 `<x/>` 返回空。
于是 `set_singleton` 替换了自闭合那份，却看不见旁边的成对副本，重复照样留着。
现在用 `singleton_spans()` 一次找两种形式。

<p align="right"><sub>来源 com-pitfalls §31</sub></p>

---

### §40 切换候选合并成一份 deck → 拒开且无法归因

把 48 个切换候选合并到一份 deck 里探，PowerPoint 报 `could not open the file`，
**一个结论都拿不到** —— 一个非法子元素就足以拒开整份（同 §26），而"开不了"
不会告诉你是哪一个坏的。

**一个候选一份 deck**，"这份被拒开"才是可归因的结论。

**症状的粒度本身就是信息**：一个拒开 = 那个候选有问题；**全部**拒开 =
生成它们的那份代码有问题（第一次探测 48 个全拒，根因是我们自己把 `xmlns:a`
在 `<p:sld>` 上声明了两次 —— python-pptx 已经声明过 a/p/r）。

<p align="right"><sub>来源 com-pitfalls §40</sub></p>

### §42 切换 / 翻转 / 库 / 摩天轮 / 传送带：裸写会被拒开整份文件

`<p14:switch/>`、`<p14:flip/>`、`<p14:gallery/>`、`<p14:ferris/>`、`<p14:conveyor/>`
五个裸写全部被拒开。**这五个必须带 `dir`**，PowerPoint 自己永远写 `dir="l"`。

**"拒开"只说明这个写法不行，说不出正确的写法。** 猜名字到此为止 ——
改用枚举扫描让 PowerPoint 自己写一遍（同 §23 的做法），答案就出来了。

<p align="right"><sub>来源 com-pitfalls §42</sub></p>

### §43.2 一个 `<p:transition>` 里塞两个子元素 → 拒开整份文件

`<p:push/>` + `<p159:morph/>`、`<p:push/>` + `<p:wipe/>`，两组都拒开。
**不是丢一个、不是留第一个，是整份打不开** —— 与 §42 同一类失败。

于是「整页怎么过去」只有一个名额：**形状形变与整页位移不能同时表达**，
页内的个体行为必须走 `<p:timing>`。详见
[`transition-model.md`](transition-model.md) 二。

<p align="right"><sub>来源 com-pitfalls §43.2</sub></p>

---
