# 踩坑 · 文件打不开 / 报损坏

> 文件打不开 / 报损坏 —— **结构非法**的全部成因
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

**一条规则能防掉这里大部分**：序列型元素只能出现一次 —— 存在就替换，不存在才插入（§31）。

---

## 1. 多余 preset 包装层 → PowerPoint 拒开（最致命）

**现象**：打开直接报「需要修复」，`E_FAIL`，连从 PowerPoint 自己文件里抠出来的原封节点也不行

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
## 2. 重复属性 → XML 非法

**现象**：lxml 报 `Attribute id redefined`，XML 非法

模板的外层 cTn 自带 `id="5"`、`dur`、`fill`。合并属性到骨架 cTn 时如果不过滤，
就会产生 `id="4" … id="5"` 这种**重复属性**，lxml 直接报
`Attribute id redefined`。必须过滤 `id|dur|fill`。
## 4. 结构校验通过 ≠ PowerPoint 能打开

**现象**：lxml 结构校验全绿，PowerPoint 打开却 `E_FAIL`

lxml 只保证 well-formed。实例：v3（自建 timing）结构校验 OK，PowerPoint `E_FAIL`。
**任何 OOXML 手改都必须过一遍 `motion.ps1` 真开**。
## 12. `<p:sld>` 子元素是 sequence，位置错了切换会被静默丢弃（最隐蔽）

**现象**：`EntryEffect` 读回 `0x0`，再次 `Save()` 后切换消失，全程不报错

`<p:sld>` 的元素模型是 `xsd:sequence`：

```
cSld, clrMapOvr?, transition?, timing?, extLst?
```

`<p:transition>` 必须排在 `<p:timing>` **之前**。把它放在后面（比如直接追加到
`</p:sld>` 前）会发生什么：

| 现象 | 说明 |
| --- | --- |
| PowerPoint 打开**不报错** | 结构校验（lxml）也通过 |
| `SlideShowTransition.EntryEffect` 读回 **0** | PowerPoint 根本没绑定这个元素 |
| 再次 `Save()` 后切换**消失** | 被静默写没了 |

一次 `apply` 里 `<p:timing>` 先插到 `</p:sld>` 前，切换插到"它前面"，结果切换就落进了
timing 内部。修法是按 sequence 顺序插入（`motion.py` 的 `insert_in_slide_order()`）。
实测（`engine2.pptx` → PowerPoint 往返）：修好后 `EntryEffect` 从 `0x0` 变成
`0xF09/fade`、`0xF0F/push`、`0xB01/wipe`，`<p:fade/>` 原样保留。

顺带一个读法问题：PowerPoint 自己保存时会把切换写成
`mc:AlternateContent{ Choice(p14:dur) , Fallback(无 dur) }`，所以**一份健康的
PowerPoint 文件里一个切换会有两个 `<p:transition>` 元素**。数元素个数会误报。
`verify_motion.py` 因此只数"生效"的那个（`motion.active_transition_blocks()`）。
## 26. `a:rot` 的 `lat/lon` 必须是 0~21600000，负数会让**整个文件**损坏

**现象**：3D 旋转角度写成负数后，整份文件打不开

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
## 31. 「往序列里插一个已经存在的元素」—— 本项目**所有**损坏文件都是这一个原因

**现象**：所有打不开的文件，最后都是「往序列里插了一个已经存在的元素」

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

---
## 40. 切换的候选必须**一个一份 deck** —— 合并起来探就没有结论可言

**现象**：一次把多个切换候选放进同一份 deck 去探，PowerPoint 直接拒开

**症状**：把 48 个切换候选合并成一份 97 页的 deck 去探，PowerPoint 报
`could not open the file`，**一个结论都拿不到**。

**根因**：一个非法的 `<p:transition>` 子元素会让 PowerPoint 拒开**整份文件**
（同 §26）。48 个候选里只要有一个坏，整份就开不了 —— 而"开不了"不会告诉你
是哪一个坏的。

**修法**：每个候选一份 2 页 deck，逐份开合。**"PowerPoint 拒开这份"是一等公民的
结论**，不是要吞掉的异常。

**这条还有第二层**：第一批探测里 48 个**全部**被拒开，包括最普通的 `<p:fade/>`。
全部拒开就意味着问题**不在候选**，而在构建器 —— 一查是我们自己把 `xmlns:a`
在 `<p:sld>` 上声明了两次（python-pptx 已经声明过 a/p/r），part 不合法。

> **症状的"粒度"本身就是信息**：一个拒开 = 那个候选有问题；
> 全部拒开 = 生成它们的那份代码有问题。

---
## 42. 切换 / 翻转 / 库 / 摩天轮 / 传送带 **缺 `dir` 会被拒开整份文件**

**现象**：切换 / 翻转 / 库 / 摩天轮 / 传送带 一写，PowerPoint 拒开整份文件

**症状**：`<p14:switch/>`、`<p14:flip/>`、`<p14:gallery/>`、`<p14:ferris/>`、
`<p14:conveyor/>` 五个裸写全部被拒开。元素名是从 [MS-PPTX] 的 p14 元素表来的，
看起来完全合理。

**根因**：这五个**必须**带 `dir`。PowerPoint 自己永远写 `dir="l"`。

**关键教训**：**"拒开"只能说这个写法不行，说不出正确的写法是什么。**
猜名字这条路走到这里就断了 —— 换个名字继续猜是在浪费时间。

**正确的做法（照抄 §23「让 PowerPoint 自己写一遍」）**：造一份每页一个
`PpEntryEffect` 值的 deck，用 COM 把值逐页设上去、保存、读回 XML。
PowerPoint 自己给出的就是答案：

```text
3880  <p14:gallery dir="l"/>      3901  <p14:switch dir="l"/>
3882  <p14:conveyor dir="l"/>     3905  <p14:flip dir="l"/>
3899  <p14:ferris dir="l"/>
```

**枚举扫描还顺手纠正了旧的 COM 枚举表**：旧表把 fade 与 strips 都记成 `0x0A01`、
push 与 randombar 都记成 `0x0901`。实测是 fade=3849 / strips=2561 /
push(dir=u)=3855 / zoom(dir=in)=3074 —— 旧表那套 2003 年代的值与本机完全不同。

完整的 48 项对照表见 [`transitions.md`](transitions.md)。

---
