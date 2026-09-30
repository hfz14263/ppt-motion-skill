# 症状 · XML 里有，PowerPoint 不认（静默丢弃）

> 文件能开、不报错，但 PowerPoint **安静地丢掉**你的元素或属性。这一栏的共同点就是「不报错」，所以判据只能是「真开一遍看效果在不在」。
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## XML 里有，PowerPoint 不认（静默丢弃）

最隐蔽的一类：文件能打开、不报错、不提示修复，**元素就在 XML 里，但被忽略**。字符串检查、结构检查、lxml 解析**全部会通过**。判据只有一个：**用 PowerPoint 读回**（对象模型 / 真渲染），不要看字符串。

### §8 `p14:dur` 需要声明前缀

想带毫秒时长要写 `p14:dur="800"`，但 pptd 导出的 slide 根元素**没有** `xmlns:p14`。
必须自己补上，否则是未绑定前缀（非法 XML）：

```xml
<p:sld … xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
```

<p align="right"><sub>来源 com-pitfalls §8</sub></p>

### §12 `<p:sld>` 子元素是 sequence，位置错了切换会被静默丢弃（最隐蔽）

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

<p align="right"><sub>来源 com-pitfalls §12</sub></p>

### §13 `<p:cTn id>` 必须按文档顺序递增，否则强调动画会被静默丢弃

`<p:cTn id>` 不只是主键：PowerPoint 靠它重建时间线，**每个节点必须比自己的祖先编号大**。
如果先实例化效果的子节点、再分配外层骨架的两个 id，就会得到

```
..., 12, 13, 10, 11, 15, 16, 14
```

最后一个效果自己的动画节点（14）排在了它的父骨架（15,16）**前面**。PowerPoint 打开时
照单全收（`MainSequence.Count` 是对的），但 `Save()` 后**这个效果直接消失**，没有任何
提示。修法：`build_timing()` 先取外层两个 id，再实例化子节点。修好后整页 id 单调递增
（`1,2,5,6,7,8,...,16`）。

<p align="right"><sub>来源 com-pitfalls §13</sub></p>

### §19 「静默丢弃」的第三种成因：元素放错了父节点

§5 讲过"改完 OOXML 必须用 PowerPoint 真开一次"。这里补一个更隐蔽的变体：
**元素写在了正确的文档里、正确的元素附近，但父节点不对，PowerPoint 不报错、直接丢掉。**

实测样例（图片亮度 −40）：

```xml
错的：<a:blip r:embed="rId2"/><a:lum bright="-80000"/>
      └──── 自闭合 ────┘  于是 <a:lum> 成了 <a:blip> 的「兄弟」

对的：<a:blip r:embed="rId2"><a:lum bright="-80000"/></a:blip>
                            └──── 必须在 <a:blip>「内部」 ────┘
```

**症状和"这个版本不支持该功能"一模一样**：打开正常、保存后元素消失。
所以很容易误诊成版本问题，然后放弃一个其实能做的功能。

### 排查方法：让 PowerPoint 自己写一遍

不要读 schema 猜，也不要只试一种写法。**让 PowerPoint 用它自己的对象模型设一次，
再把它写出来的 XML 读回来对比。** 这是唯一能把"我写错了"和"它不支持"分开的办法：

```powershell
$pf = $sh.PictureFormat
$pf.Brightness = 0.1                 # COM 侧：0.0~1.0，0.5 = 0%
$pres.SaveAs("out.pptx")             # 然后读 out.pptx 里它写了什么
```

对比时**逐字符看父节点的开闭**，不要只看元素本身在不在。

### 顺带记住两个数值口径

| 界面 | COM `PictureFormat.Brightness` | OOXML `a:lum/@bright` |
| --- | --- | --- |
| 0% | 0.5 | 省略 |
| −40% | 0.1 | `-80000` |
| +20% | 0.6 | `20000` |

换算：`bright = (COM值 − 0.5) × 200000`。

**别在两个口径之间混用**：COM 传 `-0.4` 会得到"参数无效"，
而按界面百分比猜 XML 会写成 `-40000`（少一倍）。两个错叠加起来，
看起来就像"这个功能根本不存在"。

<p align="right"><sub>来源 com-pitfalls §19</sub></p>

### §20 Morph（平滑）**可以注入** —— 旧结论是写法错误，不是版本限制

上游曾把 Morph 判为"本机不可用"，并据此推论"依赖 Morph 补间的效果无法用注入复现"。
**那个结论是错的。** 复核后确认：三条"证据"里第一条就是根因。

### 20.1 根因：元素名写错了

旧写法是

```xml
<p:transition><p:morph/></p:transition>     <!-- 错 -->
```

**`p:` 命名空间里根本没有 `morph` 元素**。它不在 ECMA-376 的
`CT_SlideTransition` 里；morph 是 Microsoft 的 p159 扩展。所以 PowerPoint
打开时当作未知元素**静默丢弃**，`SaveAs` 之后连 `<p:transition>` 都不剩 ——
看起来就像"这个版本不支持"。

### 20.2 正确写法（取自 `material/template1` 与 `template2` 的真实文件）

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="2000">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="slow"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

要点：

| 项 | 值 |
| --- | --- |
| 元素 | `p159:morph`（**不是** `p:morph`） |
| 命名空间 | `http://schemas.microsoft.com/office/powerpoint/2015/09/main`（**2015/09**，不是 2019） |
| 必需包裹 | `mc:AlternateContent` + `mc:Choice Requires="p159"` |
| 降级 | `mc:Fallback` 里放 `<p:fade/>`，旧版按淡入处理 |
| 选项 | `option="byObject"`（默认，按对象匹配并补间） |
| 时长 | `p14:dur`（毫秒），需 `xmlns:p14` 声明 |
| 位置 | `<p:sld>` 的 `sequence` 里，与其他切换同位置 |

### 20.3 实测证据（Office LTSC 2024，16.0.17928.20148）

| 检查 | 结果 |
| --- | --- |
| PowerPoint 打开 | 正常，无修复提示 |
| `SlideShowTransition.EntryEffect` 读回 | **`0xF72`（3842）** —— 它认得 |
| `SaveCopyAs` 往返后 | slide XML 里 `p159:morph` **仍在** |
| 两页间的形状差异 | 位置/尺寸/3D 角度不同 → 有可补间的差值 |

所以第 2、3 条旧"证据"也解释通了：

- **"对象模型里没有它"** —— 对，`SlideShowTransition` 只暴露旧式成员，
  **但这不影响注入**。注入是写 XML，不是设属性。
- **"按名字设 `EntryEffect = "ppEffectMorph"` 失败"** —— 因为 morph 不是一个
  可赋值的枚举成员。但**读**得回来：正确写入后 PowerPoint 报 `0xF72`。
  **可读 ≠ 可写**，旧结论把这两件事混为一谈了。

### 20.4 Morph 的匹配规则（来自两份 material 模板）

morph 跨两页配对形状，**按形状名/id 匹配**，然后补间差异。所以做动效的方式是：

1. 把第 1 页整页复制成第 2 页（形状名与 id 保持一致）；
2. 在第 2 页改动形状的 **位置 / 尺寸 / 旋转 / 3D 相机角度**；
3. 给**第 2 页**（被进入的那页）加 `p159:morph`。

`material/template1` 就是这么做的：同一个 `图片 8`，
第 1 页 `camera lat="0"`、第 2 页 `camera lat="17400000"`（290°），
其余不变 —— 播放时平面平滑"翻倒"。

### 20.5 仍需注意

- 从 `material/template2` 实测：五个矩形组在第 1 页位于 `left=-180.8pt`（画板外），
  第 2 页铺开 —— **两页确实不同，morph 补间出"矩形从左侧滑入"的效果**，
  是可直接运行的真实范例。
- ⚠️ 但读这些坐标时要注意**分组**：形状在 `<p:grpSp>` 里，用正则按
  `<p:sp>` 切分 spTree 会读到组变换而不是子形状的 `<a:off>`，结论会完全错。
  本 skill 的 `geometry_fingerprint` 不受影响（它两个 `grpSp` 与子 `sp` 都取）。
  **做 morph 必须保证两页之间存在真实差异**，否则什么都不会动。
- `mc:Fallback` 不能省：WPS / 旧版 PowerPoint / 部分在线预览只认 fallback，
  没有它时这些环境可能整页切换失效。

<p align="right"><sub>来源 com-pitfalls §20</sub></p>

### §22 判断"不支持"之前，先排除"我写错了"

§19 和 §20 是同一个教训的两面，值得单独留一条。
**而 §20 后来被证明正是这个坑的实例** —— 它把"元素名写错"（`p:morph`
应为 `p159:morph`）判成了版本不支持，直到拿到 PowerPoint 自己写的 morph 文件
（`material/template1`、`template2`）才纠正。**所以这一条不是抽象原则，是踩过的。**

遇到「元素写进去、保存后消失」，**不要立刻下结论说是版本不支持**。
两个成因的症状完全一样，但处理方式相反：

| 成因 | 判据 | 处理 |
| --- | --- | --- |
| **我写错了**（父节点 / 属性 / 值域） | **让 PowerPoint 自己写一遍**，对比它写的 XML —— 它能写出来，就说明它支持 | 照它写的形式改注入 |
| **确实不支持** | 用它自己的对象模型也**设不上**，或它自己保存后**也不写**这个元素 | 换路线 |

**顺序很重要。** §19 那个亮度问题，我一开始跳过了这一步，直接判成"未攻克"，
结果它只是父节点写错——**一个本来能做的功能差点被我放弃。**

**说"不支持"要三条同时成立**：

1. 手写能正常打开，但保存后元素消失；
2. 对象模型里**没有**对应成员；
3. 按名字设**也失败**，且让 PowerPoint 自己保存后，它**也不写**这个元素。

只满足第 1 条，多半是 §19 那个坑。

⚠️ **§20（Morph）就是一次误判**：它当时声称三条都满足，但复查发现
第 1 条的真正原因是元素名写错（`p:morph` 应为 `p159:morph`），
第 2、3 条只说明"属性模型里没有"，**并不能推出"注入也不行"**。
所以这个三条件清单还要补一条前提：**第 2、3 条只对"用对象模型写"有效，
对"直接写 XML"不成立** —— 判断注入能力必须用第 1 条的做法
（让 PowerPoint 自己写一遍来对比）。

<p align="right"><sub>来源 com-pitfalls §22</sub></p>

### §29 窗口化图片填充：**一个函数里踩出四种静默失败**

做"用形状切割图片"（配方 §8）时，`<a:blipFill>` 的注入连着错了四次。**四次都产出了
能打开、不报错、不提示修复、但渲染结果错误的文件**，而且字符串检查和结构检查全部通过。
只有渲染才暴露。

| # | 错误 | 现象 | 为什么静默 |
| --- | --- | --- | --- |
| **A** | 写成 `<p:blipFill>` | 回落到形状原有的填充色 | `blipFill` 属于 **drawingml `a:`**，`p:blipFill` 是未知元素，直接忽略 |
| **B** | 插在 `<a:ln>` **之后** | 同上 | `CT_ShapeProperties` 是 sequence，顺序非法就丢弃 |
| **C** | 没有删掉原有的 `<a:solidFill>` | **保留原色**（品红），图片不出现 | 填充是 `xsd:choice`，**只能有一个，且取第一个** |
| **D** | `<a:blip r:embed="rId9">` 是编的 id | 画**缺图占位符** | 关系不存在，但你看到的只是一个颜色 |

### 最该记住的两条

**A 是 `p:morph` vs `p159:morph` 的同一类错。** 前缀不是一个可以随便选的装饰 ——
它决定元素属于哪个命名空间，写错就是不同的元素。**判断依据应当是"这个元素定义在哪个
schema 里"，不是"它出现在 `p:` 元素内部"。**

**C 违反了"填充只有一个"这条**，而且症状具有欺骗性：形状照常渲染、只是颜色不对，
看起来像"主题色没改"，不像"注入失败"。

### D 差点让诊断走偏

早期测出三个窗口显示**同一个颜色 `(120,164,232)`**，我判成"PowerPoint 的缺图占位符"，
于是去找关系问题。**但那张测试图（红绿蓝三带 + 白）的平均色恰好接近那个蓝**
（实测平均 `(103,96,116)`，但仍接近），所以**"图片缺失"和"图片存在但被极度放大到只剩
一片平色"这两种截然不同的情况，在这个判据下无法区分**。

**对策**：不要用"颜色像不像占位符"当判据。用**差分**——
同一形状给不同内缩值，若渲染结果**完全相同**，才是真的没生效。

### 正确性怎么确立的

不再推理，而是**让 PowerPoint 自己写一遍**：COM 的 `Shape.Fill.UserPicture("图")`
会生成官方写法，读回来对照即可。它写的是：

```xml
<p:spPr>…<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
</p:spPr>
```

**与实测的差值只有那个 `p:` / `a:`。** 改成 `a:blipFill` 后三个窗口立刻分别显示
红/绿/蓝三条不同色带 —— 既证明了填充生效，也顺带证明了窗口位置正确。

> 这与 §20 的教训是同一条：**"我的写法不生效" ≠ "这个功能不支持"。**
> 先让 PowerPoint 写一遍，比自己读文档或推理都快。

<p align="right"><sub>来源 com-pitfalls §29</sub></p>

---
