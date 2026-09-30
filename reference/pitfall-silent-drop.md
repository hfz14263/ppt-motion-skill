# 踩坑 · XML 里有、PowerPoint 不认

> XML 里有、PowerPoint 不认 —— 静默丢弃与静默退回
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

这一类的共同点是**不报错**：XML 合法、结构自检通过，PowerPoint 只是安静地把你的东西丢掉或退回默认。判据永远不是「校验通过」，而是「真开一遍看效果在不在」。

---

## 13. `<p:cTn id>` 必须按文档顺序递增，否则强调动画会被静默丢弃

`<p:cTn id>` 不只是主键：PowerPoint 靠它重建时间线，**每个节点必须比自己的祖先编号大**。
如果先实例化效果的子节点、再分配外层骨架的两个 id，就会得到

```
..., 12, 13, 10, 11, 15, 16, 14
```

最后一个效果自己的动画节点（14）排在了它的父骨架（15,16）**前面**。PowerPoint 打开时
照单全收（`MainSequence.Count` 是对的），但 `Save()` 后**这个效果直接消失**，没有任何
提示。修法：`build_timing()` 先取外层两个 id，再实例化子节点。修好后整页 id 单调递增
（`1,2,5,6,7,8,...,16`）。
## 19. 「静默丢弃」的第三种成因：元素放错了父节点

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
## 41. `<mc:Fallback>` 不总是 `<p:fade/>` —— core 元素在 Fallback 里**重复自己**

**症状**：16 个 core 切换候选往返后全部被判「被改写」，看起来像 PowerPoint
不认这些元素。但子元素一个都没变。

**根因**：给 core 元素配了 `<p:fade/>` 作为 Fallback。PowerPoint 的逻辑是
**core 元素本来就向后兼容，Fallback 应该重复它自己**：

```xml
<!-- ✗ 写的：Fallback 是 fade -->
<mc:Choice Requires="p14"><p:transition p14:dur="800"><p:push dir="u"/></p:transition></mc:Choice>
<mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>

<!-- ✓ PowerPoint 保存后：Fallback 是 push 自己 -->
<mc:Choice Requires="p14"><p:transition p14:dur="800"><p:push dir="u"/></p:transition></mc:Choice>
<mc:Fallback><p:transition><p:push dir="u"/></p:transition></mc:Fallback>
```

只有 p14/p15/p159 的子元素才降级成 `<p:fade/>`。

**顺带两条同源的**：

* **每个切换都会被包进 `mc:AlternateContent`**，连 core 的也是 —— 因为 `p14:dur`
  是 2010 年的属性，PowerPoint 选择包住整个 `<p:transition>` 而不是丢掉时长。
  以为"只有 morph 需要包"是错的。
* **PowerPoint 会删掉值等于默认值的属性**：`<p:push dir="l"/>` 存出来是 `<p:push/>`，
  `<p:split orient="horz" dir="out"/>` 存出来是 `<p:split/>`。无害，但不会留下。

---
## 49. 属性不是到处都能加：`dir` 写错地方会让 PowerPoint 直接打不开

补测「第一行其余成员」的镜像性时，烘 42 份 deck 渲染，其中 4 份
（`box` 的两个方向、`comb` 的两个方向）PowerPoint **拒绝打开**：

```
  box_dir_l        OPEN-FAIL  PowerPoint could not open the file.
  box_dir_r        OPEN-FAIL  PowerPoint could not open the file.
  comb_dir_d       OPEN-FAIL  PowerPoint could not open the file.
  comb_dir_u       OPEN-FAIL  PowerPoint could not open the file.
```

### 49.1 先做隔离，别急着改代码

"打不开"有很多种可能（XML 拼错 / 编码 / 元素名错）。做**最小隔离**：

| 条件 | 结果 |
| --- | --- |
| `box`（`<p:zoom/>`）**不写** `dir` | ✅ 正常打开、正常渲染 |
| `box` 写 `dir="l"` / `dir="r"` | ❌ 打不开 |
| `comb`（`<p:comb/>`）**不写** `dir` | ✅ 正常打开 |
| `comb` 写 `dir="u"` / `dir="d"` | ❌ 打不开 |

**同一个效果、同一份代码、只差一个属性** —— 变量唯一，结论就确定了：
**`<p:zoom>` 与 `<p:comb>` 的元素定义不接受 `dir` 属性。**

### 49.2 为什么这条容易踩

前面已经知道"不是所有效果都支持方向"，但那是**"加了会被忽略"**——
不报错、不崩、只是没效果（§三 早先的说法）。

**这次是另一种失败模式：加了直接打不开。** 两者危害完全不同：

- **被忽略**：最坏是"没达到预期"，文件还能用，肉眼能发现。
- **打不开**：交付物直接废掉，而且是在**渲染/放映阶段**才暴露 ——
  生成时一切正常（XML 是我们自己写的，zip 也打得开），
  只有真正让 PowerPoint 打开才炸。

**教训**：往已知元素上加**非它定义的属性**，属于"写坏文件"而不是"写了无效值"。
凡是要加属性，先确认**这个元素接受它** —— 不能因为"同族别的元素接受"就推过来
（`push` 接受 `dir`，不代表 `zoom`/`comb` 也接受）。

### 49.3 顺带：`box` 的真实身份

这次顺带发现界面名与 XML 元素**不是一对一**：
界面名「框」的 `spec=box`，写出的元素其实是 **`<p:zoom>`**（另一处
`spec=zoom` 才是 `<p:zoom dir="in"/>`）。测试里那句"文件名对但属性没写"
的检查这回派上了用场 —— 它读出 XML 才发现 `box` 根本不是 `<p:box>`。

**只按界面名/效果名推断 XML 结构，一定会错。**

### 49.4 沉淀

- **加属性前先确认元素接受它**；同族别的元素接受不代表它接受。
- 出现"打不开"时**先做单变量隔离**（同一输入、只差待查的那一项），
  再动代码。
- "被忽略"和"打不开"是两种失败，**都要在文档里写明是哪一种**
  —— 只写"不支持"读者会以为是前者。

---
## 50. 非法属性值不报错，只是**静默退回默认** —— 你以为写了，其实没写

补"属性取值全集"时烘了 `clock` 的各个 `spokes` 值。其中 `spokes="6"` 是
**枚举里没有的**（enum 只有 1/2/3/8，加默认 4）。渲染结果：

```
clock_attr_spokes4   72125 bytes     <- 显式写 4
clock_attr_spokes6   72125 bytes     <- enum 里没有的 6
clock_attr_plain     72125 bytes     <- 完全不写
逐帧像素差：spokes4 vs plain = 0.000；spokes6 vs plain = 0.000
```

**三个渲染逐帧完全相同。** 结论两条：

1. **默认值 = 4 根**（不写 ≡ `spokes="4"`，PowerPoint 保存时也会把 4 删掉）。
2. **`spokes="6"` 被静默退回默认** —— 文件照常打开、照常渲染，
   **没有任何报错或警告**，只是你写的值被丢掉了。

### 50.1 为什么危险

这是**第三种**"写了属性但没生效"的形态，前两种已经记过：

| 形态 | 症状 | 举例 | 记在哪 |
| --- | --- | --- | --- |
| 被忽略 | 文件正常，属性不产生任何效果 | `dir` 加在 `page_curl` 上 | §九 / §49 |
| 打不开 | PowerPoint 拒开 | `dir` 加在 `<p:zoom>`/`<p:comb>` 上 | §49 |
| **静默退回默认** | 文件正常，属性**看似生效**（其实是默认值） | `spokes="6"` | 本节 |

第三种最隐蔽：**你写了一个值、它也确实渲染出来了、看起来一切正常** ——
但它渲染的是**默认**，不是你要的值。如果那个"值"恰好接近默认，
肉眼根本分不出来。

### 50.2 判据

**只信枚举里列过的值。** 本文档这套流程已经能把枚举抓全
（`enumdeck` + 枚举扫描 + `enumread`，1..4200 跑完共 149 个值）——
所以"这个属性有哪些合法值"是可以**查**的，不用试。

要确认某个值有没有生效，**必须渲染比对**，不能只看"文件能打开"：

```
attrdeck <dir>            # 每个取值一份 deck
probe_shapes.ps1 -Dir <dir>
attrdiff <dir> --out j.json   # 逐帧像素差 vs 默认
```

`max_frame_diff = 0.000` 就是"等于默认"，**不管那个值看起来多合理**。

### 50.3 沉淀

- **枚举里没有的属性值 = 默认值**（或拒开）。不要凭"界面里好像有"去写。
- 验证一个属性值是否生效，**判据是渲染差异，不是文件能否打开**。
- "文件能打开"只排除了"写坏"，**不证明"写对"**。

---
