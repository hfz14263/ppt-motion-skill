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

**现象**：写 `p14:dur="800"` 后 lxml 报未绑定前缀，整份文件非法

**根因与修法**：见 [`pitfall-ooxml.md`](pitfall-ooxml.md) §8。

<p align="right"><sub>来源 com-pitfalls §8</sub></p>

### §12 `<p:sld>` 子元素是 sequence，位置错了切换会被静默丢弃（最隐蔽）

**现象**：`EntryEffect` 读回 `0x0`，再次 `Save()` 后切换消失，全程不报错

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §12。

<p align="right"><sub>来源 com-pitfalls §12</sub></p>

### §13 `<p:cTn id>` 必须按文档顺序递增，否则强调动画会被静默丢弃

**现象**：强调动画设了，打开正常、`Save()` 之后效果直接消失，没有任何提示

**根因与修法**：见 [`pitfall-silent-drop.md`](pitfall-silent-drop.md) §13。

<p align="right"><sub>来源 com-pitfalls §13</sub></p>

### §19 「静默丢弃」的第三种成因：元素放错了父节点

**现象**：元素明明写进了 XML，PowerPoint 保存后它不见了

**根因与修法**：见 [`pitfall-silent-drop.md`](pitfall-silent-drop.md) §19。

<p align="right"><sub>来源 com-pitfalls §19</sub></p>

### §20 Morph（平滑）**可以注入** —— 旧结论是写法错误，不是版本限制

**现象**：一直认为 Morph 无法注入，只能退回硬切

**根因与修法**：见 [`pitfall-morph-3d.md`](pitfall-morph-3d.md) §20。

<p align="right"><sub>来源 com-pitfalls §20</sub></p>

### §22 判断"不支持"之前，先排除"我写错了"

**现象**：试了一种写法，PowerPoint 忽略了它，误判成「这个版本不支持」

**根因与修法**：见 [`pitfall-morph-3d.md`](pitfall-morph-3d.md) §22。

<p align="right"><sub>来源 com-pitfalls §22</sub></p>

### §29 窗口化图片填充：**一个函数里踩出四种静默失败**

**现象**：窗口化图片填充写完不报错，但四种填充法里有几种的效果明显不对

**根因与修法**：见 [`pitfall-morph-3d.md`](pitfall-morph-3d.md) §29。

<p align="right"><sub>来源 com-pitfalls §29</sub></p>
