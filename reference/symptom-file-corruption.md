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

**现象**：打开直接报「需要修复」，`E_FAIL`，连从 PowerPoint 自己文件里抠出来的原封节点也不行

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §1。

<p align="right"><sub>来源 com-pitfalls §1</sub></p>

### §2 重复属性 → XML 非法

**现象**：lxml 报 `Attribute id redefined`，XML 非法

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §2。

<p align="right"><sub>来源 com-pitfalls §2</sub></p>

### §26 `a:rot` 的 `lat/lon` 必须是 0~21600000，负数会让**整个文件**损坏

**现象**：3D 旋转角度写成负数后，整份文件打不开

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §26。

<p align="right"><sub>来源 com-pitfalls §26</sub></p>

### §31 「往序列里插一个已经存在的元素」—— 本项目**所有**损坏文件都是这一个原因

**现象**：所有打不开的文件，最后都是「往序列里插了一个已经存在的元素」

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §31。

<p align="right"><sub>来源 com-pitfalls §31</sub></p>

### §40 切换候选合并成一份 deck → 拒开且无法归因

**现象**：一次把多个切换候选放进同一份 deck 去探，PowerPoint 直接拒开

把 48 个切换候选合并到一份 deck 里探，PowerPoint 报 `could not open the file`，
**一个结论都拿不到** —— 一个非法子元素就足以拒开整份（同 §26），而"开不了"
不会告诉你是哪一个坏的。

**一个候选一份 deck**，"这份被拒开"才是可归因的结论。

**症状的粒度本身就是信息**：一个拒开 = 那个候选有问题；**全部**拒开 =
生成它们的那份代码有问题（第一次探测 48 个全拒，根因是我们自己把 `xmlns:a`
在 `<p:sld>` 上声明了两次 —— python-pptx 已经声明过 a/p/r）。

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §40。

<p align="right"><sub>来源 com-pitfalls §40</sub></p>

### §42 切换 / 翻转 / 库 / 摩天轮 / 传送带：裸写会被拒开整份文件

**现象**：切换 / 翻转 / 库 / 摩天轮 / 传送带 一写，PowerPoint 拒开整份文件

`<p14:switch/>`、`<p14:flip/>`、`<p14:gallery/>`、`<p14:ferris/>`、`<p14:conveyor/>`
五个裸写全部被拒开。**这五个必须带 `dir`**，PowerPoint 自己永远写 `dir="l"`。

**"拒开"只说明这个写法不行，说不出正确的写法。** 猜名字到此为止 ——
改用枚举扫描让 PowerPoint 自己写一遍（同 §23 的做法），答案就出来了。

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §42。

<p align="right"><sub>来源 com-pitfalls §42</sub></p>

### §43.2 一个 `<p:transition>` 里塞两个子元素 → 拒开整份文件

**现象**：切换效果整体错位一个页边界；一个 `<p:transition>` 里塞两个子元素就拒开；时长在 WPS / 旧版 / 在线预览里全变成「中等」（§43 的第 2 条）

`<p:push/>` + `<p159:morph/>`、`<p:push/>` + `<p:wipe/>`，两组都拒开。
**不是丢一个、不是留第一个，是整份打不开** —— 与 §42 同一类失败。

于是「整页怎么过去」只有一个名额：**形状形变与整页位移不能同时表达**，
页内的个体行为必须走 `<p:timing>`。详见
[`transition-model.md`](transition-model.md) 二。

<p align="right"><sub>来源 com-pitfalls §43.2</sub></p>

---

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §43（根因条目 §43.2）。

<p align="right"><sub>来源 com-pitfalls §43.2</sub></p>
