# 症状 · 工具与环境

> 不是文档写错了，是**工具或环境**在坑你 —— 脚本、编码、导出。
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 工具与环境

与 PowerPoint 无关，是**工具链本身**的坑：编码、解析、断言口径、COM 约束、以及「默认只读」这类安全设计。

### §3 lxml 不能直接 parse 带编码声明的 str

**现象**：`ET.fromstring()` 抛 `ValueError: Unicode strings with encoding declaration are not supported`

```python
ET.fromstring(z.read(part).decode('utf-8'))        # ValueError
ET.fromstring(z.read(part))                        # bytes，正常
```

**根因与修法**：见 [`pitfall-ooxml.md`](pitfall-ooxml.md) §3。

<p align="right"><sub>来源 com-pitfalls §3</sub></p>

### §4 结构校验通过 ≠ PowerPoint 能打开

**现象**：lxml 结构校验全绿，PowerPoint 打开却 `E_FAIL`

lxml 只保证 well-formed。实例：v3（自建 timing）结构校验 OK，PowerPoint `E_FAIL`。
**任何 OOXML 手改都必须过一遍 `motion.ps1` 真开**。

**根因与修法**：见 [`pitfall-file-corruption.md`](pitfall-file-corruption.md) §4。

<p align="right"><sub>来源 com-pitfalls §4</sub></p>

### §7 切换的旧式枚举几乎不可用

**现象**：设了 `EntryEffect = 0x0A01` 想做 fade，读回是 `<p:strips/>`；`0x1701` 直接报非法枚举

本机实测（`SlideShowTransition.EntryEffect`）：

| 枚举 | 实际写出的元素 |
| --- | --- |
| `0x0A01`（以为 fade） | `<p:strips/>` |
| `0x0901` | `<p:randomBar/>` |
| `0x0801` | `<p:pull/>` |
| `0x0C01` | `<p:zoom/>` |
| `0x1701`（以为 zoom） | **非法枚举，报错** |

所以**以写出的 `<p:transition>` 元素为准**，COM 属性只用于复核读数。实测 PowerPoint 会按我们
写的元素（如 `<p:fade/>`）保存。

**这张表已经不够用了**：它给出的值还是 2003 年代那套，而且**多个元素撞成同一个值**
（fade 与 strips 都记 `0x0A01`）。枚举扫描量到本机的真实值是
fade=3849 / strips=2561 / push(dir=u)=3855 / zoom(dir=in)=3074，和上面完全不同。

**48 项界面切换的完整实测对照表见 [`transitions.md`](transitions.md)** ——
界面名 ↔ PowerPoint 写出的子元素 ↔ 完整 XML 块 ↔ 实测枚举，四列对齐。

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §7。

<p align="right"><sub>来源 com-pitfalls §7</sub></p>

### §9 Office COM 的几个约束

**现象**：COM 调用在别人机器上好用，在本机报`自动化错误`或对象不存在

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §9。

<p align="right"><sub>来源 com-pitfalls §9</sub></p>

### §10 Windows PowerShell 5.1 的编码坑

**现象**：脚本输出中文全是乱码，或`Get-Content` 读出来是乱码

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §10。

<p align="right"><sub>来源 com-pitfalls §10</sub></p>

### §11 颜色/几何断言要按"布局帧"口径

**现象**：同一个几何值，脚本判绿、PowerPoint 真开却肉眼可见地错位

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §11。

<p align="right"><sub>来源 com-pitfalls §11</sub></p>

### §15 `motion.ps1` 默认只读，绝不回写输入（旧版会毁掉源文件）

**现象**：用旧版 `motion.ps1` 跑过之后，源 pptx 里的效果没了

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §15。

<p align="right"><sub>来源 com-pitfalls §15</sub></p>

### §16 静态导出与逐形状导出（做"动效预览"必需）

**现象**：想看动效长什么样，只能看到一堆静态图，分不清哪帧是过渡

**根因与修法**：见 [`pitfall-com.md`](pitfall-com.md) §16。

<p align="right"><sub>来源 com-pitfalls §16</sub></p>

### §32 `CreateVideo` 报"完成"却拿不到文件 —— 三个参数全错

**现象**：`CreateVideo` 报「完成」，但目标文件不存在或打不开

轮询**立刻退出**，日志写 `status=2 exists=True`，几秒后文件**不存在**。

根因：`PpMediaTaskStatus` 里 **2 = Failed、3 = Done**，我把 `2` 读成了完成；
`exists=True` 是编码器刚落的一个**空壳文件**，Quit 之后被清掉。三处错叠加：

| 错 | 后果 |
| --- | --- |
| `WithWindow = 0` 打开 | 编码器起不来（probe 脚本用的是 `WithWindow = 1`） |
| 没重定向 `TEMP/TMP` | 沙箱 TEMP 会让编码器找不到可写的临时目录 |
| `$pres.CreateVideo($f,$true,6)` 三参直调 | 部分构建上 late-binding 丢参，要 `.Invoke(@($f,$true,6,720,30,85))` 六参全给 |
| 输出路径超过 255 字符 | 静默不产文件；换 `C:\t1w\mine.mp4` 这类短路径立刻成功 |

**正确姿势**：开窗打开 + TEMP 重定向 + 六参 `Invoke` + 轮询到 `status -eq 3 -or -eq 4`。
同一台机器 15 秒出片（720p/30fps/q85，11.8 MB / 78 秒成片）。

状态码口径：**0/1 = 没开始/进行中，2 = 失败，3/4 = 完成。**

<p align="right"><sub>来源 com-pitfalls §32</sub></p>

---

**根因与修法**：见 [`pitfall-tooling.md`](pitfall-tooling.md) §32。

<p align="right"><sub>来源 com-pitfalls §32</sub></p>

### §43.1 切换效果整体错位一个页边界

**现象**：想让第 5→6 页有切换，写在了第 5 页，结果 5→6 是硬切，
而 4→5 那一页边界动了起来。

**根因**：`<p:transition>` 挂在**终点页**上 —— 写在第 N 页，动的是「进入第 N 页」。
spec 里 `slides[].page` 是终点页，不是出发页。

**为什么查不出来**：两种解释下"切换在第 N 页上"都成立，读 XML 一分都分不出来。
判别必须数渲染帧 —— 3 页 deck、只有第 2 页写 800ms push，
burst 落在边界 1（1→2）而不是边界 2。详见 [`transition-model.md`](transition-model.md) 一。

**附带**：第 1 页没有前驱，它的切换只用在**放映起步从黑场进入开场**那一次，
不是"被忽略"；此后真正的第 1→2 边界仍是硬切。

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §43（根因条目 §43.1）。

<p align="right"><sub>来源 com-pitfalls §43.1</sub></p>

### §43.3 时长在 WPS / 旧版 / 在线预览里全变成"中等"

**现象**：本机 PowerPoint 里时长完全正确（毫秒级），但换到 WPS 或旧版 PowerPoint
播放，**任何时长都跑成差不多的 ~0.75 秒**。

**根因**：`p14:dur` 只在 `mc:Choice` 分支里（它需要 2010 命名空间），
而 `mc:Fallback` 分支**不可能带它**，那里唯一说话的属性是 `spd`（三档）。
如果生成器把 `spd` 写死成 `"med"`，降级世界的时长就全部塌到 med。

**正确的做法**：让 `spd` 跟着 `duration` 取最接近的档位（显式 `speed=` 优先）。
三档的实测值是 **slow ≈ 1000ms / med ≈ 767ms / fast ≈ 500ms**
（不是界面上传说的 3s/2s/1s），由 `tests/test_transition_table.py` 钉住。

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §43（根因条目 §43.3）。

<p align="right"><sub>来源 com-pitfalls §43.3</sub></p>

### §44.4 形态表里 `fade` / `wipe` / `dissolve` 分不出来

**现象**：给 48 个效果量"运动形态"，`fade`、`wipe`、`dissolve`、`split`、`shape`
的径向剖面**几乎完全一样**（差值 <0.03），无法区分。

**根因**：探测 deck 里两页之间**只有中心的圆在变**，所以不论什么效果差异都集中在中心。
它们的差别在**变化怎么扩散**，不在**变化在哪里**。

**修法**：加**第二个探测 deck** —— 两页都是满屏随机噪点（平均亮度相同、空间图案不同）。
这样 `wipe` 的差异成一条**移动的带**（`band_travel = −0.747`），
`fade` 的差异**均匀**（`band_travel = 0.015`），`push` 的差异**铺满全屏**（`coverage = 0.957`）。

**推论**：**没有一个探测 deck 能同时看清所有效果**。先做一个 probe 聚类，
把分不开的那些挑出来给第二个 probe —— 这份名单是聚类的产物，不是猜的。
详见 [`transition-shapes.md`](transition-shapes.md) §四。

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §44（根因条目 §44.4）。

<p align="right"><sub>来源 com-pitfalls §44.4</sub></p>

### §44.3 用 `random` 做切换 → 每次放映观感都不一样

**现象**：同一个 deck 放映两次，页面切换的运动方式不同。

**根因**：`<p:random/>` **每次放映随机挑一个效果**。实测两次探测给出不同结果；
`random` 与 `clock` 的 mp4 **大小完全相同（186,629 字节）但 md5 不同** ——
说明 `random` 那一次恰好抽到了与 `clock` 相同的效果，但两者文件本身不同。

**实践含义**：用 `random` **无法保证任何观感一致性**。商务场合不要用；
需要"不重复"就手动轮换几个确定的切换。

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §44（根因条目 §44.3）。

<p align="right"><sub>来源 com-pitfalls §44.3</sub></p>

### §44.7 标定样本自己错了 → 测试"通过"但什么都没测到

**现象**：分析器的标定测试全绿，但分析器的方向判据其实是坏的。

**根因**：测试的两个合成样本**生成的帧完全相同**（都是"白色从左边长出来"），
于是左右两次都判 `l->r` —— 而期望值也被写反了，**测试就"通过"了**。

**修法**：把右侧样本改成"白色从**右边**长出来"，测试立刻报 FAIL，才暴露问题。

**教训**：造标定样本时先确认**两个样本真的不同**。同构的输入会让恒等映射也通过 ——
**测试通过不等于测试有效**。

<p align="right"><sub>来源 com-pitfalls §44.7</sub></p>

---

**根因与修法**：见 [`pitfall-transition.md`](pitfall-transition.md) §44（根因条目 §44.7）。

<p align="right"><sub>来源 com-pitfalls §44.7</sub></p>
