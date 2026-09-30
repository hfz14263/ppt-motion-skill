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

```python
ET.fromstring(z.read(part).decode('utf-8'))        # ValueError
ET.fromstring(z.read(part))                        # bytes，正常
```

<p align="right"><sub>来源 com-pitfalls §3</sub></p>

### §4 结构校验通过 ≠ PowerPoint 能打开

lxml 只保证 well-formed。实例：v3（自建 timing）结构校验 OK，PowerPoint `E_FAIL`。
**任何 OOXML 手改都必须过一遍 `motion.ps1` 真开**。

<p align="right"><sub>来源 com-pitfalls §4</sub></p>

### §7 切换的旧式枚举几乎不可用

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

<p align="right"><sub>来源 com-pitfalls §7 / §42</sub></p>

### §9 Office COM 的几个约束

- 媒体插入（`AddMediaObject2`）依赖 PowerPoint 自己的转码管线；MP4/WAV 可用，
  **未压缩 AVI 会因缺解码器失败**。手写 OOXML 做不了这件事，必须走 COM。
- 受限沙箱下媒体插入会统一报 "cannot insert the file you specified"；放宽权限后立刻成功。
  这是权限问题，不是格式问题。
- `Presentations.Open(path, ReadOnly, Untitled, WithWindow)` 的 `WithWindow` 传 0 在本机
  **打不开**，必须传 1（会闪一下窗口，正常）。
- 失败的自动化会残留 `POWERPNT` 进程并锁住 pptx（再报 "could not open the file"）。
  先 `Stop-Process POWERPNT -Force`。
- 自动化会在 `%TEMP%` 残留 `*- OProcSessId.dat` 与 `*.tmp`，可清理。
- **PowerPoint 会原样保留注入的动画**：注入 31 个 effect → PowerPoint 打开读到 31 个 →
  `Save()` 后仍是 31 个（已实测）。所以 COM 往返不会吃掉动画。

<p align="right"><sub>来源 com-pitfalls §9</sub></p>

### §10 Windows PowerShell 5.1 的编码坑

- 无 BOM 的 `.ps1` 里写中文会被按 ANSI 读成乱码，**甚至引发语法错误**（中文注释吃掉引号）。
  脚本一律纯 ASCII，中文用 `[char]0xXXXX` 拼。
- 用 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File` 调用。**不要假定 `pwsh` 存在**：
  Windows PowerShell 5.1 是 Windows 自带的，PowerShell 7（`pwsh`）是另装的，
  很多机器上只有前者。脚本按 5.1 的语法子集写，别用 7 才有的东西。
- 解压/打包优先用 .NET 的 `ZipFile` 而不是外部 `tar`：`tar` 在受限环境里可能被策略拦住，
  而 `ZipFile` 不依赖任何外部进程。

<p align="right"><sub>来源 com-pitfalls §10</sub></p>

### §11 颜色/几何断言要按"布局帧"口径

不要把整份 slide XML 做哈希——`<p:timing>` 本来就会变。只哈希
**spTree 内每个形状的 `tag + cNvPr/@id + @name + xfrm(x,y,cx,cy,rot)` 序列**，
这样"动画注入不改版面"才能被机器证明。

<p align="right"><sub>来源 com-pitfalls §11</sub></p>

### §15 `motion.ps1` 默认只读，绝不回写输入（旧版会毁掉源文件）

COM 层是**复核**层，`Presentations.Open(..., ReadOnly=msoTrue, ...)` 打开：
- 它不会再覆盖你传给它的 pptx（旧版用 `$pres.Save()` 就地保存，把 OOXML 引擎写好的
  切换换成了枚举写出的 `<p:strips/>`，并且直接改写了输入文件）；
- 需要持久化的东西（内嵌媒体）写到 `OutDir` 下的 `*.com.pptx`；
- 切换默认**不**走 COM 枚举（那反而会覆盖正确结果，见 §7）；要探测属性模型用
  `-SyncTransitions` 显式开启。

另外 `Slide.Export` 是 COM 调用，**相对路径会按 PowerPoint 自己的工作目录解析**
（不是 PowerShell 的当前目录），于是报"找不到 <你要求的路径>"。`OutDir` 必须绝对化。

<p align="right"><sub>来源 com-pitfalls §15</sub></p>

### §16 静态导出与逐形状导出（做"动效预览"必需）

### 16.1 `Slide.Export` 无视 `Shape.Visible`

```
shape.Visible = 0        # msoFalse
slide.Export(path,...)   # 形状照样画出来
```

所以**不能**靠隐藏其它形状来单独导出一个形状。
（另注意 `msoTrue = -1`，不是 `1`；写 `1` 是无效值。）

### 16.2 `Shape.Export` 才认 `Visible`，且给真 alpha

```python
shape.Export(path, 2)     # 2 = ppShapeFormatPNG -> RGBA，带透明
```

- 第二个参数是**数字枚举**；传 `"PNG"` 抛 `invalid literal for int() with base 10: 'PNG'`。
- 宽高可省略，省了就按 `Shape.ScaleWidth/ScaleHeight` 默认导出（实测 1.45× 形状尺寸，够清晰）。
- 带 alpha 是它能当图层用的前提。

### 16.3 背景要"先导图层、再删形状、最后导整页"

想让"形状逐个浮现"可看，需要 base + 每形状一层。base 必须是**删掉这些形状之后**的整页：

```python
for sid in wanted: shape.Export(...)      # 1. 先导图层
for sid in wanted: shape.Delete()         # 2. 再删掉它们
slide.Export(base, ...)                   # 3. 剩下的当背景
```

顺序反了就会重影：base 里已经烘焙了一份，图层淡入到位后正好叠在自己的副本上。

### 16.4 `Presentation.CreateVideo`：**取决于 Office 构建，必须本机实测**

上游在它的机器上观察到**所有参数组合都失败**：

| Quality | 上游观察 |
| --- | --- |
| 0 | 返回成功码但**不产文件** |
| 1 / 2 | `E_INVALIDARG` |
| VertResolution 480/540/720/1080 | 均失败 |
| ReadOnly / ReadWrite 打开 | 均失败 |

**但另一台机器上完全可用**（Office LTSC 2024 ProPlus Retail x64，构建
`16.0.17928.20148`，POWERPNT.exe 同版本）：

| 参数 | 实测结果 |
| --- | --- |
| 480p / 720p / 1080p @30fps, q85 | 成功，**1.19 MB / 2.49 MB / 4.11 MB**（同一 busy deck） |
| 720p @30fps, quality 1 | 成功，0.54 MB |
| 720p @15fps, q85 | 成功，1.86 MB |
| `WithWindow` = 0 或 1 | **都能导出** |
| `CreateVideoStatus` | 轮询到 `3`（done）即完成，1–9 秒 |

所以**分辨率、质量、帧率参数确实生效**（文件大小随参数单调变化）。

**两边观察有交叉也有冲突，这本身就是关键线索：**

| 参数 | 上游 | 本机（LTSC 2024, 16.0.17928.20148） |
| --- | --- | --- |
| quality 0 | 返回成功码但**不产文件** | **也失败**：`E_INVALIDARG` |
| quality 1 / 2 | `E_INVALIDARG` | **成功**（q1 = 0.54 MB） |
| 480 / 720 / 1080p | 均失败 | **均成功**（1.19 / 2.49 / 4.11 MB） |
| ReadOnly / ReadWrite | 均失败 | **均成功** |

`quality 0` 两边都坏 —— 说明这不是随机的，而是**某个参数值本身有问题**；
其余参数本机可用而上游不可用，指向**构建 / COM 驱动差异**。

**结论（环境限定）**：`CreateVideo` 的可用性**不是 Office 的普遍属性**。差异可能来自
Office 版本、位数、COM 驱动或媒体子系统状态 —— 具体原因未定论，两边观察可以同时为真，
所以**不要用任何一方的结论去推断另一台机器**。

**因此：不要假定，先探测。** 用 `scripts/probe_createvideo.ps1` 在**你自己的机器**上跑一遍：

```powershell
powershell -NoProfile -File scripts/probe_createvideo.ps1
```

它打印本机 Office 构建指纹，并逐个参数组合尝试导出，报告哪些成功。
- 有任一组合成功 → 可以导 MP4 复核动效
- 全部失败 → 用 §16.1–§16.3 的图层方案，或 `motion.py player`

无论哪种情况，`player` 都仍有独立价值：它做**逐形状图层 + 可控时序**，
能暂停在任意时刻单看某一层，这是 MP4 做不到的。

（`CreateVideo` 本身是异步的：即使调用成功也要轮询文件/状态，不能只等返回值。）

<p align="right"><sub>来源 com-pitfalls §16</sub></p>

### §32 `CreateVideo` 报"完成"却拿不到文件 —— 三个参数全错

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

<p align="right"><sub>来源 com-pitfalls §44.4</sub></p>

### §44.3 用 `random` 做切换 → 每次放映观感都不一样

**现象**：同一个 deck 放映两次，页面切换的运动方式不同。

**根因**：`<p:random/>` **每次放映随机挑一个效果**。实测两次探测给出不同结果；
`random` 与 `clock` 的 mp4 **大小完全相同（186,629 字节）但 md5 不同** ——
说明 `random` 那一次恰好抽到了与 `clock` 相同的效果，但两者文件本身不同。

**实践含义**：用 `random` **无法保证任何观感一致性**。商务场合不要用；
需要"不重复"就手动轮换几个确定的切换。

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
