# 踩坑手册（全部为本机实测）

环境：Windows 10 + Microsoft Office 16.0（Office16）+ Windows PowerShell 5.1 + Python 3.9。

## 1. 多余 preset 包装层 → PowerPoint 拒开（最致命）

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

模板的外层 cTn 自带 `id="5"`、`dur`、`fill`。合并属性到骨架 cTn 时如果不过滤，
就会产生 `id="4" … id="5"` 这种**重复属性**，lxml 直接报
`Attribute id redefined`。必须过滤 `id|dur|fill`。

## 3. lxml 不能直接 parse 带编码声明的 str

```python
ET.fromstring(z.read(part).decode('utf-8'))        # ValueError
ET.fromstring(z.read(part))                        # bytes，正常
```

## 4. 结构校验通过 ≠ PowerPoint 能打开

lxml 只保证 well-formed。实例：v3（自建 timing）结构校验 OK，PowerPoint `E_FAIL`。
**任何 OOXML 手改都必须过一遍 `motion.ps1` 真开**。

## 5. 静态预览：为什么要 `preview`（实测修正）

原先这里写的是"带入场动画的形状在 PNG/PDF 里是隐藏的"。**本机实测不成立**：本 skill 的
`fly`/`fade` 模板用的是把 `style.visibility` 设为 `visible` 的 `<p:set>`，从不写隐藏，
所以 `Slide.Export` 直接把形状画出来了（自测 deck 的 animated 版与 preview 版首屏 PNG
逐字节相同）。

那 `preview` 还有什么用：

- **去掉切换**：否则导出的"静态副本"其实还带着翻页效果；
- **让静态导出与播放顺序解耦**：动画链不再影响导出时机，结果可复现；
- 需要"动画开始前"那一帧时，只有 preview 版本是确定的状态。

所以 §5 的正确说法是：`preview` 产出的是一份**无 timing、无 transition** 的干净副本，
用它可以得到确定性的截图/PDF；而不是"不 preview 就看不到形状"。

```bash
python scripts/motion.py preview --pptx animated.pptx --out static.pptx
# 再对 static.pptx 走 motion.ps1 -NoSave -ExportPdf
```

## 6. 颜色是 BGR，且不能用算术表达式生成

- `Shape.Fill.ForeColor.RGB = 0x0B1020`（想要深蓝）→ XML 里存成 `20100B`（红蓝互换）。
  想显示 `#FFD24A` 必须传 `0x4AD2FF`。
- `0x30 * $i * 65536 + …` 这类算术，PowerPoint 收到 Double 会**截断**：
  `#14304F` 被存成 `#13304F`（红通道 −1）。用常量 + `[int]` 强转。

```powershell
function RGBv([int]$displayRGB) {   # 显示色 -> PowerPoint 存的 BGR 值
  $r = $displayRGB -band 0xFF; $g = ($displayRGB -shr 8) -band 0xFF; $b = ($displayRGB -shr 16) -band 0xFF
  return [int](($r -shl 16) -bor ($g -shl 8) -bor $b)
}
```

## 7. 切换的旧式枚举几乎不可用

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

## 8. `p14:dur` 需要声明前缀

想带毫秒时长要写 `p14:dur="800"`，但 pptd 导出的 slide 根元素**没有** `xmlns:p14`。
必须自己补上，否则是未绑定前缀（非法 XML）：

```xml
<p:sld … xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">
```

## 9. Office COM 的几个约束

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

## 10. Windows PowerShell 5.1 的编码坑

- 无 BOM 的 `.ps1` 里写中文会被按 ANSI 读成乱码，**甚至引发语法错误**（中文注释吃掉引号）。
  脚本一律纯 ASCII，中文用 `[char]0xXXXX` 拼。
- 用 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File` 调用。**不要假定 `pwsh` 存在**：
  Windows PowerShell 5.1 是 Windows 自带的，PowerShell 7（`pwsh`）是另装的，
  很多机器上只有前者。脚本按 5.1 的语法子集写，别用 7 才有的东西。
- 解压/打包优先用 .NET 的 `ZipFile` 而不是外部 `tar`：`tar` 在受限环境里可能被策略拦住，
  而 `ZipFile` 不依赖任何外部进程。

## 11. 颜色/几何断言要按"布局帧"口径

不要把整份 slide XML 做哈希——`<p:timing>` 本来就会变。只哈希
**spTree 内每个形状的 `tag + cNvPr/@id + @name + xfrm(x,y,cx,cy,rot)` 序列**，
这样"动画注入不改版面"才能被机器证明。

## 12. `<p:sld>` 子元素是 sequence，位置错了切换会被静默丢弃（最隐蔽）

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

## 14. 同一形状上"入场 + 强调"不兼容 PowerPoint 往返（实测限制，未解决）

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

## 15. `motion.ps1` 默认只读，绝不回写输入（旧版会毁掉源文件）

COM 层是**复核**层，`Presentations.Open(..., ReadOnly=msoTrue, ...)` 打开：
- 它不会再覆盖你传给它的 pptx（旧版用 `$pres.Save()` 就地保存，把 OOXML 引擎写好的
  切换换成了枚举写出的 `<p:strips/>`，并且直接改写了输入文件）；
- 需要持久化的东西（内嵌媒体）写到 `OutDir` 下的 `*.com.pptx`；
- 切换默认**不**走 COM 枚举（那反而会覆盖正确结果，见 §7）；要探测属性模型用
  `-SyncTransitions` 显式开启。

另外 `Slide.Export` 是 COM 调用，**相对路径会按 PowerPoint 自己的工作目录解析**
（不是 PowerShell 的当前目录），于是报"找不到 <你要求的路径>"。`OutDir` 必须绝对化。

## 16. 静态导出与逐形状导出（做"动效预览"必需）

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

## 17. 动画的**中间态**在本机不可观测

想验证 `wipe(left)` 到底往哪个方向擦，必须看到动画跑到一半的样子。三条路都试过，
**全部不通**，记在这里省得重走：

| 探针 | 结果 |
| --- | --- |
| `SlideShowView.Export` | 这个 build **没有这个成员**：调用报 `<unknown>.Export` |
| `Slide.Export` / `Shape.Export` | 导出的是**终态**，完全无视实时 `Visible` 与 filter 状态。实测：6 秒的 wipe 在 2.4 秒采样，八个方向**全部 100% 不透明** |
| 截屏 `ImageGrab` / `PrintWindow` | 需要桌面真的在渲染。实测抓到的整屏 97% 接近纯黑、只有 0.55% 亮于灰 200，即屏幕没有输出（休眠/锁定）。`PrintWindow(hwnd, dc, PW_RENDERFULLCONTENT)` 返回的是**桌面 DC**，不是窗口自身内容 |

§16.2 说 `Shape.Export` 认 `Visible`——那指的是**静态**的 `Shape.Visible` 属性，
不是动画运行中的可见性状态。两者是两回事，实测以终态为准。

**推论**：方向性擦除（`dir`）的效果只能靠**人眼看一次**确认；
`showcase/qa/calib_wipe.py` 只验证能得到证的部分（`dir` 确实写进了 `filter=`、
几何字节不变、给无方向效果写 `dir` 会报错）。

## 18. Round-trip 会以第二种方式咬"一个形状多个效果"

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
什么都折叠不了，版面看起来完全一样。见 `references/authoring-rules.md` §J。




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

## 20. Morph（平滑）**可以注入** —— 旧结论是写法错误，不是版本限制

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

- 从 `material/template2` 实测：它的**两页形状几何完全相同**，所以那份文件
  本身不产生任何补间 —— 它是一份"设计理念说明"，不是可运行范例。
  **做 morph 必须保证两页之间存在真实差异**，否则什么都不会动。
- `mc:Fallback` 不能省：WPS / 旧版 PowerPoint / 部分在线预览只认 fallback，
  没有它时这些环境可能整页切换失效。

## 21. 3D 相机（`scene3d`）可以被注入，且角度原样保留

和 §20 相反，这一项**能**做，实测确认（PowerPoint 16.0 build 20228）：

```xml
<a:scene3d>
  <a:camera prst="perspectiveRelaxedModerately">
    <a:rot lat="0" lon="17400000" rev="0"/>   <!-- 1/60000 度；17400000 = 290° -->
  </a:camera>
  <a:lightRig rig="threePt" dir="t"/>
</a:scene3d>
```

放进 `<p:spPr>` 内、写在几何之后。`SaveAs` 往返后 `lon` **一位不差**。

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

## 22. 判断"不支持"之前，先排除"我写错了"

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

**说"不支持"要三条同时成立**（§20 就是这么确认的）：

1. 手写能正常打开，但保存后元素消失；
2. 对象模型里**没有**对应成员；
3. 按名字设**也失败**，且让 PowerPoint 自己保存后，它**也不写**这个元素。

只满足第 1 条，多半是 §19 那个坑。

## 23. 「三维旋转」的界面角度 ≠ OOXML 的 `lat/lon`

PowerPoint 格式窗格里填的是 **RotationX / RotationY**，而 OOXML 存的是
`<a:camera><a:rot lat lon rev>`，**两者不是同一组数字，轴也不对应**。

实测映射（让 PowerPoint 自己设角度、读回它写的 XML 得到）：

```
窗格 RotationX = -32, RotationY = -20
   ↓ PowerPoint 写出
<a:camera prst="orthographicFront">
  <a:rot lat="20400000" lon="1200000" rev="0"/>     ← 340° / 20°，单位 1/60000 度
</a:camera>

lon = |RotationX|
lat = 360 - |RotationY|        （RotationY 为负时）
投影 = orthographicFront，即平行投影，不是透视
```

### 两个会静默产生错误结果的写法

| 写法 | 得到什么 |
| --- | --- |
| 把窗格角度**直接写进 lat/lon** | **竖起来的菱形**，不是躺平的板——轴不对应 |
| 用 `<a:xfrm rot="...">` 代替 | **水平错切**的平行四边形——那是**平面内**旋转，不是 3D |

**两种写法都结构合法、PowerPoint 都正常打开**，所以什么都不会报错。
只有**渲染出来看**才能发现。

### 正确做法

**让 PowerPoint 写，不要自己推。**

```powershell
$sh.ThreeD.RotationX = -32
$sh.ThreeD.RotationY = -20
$pres.SaveAs("out.pptx")        # 它自己会转成合法的 lat/lon
```

它的输出就是权威；`lat/lon` 只用于**读**和**校核**，不用于猜。

### 而且角度值要**扫**出来，不要算

预设相机表里那些角度（LibreOffice `scene3dhelper.cxx` 有一份整理）**与渲染结果
对不上**（§21 记过：单位换算差约 725 倍）。所以确定姿态的办法是：
**扫一组候选 → 各渲一帧 → 跟参考图比**。我最后用的 `RotationX -32 / RotationY -20`
就是这么定的，不是算出来的。

跟目标图比对时，看三个特征最快：**近边是否比远边宽**（有无透视）、
**左右竖边是否向上收敛**（俯角方向）、**板面是否被压扁**（倾角大小）。

## 24. 造场景图时，图的"结构"决定效果成不成立

立体页要求那个平面能**读成地面**。这取决于图片自己有没有**地平线**和**前景纵深**，
不取决于 3D 角度调得多准。

实测：同一组角度下，
- 用**带山脊 + 麦田**的图 → 一眼就是躺平的地面
- 用**纯网格标定图** → 只看到一个倾斜的方块，读不出"地面"

**所以这类效果里，素材的结构比参数更关键。** 换成纯色或图形素材，
再多迭代角度也出不来那个感觉。

**推论**：做立体页之前先问「这张图有没有地平线」。没有的话，
要么换图，要么先给图加一条。

## 25. `ThreeD.RotationX/Y` 写出的是**平行投影**，永远做不出"躺平的地面"

这是"立体版"那个效果真正卡住的地方，比 §23 的角度映射更根本。

**PowerPoint 的对象模型**：

```powershell
$sh.ThreeD.RotationX = -32
$sh.ThreeD.RotationY = -20
# 它写出：
# <a:camera prst="orthographicFront"><a:rot lat="340" lon="20" rev="0"/>
```

**`orthographicFront` 是平行投影，没有灭点。** 实测（白板 + 深色底，按阈值分割剪影，
量远边宽度 / 近边宽度）：

| 相机 | 收敛比（远/近） |
| --- | --- |
| `orthographicFront`（COM 的默认） | **1.000** —— 每一个角度都是 1.000 |
| `perspectiveFront` / `perspectiveRelaxed` / `perspectiveAbove` … | **0.784 ~ 0.880** |

**所以只调 RotationX/Y 是得不到透视的**：它们默认走平行投影，远边和近边永远等宽。（补充：COM 其实有开关，见 §27——ThreeD.Perspective = -1。我这次漏测了它，因为成员列表输出被截断。）
平面看起来就是"立着但被压扁"，不是"躺下"。

**正确做法：手写 `prst`，用透视预设。**

```xml
<a:scene3d>
  <a:camera prst="perspectiveFront">
    <a:rot lat="18900000" lon="0" rev="0"/>    <!-- lat 315° = 相机俯角 -45° -->
  </a:camera>
  <a:lightRig rig="threePt" dir="t"/>
</a:scene3d>
```

`lat` 就是俯角，实测：

| lat | 收敛比 | 压扁程度 |
| --- | --- | --- |
| 330°（−30°） | 0.880 | 0.485 |
| **315°（−45°）** | **0.839** | **0.396** |
| 300°（−60°） | 0.805 | 0.280 |
| 285°（−75°） | 0.784 | 0.144 |

**315° 是平衡点**：收敛明显，但没压扁到看不出是一块面。

**怎么判断"躺下了没有"**：不要看"像不像平行四边形"。
**量远边是不是明显比近边窄。** 等宽就是没躺下。

## 26. `a:rot` 的 `lat/lon` 必须是 0~21600000，负数会让**整个文件**损坏

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

## 27. `ThreeD.Perspective` 是**开关**，不是强度——而且 `0` 和 `1` 等价

§25 说"COM 只能出平行投影"，**那一半是错的**。错因是我那次的成员列表输出被截断，
漏掉了 `PresetCamera / Perspective / FieldOfView` 三个成员，于是我漏测了它们。

完整测下来（白板 + 深色底，剪影分割，量远近边比）：

| `ThreeD.Perspective` | 写出的 `prst` | 收敛比 |
| --- | --- | --- |
| `0` | `legacyObliqueFront` | **1.000**（平行） |
| `1` | `legacyObliqueFront` | **1.000**（平行，与 0 完全等价） |
| `-1` | `legacyPerspectiveFront` | **透视** |

`Perspective = -1` 且 `RotationX = -45` 时，**收敛比 0.543** —— 真的产生了灭点。

**三个反直觉点：**

1. **它是开关，不是强度。** 名字叫 Perspective，看着像"透视强度 0~100"，
   实际只接受 `{-3, -1, 0, 1}`，而且 **`0` 和 `1` 都表示关闭**。
2. **要开透视得写 `-1`。** 按名字和直觉都会先试 `1`，那正好是关。
3. **`FieldOfView` 默认 45，但它不控制投影类型**，只影响透视的强度感。

所以「界面能直接做出透视」是成立的 —— 走 `ThreeD.Perspective = -1`。
§25 的结论要按这条修正：**"必须手写 XML"不成立；手写只是更可控**（可以直接选
`perspectiveFront` 等现代预设，而不是被限制在 `legacyPerspectiveFront`）。

## 28. 把"参数 → 渲染结果"当作待测对象，而不是待查文档

这一整轮 3D 效果踩的坑，可以归成一类，值得当成方法论写下来：

**参数被接受 ≠ 参数有效 ≠ 结果如你所想。**

实测到的四种脱节，每一种都让"读文档 → 写参数"的做法失效：

| 类型 | 实例 |
| --- | --- |
| **参数变了，结果没变** | `PresetThreeDFormat` 接受值，但写出的投影不变；五个透视预设渲染完全相同 |
| **参数名和语义相反** | `Perspective = 0` 与 `= 1` 都是"关"，`-1` 才是"开" |
| **两个参数系统轴不对应** | 窗格 `RotationX/Y` 与 OOXML `lat/lon` 不是同一组数（§23） |
| **值域不合法时不是忽略，是拒绝** | `lat="-1800000"` 让 PowerPoint 判整个文件损坏（§26） |

**所以顺序应该是：**

1. **先扫参数空间**，把每个值渲染出来
2. **量一个可测的特征**（不是"看着像"）
3. **建立"参数 → 实测特征"的表**，需要时再回填文档里的语义

**关键在于第二步要选可量化的特征。** 这一轮用过的：

| 特征 | 量什么 | 判什么 |
| --- | --- | --- |
| **收敛比** = 远边宽 / 近边宽 | 剪影上下沿宽度 | 是否透视（1.000 = 平行） |
| **压扁度** = 高 / 中宽 | 剪影包围盒 | 倾角大小 |
| **错切方向** | 左右竖边长度差 | 深度往哪边退 |

**反面教材是我自己**：我先用"像不像平行四边形"挑姿态，而平行投影的矩形**在任何角度
都像平行四边形**，所以这个判据零区分力，直接导致选中了一个根本没躺下的姿态。
**换成"量远边是不是更窄"之后，一次就选中了。**
