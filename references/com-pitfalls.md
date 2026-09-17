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
- 本机没有 `pwsh`，用 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File`。
- 本机没有 `git`/`tar`（`tar` 被策略拦），解压用 .NET 的 `ZipFile`。

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

### 16.4 `Presentation.CreateVideo` 在本机完全不可用

| Quality | 结果 |
| --- | --- |
| 0 | 返回成功码但**不产文件** |
| 1 / 2（文档的 medium/high） | `E_INVALIDARG` |
| VertResolution 480/540/720/1080 | 均失败 |
| ReadOnly / ReadWrite 打开 | 均失败 |

**结论：别指望在本机导 MP4**，用 §16.1–16.3 的图层方案做预览。
（`CreateVideo` 本身是异步的：即使调用成功也要轮询文件，不能只等返回值。）

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


