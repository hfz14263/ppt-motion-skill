# 踩坑 · COM 与脚本环境

> COM 与脚本环境 —— 界面层的约束、编码、真渲染
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

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
