# 切换效果实测对照表

**界面上一个名字 → PowerPoint 真正写进文件的 XML。**

本表不是从 ECMA-376 或 MS-PPTX 抄来的，是**在这台机器的 PowerPoint 上量出来的**。
理由见下面的「为什么必须量」。

- 机器可读：`scripts/transition_reference.json`
- 重新生成：`python scripts/build_transition_table.py`（配合 `scripts/probe_transitions.ps1`
  与 `scripts/probe_enum_scan.ps1`，两者都需要真实 PowerPoint）
- `scripts/motion.py` 的切换模板**直接读这张表**，不再有手写的 XML 字符串。

---

## 1. 为什么必须量，不能查

1. **COM 枚举不可信。** `SlideShowTransition.EntryEffect` 是 2003 年代的枚举，
   本机映射与直觉差得很远 —— `0x0A01` 名曰 fade，实际产出 `<p:strips/>`。
   而且**多个 XML 元素会读回同一个枚举值**（fade / strips 都曾是 0x0A01），
   所以枚举永远无法建立「界面名 → 元素」的映射。
2. **界面分组 ≠ XML 命名空间。** 细微组里混着 core ECMA 元素、p14（显示/闪光）、
   p159（平滑）；华丽组里 2013 年新增的那批全是 p15 `prstTrans`，2010 年那批是 p14；
   动态内容组全是 p14。**想知道属于哪个命名空间，只能写一个候选看 PowerPoint 怎么处理。**
3. **默认值没人写下来。** ECMA 说 `<p:push/>` 可以带 `dir`，但没说作者没碰「效果选项」时
   PowerPoint 到底写什么。省略属性再读回保存后的文件，是知道默认值的唯一诚实办法。
4. **猜错的元素名不会安静失败。** 一个非法的切换子元素会让 PowerPoint **拒开整个文件**
   （同 [`com-pitfalls.md`](com-pitfalls.md) §26）。所以候选必须**一个一份 deck** 去探，
   让「打不开」成为一等公民的结论，而不是被吞掉的异常。

### 两次测量的分工

| 手段 | 做法 | 能回答 | 答不了 |
| --- | --- | --- | --- |
| **枚举扫描** | 一页一个 `PpEntryEffect` 值，用 COM 设上去、保存、读回 | PowerPoint 自己用哪个元素表达哪个值 | 手写这个元素会不会被接受 |
| **往返探测** | 一个候选一份 2 页 deck，打开、保存、比对 | 手写的块是被保留、被改写还是被丢；拒开也算一种结论 | 猜错的名字**正确**写法是什么 |

**往返只能说"有效/拒开"，说不出正确的名字。** 所以 5 个猜错的（切换/翻转/库/摩天轮/传送带）
是靠枚举扫描找回来的：PowerPoint 自己永远写 `dir="l"`，裸写 `<p14:flip/>` 才被拒。

---

## 2. 七条实测规则

这七条比下面那张表更容易被违反，而且每一条都是量出来的，不是读规范读出来的。

1. **命名空间有四种，界面分组不告诉你属于哪种。** 见上。
2. **每个切换最终都在 `mc:AlternateContent` 里 —— 连最普通的 core 切换也是。**
   原因是 `p14:dur`：这是一个 2010 年的属性，PowerPoint 选择把整个 `<p:transition>`
   包起来，而不是把时长丢掉。
3. **`mc:Fallback` 并不总是 `<p:fade/>`。** core 元素在 Fallback 里**重复自己**
   （它本来就向后兼容）；只有 p14/p15/p159 的子元素才降级成 fade。
   给 `<p:push/>` 配一个 `<p:fade/>` 的 Fallback 会被接受，但 PowerPoint 保存时会改掉
   —— 第一批 48 个候选里有 16 个纯粹因为这个被判成「被改写」。
4. **PowerPoint 会删掉值等于默认值的属性。** `<p:push dir="l"/>` 存出来是 `<p:push/>`，
   `<p:split orient="horz" dir="out"/>` 存出来是 `<p:split/>`。写了无害，但保存后就不在了。
5. **切换 / 翻转 / 库 / 摩天轮 / 传送带 缺 `dir` 会被拒开** —— 裸的 `<p14:flip/>`
   让 PowerPoint 拒的是**整份文件**，不是那一页。PowerPoint 自己永远写 `dir="l"`。
6. **命名空间声明可以挂在 `<mc:AlternateContent>`、`<mc:Choice>` 或 `<p:transition>` 上**，
   PowerPoint 会随意搬动它们，还会在 `<mc:Fallback>` 上加一个多余的 `xmlns=""`。
   所以**按字符串比较保存前后的 XML 会报一堆假差异** —— 要比较元素结构。
7. **`<p:sld>` 上重复声明 xmlns 会让整个 part 不合法，PowerPoint 拒开整份 deck。**
   补命名空间时只补缺的那个前缀。

---

## 3. 对照表（48 项，按界面分组与顺序）

`spec` 列是 `scripts/transition_reference.json` 里的键，也是新增 spec 名时用的名字。
`EntryEffect` 列是枚举扫描量到的 COM 复核值 —— **只用于复核**，不要拿它去"同步"。

#### 细微
| spec | 界面名 | English | PowerPoint 写出的子元素 | EntryEffect |
| --- | --- | --- | --- | --- |
| `fade` | 淡入/淡出 | Fade | `<p:fade/>` | 3849 |
| `push` | 推入 | Push | `<p:push/>` | 3853 |
| `wipe` | 擦除 | Wipe | `<p:wipe/>` | 2817 |
| `split` | 分割 | Split | `<p:split/>` | 3585 |
| `reveal` | 显示 | Reveal | `<p14:reveal/>` | 3894 |
| `cut` | 切入 | Cut | `<p:cut/>` | 257 |
| `randombar` | 随机线条 | Random Bars | `<p:randomBar/>` | 2305 |
| `shape` | 形状 | Shape | `<p:circle/>` | 3845 |
| `uncover` | 揭开 | Uncover | `<p:pull/>` | 2049 |
| `cover` | 覆盖 | Cover | `<p:cover/>` | 1281 |
| `flash` | 闪光 | Flash | `<p14:flash/>` | 3909 |
| `morph` | 平滑 | Morph | `<p159:morph option="byObject"/>` | 3954 |

#### 华丽
| spec | 界面名 | English | PowerPoint 写出的子元素 | EntryEffect |
| --- | --- | --- | --- | --- |
| `fall_over` | 跌落 | Fall Over | `<p15:prstTrans prst="fallOver"/>` | 3934 |
| `drape` | 悬挂 | Drape | `<p15:prstTrans prst="drape"/>` | 3936 |
| `curtains` | 帘式 | Curtains | `<p15:prstTrans prst="curtains"/>` | 3938 |
| `wind` | 风 | Wind | `<p15:prstTrans prst="wind"/>` | 3940 |
| `prestige` | 上拉帷幕 | Prestige | `<p15:prstTrans prst="prestige"/>` | 3941 |
| `fracture` | 折断 | Fracture | `<p15:prstTrans prst="fracture"/>` | 3942 |
| `crush` | 压碎 | Crush | `<p15:prstTrans prst="crush"/>` | 3943 |
| `peel_off` | 剥离 | Peel Off | `<p15:prstTrans prst="peelOff"/>` | 3944 |
| `page_curl` | 页面卷曲 | Page Curl | `<p15:prstTrans prst="pageCurlSingle"/>` | 3946 |
| `airplane` | 飞机 | Airplane | `<p15:prstTrans prst="airplane"/>` | 3951 |
| `origami` | 日式折纸 | Origami | `<p15:prstTrans prst="origami"/>` | 3953 |
| `dissolve` | 溶解 | Dissolve | `<p:dissolve/>` | 1537 |
| `checkerboard` | 棋盘 | Checkerboard | `<p:checker/>` | 1025 |
| `blinds` | 百叶窗 | Blinds | `<p:blinds/>` | 769 |
| `clock` | 时钟 | Clock | `<p:wheel spokes="1"/>` | 3857 |
| `ripple` | 涟漪 | Ripple | `<p14:ripple/>` | 3867 |
| `honeycomb` | 蜂巢 | Honeycomb | `<p14:honeycomb/>` | 3898 |
| `glitter` | 闪罐 | Glitter | `<p14:glitter/>` | 3874 |
| `vortex` | 涡流 | Vortex | `<p14:vortex/>` | 3863 |
| `shred` | 碎片 | Shred | `<p14:shred/>` | 3910 |
| `switch` | 切换 | Switch | `<p14:switch dir="l"/>` | 3901 |
| `flip` | 翻转 | Flip | `<p14:flip dir="l"/>` | 3905 |
| `gallery` | 库 | Gallery | `<p14:gallery dir="l"/>` | 3880 |
| `cube` | 立方体 | Cube | `<p14:prism/>` | 3914 |
| `doors` | 门 | Doors | `<p14:doors/>` | 3885 |
| `box` | 框 | Box | `<p:zoom/>` | 3073 |
| `comb` | 梳理 | Comb | `<p:comb/>` | 3847 |
| `zoom2` | 缩放 | Zoom | `<p14:warp/>` | 3889 |
| `random` | 随机 | Random | `<p:random/>` | 513 |

#### 动态内容
| spec | 界面名 | English | PowerPoint 写出的子元素 | EntryEffect |
| --- | --- | --- | --- | --- |
| `pan` | 平移 | Pan | `<p14:pan/>` | 3930 |
| `ferris_wheel` | 摩天轮 | Ferris Wheel | `<p14:ferris dir="l"/>` | 3899 |
| `conveyor` | 传送带 | Conveyor | `<p14:conveyor dir="l"/>` | 3882 |
| `rotate` | 旋转 | Rotate | `<p14:prism isContent="1"/>` | 3918 |
| `window` | 窗口 | Window | `<p14:window/>` | 3887 |
| `orbit` | 轨道 | Orbit | `<p14:prism isContent="1" isInverted="1"/>` | 3926 |
| `fly_through` | 飞过 | Fly Through | `<p14:flythrough/>` | 3890 |

**三组之外的第 13 个细微项是「无」** —— 不写 `<p:transition>`，界面上就是「无」。

**同一元素靠属性区分的**：立方体 / 旋转 / 轨道 三个都是 `<p14:prism>`，
靠 `isContent` 与 `isInverted` 区分；这是**唯一**靠属性区分三个界面项的地方。

---

## 4. 写进文件的完整块

表里的"子元素"只是最里层。要用的完整块在 `scripts/transition_reference.json`
的 `xml` 字段里（`{spd}` 与 `{dur}` 是占位符），形状如下：

```xml
<!-- core 元素：Fallback 重复自己 -->
<mc:AlternateContent xmlns:mc="..." xmlns:p14="...">
  <mc:Choice Requires="p14">
    <p:transition spd="med" p14:dur="800"><p:push dir="u"/></p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition spd="med"><p:push dir="u"/></p:transition></mc:Fallback>
</mc:AlternateContent>

<!-- 扩展元素：Fallback 降级为 fade -->
<mc:AlternateContent xmlns:mc="..." xmlns:p15="...">
  <mc:Choice Requires="p15">
    <p:transition spd="med" p14:dur="800"><p15:prstTrans prst="wind"/></p:transition>
  </mc:Choice>
  <mc:Fallback><p:transition spd="med"><p:fade/></p:transition></mc:Fallback>
</mc:AlternateContent>
```

**块自带命名空间声明**，所以不需要在 `<p:sld>` 根上补 `xmlns:p14`。
位置约束（`<p:transition>` 必须在 `clrMapOvr` 之后、`timing` 之前）见
[`com-pitfalls.md`](com-pitfalls.md) §12。

`mc:Choice` + `mc:Fallback` 会写出**两个** `<p:transition>`，这是合法的、不是重复
（[`com-pitfalls.md`](com-pitfalls.md) §20/§31）；判断"有几个切换"要数**生效的那个**，
即 `motion.py` 的 `active_transition_blocks()`。

---

## 5. 只在 XML 里存在、界面上没有的（7 个）

枚举扫描在 1..4200 上跑完（3956 以上一个都没有），共量到 149 个值。
下面这些 PowerPoint **能写也能读**，但界面库里没有对应项 ——
**不要**为了"用上新功能"去写它们，界面上无法复核，也无法通过「效果选项」调整。

| PowerPoint 写出的子元素 | EntryEffect | 说明 |
| --- | --- | --- |
| `<p:diamond/>` | 3846 | 2003 年代的「菱形」 |
| `<p:newsflash/>` | 3850 | 2003 年代的「新闻快报」 |
| `<p:plus/>` | 3851 | 2003 年代的「加号」 |
| `<p:wedge/>` | 3856 | 2003 年代的「楔入」 |
| `<p14:wheelReverse spokes="1"/>` | 3862 | 时钟的反向变体 |
| `<p15:prstTrans prst="pageCurlDouble"/>` | 3948 | 界面上的「页面卷曲」是 Single |
| `<p15:prstTrans prst="pageCurlDouble" invX="1"/>` | 3949 | 同上，镜像 |

`<p:strips/>`（2561）不在本表里，因为 `motion.py` 的 `strips` spec 用了它 ——
它是历史遗留 spec，不是界面项。

---

## 6. 想加一个新切换时

1. 在上面第 3 节查有没有。**有**就直接用 spec 名。
2. **没有**就别猜元素名 —— 跑一次枚举扫描（`enumdeck` + `probe_enum_scan.ps1` + `enumread`），
   让 PowerPoint 自己写一遍。
3. 拿到元素后跑一次往返（`build` + `probe_transitions.ps1` + `collect`），
   确认**手写**这个块也会被原样保留。
4. 写进 `scripts/build_transition_table.py` 的 `HYPOTHESES`，重新生成表和本文档。

<p align="right"><sub>来源 com-pitfalls §40–§42</sub></p>
