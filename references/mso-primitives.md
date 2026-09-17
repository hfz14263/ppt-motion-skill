# 动画效果索引与切换语义

## 1. 查效果:用 CLI,不要查文档

```bash
python scripts/motion.py catalog                 # 137 个别名全表
python scripts/motion.py catalog --kind path     # 只看动作路径
python scripts/motion.py catalog --json          # 机器可读
```

CLI 的输出来自 `scripts/motion_catalog.json`,而那个文件里的
`presetID / presetClass / presetSubtype` 是**从真实 PowerPoint 输出提取**的
(用 COM 逐个 `AddEffect` 后回读 XML),不是按枚举值猜的。**它是唯一权威。**

> 本文档以前内嵌了同一份 137 行表格,那是纯复制——CLI 一改文档就漂。
> 现在表格删掉了,只留 CLI 表达不了的东西。

---

## 2. `enum` ≠ `presetID`(这是本文件存在的首要理由)

`catalog` 打出两列:`enum`(MsoAnimEffect 枚举值)和 `presetID`(写进 XML 的那个)。
**多数别名两者相同,但动作路径和强调全部不同**,因为枚举把类别也编进了值:

| 类别 | enum 起止 | presetID 起止 | 说明 |
| --- | --- | --- | --- |
| 入场 | 1 | 1 | 一致 |
| 强调 | 54 | 1 | **偏移 53**:`spin` 是 enum 61 / presetID 8 |
| 动作路径 | 86 | 1 | **偏移 85**:`pathCircle` 是 enum 86 / presetID 1 |

**所以注入必须用 `presetID`。** 你从微软文档或 VBA 里抄到的枚举值,不能直接写进
OOXML —— 那会让你得到完全不同的效果,而且结构校验全过。

`catalog` 会同时打出 `enum`、`enumName` 和 `presetID`,对照着看就不会错。

---

## 3. 未收录的枚举值

这些 MsoAnimEffect 有值,但**没有提取到模板**,写进 spec 会报
`unknown effect alias`。其中媒体类那几个不是"动画",要走
`scripts/motion.ps1` 的媒体路径,不要当效果注入。

| 枚举 | 名称 | 备注 |
| --- | --- | --- |
| 0 | msoAnimEffectCustom | 自定义,无固定模板 |
| 24 | msoAnimEffectRandomEffects | |
| 25 | msoAnimEffectBoomerang | |
| 27 | msoAnimEffectColorReveal | |
| 29 | msoAnimEffectEaseIn | |
| 36 | msoAnimEffectThinLine | |
| 41 | msoAnimEffectFadedSwivel | |
| 45 | msoAnimEffectStretchy | |
| 71 | msoAnimEffectComplementaryColor2 | |
| 79 | msoAnimEffectStyleEmphasis | |
| 83 / 84 / 85 / 150 | msoAnimEffectMediaPlay / Pause / Stop / PlayFromBookmark | **走媒体层,不是效果** |

要补某个:把枚举值加进 `scripts/alias_probe.ps1` 的映射表,重跑提取。

---

## 4. 触发与时序

| spec 写法 | PowerPoint `nodeType` | 语义 |
| --- | --- | --- |
| `trigger: click` | `clickEffect` | 点击触发 |
| `trigger: with` | `withEffect` | 与上一个**同时**开始 |
| `trigger: after` | `afterEffect` | 上一个**整组**结束后（默认） |

⚠️ **`with` 有三个反直觉之处**，都会导致动效和预期不符:

1. **`with` 不能作为一组的第一个效果。** 它要和"上一个"同时,前面没有就退化成
   `after`。把 `with` 放在首位,等于写错了触发方式而不报错。
2. **`after` 等的是整组的最大值,不是最后一个效果的结束。** 一个 `with` 效果可以比
   它伴随的那个更长,后面的 `after` 必须等更慢的那个。`motion.py` 的
   `schedule_spec()` 就是按这个规则算时间的。
3. **`after` 之后跟 `with`,那组的时间戳回退到组的起点。** 这是"一次出现多个元素"
   的正确写法,但很容易误读成"它们各自延时"。

其余时序字段:`delay`（秒，写到外层 `stCondLst`）、`duration`（秒，写到叶子行为
`cTn` 的毫秒）、`repeat`（`repeatCount = 次数×1000`）、`autoReverse`（`autoRev="1"`）、
`smooth`（0–1，同时设 `accel` 和 `decel`）。

---

## 5. 切换效果

**以写出的 `<p:transition>` 元素为准。** COM 的
`SlideShowTransition.EntryEffect` 枚举在本机与元素**不一致**
（`0x0A01` 实际产出 `<p:strips/>`），所以它只用于复核读数，**不要用它去"同步"**。

| spec 名 | 写出元素 | COM 复核枚举 |
| --- | --- | --- |
| `fade` | `<p:fade/>` | 0x0A01 |
| `fadeblack` | `<p:fade thruBlk="1"/>` | 0x0B01 |
| `push` | `<p:push dir="u"/>` | 0x0901 |
| `pushleft` | `<p:push dir="l"/>` | 0x0901 |
| `wipe` | `<p:wipe dir="l"/>` | 0x0801 |
| `cover` | `<p:cover dir="l"/>` | 0x0C01 |
| `split` | `<p:split orient="horz" dir="out"/>` | 0x0701 |
| `zoom` | `<p:zoom dir="in"/>` | 0x0C01 |
| `dissolve` | `<p:dissolve/>` | 0x0D01 |
| `strips` | `<p:strips/>` | 0x0A01 |
| `pull` | `<p:pull/>` | 0x0801 |
| `randombar` | `<p:randomBar/>` | 0x0901 |

`<p:transition>` 还有两个前缀:**`mc:AlternateContent` 会让 PowerPoint 为一次切换
写两个 `<p:transition>`**(Choice + Fallback),这是合法的,不是重复。
位置约束见 [`com-pitfalls.md`](com-pitfalls.md) §12。
