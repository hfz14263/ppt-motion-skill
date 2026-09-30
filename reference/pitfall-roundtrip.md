# 踩坑 · 往返后被改掉或吃掉

> 往返（round-trip）后被改掉或吃掉的效果
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

往返是「PowerPoint 帮我重写一遍文件」。它能救活你写不对的东西，也会顺手改掉你写对的东西 —— 两个方向都会咬人。

---

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
什么都折叠不了，版面看起来完全一样。见 `reference/authoring-rules.md` §J。
