# 症状 · 往返后效果被吃掉

> PowerPoint 打开再保存一遍（round-trip），你的效果就变了或没了。注意：往返**既会救活**你写不对的东西，**也会改掉**你写对的东西。
>
> 本页按**现象**写，不按发现顺序。每节结尾标着它对应的踩坑编号
> （`来源 com-pitfalls §n`）—— 编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 目录见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

---

## 往返后效果被吃掉

注入时好好的，**PowerPoint 打开再保存之后就没了** —— 因为 PowerPoint 会按自己的理解重写 XML。判据：`motion.ps1 -Strict` 往返普查。

### §14 同一形状上"入场 + 强调"不兼容 PowerPoint 往返（实测限制，未解决）

**现象**：同一形状上同时放入场和强调，PowerPoint 往返后其中一个效果不见了

**根因与修法**：见 [`pitfall-roundtrip.md`](pitfall-roundtrip.md) §14。

<p align="right"><sub>来源 com-pitfalls §14</sub></p>

### §18 Round-trip 会以第二种方式咬"一个形状多个效果"

**现象**：修好一个效果，另一个效果在往返后消失了

**根因与修法**：见 [`pitfall-roundtrip.md`](pitfall-roundtrip.md) §18。

<p align="right"><sub>来源 com-pitfalls §18</sub></p>

### §21 3D 相机（`scene3d`）可以被注入，且角度原样保留

**现象**：一直认为 3D 相机无法注入

**根因与修法**：见 [`pitfall-morph-3d.md`](pitfall-morph-3d.md) §21。

<p align="right"><sub>来源 com-pitfalls §21</sub></p>

### §41 core 切换的 Fallback 被 PowerPoint 改掉

**现象**：core 切换的 Fallback 被 PowerPoint 改掉了

16 个 core 切换往返后全被判「被改写」，但子元素一个都没变。原因在 Fallback：
**core 元素本来就向后兼容，PowerPoint 要求 Fallback 重复它自己**，
只有 p14/p15/p159 的子元素才降级成 `<p:fade/>`。

```xml
<mc:Fallback><p:transition><p:fade/></p:transition></mc:Fallback>   <!-- 写的 -->
<mc:Fallback><p:transition><p:push dir="u"/></p:transition></mc:Fallback>  <!-- 保存后 -->
```

同源两条：**每个切换都会被包进 `mc:AlternateContent`**（连 core 的也是，因为
`p14:dur` 是 2010 年属性）；**值等于默认值的属性会被删掉**（`<p:push dir="l"/>` →
`<p:push/>`）。

完整 48 项对照表见 [`transitions.md`](transitions.md)。

**根因与修法**：见 [`pitfall-silent-drop.md`](pitfall-silent-drop.md) §41。

<p align="right"><sub>来源 com-pitfalls §41</sub></p>
