# PPT 动效的总纲领：有，但要经过一道翻译

**结论先说：网上没有"PPT 动效总纲领"这种文件。** 权威纲领全部来自**屏幕 UI 动效**
领域（Material、Carbon、Apple HIG），而它们**大部分不能直接搬到 PPT** ——
不是理念不对，是 **PPT 的动效原语少得多**。

所以这份文档做两件事：**引权威纲领**，然后**逐条标注在 PPT 里能落到什么程度**。

---

## 1. 权威纲领（按可迁移程度排序）

### 1.1 IBM Carbon —— **最可迁移的一份**

[Carbon: Motion](https://carbondesignsystem.com/elements/motion/overview/) 的核心是
**三种缓动语义**，而这三条 PPT 恰好都能表达：

| 缓动 | 定义 | 用在哪 | **PPT 对应** |
| --- | --- | --- | --- |
| **entrance** | 快速出现，**减速停稳** | 弹窗、提示、下拉展开 | 入场动画 + `smooth` |
| **exit** | **加速离开**，暗示"走远了" | 关闭、消失 | 退场动画 |
| **standard** | 全程可见的位移，两端都缓 | 展开、重排 | 动作路径 + `smooth` |

> 例外条款很实用：**"如果元素只是离开视野但还会回来（如侧边栏），用 standard 而不是 exit"**
> —— 因为它要"停在视野外待命"，不该加速。

**两条风格原则**（原文）：

- **productive vs expressive 必须分开。** "把 expressive 留给**少数重要时刻**，
>   以便抓住注意力，并给 productive 体验一个节奏上的停顿。"
- **"严格线性运动在肉眼看来不自然。"** 元素应当**快起、平滑减速**，
>   符合轻质材料的物理。
- **反例**："不要用暗示**弹跳、拉伸或急停**的曲线。"
- **时长随距离/尺寸非线性增长**：动的越多，时长越长。

### 1.2 Material 3

[Material: Easing and duration](https://m3.material.io/styles/motion/easing-and-duration/tokens-specs)
提供**令牌化的时长阶梯**（short / medium / long / extra-long，各 4 档）与
**强调型缓动**（emphasized easing）。**理念可迁移：时长要成套、不要逐页拍脑袋。**

### 1.3 Disney 十二原则

`squash&stretch` / `anticipation` / `follow-through` 等。
**这一份最不适合搬进 PPT** —— 它面向逐帧角色动画，而 PPT 没有关键帧插值。
能借的只有两条：**anticipation**（动之前先有个预备动作）和 **staging**（一次只讲一件事）。

---

## 2. PPT 的能力边界（这一节决定上面哪些能用）

| 纲领里的概念 | PPT 能不能做 | 说明 |
| --- | --- | --- |
| 缓动曲线（cubic-bezier） | ❌ **不能** | PPT 只有 `smooth`（加速/减速的粗略开关），**没有曲线控制** |
| 精确时长 | ✅ 能 | 每个效果可设时长 |
| 延迟 / 错峰（stagger） | ✅ 能 | `delay`，是本 skill 最有力的手段 |
| 触发（click/with/after） | ✅ 能 | 决定"观众等不等" |
| 动作路径 | ✅ 能 | 相当于 standard easing 的位移 |
| 属性插值（宽高/透明度补间） | ⚠️ 部分 | 只有特定效果：`growTurn`、`fade`、`zoom`… |
| 补间任意属性（Morph） | ⚠️ 见下 | `p159:morph` 可用，但只在**切页**之间 |
| clip-path / mask 动画 | ❌ 不能 | **但"形状切割图片"可以做出等价效果**（见配方 §8） |
| 滚动联动 / 视差 | ❌ 不能 | 无滚动概念 |

### 因此结论是反直觉的

**PPT 的动效上限不在"效果多花哨"，而在"编排"。**
既然**没有真正的缓动曲线**，那么质感就只能来自：

1. **时机**（delay / trigger / 时长阶梯）
2. **顺序**（先动谁、后动谁）
3. **幅度**（动的距离/尺寸变化量）

**换句话说：PPT 动效的质量 ≈ 编排质量，而不是效果选择。**
这一条解释了一个常见现象：**同一套 `fade`+`wipe`，编排好的是高级感，编排差的是廉价感。**

---

## 3. 落到本 skill 的规则

`references/motion-design-spec.md` 已经把上面这些量化了。**对照 Carbon 后要补三条**：

1. **进/出场用不同语义**（Carbon 的核心）：
   入场用**减速停稳**（`smooth` 高、时长中），退场用**加速**（时长短）。
   本 skill 目前只写了时长阶梯，**没有区分进出场的缓动语义** —— 见第 4 节待办。
2. **expressive 要稀缺**。每份 deck 只允许少数几个"重要时刻"用 T2/T3 时长，
   其余全部 T1。**一页上不允许有两个元素同时在 T2。**
3. **禁止弹跳/拉伸/急停**。这一条本 skill 已在慎用清单里列了
   （`bounce` `boomerang` `pinwheel`…），但**当时没有依据**；现在有：
   Carbon 明确说这类曲线"不自然、分散注意、纯装饰"。

---

## 4. 还没做到的（诚实列出）

- **进/出场缓动语义分离** —— spec 有 `smooth` 字段，但没有"入场用高 smooth / 退场用短时长"
  这样的**语义层**。要在 `motion.py` 里加成默认值。
- **非线性时长** —— Carbon 按移动距离算时长。本 skill 的时长是**手写**的，
  没有"距离越长时间越长"的自动规则。
- **权威纲领的本地化** —— 本文档是摘要，没有逐条把 Carbon/Material 的令牌表抄进来。

---

## 5. 来源

| 内容 | 来源 |
| --- | --- |
| 三种缓动语义、productive/expressive、禁弹跳 | [IBM Carbon — Motion](https://carbondesignsystem.com/elements/motion/overview/)（[原始 mdx](https://raw.githubusercontent.com/carbon-design-system/carbon-website/e8d67b412ea443dec338e12f5bc3f314be385aac/src/pages/guidelines/motion/basics.mdx)） |
| 时长令牌阶梯、emphasized easing | [Material 3 — Easing and duration](https://m3.material.io/styles/motion/easing-and-duration/tokens-specs)、[Material 1 — Duration & easing](https://m1.material.io/motion/duration-easing.html) |
| 十二原则 | Disney，经动画业界通用转述 |
| **PPT 能力边界与"编排即质感"的判断** | **本机实测**（见 `com-pitfalls.md`、`camera-reference.md`） |
