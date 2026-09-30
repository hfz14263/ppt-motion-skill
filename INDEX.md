# 索引

**六个入口,按"想做什么"分。** 每条都能填进"我想 ____"这个句子 ——
填不进去的东西属于 L2/L3,不属于这里。

```
我想  设计动效        → ① 动效设计知识
我想  用界面做（COM）  → ② COM 操作技术
我想  写进文件        → ③ OOXML 注入技术
我想  验证对不对      → ④ 验证与判据
我想  定版面长相      → ⑤ 版面与风格（静态）
我想  查一个现象      → ⑥ 参考索引（症状 / 参数 / 符号）
```

**一条边界判据**：**⑤ 只管静态版面**（色板、排版、密度）。
**动效一律归 ①**。上游设计系统自己写着 *static layout only; motion governed by
motion-design-spec*，这里沿用。

---

## ① 动效设计知识 —— 该怎么动

**先读** [`reference/recipe-library.md`](reference/recipe-library.md)
—— 6 条**原理**（不是效果清单）。新素材从原理推导，不查表。

| 想知道 | 去哪 |
| --- | --- |
| 有哪些**原理**可套 | [`recipe-library.md`](reference/recipe-library.md) |
| **时长/节奏/密度**的量化规范 | [`motion-design-spec.md`](reference/motion-design-spec.md) |
| **权威纲领**（Carbon/Material）与 PPT 能力边界 | [`motion-principles.md`](reference/motion-principles.md) |
| 现成模板**从零设计**动效 | [`template-patterns.md`](reference/template-patterns.md) |
| **效果→语义**映射（该用哪个效果） | [`motion-design-spec.md`](reference/motion-design-spec.md) §3 |

---

## ② COM 操作技术 —— 界面能做什么

**先读** [`reference/capabilities.md`](reference/capabilities.md)
—— 哪些事只有 COM 能做、哪些 COM 做不到，以及**为什么"查不到"不等于"不支持"**。

| 想知道 | 去哪 |
| --- | --- |
| COM 能做什么 / 不能做什么 | [`reference/capabilities.md`](reference/capabilities.md) |
| 无头操作、真渲染、导出 | [`SKILL.md`](SKILL.md)（`motion.ps1` 层） |
| 媒体（音视频）怎么嵌入 | [`SKILL.md`](SKILL.md) 媒体层 |
| 为什么属性模型和 XML 是**两条路** | [`com-pitfalls.md`](reference/com-pitfalls.md) 「当时的误解」段 |
| COM / 脚本环境的坑 | [`pitfall-com.md`](reference/pitfall-com.md) |

---

## ③ OOXML 注入技术 —— 怎么写进文件

**先读** [`SKILL.md`](SKILL.md) §3（spec 字段全表）。
**写注入代码前必读** [`authoring-rules.md`](reference/authoring-rules.md)（13 条硬约束）。

| 想知道 | 去哪 |
| --- | --- |
| spec 怎么写（YAML 全表） | [`SKILL.md`](SKILL.md) §3 |
| **元素顺序 / 命名空间**的硬约束 | [`authoring-rules.md`](reference/authoring-rules.md) |
| Morph / 3D / 窗口填充的**写法** | [`morph-and-3d-recipes.md`](reference/morph-and-3d-recipes.md) |
| 效果别名与 presetID 对照 | [`facts/symbols.json`](facts/symbols.json) |
| **一个元素只能出现一次**（本项目所有损坏的根因） | [`com-pitfalls.md`](reference/com-pitfalls.md) §31 |
---

## ④ 验证与判据 —— 怎么知道对不对

**先读** [`review-checklist.md`](reference/review-checklist.md) §0
—— **判据可信度分级**：哪条判据已知会骗人。

| 想知道 | 去哪 |
| --- | --- |
| 哪些判据**不可信** | [`review-checklist.md`](reference/review-checklist.md) §0 |
| 交付前的完整清单 | [`review-checklist.md`](reference/review-checklist.md) |
| 动效规范审计（机器可判的那半） | `scripts/design_audit.py` |
| 重复单例检查（**所有损坏的根因**） | `scripts/verify_singletons.py` |
| 结构 + 几何自证 | `scripts/verify_motion.py` |
| 「同图双版本」版式校验（底图/前景是否同图、前景是否被拉伸） | `scripts/verify_dual_photo.py` |
| 从视频量动效的**能力边界** | [`video-analysis-limits.md`](reference/video-analysis-limits.md) |

---

## ⑤ 版面与风格（**静态**）—— 长什么样

**先读** [`design-system/README.md`](reference/design-system/README.md)
—— 10 套内置 + 34 套按需抓取。**先按「分类 + 适用」挑一套，再读全文。**

| 想知道 | 去哪 |
| --- | --- |
| 挑哪一套风格 | [`design-system/README.md`](reference/design-system/README.md) |
| 让设计系统**可执行**（解析调色板生成页面） | `scripts/design_compose.py` |
| 造「同图双版本」样例（含故意拉胖的负对照） | `scripts/build_dual_photo.py` |
| 与 `dsh-ppt-studio` 的对接契约 | [`ppt-studio-integration.md`](reference/ppt-studio-integration.md) |
| 为什么这样分层 | [`architecture.md`](reference/architecture.md) |

**禁止**：混搭多套风格（上游明确要求）。

---

## ⑥ 参考索引 —— 查一个现象

**遇到问题时，你知道的是"现象"，不是"根因"。所以这一栏按现象查。**

| 我看到 | 去哪 |
| --- | --- |
| **文件打不开 / 报损坏** | [`reference/symptoms.md`](reference/symptoms.md) → 「文件损坏」 |
| **值越界**（角度、百分比） | [`reference/symptoms.md`](reference/symptoms.md) → 「值越界」 |
| **XML 里有，PowerPoint 不认** | [`reference/symptoms.md`](reference/symptoms.md) → 「静默丢弃」 |
| **文件正常但结果不对** | [`reference/symptoms.md`](reference/symptoms.md) → 「结果不对」 |
| **效果被吃掉**（往返后丢失） | [`reference/symptoms.md`](reference/symptoms.md) → 「往返丢失」 |
| 参数**实测对照表**（相机） | [`camera-reference.md`](reference/camera-reference.md) |
| **页面切换**：48 项实测对照 / 机制怎么运作 / 看起来在做什么 / 该用哪个 | [`transitions.md`](reference/transitions.md) · [`transition-model.md`](reference/transition-model.md) · [`transition-shapes.md`](reference/transition-shapes.md) · [`transition-choice.md`](reference/transition-choice.md) |
| 效果**别名 ↔ presetID** | [`facts/symbols.json`](facts/symbols.json) |
| **界面角度 ↔ lat/lon** | [`facts/axes.json`](facts/axes.json) |
| 效果与切换的**原始语义** | [`mso-primitives.md`](reference/mso-primitives.md) |

> **症状与踩坑这两族都拆成了多份**，各有目录页分流 ——
> 症状见 [`symptoms.md`](reference/symptoms.md)，踩坑见
> [`com-pitfalls.md`](reference/com-pitfalls.md)。两份目录页都能在
> 一屏内选完，所以这里不再重复它们的表（重复的表迟早会漏更新）。
> 看到 `§n` 形式的引用，查 [`pitfall-map.md`](reference/pitfall-map.md)。

### 踩坑手册（已按主题拆成 10 份）

**入口是 [`com-pitfalls.md`](reference/com-pitfalls.md)**（目录页，按"你现在的处境"分流）。
看到 `§n` 这种引用、想知道它在哪份文件里，查
[`pitfall-map.md`](reference/pitfall-map.md)。

| 处境 | 去哪 |
| --- | --- |
| 文件打不开 / 报损坏 | [`pitfall-file-corruption.md`](reference/pitfall-file-corruption.md) |
| XML 里有、PowerPoint 不认（不报错） | [`pitfall-silent-drop.md`](reference/pitfall-silent-drop.md) |
| 往返后效果没了 / 被改了 | [`pitfall-roundtrip.md`](reference/pitfall-roundtrip.md) |
| 注入写法（元素顺序、属性、前缀） | [`pitfall-ooxml.md`](reference/pitfall-ooxml.md) |
| COM / 脚本环境 | [`pitfall-com.md`](reference/pitfall-com.md) |
| Morph / 3D / 图片填充 | [`pitfall-morph-3d.md`](reference/pitfall-morph-3d.md) |
| 测量与判据（指标骗人、口径错） | [`pitfall-measurement.md`](reference/pitfall-measurement.md) |
| 素材与观感（截图/照片、抠图） | [`pitfall-media.md`](reference/pitfall-media.md) |
| 页面切换 | [`pitfall-transition.md`](reference/pitfall-transition.md) |
| 工具链（推送、导出、视频） | [`pitfall-tooling.md`](reference/pitfall-tooling.md) |
| `§n` 到底在哪份文件 | [`pitfall-map.md`](reference/pitfall-map.md) |

---

## 阅读路径（如果你想从头理解）

不是按文件，是按**问题在长什么样**：

```
第 1 段  给一份现成 pptx 加动画
         → SKILL.md 的 S0–S8 流程 → ④ 验证
第 2 段  搞清楚"界面能做什么、XML 要写什么"
         → ② COM 操作技术 → ③ OOXML 注入技术 → ⑥ 查现象
第 3 段  从"能加动画"到"动得好看"
         → ① 动效设计知识 → ⑤ 挑一套静态风格
```
