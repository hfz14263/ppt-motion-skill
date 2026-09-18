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

**先读** [`references/recipe-library.md`](references/recipe-library.md)
—— 6 条**原理**（不是效果清单）。新素材从原理推导，不查表。

| 想知道 | 去哪 |
| --- | --- |
| 有哪些**原理**可套 | [`recipe-library.md`](references/recipe-library.md) |
| **时长/节奏/密度**的量化规范 | [`motion-design-spec.md`](references/motion-design-spec.md) |
| **权威纲领**（Carbon/Material）与 PPT 能力边界 | [`motion-principles.md`](references/motion-principles.md) |
| 现成模板**从零设计**动效 | [`template-patterns.md`](references/template-patterns.md) |
| **效果→语义**映射（该用哪个效果） | [`motion-design-spec.md`](references/motion-design-spec.md) §3 |

---

## ② COM 操作技术 —— 界面能做什么

**先读** [`reference/capabilities.md`](reference/capabilities.md)
—— 哪些事只有 COM 能做、哪些 COM 做不到，以及**为什么"查不到"不等于"不支持"**。

| 想知道 | 去哪 |
| --- | --- |
| COM 能做什么 / 不能做什么 | [`reference/capabilities.md`](reference/capabilities.md) |
| 无头操作、真渲染、导出 | [`SKILL.md`](SKILL.md)（`motion.ps1` 层） |
| 媒体（音视频）怎么嵌入 | [`SKILL.md`](SKILL.md) 媒体层 |
| 为什么属性模型和 XML 是**两条路** | [`com-pitfalls.md`](references/com-pitfalls.md) 「当时的误解」段 |

---

## ③ OOXML 注入技术 —— 怎么写进文件

**先读** [`SKILL.md`](SKILL.md) §3（spec 字段全表）。
**写注入代码前必读** [`authoring-rules.md`](references/authoring-rules.md)（13 条硬约束）。

| 想知道 | 去哪 |
| --- | --- |
| spec 怎么写（YAML 全表） | [`SKILL.md`](SKILL.md) §3 |
| **元素顺序 / 命名空间**的硬约束 | [`authoring-rules.md`](references/authoring-rules.md) |
| Morph / 3D / 窗口填充的**写法** | [`morph-and-3d-recipes.md`](references/morph-and-3d-recipes.md) |
| 效果别名与 presetID 对照 | [`facts/symbols.json`](facts/symbols.json) |
| **一个元素只能出现一次**（本项目所有损坏的根因） | [`com-pitfalls.md`](references/com-pitfalls.md) §31 |

---

## ④ 验证与判据 —— 怎么知道对不对

**先读** [`review-checklist.md`](references/review-checklist.md) §0
—— **判据可信度分级**：哪条判据已知会骗人。

| 想知道 | 去哪 |
| --- | --- |
| 哪些判据**不可信** | [`review-checklist.md`](references/review-checklist.md) §0 |
| 交付前的完整清单 | [`review-checklist.md`](references/review-checklist.md) |
| 动效规范审计（机器可判的那半） | `scripts/design_audit.py` |
| 重复单例检查（**所有损坏的根因**） | `scripts/verify_singletons.py` |
| 结构 + 几何自证 | `scripts/verify_motion.py` |
| 从视频量动效的**能力边界** | [`video-analysis-limits.md`](references/video-analysis-limits.md) |

---

## ⑤ 版面与风格（**静态**）—— 长什么样

**先读** [`design-system/README.md`](references/design-system/README.md)
—— 10 套内置 + 34 套按需抓取。**先按「分类 + 适用」挑一套，再读全文。**

| 想知道 | 去哪 |
| --- | --- |
| 挑哪一套风格 | [`design-system/README.md`](references/design-system/README.md) |
| 让设计系统**可执行**（解析调色板生成页面） | `scripts/design_compose.py` |
| 与 `dsh-ppt-studio` 的对接契约 | [`ppt-studio-integration.md`](references/ppt-studio-integration.md) |
| 为什么这样分层 | [`architecture.md`](references/architecture.md) |

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
| 参数**实测对照表**（相机） | [`camera-reference.md`](references/camera-reference.md) |
| 效果**别名 ↔ presetID** | [`facts/symbols.json`](facts/symbols.json) |
| **界面角度 ↔ lat/lon** | [`facts/axes.json`](facts/axes.json) |
| 效果与切换的**原始语义** | [`mso-primitives.md`](references/mso-primitives.md) |

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
