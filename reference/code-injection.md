# 代码索引 ·注入层

> 上一层：[`CODE_INDEX.md`](../CODE_INDEX.md) · 同层：[探针层](code-probe.md) · [校验层](code-verify.md) · [工具层](code-tooling.md)
>
> **表里的行数与依赖关系由 `tools/gen_code_tables.py` 扫出来，不是手抄的。**
> 重新生成：`python tools/gen_code_tables.py injection`

## 这一层负责什么

把spec 变成pptx 里真的动起来的 OOXML—— 注入引擎本体，和它的浏览器端重放。

**这一层是全库耦合的顶点**，也是唯一一层「改一行可能影响别人」的地方：

| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |
| --- | --- | --- | --- | --- | --- |
| `scripts/motion.py` | dsh-ppt-office-motion :: OOXML motion injection engine | 核心注入引擎：动画/切换/3D/图片填充 + CLI | —（叶子） | `check_coverage`、`design_audit`、`player`、`selftest`、`verify_motion`、`verify_singletons`、`test_camera`、`test_fill_window`、`test_morph`、`test_singletons` | 1914 行 ⚠️ 超 800 |
| `scripts/player.py` | dsh-ppt-office-motion :: build a browser motion preview from a spec | 浏览器端动效重放（HTML） | `motion` | `check_coverage`、`design_audit`、`selftest` | 578 行 |

共 2 个文件，其中 Python 2492 行。

---

## 核心逻辑与数据流

### `motion.py`—— 一次 `apply` 的完整路径

```text
spec (JSON/YAML)
   │
   ▼
normalize_spec()              字段校验 + 补默认值；字段含义同时写在 SKILL.md
   │
   ├──▶ catalog()            查 137 个效果的 presetID 与 XML 模板（只读）
   │
   ▼
open pptx as zip             只有 motion.py 碰 zipfile / lxml
   │
   ├──▶ 动画：改<p:timing>       形状级动画树，挂在幻灯片节点上
   ├──▶ 切换：写<p:transition>   ★ 挂在【终点页】上 —— 写在第 N 页，动的是「进入第 N 页」
   ├──▶ 3D：写<a:sp3d>          相机 presetID + 路径（形态定义在 facts/）
   └──▶ 图片填充：<p:blipFill>  淡化整图铺底 + 原色塞进形状当前景
   │
   ▼
写回zip → 落盘
```

**三条最容易踩的**（都指向 `pitfall-*`，不在这里展开）：

| 症状 | 根因 |
| --- | --- |
| 切换整体错位一个页边界 | `<p:transition>` 挂在终点页 |
| `<p:timing>` 静默消失 | transition 与timing 互相吃掉，见 `probe_timing_roundtrip.ps1` |
| PowerPoint 拒开 `E_FAIL` | 多余 preset 包装层 / schema顺序，见 `pitfall-file-corruption.md` |

### `player.py` —— 浏览器端重放

依赖 `motion.py` 的 `normalize_spec`，把同一份 spec 渲染成 HTML 预览。
**不重新解释 spec 语义** —— 两个入口共用同一个 spec，是有意的：
spec 解释权只留一份，否则预览和真注入会漂。

---

## 改动这一层要过什么

按 [`CODE_INDEX.md` §4.1](../CODE_INDEX.md#41-四道门验证加新能力时想清楚要过哪几道)：

| 改动类型 | 必过的门 |
| --- | --- |
| 几何 / 布局相关 | `motion.py apply --assert-geometry` |
| 动效相关 | `motion.py check` |
| XML 结构相关 | `verify_motion.py` |
| **涉及真渲染的（改切换、3D、media）** | **`motion.ps1 -Strict`** |

第四道门最贵也最重要：**lxml 全绿不代表 PowerPoint 认。**
改完只跑前三道门就收工，是这个项目最容易犯的错。

---

## 已知热点

`motion.py` **1914 行，逼近 2000 行硬上限**，被10 个模块依赖。

**本轮决定：只立规范，不拆。** 理由：
拆它要同时动 10 个依赖者，不是一次改动能验证的；
而现在它的内聚性是好的（所有注入相关的东西都在一处，职责清晰）。

新增逻辑时：

1. 优先放进现有职责区段（标注清楚属于哪一块）
2. 实在放不下，**先看两个拆分样板**（都在 `tools/`，都验证过）：
   - `split_transition_probe.py` —— 3048 行单体 → 多层包（拆**模块**）
   - `split_commands.py` —— 1367 行模块 → 5 个同层文件（拆**一层**）
3. 拆分四条硬约束见 [`CONTRIBUTING.md`](../CONTRIBUTING.md) —— 外部接口不许变、不许改调用方向、一次只拆一层、搬代码前先 grep `__file__` 与相对路径拼装

> **别"顺手"把 `motion ⇄ check_coverage` 的表面环修掉。**
> 那个环是靠 `motion.py check` 里的**函数内延迟导入**刻意打断的，是设计信号。
> 依赖扫描只统计模块级import，就是为了不把它当成违规。

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种动画效果 | 改 `facts/effects.json`，**不要**动 `motion.py` 的模板表 |
| 加一种切换形态 | 改 `facts/transitions.json` + `transition_probe/data.py` |
| 改注入逻辑 | `scripts/motion.py`，先读上面「已知热点」 |
| 改 spec 字段 | `normalize_spec`，字段含义同步写 `SKILL.md` |
| 加一个 CLI 子命令 | `motion.py` 的 CLI 分发 —— **子命令名是接口，改名要改全仓库** |