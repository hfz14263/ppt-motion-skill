# 代码索引 · 校验层

> 上一层：[`CODE_INDEX.md`](../CODE_INDEX.md) · 同层：[注入层](code-injection.md) · [探针层](code-probe.md) · [工具层](code-tooling.md)
>
> **表里的行数与依赖关系由 `tools/gen_code_tables.py` 扫出来，不是手抄的。**
> 重新生成：`python tools/gen_code_tables.py verify`

## 这一层负责什么

**证明「它真的生效了」。** 产出是**结论**：过 / 不过+ 原因。

这一层的存在理由是本项目最核心的一条认识：

> **不报错 ≠ 生效。**
> 本项目最危险的失败模式是**静默丢失** —— lxml 全绿、`verify_motion` 通过、
> 文件能打开，但动效没了或者换了个样子。
> 所以校验层不能只查"结构合法"，必须查"效果还在"。

---

## 速查表

| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |
| --- | --- | --- | --- | --- | --- |
| `scripts/verify_motion.py` | Structural self-check for a motion-injected deck | XML 结构合法性 | `motion` | —（没人依赖） | 125 行 |
| `scripts/verify_singletons.py` | Check for duplicate singletons in a deck, and fail loudly | singleton 约束 | `motion` | —（没人依赖） | 85 行 |
| `scripts/verify_dual_photo.py` | 校验"同图双版本"版式：淡化整图铺底 + 原色塞进形状当前景 | 双色照片方案是否真的有效 | —（叶子） | `test_dual_photo` | 373 行 |
| `scripts/verify_recipes.py` | Validate the recipe library: structure, and the rule that keeps it from rotting | 配方库结构 | —（叶子） | —（没人依赖） | 106 行 |
| `scripts/verify_docs.py` | Validate the documentation structure: references resolve, no orphans, facts unique | 文档结构与登记 | —（叶子） | —（没人依赖） | 237 行 |
| `scripts/check_coverage.py` | dsh-ppt-office-motion :: spec coverage audit | spec 覆盖率审计 | `motion`、`player` | `motion`（函数内延迟导入） | 132 行 |
| `scripts/check_structure.py` | 结构体检：把 CONTRIBUTING.md 的规则变成可执行的检查 | 仓库结构体检（行数/命名/体积/**import 环/热点**） | —（叶子） | —（没人依赖） | 315 行 |
| `scripts/inspect_pptx.py` | inspect_pptx: is this .pptx actually animated, and how? | 看一个 pptx 到底有没有动效 | —（叶子） | —（没人依赖） | 187 行 |
| `scripts/design_audit.py` | Audit a motion spec against the design spec's HARD constraints, and say plainly what it cannot judge | 设计层审计 | `motion`、`player` | —（没人依赖） | 378 行 |
| `scripts/design_compose.py` | Read a design system and compose pages that follow it | 设计层合成 | —（叶子） | —（没人依赖） | 412 行 |
| `scripts/selftest.py` | Regression tests for dsh-ppt-office-motion | 回归测试入口 | `motion`、`player` | —（没人依赖） | 457 行 |
| `scripts/review_assist.py` | dsh-ppt-office-motion :: review assistant | 自动复核 | —（叶子） | —（没人依赖） | 317 行 |
| `tools/scan_deps.py` | 只读：扫模块级 import，得出真实依赖图与环 | 依赖图 / 环检测 / 改动波及面（只读） | —（叶子） | —（没人依赖） | 183 行 |

共 13 个文件，其中 Python 3266 行。

---

## 四类校验，各自证明什么

**别把它们当同一种东西。** 问的问题完全不同：

| 类别 | 文件 | 证明什么 | 不证明什么 |
| --- | --- | --- | --- |
| **结构合法** | `verify_motion.py`、`verify_singletons.py` | XML 符合 schema、没有重复单例 | 不证明 PowerPoint 认|
| **内容对不对** | `inspect_pptx.py`、`check_coverage.py` | 文件里到底有没有动效、spec 覆盖了多少 | 不证明渲染出来好看 |
| **设计合规** | `design_audit.py`、`design_compose.py`、`verify_recipes.py` | 符合设计规范的硬约束 | 不证明 PowerPoint 认 |
| **仓库健康** | `check_structure.py`、`verify_docs.py`、`scan_deps.py` | 结构、命名、登记、依赖方向没烂 | 不证明任何单份pptx |

> **`design_audit.py` 有意在输出里说「我不能判断什么」** ——
> 这不是谦虚，是防止它被当成万能判官。校验工具说自己测不了什么，
> 比让人误以为它测过了更安全。

### 唯一必须过 PowerPoint 的门

上面四类都跑完，**还是不能证明 PowerPoint 认**。

真正的那道门在工具层：`motion.ps1 -Strict` ——
让 PowerPoint 真开、真存、再看效果还在不在。
它是[lxml 测不出来的那一类失败](../reference/pitfall-silent-drop.md)的唯一防线。

---

## `motion ⇄ check_coverage` 那个环

`motion.py check` 子命令用**函数内延迟导入**转发到 `check_coverage.main()`。
所以：

- 依赖图上这里有一个**表面环**
- `scan_deps.py` 只统计**模块级** import，所以它报 `0 处环`—— 这是**正确**的
- 上表`check_coverage` 的「被谁依赖」写作 `motion`（函数内延迟导入），
  就是为了让这件事在文档里留痕，而不是让人以为是漏扫

> **这不是待修的bug，是设计。** 删掉那个延迟导入会让 `motion check` 变成
> 独立脚本，而它作为子命令存在是有理由的。见 `pitfall-tooling.md`。

---

## `check_structure.py` 承担的特殊角色

它是 [`CONTRIBUTING.md`](../CONTRIBUTING.md) 的**执行者**——
规范写在文档里，判定在这个脚本里。

它现在守四类：体积上限、长文档目录、代码内容表、新目录登记，
外加 **模块级 import 环** 与 **热点模块**（被依赖 ≥ 5）。

所以有一条规矩：

> **加一条规范，必须同时改 `check_structure.py`，否则规范会分叉。**
> 文档说「不许超过 800 行」而体检不查，那条规范就是一句空话。

反过来也一样：体检里新增的检查，必须在 `CONTRIBUTING.md` 有对应条目。
两边都不许单方面漂移。

### 环检测为什么只扫模块级

`check_import_cycles()` 只统计 `ast`的 `tree.body`。
**函数内延迟导入不算依赖边** —— 那是**故意打断环**的手法。

这个区分不是洁癖，是本项目踩过的坑：`motion ⇄ check_coverage`
是靠 `motion.py check` 里的函数内导入打断的，**刻意且正确**。
判据不认这个区分，它就会天天误报，然后有人会去"顺手修掉"那个延迟导入——
**那就把设计改坏了**。

依赖图解析复用 `tools/scan_deps.build_graph()`，
**不在这里重写第二份** —— 详见 `CONTRIBUTING.md` §六.8 的真实教训。

---

## 改动这一层要过什么

| 改动类型 | 必过的门 |
| --- | --- |
| 改 `verify_*.py` 的判据 | 拿**已知答案**样本跑一遍（好样本 + 坏样本各一个） |
| 改 `check_structure.py` | 同步 `CONTRIBUTING.md`，两条一起改 |
| 改 `SCAN_DIRS` / 扫描范围 | **`不同步登记 = 该目录完全不被检查**（静默失效） |
| 改 `scan_deps.py` 的解析 | 跑 `gen_code_tables.py`，确认依赖列没变成"全是叶子" |

最后一条是真实踩过的坑：简化的 import 匹配认不出 `import motion` 指的是
`scripts.motion`，整张表的依赖列会全变成"叶子"——
**一张把所有依赖都显示成没有依赖的表，比没有表更坏。**

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种验证 | `scripts/verify_*.py`，先想清楚它在证明什么、不能证明什么 |
| 加一条结构规则 | `check_structure.py` **+** `CONTRIBUTING.md`，一起改 |
| 改文档扫描范围 | `verify_docs.py` 的 `SCAN_DIRS` |
| 想知道真实依赖 | `python tools/scan_deps.py`（只读，放心跑） |
| 生成/刷新代码索引表 | `python tools/gen_code_tables.py [层名]` |