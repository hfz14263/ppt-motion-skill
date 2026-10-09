# 代码地图（CODE_INDEX）

**这是代码的地图，不是代码的说明书。**

它回答四个问题，每个问题对应下面一节：

| 你想知道 | 看哪节 |
| --- | --- |
| 这个项目分几层、依赖往哪走 | [§1 分层依赖图](#1-分层依赖图) |
| 有哪些入口、从哪跑起 | [§2 入口清单](#2-入口清单) |
| 某个文件是干什么的 | [§3 模块速查表](#3-模块速查表) |
| **我要做 X，该改哪个文件** | [§4 扩展点](#4-扩展点我要做-x-改哪个文件) |

**不回答**函数的实现细节 —— 那在代码里；
**不回答** OOXML / COM / 排版知识 —— 那在 [`reference/architecture.md`](reference/architecture.md)；
**不回答**该怎么改代码 —— 那在 [`CONTRIBUTING.md`](CONTRIBUTING.md)。

> **改这个文件前先查 §4**，别凭文件名猜。
> 本索引的数字（行数、依赖数）由 `tests/test_code_index.py` 守着，
> 代码变了它会报「索引已过期」—— 所以这份文档不会悄悄腐烂。

---

## 1. 分层依赖图

依赖**只能从上往下**。下层不许反向 import 上层；同层之间尽量不 import。

```text
Layer 3  入口
         SKILL.md 描述的 S0–S8 流程 · install.ps1
              │
              ▼
Layer 2  校验与探针
         verify_motion.py · verify_singletons.py · check_coverage.py
         verify_docs.py · check_structure.py · verify_recipes.py
         verify_dual_photo.py
         transition_probe/（common → data → decks → analysis → commands）
         build_*.py · analyze_video.py · design_audit.py · design_compose.py
              │
              ▼
Layer 1  基础设施
         motion.py（1914 行 · 被 10 个模块依赖）· player.py
              │
              ▼
Layer 0  数据（无本地依赖）
         motion_catalog.json · transition_reference.json · facts/*.json
```

**实测：模块级 import 环 = 0 处。**
全库共 **58 个 Python 模块**（`scripts/` 22+6 · `tools/` 17 · `tests/` 11 + 其他），
另有 **9 个 `.ps1`** 不进 AST 图（它们被命令行调用，不是被 import）。

唯一被设计出来的**表面环**是 `motion ⇄ check_coverage`：
`motion.py check` 子命令用**函数内延迟导入**转发到 `check_coverage.main()`。
这是**刻意打断环**的手法，不是违规 —— 所以依赖扫描只统计模块级 import。

> 想知道谁依赖谁、被依赖多少：跑 `python tools/scan_deps.py`（只读）。

---

## 2. 入口清单

### 2.1 命令行入口（用户在终端敲的）

| 命令 | 干什么 | 备注 |
| --- | --- | --- |
| `motion.py apply` | **注入动效**：动画 / 切换 / 3D / 图片填充 | ★ 核心入口 |
| `motion.py inspect` | 看一个 pptx 到底有没有动画 | 只读 |
| `motion.py preview` | 导出无timing、无 transition 的干净副本 | 静态导出用 |
| `motion.py catalog` | 查 137 个效果的 presetID 与 XML 模板 | 只读 |
| `motion.py player` | 生成浏览器端动效重放 | 依赖 `player.py` |
| `motion.py check` | spec 覆盖率审计 | 转发到 `check_coverage.py` |
| `build_transition_table.py` | 切换实测，**19 个子命令** | 入口转发，实体在 `transition_probe/` |
| `selftest.py` | 回归测试 | |
| `review_assist.py` | 自动复核（结构可信 + 启发式标注） | |
| `install.ps1` | 安装到 DSH skill 目录 + 环境预检 | |

### 2.2 编程入口（被别的模块 import）

| 模块 | 谁用它 | 稳定性承诺 |
| --- | --- | --- |
| `motion` | **10 个模块** | 稳定。函数名、CLI 子命令名、退出码都不许变 |
| `player` | 3 个模块 | 稳定 |
| `transition_probe` | 仅 `build_transition_table.py` | 稳定 |
| `motion_catalog.json` | `motion.py` + `transition_probe` | 数据契约，改动要同步 `facts/` |

### 2.3 文档入口

| 想知道 | 去哪 |
| --- | --- |
| 端到端怎么用（spec 全文、命令、坑） | [`SKILL.md`](SKILL.md) |
| 文档总入口 | [`INDEX.md`](INDEX.md) |
| 代码怎么改 | [`CONTRIBUTING.md`](CONTRIBUTING.md) |
| 长期交接与未完成事项 | [`HANDOVER.md`](HANDOVER.md) |
| 已做完的事 + 当时怎么想错的 | [`history/`](history/) |

---

## 3. 模块速查表

按功能域分四份，**每份都是同一个模板**。点文件看详细版：

| 分层 | 文件 | 装什么 | 详细版 |
| --- | --- | --- | --- |
| 注入层 | `motion.py` / `player.py` | 改 OOXML、生成 HTML 重放 | [`reference/code-injection.md`](reference/code-injection.md) |
| 探针层 | `transition_probe/` / `build_*` / `analyze_video.py` | 造 deck、测参数、量帧 | [`reference/code-probe.md`](reference/code-probe.md) |
| 校验层 | `verify_*` / `check_*` | 证明「它真的生效了」 | [`reference/code-verify.md`](reference/code-verify.md) |
| 工具层 | `push_via_api.py` / `vendor_themes.py` 等 | 不产动效，但项目运转要靠 | [`reference/code-tooling.md`](reference/code-tooling.md) |

**为什么分四份而不是一张表**：`scripts/` 有 22 个 `.py` + 9 个 `.ps1` + 2 个 `.json` = **33 个文件**。
33 行 × 7 列的表一屏放不下，必须滚动 —— 所以按功能域分层。
分层阈值是 `文件数 > 25`。

---

## 4. 扩展点：我要做 X，改哪个文件

**这一节是本索引存在的核心理由。** 只说「有什么」，读者还是得自己猜改哪儿。

| 我要做的事 | 改这里 | 注意 |
| --- | --- | --- |
| 加一种动画效果 | `facts/effects.json` → 再确认 `scripts/motion_catalog.json` | 数据驱动；**不要**手改 `motion.py` 的模板表 |
| 加一个切换形态 | `facts/transitions.json` + `scripts/transition_probe/data.py` | 形态定义走数据，实测走探针 |
| 改注入逻辑（写 OOXML） | `scripts/motion.py` | 已 1914 行，**接近 2000 硬上限**；新逻辑优先放进现有职责区段，实在放不下才拆 |
| 加一个 CLI 子命令 | `scripts/motion.py` 的CLI 分发 | 子命令名是接口，**改名要改全仓库** |
| 改 spec 字段 | `scripts/motion.py` 的 `normalize_spec` | 字段含义同时写在 `SKILL.md` |
| 加一种验证 | `scripts/verify_*.py` | 四道门各有位置，见下|
| 加一条体检规则 | `scripts/check_structure.py` | **必须同时更新 `CONTRIBUTING.md`**，否则体检和文档会分叉 |
| 改文档结构 / 加文档 | `scripts/verify_docs.py` 的 `SCAN_DIRS` | **不同步登记 = 该目录完全不被检查** |
| 加探针实测 | `scripts/transition_probe/` | 严格分层：`common → data → decks → analysis → commands` |
| 改页面切换机制 | `reference/transition-model.md` | 机制层文档是**唯一真相源**，`facts/transitions.json` 声明它 |
| 改设计规范 | `reference/motion-design-spec.md` | 与 `facts/` 冲突时以 `facts/` 为准 |

### 4.1 四道门验证（加新能力时想清楚要过哪几道）

| 门 | 证明什么 | 入口 |
| --- | --- | --- |
| `motion.py apply --assert-geometry` | 几何没动 | 布局相关的改动必须过 |
| `motion.py check` | 没有零效果的形状 | 动效相关必过 |
| `verify_motion.py` | XML 结构合法 | 结构相关必过 |
| `motion.ps1 -Strict` | **PowerPoint 真开保存后效果还在** | 任何涉及真渲染的改动都必过 |

> 第四道门是本项目最贵也最重要的一道 —— lxml 全绿不代表 PowerPoint 认。

---

## 5. 已知热点

**这些是体检会持续标记的地方，不是待办清单** —— 记录在此是为了让人知道「为什么它特殊」。

| 文件 / 位置 | 症状 | 为什么特殊 | 现状 |
| --- | --- | --- | --- |
| `scripts/motion.py` | 1914 行，**逼近 2000 硬上限** | 全库被依赖最多的文件（10 个模块） | 已定：**本轮只立规范，不拆**。拆它要动 10 个依赖者，不是单次改动能验证的 |
| `scripts/transition_probe/commands.py` | 1367 行，超 800 软上限 | 19 个 CLI 子命令都在这里 | 拆分样板（见 `history/`）已验证可行，待拆时照抄 |
| `scripts/motion.py ⇄ check_coverage.py` | 表面 import 环 | 靠**函数内延迟导入**刻意打断 | **刻意设计，不是 bug**。别"顺手修掉" |
| `tools/` | 15 个脚本，含一次性拆分工具 | 一次性工具拆完即删 | 删不删由我逐个问你 —— 见 [`CONTRIBUTING.md`](CONTRIBUTING.md) §九 |

---

## 6. 维护

| 谁 | 管什么 |
| --- | --- |
| `tools/scan_deps.py` | 生成依赖图、环检测、波及面（**只读**） |
| `tests/test_code_index.py` | 守本索引与真实代码同步 |
| [`CONTRIBUTING.md`](CONTRIBUTING.md) | 体积阈值、拆分约束、依赖方向的**规范来源** |

**规矩**：改代码结构（加文件、拆文件、改依赖）必须同步更新本索引，
否则 `test_code_index.py` 会失败 —— 那是故意的：**索引腐烂要有人被看见**。
