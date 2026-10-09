# 代码索引 · 探针层

> 上一层：[`CODE_INDEX.md`](../CODE_INDEX.md) · 同层：[注入层](code-injection.md) · [校验层](code-verify.md) · [工具层](code-tooling.md)
>
> **表里的行数与依赖关系由 `tools/gen_code_tables.py` 扫出来，不是手抄的。**
> 重新生成：`python tools/gen_code_tables.py probe`

## 这一层负责什么

**造样本、跑实测、量结果。** 这一层产出的是**证据**，不是动效——
所以它可以放心地写"我猜错了"，那正是它的价值。

判据只有一条：

> **PowerPoint 自己写回来的东西，才是权威。**
> 我们的表格、我们的推断、我们的 lxml 校验，都只能算候选。

---

## 速查表

| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |
| --- | --- | --- | --- | --- | --- |
| `scripts/build_transition_table.py` | Build the slide-transition reference table（入口） | 切换实测入口，转发到 transition_probe/ | `transition_probe` | —（没人依赖） | 143 行 |
| `scripts/transition_probe/__init__.py` | transition_probe — 切换实测探针工具包（从单文件拆出） | 包标记 | `analysis`、`commands_attr`、`commands_mechanism`、`commands_shape`、`commands_table`、`commands_timing`、`common`、`data`、`decks` | `build_transition_table` | 24 行 |
| `scripts/transition_probe/common.py` | transition_probe.common — 共享常量与通用工具 | COM 会话与共用小工具 | —（叶子） | `transition_probe`、`analysis`、`commands_attr`、`commands_mechanism`、`commands_shape`、`commands_table`、`commands_timing`、`decks` | 108 行 |
| `scripts/transition_probe/data.py` | transition_probe.data — 探测数据表 | 切换形态定义 | —（叶子） | `transition_probe`、`analysis`、`commands_attr`、`commands_shape`、`commands_table`、`commands_timing`、`decks` | 397 行 |
| `scripts/transition_probe/decks.py` | transition_probe.decks — deck 构造器 | 造deck | `common`、`data` | `transition_probe`、`commands_attr`、`commands_mechanism`、`commands_shape`、`commands_table`、`commands_timing` | 521 行 |
| `scripts/transition_probe/analysis.py` | transition_probe.analysis — 帧分析与表构造 | 量帧与数据分析 | `common`、`data` | `transition_probe`、`commands_attr`、`commands_shape`、`commands_table`、`commands_timing` | 629 行 |
| `scripts/transition_probe/commands_table.py` | transition_probe.commands_table — 建表主链路 | 建表主链路：造 deck → 枚举扫描 → 合并成表 | `analysis`、`common`、`data`、`decks` | `transition_probe` | 395 行 |
| `scripts/transition_probe/commands_mechanism.py` | transition_probe.commands_mechanism — 机制层 | 机制层：切换挂在哪一页 | `common`、`decks` | `transition_probe` | 206 行 |
| `scripts/transition_probe/commands_shape.py` | transition_probe.commands_shape — 形态层与方向 | 形态层与方向：效果看起来在做什么 | `analysis`、`common`、`data`、`decks` | `transition_probe` | 351 行 |
| `scripts/transition_probe/commands_attr.py` | transition_probe.commands_attr — 属性取值全集 | 属性取值全集 | `analysis`、`common`、`data`、`decks` | `transition_probe` | 160 行 |
| `scripts/transition_probe/commands_timing.py` | transition_probe.commands_timing — 切换 × 页内动画 | 切换 × 页内动画：结构平行、时间串行 | `analysis`、`common`、`data`、`decks` | `transition_probe` | 315 行 |
| `scripts/build_camera_table.py` | Build the 3D camera reference table: parameter -> MEASURED rendered result | 相机路径实测表 | —（叶子） | —（没人依赖） | 664 行 |
| `scripts/build_shape_evidence.py` | 把 3a 的【文字判定】与【渲染实物】放进同一份 deck，供人眼复核 | 形状证据实测 | —（叶子） | —（没人依赖） | 361 行 |
| `scripts/build_dual_photo.py` | 构造"同图双版本"版式的样例页，供 verify_dual_photo.py 做正/负对照 | 双色照片对照实验 | —（叶子） | `test_dual_photo` | 184 行 |
| `scripts/analyze_video.py` | dsh-ppt-office-motion :: motion analysis from a video recording | 视频逐帧分析 | —（叶子） | —（没人依赖） | 424 行 |
| `scripts/build_pitfall_map.py` | 生成 reference/pitfall-map.md —— §编号 → 文件的唯一映射 | 生成症状→根因映射表 | —（叶子） | —（没人依赖） | 138 行 |

共 16 个文件，其中 Python 5020 行。

> **包内依赖现在是可见的（2026-10-09 修复）。**
> `scan_deps` 曾把相对导入 `from .common import X` 拼成
> `...transition_probecommon`（**缺点号**），于是在依赖图上整个包的内部依赖
> **静默消失** —— 看起来像"包内没有依赖"。
>
> 更糟的是：当时给这个现象编了个听起来合理的解释（"扫描按包边界处理"）
> 写在这里 —— **一个 bug 就这么被解释成了设计**。
> 这与 `CODE_RULES.md` §十.4 记的是同一类错：**度量工具错了，结论就反了。**
>
> 修好后实测（下面那张图现在是**跑出来的**，不是读代码看出来的）：
> 包内**严格分层、0 环**。

---

## `transition_probe/` 的内部分层

拆开之后单文件都不大（最大 629 行），但**依赖必须单向**：

```text
  commands_table.py      建表主链路：造 deck → 枚举扫描 → 合并成表
  commands_mechanism.py  机制层：切换挂在哪一页
  commands_shape.py      形态层与方向：效果看起来在做什么
  commands_attr.py       属性取值全集
  commands_timing.py     切换 × 页内动画：结构平行、时间串行
      ▲  19 个 CLI 子命令（同一层的五个兄弟，互不调用）
      │
  analysis.py            量帧、算差异、构造结果表      （被 5 个依赖）
      ▲
  decks.py               造 deck（每候选一个）          （被 6 个依赖）
      ▲
  data.py                形态定义（哪些要测）            （被 7 个依赖）
      ▲
  common.py              COM 会话、常量、通用小工具      （被 8 个依赖）
```

**这张图是 `python tools/scan_deps.py` 跑出来的**（0 环），不是读代码看出来的。
「被 N 个依赖」就是**改动波及面**：改 `common.py` 会影响 8 个模块 ——
它虽然只有 108 行，但它是全包的地基。

> 依赖者里包含 `__init__.py`（它 import 全部五个层是为了再导出，那是它的职责）。
> 所以这里的 N 比"业务模块"数多 1 —— **不是虚高**：改 `common.py` 时
> `__init__.py` 的再导出行确实也要看一眼。

> **包内也有热点，这是正常的。** 包**对外**只有 1 个依赖者
> （`build_transition_table.py`），把复杂度关在了包里 —— 这正是拆包的目的。
> 但**包内**的 `common` / `data` / `decks` / `analysis` 被依赖 ≥5，
> 体检会报 `hot-module` ADVISE。那不是问题，是提醒"改之前先看清谁在用"。

**唯一入口**：`build_transition_table.py`转发到包。
其他模块**不许**直接 import `commands_*` —— 想跑实测请走入口，那是接口。

> **五个 `commands_*` 是同一层的兄弟，不是一个调用链。**
> 实测：19 个函数**没有一个调用另一个**，它们都只依赖下面四层。
> 所以拆它们不影响分层 —— 是"五个并列的门面"，不是"五级台阶"。

### 为什么"一候选一个 deck"

一个 schema 非法的 `<p:transition>` 子元素会让PowerPoint **拒开整个文件**。
把所有候选塞进一个 deck，一次失败就什么都测不出来。
→ 批量实测必须一候选一deck，代价是慢，换来的是每次失败只污染一个样本。

---

## PowerShell 侧的探针（在工具层，跑的是这一层的活）

7 个 `probe_*.ps1` + `motion.ps1` 见 [`code-tooling.md`](code-tooling.md#powershell-侧)。
它们和Python 探针的分工：

| 谁来问 | 用哪个 | 为什么 |
| --- | --- | --- |
| "PowerPoint 认不认我写的切换？" | `probe_roundtrip.ps1` | 单 deck 往返，读回 EntryEffect / Duration |
| "我猜的元素名对不对？" | `probe_enum_scan.ps1` | **只有它能反查**。往返只能回答"合法/拒开"，拒开不告诉你正确的名字 |
| "切换挂在出发页还是终点页？" | `probe_anchor.ps1` | 问的是**位置**，不是合法性 |
| "`transition` 和 `timing` 会不会互相吃掉？" | `probe_timing_roundtrip.ps1` | 整目录批量，专治**静默丢失** |
| "观众实际看到的是什么？" | `probe_shapes.ps1` | 48 个 deck 各出一段 mp4，给人眼看 |

> **`probe_shapes.ps1` 为什么不并进 `probe_anchor.ps1`**：
> 两个活混在一起会让 anchor 脚本把 ~48 段slide XML 无理由地拖进报告。
> 这是一次**有意的分离** —— 见该文件头部注释。

---

## 改动这一层要过什么

| 改动类型 | 必过的门 |
| --- | --- |
| 改了 `facts/` 里的形态或参数 | 重跑对应实测，**把新结果贴进本索引旁的表**，不许只在代码里改 |
| 改了 deck 构造 | 跑一次往返，确认 PowerPoint 还认|
| 改了 `analysis.py` 的量帧逻辑 | 拿已知答案样本（`make_calibration.ps1` 的产物）验证 |

**这一层的产出是数字和表格，数字过期比代码腐烂更危险** ——
代码坏了会报错，一份过期的实测表只会安静地把人带错方向。

---

## 已知热点

**`commands.py` 已拆完（2026-10-09）。** 1367 行 / 19 个子命令 →
5 个 `commands_*.py`（最大 395 行），按**探测维度**分。

**两次拆分的样板都在 `tools/`，留着照抄**：

| 样板 | 干了什么 | 适用场景 |
| --- | --- | --- |
| `tools/split_transition_probe.py` | 3048 行单体 → 包（`common/data/decks/analysis/commands`） | 拆**整个模块**成多层包 |
| `tools/split_commands.py` | 1367 行模块 → 5 个平级文件 | 拆**一个层**成同层兄弟 |

**为什么按"探测维度"分而不是按行数**：这五个维度是它自己 docstring
里那张内容表**已经分好的组** —— 建表主链路 / 机制层 / 形态层 / 属性 / 时序。
按行数切出来的文件职责是乱的；按维度切出来的文件，名字就说明了它装什么。

**无损证据**：拆前 19 个函数 → 拆后 19 个，**逐字节对比无一字差异**。

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种切换形态 | `facts/transitions.json` + `transition_probe/data.py` |
| 加一种相机路径 | 改 `scripts/build_camera_table.py` 重跑 → 结果写进 `reference/camera-reference.md`（它由 `facts/axes.json` 声明为唯一真相源） |
| 加一种形状证据实验 | `build_shape_evidence.py` + `probe_shapes.ps1` |
| 改页面切换机制（为什么挂终点页） | `reference/transition-model.md` —— **唯一真相源** |
| 加探针子命令 | 按维度选一个 `transition_probe/commands_*.py`（table/mechanism/shape/attr/timing）|