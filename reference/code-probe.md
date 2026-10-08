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
| `scripts/build_transition_table.py` | Build the slide-transition reference table（入口） | 切换实测入口，转发到 `transition_probe/` | `transition_probe` | —（没人依赖） | 139 行 |
| `scripts/transition_probe/__init__.py` | transition_probe — 切换实测探针工具包（从单文件拆出） | 包标记 | —（叶子） | `build_transition_table` | 20 行 |
| `scripts/transition_probe/common.py` | transition_probe.common — 共享常量与通用工具 | COM 会话与共用小工具 | —（叶子） | —（没人依赖） | 108 行 |
| `scripts/transition_probe/data.py` | transition_probe.data — 探测数据表 | 切换形态定义 | —（叶子） | —（没人依赖） | 397 行 |
| `scripts/transition_probe/decks.py` | transition_probe.decks — deck 构造器 | 造 deck | —（叶子） | —（没人依赖） | 521 行 |
| `scripts/transition_probe/analysis.py` | transition_probe.analysis — 帧分析与表构造 | 量帧与数据分析 | —（叶子） | —（没人依赖） | 629 行 |
| `scripts/transition_probe/commands.py` | transition_probe.commands — CLI 子命令 | 19 个 CLI 子命令的实现 | —（叶子） | —（没人依赖） | 1367 行 ⚠️ 超 800 |
| `scripts/build_camera_table.py` | Build the 3D camera reference table: parameter -> MEASURED rendered result | 相机路径实测表 | —（叶子） | —（叶子） | 664 行 |
| `scripts/build_shape_evidence.py` | 把 3a 的【文字判定】与【渲染实物】放进同一份 deck，供人眼复核 | 形状证据实测 | —（叶子） | —（叶子） | 361 行 |
| `scripts/build_dual_photo.py` | 构造"同图双版本"版式的样例页，供 verify_dual_photo.py 做正/负对照 | 双色照片对照实验 | —（叶子） | `test_dual_photo` | 184 行 |
| `scripts/analyze_video.py` | dsh-ppt-office-motion :: motion analysis from a video recording | 视频逐帧分析 | —（叶子） | —（没人依赖） | 424 行 |
| `scripts/build_pitfall_map.py` | 生成 reference/pitfall-map.md —— §编号 → 文件的唯一映射 | 生成症状→根因映射表 | —（叶子） | —（没人依赖） | 138 行 |

共 12 个文件，其中 Python 4952 行。

> **包内为什么显示"叶子"**：`transition_probe/` 各模块之间用**相对导入**
> （`from .common import ...`），扫描按包边界处理，所以对外呈现为叶子节点。
> 这不是依赖缺失 —— 包内**严格分层**：`common → data → decks → analysis → commands`，
> 下层不反向 import 上层。这条分层是有意的，别为了"让依赖图好看"改掉。

---

## `transition_probe/` 的内部分层

拆开之后单文件都不大（最大 629 行），但**依赖必须单向**：

```text
  commands.py        19 个 CLI 子命令，只做参数解析与调度
      ▲
  analysis.py        量帧、算差异、构造结果表
      ▲
  decks.py           造 deck（每候选一个）
      ▲
  data.py            形态定义（哪些要测）
      ▲
  common.py          COM 会话、常量、通用小工具
```

**唯一入口**：`build_transition_table.py`转发到包。
其他模块**不许**直接 import `commands`—— 想跑实测请走入口，那是接口。

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

`transition_probe/commands.py` **1367 行，超 800 软上限** —— 19 个 CLI 子命令都在这里。

**拆分样板已经验证过**：`tools/split_transition_probe.py` 就是把3048 行的单体
拆成现在这个包的脚本，拆完全绿。真要拆 `commands.py` 时照抄它的做法，
四条硬约束见 [`CONTRIBUTING.md`](../CONTRIBUTING.md)。

拆的时候按**子命令的领域**分（enum / deck / roundtrip / camera / shape），
不要按行数硬切—— 按行数切出来的文件职责是乱的。

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种切换形态 | `facts/transitions.json` + `transition_probe/data.py` |
| 加一种相机路径 | `facts/cameras.json` + `transition_probe/data.py` |
| 加一种形状证据实验 | `build_shape_evidence.py` + `probe_shapes.ps1` |
| 改页面切换机制（为什么挂终点页） | `reference/transition-model.md` —— **唯一真相源** |
| 加探针子命令 | `transition_probe/commands.py`，注意分层 |