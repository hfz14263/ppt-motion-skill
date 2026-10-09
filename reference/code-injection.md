# 代码索引 ·注入层

> 上一层：[`CODE_INDEX.md`](../CODE_INDEX.md) · 同层：[探针层](code-probe.md) · [校验层](code-verify.md) · [工具层](code-tooling.md)
>
> **表里的行数与依赖关系由 `tools/gen_code_tables.py` 扫出来，不是手抄的。**
> 重新生成：`python tools/gen_code_tables.py injection`

## 这一层负责什么

把 spec 变成 pptx 里真的动起来的 OOXML —— 注入引擎本体，和它的浏览器端重放。

**这一层是全库耦合的顶点**，也是唯一一层「改一行可能影响别人」的地方：

| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |
| --- | --- | --- | --- | --- | --- |
| `scripts/motion.py` | **门面**：apply 主流程 + CLI + 重新导出下面四层 | `apply` / `inspect` / `preview` / `catalog` / `player` / `check` | `motion_media`、`motion_spec`、`motion_timing`、`motion_xml` | 10 个模块（见下） | 545 行 |
| `scripts/motion_xml.py` | motion.xml — 常量与 XML 底层（**整个引擎的地基**） | 常量 + zip + XML 字符串手术 + 单例 + 形状索引 | —（叶子） | `motion`、`motion_media`、`motion_spec`、`motion_timing` | 528 行 |
| `scripts/motion_timing.py` | motion.timing — timing 与 transition 的 XML 生成 | timing / transition 生成 | `motion_xml` | `motion` | 440 行 |
| `scripts/motion_media.py` | motion.media — 3D 相机与图片填充 | 3D 相机 + 图片填充 | `motion_xml` | `motion` | 338 行 |
| `scripts/motion_spec.py` | motion.spec — spec 处理与**结构校验** | spec 处理 + 三个校验函数 | `motion_xml` | `motion` | 228 行 |
| `scripts/player.py` | 浏览器端动效重放（HTML） | 生成 HTML 预览 | `motion` | `check_coverage`、`design_audit`、`selftest` | 578 行 |

共 6 个文件，其中 Python 2657 行。

**10 个依赖者全部只 import `motion`**（不直接 import 那四层）——
`test_camera`、`test_fill_window`、`test_morph`、`test_singletons`
· `check_coverage`、`design_audit`、`player`、`selftest`
· `verify_motion`、`verify_singletons`。

> **这是"门面"的价值**：四层是 2026-10-09 才拆出来的，
> 但 **10 个依赖者一行都没改** —— 它们仍然只 `import motion`。

---

## 四个层各装什么

```text
  motion.py          门面：apply_motion / strip_animations / inspect / main
                     ＋ 重新导出下面四层（外部接口零变化的关键）
      ▲
      ├── motion_timing.py   timing / transition 的 XML 生成
      ├── motion_media.py    3D 相机 + 图片填充
      └── motion_spec.py     spec 处理 + 结构校验
              ▲  三层互不调用 —— 实测跨组调用 0 处
              │
          motion_xml.py      常量 + zip + XML 手术 + 单例 + 形状索引
                             ★ 地基：被上面四层全部依赖，自己无内部依赖
```

**为什么这么分**：按**依赖方向**分，不按行数。实测的跨组调用：

| 从 | 到 | 处数 |
| --- | --- | --- |
| `apply`（motion.py） | xml / timing / media / spec | 18 / 5 / 7 / 4 |
| `timing` / `media` / `spec` | xml | 4 / 4 / 6 |
| **反向调用** | —— | **0（严格 DAG）** |

`motion_xml.py` 里的 `element_spans` **被 10 处调用** ——
它是全引擎唯一的重心：所有 XML 改写都走它（只动属性值与文本，不重排 XML）。

---

## 核心逻辑与数据流

### 一次 `apply` 的完整路径

```text
spec (JSON/YAML)
   │
   ▼
normalize_spec()              字段校验 + 补默认值；字段含义同时写在 SKILL.md
   │                          （motion_spec.py）
   ├──▶ catalog()            查 137 个效果的 presetID 与 XML 模板（只读）
   │
   ▼
read_parts()                  开 zip；**不编辑的 part 逐字节复制**
   │                          （motion_xml.py —— 只有它碰 zipfile）
   ├──▶ 动画：改<p:timing>       形状级动画树（motion_timing.py）
   ├──▶ 切换：写<p:transition>   ★ 挂在【终点页】上（motion_timing.py）
   ├──▶ 3D：写<a:sp3d>          相机 presetID + 路径（motion_media.py）
   └──▶ 图片填充：<p:blipFill>  淡化整图铺底 + 原色塞进形状当前景（motion_media.py）
   │
   ▼
write_parts() → 落盘
```

**三条最容易踩的**（都指向 `pitfall-*`，不在这里展开）：

| 症状 | 根因 |
| --- | --- |
| 切换整体错位一个页边界 | `<p:transition>` 挂在终点页 |
| `<p:timing>` 静默消失 | transition 与 timing 互相吃掉，见 `probe_timing_roundtrip.ps1` |
| PowerPoint 拒开 `E_FAIL` | 多余 preset 包装层 / schema 顺序，见 `pitfall-file-corruption.md` |

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

**`motion.py` 已拆完（2026-10-09）**：1914 行 / 81 符号 →
门面 545 行 + 四层（528 / 440 / 338 / 228）。

**这次为什么推翻了 2026-10-08 的「不拆」**：
当时的理由是「要动 10 个依赖者，不是单次改动能验证的」—— **理由是对的，
但前提变了**：先量清了内部调用图，发现是**严格 DAG**（零反向、零环），
而且 10 个依赖者用到的都是公开名、只走 `import motion` 一个口。
**拆它从"不可验证"变成了"一次可验证的搬运"**：
拆完 81 个符号逐字节相同、83 条顶层语句零丢失、
5 份测试（359+20+29+18+16 项）+ selftest 63 项全绿。

> **教训（已写进 `CONTRIBUTING.md` §六.2）**：
> 「不能拆」往往有**前提**。前提变了就该重看，而不是把它当成永久结论。

**当时那条「逐符号验证」不够 —— 漏了两行**（`TRANSITIONS["none"] = ""`
这类**下标赋值**不是 `Name` 赋值，第一版搬运逻辑只认 `Name`）。
现在 `tools/split_motion.py` 有 `assert_fully_consumed()`：
**每条顶层语句没被搬走就中止**。这条断言放在拆分时，不放到事后验证 ——
因为事后验证的脚本和搬运逻辑**共享同一个盲区**，两边都会说"没问题"。

新增逻辑时按层放：

| 我要改什么 | 放哪 |
| --- | --- |
| 动画 / 切换的 XML 生成 | `motion_timing.py` |
| 3D 相机 / 图片填充 | `motion_media.py` |
| spec 字段 / 校验判据 | `motion_spec.py` |
| XML 手术 / 单例 / 形状索引 | `motion_xml.py`（**改这里要过 4 个依赖者的眼**） |
| apply 主流程 / CLI | `motion.py` |

**三个拆分样板**（都在 `tools/`，都验证过）：

| 样板 | 干什么 |
| --- | --- |
| `split_transition_probe.py` | 3048 行单体 → 多层包（拆**模块**） |
| `split_commands.py` | 1367 行模块 → 5 个同层文件（拆**一层**） |
| `split_motion.py` | 1914 行模块 → 门面 + 四层（拆**分层**；带完整性断言） |

> **别"顺手"把 `motion ⇄ check_coverage` 的表面环修掉。**
> 那个环是靠 `motion.py check` 里的**函数内延迟导入**刻意打断的，是设计信号。
> 依赖扫描只统计模块级 import，就是为了不把它当成违规。
> 拆分时那条 `import check_coverage` 跟着 `main()` 留在门面里 ——
> **它有专门的验证**（`tools/_verify_motion_split.py` 会检查它没被提到模块级）。

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种动画效果 | 改 `scripts/motion_catalog.json`（**权威表**，COM 提取的真实 presetID），**不要**动 `motion_timing.py` 的模板 |
| 加一种切换形态 | 改 `facts/transitions.json` + `transition_probe/data.py` + `motion_timing.py` |
| 改 spec 字段 | `motion_spec.py` 的 `normalize_spec`，字段含义同步写 `SKILL.md` |
| 加一个 CLI 子命令 | `motion.py` 的 CLI 分发 —— **子命令名是接口，改名要改全仓库** |
| 改 XML 改写逻辑 | `motion_xml.py` —— **先看清 4 个依赖者**（含 3 个兄弟层） |
| 3D / 填充 | `motion_media.py`（角度对外是度，内部换算 1/60000） |
