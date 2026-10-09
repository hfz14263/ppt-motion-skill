# 代码索引 · 工具层

> 上一层：[`CODE_INDEX.md`](../CODE_INDEX.md) · 同层：[注入层](code-injection.md) · [探针层](code-probe.md) · [校验层](code-verify.md)
>
> **表里的行数与依赖关系由 `tools/gen_code_tables.py` 扫出来，不是手抄的。**
> 重新生成：`python tools/gen_code_tables.py tooling`
>
> ⚠️ **往 `tools/` 加脚本前，先查这张表**（[`CODE_RULES.md`](../CODE_RULES.md) §六.8）——
> 做过类似的事就改现有的；确要新建，**必须登记进 LAYERS**，否则体检会拦。
> 登记表就是"有没有现成的"的对照清单。

## 这一层负责什么

**不产动效，但项目运转要靠它们。**

三类成员：

1. **PowerShell / COM 层** —— Python 碰不到真渲染，只有这一层能
2. **一次性拆分/归档工具** —— `tools/split_*.py` / `archive_*`，**拆完留档**（下次照抄）
3. **杂项工具** —— 推送、主题入库、目录与索引生成

---

## 速查表

| 文件 | 职责 | 对外接口 / 入口 | 依赖谁 | 被谁依赖 | 体积 |
| --- | --- | --- | --- | --- | --- |
| `scripts/push_via_api.py` | 当 github.com 被代理挡住、但 api.github.com 通时，用 Git Data API 推送 | 走 API 推送 | —（叶子） | —（没人依赖） | 203 行 |
| `scripts/vendor_themes.py` | Vendor a curated set of design systems from open-kimi-ppt-skill (MIT) | 主题文件入库 | —（叶子） | —（没人依赖） | 89 行 |
| `scripts/motion.ps1` | dsh-ppt-office-motion :: Office COM layer | Office COM 层：media 插入 + 真渲染 | —（叶子） | —（没人依赖） | 482 行 |
| `scripts/make_calibration.ps1` | dsh-ppt-office-motion :: build the known-answer calibration deck + video | 造标定deck + 标定视频 | —（叶子） | —（没人依赖） | 103 行 |
| `scripts/probe_anchor.ps1` | dsh-ppt-office-motion :: measure WHERE a slide transition actually takes effect | 测 `<p:transition>` 锚点在终点页还是出发页 | —（叶子） | —（没人依赖） | 185 行 |
| `scripts/probe_createvideo.ps1` | dsh-ppt-office-motion :: probe whether Presentation.CreateVideo works HERE | 探测本机 CreateVideo 是否可用 | —（叶子） | —（没人依赖） | 146 行 |
| `scripts/probe_enum_scan.ps1` | dsh-ppt-office-motion :: enum scan -- let PowerPoint name the transitions | 让 PowerPoint 自己给切换命名（反查猜错的元素名） | —（叶子） | —（没人依赖） | 70 行 |
| `scripts/probe_roundtrip.ps1` | dsh-ppt-office-motion :: roundtrip an applied deck and read the transition back | 单deck 往返：PowerPoint 认不认我写的切换 | —（叶子） | —（没人依赖） | 76 行 |
| `scripts/probe_shapes.ps1` | dsh-ppt-office-motion :: render every shape-probe deck to a video | 把每个形状探针 deck 渲染成视频 | —（叶子） | —（没人依赖） | 109 行 |
| `scripts/probe_timing_roundtrip.ps1` | dsh-ppt-office-motion :: roundtrip every deck in a dir and report what survived | 整目录往返：transition 与 timing 会不会互相吃掉 | —（叶子） | —（没人依赖） | 134 行 |
| `scripts/probe_transitions.ps1` | dsh-ppt-office-motion :: transition probe -- PowerPoint roundtrip + enum readback | 切换探针：往返 + enum 回读 | —（叶子） | —（没人依赖） | 96 行 |
| `tools/split_transition_probe.py` | 一次性拆分器：把 build_transition_table.py（3048 行）拆成 transition_probe/ 包 | 一次性：拆 transition_probe 成包 | —（叶子） | —（没人依赖） | 296 行 |
| `tools/split_commands.py` | 一次性拆分器：把 transition_probe/commands.py（1367 行 / 19 个 cmd）按功能域拆开 | 一次性：把 commands.py 按探测维度拆成 5 份 | —（叶子） | —（没人依赖） | 216 行 |
| `tools/split_motion.py` | 一次性拆分器：把 scripts/motion.py（1914 行 / 81 个顶层符号）按分层拆开 | 一次性：把 motion.py 按分层拆成门面 + 四层（带完整性断言） | —（叶子） | —（没人依赖） | 467 行 |
| `tools/split_contributing.py` | 一次性拆分器：把 CONTRIBUTING.md 按**读者是谁**拆成两份 | 一次性：把 CONTRIBUTING.md 按读者拆成两份 | —（叶子） | —（没人依赖） | 155 行 |
| `tools/split_pitfalls.py` | 一次性的拆分器：把 com-pitfalls.md 按主题切成多份，编号保持不变 | 一次性：拆 pitfall 族 | —（叶子） | —（没人依赖） | 160 行 |
| `tools/split_symptoms.py` | 一次性拆分器：把 symptoms.md 按「现象栏」切成多份 | 一次性：拆 symptom 族 | —（叶子） | —（没人依赖） | 119 行 |
| `tools/split_morph_recipes.py` | 一次性拆分器：把 morph-and-3d-recipes.md 按主题切成 3 份 + 目录页 | 一次性：拆 morph 配方 | —（叶子） | —（没人依赖） | 105 行 |
| `tools/split_tests.py` | 一次性拆分器：把 test_transition_table.py（1633 行 / 22 组）拆成三份 | 一次性：拆 test_transition_table.py 成三份 | —（叶子） | —（没人依赖） | 243 行 |
| `tools/archive_handover_85.py` | 一次性归档器：把 HANDOVER.md §8 里**已完成**的三节搬进 history/ | 一次性：把 HANDOVER §8 已完成的三节搬进 history/ | —（叶子） | —（没人依赖） | 138 行 |
| `tools/gen_code_tables.py` | 生成 reference/code-*.md 里的速查表 —— 表格数据不手抄 | 生成 / 刷新四份 code-*.md 的速查表（表的数据唯一来源） | —（叶子） | `test_code_index` | 228 行 |
| `tools/add_toc.py` | 给长文档补一张目录（TOC）—— 只加导航，不动正文一个字 | 给长文档加目录 | —（叶子） | —（没人依赖） | 172 行 |
| `tools/collapse_symptom.py` | 分级收敛 symptom 侧正文：把重复正文换成指向根因的链接 | 根因收敛 1/3：重复节收敛为索引（幂等） | —（叶子） | —（没人依赖） | 199 行 |
| `tools/add_symptom_view.py` | 给已收敛的 symptom 节补「现象」导语 | 根因收敛 2/3：给已收敛节补现象导语（幂等） | —（叶子） | —（没人依赖） | 109 行 |
| `tools/link_symptom_to_pitfall.py` | 给档3（保留现象正文的节）补根因链接 | 根因收敛 3/3：给保留正文的节补根因链接（幂等） | —（叶子） | —（没人依赖） | 123 行 |
| `tools/add_symptom_leadin.py` | 给 pitfall 顶层编号节插入「现象」导语 | 按 TSV 给 pitfall 顶层节插现象导语 | —（叶子） | —（没人依赖） | 145 行 |
| `reference/design-system/build_index.py` | Generate the design-system index for this skill | 生成 / 校验 design-system 索引（--check 只读） | —（叶子） | —（没人依赖） | 299 行 |

共 27 个文件，其中 Python 3466 行。

> **`.ps1` 显示"叶子"是正常的** —— PowerShell 脚本之间不 import，
> 它们被**人**从命令行调用，不是被代码依赖。上表的依赖分析只对 `.py` 有效。

---

## PowerShell侧

### `motion.ps1` —— 唯一能真渲染的东西

**OOXML 引擎是 `motion.py`，这里只做 Python 碰不到的两件事**：
media（视频/音频）插入，和让 PowerPoint 真开真存。

| 用法 | 干什么 |
| --- |
| `motion.ps1` | 常规：插入 media、真渲染 |
| `motion.ps1 -Strict` | **第四道门**：开→存→复查效果还在不在 |

两条硬规则（写在文件头，来自实测，不是风格偏好）：

| 规则 | 后果 |
| --- | --- |
| `WithWindow` 必须为真（这个构建上） | 否则 COM 静默失败，见 `pitfall-com.md` |
| 输出路径保持在 ~200 字符内 | 超了会失败 |

### 7 个 `probe_*.ps1` 各自问什么

它们跑的是[探针层的活](code-probe.md#powershell-侧的探针在工具层跑的是这一层的活)，
问题是**分开的**，别合并：

| 脚本 | 问的问题 | 独有之处|
| --- | --- | --- |
| `probe_enum_scan.ps1` | 元素名我猜错了，正确的是什么？ | **唯一能反查名字**的。往返只说"拒开"，不说正确名字 |
| `probe_roundtrip.ps1` | PowerPoint 认不认我写的切换？ | 单 deck，读回 EntryEffect / Duration |
| `probe_transitions.ps1` | 同上 + enum 回读 | 批量版 |
| `probe_anchor.ps1` | 切换挂在出发页还是终点页？ | 问**位置**，不是合法性 |
| `probe_timing_roundtrip.ps1` | transition 和 timing 会不会互相吃掉？ | 整目录，专治**静默丢失** |
| `probe_shapes.ps1` | 观众实际看到什么？ | 48 个 mp4 给人眼看 |
| `probe_createvideo.ps1` | 本机 `CreateVideo` 能用吗？ | 环境探测，一次性结论 |

**它们分开是有理由的**，理由都写在各自文件头。合并的代价已经付过一次：
`probe_anchor.ps1` 曾经顺带渲染 48 个 deck，把 48 段 slide XML 无理由拖进报告。

---

## 一次性拆分/归档工具（`tools/split_*.py`、`archive_*`）

**它们的作用已经完成，但脚本要留着** —— 下次拆文件时照抄它的做法比自己想靠谱。

**按"拆什么"分两类，各有一个样板**：

| 脚本 | 干了什么 | 拆前体积 | 适用场景 |
| --- | --- | --- | --- |
| `split_transition_probe.py` | `build_transition_table.py` → `transition_probe/` 五层包 | 3048 行 | 拆**整个模块**成多层包 |
| `split_commands.py` | `transition_probe/commands.py` → 5 个平级 `commands_*.py` | 1367 行 | 拆**一个层**成同层兄弟 |
| `split_motion.py` | `motion.py` → 门面 + 四层（带完整性断言） | 1914 行 | 拆**分层**（10 个依赖者一行没改） |
| `split_tests.py` | `test_transition_table.py` → 三份测试 + `helpers.py` | 1633 行 | 拆**测试文件**（带 `helpers.py` 提取） |
| `split_pitfalls.py` | `com-pitfalls.md` → 10 份 `pitfall-*.md` | 单文件巨长 | 拆**文档**（编号保住） |
| `split_symptoms.py` | `symptoms.md` → 5 份 `symptom-*.md` | 同上 | 同上 |
| `split_morph_recipes.py` | `morph-and-3d-recipes.md` → 3 份 + 目录页 | 同上 | 同上 |
| `split_contributing.py` | `CONTRIBUTING.md` → 按读者分两份 | 19924 字符 | 拆**文档**（按读者，编号不重排） |
| `archive_handover_85.py` | `HANDOVER.md` §8 的三节 → `history/` | 21901 字符 | **归档**（不是拆分：内容整块搬走） |

**四条硬约束（每一条都是踩出来的）**：

1. **外部接口零变化** —— 函数名 / CLI 子命令名 / 退出码全不变
2. **逐符号搬运，逐字节验证** —— `split_commands.py` 的验收是
   "19 个函数拆前拆后**逐字节相同**"，不是"看起来对"
3. **拆分脚本的产物会覆盖自己的输入** —— 所以**必须从 git 取源**
   （`split_tests.py` 第一版踩过：第一次成功、第二次切出乱码）
4. **搬完必须跑测试** —— 代码依赖自身位置时（`__file__` 拼路径）
   搬运会改变行为，静态检查看不见

**为什么不删**：
它们是**怎么做拆分**的唯一完整记录，删了下次就是重新踩一遍。
体积代价（约 1900 行 Python）是值得的。

**2026-10-09 维护记录**：删了 2 个已消耗的 `rewrite_*.py`（一次性改写，模式与
§二十一「打补丁，不重写」相悖，没有复用场景）与 1 个被取代的
`archive_handover_backlog.py`（同类的 `archive_handover_85` 是更好的样板：
从 git 取源 + 幂等）。判据是「职责已完成 + 无复用场景 + 只在索引里被提过」。
拆分样板一个没动 —— 它们各有不同的适用场景（见上表）。

> **但这不等于 `tools/` 里的东西都不许删。**
> 判断标准见 [`CODE_RULES.md`](../CODE_RULES.md) §六.6.1：
> **我（AI）认为当前任务已完成、为此生成的临时文件可以删时，
> 会先逐个问你，得到答复才删。**

---

## `tools/` 里还有什么

除了一次性拆分器，还有幂等的维护脚本 ——
它们**每次跑都必须安全重跑**，这是硬要求：

| 脚本 | 作用 | 幂等保障 |
| --- | --- | --- |
| `scan_deps.py` | 依赖图 / 环 / 波及面 | 只读 |
| `gen_code_tables.py` | 生成本索引的表格 | 只输出，不写文件 |
| `measure_overlap.py` | 两族文档逐节相似度 | 只读 |
| `check_symptom_coverage.py` | 现象导语覆盖度 | 只读 |
| `collapse_symptom.py` / `add_symptom_view.py` / `link_symptom_to_pitfall.py` | 根因收敛三步 | **有显式幂等判据**（已处理的节跳过） |
| `add_symptom_leadin.py` | 按 TSV 插导语 | 编号必须一一对应，错位就中止 |
| `add_toc.py` | 给长文档加目录 | |

> **为什么幂等要写成显式判据、不靠"文件已存在就跳过"**：
> 收敛脚本分三步跑，中间任何一步都可能重跑。
> 真正的幂等判据是**内容形态**（"正文里已有根因链接"），
> 不是文件名或时间戳。

---

## 改动这一层要过什么

| 改动类型 | 必过的门 |
| --- | --- |
| 改 `motion.ps1` | `-Strict` 走一遍，**必须真的开 PowerPoint** |
| 改任何 `probe_*.ps1` | 拿已知答案样本（`make_calibration.ps1` 产物）验证 |
| 改 `tools/` 里幂等脚本 | **连跑两遍**，第二遍必须输出"0 处改动" |
| 新增 `tools/` 脚本 | **先查这张表有没有现成的**（§六.8）→ 登记进 LAYERS → 说明它是一次性的还是常驻的 |

最后一条很重要：`tools/` 是开发期工具，**不进发布包**。
`split_*.py` 留着是为了照抄，不是为了让用户拿到手。

---

## 往上一层

| 我要做的事 | 去哪 |
| --- | --- |
| 加一种 COM 操作 | `motion.ps1`，注意 `WithWindow` 和路径长度两条规则 |
| 加一种探针实验 | `scripts/probe_*.ps1` + 对应的 `build_*.py` |
| 拆一个大文件 | 抄 `tools/split_*.py`，四条硬约束见 `CONTRIBUTING.md` |
| 改一次性工具 | 先确认它是不是幂等的 —— 不幂等的工具比没有工具更危险 |