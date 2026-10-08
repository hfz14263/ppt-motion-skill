# 交接文档

> 给接手这份工作的人（或下一个会话的 agent）。
> 先读本文，再读 `SKILL.md`（操作手册）与 `reference/review-checklist.md`（交付复核）。

---

## 1. 这是什么

`ppt-office-motion` —— 给**任意** `.pptx` 叠加动画、切换、3D 相机与内嵌音视频，
**而完全不改动它的版面几何**。

存在理由是上游工作流的一个缺口：版面刚被门禁校验通过，一旦用 PowerPoint
打开→编辑→保存，PowerPoint 会跑自己的排版引擎（字体解析、auto-fit 重排），
刚通过的重叠/溢出结论立刻作废。本工具的全部设计就是为了绕开这一步。

**核心承诺**：注入只碰三样东西 ——
`ppt/slides/slideN.xml` 的 `<p:transition>` / `<p:timing>`、`[Content_Types].xml`、
`ppt/media/*`。形状的 `<p:spTree>` 及其 `<a:xfrm>` **一个字节都不动**，
因此上游门禁结论可以**直接继承**，且能用几何指纹做机器自证。

---

## 2. 当前状态（已验证能力）

| 能力 | 状态 | 验证方式 |
| --- | --- | --- |
| OOXML 增量注入（137 个效果别名） | ✅ | `selftest` 63 断言 |
| 几何自证（布局帧字节不变） | ✅ | `motion.py apply --assert-geometry` |
| 结构自证（well-formed / spid / cTn id） | ✅ | `verify_motion.py` |
| 覆盖率审计（无"零效果"形状） | ✅ | `motion.py check` |
| 往返普查（PowerPoint 保存后效果没丢） | ✅ | `motion.ps1 -Strict` |
| **Morph（平滑）切换** | ✅ | `tests/test_morph.py` 20 断言 |
| **3D 相机（度数 + 半自动）** | ✅ | `tests/test_camera.py` 30 断言 |
| **窗口化图片填充（`fills:`）** | ✅ | `tests/test_fill_window.py` 28 断言 |
| 方向性擦除 `dir:` | ✅ | 上游 `selftest` |
| 艺术效果（虚化 / 亮度，走 COM） | ✅ | 与 template1 的 XML 逐字节同构 |
| 内嵌 MP4 / WAV | ✅ | `motion.ps1` 媒体层 |
| 从视频反推动效节奏 | ✅ | `tests/calib` 已知答案标定 |
| 动效预览（HTML 重放） | ✅ | `motion.py player` |
| 交付复核（含美感判据） | ✅ | `review_assist.py` + `review-checklist.md` |

**测试总览**：`selftest 63` + `test_morph 20` + `test_camera 30` +
`test_fill_window 28` + `smoke` 三场景，全绿。

---

## 3. 架构：四层，各管一件事

```
上游（版面）           dsh-ppt-studio / open-kimi-ppt / 任意现成 pptx
   │ 产出带稳定 elementId 的静态 .pptx
   ▼
① 注入层  scripts/motion.py         纯 OOXML，几何零改写
   │      读 motion spec，写 <p:timing> / <p:transition> / <a:scene3d>
   ▼
② 媒体层  scripts/motion.ps1        Office COM：AddMediaObject2 + 真渲染 + PDF
   │      唯一必须真 PowerPoint 的一层
   ▼
③ 复核层  verify_motion.py          结构自证
   │      review_assist.py          自动复核 + 明说判不了什么
   ▼
④ 设计层  reference/               规范与配方（不产生代码）
          motion-design-spec.md     该怎么动
          review-checklist.md       怎么验（含美感判据）
          morph-and-3d-recipes.md   三项技术配方
          design-system/            10 套静态版面设计系统
```

**为什么要分层**：每一层能证明的东西不同。①只证"XML 合法且几何没动"，
③才证"PowerPoint 采纳了"。把①的结论当③用，就会犯第 6 节那些错。

### 3.1 文档自身的"规模分层"（2026-09 整理）

四层架构管的是**知识**，还有一条**规模**上的规则：

> **一个文件长到"查它要翻屏"的时候，它就失效了 —— 因为查它的时刻
> 通常是你已经出错的时候。**

**但"太大"有两种，治法不同**：

| 症状 | 治法 | 什么时候用 |
| --- | --- | --- |
| **一个文件装了多件事** | **拆** | 名字对不上内容（`morph-and-3d-recipes.md` 装着素材派生和窗口填充）、或一节涨成一份文档（§51 / 2027 行） |
| **一份文档装了同族的很多节** | **加目录** | 内聚的层文档 —— 形态层/选择层是 L1–L4 的骨架，**拆开就把架构打散了** |

**判据是"读者能不能一眼知道里面有什么"，不是行数。**

#### 拆

| 原文件 | 拆成 | 拆法 | 引用怎么活下来 |
| --- | --- | --- | --- |
| `com-pitfalls.md` 2027 行 | 10 份 `pitfall-*.md` | 按**主题** | **编号不变** —— `§44` 永远指"切换的形态"；编号→文件表在 [`reference/pitfall-map.md`](reference/pitfall-map.md)（**脚本扫描生成，不是手抄**） |
| `symptoms.md` 1394 行 | 6 份 `symptom-*.md` | 按**现象** | **栏名不变** —— 目录页保留同名分流 |
| `morph-and-3d-recipes.md` 683 行 | 3 份 `recipe-*.md` | 按**主题** | **§ 编号不变**；目录页开头那张表就是 `§n` 的解释器 |
| `HANDOVER.md` §8 299 行 | 3 份 `history/*.md` | 按**主题**（不是时间） | 原编号保留；§8 只留**真·待办** |

三条设计取舍（都是踩过才知道的）：

1. **原文件名必须保留**（退化成目录页，不是删除）。`com-pitfalls.md` 有
   **91 处**引用写着它；删掉 = 91 个死链。**改引用比留入口贵得多，也更容易漏。**
2. **编号只增不减、只挪不改**（⚠️ 硬规矩）。编号一旦复用，历史引用会
   **静默**指向另一个坑 —— 比 404 更糟，因为它看起来是好的。
3. **日志和待办不能混在一节里**。§8 把"做完了什么"和"还剩什么"混在一起，
   结果 8 条里 7 条早已完成，读者却无法一眼看出还剩什么，只能逐条读 ✅。
   **混在一起的两类东西会一起失效。**

**拆分必须逐字节验证**，不能抽查：拆完的每一节与拆分前 `git show` 出来的
原文比对（51 节 + 10 栏 + 8 条 + 8 节，全部"不一致 = 0"）。

#### 加目录

`>=300 行且有 >=4 个标题` 的文档必须带目录 —— 由 `tools/add_toc.py` 加
（**幂等**：有 `<!-- toc -->` 标记就替换，不在人写的目录上叠）。
现有 10 份带目录，`tests/test_transition_table.py` 第 19 组守着这条线。

#### 工具与分发边界

拆分/加目录脚本放在 `tools/`，**不随技能包分发**
（`tests/test_install_manifest.py` 把它归入 `DEV_ONLY`）。
长期活着的只有两个生成器：`scripts/build_pitfall_map.py`（编号表）与
`tools/add_toc.py`（目录）。**映射表不能手抄** —— 手抄的表迟早和文件漂开，
而漂开时没人会发现。`history/` 则随包分发（`HANDOVER.md` 链着它，不发就是死链）。

---

## 4. 目录与关键文件

```
SKILL.md                      操作手册：S0–S8 流程、spec 全文、命令、九条坑
README.md                     项目说明与文档索引（对外）
INDEX.md                      ★ 唯一入口 —— 六个「我想 ____」+ 项目自身
HANDOVER.md                   本文：长期交接
CONTRIBUTING.md               ★ 变更规范 —— 体积上限 / 分诊 / 命名 / 登记（加东西前先读）
install.ps1                   安装到 DSH skill 目录 + 环境预检
requirements.txt              PyYAML / lxml

history/                      归档：已做完的事 + 当时怎么想错的（不是待办）
  README.md                   ★ 归档索引
  early-capabilities.md       早期能力建设（原 §8 的 1、3、4）
  doc-restructure.md          三层文档结构重整（原 §8 的 5）
  transition-layers.md        页面切换四层（原 §8 的 6、7、8）

scripts/
  motion.py                   ★ 核心引擎：注入 / 预览 / 覆盖率 / spec 解析
  motion.ps1                  ★ Office COM 层：媒体 + 真渲染 + 往返普查
  verify_motion.py            结构自证
  review_assist.py            ★ 自动复核（结构可信 + 启发式标注）
  player.py                   动效 HTML 重放
  check_coverage.py           覆盖率审计（供 motion.py check 调用）
  inspect_pptx.py             判断一个 pptx 到底有没有动画
  selftest.py                 63 断言回归
  analyze_video.py            从视频量动效节奏
  make_calibration.ps1        生成已知答案的标定样本
  probe_createvideo.ps1       探测本机 CreateVideo 是否可用
  build_camera_table.py       3D 相机参数实测表生成
  build_transition_table.py   切换实测的**入口**（19 个子命令；实体在下面的包里）
  transition_probe/           ★ 探针工具包（2026-09-30 从上面那份 3048 行单文件拆出）
    common.py                 共享常量 + 通用小工具（含 SCRIPTS_DIR）
    data.py                   探测数据：假设表与各 PROBE 名单
    decks.py                  deck 构造（假设 → slide XML，纯设置）
    analysis.py               帧分析与表构造
    commands.py               19 个 CLI 子命令（名字是接口，不能改）
  build_shape_evidence.py     形态层正确性证据（左栏判定 + 右栏真实渲染 GIF）
  build_pitfall_map.py        ★ `§n` → 文件的映射表**生成器**（表不能手抄）
  verify_docs.py              ★ 文档结构校验：悬空引用 / 孤儿 / 入口预算
  check_structure.py          ★ 结构体检：体积上限 / 索引 / 目录登记（**只读**）
  push_via_api.py             github.com 被挡时走 API 推送
  vendor_themes.py            批量下载上游设计系统
  motion_catalog.json         137 个效果的真实 presetID + XML 模板

tools/                        **不随包分发** —— 开发期一次性脚本（见 CONTRIBUTING §五）
  add_toc.py                  给长文档加目录（幂等 + 标题层级自适应）
  split_pitfalls.py           踩坑手册 → 10 份（一次性）
  split_symptoms.py           症状文档 → 6 份（一次性）
  split_morph_recipes.py      配方文档 → 3 份（一次性）
  archive_handover_backlog.py HANDOVER §8 → history/（一次性）
  rewrite_*.py                拆分后改写目录页（一次性）

reference/
  authoring-rules.md          13 条硬约束（每条对应一次真实翻车）
  com-pitfalls.md             ★ 踩坑**目录页** —— 按"你现在的处境"分流到下面 10 份
  pitfall-map.md              ★ `§n` 的解释器：编号 → 文件（拆分后 90+ 处引用靠它活）
  pitfall-file-corruption.md  踩坑 · 文件打不开 / 报损坏（§1 §2 §4 §12 §26 §31 §40 §42）
  pitfall-silent-drop.md      踩坑 · XML 里有、PowerPoint 不认（§13 §19 §41 §49 §50）
  pitfall-roundtrip.md        踩坑 · 往返后被改掉或吃掉（§14 §18）
  pitfall-ooxml.md            踩坑 · OOXML 注入写法（§3 §8）
  pitfall-com.md              踩坑 · COM 与脚本环境（§5–§7 §9–§11 §15–§17）
  pitfall-morph-3d.md         踩坑 · Morph / 3D / 图片填充（§20–§25 §27 §29 §30 §34）
  pitfall-measurement.md      踩坑 · 测量与判据（§28 §33 §38）
  pitfall-media.md            踩坑 · 素材与观感（§35–§37 §39）
  pitfall-transition.md       踩坑 · 页面切换（§43 §44 §46 §47 §48 §51）
  pitfall-tooling.md          踩坑 · 工具链（§32 §45）
  symptoms.md                 ★ 症状**目录页** —— 按现象分流到下面 6 份
  symptom-file-corruption.md  症状 · 文件打不开 / 报损坏
  symptom-silent-drop.md      症状 · XML 里有、PowerPoint 不认
  symptom-wrong-result.md     症状 · 文件正常但结果不对
  symptom-roundtrip.md        症状 · 往返后效果被吃掉
  symptom-tooling.md          症状 · 工具与环境
  symptom-evidence.md         症状 · 证据与判断（值越界 + 四个"我以为我看见了"）
  review-checklist.md         ★ 交付复核清单（判据可信度分级 + 美感判据）
  motion-design-spec.md       动效设计规范
  morph-and-3d-recipes.md     ★ 配方**目录页** —— §n 的解释器（原 683 行，见下 3 份）
  recipe-morph-camera.md      配方 · Morph 与 3D 相机 + 两个测量陷阱（§1 §2 §3）
  recipe-asset-derivation.md  配方 · 素材派生（模板提取 / 艺术效果 / 照片复现，§4 §6 §7）
  recipe-fill-window.md       配方 · 窗口化图片填充（形状切割，§8）
  recipe-library.md           **原理**库（按原理不按效果）—— 与上面三份"具体配方"不同
  camera-reference.md         62 个相机预设的参数→实测对照
  video-analysis-limits.md    视频分析的实测能力边界
  architecture.md             为什么这样设计
  template-patterns.md        给静态模板从零设计动效
  mso-primitives.md           效果与切换的原始语义
  ppt-studio-integration.md   与 deck.yaml + elementId 的对接契约
  design-system/              10 套内置设计系统 + 44 套索引

tests/
  test_morph.py               20 断言：Morph 注入
  test_camera.py              30 断言：3D 相机（度数换算/损坏值防护/插入位置）
  test_fill_window.py         28 断言：窗口化填充（命名空间/兄弟顺序/单一填充/算术）
  smoke.py                    三场景端到端（pptd 导出 / PowerPoint 原生 / 真开）
  privacy_audit.py            隐私与泄漏审计（文本 + 容器内 + git 历史）
  fixtures/                   回归样本
  motion.yaml, fx-pro.pptx    回归 spec 与样本

examples/
  motion.example.yaml         spec 全字段示例
  material/                   ★ 原始素材：两份参考模板 + 作业笔记（上面那些结论的推导来源）
```

---

## 5. 快速上手

```bash
pip install -r requirements.txt
powershell -NoProfile -File install.ps1 -Force     # 装进 DSH skill 目录
```

```bash
# 看清 deck：有哪些形状、什么 id、已有哪些动画
python scripts/motion.py inspect --pptx deck.pptx

# 注入 + 几何自证
python scripts/motion.py apply --pptx in.pptx --spec m.yaml --out out.pptx --assert-geometry

# 结构自证
python scripts/verify_motion.py --pptx out.pptx --source in.pptx

# 覆盖率 / 往返普查 / 真渲染
python scripts/motion.py check --pptx out.pptx --spec m.yaml
powershell -NoProfile -File scripts/motion.ps1 -Pptx out.pptx -Spec m.yaml -OutDir review -ExportPdf -Strict

# 动效预览（唯一能"看见动"的手段，静态图看不出来）
python scripts/motion.py player --pptx out.pptx --spec m.yaml --outdir preview

# 交付复核
python scripts/review_assist.py --pptx out.pptx --source in.pptx
```

spec 字段全表见 `SKILL.md` §3，含 `cameras:`（3D 相机，度数为单位）与
`transition: {type: morph}`。

---

## 6. ⚠️ 本项目最重要的部分：四次真实误判

**这四条比任何功能说明都重要。** 它们形态完全相同 ——
**拿单一自动化判据当结论，没做交叉验证** —— 而症状都是"自信地给出错误答案"
（第 4 条是"自信地给出**通过**"）。

| # | 当时结论 | 实际 | 根因 |
| --- | --- | --- | --- |
| 1 | "本机不支持 Morph" | **支持**，`EntryEffect=0xF72` | 元素名写成 `p:morph`，正确是 `p159:morph`；查 COM 枚举自然查不到 |
| 2 | "template2 两页几何完全相同、不产生补间" | 矩形组 left 从 **−180.8** 变到铺开 | 解析脚本**没处理 `<p:grpSp>`**，读到组变换而不是子形状 `<a:off>` |
| 3 | "原件里人物不在该位置"（模板匹配 0.401） | 人物**就在**那里 | 模板匹配在低对比/缩放差异下不可靠 |
| 4 | "隐私审计 CLEAN" | 一个**已发布**的回归样本里带着真实姓名 | 审计 `SKIP_EXT` 跳过 `.pptx`/`.png`，而 `.pptx` 是 ZIP、PNG 像素是压缩流 —— **字节搜索根本看不见里面** |

**第 1 条的完整教训**：查不到某个能力时，先怀疑自己写错 ——
"对象模型里没有"**不等于**"不能注入"。注入是写 XML，与属性模型是两条路。

**第 2 条的完整教训**：断言"没有差异"之前，必须先证明**解析器看得见差异**。
把已知有差异的样本喂进去，看它是否报不同。

**第 4 条的完整教训**：说 CLEAN 之前，先证明审计**看得见**它要查的东西。
把二进制排除在文本扫描之外是对的（否则满屏误报），但排除之后就**再没有东西**在看它们了。
现在 `privacy_audit.py` 有第二遍 **container scan**：解压 OOXML 条目、只读 PNG 的
**文本块**（`tEXt/iTXt/zTXt/eXIf/tIME`），并且**先做正向对照** ——
把名字注回去，确认它报警，再还原。

> 顺带一个反直觉的点：PNG 里"搜到"姓名往往是**假警报** ——
> 那几个字节落在 `IDAT` 压缩流里，纯属巧合。所以扫描器只读文本块，
> 不读像素。同理，`.pptx` 上直接 `grep` 永远是 0 命中，无论里面写了什么。

**因此 `reference/review-checklist.md` §5 的三条规则不是形式主义**：

1. 任何"通过"结论要**两个独立判据**（"几何没变"+"结构合法"不算，都出自同一份 XML）
2. 机器判"没有 / 没变"时，必须找**正向判据**佐证
3. 涉及**分组 / 嵌套 / 扩展命名空间**的解析，必须与**对象模型交叉核对**

**判据可信度分级表**（在复核清单 §0）请务必先看：它写明了每条判据能证明什么、
**已知在哪失效**。越便宜的判据越弱，不要用上层结论替代下层验证。

---

## 7. 环境依赖与已知限制

**必需**
- Python 3.7+，`PyYAML`（YAML spec）、`lxml`（校验闸）
- Windows PowerShell 5.1+（COM 层）
- Windows + **Microsoft PowerPoint 桌面版**（媒体内嵌、真渲染）；
  **WPS 不行** —— 不暴露同一套 COM

**可选**
- `opencv-python` + `numpy`：视频动效分析
- Pillow + numpy：`review_assist.py` 的素材卫生检查
- Chromium/Edge：截图类辅助

**已知限制**
- **同一形状上"入场 + 强调"不兼容 PowerPoint 往返**（本机实测限制，未解决）。
  要叠就用不同形状。
- **`Presentation.CreateVideo` 的可用性随 Office 构建变化**，不要套用别人的结论；
  换机器先跑 `probe_createvideo.ps1`。
- **单页内"扫描"效果**：窗口带着填充一起移动是"平移"而非"扫描"；
  真扫描要把图片单独放底层、窗口当遮罩。
- **静态导出 ≠ 动态效果**：入场动画在导出图里是隐藏的，动作路径显示起点。
- **`grabCut` 是经典分割**：需要矩形提示，人物与背景同色/背景杂乱会失效。
  生产环境建议换专门的人像抠图/生成式 AI。
- 本机 `.ps1` 若含中文，**必须写 UTF-8 BOM**，否则 PowerShell 5.1 按 GBK 读会打乱字符串。

---

## 8. 还没做的

**只有这一条是真·待办** —— 其余 7 条都已完成，归档在 [`history/`](history/README.md)。

2. **多 API 协作** —— 素材生成（尤其抠图/插画）接专门的图像生成 AI，
   替代 `grabCut` 这类兜底方案。

### 8.1 已经做完的 → 归档在 [`history/`](history/README.md)

这一节原本有 **299 行**（占全文 593 行的一半），8 条里 **7 条已打 ✅**。
它实际上是一份**开发日志**，不是待办清单 —— 而日志留在概览文档里，
代价由每个只想读概览的人支付。

> **日志没错，放错地方才是。** 所以搬进 `history/`：
> 想查历史的人去查，想读概览的人不被挡。

**搬走时逐条原文照搬，一个字没改** —— 那些记录的价值全在
「当时为什么想错了」的细节上，改写就等于毁证。
三份归档**按主题分**（不是按时间）：同一条线的工作横跨好几天，
按日排会让人来回跳。

| 归档 | 装什么 | 原编号 |
| --- | --- | --- |
| [`history/early-capabilities.md`](history/early-capabilities.md) | 美感判据执行化 · `fills:` spec 扩展 · 按原理组织的配方库 | §8 的 1、3、4 |
| [`history/doc-restructure.md`](history/doc-restructure.md) | 三层文档结构重整（文档曾长到 27 文件 / 261k 字符） | §8 的 5 |
| [`history/transition-layers.md`](history/transition-layers.md) | 页面切换四层：数据表 → 机制 → 形态 → 选择 | §8 的 6、7、8 |

> 用**原编号**而不是行数：行数会漂，编号不会。归档文件里每条都标着它原来的
> 编号，所以「§8 第 7 条」永远指得动。

**这一节本身也是个教训的实例**：把日志和待办混在一起，
结果是**待办被淹没在历史里** —— 8 条里 7 条早就做完了，
但读者无法一眼看出"还剩什么"，只能逐条读 ✅。
**混在一起的两类东西，会一起失效。**

### 8.2 验证文化（请保持）

- **加任何东西之前先读 [`CONTRIBUTING.md`](CONTRIBUTING.md)** ——
  体积上限、拆/加目录/归档的分诊、新目录必须登记两处。
  它和 `check_structure.py` 的数字由测试第 20 组钉在一起，**不会各说各话**。
- 改 OOXML 后**必须**过 PowerPoint 真开一次 —— 不是可选步骤
- 改分析器/判据后**必须**用已知答案样本对表
  （`make_calibration.ps1` 生成标定视频；`tests/` 里有回归样本）
- 提交前跑：`selftest` + `test_morph` + `test_camera` + `test_dual_photo` + `smoke`
  + `privacy_audit` + **`verify_docs`** + **`test_install_manifest`**
  + **`test_transition_table`** + **`check_structure`**
  （`test_dual_photo` 是版式校验器「同图双版本」的回归，自带正负对照，不需要 PowerPoint）
- 后两个是 2026-09-28 复验时才补的，补它们是因为**前六个一项都没报错、问题却真的存在**：
  - `verify_docs` 新增第 5 条规则：**旧目录名不得作为路径存活**（旧名是复数拼写）。
    markdown 链接检查看不见 `os.path.join(ROOT, <目录名>, ...)` 这种裸字符串写法 ——
    2026-09-28 的目录统一就漏了 6 处，只有真跑脚本才暴露
    （`verify_recipes` 直接 FileNotFoundError）。
  - `test_install_manifest` 守 `install.ps1` 的复制清单：三层结构后的 `INDEX.md`
    与 `facts/` **从来没进过清单**，装出去的技能包没有入口页也没有唯一真相源。
- **教训**：改路径之后，光 grep `旧名/` 不够 —— 要 grep 裸字符串，更要**跑一遍所有脚本**。
- `test_transition_table` 也是 2026-09-28 补的：切换表是**量**出来的，`motion.py` 直接读它。
  两者一旦漂移，症状不是报错，而是**写出一个 PowerPoint 会悄悄改掉的块** ——
  文件一切正常，内容就是不对。这种漂移只有对照测试能发现。
- 再一条同类教训：**"全部失败"要怀疑生成器，不是被测对象**。第一次探测 48 个切换候选
  **全部**被拒开，包括最普通的 `<p:fade/>` —— 全拒说明问题不在候选，一查是我们自己把
  `xmlns:a` 在 `<p:sld>` 上声明了两次（python-pptx 已声明 a/p/r），part 不合法。
  **症状的粒度本身就是信息**：一个失败 = 那个用例有问题；全部失败 = 生成用例的代码有问题。
- 第三条同类：**"能打开"离"没被改过"很远**。机制层那两个 bug 写出来的块
  （`spd="{spd}"`）PowerPoint 全都能开，还按自己的规则补成合法值 ——
  任何结构检查、解析、回归断言都通过。**只有把"写入的文件"和"PowerPoint 存回的文件"
  逐项比对**才看得见（`scripts/probe_roundtrip.ps1`）。
  改 OOXML 生成侧时，往返对比不是可选项。
- 第四条：**别用整帧均值判断过渡，也别把"某个元素的合法写法"推广到另一个元素。**
  前者让 1.5s 的 wipe 整段隐形、也让 wipe 冒充淡入（要看空间剖面）；
  后者让我照着 morph 存回来的数字形态给 `spd` 写了个 `spd="800"`，
  直接让 PowerPoint 拒开整份 —— 一个元素的 saved 形态不构成另一个元素的契约。
- **发现之前写错了 → 就地打补丁，不要重写 / 新开"修正版"。**
  用户定的守则，写进 [`CONTRIBUTING.md`](CONTRIBUTING.md) §五。
  起因正是上面 §8.3 与 §8.4 的冲突：结论错了就**在原地改那句**，
  再留一行"当时为什么判断错了"。**两份说法比一份错更糟** ——
  它和本项目最怕的"一处说 X、另一处说 not-X"是同一类病。
  **唯一例外**：`history/` 里的归档不补丁（那些是"当时为什么想错"的记录，
  改它等于毁证）。要更正的是**当前的结论**，不是**历史的事实**。

### 8.3 无人值守：能做哪一半，为什么另一半不能

有用户问过：**"这套整理能不能在我下班后自动跑完？"**
**答案：检查能，改动不能。**

| 动作 | 无人值守 | 为什么 |
| --- | --- | --- |
| 跑体检（`check_structure` / `verify_docs` / 各测试） | ✅ **能** | **只读**。最坏结果是打一份报告 |
| 按报告去拆 / 归档 / 重构 | ❌ **不能** | 见下面三条 |

**三条理由（不是"我做不到"，是"这么做会出事"）：**

**① 这些动作本质是删除。** "拆"把 683 行的文件变成 74 行；"归档"把 299 行
从概览里搬走。**内容没丢**（每一步都有逐字节校验），但**读者看到的东西变了** ——
而"变得对不对"要靠人看。在无人确认时执行删除类操作，
等于把不可逆的那部分留给一个不在场的人。

**② 有些判据在这里根本跑不了。** 本项目的验证有一半依赖**真 PowerPoint**
（渲染帧、往返比对）—— 那部分流水线在无头环境里跑不通。
任何**改变了代码路径**的重构，都会在"验不到的那一半"留下盲区。

> ⚠️ **这条我一开始用错了，见 §8.4。** 它说的是"跑不通流水线"，
> 不等于"验不了搬家"：**纯搬运不改变行为**，要验的是模块结构（import /
> 名字 / CLI 分发 / 测试），那四件本地全能验。我当时把两者混为一谈，
> 因此把一个**可以做**的拆分判成了"不能做"。
> **教训：判"能不能自动做"之前，先分清这是"改变逻辑"还是"改变位置"。**

**③ 分诊是判断题，不是计算题。** 同一份 600 行文档，"该拆"还是"该加目录"
取决于**它是不是某一层架构的骨架** —— 这需要理解上下文，
而体检报告只能给出体积和标题数。`tools/add_toc.py` 能算目录，
**算不出"这份该不该拆"**。

**所以约定是：**

- **定时任务只跑检查**，把结果写进报告（`check_structure.py --json`），
  **绝不动文件** —— 这样它半夜出错的代价只是报告写歪。
- **改动一律在有人时做，一步一提交** —— 拆 / 归档 / 加目录各一个 commit，
  出问题能精确回滚。
- **豁免不会悄悄消失**：写进例外清单的条目，每次体检都以 ADVISE 重新报出来，
  直到待办被清掉（见 `CODE_SPLIT_EXEMPT` 的 `todo` 字段）。

**当前挂着的豁免**：

| 项 | 状态 |
| --- | --- |
| ~~`scripts/build_transition_table.py` 3048 行~~ | ✅ **已拆**（2026-09-30）→ `transition_probe/` 包五层。豁免已撤销，见 §8.4 |
| `reference/symptom-*.md` 与 `reference/pitfall-*.md` 内容重叠 | ⏸️ **挂起**（2026-10-08，用户决定）。重叠量已测清（见下），但**先不动** —— 等对项目结构的认知理清之后再定|

> **为什么挂起而不是删掉这条记录**：测出来的数字不会过期，
> 而"当时为什么没立刻处理"也是上下文的一部分。

### 8.3.1 这条到底重叠在哪（2026-10-08 实测，取代原先"50% 段落"的粗略说法）

先说清两族是什么关系 —— 它们**不是两批内容，是同一批内容的两种排序**：

- `pitfall-*`（11 份 / 51 节 / 2251 行）按**「你现在的处境」**分流：
  文件损坏 / 静默丢弃 / 往返 / OOXML / COM / Morph-3D / 测量 / 素材 / 切换 / 工具链
- `symptom-*`（6 份 / 42 节 / 1527 行）按**「你看到的现象」**分流：
  打不开 / 不认 / 结果不对 / 往返被吃 / 工具环境 / 证据不足

**症状页的 42 个编号全部来自踩坑页**（`§1`–`§42`，零个自有编号），
每个节尾都挂着 `<sub>来源 com-pitfalls §n</sub>`。踩坑页有 9 个编号（`§43`–`§51`）
不进症状页 —— 它们是**规则与机制**，不是现象。

**重叠的准确形态**（不是"50% 段落逐字重复"这么粗）：

| 形态 | 节数 | 说明 |
| --- | --- | --- |
| 整节逐字相同 | **0** | —— |
| 去掉代码块后**全部**段落都在对方那里 | **24** | 差异**只在代码块与标题措辞** |
| 有症状页自己的段落 | **18** | 详见下|

**字符量几乎相同**：症状页 42 节正文 35,637 字符 vs 踩坑页对应 42 节 35,782 字符
—— **压缩率 100%**。症状页**没有真正瘦身**。

所以真实代价不是"删掉一半正文"，而是：**同一段解释被写了两遍，
改一处要改两处，漏一处就自相矛盾** —— 这正是本项目最怕的
"一处说 X、另一处说 not-X"。

**18 节独有内容的规律**（这是 (a) 方案的可行性依据）：
独有段落几乎都是**症状视角的引导语** —— 先说"你看到什么"、再说"根因在§n"。
例如 `§41`：

- symptom 版：*"16 个 core 切换往返后全被判「被改写」，但子元素一个都没变。原因在 Fallback：…"*
- pitfall 版：结构化成 **症状 / 根因 / 修法** 三段，还给了 ✗/✓ 对照的完整 XML

**两个方向**（都需要人拍板，我不动）：

- **(a) 症状页瘦身成转诊页** —— 只留"现象 → 根因在 §n"，正文只留在踩坑页。
  **省下的其实不多**（压缩率 100%，能省的主要是那 24 节里重复的解释段），
  但会让症状页从"可读"退化成"索引"，与它存在的理由
  （*"查症状的时刻通常是你已经出错的时候，那时最不该付出的成本是找"*）
  **正面冲突**。属删除类。
- **(b) 承认这是同一内容的两种排序** —— 保留两份，但加一条机器检查：
  凡症状页出现 `§n`，该`§n` 必须存在于踩坑页，且两边的关键结论不许分叉。
  **不删任何内容**，代价是多一条检查要维护。

> ⚠️ 我不替你选。**(a) 与症状页的存在理由冲突**，(b) 承认重复是长期负担 ——
> 这是内容决策，不是技术决策。

### 8.4 上面那条「不能自动做」的结论，有一半是错的（更正）

`build_transition_table.py` 当初被我判为"不能拆"，理由写在上面 §8.3 的 ②：
*"搬完 3000 行只验得了一半的路径，验不全的重构不能自动做。"*

**那个理由说对了一件事，说错了一件事**，值得留档：

- ✅ **对**：那段代码的**探测流水线**（deck → PowerPoint 渲染 → 分析帧）
  确实要真 PowerPoint，本机跑不通。
- ❌ **错**：把"跑不通流水线"当成了"验不了搬家"。**这次做的是纯搬运** ——
  把整块函数从 A 文件挪到 B 文件，**行为按构造不变**。需要验证的不是流水线，
  是**模块结构**：import 能否解析、名字全不全、CLI 分发对不对、测试过不过。
  **这四件本地全能验。**

**实际拆完的验证（全部本地通过）**：

| 检查 | 结果 |
| --- | --- |
| 顶层名字（拆前 AST 扫出 vs 拆后可导入） | **64 / 64** 一致 |
| 逐符号原文比对（拆前 `git show` vs 拆后） | **64 / 64 逐字节相同**，无增无减 |
| CLI 子命令（写在 `facts/` 与 `reference/` 里的） | **19 / 19** 全在 |
| 测试套件 | 全绿；`check_structure` **0 FAIL** |

**但确实有一类真实风险，而且被抓到了**：
**代码依赖自身位置时，搬运会改变行为。** 实测撞了一次 ——
`decks.py` 里用 `os.path.dirname(__file__)` 拼 `motion_catalog.json`，
在原单文件里它恰好住在 `scripts/` 下所以是对的；搬进 `transition_probe/`
之后，**同一个表达式指向了一个不存在的地方**。

- 已改用 `common.SCRIPTS_DIR`（明确指 `scripts/`，不随模块位置漂）。
- ⚠️ **这类错静态检查完全看不见**，只能靠跑测试。所以：
  **搬代码前先 grep `__file__` 与相对路径拼装** —— 这是纯搬运唯一的真陷阱。

> **"纯搬运是安全的"这个判断，95% 成立。剩下 5% 就是这一条。**
> 它不会让搬运变得不值得做，但它意味着**搬完必须跑测试**，不能只看 import 通不通。

---

## 9. 来源与许可

- 本 skill 基于上游 [`hfz14263/ppt-motion-skill`](https://github.com/hfz14263/ppt-motion-skill)
  的成熟版本演进（该线已有 63 断言自测、13 条 authoring-rules、62 相机实测表、
  `player.py`）。本仓库在其基线上叠加：动效设计规范、设计系统、视频分析、
  交付复核清单、camera/morph spec 支持。
- `reference/design-system/` 的 10 套设计系统来自
  [`open-kimi-ppt-skill`](https://github.com/acnlie/open-kimi-ppt-skill)（**MIT**，
  Copyright (c) 2026 Binaryify Zhuang），逐字保留并附 `LICENSE-upstream.txt`。
- 本项目 MIT。

**注**：仓库 git 历史曾做过一次改写（统一为
`hfz14263 <hfz14263@users.noreply.github.com>`），因此全部提交署同一身份。
改写原因见第 6 节同一套"先自查"原则 —— 早期提交信息里带了机器路径，
且作者邮箱是个人邮箱。改写前的完整备份在仓库外（`*-BACKUP-pre-rewrite.bundle`），
确认无需回滚后可删除。
