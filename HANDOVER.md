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

### 3.1 文档自身的"规模分层"（2026-09 拆分）

四层架构管的是**知识**，还有一条**规模**上的规则：

> **一个文件长到"查它要翻屏"的时候，它就失效了 —— 因为查它的时刻
> 通常是你已经出错的时候。**

触发点很具体：`com-pitfalls.md` 涨到 §51 / 2027 行，`symptoms.md` 1394 行。
两份都被拆了：

| 原文件 | 拆成 | 拆法 | 引用怎么活下来 |
| --- | --- | --- | --- |
| `com-pitfalls.md` | 10 份 `pitfall-*.md` | 按**主题**（文件损坏 / 静默丢弃 / 往返 / …） | **编号不变** —— `§44` 永远指"切换的形态"；编号→文件表在 [`reference/pitfall-map.md`](reference/pitfall-map.md)（由 `scripts/build_pitfall_map.py` 扫描生成，不是手抄） |
| `symptoms.md` | 6 份 `symptom-*.md` | 按**现象**（打不开 / 不认 / 结果不对 / …） | **栏名不变** —— INDEX 和各处写着 "`symptoms.md` → 「文件损坏」"，目录页保留同名分流 |

两条设计取舍值得记下来：

1. **原文件名必须保留**（退化成目录页，不是删除）。全仓库 90+ 处引用写着
   `com-pitfalls.md`；删掉 = 90+ 个死链。**改引用比留入口贵得多，也更容易漏。**
2. **编号只增不减、只挪不改**（⚠️ 这条是硬规矩）。一个编号永远指同一件事，
   哪怕它换了文件。编号一旦被复用，历史引用会**静默**指向另一个坑 ——
   比 404 更糟，因为它看起来是好的。

`tests/test_transition_table.py` 第 16/17 组守着这件事：编号完整、编号唯一、
映射表与真实归属一致、目录页链全了族文件。**拆分本身也做了逐字节校验**
（拆完的每一节/栏与拆分前 `git show` 出来的原文比对，51 节 + 10 栏全部一致）。

拆分脚本（`split_pitfalls.py` / `split_symptoms.py`）放在 `tools/` ——
它们是一次性搬迁工具，**不随技能包分发**（`tests/test_install_manifest.py`
把 `tools/` 归入 DEV_ONLY，与 `tests/` 同级）。长期活着的只有映射表生成器。

---

## 4. 目录与关键文件

```
SKILL.md                      操作手册：S0–S8 流程、spec 全文、命令、九条坑
README.md                     项目说明与文档索引
HANDOVER.md                   本文
install.ps1                   安装到 DSH skill 目录 + 环境预检
requirements.txt              PyYAML / lxml

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
  vendor_themes.py            批量下载上游设计系统
  motion_catalog.json         137 个效果的真实 presetID + XML 模板

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
  morph-and-3d-recipes.md     Morph / 3D 相机 / 艺术效果 / 形状切割 配方
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

## 8. 待办与建议方向

**已识别但未做**
1. ~~**美感判据的执行化**~~ ✅ **已做（`scripts/design_audit.py`）** ——
   `motion-design-spec` §6 的五条硬上限与时长/切换判据已可执行，输出分
   FAIL / ADVISE / HUMAN 三类，并把"判不了"的十项每次打印。
   **过程中自己踩了一次同类错**：第一版把限值套在"组"上，而 §6 原文管的是
   "观众会等动画"即**逐个揭示**；一串 `after` 在 PowerPoint 里只算一组，
   于是 14 页被误报为违规。判据的定义本身要先验证。
   **§4 的页型剧本仍未执行化**（它需要知道每页的页型，spec 里没有这个信息）。
2. **多 API 协作** —— 素材生成（尤其抠图/插画）接专门的图像生成 AI，
   替代 `grabCut` 这类兜底方案。
3. ~~**morph / 3D 之外的 spec 扩展**~~ ✅ **已做（`fills:`）** ——
   单页内动画的"形状切割图片"不再需要手写注入：窗口几何入 spec，
   扫描动画本来就是 `wipe` / `pathRight`（早已支持）。
   **过程中一个函数里踩出四种静默失败**，全部产物都能正常打开、渲染却错：
   命名空间写成 `p:blipFill`（应为 `a:`）、插在 `<a:ln>` 之后、没删旧填充（填充是
   `xsd:choice`，取第一个）、`r:embed` 是编的 id 画占位符。详见
   `reference/com-pitfalls.md` §29–30 与 `tests/test_fill_window.py`（28 条）。
   **结论：正确性是让 PowerPoint 自己写一遍对照出来的，不是推出来的。**

4. **动效配方库（按原理，不按效果）** —— `reference/recipe-library.md` +
   `recipes.json`，6 条已实测原理（页间差异 / 动遮罩不动图 / 形状拼切整图 /
   平面→立体是两页 / 手写中间帧 / 图×形状是设计层）。
   每条带 `constraints`、**`breaks_how`（坏掉时的可见症状）**、`evidence`。
   **为什么按原理**：按效果组织是学不完的路；原理有限，新素材从原理推导。
   `scripts/verify_recipes.py` 校验结构，**并强制 `breaks_how` 写症状、`evidence`
   指向真实文件** —— 防止库退化成效果清单或传闻集。
   **下一步**：遇到新素材时按 `recipe-library.md` §3 的顺序做，加条目前先确认
   它是不是已有原理的实例。

5. ~~**三层文档结构重整**~~ ✅ **已做（2026-09-18 `fac93ec` 开启，2026-09-28 收尾）** ——
   动机：文档曾长到 **27 文件 / 261k 字符**，超过任何模型的上下文。这不只是"读不完"：
   没人能同时持有全部内容，于是**一处说 X、另一处说 not-X 的矛盾根本不可能被发现**
   （同一轮里就发生了：`lat/lon` 值域出现在三个文件、一个"静默丢弃"概念散在十处）。
   结构：
   - **L1 `INDEX.md`** —— 唯一入口，按"我想____"分六个入口，**预算 6000 字符**
     （现 4257，由 `verify_docs` 报出，不要手抄）。入口一旦需要翻页就不再是入口。
   - **L2 `facts/*.json`** —— 权威来源 + 跨来源规则，机器可校验。
   - **L3 `reference/*.md`** —— 细节**只写一次**，正文在这里、别处只引。
   收尾时补的两件事：目录名统一为单数 `reference/`（全局引用已改，
   `scripts/verify_docs.py` 的 `SCAN_DIRS` 也去掉了复数兼容项，防止文件被写回旧目录后
   仍被判为可达）；`reference/symptoms.md` 补齐 **§32–§39**，至此症状表与
   `com-pitfalls` 正文 **§1–§39 一一对应**。
   **闸**：`python scripts/verify_docs.py` 查悬空引用 / 孤儿文档 / 入口体积，
   现跑为 OK。**它只证明结构自洽，不证明内容正确。**
   **教训（为什么这项会被漏掉半截）**：规划只存在于提交信息里，没进 HANDOVER、
   没开分支、没建文档，两周后就查不到"还有个尾巴"。**跨会话的规划必须落在
   可检索的文件里**，提交信息不算。

6. ~~**切换效果实测对照表**~~ ✅ **已做（2026-09-28）** ——
   界面库 48 项（细微 12 + 华丽 29 + 动态内容 7，不含"无"）逐项量出
   **界面名 ↔ PowerPoint 写出的子元素 ↔ 完整 XML 块 ↔ 实测 COM 枚举**。
   产物：`scripts/transition_reference.json`（机读）+ `reference/transitions.md`（L3）
   + `facts/transitions.json`（L2，不存数据）+ 新增 `com-pitfalls` §40–42 与
   `symptoms` §40–42。
   **为什么必须量**：COM 枚举是 2003 年代的，本机与直觉差很远（`0x0A01` 名曰 fade
   实际产出 `<p:strips/>`），且**多个元素撞成同一个值**；界面分组 ≠ XML 命名空间；
   默认值没人写下来。**旧的 `motion.py` 手写 14 个模板，每一个都有错。**
   **方法（照抄相机的做法）**：两次测量分工 ——
   *枚举扫描*（一页一个 `PpEntryEffect`，COM 设上去保存读回，扫完 1..4200，
   3956 以上无切换）解决"PowerPoint 用哪个元素表达哪个值"；
   *往返探测*（一个候选一份 deck）解决"手写会被保留/改写/丢弃/拒开"。
   **往返只能说有效或拒开，说不出正确的名字** —— 5 个猜错的（切换/翻转/库/摩天轮/传送带
   裸写被拒开整份文件）是靠枚举扫描找回来的，它们必须带 `dir="l"`。
   **顺手改掉的**：`motion.py` 的 13 个手写模板改为读实测表（行为不变，已用真机
   往返验证：4 个切换全部原样保留、COM 读回 3849/3855/2817 与表一致）；
   `mso-primitives.md` §5 的旧小表（枚举列全错）退化为指向新表。
   **闸**：`python tests/test_transition_table.py`（表与 `motion.py` 不漂移、
   块合法且自带命名空间、Choice+Fallback 只算一个切换、重复注入不累积、
   `spd` 取档边界、带 `advTm` 后无占位符残留）。

7. **切换的机制层（2026-09-28，数据表之上的一层）**
   用户原话："下一步做的则是类似于我们在平滑中总结的：第一页是初始状态，
   第二页是终态，中间依赖平滑补充。这种理解性的内容，指导这些方法使用。"
   产出一份 `reference/transition-model.md`（四条模型 + 决策表），
   `com-pitfalls` §43、`symptoms` §43.1/§43.2/§43.3、`facts/transitions.json` 5 条规则。
   **四条模型都是从渲染帧量出来的，不是从文档读的**：
   - ① `<p:transition>` **挂在终点页**（写在第 N 页 = 动「进入第 N 页」）。
     **读 XML 分不出两种解释**（两种下"切换在第 N 页上"都成立），
     所以用一份能推翻它的 deck：3 页只有第 2 页写 push 800ms →
     burst 落在边界 1（1→2）而非边界 2。第 1 页的切换则只用在放映起步。
   - ② `<p:transition>` **只有一个槽位**：塞两个子元素（core+core 与 core+p159
     两组都试）→ **整份文件拒开**。所以「形状形变」与「整页位移」不可同时表达。
   - ③ `p14:dur`（毫秒）压过 `spd`；让二者打架即见分晓（200ms 写 spd="slow" →
     渲染 200ms）。**但 Fallback 里没有 `p14:dur`**，那里只认 `spd`，
     实测 slow≈1000 / med≈767 / fast≈500 ms。**PowerPoint 会重算 Choice 的
     `spd` 而毫秒原样保留** —— 往返后看到 spd 变化不代表时长丢了。
   - ④ `<p:transition>` 与 `<p:cSld>` 平级，**看不见任何形状**。
   **顺手修掉两个真 bug**：`build_transition` 写死 `spd="med"` 且静默忽略
   `speed=`（降级世界时长全塌）；附加属性先贴再 `format()`，**只在同时给
   `advTm` 时**写出 `spd="{spd}"` 半成品，而 PowerPoint 照开不误。
   **怎么发现的**：改完先做真机往返（`scripts/probe_roundtrip.ps1`），
   比对"写入 vs 存回"逐项 —— 纯结构检查对此完全无感。

8. **效果层的形态描述 + 选择指导（下一步，用户 2026-09-28 定路线）**
   用户把动效实现拆成三步，并要复核。**复核结论：1、2 已完成，第 3 步要拆成两半。**
   - **① 建立参数表、明确怎么调用** —— ✅ 已做。
     参数表 = `scripts/transition_reference.json`（48 项）+ morph / camera /
     fills / pathRight 等 spec 键；"怎么调用" = `motion.yaml` 的 spec 契约 +
     `scripts/motion.py` 的 `normalize_spec`。**表是量出来的，不是抄的。**
   - **② 确定整体模块运行** —— ✅ 已做。
     `motion.yaml` → `applied.pptx` 真机跑通（12 页 / 31 效果 / 4 切换），
     并过了"写入 vs 存回"往返比对。
   - **③ 细化到具体模块、确定效果与如何选择** —— 拆两半，**3a 已完成**：
     - **3a 形态描述（可量测）** —— ✅ **已做（2026-09-28）**。
       产物 `reference/transition-shapes.md`（48 项形态表 + 6 族分类 + 4 个指标
       + 5 个坑），`com-pitfalls` §44（7 个子节），`symptoms` §44.3/§44.4/§44.7，
       `facts/transitions.json` 4 条规则 + 声明形态层为第三个真相源。
       工具：`build_transition_table.py` 新增 `shapedeck` / `shapedeck2` /
       `shapeanalyze` / `shapes`；新增 `scripts/probe_shapes.ps1`（批量渲染）。
       **核心发现：一个探测 deck 分不开全部效果** —— deck1（网格+移动标记）下
       `fade`/`wipe`/`dissolve`/`split`/`shape`/`randombar`/`box` 的径向剖面
       差值 <0.03；必须加 deck2（满屏噪点）才分得开。
       **因此 `DECK2_SPECS` 那份 12 个的名单是聚类的产物，不是猜的。**
       三个指标各管一族：位移（相位相关）认整页平移、漂移（重心轨迹）认方向揭示、
       径向剖面认中心/均匀。**单一指标会漏掉一整族效果。**
       这一层共修掉 **5 个"能跑但测错"的坑**（探测 deck 每格变色→方向被平均掉、
       编码器噪声被当过渡、首尾帧相位相关、单 deck 分不开、测试样本同构所以
       测试"通过"但无效）。全部记在 `com-pitfalls` §44。判据形态是 FAIL：
       `tests/test_transition_table.py` 第 7–9 组（deck 生成器、方向标定、
       48 项覆盖、facts 一致性）。
     - **3a 正确性证据（用户质疑 → 已闭环）** —— ✅ **已做（2026-09-28）**。
       用户质疑："你如何确保 3a 写入的内容是正确的？" 缺口是真实的：
       数字有标定样本核对，但**把数字翻译成中文这一层没有任何机器验证**，
       而那恰恰是读者唯一会用的东西。
       产物 `scripts/build_shape_evidence.py` → 14 页 deck，左栏我的判定、
       右栏该效果**真实渲染的循环动画**。
       ⚠️ **第一版交的是静态抽帧，被用户当场退回**："压根是静态的，看不到动效"。
       抽帧只能证明"某一刻有什么"，证明不了"怎么动" —— 而形态层判的就是运动。
       改成嵌循环 GIF。随后又发现更隐蔽的坑：**帧数多 ≠ 看得见** ——
       `wipe`/`fade`/`split`/`random` 的 GIF 有 21 帧、无限循环、结构完全合法，
       但帧间差异只有 0.2（deck1 只让一个小圆盘动）＝静止。换 deck2 升到 3.2。
       所以效果与探测 deck 是**配对关系**，写死 `DECK2_ONLY` 名单 +
       `MIN_VISIBLE` 阈值自动拦截"有帧数但看不出运动"的动画。
       全部记在 `com-pitfalls` §46。判据：`test_transition_table.py` 第 10 组
       （含"不许回退成静态抽帧"的负向检查）。
     - **3a 形态修正（用户目视复核）** —— ✅ **已做（2026-09-28）**。
       用户逐页看动画后指出三条，**全部属实，全部已改**：
       1. `uncover` 起点应在**右缘**（我原写"起点在中间"）—— 错因是
          **只抽样了轨迹后半段**；改用基准帧差后能量峰值（帧 70，重心 0.40）
          才是真正中点。
       2. `curtains` 是**左右对开**（我原写 `t→b` 上下）—— 变化重心全程钉在
          **0.50**（对称）即为证。
       3. `random` 表现为"从 y 轴一条线顺时针扫过"——**描述正确**，
          但那一次它**抽中了 wheel**；与 `clock` 的扇区轨迹逐帧相同即为证。
       → 由此发现**第四类指标缺口**：旋转同时躲过位移/漂移/径向三类指标，
       导致 `clock` 被静默归进「中心扩散」族、方向错记 `t→b`。
       **指标盲区不报错，只给错答案** —— 这是本层最值钱的一条教训。
       新增**角度扇区轨迹**指标（`rot_span ≥12` + `rot_mono ≥0.7`，
       第 8b 组三条标定：wheel→True、平移→False、静止→False）。
       族数 6 → **7**（新增「旋转」）。全部记在 `com-pitfalls` §47、
       `transition-shapes` §三之补 + §八。facts 新增 2 条规则。
     - **3b 选择指导（不可量测）** —— ✅ **已做（2026-09-28）**。
       产出 **`reference/transition-choice.md`**（选择层，第四层）。这是**设计判断，
       不是测量结论**，不由渲染帧推出：产出「意图 → 效果」映射，**天然带场景条件**。
       - **结构**：顶部先声明"这一层是判断，不是测量"，并用双向标记把可反驳性写死 ——
         **[实测]**（引用形态层事实，不可反驳，必须可回溯）vs **［判断］**（设计建议，欢迎反驳）。
         §〇 闸门（沿用 §0/§6 全片种类 ≤ 3）；§一 场景字典（维度 A 页间关系 5 种
         = 递进/并列/转折/回归/复位，决定**方向**；维度 B 页型 8 种，沿用 §4，
         决定**强度**）；§二 映射表 6 张（每行含 首选 / dir / 时长 / 备选 / 别用，
         可直接执行）；§三 `dir` 方向语义 + **诚实标注缺口**；§四 禁配清单（三层根据）；
         §五 默认策略（全片 3 种）；§六 与三层的接口。
       - **方向断层（3b 顺带修掉的旧账）**：`push`/`wipe`/`cover`/`uncover` 的默认方向
         **全是 `r→l`**（`wipe` 的 UI 名就是"擦除(左)"）。所以 §5 旧推荐
         "横向对比用 wipe / 章节推进用 push" 在**默认参数下看不出区别** ——
         递进必须显式写 `dir=l→r`，否则被表达成"倒退"。已改写 §5（加 `方向` 列 +
         ⚠️ 默认陷阱 + 指向本层的指针）。
       - **锁定**：`tests/test_transition_table.py` 第 11 组 —— 校验 5 种关系 token、
         8 种页型 token、两种标记都在、**所有反引号切换名都是真实 spec**
         （引不存在的效果照写不报错）、**所有 [实测] 引用的数字与方向在
         `transition-shapes.md` 里逐字存在**（这是"不可反驳"承诺的唯一技术保障）、
         沿用形态层族名（不许自造）、INDEX 能到达本层。facts 新增 2 条规则
         （21 条）：`choice-layer-is-judgement-not-measurement`、
         `choice-layer-cannot-override-shape-layer`，并把
         `transition-choice.md` 登记为 `kind: design-judgement` 来源。
       - **测试当场抓到的真实缺口**：文档讲 `wipe` 时写"扫描族"，与形态层族名
         "方向揭示"不一致 —— 同一件事两套说法，读者无法互校。已统一。
     - **3b 缺口闭合：`dir` 到底会不会镜像** —— ✅ **已做（2026-09-29）**。
       原缺口是"形态层只量了默认形态，换 `dir` 是否镜像无实测"，而 §二/§五
       整张推荐表都建在这个假设上。现在测了。
       - **范围收敛**（用户裁定）：只测选择层实际推荐的 4 个效果 × 2 方向
       - **`curtains` 的对称性**（§三 的另一条声明），共 **10 份 deck**。
         没有全测 48 个 —— 那是重做形态层，不叫补缺口。
       - **结果：4/4 全部镜像成立，`curtains` 确为对称（`dir` 被忽略）**：
         `push` −1.004/+1.009（累积位移）、`wipe` r=−0.994、
         `uncover` r=−0.986、`cover` r=−0.978、`curtains` r=**+1.000** 且 tx=0。
       - **方法论收获（本轮最值钱的一条）**：判镜像**必须两路并行**。
         整页平移（`push`）的每一列变化量相等 → **变化重心根本不动**
         （实测 `dx`=0.011，会被读成"静止"）；而扫描类（`wipe`）的总质量位移
         ≈0，累积位移看不见它。**单一指标必有一族盲区，且盲区是静默的**
         —— `push` 第一次就被判成 `ambiguous`，原因不是它怪，是仪器选错了。
         这条与形态层 §八（旋转躲过前三类指标）是同一个教训的第二次出现。
       - **顺带修掉一个真 bug**：分析窗口原为 `bnd-2s .. bnd+3s`，在 148 帧的
         渲染上会**吞掉整段片子**（`bnd+3*fps=150 > 148` 且 `bnd-2*fps=0`），
         把第 1 页的停留也算进来 → `curtains` 被读成 `t->b`、活跃窗口只剩 1 帧。
         改为 `bnd ± 1.5s`（过渡只有 0.5–1.0s，窗口必须紧）。
         **`cmd_shapes` 与 `cmd_dirmirror` 同源同修**，否则两层口径会漂。
       - 新增命令：`dirdeck`（生成方向变体 deck）、`dirmirror`（配对判镜像）；
         `_profile_metrics` 新增返回逐帧重心轨迹 `cx_trace`/`cy_trace`。
       - **测试**：第 12 组 —— 用**合成帧**先把判据标定住（左→右 vs 右→左
         必须判 mirror、自身对自身必须不判 mirror、静止必须判 no-motion），
         再校验 DIR_PROBE 名单与"dir 真的写进了 XML"。
     - **3b 缺口闭合（第二轮）：把"第一行其余成员"全测掉** —— ✅ **已做（2026-09-29）**。
       上一轮只测了 4 个推荐效果，其余成员是推断。本轮把**两个方向性族全部
       19 个效果**逐一测掉（各 2 方向，共 **42 份 deck**）。
       - **范围纪律（写进测试的判据，不是拍份数）**：只测两族成员，
         减去 §四 禁配清单，中心扩散/旋转/淡变/随机不测（没有方向轴可翻）。
       - **结果：19 个干净地分成两堆**（`dirprobe2`）：
         - **会镜像 8 个**：`push` `pan` `switch` `wipe` `cover` `uncover`
           `reveal` `glitter`。
         - **被忽略 11 个**：`page_curl` `peel_off` `shape` `wind` `curtains`
           `airplane` `crush` `drape` `fall_over` `prestige` `origami`
           —— 两个方向**逐帧像素差 = 0.000**。
       - **两路独立判据 19/19 完全一致**：逐帧像素差（变没变）与 `band_travel`
         反号（怎么变）判出的集合**逐一相同**；分离度极大，`max_frame_diff`
         取值 0.000 → 4.10 → 23.9+，**没有灰色地带**。
       - **⭐ 纠正了选择层两处早先的错**（都是一次全量实测量出来的）：
         `curtains` 不是"左右对称"而是**竖向条带对开**；
         `origami` 默认是**水平 `l→r`**（早先误归 `t→b`）。
         另外把 `split` 从"能定向"表里**移出**（实测是中心扩散，`dir` 无效）。
       - **⭐ 新发现：`dir` 加错地方会让 PowerPoint 直接打不开**（不是"被忽略"）。
         `box`（元素其实是 `<p:zoom>`）与 `comb`（`<p:comb>`）不接受 `dir`。
         **单变量隔离**证实：不写能开、写了开不了。详见 com-pitfalls **§49**。
         这条对技能包是**硬约束**，写配置不能"所有效果都加 dir"。
       - **"被忽略" ≠ "没有方向"**：`page_curl`/`peel_off`/`shape`/`wind`
         方向都很明确，只是**写死、`dir` 拧不动** —— 要反向只能换效果。
       - 归档：`tests/fixtures/dir/dirmirror_full.json`（19 条全量）、
         `tests/fixtures/dir/dir_compare_full.png`（目视对照）。
       - **剩余未知**：只有禁配效果与"无方向轴"的四族未测（本就不该测）。
         两个方向性族已**全测完**，选择层的方向建议现在**全部有实测支撑**。
   - **顺序不可颠倒**：先有形态（"wipe(dir=r) 是右进左出"）才谈选择
     （"向右推进适合表达时间前进"）。反过来会退化成效果清单 + 形容词 ——
     正是 `recipe-library.md`「按原理不按效果」要防的那件事。
   - **范围界定**（回答"机制层是否覆盖了 48 项全部"）：否。机制层是 47 个普通切换
     **共用**的 4 条底层机制，**一个具体效果的"长相"都没描述**；
     48 项之间的取舍、`<p:timing>` × 切换的配合，都还没有。
     - **3a 遗留之一「属性取值全集」** —— ✅ **已做（2026-09-29）**。
       名字层原先每个效果只记一个子元素（如 `clock -> <p:wheel spokes="1"/>`），
       但枚举里 149 个值有一大半是**同一元素换属性值**，此前没人量过它们干什么。
       - **做法**：新增 `ATTR_PROBE`（26 份 deck）+ `attrdeck` / `attrdiff`。
         判据用**第三路仪器：逐帧像素差** —— `spokes`/`pattern`/`isContent`
         本来就不是方向，用轴判据会**问错问题**。
       - **⭐ 最重要的发现：非法属性值静默退回默认**（com-pitfalls **§50**）。
         `clock` 的 `spokes="6"`（enum 里没有）与不写、与 `spokes="4"`
         **逐帧像素差 0.000** —— 文件照开照渲染，你写的值被悄悄丢掉。
         这是第三种"写了不生效"（前两种：被忽略 / 打不开）。**只验"能打开"不够。**
       - **clock 默认是 4 根**，不是名字层原先记的 1 根（`spokes="1"` 只是
         **一个取值**，不是默认）。
       - **多个效果换属性值 = 换界面项**：`split` 的 `orient`×`dir` 四组合全不同、
         `comb` 的 `dir`（horz=默认 / vert 差 64.0）、`p14:prism` 的
         `isContent`×`isInverted` 四组合（像素差 62–105）全不同、
         `glitter`/`shred` 的 `pattern` 换粒子形状但**不改方向**。
       - **⭐ `invX` 才是 `p15:prstTrans` 系列的反向属性**（wind 82.5 /
         peel_off 42.4 / fall_over 12.4），配合 §九 的"`dir` 对它们无效"，
         得到完整规则：**同一个"反向"意图，core 元素用 `dir`，`prstTrans` 用 `invX`。**
       - **纠正名字层一处含糊**：原写"prism 靠 isContent/isInverted 区分三个界面项"，
         **没点明是哪两个位**。实测「轨道」= 两个都置位，而 `isInverted` 单独置位
         是**界面上没有的第四形态**。已写准。
       - 产物：`transitions.md` **§5.5 属性取值全集**、`transition-shapes.md`
         **§十**、`com-pitfalls` **§50**、facts +5 条（32 条）、测试**第 13 组**。
         归档 `tests/fixtures/dir/attrdiff_full.json` + `attr_compare_full.png`。
     - **仍剩**：48 项之间的取舍（属选择层）、`<p:timing>` × 切换的配合。

**验证文化（请保持）**
- 改 OOXML 后**必须**过 PowerPoint 真开一次 —— 不是可选步骤
- 改分析器/判据后**必须**用已知答案样本对表
  （`make_calibration.ps1` 生成标定视频；`tests/` 里有回归样本）
- 提交前跑：`selftest` + `test_morph` + `test_camera` + `test_dual_photo` + `smoke`
  + `privacy_audit` + **`verify_docs`** + **`test_install_manifest`**
  + **`test_transition_table`**
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
