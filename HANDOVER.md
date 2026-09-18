# 交接文档

> 给接手这份工作的人（或下一个会话的 agent）。
> 先读本文，再读 `SKILL.md`（操作手册）与 `references/review-checklist.md`（交付复核）。

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
| 方向性擦除 `dir:` | ✅ | 上游 `selftest` |
| 艺术效果（虚化 / 亮度，走 COM） | ✅ | 与 template1 的 XML 逐字节同构 |
| 内嵌 MP4 / WAV | ✅ | `motion.ps1` 媒体层 |
| 从视频反推动效节奏 | ✅ | `tests/calib` 已知答案标定 |
| 动效预览（HTML 重放） | ✅ | `motion.py player` |
| 交付复核（含美感判据） | ✅ | `review_assist.py` + `review-checklist.md` |

**测试总览**：`selftest 63` + `test_morph 20` + `test_camera 30` + `smoke` 三场景，全绿。

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
④ 设计层  references/               规范与配方（不产生代码）
          motion-design-spec.md     该怎么动
          review-checklist.md       怎么验（含美感判据）
          morph-and-3d-recipes.md   三项技术配方
          design-system/            10 套静态版面设计系统
```

**为什么要分层**：每一层能证明的东西不同。①只证"XML 合法且几何没动"，
③才证"PowerPoint 采纳了"。把①的结论当③用，就会犯第 6 节那些错。

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

references/
  authoring-rules.md          13 条硬约束（每条对应一次真实翻车）
  com-pitfalls.md             COM / OOXML 坑，含"静默失败"四类
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
  smoke.py                    三场景端到端（pptd 导出 / PowerPoint 原生 / 真开）
  privacy_audit.py            隐私与泄漏审计（扫描内容 + git 历史）
  fixtures/                   回归样本
  motion.yaml, fx-pro.pptx    回归 spec 与样本
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

## 6. ⚠️ 本项目最重要的部分：三次真实误判

**这三条比任何功能说明都重要。** 它们形态完全相同 ——
**拿单一自动化判据当结论，没做交叉验证** —— 而症状都是"自信地给出错误答案"。

| # | 当时结论 | 实际 | 根因 |
| --- | --- | --- | --- |
| 1 | "本机不支持 Morph" | **支持**，`EntryEffect=0xF72` | 元素名写成 `p:morph`，正确是 `p159:morph`；查 COM 枚举自然查不到 |
| 2 | "template2 两页几何完全相同、不产生补间" | 矩形组 left 从 **−180.8** 变到铺开 | 解析脚本**没处理 `<p:grpSp>`**，读到组变换而不是子形状 `<a:off>` |
| 3 | "原件里人物不在该位置"（模板匹配 0.401） | 人物**就在**那里 | 模板匹配在低对比/缩放差异下不可靠 |

**第 1 条的完整教训**：查不到某个能力时，先怀疑自己写错 ——
"对象模型里没有"**不等于**"不能注入"。注入是写 XML，与属性模型是两条路。

**第 2 条的完整教训**：断言"没有差异"之前，必须先证明**解析器看得见差异**。
把已知有差异的样本喂进去，看它是否报不同。

**因此 `references/review-checklist.md` §5 的三条规则不是形式主义**：

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
1. **美感判据的执行化** —— `review_assist.py` 目前只覆盖素材卫生 + 少量启发式。
   `motion-design-spec` 的 §6 密度上限、§4 页型剧本可以做成**逐页核对表**，
   让每页生成时就带上自检。
2. **多 API 协作** —— 素材生成（尤其抠图/插画）接专门的图像生成 AI，
   替代 `grabCut` 这类兜底方案。
3. **morph / 3D 之外的 spec 扩展** —— 目前 `cameras:` 与 `morph` 已入 spec，
   单页内动画（形状切割）仍是手写注入。

**验证文化（请保持）**
- 改 OOXML 后**必须**过 PowerPoint 真开一次 —— 不是可选步骤
- 改分析器/判据后**必须**用已知答案样本对表
  （`make_calibration.ps1` 生成标定视频；`tests/` 里有回归样本）
- 提交前跑：`selftest` + `test_morph` + `test_camera` + `smoke` + `privacy_audit`

---

## 9. 来源与许可

- 本 skill 基于上游 [`hfz14263/ppt-motion-skill`](https://github.com/hfz14263/ppt-motion-skill)
  的成熟版本演进（该线已有 63 断言自测、13 条 authoring-rules、62 相机实测表、
  `player.py`）。本仓库在其基线上叠加：动效设计规范、设计系统、视频分析、
  交付复核清单、camera/morph spec 支持。
- `references/design-system/` 的 10 套设计系统来自
  [`open-kimi-ppt-skill`](https://github.com/acnlie/open-kimi-ppt-skill)（**MIT**，
  Copyright (c) 2026 Binaryify Zhuang），逐字保留并附 `LICENSE-upstream.txt`。
- 本项目 MIT。

**注**：仓库 git 历史曾做过一次改写（统一为
`hfz14263 <hfz14263@users.noreply.github.com>`），因此全部提交署同一身份。
改写原因见第 6 节同一套"先自查"原则 —— 早期提交信息里带了机器路径，
且作者邮箱是个人邮箱。改写前的完整备份在仓库外（`*-BACKUP-pre-rewrite.bundle`），
确认无需回滚后可删除。
