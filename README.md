# ppt-office-motion

> 让你的 PPT 动起来。

给**任意** `.pptx` 叠加动画、切换与内嵌音视频,**而完全不改动它的版面几何**。

上游刚把版面调好,一旦用 PowerPoint 打开→编辑→保存,PowerPoint 会跑自己的排版引擎
(字体解析、auto-fit 重排),上游刚通过的重叠/溢出结论立刻作废。
这个工具的全部设计就是为了绕开这一步。

---

## 核心承诺

注入只碰三样东西:

```
ppt/slides/slideN.xml      ← 只加 <p:transition> 和 <p:timing>
[Content_Types].xml        ← 只加媒体 MIME 声明
ppt/media/*                ← 只加内嵌媒体
```

形状的 `<p:spTree>` 及其 `<a:xfrm>` **一个字节都不动**。所以:

- 上游门禁的通过结论**可以直接继承**,不必重跑
- 注入前后可以算**几何指纹**做机器自证,输出 `geometry: UNCHANGED`

## 安装

```bash
pip install -r requirements.txt
```

Windows 上还可以装成技能:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Force  # 覆盖已有
```

**只用 Python 侧也够用**:`apply` / `verify` / `check` / `catalog` / `player` 都不需要
PowerPoint。只有真渲染、媒体内嵌、往返普查需要 Windows + PowerPoint。

## 用法

```bash
python scripts/motion.py inspect --pptx deck.pptx          # 看清有哪些形状和 id
python scripts/motion.py apply --pptx in.pptx --spec m.yaml \
       --out out.pptx --assert-geometry                    # 注入 + 几何自证
python scripts/verify_motion.py --pptx out.pptx --source in.pptx
python scripts/motion.py check --pptx out.pptx --spec m.yaml
python scripts/motion.py player --pptx out.pptx --spec m.yaml --outdir preview
```

spec 的完整字段见 [`examples/motion.example.yaml`](examples/motion.example.yaml)。

### ⚠️ 三道结构闸通过,不等于动效正确

这个工具**证明不了**四件事,而这四件才是实际翻车的地方:

1. **该动的都动了** — 漏掉的形状在放映开始就可见,**不报任何错**
2. **顺序对** — 标题在内容前、结论在论据后
3. **单位对** — EMU / pt / inch 混用会静默错位
4. **方向对** — 方向性擦除(`dir:`)的八个取值预设三元组完全一致,
   装反了所有结构闸**全部通过**

所以交付前**必须**跑覆盖率审计(`check`)和往返普查(`motion.ps1 -Strict`),
并且**人眼看一遍方向和顺序**。清单见
[`references/authoring-rules.md`](references/authoring-rules.md) §H。

## 文档

| 文件 | 讲什么 |
| --- | --- |
| [`SKILL.md`](SKILL.md) | **操作手册**:分层职责、S0–S7 流程、spec 全文、命令、九条坑 |
| [`references/architecture.md`](references/architecture.md) | 为什么这样设计:OOXML / COM / 渲染行为的知识构成 |
| [`references/authoring-rules.md`](references/authoring-rules.md) | 13 条硬约束,每条对应一次真实翻车 + 人眼清单 |
| [`references/com-pitfalls.md`](references/com-pitfalls.md) | 28 条 COM / OOXML 坑,含"静默失败"四类 |
| [`references/morph-and-3d-recipes.md`](references/morph-and-3d-recipes.md) | **Morph(平滑)与 3D 相机的注入配方**:元素写法、命名空间、跨页匹配规则、实测对照、两个测量陷阱 |
| [`references/template-patterns.md`](references/template-patterns.md) | 给"一个动画都没有"的现成模板从零设计动效 |
| [`references/mso-primitives.md`](references/mso-primitives.md) | PowerPoint 原生效果与切换的原始语义 |
| [`references/camera-reference.md`](references/camera-reference.md) | **3D 相机参数 → 实测结果**:62 个预设、lat/lon、fov、zoom 的实测对照表 |
| [`references/ppt-studio-integration.md`](references/ppt-studio-integration.md) | 与 `dsh-ppt-studio`(`deck.yaml` + `elementId`)的对接契约 |
| [`references/motion-design-spec.md`](references/motion-design-spec.md) | **动效怎么设计**:动效闸门、时长阶梯 T0–T3、效果语义映射、7 类页型编排剧本、密度上限、自检清单 |
| [`references/design-system/README.md`](references/design-system/README.md) | **静态版面**设计系统:内置 10 套(咨询/财务/汇报/推广/学术)+ 上游 34 套按需抓取 |
| [`references/video-analysis-limits.md`](references/video-analysis-limits.md) | 从视频量动效的**实测能力边界**:起点 ±0.10s、时长 +0.02~+0.16s、哪些测不出 |
| [`references/review-checklist.md`](references/review-checklist.md) | **交付复核清单**:判据可信度分级(哪条判据已知会骗人)、结构/素材卫生/版面/动效/一致性检查项、交叉验证规则 |

## 找参考素材时先跑一下 `inspect_pptx.py`

`scripts/inspect_pptx.py` 回答一个问题:**这个 pptx 到底有没有动画?**

因为标签会骗人:Canva 导出的 pptx 是**静态**的(它的动效是渲染层功能),
模板站的 "animated" 分类通常只表示**有页面切换**,而一个文件夹名叫
"Carousel Animation" 的仓库实测 **0 个效果**。视频更没法读——看得到动,读不出任何时间参数。

```bash
python scripts/inspect_pptx.py deck.pptx          # 单文件
python scripts/inspect_pptx.py refs/ -r           # 递归扫一个目录
python scripts/inspect_pptx.py *.pptx --json      # 机器可读
```

它报告:效果总数与分类(入场/强调/路径)、时长与延迟的分布、触发方式、
逐页效果数,并给出结论。判据是**带 `<p:timing>` 的页数占比 ≥ 1/3** ——
只有切换的文件是页面装饰,有 timing 的才是设计。

## 自测

```bash
python scripts/selftest.py      # 63 passed, 0 failed
```

## 交付前复核

```bash
python scripts/review_assist.py --pptx out.pptx --source in.pptx
```

它区分两类结论：**结构项**是机器可判、可信的；**素材卫生与版面项**是启发式信号，
**必须人眼确认**。脚本还会列出**它判不了的**（视觉主角、留白、节奏、方向、是否好看…）。

实测它能自动抓出真实事故：`material/template1` 的旋转图层是全屏截图，含 PowerPoint
界面（69% 近白像素、547 行近白），以及 `image1.png` 被放大 1.47 倍显示。

判据可信度分级与三条交叉验证规则见
[`references/review-checklist.md`](references/review-checklist.md)。

## 本机能力探测

两项能力**随 Office 构建变化，不要套用别人的结论**:

```powershell
# CreateVideo 能否导出 MP4(上游不可用、另一台机器可用;quality 0 两边都失败)
powershell -NoProfile -File scripts/probe_createvideo.ps1
```

```bash
# 从视频量动效节奏(拿不到 pptx 时的兜底;先跑标定对表)
python scripts/analyze_video.py --video clip.mp4 --json
python scripts/analyze_video.py --video tests/calib/calib.mp4 --json   # 已知答案样本
```

## 许可

[MIT](LICENSE)。
