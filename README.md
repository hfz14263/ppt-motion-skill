# ppt-office-motion

> 让你的 PPT 动起来。

给**任意** `.pptx` 叠加动画、切换与内嵌音视频,**而完全不改动它的版面几何**。

> 上游刚把版面调好,一旦用 PowerPoint 打开→编辑→保存,PowerPoint 会跑自己的排版引擎
> (字体解析、auto-fit 重排),上游刚通过的重叠/溢出结论立刻作废。
> 这个 skill 的全部设计就是为了绕开这一步。

---

## 核心承诺

注入只碰三样东西:

```
ppt/slides/slideN.xml      ← 只加 <p:transition> 和 <p:timing>
[Content_Types].xml        ← 只加媒体 MIME 声明
ppt/media/*                ← 只加内嵌媒体
```

形状的 `<p:spTree>` 及其 `<a:xfrm>` **一个字节都不动**。所以:

- 上游门禁(`ppt_verify` 之类)的通过结论**可以直接继承**,不必重跑
- 注入前后可以算**几何指纹**做机器自证,输出 `geometry: UNCHANGED`

---

## 快速开始

```bash
pip install -r requirements.txt

# 0) 看清这个 deck 有什么
python scripts/motion.py inspect --pptx deck.pptx

# 1) 写 spec(见 examples/motion.example.yaml),然后注入 + 几何自证
python scripts/motion.py apply --pptx in.pptx --spec m.yaml --out out.pptx --assert-geometry

# 2) 结构自证:well-formed / cTn id 唯一递增 / 几何不变
python scripts/verify_motion.py --pptx out.pptx --source in.pptx

# 3) 覆盖率审计:抓"没有任何效果"的形状(它们在放映开始就可见,是静默错位的主因)
python scripts/motion.py check --pptx out.pptx --spec m.yaml

# 4) 真渲染 + PDF + PowerPoint 往返普查(本机需要装 PowerPoint)
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/motion.ps1 `
           -Pptx out.pptx -Spec m.yaml -OutDir review -ExportPdf -Strict

# 5) 动效预览:把 spec 时序在浏览器里重放(唯一能"看见动效"的手段)
python scripts/motion.py player --pptx out.pptx --spec m.yaml --outdir preview
```

**第 3 步和第 4 步不要跳。** 第 3 步抓的是"该动的没动",第 4 步抓的是
"PowerPoint 保存后把动画吃了"——这两类问题**前面所有的结构校验都发现不了**。

---

## motion spec 长什么样

```yaml
version: 1
slides:
  - page: 1
    transition: {type: fade, duration: 0.8}
    effects:
      - {target: ground, effect: wipe, duration: 0.6, dir: down}
      - {target: title,  effect: wipe, duration: 0.8, dir: left}
      - {target: badge,  effect: spin, duration: 1.0, repeat: 2}
    media:
      - {src: media/clip.mp4, bounds: [100, 90, 520, 293], loop: false}
```

| 字段 | 说明 |
| --- | --- |
| `target` | elementId / 形状 name / 数字 id / `[多个]` |
| `effect` | 137 个别名,`motion.py catalog` 列全(入场 / 强调 / 动作路径) |
| `dir` | `down left right up upleft upright downleft downright` — **方向性擦除**,只对有 filter 的效果有效 |
| `trigger` | `click` \| `with` \| `after` |
| `delay` `duration` `repeat` `autoReverse` `smooth` | 时序与节奏 |
| `media.*` | src / bounds / loop / rewind / mute / volume / autoplay / elementId |

### 方向性擦除(`dir`)

`wipe` 默认是 `wipe(down)`。加 `dir: left` 才能让一条**横向数据线被"画"出来**,
而不是掉下来。这是唯一能把"方向"变成设计决定的手段。

⚠️ **但方向装反了没有任何结构闸能发现**:八个方向的
`presetID/presetClass/presetSubtype` 完全一致,只差一个 `filter` 字符串。
`--assert-geometry`、`verify_motion`、往返普查**全部会通过**。
交付前必须人眼看一次。详见 `references/authoring-rules.md` §H/§I。

---

## 目录结构

```
SKILL.md                      技能主文档:分层职责、S0-S7 流程、九条坑、对接契约
references/
  architecture.md             知识构成与处理流程总览(OOXML / COM / 渲染行为)
  authoring-rules.md          13 条硬约束,每条对应一次真实翻车 + 三个人眼清单
  com-pitfalls.md             18 条 COM / OOXML 坑,含"静默失败"四类
  template-patterns.md        给"一个动画都没有"的现成模板从零设计动效
  mso-primitives.md           PowerPoint 原生效果与切换的原始语义
  ppt-studio-integration.md   与 dsh-ppt-studio(deck.yaml + elementId)的对接契约
scripts/
  motion.py                   OOXML 引擎:apply / inspect / preview / catalog / player / check
  motion_catalog.json         137 个效果的逐字模板(从真实 PowerPoint 输出提取)
  player.py                   逐形状图层 + 自包含 HTML 播放器
  motion.ps1                  Office COM 层:媒体内嵌 / 真渲染 / PDF / 往返普查
  verify_motion.py            结构 + 几何自证
  check_coverage.py           覆盖率审计
  selftest.py                 63 项回归测试
examples/motion.example.yaml
install.ps1                   安装到 %USERPROFILE%\.dsh\skills
```

---

## 为什么需要这么多道闸

PowerPoint 有一类失败方式:**文件结构完全合法、打开不报错、不弹修复,动画就是没了。**
已知四种:

1. `<p:transition>` 位置违反 `CT_Slide` 的 sequence → 切换静默消失
2. `<p:cTn id>` 非按文档顺序递增 → 效果静默消失
3. 同一形状"入场 + 强调" → 强调被丢掉(本机固有限制)
4. 同一形状挂多个效果(逐段揭示)→ 往返时折叠成一行,多出来的丢掉

所以有 `motion.ps1 -Strict`:把"保存后动画变少"从警告变成**失败退出**。

### 门禁通过 ≠ 动效正确

这个工具**证明不了**四件事,而这四件才是实际翻车的地方:

1. **该动的都动了** — 漏掉的形状在放映一开始就可见,不报任何错
2. **顺序对** — 标题在内容前、结论在论据后
3. **单位对** — EMU / pt / inch 混用静默错位
4. **方向对** — 见上面的 ⚠️

---

## 环境要求

- Python 3.7+(`python-pptx` / `lxml` / `PyYAML` / `Pillow` / `numpy`)
- **Windows + PowerPoint**(`motion.ps1` 那一层需要;真渲染、媒体内嵌、往返普查都依赖它)
- Python 侧的 `apply` / `verify` / `check` / `catalog` **不需要** PowerPoint

本机(Windows + PowerPoint 16.0)实测的两个限制,写在这里免得别人重走:

- **导不出 MP4**:`Presentation.CreateVideo` 对所有参数组合都失败
  (Quality 0 返回成功但不产文件;1/2 报 `E_INVALIDARG`)。所以用 `player` 做动效预览。
- **动画的中间态不可观测**:`Slide.Export` / `Shape.Export` 只渲染终态,
  `SlideShowView.Export` 在部分 build 上根本不存在。详见 `references/com-pitfalls.md` §17。

---

## 安装到 DSH

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
# 已存在时需显式覆盖:
powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -Force
```

装到 `%USERPROFILE%\.dsh\skills\dsh-ppt-office-motion`,然后重启 dsh web 让技能目录刷新。

---

## 自测

```bash
python scripts/selftest.py      # 63 passed, 0 failed
```

覆盖:transition/timing 顺序、cTn id 单调性、几何不变、预览剥离、
`player` 的图片预解码、方向性擦除到达 `filter=`、以及给无方向效果写 `dir` 会报错。

---

## 许可

未附许可文件。若要公开分发,请先补一个。
