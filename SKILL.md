---
name: ppt-office-motion
description: 给任意 .pptx 叠加动画、切换与内嵌音视频，而完全不动版面几何。用声明式 motion spec（按 elementId 锚定）驱动 OOXML 增量注入 + Office COM 真渲染复核；动画/媒体注入后布局帧字节不变，可直接继承上游 dsh-ppt-studio 的 ppt_verify 门禁结论。
whenToUse: 需要给已定稿的 PPT 动态化时——加入场/强调/动作路径动画、页面切换、内嵌 MP4/WAV；或需要验证"动画没有把版面搞坏"；或要与 dsh-ppt-studio 的 deck.yaml 工程配合，把 elementId 变成动画目标。
---

# PPT Office Motion

把一个已经过版面校验的 `.pptx` **动起来**，而不破坏它的版面。

## 1. 为什么不是"再用 PowerPoint 加工一遍"

用 PowerPoint 打开→编辑→保存会触发 PowerPoint 自己的排版引擎（字体解析、auto-fit 重排），
上游刚通过的重叠/溢出/门禁结论就作废了。本 skill 的分工是：

| 层 | 归属 | 职责 |
| --- | --- | --- |
| 版面 | 上游（如 dsh-ppt-studio 的 PPTD + `ppt_verify`） | 元素不打架、不溢出 |
| 动画/切换 | 本 skill 的 `scripts/motion.py` | **纯 OOXML 增量注入**，几何零改写 |
| 媒体 | 本 skill 的 `scripts/motion.ps1` | `AddMediaObject2` 内嵌 MP4/WAV |
| 复核 | `verify_motion.py` + `motion.ps1` | 布局帧 SHA256 自证 + Office 真渲染 |

因为注入只碰 `ppt/slides/slideN.xml` 的 `<p:timing>` / `<p:transition>` 和 `[Content_Types].xml`，
spTree 的形状及其 `a:xfrm` 原封不动，所以 **`ppt_verify` 的通过结论可以直接继承**，不必重跑。

## 2. 标准作业流程

```
S0  确认上游已有 .pptx（已过 verify / 或任意现成 pptx）
S1  motion.py inspect --pptx deck.pptx        # 看清 elementId 与形状 id 的对应
S2  写 motion spec（按 elementId 声明效果；见 §3）
S3  motion.py apply --pptx in.pptx --spec m.yaml --out out.pptx --assert-geometry
S4  verify_motion.py --pptx out.pptx --source in.pptx     # 结构 + 几何自证
S5  motion.ps1 -Pptx out.pptx -Spec m.yaml -OutDir review -ExportPdf -Strict
                                              # 真渲染 + 媒体 + PDF + 往返普查
S6  motion.py player --pptx out.pptx --spec m.yaml --outdir preview
                                              # 动效预览，浏览器打开 preview.html
S7  read_image 看 review/render/*.png         # 视觉审阅（版面）
```

**S3 必须带 `--assert-geometry`**，S4 必须比对 `--source`。这两步是本 skill 的核心承诺，
输出 `geometry: UNCHANGED` 才算通过；一旦报 CHANGED，说明注入了不该注入的东西，停下排查。

**S5 建议带 `-Strict`**：只有它会把"PowerPoint 保存后动画变少"变成失败退出。不带时
`LOSS` 只作为警告打印，退出码仍是 0。

**S6 是唯一能看到"动"的一步**，S7 的静态图看不出动画（见 §4 的 `player`）。

### ⚠️ 门禁通过 ≠ 动效正确

S3 的 `--assert-geometry`、S4 的 `verify_motion`、S5 的往返普查，
只证明 **OOXML 合法 + 几何没动 + 效果没被 PowerPoint 吃掉**。
它们**证明不了**下面这三件事——而这三件正是实际翻车的地方：

1. **该动的都动了**（漏掉的形状在放映开始就可见，不会报错）
2. **顺序对**（标题在内容前、结论在论据后）
3. **单位对**（EMU / pt / inch 混用会静默错位）

**动手前先读 `references/authoring-rules.md`**（13 条硬约束，每条对应一次真实翻车）。
它给出三个额外校验器：覆盖率、目标有效性、版面（溢出/越界/压页脚）。

特别是这两条最贵：

- **A1**：spec 的形状 id **禁止手写**，必须由实测几何生成。deck 一改版式 id 就重排，
  手写的会静默指向别的形状（曾出现"正好差一位"，导致整张卡没动画）。
- **G**：改了 deck 就必须**重新生成 spec**。"觉得 spec 没动"就复用旧的，是上面那个 bug 的成因。

## 3. motion spec

```yaml
version: 1
slides:
  - page: 1                      # 1-based 页号；也支持 slide: 或 index:
    transition: {type: fade, duration: 0.8}
    effects:
      - {target: band,  effect: fade, duration: 1.0, trigger: with}
      - {target: title, effect: fly,  duration: 0.7, trigger: after, delay: 0.2}
      - {target: badge, effect: spin, duration: 1.0, trigger: after, repeat: 2}
      - {target: glow,  effect: pathSineWave, duration: 3.0, trigger: after, autoReverse: true, smooth: 0.5}
    media:
      - {src: media/clip.mp4, bounds: [100, 90, 520, 293], loop: false, rewind: true,
         volume: 0.8, autoplay: true, elementId: VIDEO}
```

字段：

| 字段 | 取值 | 说明 |
| --- | --- | --- |
| `target` | elementId / 形状 name / 数字 id / `[多个]` | dsh-ppt-studio 导出里 elementId 就是 cNvPr/@name |
| `effect` | 137 个别名，`motion.py catalog` 列全 | 入场 / 强调 / 动作路径 |
| `dir` | `down` `left` `right` `up` `upleft` `upright` `downleft` `downright` | **方向性擦除**，只对有 filter 的效果有效（见 §3.1） |
| `trigger` | `click` \| `with` \| `after` | 默认 `after` |
| `delay` | 秒 | 相对延时 |
| `duration` | 秒 | 写到叶子行为 cTn（毫秒） |
| `repeat` | 次数 | `repeatCount = 次数×1000` |
| `autoReverse` | true/false | 往复播放 |
| `smooth` | 0–1 | accel/decel |
| `transition.type` | `fade` `push` `wipe` `cover` `split` `zoom` `dissolve` `strips` `pull` `randombar` `none` | 见 `references/mso-primitives.md` |
| `media.*` | src/bounds/loop/rewind/mute/volume/autoplay/elementId | 媒体必须走 COM 层 |

### 3.0 3D 相机与旋转（本 skill **尚未**纳入 spec）

`<a:scene3d>` 里的相机**不在 spec 字段里**。要做立体效果只能手写注入，
实测对照表见 [`references/camera-reference.md`](references/camera-reference.md)，
测量脚本 `scripts/build_camera_table.py`。

**三条最容易翻车的，先看这三条：**

1. **只有 `prst="perspective*"` / `"legacyPerspective*"` 会产生透视。**
   另外 47 个（`oblique*` / `isometric*` / `orthographicFront`）
   **全是平行投影、永远没有灭点** —— 平面不可能"躺下去"，
   无论角度怎么调。
2. **`a:rot` 的 `lat` / `lon` / `rev` 必须落在 `0..21599999`。**
   规范写的上界 `21600000` **会让 PowerPoint 报整个文件损坏（`0x80070570`）**，
   负数同样。所以 `-45°` 要写成 `18900000`。
3. **判断"躺下没有"要量收敛比（远边宽 / 近边宽）：`1.000` = 平行，明显小于 1 才是透视。**
   **不要看"像不像平行四边形"** —— 平行投影的矩形在任何角度都像，
   这个判据零区分力。

`fov`（0~180°）是**透视强度的连续旋钮**，越大透视越强；
实测它对压扁度几乎无影响，所以**透视强度和俯角可以分别调**。

### 3.1 方向性擦除（`dir`）

`wipe` 默认是 `wipe(down)`。给带 filter 的效果加 `dir:` 改方向，这是**唯一**能让
"一条线被画出来"而不是"掉下来"的手段：

```yaml
effects:
  - {target: 12, effect: wipe, duration: 1.6, dir: left}   # 横向数据线，从左画到右
  - {target: 30, effect: wipe, duration: 0.9, delay: 0.3, dir: down}
```

- **只有有方向的 filter 接受它**：`wipe / barn / checkerboard / circle / diamond /
  plus / wheel / wedge / strips`。给 `fade`、`fly`、动作路径写 `dir:` 会**直接报错退出**，
  不静默忽略——静默忽略等于交付了另一个动画。
- 方向是**逐效果覆盖**，不是独立 alias：`wipe(left)` 没有自己的 presetID，
  方向只存在于 `<p:animEffect filter="...">` 的取值里。
- ⚠️ **`dir` 装反了没有任何结构闸能发现**。八个方向的
  `presetID/presetClass/presetSubtype` 完全一致，只差 `filter` 字符串，所以
  `--assert-geometry`、`verify_motion`、往返普查**全部通过**。本机也无法自动校准：
  COM 导出看不到动画中间态、截屏拿不到画面（三条死路见
  `references/authoring-rules.md` §I3）。**交付前必须人眼看一次方向**（同文件 §H）。
- `player` 会按方向渲染起始裁剪（`player.py` 的 `WIPE_CLIP`）。这一步是必须的：如果
  预览把所有擦除都画成"从左边长出来"，那它验证的就是错的东西。

## 4. 命令

```bash
# 看清 deck：哪些形状、什么 id、已有哪些动画
python scripts/motion.py inspect --pptx deck.pptx [--json]

# 列出可用效果（P=带动作路径 F=带滤镜 V=带可见性设置）
python scripts/motion.py catalog [--kind entrance|emphasis|path]

# 注入 + 几何自证
python scripts/motion.py apply --pptx in.pptx --spec m.yaml --out out.pptx \
       --assert-geometry --report report.json

# 结构自证（well-formed / spid 有效 / cTn id 唯一 / 几何不变）
python scripts/verify_motion.py --pptx out.pptx --source in.pptx

# 覆盖率审计：报告"没有任何效果"的形状（它们在放映开始就可见，是静默错位的主因）
python scripts/motion.py check --pptx out.pptx --spec m.yaml

# 生成去动画的静态副本（导出确定性截图/PDF 用）
# 同时清掉 timing 与 transition，并回收空掉的 mc:AlternateContent
python scripts/motion.py preview --pptx animated.pptx --out static.pptx

# 动效预览：把 spec 的时序在浏览器里重放（不需要视频）
# 产出 <outdir>.html + <outdir>/stage/*.png（每形状一层，带透明通道）
python scripts/motion.py player --pptx animated.pptx --spec m.yaml --outdir preview

# Office COM 层：媒体内嵌 + 真渲染 + PDF（默认只读打开，绝不回写输入）
powershell -NoProfile -ExecutionPolicy Bypass -File scripts/motion.ps1 `
           -Pptx out.pptx -Spec m.yaml -OutDir review -ExportPdf
```

**`player` 是唯一能"看见动效"的手段。** 静态 PNG/PDF 看不出动画，而本机
`Presentation.CreateVideo` 对所有参数组合都失败（导不出 MP4，见
`references/com-pitfalls.md` §16），所以用它把动效按 spec 时序在网页里重放。
生成的 HTML 旁边会有一个同名 `stage/` 目录放图层 PNG，**两者要一起移动**。
它只做数字 id 目标（`--pptx inspect --json` 查 id）。

`motion.ps1` 的关键开关：

| 开关 | 作用 |
| --- | --- |
| （默认） | **只读**打开，绝不改写输入 pptx；需要持久化的内容写 `OutDir\*.com.pptx` |
| `-Strict` | 往返普查发现效果丢失时以非零退出（CI 用） |
| `-SyncTransitions` | 用旧式 COM 枚举重写切换。**会覆盖正确结果**，仅用于探测属性模型 |
| `-NoRender` / `-NoSave` | 跳过 PNG 渲染 / 跳过写回副本 |
| `-Media map.json` | 用显式媒体 map，而不从 spec 推导 |

因为 `Slide.Export` 是 COM 调用，**`OutDir` 传相对路径会按 PowerPoint 自己的工作目录解析**
（不是当前 shell 的目录）；脚本会自己绝对化，但你自己调 COM 时要留意。

## 5. 九个必须知道的坑

1. **多余 preset 包装层会让 PowerPoint 拒绝打开**（本 skill 已修）。PowerPoint 自己写的
   `<p:cTn presetID=... nodeType="afterEffect">` 不能直接嵌在骨架 `<p:cTn>` 下；效果子节点必须
   **内联**在骨架 cTn 里，preset 属性合并到骨架 cTn 上。详见 `references/com-pitfalls.md`。
2. **`<p:sld>` 子元素是 sequence，`<p:transition>` 必须在 `<p:timing>` 之前**（本 skill 已修）。
   位置错了 PowerPoint 不报错、`EntryEffect` 读回 0，`Save()` 后**切换静默消失**。详见
   `com-pitfalls.md` §12。
3. **`<p:cTn id>` 必须按文档顺序递增，父节点先于子节点编号**（本 skill 已修）。顺序错了
   PowerPoint 打开时还数得出来、`Save()` 后**效果静默消失**。详见 `com-pitfalls.md` §13。
4. **同一形状上「入场 + 强调」不兼容 PowerPoint 往返**（本机实测限制，未解决）：单效果 `spin`
   能留住，`fly` + `growShrink`/`spin` 同形状则强调必丢。要叠就用**不同形状**。`motion.ps1`
   的往返普查会报 `LOSS`，加 `-Strict` 会因此失败。详见 `com-pitfalls.md` §14。
5. **动手改 OOXML 后必须用 PowerPoint 真开一次**。结构校验（lxml）能过、PowerPoint 仍可能报
   `E_FAIL` 或静默丢动画。`motion.ps1` 就是这道闸。
6. **静态截图/PDF 走 `motion.py preview` 得到确定性副本**。本 skill 的入场模板从不写隐藏
   （`style.visibility` 只被设成 `visible`），所以实测导出 PNG 并不会"丢形状"——preview 的
   真正作用是**去掉切换与整条动画链**，让导出结果可复现。详见 `com-pitfalls.md` §5。
7. **颜色是 BGR**：`ForeColor.RGB = 0xRRGGBB` 会被存成 `BBGGRR`；且不要用算术表达式生成颜色
   （PowerPoint 收到 Double 会截断）。
8. **切换的旧式枚举几乎不可用**：本机 `0x0A01` 写出的其实是 `<p:strips/>`，`0x1701` 直接非法。
   以 `<p:transition>` 元素为准，COM 属性只用于复核——**不要用枚举去「同步」**，那会把正确的元素
   覆盖成错的（`-SyncTransitions` 默认关闭）。
9. **媒体需要真 PowerPoint**：`AddMediaObject2` 依赖 PowerPoint 自己的转码管线，手写 OOXML 做不到；
   MP4/WAV 可用，未压缩 AVI 会因缺解码器失败。

## 6. 给"静态模板"做动画

现成模板（尤其从网页/设计站下载的）常常**一个动画都没有**：网页端的"动态"是渲染层效果，
导出 pptx 不会变成 OOXML 动画。这时候本 skill 的用法是"从零设计动效"而不是"叠加动效"：

- 形状名是 UI 默认名（`Group 8` / `Freeform 2`），`target` 只能写**数字 id**；
- 按几何 + 文本给形状分角色（eyebrow / title / photo-card / chip / panel）再配效果；
- 注意 `spTree` 顺序是**绘制顺序**不是阅读顺序，spec 按阅读顺序写以决定时间轴；
- 别在同一形状上叠"入场 + 强调"（§5 第 4 条）。

完整方法、形状分类脚本、以及"怎么让人看见动效"（本机导不出 MP4，改用逐形状图层 +
自包含 HTML 播放器）见 `references/template-patterns.md`；
**写 spec 的硬约束见 `references/authoring-rules.md`**；
静态/逐形状导出的 COM 坑见 `references/com-pitfalls.md` §16。

## 7. 与 dsh-ppt-studio 的对接契约

- 坐标：deck.yaml 的 **960×540 pt、1px=1pt、原点左上** = Office COM 坐标，1:1。
- 锚点：`export-pptx.js` 写 `<p:cNvPr id="{nid}" name="{elementId}"/>`，`nid` 从 1000 起全局递增
  → **动画直接按 deck.yaml 的 `elementId` 锚定**。
- 图表：矢量拼绘成多个形状，name 为 `chart-dN-0` 等；`target` 写完整 name 即可。
- 上游门禁结论可继承：本 skill 不改几何，`ppt_verify` 不必重跑。
- 因此**本 skill 不依赖 dsh-ppt-studio 也能用**（对任何 pptx 都成立），装了则体验更顺。
