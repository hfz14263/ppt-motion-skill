# 给"静态模板"做动效：分类 + 时序 + 预览

一次真实作业的完整记录：拿到一个**没有任何动画**的高质量静态模板
（`Quarterly Business Report`，18 页、1440×810pt、86 个媒体文件、全 SVG + 照片），
给前 3 页注入动效，并让动效**能被看见**。

## 1. 先确认它到底有没有动画

不要相信"网页上看着会动"。下载成 .pptx 后把整个包扫一遍（关键是
`p:transition` / `p:timing` / `presetID` / `p:anim*`）：

```bash
python scripts/motion.py inspect --pptx deck.pptx     # 逐页 timing / effects / transition
```

想看得更全（连 notes、master、extLst 一起扫）：

```bash
python - <<'PY'
import zipfile, collections
z = zipfile.ZipFile('deck.pptx')
keys = ['p:transition','p:timing','p:animEffect','p:animMotion','p:anim ',
        'p:animRot','p:animScale','presetID','p:bldLst','p:cond ','morph']
hits = collections.Counter()
for n in z.namelist():
    if n.endswith(('.xml','.rels')):
        x = z.read(n).decode('utf-8','replace')
        for k in keys: hits[k] += x.count(k)
print(dict(hits))
PY
```

该模板的结果是**全 0**。网页端的"动态"（滚动视差、悬停、逐条浮现）属于渲染层，
导出 pptx **不会**变成 OOXML 动画。这是常态，不是文件损坏，也不是下载出错。

再用 COM 复核 PowerPoint 自己的时间线（顺手证明它确实读得到动画）：

```bash
powershell -File scripts/motion.ps1 -Pptx deck.pptx -OutDir review -NoSave -NoRender
# -> "as PowerPoint loads it: 0 timeline effect(s) across 18 slide(s)"
```

## 2. 写 spec：先分类，再配效果

这类模板的形状名是 UI 默认名（`Group 8` / `Freeform 2` / `TextBox 23`），
**没有 elementId 可锚定**，所以 `target` 只能写数字 id。

`player.classify()`（以及 `scripts/player.py` 里的 `shape_index()`）就是做这件事的：
取 `p:spTree` 的**顶层**形状（深度追踪——`p:sp` 也会作为 `<a:sp>` 出现在几何定义里，
形状还会嵌在 group 内），按几何 + 文本给角色：

| 角色 | 判据 | 动画倾向 |
| --- | --- | --- |
| `eyebrow` | 文本、`y < 0.12H` | fade（顶栏标签） |
| `title` | 文本、`h ≥ 70` 且 `y > 0.45H` | riseUp（主视觉大字） |
| `headline` | 文本、`h ≥ 70` | zoom |
| `label` | 其它文本 | fade |
| `photo-card` | `w ≥ 0.28W` 且 `h ≥ 0.35H` | fadedZoom / glide |
| `chip` | `w ≤ 0.22W` 且 `h ≤ 0.12H` | zoom（徽标/页码） |
| `backdrop` | 最后绘制的块 | fade |
| `panel` | 其余带填充的块 | wipe（色条/卡片） |

三个实测容易踩的点：

1. **读序 ≠ z 序。** 封面两个色条在 XML 里是 `Group 5` 然后 `Group 2`，但视觉上先左后右。
   `spTree` 顺序是**绘制顺序**；spec 按**阅读顺序**写（决定时间轴），z 序由 PowerPoint
   自己的 spTree 顺序决定，注入不会改它。
2. **`push` 是切换不是动画。** 写进 `effects` 会直接报 `unknown effect alias 'push'`。
   入场用 `wipe` / `fly` / `riseUp` / `glide` / `fadedZoom` / `zoom`
   （`motion.py catalog --kind entrance` 全列）。
3. **别在同一形状上叠"入场 + 强调"**（`com-pitfalls.md` §14，会被 PowerPoint 静默丢掉）。
   本例每个形状只用一个入场效果，25 个效果全部通过往返普查。

## 3. 时序怎么算

`motion.schedule_spec()` 复刻 `build_timing()` 的分组规则：

- 每个效果各自一个 `<p:par>`；
- `trigger: with` **共享当前组的开始时刻**（所以 spec 里 `with` 要紧跟在它要伴随的效果后面）；
- 下一个 `trigger: after` 从**整组结束时刻**（组内各效果的 max end，不是最后一个的 end）
  起算，再加自己的 `delay`。

最后一条容易被忽略：伴随效果可能比被伴随的更长，后面的 `after` 必须等慢的那个。

## 4. 让人看见动效：`motion.py player`

静态 PNG/PDF 看不出动画，而本机 **`Presentation.CreateVideo` 对所有参数组合都失败**
（详见 `com-pitfalls.md` §16），导不出 MP4。所以用：

```bash
python scripts/motion.py player --pptx animated.pptx --spec motion.yaml --outdir preview
# -> preview/index.html + preview/stage/*.png
```

产物是**每形状一层的自包含播放器**：Tab 切页、▶ 重放、⏭ 连播、右侧时间轴随动效高亮。
HTML 很轻（十几 KB），图层 PNG 在旁边的 `stage/` 目录。HTML 在 `outdir` **里面**，
`src` 写相对路径 `stage/xxx.png`，所以**整个 `outdir` 文件夹可随意移动**。

实现要点（都踩过）：

- `Slide.Export` **无视** `Shape.Visible`，隐藏形状照样画出来 → 不能靠隐藏来隔离形状；
  `Shape.Export(path, 2)` 才认 `Visible` 且给真 alpha（RGBA）。
- base 必须**先导图层、再删形状、最后导整页**，否则图层淡入到位会叠在自己的副本上，
  文字重影。
- 图层和整页导出都走 COM，**相对路径会按 PowerPoint 的工作目录解析**，必须绝对化
  （这一条在 `motion.ps1` 和 `player.py` 里各踩过一次，selftest 有回归覆盖）。
- 图片按**固有比例**放置：只给 width（或只给 height）。同时给 w 和 h 会拉伸变形；
  而图的比例和版面槽位不一致时，按宽度缩放会让高度溢出，压到下面的元素上
  （本 deck 的 P5/P9 各中过一次，解法是**把图按槽位的长宽比生成**）。
- 浏览器侧：**别用 `@keyframes`**（早期版本图层始终不显示）；改用"建元素时设
  `opacity:0`，`setTimeout` 后加 class 触发 CSS transition"。并且**必须等所有图片解码完
  再开始**，否则过渡跑在未解码元素上，图层永久半透明——看起来和被导成空白一模一样。
  （另外 headless Edge 会拒掉几十个并行 `file://` 子资源，所以别把图片散着引用；
  需要单文件分发时可把 PNG 内联成 data URI，但那样 HTML 会到十几 MB。）
- **`--outdir` 里 HTML 与 `stage/` 的相对关系只有一种是对的**：HTML 放 `outdir/index.html`，
  图层放 `outdir/stage/`。写成 `<outdir>.html` 的兄弟文件会让 `src` 全部 404
  （`player.py` 的缺失图片自检会当场报出来，别忽略它）。

## 5. 中文字体与排版（自制图 / 文本框都适用）

- **中文字体没有下标字符**。`q(xₜ|xₜ₋₁)` 里的 `ₜ`(U+209C) `₋`(U+208B) 在微软雅黑/思源黑体里
  都缺字形，matplotlib 会渲染成豆腐块（`Glyph 8348 missing from current font`）。
  解法：数学部分交给 mathtext（`$q(x_t|x_{t-1})$`），中文留在 `$...$` 外面。
- **`Consolas` 也没有下标字形**。把 `x_t` 直接打进 PowerPoint 文本框会显示成字面下划线。
  要真正的下标就**把公式渲染成透明 PNG 再插入**（本 deck 的 P5 就是这么做的）。
- **文字宽度必须实测**，不要目测：用 PIL 载入真实字体
  （`%WINDIR%\Fonts\msyh.ttc` / `consola.ttf`）调 `font.getlength()`，
  按宽度贪心换行算行数。估算换行数是"孤字/溢出"类 bug 的主要来源。
- **按"目标机器有的字体"选字体**，不要按你机器上有的选。Windows 自带的
  微软雅黑 / 黑体 / 宋体基本到处都是；很多设计字体（例如思源黑体）默认没装，
  一旦缺失，别的机器打开就掉字体、重新换行、版面崩。**排版前先查字体是否存在**，
  拿不到首选就按系统自带的等效体量宽度。
- **`python-pptx` 的 `sldSz` 即使 cx/cy 是 16:9，`type` 属性仍留 `screen4x3`**，
  PowerPoint 的"幻灯片大小"对话框会因此显示 4:3。保存后改写该属性即可。

## 6. 校验器：写给自己用的三道闸

机器查不出"该动的没动 / 顺序错 / 单位错"，所以生成 spec 之外还要有校验：

| 校验器 | 查什么 | 抓到的真实 bug |
| --- | --- | --- |
| 覆盖率 | 每个非页脚形状都有且仅有一个效果 | P4 整列、P7 右栏、P8 整表、P6 整卡无动画 |
| 目标有效性 | target 在页面上存在、不重复 | spec 指向不存在的 id、重复效果 |
| 版面 | 出界 / 溢出 / 压页脚 / 文字不在容器内 | P8+P9 底条压页脚、P7 文字浮在框上 |

**校验器自己也会误报**：我的"压页脚"规则第一版把页脚自身、全出血背景、以及
0.03in 余量的正常情况全报了（23 条）。规则必须先定义清楚"什么才算违规"并给出豁免，
且**改完要能抓回它本来要抓的那个 bug**，否则它已经退化成噪声。详见
`references/authoring-rules.md` §F。

## 7. 这次作业的产物

- `animated.pptx` —— 注入后 18 页；第 1–3 页 25 个动效 + 3 个切换，几何未变。
- `preview/index.html` + `preview/stage/` —— 动效播放器（`motion.py player` 生成）。
- `motion.yaml` —— 动效声明，可作为同类模板的起点。
