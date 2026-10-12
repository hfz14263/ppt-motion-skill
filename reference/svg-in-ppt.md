# SVG 在 PPT 里的能力边界（2026-10-10 实测）

> 本文回答三个问题，每条都有 PowerPoint 实测或官方原文支撑：
> 1. SVG 能做什么、不能做什么？
> 2. 它和 Morph 是什么关系？
> 3. 本项目怎么用它？

---

## 1. 结论先说

| 问题 | 答案 |
| --- | --- |
| SVG 能保持放大不糊吗 | **能**。这是它唯一的、但确实很值钱的用途 |
| SVG 能自己动吗 | **不能**。CSS `@keyframes` 与 SMIL 在插入时被剥离 |
| SVG 动态需要 Morph 吗 | **不需要，两者无关**。Morph 是页面切换，SVG 是页面内的静态矢量 |
| 要逐元素动画怎么办 | 先在 PowerPoint 里「转换为形状」，之后按普通形状处理 |
| Morph 时会糊吗 | **会**。过渡期间 PowerPoint 拿 PNG 兜底做插值 |
| 本项目怎么注入 | 走 COM 层 `AddPicture`，**不手写 XML** |

---

<a id="s-now"></a>

## 2. 不能自己动 —— 这是最容易误解的一点

**微软官方 Q&A 原文**：

> "while all Office applications now include support for SVG images,
> **they do not support SVG Animation**. If you need to animate the elements of
> an SVG image in PowerPoint, the animation would need to be done in
> PowerPoint."

所以网上那些"SVG 动画 PPT"**不是 PowerPoint 做到的**。它们通常是三条路之一：

| 做法 | 实际是什么 | 代价 |
| --- | --- | --- |
| 录屏转视频 | 视频 | 无交互、占体积 |
| GIF | 位图动画 | 画质低、无透明度 |
| 网页版 PPT（reveal.js 等） | 浏览器 | **不是 pptx**，PowerPoint 打不开 |

**"动态 SVG PPT"这个选项在 PowerPoint 里不存在。** 说清这一点，
比反复试错省时间。

**动效只能来自 `<p:timing>`** —— 也就是本项目注入层写的那套 `<p:timing>`。
一个 SVG 整体可以当普通形状加效果；但 SVG 内部每条路径不能。

<a id="s-morph"></a>

## 3. 和 Morph 的关系：无

这三件事经常被混在一起谈，但它们在 OOXML 里落在三个完全不同的地方：

| 你想要的效果 | 机制 | OOXML 落点 | 本项目支持 |
| --- | --- | --- | --- |
| 页面整体动一下、形状补间 | **Morph** | `slideN.xml` 的 `<p:transition>` + `p159:morph` | ✅ 配方 1 |
| 图形放大不糊 | **SVG** | `ppt/media/*.svg` + `asvg:svgBlip` 扩展 | ✅ 本文档 |
| 点一下跳到某页并放大 | **Zoom** | 独立 part（`/slideN/zoomM`） | ❌ 未实现 |

**「动态背景」的正确做法（两条，都不需要 SVG）**：

- **Morph 路线**：两页放同位置的形状，只改数值 → 自动补间。最优雅，已验证。
- **视频路线**：内嵌 MP4 循环播放。已支持，但无交互、占体积。

<a id="s-how"></a>

## 4. 本项目怎么用：`svg:` 段

```yaml
slides:
  - page: 3
    svg:
      - src: icon.svg
        bounds: [60, 60, 240, 240]      # 左上宽高（pt）
        elementId: ICON                  # 命名后可在 effects 里引用
```

跑法与媒体层一致：

```powershell
pwsh -NoProfile -File scripts/motion.ps1 -Pptx deck.pptx -Spec spec.yaml -OutDir out
```

输出落在 `out/*.com.pptx`。

<a id="s-why-com"></a>

## 5. 为什么走 COM 层，而不是手写 XML

**第一版确实是手写 XML 的**（在 `asvg:svgBlip` 扩展里塞 svg 关系 +
`Default Extension="svg"`），而且往返无损、结构自证通过。

但那是在跟一个「**结构合法 ≠ PowerPoint 自己那样写**」的问题较劲。
实测对比两份文件的同一处：

```xml
<!-- COM 版（PowerPoint 自己写的） -->
<p:cNvPr id="3" name="ICON">
  <a:extLst><a:ext uri="{FF2B5EF4-...}">
    <a16:creationId id="{83811D3A-...}"/>   <!-- 多出来的这一段 -->
  </a:ext></a:extLst>
</p:cNvPr>
```

手写版没有 `a16:creationId`。它不致命，但意味着**每次都要靠往返验证
才知道自己有没有写对**。而一行 COM 代码就解决了：

```powershell
$shape = $slide.Shapes.AddPicture($src, 0, -1, $x, $y, $w, $h)
```

**让 PowerPoint 写它自己的格式**，整类「合法但微妙不对」的风险就没了。
这与项目既有分工一致：媒体和 SVG 都归 COM，`<p:timing>`/`<p:transition>`
才归 XML 层（因为那些必须不动几何）。

`motion.py` 侧只负责两件轻活：解析 `svg:` 段（供 `check` 报告）、
给 `[Content_Types].xml` 加 `Default Extension="svg"`（缺了 PowerPoint 拒收）。

<a id="s-pitfall"></a>

## 6. 三个坑

1. **Morph 会让 SVG 发虚。** 过渡时用 PNG 兜底插值，动画结束才切回矢量。
   大幅放大的图形这段明显糊。→ 想要真形状 Morph，先「转换为形状」。
2. **「转换为形状」是手动操作，不是 spec 字段。** 转换后 SVG 每条路径变成
   原生 shape，才能单独加效果。批量转换要另写 COM 脚本，本项目未做。
3. **复杂 SVG 会卡。** 路径数量多的插画插入后动画掉帧。→ 导入前简化路径
   （Inkscape 的 `Path > Simplify`），文字转曲。

<a id="s-render"></a>

## 7. 渲染器兼容（顺带记录）

非 PowerPoint 的渲染器对 SVG 支持参差，实测/已知情况：

- **LibreOffice**：能正常渲染 SVG
- **浏览器内核渲染器**：`createImageBitmap` 不认 `image/svg+xml` → **画不出来**，
  只剩轮廓和阴影（除非显式声明 MIME）
- **纯位图解码器**（如 Rust `image` crate）：无法嗅探 SVG 格式 → 计入
  `skipped_images`，同样是空洞

**所以 PPT 里放了 SVG，导出预览/服务端渲染时可能少一块。** 这是"能生成"
和"到处都能看"的差别，报交付时要提。

---

**验证方式**：2026-10-10 实测 —— COM 插入 SVG 后 PowerPoint 识别为
`type=28 msoGraphic`（普通图片是 `type=13`），产出 `ppt/media/image2.svg`
+ `asvg:svgBlip` 扩展，往返保存后扩展仍在。
