# 配方 · 窗口化图片填充（用形状切割图片）

> **单页内部**动画：让一张图片像被窗口「扫过」一样露出来。关键在「默认填充是拉伸的，要做窗口必须给负偏移」。
>
> **`§` 编号沿用拆分前的原文**（`§1.2`、`§2.3`、`§8.3` 等，别处正引着它们）。
> 完整目录见 [`morph-and-3d-recipes.md`](morph-and-3d-recipes.md)。

---

## 8. 单页内部动画：用形状切割图片 + 从左向右的直线

summary 第 3 条，不依赖页面切换。两条路线都成立，实测都做了出来。

### 8.1 形状怎么"切割"图片

**形状的填充就是图片**，于是形状的轮廓决定了图片露出哪一部分 —— 这正是
PowerPoint「合并形状 → 相交」产出的结构：一个 `<p:sp>`，它的 `spPr` 里带
`<a:blipFill>`。

COM 侧也能直接做：

```powershell
$shape.Fill.UserPicture($img)     # 形状的填充变成图片，写出 <a:blipFill>
```

### 8.2 关键：默认填充是**拉伸**的，要做窗口必须给负偏移

`Fill.UserPicture` 写出的是

```xml
<a:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>
```

空的 `<a:fillRect/>` 意味着**图片被压扁塞进形状**——那不是"窗口"，只是"缩略图"。

要做成真正的窗口（形状只露出大图的一角），要给 `fillRect` **负的内缩**。
单位是**形状尺寸的千分之一百分比**（1000ths of a percent）：

```
窗口位于 x、宽 w，图片要铺满 0..slideW：
    l = -(x / w) * 100000
    r = -((slideW - x - w) / w) * 100000
```

实算核对（x=380, w=200, slideW=960）：
`l = -190% × 200pt = -380pt` → 填充左边界落在画布 x=0；
`r = -190% × 200pt = -380pt` → 填充右边界落在画布 x=960。正确。

实测：三个窗口分别取 `l/r = -90000/-290000`、`-190000/-190000`、`-290000/-90000`，
渲染后视觉复核确认**三个窗口显示的是同一张照片的不同部分**（像三块拼图），
而不是三张压扁的完整照片。

### 8.3 两条动画路线

| 路线 | 做法 | 效果 |
| --- | --- | --- |
| **A 静止窗口 + 擦除** | 窗口位置固定，给窗口加 `wipe`（`dir: right`） | 该切片**从左向右**逐渐显现；多个窗口按 delay 依次擦出，形成扫描序列 |
| **B 窗口移动** | 给窗口加**动作路径**（`pathRight`），窗口自身平移 | 可见区域**横扫**过图片，像取景框/放大镜扫过 |

**窗口本身现在也入 spec 了**（`fills:`），不必再手写注入：

```yaml
fills:
  - {target: WIN1, window: [60, 150, 200, 300]}     # [x, y, w, h]，单位 pt
  - {target: WIN2, window: [380, 150, 200, 300]}
  - {target: WIN3, window: [700, 150, 200, 300]}
effects:
  - {target: WIN1, effect: wipe, dir: right, duration: 0.8, trigger: after}
  - {target: WIN2, effect: wipe, dir: right, duration: 0.8, trigger: with}
  - {target: WIN3, effect: wipe, dir: right, duration: 0.8, trigger: with}
```

`fills` 和 `cameras` 一样是**形状属性，不进时间轴**，所以与 `effects` 分开写。
`window` 缺省时 `picture` 取整个画布（全出血），也就是教程里的用法；
要给**非全出血**的图片开窗就显式写 `picture: [x, y, w, h]`。
负内缩由 `motion.window_insets()` 算，实测与 PowerPoint 自己写的**数值一致**。

> 一个前提：`<a:blip r:embed>` 必须是**真实存在的关系**。手写一个假 id
> 会让 PowerPoint 画「缺图占位符」，而那个占位符看起来和「内缩被忽略」一模一样 ——
> 这条曾让诊断走偏。图片的嵌入属于媒体层，不在本函数职责内。

回归：`tests/test_fill_window.py`（28 条断言）。

### 8.4 注意

- **动作路径是"起点"**：静态导出（PNG/PDF）看到的是路径**起点**的窗口位置，
  不是运动过程。要看"动"的效果要用 `motion.py player` 重放，或导 MP4。
- 路线 B 里如果窗口**移动而填充随之移动**，窗口内容不变、只是整体平移；
  要让内容"扫描"，要么把图片单独放底层、窗口作为遮罩，要么同时动画化填充偏移。
