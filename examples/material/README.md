# 原始素材（source material）

这里放的是**驱动本技能动效设计的原始素材**：两份参考模板 + 当时的作业笔记。

技能里的 `references/morph-and-3d-recipes.md` §4 / §6 / §7 / §8 和
`references/template-patterns.md` 都是从这批素材反推出来的。素材在手，
那些结论就可以被**重新核对**，而不是只能相信文档。

> 这批文件是**素材**，不是技能运行时需要的东西。它进仓库是为了让推导过程可复查；
> 安装技能时会一并复制（`install.ps1` 复制 `examples/`），不需要的话可以删掉这一层。

---

## 文件清单

| 文件 | 大小 | 是什么 |
| --- | --- | --- |
| `summary.txt` | 391 B | 三件事的总纲（见下） |
| `template1/description1.txt` | 1019 B | template1 的**逐步作业笔记**（9 步，含踩坑） |
| `template1/template1.PNG` | 620 KB | 726×406，合成好的底图（麦田 + 人物） |
| `template1/template1.1.PNG` | 2.0 MB | 2560×1440，**全屏截图**，被当作旋转图层用 |
| `template1/template1.2.png` | 62 KB | 320×370，**抠好的人物**（带 alpha） |
| `template1/template1.pptx` | 2.5 MB | template1 成品，2 页 |
| `template2/description2.txt` | 754 B | template2 的作业笔记（7 步） |
| `template2/template2.pptx` | 34 KB | template2 成品，2 页 |

合计约 5.1 MiB。

## summary.txt 说的三件事

| # | 手法 | 落到哪里 |
| --- | --- | --- |
| 1 | **图片本身变化**（同一张图两页，3D 旋转 + 平滑） | `references/morph-and-3d-recipes.md` §4、§6、§7 |
| 2 | **用形状实现效果**（矩形组从画板外滑入展开） | `references/morph-and-3d-recipes.md` §4 |
| 3 | **形状切割图片 + 从左向右的直线**（单页内部动画） | `references/morph-and-3d-recipes.md` §8 |

## 两份 deck 实测长什么样

用 `scripts/motion.ps1 -NoSave -NoRender` 读回来（PowerPoint 自己认的，不是我们的推断）：

```
template1: 2 slides, 960x540pt
  slide 1: shapes=3  effects=0  transition=0xF72
  slide 2: shapes=4  effects=0  transition=0xF72
template2: 2 slides, 960x540pt
  slide 1: shapes=5  effects=0  transition=0xF72
  slide 2: shapes=5  effects=0  transition=0xF72
```

`0xF72` = 3842 = **平滑（Morph）**。注意两点：

- **两页都靠 morph，没有任何时间轴动画**（`effects=0`）。这是"两份 PPT 讲一个变化"
  的典型做法：变化本身是**页间**的，不是页内的。
- template1 两页形状数不同（3 / 4），template2 相同（5 / 5）—— 形状数不同**也可以**
  morph，靠的是形状的匹配规则（见 §1.2），不是逐个对应。

---

## ⚠️ 入库前做过的清理

两个 `.pptx` 里的 `docProps/core.xml` 原本带作者真名
（`dc:creator` 与 `cp:lastModifiedBy` 各一处，**每个文件 2 处**）。
入库的副本已把这两处**清空**。

- **仓库外的原始素材目录里的原件没有改动**，清的是入库的这一份副本。
- `.pptx` 是 ZIP，名字在压缩流里，**普通文本搜索/`grep` 搜不到**——
  必须解压 `docProps/core.xml` 才能看见。同理，PNG 里出现的那串字节是
  **IDAT 压缩数据中的巧合字节序列**，不是文本；这些 PNG **没有任何**
  文本块（只有 `IHDR/gAMA/sRGB/pHYs/IDAT/IEND`），也没有 EXIF。
- 除 `docProps/core.xml` 外，入库副本与原件的**每一个 ZIP entry 都逐字节相同**。

复核命令（`tests/privacy_audit.py` 已覆盖容器内扫描）：

```bash
python tests/privacy_audit.py
```

## 人物图的使用注意

`template1.PNG` / `template1.1.PNG` / `template1.2.png` 里有一位人物。
按 `description1.txt` 第 4.1 步自己的说法，这类"带白边的人物图"是
**从图像生成 AI 拿的**，属于合成素材而非真实人物照片。即便如此，
**对外使用前请自行确认素材来源与授权**，本仓库不对这批素材的权利状态作任何声明。
