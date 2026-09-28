# ppt-office-motion 的知识构成与处理流程

给一个已经很完整的 pptx 叠动画，而**完全不动版面几何**。

---

## 一、这个 skill 要解决的核心矛盾

上游（dsh-ppt-studio 这类工具）刚把版面调好：元素不打架、文字不溢出、页脚不压。
一旦用 PowerPoint 打开、编辑、保存，PowerPoint 会跑**自己的排版引擎**（字体解析、autofit 重排、
行距重算），上游刚通过的结论立刻作废。

所以分工被切成三层，边界必须干净：

| 层 | 归属 | 职责 |
| --- | --- | --- |
| 版面 | 上游 | 元素位置、字号、溢出 |
| 动画 / 切换 | **本 skill** | 纯 OOXML 增量注入，几何零改写 |
| 媒体 | 本 skill（COM 层） | 内嵌 MP4/WAV |

能做到"几何零改写"的原因是：注入只碰 `ppt/slides/slideN.xml` 里的
`<p:timing>`（动画）和 `<p:transition>`（切换），以及 `[Content_Types].xml`。
形状的 `<p:spTree>` 及其 `<a:xfrm>`（几何）**一个字节都不动**，
所以上游的门禁结论可以直接继承，不必重跑。

---

## 二、用到的知识

### 1. OOXML / PresentationML 结构

必须知道 `<p:sld>` 的子元素是一个 **`xsd:sequence`**，顺序固定：

```
cSld → clrMapOvr? → transition? → timing? → extLst?
```

`transition` 必须在 `timing` **之前**。位置写错，PowerPoint 不报错、不弹修复，
但 `EntryEffect` 读回 0，`Save()` 之后**切换静默消失**。

`<p:timing>` 是一棵时间树：

```
tmRoot → mainSeq → 每个效果一个 <p:par>
                      └─ <p:cTn presetID= presetClass= presetSubtype=>  ← 预设属性
                           └─ <p:childTnLst>
                                ├─ <p:set>         style.visibility=visible（入场先显形）
                                └─ <p:animEffect filter="...">  ← 真正的效果 + 方向
```

关键约束：**`<p:cTn id>` 必须按文档顺序递增，父节点先于子节点编号**。
顺序错了 PowerPoint 打开时还数得出来、`Save()` 后效果静默消失。

### 2. PowerPoint COM 对象模型（win32com / pywin32）

自动化那一层。难点不在"会不会调"，而在**哪些属性可信**：

- `SlideShowTransition.EntryEffect` 是**旧式枚举**，本机映射离预设号很远
  （`0x0A01` 号称 fade，实际写出 `<p:strips/>`；`0x1701` 直接非法）。
  所以 **XML 元素是权威，COM 属性只用于复核**。用它去"同步"会把正确结果覆盖成错的。
- `Shape.Visible` 是**静态属性**，不是动画运行中的可见性。
- 颜色是 **BGR**：`ForeColor.RGB = 0xRRGGBB` 会存成 `BBGGRR`；
  也不要用算术表达式生成颜色（收到 Double 会截断）。
- COM 调用的相对路径按 **PowerPoint 自己的工作目录**解析，不是 shell 的 cwd。

### 3. 导出的真实行为（决定验证策略）

这一块最反直觉，也是整个验证体系的地基：

| 方法 | 行为 |
| --- | --- |
| `Slide.Export` | 导出整页，但**无视 `Shape.Visible`** |
| `Shape.Export(path, 2)` | 认 `Visible`，且给**真 alpha**（RGBA），所以能当图层用 |
| `SlideShowView.Export` | 本机**根本没有这个成员** |
| `Slide.Export` / `Shape.Export` | 一律渲染**终态**，无视动画中间态 |
| `Presentation.CreateVideo` | 本机**完全不可用**（Quality 0 返回成功但不产文件；1/2 报 `E_INVALIDARG`） |

两个直接后果：

- **导不出 MP4**，所以"让人看见动效"只能另想办法（见下面 player）。
- **动画的中间态不可观测**，所以方向性擦除（`wipe(left)` 到底往哪边擦）
  无法自动校准，只能靠人眼看一次。

### 4. 渲染语义：透明、合成、图层

- 入场动画的 `<p:set>` 把 `style.visibility` 设成 `visible`，
  **模板里从不写 `hidden`**，所以静态导出的 PNG 不会"丢形状"。
- **半透明填充会和背后的底色相乘。** 一张为白底画的图（曲线下的淡色填充）
  放到深蓝底上会糊成一块灰。所以"给深色底用"的图必须**单独画成不透明面板**，
  而不是把白底图调透明度。
- 逐形状导出做图层：**先导图层 → 再删形状 → 最后导整页当背景**。
  顺序反了就重影（背景里已烘焙一份，图层淡入后正好叠在自己副本上）。

### 5. 字体与排版工程

- **单位**：`1 pt = 12700 EMU`；`1 in = 72 pt`。960×540 pt = 13.333×7.5 in = 16:9。
  三种单位混用会静默错位，这是本项目最容易翻车的地方。
- **字距**：OOXML 用 `a:rPr/@spc`，单位是**百分之一磅**；python-pptx 没有 API，要直接写 XML。
  大字号需要**负字距**，小标签需要**正字距**。
- **没有 autofit**：每个文字框必须按**真实字形**量过（PIL 读 ttc/ttf 算 advance width），
  估算是"文字压到底部色块上"的成因。
- 字体可用性要先查。**"这个字体在不在"是环境问题，不是代码问题**：
  目标是思源黑体时，很多 Windows 机器上并没有，需要用系统自带的等效细体顶替，
  并且**按取到的实际字体量宽度**，不要按你想要的那个字体量。

### 6. 设计系统知识（为什么"规范"反而难看）

- **Swiss / 国际主义排版**：靠数学网格、左对齐右参差、非对称构图；把图当**数据**，不当装饰。
- **TED 的幻灯片指南**：一页一个论点；"消灭标题+项目符号"；每页最多六行。
- **冲击力来自 SCALE，不来自字重**：大字号用**细体**+紧字距，不要中等字号加粗。
  粗体只留给小标签，这样它才真的读作"强调"。
- **口音色要省着用**（全篇 ≤8 次）。到处都是重音，等于没有重音。

### 7. 脚本与工具链

python-pptx（建 deck / 读形状）、lxml（解析 slide XML）、zipfile（改包内部件）、
Pillow + matplotlib（出图，中文字体配置）、pywin32（COM）、YAML（spec）、
PowerShell（COM 脚本宿主，用 `powershell.exe`；5.1 是 Windows 自带的，`pwsh` 不一定有）。

---

## 三、整体处理流程

### 标准作业流程 S0–S7

```
S0  确认上游已有 .pptx（已过 verify，或任意现成 pptx）
S1  motion.py inspect --pptx deck.pptx        # 看清 elementId ↔ 形状 id
S2  写 motion spec（按 elementId 声明效果）
S3  motion.py apply --pptx in --spec m.yaml --out out --assert-geometry
S4  verify_motion.py --pptx out --source in   # 结构 + 几何自证
S5  motion.ps1 -Pptx out -Spec m.yaml -OutDir review -ExportPdf -Strict
                                              # 真渲染 + 媒体 + PDF + 往返普查
S6  motion.py player --pptx out --spec m.yaml --outdir preview
                                              # 动效预览，浏览器打开
S7  read_image 看 review/render/*.png         # 视觉审阅
```

**S3 必须带 `--assert-geometry`，S4 必须比对 `--source`。**
这两步输出 `geometry: UNCHANGED`，才是本 skill 的核心承诺兑现。
一旦报 CHANGED，说明注入了不该注入的东西，停下排查。

⚠️ **加上 S5 的 `-Strict` 和第 3 步的覆盖率审计，一共四道闸——但它们仍然证不了
"动效正确"。** 这一点见 [`authoring-rules.md`](authoring-rules.md) §H。

### 为什么需要这么多道闸：PowerPoint 的"静默失败"

**这是整个设计的动机。** PowerPoint 有一类失败方式：
文件结构完全合法、打开不报错、不弹修复，**动画就是没了**。

已知的四种：

1. `<p:transition>` 位置违反 sequence → 切换静默消失
2. `<p:cTn id>` 非递增 → 效果静默消失
3. 同一形状上"入场 + 强调" → 强调静默消失（本机固有限制，未解决）
4. 同一形状挂多个效果（逐段揭示）→ 往返时折叠成一行，多出来的丢掉

所以：

- **lxml 结构校验能过的东西，PowerPoint 仍可能拒绝或丢内容。**
- **必须用真 PowerPoint 开一次**（`motion.ps1` 就是这道闸）。
- `motion.ps1 -Strict` 把"保存后动画变少"变成**失败退出**，不然它只是打印个警告。

### 几何自证怎么做的

对每张 slide 提取所有形状的 `(x, y, cx, cy, rot)`，算一个指纹；
注入前后各算一次，**必须完全相同**。这是"零改写"的机器证明，不是承诺。

### ⚠️ 门禁通过 ≠ 动效正确

必须说清这个 skill **证明不了**什么，因为这三件才是实际翻车的地方：

1. **该动的都动了** —— 漏掉的形状在放映一开始就可见，不会报任何错
2. **顺序对** —— 标题在内容前、结论在论据后
3. **单位对** —— EMU / pt / inch 混用静默错位
4. （新增）**方向对** —— 八个 `dir` 的预设三元组完全一致，只差 `filter` 字符串；
   装反了三个结构闸**全部通过**

所以流程里还有三个**额外校验器**：

- **覆盖率**：报告"没有任何效果"的形状。这是静默错位的主因——
  一个漏掉的形状在放映开始就可见，**不会报任何错**，症状是"这一栏从不动画"。
- **目标有效性**：每个 target 必须在该页存在；几何不能出界、不能压页脚。
- **版面**：溢出 / 越界 / 压页脚，外加**按实测文字宽度**（不是框宽）比较的重叠检测。

### 锚定与"不许手写 id"

spec 的形状 id **禁止手写**，必须由实测几何或**摆放时注册的语义角色**生成。
原因是一次真实事故：手写 id 列表在每次改版后静默指向别的形状
（出现过"正好差一位"，导致整张卡没有动画）。

推论（这条最贵）：**改了 deck 就必须重新生成 spec**。
"觉得 spec 没动"就复用旧的，正是上面那个 bug 的成因。

### 动效预览：player

因为导不出 MP4、静态图看不出动画，所以把 spec 的时序在浏览器里重放：

- 逐形状用 `Shape.Export` 导成**带 alpha 的图层**，base 是删掉这些形状后的整页
- 生成自包含 HTML：每个形状一层，用 opacity + `setTimeout` + CSS transition 驱动
- **不用 `@keyframes`**（会和"播放中"状态冲突），且**所有图片必须预解码**
  （否则 transition 在未解码的图上跑完，图层永远停在半透明——看起来像导出坏了）
- 图层必须**按方向渲染起始裁剪**，否则预览会把所有擦除画成一个方向，**验证的就是错的东西**

---

## 四、一句话总结

**技术上**，是 OOXML 结构 + PowerPoint COM + 渲染真实行为三样东西的交叉，
难点几乎全在"哪些 API 的表面行为是假的"。

**方法上**，是一套**分层验证**：结构自证 → 几何自证 → 真渲染往返 → 覆盖率 → 版面，
每道闸对应一次真实翻车，而且**每一道都要先校准**——
因为校验器本身也会误报（版面重叠检测第一版误报 6 条，全是标签框比文字宽）。
