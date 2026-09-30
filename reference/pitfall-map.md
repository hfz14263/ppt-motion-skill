# 踩坑编号 → 文件映射

**这一页是 `§n` 这个写法的解释器。** 全仓库（`facts/`、`tests/`、各 reference）有 90+ 处引用写成 `com-pitfalls.md §44` 这种形式。踩坑手册按主题拆成多份之后，**编号不变、含义不变**，只是编号到文件的对应关系挪到了这里。

> **编号只增不减、只挪不改。** 一个编号永远指同一件事，哪怕它换了文件。
> 这是历史引用还能用的唯一原因 —— 编号一旦被复用，旧引用会**静默**指向
> 另一个坑，比链接 404 更糟。

**怎么用**：看到 `§44`，在本表查到它在哪份文件，再 `grep` 那个文件里的
`## 44.`。新写引用时请**直接写文件名**（例如 [`pitfall-transition.md`](pitfall-transition.md) §44），不要再走编号。

本页由 `scripts/build_pitfall_map.py` 从文件标题扫描生成 —— 不是手抄，所以不会漂。

---

| § | 标题 | 在哪个文件 |
| --- | --- | --- |
| §1 | 多余 preset 包装层 → PowerPoint 拒开（最致命） | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §2 | 重复属性 → XML 非法 | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §3 | lxml 不能直接 parse 带编码声明的 str | [`pitfall-ooxml.md`](pitfall-ooxml.md) |
| §4 | 结构校验通过 ≠ PowerPoint 能打开 | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §5 | 静态预览：为什么要 `preview`（实测修正） | [`pitfall-com.md`](pitfall-com.md) |
| §6 | 颜色是 BGR，且不能用算术表达式生成 | [`pitfall-com.md`](pitfall-com.md) |
| §7 | 切换的旧式枚举几乎不可用 | [`pitfall-com.md`](pitfall-com.md) |
| §8 | `p14:dur` 需要声明前缀 | [`pitfall-ooxml.md`](pitfall-ooxml.md) |
| §9 | Office COM 的几个约束 | [`pitfall-com.md`](pitfall-com.md) |
| §10 | Windows PowerShell 5.1 的编码坑 | [`pitfall-com.md`](pitfall-com.md) |
| §11 | 颜色/几何断言要按"布局帧"口径 | [`pitfall-com.md`](pitfall-com.md) |
| §12 | `<p:sld>` 子元素是 sequence，位置错了切换会被静默丢弃（最隐蔽） | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §13 | `<p:cTn id>` 必须按文档顺序递增，否则强调动画会被静默丢弃 | [`pitfall-silent-drop.md`](pitfall-silent-drop.md) |
| §14 | 同一形状上"入场 + 强调"不兼容 PowerPoint 往返（实测限制，未解决） | [`pitfall-roundtrip.md`](pitfall-roundtrip.md) |
| §15 | `motion.ps1` 默认只读，绝不回写输入（旧版会毁掉源文件） | [`pitfall-com.md`](pitfall-com.md) |
| §16 | 静态导出与逐形状导出（做"动效预览"必需） | [`pitfall-com.md`](pitfall-com.md) |
| §17 | 动画的**中间态**在本机不可观测 | [`pitfall-com.md`](pitfall-com.md) |
| §18 | Round-trip 会以第二种方式咬"一个形状多个效果" | [`pitfall-roundtrip.md`](pitfall-roundtrip.md) |
| §19 | 「静默丢弃」的第三种成因：元素放错了父节点 | [`pitfall-silent-drop.md`](pitfall-silent-drop.md) |
| §20 | Morph（平滑）**可以注入** —— 旧结论是写法错误，不是版本限制 | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §21 | 3D 相机（`scene3d`）可以被注入，且角度原样保留 | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §22 | 判断"不支持"之前，先排除"我写错了" | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §23 | 「三维旋转」的界面角度 ≠ OOXML 的 `lat/lon` | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §24 | 造场景图时，图的"结构"决定效果成不成立 | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §25 | `ThreeD.RotationX/Y` 写出的是**平行投影**，永远做不出"躺平的地面" | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §26 | `a:rot` 的 `lat/lon` 必须是 0~21600000，负数会让**整个文件**损坏 | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §27 | `ThreeD.Perspective` 是**开关**，不是强度——而且 `0` 和 `1` 等价 | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §28 | 把"参数 → 渲染结果"当作待测对象，而不是待查文档 | [`pitfall-measurement.md`](pitfall-measurement.md) |
| §29 | 窗口化图片填充：**一个函数里踩出四种静默失败** | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §30 | 内缩公式：独立推导 + 实测吻合，但**资料里的简化式不可照抄** | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §31 | 「往序列里插一个已经存在的元素」—— 本项目**所有**损坏文件都是这一个原因 | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §32 | `CreateVideo` 报"完成"却拿不到文件 —— 三个参数全错 | [`pitfall-tooling.md`](pitfall-tooling.md) |
| §33 | 测量值必须和它的**自变量**一起记录 —— 否则会写出错误的"规律" | [`pitfall-measurement.md`](pitfall-measurement.md) |
| §34 | 「翻过去」是**下沉 + 三维旋转**两件事 —— 只做旋转会得到"被压扁的一页" | [`pitfall-morph-3d.md`](pitfall-morph-3d.md) |
| §35 | 抠图**盖不住**照片里那个人时，让它"虚" —— 不要硬撑尺寸 | [`pitfall-media.md`](pitfall-media.md) |
| §36 | 素材是"截图"还是"照片"，会改变整个观感 —— 先导出 `ppt/media/` 看一眼 | [`pitfall-media.md`](pitfall-media.md) |
| §37 | "色块"替代不了"细节" —— 要复现一张截图，就去拿**真的**截图 | [`pitfall-media.md`](pitfall-media.md) |
| §38 | 页间差值**有方向**：y 轴向下，别把加减号写反 | [`pitfall-measurement.md`](pitfall-measurement.md) |
| §39 | 露出的底（BG）**必须有可辨认的内容**，否则就是一块空白板 | [`pitfall-media.md`](pitfall-media.md) |
| §40 | 切换的候选必须**一个一份 deck** —— 合并起来探就没有结论可言 | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §41 | `<mc:Fallback>` 不总是 `<p:fade/>` —— core 元素在 Fallback 里**重复自己** | [`pitfall-silent-drop.md`](pitfall-silent-drop.md) |
| §42 | 切换 / 翻转 / 库 / 摩天轮 / 传送带 **缺 `dir` 会被拒开整份文件** | [`pitfall-file-corruption.md`](pitfall-file-corruption.md) |
| §43 | 切换的机制：**写在哪一页、能装几个、时长听谁的** | [`pitfall-transition.md`](pitfall-transition.md) |
| §44 | 切换的**形态**：观众看到的是什么运动 | [`pitfall-transition.md`](pitfall-transition.md) |
| §45 | 推送：github.com 被挡时，改走 api.github.com | [`pitfall-tooling.md`](pitfall-tooling.md) |
| §46 | 证明"运动"不能用静态帧 —— 抽帧再多也证明不了方向 | [`pitfall-transition.md`](pitfall-transition.md) |
| §47 | 指标有盲区，而且盲区是"静默"的 —— 旋转从三个指标里漏过去了 | [`pitfall-transition.md`](pitfall-transition.md) |
| §48 | §47 的第二次出现：判"镜像"时单路指标又把 `push` 判成了怪东西 | [`pitfall-transition.md`](pitfall-transition.md) |
| §49 | 属性不是到处都能加：`dir` 写错地方会让 PowerPoint 直接打不开 | [`pitfall-silent-drop.md`](pitfall-silent-drop.md) |
| §50 | 非法属性值不报错，只是**静默退回默认** —— 你以为写了，其实没写 | [`pitfall-silent-drop.md`](pitfall-silent-drop.md) |
| §51 | 切换与页内动画**不是正交的** —— 动画要排队等切换演完 | [`pitfall-transition.md`](pitfall-transition.md) |

---

## 文件职责

| 文件 | 管什么 |
| --- | --- |
| [`pitfall-file-corruption.md`](pitfall-file-corruption.md) | 文件打不开 / 报损坏 —— 结构非法的全部成因 |
| [`pitfall-silent-drop.md`](pitfall-silent-drop.md) | XML 里有、PowerPoint 不认 —— 静默丢弃与静默退回 |
| [`pitfall-roundtrip.md`](pitfall-roundtrip.md) | 往返（round-trip）后被改掉或吃掉 |
| [`pitfall-com.md`](pitfall-com.md) | COM 与脚本环境 —— 界面层的约束与陷阱 |
| [`pitfall-ooxml.md`](pitfall-ooxml.md) | OOXML 注入写法 —— 元素、属性、命名空间 |
| [`pitfall-measurement.md`](pitfall-measurement.md) | 测量与判据 —— 指标盲区、单变量纪律 |
| [`pitfall-morph-3d.md`](pitfall-morph-3d.md) | Morph / 3D / 图片填充 —— 能力判定与写法 |
| [`pitfall-media.md`](pitfall-media.md) | 素材与观感 —— 截图/照片、抠图、露底 |
| [`pitfall-transition.md`](pitfall-transition.md) | 页面切换 —— 机制、形态、方向、与动画的关系 |
| [`pitfall-tooling.md`](pitfall-tooling.md) | 工具链 —— 推送、导出、视频、运行方式 |
