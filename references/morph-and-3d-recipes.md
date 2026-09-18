# Morph 与 3D 相机的注入配方

来源：两份真实 PowerPoint 文件（`template1` / `template2`，取自本次分析用的素材目录），
外加本机实测复核。**这两份文件是 PowerPoint 自己写的**，所以它们的 XML 就是
"官方写法"的地面真值 —— 比任何文档都可靠。

---

## 1. Morph（平滑）切换

### 1.1 元素写法

```xml
<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
  <mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" Requires="p159">
    <p:transition spd="slow" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" p14:dur="2000">
      <p159:morph option="byObject"/>
    </p:transition>
  </mc:Choice>
  <mc:Fallback>
    <p:transition spd="slow"><p:fade/></p:transition>
  </mc:Fallback>
</mc:AlternateContent>
```

| 项 | 值 | 备注 |
| --- | --- | --- |
| 元素 | `p159:morph` | **不是 `p:morph`**，后者在 `p:` 命名空间里不存在，会被静默丢弃 |
| 命名空间 | `…/powerpoint/2015/09/main` | 是 **2015/09**，不是 2019 |
| 包裹 | `mc:AlternateContent` + `mc:Choice Requires="p159"` | 必须 |
| 降级 | `mc:Fallback` → `<p:fade/>` | 不能省略，WPS/旧版/在线预览靠它 |
| 时长 | `p14:dur="2000"`（毫秒） | 需 `xmlns:p14` |
| 选项 | `option="byObject"` | 默认按对象匹配 |

### 1.2 匹配规则

morph **按形状名/id 跨页配对**，然后补间两页之间的差异。所以流程是：

1. 复制第 1 页 → 第 2 页（**形状名与 id 必须保持一致**）；
2. 在第 2 页改形状的**位置 / 尺寸 / 旋转 / 3D 相机角度**；
3. 把 morph 加在**第 2 页**（被进入的那页）。

**两页必须存在真实差异**，否则什么都不动。
（`material/template2` 的两页**不同**：五个矩形组在第 1 页位于 `left=-180.8pt`，即**画板左侧外**；第 2 页铺开到 `55.8 / 236.7 / 417.5 / 598.3 / 779.2pt`。morph 把这一位移补间出来 —— 这正是 description2 第 6 步的做法。）

### 1.3 验证

```bash
# 写进去没有
python -c "import zipfile,re; x=zipfile.ZipFile('out.pptx').read('ppt/slides/slide2.xml').decode('utf8'); print('p159:morph' in x)"

# PowerPoint 认不认（读回 3842 = 0xF72 即成功）
powershell -NoProfile -File scripts/motion.ps1 -Pptx out.pptx -OutDir review
```

---

## 2. 3D 相机（scene3d）

### 2.1 写法

```xml
<a:scene3d>
  <a:camera prst="perspectiveRelaxedModerately">
    <a:rot lat="17400000" lon="0" rev="0"/>
  </a:camera>
  <a:lightRig rig="threePt" dir="t"/>
</a:scene3d>
```

放在形状 `<p:spPr>` 内，**追加在 `</p:spPr>` 之前**。

### 2.2 三条硬约束

1. **只有 `prst="perspective*"` / `"legacyPerspective*"` 有透视。**
   `orthographicFront`（COM 设 `ThreeD.RotationX/Y` 时的默认输出）是**平行投影**，
   永远没有灭点，角度怎么调都不会"躺下去"。
2. **`lat/lon/rev` 必须是 `0..21599999`。** 规范写的上界 `21600000` 会让
   PowerPoint 报**整个文件损坏**（`0x80070570`）。负角度写成 `360 − x`：
   290° → `17400000`。
3. **`scene3d` 必须在 `<a:extLst>` 之前。** 追加在 `</p:spPr>` 前是对的；
   但**不要**在整段 `<p:sp>` 里搜 `<a:extLst>` 当锚点 —— `<p:cNvPr>` 里也有一个
   （`a16:creationId`），插到那里会变成非法位置而被静默丢弃。

### 2.3 实测对照（Office LTSC 2024，16.0.17928.20148）

图片 + `scene3d`，两页只有相机 `lat` 不同：

| 页 | `lat` | 渲染 bbox | 远/近边比 | PowerPoint 读回 |
| --- | --- | --- | --- | --- |
| 1 | `0` | 534×400 | 1.000 | `PresetCamera=62, RotationY=0` |
| 2 | `17400000` | **612×140** | **1.249** | `PresetCamera=62, RotationY=-70` |

结论：
- 相机角度**确实生效**，平面被压扁 65%（"翻倒"）；
- PowerPoint 把 `290°` **归一化成 `-70°`**（`290 − 360`）—— 这是正常的，不是写错；
- `PresetCamera` 读回 **62** 代表 `perspectiveRelaxedModerately`；
  读回 **`-2`** 代表 PowerPoint **没有接受**你写的 scene3d（位置或值有问题）。

---

## 3. 两个测量陷阱（都会让你误判"功能不生效"）

### 3.1 形状被画布裁切 → 收敛比恒为 1.000

判断"是否有透视"要量**远边宽 / 近边宽**。但如果形状底部超出画布，
两条采样带都落在裁切边上，宽度相同 → **ratio 永远 1.000**，
看起来就像"相机没生效"。

**第一次测 autoshape 时就栽在这里**：矩形 y 到 420pt 而画布只有 405pt。

**对策**：把被测形状完整放进画布内，并在测量里显式检查是否触碰边缘。

### 3.2 元素写在合法位置之外 → XML 存在但被忽略

`scene3d` 写进了文件、字符串检查也能找到，但 PowerPoint 读回 `PresetCamera=-2`。
**"XML 里有" ≠ "PowerPoint 采纳了"。**

**判据**：永远用对象模型**读回**确认（`PresetCamera` / `EntryEffect`），
不要只看字符串。

---

## 4. 从两份 material 模板提取的手法

| 模板 | 手法 | 拆解 |
| --- | --- | --- |
| `template1` | 图片 3D 翻倒 | 同一张图两页；第 1 页 `camera lat=0`，第 2 页 `lat=17400000`；两页都加 morph。另含"虚化+亮度校正做背景层、抠人物做前景层"的图层技巧 |
| `template2` | 目录卡片展开 | **真实可运行**：5 个矩形组第 1 页在 `left=-180.8pt`（画板外），第 2 页铺开成目录；两页加平滑 → 矩形**从左侧滑入展开** |

`template1` 的额外技巧（非 3D 部分）：

- 背景图先做**虚化**（半径 8–10，界面里的"模糊"可能叫"虚化"）+ **亮度 −40**；
- 再叠一张同尺寸原图；
- 人物单独**抠图**（`图片格式 → 删除背景`，或先用 AI 生成带白边的人物图更好抠）；
- 背景层加 3D 旋转做纵深，人物层放大做层次。

---

## 5. 和 spec 的关系

**Morph 已经纳入 spec**（`build_transition` 支持 `type: morph`）：

```yaml
transition: {type: morph, duration: 2.0, option: byObject}   # option: byObject / byWord / byChar
```

**3D 相机还没有** —— 需要手写注入 `<a:scene3d>`（配方见上面第 2 节），
或等 spec 扩展。

注入后仍要走三道闸：
`motion.py apply --assert-geometry` → `verify_motion.py` → `motion.ps1 -Strict`。
**几何指纹不变**这条对 morph 同样成立：morph 只写 `<p:transition>`，
不动 `<p:spTree>` 的任何 `<a:xfrm>`。

回归测试：`tests/test_morph.py`（20 条断言，覆盖元素名/命名空间/降级/时长/
选项校验/置于 `<p:timing>` 之前/重复注入不累积/XML 合法性）。

---

## 6. 艺术效果（虚化 / 亮度）—— 也能通过 COM 做

template1 的背景不是一张模糊好的图，而是**图片 + 艺术效果**。它存在
`<a:blip>` 的扩展里：

```xml
<a:blip r:embed="rId2">                     <!-- 原始图 -->
  <a:extLst><a:ext uri="{BEBA8EAE-…}">
    <a14:imgProps><a14:imgLayer r:embed="rId3">   <!-- 渲染结果（JPEG XR） -->
      <a14:imgEffect><a14:artisticBlur radius="8"/></a14:imgEffect>
      <a14:imgEffect><a14:brightnessContrast bright="-40000"/></a14:imgEffect>
    </a14:imgLayer></a14:imgProps>
  </a:ext></a:extLst>
</a:blip>
```

要点：

- **效果参数在 XML 里**（`radius`、`bright`），但**渲染结果被另存**成一个
  `.wdp`（JPEG XR）文件，用 `.../2007/relationships/hdphoto` 关系引用。
  显示的是那个 `.wdp`，原始图保持不动。
- `bright="-40000"` 是**千分之一百分比**，即 −40%。
- PIL **读写不了 `.wdp`**（JPEG XR），所以想手写这一层很难。

### 但 COM 能做，不用手写

`Shape.Fill.PictureEffects` **是可用的**，PowerPoint 会自己生成
`a14:imgLayer` + `.wdp`：

```powershell
$fx = $shape.Fill.PictureEffects.Insert(2)        # 2 = msoEffectBlur
for ($i = 1; $i -le $fx.EffectParameters.Count; $i++) {
  $ep = $fx.EffectParameters.Item($i)
  if ($ep.Name -eq 'Radius') { $ep.Value = 8 }    # 对应 artisticBlur radius="8"
}
$fx2 = $shape.Fill.PictureEffects.Insert(3)       # 3 = msoEffectBrightnessContrast
for ($i = 1; $i -le $fx2.EffectParameters.Count; $i++) {
  $ep = $fx2.EffectParameters.Item($i)
  if ($ep.Name -eq 'Brightness') { $ep.Value = -0.4 }   # 对应 bright="-40000"
}
```

实测：写出的 XML 与 template1 **逐字节同构**（同样两个 `imgEffect`、
同样 `radius="8"` / `bright="-40000"`、同样 `hdphoto` 关系）。

`Insert` 的类型号：`2`=模糊，`3`=亮度/对比度，`4`=混凝土，`5`=粉笔素描。
参数用**名字**取，不要按下标写死。

### 复现时踩到的两个坑（都会让结果偏离）

1. **别拿 deck 里那张已压暗的图当底图。** template1 存的 `image1.png`
   平均 RGB 是 `(120,107,71)`，而原图 `template1.PNG` 是 `(209,191,133)` ——
   前者已经是 −40% 亮度版（`209×0.57≈119`）。再叠一次 −40% 会变成近黑。
   **要用原图，让效果只作用一次。**
2. **背景色要跟原件一致。** template1 的 slide **没有 `<p:bg>`**，继承版式的
   白底；如果自作主张设成黑底，导出会整体偏暗。

改掉这两点后，复现结果与原件**逐像素一致**（平均差 0.08–0.14，
差异像素 0.03%–0.05%，即渲染噪声级别）。

---

## 7. 只有原始照片时，能不能复现？（素材派生链）

`material/template1` 里只有 `template1.PNG` 是**原始素材**，另两张都是派生出来的：

| 文件 | 性质 | 判定依据 | 复现方式 |
| --- | --- | --- | --- |
| `template1.PNG` | **原始照片** | 726×406，0% 透明、0% 白 | 源头 |
| `template1.1.PNG` | **全屏截图** | **2560×1440**（屏幕分辨率），**69% 近白像素**（含 PowerPoint 界面） | **不需要**，见下 |
| `template1.2.png` | **人物抠图** | **77.5% 透明**，不透明区 130×294 | **`cv2.grabCut` 可重建** |

### 7.1 截图那张（1.1）**不需要复现**

description 第 4 步是"全屏截图"——那是个把图层压平的**土办法**，副作用是
**把 PowerPoint 自己的界面也截进去了**（2560×1440 加 69% 白）。它不是设计意图。

要做同样的效果，直接把**清晰照片本身**当作被旋转的图层即可 —— 这正是教程本意。
所以这一环不构成复现障碍，而且去掉截图后画面更干净。

### 7.2 人物那张（1.2）需要分割，但能做到

原始照片里直接抠人，用 OpenCV 的 `grabCut` 即可（无需 AI 模型）：

```python
rect = (int(0.25*w), int(0.25*h), int(0.50*w), int(0.75*h))   # 粗略框住人
mask = np.zeros((h, w), np.uint8)
cv2.grabCut(img, mask, rect, bgm, fgm, 8, cv2.GC_INIT_WITH_RECT)
fg = np.where((mask == cv2.GC_FGD) | (mask == cv2.GC_PR_FGD), 255, 0).astype(np.uint8)
# 形态学去噪 + 取最大连通域 + 按 bbox 裁剪，即可作为透明 PNG 使用
```

实测结果：bbox **143×305**，而参考抠图是 **130×294** —— 宽高比一致（0.47 vs 0.44），
视觉复核确认"人物完整、透明背景、无残留噪点"。

**注意**：`grabCut` 需要**矩形提示**，且对背景复杂/人物与背景同色的图会失效。
它是经典分割，不是语义分割；要求更高时应当换专门的人像抠图工具（或 AI 抠图）。

### 7.3 结论

**只有原始照片，可以完整复现**：

1. 背景 = 原始照片 + 艺术效果（`artisticBlur radius=8` + `brightness -40`，走 COM）
2. 被旋转的图层 = **原始照片本身**（替代那张截图）
3. 人物 = `grabCut` 抠图
4. 动效 = `cameras: {tilt: 290}` + `type: morph`

实测产出：两页均 `EntryEffect=0xF72`（morph）、`LAYER.camera=62`（透视），
几何不变；视觉复核为"模糊背景 + 悬浮人物 + 照片三维倾斜"。

**但要注意一个区别**：上一轮的**逐像素一致**（平均差 0.08）之所以能做到，
是因为用了那两张派生素材。**只用原图时，动效完全一致，画面内容必然不同**
—— 原件的旋转图层是一张含界面白底的截图，而我们从原图重建的图层是干净照片。
这是"复现手法"与"复现原始文件"的区别。
