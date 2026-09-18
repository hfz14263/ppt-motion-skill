# Morph 与 3D 相机的注入配方

来源：`D:\idea\material` 的两份真实 PowerPoint 文件（`template1` / `template2`），
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
（`material/template2` 的两页几何完全相同 —— 那份文件不产生补间，它只是理念说明。）

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
| `template2` | 目录卡片展开 | **两页几何相同，实际不产生补间** —— 是一份理念说明：矩形铺满、计算等分宽度、左对齐挪出画板、两页加平滑 |

`template1` 的额外技巧（非 3D 部分）：

- 背景图先做**虚化**（半径 8–10，界面里的"模糊"可能叫"虚化"）+ **亮度 −40**；
- 再叠一张同尺寸原图；
- 人物单独**抠图**（`图片格式 → 删除背景`，或先用 AI 生成带白边的人物图更好抠）；
- 背景层加 3D 旋转做纵深，人物层放大做层次。

---

## 5. 和 spec 的关系

本 skill 的 motion spec 目前**没有** morph / 3D 字段 —— 这两项要么手写注入，
要么等 spec 扩展。注入后仍要走三道闸：
`motion.py apply --assert-geometry` → `verify_motion.py` → `motion.ps1 -Strict`。
**几何指纹不变**这条对 morph 同样成立：morph 只写 `<p:transition>`，
不动 `<p:spTree>` 的任何 `<a:xfrm>`。
