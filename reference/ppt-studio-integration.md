# 与 dsh-ppt-studio 的集成契约

本 skill 对**任意 .pptx** 都成立；装上 `dsh-ppt-studio` 后体验更顺，原因是两者的坐标与锚点天然对齐。

## 1. 锚点：elementId → cNvPr/@name

`lib/pptd/export-pptx.js` 每个形状都这样写：

```js
const id = nid()
'<p:sp><p:nvSpPr><p:cNvPr id="' + id + '" name="' + xm(el.id) + '"/>…'
```

即 `<p:cNvPr id="1002" name="title"/>`——**deck.yaml 的 `elementId` 原样成为 pptx 形状 name**。

`nid()` 定义：

```js
let UID = 1
const nid = () => UID++
// exportPptx() 开头：UID = 1000
```

→ id 从 **1000** 起、跨页全局递增（slide2 接着 slide1 往下排），页内唯一。

因此 `motion.py` 的 `target` 直接写 deck.yaml 里的 `elementId` 即可：

```yaml
- {target: title, effect: fly, duration: 0.7}     # -> <p:spTgt spid="1002"/>
```

实测（examples/fx/fx-pro.pptx）：

```
slide 1: id=1000 band  id=1001 glow  id=1002 title  id=1003 sub  id=1004 badge  id=1005 badge_t
```

## 2. 坐标：完全一致，无需换算

| 项 | dsh-ppt-studio（`lib/pptd/layout.js` / `export-pptx.js`） | Office COM |
| --- | --- | --- |
| 画布 | 960×540（deck.yaml `size`） | `PageSetup.SlideWidth/Height` = 960×540 pt |
| 单位 | 1px = 1pt，EMU = pt × 12700 | 1pt，内部 EMU |
| 原点 | 左上 | 左上 |

所以 `layout.json` 里的 `bounds: [x, y, w, h]` 可以直接用作 COM 的 `AddMediaObject2` 参数。

## 3. 图表：一个 chart 拆成多个形状

`export-pptx.js` 里 chart 是矢量拼绘，每个数据点/连线都是独立形状，name 形如
`chart-d0-0`、`chart-d1-0`、`chart-L1`：

```
slide 4: id=1018 c_head  id=1022 chart-d0-0  id=1023 chart-d1-0
         id=1024 chart-d2-0  id=1025 chart-d3-0  id=1019 chart-L1 …
```

`resolve_targets()` 支持三种写法：

- `target: chart-d0-0` —— 精确命中单个柱
- `target: chart` —— 命中所有 `chart-p*` 前缀片（本机 pptd 用的是 `-dN-M`，故建议写全名）
- `target: [chart-d0-0, chart-d1-0]` —— 列表

柱状图逐条生长：

```yaml
- {target: "chart-d0-0", effect: stretch, duration: 0.5, trigger: after}
- {target: "chart-d1-0", effect: stretch, duration: 0.5, trigger: after, delay: 0.15}
```

## 4. 门禁结论可以继承（最关键的收益）

`ppt_verify` 判的是重叠/出界/溢出，依据是 spTree 的形状几何。
本 skill **只写 `<p:timing>` / `<p:transition>` / `[Content_Types].xml`**，几何一律不动，
并由 `--assert-geometry` 逐页给出 SHA256 证明：

```
slides=12 effects=31 transitions=4 media=0
geometry: UNCHANGED over 12 slides
```

→ **上游跑完 `ppt_verify` 后，动效化不需要重跑门禁**。这是把动效层做成 OOXML 补丁、
而不是"再开 PowerPoint 存一次"的全部理由。

## 5. 与上游工作流的拼接

```
dsh-ppt-studio:  S0 规格 → S1 大纲 → S2 定调 → S3 逐页 → S4 页审 → S5 整体审
                 ppt_render → ppt_verify（ERROR 清零）→ ppt_shot 视觉审阅
                 ppt_export  → deck.pptx
        ↓
本 skill:        motion.py inspect           # 拿 elementId 清单
                 motion.py apply --assert-geometry
                 verify_motion.py --source
                 motion.ps1 -Spec（媒体 + 真渲染复核）
                 read_image 视觉审阅
        ↓
                 animated.pptx（版面与 deck.pptx 逐页几何一致）
```

媒体建议放 `media/`（deck.yaml 工程的既有约定），spec 里写相对路径。

## 6. 已知边界

- 上游 README 明确"导入不支持动画"：`ppt_import` 读别人的 pptx 时不会带出动画。
  本 skill 是**从零注入**，不依赖导入。
- 上游 `ppt_splice` / `ppt_slice` 做的是"整页替换/单页裁剪"，会**保留**该页原有 timing；
  若先用本 skill 注入再 splice，timing 会随之被搬运（同一页条目保持）。混用前建议先
  `motion.py preview` 存一份静态版备用。
- 若上游以后改了 `cNvPr/@name` 的写法（不再等于 elementId），本 skill 仍可用数字 id 或
  真实 name 作为 target，不至于失效。
