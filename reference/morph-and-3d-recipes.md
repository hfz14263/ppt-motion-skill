# 注入配方 + 素材派生 · 目录页

> **这一页是目录，不是正文。** 2026-09-30 它从一份 **683 行**的文件拆成了 3 份 ——
> 原因**不是篇幅，是文件名与内容不符**：名字只说了「Morph 与 3D 相机」，
> 实际还装着素材派生（§4 / §6 / §7）和窗口化填充（§8）。
> **找「形状切割」的人不会想到打开一个叫 morph-and-3d-recipes 的文件** ——
> 这是导航失败，光加目录治不了，得让文件名对上内容。
>
> **`§` 编号沿用原文、不重排** —— `§1.2`、`§2.3`、`§4`、`§6`、`§7`、`§8`
> 都在别处被引着（`examples/material/README.md`、`SKILL.md`、
> `recipe-library.md`、`pitfall-*` …）。**编号是接口。**

## 三份配方

| 想做的事 | 去哪 | 原 § |
| --- | --- | --- |
| 两页之间**同一个形状平滑变形**（Morph） | [`recipe-morph-camera.md`](recipe-morph-camera.md) | §1 |
| 加 **3D 相机 / 角度**（`cameras:`） | [`recipe-morph-camera.md`](recipe-morph-camera.md) | §2 |
| 「我写了，但看起来**没生效**」—— 两个测量陷阱 | [`recipe-morph-camera.md`](recipe-morph-camera.md) | §3 |
| 从**参考模板**里把素材抠出来 | [`recipe-asset-derivation.md`](recipe-asset-derivation.md) | §4 |
| **艺术效果**（虚化 / 亮度） | [`recipe-asset-derivation.md`](recipe-asset-derivation.md) | §6 |
| 只有一张**原始照片**，能不能复现 | [`recipe-asset-derivation.md`](recipe-asset-derivation.md) | §7 |
| **用形状切割图片**（单页内「扫过」） | [`recipe-fill-window.md`](recipe-fill-window.md) | §8 |

**这张表就是 `§n` 的解释器** —— 别处写 `morph-and-3d-recipes.md §8` 时，
在这里查到 `§8 → recipe-fill-window.md`，再进那份文件找 `## 8.`。

> **为什么按「想做的事」排而不按 § 号排**：你会带着一个动作来查，
> 不会带着一个编号来查。编号只在**别人引用给你**的时候有用 ——
> 所以它留着，但不做导航主线。

---

## 5. 和 spec 的关系

**Morph 和 3D 相机都已纳入 spec。** 两者都是声明式的，不需要手写注入：

```yaml
transition: {type: morph, duration: 2.0, option: byObject}   # option: byObject / byWord / byChar

cameras:
  - {target: LAYER, tilt: 290}                                 # 半自动：只给角度
  - {target: HERO, prst: perspectiveRelaxed, lat: -70, lon: 10} # 全手动
```

**3D 相机已纳入 spec**（`cameras:`，度数为单位）；**窗口化图片填充已纳入 spec**
（`fills:`，见 §8.3）。三者都是**形状属性**，与 `effects` 平级但各自独立。

注入后仍要走三道闸：
`motion.py apply --assert-geometry` → `verify_motion.py` → `motion.ps1 -Strict`。
**几何指纹不变**这条对它们同样成立：morph 只写 `<p:transition>`；
`fills` 只改 `<p:spPr>` 的**填充**（形状的 `<a:xfrm>` 不碰 —— 窗口是"露出图片哪一部分"，
不是"改变形状位置"），所以上游门禁结论仍可继承。

回归测试：`tests/test_morph.py`（20 条）、`tests/test_camera.py`（30 条）、
`tests/test_fill_window.py`（28 条）。

> **§5 为什么留在目录页**：它讲 morph / `cameras:` / `fills:` **三者共同**
> 怎么落到 spec 上、为什么都不动几何 —— **横跨三份**，
> 放进任何一份都是错位。

---

## 这三份的共同点

**① 三份都是「声明式」的** —— 能写进 `motion.yaml` 的 spec，就不该手写注入。
手写只在 spec 还没覆盖时才做，做完就该补进 spec。

**② 三份都经过「真开一遍 + 往返比对」** —— 不是读文档推出来的。
§3 的两个测量陷阱就是给「我以为没生效」准备的：
**「我的写法不生效」≠「这功能不支持」**，先让 PowerPoint 自己写一遍再判。

**③ 都遵守「几何零改写」** —— 形状的 `<a:xfrm>` 一个字节不动。
所以上游的门禁结论可以**直接继承**，能用几何指纹做机器自证。
