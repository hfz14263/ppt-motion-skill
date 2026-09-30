#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性：把 morph-and-3d-recipes.md 改写成目录页（并入 §5 原文）。"""
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HUB = os.path.join(ROOT, "reference", "morph-and-3d-recipes.md")
TMP = os.environ.get("TEMP") or "/tmp"

HEAD = """# 注入配方 + 素材派生 · 目录页

> **这一页是目录，不是正文。** 2026-09-30 它从一份 **683 行**的文件拆成了 3 份 ——
> 原因**不是篇幅，是文件名与内容不符**：名字只说了「Morph 与 3D 相机」，
> 实际还装着素材派生（§4/§6/§7）和窗口化填充（§8）。
> **找"形状切割"的人不会想到打开一个叫 morph-and-3d-recipes 的文件** ——
> 这是导航失败，光靠加目录治不了，得让文件名对上内容。
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
| **用形状切割图片**（单页内"扫过"） | [`recipe-fill-window.md`](recipe-fill-window.md) | §8 |

> **为什么按"想做的事"排而不按 § 号排**：你会带着一个动作来查，
> 不会带着一个编号来查。编号只在**别人引用给你**的时候有用 ——
> 所以它留着，但不做导航主线。

---

## §5 三者与 spec 的关系（留在本页）

这一节讲 morph / `cameras:` / `fills:` **三者共同**怎么落到 spec 上，
以及为什么它们仍然不动几何 —— **横跨三份**，放进任何一份都是错位，
所以留在目录页。

"""


def main():
    sec5 = io.open(os.path.join(TMP, "morph_hub_sec5.txt"),
                   encoding="utf-8").read().rstrip() + "\n"
    pre = io.open(os.path.join(TMP, "morph_hub_preamble.txt"),
                  encoding="utf-8").read()

    # 前言里的 H1 与旧的"三项技术配方"式引导语都不要了，只保留一句环境说明
    env = ""
    for line in pre.splitlines():
        if "环境" in line or "Office" in line:
            env = line.strip()
            break

    out = HEAD
    if env:
        out += "> " + env + "\n\n"
    out += sec5
    out += "\n---\n\n## 这三份的共同点\n\n"
    out += (
        "**① 三份都是「声明式」的** —— 能写进 `motion.yaml` 的 spec，就不该手写注入。\n"
        "手写只在 spec 还没覆盖时才做，做完就该补进 spec。\n\n"
        "**② 三份都经过「真开一遍 + 往返比对」** —— 不是读文档推出来的。\n"
        "§3 的两个测量陷阱就是给「我以为没生效」准备的：\n"
        "**「我的写法不生效」≠「这功能不支持」**，先让 PowerPoint 自己写一遍再判。\n\n"
        "**③ 都遵守「几何零改写」** —— 形状的 `<a:xfrm>` 一个字节不动。\n"
        "所以上游的门禁结论可以**直接继承**，能用几何指纹做机器自证。\n"
    )

    io.open(HUB, "w", encoding="utf-8", newline="\n").write(out.rstrip() + "\n")
    print("已改写 %s" % os.path.relpath(HUB, ROOT))
    print("新行数: %d" % out.count("\n"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
