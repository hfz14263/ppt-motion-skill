#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 morph-and-3d-recipes.md 按主题切成 3 份 + 目录页。

**为什么拆**：683 行，是 reference/ 里最大的一份。更糟的是**文件名与内容不符** ——
名字说"Morph 与 3D 相机的注入配方"，实际还装着：
  §4 从 material 模板提取的手法    §6 艺术效果（虚化/亮度）
  §7 只有原始照片时能不能复现       §8 用形状切割图片（窗口化填充）
找"形状切割"的人**不会想到**去打开一个叫 morph-and-3d-recipes 的文件。这是
导航失败，不是篇幅问题。

**沿用踩坑手册那套纪律**（见 scripts/build_pitfall_map.py 的说明）：
  1. **原文件名保留成目录页** —— 21 处引用写着它，删掉就是 21 个死链。
  2. **§编号不变、不重排** —— `morph-and-3d-recipes.md §1.2`、
     `§2.3`、`§4`、`§6`、`§7`、`§8` 都在别处被引着。编号是接口。
  3. **全文原文照搬**，不改一个字（这些记录的价值在细节里）。

**§5 留在目录页**：它讲 morph / cameras / fills 三者与 spec 的关系，
是**横跨三份**的整合视图 —— 放进任何一份都是错位。
"""
import io
import os
import re
import sys
from collections import OrderedDict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
SRC = os.path.join(REF, "morph-and-3d-recipes.md")

GROUPS = OrderedDict([
    ("recipe-morph-camera.md", {
        "title": "配方 · Morph 与 3D 相机（注入）",
        "duty": "两项**声明式**能力：`transition: {type: morph}` 与 `cameras:`。"
                "写进去要多难？元素写法、跨页匹配规则、三条硬约束、实测对照。",
        "nums": [1, 2, 3],
    }),
    ("recipe-asset-derivation.md", {
        "title": "配方 · 素材派生（从参考模板到能用的素材）",
        "duty": "手上只有参考模板（甚至只有一张**原始照片**）时，怎么把素材做出来。"
                "含两次真实更正 —— 都是「我以为能凑，其实不能」。",
        "nums": [4, 6, 7],
    }),
    ("recipe-fill-window.md", {
        "title": "配方 · 窗口化图片填充（用形状切割图片）",
        "duty": "**单页内部**动画：让一张图片像被窗口「扫过」一样露出来。"
                "关键在「默认填充是拉伸的，要做窗口必须给负偏移」。",
        "nums": [8],
    }),
])

HEAD = """# {title}

> {duty}
>
> **`§` 编号沿用拆分前的原文**（`§1.2`、`§2.3`、`§8.3` 等，别处正引着它们）。
> 完整目录见 [`morph-and-3d-recipes.md`](morph-and-3d-recipes.md)。

---

"""

H2 = re.compile(r"^## (\d+)\. .*$", re.M)


def main():
    text = io.open(SRC, encoding="utf-8").read()
    marks = list(H2.finditer(text))
    if not marks:
        raise SystemExit("找不到 `## n.` 标题")
    preamble = text[:marks[0].start()]
    secs = OrderedDict()
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        secs[int(m.group(1))] = text[m.start():end].rstrip() + "\n"

    print("原文 §：%s" % sorted(secs))
    assigned = [n for _, s in GROUPS.items() for n in s["nums"]]
    if sorted(assigned + [5]) != sorted(secs):
        raise SystemExit("分区不完整：搬走 %s + 目录页留 [5] != 全部 %s"
                         % (assigned, sorted(secs)))
    print("分区完整 ✓（%d 节搬走 + §5 留目录页 = %d）\n"
          % (len(assigned), len(secs)))

    for fname, spec in GROUPS.items():
        body = "".join(secs[n] for n in sorted(spec["nums"]))
        out = HEAD.format(title=spec["title"], duty=spec["duty"]) + body
        p = os.path.join(REF, fname)
        io.open(p, "w", encoding="utf-8", newline="\n").write(out.rstrip() + "\n")
        print("  %-32s §%s  (%d 行)"
              % (fname, " §".join(str(n) for n in spec["nums"]),
                 out.count("\n")))

    # 目录页 + §5 的正文片段，给改写用
    tmp = os.environ.get("TEMP") or "/tmp"
    io.open(os.path.join(tmp, "morph_hub_sec5.txt"), "w",
            encoding="utf-8", newline="\n").write(secs[5])
    io.open(os.path.join(tmp, "morph_hub_preamble.txt"), "w",
            encoding="utf-8", newline="\n").write(preamble)
    print("\n§5 与前言已导出到 %s/morph_hub_*.txt（供改写目录页）" % tmp)
    return 0


if __name__ == "__main__":
    sys.exit(main())
