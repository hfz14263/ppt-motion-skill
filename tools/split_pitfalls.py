#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性的拆分器：把 com-pitfalls.md 按主题切成多份，编号保持不变。

**这是一次性脚本**（放在 scripts/ 之外，见 tools/），拆完即完成使命。
真正的长期工具是 build_pitfall_map.py（映射表必须能重建）。

设计要点：
  1. **编号不动**。§n 的含义跨文件不变。文件内按编号升序排，便于 grep。
  2. 每份新文件自带一小段头部：这一份管什么、编号范围、指向映射表与索引。
  3. 原文件退化成**导航页**（保留全名 com-pitfalls.md），不是删除 —— 90+ 处
     引用写着这个名字，删掉就是 90+ 个死链。
  4. 切口必须落在 `## ` 之前，正文里的子标题 `### ` 不动。
"""
import io
import os
import re
import sys
from collections import OrderedDict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
SRC = os.path.join(REF, "com-pitfalls.md")

# ---- 分区表：编号 -> 目标文件 -------------------------------------------
# 依据是"看这一节时你想解决的是哪一类问题"，不是发现顺序。
# 一组里的编号不必连续（§31 讲损坏，但它排在文中 §31 的位置）。
GROUPS = OrderedDict([
    ("pitfall-file-corruption.md", {
        "duty": "文件打不开 / 报损坏 —— **结构非法**的全部成因",
        "nums": [1, 2, 4, 12, 26, 31, 40, 42],
    }),
    ("pitfall-silent-drop.md", {
        "duty": "XML 里有、PowerPoint 不认 —— 静默丢弃与静默退回",
        "nums": [13, 19, 41, 49, 50],
    }),
    ("pitfall-roundtrip.md", {
        "duty": "往返（round-trip）后被改掉或吃掉的效果",
        "nums": [14, 18],
    }),
    ("pitfall-ooxml.md", {
        "duty": "OOXML 注入写法 —— 元素顺序、属性、命名空间",
        "nums": [3, 8],
    }),
    ("pitfall-com.md", {
        "duty": "COM 与脚本环境 —— 界面层的约束、编码、真渲染",
        "nums": [5, 6, 7, 9, 10, 11, 15, 16, 17],
    }),
    ("pitfall-morph-3d.md", {
        "duty": "Morph / 3D 相机 / 图片填充 —— 能力判定与写法",
        "nums": [20, 21, 22, 23, 24, 25, 27, 29, 30, 34],
    }),
    ("pitfall-measurement.md", {
        "duty": "测量与判据 —— 指标盲区、单变量纪律、数字口径",
        "nums": [28, 33, 38],
    }),
    ("pitfall-media.md", {
        "duty": "素材与观感 —— 截图/照片、抠图、露出的底",
        "nums": [35, 36, 37, 39],
    }),
    ("pitfall-transition.md", {
        "duty": "页面切换 —— 机制、形态、方向、与页内动画的关系",
        "nums": [43, 44, 46, 47, 48, 51],
    }),
    ("pitfall-tooling.md", {
        "duty": "工具链 —— 推送通道、导出、视频、运行方式",
        "nums": [32, 45],
    }),
])

# 每份新文件的"先读"提示（可选）
POINTERS = {
    "pitfall-file-corruption.md":
        "**一条规则能防掉这里大部分**：序列型元素只能出现一次 —— 存在就替换，"
        "不存在才插入（§31）。",
    "pitfall-silent-drop.md":
        "这一类的共同点是**不报错**：XML 合法、结构自检通过，PowerPoint 只是安静地"
        "把你的东西丢掉或退回默认。判据永远不是「校验通过」，而是「真开一遍看效果在不在」。",
    "pitfall-roundtrip.md":
        "往返是「PowerPoint 帮我重写一遍文件」。它能救活你写不对的东西，也会顺手"
        "改掉你写对的东西 —— 两个方向都会咬人。",
    "pitfall-transition.md":
        "机制在 [`transition-model.md`](transition-model.md)，"
        "形态在 [`transition-shapes.md`](transition-shapes.md)，"
        "选型在 [`transition-choice.md`](transition-choice.md)。"
        "这一份只收「踩过坑」的那部分。",
}

HEAD = """# {title}

> {duty}
>
> **编号沿用拆分前的 `§n`**（`§44` 仍是"切换的形态"，哪怕它换了文件）。
> 完整编号表见 [`pitfall-map.md`](pitfall-map.md)。
> 按现象查见 [`symptoms.md`](symptoms.md)；入口见 [`../INDEX.md`](../INDEX.md)。

{pointer}---

"""

SEC = re.compile(r"^## (\d+)\. .*$", re.M)


def parse():
    text = io.open(SRC, encoding="utf-8").read()
    # 去掉原文件的 H1 与引言，正文从第一个 ## 开始
    marks = list(SEC.finditer(text))
    if not marks:
        raise SystemExit("找不到任何 `## n. ` 标题")
    preamble = text[:marks[0].start()]
    secs = OrderedDict()
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        secs[int(m.group(1))] = text[m.start():end].rstrip() + "\n"
    return preamble, secs


def main():
    preamble, secs = parse()
    for fname, spec in GROUPS.items():
        missing = [n for n in spec["nums"] if n not in secs]
        if missing:
            raise SystemExit("%s 要的编号不存在: %s" % (fname, missing))

    assigned = [n for spec in GROUPS.values() for n in spec["nums"]]
    if sorted(assigned) != sorted(secs):
        only_src = sorted(set(secs) - set(assigned))
        only_map = sorted(set(assigned) - set(secs))
        raise SystemExit("分区表与正文不一致\n  漏掉: %s\n  多出: %s"
                         % (only_src, only_map))

    titles = {
        "pitfall-file-corruption.md": "踩坑 · 文件打不开 / 报损坏",
        "pitfall-silent-drop.md": "踩坑 · XML 里有、PowerPoint 不认",
        "pitfall-roundtrip.md": "踩坑 · 往返后被改掉或吃掉",
        "pitfall-ooxml.md": "踩坑 · OOXML 注入写法",
        "pitfall-com.md": "踩坑 · COM 与脚本环境",
        "pitfall-morph-3d.md": "踩坑 · Morph / 3D / 图片填充",
        "pitfall-measurement.md": "踩坑 · 测量与判据",
        "pitfall-media.md": "踩坑 · 素材与观感",
        "pitfall-transition.md": "踩坑 · 页面切换",
        "pitfall-tooling.md": "踩坑 · 工具链",
    }

    for fname, spec in GROUPS.items():
        nums = sorted(spec["nums"])
        ptr = POINTERS.get(fname, "")
        if ptr:
            ptr = ptr + "\n\n"
        body = "".join(secs[n] for n in nums)
        out = HEAD.format(title=titles[fname], duty=spec["duty"], pointer=ptr) + body
        p = os.path.join(REF, fname)
        io.open(p, "w", encoding="utf-8", newline="\n").write(out.rstrip() + "\n")
        print("  %-38s §%s" % (fname, " §".join(str(n) for n in nums)))

    return 0


if __name__ == "__main__":
    sys.exit(main())
