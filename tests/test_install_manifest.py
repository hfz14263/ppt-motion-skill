#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression for install.ps1's copy list: it must ship the whole skill.

为什么需要它
------------
2026-09-28 把 L3 目录统一为单数（旧名是复数拼写）之后复验时发现：
install.ps1 的复制清单里仍写着复数的旧目录名，而三层结构（fac93ec）新增的两个
顶层项 —— `INDEX.md`（唯一入口）和 `facts/`（唯一真相源）—— 从来没被加进清单。

后果：装出来的技能包**没有入口页、也没有权威事实**，而 SKILL.md 和文档里的引用
全是好的，所以任何文档级校验都查不出来。只有真跑一次安装、对照目录，才会发现。

清单本身是可解析的，所以这件事不需要跑 PowerShell 也能守住：
  1. 清单里每一项都必须真实存在（防止挂着已改名的旧目录）；
  2. 顶层每一项（除开发期目录）都必须被清单覆盖（防止新增顶层项忘了登记）。

第 2 条是关键 —— 第 1 条只能防止"清单里有垃圾"，防不住"清单漏了东西"。
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# 开发期才需要、明确不该随技能包分发的目录。install.ps1 也会主动删掉 tests。
#
# `tools/` 是一次性搬迁/重构脚本的落脚处（例如 split_pitfalls.py 把踩坑手册
# 切分成 10 份）。它们对**使用者**没有意义 —— 装出去只会让人以为要运行它们。
# 与之相对，`scripts/` 是运行时工具，必须随包分发。这条界线由本集合表达。
DEV_ONLY = {"tests", "tools", "__pycache__", "probe-createvideo", ".workbuddy"}


def install_items():
    """Parse the `foreach ($item in 'a', 'b', ...)` list out of install.ps1."""
    src = io.open(os.path.join(ROOT, "install.ps1"), encoding="utf-8").read()
    m = re.search(r"foreach \(\$item in(.*?)\)\s*\{", src, re.S)
    if not m:
        raise AssertionError("install.ps1 里找不到复制清单 —— 是不是改写了 foreach？")
    return re.findall(r"'([^']+)'", m.group(1))


def main():
    fails = []
    items = install_items()

    if not items:
        fails.append("复制清单是空的")

    # 1. every listed item must exist
    for i in items:
        if not os.path.exists(os.path.join(ROOT, i)):
            fails.append("清单里的 %r 在仓库里不存在" % i)

    # 2. every shipped top-level entry must be listed
    top = sorted(e for e in os.listdir(ROOT)
                 if not e.startswith(".") and e not in DEV_ONLY)
    for e in top:
        if e not in items:
            fails.append("顶层项 %r 没有进复制清单 —— 装出去会缺它" % e)

    # a skill dir is any folder containing SKILL.md; the entry point must ship
    if "SKILL.md" not in items:
        fails.append("SKILL.md 不在清单里 —— 装出去就不算一个 directory skill")
    if "INDEX.md" not in items:
        fails.append("INDEX.md 不在清单里 —— 三层结构的唯一入口会丢")

    print("== install.ps1 复制清单 ==")
    for i in items:
        print("   %-22s %s" % (i, "ok" if os.path.exists(os.path.join(ROOT, i))
                               else "MISSING"))
    print()
    if fails:
        print("-- FAIL --")
        for f in fails:
            print("   " + f)
        return 1
    print("install manifest ok  (%d 项，顶层 %d 项全部覆盖)" % (len(items), len(top)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
