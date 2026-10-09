#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 CONTRIBUTING.md 按**读者是谁**拆成两份。

**为什么拆**：它涨到 19924 字符，离 20000 硬上限只剩 **76** ——
再写一段话就超。而它还在长（每轮都可能加规则）。
硬上限的意义是"到这儿就必须处理"，不是"忍一忍还能塞"。

**按什么拆**：按**读者要干什么**，不按篇幅：

    CONTRIBUTING.md   一~五 + 七    加文档 / 加目录 / 改文档的人
    CODE_RULES.md     六 + 八~十    改代码的人

    这两类读者本来就不重合 —— 写文档的人不需要读「依赖方向」，
    改代码的人不需要读「长文档该加目录还是拆」。

**两条硬纪律**（与前几次拆分同一套）：

  1. **编号沿用拆开前的原编号，不重排。** `## 六` 在新文件里还是 `## 六`，
     即使它前面没有一~五。理由：**编号是接口** ——
     `reference/code-*.md` 里写着 `§六.6.1`、`§六.2`、`§十.4` 这类引用，
     重排会让它们全部指错。每份文件开头都说明这一点。
  2. **逐字节搬运**：每个章节的内容拆前拆后必须逐字节相同，
     只允许改「章节之间的衔接」（那是拆分的必要代价，且可复核）。

用法：
    python tools/split_contributing.py --dry
    python tools/split_contributing.py --apply
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "CONTRIBUTING.md")
NEW = os.path.join(ROOT, "CODE_RULES.md")

# 章节 -> 去哪份文件（1-based 章号）
A, B = "CONTRIBUTING.md", "CODE_RULES.md"
CHAPTER_TO_FILE = {1: A, 2: A, 3: A, 4: A, 5: A, 6: B, 7: A, 8: B, 9: B, 10: B}

HEAD_A = """# 文档与结构规范

> **这份文件回答一个问题：往这个仓库里加文档 / 加目录时，该怎么加。**
> 它针对的是三个已经真实发生过的病症 ——
> **随意生成**（同一件事两处各写一份）、**无限堆积**（日志挤垮概览）、
> **结构混乱**（新目录没人知道、引用悄悄失效）。
>
> **它不是原则宣言，是可执行规则。** 每一条都有对应的机器检查，
> 见 [`scripts/check_structure.py`](scripts/check_structure.py) 与
> [`tests/test_docs.py`](tests/test_docs.py)。

**改代码的规则在 [`CODE_RULES.md`](CODE_RULES.md)** —— 动代码前读那一份。

## 章节索引

| 章 | 在 | 装什么 |
| --- | --- | --- |
| 一~五 | **本文件** | 统一结构 · 体积上限 · 分诊 · 命名与归档 · 变更流程 |
| 六 | [`CODE_RULES.md`](CODE_RULES.md) | 代码长期规则（拆分 / 依赖 / 内聚 / 临时代码……） |
| 七 | **本文件** | 反模式（明确禁止） |
| 八~十 | [`CODE_RULES.md`](CODE_RULES.md) | 怎么被强制 · 代码索引规范 · 补充规则 |

> **编号沿用拆开前的原编号，不连续是正常的 —— 别重排。**
> 理由见 [`CODE_RULES.md`](CODE_RULES.md) 开头：外部引用写着 `§六.6.1` 这类编号。

---

"""

HEAD_B = """# 代码规范

> **这份文件回答一个问题：改这个仓库的代码时，该怎么改。**
>
> 它是 [`CONTRIBUTING.md`](CONTRIBUTING.md)（文档与结构规范）的**代码篇** ——
> 2026-10-09 拆出来，因为合在一起已经涨到 19924 字符、离硬上限只剩 76。
>
> 配套的两份材料：
> - [`CODE_INDEX.md`](CODE_INDEX.md) —— 代码地图（「我要做 X，改哪个文件」）
> - [`scripts/check_structure.py`](scripts/check_structure.py) —— 本规范的**执行者**
>
> ⚠️ **编号沿用拆开前的原编号（一~五、七 在 [`CONTRIBUTING.md`](CONTRIBUTING.md)），
> 不连续是正常的 —— 别重排。**
> `reference/code-*.md` 里写着 `§六.2`、`§六.6.1`、`§十.4` 这类引用，
> 重排会让它们**静默指向错误的内容**（比 404 更难发现）。
>
> 目录由 `python tools/add_toc.py CODE_RULES.md` 生成（**别手抄**）：
> 中文标题的自动 slug 在不同渲染器下不一致，所以每节标题带**显式锚点**。

---

"""


def split_chapters(text):
    """按 `## <中文数字>、` 切章；返回 (头部, [(章号, 原文)])"""
    marks = [(m.start(), m.group(0))
             for m in re.finditer(r"^## [一二三四五六七八九十]+、.*$", text, re.M)]
    if not marks:
        raise SystemExit("找不到 `## 一、` 这类章节标题")
    head = text[:marks[0][0]]
    out = []
    for i, (pos, title) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        cn = re.match(r"## ([一二三四五六七八九十]+)、", title).group(1)
        num = "一二三四五六七八九十".index(cn) + 1
        out.append((num, text[pos:end].rstrip() + "\n"))
    return head, out


def main():
    apply_ = "--apply" in sys.argv
    text = io.open(SRC, encoding="utf-8").read()
    head, chapters = split_chapters(text)
    print("源文件 %d 字符，%d 章" % (len(text), len(chapters)))
    print()

    bucket = {A: [], B: []}
    for num, body in chapters:
        bucket[CHAPTER_TO_FILE[num]].append((num, body))

    for f, items in bucket.items():
        n = sum(len(b) for _num, b in items)
        print("  %-18s %d 章 %6d 字符  （章：%s）"
              % (f, len(items), n,
                 "、".join("一二三四五六七八九十"[num - 1] for num, _ in items)))
    print()

    if not apply_:
        print("（--dry：没有写任何文件）")
        return 0

    # A：原文件，换头
    a_body = "\n\n".join(b.rstrip() for _num, b in bucket[A])
    io.open(SRC, "w", encoding="utf-8", newline="\n").write(
        HEAD_A + a_body.rstrip() + "\n")
    # B：新文件
    b_body = "\n\n".join(b.rstrip() for _num, b in bucket[B])
    io.open(NEW, "w", encoding="utf-8", newline="\n").write(
        HEAD_B + b_body.rstrip() + "\n")

    for f, p in ((A, SRC), (B, NEW)):
        n = len(io.open(p, encoding="utf-8").read())
        print("  写入 %-18s %6d 字符" % (f, n))
    print()
    print("接下来（本脚本不自动做）：")
    print("  · 改带章节号的外部引用（§六.x / §十.x 现在在 CODE_RULES.md）")
    print("  · install.ps1 的复制清单加 CODE_RULES.md")
    print("  · 跑 scripts/check_structure.py 与 tests/test_install_manifest.py")
    return 0


if __name__ == "__main__":
    sys.exit(main())
