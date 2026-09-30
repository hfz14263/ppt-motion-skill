#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""给长文档补一张目录（TOC）—— 只加导航，不动正文一个字。

**为什么需要**：`reference/` 里有几份 400–600 行的文档**一个目录都没有**。
它们结构其实很清楚（§一…§十、43/44/46…），但读者打开只看到第一屏，
**不知道里面有什么**，只能滚。这与"查它要翻屏"是同一个病的两种症状：
拆分管的是"文件太大"，目录管的是"文件不小但应该有地图"。

**为什么不拆这些**：它们是**内聚的层文档**（形态层 / 选择层 / 机制层），
拆开会把 L1–L4 的架构打散。内聚的长文档正确解法是**加目录**，不是切碎。

**沿用本仓库既有约定**（见 `transition-model.md`）：
  - 目录是**表格**，不是嵌套列表（一行一个模型/章节，带一句话摘要）
  - 每节标题后插一个**显式锚点** `<a id="sN"></a>`
    —— 中文标题的自动 slug 在不同渲染器下不一致，显式锚点才可靠

**幂等**：已有 `<!-- toc -->` 标记就替换那块，不重复插入。

用法：
    python tools/add_toc.py reference/transition-shapes.md [more...]
    python tools/add_toc.py --check reference/*.md     # 只报告谁缺目录
"""
import io
import os
import re
import sys

BEGIN = "<!-- toc -->"
END = "<!-- /toc -->"

# 一句话摘要的来源：标题本身可能就够；需要人工补的写在这里（按标题前缀匹配）

H2 = re.compile(r"^## (?!#)(.+)$", re.M)
ANCHOR = re.compile(r'^<a id="s\d+"></a>$', re.M)


def in_code_blocks_finder(text):
    """返回一个函数：判断某个位置是否落在 ``` 代码块里。"""
    spans = [(m.start(), m.end())
             for m in re.finditer(r"```.*?```", text, re.S)]
    return lambda p: any(a <= p < b for a, b in spans)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    check = "--check" in argv
    files = [a for a in argv if not a.startswith("--")]
    if not files:
        print("用法: python tools/add_toc.py <md> [<md>...] [--check]")
        return 2

    for rel in files:
        if not os.path.exists(rel):
            print("跳过（不存在）: %s" % rel)
            continue
        text = io.open(rel, encoding="utf-8").read()
        fname = os.path.basename(rel)
        inside = in_code_blocks_finder(text)

        # 标题层级自适应：优先用 H2；H2 太少说明这份文档是**一个 H2 装一叠 H3**
        # （拆出来的 pitfall-*/symptom-* 都是这种：`## 文件打不开` 下面挂一列
        #  `### §1 / §2 / …`）。那种情况只有 H3 能当目录，否则目录只有一行。
        h2 = [(m.start(), m.group(1).strip())
              for m in re.compile(r"^## (?!#)(.+)$", re.M).finditer(text)
              if not inside(m.start())]
        h3 = [(m.start(), m.group(1).strip())
              for m in re.compile(r"^### (?!#)(.+)$", re.M).finditer(text)
              if not inside(m.start())]
        if len(h2) >= 4:
            heads, level = h2, 2
        elif len(h3) >= 4:
            heads, level = h3, 3
        else:
            print("跳过（标题少于 4 个，不需要目录）: %s" % rel)
            continue

        # 三种情况要分清：
        #   marked=True  本工具生成的（有 <!-- toc -->）→ **重新生成**，便于改格式
        #   handwritten  人写的目录（无标记）        → **跳过**，别在好目录上叠一张
        #   neither      没有目录                    → 生成
        marked = BEGIN in text
        handwritten = False
        if not marked:
            head40 = "\n".join(text.splitlines()[:40])
            handwritten = bool(re.search(r"^\|.*\]\(#", head40, re.M))
        if check:
            state = "有(本工具)" if marked else ("有(手写)" if handwritten else "**缺**")
            print("%-40s H%d=%2d  目录=%s" % (rel, level, len(heads), state))
            continue
        if handwritten:
            print("跳过（已有手写目录，别叠）: %s" % rel)
            continue

        # 1) 补锚点：每个标题之后插入 <a id="sN"></a>（已存在则跳过）
        out = text
        rows = []
        for i, (pos, title) in enumerate(heads, 1):
            aid = "s%d" % i
            # 计算标题行结束位置
            line_end = out.index("\n", pos) + 1
            nxt = out[line_end:line_end + 80]
            if not re.match(r"\s*<a id=\"s\d+\"></a>", nxt):
                out = (out[:line_end] + "\n<a id=\"%s\"></a>\n" % aid
                       + out[line_end:])
            rows.append((title, aid))

        # 2) 造目录。**单列链接列表**，不配摘要 —— 标题本身就是摘要，
        #    再加一列只会把同一句话写两遍（第一版就是这样，读起来像口吃）。
        lines = []
        for title, aid in rows:
            lines.append("- [%s](#%s)" % (title, aid))
        toc = BEGIN + "\n" + "\n".join(lines) + "\n" + END + "\n"

        # 3) 插在**第一个标题之前**；若它前面不远处有 `---` 分隔线，就插在分隔线
        #    之前（和 transition-model.md 的手写目录位置一致：导言 → 目录 → 分隔线）
        if BEGIN in out:
            out = re.sub(re.escape(BEGIN) + r".*?" + re.escape(END) + r"\n",
                         toc, out, flags=re.S)
        else:
            marker = "\n" + "#" * level + " "
            first = out.index(marker)
            sep = out.rfind("\n---\n", 0, first)
            # 「不远处」= 两者之间不超过 25 行（导言 + 表格通常在这个量级）
            near = sep != -1 and out.count("\n", sep, first) <= 25
            if near:
                out = out[:sep + 1] + toc + out[sep + 1:]
            else:
                out = out[:first + 1] + "\n" + toc + "\n---\n" + out[first + 1:]

        io.open(rel, "w", encoding="utf-8", newline="\n").write(out)
        print("%-40s 已加目录（%d 节）" % (rel, len(rows)))

    return 0


if __name__ == "__main__":
    sys.exit(main())
