#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 reference/pitfall-map.md —— §编号 → 文件的唯一映射。

为什么需要这个脚本（而不是手写一张表）：

    `com-pitfalls.md` 曾经是**一份**文件，§1..§51 顺序编号。全仓库有
    90+ 处引用写成 `com-pitfalls.md §44`。当这份文件涨到 2000 行、51 节时，
    它必须被拆成若干份（按主题），但**拆分不能让那 90+ 处引用失效** ——
    它们散落在 facts/、tests/、其它 reference/ 里，逐个改一遍既费力又
    必然漏掉几处。

    做法：拆完之后**保留 §编号**，编号的含义（§44 = 切换的形态）不变，
    只把「编号 → 文件」的映射集中到 `pitfall-map.md`。引用者仍可写
    `com-pitfalls.md §44` 并由映射表解析；新引用则直接写新文件名。

    ⚠️ 编号**只增不减、只挪不改**。一个编号永远指同一件事，哪怕它换了文件。
       这是这套引用能活下来的唯一原因 —— 编号一旦被复用，历史引用就会
       静默指向另一个坑，比链接失效更糟。

本脚本从拆分后的文件里**扫描** `## <n>. ` 标题，所以映射表永远是
拆分的真实结果，不是另抄一份（抄一份就会漂）。

用法：
    python scripts/build_pitfall_map.py            # 生成/更新映射表
    python scripts/build_pitfall_map.py --check    # 只校验，不写盘（CI 用）
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
OUT = os.path.join(REF, "pitfall-map.md")

# 拆分后的目标文件（按主题）。顺序即本文件的呈现顺序。
# 每项: (文件名, 一句话职责)
FILES = [
    ("pitfall-file-corruption.md", "文件打不开 / 报损坏 —— 结构非法的全部成因"),
    ("pitfall-silent-drop.md", "XML 里有、PowerPoint 不认 —— 静默丢弃与静默退回"),
    ("pitfall-roundtrip.md", "往返（round-trip）后被改掉或吃掉"),
    ("pitfall-com.md", "COM 与脚本环境 —— 界面层的约束与陷阱"),
    ("pitfall-ooxml.md", "OOXML 注入写法 —— 元素、属性、命名空间"),
    ("pitfall-measurement.md", "测量与判据 —— 指标盲区、单变量纪律"),
    ("pitfall-morph-3d.md", "Morph / 3D / 图片填充 —— 能力判定与写法"),
    ("pitfall-media.md", "素材与观感 —— 截图/照片、抠图、露底"),
    ("pitfall-transition.md", "页面切换 —— 机制、形态、方向、与动画的关系"),
    ("pitfall-tooling.md", "工具链 —— 推送、导出、视频、运行方式"),
]

SEC = re.compile(r"^## (\d+)\. (.+)$", re.M)


def scan():
    """{num: (file, title)} —— 从真实文件里扫，不手抄。"""
    found = {}
    dups = []
    for fname, _ in FILES:
        p = os.path.join(REF, fname)
        if not os.path.exists(p):
            continue
        for m in SEC.finditer(io.open(p, encoding="utf-8").read()):
            n = int(m.group(1))
            if n in found:
                dups.append((n, found[n][0], fname))
            found[n] = (fname, m.group(2).strip())
    return found, dups


def render(found):
    lines = [
        "# 踩坑编号 → 文件映射",
        "",
        "**这一页是 `§n` 这个写法的解释器。** 全仓库（`facts/`、`tests/`、"
        "各 reference）有 90+ 处引用写成 `com-pitfalls.md §44` 这种形式。"
        "踩坑手册按主题拆成多份之后，**编号不变、含义不变**，只是编号到文件的"
        "对应关系挪到了这里。",
        "",
        "> **编号只增不减、只挪不改。** 一个编号永远指同一件事，哪怕它换了文件。",
        "> 这是历史引用还能用的唯一原因 —— 编号一旦被复用，旧引用会**静默**指向",
        "> 另一个坑，比链接 404 更糟。",
        "",
        "**怎么用**：看到 `§44`，在本表查到它在哪份文件，再 `grep` 那个文件里的",
        "`## 44.`。新写引用时请**直接写文件名**（例如 "
        "[`pitfall-transition.md`](pitfall-transition.md) §44），不要再走编号。",
        "",
        "本页由 `scripts/build_pitfall_map.py` 从文件标题扫描生成 —— 不是手抄，"
        "所以不会漂。",
        "",
        "---",
        "",
        "| § | 标题 | 在哪个文件 |",
        "| --- | --- | --- |",
    ]
    for n in sorted(found):
        fname, title = found[n]
        t = title.replace("|", "\\|")
        lines.append("| §%d | %s | [`%s`](%s) |" % (n, t, fname, fname))
    lines += ["", "---", "", "## 文件职责", "", "| 文件 | 管什么 |", "| --- | --- |"]
    for fname, duty in FILES:
        if os.path.exists(os.path.join(REF, fname)):
            lines.append("| [`%s`](%s) | %s |" % (fname, fname, duty))
    lines.append("")
    return "\n".join(lines)


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    check = "--check" in argv
    found, dups = scan()
    if not found:
        print("没有扫到任何 `## n. ` 标题 —— 拆分还没做，或文件名变了")
        return 1
    if dups:
        for n, a, b in dups:
            print("FAIL 编号 §%d 同时出现在 %s 和 %s" % (n, a, b))
        return 1
    body = render(found)
    if check:
        old = io.open(OUT, encoding="utf-8").read() if os.path.exists(OUT) else ""
        if old != body:
            print("FAIL pitfall-map.md 与拆分结果不一致 —— 重跑不带 --check")
            return 1
        print("OK  映射表与 %d 个文件一致（%d 条编号）" % (len(FILES), len(found)))
        return 0
    io.open(OUT, "w", encoding="utf-8", newline="\n").write(body)
    gaps = [n for n in range(1, max(found) + 1) if n not in found]
    print("已写 %s：%d 条编号，跨 %d 份文件"
          % (os.path.relpath(OUT, ROOT), len(found),
             len({f for f, _ in found.values()})))
    if gaps:
        print("⚠️ 缺号（编号必须连续，除非是有意废弃）: %s" % gaps)
    return 0


if __name__ == "__main__":
    sys.exit(main())
