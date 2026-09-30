#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性归档器：把 HANDOVER.md §8 里"已完成项的记录"搬进 history/。

**为什么做这件事**：HANDOVER.md 是项目的**长期概览**（这是什么、架构、
目录、误判史、环境、许可）—— 它应该稳定、可通读。
但 §8「待办与建议方向」占了 593 行里的 **299 行**，而且其中 8 项有 7 项
已经打上 ✅ —— 它实际是一份**会无限增长的开发日志**，把概览挤成了日志。

日志不是坏的，**放错地方才是**：
  - 留在 HANDOVER → 每次要读概览的人先翻 300 行历史
  - 搬进 history/ → 想查历史的人去查，想读概览的人不被挡

**留下什么**：
  - item 2（多 API 协作）**真的还没做** → 留在 HANDOVER
  - 「验证文化（请保持）」是**常驻规则**，不是历史 → 留在 HANDOVER

**搬运纪律**：逐条**原文照搬**，不改一个字 —— 这些记录里有大量
"当时为什么想错了"的细节，那正是它们的价值；改写就等于毁证。

用法：
    python tools/archive_handover_backlog.py          # 执行
    python tools/archive_handover_backlog.py --check  # 只校验边界
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "HANDOVER.md")
HIST = os.path.join(ROOT, "history")

SEC8 = "## 8. 待办与建议方向"
SEC9 = "## 9. 来源与许可"

# item 编号 -> 目标文件。留下不搬的写 None。
GROUPS = [
    ("early-capabilities.md", {
        "title": "早期能力建设（美感判据 / spec 扩展 / 配方库）",
        "items": [1, 3, 4],
    }),
    ("doc-restructure.md", {
        "title": "三层文档结构重整",
        "items": [5],
    }),
    ("transition-layers.md", {
        "title": "页面切换四层（数据表 → 机制 → 形态 → 选择）",
        "items": [6, 7, 8],
    }),
]

HEAD = """# {title}

> 本文件是**归档**，不是待办。原属 `HANDOVER.md` §8「待办与建议方向」，
> 2026-09-30 因该节涨到 299 行（占全文一半）而搬出。
> 索引见 [`README.md`](README.md)；项目概览见 [`../HANDOVER.md`](../HANDOVER.md)。

**为什么这些内容还留着**：每一条都带着"当时为什么想错了"的细节 ——
那是它们唯一的价值，也是下次不重犯的依据。**改写就等于毁证。**

---

"""


def split_items(sec):
    """把 §8 切成 {编号: 正文}，外加 intro / 验证文化。"""
    marks = []
    for m in re.finditer(r"^(\d+)\. ", sec, re.M):
        marks.append((m.start(), ("item", int(m.group(1)))))
    vc = sec.find("**验证文化（请保持）**")
    if vc >= 0:
        marks.append((vc, ("culture",)))
    marks.sort()
    if not marks:
        raise SystemExit("§8 里找不到任何顶层编号项")
    out = {"intro": sec[:marks[0][0]]}
    for i, (pos, tag) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(sec)
        key = "item%d" % tag[1] if tag[0] == "item" else "culture"
        out[key] = sec[pos:end].rstrip() + "\n"
    return out


def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    check = "--check" in argv

    text = io.open(SRC, encoding="utf-8").read()
    s = text.index(SEC8)
    e = text.index(SEC9)
    sec = text[s:e]
    parts = split_items(sec)

    nums = sorted(int(k[4:]) for k in parts if k.startswith("item"))
    print("§8 共 %d 个条目: %s" % (len(nums), nums))
    print("§8 总行数: %d" % sec.count("\n"))

    assigned = [n for _, spec in GROUPS for n in spec["items"]]
    open_nums = [n for n in nums if n not in assigned]
    print("留在 HANDOVER 的（未完成）: %s" % open_nums)
    if sorted(assigned + open_nums) != nums:
        raise SystemExit("分区不完整：搬走 %s + 留下 %s != 全部 %s"
                         % (assigned, open_nums, nums))
    print("分区完整 ✓（搬走 %d + 留下 %d = %d）"
          % (len(assigned), len(open_nums), len(nums)))
    if check:
        return 0

    os.makedirs(HIST, exist_ok=True)
    for fname, spec in GROUPS:
        body = "".join(parts["item%d" % n] for n in sorted(spec["items"]))
        out = HEAD.format(title=spec["title"]) + body
        p = os.path.join(HIST, fname)
        io.open(p, "w", encoding="utf-8", newline="\n").write(out.rstrip() + "\n")
        print("  写入 %-28s §item %s" % (fname, spec["items"]))

    # 中间件写到系统临时目录，**不要写进 history/** —— 归档目录里只该有
    # 真正要长期保留的东西，混进两个 .txt 会让人以为它们也是文档。
    tmp = os.environ.get("TEMP") or "/tmp"
    io.open(os.path.join(tmp, "handover_culture.txt"), "w",
            encoding="utf-8", newline="\n").write(parts["culture"])
    io.open(os.path.join(tmp, "handover_open_items.txt"), "w",
            encoding="utf-8", newline="\n").write(
                "".join(parts["item%d" % n] for n in open_nums))
    print()
    print("已写出 history/ 。中间件在 %s/handover_{culture,open_items}.txt" % tmp)
    return 0

if __name__ == "__main__":
    sys.exit(main())
