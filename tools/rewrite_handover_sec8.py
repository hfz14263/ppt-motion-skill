#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性：把 HANDOVER.md §8 重写成「还没做的 + 归档指针 + 验证文化」。

原本 §8 = 8 条（7 条已完成）+ 验证文化。已完成项已由
archive_handover_backlog.py 搬进 history/。这里只做 §8 的改写。

**不做的事**：不动 §1–§7 与 §9。
**保留的事**：item 2 原文、验证文化全文 —— 一个字不改。
"""
import io
import os
import sys
from datetime import date

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, "HANDOVER.md")
TMP = os.environ.get("TEMP") or "/tmp"

SEC8 = "## 8. 待办与建议方向"
SEC9 = "## 9. 来源与许可"

NEW8 = """## 8. 还没做的

**只有这一条是真·待办** —— 其余 7 条都已完成，归档在 [`history/`](history/README.md)。

2. **多 API 协作** —— 素材生成（尤其抠图/插画）接专门的图像生成 AI，
   替代 `grabCut` 这类兜底方案。

### 8.1 已经做完的 → 归档在 [`history/`](history/README.md)

这一节原本有 **299 行**（占全文 593 行的一半），8 条里 **7 条已打 ✅**。
它实际上是一份**开发日志**，不是待办清单 —— 而日志留在概览文档里，
代价由每个只想读概览的人支付。

> **日志没错，放错地方才是。** 所以搬进 `history/`：
> 想查历史的人去查，想读概览的人不被挡。

**搬走时逐条原文照搬，一个字没改** —— 那些记录的价值全在
「当时为什么想错了」的细节上，改写就等于毁证。
三份归档**按主题分**（不是按时间）：同一条线的工作横跨好几天，
按日排会让人来回跳。

| 归档 | 装什么 | 规模 |
| --- | --- | --- |
| [`history/early-capabilities.md`](history/early-capabilities.md) | 美感判据执行化 · `fills:` spec 扩展 · 按原理组织的配方库 | 40 行 |
| [`history/doc-restructure.md`](history/doc-restructure.md) | 三层文档结构重整（文档曾长到 27 文件 / 261k 字符） | 35 行 |
| [`history/transition-layers.md`](history/transition-layers.md) | 页面切换四层：数据表 → 机制 → 形态 → 选择 | 226 行 |

**这一节本身也是个教训的实例**：把日志和待办混在一起，
结果是**待办被淹没在历史里** —— 8 条里 7 条早就做完了，
但读者无法一眼看出"还剩什么"，只能逐条读 ✅。
**混在一起的两类东西，会一起失效。**

### 8.2 验证文化（请保持）

"""


def main():
    text = io.open(SRC, encoding="utf-8").read()
    s = text.index(SEC8)
    e = text.index(SEC9)

    culture = io.open(os.path.join(TMP, "handover_culture.txt"),
                      encoding="utf-8").read()
    # 原文第一行是加粗标题 `**验证文化（请保持）**`，现已提升为 8.2 的标题
    first_nl = culture.index("\n")
    assert "验证文化" in culture[:first_nl], culture[:80]
    culture_body = culture[first_nl:].lstrip("\n")

    new_sec = NEW8 + culture_body
    if not new_sec.endswith("\n"):
        new_sec += "\n"

    out = text[:s] + new_sec + "\n---\n\n" + text[e:]
    io.open(SRC, "w", encoding="utf-8", newline="\n").write(out)

    after = io.open(SRC, encoding="utf-8").read()
    print("HANDOVER.md: %d 行 -> %d 行" % (text.count("\n"), after.count("\n")))
    print("§8 现在 %d 行" % (after.count("\n") - out.count("\n") +
                             0 if False else
                             after[after.index(SEC8):after.index(SEC9)].count("\n")))
    return 0


if __name__ == "__main__":
    sys.exit(main())
