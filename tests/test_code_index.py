#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""代码索引守护：CODE_INDEX.md 与四份子索引必须与真实代码同步。

为什么需要这份测试：索引腐烂是**静默**的。
改了代码不更新索引，文档不会报错、测试不会失败 —— 只是**安静地指错方向**。
人查索引时已经错了，所以没人知道索引已经烂了半年。

这个测试的全部意义就是：**让索引腐烂有人被看见。**

它守四件机械可判的事（不判断索引写得好不好，只判断它**还对不对**）：
  ① 每个 `.py`/`.ps1` 都被某份索引登记        —— 漏登记 = 存在没有记录的文件
  ② 索引里的行数与磁盘一致                 —— 数字过期是最常见的腐烂
  ③ 索引里的依赖与 `scan_deps.build_graph()` 一致
  ④ 分层数字（模块总数、环数）与实测一致
  ⑤ 子索引链回父层，总览链到全部子索引

> **2026-10-08 新建。** 配套 `CODE_INDEX.md` + `reference/code-*.md`。
> 表格数据由 `tools/gen_code_tables.py` 扫出来，本测试**核对**而不是生成 ——
> 生成器负责"写对"，本测试负责"发现写错"。
"""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from helpers import HERE, ROOT, read_text, make_checker, report

sys.path.insert(0, os.path.join(ROOT, "tools"))
import gen_code_tables as G
import scan_deps

TITLE = "代码索引同步"


def doc_lines(rel):
    """文档里写的行数：任何出现 `` `path` `` 且同行带 `N 行` 的表格行。

    实现要点（踩过的坑）：
    第一版用「恰好 6 个 `|`」来定位体积列，结果**连 `motion.py` 都解析不出来**
    —— 因为表格列数会随文档变化，靠数竖线太脆。
      **一次"度量工具出错 → 结论反向"的现场**：
      31 个文件全被判"未登记"，看起来像索引全烂了，其实是我的正则不对。
    正确做法：**只找「同一行里出现的 `N 行`」**，不关心它在哪一列。
    """
    full = os.path.join(ROOT, rel)
    if not os.path.exists(full):
        return {}
    out = {}
    for ln in read_text(full).split("\n"):
        m = re.match(r"^\|\s*`([\w/.\-]+\.(?:py|ps1))`", ln)
        if not m:
            continue
        # 必须取**最后一列**。第二版用 `(\d+)\s*行` 全行搜索，抓到了
        # 职责栏里的「把 build_transition_table.py（3048 行）拆成…」
        # —— 把被拆对象的体积当成了这个文件自己的体积。
        # 判据：切最后一个 `|` 之后的内容，只在那里找行数。
        tail = ln.rsplit("|", 2)[-2] if ln.count("|") >= 2 else ""
        n = re.search(r"(\d+)\s*行", tail)
        if n:
            out[m.group(1)] = int(n.group(1))
    return out


def all_index_docs():
    rels = ["CODE_INDEX.md"]
    for k in G.LAYERS:
        for p, _ in G.LAYERS[k]["files"]:
            pass
    # 子索引文件名与层名一一对应
    for k in G.LAYERS:
        rels.append("reference/code-%s.md" % k)
    return rels


def code_files():
    """应该被登记的源文件：scripts/ 下的 .py 与 .ps1（不含 __pycache__）。"""
    out = []
    sd = os.path.join(ROOT, "scripts")
    for dp, dn, fn in os.walk(sd):
        dn[:] = [d for d in dn if d != "__pycache__"]
        for f in fn:
            if f.endswith((".py", ".ps1")):
                rel = os.path.relpath(os.path.join(dp, f), ROOT)
                out.append(rel.replace(os.sep, "/"))
    return sorted(out)


def main():
    fails, check = make_checker(TITLE)
    print()

    # ── ① 文件都被登记 ────────────────────────────────────────────
    idx_p = os.path.join(ROOT, "CODE_INDEX.md")
    check("CODE_INDEX.md 存在", os.path.exists(idx_p))
    registered = set()
    for rel in all_index_docs():
        registered |= set(doc_lines(rel))
    # tests/ 与 tools/ 在索引里以目录形式出现，不逐个登记
    for p in code_files():
        check("已登记: %s" % p, p in registered)

    # ── ② 行数与磁盘一致 ──────────────────────────────────────────
    for rel in all_index_docs():
        for path, n in sorted(doc_lines(rel).items()):
            full = os.path.join(ROOT, path)
            if not os.path.exists(full):
                continue
            real = sum(1 for _ in io.open(full, encoding="utf-8", errors="replace"))
            check("行数一致 %s = %d" % (path, real), real == n,
                  "索引写 %d，实际 %d —— 跑 tools/gen_code_tables.py 刷新" % (n, real))

    # ── ③ 依赖关系与 scan_deps 一致 ───────────────────────────────
    edges, rev, _, mods = scan_deps.build_graph()
    for k in G.LAYERS:
        for path, _human in G.LAYERS[k]["files"]:
            if not path.endswith(".py"):
                continue
            mn = G.mod_name(path)
            real = sorted(x.split(".")[-1] for x in edges.get(mn, ()))
            if not real:
                continue
            listed = listed_deps("reference/code-%s.md" % k, path)
            check("依赖一致 %s -> %s" % (path, ", ".join(real)),
                  listed is not None and set(real) <= set(listed),
                  "索引写 %s，实测 %s" % (listed, real))

    # ── ④ 分层数字 ────────────────────────────────────────────────
    idx = read_text(idx_p) if os.path.exists(idx_p) else ""
    cycles = scan_deps.find_cycles(edges)
    check("环数与实测一致（模块级 import 环 = %d 处）" % len(cycles),
          ("%d 处" % len(cycles)) in idx)
    check("模块总数与实测一致（%d）" % len(mods),
          re.search(r"\b%d\b" % len(mods), idx) is not None,
          "索引里没写 %d 这个数" % len(mods))
    # motion.py 被依赖数
    mrev = rev.get("scripts.motion", set())
    m = re.search(r"(\d+)\s*个模块依赖", idx)
    check("motion.py 依赖者数与实测一致（%d）" % len(mrev),
          m and int(m.group(1)) == len(mrev),
          "索引写 %s，实测 %d" % (m and m.group(1), len(mrev)))

    # ── ⑤ 层级规则 ────────────────────────────────────────────────
    for k in G.LAYERS:
        sub = "reference/code-%s.md" % k
        p = os.path.join(ROOT, sub)
        check("子索引存在: %s" % sub, os.path.exists(p))
        if os.path.exists(p):
            check("子索引链回父层: %s" % sub, "CODE_INDEX.md" in read_text(p))
        check("总览链到子索引: %s" % sub, sub in idx)

    return report(TITLE, fails)


def listed_deps(doc_rel, target):
    """抽出某份索引里「文件 X 那一行」声明的【依赖谁】一栏。

    ⚠️ 表格列序固定为：文件 | 职责 | 对外接口 | **依赖谁** | 被谁依赖 | 体积
    取的是**第 4 列**（索引 3）。第一版用"行内所有反引号"，
    结果拿到的是「被谁依赖」那一列—— 于是 7 条全判 FAIL，
    看起来像索引的依赖全错了，其实是我读错了列。
    """
    p = os.path.join(ROOT, doc_rel)
    if not os.path.exists(p):
        return None
    for ln in read_text(p).split("\n"):
        m = re.match(r"^\|\s*`%s`\s*\|" % re.escape(target), ln)
        if not m:
            continue
        cols = [c.strip() for c in ln.strip().strip("|").split("|")]
        if len(cols) < 4:
            return None
        return sorted(re.findall(r"`([\w.]+)`", cols[3]))
    return None


if __name__ == "__main__":
    sys.exit(main())