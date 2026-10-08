"""只读：扫模块级 import，得出真实依赖图与环。

为什么要单独写：CONTRIBUTING §六 已经定了「依赖只能从上往下」的规则，
但**没有机器判据**。规则没有判据就会腐烂 —— 所以先量出真实形状，
才知道规范该写多少、哪些是现状问题而不是规范问题。

关键：只统计**模块级** import（ast 的 tree.body）。
函数内延迟导入是设计信号（motion ⇄ check_coverage靠它打断环），不算依赖边。

    python tools/scan_deps.py
"""

import ast
import os
import sys
from collections import defaultdict

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")


def module_name(path):
    rel = os.path.relpath(path, ROOT)
    rel = rel[:-3] if rel.endswith(".py") else rel
    rel = rel.replace(os.sep, ".")
    if rel.endswith(".__init__"):
        rel = rel[: -len(".__init__")]
    return rel


def local_name(mod):
    """模块名最后一段（用于 import 匹配）。"""
    return mod.rsplit(".", 1)[-1]


def module_level_imports(path):
    """返回 [(目标模块名, 是否相对导入)]，只取模块级。"""
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read(), filename=path)
    out = []
    for node in tree.body:  # 只扫模块级，函数内的刻意忽略
        if isinstance(node, ast.Import):
            for a in node.names:
                out.append((a.name, False))
        elif isinstance(node, ast.ImportFrom):
            if node.level and node.level > 0:
                mod = node.module or ""
                out.append(("." * node.level + mod, True))
            elif node.module:
                out.append((node.module, False))
    return out


def build_graph():
    """扫全仓库，返回 (edges, rev, unresolved, mods)。

    **抽出来是为了让别的脚本复用**。import 解析逻辑（`import motion`
    要认得出是 `scripts.motion`）很容易写错，写错了排查起来很费时间 ——
    所以只留这一份，`tools/gen_code_tables.py` 生成速查表时调的是它。

    返回：
      edges      {模块: {它依赖的模块}}      只含模块级依赖
      rev        {模块: {依赖它的模块}}      改动波及面
      unresolved {模块: [认不出的目标]}     多半是标准库/第三方
      mods       {模块: 绝对路径}
    """
    files = []
    for base in ("scripts", "tools", "tests"):
        for dirpath, dirnames, filenames in os.walk(os.path.join(ROOT, base)):
            dirnames[:] = [d for d in dirnames if d != "__pycache__"]
            for fn in filenames:
                if fn.endswith(".py"):
                    files.append(os.path.join(dirpath, fn))

    mods = {}
    for p in files:
        mods[module_name(p)] = p

    edges = defaultdict(set)
    unresolved = defaultdict(list)
    for mod, path in sorted(mods.items()):
        pkg = mod.rsplit(".", 1)[0] if "." in mod else mod
        for target, is_rel in module_level_imports(path):
            hit = None
            if is_rel:
                base = pkg
                for _ in range(target.count(".") - 1):
                    base = base.rsplit(".", 1)[0] if "." in base else base
                cand = (base + target.lstrip(".")).strip(".")
                hit = cand if cand in mods else None
            else:
                top = target.split(".")[0]
                if target in mods:
                    hit = target
                elif top in mods:
                    hit = top
                else:
                    for m in mods:
                        if local_name(m) == top:
                            hit = m
                            break
            if hit and hit != mod:
                edges[mod].add(hit)
            elif not hit:
                unresolved[mod].append(target)

    rev = defaultdict(set)
    for a, bs in edges.items():
        for b in bs:
            rev[b].add(a)
    return edges, rev, unresolved, mods


def find_cycles(edges):
    """模块级 import 环（DFS 找回边）。"""
    cycles = []

    def dfs(node, stack, onstack):
        stack.append(node)
        onstack.add(node)
        for nxt in sorted(edges.get(node, ())):
            if nxt not in edges:
                continue
            if nxt in onstack:
                cycles.append(stack[stack.index(nxt):] + [nxt])
            elif nxt not in stack:
                dfs(nxt, stack, onstack)
        stack.pop()
        onstack.discard(node)

    for m in sorted(edges):
        dfs(m, [], set())
    uniq = []
    seen = set()
    for c in cycles:
        key = tuple(sorted(set(c)))
        if key not in seen:
            seen.add(key)
            uniq.append(c)
    return uniq


def main():
    edges, rev, unresolved, mods = build_graph()
    print("模块数: %d（scripts / tools / tests）" % len(mods))
    print()

    print("== 有本地依赖的模块（按依赖数排序）")
    for mod in sorted(edges, key=lambda m: -len(edges[m])):
        print("  %-34s -> %s" % (mod, ", ".join(sorted(edges[mod]))))
    print()

    # 环检测
    uniq = find_cycles(edges)

    print("== 模块级 import 环: %d 处" % len(uniq))
    for c in uniq:
        print("  %s" % " -> ".join(c))
    print()

    # 反向依赖（谁依赖我）—— 用来判断改动波及面
    print("== 被依赖最多的模块（改动波及面）")
    for m in sorted(rev, key=lambda x: -len(rev[x]))[:12]:
        lines = 0
        if m in mods:
            with open(mods[m], encoding="utf-8") as f:
                lines = sum(1 for _ in f)
        print("  %-34s 被 %d 个模块依赖  (%d 行)" % (m, len(rev[m]), lines))
    print()

    ext = defaultdict(set)
    for mod, targets in unresolved.items():
        for t in targets:
            ext[t.split(".")[0]].add(mod)
    print("== 外部依赖（第三方/标准库）")
    for name in sorted(ext):
        who = ", ".join(sorted(ext[name]))
        print("  %-14s <- %s" % (name, who if len(who) < 70 else who[:67] + "..."))
    return 0


if __name__ == "__main__":
    sys.exit(main())
