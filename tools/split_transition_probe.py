#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 build_transition_table.py（3048 行）拆成 transition_probe/ 包。

**为什么拆**：它是全仓库唯一超过代码硬上限（2000 行）的模块。
但拆它的**理由不止是行数** —— 它同时装着四件不同的事：
  · 探测数据（假设表 / 各种 PROBE 名单）—— 改数据不该碰到分析代码
  · deck 构造（把假设写成 xml）—— 纯 python-pptx，可本地验证
  · 帧分析（读渲染结果、算指标）—— 需要 PowerPoint 产物
  · 19 个 CLI 子命令 —— 编排

**为什么这次敢拆**（推翻我上一轮的判断）：
上一轮我写"关键路径需要真 PowerPoint，本机无法无头验证"。**那句话对，
但结论下错了** —— 它说的是"跑不通整条探测流水线"，而**这次做的是纯搬运**：
把整块函数从 A 文件挪到 B 文件，**行为按构造不变**。需要验证的不是流水线，
而是模块结构：import 能不能解析、名字全不全、CLI 分发对不对、测试过不过。
**这四件事本地全能验。** 所以搬运是安全的；重跑探测才需要 PowerPoint。

**验收纪律**（与文档拆分同一套）：
  1. **外部接口零变化** —— `__init__.py` 重新导出全部名字，
     `import build_transition_table as B` 与 `python build_transition_table.py <cmd>`
     都必须照旧。19 个 CLI 子命令写在 facts/ 与 reference/ 里，一个都不能少。
  2. **逐符号搬运**，不改逻辑一个字。
  3. **名字清单必须对得上**：拆前 AST 扫出的顶层名字 == 拆后可导入的名字。

用法：
    python tools/split_transition_probe.py --dry     # 只看计划
    python tools/split_transition_probe.py           # 执行
"""
import argparse
import ast
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
SRC = os.path.join(SCRIPTS, "build_transition_table.py")
PKG = os.path.join(SCRIPTS, "transition_probe")

# ---- 分区：模块 → 顶层符号（按依赖从低到高）-------------------------------
# 分层是**严格 DAG**：common → data → decks → analysis → commands → 入口。
# 唯一的倒挂（decks 用 analysis._load_json）已经把 _load_json 提到 common 解决。
GROUPS = [
    ("common", {
        "duty": "共享常量与通用小工具 —— 不依赖本包任何东西",
        "names": ["NS", "NSDECL", "DUR_MS", "PROBE_SPD", "SLIDE_SECONDS",
                  "FPS", "GRID_COLS", "GRID_ROWS",
                  "_load_json", "_norm", "_extract_transition"],
    }),
    ("data", {
        "duty": "**探测数据** —— 假设表与各探测名单。改数据不该碰到分析代码",
        "names": ["HYPOTHESES", "MOTION_SPECS", "RULES", "DEFAULT_ATTRS",
                  "DECK2_SPECS", "DIR_PROBE", "ATTR_PROBE",
                  "TIMING_ANIM", "TIMING_PROBE"],
    }),
    ("decks", {
        "duty": "**deck 构造** —— 把假设写成 slide XML。纯设置，不需要 PowerPoint",
        "names": ["_with_dir", "wrap_transition", "_slide_xml", "_two_slide_deck",
                  "_flat_deck", "_shape_deck", "_shape_deck2", "_probe_deck",
                  "_set_attrs", "_anim_deck"],
    }),
    ("analysis", {
        "duty": "**帧分析与表构造** —— 读渲染结果、算指标、合并成表",
        "names": ["_child_tag", "_child_attrs", "_child_render", "_lookup_enum",
                  "_template", "_dir_pairs", "_read_frames", "_window",
                  "_profile_metrics", "mirror_verdict", "_energy_trace",
                  "_hot_runs", "_entry_trace", "_timing_chart"],
    }),
    ("commands", {
        "duty": "**CLI 子命令** —— 编排上面的模块。这些名字写在 facts/ 与 reference/ 里，**不能改名**",
        "names": None,          # None = 所有 cmd_*
    }),
]

MOD_DOC = """\"\"\"{title}

{duty}
\"\"\"

# 本模块由 tools/split_transition_probe.py 从单文件的 build_transition_table.py
# 搬出来（原文照搬，未改逻辑）。对外接口由 __init__.py 重新导出。
"""


def collect(src_text):
    """返回 (symbols, order, imported_names)。

    symbols: {name: {'lines': [str], 'start': int, 'end': int, 'kind': str}}
    """
    tree = ast.parse(src_text)
    lines = src_text.splitlines(keepends=True)
    symbols, order = {}, []
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            names, kind = [n.name], "def"
        elif isinstance(n, ast.Assign):
            names = [t.id for t in n.targets if isinstance(t, ast.Name)]
            kind = "const"
        else:
            continue
        for name in names:
            # 往上吸连续的注释块（分节横幅与说明都在这里）
            start = n.lineno
            i = start - 2
            while i >= 0 and lines[i].lstrip().startswith("#"):
                start = i + 1
                i -= 1
            symbols[name] = {"lines": lines[start - 1:n.end_lineno],
                             "start": start, "end": n.end_lineno, "kind": kind}
            order.append(name)

    imported = set()
    for n in tree.body:
        if isinstance(n, ast.Import):
            for a in n.names:
                imported.add((a.asname or a.name).split(".")[0])
        elif isinstance(n, ast.ImportFrom):
            for a in n.names:
                imported.add(a.asname or a.name)
    return symbols, order, imported


def used_names(src_text, symbols):
    """每个顶层符号用到了哪些顶层名字 / 哪些 import 的名字。"""
    tree = ast.parse(src_text)
    imported = set()
    for n in tree.body:
        if isinstance(n, ast.Import):
            for a in n.names:
                imported.add((a.asname or a.name).split(".")[0])
        elif isinstance(n, ast.ImportFrom):
            for a in n.names:
                imported.add(a.asname or a.name)
    out = {}
    for n in tree.body:
        if not isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            continue
        top, imp = set(), set()
        for s in ast.walk(n):
            if isinstance(s, ast.Name):
                if s.id in symbols and s.id != n.name:
                    top.add(s.id)
                elif s.id in imported:
                    imp.add(s.id)
        out[n.name] = (top, imp)
    return out, imported


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry", action="store_true")
    ns = ap.parse_args(argv)

    src_text = io.open(SRC, encoding="utf-8").read()
    symbols, order, imported = collect(src_text)
    usage, _ = used_names(src_text, symbols)

    owner = {}
    for mod, spec in GROUPS:
        names = spec["names"] if spec["names"] is not None else \
            [n for n in order if n.startswith("cmd_")]
        for n in names:
            assert n in symbols, "%s 里没有 %s" % (mod, n)
            owner[n] = mod
    unassigned = [n for n in order if n not in owner and n != "main"]
    if unassigned:
        raise SystemExit("未分组: %s" % unassigned)

    print("== 拆分计划 ==")
    for mod, spec in GROUPS:
        names = [n for n in order if owner.get(n) == mod]
        sz = sum(len(symbols[n]["lines"]) for n in names)
        print("  %-10s %2d 符号 %4d 行  %s" % (mod, len(names), sz, spec["duty"][:34]))
    print("  %-10s %2d 符号 %4d 行  （留在原文件做入口）"
          % ("main", 1, len(symbols["main"]["lines"])))
    print()

    # 每个模块需要 import 什么
    plan = {}
    for mod, spec in GROUPS:
        names = [n for n in order if owner.get(n) == mod]
        sib, std = {}, set()
        for n in names:
            top, imp = usage.get(n, (set(), set()))
            std |= imp
            for d in top:
                m2 = owner.get(d)
                if m2 and m2 != mod:
                    sib.setdefault(m2, set()).add(d)
        plan[mod] = (names, sib, sorted(std))
        print("  %-10s ← 兄弟 %s | stdlib %s"
              % (mod, {k: sorted(v) for k, v in sib.items()} or "无", sorted(std)))
    print()

    if ns.dry:
        print("（--dry：没有写任何文件）")
        return 0

    os.makedirs(PKG, exist_ok=True)
    os.makedirs(os.path.join(PKG, "__pycache__"), exist_ok=True)

    titles = {"common": "共享常量与通用工具", "data": "探测数据表",
              "decks": "deck 构造器", "analysis": "帧分析与表构造",
              "commands": "CLI 子命令"}
    for mod, spec in GROUPS:
        names, sib, std = plan[mod]
        head = MOD_DOC.format(title="transition_probe.%s — %s" % (mod, titles[mod]),
                              duty=spec["duty"])
        body = [head, ""]
        # 依赖必须按层序 import（common → data → …），否则读起来是乱的
        for dep in [g for g, _ in GROUPS if g in sib]:
            body.append("from .%s import %s" % (dep, ", ".join(sorted(sib[dep]))))
        if sib:
            body.append("")
        if std:
            body.append("\n".join("import " + s for s in sorted(std)))
            body.append("")
        for n in names:
            body.append("".join(symbols[n]["lines"]).rstrip())
            body.append("")
            body.append("")
        p = os.path.join(PKG, mod + ".py")
        io.open(p, "w", encoding="utf-8", newline="\n").write("\n".join(body).rstrip() + "\n")
        print("  写入 %-34s %d 符号" % ("scripts/transition_probe/%s.py" % mod, len(names)))

    # __init__.py：重新导出全部 —— 这是"外部接口零变化"的实现
    init = ['"""transition_probe — 切换实测探针工具包（从单文件拆出）。',
            '',
            '**对外接口由本文件维持**：`import build_transition_table as B` 里能用的',
            '每个名字，在这里都能通过 `from transition_probe import *` 拿到。',
            '拆分只是搬家，不改接口。',
            '"""',
            '']
    for mod, spec in GROUPS:
        names, _, _ = plan[mod]
        init.append("from .%s import %s  # noqa: F401" % (mod, ", ".join(names)))
    init.append("")
    init.append("__all__ = [")
    for mod, spec in GROUPS:
        names, _, _ = plan[mod]
        init.append("    " + ", ".join('"%s"' % n for n in names) + ",")
    init.append("]")
    io.open(os.path.join(PKG, "__init__.py"), "w", encoding="utf-8",
            newline="\n").write("\n".join(init) + "\n")
    print("  写入 %-34s 重新导出 %d 个名字"
          % ("scripts/transition_probe/__init__.py", len(symbols) - 1))

    # 原文件退化成薄入口
    shim = ['#!/usr/bin/env python3', '# -*- coding: utf-8 -*-',
            '"""Build the slide-transition reference table（入口）。',
            '',
            '**本文件现在只是个入口。** 2026-09-30 之前它是一份 3048 行的单文件，',
            '现在实体在 `transition_probe/` 包里（common / data / decks / analysis /',
            'commands 五层）。拆的理由不是行数 —— 是它同时装着"探测数据"、',
            '"deck 构造"、"帧分析"、"CLI 编排"四件不同的事。',
            '',
            '**为什么用入口文件而不是直接删掉**：19 个 CLI 子命令写在',
            '`facts/transitions.json` 与 `reference/*.md` 里（`regenerate_with`），',
            '`tests/test_transition_table.py` 还写着 `import build_transition_table as B`。',
            '**名字是接口** —— 保留它，所有既有调用一行都不用改。',
            '',
            'Subcommands: 见 `python build_transition_table.py --help`',
            '"""',
            'import argparse',
            'import os',
            'import sys',
            '',
            'sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))',
            '',
            '# 重新导出整个包：`B.HYPOTHESES`、`B._profile_metrics` 等照旧可用',
            'from transition_probe import *  # noqa: F401,F403',
            'from transition_probe import __all__ as _pkg_all',
            '',
            '# 显式再绑一次，让静态检查与 `from ... import X` 都能看见',
            ]
    for mod, spec in GROUPS:
        names, _, _ = plan[mod]
        shim.append("from transition_probe.%s import %s  # noqa: F401"
                    % (mod, ", ".join(names)))
    shim += ["", ""]
    shim.append("".join(symbols["main"]["lines"]).rstrip())
    shim += ["", "", 'if __name__ == "__main__":', '    sys.exit(main())', '']
    io.open(SRC, "w", encoding="utf-8", newline="\n").write("\n".join(shim))

    print("  写入 %-34s 入口 + main()" % "scripts/build_transition_table.py")
    print()
    print("原文件 %d 行 → 入口 %d 行 + 包内 %d 行"
          % (len(src_text.splitlines()), len(shim),
             sum(len(symbols[n]["lines"]) for g, _ in GROUPS
                 for n in plan[g][0])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
