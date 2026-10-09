#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 transition_probe/commands.py（1367 行 / 19 个 cmd）按功能域拆开。

**为什么拆**：1367 行超 800 软上限，且 19 个子命令其实服务**五个不同的探测维度**。
它是 `transition_probe/` 里唯一没拆过的文件 —— 上一轮拆包时它整块搬了过来
（`tools/split_transition_probe.py` 的 GROUPS 里 `commands` 那一组 `names=None`）。

**按什么拆**：按**探测维度**，不按行数、不按子命令的字母序。
依据是它自己 docstring 里那张内容表已经分好的组：

    commands_table.py      建表主链路 + 枚举扫描 + 合并成表 → 产出发布表
    commands_mechanism.py  机制层：切换挂在哪一页
    commands_shape.py      形态层 + 方向：每个效果看起来在做什么
    commands_attr.py       属性取值全集：同一元素换属性值
    commands_timing.py     切换 × 页内动画：结构平行、时间串行

**为什么这五个是一层、彼此不 import**：它们都只依赖下面的
common / data / decks / analysis 四层，**互相之间零调用**（实测：19 个函数
没有一个调用另一个）。所以拆出来是同一层的五个兄弟，不是一个调用链。

**验收纪律**（与前两次拆分同一套）：
  1. **外部接口零变化** —— 19 个 `cmd_*` 名字一个不少；
     `import build_transition_table as B` 与 `python build_transition_table.py <cmd>` 照旧。
  2. **逐符号搬运**，不改逻辑一个字。
  3. **名字清单必须对得上**：拆前 AST 扫出的顶层名 == 拆后可导入的名。

用法：
    python tools/split_commands.py --dry
    python tools/split_commands.py --apply
"""
import ast
import io
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = os.path.join(ROOT, "scripts", "transition_probe")
SRC = os.path.join(PKG, "commands.py")

# 每个新模块：文件名 → (标题, 职责, 装的 cmd_ 名单)
# 名单顺序 = 原文件里的出现顺序（搬运后读起来和原文件一致）
GROUPS = [
    ("commands_table", "建表主链路", 
     "产出**发布的实测表**：造候选 deck → 枚举扫描 → 读回 → 合并成表。",
     ["cmd_build", "cmd_collect", "cmd_enumdeck", "cmd_enumread", "cmd_table",
      "cmd_video", "cmd_sheets"]),
    ("commands_mechanism", "机制层",
     "切换**挂在哪一页**、几个槽位、时长听谁 —— 机制问题，不是效果问题。",
     ["cmd_anchordeck", "cmd_anchors"]),
    ("commands_shape", "形态层与方向",
     "每个效果**看起来在做什么**（含方向是否镜像）—— 形态表的数据来源。",
     ["cmd_shapedeck", "cmd_shapedeck2", "cmd_shapeanalyze", "cmd_shapes",
      "cmd_dirdeck", "cmd_dirmirror"]),
    ("commands_attr", "属性取值全集",
     "同一元素换属性值的后果（**不看方向** —— spokes/pattern 本来就不是方向）。",
     ["cmd_attrdeck", "cmd_attrdiff"]),
    ("commands_timing", "切换 × 页内动画",
     "两者同页时的结构平行与时间串行（页内动画排队等切换演完）。",
     ["cmd_timingdeck", "cmd_timingdiff"]),
]

MOD_DOC = '''"""{title}

{duty}

模块内容表（{n} 个命令）
----------------------------------------------------------------
{rows}

> 本模块由 `tools/split_commands.py` 从 `commands.py`（1367 行）按**探测维度**拆出
> （原文照搬，未改逻辑）。对外接口由 `../__init__.py` 重新导出 ——
> 这些 **CLI 子命令名写在 `facts/` 与 `reference/` 里，不能改名。**
"""'''

# 原文件头部那两行"由谁搬出来"的说明，新文件里换成上面 MOD_DOC 的尾注。
HEAD_SKIP = re_skip = None   # 由 collect() 实际判定


def collect(src_text):
    """返回 (symbols, order, imports, header_end)。

    symbols: {name: {'lines': [...], 'start': int, 'end': int}}
    注释块**往上吸** —— 分节横幅（`# ----` 那几行）必须跟着函数一起走，
    它们是这个函数存在的理由（"为什么要 enum scan"写在横幅下）。
    """
    tree = ast.parse(src_text)
    lines = src_text.splitlines(keepends=True)
    symbols, order = {}, []
    first_def = None
    for n in tree.body:
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            name = n.name
        elif isinstance(n, ast.Assign):
            t = [x.id for x in n.targets if isinstance(x, ast.Name)]
            if not t:
                continue
            name = t[0]
        else:
            continue
        first_def = first_def or n.lineno
        start = n.lineno
        i = start - 2
        # 往上吸连续的注释行（跳过空行不吸 —— 空行是多段注释的分界）
        while i >= 0 and lines[i].lstrip().startswith("#"):
            start = i + 1
            i -= 1
        symbols[name] = {"lines": lines[start - 1:n.end_lineno],
                         "start": start, "end": n.end_lineno}
        order.append(name)

    imports = []          # [(语句文本, 该语句提供的名字集合)]
    for n in tree.body:
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            txt = "".join(lines[n.lineno - 1:n.end_lineno]).rstrip()
            names = set()
            for a in n.names:
                names.add(a.asname or a.name.split(".")[0])
            imports.append((txt, names))
    return symbols, order, imports, (first_def or 1) - 1


def first_doc(names_lines):
    """取函数体的第一行 docstring 文本，做内容表的一行说明。"""
    txt = "".join(names_lines)
    try:
        fn = ast.parse(txt).body[0]
    except SyntaxError:
        return ""
    doc = ast.get_docstring(fn) or ""
    return doc.strip().split("\n")[0][:60]


def main():
    apply_ = "--apply" in sys.argv

    # ⚠️ 与 split_tests.py 同一个坑：这个脚本的产物会覆盖自己的输入。
    # 所以**只读一次源**，写完 5 个新文件后不再重跑自己。
    if not os.path.exists(SRC):
        raise SystemExit("找不到 %s（可能已经拆过了）" % SRC)
    src_text = io.open(SRC, encoding="utf-8").read()
    symbols, order, imports, hdr_end = collect(src_text)

    cmds = [n for n in order if n.startswith("cmd_")]
    print("源文件 %d 行，顶层符号 %d 个（其中 cmd_* %d 个）"
          % (len(src_text.split("\n")), len(order), len(cmds)))
    print()

    # 每个 cmd 归谁（重复分配或漏分配都要中止 —— 静默错分比报错更糟）
    owner = {}
    for mod, _title, _duty, names in GROUPS:
        for n in names:
            if n not in symbols:
                raise SystemExit("%s 里没有 %s" % (mod, n))
            if n in owner:
                raise SystemExit("%s 被分到两处（%s 和 %s）" % (n, owner[n], mod))
            owner[n] = mod
    missing = [c for c in cmds if c not in owner]
    if missing:
        raise SystemExit("有 cmd 没分组: %s（加进 GROUPS）" % missing)

    # 每个 cmd 用到哪些名字 —— 直接 AST walk 函数体
    used = {}
    for n in order:
        body_txt = "".join(symbols[n]["lines"])
        try:
            node = ast.parse(body_txt).body[0]
        except SyntaxError:
            used[n] = set()
            continue
        used[n] = {s.id for s in ast.walk(node) if isinstance(s, ast.Name)}

    print("== 拆分计划 ==")
    plan = {}
    for mod, title, duty, names in GROUPS:
        need = set()
        for n in names:
            need |= used[n]
        picked = [txt for txt, provided in imports if provided & need]
        nline = sum(len(symbols[n]["lines"]) for n in names)
        plan[mod] = (names, picked, nline, duty)
        print("  %-24s %d 命令 %4d 行  import %d 条"
              % (mod + ".py", len(names), nline, len(picked)))
    print()

    if not apply_:
        print("（--dry：没有写任何文件）")
        return 0

    for mod, title, duty, names in GROUPS:
        names2, picked, _n, _duty = plan[mod]
        rows = "\n".join("    %-13s %s" % (n[4:], first_doc(symbols[n]["lines"]))
                         for n in names2)
        head = MOD_DOC.format(
            title="transition_probe.%s — %s" % (mod.replace("commands_", "commands."), title),
            duty=duty, n=len(names2), rows=rows)
        body = [head, ""] + picked + ["", ""]
        for n in names2:
            body.append("".join(symbols[n]["lines"]).rstrip())
            body += ["", ""]
        p = os.path.join(PKG, mod + ".py")
        io.open(p, "w", encoding="utf-8", newline="\n").write(
            "\n".join(body).rstrip() + "\n")
        print("  写入 scripts/transition_probe/%-24s %d 命令"
              % (mod + ".py", len(names2)))

    print()
    print("接下来（本脚本不自动做，避免改错）：")
    print("  · __init__.py 与 build_transition_table.py 的 import 改指 5 个新模块")
    print("  · 删掉 commands.py（内容已全部搬走）")
    print("  · 跑 tools/gen_code_tables.py 刷新索引表")
    return 0


if __name__ == "__main__":
    sys.exit(main())