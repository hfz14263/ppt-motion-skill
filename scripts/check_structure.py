#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""结构体检：把 CONTRIBUTING.md 的规则变成可执行的检查。

⭐ **这个脚本是整条链路里唯一可以无人值守运行的部分** ——
它**只读不写**：扫一遍、打印违规、返回退出码，一个字节都不改。
所以它可以安全地挂进定时任务，因为**最坏的结果是打出一份报告**。

（"按报告去改"**不能**无人值守 —— 那是删除类操作且含设计判断，
见 HANDOVER.md 关于无人值守的说明。）

检查项（对应 CONTRIBUTING.md 的章节）：

  1. **文档体积**（§二）
     - >=10000 字符 **且** >=4 个标题 → 必须有目录            [FAIL]
     - >=20000 字符 → 必须在规范允许的例外清单里，否则拆      [FAIL]
  2. **入口预算**（§二）：INDEX.md <= 6000 字符               [FAIL]
  3. **代码体积**（§二）
     - >=800 行 → 必须有「模块内容表」                        [ADVISE]
     - >=2000 行 → 必须拆成包，或在例外清单里                 [FAIL]
  4. **新目录登记**（§一 铁律 2）
     - 顶层目录必须出现在 install.ps1 的复制清单里，
       或明确登记为 DEV_ONLY                                   [FAIL]
  5. **结构自洽**（§六）：同一张表/同一段长文在两份文件里重复   [ADVISE]

用法：
    python scripts/check_structure.py
    python scripts/check_structure.py --json
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))

# ---- 阈值：改这里就要同步改 CONTRIBUTING.md §二 --------------------------
DOC_TOC_CHARS = 10_000        # 软上限：超了要目录
DOC_TOC_HEADS = 4             #   且至少这么多标题（没结构就没法索引）
DOC_SPLIT_CHARS = 20_000      # 硬上限：超了要拆
INDEX_BUDGET = 6_000          # 入口预算
CODE_MAP_LINES = 800          # 软上限：超了要「模块内容表」
CODE_SPLIT_LINES = 2_000      # 硬上限：超了要拆成包

# ---- 例外清单：**必须写明理由**，不能只写文件名 --------------------------
# 规范允许"硬上限但在例外里"的只有一种情况：**内聚且没有再拆的余地**。
# 加进这里要同时写清楚"为什么它不该拆"，否则下一个看的人只会照抄。
#
# ⭐ 例外可以带 `todo`：**豁免不等于遗忘**。带 todo 的条目每次体检都会
#   以 ADVISE 重新报出来，直到 todo 被清掉或条目被移除。
#   —— "临时豁免"最危险的形态是它悄悄变成了永久状态。
DOC_SPLIT_EXEMPT = {
    # 例：'reference/foo.md': {'why': '单一主题的连续论述，拆开每一节都不完整'},
}
CODE_SPLIT_EXEMPT = {
    "scripts/build_transition_table.py": {
        "why": "它是**探针工具**（生成探测 deck → PowerPoint 渲染 → 分析帧），"
               "不是运行时引擎。它的关键路径需要真 PowerPoint，本机无法在"
               "无头环境里验证；搬 3000 行代码却只验得了一半的路径，"
               "风险大于收益。另外它各节之间共享 NS / HYPOTHESES 等模块级常量，"
               "拆包需要先理清依赖顺序。",
        "todo": "在能跑 PowerPoint 的会话里拆成包：probe/{data,decks,analyze}.py "
                "+ __init__ 重新导出（保持 `import build_transition_table as B` 不变），"
                "每拆一步跑一次完整测试。见 HANDOVER.md「无人值守」一节。",
    },
}

# 这些顶层目录是开发期的，明确不随包分发（与 test_install_manifest 同源）
DEV_ONLY_DIRS = {"tests", "tools", "__pycache__", ".workbuddy", ".git"}

SKIP_DIRS = {".git", "__pycache__", ".workbuddy", "node_modules"}
# 上游 vendored 内容不算我们的（它们有自己的许可与 README）
UPSTREAM = re.compile(r"vendored from|upstream content", re.I)


def read(p):
    return io.open(p, encoding="utf-8").read()


def strip_code(text):
    return re.sub(r"```.*?```", "", text, flags=re.S)


def has_toc(body):
    """目录有两种合法形态：本工具生成的（有标记）与人写的（表格 + 站内锚点）。"""
    if "<!-- toc -->" in body:
        return True
    head = "\n".join(body.splitlines()[:40])
    return bool(re.search(r"^\|.*\]\(#", head, re.M))


def has_contents_map(body):
    """代码的「内容表」= 模块 docstring 里列出分节（不依赖行号，不会漂）。"""
    head = body[:2000]
    return bool(re.search(r"内容表|Contents|CONTENTS|模块结构", head))


def walk_files(exts):
    for base, dirs, files in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in SKIP_DIRS]
        for f in files:
            if f.endswith(exts):
                yield os.path.relpath(os.path.join(base, f), ROOT).replace("\\", "/")


def install_items():
    """install.ps1 的复制清单 —— 决定哪些顶层目录随包。"""
    p = os.path.join(ROOT, "install.ps1")
    if not os.path.exists(p):
        return set()
    m = re.search(r"foreach \(\$item in(.*?)\)\s*\{", read(p), re.S)
    return set(re.findall(r"'([^']+)'", m.group(1))) if m else set()


def _exempt(rel, table):
    """返回 (是否豁免, todo 文本)。兼容 str 与 dict 两种写法。"""
    if rel not in table:
        return False, ""
    v = table[rel]
    if isinstance(v, dict):
        return True, v.get("todo", "")
    return True, ""


def check_docs(fails, advises):
    for rel in sorted(walk_files(".md")):
        body = read(os.path.join(ROOT, rel))
        n_chars, n_lines = len(body), body.count("\n")
        if UPSTREAM.search(body[:600]):
            continue
        heads = len(re.findall(r"^#{2,3} (?!#)", strip_code(body), re.M))

        exempt, todo = _exempt(rel, DOC_SPLIT_EXEMPT)
        if n_chars >= DOC_SPLIT_CHARS and not exempt:
            fails.append({
                "rule": "doc-over-hard-limit", "file": rel,
                "chars": n_chars, "limit": DOC_SPLIT_CHARS,
                "note": "超过硬上限：按 CONTRIBUTING §3.1 拆，"
                        "或写进 check_structure.py 的 DOC_SPLIT_EXEMPT 并说明理由",
            })
        elif exempt and todo:
            advises.append({"rule": "waiver-open", "file": rel, "todo": todo})
        if n_chars >= DOC_TOC_CHARS and heads >= DOC_TOC_HEADS and not has_toc(body):
            fails.append({
                "rule": "doc-needs-toc", "file": rel,
                "chars": n_chars, "heads": heads,
                "note": "超过软上限且有足够结构：加目录（tools/add_toc.py）",
            })

    idx = os.path.join(ROOT, "INDEX.md")
    if os.path.exists(idx):
        n = len(read(idx))
        if n > INDEX_BUDGET:
            fails.append({"rule": "index-too-big", "file": "INDEX.md",
                          "chars": n, "budget": INDEX_BUDGET,
                          "note": "入口一旦需要翻页就不再是入口"})


def check_code(fails, advises):
    for rel in sorted(walk_files((".py", ".ps1"))):
        # 一次性搬迁脚本（tools/）不受体积约束 —— 它们活一次就完成使命
        if rel.startswith("tools/"):
            continue
        body = read(os.path.join(ROOT, rel))
        n = body.count("\n")
        exempt, todo = _exempt(rel, CODE_SPLIT_EXEMPT)
        if n >= CODE_SPLIT_LINES and not exempt:
            fails.append({
                "rule": "code-over-hard-limit", "file": rel, "lines": n,
                "limit": CODE_SPLIT_LINES,
                "note": "超过硬上限：拆成包（__init__ 重新导出以保持 import 不变），"
                        "或写进 CODE_SPLIT_EXEMPT 并说明理由",
            })
        elif exempt and todo:
            advises.append({"rule": "waiver-open", "file": rel,
                            "lines": n, "todo": todo})
        elif n >= CODE_MAP_LINES and not has_contents_map(body):
            advises.append({
                "rule": "code-needs-contents-map", "file": rel, "lines": n,
                "note": "接近硬上限：加「模块内容表」（docstring 里列分节名，不写行号）",
            })


def check_dirs(fails):
    """顶层目录必须被有意识地登记 —— 否则里面的文件**完全不被检查**。"""
    items = install_items()
    verify = read(os.path.join(ROOT, "scripts", "verify_docs.py"))
    scanned = set(re.findall(r'SCAN_DIRS = \(([^)]*)\)', verify)[0].split())
    scanned = {s.strip().strip('",') for s in scanned if s.strip()}

    for e in sorted(os.listdir(ROOT)):
        p = os.path.join(ROOT, e)
        if not os.path.isdir(p) or e in DEV_ONLY_DIRS or e.startswith("."):
            continue
        if e not in items:
            fails.append({
                "rule": "top-dir-not-registered", "dir": e,
                "note": "顶层目录没进 install.ps1 复制清单 —— 装出去会缺它；"
                        "或登记为 DEV_ONLY",
            })
        # 含文档的目录要进扫描范围，否则新文件没人管
        has_md = any(f.endswith(".md") for f in os.listdir(p))
        if has_md and e not in scanned:
            fails.append({
                "rule": "dir-not-scanned", "dir": e,
                "note": "这个目录有文档，但不在 verify_docs.py 的 SCAN_DIRS 里 —— "
                        "里面写一份没人链的文件不会有任何提示",
            })


def check_duplication(advises):
    """同一段长文出现在两处 —— 迟早只更新一处。"""
    seen = {}
    for rel in sorted(walk_files(".md")):
        body = read(os.path.join(ROOT, rel))
        for para in strip_code(body).split("\n\n"):
            para = para.strip()
            if len(para) < 120 or para.startswith("|") or para.startswith(">"):
                continue
            if para in seen and seen[para] != rel:
                advises.append({
                    "rule": "duplicated-block", "a": seen[para], "b": rel,
                    "note": "同一段文字在两处：%s…" % para[:40].replace("\n", " "),
                })
                break
            seen.setdefault(para, rel)


def main(argv=None):
    ap = argparse.ArgumentParser(description="结构体检（只读，可无人值守运行）")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    fails, advises = [], []
    check_docs(fails, advises)
    check_code(fails, advises)
    check_dirs(fails)
    check_duplication(advises)

    report = {"fails": fails, "advises": advises, "ok": not fails}
    if ns.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
        return 1 if fails else 0

    print("== 结构体检（只读，不改文件）==")
    print("  阈值：文档目录 >=%d 字符 且 >=%d 标题 / 拆 >=%d 字符"
          % (DOC_TOC_CHARS, DOC_TOC_HEADS, DOC_SPLIT_CHARS))
    print("        代码内容表 >=%d 行 / 拆 >=%d 行 / INDEX <=%d 字符"
          % (CODE_MAP_LINES, CODE_SPLIT_LINES, INDEX_BUDGET))
    print()
    if fails:
        print("-- FAIL --")
        for f in fails:
            print("  " + json.dumps(f, ensure_ascii=False))
    if advises:
        print("-- ADVISE（启发式，可能是有意的）--")
        for a in advises:
            print("  " + json.dumps(a, ensure_ascii=False))
    if not fails and not advises:
        print("OK  体积、索引、目录登记全部合规")
    print()
    print("注意：这只证明【结构合规】，不证明【内容正确】。")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
