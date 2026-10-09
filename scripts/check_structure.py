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
  6. **依赖方向**（§六.4）：模块级 import 环 / 热点模块          [FAIL/ADVISE]
  7. **新文件可达**（§六.5）：新文件必须被入口或调用方链到         [FAIL]
     —— 没人链到的文件等于不存在
  8. **临时代码陈旧**（§六.6.1）：tools/ 里的脚本超期未动        [ADVISE]

用法：
    python scripts/check_structure.py
    python scripts/check_structure.py --json
"""
import argparse
import io
import json
import os
import re
import subprocess
import sys
import time

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
    # 2026-09-30：`build_transition_table.py`（3048 行）已拆成 `transition_probe/`
    # 包（common / data / decks / analysis / commands 五层），豁免已撤销。
    # 原条目留在这里作参照——**豁免被清掉时应该删掉条目**，
    # 否则清单会攒成一堆没人敢动的历史遗留。
    #
    # 留档的当时理由（供下次判断参考）：
    #   它是探针工具，关键路径需真 PowerPoint。后来发现这个理由**下错了结论**——
    #   搬运是纯移动，行为按构造不变，需要验证的是模块结构（import / 名字 /
    #   CLI 分发 / 测试），而这四件本地全能验。真正需要 PowerPoint 的是
    #   **重跑探测流水线**，不是验证搬家。
    #   ⚠️ 唯一的例外确实存在：**代码依赖自身位置**时搬运会改变行为——
    #   实测撞到一次（`__file__` 拼 `motion_catalog.json`，搬进子目录后指向了
    #   不存在的地方）。已改用 `common.SCRIPTS_DIR`。这类错静态检查看不见，
    #   只能靠跑测试。
}

# 这些顶层目录是开发期的，明确不随包分发（与 test_install_manifest 同源）
DEV_ONLY_DIRS = {"tests", "tools", "__pycache__", ".workbuddy", ".git"}

# 文档里引用的路径存在性检查 —— 举例性质的登记在此，带理由
PATH_REF_EXEMPT = {
    "scripts/_probe_orphan.py": {
        "why": "CONTRIBUTING §八.1 用它举例说明「负样本验证」——"
               "这个文件从来不该存在，是造出来验证判据的",
    },
}

# 归档目录不检查路径引用：**改历史记录就是毁证**（见 history/README.md）
PATH_REF_SKIP_DIRS = {"history"}

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


def strip_strikethrough(text):
    """剥掉行内 `~~删除线~~` —— 那是"此条作废"的留痕，不是活引用。

    索引里会保留"这个文件已拆，原条目留作参照"的记录（带删除线）。
    不剥掉的话，一份**已作废的留痕**会被判成"登记了不存在的文件"
    （2026-10-09 实际撞到）。
    """
    return "\n".join(re.sub(r"~~[^~]*~~", "", ln) for ln in text.split("\n"))


def check_path_refs(fails, advises):
    """文档里写 `scripts/foo.py`，就必须真的有这个文件。

    **为什么这条重要**：指向不存在文件的指引**比没有指引更坏** ——
    照着做的人会先花时间找它，找到最后才发现得自己写。
    2026-10-09 实测：278 处路径引用里 4 处指错，包括
    `facts/effects.json` 和 `facts/cameras.json`（根本不存在的文件）
    和 `scripts/alias_probe.ps1`（**从未在这个仓库出现过**）。
    而其中两处就在 `CODE_INDEX.md` 的「扩展点」表里 ——
    **那是这份索引存在的核心理由**。

    两种豁免（都写在代码里，不靠人记）：
      · `~~删除线~~` 里的 —— 那是作废留痕
      · `history/` —— 归档，**改它就是毁证**，整目录不检查
    """
    pat = re.compile(
        r"`((?:scripts|tools|tests|reference|facts)/[\w/.\-]+\.(?:py|ps1|md|json))`")
    for rel in sorted(walk_files(".md")):
        top = rel.split("/")[0]
        if top in PATH_REF_SKIP_DIRS:
            continue
        body = strip_strikethrough(read(os.path.join(ROOT, rel)))
        for i, ln in enumerate(body.split("\n")):
            for m in pat.finditer(ln):
                ref = m.group(1)
                if os.path.exists(os.path.join(ROOT, ref)):
                    continue
                if ref in PATH_REF_EXEMPT:
                    continue
                fails.append({
                    "rule": "path-ref-missing", "file": rel, "line": i + 1,
                    "ref": ref,
                    "note": "文档里写了这个路径，但它不存在 —— "
                            "读者照着做会先花时间找它。改掉，或登记进 "
                            "check_structure.PATH_REF_EXEMPT 并说明理由",
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


def check_reachability(fails, advises):
    """新文件必须被入口或调用方链到 —— 没人链到的文件等于不存在。

    **为什么这条必须有机器判据**：`CONTRIBUTING.md` §六.5 写了
    「新文件必须被入口或调用方链到」，但"被链到"是**跨文件**关系，
    光看新文件自己永远看不出它有没有人用。靠自觉的结果是：
    写完就忘，半年后仓库里躺着一个谁都不知道是什么的 `.py`。

    **怎么判「被链到」**（三种都算，逐级放宽）：
      ① 被另一个 `.py` 模块级 import          —— 代码链路
      ② 被 `install.ps1` 的复制清单收录        —— 会随包分发
      ③ 出现在 `CODE_INDEX.md` 的速查表里     —— 有人登记了它

    **豁免**：入口文件（`selftest.py` 等）与 `tools/` 下的开发期脚本
    不适用 —— 它们本来就是给人从命令行调的。这条只管"悄悄多出来的文件"。
    """
    # 入口文件：本身就是给人敲的，不该要求被别人 import
    entrypoints = {"scripts/motion.py", "scripts/build_transition_table.py",
                   "scripts/selftest.py", "scripts/review_assist.py"}
    # 已知的一层入口，其子命令通过 argv 分发，不单独登记
    known_entry = {"scripts/inspect_pptx.py", "scripts/install.ps1"}

    sys.path.insert(0, os.path.join(ROOT, "tools"))
    try:
        import scan_deps
        edges, _, _, mods = scan_deps.build_graph()
    except ImportError:
        edges, mods = {}, {}

    referenced = set()
    for _mod, deps in edges.items():
        for d in deps:
            # scan_deps 返回 `scripts.motion` 这种形式。
            # ⚠️ 试完两种扩展名要 **break** —— 试完 `.py` 失败还继续试 `.ps1`，
            # 循环结束时会留下最后一次的结果，于是 `tools/scan_deps.py`
            # 被推成 `tools/scan_deps.ps1`（不存在）→ index-stale 假阳性。
            for ext in (".py", ".ps1"):
                p = d.replace(".", "/") + ext
                if os.path.exists(os.path.join(ROOT, p)):
                    referenced.add(p)
                    break
    # 反向补齐：包内模块的父包也算被链到（build_transition_table -> transition_probe.*）
    for mod in mods:
        if mod.startswith("scripts.transition_probe."):
            referenced.add(mod.replace(".", "/") + ".py")

    # ③ 索引登记
    # ⚠️ 先剥掉 `~~删除线~~` 里的内容 —— 那是**已废弃记录**（"这个文件已拆，
    # 原条目留作参照"），不是活引用。不剥掉的话，一份"已拆"的留痕
    # 会被判成"索引登记了不存在的文件"（2026-10-09 实际撞到）。
    # Markdown 的删除线语义就是"此条作废"，判据应该理解它。
    def live_refs(text):
        for ln in text.split("\n"):
            ln = re.sub(r"~~[^~]*~~", "", ln)      # 行内删除线
            for m in re.finditer(r"`([\w/.\-]+\.(?:py|ps1))`", ln):
                yield m.group(1)

    idx_p = os.path.join(ROOT, "CODE_INDEX.md")
    if os.path.exists(idx_p):
        referenced |= set(live_refs(read(idx_p)))
    for k in ("injection", "probe", "verify", "tooling"):
        p = os.path.join(ROOT, "reference", "code-%s.md" % k)
        if os.path.exists(p):
            referenced |= set(live_refs(read(p)))

    # ② 分发清单
    for rel in sorted(walk_files(".py")):
        if not rel.startswith("scripts/"):
            continue
        if rel in entrypoints or rel in known_entry:
            continue
        if rel.endswith("__init__.py"):
            continue
        if rel in referenced:
            continue
        # 最后一道：是不是根本没有被登记进任何一份文档
        fails.append({
            "rule": "file-unreachable", "file": rel,
            "note": "这个 .py 既没被任何模块 import，也不在 CODE_INDEX / "
                    "code-*.md / install.ps1 里 —— 没人链得到它。"
                    "要么链上，要么删掉（§六.5）",
        })

    # ④ 反向：索引里登记了但磁盘上没有 —— 索引自己腐烂了
    # 这条**不能省**：③ 依赖索引完整，而索引完整正是这条要守的。
    # 少了它，一个漏登记的 tools/ 脚本会**静默通过**（判据以索引为准，
    # 索引没写它就等于它不存在）—— 这就是 2026-10-09 实际漏掉
    # split_tests.py 与 archive_handover_85.py 的原因。
    for rel in sorted(referenced):
        if not (rel.startswith("scripts/") or rel.startswith("tools/")):
            continue
        if not os.path.exists(os.path.join(ROOT, rel)):
            fails.append({
                "rule": "index-stale", "file": rel,
                "note": "索引里登记了这个文件，但磁盘上不存在 —— "
                        "文件被删或改名了，索引没跟上（§十.5）",
            })

    # ⑤ tools/ 里也一样：磁盘上有、但四份子索引里一个都没提 → 漏登记
    indexed = set()
    for k in ("injection", "probe", "verify", "tooling"):
        p = os.path.join(ROOT, "reference", "code-%s.md" % k)
        if os.path.exists(p):
            for m in re.finditer(r"`(tools/[\w.\-]+\.py)`", read(p)):
                indexed.add(m.group(1))
    td = os.path.join(ROOT, "tools")
    if os.path.isdir(td):
        for fn in sorted(os.listdir(td)):
            if not fn.endswith(".py"):
                continue
            rel = "tools/" + fn
            if rel not in indexed:
                fails.append({
                    "rule": "tool-not-indexed", "file": rel,
                    "note": "tools/ 里的脚本没进任何一份 code-*.md —— "
                            "索引漏了它。加进 tools/gen_code_tables.py 的 LAYERS",
                })


def check_stale_tools(advises):
    """`tools/` 里的脚本多久没动了 —— 临时代码该清（§六.6.1）。

    **为什么用 git 而不是文件时间戳**：文件 mtime 会被 checkout、
    复制、批量改写全部刷新，测出来的"多久没动"是假的。
    git log 记的是**真实的最后一次内容变更**。

    **为什么只报 ADVISE 不报 FAIL**：一个脚本"该不该留"是**意图问题**，
    只有作者知道（`CONTRIBUTING.md` §六.6.1 要求先问再删）。
    机器只能发现"它很久没动了"，**不能替你决定删不删**。

    豁免：只读体检工具常驻使用，不该被当成临时代码。
    判据是**文件名**—— 带 `scan_` / `check_` / `measure_` 前缀的
    是常驻体检器，其余都算可清理候选。
    """
    STALE_DAYS = 30
    PERMANENT = ("scan_deps", "check_structure", "verify_docs", "measure_",
                 "check_symptom", "gen_code_tables")
    td = os.path.join(ROOT, "tools")
    if not os.path.isdir(td):
        return
    try:
        out = subprocess.run(
            ["git", "log", "-1", "--format=%at", "--", "tools/"],
            capture_output=True, cwd=ROOT)
        if out.returncode != 0:
            return
    except (OSError, ValueError):
        return

    for fn in sorted(os.listdir(td)):
        if not fn.endswith(".py"):
            continue
        if any(fn.startswith(p) for p in PERMANENT):
            continue
        r = subprocess.run(["git", "log", "-1", "--format=%at",
                            "--", "tools/" + fn], capture_output=True, cwd=ROOT)
        if r.returncode != 0 or not r.stdout.strip():
            continue          # 从未提交过（新文件）→ 不报
        ts = int(r.stdout.strip().split()[0])
        days = (int(time.time()) - ts) // 86400
        if days >= STALE_DAYS:
            advises.append({
                "rule": "stale-tool", "file": "tools/" + fn, "days": days,
                "note": "%d 天没动过。是一次性脚本就该删；"
                        "要留就问用户，并写清为什么留（§六.6.1）" % days,
            })


def check_import_cycles(fails, advises):
    """模块级 import 环 —— 依赖只能从上往下（CONTRIBUTING §六.4）。

    ⚠️ **判据只算模块级 import**（`ast` 的 `tree.body`）。
    函数内延迟导入**不算依赖边** —— 那是**故意打断环**的手法。

    这个区分不是洁癖，是本项目踩过的坑：本项目已知一处表面环
    `motion ⇄ check_coverage`，靠 `motion.py check` 里的函数内导入打断，
    **是刻意且正确的设计**。判据不认这个区分，它就会天天误报，
    然后有人会去"顺手修掉"那个延迟导入 —— 那就把设计改坏了。

    依赖图解析复用 `tools/scan_deps.build_graph()`，
    **不在这里重写第二份**（§六.8：一份简化实现会让结论反向）。
    """
    sys.path.insert(0, os.path.join(ROOT, "tools"))
    try:
        import scan_deps
    except ImportError:
        advises.append({
            "rule": "deps-not-checked", "note": "tools/scan_deps.py 不在，跳过依赖检查",
        })
        return
    edges, rev, _, _ = scan_deps.build_graph()
    for cyc in scan_deps.find_cycles(edges):
        fails.append({
            "rule": "import-cycle", "cycle": " -> ".join(cyc),
            "note": "模块级 import 成环 —— 下层反向依赖上层了。"
                    "注意：函数内延迟导入不算环，判据只扫 tree.body",
        })
    # 改动波及面：被依赖最多的文件，作为 ADVISE 暴露出来
    # ⚠️ 不要取 top-N 截断 —— 阈值是「≥5 就是热点」，截断会让
    # 第 4 名之后的热点**静默不报**（2026-10-09 实际发生：decks 被 5 个
    # 依赖却没报，因为写了 [:3]）。判据是阈值，不是排名。
    for mod, users in sorted(rev.items(), key=lambda kv: -len(kv[1])):
        if len(users) >= 5:
            advises.append({
                "rule": "hot-module", "module": mod, "depended_by": len(users),
                "note": "被 %d 个模块依赖 —— 改它之前先看 CODE_INDEX.md 的依赖表"
                        % len(users),
            })


def main(argv=None):
    ap = argparse.ArgumentParser(description="结构体检（只读，可无人值守运行）")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    fails, advises = [], []
    check_docs(fails, advises)
    check_code(fails, advises)
    check_dirs(fails)
    check_import_cycles(fails, advises)
    check_reachability(fails, advises)
    check_stale_tools(advises)
    check_path_refs(fails, advises)
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
