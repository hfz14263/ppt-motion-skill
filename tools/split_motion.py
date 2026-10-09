#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""一次性拆分器：把 scripts/motion.py（1914 行 / 81 个顶层符号）按分层拆开。

**为什么现在敢拆**（推翻了 2026-10-08 的决定）：
当时的理由是对的 ——「要动 10 个依赖者，不是单次改动能验证的」。
现在条件变了：先量清了**依赖方向**，发现它是**严格 DAG**，而且分层比
`commands.py` 还干净。数据（`tools/scan_deps.py` + 内部调用图实测）：

    motion.py          1914 行 / 81 符号
      ├─ 内部调用关系：apply 层往下调，下层不回调上层 —— **零反向、零环**
      ├─ 10 个外部依赖者只用到约 30 个符号，且**全部是公开名**（无下划线）
      └─ 拆法：留 apply+CLI，其余按层出去，**外部接口靠 re-export 全保**

**拆成什么**（按层，不按行数）：

    motion_xml.py     35 符号 ≈ 480 行   常量 + zip + XML 字符串手术 + 单例 + 形状索引
                                         ★ 地基：无内部依赖，被其余全部依赖
    motion_timing.py  19 符号 ≈ 410 行   timing / transition XML 生成
    motion_media.py   12 符号 ≈ 310 行   3D 相机 + 图片填充
    motion_spec.py     9 符号 ≈ 200 行   spec 处理 + **结构校验**（三个 verify 用到的函数）
                                         ↑ 这三层互不调用，都只依赖 motion_xml
    motion.py          5 符号 ≈ 560 行   apply_motion / strip_animations / inspect / main
                                         + re-export 上面四层的**全部**名字

**为什么 `transitions_are_paired` / `active_transition_blocks` 归 spec 层**：
实测它们**不被本文件任何函数调用**，只被外部（`verify_motion.py` / `selftest.py`）用 ——
它们是纯校验函数，和 `geometry_fingerprint` 同类。按用途归，不按位置归。
（`MC` 常量只被这两个函数用，跟着它们走。）

**验收纪律**（与前三次拆分同一套）：
  1. **外部接口零变化** —— `import motion` 后每个原来能用的名字都还能用
  2. **逐符号搬运**，不改逻辑一个字；拆前拆后逐字节比对
  3. **名字清单必须对得上**：拆前 81 个顶层名 == 拆后 motion 命名空间里的名

用法：
    python tools/split_motion.py --dry
    python tools/split_motion.py --apply
"""
import ast
import io
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = os.path.join(ROOT, "scripts")
SRC_REL = "scripts/motion.py"
SRC = os.path.join(ROOT, SRC_REL)

GROUPS = [
    ("motion_xml", {
        "duty": "常量与 XML 底层 —— **整个引擎的地基**。\n"
                "\n"
                "    · zip 读写：不编辑的 part 逐字节复制（改坏文件的唯一方式就是解压再压缩）\n"
                "    · 前缀安全的字符串手术：只动属性值与文本，不重排 XML\n"
                "    · 子元素顺序：`<p:sld>` 是 **sequence 不是 bag**，位置错了元素会被静默丢弃\n"
                "    · 单例：**本项目所有损坏文件都是「同一元素出现两次」**\n"
                "    · 形状索引：deck 的 elementId → shape id",
        "names": [
            "HERE", "CATALOG_PATH", "SLIDE_RE", "P_NS", "MC_NS", "SHAPE_TAGS",
            "TRIGGERS", "GEOM_ATTRS", "WIPE_DIRS", "DIRECTIONAL_FILTERS",
            "TRANSITION_REF_PATH", "_load_transition_reference", "_TRANSITION_REF",
            "TRANSITIONS", "TRANSITION_ENUM", "EXT_CONTENT_TYPES",
            "read_parts", "write_parts", "xml_unescape", "xml_escape",
            "element_spans", "SLIDE_CHILD_ORDER", "SLIDE_ROOT_CHILDREN",
            "root_child_spans", "insert_in_slide_order", "replace_or_insert",
            "remove_elements", "set_singleton", "singleton_spans",
            "direct_children", "FILL_GROUP", "find_duplicate_singletons",
            "sp_tree_block", "index_shapes", "resolve_targets",
        ],
    }),
    ("motion_timing", {
        "duty": "timing 与 transition 的 XML 生成。\n"
                "\n"
                "    · `<p:timing>` 动画树：效果节点、时长覆盖、spd/dur 换算\n"
                "    · `<p:transition>` 切换：**挂在终点页**（写在第 N 页，动的是「进入第 N 页」）\n"
                "    · mc:AlternateContent 包装：PowerPoint 自己写的形式，不是我们发明的",
        "names": [
            "NodeIds", "split_effect_template", "build_effect_node",
            "set_wipe_direction", "set_leaf_durations", "set_anchor_durations",
            "apply_overrides", "build_timing", "root_declares",
            "add_root_namespace", "MORPH_TEMPLATE", "MORPH_OPTIONS",
            "MORPH_SPEEDS", "SPD_NOTCH_MS", "_spd_for_ms", "_spd_for_label",
            "build_transition", "drop_alternate_content",
            "drop_transition_alternate_content",
        ],
    }),
    ("motion_media", {
        "duty": "3D 相机与图片填充 —— 两类「非时序」的视觉属性。\n"
                "\n"
                "    · 3D 相机（`a:scene3d`）：角度对外是**度**，内部换算成 OOXML 的 1/60000\n"
                "    · 图片填充：窗口化 = 给负偏移，把整图塞进形状当前景",
        "names": [
            "CAMERA_ANGLE_UNIT", "CAMERA_MAX_ANGLE", "CAMERA_DEFAULT_PRESET",
            "CAMERA_PERSPECTIVE_PREFIXES", "deg_to_angle", "build_camera",
            "camera_is_perspective", "insert_scene3d", "window_insets",
            "build_fill_window", "insert_blip_fill", "ensure_content_type",
        ],
    }),
    ("motion_spec", {
        "duty": "spec 处理与**结构校验**。\n"
                "\n"
                "    校验三个函数**不被本文件任何函数调用**，只被外部用\n"
                "    （`verify_motion.py` / `selftest.py`）—— 所以按用途归这里，不按位置。\n"
                "    它们和 `geometry_fingerprint` 同类：**产出结论，不产出 XML**。",
        "names": [
            "geometry_fingerprint", "load_spec", "normalize_effects",
            "catalog_path", "normalize_spec", "schedule_spec", "slide_size",
            "MC", "transitions_are_paired", "active_transition_blocks",
        ],
    }),
    ("motion", {
        "duty": "主流程与 CLI —— 本模块是**门面**：apply 编排 + 重新导出下面四层。",
        "names": [
            "apply_motion", "remove_empty_alternate_content", "strip_animations",
            "inspect", "main",
        ],
    }),
]

MOD_DOC = '''"""{title}

{duty}

模块内容表（内容表；改这一层前先读这里）
----------------------------------------------------------------
{rows}

> 本模块由 `tools/split_motion.py` 从 `motion.py`（1914 行）按**分层**拆出
> （原文照搬，未改逻辑）。**外部接口由 `../motion.py` 重新导出** ——
> 10 个模块 `import motion` 用的名字一个都没变。
"""'''


def read_source():
    """从 git 取原始的 1914 行版本。

    ⚠️ 与前两次拆分同一个坑：脚本的产物会覆盖自己的输入，所以**必须从 git 取源**，
    否则第二次跑就是拿半成品当源（`split_tests.py` 第一版实际踩过）。
    """
    r = subprocess.run(["git", "show", "HEAD:" + SRC_REL],
                       capture_output=True, cwd=ROOT)
    if r.returncode != 0:
        raise SystemExit(
            "取不到 %s 的原始内容。\n"
            "  拆分脚本的产物会覆盖自己的输入，所以必须从 git 取源。" % SRC_REL)
    return r.stdout.decode("utf-8")


def collect(src_text):
    """收集顶层符号（含上方的连续注释块）+ import 语句 + **附属语句**。

    注释块必须跟着符号走 —— 这个文件的注释是**设计理由**（"为什么角度用度"、
    "为什么 <p:sld> 是 sequence"），不是装饰。丢了注释，搬家就等于毁证。

    ⚠️ **附属语句是这个脚本踩过的坑**：`TRANSITIONS["none"] = ""` 的赋值目标是
    `Subscript`，不是 `Name`；第一版只认 `Name`，于是这两行**被静默丢掉**
    （`TRANSITIONS` 少了 `none` 这个键 → `pytest` 报"旧 spec 名一个都没丢 FAIL ['none']"）。

    后果为什么难发现：**import 不报错、测试不报错、只有用到那个键的断言会红**。
    所以现在两类都处理：
      · 元组赋值 `A, B = ...`  → 每个 `Name` 都是独立符号
      · 下标赋值 `D[k] = v`    → 作为**附属语句**附着到根名 `D` 后面（顺序保持）
    并且 `assert_fully_consumed()` 会在拆分时验证**原文件每一条顶层语句都被搬走**。
    """
    tree = ast.parse(src_text)
    lines = src_text.splitlines(keepends=True)
    symbols, order, appendix = {}, [], {}
    for n in tree.body:
        names = []
        if isinstance(n, (ast.FunctionDef, ast.ClassDef)):
            names = [n.name]
        elif isinstance(n, ast.Assign):
            for t in n.targets:
                if isinstance(t, ast.Name):
                    names.append(t.id)
                elif isinstance(t, ast.Tuple):          # A, B = ...
                    names += [e.id for e in t.elts if isinstance(e, ast.Name)]
            if not names:
                # D[k] = v / obj.attr = v —— 附属语句，挂到根名后面
                for t in n.targets:
                    if isinstance(t, ast.Subscript) and isinstance(t.value, ast.Name):
                        appendix.setdefault(t.value.id, []).append(
                            (n.lineno, n.end_lineno))
                        break
                continue
        else:
            continue
        start = n.lineno
        i = start - 2
        while i >= 0 and lines[i].lstrip().startswith("#"):
            start = i + 1
            i -= 1
        for name in names:
            symbols[name] = {"lines": lines[start - 1:n.end_lineno],
                             "start": start, "end": n.end_lineno}
            order.append(name)

    # 把附属语句追加到根名符号后面（保持原顺序）
    for base, spans in appendix.items():
        if base not in symbols:
            raise SystemExit(
                "附属语句的根名 %r 不是顶层符号 —— 它会被漏搬。请检查。" % base)
        for _ln, end in sorted(spans):
            symbols[base]["lines"].extend(
                lines[symbols[base]["end"]:end])      # 含中间空行
            symbols[base]["end"] = end

    imports = []
    for n in tree.body:
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            txt = "".join(lines[n.lineno - 1:n.end_lineno]).rstrip()
            provided = set()
            for a in n.names:
                provided.add(a.asname or a.name.split(".")[0])
            imports.append((n.lineno, n.end_lineno, txt, provided))
        elif isinstance(n, ast.Try):
            # try: import yaml / except ImportError —— 整块保留
            txt = "".join(lines[n.lineno - 1:n.end_lineno]).rstrip()
            provided = set()
            for sub in ast.walk(n):
                if isinstance(sub, (ast.Import, ast.ImportFrom)):
                    for a in sub.names:
                        provided.add(a.asname or a.name.split(".")[0])
            if provided:
                imports.append((n.lineno, n.end_lineno, txt, provided))
    return symbols, order, imports, appendix


def assert_fully_consumed(src_text, symbols, imports, appendix):
    """**原文件每一条顶层语句都必须被搬走**，否则中止。

    ⚠️ 这条断言是 2026-10-09 那次漏搬换来的。当时的验证脚本只对比
    「顶层符号名」，而漏掉的是一行 `Subscript` 赋值 —— **验证脚本和搬运逻辑
    共享同一个盲区**，所以两边都说"没问题"。
    断言放在**拆分时**（而不是事后验证）才能拦住它。
    """
    consumed = set()
    for info in symbols.values():
        consumed |= set(range(info["start"], info["end"] + 1))
    for ln, end, _t, _p in imports:
        consumed |= set(range(ln, end + 1))
    for base, spans in appendix.items():
        for ln, end in spans:
            consumed |= set(range(ln, end + 1))
    # docstring / 文件头
    tree = ast.parse(src_text)
    if tree.body and isinstance(tree.body[0], ast.Expr):
        consumed |= set(range(1, tree.body[0].end_lineno + 1))

    missed = []
    for n in tree.body:
        span = set(range(n.lineno, n.end_lineno + 1))
        if not span <= consumed:
            # `if __name__ == "__main__":` 由新文件自己生成，允许
            if isinstance(n, ast.If) and "__main__" in ast.unparse(n.test):
                continue
            missed.append((n.lineno, type(n).__name__, ast.unparse(n)[:70]))
    if missed:
        raise SystemExit(
            "有顶层语句没被搬走（会在新文件里静默消失）：\n" +
            "\n".join("  L%-5d [%s] %s" % m for m in missed) +
            "\n  修 collect() 的处理分支，别手工绕过 —— 这类遗漏最难发现。")
    return True


def used_names(sym):
    """某个符号用到的所有名字（含属性访问的根名）。"""
    txt = "".join(sym["lines"])
    try:
        node = ast.parse(txt).body[0]
    except SyntaxError:
        return set()
    out = set()
    for s in ast.walk(node):
        if isinstance(s, ast.Name):
            out.add(s.id)
        elif isinstance(s, ast.Attribute) and isinstance(s.value, ast.Name):
            out.add(s.value.id)
    return out


def main():
    apply_ = "--apply" in sys.argv
    src_text = read_source()
    symbols, order, imports, appendix = collect(src_text)
    assert_fully_consumed(src_text, symbols, imports, appendix)
    print("源文件 %d 行，顶层符号 %d 个，import 语句 %d 条"
          % (len(src_text.split("\n")), len(order), len(imports)))
    print()

    owner = {}
    for mod, spec in GROUPS:
        for n in spec["names"]:
            if n not in symbols:
                raise SystemExit("%s 里没有 %s" % (mod, n))
            if n in owner:
                raise SystemExit("%s 被分到两处（%s / %s）" % (n, owner[n], mod))
            owner[n] = mod
    missing = [n for n in order if n not in owner]
    if missing:
        raise SystemExit("有符号没分组: %s" % missing)

    # 每个符号的名字用到哪些
    use = {n: used_names(symbols[n]) for n in order}

    # 组间依赖（谁要 from 谁 import 什么）
    plan = {}
    for mod, spec in GROUPS:
        names = [n for n in order if owner[n] == mod]
        need = set()
        for n in names:
            need |= use[n]
        # 兄弟模块的符号
        sib = {}
        for other, _s in GROUPS:
            if other == mod:
                continue
            deps = sorted(d for d in need if owner.get(d) == other)
            if deps:
                sib[other] = deps
        # stdlib / 三方 import
        std = []
        for _ln, _end, txt, provided in imports:
            if provided & need:
                std.append(txt)
        plan[mod] = (names, sib, std)
        nline = sum(len(symbols[n]["lines"]) for n in names)
        print("  %-14s %2d 符号 %4d 行  ← %s"
              % (mod + ".py", len(names), nline,
                 ", ".join(sorted(sib)) or "（无兄弟依赖）"))
    print()

    if not apply_:
        print("（--dry：没有写任何文件）")
        return 0

    # 写入四个新模块
    rows_of = {
        "motion_xml": "    常量        命名空间 / 正则 / 效果表 / 切换表 / 媒体类型\n"
                      "    zip         read_parts / write_parts\n"
                      "    XML 手术     xml_escape / xml_unescape / element_spans\n"
                      "    子元素顺序   insert_in_slide_order / replace_or_insert / remove_elements\n"
                      "    单例         set_singleton / singleton_spans / find_duplicate_singletons\n"
                      "    形状索引     index_shapes / resolve_targets / sp_tree_block",
        "motion_timing": "    timing      build_effect_node / build_timing / apply_overrides\n"
                         "    时长换算     _spd_for_ms / _spd_for_label / SPD_NOTCH_MS\n"
                         "    Morph       MORPH_TEMPLATE / MORPH_OPTIONS / MORPH_SPEEDS\n"
                         "    切换         build_transition / drop_transition_alternate_content",
        "motion_media": "    3D 相机      deg_to_angle / build_camera / insert_scene3d / camera_is_perspective\n"
                        "    图片填充     window_insets / build_fill_window / insert_blip_fill\n"
                        "    内容类型     ensure_content_type",
        "motion_spec": "    校验         geometry_fingerprint / transitions_are_paired\n"
                       "                 active_transition_blocks\n"
                       "    spec 处理    load_spec / normalize_spec / schedule_spec / slide_size\n"
                       "    catalog      catalog_path / normalize_effects",
    }
    for mod, spec in GROUPS:
        if mod == "motion":
            continue
        names, sib, std = plan[mod]
        head = MOD_DOC.format(
            title="motion.%s — %s" % (mod.replace("motion_", ""), spec["duty"].split("\n")[0]),
            duty=spec["duty"], rows=rows_of[mod])
        body = [head, ""]
        for dep in ("motion_xml",):          # 只有 xml 是下层；其余互不依赖
            if dep in sib:
                body.append("from %s import %s" % (dep, ", ".join(sib[dep])))
        if sib:
            other = [d for d in sib if d != "motion_xml"]
            if other:
                raise SystemExit("%s 出现了非预期的兄弟依赖: %s" % (mod, other))
            body.append("")
        body += std
        body += ["", ""]
        for n in names:
            body.append("".join(symbols[n]["lines"]).rstrip())
            body += ["", ""]
        p = os.path.join(SCRIPTS, mod + ".py")
        io.open(p, "w", encoding="utf-8", newline="\n").write(
            "\n".join(body).rstrip() + "\n")
        print("  写入 scripts/%-20s %2d 符号" % (mod + ".py", len(names)))

    # 重写 motion.py：apply + CLI + re-export
    names, sib, std = plan["motion"]
    head = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: OOXML motion injection engine（**门面**）.

注入 / 预览 / 去动画 / 体检 —— CLI 与主流程在这里，**实现按层在下面四个模块**：

    motion_xml.py     常量 + zip + XML 字符串手术 + 单例 + 形状索引   ← 地基
    motion_timing.py  timing / transition 的 XML 生成
    motion_media.py   3D 相机 + 图片填充
    motion_spec.py    spec 处理 + 结构校验

**本文件重新导出上面四层的全部名字**，所以 `import motion` 的调用方
（10 个模块）一行都不用改。2026-10-08 曾判断「不拆」——理由是"要动 10 个
依赖者，不是单次改动能验证的"；2026-10-09 量清了内部调用图是**严格 DAG**
（apply 往下调，下层零回调），拆它变成一次可验证的搬运。

模块内容表
----------------------------------------------------------------
  1. apply_motion          主流程：读 → 改 → 写（--assert-geometry 证明没动版面）
  2. remove_empty_alternate_content / strip_animations   去动画（preview 用）
  3. inspect               看一个 pptx 到底有没有动效
  4. main                  CLI 分发
  5. re-export             下面四层的全部公开名

Usage:
    python motion.py apply   --pptx IN.pptx --spec motion.yaml --out OUT.pptx [--assert-geometry]
    python motion.py inspect --pptx IN.pptx [--json]
    python motion.py preview --pptx IN.pptx --out STATIC.pptx
    python motion.py catalog [--kind entrance|emphasis|path]
"""
import argparse
import json
import os
import re
import sys
import zipfile

try:
    from lxml import etree as _ET
except ImportError:  # pragma: no cover
    _ET = None

# 新模块与本文件同目录 —— 作为脚本直接运行时 sys.path[0] 就是它，
# 但被 `import motion` 加载时不一定，所以显式补一次（幂等，重复无害）。
# ⚠️ 用 `_HERE` 而不是 `HERE`：`HERE` 由下面从 motion_xml 导入 ——
# 两处都定义的话，后导入的会覆盖先定义的，读代码的人会以为有两个真相。
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

# ---- 重新导出：把四个层拉回本命名空间（外部接口零变化的关键）-----------
#
# ⚠️ `main()` 里的 `import player` / `import check_coverage` 是**函数内延迟导入**，
# 它刻意打断 `motion ⇄ check_coverage` 那个环（见 CODE_RULES §六.4）。
# **它跟着 main() 一起留在这里 —— 别把那个 import 提到模块级。**
'''
    body = [head]
    for dep in ("motion_xml", "motion_timing", "motion_media", "motion_spec"):
        dnames = [n for n in order if owner[n] == dep]
        body.append("from %s import (  # noqa: F401" % dep)
        for i in range(0, len(dnames), 4):
            body.append("    " + ", ".join(dnames[i:i + 4]) + ",")
        body.append(")")
    body += ["", ""]
    # apply + CLI 本体：逐符号搬，但要补上它们对兄弟层的名字引用
    for n in names:
        body.append("".join(symbols[n]["lines"]).rstrip())
        body += ["", ""]
    body += ['if __name__ == "__main__":', '    sys.exit(main())', '']
    io.open(SRC, "w", encoding="utf-8", newline="\n").write("\n".join(body))

    total_new = sum(len(plan[m][0]) for m in plan if m != "motion")
    print("  写入 scripts/motion.py          %d 符号（含四层 re-export）" % len(names))
    print()
    print("原文件 %d 行 → 门面 + 四层；搬走 %d 个符号"
          % (len(src_text.split("\n")), total_new))
    return 0


if __name__ == "__main__":
    sys.exit(main())