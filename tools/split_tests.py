"""一次性拆分器：把 test_transition_table.py（1633 行 / 22 组）拆成三份。

为什么用脚本拆，不手抄：
  1633 行手抄 = 必然抄错，且错了很难看出来。
  脚本按行号切片，**逐字节搬运**，保证内容零漂移。

**按什么拆**：按**被测对象**拆，不按行数拆。
  test_transition_table.py  1–15 组  守切换实测表（transition_reference.json）
  test_docs.py             16–22 组  守文档与规范（编号 / 目录 / 规范一致性）
  test_code_index.py                守 CODE_INDEX 与真实代码同步（新建）
  helpers.py                        三份共用的 helper

**这个脚本是一次性的**：拆完就没用了。但按 CONTRIBUTING §六.6.1，
删除它要问过用户 —— 所以留着当"怎么拆测试文件"的样板。

    python tools/split_tests.py --dry     # 只报每段行数，不写
    python tools/split_tests.py --apply   # 真拆

实现要点（踩过的坑）：
  - **边界必须是「注释块的开头」**，不能切在 `# ---- 16. ----` 那一行之后 ——
    上一组的收尾注释要跟上一组走。所以边界取**上一个 check() 调用结束**的行。
  - **每段都要重新缩进**。原文件里所有测试在 main() 内缩进 4 空格，
    拆出去后仍是 main() 内的内容，缩进不变 —— 但 def main() 之后那行
    `print("== 切换实测表回归 ==")` 要换成 make_checker 生成的标题行。
  - **迁移完成后必须逐组跑一遍**，测试组的 PASSED 计数对不上就说明切错了。
"""
import io
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_REL = "tests/test_transition_table.py"


def read_source():
    """读**原始**的 1633 行版本。

    ⚠️ **不能读磁盘上那一份** —— 它就是本脚本的输出，已经被覆盖成
    1172 行的 switch 段了。第一版就是这么写的，于是：
      第一次跑 →成功写出两份
      第二次跑 → 拿半成品当源，切出乱七八糟的东西
    **拆分脚本必须从 git 取源**，因为它的产物会覆盖掉自己的输入。

    先试备份文件（若有），再退回 git HEAD。
    """
    bak = os.path.join(ROOT, "tests", ".test_transition_table.orig.py")
    if os.path.exists(bak):
        with io.open(bak, encoding="utf-8") as f:
            return f.read()
    import subprocess
    r = subprocess.run(["git", "show", "HEAD:" + SRC_REL],
                       capture_output=True, cwd=ROOT)
    if r.returncode != 0:
        raise SystemExit(
            "取不到 %s 的原始内容（git show 失败）。\n"
            "  拆分脚本的产物会覆盖自己的输入，所以必须从 git 或备份取源。\n"
            "  可以先 `git show HEAD:%s > %s` 再重跑。"
            % (SRC_REL, SRC_REL, SRC_REL))
    return r.stdout.decode("utf-8")

# (起, 止) 1-based 行号区间，含两端。
#  1–15 组：69–1187（从 print("== 切换实测表回归 ==") 到第 15 组收尾）
# 16–22 组：1189–1620（从「# ---- 16.」注释块到第 22 组最后一行 check）
#
# ⚠️ 上界 1620 是量出来的，不是猜的。原文件 1622 起是
# `print()` + `if fails:` + `return` + `if __name__ == "__main__"` 这段
# **收尾代码**，每份测试都会自己生成一份。把它切进去会让新文件里出现两个
# `return`—— 而 `dedent_body()` 的缩进校验会先拦下来（`if __name__` 顶格，
# 不是 4 空格）。**那道校验拦住了一次静默的错切。**
SEGMENTS = [
    ("switch", 69, 1187),
    ("docs", 1189, 1620),
]

HEADER_1_15 = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""切换实测表的回归测试（1–15 组）。

为什么需要这份测试：`scripts/transition_reference.json` 是量出来的，而
`scripts/motion.py` 直接读它。两者一旦漂移，症状不是报错，而是**写出一个
PowerPoint 会悄悄改掉的块** —— 看起来一切正常，文件就是不对。所以这里把
"表还成立"和"motion.py 还认这张表"钉住。

不测的东西：本测试**不**打开 PowerPoint。真机验证在 `tests/smoke.py`。

> **2026-10-08 从test_transition_table.py 拆出。** 拆分依据是**被测对象**，
> 不是行数：1–15 组守的是切换实测表与注入引擎的一致性，
> 与文档结构无关。同批拆出的还有 `test_docs.py`（16–22 组）。
> 组号是**稳定接口**，跨文件后仍然保持原编号 —— 别重排。

模块内容表（15 组；组号是稳定接口，别重排）
------------------------------------------------
   1  表本身（48 项齐全、字段完整）
   2  motion.py 认这张表（表和引擎不能漂）
   3  每个切换算"一个"，不是两个
   4  重新生成的表不会重复累积
   5  Choice 说毫秒、Fallback 说 spd，两者讲同一个故事
  5b  附加属性必须和占位符一起合流
   6  机制层文档 + facts 与引擎一致
   7  形态探测 deck 生成器
   8  形态分析器：方向必须真的被量出来
  8b  旋转指标（第四类）—— 前三类都看不见它
   9  形态层文档 + facts 一致
  10  证据 deck 的断言必须与形态文档同源
  11  选择层：判断可以反驳，引用不能漂移
  12  方向探测（补 3b §三 的缺口）
  13  属性取值全集（attrdeck / attrdiff）
  14  切换 × 页内动画：结构平行、时间串行
  15  48 项之间的取舍：引用可回溯 + 无自造效果名
"""
import io
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from helpers import HERE, ROOT, REF, LEGACY_SPECS, load_json, make_checker, report

TITLE = "切换实测表回归"


def main():
    fails, check = make_checker(TITLE)
'''

HEADER_16_22 = '''#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""文档与规范的回归测试（16–22 组）。

为什么需要这份测试：这一组守的不是代码，是**文档与实现的一致性**。
本项目里最危险的失败模式是**静默丢失**—— 文件能开、能跑、报告全绿，
但根因编号对不上了、目录漏了、规范里的数字和实现里的常量漂了。

这一组因此**只做机械可判的事**：编号完整/唯一、映射表一致、目录存在、
阈值三处同值。它**不**判断"哪一节该在哪份文件" —— 那是设计判断，会变。

> **2026-10-08 从 test_transition_table.py 拆出。** 拆分依据是**被测对象**：
> 16–22 组守的是文档与规范，与切换实测表无关。组号**保持原编号**。

模块内容表（7 组；组号是稳定接口，与拆分前一致）
------------------------------------------------
  16  踩坑手册拆分：编号完整、唯一、映射表一致
  17  症状文档拆分：现象栏可路由
  18  配方文档拆分：名实相符 + §n 可解析
  19  长文档的目录（导航不该靠滚）
  20  规范与实现：同一组数字，三处一致
  21  更正守则（打补丁，不重写）
  22  根因收敛后的两族结构
"""
import io
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from helpers import HERE, ROOT, read_text, make_checker, report

TITLE = "文档与规范回归"


def main():
    fails, check = make_checker(TITLE)
'''

FOOTER = '''
    return report(TITLE, fails)


if __name__ == "__main__":
    sys.exit(main())
'''

# 拆分后要重命名的局部变量/常量引用。
# 原文件里的 `_load_json` 和 `io.open(...,encoding="utf-8")` 两种读法都在用，
# 拆出去后统一走 helpers —— 但**不要顺手改语义**，只改读法。
REWRITES = [
    (r"\b_load_json\(", "load_json("),
]


def cut(lines, a, b):
    """取 1-based [a, b]，去掉尾部空行。"""
    seg = lines[a - 1:b]
    while seg and not seg[-1].strip():
        seg.pop()
    return seg


def dedent_body(seg):
    """把切出来的正文（原本在 main() 内，缩进 4）保持不变。

    原文件这些行本来就在 `def main():` 里，缩进已经是 4，
    所以**不需要动**。这里只做校验：首行缩进必须是 4。
    """
    for ln in seg:
        if ln.strip() and not ln.startswith("    "):
            raise SystemExit("切出来的正文缩进不对，首个非空行: %r" % ln)
    return seg


def build():
    src = read_source()
    lines = src.split("\n")

    out = {}
    for key, a, b in SEGMENTS:
        seg = dedent_body(cut(lines, a, b))
        header = HEADER_1_15 if key == "switch" else HEADER_16_22
        body = "\n".join(seg)
        for pat, rep in REWRITES:
            body = re.sub(pat, rep, body)
        # 原 main() 开头那行大标题删掉（make_checker 会重新打印）
        body = re.sub(r'^\s*print\("== (切换实测表回归|.*?) =="\)\n', "", body)
        out[key] = header + "\n" + body + "\n" + FOOTER
    return out


def main_cli():
    args = sys.argv[1:]
    apply_ = "--apply" in args
    out = build()
    for key, text in out.items():
        path = os.path.join(ROOT, "tests",
                            "test_transition_table.py" if key == "switch"
                            else "test_docs.py")
        n = text.count("\n")
        print("%-24s %5d 行  %s" % (os.path.basename(path), n,
                                    "已写入" if apply_ else "（dry-run）"))
        if apply_:
            with io.open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(text)
    if not apply_:
        print("\n加 --apply 才真写。原文件不会自动删 —— 确认三份都绿再手动删。")
    return 0


if __name__ == "__main__":
    sys.exit(main_cli())