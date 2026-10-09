"""一次性验证：拆完 motion.py 后，**每一条顶层语句**是否都还在。

**为什么要有这个（而不只是比符号名）**：2026-10-09 第一次拆完，
比"顶层符号名"显示 81/81 全对，但 `TRANSITIONS["none"] = ""` 和
`TRANSITION_ENUM["none"] = 0` 两行**丢了** —— 因为它们是 `Subscript` 赋值，
不是 `Name` 赋值。**验证脚本和搬运逻辑共享了同一个盲区，所以两边都说没问题。**

所以这一版按 **ast.unparse 后的语句文本**逐条比对：
不关心它叫什么、属于哪个符号，只关心**这条语句还在不在**。

用法（一次性，跑完即弃）：
    python tools/_verify_motion_split.py
"""
import ast
import io
import os
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

NEW_FILES = ["motion.py", "motion_xml.py", "motion_timing.py",
             "motion_media.py", "motion_spec.py"]


def top_statements(src):
    """每条顶层语句 -> 规范化文本。不含 docstring 与 import。"""
    out = []
    for n in ast.parse(src).body:
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            continue
        if isinstance(n, ast.Try):
            continue                      # lxml 的 try/except 由新文件各自生成
        if isinstance(n, ast.If) and "__main__" in ast.unparse(n.test):
            continue                      # 门面自己生成
        if isinstance(n, ast.Expr) and isinstance(n.value, ast.Constant):
            continue                      # docstring
        out.append(ast.unparse(n))
    return out


def main():
    old = subprocess.run(["git", "show", "HEAD:scripts/motion.py"],
                         capture_output=True, cwd=ROOT).stdout.decode("utf-8")
    new_src = ""
    for f in NEW_FILES:
        new_src += io.open(os.path.join(ROOT, "scripts", f),
                           encoding="utf-8").read() + "\n"

    o, n = top_statements(old), top_statements(new_src)
    # 用集合比 —— 同一个符号可能被 unparse 成一条；丢的会被看出来
    so, sn = set(o), set(n)
    missing = sorted(so - sn)
    extra = sorted(sn - so)

    print("原文件顶层语句 %d 条（去重 %d）" % (len(o), len(so)))
    print("新文件顶层语句 %d 条（去重 %d）" % (len(n), len(sn)))
    if missing:
        print("\n！！丢了 %d 条：" % len(missing))
        for m in missing[:10]:
            print("   " + m[:100])
    else:
        print("丢失: 无 —— 每条语句都搬到了")
    if extra:
        print("\n多出 %d 条（预期只有 _HERE 那两行）：" % len(extra))
        for e in extra[:5]:
            print("   " + e[:100])

    # 顺带确认那个「刻意打断环」的延迟导入没被提到模块级
    tree = ast.parse(io.open(os.path.join(ROOT, "scripts", "motion.py"),
                             encoding="utf-8").read())
    mod_level = set()
    for node in tree.body:
        if isinstance(node, (ast.Import, ast.ImportFrom)):
            for a in node.names:
                mod_level.add((a.asname or a.name).split(".")[0])
    for name in ("player", "check_coverage"):
        ok = name not in mod_level
        print("延迟导入 %-16s %s" % (name, "保留在函数内（正确）" if ok else "！被提到模块级"))
        if not ok:
            return 1

    # 命名空间完整性
    import motion
    want = {n.name for n in ast.parse(old).body if isinstance(n, ast.FunctionDef)}
    have = set(dir(motion))
    lacked = sorted(n for n in want if n not in have)
    print("函数名缺失: %s" % (lacked or "无"))
    return 1 if (missing or lacked) else 0


if __name__ == "__main__":
    sys.exit(main())
