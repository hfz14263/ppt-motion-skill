"""一次性验证：拆 CONTRIBUTING.md 后，每一章的内容是否逐字节还在。

**为什么必须验**：拆分是"搬运"，最危险的失败是**静默丢内容** ——
文件还在、能读、不报错，但某一节不见了。
`split_contributing.py` 按 `## 中文数字、` 切章，切错的后果和拆代码时漏函数一样。

**比对口径**：剥掉两种"拆分带来的合法差异"再比 ——
  · `<!-- toc -->` 目录块（由 add_toc.py 生成，原本没有）
  · `<a id="sN"></a>` 锚点行（同上）
  · 首尾空白
其余必须逐字节相同。

用法（一次性）：
    python tools/_verify_contributing_split.py
"""
import io
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CN = "一二三四五六七八九十"


def normalize(t):
    """剥掉「拆分带来的合法差异」，再压掉空行差异。

    ⚠️ 第一版只删了 `<a id=...>` 那一行，**没处理它留下的空行** ——
    `add_toc.py` 插锚点的形态是「标题 / 空行 / 锚点 / 空行 / 正文」，
    删掉锚点行后多出一个空行，于是三章被判"内容不一致"。
    **报「有问题」比不报更坏** —— 假阳性会让人养成忽略它的习惯。
    所以这里把连续空行压成一个，再比。
    """
    t = "\n".join(l for l in t.split("\n") if not re.match(r"^\s*<a id=", l))
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def chapters(text):
    """返回 {章号: 正文}，正文已规范化。"""
    marks = [(m.start(), m.group(1))
             for m in re.finditer(r"^## ([一二三四五六七八九十]+)、.*$", text, re.M)]
    out = {}
    for i, (pos, cn) in enumerate(marks):
        end = marks[i + 1][0] if i + 1 < len(marks) else len(text)
        body = text[pos:end]
        # 剥掉目录块（由 add_toc.py 生成，原本没有）
        body = re.sub(r"<!-- toc -->.*?(?=\n## |\Z)", "", body, flags=re.S)
        out[CN.index(cn) + 1] = normalize(body)
    return out


def main():
    old = subprocess.run(["git", "show", "HEAD:CONTRIBUTING.md"],
                         capture_output=True, cwd=ROOT).stdout.decode("utf-8")
    o = chapters(old)
    n = {}
    for f in ("CONTRIBUTING.md", "CODE_RULES.md"):
        n.update(chapters(io.open(os.path.join(ROOT, f), encoding="utf-8").read()))

    print("原文件 %d 章，拆后合计 %d 章" % (len(o), len(n)))
    miss = sorted(set(o) - set(n))
    extra = sorted(set(n) - set(o))
    if miss:
        print("！丢失章节: %s" % [CN[i - 1] for i in miss])
    diff = []
    for k in sorted(o):
        if k in n and o[k] != n[k]:
            a, b = o[k], n[k]
            # 找第一处不同
            for j, (x, y) in enumerate(zip(a.split("\n"), b.split("\n"))):
                if x != y:
                    diff.append((CN[k - 1], j + 1, x[:60], y[:60]))
                    break
            else:
                diff.append((CN[k - 1], -1, "长度不同 %d vs %d" % (len(a), len(b)), ""))
    if diff:
        print("！内容不一致 %d 章：" % len(diff))
        for cn, ln, x, y in diff:
            print("   第%s章 第%d行\n     原: %s\n     新: %s" % (cn, ln, x, y))
    else:
        print("内容不一致: 无 —— 每章逐字节还在")
    print("多出章节: %s" % ([CN[i - 1] for i in extra] or "无"))
    return 1 if (miss or diff) else 0


if __name__ == "__main__":
    sys.exit(main())
