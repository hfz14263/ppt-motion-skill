"""一次性归档器：把 HANDOVER.md §8 里**已完成**的三节搬进 history/。

为什么搬：§8「还没做的」这一章在 2026-09-30 已经因为同一个原因搬过一次
（当时占了全文 299/593 行，见 `history/README.md`）。2026-10-08 又往里
写了 §8.3.1 / §8.4 / §8.5 三节**已完成**的记录，全章涨到 9716 字符、
整份 HANDOVER 21901 字符，**越过 20000 硬上限**。

**同一个坑踩两次。** 第一次搬完没留下"往§8 里写完成记录是回流"这条，
所以又流回来了。所以这次除了搬，还要在 §8 开头留一句明确的禁令。

分诊依据（`CONTRIBUTING.md` §三）：
  §8.3.1 根因收敛 / §8.4 更正 / §8.5 代码索引  → **已完成 + 当时怎么想错的**
  → 会无限增长 → 归档到 `history/code-index-system.md`
  §8.1 / §8.2 / §8.3 主体                     → **常驻规则与真待办**
  → 留在 HANDOVER

编号**沿用原编号不重排** —— 这些编号被别处指认过，
重排会让"§8.4 说的那件事"变得没法指认（同 `history/README.md` 的规矩）。

    python tools/archive_handover_85.py --dry
    python tools/archive_handover_85.py --apply
"""
import io
import os
import re
import subprocess
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HANDOVER = os.path.join(ROOT, "HANDOVER.md")
TARGET = os.path.join(ROOT, "history", "code-index-system.md")

# (§8.3.1 标题行,  §8.5.3 结束后的分隔线行) 1-based
SRC_REL = "HANDOVER.md"

ARCHIVE_HEADER = '''# 代码索引体系与根因收敛（2026-10-08）

> 本文件是**归档**，不是待办。原属 `HANDOVER.md` §8「还没做的」，
> 2026-10-08 因该章涨到 9716 字符（整份 21901，**越过 20000 硬上限**）而搬出。
> 索引见 [`README.md`](README.md)；项目概览见 [`../HANDOVER.md`](../HANDOVER.md)。
> 活着的索引是 [`../CODE_INDEX.md`](../CODE_INDEX.md)。

**为什么这些内容还留着**：三节里都带着"当时为什么这么判""踩了什么坑"的细节 ——
那是它们唯一的价值，也是下次不重犯的依据。**改写就等于毁证。**

> ⚠️ **编号沿用 `HANDOVER.md` §8 的原编号，不连续是正常的。**
> 不要重排：这些编号被别处指认过，重排会让"§8.4 说的那件事"没法再指。

> ⚠️ **同一个坑踩了两次。** 2026-09-30 已经因同一原因搬过一次
> （当时 §8 占 299/593 行），`history/README.md` 里记着。
> 搬完却没留下"往 §8 里写完成记录是回流"这条，2026-10-08 又流回来。
> → 现在 `HANDOVER.md` §8 开头有禁令。

---

'''

# 留在 §8 里的替代内容：原 8.3.1 位置放一个指向归档的短条目
REPLACEMENT = '''### 8.3.1 已完成的三项（正文已归档）

2026-10-08 这一轮做完并验证的三件事，正文连同当时的误判与踩坑记录
已搬进 [`history/code-index-system.md`](history/code-index-system.md)：

| 原编号 | 主题 | 现在的落点 |
| --- | --- | --- |
| §8.3.1 | symptom / pitfall 按根因收敛 | `history/code-index-system.md` |
| §8.4 | 「不能自动做」有一半是错的（更正） | `history/code-index-system.md` |
| §8.5 | 代码索引体系（`CODE_INDEX.md` + 四份子索引） | [`CODE_INDEX.md`](CODE_INDEX.md)（活的索引） |

> ⚠️ **别再往这一节里写"做完了什么"。**
> §8 装的是**真待办**；完成记录属于 [`history/`](history/README.md)。
> 这是 2026-09-30 和 2026-10-08 **两次**越界的原因。

'''


def read_head():
    """从 git 取当前 HANDOVER 的内容。

    ⚠️ 与 `tools/split_tests.py` 同一个坑：脚本的产物会覆盖自己的输入，
    所以必须从 git 取源，否则第二次跑就乱了。
    """
    r = subprocess.run(["git", "show", "HEAD:" + SRC_REL],
                       capture_output=True, cwd=ROOT)
    if r.returncode != 0:
        return io.open(HANDOVER, encoding="utf-8").read()
    return r.stdout.decode("utf-8")


def main():
    apply_ = "--apply" in sys.argv
    src = read_head()
    lines = src.split("\n")

    # 找 §8.3.1 标题行
    start = None
    for i, ln in enumerate(lines):
        if ln.startswith("### 8.3.1 "):
            start = i
            break
    if start is None:
        print("找不到 §8.3.1 —— 可能已经归档过了。检查 history/。")
        return 0

    # 找它之后第一个 `---` （§8 章末），那就是三节的末尾
    end = None
    for i in range(start, len(lines)):
        if lines[i].strip() == "---":
            end = i
            break
    if end is None:
        print("找不到 §8 的章末 `---`，无法确定边界。")
        return 1

    body = "\n".join(lines[start:end]).rstrip() + "\n"
    new_lines = (lines[:start]
                 + REPLACEMENT.rstrip("\n").split("\n")
                 + lines[end:])
    out = "\n".join(new_lines)

    archive = ARCHIVE_HEADER + body
    print("搬走 %d 行（第 %d–%d 行）" % (end - start, start + 1, end))
    print("  HANDOVER.md      %d -> %d 字符" % (len(src), len(out)))
    print("  history/%s  新建，%d 字符"
          % (os.path.basename(TARGET), len(archive)))
    if apply_:
        with io.open(HANDOVER, "w", encoding="utf-8", newline="\n") as f:
            f.write(out)
        with io.open(TARGET, "w", encoding="utf-8", newline="\n") as f:
            f.write(archive)
        print("\n已写入。")
    else:
        print("\n加 --apply 才真写。")
    return 0


if __name__ == "__main__":
    sys.exit(main())