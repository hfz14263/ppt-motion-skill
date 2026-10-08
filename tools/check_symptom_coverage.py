"""只读体检：pitfall 顶层编号节是否自带「现象/症状」导语。

判定规则（与 CONTRIBUTING §六 对齐）：
  - 只认顶层标题 `## <n>. xxx`；`### xxx` 是它的子节，不算独立编号节。
  - 现象导语 = 正文里出现 `### 症状` / `### 现象` / `**现象**：` / `**症状**：`。
  - 现象导语应该**紧跟顶层标题**，出现在子节里也算（子节是正文的一部分）。

输出：缺现象导语的顶层节清单，供补写时逐节处理。
"""

import glob
import os
import re
import sys

TOP = re.compile(r"^## (\d+)\.\s+(.*)$")
# 必须带 re.M：正文首行通常是空行，`^` 需要能匹配每一行行首。
SYMPTOM = re.compile(r"^#{3,4}\s*(?:症状|现象)\s*$|^\*{0,2}(?:症状|现象)\*{0,2}\s*[:：]", re.M)


def split_top_sections(text):
    """按顶层 `## n.` 切分，返回 [(编号, 标题, 完整正文)]。"""
    lines = text.split("\n")
    heads = []
    for i, ln in enumerate(lines):
        m = TOP.match(ln)
        if m:
            heads.append((i, int(m.group(1)), m.group(2)))
    out = []
    for idx, (i, num, title) in enumerate(heads):
        end = heads[idx + 1][0] if idx + 1 < len(heads) else len(lines)
        out.append((num, title, "\n".join(lines[i + 1:end])))
    return out


def main():
    total = 0
    missing = []
    present = []
    for path in sorted(glob.glob("reference/pitfall-*.md")):
        if os.path.basename(path) == "pitfall-map.md":
            continue
        with open(path, encoding="utf-8") as f:
            text = f.read()
        for num, title, body in split_top_sections(text):
            total += 1
            row = (num, os.path.basename(path), title, len(body))
            if SYMPTOM.search(body):
                present.append(row)
            else:
                missing.append(row)

    print("pitfall 顶层编号节总数:", total)
    print("已有现象导语:", len(present))
    print("缺现象导语:", len(missing))
    print()
    print("== 已有（不用补）")
    for num, name, title, size in sorted(present):
        print("  §%-3s %-28s %s" % (num, name, title[:50]))
    print()
    print("== 缺（需补）")
    for num, name, title, size in sorted(missing):
        print("  §%-3s %-28s len=%-5d %s" % (num, name, size, title[:56]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
