"""给 pitfall 顶层编号节插入「现象」导语。

用法：
    python tools/add_symptom_leadin.py            # 先干跑，只报将改哪些
    python tools/add_symptom_leadin.py --apply    # 真改

设计约束（对应 CODE_RULES §六「覆盖式修改」）：
  - **覆盖式**修改现有文件，不新建任何文件。
  - 幂等：已经有导语就跳过，不会插两次。
  - 强校验：TSV 里的编号必须与实际顶层节一一对应；多一个少一个都中止。
    编号错位会让现象指向错误的坑，比不写更糟 —— 所以这里宁可失败。
  - 只在正文最前面插入，其余内容逐字节保留。
"""

import argparse
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TSV = os.path.join(ROOT, "tools", "symptom_leadins.tsv")
TOP = re.compile(r"^## (\d+)\.\s+(.*)$")
LEADIN = re.compile(r"^\*\*现象\*\*[:：]")


def read_leadins():
    leadins = {}
    with open(TSV, encoding="utf-8-sig") as f:
        for lineno, line in enumerate(f, 1):
            line = line.rstrip("\n").strip()
            #注释行、空行、BOM 后的标题行都跳过。
            if not line or line.startswith("#") or line.startswith(">"):
                continue
            m = re.match(r"^(\d+):\s*(.+)$", line)
            if not m:
                print("TSV 第 %d 行格式不对: %r" % (lineno, line[:60]))
                return None
            num = int(m.group(1))
            if num in leadins:
                print("TSV 里 §%d 出现了两次" % num)
                return None
            leadins[num] = m.group(2).strip()
    if not leadins:
        print("TSV 里一条导语都没有")
        return None
    return leadins


def read_top_sections(path):
    """返回 [(标题行号, 编号, 标题, 正文结束行号)]，只认顶层 `## n.`。"""
    with open(path, encoding="utf-8") as f:
        lines = f.read().split("\n")
    heads = []
    for i, ln in enumerate(lines):
        m = TOP.match(ln)
        if m:
            heads.append((i, int(m.group(1)), m.group(2)))
    out = []
    for idx, (i, num, title) in enumerate(heads):
        end = heads[idx + 1][0] if idx + 1 < len(heads) else len(lines)
        out.append((i, end, num, title))
    return lines, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true", help="真正写入（默认干跑）")
    args = ap.parse_args()

    leadins = read_leadins()
    if leadins is None:
        return 1

    files = sorted(glob.glob(os.path.join(ROOT, "reference", "pitfall-*.md")))
    files = [p for p in files if os.path.basename(p) != "pitfall-map.md"]

    found = {}
    for path in files:
        lines, secs = read_top_sections(path)
        for start, end, num, title in secs:
            if num in found:
                print("§%d 在多个文件里出现：%s 与 %s"
                      % (num, os.path.basename(found[num][0]), os.path.basename(path)))
                return 1
            found[num] = (path, start, end, title)

    tsv_nums = set(leadins)
    real_nums = set(found)
    only_tsv = sorted(tsv_nums - real_nums)
    only_real = sorted(real_nums - tsv_nums)
    if only_tsv or only_real:
        print("编号不匹配，中止。")
        if only_tsv:
            print("  TSV 里有、文档里没有: %s" % only_tsv)
        if only_real:
            print("  文档里有、TSV 里没有: %s" % only_real)
        return 1

    print("校验通过：%d 个编号一一对应。" % len(found))
    print()

    total = 0
    for num in sorted(found):
        path, start, end, title = found[num]
        lines, _ = read_top_sections(path)
        body = lines[start + 1:end]
        already = any(LEADIN.match(b.strip()) for b in body[:6] if b.strip())
        if already:
            print("  §%-3d 跳过（已有导语） %s" % (num, os.path.basename(path)))
            continue
        total += 1
        print("  §%-3d 插入  %-30s %s" % (num, os.path.basename(path), title[:44]))

    print()
    print("将插入 %d 处。%s" % (total, "已写入。" if args.apply else "干跑，未写入。"))

    if not args.apply or total == 0:
        return 0

    # 真正写入：按文件分组，从后往前插，避免行号漂移。
    by_file = {}
    for num in sorted(found):
        path, start, end, title = found[num]
        lines, _ = read_top_sections(path)
        if any(LEADIN.match(b.strip()) for b in lines[start + 1:end] if b.strip()):
            continue
        by_file.setdefault(path, []).append((start, num))

    for path, items in by_file.items():
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        for start, num in sorted(items, reverse=True):
            # 标题行之后先插空行分隔，再插导语
            insert_at = start + 1
            lines.insert(insert_at, "")
            lines.insert(insert_at + 1, "**现象**：%s" % leadins[num])
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("已写入 %s（%d 处）" % (os.path.basename(path), len(items)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
