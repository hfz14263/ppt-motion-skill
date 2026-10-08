"""给档3（保留现象正文的节）补根因链接。

档 3 的正文是现象侧独有的判断与实测，不能删；但必须能走到根因，
否则「两族通过根因串联」只是编号同源，不是结构可达。

实现要点（踩过的坑）：
  - **不能靠"下一节标题的行号"定位插入点**。多次插入后行号会漂，
    结果链接跑到下一节标题之后。正确做法：先扫出每节的真实末尾
    （最后一行非空、非来源尾注的行），再从后往前插。
  - 尾注**只保留一个**。脚本重跑不能让`来源 com-pitfalls §n` 出现两次。
  - 插入后必须保证与下一节标题之间有空行。

    python tools/link_symptom_to_pitfall.py --apply
"""

import argparse
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
TOP_PIT = re.compile(r"^## (\d+)\.\s+(.*)$")
SEC = re.compile(r"^#{2,4}\s*§\s*(\d+(?:\.\d+)*)\.?\s+(.*)$")
NOTE_LINE = re.compile(r"^<p align=\"right\"><sub>来源\s*com-pitfalls\s*§.*</sub></p>$")


def read_pitfall_index():
    idx = {}
    for path in sorted(glob.glob(os.path.join(REF, "pitfall-*.md"))):
        if os.path.basename(path) == "pitfall-map.md":
            continue
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        heads = [(i, int(m.group(1)))
                 for i, ln in enumerate(lines) if (m := TOP_PIT.match(ln))]
        for k, (i, num) in enumerate(heads):
            idx[num] = os.path.basename(path)
    return idx


def analyze(lines, pit):
    """返回 [(节标题行, 真实末尾行, 编号, 目标文件, 顶层号)]，只含档3。"""
    heads = []
    for i, ln in enumerate(lines):
        m = SEC.match(ln)
        if m:
            heads.append((i, m.group(1)))
    out = []
    for k, (i, raw) in enumerate(heads):
        end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
        body = lines[i + 1:end]
        if any("根因与修法" in b for b in body):
            continue  # 档1/2 已是一行链接
        if not any(NOTE_LINE.match(b.strip()) for b in body):
            continue  # 未处理
        num = int(raw.split(".")[0])
        if num not in pit:
            return None, "§%s 在 pitfall 侧不存在" % raw

        # 真实末尾：从后往前，跳过空行与所有来源尾注。
        last = end - 1
        while last > i and (not lines[last].strip() or NOTE_LINE.match(lines[last].strip())):
            last -= 1
        out.append((i, last, raw, pit[num], num))
    return out, None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    pit = read_pitfall_index()

    plan = []
    for path in sorted(glob.glob(os.path.join(REF, "symptom-*.md"))):
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        rows, err = analyze(lines, pit)
        if err:
            print("中止：%s —— %s" % (os.path.basename(path), err))
            return 1
        plan.append((path, rows))

    total = sum(len(r) for _, r in plan)
    print("待补根因链接: %d 处" % total)
    for path, rows in plan:
        for i, last, raw, pfile, num in rows:
            print("  §%-6s %-30s -> %s §%d" % (raw, os.path.basename(path), pfile, num))
    print()
    if not args.apply:
        print("干跑，未写入。")
        return 0

    for path, rows in plan:
        if not rows:
            continue
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        for i, last, raw, pfile, num in sorted(rows, key=lambda r: -r[1]):
            sub = "" if raw == str(num) else "（根因条目 §%s）" % raw
            # 清掉本节尾部所有尾注与空行，保证只有一个尾注。
            cut = last + 1
            while cut < len(lines) and (
                    not lines[cut].strip() or NOTE_LINE.match(lines[cut].strip())):
                cut += 1
            new_tail = [
                "",
                "**根因与修法**：见 [`%s`](%s) §%d%s。" % (pfile, pfile, num, sub),
                "",
                '<p align="right"><sub>来源 com-pitfalls §%s</sub></p>' % raw,
                "",
            ]
            lines[last + 1:cut] = new_tail
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("已写入 %s（%d 处）" % (os.path.basename(path), len(rows)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
