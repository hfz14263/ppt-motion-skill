"""给已收敛的 symptom 节补「现象」导语。

背景：档2 节的正文被替换成了「根因与修法」链接，节标题是根因式表述
（`### §8 p14:dur 需要声明前缀`）。用户是**按现象**找进来的，
只看到根因标题会不知道对不对号 —— 所以每节要有一句现象描述。

现象文案复用 tools/symptom_leadins.tsv（那是按现象写的，与本项目一一对应），
所以这里只做插入，不重新造词。

    python tools/add_symptom_view.py--apply
"""

import argparse
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")
TSV = os.path.join(ROOT, "tools", "symptom_leadins.tsv")
SEC = re.compile(r"^#{2,4}\s*§\s*(\d+(?:\.\d+)*)\.?\s+(.*)$")
NOTE = re.compile(r"来源\s*com-pitfalls\s*§")


def read_leadins():
    leadins = {}
    with open(TSV, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or line.startswith(">"):
                continue
            m = re.match(r"^(\d+):\s*(.+)$", line)
            if m:
                leadins[m.group(1)] = m.group(2).strip()
    return leadins


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()
    leadins = read_leadins()

    plan = []
    for name in sorted(os.listdir(REF)):
        if not (name.startswith("symptom-") and name.endswith(".md")):
            continue
        path = os.path.join(REF, name)
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        heads = []
        for i, ln in enumerate(lines):
            m = SEC.match(ln)
            if m:
                heads.append((i, m.group(1), m.group(2)))
        for k, (i, raw, title) in enumerate(heads):
            end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
            body = lines[i + 1:end]
            if any(b.strip().startswith("**现象**") for b in body):
                continue
            if not any(NOTE.search(b) for b in body):
                continue
            # 区分档1/2 与档3 不能只看「有没有根因链接」——
            # 档3 补链接之后也有了。真正的区别是**除链接外还有没有正文**：
            # 档1/2 的正文被收敛掉了，档3 保留了现象描述。
            prose = [
                b for b in body
                if b.strip()
                and "根因与修法" not in b
                and not NOTE.search(b)
            ]
            top = raw.split(".")[0]
            if top not in leadins:
                print("中止：§%s 没有对应现象文案" % raw)
                return 1
            extra = ""
            if raw != top:
                extra = "（§%s 的第 %s 条）" % (top, raw.split(".", 1)[1])
            # prose 为空 = 档1/2，现象导语是唯一正文，必须补。
            # prose 非空 = 档3，只补统一的标记行，正文不动。
            kind = "档1/2 收敛节" if not prose else "档3 保留正文"
            plan.append((path, i, raw, title, leadins[top] + extra, kind))

    print("待补现象导语: %d 处" % len(plan))
    for path, i, raw, title, text, kind in plan:
        print("  §%-6s %-30s %-14s %s"
              % (raw, os.path.basename(path), kind, text[:44]))
    print()
    if not args.apply:
        print("干跑，未写入。")
        return 0

    by_file = {}
    for path, i, raw, title, text, kind in plan:
        by_file.setdefault(path, []).append((i, text))
    for path, items in by_file.items():
        with open(path, encoding="utf-8") as f:
            lines = f.read().split("\n")
        for i, text in sorted(items, reverse=True):
            lines.insert(i + 1, "")
            lines.insert(i + 2, "**现象**：%s" % text)
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("已写入 %s（%d 处）" % (os.path.basename(path), len(items)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
