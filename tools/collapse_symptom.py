"""分级收敛 symptom 侧正文：把重复正文换成指向根因的链接。

三档（由 tools/measure_overlap.py 实测得出，阈值写在这里可复核）：
  档1  full >= 0.99  几乎逐字相同 → 回收正文，只留「现象 → 根因 §n」
  档2  0.80 <= full < 0.99  高度相似，symptom 侧多一段现象导语
                        → 保留现象导语，正文换成根因链接
  档3  full < 0.80   真正改写过的现象描述 → 整节保留，只在末尾补根因链接

硬约束：
  - **不删除任何文件**，旧入口文件名与 §n 编号全部保留。
  - 每节必须留下 `来源 com-pitfalls §n` 尾注（历史引用依赖它）。
  - 编号与 pitfall 侧对不上时中止 —— 错链比重复更糟。
  - 幂等：已经转换过（正文只剩索引行）的节跳过。

用法：
    python tools/collapse_symptom.py            # 干跑
    python tools/collapse_symptom.py --apply
"""

import argparse
import difflib
import glob
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REF = os.path.join(ROOT, "reference")

TOP_PIT = re.compile(r"^## (\d+)\.\s+(.*)$")
SEC_ANY = re.compile(r"^#{2,4} §?(\d+(?:\.\d+)*)\.?\s+(.*)$")
SRC_NOTE = re.compile(r"来源\s*com-pitfalls\s*§\s*(\d+(?:\.\d+)*)\s*$")
CODE = re.compile(r"```.*?```", re.S)
SUB = re.compile(r"<sub>.*?</sub>", re.S)

TIER1 = 0.99
TIER2 = 0.80


def load(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def norm(text, keep_code):
    if not keep_code:
        text = CODE.sub("", text)
    text = SUB.sub("", text)
    return re.sub(r"\s+", "", text)


def read_pitfall_index():
    """§n -> (文件名, 标题)。只用顶层 `## n.`，子编号回退到顶层。"""
    idx = {}
    for path in sorted(glob.glob(os.path.join(REF, "pitfall-*.md"))):
        if os.path.basename(path) == "pitfall-map.md":
            continue
        lines = load(path).split("\n")
        heads = []
        for i, ln in enumerate(lines):
            m = TOP_PIT.match(ln)
            if m:
                heads.append((i, int(m.group(1)), m.group(2)))
        for k, (i, num, title) in enumerate(heads):
            end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
            idx[num] = (os.path.basename(path), title, "\n".join(lines[i + 1:end]))
    return idx


def top_key(raw):
    return int(raw.split(".")[0])


def split_symptom_sections(lines):
    """按**带 § 的编号节**切分。

    关键：只有 `### §n` 才是节边界。
    - `### 排查方法：…`（无编号）→ 上一节正文的一部分，不能切。
    - `### 20.1 xxx`（有数字无§）→ §20 的子节，同样不能切。
      它在 pitfall 侧对应 `### 20.1`，但对用户来说它不是独立可检索条目。
    所以这里只认`§`，其余 H3 一律并入当前节。
    """
    heads = []
    for i, ln in enumerate(lines):
        m = re.match(r"^#{2,4}\s*§\s*(\d+(?:\.\d+)*)\.?\s+(.*)$", ln)
        if m:
            heads.append((i, m.group(1), m.group(2)))
    out = []
    for k, (i, num, title) in enumerate(heads):
        end = heads[k + 1][0] if k + 1 < len(heads) else len(lines)
        out.append((i, end, num, title))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--apply", action="store_true")
    args = ap.parse_args()

    pit = read_pitfall_index()
    print("pitfall 顶层编号节:", len(pit))

    stats = {1: [], 2: [], 3: []}
    plan = []

    for path in sorted(glob.glob(os.path.join(REF, "symptom-*.md"))):
        name = os.path.basename(path)
        lines = load(path).split("\n")
        secs = split_symptom_sections(lines)
        for start, end, raw, title in secs:
            if raw is None:
                continue
            num = top_key(raw)
            if num not in pit:
                print("中止：§%s 在pitfall 侧找不到顶层节（%s）" % (raw, name))
                return 1
            body = "\n".join(lines[start + 1:end])
            pbody = pit[num][2]

            # 幂等：已经处理过的节跳过。判据是「正文里有根因链接」或
            # 「已经有现象导语」—— 二者都是处理后的形态。
            # 少了这个判断，重跑会把已收敛的节当成新节再改一遍。
            if "根因与修法" in body or "**现象**" in body:
                continue

            full = difflib.SequenceMatcher(
                None, norm(body, True), norm(pbody, True)).ratio()
            if full >= TIER1:
                tier = 1
            elif full >= TIER2:
                tier = 2
            else:
                tier = 3
            stats[tier].append((raw, name, full, title))
            plan.append((path, start, end, raw, title, tier, full, num))

    for t in (1, 2, 3):
        print("档%d: %d 节" % (t, len(stats[t])))
    print()

    by_tier = {1: [], 2: [], 3: []}
    for item in plan:
        by_tier[item[5]].append(item)
    for t in (1, 2, 3):
        if not by_tier[t]:
            continue
        print("== 档%d" % t)
        for path, start, end, raw, title, tier, full, num in by_tier[t]:
            print("  §%-6s %-30s full=%.3f%s" % (raw, os.path.basename(path), full, title[:40]))
        print()

    reclaim = 0
    for path, start, end, raw, title, tier, full, num in by_tier[1] + by_tier[2]:
        reclaim += len("\n".join(load(path).split("\n")[start + 1:end]))
    print("档1+档2 可回收字符: %d" % reclaim)
    print()

    if not args.apply:
        print("干跑，未写入。")
        return 0

    # 从后往前改，避免行号漂移。
    for path in sorted({p for p, *_ in plan}):
        items = sorted([x for x in plan if x[0] == path], key=lambda x: -x[1])
        lines = load(path).split("\n")
        for _, start, end, raw, title, tier, full, num in items:
            pfile, ptitle, _pb = pit[num]
            head = lines[start]
            note = '<p align="right"><sub>来源 com-pitfalls §%s</sub></p>' % raw
            if tier == 3:
                # 保留正文，只确保末尾有来源尾注。
                seg = lines[start + 1:end]
                while seg and not seg[-1].strip():
                    seg.pop()
                if not any(SRC_NOTE.search(x.strip()) for x in seg[-3:] if x.strip()):
                    seg += ["", note]
                else:
                    seg += [""]
                new = [head] + seg
            else:
                sub = "" if raw == str(num) else "（根因条目 §%d）" % num
                new = [
                    head,
                    "",
                    "**根因与修法**：见 [`%s`](%s) §%s%s。"
                    % (pfile, pfile, num, sub),
                    "",
                    note,
                    "",
                ]
            lines[start:end] = new
        with open(path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        print("已改写 %s" % os.path.basename(path))
    return 0


if __name__ == "__main__":
    sys.exit(main())
