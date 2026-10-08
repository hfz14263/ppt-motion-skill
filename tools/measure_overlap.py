"""只读测量：symptom-* 与 pitfall-* 两族按 §n 的正文重叠程度。

不修改任何 reference/ 内容，只输出统计。
"""

import difflib
import glob
import os
import re
import sys

CJK = re.compile(r"[\u4e00-\u9fff]")


def load(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def strip_code(text):
    return re.sub(r"```.*?```", "", text, flags=re.S)


def sections(text):
    """切节并抽取稳定编号。

    两族标题写法不同，必须都认：
      pitfall-*.md  '## 12. xxx'        -> 编号 12
      symptom-*.md  '### §12 xxx'      -> 编号 12
      子编号         '### §43.2 xxx' / '### 20.1 xxx' -> 编号 43.2 / 20.1
    返回 [(编号或 None, 标题, 正文)]。
    """
    out = []
    cur_id = None
    cur_title = None
    cur_lines = []
    seen = False
    for ln in text.split("\n"):
        sid = None
        title = None
        m_sym = re.match(r"^#{2,4} §?(\d+(?:\.\d+)*)\.?\s+(.*)$", ln)
        if m_sym:
            sid, title = m_sym.group(1), m_sym.group(2)
        else:
            m_h = re.match(r"^(#{2,4}) (.+)$", ln)
            if m_h:
                # 无编号标题（`### 症状` / `### 规则`）是上一编号的子节，
                # 必须继承编号，不能把编号清成 None —— 否则正文会被判成空。
                title = m_h.group(2)
        if title is not None:
            if seen:
                out.append((cur_id, cur_title, "\n".join(cur_lines)))
            seen = True
            if sid is not None:
                cur_id = sid
            cur_title = title
            cur_lines = []
        else:
            cur_lines.append(ln)
    if seen:
        out.append((cur_id, cur_title, "\n".join(cur_lines)))
    return out


def num_key(sid):
    """'43.2' -> (43, 2)；'43' -> (43, 0)。用于排序与子编号归并。"""
    parts = [int(x) for x in sid.split(".")]
    while len(parts) < 2:
        parts.append(0)
    return (parts[0], parts[1])


def collect(pattern, skip=()):
    table = {}
    for path in sorted(glob.glob(pattern)):
        if os.path.basename(path) in skip:
            continue
        for sid, title, body in sections(load(path)):
            if sid is not None:
                table.setdefault(num_key(sid), []).append((path, sid, title, body))
    return table


def norm(text):
    text = strip_code(text)
    text = re.sub(r"<sub>.*?</sub>", "", text, flags=re.S)
    return re.sub(r"\s+", "", text)


def norm_full(text):
    """保留代码块的完整归一化文本。"""
    text = re.sub(r"<sub>.*?</sub>", "", text, flags=re.S)
    return re.sub(r"\s+", "", text)


def ratio(a, b):
    """序列相似度。对代码块也敏感，比 Jaccard 稳。"""
    if not a or not b:
        return 0.0
    return difflib.SequenceMatcher(None, a, b).ratio()


def prose_len(text):
    """去掉代码块与尾注之后的散文长度。用于识别「这节正文全是代码」。"""
    return len(norm(text))


def main():
    pit = collect("reference/pitfall-*.md", skip=("pitfall-map.md",))
    sym = collect("reference/symptom-*.md")

    print("pitfall 有编号节:", len(pit))
    print("symptom 有编号节:", len(sym))
    print("symptom 有 / pitfall 无:", sorted(set(sym) - set(pit)))
    print("pitfall 有 / symptom 无:", sorted(set(pit) - set(sym)))
    print()

    exact = []
    near = []
    low = []
    codey = []
    total = 0
    reclaim = 0
    for sid in sorted(sym):
        for spath, sid_raw, stitle, sbody in sym[sid]:
            ns = norm(sbody)
            nf = norm_full(sbody)
            total += len(sbody)
            if prose_len(sbody) < 40 and len(nf) < 200:
                codey.append((sid_raw, os.path.basename(spath), prose_len(sbody), len(sbody)))
                continue
            best = None
            best_full = 0.0
            best_prose = 0.0
            for ppath, pid_raw, ptitle, pbody in pit.get(sid, []):
                np_ = norm(pbody)
                pf_ = norm_full(pbody)
                if ns and ns == np_:
                    best = (ppath, pid_raw)
                    best_full = 1.0
                    best_prose = 1.0
                    break
                rf = ratio(nf, pf_)
                if rf > best_full:
                    best = (ppath, pid_raw)
                    best_full = rf
                    best_prose = ratio(ns, np_)
            praw = best[1] if best else "-"
            pname = os.path.basename(best[0]) if best else "-"
            row = (sid_raw, os.path.basename(spath), praw, pname,
                   best_full, best_prose, len(sbody))
            if best_full >= 0.999:
                exact.append(row)
                reclaim += len(sbody)
            elif best_full >= 0.80:
                near.append(row)
            else:
                low.append(row)

    print("完全一致(含代码块)节数:", len(exact))
    print("高相似(>=0.80)非完全一致:", len(near))
    print("低相似(<0.80)需人工:", len(low))
    print("symptom 侧正文极短(现象侧独有形态):", len(codey))
    print("symptom 编号节正文字符(未去代码块):", total)
    print("完全一致可回收字符:", reclaim)
    print()
    print("口径说明: full=含代码块序列相似度, prose=仅散文相似度")
    print()
    for tag, rows in (("NEAR", near), ("LOW", low)):
        if not rows:
            continue
        print("==", tag)
        for sid, sname, praw, pname, rf, rp, size in rows:
            print("  §%-6s %-26s -> §%-6s %-26s full=%.3f prose=%.3f len=%d"
                  % (sid, sname, praw, pname, rf, rp, size))
        print()
    if codey:
        print("== 正文极短（现象描述 + 代码，主体在代码块里）")
        for sid, sname, plen, size in codey:
            print("  §%-6s %-26s prose=%-5d raw=%d" % (sid, sname, plen, size))
        print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
