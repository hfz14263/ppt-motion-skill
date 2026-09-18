#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate the recipe library: structure, and the rule that keeps it from rotting.

A principle library fails in one specific way: it turns back into a list of effects.
That happens when every new trick gets its own entry, and when entries carry no evidence
so nobody can tell a tested mechanism from a plausible one.

So the check enforces two rules:
  1. every entry has all six fields, and `breaks_how` entries look like SYMPTOMS
     (what you see when it is wrong), not restatements of the mechanism
  2. every entry cites evidence, and the evidence names something that exists in this
     repo -- a test file or a measurements file. An assertion with no artefact behind it
     is a rumour.

Usage:
    python scripts/verify_recipes.py
    python scripts/verify_recipes.py --json
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))
LIB = os.path.join(ROOT, "references", "recipes.json")

REQUIRED = ("id", "principle", "one_line", "mechanism", "constraints", "breaks_how",
            "evidence")

# wording that means the author restated the mechanism instead of naming a symptom
NOT_A_SYMPTOM = (r"^注意", r"^记得", r"^需要", r"^必须", r"^应当", r"^确保",
                 r"^remember", r"^note", r"^ensure", r"^must")
# wording that suggests a real observed failure
SYMPTOM_HINT = ("→", "读作", "看起来", "变成", "出现", "损坏", "静默", "退回",
                "失败", "空白", "读成", "silently", "corrupt", "looks", "reads")


def main(argv=None):
    ap = argparse.ArgumentParser(description="配方库校验")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    data = json.load(io.open(LIB, encoding="utf-8"))
    entries = data.get("entries") or []
    problems = []

    seen_ids = set()
    for e in entries:
        eid = e.get("id", "?")
        for f in REQUIRED:
            if not e.get(f):
                problems.append("%s: 缺字段 %s" % (eid, f))
        if eid in seen_ids:
            problems.append("%s: id 重复" % eid)
        seen_ids.add(eid)

        # breaks_how must name symptoms
        for b in (e.get("breaks_how") or []):
            if any(re.match(p, b) for p in NOT_A_SYMPTOM):
                problems.append("%s: breaks_how 像是机制复述而非症状 -> %s"
                                % (eid, b[:40]))
            elif not any(h in b for h in SYMPTOM_HINT):
                problems.append("%s: breaks_how 未描述可见症状 -> %s" % (eid, b[:40]))

        # every cited artefact must exist
        ev = e.get("evidence", "")
        cited = re.findall(r"(?:tests|scripts|references)/[A-Za-z0-9_./-]+", ev)
        cited += re.findall(r"(?:building|showcase|template)/[A-Za-z0-9_./-]+", ev)
        for c in cited:
            if not os.path.exists(os.path.join(ROOT, c)):
                # files may live outside the skill repo; say so rather than fail hard
                alt = os.path.abspath(os.path.join(ROOT, "..", c))
                if not os.path.exists(alt):
                    problems.append("%s: evidence 引用的文件不存在 -> %s" % (eid, c))

    if ns.json:
        print(json.dumps({"entries": len(entries), "problems": problems,
                          "ok": not problems}, ensure_ascii=False, indent=1))
        return 1 if problems else 0

    print("== 配方库校验 ==")
    print("条目 %d 条，字段齐全 %s" % (len(entries), "是" if not problems else "见下"))
    print()
    for e in entries:
        print("  %-34s %s" % (e["id"], e["principle"]))
    print()
    if problems:
        print("-- 问题 --")
        for p in problems:
            print("  " + p)
        print()
        print("规则：breaks_how 要写【坏掉时的可见症状】，不是机制复述；")
        print("      evidence 要指向真实存在的文件，否则是传闻。")
        return 1
    print("OK  结构完整，症状描述具体，证据可追溯")
    print("注意：这只证明【库是自洽的】，不证明【配方是对的】 —— "
          "对错要靠 evidence 里那些实测。")
    return 0


if __name__ == "__main__":
    sys.exit(main())
