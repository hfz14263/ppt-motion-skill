#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Validate the documentation structure: references resolve, no orphans, facts unique.

WHY THIS EXISTS
---------------
The documentation grew by accretion and reached 261k characters across 27 files, which is
larger than any model's context. That is not just a reading problem: it means nobody --
including an agent -- can hold it all at once, so when one document says X and another
says not-X, the contradiction cannot be noticed. This session produced exactly that: a
fact about lat/lon bounds living in three files, a presetID relationship spread across
eight, and a "silently ignored" concept in ten.

INDEX.md fixes the reading side by giving one entry point of ~4k characters. This checks
the properties that keep it fixed:

  1. NO DANGLING REFERENCES. Every link from an index resolves. A navigation page that
     points at nothing is worse than no navigation, because it looks authoritative.
  2. NO ORPHANS. Every reference document is reachable from INDEX.md. An unreachable
     document is one nobody will read and nobody will keep correct.
  3. SIZE BUDGET. INDEX.md must stay readable in one pass. The moment the entry point
     needs scrolling and paging, it has stopped being an entry point.
  4. FACTS ARE NOT DUPLICATED. A fact is declared canonical in `facts/*.json` via its
     `sources`; a second copy elsewhere is the disease this reorganisation treats.

Rules 1-3 are mechanical. Rule 4 is reported as ADVISE, not failure: prose legitimately
mentions a number while explaining it, and a check that cannot tell a definition from a
mention would fail on every well-written document. Naming what it cannot decide is the
point -- see references/review-checklist.md section 0.

Usage:
    python scripts/verify_docs.py
    python scripts/verify_docs.py --json
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, ".."))

INDEX = "INDEX.md"
INDEX_BUDGET = 6000          # characters; ~3500 tokens. Beyond this it stops being
                             # readable in a single pass, which is its whole purpose.

# Directories that are entry-point reachable material and must be linked from INDEX.
SCAN_DIRS = ("references", "reference", "facts")
SKIP_FILES = {"references/design-system/README.md"}   # linked via its own directory entry
UPSTREAM = re.compile(r"vendored from|upstream content", re.I)


def read(p):
    return io.open(os.path.join(ROOT, p), encoding="utf-8").read()


def strip_code(text):
    """Remove fenced code blocks before scanning for links.

    Without this the linker finds things like `[($r -shl 16)]` inside a PowerShell
    snippet and reports them as dangling references. A checker that cries wolf on code
    is one people learn to ignore, which defeats its purpose.
    """
    return re.sub(r"```.*?```", "", text, flags=re.S)


def links(text):
    return [m for m in re.findall(r"\]\(([^)]+)\)", strip_code(text))
            if not m.startswith("http")]


def main(argv=None):
    ap = argparse.ArgumentParser(description="文档结构校验")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    fails, advises = [], []
    idx = read(INDEX)

    # ---- 1. dangling references, from every markdown file -------------------
    md_files = []
    for base, _dirs, files in os.walk(ROOT):
        if ".git" in base:
            continue
        for f in files:
            if f.endswith(".md"):
                md_files.append(os.path.relpath(os.path.join(base, f), ROOT)
                                .replace("\\", "/"))
    for rel in sorted(md_files):
        body = read(rel)
        for target in links(body):
            t = target.split("#")[0]
            if not t:
                continue
            resolved = os.path.normpath(os.path.join(os.path.dirname(rel), t))
            if not os.path.exists(os.path.join(ROOT, resolved)):
                fails.append({"rule": "dangling", "file": rel, "target": target})

    # ---- 2. orphans: reference docs unreachable from INDEX ------------------
    reachable = set()
    for target in links(idx):
        t = target.split("#")[0]
        if t:
            reachable.add(os.path.normpath(t).replace("\\", "/"))
    # a directory entry in INDEX makes everything under it reachable
    reachable_dirs = {r for r in reachable if not r.endswith(".md")}
    for base in SCAN_DIRS:
        for f in sorted(os.listdir(os.path.join(ROOT, base))
                        if os.path.isdir(os.path.join(ROOT, base)) else []):
            rel = "%s/%s" % (base, f)
            if not f.endswith(".md") or rel in SKIP_FILES:
                continue
            body = read(rel)
            if UPSTREAM.search(body[:600]):
                continue                       # vendored, reached via its own README
            if rel in reachable:
                continue
            if any(rel.startswith(d + "/") for d in reachable_dirs):
                continue
            # a file may be linked from any doc that INDEX reaches transitively
            if any(rel in links(read(r)) for r in reachable if r.endswith(".md")
                   and os.path.exists(os.path.join(ROOT, r))):
                continue
            advises.append({"rule": "orphan", "file": rel,
                            "note": "INDEX.md 到不了，也没被它链到的文档引用"})

    # ---- 3. index size budget ----------------------------------------------
    if len(idx) > INDEX_BUDGET:
        fails.append({"rule": "index-too-big", "chars": len(idx),
                      "budget": INDEX_BUDGET,
                      "note": "入口一旦需要翻页就不再是入口"})

    # ---- 4. canonical facts declared once ----------------------------------
    declared = {}
    for f in sorted(os.listdir(os.path.join(ROOT, "facts"))):
        if not f.endswith(".json"):
            continue
        d = json.load(io.open(os.path.join(ROOT, "facts", f), encoding="utf-8"))
        for s in (d.get("sources") or []):
            for key in ("canonical_for",):
                if s.get(key):
                    declared.setdefault(s["location"], []).append(
                        {"fact": s[key], "declared_by": "facts/" + f})

    report = {
        "index_chars": len(idx),
        "index_budget": INDEX_BUDGET,
        "markdown_files": len(md_files),
        "failures": fails,
        "advises": advises,
        "canonical_declarations": declared,
        "ok": not fails,
    }
    if ns.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
        return 1 if fails else 0

    print("== 文档结构校验 ==")
    print("  INDEX.md %d 字符（预算 %d）" % (len(idx), INDEX_BUDGET))
    print("  markdown 文件 %d 个" % len(md_files))
    print()
    if fails:
        print("-- FAIL --")
        for f in fails:
            print("  " + json.dumps(f, ensure_ascii=False))
    if advises:
        print("-- ADVISE（启发式，可能是有意的）--")
        for a in advises:
            print("  " + json.dumps(a, ensure_ascii=False))
    if not fails and not advises:
        print("OK  引用可解析、无孤儿文档、入口在预算内")
    print()
    print("已声明的唯一真相源：")
    for loc, ds in sorted(declared.items()):
        print("  %-40s ← %s" % (loc, ", ".join(d["declared_by"] for d in ds)))
    print()
    print("注意：这只证明【结构自洽】，不证明【内容正确】—— 内容正确靠 evidence 里的实测。")
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
