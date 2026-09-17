#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: spec coverage audit.

The other gates (apply --assert-geometry, verify_motion, motion.ps1 -Strict) prove
the OOXML is legal, the geometry did not move, and PowerPoint kept every effect.
They CANNOT see the failure this catches:

    A shape with no effect is visible from the moment the slide appears.

It never disappears and nothing errors -- it just sits there while everything else
animates in. Users report it as "this column keeps showing up and never animates".
This was the single most common defect in real use: a hand-written spec silently
lost an entire table column, a whole slide column, and a whole card because shape
ids had been reassigned by a later layout edit.

Also reports:
  * targets that duplicate a shape (one effect per shape -- com-pitfalls 14), and
  * targets that do not exist on their slide.

Run this on every spec. A non-zero exit means the deck is not ready to present.
"""
import argparse
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import motion  # noqa: E402
import player  # noqa: E402

# Shapes that are legitimately static on every slide: the footer credit and the
# page number. These are defaults for convenience -- pass `--footer` with your own
# text, or an empty value to disable the text match entirely. The list is
# deliberately generic: a tool should not hard-code one project's wording.
DEFAULT_STATIC = ("汇报", "Report", "Confidential", "Confidential - Do Not Distribute")


def is_static(sh, footer_texts, bottom_y=None, right_x=None):
    """Footer credit and page number are deliberately static.

    The page number sits at the bottom-RIGHT and starts *above* the very bottom
    strip (506pt on a 540pt slide), so a bottom-only test misses it. Match on
    position AND content: a bare number in the bottom-right corner.
    """
    t = (sh.get("text") or "").strip()
    if any(k and k in t for k in footer_texts):
        return True
    if t.isdigit():
        pt = sh.get("pt")
        if pt and bottom_y is not None and right_x is not None:
            if pt[1] >= bottom_y and pt[0] >= right_x:
                return True
    return False


def main(argv=None):
    ap = argparse.ArgumentParser(description="Audit motion spec coverage")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--footer", action="append", default=None,
                    help="extra footer/static text marker (repeatable)")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    spec = motion.load_spec(ns.spec)
    entries = {e["page"]: e for e in motion.normalize_spec(spec)}
    footer_texts = tuple(ns.footer) if ns.footer else DEFAULT_STATIC

    # page-number zone: bottom strip, right-hand side (points)
    try:
        w, h = motion.slide_size(ns.pptx)
    except Exception:
        w, h = 960.0, 540.0
    bottom_y, right_x = h * 0.90, w * 0.85

    result = {"missing": {}, "duplicate": {}, "unknown": {}, "ok": True}
    for page in sorted(entries):
        shapes = player.shape_index(ns.pptx, page)
        covered, dupes = set(), []
        for eff in entries[page]["effects"]:
            t = str(eff.get("target"))
            if t in covered:
                dupes.append(t)
            covered.add(t)
        ids = {sh["id"] for sh in shapes}
        unknown = sorted(t for t in covered if t not in ids)

        missing = []
        for sh in shapes:
            if sh["id"] in covered:
                continue
            if is_static(sh, footer_texts, bottom_y, right_x):
                continue
            missing.append({"id": sh["id"],
                            "label": (sh["text"][:40] or sh["name"]),
                            "pt": sh["pt"]})
        if missing:
            result["missing"][str(page)] = missing
        if dupes:
            result["duplicate"][str(page)] = sorted(set(dupes))
        if unknown:
            result["unknown"][str(page)] = unknown

    result["ok"] = not (result["missing"] or result["duplicate"] or result["unknown"])

    if ns.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
        return 0 if result["ok"] else 1

    for page, items in sorted(result["missing"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s %d shape(s) with NO effect:" % (page, len(items)))
        for it in items:
            print("      id=%-5s %-44s %s" % (it["id"], it["label"], it["pt"]))
    for page, ts in sorted(result["duplicate"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s duplicate effect on: %s" % (page, ", ".join(ts)))
    for page, ts in sorted(result["unknown"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s target(s) not on slide: %s" % (page, ", ".join(ts)))

    if result["ok"]:
        n = sum(len(e["effects"]) for e in entries.values())
        print("OK: %d slide(s), %d effect(s); every non-footer shape has exactly "
              "one effect and every target exists" % (len(entries), n))
        return 0
    print("\nFAILED: a shape with no effect is visible before the slide's animation "
          "starts.\nFix the spec (or the generator that produced it) and re-run.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
