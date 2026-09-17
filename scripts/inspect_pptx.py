#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""inspect_pptx: is this .pptx actually animated, and how?

WHY THIS EXISTS
---------------
When you go looking for well-animated decks to learn from, most of what you find is
NOT animated, and the labels actively mislead:

  * A Canva "animated template" exported to .pptx is STATIC. Canva's motion is a
    renderer feature; the export writes positions and nothing else.
  * A screen recording / tutorial video has no animation data at all -- you can see
    the motion but cannot read a single timing value out of it.
  * A template site's "animated" category usually means "has slide transitions" or
    "has an animated GIF pasted on it", not "has entrance/emphasis effects".
  * Even a repo with a folder literally named "Carousel Animation" measured ZERO
    effects -- only transitions.

So the only reliable answer comes from the package itself. This script opens the
.pptx, reads `<p:timing>`, and reports exactly what is in there: how many effects,
of which class, with what durations and delays, on which slides.

A useful rule of thumb from probing real files: **a file is worth studying if at
least one third of its slides carry `<p:timing>`**. Transitions alone are page
decoration; timing is design.

Usage:
    python inspect_pptx.py deck.pptx
    python inspect_pptx.py *.pptx --json
    python inspect_pptx.py folder/ --recursive
"""
import argparse
import collections
import glob
import json
import os
import re
import sys
import zipfile

CLASSES = {"entr": "entrance", "emph": "emphasis", "path": "motion-path",
           "exit": "exit", "mediacall": "media", "": "unknown"}


def slides_of(z):
    return sorted((n for n in z.namelist()
                   if re.match(r"ppt/slides/slide\d+\.xml$", n)),
                  key=lambda n: int(re.search(r"(\d+)", n.split("/")[-1]).group(1)))


def analyse(path):
    r = {"file": os.path.basename(path), "path": os.path.abspath(path),
         "bytes": os.path.getsize(path), "slides": 0, "animated_slides": 0,
         "effects": 0, "by_class": {}, "transitions": 0, "motion_paths": 0,
         "media": 0, "durations_ms": [], "delays_ms": [], "triggers": {},
         "notes": 0, "verdict": "", "per_slide": []}
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "ppt/presentation.xml" not in names:
                r["verdict"] = "NOT A PPTX (no ppt/presentation.xml)"
                return r
            slides = slides_of(z)
            r["slides"] = len(slides)
            r["media"] = len([n for n in names if n.startswith("ppt/media/")])
            r["notes"] = len([n for n in names if re.match(r"ppt/notesSlides/", n)])
            for n in slides:
                x = z.read(n).decode("utf-8", "replace")
                idx = int(re.search(r"(\d+)", n.split("/")[-1]).group(1))
                eff = trans = paths = 0
                cls_here = collections.Counter()
                if "<p:timing" in x:
                    r["animated_slides"] += 1
                for cls in re.findall(r'presetClass="(\w+)"', x):
                    cls_here[cls] += 1
                    r["by_class"][cls] = r["by_class"].get(cls, 0) + 1
                    eff += 1
                trans = len(re.findall(r"<p:transition\b", x))
                paths = len(re.findall(r"<p:animMotion\b", x))
                r["effects"] += eff
                r["transitions"] += trans
                r["motion_paths"] += paths
                r["durations_ms"] += [int(v) for v in
                                      re.findall(r'<p:cTn\b[^>]*\bdur="(\d+)"', x)
                                      if v not in ("0", "1")]
                r["delays_ms"] += [int(v) for v in
                                   re.findall(r'<p:cond\b[^>]*\bdelay="(\d+)"', x)]
                for t in re.findall(r'nodeType="(\w+)"', x):
                    r["triggers"][t] = r["triggers"].get(t, 0) + 1
                r["per_slide"].append({"slide": idx, "effects": eff,
                                       "by_class": dict(cls_here),
                                       "transitions": trans, "motion_paths": paths})
    except zipfile.BadZipFile:
        r["verdict"] = "CORRUPT (not a zip)"
        return r

    # The verdict is the whole point: "transitions only" is the trap, because a
    # template site's "animated" badge means exactly that.
    if r["slides"] == 0:
        r["verdict"] = "EMPTY (no slides)"
    elif r["effects"] == 0 and r["transitions"] == 0:
        r["verdict"] = "STATIC"
    elif r["effects"] == 0:
        r["verdict"] = "TRANSITIONS ONLY - not worth studying for motion design"
    elif r["animated_slides"] / float(r["slides"]) < 0.34:
        r["verdict"] = ("LIGHTLY ANIMATED (%d/%d slides) - probably a one-off "
                        "effect" % (r["animated_slides"], r["slides"]))
    else:
        r["verdict"] = "ANIMATED - worth studying (%d effects over %d/%d slides)" % (
            r["effects"], r["animated_slides"], r["slides"])
    return r


def pct(vals, q):
    if not vals:
        return None
    s = sorted(vals)
    return s[min(len(s) - 1, int(len(s) * q))]


def show(r):
    print("=" * 74)
    print(r["file"])
    print("  %s" % r["verdict"])
    if r.get("slides"):
        print("  slides %d | animated %d | effects %d | transitions %d | "
              "media %d | notes %d"
              % (r["slides"], r["animated_slides"], r["effects"],
                 r["transitions"], r["media"], r["notes"]))
    if r.get("by_class"):
        print("  by class: %s" % ", ".join(
            "%s=%d" % (CLASSES.get(k, k), v) for k, v in
            sorted(r["by_class"].items(), key=lambda kv: -kv[1])))
    if r.get("durations_ms"):
        print("  effect duration ms: min %d / median %d / max %d  (n=%d)"
              % (min(r["durations_ms"]), pct(r["durations_ms"], 0.5),
                 max(r["durations_ms"]), len(r["durations_ms"])))
    if r.get("delays_ms"):
        nonzero = [d for d in r["delays_ms"] if d]
        if nonzero:
            print("  stagger delay ms: median %d / max %d  (n=%d nonzero)"
                  % (pct(nonzero, 0.5), max(nonzero), len(nonzero)))
    if r.get("triggers"):
        print("  triggers: %s" % r["triggers"])
    if r.get("per_slide") and r["slides"] > 1:
        rich = [s for s in r["per_slide"] if s["effects"]]
        if rich:
            print("  busiest slides: %s" % ", ".join(
                "#%d(%d)" % (s["slide"], s["effects"])
                for s in sorted(rich, key=lambda s: -s["effects"])[:6]))
    print()


def main():
    ap = argparse.ArgumentParser(description="Is this pptx actually animated?")
    ap.add_argument("targets", nargs="+", help="pptx file(s), glob(s) or folder(s)")
    ap.add_argument("--recursive", "-r", action="store_true",
                    help="descend into folders")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    ns = ap.parse_args()

    paths = []
    for t in ns.targets:
        if os.path.isdir(t):
            pat = os.path.join(t, "**", "*.pptx") if ns.recursive \
                else os.path.join(t, "*.pptx")
            paths += glob.glob(pat, recursive=ns.recursive)
        else:
            paths += glob.glob(t) or [t]
    paths = sorted(set(p for p in paths if os.path.isfile(p)))
    if not paths:
        print("no .pptx found")
        return 1

    results = [analyse(p) for p in paths]
    if ns.json:
        print(json.dumps(results, ensure_ascii=False, indent=1))
    else:
        for r in results:
            show(r)
        ok = [r for r in results if r["verdict"].startswith("ANIMATED")]
        print("%d/%d file(s) are animated enough to learn from" % (len(ok), len(results)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
