#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression for windowed picture fills (the single-slide "shape cuts image" recipe).

Four separate bugs were found while building this, and every one of them produced a
file that opened without complaint and rendered the wrong thing. They are the reason
each assertion below exists:

  A. <p:blipFill> instead of <a:blipFill>
     blipFill lives in the DRAWINGML namespace. With the p: prefix it is an unknown
     element and PowerPoint drops it in silence -- the shape falls back to whatever
     fill it had. Same family as p:morph vs p159:morph.
  B. the fill placed AFTER <a:ln>
     CT_ShapeProperties is a sequence: xfrm, geom, fill, ln, effectLst, ... An illegal
     order is dropped in silence.
  C. a leftover <a:solidFill> left beside the new fill
     The fill choices are an xsd:choice, so only one applies and the FIRST wins. The
     shape kept its old colour and the picture never appeared.
  D. an invented r:embed
     PowerPoint paints its missing-image placeholder, which is easy to mistake for
     "the inset was ignored" -- and it is what made an early diagnosis wrong.

Nothing here needs PowerPoint: A/B/C/D are all decidable from the XML, and the rendered
proof lives in showcase/qa/fill_shipped.py which checks three windows show three
different bands of one picture.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))

import motion                                          # noqa: E402

PASS = FAIL = 0


def check(label, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  %-58s ok" % label)
    else:
        FAIL += 1
        print("  %-58s FAIL %s" % (label, detail))


def sp(label, xml):
    return '<p:sp><p:nvSpPr><p:cNvPr id="7" name="%s"/></p:nvSpPr>%s</p:sp>' % (label, xml)


def direct_children(body):
    """Top-level child tags of one element, ignoring nested content."""
    out, depth, i = [], 0, 0
    inner = body[body.index(">") + 1:body.rindex("<")]
    pat = re.compile(r"<(/?)([A-Za-z0-9:]+)([^>]*?)(/?)>")
    while i < len(inner):
        m = pat.search(inner, i)
        if not m:
            break
        closing, tag, _a, selfclose = m.groups()
        if closing:
            depth -= 1
        elif depth == 0:
            out.append(tag)
            if not selfclose:
                depth += 1
        elif not selfclose:
            depth += 1
        i = m.end()
    return out


RANK = {"a:xfrm": 0, "a:prstGeom": 1, "a:custGeom": 1, "a:blipFill": 2,
        "a:solidFill": 2, "a:noFill": 2, "a:gradFill": 2, "a:pattFill": 2,
        "a:grpFill": 2, "a:ln": 3, "a:effectLst": 4, "a:effectDag": 4,
        "a:scene3d": 5, "a:sp3d": 6, "a:extLst": 7}


print("== 1. window_insets 算术 ==")
# a full-bleed picture: the window's own edges must map onto the picture's edges
l, t, r, b = motion.window_insets([380, 0, 200, 540], [0, 0, 960, 540], 960)
check("full-bleed window at x=380 w=200 -> l = r = -190000",
      (l, r) == (-190000, -190000), (l, r))
check("insets are symmetric for a centred window", l == r, (l, r))
l2, t2, r2, b2 = motion.window_insets([0, 0, 320, 540], [0, 0, 960, 540], 960)
check("leftmost window: l = 0, r = -200000", (l2, r2) == (0, -200000), (l2, r2))
# a picture that does NOT span the slide -- the general case the doc's formula omits
l3, t3, r3, b3 = motion.window_insets([400, 100, 200, 200], [300, 0, 400, 400], 960)
check("picture smaller than slide: l = r = -50000", (l3, r3) == (-50000, -50000),
      (l3, r3))
check("vertical insets use the shape HEIGHT as denominator",
      t3 == -50000 and b3 == -50000, (t3, b3))

print("\n== 2. build_fill_window 的命名空间（bug A）==")
frag = motion.build_fill_window({"target": "W", "window": [60, 150, 200, 300],
                                 "picture": [0, 0, 960, 540]}, 960, 540)
check("uses a:blipFill, NOT p:blipFill", "<a:blipFill>" in frag, frag[:60])
check("never emits p:blipFill", "p:blipFill" not in frag, frag[:60])
check("carries a fillRect", "<a:fillRect" in frag, frag[:90])
check("is closed with </a:blipFill>", "</a:blipFill>" in frag)

print("\n== 3. 插入位置（bug B）==")
CASES = {
    "geom then ln": '<p:spPr><a:xfrm/><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                    '<a:ln><a:solidFill/></a:ln></p:spPr>',
    "geom, no ln": '<p:spPr><a:xfrm/><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                   "</p:spPr>",
    "geom + ln + effectLst": '<p:spPr><a:xfrm/><a:prstGeom prst="rect"><a:avLst/>'
                             '</a:prstGeom><a:ln><a:solidFill/></a:ln><a:effectLst/>'
                             "</p:spPr>",
}
for label, body in CASES.items():
    out, st = motion.insert_blip_fill(sp(label, body), 7, frag)
    kids = direct_children(re.search(r"<p:spPr>.*?</p:spPr>", out, re.S).group(0))
    ranks = [RANK.get(k, 99) for k in kids]
    check("%s: order legal (%s)" % (label, ",".join(k.split(":")[-1] for k in kids)),
          ranks == sorted(ranks) and st == "ok", kids)
    if "a:blipFill" in kids:
        check("%s: fill before a:ln" % label,
              "a:ln" not in kids or kids.index("a:blipFill") < kids.index("a:ln"), kids)

print("\n== 4. 已有填充必须被替换（bug C）==")
body = ('<p:spPr><a:xfrm/><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>'
        '<a:ln><a:solidFill/></a:ln></p:spPr>')
out, st = motion.insert_blip_fill(sp("W", body), 7, frag)
inner = re.search(r"<p:spPr>.*?</p:spPr>", out, re.S).group(0)
kids = direct_children(inner)
fills = [k for k in kids if k in ("a:solidFill", "a:blipFill", "a:noFill",
                                  "a:gradFill", "a:pattFill", "a:grpFill")]
check("exactly ONE fill remains", len(fills) == 1, fills)
check("the surviving fill is the new one", fills == ["a:blipFill"], fills)
check("the old colour is gone", "FF00FF" not in inner, inner[:120])
check("the line's own solidFill survived", inner.count("<a:solidFill") == 1,
      inner.count("<a:solidFill"))

print("\n== 5. 幂等与状态 ==")
once, st1 = motion.insert_blip_fill(sp("W", CASES["geom then ln"]), 7, frag)
twice, st2 = motion.insert_blip_fill(once, 7, frag)
check("first insert reports ok", st1 == "ok", st1)
check("second insert refuses", st2 == "already-has-blipFill", st2)
check("second insert did not duplicate", twice.count("<a:blipFill>") == 1,
      twice.count("<a:blipFill>"))
out, st = motion.insert_blip_fill(sp("W", CASES["geom then ln"]), 999, frag)
check("unknown shape id -> shape-not-found", st == "shape-not-found", st)
check("unknown id leaves xml untouched",
      out.count("<a:blipFill>") == 0, out.count("<a:blipFill>"))

print("\n== 6. 参数校验 ==")
for bad, why in (({"target": "W", "window": [1, 2, 3]}, "window with 3 numbers"),
                 ({"target": "W", "window": [1, 2, 0, 4]}, "zero-width window"),
                 ({"target": "W", "window": [1, 2, 4, 0]}, "zero-height window"),
                 ({"target": "W"}, "no window at all")):
    try:
        motion.build_fill_window(bad, 960, 540)
        check("rejects %s" % why, False, "no error raised")
    except ValueError:
        check("rejects %s" % why, True)

print("\n%d passed, %d failed" % (PASS, FAIL))
sys.exit(1 if FAIL else 0)
