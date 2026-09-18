#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression for the duplicate-singleton check.

The check exists because EVERY corrupted file this project produced came from one cause:
an element inserted into a position its sequence type already held. PowerPoint's only
complaint is "file or directory is corrupt" (0x80070570), with no hint which element.

Three bugs were made while BUILDING this check, and each is asserted below, because a
validator that cannot fail is worse than no validator:

  1. motion.element_spans prefixes the namespace itself (it searches "<p:" + tag), so
     callers pass the BARE name. Passing "p:spPr" searched for "<p:p:spPr", matched
     nothing, and the validator reported OK on a file PowerPoint rejected.
  2. the direct-child scan checked depth == 1. Inside spPr the direct children sit at
     depth 0, so it never matched anything.
  3. counting "<p:transition" globally flagged a CORRECT morph as a duplicate, because
     morph legitimately carries two transitions -- one in mc:Choice, one in mc:Fallback.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))
import motion                                             # noqa: E402

PASS = FAIL = 0


def check(label, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print("  %-62s ok" % label)
    else:
        FAIL += 1
        print("  %-62s FAIL %s" % (label, detail))


def wrap(sp_pr_body):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="7" name="P"/></p:nvSpPr>'
            "<p:spPr>" + sp_pr_body + "</p:spPr></p:sp>")


print("== 1. element_spans 的前缀约定（bug 1）==")
X = wrap('<a:xfrm/><a:ln><a:noFill/></a:ln>')
check("bare name 'spPr' matches", len(motion.element_spans(X, "spPr")) == 1,
      motion.element_spans(X, "spPr"))
check("prefixed name 'p:spPr' matches NOTHING (the trap)",
      motion.element_spans(X, "p:spPr") == [])

print("\n== 2. 重复检测：直接子元素（bug 2）==")
BROKEN = wrap('<a:xfrm/><a:prstGeom prst="pie"><a:avLst/></a:prstGeom>'
              "<a:ln><a:noFill/></a:ln>"
              '<a:effectLst/><a:effectLst><a:innerShdw/></a:effectLst>')
dups = motion.find_duplicate_singletons(BROKEN, "spPr")
check("two effectLst are reported", dups == [("a:effectLst", 2)], dups)

GOOD = wrap('<a:xfrm/><a:prstGeom prst="pie"><a:avLst/></a:prstGeom>'
            '<a:ln><a:noFill/></a:ln><a:effectLst><a:innerShdw/></a:effectLst>')
check("one effectLst is clean", motion.find_duplicate_singletons(GOOD, "spPr") == [])

NESTED = wrap('<a:xfrm/><a:effectLst><a:innerShdw/></a:effectLst>'
              '<a:sp3d><a:bevelT/></a:sp3d>')
check("nested children are not miscounted as siblings",
      motion.find_duplicate_singletons(NESTED, "spPr") == [])

TWO_FILLS = wrap('<a:xfrm/><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                 "<a:solidFill/><a:blipFill/>")
dups2 = motion.find_duplicate_singletons(TWO_FILLS, "spPr",
                                         ("a:solidFill", "a:blipFill"))
check("solidFill beside blipFill is reported (xsd:choice)",
      sorted(t for t, _ in dups2) == ["a:blipFill", "a:solidFill"], dups2)

print("\n== 3. morph 的两个 transition 不算重复（bug 3）==")
# morph is mc:AlternateContent > mc:Choice > p:transition, and mc:Fallback > p:transition
SLIDE = ('<p:sld><p:cSld/><p:clrMapOvr/>'
         '<mc:AlternateContent><mc:Choice Requires="p159">'
         '<p:transition spd="slow"><p159:morph option="byObject"/></p:transition>'
         "</mc:Choice><mc:Fallback>"
         '<p:transition spd="slow"><p:fade/></p:transition>'
         "</mc:Fallback></mc:AlternateContent></p:sld>")
check("a correct morph reports no duplicate",
      motion.find_duplicate_singletons(SLIDE, "sld",
                                      ("p:cSld", "p:clrMapOvr", "p:timing",
                                       "p:transition")) == [])

TWO_TRANS = ('<p:sld><p:cSld/><p:clrMapOvr/>'
             '<p:transition><p:fade/></p:transition>'
             '<p:transition><p:fade/></p:transition></p:sld>')
check("two direct transitions ARE reported",
      motion.find_duplicate_singletons(TWO_TRANS, "sld", ("p:transition",))
      == [("p:transition", 2)])

print("\n== 4. set_singleton：替换而不是插入 ==")
body = '<a:xfrm/><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:effectLst/>'
out, st = motion.set_singleton(body, "a:effectLst",
                               '<a:effectLst><a:innerShdw/></a:effectLst>',
                               before=("a:scene3d", "a:extLst"))
check("empty self-closing effectLst is REPLACED", out.count("<a:effectLst") == 1, out)
check("the shadow is present", "innerShdw" in out)
check("status reports replaced", st == "replaced", st)

out2, st2 = motion.set_singleton('<a:xfrm/>', "a:effectLst",
                                 "<a:effectLst><a:innerShdw/></a:effectLst>",
                                 before=("a:extLst",))
check("absent effectLst is INSERTED", out2.count("<a:effectLst") == 1, out2)

out3, st3 = motion.set_singleton(BROKEN, "a:effectLst",
                                 "<a:effectLst><a:innerShdw/></a:effectLst>",
                                 inside="spPr")
# the result carries exactly one effectLst element, in its paired form
check("duplicates in a real shape are collapsed to one",
      len(motion.singleton_spans(out3, "a:effectLst")) == 1, out3)
check("the empty self-closing copy is gone",
      "<a:effectLst/>" not in out3, out3)
check("no duplicate reported after the fix",
      motion.find_duplicate_singletons(out3, "spPr") == [],
      motion.find_duplicate_singletons(out3, "spPr"))

print("\n%d passed, %d failed" % (PASS, FAIL))
sys.exit(1 if FAIL else 0)
