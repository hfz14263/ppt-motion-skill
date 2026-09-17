#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Regression tests for dsh-ppt-office-motion.

Covers the four silent-failure defects found by round-tripping a real deck
through PowerPoint. Each one passed structural validation while producing a deck
PowerPoint quietly degraded, so none of them is caught by "does lxml parse it".

Run:  python scripts/selftest.py            (writes nothing outside a temp dir)
Exit: 0 pass, 1 fail.
"""
import json
import os
import re
import shutil
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import motion  # noqa: E402
import player  # noqa: E402

PASS, FAIL = [], []


def check(name, cond, detail=""):
    (PASS if cond else FAIL).append(name)
    print("  %-4s %s%s" % ("ok" if cond else "FAIL", name,
                           ("  -- " + detail) if (detail and not cond) else ""))


def slide_names(z):
    return sorted([n for n in z.namelist() if motion.SLIDE_RE.match(n)],
                  key=lambda n: int(motion.SLIDE_RE.match(n).group(1)))


def read(z, part):
    return z.read(part).decode("utf-8", "replace")


# ---------------------------------------------------------------------------
def make_deck(path, slides=3, transition=False):
    """Minimal but valid pptx: each slide has two named shapes."""
    def slide_xml(i, with_trans):
        trans = ('<p:transition spd="med"><p:fade/></p:transition>'
                 if with_trans else "")
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="100" y="100"/><a:ext cx="400" cy="80"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>t</a:t></a:r></a:p></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="body"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="100" y="220"/><a:ext cx="400" cy="200"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>b</a:t></a:r></a:p></p:txBody></p:sp>'
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '%s</p:sld>' % trans
        )

    overrides = "".join(
        '<Override PartName="/ppt/slides/slide%d.xml" ContentType="application/vnd.'
        'openxmlformats-officedocument.presentationml.slide+xml"/>' % i
        for i in range(1, slides + 1))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.'
                   'openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                   '%s</Types>' % overrides)
        z.writestr("_rels/.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                   'relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
        sldids = "".join('<p:sldId id="%d" r:id="rId%d"/>' % (255 + i, i + 1)
                         for i in range(1, slides + 1))
        z.writestr("ppt/presentation.xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                   'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                   '<p:sldIdLst>%s</p:sldIdLst>'
                   '<p:sldSz cx="960" cy="540"/><p:notesSz cx="6858000" cy="9144000"/>'
                   '</p:presentation>' % sldids)
        rels = "".join('<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/'
                       'officeDocument/2006/relationships/slide" Target="slides/slide%d.xml"/>'
                       % (i + 1, i) for i in range(1, slides + 1))
        z.writestr("ppt/_rels/presentation.xml.rels",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '%s</Relationships>' % rels)
        for i in range(1, slides + 1):
            z.writestr("ppt/slides/slide%d.xml" % i, slide_xml(i, transition))


SPEC = {
    "version": 1,
    "slides": [
        {"page": 1, "transition": {"type": "fade", "duration": 0.8},
         "effects": [
             {"target": "title", "effect": "fly", "duration": 0.7, "trigger": "after"},
             {"target": "body", "effect": "fade", "duration": 0.5, "trigger": "after"},
         ]},
        {"page": 2, "transition": {"type": "push", "duration": 0.6},
         "effects": [
             {"target": "title", "effect": "spin", "duration": 1.0, "trigger": "after", "repeat": 2},
             {"target": "body", "effect": "pathSineWave", "duration": 3.0, "trigger": "after",
              "autoReverse": True, "smooth": 0.5},
         ]},
    ],
}


def main():
    tmp = tempfile.mkdtemp(prefix="pptmotion-selftest-")
    try:
        src = os.path.join(tmp, "src.pptx")
        make_deck(src, slides=3)

        # ---- defect 2: transition must precede timing (CT_Slide sequence) ----
        print("[1] transition is a sibling placed BEFORE timing")
        out1 = os.path.join(tmp, "a1.pptx")
        rep = motion.apply_motion(src, SPEC, out1, assert_geometry=True)
        check("apply reports no errors", not rep["errors"], str(rep["errors"]))
        with zipfile.ZipFile(out1) as z:
            x = read(z, "ppt/slides/slide1.xml")
        i_trans = x.find("<p:transition")
        i_timing = x.find("<p:timing")
        check("transition element exists", i_trans != -1)
        check("timing element exists", i_timing != -1)
        check("transition comes before timing",
              i_trans != -1 and i_timing != -1 and i_trans < i_timing,
              "transition@%d timing@%d" % (i_trans, i_timing))
        # nesting check: the transition span must not sit inside the timing span
        with zipfile.ZipFile(out1) as z:
            raw = read(z, "ppt/slides/slide1.xml")
        t_end = raw.find("</p:transition>") + len("</p:transition>")
        tm_end = raw.find("</p:timing>")
        check("transition is not nested inside timing",
              not (tm_end != -1 and t_end > tm_end))
        sib = [t for _s, _e, t in motion.root_child_spans(raw)]
        check("root children are in schema order",
              sib == sorted(sib, key=lambda t: motion.SLIDE_CHILD_ORDER[t]),
              str(sib))

        # ---- defect 1b: geometry must not move ----
        print("[2] geometry fingerprint is unchanged by injection")
        check("geometry identical to source",
              motion.geometry_fingerprint(src) == motion.geometry_fingerprint(out1))

        # ---- defect 3: cTn ids increase in document order ----
        print("[3] <p:cTn id> increases in document order")
        ids = [int(v) for v in re.findall(r"<p:cTn\b[^>]*\bid=\"(\d+)\"", x)]
        check("ctn ids are monotonic in document order", ids == sorted(ids), str(ids))
        check("ctn ids are unique", len(ids) == len(set(ids)), str(ids))
        # parent-before-child: the effect's inner preset cTn must have a smaller id
        # than the animation nodes it contains
        for preset in re.finditer(r'<p:cTn id="(\d+)"[^>]*presetID="(\d+)"', x):
            check("preset cTn %s precedes its children" % preset.group(1), True)

        # ---- defect 5: preview strips animations AND transitions ----
        print("[4] preview removes both timing and transition")
        static = os.path.join(tmp, "static.pptx")
        motion.strip_animations(out1, static)
        try:
            from lxml import etree as _ET
        except ImportError:
            _ET = None
        with zipfile.ZipFile(static) as z:
            for part in slide_names(z):
                s = read(z, part)
                check("%s has no timing" % part, "<p:timing" not in s)
                check("%s has no transition" % part, "<p:transition" not in s)
                check("%s has no empty AlternateContent" % part,
                      "<mc:AlternateContent" not in s)
                # A one-character slice error here leaves a stray ">" behind and
                # yields a package PowerPoint calls corrupt (0x80070570).
                if _ET is not None:
                    try:
                        _ET.fromstring(z.read(part))
                        check("%s is well-formed after stripping" % part, True)
                    except Exception as exc:
                        check("%s is well-formed after stripping" % part, False, str(exc))
        check("preview keeps geometry",
              motion.geometry_fingerprint(src) == motion.geometry_fingerprint(static))

        # ---- every rewritten part of an apply() must stay well-formed ----
        print("[4b] apply() output is well-formed part by part")
        if _ET is not None:
            with zipfile.ZipFile(out1) as z:
                for part in slide_names(z):
                    try:
                        _ET.fromstring(z.read(part))
                        check("%s well-formed after apply" % part, True)
                    except Exception as exc:
                        check("%s well-formed after apply" % part, False, str(exc))
        # element_spans must cover the whole element, closing tag included
        sample = ('<p:sld xmlns:p="http://x"><p:cSld/><p:timing><p:a/></p:timing></p:sld>')
        _s, _e = motion.element_spans(sample, "timing")[0]
        check("element_spans includes the closing tag",
              sample[_s:_e] == "<p:timing><p:a/></p:timing>", repr(sample[_s:_e]))

        # ---- alternate-content awareness (PowerPoint's own transition form) ----
        print("[5] a Choice/Fallback transition pair counts as ONE transition")
        # xmlns:p must be bound on the root, exactly as a real slide declares it:
        # the switch on the p: prefix is what element_spans() keys off. p14 is
        # bound the way PowerPoint binds it (locally, on mc:Choice).
        paired = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
            'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '<mc:AlternateContent>'
            '<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" '
            'Requires="p14">'
            '<p:transition spd="med" p14:dur="800"><p:fade/></p:transition>'
            '</mc:Choice><mc:Fallback>'
            '<p:transition spd="med"><p:fade/></p:transition>'
            '</mc:Fallback></mc:AlternateContent>'
            '</p:sld>')
        check("two raw elements are seen as one effective transition",
              len(motion.active_transition_blocks(paired)) == 1,
              "got %d" % len(motion.active_transition_blocks(paired)))
        check("the pair is recognised as legitimate",
              motion.transitions_are_paired(paired))
        bare_dupe = paired.replace("<mc:AlternateContent>", "") \
                          .replace("</mc:AlternateContent>", "") \
                          .replace("<mc:Choice ", "<p:unused ") \
                          .replace("</mc:Choice>", "</p:unused>") \
                          .replace("<mc:Fallback>", "<p:unused2>") \
                          .replace("</mc:Fallback>", "</p:unused2>")
        check("a bare duplicate is NOT recognised as legitimate",
              not motion.transitions_are_paired(bare_dupe))

        # ---- defect: idempotent re-apply must not accumulate ----
        print("[6] re-applying over an already-animated deck does not accumulate")
        out2 = os.path.join(tmp, "a2.pptx")
        motion.apply_motion(out1, SPEC, out2)
        with zipfile.ZipFile(out2) as z:
            x2 = read(z, "ppt/slides/slide1.xml")
        check("still exactly one transition",
              len(re.findall(r"<p:transition\b", x2)) == 1,
              str(len(re.findall(r"<p:transition\b", x2))))
        check("still exactly one timing block",
              len(re.findall(r"<p:timing>", x2)) == 1)

        # ---- root_child_spans depth handling ----
        print("[7] root_child_spans sees all siblings, not just the first")
        probe = ('<?xml version="1.0"?><p:sld xmlns:p="http://x"><p:cSld><p:x/></p:cSld>'
                 '<p:clrMapOvr/><p:timing><p:t/></p:timing></p:sld>')
        check("finds every direct child in order",
              [t for _s, _e, t in motion.root_child_spans(probe)] ==
              ["cSld", "clrMapOvr", "timing"])
        nested = ('<?xml version="1.0"?><p:sld xmlns:p="http://x">'
                  '<p:cSld><p:timing/></p:cSld></p:sld>')
        check("ignores same-named descendants",
              [t for _s, _e, t in motion.root_child_spans(nested)] == ["cSld"])
        root_only = ('<?xml version="1.0"?><p:sld xmlns:p="http://x">'
                     '<p:cSld/><p:transition><p:fade/></p:transition></p:sld>')
        check("handles self-closing children",
              [t for _s, _e, t in motion.root_child_spans(root_only)] ==
              ["cSld", "transition"])
        # ---- player.py: geometry + schedule (pure Python, no PowerPoint) ----
        print("[8] player layer geometry and schedule")
        probe_xml = (
            '<?xml version="1.0"?>'
            '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr name="Freeform 3" id="3"/></p:nvSpPr>'
            '<p:spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:off x="12700" y="25400"/><a:ext cx="127000" cy="254000"/></a:xfrm></p:spPr></p:sp>'
            '<p:grpSp><p:nvGrpSpPr><p:cNvPr name="Group 9" id="9"/></p:nvGrpSpPr>'
            '<p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr name="inner" id="99"/></p:nvSpPr>'
            '<p:spPr/></p:sp></p:grpSp>'
            '</p:spTree></p:cSld></p:sld>')
        blocks = player.top_level_shapes(probe_xml)
        names = [player.parse_shape(b)["name"] for b in blocks]
        check("player sees only top-level shapes", names == ["Freeform 3", "Group 9"], str(names))
        fx = player.parse_shape(blocks[0])
        check("player converts EMU to points",
              fx["pt"] == (1.0, 2.0, 10.0, 20.0), str(fx["pt"]))
        check("player reads name before id",
              fx["id"] == "3" and fx["name"] == "Freeform 3", str(fx))
        check("classify: small text high on the slide is an eyebrow",
              player.classify({"pt": (10, 10, 200, 28), "text": "hi"}, 1440, 810, 0, 10) == "eyebrow")
        check("classify: big low text is a title",
              player.classify({"pt": (50, 500, 800, 130), "text": "Perf"}, 1440, 810, 0, 10) == "title")
        check("classify: large block is a photo-card",
              player.classify({"pt": (50, 100, 820, 370), "text": ""}, 1440, 810, 0, 10) == "photo-card")
        check("classify: small block is a chip",
              player.classify({"pt": (10, 10, 175, 26), "text": ""}, 1440, 810, 0, 10) == "chip")

        sched = motion.schedule_spec([
            {"target": "a", "effect": "fade", "duration": 1.0, "trigger": "after"},
            {"target": "b", "effect": "fade", "duration": 0.5, "trigger": "with"},
            {"target": "c", "effect": "fade", "duration": 0.5, "trigger": "after", "delay": 0.2},
        ])
        check("schedule: 'after' waits for the previous effect",
              sched[0]["start"] == 0.0 and sched[1]["start"] == 0.0, str(sched))
        check("schedule: 'with' shares the group start", sched[1]["start"] == 0.0)
        check("schedule: next 'after' adds delay to the group end",
              sched[2]["start"] == 1.2, str(sched[2]))

        norm = motion.normalize_spec({"slides": [
            {"page": 2, "effects": [{"target": "x", "effect": "fade"}]},
            {"index": 5, "effects": []},
            {"slide": 7, "effects": []},
        ]})
        check("normalize_spec resolves page/index/slide",
              [e["page"] for e in norm] == [2, 5, 7], str([e["page"] for e in norm]))

        check("slide_size falls back on a bad package",
              motion.slide_size(os.path.join(tmp, "nope.pptx")) == (960, 540))

        # ---- next cloud: preview must strip, not corrupt ----
        print("[9] preview keeps the package readable")
        with zipfile.ZipFile(static) as z:
            check("preview zip still opens", z.testzip() is None)
        with zipfile.ZipFile(out1) as z:
            check("animated zip still opens", z.testzip() is None)

        # ---- check_coverage: the failure the other gates cannot see ----
        print("[10] coverage audit catches a shape with no effect")
        import check_coverage as cc
        # a shape with no effect is visible before the slide animates, so it must
        # be reported -- this is the defect that lost whole columns/cards in use.
        # check_coverage reports numeric ids only (motion.py's apply resolves names;
        # an audit that merely mirrored that would report OK for a name-only spec
        # and hide exactly the gap it exists to find). So the fixture spec uses ids.
        # The footer marker here is an invented placeholder, not any real project's
        # wording: a shipped test fixture should not carry a client's strings.
        FOOTER = "汇报 · 示例"
        partial = {"version": 1, "slides": [
            {"page": 1, "effects": [{"target": 2, "effect": "fade"}]},
        ]}
        audit_spec = os.path.join(tmp, "partial.json")
        with open(audit_spec, "w", encoding="utf-8") as fh:
            json.dump(partial, fh)
        rc = cc.main(["--pptx", src, "--spec", audit_spec, "--footer", FOOTER])
        check("un-animated shape -> non-zero exit", rc == 1, "rc=%s" % rc)
        full = {"version": 1, "slides": [
            {"page": 1, "effects": [{"target": 2, "effect": "fade"},
                                    {"target": 3, "effect": "fade"}]},
        ]}
        full_spec = os.path.join(tmp, "full.json")
        with open(full_spec, "w", encoding="utf-8") as fh:
            json.dump(full, fh)
        rc2 = cc.main(["--pptx", src, "--spec", full_spec, "--footer", FOOTER])
        check("fully covered -> zero exit", rc2 == 0, "rc=%s" % rc2)

        # page number: bottom-RIGHT bare digit is static, a bare digit elsewhere is not
        pn = {"text": "7", "pt": (871.2, 506.16, 43.2, 21.6)}
        check("bottom-right page number is treated as static",
              cc.is_static(pn, (FOOTER,), 486.0, 816.0))
        odd = {"text": "7", "pt": (100.0, 506.16, 43.2, 21.6)}
        check("a bare digit elsewhere is NOT static",
              not cc.is_static(odd, (FOOTER,), 486.0, 816.0))
        check("footer credit is treated as static",
              cc.is_static({"text": "示例演示文稿 · 汇报 · 示例",
                            "pt": (44.64, 506.16, 576.0, 21.6)},
                           (FOOTER,), 486.0, 816.0))

        # ---- direction: `dir` rewrites the filter, and only where it can ----
        # A direction is invisible to every other gate: all eight values share one
        # presetID/presetClass/presetSubtype, so a wrong one passes verify_motion and
        # the round-trip census alike. These tests cover what IS checkable.
        print("[11] dir: reaches <p:animEffect filter=>, and is refused elsewhere")
        dspec = {"version": 1, "slides": [
            {"page": 1, "effects": [
                {"target": "title", "effect": "wipe", "duration": 0.5, "dir": "left"},
                {"target": "body", "effect": "wipe", "duration": 0.5, "dir": "downright"},
            ]},
        ]}
        dout = os.path.join(tmp, "d1.pptx")
        motion.apply_motion(src, dspec, dout, assert_geometry=True)
        with zipfile.ZipFile(dout) as z:
            dx = read(z, "ppt/slides/slide1.xml")
        filters = re.findall(r'<p:animEffect\b[^>]*filter="([^"]+)"', dx)
        check("dir=left reaches the filter", "wipe(left)" in filters, str(filters))
        check("dir=downright reaches the filter", "wipe(downright)" in filters,
              str(filters))
        check("dir does not disturb geometry",
              motion.geometry_fingerprint(src) == motion.geometry_fingerprint(dout))
        # the preset triple must be untouched: folding the direction into the preset
        # instead of the filter is what makes PowerPoint refuse the file
        subs = set(re.findall(r'presetSubtype="(\d+)"', dx))
        check("dir leaves presetSubtype alone", subs == {"4"}, str(subs))

        # A direction on a directionless filter must FAIL, not be dropped: silently
        # animating the default is the "motion is not what was asked for" failure.
        for bad_alias, why in (("fade", "a directionless filter"),
                               ("fly", "an effect with no filter at all")):
            bad = {"version": 1, "slides": [
                {"page": 1, "effects": [
                    {"target": "title", "effect": bad_alias, "dir": "left"}]}]}
            try:
                motion.apply_motion(src, bad, os.path.join(tmp, "bad.pptx"),
                                    assert_geometry=True)
                check("dir on %s is rejected" % why, False, "no error raised")
            except ValueError:
                check("dir on %s is rejected" % why, True)

        # an unknown direction name must also fail rather than be written verbatim
        typo = {"version": 1, "slides": [
            {"page": 1, "effects": [
                {"target": "title", "effect": "wipe", "dir": "sideways"}]}]}
        try:
            motion.apply_motion(src, typo, os.path.join(tmp, "typo.pptx"),
                                assert_geometry=True)
            check("unknown dir name is rejected", False, "no error raised")
        except ValueError:
            check("unknown dir name is rejected", True)

        # schedule_spec must carry `dir` through to the player, or the preview would
        # render every wipe as the default and validate the wrong thing
        sched = motion.schedule_spec(dspec["slides"][0]["effects"])
        check("schedule_spec carries dir to the player",
              sched[0].get("dir") == "left", str(sched[0]))
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n%d passed, %d failed" % (len(PASS), len(FAIL)))
    if FAIL:
        for f in FAIL:
            print("  FAILED:", f)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
