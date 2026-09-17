# -*- coding: utf-8 -*-
"""Structural self-check for a motion-injected deck.

Verifies, without needing PowerPoint:
  1. every slide part is well-formed XML
  2. animation node counts and preset params per slide
  3. every spid referenced by a timing node exists as a shape in that slide
  4. every cTn id inside one slide's timing is unique
  5. geometry fingerprint equals the source deck
"""
import argparse, json, os, re, sys, zipfile
import lxml.etree as ET

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "scripts"))
import motion  # noqa: E402

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def check(pptx, source=None):
    errs, warns, info = [], [], {}
    z = zipfile.ZipFile(pptx)
    slides = sorted([n for n in z.namelist() if motion.SLIDE_RE.match(n)],
                    key=lambda n: int(motion.SLIDE_RE.match(n).group(1)))
    total_effects = 0
    total_paths = 0
    detail = []
    for name in slides:
        raw = z.read(name)
        try:
            root = ET.fromstring(raw)
        except ET.XMLSyntaxError as exc:
            errs.append("%s: malformed XML: %s" % (name, exc))
            continue
        xml = raw.decode("utf-8")

        # shape ids present in this slide
        shape_ids = set()
        for cNvPr in root.iter(f"{P}cNvPr"):
            if cNvPr.get("id"):
                shape_ids.add(cNvPr.get("id"))

        # referenced shape ids
        refd = re.findall(r"<p:spTgt\s+spid=\"(\d+)\"", xml)
        missing = sorted(set(refd) - shape_ids)
        if missing:
            errs.append("%s: timing references unknown shape ids %s" % (name, missing))

        # cTn id uniqueness
        ctns = re.findall(r"<p:cTn\b[^>]*\bid=\"(\d+)\"", xml)
        dupes = sorted({i for i in ctns if ctns.count(i) > 1})
        if dupes:
            errs.append("%s: duplicate cTn ids %s" % (name, dupes))

        presets = re.findall(r"presetID=\"(\d+)\"", xml)
        paths = re.findall(r"<p:animMotion\b", xml)
        raw_trans = re.findall(r"<p:transition\b", xml)
        # PowerPoint writes a transition as mc:AlternateContent{Choice, Fallback},
        # so a healthy PowerPoint-saved deck legitimately holds two <p:transition>
        # elements for one transition. Judge the *effective* count, and only treat
        # an unwrapped duplicate as an error.
        trans = motion.active_transition_blocks(xml)
        total_effects += len(presets)
        total_paths += len(paths)
        detail.append({
            "part": name, "shapes": len(shape_ids), "effects": len(presets),
            "motionPaths": len(paths), "transition": bool(trans),
        })
        if len(trans) > 1:
            errs.append("%s: %d effective transition elements (expected <=1)"
                        % (name, len(trans)))
        elif len(raw_trans) > len(trans):
            if not motion.transitions_are_paired(xml):
                errs.append("%s: %d stray <p:transition> element(s) outside "
                            "mc:AlternateContent" % (name, len(raw_trans) - len(trans)))
        if slide_with_anim(name, presets) and "<p:timing>" not in xml:
            errs.append("%s: preset nodes without <p:timing>" % name)

    info["slides"] = detail
    info["effects_total"] = total_effects
    info["motion_paths_total"] = total_paths

    if source:
        b = motion.geometry_fingerprint(source)
        a = motion.geometry_fingerprint(pptx)
        diffs = [k for k in sorted(set(b) | set(a)) if b.get(k) != a.get(k)]
        info["geometry_ok"] = not diffs
        info["geometry_diff"] = diffs
        if diffs:
            errs.append("geometry differs from source: %s" % diffs)
    return {"ok": not errs, "errors": errs, "warnings": warns, "info": info}


def slide_with_anim(name, presets):
    return bool(presets)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--source")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args()
    res = check(ns.pptx, ns.source)
    if ns.json:
        print(json.dumps(res, ensure_ascii=False, indent=1))
    else:
        print("OK" if res["ok"] else "FAILED")
        print("effects=%d motionPaths=%d" % (res["info"]["effects_total"],
                                            res["info"]["motion_paths_total"]))
        if "geometry_ok" in res["info"]:
            print("geometry vs source: %s" % ("identical" if res["info"]["geometry_ok"] else "DIFFERS"))
        for d in res["info"]["slides"]:
            if d["effects"] or d["transition"]:
                print("  %-22s effects=%-3d paths=%-3d transition=%s" % (
                    d["part"], d["effects"], d["motionPaths"], d["transition"]))
        for e in res["errors"]:
            print("ERROR:", e)
        for w in res["warnings"]:
            print("WARN:", w)
    return 0 if res["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
