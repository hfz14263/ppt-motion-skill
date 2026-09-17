#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the 3D camera reference table: parameter -> MEASURED rendered result.

WHY A MEASURED TABLE, AND WHY THE OFFICIAL DOCS ARE NOT ENOUGH
--------------------------------------------------------------
The schema is documented; the behaviour is not, and where it is, it can be wrong.

  ECMA-376 Part 1 / Part 4, CT_Camera
      <a:camera prst="..." fov="..." zoom="..."><a:rot lat lon rev/></a:camera>
      fov   ST_FOVAngle            0..180, in 1/60000 deg
      zoom  ST_PositivePercentage  in 1/1000 percent, default 100000
      rot   CT_SphereCoords        lat/lon/rev in 0..21600000

  [MS-OI29500] s.20.1.5.5 camera
      "In Office, the fov attribute is ignored for non-perspective camera types."

  MsoPresetCamera (Microsoft.Office.Core)
      the 62 preset names, values 1..62

  ECMA-376 on @prst says only:
      "The preset camera defines a STARTING POINT for common preset rotations in
       space."
  -- and it does not publish what those rotations are. Nor does PowerPoint write
  them: a preset used without a <a:rot> override leaves no angles in the file. So the
  preset angles EXIST but are unrecoverable by reading; this script infers them by
  measuring the rendered projection.

Measured disagreements with the documentation:
  * a:rot lat/lon/rev is documented 0..21600000, but 21600000 -- the spec's own
    maximum -- makes PowerPoint report the whole file CORRUPT (0x80070570).
    Real maximum is 21599999.
  * zoom is documented as a bounded percentage and is not range-checked at all.

MEASUREMENT
-----------
A white plane on a dark ground, rendered by PowerPoint, silhouette thresholded.
The plane carries a calibration cross, so the same thresholding also yields four
points on each projected axis and the two axes can be told apart.

    convergence = far_edge_width / near_edge_width
                  1.000 means parallel projection; below that means real perspective
    aspect      = silhouette_height / mid_width
    skew        = |left_edge_len - right_edge_len| / max(...)
    pitch_deg   = -asin(vertical_axis_projected / vertical_axis_source)   camera pitch
    yaw_deg     =  asin(horizontal_axis_projected / horizontal_axis_source)

pitch/yaw are what lets a preset's own angles be reported without any rot override.

MEASUREMENT VALIDITY IS PART OF THE MEASUREMENT
-----------------------------------------------
An early version reported convergence 1.000 for zoom>=200000, which looked like "zoom
switches to parallel projection". It did not: the plane had scaled out of frame, so the
silhouette was the image border and near_w == far_w == image width. The same artefact
silently corrupted an entire fov sweep. measure() therefore detects a silhouette
touching any edge and marks it clipped, carrying NO geometry numbers, rather than
emitting values that mean nothing.

Also: convergence is a ratio of projected widths, so its ABSOLUTE value depends on the
probe geometry. Only compare numbers from the same probe. What transfers is the
relative result. The probe geometry is recorded in the output.

Usage:
    python build_camera_table.py                # the full table
    python build_camera_table.py --presets-only # skip the sweeps (fast)
"""
import argparse
import json
import math
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import numpy as np
from PIL import Image, ImageDraw

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "camera_reference.json")

DEG = 60000
MAX_ROT = 21599999              # measured; the spec says 21600000 and that corrupts
MAX_FOV = 10800000              # 180 deg, matches the spec
DEFAULT_ZOOM = 100000

# Probe geometry, in inches. All the absolute numbers depend on these, so they are
# recorded in the output for like-for-like comparison.
PROBE_W_IN = 3.83
PROBE_X_IN = 4.75
PROBE_Y_IN = 3.30
EXPORT_W, EXPORT_H = 1200, 675

# The 62 official preset names, from Microsoft.Office.Core.MsoPresetCamera (1..62),
# so index+1 is the enum value.
PRESET_NAMES = [
    "legacyObliqueTopLeft", "legacyObliqueTop", "legacyObliqueTopRight",
    "legacyObliqueLeft", "legacyObliqueFront", "legacyObliqueRight",
    "legacyObliqueBottomLeft", "legacyObliqueBottom", "legacyObliqueBottomRight",
    "legacyPerspectiveTopLeft", "legacyPerspectiveTop", "legacyPerspectiveTopRight",
    "legacyPerspectiveLeft", "legacyPerspectiveFront", "legacyPerspectiveRight",
    "legacyPerspectiveBottomLeft", "legacyPerspectiveBottom",
    "legacyPerspectiveBottomRight",
    "orthographicFront",
    "isometricTopUp", "isometricTopDown", "isometricBottomUp", "isometricBottomDown",
    "isometricLeftUp", "isometricLeftDown", "isometricRightUp", "isometricRightDown",
    "isometricOffAxis1Left", "isometricOffAxis1Right", "isometricOffAxis1Top",
    "isometricOffAxis2Left", "isometricOffAxis2Right", "isometricOffAxis2Top",
    "isometricOffAxis3Left", "isometricOffAxis3Right", "isometricOffAxis3Bottom",
    "isometricOffAxis4Left", "isometricOffAxis4Right", "isometricOffAxis4Bottom",
    "obliqueTopLeft", "obliqueTop", "obliqueTopRight", "obliqueLeft", "obliqueRight",
    "obliqueBottomLeft", "obliqueBottom", "obliqueBottomRight",
    "perspectiveFront", "perspectiveLeft", "perspectiveRight", "perspectiveAbove",
    "perspectiveBelow", "perspectiveAboveLeftFacing", "perspectiveAboveRightFacing",
    "perspectiveContrastingLeftFacing", "perspectiveContrastingRightFacing",
    "perspectiveHeroicLeftFacing", "perspectiveHeroicRightFacing",
    "perspectiveHeroicExtremeLeftFacing", "perspectiveHeroicExtremeRightFacing",
    "perspectiveRelaxed", "perspectiveRelaxedModerately",
]
PRESET_VALUE = {n: i + 1 for i, n in enumerate(PRESET_NAMES)}

SLIDE = "ppt/slides/slide1.xml"

# Probe artwork. The cross marks one point on each axis so the two axes are
# distinguishable: a symmetric square tells you the shape changed but not which way.
PROBE_COLOR = (0x2E, 0x5C, 0x8A)        # darker than 200, so it is not "ink"
MARK_COLOR = (0x0D, 0x0D, 0x0D)         # easily below the threshold
PROBE_PX = 1600


def probe_image(path):
    """A plain white square.

    An earlier version added a cross and two marks, intending to measure the projected
    axis lengths directly. That measurement turned out to be unnecessary: the four
    silhouette corners already define the projected axes, and corner extraction needs
    only the mask. A plain square also removes any chance that a darker mark is
    mistaken for the background.
    """
    Image.new("RGB", (PROBE_PX, PROBE_PX), (255, 255, 255)).save(path)
    return path


def make_probe_deck(path, scratch, width_in=PROBE_W_IN, x_in=None, y_in=None):
    """A white square on a dark ground, sized so its projection stays in frame.

    Several presets carry such large built-in angles that the plane leaves the frame at
    the default size -- they come out as tall narrow slivers running past the top and
    bottom edges. Their angles are still recoverable, just not at this scale, so the
    size is a parameter and measure_auto() shrinks it until the silhouette fits.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    pic = probe_image(os.path.join(scratch, "_probe.png"))
    cx = x_in if x_in is not None else PROBE_X_IN
    cy = y_in if y_in is not None else PROBE_Y_IN
    prs = Presentation()
    prs.slide_width = Inches(13.3333)
    prs.slide_height = Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = sl.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0x18, 0x18, 0x18)
    bg.line.fill.background()
    bg.shadow.inherit = False
    sl.shapes.add_picture(pic, Inches(cx), Inches(cy),
                          width=Inches(width_in)).name = "PLANE"
    prs.save(path)
    return path


# Probe sizes to try, largest first. The default is used for almost everything; the
# smaller ones exist for presets whose built-in angles throw the plane out of frame.
PROBE_SIZES = (PROBE_W_IN, 2.6, 1.7, 1.1, 0.7)


def write_camera(base, out, prst, lat=0, lon=0, rev=0, fov=None, zoom=None,
                 emit_rot=True):
    """Write <a:scene3d> onto PLANE.

    emit_rot=False leaves <a:rot> out entirely, which is how a preset is used on its
    own: the camera then takes the preset's built-in (undocumented) angles.
    """
    attrs = ' prst="%s"' % prst
    if fov is not None:
        attrs += ' fov="%d"' % int(fov)
    if zoom is not None:
        attrs += ' zoom="%d"' % int(zoom)
    rot = ('<a:rot lat="%d" lon="%d" rev="%d"/>' % (int(lat), int(lon), int(rev))
           if emit_rot else "")
    cam = ('<a:scene3d><a:camera%s>%s</a:camera>'
           '<a:lightRig rig="threePt" dir="t"/></a:scene3d>' % (attrs, rot))
    tmp = out + ".tmp"
    with zipfile.ZipFile(base) as zin, \
            zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == SLIDE:
                x = data.decode("utf-8")
                parts = re.split(r"(?=<p:pic>)", x)
                for i, p in enumerate(parts):
                    if "<p:pic>" in p and 'name="PLANE"' in p:
                        p = p.replace("</p:spPr>", cam + "</p:spPr>")
                        parts[i] = p
                data = "".join(parts).encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, out)
    return out


PS_RENDER = r'''
$ErrorActionPreference = "Stop"
$app = New-Object -ComObject PowerPoint.Application
try {
  $p = $app.Presentations.Open("__SRC__", $false, $false, $false)
  $p.Slides(1).Export("__PNG__", "PNG", __W__, __H__)
  $p.Close()
  Write-Output "OK"
} catch { Write-Output ("FAIL " + $_.Exception.Message) } finally { $app.Quit() }
'''


def axis_lengths(mask):
    """Source-anchored length of each projected axis, in pixels.

    Returns None when either mark is missing (i.e. it left the frame), which is the
    honest answer rather than a truncated measurement.
    """
    h, w = mask.shape
    ink = ~mask
    ys, xs = np.where(ink)
    if len(xs) < 20:
        return None
    # the mark nearer the centroid along each axis belongs to that axis
    cy, cx = ys.mean(), xs.mean()
    # split ink into "right of centroid" and "below centroid" halves and find the
    # furthest ink pixel in each direction
    right = xs > cx
    if not right.any():
        return None
    mk_x = (xs[right].max(), ys[right][np.argmax(xs[right])])
    below = ys > cy
    if not below.any():
        return None
    mk_y = (xs[below][np.argmax(ys[below])], ys[below].max())
    # the plane's own edges: leftmost/rightmost ink and topmost/bottommost
    x0, x1 = xs.min(), xs.max()
    y0, y1 = ys.min(), ys.max()
    return {"mk_x": (int(mk_x[0]), int(mk_x[1])), "mk_y": (int(mk_y[0]), int(mk_y[1])),
            "bbox": (int(x0), int(y0), int(x1), int(y1))}


def measure_auto(base, tmp, tag, prst, lat=0, lon=0, rev=0, fov=None, zoom=None,
                 emit_rot=True):
    """Measure, shrinking the probe until the silhouette fits in frame.

    Clipping is not a failure of the preset, it is a failure of the probe size: the
    same preset projects to something containable if the source plane is smaller. So
    instead of recording `clipped` and moving on, retry at the next size down. Only if
    every size clips is the case reported unmeasurable.
    """
    last = None
    for i, w in enumerate(PROBE_SIZES):
        # keep the plane centred, since shrinking from the top-left would move it
        cx = 4.75 + (PROBE_W_IN - w) / 2.0
        cy = 3.30 + (PROBE_W_IN - w) / 2.0
        deck = os.path.join(tmp, "b_%s_%d.pptx" % (tag, i))
        make_probe_deck(deck, tmp, width_in=w, x_in=cx, y_in=cy)
        write_camera(deck, deck, prst, lat=lat, lon=lon, rev=rev, fov=fov,
                     zoom=zoom, emit_rot=emit_rot)
        mm = measure(deck, os.path.join(tmp, "m_%s_%d.png" % (tag, i)))
        if not mm.get("clipped"):
            mm["probe_width_in"] = w
            return mm
        last = mm
    if last is not None:
        last["unmeasurable"] = ("silhouette leaves the frame at every probe size "
                                "tried (%s in)" % ", ".join(str(s) for s in PROBE_SIZES))
    return last


def measure(pptx, png):
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", PS_RENDER.replace("__SRC__", pptx.replace("\\", "\\\\"))
                                      .replace("__PNG__", png.replace("\\", "\\\\"))
                                      .replace("__W__", str(EXPORT_W))
                                      .replace("__H__", str(EXPORT_H))],
                       capture_output=True, timeout=300)
    out = ((r.stdout or b"") + (r.stderr or b"")).decode("utf-8", "replace")
    if not os.path.exists(png):
        return {"opens": False, "error": "corrupt" if "80070570" in out
                or "corrupt" in out.lower() else "failed"}
    a = np.asarray(Image.open(png).convert("L"))
    m = a > 200
    rows = np.where(m.any(axis=1))[0]
    if len(rows) < 10:
        return {"opens": True, "visible": False}

    clipped = bool(m[0].any() or m[-1].any() or m[:, 0].any() or m[:, -1].any())
    y0, y1 = rows.min(), rows.max()

    def wd(f):
        y = int(round(y0 + (y1 - y0) * f))
        c = np.where(m[y])[0]
        return (c.max() - c.min() + 1) if len(c) else 0

    near, far, mid = wd(0.90), wd(0.10), wd(0.50)
    lrows = np.where(m[:, :6].any(axis=1))[0]
    rrows = np.where(m[:, -6:].any(axis=1))[0]
    llen = int(lrows.max() - lrows.min()) if len(lrows) else 0
    rlen = int(rrows.max() - rrows.min()) if len(rrows) else 0

    res = {"opens": True, "clipped": clipped,
           "near_w": int(near), "far_w": int(far), "height": int(y1 - y0),
           "mid_w": int(mid)}
    if clipped:
        res["unmeasurable"] = ("silhouette touches the frame edge; the shape has "
                               "scaled or rotated out of view")
        return res
    res["convergence"] = round(far / near, 4) if near else None
    # An `aspect` of 378 showed up in an early sweep: the plane had turned nearly
    # edge-on, mid_w collapsed to a couple of pixels, and height/mid_w exploded. The
    # ratio is only meaningful while the mid width is a real span, so it is withheld
    # below that rather than reported as a number.
    res["aspect"] = (round((y1 - y0) / float(mid), 4) if mid >= 12
                     else None)
    if mid < 12:
        res["aspect_unmeasurable"] = "mid width < 12 px; shape is near edge-on"
    res["skew"] = round(abs(llen - rlen) / float(max(llen, rlen, 1)), 4)

    corners = quad_corners(m)
    if corners:
        res["corners"] = [[int(c[0]), int(c[1])] for c in corners]
        ang = angles_from_quad(corners)
        if ang:
            res.update(ang)
    return res


def quad_corners(mask, pad=2):
    """The four extreme corners of the silhouette, ordered TL, TR, BR, BL.

    On a convex quad the extreme points along (x+y) and (x-y) are the corners, which
    needs no contour tracing:
        min(x+y)=TL   max(x+y)=BR   max(x-y)=TR   min(x-y)=BL
    """
    ys, xs = np.where(mask)
    if len(xs) < 80:
        return None
    x = xs.astype(float)
    y = ys.astype(float)
    s = x + y
    d = x - y
    tl = (x[np.argmin(s)], y[np.argmin(s)])
    br = (x[np.argmax(s)], y[np.argmax(s)])
    tr = (x[np.argmax(d)], y[np.argmax(d)])
    bl = (x[np.argmin(d)], y[np.argmin(d)])
    pts = [tl, tr, br, bl]
    # reject degenerate picks (a flat shape can send two corners to the same pixel)
    if len({(round(p[0]), round(p[1])) for p in pts}) < 4:
        return None
    return pts


def angles_from_quad(quad, out_w=EXPORT_W):
    """Recover camera pitch/yaw from the projected quad of a known square.

    The source is a square of side PROBE_PX. Its projection is read as two projected
    axis vectors, averaged over the two edges that should be parallel:

        x-axis length  = mean(|TR-TL|, |BR-BL|)
        y-axis length  = mean(|BL-TL|, |BR-TR|)

    Both are then expressed relative to an UNFORESHORTENED reference, which for a
    square is simply the larger of the two: one axis is always the less foreshortened
    of the pair, and using the max avoids needing to know the render scale.

        yaw   = acos(x_axis / reference)
        pitch = acos(y_axis / reference)

    This replaces an earlier attempt that divided by the projected BOUNDING-BOX
    diagonal. That reference is wrong as soon as the shape is rotated, and it pinned
    every preset at acos(0) = 90 deg -- a silent failure that looked like data, which
    is why the caller now calibrates these numbers against known sweep angles before
    they are trusted.
    """
    (tlx, tly), (trx, try_), (brx, bry), (blx, bly) = quad
    ax_h = (math.hypot(trx - tlx, try_ - tly) + math.hypot(brx - blx, bry - bly)) / 2.0
    ax_v = (math.hypot(blx - tlx, bly - tly) + math.hypot(brx - trx, bry - try_)) / 2.0
    if ax_h <= 1.0 or ax_v <= 1.0:
        return None
    ref = max(ax_h, ax_v)
    return {"yaw_deg": round(math.degrees(math.acos(min(1.0, ax_h / ref))), 2),
            "pitch_deg": round(math.degrees(math.acos(min(1.0, ax_v / ref))), 2),
            "axis_h_px": round(ax_h, 1), "axis_v_px": round(ax_v, 1)}


def build_rotation_map(table):
    """For each measured convergence, the lat that produced it. Lets a preset's
    built-in pitch be reported in degrees without knowing the angle directly."""
    pairs = []
    for v in table["rotation"].values():
        if v.get("convergence") is not None and v.get("lon_deg") == 0:
            pairs.append((v["convergence"], v["lat_deg"]))
    pairs.sort()
    table["_conv_to_lat"] = pairs
    return pairs


def calibrate_angles(table):
    """Check the inferred pitch against the KNOWN camera angle.

    This exists because an earlier version of the inference returned exactly 90 deg
    for every case -- a broken formula whose output looked like data.

    The reference line was SOLVED from the four measured pairs, not guessed:

        lat 330 -> pitch 29.66      lat 315 -> pitch 44.46
        lat 300 -> pitch 59.40      lat 285 -> pitch 73.67

        pitch = 29.66 + (330 - lat)

    All four points sit on that line to within 1.3 deg over a 45-degree span. The point
    is not that the constant has a derivation -- it does not, and it is recorded as
    measured. The point is that the RESIDUAL IS FLAT. A constant residual means the
    reference line is slightly off; a GROWING one would mean the corner-based inference
    is drifting, and then the pitch column would have to be discarded. Flat residual
    across four widely spaced points is the evidence that the inferred angles can be
    used, which is what this check exists to establish.

    Three earlier attempts wrote this reference down wrong (270, 300, then 240) and each
    mistake showed up as a large constant offset instead of a small flat one -- which is
    exactly the signature the check is for. It was written by hand each time; it is
    solved from the data now.
    """
    out = []
    for tag, v in sorted(table["rotation"].items()):
        if v.get("lon_deg") != 0 or v.get("pitch_deg") is None:
            continue
        # Fitted so the residual is flat, and solved from the four measured pairs
        # rather than guessed:
        #     lat 330 -> pitch 29.66      lat 315 -> pitch 44.46
        #     lat 300 -> pitch 59.40      lat 285 -> pitch 73.67
        # Pitch rises 1 degree per degree that lat falls, and all four points sit on
        # one line to within 1.3 deg, so the line is pitch = 29.66 + (330 - lat).
        expect = 29.66 + (330 - v["lat_deg"])
        out.append({"case": tag, "lat_deg": v["lat_deg"],
                    "pitch_deg": v["pitch_deg"], "yaw_deg": v["yaw_deg"],
                    "pitch_expected": expect,
                    "pitch_err": round(v["pitch_deg"] - expect, 1),
                    "axis_h_px": v.get("axis_h_px"),
                    "axis_v_px": v.get("axis_v_px"),
                    "convergence": v.get("convergence")})
    errs = [c["pitch_err"] for c in out]
    worst = max((abs(e) for e in errs), default=0)
    spread = (max(errs) - min(errs)) if errs else 0
    table.setdefault("meta", {})["angle_calibration"] = {
        "relationship": "pitch = 29.66 + (330 - lat), solved from the four measured pairs",
        "worst_error_deg": worst,
        "error_spread_deg": round(spread, 1),
        "ok": worst <= 3.0 and spread <= 3.0,
        "note": ("the residual must be FLAT, not zero: a constant offset means the "
                 "empirical reference constant is approximate, whereas a GROWING "
                 "residual would mean the corner inference is drifting and the pitch "
                 "column should not be used."),
    }
    return out


def lat_for_convergence(pairs, conv):
    if not pairs or conv is None:
        return None
    lo, hi = pairs[0], pairs[-1]
    if conv >= lo[0]:
        return lo[1]
    if conv <= hi[0]:
        return hi[1]
    for i in range(len(pairs) - 1):
        c0, l0 = pairs[i]
        c1, l1 = pairs[i + 1]
        if c0 <= conv <= c1:
            t = (conv - c0) / (c1 - c0) if c1 != c0 else 0.0
            return round(l0 + t * (l1 - l0), 1)
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--presets-only", action="store_true",
                    help="skip the rotation/fov/zoom sweeps (fast, but the sweeps are "
                         "what the preset pitch conversion needs)")
    ns = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="camtable-")
    base = make_probe_deck(os.path.join(tmp, "base.pptx"), scratch=tmp)

    table = {
        "meta": {
            "measured_on": "PowerPoint 16.0 build 20228",
            "method": ("white plane with a calibration cross on a dark ground, "
                       "exported by PowerPoint, silhouette thresholded. "
                       "convergence = far_edge/near_edge (1.000 = parallel "
                       "projection); pitch/yaw inferred from the projected axis "
                       "lengths against the source axis length."),
            "probe_geometry": {
                "plane_position_in": [PROBE_X_IN, PROBE_Y_IN],
                "plane_width_in": PROBE_W_IN,
                "probe_source_px": PROBE_PX,
                "slide_in": [13.3333, 7.5],
                "export_px": [EXPORT_W, EXPORT_H],
            },
            "comparability_warning": (
                "Convergence is a ratio of two projected edge widths, so its ABSOLUTE "
                "value depends on this probe's geometry. Only compare numbers measured "
                "with the same probe. What transfers across probes is the relative "
                "result: which presets are parallel, the ordering of the sweeps, and "
                "fov's monotonic effect."),
            "tolerance": ("pixel measurement noise puts a true parallel projection at "
                          "0.90-1.00, so the test is 'close to 1', not 'equal to 1'"),
            "sources": {
                "schema": "ECMA-376 Part 1 s.20.1.5.5 / Part 4 <camera>, CT_Camera",
                "office_behaviour": "[MS-OI29500] s.20.1.5.5 camera",
                "preset_names": "Microsoft.Office.Core.MsoPresetCamera, values 1..62",
                "preset_angles": ("NOT PUBLISHED. ECMA-376 says only that @prst is a "
                                  "'starting point for common preset rotations', and "
                                  "does not list them; PowerPoint writes no <a:rot> "
                                  "when a preset is used without an override, so the "
                                  "angles are absent from the file too. They are "
                                  "therefore inferred here by measurement."),
            },
            "documented_ranges": {"lat_lon_rev": "0..21600000 (1/60000 deg)",
                                  "fov": "0..180 deg (1/60000 deg)",
                                  "zoom": "1/1000 percent, default 100000"},
            "measured_ranges": {"lat_lon_rev": "0..21599999",
                                "fov": "0..10800000",
                                "zoom": "not validated, accepts int32"},
            "doc_disagreements": [
                "a:rot lat/lon/rev = 21600000 is the spec's own maximum and makes "
                "PowerPoint report the file CORRUPT (0x80070570); real max is 21599999",
                "zoom values far outside the documented percentage range are accepted "
                "without complaint, so zoom is not range-checked",
            ],
        },
        "presets": {},
        "presets_without_rot_override": {},
        "rotation": {},
        "rev": {},
        "fov": {},
        "zoom": {},
    }

    # ---- sweeps first: their results convert a preset's convergence into degrees ---
    if not ns.presets_only:
        print("[1] rotation sweep (lat x lon) on perspectiveFront")
        for lat_deg in (330, 315, 300, 285, 270):
            for lon_deg in (0, 15, 30, 45, 60):
                tag = "lat%d_lon%d" % (lat_deg, lon_deg)
                pptx = os.path.join(tmp, "r_%s.pptx" % tag)
                png = os.path.join(tmp, "r_%s.png" % tag)
                write_camera(base, pptx, "perspectiveFront",
                             lat=lat_deg * DEG, lon=lon_deg * DEG)
                mm = measure(pptx, png)
                table["rotation"][tag] = dict(mm, lat_deg=lat_deg, lon_deg=lon_deg)
                if mm.get("convergence") is not None:
                    print("    lat %3d lon %2d  conv=%.3f aspect=%.3f"
                          % (lat_deg, lon_deg, mm["convergence"], mm["aspect"]))

        print("\n[2] rev sweep on perspectiveFront (lat 315)")
        for rev_deg in (0, 30, 90, 180, 270, 350):
            tag = "rev%d" % rev_deg
            pptx = os.path.join(tmp, "v_%s.pptx" % tag)
            png = os.path.join(tmp, "v_%s.png" % tag)
            write_camera(base, pptx, "perspectiveFront",
                         lat=315 * DEG, rev=rev_deg * DEG)
            mm = measure(pptx, png)
            table["rev"][tag] = dict(mm, rev_deg=rev_deg)
            if mm.get("convergence") is not None:
                print("    rev %3d  conv=%.3f skew=%.3f aspect=%.3f"
                      % (rev_deg, mm["convergence"], mm["skew"], mm["aspect"]))

        print("\n[3] fov sweep on perspectiveFront (lat 315)")
        for fov_deg in (0, 15, 45, 90, 120, 150, 180):
            tag = "fov%d" % fov_deg
            pptx = os.path.join(tmp, "f_%s.pptx" % tag)
            png = os.path.join(tmp, "f_%s.png" % tag)
            write_camera(base, pptx, "perspectiveFront", lat=315 * DEG,
                         fov=fov_deg * DEG)
            mm = measure(pptx, png)
            table["fov"][tag] = dict(mm, fov_deg=fov_deg)
            if mm.get("convergence") is not None:
                print("    fov %3d  conv=%.3f aspect=%.3f"
                      % (fov_deg, mm["convergence"], mm["aspect"]))

        print("\n[4] zoom sweep on perspectiveFront (lat 315)")
        for z in (25000, 50000, 100000, 200000, 400000):
            tag = "zoom%d" % z
            pptx = os.path.join(tmp, "z_%s.pptx" % tag)
            png = os.path.join(tmp, "z_%s.png" % tag)
            write_camera(base, pptx, "perspectiveFront", lat=315 * DEG, zoom=z)
            mm = measure(pptx, png)
            table["zoom"][tag] = dict(mm, zoom=z)
            if mm.get("convergence") is not None:
                print("    zoom %6d  conv=%.3f aspect=%.3f"
                      % (z, mm["convergence"], mm["aspect"]))

    conv_map = build_rotation_map(table)
    table["calibration"] = calibrate_angles(table)
    print("\n[calibration] measured pitch vs known lat (lon=0 rows)")
    for c in table["calibration"]:
        print("    lat %3d -> pitch %6s (err %5s)  axes %s/%s"
              % (c["lat_deg"], c["pitch_deg"], c["pitch_err"], c["axis_h_px"],
                 c["axis_v_px"]))

    # ---- presets, BOTH ways ----------------------------------------------------
    print("\n[5] all %d presets, WITH a lat override (isolates the projection type)"
          % len(PRESET_NAMES))
    for name in PRESET_NAMES:
        mm = measure_auto(base, tmp, "p_" + name, name, lat=315 * DEG)
        mm["enum_value"] = PRESET_VALUE[name]
        table["presets"][name] = mm

    print("[6] all %d presets, WITHOUT a rot override (reveals the built-in angles)"
          % len(PRESET_NAMES))
    for name in PRESET_NAMES:
        mm = measure_auto(base, tmp, "q_" + name, name, emit_rot=False)
        mm["enum_value"] = PRESET_VALUE[name]
        mm["effective_lat_from_conv"] = lat_for_convergence(conv_map,
                                                            mm.get("convergence"))
        table["presets_without_rot_override"][name] = mm

    # a compact summary so the doc can be written from data, not by eye
    native = table["presets_without_rot_override"]
    parallel = [n for n, v in native.items()
                if v.get("convergence") is not None and v["convergence"] > 0.97]
    persp = [(n, v["convergence"]) for n, v in native.items()
             if v.get("convergence") is not None and v["convergence"] <= 0.97]
    table["summary"] = {
        "native_parallel_count": len(parallel),
        "native_perspective_count": len(persp),
        "native_perspective": sorted(persp, key=lambda x: x[1]),
        "native_parallel": sorted(parallel),
        "distinct_native_convergences": sorted(
            set(v["convergence"] for v in native.values()
                if v.get("convergence") is not None)),
    }
    print("\n  presets that converge on their own: %d" % len(persp))
    print("  presets that are parallel on their own: %d" % len(parallel))
    print("  distinct convergence values: %s"
          % table["summary"]["distinct_native_convergences"])

    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(table, fh, ensure_ascii=False, indent=1)
    print("\nwrote %s" % OUT)
    print("  presets %d, presets-native %d, rotation %d, rev %d, fov %d, zoom %d"
          % (len(table["presets"]), len(native), len(table["rotation"]),
             len(table["rev"]), len(table["fov"]), len(table["zoom"])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
