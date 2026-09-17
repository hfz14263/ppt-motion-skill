#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the 3D camera reference table: parameter -> MEASURED rendered result.

WHY A MEASURED TABLE AND NOT A DOC TABLE
----------------------------------------
The official sources give the schema but not the behaviour, and where they do give
behaviour they are not always right:

  ECMA-376, CT_Camera
      <a:camera prst="..." fov="..." zoom="..."><a:rot lat lon rev/></a:camera>
      fov   ST_FOVAngle            0..180, in 1/60000 deg
      zoom  ST_PositivePercentage  in 1/1000 percent, default 100000
      rot   CT_SphereCoords        lat/lon/rev in 0..21600000

  [MS-OI29500], camera
      "In Office, the fov attribute is ignored for non-perspective camera types."

  MsoPresetCamera enum (Microsoft.Office.Core), values 1..62, official names.

Measured against that:
  * rot upper bound 21600000 -- the spec's own maximum -- makes PowerPoint report the
    whole file as CORRUPT (0x80070570). The real maximum is 21599999. A table built
    from the spec would therefore hand out a file-destroying value.
  * fov DOES change the render on a perspective camera (convergence 0.826 -> 0.473 as
    fov goes 45 -> 150), and is genuinely ignored on a non-perspective one, which does
    match [MS-OI29500].
  * zoom beyond the documented percentage range is accepted without complaint.

So the table records what the implementation DOES, with the doc value alongside so the
disagreements are visible rather than buried.

MEASUREMENT
-----------
A white plane on a dark ground, rendered by PowerPoint, silhouette segmented by
threshold, then three numbers:

    convergence = far_edge_width / near_edge_width   1.000 means PARALLEL projection;
                                                      < 1 means real perspective
    aspect      = silhouette_height / mid_width      how squashed the plane is
    skew        = |left_edge_len - right_edge_len| / max   how much it leans sideways

Convergence is the discriminating feature. Picking a pose by "does it look like a
parallelogram" fails, because a parallel-projected rectangle looks like one at every
angle.

Usage:
    python build_camera_table.py                # writes camera_reference.json
    python build_camera_table.py --quick        # presets only, skip the sweeps
"""
import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import zipfile

import numpy as np
from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "camera_reference.json")

DEG = 60000
MAX_ROT = 21599999              # measured; the spec says 21600000 and that corrupts
MAX_FOV = 10800000              # 180 deg, matches the spec
DEFAULT_ZOOM = 100000

# The 62 official preset names, from Microsoft.Office.Core.MsoPresetCamera (1..62).
# Ordered by value so the index+1 is the enum value.
PRESET_NAMES = [
    # 1..9 legacy oblique
    "legacyObliqueTopLeft", "legacyObliqueTop", "legacyObliqueTopRight",
    "legacyObliqueLeft", "legacyObliqueFront", "legacyObliqueRight",
    "legacyObliqueBottomLeft", "legacyObliqueBottom", "legacyObliqueBottomRight",
    # 10..18 legacy perspective
    "legacyPerspectiveTopLeft", "legacyPerspectiveTop", "legacyPerspectiveTopRight",
    "legacyPerspectiveLeft", "legacyPerspectiveFront", "legacyPerspectiveRight",
    "legacyPerspectiveBottomLeft", "legacyPerspectiveBottom",
    "legacyPerspectiveBottomRight",
    # 19 orthographic
    "orthographicFront",
    # 20..39 isometric
    "isometricTopUp", "isometricTopDown", "isometricBottomUp", "isometricBottomDown",
    "isometricLeftUp", "isometricLeftDown", "isometricRightUp", "isometricRightDown",
    "isometricOffAxis1Left", "isometricOffAxis1Right", "isometricOffAxis1Top",
    "isometricOffAxis2Left", "isometricOffAxis2Right", "isometricOffAxis2Top",
    "isometricOffAxis3Left", "isometricOffAxis3Right", "isometricOffAxis3Bottom",
    "isometricOffAxis4Left", "isometricOffAxis4Right", "isometricOffAxis4Bottom",
    # 40..47 oblique
    "obliqueTopLeft", "obliqueTop", "obliqueTopRight", "obliqueLeft", "obliqueRight",
    "obliqueBottomLeft", "obliqueBottom", "obliqueBottomRight",
    # 48..62 modern perspective
    "perspectiveFront", "perspectiveLeft", "perspectiveRight", "perspectiveAbove",
    "perspectiveBelow", "perspectiveAboveLeftFacing", "perspectiveAboveRightFacing",
    "perspectiveContrastingLeftFacing", "perspectiveContrastingRightFacing",
    "perspectiveHeroicLeftFacing", "perspectiveHeroicRightFacing",
    "perspectiveHeroicExtremeLeftFacing", "perspectiveHeroicExtremeRightFacing",
    "perspectiveRelaxed", "perspectiveRelaxedModerately",
]
PRESET_VALUE = {name: i + 1 for i, name in enumerate(PRESET_NAMES)}

SLIDE = "ppt/slides/slide1.xml"


def make_probe_deck(path, scratch=None):
    """A white plane on a dark ground: the silhouette is then unambiguous.

    NOTE ON REPRODUCIBILITY: the measured convergence depends on the probe geometry --
    where the plane sits in the frame and how big it is -- because convergence is a
    ratio between two edges of a projected shape. So the ABSOLUTE numbers in a table
    are only comparable with numbers measured from the SAME probe. The findings that
    transfer are the relative ones: which presets are parallel, the ordering of the
    orientation sweep, and the monotonic response to fov.

    The script records the probe geometry in the output so a later run can be compared
    like for like.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from pptx.dml.color import RGBColor

    plane = os.path.join(scratch or HERE, "_probe_plane.png")
    if not os.path.exists(plane):
        Image.new("RGB", (1600, 900), (255, 255, 255)).save(plane)

    prs = Presentation()
    prs.slide_width = Inches(13.3333)
    prs.slide_height = Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    bg = sl.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0x18, 0x18, 0x18)
    bg.line.fill.background()
    bg.shadow.inherit = False
    # The probe plane is deliberately small and centred. An earlier version used a
    # large one, which left no room for zoom/fov to enlarge it, so those cases scaled
    # out of frame and produced meaningless numbers (see measure()).
    sl.shapes.add_picture(plane, Inches(4.75), Inches(3.30),
                          width=Inches(3.83)).name = "PLANE"
    prs.save(path)
    return path


def write_camera(base, out, prst, lat=0, lon=0, rev=0, fov=None, zoom=None):
    attrs = ' prst="%s"' % prst
    if fov is not None:
        attrs += ' fov="%d"' % int(fov)
    if zoom is not None:
        attrs += ' zoom="%d"' % int(zoom)
    cam = ('<a:scene3d><a:camera%s><a:rot lat="%d" lon="%d" rev="%d"/></a:camera>'
           '<a:lightRig rig="threePt" dir="t"/></a:scene3d>'
           % (attrs, int(lat), int(lon), int(rev)))
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
  $p.Slides(1).Export("__PNG__", "PNG", 1200, 675)
  $p.Close()
  Write-Output "OK"
} catch {
  Write-Output ("FAIL " + $_.Exception.Message)
} finally { $app.Quit() }
'''


def measure(pptx, png, w=1200, h=675):
    """Render and measure, refusing to report a number when the silhouette is clipped.

    MEASUREMENT VALIDITY IS PART OF THE MEASUREMENT. An early version of this reported
    convergence 1.000 for zoom 200000 and above, which looked like "zoom turns the
    projection parallel". It was not: the plane had been scaled past the frame, so the
    silhouette became the image border, and near_w == far_w == image width. The same
    artefact silently corrupted the fov sweep (fov 120/150/180 all read near_w = 1200).

    So a silhouette that touches any edge is reported as clipped and carries NO
    convergence value, rather than a number that means nothing.
    """
    r = subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
                        "-Command", PS_RENDER.replace("__SRC__", pptx.replace("\\", "\\\\"))
                                      .replace("__PNG__", png.replace("\\", "\\\\"))],
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
    denom = max(llen, rlen, 1)

    res = {"opens": True, "clipped": clipped,
           "near_w": int(near), "far_w": int(far), "height": int(y1 - y0)}
    if clipped:
        # no meaningful geometry: the silhouette is the frame, not the shape
        res["unmeasurable"] = ("silhouette touches the frame edge; the shape has "
                               "scaled out of view")
        return res
    res.update({"convergence": round(far / near, 4) if near else None,
                "aspect": round((y1 - y0) / float(mid or 1), 4),
                "skew": round(abs(llen - rlen) / float(denom), 4)})
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--quick", action="store_true", help="presets only")
    ns = ap.parse_args()

    tmp = tempfile.mkdtemp(prefix="camtable-")
    base = make_probe_deck(os.path.join(tmp, "base.pptx"), scratch=tmp)

    table = {
        "meta": {
            "measured_on": "PowerPoint 16.0 build 20228",
            "method": ("white plane on a dark ground, exported by PowerPoint, "
                       "silhouette thresholded; convergence = far_edge/near_edge where "
                       "1.000 means parallel projection"),
            "probe_geometry": {
                "plane_position_in": [4.75, 3.30],
                "plane_width_in": 3.83,
                "slide_in": [13.3333, 7.5],
                "export_px": [1200, 675],
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
                "schema": "ECMA-376 Part 1, CT_Camera / CT_SphereCoords / ST_FOVAngle",
                "office_behaviour": "[MS-OI29500] Part 1 s.20.1.5.5 camera",
                "preset_names": "Microsoft.Office.Core.MsoPresetCamera, values 1..62",
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
        "rotation": {},
        "fov": {},
        "zoom": {},
    }

    # ---- every preset, at a fixed downward camera -------------------------------
    print("[1] all %d presets" % len(PRESET_NAMES))
    for name in PRESET_NAMES:
        pptx = os.path.join(tmp, "p_%s.pptx" % name)
        png = os.path.join(tmp, "p_%s.png" % name)
        write_camera(base, pptx, name, lat=18600000)
        mm = measure(pptx, png)
        mm["enum_value"] = PRESET_VALUE[name]
        table["presets"][name] = mm
        if mm.get("opens") and mm.get("convergence") is not None:
            print("    %-40s conv=%.3f" % (name, mm["convergence"]))

    if ns.quick:
        with open(OUT, "w", encoding="utf-8") as fh:
            json.dump(table, fh, ensure_ascii=False, indent=1)
        print("\nwrote %s" % OUT)
        return 0

    # ---- rotation sweep on the modern perspective preset ------------------------
    print("\n[2] rotation (lat x lon) on perspectiveFront")
    for lat_deg in (330, 315, 300, 285, 270):
        for lon_deg in (0, 30, 60):
            tag = "lat%d_lon%d" % (lat_deg, lon_deg)
            pptx = os.path.join(tmp, "r_%s.pptx" % tag)
            png = os.path.join(tmp, "r_%s.png" % tag)
            write_camera(base, pptx, "perspectiveFront",
                         lat=lat_deg * DEG, lon=lon_deg * DEG)
            mm = measure(pptx, png)
            table["rotation"][tag] = dict(mm, lat_deg=lat_deg, lon_deg=lon_deg)
            if mm.get("opens") and mm.get("convergence") is not None:
                print("    lat %3d lon %3d  conv=%.3f aspect=%.3f skew=%.3f"
                      % (lat_deg, lon_deg, mm["convergence"], mm["aspect"], mm["skew"]))

    # ---- fov sweep --------------------------------------------------------------
    print("\n[3] fov on perspectiveFront (lat 310)")
    for fov_deg in (0, 15, 45, 90, 120, 150, 180):
        tag = "fov%d" % fov_deg
        pptx = os.path.join(tmp, "f_%s.pptx" % tag)
        png = os.path.join(tmp, "f_%s.png" % tag)
        write_camera(base, pptx, "perspectiveFront", lat=18600000, fov=fov_deg * DEG)
        mm = measure(pptx, png)
        table["fov"][tag] = dict(mm, fov_deg=fov_deg)
        if mm.get("opens") and mm.get("convergence") is not None:
            print("    fov %3d  conv=%.3f aspect=%.3f"
                  % (fov_deg, mm["convergence"], mm["aspect"]))

    # ---- zoom sweep ------------------------------------------------------------
    print("\n[4] zoom on perspectiveFront (lat 310)")
    for z in (25000, 50000, 100000, 200000, 400000):
        tag = "zoom%d" % z
        pptx = os.path.join(tmp, "z_%s.pptx" % tag)
        png = os.path.join(tmp, "z_%s.png" % tag)
        write_camera(base, pptx, "perspectiveFront", lat=18600000, zoom=z)
        mm = measure(pptx, png)
        table["zoom"][tag] = dict(mm, zoom=z)
        if mm.get("opens") and mm.get("convergence") is not None:
            print("    zoom %6d  conv=%.3f aspect=%.3f"
                  % (z, mm["convergence"], mm["aspect"]))

    with open(OUT, "w", encoding="utf-8") as fh:
        json.dump(table, fh, ensure_ascii=False, indent=1)
    print("\nwrote %s (%d presets, %d rotation, %d fov, %d zoom)"
          % (OUT, len(table["presets"]), len(table["rotation"]),
             len(table["fov"]), len(table["zoom"])))
    return 0


if __name__ == "__main__":
    sys.exit(main())
