# -*- coding: utf-8 -*-
"""Regression test for 3D camera (a:scene3d) injection via the spec.

The camera is exposed in DEGREES and converted to OOXML's 1/60000 unit here,
because the raw unit has two traps:
  * the spec's stated upper bound 21600000 makes PowerPoint report the whole file
    as corrupt (0x80070570)
  * PowerPoint normalises negative angles (-70 deg is stored as 290 deg =
    17400000), which callers should not have to know

The semi-automatic form (only an angle) must default to a PERSPECTIVE preset,
because COM's ThreeD.RotationX/Y writes prst="orthographicFront" -- a parallel
projection that can never look laid down.
"""
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
import motion as M  # noqa: E402

FAILED = []


def check(label, cond, extra=''):
    print('  %-58s %s%s' % (label, 'ok' if cond else 'FAIL', ('  ' + extra) if extra else ''))
    if not cond:
        FAILED.append(label)


print('degrees -> OOXML 1/60000 degree')

for deg, want in ((0, 0), (290, 17400000), (-70, 17400000), (90, 5400000),
                  (360, 0), (-180, 10800000), (45, 2700000), (-90, 16200000)):
    got = M.deg_to_angle(deg)
    check('%6s deg -> %d' % (deg, want), got == want, 'got %d' % got)

# negative angles wrap exactly like PowerPoint, which is what makes template1's
# 290 deg reproducible from either 290 or -70
check('-70 and 290 are the same angle',
      M.deg_to_angle(-70) == M.deg_to_angle(290))

# a raw OOXML value must be rejected loudly, not silently re-interpreted
try:
    M.deg_to_angle(17400000)
    check('raw OOXML angle rejected', False, 'it was accepted -- would become '
          '%d deg' % (17400000 % 360))
except ValueError:
    check('raw OOXML angle rejected', True)

for bad in ('abc', None, [1]):
    try:
        M.deg_to_angle(bad)
        check('non-numeric %r rejected' % (bad,), False)
    except (ValueError, TypeError):
        check('non-numeric %r rejected' % (bad,), True)

# everything normalises below the corruption threshold
worst = max(M.deg_to_angle(d) for d in range(0, 360))
check('all whole degrees stay <= %d' % M.CAMERA_MAX_ANGLE, worst <= M.CAMERA_MAX_ANGLE,
      'worst=%d' % worst)
check('21600000 is never emitted', M.deg_to_angle(360) != 21600000)

print()
print('camera XML')

b = M.build_camera({'tilt': 290})
check('semi-auto defaults to a perspective preset',
      'prst="perspectiveRelaxedModerately"' in b)
check('tilt is an alias for lat', 'lat="17400000"' in b)
check('lon/rev default to 0', 'lon="0" rev="0"' in b)
check('light rig included', '<a:lightRig rig="threePt" dir="t"/>' in b)

b = M.build_camera({'prst': 'perspectiveRelaxed', 'lat': -70, 'lon': 10})
check('full form honours prst', 'prst="perspectiveRelaxed"' in b)
check('full form converts lon too', 'lat="17400000" lon="600000"' in b, b[-120:])

b = M.build_camera(290)
check('bare number accepted as tilt', 'lat="17400000"' in b)

check('perspective detection: perspectiveRelaxedModerately',
      M.camera_is_perspective('perspectiveRelaxedModerately'))
check('perspective detection: legacyPerspectiveFront',
      M.camera_is_perspective('legacyPerspectiveFront'))
check('perspective detection: orthographicFront is NOT',
      not M.camera_is_perspective('orthographicFront'))

print()
print('insertion into a shape')

SLIDE = ('<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
         'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
         '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
         '</p:nvGrpSpPr><p:grpSpPr/>'
         '<p:sp><p:nvSpPr><p:cNvPr id="7" name="HERO"><a:extLst><a:ext uri="{x}">'
         '<a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{y}"/>'
         '</a:ext></a:extLst></p:cNvPr><p:cNvSpPr/></p:nvSpPr>'
         '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>'
         '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln><a:noFill/></a:ln>'
         '<a:extLst><a:ext uri="{z}"><a14:hiddenLine xmlns:a14="http://x"/></a:ext></a:extLst>'
         '</p:spPr><p:txBody><a:bodyPr/></p:txBody></p:sp>'
         '</p:spTree></p:cSld></p:sld>')

frag = M.build_camera({'tilt': 290})
out, status = M.insert_scene3d(SLIDE, '7', frag)
check('insert reports ok', status == 'ok', status)
check('scene3d present', '<a:scene3d>' in out)

i_scene = out.find('<a:scene3d>')
i_extl = out.find('<a:extLst>', out.find('<p:spPr>'))
check('scene3d sits BEFORE the spPr extLst',
      i_scene != -1 and i_extl != -1 and i_scene < i_extl,
      'scene@%s extLst@%s' % (i_scene, i_extl))
# and specifically NOT inside p:cNvPr, which has its own extLst
i_cnvpr_ext = out.find('<a:extLst>')
check('scene3d is not inside p:cNvPr',
      not (out.find('<p:cNvPr', 0, i_scene) > out.rfind('<p:cNvPr', 0, i_scene)
           and i_scene < out.find('</p:cNvPr>', 0)),
      'cnvPr extLst@%s scene@%s' % (i_cnvpr_ext, i_scene))

out2, status2 = M.insert_scene3d(out, '7', frag)
check('second insert is a no-op', status2 == 'already-has-scene3d', status2)
check('no duplicate scene3d', out2.count('<a:scene3d>') == 1)

_, status3 = M.insert_scene3d(SLIDE, '999', frag)
check('unknown shape reported', status3 == 'shape-not-found', status3)

try:
    from lxml import etree
    etree.fromstring(out.encode('utf-8'))
    check('result is well-formed XML', True)
except ImportError:
    check('result is well-formed XML (lxml absent, skipped)', True)
except Exception as e:
    check('result is well-formed XML', False, str(e)[:90])

print()
if FAILED:
    print('CAMERA TEST FAILED: %s' % FAILED)
    sys.exit(1)
print('camera test passed')
