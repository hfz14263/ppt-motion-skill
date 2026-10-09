"""motion.media — 3D 相机与图片填充 —— 两类「非时序」的视觉属性。

3D 相机与图片填充 —— 两类「非时序」的视觉属性。

    · 3D 相机（`a:scene3d`）：角度对外是**度**，内部换算成 OOXML 的 1/60000
    · 图片填充：窗口化 = 给负偏移，把整图塞进形状当前景

模块内容表（内容表；改这一层前先读这里）
----------------------------------------------------------------
    3D 相机      deg_to_angle / build_camera / insert_scene3d / camera_is_perspective
    图片填充     window_insets / build_fill_window / insert_blip_fill
    内容类型     ensure_content_type

> 本模块由 `tools/split_motion.py` 从 `motion.py`（1914 行）按**分层**拆出
> （原文照搬，未改逻辑）。**外部接口由 `../motion.py` 重新导出** ——
> 10 个模块 `import motion` 用的名字一个都没变。
"""

from motion_xml import EXT_CONTENT_TYPES, element_spans, xml_escape

import re


# ---------------------------------------------------------------------------
# 3D camera (a:scene3d)
#
# Angles are exposed in DEGREES and converted to OOXML's 1/60000 unit here.
# Two reasons the raw unit is a trap:
#   * the spec's stated upper bound 21600000 makes PowerPoint report the whole
#     file as corrupt (0x80070570); 360 deg normalises to 0 so we never emit it
#   * PowerPoint normalises negative angles, so -70 deg is stored as 290 deg
#     (17400000) -- accepting degrees means callers never have to know that
# ---------------------------------------------------------------------------
CAMERA_ANGLE_UNIT = 60000


CAMERA_MAX_ANGLE = 21599999          # 21600000 corrupts the file


# COM's ThreeD.RotationX/Y writes prst="orthographicFront", which is a PARALLEL
# projection and therefore can never look "laid down". Default to a perspective
# preset so the semi-automatic form actually produces depth.
CAMERA_DEFAULT_PRESET = "perspectiveRelaxedModerately"


CAMERA_PERSPECTIVE_PREFIXES = ("perspective", "legacyPerspective")


def deg_to_angle(deg, field="angle"):
    """Degrees -> OOXML 1/60000 deg, normalised into [0, 360) like PowerPoint."""
    try:
        v = float(deg)
    except (TypeError, ValueError):
        raise ValueError("%s must be a number in degrees, got %r" % (field, deg))
    if abs(v) > 720:
        raise ValueError(
            "%s=%r looks like a raw OOXML angle (1/60000 deg). Pass DEGREES "
            "(e.g. 290, or -70), not the raw value." % (field, deg))
    v = v % 360.0
    raw = int(round(v * CAMERA_ANGLE_UNIT))
    if raw > CAMERA_MAX_ANGLE:
        raise ValueError("%s=%r normalises to %d, past the safe maximum %d "
                         "(21600000 corrupts the file)"
                         % (field, deg, raw, CAMERA_MAX_ANGLE))
    return raw


def build_camera(spec_dict):
    """Build an <a:scene3d> from a spec entry.

    Semi-automatic form -- only an angle, everything else defaulted:
        {target: HERO, tilt: 290}
    Full form:
        {target: HERO, prst: perspectiveRelaxed, lat: 290, lon: 0, rev: 0,
         lightRig: threePt, lightDir: t}
    """
    if isinstance(spec_dict, (int, float)):
        spec_dict = {"tilt": spec_dict}
    if not isinstance(spec_dict, dict):
        raise ValueError("camera entry must be a mapping or a number, got %r" % (spec_dict,))

    prst = spec_dict.get("prst") or CAMERA_DEFAULT_PRESET
    lat = spec_dict.get("lat", spec_dict.get("tilt", 0))
    lon = spec_dict.get("lon", 0)
    rev = spec_dict.get("rev", 0)

    lat_a = deg_to_angle(lat, "lat/tilt")
    lon_a = deg_to_angle(lon, "lon")
    rev_a = deg_to_angle(rev, "rev")

    return (
        '<a:scene3d><a:camera prst="%s">'
        '<a:rot lat="%d" lon="%d" rev="%d"/>'
        '</a:camera><a:lightRig rig="%s" dir="%s"/></a:scene3d>'
        % (xml_escape(str(prst)), lat_a, lon_a, rev_a,
           xml_escape(str(spec_dict.get("lightRig", "threePt"))),
           xml_escape(str(spec_dict.get("lightDir", "t"))))
    )


def camera_is_perspective(prst):
    return str(prst or "").startswith(CAMERA_PERSPECTIVE_PREFIXES)


def insert_scene3d(xml, shape_id, fragment):
    """Put an <a:scene3d> into the spPr of one shape, by cNvPr id.

    Position matters. CT_ShapeProperties is a SEQUENCE ending
    ... effectLst?, scene3d?, sp3d?, extLst?. So:
      * if the spPr has its own <a:extLst> (PowerPoint adds one holding
        a14:hiddenLine as soon as a line is set), scene3d must go BEFORE it
      * otherwise before </p:spPr> is correct

    Do NOT instead search the whole <p:sp> for <a:extLst>: <p:cNvPr> carries one
    too (a16:creationId), and scene3d inserted there is silently ignored.

    Returns (xml, status).
    """
    for tag in ("sp", "pic", "graphicFrame", "cxnSp"):
        for s, e in element_spans(xml, tag):
            blk = xml[s:e]
            if not re.search(r'<p:cNvPr\b[^>]*\bid="%s"' % re.escape(str(shape_id)), blk):
                continue
            if "<a:scene3d" in blk:
                return xml, "already-has-scene3d"
            sp = element_spans(blk, "spPr")
            if not sp:
                return xml, "no-spPr"
            ps, pe = sp[0]
            extl = re.search(r"<a:extLst(?=[\s/>])", blk[ps:pe])
            at = ps + extl.start() if extl else pe - len("</p:spPr>")
            return xml[:s] + blk[:at] + fragment + blk[at:] + xml[e:], "ok"
    return xml, "shape-not-found"


def window_insets(window, picture, slide_w):
    """Turn a window rectangle into the negative insets that make a fillRect a window.

    THE POINT: `<a:fillRect/>` with no attributes means the picture is STRETCHED to fit
    the shape -- a thumbnail, not a window. To make the shape reveal one part of a
    larger picture, the fill rect must be pushed OUTWARD with negative insets. Units are
    thousandths of a percent OF THE SHAPE's width, and the arithmetic is:

        inset_l = -(window.x - picture.x) / window.w * 100000
        inset_r = -(picture.x + picture.w - window.x - window.w) / window.w * 100000

    In the common case the picture spans the whole slide, which reduces to
    l = -(x / w) * 100000 and r = -((slideW - x - w) / w) * 100000 -- the form written
    down in reference/morph-and-3d-recipes.md section 8.2.

    Taking the general form rather than hard-coding the slide-spanning case is
    deliberate: a window onto a picture that is NOT full-bleed is just as useful, and
    the special case is easy to get wrong silently (a wrong inset produces a squashed
    thumbnail that still renders, so nothing errors).
    """
    x, y, w, h = [float(v) for v in window[:4]]
    px, py, pw, ph = [float(v) for v in picture[:4]]
    if w <= 0 or h <= 0:
        raise ValueError("window must have positive width and height")
    if pw <= 0 or ph <= 0:
        raise ValueError("picture must have positive width and height")
    # vertical insets use the SHAPE'S HEIGHT as the denominator, not its width
    ins_l = -((x - px) / w) * 100000.0
    ins_r = -((px + pw - x - w) / w) * 100000.0
    ins_t = -((y - py) / h) * 100000.0
    ins_b = -((py + ph - y - h) / h) * 100000.0
    return (int(round(ins_l)), int(round(ins_t)),
            int(round(ins_r)), int(round(ins_b)))


def build_fill_window(spec_dict, slide_w=None, slide_h=None):
    """<a:blipFill> fragment that makes a shape a WINDOW onto a picture."""
    win = spec_dict.get("window") or spec_dict.get("bounds")
    if not win or len(win) != 4:
        raise ValueError("fillWindow needs window: [x, y, w, h]")
    pic = spec_dict.get("picture")
    if pic is None:
        if slide_w is None or slide_h is None:
            raise ValueError("fillWindow without 'picture' needs the slide size")
        pic = [0, 0, slide_w, slide_h]
    if len(pic) != 4:
        raise ValueError("fillWindow 'picture' must be [x, y, w, h]")
    l, t, r, b = window_insets(win, pic, slide_w)
    rect = ""
    if any((l, t, r, b)):
        rect = '<a:fillRect l="%d" t="%d" r="%d" b="%d"/>' % (l, t, r, b)
    else:
        rect = "<a:fillRect/>"
    return "<a:blipFill><a:blip/><a:stretch>%s</a:stretch></a:blipFill>" % rect


def insert_blip_fill(xml, shape_id, fragment):
    """Put a <a:blipFill> into one shape, by cNvPr id.

    The parent differs by shape type, and this is the easy thing to get wrong:
      * <p:pic>  is a sequence  nvPicPr, blipFill, spPr  -> blipFill is a CHILD of p:pic
      * <p:sp>   is a sequence  nvSpPr, spPr, style?, txBody?  -> blipFill goes INSIDE
                 spPr, whose own sequence is xfrm?, geom, fill, ln?, effectLst?, ...
                 so it must land after the geometry and before <a:ln>/<a:effectLst>.

    The `<a:blip/>` is left without r:embed on purpose. The picture part has to be added
    to the package and related from the slide, which is a media-layer operation; what
    this function guarantees is the WINDOW GEOMETRY, which is the part that cannot be
    expressed any other way. An embed-less blip is the shape saying "I have a picture
    fill whose window is this", which is exactly the declarative intent.

    Returns (xml, status).
    """
    FILL_TAGS = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
                 "a:grpFill")
    for tag in ("pic", "sp", "cxnSp"):
        for s, e in element_spans(xml, tag):
            blk = xml[s:e]
            if not re.search(r'<p:cNvPr\b[^>]*\bid="%s"' % re.escape(str(shape_id)), blk):
                continue
            if "<a:blipFill" in blk:
                return xml, "already-has-blipFill"
            # A shape has AT MOST ONE fill. CT_ShapeProperties lists the fill choices
            # in a <xsd:choice>, so writing blipFill beside an existing solidFill does
            # not override it -- PowerPoint keeps the first and ignores ours. Measured:
            # a window shape kept its original magenta fill and the picture never
            # appeared, while the XML looked perfectly reasonable and every structural
            # check passed. So the existing fill must be REMOVED, not merely preceded.
            if tag == "pic":
                sp = element_spans(blk, "spPr")
                if not sp:
                    return xml, "no-spPr"
                at = sp[0][0]                     # blipFill precedes spPr in p:pic
                return xml[:s] + blk[:at] + fragment + blk[at:] + xml[e:], "ok"
            # p:sp / p:cxnSp: blipFill goes INSIDE spPr, and CT_ShapeProperties is a
            # SEQUENCE:  xfrm?, geom?, fill?, ln?, effectLst?, effectDag?, scene3d?,
            # sp3d?, extLst?
            #
            # So the fill must land AFTER the geometry and BEFORE <a:ln>. Anchoring on
            # the first of `<a:ln / effectLst / scene3d / ...` and inserting BEFORE it
            # is not enough on its own, because a shape that has a line but no
            # geometry match still needs the fill after the geometry -- and a first
            # attempt here produced `<a:prstGeom/><a:ln/><p:blipFill/>`, which
            # PowerPoint IGNORES IN SILENCE (the fill simply does not paint; there is
            # no repair prompt and no error). Same failure family as the slide-level
            # `transition`-after-`timing` trap already recorded in com-pitfalls.
            sp = element_spans(blk, "spPr")
            if not sp:
                return xml, "no-spPr"
            ps, pe = sp[0]
            inner = blk[ps:pe]

            # Drop the shape's own fill -- and ONLY its own.
            #
            # The obvious implementation scans spPr for the first tag in FILL_TAGS and
            # cuts it. That is wrong, and it corrupted a deck: `<a:ln>` contains its own
            # <a:noFill/> for a no-line shape, and `a:noFill` sorts before `a:solidFill`
            # in the tag list, so the scan removed the LINE's fill and left the shape's
            # solidFill in place. Two fills in one spPr is an invalid xsd:choice, and
            # PowerPoint reports the whole file as corrupt (0x80070570).
            #
            # The unit test missed it because its fixture had an empty <a:ln></a:ln>.
            # A test whose input differs from reality in exactly the place the code
            # looks is not a test.
            #
            # So: find fills that are DIRECT children of spPr. The scan starts inside
            # spPr, so its direct children sit at depth 0 and anything nested goes
            # positive.
            inner2 = inner
            scan = re.compile(r"<(/?)([A-Za-z0-9:]+)([^>]*?)(/?)>")
            depth = 0
            i2 = inner.index(">") + 1
            end2 = inner.rindex("<")
            cut = None
            while i2 < end2:
                m = scan.search(inner, i2)
                if not m or m.start() >= end2:
                    break
                closing, tag, _a, selfclose = m.groups()
                if closing:
                    depth -= 1
                elif depth == 0:
                    if tag in FILL_TAGS:
                        if selfclose:
                            cut = (m.start(), m.end())
                        else:
                            close = inner.find("</%s>" % tag, m.end())
                            cut = (m.start(), close + len(tag) + 3 if close > 0
                                   else m.end())
                        break
                    if not selfclose:
                        depth += 1
                else:
                    if not selfclose:
                        depth += 1
                i2 = m.end()
            if cut:
                inner2 = inner[:cut[0]] + inner[cut[1]:]

            # index just past the last geometry element, or 0 if there is none
            after_geom = 0
            for geom in ("a:prstGeom", "a:custGeom"):
                m = re.search(r"<%s(?=[\s/>])" % re.escape(geom), inner2)
                if not m:
                    continue
                close = inner2.find("</%s>" % geom, m.start())
                if close > 0:
                    after_geom = max(after_geom, close + len(geom) + 3)
                else:
                    after_geom = max(after_geom, m.start())
            # first element that must FOLLOW the fill
            nxt = None
            for later in ("a:ln", "a:effectLst", "a:effectDag", "a:scene3d",
                          "a:sp3d", "a:extLst"):
                m = re.search(r"<%s(?=[\s/>])" % re.escape(later), inner2)
                if m and m.start() >= after_geom and (nxt is None or m.start() < nxt):
                    nxt = m.start()
            if nxt is not None:
                pos = nxt
            elif after_geom:
                pos = after_geom
            else:
                pos = len(inner2) - len("</p:spPr>")
            # inner2 is the spPr body INCLUDING its own opening and closing tags, so a
            # single splice into it produces the whole replacement shape. Splicing into
            # blk instead meant juggling two coordinate systems, which is how the first
            # attempt ended up one offset away from correct.
            body = inner2[:pos] + fragment + inner2[pos:]
            return xml[:s] + blk[:ps] + body + blk[pe:] + xml[e:], "ok"
    return xml, "shape-not-found"


def ensure_content_type(ct_xml, ext):
    ext = (ext or "").lower()
    if not ext or 'Extension="%s"' % ext in ct_xml:
        return ct_xml
    mime = EXT_CONTENT_TYPES.get(ext)
    if not mime:
        return ct_xml
    return ct_xml.replace(
        "</Types>",
        '<Default Extension="%s" ContentType="%s"/></Types>' % (ext, mime),
    )
