"""motion.timing — timing 与 transition 的 XML 生成。

timing 与 transition 的 XML 生成。

    · `<p:timing>` 动画树：效果节点、时长覆盖、spd/dur 换算
    · `<p:transition>` 切换：**挂在终点页**（写在第 N 页，动的是「进入第 N 页」）
    · mc:AlternateContent 包装：PowerPoint 自己写的形式，不是我们发明的

模块内容表（内容表；改这一层前先读这里）
----------------------------------------------------------------
    timing      build_effect_node / build_timing / apply_overrides
    时长换算     _spd_for_ms / _spd_for_label / SPD_NOTCH_MS
    Morph       MORPH_TEMPLATE / MORPH_OPTIONS / MORPH_SPEEDS
    切换         build_transition / drop_transition_alternate_content

> 本模块由 `tools/split_motion.py` 从 `motion.py`（1914 行）按**分层**拆出
> （原文照搬，未改逻辑）。**外部接口由 `../motion.py` 重新导出** ——
> 10 个模块 `import motion` 用的名字一个都没变。
"""

from motion_xml import DIRECTIONAL_FILTERS, TRANSITIONS, TRIGGERS, WIPE_DIRS

import re


# ---------------------------------------------------------------------------
# timing XML generation
# ---------------------------------------------------------------------------
class NodeIds(object):
    def __init__(self, start=4):
        self.n = start

    def next(self):
        self.n += 1
        return self.n


def split_effect_template(tpl):
    """Split a catalogue template into (inner children, preset attributes).

    The catalogue templates are verbatim PowerPoint output, which wraps the effect
    in an extra preset <p:cTn nodeType="afterEffect"> layer. PowerPoint refuses to
    open a deck where that wrapper sits directly under a bare timing <p:cTn>, so we
    keep only the children and lift the preset attributes onto the skeleton cTn
    that build_timing() generates. Verified against real PowerPoint: the inline
    form opens, the wrapped form does not.
    """
    m = re.search(r"<p:cTn\b([^>]*)>", tpl)
    if not m:
        return tpl, ""
    attrs = m.group(1).strip()
    gt = m.end()
    close = tpl.rfind("</p:cTn>")
    body = tpl[gt:close] if close != -1 else tpl[gt:]
    inner = re.search(r"<p:childTnLst\b[^>]*>(.*)</p:childTnLst>\s*$", body, re.S)
    children = inner.group(1) if inner else body
    return children, attrs


def build_effect_node(tpl, shape_id, ids):
    """Instantiate a template's effect children, renumbering local timing ids."""
    children, attrs = split_effect_template(tpl)
    local = {}

    def place(m):
        key = m.group(1)
        if key not in local:
            local[key] = str(ids.next())
        return local[key]

    node = re.sub(r"\{N(\d+)\}", place, children)
    return node.replace("{SHAPE_ID}", str(shape_id)), attrs


def set_wipe_direction(node, direction):
    """Rewrite the direction inside every <p:animEffect filter="..."> value.

    Only the direction token changes; the filter *family* (wipe / barn / wheel)
    is whatever the catalogue alias chose. `wipe(down)` -> `wipe(left)` is the
    same presetID/presetClass/presetSubtype, which matters: PowerPoint validates
    the preset triple, and folding the direction into the preset instead of the
    filter is what makes a deck fail to open.
    """
    d = str(direction).strip().lower()
    if d not in WIPE_DIRS:
        raise ValueError("unknown wipe direction %r; known: %s"
                         % (direction, list(WIPE_DIRS)))

    fam = re.search(r'<p:animEffect\b[^>]*\bfilter="([A-Za-z]+)', node)
    if not fam or fam.group(1) not in DIRECTIONAL_FILTERS:
        raise ValueError(
            "filter %r takes no direction, so dir=%r cannot apply; directional "
            "filters are %s"
            % (fam.group(1) if fam else None, direction,
               list(DIRECTIONAL_FILTERS)))

    pat = re.compile(r'(<p:animEffect\b[^>]*\bfilter=")([A-Za-z]+)'
                     r'(?:\([^)]*\))?(")')

    def fix(m):
        return "%s%s(%s)%s" % (m.group(1), m.group(2), d, m.group(3))

    return pat.sub(fix, node)


def set_leaf_durations(node, duration):
    """Set the visual duration on the leaf behaviour cTn nodes (milliseconds).

    PowerPoint keeps the effect duration on the leaf behaviour <p:cTn> while the
    root preset <p:cTn> carries fill="hold". The visibility <p:set> behaviour must
    keep dur="1" so the shape appears instantly instead of fading in.
    """
    if duration is None:
        return node
    ms = max(1, int(round(float(duration) * 1000)))

    def per_block(m):
        def tweak(cm):
            head = cm.group(0)
            dm = re.search(r'dur="(\d+)"', head)
            if not dm or dm.group(1) == "1":
                return head
            return head[:dm.start()] + 'dur="%d"' % ms + head[dm.end():]
        return re.sub(r"<p:cTn\b[^>]*>", tweak, m.group(0))

    return re.sub(
        r"<p:animEffect\b.*?</p:animEffect>"
        r"|<p:anim\b.*?</p:anim>"
        r"|<p:animRot\b.*?</p:animRot>"
        r"|<p:animMotion\b.*?</p:animMotion>"
        r"|<p:animScale\b.*?</p:animScale>"
        r"|<p:animClr\b.*?</p:animClr>",
        per_block, node, flags=re.S,
    )


def set_anchor_durations(node, duration):
    """Kept for API stability; the preset cTn no longer carries dur."""
    return node


def apply_overrides(attrs, eff):
    """Apply repeat / autoReverse / smooth onto the preset attribute string."""
    head = attrs

    def set_attr(head, attr, value):
        if re.search(r"\b%s=\"" % attr, head):
            return re.sub(r'\b%s="[^"]*"' % attr, '%s="%s"' % (attr, value), head)
        return (head + ' %s="%s"' % (attr, value)).strip()

    if eff.get("repeat") is not None:
        head = set_attr(head, "repeatCount", int(eff["repeat"]) * 1000)
    if eff.get("autoReverse") is not None:
        head = set_attr(head, "autoRev", "1" if eff["autoReverse"] else "0")
    if eff.get("smooth") is not None:
        v = int(float(eff["smooth"]) * 100000)
        head = set_attr(head, "accel", v)
        head = set_attr(head, "decel", v)
    return head


def build_timing(entries, ids):
    """entries: [{shape_id, tpl, effect, delay, attrs}] -> a full <p:timing> block.

    Effect children are inlined under the generated preset cTn. The catalogue's
    extra wrapper cTn is deliberately dropped: PowerPoint rejects the wrapped form.

    ID ALLOCATION ORDER MATTERS. <p:cTn id> is not just a key: PowerPoint rebuilds
    the timeline from these nodes and each node must be numbered after its own
    ancestors, so ids have to increase in document order (parent, then its
    children). Allocating a wrapper's two ids *after* instantiated the effect's
    children produced a sequence like

        ..., 12, 13, 10, 11, 15, 16, 14

    where the last effect's animation node (14) precedes its own parent cage
    (15, 16). PowerPoint loads such a slide and silently DROPS that effect -- no
    repair prompt, no error, the animation is simply gone after a save. Taking the
    cage ids first, then instantiating the children, keeps the whole slide
    monotonic and every effect survives the round-trip.
    """
    pars = []
    bld = []
    seen_bld = set()
    for ent in entries:
        eff = ent["effect"]
        o = ids.next()      # outer <p:par> wrapper cTn
        i = ids.next()      # inner preset cTn  -- before its children, see above
        node, attrs = build_effect_node(ent["tpl"], ent["shape_id"], ids)
        node = set_leaf_durations(node, eff.get("duration"))
        if eff.get("dir"):
            # set_wipe_direction validates both the direction name and whether
            # this effect's filter can carry one. Failing here is deliberate: a
            # silently ignored `dir` produces a deck whose motion is not the one
            # that was asked for, which is the failure mode this skill exists to
            # prevent.
            node = set_wipe_direction(node, eff["dir"])
        attrs = apply_overrides(attrs, eff)
        # the generated <p:cTn> already owns id/dur/fill; never let the template
        # wrapper's own copies through (a duplicate attribute is malformed XML)
        attrs = re.sub(r'\b(id|dur|fill)="[^"]*"\s*', "", attrs).strip()
        trigger = str(eff.get("trigger", "after")).lower()
        node_type = TRIGGERS.get(trigger, "afterEffect")
        attrs = re.sub(r'nodeType="[^"]*"', 'nodeType="%s"' % node_type, attrs)
        if 'nodeType="' not in attrs:
            attrs = (attrs + ' nodeType="%s"' % node_type).strip()
        attrs = re.sub(r"\s+", " ", attrs).strip()
        delay_ms = int(round(float(ent.get("delay") or 0.0) * 1000))
        if node_type == "withEffect":
            delay_ms = int(round(float(eff.get("delay") or 0.0) * 1000))
        pars.append(
            '<p:par><p:cTn id="%d" fill="hold">'
            '<p:stCondLst><p:cond delay="%d"/></p:stCondLst>'
            '<p:childTnLst><p:par><p:cTn id="%d" fill="hold"%s>'
            '<p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            '<p:childTnLst>%s</p:childTnLst>'
            '</p:cTn></p:par></p:childTnLst>'
            '</p:cTn></p:par>' % (
                o, delay_ms, i,
                (" " + attrs) if attrs else "",
                node,
            )
        )
        sid = str(ent["shape_id"])
        if sid not in seen_bld:
            seen_bld.add(sid)
            bld.append('<p:bldP spid="%s" grpId="0" animBg="1"/>' % sid)
    if not pars:
        return ""
    bld_lst = ("<p:bldLst>%s</p:bldLst>" % "".join(bld)) if bld else ""
    return (
        '<p:timing><p:tnLst><p:par>'
        '<p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">'
        '<p:childTnLst><p:seq concurrent="1" nextAc="seek">'
        '<p:cTn id="2" dur="indefinite" nodeType="mainSeq">'
        '<p:childTnLst>%s</p:childTnLst></p:cTn>'
        '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
        '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>'
        '</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>%s</p:timing>' % ("".join(pars), bld_lst)
    )


def root_declares(xml, prefix):
    """True only when the document root element itself binds this prefix.

    A descendant may declare the namespace locally (PowerPoint writes
    xmlns:p14 on <p14:creationId> inside p:extLst), and that binding does NOT
    cover the root's other children. A substring test would be fooled by it.
    """
    head = re.search(r"<\w+:[A-Za-z0-9]+\b[^>]*>", xml)
    if not head:
        return False
    return ("xmlns:%s=" % prefix) in head.group(0)


def add_root_namespace(xml, prefix, uri):
    head = re.search(r"<\w+:[A-Za-z0-9]+\b[^>]*>", xml)
    if not head:
        return xml
    tag = head.group(0)
    # the root element has no children yet, so its first '>' closes the start tag
    at = tag.rfind(">")
    if at == -1:
        return xml
    if tag[at - 1] == "/":
        at -= 1
    new_tag = tag[:at] + ' xmlns:%s="%s"' % (prefix, uri) + tag[at:]
    return xml[:head.start()] + new_tag + xml[head.end():]


# Morph (平滑) is NOT a simple <p:transition> child: it lives in the p159
# extension namespace and must be wrapped in mc:AlternateContent with an
# mc:Fallback, otherwise PowerPoint treats it as an unknown element and silently
# drops it -- which is exactly how it was once misdiagnosed as "unsupported".
# p14 is declared locally here so no root-namespace fixup is needed.
#
# NOTE 1 -- both branches share the SAME {spd} token, and {extra} sits in BOTH
# <p:transition> elements. A duration or an auto-advance that only reaches the
# Choice branch is invisible here and wrong in the fallback world (WPS / old
# PowerPoint / online preview) -- see _spd_for_ms.
MORPH_TEMPLATE = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
    '<p:transition spd="{spd}" p14:dur="{dur}"{extra}>'
    '<p159:morph option="{option}"/>'
    '</p:transition>'
    '</mc:Choice>'
    '<mc:Fallback><p:transition spd="{spd}"{extra}>'
    '<p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)


MORPH_OPTIONS = ("byobject", "byword", "bychar")


MORPH_SPEEDS = {"slow": "slow", "med": "med", "fast": "fast"}


# What `spd` means in milliseconds, measured two independent ways.
#
#   (1) AS A LENGTH. A transition carrying ONLY spd="" (no p14:dur at all, which
#       is what mc:Fallback can hold) grows/dissolves for this long
#       (build_transition_table.py anchordeck -> notch.pptx; the same at 2s and
#       5s per slide, so it is the transition and not the dwell):
#             fast ~500ms   med ~767ms   slow ~1000ms
#
#   (2) AS A READ-BACK. SlideShowTransition.Duration, after opening an applied
#       deck, reports 0.8s / 0.6s / 0.9s exactly as written -- the millisecond
#       survives the round-trip. So spd is NOT a second copy of the duration
#       competing with it; it is the coarse notch, and p14:dur always wins.
#
# FAILURE MODE WORTH KNOWING: PowerPoint recomputes the Choice branch's spd on
# save. A written `spd="med" p14:dur="800"` comes back as `spd="slow"` with the
# 800ms intact -- it rounds the duration to its nearest notch, and 0.8s is
# closer to 1.0s than to 0.767s. Nothing is refused, nothing is lost, and every
# structural check still passes; a round-trip diff just shows spd moving.
#   * writing spd="800" instead makes PowerPoint REFUSE the whole deck
#     (HRESULT E_FAIL) -- do not reach for the numeric form because a morph's
#     saved spd happens to be numeric; that one comes out of a COM value in
#     SECONDS on a different element.
SPD_NOTCH_MS = (("fast", 500.0), ("med", 767.0), ("slow", 1000.0))


def _spd_for_ms(ms, speed=None):
    """Pick the spd notch closest to a millisecond duration.

    `p14:dur` is the field that actually drives the render, so this only has to
    be honest in the one place p14:dur cannot reach: mc:Fallback, where spd IS
    the duration. Picking the nearest measured notch keeps the two branches
    telling the same story there.

    An explicit `speed` wins -- then the caller is asking for a label by name,
    and the notch values are the ones in SPD_NOTCH_MS.
    """
    if speed:
        return _spd_for_label(speed)
    return min(SPD_NOTCH_MS, key=lambda kv: abs(kv[1] - ms))[0]


def _spd_for_label(speed):
    got = MORPH_SPEEDS.get(str(speed).strip().lower())
    if got is None:
        raise ValueError("speed must be slow/med/fast, got %r" % speed)
    return got


def build_transition(name, duration=None, advance_after=None, on_click=None,
                     option=None, speed=None):
    key = str(name or "none").strip().lower()

    if key == "morph":
        opt = str(option or "byObject").strip()
        if opt.lower() not in MORPH_OPTIONS:
            raise ValueError("morph option must be one of byObject/byWord/byChar, got %r" % option)
        opt = {"byobject": "byObject", "byword": "byWord", "bychar": "byChar"}[opt.lower()]
        ms = int(round(float(duration if duration is not None else 2.0) * 1000))
        spd = _spd_for_ms(ms, speed)
        extra = ""
        if advance_after:
            extra += ' advTm="%d"' % int(round(float(advance_after) * 1000))
        if on_click is False:
            extra += ' advClick="0"'
        return MORPH_TEMPLATE.format(dur=ms, option=opt, spd=spd, extra=extra)

    if key not in TRANSITIONS:
        raise ValueError("unknown transition %r; known: %s, morph" % (name, sorted(TRANSITIONS)))
    block = TRANSITIONS[key]
    if not block:
        return ""
    # `extra` (advTm / advClick) has to be merged BEFORE .format() runs, not
    # pasted in afterwards. Pasting first and formatting second is invisible
    # until extra is set: the literal "{spd}" then lives in the element that
    # the paste hit, gets replaced with a non-attribute string, and the file
    # PowerPoint writes back carries the previous spd. Nothing raises.
    extra = ""
    if advance_after:
        extra += ' advTm="%d"' % int(round(float(advance_after) * 1000))
    if on_click is False:
        extra += ' advClick="0"'
    # p14:dur is the millisecond-accurate duration. The measured block declares
    # xmlns:p14 itself, so there is no root-namespace fixup to attempt here and
    # no case where the duration has to be dropped. `extra` is merged into BOTH
    # branches (there is no count=1 here) so an auto-advance is not a
    # 2010-only feature.
    ms = int(round(float(duration if duration is not None else 0.7) * 1000))
    return block.replace("<p:transition", "<p:transition" + extra).format(
        spd=_spd_for_ms(ms, speed), dur=' p14:dur="%d"' % ms)


def drop_alternate_content(xml):
    """Remove every mc:AlternateContent block (used before re-writing morph)."""
    while True:
        m = re.search(r"<mc:AlternateContent(?=[\s/>])", xml)
        if not m:
            return xml
        gt = xml.find(">", m.end())
        if gt == -1:
            return xml
        if xml[gt - 1] == "/":
            xml = xml[:m.start()] + xml[gt + 1:]
            continue
        close = xml.find("</mc:AlternateContent>", gt)
        if close == -1:
            return xml
        xml = xml[:m.start()] + xml[close + len("</mc:AlternateContent>"):]


def drop_transition_alternate_content(xml):
    """Remove mc:AlternateContent blocks that wrap a <p:transition>, keep the rest.

    drop_alternate_content() takes every AlternateContent in the slide, which is
    too blunt to run on every transition: a slide can carry unrelated ones
    (extension lists, media). This one only removes a wrapper whose body holds a
    <p:transition>, so an already-processed deck can be re-processed without
    accumulating a second transition -- while everything else is left alone.
    """
    while True:
        m = re.search(r"<mc:AlternateContent(?=[\s/>])", xml)
        if not m:
            return xml
        gt = xml.find(">", m.end())
        if gt == -1:
            return xml
        if xml[gt - 1] == "/":
            end = gt + 1
            body = xml[m.start():end]
        else:
            close = xml.find("</mc:AlternateContent>", gt)
            if close == -1:
                return xml
            end = close + len("</mc:AlternateContent>")
            body = xml[m.start():end]
        if "<p:transition" in body:
            xml = xml[:m.start()] + xml[end:]
        else:
            # not ours -- step past it and keep looking
            rest = drop_transition_alternate_content(xml[end:])
            return xml[:end] + rest
