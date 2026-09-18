#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: OOXML motion injection engine.

Injects entrance / emphasis / motion-path animations, slide transitions and
animation-stripping into an existing .pptx by patching only the parts that carry
no geometry:

    ppt/slides/slideN.xml      <p:timing> and <p:transition> (replaced)
    [Content_Types].xml        media extension defaults (additive)

Shape geometry, text bodies, fills and the spTree element order are never
touched, so a layout that already passed a pre-flight gate stays byte-identical
inside the layout frame. `--assert-geometry` proves exactly that.

Effect parameters come from motion_catalog.json, which was extracted from real
PowerPoint output (presetID / presetClass / presetSubtype plus the exact timing
XML each effect emits), not from a guessed enum table.

Usage:
    python motion.py apply   --pptx IN.pptx --spec motion.yaml --out OUT.pptx [--assert-geometry]
    python motion.py inspect --pptx IN.pptx [--json]
    python motion.py preview --pptx IN.pptx --out STATIC.pptx
    python motion.py catalog [--kind entrance|emphasis|path]
"""
import argparse
import hashlib
import json
import os
import re
import sys
import zipfile

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

try:
    from lxml import etree as _ET
except ImportError:  # pragma: no cover
    _ET = None

HERE = os.path.dirname(os.path.abspath(__file__))
CATALOG_PATH = os.path.join(HERE, "motion_catalog.json")

SLIDE_RE = re.compile(r"^ppt/slides/slide(\d+)\.xml$")
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
MC_NS = "{http://schemas.openxmlformats.org/markup-compatibility/2006}"
SHAPE_TAGS = ("sp", "pic", "graphicFrame", "cxnSp", "grpSp")
TRIGGERS = {"click": "clickEffect", "with": "withEffect", "after": "afterEffect"}
GEOM_ATTRS = ("x", "y", "cx", "cy", "rot")

# Wipe-family filter directions. PowerPoint's `wipe` entrance is `wipe(down)`
# (presetSubtype=4), which is fine for a card but wrong for a data line: a
# left-to-right wipe is what makes a line read as *drawn*. Direction lives only
# inside the <p:animEffect filter="..."> value, and `wipe(left)` has no presetID
# of its own, so this is a per-effect override rather than a catalogue alias.
WIPE_DIRS = ("down", "left", "right", "up",
             "upleft", "upright", "downleft", "downright")

# Filter families that accept a direction argument. `fade` and `dissolve` are
# filters too, but they are directionless, so asking for a direction on them has
# to fail rather than quietly animate the default.
DIRECTIONAL_FILTERS = ("wipe", "barn", "checkerboard", "circle", "diamond",
                       "plus", "wheel", "wedge", "strips")

TRANSITIONS = {
    "fade": '<p:transition spd="med"{dur}><p:fade/></p:transition>',
    "smoothfade": '<p:transition spd="med"{dur}><p:fade/></p:transition>',
    "fadeblack": '<p:transition spd="med"{dur}><p:fade thruBlk="1"/></p:transition>',
    "push": '<p:transition spd="med"{dur}><p:push dir="u"/></p:transition>',
    "pushleft": '<p:transition spd="med"{dur}><p:push dir="l"/></p:transition>',
    "wipe": '<p:transition spd="med"{dur}><p:wipe dir="l"/></p:transition>',
    "cover": '<p:transition spd="med"{dur}><p:cover dir="l"/></p:transition>',
    "split": '<p:transition spd="med"{dur}><p:split orient="horz" dir="out"/></p:transition>',
    "zoom": '<p:transition spd="med"{dur}><p:zoom dir="in"/></p:transition>',
    "dissolve": '<p:transition spd="med"{dur}><p:dissolve/></p:transition>',
    "strips": '<p:transition spd="med"{dur}><p:strips/></p:transition>',
    "pull": '<p:transition spd="med"{dur}><p:pull/></p:transition>',
    "randombar": '<p:transition spd="med"{dur}><p:randomBar/></p:transition>',
    "none": "",
}
# Legacy SlideShowTransition.EntryEffect values. NOTE: this build of PowerPoint
# maps them far from the obvious preset numbers (0x0A01 "fade" emits <p:strips/>,
# 0x0901 emits <p:randomBar/>, 0x0C01 emits <p:zoom/>). The XML element written by
# build_transition is authoritative; these are only used to let the review layer
# report a non-zero transition via the COM property model.
TRANSITION_ENUM = {
    "fade": 0x0A01, "smoothfade": 0x0A01, "strips": 0x0A01,
    "push": 0x0901, "randombar": 0x0901,
    "wipe": 0x0801, "pull": 0x0801,
    "cover": 0x0C01, "zoom": 0x0C01,
    "split": 0x0701, "fadeblack": 0x0B01,
    "dissolve": 0x0D01,
    "none": 0,
}
EXT_CONTENT_TYPES = {
    "mp4": "video/mp4", "m4v": "video/mp4", "mov": "video/quicktime",
    "wmv": "video/x-ms-wmv", "avi": "video/avi",
    "wav": "audio/wav", "mp3": "audio/mpeg", "m4a": "audio/mp4", "wma": "audio/x-ms-wma",
}


# ---------------------------------------------------------------------------
# zip helpers: every part we do not edit is copied byte-for-byte
# ---------------------------------------------------------------------------
def read_parts(path):
    with zipfile.ZipFile(path) as z:
        infos = [(i, z.read(i.filename)) for i in z.infolist()]
    return infos


def write_parts(infos, replacements, out_path):
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        for info, data in infos:
            payload = replacements.get(info.filename, data)
            zi = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = info.external_attr
            zi.internal_attr = info.internal_attr
            zi.create_system = info.create_system
            z.writestr(zi, payload)


# ---------------------------------------------------------------------------
# prefix-safe XML helpers (only attribute values / text are manipulated)
# ---------------------------------------------------------------------------
def xml_unescape(s):
    return (s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
             .replace("&apos;", "'").replace("&amp;", "&"))


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


def element_spans(xml, tag):
    """Character spans of <p:tag ...>...</p:tag> for non-nesting tags.

    The closing tag is "</p:" + tag + ">" == len(tag) + 5 characters. Getting that
    constant wrong is silent: slicing one character short leaves the trailing ">"
    of the closing tag behind, which is *malformed XML* -- lxml rejects it, and
    PowerPoint reports the package as corrupt (0x80070570). It stayed hidden while
    every caller's output happened to be re-parsed only after a later edit, so the
    span end is now derived from the tag itself rather than a magic number.
    """
    spans = []
    for m in re.finditer(r"<p:" + tag + r"(?=[\s/>])", xml):
        start = m.start()
        gt = xml.find(">", m.end())
        if gt == -1:
            continue
        if xml[gt - 1] == "/":
            spans.append((start, gt + 1))
            continue
        closing = "</p:%s>" % tag
        close = xml.find(closing, gt)
        if close == -1:
            continue
        spans.append((start, close + len(closing)))
    return spans


# ---------------------------------------------------------------------------
# slide child ordering: CT_Slide is a *sequence*, not a bag
#
# <p:sld> is <xsd:sequence>: (cSld, clrMapOvr?, transition?, timing?, extLst?).
# Getting this wrong is silent and expensive: PowerPoint loads a slide whose
# <p:transition> sits after <p:timing> without complaining, reports
# SlideShowTransition.EntryEffect == 0 (it never bound the element), and writes
# the transition OUT again on save. So a misplaced block costs you the whole
# transition with no error anywhere. Verified by round-tripping through
# PowerPoint: transition-before-timing survives, transition-after-timing is
# dropped. A survey of 274 slides from real decks (incl. this workspace) shows
# only cSld -> clrMapOvr -> transition -> timing (p:extLst only ever appears a
# child of p:cSld, never as a sibling).
#
# p:timing also deliberately sits AFTER p:transition: the old code appended the
# transition at the very end of <p:sld>, which both mis-ordered it and, whenever
# a timing block followed it, parked it *inside* the timing element.
# ---------------------------------------------------------------------------
SLIDE_CHILD_ORDER = {"cSld": 0, "clrMapOvr": 1, "transition": 2, "timing": 3, "extLst": 4}
SLIDE_ROOT_CHILDREN = ("cSld", "clrMapOvr", "transition", "timing", "extLst")


def root_child_spans(xml):
    """Spans of <p:sld>'s direct children only (depth 1, namespace-aware).

    Depth tracking rather than a substring search, because a descendant may
    contain any of these tags (a timing node can nest <p:childTnLst> freely) and
    a naive scan would report it as a sibling.
    """
    root = re.search(r"<p:sld\b[^>]*?(/?)>", xml)
    if not root:
        return []
    if root.group(1) == "/":          # <p:sld/>
        return []
    names = "|".join(SLIDE_ROOT_CHILDREN)
    out, depth, pos = [], 0, root.end()
    while pos < len(xml):
        nxt = xml.find("<", pos)
        if nxt == -1:
            break
        if xml.startswith("<!--", nxt):
            pos = xml.find("-->", nxt)
            if pos == -1:
                break
            pos += 3
            continue
        if xml.startswith("<![CDATA[", nxt):
            pos = xml.find("]]>", nxt)
            if pos == -1:
                break
            pos += 3
            continue
        if xml.startswith("<?", nxt):
            pos = xml.find("?>", nxt)
            if pos == -1:
                break
            pos += 2
            continue
        gt = xml.find(">", nxt)
        if gt == -1:
            break
        head = xml[nxt:gt + 1]
        if head.startswith("</"):
            depth -= 1
            pos = gt + 1
            if depth < 0:
                break                     # </p:sld> itself closes the root
            continue
        m = re.match(r"<(?:p:)?(%s)\b" % names, head)
        self_closing = head.endswith("/>")
        if depth == 0 and m:
            if self_closing:
                out.append((nxt, gt + 1, m.group(1)))
            else:
                close = xml.find("</p:%s>" % m.group(1), gt)
                if close == -1:
                    break
                out.append((nxt, close + len(m.group(1)) + 4, m.group(1)))
        if not self_closing:
            depth += 1
        pos = gt + 1
    return out


def insert_in_slide_order(xml, tag, new_block):
    """Insert a direct child of <p:sld> at its schema-correct position.

    Falls back to just before </p:sld> when nothing later in the sequence
    exists, which is the only position left for, e.g., a trailing <p:timing>
    once p:extLst is the sole follower.
    """
    rank = SLIDE_CHILD_ORDER[tag]
    for s, _e, found in root_child_spans(xml):
        if SLIDE_CHILD_ORDER[found] > rank:
            return xml[:s] + new_block + xml[s:]
    close = xml.rfind("</p:sld>")
    if close == -1:
        raise RuntimeError("anchor </p:sld> not found")
    return xml[:close] + new_block + xml[close:]


def replace_or_insert(xml, tag, new_block, anchor="</p:sld>"):
    """Replace the first <p:tag> in place, else insert it in schema order.

    Replacement is positional and therefore order-preserving: a transition that
    loaded correctly stays correctly placed. Insertion goes through
    insert_in_slide_order so a brand-new block never lands inside another
    element or after a later sibling in the CT_Slide sequence.
    """
    spans = element_spans(xml, tag)
    if spans:
        s, e = spans[0]
        return xml[:s] + new_block + xml[e:]
    if tag in SLIDE_CHILD_ORDER:
        return insert_in_slide_order(xml, tag, new_block)
    idx = xml.find(anchor)
    if idx == -1:
        raise RuntimeError("anchor %s not found" % anchor)
    return xml[:idx] + new_block + xml[idx:]


def remove_elements(xml, tag):
    for s, e in reversed(element_spans(xml, tag)):
        xml = xml[:s] + xml[e:]
    return xml


def sp_tree_block(slide_xml):
    spans = element_spans(slide_xml, "spTree")
    if not spans:
        raise RuntimeError("no <p:spTree> in slide")
    s, e = spans[0]
    return slide_xml[s:e]


# ---------------------------------------------------------------------------
# shape index: deck elementId -> shape id
# ---------------------------------------------------------------------------
def index_shapes(slide_xml):
    tree = sp_tree_block(slide_xml)
    out = {}
    order = 0
    for tag in SHAPE_TAGS:
        for s, e in element_spans(tree, tag):
            block = tree[s:e]
            m = re.search(r'<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"', block)
            if m:
                sid, name = m.group(1), m.group(2)
            else:
                m2 = re.search(r'<p:cNvPr\b[^>]*\bname="([^"]*)"[^>]*\bid="(\d+)"', block)
                if not m2:
                    continue
                name, sid = m2.group(1), m2.group(2)
            order += 1
            key = xml_unescape(name)
            out[key] = {"id": sid, "tag": tag, "order": order}
    for name in list(out):
        base = re.sub(r"-p\d+$", "", name)
        if base != name and base not in out:
            out[base] = dict(out[name], pieces=True)
    return out


def resolve_targets(target, shapes):
    if isinstance(target, (list, tuple)):
        out = []
        for t in target:
            out.extend(resolve_targets(t, shapes))
        return out
    key = str(target)
    if key in shapes:
        return [shapes[key]["id"]]
    if key.isdigit() and any(v["id"] == key for v in shapes.values()):
        return [key]
    pieces = [(n, v) for n, v in shapes.items() if re.fullmatch(re.escape(key) + r"-p\d+", n)]
    if pieces:
        return [v["id"] for _, v in sorted(pieces, key=lambda kv: kv[1]["order"])]
    return []


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
MORPH_TEMPLATE = (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" '
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">'
    '<p:transition spd="{spd}" p14:dur="{dur}"{extra}>'
    '<p159:morph option="{option}"/>'
    '</p:transition>'
    '</mc:Choice>'
    '<mc:Fallback><p:transition spd="{spd}"><p:fade/></p:transition></mc:Fallback>'
    '</mc:AlternateContent>'
)
MORPH_OPTIONS = ("byobject", "byword", "bychar")
MORPH_SPEEDS = {"slow": "slow", "med": "med", "fast": "fast"}


def build_transition(name, duration=None, advance_after=None, on_click=None,
                     option=None, speed=None):
    key = str(name or "none").strip().lower()

    if key == "morph":
        opt = str(option or "byObject").strip()
        if opt.lower() not in MORPH_OPTIONS:
            raise ValueError("morph option must be one of byObject/byWord/byChar, got %r" % option)
        opt = {"byobject": "byObject", "byword": "byWord", "bychar": "byChar"}[opt.lower()]
        ms = int(round(float(duration if duration is not None else 2.0) * 1000))
        spd = MORPH_SPEEDS.get(str(speed or "slow").lower(), "slow")
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
    # p14:dur is the millisecond-accurate duration; p:sld declares xmlns:p14
    ms = int(round(float(duration if duration is not None else 0.7) * 1000))
    block = block.format(dur=' p14:dur="%d"' % ms)
    extra = ""
    if advance_after:
        extra += ' advTm="%d"' % int(round(float(advance_after) * 1000))
    if on_click is False:
        extra += ' advClick="0"'
    if extra:
        block = block.replace("<p:transition", "<p:transition" + extra, 1)
    return block


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


# ---------------------------------------------------------------------------
# geometry fingerprint: the layout-frame proof
# ---------------------------------------------------------------------------
def geometry_fingerprint(pptx_path):
    out = {}
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted([n for n in z.namelist() if SLIDE_RE.match(n)],
                       key=lambda n: int(SLIDE_RE.match(n).group(1)))
        for name in names:
            xml = z.read(name).decode("utf-8", "replace")
            tree = sp_tree_block(xml)
            rows = []
            for tag in SHAPE_TAGS:
                for s, e in element_spans(tree, tag):
                    block = tree[s:e]
                    cm = re.search(r'<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"', block)
                    ident = list(cm.groups()) if cm else ["", ""]
                    xf = re.search(r"<a:xfrm[^>]*>.*?</a:xfrm>", block, re.S)
                    geom = []
                    if xf:
                        for attr in GEOM_ATTRS:
                            am = re.search(r'\b%s="(-?\d+)"' % attr, xf.group(0))
                            geom.append(am.group(1) if am else "-")
                    rows.append([tag] + ident + ["|".join(geom)])
            payload = json.dumps(rows, ensure_ascii=False)
            out[name] = {
                "sha256": hashlib.sha256(payload.encode("utf-8")).hexdigest(),
                "shapes": len(rows),
            }
    return out


# ---------------------------------------------------------------------------
# spec handling
# ---------------------------------------------------------------------------
def load_spec(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    if path.lower().endswith((".yaml", ".yml")):
        if yaml is None:
            raise RuntimeError("PyYAML required for .yaml specs (pip install pyyaml)")
        return yaml.safe_load(text)
    return json.loads(text)


def normalize_effects(slide_entry):
    raw = slide_entry.get("effects")
    if raw is None:
        raw = slide_entry.get("animation") or []
    if isinstance(raw, dict):
        out = []
        for k, v in raw.items():
            item = dict(v) if isinstance(v, dict) else {"effect": v}
            item.setdefault("target", k)
            out.append(item)
        return out
    if not isinstance(raw, list):
        raise ValueError("effects must be a list or mapping")
    return [dict(x) for x in raw]


def catalog_path():
    return CATALOG_PATH


def normalize_spec(spec):
    """[{page, effects, transition}] with page resolved from page/index/slide.

    apply_motion() does the same resolution inline; player.py needs it too, so the
    rule lives here once instead of being reimplemented.
    """
    out = []
    for entry in (spec.get("slides") or []):
        if "page" in entry:
            page = int(entry["page"])
        elif "index" in entry:
            page = int(entry["index"])
        elif "slide" in entry:
            v = entry["slide"]
            if not isinstance(v, int):
                raise ValueError("slide: must be numeric to resolve a page number")
            page = v
        else:
            raise ValueError("slide entry needs page/index/slide: %r" % (entry,))
        out.append({"page": page,
                    "effects": normalize_effects(entry),
                    "transition": entry.get("transition")})
    return out


def schedule_spec(effects):
    """Absolute start time per effect, in seconds.

    Mirrors build_timing(): each effect becomes its own <p:par>, so `after` waits
    for its whole PRECEDING GROUP to finish, while `with` shares the current
    group's start (which is why a spec puts `with` entries directly after the
    effect they accompany).

    The group's end is the max end of its members, not the last member's end: a
    `with` effect can be longer than the one it accompanies, and the following
    `after` must wait for the slower of them.
    """
    out, group_start, group_end = [], 0.0, 0.0
    for eff in effects:
        dur = float(eff.get("duration") or 0.6)
        delay = float(eff.get("delay") or 0.0)
        if str(eff.get("trigger", "after")).lower() == "with" and out:
            start = group_start
        else:
            start = group_end + delay
            group_start = start
            group_end = start
        end = start + dur
        if end > group_end:
            group_end = end
        out.append({"target": eff.get("target"), "effect": eff.get("effect"),
                    "start": round(start, 3), "duration": round(dur, 3),
                    "dir": str(eff["dir"]).strip().lower() if eff.get("dir") else None})
    return out


def slide_size(pptx_path, default=(960, 540)):
    """(cx, cy) of the slide in points, from <p:sldSz> in presentation.xml.

    sldSz is in EMU (1 pt = 12700 EMU). Falls back to the default when the part or
    the attribute is missing rather than failing the whole run.
    """
    try:
        with zipfile.ZipFile(pptx_path) as z:
            xml = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    except (KeyError, OSError, zipfile.BadZipFile):
        return default
    m = re.search(r'<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', xml)
    if not m:
        m = re.search(r'<p:sldSz\b[^>]*\bcy="(\d+)"[^>]*\bcx="(\d+)"', xml)
        if m:
            return (int(m.group(2)) / 12700.0, int(m.group(1)) / 12700.0)
        return default
    return (int(m.group(1)) / 12700.0, int(m.group(2)) / 12700.0)


# ---------------------------------------------------------------------------
# apply
# ---------------------------------------------------------------------------
def apply_motion(pptx_in, spec, out_path, assert_geometry=False):
    with open(CATALOG_PATH, encoding="utf-8") as fh:
        catalog = json.load(fh)

    infos = read_parts(pptx_in)
    by_name = {i.filename: d for i, d in infos}
    slide_names = sorted([n for n in by_name if SLIDE_RE.match(n)],
                         key=lambda n: int(SLIDE_RE.match(n).group(1)))

    report = {
        "source": os.path.abspath(pptx_in),
        "out": os.path.abspath(out_path),
        "slides": [], "warnings": [], "errors": [],
        "media_plan": [], "effects_total": 0, "transitioned": 0,
    }

    spec_by_index, spec_by_name = {}, {}
    for entry in (spec.get("slides") or []):
        if "page" in entry:
            spec_by_index[int(entry["page"])] = entry
        elif "index" in entry:
            spec_by_index[int(entry["index"])] = entry
        elif "slide" in entry:
            v = entry["slide"]
            if isinstance(v, int):
                spec_by_index[v] = entry
            else:
                spec_by_name[str(v)] = entry
        else:
            report["errors"].append("slide entry needs page/index/slide: %r" % (entry,))

    replacements = {}
    for pos, sname in enumerate(slide_names, start=1):
        idx = int(SLIDE_RE.match(sname).group(1))
        xml = by_name[sname].decode("utf-8")
        entry = spec_by_index.get(idx) or spec_by_index.get(pos) or spec_by_name.get(sname)
        shapes = index_shapes(xml)
        sr = {"slide": idx, "part": sname, "effects": 0, "unknownTargets": []}

        if entry is not None:
            effects = normalize_effects(entry)
            entries = []
            for eff in effects:
                alias = eff.get("effect") or eff.get("name")
                if not alias:
                    report["errors"].append("slide %d: effect entry without 'effect'" % idx)
                    continue
                meta = catalog.get(str(alias))
                if meta is None:
                    report["errors"].append("slide %d: unknown effect alias %r" % (idx, alias))
                    continue
                ids = resolve_targets(eff.get("target"), shapes)
                if not ids:
                    sr["unknownTargets"].append(str(eff.get("target")))
                    report["errors"].append(
                        "slide %d: target %r not found in spTree" % (idx, eff.get("target")))
                    continue
                for sid in ids:
                    entries.append({
                        "shape_id": sid, "tpl": meta["template"],
                        "effect": eff, "delay": eff.get("delay") or 0.0,
                    })
                sr.setdefault("applied", []).append({
                    "effect": alias, "presetID": meta["presetID"],
                    "presetClass": meta["presetClass"], "target": eff.get("target"), "ids": ids,
                })
                sr["effects"] += len(ids)

            if entries:
                new_xml = replace_or_insert(xml, "timing", build_timing(entries, NodeIds(4)))
            else:
                new_xml = remove_elements(xml, "timing")

            trans = entry.get("transition")
            if trans is not None:
                if isinstance(trans, str):
                    trans = {"type": trans}
                is_morph = str(trans.get("type", "")).strip().lower() == "morph"
                block = build_transition(trans.get("type", "fade"), trans.get("duration"),
                                         trans.get("advanceAfter"), trans.get("onClick"),
                                         trans.get("option"), trans.get("speed"))
                if block and "p14:" in block and not is_morph:
                    # p14:dur needs the prefix bound on the slide ROOT. Declare it
                    # when we can, else fall back to the pre-2010 transition form
                    # (spd only, no millisecond duration) which every version opens.
                    if not root_declares(new_xml, "p14"):
                        new_xml = add_root_namespace(
                            new_xml, "p14",
                            "http://schemas.microsoft.com/office/powerpoint/2010/main")
                    if not root_declares(new_xml, "p14"):
                        block = re.sub(r'\s+p14:dur="[^"]*"', "", block)
                        report["warnings"].append(
                            "slide %d: xmlns:p14 not declarable on root, "
                            "transition duration omitted" % idx)
                new_xml = remove_elements(new_xml, "transition")
                if block:
                    if is_morph:
                        # morph is wrapped in mc:AlternateContent. Drop the whole
                        # wrapper first, then insert with the "transition" tag so
                        # SLIDE_CHILD_ORDER still places it before <p:timing>.
                        new_xml = drop_alternate_content(new_xml)
                    new_xml = replace_or_insert(new_xml, "transition", block)
                sr["transition"] = trans.get("type", "fade")
                report["transitioned"] += 1

            # ---- 3D cameras ------------------------------------------------
            # Not an animation: <a:scene3d> is a shape property, so it never
            # enters the timing tree. Applied separately from `effects`.
            for cam in (entry.get("cameras") or []):
                if not isinstance(cam, dict):
                    cam = {"target": cam}
                ids = resolve_targets(cam.get("target"), shapes)
                if not ids:
                    sr["unknownTargets"].append(str(cam.get("target")))
                    report["errors"].append(
                        "slide %d: camera target %r not found in spTree"
                        % (idx, cam.get("target")))
                    continue
                try:
                    frag = build_camera(cam)
                except ValueError as exc:
                    report["errors"].append("slide %d: %s" % (idx, exc))
                    continue
                prst = cam.get("prst") or CAMERA_DEFAULT_PRESET
                if not camera_is_perspective(prst):
                    report["warnings"].append(
                        "slide %d: camera prst=%s is a PARALLEL projection -- it "
                        "will never show a vanishing point" % (idx, prst))
                for sid in ids:
                    new_xml, status = insert_scene3d(new_xml, sid, frag)
                    sr.setdefault("cameras", []).append({
                        "target": cam.get("target"), "id": sid,
                        "prst": prst, "status": status,
                        "tilt_deg": cam.get("lat", cam.get("tilt", 0)),
                        "lon_deg": cam.get("lon", 0),
                    })
                    if status != "ok":
                        report["warnings"].append(
                            "slide %d: camera on id=%s -> %s" % (idx, sid, status))
                sr["cameras_applied"] = sr.get("cameras_applied", 0) + len(ids)

            for media in (entry.get("media") or []):
                report["media_plan"].append({
                    "slide": idx,
                    "src": media.get("src"),
                    "kind": media.get("kind", "auto"),
                    "bounds": media.get("bounds"),
                    "link": bool(media.get("link", False)),
                    "loop": bool(media.get("loop", False)),
                    "rewind": bool(media.get("rewind", True)),
                    "mute": bool(media.get("mute", False)),
                    "volume": media.get("volume"),
                    "autoplay": bool(media.get("autoplay", True)),
                    "elementId": media.get("elementId"),
                })

            if new_xml != xml:
                replacements[sname] = new_xml.encode("utf-8")

        report["effects_total"] += sr["effects"]
        report["slides"].append(sr)

    ct_name = "[Content_Types].xml"
    if ct_name in by_name:
        ct_xml = by_name[ct_name].decode("utf-8")
        original = ct_xml
        for m in report["media_plan"]:
            if m["src"]:
                ct_xml = ensure_content_type(ct_xml, os.path.splitext(m["src"])[1].lstrip("."))
        if ct_xml != original:
            replacements[ct_name] = ct_xml.encode("utf-8")

    # Final gate: every part we rewrote must still be well-formed XML. A malformed
    # part opens fine in no viewer but produces a baffling PowerPoint E_FAIL, so
    # never write one.
    if _ET is not None:
        for name in sorted(replacements):
            payload = replacements[name]
            if not name.lower().endswith(".xml"):
                continue
            try:
                _ET.fromstring(payload)
            except Exception as exc:
                report["errors"].append("refusing to write malformed XML in %s: %s" % (name, exc))
        if report["errors"]:
            return report

    write_parts(infos, replacements, out_path)

    if assert_geometry:
        before = geometry_fingerprint(pptx_in)
        after = geometry_fingerprint(out_path)
        diffs = [k for k in sorted(set(before) | set(after)) if before.get(k) != after.get(k)]
        report["geometry"] = {"ok": not diffs, "diff": diffs, "slides": len(before)}
        if diffs:
            report["errors"].append("geometry changed by motion injection: %s" % diffs)
    return report


MC = MC_NS


def transitions_are_paired(xml):
    """True when every <p:transition> sits inside an mc:AlternateContent wrapper.

    A deck may hold more <p:transition> elements than effective transitions only
    in that one sanctioned way (Choice + Fallback). Anything else is a real
    duplicate that PowerPoint may resolve unpredictably.
    """
    if _ET is None:
        return False
    try:
        root = _ET.fromstring(xml.encode("utf-8"))
    except Exception:
        return False
    total = sum(1 for _ in root.iter(P_NS + "transition"))
    wrapped = 0
    for alt in root.iter(MC + "AlternateContent"):
        wrapped += sum(1 for _ in alt.iter(P_NS + "transition"))
    return total == wrapped and total > 0


def active_transition_blocks(xml):
    """Spans of the *effective* <p:transition> elements in a slide.

    PowerPoint serialises a transition as an mc:AlternateContent pair: the
    p14:dur flavour under <mc:Choice> and a pre-2010 fallback under
    <mc:Fallback>. Both must be written -- that is how PowerPoint itself does it
    -- but only the Choice is the transition in force, so counting raw
    occurrences over-reports by one on every PowerPoint-saved deck. Fall back to
    every occurrence when there is no AlternateContent wrapper (our own output,
    or a pre-2010 deck).
    """
    spans = element_spans(xml, "transition")
    if _ET is None:
        return spans
    try:
        root = _ET.fromstring(xml.encode("utf-8"))
    except Exception:
        return spans
    effective = []
    for alt in root.iter(MC + "AlternateContent"):
        choice = alt.find(MC + "Choice")
        if choice is None:
            continue
        for tr in choice.iter(P_NS + "transition"):
            effective.append(tr)
    if not effective:
        return spans
    return spans[:len(effective)]


def strip_animations(pptx_in, out_path):
    """Remove both <p:timing> and <p:transition> so the deck renders statically.

    Entrance animations keep their shapes hidden until they start, so a PNG/PDF
    grabbed from an animated deck is missing every "fly in" shape. Transitions
    have to go too: leaving them means the "static" copy still has slide
    transitions, which is not what the preview is for. Emptying an
    mc:AlternateContent wrapper is cleaned up as well -- an AlternateContent
    with no Choice and no Fallback is schema-invalid, and PowerPoint is entitled
    to reject it.
    """
    infos = read_parts(pptx_in)
    replacements = {}
    for info, data in infos:
        if not SLIDE_RE.match(info.filename):
            continue
        xml = data.decode("utf-8", "replace")
        new = remove_elements(xml, "timing")
        new = remove_elements(new, "transition")
        new = remove_empty_alternate_content(new)
        if new != xml:
            replacements[info.filename] = new.encode("utf-8")
    write_parts(infos, replacements, out_path)
    return {"out": os.path.abspath(out_path), "slides_stripped": len(replacements)}


def remove_empty_alternate_content(xml):
    """Drop mc:AlternateContent blocks that no longer hold a Choice or Fallback."""
    out = xml
    while True:
        nxt = None
        for m in re.finditer(r"<mc:AlternateContent(?=[\s/>])", out):
            s = m.start()
            gt = out.find(">", m.end())
            if gt == -1:
                break
            if out[gt - 1] == "/":
                nxt = (s, gt + 1)
                break
            close = out.find("</mc:AlternateContent>", gt)
            if close == -1:
                break
            e = close + len("</mc:AlternateContent>")
            body = out[gt + 1:close]
            if "<mc:Choice" not in body and "<mc:Fallback" not in body:
                nxt = (s, e)
                break
        if nxt is None:
            return out
        out = out[:nxt[0]] + out[nxt[1]:]


def inspect(pptx_path):
    out = {"file": os.path.abspath(pptx_path), "slides": []}
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted([n for n in z.namelist() if SLIDE_RE.match(n)],
                       key=lambda n: int(SLIDE_RE.match(n).group(1)))
        for name in names:
            xml = z.read(name).decode("utf-8", "replace")
            rels_name = "ppt/slides/_rels/%s.rels" % os.path.basename(name)
            try:
                rels = z.read(rels_name).decode("utf-8", "replace")
            except KeyError:
                rels = ""
            shapes = index_shapes(xml)
            presets = re.findall(r'presetID="(\d+)"[^>]*presetClass="(\w+)"', xml)
            out["slides"].append({
                "index": int(SLIDE_RE.match(name).group(1)),
                "shapes": [{"name": k, "id": v["id"], "tag": v["tag"]}
                           for k, v in shapes.items() if not v.get("pieces")],
                "timing": bool(re.search(r"<p:timing>", xml)),
                "effects": [{"presetID": int(a), "presetClass": b} for a, b in presets],
                "motionPaths": len(re.findall(r"<p:animMotion\b", xml)),
                "transition": bool(re.search(r"<p:transition\b", xml)),
                "media": sorted(set(re.findall(r'relationships/(video|audio|media)"', rels))),
            })
    return out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="OOXML motion injection for .pptx")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("apply")
    a.add_argument("--pptx", required=True)
    a.add_argument("--spec", required=True)
    a.add_argument("--out", required=True)
    a.add_argument("--assert-geometry", action="store_true")
    a.add_argument("--report")
    a.add_argument("--json", action="store_true")

    i = sub.add_parser("inspect")
    i.add_argument("--pptx", required=True)
    i.add_argument("--json", action="store_true")

    p = sub.add_parser("preview")
    p.add_argument("--pptx", required=True)
    p.add_argument("--out", required=True)

    c = sub.add_parser("catalog")
    c.add_argument("--kind", choices=["entrance", "emphasis", "path"])
    c.add_argument("--json", action="store_true")

    pl = sub.add_parser("player", help="build an HTML motion preview from a spec")
    pl.add_argument("--pptx", required=True)
    pl.add_argument("--spec", required=True)
    pl.add_argument("--outdir", required=True,
                    help="writes <outdir>.html next to <outdir>/stage/*.png")
    pl.add_argument("--title")
    pl.add_argument("--json", action="store_true")

    ck = sub.add_parser("check", help="audit spec coverage (shapes with no effect)")
    ck.add_argument("--pptx", required=True)
    ck.add_argument("--spec", required=True)
    ck.add_argument("--footer", action="append",
                    help="extra static-footer text marker (repeatable)")
    ck.add_argument("--json", action="store_true")

    ns = ap.parse_args(argv)

    if ns.cmd == "apply":
        report = apply_motion(ns.pptx, load_spec(ns.spec), ns.out,
                             assert_geometry=ns.assert_geometry)
        if ns.report:
            with open(ns.report, "w", encoding="utf-8") as fh:
                json.dump(report, fh, ensure_ascii=False, indent=1)
        if ns.json:
            print(json.dumps(report, ensure_ascii=False, indent=1))
        else:
            print("slides=%d effects=%d transitions=%d media=%d" % (
                len(report["slides"]), report["effects_total"],
                report["transitioned"], len(report["media_plan"])))
            if report.get("geometry"):
                g = report["geometry"]
                print("geometry: %s over %d slides" % (
                    "UNCHANGED" if g["ok"] else "CHANGED", g["slides"]))
            for w in report["warnings"]:
                print("WARN:", w)
            for e in report["errors"]:
                print("ERROR:", e)
        return 1 if report["errors"] else 0

    if ns.cmd == "inspect":
        data = inspect(ns.pptx)
        if ns.json:
            print(json.dumps(data, ensure_ascii=False, indent=1))
        else:
            for sl in data["slides"]:
                print("slide %d: shapes=%d effects=%d paths=%d transition=%s media=%s" % (
                    sl["index"], len(sl["shapes"]), len(sl["effects"]),
                    sl["motionPaths"], sl["transition"], sl["media"] or "-"))
                for sh in sl["shapes"]:
                    print("    id=%-6s %-14s %s" % (sh["id"], sh["tag"], sh["name"]))
        return 0

    if ns.cmd == "preview":
        print(json.dumps(strip_animations(ns.pptx, ns.out), ensure_ascii=False))
        return 0

    if ns.cmd == "catalog":
        with open(CATALOG_PATH, encoding="utf-8") as fh:
            catalog = json.load(fh)
        rows = [(k, v) for k, v in catalog.items() if not ns.kind or v["kind"] == ns.kind]
        if ns.json:
            print(json.dumps({k: {kk: vv for kk, vv in v.items() if kk != "template"}
                              for k, v in rows}, ensure_ascii=False, indent=1))
        else:
            for k, v in sorted(rows, key=lambda x: (x[1]["kind"], x[1]["presetID"])):
                flags = "".join(["P" if v["hasMotionPath"] else "-",
                                 "F" if v["hasFilter"] else "-",
                                 "V" if v["hasVisibilitySet"] else "-"])
                # The MsoAnimEffect names are printed because enum != presetID:
                # `spin` is presetID 8 but enum 61, so anyone reading a Microsoft
                # reference needs the name to find the effect they mean.
                print("%-24s %-9s presetID=%-4d sub=%-3d enum=%-4s %-26s %s" % (
                    k, v["kind"], v["presetID"], v["presetSubtype"],
                    v.get("enum", ""), v.get("enumName", ""), flags))
            print("\ntotal %d aliases (P=motion path, F=filter, V=visibility set)"
                  % len(rows))
            print("enum = MsoAnimEffect value; enumName = its constant. "
                  "enum != presetID by design.")
        return 0

    if ns.cmd == "player":
        # Imported lazily: player.py imports this module, and pywin32 is only
        # needed for this one command.
        import player
        try:
            res = player.build_player(ns.pptx, ns.spec, ns.outdir, title=ns.title)
        except Exception as exc:
            print("FAILED: %s" % exc)
            return 1
        if ns.json:
            print(json.dumps(res, ensure_ascii=False, indent=1))
        print("open: %s" % res["html"])
        return 0

    if ns.cmd == "check":
        import check_coverage
        argv = ["--pptx", ns.pptx, "--spec", ns.spec]
        for f in (ns.footer or []):
            argv += ["--footer", f]
        if ns.json:
            argv.append("--json")
        return check_coverage.main(argv)
    return 2


if __name__ == "__main__":
    sys.exit(main())
