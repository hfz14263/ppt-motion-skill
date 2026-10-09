"""motion.xml — 常量与 XML 底层 —— **整个引擎的地基**。

常量与 XML 底层 —— **整个引擎的地基**。

    · zip 读写：不编辑的 part 逐字节复制（改坏文件的唯一方式就是解压再压缩）
    · 前缀安全的字符串手术：只动属性值与文本，不重排 XML
    · 子元素顺序：`<p:sld>` 是 **sequence 不是 bag**，位置错了元素会被静默丢弃
    · 单例：**本项目所有损坏文件都是「同一元素出现两次」**
    · 形状索引：deck 的 elementId → shape id

模块内容表（内容表；改这一层前先读这里）
----------------------------------------------------------------
    常量        命名空间 / 正则 / 效果表 / 切换表 / 媒体类型
    zip         read_parts / write_parts
    XML 手术     xml_escape / xml_unescape / element_spans
    子元素顺序   insert_in_slide_order / replace_or_insert / remove_elements
    单例         set_singleton / singleton_spans / find_duplicate_singletons
    形状索引     index_shapes / resolve_targets / sp_tree_block

> 本模块由 `tools/split_motion.py` 从 `motion.py`（1914 行）按**分层**拆出
> （原文照搬，未改逻辑）。**外部接口由 `../motion.py` 重新导出** ——
> 10 个模块 `import motion` 用的名字一个都没变。
"""

import json
import os
import re
import zipfile


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


# ---------------------------------------------------------------------------
# Slide transitions come from transition_reference.json, which was MEASURED on
# this PowerPoint build by scripts/build_transition_table.py (see that file's
# header). These used to be thirteen hand-written template strings, and every
# one of them was wrong in some way that only a roundtrip could show:
#
#   * none of them was wrapped in mc:AlternateContent. PowerPoint keeps a bare
#     <p:transition> but rewrites it on save, so the file we produced was never
#     in the form PowerPoint itself writes.
#   * the Fallback is not <p:fade/> for a core element -- it repeats the core
#     element. Only a p14/p15/p159 child degrades to fade.
#   * pushleft / wipe / cover / split / pull spelled out direction attributes
#     that carry the DEFAULT value; PowerPoint deletes them on save.
#   * TRANSITION_ENUM below was a table of legacy 2003-era numbers in which
#     "fade" actually resolves to <p:strips/>. The measured values are nothing
#     like it (fade is 3849, strips is 2561).
#
# The table stores PowerPoint's own saved XML with {spd} and {dur} parameterised
# back out, so the block emitted here is structurally identical to what
# PowerPoint writes, and every namespace it needs is declared inside the block
# itself -- no root-namespace fixup, no "duration omitted" fallback.
# ---------------------------------------------------------------------------
TRANSITION_REF_PATH = os.path.join(HERE, "transition_reference.json")


def _load_transition_reference():
    with open(TRANSITION_REF_PATH, encoding="utf-8-sig") as fh:
        return json.load(fh)


_TRANSITION_REF = _load_transition_reference()


TRANSITIONS = {spec: e["xml"] for spec, e in _TRANSITION_REF["by_spec"].items()}
TRANSITIONS["none"] = ""


# Measured SlideShowTransition.EntryEffect per spec, straight out of the enum
# scan. Used only so the review layer can report a non-zero transition through
# the COM property model; the XML above is what actually lands in the file.
TRANSITION_ENUM = {spec: e["entryEffect"]
                   for spec, e in _TRANSITION_REF["by_spec"].items()
                   if e["entryEffect"] is not None}
TRANSITION_ENUM["none"] = 0


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


# ---------------------------------------------------------------------------
# singletons: elements that may appear at most once inside their parent
# ---------------------------------------------------------------------------
# Every corruption in this project's history has the same shape: an element was
# INSERTED into a position that already held it. OOXML schema types are largely
# xsd:sequence with optional members, so a second copy is not "extra" -- it makes the
# document invalid and PowerPoint refuses to open the file at all (0x80070570, "file or
# directory is corrupt"), with no hint about which element is at fault.
#
# Recorded instances, all found only by opening the file in PowerPoint:
#   * <p:transition> twice, because python-pptx had written one
#   * <a:solidFill> beside <a:blipFill>, and fill is an xsd:choice so the FIRST wins
#   * <a:effectLst/> (empty, self-closing) plus a second one carrying the shadow
#   * a:rot lat/lon out of range -- a sibling problem, same silent-then-fatal pattern
#
# So the rule is: never `s.replace("</p:spPr>", X + "</p:spPr>")`. Use set_singleton,
# which replaces when the element exists and inserts only when it does not. The
# helpers below exist because ad-hoc insertion kept looking correct in the XML and
# failing in PowerPoint.
def set_singleton(xml, tag, new_block, inside=None, before=()):
    """Make `tag` appear exactly once, replacing any existing instance.

    Returns (xml, status) with status in {replaced, inserted, no-<parent>}.

    Insert-or-replace, never plain insert: appending to a position that already holds
    the element makes the document invalid and PowerPoint refuses to open the file,
    reporting only "file or directory is corrupt". Every corrupted file in this
    project's history came from that one mistake.

    `inside` scopes the operation to one parent element's span, which matters because a
    tag such as a:effectLst also occurs in contexts that must not be touched.

    `before` lists the tags this one must precede when it has to be INSERTED. It is only
    consulted then; an existing element is replaced where it already sits, so its
    position is never disturbed.
    """
    if inside is not None:
        spans = element_spans(xml, inside)
        if not spans:
            return xml, "no-%s" % inside
        s, e = spans[0]
        body, status = set_singleton(xml[s:e], tag, new_block, inside=None,
                                     before=before)
        return xml[:s] + body + xml[e:], status

    spans = singleton_spans(xml, tag)
    if spans:
        s, e = spans[0]
        out = xml[:s] + new_block + xml[e:]
        # drop any FURTHER copies: leaving one behind is the entire bug. Removed
        # back-to-front so the earlier offsets stay valid.
        for s2, e2 in reversed(spans[1:]):
            out = out[:s2] + out[e2:]
        return out, "replaced"

    # absent: place it before the first tag it must precede, else before the closing tag
    # of THIS element. Falling back to the last "</" in the whole string put the element
    # at the wrong nesting level, which is silently invalid.
    at = None
    for later in before:
        lm = re.search(r"<%s(?=[\s/>])" % re.escape(later), xml)
        if lm and (at is None or lm.start() < at):
            at = lm.start()
    if at is None:
        head = re.match(r"<[A-Za-z0-9:]+(?=[\s/>])[^>]*>", xml)
        if not head:
            return xml, "no-anchor"
        close = xml.rfind("</", head.end())
        at = close if close > 0 else len(xml)
    return xml[:at] + new_block + xml[at:], "inserted"


def singleton_spans(xml, tag):
    """Spans of a tag in BOTH forms: `<x/>` and `<x>...</x>`.

    element_spans only reports paired forms, so a self-closing `<a:effectLst/>` was
    invisible to it. That mattered: set_singleton replaced the self-closing copy through
    its fallback branch and never saw the paired duplicate sitting beside it, so the
    result still carried two effectLst and PowerPoint still refused the file. Both forms
    have to be found in one pass or the duplicate survives.
    """
    pat = re.compile(r"<%s(?=[\s/>])[^>]*?(/?)>" % re.escape(tag))
    out = []
    for m in pat.finditer(xml):
        if m.group(1) == "/":
            out.append((m.start(), m.end()))
            continue
        close = xml.find("</%s>" % tag, m.end())
        if close > 0:
            out.append((m.start(), close + len(tag) + 3))
    return out


def direct_children(body):
    """Top-level child tag names of one element, ignoring everything nested."""
    out = []
    scan = re.compile(r"<(/?)([A-Za-z0-9:]+)([^>]*?)(/?)>")
    i = body.index(">") + 1
    end = body.rindex("<")
    depth = 0
    while i < end:
        m = scan.search(body, i)
        if not m or m.start() >= end:
            break
        closing, t, _a, selfclose = m.groups()
        if closing:
            depth -= 1
        elif depth == 0:
            out.append(t)
            if not selfclose:
                depth += 1
        else:
            if not selfclose:
                depth += 1
        i = m.end()
    return out


# Fill elements are an xsd:choice, not a repeated element: two of them coexisting is
# invalid even when they have DIFFERENT names, and the first one silently wins. A
# same-name duplicate check cannot see that, which is how <a:solidFill/> beside
# <a:blipFill/> survived a "clean" report.
FILL_GROUP = ("a:noFill", "a:solidFill", "a:gradFill", "a:blipFill", "a:pattFill",
              "a:grpFill")


def find_duplicate_singletons(xml, parent="spPr", tags=None):
    """Report sequence members that appear more than once, and fill conflicts.

    A structural check for the class of bug that produced every corrupted file here.

    Three mistakes were made getting this right, all recorded because a validator that
    cannot fail is worse than none:
      * depth was tracked wrong and the scan checked `depth == 1` for a direct child.
        Inside spPr the direct children sit at depth 0, so it matched nothing and
        reported OK on a file PowerPoint rejected.
      * counting `<p:transition` with a global regex flagged a CORRECT morph as a
        duplicate, because morph carries two transitions by design -- one in mc:Choice,
        one in mc:Fallback. Only direct children of p:sld count.
      * only same-name duplicates were looked for, so two DIFFERENT fills passed. Fills
        are a choice, so any two of them conflicting is the fault.
    """
    if tags is None:
        tags = ("a:effectLst", "a:effectDag", "a:ln", "a:blipFill", "a:solidFill",
                "a:noFill", "a:gradFill", "a:pattFill", "a:grpFill", "a:scene3d",
                "a:sp3d", "a:xfrm")
    out = []
    for s, e in element_spans(xml, parent):
        kids = direct_children(xml[s:e])
        seen = {}
        for t in kids:
            if t in tags:
                seen[t] = seen.get(t, 0) + 1
        for t, n in seen.items():
            if n > 1:
                out.append((t, n))
        fills = [t for t in kids if t in FILL_GROUP]
        if len(fills) > 1:
            for t in dict.fromkeys(fills):
                out.append((t, fills.count(t)))
    return out


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
