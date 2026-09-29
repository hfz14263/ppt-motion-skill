#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the slide-transition reference table: UI gallery item -> MEASURED XML.

WHY A MEASURED TABLE, AND WHY THE DOCS ARE NOT ENOUGH
-----------------------------------------------------
The gallery is what the user sees; the file only has <p:transition>. Four things
separate them, and none of the four can be looked up:

  1. The COM enum lies. SlideShowTransition.EntryEffect is the 2003-era enum and
     this build maps it far from the obvious numbers (0x0A01 "fade" emits
     <p:strips/>); several XML elements read back the SAME enum value. So the
     enum can never establish a UI-name -> element mapping.  (com-pitfalls §7)
  2. The gallery is NOT one namespace. Some 细微 items are core ECMA elements,
     one is p159 (morph), 闪光 is p14 -- and the whole 华丽 group is p15
     prstTrans. UI grouping ≠ XML namespace; the only way to know is to write a
     candidate and see what PowerPoint does with it.
  3. Defaults are unwritten. ECMA says <p:push/> takes dir; it does not say what
     PowerPoint writes when the author never touched 效果选项. Omitting the
     attribute and reading back the saved file is the only honest way to learn
     the default.
  4. A bad candidate does not fail quietly -- it fails LOUDLY. On this build a
     single schema-invalid <p:transition> child makes PowerPoint refuse the
     WHOLE file (com-pitfalls §26). So candidates are probed ONE DECK EACH:
     "PowerPoint would not open it" is a first-class verdict, not an exception
     to swallow. Mixing 48 candidates into one deck was the first thing this
     script tried and it destroyed exactly this signal.

So: HYPOTHESES below is a hypothesis list (from the gallery screenshot + the
ECMA-376 core element list + [MS-PPTX] p14 / p15 extension names). The
roundtrip through PowerPoint turns it into a table. Anything PowerPoint
rewrites or refuses is recorded as such, never silently fixed.

Subcommands:
    build   CASES_DIR                 one 2-slide deck per candidate + manifest
    collect CASES_DIR OUT_DIR         (after probe_transitions.ps1) diff all
    video   CASES_DIR MANIFEST OUT    one combined deck for the rendering pass
    sheets  VIDEO OUT_DIR             contact sheets for the visual pass
"""
import argparse
import glob
import io
import json
import os
import re
import shutil
import sys
import zipfile

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "p15": "http://schemas.microsoft.com/office/powerpoint/2012/main",
    "p159": "http://schemas.microsoft.com/office/powerpoint/2015/09/main",
}
NSDECL = " ".join('xmlns:%s="%s"' % kv for kv in sorted(NS.items()))

# --------------------------------------------------------------------------
# The 48 hypotheses. `family` decides how the child is wrapped:
#   core  -> direct child of <p:transition>            (ECMA-376)
#   p14   -> mc:Choice Requires="p14" + fade Fallback  ([MS-PPTX] 3.1)
#   p15   -> mc:Choice Requires="p15" + fade Fallback
#   p159  -> morph, already proven by motion.py
# Attributes are written ONLY where they distinguish two gallery items
# (cube / rotate / orbit are all p14:prism) or where PowerPoint refuses the
# bare form -- 切换/翻转/库/摩天轮/传送带 are refused WITHOUT dir, which the
# roundtrip found and the enum scan explained: PowerPoint itself always writes
# them with dir="l". "Refused" alone would never have said which attribute.
# Everything else is left out ON PURPOSE so the roundtrip reveals PowerPoint's
# own default.
# --------------------------------------------------------------------------
def _with_dir(child, direction):
    """Return `child` with a `dir="..."` attribute injected.

    `dir` goes on the CHILD element (`<p:push dir="l"/>`), not on
    <p:transition> -- that is where PowerPoint itself writes it
    (com-pitfalls §41/§42). Only used by the direction probe; the default
    probe deliberately writes no dir at all.
    """
    if not direction:
        return child
    m = re.match(r"<([A-Za-z0-9:]+)((?:\s[^>]*)?)/>$", child)
    if not m:
        raise ValueError("cannot inject dir into %r" % child)
    tag, attrs = m.group(1), m.group(2) or ""
    attrs = re.sub(r'\s+dir="[^"]*"', "", attrs)      # replace, don't stack
    return '<%s%s dir="%s"/>' % (tag, attrs, direction)


HYPOTHESES = [
    # ---- 细微 (12, excluding 无) -----------------------------------------
    {"ui": "淡入/淡出", "group": "细微", "en": "Fade", "family": "core",
     "spec": "fade", "child": "<p:fade/>"},
    {"ui": "推入", "group": "细微", "en": "Push", "family": "core",
     "spec": "push", "child": "<p:push/>"},
    {"ui": "擦除", "group": "细微", "en": "Wipe", "family": "core",
     "spec": "wipe", "child": "<p:wipe/>"},
    {"ui": "分割", "group": "细微", "en": "Split", "family": "core",
     "spec": "split", "child": "<p:split/>"},
    {"ui": "显示", "group": "细微", "en": "Reveal", "family": "p14",
     "spec": "reveal", "child": "<p14:reveal/>"},
    {"ui": "切入", "group": "细微", "en": "Cut", "family": "core",
     "spec": "cut", "child": "<p:cut/>"},
    {"ui": "随机线条", "group": "细微", "en": "Random Bars", "family": "core",
     "spec": "randombar", "child": "<p:randomBar/>"},
    {"ui": "形状", "group": "细微", "en": "Shape", "family": "core",
     "spec": "shape", "child": "<p:circle/>"},
    {"ui": "揭开", "group": "细微", "en": "Uncover", "family": "core",
     "spec": "uncover", "child": "<p:pull/>"},
    {"ui": "覆盖", "group": "细微", "en": "Cover", "family": "core",
     "spec": "cover", "child": "<p:cover/>"},
    {"ui": "闪光", "group": "细微", "en": "Flash", "family": "p14",
     "spec": "flash", "child": "<p14:flash/>"},
    {"ui": "平滑", "group": "细微", "en": "Morph", "family": "p159",
     "spec": "morph", "child": '<p159:morph option="byObject"/>'},
    # ---- 华丽 (29) --------------------------------------------------------
    {"ui": "跌落", "group": "华丽", "en": "Fall Over", "family": "p15",
     "spec": "fall_over", "child": '<p15:prstTrans prst="fallOver"/>'},
    {"ui": "悬挂", "group": "华丽", "en": "Drape", "family": "p15",
     "spec": "drape", "child": '<p15:prstTrans prst="drape"/>'},
    {"ui": "帘式", "group": "华丽", "en": "Curtains", "family": "p15",
     "spec": "curtains", "child": '<p15:prstTrans prst="curtains"/>'},
    {"ui": "风", "group": "华丽", "en": "Wind", "family": "p15",
     "spec": "wind", "child": '<p15:prstTrans prst="wind"/>'},
    {"ui": "上拉帷幕", "group": "华丽", "en": "Prestige", "family": "p15",
     "spec": "prestige", "child": '<p15:prstTrans prst="prestige"/>'},
    {"ui": "折断", "group": "华丽", "en": "Fracture", "family": "p15",
     "spec": "fracture", "child": '<p15:prstTrans prst="fracture"/>'},
    {"ui": "压碎", "group": "华丽", "en": "Crush", "family": "p15",
     "spec": "crush", "child": '<p15:prstTrans prst="crush"/>'},
    {"ui": "剥离", "group": "华丽", "en": "Peel Off", "family": "p15",
     "spec": "peel_off", "child": '<p15:prstTrans prst="peelOff"/>'},
    {"ui": "页面卷曲", "group": "华丽", "en": "Page Curl", "family": "p15",
     "spec": "page_curl", "child": '<p15:prstTrans prst="pageCurlSingle"/>'},
    {"ui": "飞机", "group": "华丽", "en": "Airplane", "family": "p15",
     "spec": "airplane", "child": '<p15:prstTrans prst="airplane"/>'},
    {"ui": "日式折纸", "group": "华丽", "en": "Origami", "family": "p15",
     "spec": "origami", "child": '<p15:prstTrans prst="origami"/>'},
    {"ui": "溶解", "group": "华丽", "en": "Dissolve", "family": "core",
     "spec": "dissolve", "child": "<p:dissolve/>"},
    {"ui": "棋盘", "group": "华丽", "en": "Checkerboard", "family": "core",
     "spec": "checkerboard", "child": "<p:checker/>"},
    {"ui": "百叶窗", "group": "华丽", "en": "Blinds", "family": "core",
     "spec": "blinds", "child": "<p:blinds/>"},
    {"ui": "时钟", "group": "华丽", "en": "Clock", "family": "core",
     "spec": "clock", "child": '<p:wheel spokes="1"/>'},
    {"ui": "涟漪", "group": "华丽", "en": "Ripple", "family": "p14",
     "spec": "ripple", "child": "<p14:ripple/>"},
    {"ui": "蜂巢", "group": "华丽", "en": "Honeycomb", "family": "p14",
     "spec": "honeycomb", "child": "<p14:honeycomb/>"},
    {"ui": "闪罐", "group": "华丽", "en": "Glitter", "family": "p14",
     "spec": "glitter", "child": "<p14:glitter/>"},
    {"ui": "涡流", "group": "华丽", "en": "Vortex", "family": "p14",
     "spec": "vortex", "child": "<p14:vortex/>"},
    {"ui": "碎片", "group": "华丽", "en": "Shred", "family": "p14",
     "spec": "shred", "child": "<p14:shred/>"},
    {"ui": "切换", "group": "华丽", "en": "Switch", "family": "p14",
     "spec": "switch", "child": '<p14:switch dir="l"/>'},
    {"ui": "翻转", "group": "华丽", "en": "Flip", "family": "p14",
     "spec": "flip", "child": '<p14:flip dir="l"/>'},
    {"ui": "库", "group": "华丽", "en": "Gallery", "family": "p14",
     "spec": "gallery", "child": '<p14:gallery dir="l"/>'},
    {"ui": "立方体", "group": "华丽", "en": "Cube", "family": "p14",
     "spec": "cube", "child": "<p14:prism/>"},
    {"ui": "门", "group": "华丽", "en": "Doors", "family": "p14",
     "spec": "doors", "child": "<p14:doors/>"},
    {"ui": "框", "group": "华丽", "en": "Box", "family": "core",
     "spec": "box", "child": "<p:zoom/>"},
    {"ui": "梳理", "group": "华丽", "en": "Comb", "family": "core",
     "spec": "comb", "child": "<p:comb/>"},
    {"ui": "缩放", "group": "华丽", "en": "Zoom", "family": "p14",
     "spec": "zoom2", "child": "<p14:warp/>"},
    {"ui": "随机", "group": "华丽", "en": "Random", "family": "core",
     "spec": "random", "child": "<p:random/>"},
    # ---- 动态内容 (7) ------------------------------------------------------
    {"ui": "平移", "group": "动态内容", "en": "Pan", "family": "p14",
     "spec": "pan", "child": "<p14:pan/>"},
    {"ui": "摩天轮", "group": "动态内容", "en": "Ferris Wheel", "family": "p14",
     "spec": "ferris_wheel", "child": '<p14:ferris dir="l"/>'},
    {"ui": "传送带", "group": "动态内容", "en": "Conveyor", "family": "p14",
     "spec": "conveyor", "child": '<p14:conveyor dir="l"/>'},
    {"ui": "旋转", "group": "动态内容", "en": "Rotate", "family": "p14",
     "spec": "rotate", "child": '<p14:prism isContent="1"/>'},
    {"ui": "窗口", "group": "动态内容", "en": "Window", "family": "p14",
     "spec": "window", "child": "<p14:window/>"},
    {"ui": "轨道", "group": "动态内容", "en": "Orbit", "family": "p14",
     "spec": "orbit", "child": '<p14:prism isContent="1" isInverted="1"/>'},
    {"ui": "飞过", "group": "动态内容", "en": "Fly Through", "family": "p14",
     "spec": "fly_through", "child": "<p14:flythrough/>"},
]

# --------------------------------------------------------------------------
# The 13 templates motion.py ships today (plus morph, which it builds
# separately). They are probed alongside the gallery items because several of
# them are NOT the gallery default: motion's "push" is dir="u" while the
# gallery's 推入 is the bare (dir="l") form, and its "zoom" is dir="in" while
# the gallery's 框 is dir="out". Refactoring motion.py onto the measured table
# must not change what any of these names produce, so each one gets its own
# deck and its own measured readback -- not a guess derived from the gallery.
# --------------------------------------------------------------------------
MOTION_SPECS = [
    {"ui": "淡入/淡出", "group": "motion", "en": "Fade", "family": "core",
     "spec": "fade", "child": "<p:fade/>"},
    {"ui": "淡入/淡出", "group": "motion", "en": "Fade", "family": "core",
     "spec": "smoothfade", "child": "<p:fade/>"},
    {"ui": "淡出为黑", "group": "motion", "en": "Fade Through Black",
     "family": "core", "spec": "fadeblack", "child": '<p:fade thruBlk="1"/>'},
    {"ui": "推入(上)", "group": "motion", "en": "Push Up", "family": "core",
     "spec": "push", "child": '<p:push dir="u"/>'},
    {"ui": "推入(左)", "group": "motion", "en": "Push Left", "family": "core",
     "spec": "pushleft", "child": '<p:push dir="l"/>'},
    {"ui": "擦除(左)", "group": "motion", "en": "Wipe Left", "family": "core",
     "spec": "wipe", "child": '<p:wipe dir="l"/>'},
    {"ui": "覆盖(左)", "group": "motion", "en": "Cover Left", "family": "core",
     "spec": "cover", "child": '<p:cover dir="l"/>'},
    {"ui": "分割(外)", "group": "motion", "en": "Split Out", "family": "core",
     "spec": "split", "child": '<p:split orient="horz" dir="out"/>'},
    {"ui": "缩放(内)", "group": "motion", "en": "Zoom In", "family": "core",
     "spec": "zoom", "child": '<p:zoom dir="in"/>'},
    {"ui": "溶解", "group": "motion", "en": "Dissolve", "family": "core",
     "spec": "dissolve", "child": "<p:dissolve/>"},
    {"ui": "条带", "group": "motion", "en": "Strips", "family": "core",
     "spec": "strips", "child": "<p:strips/>"},
    {"ui": "揭开(左)", "group": "motion", "en": "Uncover Left", "family": "core",
     "spec": "pull", "child": '<p:pull dir="l"/>'},
    {"ui": "随机线条", "group": "motion", "en": "Random Bars", "family": "core",
     "spec": "randombar", "child": '<p:randomBar/>'},
    {"ui": "平滑", "group": "motion", "en": "Morph", "family": "p159",
     "spec": "morph", "child": '<p159:morph option="byObject"/>'},
]

DUR_MS = 800          # long enough to see, short enough for clean boundaries
PROBE_SPD = "slow"    # what the probe writes; parameterised back out in table
SLIDE_SECONDS = 2     # CreateVideo DefaultSlideDuration
FPS = 30              # must match what probe_anchor.ps1 exports at

# Every rule here was measured, not read out of a spec. They travel with the
# table because a table of element names without them invites exactly the
# mistakes that produced them.
RULES = [
    "A transition lives in one of four namespaces and the UI grouping does not "
    "tell you which: 细微 mixes core ECMA-376 elements with p14 (显示/闪光) and "
    "p159 (平滑); all of 华丽's 2013+ items are p15:prstTrans; the 2010 items in "
    "华丽/动态内容 are p14.",
    "Every transition ends up inside mc:AlternateContent -- even a plain core "
    "one -- because p14:dur is a 2010 attribute and PowerPoint wraps the whole "
    "<p:transition> rather than drop the duration.",
    "The mc:Fallback is NOT always <p:fade/>. A core element repeats itself in "
    "the Fallback (it is already backwards-compatible); only a p14/p15/p159 "
    "child degrades to fade.",
    "PowerPoint deletes attributes that carry the default value: <p:push "
    "dir=\"l\"/> is saved as <p:push/>, <p:split orient=\"horz\" dir=\"out\"/> "
    "as <p:split/>. Writing them is harmless but never survives a save.",
    "切换 / 翻转 / 库 / 摩天轮 / 传送带 are REFUSED without a dir attribute -- "
    "a bare <p14:flip/> makes PowerPoint reject the entire file, not just that "
    "slide. PowerPoint itself always writes dir=\"l\".",
    "Namespace declarations may ride on <mc:AlternateContent>, <mc:Choice> or "
    "<p:transition>; PowerPoint moves them around freely and adds a stray "
    "xmlns=\"\" on <mc:Fallback>. Comparing saved XML as raw strings therefore "
    "reports false differences -- compare the element structure, not the text.",
    "Duplicate xmlns declarations on <p:sld> make the part not well-formed and "
    "PowerPoint refuses the whole deck. Add only the prefixes that are missing.",
]


def wrap_transition(family, child, spd="slow", dur=DUR_MS):
    """Return the full XML block for one candidate, per its family.

    MEASURED, not assumed. Two rules came out of the roundtrip:

    * Every transition ends up inside mc:AlternateContent -- even a plain core
      one. The reason is p14:dur: a 2010 attribute needs a Choice/Fallback, and
      PowerPoint wraps the whole <p:transition> rather than dropping dur.
    * The Fallback is NOT always <p:fade/>. For a core element the Fallback
      repeats the core element itself (it is already backwards-compatible);
      only a p14/p15/p159 child degrades to fade. Writing <p:fade/> as the
      Fallback for <p:push/> is accepted but PowerPoint rewrites it -- 16 of
      the first 48 candidates came back "rewritten" purely because of this.
    """
    if family == "core":
        return (
            '<mc:AlternateContent>'
            '<mc:Choice xmlns:p14="%s" Requires="p14">'
            '<p:transition spd="%s" p14:dur="%d">%s</p:transition>'
            '</mc:Choice>'
            '<mc:Fallback><p:transition spd="%s">%s</p:transition>'
            '</mc:Fallback></mc:AlternateContent>'
        ) % (NS["p14"], spd, dur, child, spd, child)
    if family in ("p14", "p15", "p159"):
        prefix = family
        return (
            '<mc:AlternateContent>'
            '<mc:Choice xmlns:%s="%s" Requires="%s">'
            '<p:transition spd="%s" p14:dur="%d">%s</p:transition>'
            '</mc:Choice>'
            '<mc:Fallback><p:transition spd="%s"><p:fade/></p:transition>'
            '</mc:Fallback></mc:AlternateContent>'
        ) % (prefix, NS[prefix], prefix, spd, dur, child, spd)
    raise ValueError("unknown family %r" % family)


def _slide_xml(data, block):
    """Declare every extension namespace on <p:sld> and insert `block` after
    <p:clrMapOvr> (position rule: cSld, clrMapOvr?, transition?, timing? --
    a misplaced block is silently dropped on save, com-pitfalls §12).

    Only the MISSING prefixes are added. python-pptx already writes xmlns:a,
    xmlns:p and xmlns:r on <p:sld>; appending a second xmlns:a makes the part
    not well-formed, and PowerPoint's answer to that is to refuse the WHOLE
    deck (observed: all 48 candidates refused, including a plain <p:fade/>,
    which is how this was found -- §26 again, but the fault was ours)."""
    xml = data.decode("utf-8")
    m = re.search(r"<p:sld(\s[^>]*?)?>", xml)
    have = set(re.findall(r'xmlns:(\w+)=', m.group(1) or ""))
    add = " ".join('xmlns:%s="%s"' % (k, v) for k, v in sorted(NS.items())
                   if k not in have)
    def repl(mm):
        head = mm.group(0)[:-1].rstrip()
        return "%s %s>" % (head, add) if add else mm.group(0)
    xml = re.sub(r"<p:sld(\s[^>]*?)?>", repl, xml, count=1)
    anchor = "</p:clrMapOvr>"
    idx = xml.rindex(anchor) + len(anchor)
    return (xml[:idx] + block + xml[idx:]).encode("utf-8")


def _two_slide_deck(path, label, hyp=None):
    """A minimal deck: slide 1 'from', slide 2 'entering' (+ candidate)."""
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]

    def add(bg, big, small):
        s = prs.slides.add_slide(blank)
        r = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor(*bg)
        r.line.fill.background()
        r.shadow.inherit = False
        tb = s.shapes.add_textbox(Pt(60), Pt(60), prs.slide_width - Pt(120),
                                  prs.slide_height - Pt(120))
        tf = tb.text_frame
        tf.text = big
        p0 = tf.paragraphs[0]
        p0.font.size = Pt(120)
        p0.font.bold = True
        p1 = tf.add_paragraph()
        p1.text = small
        p1.font.size = Pt(20)

    add((235, 220, 200), label, "from")
    add((70, 110, 190), label, hyp["ui"] if hyp else "target")
    prs.save(path)

    if hyp is None:
        return
    buf = io.BytesIO(open(path, "rb").read())
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide2\.xml", item.filename):
                data = _slide_xml(data, wrap_transition(hyp["family"],
                                                        hyp["child"]))
            zout.writestr(item, data)
    open(path, "wb").write(out.getvalue())


def cmd_build(case_dir):
    os.makedirs(case_dir, exist_ok=True)
    for f in glob.glob(os.path.join(case_dir, "*.pptx")):
        os.remove(f)
    cases = []
    for i, h in enumerate(HYPOTHESES):
        name = "cand%02d-%s.pptx" % (i + 1, h["spec"])
        _two_slide_deck(os.path.join(case_dir, name), str(i + 1), h)
        cases.append(name)
    for i, h in enumerate(MOTION_SPECS):
        name = "mspec%02d-%s.pptx" % (i + 1, h["spec"])
        _two_slide_deck(os.path.join(case_dir, name), "m%d" % (i + 1), h)
        cases.append(name)
    manifest = {"hypotheses": HYPOTHESES, "motion_specs": MOTION_SPECS,
                "dur_ms": DUR_MS, "slide_seconds": SLIDE_SECONDS,
                "cases": cases}
    json.dump(manifest, io.open(os.path.join(case_dir, "manifest.json"), "w",
                                encoding="utf-8"), ensure_ascii=False, indent=1)
    print("built %d case decks (%d gallery + %d motion spec) -> %s" % (
        len(cases), len(HYPOTHESES), len(MOTION_SPECS), case_dir))
    return 0


# --------------------------------------------------------------------------
# collect: what did PowerPoint do to each candidate?
# --------------------------------------------------------------------------
def _extract_transition(xml):
    """Normalised <p:transition> content of a slide (AlternateContent kept
    whole: Choice + Fallback is BY DESIGN, not a duplicate -- §20/§31)."""
    m = re.search(r"<mc:AlternateContent(?=[\s/>]).*?</mc:AlternateContent>",
                  xml, re.S)
    if m:
        return re.sub(r">\s+<", "><", m.group(0))
    m = re.search(r"<p:transition\b.*?</p:transition>|<p:transition\b[^>]*/>",
                  xml, re.S)
    return re.sub(r">\s+<", "><", m.group(0)) if m else ""


def _load_json(path):
    """PowerShell's Set-Content -Encoding UTF8 writes a BOM; json.load chokes
    on it unless the codec is utf-8-sig. Every file this script reads may have
    come from either side, so always open with utf-8-sig (it is identical to
    utf-8 when there is no BOM)."""
    return json.load(io.open(path, encoding="utf-8-sig"))


def _norm(s):
    """Compare semantics, not namespace-declaration placement.

    PowerPoint moves the xmlns:* declarations off <p:sld> and onto
    <mc:AlternateContent> (and sometimes onto <p:transition>, and it adds a
    stray xmlns="" on <mc:Fallback>). None of that changes what the block
    MEANS, so comparing raw strings reports every single candidate as
    "rewritten" -- which is noise, not a finding."""
    s = re.sub(r'\s+xmlns:\w+="[^"]*"', "", s)
    s = re.sub(r'\s+xmlns=""', "", s)
    return re.sub(r">\s+<", "><", s).strip()


def cmd_collect(case_dir, out_dir, json_out, enum_path=None):
    man = _load_json(os.path.join(case_dir, "manifest.json"))
    # probe_transitions.ps1 writes {"decks":[{deck,opened,slides:[...]}, ...]}.
    # Three distinct outcomes per deck, and they must NOT be merged:
    #   not in enum.json at all  -> UNPROBED   (the probe never reached it)
    #   opened=false             -> REJECTED   (PowerPoint refused the file)
    #   present in OutDir        -> compare wrote vs readback
    decks = {}
    for cand in filter(None, [enum_path, os.path.join(out_dir, "enum.json"),
                              os.path.join(os.path.dirname(out_dir.rstrip("\\/")),
                                           "enum.json")]):
        if not os.path.exists(cand):
            continue
        for d in _load_json(cand).get("decks", []):
            decks[d["deck"]] = d
        break
    rows = []
    probed = [(h, "cand%02d-%s.pptx" % (i + 1, h["spec"]), "gallery")
              for i, h in enumerate(man["hypotheses"])]
    probed += [(h, "mspec%02d-%s.pptx" % (i + 1, h["spec"]), "motion")
               for i, h in enumerate(man.get("motion_specs", []))]
    for h, name, source in probed:
        case = os.path.join(case_dir, name)
        done = os.path.join(out_dir, name)
        d = decks.get(name)
        base = dict(ui=h["ui"], group=h["group"], en=h["en"],
                    family=h["family"], spec=h["spec"], source=source)
        if d is not None and not d.get("opened", True):
            rows.append(dict(base, verdict="REJECTED", wrote="", readback="",
                             error=d.get("error", "")))
            continue
        if not os.path.exists(done):
            rows.append(dict(base, verdict="UNPROBED", wrote="", readback=""))
            continue
        wrote = _extract_transition(zipfile.ZipFile(case).read(
            "ppt/slides/slide2.xml").decode("utf-8"))
        readback = _extract_transition(zipfile.ZipFile(done).read(
            "ppt/slides/slide2.xml").decode("utf-8"))
        e = {}
        if d:
            e = {int(s["slide"]): s for s in d.get("slides", [])}.get(2, {})
        same = _norm(wrote) == _norm(readback)
        rows.append(dict(base, wrote=wrote, readback=readback,
                         preserved=same,
                         entryEffect=e.get("entryEffect"),
                         comDuration=e.get("duration"),
                         verdict=("preserved" if same else
                                  "rewritten" if readback else "dropped")))
    report = {"measured_on": "PowerPoint COM roundtrip (open + SaveAs)",
              "rows": rows,
              "summary": {
                  "total": len(rows),
                  "preserved": sum(1 for r in rows
                                   if r["verdict"] == "preserved"),
                  "rewritten": sum(1 for r in rows if r["verdict"] == "rewritten"),
                  "dropped": sum(1 for r in rows if r["verdict"] == "dropped"),
                  "rejected": sum(1 for r in rows if r["verdict"] == "REJECTED"),
                  "unprobed": sum(1 for r in rows if r["verdict"] == "UNPROBED"),
              }}
    if json_out:
        json.dump(report, io.open(json_out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    s = report["summary"]
    print("== 切换候选 往返结果 ==")
    print("  候选 %d | 原样保留 %d | 被改写 %d | 被丢弃 %d | 拒开 %d | 未测 %d" % (
        s["total"], s["preserved"], s["rewritten"], s["dropped"],
        s["rejected"], s["unprobed"]))
    for r in rows:
        if r["verdict"] == "preserved":
            continue
        rb = r["readback"] or (r.get("error") or "(PowerPoint 写回时没了)")
        print("  [%-8s] %-6s %-6s %s" % (r["verdict"][:8], r["ui"],
                                         r["family"], rb[:110]))
    if json_out:
        print("  -> %s" % json_out)
    return 0


# --------------------------------------------------------------------------
# enumdeck / enumread: let PowerPoint write the transitions ITSELF
# --------------------------------------------------------------------------
# Roundtrip probing settles "is this child element valid", but it cannot invent
# a NAME we guessed wrong -- a wrong name comes back as "refused", which tells
# us nothing about the right one. For those, stop guessing and let PowerPoint
# write: put one slide per PpEntryEffect value in a deck, set every slide's
# EntryEffect over COM, save, and read back what element it produced. The
# pairing value -> XML is then PowerPoint's own, not ours.
def cmd_enumdeck(out_pptx, lo, hi):
    """One blank slide per enum value in [lo, hi]; slide k carries value lo+k."""
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    for _ in range(lo, hi + 1):
        prs.slides.add_slide(blank)
    prs.save(out_pptx)
    json.dump({"lo": lo, "hi": hi, "count": hi - lo + 1},
              io.open(os.path.splitext(out_pptx)[0] + ".manifest.json", "w",
                      encoding="utf-8"), indent=1)
    print("enumdeck %s: %d slides, values %d..%d" % (
        out_pptx, hi - lo + 1, lo, hi))
    return 0


def cmd_enumread(deck, json_out, manifest=None):
    """Read each slide back: enum value -> the XML PowerPoint wrote for it."""
    # The scanned deck is saved under a different name than the deck that was
    # built (probe_enum_scan.ps1 needs a separate output), so the manifest
    # lives next to the INPUT deck, not next to this one.
    if manifest is None:
        manifest = os.path.join(os.path.dirname(deck), "enumdeck.manifest.json")
    man = _load_json(manifest)
    lo, hi = man["lo"], man["hi"]
    out = []
    with zipfile.ZipFile(deck) as z:
        for v in range(lo, hi + 1):
            name = "ppt/slides/slide%d.xml" % (v - lo + 1)
            if name not in z.namelist():
                continue
            xml = z.read(name).decode("utf-8")
            blk = _extract_transition(xml)
            out.append({"value": v, "xml": blk})
    known = [r for r in out if r["xml"]]
    if json_out:
        json.dump(out, io.open(json_out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    print("enumread: %d slides, %d produced a transition block" % (
        len(out), len(known)))
    for r in known:
        print("  %5d  %s" % (r["value"], r["xml"][:150]))
    if json_out:
        print("  -> %s" % json_out)
    return 0


# --------------------------------------------------------------------------
# table: merge the two measurements into the shipped reference file
# --------------------------------------------------------------------------
# Attributes PowerPoint drops because they carry the default value. The enum
# scan cannot show these (it only ever writes the value PowerPoint chose), but
# the roundtrip can: writing <p:push dir="l"/> and saving gives back <p:push/>.
# Needed to look a motion-spec child up in the enum scan, which is keyed by
# PowerPoint's OWN normalized form.
DEFAULT_ATTRS = {("dir", "l"), ("orient", "horz"), ("dir", "out")}


def _child_tag(child):
    m = re.match(r"\s*<([\w:]+)", child)
    return m.group(1) if m else ""


def _child_attrs(child):
    return re.findall(r'([\w:]+)="([^"]*)"', child)


def _child_render(tag, attrs):
    if not attrs:
        return "<%s/>" % tag
    return "<%s %s/>" % (tag, " ".join('%s="%s"' % kv for kv in attrs))


def _lookup_enum(enum_by_child, child):
    """Find the enum value for a child, tolerating default-valued attributes.

    Tries the child as written, then with any subset of default-valued
    attributes removed. It deliberately does NOT fall back to "strip every
    attribute": <p:zoom dir="in"/> is a real, different transition from
    <p:zoom/>, so a greedy fallback would silently return the wrong number.
    """
    tag = _child_tag(child)
    attrs = _child_attrs(child)
    droppable = [i for i, kv in enumerate(attrs) if kv in DEFAULT_ATTRS]
    for mask in range(1 << len(droppable)):
        keep = [kv for i, kv in enumerate(attrs)
                if not any(droppable[j] == i and (mask >> j) & 1
                           for j in range(len(droppable)))]
        hit = enum_by_child.get(_child_render(tag, keep))
        if hit is not None:
            return hit
    return None


def _template(readback):
    """Parameterise PowerPoint's own block so callers can set speed/duration.

    The shipped table stores what PowerPoint WROTE, not what we guessed, but a
    stored 800ms/"slow" block would be useless to a caller. So the two values
    the probe fixed are turned back into placeholders -- the structure around
    them (AlternateContent, Choice/Fallback, which namespace rides where) stays
    exactly as PowerPoint emitted it.
    """
    t = readback.replace('spd="%s"' % PROBE_SPD, 'spd="{spd}"')
    t = t.replace(' p14:dur="%d"' % DUR_MS, "{dur}")
    return t


def cmd_table(report, scans, out):
    rep = _load_json(report)
    enum_by_child = {}
    for s in scans:
        for r in _load_json(s):
            if not r.get("xml"):
                continue
            m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>",
                          r["xml"], re.S)
            if not m:
                continue
            child = re.sub(r"\s+", " ", m.group(1)).strip()
            enum_by_child.setdefault(child, r["value"])

    def base_id(child):
        tag = _child_tag(child)
        attrs = dict(_child_attrs(child))
        return "%s:%s" % (tag, attrs.get("prst", ""))

    by_spec, gallery, claimed = {}, [], set()
    for n, r in enumerate(rep["rows"]):
        if not r.get("readback"):
            continue
        m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>",
                      r["readback"], re.S)
        child = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
        entry = {
            "spec": r["spec"], "ui": r["ui"], "en": r["en"],
            "group": r["group"], "family": r["family"], "child": child,
            "xml": _template(r["readback"]),
            "entryEffect": _lookup_enum(enum_by_child, child),
            "roundtrip": r["verdict"],
            "order": n,          # gallery order, so docs can render as the UI does
            "source": r.get("source"),
        }
        claimed.add(base_id(child))
        if r.get("source") == "motion":
            by_spec[r["spec"]] = entry
        else:
            gallery.append(entry)

    # Everything PowerPoint can write that no gallery item and no motion spec
    # claims: still valid in the file format, just not reachable from the UI.
    legacy = []
    for child, v in sorted(enum_by_child.items(), key=lambda kv: kv[1]):
        if base_id(child) in claimed:
            continue
        legacy.append({"child": child, "entryEffect": v})

    doc = {
        "meta": {
            "purpose": "UI gallery item -> the XML PowerPoint actually writes.",
            "measured_by": "scripts/build_transition_table.py (build/collect/"
                           "enumdeck/enumread/table) + scripts/probe_transitions.ps1 "
                           "+ scripts/probe_enum_scan.ps1",
            "method": [
                "enum scan: one slide per PpEntryEffect value, set over COM, "
                "saved, read back -- gives PowerPoint's own element per value",
                "roundtrip: one 2-slide deck per candidate, opened and saved "
                "-- gives 'accepted and kept' plus the canonical wrapper",
            ],
            "enum_range_scanned": "1..4200 (no transition exists above 3956)",
            "rules": RULES,
        },
        "namespaces": NS,
        "by_spec": by_spec,
        "gallery": gallery,
        "enum_by_child": enum_by_child,
        "not_in_gallery": legacy,
        "summary": {
            "gallery": len(gallery),
            "motion_specs": len(by_spec),
            "enum_values": len(enum_by_child),
            "not_in_gallery": len(legacy),
        },
    }
    json.dump(doc, io.open(out, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    s = doc["summary"]
    print("table: %d gallery + %d motion spec, %d enum values, "
          "%d reachable only in XML -> %s" % (
              s["gallery"], s["motion_specs"], s["enum_values"],
              s["not_in_gallery"], out))
    return 0


# --------------------------------------------------------------------------
# anchordeck: does a <p:transition> written on slide N drive N-1 -> N?
# --------------------------------------------------------------------------
# Two claims every user of the table makes, and neither had been measured:
#
#   A. A transition is written on the slide it ANIMATES INTO. So it lives on
#      slide N and drives N-1 -> N. Writing it on slide 1 should do nothing
#      visible. Prove it with rendered frames, not with the XML: the XML is
#      there either way, so "the file contains it" proves nothing.
#   B. <p:transition> takes exactly ONE child. So morph and push cannot both be
#      in force -- they compete for the same slot. (This is what decides whether
#      "shape deformation AND whole-page movement" is even expressible.)
#
# The anchor deck puts an 800ms push on slide 1 and an 800ms wipe on slide 3,
# leaving slide 2 empty. Rendered at 30fps with 2s per slide the boundaries sit
# near frames 60 and 120. If (A) holds there is a multi-frame change burst only
# at the second boundary; the first is a one-frame hard cut.
def _flat_deck(path, labels, blocks, dur=DUR_MS):
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    palette = [(200, 60, 60), (60, 150, 90), (60, 80, 190), (190, 150, 40)]
    for i, text in enumerate(labels):
        s = prs.slides.add_slide(blank)
        r = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor(*palette[i % len(palette)])
        r.line.fill.background()
        r.shadow.inherit = False
        tb = s.shapes.add_textbox(Pt(60), Pt(60), prs.slide_width - Pt(120),
                                  prs.slide_height - Pt(120))
        tf = tb.text_frame
        tf.text = text
        tf.paragraphs[0].font.size = Pt(240)
        tf.paragraphs[0].font.bold = True
    prs.save(path)

    assigned = {}
    buf = io.BytesIO(open(path, "rb").read())
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
            if m and int(m.group(1)) in blocks:
                data = _slide_xml(data, blocks[int(m.group(1))])
                assigned[int(m.group(1))] = blocks[int(m.group(1))]
            zout.writestr(item, data)
    open(path, "wb").write(out.getvalue())
    return assigned


def cmd_anchordeck(out_dir):
    """Two decks that can FALSIFY claim A, plus one that probes claim B.

    anchor.pptx -- three slides, a push written on slide 2 ONLY.
        A holds  -> change burst at boundary 1 (1->2), hard cut at boundary 2.
        A is wrong (transition drives the page it is written on OUT) ->
        hard cut at boundary 1, burst at boundary 2.
    first.pptx  -- two slides, a wipe written on slide 1 ONLY.
        A holds  -> no burst anywhere: slide 1 has no predecessor to move from.
        A is wrong -> burst at the single boundary.
    twochild.pptx -- one <p:transition> holding two children at once (claim B).
    """
    os.makedirs(out_dir, exist_ok=True)
    tr_push = wrap_transition("core", '<p:push dir="u"/>', dur=DUR_MS)
    tr_wipe = wrap_transition("core", '<p:wipe dir="l"/>', dur=DUR_MS)

    anchor = os.path.join(out_dir, "anchor.pptx")
    got = _flat_deck(anchor, ["1", "2", "3"], {2: tr_push})
    json.dump({"case": "transition written on slide 2 only",
               "transitions_on": sorted(got), "dur_ms": DUR_MS, "fps": FPS,
               "expected_boundary_frames": [1 * SLIDE_SECONDS * 30,
                                            2 * SLIDE_SECONDS * 30],
               "prediction": ("A: burst at boundary 1, hard cut at boundary 2 / "
                              "not-A: hard cut at 1, burst at 2")},
              io.open(os.path.join(out_dir, "anchor.manifest.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)

    first = os.path.join(out_dir, "first.pptx")
    got1 = _flat_deck(first, ["1", "2"], {1: tr_wipe})
    json.dump({"case": "transition written on slide 1 only",
               "transitions_on": sorted(got1), "dur_ms": DUR_MS, "fps": FPS,
               "expected_boundary_frames": [SLIDE_SECONDS * 30],
               "prediction": "A: no burst at all / not-A: burst at boundary 1"},
              io.open(os.path.join(out_dir, "first.manifest.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)

    # Claim B: two children competing for the one slot.
    two = os.path.join(out_dir, "twochild.pptx")
    both = ('<mc:AlternateContent xmlns:mc="%s" xmlns:p159="%s" xmlns:p14="%s">'
            '<mc:Choice Requires="p159">'
            '<p:transition spd="slow" p14:dur="%d">'
            '<p:push dir="u"/><p159:morph option="byObject"/>'
            '</p:transition></mc:Choice>'
            '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition>'
            '</mc:Fallback></mc:AlternateContent>'
            ) % (NS["mc"], NS["p159"], NS["p14"], DUR_MS)
    _flat_deck(two, ["one slot", "two children"], {2: both})

    # Same claim with two core children, to show it is not a p159 peculiarity.
    twoc = os.path.join(out_dir, "twochild_core.pptx")
    both_c = ('<mc:AlternateContent xmlns:mc="%s" xmlns:p14="%s">'
              '<mc:Choice Requires="p14">'
              '<p:transition spd="slow" p14:dur="%d">'
              '<p:push dir="u"/><p:wipe dir="l"/>'
              '</p:transition></mc:Choice>'
              '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition>'
              '</mc:Fallback></mc:AlternateContent>'
              ) % (NS["mc"], NS["p14"], DUR_MS)
    _flat_deck(twoc, ["one slot", "two children"], {2: both_c})

    # Claim C: spd (three coarse notches) and p14:dur (exact milliseconds) encode
    # the SAME quantity. Write them in conflict and see which one the renderer
    # obeys. slow is ~2s, fast is ~0.5s; the dur values say the opposite.
    duel = os.path.join(out_dir, "duel.pptx")
    slow_short = ('<mc:AlternateContent xmlns:mc="%s" xmlns:p14="%s">'
                  '<mc:Choice Requires="p14">'
                  '<p:transition spd="slow" p14:dur="200">'
                  '<p:push dir="u"/>'
                  '</p:transition></mc:Choice>'
                  '<mc:Fallback><p:transition spd="slow"><p:fade/>'
                  '</p:transition></mc:Fallback></mc:AlternateContent>'
                  ) % (NS["mc"], NS["p14"])
    fast_long = ('<mc:AlternateContent xmlns:mc="%s" xmlns:p14="%s">'
                 '<mc:Choice Requires="p14">'
                 '<p:transition spd="fast" p14:dur="1500">'
                 '<p:wipe dir="l"/>'
                 '</p:transition></mc:Choice>'
                 '<mc:Fallback><p:transition spd="fast"><p:fade/>'
                 '</p:transition></mc:Fallback></mc:AlternateContent>'
                 ) % (NS["mc"], NS["p14"])
    _flat_deck(duel, ["1", "2", "3"], {2: slow_short, 3: fast_long})
    json.dump({"case": "spd vs p14:dur in conflict",
               "transitions_on": [2, 3], "dur_ms": DUR_MS, "fps": FPS,
               "expected_boundary_frames": [60, 120],
               "prediction": ("dur wins -> ~6 frames at boundary 1, ~45 at 2 / "
                              "spd wins -> ~60 frames at boundary 1, ~15 at 2")},
               io.open(os.path.join(out_dir, "duel.manifest.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)

    # Claim C follow-up: the Choice branch obeys milliseconds, but mc:Fallback
    # carries NO p14:dur -- a 2010 attribute cannot appear in the fallback world.
    # There the only knob is spd, so "how long is slow/med/fast" decides what old
    # PowerPoint / WPS / online preview actually show. One element, three notches.
    notch = os.path.join(out_dir, "notch.pptx")
    ns_rows = []
    for k, spd in enumerate(("slow", "med", "fast")):
        ns_rows.append('<p:transition spd="%s"><p:dissolve/></p:transition>' % spd)
    _flat_deck(notch, ["1", "2", "3", "4"], {2: ns_rows[0], 3: ns_rows[1],
                                             4: ns_rows[2]})
    json.dump({"case": "spd notches with NO p14:dur (the fallback world)",
               "transitions_on": [2, 3, 4], "fps": FPS,
               "expected_boundary_frames": [60, 120, 180],
               "prediction": "burst length per notch = its millisecond value"},
              io.open(os.path.join(out_dir, "notch.manifest.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("anchordeck -> %s" % out_dir)
    print("  anchor.pptx   3 页，只有第 2 页写 push   <- 能否扭转「写在哪页」")
    print("  first.pptx    2 页，只有第 1 页写 wipe   <- 第一页有没有得可动")
    print("  twochild.pptx      push + morph 同槽（扩展 vs p159）")
    print("  twochild_core.pptx push + wipe  同槽（两个都是 core）")
    print("  duel.pptx     3 页，第 2/3 页故意让 spd 与 p14:dur 打架")
    print("  notch.pptx    4 页，slow/med/fast 只写 spd、不写 p14:dur（降级世界）")
    return 0


# --------------------------------------------------------------------------
# anchors: where did the rendered frames actually change?
# --------------------------------------------------------------------------
def cmd_anchors(video, manifest):
    import cv2
    import numpy as np

    man = _load_json(manifest)
    cap = cv2.VideoCapture(video)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    gray = [cv2.cvtColor(f, cv2.COLOR_BGR2GRAY) for f in frames]
    d = [0.0] + [float(np.mean(cv2.absdiff(gray[i], gray[i - 1])))
                 for i in range(1, len(frames))]

    fps = man.get("fps", 30)
    # A global absdiff threshold misses slow reveals: a 1.5s wipe moves an EDGE,
    # so per frame only ~2% of the pixels change and the mean stays under noise.
    # These probe decks are flat single-colour fills, so asking "did the average
    # colour move at all" has no such blind spot. Signature = rounded mean BGR.
    sig = []
    for f in frames:
        m = f.mean(axis=(0, 1))
        sig.append((round(float(m[0]), 1), round(float(m[1]), 1),
                    round(float(m[2]), 1)))
    changed = [i for i in range(1, len(sig)) if sig[i] != sig[i - 1]]
    if not changed:
        print("== %d 帧；全程无任何变化 ==" % len(frames))
        return 0
    events = []
    s = p = changed[0]
    for i in changed[1:]:
        if i == p + 1:
            p = i
        else:
            events.append((s, p))
            s = p = i
    events.append((s, p))

    bounds = man.get("expected_boundary_frames") or []
    print("== %s：%d 帧 ==" % (os.path.basename(video), len(frames)))
    print("  用例：%s" % man.get("case", ""))
    print("  预期：%s" % man.get("prediction", ""))
    print("  预期页边界 @%dfps：%s" % (fps, bounds))
    print()
    print("  %-14s %-6s %-8s %-10s %s" % (
        "帧区间", "帧数", "毫秒", "判定", "落在"))
    for a, b in events:
        n = b - a + 1
        ms = n * 1000.0 / fps
        if n <= 2:
            kind = "硬切"
        elif abs(ms - man.get("dur_ms", -1)) < 200:
            kind = "过渡 burst"
        else:
            kind = "多帧变化"
        near = ""
        for k, bf in enumerate(bounds):
            if a - 2 <= bf <= b + 2:
                near = "边界 %d（第%d页→第%d页）" % (k + 1, k + 1, k + 2)
        if not near:
            near = "放映起步" if a <= 3 else "非页边界"
        print("  %-14s %-6d %-8.0f %-10s %s" % (
            "%d–%d" % (a, b), n, ms, kind, near))
    return 0


# --------------------------------------------------------------------------
# video: one combined deck, for the rendering pass
# --------------------------------------------------------------------------
def cmd_video(case_dir, out_dir, out_pptx):
    """Combine the candidates that SURVIVED the roundtrip into one deck:
    slide 1 = title, then per candidate a from/target pair with the transition
    on the target. In the exported video the tested transitions appear as the
    multi-frame change bursts, in slide order."""
    man = _load_json(os.path.join(case_dir, "manifest.json"))
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]

    def add(bg, big, small):
        s = prs.slides.add_slide(blank)
        r = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor(*bg)
        r.line.fill.background()
        r.shadow.inherit = False
        tb = s.shapes.add_textbox(Pt(60), Pt(60), prs.slide_width - Pt(120),
                                  prs.slide_height - Pt(120))
        tf = tb.text_frame
        tf.text = big
        tf.paragraphs[0].font.size = Pt(120)
        tf.paragraphs[0].font.bold = True
        p1 = tf.add_paragraph()
        p1.text = small
        p1.font.size = Pt(20)

    add((150, 150, 150), "0", "probe title / none")
    blocks = []
    kept = []
    for i, h in enumerate(man["hypotheses"]):
        done = os.path.join(out_dir, man["cases"][i])
        if not os.path.exists(done):
            continue                      # rejected upstream; nothing to show
        kept.append(h)
        add((235, 220, 200), str(len(kept)), "from")
        add((70, 110, 190), h["ui"], "%s / %s" % (h["en"], h["family"]))
        blocks.append(wrap_transition(h["family"], h["child"]))
    prs.save(out_pptx)

    buf = io.BytesIO(open(out_pptx, "rb").read())
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
            if m and int(m.group(1)) % 2 == 0:
                data = _slide_xml(data, blocks[int(m.group(1)) // 2 - 1])
            zout.writestr(item, data)
    open(out_pptx, "wb").write(out.getvalue())

    json.dump({"order": [h["ui"] for h in kept],
               "slide_seconds": SLIDE_SECONDS},
              io.open(os.path.splitext(out_pptx)[0] + ".order.json", "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("video deck %s: %d slides, %d candidates" % (
        out_pptx, 1 + 2 * len(kept), len(kept)))
    return 0


# --------------------------------------------------------------------------
# sheets: contact sheets for the visual pass
# --------------------------------------------------------------------------
def cmd_sheets(video, out_dir, want=10):
    import cv2
    import numpy as np

    os.makedirs(out_dir, exist_ok=True)
    cap = cv2.VideoCapture(video)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    gray = [cv2.cvtColor(f, cv2.COLOR_BGR2GRAY) for f in frames]

    # a burst = >=3 consecutive frames that differ from their predecessor;
    # single-frame jumps are the hard cuts between untested slide pairs.
    bursts, start = [], None
    for i in range(1, len(frames)):
        d = float(np.mean(cv2.absdiff(gray[i], gray[i - 1])))
        if d > 6.0:
            if start is None:
                start = i
        elif start is not None:
            if i - start >= 3:
                bursts.append((start, i))
            start = None
    print("video %d frames, %d multi-frame bursts" % (len(frames), len(bursts)))

    order = []
    opath = os.path.splitext(video)[0] + ".order.json"
    if os.path.exists(opath):
        order = json.load(io.open(opath, encoding="utf-8"))["order"]

    for k, (a, b) in enumerate(bursts):
        lo, hi = max(0, a - 2), min(len(frames), b + 1)
        idxs = [lo + int(j * (hi - 1 - lo) / max(1, want - 1)) for j in range(want)]
        tiles = [cv2.resize(frames[i], (569, 320)) for i in idxs]
        label = order[k] if k < len(order) else "?"
        sheet = np.hstack(tiles)
        cv2.putText(sheet, "#%02d %s  f%d-%d" % (k + 1, label, a, b),
                    (8, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
        cv2.imwrite(os.path.join(out_dir, "burst%02d.png" % (k + 1)), sheet)
    print("sheets -> %s (%d)" % (out_dir, len(bursts)))
    return 0


# --------------------------------------------------------------------------
# shapedeck / shapeanalyze: WHAT does each transition look like as motion?
#
# The 48-item table answers "what element does PowerPoint write". It says
# nothing about what the viewer SEES. transition-model.md §七 already found
# that a whole-frame mean cannot see direction: a 1.5s wipe merely moves one
# edge, so the frame mean barely moves, and the same detector reads a wipe as
# a fade. Measuring SHAPE therefore requires a spatial profile, and the probe
# deck must carry spatial structure for that profile to have anything to read.
#
# Design of one probe deck (2 slides):
#   slide 1 = "FROM": a fine grid of distinct cells (each cell its own colour)
#   slide 2 = "TO"  : the same grid, each cell shifted by one step in the
#                     palette -- so every cell differs, and the difference is
#                     uniform in space. A transition that sweeps will reveal
#                     the grid in a spatial order we can recover.
# Both slides carry a large centred disc in a contrasting colour so the
# spatial centroid of change is well defined even for centre-out effects.
# --------------------------------------------------------------------------
GRID_COLS = 16
GRID_ROWS = 9


def _shape_deck(path, block, dur=DUR_MS):
    """One 2-slide deck whose 'entering' slide carries `block`.

    Returns the number of slides written so callers can assert nothing was
    silently dropped.
    """
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    w, h = prs.slide_width, prs.slide_height

    def cell_colour(i):
        # Identical on both slides: the grid is a STATIC reference frame, not
        # part of the change. An earlier version stepped the palette between
        # the two slides, which made every cell differ -- the diff map went
        # nearly uniform and the direction of the actual motion was averaged
        # away. (push measured as "stationary" because of it.) The only thing
        # that should differ between FROM and TO is the marker below.
        v = (i * 37) % 200 + 30
        return v

    for phase in (0, 1):                      # phase 0 = FROM, 1 = TO
        s = prs.slides.add_slide(blank)
        cw, ch = w // GRID_COLS, h // GRID_ROWS
        for r in range(GRID_ROWS):
            for c in range(GRID_COLS):
                i = r * GRID_COLS + c
                v = cell_colour(i)
                sh = s.shapes.add_shape(1, c * cw, r * ch, cw, ch)
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor(v, v, v)
                sh.line.fill.background()
                sh.shadow.inherit = False
        # ONE marker that moves and recolours between the two slides. Its
        # displacement makes a directional transition produce a clean, moving
        # diff centroid; its colour change guarantees a strong signal even for
        # effects that do not translate anything (fade, dissolve).
        d = min(w, h) // 3
        dx = 0 if phase == 0 else w // 3       # a known, sizeable shift
        disc = s.shapes.add_shape(9, (w - d) // 2 + dx, (h - d) // 2, d, d)
        disc.fill.solid()
        disc.fill.fore_color.rgb = RGBColor(230, 30, 30) if phase == 0 \
            else RGBColor(30, 30, 230)
        disc.line.fill.background()
        disc.shadow.inherit = False

    prs.save(path)

    buf = io.BytesIO(open(path, "rb").read())
    out = io.BytesIO()
    n = 0
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide2\.xml", item.filename):
                data = _slide_xml(data, block)
                n += 1
            zout.writestr(item, data)
    open(path, "wb").write(out.getvalue())
    return n


def _shape_deck2(path, block, dur=DUR_MS):
    """A SECOND probe deck for effects the first one cannot separate.

    Why a second deck exists at all (measured, not theorised): in deck 1 the
    only thing that differs between FROM and TO is a centred disc. Every
    effect that does not TRANSLATE therefore produces the same centre-weighted
    diff, and fade / wipe / dissolve / split / shape / randombar / box all came
    back with radial profiles within 0.03 of each other. They differ in HOW the
    change propagates, not WHERE it sits -- so the probe has to make
    propagation visible.

    Design: FROM and TO are both a fine RANDOM noise texture (different noise on
    each page, same average brightness). Now:
      * a sweeping reveal exposes the new noise region by region -> the diff
        forms a band that TRAVELS, and its direction is measurable;
      * a uniform cross-fade changes every pixel at once -> a flat diff, no
        travel;
      * a split opens from the centre outward -> the diff centroid moves
        outward in both directions.
    Average brightness is matched between the two pages on purpose: if TO were
    brighter overall, a fade would look like a brightness ramp and could be
    mistaken for a sweep.
    """
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.dml.color import RGBColor

    import random as _random

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    w, h = prs.slide_width, prs.slide_height
    cols, rows = 48, 27                       # 48x27 tiles = 1296 cells
    tw, th = w // cols, h // rows

    for phase in (0, 1):
        rng = _random.Random(1000 + phase)
        s = prs.slides.add_slide(blank)
        # Two brightness levels only, 50/50: matched mean, but the SPATIAL
        # pattern is what differs. Two levels keep the H.264 noise low.
        for r in range(rows):
            for c in range(cols):
                v = 190 if rng.random() < 0.5 else 60
                sh = s.shapes.add_shape(1, c * tw, r * th, tw, th)
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor(v, v, v)
                sh.line.fill.background()
                sh.shadow.inherit = False

    prs.save(path)

    buf = io.BytesIO(open(path, "rb").read())
    out = io.BytesIO()
    n = 0
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide2\.xml", item.filename):
                data = _slide_xml(data, block)
                n += 1
            zout.writestr(item, data)
    open(path, "wb").write(out.getvalue())
    return n


# Effects whose radial profile in deck 1 lands within 0.03 of another's, i.e.
# the first probe cannot tell them apart. Found by clustering the deck-1
# results; the list is data, not a guess.
DECK2_SPECS = ("fade", "wipe", "dissolve", "split", "shape", "randombar",
               "box", "push", "pan", "clock", "random", "cut")


def cmd_shapedeck2(out_dir, specs=None):
    """Deck 2: the noise-texture probe, for the effects deck 1 conflates."""
    os.makedirs(out_dir, exist_ok=True)
    want = set(specs or DECK2_SPECS)
    written = []
    for hyp in HYPOTHESES:
        if hyp["spec"] not in want:
            continue
        path = os.path.join(out_dir, hyp["spec"] + ".pptx")
        block = wrap_transition(hyp["family"], hyp["child"], dur=DUR_MS)
        n = _shape_deck2(path, block)
        if n != 1:
            raise RuntimeError("%s: transition not written (n=%d)"
                               % (hyp["spec"], n))
        written.append({"spec": hyp["spec"], "ui": hyp["ui"],
                        "en": hyp["en"], "family": hyp["family"]})
        json.dump({"case": "shape probe 2 (noise): %s" % hyp["ui"],
                   "spec": hyp["spec"], "ui": hyp["ui"],
                   "group": hyp["group"], "family": hyp["family"],
                   "child": hyp["child"], "dur_ms": DUR_MS, "fps": FPS,
                   "slides": 2, "transition_on": 2, "probe": "noise",
                   "expected_boundary_frames": [SLIDE_SECONDS * FPS]},
                  io.open(os.path.join(out_dir, hyp["spec"] + ".manifest.json"),
                          "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("shapedeck2 -> %s" % out_dir)
    print("  %d 份（%s）" % (len(written), "、".join(sorted(want))))
    return 0


def cmd_shapedeck(out_dir):
    """One deck per effect in HYPOTHESES, for the shape (not the name) probe.

    Default attributes only: one deck per spec. Variants (push dir=l vs r,
    wheel spokes=1 vs 4) are handled by cmd_dirdeck -- measuring the DEFAULT
    form first is what makes "what does each of the 47 do by default"
    answerable, and the direction variants are a separate, narrower question.
    """
    os.makedirs(out_dir, exist_ok=True)
    written = []
    for hyp in HYPOTHESES:
        name = hyp["spec"]
        path = os.path.join(out_dir, name + ".pptx")
        block = wrap_transition(hyp["family"], hyp["child"], dur=DUR_MS)
        n = _shape_deck(path, block)
        if n != 1:
            raise RuntimeError("%s: transition not written (n=%d)" % (name, n))
        written.append({"spec": name, "ui": hyp["ui"], "en": hyp["en"],
                        "group": hyp["group"], "family": hyp["family"],
                        "child": hyp["child"]})
        json.dump({"case": "shape probe: %s (%s)" % (hyp["ui"], hyp["en"]),
                   "spec": name, "ui": hyp["ui"], "group": hyp["group"],
                   "family": hyp["family"], "child": hyp["child"],
                   "dur_ms": DUR_MS, "fps": FPS, "slides": 2,
                   "transition_on": 2,
                   "expected_boundary_frames": [SLIDE_SECONDS * FPS],
                   "grid": [GRID_COLS, GRID_ROWS]},
                  io.open(os.path.join(out_dir, name + ".manifest.json"), "w",
                          encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump({"decks": written, "fps": FPS, "dur_ms": DUR_MS,
               "slide_seconds": SLIDE_SECONDS, "count": len(written)},
              io.open(os.path.join(out_dir, "_index.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("shapedeck -> %s" % out_dir)
    print("  %d 份 deck（HYPOTHESES 全量，默认属性）" % len(written))
    print("  每份 2 页：第 1 页 FROM 网格、第 2 页 TO 网格 + 该效果")
    return 0


# Which (spec, dir, axis) triples to render for the DIRECTION probe.
#
# Scope: the members of the two directional families in
# reference/transition-shapes.md §一 -- **整页平移** (push/pan/switch/
# page_curl/window) and **方向揭示** (19 members) -- MINUS the ones §四 of
# transition-choice.md forbids in business decks (checkerboard / honeycomb /
# window). Testing an effect nobody is allowed to use buys no decision value;
# this is a scope rule, not an oversight.
#
# Why this exists at all: transition-choice.md §三 carried an honest ⚠️ that
# the shape layer had only ever measured each effect in its DEFAULT form, so
# "dir flips the motion" was an assumption, not a measurement. This probe
# turns that assumption into data (or refutes it).
#
# ⚠️ TWO NOTATIONS, and they must never be mixed up again:
#   * `dir` here is the RAW OOXML ATTRIBUTE VALUE (`l`/`r`/`u`/`d`).
#   * `画面 x→y` is the OBSERVED MOTION, which is the OPPOSITE reading:
#     `dir=l` makes the picture travel r→l (l is the TARGET side the picture
#     moves toward, not the side it enters from). Measured on push:
#     dir=l → tx -1.004 (画面 r→l); dir=r → +1.009 (画面 l→r).
#   * An early version of this table annotated pairs as "push l→r" using the
#     SEMANTIC arrow for the ATTRIBUTE value -- literally backwards. The
#     `why` strings below now spell out the direction each pair is FOR.
#
# `axis` picks which trace mirror_verdict correlates: horizontal effects
# (default 画面 r→l / l→r) use "x"; vertical ones (comb/prestige 画面 b→t,
# airplane/crush/drape/fall_over 画面 t→b) use "y" and are probed with u/d --
# flipping an l/r attribute on a vertical effect tests nothing.
DIR_PROBE = (
    # ---- 整页平移族：水平轴 --------------------------------------------
    ("push", "l", "x", "默认值。画面 r→l"),
    ("push", "r", "x", "画面 l→r（§二 递进要的就是这个）"),
    ("pan", "l", "x", "同 push 族：默认 画面 r→l"),
    ("pan", "r", "x", "画面 l→r"),
    ("switch", "l", "x", "同 push 族：默认 画面 r→l"),
    ("switch", "r", "x", "画面 l→r"),
    ("page_curl", "l", "x", "卷曲带方向；默认 画面 r→l"),
    ("page_curl", "r", "x", "画面 l→r"),
    # ---- 方向揭示族：水平轴 --------------------------------------------
    ("wipe", "l", "x", "默认值。画面 r→l（形态层 band_travel −0.747）"),
    ("wipe", "r", "x", "画面 l→r（§二 并列首选）"),
    ("cover", "l", "x", "默认值。画面 r→l"),
    ("cover", "r", "x", "画面 l→r"),
    ("uncover", "l", "x", "默认值。画面 r→l"),
    ("uncover", "r", "x", "画面 l→r"),
    ("reveal", "l", "x", "§三 语义表标默认 r→l"),
    ("reveal", "r", "x", "画面 l→r（对照）"),
    ("box", "l", "x", "§三 语义表标默认 l→r，形态层却量到 stationary —— 必测"),
    ("box", "r", "x", "同上（对照）"),
    ("shape", "l", "x", "同 box：文档与形态层不一致"),
    ("shape", "r", "x", "同上（对照）"),
    ("wind", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 r→l ✓）"),
    ("wind", "r", "x", "画面 l→r（对照）"),
    ("glitter", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 r→l ✓）"),
    ("glitter", "r", "x", "画面 l→r（对照）"),
    ("peel_off", "l", "x", "§三 语义表标默认 r→l（形态层实测 画面 l→r，待校）"),
    ("peel_off", "r", "x", "画面 r→l（对照）"),
    ("curtains", "l", "x", "§三 声明『对称，加 dir 无方向』—— 验证是否真被忽略"),
    ("curtains", "r", "x", "同上：有了 l/r 才能配对"),
    # ---- 方向揭示族：垂直轴（默认就不是水平的）-------------------------
    ("comb", "u", "y", "默认 画面 b→t（形态层实测）—— 垂直轴，不能测 l/r"),
    ("comb", "d", "y", "画面 t→b"),
    ("prestige", "u", "y", "默认 画面 b→t（形态层实测）"),
    ("prestige", "d", "y", "画面 t→b"),
    ("airplane", "u", "y", "默认 画面 t→b（形态层实测）"),
    ("airplane", "d", "y", "画面 b→t"),
    ("crush", "u", "y", "默认 画面 t→b"),
    ("crush", "d", "y", "画面 b→t"),
    ("drape", "u", "y", "默认 画面 t→b"),
    ("drape", "d", "y", "画面 b→t"),
    ("fall_over", "u", "y", "默认 画面 t→b"),
    ("fall_over", "d", "y", "画面 b→t"),
    ("origami", "u", "y", "形态层实测 画面 t→b（§三 语义表把它归在 l→r 一组，待校）"),
    ("origami", "d", "y", "画面 b→t"),
)


def _dir_pairs(probe=None):
    """Group DIR_PROBE rows into {spec: {axis, dirs, why}} for pairing."""
    out = {}
    for spec, d, axis, why in (probe or DIR_PROBE):
        e = out.setdefault(spec, {"axis": axis, "dirs": [], "why": []})
        e["dirs"].append(d)
        e["why"].append("%s: %s" % (d, why))
    return out


# ---------------------------------------------------------------------------
# ATTR_PROBE: the "attribute value set" the mechanism layer never covered.
#
# HANDOVER 3a left this explicitly open: the name layer records ONE child per
# effect (clock -> <p:wheel spokes="1"/>), but several effects take FURTHER
# attributes, and nobody had measured what those values actually DO. The
# PowerPoint-written presetID enum (scripts/transition_reference.json,
# enum_by_child) turns out to enumerate every variant PowerPoint itself emits
# in its UI, so it tells us the **domain** of each attribute exactly -- we do
# not have to guess which values are legal.
#
# Each row is (spec, label, overrides, why):
#   * `spec`  -- the effect's spec in HYPOTHESES (keeps the UI name + family)
#   * `label` -- short tag used in the filename, e.g. "spokes1"
#   * `overrides` -- (attr, value) pairs to FORCE onto the child element.
#                    value=None means "strip the attribute" (the default form).
#   * `why`   -- what decision this variant informs.
ATTR_PROBE = (
    # -- wheel spokes: interface「时钟」+「转盘」。enum lists 1/2/3/8, and
    #    `<p:wheel/>` (no attr) is a SEPARATE presetID, so 4 is the default.
    ("clock", "spokes1", (("spokes", "1"),), "名字层记的形式；一条线扫一圈"),
    ("clock", "spokes2", (("spokes", "2"),), "界面「转盘 2 根」—— 真的是两根吗"),
    ("clock", "spokes3", (("spokes", "3"),), "界面「转盘 3 根」"),
    ("clock", "spokes4", (("spokes", "4"),), "显式写 4 —— 是否等于不写（默认）"),
    ("clock", "spokes8", (("spokes", "8"),), "界面「转盘 8 根」"),
    ("clock", "plain", (("spokes", None),), "不写 spokes —— 默认到底几根"),
    ("clock", "spokes6", (("spokes", "6"),), "enum 里没有的值 —— 合法吗、等同什么"),
    # -- split orient: horz/vert 在 enum 里是两个 presetID，必测
    ("split", "horz", (("orient", "horz"), ("dir", "out")),
     "名字层记的默认形式"),
    ("split", "vert", (("orient", "vert"), ("dir", "out")),
     "竖着开 —— 与横着开是否只是转了 90°"),
    ("split", "in_horz", (("orient", "horz"), ("dir", "in")),
     "反向：从边缘合拢"),
    ("split", "in_vert", (("orient", "vert"), ("dir", "in")),
     "竖着合拢"),
    # -- comb 的 dir：enum 里 comb 的第二个 presetID 是 dir="vert"，
    #    所以**不写 dir 才是默认**，必须把默认单独烘一份才能比。
    ("comb", "plain", (("dir", None),), "默认（不写 dir）—— 基线"),
    ("comb", "horz", (("dir", "horz"),), "显式 horz —— 是否等于默认"),
    ("comb", "vert", (("dir", "vert"),), "竖向梳理（enum 记的第二个值）"),
    # -- glitter pattern
    ("glitter", "default", (("pattern", None),), "不写 pattern 的默认花纹"),
    ("glitter", "hexagon", (("pattern", "hexagon"),), "菱形花纹是否只是换粒子形状"),
    # -- prism isContent / isInverted（名字层 §3 已记 cube/rotate/orbit 三项，
    #    但没点明是哪两个位；第四个组合只有 XML）
    ("cube", "plain", (("isContent", None), ("isInverted", None)),
     "裸 <p14:prism/> —— 界面「立方体」(3914)"),
    ("cube", "content", (("isContent", "1"),),
     "isContent=1 —— 界面「旋转」(3918)"),
    ("cube", "inverted", (("isInverted", "1"),),
     "isInverted=1 单独置位 —— 界面上没有这一项 (3922)"),
    ("cube", "both", (("isContent", "1"), ("isInverted", "1")),
     "两个都写 —— 界面「轨道」(3926)"),
    # -- p15 prstTrans invX：与 dir 完全不同的一个反向属性
    ("wind", "plain", (("invX", None),), "默认方向"),
    ("wind", "invX", (("invX", "1"),), "invX=1 是不是横向镜像（与 dir 有别）"),
    ("peel_off", "plain", (("invX", None),), "默认方向"),
    ("peel_off", "invX", (("invX", "1"),), "invX=1 的反向"),
    ("fall_over", "plain", (("invX", None),), "默认方向"),
    ("fall_over", "invX", (("invX", "1"),), "invX=1 的反向"),
)


# --------------------------------------------------------------------------
# TIMING_PROBE: <p:transition> and <p:timing> on the SAME slide.
#
# The mechanism layer asserts, at three separate places, that the two are
# orthogonal: "两者正交，互不占用，可以同时存在于一页" (transition-model.md §一/§四).
# That assertion has never been measured. It is exactly the shape of claim this
# project refuses to accept unmeasured -- and it is the load-bearing one, because
# the whole "two mechanisms" architecture rests on it.
#
# Three things have to be told apart, and they need DIFFERENT instruments:
#
#   (a) STRUCTURAL  -- does PowerPoint keep BOTH after a round-trip, or does one
#       silently eat the other?  Same failure class as §12 (misplaced child is
#       dropped on save) -- invisible in the file you wrote.
#   (b) TIMELINE    -- when the animation fires, does it land in the same frame
#       window as the transition, or after it?  A per-frame energy trace answers
#       this; the shape metrics (direction/mode) do not -- they describe HOW the
#       change looks, not WHEN it happens.
#   (c) INTERFERENCE-- with both present, does either one change shape vs. its
#       solo rendering?  Only a pixel diff against the solo control can say.
#
# So every row below comes in a PAIR: the combination, and its solo control.
# Row layout: (spec, tspec, label, why)
#   tspec is the transition written on the same slide, or None for the control.
# --------------------------------------------------------------------------
# The animation is deliberately the SAME everywhere: a fade on one shape,
# 0.5 s, trigger=with. `with` is the interesting trigger -- an `after` animation
# waits for a click and CreateVideo would never reach it, which would look like
# "the animation was eaten" when nothing of the sort happened.
TIMING_ANIM = {"effect": "fade", "duration": 0.5, "trigger": "with"}

TIMING_PROBE = (
    # ---- the pair that answers (a) and (b) -------------------------------
    ("push", "push", "both_push",
     "同页写 push 切换 + fade 动画 —— 两者是否都在、动画何时播"),
    ("push", None, "solo_anim",
     "只有动画、没有切换 —— 对照组：动画本身的帧窗口"),
    ("push", "push_only", "solo_trans",
     "只有 push 切换、没有动画 —— 同样是 push，排除切换种类这个变量"),
    # ---- a second transition family, to check the answer is not push-specific
    ("wipe", "wipe", "both_wipe",
     "换成方向揭示族 —— 结论是否只对 push 成立"),
    ("wipe", None, "solo_anim_wipe", "对照组：动画本身（wipe 组）"),
    ("wipe", "wipe_only", "solo_trans_wipe", "对照组：wipe 本身（同样是 wipe）"),
    # ---- no transition vs. a SLOW one: does transition duration delay the
    #      animation?  If the animation waited on the transition we would see
    #      its window slide right as the transition lengthens.  The transition
    #      is held FIXED (fade) across the three rows so only duration varies.
    ("push", "fade", "both_fade_short",
     "短切换（0.3s fade）+ 动画 —— 动画是否被切换推后"),
    ("push", "fade_slow", "both_fade_long",
     "长切换（1.5s fade）+ 动画 —— 动画窗口是否跟着右移"),
    ("push", "fade_only", "solo_fade_800",
     "fade(默认 0.8s) 单独 —— fade 系列自己的基线，供上面两行比"),
)


def _set_attrs(child, overrides):
    """Force (attr, value) pairs onto a child element.

    `value=None` REMOVES the attribute (so the row measures the default form).
    Removal matters: an early version of the dir probe only ever ADDED, so
    "default" rows silently carried whatever the HYPOTHESES child already had.
    """
    m = re.match(r"<([A-Za-z0-9:]+)((?:\s[^>]*)?)/>$", child)
    if not m:
        raise ValueError("cannot set attrs on %r" % child)
    tag, attrs = m.group(1), m.group(2) or ""
    for attr, val in overrides:
        attrs = re.sub(r'\s+' + re.escape(attr) + r'="[^"]*"', "", attrs)
        if val is not None:
            attrs += ' %s="%s"' % (attr, val)
    return "<%s%s/>" % (tag, attrs)


def _probe_deck(out_dir, rows, tag_key, label_of):
    """Shared writer for dirdeck / attrdeck: one deck per probe row."""
    os.makedirs(out_dir, exist_ok=True)
    by_spec = {h["spec"]: h for h in HYPOTHESES}
    written = []
    for spec, label, overrides, why in rows:
        hyp = by_spec.get(spec)
        if hyp is None:
            raise RuntimeError("probe references unknown spec %r" % spec)
        name = "%s_%s_%s" % (spec, tag_key, label)
        path = os.path.join(out_dir, name + ".pptx")
        child = _set_attrs(hyp["child"], overrides)
        block = wrap_transition(hyp["family"], child, dur=DUR_MS)
        n = _shape_deck(path, block)
        if n != 1:
            raise RuntimeError("%s: transition not written (n=%d)" % (name, n))
        entry = {"spec": spec, tag_key: label, "label": label,
                 "ui": hyp["ui"], "group": hyp["group"],
                 "family": hyp["family"], "child": child, "why": why,
                 "deck": name}
        written.append(entry)
        json.dump({"case": "%s probe: %s %s=%s" % (tag_key, spec, tag_key,
                                                   label),
                   "spec": spec, tag_key: label, "label": label,
                   "ui": hyp["ui"], "group": hyp["group"],
                   "family": hyp["family"], "child": child,
                   "dur_ms": DUR_MS, "fps": FPS, "slides": 2,
                   "transition_on": 2,
                   "expected_boundary_frames": [SLIDE_SECONDS * FPS],
                   "grid": [GRID_COLS, GRID_ROWS]},
                  io.open(os.path.join(out_dir, name + ".manifest.json"),
                          "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump({"decks": written, "fps": FPS, "dur_ms": DUR_MS,
               "slide_seconds": SLIDE_SECONDS, "count": len(written)},
              io.open(os.path.join(out_dir, "_index.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    return written


def cmd_attrdeck(out_dir):
    """One deck per ATTR_PROBE row -- the "attribute value set" probe."""
    rows = [(s, lab, ov, why) for (s, lab, ov, why) in ATTR_PROBE]
    written = _probe_deck(out_dir, rows, "attr", lambda e: e["label"])
    print("attrdeck -> %s" % out_dir)
    print("  %d 份 deck（ATTR_PROBE：属性取值全集）" % len(written))
    for e in written:
        print("    %-22s %-46s %s" % (e["deck"], e["child"], e["why"]))
    return 0


def _anim_deck(path, anim, trans_block=None):
    """One deck whose 'entering' slide carries a REAL <p:timing> animation and
    optionally a <p:transition>.

    The timing block is generated by motion.py's own build_timing(), not
    hand-written here. That matters: the question is whether the SHIPPING
    engine's output coexists with a transition, and a hand-rolled timing block
    would answer a question nobody asked.

    The animation targets whatever shape the probe deck already has, by
    position, so the deck geometry stays the shape-probe geometry (the marker
    disc is what moves -- a static grid would make a fade invisible).

    Returns (n_timing, n_transition): how many slides got each block, so the
    caller can assert nothing was silently dropped.
    """
    from pptx import Presentation
    from pptx.util import Emu
    from pptx.dml.color import RGBColor
    import motion as _motion

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    w, h = prs.slide_width, prs.slide_height

    def cell_colour(i):
        v = (i * 37) % 200 + 30
        return v

    for phase in (0, 1):
        s = prs.slides.add_slide(blank)
        cw, ch = w // GRID_COLS, h // GRID_ROWS
        for r in range(GRID_ROWS):
            for c in range(GRID_COLS):
                i = r * GRID_COLS + c
                v = cell_colour(i)
                sh = s.shapes.add_shape(1, c * cw, r * ch, cw, ch)
                sh.fill.solid()
                sh.fill.fore_color.rgb = RGBColor(v, v, v)
                sh.line.fill.background()
                sh.shadow.inherit = False
        d = min(w, h) // 3
        dx = 0 if phase == 0 else w // 3
        disc = s.shapes.add_shape(9, (w - d) // 2 + dx, (h - d) // 2, d, d)
        disc.fill.solid()
        disc.fill.fore_color.rgb = RGBColor(230, 30, 30) if phase == 0 \
            else RGBColor(30, 30, 230)
        disc.line.fill.background()
        disc.shadow.inherit = False

    prs.save(path)

    # Find the moving marker's shape id on slide 2, so the animation has a real
    # target (motion.py resolves targets by name/id from the spTree).
    buf = io.BytesIO(open(path, "rb").read())
    with zipfile.ZipFile(buf) as z:
        s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
    ids = re.findall(r'<p:cNvPr id="(\d+)"[^>]*name="([^"]*)"', s2)
    if not ids:
        raise RuntimeError("no shapes found on slide 2 of %s" % path)
    # The LAST shape is the disc (grid is written first).
    marker_id = ids[-1][0]

    timing_block = ""
    if anim is not None:
        catalog = _load_json(os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                          "motion_catalog.json"))
        meta = catalog.get(str(anim["effect"]))
        if meta is None:
            raise RuntimeError("unknown anim effect %r" % anim["effect"])
        entries = [{"shape_id": marker_id, "tpl": meta["template"],
                    "effect": anim, "delay": anim.get("delay") or 0.0}]
        timing_block = _motion.build_timing(entries, _motion.NodeIds(4))

    n_t, n_x = 0, 0
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if re.fullmatch(r"ppt/slides/slide2\.xml", item.filename):
                xml = data.decode("utf-8")
                # Declare extension prefixes on <p:sld> (only the missing ones:
                # a duplicate makes the part malformed, com-pitfalls §26/§12).
                m = re.search(r"<p:sld(\s[^>]*?)?>", xml)
                have = set(re.findall(r'xmlns:(\w+)=', m.group(1) or ""))
                add = " ".join('xmlns:%s="%s"' % (k, v)
                               for k, v in sorted(NS.items()) if k not in have)
                if add:
                    xml = re.sub(r"<p:sld(\s[^>]*?)?>",
                                 lambda mm: "%s %s>" % (mm.group(0)[:-1].rstrip(), add),
                                 xml, count=1)
                # ORDER: cSld, clrMapOvr?, transition?, timing?, extLst?
                # Insert by walking the anchor forward, never by re-deriving an
                # index from string searches -- that is exactly how a timing
                # block ends up INSIDE the transition (com-pitfalls §12/§19).
                anchor = "</p:clrMapOvr>"
                idx = xml.rindex(anchor) + len(anchor)
                pieces = []
                if trans_block:
                    pieces.append(trans_block)
                    n_x += 1
                if timing_block:
                    pieces.append(timing_block)
                    n_t += 1
                if pieces:
                    xml = xml[:idx] + "".join(pieces) + xml[idx:]
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    open(path, "wb").write(out.getvalue())
    return n_t, n_x


def cmd_timingdeck(out_dir):
    """One deck per TIMING_PROBE row: transition + timing on the same slide.

    Each row that names a transition is paired with a solo control, so the
    analyser can always answer "did the combination change either one?"
    without a baseline it invented.
    """
    os.makedirs(out_dir, exist_ok=True)
    by_spec = {h["spec"]: h for h in HYPOTHESES}
    # transition written by the solo/combination rows. `none` = the row has a
    # transition but it is the plain fade (so "solo_trans" means "transition
    # only"); None = no transition block at all.
    written = []
    for spec, tspec, label, why in TIMING_PROBE:
        hyp = by_spec.get(spec)
        if hyp is None:
            raise RuntimeError("TIMING_PROBE references unknown spec %r" % spec)
        name = "timing_%s" % label
        path = os.path.join(out_dir, name + ".pptx")

        trans_block = None
        trans_child = None
        want_trans = tspec is not None
        # The control rows MUST use the same transition element as their
        # combination row, otherwise the comparison moves two variables at once
        # (an earlier version used <p:fade/> for "solo_trans" and <p:push/> for
        # "both_push", which made the control useless -- the low fade energy
        # looked like "the transition vanished", when it was simply a different
        # transition).
        trans_fam = "core"
        if tspec == "push":
            trans_child = hyp["child"]
        elif tspec == "push_only":
            trans_child = by_spec["push"]["child"]
        elif tspec == "wipe":
            trans_child = by_spec["wipe"]["child"]
        elif tspec == "wipe_only":
            trans_child = by_spec["wipe"]["child"]
        elif tspec in ("fade", "fade_slow", "fade_only"):
            trans_child = "<p:fade/>"
        if want_trans:
            if tspec == "wipe":
                trans_fam = by_spec["wipe"]["family"]
            elif tspec == "wipe_only":
                trans_fam = by_spec["wipe"]["family"]
            dur = {"fade": 300, "fade_slow": 1500, "fade_only": DUR_MS}.get(
                tspec, DUR_MS)
            trans_block = wrap_transition(trans_fam, trans_child, dur=dur)

        anim = TIMING_ANIM if not str(label).startswith("solo_trans") \
            and label != "solo_fade_800" else None
        n_t, n_x = _anim_deck(path, anim, trans_block)
        if want_trans and n_x != 1:
            raise RuntimeError("%s: transition not written (n=%d)" % (name, n_x))
        if anim is not None and n_t != 1:
            raise RuntimeError("%s: timing not written (n=%d)" % (name, n_t))

        entry = {"spec": spec, "tspec": tspec, "label": label, "ui": hyp["ui"],
                 "group": hyp["group"], "family": hyp["family"],
                 "has_transition": bool(want_trans), "has_timing": anim is not None,
                 "trans_child": trans_child, "anim": anim, "why": why,
                 "deck": name}
        written.append(entry)
        json.dump({
            "case": "timing probe: %s" % label, "spec": spec, "label": label,
            "tspec": tspec, "ui": hyp["ui"], "group": hyp["group"],
            "family": hyp["family"],
            "has_transition": bool(want_trans), "has_timing": anim is not None,
            "trans_child": trans_child, "anim": anim, "dur_ms": DUR_MS,
            "fps": FPS, "slides": 2, "transition_on": 2,
            "expected_boundary_frames": [SLIDE_SECONDS * FPS],
            "grid": [GRID_COLS, GRID_ROWS],
        }, io.open(os.path.join(out_dir, name + ".manifest.json"), "w",
                   encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump({"decks": written, "fps": FPS, "dur_ms": DUR_MS,
               "slide_seconds": SLIDE_SECONDS, "count": len(written)},
              io.open(os.path.join(out_dir, "_index.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("timingdeck -> %s" % out_dir)
    print("  %d 份 deck（TIMING_PROBE：切换 × 页内动画同页）" % len(written))
    for e in written:
        print("    %-24s trans=%-5s anim=%-5s %s" % (
            e["deck"], e["has_transition"], e["has_timing"], e["why"]))
    return 0


def cmd_dirdeck(out_dir):
    """One deck per (spec, dir) row in DIR_PROBE.
    Filenames are `<spec>_dir_<d>.pptx` so a plain directory listing shows
    which pairs exist, and so the analyser can pair `x_dir_l` with `x_dir_r`
    by filename without consulting the manifest.
    """
    os.makedirs(out_dir, exist_ok=True)
    by_spec = {h["spec"]: h for h in HYPOTHESES}
    written = []
    for spec, direction, axis, why in DIR_PROBE:
        hyp = by_spec.get(spec)
        if hyp is None:
            raise RuntimeError("DIR_PROBE references unknown spec %r" % spec)
        name = "%s_dir_%s" % (spec, direction)
        path = os.path.join(out_dir, name + ".pptx")
        child = _with_dir(hyp["child"], direction)
        block = wrap_transition(hyp["family"], child, dur=DUR_MS)
        n = _shape_deck(path, block)
        if n != 1:
            raise RuntimeError("%s: transition not written (n=%d)" % (name, n))
        entry = {"spec": spec, "dir": direction, "axis": axis, "ui": hyp["ui"],
                 "group": hyp["group"], "family": hyp["family"],
                 "child": child, "why": why, "deck": name}
        written.append(entry)
        json.dump({"case": "dir probe: %s dir=%s" % (spec, direction),
                   "spec": spec, "dir": direction, "axis": axis,
                   "ui": hyp["ui"],
                   "group": hyp["group"], "family": hyp["family"],
                   "child": child, "dur_ms": DUR_MS, "fps": FPS, "slides": 2,
                   "transition_on": 2,
                   "expected_boundary_frames": [SLIDE_SECONDS * FPS],
                   "grid": [GRID_COLS, GRID_ROWS]},
                  io.open(os.path.join(out_dir, name + ".manifest.json"),
                          "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump({"decks": written, "fps": FPS, "dur_ms": DUR_MS,
               "slide_seconds": SLIDE_SECONDS, "count": len(written)},
              io.open(os.path.join(out_dir, "_index.json"), "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("dirdeck -> %s" % out_dir)
    print("  %d 份 deck（DIR_PROBE：选择层真正用到的 dir 组合）" % len(written))
    for e in written:
        print("    %-14s %s" % (e["deck"], e["why"]))
    return 0


def _profile_metrics(frames):
    """Recover the SPATIAL story from a rendered transition.

    Whole-frame means are not enough (transition-model.md §七): a directional
    reveal moves an edge, not the mean. So we work with a coarse block grid
    and ask, per frame, where the change is concentrated and how that centre
    travels.
    """
    import cv2
    import numpy as np

    if len(frames) < 2:
        return None
    # Work at block resolution: 16x9 blocks, same as the probe grid.
    small = [cv2.resize(cv2.cvtColor(f, cv2.COLOR_BGR2GRAY),
                        (GRID_COLS * 4, GRID_ROWS * 4),
                        interpolation=cv2.INTER_AREA).astype(np.float32)
             for f in frames]
    diffs = [np.abs(small[i] - small[i - 1]) for i in range(1, len(small))]
    energy = np.array([float(d.sum()) for d in diffs])

    # Active window. Two traps here, both met in practice:
    #
    # 1. A relative-only threshold. H.264 leaves isolated noise spikes on
    #    otherwise identical frames (measured on a still grid: energy 84-170
    #    on single frames), while the real transition of a small-area effect
    #    can peak at ~800. 10% of 823 is 82 -- right inside the noise band,
    #    which is how wipe/fade first reported a 120-frame window covering the
    #    whole dwell.
    # 2. The noise floor from median+MAD is USELESS when most frames are
    #    pixel-identical (median=0, MAD=0 -> floor=0), which is exactly the
    #    case for a flat probe deck.
    #
    # So: threshold on prominence, then require CONTIGUITY. A transition is a
    # contiguous run of frames; noise is isolated single frames. Keep only runs
    # of >= 3 frames, and pick the strongest.
    thr = max(energy.max() * 0.10, 1e-6)
    hot = np.where(energy > thr)[0]
    runs = []
    if hot.size:
        s = p = int(hot[0])
        for i in hot[1:]:
            i = int(i)
            if i <= p + 2:                 # bridge 1-frame gaps inside a run
                p = i
            else:
                runs.append((s, p))
                s = p = i
        runs.append((s, p))
    runs = [(s, e) for s, e in runs if e - s >= 2]     # >= 3 frames
    if not runs:
        # No run long enough to be a motion. If there is still a strong single-
        # frame spike, that is a HARD CUT -- one frame, real change, no motion
        # to describe. Reporting a bare "none" would conflate it with "nothing
        # happened at all", which is a different thing (measured: cut fires
        # exactly one 33 ms frame at energy 7601, versus a still deck's 0).
        if hot.size and energy.max() > 1e-6:
            return {"direction": "none", "symmetry": "none", "mode": "cut",
                    "energy": energy.tolist(), "peak_frame": int(hot[0]),
                    "span": 0, "dx": 0.0, "dy": 0.0, "translate": "none",
                    "tx": 0.0, "ty": 0.0, "symmetry_v": 0.0, "symmetry_h": 0.0,
                    "rot_span": 0, "rot_mono": 0.0, "is_rotation": False,
                    "ang_trace": [], "cx_trace": [], "cy_trace": [],
                    "radial_profile": [], "coverage": 0.0,
                    "peak_to_mean": 0.0, "band_travel": 0.0,
                    "active_frames": int(hot.size)}
        return {"direction": "none", "symmetry": "none", "mode": "none",
                "energy": energy.tolist(), "peak_frame": 0, "span": 0,
                "dx": 0.0, "dy": 0.0, "translate": "none",
                "symmetry_v": 0.0, "symmetry_h": 0.0,
                "rot_span": 0, "rot_mono": 0.0, "is_rotation": False,
                "ang_trace": [], "cx_trace": [], "cy_trace": [],
                "radial_profile": [], "active_frames": 0}
    # Strongest run by integrated energy.
    a, b = max(runs, key=lambda r: energy[r[0]:r[1] + 1].sum())
    active = np.arange(a, b + 1)

    # Dense frames = the ones where the reveal front is actually travelling.
    dense = [i for i in active if energy[i] > energy.max() * 0.25]
    sub = [diffs[i] for i in dense]
    if not sub:
        sub = [diffs[i] for i in active]
    mean_map = np.mean(sub, axis=0)
    H, W = mean_map.shape

    ys, xs = np.mgrid[0:H, 0:W]
    tot = mean_map.sum() + 1e-9
    cx = float((xs * mean_map).sum() / tot) / (W - 1)
    cy = float((ys * mean_map).sum() / tot) / (H - 1)

    # --- direction: does the change centre drift across the window? --------
    cxs, cys = [], []
    for i in dense:
        m = diffs[i]
        t = m.sum() + 1e-9
        cxs.append(float((xs * m).sum() / t) / (W - 1))
        cys.append(float((ys * m).sum() / t) / (H - 1))
    if len(cxs) >= 3:
        dx = cxs[-1] - cxs[0]
        dy = cys[-1] - cys[0]
    else:
        dx = dy = 0.0
    if max(abs(dx), abs(dy)) < 0.12:
        direction = "stationary"          # change happens but does not travel
    elif abs(dx) >= abs(dy):
        direction = "l->r" if dx > 0 else "r->l"
    else:
        direction = "t->b" if dy > 0 else "b->t"

    # --- symmetry: mirror residual about vertical / horizontal axes --------
    def mirror_resid(m):
        fv = m[:, ::-1]
        fh = m[::-1, :]
        den = m.sum() + 1e-9
        return (np.abs(m - fv).sum() / den, np.abs(m - fh).sum() / den)
    rv, rh = mirror_resid(mean_map)

    # --- mode: linear sweep vs radial vs rotating vs uniform --------------
    # Radial test: energy ringed around the centre at constant radius.
    rr = np.sqrt(((xs - (W - 1) / 2.0) / (W / 2.0)) ** 2 +
                 ((ys - (H - 1) / 2.0) / (H / 2.0)) ** 2)
    r_bins = np.linspace(0, 1.4, 8)
    radial_profile = []
    for k in range(len(r_bins) - 1):
        m_sel = (rr >= r_bins[k]) & (rr < r_bins[k + 1])
        radial_profile.append(float(mean_map[m_sel].mean())
                              if m_sel.any() else 0.0)
    rp = np.array(radial_profile)
    rp_n = rp / (rp.max() + 1e-9)

    # How many frames actually carried the change. One frame is a HARD CUT --
    # there is no motion to describe, and calling it "centre-out" (as an early
    # version did) invents a shape that does not exist. `cut` measured exactly
    # one 33 ms frame; so does nothing else.
    n_active = int(active.size)
    if b - a <= 1 or n_active <= 1:
        mode = "cut"
    elif direction != "stationary" and abs(dx) + abs(dy) > 0.35:
        mode = "sweep"
    elif rp_n[1:4].mean() > rp_n[4:].mean() * 1.5:
        mode = "centre-out"
    elif float(rp_n.std()) < 0.10:
        mode = "uniform"
    else:
        mode = "radial/other"

    # --- translation: accumulated per-frame displacement --------------------
    # The centroid method above cannot see a WHOLE-PAGE translation (push, pan,
    # switch, peel): every column changes equally, so the change centroid never
    # moves. Measured: push reported "stationary" while its column bands were
    # visibly scrambling between frames 65 and 80.
    #
    # Phase correlation between the two ENDPOINT frames does not help either --
    # for a slide-in both endpoints are mid-motion frames that look alike, so
    # the net shift reads ~0. What works is to accumulate the per-frame shift
    # ACROSS the window: consecutive frames give a clean signal (measured on
    # push: a steady -1 to -5 px per frame, response ~1.0), and summing them
    # recovers the total travel and its sign.
    tot_x = tot_y = 0.0
    n_pairs = 0
    for i in range(a, min(b + 1, len(small) - 1)):
        try:
            (sx, sy), resp = cv2.phaseCorrelate(
                np.float32(small[i]), np.float32(small[i + 1]))
        except Exception:
            continue
        tot_x += float(sx)
        tot_y += float(sy)
        n_pairs += 1
    tx = tot_x / max(1, W - 1)
    ty = tot_y / max(1, H - 1)
    shift_mag = (tx ** 2 + ty ** 2) ** 0.5
    if shift_mag > 0.15:
        if abs(tx) >= abs(ty):
            translate = "l->r" if tx > 0 else "r->l"
        else:
            translate = "t->b" if ty > 0 else "b->t"
    else:
        translate = "none"

    # --- localisation: how SPREAD OUT is the diff, and does it travel? ------
    # This is the metric deck 2 exists for. A uniform cross-fade lights up the
    # whole frame at once (high coverage, low peak-to-mean). A sweeping reveal
    # lights up a moving band (varying coverage, high peak-to-mean on a coarse
    # grid). `band_travel` is the signed drift of the diff's column profile,
    # which is exactly "which way is the reveal front going".
    mean_map_norm = mean_map / (mean_map.max() + 1e-9)
    coverage = float((mean_map_norm > 0.35).mean())
    p2m = float(mean_map_norm.mean()) or 1e-9
    peak_to_mean = float(1.0 / p2m) if p2m else 0.0
    # Column profile of the mean diff, as a fraction of frame width, and its
    # drift from the first dense frame to the last. A front moving left->right
    # shifts the column centroid the same way.
    colprof = mean_map.sum(axis=0)
    band_travel = 0.0
    if len(dense) >= 3:
        def colcen(m):
            c = m.sum(axis=0)
            idx = np.arange(len(c))
            return float((idx * c).sum() / (c.sum() + 1e-9)) / max(1, len(c) - 1)
        band_travel = colcen(diffs[dense[-1]]) - colcen(diffs[dense[0]])

    # --- rotation: does the diff sweep around the centre? -------------------
    # 前三类指标都看不见旋转：一圈扫下来重心回原点（漂移≈0）、不是整页平移
    # （位移≈0）、径向剖面又和中心扩散几乎一样 —— 结果 clock 被就近塞进
    # 「中心扩散」，方向还错记成 t->b。这个指标是补上那条缝。
    #
    # 做法：绕中心切 24 个 15° 扇区，逐帧取"当前变化最集中的扇区"，
    # 得到一条角度序列。扫一圈的效果会表现为角度【单调推进】（顺时针递减）。
    # 判据用"覆盖了多少个不同扇区"+"轨迹的单调性"，避免把随机抖动认成旋转。
    ang_trace = []
    if len(dense) >= 4:
        yy, xx = np.indices((H, W))
        ang = (np.degrees(np.arctan2(-(yy - (H / 2.0)), xx - (W / 2.0)))
               + 360.0) % 360.0
        sec_id = (ang // 15.0).astype(np.int32)          # 0..23
        for di in dense:
            dm = diffs[di]
            tot = float(dm.sum())
            if tot <= 1e-9:
                continue
            per = np.bincount(sec_id.ravel(), weights=dm.ravel(), minlength=24)
            ang_trace.append(int(per.argmax()) * 15)

    rot_span = 0        # 覆盖过多少个不同扇区
    rot_mono = 0.0      # 顺时针推进的一致性（0..1）
    if len(ang_trace) >= 4:
        rot_span = len(set(ang_trace))
        # 相邻两步的"顺时针增量"（角度递减），归一化到 [-0.5, 0.5]
        steps = []
        for a, b in zip(ang_trace, ang_trace[1:]):
            d = ((a - b + 180) % 360) - 180           # 正数=顺时针
            if abs(d) <= 90:                          # 忽略跳变（扇区换错峰）
                steps.append(d)
        if steps:
            pos = sum(1 for s in steps if s > 0)
            rot_mono = float(pos) / len(steps)

    # 旋转的判据：扫过 ≥12 个扇区（半圈以上）且顺时针一致 ≥0.7。
    # 门槛取这么高是因为"局部纹理变化"也会零星点亮不同扇区，但那不单调。
    is_rotation = (rot_span >= 12 and rot_mono >= 0.7)

    return {"direction": direction, "dx": round(dx, 3), "dy": round(dy, 3),
            "translate": translate, "tx": round(tx, 3), "ty": round(ty, 3),
            "rot_span": rot_span, "rot_mono": round(rot_mono, 3),
            "is_rotation": is_rotation, "ang_trace": ang_trace,
            "symmetry_v": round(float(rv), 3), "symmetry_h": round(float(rh), 3),
            "mode": mode, "radial_profile": [round(x, 3) for x in rp_n],
            "coverage": round(coverage, 3),
            "peak_to_mean": round(peak_to_mean, 2),
            "band_travel": round(band_travel, 3),
            # Per-frame change centroid over the DENSE window. Exposed so the
            # mirror test can compare two runs frame-by-frame instead of
            # comparing two summary numbers (a summary can agree while the
            # trajectories disagree, and vice versa).
            "cx_trace": [round(v, 4) for v in cxs],
            "cy_trace": [round(v, 4) for v in cys],
            "peak_frame": a, "span": b - a, "active_frames": n_active,
            "energy": energy.tolist()}


def mirror_verdict(a_res, b_res, axis="x"):
    """Decide whether two runs of the same effect are MIRROR IMAGES.

    The question transition-choice.md §三 could not answer: if you take
    `<p:push dir="l"/>` and flip it to `dir="r"`, do you get the same motion
    backwards, or something else entirely?

    TWO COMPLEMENTARY INSTRUMENTS, because each is blind to one family:

    * `cx_trace` -- the per-frame CHANGE CENTROID. Sees a travelling reveal
      front (wipe, cover, uncover). BLIND to whole-page translation: every
      column changes equally so the centroid never moves. Measured on push:
      dx = 0.011 (i.e. "stationary") while the page demonstrably slid.
    * `tx` -- the ACCUMULATED per-frame shift from phase correlation. Sees
      whole-page translation (push) in one number. Blind to a sweep: a wipe
      front crossing the frame moves very little total mass.

    So we run both and take whichever one actually has signal. This is the
    §八 lesson restated: every instrument has a family it cannot see, and the
    failure is silent -- push came back "stationary / ambiguous" purely
    because the first instrument was the wrong one, not because push is odd.

    A true mirror means: one run goes the other way at the corresponding
    moment. We accept it if EITHER instrument shows that, and we record which.
    """
    import numpy as np

    out = {"axis": axis,
           "dir_a": a_res.get("direction"), "dir_b": b_res.get("direction"),
           "translate_a": a_res.get("translate"),
           "translate_b": b_res.get("translate")}

    # ---- instrument 2: accumulated shift (works for whole-page moves) ----
    # `tx`/`ty` are already normalised by frame width/height, so their
    # magnitude is comparable across effects. Signal threshold mirrors the one
    # `_profile_metrics` itself uses to call something a translation (0.15).
    txa, txb = a_res.get("tx") or 0.0, b_res.get("tx") or 0.0
    tya, tyb = a_res.get("ty") or 0.0, b_res.get("ty") or 0.0
    shift_mag = max((txa ** 2 + tya ** 2) ** 0.5, (txb ** 2 + tyb ** 2) ** 0.5)
    out["tx_a"], out["tx_b"] = txa, txb
    out["shift_mag"] = round(shift_mag, 3)
    by_shift = None
    if shift_mag > 0.15:
        # Opposite signs = the two runs travelled opposite ways. Ratio near 1
        # means they travelled the SAME DISTANCE, which is what "mirror" means
        # -- a l→r that goes 3x as far as its r→l twin is not a mirror.
        opp = (txa * txb) < 0
        denom = max(abs(txa), abs(txb))
        ratio = min(abs(txa), abs(txb)) / denom if denom > 1e-9 else 0.0
        by_shift = bool(opp and ratio >= 0.7)
        out["shift_opposite"] = bool(opp)
        out["shift_ratio"] = round(ratio, 3)

    # ---- instrument 1: change-centroid trajectory ------------------------
    ta = a_res.get("cx_trace" if axis == "x" else "cy_trace") or []
    tb = b_res.get("cx_trace" if axis == "x" else "cy_trace") or []
    out["n_a"], out["n_b"] = len(ta), len(tb)
    r = None
    if len(ta) >= 4 and len(tb) >= 4:
        n = 24
        ya = np.interp(np.linspace(0, 1, n), np.linspace(0, 1, len(ta)),
                       np.array(ta, dtype=float))
        yb = np.interp(np.linspace(0, 1, n), np.linspace(0, 1, len(tb)),
                       np.array(tb, dtype=float))
        if ya.std() >= 1e-4 and yb.std() >= 1e-4:
            r = float(np.corrcoef(ya, yb)[0, 1])
    out["r"] = None if r is None else round(r, 3)

    da, db = a_res.get("direction"), b_res.get("direction")
    opposite = (da != db) or "stationary" in (da, db)
    out["labels_opposite"] = bool(opposite)

    # ---- combine: either instrument may carry the verdict ----------------
    by_trace = None
    if r is not None and r <= -0.7 and opposite:
        by_trace = True
    elif r is not None and r >= 0.7 and not opposite:
        by_trace = False
    if by_shift is True:
        out["verdict"] = "mirror"
        out["via"] = "shift" if not by_trace else "both"
    elif by_trace is True:
        out["verdict"] = "mirror"
        out["via"] = "trace"
    elif by_trace is False:
        out["verdict"] = "same"
        out["via"] = "trace"
    elif shift_mag <= 0.15 and r is None:
        out["verdict"] = "no-motion"
    else:
        out["verdict"] = "ambiguous"
        out["via"] = "none"
    return out


def cmd_dirmirror(video_dir, out=None):
    """Pair up the two dir renders of each spec and judge whether dir mirrors.

    Reads the manifest to learn each deck's spec/dir/axis, runs the SAME shape
    analyser the default probe uses (so the two layers cannot drift), then
    compares the two trajectories along that spec's axis.

    Pairing is by spec (whatever two dir values DIR_PROBE declared for it), not
    hard-coded to l/r -- horizontal effects are probed l/r, vertical ones u/d.
    """
    import glob as _glob
    import cv2

    videos = sorted(_glob.glob(os.path.join(video_dir, "*.mp4")))
    if not videos:
        print("!! %s 下没有 mp4" % video_dir)
        return 1

    want_axis = {s: e["axis"] for s, e in _dir_pairs().items()}
    by_spec = {}
    for v in videos:
        stem = os.path.splitext(os.path.basename(v))[0]
        man = os.path.join(video_dir, stem + ".manifest.json")
        if not os.path.exists(man):
            print("-- %s: 无 manifest，跳过" % stem)
            continue
        m = _load_json(man)
        cap = cv2.VideoCapture(v)
        frames = []
        while True:
            ok, f = cap.read()
            if not ok:
                break
            frames.append(f)
        cap.release()
        if not frames:
            continue
        fps = m.get("fps", FPS)
        bnd = (m.get("expected_boundary_frames") or [SLIDE_SECONDS * fps])[0]
        # Window centred on the slide boundary, +/- 1.5 s. An earlier
        # -2s..+3s window silently swallowed the WHOLE clip on these renders
        # (148 frames: bnd=60, so bnd+3*fps=150 > 148 and bnd-2*fps=0), which
        # dragged the dwell of slide 1 into the analysis and made `curtains`
        # read as t->b on a one-frame active window. A transition is a
        # ~0.5-1.0 s event; the window has to be tight around it.
        seg = frames[max(0, bnd - int(1.5 * fps)):
                     min(len(frames), bnd + int(1.5 * fps))]
        res = _profile_metrics(seg) or {}
        res.update({"spec": m.get("spec"), "dir": m.get("dir"),
                    "axis": m.get("axis") or want_axis.get(m.get("spec"), "x"),
                    "ui": m.get("ui"), "deck": stem})
        by_spec.setdefault(m.get("spec"), {})[m.get("dir")] = res

    rows = []
    for spec in sorted(by_spec):
        pair = by_spec[spec]
        if len(pair) < 2:
            # Report a lone deck so a missing half is visible instead of
            # silently absent.
            rows.append({"spec": spec, "dirs": sorted(pair),
                         "verdict": "unpaired",
                         "note": "只有 %s，无法配对比较"
                                 % "/".join(sorted(pair))})
            continue
        # Take the two dir values DIR_PROBE declared, in a stable order.
        decl = [d for d, _w in
                [(d, w) for _s, d, _a, w in DIR_PROBE if _s == spec]]
        keys = [d for d in decl if d in pair] or sorted(pair)
        ka, kb = keys[0], keys[1]
        ra, rb = pair[ka], pair[kb]
        axis = ra.get("axis") or "x"
        v = mirror_verdict(ra, rb, axis=axis)
        v.update({"spec": spec, "ui": ra.get("ui"), "axis": axis,
                  "dir_a": ka, "dir_b": kb,
                  "translate_a": ra.get("translate"),
                  "translate_b": rb.get("translate"),
                  "band_a": ra.get("band_travel"),
                  "band_b": rb.get("band_travel")})
        rows.append(v)

    print("\n%-10s %-8s %-6s %-9s %-6s %-7s %-10s %-10s %s" % (
        "spec", "界面名", "轴", "判定", "via", "r",
        "dir_a 位移", "dir_b 位移", "位移量(tx,ty)"))
    print("-" * 108)
    for r in rows:
        if r.get("verdict") == "unpaired":
            print("%-10s %-8s %-6s %-9s %-6s %-7s %-10s %-10s %s" % (
                r["spec"], "", "-", "unpaired", "-", "-", "-", "-",
                r.get("note", "")))
            continue
        print("%-10s %-8s %-6s %-9s %-6s %-7s %-10s %-10s %s" % (
            r["spec"], r.get("ui") or "", r.get("axis") or "-",
            r.get("verdict"), r.get("via") or "-", r.get("r"),
            r.get("translate_a") or "-", r.get("translate_b") or "-",
            "dir=%-6s tx=(%+.3f,%+.3f) shift=%s"
            % ("%s/%s" % (r.get("dir_a"), r.get("dir_b")),
               r.get("tx_a") or 0, r.get("tx_b") or 0,
               r.get("shift_mag"))))

    if out:
        json.dump({"mirror": rows, "dir_probe": list(DIR_PROBE)},
                  io.open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("\n-> %s" % out)
    return 0


def _read_frames(path):
    import cv2
    cap = cv2.VideoCapture(path)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    return frames


def _window(frames, man):
    fps = man.get("fps", FPS)
    bnd = (man.get("expected_boundary_frames") or [SLIDE_SECONDS * fps])[0]
    return frames[max(0, bnd - int(1.5 * fps)):min(len(frames), bnd + int(1.5 * fps))]


def cmd_attrdiff(video_dir, out=None):
    """Compare the attribute variants of each spec: does the attribute DO anything?

    The question the name layer could not answer: it records ONE child per
    effect (`clock -> <p:wheel spokes="1"/>`), while the PowerPoint-written
    presetID enum shows several effects carry FURTHER attribute values. Two
    distinct outcomes must be told apart, exactly as in the dir probe:

      * the variants render IDENTICALLY  -> the attribute value is cosmetic /
        ignored (or the value is equivalent to the default);
      * the variants render DIFFERENTLY  -> the attribute really changes the
        shape, and the shape layer must describe each value.

    Uses a **third, independent instrument** as the primary signal: the
    per-frame PIXEL DIFF of the two renderings. Unlike cx_trace / tx it needs
    no axis assumption -- an attribute is not necessarily a direction
    (spokes, orient, pattern, isContent are not), so an axis-based verdict
    would be the wrong question. Shape metrics are reported alongside for
    description, not for the "did it change" verdict.
    """
    import glob as _glob
    import numpy as np
    import cv2

    videos = sorted(_glob.glob(os.path.join(video_dir, "*.mp4")))
    if not videos:
        print("!! %s 下没有 mp4" % video_dir)
        return 1

    by_spec = {}
    for v in videos:
        stem = os.path.splitext(os.path.basename(v))[0]
        man = os.path.join(video_dir, stem + ".manifest.json")
        if not os.path.exists(man):
            print("-- %s: 无 manifest，跳过" % stem)
            continue
        m = _load_json(man)
        frames = _read_frames(v)
        if not frames:
            continue
        seg = _window(frames, m)
        prof = _profile_metrics(seg) or {}
        by_spec.setdefault(m.get("spec"), {})[m.get("label")] = {
            "label": m.get("label"), "child": m.get("child"),
            "ui": m.get("ui"), "deck": stem,
            "frames": seg, "profile": prof}

    rows = []
    for spec in sorted(by_spec):
        variants = by_spec[spec]
        if len(variants) < 2:
            rows.append({"spec": spec, "labels": sorted(variants),
                         "verdict": "single",
                         "note": "只有一个变体，无从比较"})
            continue
        labels = sorted(variants)
        # baseline = a row that means "no attribute / the default form" if one
        # exists, else the first alphabetically (stable, and reported
        # explicitly). Order matters: `plain`/`default`/`spokesNONE` all mean
        # "I stripped the attribute", and `horz` only means default for split
        # -- it must NOT outrank `plain` (an earlier version had horz first and
        # silently used <p:comb dir="horz"/> as comb's baseline).
        base = next((l for l in ("plain", "default", "spokesNONE")
                     if l in labels), None)
        if base is None:
            base = next((l for l in ("horz",) if l in labels), labels[0])
        bf = variants[base]["frames"]
        pairs = []
        for l in labels:
            if l == base:
                continue
            af = variants[l]["frames"]
            n = min(len(af), len(bf))
            diffs = [float(np.mean(cv2.absdiff(af[i], bf[i])))
                     for i in range(n)]
            mx = max(diffs) if diffs else 0.0
            mean = sum(diffs) / len(diffs) if diffs else 0.0
            p = variants[l]["profile"]
            bp = variants[base]["profile"]
            pairs.append({
                "label": l, "child": variants[l]["child"],
                "max_frame_diff": round(mx, 3),
                "mean_frame_diff": round(mean, 3),
                "changed_frames": sum(1 for x in diffs if x > 0.5),
                "differs": bool(mx > 0.5),
                "band_travel": p.get("band_travel"),
                "base_band_travel": bp.get("band_travel"),
                "direction": p.get("direction"),
                "base_direction": bp.get("direction"),
                "span": p.get("span"), "base_span": bp.get("span"),
            })
        rows.append({"spec": spec, "labels": labels, "base": base,
                     "ui": variants[base]["ui"],
                     "base_child": variants[base]["child"],
                     "base_direction": variants[base]["profile"].get("direction"),
                     "base_band": variants[base]["profile"].get("band_travel"),
                     "pairs": pairs})

    print("\n%-11s %-24s %-9s %-14s %-9s %s" % (
        "spec", "variant", "differs?", "maxFrameDiff", "方向", "child"))
    print("-" * 118)
    for r in rows:
        if r.get("verdict") == "single":
            print("%-11s %-24s %-9s %-14s %-9s %s" % (
                r["spec"], "-", "single", "-", "-", r.get("note", "")))
            continue
        print("%-11s %-24s %-9s %-14s %-9s %s" % (
            r["spec"], "%s (base)" % r["base"], "-", "-",
            r.get("base_direction"), r.get("base_child")))
        for p in r["pairs"]:
            print("%-11s %-24s %-9s %-14s %-9s %s" % (
                "", p["label"], "DIFF" if p["differs"] else "same",
                p["max_frame_diff"],
                "%s->%s" % (r.get("base_direction"), p["direction"]),
                p["child"]))

    if out:
        slim = []
        for r in rows:
            rr = {k: v for k, v in r.items() if k != "frames"}
            slim.append(rr)
        json.dump({"attrs": slim, "attr_probe": [list(x) for x in ATTR_PROBE]},
                  io.open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("\n-> %s" % out)
    return 0


def _energy_trace(frames):
    """Per-frame frame-to-frame change energy, over the WHOLE clip.

    Unlike _profile_metrics this does not decide "where is the active window and
    what shape is it" -- it answers "WHEN is anything moving". The timing probe
    needs exactly that: a transition and an animation are two separate events in
    time, and the question is whether they overlap, abut, or are one event.

    Localised to the centre region on purpose: the animating marker lives in the
    middle, while a whole-page transition lights up every column. Reporting both
    a full-frame and a centre trace lets the analyser tell "the animation played"
    from "the page slid" without assuming which is which.
    """
    import cv2
    import numpy as np

    if len(frames) < 2:
        return None
    gray = [cv2.cvtColor(f, cv2.COLOR_BGR2GRAY).astype(np.float32) for f in frames]
    h, w = gray[0].shape
    y0, y1 = int(h * 0.35), int(h * 0.65)
    x0, x1 = int(w * 0.35), int(w * 0.65)
    full, centre = [], []
    for i in range(1, len(gray)):
        d = np.abs(gray[i] - gray[i - 1])
        full.append(float(d.mean()))
        centre.append(float(d[y0:y1, x0:x1].mean()))
    return {"full": full, "centre": centre,
            "full_max": max(full) if full else 0.0,
            "centre_max": max(centre) if centre else 0.0}


def _hot_runs(trace, frac=0.15, min_len=2):
    """Contiguous runs where the trace exceeds `frac` of its own peak.

    Returns [(start, end, peak_index, peak_value)]. Runs, not a threshold count,
    because both a transition and an animation are contiguous events; counting
    scattered hot frames would merge two events into one blurry one -- the exact
    mistake _profile_metrics warns about (§七).
    """
    if not trace:
        return []
    peak = max(trace)
    if peak <= 1e-9:
        return []
    thr = peak * frac
    runs, s, p = [], None, None
    for i, v in enumerate(trace):
        if v > thr:
            if s is None:
                s = i
            p = i
        else:
            if s is not None and p - s >= min_len - 1:
                runs.append((s, p))
            s = None
    if s is not None and p - s >= min_len - 1:
        runs.append((s, p))
    out = []
    for s, e in runs:
        seg = trace[s:e + 1]
        pk = s + max(range(len(seg)), key=lambda i: seg[i])
        out.append((s, e, pk, trace[pk]))
    return out


def _entry_trace(frames, lo=0, hi=None):
    """Track the ENTERING shape's colour mass over time.

    Why not the generic energy trace: the fade-in of a mid-tone disc against a
    mid-tone grid moves the frame-mean by <2/255, which is under every
    reasonable energy threshold -- the `solo_anim` deck reported ZERO active
    frames that way, which reads as "the animation never ran". It did run; the
    instrument was blind to it.

    The probe deck is built so the entering slide's marker is a DIFFERENT hue
    (page 1 red, page 2 blue). That makes "has the new page's shape appeared
    yet" a direct colour measurement, with no threshold heuristics: the mass of
    `max(0, B - R)` is exactly how much blue-vs-red is on screen.

    Returns the per-frame mass plus the ramp window, which is the number that
    actually answers "when did the animation start".
    """
    import numpy as np

    mass = []
    for f in frames:
        b = f[:, :, 0].astype(np.int32)
        r = f[:, :, 2].astype(np.int32)
        mass.append(float(np.maximum(0, b - r).sum()))
    if not mass:
        return None
    peak = max(mass)
    if peak <= 0:
        return {"mass": mass, "peak": 0.0, "ramp_start": None,
                "ramp_end": None, "ramp_ms": None}
    two = next((i for i, v in enumerate(mass) if v > peak * 0.02), None)
    sat = next((i for i, v in enumerate(mass) if v > peak * 0.97), None)
    return {
        "mass": mass, "peak": peak, "ramp_start": two, "ramp_end": sat,
        "ramp_ms": (round((sat - two) * 1000.0 / FPS)
                    if two is not None and sat is not None else None),
    }


def cmd_timingdiff(video_dir, out=None, png=None):
    """Does <p:transition> coexist with <p:timing> -- and if so, WHO goes first?

    Three questions, three instruments (see the TIMING_PROBE comment):

      (a) both blocks present after a round-trip?  -> read the saved XML
      (b) when does each event fire?               -> per-frame energy runs
      (c) does either one change shape?            -> pixel diff vs solo control

    (a) cannot be answered from the deck we wrote -- PowerPoint may drop one on
    save without a prompt (com-pitfalls §12), so it is read from the *-saved
    decks when they exist. If they do not, the verdict is reported as unmeasured
    rather than assumed.
    """
    import glob as _glob
    import numpy as np
    import cv2

    videos = sorted(_glob.glob(os.path.join(video_dir, "*.mp4")))
    if not videos:
        print("!! %s 下没有 mp4" % video_dir)
        return 1

    rows = []
    by_label = {}
    for v in videos:
        stem = os.path.splitext(os.path.basename(v))[0]
        man = os.path.join(video_dir, stem + ".manifest.json")
        if not os.path.exists(man):
            print("-- %s: 无 manifest，跳过" % stem)
            continue
        m = _load_json(man)
        frames = _read_frames(v)
        if not frames:
            continue
        et = _energy_trace(frames)
        if et is None:
            continue
        # The marker lives in the middle third; the transition lights the whole
        # frame. A run that shows up ONLY in `centre` is the animation; one that
        # shows up in `full` is the transition (or both, when they overlap).
        centre_runs = _hot_runs(et["centre"])
        full_runs = _hot_runs(et["full"])
        ent = _entry_trace(frames) or {}
        bnd = (m.get("expected_boundary_frames") or [SLIDE_SECONDS * FPS])[0]
        delay = None
        if ent.get("ramp_start") is not None:
            delay = ent["ramp_start"] - int(bnd)
        rec = {
            "label": m.get("label"), "spec": m.get("spec"),
            "tspec": m.get("tspec"),
            "has_transition": m.get("has_transition"),
            "has_timing": m.get("has_timing"),
            "deck": stem, "frames": len(frames),
            "centre_runs": [{"start": s, "end": e, "peak": pk, "peak_value": round(pv, 2)}
                            for s, e, pk, pv in centre_runs],
            "full_runs": [{"start": s, "end": e, "peak": pk, "peak_value": round(pv, 2)}
                          for s, e, pk, pv in full_runs],
            "centre_max": round(et["centre_max"], 3),
            "full_max": round(et["full_max"], 3),
            "boundary_frame": int(bnd),
            "entry_ramp_start": ent.get("ramp_start"),
            "entry_ramp_end": ent.get("ramp_end"),
            "entry_ramp_ms": ent.get("ramp_ms"),
            # The headline number: how long AFTER the slide boundary does the
            # entering shape begin to appear?  `solo_trans` (no animation) is
            # the baseline this is read against.
            "entry_delay_frames": delay,
            "entry_delay_ms": (round(delay * 1000.0 / FPS) if delay is not None else None),
            "energy_centre": [round(x, 3) for x in et["centre"]],
            "energy_full": [round(x, 3) for x in et["full"]],
            "_frames": frames,
        }
        rows.append(rec)
        by_label[m.get("label")] = rec

    # (a) structural: did PowerPoint keep both?  Look for the saved decks.
    saved = {}
    for p in _glob.glob(os.path.join(video_dir, "*-saved.pptx")) + \
            _glob.glob(os.path.join(video_dir, "*_saved.pptx")):
        stem = os.path.splitext(os.path.basename(p))[0]
        try:
            with zipfile.ZipFile(p) as z:
                s2 = z.read("ppt/slides/slide2.xml").decode("utf-8")
        except Exception:
            continue
        saved[stem] = {
            "has_transition": bool(re.search(r"<p:transition\b", s2)),
            "has_timing": bool(re.search(r"<p:timing>", s2)),
            # ORDER is part of the question: the sequence allows transition then
            # timing only. Report which came first in the SAVED file.
            "order": ("transition-before-timing"
                      if s2.find("<p:transition") < s2.find("<p:timing>")
                      else "timing-before-transition"
                      if "<p:timing>" in s2 and "<p:transition" in s2
                      else "n/a"),
            "moved_to_extLst": s2.find("<p:timing>") > s2.find("<p:extLst"),
        }

    # (c) interference: combination vs its solo controls, per pair.
    def _diff(a, b):
        n = min(len(a["_frames"]), len(b["_frames"]))
        if n < 2:
            return None
        dfs = [float(np.mean(cv2.absdiff(a["_frames"][i], b["_frames"][i])))
               for i in range(n)]
        return {"max": round(max(dfs), 3), "mean": round(sum(dfs) / len(dfs), 3)}

    verdicts = []
    for spec, both_lbl, anim_lbl, trans_lbl in (
            ("push", "both_push", "solo_anim", "solo_trans"),
            ("wipe", "both_wipe", "solo_anim_wipe", "solo_trans_wipe"),
            ("fade", "both_fade_short", None, "solo_fade_800")):
        both = by_label.get(both_lbl)
        anim = by_label.get(anim_lbl) if anim_lbl else None
        trans = by_label.get(trans_lbl)
        v = {"spec": spec, "both": both_lbl,
             "anim_control": anim_lbl, "trans_control": trans_lbl}
        if both and (anim or trans):
            # Timeline: how many distinct energy runs, and where are they?
            br = both["centre_runs"]
            ar = anim["centre_runs"] if anim else []
            tr = trans["full_runs"] if trans else []
            v["both_n_runs"] = len(br)
            v["anim_peak_frame"] = ar[0]["peak"] if ar else None
            v["trans_peak_frame"] = tr[0]["peak"] if tr else None
            v["both_centre_peak"] = br[0]["peak"] if br else None
            v["anim_fires_with_transition"] = bool(br)
            # Did the animation window survive?  Compare peak positions.
            if ar and br:
                v["peak_shift_frames"] = br[0]["peak"] - ar[0]["peak"]
            if anim:
                v["combined_vs_solo_anim_max"] = _diff(both, anim) or {}
            if trans:
                v["combined_vs_solo_trans_max"] = _diff(both, trans) or {}
        verdicts.append(v)

    print("\n%-10s %-24s %-8s %-8s %-11s %-11s" % (
        "spec", "deck", "trans?", "anim?", "centreRuns", "fullRuns"))
    print("-" * 88)
    for r in rows:
        print("%-10s %-24s %-8s %-8s %-11d %-11d" % (
            r["spec"], r["label"], r["has_transition"], r["has_timing"],
            len(r["centre_runs"]), len(r["full_runs"])))

    print("\n== 进入形状什么时候出现（boundary 在 f%d） ==" % (
        rows[0]["boundary_frame"] if rows else 0))
    print("  %-24s %-11s %-8s %-9s %-9s %s" % (
        "deck", "trans?", "anim?", "rampStart", "delay(ms)", "ramp(ms)"))
    print("  " + "-" * 74)
    for r in sorted(rows, key=lambda x: (x.get("spec") or "", x["label"])):
        delay = r.get("entry_delay_ms")
        print("  %-24s %-11s %-8s %-9s %-9s %s" % (
            r["label"], "yes" if r["has_transition"] else "no",
            "yes" if r["has_timing"] else "no",
            r.get("entry_ramp_start"),
            "%+d" % delay if delay is not None else "—",
            r.get("entry_ramp_ms")))

    print("\n== 时间线（centre 通道里的事件峰） ==")
    for r in rows:
        cr = ", ".join("f%d-%d(peak f%d, %.1f)" % (
            x["start"], x["end"], x["peak"], x["peak_value"])
            for x in r["centre_runs"]) or "—"
        fr = ", ".join("f%d-%d" % (x["start"], x["end"]) for x in r["full_runs"]) or "—"
        print("  %-24s centre: %-46s full: %s" % (r["label"], cr, fr))

    print("\n== 结构（PowerPoint 存回后） ==")
    if saved:
        for stem in sorted(saved):
            s = saved[stem]
            print("  %-30s transition=%s timing=%s order=%s" % (
                stem, s["has_transition"], s["has_timing"], s["order"]))
    else:
        print("  （没有 *-saved.pptx；未测 —— 见 probe_roundtrip.ps1）")

    print("\n== 组合 vs 单独 ==")
    for v in verdicts:
        print("  %-6s 切换单独峰 f%s / 组合里 centre 事件 %s 个" % (
            v["spec"], v.get("trans_peak_frame"), v.get("both_n_runs")))
        if v.get("anim_peak_frame") is not None:
            print("         动画单独峰 f%s" % v["anim_peak_frame"])
        d = v.get("combined_vs_solo_anim_max") or {}
        if d:
            print("         组合 vs 单独动画，逐帧像素差 max=%.3f" % d.get("max", 0))
        d = v.get("combined_vs_solo_trans_max") or {}
        if d:
            print("         组合 vs 单独切换，逐帧像素差 max=%.3f" % d.get("max", 0))

    if out:
        slim = []
        for r in rows:
            rr = {k: val for k, val in r.items() if k != "_frames"}
            slim.append(rr)
        json.dump({"timing": slim, "verdicts": verdicts, "saved": saved,
                   "timing_probe": [list(x) for x in TIMING_PROBE]},
                  io.open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("\n-> %s" % out)

    if png:
        _timing_chart(rows, png)
        print("-> %s" % png)
    return 0


def _timing_chart(rows, png, series=("solo_anim", "both_fade_short",
                                     "both_push", "both_fade_long")):
    """Plot the entering shape's colour mass over time for a few decks.

    The delay law is a TEMPORAL fact, and a temporal fact is best shown as a
    timeline -- a table of "delay = 933 ms" invites the reader to doubt the
    definition of "delay", while a curve that visibly starts 900 ms after the
    boundary does not.

    Draws the slide boundary as a vertical line so the eye can measure the gap
    itself, and marks each curve's ramp start.
    """
    import numpy as np
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
    except Exception as exc:              # matplotlib is optional
        print("!! 画图需要 matplotlib：%s" % exc)
        return

    have = [r for r in rows if r["label"] in series]
    if not have:
        print("!! 没有可画的序列")
        return
    # CJK fonts are not guaranteed; label in ASCII so the chart is readable
    # everywhere rather than showing tofu boxes.
    labels = {
        "solo_anim": "no transition (hard cut)",
        "both_fade_short": "fade 300 ms + anim",
        "both_push": "push 800 ms + anim",
        "both_wipe": "wipe 800 ms + anim",
        "both_fade_long": "fade 1500 ms + anim",
    }
    fig, ax = plt.subplots(figsize=(11, 5.2), dpi=110)
    colours = ["#2c7fb8", "#31a354", "#d95f0e", "#756bb1", "#636363"]
    bnd = have[0]["boundary_frame"]
    ax.axvline(bnd, color="#999999", ls="--", lw=1.2)
    ax.text(bnd + 1, 0, " slide boundary (f%d)" % bnd, color="#666666",
            fontsize=8, va="bottom")

    for i, r in enumerate(have):
        et = _entry_trace(r["_frames"])
        if not et or not et["mass"]:
            continue
        m = np.array(et["mass"], dtype=float)
        if m.max() <= 0:
            continue
        m = m / m.max()
        n = len(m)
        ax.plot(range(n), m, color=colours[i % len(colours)], lw=1.8,
                label=labels.get(r["label"], r["label"]))
        rs = r.get("entry_ramp_start")
        if rs is not None:
            ax.plot([rs], [m[rs]], "o", color=colours[i % len(colours)], ms=6)
            # Stagger the delay arrows or they overprint each other in the header.
            yb = 1.02 + i * 0.045
            ax.annotate("", xy=(rs, yb), xytext=(bnd, yb),
                        arrowprops=dict(arrowstyle="<->",
                                        color=colours[i % len(colours)], lw=1.1))
            ax.text((bnd + rs) / 2.0, yb + 0.015,
                    "%+d ms" % (r.get("entry_delay_ms") or 0),
                    ha="center", fontsize=8,
                    color=colours[i % len(colours)])

    ax.set_xlabel("frame (30 fps)")
    ax.set_ylabel("entering shape mass (normalised)")
    ax.set_title("When does the animated shape appear?\n"
                 "entrance animation start is pushed back by the transition duration")
    ax.set_ylim(-0.05, 1.26)
    ax.grid(alpha=0.25)
    ax.legend(fontsize=8, loc="center right")
    fig.tight_layout()
    fig.savefig(png)
    plt.close(fig)


def cmd_shapeanalyze(video, manifest, out=None):
    """Report the shape of one rendered transition."""
    import cv2

    man = _load_json(manifest)
    cap = cv2.VideoCapture(video)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    if not frames:
        print("!! %s 读不出帧" % video)
        return 1
    fps = man.get("fps", FPS)
    # Only the boundary region matters; trim to the known boundary +- 2s.
    bnd = (man.get("expected_boundary_frames") or [SLIDE_SECONDS * fps])[0]
    lo = max(0, bnd - 2 * fps)
    hi = min(len(frames), bnd + 3 * fps)
    seg = frames[lo:hi]

    res = _profile_metrics(seg)
    res = dict(res or {})
    res.update({"spec": man.get("spec"), "ui": man.get("ui"),
                "group": man.get("group"), "family": man.get("family"),
                "child": man.get("child"), "video_frames": len(frames)})
    line = ("%-14s %-8s dir=%-11s mode=%-12s sym_v=%.2f sym_h=%.2f span=%d")
    print(line % (res.get("spec"), res.get("ui"), res.get("direction"),
                  res.get("mode"), res.get("symmetry_v", 0),
                  res.get("symmetry_h", 0), res.get("span", 0)))
    if out:
        all_res = []
        if os.path.exists(out):
            try:
                all_res = _load_json(out).get("shapes", [])
            except Exception:
                all_res = []
        all_res = [r for r in all_res if r.get("spec") != res["spec"]]
        all_res.append(res)
        all_res.sort(key=lambda r: str(r.get("spec")))
        json.dump({"shapes": all_res},
                  io.open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    return 0


def cmd_shapes(video_dir, out):
    """Analyse every <dir>/*.mp4 that has a matching manifest."""
    import glob as _glob

    videos = sorted(_glob.glob(os.path.join(video_dir, "*.mp4")))
    if not videos:
        print("!! %s 下没有 mp4" % video_dir)
        return 1
    results = []
    for v in videos:
        stem = os.path.splitext(os.path.basename(v))[0]
        man = os.path.join(video_dir, stem + ".manifest.json")
        if not os.path.exists(man):
            print("-- %s: 无 manifest，跳过" % stem)
            continue
        import cv2
        cap = cv2.VideoCapture(v)
        frames = []
        while True:
            ok, f = cap.read()
            if not ok:
                break
            frames.append(f)
        cap.release()
        if not frames:
            continue
        m = _load_json(man)
        fps = m.get("fps", FPS)
        bnd = (m.get("expected_boundary_frames") or [SLIDE_SECONDS * fps])[0]
        # Window centred on the slide boundary, +/- 1.5 s. An earlier
        # -2s..+3s window silently swallowed the WHOLE clip on these renders
        # (148 frames: bnd=60, so bnd+3*fps=150 > 148 and bnd-2*fps=0), which
        # dragged the dwell of slide 1 into the analysis and made `curtains`
        # read as t->b on a one-frame active window. A transition is a
        # ~0.5-1.0 s event; the window has to be tight around it.
        seg = frames[max(0, bnd - int(1.5 * fps)):
                     min(len(frames), bnd + int(1.5 * fps))]
        res = _profile_metrics(seg) or {}
        res.update({"spec": m.get("spec"), "ui": m.get("ui"),
                    "group": m.get("group"), "family": m.get("family"),
                    "child": m.get("child")})
        results.append(res)
    results.sort(key=lambda r: str(r.get("spec")))
    print("\n%-14s %-10s %-10s %-10s %-13s %s" % (
        "spec", "界面名", "位移", "变化漂移", "推进模式", "对称(v/h) 帧跨度"))
    print("-" * 92)
    for r in results:
        print("%-14s %-10s %-10s %-10s %-13s %.2f / %.2f  %s" % (
            r.get("spec"), r.get("ui"), r.get("translate"),
            r.get("direction"), r.get("mode"),
            r.get("symmetry_v", 0), r.get("symmetry_h", 0), r.get("span")))
    if out:
        json.dump({"shapes": results}, io.open(out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
        print("\n-> %s（%d 条）" % (out, len(results)))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="切换效果实测表")
    sub = ap.add_subparsers(dest="cmd", required=True)
    b = sub.add_parser("build")
    b.add_argument("case_dir")
    c = sub.add_parser("collect")
    c.add_argument("case_dir")
    c.add_argument("out_dir")
    c.add_argument("--json")
    c.add_argument("--enum")
    e = sub.add_parser("enumdeck")
    e.add_argument("out_pptx")
    e.add_argument("lo", type=int)
    e.add_argument("hi", type=int)
    er = sub.add_parser("enumread")
    er.add_argument("deck")
    er.add_argument("--json")
    er.add_argument("--manifest")
    v = sub.add_parser("video")
    v.add_argument("case_dir")
    v.add_argument("out_dir")
    v.add_argument("out_pptx")
    s = sub.add_parser("sheets")
    s.add_argument("video")
    s.add_argument("out_dir")
    s.add_argument("--want", type=int, default=10)
    t = sub.add_parser("table")
    t.add_argument("report")
    t.add_argument("out")
    t.add_argument("--scans", nargs="+", required=True)
    ad = sub.add_parser("anchordeck")
    ad.add_argument("out_dir")
    an = sub.add_parser("anchors")
    an.add_argument("video")
    an.add_argument("manifest")
    sd = sub.add_parser("shapedeck")
    sd.add_argument("out_dir")
    sd2 = sub.add_parser("shapedeck2")
    sd2.add_argument("out_dir")
    sd2.add_argument("--specs", nargs="+")
    sa = sub.add_parser("shapeanalyze")
    sa.add_argument("video")
    sa.add_argument("manifest")
    sa.add_argument("--out")
    sh = sub.add_parser("shapes")
    sh.add_argument("video_dir")
    sh.add_argument("--out")
    dd = sub.add_parser("dirdeck")
    dd.add_argument("out_dir")
    dm = sub.add_parser("dirmirror")
    dm.add_argument("video_dir")
    dm.add_argument("--out")
    at = sub.add_parser("attrdeck")
    at.add_argument("out_dir")
    am = sub.add_parser("attrdiff")
    am.add_argument("video_dir")
    am.add_argument("--out")
    td = sub.add_parser("timingdeck")
    td.add_argument("out_dir")
    tm = sub.add_parser("timingdiff")
    tm.add_argument("video_dir")
    tm.add_argument("--out")
    tm.add_argument("--png")
    ns = ap.parse_args(argv)
    if ns.cmd == "build":
        return cmd_build(ns.case_dir)
    if ns.cmd == "collect":
        return cmd_collect(ns.case_dir, ns.out_dir, ns.json, ns.enum)
    if ns.cmd == "enumdeck":
        return cmd_enumdeck(ns.out_pptx, ns.lo, ns.hi)
    if ns.cmd == "enumread":
        return cmd_enumread(ns.deck, ns.json, ns.manifest)
    if ns.cmd == "video":
        return cmd_video(ns.case_dir, ns.out_dir, ns.out_pptx)
    if ns.cmd == "table":
        return cmd_table(ns.report, ns.scans, ns.out)
    if ns.cmd == "anchordeck":
        return cmd_anchordeck(ns.out_dir)
    if ns.cmd == "anchors":
        return cmd_anchors(ns.video, ns.manifest)
    if ns.cmd == "shapedeck":
        return cmd_shapedeck(ns.out_dir)
    if ns.cmd == "shapedeck2":
        return cmd_shapedeck2(ns.out_dir, ns.specs)
    if ns.cmd == "shapeanalyze":
        return cmd_shapeanalyze(ns.video, ns.manifest, ns.out)
    if ns.cmd == "shapes":
        return cmd_shapes(ns.video_dir, ns.out)
    if ns.cmd == "dirdeck":
        return cmd_dirdeck(ns.out_dir)
    if ns.cmd == "dirmirror":
        return cmd_dirmirror(ns.video_dir, ns.out)
    if ns.cmd == "attrdeck":
        return cmd_attrdeck(ns.out_dir)
    if ns.cmd == "attrdiff":
        return cmd_attrdiff(ns.video_dir, ns.out)
    if ns.cmd == "timingdeck":
        return cmd_timingdeck(ns.out_dir)
    if ns.cmd == "timingdiff":
        return cmd_timingdiff(ns.video_dir, ns.out, ns.png)
    return cmd_sheets(ns.video, ns.out_dir, ns.want)


if __name__ == "__main__":
    sys.exit(main())
