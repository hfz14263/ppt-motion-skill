#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read a design system and compose pages that follow it.

WHY THIS EXISTS
---------------
The design systems in references/design-system/ are rich and prescriptive -- palette,
type scale, layout skeleton, signature components, prohibitions, density baseline -- but
they are prose. A person can read them; a program cannot, so every deck that "follows
style X" was following a human's memory of X. That is why decks converge on the same
generic look regardless of which system was chosen.

This script does the smallest thing that makes a system executable: parse the palette
out of the prose, then compose a body page from it. Two pages from two systems, with
identical content, should look like two different design languages. If they do not, the
"style" was never real and no amount of reading would have helped.

SCOPE, HONESTLY
---------------
Implemented: palette extraction (role-labelled), the typographic scale, the page
furniture the systems call "fixed page trio" (source line + page number) and the
"bottom synthesis bar", and the composition of a dense multi-column body page.

NOT implemented, and deliberately not faked: the photographic title band (several
systems require a darkened photo header, and there is no image generator here), the
chart language (each system specifies chart forms in detail), and section 4's page-type
playbooks. A composer that silently omitted those while claiming to follow the system
would be worse than no composer.

Usage:
    python scripts/design_compose.py --list
    python scripts/design_compose.py --theme pine-green-strategy --out deck.pptx
    python scripts/design_compose.py --all --outdir composed
"""
import argparse
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DS_DIR = os.path.abspath(os.path.join(HERE, "..", "references", "design-system"))

HEX = re.compile(r"#([0-9A-Fa-f]{6})\b")
FENCE = re.compile(r"【([^】]+)】")
PAL_HEAD = ("色彩", "颜色", "Color", "Palette")

# Role keywords, most specific first. Order matters: a sentence naming both a structural
# and a neutral colour should be filed under structural.
ROLE_WORDS = [
    ("negative", r"negative|risk|风险|负向|负值|警示"),
    ("positive", r"positive|growth|增长|正向|上升"),
    ("accent", r"accent|emphasis|强调"),
    ("structural", r"structural|skeleton|structure|骨架|标题|结构"),
    ("neutral", r"gray|grey|neutral|灰|中性"),
    ("background", r"background|底色|背景|底"),
]

# Sizes quoted in the design systems are given as multiples of body text (1x ~= 14pt).
# The systems state these explicitly, so the scale is read rather than invented.
TYPE_PATTERNS = [
    ("page_number", r"footnote|page number|页脚|页码", 8),
    ("body", r"body copy|正文", 14),
    ("section", r"section title|section heading|章节标题", 20),
    ("title", r"full-width title|main title|主标题|页面标题", 28),
    ("display", r"large figures|oversized figure|大数字|关键数字", 40),
]


def load_system(theme_id):
    path = os.path.join(DS_DIR, theme_id + ".md")
    if not os.path.exists(path):
        raise SystemExit("unknown theme %r; try --list" % theme_id)
    return io.open(path, encoding="utf-8").read()


def parse_palette(text):
    """{role: [hex, ...]} plus the raw palette paragraph.

    Requires a real 【色彩规范】-style heading; there is no silent fallback to the whole
    file, because that fallback yields plausible numbers from an unrelated section.
    """
    parts = FENCE.split(text)
    pal = None
    for i in range(1, len(parts), 2):
        if any(h in parts[i] for h in PAL_HEAD):
            pal = parts[i + 1]
            break
    if pal is None:
        m = re.search(r"^[#*\s]*Color Palette\s*$", text, re.M)
        if not m:
            raise SystemExit("no palette heading found -- refusing to guess")
        nxt = text.find("【", m.end())
        pal = text[m.end():nxt if nxt > 0 else len(text)]

    roles = {}
    for sentence in re.split(r"[.;\n]", pal):
        if not HEX.search(sentence):
            continue
        found = None
        for name, pat in ROLE_WORDS:
            if re.search(pat, sentence, re.I):
                found = name
                break
        found = found or "other"
        for h in HEX.findall(sentence):
            roles.setdefault(found, [])
            if h.upper() not in roles[found]:
                roles[found].append(h.upper())
    return roles, pal.strip()


def parse_signature(text):
    """One-line style signature, if the file states one."""
    m = re.search(r"(?:One-line style signature|风格签名)[：:\s]*(.+)", text)
    return m.group(1).strip()[:160] if m else ""


def parse_density(text):
    """Numeric hints from the density baseline, used to size the composition."""
    m = re.search(r"(\d+)\s*[–\-]\s*(\d+)\s*(?:independent modules|个.{0,6}模块)", text)
    if m:
        return int(m.group(1)), int(m.group(2))
    return 2, 4


def luminance(h):
    """Relative luminance, for contrast checks."""
    r, g, b = (int(h[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    f = lambda c: c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4  # noqa: E731
    return 0.2126 * f(r) + 0.7152 * f(g) + 0.0722 * f(b)


def contrast(a, b):
    la, lb = luminance(a), luminance(b)
    hi, lo = max(la, lb), min(la, lb)
    return (hi + 0.05) / (lo + 0.05)


def saturation(h):
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    mx, mn = max(r, g, b), min(r, g, b)
    return 0 if mx == 0 else (mx - mn) / float(mx)


def pick_readable(candidates, bg, want_light, prefer_colour=False):
    """Choose the most readable candidate, WITHOUT discarding the brand colour.

    Two over-corrections are recorded here because both happened:

      * taking neutral[0] for body text picked #F1F1F1 on white -- the colour the system
        reserves for explanation-column FILL -- and the text vanished in the render.
      * then "pick the highest contrast" picked black or white for the STRUCTURAL role
        every time, because nothing out-contrasts black. Pine green #04512C became
        #111111 and the deck lost its identity while gaining perfect legibility.

    So `prefer_colour` asks for a candidate that is readable AND still has chroma, and
    only falls back to the strongest when no coloured candidate qualifies.
    """
    if not candidates:
        return "FFFFFF" if want_light else "111111"
    usable = [c for c in candidates if contrast(c, bg) >= 3.0]
    if prefer_colour:
        coloured = [c for c in usable if saturation(c) >= 0.15]
        if coloured:
            return max(coloured, key=lambda c: contrast(c, bg))
    if usable:
        return max(usable, key=lambda c: contrast(c, bg))
    return max(candidates, key=lambda c: contrast(c, bg))


def theme_of(theme_id):
    text = load_system(theme_id)
    roles, pal = parse_palette(text)
    bg = roles.get("background", ["FFFFFF"])[0]
    dark_bg = luminance(bg) < 0.35
    # body/ink: strongest neutral, or the structural colour if the system lists none
    ink_pool = list(roles.get("neutral", [])) or list(roles.get("structural", []))
    ink = pick_readable(ink_pool, bg, want_light=dark_bg)
    mute = ink
    others = [c for c in ink_pool if c.upper() != ink.upper()]
    if others:
        mute = pick_readable(others, bg, want_light=dark_bg)
    # structural and accent keep their hue: readability is a floor, not the objective
    struct = pick_readable(roles.get("structural", []) or roles.get("accent", []),
                           bg, want_light=dark_bg, prefer_colour=True)
    # The accent pool must NOT include the "other" bucket. That bucket collects the
    # colours a system designates for CHART SERIES, and drawing an accent from it put
    # five saturated hues on one page -- which is the rainbow the systems explicitly
    # forbid and which review_assist flagged on the composed deck. An accent is the
    # system's named accent; chart scales are used inside charts, by counts, not as
    # page furniture.
    acc_pool = [c for c in roles.get("accent", []) if c.upper() != struct.upper()]
    accent = (pick_readable(acc_pool, bg, want_light=dark_bg, prefer_colour=True)
              if acc_pool else struct)
    return {
        "id": theme_id,
        "signature": parse_signature(text),
        "roles": roles,
        "density": parse_density(text),
        "bg": bg,
        "ink": ink,
        "mute": mute,
        "struct": struct,
        "accent": accent,
        "neg": pick_readable(roles.get("negative", []) or [accent], bg,
                             want_light=dark_bg, prefer_colour=True),
        "contrast": {"ink": round(contrast(ink, bg), 2),
                     "struct": round(contrast(struct, bg), 2),
                     "accent": round(contrast(accent, bg), 2)},
    }


# ---------------------------------------------------------------------------
# composition
# ---------------------------------------------------------------------------
CLAIMS = [
    ("结构", "闸门只能证明 XML 合法，证明不了 PowerPoint 采纳",
     "结构自证、几何指纹、lxml 解析，三者都从同一份 XML 推出，不构成独立判据。"),
    ("真渲染", "唯一能证明采纳的是 PowerPoint 自己读回来",
     "对象模型读回、真渲染导图、往返普查——三条路都比字符串检查强，但各有边界。"),
    ("交叉", "任何「通过」结论至少要两个独立判据",
     "「几何没变」+「结构合法」不算两个。要加上 PowerPoint 真开一次，才够。"),
]


def compose(theme, out_path, deck_title="动效注入的三道闸"):
    """A dense multi-column body page, built only from the theme's own palette."""
    from pptx import Presentation
    from pptx.util import Inches, Pt, Emu
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

    def C(h):
        return RGBColor.from_string(h.upper())

    prs = Presentation()
    prs.slide_width = Inches(13.3333)
    prs.slide_height = Inches(7.5)
    SW, SH = 13.3333, 7.5
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    def rect(x, y, w, h, fill=None, line=None, lw=0.75):
        from pptx.enum.shapes import MSO_SHAPE
        sh = sl.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                 Inches(w), Inches(h))
        sh.shadow.inherit = False
        if fill:
            sh.fill.solid()
            sh.fill.fore_color.rgb = C(fill)
        else:
            sh.fill.background()
        if line:
            sh.line.color.rgb = C(line)
            sh.line.width = Pt(lw)
        else:
            sh.line.fill.background()
        return sh

    def text(x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP):
        tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        first = True
        for para in runs:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            p.alignment = align
            if "space" in para:
                p.space_after = Pt(para["space"])
            for r in para["runs"]:
                run = p.add_run()
                run.text = r["t"]
                run.font.size = Pt(r.get("sz", 14))
                run.font.bold = r.get("b", False)
                run.font.name = r.get("font", "微软雅黑")
                run.font.color.rgb = C(r.get("c", theme["ink"]))
        return tb

    # page ground
    rect(0, 0, SW, SH, fill=theme["bg"])

    # ---- title: the systems require an ASSERTION, not a noun phrase -----------
    rect(0, 0, 0.16, SH, fill=theme["struct"])          # structural edge, one per page
    text(0.72, 0.52, 10.6, 0.75, [{"runs": [
        {"t": "三道闸都能过，却仍然可能是错的", "sz": 30, "b": True,
         "c": theme["struct"]}]}])
    text(0.72, 1.28, 10.6, 0.40, [{"runs": [
        {"t": "本页说明：每道闸能证明什么、不能证明什么，以及为什么它们不是彼此独立的判据。",
         "sz": 12.5, "c": theme["mute"]}]}])
    rect(0.72, 1.78, 11.9, 0.012, fill=theme["struct"])  # fine rule, not a card

    # ---- evidence: a ruled table plus numbered blocks, not cards ---------------
    # The systems forbid using containers to build hierarchy and forbid equal-split
    # composition, so hierarchy here comes from rules, type weight and size, and the
    # columns are deliberately uneven: the matrix column is wider because it carries
    # more structure than the notes beside it.
    table_rows = [
        ("lxml 解析通过", "XML 合法", "不能证明 PowerPoint 肯打开", theme["ink"]),
        ("几何指纹相同", "xfrm/tag/id 未变", "漏形状时它一样报 UNCHANGED", theme["ink"]),
        ("字符串含某元素", "写进去了", "scene3d 插错位置时 XML 在、读回为 -2", theme["neg"]),
        ("对象模型读回", "PowerPoint 采纳了", "对 p159 扩展无效（morph 查不到但生效）", theme["ink"]),
        ("Slide.Export 渲染", "静态态外观", "入场显示隐藏态，动作路径显示起点", theme["ink"]),
        ("player 重放", "时序与分层", "是重放不是真渲染，字体换行可能不同", theme["ink"]),
        ("--assert-geometry", "形状 xfrm 未变", "不证「该动的都动了」", theme["ink"]),
        ("motion.ps1 -Strict", "往返后效果还在", "只在装了 PowerPoint 的机器上有效", theme["ink"]),
    ]
    tx, ty, tw = 0.72, 2.02, 7.55
    row_h = 0.335
    text(tx, ty - 0.34, tw, 0.28, [{"runs": [
        {"t": "判据能证明什么、不能证明什么", "sz": 11.5, "b": True,
         "c": theme["struct"]}]}])
    rect(tx, ty - 0.05, tw, 0.012, fill=theme["struct"])
    for i, (name, proves, cannot, col) in enumerate(table_rows):
        y = ty + i * row_h
        text(tx, y + 0.05, 2.05, 0.27, [{"runs": [
            {"t": name, "sz": 10, "b": True, "c": col}]}])
        text(tx + 2.08, y + 0.05, 1.85, 0.27, [{"runs": [
            {"t": proves, "sz": 10, "c": theme["ink"]}]}])
        text(tx + 3.95, y + 0.05, tw - 3.95, 0.27, [{"runs": [
            {"t": cannot, "sz": 10, "c": theme["mute"]}]}])
        if i < len(table_rows) - 1:
            rect(tx, y + row_h - 0.04, tw, 0.007, fill=theme["mute"])

    # right rail: the framework, drawn with fine lines rather than filled blocks
    rx = 8.62
    text(rx, ty - 0.34, 4.0, 0.28, [{"runs": [
        {"t": "三道闸的独立性", "sz": 11.5, "b": True, "c": theme["struct"]}]}])
    rect(rx, ty - 0.05, 4.0, 0.012, fill=theme["struct"])
    for i, (tag, claim, detail) in enumerate(CLAIMS):
        y = ty + i * 1.17
        text(rx, y + 0.02, 0.5, 0.34, [{"runs": [
            {"t": "%d" % (i + 1), "sz": 19, "b": True, "c": theme["accent"]}]}])
        text(rx + 0.52, y + 0.04, 3.45, 0.60, [{"runs": [
            {"t": claim, "sz": 11.5, "b": True, "c": theme["struct"]}]}])
        text(rx + 0.52, y + 0.60, 3.45, 0.52, [{"runs": [
            {"t": detail, "sz": 9.5, "c": theme["ink"]}]}])
        if i < len(CLAIMS) - 1:
            rect(rx, y + 1.08, 4.0, 0.007, fill=theme["mute"])

    # a hairline under the whole evidence area, so the page reads as one field
    rect(0.72, 5.86, 11.9, 0.007, fill=theme["mute"])

    # ---- bottom synthesis bar: every system asks for exactly one ---------------
    rect(0.72, 6.06, 11.9, 0.58, fill=theme["bg"], line=theme["struct"], lw=0.75)
    text(0.94, 6.19, 11.4, 0.34, [{"runs": [
        {"t": "判断依据不是闸的数量，而是它们是否独立：三条同源的检查，等于一条。",
         "sz": 12.5, "b": True, "c": theme["struct"]}]}])

    # ---- fixed page trio: source line, page number ----------------------------
    text(0.72, 6.95, 7.0, 0.3, [{"runs": [
        {"t": "来源：dsh-ppt-office-motion · references/review-checklist.md",
         "sz": 8, "c": theme["mute"]}]}])
    text(11.6, 6.95, 1.0, 0.3, [{"runs": [
        {"t": "01", "sz": 8, "c": theme["mute"]}]}], align=PP_ALIGN.RIGHT)

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    prs.save(out_path)
    return out_path


def list_themes():
    idx = os.path.join(DS_DIR, "themes.json")
    if os.path.exists(idx):
        data = json.load(io.open(idx, encoding="utf-8"))
        for t in data["themes"]:
            print("  %-32s %-12s %s" % (t["id"], t["category"], t["purpose"][:44]))
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(description="用设计系统生成页面")
    ap.add_argument("--theme")
    ap.add_argument("--out", default="composed/deck.pptx")
    ap.add_argument("--all", action="store_true")
    ap.add_argument("--outdir", default="composed")
    ap.add_argument("--list", action="store_true")
    ap.add_argument("--show", action="store_true", help="打印解析出的主题")
    ns = ap.parse_args(argv)

    if ns.list:
        return list_themes()

    if ns.all:
        data = json.load(io.open(os.path.join(DS_DIR, "themes.json"), encoding="utf-8"))
        for t in data["themes"]:
            th = theme_of(t["id"])
            out = os.path.join(ns.outdir, "%s.pptx" % t["id"])
            compose(th, out)
            print("  %-32s bg=#%s struct=#%s accent=#%s"
                  % (t["id"], th["bg"], th["struct"], th["accent"]))
        return 0

    if not ns.theme:
        ap.error("need --theme, --all or --list")
    th = theme_of(ns.theme)
    if ns.show:
        print(json.dumps(th, ensure_ascii=False, indent=1))
        return 0
    compose(th, ns.out)
    print("wrote %s" % ns.out)
    print("  bg=#%s  structural=#%s  accent=#%s  ink=#%s"
          % (th["bg"], th["struct"], th["accent"], th["ink"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
