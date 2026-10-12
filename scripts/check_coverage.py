#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: spec coverage audit.

The other gates (apply --assert-geometry, verify_motion, motion.ps1 -Strict) prove
the OOXML is legal, the geometry did not move, and PowerPoint kept every effect.
They CANNOT see the failures this catches.

Two failure modes, one per kind of slide:

**Plain slide (no morph transition)** -- the classic one:

    A shape with no effect is visible from the moment the slide appears.

It never disappears and nothing errors -- it just sits there while everything else
animates in. Users report it as "this column keeps showing up and never animates".
This was the single most common defect in real use: a hand-written spec silently
lost an entire table column, a whole slide column, and a whole card because shape
ids had been reassigned by a later layout edit.

**Morph slide** -- the rule is different, and it took a lab measurement to learn
it (2026-10-12, see examples/solar-motion.yaml for the frame evidence):

    On a morph page, page-internal entrance animations do NOT blend with the
    morph. A shape that carries one drops out of the interpolated transition
    entirely: it fades out, then re-enters after the morph ends. So the correct
    design is NO effects on the page at all, letting the transition carry the
    whole performance -- and the failure to look for is no longer "no effect"
    but "no cross-page difference":

      * identical geometry AND identical text as the previous page
        -> the shape will not move at all during the morph ("static"); and
      * no shape with the same id on the previous page
        -> it can only fade in/out, never interpolate ("unpaired").

    Shapes fully OFF the canvas on BOTH pages are skipped: nobody can see them,
    so their being static is not a defect.

Also reports:
  * targets that duplicate a shape (one effect per shape -- com-pitfalls 14), and
  * targets that do not exist on their slide.

Run this on every spec. A non-zero exit means the deck is not ready to present.
"""
import argparse
import json
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import motion  # noqa: E402
import player  # noqa: E402

# Shapes that are legitimately static on every slide: the footer credit and the
# page number. These are defaults for convenience -- pass `--footer` with your own
# text, or an empty value to disable the text match entirely. The list is
# deliberately generic: a tool should not hard-code one project's wording.
DEFAULT_STATIC = ("汇报", "Report", "Confidential", "Confidential - Do Not Distribute")


def is_static(sh, footer_texts, bottom_y=None, right_x=None):
    """Footer credit and page number are deliberately static.

    The page number sits at the bottom-RIGHT and starts *above* the very bottom
    strip (506pt on a 540pt slide), so a bottom-only test misses it. Match on
    position AND content: a bare number in the bottom-right corner.
    """
    t = (sh.get("text") or "").strip()
    if any(k and k in t for k in footer_texts):
        return True
    if t.isdigit():
        pt = sh.get("pt")
        if pt and bottom_y is not None and right_x is not None:
            if pt[1] >= bottom_y and pt[0] >= right_x:
                return True
    return False


def is_morph_slide(pptx, page):
    """该页**实际的**切换是不是 morph —— 读包里的 slide XML，不看 spec 的意图。

    判据要核验产物：spec 里写了 morph 不等于包里真有（apply 可能被跳过，
    或者 deck 后来被人手工改过）。morph 以 p159 扩展元素出现，
    mc:AlternateContent 里带 `p159:morph`（见 reference/svg-in-ppt.md §3）。
    """
    try:
        with zipfile.ZipFile(pptx) as z:
            xml = z.read("ppt/slides/slide%d.xml" % page).decode("utf-8", "replace")
    except (KeyError, OSError):
        return False
    return "p159:morph" in xml


def off_canvas(sh, w, h):
    """形状是否**完全**在画布外（与画布无交集）。

    这是「挪出画布」退场手法的判据基础：完全在外 = 观众看不到它，
    因此它在相邻两页之间「静止」不构成缺陷。没有几何信息时保守返回 False
    （当作在画布内，参与判定）。
    """
    pt = sh.get("pt")
    if not pt:
        return False
    x, y, cx, cy = pt
    return (x + cx <= 0) or (x >= w) or (y + cy <= 0) or (y >= h)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Audit motion spec coverage")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--footer", action="append", default=None,
                    help="extra footer/static text marker (repeatable)")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    spec = motion.load_spec(ns.spec)
    entries = {e["page"]: e for e in motion.normalize_spec(spec)}
    footer_texts = tuple(ns.footer) if ns.footer else DEFAULT_STATIC

    # page-number zone: bottom strip, right-hand side (points)
    try:
        w, h = motion.slide_size(ns.pptx)
    except Exception:
        w, h = 960.0, 540.0
    bottom_y, right_x = h * 0.90, w * 0.85

    result = {"missing": {}, "static": {}, "unpaired": {},
              "duplicate": {}, "unknown": {}, "ok": True}
    morph_pages = 0
    for page in sorted(entries):
        shapes = player.shape_index(ns.pptx, page)
        covered, dupes = set(), []
        for eff in entries[page]["effects"]:
            t = str(eff.get("target"))
            if t in covered:
                dupes.append(t)
            covered.add(t)
        # spec 的 target 写的是 shape **name**（authoring-rules §A1：
        # "不要手写 id"），所以这里必须按 name 判存在性。
        # 原来写成 ids = {sh["id"]} —— 拿数字 id 去比 name，交集恒为空，
        # 于是每一页每一个 target 都被误报成 "not on slide"。
        # verify_motion 明明报 109 个效果全在，两边矛盾时先怀疑判据（§4.1）。
        # id 仍然接受：有人确实会写 id，spec 里两种都支持。
        names = {sh["name"] for sh in shapes}
        idents = {sh["id"] for sh in shapes if sh["id"]}
        unknown = sorted(t for t in covered if t not in names and t not in idents)

        # 第 1 页没有"上一页"，morph 无从配对 —— 它只能走 timing 那条路。
        if page > 1 and is_morph_slide(ns.pptx, page):
            # ── morph 页面：判「跨页差异」，不判「timing 覆盖」 ─────────
            # 实测结论：morph 页面上的页内 entrance 动画会把形状踢出补间
            # （先淡出、morph 完再重新入场）。所以这里要的是**差异**。
            morph_pages += 1
            prev = {sh["id"]: sh for sh in player.shape_index(ns.pptx, page - 1)}
            static, unpaired = [], []
            for sh in shapes:
                if is_static(sh, footer_texts, bottom_y, right_x):
                    continue
                prv = prev.get(sh["id"])
                if prv is None:
                    unpaired.append({"id": sh["id"],
                                     "label": (sh["text"][:40] or sh["name"]),
                                     "pt": sh["pt"]})
                    continue
                if off_canvas(sh, w, h) and off_canvas(prv, w, h):
                    continue          # 两页都在画布外：看不见，静止无妨
                same = (sh["pt"] == prv["pt"]
                        and (sh["text"] or "") == (prv["text"] or ""))
                if same:
                    static.append({"id": sh["id"],
                                   "label": (sh["text"][:40] or sh["name"]),
                                   "pt": sh["pt"]})
            if static:
                result["static"][str(page)] = static
            if unpaired:
                result["unpaired"][str(page)] = unpaired
        else:
            # ── 普通页面：每个可见形状都要有 timing 覆盖（原有判据） ────
            missing = []
            for sh in shapes:
                if sh["name"] in covered or (sh["id"] and sh["id"] in covered):
                    continue
                if is_static(sh, footer_texts, bottom_y, right_x):
                    continue
                missing.append({"id": sh["id"],
                                "label": (sh["text"][:40] or sh["name"]),
                                "pt": sh["pt"]})
            if missing:
                result["missing"][str(page)] = missing

        if dupes:
            result["duplicate"][str(page)] = sorted(set(dupes))
        if unknown:
            result["unknown"][str(page)] = unknown

    result["ok"] = not any(result[k] for k in
                           ("missing", "static", "unpaired",
                            "duplicate", "unknown"))

    if ns.json:
        print(json.dumps(result, ensure_ascii=False, indent=1))
        return 0 if result["ok"] else 1

    for page, items in sorted(result["missing"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s %d shape(s) with NO timing effect:" % (page, len(items)))
        for it in items:
            print("      id=%-5s %-44s %s" % (it["id"], it["label"], it["pt"]))
    for page, items in sorted(result["static"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s %d shape(s) IDENTICAL to previous page (no morph difference):"
              % (page, len(items)))
        for it in items:
            print("      id=%-5s %-44s %s" % (it["id"], it["label"], it["pt"]))
    for page, items in sorted(result["unpaired"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s %d shape(s) with NO same-id shape on previous page "
              "(morph can only fade them):" % (page, len(items)))
        for it in items:
            print("      id=%-5s %-44s %s" % (it["id"], it["label"], it["pt"]))
    for page, ts in sorted(result["duplicate"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s duplicate effect on: %s" % (page, ", ".join(ts)))
    for page, ts in sorted(result["unknown"].items(), key=lambda kv: int(kv[0])):
        print("P%-3s target(s) not on slide: %s" % (page, ", ".join(ts)))

    if result["ok"]:
        n = sum(len(e["effects"]) for e in entries.values())
        print("OK: %d slide(s) (%d morph), %d effect(s); every plain slide's shape "
              "has an effect, every morph slide's shape differs from its "
              "predecessor" % (len(entries), morph_pages, n))
        return 0
    print("\nFAILED. On a plain slide a shape with no effect is visible before the "
          "animation starts;\non a morph slide a shape with no cross-page "
          "difference never moves at all.\nFix the spec (or the generator that "
          "produced it) and re-run.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
