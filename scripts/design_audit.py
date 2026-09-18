#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Audit a motion spec against the design spec's HARD constraints, and say plainly
what it cannot judge.

WHY THIS EXISTS
---------------
`references/motion-design-spec.md` section 6 lists five hard limits and section 8 a
ten-item pre-delivery checklist. Until now both were things a person had to remember
and count by hand, which is exactly the kind of instruction that rots. This turns the
mechanical half into a command.

WHAT IT CAN AND CANNOT DECIDE
-----------------------------
The project has a documented history of automated checks giving confident wrong
answers (references/review-checklist.md section 6, and HANDOVER.md section 6). A design
linter is a fresh opportunity to make that mistake, so the output is split in three:

  FAIL    a mechanical limit is broken. Cheap, reliable, and the rule is written down
          in the spec, so there is nothing to interpret.
  ADVISE  a heuristic fired. Informational only -- it may be exactly what the author
          intended. Never treat one of these as a defect on its own.
  HUMAN   checks that are NOT decidable from the spec alone. Printed every run so the
          tool never looks like it has covered them.

The third list is the point. Nine of the ten items in the design spec's section 8
checklist are judgement calls -- "can you say what this animation directs attention
to", "does it still read with the motion removed", "is the motion direction the same
as the reading direction". A tool that silently skipped them would be worse than no
tool, because a green run would read as approval.

Usage:
    python scripts/design_audit.py --pptx deck.pptx --spec motion.yaml
    python scripts/design_audit.py --pptx deck.pptx --spec motion.yaml --json
"""
import argparse
import json
import os
import re
import sys
from collections import Counter, OrderedDict

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import motion                                          # noqa: E402
import player                                          # noqa: E402

# --- limits, quoted from motion-design-spec.md section 6 ---------------------
# "组" is read as a PowerPoint ANIMATION GROUP, which is the unit that paragraph is
# about ("每页 5 组，超出请合并或删减" = how many things the audience waits through).
#
# The two units are NOT the same number, and conflating them made this linter report
# fourteen false failures on its first run:
#
#   reveals = one per effect. What the audience watches appear, one item at a time.
#   groups  = PowerPoint's click/auto-advance boundaries. A run of `after` effects
#             with no `click` is ONE group that plays through; a `with` never starts
#             a boundary.
#
# The pacing constraint applies to reveals. Groups are reported alongside for
# information, because a page can have few groups and still be a slideshow of
# fourteen separate appearances.
MAX_REVEALS_PER_PAGE = 5
MAX_EFFECT_KINDS_PER_PAGE = 3
MAX_LOOPS_PER_PAGE = 1
MAX_PAGE_SECONDS = 12.0
MAX_TRANSITION_KINDS = 3

# --- timing ladder, section 2 ------------------------------------------------
LADDER = ((0.20, 0.35, "T0"), (0.40, 0.80, "T1"),
          (0.90, 1.50, "T2"), (1.20, 3.00, "T3"))
DELAY_MIN, DELAY_MAX = 0.10, 0.30

# --- the "use with caution" list, section 3 ----------------------------------
DISCOURAGED = {
    "bounce", "boomerang", "pinwheel", "buzzsaw", "checkerboard", "blinds",
    "randomBars", "swivel", "spinner", "flash", "flicker",
}
# Effects whose semantic is "keep running", so a loop on them is expected.
LOOP_FRIENDLY = {
    "pathSineWave", "pathCircle", "pathHorizontalFigure8", "pathVerticalFigure8",
    "pathFigure8Four", "spin", "wave", "shimmer", "pulse",
}
# A shape covering this fraction of the slide or more is treated as a background,
# which is what section 1's "first screen must not be empty" rule is about.
GROUND_COVERAGE = 0.90


def ladder_of(duration):
    for lo, hi, name in LADDER:
        if lo <= duration <= hi:
            return name
    return None


def index_shape_roles(pptx):
    """{page: {shape_id: {"name", "coverage", "text", "cx"}}}

    Geometry comes from player.shape_index, not motion.inspect: inspect reports only
    name/id/tag per shape, with no coordinates, so coverage and horizontal centre have
    to be read from the shape index. Slide size comes from motion.slide_size so the
    coverage fraction is correct for decks that are not 960x540.
    """
    sw, sh_h = motion.slide_size(pptx)
    out = {}
    for page in range(1, len(motion.inspect(pptx).get("slides") or []) + 1):
        per = {}
        for sh in player.shape_index(pptx, page):
            pt = sh.get("pt")
            if not pt:
                continue
            x, y, w, h = pt
            cov = (w * h) / float(sw * sh_h) if sw and sh_h else 0.0
            per[int(sh["id"])] = {"name": sh.get("name", ""),
                                  "coverage": round(cov, 4),
                                  "text": (sh.get("text") or "").strip()[:24],
                                  "cx": x + w / 2.0}
        out[page] = per
    return out


def transition_kind_of(pptx, page):
    """Read the transition actually present in slide XML."""
    import zipfile
    try:
        with zipfile.ZipFile(pptx) as z:
            x = z.read("ppt/slides/slide%d.xml" % page).decode("utf-8", "replace")
    except (KeyError, OSError):
        return None
    if "<p:transition" not in x:
        return None
    if "p159:morph" in x or "<p:morph" in x:
        return "morph"
    m = re.search(r"<p:transition\b[^>]*>\s*<p:(\w+)", x)
    return m.group(1) if m else "unknown"


def audit(pptx, spec_path, reveal_limit=MAX_REVEALS_PER_PAGE):
    spec = motion.load_spec(spec_path)
    entries = motion.normalize_spec(spec)
    roles = index_shape_roles(pptx)
    cat = json.load(open(motion.catalog_path(), encoding="utf-8"))

    findings = []          # (level, code, page, message, ref)
    pages = OrderedDict()

    trans_kinds = Counter()
    for entry in entries:
        page = int(entry["page"])
        effs = entry["effects"]
        sched = motion.schedule_spec(effs)
        page_roles = roles.get(page, {})

        effects = [e for e in effs]
        kinds = Counter(str(e.get("effect")) for e in effects)
        loops = [e for e in effects
                 if int(e.get("repeat") or 0) > 1 or e.get("autoReverse")]

        # A "group" is a click or auto-advance boundary. `with` belongs to the group
        # it accompanies, and a run of `after` is a single group that plays through --
        # so only `click` starts a boundary here. `reveals` counts appearances.
        groups = 1 if effects else 0
        for e in effects:
            if str(e.get("trigger", "after")).lower() == "click":
                groups += 1
        reveals = len(effects)

        total = max((s["start"] + s["duration"] for s in sched), default=0.0)

        info = {"effects": reveals, "groups": groups,
                "kinds": dict(kinds), "loops": len(loops),
                "total_seconds": round(total, 2),
                "transition": transition_kind_of(pptx, page)}
        pages[page] = info
        if info["transition"]:
            trans_kinds[info["transition"]] += 1

        # ---- FAIL: the hard limits. Reveals is the pacing constraint; groups is
        # reported for information only (see the note on MAX_REVEALS_PER_PAGE).
        if reveals > reveal_limit:
            findings.append(("FAIL", "reveals", page,
                             "%d 次揭示，超过上限 %d —— 规范原文是「超过 6 组观众会"
                             "开始等动画而不是听你讲」，逐个出现正是这种等待"
                             % (reveals, reveal_limit), "§6/§1.4"))
        if groups > 5:
            findings.append(("ADVISE", "groups", page,
                             "%d 个点击/自动推进边界，偏多" % groups, "§6"))
        if len(kinds) > MAX_EFFECT_KINDS_PER_PAGE:
            findings.append(("FAIL", "kinds", page,
                             "%d 种效果（%s），超过上限 %d"
                             % (len(kinds), ", ".join(sorted(kinds)),
                                MAX_EFFECT_KINDS_PER_PAGE), "§6"))
        if len(loops) > MAX_LOOPS_PER_PAGE:
            findings.append(("FAIL", "loops", page,
                             "%d 个循环元素，超过上限 %d"
                             % (len(loops), MAX_LOOPS_PER_PAGE), "§6"))
        if total > MAX_PAGE_SECONDS:
            findings.append(("FAIL", "seconds", page,
                             "总时长 %.1fs，超过上限 %.0fs"
                             % (total, MAX_PAGE_SECONDS), "§6"))

        # ---- FAIL: a loop on content rather than decoration
        for e in loops:
            alias = str(e.get("effect"))
            if alias not in LOOP_FRIENDLY:
                findings.append(("FAIL", "loop-on-content", page,
                                 "`%s` 带 repeat/autoReverse，但它是内容型效果；"
                                 "循环只允许给装饰件" % alias, "§2/§4"))
            tid = int(e.get("target")) if str(e.get("target", "")).isdigit() else None
            sh = page_roles.get(tid) if tid is not None else None
            if sh and sh["text"] and alias not in LOOP_FRIENDLY:
                findings.append(("ADVISE", "loop-on-text", page,
                                 "循环落在含文字的 `%s` 上（%r）"
                                 % (sh["name"], sh["text"]), "§2"))

        # ---- FAIL: timing ladder and delay band
        for e in effects:
            d = float(e.get("duration") or 0.5)
            if ladder_of(d) is None:
                findings.append(("ADVISE", "ladder", page,
                                 "时长 %.2fs 不落在任何阶梯档位 "
                                 "(T0 .2-.35 / T1 .4-.8 / T2 .9-1.5 / T3 1.2-3.0)"
                                 % d, "§2"))
            dl = float(e.get("delay") or 0.0)
            if dl and not (DELAY_MIN <= dl <= DELAY_MAX):
                findings.append(("ADVISE", "delay", page,
                                 "delay %.2fs 超出 0.10–0.30s 建议区间" % dl, "§2"))

        # ---- ADVISE: effects the spec says to avoid
        for alias in kinds:
            if alias in DISCOURAGED:
                findings.append(("ADVISE", "discouraged", page,
                                 "`%s` 在慎用清单里" % alias, "§3"))

        # ---- ADVISE: first-screen emptiness (heuristic)
        if effects:
            first = effects[0]
            tid = int(first.get("target")) if str(first.get("target", "")).isdigit() else None
            sh = page_roles.get(tid) if tid is not None else None
            is_ground = bool(sh and sh["coverage"] >= GROUND_COVERAGE)
            if str(first.get("trigger", "")).lower() == "with":
                findings.append(("ADVISE", "with-first", page,
                                 "第一个效果就是 `with`；`with` 只应用于背景层，"
                                 "放首位会退化成 after", "§1"))
            if not is_ground and len(effects) > 1:
                findings.append(("ADVISE", "ground-first", page,
                                 "第一个效果作用在非背景形状上"
                                 "（`%s` 覆盖 %.0f%%）；若非有意，背景应先出现"
                                 % ((sh or {}).get("name", "?"),
                                    100 * (sh or {}).get("coverage", 0)), "§1/§4"))

        # ---- ADVISE: symmetric-looking layout with asymmetric effects
        left = [e for e in effects
                if _centre_x(e, page_roles) is not None
                and _centre_x(e, page_roles) < 480]
        right = [e for e in effects
                 if _centre_x(e, page_roles) is not None
                 and _centre_x(e, page_roles) >= 480]
        if len(left) >= 2 and len(right) >= 2:
            lk = sorted({str(e.get("effect")) for e in left})
            rk = sorted({str(e.get("effect")) for e in right})
            if lk != rk and len(lk) == 1 and len(rk) == 1:
                findings.append(("ADVISE", "symmetry", page,
                                 "左右两栏效果不同（左 %s / 右 %s）；"
                                 "对称结构宜用对称动效" % (lk[0], rk[0]), "§4"))
        if len(left) >= 2 and len(right) >= 2:
            ld = {round(float(e.get("duration") or 0), 2) for e in left}
            rd = {round(float(e.get("duration") or 0), 2) for e in right}
            if ld != rd and len(ld) == 1 and len(rd) == 1:
                findings.append(("ADVISE", "symmetry-time", page,
                                 "左右两栏时长不同（左 %ss / 右 %ss）"
                                 % (list(ld)[0], list(rd)[0]), "§4"))

        # ---- ADVISE: many bars of one kind all animating = a chart with no focus
        same = [k for k, v in kinds.items() if v >= 4]
        if same:
            findings.append(("ADVISE", "chart-focus", page,
                             "同一效果 `%s` 用了 %d 次；图表页宜整块入、只强调一个点"
                             % (same[0], kinds[same[0]]), "§4"))

    if len(trans_kinds) > MAX_TRANSITION_KINDS:
        findings.append(("FAIL", "transitions", None,
                         "全片 %d 种切换（%s），超过上限 %d"
                         % (len(trans_kinds), ", ".join(sorted(trans_kinds)),
                            MAX_TRANSITION_KINDS), "§6"))

    return {"pages": pages, "findings": findings,
            "transition_kinds": dict(trans_kinds)}


def _centre_x(entry, page_roles):
    """Horizontal centre of the target shape, in points. None when unresolvable."""
    t = entry.get("target")
    if not str(t).isdigit():
        return None
    sh = page_roles.get(int(t))
    return (sh or {}).get("cx")


def human_checks():
    """The design spec's section 8 items that NO tool can settle.

    Printed on every run. Nine of the ten are judgement calls; if this list were
    hidden, a clean mechanical pass would look like approval of the whole checklist.
    """
    return [
        ("每处动画能说出它想把注意力引到哪里", "motion-design-spec §8 / review-checklist §4"),
        ("删掉全部动画后静态版面依然成立", "§8 —— 跑 `motion.py preview` 导出去看"),
        ("一页同一时刻只有一个元素抢眼", "§8"),
        ("对称结构用了对称动效（含视觉效果，不只是参数）", "§8"),
        ("图表页只强调一个结论", "§8"),
        ("方向性擦除的方向与阅读方向一致", "§8 —— 结构闸查不出方向错误，必须人眼看"),
        ("时长阶梯全篇统一，没有随手感", "§8"),
        ("切换统一且与章节节奏一致", "§8"),
        ("阅读稿 / 打印稿 / PDF 版已确认不动画，或已生成去动画副本", "§8"),
        ("素材里没有工具界面残留 / 水印 / 占位文字", "review-checklist §2"),
    ]


def main(argv=None):
    ap = argparse.ArgumentParser(description="动效设计规范硬约束执行化")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--reveal-limit", type=int, default=MAX_REVEALS_PER_PAGE,
                    help="每页揭示次数上限（默认 %d，来自 §6；"
                         "确定要密集 build 时可放宽）" % MAX_REVEALS_PER_PAGE)
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)

    rep = audit(ns.pptx, ns.spec, reveal_limit=ns.reveal_limit)

    if ns.json:
        rep["human_checks"] = human_checks()
        print(json.dumps(rep, ensure_ascii=False, indent=1))
        return 0 if not [f for f in rep["findings"] if f[0] == "FAIL"] else 1

    fails = [f for f in rep["findings"] if f[0] == "FAIL"]
    advises = [f for f in rep["findings"] if f[0] == "ADVISE"]

    print("== 动效设计审计 ==")
    print("（reveals = 揭示次数，逐个出现；组 = 点击/自动推进边界。两者不是一个数，"
          "详见脚本头部注释）")
    print()
    print("%-6s %-8s %-6s %-7s %-6s %s"
          % ("页", "reveals", "组", "种类", "循环", "总时长 / 切换"))
    for pg, i in rep["pages"].items():
        print("%-6d %-8d %-6d %-7d %-6d %.1fs / %s"
              % (pg, i["effects"], i["groups"], len(i["kinds"]), i["loops"],
                 i["total_seconds"], i["transition"] or "-"))
    print()
    print("全片切换种类: %d %s"
          % (len(rep["transition_kinds"]),
             "(" + ", ".join(sorted(rep["transition_kinds"])) + ")" or ""))

    if fails:
        print("\n-- FAIL：超过规范 §6 的硬上限 --")
        for lvl, code, pg, msg, ref in fails:
            print("  P%s  %-18s %s   [%s]" % (pg if pg else "-", code, msg, ref))
    if advises:
        print("\n-- ADVISE：启发式提示，可能正是你要的，不要当成缺陷 --")
        for lvl, code, pg, msg, ref in advises:
            print("  P%s  %-18s %s   [%s]" % (pg if pg else "-", code, msg, ref))

    print("\n-- 只能人眼判的（本工具不覆盖，逐条过） --")
    for i, (what, where) in enumerate(human_checks(), 1):
        print("  %2d. %-46s %s" % (i, what, where))

    print("\n结论：%d 项硬约束违规，%d 项启发式提示，%d 项需人工确认。"
          % (len(fails), len(advises), len(human_checks())))
    if not fails:
        print("注意：**这不等于通过** —— 上面那 %d 项没被判过。"
              % len(human_checks()))
    return 1 if fails else 0


if __name__ == "__main__":
    sys.exit(main())
