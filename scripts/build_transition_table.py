#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build the slide-transition reference table（入口）。

**本文件现在只是个入口。** 2026-09-30 之前它是一份 3048 行的单文件，
现在实体在 `transition_probe/` 包里（common / data / decks / analysis /
commands 五层）。拆的理由不是行数 —— 是它同时装着"探测数据"、
"deck 构造"、"帧分析"、"CLI 编排"四件不同的事。

**为什么用入口文件而不是直接删掉**：19 个 CLI 子命令写在
`facts/transitions.json` 与 `reference/*.md` 里（`regenerate_with`），
`tests/test_transition_table.py` 还写着 `import build_transition_table as B`。
**名字是接口** —— 保留它，所有既有调用一行都不用改。

Subcommands: 见 `python build_transition_table.py --help`
"""
import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# 重新导出整个包：`B.HYPOTHESES`、`B._profile_metrics` 等照旧可用
from transition_probe import *  # noqa: F401,F403
from transition_probe import __all__ as _pkg_all

# 显式再绑一次，让静态检查与 `from ... import X` 都能看见
from transition_probe.common import NS, NSDECL, DUR_MS, PROBE_SPD, SLIDE_SECONDS, FPS, _extract_transition, _load_json, _norm, GRID_COLS, GRID_ROWS  # noqa: F401
from transition_probe.data import HYPOTHESES, MOTION_SPECS, RULES, DEFAULT_ATTRS, DECK2_SPECS, DIR_PROBE, ATTR_PROBE, TIMING_ANIM, TIMING_PROBE  # noqa: F401
from transition_probe.decks import _with_dir, wrap_transition, _slide_xml, _two_slide_deck, _flat_deck, _shape_deck, _shape_deck2, _set_attrs, _probe_deck, _anim_deck  # noqa: F401
from transition_probe.analysis import _child_tag, _child_attrs, _child_render, _lookup_enum, _template, _dir_pairs, _profile_metrics, mirror_verdict, _read_frames, _window, _energy_trace, _hot_runs, _entry_trace, _timing_chart  # noqa: F401
from transition_probe.commands import cmd_build, cmd_collect, cmd_enumdeck, cmd_enumread, cmd_table, cmd_anchordeck, cmd_anchors, cmd_video, cmd_sheets, cmd_shapedeck2, cmd_shapedeck, cmd_attrdeck, cmd_timingdeck, cmd_dirdeck, cmd_dirmirror, cmd_attrdiff, cmd_timingdiff, cmd_shapeanalyze, cmd_shapes  # noqa: F401


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
