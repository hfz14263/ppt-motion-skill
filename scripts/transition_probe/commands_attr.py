"""transition_probe.commands_attr — 属性取值全集

同一元素换属性值的后果（**不看方向** —— spokes/pattern 本来就不是方向）。

模块内容表（2 个命令）
----------------------------------------------------------------
    attrdeck      One deck per ATTR_PROBE row -- the "attribute value set" pro
    attrdiff      Compare the attribute variants of each spec: does the attrib

> 本模块由 `tools/split_commands.py` 从 `commands.py`（1367 行）按**探测维度**拆出
> （原文照搬，未改逻辑）。对外接口由 `../__init__.py` 重新导出 ——
> 这些 **CLI 子命令名写在 `facts/` 与 `reference/` 里，不能改名。**
"""

from .common import DUR_MS, FPS, GRID_COLS, GRID_ROWS, NS, SLIDE_SECONDS, _extract_transition, _load_json, _norm
from .data import ATTR_PROBE, DECK2_SPECS, DIR_PROBE, HYPOTHESES, MOTION_SPECS, RULES, TIMING_ANIM, TIMING_PROBE
from .decks import _anim_deck, _flat_deck, _probe_deck, _shape_deck, _shape_deck2, _slide_xml, _two_slide_deck, _with_dir, wrap_transition
from .analysis import _child_attrs, _child_tag, _dir_pairs, _energy_trace, _entry_trace, _hot_runs, _lookup_enum, _profile_metrics, _read_frames, _template, _timing_chart, _window, mirror_verdict
import io
import json
import os


def cmd_attrdeck(out_dir):
    """One deck per ATTR_PROBE row -- the "attribute value set" probe."""
    rows = [(s, lab, ov, why) for (s, lab, ov, why) in ATTR_PROBE]
    written = _probe_deck(out_dir, rows, "attr", lambda e: e["label"])
    print("attrdeck -> %s" % out_dir)
    print("  %d 份 deck（ATTR_PROBE：属性取值全集）" % len(written))
    for e in written:
        print("    %-22s %-46s %s" % (e["deck"], e["child"], e["why"]))
    return 0


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
