"""transition_probe.commands_timing — 切换 × 页内动画

两者同页时的结构平行与时间串行（页内动画排队等切换演完）。

模块内容表（2 个命令）
----------------------------------------------------------------
    timingdeck    One deck per TIMING_PROBE row: transition + timing on the sa
    timingdiff    Does <p:transition> coexist with <p:timing> -- and if so, WH

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
import re
import zipfile


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
