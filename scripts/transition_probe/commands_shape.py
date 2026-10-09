"""transition_probe.commands_shape — 形态层与方向

每个效果**看起来在做什么**（含方向是否镜像）—— 形态表的数据来源。

模块内容表（6 个命令）
----------------------------------------------------------------
    shapedeck     One deck per effect in HYPOTHESES, for the shape (not the na
    shapedeck2    Deck 2: the noise-texture probe, for the effects deck 1 conf
    shapeanalyze  Report the shape of one rendered transition.
    shapes        Analyse every <dir>/*.mp4 that has a matching manifest.
    dirdeck       One deck per (spec, dir) row in DIR_PROBE.
    dirmirror     Pair up the two dir renders of each spec and judge whether d

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
