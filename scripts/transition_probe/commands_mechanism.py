"""transition_probe.commands_mechanism — 机制层

切换**挂在哪一页**、几个槽位、时长听谁 —— 机制问题，不是效果问题。

模块内容表（2 个命令）
----------------------------------------------------------------
    anchordeck    Two decks that can FALSIFY claim A, plus one that probes cla
    anchors       

> 本模块由 `tools/split_commands.py` 从 `commands.py`（1367 行）按**探测维度**拆出
> （原文照搬，未改逻辑）。对外接口由 `../__init__.py` 重新导出 ——
> 这些 **CLI 子命令名写在 `facts/` 与 `reference/` 里，不能改名。**
"""

from .common import DUR_MS, FPS, GRID_COLS, GRID_ROWS, NS, SLIDE_SECONDS, _extract_transition, _load_json, _norm
from .decks import _anim_deck, _flat_deck, _probe_deck, _shape_deck, _shape_deck2, _slide_xml, _two_slide_deck, _with_dir, wrap_transition
import io
import json
import os


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
