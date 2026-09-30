"""transition_probe.analysis — 帧分析与表构造

**帧分析与表构造** —— 读渲染结果、算指标、合并成表
"""

# 本模块由 tools/split_transition_probe.py 从单文件的 build_transition_table.py
# 搬出来（原文照搬，未改逻辑）。对外接口由 __init__.py 重新导出。


from .common import DUR_MS, FPS, GRID_COLS, GRID_ROWS, PROBE_SPD, SLIDE_SECONDS
from .data import DEFAULT_ATTRS, DIR_PROBE

import re

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


def _dir_pairs(probe=None):
    """Group DIR_PROBE rows into {spec: {axis, dirs, why}} for pairing."""
    out = {}
    for spec, d, axis, why in (probe or DIR_PROBE):
        e = out.setdefault(spec, {"axis": axis, "dirs": [], "why": []})
        e["dirs"].append(d)
        e["why"].append("%s: %s" % (d, why))
    return out


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
