#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: motion analysis from a video recording.

Reads a screen/export recording of a deck and measures what actually moved:

  * a motion timeline  - when each burst starts/ends, how long it lasts,
                         how long the static gaps between bursts are
  * per-slide splitting - long static runs separate slides
  * movement signatures - bbox / centroid / changed-area trajectories per burst,
                         turned into a probable effect family
  * aggregate statistics - burst count per slide, duration and interval
                         distributions, animated-time share, loop detection

Why this exists: effect *names* cannot be recovered from pixels, but timing,
order, direction and whether something loops can. That is exactly the evidence
the motion design spec needs, and it is far more reliable than guessing from
screenshots.

Usage:
    python analyze_video.py --video clip.mp4 [--out out_dir] [--fps 30]
                            [--json] [--max-slides 0]

Requires: opencv-python (cv2) + numpy.
"""
import argparse
import json
import os
import sys

try:
    import cv2
    import numpy as np
except ImportError as exc:  # pragma: no cover
    print("need opencv-python and numpy: %s" % exc, file=sys.stderr)
    sys.exit(2)


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def to_gray(frame):
    if frame.ndim == 3:
        return cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    return frame


def downscale(gray, width):
    h, w = gray.shape[:2]
    if w <= width:
        return gray, 1.0
    scale = width / float(w)
    return cv2.resize(gray, (width, int(round(h * scale))), interpolation=cv2.INTER_AREA), scale


def bbox_of(mask, min_pixels=4):
    ys, xs = np.nonzero(mask)
    if len(xs) < min_pixels:
        return None
    return int(xs.min()), int(ys.min()), int(xs.max()), int(ys.max())


def hysteresis(signal, high, low):
    """Boolean mask: enter on > high, leave on < low."""
    on = np.zeros(len(signal), dtype=bool)
    active = False
    for i, v in enumerate(signal):
        if active:
            if v < low:
                active = False
        else:
            if v > high:
                active = True
        on[i] = active
    return on


def smooth(signal, k):
    if k <= 1:
        return signal
    kernel = np.ones(k, dtype=np.float64) / k
    return np.convolve(signal, kernel, mode="same")


# --------------------------------------------------------------------------
# activity signal
#
# Measured on real PowerPoint-exported H.264 (see tests/calibration-video.md):
#   * the block mean and any low pixel-diff threshold are dominated by encoder
#     noise -- on a completely static slide they stay at a constant non-zero
#     level, so they cannot separate motion from stillness;
#   * the per-frame MAXIMUM absolute pixel difference does separate them, and it
#     ramps monotonically for the whole length of a fade / move / grow;
#   * a low-contrast animation (small colour delta against the background) can
#     still fall near the noise floor -- those need a high-bitrate recording or
#     the source deck, not the video.
#
# So: signal = max abs diff, threshold = noise floor measured from the quietest
# part of the clip, and we additionally measure "settle time" (how long until the
# frame stops changing after a burst starts), which tracks a fade's real length
# far better than the burst itself.
# --------------------------------------------------------------------------
def _robust_floor(x):
    x = np.asarray(x, dtype=np.float64)
    if x.size == 0:
        return 0.0
    return float(np.percentile(x, 20))


def analyze(video, sample_fps=None, analysis_width=640, min_abs=4.0):
    cap = cv2.VideoCapture(video)
    if not cap.isOpened():
        raise RuntimeError("cv2 could not open %s" % video)

    src_fps = cap.get(cv2.CAP_PROP_FPS) or 30.0
    w = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
    h = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
    step = max(1, int(round(src_fps / sample_fps))) if sample_fps else 1
    eff_fps = src_fps / step

    prev = None
    times, mag, boxes, areas, centroids = [], [], [], [], []
    idx = 0
    while True:
        if not cap.grab():
            break
        if idx % step == 0:
            ok, frame = cap.retrieve()
            if not ok:
                break
            g, _ = downscale(to_gray(frame), analysis_width)
            if prev is not None:
                diff = cv2.absdiff(g, prev)
                times.append(idx / src_fps)
                mag.append(int(diff.max()))
                mask = (diff >= max(8, int(min_abs) * 2)).astype(np.uint8)
                bb = bbox_of(mask)
                boxes.append(bb)
                areas.append(int(mask.sum()))
                if bb:
                    x0, y0, x1, y1 = bb
                    centroids.append(((x0 + x1) / 2.0, (y0 + y1) / 2.0))
                else:
                    centroids.append(None)
            prev = g
        idx += 1
    cap.release()

    if not mag:
        raise RuntimeError("no frames analysed")

    raw = np.asarray(mag, dtype=np.float64)
    n = len(raw)
    sm = smooth(raw, max(1, int(round(eff_fps * 0.05))))
    floor = _robust_floor(raw)
    # Headroom must ignore hard cuts (slide changes spike to ~255), otherwise one
    # cut inflates every threshold and subtle animations disappear. Use a high
    # percentile of the non-cut population instead of the raw maximum.
    body = raw[raw < 0.6 * raw.max()] if raw.max() > 0 else raw
    headroom = float(np.percentile(body, 98)) if body.size else float(raw.max())
    span = max(1.0, headroom - floor)
    high = floor + max(min_abs, 0.30 * span)
    low = floor + max(min_abs * 0.5, 0.12 * span)
    active = hysteresis(sm, high, low)
    # a hard cut is a single-frame spike far above the animation band: mark it so
    # it is reported as a slide change rather than as an animation
    cut_level = max(high * 3.0, 0.55 * raw.max()) if raw.max() > 0 else high * 3.0
    is_cut = raw >= cut_level

    # ---- bursts ----
    bursts = []
    i = 0
    while i < n:
        if not active[i]:
            i += 1
            continue
        j = i
        while j + 1 < n and active[j + 1]:
            j += 1
        bursts.append([i, j])
        i = j + 1

    merged = []
    for b in bursts:
        if merged and (times[b[0]] - times[merged[-1][1]]) < 0.08:
            merged[-1][1] = b[1]
        else:
            merged.append(list(b))
    bursts = merged

    events = []
    for bi, (s, e) in enumerate(bursts):
        # settle: keep extending while the signal stays clearly above the floor
        st = e
        settle_limit = floor + max(min_abs, 0.18 * span)
        while st + 1 < n and raw[st + 1] > settle_limit:
            st += 1
        dur = times[st] - times[s] + (1.0 / eff_fps)
        bb_list = [boxes[k] for k in range(s, st + 1) if boxes[k]]
        area_list = [areas[k] for k in range(s, st + 1)]
        cent_list = [centroids[k] for k in range(s, st + 1) if centroids[k] is not None]
        peak_i = int(np.argmax(raw[s:st + 1])) + s
        cut = bool(is_cut[s:st + 1].any())

        ev = {
            "index": bi + 1,
            "start": round(times[s], 3),
            "end": round(times[st] + 1.0 / eff_fps, 3),
            "duration": round(dur, 3),
            "peak_time": round(times[peak_i], 3),
            "peak_magnitude": int(raw[peak_i]),
            "snr": round(float(raw[peak_i]) / max(1.0, floor), 2),
            "changed_px_peak": int(max(area_list)) if area_list else 0,
            "slide_change": cut,
        }

        if bb_list:
            x0 = min(b[0] for b in bb_list); y0 = min(b[1] for b in bb_list)
            x1 = max(b[2] for b in bb_list); y1 = max(b[3] for b in bb_list)
            ev["bbox"] = [x0, y0, x1, y1]
            ev["bbox_span"] = [x1 - x0, y1 - y0]
            if cent_list:
                xs = [c[0] for c in cent_list]; ys = [c[1] for c in cent_list]
                ev["centroid_travel"] = [round(max(xs) - min(xs), 1), round(max(ys) - min(ys), 1)]
                ev["centroid_net"] = [round(xs[-1] - xs[0], 1), round(ys[-1] - ys[0], 1)]
                rev = 0
                for k in range(2, len(xs)):
                    d1 = xs[k - 1] - xs[k - 2]
                    d2 = xs[k] - xs[k - 1]
                    if d1 * d2 < 0:
                        rev += 1
                ev["x_reversals"] = rev
            # growth must use the PEAK area, not the last frame: by the time a
            # fade finishes the frame-to-frame diff has decayed back to zero, so
            # last/first would report ~0 even for a large reveal
            ev["area_growth"] = (round(float(max(area_list)) / max(1, area_list[0]), 2)
                                 if area_list else None)

            span_w = x1 - x0
            travel = ev.get("centroid_travel", [0, 0])
            net = ev.get("centroid_net", [0, 0])
            growth = ev.get("area_growth") or 1.0
            revs = ev.get("x_reversals", 0)

            # Loop detection is deliberately strict (>=3 direction reversals with a
            # small net displacement AND little area growth): a slow fade jitters
            # the mask centroid and would otherwise be misread as an oscillation.
            if ev["snr"] < 1.6 or dur <= 5.0 / max(1.0, eff_fps):
                guess, why = "cut / encoder artifact (ignore)", "too short or too close to the noise floor to be an animation"
            elif (dur >= 1.0 and revs >= 3 and growth < 1.5
                  and max(abs(net[0]), abs(net[1])) < 0.35 * max(1, span_w)):
                guess, why = "loop / oscillate", "centroid reverses >=3x with small net travel and no growth"
            # growth is checked before translation: a fade-in reveals a shape whose
            # changed-area grows from ~0, which also drags the centroid around and
            # would otherwise be misread as a translate
            elif growth > 2.5:
                guess, why = "fade / zoom / grow (area grows)", "changed area grows >2.5x over the burst"
            elif max(abs(net[0]), abs(net[1])) > 0.25 * max(1, span_w):
                guess, why = "fly / push (translate)", "centroid travels a long way in one direction"
            elif growth < 0.6:
                guess, why = "shrink / exit", "changed area shrinks over the burst"
            else:
                guess, why = "fade / wipe (in place)", "area roughly stable, no large translation"
            ev["probable"] = guess
            ev["probable_because"] = why
        events.append(ev)

    total_span = times[-1] - times[0] + (1.0 / eff_fps) if times else 0.0
    animated = float(np.sum(active)) / eff_fps if len(active) else 0.0

    # ---- split into slides: a static run longer than `slide_gap` ----
    slide_gap = 1.2
    slide_bounds = []
    cur_start = times[0]
    k = 0
    while k < n:
        if not active[k]:
            j = k
            while j + 1 < n and not active[j + 1]:
                j += 1
            if times[j] - times[k] >= slide_gap:
                slide_bounds.append((cur_start, times[k]))
                cur_start = times[j]
            k = j + 1
        else:
            k += 1
    slide_bounds.append((cur_start, times[-1] + 1.0 / eff_fps))

    slides = []
    for si, (s0, s1) in enumerate(slide_bounds, start=1):
        inner = [e for e in events if s0 - 1e-6 <= e["start"] <= s1 + 1e-6]
        slides.append({
            "slide": si,
            "start": round(s0, 3),
            "end": round(s1, 3),
            "span": round(s1 - s0, 3),
            "motion_count": len(inner),
            "animated_time": round(sum(e["duration"] for e in inner), 3),
            "effects": [{"start": e["start"], "duration": e["duration"],
                         "probable": e.get("probable"), "bbox": e.get("bbox")} for e in inner],
        })

    durs = [e["duration"] for e in events]
    gaps = []
    for a, b in zip(events, events[1:]):
        gaps.append(round(b["start"] - a["end"], 3))

    types = {}
    for e in events:
        types[e.get("probable", "?")] = types.get(e.get("probable", "?"), 0) + 1

    overview = {
        "video": os.path.abspath(video),
        "size": [w, h],
        "source_fps": round(src_fps, 3),
        "analysis_fps": round(eff_fps, 3),
        "frames_analysed": n,
        "video_seconds": round(total_span, 3),
        "motion_bursts": len(events),
        "bursts_per_slide_median": _median([s["motion_count"] for s in slides]),
        "duration_median": _median(durs),
        "duration_min": min(durs) if durs else None,
        "duration_max": max(durs) if durs else None,
        "gap_median": _median([g for g in gaps if g > 0]) if any(g > 0 for g in gaps) else None,
        "animated_time": round(animated, 3),
        "animated_share": round(animated / total_span, 3) if total_span else None,
        "probable_types": types,
    }
    return {"overview": overview, "slides": slides, "events": events,
            "timeline": {"t": [round(t, 3) for t in times],
                         "activity": [int(v) for v in raw],
                         "active": [bool(v) for v in active]}}


def _median(xs):
    xs = [x for x in xs if x is not None]
    if not xs:
        return None
    s = sorted(xs)
    m = len(s) // 2
    v = s[m] if len(s) % 2 else (s[m - 1] + s[m]) / 2.0
    return round(v, 3)


# --------------------------------------------------------------------------
# markdown report
# --------------------------------------------------------------------------
def to_markdown(res):
    o = res["overview"]
    L = []
    L.append("# 动效分析报告")
    L.append("")
    L.append("视频：`%s`" % o["video"])
    L.append("")
    L.append("分辨率 %sx%s · 源 %s fps · 分析 %s fps · 时长 %ss"
             % (o["size"][0], o["size"][1], o["source_fps"], o["analysis_fps"], o["video_seconds"]))
    L.append("")
    L.append("## 总览")
    L.append("")
    L.append("| 指标 | 值 |")
    L.append("| --- | --- |")
    L.append("| 动效个数 | **%d** |" % o["motion_bursts"])
    L.append("| 每页动效数（中位） | %s |" % o["bursts_per_slide_median"])
    L.append("| 单次时长 中位 / 最短 / 最长 | %s / %s / %s s"
             % (o["duration_median"], o["duration_min"], o["duration_max"]))
    L.append("| 动效间隔（中位） | %s s |" % o["gap_median"])
    L.append("| 有动作的时间占比 | %s |" % o["animated_share"])
    L.append("| 疑似类型分布 | %s |" % ", ".join("%s×%d" % (k, v) for k, v in sorted(o["probable_types"].items())))
    L.append("")
    L.append("## 分页时间线")
    L.append("")
    L.append("| 页 | 起 | 止 | 时长 | 动效数 | 动画时长 |")
    L.append("| --- | --- | --- | --- | --- | --- |")
    for s in res["slides"]:
        L.append("| %d | %.2f | %.2f | %.2f | %d | %.2f |"
                 % (s["slide"], s["start"], s["end"], s["span"], s["motion_count"], s["animated_time"]))
    L.append("")
    L.append("## 逐个动效")
    L.append("")
    L.append("| # | 起始 | 时长 | 疑似效果 | 依据 | 变化区域 (x0,y0,x1,y1) | 质点位移 |")
    L.append("| --- | --- | --- | --- | --- | --- | --- |")
    for e in res["events"]:
        L.append("| %d | %.2fs | %.2fs | %s | %s | %s | %s |"
                 % (e["index"], e["start"], e["duration"], e.get("probable", "-"),
                    e.get("probable_because", "-"),
                    e.get("bbox", "-"), e.get("centroid_travel", "-")))
    L.append("")
    L.append("> 说明：`疑似效果` 由像素轨迹推断，**不是** PowerPoint 里的真实效果名。")
    L.append("> 时长、间隔、顺序、位移方向是直接测量值，可以放心使用。")
    L.append("")
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="measure motion in a recorded deck")
    ap.add_argument("--video", required=True)
    ap.add_argument("--out")
    ap.add_argument("--fps", type=float, default=None, help="analysis fps (default: source fps)")
    ap.add_argument("--width", type=int, default=640, help="analysis downscale width")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args()

    res = analyze(ns.video, sample_fps=ns.fps, analysis_width=ns.width)

    if ns.out:
        os.makedirs(ns.out, exist_ok=True)
        p = os.path.join(ns.out, "motion-analysis.json")
        with open(p, "w", encoding="utf-8") as fh:
            json.dump(res, fh, ensure_ascii=False, indent=1)
        m = os.path.join(ns.out, "motion-analysis.md")
        with open(m, "w", encoding="utf-8") as fh:
            fh.write(to_markdown(res))
        print("wrote %s" % p)
        print("wrote %s" % m)

    if ns.json:
        print(json.dumps(res["overview"], ensure_ascii=False, indent=1))
    else:
        print(to_markdown(res))
    return 0


if __name__ == "__main__":
    sys.exit(main())
