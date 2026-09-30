"""transition_probe.commands — CLI 子命令

**CLI 子命令** —— 编排上面的模块。这些名字写在 facts/ 与 reference/ 里，**不能改名**
（`facts/transitions.json` 的 `regenerate_with`、`reference/transitions.md` 等
都按名字指过来；改名等于让那些"怎么复现"的说明失效）。

模块内容表（19 个命令，按它服务的那一层分组；顺序与文件内一致）
----------------------------------------------------------------
**建表主链路**（名字层 48 项实测表）
    build       每个候选一份 2 页 deck + manifest
    collect     （跑完 probe_transitions.ps1 后）逐份 diff
    video       合成一份 deck 供渲染
    sheets      接触表，供目视复核

**枚举扫描**（验证 COM 枚举不可用）
    enumdeck    一页一个 PpEntryEffect，批量扫
    enumread    读回 PowerPoint 自己写了什么

**合并成表**
    table       把两次测量合并成发布的 reference 文件

**机制层**（切换挂在哪页、几个槽位、时长听谁）
    anchordeck  第 2 页写切换，看 burst 落在哪个边界
    anchors     从渲染帧里找真正的变化点

**形态层**（每个效果看起来在做什么）
    shapedeck   48 份网格+移动标记（deck1）
    shapedeck2  12 份满屏噪点（deck2）—— 有些效果 deck1 分不开
    shapeanalyze / shapes   逐份细看 / 合并成形态表

**方向**（`dir` 换值会不会镜像）
    dirdeck     生成方向变体 deck
    dirmirror   配对判镜像（两路判据：变化重心 + 累积位移）

**属性取值全集**
    attrdeck    26 份 deck，量"同一元素换属性值"
    attrdiff    逐帧像素差判定（**不看方向** —— spokes/pattern 本来就不是方向）

**切换 × 页内动画**
    timingdeck  deck 生成
    timingdiff  三答案：结构（两块都在吗）/ 时间（延迟多少）/ 干扰（组合 vs 单独）
"""

# 本模块由 tools/split_transition_probe.py 从单文件的 build_transition_table.py
# 搬出来（原文照搬，未改逻辑）。对外接口由 __init__.py 重新导出。


from .common import DUR_MS, FPS, GRID_COLS, GRID_ROWS, NS, SLIDE_SECONDS, _extract_transition, _load_json, _norm
from .data import ATTR_PROBE, DECK2_SPECS, DIR_PROBE, HYPOTHESES, MOTION_SPECS, RULES, TIMING_ANIM, TIMING_PROBE
from .decks import _anim_deck, _flat_deck, _probe_deck, _shape_deck, _shape_deck2, _slide_xml, _two_slide_deck, _with_dir, wrap_transition
from .analysis import _child_attrs, _child_tag, _dir_pairs, _energy_trace, _entry_trace, _hot_runs, _lookup_enum, _profile_metrics, _read_frames, _template, _timing_chart, _window, mirror_verdict

import glob
import io
import json
import os
import re
import zipfile

def cmd_build(case_dir):
    os.makedirs(case_dir, exist_ok=True)
    for f in glob.glob(os.path.join(case_dir, "*.pptx")):
        os.remove(f)
    cases = []
    for i, h in enumerate(HYPOTHESES):
        name = "cand%02d-%s.pptx" % (i + 1, h["spec"])
        _two_slide_deck(os.path.join(case_dir, name), str(i + 1), h)
        cases.append(name)
    for i, h in enumerate(MOTION_SPECS):
        name = "mspec%02d-%s.pptx" % (i + 1, h["spec"])
        _two_slide_deck(os.path.join(case_dir, name), "m%d" % (i + 1), h)
        cases.append(name)
    manifest = {"hypotheses": HYPOTHESES, "motion_specs": MOTION_SPECS,
                "dur_ms": DUR_MS, "slide_seconds": SLIDE_SECONDS,
                "cases": cases}
    json.dump(manifest, io.open(os.path.join(case_dir, "manifest.json"), "w",
                                encoding="utf-8"), ensure_ascii=False, indent=1)
    print("built %d case decks (%d gallery + %d motion spec) -> %s" % (
        len(cases), len(HYPOTHESES), len(MOTION_SPECS), case_dir))
    return 0


def cmd_collect(case_dir, out_dir, json_out, enum_path=None):
    man = _load_json(os.path.join(case_dir, "manifest.json"))
    # probe_transitions.ps1 writes {"decks":[{deck,opened,slides:[...]}, ...]}.
    # Three distinct outcomes per deck, and they must NOT be merged:
    #   not in enum.json at all  -> UNPROBED   (the probe never reached it)
    #   opened=false             -> REJECTED   (PowerPoint refused the file)
    #   present in OutDir        -> compare wrote vs readback
    decks = {}
    for cand in filter(None, [enum_path, os.path.join(out_dir, "enum.json"),
                              os.path.join(os.path.dirname(out_dir.rstrip("\\/")),
                                           "enum.json")]):
        if not os.path.exists(cand):
            continue
        for d in _load_json(cand).get("decks", []):
            decks[d["deck"]] = d
        break
    rows = []
    probed = [(h, "cand%02d-%s.pptx" % (i + 1, h["spec"]), "gallery")
              for i, h in enumerate(man["hypotheses"])]
    probed += [(h, "mspec%02d-%s.pptx" % (i + 1, h["spec"]), "motion")
               for i, h in enumerate(man.get("motion_specs", []))]
    for h, name, source in probed:
        case = os.path.join(case_dir, name)
        done = os.path.join(out_dir, name)
        d = decks.get(name)
        base = dict(ui=h["ui"], group=h["group"], en=h["en"],
                    family=h["family"], spec=h["spec"], source=source)
        if d is not None and not d.get("opened", True):
            rows.append(dict(base, verdict="REJECTED", wrote="", readback="",
                             error=d.get("error", "")))
            continue
        if not os.path.exists(done):
            rows.append(dict(base, verdict="UNPROBED", wrote="", readback=""))
            continue
        wrote = _extract_transition(zipfile.ZipFile(case).read(
            "ppt/slides/slide2.xml").decode("utf-8"))
        readback = _extract_transition(zipfile.ZipFile(done).read(
            "ppt/slides/slide2.xml").decode("utf-8"))
        e = {}
        if d:
            e = {int(s["slide"]): s for s in d.get("slides", [])}.get(2, {})
        same = _norm(wrote) == _norm(readback)
        rows.append(dict(base, wrote=wrote, readback=readback,
                         preserved=same,
                         entryEffect=e.get("entryEffect"),
                         comDuration=e.get("duration"),
                         verdict=("preserved" if same else
                                  "rewritten" if readback else "dropped")))
    report = {"measured_on": "PowerPoint COM roundtrip (open + SaveAs)",
              "rows": rows,
              "summary": {
                  "total": len(rows),
                  "preserved": sum(1 for r in rows
                                   if r["verdict"] == "preserved"),
                  "rewritten": sum(1 for r in rows if r["verdict"] == "rewritten"),
                  "dropped": sum(1 for r in rows if r["verdict"] == "dropped"),
                  "rejected": sum(1 for r in rows if r["verdict"] == "REJECTED"),
                  "unprobed": sum(1 for r in rows if r["verdict"] == "UNPROBED"),
              }}
    if json_out:
        json.dump(report, io.open(json_out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    s = report["summary"]
    print("== 切换候选 往返结果 ==")
    print("  候选 %d | 原样保留 %d | 被改写 %d | 被丢弃 %d | 拒开 %d | 未测 %d" % (
        s["total"], s["preserved"], s["rewritten"], s["dropped"],
        s["rejected"], s["unprobed"]))
    for r in rows:
        if r["verdict"] == "preserved":
            continue
        rb = r["readback"] or (r.get("error") or "(PowerPoint 写回时没了)")
        print("  [%-8s] %-6s %-6s %s" % (r["verdict"][:8], r["ui"],
                                         r["family"], rb[:110]))
    if json_out:
        print("  -> %s" % json_out)
    return 0


# --------------------------------------------------------------------------
# enumdeck / enumread: let PowerPoint write the transitions ITSELF
# --------------------------------------------------------------------------
# Roundtrip probing settles "is this child element valid", but it cannot invent
# a NAME we guessed wrong -- a wrong name comes back as "refused", which tells
# us nothing about the right one. For those, stop guessing and let PowerPoint
# write: put one slide per PpEntryEffect value in a deck, set every slide's
# EntryEffect over COM, save, and read back what element it produced. The
# pairing value -> XML is then PowerPoint's own, not ours.
def cmd_enumdeck(out_pptx, lo, hi):
    """One blank slide per enum value in [lo, hi]; slide k carries value lo+k."""
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]
    for _ in range(lo, hi + 1):
        prs.slides.add_slide(blank)
    prs.save(out_pptx)
    json.dump({"lo": lo, "hi": hi, "count": hi - lo + 1},
              io.open(os.path.splitext(out_pptx)[0] + ".manifest.json", "w",
                      encoding="utf-8"), indent=1)
    print("enumdeck %s: %d slides, values %d..%d" % (
        out_pptx, hi - lo + 1, lo, hi))
    return 0


def cmd_enumread(deck, json_out, manifest=None):
    """Read each slide back: enum value -> the XML PowerPoint wrote for it."""
    # The scanned deck is saved under a different name than the deck that was
    # built (probe_enum_scan.ps1 needs a separate output), so the manifest
    # lives next to the INPUT deck, not next to this one.
    if manifest is None:
        manifest = os.path.join(os.path.dirname(deck), "enumdeck.manifest.json")
    man = _load_json(manifest)
    lo, hi = man["lo"], man["hi"]
    out = []
    with zipfile.ZipFile(deck) as z:
        for v in range(lo, hi + 1):
            name = "ppt/slides/slide%d.xml" % (v - lo + 1)
            if name not in z.namelist():
                continue
            xml = z.read(name).decode("utf-8")
            blk = _extract_transition(xml)
            out.append({"value": v, "xml": blk})
    known = [r for r in out if r["xml"]]
    if json_out:
        json.dump(out, io.open(json_out, "w", encoding="utf-8"),
                  ensure_ascii=False, indent=1)
    print("enumread: %d slides, %d produced a transition block" % (
        len(out), len(known)))
    for r in known:
        print("  %5d  %s" % (r["value"], r["xml"][:150]))
    if json_out:
        print("  -> %s" % json_out)
    return 0


def cmd_table(report, scans, out):
    rep = _load_json(report)
    enum_by_child = {}
    for s in scans:
        for r in _load_json(s):
            if not r.get("xml"):
                continue
            m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>",
                          r["xml"], re.S)
            if not m:
                continue
            child = re.sub(r"\s+", " ", m.group(1)).strip()
            enum_by_child.setdefault(child, r["value"])

    def base_id(child):
        tag = _child_tag(child)
        attrs = dict(_child_attrs(child))
        return "%s:%s" % (tag, attrs.get("prst", ""))

    by_spec, gallery, claimed = {}, [], set()
    for n, r in enumerate(rep["rows"]):
        if not r.get("readback"):
            continue
        m = re.search(r"<p:transition[^>]*>(.*?)</p:transition>",
                      r["readback"], re.S)
        child = re.sub(r"\s+", " ", m.group(1)).strip() if m else ""
        entry = {
            "spec": r["spec"], "ui": r["ui"], "en": r["en"],
            "group": r["group"], "family": r["family"], "child": child,
            "xml": _template(r["readback"]),
            "entryEffect": _lookup_enum(enum_by_child, child),
            "roundtrip": r["verdict"],
            "order": n,          # gallery order, so docs can render as the UI does
            "source": r.get("source"),
        }
        claimed.add(base_id(child))
        if r.get("source") == "motion":
            by_spec[r["spec"]] = entry
        else:
            gallery.append(entry)

    # Everything PowerPoint can write that no gallery item and no motion spec
    # claims: still valid in the file format, just not reachable from the UI.
    legacy = []
    for child, v in sorted(enum_by_child.items(), key=lambda kv: kv[1]):
        if base_id(child) in claimed:
            continue
        legacy.append({"child": child, "entryEffect": v})

    doc = {
        "meta": {
            "purpose": "UI gallery item -> the XML PowerPoint actually writes.",
            "measured_by": "scripts/build_transition_table.py (build/collect/"
                           "enumdeck/enumread/table) + scripts/probe_transitions.ps1 "
                           "+ scripts/probe_enum_scan.ps1",
            "method": [
                "enum scan: one slide per PpEntryEffect value, set over COM, "
                "saved, read back -- gives PowerPoint's own element per value",
                "roundtrip: one 2-slide deck per candidate, opened and saved "
                "-- gives 'accepted and kept' plus the canonical wrapper",
            ],
            "enum_range_scanned": "1..4200 (no transition exists above 3956)",
            "rules": RULES,
        },
        "namespaces": NS,
        "by_spec": by_spec,
        "gallery": gallery,
        "enum_by_child": enum_by_child,
        "not_in_gallery": legacy,
        "summary": {
            "gallery": len(gallery),
            "motion_specs": len(by_spec),
            "enum_values": len(enum_by_child),
            "not_in_gallery": len(legacy),
        },
    }
    json.dump(doc, io.open(out, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    s = doc["summary"]
    print("table: %d gallery + %d motion spec, %d enum values, "
          "%d reachable only in XML -> %s" % (
              s["gallery"], s["motion_specs"], s["enum_values"],
              s["not_in_gallery"], out))
    return 0


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


# --------------------------------------------------------------------------
# video: one combined deck, for the rendering pass
# --------------------------------------------------------------------------
def cmd_video(case_dir, out_dir, out_pptx):
    """Combine the candidates that SURVIVED the roundtrip into one deck:
    slide 1 = title, then per candidate a from/target pair with the transition
    on the target. In the exported video the tested transitions appear as the
    multi-frame change bursts, in slide order."""
    man = _load_json(os.path.join(case_dir, "manifest.json"))
    from pptx import Presentation
    from pptx.util import Pt, Emu
    from pptx.dml.color import RGBColor

    prs = Presentation()
    prs.slide_width = Emu(9144000)
    prs.slide_height = Emu(5143500)
    blank = prs.slide_layouts[6]

    def add(bg, big, small):
        s = prs.slides.add_slide(blank)
        r = s.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
        r.fill.solid()
        r.fill.fore_color.rgb = RGBColor(*bg)
        r.line.fill.background()
        r.shadow.inherit = False
        tb = s.shapes.add_textbox(Pt(60), Pt(60), prs.slide_width - Pt(120),
                                  prs.slide_height - Pt(120))
        tf = tb.text_frame
        tf.text = big
        tf.paragraphs[0].font.size = Pt(120)
        tf.paragraphs[0].font.bold = True
        p1 = tf.add_paragraph()
        p1.text = small
        p1.font.size = Pt(20)

    add((150, 150, 150), "0", "probe title / none")
    blocks = []
    kept = []
    for i, h in enumerate(man["hypotheses"]):
        done = os.path.join(out_dir, man["cases"][i])
        if not os.path.exists(done):
            continue                      # rejected upstream; nothing to show
        kept.append(h)
        add((235, 220, 200), str(len(kept)), "from")
        add((70, 110, 190), h["ui"], "%s / %s" % (h["en"], h["family"]))
        blocks.append(wrap_transition(h["family"], h["child"]))
    prs.save(out_pptx)

    buf = io.BytesIO(open(out_pptx, "rb").read())
    out = io.BytesIO()
    with zipfile.ZipFile(buf) as zin, zipfile.ZipFile(
            out, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            m = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", item.filename)
            if m and int(m.group(1)) % 2 == 0:
                data = _slide_xml(data, blocks[int(m.group(1)) // 2 - 1])
            zout.writestr(item, data)
    open(out_pptx, "wb").write(out.getvalue())

    json.dump({"order": [h["ui"] for h in kept],
               "slide_seconds": SLIDE_SECONDS},
              io.open(os.path.splitext(out_pptx)[0] + ".order.json", "w",
                      encoding="utf-8"), ensure_ascii=False, indent=1)
    print("video deck %s: %d slides, %d candidates" % (
        out_pptx, 1 + 2 * len(kept), len(kept)))
    return 0


# --------------------------------------------------------------------------
# sheets: contact sheets for the visual pass
# --------------------------------------------------------------------------
def cmd_sheets(video, out_dir, want=10):
    import cv2
    import numpy as np

    os.makedirs(out_dir, exist_ok=True)
    cap = cv2.VideoCapture(video)
    frames = []
    while True:
        ok, f = cap.read()
        if not ok:
            break
        frames.append(f)
    cap.release()
    gray = [cv2.cvtColor(f, cv2.COLOR_BGR2GRAY) for f in frames]

    # a burst = >=3 consecutive frames that differ from their predecessor;
    # single-frame jumps are the hard cuts between untested slide pairs.
    bursts, start = [], None
    for i in range(1, len(frames)):
        d = float(np.mean(cv2.absdiff(gray[i], gray[i - 1])))
        if d > 6.0:
            if start is None:
                start = i
        elif start is not None:
            if i - start >= 3:
                bursts.append((start, i))
            start = None
    print("video %d frames, %d multi-frame bursts" % (len(frames), len(bursts)))

    order = []
    opath = os.path.splitext(video)[0] + ".order.json"
    if os.path.exists(opath):
        order = json.load(io.open(opath, encoding="utf-8"))["order"]

    for k, (a, b) in enumerate(bursts):
        lo, hi = max(0, a - 2), min(len(frames), b + 1)
        idxs = [lo + int(j * (hi - 1 - lo) / max(1, want - 1)) for j in range(want)]
        tiles = [cv2.resize(frames[i], (569, 320)) for i in idxs]
        label = order[k] if k < len(order) else "?"
        sheet = np.hstack(tiles)
        cv2.putText(sheet, "#%02d %s  f%d-%d" % (k + 1, label, a, b),
                    (8, 26), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)
        cv2.imwrite(os.path.join(out_dir, "burst%02d.png" % (k + 1)), sheet)
    print("sheets -> %s (%d)" % (out_dir, len(bursts)))
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


def cmd_attrdeck(out_dir):
    """One deck per ATTR_PROBE row -- the "attribute value set" probe."""
    rows = [(s, lab, ov, why) for (s, lab, ov, why) in ATTR_PROBE]
    written = _probe_deck(out_dir, rows, "attr", lambda e: e["label"])
    print("attrdeck -> %s" % out_dir)
    print("  %d 份 deck（ATTR_PROBE：属性取值全集）" % len(written))
    for e in written:
        print("    %-22s %-46s %s" % (e["deck"], e["child"], e["why"]))
    return 0


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
