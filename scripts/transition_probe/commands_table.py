"""transition_probe.commands_table — 建表主链路

产出**发布的实测表**：造候选 deck → 枚举扫描 → 读回 → 合并成表。

模块内容表（7 个命令）
----------------------------------------------------------------
    build         
    collect       
    enumdeck      One blank slide per enum value in [lo, hi]; slide k carries 
    enumread      Read each slide back: enum value -> the XML PowerPoint wrote
    table         
    video         Combine the candidates that SURVIVED the roundtrip into one 
    sheets        

> 本模块由 `tools/split_commands.py` 从 `commands.py`（1367 行）按**探测维度**拆出
> （原文照搬，未改逻辑）。对外接口由 `../__init__.py` 重新导出 ——
> 这些 **CLI 子命令名写在 `facts/` 与 `reference/` 里，不能改名。**
"""

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
