#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: OOXML motion injection engine（**门面**）.

注入 / 预览 / 去动画 / 体检 —— CLI 与主流程在这里，**实现按层在下面四个模块**：

    motion_xml.py     常量 + zip + XML 字符串手术 + 单例 + 形状索引   ← 地基
    motion_timing.py  timing / transition 的 XML 生成
    motion_media.py   3D 相机 + 图片填充
    motion_spec.py    spec 处理 + 结构校验

**本文件重新导出上面四层的全部名字**，所以 `import motion` 的调用方
（10 个模块）一行都不用改。2026-10-08 曾判断「不拆」——理由是"要动 10 个
依赖者，不是单次改动能验证的"；2026-10-09 量清了内部调用图是**严格 DAG**
（apply 往下调，下层零回调），拆它变成一次可验证的搬运。

模块内容表
----------------------------------------------------------------
  1. apply_motion          主流程：读 → 改 → 写（--assert-geometry 证明没动版面）
  2. remove_empty_alternate_content / strip_animations   去动画（preview 用）
  3. inspect               看一个 pptx 到底有没有动效
  4. main                  CLI 分发
  5. re-export             下面四层的全部公开名

Usage:
    python motion.py apply   --pptx IN.pptx --spec motion.yaml --out OUT.pptx [--assert-geometry]
    python motion.py inspect --pptx IN.pptx [--json]
    python motion.py preview --pptx IN.pptx --out STATIC.pptx
    python motion.py catalog [--kind entrance|emphasis|path]
"""
import argparse
import json
import os
import re
import sys
import zipfile

try:
    from lxml import etree as _ET
except ImportError:  # pragma: no cover
    _ET = None

# 新模块与本文件同目录 —— 作为脚本直接运行时 sys.path[0] 就是它，
# 但被 `import motion` 加载时不一定，所以显式补一次（幂等，重复无害）。
# ⚠️ 用 `_HERE` 而不是 `HERE`：`HERE` 由下面从 motion_xml 导入 ——
# 两处都定义的话，后导入的会覆盖先定义的，读代码的人会以为有两个真相。
_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

# ---- 重新导出：把四个层拉回本命名空间（外部接口零变化的关键）-----------
#
# ⚠️ `main()` 里的 `import player` / `import check_coverage` 是**函数内延迟导入**，
# 它刻意打断 `motion ⇄ check_coverage` 那个环（见 CONTRIBUTING §六.4）。
# **它跟着 main() 一起留在这里 —— 别把那个 import 提到模块级。**

from motion_xml import (  # noqa: F401
    HERE, CATALOG_PATH, SLIDE_RE, P_NS,
    MC_NS, SHAPE_TAGS, TRIGGERS, GEOM_ATTRS,
    WIPE_DIRS, DIRECTIONAL_FILTERS, TRANSITION_REF_PATH, _load_transition_reference,
    _TRANSITION_REF, TRANSITIONS, TRANSITION_ENUM, EXT_CONTENT_TYPES,
    read_parts, write_parts, xml_unescape, xml_escape,
    element_spans, SLIDE_CHILD_ORDER, SLIDE_ROOT_CHILDREN, root_child_spans,
    insert_in_slide_order, replace_or_insert, remove_elements, set_singleton,
    singleton_spans, direct_children, FILL_GROUP, find_duplicate_singletons,
    sp_tree_block, index_shapes, resolve_targets,
)
from motion_timing import (  # noqa: F401
    NodeIds, split_effect_template, build_effect_node, set_wipe_direction,
    set_leaf_durations, set_anchor_durations, apply_overrides, build_timing,
    root_declares, add_root_namespace, MORPH_TEMPLATE, MORPH_OPTIONS,
    MORPH_SPEEDS, SPD_NOTCH_MS, _spd_for_ms, _spd_for_label,
    build_transition, drop_alternate_content, drop_transition_alternate_content,
)
from motion_media import (  # noqa: F401
    CAMERA_ANGLE_UNIT, CAMERA_MAX_ANGLE, CAMERA_DEFAULT_PRESET, CAMERA_PERSPECTIVE_PREFIXES,
    deg_to_angle, build_camera, camera_is_perspective, insert_scene3d,
    window_insets, build_fill_window, insert_blip_fill, ensure_content_type,
)
from motion_spec import (  # noqa: F401
    geometry_fingerprint, load_spec, normalize_effects, catalog_path,
    normalize_spec, schedule_spec, slide_size, MC,
    transitions_are_paired, active_transition_blocks,
)


# ---------------------------------------------------------------------------
# apply
# ---------------------------------------------------------------------------
def apply_motion(pptx_in, spec, out_path, assert_geometry=False):
    with open(CATALOG_PATH, encoding="utf-8") as fh:
        catalog = json.load(fh)

    infos = read_parts(pptx_in)
    by_name = {i.filename: d for i, d in infos}
    slide_names = sorted([n for n in by_name if SLIDE_RE.match(n)],
                         key=lambda n: int(SLIDE_RE.match(n).group(1)))
    # Slide size in points, read once. Needed by windowed fills to default the picture
    # rectangle to the full bleed; taking it from the package rather than assuming
    # 960x540 keeps the inset arithmetic correct on other slide sizes.
    sw, sh = slide_size(pptx_in)

    report = {
        "source": os.path.abspath(pptx_in),
        "out": os.path.abspath(out_path),
        "slides": [], "warnings": [], "errors": [],
        "media_plan": [], "effects_total": 0, "transitioned": 0,
    }

    spec_by_index, spec_by_name = {}, {}
    for entry in (spec.get("slides") or []):
        if "page" in entry:
            spec_by_index[int(entry["page"])] = entry
        elif "index" in entry:
            spec_by_index[int(entry["index"])] = entry
        elif "slide" in entry:
            v = entry["slide"]
            if isinstance(v, int):
                spec_by_index[v] = entry
            else:
                spec_by_name[str(v)] = entry
        else:
            report["errors"].append("slide entry needs page/index/slide: %r" % (entry,))

    replacements = {}
    for pos, sname in enumerate(slide_names, start=1):
        idx = int(SLIDE_RE.match(sname).group(1))
        xml = by_name[sname].decode("utf-8")
        entry = spec_by_index.get(idx) or spec_by_index.get(pos) or spec_by_name.get(sname)
        shapes = index_shapes(xml)
        sr = {"slide": idx, "part": sname, "effects": 0, "unknownTargets": []}

        if entry is not None:
            effects = normalize_effects(entry)
            entries = []
            for eff in effects:
                alias = eff.get("effect") or eff.get("name")
                if not alias:
                    report["errors"].append("slide %d: effect entry without 'effect'" % idx)
                    continue
                meta = catalog.get(str(alias))
                if meta is None:
                    report["errors"].append("slide %d: unknown effect alias %r" % (idx, alias))
                    continue
                ids = resolve_targets(eff.get("target"), shapes)
                if not ids:
                    sr["unknownTargets"].append(str(eff.get("target")))
                    report["errors"].append(
                        "slide %d: target %r not found in spTree" % (idx, eff.get("target")))
                    continue
                for sid in ids:
                    entries.append({
                        "shape_id": sid, "tpl": meta["template"],
                        "effect": eff, "delay": eff.get("delay") or 0.0,
                    })
                sr.setdefault("applied", []).append({
                    "effect": alias, "presetID": meta["presetID"],
                    "presetClass": meta["presetClass"], "target": eff.get("target"), "ids": ids,
                })
                sr["effects"] += len(ids)

            if entries:
                new_xml = replace_or_insert(xml, "timing", build_timing(entries, NodeIds(4)))
            else:
                new_xml = remove_elements(xml, "timing")

            trans = entry.get("transition")
            if trans is not None:
                if isinstance(trans, str):
                    trans = {"type": trans}
                is_morph = str(trans.get("type", "")).strip().lower() == "morph"
                block = build_transition(trans.get("type", "fade"), trans.get("duration"),
                                         trans.get("advanceAfter"), trans.get("onClick"),
                                         trans.get("option"), trans.get("speed"))
                # No xmlns:p14 root fixup any more: the measured block declares
                # every prefix it uses inside itself. The old code declared p14
                # on <p:sld> and, failing that, STRIPPED p14:dur and warned --
                # a silent loss of the transition duration that can no longer
                # happen, because the block carries its own declaration.
                #
                # Every measured transition is wrapped in mc:AlternateContent
                # (morph always was; the core ones are too -- PowerPoint wraps
                # them to host p14:dur). So the previous wrapper has to go for
                # ALL of them, not just morph, or re-running this tool on an
                # already-processed deck leaves two transitions behind.
                new_xml = drop_transition_alternate_content(new_xml)
                new_xml = remove_elements(new_xml, "transition")
                if block:
                    new_xml = replace_or_insert(new_xml, "transition", block)
                sr["transition"] = trans.get("type", "fade")
                report["transitioned"] += 1

            # ---- 3D cameras ------------------------------------------------
            # Not an animation: <a:scene3d> is a shape property, so it never
            # enters the timing tree. Applied separately from `effects`.
            for cam in (entry.get("cameras") or []):
                if not isinstance(cam, dict):
                    cam = {"target": cam}
                ids = resolve_targets(cam.get("target"), shapes)
                if not ids:
                    sr["unknownTargets"].append(str(cam.get("target")))
                    report["errors"].append(
                        "slide %d: camera target %r not found in spTree"
                        % (idx, cam.get("target")))
                    continue
                try:
                    frag = build_camera(cam)
                except ValueError as exc:
                    report["errors"].append("slide %d: %s" % (idx, exc))
                    continue
                prst = cam.get("prst") or CAMERA_DEFAULT_PRESET
                if not camera_is_perspective(prst):
                    report["warnings"].append(
                        "slide %d: camera prst=%s is a PARALLEL projection -- it "
                        "will never show a vanishing point" % (idx, prst))
                for sid in ids:
                    new_xml, status = insert_scene3d(new_xml, sid, frag)
                    sr.setdefault("cameras", []).append({
                        "target": cam.get("target"), "id": sid,
                        "prst": prst, "status": status,
                        "tilt_deg": cam.get("lat", cam.get("tilt", 0)),
                        "lon_deg": cam.get("lon", 0),
                    })
                    if status != "ok":
                        report["warnings"].append(
                            "slide %d: camera on id=%s -> %s" % (idx, sid, status))
                sr["cameras_applied"] = sr.get("cameras_applied", 0) + len(ids)

            # ---- windowed picture fills -------------------------------------
            # Also a shape property rather than an animation, so it lives beside
            # `cameras` and never enters the timing tree. This is what turns a shape
            # into a WINDOW onto a picture instead of a stretched thumbnail; the
            # animation that then scans it is a plain `wipe` / `pathRight` in
            # `effects`, which the spec already supported.
            for fw in (entry.get("fills") or []):
                if not isinstance(fw, dict):
                    report["errors"].append(
                        "slide %d: each fills entry must be a mapping with "
                        "target + window" % idx)
                    continue
                ids = resolve_targets(fw.get("target"), shapes)
                if not ids:
                    sr["unknownTargets"].append(str(fw.get("target")))
                    report["errors"].append(
                        "slide %d: fill target %r not found in spTree"
                        % (idx, fw.get("target")))
                    continue
                try:
                    frag = build_fill_window(fw, sw, sh)
                except ValueError as exc:
                    report["errors"].append("slide %d: %s" % (idx, exc))
                    continue
                for sid in ids:
                    new_xml, status = insert_blip_fill(new_xml, sid, frag)
                    sr.setdefault("fills", []).append({
                        "target": fw.get("target"), "id": sid,
                        "window": fw.get("window") or fw.get("bounds"),
                        "insets": re.search(r'<a:fillRect\b([^/]*)/>', frag).group(1).strip()
                                  if "<a:fillRect" in frag else "",
                        "status": status,
                    })
                    if status != "ok":
                        report["warnings"].append(
                            "slide %d: fill on id=%s -> %s" % (idx, sid, status))
                sr["fills_applied"] = sr.get("fills_applied", 0) + len(ids)

            for media in (entry.get("media") or []):
                report["media_plan"].append({
                    "slide": idx,
                    "src": media.get("src"),
                    "kind": media.get("kind", "auto"),
                    "bounds": media.get("bounds"),
                    "link": bool(media.get("link", False)),
                    "loop": bool(media.get("loop", False)),
                    "rewind": bool(media.get("rewind", True)),
                    "mute": bool(media.get("mute", False)),
                    "volume": media.get("volume"),
                    "autoplay": bool(media.get("autoplay", True)),
                    "elementId": media.get("elementId"),
                })

            if new_xml != xml:
                replacements[sname] = new_xml.encode("utf-8")

        report["effects_total"] += sr["effects"]
        report["slides"].append(sr)

    ct_name = "[Content_Types].xml"
    if ct_name in by_name:
        ct_xml = by_name[ct_name].decode("utf-8")
        original = ct_xml
        for m in report["media_plan"]:
            if m["src"]:
                ct_xml = ensure_content_type(ct_xml, os.path.splitext(m["src"])[1].lstrip("."))
        if ct_xml != original:
            replacements[ct_name] = ct_xml.encode("utf-8")

    # Final gate: every part we rewrote must still be well-formed XML. A malformed
    # part opens fine in no viewer but produces a baffling PowerPoint E_FAIL, so
    # never write one.
    if _ET is not None:
        for name in sorted(replacements):
            payload = replacements[name]
            if not name.lower().endswith(".xml"):
                continue
            try:
                _ET.fromstring(payload)
            except Exception as exc:
                report["errors"].append("refusing to write malformed XML in %s: %s" % (name, exc))
        if report["errors"]:
            return report

    write_parts(infos, replacements, out_path)

    if assert_geometry:
        before = geometry_fingerprint(pptx_in)
        after = geometry_fingerprint(out_path)
        diffs = [k for k in sorted(set(before) | set(after)) if before.get(k) != after.get(k)]
        report["geometry"] = {"ok": not diffs, "diff": diffs, "slides": len(before)}
        if diffs:
            report["errors"].append("geometry changed by motion injection: %s" % diffs)
    return report


def strip_animations(pptx_in, out_path):
    """Remove both <p:timing> and <p:transition> so the deck renders statically.

    Entrance animations keep their shapes hidden until they start, so a PNG/PDF
    grabbed from an animated deck is missing every "fly in" shape. Transitions
    have to go too: leaving them means the "static" copy still has slide
    transitions, which is not what the preview is for. Emptying an
    mc:AlternateContent wrapper is cleaned up as well -- an AlternateContent
    with no Choice and no Fallback is schema-invalid, and PowerPoint is entitled
    to reject it.
    """
    infos = read_parts(pptx_in)
    replacements = {}
    for info, data in infos:
        if not SLIDE_RE.match(info.filename):
            continue
        xml = data.decode("utf-8", "replace")
        new = remove_elements(xml, "timing")
        # Both forms have to go: the measured blocks wrap the transition in
        # mc:AlternateContent (that is how PowerPoint writes every one of them),
        # and a deck written by an older build of this tool carries a bare
        # <p:transition>. remove_elements() alone leaves an EMPTY mc:Choice
        # behind, and an AlternateContent whose Choice survives is not "empty",
        # so remove_empty_alternate_content would not clean it up either --
        # the wrapper has to be taken together with its transition.
        new = drop_transition_alternate_content(new)
        new = remove_elements(new, "transition")
        new = remove_empty_alternate_content(new)
        if new != xml:
            replacements[info.filename] = new.encode("utf-8")
    write_parts(infos, replacements, out_path)
    return {"out": os.path.abspath(out_path), "slides_stripped": len(replacements)}


def remove_empty_alternate_content(xml):
    """Drop mc:AlternateContent blocks that no longer hold a Choice or Fallback."""
    out = xml
    while True:
        nxt = None
        for m in re.finditer(r"<mc:AlternateContent(?=[\s/>])", out):
            s = m.start()
            gt = out.find(">", m.end())
            if gt == -1:
                break
            if out[gt - 1] == "/":
                nxt = (s, gt + 1)
                break
            close = out.find("</mc:AlternateContent>", gt)
            if close == -1:
                break
            e = close + len("</mc:AlternateContent>")
            body = out[gt + 1:close]
            if "<mc:Choice" not in body and "<mc:Fallback" not in body:
                nxt = (s, e)
                break
        if nxt is None:
            return out
        out = out[:nxt[0]] + out[nxt[1]:]


def inspect(pptx_path):
    out = {"file": os.path.abspath(pptx_path), "slides": []}
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted([n for n in z.namelist() if SLIDE_RE.match(n)],
                       key=lambda n: int(SLIDE_RE.match(n).group(1)))
        for name in names:
            xml = z.read(name).decode("utf-8", "replace")
            rels_name = "ppt/slides/_rels/%s.rels" % os.path.basename(name)
            try:
                rels = z.read(rels_name).decode("utf-8", "replace")
            except KeyError:
                rels = ""
            shapes = index_shapes(xml)
            presets = re.findall(r'presetID="(\d+)"[^>]*presetClass="(\w+)"', xml)
            out["slides"].append({
                "index": int(SLIDE_RE.match(name).group(1)),
                "shapes": [{"name": k, "id": v["id"], "tag": v["tag"]}
                           for k, v in shapes.items() if not v.get("pieces")],
                "timing": bool(re.search(r"<p:timing>", xml)),
                "effects": [{"presetID": int(a), "presetClass": b} for a, b in presets],
                "motionPaths": len(re.findall(r"<p:animMotion\b", xml)),
                "transition": bool(re.search(r"<p:transition\b", xml)),
                "media": sorted(set(re.findall(r'relationships/(video|audio|media)"', rels))),
            })
    return out


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser(description="OOXML motion injection for .pptx")
    sub = ap.add_subparsers(dest="cmd", required=True)

    a = sub.add_parser("apply")
    a.add_argument("--pptx", required=True)
    a.add_argument("--spec", required=True)
    a.add_argument("--out", required=True)
    a.add_argument("--assert-geometry", action="store_true")
    a.add_argument("--report")
    a.add_argument("--json", action="store_true")

    i = sub.add_parser("inspect")
    i.add_argument("--pptx", required=True)
    i.add_argument("--json", action="store_true")

    p = sub.add_parser("preview")
    p.add_argument("--pptx", required=True)
    p.add_argument("--out", required=True)

    c = sub.add_parser("catalog")
    c.add_argument("--kind", choices=["entrance", "emphasis", "path"])
    c.add_argument("--json", action="store_true")

    pl = sub.add_parser("player", help="build an HTML motion preview from a spec")
    pl.add_argument("--pptx", required=True)
    pl.add_argument("--spec", required=True)
    pl.add_argument("--outdir", required=True,
                    help="writes <outdir>.html next to <outdir>/stage/*.png")
    pl.add_argument("--title")
    pl.add_argument("--json", action="store_true")

    ck = sub.add_parser("check", help="audit spec coverage (shapes with no effect)")
    ck.add_argument("--pptx", required=True)
    ck.add_argument("--spec", required=True)
    ck.add_argument("--footer", action="append",
                    help="extra static-footer text marker (repeatable)")
    ck.add_argument("--json", action="store_true")

    ns = ap.parse_args(argv)

    if ns.cmd == "apply":
        report = apply_motion(ns.pptx, load_spec(ns.spec), ns.out,
                             assert_geometry=ns.assert_geometry)
        if ns.report:
            with open(ns.report, "w", encoding="utf-8") as fh:
                json.dump(report, fh, ensure_ascii=False, indent=1)
        if ns.json:
            print(json.dumps(report, ensure_ascii=False, indent=1))
        else:
            print("slides=%d effects=%d transitions=%d media=%d" % (
                len(report["slides"]), report["effects_total"],
                report["transitioned"], len(report["media_plan"])))
            if report.get("geometry"):
                g = report["geometry"]
                print("geometry: %s over %d slides" % (
                    "UNCHANGED" if g["ok"] else "CHANGED", g["slides"]))
            for w in report["warnings"]:
                print("WARN:", w)
            for e in report["errors"]:
                print("ERROR:", e)
        return 1 if report["errors"] else 0

    if ns.cmd == "inspect":
        data = inspect(ns.pptx)
        if ns.json:
            print(json.dumps(data, ensure_ascii=False, indent=1))
        else:
            for sl in data["slides"]:
                print("slide %d: shapes=%d effects=%d paths=%d transition=%s media=%s" % (
                    sl["index"], len(sl["shapes"]), len(sl["effects"]),
                    sl["motionPaths"], sl["transition"], sl["media"] or "-"))
                for sh in sl["shapes"]:
                    print("    id=%-6s %-14s %s" % (sh["id"], sh["tag"], sh["name"]))
        return 0

    if ns.cmd == "preview":
        print(json.dumps(strip_animations(ns.pptx, ns.out), ensure_ascii=False))
        return 0

    if ns.cmd == "catalog":
        with open(CATALOG_PATH, encoding="utf-8") as fh:
            catalog = json.load(fh)
        rows = [(k, v) for k, v in catalog.items() if not ns.kind or v["kind"] == ns.kind]
        if ns.json:
            print(json.dumps({k: {kk: vv for kk, vv in v.items() if kk != "template"}
                              for k, v in rows}, ensure_ascii=False, indent=1))
        else:
            for k, v in sorted(rows, key=lambda x: (x[1]["kind"], x[1]["presetID"])):
                flags = "".join(["P" if v["hasMotionPath"] else "-",
                                 "F" if v["hasFilter"] else "-",
                                 "V" if v["hasVisibilitySet"] else "-"])
                # The MsoAnimEffect names are printed because enum != presetID:
                # `spin` is presetID 8 but enum 61, so anyone reading a Microsoft
                # reference needs the name to find the effect they mean.
                print("%-24s %-9s presetID=%-4d sub=%-3d enum=%-4s %-26s %s" % (
                    k, v["kind"], v["presetID"], v["presetSubtype"],
                    v.get("enum", ""), v.get("enumName", ""), flags))
            print("\ntotal %d aliases (P=motion path, F=filter, V=visibility set)"
                  % len(rows))
            print("enum = MsoAnimEffect value; enumName = its constant. "
                  "enum != presetID by design.")
        return 0

    if ns.cmd == "player":
        # Imported lazily: player.py imports this module, and pywin32 is only
        # needed for this one command.
        import player
        try:
            res = player.build_player(ns.pptx, ns.spec, ns.outdir, title=ns.title)
        except Exception as exc:
            print("FAILED: %s" % exc)
            return 1
        if ns.json:
            print(json.dumps(res, ensure_ascii=False, indent=1))
        print("open: %s" % res["html"])
        return 0

    if ns.cmd == "check":
        import check_coverage
        argv = ["--pptx", ns.pptx, "--spec", ns.spec]
        for f in (ns.footer or []):
            argv += ["--footer", f]
        if ns.json:
            argv.append("--json")
        return check_coverage.main(argv)
    return 2


if __name__ == "__main__":
    sys.exit(main())
