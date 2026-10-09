"""motion.spec — spec 处理与**结构校验**。

spec 处理与**结构校验**。

    校验三个函数**不被本文件任何函数调用**，只被外部用
    （`verify_motion.py` / `selftest.py`）—— 所以按用途归这里，不按位置。
    它们和 `geometry_fingerprint` 同类：**产出结论，不产出 XML**。

模块内容表（内容表；改这一层前先读这里）
----------------------------------------------------------------
    校验         geometry_fingerprint / transitions_are_paired
                 active_transition_blocks
    spec 处理    load_spec / normalize_spec / schedule_spec / slide_size
    catalog      catalog_path / normalize_effects

> 本模块由 `tools/split_motion.py` 从 `motion.py`（1914 行）按**分层**拆出
> （原文照搬，未改逻辑）。**外部接口由 `../motion.py` 重新导出** ——
> 10 个模块 `import motion` 用的名字一个都没变。
"""

from motion_xml import CATALOG_PATH, GEOM_ATTRS, MC_NS, P_NS, SHAPE_TAGS, SLIDE_RE, element_spans, sp_tree_block

import hashlib
import json
import re
import zipfile
try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None
try:
    from lxml import etree as _ET
except ImportError:  # pragma: no cover
    _ET = None


# ---------------------------------------------------------------------------
# geometry fingerprint: the layout-frame proof
# ---------------------------------------------------------------------------
def geometry_fingerprint(pptx_path):
    out = {}
    with zipfile.ZipFile(pptx_path) as z:
        names = sorted([n for n in z.namelist() if SLIDE_RE.match(n)],
                       key=lambda n: int(SLIDE_RE.match(n).group(1)))
        for name in names:
            xml = z.read(name).decode("utf-8", "replace")
            tree = sp_tree_block(xml)
            rows = []
            for tag in SHAPE_TAGS:
                for s, e in element_spans(tree, tag):
                    block = tree[s:e]
                    cm = re.search(r'<p:cNvPr\b[^>]*\bid="(\d+)"[^>]*\bname="([^"]*)"', block)
                    ident = list(cm.groups()) if cm else ["", ""]
                    xf = re.search(r"<a:xfrm[^>]*>.*?</a:xfrm>", block, re.S)
                    geom = []
                    if xf:
                        for attr in GEOM_ATTRS:
                            am = re.search(r'\b%s="(-?\d+)"' % attr, xf.group(0))
                            geom.append(am.group(1) if am else "-")
                    rows.append([tag] + ident + ["|".join(geom)])
            payload = json.dumps(rows, ensure_ascii=False)
            out[name] = {
                "sha256": hashlib.sha256(payload.encode("utf-8")).hexdigest(),
                "shapes": len(rows),
            }
    return out


# ---------------------------------------------------------------------------
# spec handling
# ---------------------------------------------------------------------------
def load_spec(path):
    with open(path, encoding="utf-8") as fh:
        text = fh.read()
    if path.lower().endswith((".yaml", ".yml")):
        if yaml is None:
            raise RuntimeError("PyYAML required for .yaml specs (pip install pyyaml)")
        return yaml.safe_load(text)
    return json.loads(text)


def normalize_effects(slide_entry):
    raw = slide_entry.get("effects")
    if raw is None:
        raw = slide_entry.get("animation") or []
    if isinstance(raw, dict):
        out = []
        for k, v in raw.items():
            item = dict(v) if isinstance(v, dict) else {"effect": v}
            item.setdefault("target", k)
            out.append(item)
        return out
    if not isinstance(raw, list):
        raise ValueError("effects must be a list or mapping")
    return [dict(x) for x in raw]


def catalog_path():
    return CATALOG_PATH


def normalize_spec(spec):
    """[{page, effects, transition}] with page resolved from page/index/slide.

    apply_motion() does the same resolution inline; player.py needs it too, so the
    rule lives here once instead of being reimplemented.
    """
    out = []
    for entry in (spec.get("slides") or []):
        if "page" in entry:
            page = int(entry["page"])
        elif "index" in entry:
            page = int(entry["index"])
        elif "slide" in entry:
            v = entry["slide"]
            if not isinstance(v, int):
                raise ValueError("slide: must be numeric to resolve a page number")
            page = v
        else:
            raise ValueError("slide entry needs page/index/slide: %r" % (entry,))
        out.append({"page": page,
                    "effects": normalize_effects(entry),
                    "transition": entry.get("transition")})
    return out


def schedule_spec(effects):
    """Absolute start time per effect, in seconds.

    Mirrors build_timing(): each effect becomes its own <p:par>, so `after` waits
    for its whole PRECEDING GROUP to finish, while `with` shares the current
    group's start (which is why a spec puts `with` entries directly after the
    effect they accompany).

    The group's end is the max end of its members, not the last member's end: a
    `with` effect can be longer than the one it accompanies, and the following
    `after` must wait for the slower of them.
    """
    out, group_start, group_end = [], 0.0, 0.0
    for eff in effects:
        dur = float(eff.get("duration") or 0.6)
        delay = float(eff.get("delay") or 0.0)
        if str(eff.get("trigger", "after")).lower() == "with" and out:
            start = group_start
        else:
            start = group_end + delay
            group_start = start
            group_end = start
        end = start + dur
        if end > group_end:
            group_end = end
        out.append({"target": eff.get("target"), "effect": eff.get("effect"),
                    "start": round(start, 3), "duration": round(dur, 3),
                    "dir": str(eff["dir"]).strip().lower() if eff.get("dir") else None})
    return out


def slide_size(pptx_path, default=(960, 540)):
    """(cx, cy) of the slide in points, from <p:sldSz> in presentation.xml.

    sldSz is in EMU (1 pt = 12700 EMU). Falls back to the default when the part or
    the attribute is missing rather than failing the whole run.
    """
    try:
        with zipfile.ZipFile(pptx_path) as z:
            xml = z.read("ppt/presentation.xml").decode("utf-8", "replace")
    except (KeyError, OSError, zipfile.BadZipFile):
        return default
    m = re.search(r'<p:sldSz\b[^>]*\bcx="(\d+)"[^>]*\bcy="(\d+)"', xml)
    if not m:
        m = re.search(r'<p:sldSz\b[^>]*\bcy="(\d+)"[^>]*\bcx="(\d+)"', xml)
        if m:
            return (int(m.group(2)) / 12700.0, int(m.group(1)) / 12700.0)
        return default
    return (int(m.group(1)) / 12700.0, int(m.group(2)) / 12700.0)


MC = MC_NS


def transitions_are_paired(xml):
    """True when every <p:transition> sits inside an mc:AlternateContent wrapper.

    A deck may hold more <p:transition> elements than effective transitions only
    in that one sanctioned way (Choice + Fallback). Anything else is a real
    duplicate that PowerPoint may resolve unpredictably.
    """
    if _ET is None:
        return False
    try:
        root = _ET.fromstring(xml.encode("utf-8"))
    except Exception:
        return False
    total = sum(1 for _ in root.iter(P_NS + "transition"))
    wrapped = 0
    for alt in root.iter(MC + "AlternateContent"):
        wrapped += sum(1 for _ in alt.iter(P_NS + "transition"))
    return total == wrapped and total > 0


def active_transition_blocks(xml):
    """Spans of the *effective* <p:transition> elements in a slide.

    PowerPoint serialises a transition as an mc:AlternateContent pair: the
    p14:dur flavour under <mc:Choice> and a pre-2010 fallback under
    <mc:Fallback>. Both must be written -- that is how PowerPoint itself does it
    -- but only the Choice is the transition in force, so counting raw
    occurrences over-reports by one on every PowerPoint-saved deck. Fall back to
    every occurrence when there is no AlternateContent wrapper (our own output,
    or a pre-2010 deck).
    """
    spans = element_spans(xml, "transition")
    if _ET is None:
        return spans
    try:
        root = _ET.fromstring(xml.encode("utf-8"))
    except Exception:
        return spans
    effective = []
    for alt in root.iter(MC + "AlternateContent"):
        choice = alt.find(MC + "Choice")
        if choice is None:
            continue
        for tr in choice.iter(P_NS + "transition"):
            effective.append(tr)
    if not effective:
        return spans
    return spans[:len(effective)]
