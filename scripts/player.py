#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: build a browser motion preview from a spec.

Answers "I injected animation -- now how do I actually look at it?".

Why this exists instead of an MP4: Presentation.CreateVideo is unusable on the
machine this was measured on. Quality must be 0 (1 and 2 raise E_INVALIDARG),
and even Quality 0 returns a success code without ever producing a file -- see
references/com-pitfalls.md 16. So the motion is rebuilt as layered PNGs.

How it works
------------
For every animated shape of every slide named in the spec we produce:

    player/stage/sN_base.png   the slide with all animated shapes DELETED
    player/stage/sN_<id>.png   one shape, transparent (RGBA), at its own size

The player places each layer at the shape's real (Left, Top) on a slide-sized
stage and reveals it on the spec's schedule. Revealing a layer therefore shows
exactly what PowerPoint shows at that moment.

Three COM behaviours drive the implementation, all measured (see com-pitfalls 16):
  * Slide.Export ignores Shape.Visible -- hiding a shape does not remove it from
    a static export, so it cannot isolate a shape.
  * Shape.Export(path, 2) does honour Visible and yields real alpha (RGBA).
  * The base must be exported AFTER deleting the animated shapes, or every
    revealed layer lands on a baked-in copy of itself and the text ghosts.
"""
import argparse
import json
import os
import re
import sys

try:
    import yaml
except ImportError:  # pragma: no cover
    yaml = None

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import motion  # noqa: E402
from motion import catalog_path, load_spec, schedule_spec  # noqa: E402

EMU = 12700.0

# Entrance feel per effect alias: (start transform, use clip-path wipe).
# Unknown aliases fall back to a plain fade.
FEEL = {
    "appear":    ("", False),
    "fade":      ("translateY(0) scale(1)", False),
    "fadedZoom": ("scale(.94)", False),
    "dissolve":  ("", False),
    "glide":     ("translateX(48px)", False),
    "wipe":      ("", True),
    "zoom":      ("scale(.62)", False),
    "riseUp":    ("translateY(48px)", False),
    "fly":       ("translateY(-48px)", False),
    "float":     ("translateY(40px)", False),
    "swivel":    ("scale(.7) rotate(-6deg)", False),
    "spinner":   ("scale(.5) rotate(-90deg)", False),
}

# Start-state clip for each wipe direction, as CSS inset(top right bottom left).
# "wipe(left)" means the reveal travels leftwards, so the shape is initially cut
# away on its left edge... except PowerPoint names the filter by the direction the
# EDGE MOVES, so wipe(left) uncovers from the right edge inward. These values were
# read off a rendered PowerPoint run, not derived: getting it backwards makes the
# preview disagree with the real show. Default is `left`, matching the catalogue's
# presetSubtype=4 wipe.
WIPE_CLIP = {
    "left":      "0 0 0 100%",
    "right":     "0 100% 0 0",
    "up":        "100% 0 0 0",
    "down":      "0 0 100% 0",
    "upleft":    "100% 0 0 100%",
    "upright":   "100% 100% 0 0",
    "downleft":  "0 0 100% 100%",
    "downright": "0 100% 100% 0",
}

HTML_TMPL = r"""<!DOCTYPE html>
<html lang="zh"><head><meta charset="utf-8">
<title>PPT 动效预览 — __TITLE__</title>
<style>
*{box-sizing:border-box}
body{margin:0;background:#14100e;color:#efe6dc;
     font:15px/1.5 "Segoe UI",system-ui,-apple-system,"Microsoft YaHei",sans-serif}
header{padding:22px 28px 6px}
h1{margin:0 0 4px;font-size:21px;font-weight:600}
.sub{color:#b9a795;font-size:13.5px;margin-bottom:14px}
.sub code{color:#e8c9a0}
.tabs{display:flex;gap:8px;padding:0 28px 14px;flex-wrap:wrap}
.tab{padding:8px 15px;border:1px solid #453a31;border-radius:999px;background:#1e1815;
     color:#cfbfae;cursor:pointer;font-size:13.5px}
.tab.on{background:#945028;border-color:#945028;color:#fff}
.ctrl{display:flex;gap:10px;align-items:center;padding:0 28px 16px;flex-wrap:wrap}
button{padding:9px 17px;border-radius:8px;border:1px solid #5b4a3c;background:#2a211b;
       color:#efe6dc;cursor:pointer;font-size:14px}
button.pri{background:#945028;border-color:#945028}
button:hover{filter:brightness(1.15)}
.wrap{padding:0 28px 30px;display:flex;gap:22px;align-items:flex-start;flex-wrap:wrap}
.stage{position:relative;border-radius:12px;overflow:hidden;background:#222;
       box-shadow:0 18px 50px rgba(0,0,0,.55);flex:0 0 auto}
.stage img.base{position:absolute;left:0;top:0;width:100%;height:100%}
.stage img.lyr{position:absolute;opacity:0;transition:opacity .001s linear}
.tl{min-width:300px;flex:1 1 320px}
.tl h2{font-size:14px;margin:2px 0 10px;color:#cbb9a6;font-weight:600;
       text-transform:uppercase;letter-spacing:.8px}
.row{padding:7px 10px;border-left:3px solid #3a2f27;margin-bottom:5px;border-radius:0 6px 6px 0;
     background:#1b1613;font-size:13px}
.row.hot{border-left-color:#e8a33d;background:#241b14}
.row .t{color:#e8a33d;font-variant-numeric:tabular-nums;margin-right:8px}
.row .n{color:#9fd3a0}
.row .m{color:#8d7c6c;font-size:12px}
#probe{color:#ff9b6a;font:12px/1.5 monospace;padding:0 28px 6px;word-break:break-all}
.note{padding:0 28px 34px;color:#8d7c6c;font-size:12.5px;max-width:920px}
</style></head>
<body>
<header>
  <h1>PPT 动效预览 · __TITLE__</h1>
  <div class="sub">由 dsh <code>ppt-office-motion player</code> 生成：把 spec 里的动效
    按真实时间轴在浏览器里重放。共 __NSLIDES__ 页 / __NEFFECTS__ 个动效。</div>
</header>
<div class="tabs" id="tabs"></div>
<div class="ctrl">
  <button class="pri" id="replay">▶ 重放本页</button>
  <button id="auto">⏭ 连播全部</button>
  <span id="status" style="color:#8d7c6c;font-size:13px"></span>
</div>
<div class="wrap">
  <div class="stage" id="stage"><img class="base" id="base" alt=""></div>
  <div class="tl"><h2>时间轴</h2><div id="tl"></div></div>
</div>
<div id="probe"></div>
<div class="note">
  图层是 PowerPoint 真实渲染的每个形状（PNG 带透明通道），按 spec 的时序重放；
  不是视频，也不改版面 —— 几何不变已由 <code>apply --assert-geometry</code> 断言。
  本机 <code>Presentation.CreateVideo</code> 对所有参数组合都返回 E_INVALIDARG，故不导 MP4。
</div>
<script>
var DATA = /*DATA*/;
var FEEL = /*FEEL*/;
var WIPE_CLIP = /*WIPE_CLIP*/;
var stage = document.getElementById('stage');
var baseEl = document.getElementById('base');
var tlEl = document.getElementById('tl');
var tabsEl = document.getElementById('tabs');
var statusEl = document.getElementById('status');
var probeEl = document.getElementById('probe');
var cur = 0, timers = [], shown = [];
var SW = __SW__, SH = __SH__;
stage.style.width = SW + 'px';
stage.style.height = SH + 'px';

function clearTimers(){ for (var i=0;i<timers.length;i++) clearTimeout(timers[i]); timers = []; }
function probe(msg){ probeEl.textContent = msg; }

function render(i){
  cur = i; clearTimers();
  var d = DATA[i];
  for (var k=0;k<tabsEl.children.length;k++) tabsEl.children[k].className = (k===i?'tab on':'tab');
  baseEl.onload  = function(){ probe('base '+d.slide+' loaded '+baseEl.naturalWidth+'x'+baseEl.naturalHeight); };
  baseEl.onerror = function(){ probe('BASE FAILED: '+d.base); };
  baseEl.src = d.base;

  var old = stage.querySelectorAll('img.lyr');
  for (var q=0;q<old.length;q++) old[q].parentNode.removeChild(old[q]);

  var lead = parseFloat((d.transition && d.transition.duration) || 0);
  var items = [];

  d.layers.forEach(function(L){
    var it = null;
    for (var z=0;z<d.items.length;z++) if (String(d.items[z].target)===String(L.id)) it = d.items[z];
    if (!it) it = {start:0,duration:.6,effect:'fade'};
    var img = document.createElement('img');
    img.className = 'lyr'; img.alt = L.name;
    img.style.left   = (L.x/SW*100)+'%';
    img.style.top    = (L.y/SH*100)+'%';
    img.style.width  = (L.w/SW*100)+'%';
    img.style.height = (L.h/SH*100)+'%';
    img.onerror = function(){ probe('LAYER FAILED: '+L.name+' ('+L.src+')'); };
    img.src = L.src;
    stage.appendChild(img);

    var feel = FEEL[it.effect] || FEEL.fade;
    var tr = feel[0], wipe = feel[1];
    img.style.transition = 'opacity '+it.duration+'s cubic-bezier(.22,.7,.3,1)'
      + (wipe ? '' : ', transform '+it.duration+'s cubic-bezier(.22,.7,.3,1)');
    // A wipe's start state is direction-dependent. Showing every wipe as
    // "grow from the left" made a top-down reveal and a left-to-right line draw
    // look identical in the preview, which is the one place motion is visible --
    // so the preview would have validated the wrong thing.
    var clip = WIPE_CLIP[it.dir] || WIPE_CLIP.left;
    if (wipe) img.style.clipPath = 'inset('+clip+')';
    else if (tr) img.style.transform = tr;
    items.push({el:img, it:it, tr:tr, wipe:wipe});
  });

  items.forEach(function(S){
    timers.push(setTimeout(function(){
      S.el.style.opacity = '1';
      if (S.wipe){ S.el.style.transition += ', clip-path '+S.it.duration+'s ease-out';
                   S.el.style.clipPath = 'inset(0 0 0 0)'; }
      else if (S.tr) S.el.style.transform = 'none';
    }, (S.it.start + lead) * 1000));
  });

  tlEl.innerHTML = d.items.map(function(x){
    var nm = '';
    for (var z=0;z<d.layers.length;z++) if (String(d.layers[z].id)===String(x.target)) nm = d.layers[z].name;
    return '<div class="row"><span class="t">'+x.start.toFixed(2)+'s</span>'
      + '<span class="n">'+(x.effect||'')+'</span> <span class="m">'+nm+' · id '+x.target+'</span></div>';
  }).join('');
  var total = 1;
  d.items.forEach(function(x){ total = Math.max(total, x.start + x.duration); });
  statusEl.textContent = '本页 '+d.items.length+' 个动效，约 '+total.toFixed(1)+' 秒';

  var rows = tlEl.querySelectorAll('.row');
  d.items.forEach(function(x,k){
    timers.push(setTimeout(function(){
      for (var r=0;r<rows.length;r++) rows[r].className = 'row';
      if (rows[k]) rows[k].className = 'row hot';
    }, (x.start + lead) * 1000));
  });
  timers.push(setTimeout(function(){
    for (var r=0;r<rows.length;r++) rows[r].className = 'row';
  }, (total + lead + 1.2) * 1000));
}

DATA.forEach(function(d,i){
  var b = document.createElement('button');
  b.className = 'tab'; b.textContent = '第 '+d.slide+' 页 · '+d.items.length+' 动效';
  b.onclick = function(){ render(i); };
  tabsEl.appendChild(b);
});
document.getElementById('replay').onclick = function(){ render(cur); };
document.getElementById('auto').onclick = function(){
  var k = 0;
  (function step(){
    if (k >= DATA.length) return;
    render(k);
    var d = DATA[k], total = 1;
    d.items.forEach(function(x){ total = Math.max(total, x.start + x.duration); });
    var lead = parseFloat((d.transition && d.transition.duration) || 0);
    k++;
    setTimeout(step, (total + lead + 1.6) * 1000);
  })();
};

// Decode every image before the first render. Without this the browser starts the
// reveal transition against an element it has not decoded yet, the transition
// completes on an undecoded image, and the layer sits semi-transparent forever --
// which looks exactly like a broken export.
var urls = [];
DATA.forEach(function(d){
  urls.push(d.base);
  d.layers.forEach(function(L){ urls.push(L.src); });
});
var pending = urls.length, failed = 0;
urls.forEach(function(u){
  var im = new Image();
  im.onload = function(){ if (--pending === 0) start(); };
  im.onerror = function(){ failed++; if (--pending === 0) start(); };
  im.src = u;
});
function start(){
  probe(failed ? (failed+' IMAGE(S) FAILED TO LOAD')
               : ('assets ready: '+urls.length+' images'));
  render(0);
}
</script>
</body></html>
"""


# --------------------------------------------------------------------- geometry
def top_level_shapes(xml):
    """The direct children of <p:spTree>, i.e. the real shape list.

    Depth tracking, not a substring scan: p:sp also appears as <a:sp> inside
    geometry definitions, and shapes nest inside groups.
    """
    s = xml.find("<p:spTree")
    if s == -1:
        return []
    s = xml.find(">", s) + 1
    tree = xml[s:xml.find("</p:spTree>")]
    out, depth = [], 0
    for m in re.finditer(r"<(/?)([a-zA-Z0-9]+):([a-zA-Z0-9]+)([^>]*?)(/?)>", tree):
        closing, _ns, local, attrs, selfclose = m.groups()
        if closing:
            depth -= 1
            continue
        if depth == 0 and local in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
            end = m.end() if selfclose else tree.find("</p:%s>" % local, m.end())
            out.append(tree[m.start():end] if end != -1 else tree[m.start():])
        if not selfclose:
            depth += 1
    return out


def parse_shape(block):
    nm = re.search(r"<p:cNvPr\b[^>]*?/?>", block)
    name = sid = ""
    if nm:
        n = re.search(r'\bname="([^"]*)"', nm.group(0))
        i = re.search(r'\bid="(\d+)"', nm.group(0))
        name = n.group(1) if n else ""
        sid = i.group(1) if i else ""
    off = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', block)
    ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', block)
    txt = "".join(re.findall(r"<a:t>([^<]*)</a:t>", block)).strip()
    box = None
    if off and ext:
        box = tuple(round(int(v) / EMU, 2) for v in
                    (off.group(1), off.group(2), ext.group(1), ext.group(2)))
    return {"id": sid, "name": name, "pt": box, "text": txt,
            "hasImage": bool(re.search(r'r:embed="', block)),
            "tag": "pic" if "<p:pic" in block[:80] else "sp"}


def shape_index(pptx, slide):
    """Ordered shape records for one slide, read straight from the package."""
    import zipfile
    with zipfile.ZipFile(pptx) as z:
        xml = z.read("ppt/slides/slide%d.xml" % slide).decode("utf-8", "replace")
    return [parse_shape(b) for b in top_level_shapes(xml)]


def shape_paragraphs(pptx, slide):
    """{shape_id: [paragraph text, ...]} for one slide, empty strings included.

    A textbox animated as a whole arrives as one block, which is fine for a label
    and flat for a lead. PowerPoint staggers a text block by PARAGRAPH, so a page
    that wants "line one, then line two" needs to know where the paragraph breaks
    are -- the shape record alone cannot say, because it reports the box's
    concatenated text with the newlines flattened away.
    """
    import zipfile
    from lxml import etree
    with zipfile.ZipFile(pptx) as z:
        xml = z.read("ppt/slides/slide%d.xml" % slide).decode("utf-8", "replace")
    out = {}
    root = etree.fromstring(xml.encode("utf-8"))
    ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main",
          "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
    for sp in root.iter("{%s}sp" % ns["p"]):
        cNvPr = sp.find(".//{%s}cNvPr" % ns["p"])
        txBody = sp.find(".//{%s}txBody" % ns["p"])
        if cNvPr is None or txBody is None:
            continue
        paras = []
        for p in txBody.findall("{%s}p" % ns["a"]):
            paras.append("".join(t.text or "" for t in p.iter("{%s}t" % ns["a"])))
        out[str(cNvPr.get("id"))] = paras
    return out


def classify(sh, slide_w, slide_h, z, total):
    """Guess a design role from geometry + text.

    Only needed when writing a spec from scratch: a deck authored in PowerPoint
    has no meaningful shape names to target, so effects get picked by role.
    """
    if not sh["pt"]:
        return "unknown"
    x, y, w, h = sh["pt"]
    if sh["text"]:
        if h >= 70 and y > slide_h * 0.45:
            return "title"
        if h >= 70:
            return "headline"
        if y < slide_h * 0.12:
            return "eyebrow"
        return "label"
    if w >= slide_w * 0.28 and h >= slide_h * 0.35:
        return "photo-card"
    if w <= slide_w * 0.22 and h <= slide_h * 0.12:
        return "chip"
    if z >= total - 2:
        return "backdrop"
    return "panel"


# -------------------------------------------------------------------- rendering
def render_layers(pptx, plan, outdir, width, height, quiet=False):
    """Write <outdir>/stage/*.png. Returns {slide: [layer, ...]}.

    Each layer's `src` is relative to the HTML's folder, which is the parent of
    `outdir` -- see build_player().
    """
    try:
        import win32com.client as w
    except ImportError:
        raise RuntimeError("pywin32 is required for `player` (pip install pywin32)")

    outdir = os.path.abspath(outdir)
    # srcs are relative to the HTML, which sits in outdir; the PNGs sit in
    # outdir/stage. So the prefix is just "stage/", NOT outdir's own name.
    rel = "stage/"
    stage = os.path.join(outdir, "stage")
    if os.path.isdir(stage):
        for f in os.listdir(stage):
            if f.endswith(".png"):
                os.remove(os.path.join(stage, f))
    else:
        os.makedirs(stage)

    app = w.Dispatch("PowerPoint.Application")
    pres = app.Presentations.Open(os.path.abspath(pptx), ReadOnly=True,
                                  Untitled=False, WithWindow=True)
    layers = {}
    try:
        for slide, targets in plan:
            s = pres.Slides(slide)
            by_id = {sh["id"]: sh for sh in shape_index(pptx, slide) if sh["pt"]}
            by_com = {sh.Id: sh for sh in s.Shapes}

            for tid in targets:
                sh = by_id.get(str(tid))
                target = by_com.get(int(tid))
                if sh is None or target is None:
                    if not quiet:
                        print("  ! slide %d: shape %s not found" % (slide, tid))
                    continue
                name = "s%d_%s.png" % (slide, tid)
                # ppShapeFormatPNG = 2. The filter is a numeric enum; passing the
                # string "PNG" raises "invalid literal for int()".
                target.Export(os.path.join(stage, name), 2)
                x, y, cw, ch = sh["pt"]
                layers.setdefault(slide, []).append({
                    "id": int(tid), "img": name, "name": sh["name"],
                    "src": rel + name,
                    "x": x, "y": y, "w": cw, "h": ch,
                })

            # Delete the animated shapes, THEN export what is left as the base.
            # Exporting the full slide as the base and layering on top
            # double-renders every animated shape (the revealed layer lands on its
            # own baked-in copy and the text ghosts).
            for tid in targets:
                target = by_com.get(int(tid))
                if target is not None:
                    try:
                        target.Delete()
                    except Exception as exc:
                        if not quiet:
                            print("  ! could not delete %s: %s" % (tid, exc))
            s.Export(os.path.join(stage, "s%d_base.png" % slide), "PNG", width, height)
            if not quiet:
                print("  slide %d: %d layer(s)" % (slide, len(targets)))
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            app.Quit()
        except Exception:
            pass

    for sl in layers:
        layers[sl][0]["base"] = "stage/s%d_base.png" % sl
    return layers


def build_player(pptx, spec_path, outdir, width=None, height=None,
                 title=None, quiet=False):
    spec = load_spec(spec_path)
    entries = motion.normalize_spec(spec)

    # Slide design size drives the stage; fall back to the spec's usual 960x540.
    if width is None or height is None:
        w, h = motion.slide_size(pptx)
        width = width or int(w)
        height = height or int(h)

    plan = []
    for e in entries:
        targets = []
        for eff in e.get("effects") or []:
            t = eff.get("target")
            if isinstance(t, (list, tuple)):
                targets.extend(str(v) for v in t if str(v).isdigit())
            elif str(t).isdigit():
                targets.append(str(t))
        if targets:
            plan.append((e["page"], targets))

    if not plan:
        raise RuntimeError(
            "spec has no numeric targets; `player` needs shape ids "
            "(run `motion.py inspect --json` to find them)")

    os.makedirs(outdir, exist_ok=True)
    if not quiet:
        print("== building motion player ==")
        print("deck: %s" % os.path.abspath(pptx))
        print("stage: %dx%d pt" % (width, height))
    layers = render_layers(pptx, plan, outdir, width, height, quiet=quiet)

    payload = []
    for e in entries:
        sl = e["page"]
        if sl not in layers:
            continue
        payload.append({
            "slide": sl,
            "transition": e.get("transition") or {},
            "items": schedule_spec(e.get("effects") or []),
            "base": layers[sl][0]["base"],
            "layers": layers[sl],          # spec order == stage z-order
        })

    # Layer srcs are relative to the HTML's folder, so the HTML lives INSIDE
    # --outdir (as index.html) with the layers in --outdir/stage. The whole folder
    # stays relocatable that way; a sibling layout broke when the outdir name and
    # the html name diverged.
    outdir = os.path.abspath(outdir)
    html = (HTML_TMPL
            .replace("/*DATA*/", json.dumps(payload, ensure_ascii=False, indent=1))
            .replace("/*FEEL*/", json.dumps(FEEL, ensure_ascii=False))
            .replace("/*WIPE_CLIP*/", json.dumps(WIPE_CLIP, ensure_ascii=False))
            .replace("__TITLE__", title or os.path.basename(pptx))
            .replace("__NSLIDES__", str(len(payload)))
            .replace("__NEFFECTS__", str(sum(len(p["items"]) for p in payload)))
            .replace("__SW__", str(width)).replace("__SH__", str(height)))

    path = os.path.join(outdir, "index.html")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)

    missing = []
    for p in payload:
        for one in [p["base"]] + [L["src"] for L in p["layers"]]:
            # srcs are relative to outdir, which is where index.html lives
            if not os.path.exists(os.path.join(outdir, one.replace("/", os.sep))):
                missing.append(one)
    if missing:
        raise RuntimeError("missing layer images: %s" % missing[:5])

    if not quiet:
        print("player: %s (%d slides, %d effects, %d layers)" % (
            path, len(payload), sum(len(p["items"]) for p in payload),
            sum(len(p["layers"]) for p in payload)))
    return {"html": path, "stage": os.path.join(outdir, "stage"),
            "slides": len(payload),
            "effects": sum(len(p["items"]) for p in payload)}


# ------------------------------------------------------------------------- CLI
def main(argv=None):
    ap = argparse.ArgumentParser(description="Build an HTML motion preview from a spec")
    ap.add_argument("--pptx", required=True)
    ap.add_argument("--spec", required=True)
    ap.add_argument("--outdir", required=True,
                    help="e.g. preview -> writes preview.html and preview/stage/*.png")
    ap.add_argument("--width", type=int)
    ap.add_argument("--height", type=int)
    ap.add_argument("--title")
    ap.add_argument("--json", action="store_true")
    ns = ap.parse_args(argv)
    try:
        res = build_player(ns.pptx, ns.spec, ns.outdir, ns.width, ns.height, ns.title)
    except Exception as exc:
        print("FAILED: %s" % exc)
        return 1
    if ns.json:
        print(json.dumps(res, ensure_ascii=False, indent=1))
    print("open: %s" % res["html"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
