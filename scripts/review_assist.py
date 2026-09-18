#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-ppt-office-motion :: review assistant.

Automates the parts of references/review-checklist.md that CAN be judged
mechanically, and -- just as important -- prints what it CANNOT judge.

Design rule learned the hard way (three real misjudgements this project):
a machine verdict is only ever a *signal*. Anything reported here as HEURISTIC
must be confirmed by eye before it is acted on, and any "nothing found" verdict
must not be trusted on its own.

Nothing here is destructive: it reads a .pptx and reports.

Usage:
    python review_assist.py --pptx deck.pptx [--source orig.pptx] [--json]
"""
import argparse
import json
import os
import re
import sys
import zipfile

P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
SLIDE_RE = re.compile(r'^ppt/slides/slide(\d+)\.xml$')
SHAPE_TAGS = ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp')
EMU_PT = 12700

# theme colours that make a page read as "rainbow"
RAINBOW_HUES = 4          # distinct strong hues on one slide
MANY_FONTSIZES = 7        # distinct font sizes on one slide


def load(pptx):
    z = zipfile.ZipFile(pptx)
    out = {'zip': z}
    out['slides'] = sorted((n for n in z.namelist() if SLIDE_RE.match(n)),
                           key=lambda s: int(SLIDE_RE.match(s).group(1)))
    out['xml'] = {n: z.read(n).decode('utf-8', 'replace') for n in out['slides']}
    out['media'] = {n: z.getinfo(n).file_size
                    for n in z.namelist() if n.startswith('ppt/media/')}
    return out


def rels_for(z, slide):
    p = 'ppt/slides/_rels/%s.rels' % os.path.basename(slide)
    try:
        return z.read(p).decode('utf-8', 'replace')
    except KeyError:
        return ''


# --------------------------------------------------------------------------
# material hygiene -- the checks that came out of a real accident
# --------------------------------------------------------------------------
def check_material_hygiene(d):
    """template1's rotating layer was a full-screen SCREENSHOT with 69% near-white
    pixels, i.e. it carried PowerPoint's own UI. Catch that shape of accident."""
    findings = []
    try:
        from PIL import Image
        import numpy as np
        import io
    except ImportError:
        return [('SKIP', 'Pillow/numpy absent, image hygiene not checked',
                 'install Pillow + numpy to enable')], []

    used = {}
    for s in d['slides']:
        for m in re.finditer(r'Target="\.\./media/([^"]+)"', rels_for(d['zip'], s)):
            used.setdefault(m.group(1), set()).add(os.path.basename(s))

    imgs = []
    for name in d['media']:
        base = name.split('/')[-1]
        if not re.search(r'\.(png|jpe?g|gif|bmp|tiff?)$', base, re.I):
            continue
        try:
            raw = d['zip'].read(name)
            im = Image.open(io.BytesIO(raw))
            a = np.asarray(im.convert('RGB')).astype(np.int16)
            h, w = a.shape[:2]
            near_white = float(((a > 235).all(axis=2)).mean())
            row_white = ((a > 235).all(axis=2).mean(axis=1))
            # a UI band = a long run of rows that are nearly all white
            band = int((row_white > 0.9).sum())
            imgs.append({'file': base, 'size': [w, h], 'bytes': len(raw),
                         'near_white': round(near_white, 3),
                         'white_rows': band, 'mode': im.mode})
            if near_white > 0.35 or band > h * 0.08:
                findings.append((
                    'HEURISTIC', 'possible screenshot / UI chrome: %s' % base,
                    '%.0f%% near-white pixels, %d near-white rows of %d '
                    '(this is how template1 shipped PowerPoint UI in a layer)'
                    % (near_white * 100, band, h)))
            # unusual aspect ratios are fine, but flag extreme ones
            if w and h and (w / h > 4 or h / w > 4):
                findings.append(('HEURISTIC', 'very extreme aspect: %s' % base,
                                 '%dx%d' % (w, h)))
        except Exception as e:
            findings.append(('SKIP', 'cannot open %s' % base, str(e)[:80]))
    return findings, imgs


def check_display_vs_source(d, imgs):
    """Is any picture displayed larger than its pixel size (i.e. will look soft)?"""
    out = []
    by_size = {i['file']: i['size'] for i in imgs}
    for s in d['slides']:
        rels = rels_for(d['zip'], s)
        rid2file = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="\.\./media/([^"]+)"', rels))
        for blk in re.split(r'(?=<p:pic[ >])', d['xml'][s]):
            emb = re.search(r'r:embed="(rId\d+)"', blk)
            if not emb:
                continue
            f = os.path.basename(rid2file.get(emb.group(1), ''))
            if f not in by_size:
                continue
            w_pt = h_pt = 0
            ext = re.search(r'<a:ext cx="(\d+)" cy="(\d+)"/>', blk)
            if ext:
                w_pt = int(ext.group(1)) / EMU_PT
                h_pt = int(ext.group(2)) / EMU_PT
            sw, sh = by_size[f]
            # 1pt on a 960x540 slide exported at 1280 wide ~= 1.33 px
            need_w = w_pt * (1280.0 / 960.0)
            upscale = need_w / sw if sw else 0
            src_ar = sw / sh if sh else 0
            dst_ar = w_pt / h_pt if h_pt else 0
            if upscale > 1.15:
                out.append(('HEURISTIC', '%s is upscaled %.2fx' % (f, upscale),
                            '%dx%d px shown at %.0f pt wide' % (sw, sh, w_pt)))
            if src_ar and dst_ar and (max(src_ar, dst_ar) / min(src_ar, dst_ar)) > 1.06:
                out.append(('HEURISTIC', '%s aspect distorted %.0f%%' % (
                    f, (max(src_ar, dst_ar) / min(src_ar, dst_ar) - 1) * 100),
                    'source %.2f vs displayed %.2f' % (src_ar, dst_ar)))
    return out


# --------------------------------------------------------------------------
# page aesthetics -- heuristics only, and labelled as such
# --------------------------------------------------------------------------
def check_page_aesthetics(d):
    findings = []
    for s in d['slides']:
        x = d['xml'][s]
        n = int(SLIDE_RE.match(s).group(1))
        sizes = sorted({int(v) / 100 for v in re.findall(r'sz="(\d+)"', x)})
        if len(sizes) > MANY_FONTSIZES:
            findings.append(('HEURISTIC', 'slide %d: %d distinct font sizes' % (n, len(sizes)),
                             ', '.join('%.0f' % v for v in sizes)))
        cols = re.findall(r'<a:srgbClr val="([0-9A-Fa-f]{6})"/>', x)
        strong = []
        for c in cols:
            r_, g_, b_ = (int(c[i:i + 2], 16) for i in (0, 2, 4))
            mx, mn = max(r_, g_, b_), min(r_, g_, b_)
            if mx > 90 and (mx - mn) > 60:          # saturated, not near-grey
                strong.append(c.upper())
        uniq = sorted(set(strong))
        if len(uniq) > RAINBOW_HUES:
            findings.append(('HEURISTIC', 'slide %d: %d saturated colours' % (n, len(uniq)),
                             ', '.join(uniq[:10]) + '  (rainbow risk)'))
        rounds = len(re.findall(r'<a:prstGeom prst="roundRect"', x))
        if rounds >= 3:
            findings.append(('HEURISTIC', 'slide %d: %d rounded rectangles' % (n, rounds),
                             'card-stack look -- the strongest "AI deck" signal'))
    return findings


# --------------------------------------------------------------------------
# structure -- these are the trustworthy, mechanical ones
# --------------------------------------------------------------------------
def check_structure(d, source=None):
    checks = []
    for s in d['slides']:
        x = d['xml'][s]
        n = int(SLIDE_RE.match(s).group(1))
        morph = 'p159:morph' in x
        i_t = x.find('<p:timing>')
        i_tr = x.find('<p:transition')
        i_ac = x.find('<mc:AlternateContent')
        eff = len(re.findall(r'presetID="\d+"', x))
        paths = len(re.findall(r'<p:animMotion', x))
        has_trans = i_tr != -1 or i_ac != -1
        # If there is NO transition, there is no ordering to check -- reporting a
        # failure then is a false positive. (This very bug fired on a deck that
        # had effects but no transitions at all.)
        order_ok = True
        if has_trans and i_t != -1:
            first_trans = min(p for p in (i_tr, i_ac) if p != -1)
            order_ok = first_trans < i_t
        checks.append({'slide': n, 'morph': morph, 'effects': eff,
                       'motion_paths': paths,
                       'has_timing': i_t != -1, 'has_transition': has_trans,
                       'transition_before_timing': order_ok})
    problems = []
    for c in checks:
        if not c['transition_before_timing']:
            problems.append(('FAIL', 'slide %d: transition is NOT before <p:timing>' % c['slide'],
                             'CT_Slide is a sequence; wrong order makes it vanish on save'))
    if source and os.path.exists(source):
        try:
            sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
            import motion as M
            a = M.geometry_fingerprint(source)
            b = M.geometry_fingerprint(os.path.abspath(d['path']))
            diff = [k for k in sorted(set(a) | set(b)) if a.get(k) != b.get(k)]
            if diff:
                problems.append(('FAIL', 'geometry differs from source', ', '.join(diff)))
            else:
                problems.append(('OK', 'geometry identical to source',
                                 'over %d slides' % len(a)))
        except Exception as e:
            problems.append(('SKIP', 'geometry compare failed', str(e)[:100]))
    return checks, problems


CANNOT_JUDGE = [
    '视觉主角是否明确、层级是否清晰',
    '留白分布是否舒服、是否"实心空白块"',
    '标题是否结论句、文案是否可读',
    '配色是否协调（只能查出"彩虹风险"，判不出好坏）',
    '动效节奏是否舒服、每处动画是否有意图',
    '方向性擦除的方向是否正确（结构闸查不出，必须人眼）',
    '删掉动画后静态版面是否依然成立',
    '抠图边缘是否干净、有无色边',
    '整体是否"好看"',
]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--pptx', required=True)
    ap.add_argument('--source', help='original deck, to compare geometry')
    ap.add_argument('--json', action='store_true')
    ns = ap.parse_args()

    d = load(ns.pptx)
    d['path'] = ns.pptx

    struct, problems = check_structure(d, ns.source)
    hygiene, imgs = check_material_hygiene(d)
    display = check_display_vs_source(d, imgs)
    pages = check_page_aesthetics(d)

    report = {
        'pptx': os.path.abspath(ns.pptx),
        'slides': len(d['slides']),
        'media': imgs,
        'structure': struct,
        'problems': problems,
        'material_hygiene': hygiene,
        'display_scale': display,
        'page_aesthetics': pages,
        'cannot_judge': CANNOT_JUDGE,
    }

    if ns.json:
        print(json.dumps(report, ensure_ascii=False, indent=1))
        return 0

    print('=' * 78)
    print('REVIEW ASSIST  %s' % os.path.basename(ns.pptx))
    print('=' * 78)
    print('%d slides, %d readable image(s) of %d media entries'
          % (len(d['slides']), len(imgs), len(d['media'])))
    print()

    print('-- STRUCTURE (mechanical, trustworthy) --')
    for tag, msg, extra in problems:
        print('   [%s] %s' % (tag, msg))
        if extra:
            print('        %s' % extra)
    for c in struct:
        print('   slide %-3d effects=%-3d paths=%-2d morph=%-5s transition=%-5s'
              % (c['slide'], c['effects'], c['motion_paths'], c['morph'], c['has_transition']))
    print()

    print('-- MATERIAL HYGIENE (heuristic -- confirm by eye) --')
    if not hygiene:
        print('   nothing flagged  (absence of a finding is NOT proof of cleanliness)')
    for tag, msg, extra in hygiene:
        print('   [%s] %s' % (tag, msg))
        if extra:
            print('        %s' % extra)
    print()

    print('-- DISPLAY SCALE / ASPECT (heuristic) --')
    if not display:
        print('   nothing flagged')
    for tag, msg, extra in display:
        print('   [%s] %s  -- %s' % (tag, msg, extra))
    print()

    print('-- PAGE AESTHETICS (heuristic ONLY; these are signals, not verdicts) --')
    if not pages:
        print('   nothing flagged')
    for tag, msg, extra in pages:
        print('   [%s] %s' % (tag, msg))
        print('        %s' % extra)
    print()

    print('-- THIS TOOL CANNOT JUDGE (must be reviewed by eye) --')
    for item in CANNOT_JUDGE:
        print('   - %s' % item)
    print()
    print('Reminder: a machine verdict is a signal. See references/review-checklist.md')
    print('section 5 for the cross-validation rules that came out of three real')
    print('misjudgements.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
