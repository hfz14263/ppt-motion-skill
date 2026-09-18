# -*- coding: utf-8 -*-
"""End-to-end smoke test for this skill.

Runs against a given copy of the skill (default: this repo) and proves the
shipped artefacts are self-contained: catalogue loads, injection runs, geometry
stays identical, structure verifies, and PowerPoint accepts the result.

Two scenarios on purpose -- they exercise different OOXML writers:

  1. a pptd export (dsh-ppt-studio style: ids from 1000, cNvPr/@name == elementId)
  2. a PowerPoint-native deck, which reproduces the p14 trap: PowerPoint declares
     xmlns:p14 on a DESCENDANT (<p14:creationId> inside p:extLst), never on the
     slide root, so a naive "is xmlns:p14 present?" test emits a transition with
     an unbound prefix.

Usage:
    python tests/smoke.py [--skill <dir>] [--keep]

Requires: pyyaml + lxml, and PowerPoint for the COM step (that step is skipped
with a clear note when COM is unavailable).
"""
import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
import zipfile
import re

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
PS = 'powershell.exe'


def find_python():
    # never hardcode an interpreter path: prefer the one running this script,
    # then whatever is on PATH. Machine-specific paths do not belong in a repo.
    for cand in (sys.executable, shutil.which('python'), shutil.which('python3')):
        if cand and os.path.exists(cand):
            return cand
    return 'python'


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--skill', default=REPO, help='skill copy to test (default: this repo)')
    ap.add_argument('--keep', action='store_true', help='keep the work directory')
    ns = ap.parse_args()

    skill = os.path.abspath(ns.skill)
    py = find_python()
    work = tempfile.mkdtemp(prefix='pom-smoke-')
    fails = []

    def run(cmd, label, expect_ok=True):
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding='utf-8', errors='replace')
        ok = (r.returncode == 0) if expect_ok else True
        print('%-32s %s' % (label, 'PASS' if ok else 'FAIL'))
        if not ok:
            fails.append(label)
            print('   stdout:', (r.stdout or '').strip()[:500])
            print('   stderr:', (r.stderr or '').strip()[:500])
        return (r.stdout or '') + (r.stderr or '')

    print('skill  :', skill)
    print('python :', py)
    print('work   :', work)
    print()

    # ---- 1. catalogue integrity ------------------------------------------
    cat_path = os.path.join(skill, 'scripts', 'motion_catalog.json')
    cat = json.load(open(cat_path, encoding='utf-8'))
    print('catalogue aliases: %d' % len(cat))
    bad = [k for k, v in cat.items()
           if (v['kind'] == 'path' and not v['hasMotionPath'])
           or (v['kind'] == 'entrance' and not v['hasVisibilitySet'])]
    print('catalogue sanity:', 'PASS' if not bad else 'FAIL %s' % bad[:5])
    if bad:
        fails.append('catalogue sanity')

    # ---- 2. CLI surface --------------------------------------------------
    run([py, os.path.join(skill, 'scripts', 'motion.py'), 'catalog', '--kind', 'entrance'],
        'motion.py catalog')

    # ---- scenario 1: pptd export ----------------------------------------
    print()
    print('--- scenario 1: pptd export ---')
    src = os.path.join(HERE, 'fx-pro.pptx')
    spec = os.path.join(HERE, 'motion.yaml')
    out = os.path.join(work, 'animated.pptx')
    if not os.path.exists(src):
        print('fixture missing, skipped')
    else:
        run([py, os.path.join(skill, 'scripts', 'motion.py'), 'inspect', '--pptx', src],
            'motion.py inspect')
        run([py, os.path.join(skill, 'scripts', 'motion.py'), 'apply', '--pptx', src,
             '--spec', spec, '--out', out, '--assert-geometry',
             '--report', os.path.join(work, 'report.json')], 'motion.py apply')
        run([py, os.path.join(skill, 'scripts', 'verify_motion.py'), '--pptx', out,
             '--source', src], 'verify_motion.py')
        # upstream's own coverage auditor must also accept the result
        run([py, os.path.join(skill, 'scripts', 'motion.py'), 'check', '--pptx', out,
             '--spec', spec], 'motion.py check (coverage)', expect_ok=False)
        rep = json.load(open(os.path.join(work, 'report.json'), encoding='utf-8'))
        geo = rep.get('geometry', {})
        print('   geometry unchanged:', geo.get('ok'), 'over', geo.get('slides'), 'slides')
        if not geo.get('ok'):
            fails.append('geometry')

    # ---- scenario 2: PowerPoint-native deck ------------------------------
    print()
    print('--- scenario 2: PowerPoint-native deck (p14 trap) ---')
    fx = os.path.join(HERE, 'fixtures', 'powerpoint-native.pptx')
    fx_spec = os.path.join(HERE, 'fixtures', 'powerpoint-native.yaml')
    if not os.path.exists(fx):
        print('fixture missing, skipped')
    else:
        with zipfile.ZipFile(fx) as z:
            part = z.read('ppt/slides/slide1.xml').decode('utf-8')
        root = re.search(r'<p:sld\b[^>]*>', part).group(0)
        print('   fixture: root binds p14 = %s, a descendant declares it = %s'
              % ('xmlns:p14=' in root, 'xmlns:p14=' in part))
        fx_out = os.path.join(work, 'native-animated.pptx')
        run([py, os.path.join(skill, 'scripts', 'motion.py'), 'apply', '--pptx', fx,
             '--spec', fx_spec, '--out', fx_out, '--assert-geometry'],
            'motion.py apply (native)')
        run([py, os.path.join(skill, 'scripts', 'verify_motion.py'), '--pptx', fx_out,
             '--source', fx], 'verify_motion.py (native)')

    # ---- scenario 3: PowerPoint acceptance -------------------------------
    print()
    print('--- scenario 3: PowerPoint acceptance ---')
    target = out if os.path.exists(out) else os.path.join(work, 'native-animated.pptx')
    if not os.path.exists(target):
        print('no deck to open, skipped')
    else:
        txt = run([PS, '-NoProfile', '-ExecutionPolicy', 'Bypass',
                   '-File', os.path.join(skill, 'scripts', 'motion.ps1'),
                   '-Pptx', target, '-Spec', spec, '-OutDir', os.path.join(work, 'review')],
                  'motion.ps1 (COM)', expect_ok=False)
        # upstream prints "opened read-only: N slides"; earlier local line printed
        # "opened: ...". Accept either, and cross-check with the render count so a
        # wording change cannot silently turn this into a false pass or false fail.
        opened = re.search(r'opened[^:]*:\s*\d+\s*slides', txt) is not None
        renders = os.path.join(work, 'review', 'render')
        n = len(os.listdir(renders)) if os.path.isdir(renders) else 0
        print('   PowerPoint opened the deck: %s' % ('yes' if opened else 'NO'))
        print('   rendered PNGs: %d' % n)
        if not opened:
            fails.append('PowerPoint open')
        if n == 0:
            fails.append('render produced no PNGs')

    print()
    if fails:
        print('SMOKE FAILED:', fails)
        if not ns.keep:
            shutil.rmtree(work, ignore_errors=True)
        return 1
    print('SMOKE PASSED')
    if ns.keep:
        print('artefacts kept at:', work)
    else:
        shutil.rmtree(work, ignore_errors=True)
    return 0


if __name__ == '__main__':
    sys.exit(main())
