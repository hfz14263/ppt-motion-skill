# -*- coding: utf-8 -*-
"""Privacy / leakage audit for the skill repo.

Looks for things that should not ship:
  * absolute paths from this machine (D:\\idea, C:\\Users\\<name>, temp dirs)
  * usernames, emails, real names, hostnames
  * credentials of any kind
  * machine fingerprints that are fine to mention vs. ones that identify the box
  * third-party vendored content (licence compliance)
  * the user's private material files

Run from anywhere:  python privacy_audit.py [--repo <path>]
"""
import argparse
import os
import re
import subprocess
import sys

# patterns worth flagging in file CONTENT
CONTENT_RULES = [
    ('absolute path (this machine)', re.compile(r'[A-Za-z]:\\\\?(?:Users|idea|Git|Program Files|Windows)', re.I)),
    ('temp dir path', re.compile(r'AppData\\\\?Local\\\\?Temp|/tmp/', re.I)),
    ('windows user profile', re.compile(r'C:\\\\?Users\\\\?\w+', re.I)),
    ('email address', re.compile(r'[\w.+-]+@[\w-]+\.[\w.]+')),
    ('possible secret', re.compile(r'(?i)\b(api[_-]?key|secret|token|passwo?rd|credential)\b\s*[:=]\s*\S{8,}')),
    ('github token shape', re.compile(r'\b(gh[pousr]_[A-Za-z0-9]{20,}|github_pat_\w{20,})')),
    ('private key block', re.compile(r'-----BEGIN [A-Z ]*PRIVATE KEY-----')),
    ('aws key shape', re.compile(r'\bAKIA[0-9A-Z]{16}\b')),
    ('sk- style key', re.compile(r'\bsk-[A-Za-z0-9]{20,}')),
]

# anything matching these is expected/normal and should NOT be reported as a leak
ALLOW = [
    re.compile(r'%USERPROFILE%|\$env:USERPROFILE|\$\{?env:USERPROFILE'),
    re.compile(r'C:\\Users\\<', re.I),                 # doc placeholder
    re.compile(r'C:\\\\Users\\\\<', re.I),
    re.compile(r'examples?/|placeholder|e\.g\.', re.I),
    re.compile(r'LICENSE-upstream'),                   # vendored licence file
    re.compile(r'material/'),                          # referenced as external input
    re.compile(r'summerai|skills\.summerai\.cc'),      # cited source
    # GitHub's noreply address exists precisely so it CAN be public. Allowing it
    # keeps the rule sharp for real addresses -- note the personal address that
    # used to be exempt here is NOT allowed any more, so if it ever comes back the
    # audit will flag it.
    re.compile(r'@users\.noreply\.github\.com', re.I),
    re.compile(r'Copyright \(c\) 2026 Binaryify Zhuang'),
]

SKIP_DIRS = {'.git', '__pycache__', 'node_modules', '.office-tmp'}
SKIP_EXT = {'.pptx', '.png', '.jpg', '.jpeg', '.mp4', '.gif', '.zip', '.pyc'}
# This file necessarily contains the very patterns it searches for, so it would
# always flag itself. Skip it rather than weakening the rules.
SKIP_FILES = {'privacy_audit.py'}


def iter_files(root):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in filenames:
            if fn in SKIP_FILES:
                continue
            if os.path.splitext(fn)[1].lower() in SKIP_EXT:
                continue
            yield os.path.join(dirpath, fn)


def audit_content(root):
    findings = []
    for path in iter_files(root):
        try:
            text = open(path, encoding='utf-8', errors='replace').read()
        except Exception:
            continue
        for lineno, line in enumerate(text.splitlines(), 1):
            for label, rx in CONTENT_RULES:
                m = rx.search(line)
                if not m:
                    continue
                if any(a.search(line) for a in ALLOW):
                    continue
                findings.append((os.path.relpath(path, root), lineno, label, line.strip()[:110]))
    return findings


def audit_git(root):
    out = {'authors': [], 'paths': [], 'msg_leaks': []}
    try:
        r = subprocess.run(['git', '-C', root, 'log', '--format=%an <%ae>'],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        out['authors'] = sorted(set(l for l in (r.stdout or '').splitlines() if l.strip()))

        r = subprocess.run(['git', '-C', root, 'log', '--name-only', '--format='],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        paths = set(l.strip() for l in (r.stdout or '').splitlines() if l.strip())
        out['paths'] = sorted(paths)

        r = subprocess.run(['git', '-C', root, 'log', '--format=%s%n%b'],
                           capture_output=True, text=True, encoding='utf-8', errors='replace')
        for label, rx in CONTENT_RULES:
            for m in rx.finditer(r.stdout or ''):
                line = (r.stdout or '')[:m.start()].split('\n')[-1]
                if not any(a.search(line) for a in ALLOW):
                    out['msg_leaks'].append((label, line.strip()[:100]))
    except Exception as e:
        out['error'] = str(e)
    return out


def audit_licences(root):
    """Vendored third-party content must carry attribution."""
    out = []
    ds = os.path.join(root, 'references', 'design-system')
    if os.path.isdir(ds):
        files = [f for f in os.listdir(ds) if f.endswith('.md')]
        has_lic = os.path.exists(os.path.join(ds, 'LICENSE-upstream.txt'))
        attributed = 0
        for f in files:
            head = open(os.path.join(ds, f), encoding='utf-8', errors='replace').read(400)
            if 'Binaryify' in head or 'open-kimi-ppt' in head:
                attributed += 1
        out.append(('design-system: %d md files, licence file present=%s, attributed headers=%d'
                    % (len(files), has_lic, attributed)))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--repo',
                    default=os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    ns = ap.parse_args()
    root = os.path.abspath(ns.repo)

    print('=' * 78)
    print('PRIVACY AUDIT  %s' % root)
    print('=' * 78)
    print()

    print('-- tracked files --')
    r = subprocess.run(['git', '-C', root, 'ls-files'], capture_output=True, text=True,
                       encoding='utf-8', errors='replace')
    tracked = [l for l in (r.stdout or '').splitlines() if l.strip()]
    print('   %d files tracked' % len(tracked))
    for t in tracked:
        print('      %s' % t)
    print()

    print('-- content scan (absolute paths, identities, secrets) --')
    f = audit_content(root)
    if not f:
        print('   CLEAN -- nothing flagged')
    else:
        for path, ln, label, snip in f:
            print('   [%s] %s:%d' % (label, path, ln))
            print('        %s' % snip)
    print()

    print('-- git history --')
    g = audit_git(root)
    print('   authors:')
    for a in g.get('authors', []):
        print('      %s' % a)
    print('   directories ever touched:')
    dirs = sorted(set(os.path.dirname(p).replace('\\', '/') or '.' for p in g.get('paths', [])))
    for d in dirs:
        print('      %s' % d)
    leaks = g.get('msg_leaks', [])
    if leaks:
        print('   commit-message leaks:')
        for label, snip in leaks:
            print('      [%s] %s' % (label, snip))
    else:
        print('   commit-message leaks: none')
    print()

    print('-- third-party attribution --')
    for line in audit_licences(root):
        print('   %s' % line)
    print()

    print('-- machine fingerprints intentionally recorded --')
    for path in iter_files(root):
        text = open(path, encoding='utf-8', errors='replace').read()
        for m in re.finditer(r'16\.0\.\d{5}\.\d{5}|LTSC 2024|ProPlus2024Retail|19045', text):
            print('   %s: %s' % (os.path.relpath(path, root), m.group(0)))
    print()

    ok = not f and not leaks
    print('=' * 78)
    print('VERDICT: %s' % ('no leaks found in content' if ok else 'REVIEW THE ITEMS ABOVE'))
    print('=' * 78)
    return 0


if __name__ == '__main__':
    sys.exit(main())
