# -*- coding: utf-8 -*-
"""Regression test for morph (平滑) transition injection.

Morph was once misdiagnosed as "unsupported on this machine". The real cause was
a wrong element name: `<p:morph/>` does not exist in the p: namespace (morph is a
p159 extension), so PowerPoint silently dropped it and the deck looked like a
version limitation.

This test locks in the correct form so it cannot regress:
  * element is p159:morph inside mc:AlternateContent / mc:Choice Requires="p159"
  * the p159 namespace is the 2015/09 one
  * an mc:Fallback carries a plain fade
  * duration travels in p14:dur
  * option accepts byObject / byWord / byChar and rejects anything else
  * injection places the block BEFORE <p:timing> (CT_Slide is a sequence)
  * re-applying does not leave a stale AlternateContent wrapper behind
"""
import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'scripts'))
import motion as M  # noqa: E402

FAILED = []


def check(label, cond, extra=''):
    print('  %-56s %s%s' % (label, 'ok' if cond else 'FAIL', ('  ' + extra) if extra else ''))
    if not cond:
        FAILED.append(label)


print('morph transition generation')

b = M.build_transition('morph', 2.0)
check('wrapped in mc:AlternateContent', b.startswith('<mc:AlternateContent'))
check('has mc:Choice Requires="p159"', 'Requires="p159"' in b)
check('p159 namespace is 2015/09',
      'http://schemas.microsoft.com/office/powerpoint/2015/09/main' in b)
check('element is p159:morph (not p:morph)',
      '<p159:morph option="byObject"/>' in b and '<p:morph' not in b)
check('has mc:Fallback with fade', '<mc:Fallback>' in b and '<p:fade/>' in b)
check('duration in p14:dur', 'p14:dur="2000"' in b)
check('p14 declared locally', 'xmlns:p14=' in b)

for opt in ('byObject', 'byWord', 'byChar'):
    bb = M.build_transition('morph', 1.5, option=opt)
    check('option %s honoured' % opt, ('option="%s"' % opt) in bb)

for low, want in (('byobject', 'byObject'), ('BYWORD', 'byWord'), ('bychar', 'byChar')):
    bb = M.build_transition('morph', 1.0, option=low)
    check('option %r normalised to %s' % (low, want), ('option="%s"' % want) in bb)

try:
    M.build_transition('morph', 1.0, option='bogus')
    check('invalid option rejected', False)
except ValueError:
    check('invalid option rejected', True)

check('default option is byObject',
      'option="byObject"' in M.build_transition('morph'))

print()
print('injection into a real slide')

SLIDE = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
         '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
         ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
         ' xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
         '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
         '</p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>'
         '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')

# simulate a slide that already has timing (animation) in it
with_timing = SLIDE.replace('</p:sld>',
                            '<p:timing><p:tnLst/></p:timing></p:sld>')
out = M.replace_or_insert(M.drop_alternate_content(with_timing), 'transition',
                          M.build_transition('morph', 2.0))
i_trans = out.find('<mc:AlternateContent')
i_timing = out.find('<p:timing>')
check('morph present', i_trans != -1)
check('morph placed BEFORE <p:timing>', i_trans != -1 and i_timing != -1 and i_trans < i_timing,
      'transition@%s timing@%s' % (i_trans, i_timing))

# re-apply: the stale wrapper must not accumulate
once = M.replace_or_insert(M.drop_alternate_content(SLIDE), 'transition',
                           M.build_transition('morph', 2.0))
twice = M.drop_alternate_content(once)
twice = M.replace_or_insert(twice, 'transition', M.build_transition('morph', 2.0))
check('re-apply leaves exactly one AlternateContent',
      twice.count('<mc:AlternateContent') == 1,
      'count=%d' % twice.count('<mc:AlternateContent'))
check('re-apply leaves exactly one p159:morph',
      twice.count('<p159:morph') == 1, 'count=%d' % twice.count('<p159:morph'))

# and it must still be well-formed XML
try:
    from lxml import etree
    etree.fromstring(out.encode('utf-8'))
    check('injected slide is well-formed XML', True)
except ImportError:
    check('injected slide is well-formed XML (lxml absent, skipped)', True)
except Exception as e:
    check('injected slide is well-formed XML', False, str(e)[:80])

print()
if FAILED:
    print('MORPH TEST FAILED: %s' % FAILED)
    sys.exit(1)
print('morph test passed')
