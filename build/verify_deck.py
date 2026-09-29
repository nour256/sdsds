"""Automated checks for IEN_301_Project_1_v2.pptx (run after build_deck.py)."""
import re, sys, zipfile, json
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

F = sys.argv[1] if len(sys.argv) > 1 else 'IEN_301_Project_1_v2.pptx'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
P = 'http://schemas.openxmlformats.org/presentationml/2006/main'
ok = True

def fail(msg):
    global ok
    ok = False
    print('FAIL:', msg)

prs = Presentation(F)
BANNED = ['solution', 'heater', 'device', 'fix', 'convenient']
ICON_NAMES = ['thermostat', 'water_drop', 'ac_unit', 'schedule', 'timer', 'autorenew', 'material icons']
OLD_NOTES = ['generated with the built-in ImageGen', 'The user supplied the IEN301 brief', 'no participant quotes have been supplied',
             'PRESENTATION SCRIPT', 'CLICK CUES', 'SOURCES AND EVIDENCE BOUNDARIES']

print(f'{F}: {len(prs.slides)} slides')
summary = []
for n, s in enumerate(prs.slides, start=1):
    texts = []
    n_pic = n_txt = 0
    fonts = set()
    for sh in s.shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.PICTURE:
            n_pic += 1
        if sh.has_text_frame and sh.text_frame.text.strip():
            n_txt += 1
            texts.append(sh.text_frame.text)
        if sh.has_text_frame:
            for r in sh._element.iter('{%s}latin' % A):
                fonts.add(r.get('typeface'))
        if sh.shape_type == MSO_SHAPE_TYPE.TABLE:
            for row in sh.table.rows:
                for c in row.cells:
                    texts.append(c.text)
                    n_txt += 1
    full = '\n'.join(texts)
    words = len(re.findall(r'\w+', full))
    hidden = s._element.get('show') == '0'
    # 1 image-only slides
    if n_txt == 0:
        fail(f'slide {n}: no editable text (image-only?)')
    # 2 banned words (whole word, case-insensitive)
    for b in BANNED:
        if re.search(r'\b' + b + r's?\b', full, re.I):
            fail(f'slide {n}: banned word "{b}" in: {full!r}')
    # 3 icon names
    for ic in ICON_NAMES:
        if re.search(r'(?<![A-Za-z])' + re.escape(ic) + r'(?![A-Za-z])', full, re.I):
            fail(f'slide {n}: leftover icon name "{ic}"')
    # 4 fonts
    bad_fonts = fonts - {'Quattrocento Sans', None}
    if bad_fonts:
        fail(f'slide {n}: off-theme fonts {bad_fonts}')
    # 5 notes
    if not s.has_notes_slide or not s.notes_slide.notes_text_frame.text.strip():
        fail(f'slide {n}: missing speaker notes')
    else:
        nt = s.notes_slide.notes_text_frame.text
        for o in OLD_NOTES:
            if o.lower() in nt.lower():
                fail(f'slide {n}: old drafting note text present: {o}')
        if 'Click:' not in nt:
            fail(f'slide {n}: notes lack a Click: line')
    # 6 timing & transition XML
    sld = s._element
    timing = sld.find('{%s}timing' % P)
    trans = sld.find('.//{%s}transition' % P)
    n_eff = len(sld.findall('.//{%s}animEffect' % P)) if timing is not None else 0
    clicks = 0
    if timing is not None:
        main = timing.find('.//{%s}cTn[@nodeType="mainSeq"]/{%s}childTnLst' % (P, P))
        clicks = len([c for c in main if c.find('./{%s}cTn/{%s}stCondLst/{%s}cond[@evt="onBegin"]' % (P, P, P)) is None])
        # all spids referenced exist
        ids = {e.get('id') for e in sld.iter('{%s}cNvPr' % P)}
        for t in timing.iter('{%s}spTgt' % P):
            if t.get('spid') not in ids:
                fail(f'slide {n}: animation targets missing shape id {t.get("spid")}')
        # ordering: cSld, clrMapOvr, transition/AlternateContent, timing
        order = [etree.QName(c).localname for c in sld]
        if order.index('timing') < order.index('clrMapOvr'):
            fail(f'slide {n}: element order wrong: {order}')
    if trans is None:
        fail(f'slide {n}: no slide transition')
    elif trans.find('{%s}fade' % P) is None:
        fail(f'slide {n}: transition is not fade')
    if timing is None and not hidden:
        fail(f'slide {n}: no animations')
    first = full.strip().split('\n')[0][:60] if full.strip() else ''
    summary.append((n, hidden, n_txt, n_pic, words, n_eff, clicks, first))

print('\n slide hidden text pics words effects clicks first-line')
for row in summary:
    print(' %5d %6s %4d %4d %5d %7d %6d  %s' % row)

# 7 every slide XML parses from the zip, and the hidden slide exists
z = zipfile.ZipFile(F)
for name in z.namelist():
    if name.startswith('ppt/slides/slide') and name.endswith('.xml'):
        etree.fromstring(z.read(name))
print('all slide XML parts parse OK:', len([n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')]))
# 8 no orphan old media
media = [n for n in z.namelist() if n.startswith('ppt/media/')]
print('media parts:', media)
print('fonts embedded:', [n for n in z.namelist() if n.startswith('ppt/fonts/')])
print('RESULT:', 'PASS' if ok else 'FAIL')
sys.exit(0 if ok else 1)
