"""Rebuild IEN_301_Project_1.pptx -> IEN_301_Project_1_v2.pptx (problem framing only).

Keeps original slides 2, 3(bg), 4, 5 and their exact theme; rebuilds everything else natively.
Run:  python3 build_deck.py  (from the scratchpad directory)
"""
import json, re, sys, copy
from lxml import etree
from pptx import Presentation
from deck_lib import *

SRC = 'orig.pptx'
OUT = 'IEN_301_Project_1_v2.pptx'
MEDIA = 'orig_unpacked/ppt/media/'
IMG_SHOWER_DRAIN = MEDIA + 'image1.png'   # slide 2 photo (dark)
IMG_BUCKET = MEDIA + 'image7.png'         # slide 9 photo (light, shower + bucket)
IMG_DUSK = MEDIA + 'image2.png'           # slide 8 photo (dusk landscape; RV cropped out)

TEAM = ['Omar Hanifa', 'Mohamed Alkaabi', 'Khaled Kittaneh', 'Nour Helal Yousifi', 'Mohamed Haque']

# ----------------------------------------------------------------------------- numbers (computed)
FLOW = 8.7                 # L/min, RSB 2015 audit via ADDC/Tarsheed 2017
WAITS = {'low': 3, 'mid': 4, 'high': 5}   # minutes reported: I1 3-4 min, I2 3-5 min
SHOWERS_PER_DAY = 1
per_shower = {k: FLOW * v * SHOWERS_PER_DAY for k, v in WAITS.items()}
per_year = {k: v * 365 for k, v in per_shower.items()}
V1_PER_SHOWER = FLOW * 0.5
NUMBERS = {'per_shower': per_shower, 'per_year': per_year, 'v1_per_shower': V1_PER_SHOWER,
           'ratio': (WAITS['low'] * 60 / 30, WAITS['high'] * 60 / 30)}
print('Computed numbers:', json.dumps(NUMBERS, indent=1))


def fmt_int(x):
    return f'{int(round(x, -2)):,}'   # round to nearest 100 for the slide (exact values in evidence.md)


# ----------------------------------------------------------------------------- notes
NOTES = {}
CLAIMS = []   # (slide, claim, interview, timestamp, verbatim)


def note(n, spoken, click, cite=None):
    txt = spoken.strip()
    txt += '\n\nClick: ' + click.strip()
    if cite:
        txt += '\n\nRecording refs: ' + cite.strip()
    NOTES[n] = txt
    return txt


# ============================================================================= build
prs = Presentation(SRC)
orig = list(prs.slides)
layout = prs.slide_masters[1].slide_layouts[0]   # empty "TITLE" layout used by slides 2-10

# ---- helper to find a shape by original id on kept slides
def sp_by_id(slide, sid):
    for sh in slide.shapes:
        if sh.shape_id == sid:
            return sh
    raise KeyError(sid)


def replace_text(slide, sid, old, new):
    sh = sp_by_id(slide, sid)
    found = False
    for t in sh._element.iter('{%s}t' % A):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            found = True
    assert found, (sid, old)


def move(slide, sid, y=None, x=None, h=None, w=None):
    sh = sp_by_id(slide, sid)
    if y is not None: sh.top = emu(y)
    if x is not None: sh.left = emu(x)
    if h is not None: sh.height = emu(h)
    if w is not None: sh.width = emu(w)


# ============================================================================= SLIDE 1 - TITLE (new)
s1 = prs.slides.add_slide(layout)
set_background(s1, DARK_BG)
picture(s1, IMG_BUCKET, 5.81, 0, 4.19, 5.625, crop=(0.0, 0.05, 0.0, 0.05), name='Photo: shower and bucket')
# soften the photo edge with a dark overlay strip so the light photo reads as a dark hero slide
s1_ov = shape(s1, 'rect', 5.81, 0, 4.19, 5.625, fill=DARK_BG, alpha=38, name='Photo overlay')
s1_lab = label(s1, 'IEN301  ·  PROJECT 1: PROBLEM FRAMING', dark=True)
s1_t1 = textbox(s1, 0.61, 1.30, 5.1, 0.95, [para([('FIRST DROP', 54, WHITE, True)])], name='Title')
s1_t2 = textbox(s1, 0.61, 2.30, 5.1, 1.30, [para([('The water that runs', 30, SKY, True)]), para([('before the shower begins', 30, SKY, True)])], name='Subtitle')
s1_team = textbox(s1, 0.62, 3.95, 5.1, 0.95,
                  [para([('TEAM', 11, SKY, True)], spc_aft=4),
                   para([(' · '.join(TEAM[:3]), 12, MIST)]),
                   para([(' · '.join(TEAM[3:]), 12, MIST)])], name='Team')
s1_foot = footer(s1, 'Problem framing only. No product or intervention is proposed in this deck.', dark=True)
s1_num = page_number(s1, 1, dark=True)
set_transition(s1)
Timing().add((s1_lab, 'fade'), (s1_t1, 'fade'), auto=True).add((s1_t2, 'fade')).add((s1_team, 'fade')).apply(s1)
note(1, "Good morning. We are team First Drop, and this is our problem framing for IEN301. "
        "Our topic is a small moment everyone with a warm shower knows: the water is running, but nobody is washing yet. "
        "Today we will define who lives this moment, show what two interviewees told us about it, and explain why it persists.",
     "title and label appear with the slide. Click 1: subtitle. Click 2: team names.")

# ============================================================================= SLIDE 2 - THE MOMENT (kept as-is)
s2 = orig[1]
set_transition(s2)
Timing().add((64, 'fade'), (65, 'fade'), auto=True).add((66, 'fade')).apply(s2)
note(2, "Picture turning on your shower. Water is flowing. Washing has not begun. "
        "That interval, between the tap and the first wash, is the whole of our project. "
        "Everything on the next slides is about what happens to the water and to the person during those minutes.",
     "'Water is flowing' is visible on entry. Click 1: 'Washing has not begun'.")

# ============================================================================= SLIDE 3 - USER & CONTEXT (rebuilt on the original slide)
s3 = orig[2]
clear_shapes(s3)
set_background(s3, LIGHT_BG)
s3_lab = label(s3, 'USER & CONTEXT DEFINITION')
s3_hd = headline(s3, ['Adults in UAE homes', 'who shower with warm water', 'and wait 3–5 minutes for it to arrive'], y=1.05, sz=30)
s3_sub = textbox(s3, 0.62, 3.05, 8.7, 0.42, [para([('Situation: the tap is on, cold water runs to the drain, and they wait', 17, BLUE, True)])], name='Situation')
s3_body = textbox(s3, 0.62, 3.50, 8.7, 1.05,
                  [para([('Outside the bathroom  (“I’m waiting outside” – Interviewee 1)', 14, GREY)], spc_aft=3),
                   para([('or standing by  (“I just let the water run… and just sit there” – Interviewee 2)', 14, GREY)], spc_aft=3),
                   para([('Context: UAE / MENA homes with piped hot water; interviewed 27 Sept 2026', 14, GREY)])], name='Context')
s3_foot = footer(s3, 'Based on 2 recorded interviews. Home type, water-heating setup and tenure were not asked yet — to confirm in interview 3.')
s3_num = page_number(s3, 3)
set_transition(s3)
Timing().add((s3_lab, 'fade'), (s3_hd, 'fade'), auto=True).add((s3_sub, 'fade')).add((s3_body, 'wipeLeft')).apply(s3)
note(3, "Our user is not 'people' or 'residents' in general. It is adults in UAE homes who shower with warm water and wait three to five minutes for it to arrive. "
        "Both interviewees gave us that range unprompted. The situation is the same for both: the tap is on, cold water runs to the drain, and they wait. "
        "How they wait differs, and that matters later.",
     "headline on entry. Click 1: situation line. Click 2: the two waiting behaviours and context.",
     "I1 00:22 '3 to 4 minutes'; I2 00:16 '3 to 5 minutes'; I1 00:45; I2 00:40.")
CLAIMS += [(3, 'Wait 3–4 min', 1, '00:22', 'About 3 to 4 minutes.'),
           (3, 'Wait 3–5 min', 2, '00:16', 'Three to five minutes.'),
           (3, 'Waits outside the bathroom', 1, '00:45', "I'm waiting outside."),
           (3, 'Stays in the bathroom', 2, '00:40', 'I just let the water run until it gets warm and just sit there.')]

# ============================================================================= SLIDE 4 - WHY WATER RUNS (kept, one wording edit)
s4 = orig[3]
replace_text(s4, 88, 'from the heater to the shower', 'from where it is heated to the shower')
set_transition(s4)
Timing().add((85, 'fade'), (86, 'fade'), (87, 'fade'), (88, 'fade'), auto=True).add((89, 'fade'), (90, 'fade'), (91, 'fade')).apply(s4)
note(4, "Why does water run before use? First, the water sitting in the pipe has cooled since the last shower, and it has to clear before hot water arrives. "
        "Second, once the water is warm, the person may not be there to notice. Both intervals send clean water to the drain.",
     "stage 01 on entry. Click 1: stage 02.")

# ============================================================================= SLIDE 5 - ILLUSTRATIVE CALCULATION (kept + interview line)
s5 = orig[4]
move(s5, 104, y=3.72, h=0.42)
s5_link = textbox(s5, 0.62, 4.20, 8.67, 0.66,
                  [para([('Interviews reported 3–5 minute waits, not 30 seconds.', 15, BLUE, True)], spc_aft=2),
                   para([('8.7 L/min × 4 min = about 35 L per shower', 15, BLUE, True)])],
                  name='Interview link')
set_transition(s5)
Timing().add((99, 'fade'), (100, 'fade'), (101, 'fade'), auto=True).add((102, 'flyUpFade'), (103, 'fade'), (104, 'fade')).add((s5_link, 'fade')).apply(s5)
note(5, "Before the interviews, we assumed a thirty-second wait. At the RSB flow rate of 8.7 litres a minute, that is 4.35 litres per shower. "
        "Then we listened. Both interviewees reported three to five minutes, so the same flow rate gives about 35 litres per shower. "
        "The number is illustrative, but the order of magnitude changed.",
     "'30 seconds' on entry. Click 1: 4.35 litres and the formula. Click 2: the interview line.",
     "I1 00:22; I2 00:16. 8.7 × 4 = 34.8 L.")
CLAIMS += [(5, 'Reported waits 3–5 min; ~35 L per shower at 4 min', 1, '00:22', 'About 3 to 4 minutes.')]

# ============================================================================= SLIDE 6 - EVIDENCE: METHOD (new)
s6 = prs.slides.add_slide(layout)
set_background(s6, LIGHT_BG)
s6_lab = label(s6, 'EVIDENCE — HOW WE LISTENED')
s6_n1 = textbox(s6, 0.55, 1.05, 3.0, 1.2, [para([('2', 84, INK, True)])], name='Big number interviews')
s6_l1 = textbox(s6, 0.62, 2.30, 3.4, 0.6, [para([('recorded interviews', 17, GREY)]), para([('27 September 2026', 14, GREY)])], name='Label interviews')
s6_n2 = textbox(s6, 3.75, 1.05, 3.0, 1.2, [para([('5:15', 84, BLUE, True)])], name='Big number minutes')
s6_l2 = textbox(s6, 3.82, 2.30, 3.4, 0.6, [para([('minutes of recording', 17, BLUE)]), para([('face-to-face, phone-recorded', 14, GREY)])], name='Label minutes')
s6_method = textbox(s6, 7.0, 1.15, 2.4, 1.9,
                    [para([('SAME 5 QUESTIONS', 11, BLUE, True)], spc_aft=4),
                     para([('Warm or cold?', 13, GREY)], spc_aft=2), para([('How long do you wait?', 13, GREY)], spc_aft=2),
                     para([('How much is lost?', 13, GREY)], spc_aft=2), para([('What do you do meanwhile?', 13, GREY)], spc_aft=2),
                     para([('What happens to the water?', 13, GREY)])], name='Questions')
s6_who = textbox(s6, 0.62, 3.25, 8.7, 1.5,
                 [para([('WHO WE INTERVIEWED', 11, BLUE, True)], spc_aft=6),
                  para([('Interviewee 1 — “Mr Ali”: ', 13, INK, True), ('older man · warm showers · waits 3–4 min · waits outside the bathroom · nothing collected', 12, GREY)], spc_aft=7),
                  para([('Interviewee 2 — “Ryan”: ', 13, INK, True), ('young man · warm showers · waits 3–5 min · stays in the bathroom · nothing collected', 12, GREY)], spc_aft=7),
                  para([('Interviewee 3: ', 13, INK, True), ('[PLACEHOLDER — recording not received; add descriptor when transcribed]', 12, GREY, False, True)])], name='Interviewees')
s6_foot = footer(s6, 'Home type, water-heating setup and tenure were not asked. Recordings and timestamps are available for Q&A.')
s6_num = page_number(s6, 6)
set_transition(s6)
Timing().add((s6_lab, 'fade'), (s6_n1, 'fade'), (s6_l1, 'fade'), auto=True).add((s6_n2, 'flyUpFade'), (s6_l2, 'fade')).add((s6_method, 'fade')).add((s6_who, 'wipeLeft')).apply(s6)
note(6, "Our evidence comes from two recorded, face-to-face interviews on 27 September, about two and a half minutes each, five minutes in total. "
        "We asked both people the same five questions. Mr Ali is an older man who waits outside the bathroom; Ryan is a young man who stays and waits. "
        "A third interview is planned and its findings are not on these slides yet.",
     "'2 interviews' on entry. Click 1: '5:15 minutes'. Click 2: the five questions. Click 3: who we interviewed.",
     "Video files 16.41.50 (2:48) and 16.42.14 (2:27).")

# ============================================================================= SLIDE 7 - EVIDENCE: QUOTES (new)
s7 = prs.slides.add_slide(layout)
set_background(s7, LIGHT_BG)
s7_lab = label(s7, 'EVIDENCE — IN THEIR WORDS')
s7_hd = textbox(s7, 0.59, 0.78, 8.9, 0.5, [para([('Four things we heard', 26, INK, True)])], name='Headline')
QUOTES = [
    ('“About 3 to 4 minutes.”', 'Interviewee 1 · older man, warm showers · 00:22', 'How long do you wait for warm water?'),
    ('“I just let the water run until it gets warm and just sit there.”', 'Interviewee 2 · young man · 00:40', 'What do you do while you wait?'),
    ('“I’m waiting outside.”', 'Interviewee 1 · 00:45', 'What do you do while you wait?'),
    ('“No idea, but… I would say at least a bucket.”', 'Interviewee 1 · 00:34', 'How much water is lost?'),
]
s7_cards = []
cw, ch, gx, gy = 4.20, 1.62, 0.28, 0.22
for i, (q, who, ask) in enumerate(QUOTES):
    cx = 0.62 + (i % 2) * (cw + gx)
    cy = 1.45 + (i // 2) * (ch + gy)
    card = shape(s7, 'rect', cx, cy, cw, ch, fill=CARD, line=(RULE, 7150), name=f'Quote card {i+1}')
    mark = textbox(s7, cx + 0.15, cy + 0.02, 0.5, 0.6, [para([('“', 44, BLUE, True)])], name=f'Quote mark {i+1}')
    qt = textbox(s7, cx + 0.52, cy + 0.18, cw - 0.7, 0.85, [para([(q, 16, INK, True)], ln_spc=1.05)], name=f'Quote {i+1}')
    at = textbox(s7, cx + 0.52, cy + ch - 0.52, cw - 0.7, 0.45, [para([(ask, 10, GREY, False, True)], spc_aft=1), para([('— ' + who, 10, GREY)])], name=f'Attribution {i+1}')
    s7_cards.append((card, mark, qt, at))
s7_foot = footer(s7, 'Verbatim from the recordings (Whisper transcript, cross-checked). Timestamps are mm:ss in each video.')
s7_num = page_number(s7, 7)
set_transition(s7)
t = Timing().add((s7_lab, 'fade'), (s7_hd, 'fade'), auto=True)
for card, mark, qt, at in s7_cards:
    t.add((card, 'fade'), (mark, 'fade'), (qt, 'fade'), (at, 'fade'))
t.apply(s7)
note(7, "Here is what we heard, word for word. Mr Ali: about three to four minutes. Ryan: I just let the water run until it gets warm and just sit there. "
        "Mr Ali again: I'm waiting outside. And when we asked how much is lost: no idea, at least a bucket. "
        "Nobody measures the wait, and nobody is doing anything with the water.",
     "headline on entry. Click 1–4: one quote per click, left to right, top to bottom.",
     "I1 00:22; I2 00:40; I1 00:45; I1 00:34. Also I2 00:16 '3 to 5 minutes', I1 00:56 'goes down the drain', I2 00:30 'probably a lot'.")
CLAIMS += [(7, 'Quote: 3–4 minutes', 1, '00:22', 'About 3 to 4 minutes.'),
           (7, 'Quote: let it run and sit there', 2, '00:40', 'I just let the water run until it gets warm and just sit there.'),
           (7, 'Quote: waiting outside', 1, '00:45', "I'm waiting outside."),
           (7, 'Quote: at least a bucket', 1, '00:34', "No idea, but I think it's about… I would say at least a bucket.")]

# ============================================================================= SLIDE 8 - WHAT SURPRISED US (new, dark like slide 4)
s8 = prs.slides.add_slide(layout)
set_background(s8, DARK_BG)
s8_lab = label(s8, 'EVIDENCE — WHAT WE DID NOT KNOW AT THE START', dark=True)
SURPRISES = [
    ('Minutes, not seconds', ['We had assumed about 30 seconds.', 'Both interviewees said 3–5 minutes: six to ten times longer.']),
    ('Nobody measures the loss', ['“No idea… at least a bucket.”  “I don’t know the exact number.”', 'The run-off is invisible while it happens.']),
    ('Waiting looks different in each home', ['One leaves the room (“I’m waiting outside”); one stays (“just sit there”).', 'Either way, the water goes down the drain.']),
]
s8_items = []
for i, (title, body) in enumerate(SURPRISES):
    y = 0.98 + i * 1.38
    num = textbox(s8, 0.61, y, 1.09, 0.73, [para([(f'0{i+1}', 38, SKY, True)])], name=f'Number {i+1}')
    tt = textbox(s8, 1.93, y + 0.04, 7.42, 0.5, [para([(title, 24, WHITE, True)])], name=f'Surprise title {i+1}')
    bb = textbox(s8, 1.95, y + 0.58, 7.3, 0.75, [para([(body[0], 15, MIST)]), para([(body[1], 15, MIST)])], name=f'Surprise body {i+1}')
    s8_items.append((num, tt, bb))
s8_foot = footer(s8, 'Winter effect, home type and water-heating setup were not covered in the two recordings — to test in interview 3.', dark=True)
s8_num = page_number(s8, 8, dark=True)
set_transition(s8)
t = Timing().add((s8_lab, 'fade'), auto=True)
for num, tt, bb in s8_items:
    t.add((num, 'fade'), (tt, 'fade'), (bb, 'fade'))
t.apply(s8)
note(8, "Three things we did not know at the start. One: the wait is minutes, not seconds; our thirty-second assumption was off by six to ten times. "
        "Two: nobody measures the loss; both people could only guess. "
        "Three: waiting looks different in each home. Mr Ali leaves the bathroom; Ryan stays. In both cases the water goes down the drain.",
     "label on entry. Click 1–3: one finding per click.",
     "I1 00:22, I2 00:16 (wait); I1 00:34, I2 00:30 (no idea); I1 00:45, I2 00:40 (behaviour); I1 00:56, I2 00:52 (drain).")
CLAIMS += [(8, 'Wait 6–10× the 30 s assumption', 1, '00:22', 'About 3 to 4 minutes.'),
           (8, 'Nobody measures the loss', 2, '00:30', "I don't know the exact number, but probably a lot. [unclear]"),
           (8, 'Water goes down the drain (both)', 2, '00:52', 'Nothing, it just… / Yeah.')]

# ============================================================================= SLIDE 9 - KEY NUMBERS (new)
s9 = prs.slides.add_slide(layout)
set_background(s9, LIGHT_BG)
s9_lab = label(s9, 'KEY NUMBERS — ILLUSTRATIVE')
s9_hd = textbox(s9, 0.59, 0.78, 8.9, 0.5, [para([('One shower’s warm-up, scaled to a year', 26, INK, True)])], name='Headline')
s9_formula = textbox(s9, 0.62, 1.38, 8.7, 0.42, [para([('8.7 L/min  ×  wait (min)  ×  1 shower/day  ×  365 days', 19, INK)])], name='Formula')
COLS = [('low', '3 min wait', 'low end reported', INK), ('mid', '4 min wait', 'midpoint of 3–5 min', BLUE), ('high', '5 min wait', 'high end reported', INK)]
s9_cols = []
for i, (k, ttl, sub, col) in enumerate(COLS):
    x = 0.55 + i * 2.95
    big = textbox(s9, x, 2.05, 2.85, 1.05, [para([(fmt_int(per_year[k]), 54, col, True)])], name=f'Big number {k}')
    unit = textbox(s9, x + 0.07, 3.08, 2.8, 0.35, [para([('litres per year', 17, col)])], name=f'Unit {k}')
    det = textbox(s9, x + 0.07, 3.50, 2.8, 0.85, [para([(ttl, 14, INK, True)], spc_aft=2), para([(sub, 12, GREY)], spc_aft=2), para([(f'{per_shower[k]:.1f} L per shower', 12, GREY)])], name=f'Detail {k}')
    s9_cols.append((big, unit, det))
s9_cmp = textbox(s9, 0.62, 4.50, 8.7, 0.45, [para([('For comparison, the 30-second wait we first assumed would be about ', 12, GREY), (f'{int(round(V1_PER_SHOWER*365)):,} L per year', 12, GREY, True), ('.', 12, GREY)])], name='Comparison')
s9_foot = footer(s9, 'Illustrative. Flow: RSB 2015 audit (via ADDC/Tarsheed 2017). Wait times: interviews 1–2. One shower/day assumed. Winter not measured.')
s9_num = page_number(s9, 9)
set_transition(s9)
t = Timing().add((s9_lab, 'fade'), (s9_hd, 'fade'), (s9_formula, 'fade'), auto=True)
for big, unit, det in s9_cols:
    t.add((big, 'flyUpFade'), (unit, 'fade'), (det, 'fade'))
t.add((s9_cmp, 'fade'))
t.apply(s9)
note(9, "We replaced the unsourced yearly figures from our first draft with a calculation you can check on the slide. "
        "8.7 litres a minute, times the reported wait, times one shower a day, times 365. "
        "Three minutes gives about nine and a half thousand litres a year; four minutes about twelve and a half thousand; five minutes about sixteen thousand. "
        "Our original thirty-second assumption would have been under sixteen hundred.",
     "formula on entry. Click 1–3: the three yearly figures. Click 4: the comparison line.",
     "Wait times: I1 00:22, I2 00:16. Computation in evidence.md §4.")
CLAIMS += [(9, 'Wait range 3–5 min used in calculation', 2, '00:16', 'Three to five minutes.')]

# ============================================================================= SLIDE 10 - POV STATEMENT (new)
s10 = prs.slides.add_slide(layout)
set_background(s10, LIGHT_BG)
s10_lab = label(s10, 'POINT OF VIEW')
POV_PARTS = [('Adults in UAE homes who shower with warm water', INK), (' need a way to ', BLUE),
             ('stop the 3–5 minute warm-up from sending clean water down the drain', INK), (' because ', BLUE),
             ('the wait is passive time — they step out or wait it out — and nobody sees how much is lost.', INK)]
s10_pov = textbox(s10, 0.59, 0.85, 8.9, 2.25, [para([(txt, 22, col, True) for txt, col in POV_PARTS], ln_spc=1.08)], name='POV statement')
# 4-step flow as native shapes
STEPS = [('Turn on the shower', 'power'), ('Cold water runs to the drain', 'drop'), ('Wait 3–5 minutes', 'clock'), ('Shower', 'shower')]
s10_steps = []
sx0, sy, sw = 0.62, 3.45, 2.05
for i, (cap, icon) in enumerate(STEPS):
    x = sx0 + i * (sw + 0.24)
    ids = []
    cx, cy, r = x + 0.36, sy + 0.36, 0.36
    ids.append(shape(s10, 'ellipse', x, sy, 0.72, 0.72, fill=CARD, line=(RULE, 7150), name=f'Icon circle {i+1}'))
    if icon == 'power':
        ids.append(shape(s10, 'blockArc', cx - 0.2, cy - 0.2, 0.4, 0.4, fill=BLUE, adj={'adj1': 18000000, 'adj2': 14400000, 'adj3': 14000}, name='Icon power arc'))
        ids.append(shape(s10, 'roundRect', cx - 0.03, cy - 0.24, 0.06, 0.24, fill=BLUE, name='Icon power bar'))
    elif icon == 'drop':
        ids.append(shape(s10, 'teardrop', cx - 0.16, cy - 0.18, 0.32, 0.32, fill=BLUE, rot=18900000, name='Icon drop'))
    elif icon == 'clock':
        ids.append(shape(s10, 'ellipse', cx - 0.2, cy - 0.2, 0.4, 0.4, line=(BLUE, 22225), name='Icon clock face'))
        ids.append(line(s10, cx, cy, cx, cy - 0.13, BLUE, 22225, name='Icon clock hand 1'))
        ids.append(line(s10, cx, cy, cx + 0.11, cy, BLUE, 22225, name='Icon clock hand 2'))
    elif icon == 'shower':
        ids.append(shape(s10, 'roundRect', cx - 0.03, cy - 0.24, 0.06, 0.14, fill=BLUE, name='Icon shower stem'))
        ids.append(shape(s10, 'roundRect', cx - 0.17, cy - 0.11, 0.34, 0.08, fill=BLUE, name='Icon shower head'))
        for j, dx in enumerate((-0.11, 0.0, 0.11)):
            ids.append(line(s10, cx + dx, cy + 0.02, cx + dx, cy + 0.20, BLUE, 19050, name=f'Icon shower drop {j+1}'))
    ids.append(textbox(s10, x - 0.1, sy + 0.82, sw - 0.1, 0.5, [para([(cap, 12, GREY)])], name=f'Step caption {i+1}'))
    if i < 3:
        ids.append(shape(s10, 'chevron', x + sw - 0.10, sy + 0.24, 0.16, 0.24, fill=MIST, name=f'Step arrow {i+1}'))
    s10_steps.append(ids)
s10_foot = footer(s10, 'Grounded in interviews 1–2 (27 Sept 2026): waits of 3–4 and 3–5 min; “I’m waiting outside”; “just sit there”; “no idea… at least a bucket”.')
s10_num = page_number(s10, 10)
set_transition(s10)
t = Timing().add((s10_lab, 'fade'), auto=True).add((s10_pov, 'fade'))
for ids in s10_steps:
    t.add(*[(i, 'fade') for i in ids])
t.apply(s10)
note(10, "This brings us to our point of view. Adults in UAE homes who shower with warm water need a way to stop the three-to-five-minute warm-up from sending clean water down the drain, "
         "because the wait is passive time, they step out or wait it out, and nobody sees how much is lost. "
         "The sequence is always the same: turn on, cold water drains, wait, shower.",
     "label on entry. Click 1: the POV statement. Click 2–5: the four steps of the flow.",
     "I1 00:22, 00:34, 00:45, 00:56; I2 00:16, 00:30, 00:40, 00:52.")
CLAIMS += [(10, 'POV: 3–5 minute warm-up', 1, '00:22', 'About 3 to 4 minutes.'),
           (10, 'POV: they step out', 1, '00:45', "I'm waiting outside."),
           (10, 'POV: or wait it out', 2, '00:40', 'I just let the water run until it gets warm and just sit there.'),
           (10, 'POV: nobody sees how much is lost', 1, '00:34', "No idea… I would say at least a bucket.")]

# ============================================================================= SLIDE 11 - CONSTRAINTS (new)
s11 = prs.slides.add_slide(layout)
set_background(s11, LIGHT_BG)
s11_lab = label(s11, 'CONSTRAINTS — WHY THE PROBLEM PERSISTS')
s11_hd = textbox(s11, 0.59, 0.78, 8.9, 0.5, [para([('Why the water keeps running', 26, INK, True)])], name='Headline')
CARDS = [
    ('HABIT', 'Turn it on, then wait — outside the room or standing by. The run-off is never handled.', 'Interviews 1–2'),
    ('COST SIGNAL', 'Neither person could estimate the loss (“at least a bucket”). Nothing in the bill or the bathroom shows it.', 'Interviews 1–2 · context'),
    ('SYSTEMS', 'Water is heated far from the shower; the pipe must clear every time. Re-plumbing is rarely an option.', 'Context (EPA 2026) · to confirm'),
    ('CLIMATE', 'Colder inlet water in winter should lengthen the wait. Not covered in either interview.', 'To verify — interview 3'),
    ('ROUTINE', 'The wait sits inside getting ready; extra steps compete with time. Not raised by either interviewee.', 'Assumption — to verify'),
]
s11_cards = []
cw, ch, gx = 1.68, 2.80, 0.115
for i, (ttl, body, tag) in enumerate(CARDS):
    x = 0.62 + i * (cw + gx)
    y = 1.45
    card = shape(s11, 'rect', x, y, cw, ch, fill=CARD, line=(RULE, 7150), name=f'Card {i+1}')
    num = textbox(s11, x + 0.15, y + 0.14, 1.0, 0.4, [para([(f'0{i+1}', 20, BLUE, True)])], name=f'Card number {i+1}')
    tt = textbox(s11, x + 0.15, y + 0.55, cw - 0.3, 0.3, [para([(ttl, 12, INK, True)])], name=f'Card title {i+1}')
    bb = textbox(s11, x + 0.15, y + 0.88, cw - 0.3, 1.40, [para([(body, 11, GREY)], ln_spc=1.05)], name=f'Card body {i+1}')
    tg = textbox(s11, x + 0.15, y + ch - 0.42, cw - 0.3, 0.36, [para([(tag, 9, BLUE, True)])], name=f'Card tag {i+1}')
    s11_cards.append((card, num, tt, bb, tg))
s11_note = textbox(s11, 0.62, 4.42, 8.7, 0.55, [para([('Elsewhere (scheduled supply in Jordan, RV tanks) the same wait draws on a finite reserve — context, not a finding about our users.', 11, GREY, False, True)])], name='Global footnote')
s11_foot = footer(s11, 'Background: EPA 2026 (hot-water delivery). Tags show where each constraint comes from.')
s11_num = page_number(s11, 11)
set_transition(s11)
t = Timing().add((s11_lab, 'fade'), (s11_hd, 'fade'), auto=True)
for ids in s11_cards:
    t.add(*[(i, 'wipeLeft') for i in ids])
t.add((s11_note, 'fade'))
t.apply(s11)
note(11, "Why does this persist? Habit: you turn it on and wait; the run-off is never handled. Cost signal: nobody can see the loss, so nothing prompts a change. "
         "Systems: the water is heated away from the shower, and re-plumbing is out of most residents' hands. "
         "Climate and routine are plausible constraints we still have to verify; we are honest about that on the slide.",
     "headline on entry. Click 1–5: one card per click. Click 6: the context footnote.",
     "I1 00:45, 00:34, 00:56; I2 00:40, 00:30, 00:52.")
CLAIMS += [(11, 'Habit: turn on and wait, run-off never handled', 1, '00:56', 'No, it just goes down the drain.'),
           (11, 'Cost signal: cannot estimate the loss', 1, '00:34', "No idea… I would say at least a bucket.")]

# ============================================================================= SLIDE 12 - WHAT USERS DO TODAY (new, table-style rows)
s12 = prs.slides.add_slide(layout)
set_background(s12, LIGHT_BG)
s12_lab = label(s12, 'WHAT USERS DO TODAY — AND WHY IT DOES NOT WORK')
s12_hd = textbox(s12, 0.59, 0.78, 8.9, 0.5, [para([('Today’s options all leave the water running', 26, INK, True)])], name='Headline')
ROWS = [
    ('Turn it on and walk away', 'Interviewee 1: “I’m waiting outside.”', 'Nobody is in the room when warm water arrives, so the run continues until they return.'),
    ('Stand by and wait', 'Interviewee 2: “just sit there.”', 'Passive minutes with the water in view — and it still goes down the drain.'),
    ('Collect it in a bucket', 'Neither does it: “it just goes down the drain.”', 'Extra handling and storage in a rushed moment; one user is not even in the room.'),
    ('Thermostatic shut-off valve', 'Published approach (EPA 2026)', 'Stops the flow once warm water arrives; the cold slug still drains first, and it needs installing.'),
    ('Recirculation loop', 'Published approach (EPA 2026)', 'Keeps the pipe warm, but it is a plumbing change — out of reach for tenants; ongoing energy cost.'),
]
x1, x2, w1, w2 = 0.62, 4.05, 3.25, 5.35
y = 1.42
hdr1 = textbox(s12, x1, y, w1, 0.32, [para([('What users do or could do', 12, BLUE, True)])], anchor='ctr', name='Header 1')
hdr2 = textbox(s12, x2, y, w2, 0.32, [para([('Why it doesn’t work for these users', 12, BLUE, True)])], anchor='ctr', name='Header 2')
hdr_line = line(s12, x1, y + 0.34, x2 + w2, y + 0.34, RULE, 7150, name='Header rule')
y += 0.40
s12_rows = []
rh = 0.60
for i, (what, src, why) in enumerate(ROWS):
    a = textbox(s12, x1, y, w1, rh, [para([(what, 13, INK, True)], spc_aft=1), para([(src, 10, GREY, False, True)])], anchor='ctr', name=f'Row {i+1} what')
    b = textbox(s12, x2, y, w2, rh, [para([(why, 12, GREY)], ln_spc=1.02)], anchor='ctr', name=f'Row {i+1} why')
    rl = line(s12, x1, y + rh + 0.02, x2 + w2, y + rh + 0.02, RULE, 7150, name=f'Row {i+1} rule')
    s12_rows.append((a, b, rl))
    y += rh + 0.06
s12_foot = footer(s12, 'Rows 1–3 from interviews 1–2. Rows 4–5 are published approaches (EPA 2026), listed as context, not as proposals.')
s12_num = page_number(s12, 12)
set_transition(s12)
t = Timing().add((s12_lab, 'fade'), (s12_hd, 'fade'), (hdr1, 'fade'), (hdr2, 'fade'), (hdr_line, 'fade'), auto=True)
for a, b, rl in s12_rows:
    t.add((a, 'wipeLeft'), (b, 'wipeLeft'), (rl, 'wipeLeft'))
t.apply(s12)
note(12, "What do users do today? Mr Ali turns it on and walks away; Ryan stands by and waits. Neither collects the water; a bucket means handling, and one of them is not even in the room. "
         "The published options do not fit these users either: a shut-off valve still drains the cold slug first and needs installing, and a recirculation loop is a plumbing change most tenants cannot make.",
     "headers on entry. Click 1–5: one row per click.",
     "I1 00:45, 00:56; I2 00:40, 00:52. EPA 2026 for rows 4–5.")
CLAIMS += [(12, 'Turn it on and walk away', 1, '00:45', "I'm waiting outside."),
           (12, 'Stand by and wait', 2, '00:40', 'I just let the water run until it gets warm and just sit there.'),
           (12, 'Nobody collects the water', 1, '00:56', 'No, it just goes down the drain.')]

# ============================================================================= SLIDE 13 - REFLECTION 01: ASSUMPTION (dark hero)
s13 = prs.slides.add_slide(layout)
set_background(s13, DARK_BG)
picture(s13, IMG_SHOWER_DRAIN, 0, 0, 10, 5.625, crop=(0.0, 0.0, 0.0, 0.44), name='Photo: shower')
s13_ov = shape(s13, 'rect', 0, 0, 10, 5.625, fill=DARK_BG, alpha=62, name='Photo overlay')
s13_lab = label(s13, 'REFLECTION — 01 / WHAT WE ASSUMED', dark=True)
s13_hd = headline(s13, ['We assumed a', '30-second wait'], y=1.20, sz=41, dark=True, accent_from=1, w=6)
s13_body = textbox(s13, 0.62, 3.15, 6.0, 1.3,
                   [para([('…and a few litres lost: 4.35 L per shower.', 18, WHITE)], spc_aft=6),
                    para([('A minor annoyance, not a routine.', 18, WHITE)])], name='Assumption body')
s13_foot = footer(s13, 'Our starting assumption, before any interview (see slide 5).', dark=True)
s13_num = page_number(s13, 13, dark=True)
set_transition(s13)
Timing().add((s13_lab, 'fade'), (s13_hd, 'fade'), auto=True).add((s13_body, 'fade')).apply(s13)
note(13, "Now, what we learned about our own thinking. We started with one assumption: the wait is about thirty seconds, so the loss is a few litres. "
         "We treated it as a minor annoyance rather than a routine.",
     "headline on entry. Click 1: the two assumption lines.")

# ============================================================================= SLIDE 14 - REFLECTION 02: EVIDENCE THAT CHANGED OUR MIND (dark)
s14 = prs.slides.add_slide(layout)
set_background(s14, DARK_BG)
s14_lab = label(s14, 'REFLECTION — 02 / THE EVIDENCE THAT CHANGED IT', dark=True)
EV = [('“About 3 to 4 minutes.”', 'Interviewee 1 · 00:22'),
      ('“Three to five minutes.”', 'Interviewee 2 · 00:16'),
      ('“No idea… I would say at least a bucket.”', 'Interviewee 1 · 00:34'),
      ('“I’m waiting outside.”  /  “…and just sit there.”', 'Interviewee 1 · 00:45  /  Interviewee 2 · 00:40')]
s14_q = []
for i, (q, who) in enumerate(EV):
    y = 0.98 + i * 0.98
    qt = textbox(s14, 0.62, y, 5.6, 0.5, [para([(q, 20, WHITE, True)])], name=f'Evidence quote {i+1}')
    at = textbox(s14, 0.62, y + 0.48, 5.6, 0.3, [para([('— ' + who, 11, MIST)])], name=f'Evidence attribution {i+1}')
    s14_q.append((qt, at))
s14_big = textbox(s14, 6.45, 1.35, 3.2, 1.3, [para([('6–10×', 66, SKY, True)])], name='Big number ratio')
s14_bigl = textbox(s14, 6.5, 2.75, 3.0, 0.9, [para([('longer than we assumed', 17, WHITE)]), para([('3–5 min reported vs 30 s', 14, MIST)])], name='Ratio label')
s14_foot = footer(s14, 'Both interviews, 27 Sept 2026. Full transcripts and timestamps in evidence.md.', dark=True)
s14_num = page_number(s14, 14, dark=True)
set_transition(s14)
t = Timing().add((s14_lab, 'fade'), auto=True)
for qt, at in s14_q:
    t.add((qt, 'fade'), (at, 'fade'))
t.add((s14_big, 'flyUpFade'), (s14_bigl, 'fade'))
t.apply(s14)
note(14, "Four lines changed our mind. Three to four minutes. Three to five minutes. No idea, at least a bucket. And two different ways of waiting: outside the room, or sitting there. "
         "The reported wait is six to ten times longer than we assumed, and the loss is invisible to the people living it.",
     "label on entry. Click 1–4: one quote per click. Click 5: the '6–10×' number.",
     "I1 00:22, 00:34, 00:45; I2 00:16, 00:40.")
CLAIMS += [(14, '6–10× longer than the 30 s assumption', 2, '00:16', 'Three to five minutes.')]

# ============================================================================= SLIDE 15 - REFLECTION 03: HOW THE POV SHARPENED (dark)
s15 = prs.slides.add_slide(layout)
set_background(s15, DARK_BG)
s15_lab = label(s15, 'REFLECTION — 03 / HOW THE POV SHARPENED', dark=True)
s15_b_lab = textbox(s15, 0.62, 1.05, 4.0, 0.3, [para([('BEFORE', 11, MIST, True)])], name='Before label')
s15_b = textbox(s15, 0.62, 1.40, 4.0, 1.6, [para([('“Residents waste water while waiting for hot water.”', 22, WHITE, True)], ln_spc=1.05)], name='Before statement')
s15_b_gaps = textbox(s15, 0.62, 3.15, 4.0, 1.5,
                     [para([('Any resident, anywhere', 14, MIST)], spc_aft=4), para([('No measure of the wait', 14, MIST)], spc_aft=4),
                      para([('No reason it persists', 14, MIST)], spc_aft=4), para([('Nothing a team could act on', 14, MIST)])], name='Before gaps')
s15_arrow = shape(s15, 'chevron', 4.85, 2.45, 0.3, 0.5, fill=SKY, name='Arrow')
s15_a_lab = textbox(s15, 5.45, 1.05, 4.0, 0.3, [para([('AFTER', 11, SKY, True)])], name='After label')
s15_a = textbox(s15, 5.45, 1.40, 4.0, 2.9,
                [para([(txt, 15, SKY if col == BLUE else WHITE, True) for txt, col in POV_PARTS], ln_spc=1.08)], name='After statement')
s15_a_gaps = textbox(s15, 5.45, 4.35, 4.0, 0.7,
                     [para([('Specific user · measured wait · observed behaviour · a reason it persists', 12, MIST)])], name='After gains')
s15_foot = footer(s15, 'Before: our v1 framing. After: grounded in interviews 1–2 (27 Sept 2026).', dark=True)
s15_num = page_number(s15, 15, dark=True)
set_transition(s15)
Timing().add((s15_lab, 'fade'), (s15_b_lab, 'fade'), (s15_b, 'fade'), auto=True).add((s15_b_gaps, 'wipeLeft')).add((s15_arrow, 'fade'), (s15_a_lab, 'fade'), (s15_a, 'fade')).add((s15_a_gaps, 'fade')).apply(s15)
note(15, "Here is how the point of view sharpened. Before: residents waste water while waiting for hot water. Any resident, no measure, no cause. "
         "After: adults in UAE homes who shower with warm water need a way to stop the three-to-five-minute warm-up from sending clean water down the drain, because the wait is passive time and nobody sees how much is lost. "
         "Same moment, but now with a user, a measure, a behaviour and a reason.",
     "'Before' statement on entry. Click 1: its gaps. Click 2: the 'After' statement. Click 3: what it gained.")

# ============================================================================= SLIDE 16 - CLOSING (dark)
s16 = prs.slides.add_slide(layout)
set_background(s16, DARK_BG)
picture(s16, IMG_DUSK, 0, 0, 10, 5.625, crop=(0.0, 0.0, 0.42, 0.0), name='Photo: dusk')
s16_ov = shape(s16, 'rect', 0, 0, 10, 5.625, fill=DARK_BG, alpha=70, name='Photo overlay')
s16_lab = label(s16, 'FIRST DROP  ·  IEN301 PROJECT 1', dark=True)
s16_hd = textbox(s16, 0.61, 1.25, 8.0, 1.1, [para([('Questions?', 54, WHITE, True)])], name='Questions')
s16_pov = textbox(s16, 0.62, 2.55, 8.6, 1.5,
                  [para([('OUR POINT OF VIEW', 11, SKY, True)], spc_aft=6),
                   para([(txt, 14, SKY if col == BLUE else WHITE, col == BLUE) for txt, col in POV_PARTS], ln_spc=1.1)], name='POV small')
s16_foot = footer(s16, '  ·  '.join(TEAM) + '   ·   Transcripts on request', dark=True)
s16_num = page_number(s16, 16, dark=True)
set_transition(s16)
Timing().add((s16_lab, 'fade'), (s16_hd, 'fade'), auto=True).add((s16_pov, 'fade')).apply(s16)
note(16, "That is our problem framing. We leave the point of view on screen and welcome your questions. "
         "We have both recordings and the transcripts with timestamps if you want to check any quote.",
     "'Questions?' on entry. Click 1: the POV in small text.")

# ============================================================================= SLIDE 17 - APPENDIX (hidden): claim -> recording map (real table)
s17 = prs.slides.add_slide(layout)
set_background(s17, LIGHT_BG)
s17._element.set('show', '0')
s17_lab = label(s17, 'APPENDIX (HIDDEN) — CLAIM TO RECORDING MAP')
s17_hd = textbox(s17, 0.59, 0.72, 8.9, 0.4, [para([('Every on-slide claim, and where to hear it', 20, INK, True)])], name='Headline')
# group claims by recording moment (interview, timestamp): one row per moment, listing the slides that use it
from collections import OrderedDict
groups = OrderedDict()
for sl, claim, iv, ts, verb in CLAIMS:
    key = (iv, ts)
    g = groups.setdefault(key, {'slides': [], 'claims': [], 'verb': verb})
    if sl not in g['slides']:
        g['slides'].append(sl)
    short = claim.split(':')[-1].strip()
    if short not in g['claims']:
        g['claims'].append(short)
rows = []
for (iv, ts), g in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
    rows.append((', '.join(str(x) for x in g['slides']), '; '.join(g['claims'][:2]), f'Interview {iv}', ts, g['verb'].replace(' [unclear]', '')))
rows.append(('3, 6, 8, 11', 'Winter, home type, heating setup, tenure, routine — NOT covered', '—', '—', 'Placeholders on slides; see TODO.md'))
rows.append(('—', 'Interviewer-supplied figures (6–12 L/min; 30,000 L/yr) — NOT used', 'Interviews 1, 2', '01:03; 00:58', 'Said by the interviewer, unsourced'))
from pptx.util import Emu as _Emu
nrows, ncols = len(rows) + 1, 5
gf = s17.shapes.add_table(nrows, ncols, emu(0.62), emu(1.2), emu(8.7), emu(0.25 * nrows))
gf.name = 'Claim map table'
tbl = gf.table
widths = [0.75, 3.1, 0.95, 0.75, 3.15]
for i, w in enumerate(widths):
    tbl.columns[i].width = emu(w)
HEAD = ['Slides', 'Claim(s) on slide', 'Source', 'Time', 'Verbatim']
def style_cell(cell, text, bold, color, sz):
    tf = cell.text_frame
    tf.text = ''
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = text
    r.font.size = Pt(sz); r.font.bold = bold; r.font.name = FONT
    from pptx.dml.color import RGBColor
    r.font.color.rgb = RGBColor.from_string(color)
    cell.margin_left = emu(0.0); cell.margin_right = emu(0.08); cell.margin_top = emu(0.03); cell.margin_bottom = emu(0.03)
    # fill + borders like slide 10 (transparent sides, light rule at the bottom)
    tcPr = cell._tc.get_or_add_tcPr()
    for tag in ('a:lnL', 'a:lnR', 'a:lnT', 'a:lnB'):
        ln = E(tag, {'w': '7150', 'cap': 'flat', 'cmpd': 'sng'})
        if tag == 'a:lnB':
            ln.append(E('a:solidFill', None, E('a:srgbClr', {'val': RULE})))
        else:
            ln.append(E('a:noFill'))
        tcPr.append(ln)
    tcPr.append(E('a:solidFill', None, E('a:srgbClr', {'val': LIGHT_BG})))
for c, h in enumerate(HEAD):
    style_cell(tbl.cell(0, c), h, True, BLUE, 10)
for r, row in enumerate(rows, start=1):
    for c, val in enumerate(row):
        style_cell(tbl.cell(r, c), val, c == 1, INK if c in (0, 1) else GREY, 9)
for r in range(nrows):
    tbl.rows[r].height = emu(0.24)
# drop the table style banding flags python-pptx sets, keep the original deck's table style id
tblPr = tbl._tbl.tblPr
for k in ('firstRow', 'bandRow'):
    if tblPr.get(k) is not None:
        del tblPr.attrib[k]
sid_el = tblPr.find('{%s}tableStyleId' % A)
if sid_el is None:
    sid_el = etree.SubElement(tblPr, '{%s}tableStyleId' % A)
sid_el.text = '{609977B7-E36E-42CD-BEEE-EC8A0416F217}'
s17_foot = footer(s17, 'Hidden appendix for Q&A. Timestamps are mm:ss in each video. Full transcripts: interview_1.md, interview_2.md; evidence.md.')
s17_num = page_number(s17, 17)
set_transition(s17)
note(17, "Hidden appendix. Not presented; used during Q&A to point to the exact moment in each recording.", "no builds.")

# ============================================================================= ORDER + DELETE
final = [s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16, s17]
keep_parts = {s.part for s in final}
sldIdLst = prs.slides._sldIdLst
# delete dropped slides
for sldId in list(sldIdLst):
    part = prs.part.related_part(sldId.rId)
    if part not in keep_parts:
        prs.part.drop_rel(sldId.rId)
        sldIdLst.remove(sldId)
# reorder
id_by_part = {prs.part.related_part(sldId.rId): sldId for sldId in list(sldIdLst)}
for sldId in list(sldIdLst):
    sldIdLst.remove(sldId)
for s in final:
    sldIdLst.append(id_by_part[s.part])

# ============================================================================= NOTES (fresh for every slide)
for n, s in enumerate(final, start=1):
    set_notes(s, NOTES[n])

prs.save(OUT)
print('saved', OUT)

# ============================================================================= REPORT
def spoken_words(txt):
    body = txt.split('\n\nClick:')[0]
    return len(re.findall(r"[A-Za-z0-9’'\-–]+", body))
total = 0
print('\nSpeaker-notes word counts (spoken part only):')
for n in range(1, 18):
    w = spoken_words(NOTES[n]); total += w
    print(f'  slide {n:2d}: {w:3d} words')
print(f'  TOTAL: {total} words  ->  {total/130:.1f} min at 130 wpm')
json.dump({'notes': NOTES, 'claims': CLAIMS, 'numbers': NUMBERS, 'total_words': total}, open('build_report.json', 'w'), indent=1, ensure_ascii=False)
