#!/usr/bin/env python3
"""Generate the Quiz Night deck (16 June, Permanent Representation of
Lithuania to the EU) from the questions drafted in `Klausimai protmūšiui.docx`.

Format: one slide per question, followed by one slide with the answer.
Illustrations are OpenMoji (CC BY-SA 4.0) rendered to `assets/` by
`fetch_assets.py`; answer-slide photos live in `photos/` (see
photos/CREDITS.md). Production notes and source links live in speaker notes.

Usage:  pip install python-pptx Pillow && python3 generate_slides.py
Output: Quiz_Night_2026-06-16.pptx
"""

from PIL import Image
from lxml import etree
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette --
BG = RGBColor(0x0B, 0x0E, 0x15)         # near-black background (every slide)
PANEL = RGBColor(0x17, 0x1C, 0x27)      # option cards / fact panels
PAPER = RGBColor(0xFF, 0xFF, 0xFF)      # headline text
LIGHT = RGBColor(0xDD, 0xE2, 0xEB)      # body text on dark
FOG = RGBColor(0x95, 0x9E, 0xAF)        # secondary text on dark
INK = RGBColor(0x10, 0x13, 0x18)        # text on bright slabs
CIRCLE = RGBColor(0xF7, 0xF3, 0xE9)     # icon circles


def mix(a, b, t):
    """Blend colour a toward colour b by factor t (0..1)."""
    return RGBColor(*(round(a[i] + (b[i] - a[i]) * t) for i in range(3)))


# One saturated colour per round: EU blue, then the Lithuanian tricolour.
# slab = full-bleed divider background · bright = accent on dark slides ·
# on = text colour that reads on the slab.
THEMES = {
    "R1": {"slab": RGBColor(0x0A, 0x32, 0x8C),
           "bright": RGBColor(0x7B, 0xA7, 0xFF), "on": PAPER},
    "R2": {"slab": RGBColor(0xFD, 0xB9, 0x13),
           "bright": RGBColor(0xFF, 0xC9, 0x2E), "on": INK},
    "R3": {"slab": RGBColor(0x00, 0x6A, 0x44),
           "bright": RGBColor(0x46, 0xD2, 0x90), "on": PAPER},
    "RB": {"slab": RGBColor(0xC1, 0x27, 0x2D),
           "bright": RGBColor(0xFF, 0x70, 0x61), "on": PAPER},
}
TRICOLOUR = [THEMES["R2"]["slab"], THEMES["R3"]["slab"], THEMES["RB"]["slab"]]
CURRENT = THEMES["R2"]  # active round theme; dividers switch it

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
FONT = "Calibri"
DISPLAY = "Arial Black"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]

FOOTER = "Quiz Night · 16 June 2026 · Permanent Representation of Lithuania to the EU"


# ---------------------------------------------------------------- helpers --
def add_slide(bg=None, notes=None, footer=None):
    slide = prs.slides.add_slide(BLANK)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = bg or BG
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    text(slide, Inches(0.7), Inches(7.08), Inches(10), Inches(0.3), FOOTER,
         size=9.5, color=footer or mix(FOG, BG, 0.35))
    return slide


def shape(slide, kind, x, y, w, h, fill=None, radius=None):
    sp = slide.shapes.add_shape(kind, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    sp.line.fill.background()
    if radius is not None:
        sp.adjustments[0] = radius
    return sp


def text(slide, x, y, w, h, content, size=16, color=LIGHT, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=None,
         space_after=None, font=FONT):
    """content: str | list of paragraphs; a paragraph is str | list of
    (text, fmt) run tuples with fmt keys bold/color/size/italic."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(content, str):
        content = [content]
    for i, para in enumerate(content):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if spacing:
            p.line_spacing = spacing
        if space_after is not None:
            p.space_after = space_after
        if isinstance(para, str):
            para = [(para, {})]
        for run_text, fmt in para:
            r = p.add_run()
            r.text = run_text
            r.font.name = font
            r.font.size = Pt(fmt.get("size", size))
            r.font.bold = fmt.get("bold", bold)
            r.font.italic = fmt.get("italic", False)
            r.font.color.rgb = fmt.get("color", color)
    return tb


def icon_circle(slide, name, cx, cy, d):
    """Icon in a soft circle, centred at (cx, cy) inches, diameter d."""
    cx, cy, d = Inches(cx), Inches(cy), Inches(d)
    shape(slide, MSO_SHAPE.OVAL, cx - d / 2, cy - d / 2, d, d, fill=CIRCLE)
    pic = Emu(int(d * 0.72))
    slide.shapes.add_picture(f"assets/{name}.png", cx - pic / 2, cy - pic / 2,
                             pic, pic)


def giant(slide, txt, color, x, y, w, h, size, align=PP_ALIGN.RIGHT,
          anchor=MSO_ANCHOR.TOP):
    """Oversized display glyph (question numbers, divider numerals)."""
    text(slide, Inches(x), Inches(y), Inches(w), Inches(h), txt, size=size,
         color=color, align=align, anchor=anchor, font=DISPLAY)


def edge_bar(slide):
    """Full-height accent bar on the left edge of content slides."""
    shape(slide, MSO_SHAPE.RECTANGLE, 0, 0, Inches(0.14), SLIDE_H,
          fill=CURRENT["bright"])


def chip(slide, letter, x, y, side=0.5):
    """Bright square chip with the option letter."""
    sp = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
               Inches(side), Inches(side), fill=CURRENT["bright"], radius=0.3)
    tf = sp.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = letter
    r.font.name = DISPLAY
    r.font.size = Pt(15)
    r.font.color.rgb = BG


def tricolour_bars(slide, cy, w=1.05, h=0.14, gap=0.2):
    """Lithuanian tricolour, centred horizontally at height cy (inches)."""
    x = (13.333 - (3 * w + 2 * gap)) / 2
    for c in TRICOLOUR:
        shape(slide, MSO_SHAPE.RECTANGLE, Inches(x), Inches(cy), Inches(w),
              Inches(h), fill=c)
        x += w + gap


def photo_frame(slide, name, x, y, w, h):
    """Photo from photos/, cropped to fill the box, with rounded corners."""
    path = f"photos/{name}"
    pic = slide.shapes.add_picture(path, Inches(x), Inches(y), Inches(w),
                                   Inches(h))
    src_w, src_h = Image.open(path).size
    target, source = w / h, src_w / src_h
    if source > target:
        f = 1 - target / source
        pic.crop_left = pic.crop_right = f / 2
    elif source < target:
        f = 1 - source / target
        pic.crop_top = pic.crop_bottom = f / 2
    spPr = pic._element.spPr
    ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for geom in spPr.findall(f"{{{ns}}}prstGeom"):
        spPr.remove(geom)
    geom = etree.SubElement(spPr, f"{{{ns}}}prstGeom")
    geom.set("prst", "roundRect")
    av = etree.SubElement(geom, f"{{{ns}}}avLst")
    gd = etree.SubElement(av, f"{{{ns}}}gd")
    gd.set("name", "adj")
    gd.set("fmla", "val 6000")
    pic.line.color.rgb = CURRENT["bright"]
    pic.line.width = Pt(1.5)
    pic.shadow.inherit = False
    return pic


def icon_row(slide, names, cy, d=1.5, gap=0.55):
    span = len(names) * d + (len(names) - 1) * gap
    x = (13.333 - span) / 2 + d / 2
    for n in names:
        icon_circle(slide, n, x, cy, d)
        x += d + gap


def kicker(slide, left, right=None):
    text(slide, Inches(0.7), Inches(0.55), Inches(8.5), Inches(0.35), left,
         size=13, bold=True, color=CURRENT["bright"])
    if right:
        text(slide, Inches(9.2), Inches(0.55), Inches(3.43), Inches(0.35),
             right, size=13, bold=True, color=FOG, align=PP_ALIGN.RIGHT)


def video_pill(slide, label, url, y):
    pill = shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(y),
                 Inches(2.55), Inches(0.46), fill=CURRENT["bright"],
                 radius=0.5)
    tf = pill.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = "▶   " + label
    r.font.name = FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = BG
    text(slide, Inches(3.45), Inches(y + 0.11), Inches(5.5), Inches(0.3), url,
         size=12, color=FOG)


def fact_panel(slide, fact, x, y, w, h):
    shape(slide, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
          Inches(w), Inches(h), fill=PANEL, radius=0.07)
    shape(slide, MSO_SHAPE.RECTANGLE, Inches(x), Inches(y + 0.14),
          Inches(0.07), Inches(h - 0.28), fill=CURRENT["bright"])
    text(slide, Inches(x + 0.4), Inches(y + 0.3), Inches(w - 0.75),
         Inches(0.3), "DID YOU KNOW?", size=12.5, bold=True,
         color=CURRENT["bright"])
    text(slide, Inches(x + 0.4), Inches(y + 0.72), Inches(w - 0.75),
         Inches(h - 1.0), fact, size=16, color=LIGHT, spacing=1.15)


# ----------------------------------------------------------------- slides --
def title_slide():
    s = add_slide(notes=(
        "Host notes: welcome everyone and explain the team draw — slips with "
        "names of Lithuanian rivers and lakes at the door."))
    dim = mix(THEMES["R2"]["bright"], BG, 0.84)
    giant(s, "?", dim, 10.4, 0.0, 2.6, 3.4, size=230)
    giant(s, "?", dim, 0.35, 4.1, 2.6, 3.4, size=230, align=PP_ALIGN.LEFT)
    text(s, Inches(0.8), Inches(2.2), Inches(11.73), Inches(0.45),
         "BRUSSELS · 16 JUNE 2026", size=16, bold=True,
         color=THEMES["R2"]["bright"], align=PP_ALIGN.CENTER)
    text(s, Inches(0.4), Inches(2.7), Inches(12.53), Inches(1.65),
         "QUIZ NIGHT", size=92, color=PAPER, align=PP_ALIGN.CENTER,
         font=DISPLAY)
    tricolour_bars(s, 4.6)
    text(s, Inches(1.5), Inches(5.05), Inches(10.33), Inches(0.45),
         "Permanent Representation of Lithuania to the European Union",
         size=16, color=FOG, align=PP_ALIGN.CENTER)


def rules_slide():
    s = add_slide(notes=(
        "Team assignment: everyone draws a slip when entering — the slips "
        "carry names of Lithuanian rivers and lakes."))
    text(s, Inches(0.7), Inches(0.6), Inches(12), Inches(0.85),
         "HOW TONIGHT WORKS", size=38, color=PAPER, font=DISPLAY)
    cards = [
        ("teams", "TEAMS", "R1",
         "Draw a slip at the door — the river or lake on it is your team."),
        ("target", "ROUNDS", "R2",
         "Three rounds plus a bonus round. One point per question unless "
         "marked otherwise."),
        ("memo", "ANSWERS", "R3",
         "Write your answers down — we reveal and score after each round."),
        ("no_phone", "FAIR PLAY", "RB",
         "No phones. Diplomatic immunity does not cover googling."),
    ]
    grid = [(0.7, 1.9), (6.95, 1.9), (0.7, 4.35), (6.95, 4.35)]
    for (gx, gy), (icon, head, theme, body) in zip(grid, cards):
        accent = THEMES[theme]["bright"]
        shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy),
              Inches(5.7), Inches(2.05), fill=PANEL, radius=0.09)
        shape(s, MSO_SHAPE.RECTANGLE, Inches(gx), Inches(gy + 0.16),
              Inches(0.07), Inches(1.73), fill=accent)
        icon_circle(s, icon, gx + 1.0, gy + 1.02, 1.15)
        text(s, Inches(gx + 1.85), Inches(gy + 0.38), Inches(3.65),
             Inches(0.42), head, size=17, bold=True, color=accent)
        text(s, Inches(gx + 1.85), Inches(gy + 0.85), Inches(3.65),
             Inches(1.05), body, size=14.5, color=LIGHT, spacing=1.1)


def divider(theme, round_no, title, sub, icons, num, notes=None, size=54):
    global CURRENT
    CURRENT = THEMES[theme]
    slab, on = CURRENT["slab"], CURRENT["on"]
    s = add_slide(bg=slab, notes=notes, footer=mix(on, slab, 0.5))
    giant(s, num, mix(on, slab, 0.85), 6.8, 2.0, 5.83, 5.5, size=300,
          anchor=MSO_ANCHOR.BOTTOM)
    text(s, Inches(0.9), Inches(1.45), Inches(10), Inches(0.45), round_no,
         size=17, bold=True, color=mix(on, slab, 0.2))
    text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(2.3),
         title.upper(), size=size, color=on, font=DISPLAY, spacing=1.02)
    text(s, Inches(0.9), Inches(4.5), Inches(10), Inches(0.45), sub,
         size=17, bold=True, color=mix(on, slab, 0.25))
    x = 0.9 + 0.55
    for n in icons:
        icon_circle(s, n, x, 6.0, 1.1)
        x += 1.1 + 0.4


def question_slide(round_label, q_no, q_total, question, icon, options=None,
                   hint=None, video=None, notes=None, points="1 point",
                   q_size=26):
    s = add_slide(notes=notes)
    edge_bar(s)
    kicker(s, round_label, f"QUESTION {q_no} OF {q_total} · {points.upper()}")
    bright = CURRENT["bright"]
    dim = mix(bright, BG, 0.8)
    if options:
        giant(s, str(q_no), dim, 10.6, 0.45, 2.33, 2.1, size=140)
        text(s, Inches(0.7), Inches(1.25), Inches(9.7), Inches(2.4),
             question, size=q_size, bold=True, color=PAPER, spacing=1.12)
        full_row = any(len(o[1]) > 38 for o in options)
        if full_row:
            y = 4.0 if len(options) == 3 else 3.7
            for letter, opt in options:
                shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), Inches(y),
                      Inches(11.93), Inches(0.78), fill=PANEL, radius=0.16)
                chip(s, letter, 0.94, y + 0.16, side=0.46)
                text(s, Inches(1.62), Inches(y), Inches(10.7), Inches(0.78),
                     opt, size=17, bold=True, color=LIGHT,
                     anchor=MSO_ANCHOR.MIDDLE)
                y += 0.94
        else:
            gy0 = 3.55 if len(question) < 130 else 4.15
            grid = [(0.7, gy0), (6.78, gy0), (0.7, gy0 + 1.2),
                    (6.78, gy0 + 1.2)]
            for (gx, gy), (letter, opt) in zip(grid, options):
                shape(s, MSO_SHAPE.ROUNDED_RECTANGLE, Inches(gx), Inches(gy),
                      Inches(5.85), Inches(0.95), fill=PANEL, radius=0.14)
                chip(s, letter, gx + 0.27, gy + 0.225)
                text(s, Inches(gx + 1.02), Inches(gy), Inches(4.6),
                     Inches(0.95), opt, size=18, bold=True, color=LIGHT,
                     anchor=MSO_ANCHOR.MIDDLE)
    else:
        giant(s, str(q_no), dim, 10.6, 0.45, 2.33, 1.9, size=120)
        text(s, Inches(0.7), Inches(1.55), Inches(7.9), Inches(4.4),
             question, size=q_size, bold=True, color=PAPER, spacing=1.15)
        icon_circle(s, icon, 10.85, 4.1, 2.7)
        if video:
            video_pill(s, "Play the clip", video, 5.95)
    if hint:
        text(s, Inches(0.7), Inches(6.45), Inches(11.9), Inches(0.4),
             [[("HINT   ", {"bold": True, "color": bright, "size": 14}),
               (hint, {"size": 14, "color": FOG, "italic": True})]])


def answer_slide(round_label, q_no, answer, icon=None, icons=None, fact=None,
                 video=None, notes=None, a_size=38, photo=None, credit=None,
                 photo_h=None):
    s = add_slide(notes=notes)
    edge_bar(s)
    kicker(s, round_label, f"ANSWER · QUESTION {q_no}")
    text_w = 7.6 if photo else 8.4
    text(s, Inches(0.7), Inches(1.45), Inches(text_w), Inches(1.7), answer,
         size=a_size, bold=True, color=CURRENT["bright"], spacing=1.05)
    if fact:
        fact_panel(s, fact, 0.7, 3.65, 7.6 if photo else 7.9, 2.95)
    if photo:
        names = photo if isinstance(photo, list) else [photo]
        credit_y = 6.44
        if len(names) == 1:
            h = photo_h or 4.9
            y = 1.45 + (4.9 - h) / 2
            photo_frame(s, names[0], 8.75, y, 3.88, h)
            credit_y = y + h + 0.08
        else:
            photo_frame(s, names[0], 8.75, 1.45, 3.88, 2.33)
            photo_frame(s, names[1], 8.75, 4.02, 3.88, 2.33)
        if credit:
            text(s, Inches(8.75), Inches(credit_y), Inches(3.88),
                 Inches(0.3), credit, size=8.5, color=FOG,
                 align=PP_ALIGN.RIGHT)
    elif icons:
        d, gap = 1.28, 0.22
        cy = 3.55
        cx = 10.85 - (len(icons) - 1) * (d + gap) / 2
        for n in icons:
            icon_circle(s, n, cx, cy, d)
            cx += d + gap
    elif icon:
        icon_circle(s, icon, 10.85, 3.55, 3.0)
    if video:
        video_pill(s, "Watch the story", video, 2.85)


def closing_slide():
    s = add_slide(notes="Count the points and announce the winning team.")
    icon_row(s, ["trophy"], 1.75, d=1.7)
    text(s, Inches(0.7), Inches(2.9), Inches(11.93), Inches(2.2),
         ["THANK YOU", "FOR PLAYING"], size=60, color=PAPER,
         align=PP_ALIGN.CENTER, font=DISPLAY, spacing=1.02)
    text(s, Inches(1.5), Inches(5.35), Inches(10.33), Inches(0.5),
         "Time to count the points", size=19, bold=True,
         color=THEMES["R2"]["bright"], align=PP_ALIGN.CENTER)
    tricolour_bars(s, 6.25)


# ------------------------------------------------------------------ deck ---
title_slide()
rules_slide()

# ===== ROUND 1 ==============================================================
R1 = "ROUND 1 · EUROPE & FUN FACTS"
divider("R1", "ROUND 1", "Europe & Fun Facts",
        "Multiple choice · 6 questions · 1 point each",
        ["globe", "euro_note", "jeans"], num="1")

question_slide(
    R1, 1, 6, "Which European island changes sovereignty every six months?",
    "globe",
    options=[("A", "Heligoland"), ("B", "Pheasant Island"), ("C", "Jersey"),
             ("D", "Bornholm")])
answer_slide(
    R1, 1, "B — Pheasant Island",
    photo="pheasant_island.jpg", credit="Photo: Wikimedia Commons",
    fact="Under the 1659 Treaty of the Pyrenees, France and Spain share "
         "this tiny uninhabited island in the Bidasoa river. Administration "
         "alternates between the two countries every six months — the "
         "world's oldest condominium.")

question_slide(
    R1, 2, 6, "Which EU symbol was publicly unveiled about 40 years ago?",
    "star",
    options=[("A", "Euro coins"), ("B", "European flag"),
             ("C", "European anthem"), ("D", "Schengen passport")])
answer_slide(
    R1, 2, "B — The European flag",
    photo="berlaymont_flag_1986.jpg",
    credit="Photo: © European Communities, 1986",
    fact="In May 1986 the European flag was raised for the first time "
         "outside the Berlaymont building — twelve gold stars in a circle, "
         "a symbol of unity, unchanged ever since.")

question_slide(
    R1, 3, 6, "What major international award did the European Union "
              "receive in 2012?",
    "medal",
    options=[("A", "Sakharov Prize"), ("B", "Nobel Peace Prize"),
             ("C", "Charlemagne Prize"), ("D", "Right Livelihood Award")])
answer_slide(
    R1, 3, "B — The Nobel Peace Prize",
    photo="nobel_peace_prize_2012.jpg", photo_h=2.2, credit="Photo: Reuters",
    fact="Awarded to the EU in 2012 for over six decades of advancing "
         "peace, reconciliation, democracy and human rights in Europe.")

question_slide(
    R1, 4, 6, "Which euro banknote was nicknamed “Bin Laden”?",
    "euro_note",
    options=[("A", "€50"), ("B", "€100"), ("C", "€200"), ("D", "€500")],
    notes="Source: https://www.theguardian.com/business/2016/may/04/"
          "500-euro-banknote-could-be-scrapped-crime")
answer_slide(
    R1, 4, "D — The €500 note", photo="euro_500_note.jpg", photo_h=2.6,
    fact="The ECB stopped issuing the €500 note in 2019 over its popularity "
         "with money launderers — and because, like its namesake, everyone "
         "knew what it looked like but almost no one had ever seen one.")

question_slide(
    R1, 5, 6, "In 1984, former Nigerian minister Umaru Dikko was kidnapped "
              "in London. His captors planned to smuggle him out of the UK "
              "in a diplomatic crate, relying on the Vienna Convention rule "
              "that diplomatic bags cannot be opened or detained.\n\nWhat "
              "went wrong?",
    "crate", q_size=21,
    options=[("A", "The crate was improperly labelled as diplomatic "
                   "baggage"),
             ("B", "Dikko did not fit inside the crate"),
             ("C", "The aircraft meant to transport him never arrived")],
    notes="Background video: https://www.youtube.com/watch?v=N83Idy9IOmU")
answer_slide(
    R1, 5, "A — The crate was improperly labelled", a_size=32, icon="crate",
    fact="With no official diplomatic markings, customs officers at "
         "Stansted were entitled to open the crate — and found Dikko "
         "unconscious, accompanied by an anaesthetist.",
    video="youtube.com/watch?v=N83Idy9IOmU",
    notes="Video: https://www.youtube.com/watch?v=N83Idy9IOmU")

question_slide(
    R1, 6, 6, "One of the inventors of modern blue jeans was born in "
              "present-day Latvia. He developed a way to make work trousers "
              "much more durable and partnered with Levi Strauss to patent "
              "the idea.\n\nWhat simple innovation made the trousers "
              "famous?",
    "jeans", q_size=22,
    options=[("A", "A zipper"), ("B", "Metal rivets"),
             ("C", "Waterproof fabric"), ("D", "Belt loops")])
answer_slide(
    R1, 6, "B — Metal rivets", photo="jeans_rivets.jpg",
    fact="Jacob Davis, a tailor born in Riga in 1831, reinforced the stress "
         "points of work trousers with copper rivets. He and Levi Strauss "
         "patented the idea in 1873 — and blue jeans were born.")

# ===== ROUND 2 ==============================================================
R2 = "ROUND 2 · LITHUANIA"
divider("R2", "ROUND 2", "Lithuania", "8 questions · 1 point each",
        ["microphone", "mountain", "soup"], num="2", size=76)

question_slide(
    R2, 1, 8, "Every year Lithuania plans to win Eurovision — although "
              "officially we never have.\n\nIn 2006, however, Lithuania "
              "tried a different strategy: entering a song that simply "
              "declared victory before the contest was even over.\n\nWhat "
              "was the title of the song?",
    "microphone", q_size=22,
    video="youtube.com/watch?v=DBAdOlQPbwg",
    notes="Video: https://www.youtube.com/watch?v=DBAdOlQPbwg — decide "
          "which part of the clip to show. Option: move this question to a "
          "music/visual round.")
answer_slide(
    R2, 1, "“We Are The Winners”",
    photo=["lt_united_eurovision_2006.jpg", "lena_valaitis_1981.jpg"],
    credit="Photos: EBU / eurovision.tv",
    fact="Honourable mention: Lithuanian-German singer Lena Valaitis took "
         "second place at Eurovision 1981 — proof that we keep looking for "
         "that victory by all possible means.",
    notes="Mention Lena Valaitis as a fun fact, not a question.")

question_slide(
    R2, 2, 8, "Two Lithuanian mountaineers carried a symbolic national "
              "object to the summit of Mount Everest and scattered it from "
              "the world's highest peak.\n\nWhat was it?",
    "climber",
    notes="Tell the full story and the exact years out loud; keep the slide "
          "short.")
answer_slide(
    R2, 2, "Amber",
    photo=["vitkauskas_everest_1995.jpg", "baltic_amber.jpg"],
    credit="Top photo: V. Vitkauskas archive",
    fact="Vladas Vitkauskas (1995) and Saulius Vilius (2003) both carried "
         "Baltic amber to the top of the world and scattered it from the "
         "summit.")

question_slide(
    R2, 3, 8, "When Lithuanians learn about the interwar period, one "
              "statistic is proudly repeated in schools: in 1938, Lithuania "
              "ranked second in Europe in butter and bacon production — "
              "despite being one of the poorest countries on the "
              "continent.\n\nWhich country ranked first?",
    "butter", q_size=22,
    hint="Think of a recent Council Presidency.")
answer_slide(
    R2, 3, "Denmark", photo="interwar_lithuania_1930s.jpg",
    fact="A statistic still celebrated in Lithuanian classrooms. Denmark, "
         "for its part, no longer wants to be the EU's bacon factory — as "
         "Politico once put it.",
    notes="Politico: “Denmark no longer wants to be EU bacon factory”.")

question_slide(
    R2, 4, 8, "Across Europe people celebrate Midsummer with bonfires.\n\n"
              "In Lithuania, one tradition involves searching for a "
              "mythical flower that supposedly blooms only on that "
              "night.\n\nWhich flower?",
    "bonfire")
answer_slide(
    R2, 4, "The fern flower",
    photo="fern_fiddleheads.jpg", credit="Photo: Wikimedia Commons",
    fact="According to tradition, the fern blooms only on Midsummer night — "
         "whoever finds it gains happiness and wisdom. Botanists remain "
         "unconvinced.")

question_slide(
    R2, 5, 8, "1918 was a big year for Lithuania — also for "
              "democracy.\n\nLithuanian women gained something that many "
              "women in Western Europe had to wait decades for.\n\nWhat "
              "was it?",
    "parliament",
    notes="Alternative version: show the centenary stamp (with some details "
          "removed) and ask what occasion it marks. Stamp: "
          "https://manoteises.lt/straipsnis/isleidziamas-pasto-zenklas-"
          "moteru-balsavimo-simtmeciui-lietuvoje-pamineti/")
answer_slide(
    R2, 5, "Voting rights",
    photo="suffrage_stamp_2018.jpg",
    credit="Stamp: Lietuvos paštas, 2018 · design J. Dadonas",
    fact="Lithuanian women gained the vote on 2 November 1918, when the "
         "provisional constitution enshrined equal suffrage — ahead of "
         "much of Western Europe. A commemorative stamp marked the "
         "centenary.")

question_slide(
    R2, 6, 8, "A 16th-century Lithuanian nobleman travelled through Egypt "
              "and bought several unusual souvenirs.\n\nDuring a storm at "
              "sea, terrified sailors blamed the bad weather on them and "
              "threw them overboard.\n\nWhat were the souvenirs?",
    "wave", q_size=24)
answer_slide(
    R2, 6, "Egyptian mummies",
    photo="radvila_the_orphan.jpg",
    credit="Anonymous portrait, c. 1590 (public domain)",
    fact="Mikalojus Kristupas Radvila the Orphan brought two mummies back "
         "from his pilgrimage. When storms battered the ship, the crew "
         "blamed the cargo — and overboard they went. His travel diary "
         "became a European bestseller.")

question_slide(
    R2, 7, 8, "In a humorous Lithuanian promotional video from the early "
              "2000s, traditional dishes such as cepelinai are discussed "
              "with the idea that they might one day become popular across "
              "Europe as street food.\n\nThe video reflects a moment when "
              "Lithuania was preparing for a major change. What event was "
              "approaching?",
    "dumpling", q_size=22,
    video="youtube.com/watch?v=YgvHcenDYcU",
    notes="Video: https://www.youtube.com/watch?v=YgvHcenDYcU — add English "
          "subtitles to the clip. Consider whether to keep “early 2000s” in "
          "the wording.")
answer_slide(
    R2, 7, "Joining the European Union", a_size=34,
    photo="eu_enlargement_2004_map.png", credit="Map: Wikimedia Commons",
    fact="Lithuania joined the EU on 1 May 2004, in the largest enlargement "
         "in the Union's history — ten countries at once. The cepelinai "
         "street-food revolution is still pending.")

question_slide(
    R2, 8, 8, "Pink is strongly associated with Lithuania because of a "
              "traditional summer dish.\n\nWhich dish?",
    "blossom",
    options=[("A", "Cepelinai"), ("B", "Šakotis"), ("C", "Šaltibarščiai"),
             ("D", "Kibinai")])
answer_slide(
    R2, 8, "C — Šaltibarščiai", photo="saltibarsciai.jpg",
    fact="The electric-pink cold beet soup is Lithuania's unofficial "
         "summer flag — best served with hot potatoes and a sunny terrace.")

# ===== ROUND 3 ==============================================================
R3 = "ROUND 3 · ENVIRONMENT & NATURE"
divider("R3", "ROUND 3", "Environment & Nature",
        "4 questions · 1 point each",
        ["headphones", "fern", "beaver"], num="3", size=50)

question_slide(
    R3, 1, 4, "In 2013 Metallica became the first band to perform on all "
              "seven continents.\n\nDuring their Antarctic concert, the "
              "audience used 120 of what?",
    "guitar")
answer_slide(
    R3, 1, "Headphones",
    photo="metallica_freeze_em_all.jpg",
    credit="Artwork: © Blackened Recordings",
    fact="To comply with Antarctic environmental rules, the band played "
         "without amplifiers — the audience of about 120 listened through "
         "headphones. The concert was fittingly called “Freeze 'Em All”.")

question_slide(
    R3, 2, 4, "Three famous scientists — Birutė Galdikas, Dian Fossey and "
              "Jane Goodall — devoted their careers to studying these "
              "animals in the wild.\n\nCollectively, they became known as "
              "the “Trimates”.\n\nWhat group of animals are these?",
    "microscope", q_size=24,
    notes="Photo idea: https://www.themarysue.com/the-trimates-three-women-"
          "that-made-science-history/")
answer_slide(
    R3, 2, "Great apes",
    photo="trimates_goodall_fossey_galdikas.jpg",
    fact="Goodall (chimpanzees), Fossey (gorillas) and Galdikas "
         "(orangutans) were recruited by Louis Leakey — hence also "
         "“Leakey's Angels”. Galdikas, of Lithuanian descent, still works "
         "in Borneo. “Orangutan” means “person of the forest”.")

question_slide(
    R3, 3, 4, "Which EU law unexpectedly entered the headlines after the "
              "deaths of four US soldiers, whose vehicle sank in a military "
              "training area near the Lithuanian–Belarusian border?",
    "newspaper")
answer_slide(
    R3, 3, "The Nature Restoration Law", a_size=34,
    photo="cepkeliai_marsh.jpg", credit="Photo: Wikimedia Commons",
    fact="Commentary on the 2025 tragedy near Pabradė linked the swampy "
         "terrain to wetlands restored under EU environmental policy — "
         "putting the Nature Restoration Law unexpectedly in the news.")

question_slide(
    R3, 4, 4, "This animal is often called an “ecosystem engineer” because "
              "it creates wetlands that benefit countless other "
              "species.\n\nWhat animal is it?",
    "tools")
answer_slide(
    R3, 4, "The beaver", photo="beaver_at_dam.jpg",
    fact="Beaver dams create wetlands that store water, filter pollution "
         "and host countless species. Lithuania's beaver population has "
         "grown from near extinction to one of the densest in Europe.")

# ===== BONUS ROUND ==========================================================
RB = "BONUS · THE BRUSSELS BUBBLE"
divider("RB", "BONUS ROUND", "The Brussels Bubble",
        "Insider questions · for those who were in the room",
        ["recycle", "speech", "ship"], num="B", size=50,
        notes="Possible extra questions: the CBAM question about nails "
              "(still needs wording — a funny one); The Schumann Show about "
              "the institutions: https://www.youtube.com/watch?v=D0fBh-0Eiy0")

question_slide(
    RB, 1, 3, "During negotiations on which file did Violeta Dragu "
              "(Romania) say:\n\n“I have never seen Lithuania being so "
              "active during the negotiations”?",
    "speech",
    notes="Can be made multiple choice to make it easier — or harder, by "
          "also asking what the problem was.")
answer_slide(
    RB, 1, "The Waste Framework Directive", a_size=34, icon="recycle")

question_slide(
    RB, 2, 3, "During a Coreper I meeting, the Council was split on a "
              "legislative file, with both sides carefully counting "
              "votes.\n\nThe Deputy Permanent Representative of Lithuania "
              "intervened and said… what?\n\nAnd which file was being "
              "discussed?",
    "abacus", q_size=23, points="2 points",
    notes="Optional hint: give the exact date — the 14 June Coreper I (or "
          "the earlier one).")
answer_slide(
    RB, 2, "“Scrutiny reservation” — on the Nature Restoration Law",
    a_size=30, photo="europa_building_room.jpg",
    fact="A scrutiny reservation lets a member state hold its position "
         "while procedures are completed back home — a small phrase that "
         "can pause a finely balanced vote. Pictured: the Europa building, "
         "where Coreper meets.")

question_slide(
    RB, 3, 3, "1985: Lithuanian, Latvian and Estonian dissidents organised "
              "the Peace and Freedom Cruise aboard the ship Baltic Star — "
              "sailing from Stockholm along the edge of Soviet territorial "
              "waters and ending with a widely reported demonstration in "
              "Helsinki.\n\nThe Soviet Union tried to derail the voyage. "
              "What did it do?",
    "ship", q_size=22,
    notes="Background: per Tomas Venclova, Aleksandras Štromas was the "
          "cruise's spiritus movens; the Baltic Tribunal also took place in "
          "1985. Keep names off the slide. Ship photo to be added later.")
answer_slide(
    RB, 3, "A rumour that a bomb was on board", a_size=32,
    photo="baltic_star_birger_jarl.jpg", photo_h=2.6,
    credit="Photo: Wikimedia Commons",
    fact="The rumour failed: the cruise went ahead, and the Helsinki "
         "demonstration was reported across the Scandinavian and wider "
         "European press — a loud reminder that the Baltic states had not "
         "been forgotten.")

closing_slide()

OUT = "Quiz_Night_2026-06-16.pptx"
prs.save(OUT)
print(f"Saved {OUT} with {len(prs.slides._sldIdLst)} slides")
