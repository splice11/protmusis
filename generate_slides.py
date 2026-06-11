#!/usr/bin/env python3
"""Generate the Protmūšis quiz deck (16 June, Permanent Representation of
Lithuania to the EU) from the questions drafted in `Klausimai protmūšiui.docx`.

Format follows the brief: one slide per question, followed by one slide with
the answer. Media still to be collected is marked with placeholder frames on
the slides; production notes and source links live in the speaker notes.

Usage:  python3 generate_slides.py
Output: Protmusis_2026-06-16.pptx
"""

import math

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette --
NAVY = RGBColor(0x0B, 0x22, 0x40)        # deep diplomatic navy (background)
NAVY_SOFT = RGBColor(0x14, 0x32, 0x5C)   # panel fill
GOLD = RGBColor(0xFF, 0xCC, 0x00)        # EU gold (accents, answers)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
MIST = RGBColor(0xC9, 0xD6, 0xE8)        # muted body text
LT_YELLOW = RGBColor(0xFD, 0xB9, 0x13)   # Lithuanian tricolour
LT_GREEN = RGBColor(0x00, 0x6A, 0x44)
LT_RED = RGBColor(0xC1, 0x27, 0x2D)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

TITLE_FONT = "Georgia"
BODY_FONT = "Calibri"

prs = Presentation()
prs.slide_width = SLIDE_W
prs.slide_height = SLIDE_H
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------- helpers --
def add_slide(notes=None):
    slide = prs.slides.add_slide(BLANK)
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = NAVY
    if notes:
        slide.notes_slide.notes_text_frame.text = notes
    return slide


def box(slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE,
        line_w=None, dash=None):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = line_w or Pt(1)
        if dash:
            sp.line.dash_style = dash
    return sp


def text(slide, x, y, w, h, runs, size=18, color=WHITE, bold=False,
         align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, font=BODY_FONT,
         spacing=None, space_after=None):
    """runs: str, or list of paragraphs; each paragraph is str or
    list of (text, {bold/color/size/italic/font}) run tuples."""
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    if isinstance(runs, str):
        runs = [runs]
    for i, para in enumerate(runs):
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
            r.font.name = fmt.get("font", font)
            r.font.size = Pt(fmt.get("size", size))
            r.font.bold = fmt.get("bold", bold)
            r.font.italic = fmt.get("italic", False)
            r.font.color.rgb = fmt.get("color", color)
    return tb


def tricolour_footer(slide, label):
    third = Emu(int(SLIDE_W) // 3)
    h = Inches(0.07)
    y = SLIDE_H - h
    for i, c in enumerate((LT_YELLOW, LT_GREEN, LT_RED)):
        box(slide, Emu(int(third) * i), y, third, h, fill=c)
    text(slide, Inches(0.55), Inches(6.98), Inches(12.2), Inches(0.35), label,
         size=10.5, color=MIST)


def chip(slide, label, fill=GOLD, color=NAVY, x=Inches(0.55), y=Inches(0.5),
         w=Inches(4.4)):
    c = box(slide, x, y, w, Inches(0.42), fill=fill,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    c.adjustments[0] = 0.5
    tf = c.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name = BODY_FONT
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = color


def media_placeholder(slide, caption, x, y, w, h, kind="PHOTO"):
    fr = box(slide, x, y, w, h, fill=NAVY_SOFT, line=GOLD, line_w=Pt(1.25),
             dash=MSO_LINE.DASH, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
    fr.adjustments[0] = 0.06
    tf = fr.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_left = tf.margin_right = Inches(0.25)
    icon = "🎬" if kind == "VIDEO" else "📷"
    for i, (t, sz, b, col) in enumerate(
            [(f"{icon}  {kind} TO ADD", 14, True, GOLD),
             (caption, 13, False, MIST)]):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = t
        r.font.name = BODY_FONT
        r.font.size = Pt(sz)
        r.font.bold = b
        r.font.color.rgb = col


def star_ring(slide, cx, cy, radius, star_size=Inches(0.30), color=GOLD):
    for k in range(12):
        a = math.radians(k * 30 - 90)
        x = Emu(int(cx + radius * math.cos(a) - star_size / 2))
        y = Emu(int(cy + radius * math.sin(a) - star_size / 2))
        box(slide, x, y, star_size, star_size, fill=color,
            shape=MSO_SHAPE.STAR_5_POINT)


FOOTER = "PROTMŪŠIS  ·  16 June 2026  ·  Permanent Representation of Lithuania to the EU"


# ------------------------------------------------------------ title slide --
def title_slide():
    s = add_slide(
        "Host notes: welcome everyone, explain the team draw (slips with names "
        "of Lithuanian rivers and lakes at the door).")
    star_ring(s, int(SLIDE_W / 2), Inches(2.05), Inches(1.15))
    text(s, Inches(1.5), Inches(3.1), Inches(10.33), Inches(1.3), "PROTMŪŠIS",
         size=66, bold=True, color=WHITE, align=PP_ALIGN.CENTER,
         font=TITLE_FONT)
    text(s, Inches(1.5), Inches(4.35), Inches(10.33), Inches(0.5),
         "QUIZ NIGHT", size=22, color=GOLD, align=PP_ALIGN.CENTER, bold=True)
    text(s, Inches(1.5), Inches(5.15), Inches(10.33), Inches(0.9),
         ["Permanent Representation of Lithuania to the European Union",
          "Brussels  ·  16 June 2026"],
         size=17, color=MIST, align=PP_ALIGN.CENTER, spacing=1.25)
    tricolour_footer(s, "")


def welcome_slide():
    s = add_slide(
        "Priskyrimas į komandas – lapelių traukimas įeinant, ant jų LT "
        "upių/ežerų pavadinimai.")
    chip(s, "BEFORE WE START", w=Inches(2.6))
    text(s, Inches(0.55), Inches(1.25), Inches(12.2), Inches(0.9),
         "How tonight works", size=38, bold=True, font=TITLE_FONT)
    rules = [
        ("Teams", "Draw a slip at the door — the Lithuanian river or lake on "
                  "it is your team."),
        ("Rounds", "Three rounds plus a Brussels-bubble bonus round. One "
                   "point per question unless marked otherwise."),
        ("Answers", "Write your answers down; we reveal and score after each "
                    "round."),
        ("Fair play", "No phones — diplomatic immunity does not cover "
                      "googling."),
    ]
    y = 2.45
    for head, body_t in rules:
        box(s, Inches(0.55), Inches(y + 0.08), Inches(0.09), Inches(0.78),
            fill=GOLD)
        text(s, Inches(0.9), Inches(y), Inches(11.6), Inches(1.0),
             [[(head + "   ", {"bold": True, "color": GOLD, "size": 19}),
               (body_t, {"size": 19, "color": WHITE})]], spacing=1.1)
        y += 1.06
    tricolour_footer(s, FOOTER)


def divider(round_no, title, sub, notes=None):
    s = add_slide(notes)
    star_ring(s, int(SLIDE_W / 2), Inches(1.7), Inches(0.85),
              star_size=Inches(0.2))
    text(s, Inches(1.5), Inches(2.8), Inches(10.33), Inches(0.6), round_no,
         size=20, bold=True, color=GOLD, align=PP_ALIGN.CENTER)
    text(s, Inches(0.8), Inches(3.35), Inches(11.73), Inches(1.2), title,
         size=48, bold=True, align=PP_ALIGN.CENTER, font=TITLE_FONT)
    text(s, Inches(1.5), Inches(4.7), Inches(10.33), Inches(0.6), sub,
         size=17, color=MIST, align=PP_ALIGN.CENTER)
    tricolour_footer(s, FOOTER)


def question_slide(round_label, q_no, question, options=None, hint=None,
                   media=None, notes=None, points=None):
    s = add_slide(notes)
    chip(s, round_label, w=Inches(4.6))
    pts = f"  ·  {points}" if points else ""
    text(s, Inches(9.0), Inches(0.42), Inches(3.78), Inches(0.6),
         f"QUESTION {q_no}{pts}", size=16, bold=True, color=GOLD,
         align=PP_ALIGN.RIGHT)
    text_w = Inches(7.7) if media else Inches(12.2)
    text(s, Inches(0.55), Inches(1.45), text_w, Inches(3.1), question,
         size=24, spacing=1.15)
    y = 4.65
    if options:
        for letter, opt in options:
            text(s, Inches(0.8), Inches(y), text_w - Inches(0.3), Inches(0.5),
                 [[(letter + ")  ", {"bold": True, "color": GOLD, "size": 20}),
                   (opt, {"size": 20, "color": WHITE})]])
            y += 0.52
    if hint:
        text(s, Inches(0.55), Inches(6.25), Inches(12.2), Inches(0.5),
             [[("Hint:  ", {"bold": True, "color": GOLD, "size": 15}),
               (hint, {"size": 15, "color": MIST, "italic": True})]])
    if media:
        kind, caption = media
        media_placeholder(s, caption, Inches(8.6), Inches(1.45), Inches(4.15),
                          Inches(4.6), kind=kind)
    tricolour_footer(s, FOOTER)


def answer_slide(round_label, q_no, answer, fact=None, media=None, notes=None):
    s = add_slide(notes)
    chip(s, round_label, w=Inches(4.6))
    text(s, Inches(9.0), Inches(0.42), Inches(3.78), Inches(0.6),
         f"ANSWER {q_no}", size=16, bold=True, color=GOLD,
         align=PP_ALIGN.RIGHT)
    text_w = Inches(7.7) if media else Inches(12.2)
    text(s, Inches(0.55), Inches(1.7), text_w, Inches(1.6), answer,
         size=40, bold=True, color=GOLD, font=TITLE_FONT, spacing=1.05)
    if fact:
        text(s, Inches(0.55), Inches(3.9), text_w, Inches(2.6),
             [[("Did you know?  ", {"bold": True, "color": GOLD, "size": 17})],
              [(fact, {"size": 17, "color": MIST})]],
             spacing=1.2, space_after=Pt(6))
    if media:
        kind, caption = media
        media_placeholder(s, caption, Inches(8.6), Inches(1.45), Inches(4.15),
                          Inches(4.6), kind=kind)
    tricolour_footer(s, FOOTER)


def closing_slide():
    s = add_slide("Announce the scores and the winning team.")
    star_ring(s, int(SLIDE_W / 2), Inches(2.0), Inches(1.0),
              star_size=Inches(0.24))
    text(s, Inches(1.5), Inches(3.1), Inches(10.33), Inches(1.2), "AČIŪ!",
         size=60, bold=True, align=PP_ALIGN.CENTER, font=TITLE_FONT)
    text(s, Inches(1.5), Inches(4.35), Inches(10.33), Inches(0.6),
         "Thank you for playing — time to count the points.",
         size=20, color=GOLD, align=PP_ALIGN.CENTER)
    tricolour_footer(s, FOOTER)


# ------------------------------------------------------------------ deck ---
title_slide()
welcome_slide()

# ===== ROUND 1 ==============================================================
R1 = "ROUND 1 · EUROPE & FUN FACTS"
divider("ROUND 1", "Europe & Fun Facts", "Multiple choice · 6 questions · 1 point each")

question_slide(
    R1, 1, "Which European island changes sovereignty every six months?",
    options=[("A", "Heligoland"), ("B", "Pheasant Island"), ("C", "Jersey"),
             ("D", "Bornholm")])
answer_slide(
    R1, 1, "B)  Pheasant Island",
    fact="Under the 1659 Treaty of the Pyrenees, France and Spain share this "
         "tiny uninhabited island in the Bidasoa river. Administration "
         "alternates between the two countries every six months — the "
         "world's oldest condominium.",
    media=("PHOTO", "Pheasant Island (Île des Faisans / Isla de los "
                    "Faisanes) — aerial view"))

question_slide(
    R1, 2, "Which EU symbol was publicly unveiled about 40 years ago?",
    options=[("A", "Euro coins"), ("B", "European flag"),
             ("C", "European anthem"), ("D", "Schengen passport")])
answer_slide(
    R1, 2, "B)  European flag",
    fact="In May 1986 the European flag was raised for the first time "
         "outside the Berlaymont building — twelve gold stars in a circle, "
         "a symbol of unity, unchanged ever since.",
    media=("PHOTO", "1986 flag-raising ceremony at the Berlaymont"))

question_slide(
    R1, 3, "What major international award did the European Union receive "
           "in 2012?",
    options=[("A", "Sakharov Prize"), ("B", "Nobel Peace Prize"),
             ("C", "Charlemagne Prize"), ("D", "Right Livelihood Award")])
answer_slide(
    R1, 3, "B)  Nobel Peace Prize",
    fact="Awarded to the EU in 2012 for over six decades of advancing "
         "peace, reconciliation, democracy and human rights in Europe.")

question_slide(
    R1, 4, "Which euro banknote was nicknamed “Bin Laden”?",
    options=[("A", "€50"), ("B", "€100"), ("C", "€200"), ("D", "€500")],
    notes="Source: https://www.theguardian.com/business/2016/may/04/"
          "500-euro-banknote-could-be-scrapped-crime")
answer_slide(
    R1, 4, "D)  €500",
    fact="The ECB stopped issuing the €500 note in 2019 over its popularity "
         "with money launderers — and because, like its namesake, everyone "
         "knew what it looked like but almost no one had ever seen one.",
    media=("PHOTO", "The €500 banknote"))

question_slide(
    R1, 5, "In 1984, former Nigerian minister Umaru Dikko was kidnapped in "
           "London. His captors planned to smuggle him out of the UK in a "
           "diplomatic crate, relying on the Vienna Convention rule that "
           "diplomatic bags cannot be opened or detained.\n\nWhat went "
           "wrong?",
    options=[("A", "The crate was improperly labelled as diplomatic baggage"),
             ("B", "Dikko did not fit inside the crate"),
             ("C", "The aircraft meant to transport him never arrived")],
    notes="Video background: https://www.youtube.com/watch?v=N83Idy9IOmU")
answer_slide(
    R1, 5, "A)  The crate was improperly\nlabelled",
    fact="With no official diplomatic markings, customs officers at "
         "Stansted were entitled to open the crate — and found Dikko "
         "unconscious, accompanied by an anaesthetist.",
    notes="Video: https://www.youtube.com/watch?v=N83Idy9IOmU")

question_slide(
    R1, 6, "One of the inventors of modern blue jeans was born in "
           "present-day Latvia. He developed a way to make work trousers "
           "much more durable and partnered with Levi Strauss to patent the "
           "idea.\n\nWhat simple innovation made the trousers famous?",
    options=[("A", "A zipper"), ("B", "Metal rivets"),
             ("C", "Waterproof fabric"), ("D", "Belt loops")])
answer_slide(
    R1, 6, "B)  Metal rivets",
    fact="Jacob Davis, a tailor born in Riga in 1831, reinforced the stress "
         "points of work trousers with copper rivets. He and Levi Strauss "
         "patented the idea in 1873 — and blue jeans were born.")

# ===== ROUND 2 ==============================================================
R2 = "ROUND 2 · LITHUANIA"
divider("ROUND 2", "Lithuania", "8 questions · 1 point each")

question_slide(
    R2, 1, "Every year Lithuania plans to win Eurovision — although "
           "officially we never have.\n\nIn 2006, however, Lithuania tried a "
           "different strategy: entering a song that simply declared victory "
           "before the contest was even over.\n\nWhat was the title of the "
           "song?",
    media=("VIDEO", "Song excerpt — LT United, Eurovision 2006 (choose "
                    "which part to show)"),
    notes="Video: https://www.youtube.com/watch?v=DBAdOlQPbwg\n"
          "Pasitarkime, kurią dalį rodysime.\n"
          "[Svarstyta perkelti prie muzikinio/vaizdų turo.]")
answer_slide(
    R2, 1, "“We Are The Winners”",
    fact="Honourable mention: Lithuanian-German singer Lena Valaitis took "
         "second place at Eurovision 1981 — proof that we keep looking for "
         "that victory by all possible means.",
    media=("PHOTO", "LT United on the Eurovision 2006 stage / Lena Valaitis, "
                    "1981"),
    notes="Lena Valaitis paminėti kaip faktą, ne kaip klausimą.")

question_slide(
    R2, 2, "Two Lithuanian mountaineers carried a symbolic national object "
           "to the summit of Mount Everest and scattered it from the "
           "world's highest peak.\n\nWhat was it?",
    notes="Pačią istoriją ir konkrečius metus papasakoti žodžiu, skaidrėse "
          "rašyti lakoniškiau.")
answer_slide(
    R2, 2, "Amber",
    fact="Vladas Vitkauskas (1995) and Saulius Vilius (2003) both carried "
         "Baltic amber to the top of the world and scattered it from the "
         "summit.",
    media=("PHOTO", "The climbers and Baltic amber"))

question_slide(
    R2, 3, "When Lithuanians learn about the interwar period, one statistic "
           "is proudly repeated in schools: in 1938, Lithuania ranked second "
           "in Europe in butter and bacon production — despite being one of "
           "the poorest countries on the continent.\n\nWhich country ranked "
           "first?",
    hint="Think of a recent Council Presidency.")
answer_slide(
    R2, 3, "Denmark",
    fact="A statistic still celebrated in Lithuanian classrooms. Denmark, "
         "for its part, no longer wants to be the EU's bacon factory "
         "(Politico).",
    media=("PHOTO", "Interwar statistics — or simply butter and bacon"),
    notes="Politico: “Denmark no longer wants to be EU bacon factory”.")

question_slide(
    R2, 4, "Across Europe people celebrate Midsummer with bonfires. In "
           "Lithuania, one tradition involves searching for a mythical "
           "flower that supposedly blooms only on that night.\n\nWhich "
           "flower?")
answer_slide(
    R2, 4, "The fern flower",
    fact="According to tradition, the fern blooms only on Midsummer night "
         "(Joninės) — whoever finds it gains happiness and wisdom. "
         "Botanists remain unconvinced.",
    media=("PHOTO", "Fern flower (paparčio žiedas)"))

question_slide(
    R2, 5, "1918 was a big year for Lithuania — also for democracy.\n\n"
           "Lithuanian women gained something that many women in Western "
           "Europe had to wait decades for.\n\nWhat was it?",
    notes="Alternatyva: rodyti pašto ženklą (kai ką pašalinus) ir klausti, "
          "kokia proga jis išleistas.\n"
          "Stamp: https://manoteises.lt/straipsnis/isleidziamas-pasto-"
          "zenklas-moteru-balsavimo-simtmeciui-lietuvoje-pamineti/")
answer_slide(
    R2, 5, "Voting rights",
    fact="Lithuanian women gained the vote on 2 November 1918, when the "
         "provisional constitution enshrined equal suffrage — ahead of much "
         "of Western Europe. A commemorative stamp marked the centenary.",
    media=("PHOTO", "Centenary postage stamp — women's suffrage in "
                    "Lithuania"))

question_slide(
    R2, 6, "A 16th-century Lithuanian nobleman travelled through Egypt and "
           "bought several unusual souvenirs. During a storm at sea, "
           "terrified sailors blamed the bad weather on them and threw them "
           "overboard.\n\nWhat were the souvenirs?")
answer_slide(
    R2, 6, "Egyptian mummies",
    fact="Mikalojus Kristupas Radvila Našlaitėlis brought two mummies back "
         "from his pilgrimage. When storms battered the ship, the crew "
         "blamed the cargo — and overboard they went. His travel diary "
         "became a European bestseller.",
    media=("PHOTO", "Radvila Našlaitėlis portrait · book illustration · the "
                    "mummies"))

question_slide(
    R2, 7, "In a humorous Lithuanian promotional video from the early "
           "2000s, traditional dishes such as cepelinai are discussed with "
           "the idea that they might one day become popular across Europe "
           "as street food.\n\nThe video reflects a moment when Lithuania "
           "was preparing for a major change. What event was approaching?",
    media=("VIDEO", "Promo video excerpt — add English subtitles"),
    notes="Video: https://www.youtube.com/watch?v=YgvHcenDYcU\n"
          "Uždėti angliškus subtitrus klausimui. Pagalvoti, ar įdėsime "
          "frazę „early 2000s“.")
answer_slide(
    R2, 7, "Joining the European Union",
    fact="Lithuania joined the EU on 1 May 2004, in the largest enlargement "
         "in the Union's history — ten countries at once. The cepelinai "
         "street-food revolution is still pending.",
    media=("PHOTO", "Map of the 2004 enlargement — the EU grows"))

question_slide(
    R2, 8, "Pink is strongly associated with Lithuania because of a "
           "traditional summer dish.\n\nWhich dish?",
    options=[("A", "Cepelinai"), ("B", "Šakotis"), ("C", "Šaltibarščiai"),
             ("D", "Kibinai")])
answer_slide(
    R2, 8, "C)  Šaltibarščiai",
    fact="The electric-pink cold beet soup is Lithuania's unofficial summer "
         "flag — best served with hot potatoes and a sunny terrace.",
    media=("PHOTO", "Šaltibarščiai in full pink glory"))

# ===== ROUND 3 ==============================================================
R3 = "ROUND 3 · ENVIRONMENT & NATURE"
divider("ROUND 3", "Environment & Nature", "4 questions · 1 point each")

question_slide(
    R3, 1, "In 2013 Metallica became the first band to perform on all seven "
           "continents. During their Antarctic concert, the audience used "
           "120 of what?")
answer_slide(
    R3, 1, "Headphones",
    fact="To comply with Antarctic environmental rules, the band played "
         "without amplifiers — the audience of about 120 listened through "
         "headphones. The concert was fittingly called “Freeze 'Em All”.",
    media=("PHOTO", "Metallica's Antarctic concert (photo or video)"))

question_slide(
    R3, 2, "Three famous scientists — Birutė Galdikas, Dian Fossey and Jane "
           "Goodall — devoted their careers to studying these animals in "
           "the wild. Collectively, they became known as the "
           "“Trimates”.\n\nWhat group of animals are these?")
answer_slide(
    R3, 2, "Great apes",
    fact="Goodall (chimpanzees), Fossey (gorillas) and Galdikas "
         "(orangutans) were recruited by Louis Leakey — hence also "
         "“Leakey's Angels”. Galdikas, of Lithuanian descent, still works "
         "in Borneo. “Orangutan” means “person of the forest”.",
    media=("PHOTO", "The three scientists"),
    notes="Photos: https://www.themarysue.com/the-trimates-three-women-that-"
          "made-science-history/")

question_slide(
    R3, 3, "Which EU law unexpectedly entered the headlines after the "
           "deaths of four US soldiers, whose vehicle sank in a military "
           "training area near the Lithuanian–Belarusian border?")
answer_slide(
    R3, 3, "The Nature Restoration Law",
    fact="Commentary on the 2025 tragedy near Pabradė linked the swampy "
         "terrain to wetlands restored under EU environmental policy — "
         "putting the Nature Restoration Law unexpectedly in the news.")

question_slide(
    R3, 4, "This animal is often called an “ecosystem engineer” because it "
           "creates wetlands that benefit countless other species.\n\nWhat "
           "animal is it?")
answer_slide(
    R3, 4, "The beaver",
    fact="Beaver dams create wetlands that store water, filter pollution "
         "and host countless species. Lithuania's beaver population has "
         "grown from near extinction to one of the densest in Europe.",
    media=("PHOTO", "Beaver — plus population statistics if available"))

# ===== BONUS ROUND ==========================================================
RB = "BONUS · THE BRUSSELS BUBBLE"
divider("BONUS ROUND", "The Brussels Bubble",
        "Insider questions · for those who were in the room",
        notes="Galimi papildomi klausimai: CBAM apie vinis (paformuluoti), "
              "the Schumann show apie institucijas "
              "(https://www.youtube.com/watch?v=D0fBh-0Eiy0).")

question_slide(
    RB, 1, "During negotiations on which file did Violeta Dragu (Romania) "
           "say:\n\n“I have never seen Lithuania being so active during the "
           "negotiations”?",
    points="1 pt",
    notes="Galima multiple choice, jei norime palengvinti. Galima "
          "pasunkinti ir paklausti, kokia buvo problema.")
answer_slide(
    RB, 1, "The Waste Framework\nDirective")

question_slide(
    RB, 2, "During a Coreper I meeting, the Council was split on a "
           "legislative file, with both sides carefully counting votes.\n\n"
           "The Deputy Permanent Representative of Lithuania intervened and "
           "said… what?\n\nAnd which file was being discussed?",
    points="2 pts",
    notes="Jei norime su užuominomis – galime nurodyti konkrečią datą: "
          "birželio 14 d. Coreper I (arba dar vienas anksčiau buvęs).")
answer_slide(
    RB, 2, "“Scrutiny reservation” —\non the Nature Restoration Law")

question_slide(
    RB, 3, "1985: Lithuanian, Latvian and Estonian dissidents organised the "
           "Peace and Freedom Cruise aboard the ship Baltic Star — sailing "
           "from Stockholm along the edge of Soviet territorial waters, "
           "past Neringa, Königsberg, Palanga and Liepāja, and ending with "
           "a widely reported demonstration in Helsinki.\n\nThe Soviet "
           "Union tried to derail the voyage. What did it do?",
    points="1 pt",
    notes="Anot T. Venclovos, A. Štromo buvo žygio spiritus movens; 1985 m. "
          "taip pat vyko Baltijos tribunolas. Pavardžių skaidrėje neminėti. "
          "Laivo nuotrauka bus atsiųsta vėliau.")
answer_slide(
    RB, 3, "A rumour that a bomb was on board",
    fact="The rumour failed: the cruise went ahead, and the Helsinki "
         "demonstration was reported across Scandinavian and wider European "
         "press — a loud reminder that the Baltic states had not been "
         "forgotten.",
    media=("PHOTO", "The ship Baltic Star (photo to be sent)"))

closing_slide()

OUT = "Protmusis_2026-06-16.pptx"
prs.save(OUT)
print(f"Saved {OUT} with {len(prs.slides._sldIdLst)} slides")
