# Trivia Night — slides

Slide deck for the **16 June trivia night** at the Permanent Representation of
Lithuania to the EU, built from the questions drafted in
`Klausimai protmūšiui.docx`.

Format: **one slide — question, next slide — answer.** The deck is fully in
English for an international audience. Production notes from the draft
(source links, hint options, alternative question versions) live in the
**speaker notes** of each slide.

## Files

| File | Purpose |
|---|---|
| `Trivia_Night_2026-06-16.pptx` | The deck (49 slides, 16:9) — open in PowerPoint |
| `Trivia_Night_2026-06-16.pdf` | PDF preview for quick review |
| `generate_slides.py` | Builds the deck: `pip install python-pptx && python3 generate_slides.py` |
| `fetch_assets.py` | Downloads/renders the illustrations into `assets/`: `pip install cairosvg && python3 fetch_assets.py` |
| `assets/` | Rendered illustrations (committed, so the deck builds offline) |
| `Klausimai protmūšiui.docx` | Original question draft |

## Design

- Bold dark theme: near-black slides, with one saturated colour per round —
  EU blue for Round 1, then the Lithuanian tricolour (yellow, green, red)
  for Rounds 2–3 and the bonus round.
- Full-bleed colour divider slides with oversized numerals; question slides
  carry a giant question number and a colour-coded edge bar; answers are
  revealed in the round colour.
- Sharp, rectangular geometry throughout — no rounded corners.
- [Montserrat](https://github.com/JulietaUla/Montserrat) typography
  (SIL OFL 1.1), **embedded in the .pptx** — the deck renders identically
  on machines that don't have the font installed.
- Real photographs (`photos/`, see `photos/CREDITS.md`) illustrate most
  questions and answers, placed at native aspect ratio — never cropped —
  and sized to fill the right side of the slide; remaining icons are
  [OpenMoji](https://openmoji.org) (CC BY-SA 4.0).

## Deck structure

1. **Title + rules** — team draw at the door (slips with Lithuanian rivers/lakes)
2. **Round 1 — Europe & Fun Facts** · 6 multiple-choice questions
3. **Round 2 — Lithuania** · 8 questions
4. **Round 3 — Environment & Nature** · 4 questions
5. **Bonus — The Brussels Bubble** · 3 insider questions (1–2 points)
6. **Closing / scoring**

## Videos

Questions built around a clip have a **“Play the clip”** button on the slide
with the URL next to it (and in the speaker notes):

- Round 2 Q1 — LT United, *We Are The Winners* (Eurovision 2006):
  <https://www.youtube.com/watch?v=DBAdOlQPbwg> — decide which part to show
- Round 2 Q7 — EU accession promo video (needs English subtitles):
  <https://www.youtube.com/watch?v=YgvHcenDYcU>
- Round 1 Q5 answer — the Umaru Dikko story:
  <https://www.youtube.com/watch?v=N83Idy9IOmU>

To embed the clips offline, download the excerpts and insert them over the
play-button slides in PowerPoint (Insert → Video).

## Question pool not yet on slides (from the draft — to discuss)

- CBAM question about nails (needs wording)
- The Schumann Show / the institutions ([video](https://www.youtube.com/watch?v=D0fBh-0Eiy0))
- One question per member state / quick-fire round idea
- Questions voiced by the Minister and Selemonas Paltanavičius (incl. the Belém fire)
- Napoleon / St Anne's church; Vytis; the Pacai–Pazzi family link; Bona Sforza
- Čiurlionis + the oceans act / a music question; German street in Vilnius
- Kaunas modernism / Vilnius baroque; football (France / World Cup angle)
- Adamkus 100 / Attenborough 100; the bears offered by Slovenia
- Baltic Way symbols in the European Parliament
