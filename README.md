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
| `Trivia_Night_2026-06-16.pptx` | The deck (60 slides, 16:9) — open in PowerPoint |
| `Trivia_Night_2026-06-16.pdf` | Static preview — **re-export from PowerPoint** after rebuilding; a PDF cannot play the embedded audio/video |
| `Trivia_Night_2026-06-16_run_of_show.docx` | Run of show: every round/question/answer in presentation order, plus the changes made relative to the draft |
| `generate_slides.py` | Builds the deck: `pip install python-pptx Pillow && python3 generate_slides.py` |
| `generate_rundown.py` | Builds the run-of-show docx: `pip install python-docx && python3 generate_rundown.py` |
| `fetch_assets.py` | Downloads/renders the illustrations into `assets/`: `pip install cairosvg && python3 fetch_assets.py` |
| `assets/` | Rendered illustrations (committed, so the deck builds offline) |
| `photos/` | Photographs and album art (`song_*`) used on the slides |
| `media/audio/` | Song excerpts embedded in the Music round |
| `media/video/` | Video clips embedded in the deck |
| `Klausimai protmūšiui.docx` | Original question draft |

## Design

- Bold dark theme: near-black slides, with one saturated colour per round —
  EU blue for Round 1, then the Lithuanian tricolour (yellow, green, red)
  for Rounds 2–3 and the bonus round.
- Full-bleed colour divider slides with oversized numerals; question slides
  carry a giant question number and a colour-coded edge bar; answers are
  revealed in the round colour.
- Sharp, rectangular geometry throughout — no rounded corners.
- **Segoe UI** typography (Segoe UI Black for the display headings) — a
  font installed by default on Windows, so nothing has to be embedded and
  managed/locked-down Office installs render the deck as intended.
- Real photographs (`photos/`, see `photos/CREDITS.md`) illustrate most
  questions and answers, placed at native aspect ratio — never cropped —
  and sized to fill the right side of the slide; remaining icons are
  [OpenMoji](https://openmoji.org) (CC BY-SA 4.0).

## Deck structure

1. **Title + rules** — team draw at the door (slips with Lithuanian rivers/lakes)
2. **Round 1 — Europe & Fun Facts** · 6 multiple-choice questions
3. **Round 2 — Lithuania** · 8 questions
4. **Round 3 — Environment & Nature** · 4 questions
5. **Round 4 — Music (Name That Tune)** · 5 song excerpts — all secretly
   about the sea (a nod to Čiurlionis's *Jūra*)
6. **Bonus — The Brussels Bubble** · 3 insider questions (1–2 points)
7. **Closing / scoring**

## Embedded audio & video

All clips are **embedded in the .pptx and play inline** in PowerPoint
(offline). Each sits behind an on-brand poster frame; click the play
button to start. **Videos are set to play full screen.** A **hyperlinked
title sits under each video** (and the URL is in the speaker notes) so you
can fall back to YouTube if a clip won't play on the night.

- **Music round** — five excerpts in `media/audio/` (How Far I'll Go,
  My Heart Will Go On, Jūra, Yellow Submarine, Ocean Eyes), revealed with
  album art on the answer slides.
- Round 2 Q1 **answer** — LT United, *We Are The Winners*
  (`media/video/eurovision_we_are_the_winners.mp4`). The clip lives on the
  **answer** slide so it doesn't give the title away; the **question** keeps
  just the Eurovision logo. Trim to the part you want in PowerPoint
  (Playback ▸ Trim Video).
- Round 2 Q7 — EU-accession promo
  (`media/video/eu_accession_spirgi.mp4`) — add English subtitles before
  the night.
- Round 1 Q5 answer — the Umaru Dikko story
  (`media/video/umaru_dikko_kidnap.mp4`).

> The clips are sized to a 16:9 frame. If a source clip has a different
> aspect ratio it will letterbox/stretch — drag the frame's handles in
> PowerPoint to adjust.

## Question pool not yet on slides (from the draft — to discuss)

- CBAM question about nails (needs wording)
- The Schumann Show / the institutions ([video](https://www.youtube.com/watch?v=D0fBh-0Eiy0))
- One question per member state / quick-fire round idea
- Questions voiced by the Minister and Selemonas Paltanavičius (incl. the Belém fire)
- Napoleon / St Anne's church; Vytis; the Pacai–Pazzi family link; Bona Sforza
- German street in Vilnius
- Kaunas modernism / Vilnius baroque; football (France / World Cup angle)
- Adamkus 100 / Attenborough 100; the bears offered by Slovenia
- Baltic Way symbols in the European Parliament
